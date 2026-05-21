#!/usr/bin/env python
"""
ScienceDirect Paper Scraper v2.0
=================================
Uses curl_cffi to mimic Chrome TLS fingerprint and bypass Cloudflare bot detection.
Supports multiple search modes; results saved as CSV / JSON / XLSX.

Supported Search Modes
----------------------
  keyword         — search by keyword
  journal         — browse by journal name
  journal_keyword — search by keyword within a specific journal
  author          — search by author name
  issn            — search by journal ISSN
  advanced        — advanced search (combine multiple criteria)

Quick Start (CLI)
-----------------
  python sd_scraper_en.py -m keyword -q "machine learning" -n 100 --browser-cookies
  python sd_scraper_en.py -m journal  -j "Energy" -n 50 --browser-cookies
  python sd_scraper_en.py -m journal_keyword -j "Renewable Energy" -q "solar cell" -n 50 --browser-cookies

Interactive Wizard
------------------
  python sd_scraper_en.py

Install Dependencies
--------------------
  pip install curl_cffi browser-cookie3
"""

import json
import csv
import os
import sys
import re
import time
import random
import argparse
from datetime import datetime
from urllib.parse import urlencode

from curl_cffi import requests as curl_requests

try:
    import browser_cookie3
    HAS_BROWSER_COOKIE3 = True
except ImportError:
    HAS_BROWSER_COOKIE3 = False

try:
    from openpyxl import Workbook
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


# ──────────────────────────────────────────────────────────────────────────────
# DevTools PDF capture (pure websocket-client, no Playwright required)
# ──────────────────────────────────────────────────────────────────────────────

def _dt_get_header(headers, name: str) -> str:
    target = name.lower()
    if isinstance(headers, dict):
        for k, v in headers.items():
            if str(k).lower() == target:
                return str(v).lower()
    elif isinstance(headers, list):
        for entry in headers:
            if isinstance(entry, dict) and str(entry.get("name", "")).lower() == target:
                return str(entry.get("value", "")).lower()
    return ""


def _dt_is_pdf_url(url: str) -> bool:
    low = (url or "").lower()
    return ".pdf" in low or "pdf.sciencedirectassets.com" in low


def _dt_is_pdf_response(response: dict) -> bool:
    mime = (response.get("mimeType") or "").lower()
    ct = _dt_get_header(response.get("headers") or {}, "content-type")
    return "pdf" in mime or "pdf" in ct or _dt_is_pdf_url(response.get("url", ""))


def _dt_is_pdf_fetch_response(url: str, status, headers) -> bool:
    if status is None or status < 200 or status >= 400:
        return False
    ct = _dt_get_header(headers, "content-type")
    return "pdf" in ct or _dt_is_pdf_url(url)


BLOCK_PAGE_SIGNALS = (
    "there was a problem providing",
    "verify you are human",
    "captcha",
    "are you a robot",
    "robot or human",
    "challenge-platform",
    "cf-browser-verification",
    "access denied",
)

CAPTCHA_SIGNALS = (
    "are you a robot",
    "verify you are human",
    "captcha",
    "robot or human",
    "challenge-platform",
    "cf-browser-verification",
)

# Injected before each page load to hide Chrome automation fingerprints
_STEALTH_JS = """
(function() {
    // 1. Hide webdriver flag (most common detection point)
    Object.defineProperty(navigator, 'webdriver', { get: () => undefined });

    // 2. Populate navigator.plugins (empty in controlled Chrome)
    if (navigator.plugins.length === 0) {
        Object.defineProperty(navigator, 'plugins', {
            get: () => {
                const arr = [
                    { name: 'Chrome PDF Plugin',     filename: 'internal-pdf-viewer' },
                    { name: 'Chrome PDF Viewer',     filename: 'mhjfbmdgcfjbbpaeojofohoefgiehjai' },
                    { name: 'Native Client',         filename: 'internal-nacl-plugin' },
                ];
                arr.item = (i) => arr[i];
                arr.namedItem = (n) => arr.find(p => p.name === n) || null;
                arr.refresh = () => {};
                Object.setPrototypeOf(arr, PluginArray.prototype);
                return arr;
            }
        });
    }

    // 3. Populate navigator.languages
    if (!navigator.languages || navigator.languages.length === 0) {
        Object.defineProperty(navigator, 'languages', { get: () => ['en-US', 'en', 'zh-CN', 'zh'] });
    }

    // 4. Populate window.chrome (sometimes missing in controlled Chrome)
    if (!window.chrome) {
        window.chrome = {
            runtime: {},
            loadTimes: function() {},
            csi: function() {},
            app: {},
        };
    }

    // 5. Keep Notification.permission as 'default'
    if (window.Notification) {
        try {
            Object.defineProperty(Notification, 'permission', { get: () => 'default' });
        } catch (_) {}
    }

    // 6. Remove CDP-injected global variable traces
    try { delete window.__nightmare; } catch (_) {}
    try { delete window._phantom;    } catch (_) {}
    try { delete window.callPhantom;  } catch (_) {}
    try { delete document.__defineGetter__; } catch (_) {}
})();
"""


def _dt_capture_pdf(ws_url: str, url: str, timeout: int = 35):
    """
    Navigate to url in an existing DevTools tab and capture PDF bytes via
    Network/Fetch interception.
    Returns (bytes | None, note_str).
    note_str starting with "blocked:" means a bot-detection page was hit.
    """
    try:
        import websocket as _ws
    except ImportError:
        return None, "websocket-client not installed"

    import base64 as _b64

    ws = _ws.create_connection(ws_url, timeout=180, suppress_origin=True)
    ws.settimeout(1)
    msg_id = 200

    def send(method, params=None):
        nonlocal msg_id
        msg_id += 1
        ws.send(json.dumps({"id": msg_id, "method": method, "params": params or {}}))
        return msg_id

    pdf_req_ids: set = set()
    body_reqs: dict = {}
    fetch_body_reqs: dict = {}
    fetch_meta: dict = {}
    evaluate_reqs: dict = {}
    last_pdf_url = ""
    last_error = ""

    try:
        send("Page.enable")
        send("Page.addScriptToEvaluateOnNewDocument", {"source": _STEALTH_JS})
        send("Fetch.enable", {"patterns": [
            {"urlPattern": "*pdf.sciencedirectassets.com/*", "requestStage": "Response"},
            {"urlPattern": "*pdfft*", "requestStage": "Response"},
        ]})
        send("Network.enable", {
            "maxTotalBufferSize": 120 * 1024 * 1024,
            "maxResourceBufferSize": 100 * 1024 * 1024,
        })
        send("Network.setCacheDisabled", {"cacheDisabled": True})
        send("Page.navigate", {"url": url})

        deadline = time.time() + max(5, timeout)
        while time.time() < deadline:
            ws.settimeout(max(0.5, min(2.0, deadline - time.time())))
            try:
                raw = ws.recv()
            except Exception:
                continue
            try:
                msg = json.loads(raw)
            except Exception:
                continue

            method = msg.get("method")
            if method == "Fetch.requestPaused":
                p = msg.get("params", {})
                req_id = p.get("requestId")
                req_url = (p.get("request") or {}).get("url", "")
                status = p.get("responseStatusCode")
                hdrs = p.get("responseHeaders") or []
                if req_id and _dt_is_pdf_fetch_response(req_url, status, hdrs):
                    fetch_meta[req_id] = {"url": req_url}
                    bid = send("Fetch.getResponseBody", {"requestId": req_id})
                    fetch_body_reqs[bid] = req_id
                elif req_id:
                    send("Fetch.continueRequest", {"requestId": req_id})
            elif method == "Network.responseReceived":
                p = msg.get("params", {})
                resp = p.get("response", {})
                if _dt_is_pdf_response(resp):
                    rid = p.get("requestId")
                    if rid:
                        pdf_req_ids.add(rid)
                        last_pdf_url = resp.get("url", "") or last_pdf_url
            elif method == "Network.loadingFinished":
                rid = msg.get("params", {}).get("requestId")
                if rid in pdf_req_ids and rid not in body_reqs.values():
                    bid = send("Network.getResponseBody", {"requestId": rid})
                    body_reqs[bid] = rid
            elif method == "Network.loadingFailed":
                p = msg.get("params", {})
                rid = p.get("requestId")
                if rid in pdf_req_ids:
                    last_error = p.get("errorText") or "pdf_loading_failed"
                    pdf_req_ids.discard(rid)
            elif method == "Page.loadEventFired":
                bid = send("Runtime.evaluate", {
                    "expression": "document.body ? document.body.innerText.slice(0,2000).toLowerCase() : ''",
                    "returnByValue": True,
                })
                evaluate_reqs[bid] = "block_check"
            elif msg.get("id") in evaluate_reqs:
                evaluate_reqs.pop(msg["id"])
                if "error" not in msg:
                    text = msg.get("result", {}).get("result", {}).get("value", "") or ""
                    if any(sig in text for sig in BLOCK_PAGE_SIGNALS):
                        return None, f"blocked:{text[:120]}"
            elif msg.get("id") in fetch_body_reqs:
                req_id = fetch_body_reqs.pop(msg["id"])
                result = msg.get("result", {})
                if "error" not in msg:
                    body = result.get("body")
                    if body:
                        data = _b64.b64decode(body) if result.get("base64Encoded") else body.encode("latin-1", errors="ignore")
                        if data[:4] == b"%PDF":
                            return data, fetch_meta.get(req_id, {}).get("url") or last_pdf_url or url
                try:
                    send("Fetch.continueRequest", {"requestId": req_id})
                except Exception:
                    pass
            elif msg.get("id") in body_reqs:
                body_reqs.pop(msg["id"])
                result = msg.get("result", {})
                if "error" not in msg:
                    body = result.get("body")
                    if body:
                        data = _b64.b64decode(body) if result.get("base64Encoded") else body.encode("latin-1", errors="ignore")
                        if data[:4] == b"%PDF":
                            return data, last_pdf_url or url
    finally:
        try:
            ws.close()
        except Exception:
            pass

    return None, last_error or "network_pdf_not_captured"


# ──────────────────────────────────────────────────────────────────────────────
# Core scraper class
# ──────────────────────────────────────────────────────────────────────────────

class ScienceDirectScraper:
    BASE_URL = "https://www.sciencedirect.com"
    SEARCH_API = "https://www.sciencedirect.com/search/api"

    def __init__(self, cookies_file=None, use_browser_cookies=False, delay_range=(2, 5)):
        self.session = curl_requests.Session(impersonate="chrome124")
        self.delay_range = delay_range
        self._search_token = None
        self._cookie_dict = {}
        self._session_cookies = {}

        if use_browser_cookies:
            self._load_browser_cookies()
        elif cookies_file:
            self._load_cookies(cookies_file)

    # ── Cookie support ────────────────────────────────────────────────────────

    def _apply_cookie_header(self):
        if self._cookie_dict:
            self.session.headers["Cookie"] = "; ".join(
                f"{k}={v}" for k, v in self._cookie_dict.items()
            )

    def _load_browser_cookies(self):
        """Read sciencedirect.com cookies directly from the local Chrome installation."""
        if not HAS_BROWSER_COOKIE3:
            print("[Error] browser-cookie3 not installed. Run: pip install browser-cookie3")
            return
        try:
            jar = browser_cookie3.chrome(domain_name='.sciencedirect.com')
            for c in jar:
                self._cookie_dict[c.name] = c.value
            if self._cookie_dict:
                self._apply_cookie_header()
                print(f"[Info] Loaded {len(self._cookie_dict)} cookies from Chrome (institutional mode)")
            else:
                print("[Warning] No sciencedirect.com cookies found in Chrome. Please log in via Chrome first.")
        except Exception as e:
            print(f"[Warning] Failed to read Chrome cookies: {e}")
            print("         On macOS, a Keychain permission dialog may appear — click Allow.")

    def _load_cookies(self, cookies_file):
        """Load cookies from a JSON file."""
        if not os.path.exists(cookies_file):
            print(f"[Warning] Cookie file not found: {cookies_file}. Running in guest mode.")
            return
        with open(cookies_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            for c in data:
                name = c.get("name") or c.get("Name", "")
                value = c.get("value") or c.get("Value", "")
                if name and value:
                    self._cookie_dict[name] = value
        elif isinstance(data, dict):
            self._cookie_dict.update({k: str(v) for k, v in data.items()})
        self._apply_cookie_header()
        print(f"[Info] Loaded {len(self._cookie_dict)} cookies from file (institutional mode)")

    # ── Internal utilities ────────────────────────────────────────────────────

    def _delay(self):
        time.sleep(random.uniform(*self.delay_range))

    def _fetch_search_token(self, params: dict):
        """
        Fetch the session cookie and searchToken required for the search API.
        Uses a clean session (no institutional cookies) to avoid session conflicts.
        """
        url = self.BASE_URL + "/search?" + urlencode(params)
        try:
            resp = self.session.get(url, timeout=25, headers={"Cookie": ""})
            if resp.status_code != 200:
                print(f"  [Warning] Search page returned HTTP {resp.status_code}")
                return None

            self._session_cookies = {}
            for k, v in resp.headers.items():
                if k.lower() == "set-cookie":
                    part = v.split(";")[0]
                    if "=" in part:
                        name, val = part.split("=", 1)
                        self._session_cookies[name.strip()] = val.strip()
            self.session.headers["Cookie"] = "; ".join(
                f"{k}={v}" for k, v in self._session_cookies.items()
            )

            m = re.search(r'"searchToken":"([^"]+)"', resp.text)
            if m:
                self._search_token = m.group(1)
                return self._search_token
            else:
                print("  [Warning] searchToken not found in page")
                return None
        except Exception as e:
            print(f"  [Network error] {e}")
            return None

    def _search(self, params: dict, max_count: int = 100):
        """
        Two-step search:
        Step 1: Fetch search page HTML → get csrf_token cookie + searchToken
        Step 2: Call /search/api with token, paginating as needed
        """
        results = []
        offset = 0
        per_page = 25
        total_known = None

        token_params = {k: v for k, v in params.items()}
        token_params["offset"] = 0
        token_params["show"] = per_page
        token = self._fetch_search_token(token_params)
        if not token:
            print("  [Error] Could not obtain search token. Check network or cookies.")
            return results

        self._delay()

        _token_retries = 0
        while len(results) < max_count:
            api_params = {k: v for k, v in params.items()}
            api_params["offset"] = offset
            api_params["show"] = per_page
            api_params["t"] = token
            api_params["hostname"] = "www.sciencedirect.com"

            api_url = self.SEARCH_API + "?" + urlencode(api_params)

            try:
                resp = self.session.get(
                    api_url,
                    timeout=20,
                    headers={
                        "Accept": "application/json, text/javascript, */*; q=0.01",
                        "Referer": self.BASE_URL + "/search?" + urlencode(params),
                        "X-Requested-With": "XMLHttpRequest",
                    },
                )
                if resp.status_code == 401:
                    if _token_retries >= 2:
                        print("  [Error] Token invalid after multiple refreshes. Stopping.")
                        break
                    _token_retries += 1
                    wait = 5 * _token_retries
                    print(f"  [Info] Token expired. Waiting {wait}s then refreshing...")
                    time.sleep(wait)
                    token_params["offset"] = offset
                    token = self._fetch_search_token(token_params)
                    if not token:
                        print("  [Error] Could not refresh token.")
                        break
                    continue
                if resp.status_code == 429:
                    print("  [Rate limit] Too many requests. Waiting 30s...")
                    time.sleep(30)
                    continue
                if resp.status_code != 200:
                    print(f"  [HTTP {resp.status_code}] API request failed")
                    break
                _token_retries = 0
                data = resp.json()
            except json.JSONDecodeError:
                print("  [Error] Response is not JSON — possible bot detection")
                break
            except Exception as e:
                print(f"  [Network error] {e}")
                break

            items = data.get("searchResults", [])
            if total_known is None:
                total_known = int(data.get("resultsFound", data.get("totalResults", 0)))
                actual_max = min(max_count, total_known) if total_known else max_count
                if total_known:
                    print(f"  Found {total_known} results. Planning to fetch {actual_max}.")
                else:
                    print(f"  Page returned {len(items)} results.")

            if not items:
                if total_known == 0:
                    print("  No matching results.")
                else:
                    print("  No more results.")
                break

            for item in items:
                if len(results) >= max_count:
                    break
                article = self._parse_article(item)
                results.append(article)
                idx = len(results)
                actual_max = min(max_count, total_known or max_count)
                title_preview = (article["title"] or "(no title)")[:60]
                print(f"  [{idx}/{actual_max}] {title_preview}")

            offset += per_page
            if total_known and offset >= total_known:
                break
            if len(results) >= max_count:
                break

            self._delay()

        return results

    def _parse_article(self, item: dict) -> dict:
        """Parse a single search result item."""
        authors_raw = item.get("authors", [])
        if isinstance(authors_raw, list):
            names = []
            for a in authors_raw:
                if isinstance(a, dict):
                    name = a.get("name", "")
                    if not name:
                        name = f"{a.get('givenName', '')} {a.get('surname', '')}".strip()
                    if name:
                        names.append(name)
                elif isinstance(a, str):
                    names.append(a)
            authors_str = "; ".join(names)
        elif isinstance(authors_raw, dict):
            author_list = authors_raw.get("authorList", [])
            authors_str = "; ".join(
                f"{a.get('givenName', '')} {a.get('surname', '')}".strip()
                for a in author_list
            )
        else:
            authors_str = ""

        doi = item.get("doi") or item.get("prism:doi", "")
        link = item.get("link", "")
        if not link:
            link = f"https://doi.org/{doi}" if doi else ""
        elif not link.startswith("http"):
            link = self.BASE_URL + link

        sort_date = item.get("sortDate", "")
        date_str = sort_date[:10] if sort_date else ""
        year = date_str[:4] if date_str else ""

        volume_issue = item.get("volumeIssue", "")
        volume = volume_issue.replace("Volume ", "").strip() if volume_issue else ""

        pdf_info = item.get("pdf", {}) or {}
        pdf_link = pdf_info.get("downloadLink", "")
        if pdf_link and not pdf_link.startswith("http"):
            pdf_link = self.BASE_URL + pdf_link
        pii = item.get("pii", "")

        journal_raw = item.get("sourceTitle", "") or item.get("publicationName", "")
        journal_name = re.sub(r"<[^>]+>", "", journal_raw)

        return {
            "title":        item.get("title", ""),
            "authors":      authors_str,
            "journal":      journal_name,
            "volume":       volume,
            "issue":        item.get("issue", ""),
            "year":         year,
            "date":         date_str,
            "doi":          doi,
            "abstract":     item.get("abstract", ""),
            "article_type": item.get("articleType", ""),
            "open_access":  bool(item.get("openAccess") or item.get("openArchive")),
            "url":          link,
            "pdf_url":      pdf_link,
            "pii":          pii,
        }

    # ── Search modes ──────────────────────────────────────────────────────────

    def search_by_keyword(self, query, count=100, sort_by="relevance",
                          date_range=None, article_type=None):
        """Search by keyword (supports Boolean operators AND / OR / NOT)."""
        print(f"\n[Keyword Search]  Query: {query}")
        params = {"qs": query, "sortBy": sort_by}
        if date_range:
            params["date"] = date_range
        if article_type:
            params["articleTypes"] = article_type
        return self._search(params, max_count=count)

    def search_by_journal(self, journal_name, count=100, sort_by="date",
                          date_range=None):
        """Browse a journal by name."""
        print(f"\n[Journal Browse]  Journal: {journal_name}")
        params = {"pub": journal_name, "sortBy": sort_by}
        if date_range:
            params["date"] = date_range
        return self._search(params, max_count=count)

    def search_by_journal_keyword(self, journal_name, query, count=100,
                                  sort_by="relevance", date_range=None):
        """Search by keyword within a specific journal."""
        print(f"\n[Journal + Keyword]  Journal: {journal_name}  Query: {query}")
        params = {"pub": journal_name, "qs": query, "sortBy": sort_by}
        if date_range:
            params["date"] = date_range
        return self._search(params, max_count=count)

    def search_by_author(self, author_name, count=100, sort_by="date"):
        """Search by author name."""
        print(f"\n[Author Search]  Author: {author_name}")
        params = {"au": author_name, "sortBy": sort_by}
        return self._search(params, max_count=count)

    def search_by_issn(self, issn, count=100, sort_by="date", date_range=None):
        """Search by journal ISSN."""
        print(f"\n[ISSN Search]  ISSN: {issn}")
        params = {"issn": issn, "sortBy": sort_by}
        if date_range:
            params["date"] = date_range
        return self._search(params, max_count=count)

    def search_advanced(self, query=None, journal=None, author=None,
                        issn=None, date_range=None, article_type=None,
                        open_access_only=False, count=100, sort_by="relevance"):
        """Advanced search: combine multiple criteria."""
        print("\n[Advanced Search]")
        params = {"sortBy": sort_by}
        if query:
            params["qs"] = query;       print(f"  Keyword:      {query}")
        if journal:
            params["pub"] = journal;    print(f"  Journal:      {journal}")
        if author:
            params["au"] = author;      print(f"  Author:       {author}")
        if issn:
            params["issn"] = issn;      print(f"  ISSN:         {issn}")
        if date_range:
            params["date"] = date_range; print(f"  Date range:   {date_range}")
        if article_type:
            params["articleTypes"] = article_type; print(f"  Article type: {article_type}")
        if open_access_only:
            params["openAccess"] = "true"; print("  Open access only: Yes")
        return self._search(params, max_count=count)

    # ── Save results ──────────────────────────────────────────────────────────

    FIELDS = ["title", "authors", "journal", "volume", "issue",
              "year", "date", "doi", "abstract", "article_type",
              "open_access", "url", "pdf_url", "pii"]

    def save_to_csv(self, results, filename, output_dir):
        os.makedirs(output_dir, exist_ok=True)
        path = os.path.join(output_dir, filename)
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=self.FIELDS, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(results)
        print(f"\n[CSV] Saved → {path}  ({len(results)} papers)")
        return path

    def save_to_json(self, results, filename, output_dir):
        os.makedirs(output_dir, exist_ok=True)
        path = os.path.join(output_dir, filename)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        print(f"[JSON] Saved → {path}  ({len(results)} papers)")
        return path

    def save_to_xlsx(self, results, filename, output_dir):
        if not HAS_OPENPYXL:
            print("[Warning] openpyxl not installed. Falling back to CSV.")
            csv_name = os.path.splitext(filename)[0] + ".csv"
            return self.save_to_csv(results, csv_name, output_dir)

        os.makedirs(output_dir, exist_ok=True)
        path = os.path.join(output_dir, filename)
        wb = Workbook()
        ws = wb.active
        ws.title = "papers"
        ws.append(self.FIELDS)
        for item in results:
            ws.append([item.get(field, "") for field in self.FIELDS])
        wb.save(path)
        print(f"[XLSX] Saved → {path}  ({len(results)} papers)")
        return path

    @staticmethod
    def _make_pdf_filename(idx, article):
        """Generate filename: {index}_{first_author_surname}_{year}_{title}.pdf"""
        authors = article.get("authors", "")
        first_author = (authors.split(";")[0].strip().split()[-1]
                        if authors else "Unknown")
        year = article.get("year", "")
        safe_title = re.sub(r'[\\/*?:"<>|]', "", article.get("title", ""))[:60].strip()
        return f"{idx:03d}_{first_author}_{year}_{safe_title}.pdf"

    # ── PDF download helpers ──────────────────────────────────────────────────

    def _get_cookies_via_cdp(self, ctx):
        """Extract all relevant cookies from a CDP Chrome context."""
        try:
            all_c = ctx.cookies([
                "https://www.sciencedirect.com",
                "https://www.sciencedirectassets.com",
                "https://pdf.sciencedirectassets.com",
                "https://www.elsevier.com",
            ])
            d = {c["name"]: c["value"] for c in all_c}
            h = "; ".join(f"{k}={v}" for k, v in d.items())
            return d, h
        except Exception:
            return {}, ""

    def _get_cookies_via_browser_cookie3(self):
        """Read persisted cookies from disk (fallback; lacks session cookies)."""
        if not HAS_BROWSER_COOKIE3:
            return {}, ""
        try:
            all_cookies = {}
            for domain in (".sciencedirect.com", ".elsevier.com", ".sciencedirectassets.com"):
                jar = browser_cookie3.chrome(domain_name=domain)
                for c in jar:
                    all_cookies[c.name] = c.value
            h = "; ".join(f"{k}={v}" for k, v in all_cookies.items())
            return all_cookies, h
        except Exception as e:
            print(f"  [Warning] Failed to read Chrome cookies: {e}")
            return {}, ""

    # ── DevTools PDF download (primary method, no Playwright required) ────────

    def download_pdfs_devtools(self, results, output_dir, debug_port=9222):
        """
        Download PDFs via Chrome DevTools Protocol (pure websocket-client).

        How it works:
        1. Connect to a running Chrome debug instance (port 9222)
        2. For each paper, open the pdfft URL in a persistent tab
        3. Intercept PDF response bytes via Network/Fetch DevTools events
        4. Write directly to disk — no Save dialog, no Playwright dependency

        Prerequisite: Chrome must be logged in via your institutional account
        (CARSI / university SSO) on ScienceDirect.
        """
        try:
            import websocket
        except ImportError:
            print("[Error] websocket-client required: pip install websocket-client")
            return

        from urllib.request import Request as _Req, urlopen as _urlopen
        from urllib.parse import quote as _quote

        pdf_dir = os.path.join(output_dir, "pdfs")
        os.makedirs(pdf_dir, exist_ok=True)
        total = len(results)
        success = skip = fail = 0

        print(f"\n[DevTools PDF Download]  {total} papers → {pdf_dir}")

        chrome_was_fresh = not self._is_chrome_debug_ready()
        if chrome_was_fresh:
            print("  Debug port not ready. Launching Chrome automatically...")
            self._launch_chrome_with_debug()
            if not self._is_chrome_debug_ready():
                print("[Error] Chrome debug port unavailable. PDF download aborted.")
                return
        else:
            print("  Chrome debug port detected ✓")

        base_url = f"http://127.0.0.1:{debug_port}"

        def open_tab(url):
            req = _Req(f"{base_url}/json/new?{_quote(url, safe=':/?&=%')}", method="PUT")
            with _urlopen(req, timeout=20) as r:
                return json.loads(r.read())

        def close_tab(page_id):
            try:
                _urlopen(_Req(f"{base_url}/json/close/{page_id}"), timeout=10)
            except Exception:
                pass

        def _prompt_login():
            print("\n" + "=" * 60)
            print("  Please log in with your institutional account in Chrome:")
            print("  1. Open any ScienceDirect article")
            print("  2. Click 'View PDF' → log in via your university (SSO/CARSI)")
            print("  3. Confirm the PDF is visible (not a login page)")
            print("  4. Return here and press Enter to continue")
            print("=" * 60)
            try:
                input("  >>> Press Enter when login is complete: ")
            except EOFError:
                print("  [Non-interactive mode] Continuing automatically (ensure login is done).")

        def _check_institutional_access(pii: str) -> bool:
            """Returns True if institutional download access is confirmed."""
            test_pdf_url = f"{self.BASE_URL}/science/article/pii/{pii}/pdfft"
            tab = None
            try:
                tab = open_tab(test_pdf_url)
                time.sleep(10)
                ws = websocket.create_connection(
                    tab["webSocketDebuggerUrl"], timeout=30, suppress_origin=True)
                ws.send(json.dumps({"id": 1, "method": "Runtime.evaluate",
                                    "params": {"expression": "location.href", "returnByValue": True}}))
                final_url = ""
                for _ in range(15):
                    msg = json.loads(ws.recv())
                    if msg.get("id") == 1:
                        final_url = msg.get("result", {}).get("result", {}).get("value", "")
                        break
                ws.close()
                return "sciencedirectassets.com" in final_url or "pdf" in final_url.lower()
            except Exception as e:
                print(f"  [Warning] Access check error: {e}")
                return False
            finally:
                if tab:
                    close_tab(tab["id"])

        test_pii = next((a["pii"] for a in results if a.get("pii")), None)

        if chrome_was_fresh:
            _prompt_login()

        if test_pii:
            print("  Checking institutional access (navigating to PDF URL)...")
            has_access = _check_institutional_access(test_pii)
            if has_access:
                print("  Institutional access confirmed ✓")
            else:
                print("  Institutional download access not detected.")
                _prompt_login()
                print("  Re-checking access...")
                if _check_institutional_access(test_pii):
                    print("  Institutional access confirmed ✓")
                else:
                    print("  [Warning] Access still not confirmed. Attempting download anyway (may fail).")

        print()

        INTER_MIN         = 12
        INTER_MAX         = 22
        SESSION_BREAK_N   = 8
        SESSION_BREAK_T   = 150
        BLOCK_WAIT_1      = 270
        BLOCK_WAIT_2      = 420

        p_tab = open_tab("about:blank")

        def _tab_navigate(url, wait=5.0):
            try:
                ws2 = websocket.create_connection(
                    p_tab["webSocketDebuggerUrl"], timeout=30, suppress_origin=True)
                ws2.settimeout(2)
                mid = [0]

                def _s(method, params=None):
                    mid[0] += 1
                    ws2.send(json.dumps({"id": mid[0], "method": method,
                                        "params": params or {}}))

                _s("Page.enable")
                _s("Page.addScriptToEvaluateOnNewDocument", {"source": _STEALTH_JS})
                _s("Page.navigate", {"url": url})
                deadline2 = time.time() + wait + 10
                while time.time() < deadline2:
                    try:
                        msg = json.loads(ws2.recv())
                        if msg.get("method") == "Page.loadEventFired":
                            break
                    except Exception:
                        pass
                ws2.close()
            except Exception:
                pass
            time.sleep(wait)

        def _ensure_tab():
            nonlocal p_tab
            try:
                ws_test = websocket.create_connection(
                    p_tab["webSocketDebuggerUrl"], timeout=5, suppress_origin=True)
                ws_test.close()
            except Exception:
                try:
                    close_tab(p_tab["id"])
                except Exception:
                    pass
                p_tab = open_tab("about:blank")
                time.sleep(1)

        def _fetch_one(pii, pdf_url):
            _ensure_tab()
            article_url = f"{self.BASE_URL}/science/article/pii/{pii}"
            _tab_navigate(article_url, wait=random.uniform(4, 6))
            try:
                return _dt_capture_pdf(p_tab["webSocketDebuggerUrl"], pdf_url, timeout=45)
            except Exception as exc:
                return None, str(exc)

        downloads_since_break = 0

        try:
            for idx, article in enumerate(results, 1):
                pii = article.get("pii", "")
                title_short = (article.get("title") or "")[:55]

                if not pii:
                    print(f"  [{idx}/{total}] Skipped (no PII): {title_short}")
                    skip += 1
                    continue

                filename = self._make_pdf_filename(idx, article)
                filepath = os.path.join(pdf_dir, filename)

                if os.path.exists(filepath) and os.path.getsize(filepath) > 0:
                    print(f"  [{idx}/{total}] Already exists, skipping: {filename}")
                    skip += 1
                    continue

                pdf_url = article.get("pdf_url") or ""
                if not pdf_url or "pdfft" not in pdf_url:
                    pdf_url = f"{self.BASE_URL}/science/article/pii/{pii}/pdfft"

                pdf_bytes, note = _fetch_one(pii, pdf_url)

                if pdf_bytes is None and str(note).startswith("blocked:"):
                    note_low = str(note).lower()
                    is_captcha = any(s in note_low for s in CAPTCHA_SIGNALS)

                    if is_captcha:
                        print(f"\n  [{idx}/{total}] CAPTCHA detected — human verification required")
                        print("  ─────────────────────────────────────────────")
                        print("  Please switch to the Chrome window and complete verification:")
                        print("  · Check 'I'm not a robot' or solve the image challenge")
                        print("  · Return here and press Enter to continue")
                        print("  ─────────────────────────────────────────────")
                        try:
                            input("  >>> Press Enter after completing verification: ")
                        except EOFError:
                            time.sleep(30)
                        pdf_bytes, note = _fetch_one(pii, pdf_url)
                    else:
                        print(f"  [{idx}/{total}] Rate limited. Waiting {BLOCK_WAIT_1}s (~{BLOCK_WAIT_1//60} min) before retry...")
                        _tab_navigate("about:blank", wait=2)
                        time.sleep(BLOCK_WAIT_1)
                        pdf_bytes, note = _fetch_one(pii, pdf_url)

                        if pdf_bytes is None and str(note).startswith("blocked:"):
                            print(f"  [{idx}/{total}] Still rate limited. Waiting {BLOCK_WAIT_2}s (~{BLOCK_WAIT_2//60} min)...")
                            _tab_navigate("about:blank", wait=2)
                            time.sleep(BLOCK_WAIT_2)
                            pdf_bytes, note = _fetch_one(pii, pdf_url)

                if pdf_bytes and pdf_bytes[:4] == b"%PDF":
                    with open(filepath, "wb") as f:
                        f.write(pdf_bytes)
                    size_kb = len(pdf_bytes) // 1024
                    print(f"  [{idx}/{total}] ✓ {filename}  ({size_kb} KB)")
                    success += 1
                    downloads_since_break += 1
                else:
                    is_blocked = str(note).startswith("blocked:")
                    tag = "Blocked" if is_blocked else "PDF not captured"
                    print(f"  [{idx}/{total}] ✗ {tag}: {title_short[:40]}  ({str(note)[:80]})")
                    fail += 1

                if idx < total:
                    if downloads_since_break >= SESSION_BREAK_N:
                        print(f"  [Throttle] {downloads_since_break} consecutive downloads. Resting {SESSION_BREAK_T}s ...")
                        _tab_navigate("about:blank", wait=2)
                        time.sleep(SESSION_BREAK_T)
                        downloads_since_break = 0
                    else:
                        time.sleep(random.uniform(INTER_MIN, INTER_MAX))
        finally:
            close_tab(p_tab["id"])

        print(f"\n[Done] Success: {success}  Failed: {fail}  Skipped: {skip}")

    # ── Chrome CDP utilities ──────────────────────────────────────────────────

    CHROME_BIN = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
    CHROME_DBG_PROFILE = "/tmp/chrome_dbg_profile"
    CHROME_DBG_PORT = 9222

    def _is_chrome_debug_ready(self):
        """Check whether the Chrome debug port is available."""
        import urllib.request
        try:
            urllib.request.urlopen(
                f"http://127.0.0.1:{self.CHROME_DBG_PORT}/json/version", timeout=2)
            return True
        except Exception:
            return False

    def _launch_chrome_with_debug(self):
        """
        Launch Chrome in debug mode:
        1. Copy key browser state from the default Profile to a temp directory
        2. Start Chrome with --user-data-dir + --remote-debugging-port
        3. Wait up to 40s for the debug port to become ready
        Returns the Popen object, or None on failure.
        """
        import subprocess, shutil

        default_profile = os.path.expanduser(
            "~/Library/Application Support/Google/Chrome/Default"
        )
        tmp_default = os.path.join(self.CHROME_DBG_PROFILE, "Default")
        os.makedirs(tmp_default, exist_ok=True)

        files_to_copy = (
            "Cookies", "Cookies-journal", "Preferences", "Secure Preferences",
            "History", "Visited Links", "Web Data", "Login Data",
        )
        dirs_to_copy = (
            "Network", "Local Storage", "Session Storage",
            "IndexedDB", "SharedStorage", "WebStorage",
        )

        for fname in files_to_copy:
            src = os.path.join(default_profile, fname)
            dst = os.path.join(tmp_default, fname)
            if os.path.exists(src):
                try:
                    shutil.copy2(src, dst)
                except Exception as e:
                    print(f"  [Warning] Could not copy {fname}: {e} (continuing)")

        for dname in dirs_to_copy:
            src = os.path.join(default_profile, dname)
            dst = os.path.join(tmp_default, dname)
            if os.path.exists(src):
                try:
                    shutil.copytree(src, dst, dirs_exist_ok=True)
                except Exception as e:
                    print(f"  [Warning] Could not sync {dname}: {e} (continuing)")

        for lock in ("SingletonLock", "SingletonCookie", "SingletonSocket"):
            lf = os.path.join(self.CHROME_DBG_PROFILE, lock)
            try:
                os.remove(lf)
            except FileNotFoundError:
                pass

        cmd = [
            self.CHROME_BIN,
            f"--remote-debugging-port={self.CHROME_DBG_PORT}",
            f"--remote-allow-origins=http://127.0.0.1:{self.CHROME_DBG_PORT}",
            f"--user-data-dir={self.CHROME_DBG_PROFILE}",
            "--disable-blink-features=AutomationControlled",
            "--no-first-run",
            "--no-default-browser-check",
        ]

        log_path = "/tmp/chrome_debug.log"
        proc = None
        try:
            with open(log_path, "w") as log_f:
                proc = subprocess.Popen(cmd, stdout=log_f, stderr=subprocess.STDOUT)
            print(f"  Chrome launched (PID {proc.pid}). Waiting for debug port...")
        except FileNotFoundError:
            print(f"  [Error] Chrome not found at: {self.CHROME_BIN}")
            return None

        for i in range(40):
            time.sleep(1)
            if self._is_chrome_debug_ready():
                print(f"  Chrome debug port ready ({i+1}s) ✓")
                return proc
            if (i + 1) % 5 == 0:
                print(f"  Waiting for Chrome... ({i+1}s)")

        print(f"  [Warning] Chrome not ready after 40s. Check log: cat {log_path}")
        return None

    def open_chrome_for_login(self, target_url=None, keep_page_open=True):
        """Open a debug Chrome window so the user can complete institutional login."""
        target_url = target_url or self.BASE_URL
        if not self._is_chrome_debug_ready():
            print("  Launching debug Chrome...")
            self._launch_chrome_with_debug()

        if not self._is_chrome_debug_ready():
            print("\n[Error] Chrome debug port unavailable. Check: cat /tmp/chrome_debug.log")
            return False

        try:
            from playwright.sync_api import sync_playwright
        except ImportError:
            print("[Error] Playwright required: pip install playwright")
            return False

        with sync_playwright() as p:
            browser = p.chromium.connect_over_cdp(
                f"http://127.0.0.1:{self.CHROME_DBG_PORT}"
            )
            ctx = browser.contexts[0] if browser.contexts else browser.new_context()
            page = ctx.new_page()
            page.add_init_script(self._STEALTH_SCRIPT)
            try:
                page.goto(target_url, timeout=30000, wait_until="domcontentloaded")
            except Exception:
                pass

            print("\n" + "=" * 60)
            print("  Chrome is open. Please:")
            print("  1. Log in to ScienceDirect via your institutional account")
            print("  2. Confirm you can open a full-text article")
            print("  3. Return here and press Enter to continue")
            print("=" * 60)
            input()

            if not keep_page_open:
                try:
                    page.close()
                except Exception:
                    pass
        return True

    # ── Playwright-based PDF download (alternative method) ───────────────────

    _STEALTH_SCRIPT = """
        Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
        delete navigator.__proto__.webdriver;
        window.chrome = window.chrome || { runtime: {} };
        Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3,4,5]});
        Object.defineProperty(navigator, 'languages', {get: () => ['en-US','en','zh-CN']});
    """

    def _check_sd_access(self, ctx, test_pii=None):
        """
        Check whether the current Chrome session has institutional full-text access.
        Returns True if access is confirmed.
        """
        check_page = ctx.new_page()
        try:
            check_page.add_init_script(self._STEALTH_SCRIPT)
            if test_pii:
                url = f"{self.BASE_URL}/science/article/pii/{test_pii}"
            else:
                url = self.BASE_URL
            check_page.goto(url, timeout=25000, wait_until="domcontentloaded")
            time.sleep(2)
            current_url = check_page.url
            if "sciencedirect.com" not in current_url:
                return False
            if test_pii and "/abs/" in current_url:
                return False
            body_text = check_page.inner_text("body")
            if "Sign in" in body_text and "Remote access" not in body_text:
                return False
            return True
        except Exception:
            return False
        finally:
            check_page.close()

    def _wait_for_login(self, ctx, test_pii=None):
        """Prompt the user to complete institutional login in the Chrome window."""
        login_page = ctx.new_page()
        try:
            login_page.add_init_script(self._STEALTH_SCRIPT)
            target = (
                f"{self.BASE_URL}/science/article/pii/{test_pii}"
                if test_pii else self.BASE_URL
            )
            login_page.goto(target, timeout=20000, wait_until="domcontentloaded")
        except Exception:
            pass

        print("\n" + "="*60)
        print("  Chrome may have redirected to your institutional login page.")
        print("  Please complete login in Chrome:")
        print("  · If shown a university login page: enter your credentials")
        print("  · If on ScienceDirect: click Sign in → Access through your institution")
        print("  Return here and press Enter when done.")
        print("="*60)
        input()

        try:
            login_page.close()
        except Exception:
            pass

    def download_pdfs_via_chrome(self, results, output_dir):
        """
        Download PDFs by driving Chrome via CDP + Playwright.
        The script will:
          1. Copy persistent cookies from default Chrome Profile
          2. Launch Chrome in debug mode
          3. Check/prompt for institutional access (one-time)
          4. For each paper: load article page, click View PDF, download via curl_cffi
        """
        try:
            from playwright.sync_api import sync_playwright
        except ImportError:
            print("[Error] Playwright required: pip install playwright && playwright install chromium")
            return

        pdf_dir = os.path.join(output_dir, "pdfs")
        os.makedirs(pdf_dir, exist_ok=True)
        total = len(results)
        success = skip = fail = 0

        print(f"\n[Chrome PDF Download]  {total} papers → {pdf_dir}")

        chrome_proc = None
        if not self._is_chrome_debug_ready():
            print("  Debug port not ready. Launching Chrome...")
            chrome_proc = self._launch_chrome_with_debug()
            if not self._is_chrome_debug_ready():
                print("\n[Error] Chrome debug port still unavailable.")
                print("  Check log: cat /tmp/chrome_debug.log")
                return
        else:
            print("  Chrome debug port detected ✓")

        with sync_playwright() as p:
            try:
                browser = p.chromium.connect_over_cdp(
                    f"http://127.0.0.1:{self.CHROME_DBG_PORT}"
                )
            except Exception as e:
                print(f"\n[Error] Failed to connect to Chrome: {e}")
                return

            ctx = browser.contexts[0] if browser.contexts else browser.new_context()
            print("  Connected to Chrome ✓")

            test_pii = next((a["pii"] for a in results if a.get("pii")), None)
            print("  Checking institutional access...")
            if not self._check_sd_access(ctx, test_pii):
                self._wait_for_login(ctx)
                if not self._check_sd_access(ctx, test_pii):
                    print("\n[Warning] Institutional access not confirmed. Attempting download anyway.")
                else:
                    print("  Institutional access confirmed ✓")
            else:
                print("  Institutional access confirmed ✓")

            print()

            browser_cookies = ctx.cookies([
                "https://www.sciencedirect.com",
                "https://www.sciencedirectassets.com",
                "https://pdf.sciencedirectassets.com",
                "https://www.elsevier.com",
            ])
            cookie_header = "; ".join(
                f"{c['name']}={c['value']}" for c in browser_cookies
            )
            dl_session = curl_requests.Session(impersonate="chrome124")
            _DL_HEADERS = {
                "Cookie": cookie_header,
                "User-Agent": (
                    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/146.0.0.0 Safari/537.36"
                ),
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,"
                          "image/avif,image/webp,*/*;q=0.8",
                "Accept-Language": "en-US,en;q=0.9",
                "sec-fetch-dest": "document",
                "sec-fetch-mode": "navigate",
            }

            for idx, article in enumerate(results, 1):
                pii = article.get("pii", "")
                title_short = article.get("title", "")[:55]
                if not pii:
                    print(f"  [{idx}/{total}] Skipped (no PII): {title_short}")
                    skip += 1
                    continue

                filename = self._make_pdf_filename(idx, article)
                filepath = os.path.join(pdf_dir, filename)

                if os.path.exists(filepath):
                    print(f"  [{idx}/{total}] Already exists, skipping: {filename}")
                    skip += 1
                    continue

                article_url = f"{self.BASE_URL}/science/article/pii/{pii}"

                # Strategy A: direct HTTP download via search result pdf_url
                pdf_url_from_search = article.get("pdf_url", "")
                downloaded = False

                if pdf_url_from_search and "pdfft" in pdf_url_from_search:
                    try:
                        headers_a = dict(_DL_HEADERS)
                        headers_a["Referer"] = article_url
                        headers_a["sec-fetch-site"] = "same-origin"
                        resp = dl_session.get(
                            pdf_url_from_search, headers=headers_a,
                            allow_redirects=True, timeout=60,
                        )
                        ct = resp.headers.get("content-type", "")
                        if "pdf" in ct.lower() or resp.content[:4] == b"%PDF":
                            with open(filepath, "wb") as f:
                                f.write(resp.content)
                            size_kb = os.path.getsize(filepath) // 1024
                            print(f"  [{idx}/{total}] ✓ {filename}  ({size_kb} KB)  [direct]")
                            success += 1
                            downloaded = True
                    except Exception:
                        pass

                if downloaded:
                    if idx < total:
                        time.sleep(random.uniform(3, 6))
                    continue

                # Strategy B: CDP navigate + click View PDF → curl_cffi download
                page = ctx.new_page()
                try:
                    page.add_init_script(self._STEALTH_SCRIPT)
                    page.goto(article_url, timeout=30000, wait_until="networkidle")
                    time.sleep(1)

                    try:
                        page.wait_for_selector('a[href*="pdfft"]', timeout=8000)
                    except Exception:
                        page_text = ""
                        try:
                            page_text = page.inner_text("body")[:300]
                        except Exception:
                            pass
                        if "problem providing" in page_text or "crasolve" in page.url:
                            print(f"  [{idx}/{total}] ✗ Bot detection on article page")
                            page.close()
                            fail += 1
                            continue
                        print(f"  [{idx}/{total}] ✗ No PDF button (no access?): {title_short[:40]}")
                        fail += 1
                        page.close()
                        continue

                    pdf_assets_url = None
                    try:
                        with page.expect_popup(timeout=15000) as popup_info:
                            page.click('a[href*="pdfft"]')
                        popup = popup_info.value
                        pdf_assets_url = popup.url
                        popup.close()
                    except Exception:
                        try:
                            pdf_assets_url = page.eval_on_selector(
                                'a[href*="pdfft"]', 'el => el.href'
                            )
                        except Exception:
                            pass
                    page.close()

                    if not pdf_assets_url:
                        raise RuntimeError("Could not obtain PDF link")

                    headers_b = dict(_DL_HEADERS)
                    headers_b["Referer"] = article_url
                    headers_b["sec-fetch-site"] = "cross-site"
                    resp = dl_session.get(
                        pdf_assets_url, headers=headers_b,
                        allow_redirects=True, timeout=60,
                    )
                    ct = resp.headers.get("content-type", "")
                    if "pdf" in ct.lower() or resp.content[:4] == b"%PDF":
                        with open(filepath, "wb") as f:
                            f.write(resp.content)
                        size_kb = os.path.getsize(filepath) // 1024
                        print(f"  [{idx}/{total}] ✓ {filename}  ({size_kb} KB)  [CDP+direct]")
                        success += 1
                    else:
                        print(f"  [{idx}/{total}] ✗ Non-PDF response ({ct[:30]}): {title_short[:35]}")
                        fail += 1

                except Exception as e:
                    print(f"  [{idx}/{total}] ✗ Failed: {title_short[:40]}  ({e})")
                    fail += 1
                    try:
                        page.close()
                    except Exception:
                        pass

                if idx < total:
                    time.sleep(random.uniform(3, 6))

        print(f"\n[Done] Success: {success}  Failed: {fail}  Skipped: {skip}")
        if chrome_proc is not None:
            print(f"\n  Note: Debug Chrome is still running (PID {chrome_proc.pid})")
            print("  To close it: pkill -f 'Google Chrome'")


# ──────────────────────────────────────────────────────────────────────────────
# Interactive wizard
# ──────────────────────────────────────────────────────────────────────────────

def _input_int(prompt, default):
    raw = input(f"{prompt} [default {default}]: ").strip()
    try:
        return int(raw) if raw else default
    except ValueError:
        return default

def _input_optional(prompt):
    raw = input(f"{prompt} [leave blank to skip]: ").strip()
    return raw if raw else None

def interactive_mode():
    print("=" * 60)
    print("  ScienceDirect Paper Scraper v2.0")
    print("=" * 60)

    # ── Institutional cookies ─────────────────────────────────────────────────
    print("\n[Institutional Account]")
    print("  1. Auto-read from Chrome (recommended)")
    print("  2. Specify a cookie JSON file")
    print("  3. Skip — run as guest")
    cookie_choice = input("  Choose [1/2/3, default 1]: ").strip() or "1"

    cookies_file = None
    use_browser_cookies = False
    if cookie_choice == "1":
        use_browser_cookies = True
    elif cookie_choice == "2":
        cookies_file = _input_optional("  Cookie file path (e.g. cookies.json)")

    scraper = ScienceDirectScraper(cookies_file=cookies_file, use_browser_cookies=use_browser_cookies)

    # ── Search mode ───────────────────────────────────────────────────────────
    print("\n[Search Mode]")
    modes = {
        "1": ("keyword",         "Search by keyword"),
        "2": ("journal",         "Browse by journal name"),
        "3": ("journal_keyword", "Keyword search within a journal"),
        "4": ("author",          "Search by author name"),
        "5": ("issn",            "Search by journal ISSN"),
        "6": ("advanced",        "Advanced search (combine criteria)"),
    }
    for k, (_, desc) in modes.items():
        print(f"  {k}. {desc}")
    choice = input("  Choose [1-6]: ").strip()
    mode = modes.get(choice, ("keyword", ""))[0]

    # ── Search parameters ─────────────────────────────────────────────────────
    print("\n[Search Parameters]")
    query      = _input_optional("  Keywords (e.g. machine learning)")
    journal    = _input_optional("  Journal name (e.g. Energy)")
    author     = _input_optional("  Author name (e.g. Zhang Wei)")
    issn       = _input_optional("  Journal ISSN (e.g. 0360-5442)")
    date_range = _input_optional("  Year range (e.g. 2020-2024)")
    print("  Article type: FLA=Full Article  REV=Review  SCO=Short Comm.  (blank=all)")
    article_type = _input_optional("  Article type")
    count      = _input_int("  Max papers to fetch", 50)
    sort_raw   = input("  Sort by: relevance / date [default relevance]: ").strip()
    sort_by    = sort_raw if sort_raw in ("relevance", "date") else "relevance"

    # ── Execute search ────────────────────────────────────────────────────────
    results = []
    if mode == "keyword":
        results = scraper.search_by_keyword(query or "", count, sort_by, date_range, article_type)
    elif mode == "journal":
        results = scraper.search_by_journal(journal or "", count, sort_by, date_range)
    elif mode == "journal_keyword":
        results = scraper.search_by_journal_keyword(journal or "", query or "", count, sort_by, date_range)
    elif mode == "author":
        results = scraper.search_by_author(author or "", count, sort_by)
    elif mode == "issn":
        results = scraper.search_by_issn(issn or "", count, sort_by, date_range)
    elif mode == "advanced":
        results = scraper.search_advanced(query, journal, author, issn,
                                          date_range, article_type,
                                          count=count, sort_by=sort_by)

    if not results:
        print("\nNo results found. Check your parameters or cookies.")
        return

    # ── Output format ─────────────────────────────────────────────────────────
    print("\n[Output Format]")
    fmt_raw = input("  Format: xlsx / csv / json / all [default xlsx]: ").strip().lower()
    fmt = fmt_raw if fmt_raw in ("xlsx", "csv", "json", "all") else "xlsx"

    print("\n[PDF Download]")
    print("  1. Download PDFs (recommended)")
    print("     Reads cookies from Chrome and downloads via direct HTTP.")
    print("     Prerequisite: Chrome logged in via institutional account on ScienceDirect.")
    print("  2. Skip — save metadata only")
    dl_choice = input("  Choose [1/2, default 2]: ").strip() or "2"
    download_pdfs = (dl_choice == "1")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    parts = [mode]
    if query:   parts.append(query.replace(" ", "_")[:20])
    if journal: parts.append(journal.replace(" ", "_")[:20])
    base = "_".join(parts) + f"_{timestamp}"
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results", base)

    if fmt in ("xlsx", "all"):
        scraper.save_to_xlsx(results, base + ".xlsx", output_dir)
    if fmt in ("csv", "all"):
        scraper.save_to_csv(results, base + ".csv", output_dir)
    if fmt in ("json", "all"):
        scraper.save_to_json(results, base + ".json", output_dir)
    if download_pdfs:
        scraper.download_pdfs_devtools(results, output_dir)


# ──────────────────────────────────────────────────────────────────────────────
# CLI entry point
# ──────────────────────────────────────────────────────────────────────────────

def build_parser():
    parser = argparse.ArgumentParser(
        description="ScienceDirect Paper Scraper v2.0",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python sd_scraper_en.py --interactive
  python sd_scraper_en.py --open-browser-login
  python sd_scraper_en.py -m keyword -q "machine learning" -n 100 --browser-cookies --format xlsx --download-pdfs
  python sd_scraper_en.py -m journal -j "Energy" -n 50 --browser-cookies --sort date --format xlsx
  python sd_scraper_en.py -m journal_keyword -j "Renewable Energy" -q "solar cell" -n 50 --browser-cookies --download-pdfs
  python sd_scraper_en.py -m author -a "Zhang Wei" -n 30 --browser-cookies --format all
  python sd_scraper_en.py -m advanced -q "deep learning" --date 2021-2024 --type REV -n 50 --browser-cookies --download-pdfs
        """,
    )
    parser.add_argument("--interactive", action="store_true", help="Launch interactive wizard")
    parser.add_argument("--open-browser-login", action="store_true",
                        help="Open Chrome for manual institutional login before scraping")
    parser.add_argument("--login-only", action="store_true",
                        help="Open Chrome for login only, without running a search")
    parser.add_argument("-m", "--mode",
                        choices=["keyword", "journal", "journal_keyword",
                                 "author", "issn", "advanced"],
                        help="Search mode")
    parser.add_argument("-q", "--query",   help="Search keywords (supports AND/OR/NOT)")
    parser.add_argument("-j", "--journal", help="Journal name")
    parser.add_argument("-a", "--author",  help="Author name")
    parser.add_argument("--issn",          help="Journal ISSN")
    parser.add_argument("-n", "--count",   type=int, default=50, help="Max papers to fetch (default 50)")
    parser.add_argument("--date",          help="Year range, e.g. 2020-2024")
    parser.add_argument("--sort",          choices=["relevance", "date"], default="relevance")
    parser.add_argument("--type",  dest="article_type",
                        choices=["FLA", "REV", "SCO", "EDB", "ERR", "COR"],
                        help="Article type: FLA=Full Article / REV=Review / SCO=Short Comm.")
    parser.add_argument("--open-access",   action="store_true", help="Fetch open-access articles only")
    parser.add_argument("--browser-cookies", dest="browser_cookies", action="store_true",
                        help="Auto-read cookies from local Chrome")
    parser.add_argument("--cookies",       help="Path to a cookie JSON file")
    parser.add_argument("--format",        choices=["xlsx", "csv", "json", "all"], default="xlsx")
    parser.add_argument("--download-pdfs", action="store_true",
                        help="Download PDFs after saving the paper list")
    parser.add_argument("--output",        help="Output directory (default: ./results/)")
    parser.add_argument("--filename",      help="Custom output filename (without extension)")
    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()

    if args.interactive or len(sys.argv) == 1:
        interactive_mode()
        return

    if args.login_only and args.mode:
        print("Note: --login-only ignores search parameters and only opens Chrome for login.")
    if args.login_only and not args.open_browser_login:
        args.open_browser_login = True

    if not args.mode and not args.login_only:
        parser.print_help()
        return

    scraper = ScienceDirectScraper(
        cookies_file=args.cookies,
        use_browser_cookies=args.browser_cookies
    )

    if args.open_browser_login:
        ok = scraper.open_chrome_for_login(keep_page_open=True)
        if not ok:
            return
        if args.login_only:
            print("\nChrome will remain open with your login session.")
            print("Now run your actual scrape command, e.g.:")
            print('python sd_scraper_en.py -m keyword -q "machine learning" -n 20 --browser-cookies --format xlsx --download-pdfs')
            return

    results = []
    if args.mode == "keyword":
        if not args.query:
            print("Error: keyword mode requires -q"); return
        results = scraper.search_by_keyword(
            args.query, args.count, args.sort, args.date, args.article_type)
    elif args.mode == "journal":
        if not args.journal:
            print("Error: journal mode requires -j"); return
        results = scraper.search_by_journal(
            args.journal, args.count, args.sort, args.date)
    elif args.mode == "journal_keyword":
        if not args.journal or not args.query:
            print("Error: journal_keyword mode requires -j and -q"); return
        results = scraper.search_by_journal_keyword(
            args.journal, args.query, args.count, args.sort, args.date)
    elif args.mode == "author":
        if not args.author:
            print("Error: author mode requires -a"); return
        results = scraper.search_by_author(args.author, args.count, args.sort)
    elif args.mode == "issn":
        if not args.issn:
            print("Error: issn mode requires --issn"); return
        results = scraper.search_by_issn(
            args.issn, args.count, args.sort, args.date)
    elif args.mode == "advanced":
        results = scraper.search_advanced(
            query=args.query, journal=args.journal, author=args.author,
            issn=args.issn, date_range=args.date, article_type=args.article_type,
            open_access_only=args.open_access,
            count=args.count, sort_by=args.sort)

    if not results:
        print("\nNo results found.")
        return

    if args.filename:
        base = args.filename
    else:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        parts = [args.mode]
        if args.query:   parts.append(args.query.replace(" ", "_")[:20])
        if args.journal: parts.append(args.journal.replace(" ", "_")[:20])
        base = "_".join(parts) + f"_{timestamp}"

    output_dir = args.output or os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "results")

    if args.format in ("xlsx", "all"):
        scraper.save_to_xlsx(results, base + ".xlsx", output_dir)
    if args.format in ("csv", "all"):
        scraper.save_to_csv(results, base + ".csv", output_dir)
    if args.format in ("json", "all"):
        scraper.save_to_json(results, base + ".json", output_dir)
    if args.download_pdfs:
        scraper.download_pdfs_devtools(results, output_dir)


if __name__ == "__main__":
    main()
