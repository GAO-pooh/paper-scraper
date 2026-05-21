#!/usr/bin/env python
"""
INFORMS PubsOnLine Paper Scraper v1.1
======================================
Based on the Atypon/Literatum platform. Uses curl_cffi to mimic Chrome TLS fingerprint.
Supports keyword search, journal browsing, and volume/issue table-of-contents scraping.
Results saved as CSV / JSON / XLSX. Supports bulk PDF download (member account, direct HTTP).

New: --chrome-login mode
  Launches a real Chrome window → you log in manually → script auto-extracts session
  cookies → scraping begins. No manual cookie export needed. Most reliable method.

INFORMS Journal Codes
---------------------
  mnsc   — Management Science
  opre   — Operations Research
  ijoc   — INFORMS Journal on Computing
  mksc   — Marketing Science
  msom   — Manufacturing & Service Operations Management
  trsc   — Transportation Science
  orsc   — Organization Science
  isre   — Information Systems Research
  deca   — Decision Analysis
  stsy   — Stochastic Systems
  ijds   — INFORMS Journal on Data Science
  serv   — Service Science
  inte   — INFORMS Journal on Applied Analytics (formerly Interfaces)
  educ   — INFORMS Transactions on Education

Quick Start
-----------
  # Recommended: log in via browser first, then use browser-cookies
  python informs_scraper_en.py -m keyword -q "machine learning" -n 100 --browser-cookies

  # Or provide member credentials directly
  python informs_scraper_en.py -m keyword -q "supply chain" -n 50 --member YOUR_ID --password YOUR_PWD

  # Browse latest articles in a journal
  python informs_scraper_en.py -m journal -j mnsc -n 200 --browser-cookies

  # Browse a specific volume/issue
  python informs_scraper_en.py -m toc -j mnsc -v 71 -i 3 --browser-cookies

  # Search + download PDFs
  python informs_scraper_en.py -m keyword -q "inventory" -n 30 --browser-cookies --download-pdf

  # ★ Best option: pop up Chrome, log in manually
  python informs_scraper_en.py -m keyword -q "machine learning" -n 50 --chrome-login --download-pdf

Install Dependencies
--------------------
  pip install curl_cffi browser-cookie3 beautifulsoup4 lxml openpyxl
"""

import json
import csv
import os
import re
import sys
import time
import random
import argparse
from datetime import datetime
from urllib.parse import urlencode, quote

from curl_cffi import requests as curl_requests

try:
    from bs4 import BeautifulSoup
    HAS_BS4 = True
except ImportError:
    HAS_BS4 = False

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


# ─────────────────────────────────────────────────────────────────────────────
# Journal code → full name mapping
# ─────────────────────────────────────────────────────────────────────────────
JOURNAL_NAMES = {
    "mnsc":  "Management Science",
    "opre":  "Operations Research",
    "ijoc":  "INFORMS Journal on Computing",
    "mksc":  "Marketing Science",
    "msom":  "Manufacturing & Service Operations Management",
    "trsc":  "Transportation Science",
    "orsc":  "Organization Science",
    "isre":  "Information Systems Research",
    "deca":  "Decision Analysis",
    "stsy":  "Stochastic Systems",
    "ijds":  "INFORMS Journal on Data Science",
    "serv":  "Service Science",
    "inte":  "INFORMS Journal on Applied Analytics",
    "educ":  "INFORMS Transactions on Education",
    "nets":  "INFORMS Journal on Optimization",
}

PAGE_SIZE = 20

_UA = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)


# ─────────────────────────────────────────────────────────────────────────────
# Main scraper class
# ─────────────────────────────────────────────────────────────────────────────

class INFORMSScraper:
    BASE_URL = "https://pubsonline.informs.org"

    LOGIN_PAGE  = "/literatumuserslogin"
    LOGIN_POST  = "/action/doLogin"
    SEARCH_URL  = "/action/doSearch"
    TOC_URL     = "/toc/{journal}/{vol}/{issue}"
    JOURNAL_URL = "/loi/{journal}"

    CHROME_BIN         = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
    CHROME_DBG_PORT    = 9222
    CHROME_DBG_PROFILE = "/tmp/chrome_informs_profile"

    def __init__(
        self,
        member_id=None,
        password=None,
        cookies_file=None,
        use_browser_cookies=False,
        use_chrome_login=False,
        delay_range=(2, 5),
    ):
        self.session = curl_requests.Session(impersonate="chrome124")
        self.session.headers.update({
            "User-Agent": _UA,
            "Accept-Language": "en-US,en;q=0.9",
        })
        self.delay_range = delay_range
        self._cookie_str = ""

        if use_chrome_login:
            self._chrome_login_flow()
        elif use_browser_cookies:
            self._load_browser_cookies()
        elif cookies_file:
            self._load_cookies_file(cookies_file)
        elif member_id and password:
            self._login(member_id, password)

    # ── Chrome debug mode ─────────────────────────────────────────────────────

    def _is_chrome_debug_ready(self) -> bool:
        import urllib.request
        try:
            urllib.request.urlopen(
                f"http://127.0.0.1:{self.CHROME_DBG_PORT}/json/version", timeout=2)
            return True
        except Exception:
            return False

    def _launch_chrome_with_debug(self):
        """
        Launch Chrome in debug mode, reusing the default Profile (saved
        passwords, bookmarks, history).
        Returns Popen object, or None on failure.
        """
        import subprocess
        import shutil

        default_profile = os.path.expanduser(
            "~/Library/Application Support/Google/Chrome/Default"
        )
        tmp_default = os.path.join(self.CHROME_DBG_PROFILE, "Default")
        os.makedirs(tmp_default, exist_ok=True)

        for fname in ("Cookies", "Cookies-journal", "Preferences",
                      "Login Data", "Web Data"):
            src = os.path.join(default_profile, fname)
            dst = os.path.join(tmp_default, fname)
            if os.path.exists(src) and not os.path.exists(dst):
                try:
                    shutil.copy2(src, dst)
                except Exception:
                    pass

        for lock in ("SingletonLock", "SingletonCookie", "SingletonSocket"):
            try:
                os.remove(os.path.join(self.CHROME_DBG_PROFILE, lock))
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

        try:
            proc = subprocess.Popen(
                cmd,
                stdout=open("/tmp/chrome_informs.log", "w"),
                stderr=subprocess.STDOUT,
            )
            print(f"  Chrome launched (PID {proc.pid}). Waiting for debug port...")
        except FileNotFoundError:
            print(f"  [Error] Chrome not found at: {self.CHROME_BIN}")
            return None

        for i in range(40):
            time.sleep(1)
            if self._is_chrome_debug_ready():
                print(f"  Debug port ready ({i + 1}s) ✓")
                return proc
            if (i + 1) % 5 == 0:
                print(f"  Waiting for Chrome... ({i + 1}s)")

        print("  [Warning] Not ready after 40s. Check log: cat /tmp/chrome_informs.log")
        return None

    def _extract_cookies_via_cdp(self) -> str:
        """
        Extract pubsonline.informs.org cookies from running Chrome via CDP websocket.
        Returns a formatted Cookie header string. No Playwright required.
        """
        import urllib.request
        try:
            with urllib.request.urlopen(
                f"http://127.0.0.1:{self.CHROME_DBG_PORT}/json/list", timeout=5
            ) as r:
                pages = json.loads(r.read())
        except Exception as e:
            print(f"  [Error] Cannot retrieve CDP tab list: {e}")
            return ""

        ws_url = ""
        for page in pages:
            if "informs" in page.get("url", "").lower():
                ws_url = page.get("webSocketDebuggerUrl", "")
                break
        if not ws_url:
            for page in pages:
                if page.get("webSocketDebuggerUrl"):
                    ws_url = page["webSocketDebuggerUrl"]
                    break

        if not ws_url:
            print("  [Error] No available CDP tab found")
            return ""

        try:
            import websocket as _ws
        except ImportError:
            print("  [Note] websocket-client not installed. Falling back to browser_cookie3.")
            return self._load_browser_cookies(silent=True)

        try:
            ws = _ws.create_connection(ws_url, timeout=10, suppress_origin=True)
            ws.send(json.dumps({
                "id": 1,
                "method": "Network.getCookies",
                "params": {"urls": [
                    "https://pubsonline.informs.org",
                    "https://informs.org",
                ]}
            }))
            cookies = {}
            deadline = time.time() + 8
            while time.time() < deadline:
                ws.settimeout(max(0.3, deadline - time.time()))
                try:
                    msg = json.loads(ws.recv())
                except Exception:
                    break
                if msg.get("id") == 1:
                    for c in msg.get("result", {}).get("cookies", []):
                        cookies[c["name"]] = c["value"]
                    break
            ws.close()

            if cookies:
                cookie_str = "; ".join(f"{k}={v}" for k, v in cookies.items())
                print(f"  Extracted {len(cookies)} cookies via CDP ✓")
                return cookie_str
            else:
                print("  [Warning] CDP returned 0 cookies (not logged in yet?)")
                return ""
        except Exception as e:
            print(f"  [Warning] CDP cookie extraction failed: {e}")
            return ""

    def _open_informs_tab(self):
        """Open the INFORMS login page in the debug Chrome."""
        import urllib.request
        from urllib.parse import quote as _q
        login_url = self.BASE_URL + self.LOGIN_PAGE
        try:
            req = urllib.request.Request(
                f"http://127.0.0.1:{self.CHROME_DBG_PORT}/json/new?{_q(login_url, safe=':/?&=%')}",
                method="PUT",
            )
            urllib.request.urlopen(req, timeout=10)
        except Exception as e:
            print(f"  [Warning] Could not open new tab: {e}")

    def _chrome_login_flow(self):
        """
        Full Chrome login flow:
        1. Launch debug Chrome (reuses your default Profile)
        2. Open INFORMS login page
        3. Wait for you to log in manually, then press Enter
        4. Extract session cookies via CDP
        5. Set cookies on the session
        """
        print("\n[Chrome Login]  Launching Chrome — please log in to INFORMS in the window...")

        if not self._is_chrome_debug_ready():
            proc = self._launch_chrome_with_debug()
            if not self._is_chrome_debug_ready():
                print("  [Error] Chrome failed to start. Check: cat /tmp/chrome_informs.log")
                return
        else:
            print("  Existing debug Chrome detected ✓")

        self._open_informs_tab()
        time.sleep(1.5)

        print()
        print("=" * 60)
        print("  Chrome has opened pubsonline.informs.org.")
        print()
        print("  Please:")
        print("  1. Log in with your member ID and password")
        print("  2. Confirm your name appears or you are on the journal homepage")
        print("  3. Return here and press Enter to continue")
        print("=" * 60)

        try:
            input("\n  >>> Press Enter when login is complete: ")
        except EOFError:
            print("  [Non-interactive mode] Waiting 30s then continuing...")
            time.sleep(30)

        cookie_str = self._extract_cookies_via_cdp()
        if cookie_str:
            self._cookie_str = cookie_str
            self.session.headers["Cookie"] = self._cookie_str
            print("  [✓] Cookies set. Warming up session...")
            self._warmup()
            print("  [✓] Ready. Starting scrape.")
        else:
            print("  [!] CDP extraction failed. Trying browser_cookie3 fallback...")
            self._load_browser_cookies()
            if self._cookie_str:
                self._warmup()

    # ── Authentication ────────────────────────────────────────────────────────

    def _load_browser_cookies(self, silent=False) -> str:
        """
        Read persistent pubsonline.informs.org cookies from local Chrome.
        Note: browser_cookie3 only reads disk cookies, not session cookies.
        Use --chrome-login for the most complete cookie set.
        """
        if not HAS_BROWSER_COOKIE3:
            if not silent:
                print("[Error] browser-cookie3 not installed: pip install browser-cookie3")
            return ""
        try:
            cookies = {}
            for domain in (".informs.org", "pubsonline.informs.org"):
                jar = browser_cookie3.chrome(domain_name=domain)
                for c in jar:
                    cookies[c.name] = c.value
            if cookies:
                self._cookie_str = "; ".join(f"{k}={v}" for k, v in cookies.items())
                self.session.headers["Cookie"] = self._cookie_str
                if not silent:
                    print(f"[Info] Loaded {len(cookies)} cookies from Chrome")
                    print("  Tip: If you encounter 403 errors, use --chrome-login instead.")
                return self._cookie_str
            else:
                if not silent:
                    print("[Warning] No informs.org cookies found in Chrome.")
                    print("  Please log in to pubsonline.informs.org in Chrome first, or use --chrome-login.")
                return ""
        except Exception as e:
            if not silent:
                print(f"[Warning] Failed to read Chrome cookies: {e}")
                print("         On macOS, a Keychain permission dialog may appear — click Allow.")
            return ""

    def _load_cookies_file(self, path):
        """Load cookies from a JSON file (Netscape or dict format)."""
        if not os.path.exists(path):
            print(f"[Warning] Cookie file not found: {path}")
            return
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            cookies = {c.get("name", ""): c.get("value", "") for c in data if c.get("name")}
        elif isinstance(data, dict):
            cookies = data
        else:
            print("[Warning] Unsupported cookie file format")
            return
        self._cookie_str = "; ".join(f"{k}={v}" for k, v in cookies.items() if k and v)
        self.session.headers["Cookie"] = self._cookie_str
        print(f"[Info] Loaded {len(cookies)} cookies from file")

    def _login(self, member_id: str, password: str):
        """
        Log in with member credentials via the standard Atypon flow:
        1. GET login page, extract CSRF token
        2. POST credentials to /action/doLogin
        3. Collect Set-Cookie headers
        """
        print(f"[Login] Logging in as member {member_id}...")
        login_page_url = self.BASE_URL + self.LOGIN_PAGE
        try:
            resp = self.session.get(login_page_url, timeout=20)
            if resp.status_code != 200:
                print(f"  [Warning] Login page returned HTTP {resp.status_code}")
        except Exception as e:
            print(f"  [Error] Cannot reach login page: {e}")
            return

        csrf_token = ""
        for pattern in [
            r'name="csrf"[^>]*value="([^"]+)"',
            r'name="csrfToken"[^>]*value="([^"]+)"',
            r'"csrf_token":"([^"]+)"',
            r'name="_csrf"[^>]*value="([^"]+)"',
        ]:
            m = re.search(pattern, resp.text)
            if m:
                csrf_token = m.group(1)
                break

        session_cookies = {}
        for k, v in resp.headers.items():
            if k.lower() == "set-cookie":
                part = v.split(";")[0]
                if "=" in part:
                    name, val = part.split("=", 1)
                    session_cookies[name.strip()] = val.strip()

        post_url = self.BASE_URL + self.LOGIN_POST
        payload = {
            "login":       member_id,
            "password":    password,
            "redirectUri": "/",
            "action":      "login",
        }
        if csrf_token:
            payload["csrf"] = csrf_token
            payload["csrfToken"] = csrf_token

        post_headers = {
            "Content-Type": "application/x-www-form-urlencoded",
            "Referer":      login_page_url,
            "Origin":       self.BASE_URL,
            "Cookie":       "; ".join(f"{k}={v}" for k, v in session_cookies.items()),
        }

        try:
            resp2 = self.session.post(
                post_url, data=urlencode(payload), headers=post_headers,
                timeout=25, allow_redirects=True,
            )
        except Exception as e:
            print(f"  [Error] Login request failed: {e}")
            return

        all_cookies = dict(session_cookies)
        for h_name, h_val in resp2.headers.items():
            if h_name.lower() == "set-cookie":
                part = h_val.split(";")[0]
                if "=" in part:
                    name, val = part.split("=", 1)
                    all_cookies[name.strip()] = val.strip()

        auth_cookie_names = {"literatumJwt", "JSESSIONID", "literatumSession", "SESSION"}
        found_auth = any(k in all_cookies for k in auth_cookie_names)

        if not found_auth:
            if "logout" in resp2.text.lower() or member_id.lower() in resp2.text.lower():
                found_auth = True

        if found_auth or all_cookies:
            self._cookie_str = "; ".join(f"{k}={v}" for k, v in all_cookies.items())
            self.session.headers["Cookie"] = self._cookie_str
            if found_auth:
                print("  [✓] Login successful")
            else:
                print("  [?] Login status uncertain. Saved cookies and continuing.")
        else:
            print("  [✗] Login failed. Check member ID and password.")

    # ── Internal utilities ────────────────────────────────────────────────────

    def _delay(self):
        time.sleep(random.uniform(*self.delay_range))

    _NAV_HEADERS = {
        "Accept":                    "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language":           "en-US,en;q=0.9",
        "sec-fetch-dest":            "document",
        "sec-fetch-mode":            "navigate",
        "sec-fetch-site":            "same-origin",
        "sec-fetch-user":            "?1",
        "upgrade-insecure-requests": "1",
    }

    def _warmup(self):
        """Visit the homepage to establish a clean session and collect XSRF cookies."""
        try:
            resp = self.session.get(
                self.BASE_URL, timeout=20,
                headers={**self._NAV_HEADERS, "sec-fetch-site": "none"},
            )
            for k, v in resp.headers.items():
                if k.lower() == "set-cookie":
                    part = v.split(";")[0]
                    if "=" in part:
                        name, val = part.split("=", 1)
                        name, val = name.strip(), val.strip()
                        if name and val and name not in self._cookie_str:
                            self._cookie_str = (
                                self._cookie_str + f"; {name}={val}"
                                if self._cookie_str else f"{name}={val}"
                            )
            self.session.headers["Cookie"] = self._cookie_str
        except Exception:
            pass

    # ── CDP browser navigation (ultimate anti-bot: let real Chrome make requests) ──

    def _cdp_fetch_html(self, url: str, wait_seconds: float = 6.0,
                        debug_save: str = None) -> str:
        """
        Open url in debug Chrome via CDP, wait for JS rendering, and return full HTML.
        """
        try:
            import websocket as _ws
        except ImportError:
            print("  [CDP] websocket-client required: pip install websocket-client")
            return ""

        import urllib.request
        from urllib.parse import quote as _q

        try:
            req = urllib.request.Request(
                f"http://127.0.0.1:{self.CHROME_DBG_PORT}/json/new?{_q(url, safe=':/?&=%')}",
                method="PUT",
            )
            with urllib.request.urlopen(req, timeout=10) as r:
                tab = json.loads(r.read())
        except Exception as e:
            print(f"  [CDP] Cannot open new tab: {e}")
            return ""

        tab_id = tab.get("id", "")
        ws_url = tab.get("webSocketDebuggerUrl", "")

        def _close():
            try:
                urllib.request.urlopen(
                    f"http://127.0.0.1:{self.CHROME_DBG_PORT}/json/close/{tab_id}",
                    timeout=5,
                )
            except Exception:
                pass

        if not ws_url:
            _close()
            return ""

        html = ""
        try:
            ws  = _ws.create_connection(ws_url, timeout=20, suppress_origin=True)
            mid = [0]

            def send(method, params=None):
                mid[0] += 1
                ws.send(json.dumps({"id": mid[0], "method": method,
                                    "params": params or {}}))
                return mid[0]

            send("Page.enable")
            nav_id = send("Page.navigate", {"url": url})

            load_done  = False
            deadline   = time.time() + 30
            while not load_done and time.time() < deadline:
                ws.settimeout(max(0.3, deadline - time.time()))
                try:
                    msg = json.loads(ws.recv())
                except Exception:
                    continue
                ev = msg.get("method", "")
                if ev in ("Page.loadEventFired", "Page.frameStoppedLoading"):
                    load_done = True
                elif msg.get("id") == nav_id and "result" in msg:
                    load_done = True

            if not load_done:
                print("  [CDP] Page did not finish loading within 30s. Continuing anyway.")

            print(f"  [CDP] Page loaded. Waiting {wait_seconds:.0f}s for JS rendering...")
            time.sleep(wait_seconds)

            html_id = send("Runtime.evaluate", {
                "expression":    "document.documentElement.outerHTML",
                "returnByValue": True,
            })
            deadline2 = time.time() + 10
            while time.time() < deadline2:
                ws.settimeout(max(0.3, deadline2 - time.time()))
                try:
                    msg = json.loads(ws.recv())
                except Exception:
                    continue
                if msg.get("id") == html_id:
                    html = msg.get("result", {}).get("result", {}).get("value", "")
                    break

            ws.close()
        except Exception as e:
            print(f"  [CDP] Error: {e}")
        finally:
            _close()

        if debug_save:
            try:
                with open(debug_save, "w", encoding="utf-8") as f:
                    f.write(html or "(empty)")
                print(f"  [Debug] HTML saved → {debug_save}  ({len(html)} chars)")
            except Exception:
                pass

        return html

    def _get(self, url, referer=None, **kwargs):
        """
        GET request with retry, rate limiting, and browser navigation headers.
        Falls back to CDP (real Chrome) automatically on 403.
        """
        nav_headers = dict(self._NAV_HEADERS)
        nav_headers["Referer"] = referer or self.BASE_URL
        if "headers" in kwargs:
            nav_headers.update(kwargs.pop("headers"))

        for attempt in range(2):
            try:
                resp = self.session.get(
                    url, timeout=25, headers=nav_headers, **kwargs
                )
                if resp.status_code == 429:
                    wait = 30 * (attempt + 1)
                    print(f"  [Rate limit] Waiting {wait}s...")
                    time.sleep(wait)
                    continue
                if resp.status_code == 403:
                    return self._get_via_cdp(url)
                return resp
            except Exception as e:
                if attempt < 1:
                    time.sleep(3)
                else:
                    print(f"  [Network error] {e}")
                    return None
        return None

    def _get_via_cdp(self, url: str):
        """
        Fallback when curl_cffi is blocked (403): let debug Chrome make the request.
        Returns a mock response object with .status_code and .text.
        """
        if not self._is_chrome_debug_ready():
            print("  [CDP] Debug Chrome not running. Use --chrome-login to start it.")
            return None

        print("  [CDP] curl_cffi blocked — switching to Chrome browser (JS rendering)...")

        debug_path = "/tmp/informs_search_debug.html"
        html = self._cdp_fetch_html(url, wait_seconds=6.0, debug_save=debug_path)
        if not html:
            print(f"  [CDP] No content retrieved for: {url}")
            return None

        class _FakeResp:
            status_code = 200
            def __init__(self, text):
                self.text = text

        return _FakeResp(html)

    # ── HTML parsing ──────────────────────────────────────────────────────────

    def _parse_article_card(self, item) -> dict:
        """Parse a single article card from BeautifulSoup."""
        def text(tag):
            return tag.get_text(separator=" ", strip=True) if tag else ""

        title = ""
        for sel in [".hlFld-Title", "h5.card__title a", "h3.article-title a",
                    ".meta__title a", "h5 a", "h4 a", "h3 a"]:
            t = item.select_one(sel)
            if t:
                title = text(t)
                break

        doi = ""
        url = ""
        title_a = item.select_one(".hlFld-Title a, .meta__title a")
        if not title_a:
            title_a = item.select_one("a[href*='/doi/10.']")
        if title_a:
            href = title_a.get("href", "")
            m = re.search(r"/doi/(?:abs/|full/|epdf/|pdf/)?(10\.\d{4,}/[^\s?#\"]+)", href)
            if m:
                doi = m.group(1).rstrip("/")
                url = self.BASE_URL + href if not href.startswith("http") else href
        if not doi:
            for a in item.select("a[href*='/doi/']"):
                href = a.get("href", "")
                m = re.search(r"/doi/(?:abs/|full/|epdf/|pdf/)?(10\.\d{4,}/[^\s?#\"]+)", href)
                if m:
                    doi = m.group(1).rstrip("/")
                    url = self.BASE_URL + href if not href.startswith("http") else href
                    break

        authors = ""
        author_tags = item.select(".entryAuthor.hlFld-ContribAuthor")
        if author_tags:
            authors = "; ".join(text(a) for a in author_tags if text(a))
        if not authors:
            for sel in [".meta__authors .entryAuthor", ".meta__authors li",
                        ".loa-author-name", ".author-name"]:
                tags = item.select(sel)
                if tags:
                    authors = "; ".join(text(a) for a in tags if text(a))
                    break
        if not authors:
            t = item.select_one(".meta__authors")
            if t:
                authors = text(t)

        journal = ""
        t = item.select_one(".meta__serial")
        if t:
            journal = text(t)
        if not journal:
            for code, name in JOURNAL_NAMES.items():
                if doi and f"/{code}." in doi.lower():
                    journal = name
                    break

        year = volume = issue = date_str = ""
        t = item.select_one(".publicationYear")
        if t:
            year = text(t).strip("()")

        t = item.select_one(".meta__details, .rlist--inline.separator.toc-item__detail")
        if t:
            meta_text = text(t)
            if not year:
                m = re.search(r"\b(19|20)\d{2}\b", meta_text)
                if m:
                    year = m.group(0)
            m_vol = re.search(r"[Vv]ol(?:ume)?\.?\s*(\d+)", meta_text)
            if m_vol:
                volume = m_vol.group(1)
            m_iss = re.search(r"[Nn]o\.?\s*(\d+)|[Ii]ss(?:ue)?\.?\s*(\d+)", meta_text)
            if m_iss:
                issue = m_iss.group(1) or m_iss.group(2)

        if year:
            date_str = year

        abstract = ""
        t = item.select_one(".hlFld-Abstract, .card__abstract, .article-abstract")
        if t:
            abstract = text(t)

        pdf_url = ""
        pdf_a = item.select_one("a.pdfLink, a[href*='/doi/pdf/'], a[href*='pdfLink']")
        if pdf_a:
            href = pdf_a.get("href", "")
            pdf_url = self.BASE_URL + href if not href.startswith("http") else href
            if "download=true" not in pdf_url:
                pdf_url += ("&" if "?" in pdf_url else "?") + "download=true"
        elif doi:
            pdf_url = f"{self.BASE_URL}/doi/pdf/{doi}?download=true"

        return {
            "title":    title,
            "authors":  authors,
            "journal":  journal,
            "volume":   volume,
            "issue":    issue,
            "year":     year,
            "date":     date_str,
            "doi":      doi,
            "abstract": abstract,
            "url":      url,
            "pdf_url":  pdf_url,
        }

    def _parse_search_html(self, html: str) -> list:
        """Parse search result page HTML and return a list of article dicts."""
        if not HAS_BS4:
            print("[Error] beautifulsoup4 required: pip install beautifulsoup4 lxml")
            return []

        soup = BeautifulSoup(html, "lxml")
        articles = []

        SELECTORS = [
            "li.search__item",
            "li[class*='search__item']",
            "li.search-result-item",
            "li.issue-item",
            "div.issue-item",
            "li.card",
            "article.card",
            "li[class*='search-result']",
            "div[class*='search-result']",
            "ul.rlist li",
            "li[data-doi]",
        ]

        containers = []
        matched_sel = None
        for sel in SELECTORS:
            found = soup.select(sel)
            if found:
                containers = found
                matched_sel = sel
                break

        if not containers:
            containers = [
                tag for tag in soup.find_all(["li", "article", "div"])
                if (tag.find("a", href=re.compile(r"/doi/10\."))
                    and tag.find(["h3", "h4", "h5"]))
            ]
            if containers:
                matched_sel = "fallback(doi+heading)"

        if not containers:
            print("  [Debug] No article containers found. Page excerpt:")
            body_text = soup.get_text()[:300].replace("\n", " ").strip()
            print(f"    Text: {body_text}")
            doi_links = soup.find_all("a", href=re.compile(r"/doi/"))
            print(f"    /doi/ links found: {len(doi_links)}")
            classes_found = set()
            for tag in soup.find_all(["li", "article", "div"], limit=50):
                cls = " ".join(tag.get("class", []))
                if cls:
                    classes_found.add(cls[:60])
            if classes_found:
                print("    Block element class samples:")
                for c in list(classes_found)[:15]:
                    print(f"      {c}")
            print("  [Tip] Raw HTML saved to /tmp/informs_search_debug.html for inspection.")
        else:
            print(f"  [Parse] Matched selector: {matched_sel} — {len(containers)} containers")

        for item in containers:
            article = self._parse_article_card(item)
            if article["doi"] or article["title"]:
                articles.append(article)

        return articles

    def _parse_toc_html(self, html: str) -> list:
        """Parse a journal TOC page HTML."""
        if not HAS_BS4:
            print("[Error] beautifulsoup4 required: pip install beautifulsoup4 lxml")
            return []

        soup = BeautifulSoup(html, "lxml")

        containers = (
            soup.select("li.issue-item")
            or soup.select("div.issue-item")
            or soup.select("li.card")
            or soup.select("article.card")
        )

        if not containers:
            containers = [
                tag for tag in soup.find_all(["li", "article", "div"])
                if tag.find("a", href=re.compile(r"/doi/abs/|/doi/10\."))
            ]

        articles = []
        for item in containers:
            article = self._parse_article_card(item)
            if article["doi"] or article["title"]:
                articles.append(article)

        return articles

    def _get_total_results(self, html: str) -> int:
        """Extract total result count from a search page."""
        if not HAS_BS4:
            return 0
        soup = BeautifulSoup(html, "lxml")
        for sel in [".result-count", ".search-results-count",
                    "[class*='result-count']", "span[class*='found']"]:
            t = soup.select_one(sel)
            if t:
                m = re.search(r"[\d,]+", t.get_text())
                if m:
                    return int(m.group(0).replace(",", ""))
        m = re.search(r"([\d,]+)\s+results?", html, re.I)
        if m:
            return int(m.group(1).replace(",", ""))
        return 0

    # ── Search modes ──────────────────────────────────────────────────────────

    def search_by_keyword(self, query: str, count: int = 100,
                          sort_by: str = "relevance", date_range: str = None,
                          journal_code: str = None) -> list:
        """
        Search by keyword (supports AND/OR/NOT Boolean operators).

        Parameters
        ----------
        query        : search string, e.g. "machine learning" or "supply AND chain"
        count        : maximum papers to fetch
        sort_by      : "relevance" or "date"
        date_range   : year range, e.g. "2020-2024" (optional)
        journal_code : restrict to a journal, e.g. "mnsc" (optional)
        """
        print(f"\n[Keyword Search]  Query: {query}  Max: {count}")
        if journal_code:
            print(f"  Restricted to journal: {JOURNAL_NAMES.get(journal_code, journal_code)}")

        self._warmup()

        results = []
        start_page = 0
        total_known = None

        while len(results) < count:
            params = {
                "AllField":  query,
                "startPage": start_page,
                "pageSize":  PAGE_SIZE,
            }
            if sort_by == "date":
                params["sortBy"] = "Earliest_First"
            if date_range:
                parts = date_range.split("-")
                if len(parts) == 2:
                    params["startYear"] = parts[0]
                    params["endYear"]   = parts[1]
                elif len(parts) == 1:
                    params["startYear"] = parts[0]
                    params["endYear"]   = parts[0]
            if journal_code:
                params["SeriesKey"] = journal_code

            url = self.BASE_URL + self.SEARCH_URL + "?" + urlencode(params)
            resp = self._get(url)
            if not resp or resp.status_code != 200:
                print("  [Stop] Could not retrieve search results")
                break

            page_results = self._parse_search_html(resp.text)

            if total_known is None:
                total_known = self._get_total_results(resp.text)
                actual_max = min(count, total_known) if total_known > 0 else count
                if total_known > 0:
                    print(f"  Found {total_known} results. Planning to fetch {actual_max}.")
                else:
                    print(f"  Page returned {len(page_results)} results (total unknown).")

            if not page_results:
                print("  No more results.")
                break

            for art in page_results:
                if len(results) >= count:
                    break
                results.append(art)
                idx = len(results)
                actual_max = min(count, total_known or count)
                title_preview = (art["title"] or "(no title)")[:60]
                print(f"  [{idx}/{actual_max}] {title_preview}")

            start_page += 1
            if total_known and (start_page * PAGE_SIZE) >= total_known:
                break
            if len(results) >= count:
                break

            self._delay()

        return results

    def browse_journal(self, journal_code: str, count: int = 100) -> list:
        """
        Browse all articles in a journal (most recent first).
        Fetches the issue list from /loi/{journal}, then iterates through TOCs.
        """
        journal_name = JOURNAL_NAMES.get(journal_code, journal_code)
        print(f"\n[Journal Browse]  {journal_name} ({journal_code})  Max: {count}")

        self._warmup()
        loi_url = self.BASE_URL + f"/loi/{journal_code}"
        resp = self._get(loi_url)
        if not resp or resp.status_code != 200:
            print(f"  [Error] Cannot access journal page: {loi_url}")
            return []

        if not HAS_BS4:
            print("[Error] beautifulsoup4 required")
            return []

        soup = BeautifulSoup(resp.text, "lxml")
        toc_links = []
        for a in soup.select("a[href*='/toc/']"):
            href = a.get("href", "")
            m = re.match(r"/toc/(\w+)/(\d+)/(\d+)", href)
            if m and m.group(1) == journal_code:
                toc_links.append((int(m.group(2)), int(m.group(3)), href))

        if not toc_links:
            for a in soup.find_all("a", href=re.compile(r"/toc/")):
                href = a.get("href", "")
                m = re.match(r"/toc/\w+/(\d+)/(\d+)", href)
                if m:
                    toc_links.append((int(m.group(1)), int(m.group(2)), href))

        toc_links.sort(key=lambda x: (x[0], x[1]), reverse=True)
        print(f"  Found {len(toc_links)} issues. Fetching from most recent...")

        results = []
        for vol, iss, href in toc_links:
            if len(results) >= count:
                break
            toc_url = self.BASE_URL + href if not href.startswith("http") else href
            print(f"  → Vol. {vol} No. {iss}: {toc_url}")
            toc_resp = self._get(toc_url)
            if not toc_resp or toc_resp.status_code != 200:
                continue

            page_arts = self._parse_toc_html(toc_resp.text)
            for art in page_arts:
                if len(results) >= count:
                    break
                if not art.get("journal"):
                    art["journal"] = journal_name
                if not art.get("volume"):
                    art["volume"] = str(vol)
                if not art.get("issue"):
                    art["issue"] = str(iss)
                results.append(art)
                print(f"    [{len(results)}/{count}] {(art['title'] or '(no title)')[:60]}")

            self._delay()

        return results

    def browse_toc(self, journal_code: str, volume: int, issue: int) -> list:
        """
        Fetch the full table of contents for a specific volume/issue.

        Parameters
        ----------
        journal_code : e.g. "mnsc"
        volume       : e.g. 71
        issue        : e.g. 3
        """
        journal_name = JOURNAL_NAMES.get(journal_code, journal_code)
        print(f"\n[TOC Browse]  {journal_name} Vol.{volume} No.{issue}")

        self._warmup()
        toc_url = self.BASE_URL + self.TOC_URL.format(
            journal=journal_code, vol=volume, issue=issue
        )
        resp = self._get(toc_url)
        if not resp or resp.status_code != 200:
            print(f"  [Error] Cannot access: {toc_url}")
            return []

        articles = self._parse_toc_html(resp.text)
        for art in articles:
            if not art.get("journal"):
                art["journal"] = journal_name
            if not art.get("volume"):
                art["volume"] = str(volume)
            if not art.get("issue"):
                art["issue"] = str(issue)

        print(f"  {len(articles)} papers in this issue")
        for i, art in enumerate(articles, 1):
            print(f"  [{i}] {(art['title'] or '(no title)')[:65]}")

        return articles

    def search_advanced(
        self,
        query: str = None,
        journal_code: str = None,
        author: str = None,
        date_range: str = None,
        count: int = 100,
        sort_by: str = "relevance",
    ) -> list:
        """Advanced search: combine keyword, journal, author, and date range."""
        print("\n[Advanced Search]")
        if query:
            print(f"  Keyword:    {query}")
        if journal_code:
            print(f"  Journal:    {JOURNAL_NAMES.get(journal_code, journal_code)}")
        if author:
            print(f"  Author:     {author}")
        if date_range:
            print(f"  Date range: {date_range}")

        results = []
        start_page = 0
        total_known = None

        while len(results) < count:
            params = {"startPage": start_page, "pageSize": PAGE_SIZE}
            if query:
                params["AllField"] = query
            if author:
                params["ContribAuthor"] = author
            if journal_code:
                params["SeriesKey"] = journal_code
            if sort_by == "date":
                params["sortBy"] = "Earliest_First"
            if date_range:
                parts = date_range.split("-")
                if len(parts) == 2:
                    params["startYear"] = parts[0]
                    params["endYear"]   = parts[1]
                elif len(parts) == 1:
                    params["startYear"] = parts[0]
                    params["endYear"]   = parts[0]

            url = self.BASE_URL + self.SEARCH_URL + "?" + urlencode(params)
            resp = self._get(url)
            if not resp or resp.status_code != 200:
                break

            page_results = self._parse_search_html(resp.text)
            if total_known is None:
                total_known = self._get_total_results(resp.text)
                actual_max = min(count, total_known) if total_known > 0 else count
                if total_known > 0:
                    print(f"  Found {total_known} results. Planning to fetch {actual_max}.")

            if not page_results:
                break

            for art in page_results:
                if len(results) >= count:
                    break
                results.append(art)
                idx = len(results)
                actual_max = min(count, total_known or count)
                print(f"  [{idx}/{actual_max}] {(art['title'] or '(no title)')[:60]}")

            start_page += 1
            if len(results) >= count:
                break

            self._delay()

        return results

    # ── Save results ──────────────────────────────────────────────────────────

    FIELDS = [
        "title", "authors", "journal", "volume", "issue",
        "year", "date", "doi", "abstract", "url", "pdf_url",
    ]

    def save_to_csv(self, results: list, filename: str, output_dir: str) -> str:
        os.makedirs(output_dir, exist_ok=True)
        path = os.path.join(output_dir, filename)
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=self.FIELDS, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(results)
        print(f"\n[CSV] Saved → {path}  ({len(results)} papers)")
        return path

    def save_to_json(self, results: list, filename: str, output_dir: str) -> str:
        os.makedirs(output_dir, exist_ok=True)
        path = os.path.join(output_dir, filename)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        print(f"[JSON] Saved → {path}  ({len(results)} papers)")
        return path

    def save_to_xlsx(self, results: list, filename: str, output_dir: str) -> str:
        if not HAS_OPENPYXL:
            print("[Warning] openpyxl not installed. Falling back to CSV.")
            return self.save_to_csv(results, filename.replace(".xlsx", ".csv"), output_dir)
        os.makedirs(output_dir, exist_ok=True)
        path = os.path.join(output_dir, filename)
        wb = Workbook()
        ws = wb.active
        ws.title = "papers"
        ws.append(self.FIELDS)
        for item in results:
            ws.append([str(item.get(f, "") or "") for f in self.FIELDS])
        wb.save(path)
        print(f"[XLSX] Saved → {path}  ({len(results)} papers)")
        return path

    # ── PDF download ──────────────────────────────────────────────────────────

    @staticmethod
    def _make_pdf_filename(idx: int, article: dict) -> str:
        """Generate filename: {index}_{first_author_surname}_{year}_{title}.pdf"""
        authors = article.get("authors", "")
        first_author = (
            authors.split(";")[0].strip().split()[-1] if authors else "Unknown"
        )
        year  = article.get("year", "")
        title = re.sub(r'[\\/*?:"<>|]', "", article.get("title", ""))[:60].strip()
        return f"{idx:03d}_{first_author}_{year}_{title}.pdf"

    def download_pdfs(self, results: list, output_dir: str):
        """
        Batch PDF download.

        How it works
        ------------
        INFORMS PubsOnLine PDF endpoint:
          GET /doi/pdf/{doi}?download=true

        A valid session cookie is sufficient — no Playwright or DevTools required.
        The Atypon platform does not deploy Cloudflare JS challenges on PDF endpoints,
        making this much simpler than ScienceDirect.

        Prerequisite
        ------------
        Logged in via member credentials (--member + --password) or Chrome cookies
        (--browser-cookies / --chrome-login).
        """
        if not self._cookie_str:
            print("[Warning] No session cookie detected. PDF download may fail.")
            print("         Use --browser-cookies or --chrome-login.")

        pdf_dir = os.path.join(output_dir, "pdfs")
        os.makedirs(pdf_dir, exist_ok=True)

        total   = len(results)
        success = skip = fail = 0

        print(f"\n[PDF Download]  {total} papers → {pdf_dir}")

        dl_session = curl_requests.Session(impersonate="chrome124")
        dl_headers = {
            "User-Agent":      _UA,
            "Accept":          "application/pdf,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
            "Referer":         self.BASE_URL,
            "Cookie":          self._cookie_str,
        }

        for idx, article in enumerate(results, 1):
            doi        = article.get("doi", "")
            pdf_url    = article.get("pdf_url", "")
            title_short = (article.get("title", "") or "")[:50]

            if not doi and not pdf_url:
                print(f"  [{idx}/{total}] Skipped (no DOI): {title_short}")
                skip += 1
                continue

            if not pdf_url:
                pdf_url = f"{self.BASE_URL}/doi/pdf/{doi}?download=true"

            filename = self._make_pdf_filename(idx, article)
            filepath = os.path.join(pdf_dir, filename)

            if os.path.exists(filepath) and os.path.getsize(filepath) > 10_000:
                print(f"  [{idx}/{total}] Already exists, skipping: {filename}")
                skip += 1
                continue

            try:
                resp = dl_session.get(
                    pdf_url, headers=dl_headers, allow_redirects=True, timeout=60,
                )
                ct = resp.headers.get("content-type", "")

                if resp.status_code == 403:
                    print(f"  [{idx}/{total}] ✗ 403 Forbidden (session expired?): {title_short}")
                    fail += 1
                    if fail >= 3 and success == 0:
                        print("\n  [Stop] 3 consecutive 403 errors. Cookie may be invalid.")
                        print("  Please log in again to pubsonline.informs.org and retry.")
                        break
                    continue

                if resp.status_code == 401:
                    print(f"  [{idx}/{total}] ✗ 401 Unauthorized: {title_short}")
                    fail += 1
                    continue

                if resp.status_code != 200:
                    print(f"  [{idx}/{total}] ✗ HTTP {resp.status_code}: {title_short}")
                    fail += 1
                    continue

                content = resp.content
                if content[:4] == b"%PDF" or "pdf" in ct.lower():
                    with open(filepath, "wb") as f:
                        f.write(content)
                    size_kb = len(content) // 1024
                    print(f"  [{idx}/{total}] ✓ {filename}  ({size_kb} KB)")
                    success += 1
                elif b"login" in content[:2000].lower() or b"sign in" in content[:2000].lower():
                    print(f"  [{idx}/{total}] ✗ Redirected to login page (session expired): {title_short}")
                    fail += 1
                    if fail >= 3 and success == 0:
                        print("\n  [Stop] Session invalid. Please re-login and retry.")
                        break
                else:
                    print(f"  [{idx}/{total}] ✗ Non-PDF response ({ct[:40]}): {title_short}")
                    fail += 1

            except Exception as e:
                print(f"  [{idx}/{total}] ✗ Download error: {title_short}  ({e})")
                fail += 1

            if idx < total:
                time.sleep(random.uniform(*self.delay_range))

        print(f"\n[Done] Success: {success}  Skipped: {skip}  Failed: {fail}")


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def _build_parser():
    p = argparse.ArgumentParser(
        description="INFORMS PubsOnLine Paper Scraper",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Keyword search (read cookies from Chrome)
  python informs_scraper_en.py -m keyword -q "machine learning" -n 100 --browser-cookies

  # Keyword search + restrict journal + download PDFs
  python informs_scraper_en.py -m keyword -q "inventory" -j mnsc -n 50 --browser-cookies --download-pdf

  # Browse latest articles in a journal
  python informs_scraper_en.py -m journal -j opre -n 200 --browser-cookies

  # Specific volume/issue TOC
  python informs_scraper_en.py -m toc -j mnsc -v 71 -i 3 --browser-cookies

  # Member credentials
  python informs_scraper_en.py -m keyword -q "supply chain" -n 100 --member 123456 --password MyPwd

  # Advanced search: author + journal + year range
  python informs_scraper_en.py -m advanced -q "reinforcement learning" -j ijoc --author "Powell" --date 2020-2024 -n 50
        """,
    )

    p.add_argument(
        "-m", "--mode",
        choices=["keyword", "journal", "toc", "advanced"],
        default="keyword",
        help="Search mode (default: keyword)",
    )
    p.add_argument("-q", "--query",   help="Search keywords (keyword/advanced mode)")
    p.add_argument("-j", "--journal", help="Journal code, e.g. mnsc / opre / ijoc")
    p.add_argument("-v", "--volume",  type=int, help="Volume number (toc mode)")
    p.add_argument("-i", "--issue",   type=int, help="Issue number (toc mode)")
    p.add_argument("--author",        help="Author name (advanced mode)")
    p.add_argument("--date",          help="Year range, e.g. 2020-2024")
    p.add_argument(
        "-n", "--count", type=int, default=100,
        help="Max papers to fetch (default: 100)",
    )
    p.add_argument(
        "--sort", choices=["relevance", "date"], default="relevance",
        help="Sort order (default: relevance)",
    )

    auth = p.add_mutually_exclusive_group()
    auth.add_argument(
        "--chrome-login", action="store_true",
        help="★ Pop up Chrome for manual login; auto-extracts session cookies (most reliable)",
    )
    auth.add_argument(
        "--browser-cookies", action="store_true",
        help="Read persistent cookies from local Chrome (use --chrome-login if 403 occurs)",
    )
    auth.add_argument(
        "--cookies-file", metavar="PATH",
        help="Load cookies from a JSON file (Cookie Editor export)",
    )
    p.add_argument("--member",   help="INFORMS member ID")
    p.add_argument("--password", help="Account password (used with --member)")

    _default_out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "informs_result")
    p.add_argument(
        "-o", "--output-dir",
        default=_default_out,
        help=f"Output directory (default: {_default_out})",
    )
    p.add_argument(
        "--format", choices=["csv", "json", "xlsx"], default="csv",
        help="Metadata output format (default: csv)",
    )
    p.add_argument(
        "--download-pdf", action="store_true",
        help="Download PDFs after scraping",
    )
    p.add_argument(
        "--delay", type=float, nargs=2, metavar=("MIN", "MAX"), default=[2.0, 5.0],
        help="Request delay in seconds, default: 2 5",
    )

    return p


def _interactive_mode():
    """Interactive wizard (launched when no arguments are given)."""
    print("=" * 60)
    print("  INFORMS PubsOnLine Paper Scraper v1.0")
    print("=" * 60)
    print()
    print("Search mode:")
    print("  1. Keyword search")
    print("  2. Browse by journal")
    print("  3. Specific volume/issue TOC")
    print("  4. Advanced search (combine criteria)")
    mode_map = {"1": "keyword", "2": "journal", "3": "toc", "4": "advanced"}
    mode_choice = input("\nChoose (1-4, default 1): ").strip() or "1"
    mode = mode_map.get(mode_choice, "keyword")

    query = journal = volume = issue = author = date_range = None

    if mode in ("keyword", "advanced"):
        query = input("Search keywords (supports AND/OR/NOT): ").strip()

    if mode == "journal":
        print(f"\nJournal codes: {', '.join(JOURNAL_NAMES.keys())}")
        journal = input("Journal code: ").strip().lower()

    if mode == "toc":
        print(f"\nJournal codes: {', '.join(JOURNAL_NAMES.keys())}")
        journal = input("Journal code: ").strip().lower()
        volume  = int(input("Volume number (e.g. 71): ").strip())
        issue   = int(input("Issue number (e.g. 3): ").strip())

    if mode == "advanced":
        j = input("Restrict to journal code (optional, press Enter to skip): ").strip().lower()
        journal = j or None
        a = input("Restrict to author (optional, press Enter to skip): ").strip()
        author = a or None
        d = input("Year range (e.g. 2020-2024, optional): ").strip()
        date_range = d or None

    count = int(input("\nMax papers to fetch (default 100): ").strip() or "100")

    print("\nAuthentication:")
    print("  1. Pop up Chrome for manual login (★ most reliable, full session cookies)")
    print("  2. Read cookies from local Chrome (must be logged in already)")
    print("  3. Enter member ID + password (auto-login)")
    print("  4. None (metadata only, no PDF download)")
    auth_choice = input("Choose (1-4, default 1): ").strip() or "1"

    use_chrome_login = use_browser_cookies = False
    member_id = password = None

    if auth_choice == "1":
        use_chrome_login = True
    elif auth_choice == "2":
        use_browser_cookies = True
    elif auth_choice == "3":
        member_id = input("Member ID: ").strip()
        password  = input("Password: ").strip()

    dl_pdf = input("\nDownload PDFs? (y/N): ").strip().lower() == "y"
    _default_out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "informs_result")
    output_dir = input(f"Output directory (default {_default_out}): ").strip() or _default_out
    fmt = input("Output format (csv/json/xlsx, default csv): ").strip().lower() or "csv"

    return {
        "mode":               mode,
        "query":              query,
        "journal":            journal,
        "volume":             volume,
        "issue":              issue,
        "author":             author,
        "date":               date_range,
        "count":              count,
        "sort":               "relevance",
        "chrome_login":       use_chrome_login,
        "browser_cookies":    use_browser_cookies,
        "member":             member_id,
        "password":           password,
        "output_dir":         output_dir,
        "format":             fmt,
        "download_pdf":       dl_pdf,
        "delay":              [2.0, 5.0],
        "cookies_file":       None,
    }


def main():
    if len(sys.argv) == 1:
        cfg = _interactive_mode()
    else:
        parser = _build_parser()
        args   = parser.parse_args()
        cfg = vars(args)

    scraper = INFORMSScraper(
        member_id          = cfg.get("member"),
        password           = cfg.get("password"),
        cookies_file       = cfg.get("cookies_file"),
        use_browser_cookies= cfg.get("browser_cookies", False),
        use_chrome_login   = cfg.get("chrome_login", False),
        delay_range        = tuple(cfg.get("delay", [2.0, 5.0])),
    )

    mode    = cfg.get("mode", "keyword")
    results = []

    if mode == "keyword":
        if not cfg.get("query"):
            print("[Error] keyword mode requires -q / --query")
            sys.exit(1)
        results = scraper.search_by_keyword(
            query       = cfg["query"],
            count       = cfg.get("count", 100),
            sort_by     = cfg.get("sort", "relevance"),
            date_range  = cfg.get("date"),
            journal_code= cfg.get("journal"),
        )

    elif mode == "journal":
        if not cfg.get("journal"):
            print("[Error] journal mode requires -j / --journal")
            sys.exit(1)
        results = scraper.browse_journal(
            journal_code= cfg["journal"],
            count       = cfg.get("count", 100),
        )

    elif mode == "toc":
        if not all([cfg.get("journal"), cfg.get("volume"), cfg.get("issue")]):
            print("[Error] toc mode requires -j / --journal, -v / --volume, -i / --issue")
            sys.exit(1)
        results = scraper.browse_toc(
            journal_code= cfg["journal"],
            volume      = cfg["volume"],
            issue       = cfg["issue"],
        )

    elif mode == "advanced":
        results = scraper.search_advanced(
            query       = cfg.get("query"),
            journal_code= cfg.get("journal"),
            author      = cfg.get("author"),
            date_range  = cfg.get("date"),
            count       = cfg.get("count", 100),
            sort_by     = cfg.get("sort", "relevance"),
        )

    if not results:
        print("\nNo papers found. Exiting.")
        return

    output_dir = cfg.get("output_dir",
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "informs_result"))
    timestamp  = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name  = f"informs_{mode}_{timestamp}"
    fmt        = cfg.get("format", "csv")

    if fmt == "json":
        scraper.save_to_json(results, base_name + ".json", output_dir)
    elif fmt == "xlsx":
        scraper.save_to_xlsx(results, base_name + ".xlsx", output_dir)
    else:
        scraper.save_to_csv(results, base_name + ".csv", output_dir)

    if cfg.get("download_pdf"):
        scraper.download_pdfs(results, output_dir)

    print("\n[All done]")


if __name__ == "__main__":
    main()
