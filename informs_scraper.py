#!/usr/bin/env python
"""
INFORMS PubsOnLine 论文抓取工具 v1.1
=====================================
基于 Atypon/Literatum 平台，使用 curl_cffi 模拟 Chrome TLS 指纹。
支持关键词搜索、按期刊浏览、卷期目录抓取，结果保存为 CSV/JSON/XLSX，
支持 PDF 批量下载（会员账号直连，无需 Playwright）。

新增：--chrome-login 模式
  自动弹出真实 Chrome 窗口 → 你手动登录 → 脚本自动提取 session cookie → 开始抓取。
  无需手动导出/复制 cookie，最可靠。

INFORMS 期刊代码
----------------
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

快速上手
--------
  # 先用浏览器登录再用 browser-cookies（推荐）
  python informs_scraper.py -m keyword -q "machine learning" -n 100 --browser-cookies

  # 或直接提供会员账号
  python informs_scraper.py -m keyword -q "supply chain" -n 50 --member YOUR_ID --password YOUR_PWD

  # 按期刊浏览最新文章
  python informs_scraper.py -m journal -j mnsc -n 200 --browser-cookies

  # 浏览指定卷期
  python informs_scraper.py -m toc -j mnsc -v 71 -i 3 --browser-cookies

  # 搜索 + 下载 PDF
  python informs_scraper.py -m keyword -q "inventory" -n 30 --browser-cookies --download-pdf

  # ★ 弹出 Chrome 窗口手动登录（最推荐，cookie 最完整）
  python informs_scraper.py -m keyword -q "machine learning" -n 50 --chrome-login --download-pdf

依赖安装
--------
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
# 期刊代码 → 全名映射（用于展示）
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

# 每页结果数（Atypon 默认 20，可设 50）
PAGE_SIZE = 20

# Chrome UA
_UA = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)


# ─────────────────────────────────────────────────────────────────────────────
# 主爬虫类
# ─────────────────────────────────────────────────────────────────────────────

class INFORMSScraper:
    BASE_URL = "https://pubsonline.informs.org"

    # Atypon 标准端点
    LOGIN_PAGE  = "/literatumuserslogin"          # GET → CSRF + 表单
    LOGIN_POST  = "/action/doLogin"               # POST → 登录
    SEARCH_URL  = "/action/doSearch"              # GET → 搜索结果（HTML）
    TOC_URL     = "/toc/{journal}/{vol}/{issue}"  # GET → 卷期目录（HTML）
    JOURNAL_URL = "/loi/{journal}"                # GET → 期刊文章列表（HTML）

    # Chrome 调试配置
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
            "Accept-Language": "zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7",
        })
        self.delay_range = delay_range
        self._cookie_str = ""   # 拼好的 Cookie 请求头

        if use_chrome_login:
            self._chrome_login_flow()
        elif use_browser_cookies:
            self._load_browser_cookies()
        elif cookies_file:
            self._load_cookies_file(cookies_file)
        elif member_id and password:
            self._login(member_id, password)

    # ── Chrome 调试模式 ───────────────────────────────────────────────────────

    def _is_chrome_debug_ready(self) -> bool:
        """检查 Chrome 远程调试端口是否就绪。"""
        import urllib.request
        try:
            urllib.request.urlopen(
                f"http://127.0.0.1:{self.CHROME_DBG_PORT}/json/version", timeout=2)
            return True
        except Exception:
            return False

    def _launch_chrome_with_debug(self):
        """
        以调试模式启动 Chrome，复用默认 Profile（带已保存的密码/书签/历史）。
        返回 Popen 对象，失败返回 None。
        """
        import subprocess
        import shutil

        default_profile = os.path.expanduser(
            "~/Library/Application Support/Google/Chrome/Default"
        )
        tmp_default = os.path.join(self.CHROME_DBG_PROFILE, "Default")
        os.makedirs(tmp_default, exist_ok=True)

        # 复制关键文件：Cookie、Preferences、登录状态
        for fname in ("Cookies", "Cookies-journal", "Preferences",
                      "Login Data", "Web Data"):
            src = os.path.join(default_profile, fname)
            dst = os.path.join(tmp_default, fname)
            if os.path.exists(src) and not os.path.exists(dst):
                try:
                    shutil.copy2(src, dst)
                except Exception:
                    pass

        # 清除残留锁文件
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
            print(f"  Chrome 已启动 (PID {proc.pid})，等待调试端口就绪...")
        except FileNotFoundError:
            print(f"  [错误] 找不到 Chrome：{self.CHROME_BIN}")
            return None

        for i in range(40):
            time.sleep(1)
            if self._is_chrome_debug_ready():
                print(f"  调试端口就绪 ({i + 1}s) ✓")
                return proc
            if (i + 1) % 5 == 0:
                print(f"  等待 Chrome 启动... ({i + 1}s)")

        print("  [警告] 40s 内未就绪，查看日志：cat /tmp/chrome_informs.log")
        return None

    def _extract_cookies_via_cdp(self) -> str:
        """
        通过 CDP websocket 从运行中的 Chrome 提取 pubsonline.informs.org 的所有 cookie。
        返回拼好的 Cookie 请求头字符串。无需 Playwright。
        """
        import urllib.request
        try:
            with urllib.request.urlopen(
                f"http://127.0.0.1:{self.CHROME_DBG_PORT}/json/list", timeout=5
            ) as r:
                pages = json.loads(r.read())
        except Exception as e:
            print(f"  [错误] 无法获取 CDP 标签列表：{e}")
            return ""

        # 优先选 pubsonline 标签，其次选第一个有 ws url 的标签
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
            print("  [错误] 没有可用的 CDP 标签页")
            return ""

        try:
            import websocket as _ws
        except ImportError:
            print("  [提示] websocket-client 未安装，改用 browser_cookie3 读取")
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
                print(f"  已通过 CDP 提取 {len(cookies)} 个 cookie ✓")
                return cookie_str
            else:
                print("  [警告] CDP 返回 0 个 cookie（可能尚未登录）")
                return ""
        except Exception as e:
            print(f"  [警告] CDP cookie 提取失败：{e}")
            return ""

    def _open_informs_tab(self):
        """在调试 Chrome 中打开 INFORMS 登录页。"""
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
            print(f"  [警告] 无法打开新标签：{e}")

    def _chrome_login_flow(self):
        """
        完整的 Chrome 登录流程：
        1. 启动调试 Chrome（复用你的默认 Profile）
        2. 打开 INFORMS 登录页
        3. 等你手动登录，按 Enter 确认
        4. 通过 CDP 提取 session cookie
        5. 设置到 session headers
        """
        print("\n[Chrome 登录]  弹出 Chrome 窗口，请在其中登录 INFORMS...")

        if not self._is_chrome_debug_ready():
            proc = self._launch_chrome_with_debug()
            if not self._is_chrome_debug_ready():
                print("  [错误] Chrome 启动失败，查看：cat /tmp/chrome_informs.log")
                return
        else:
            print("  检测到已运行的调试 Chrome ✓")

        # 打开登录页
        self._open_informs_tab()
        time.sleep(1.5)

        print()
        print("=" * 60)
        print("  Chrome 窗口已打开 pubsonline.informs.org 登录页。")
        print()
        print("  请：")
        print("  1. 在 Chrome 中用会员号 + 密码登录")
        print("  2. 确认页面显示你的名字或进入期刊主页")
        print("  3. 回到这里按 Enter 继续")
        print("=" * 60)

        try:
            input("\n  >>> 登录完成后按 Enter：")
        except EOFError:
            print("  [非交互模式] 等待 30s 后自动继续...")
            time.sleep(30)

        # 提取 cookie
        cookie_str = self._extract_cookies_via_cdp()
        if cookie_str:
            self._cookie_str = cookie_str
            self.session.headers["Cookie"] = self._cookie_str
            print("  [✓] Cookie 已设置，暖身中...")
            self._warmup()   # 获取 XSRF 等服务器 cookie
            print("  [✓] 就绪，开始抓取")
        else:
            print("  [!] CDP 提取失败，尝试从磁盘读取 Chrome cookie...")
            self._load_browser_cookies()
            if self._cookie_str:
                self._warmup()

    # ── 认证 ─────────────────────────────────────────────────────────────────

    def _load_browser_cookies(self, silent=False) -> str:
        """
        从本机 Chrome 读取 pubsonline.informs.org 的持久化 Cookie。
        注意：browser_cookie3 只能读磁盘 cookie，不含 session cookie，
        因此效果不如 CDP 提取（--chrome-login）。
        """
        if not HAS_BROWSER_COOKIE3:
            if not silent:
                print("[错误] 未安装 browser-cookie3：pip install browser-cookie3")
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
                    print(f"[信息] 已从 Chrome 读取 {len(cookies)} 个 cookie")
                    print("  提示：若遇 403，改用 --chrome-login（提取完整 session cookie）")
                return self._cookie_str
            else:
                if not silent:
                    print("[警告] Chrome 中未找到 informs.org 的 cookie")
                    print("  请先在 Chrome 登录 pubsonline.informs.org，或改用 --chrome-login")
                return ""
        except Exception as e:
            if not silent:
                print(f"[警告] 读取 Chrome cookie 失败：{e}")
                print("       macOS 可能弹出钥匙串权限，请点「允许」")
            return ""

    def _load_cookies_file(self, path):
        """从 JSON 文件（Netscape 或字典格式）加载 Cookie。"""
        if not os.path.exists(path):
            print(f"[警告] 找不到 cookie 文件：{path}")
            return
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            cookies = {c.get("name", ""): c.get("value", "") for c in data if c.get("name")}
        elif isinstance(data, dict):
            cookies = data
        else:
            print("[警告] cookie 文件格式不支持")
            return
        self._cookie_str = "; ".join(f"{k}={v}" for k, v in cookies.items() if k and v)
        self.session.headers["Cookie"] = self._cookie_str
        print(f"[信息] 已从文件加载 {len(cookies)} 个 cookie")

    def _login(self, member_id: str, password: str):
        """
        用会员账号密码登录 INFORMS PubsOnLine（Atypon 标准流程）。
        步骤：
          1. GET 登录页，提取 CSRF token
          2. POST 凭证到 /action/doLogin
          3. 收集 Set-Cookie 存入 session
        """
        print(f"[登录] 使用会员号 {member_id} 登录...")
        login_page_url = self.BASE_URL + self.LOGIN_PAGE
        try:
            resp = self.session.get(login_page_url, timeout=20)
            if resp.status_code != 200:
                print(f"  [警告] 登录页面返回 HTTP {resp.status_code}")
        except Exception as e:
            print(f"  [错误] 无法访问登录页：{e}")
            return

        # 提取 CSRF token（Atypon 常见字段名）
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

        # 收集登录页下发的 cookie
        session_cookies = {}
        for k, v in resp.headers.items():
            if k.lower() == "set-cookie":
                part = v.split(";")[0]
                if "=" in part:
                    name, val = part.split("=", 1)
                    session_cookies[name.strip()] = val.strip()

        # POST 登录
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
                post_url,
                data=urlencode(payload),
                headers=post_headers,
                timeout=25,
                allow_redirects=True,
            )
        except Exception as e:
            print(f"  [错误] 登录请求失败：{e}")
            return

        # 收集所有 Set-Cookie（包括重定向链）
        all_cookies = dict(session_cookies)
        for h_name, h_val in resp2.headers.items():
            if h_name.lower() == "set-cookie":
                part = h_val.split(";")[0]
                if "=" in part:
                    name, val = part.split("=", 1)
                    all_cookies[name.strip()] = val.strip()

        # 检查是否登录成功（Atypon 登录成功后会有 literatumJwt 或 JS_SESSION 等 cookie）
        auth_cookie_names = {"literatumJwt", "JSESSIONID", "literatumSession", "SESSION"}
        found_auth = any(k in all_cookies for k in auth_cookie_names)

        # 如果没有明确的 auth cookie，也检查响应里是否有用户名
        if not found_auth:
            if "logout" in resp2.text.lower() or member_id.lower() in resp2.text.lower():
                found_auth = True

        if found_auth or all_cookies:
            self._cookie_str = "; ".join(f"{k}={v}" for k, v in all_cookies.items())
            self.session.headers["Cookie"] = self._cookie_str
            if found_auth:
                print("  [✓] 登录成功")
            else:
                print("  [?] 登录状态不确定，已保存 cookie，继续尝试")
        else:
            print("  [✗] 登录失败，可能密码错误或账号格式有误")
            print("      提示：会员号格式如 123456，密码区分大小写")

    # ── 内部工具 ─────────────────────────────────────────────────────────────

    def _delay(self):
        time.sleep(random.uniform(*self.delay_range))

    # 浏览器正常导航时携带的头部——Atypon 会检查这些
    _NAV_HEADERS = {
        "Accept":                  "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language":         "zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7",
        "sec-fetch-dest":          "document",
        "sec-fetch-mode":          "navigate",
        "sec-fetch-site":          "same-origin",
        "sec-fetch-user":          "?1",
        "upgrade-insecure-requests": "1",
    }

    def _warmup(self):
        """
        访问主页让服务器建立干净 session，收集服务器下发的 XSRF cookie。
        """
        try:
            resp = self.session.get(
                self.BASE_URL,
                timeout=20,
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

    # ── CDP 浏览器内导航（终极反反爬：让真实 Chrome 做请求）────────────────────

    def _cdp_fetch_html(self, url: str, wait_seconds: float = 6.0,
                        debug_save: str = None) -> str:
        """
        通过 CDP 在调试 Chrome 里打开 url，等待 JS 渲染后返回完整 HTML。

        策略：导航 → 等 loadEventFired → 再等 wait_seconds（JS 渲染）→ 取 HTML。
        全程只维护一个消息循环，无嵌套轮询，逻辑简单可靠。
        """
        try:
            import websocket as _ws
        except ImportError:
            print("  [CDP] 需要 websocket-client：pip install websocket-client")
            return ""

        import urllib.request
        from urllib.parse import quote as _q

        # ── 打开新标签 ──
        try:
            req = urllib.request.Request(
                f"http://127.0.0.1:{self.CHROME_DBG_PORT}/json/new?{_q(url, safe=':/?&=%')}",
                method="PUT",
            )
            with urllib.request.urlopen(req, timeout=10) as r:
                tab = json.loads(r.read())
        except Exception as e:
            print(f"  [CDP] 无法打开新标签：{e}")
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

            # ── 阶段 1：等待导航完成（Page.loadEventFired 或 frameStoppedLoading）──
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
                # 备用：navigate 的 result 回来也算
                elif msg.get("id") == nav_id and "result" in msg:
                    load_done = True

            if not load_done:
                print("  [CDP] 页面未在 30s 内加载完成，强制继续")

            # ── 阶段 2：等 JS 渲染 ──
            print(f"  [CDP] 页面已加载，等待 JS 渲染 {wait_seconds:.0f}s...")
            time.sleep(wait_seconds)

            # ── 阶段 3：取 HTML ──
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
            print(f"  [CDP] 错误：{e}")
        finally:
            _close()

        if debug_save:
            try:
                with open(debug_save, "w", encoding="utf-8") as f:
                    f.write(html or "(empty)")
                print(f"  [调试] HTML 已保存 → {debug_save}  ({len(html)} chars)")
            except Exception:
                pass

        return html

    def _get(self, url, referer=None, **kwargs):
        """
        带重试、限速、浏览器导航头的 GET 请求。
        若 curl_cffi 被 403 拦截，自动回退到 CDP（真实 Chrome）发请求。
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
                    print(f"  [限速] 等待 {wait}s...")
                    time.sleep(wait)
                    continue
                if resp.status_code == 403:
                    # curl_cffi 被拦 → 用 CDP 浏览器重试
                    return self._get_via_cdp(url)
                return resp
            except Exception as e:
                if attempt < 1:
                    time.sleep(3)
                else:
                    print(f"  [网络错误] {e}")
                    return None
        return None

    def _get_via_cdp(self, url: str):
        """
        curl_cffi 被 403 时的备用方案：让调试 Chrome 发请求，返回一个模拟 Response。
        搜索结果页是 JS 渲染，传入 wait_selector 等待结果容器出现。
        同时保存原始 HTML 到临时文件方便排查选择器问题。
        """
        if not self._is_chrome_debug_ready():
            print("  [CDP] 调试 Chrome 未运行，请用 --chrome-login 启动。")
            return None

        print("  [CDP] curl_cffi 被拦，改用 Chrome 浏览器请求（等待 JS 渲染）...")

        debug_path = "/tmp/informs_search_debug.html"
        html = self._cdp_fetch_html(
            url,
            wait_seconds=6.0,
            debug_save=debug_path,
        )
        if not html:
            print(f"  [CDP] 未获取到内容：{url}")
            return None

        class _FakeResp:
            status_code = 200
            def __init__(self, text):
                self.text = text

        return _FakeResp(html)

    # ── HTML 解析 ─────────────────────────────────────────────────────────────

    def _parse_article_card(self, item) -> dict:
        """
        从 BeautifulSoup Tag 解析单篇文章元数据。
        优先使用 INFORMS PubsOnLine 的真实 class 名，备用通用选择器。
        """
        def text(tag):
            return tag.get_text(separator=" ", strip=True) if tag else ""

        # ── 标题 ──
        title = ""
        for sel in [".hlFld-Title", "h5.card__title a", "h3.article-title a",
                    ".meta__title a", "h5 a", "h4 a", "h3 a"]:
            t = item.select_one(sel)
            if t:
                title = text(t)
                break

        # ── DOI / URL ──
        doi = ""
        url = ""
        # 先从 .hlFld-Title 的 <a> 取
        title_a = item.select_one(".hlFld-Title a, .meta__title a")
        if not title_a:
            title_a = item.select_one("a[href*='/doi/10.']")
        if title_a:
            href = title_a.get("href", "")
            m = re.search(r"/doi/(?:abs/|full/|epdf/|pdf/)?(10\.\d{4,}/[^\s?#\"]+)", href)
            if m:
                doi = m.group(1).rstrip("/")
                url = self.BASE_URL + href if not href.startswith("http") else href
        # 备用：任意含 DOI 的链接
        if not doi:
            for a in item.select("a[href*='/doi/']"):
                href = a.get("href", "")
                m = re.search(r"/doi/(?:abs/|full/|epdf/|pdf/)?(10\.\d{4,}/[^\s?#\"]+)", href)
                if m:
                    doi = m.group(1).rstrip("/")
                    url = self.BASE_URL + href if not href.startswith("http") else href
                    break

        # ── 作者 ──
        authors = ""
        # INFORMS 真实 class
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

        # ── 期刊 ──
        journal = ""
        t = item.select_one(".meta__serial")
        if t:
            journal = text(t)
        if not journal:
            for code, name in JOURNAL_NAMES.items():
                if doi and f"/{code}." in doi.lower():
                    journal = name
                    break

        # ── 年份 / 卷 / 期 ──
        year = volume = issue = date_str = ""
        t = item.select_one(".publicationYear")
        if t:
            year = text(t).strip("()")

        # meta__details 含 "Vol. XX No. X"
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

        # ── 摘要 ──
        abstract = ""
        t = item.select_one(".hlFld-Abstract, .card__abstract, .article-abstract")
        if t:
            abstract = text(t)

        # ── PDF URL（优先用页面中的 .pdfLink，其次构造）──
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
        """解析搜索结果页 HTML，返回文章列表。"""
        if not HAS_BS4:
            print("[错误] 需要 beautifulsoup4：pip install beautifulsoup4 lxml")
            return []

        soup = BeautifulSoup(html, "lxml")
        articles = []

        # 按优先级尝试各种 Atypon 搜索结果容器选择器
        # INFORMS 真实 class（从页面 HTML 分析得到）排在最前
        SELECTORS = [
            "li.search__item",               # INFORMS 真实容器 (<li> not <div>)
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
            # 最终备用：所有含 /doi/ 链接且有标题的块级元素
            containers = [
                tag for tag in soup.find_all(["li", "article", "div"])
                if (tag.find("a", href=re.compile(r"/doi/10\."))
                    and tag.find(["h3", "h4", "h5"]))
            ]
            if containers:
                matched_sel = "fallback(doi+heading)"

        if not containers:
            # 打印调试信息帮助排查
            print("  [调试] 未找到文章容器，页面摘要：")
            body_text = soup.get_text()[:300].replace("\n", " ").strip()
            print(f"    文本片段: {body_text}")
            doi_links = soup.find_all("a", href=re.compile(r"/doi/"))
            print(f"    /doi/ 链接数: {len(doi_links)}")
            classes_found = set()
            for tag in soup.find_all(["li", "article", "div"], limit=50):
                cls = " ".join(tag.get("class", []))
                if cls:
                    classes_found.add(cls[:60])
            if classes_found:
                print(f"    前50个块级元素 class 样本:")
                for c in list(classes_found)[:15]:
                    print(f"      {c}")
            print("  [提示] 原始 HTML 已保存至 /tmp/informs_search_debug.html，可手动检查")
        else:
            print(f"  [解析] 匹配选择器: {matched_sel}，找到 {len(containers)} 个容器")

        for item in containers:
            article = self._parse_article_card(item)
            if article["doi"] or article["title"]:
                articles.append(article)

        return articles

    def _parse_toc_html(self, html: str) -> list:
        """解析期刊 TOC（目录）页 HTML。结构与搜索结果类似但略有不同。"""
        if not HAS_BS4:
            print("[错误] 需要 beautifulsoup4：pip install beautifulsoup4 lxml")
            return []

        soup = BeautifulSoup(html, "lxml")

        # TOC 页典型结构
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
        """从搜索结果页提取总结果数。"""
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
        # fallback：在页面文本中找 "X results"
        m = re.search(r"([\d,]+)\s+results?", html, re.I)
        if m:
            return int(m.group(1).replace(",", ""))
        return 0

    # ── 搜索模式 ─────────────────────────────────────────────────────────────

    def search_by_keyword(self, query: str, count: int = 100,
                          sort_by: str = "relevance", date_range: str = None,
                          journal_code: str = None) -> list:
        """
        按关键词搜索（支持 AND/OR/NOT 布尔运算符）。

        参数
        ----
        query       : 搜索词，如 "machine learning" 或 "supply AND chain"
        count       : 最多抓取篇数
        sort_by     : 排序，"relevance"（相关性）或 "date"（时间）
        date_range  : 年份范围，如 "2020-2024"（可选）
        journal_code: 限定期刊代码，如 "mnsc"（可选）
        """
        print(f"\n[关键词搜索]  关键词: {query}  最多: {count} 篇")
        if journal_code:
            print(f"  限定期刊: {JOURNAL_NAMES.get(journal_code, journal_code)}")

        self._warmup()   # 确保 XSRF/session cookie 最新

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
                # Atypon 格式：startYear=2020&endYear=2024
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
                print("  [停止] 无法获取搜索结果")
                break

            page_results = self._parse_search_html(resp.text)

            if total_known is None:
                total_known = self._get_total_results(resp.text)
                actual_max = min(count, total_known) if total_known > 0 else count
                if total_known > 0:
                    print(f"  共找到 {total_known} 篇，计划抓取 {actual_max} 篇")
                else:
                    print(f"  本页返回 {len(page_results)} 篇（总数未知）")

            if not page_results:
                print("  没有更多结果")
                break

            for art in page_results:
                if len(results) >= count:
                    break
                results.append(art)
                idx = len(results)
                actual_max = min(count, total_known or count)
                title_preview = (art["title"] or "（无标题）")[:60]
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
        浏览期刊所有文章（按时间倒序）。
        通过 /loi/{journal} 页面获取卷期列表，再逐期抓取 TOC。
        """
        journal_name = JOURNAL_NAMES.get(journal_code, journal_code)
        print(f"\n[期刊浏览]  期刊: {journal_name} ({journal_code})  最多: {count} 篇")

        self._warmup()
        loi_url = self.BASE_URL + f"/loi/{journal_code}"
        resp = self._get(loi_url)
        if not resp or resp.status_code != 200:
            print(f"  [错误] 无法访问期刊页面：{loi_url}")
            return []

        # 解析卷期链接
        if not HAS_BS4:
            print("[错误] 需要 beautifulsoup4")
            return []

        soup = BeautifulSoup(resp.text, "lxml")
        # Atypon loi 页面：每个 issue 有 <a href="/toc/{journal}/{vol}/{issue}">
        toc_links = []
        for a in soup.select("a[href*='/toc/']"):
            href = a.get("href", "")
            m = re.match(r"/toc/(\w+)/(\d+)/(\d+)", href)
            if m and m.group(1) == journal_code:
                toc_links.append((int(m.group(2)), int(m.group(3)), href))

        if not toc_links:
            # 备用：找所有 /toc/ href
            for a in soup.find_all("a", href=re.compile(r"/toc/")):
                href = a.get("href", "")
                m = re.match(r"/toc/\w+/(\d+)/(\d+)", href)
                if m:
                    toc_links.append((int(m.group(1)), int(m.group(2)), href))

        # 按 (vol, issue) 倒序
        toc_links.sort(key=lambda x: (x[0], x[1]), reverse=True)
        print(f"  发现 {len(toc_links)} 期，从最新开始抓取...")

        results = []
        for vol, iss, href in toc_links:
            if len(results) >= count:
                break
            toc_url = self.BASE_URL + href if not href.startswith("http") else href
            print(f"  → 卷 {vol} 期 {iss}：{toc_url}")
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
                print(f"    [{len(results)}/{count}] {(art['title'] or '无标题')[:60]}")

            self._delay()

        return results

    def browse_toc(self, journal_code: str, volume: int, issue: int) -> list:
        """
        抓取指定卷期的完整目录。

        参数
        ----
        journal_code : 期刊代码，如 "mnsc"
        volume       : 卷号，如 71
        issue        : 期号，如 3
        """
        journal_name = JOURNAL_NAMES.get(journal_code, journal_code)
        print(f"\n[目录浏览]  {journal_name} Vol.{volume} No.{issue}")

        self._warmup()
        toc_url = self.BASE_URL + self.TOC_URL.format(
            journal=journal_code, vol=volume, issue=issue
        )
        resp = self._get(toc_url)
        if not resp or resp.status_code != 200:
            print(f"  [错误] 无法访问：{toc_url}")
            return []

        articles = self._parse_toc_html(resp.text)
        for art in articles:
            if not art.get("journal"):
                art["journal"] = journal_name
            if not art.get("volume"):
                art["volume"] = str(volume)
            if not art.get("issue"):
                art["issue"] = str(issue)

        print(f"  本期共 {len(articles)} 篇")
        for i, art in enumerate(articles, 1):
            print(f"  [{i}] {(art['title'] or '无标题')[:65]}")

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
        """高级搜索：组合关键词、期刊、作者、时间范围。"""
        print("\n[高级搜索]")
        if query:
            print(f"  关键词:   {query}")
        if journal_code:
            print(f"  期刊:     {JOURNAL_NAMES.get(journal_code, journal_code)}")
        if author:
            print(f"  作者:     {author}")
        if date_range:
            print(f"  时间范围: {date_range}")

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
                    print(f"  共找到 {total_known} 篇，计划抓取 {actual_max} 篇")

            if not page_results:
                break

            for art in page_results:
                if len(results) >= count:
                    break
                results.append(art)
                idx = len(results)
                actual_max = min(count, total_known or count)
                print(f"  [{idx}/{actual_max}] {(art['title'] or '无标题')[:60]}")

            start_page += 1
            if len(results) >= count:
                break

            self._delay()

        return results

    # ── 保存结果 ─────────────────────────────────────────────────────────────

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
        print(f"\n[CSV] 已保存 → {path}  （共 {len(results)} 篇）")
        return path

    def save_to_json(self, results: list, filename: str, output_dir: str) -> str:
        os.makedirs(output_dir, exist_ok=True)
        path = os.path.join(output_dir, filename)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        print(f"[JSON] 已保存 → {path}  （共 {len(results)} 篇）")
        return path

    def save_to_xlsx(self, results: list, filename: str, output_dir: str) -> str:
        if not HAS_OPENPYXL:
            print("[警告] 未安装 openpyxl，回退保存为 CSV")
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
        print(f"[XLSX] 已保存 → {path}  （共 {len(results)} 篇）")
        return path

    # ── PDF 下载 ─────────────────────────────────────────────────────────────

    @staticmethod
    def _make_pdf_filename(idx: int, article: dict) -> str:
        """生成文件名：{序号}_{第一作者姓}_{年份}_{标题截断}.pdf"""
        authors = article.get("authors", "")
        first_author = (
            authors.split(";")[0].strip().split()[-1] if authors else "Unknown"
        )
        year  = article.get("year", "")
        title = re.sub(r'[\\/*?:"<>|]', "", article.get("title", ""))[:60].strip()
        return f"{idx:03d}_{first_author}_{year}_{title}.pdf"

    def download_pdfs(self, results: list, output_dir: str):
        """
        批量下载 PDF。

        原理
        ----
        INFORMS PubsOnLine 的 PDF 端点：
          GET /doi/pdf/{doi}?download=true

        有效的登录 session cookie 即可直接下载，无需 Playwright 或 DevTools。
        比 ScienceDirect 简单得多——Atypon 平台不在 PDF 端点部署 Cloudflare JS 验证。

        前提
        ----
        已通过会员账号登录（--member + --password）或从 Chrome 读取 cookie（--browser-cookies）。
        """
        if not self._cookie_str:
            print("[警告] 未检测到登录 cookie，PDF 下载可能失败（请用 --browser-cookies 或 --member）")

        pdf_dir = os.path.join(output_dir, "pdfs")
        os.makedirs(pdf_dir, exist_ok=True)

        total   = len(results)
        success = skip = fail = 0

        print(f"\n[PDF 下载]  共 {total} 篇，保存至 {pdf_dir}")

        dl_session = curl_requests.Session(impersonate="chrome124")
        dl_headers = {
            "User-Agent":      _UA,
            "Accept":          "application/pdf,*/*;q=0.8",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
            "Referer":         self.BASE_URL,
            "Cookie":          self._cookie_str,
        }

        for idx, article in enumerate(results, 1):
            doi        = article.get("doi", "")
            pdf_url    = article.get("pdf_url", "")
            title_short = (article.get("title", "") or "")[:50]

            if not doi and not pdf_url:
                print(f"  [{idx}/{total}] 跳过（无 DOI）: {title_short}")
                skip += 1
                continue

            if not pdf_url:
                pdf_url = f"{self.BASE_URL}/doi/pdf/{doi}?download=true"

            filename = self._make_pdf_filename(idx, article)
            filepath = os.path.join(pdf_dir, filename)

            if os.path.exists(filepath) and os.path.getsize(filepath) > 10_000:
                print(f"  [{idx}/{total}] 已存在，跳过: {filename}")
                skip += 1
                continue

            try:
                resp = dl_session.get(
                    pdf_url,
                    headers=dl_headers,
                    allow_redirects=True,
                    timeout=60,
                )
                ct = resp.headers.get("content-type", "")

                if resp.status_code == 403:
                    print(f"  [{idx}/{total}] ✗ 403 权限不足（登录已过期？）: {title_short}")
                    fail += 1
                    # 连续 403 停止
                    if fail >= 3 and success == 0:
                        print("\n  [停止] 连续 3 次 403，cookie 可能已失效。")
                        print("  请重新在 Chrome 登录 pubsonline.informs.org 后再运行。")
                        break
                    continue

                if resp.status_code == 401:
                    print(f"  [{idx}/{total}] ✗ 401 未授权: {title_short}")
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
                    print(f"  [{idx}/{total}] ✗ 被重定向到登录页（session 过期）: {title_short}")
                    fail += 1
                    if fail >= 3 and success == 0:
                        print("\n  [停止] Session 已失效，请重新登录后再运行。")
                        break
                else:
                    print(f"  [{idx}/{total}] ✗ 非 PDF 响应（{ct[:40]}）: {title_short}")
                    fail += 1

            except Exception as e:
                print(f"  [{idx}/{total}] ✗ 下载异常: {title_short}  ({e})")
                fail += 1

            if idx < total:
                time.sleep(random.uniform(*self.delay_range))

        print(f"\n[完成] 成功: {success}  跳过: {skip}  失败: {fail}")


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def _build_parser():
    p = argparse.ArgumentParser(
        description="INFORMS PubsOnLine 论文抓取工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 关键词搜索（从 Chrome 读取 cookie）
  python informs_scraper.py -m keyword -q "machine learning" -n 100 --browser-cookies

  # 关键词搜索 + 限定期刊 + 下载 PDF
  python informs_scraper.py -m keyword -q "inventory" -j mnsc -n 50 --browser-cookies --download-pdf

  # 浏览期刊最新文章
  python informs_scraper.py -m journal -j opre -n 200 --browser-cookies

  # 指定卷期目录
  python informs_scraper.py -m toc -j mnsc -v 71 -i 3 --browser-cookies

  # 使用会员账号密码
  python informs_scraper.py -m keyword -q "supply chain" -n 100 --member 123456 --password MyPwd

  # 高级搜索：指定作者 + 期刊 + 年份
  python informs_scraper.py -m advanced -q "reinforcement learning" -j ijoc --author "Powell" --date 2020-2024 -n 50
        """,
    )

    p.add_argument(
        "-m", "--mode",
        choices=["keyword", "journal", "toc", "advanced"],
        default="keyword",
        help="搜索模式（默认: keyword）",
    )
    p.add_argument("-q", "--query",   help="搜索关键词（keyword/advanced 模式）")
    p.add_argument("-j", "--journal", help="期刊代码，如 mnsc / opre / ijoc")
    p.add_argument("-v", "--volume",  type=int, help="卷号（toc 模式）")
    p.add_argument("-i", "--issue",   type=int, help="期号（toc 模式）")
    p.add_argument("--author",        help="作者姓名（advanced 模式）")
    p.add_argument("--date",          help="年份范围，如 2020-2024")
    p.add_argument(
        "-n", "--count", type=int, default=100,
        help="最多抓取篇数（默认: 100）",
    )
    p.add_argument(
        "--sort", choices=["relevance", "date"], default="relevance",
        help="排序方式（默认: relevance）",
    )

    # 认证
    auth = p.add_mutually_exclusive_group()
    auth.add_argument(
        "--chrome-login", action="store_true",
        help="★ 弹出 Chrome 窗口手动登录，自动提取 session cookie（最推荐）",
    )
    auth.add_argument(
        "--browser-cookies", action="store_true",
        help="从本机 Chrome 磁盘读取持久化 cookie（若遇 403 改用 --chrome-login）",
    )
    auth.add_argument(
        "--cookies-file", metavar="PATH",
        help="从 JSON 文件加载 cookie（EditThisCookie / Cookie Editor 导出）",
    )
    p.add_argument("--member",   help="INFORMS 会员号")
    p.add_argument("--password", help="账号密码（与 --member 配合使用）")

    # 输出
    p.add_argument(
        "-o", "--output-dir",
        default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "informs_result"),
        help="输出目录（默认: 脚本所在目录下的 informs_result/）",
    )
    p.add_argument(
        "--format", choices=["csv", "json", "xlsx"], default="csv",
        help="元数据保存格式（默认: csv）",
    )
    p.add_argument(
        "--download-pdf", action="store_true",
        help="搜索后自动下载 PDF",
    )
    p.add_argument(
        "--delay", type=float, nargs=2, metavar=("MIN", "MAX"), default=[2.0, 5.0],
        help="请求间隔（秒），默认 2 5",
    )

    return p


def _interactive_mode():
    """交互式向导（无参数运行时启动）。"""
    print("=" * 60)
    print("  INFORMS PubsOnLine 论文抓取工具 v1.0")
    print("=" * 60)
    print()
    print("搜索模式：")
    print("  1. 关键词搜索")
    print("  2. 按期刊浏览")
    print("  3. 指定卷期目录")
    print("  4. 高级搜索（组合条件）")
    mode_map = {"1": "keyword", "2": "journal", "3": "toc", "4": "advanced"}
    mode_choice = input("\n请选择（1-4，默认 1）: ").strip() or "1"
    mode = mode_map.get(mode_choice, "keyword")

    query      = None
    journal    = None
    volume     = None
    issue      = None
    author     = None
    date_range = None

    if mode in ("keyword", "advanced"):
        query = input("搜索关键词（支持 AND/OR/NOT）: ").strip()

    if mode == "journal":
        print(f"\n期刊代码：{', '.join(JOURNAL_NAMES.keys())}")
        journal = input("期刊代码: ").strip().lower()

    if mode == "toc":
        print(f"\n期刊代码：{', '.join(JOURNAL_NAMES.keys())}")
        journal = input("期刊代码: ").strip().lower()
        volume  = int(input("卷号（如 71）: ").strip())
        issue   = int(input("期号（如 3）: ").strip())

    if mode == "advanced":
        j = input("限定期刊代码（可选，回车跳过）: ").strip().lower()
        journal = j or None
        a = input("限定作者（可选，回车跳过）: ").strip()
        author = a or None
        d = input("年份范围（如 2020-2024，可选）: ").strip()
        date_range = d or None

    count = int(input("\n最多抓取多少篇（默认 100）: ").strip() or "100")

    print("\n认证方式：")
    print("  1. 弹出 Chrome 窗口手动登录（★ 最推荐，session cookie 最完整）")
    print("  2. 从 Chrome 磁盘读取 cookie（需先在普通 Chrome 登录过）")
    print("  3. 输入会员号 + 密码（程序自动登录）")
    print("  4. 不登录（仅抓取元数据，无法下载 PDF）")
    auth_choice = input("请选择（1-4，默认 1）: ").strip() or "1"

    use_chrome_login    = False
    use_browser_cookies = False
    member_id = password = None

    if auth_choice == "1":
        use_chrome_login = True
    elif auth_choice == "2":
        use_browser_cookies = True
    elif auth_choice == "3":
        member_id = input("会员号: ").strip()
        password  = input("密码: ").strip()

    dl_pdf = input("\n是否下载 PDF？（y/N）: ").strip().lower() == "y"
    _default_out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "informs_result")
    output_dir = (
        input(f"输出目录（默认 {_default_out}）: ").strip()
        or _default_out
    )
    fmt = input("保存格式（csv/json/xlsx，默认 csv）: ").strip().lower() or "csv"

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
    # 无参数 → 交互式向导
    if len(sys.argv) == 1:
        cfg = _interactive_mode()
    else:
        parser = _build_parser()
        args   = parser.parse_args()
        cfg = vars(args)

    # ── 初始化爬虫 ──────────────────────────────────────────────────────────
    scraper = INFORMSScraper(
        member_id          = cfg.get("member"),
        password           = cfg.get("password"),
        cookies_file       = cfg.get("cookies_file"),
        use_browser_cookies= cfg.get("browser_cookies", False),
        use_chrome_login   = cfg.get("chrome_login", False),
        delay_range        = tuple(cfg.get("delay", [2.0, 5.0])),
    )

    # ── 执行搜索 ─────────────────────────────────────────────────────────────
    mode    = cfg.get("mode", "keyword")
    results = []

    if mode == "keyword":
        if not cfg.get("query"):
            print("[错误] keyword 模式需要 -q / --query 参数")
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
            print("[错误] journal 模式需要 -j / --journal 参数")
            sys.exit(1)
        results = scraper.browse_journal(
            journal_code= cfg["journal"],
            count       = cfg.get("count", 100),
        )

    elif mode == "toc":
        if not all([cfg.get("journal"), cfg.get("volume"), cfg.get("issue")]):
            print("[错误] toc 模式需要 -j / --journal、-v / --volume、-i / --issue")
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
        print("\n没有抓取到任何文章，程序退出。")
        return

    # ── 保存元数据 ──────────────────────────────────────────────────────────
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

    # ── 下载 PDF ─────────────────────────────────────────────────────────────
    if cfg.get("download_pdf"):
        scraper.download_pdfs(results, output_dir)

    print("\n[全部完成]")


if __name__ == "__main__":
    main()
