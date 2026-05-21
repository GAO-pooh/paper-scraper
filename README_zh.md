<div align="right">

[English](README.md) | [简体中文](README_zh.md)

</div>

# 学术论文抓取工具

针对 **ScienceDirect** 和 **INFORMS PubsOnLine** 的自动化爬虫——支持按关键词、作者、期刊搜索，并通过机构账号批量下载 PDF。

---

## 支持平台

| 脚本 | 平台 | PDF 下载方式 |
|------|------|-------------|
| `sd_scraper.py` / `sd_scraper_en.py` | [ScienceDirect](https://www.sciencedirect.com)（Elsevier） | Chrome DevTools 协议（绕过 Cloudflare） |
| `informs_scraper.py` / `informs_scraper_en.py` | [INFORMS PubsOnLine](https://pubsonline.informs.org) | Session Cookie 直连下载 |

`_en.py` 结尾为英文界面版，其余为中文界面版，逻辑完全相同。

---

## 功能特点

- **多种搜索模式**：关键词、期刊浏览、期刊内关键词、作者、ISSN、高级搜索（任意组合）
- **机构账号批量下载 PDF**（支持高校 SSO / CARSI 登录）
- **反爬绕过**：`curl_cffi` 模拟 Chrome TLS 指纹、隐身 JS 注入、自动处理限速封锁
- **ScienceDirect**：Chrome DevTools Protocol (CDP) 直接捕获 PDF 字节，无需 Playwright
- **INFORMS**：Atypon 平台 PDF 端点无 JS 验证，Session Cookie 直连即可
- 输出格式：**CSV / JSON / XLSX**
- 支持交互式向导（无参数运行即启动）

---

## 安装

**第一步 — 安装依赖**

```bash
pip install curl_cffi websocket-client browser-cookie3 openpyxl
```

如需使用 INFORMS 爬虫，额外安装：

```bash
pip install beautifulsoup4 lxml
```

**第二步 — 运行**

```bash
# ScienceDirect（交互式向导）
python3 sd_scraper.py

# INFORMS（交互式向导）
python3 informs_scraper.py
```

> **目前仅支持 macOS**（Chrome 路径硬编码为 `/Applications/Google Chrome.app`）。  
> Linux / Windows 用户需修改脚本类中的 `CHROME_BIN` 变量。

完整依赖列表见 [`requirements.txt`](requirements.txt)。

---

## 快速上手

### ScienceDirect

```bash
# 交互式向导（首次使用推荐）
python sd_scraper.py

# 关键词搜索，保存为 XLSX
python sd_scraper.py -m keyword -q "machine learning" -n 100 --browser-cookies

# 关键词搜索 + 下载 PDF
python sd_scraper.py -m keyword -q "deep learning" -n 50 --browser-cookies --download-pdfs

# 浏览期刊（按时间倒序）
python sd_scraper.py -m journal -j "Energy" -n 200 --browser-cookies --sort date

# 在指定期刊内按关键词搜索
python sd_scraper.py -m journal_keyword -j "Renewable Energy" -q "solar cell" -n 50 --browser-cookies

# 按作者搜索
python sd_scraper.py -m author -a "Zhang Wei" -n 30 --browser-cookies

# 高级搜索（组合条件）
python sd_scraper.py -m advanced -q "deep learning" --date 2021-2024 --type REV -n 50 --browser-cookies --download-pdfs
```

### INFORMS PubsOnLine

```bash
# 交互式向导
python informs_scraper.py

# 关键词搜索
python informs_scraper.py -m keyword -q "supply chain" -n 100 --browser-cookies

# 浏览期刊
python informs_scraper.py -m journal -j mnsc -n 200 --browser-cookies

# 指定卷期目录
python informs_scraper.py -m toc -j mnsc -v 71 -i 3 --browser-cookies

# 关键词搜索 + 下载 PDF
python informs_scraper.py -m keyword -q "inventory" -n 50 --browser-cookies --download-pdf

# 最推荐的登录方式：弹出 Chrome 窗口手动登录
python informs_scraper.py -m keyword -q "machine learning" -n 50 --chrome-login --download-pdf
```

#### INFORMS 期刊代码

| 代码 | 期刊名称 |
|------|---------|
| `mnsc` | Management Science |
| `opre` | Operations Research |
| `ijoc` | INFORMS Journal on Computing |
| `mksc` | Marketing Science |
| `msom` | Manufacturing & Service Operations Management |
| `trsc` | Transportation Science |
| `isre` | Information Systems Research |
| `orsc` | Organization Science |

---

## PDF 下载原理

### ScienceDirect

Cloudflare 会拦截对 PDF 端点的直接 HTTP 请求。本工具通过 Chrome DevTools Protocol (CDP) 绕过这一机制：

1. 脚本自动以调试模式启动 Chrome（复用现有 Profile，无需重新登录）
2. 在文章页面建立 Cookie 上下文
3. 通过 `Network` / `Fetch` DevTools 事件拦截 PDF 响应字节
4. 直接写入磁盘，无 Save 对话框，无需 Playwright

**首次使用**：用 Chrome 通过机构账号（SSO / CARSI）登录 ScienceDirect，并打开任意一篇有权限的文章 PDF 确认访问正常。之后脚本全自动处理。

```bash
# 若 Chrome 尚未登录，先运行：
python sd_scraper.py --open-browser-login
# 再执行抓取：
python sd_scraper.py -m keyword -q "turbine" -n 50 --browser-cookies --download-pdfs
```

### INFORMS

INFORMS（Atypon 平台）的 PDF 端点没有 JS 验证，有效的 Session Cookie 即可直连下载。

```bash
# 方式一：从 Chrome 读取 Cookie（需提前在 Chrome 中登录）
python informs_scraper.py -m keyword -q "inventory" -n 30 --browser-cookies --download-pdf

# 方式二：弹出 Chrome 窗口手动登录，自动提取 Cookie（最推荐）
python informs_scraper.py -m keyword -q "inventory" -n 30 --chrome-login --download-pdf

# 方式三：会员账号直接登录
python informs_scraper.py -m keyword -q "inventory" -n 30 --member 123456 --password MyPwd --download-pdf
```

---

## 输出目录结构

默认输出到 `./results/`（ScienceDirect）或 `./informs_result/`（INFORMS）。

```
results/
└── keyword_machine_learning_20250101_120000/
    ├── keyword_machine_learning_20250101_120000.xlsx   ← 元数据
    └── pdfs/
        ├── 001_Zhang_2024_Deep learning for...pdf
        ├── 002_Li_2023_Transfer learning in...pdf
        └── ...
```

---

## 完整命令行参数

### ScienceDirect（`sd_scraper.py`）

| 参数 | 说明 |
|------|------|
| `-m`, `--mode` | `keyword` / `journal` / `journal_keyword` / `author` / `issn` / `advanced` |
| `-q`, `--query` | 搜索关键词（支持 `AND` / `OR` / `NOT`） |
| `-j`, `--journal` | 期刊名称 |
| `-a`, `--author` | 作者姓名 |
| `--issn` | 期刊 ISSN |
| `-n`, `--count` | 最大抓取数量（默认 50） |
| `--date` | 年份范围，如 `2020-2024` |
| `--sort` | 排序：`relevance`（默认）/ `date` |
| `--type` | 文章类型：`FLA` / `REV` / `SCO` |
| `--browser-cookies` | 自动从本机 Chrome 读取 Cookie |
| `--cookies` | 指定 Cookie JSON 文件路径 |
| `--format` | 输出格式：`xlsx`（默认）/ `csv` / `json` / `all` |
| `--download-pdfs` | 搜索后下载 PDF |
| `--output` | 自定义输出目录 |
| `--open-browser-login` | 打开 Chrome 供机构登录 |
| `--interactive` | 启动交互式向导 |

### INFORMS（`informs_scraper.py`）

| 参数 | 说明 |
|------|------|
| `-m`, `--mode` | `keyword` / `journal` / `toc` / `advanced` |
| `-q`, `--query` | 搜索关键词 |
| `-j`, `--journal` | 期刊代码（如 `mnsc`） |
| `-v`, `--volume` | 卷号（toc 模式） |
| `-i`, `--issue` | 期号（toc 模式） |
| `--author` | 作者姓名（advanced 模式） |
| `--date` | 年份范围，如 `2020-2024` |
| `-n`, `--count` | 最大抓取数量（默认 100） |
| `--chrome-login` | ★ 弹出 Chrome 手动登录（最推荐） |
| `--browser-cookies` | 从本机 Chrome 读取 Cookie |
| `--cookies-file` | 指定 Cookie JSON 文件路径 |
| `--member` | INFORMS 会员号 |
| `--password` | 账号密码 |
| `--format` | `csv`（默认）/ `json` / `xlsx` |
| `--download-pdf` | 搜索后下载 PDF |
| `-o`, `--output-dir` | 输出目录 |

---

## 注意事项

- **需要机构订阅权限**才能下载完整 PDF。开放获取文章无需登录。
- **请求限速**：脚本在每次请求之间加入随机延迟（默认 2–5 秒），请勿降低此数值。
- **仅支持 macOS**：Chrome 路径硬编码为 `/Applications/Google Chrome.app/...`。Linux / Windows 用户需修改 `CHROME_BIN`。
- **Cookie 有时效性**：Session Cookie 通常几天到几周后过期。遇到 403 错误时重新登录即可。
- 请遵守所在机构及各出版商的使用条款，仅用于个人学术研究。

---

## 开源协议

MIT
