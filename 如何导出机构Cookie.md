# 如何导出机构 Cookie 以使用学校账号抓取 ScienceDirect

## 为什么需要 Cookie？

ScienceDirect 的部分文章（摘要、全文）需要机构订阅权限。通过导出浏览器已登录的 Cookie，
抓取脚本就能"继承"你的机构权限，从而访问完整内容。

---

## 步骤一：在浏览器中通过学校账号登录

1. 打开 Chrome 或 Firefox，进入 https://www.sciencedirect.com
2. 点击右上角 **"Sign in"**
3. 选择 **"Access through your institution"（通过机构访问）**
4. 搜索你的学校名称，点击跳转
5. 在学校的统一身份认证页面用学号/工号登录
6. 成功返回 ScienceDirect 后，页面右上角会显示你的机构名称，说明登录成功

---

## 步骤二：安装 Cookie Editor 插件

- **Chrome**: 在 Chrome 应用商店搜索 "Cookie Editor"（by cgagnier），安装
- **Firefox**: 在 Firefox 附加组件页面搜索 "Cookie Editor"，安装

---

## 步骤三：导出 Cookie

1. 保持在 `sciencedirect.com` 页面
2. 点击浏览器工具栏中的 Cookie Editor 图标
3. 点击底部的 **"Export"** 按钮
4. 选择 **"Export as JSON"**
5. 将内容粘贴到一个新文件，命名为 `cookies.json`
6. 把 `cookies.json` 放到和 `sd_scraper.py` 同一个目录下

---

## 步骤四：运行抓取脚本时指定 Cookie 文件

```bash
# 示例：用机构账号抓取关键词 "turbine blade" 的 200 篇文章
python sd_scraper.py -m keyword -q "turbine blade" -n 200 --cookies cookies.json

# 交互式向导也支持输入 cookie 文件路径
python sd_scraper.py --interactive
```

---

## 注意事项

- Cookie 有时效性，一般几天到几周不等。如果抓取时报 403 错误，请重新导出 Cookie。
- 请勿将 `cookies.json` 上传到公开平台（如 GitHub），以免账号被他人盗用。
- 仅用于个人学术研究，请遵守所在机构和 ScienceDirect 的使用条款。
