# 简历优化工具

一个基于 Flask + DeepSeek 的中文简历优化 Web 应用。支持上传 PDF/DOCX 简历，粘贴或上传 JD，结合岗位理解、匹配分析、过往成果材料和 OCR，生成优化后的 Word 简历。

## 功能

- 上传 PDF 或 DOCX 简历。
- 粘贴岗位 JD，或上传岗位页面/JD 文件。
- 支持图片型 PDF/JD 截图 OCR 识别。
- 上传过往成果材料，支持 PDF、DOCX、PPTX、TXT、MD、CSV、JSON、HTML。
- 生成优化后的 Word 文件。
- 展示岗位理解、匹配分析、具体不匹配诊断、修改清单。
- 可选访问口令，方便给朋友使用时做简单保护。
- 自动清理 24 小时前生成的 Word 文件。

## 本地安装

```powershell
pip install -r requirements.txt
```

复制环境变量示例：

```powershell
copy .env.example .env
```

编辑 `.env`：

```text
DEEPSEEK_API_KEY=你的 DeepSeek API Key
DEEPSEEK_API_BASE=https://api.deepseek.com/v1
DEEPSEEK_MODEL=deepseek-chat
DEEPSEEK_REQUEST_TIMEOUT=60

APP_PASSWORD=给朋友的访问口令
APP_SECRET_KEY=一段足够长的随机字符串
OUTPUT_TTL_SECONDS=86400
```

如果不设置 `APP_PASSWORD`，本地访问不会要求登录。

## 本地运行

```powershell
python app.py
```

浏览器访问：

```text
http://127.0.0.1:5000/
```

## 给同一 Wi-Fi 下的朋友试用

把 `app.py` 最后一行改成：

```python
app.run(host="0.0.0.0", port=5000, debug=False)
```

启动后用 `ipconfig` 查看你的 IPv4 地址，例如：

```text
192.168.1.23
```

朋友访问：

```text
http://192.168.1.23:5000/
```

注意：需要 Windows 防火墙允许 Python 访问网络。

## 部署到 Render

1. 把项目推送到 GitHub。
2. 在 Render 新建 Web Service。
3. Build Command:

```bash
pip install -r requirements.txt
```

4. Start Command:

```bash
gunicorn app:app
```

5. 设置环境变量：

```text
DEEPSEEK_API_KEY
DEEPSEEK_API_BASE
DEEPSEEK_MODEL
DEEPSEEK_REQUEST_TIMEOUT
APP_PASSWORD
APP_SECRET_KEY
OUTPUT_TTL_SECONDS
```

## 部署到 Railway

1. 把项目推送到 GitHub。
2. Railway 选择从 GitHub 部署。
3. Railway 会读取 `Procfile`：

```text
web: gunicorn app:app
```

4. 在 Variables 中设置和 Render 相同的环境变量。

## 部署到腾讯云/阿里云轻量服务器

服务器上安装 Python 后：

```bash
git clone <your-repo-url>
cd <your-repo>
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

设置环境变量后启动：

```bash
gunicorn -w 2 -b 0.0.0.0:5000 app:app
```

生产环境建议再用 Nginx 反向代理到 80/443，并配置 HTTPS。

## 隐私建议

- 不要把 `.env`、API Key、个人配置提交到 GitHub。
- 给朋友使用时一定设置 `APP_PASSWORD`。
- 生成文件默认 24 小时后自动清理，可通过 `OUTPUT_TTL_SECONDS` 修改。
- 简历和 JD 属于敏感信息，不建议公开无密码访问。

