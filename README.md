# 域名整理工具

一个轻量 Flask 应用：登录后可上传 TXT/LOG 文件或直接粘贴文本，服务端按业务分类整理并返回 Excel。

## 本地启动（PowerShell）

```powershell
cd <项目目录>
py -3.10 -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt

$env:TINKER_USERNAME = "admin"
$env:TINKER_PASSWORD = "请修改为安全密码"
$env:TINKER_SECRET_KEY = "请使用足够长的随机字符串"
python app.py
```

浏览器访问 <http://127.0.0.1:5000>。
应用没有内置默认密码；未设置 `TINKER_PASSWORD` 时会拒绝启动。

## Docker 启动

首次使用时复制环境变量示例，并在 `.env` 中填写密码和随机密钥：

```powershell
Copy-Item .env.example .env
# 编辑 .env，填写 TINKER_PASSWORD 和 TINKER_SECRET_KEY
docker compose up -d --build
```

浏览器访问 <http://127.0.0.1:5000>。用户名默认为 `admin`，密码使用 `.env` 中的 `TINKER_PASSWORD`。

查看运行状态和日志：

```powershell
docker compose ps
docker compose logs -f
```

停止服务：

```powershell
docker compose down
```

`.env` 不会被提交。`TINKER_SECRET_KEY` 应使用足够长的随机字符串；启用 HTTPS 后，将 `TINKER_SECURE_COOKIE` 设置为 `1`。

### 通过代理构建镜像

Docker daemon 的代理只负责拉取基础镜像，Dockerfile 中的 `pip install` 还需要构建代理。可在服务器未提交的 `.env` 中配置：

```dotenv
HTTP_PROXY=http://proxy-host:port
HTTPS_PROXY=http://proxy-host:port
NO_PROXY=localhost,127.0.0.1
```

然后重新构建：

```bash
docker compose build --no-cache
docker compose up -d
```

如果基础镜像也无法拉取，需要同时为 Linux Docker daemon 配置代理。执行 `sudo systemctl edit docker`，填写：

```ini
[Service]
Environment="HTTP_PROXY=http://proxy-host:port"
Environment="HTTPS_PROXY=http://proxy-host:port"
Environment="NO_PROXY=localhost,127.0.0.1"
```

保存后应用配置：

```bash
sudo systemctl daemon-reload
sudo systemctl restart docker
sudo systemctl show --property=Environment docker
```

变量名和代理 URL 必须是纯文本，不要包含 Markdown 的反斜杠或 `[地址](地址)` 形式。代理凭据和真实内网地址不要提交到仓库。

## 生产运行

Windows 可以使用 Waitress：

```powershell
waitress-serve --call --listen=0.0.0.0:5000 app:create_app
```

对外部署时应启用 HTTPS，并设置：

```powershell
$env:TINKER_SECURE_COOKIE = "1"
```

## HTTP 接口

- `POST /api/login`：用户名和密码登录。
- `POST /api/convert`：通过 `multipart/form-data` 传入 `file` 或 `text`，以及可选的 `title`，响应为 Excel 文件。
- `POST /api/logout`：退出当前会话。

上传内容最大 2 MB，目前支持 UTF-8、UTF-8 BOM 和 GB18030 编码的 `.txt`、`.log` 文件。

系统优先使用内置业务分类。无法识别内置分类时，会将域名前一行的第一个字段作为动态分类，例如 `DEMO-SEO 备案域名` 会归入 `DEMO-SEO`。
