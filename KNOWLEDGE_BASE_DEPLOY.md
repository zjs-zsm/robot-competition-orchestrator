# 机器人竞赛知识库服务

这是独立知识库服务，部署在 `knowledge-base-service` 分支，不修改 V4.1.0 主分支。

## Render 配置

- Runtime: Python
- Branch: `knowledge-base-service`
- Build command: `pip install -r requirements-kb.txt`
- Start command: `uvicorn knowledge_base:app --host 0.0.0.0 --port $PORT`
- 推荐连接 Render PostgreSQL，并在服务环境变量中设置其 Internal Database URL。

## 必需环境变量

- `SESSION_SECRET`: 至少 32 字符的随机秘密，用于签名登录 cookie。
- `ADMIN_USERNAME`: 管理员账号。
- `ADMIN_PASSWORD`: 管理员强密码。
- `TEACHER_USERNAME`: 老师账号。
- `TEACHER_PASSWORD`: 老师强密码。
- `KB_API_KEY`: V4.1.0 调用检索 API 的长随机密钥。
- `DATABASE_URL`: Render PostgreSQL 的 Internal Database URL（推荐）。未设置时使用 SQLite，但临时磁盘数据可能在重新部署/重启后丢失。
- `COOKIE_SECURE`: 公网 HTTPS 保持 true。

## API

- `GET /health`
- `POST /api/v1/kb/upload` (登录 cookie)
- `GET /api/v1/kb/documents` (登录 cookie)
- `DELETE /api/v1/kb/documents/{id}` (登录 cookie)
- `POST /api/v1/knowledge/search` (Bearer KB_API_KEY 或已登录 cookie)

## 注意

1. 上传后直接入库，无人工审核。
2. 文件大小默认 25 MB，单次最多 30 个。
3. 当前使用关键词/字符 n-gram 检索，不是语义向量检索。
4. 扫描版 PDF 和图片会保存，但当前版本尚未接 OCR，因此不能从纯图片中自动检索文字。
5. 公开部署前必须设置独立且强的账号密码、SESSION_SECRET 和 KB_API_KEY。
