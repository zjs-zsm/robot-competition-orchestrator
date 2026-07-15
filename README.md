# 智能机器人创意竞赛助手 V2：中央编排器 v0.1

第一阶段先打通：

超星任意自然语言输入 → 自建 HTTP 插件 → 统一 JSON 返回。

## 接口

- `GET /health`
- `POST /api/v1/robot-competition/chat`

请求示例：

```json
{
  "session_id": "test-session-001",
  "user_id": "test-user",
  "message": "我想设计一个面向独居老人的家庭陪伴机器人",
  "attachments": []
}
```

## 本地启动

```bash
pip install -r requirements.txt
uvicorn main:app --reload
```

## Render 部署参数

- Language: Python 3
- Build Command: `pip install -r requirements.txt`
- Start Command: `uvicorn main:app --host 0.0.0.0 --port $PORT`
