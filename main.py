from __future__ import annotations

import re
from datetime import datetime, timezone
from typing import Any

from fastapi import FastAPI
from pydantic import BaseModel, Field

app = FastAPI(
    title="Robot Competition Orchestrator",
    version="0.1.0",
    description="智能机器人创意竞赛助手 V2 的中央编排器测试接口",
)

SESSION_STORE: dict[str, dict[str, Any]] = {}

class ChatRequest(BaseModel):
    session_id: str = Field(..., min_length=1, description="会话或任务流唯一标识")
    user_id: str = Field(default="anonymous", description="用户标识")
    message: str = Field(..., min_length=1, description="用户本次任意自然语言输入")
    attachments: list[dict[str, Any]] = Field(default_factory=list)

class ChatResponse(BaseModel):
    success: bool
    stage: str
    intent: str
    message: str
    data: dict[str, Any]
    files: list[dict[str, Any]]
    suggested_actions: list[str]
    version: int
    updated_at: str

def detect_intent(message: str, state: dict[str, Any]) -> str:
    text = message.strip()
    if re.search(r"(导出|下载).*(word|Word|文档)", text):
        return "export_word"
    if re.search(r"(导出|下载).*(ppt|PPT|演示)", text):
        return "export_pptx"
    if re.search(r"(第[一二三四五六七1-7]页|修改|改成|简化|突出|替换图片|重画)", text):
        return "modify_report"
    if re.search(r"(生成|写|制作).*(报告|七页|7页)", text):
        return "generate_report"
    if re.search(r"(重新生成|换一批|不满意|再来).*(题目|标题|名称)", text):
        return "regenerate_titles"
    if re.search(r"(选择|选|用|确定).*(第?[一二三123]|1|2|3).*(题目|标题|个)?", text):
        return "select_title"
    if not state.get("raw_idea"):
        return "create_project"
    return "supplement_idea"

def parse_title_index(message: str) -> int | None:
    mapping = {"一": 1, "二": 2, "三": 3}
    match = re.search(r"(?:第)?([一二三123])(?:个|项|题目|标题)?", message)
    if not match:
        return None
    token = match.group(1)
    return int(token) if token.isdigit() else mapping[token]

def get_or_create_state(session_id: str, user_id: str) -> dict[str, Any]:
    if session_id not in SESSION_STORE:
        SESSION_STORE[session_id] = {
            "session_id": session_id,
            "user_id": user_id,
            "stage": "new",
            "raw_idea": "",
            "candidate_titles": [],
            "selected_title": "",
            "evaluation": {},
            "report_json": {},
            "images": {},
            "version": 0,
            "history": [],
        }
    return SESSION_STORE[session_id]

@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok", "service": "robot_competition_orchestrator"}

@app.post("/api/v1/robot-competition/chat", response_model=ChatResponse)
def chat(req: ChatRequest) -> ChatResponse:
    state = get_or_create_state(req.session_id, req.user_id)
    intent = detect_intent(req.message, state)
    state["history"].append({"role": "user", "content": req.message})
    state["version"] += 1

    files: list[dict[str, Any]] = []
    data: dict[str, Any] = {}

    if intent == "create_project":
        state["raw_idea"] = req.message
        state["stage"] = "idea_received"
        reply = "已收到你的机器人创意。当前接口已成功识别为新建创意。"
        actions = ["生成3个候选题目", "补充使用场景或核心功能"]
    elif intent == "supplement_idea":
        state["raw_idea"] = f'{state.get("raw_idea", "")}\n补充要求：{req.message}'.strip()
        state["stage"] = "idea_updated"
        reply = "已把本次内容作为创意补充要求保存。"
        actions = ["生成3个候选题目", "继续补充要求"]
    elif intent == "regenerate_titles":
        state["stage"] = "title_regeneration_requested"
        reply = "已识别为重新生成候选题目。"
        actions = ["突出差异化重新生成", "补充新的命名要求"]
    elif intent == "select_title":
        index = parse_title_index(req.message)
        titles = state.get("candidate_titles", [])
        if index and len(titles) >= index:
            state["selected_title"] = titles[index - 1]
            reply = f"已选择第{index}个题目：{state['selected_title']}"
        else:
            reply = "已识别为题目选择，但当前测试版本还没有生成候选题目数据。"
        state["stage"] = "title_selected"
        data["selected_index"] = index
        actions = ["生成7页报告", "修改题目名称"]
    elif intent == "generate_report":
        state["stage"] = "report_generation_requested"
        reply = "已识别为生成7页报告。"
        actions = ["生成报告", "先查看评分依据"]
    elif intent == "modify_report":
        state["stage"] = "report_modification_requested"
        reply = "已识别为局部修改请求。"
        actions = ["确认修改", "查看当前报告版本"]
    elif intent == "export_word":
        state["stage"] = "word_export_requested"
        reply = "已识别为Word导出请求。"
        actions = ["导出Word", "同时导出PPT"]
    elif intent == "export_pptx":
        state["stage"] = "pptx_export_requested"
        reply = "已识别为PPTX导出请求。"
        actions = ["导出PPT", "同时导出Word"]
    else:
        state["stage"] = "unknown"
        reply = "暂未识别本次操作，请补充说明。"
        actions = ["新建创意", "生成题目", "修改报告", "导出文件"]

    state["history"].append({"role": "assistant", "content": reply})
    updated_at = datetime.now(timezone.utc).isoformat()
    data.update({
        "raw_idea": state.get("raw_idea", ""),
        "selected_title": state.get("selected_title", ""),
        "current_stage": state["stage"],
    })

    return ChatResponse(
        success=True,
        stage=state["stage"],
        intent=intent,
        message=reply,
        data=data,
        files=files,
        suggested_actions=actions,
        version=state["version"],
        updated_at=updated_at,
    )
