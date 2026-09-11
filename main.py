from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import base64
import hashlib
import io
import json
import os
import random
import re
import sqlite3
import textwrap
import urllib.request
import urllib.error
import urllib.parse

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Cm, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image, ImageDraw, ImageFont

APP_VERSION = "2.0.0"

app = FastAPI(
    title="Robot Competition Orchestrator",
    version=APP_VERSION,
    description="智能机器人创意竞赛助手 V1.0 正式版中央编排器",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ============================================================
# 数据模型：保持与学习通插件现有字段兼容
# ============================================================
class ChatRequest(BaseModel):
    session_id: str = Field(..., description="会话ID")
    user_id: str = Field(..., description="用户ID")
    message: str = Field(..., description="用户输入")
    attachments: Optional[List[Dict[str, Any]]] = Field(default_factory=list)


class ChatResponse(BaseModel):
    success: bool
    stage: str
    intent: str
    message: str
    payload_json: str = ""
    download_url: str = ""
    data: Dict[str, Any] = Field(default_factory=dict)
    files: List[Dict[str, Any]] = Field(default_factory=list)
    suggested_actions: List[str] = Field(default_factory=list)
    version: int = 10
    updated_at: str


# ============================================================
# 基础配置
# ============================================================
PUBLIC_BASE_URL = os.getenv(
    "PUBLIC_BASE_URL",
    "https://robot-competition-orchestrator.onrender.com",
).rstrip("/")

DATA_DIR = Path(os.getenv("DATA_DIR", "/tmp/robot_competition_data"))
DATA_DIR.mkdir(parents=True, exist_ok=True)

EXPORT_DIR = Path(os.getenv("EXPORT_DIR", str(DATA_DIR / "exports")))
EXPORT_DIR.mkdir(parents=True, exist_ok=True)

SESSION_DB_PATH = os.getenv(
    "SESSION_DB_PATH",
    str(DATA_DIR / "sessions.sqlite3"),
)

# 可选：任意 OpenAI-compatible Chat Completions 服务。
# 没有配置时，系统会自动使用本地动态生成，不影响完整流程。
LLM_API_KEY = os.getenv("LLM_API_KEY", "").strip()
LLM_API_BASE = os.getenv("LLM_API_BASE", "").strip().rstrip("/")
LLM_MODEL = os.getenv("LLM_MODEL", "").strip()
LLM_TIMEOUT = int(os.getenv("LLM_TIMEOUT", "45"))

# 未来老师资料知识库接口：目前可留空。
KNOWLEDGE_API_URL = os.getenv("KNOWLEDGE_API_URL", "").strip()
KNOWLEDGE_API_KEY = os.getenv("KNOWLEDGE_API_KEY", "").strip()

# 可选图片生成/检索接口。留空时自动生成项目相关结构示意图。
IMAGE_API_URL = os.getenv("IMAGE_API_URL", "").strip()
IMAGE_API_KEY = os.getenv("IMAGE_API_KEY", "").strip()

# 推荐图像生成服务。配置后可自动生成“产品效果图 + 应用场景图”。
# 未配置时仍会生成本地概念插画与原生Word架构/流程可视化，不影响主流程。
POLLINATIONS_API_KEY = os.getenv("POLLINATIONS_API_KEY", "").strip()
POLLINATIONS_IMAGE_MODEL = os.getenv("POLLINATIONS_IMAGE_MODEL", "flux").strip() or "flux"
IMAGE_TIMEOUT = int(os.getenv("IMAGE_TIMEOUT", "90"))


# ============================================================
# 样例案例库：仅在没有真实知识库时用于流程验证
# ============================================================
REFERENCE_CASES = [
    {
        "title": "智能老人陪伴与健康监测机器人",
        "keywords": ["老人", "陪伴", "健康", "跌倒", "吃药", "家庭", "语音"],
        "track": "服务机器人",
    },
    {
        "title": "校园智能垃圾分类与回收机器人",
        "keywords": ["垃圾分类", "校园", "回收", "识别", "环保", "移动"],
        "track": "服务机器人",
    },
    {
        "title": "公共空间智能消毒巡检机器人",
        "keywords": ["消毒", "巡检", "公共空间", "环境", "导航", "安全"],
        "track": "特种机器人",
    },
    {
        "title": "导盲辅助与道路安全提醒机器人",
        "keywords": ["导盲", "道路", "安全", "避障", "语音", "辅助"],
        "track": "服务机器人",
    },
    {
        "title": "农业果蔬采摘与成熟度识别机器人",
        "keywords": ["农业", "采摘", "果蔬", "成熟度", "识别", "机械臂"],
        "track": "农业机器人",
    },
    {
        "title": "仓储物流自主搬运机器人",
        "keywords": ["物流", "仓储", "搬运", "路径规划", "导航", "调度"],
        "track": "工业机器人",
    },
    {
        "title": "水质检测与河道巡航机器人",
        "keywords": ["水质", "检测", "河道", "巡航", "传感器", "环保"],
        "track": "特种机器人",
    },
    {
        "title": "家庭教育陪伴与学习监督机器人",
        "keywords": ["儿童", "教育", "学习", "陪伴", "语音", "家庭"],
        "track": "服务机器人",
    },
]


# ============================================================
# 通用工具
# ============================================================
def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", "", (text or "").strip()).lower()


def clean_message(text: str) -> str:
    result = (text or "").strip()
    prefixes = [
        "启动V2中央编排器",
        "启动v2中央编排器",
        "启动机器人竞赛助手",
        "启动达标版竞赛助手",
        "运行机器人竞赛中央编排器",
    ]
    for prefix in prefixes:
        result = result.replace(prefix, "").strip()
    result = re.sub(r"/no_think\b", "", result, flags=re.I).strip()
    return result if result else (text or "").strip()


def contains_any(text: str, words: List[str]) -> bool:
    lowered = (text or "").lower()
    return any((word or "").lower() in lowered for word in words)


def clamp(value: float, low: float = 0, high: float = 100) -> float:
    return max(low, min(high, value))


def stable_seed(text: str) -> int:
    return int(hashlib.sha1(text.encode("utf-8")).hexdigest()[:8], 16)


def unique(items: List[str]) -> List[str]:
    return list(dict.fromkeys([x for x in items if x]))


def safe_filename(text: str, limit: int = 40) -> str:
    text = re.sub(r"[^\w\u4e00-\u9fff-]+", "_", text or "report")
    return text.strip("_")[:limit] or "report"


# ============================================================
# 会话持久化：SQLite，避免进程内变量导致短暂重启后完全丢失
# ============================================================
def _db() -> sqlite3.Connection:
    conn = sqlite3.connect(SESSION_DB_PATH, timeout=10)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS sessions (
            session_key TEXT PRIMARY KEY,
            payload TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )
    return conn


def build_session_key(req: ChatRequest) -> str:
    user_id = (req.user_id or "").strip()
    session_id = (req.session_id or "").strip()

    placeholder_users = {
        "", "anonymous", "current_user", "user_456", "chaoxing-test-user"
    }
    placeholder_sessions = {
        "", "default", "current_session", "current_session_123", "chaoxing-test-001"
    }

    if user_id in placeholder_users:
        user_id = "anonymous"
    if session_id in placeholder_sessions:
        session_id = "default"

    return f"{user_id}:{session_id}"


def load_session(session_key: str) -> Dict[str, Any]:
    try:
        with _db() as conn:
            row = conn.execute(
                "SELECT payload FROM sessions WHERE session_key = ?",
                (session_key,),
            ).fetchone()
        if row:
            return json.loads(row[0])
    except Exception as exc:
        print("SESSION_LOAD_ERROR:", repr(exc), flush=True)
    return {}


def save_session(session_key: str, session: Dict[str, Any]) -> None:
    try:
        payload = json.dumps(session, ensure_ascii=False)
        with _db() as conn:
            conn.execute(
                """
                INSERT INTO sessions(session_key, payload, updated_at)
                VALUES(?, ?, ?)
                ON CONFLICT(session_key) DO UPDATE SET
                    payload = excluded.payload,
                    updated_at = excluded.updated_at
                """,
                (session_key, payload, now_iso()),
            )
    except Exception as exc:
        print("SESSION_SAVE_ERROR:", repr(exc), flush=True)


# ============================================================
# 可选大模型调用
# ============================================================
def llm_available() -> bool:
    return bool(LLM_API_KEY and LLM_API_BASE and LLM_MODEL)


def call_llm_json(system_prompt: str, user_prompt: str) -> Optional[Dict[str, Any]]:
    if not llm_available():
        return None

    url = LLM_API_BASE
    if not url.endswith("/chat/completions"):
        url = url + "/chat/completions"

    body = {
        "model": LLM_MODEL,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": 0.7,
        "response_format": {"type": "json_object"},
    }

    req = urllib.request.Request(
        url,
        data=json.dumps(body, ensure_ascii=False).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {LLM_API_KEY}",
            "Content-Type": "application/json",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=LLM_TIMEOUT) as resp:
            raw = json.loads(resp.read().decode("utf-8"))
        content = raw["choices"][0]["message"]["content"]
        content = re.sub(r"^```(?:json)?|```$", "", content.strip(), flags=re.I).strip()
        return json.loads(content)
    except Exception as exc:
        print("LLM_CALL_ERROR:", repr(exc), flush=True)
        return None


# ============================================================
# 创意理解：本地动态解析 + 可选 LLM 增强
# ============================================================
SCENE_RULES = {
    "家庭/居家": ["家庭", "居家", "家里", "家用", "卧室", "客厅", "厨房", "阳台"],
    "校园/教育": ["校园", "学校", "教室", "学生", "宿舍", "实验室"],
    "医疗/养老": ["医院", "养老院", "老人", "康复", "护理", "患者"],
    "工业/仓储": ["工厂", "工业", "仓库", "物流", "产线", "搬运"],
    "农业/园艺": ["农业", "农田", "果园", "温室", "花盆", "植物", "多肉"],
    "环保/水域": ["水质", "河道", "湖泊", "环保", "垃圾", "污染"],
    "公共服务": ["社区", "商场", "车站", "公共", "道路", "景区"],
    "宠物照护": ["宠物", "猫", "狗", "猫砂", "喂食", "陪玩"],
}

TARGET_RULES = {
    "老年人": ["老人", "老年", "独居", "空巢"],
    "儿童/学生": ["儿童", "孩子", "学生", "小朋友"],
    "残障/特殊人群": ["残障", "盲人", "视障", "行动不便", "听障"],
    "宠物家庭": ["宠物", "猫", "狗"],
    "家庭用户": ["家庭", "家用", "居家", "家里"],
    "农业从业者": ["农民", "农业", "果园", "温室"],
    "工业/物流人员": ["工厂", "仓库", "物流", "工人"],
}

FUNCTION_RULES = {
    "用药提醒": ["吃药", "服药", "用药", "药物提醒"],
    "跌倒/姿态检测": ["跌倒", "摔倒", "姿态"],
    "语音交互与陪伴": ["语音", "聊天", "对话", "陪伴"],
    "健康监测": ["心率", "血压", "体温", "健康", "血氧"],
    "危险预警": ["危险", "报警", "异常", "烟雾", "燃气", "关火", "高温", "火灾"],
    "自主导航与避障": ["导航", "避障", "巡航", "移动", "跟随"],
    "视觉识别": ["识别", "视觉", "摄像头", "图像", "目标检测"],
    "机械执行": ["机械臂", "抓取", "搬运", "递送", "开关", "关闭燃气"],
    "远程通知": ["远程", "通知", "手机", "家属", "推送"],
    "智能家居联动": ["智能家居", "灯光", "空调", "门锁", "燃气阀"],
    "清洁整理": ["清洁", "扫地", "拖地", "收纳", "整理"],
    "宠物照护": ["猫砂", "补粮", "喂食", "陪玩", "宠物"],
    "植物养护": ["浇水", "植物", "多肉", "光照", "土壤", "施肥"],
    "垃圾分类与回收": ["垃圾分类", "回收", "垃圾"],
    "水质/环境检测": ["水质", "空气", "温湿度", "污染", "环境监测"],
    "学习辅助": ["学习", "作业", "教学", "辅导", "监督"],
}

TECH_RULES = {
    "多模态传感器融合": ["传感器", "跌倒", "健康", "烟雾", "高温", "燃气", "环境", "检测"],
    "计算机视觉识别": ["摄像头", "视觉", "图像", "识别", "目标检测", "姿态"],
    "语音识别与自然语言交互": ["语音", "聊天", "对话", "陪伴", "提醒"],
    "移动底盘与路径规划": ["移动", "导航", "避障", "巡航", "跟随"],
    "机械臂/执行机构控制": ["机械臂", "抓取", "搬运", "递送", "开关", "阀"],
    "物联网与远程通信": ["远程", "通知", "手机", "家属", "智能家居", "联网"],
    "边缘计算与本地决策": ["本地", "低延迟", "离线", "实时", "异常"],
    "智能规划与任务调度": ["自动", "规划", "调度", "定时", "任务"],
}


def _match_rules(text: str, rules: Dict[str, List[str]]) -> List[str]:
    return [name for name, keys in rules.items() if contains_any(text, keys)]


def _extract_free_keywords(raw: str) -> List[str]:
    cleaned = re.sub(r"[，。！？；、,.!?;:\n\t（）()\[\]【】]", " ", raw)
    parts = [p.strip() for p in cleaned.split() if 2 <= len(p.strip()) <= 12]
    stop = {"我想", "一个", "机器人", "智能机器人", "主要", "用于", "希望", "可以", "能够", "这个", "一种"}
    return unique([p for p in parts if p not in stop])[:12]


def local_extract_idea_fields(raw_idea: str) -> Dict[str, Any]:
    text = normalize_text(raw_idea)
    targets = _match_rules(text, TARGET_RULES) or ["目标场景用户"]
    scenes = _match_rules(text, SCENE_RULES) or ["实际应用场景"]
    functions = _match_rules(text, FUNCTION_RULES)
    tech = _match_rules(text, TECH_RULES)

    if not functions:
        functions = ["智能感知", "人机交互", "自主服务"]
    if not tech:
        tech = ["传感器采集与智能决策", "人机交互控制"]

    keywords = unique(
        _extract_free_keywords(raw_idea)
        + [k for name, keys in FUNCTION_RULES.items() for k in keys if k in text]
    )[:12]

    pain_points = []
    if contains_any(text, ["危险", "跌倒", "异常", "燃气", "火", "安全"]):
        pain_points.append("现有人工监护或被动告警存在发现不及时的问题")
    if contains_any(text, ["老人", "儿童", "残障", "宠物"]):
        pain_points.append("目标用户需要更低门槛、更主动的持续服务")
    if contains_any(text, ["自动", "提醒", "清洁", "搬运", "照护"]):
        pain_points.append("重复性任务占用时间，缺少持续自动化执行")
    if not pain_points:
        pain_points.append("现有方案智能化、连续服务和闭环执行能力不足")

    return {
        "raw_idea": raw_idea.strip(),
        "target_groups": unique(targets),
        "scenarios": unique(scenes),
        "core_functions": unique(functions),
        "tech_modules": unique(tech),
        "keywords": unique(keywords) or unique(functions)[:6],
        "pain_points": unique(pain_points),
        "design_goal": f"围绕{scenes[0]}中的真实需求，构建可感知、可判断、可执行、可反馈的机器人服务闭环。",
    }


def extract_idea_fields(raw_idea: str) -> Dict[str, Any]:
    local = local_extract_idea_fields(raw_idea)

    ai = call_llm_json(
        """你是大学生智能机器人创意竞赛方案分析专家。请只返回JSON，不要Markdown。
字段必须包含：target_groups(list), scenarios(list), core_functions(list), tech_modules(list),
keywords(list), pain_points(list), design_goal(str)。要求忠实于用户原始创意，不得强行改成老人陪护。""",
        f"用户创意：{raw_idea}\n请提取竞赛设计字段。",
    )

    if not ai:
        return local

    merged = dict(local)
    for key in ["target_groups", "scenarios", "core_functions", "tech_modules", "keywords", "pain_points"]:
        value = ai.get(key)
        if isinstance(value, list) and value:
            merged[key] = unique([str(x).strip() for x in value if str(x).strip()])[:10]
    if isinstance(ai.get("design_goal"), str) and ai["design_goal"].strip():
        merged["design_goal"] = ai["design_goal"].strip()
    merged["raw_idea"] = raw_idea.strip()
    return merged


# ============================================================
# 知识库接口：未来老师资料直接接这里
# ============================================================
def search_competition_knowledge(fields: Dict[str, Any], top_k: int = 5) -> Dict[str, Any]:
    query = " ".join(
        fields.get("keywords", [])
        + fields.get("core_functions", [])
        + fields.get("scenarios", [])
    )

    if KNOWLEDGE_API_URL:
        try:
            body = json.dumps({"query": query, "top_k": top_k}, ensure_ascii=False).encode("utf-8")
            headers = {"Content-Type": "application/json"}
            if KNOWLEDGE_API_KEY:
                headers["Authorization"] = f"Bearer {KNOWLEDGE_API_KEY}"
            req = urllib.request.Request(KNOWLEDGE_API_URL, data=body, headers=headers, method="POST")
            with urllib.request.urlopen(req, timeout=20) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            cases = data.get("cases") or data.get("results") or []
            return {
                "knowledge_base_used": True,
                "similar_award_cases": cases[:top_k],
                "knowledge_sources": data.get("sources", []),
                "source_type": "external_knowledge_base",
            }
        except Exception as exc:
            print("KNOWLEDGE_API_ERROR:", repr(exc), flush=True)

    # 无老师资料时，样例库仅作占位，不冒充真实获奖数据
    scored = []
    kws = set(fields.get("keywords", []))
    for case in REFERENCE_CASES:
        case_kws = set(case["keywords"])
        union = kws | case_kws
        sim = len(kws & case_kws) / len(union) if union else 0
        scored.append({**case, "similarity": round(sim, 3), "is_sample": True})
    scored.sort(key=lambda x: x["similarity"], reverse=True)
    return {
        "knowledge_base_used": False,
        "similar_award_cases": scored[:top_k],
        "knowledge_sources": [],
        "source_type": "sample_case_library",
    }


# ============================================================
# 动态候选题目
# ============================================================
def score_candidate(
    fields: Dict[str, Any],
    candidate: Dict[str, Any],
    knowledge: Dict[str, Any],
    candidate_index: int,
) -> Dict[str, Any]:
    function_count = len(fields.get("core_functions", []))
    tech_count = len(fields.get("tech_modules", []))
    keyword_count = len(fields.get("keywords", []))

    similar_cases = knowledge.get("similar_award_cases", [])
    highest_similarity = 0.0
    similar_case = {}
    if similar_cases:
        highest_similarity = float(similar_cases[0].get("similarity", 0) or 0)
        similar_case = similar_cases[0]

    # 不是实际获奖概率，是可解释竞争力评分
    innovation = 72 + min(12, function_count * 2.4) + candidate_index * 1.8
    scientific = 70 + min(15, tech_count * 2.8)
    application = 72 + min(14, len(fields.get("scenarios", [])) * 4) + min(5, len(fields.get("target_groups", [])) * 2)
    expression = 72 + min(12, keyword_count * 1.2)

    if "闭环" in candidate.get("positioning", ""):
        innovation += 3
    if candidate.get("differentiator"):
        innovation += 2

    similarity_penalty = highest_similarity * (12 if knowledge.get("knowledge_base_used") else 6)
    feasibility_penalty = 0
    if function_count >= 7 and tech_count <= 2:
        feasibility_penalty = 8
    elif function_count >= 6 and tech_count <= 2:
        feasibility_penalty = 5

    base = innovation * 0.35 + scientific * 0.30 + application * 0.25 + expression * 0.10
    competitiveness = clamp(base - similarity_penalty - feasibility_penalty)

    return {
        "innovation_score": round(clamp(innovation), 1),
        "scientific_score": round(clamp(scientific), 1),
        "application_score": round(clamp(application), 1),
        "expression_score": round(clamp(expression), 1),
        "base_total": round(base, 1),
        "highest_similarity": round(highest_similarity, 3),
        "similar_case": similar_case,
        "similarity_penalty": round(similarity_penalty, 1),
        "feasibility_penalty": round(feasibility_penalty, 1),
        "competitiveness_score": round(competitiveness, 1),
        "score_note": "可解释竞赛竞争力评分，不等同于真实获奖概率。",
    }


def _project_theme(fields: Dict[str, Any]) -> Tuple[str, str, str]:
    funcs = fields.get("core_functions", [])
    scene = fields.get("scenarios", ["应用场景"])[0]
    target = fields.get("target_groups", ["目标用户"])[0]

    if any("宠物" in f for f in funcs) or "宠物" in scene:
        return "宠护", "宠物照护", "智能照护"
    if any("植物" in f for f in funcs) or "园艺" in scene:
        return "植养", "植物养护", "智能园艺"
    if any("危险" in f for f in funcs) or any(x in fields.get("raw_idea", "") for x in ["燃气", "关火", "高温", "火灾"]):
        return "安防", "风险预警", "主动安全"
    if any("学习" in f for f in funcs):
        return "学伴", "学习辅助", "智能教育"
    if any("清洁" in f for f in funcs):
        return "净居", "家务服务", "自主清洁"
    if any("水质" in f or "环境" in f for f in funcs):
        return "环巡", "环境监测", "智能巡检"
    if any("机械" in f for f in funcs):
        return "智作", "任务执行", "自主作业"
    return "智服", f"{target}服务", f"{scene}智能服务"


def local_generate_candidates(fields: Dict[str, Any], knowledge: Dict[str, Any]) -> List[Dict[str, Any]]:
    seed = stable_seed(fields.get("raw_idea", ""))
    rng = random.Random(seed)
    prefix, domain, value = _project_theme(fields)
    target = fields.get("target_groups", ["目标用户"])[0]
    scene = fields.get("scenarios", ["实际应用场景"])[0]
    funcs = fields.get("core_functions", ["智能感知", "自主服务"])
    tech = fields.get("tech_modules", ["智能感知与决策"])
    function_text = "、".join(funcs[:3])
    tech_text = "、".join(tech[:3])

    name_stems = [
        ("星", "智"), ("卫", "慧"), ("伴", "灵"), ("盾", "安"),
        ("擎", "智"), ("航", "慧"), ("联", "云"), ("芯", "睿"),
    ]
    rng.shuffle(name_stems)

    names = [
        f"{prefix}{name_stems[0][0]}",
        f"{name_stems[1][1]}{domain[:2]}",
        f"{prefix}{name_stems[2][1]}",
    ]

    candidates = [
        {
            "id": 1,
            "title": f"“{names[0]}”——面向{target}的{domain}机器人",
            "positioning": f"围绕{function_text}形成感知—决策—执行—反馈闭环，突出实际需求解决能力。",
            "core_tech": tech_text,
            "differentiator": "场景需求闭环",
        },
        {
            "id": 2,
            "title": f"“{names[1]}”——基于多模态感知的{scene}{value}机器人",
            "positioning": f"以{tech_text}为核心，强调多源信息融合、主动判断与可靠执行。",
            "core_tech": tech_text,
            "differentiator": "多模态感知与主动决策",
        },
        {
            "id": 3,
            "title": f"“{names[2]}”——融合智能联动与自主服务的{domain}机器人",
            "positioning": f"突出{function_text}之间的协同联动，强调展示性、可扩展性和工程落地。",
            "core_tech": tech_text,
            "differentiator": "模块协同与可扩展设计",
        },
    ]

    for idx, item in enumerate(candidates):
        item["scores"] = score_candidate(fields, item, knowledge, idx)
    return candidates


def generate_candidates(fields: Dict[str, Any], knowledge: Dict[str, Any]) -> List[Dict[str, Any]]:
    knowledge_summary = [
        {"title": c.get("title"), "similarity": c.get("similarity", 0)}
        for c in knowledge.get("similar_award_cases", [])[:3]
    ]

    ai = call_llm_json(
        """你是中国大学生智能机器人创意竞赛命题专家。只返回JSON。
必须根据本次用户创意动态生成3个明显不同的题目，禁止固定输出“智护星/安居守护者/慧联家护”。
格式：{"candidates":[{"id":1,"title":"...","positioning":"...","core_tech":"...","differentiator":"..."},...]}。
题目要像正式竞赛作品名，技术路线可实现，三题侧重点不同。""",
        json.dumps({"fields": fields, "knowledge": knowledge_summary}, ensure_ascii=False),
    )

    if ai and isinstance(ai.get("candidates"), list) and len(ai["candidates"]) >= 3:
        result = []
        for idx, item in enumerate(ai["candidates"][:3]):
            candidate = {
                "id": idx + 1,
                "title": str(item.get("title", "")).strip(),
                "positioning": str(item.get("positioning", "")).strip(),
                "core_tech": str(item.get("core_tech", "")).strip() or "、".join(fields.get("tech_modules", [])[:3]),
                "differentiator": str(item.get("differentiator", "")).strip(),
            }
            if not candidate["title"]:
                return local_generate_candidates(fields, knowledge)
            candidate["scores"] = score_candidate(fields, candidate, knowledge, idx)
            result.append(candidate)
        return result

    return local_generate_candidates(fields, knowledge)


def format_candidate_message(fields: Dict[str, Any], candidates: List[Dict[str, Any]], prefix: str, knowledge: Dict[str, Any]) -> str:
    lines = [
        prefix,
        "",
        "【创意理解】",
        "目标用户：" + "、".join(fields.get("target_groups", [])),
        "应用场景：" + "、".join(fields.get("scenarios", [])),
        "核心功能：" + "、".join(fields.get("core_functions", [])),
        "技术模块：" + "、".join(fields.get("tech_modules", [])),
        "核心痛点：" + "；".join(fields.get("pain_points", [])),
        "",
        "【3个动态候选题目】",
    ]

    for c in candidates:
        s = c["scores"]
        lines += [
            "",
            f"{c['id']}. {c['title']}",
            f"定位：{c['positioning']}",
            f"差异化：{c.get('differentiator', '')}",
            f"核心技术：{c['core_tech']}",
            f"评分：创新性{s['innovation_score']}，科学性{s['scientific_score']}，应用前景{s['application_score']}，设计表达{s['expression_score']}",
            f"竞赛竞争力：{s['competitiveness_score']} / 100",
        ]

    if knowledge.get("knowledge_base_used"):
        lines += ["", "已接入真实知识库，本轮评分已参考知识库检索结果。"]
    else:
        lines += ["", "说明：当前尚未接入老师提供的真实往届资料，相似案例仅用于流程验证；评分不是实际获奖概率。"]

    lines += ["", "请输入 1、2、3 选择题目；输入“重新生成”换一批；也可以继续补充你的创意要求。"]
    return "\n".join(lines)


# ============================================================
# 7页动态报告
# ============================================================
def _local_report(fields: Dict[str, Any], selected: Dict[str, Any], knowledge: Dict[str, Any]) -> Dict[str, Any]:
    """生成内容较完整的7页报告底稿。

    这一层必须在没有外部LLM时也能独立工作；若配置LLM，build_report_json会再做一次增强。
    """
    title = selected["title"]
    target_list = fields.get("target_groups", []) or ["目标用户"]
    scene_list = fields.get("scenarios", []) or ["实际应用场景"]
    funcs = fields.get("core_functions", []) or ["智能感知", "任务决策", "自主服务"]
    techs = fields.get("tech_modules", []) or ["传感器采集", "智能决策", "执行控制"]
    pain = fields.get("pain_points", []) or ["现有方案连续服务能力不足"]
    target = "、".join(target_list)
    scenes = "、".join(scene_list)
    func_text = "、".join(funcs[:6])
    tech_text = "、".join(techs[:6])
    raw_idea = fields.get("raw_idea", "")

    score = selected.get("scores", {}).get("competitiveness_score")
    positioning = selected.get("positioning", "")
    differentiator = selected.get("differentiator", "")

    pages = {
        "page_1": {
            "page_no": 1,
            "module": "封面",
            "title": title,
            "content": {
                "subtitle": "智能机器人创意竞赛 · 项目设计报告",
                "project_positioning": positioning,
                "design_statement": (
                    f"本项目面向{target}，聚焦{scenes}，以{func_text}为主要功能，"
                    f"通过{tech_text}构建从感知、判断到执行和反馈的完整机器人服务闭环。"
                ),
                "keywords": fields.get("keywords", [])[:8],
                "competition_score": score,
            },
        },
        "page_2": {
            "page_no": 2,
            "module": "设计背景",
            "title": "设计背景与用户需求",
            "content": {
                "background_paragraphs": [
                    f"随着家庭服务、智能感知与物联网技术的发展，{scenes}正在从单点智能设备逐步向能够主动感知、持续服务和协同执行的机器人系统演进。"
                    f"对于{target}而言，真正有价值的产品并不是简单增加一个提醒或遥控功能，而是能够在无人持续操作的情况下完成连续任务，并在异常时及时反馈。",
                    f"本项目来源于用户原始创意：{raw_idea}。围绕这一创意，方案将需求拆分为“感知对象—状态判断—任务决策—执行动作—结果反馈”五个环节，避免只做功能堆叠，强调完整服务闭环与竞赛展示效果。",
                    f"现有同类产品常见问题包括：功能相互割裂、对真实场景适应性不足、异常处理依赖人工以及长期运行稳定性不高。因此本项目将{func_text}进行协同设计，并把可量化测试指标纳入方案。",
                ],
                "pain_points": pain,
                "user_needs": [
                    f"低门槛：面向{target}，交互方式应直观，减少复杂设置。",
                    f"持续性：在{scenes}中支持长期、重复任务自动执行。",
                    "主动性：不仅等待用户指令，还能够识别关键事件并主动触发服务。",
                    "可靠性：关键操作应有状态确认、失败重试和人工兜底。",
                    "可扩展性：预留传感器、执行机构、联网设备和知识库接口。",
                ],
                "design_goals": [
                    fields.get("design_goal", ""),
                    "构建可感知、可判断、可执行、可反馈的机器人服务闭环。",
                    "兼顾竞赛展示性、工程可实现性、用户体验和后续扩展空间。",
                    "通过明确的功能指标和场景测试，让创新点可以被验证而不是停留在概念描述。",
                ],
            },
        },
        "page_3": {
            "page_no": 3,
            "module": "产品整体结构",
            "title": "机器人整体结构设计",
            "content": {
                "overview": (
                    f"机器人总体采用模块化分层架构。系统围绕{scenes}中的任务需求，把硬件感知、边缘计算、任务决策、执行机构、"
                    f"人机交互与远程服务统一到同一控制链路中。核心模块包括{func_text}。"
                ),
                "system_layers": [
                    "感知层：通过摄像头、环境传感器、状态传感器或用户输入采集场景信息，并完成数据时间同步与基础过滤。",
                    "认知决策层：对感知结果进行融合，完成目标/状态识别、异常判断、任务优先级排序和执行策略生成。",
                    "执行层：根据任务类型驱动移动底盘、机械执行机构、语音提示、喂食/清洁/开关等功能模块完成动作。",
                    "交互与联网层：面向用户提供提示、确认、任务进度、异常消息和远程信息同步，并记录关键运行日志。",
                ],
                "core_modules": funcs,
                "mechanical_concept": (
                    f"本体结构围绕{scenes}进行模块化设计。主体可采用移动底盘或固定式底座，根据实际任务配置摄像头、传感器阵列、"
                    "主控计算单元、执行机构与交互终端。模块之间采用标准化供电与通信接口，便于竞赛阶段快速替换和迭代。"
                ),
                "design_principles": [
                    "重心与运动机构优先保证稳定性和安全性。",
                    "传感器布置兼顾视野、遮挡、维护和线缆走向。",
                    "高频使用模块采用快拆或抽拉结构，方便清洁与维护。",
                    "外观与结构突出项目主题，使评委能够快速理解功能分区。",
                ],
            },
        },
        "page_4": {
            "page_no": 4,
            "module": "软硬件功能设计",
            "title": "硬件与软件功能设计",
            "content": {
                "hardware": [
                    {"name": "感知单元", "description": f"围绕{func_text}配置视觉、距离、环境或状态传感器；采样频率与安装位置按场景需求确定。"},
                    {"name": "主控计算单元", "description": "负责数据预处理、状态融合、任务调度与本地决策；关键基础功能在网络异常时仍可工作。"},
                    {"name": "执行机构", "description": "根据任务配置移动底盘、机械臂、舵机、电机、泵阀、喂食/清洁等专用机构，并增加限位和过流保护。"},
                    {"name": "通信与供电", "description": "通过Wi-Fi/Bluetooth/蜂窝网络或局域网与移动端、云端或智能家居设备交互，并配置电量检测与低电量策略。"},
                ],
                "software": [
                    {"name": "感知服务", "description": f"完成{tech_text}相关数据采集、预处理与质量检测。"},
                    {"name": "状态识别", "description": "将单一传感器结果转化为可用于决策的状态标签，并对不确定结果保留置信度。"},
                    {"name": "任务调度", "description": f"根据{func_text}设置优先级、触发条件、互斥条件和失败重试策略。"},
                    {"name": "用户交互", "description": "提供状态提示、任务确认、异常提醒、远程查看与历史记录，减少用户操作负担。"},
                ],
                "interaction_logic": "用户/环境事件 → 多源感知 → 数据预处理 → 状态识别 → 风险与任务判断 → 执行动作 → 结果确认 → 记录与远程反馈",
                "functional_matrix": [
                    {"function": f, "trigger": "定时/事件/用户指令", "feedback": "状态提示 + 执行结果 + 异常记录"}
                    for f in funcs[:6]
                ],
            },
        },
        "page_5": {
            "page_no": 5,
            "module": "关键技术",
            "title": "关键技术与实现路线",
            "content": {
                "technical_summary": (
                    f"项目技术路线以{tech_text}为核心。实现时不追求复杂算法堆叠，而是把每项技术与具体任务、输入数据、判断逻辑、"
                    "执行动作和测试指标对应起来，保证方案能够从概念顺利过渡到样机。"
                ),
                "key_technologies": [
                    {
                        "name": t,
                        "implementation": (
                            f"围绕“{t}”设计输入数据、预处理方式、阈值/模型判断、输出动作和失败兜底。"
                            "竞赛阶段优先采用可复现、可解释、可快速迭代的实现方案，并记录关键日志用于调参。"
                        ),
                    }
                    for t in techs[:6]
                ],
                "technical_route": ["数据采集", "质量检查与预处理", "状态/目标识别", "多源信息融合", "任务决策", "执行控制", "反馈确认与日志记录"],
                "engineering_metrics": [
                    "识别类：准确率、召回率、误报率与复杂光照/遮挡条件下稳定性。",
                    "实时类：感知到动作的端到端响应延迟。",
                    "执行类：任务执行成功率、重复定位误差、机构卡滞/失败率。",
                    "通信类：异常消息送达率、断网后的本地降级能力。",
                    "可靠性：连续运行时长、异常恢复能力和关键模块故障隔离。",
                ],
                "risk_control": [
                    "涉及安全的动作设置人工确认、软硬件限位和紧急停止。",
                    "网络异常时保留本地基础任务，恢复连接后再同步日志。",
                    "模型输出不确定时进入保守策略，避免直接执行高风险动作。",
                    "涉及图像、语音与用户信息时设置最小化采集、权限管理和本地化存储策略。",
                ],
            },
        },
        "page_6": {
            "page_no": 6,
            "module": "项目创新点",
            "title": "项目创新点与差异化分析",
            "content": {
                "innovation_intro": (
                    f"本项目的创新重点不是单独增加一个新功能，而是针对{target}在{scenes}中的连续需求，"
                    f"把{func_text}组成可协同、可验证、可扩展的机器人服务系统。"
                ),
                "innovation_points": [
                    {"name": "需求驱动", "description": f"从{target}在{scenes}中的真实痛点出发，功能设计与场景任务一一对应。"},
                    {"name": "闭环服务", "description": "把感知、判断、执行、确认和记录连接起来，避免停留在提醒或遥控层面。"},
                    {"name": "多模块协同", "description": f"让{'、'.join(funcs[:4])}按任务优先级协同工作，而不是相互独立展示。"},
                    {"name": "工程化设计", "description": "采用模块化软硬件架构、失败重试、状态确认与安全兜底，提高样机可实现性。"},
                    {"name": "持续演进", "description": "预留知识库、智能家居/移动端接口和新传感器接入能力，便于后续迭代。"},
                ],
                "comparison": [
                    {"dimension": "需求理解", "common": "以功能清单为主", "ours": "从场景任务和用户痛点反推功能"},
                    {"dimension": "交互方式", "common": "用户主动下达指令", "ours": "主动感知 + 事件触发 + 用户确认"},
                    {"dimension": "系统结构", "common": "模块相互独立", "ours": "统一任务调度与状态闭环"},
                    {"dimension": "可靠性", "common": "重演示、轻异常处理", "ours": "加入失败重试、日志和安全兜底"},
                    {"dimension": "扩展能力", "common": "功能固定", "ours": "模块接口化并预留知识库与外部设备接口"},
                ],
                "differentiation_analysis": {
                    "knowledge_base_used": knowledge.get("knowledge_base_used", False),
                    "similar_award_cases": knowledge.get("similar_award_cases", [])[:3],
                    "current_boundary": "未接入老师提供的真实获奖资料时，不把样例相似度宣传为真实获奖证据。",
                    "selected_differentiator": differentiator,
                },
            },
        },
        "page_7": {
            "page_no": 7,
            "module": "行业应用前景",
            "title": "行业应用前景与落地路径",
            "content": {
                "prospect_paragraphs": [
                    f"从应用价值看，项目面向{target}，可首先在{scenes}完成小规模验证。与单一智能设备相比，机器人具备移动/执行、持续感知和多任务协同能力，"
                    "因此更适合承担具有连续性、重复性和异常处置要求的服务任务。",
                    "从产品化路径看，应优先完成核心任务闭环，再逐步增加高级功能。竞赛样机阶段重点证明功能可行和场景价值；后续可通过模块化版本、移动端服务、"
                    "智能家居或行业平台接入形成更完整的产品体系。",
                ],
                "application_scenarios": scene_list,
                "target_users": target_list,
                "deployment_paths": [
                    "阶段1：完成核心功能最小可行样机，验证主任务闭环。",
                    "阶段2：在目标场景开展连续测试，记录成功、失败和异常样本。",
                    "阶段3：根据测试结果优化结构、算法、交互和安全策略。",
                    "阶段4：形成模块化版本，适配不同家庭/场景和成本档位。",
                    "阶段5：接入移动端、智能家居、知识库或行业平台，构建持续服务能力。",
                ],
                "social_value": [
                    "降低重复性人工工作负担，提高用户时间利用效率。",
                    "提高异常发现和信息反馈速度，减少关键事件漏报。",
                    "通过持续服务与低门槛交互提升目标用户使用体验。",
                    "形成可扩展的数据与服务闭环，为后续算法迭代提供基础。",
                ],
                "future_iterations": [
                    "接入老师提供的真实往届获奖作品知识库，完善相似度与竞争力分析。",
                    "增加真实场景测试数据和可量化指标，形成更有说服力的竞赛证据。",
                    "优化机器人外观、结构与交互细节，提升作品完成度和展示效果。",
                    "补齐答辩PPT、演示视频、结构图纸和测试记录，形成完整参赛材料。",
                ],
            },
        },
    }

    return {
        "schema_version": "2.0",
        "project_title": title,
        "raw_idea": raw_idea,
        "selected_candidate": selected,
        "idea_fields": fields,
        "knowledge": {
            "knowledge_base_used": knowledge.get("knowledge_base_used", False),
            "similar_award_cases": knowledge.get("similar_award_cases", [])[:5],
            "similarity_score": (knowledge.get("similar_award_cases") or [{}])[0].get("similarity") if knowledge.get("similar_award_cases") else None,
            "competition_probability": None,
            "knowledge_sources": knowledge.get("knowledge_sources", []),
        },
        "pages": pages,
        "generated_at": now_iso(),
    }

def build_report_json(fields: Dict[str, Any], selected: Dict[str, Any], knowledge: Dict[str, Any]) -> Dict[str, Any]:
    base = _local_report(fields, selected, knowledge)

    ai = call_llm_json(
        """你是智能机器人竞赛设计报告专家。只返回JSON。
必须保持严格7页：page_1封面、page_2设计背景、page_3整体结构、page_4软硬件功能、page_5关键技术、page_6创新点、page_7应用前景。
内容必须完全围绕给定用户创意和已选题目，禁止套用老人陪护固定模板。每页内容要专业、具体、可落地。
返回格式：{"pages":{...}}，不要删除页码/module/title/content字段。""",
        json.dumps({"base_report": base, "fields": fields, "selected": selected}, ensure_ascii=False),
    )

    if ai and isinstance(ai.get("pages"), dict):
        # 只替换7页内容，保留元数据和知识库字段
        for key in [f"page_{i}" for i in range(1, 8)]:
            if isinstance(ai["pages"].get(key), dict):
                page = ai["pages"][key]
                page.setdefault("page_no", int(key.split("_")[1]))
                page.setdefault("module", base["pages"][key]["module"])
                page.setdefault("title", base["pages"][key]["title"])
                page.setdefault("content", base["pages"][key]["content"])
                base["pages"][key] = page
    return base


def format_report_summary(report_json: Dict[str, Any]) -> str:
    lines = [
        "7页结构化竞赛报告已经生成完成。",
        "",
        f"项目题目：{report_json.get('project_title', '')}",
        "",
        "报告固定结构：",
    ]
    for i in range(1, 8):
        page = report_json["pages"][f"page_{i}"]
        lines.append(f"{i}. {page.get('title', page.get('module', ''))}")
    lines += [
        "",
        "结构化数据已保存在 payload_json 中。",
        "输入“下载报告”即可生成图文并茂的 Word 文档。",
    ]
    if not report_json.get("knowledge", {}).get("knowledge_base_used"):
        lines += ["", "当前尚未接入老师提供的真实获奖作品知识库；接口已预留，后续直接接入即可。"]
    return "\n".join(lines)


# ============================================================
# 动态图片：项目相关概念图、结构图、流程图、创新图
# ============================================================
def _domain_hint(report: Dict[str, Any]) -> str:
    fields = report.get("idea_fields", {})
    raw = fields.get("raw_idea", "")
    text = normalize_text(raw + " " + " ".join(fields.get("core_functions", [])))
    if contains_any(text, ["宠物", "猫", "狗", "猫砂", "喂食"]):
        return "pet care robot in a modern home, cat or dog nearby"
    if contains_any(text, ["老人", "养老", "跌倒", "用药"]):
        return "home elder-care companion robot assisting an older adult"
    if contains_any(text, ["厨房", "燃气", "关火", "烹饪"]):
        return "home kitchen safety robot near a stove and cooking area"
    if contains_any(text, ["植物", "浇水", "多肉", "园艺"]):
        return "home plant-care robot tending potted plants on a balcony"
    if contains_any(text, ["儿童", "学习", "教育", "作业"]):
        return "family educational companion robot helping a child study"
    if contains_any(text, ["清洁", "扫地", "拖地", "收纳"]):
        return "home service robot doing cleaning and organizing"
    if contains_any(text, ["农业", "果园", "温室", "采摘"]):
        return "agricultural service robot working in a greenhouse or orchard"
    if contains_any(text, ["仓库", "物流", "搬运"]):
        return "autonomous warehouse service robot moving goods"
    return "smart service robot operating in its intended real-world environment"


def _visual_prompt(report: Dict[str, Any], kind: str) -> str:
    fields = report.get("idea_fields", {})
    title = report.get("project_title", "smart service robot")
    funcs = ", ".join(fields.get("core_functions", [])[:5])
    techs = ", ".join(fields.get("tech_modules", [])[:4])
    domain = _domain_hint(report)
    if kind == "product":
        return (
            f"Professional industrial design concept render for a university robotics competition project titled {title}. "
            f"Show a realistic, buildable smart robot, {domain}. Core functions: {funcs}. Technologies: {techs}. "
            "Clean modern product design, white studio plus subtle real environment, no text, no logo, realistic materials, engineering prototype quality, 16:9."
        )
    if kind == "scene":
        return (
            f"Realistic application scene for a university robotics competition project titled {title}: {domain}. "
            f"Clearly show the robot performing useful tasks related to {funcs}. Human-safe, practical, believable home/service environment, no text, no logo, documentary product photography style, 16:9."
        )
    return (
        f"Realistic close-up functional detail of a smart robot project titled {title}, showing sensors and actuators for {funcs}; "
        f"technologies {techs}; engineering prototype, no text, no logo, 16:9."
    )


def _download_binary(url: str, dest: Path, headers: Optional[Dict[str, str]] = None, timeout: int = 90) -> bool:
    try:
        req = urllib.request.Request(url, headers=headers or {}, method="GET")
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = resp.read()
        if len(data) < 5000:
            return False
        dest.write_bytes(data)
        return True
    except Exception as exc:
        print("IMAGE_DOWNLOAD_ERROR:", repr(exc), flush=True)
        return False


def _pollinations_image(prompt: str, dest: Path) -> bool:
    if not POLLINATIONS_API_KEY:
        return False
    encoded = urllib.parse.quote(prompt[:1800], safe="")
    url = (
        f"https://gen.pollinations.ai/image/{encoded}"
        f"?model={urllib.parse.quote(POLLINATIONS_IMAGE_MODEL)}&width=1280&height=768&nologo=true"
    )
    return _download_binary(
        url,
        dest,
        headers={"Authorization": f"Bearer {POLLINATIONS_API_KEY}"},
        timeout=IMAGE_TIMEOUT,
    )


def _commons_search_terms(report: Dict[str, Any], kind: str) -> str:
    hint = _domain_hint(report)
    if "pet care" in hint:
        return "pet care robot cat feeder robot" if kind == "product" else "cat home smart pet feeder"
    if "elder-care" in hint:
        return "service robot elder care" if kind == "product" else "elderly home service robot"
    if "kitchen" in hint:
        return "kitchen service robot" if kind == "product" else "smart kitchen robot"
    if "plant-care" in hint:
        return "agricultural robot plant care" if kind == "product" else "smart irrigation potted plants"
    if "educational" in hint:
        return "educational robot" if kind == "product" else "robot child education"
    if "cleaning" in hint:
        return "home cleaning robot" if kind == "product" else "robot vacuum home"
    if "agricultural" in hint:
        return "agricultural robot" if kind == "product" else "farm robot greenhouse"
    if "warehouse" in hint:
        return "warehouse robot" if kind == "product" else "autonomous mobile robot warehouse"
    return "service robot prototype" if kind == "product" else "service robot home"


def _wikimedia_commons_image(report: Dict[str, Any], kind: str, dest: Path) -> bool:
    """无需API Key的真实图片兜底。仅从Wikimedia Commons开放资源检索。"""
    query = _commons_search_terms(report, kind)
    params = {
        "action": "query",
        "generator": "search",
        "gsrsearch": query,
        "gsrnamespace": "6",
        "gsrlimit": "8",
        "prop": "imageinfo",
        "iiprop": "url|mime",
        "format": "json",
        "origin": "*",
    }
    url = "https://commons.wikimedia.org/w/api.php?" + urllib.parse.urlencode(params)
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "RobotCompetitionAssistant/2.0"})
        with urllib.request.urlopen(req, timeout=20) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        pages = list((data.get("query", {}).get("pages", {}) or {}).values())
        for page in pages:
            infos = page.get("imageinfo") or []
            if not infos:
                continue
            info = infos[0]
            mime = str(info.get("mime", "")).lower()
            image_url = info.get("url")
            if image_url and mime in {"image/jpeg", "image/png", "image/webp"}:
                if _download_binary(image_url, dest, headers={"User-Agent": "RobotCompetitionAssistant/2.0"}, timeout=30):
                    return True
    except Exception as exc:
        print("WIKIMEDIA_IMAGE_ERROR:", repr(exc), flush=True)
    return False


def _generic_image_api(prompt: str, dest: Path) -> bool:
    """兼容用户未来自建/其他图片服务。
    约定POST JSON {prompt,width,height}，返回图片二进制，或JSON中的url字段。
    """
    if not IMAGE_API_URL:
        return False
    body = json.dumps({"prompt": prompt, "width": 1280, "height": 768}, ensure_ascii=False).encode("utf-8")
    headers = {"Content-Type": "application/json"}
    if IMAGE_API_KEY:
        headers["Authorization"] = f"Bearer {IMAGE_API_KEY}"
    try:
        req = urllib.request.Request(IMAGE_API_URL, data=body, headers=headers, method="POST")
        with urllib.request.urlopen(req, timeout=IMAGE_TIMEOUT) as resp:
            ctype = (resp.headers.get("Content-Type") or "").lower()
            data = resp.read()
        if "image" in ctype and len(data) > 5000:
            dest.write_bytes(data)
            return True
        raw = json.loads(data.decode("utf-8"))
        image_url = raw.get("url") or raw.get("image_url") or raw.get("data", {}).get("url")
        if image_url:
            return _download_binary(image_url, dest, timeout=IMAGE_TIMEOUT)
    except Exception as exc:
        print("GENERIC_IMAGE_API_ERROR:", repr(exc), flush=True)
    return False


def _fallback_robot_illustration(report: Dict[str, Any], dest: Path, scene: bool = False) -> None:
    """离线兜底插画：不写中文文字，避免Render缺少中文字体出现方框。"""
    W, H = 1280, 768
    img = Image.new("RGB", (W, H), (246, 248, 251))
    d = ImageDraw.Draw(img)
    # environment
    d.rectangle((0, 560, W, H), fill=(226, 232, 239))
    d.rectangle((60, 90, 1220, 590), outline=(185, 197, 210), width=4)
    # robot body
    cx, cy = 650, 385
    d.rounded_rectangle((cx-170, cy-150, cx+170, cy+135), radius=55, fill=(235, 240, 248), outline=(68, 91, 126), width=7)
    d.rounded_rectangle((cx-115, cy-245, cx+115, cy-130), radius=45, fill=(248, 250, 253), outline=(68, 91, 126), width=7)
    d.ellipse((cx-62, cy-210, cx-22, cy-170), fill=(70, 116, 176))
    d.ellipse((cx+22, cy-210, cx+62, cy-170), fill=(70, 116, 176))
    d.rounded_rectangle((cx-95, cy-70, cx+95, cy+20), radius=22, fill=(214, 225, 239), outline=(110, 130, 160), width=4)
    d.ellipse((cx-160, cy+110, cx-80, cy+190), fill=(66, 77, 90))
    d.ellipse((cx+80, cy+110, cx+160, cy+190), fill=(66, 77, 90))
    # sensor rays
    for dx in (-260, -210, 210, 260):
        d.line((cx + (110 if dx>0 else -110), cy-180, cx+dx, cy-250), fill=(124, 151, 185), width=3)
    # domain object cues
    hint = _domain_hint(report)
    if "pet care" in hint:
        # pet bowl + cat silhouette
        d.ellipse((200, 560, 360, 620), fill=(174, 191, 214), outline=(90, 110, 140), width=3)
        d.polygon([(980,520),(1030,470),(1080,520)], fill=(115, 130, 150))
        d.ellipse((940,500,1120,650), fill=(130,145,165))
        d.polygon([(960,505),(980,455),(1015,505)], fill=(130,145,165))
        d.polygon([(1045,505),(1080,455),(1100,510)], fill=(130,145,165))
    elif "plant-care" in hint:
        d.rectangle((160,540,340,670), fill=(181,149,118))
        d.line((250,540,250,420), fill=(89,122,82), width=12)
        d.ellipse((180,430,250,510), fill=(116,157,105)); d.ellipse((250,410,340,500), fill=(116,157,105))
    elif "kitchen" in hint:
        d.rectangle((100,420,360,620), fill=(203,210,218), outline=(110,120,130), width=4)
        d.ellipse((150,450,220,520), outline=(160,80,60), width=5)
    elif "educational" in hint:
        d.rectangle((120,480,380,620), fill=(196,206,219), outline=(110,120,140), width=4)
        d.rectangle((160,420,340,500), fill=(236,240,246), outline=(100,120,150), width=3)
    if scene:
        # add phone/status panel on right to suggest remote interaction
        d.rounded_rectangle((1040,140,1180,390), radius=25, fill=(252,253,255), outline=(82,105,138), width=5)
        d.ellipse((1090,170,1130,210), fill=(95,137,189))
        for y in (245,285,325):
            d.rounded_rectangle((1070,y,1150,y+16), radius=8, fill=(170,188,210))
    img.save(dest, format="PNG")


def generate_project_images(report: Dict[str, Any], session_key: str) -> Dict[str, str]:
    digest = hashlib.sha1(f"{session_key}-{report.get('project_title','')}".encode("utf-8")).hexdigest()[:10]
    result: Dict[str, str] = {}
    for kind in ["product", "scene"]:
        path = EXPORT_DIR / f"visual_{kind}_{digest}.png"
        prompt = _visual_prompt(report, kind)
        ok = (
            _generic_image_api(prompt, path)
            or _pollinations_image(prompt, path)
            or _wikimedia_commons_image(report, kind, path)
        )
        if not ok:
            _fallback_robot_illustration(report, path, scene=(kind == "scene"))
        result[kind] = str(path)
    return result

# ============================================================
# Word 美化与图文排版
# ============================================================
def _set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def _set_run_font(run, name: str = "Microsoft YaHei", size: float = 10.5, bold: bool = False) -> None:
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold


def _add_title(doc: Document, text: str, size: float = 22) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    _set_run_font(r, size=size, bold=True)


def _add_heading(doc: Document, text: str, level: int = 1) -> None:
    p = doc.add_paragraph()
    r = p.add_run(text)
    _set_run_font(r, size=16 if level == 1 else 13, bold=True)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(3)


def _add_paragraph(doc: Document, text: Any, bold_prefix: Optional[str] = None, first_line: bool = True) -> None:
    if text is None or str(text).strip() == "":
        return
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.28
    p.paragraph_format.space_after = Pt(3)
    if first_line:
        p.paragraph_format.first_line_indent = Cm(0.74)
    if bold_prefix and str(text).startswith(bold_prefix):
        r1 = p.add_run(bold_prefix)
        _set_run_font(r1, bold=True)
        r2 = p.add_run(str(text)[len(bold_prefix):])
        _set_run_font(r2)
    else:
        r = p.add_run(str(text))
        _set_run_font(r)


def _add_paragraphs(doc: Document, items: Any, limit: int = 10) -> None:
    if isinstance(items, list):
        for item in items[:limit]:
            _add_paragraph(doc, item)
    elif items:
        _add_paragraph(doc, items)


def _add_bullets(doc: Document, items: Any, limit: int = 10) -> None:
    if not isinstance(items, list):
        if items:
            _add_paragraph(doc, items)
        return
    for item in items[:limit]:
        p = doc.add_paragraph(style=None)
        p.paragraph_format.left_indent = Cm(0.55)
        p.paragraph_format.first_line_indent = Cm(-0.3)
        p.paragraph_format.line_spacing = 1.2
        p.paragraph_format.space_after = Pt(2)
        if isinstance(item, dict):
            name = item.get("name") or item.get("title") or item.get("dimension") or item.get("function") or ""
            desc = item.get("description") or item.get("implementation") or item.get("ours") or item.get("feedback") or ""
            value = f"• {name}：{desc}" if desc else f"• {name}"
        else:
            value = f"• {item}"
        r = p.add_run(value)
        _set_run_font(r, size=10.2)


def _add_image(doc: Document, path: Optional[str], caption: str, width_cm: float = 15.5) -> None:
    if not path or not Path(path).exists():
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run()
    try:
        r.add_picture(path, width=Cm(width_cm))
    except Exception as exc:
        print("WORD_IMAGE_INSERT_ERROR:", repr(exc), flush=True)
        return
    c = doc.add_paragraph()
    c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c.paragraph_format.space_after = Pt(3)
    rr = c.add_run(caption)
    _set_run_font(rr, size=9)


def _style_table(table, header_fill: str = "DCE6F1", body_fill: Optional[str] = None) -> None:
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for ridx, row in enumerate(table.rows):
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if ridx == 0:
                _set_cell_shading(cell, header_fill)
            elif body_fill:
                _set_cell_shading(cell, body_fill)
            for p in cell.paragraphs:
                p.paragraph_format.space_after = Pt(0)
                for run in p.runs:
                    _set_run_font(run, size=9.5, bold=(ridx == 0))


def _add_architecture_table(doc: Document, layers: List[str]) -> None:
    _add_heading(doc, "系统总体架构", 2)
    table = doc.add_table(rows=1, cols=2)
    table.rows[0].cells[0].text = "层级"
    table.rows[0].cells[1].text = "主要职责"
    for item in layers[:6]:
        if "：" in str(item):
            a, b = str(item).split("：", 1)
        else:
            a, b = "模块", str(item)
        cells = table.add_row().cells
        cells[0].text = a
        cells[1].text = b
    _style_table(table)


def _add_flow_table(doc: Document, steps: List[str]) -> None:
    _add_heading(doc, "技术流程", 2)
    steps = [str(s) for s in steps[:8] if s]
    if not steps:
        return
    table = doc.add_table(rows=2, cols=len(steps))
    for i, step in enumerate(steps):
        table.rows[0].cells[i].text = f"步骤{i+1}"
        table.rows[1].cells[i].text = step
    _style_table(table)


def _add_comparison_table(doc: Document, rows: List[Dict[str, Any]]) -> None:
    _add_heading(doc, "差异化对比", 2)
    table = doc.add_table(rows=1, cols=3)
    table.rows[0].cells[0].text = "比较维度"
    table.rows[0].cells[1].text = "常规方案"
    table.rows[0].cells[2].text = "本项目"
    for item in rows[:8]:
        cells = table.add_row().cells
        cells[0].text = str(item.get("dimension", ""))
        cells[1].text = str(item.get("common", ""))
        cells[2].text = str(item.get("ours", ""))
    _style_table(table)


def _add_function_matrix(doc: Document, rows: List[Dict[str, Any]]) -> None:
    _add_heading(doc, "核心功能矩阵", 2)
    table = doc.add_table(rows=1, cols=3)
    table.rows[0].cells[0].text = "功能"
    table.rows[0].cells[1].text = "触发方式"
    table.rows[0].cells[2].text = "反馈方式"
    for item in rows[:7]:
        cells = table.add_row().cells
        cells[0].text = str(item.get("function", ""))
        cells[1].text = str(item.get("trigger", ""))
        cells[2].text = str(item.get("feedback", ""))
    _style_table(table)


def _prepare_document() -> Document:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(1.45)
    section.bottom_margin = Cm(1.35)
    section.left_margin = Cm(1.55)
    section.right_margin = Cm(1.55)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Microsoft YaHei"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    normal.font.size = Pt(10.3)
    return doc


def _render_cover(doc: Document, report: Dict[str, Any], images: Dict[str, str]) -> None:
    doc.add_paragraph("\n")
    _add_title(doc, report.get("project_title", "智能机器人竞赛项目"), 22)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("智能机器人创意竞赛 · 项目设计报告")
    _set_run_font(r, size=13)
    _add_image(doc, images.get("product"), "图1  项目产品概念效果图", 15.8)

    content = report.get("pages", {}).get("page_1", {}).get("content", {})
    _add_paragraph(doc, content.get("design_statement", ""), first_line=False)

    selected = report.get("selected_candidate", {})
    table = doc.add_table(rows=4, cols=2)
    pairs = [
        ("项目定位", selected.get("positioning", "")),
        ("核心技术", selected.get("core_tech", "")),
        ("差异化", selected.get("differentiator", "")),
        ("竞赛竞争力", f"{selected.get('scores', {}).get('competitiveness_score', '')} / 100"),
    ]
    for row, (k, v) in zip(table.rows, pairs):
        row.cells[0].text = k
        row.cells[1].text = str(v)
    _style_table(table, header_fill="E8EEF7")


def _render_page(doc: Document, report: Dict[str, Any], page_no: int, images: Dict[str, str]) -> None:
    page = report["pages"][f"page_{page_no}"]
    content = page.get("content", {})
    _add_heading(doc, f"第{page_no}页  {page.get('title', '')}", 1)

    if page_no == 2:
        _add_image(doc, images.get("scene"), "图2  项目典型应用场景示意图", 15.2)
        _add_heading(doc, "设计背景", 2)
        _add_paragraphs(doc, content.get("background_paragraphs") or content.get("background", []), 4)
        _add_heading(doc, "用户痛点", 2)
        _add_bullets(doc, content.get("pain_points", []), 6)
        _add_heading(doc, "用户需求", 2)
        _add_bullets(doc, content.get("user_needs", []), 6)
        _add_heading(doc, "设计目标", 2)
        _add_bullets(doc, content.get("design_goals", []), 6)

    elif page_no == 3:
        _add_paragraph(doc, content.get("overview", ""))
        _add_architecture_table(doc, content.get("system_layers", []))
        _add_heading(doc, "核心模块", 2)
        _add_bullets(doc, content.get("core_modules", []), 7)
        _add_heading(doc, "本体结构思路", 2)
        _add_paragraph(doc, content.get("mechanical_concept", ""))
        _add_heading(doc, "结构设计原则", 2)
        _add_bullets(doc, content.get("design_principles", []), 6)

    elif page_no == 4:
        _add_heading(doc, "硬件设计", 2)
        _add_bullets(doc, content.get("hardware", []), 8)
        _add_heading(doc, "软件功能", 2)
        _add_bullets(doc, content.get("software", []), 8)
        _add_function_matrix(doc, content.get("functional_matrix", []))
        _add_heading(doc, "交互逻辑", 2)
        _add_paragraph(doc, content.get("interaction_logic", ""), first_line=False)

    elif page_no == 5:
        _add_paragraph(doc, content.get("technical_summary", ""))
        _add_flow_table(doc, content.get("technical_route", []))
        _add_heading(doc, "关键技术", 2)
        _add_bullets(doc, content.get("key_technologies", []), 6)
        _add_heading(doc, "量化指标", 2)
        _add_bullets(doc, content.get("engineering_metrics", []), 6)
        _add_heading(doc, "风险控制", 2)
        _add_bullets(doc, content.get("risk_control", []), 6)

    elif page_no == 6:
        _add_paragraph(doc, content.get("innovation_intro", ""))
        _add_heading(doc, "创新点", 2)
        _add_bullets(doc, content.get("innovation_points", []), 7)
        _add_comparison_table(doc, content.get("comparison", []))
        diff = content.get("differentiation_analysis", {})
        note = diff.get("current_boundary") if isinstance(diff, dict) else None
        if note:
            _add_heading(doc, "知识库说明", 2)
            _add_paragraph(doc, note)

    elif page_no == 7:
        _add_paragraphs(doc, content.get("prospect_paragraphs", []), 4)
        _add_heading(doc, "典型应用场景", 2)
        _add_bullets(doc, content.get("application_scenarios", []), 6)
        _add_heading(doc, "落地路径", 2)
        _add_bullets(doc, content.get("deployment_paths", []), 7)
        _add_heading(doc, "社会价值", 2)
        _add_bullets(doc, content.get("social_value", []), 6)
        _add_heading(doc, "后续迭代", 2)
        _add_bullets(doc, content.get("future_iterations", []), 6)


def export_report_to_word(report_json: Dict[str, Any], session_key: str) -> Dict[str, str]:
    images = generate_project_images(report_json, session_key)
    doc = _prepare_document()
    _render_cover(doc, report_json, images)

    for page_no in range(2, 8):
        doc.add_page_break()
        _render_page(doc, report_json, page_no, images)

    digest = hashlib.sha1(
        f"{session_key}-{report_json.get('project_title','')}-{now_iso()}".encode("utf-8")
    ).hexdigest()[:12]

    # 下载URL永远使用ASCII文件名，避免学习通/浏览器截断中文链接。
    filename = f"robot_competition_report_{digest}.docx"
    path = EXPORT_DIR / filename
    doc.save(path)

    return {
        "filename": filename,
        "file_path": str(path),
        "download_url": f"{PUBLIC_BASE_URL}/api/v1/robot-competition/download/{filename}",
        "visual_mode": "ai_generated" if (POLLINATIONS_API_KEY or IMAGE_API_URL) else "offline_fallback",
    }

# ============================================================
# 意图识别与会话流程
# ============================================================
def detect_page_number(message: str) -> Optional[int]:
    text = normalize_text(message)
    for pattern in [r"第([1-7])页", r"page([1-7])", r"页([1-7])"]:
        m = re.search(pattern, text, re.I)
        if m:
            return int(m.group(1))
    return None


def detect_intent(message: str, session: Dict[str, Any]) -> str:
    text = normalize_text(message)

    if contains_any(text, ["查看报告结构", "查看结构化数据", "查看json"]):
        return "view_report_json"
    if contains_any(text, ["修改第", "调整第", "改第", "补充第"]) and detect_page_number(message):
        return "modify_page_request"
    if contains_any(text, ["重新生成", "重生成", "再生成", "换一批", "不满意"]):
        return "regenerate_titles"
    if re.fullmatch(r"[123]", text) or contains_any(text, ["选1", "选2", "选3", "第一个", "第二个", "第三个"]):
        return "select_title"
    if contains_any(text, ["生成报告", "写报告", "生成方案书"]):
        return "generate_report"
    if contains_any(text, ["下载", "导出", "生成word", "word报告", "word文档"]) and contains_any(text, ["报告", "word", "文档", "方案书"]):
        return "word_export_pending"
    if "raw_idea" not in session:
        return "create_project"
    return "supplement_idea"


def parse_selection(message: str) -> Optional[int]:
    text = normalize_text(message)
    if text == "1" or "选1" in text or "第一个" in text:
        return 1
    if text == "2" or "选2" in text or "第二个" in text:
        return 2
    if text == "3" or "选3" in text or "第三个" in text:
        return 3
    return None


def response(**kwargs: Any) -> ChatResponse:
    kwargs.setdefault("payload_json", "")
    kwargs.setdefault("download_url", "")
    kwargs.setdefault("data", {})
    kwargs.setdefault("files", [])
    kwargs.setdefault("suggested_actions", [])
    kwargs.setdefault("updated_at", now_iso())
    return ChatResponse(**kwargs)


def handle_chat(req: ChatRequest) -> ChatResponse:
    session_key = build_session_key(req)
    session = load_session(session_key)
    user_message = clean_message(req.message)
    intent = detect_intent(user_message, session)

    print("DEBUG_MESSAGE:", user_message, "INTENT:", intent, flush=True)
    print("DEBUG_SESSION:", session_key, "SESSION_KEYS:", list(session.keys()), flush=True)

    if intent in {"create_project", "supplement_idea", "regenerate_titles"}:
        if intent == "create_project":
            raw_idea = user_message
            prefix = "已收到你的机器人创意，正在根据本次输入动态生成3个候选题目。"
        elif intent == "supplement_idea":
            raw_idea = (session.get("raw_idea", "") + "\n补充要求：" + user_message).strip()
            prefix = "已把补充要求合并到原始创意，下面重新生成候选题目。"
        else:
            raw_idea = session.get("raw_idea", user_message)
            # 增加扰动，保证“重新生成”能产生不同命名倾向
            raw_idea = raw_idea + f"\n重新生成批次：{now_iso()}"
            prefix = "已根据当前创意重新生成一批候选题目。"

        fields = extract_idea_fields(raw_idea)
        knowledge = search_competition_knowledge(fields)
        candidates = generate_candidates(fields, knowledge)

        session.update({
            "raw_idea": raw_idea,
            "fields": fields,
            "knowledge": knowledge,
            "candidates": candidates,
            "selected_title": None,
            "report_json": None,
            "stage": "candidates_ready",
        })
        save_session(session_key, session)

        payload = {"fields": fields, "candidates": candidates, "knowledge": knowledge}
        return response(
            success=True,
            stage="candidates_ready",
            intent=intent,
            message=format_candidate_message(fields, candidates, prefix, knowledge),
            payload_json=json.dumps(payload, ensure_ascii=False, indent=2),
            data=payload,
            suggested_actions=["输入1", "输入2", "输入3", "重新生成", "补充要求"],
        )

    if intent == "select_title":
        selected_id = parse_selection(user_message)
        candidates = session.get("candidates", [])
        selected = next((c for c in candidates if c.get("id") == selected_id), None)
        if not selected:
            return response(
                success=False,
                stage="selection_failed",
                intent=intent,
                message="没有找到对应候选题目。请先输入一个完整机器人创意，然后选择1、2或3。",
                suggested_actions=["输入1", "输入2", "输入3"],
            )

        session["selected_title"] = selected
        session["stage"] = "title_confirmed"
        save_session(session_key, session)
        payload = {"selected_title": selected, "current_stage": "title_confirmed"}
        return response(
            success=True,
            stage="title_confirmed",
            intent=intent,
            message=(
                f"已确认最终竞赛题目：\n\n{selected['title']}\n\n"
                f"竞赛竞争力评分：{selected['scores']['competitiveness_score']} / 100\n"
                f"核心技术：{selected['core_tech']}\n\n"
                "请输入“生成报告”，系统将围绕本项目生成固定7页结构、动态内容的竞赛报告。"
            ),
            payload_json=json.dumps(payload, ensure_ascii=False, indent=2),
            data=payload,
            suggested_actions=["生成报告", "重新生成题目"],
        )

    if intent == "generate_report":
        selected = session.get("selected_title")
        fields = session.get("fields")
        knowledge = session.get("knowledge") or {}
        if not selected or not fields:
            return response(
                success=False,
                stage="report_failed",
                intent=intent,
                message="还没有确认最终题目。请先输入机器人创意并选择1、2或3。",
                suggested_actions=["输入1", "输入2", "输入3"],
            )

        report_json = build_report_json(fields, selected, knowledge)
        session["report_json"] = report_json
        session["report_revision"] = int(session.get("report_revision", 0)) + 1
        session["stage"] = "report_json_ready"
        save_session(session_key, session)

        return response(
            success=True,
            stage="report_json_ready",
            intent=intent,
            message=format_report_summary(report_json),
            payload_json=json.dumps(report_json, ensure_ascii=False, indent=2),
            data={"report_json": report_json, "revision": session["report_revision"]},
            suggested_actions=["下载报告", "查看报告结构数据", "修改第2页"],
        )

    if intent == "view_report_json":
        report_json = session.get("report_json")
        if not report_json:
            return response(
                success=False,
                stage="report_not_ready",
                intent=intent,
                message="当前还没有报告数据。请先选择题目并输入“生成报告”。",
            )
        return response(
            success=True,
            stage="report_json_ready",
            intent=intent,
            message="报告结构化JSON已返回到 payload_json 字段。",
            payload_json=json.dumps(report_json, ensure_ascii=False, indent=2),
            data={"report_json": report_json},
        )

    if intent == "modify_page_request":
        report_json = session.get("report_json")
        if not report_json:
            return response(
                success=False,
                stage="report_not_ready",
                intent=intent,
                message="当前还没有报告。请先生成报告。",
            )
        page_no = detect_page_number(user_message)
        if not page_no:
            return response(success=False, stage="modify_failed", intent=intent, message="请明确要修改第1至第7页中的哪一页。")

        key = f"page_{page_no}"
        page = report_json["pages"].get(key, {})
        page.setdefault("revision_notes", [])
        page["revision_notes"].append(user_message)

        # 有LLM时按用户要求重写该页；无LLM时保留要求供后续人工/再次生成使用。
        ai = call_llm_json(
            "你是竞赛报告编辑专家。只返回JSON对象，保持page_no/module/title/content结构，按修改要求重写这一页。",
            json.dumps({"page": page, "request": user_message, "project_title": report_json.get("project_title")}, ensure_ascii=False),
        )
        if ai and isinstance(ai, dict):
            ai.setdefault("page_no", page_no)
            ai.setdefault("module", page.get("module"))
            ai.setdefault("title", page.get("title"))
            ai.setdefault("content", page.get("content"))
            ai["revision_notes"] = page.get("revision_notes", [])
            report_json["pages"][key] = ai

        session["report_json"] = report_json
        session["report_revision"] = int(session.get("report_revision", 1)) + 1
        save_session(session_key, session)
        return response(
            success=True,
            stage="report_json_ready",
            intent=intent,
            message=f"已记录并处理第{page_no}页修改要求。当前版本：{session['report_revision']}。",
            payload_json=json.dumps(report_json, ensure_ascii=False, indent=2),
            data={"report_json": report_json, "revision": session["report_revision"]},
            suggested_actions=["下载报告", "继续修改其他页面"],
        )

    if intent == "word_export_pending":
        report_json = session.get("report_json")
        if not report_json:
            return response(
                success=False,
                stage="word_export_failed",
                intent=intent,
                message="当前还没有报告数据。请先完成选题并输入“生成报告”。",
                suggested_actions=["生成报告"],
            )

        try:
            export_info = export_report_to_word(report_json, session_key)
            print("DEBUG_WORD_SUCCESS:", export_info, flush=True)
        except Exception as exc:
            print("DEBUG_WORD_ERROR:", repr(exc), flush=True)
            return response(
                success=False,
                stage="word_export_failed",
                intent=intent,
                message=f"Word生成失败：{type(exc).__name__}。请查看Render日志。",
                payload_json=json.dumps(report_json, ensure_ascii=False, indent=2),
                data={"error": str(exc)},
            )

        session["word_file"] = export_info
        session["stage"] = "word_ready"
        save_session(session_key, session)

        return response(
            success=True,
            stage="word_ready",
            intent=intent,
            message=(
                "Word报告已经生成完成。\n\n"
                "文档已按7页竞赛报告结构排版，并自动加入与本项目相关的概念图、系统架构图、技术流程图和创新点图。\n\n"
                "请点击下方链接下载 .docx 文件："
            ),
            payload_json=json.dumps(report_json, ensure_ascii=False, indent=2),
            download_url=export_info["download_url"],
            data={"word_file": export_info, "report_json": report_json},
            files=[{"name": export_info["filename"], "url": export_info["download_url"], "type": "docx"}],
            suggested_actions=["下载报告", "修改指定页面后重新导出"],
        )

    return response(
        success=False,
        stage="unknown",
        intent="unknown",
        message="暂时无法识别该操作。你可以输入机器人创意、1/2/3、生成报告或下载报告。",
    )


# ============================================================
# API 路由
# ============================================================
@app.get("/")
def root() -> Dict[str, Any]:
    return {
        "service": "robot-competition-orchestrator",
        "version": APP_VERSION,
        "status": "ok",
        "llm_enabled": llm_available(),
        "knowledge_base_enabled": bool(KNOWLEDGE_API_URL),
    }


@app.get("/health")
def health() -> Dict[str, Any]:
    return {"ok": True, "version": APP_VERSION, "time": now_iso()}


@app.post("/api/v1/robot-competition/chat", response_model=ChatResponse)
def chat(req: ChatRequest) -> ChatResponse:
    return handle_chat(req)


@app.get("/api/v1/robot-competition/download/{filename}")
def download(filename: str):
    safe = Path(filename).name
    file_path = EXPORT_DIR / safe
    if not file_path.exists() or not file_path.is_file():
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(
        path=str(file_path),
        filename=safe,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
