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

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Cm, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image, ImageDraw, ImageFont

APP_VERSION = "1.0.0"

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

EXPORT_DIR = Path(os.getenv("EXPORT_DIR", "/tmp/robot_competition_exports"))
EXPORT_DIR.mkdir(parents=True, exist_ok=True)

SESSION_DB_PATH = os.getenv(
    "SESSION_DB_PATH",
    "/tmp/robot_competition_sessions.sqlite3",
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
    title = selected["title"]
    target = "、".join(fields.get("target_groups", []))
    scenes = "、".join(fields.get("scenarios", []))
    funcs = fields.get("core_functions", [])
    techs = fields.get("tech_modules", [])
    pain = fields.get("pain_points", [])

    pages = {
        "page_1": {
            "page_no": 1,
            "module": "封面",
            "title": title,
            "content": {
                "subtitle": "智能机器人创意竞赛项目设计报告",
                "project_positioning": selected.get("positioning", ""),
                "keywords": fields.get("keywords", [])[:8],
                "competition_score": selected.get("scores", {}).get("competitiveness_score"),
            },
        },
        "page_2": {
            "page_no": 2,
            "module": "设计背景",
            "title": "设计背景与用户需求",
            "content": {
                "background": [
                    f"项目面向{target}，聚焦{scenes}中的真实使用需求。",
                    f"用户原始创意：{fields.get('raw_idea', '')}",
                ],
                "pain_points": pain,
                "design_goals": [
                    fields.get("design_goal", ""),
                    "从单一功能演示升级为可持续运行的机器人服务闭环。",
                    "兼顾竞赛展示性、工程可实现性和后续扩展空间。",
                ],
            },
        },
        "page_3": {
            "page_no": 3,
            "module": "产品整体结构",
            "title": "机器人整体结构设计",
            "content": {
                "system_layers": [
                    "感知层：采集视觉、环境、状态与用户输入",
                    "认知层：完成识别、融合、风险判断和任务规划",
                    "执行层：通过移动底盘、机械执行机构、语音或联网设备完成动作",
                    "交互层：向用户提供提示、确认、反馈与远程信息同步",
                ],
                "core_modules": funcs,
                "mechanical_concept": f"围绕{scenes}构建模块化机器人本体，可按实际任务配置移动底盘、传感组件、交互终端和执行机构。",
            },
        },
        "page_4": {
            "page_no": 4,
            "module": "软硬件功能设计",
            "title": "硬件与软件功能设计",
            "content": {
                "hardware": [f"传感与采集：{t}" for t in techs[:3]] + ["主控与通信：负责数据融合、任务调度和远程连接"],
                "software": [f"功能模块：{f}" for f in funcs],
                "interaction_logic": "用户输入/环境事件 → 传感采集 → 智能判断 → 执行动作 → 反馈确认 → 数据记录",
            },
        },
        "page_5": {
            "page_no": 5,
            "module": "关键技术",
            "title": "关键技术与实现路线",
            "content": {
                "key_technologies": [
                    {"name": t, "implementation": f"围绕{t}设计输入、算法处理、阈值/模型判定和输出动作，并设置可量化测试指标。"}
                    for t in techs
                ],
                "technical_route": ["数据采集", "预处理", "特征/状态识别", "多源信息融合", "任务决策", "执行控制", "反馈与记录"],
                "engineering_metrics": ["识别准确率/误报率", "系统响应延迟", "任务执行成功率", "异常告警送达率", "连续运行稳定性"],
                "risk_control": ["关键安全事件设置人工确认与兜底", "网络异常时保留本地基础能力", "避免把概念功能写成已完成实测", "保护图像、语音和用户数据"],
            },
        },
        "page_6": {
            "page_no": 6,
            "module": "项目创新点",
            "title": "项目创新点与差异化分析",
            "content": {
                "innovation_points": [
                    {"name": "需求驱动", "description": f"从{target}在{scenes}中的具体痛点出发，而不是简单堆叠功能。"},
                    {"name": "机器人服务闭环", "description": "把感知、判断、执行、反馈串成完整闭环。"},
                    {"name": "多模块协同", "description": f"将{'、'.join(funcs[:4])}进行任务级协同。"},
                    {"name": "可扩展工程设计", "description": "采用模块化结构，便于后续增加传感器、执行机构和知识库。"},
                ],
                "differentiation_analysis": {
                    "knowledge_base_used": knowledge.get("knowledge_base_used", False),
                    "similar_award_cases": knowledge.get("similar_award_cases", [])[:3],
                    "current_boundary": "未接入真实获奖资料时，不把样例相似度宣传为真实获奖证据。",
                },
            },
        },
        "page_7": {
            "page_no": 7,
            "module": "行业应用前景",
            "title": "行业应用前景与落地路径",
            "content": {
                "application_scenarios": fields.get("scenarios", []),
                "target_users": fields.get("target_groups", []),
                "deployment_paths": ["完成核心功能样机", "开展目标场景测试", "记录失败案例并迭代", "形成模块化版本", "与行业平台/智能设备接口联动"],
                "social_value": ["降低重复性人工工作负担", "提高异常发现与处置效率", "提升目标用户安全性和便利性", "形成可持续迭代的数据与服务闭环"],
                "future_iterations": ["接入真实往届获奖作品知识库", "增加真实场景测试数据", "优化机器人外观与结构设计", "完善竞赛答辩材料和演示视频"],
            },
        },
    }

    return {
        "schema_version": "1.0",
        "project_title": title,
        "raw_idea": fields.get("raw_idea", ""),
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
def _find_cjk_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = [
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJKsc-Regular.otf",
        "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for path in candidates:
        if Path(path).exists():
            try:
                return ImageFont.truetype(path, size=size)
            except Exception:
                pass
    return ImageFont.load_default()


def _wrap_text(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, max_width: int) -> List[str]:
    if not text:
        return []
    lines: List[str] = []
    current = ""
    for ch in text:
        trial = current + ch
        bbox = draw.textbbox((0, 0), trial, font=font)
        if bbox[2] - bbox[0] > max_width and current:
            lines.append(current)
            current = ch
        else:
            current = trial
    if current:
        lines.append(current)
    return lines


def _save_diagram(report: Dict[str, Any], kind: str, digest: str) -> Path:
    width, height = 1400, 820
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)
    font_title = _find_cjk_font(44)
    font_head = _find_cjk_font(30)
    font_body = _find_cjk_font(24)

    title = report.get("project_title", "智能机器人项目")
    fields = report.get("idea_fields", {})
    funcs = fields.get("core_functions", [])[:5]
    techs = fields.get("tech_modules", [])[:5]

    draw.rounded_rectangle((35, 35, width - 35, height - 35), radius=30, outline=(70, 90, 120), width=4)
    for idx, line in enumerate(_wrap_text(draw, title, font_title, width - 180)[:2]):
        draw.text((90, 70 + idx * 58), line, font=font_title, fill=(30, 40, 60))

    top = 210
    if kind == "overview":
        columns = [
            ("目标用户", fields.get("target_groups", [])),
            ("应用场景", fields.get("scenarios", [])),
            ("核心功能", funcs),
            ("关键技术", techs),
        ]
        box_w = 285
        for i, (head, items) in enumerate(columns):
            x = 70 + i * 325
            draw.rounded_rectangle((x, top, x + box_w, 700), radius=22, outline=(110, 125, 145), width=3)
            draw.text((x + 20, top + 25), head, font=font_head, fill=(25, 45, 75))
            y = top + 90
            for item in items[:5]:
                for line in _wrap_text(draw, "• " + str(item), font_body, box_w - 35)[:2]:
                    draw.text((x + 18, y), line, font=font_body, fill=(60, 65, 75))
                    y += 38
                y += 8

    elif kind == "architecture":
        layers = [
            ("感知层", "视觉 / 环境 / 状态 / 用户输入"),
            ("认知决策层", "识别 / 信息融合 / 风险判断 / 任务规划"),
            ("执行层", "移动 / 机械执行 / 语音 / 联网控制"),
            ("反馈层", "提醒 / 确认 / 记录 / 远程通知"),
        ]
        y = top
        for head, body in layers:
            draw.rounded_rectangle((190, y, 1210, y + 105), radius=22, outline=(90, 110, 140), width=3)
            draw.text((230, y + 24), head, font=font_head, fill=(25, 45, 75))
            draw.text((470, y + 30), body, font=font_body, fill=(60, 65, 75))
            y += 135
            if y < 720:
                draw.line((700, y - 30, 700, y - 8), fill=(90, 110, 140), width=5)
                draw.polygon([(690, y - 15), (710, y - 15), (700, y - 3)], fill=(90, 110, 140))

    elif kind == "workflow":
        steps = ["数据采集", "预处理", "识别/融合", "任务决策", "执行控制", "反馈记录"]
        x = 85
        y = 360
        for i, step in enumerate(steps):
            bw = 175
            draw.rounded_rectangle((x, y, x + bw, y + 115), radius=20, outline=(90, 110, 140), width=3)
            draw.text((x + 20, y + 35), step, font=font_body, fill=(35, 50, 75))
            if i < len(steps) - 1:
                draw.line((x + bw + 10, y + 58, x + bw + 45, y + 58), fill=(90, 110, 140), width=4)
                draw.polygon([(x + bw + 38, y + 48), (x + bw + 52, y + 58), (x + bw + 38, y + 68)], fill=(90, 110, 140))
            x += 215

    else:  # innovation
        items = [
            "真实需求驱动",
            "感知-决策-执行闭环",
            "多模块协同",
            "工程可实现",
            "可扩展知识库",
        ]
        center = (700, 470)
        draw.ellipse((540, 350, 860, 590), outline=(70, 90, 120), width=4)
        draw.text((600, 440), "项目创新核心", font=font_head, fill=(30, 45, 70))
        positions = [(120, 250), (980, 250), (80, 590), (1010, 590), (535, 680)]
        for item, (x, y) in zip(items, positions):
            draw.rounded_rectangle((x, y, x + 310, y + 90), radius=20, outline=(110, 125, 145), width=3)
            draw.text((x + 18, y + 28), item, font=font_body, fill=(50, 60, 80))
            draw.line((x + 155, y + 45, center[0], center[1]), fill=(165, 170, 180), width=2)

    path = EXPORT_DIR / f"diagram_{kind}_{digest}.png"
    img.save(path, format="PNG")
    return path


def generate_project_images(report: Dict[str, Any], session_key: str) -> Dict[str, str]:
    digest = hashlib.sha1(f"{session_key}-{report.get('project_title','')}".encode("utf-8")).hexdigest()[:10]
    result = {}
    for kind in ["overview", "architecture", "workflow", "innovation"]:
        try:
            result[kind] = str(_save_diagram(report, kind, digest))
        except Exception as exc:
            print("DIAGRAM_ERROR", kind, repr(exc), flush=True)
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
    p.space_before = Pt(8)
    p.space_after = Pt(4)


def _add_paragraph(doc: Document, text: Any, bold_prefix: Optional[str] = None) -> None:
    if text is None:
        return
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.35
    p.paragraph_format.space_after = Pt(4)
    if bold_prefix and str(text).startswith(bold_prefix):
        r1 = p.add_run(bold_prefix)
        _set_run_font(r1, bold=True)
        r2 = p.add_run(str(text)[len(bold_prefix):])
        _set_run_font(r2)
    else:
        r = p.add_run(str(text))
        _set_run_font(r)


def _add_bullets(doc: Document, items: Any, limit: int = 8) -> None:
    if not isinstance(items, list):
        if items:
            _add_paragraph(doc, items)
        return
    for item in items[:limit]:
        p = doc.add_paragraph(style=None)
        p.paragraph_format.left_indent = Cm(0.55)
        p.paragraph_format.first_line_indent = Cm(-0.3)
        p.paragraph_format.line_spacing = 1.25
        if isinstance(item, dict):
            text = item.get("name") or item.get("title") or ""
            desc = item.get("description") or item.get("implementation") or ""
            value = f"• {text}：{desc}" if desc else f"• {text}"
        else:
            value = f"• {item}"
        r = p.add_run(value)
        _set_run_font(r)


def _add_image(doc: Document, path: Optional[str], caption: str) -> None:
    if not path or not Path(path).exists():
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run()
    r.add_picture(path, width=Cm(15.5))
    c = doc.add_paragraph()
    c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rr = c.add_run(caption)
    _set_run_font(rr, size=9)


def _prepare_document() -> Document:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(1.6)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Microsoft YaHei"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    normal.font.size = Pt(10.5)
    return doc


def _render_cover(doc: Document, report: Dict[str, Any], images: Dict[str, str]) -> None:
    doc.add_paragraph("\n\n")
    _add_title(doc, report.get("project_title", "智能机器人竞赛项目"), 24)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("智能机器人创意竞赛 · 项目设计报告")
    _set_run_font(r, size=14)
    doc.add_paragraph("")
    _add_image(doc, images.get("overview"), "图1  项目概念与核心模块概览")
    selected = report.get("selected_candidate", {})
    table = doc.add_table(rows=3, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    pairs = [
        ("项目定位", selected.get("positioning", "")),
        ("核心技术", selected.get("core_tech", "")),
        ("竞赛竞争力", f"{selected.get('scores', {}).get('competitiveness_score', '')} / 100"),
    ]
    for row, (k, v) in zip(table.rows, pairs):
        row.cells[0].text = k
        row.cells[1].text = str(v)
        _set_cell_shading(row.cells[0], "E8EEF7")
        for c in row.cells:
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for p in c.paragraphs:
                for run in p.runs:
                    _set_run_font(run, size=10, bold=(c is row.cells[0]))


def _render_page(doc: Document, report: Dict[str, Any], page_no: int, images: Dict[str, str]) -> None:
    page = report["pages"][f"page_{page_no}"]
    content = page.get("content", {})
    _add_heading(doc, f"第{page_no}页  {page.get('title', '')}", 1)

    if page_no == 2:
        _add_heading(doc, "设计背景", 2)
        _add_bullets(doc, content.get("background", []))
        _add_heading(doc, "用户痛点", 2)
        _add_bullets(doc, content.get("pain_points", []))
        _add_heading(doc, "设计目标", 2)
        _add_bullets(doc, content.get("design_goals", []))

    elif page_no == 3:
        _add_image(doc, images.get("architecture"), "图2  机器人系统总体架构")
        _add_heading(doc, "系统分层", 2)
        _add_bullets(doc, content.get("system_layers", []))
        _add_heading(doc, "核心模块", 2)
        _add_bullets(doc, content.get("core_modules", []))
        _add_heading(doc, "本体结构思路", 2)
        _add_paragraph(doc, content.get("mechanical_concept", ""))

    elif page_no == 4:
        _add_heading(doc, "硬件设计", 2)
        _add_bullets(doc, content.get("hardware", []))
        _add_heading(doc, "软件功能", 2)
        _add_bullets(doc, content.get("software", []))
        _add_heading(doc, "交互逻辑", 2)
        _add_paragraph(doc, content.get("interaction_logic", ""))

    elif page_no == 5:
        _add_image(doc, images.get("workflow"), "图3  项目关键技术与运行流程")
        _add_heading(doc, "关键技术", 2)
        _add_bullets(doc, content.get("key_technologies", []))
        _add_heading(doc, "技术路线", 2)
        _add_bullets(doc, content.get("technical_route", []))
        _add_heading(doc, "量化指标", 2)
        _add_bullets(doc, content.get("engineering_metrics", []))
        _add_heading(doc, "风险控制", 2)
        _add_bullets(doc, content.get("risk_control", []))

    elif page_no == 6:
        _add_image(doc, images.get("innovation"), "图4  项目创新点与差异化设计")
        _add_heading(doc, "创新点", 2)
        _add_bullets(doc, content.get("innovation_points", []))
        diff = content.get("differentiation_analysis", {})
        _add_heading(doc, "知识库与差异化边界", 2)
        _add_paragraph(doc, json.dumps(diff, ensure_ascii=False, indent=2))

    elif page_no == 7:
        _add_heading(doc, "应用场景", 2)
        _add_bullets(doc, content.get("application_scenarios", []))
        _add_heading(doc, "落地路径", 2)
        _add_bullets(doc, content.get("deployment_paths", []))
        _add_heading(doc, "社会价值", 2)
        _add_bullets(doc, content.get("social_value", []))
        _add_heading(doc, "后续迭代", 2)
        _add_bullets(doc, content.get("future_iterations", []))


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

    filename = f"robot_competition_report_{digest}.docx"

    path = EXPORT_DIR / filename
    doc.save(path)

    return {
    "filename": filename,
    "file_path": str(path),
    "download_url": f"{PUBLIC_BASE_URL}/api/v1/robot-competition/download/{filename}",
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
