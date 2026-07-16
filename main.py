from datetime import datetime, timezone
from typing import Any, Dict, List, Optional
import math
import re

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field


app = FastAPI(
    title="Robot Competition Orchestrator",
    version="0.2.0",
    description="智能机器人创意竞赛助手 V2 中央编排器测试接口"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# =========================
# 数据模型
# =========================

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
    data: Dict[str, Any] = Field(default_factory=dict)
    files: List[Dict[str, Any]] = Field(default_factory=list)
    suggested_actions: List[str] = Field(default_factory=list)
    version: int = 2
    updated_at: str


# =========================
# 简易会话存储
# Render 免费版重启后会清空
# 后续正式版可替换为数据库
# =========================

SESSION_STORE: Dict[str, Dict[str, Any]] = {}


# =========================
# 内置案例库 v0.2
# 注意：这里只是调试用的“样例案例库”
# 后续要替换为真实的往届获奖作品知识库
# =========================

REFERENCE_CASES = [
    {
        "title": "智能老人陪伴与健康监测机器人",
        "keywords": ["老人", "陪伴", "健康", "跌倒", "吃药", "家庭", "语音"],
        "track": "服务机器人"
    },
    {
        "title": "校园智能垃圾分类与回收机器人",
        "keywords": ["垃圾分类", "校园", "回收", "识别", "环保", "移动"],
        "track": "服务机器人"
    },
    {
        "title": "公共空间智能消毒巡检机器人",
        "keywords": ["消毒", "巡检", "公共空间", "环境", "导航", "安全"],
        "track": "特种机器人"
    },
    {
        "title": "导盲辅助与道路安全提醒机器人",
        "keywords": ["导盲", "道路", "安全", "避障", "语音", "辅助"],
        "track": "服务机器人"
    },
    {
        "title": "农业果蔬采摘与成熟度识别机器人",
        "keywords": ["农业", "采摘", "果蔬", "成熟度", "识别", "机械臂"],
        "track": "农业机器人"
    },
    {
        "title": "仓储物流自主搬运机器人",
        "keywords": ["物流", "仓储", "搬运", "路径规划", "导航", "调度"],
        "track": "工业机器人"
    },
    {
        "title": "水质检测与河道巡航机器人",
        "keywords": ["水质", "检测", "河道", "巡航", "传感器", "环保"],
        "track": "特种机器人"
    },
    {
        "title": "家庭教育陪伴与学习监督机器人",
        "keywords": ["儿童", "教育", "学习", "陪伴", "语音", "家庭"],
        "track": "服务机器人"
    }
]


# =========================
# 基础工具函数
# =========================

def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", "", text.strip())


def contains_any(text: str, words: List[str]) -> bool:
    return any(w in text for w in words)


def keyword_hit_count(text: str, words: List[str]) -> int:
    return sum(1 for w in words if w in text)


def clamp(value: float, low: float = 0, high: float = 100) -> float:
    return max(low, min(high, value))


def jaccard_similarity(a: List[str], b: List[str]) -> float:
    set_a = set(a)
    set_b = set(b)
    if not set_a or not set_b:
        return 0.0
    return len(set_a & set_b) / len(set_a | set_b)


# =========================
# 创意字段提取
# =========================

def extract_idea_fields(raw_idea: str) -> Dict[str, Any]:
    text = normalize_text(raw_idea)

    target_groups = []
    if contains_any(text, ["老人", "老年", "独居", "空巢"]):
        target_groups.append("独居老人/老年人")
    if contains_any(text, ["儿童", "学生", "孩子"]):
        target_groups.append("儿童/学生")
    if contains_any(text, ["残障", "盲人", "视障", "行动不便"]):
        target_groups.append("特殊人群")
    if contains_any(text, ["家庭", "家居", "居家"]):
        target_groups.append("家庭用户")
    if not target_groups:
        target_groups.append("普通用户")

    scenarios = []
    if contains_any(text, ["家庭", "家居", "居家"]):
        scenarios.append("家庭场景")
    if contains_any(text, ["校园", "学校", "教室"]):
        scenarios.append("校园场景")
    if contains_any(text, ["医院", "养老院", "社区"]):
        scenarios.append("医疗/养老/社区场景")
    if contains_any(text, ["工厂", "仓库", "物流"]):
        scenarios.append("工业/物流场景")
    if not scenarios:
        scenarios.append("通用服务场景")

    functions = []
    function_map = {
        "提醒吃药": ["吃药", "服药", "用药", "药物提醒"],
        "跌倒检测": ["跌倒", "摔倒"],
        "语音交互": ["语音", "聊天", "对话"],
        "智能家居联动": ["智能家居", "家居", "灯光", "空调", "门锁"],
        "健康监测": ["健康", "心率", "血压", "体温", "监测"],
        "环境感知": ["环境", "烟雾", "温湿度", "空气"],
        "自主导航": ["导航", "避障", "移动", "巡航"],
        "远程通知": ["远程", "通知", "报警", "家属", "手机"],
        "图像识别": ["识别", "视觉", "摄像头", "图像"],
        "机械执行": ["机械臂", "抓取", "递送", "搬运"]
    }

    for func, keys in function_map.items():
        if contains_any(text, keys):
            functions.append(func)

    if not functions:
        functions.append("基础人机交互")

    tech_modules = []
    tech_map = {
        "语音识别与自然语言交互": ["语音", "聊天", "对话"],
        "多传感器环境感知": ["传感器", "环境", "温湿度", "烟雾", "空气"],
        "视觉识别": ["摄像头", "视觉", "图像", "识别"],
        "姿态/跌倒检测算法": ["跌倒", "摔倒", "姿态"],
        "物联网与智能家居控制": ["智能家居", "家居", "灯光", "空调", "门锁"],
        "移动底盘与路径规划": ["移动", "导航", "避障", "巡航"],
        "远程通信与报警": ["远程", "报警", "通知", "家属"]
    }

    for module, keys in tech_map.items():
        if contains_any(text, keys):
            tech_modules.append(module)

    if not tech_modules:
        tech_modules.append("基础传感器采集与人机交互模块")

    keywords = []
    for word in [
        "老人", "独居", "家庭", "陪伴", "健康", "吃药", "跌倒", "语音", "智能家居",
        "导航", "识别", "报警", "家属", "儿童", "校园", "农业", "物流", "环保"
    ]:
        if word in text:
            keywords.append(word)

    if not keywords:
        keywords = functions[:3]

    return {
        "raw_idea": raw_idea,
        "target_groups": list(dict.fromkeys(target_groups)),
        "scenarios": list(dict.fromkeys(scenarios)),
        "core_functions": list(dict.fromkeys(functions)),
        "tech_modules": list(dict.fromkeys(tech_modules)),
        "keywords": list(dict.fromkeys(keywords))
    }


# =========================
# 相似度与评分
# =========================

def find_most_similar_case(keywords: List[str]) -> Dict[str, Any]:
    best_case = None
    best_score = 0.0

    for case in REFERENCE_CASES:
        score = jaccard_similarity(keywords, case["keywords"])
        if score > best_score:
            best_score = score
            best_case = case

    if best_case is None:
        best_case = REFERENCE_CASES[0]

    return {
        "title": best_case["title"],
        "track": best_case["track"],
        "similarity": round(best_score, 3)
    }


def score_candidate(fields: Dict[str, Any], candidate_index: int) -> Dict[str, Any]:
    functions = fields["core_functions"]
    tech_modules = fields["tech_modules"]
    keywords = fields["keywords"]

    similar_case = find_most_similar_case(keywords)
    highest_similarity = similar_case["similarity"]

    function_count = len(functions)
    tech_count = len(tech_modules)

    # 四项基础评分
    innovation_score = 70 + min(12, function_count * 3) + candidate_index * 2
    scientific_score = 68 + min(16, tech_count * 3) + candidate_index
    application_score = 72 + min(12, len(fields["target_groups"]) * 4) + min(8, len(fields["scenarios"]) * 3)
    expression_score = 70 + min(15, len(keywords) * 2)

    # 同质化与相似度惩罚
    similarity_penalty = highest_similarity * 18
    homogeneity_penalty = 8 if highest_similarity >= 0.45 else 3 if highest_similarity >= 0.25 else 0

    # 可实现性惩罚：功能太多但技术模块不足时惩罚
    feasibility_penalty = 0
    if function_count >= 6 and tech_count <= 3:
        feasibility_penalty = 8
    elif function_count >= 5 and tech_count <= 2:
        feasibility_penalty = 12

    # 资料完整度
    completeness_items = 0
    completeness_items += 1 if fields["target_groups"] else 0
    completeness_items += 1 if fields["scenarios"] else 0
    completeness_items += 1 if fields["core_functions"] else 0
    completeness_items += 1 if fields["tech_modules"] else 0
    data_completeness_bonus = completeness_items * 1.5

    # 差异化创新加分
    differentiation_bonus = 0
    if contains_any("".join(functions + tech_modules), ["多传感器", "姿态", "物联网", "路径规划"]):
        differentiation_bonus += 4
    if highest_similarity < 0.25:
        differentiation_bonus += 4
    elif highest_similarity < 0.45:
        differentiation_bonus += 2

    base_total = (
        innovation_score * 0.35
        + scientific_score * 0.30
        + application_score * 0.25
        + expression_score * 0.10
    )

    competitiveness_score = base_total - similarity_penalty - homogeneity_penalty - feasibility_penalty + data_completeness_bonus + differentiation_bonus
    competitiveness_score = clamp(competitiveness_score)

    return {
        "innovation_score": round(clamp(innovation_score), 1),
        "scientific_score": round(clamp(scientific_score), 1),
        "application_score": round(clamp(application_score), 1),
        "expression_score": round(clamp(expression_score), 1),
        "base_total": round(base_total, 1),
        "highest_similarity": highest_similarity,
        "similar_case": similar_case,
        "similarity_penalty": round(similarity_penalty, 1),
        "homogeneity_penalty": round(homogeneity_penalty, 1),
        "feasibility_penalty": round(feasibility_penalty, 1),
        "data_completeness_bonus": round(data_completeness_bonus, 1),
        "differentiation_bonus": round(differentiation_bonus, 1),
        "competitiveness_score": round(competitiveness_score, 1)
    }


def generate_candidates(fields: Dict[str, Any]) -> List[Dict[str, Any]]:
    target = fields["target_groups"][0]
    scenario = fields["scenarios"][0]
    functions = fields["core_functions"]
    tech = fields["tech_modules"]

    func_short = "、".join(functions[:3])
    tech_short = "、".join(tech[:2])

    candidates = [
        {
            "id": 1,
            "title": f"“智护星”——面向{target}的{scenario}安全陪伴机器人",
            "positioning": f"突出{func_short}，强调家庭安全与陪伴服务。",
            "core_tech": tech_short
        },
        {
            "id": 2,
            "title": f"“安居守护者”——基于多模态感知的{target}健康监护机器人",
            "positioning": f"突出健康监测、异常识别和远程通知，适合做成完整竞赛方案。",
            "core_tech": tech_short
        },
        {
            "id": 3,
            "title": f"“慧联家护”——融合智能家居联动的{target}主动服务机器人",
            "positioning": f"突出人机交互、环境感知与智能家居联动，展示性较强。",
            "core_tech": tech_short
        }
    ]

    for idx, candidate in enumerate(candidates):
        candidate["scores"] = score_candidate(fields, idx)

    return candidates


# =========================
# 意图识别与回复生成
# =========================

def detect_intent(message: str, session: Dict[str, Any]) -> str:
    text = normalize_text(message)

    if contains_any(text, ["重新生成", "重生成", "再生成", "换一批", "不满意"]):
        return "regenerate_titles"

    if re.fullmatch(r"[123]", text) or contains_any(text, ["选1", "选2", "选3", "第一个", "第二个", "第三个"]):
        return "select_title"

    if contains_any(text, ["生成报告", "写报告", "生成word", "导出word"]):
        return "generate_report"

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


def format_candidate_message(fields: Dict[str, Any], candidates: List[Dict[str, Any]], prefix: str) -> str:
    lines = []
    lines.append(prefix)
    lines.append("")
    lines.append("【创意字段提取】")
    lines.append(f"目标用户：{'、'.join(fields['target_groups'])}")
    lines.append(f"应用场景：{'、'.join(fields['scenarios'])}")
    lines.append(f"核心功能：{'、'.join(fields['core_functions'])}")
    lines.append(f"技术模块：{'、'.join(fields['tech_modules'])}")
    lines.append("")
    lines.append("【3个候选题目与获奖竞争力预测】")

    for c in candidates:
        s = c["scores"]
        lines.append("")
        lines.append(f"{c['id']}. {c['title']}")
        lines.append(f"定位：{c['positioning']}")
        lines.append(f"核心技术：{c['core_tech']}")
        lines.append(
            f"评分：创新性{s['innovation_score']}，科学性{s['scientific_score']}，"
            f"应用前景{s['application_score']}，设计表达{s['expression_score']}"
        )
        lines.append(
            f"加权基础分：{s['base_total']}；最高相似度：{s['highest_similarity']}；"
            f"相似案例：{s['similar_case']['title']}"
        )
        lines.append(
            f"惩罚/加分：相似度惩罚-{s['similarity_penalty']}，"
            f"同质化惩罚-{s['homogeneity_penalty']}，"
            f"可实现性惩罚-{s['feasibility_penalty']}，"
            f"资料完整度+{s['data_completeness_bonus']}，"
            f"差异化创新+{s['differentiation_bonus']}"
        )
        lines.append(f"获奖竞争力预测：{s['competitiveness_score']} / 100")

    lines.append("")
    lines.append("说明：这里的“获奖竞争力预测”不是获奖概率，而是基于公式的可解释竞争力评分。")
    lines.append("当前v0.2使用内置样例案例库进行相似度测试；后续接入真实往届获奖作品知识库后，会替换为真实案例对比。")
    lines.append("")
    lines.append("请选择：输入 1、2、3 确认题目；输入“重新生成”换一批；或直接补充你的要求。")

    return "\n".join(lines)


def handle_chat(req: ChatRequest) -> ChatResponse:
    session = SESSION_STORE.setdefault(req.session_id, {})
    user_message = req.message.strip()
    intent = detect_intent(user_message, session)

    if intent in ["create_project", "supplement_idea", "regenerate_titles"]:
        if intent == "create_project":
            raw_idea = user_message
            prefix = "已收到你的机器人创意，下面自动生成3个候选题目和初评结果。"
        elif intent == "supplement_idea":
            previous = session.get("raw_idea", "")
            raw_idea = previous + "\n补充要求：" + user_message
            prefix = "已把本次内容作为补充要求合并，下面重新生成候选题目和评分。"
        else:
            raw_idea = session.get("raw_idea", user_message)
            prefix = "已根据当前创意重新生成3个候选题目和评分。"

        fields = extract_idea_fields(raw_idea)
        candidates = generate_candidates(fields)

        session["raw_idea"] = raw_idea
        session["fields"] = fields
        session["candidates"] = candidates
        session["stage"] = "candidates_ready"

        return ChatResponse(
            success=True,
            stage="candidates_ready",
            intent=intent,
            message=format_candidate_message(fields, candidates, prefix),
            data={
                "fields": fields,
                "candidates": candidates,
                "current_stage": "candidates_ready"
            },
            files=[],
            suggested_actions=["输入1/2/3选择题目", "重新生成", "补充要求"],
            updated_at=now_iso()
        )

    if intent == "select_title":
        selected_id = parse_selection(user_message)
        candidates = session.get("candidates", [])

        if not selected_id or not candidates:
            return ChatResponse(
                success=False,
                stage="selection_failed",
                intent="select_title",
                message="暂时没有可选择的候选题目。请先输入你的机器人创意，生成3个候选题目。",
                data={},
                files=[],
                suggested_actions=["重新输入创意"],
                updated_at=now_iso()
            )

        selected = next((c for c in candidates if c["id"] == selected_id), None)
        if selected is None:
            return ChatResponse(
                success=False,
                stage="selection_failed",
                intent="select_title",
                message="没有找到对应编号的题目。请输入 1、2 或 3。",
                data={},
                files=[],
                suggested_actions=["输入1", "输入2", "输入3"],
                updated_at=now_iso()
            )

        session["selected_title"] = selected
        session["stage"] = "title_confirmed"

        msg = (
            f"已确认最终题目：\n\n"
            f"{selected['title']}\n\n"
            f"该题目的获奖竞争力预测为：{selected['scores']['competitiveness_score']} / 100。\n\n"
            f"下一步可以输入“生成报告”，系统将生成结构化7页报告JSON。"
        )

        return ChatResponse(
            success=True,
            stage="title_confirmed",
            intent="select_title",
            message=msg,
            data={
                "selected_title": selected,
                "current_stage": "title_confirmed"
            },
            files=[],
            suggested_actions=["生成报告", "重新生成题目", "补充要求"],
            updated_at=now_iso()
        )

    if intent == "generate_report":
        selected = session.get("selected_title")
        fields = session.get("fields")

        if not selected or not fields:
            return ChatResponse(
                success=False,
                stage="report_failed",
                intent="generate_report",
                message="还没有确认最终题目。请先输入 1、2 或 3 选择一个候选题目。",
                data={},
                files=[],
                suggested_actions=["输入1", "输入2", "输入3"],
                updated_at=now_iso()
            )

        report_json = {
            "page_1": "封面：作品名称、参赛方向、团队信息、关键词",
            "page_2": "设计背景：用户痛点、现实需求、竞赛价值",
            "page_3": "产品整体结构：硬件结构、软件架构、交互流程",
            "page_4": "软硬件功能设计：传感器、执行器、通信模块与功能闭环",
            "page_5": "关键技术：算法、控制、感知、数据处理与安全策略",
            "page_6": "项目创新点：差异化设计、技术创新、应用创新",
            "page_7": "行业应用前景：推广场景、社会价值、商业化与迭代方向"
        }

        session["report_json"] = report_json
        session["stage"] = "report_json_ready"

        msg = (
            f"已基于最终题目生成结构化7页报告JSON框架。\n\n"
            f"最终题目：{selected['title']}\n\n"
            f"报告结构：\n"
            f"1. 封面\n"
            f"2. 设计背景\n"
            f"3. 产品整体结构\n"
            f"4. 软硬件功能设计\n"
            f"5. 关键技术\n"
            f"6. 项目创新点\n"
            f"7. 行业应用前景\n\n"
            f"下一版将继续加入：6张配图生成、固定模板排版、Word文件导出、指定页面局部修改。"
        )

        return ChatResponse(
            success=True,
            stage="report_json_ready",
            intent="generate_report",
            message=msg,
            data={
                "report_json": report_json,
                "current_stage": "report_json_ready"
            },
            files=[],
            suggested_actions=["生成配图", "导出Word", "修改第3页"],
            updated_at=now_iso()
        )

    return ChatResponse(
        success=False,
        stage="unknown",
        intent="unknown",
        message="暂时无法识别你的操作。你可以输入机器人创意、输入1/2/3选择题目，或输入“重新生成”。",
        data={},
        files=[],
        suggested_actions=["输入创意", "重新生成", "输入1/2/3"],
        updated_at=now_iso()
    )


# =========================
# API
# =========================

@app.get("/health")
def health() -> Dict[str, str]:
    return {
        "status": "ok",
        "service": "robot_competition_orchestrator",
        "version": "0.2.0"
    }


@app.post("/api/v1/robot-competition/chat", response_model=ChatResponse)
def chat(req: ChatRequest) -> ChatResponse:
    return handle_chat(req)
