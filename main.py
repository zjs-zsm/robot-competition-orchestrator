from datetime import datetime, timezone
from typing import Any, Dict, List, Optional
import json
import re

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field


APP_VERSION = "0.4.0"

app = FastAPI(
    title="Robot Competition Orchestrator",
    version=APP_VERSION,
    description="智能机器人创意竞赛助手 V2 中央编排器"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


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
    version: int = 4
    updated_at: str


SESSION_STORE: Dict[str, Dict[str, Any]] = {}


# 当前仍是“样例案例库”。
# 下一阶段可把这里替换为真实往届获奖作品知识库 / 外部检索结果。
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


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", "", (text or "").strip())


def clean_message(text: str) -> str:
    prefixes = [
        "启动V2中央编排器",
        "启动v2中央编排器",
        "启动机器人竞赛助手",
        "启动达标版竞赛助手",
        "运行机器人竞赛中央编排器"
    ]

    result = (text or "").strip()

    for prefix in prefixes:
        result = result.replace(prefix, "").strip()

    return result if result else (text or "").strip()


def contains_any(text: str, words: List[str]) -> bool:
    return any(word in text for word in words)


def clamp(value: float, low: float = 0, high: float = 100) -> float:
    return max(low, min(high, value))


def jaccard_similarity(first_words: List[str], second_words: List[str]) -> float:
    first_set = set(first_words)
    second_set = set(second_words)

    if not first_set or not second_set:
        return 0.0

    return len(first_set & second_set) / len(first_set | second_set)


def build_session_key(req: ChatRequest) -> str:
    """
    关键修复：
    不再使用固定 session_key。
    同一用户 + 同一会话使用同一状态；不同会话互不污染。
    """
    user_id = (req.user_id or "anonymous").strip()
    session_id = (req.session_id or "default").strip()
    return f"{user_id}:{session_id}"


def extract_idea_fields(raw_idea: str) -> Dict[str, Any]:
    text = normalize_text(raw_idea)

    target_groups: List[str] = []

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

    scenarios: List[str] = []

    if contains_any(text, ["家庭", "家居", "居家"]):
        scenarios.append("家庭场景")
    if contains_any(text, ["校园", "学校", "教室"]):
        scenarios.append("校园场景")
    if contains_any(text, ["医院", "养老院", "社区"]):
        scenarios.append("医疗/养老/社区场景")
    if contains_any(text, ["工厂", "仓库", "物流"]):
        scenarios.append("工业/物流场景")
    if contains_any(text, ["农业", "果园", "温室", "农田"]):
        scenarios.append("农业场景")
    if contains_any(text, ["河道", "水质", "湖泊", "环境监测"]):
        scenarios.append("环境监测场景")
    if not scenarios:
        scenarios.append("通用服务场景")

    function_map = {
        "提醒吃药": ["吃药", "服药", "用药", "药物提醒"],
        "跌倒检测": ["跌倒", "摔倒"],
        "语音交互": ["语音", "聊天", "对话", "陪伴"],
        "智能家居联动": ["智能家居", "家居", "灯光", "空调", "门锁"],
        "健康监测": ["健康", "心率", "血压", "体温", "监测"],
        "环境感知": ["环境", "烟雾", "温湿度", "空气"],
        "自主导航": ["导航", "避障", "移动", "巡航"],
        "远程通知": ["远程", "通知", "报警", "家属", "手机"],
        "图像识别": ["识别", "视觉", "摄像头", "图像"],
        "机械执行": ["机械臂", "抓取", "递送", "搬运"],
        "垃圾分类": ["垃圾分类", "回收"],
        "水质检测": ["水质", "河道", "水体"],
        "成熟度识别": ["成熟度", "果蔬", "采摘"]
    }

    functions: List[str] = []
    for function_name, keys in function_map.items():
        if contains_any(text, keys):
            functions.append(function_name)
    if not functions:
        functions.append("基础人机交互")

    tech_map = {
        "语音识别与自然语言交互": ["语音", "聊天", "对话", "陪伴"],
        "多传感器环境感知": ["传感器", "环境", "温湿度", "烟雾", "空气"],
        "视觉识别": ["摄像头", "视觉", "图像", "识别"],
        "姿态/跌倒检测算法": ["跌倒", "摔倒", "姿态"],
        "物联网与智能家居控制": ["智能家居", "家居", "灯光", "空调", "门锁"],
        "移动底盘与路径规划": ["移动", "导航", "避障", "巡航"],
        "远程通信与报警": ["远程", "报警", "通知", "家属"],
        "机械臂控制": ["机械臂", "抓取", "递送", "搬运"],
        "边缘计算与本地决策": ["边缘", "本地", "低延迟", "离线"]
    }

    tech_modules: List[str] = []
    for module_name, keys in tech_map.items():
        if contains_any(text, keys):
            tech_modules.append(module_name)
    if not tech_modules:
        tech_modules.append("基础传感器采集与人机交互模块")

    keyword_candidates = [
        "老人", "独居", "家庭", "陪伴", "健康", "吃药", "跌倒", "语音",
        "智能家居", "导航", "识别", "报警", "家属", "儿童", "校园",
        "农业", "物流", "环保", "水质", "河道", "机械臂", "垃圾分类"
    ]

    keywords = [word for word in keyword_candidates if word in text]
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


def find_most_similar_case(keywords: List[str]) -> Dict[str, Any]:
    best_case = REFERENCE_CASES[0]
    best_score = 0.0

    for case in REFERENCE_CASES:
        score = jaccard_similarity(keywords, case["keywords"])
        if score > best_score:
            best_score = score
            best_case = case

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

    innovation_score = 70 + min(12, function_count * 3) + candidate_index * 2
    scientific_score = 68 + min(16, tech_count * 3) + candidate_index
    application_score = (
        72
        + min(12, len(fields["target_groups"]) * 4)
        + min(8, len(fields["scenarios"]) * 3)
    )
    expression_score = 70 + min(15, len(keywords) * 2)

    similarity_penalty = highest_similarity * 18

    if highest_similarity >= 0.45:
        homogeneity_penalty = 8
    elif highest_similarity >= 0.25:
        homogeneity_penalty = 3
    else:
        homogeneity_penalty = 0

    feasibility_penalty = 0
    if function_count >= 6 and tech_count <= 3:
        feasibility_penalty = 8
    elif function_count >= 5 and tech_count <= 2:
        feasibility_penalty = 12

    completeness_items = sum([
        bool(fields["target_groups"]),
        bool(fields["scenarios"]),
        bool(fields["core_functions"]),
        bool(fields["tech_modules"])
    ])
    data_completeness_bonus = completeness_items * 1.5

    differentiation_bonus = 0
    combined_modules = "".join(functions + tech_modules)
    if contains_any(combined_modules, ["多传感器", "姿态", "物联网", "路径规划", "边缘计算"]):
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

    competitiveness_score = (
        base_total
        - similarity_penalty
        - homogeneity_penalty
        - feasibility_penalty
        + data_completeness_bonus
        + differentiation_bonus
    )

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
        "competitiveness_score": round(clamp(competitiveness_score), 1)
    }


def generate_candidates(fields: Dict[str, Any]) -> List[Dict[str, Any]]:
    target = fields["target_groups"][0]
    scenario = fields["scenarios"][0]
    functions = fields["core_functions"]
    tech_modules = fields["tech_modules"]

    function_summary = "、".join(functions[:3])
    technology_summary = "、".join(tech_modules[:3])

    candidates = [
        {
            "id": 1,
            "title": f"“智护星”——面向{target}的{scenario}安全陪伴机器人",
            "positioning": f"突出{function_summary}，强调安全与陪伴服务闭环。",
            "core_tech": technology_summary
        },
        {
            "id": 2,
            "title": f"“安居守护者”——基于多模态感知的{target}健康监护机器人",
            "positioning": "突出健康监测、异常识别和远程通知，适合形成完整竞赛方案。",
            "core_tech": technology_summary
        },
        {
            "id": 3,
            "title": f"“慧联家护”——融合环境感知与智能联动的{target}主动服务机器人",
            "positioning": "突出人机交互、环境感知与设备联动，强调展示性和主动服务。",
            "core_tech": technology_summary
        }
    ]

    for index, candidate in enumerate(candidates):
        candidate["scores"] = score_candidate(fields, index)

    return candidates


def format_candidate_message(
    fields: Dict[str, Any],
    candidates: List[Dict[str, Any]],
    prefix: str
) -> str:
    lines = [
        prefix,
        "",
        "【创意字段提取】",
        "目标用户：" + "、".join(fields["target_groups"]),
        "应用场景：" + "、".join(fields["scenarios"]),
        "核心功能：" + "、".join(fields["core_functions"]),
        "技术模块：" + "、".join(fields["tech_modules"]),
        "",
        "【3个候选题目与获奖竞争力预测】"
    ]

    for candidate in candidates:
        scores = candidate["scores"]
        lines.extend([
            "",
            f"{candidate['id']}. {candidate['title']}",
            f"定位：{candidate['positioning']}",
            f"核心技术：{candidate['core_tech']}",
            (
                "评分："
                f"创新性{scores['innovation_score']}，"
                f"科学性{scores['scientific_score']}，"
                f"应用前景{scores['application_score']}，"
                f"设计表达{scores['expression_score']}"
            ),
            (
                f"加权基础分：{scores['base_total']}；"
                f"最高相似度：{scores['highest_similarity']}；"
                f"相似案例：{scores['similar_case']['title']}"
            ),
            (
                "惩罚/加分："
                f"相似度惩罚-{scores['similarity_penalty']}，"
                f"同质化惩罚-{scores['homogeneity_penalty']}，"
                f"可实现性惩罚-{scores['feasibility_penalty']}，"
                f"资料完整度+{scores['data_completeness_bonus']}，"
                f"差异化创新+{scores['differentiation_bonus']}"
            ),
            f"获奖竞争力预测：{scores['competitiveness_score']} / 100"
        ])

    lines.extend([
        "",
        "说明：这里的“获奖竞争力预测”是可解释竞争力评分，不是实际获奖概率。",
        f"当前 v{APP_VERSION} 仍使用内置样例案例库；下一阶段将替换为真实往届获奖作品知识库。",
        "",
        "请选择：输入1、2、3确认题目；输入“重新生成”换一批；或直接补充你的要求。"
    ])

    return "\n".join(lines)


def image_spec(image_type: str, purpose: str, prompt: str) -> Dict[str, Any]:
    return {
        "required": True,
        "image_type": image_type,
        "purpose": purpose,
        "prompt": prompt,
        "aspect_ratio": "16:9",
        "file_url": "",
        "generation_status": "pending"
    }


def build_report_json(fields: Dict[str, Any], selected: Dict[str, Any]) -> Dict[str, Any]:
    """
    竞赛级固定 7 页结构。
    核心原则：不做普通产品介绍，而是突出
    感知 -> 决策 -> 执行 -> 反馈 的机器人闭环。
    """
    title = selected["title"]
    scores = selected["scores"]

    target = "、".join(fields.get("target_groups", []))
    scenario = "、".join(fields.get("scenarios", []))
    core_functions = fields.get("core_functions", [])
    tech_modules = fields.get("tech_modules", [])
    keywords = fields.get("keywords", [])

    function_text = "、".join(core_functions)
    technology_text = "、".join(tech_modules)

    report = {
        "schema_version": f"robot-competition-report-v{APP_VERSION}",
        "project_title": title,
        "revision": 1,
        "generated_at": now_iso(),
        "quality_standard": {
            "orientation": "大学生机器人创新竞赛",
            "core_dimensions": ["创新性", "科学性", "工程可实现性", "应用价值", "设计表达"],
            "robotic_closed_loop": ["感知", "决策", "执行", "反馈"],
            "anti_hallucination": [
                "不得把样例案例写成真实获奖作品",
                "不得虚构市场规模和政策数据",
                "不得把概念功能描述为已完成实测"
            ]
        },
        "competition_prediction": {
            "name": "获奖竞争力预测",
            "score": scores.get("competitiveness_score", 0),
            "base_total": scores.get("base_total", 0),
            "innovation_score": scores.get("innovation_score", 0),
            "scientific_score": scores.get("scientific_score", 0),
            "application_score": scores.get("application_score", 0),
            "expression_score": scores.get("expression_score", 0),
            "highest_similarity": scores.get("highest_similarity", 0),
            "similar_case": scores.get("similar_case", {}),
            "similarity_penalty": scores.get("similarity_penalty", 0),
            "homogeneity_penalty": scores.get("homogeneity_penalty", 0),
            "feasibility_penalty": scores.get("feasibility_penalty", 0),
            "data_completeness_bonus": scores.get("data_completeness_bonus", 0),
            "differentiation_bonus": scores.get("differentiation_bonus", 0),
            "disclaimer": "该数值是基于固定公式计算的获奖竞争力预测，不是实际获奖概率。"
        },
        "project_fields": {
            "raw_idea": fields.get("raw_idea", ""),
            "target_groups": fields.get("target_groups", []),
            "scenarios": fields.get("scenarios", []),
            "core_functions": core_functions,
            "tech_modules": tech_modules,
            "keywords": keywords
        },
        "pages": {
            "page_1": {
                "page_no": 1,
                "module": "封面",
                "title": title,
                "subtitle": "智能机器人创新竞赛项目报告",
                "content": {
                    "competition_name": "智能机器人创新竞赛",
                    "project_name": title,
                    "design_concept": f"面向{target}，以“主动感知、智能判断、可靠执行、及时反馈”为核心设计理念。",
                    "track": "待根据正式竞赛规则确认",
                    "keywords": keywords,
                    "team_name": "待填写",
                    "school_name": "待填写",
                    "members": "待填写",
                    "advisor": "待填写",
                    "date": "待填写"
                },
                "image": {
                    "required": False,
                    "image_type": "封面主视觉",
                    "purpose": "由最终报告模板或效果图模块完成",
                    "prompt": "",
                    "aspect_ratio": "A4",
                    "file_url": "",
                    "generation_status": "not_required"
                }
            },
            "page_2": {
                "page_no": 2,
                "module": "设计背景",
                "title": "设计背景与用户需求",
                "content": {
                    "summary": f"本项目面向{target}，聚焦{scenario}，围绕{function_text}等高频需求设计机器人服务闭环。",
                    "user_pain_points": [
                        "持续照护或重复性服务资源不足，人工成本与时间成本较高。",
                        "传统单一设备彼此割裂，缺少统一的感知、判断和联动能力。",
                        "异常事件发生后，发现、确认、通知和处置链路可能不够及时。",
                        "现有产品的人机交互门槛、主动服务能力和个性化适配仍可提升。"
                    ],
                    "design_objectives": [
                        f"为{target}提供低门槛、自然、清晰的人机交互方式。",
                        f"实现{function_text}等核心功能的感知—决策—执行闭环。",
                        "提高异常状态识别、主动提醒、远程联动与人工确认能力。",
                        "在保证安全和可实现性的前提下突出竞赛展示性与可扩展性。"
                    ],
                    "evidence_requirements": [
                        "正式版本需接入竞赛规则知识库，自动校验参赛方向与格式。",
                        "正式版本需接入真实往届获奖作品，完成差异化对比。",
                        "涉及市场、政策和统计数据时必须保留来源字段。"
                    ]
                },
                "image": image_spec(
                    "用户痛点场景图",
                    "展示目标用户在真实使用场景中的核心痛点",
                    f"面向{target}的{scenario}用户需求场景图，突出{function_text}相关痛点，科技竞赛报告风格，少文字、强信息图表达。"
                )
            },
            "page_3": {
                "page_no": 3,
                "module": "产品整体结构",
                "title": "机器人整体结构设计",
                "content": {
                    "architecture_summary": "系统采用感知层、决策层、执行层、通信层和应用层构成分层机器人架构。",
                    "hardware_structure": [
                        "机器人主体 / 移动底盘或固定终端",
                        "人体状态与环境感知传感器",
                        "语音与视觉交互模块",
                        "边缘计算控制器",
                        "无线通信与物联网模块",
                        "声光提醒、显示与必要执行机构"
                    ],
                    "software_structure": tech_modules,
                    "robot_closed_loop": {
                        "sense": "采集用户、环境及设备状态",
                        "decide": "完成状态识别、风险判断和服务策略选择",
                        "act": "执行提醒、导航、联动、报警或机械动作",
                        "feedback": "将执行结果反馈给用户、家属或管理端"
                    },
                    "data_flow": [
                        "传感器采集原始信息",
                        "边缘端进行数据清洗、融合与初步识别",
                        "决策模块判断需求、风险等级与服务动作",
                        "执行模块完成提醒、设备联动、移动或报警",
                        "结果回传用户端并记录关键运行日志"
                    ],
                    "external_interfaces": [
                        "智能家居设备接口",
                        "家属 / 管理人员移动端接口",
                        "云端数据管理接口",
                        "紧急联系人通知接口"
                    ]
                },
                "image": image_spec(
                    "产品总体结构图",
                    "展示机器人软硬件组成、模块关系与数据流",
                    f"绘制{title}总体结构图，含感知层、决策层、执行层、通信层和应用层，用箭头表示信息流与控制流，专业工程框图风格。"
                )
            },
            "page_4": {
                "page_no": 4,
                "module": "软硬件功能设计",
                "title": "硬件与软件功能设计",
                "content": {
                    "core_functions": core_functions,
                    "function_modules": [
                        {
                            "module_name": function_name,
                            "input": "用户指令、传感器数据或环境状态",
                            "processing": "数据融合、状态识别、规则判断或智能决策",
                            "output": "提醒、反馈、联动、报警、导航或执行动作",
                            "verification": "设计对应的功能测试与异常场景测试"
                        }
                        for function_name in core_functions
                    ],
                    "hardware_design_principles": [
                        "传感器与实际识别任务一一对应，避免堆砌无关硬件。",
                        "核心功能尽可能支持边缘端或离线运行，降低网络依赖。",
                        "安全相关功能设计冗余检测或二次确认。",
                        "结构设计兼顾稳定性、安全性、维护性和目标用户使用习惯。"
                    ],
                    "software_design_principles": [
                        "采用模块化架构，支持功能独立测试、替换和升级。",
                        "对用户身份、图像、音频和健康数据设置权限控制。",
                        "保留个性化参数，如提醒时间、联系人和交互偏好。",
                        "保留运行日志，支持故障定位、效果评估和迭代优化。"
                    ]
                },
                "image": image_spec(
                    "软硬件功能框图",
                    "展示传感器、控制器、算法、通信和执行模块之间的关系",
                    f"绘制{title}软硬件功能框图，突出{function_text}，清楚展示输入、处理、决策、执行和反馈链路。"
                )
            },
            "page_5": {
                "page_no": 5,
                "module": "关键技术",
                "title": "关键技术与实现路线",
                "content": {
                    "key_technologies": tech_modules,
                    "technical_route": [
                        {
                            "technology": technology,
                            "role": "支撑机器人感知、理解、决策、通信或执行",
                            "implementation": "优先选择可在现有硬件和开源框架上实现的方案",
                            "verification": "通过模块测试、场景测试和系统联调验证"
                        }
                        for technology in tech_modules
                    ],
                    "engineering_metrics": [
                        "识别类模块：准确率、召回率、误报率",
                        "交互类模块：响应延迟、成功率、可理解性",
                        "移动类模块：避障成功率、路径完成率、定位误差",
                        "告警类模块：触发延迟、送达率、重复告警率"
                    ],
                    "implementation_focus": [
                        "明确每项算法的输入、输出和运行位置。",
                        "优先采用成熟、可复现、可在现有硬件上运行的技术。",
                        "关键识别模块必须设置量化指标，而不是只写“智能识别”。",
                        "建立网络异常、传感器异常和识别不确定时的降级策略。"
                    ],
                    "risk_control": [
                        "避免将概念性功能写成已经完成实测。",
                        "明确原型验证、竞赛样机和最终产品之间的差异。",
                        "涉及健康、安全和紧急情况时增加人工确认与兜底。",
                        "保护用户音频、图像、定位和健康数据。"
                    ]
                },
                "image": image_spec(
                    "关键技术流程图",
                    "展示从数据采集到决策执行的技术路线",
                    f"生成{title}关键技术流程图，核心技术包括{technology_text}，按采集、预处理、识别、融合、决策、执行、反馈顺序展示。"
                )
            },
            "page_6": {
                "page_no": 6,
                "module": "项目创新点",
                "title": "项目创新点与差异化分析",
                "content": {
                    "innovation_points": [
                        {
                            "name": "机器人服务闭环",
                            "description": f"将{function_text}整合为感知、决策、执行和反馈闭环，而非多个独立功能简单叠加。"
                        },
                        {
                            "name": "主动服务模式",
                            "description": "由被动等待指令升级为主动感知、主动判断、主动提醒与主动联动。"
                        },
                        {
                            "name": "多模态信息融合",
                            "description": "通过多类传感信息交叉验证，降低单一传感器误判对安全功能的影响。"
                        },
                        {
                            "name": "适配目标用户的人机交互",
                            "description": f"围绕{target}的实际使用能力优化语音、界面、提醒方式和异常反馈机制。"
                        },
                        {
                            "name": "可解释竞争力评估",
                            "description": "对创新性、科学性、应用前景和设计表达进行分项评分，并显式加入同质化和可实现性修正。"
                        }
                    ],
                    "differentiation_analysis": {
                        "reference_type": "sample_case_library",
                        "similar_case": scores.get("similar_case", {}),
                        "highest_similarity": scores.get("highest_similarity", 0),
                        "homogeneity_penalty": scores.get("homogeneity_penalty", 0),
                        "differentiation_bonus": scores.get("differentiation_bonus", 0),
                        "current_gap": "当前案例库仅用于流程验证，不能作为正式往届获奖证据。",
                        "next_upgrade": "接入真实往届获奖作品后，输出相似作品、重合点、差异点和可改进方向。"
                    }
                },
                "image": image_spec(
                    "创新点对比图",
                    "对比传统方案、样例相似方案和本项目的差异",
                    f"制作{title}创新点对比信息图，突出多模态感知、主动服务、机器人闭环、联动能力和目标用户适配。"
                )
            },
            "page_7": {
                "page_no": 7,
                "module": "行业应用前景",
                "title": "行业应用前景与落地路径",
                "content": {
                    "application_scenarios": fields.get("scenarios", []),
                    "target_users": fields.get("target_groups", []),
                    "deployment_paths": [
                        "个人 / 家庭场景原型部署",
                        "社区或公共服务节点试点",
                        "机构级辅助服务部署",
                        "与智能设备、物联网平台或行业系统进行接口联动",
                        "通过模块化方案扩展到相邻应用场景"
                    ],
                    "social_value": [
                        "减少高频、重复性基础服务工作压力。",
                        "提高异常事件发现、确认、通知和处置效率。",
                        "提升目标用户在真实场景中的安全感、便利性和自主性。",
                        "形成可扩展的机器人服务入口和数据闭环。"
                    ],
                    "commercialization_path": [
                        "完成核心功能原型与关键指标验证。",
                        "开展目标用户场景测试，记录失败案例并迭代。",
                        "形成基础版、增强版和机构版等模块化配置。",
                        "逐步建设设备管理、数据服务、维护和售后体系。"
                    ],
                    "future_iterations": [
                        "扩展更多传感器或执行器接口。",
                        "提高个性化识别与交互能力。",
                        "增加跨设备协同和多机器人协作能力。",
                        "引入长期运行数据评估与持续优化机制。"
                    ],
                    "evidence_boundary": "市场规模、用户数量、政策支持等具体数字必须在接入真实资料后再填写。"
                },
                "image": image_spec(
                    "行业应用生态图",
                    "展示机器人与用户、家庭/机构、移动端和服务平台之间的关系",
                    f"生成{title}行业应用生态图，中心为机器人，周围连接目标用户、家庭或机构、管理端、移动端、物联网设备和服务平台。"
                )
            }
        },
        "image_plan": {
            "total_images": 6,
            "pages": [2, 3, 4, 5, 6, 7],
            "status": "pending"
        },
        "real_case_upgrade": {
            "status": "pending",
            "required_fields": [
                "award_title",
                "competition_name",
                "award_level",
                "year",
                "source_url",
                "keywords",
                "summary"
            ]
        },
        "word_export": {
            "template_name": "robot_competition_report_v1",
            "page_count": 7,
            "status": "pending",
            "file_url": ""
        },
        "revision_history": [
            {
                "revision": 1,
                "action": "initial_generation",
                "updated_pages": [1, 2, 3, 4, 5, 6, 7],
                "updated_at": now_iso()
            }
        ]
    }

    return report


def format_report_summary(report_json: Dict[str, Any]) -> str:
    title = report_json.get("project_title", "")
    prediction = report_json.get("competition_prediction", {})
    return (
        "已生成严格固定为7页的结构化竞赛报告数据。\n\n"
        f"最终题目：{title}\n"
        f"获奖竞争力预测：{prediction.get('score', 0)} / 100\n\n"
        "报告页序：\n"
        "1. 封面\n"
        "2. 设计背景与用户需求\n"
        "3. 机器人整体结构设计\n"
        "4. 硬件与软件功能设计\n"
        "5. 关键技术与实现路线\n"
        "6. 项目创新点与差异化分析\n"
        "7. 行业应用前景与落地路径\n\n"
        "本版已经强化：机器人闭环、工程指标、风险控制、差异化分析和图示规划。\n"
        "结构化数据保存在 payload_json 中。\n\n"
        "下一阶段：接入真实往届获奖作品知识库，再把样例相似度替换为真实案例对比。"
    )


def detect_page_number(message: str) -> Optional[int]:
    text = normalize_text(message)

    patterns = [
        r"第([1-7])页",
        r"page([1-7])",
        r"页([1-7])"
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return int(match.group(1))

    return None


def detect_intent(message: str, session: Dict[str, Any]) -> str:
    text = normalize_text(message)

    if contains_any(text, ["查看报告结构", "查看结构化数据", "查看json", "查看JSON"]):
        return "view_report_json"

    if (
        contains_any(text, ["修改第", "调整第", "改第", "补充第"])
        and detect_page_number(message)
    ):
        return "modify_page_request"

    if contains_any(text, ["重新生成", "重生成", "再生成", "换一批", "不满意"]):
        return "regenerate_titles"

    if (
        re.fullmatch(r"[123]", text)
        or contains_any(text, ["选1", "选2", "选3", "第一个", "第二个", "第三个"])
    ):
        return "select_title"

    if contains_any(text, ["生成报告", "写报告"]):
        return "generate_report"

    # Word 导出暂不伪装成“已完成”
    if contains_any(text, ["生成word", "导出word", "导出Word", "生成Word"]):
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


def handle_chat(req: ChatRequest) -> ChatResponse:
    # 使用真实 user_id + session_id 做隔离，防止多轮状态丢失或不同会话串线。
    session_key = build_session_key(req)
    session = SESSION_STORE.setdefault(session_key, {})

    user_message = clean_message(req.message)
    intent = detect_intent(user_message, session)

    if intent in ["create_project", "supplement_idea", "regenerate_titles"]:
        if intent == "create_project":
            raw_idea = user_message
            prefix = "已收到你的机器人创意，下面自动生成3个候选题目和初评结果。"

        elif intent == "supplement_idea":
            previous_idea = session.get("raw_idea", "")
            raw_idea = previous_idea + "\n补充要求：" + user_message
            prefix = "已把本次内容作为补充要求合并，下面重新生成候选题目和评分。"

        else:
            raw_idea = session.get("raw_idea", user_message)
            prefix = "已根据当前创意重新生成3个候选题目和评分。"

        fields = extract_idea_fields(raw_idea)
        candidates = generate_candidates(fields)

        session["raw_idea"] = raw_idea
        session["fields"] = fields
        session["candidates"] = candidates
        session["selected_title"] = None
        session["report_json"] = None
        session["report_revision"] = 0
        session["stage"] = "candidates_ready"

        payload = {
            "fields": fields,
            "candidates": candidates,
            "current_stage": "candidates_ready"
        }

        return ChatResponse(
            success=True,
            stage="candidates_ready",
            intent=intent,
            message=format_candidate_message(fields, candidates, prefix),
            payload_json=json.dumps(payload, ensure_ascii=False, indent=2),
            download_url="",
            data=payload,
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
                payload_json="",
                download_url="",
                data={},
                files=[],
                suggested_actions=["重新输入创意"],
                updated_at=now_iso()
            )

        selected = next(
            (candidate for candidate in candidates if candidate["id"] == selected_id),
            None
        )

        if selected is None:
            return ChatResponse(
                success=False,
                stage="selection_failed",
                intent="select_title",
                message="没有找到对应编号的题目。请输入1、2或3。",
                payload_json="",
                download_url="",
                data={},
                files=[],
                suggested_actions=["输入1", "输入2", "输入3"],
                updated_at=now_iso()
            )

        session["selected_title"] = selected
        session["stage"] = "title_confirmed"

        selected_payload = {
            "selected_title": selected,
            "current_stage": "title_confirmed"
        }

        message = (
            "已确认最终竞赛题目：\n\n"
            f"{selected['title']}\n\n"
            f"该题目的获奖竞争力预测为 {selected['scores']['competitiveness_score']} / 100，"
            f"核心技术方向为：{selected['core_tech']}。\n\n"
            "请输入「生成报告」，系统会生成严格固定为7页的结构化竞赛报告数据。"
        )

        return ChatResponse(
            success=True,
            stage="title_confirmed",
            intent="select_title",
            message=message,
            payload_json=json.dumps(selected_payload, ensure_ascii=False, indent=2),
            download_url="",
            data=selected_payload,
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
                message="还没有确认最终题目。请先输入1、2或3选择一个候选题目。",
                payload_json="",
                download_url="",
                data={},
                files=[],
                suggested_actions=["输入1", "输入2", "输入3"],
                updated_at=now_iso()
            )

        report_json = build_report_json(fields, selected)
        report_json_text = json.dumps(report_json, ensure_ascii=False, indent=2)

        session["report_json"] = report_json
        session["report_revision"] = 1
        session["stage"] = "report_json_ready"

        return ChatResponse(
            success=True,
            stage="report_json_ready",
            intent="generate_report",
            message=format_report_summary(report_json),
            payload_json=report_json_text,
            download_url="",
            data={
                "report_json": report_json,
                "current_stage": "report_json_ready",
                "revision": 1
            },
            files=[],
            suggested_actions=[
                "查看报告结构数据",
                "修改指定页面",
                "下一步接入真实获奖案例"
            ],
            updated_at=now_iso()
        )

    if intent == "view_report_json":
        report_json = session.get("report_json")

        if not report_json:
            return ChatResponse(
                success=False,
                stage="report_not_ready",
                intent="view_report_json",
                message="当前还没有报告数据。请先选择题目并输入“生成报告”。",
                payload_json="",
                download_url="",
                data={},
                files=[],
                suggested_actions=["生成报告"],
                updated_at=now_iso()
            )

        report_json_text = json.dumps(report_json, ensure_ascii=False, indent=2)

        return ChatResponse(
            success=True,
            stage="report_json_ready",
            intent="view_report_json",
            message="报告结构化 JSON 已返回到 payload_json 字段。",
            payload_json=report_json_text,
            download_url="",
            data={
                "report_json": report_json,
                "revision": session.get("report_revision", 1)
            },
            files=[],
            suggested_actions=["修改指定页面", "下一步接入真实获奖案例"],
            updated_at=now_iso()
        )

    if intent == "modify_page_request":
        report_json = session.get("report_json")
        page_no = detect_page_number(user_message)

        if not report_json or not page_no:
            return ChatResponse(
                success=False,
                stage="modify_page_failed",
                intent="modify_page_request",
                message="当前还没有可修改的报告，或没有识别到页码。请先生成报告，并使用“修改第3页：……”这样的表达。",
                payload_json="",
                download_url="",
                data={},
                files=[],
                suggested_actions=["生成报告"],
                updated_at=now_iso()
            )

        # 当前版本不调用外部大模型擅自改写正文，避免把用户需求误改坏。
        # 先把修改指令结构化保存，下一阶段接入 LLM 后执行定向改写。
        revision = int(session.get("report_revision", 1)) + 1
        request_item = {
            "page_no": page_no,
            "request": user_message,
            "revision": revision,
            "created_at": now_iso()
        }

        report_json.setdefault("pending_page_edits", []).append(request_item)
        report_json["revision"] = revision
        report_json.setdefault("revision_history", []).append({
            "revision": revision,
            "action": "page_edit_requested",
            "updated_pages": [page_no],
            "request": user_message,
            "updated_at": now_iso()
        })

        session["report_json"] = report_json
        session["report_revision"] = revision
        session["stage"] = "page_edit_queued"

        return ChatResponse(
            success=True,
            stage="page_edit_queued",
            intent="modify_page_request",
            message=(
                f"已记录第{page_no}页的修改要求。\n"
                "当前 v0.4.0 先保证修改指令不丢失；下一阶段接入大模型定向改写后，会只改指定页面，不重生成整份报告。"
            ),
            payload_json=json.dumps(report_json, ensure_ascii=False, indent=2),
            download_url="",
            data={
                "page_no": page_no,
                "revision": revision,
                "request": user_message
            },
            files=[],
            suggested_actions=["继续修改其他页面", "查看报告结构数据"],
            updated_at=now_iso()
        )

    if intent == "word_export_pending":
        report_json = session.get("report_json")

        if not report_json:
            return ChatResponse(
                success=False,
                stage="word_export_failed",
                intent="word_export_pending",
                message="当前还没有报告数据。请先完成选题并输入“生成报告”。",
                payload_json="",
                download_url="",
                data={},
                files=[],
                suggested_actions=["生成报告"],
                updated_at=now_iso()
            )

        return ChatResponse(
            success=True,
            stage="word_export_pending",
            intent="word_export_pending",
            message=(
                "7页结构化报告数据已经准备好，但当前 v0.4.0 尚未伪装成已生成 Word 文件。\n"
                "下一阶段我们会新增真正的 Word 导出接口，读取当前 report_json 并生成可下载 .docx。"
            ),
            payload_json=json.dumps(report_json, ensure_ascii=False, indent=2),
            download_url="",
            data={
                "current_stage": "word_export_pending",
                "report_revision": session.get("report_revision", 1)
            },
            files=[],
            suggested_actions=["下一步开发Word导出", "查看报告结构数据"],
            updated_at=now_iso()
        )

    return ChatResponse(
        success=False,
        stage="unknown",
        intent="unknown",
        message="暂时无法识别你的操作。你可以输入机器人创意、输入1/2/3选择题目、输入“生成报告”，或输入“查看报告结构”。",
        payload_json="",
        download_url="",
        data={},
        files=[],
        suggested_actions=["输入创意", "重新生成", "输入1/2/3", "生成报告"],
        updated_at=now_iso()
    )


@app.get("/health")
def health() -> Dict[str, str]:
    return {
        "status": "ok",
        "service": "robot_competition_orchestrator",
        "version": APP_VERSION
    }


@app.post(
    "/api/v1/robot-competition/chat",
    response_model=ChatResponse
)
def chat(req: ChatRequest) -> ChatResponse:
    return handle_chat(req)
