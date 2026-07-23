from datetime import datetime, timezone
from typing import Any, Dict, List, Optional
import json
import re

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field


app = FastAPI(
    title="Robot Competition Orchestrator",
    version="0.3.2",
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
    version: int = 3
    updated_at: str


SESSION_STORE: Dict[str, Dict[str, Any]] = {}


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
    return re.sub(r"\s+", "", text.strip())


def clean_message(text: str) -> str:
    prefixes = [
        "启动V2中央编排器",
        "启动v2中央编排器",
        "启动机器人竞赛助手",
        "启动达标版竞赛助手",
        "运行机器人竞赛中央编排器"
    ]

    result = text.strip()

    for prefix in prefixes:
        result = result.replace(prefix, "").strip()

    return result if result else text.strip()


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

    functions = []

    for function_name, keys in function_map.items():
        if contains_any(text, keys):
            functions.append(function_name)

    if not functions:
        functions.append("基础人机交互")

    tech_map = {
        "语音识别与自然语言交互": ["语音", "聊天", "对话"],
        "多传感器环境感知": ["传感器", "环境", "温湿度", "烟雾", "空气"],
        "视觉识别": ["摄像头", "视觉", "图像", "识别"],
        "姿态/跌倒检测算法": ["跌倒", "摔倒", "姿态"],
        "物联网与智能家居控制": ["智能家居", "家居", "灯光", "空调", "门锁"],
        "移动底盘与路径规划": ["移动", "导航", "避障", "巡航"],
        "远程通信与报警": ["远程", "报警", "通知", "家属"]
    }

    tech_modules = []

    for module_name, keys in tech_map.items():
        if contains_any(text, keys):
            tech_modules.append(module_name)

    if not tech_modules:
        tech_modules.append("基础传感器采集与人机交互模块")

    keywords = []

    keyword_candidates = [
        "老人",
        "独居",
        "家庭",
        "陪伴",
        "健康",
        "吃药",
        "跌倒",
        "语音",
        "智能家居",
        "导航",
        "识别",
        "报警",
        "家属",
        "儿童",
        "校园",
        "农业",
        "物流",
        "环保"
    ]

    for word in keyword_candidates:
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

    completeness_items = 0

    if fields["target_groups"]:
        completeness_items += 1

    if fields["scenarios"]:
        completeness_items += 1

    if fields["core_functions"]:
        completeness_items += 1

    if fields["tech_modules"]:
        completeness_items += 1

    data_completeness_bonus = completeness_items * 1.5

    differentiation_bonus = 0
    combined_modules = "".join(functions + tech_modules)

    if contains_any(combined_modules, ["多传感器", "姿态", "物联网", "路径规划"]):
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
    technology_summary = "、".join(tech_modules[:2])

    candidates = [
        {
            "id": 1,
            "title": f"“智护星”——面向{target}的{scenario}安全陪伴机器人",
            "positioning": f"突出{function_summary}，强调家庭安全与陪伴服务。",
            "core_tech": technology_summary
        },
        {
            "id": 2,
            "title": f"“安居守护者”——基于多模态感知的{target}健康监护机器人",
            "positioning": "突出健康监测、异常识别和远程通知，适合做成完整竞赛方案。",
            "core_tech": technology_summary
        },
        {
            "id": 3,
            "title": f"“慧联家护”——融合智能家居联动的{target}主动服务机器人",
            "positioning": "突出人机交互、环境感知与智能家居联动，展示性较强。",
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
        "说明：这里的“获奖竞争力预测”不是获奖概率，而是基于固定公式计算的可解释竞争力评分。",
        "当前v0.3.2使用内置样例案例库进行相似度测试；后续接入真实往届获奖作品知识库后，会替换为真实案例对比。",
        "",
        "请选择：输入1、2、3确认题目；输入“重新生成”换一批；或直接补充你的要求。"
    ])

    return "\n".join(lines)


def build_report_json(fields: Dict[str, Any], selected: Dict[str, Any]) -> Dict[str, Any]:
    title = selected["title"]
    scores = selected["scores"]

    target = "、".join(fields.get("target_groups", []))
    scenario = "、".join(fields.get("scenarios", []))
    core_functions = fields.get("core_functions", [])
    tech_modules = fields.get("tech_modules", [])
    keywords = fields.get("keywords", [])

    function_text = "、".join(core_functions)
    technology_text = "、".join(tech_modules)

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

    report = {
        "schema_version": "robot-competition-report-v0.3.2",
        "project_title": title,
        "revision": 1,
        "generated_at": now_iso(),
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
                "subtitle": "智能机器人创意竞赛项目报告",
                "content": {
                    "competition_name": "智能机器人创意竞赛",
                    "project_name": title,
                    "track": "服务机器人赛道（待根据竞赛规则确认）",
                    "keywords": keywords,
                    "team_name": "待填写",
                    "school_name": "待填写",
                    "members": "待填写",
                    "advisor": "待填写",
                    "date": "待填写"
                },
                "image": {
                    "required": False,
                    "image_type": "封面背景",
                    "purpose": "由固定Word模板完成封面设计",
                    "prompt": "",
                    "aspect_ratio": "A4",
                    "file_url": "",
                    "generation_status": "not_required"
                }
            },
            "page_2": {
                "page_no": 2,
                "module": "设计背景",
                "title": "设计背景",
                "content": {
                    "summary": f"本项目面向{target}，重点服务于{scenario}，围绕{function_text}等需求设计智能机器人。",
                    "user_pain_points": [
                        "目标用户可能面临持续照护资源不足的问题。",
                        "传统单一设备之间缺乏统一感知、决策和联动能力。",
                        "异常事件发生后，信息通知和应急响应可能不及时。",
                        "现有产品在适老化交互和长期陪伴方面仍有改进空间。"
                    ],
                    "design_objectives": [
                        f"为{target}提供低门槛、易操作的人机交互方式。",
                        f"实现{function_text}等核心功能的闭环服务。",
                        "提高异常事件识别、主动提醒与远程联动能力。",
                        "保证方案具备可实现性、可扩展性和竞赛展示性。"
                    ],
                    "evidence_requirements": [
                        "正式版本必须从竞赛规则知识库提取参赛约束。",
                        "必须引用真实往届获奖作品进行差异化分析。",
                        "不得虚构市场规模、用户数量或获奖概率。"
                    ]
                },
                "image": image_spec(
                    "用户痛点场景图",
                    "展示目标用户在真实使用场景中的核心痛点",
                    f"绘制一张面向{target}的{scenario}需求场景图，体现{function_text}相关痛点，采用科技竞赛报告风格，结构清晰，不出现大段文字。"
                )
            },
            "page_3": {
                "page_no": 3,
                "module": "产品整体结构",
                "title": "产品整体结构",
                "content": {
                    "architecture_summary": "系统采用感知层、通信层、决策层、执行层和应用层组成的分层架构。",
                    "hardware_structure": [
                        "机器人主体或移动底盘",
                        "环境与人体状态感知传感器",
                        "语音交互模块",
                        "边缘计算控制器",
                        "无线通信与物联网模块",
                        "声光提醒及必要的执行机构"
                    ],
                    "software_structure": tech_modules,
                    "data_flow": [
                        "传感器采集用户与环境信息",
                        "边缘端进行数据清洗和初步识别",
                        "决策模块判断用户需求与异常状态",
                        "执行模块完成提醒、联动或报警",
                        "结果反馈至用户、家属或管理平台"
                    ],
                    "external_interfaces": [
                        "智能家居设备接口",
                        "移动端或家属端接口",
                        "云端数据管理接口",
                        "紧急联系人通知接口"
                    ]
                },
                "image": image_spec(
                    "产品总体结构图",
                    "展示机器人软硬件总体组成及模块连接关系",
                    f"生成{title}的产品总体结构图，包含感知层、决策层、执行层、通信层和应用层，用箭头表示数据流，采用专业工程框图风格。"
                )
            },
            "page_4": {
                "page_no": 4,
                "module": "软硬件功能设计",
                "title": "软硬件功能设计",
                "content": {
                    "core_functions": core_functions,
                    "function_modules": [
                        {
                            "module_name": function_name,
                            "input": "传感器数据或用户指令",
                            "processing": "状态识别、规则判断或智能决策",
                            "output": "提醒、反馈、联动、报警或执行动作"
                        }
                        for function_name in core_functions
                    ],
                    "hardware_design_principles": [
                        "传感器选择应与实际识别任务对应。",
                        "核心功能应尽量支持离线或边缘端运行。",
                        "重要异常事件应设计冗余检测机制。",
                        "结构设计应考虑安全、稳定和适老化使用。"
                    ],
                    "software_design_principles": [
                        "采用模块化软件架构，便于单独测试和升级。",
                        "对用户信息设置权限控制和隐私保护。",
                        "为不同用户保留可配置的提醒和交互参数。",
                        "保留运行日志，支持故障定位和功能评估。"
                    ]
                },
                "image": image_spec(
                    "软硬件功能框图",
                    "展示各传感器、控制器和软件模块之间的关系",
                    f"绘制{title}的软硬件功能框图，突出{function_text}，展示传感器、控制器、通信模块、执行模块及用户端之间的连接。"
                )
            },
            "page_5": {
                "page_no": 5,
                "module": "关键技术",
                "title": "关键技术",
                "content": {
                    "key_technologies": tech_modules,
                    "technical_route": [
                        {
                            "technology": technology,
                            "role": "支撑机器人感知、理解、决策或执行功能",
                            "verification": "通过模块测试、场景测试和系统联调验证"
                        }
                        for technology in tech_modules
                    ],
                    "implementation_focus": [
                        "明确每项算法的输入、输出和运行位置。",
                        "优先采用可以在现有硬件上实现的成熟技术。",
                        "对关键识别模块设置准确率、延迟和稳定性指标。",
                        "建立异常情况和传感器失效时的降级策略。"
                    ],
                    "risk_control": [
                        "避免将概念性功能描述为已经完全实现。",
                        "明确原型验证与最终产品之间的差异。",
                        "对涉及健康和安全的判断增加人工确认机制。",
                        "保护用户音频、图像和健康数据。"
                    ]
                },
                "image": image_spec(
                    "关键技术流程图",
                    "展示从数据采集到决策执行的技术路线",
                    f"生成{title}的关键技术流程图，核心技术包括{technology_text}，按数据采集、预处理、识别、决策、执行、反馈顺序展示。"
                )
            },
            "page_6": {
                "page_no": 6,
                "module": "项目创新点",
                "title": "项目创新点",
                "content": {
                    "innovation_points": [
                        {
                            "name": "多功能闭环融合",
                            "description": f"将{function_text}整合到同一机器人系统中，形成感知、判断、执行和反馈闭环。"
                        },
                        {
                            "name": "主动服务模式",
                            "description": "由被动响应升级为主动识别需求、主动提醒和主动联动。"
                        },
                        {
                            "name": "适老化人机交互",
                            "description": "通过自然语音、大字体、低层级操作和异常自动反馈降低使用门槛。"
                        },
                        {
                            "name": "可解释竞争力分析",
                            "description": "采用创新性、科学性、应用前景和设计表达加权评分，并加入同质化与可实现性惩罚。"
                        }
                    ],
                    "differentiation_analysis": {
                        "similar_case": scores.get("similar_case", {}),
                        "highest_similarity": scores.get("highest_similarity", 0),
                        "homogeneity_penalty": scores.get("homogeneity_penalty", 0),
                        "differentiation_bonus": scores.get("differentiation_bonus", 0),
                        "note": "当前相似度数据来自样例案例库，正式版本必须替换为真实往届获奖作品数据。"
                    }
                },
                "image": image_spec(
                    "创新点对比图",
                    "对比本方案与传统单功能设备之间的差异",
                    f"制作{title}的创新点对比信息图，左侧为传统单功能设备，右侧为本项目的多模态感知、主动服务、智能家居联动和适老化交互。"
                )
            },
            "page_7": {
                "page_no": 7,
                "module": "行业应用前景",
                "title": "行业应用前景",
                "content": {
                    "application_scenarios": fields.get("scenarios", []),
                    "target_users": fields.get("target_groups", []),
                    "deployment_paths": [
                        "家庭独立使用",
                        "社区养老服务站部署",
                        "养老机构辅助照护",
                        "与智能家居厂商联合部署",
                        "与健康管理平台进行接口对接"
                    ],
                    "social_value": [
                        "降低高频、重复性基础照护工作的压力。",
                        "提升异常事件发现和通知效率。",
                        "改善目标用户居家生活的安全感和便利性。",
                        "为社区和家庭提供可扩展的智能照护入口。"
                    ],
                    "commercialization_path": [
                        "首先完成核心功能原型和场景验证。",
                        "根据真实用户反馈迭代硬件与交互设计。",
                        "形成基础版、家庭联动版和机构服务版。",
                        "逐步建设设备管理、数据服务和售后体系。"
                    ],
                    "future_iterations": [
                        "增加更丰富的生命体征检测设备接口。",
                        "优化多用户身份识别和个性化交互能力。",
                        "扩展更多智能家居协议。",
                        "引入长期使用数据评估与持续学习机制。"
                    ]
                },
                "image": image_spec(
                    "行业应用生态图",
                    "展示家庭、社区、养老机构与服务平台的应用生态",
                    f"生成{title}的行业应用生态图，中心为机器人，周围连接家庭、社区、养老机构、家属移动端、智能家居和健康服务平台。"
                )
            }
        },
        "image_plan": {
            "total_images": 6,
            "pages": [2, 3, 4, 5, 6, 7],
            "status": "pending"
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


def detect_intent(message: str, session: Dict[str, Any]) -> str:
    text = normalize_text(message)

    if contains_any(text, ["重新生成", "重生成", "再生成", "换一批", "不满意"]):
        return "regenerate_titles"

    if (
        re.fullmatch(r"[123]", text)
        or contains_any(text, ["选1", "选2", "选3", "第一个", "第二个", "第三个"])
    ):
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


def handle_chat(req: ChatRequest) -> ChatResponse:
    session_key = "chaoxing-robot-competition-demo"
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

        message = (
            "已生成严格固定为7页的结构化报告数据。\n\n"
            f"最终题目：{selected['title']}\n\n"
            "报告页序：\n"
            "1. 封面\n"
            "2. 设计背景\n"
            "3. 产品整体结构\n"
            "4. 软硬件功能设计\n"
            "5. 关键技术\n"
            "6. 项目创新点\n"
            "7. 行业应用前景\n\n"
            "当前阶段已完成：\n"
            "输入创意 → 生成3个候选题目 → 选择最终题目 → 生成七页报告结构数据。\n\n"
            "结构化数据已保存到插件输出字段 payload_json 中。\n"
            "当前展示节点只显示简要说明，不直接展开完整JSON。\n\n"
            "下一步需要新增后端导出接口，读取 payload_json 后生成固定格式Word文件。"
        )

        return ChatResponse(
            success=True,
            stage="report_json_ready",
            intent="generate_report",
            message=message,
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
                "进入Word导出阶段",
                "修改指定页面"
            ],
            updated_at=now_iso()
        )

    return ChatResponse(
        success=False,
        stage="unknown",
        intent="unknown",
        message="暂时无法识别你的操作。你可以输入机器人创意、输入1/2/3选择题目，或输入“重新生成”。",
        payload_json="",
        download_url="",
        data={},
        files=[],
        suggested_actions=["输入创意", "重新生成", "输入1/2/3"],
        updated_at=now_iso()
    )


@app.get("/health")
def health() -> Dict[str, str]:
    return {
        "status": "ok",
        "service": "robot_competition_orchestrator",
        "version": "0.3.2"
    }


@app.post(
    "/api/v1/robot-competition/chat",
    response_model=ChatResponse
)
def chat(req: ChatRequest) -> ChatResponse:
    return handle_chat(req)
