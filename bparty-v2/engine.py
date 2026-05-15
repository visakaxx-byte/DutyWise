from __future__ import annotations

import math
import json
import re
import uuid
import asyncio
import hashlib
import inspect
import time
from copy import copy
from dataclasses import asdict, dataclass, field, replace
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Optional

from openpyxl import load_workbook
from pypdf import PdfReader

from crawler_client import StrictTaxCrawler
from llm_client import LLMClient


APP_DIR = Path(__file__).resolve().parent
LOCAL_TEMPLATE_PATH = APP_DIR / "templates/清关模板.xlsx"
REFERENCE_ROOT = APP_DIR / "reference"
REPLACEMENT_WORKBOOK_PATH = APP_DIR / "docs/海关编码查找.xlsx"
RULES_ROOT = APP_DIR / "rules"
MAX_OUTPUT_ITEMS = 30
BASE_TAX_LIMIT = 0.2
TAX_TOLERANCE_USD = 1.0
TAX_UNDER_TARGET_ALLOWANCE_USD = 100.0
DEFAULT_KG_PER_CTN_MIN = 0.5
DEFAULT_KG_PER_CTN_MAX = 80.0
DEFAULT_KG_PER_PC_MIN = 0.01
DEFAULT_KG_PER_PC_MAX = 50.0
DEFAULT_UNIT_PRICE_MIN = 0.05
DEFAULT_UNIT_PRICE_MAX = 50.0
PRECHECK_MANIFEST_SAMPLE_LIMIT = 12
PRECHECK_REPLACEMENT_SAMPLE_LIMIT = 6
PRECHECK_GREEN_THRESHOLD = 0.7
PRECHECK_YELLOW_THRESHOLD = 0.35

DEFAULT_REFERENCE_STYLE_ROWS = [
    {"中文品名": "花园围栏", "英文品名": "Garden fence", "商品编码": "3926400090", "材质": "Plastic", "用途": "HOME", "单价": 1.3},
    {"中文品名": "塑料罩", "英文品名": "Plastic cover", "商品编码": "3924901050", "材质": "Plastic", "用途": "HOME", "单价": 1},
    {"中文品名": "键鼠套装", "英文品名": "Keyboard and mouse set", "商品编码": "8471602000", "材质": "Plastic", "用途": "HOME", "单价": 2},
    {"中文品名": "窗帘", "英文品名": "Curtains", "商品编码": "6303192120", "材质": "Polyester", "用途": "HOME", "单价": 2},
    {"中文品名": "LED灯", "英文品名": "LED lights", "商品编码": "8539520091", "材质": "Iron/Glass", "用途": "HOME", "单价": 2},
    {"中文品名": "杯架", "英文品名": "Cup holder", "商品编码": "3924104000", "材质": "Plastic", "用途": "HOME", "单价": 0.8},
    {"中文品名": "毛巾", "英文品名": "Towel", "商品编码": "6302991520", "材质": "Nylon", "用途": "HOME", "单价": 0.6},
    {"中文品名": "兵乒球拍", "英文品名": "Table tennis paddle", "商品编码": "9506400000", "材质": "Plastic", "用途": "HOME", "单价": 0.7},
    {"中文品名": "背景板", "英文品名": "Backdrop", "商品编码": "3926400090", "材质": "Acrylic", "用途": "HOME", "单价": 1},
    {"中文品名": "花架", "英文品名": "Flower stand", "商品编码": "8306290000", "材质": "Iron", "用途": "HOME", "单价": 2},
    {"中文品名": "手机支架", "英文品名": "Phone holder", "商品编码": "3926100000", "材质": "Plastic", "用途": "HOME", "单价": 0.6},
    {"中文品名": "充气泵", "英文品名": "Inflatable pump", "商品编码": "8414904190", "材质": "Plastic/Iron", "用途": "HOME", "单价": 3},
    {"中文品名": "键盘", "英文品名": "Keyboard", "商品编码": "8471602000", "材质": "Plastic", "用途": "HOME", "单价": 2},
    {"中文品名": "花园装饰", "英文品名": "Garden decoration", "商品编码": "8306290000", "材质": "Iron", "用途": "HOME", "单价": 1},
    {"中文品名": "戒尺", "英文品名": "Ruler", "商品编码": "9017800000", "材质": "Iron", "用途": "HOME", "单价": 0.7},
    {"中文品名": "卡扣", "英文品名": "Buckle", "商品编码": "8302426000", "材质": "Iron", "用途": "HOME", "单价": 0.5},
    {"中文品名": "装饰雕塑", "英文品名": "Decorative sculpture", "商品编码": "3926400090", "材质": "Plastic", "用途": "HOME", "单价": 1},
    {"中文品名": "塑料花", "英文品名": "Plastic flowers", "商品编码": "6702104000", "材质": "Plastic", "用途": "HOME", "单价": 0.15},
    {"中文品名": "卡片", "英文品名": "Card", "商品编码": "4909004000", "材质": "Paper", "用途": "HOME", "单价": 0.1},
]

HEADERS = (
    "中文品名",
    "英文品名",
    "商品编码",
    "材质",
    "用途",
    "箱数",
    "数量",
    "单位",
    "币制",
    "单价",
    "总价",
    "净重",
    "毛重",
    "原产国",
)

FIELD_TRANSLATION_ACRONYMS = {
    "ABS",
    "CBM",
    "DIY",
    "EVA",
    "LED",
    "OPP",
    "PC",
    "PE",
    "PP",
    "PU",
    "PVC",
    "TPE",
}

MATERIAL_TRANSLATIONS = (
    ("不锈钢", "Stainless steel"),
    ("合成皮革", "Synthetic leather"),
    ("聚酯纤维", "Polyester fiber"),
    ("牛津布", "Oxford cloth"),
    ("无纺布", "Non-woven fabric"),
    ("镀锌板", "Galvanized sheet"),
    ("锌合金", "Zinc alloy"),
    ("矿物质粉", "Mineral powder"),
    ("人造纤维", "Synthetic fiber"),
    ("人造丝", "Rayon"),
    ("亚克力", "Acrylic"),
    ("塑料片", "Plastic sheet"),
    ("塑胶", "Plastic"),
    ("塑料", "Plastic"),
    ("涤纶", "Polyester"),
    ("尼龙", "Nylon"),
    ("金属", "Metal"),
    ("钢铁", "Iron and steel"),
    ("黄铜", "Brass"),
    ("橡胶", "Rubber"),
    ("树脂", "Resin"),
    ("玻璃", "Glass"),
    ("皮革", "Leather"),
    ("纤维", "Fiber"),
    ("纸", "Paper"),
    ("铁", "Iron"),
    ("铜", "Copper"),
    ("木制", "Wood"),
    ("木架", "Wood frame"),
    ("木", "Wood"),
    ("布料", "Fabric"),
    ("布", "Cloth"),
    ("海绵", "Sponge"),
    ("陶瓷", "Ceramic"),
    ("硅胶", "Silicone"),
    ("泡沫", "Foam"),
    ("绒面", "Suede"),
    ("棉布", "Cotton cloth"),
    ("棉", "Cotton"),
    ("鞋面", "Upper"),
    ("鞋底", "Sole"),
    ("人造", "Artificial"),
)

USAGE_TRANSLATIONS = (
    ("高尔夫、电力工程、室内装修、户外广告安装等必备测量工具", "Measurement tool"),
    ("用于焊接配件", "Welding accessories"),
    ("家居窗户装饰", "Window treatment"),
    ("汽车行驶", "Automobile driving"),
    ("厨房用品", "Kitchen supplies"),
    ("宠物用品", "Pet supplies"),
    ("庭院用品", "Garden supplies"),
    ("户外用品", "Outdoor supplies"),
    ("生活用品", "Daily necessities"),
    ("日常使用", "Daily use"),
    ("家居用品", "Home supplies"),
    ("装饰用品", "Decoration supplies"),
    ("园艺作业", "Gardening"),
    ("测量戒指大小", "Ring size measurement"),
    ("测量雨量", "Rain gauge"),
    ("装饰照明", "Decorative lighting"),
    ("礼品包装", "Gift packaging"),
    ("居家生活用", "Home use"),
    ("茶具套装", "Tea set"),
    ("餐椅使用", "Dining chair use"),
    ("过滤土", "Soil filtering"),
    ("刮刀工具", "Caulk remover tool"),
    ("切药器", "Pill cutter"),
    ("输入数据", "Data input"),
    ("挂墙展示", "Wall display"),
    ("球存储", "Ball storage"),
    ("写贺卡", "Greeting card writing"),
    ("花园装饰", "Garden decoration"),
    ("生活日用", "Daily use"),
    ("化妆用", "Makeup"),
    ("穿着服饰", "Clothing"),
    ("化妆", "Makeup"),
    ("穿戴", "Wearing"),
    ("穿着", "Wearing"),
    ("家居", "Home use"),
    ("家用", "Home use"),
    ("户外", "Outdoor use"),
    ("办公", "Office use"),
    ("装饰", "Decoration"),
    ("保护", "Protection"),
    ("娱乐", "Entertainment"),
    ("防虫", "Insect protection"),
    ("喂鸟", "Bird feeding"),
    ("打字", "Typing"),
    ("摆件", "Decoration"),
    ("种菜", "Gardening"),
    ("健身", "Fitness"),
    ("凳子", "Stool"),
    ("收纳", "Storage"),
    ("过滤", "Filtering"),
    ("支撑", "Support"),
    ("饮具", "Drinkware"),
    ("粘贴", "Adhesive use"),
    ("眼扣", "Eyelet"),
    ("穿", "Wearing"),
)

ENGLISH_TERM_NORMALIZATIONS = (
    ("home supplies", "Home supplies"),
    ("home supply", "Home supplies"),
    ("household", "Household"),
    ("houshold", "Household"),
    ("daily necessities", "Daily necessities"),
    ("daily use", "Daily use"),
    ("outdoor supplies", "Outdoor supplies"),
    ("outdoor use", "Outdoor use"),
    ("garden supplies", "Garden supplies"),
    ("kitchen supplies", "Kitchen supplies"),
    ("pet supplies", "Pet supplies"),
    ("automobile driving", "Automobile driving"),
    ("decoration supplies", "Decoration supplies"),
    ("stainless steel", "Stainless steel"),
    ("zinc alloy", "Zinc alloy"),
    ("synthetic leather", "Synthetic leather"),
    ("polyester fiber", "Polyester fiber"),
    ("oxford cloth", "Oxford cloth"),
    ("non-woven fabric", "Non-woven fabric"),
    ("cotton cloth", "Cotton cloth"),
    ("iron and steel", "Iron and steel"),
    ("wood frame", "Wood frame"),
    ("plasticr", "Plastic"),
    ("plastics", "Plastic"),
    ("polyster", "Polyester"),
    ("snylon", "Nylon"),
    ("fibre", "Fiber"),
    ("plastic", "Plastic"),
    ("polyester", "Polyester"),
    ("nylon", "Nylon"),
    ("metal", "Metal"),
    ("iron", "Iron"),
    ("steel", "Steel"),
    ("glass", "Glass"),
    ("paper", "Paper"),
    ("wooden", "Wood"),
    ("wood", "Wood"),
    ("fabric", "Fabric"),
    ("cloth", "Cloth"),
    ("sponge", "Sponge"),
    ("foam", "Foam"),
    ("rubber", "Rubber"),
    ("acrylic", "Acrylic"),
    ("cotton", "Cotton"),
    ("canvas", "Canvas"),
    ("fiber", "Fiber"),
    ("velvet", "Velvet"),
    ("leather", "Leather"),
    ("ceramic", "Ceramic"),
    ("silicone", "Silicone"),
    ("brass", "Brass"),
    ("copper", "Copper"),
    ("resin", "Resin"),
    ("upper", "Upper"),
    ("sole", "Sole"),
    ("suede", "Suede"),
    ("wearing", "Wearing"),
    ("dress", "Wearing"),
    ("decoration", "Decoration"),
    ("decorate", "Decoration"),
    ("filtered", "Filtering"),
    ("filtering", "Filtering"),
    ("makeup", "Makeup"),
    ("massage", "Massage"),
    ("coaster", "Coaster"),
    ("tool", "Tool"),
    ("regular", "General use"),
    ("toy", "Toy"),
)

@dataclass
class ManifestItem:
    row: int
    zh: str
    en: str
    hs: str
    material: str
    usage: str
    ctns: Optional[float]
    qty: Optional[float]
    unit_price: Optional[float]
    declared_value: Optional[float]
    real_weight: Optional[float]
    gross_weight: Optional[float]


@dataclass
class ManifestSummary:
    filename: str
    row_count: int
    total_ctns: float
    total_real_weight: float
    total_declared_value: float
    categories: list[str]
    items: list[ManifestItem] = field(default_factory=list)
    weight_source: str = ""
    weight_evidence: str = ""
    weight_confidence: float = 0.0


@dataclass
class ManifestWeightInfo:
    total_weight_kg: Optional[float]
    source: str = ""
    evidence: str = ""
    confidence: float = 0.0


@dataclass
class BillProduct:
    name: str
    hs_code_hint: str = ""
    evidence: str = ""
    confidence: float = 1.0


@dataclass
class BillInfo:
    filename: str
    raw_text: str
    products: list[str]
    shipper: str = ""
    consignee: str = ""
    shipment_no: str = ""
    eta: str = ""
    cartons: Optional[float] = None
    gross_weight: Optional[float] = None
    cbm: Optional[float] = None
    product_entries: list[BillProduct] = field(default_factory=list)


@dataclass(frozen=True)
class ProcessingOptions:
    target_tax_amount: float
    target_item_count: int
    requested_profile: str = "auto"


@dataclass
class SelectionRules:
    allowed_certifications: list[str] = field(default_factory=lambda: ["Lacey Act", "TSCA"])
    blocked_certifications: list[str] = field(default_factory=list)
    allowed_products: list[str] = field(default_factory=list)
    blocked_products: list[str] = field(default_factory=list)


@dataclass
class ProductCandidate:
    source: str
    source_label: str
    zh: str
    en: str
    hs: str
    material: str
    usage: str
    ctns: float = 0.0
    qty: float = 0.0
    unit_price: float = 0.0
    declared_value: float = 0.0
    real_weight: float = 0.0
    gross_weight: float = 0.0
    source_rows: list[int] = field(default_factory=list)
    tax_data: dict[str, Any] = field(default_factory=dict)
    base_tax_rate: float = 0.0
    effective_tax_rate: float = 0.0
    tax_match_source: str = ""
    certification_texts: list[str] = field(default_factory=list)
    filter_reason: str = ""
    original_hs: str = ""
    query_name: str = ""
    query_material: str = ""
    plausibility_range: Optional["PlausibilityRange"] = None
    plausibility_confidence: float = 0.0
    plausibility_basis: str = ""
    llm_reason: str = ""

    @property
    def score(self) -> float:
        return self.real_weight + self.gross_weight + self.ctns * 8 + self.declared_value / 100 + self.qty / 20


@dataclass(frozen=True)
class PlausibilityRange:
    kg_per_ctn_min: Optional[float] = None
    kg_per_ctn_max: Optional[float] = None
    kg_per_pc_min: Optional[float] = None
    kg_per_pc_max: Optional[float] = None
    unit_price_min: Optional[float] = None
    unit_price_max: Optional[float] = None
    ctns_min: Optional[float] = None
    ctns_max: Optional[float] = None
    qty_per_ctn_min: Optional[float] = None
    qty_per_ctn_max: Optional[float] = None
    source: str = ""


@dataclass(frozen=True)
class RowPlan:
    candidate: ProductCandidate
    ctns: int
    qty: int
    gross_weight: float
    total_value: float
    unit_price: float
    plausibility: PlausibilityRange
    warnings: tuple[str, ...] = ()


QueryCache = dict[str, dict[str, Any]]
ProgressCallback = Callable[[dict[str, Any]], Any]


async def build_clearance(
    manifest_path: str | Path,
    bill_path: str | Path,
    output_dir: str | Path,
    requested_profile: str = "auto",
    target_tax_amount: Optional[float] = None,
    target_item_count: Optional[int] = None,
    query_cache: Optional[QueryCache] = None,
    bill_parser: Optional[LLMClient] = None,
    manifest_parser: Optional[LLMClient] = None,
    llm: Optional[LLMClient] = None,
    progress_callback: Optional[ProgressCallback] = None,
) -> dict[str, Any]:
    options = validate_processing_options(target_tax_amount, target_item_count, requested_profile)
    query_cache = query_cache if query_cache is not None else {"product": {}, "hs": {}, "bill": {}}
    llm_client = llm or bill_parser or manifest_parser or LLMClient()
    llm_parse_cache_before = bool(query_cache.get("manifest") or query_cache.get("bill"))
    manifest = await parse_manifest(manifest_path, manifest_parser or llm_client, query_cache)
    bill = await parse_bill(bill_path, bill_parser or llm_client, query_cache)
    llm_parse_used = bool(query_cache.get("manifest") or query_cache.get("bill"))
    llm_plausibility_used = False
    llm_generation_used = False
    if manifest.total_real_weight <= 0:
        raise RuntimeError("清单 Excel 未识别到有效总重量，不能继续生成")

    flow: list[dict[str, Any]] = [
        {
            "stage": "parse",
            "status": "ok",
            "manifest_rows": manifest.row_count,
            "bill_products": len(bill.products),
            "manifest_total_weight": manifest.total_real_weight,
            "target_tax_amount": options.target_tax_amount,
            "target_item_count": options.target_item_count,
        }
    ]
    await emit_progress(
        progress_callback,
        {
            "stage": "parse",
            "status": "ok",
            "progress": 5,
            "message": "文件解析完成",
            "manifest_rows": manifest.row_count,
            "manifest_total_weight": manifest.total_real_weight,
        },
    )

    rules = load_selection_rules()
    crawler = StrictTaxCrawler()
    manifest_candidates = build_manifest_candidates(manifest)
    qualified_manifest: list[ProductCandidate] = []
    manifest_filtered: list[ProductCandidate] = []
    flow.append(
        {
            "stage": "crawler_manifest_products",
            "status": "skipped",
            "queried_rows": 0,
            "qualified": 0,
            "filtered": 0,
            "message": "当前生成逻辑只使用 Excel 总重量，客户清单品名不作为输出候选池",
        }
    )

    replacement_pool = load_replacement_candidates()
    plausibility_ranges = load_plausibility_ranges()
    if len(bill.products) > options.target_item_count:
        raise RuntimeError(
            f"提单品类 {len(bill.products)} 个超过目标输出行数 {options.target_item_count}，"
            "请提高最终生成条目数"
        )
    bill_required: list[ProductCandidate] = []
    bill_filtered: list[ProductCandidate] = []
    if bill.products:
        await emit_progress(
            progress_callback,
            {
                "stage": "bill_products",
                "status": "running",
                "progress": 58,
                "message": "正在按提单品类查询 HS/税率",
            },
        )
        bill_required, bill_filtered = await qualify_bill_product_candidates(
            crawler,
            bill,
            qualified_manifest,
            replacement_pool,
            rules,
            options,
            query_cache=query_cache,
            llm=llm_client,
        )
        manifest_filtered.extend(bill_filtered)
        flow.append(
            {
                "stage": "bill_products",
                "status": "ok" if len(bill_required) == len(bill.products) else "insufficient",
                "bill_products": len(bill.products),
                "qualified": len(bill_required),
                "filtered": len(bill_filtered),
                "message": "提单品类必须来自提单 HS、品类查询或 LLM 查询扩展的真实 codeflagai 候选",
            }
        )
        if len(bill_required) < len(bill.products):
            missing = [item.zh for item in bill_filtered if item.source == "bill"] or bill.products
            raise RuntimeError(
                "提单品类缺少合格 HS/税率候选，不能套用无关品名；"
                f"请补充 HS 或调整品类: {', '.join(missing)}"
            )

    selected = select_initial_candidates([], bill_required, rules, options.target_item_count)
    replacement_filtered: list[ProductCandidate] = []
    replacement_used = 0
    if len(selected) < options.target_item_count:
        needed = options.target_item_count - len(selected)
        replacements = sorted(replacement_pool, key=lambda item: selection_score(item, rules), reverse=True)
        selected_keys = {candidate_identity(candidate) for candidate in selected}
        replacement_attempts = 0
        for replacement in replacements:
            if len(selected) >= options.target_item_count:
                break
            if candidate_identity(replacement) in selected_keys:
                continue
            if replacement_attempts > 0:
                await asyncio.sleep(crawler.settings.delay)
            replacement_attempts += 1
            await emit_progress(
                progress_callback,
                {
                    "stage": "replacement_products",
                    "status": "running",
                    "progress": min(88, 62 + replacement_attempts),
                    "message": f"正在查询替换候选 {replacement_attempts}",
                    "current": replacement_attempts,
                    "selected": len(selected),
                    "target": options.target_item_count,
                },
            )
            qualified, filtered = await qualify_candidates(
                crawler,
                [replacement],
                rules,
                query_cache=query_cache,
            )
            replacement_filtered.extend(filtered)
            if not qualified:
                continue
            selected.append(qualified[0])
            selected_keys.add(candidate_identity(qualified[0]))
            replacement_used += 1
        flow.append(
            {
                "stage": "replacement_products",
                "status": "ok" if len(selected) >= options.target_item_count else "insufficient",
                "needed": needed,
                "used": replacement_used,
                "filtered": len(replacement_filtered),
            }
        )

    if len(selected) < options.target_item_count:
        raise RuntimeError(
            f"合格品名不足，目标 {options.target_item_count} 行，当前仅 {len(selected)} 行；"
            "请补充常用替换清单或放宽规则"
        )

    await emit_progress(
        progress_callback,
        {
            "stage": "optimize_output",
            "status": "running",
            "progress": 90,
            "message": "正在让 LLM 生成合理草案并执行代码校验",
        },
    )
    selected, llm_plausibility_used = await ensure_candidate_plausibility_ranges(
        selected,
        manifest,
        bill,
        plausibility_ranges,
        llm_client,
        query_cache,
    )
    rows, draft_attempts, draft_feedback = await generate_valid_output_rows_with_llm(
        llm_client,
        selected,
        manifest,
        bill,
        options,
    )
    llm_generation_used = True
    validate_output_rows(rows)
    estimated_tax = round(sum((to_float(row.get("总价")) or 0) * (to_float(row.get("综合税率")) or 0) for row in rows), 2)
    tax_gap = round(estimated_tax - options.target_tax_amount, 2)
    flow.append(
        {
            "stage": "optimize_output",
            "status": "ok",
            "rows": len(rows),
            "estimated_tax": estimated_tax,
            "tax_gap": tax_gap,
            "llm_draft_attempts": draft_attempts,
            "draft_feedback": draft_feedback[-1] if draft_feedback else "",
        }
    )

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"清关_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}.xlsx"
    await emit_progress(
        progress_callback,
        {
            "stage": "write_workbook",
            "status": "running",
            "progress": 96,
            "message": "正在写入 Excel",
        },
    )
    write_workbook(bill, rows, output_path, metadata={})
    await emit_progress(
        progress_callback,
        {
            "stage": "done",
            "status": "ok",
            "progress": 100,
            "message": "处理完成",
        },
    )

    total_value = round(sum(to_float(row.get("总价")) or 0 for row in rows), 2)
    stats = {
        "profile_hint": options.requested_profile or "auto",
        "constraint_status": "passed",
        "weight_source": manifest.weight_source or "manifest_total_weight",
        "weight_evidence": manifest.weight_evidence,
        "weight_confidence": manifest.weight_confidence,
        "manifest_rows": manifest.row_count,
        "input_categories": len(manifest.categories),
        "target_item_count": options.target_item_count,
        "output_rows": len(rows),
        "target_tax_amount": options.target_tax_amount,
        "estimated_tax_amount": estimated_tax,
        "tax_gap": tax_gap,
        "input_ctns": manifest.total_ctns,
        "output_ctns": round(sum(to_float(row.get("箱数")) or 0 for row in rows), 2),
        "input_real_weight": manifest.total_real_weight,
        "output_gross_weight": round(sum(to_float(row.get("毛重")) or 0 for row in rows), 2),
        "bill_gross_weight_ignored": bill.gross_weight,
        "input_declared_value": manifest.total_declared_value,
        "total_value_usd": total_value,
        "plausibility_warnings": sum(1 for row in rows if clean_text(row.get("约束提示"))),
        "bill_products": len(bill.products),
        "qualified_manifest_candidates": len(qualified_manifest),
        "replacement_candidates_used": replacement_used,
        "filtered_candidates": len(manifest_filtered) + len(replacement_filtered),
        "realism_status": "passed",
        "realism_warnings": sum(1 for row in rows if clean_text(row.get("约束提示"))),
        "llm_used": llm_parse_used or llm_plausibility_used or llm_generation_used,
        "llm_parse_used": llm_parse_used,
        "llm_parse_cache_reused": llm_parse_cache_before,
        "llm_plausibility_used": llm_plausibility_used,
        "llm_generation_used": llm_generation_used,
        "llm_draft_attempts": draft_attempts,
        "crawler_used": True,
    }

    return {
        "stats": stats,
        "flow": flow,
        "manifest": manifest_to_public_dict(manifest),
        "bill": bill_to_public_dict(bill),
        "bill_categories": bill.products,
        "output_rows": rows,
        "input_tax_records": {candidate.hs: candidate.tax_data for candidate in qualified_manifest if candidate.hs},
        "output_tax_records": {normalize_hs(row.get("商品编码")): row.get("爬虫品名") for row in rows},
        "filter_summary": summarize_filter_reasons([*manifest_filtered, *replacement_filtered]),
        "output_file": str(output_path),
    }


def validate_processing_options(
    target_tax_amount: Optional[float],
    target_item_count: Optional[int],
    requested_profile: str,
) -> ProcessingOptions:
    tax_amount = to_float(target_tax_amount)
    if tax_amount is None or tax_amount <= 0:
        raise RuntimeError("期望税金必须大于 0")
    try:
        item_count = int(target_item_count) if target_item_count is not None else 0
    except (TypeError, ValueError) as exc:
        raise RuntimeError("最终生成条目数必须是整数") from exc
    if item_count < 1 or item_count > MAX_OUTPUT_ITEMS:
        raise RuntimeError(f"最终生成条目数必须在 1-{MAX_OUTPUT_ITEMS} 之间")
    return ProcessingOptions(
        target_tax_amount=round(tax_amount, 2),
        target_item_count=item_count,
        requested_profile=requested_profile or "auto",
    )


async def build_precheck(
    manifest_path: str | Path,
    bill_path: str | Path,
    requested_profile: str = "auto",
    target_tax_amount: Optional[float] = None,
    target_item_count: Optional[int] = None,
    crawler: Optional[StrictTaxCrawler] = None,
    bill_parser: Optional[LLMClient] = None,
    manifest_parser: Optional[LLMClient] = None,
    llm: Optional[LLMClient] = None,
    query_cache: Optional[QueryCache] = None,
    manifest_sample_limit: int = PRECHECK_MANIFEST_SAMPLE_LIMIT,
    replacement_sample_limit: int = PRECHECK_REPLACEMENT_SAMPLE_LIMIT,
) -> dict[str, Any]:
    options = validate_processing_options(target_tax_amount, target_item_count, requested_profile)
    query_cache = query_cache if query_cache is not None else {"product": {}, "hs": {}, "bill": {}}
    llm_client = llm or bill_parser or manifest_parser or LLMClient()
    manifest = await parse_manifest(manifest_path, manifest_parser or llm_client, query_cache)
    bill = await parse_bill(bill_path, bill_parser or llm_client, query_cache)
    rules = load_selection_rules()
    manifest_candidates: list[ProductCandidate] = []
    replacement_candidates = sorted(load_replacement_candidates(), key=lambda item: selection_score(item, rules), reverse=True)

    early_reasons: list[str] = []
    if manifest.total_real_weight <= 0:
        early_reasons.append("清单 Excel 未识别到有效总重量")
    if len(bill.products) > options.target_item_count:
        early_reasons.append(f"提单品类 {len(bill.products)} 个超过目标输出行数 {options.target_item_count}")
    total_candidate_capacity = len(bill.products) + len(manifest_candidates) + len(replacement_candidates)
    if options.target_item_count > total_candidate_capacity:
        early_reasons.append(
            f"目标条目数 {options.target_item_count} 超过客户候选和替换候选总数 {total_candidate_capacity}"
        )

    sampled_manifest: list[ProductCandidate] = []
    sampled_replacements: list[ProductCandidate] = []
    manifest_filtered: list[ProductCandidate] = []
    replacement_filtered: list[ProductCandidate] = []
    bill_required: list[ProductCandidate] = []
    bill_filtered: list[ProductCandidate] = []
    sample_elapsed = 0.0

    if not early_reasons:
        crawler = crawler or StrictTaxCrawler()
        sampled_manifest_source = manifest_candidates[: max(0, manifest_sample_limit)]
        sampled_replacement_source = replacement_candidates[: max(0, replacement_sample_limit)]
        start = time.perf_counter()
        bill_required, bill_filtered = await qualify_bill_product_candidates(
            crawler,
            bill,
            [],
            replacement_candidates,
            rules,
            options,
            query_cache=query_cache,
            llm=llm_client,
        )
        sampled_manifest, manifest_filtered = await qualify_candidates(
            crawler,
            sampled_manifest_source,
            rules,
            query_cache=query_cache,
        )
        sampled_replacements, replacement_filtered = await qualify_candidates(
            crawler,
            sampled_replacement_source,
            rules,
            query_cache=query_cache,
        )
        sample_elapsed = time.perf_counter() - start

    summary = estimate_precheck_summary(
        options=options,
        manifest=manifest,
        bill=bill,
        manifest_candidate_count=len(manifest_candidates),
        replacement_candidate_count=len(replacement_candidates),
        manifest_sample_count=min(len(manifest_candidates), max(0, manifest_sample_limit)),
        replacement_sample_count=min(len(replacement_candidates), max(0, replacement_sample_limit)),
        sampled_manifest=sampled_manifest,
        sampled_replacements=sampled_replacements,
        manifest_filtered=manifest_filtered,
        replacement_filtered=replacement_filtered,
        bill_required=bill_required,
        bill_filtered=bill_filtered,
        sample_elapsed=sample_elapsed,
        crawler_delay=(crawler.settings.delay if crawler else 0.0),
        early_reasons=early_reasons,
    )
    return {
        "precheck": summary,
        "manifest": manifest_to_public_dict(manifest),
        "bill": bill_to_public_dict(bill),
        "query_cache": query_cache,
        "input": {
            "manifest_path": str(manifest_path),
            "bill_path": str(bill_path),
            "target_tax_amount": options.target_tax_amount,
            "target_item_count": options.target_item_count,
            "requested_profile": options.requested_profile,
        },
    }


def estimate_precheck_summary(
    *,
    options: ProcessingOptions,
    manifest: ManifestSummary,
    bill: BillInfo,
    manifest_candidate_count: int,
    replacement_candidate_count: int,
    manifest_sample_count: int,
    replacement_sample_count: int,
    sampled_manifest: list[ProductCandidate],
    sampled_replacements: list[ProductCandidate],
    manifest_filtered: list[ProductCandidate],
    replacement_filtered: list[ProductCandidate],
    bill_required: list[ProductCandidate],
    bill_filtered: list[ProductCandidate],
    sample_elapsed: float,
    crawler_delay: float,
    early_reasons: list[str],
) -> dict[str, Any]:
    sampled_qualified = [*bill_required, *sampled_manifest, *sampled_replacements]
    sampled_count = len(bill_required) + len(bill_filtered) + manifest_sample_count + replacement_sample_count
    manifest_rate = len(sampled_manifest) / manifest_sample_count if manifest_sample_count else 0.0
    replacement_rate = len(sampled_replacements) / replacement_sample_count if replacement_sample_count else 0.0
    positive_rate = (
        len([candidate for candidate in sampled_qualified if candidate_tax_rate(candidate) > 0]) / len(sampled_qualified)
        if sampled_qualified
        else 0.0
    )

    projected_manifest = math.floor(manifest_candidate_count * manifest_rate * 0.85)
    projected_replacements = math.floor(replacement_candidate_count * replacement_rate * 0.85)
    estimated_usable_items = max(len(sampled_qualified), projected_manifest + projected_replacements)
    estimated_positive_items = math.floor(estimated_usable_items * positive_rate)

    if options.target_item_count <= 0:
        success_probability = 0.0
    else:
        success_probability = min(1.0, estimated_usable_items / options.target_item_count)

    risk_reasons = list(early_reasons)
    missing_bill_products = [candidate.zh for candidate in bill_filtered if candidate.source == "bill"]
    if missing_bill_products:
        risk_reasons.append("提单品类预查询未找到合格 HS/税率候选: " + ", ".join(missing_bill_products))
    if not early_reasons and estimated_usable_items < options.target_item_count:
        risk_reasons.append(f"预计合格品名约 {estimated_usable_items} 个，低于目标 {options.target_item_count} 个")
    if not early_reasons and sampled_count and estimated_positive_items <= 0:
        risk_reasons.append("抽样未发现正综合税率候选，可能无法按期望税金反推申报价")

    if early_reasons or missing_bill_products or success_probability < PRECHECK_YELLOW_THRESHOLD or (sampled_count and estimated_positive_items <= 0):
        status = "red"
        decision = "blocked"
    elif success_probability < PRECHECK_GREEN_THRESHOLD:
        status = "yellow"
        decision = "allowed"
    else:
        status = "green"
        decision = "allowed"

    average_candidate_seconds = sample_elapsed / sampled_count if sampled_count else 0.0
    average_candidate_seconds = max(average_candidate_seconds, crawler_delay)
    expected_replacement_attempts = estimate_replacement_attempts(
        options.target_item_count,
        manifest_candidate_count,
        replacement_candidate_count,
        manifest_rate,
        replacement_rate,
    )
    estimated_query_items = manifest_candidate_count + expected_replacement_attempts
    estimated_seconds = math.ceil(max(1.0, average_candidate_seconds * estimated_query_items * 1.15))
    if estimated_seconds >= 600:
        risk_reasons.append("预计执行超过 10 分钟，正式处理将进入后台任务")

    suggested_item_count = max(1, min(options.target_item_count, estimated_usable_items or len(sampled_qualified) or 1))
    filtered_reasons = summarize_filter_reasons([*bill_filtered, *manifest_filtered, *replacement_filtered])
    return {
        "status": status,
        "decision": decision,
        "can_start": decision == "allowed",
        "success_probability": round(success_probability, 4),
        "success_probability_percent": round(success_probability * 100, 1),
        "estimated_seconds": estimated_seconds,
        "estimated_minutes": round(estimated_seconds / 60, 1),
        "estimated_usable_items": estimated_usable_items,
        "estimated_positive_tax_items": estimated_positive_items,
        "suggested_item_count": suggested_item_count,
        "target_item_count": options.target_item_count,
        "target_tax_amount": options.target_tax_amount,
        "manifest_rows": manifest.row_count,
        "manifest_candidate_count": manifest_candidate_count,
        "replacement_candidate_count": replacement_candidate_count,
        "bill_product_count": len(bill.products),
        "qualified_bill_products": len(bill_required),
        "filtered_bill_products": len(bill_filtered),
        "manifest_sample_count": manifest_sample_count,
        "replacement_sample_count": replacement_sample_count,
        "sampled_qualified_manifest": len(sampled_manifest),
        "sampled_qualified_replacement": len(sampled_replacements),
        "manifest_sample_pass_rate": round(manifest_rate, 4),
        "replacement_sample_pass_rate": round(replacement_rate, 4),
        "sample_elapsed_seconds": round(sample_elapsed, 2),
        "estimated_query_items": estimated_query_items,
        "manifest_total_weight": manifest.total_real_weight,
        "manifest_weight_source": manifest.weight_source,
        "manifest_weight_evidence": manifest.weight_evidence,
        "bill_gross_weight": bill.gross_weight,
        "risk_reasons": risk_reasons,
        "filter_summary": filtered_reasons,
    }


def estimate_replacement_attempts(
    target_item_count: int,
    manifest_candidate_count: int,
    replacement_candidate_count: int,
    manifest_rate: float,
    replacement_rate: float,
) -> int:
    expected_manifest_usable = manifest_candidate_count * manifest_rate * 0.85
    needed = max(0.0, target_item_count - expected_manifest_usable)
    if needed <= 0:
        return 0
    if replacement_rate <= 0:
        return replacement_candidate_count
    return min(replacement_candidate_count, math.ceil(needed / replacement_rate))


async def qualify_bill_product_candidates(
    crawler: StrictTaxCrawler,
    bill: BillInfo,
    qualified_manifest: list[ProductCandidate],
    replacement_candidates: list[ProductCandidate],
    rules: SelectionRules,
    options: ProcessingOptions,
    query_cache: Optional[QueryCache] = None,
    llm: Optional[LLMClient] = None,
) -> tuple[list[ProductCandidate], list[ProductCandidate]]:
    required: list[ProductCandidate] = []
    filtered: list[ProductCandidate] = []
    seen: set[tuple[str, str, str]] = set()
    for product in bill.products:
        rejected_matches = [
            candidate
            for candidate in qualified_manifest
            if candidate_matches_single_bill_product(candidate, product)
            and not candidate_has_product_tax_match(candidate)
        ]
        filtered.extend(
            replace(
                candidate,
                filter_reason=f"提单品类候选未通过品名+材质查询确认，不能用于提单品类: {product}",
            )
            for candidate in rejected_matches
        )
        match = next(
            (
                candidate
                for candidate in qualified_manifest
                if candidate_matches_single_bill_product(candidate, product)
                and candidate_has_product_tax_match(candidate)
            ),
            None,
        )
        if not match:
            bill_entry = bill_product_entry_for(bill, product)
            if bill_entry and bill_entry.hs_code_hint:
                result = await qualify_bill_hs_candidate(
                    crawler,
                    build_bill_product_candidate(bill, bill_entry, options),
                    rules,
                    query_cache=query_cache,
                )
                if result.filter_reason:
                    filtered.append(result)
                else:
                    match = result
        if not match:
            result = await qualify_bill_product_query_candidate(
                crawler,
                bill,
                product,
                product,
                "",
                rules,
                options,
                query_cache=query_cache,
                match_source="bill_product",
            )
            if result.filter_reason:
                filtered.append(result)
            else:
                match = result
        if not match:
            attempts = 0
            for replacement in replacement_candidates:
                if not candidate_matches_single_bill_product(replacement, product):
                    continue
                if attempts > 0:
                    await asyncio.sleep(crawler.settings.delay)
                attempts += 1
                result = await qualify_single_candidate(crawler, replacement, rules, query_cache=query_cache)
                if result.filter_reason:
                    filtered.append(result)
                    continue
                if not candidate_has_product_tax_match(result):
                    filtered.append(
                        replace(
                            result,
                            filter_reason=f"提单品类候选未通过品名+材质查询确认，不能用于提单品类: {product}",
                        )
                    )
                    continue
                match = result
                break
        if not match and llm is not None:
            terms = await generate_bill_product_query_terms(
                llm,
                bill,
                product,
                [candidate.filter_reason for candidate in filtered[-6:] if candidate.filter_reason],
                query_cache,
            )
            attempts = 0
            for term in terms:
                if attempts > 0:
                    await asyncio.sleep(crawler.settings.delay)
                attempts += 1
                result = await qualify_bill_product_query_candidate(
                    crawler,
                    bill,
                    product,
                    term.get("product_name") or product,
                    term.get("material") or "",
                    rules,
                    options,
                    query_cache=query_cache,
                    match_source="llm_query",
                    llm_reason=term.get("reason") or "",
                )
                if result.filter_reason:
                    filtered.append(result)
                    continue
                match = result
                break
        if not match:
            filtered.append(
                ProductCandidate(
                    source="bill",
                    source_label=bill.filename,
                    zh=product,
                    en=product,
                    hs="",
                    material="",
                    usage="",
                    filter_reason=f"提单品类未在合格客户清单或常用替换清单中找到可匹配项: {product}",
                )
            )
            continue
        key = candidate_identity(match)
        if key not in seen:
            seen.add(key)
            required.append(match)
    if len(required) > options.target_item_count:
        raise RuntimeError(f"提单品类 {len(required)} 个超过目标输出行数 {options.target_item_count}")
    return required, filtered


async def qualify_bill_product_query_candidate(
    crawler: StrictTaxCrawler,
    bill: BillInfo,
    product: str,
    query_name: str,
    material: str,
    rules: SelectionRules,
    options: ProcessingOptions,
    *,
    query_cache: Optional[QueryCache] = None,
    match_source: str,
    llm_reason: str = "",
) -> ProductCandidate:
    candidate = build_bill_product_query_candidate(bill, product, query_name, material, options, llm_reason)
    product_reason = product_rule_reason(candidate, rules)
    if product_reason:
        return replace(candidate, filter_reason=product_reason)
    try:
        product_results = await cached_search_product(crawler, query_name or product, material, query_cache)
        selected = select_qualified_tax_data(product_results, rules)
        if selected:
            return attach_tax_data(candidate, selected, match_source)
        return replace(candidate, filter_reason=f"提单品类查询无合格税率/认证结果: {product} -> {query_name}")
    except Exception as exc:
        return replace(candidate, filter_reason=f"提单品类查询失败: {product} -> {query_name}: {exc}")


def build_bill_product_query_candidate(
    bill: BillInfo,
    product: str,
    query_name: str,
    material: str,
    options: ProcessingOptions,
    llm_reason: str = "",
) -> ProductCandidate:
    row_count = max(1, options.target_item_count)
    ctns = max(1.0, (bill.cartons or row_count) / row_count)
    gross_weight = max(1.0, row_count)
    qty = max(1.0, ctns)
    name = clean_text(product)
    query = clean_text(query_name) or name
    return ProductCandidate(
        source="bill" if query == name else "llm_query",
        source_label=bill.filename,
        zh=name,
        en=name.title() if name.isupper() else name,
        hs="",
        material=material or "General",
        usage="HOME",
        ctns=ctns,
        qty=qty,
        unit_price=0.0,
        declared_value=0.0,
        real_weight=gross_weight,
        gross_weight=gross_weight,
        query_name=query,
        query_material=material,
        llm_reason=llm_reason,
    )


def bill_product_entry_for(bill: BillInfo, product: str) -> Optional[BillProduct]:
    product_key = normalize_bill_product_text(product)
    for entry in bill.product_entries:
        if normalize_bill_product_text(entry.name) == product_key:
            return entry
    return None


def build_bill_product_candidate(
    bill: BillInfo,
    entry: BillProduct,
    options: ProcessingOptions,
) -> ProductCandidate:
    row_count = max(1, options.target_item_count)
    ctns = max(1.0, (bill.cartons or row_count) / row_count)
    gross_weight = max(1.0, (bill.gross_weight or row_count) / row_count)
    qty = max(1.0, ctns)
    return ProductCandidate(
        source="bill",
        source_label=bill.filename,
        zh=entry.name,
        en=entry.name.title() if entry.name.isupper() else entry.name,
        hs=entry.hs_code_hint,
        material="General",
        usage="HOME",
        ctns=ctns,
        qty=qty,
        unit_price=0.0,
        declared_value=0.0,
        real_weight=max(0.01, gross_weight - ctns),
        gross_weight=gross_weight,
    )


async def qualify_bill_hs_candidate(
    crawler: StrictTaxCrawler,
    candidate: ProductCandidate,
    rules: SelectionRules,
    query_cache: Optional[QueryCache] = None,
) -> ProductCandidate:
    product_reason = product_rule_reason(candidate, rules)
    if product_reason:
        return replace(candidate, filter_reason=product_reason)
    if not candidate.hs:
        return replace(candidate, filter_reason=f"提单品类缺少 HS hint: {candidate.zh or candidate.en}")

    try:
        hs_results = await cached_search(crawler, candidate.hs, query_cache)
        selected = select_qualified_tax_data(hs_results, rules)
        if selected:
            return attach_tax_data(candidate, selected, "bill_hs")
        return replace(candidate, filter_reason="提单 HS hint 查询无合格税率/认证结果")
    except Exception as exc:
        return replace(candidate, filter_reason=f"提单 HS hint 查询失败: {exc}")


def select_initial_candidates(
    qualified_manifest: list[ProductCandidate],
    bill_required: list[ProductCandidate],
    rules: SelectionRules,
    target_item_count: int,
) -> list[ProductCandidate]:
    selected: list[ProductCandidate] = []
    selected_keys: set[tuple[str, str, str]] = set()
    for candidate in bill_required:
        key = candidate_identity(candidate)
        if key not in selected_keys:
            selected.append(candidate)
            selected_keys.add(key)
    for candidate in sorted(qualified_manifest, key=lambda item: selection_score(item, rules), reverse=True):
        if len(selected) >= target_item_count:
            break
        key = candidate_identity(candidate)
        if key in selected_keys:
            continue
        selected.append(candidate)
        selected_keys.add(key)
    return selected


def candidate_matches_single_bill_product(candidate: ProductCandidate, product: str) -> bool:
    return row_matches_single_bill_product({"中文品名": candidate.zh, "英文品名": candidate.en}, product)


def candidate_has_product_tax_match(candidate: ProductCandidate) -> bool:
    return candidate.tax_match_source in {"product", "bill_hs", "bill_product", "llm_query"} and bool(normalize_hs(candidate.hs))


def load_selection_rules(rules_dir: Path = RULES_ROOT) -> SelectionRules:
    return SelectionRules(
        allowed_certifications=read_rule_lines(rules_dir / "allowed_certifications.txt") or ["Lacey Act", "TSCA"],
        blocked_certifications=read_rule_lines(rules_dir / "blocked_certifications.txt"),
        allowed_products=read_rule_lines(rules_dir / "allowed_products.txt"),
        blocked_products=read_rule_lines(rules_dir / "blocked_products.txt"),
    )


def read_rule_lines(path: Path) -> list[str]:
    if not path.exists():
        return []
    lines: list[str] = []
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        lines.append(line)
    return lines


def build_manifest_candidates(manifest: ManifestSummary) -> list[ProductCandidate]:
    groups: dict[tuple[str, str, str, str, str], ProductCandidate] = {}
    for item in manifest.items:
        key = (item.zh, item.en, item.hs, item.material, item.usage)
        if key not in groups:
            groups[key] = ProductCandidate(
                source="manifest",
                source_label=manifest.filename,
                zh=item.zh,
                en=item.en or item.zh,
                hs=item.hs,
                material=item.material,
                usage=item.usage or "HOME",
            )
        candidate = groups[key]
        candidate.ctns += item.ctns or 0
        candidate.qty += item.qty or 0
        candidate.unit_price = candidate.unit_price or item.unit_price or 0
        candidate.declared_value += item.declared_value or 0
        candidate.real_weight += item.real_weight or item.gross_weight or 0
        candidate.gross_weight += item.gross_weight or item.real_weight or 0
        candidate.source_rows.append(item.row)

    return sorted(groups.values(), key=lambda item: item.score, reverse=True)


async def qualify_candidates(
    crawler: StrictTaxCrawler,
    candidates: list[ProductCandidate],
    rules: SelectionRules,
    query_cache: Optional[QueryCache] = None,
    progress_callback: Optional[ProgressCallback] = None,
    progress_stage: str = "",
    progress_start: float = 0,
    progress_end: float = 100,
) -> tuple[list[ProductCandidate], list[ProductCandidate]]:
    qualified: list[ProductCandidate] = []
    filtered: list[ProductCandidate] = []
    for index, candidate in enumerate(candidates):
        if index > 0:
            await asyncio.sleep(crawler.settings.delay)
        result = await qualify_single_candidate(crawler, candidate, rules, query_cache=query_cache)
        if result.filter_reason:
            filtered.append(result)
        else:
            qualified.append(result)
        if progress_callback and progress_stage:
            progress = progress_start
            if candidates:
                progress = progress_start + (progress_end - progress_start) * ((index + 1) / len(candidates))
            await emit_progress(
                progress_callback,
                {
                    "stage": progress_stage,
                    "status": "running",
                    "progress": round(progress, 2),
                    "message": f"已查询 {index + 1}/{len(candidates)} 个候选",
                    "current": index + 1,
                    "total": len(candidates),
                    "qualified": len(qualified),
                    "filtered": len(filtered),
                },
            )
    return qualified, filtered


async def qualify_single_candidate(
    crawler: StrictTaxCrawler,
    candidate: ProductCandidate,
    rules: SelectionRules,
    query_cache: Optional[QueryCache] = None,
) -> ProductCandidate:
    product_reason = product_rule_reason(candidate, rules)
    if product_reason:
        return replace(candidate, filter_reason=product_reason)

    errors: list[str] = []
    if candidate.source == "replacement" and normalize_hs(candidate.hs):
        try:
            hs_results = await cached_search(crawler, candidate.hs, query_cache)
            selected = select_qualified_tax_data(hs_results, rules, required_hs=candidate.hs)
            if selected:
                return attach_tax_data(candidate, selected, "replacement_hs")
            errors.append("替换清单原始 HTS 查询无合格结果")
        except Exception as exc:
            errors.append(f"替换清单原始 HTS 查询失败: {exc}")

    try:
        product_results = await cached_search_product(crawler, candidate.zh or candidate.en, candidate.material, query_cache)
        selected = select_qualified_tax_data(
            product_results,
            rules,
            required_hs=(candidate.hs if candidate.source == "replacement" else ""),
        )
        if selected:
            return attach_tax_data(candidate, selected, "product")
        errors.append(
            "品名+材质查询无合格税率/认证结果"
            + ("，或返回 HS 与替换清单原始 HS 不一致" if candidate.source == "replacement" and candidate.hs else "")
        )
    except Exception as exc:
        errors.append(f"品名+材质查询失败: {exc}")

    if candidate.hs and candidate.source != "replacement":
        try:
            hs_results = await cached_search(crawler, candidate.hs, query_cache)
            selected = select_qualified_tax_data(hs_results, rules, required_hs=candidate.hs)
            if selected:
                return attach_tax_data(candidate, selected, "hs")
            errors.append("原始 HTS 兜底查询无合格结果")
        except Exception as exc:
            errors.append(f"原始 HTS 兜底查询失败: {exc}")

    return replace(candidate, filter_reason="；".join(errors) or "无合格税率/认证结果")


async def cached_search_product(
    crawler: StrictTaxCrawler,
    product_name: str,
    material: str,
    query_cache: Optional[QueryCache] = None,
) -> dict[str, dict[str, Any]]:
    if query_cache is None:
        return await crawler.search_product(product_name, material)
    section = query_cache.setdefault("product", {})
    key = product_cache_key(product_name, material)
    if key not in section:
        section[key] = await crawler.search_product(product_name, material)
    return section[key]


async def cached_search(
    crawler: StrictTaxCrawler,
    hs_code: str,
    query_cache: Optional[QueryCache] = None,
) -> dict[str, dict[str, Any]]:
    if query_cache is None:
        return await crawler.search(hs_code)
    section = query_cache.setdefault("hs", {})
    key = normalize_hs(hs_code)
    if key not in section:
        section[key] = await crawler.search(hs_code)
    return section[key]


async def generate_bill_product_query_terms(
    llm: LLMClient,
    bill: BillInfo,
    product: str,
    failures: list[str],
    query_cache: Optional[QueryCache] = None,
) -> list[dict[str, str]]:
    cache_key = hashlib.sha256(
        json.dumps(
            {
                "kind": "bill_query_terms",
                "product": product,
                "hs_hint": (bill_product_entry_for(bill, product).hs_code_hint if bill_product_entry_for(bill, product) else ""),
                "failures": failures[-6:],
            },
            ensure_ascii=False,
            sort_keys=True,
        ).encode("utf-8", errors="ignore")
    ).hexdigest()
    if query_cache is not None:
        section = query_cache.setdefault("llm_query_terms", {})
        if cache_key in section and isinstance(section[cache_key], dict):
            return normalize_llm_query_terms(section[cache_key], product)

    payload = await llm.chat_json(build_bill_query_term_messages(bill, product, failures), temperature=0.1)
    terms = normalize_llm_query_terms(payload, product)
    if query_cache is not None:
        query_cache.setdefault("llm_query_terms", {})[cache_key] = payload
    return terms


def build_bill_query_term_messages(bill: BillInfo, product: str, failures: list[str]) -> list[dict[str, str]]:
    context = {
        "bill_product": product,
        "bill_products": bill.products,
        "hs_hint": (bill_product_entry_for(bill, product).hs_code_hint if bill_product_entry_for(bill, product) else ""),
        "evidence": (bill_product_entry_for(bill, product).evidence if bill_product_entry_for(bill, product) else ""),
        "failed_reasons": failures[-6:],
    }
    return [
        {
            "role": "system",
            "content": (
                "你是美国进口商品归类查询词专家。你的任务是把提单货物品类转换成 codeflagai 更容易查到的"
                "语义等价或高度相近查询词。不要改变成无关品类。只返回 JSON object。"
            ),
        },
        {
            "role": "user",
            "content": (
                "请给出 3-5 个查询词，优先英文，其次中英混合。每个查询词需包含 product_name、material、reason。\n"
                "要求：必须和提单品类语义一致或商业上可申报为同类；不能为了低税率改成无关商品。\n"
                "JSON格式：{\"queries\":[{\"product_name\":\"\",\"material\":\"\",\"reason\":\"\"}]}\n"
                f"上下文：{json.dumps(context, ensure_ascii=False, separators=(',', ':'))}"
            ),
        },
    ]


def normalize_llm_query_terms(payload: dict[str, Any], product: str) -> list[dict[str, str]]:
    raw_terms = payload.get("queries") or payload.get("terms") or payload.get("items")
    if not isinstance(raw_terms, list):
        raw_terms = []
    result: list[dict[str, str]] = []
    seen: set[str] = set()
    for raw in raw_terms:
        if not isinstance(raw, dict):
            name = clean_text(raw)
            material = ""
            reason = ""
        else:
            name = clean_text(raw.get("product_name") or raw.get("name") or raw.get("query"))
            material = clean_text(raw.get("material"))
            reason = clean_text(raw.get("reason"))
        if not name:
            continue
        key = product_cache_key(name, material)
        if key in seen:
            continue
        seen.add(key)
        result.append({"product_name": name, "material": material, "reason": reason})
        if len(result) >= 5:
            break
    direct = {"product_name": product, "material": "", "reason": "提单原始品类"}
    direct_key = product_cache_key(direct["product_name"], direct["material"])
    return ([direct] if direct_key not in seen else []) + result


def product_cache_key(product_name: str, material: str = "") -> str:
    return f"{normalize_text(product_name)}|{normalize_text(material)}"


async def emit_progress(callback: Optional[ProgressCallback], payload: dict[str, Any]) -> None:
    if not callback:
        return
    result = callback(payload)
    if inspect.isawaitable(result):
        await result


def product_rule_reason(candidate: ProductCandidate, rules: SelectionRules) -> str:
    name = normalize_text(f"{candidate.zh} {candidate.en}")
    for blocked in rules.blocked_products:
        if normalize_text(blocked) and normalize_text(blocked) in name:
            return f"命中禁用品名库: {blocked}"
    return ""


def product_is_allowed(candidate: ProductCandidate, rules: SelectionRules) -> bool:
    name = normalize_text(f"{candidate.zh} {candidate.en}")
    return any(normalize_text(item) and normalize_text(item) in name for item in rules.allowed_products)


def selection_score(candidate: ProductCandidate, rules: SelectionRules) -> tuple[int, float]:
    return (1 if product_is_allowed(candidate, rules) else 0, candidate.score)


def select_qualified_tax_data(
    candidates: dict[str, dict[str, Any]],
    rules: SelectionRules,
    required_hs: str = "",
) -> Optional[dict[str, Any]]:
    ranked: list[tuple[int, float, dict[str, Any]]] = []
    for data in candidates.values():
        if required_hs and not tax_candidate_matches_required_hs(data, required_hs):
            continue
        reason = tax_filter_reason(data, rules)
        if reason:
            continue
        anti_dumping_penalty = 1 if data.get("anti_dumping") else 0
        ranked.append((anti_dumping_penalty, base_tax_rate(data) or 0, data))
    if not ranked:
        return None
    return sorted(ranked, key=lambda item: (item[0], item[1]))[0][2]


def tax_candidate_matches_required_hs(data: dict[str, Any], required_hs: str) -> bool:
    required = normalize_hs(required_hs)
    returned = normalize_hs(data.get("hs_code_us"))
    if not required or not returned:
        return True
    if returned == required:
        return True
    if len(required) < 10 and returned.startswith(required):
        return True
    if len(returned) < 10 and required.startswith(returned):
        return True
    return False


def tax_filter_reason(data: dict[str, Any], rules: SelectionRules) -> str:
    hs = normalize_hs(data.get("hs_code_us"))
    if len(hs) != 10:
        return "美国 HTS 不是 10 位"
    base_rate = base_tax_rate(data)
    if base_rate is None:
        return "基础税率为空或无法解析"
    if base_rate >= BASE_TAX_LIMIT:
        return f"基础税率 {format_rate(base_rate)} 不小于 20%"
    cert_reason = certification_filter_reason(data.get("certification_texts") or [], rules)
    if cert_reason:
        return cert_reason
    return ""


def attach_tax_data(candidate: ProductCandidate, data: dict[str, Any], match_source: str) -> ProductCandidate:
    hs = normalize_hs(data.get("hs_code_us")) or candidate.hs
    return replace(
        candidate,
        hs=hs,
        tax_data=data,
        base_tax_rate=base_tax_rate(data) or 0,
        effective_tax_rate=effective_tax_rate(data),
        tax_match_source=match_source,
        certification_texts=list(data.get("certification_texts") or []),
        filter_reason="",
    )


def normalize_generated_candidate(candidate: ProductCandidate) -> ProductCandidate:
    return replace(
        candidate,
        source="generated",
        ctns=1,
        qty=20,
        unit_price=1.0,
        declared_value=20.0,
        real_weight=10.0,
        gross_weight=10.0,
        source_rows=[],
    )


async def ensure_candidate_plausibility_ranges(
    candidates: list[ProductCandidate],
    manifest: ManifestSummary,
    bill: BillInfo,
    known_ranges: dict[tuple[str, str, str], PlausibilityRange],
    llm: LLMClient,
    query_cache: Optional[QueryCache] = None,
) -> tuple[list[ProductCandidate], bool]:
    result: list[ProductCandidate] = []
    llm_used = False
    for candidate in candidates:
        known = lookup_plausibility_range(candidate, known_ranges)
        if known and plausibility_range_is_complete(known):
            result.append(replace(candidate, plausibility_range=with_default_plausibility_bounds(known)))
            continue
        if not should_use_llm_for_plausibility(candidate, known):
            derived = derive_candidate_plausibility_range(candidate)
            if plausibility_range_is_complete(derived):
                result.append(replace(candidate, plausibility_range=with_default_plausibility_bounds(derived)))
                continue
        estimated = await estimate_candidate_plausibility_with_llm(llm, candidate, manifest, bill, query_cache)
        llm_used = True
        if not plausibility_range_is_complete(estimated):
            raise RuntimeError(f"LLM 未能给出完整合理范围，不能生成: {candidate.zh}/{candidate.en}")
        result.append(
            replace(
                candidate,
                plausibility_range=with_default_plausibility_bounds(estimated),
                plausibility_confidence=range_confidence(estimated),
                plausibility_basis=estimated.source,
            )
        )
    return result, llm_used


def should_use_llm_for_plausibility(candidate: ProductCandidate, known: Optional[PlausibilityRange]) -> bool:
    if candidate.source in {"bill", "llm_query"}:
        return True
    return known is not None and not plausibility_range_is_complete(known)


def lookup_plausibility_range(
    candidate: ProductCandidate,
    ranges: dict[tuple[str, str, str], PlausibilityRange],
) -> Optional[PlausibilityRange]:
    direct = ranges.get(candidate_identity(candidate))
    if direct:
        return direct
    hs = normalize_hs(candidate.hs)
    name_keys = {normalize_text(candidate.zh), normalize_text(candidate.en)}
    for (zh_key, en_key, hs_key), value in ranges.items():
        if hs and hs_key == hs and (zh_key in name_keys or en_key in name_keys):
            return value
    return None


def derive_candidate_plausibility_range(candidate: ProductCandidate) -> PlausibilityRange:
    gross = candidate.gross_weight or candidate.real_weight or 0
    kg_per_ctn = gross / candidate.ctns if gross and candidate.ctns else None
    kg_per_pc = gross / candidate.qty if gross and candidate.qty else None
    qty_per_ctn = candidate.qty / candidate.ctns if candidate.qty and candidate.ctns else None
    return build_plausibility_range(
        kg_per_ctn=kg_per_ctn,
        kg_per_pc=kg_per_pc,
        unit_price=candidate.unit_price or None,
        ctns=candidate.ctns or None,
        qty_per_ctn=qty_per_ctn,
        source=f"{candidate.source_label} 派生范围",
    )


def plausibility_range_is_complete(value: Optional[PlausibilityRange]) -> bool:
    if value is None:
        return False
    fields = (
        value.kg_per_ctn_min,
        value.kg_per_ctn_max,
        value.kg_per_pc_min,
        value.kg_per_pc_max,
        value.unit_price_min,
        value.unit_price_max,
        value.qty_per_ctn_min,
        value.qty_per_ctn_max,
    )
    return all(item is not None and item > 0 for item in fields)


def range_confidence(value: PlausibilityRange) -> float:
    match = re.search(r"confidence=([0-9.]+)", value.source)
    if not match:
        return 0.0
    parsed = to_float(match.group(1))
    return max(0.0, min(1.0, parsed or 0.0))


async def estimate_candidate_plausibility_with_llm(
    llm: LLMClient,
    candidate: ProductCandidate,
    manifest: ManifestSummary,
    bill: BillInfo,
    query_cache: Optional[QueryCache] = None,
) -> PlausibilityRange:
    payload_context = {
        "candidate": candidate_to_llm_dict(candidate),
        "manifest_total_weight_kg": manifest.total_real_weight,
        "manifest_total_ctns": manifest.total_ctns,
        "bill_products": bill.products,
    }
    cache_key = hashlib.sha256(
        json.dumps({"kind": "plausibility", **payload_context}, ensure_ascii=False, sort_keys=True).encode("utf-8", errors="ignore")
    ).hexdigest()
    if query_cache is not None:
        section = query_cache.setdefault("llm_plausibility", {})
        if cache_key in section and isinstance(section[cache_key], dict):
            return normalize_llm_plausibility_payload(section[cache_key], candidate)

    payload = await llm.chat_json(build_plausibility_messages(payload_context), temperature=0.1)
    plausibility = normalize_llm_plausibility_payload(payload, candidate)
    if query_cache is not None:
        query_cache.setdefault("llm_plausibility", {})[cache_key] = payload
    return plausibility


def build_plausibility_messages(context: dict[str, Any]) -> list[dict[str, str]]:
    return [
        {
            "role": "system",
            "content": (
                "你是美国清关商业发票合理性审核专家。你负责根据品类、材质、HS和常识估算合理的申报范围。"
                "不要为了满足税金而压低单价或单重；只返回 JSON object。"
            ),
        },
        {
            "role": "user",
            "content": (
                "请为该候选估算合理范围，单位为 USD、kg、PCS/CTN。\n"
                "必须返回正数区间，min <= max。电器、机器、家居用品等要符合常识；不能把加湿器、挂烫机估成 1 美元。\n"
                "JSON格式：{\"unit_price_min\":0.0,\"unit_price_max\":0.0,"
                "\"kg_per_pc_min\":0.0,\"kg_per_pc_max\":0.0,"
                "\"qty_per_ctn_min\":0.0,\"qty_per_ctn_max\":0.0,"
                "\"kg_per_ctn_min\":0.0,\"kg_per_ctn_max\":0.0,"
                "\"confidence\":0.0,\"basis\":\"\"}\n"
                f"上下文：{json.dumps(context, ensure_ascii=False, separators=(',', ':'))}"
            ),
        },
    ]


def normalize_llm_plausibility_payload(payload: dict[str, Any], candidate: ProductCandidate) -> PlausibilityRange:
    confidence = to_float(payload.get("confidence"))
    confidence = max(0.0, min(1.0, confidence if confidence is not None else 0.0))
    basis = clean_text(payload.get("basis") or payload.get("reason") or "LLM 合理性估算")
    value = PlausibilityRange(
        kg_per_ctn_min=positive_float(payload.get("kg_per_ctn_min")),
        kg_per_ctn_max=positive_float(payload.get("kg_per_ctn_max")),
        kg_per_pc_min=positive_float(payload.get("kg_per_pc_min")),
        kg_per_pc_max=positive_float(payload.get("kg_per_pc_max")),
        unit_price_min=positive_float(payload.get("unit_price_min")),
        unit_price_max=positive_float(payload.get("unit_price_max")),
        qty_per_ctn_min=positive_float(payload.get("qty_per_ctn_min")),
        qty_per_ctn_max=positive_float(payload.get("qty_per_ctn_max")),
        source=f"LLM合理性估算 confidence={confidence}: {candidate.zh}/{candidate.en}; {basis}",
    )
    return normalize_plausibility_range_bounds(value)


def normalize_plausibility_range_bounds(value: PlausibilityRange) -> PlausibilityRange:
    kg_ctn_min, kg_ctn_max = ordered_optional(value.kg_per_ctn_min, value.kg_per_ctn_max)
    kg_pc_min, kg_pc_max = ordered_optional(value.kg_per_pc_min, value.kg_per_pc_max)
    price_min, price_max = ordered_optional(value.unit_price_min, value.unit_price_max)
    qty_min, qty_max = ordered_optional(value.qty_per_ctn_min, value.qty_per_ctn_max)
    return PlausibilityRange(
        kg_per_ctn_min=kg_ctn_min,
        kg_per_ctn_max=kg_ctn_max,
        kg_per_pc_min=kg_pc_min,
        kg_per_pc_max=kg_pc_max,
        unit_price_min=price_min,
        unit_price_max=price_max,
        ctns_min=value.ctns_min,
        ctns_max=value.ctns_max,
        qty_per_ctn_min=qty_min,
        qty_per_ctn_max=qty_max,
        source=value.source,
    )


def ordered_optional(left: Optional[float], right: Optional[float]) -> tuple[Optional[float], Optional[float]]:
    if left is None or right is None:
        return left, right
    return min(left, right), max(left, right)


def positive_float(value: Any) -> Optional[float]:
    parsed = to_float(value)
    if parsed is None or parsed <= 0:
        return None
    return parsed


def base_tax_rate(data: dict[str, Any]) -> Optional[float]:
    return parse_tax_rate(data.get("tax_rate") or data.get("column_rate_of_duty"))


def effective_tax_rate(data: dict[str, Any]) -> float:
    base = base_tax_rate(data)
    if base is None:
        base = 0
    return base + additional_tax_rate(data)


def additional_tax_rate(data: dict[str, Any]) -> float:
    return parse_non_exempt_additional_tax_rate(
        data.get("additional_tax_rate"),
        data.get("additional_tax_details"),
    )


def parse_non_exempt_additional_tax_rate(value: Any, details: Any = None) -> float:
    if isinstance(details, list):
        total = 0.0
        for item in details:
            if not isinstance(item, dict):
                total += parse_non_exempt_additional_tax_rate(item)
                continue
            label = clean_text(item.get("label") or item.get("name") or item.get("type") or item.get("title"))
            rate_value = item.get("rate") or item.get("value") or item.get("tax_rate") or item.get("additionalTaxRate")
            text = " ".join(clean_text(part) for part in (label, rate_value, item.get("text"), item.get("remark")) if clean_text(part))
            if is_base_tax_detail_text(text):
                continue
            if is_exempt_additional_tax_text(text):
                continue
            total += parse_tax_rate(rate_value if rate_value not in (None, "") else text) or 0
        return total
    if isinstance(details, dict):
        return parse_non_exempt_additional_tax_rate([details])

    text = " ".join(clean_text(part) for part in (value, details) if clean_text(part))
    if not text:
        return 0.0
    total = 0.0
    parts = [part.strip() for part in re.split(r"[;\n,，]+", text) if part.strip()]
    if len(parts) <= 1 and "+" in text:
        parts = [part.strip() for part in text.split("+") if part.strip()]
    for part in parts or [text]:
        if is_base_tax_detail_text(part):
            continue
        if is_exempt_additional_tax_text(part):
            continue
        total += parse_tax_rate(part) or 0
    return total


def is_base_tax_detail_text(value: Any) -> bool:
    text = normalize_text(value)
    compact = text.replace(" ", "")
    if not text:
        return False
    markers = (
        "通用税率",
        "普通税率",
        "基础税率",
        "基本税率",
        "进口税率",
        "generalrate",
        "general",
        "baserate",
        "basicrate",
        "column1",
        "columnone",
        "mfn",
        "normaltrade",
    )
    return any(marker in compact for marker in markers)


def is_exempt_additional_tax_text(value: Any) -> bool:
    text = normalize_text(value)
    compact = text.replace(" ", "")
    if not text:
        return False
    exempt_markers = (
        "条件豁免",
        "可豁免",
        "豁免",
        "排除",
        "exempt",
        "exemption",
        "excluded",
        "exclusion",
        "conditionalexemption",
    )
    non_exempt_markers = (
        "无豁免",
        "不豁免",
        "未豁免",
        "nonexempt",
        "noexemption",
        "without exemption",
        "notexempt",
    )
    if any(marker in compact for marker in non_exempt_markers) or any(marker in text for marker in non_exempt_markers):
        return False
    return any(marker in compact for marker in exempt_markers) or any(marker in text for marker in exempt_markers)


def parse_tax_rate(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        if math.isnan(float(value)):
            return None
        number = float(value)
        return number if abs(number) <= 1 else number / 100

    text = str(value).strip()
    if not text:
        return None
    lowered = text.lower()
    if lowered in {"n/a", "na", "none", "null", "-"}:
        return None
    if lowered in {"free", "免税", "无"}:
        return 0.0

    matches = re.findall(r"-?\d+(?:\.\d+)?", text.replace(",", ""))
    if not matches:
        return None
    numbers = [float(match) for match in matches]
    number = sum(numbers) if "+" in text else numbers[0]
    return number / 100 if "%" in text or abs(number) > 1 else number


def certification_filter_reason(texts: list[str], rules: SelectionRules) -> str:
    cleaned = [clean_text(text) for text in texts if clean_text(text)]
    if not cleaned:
        return ""

    joined = " ".join(cleaned)
    joined_key = normalize_certification_text(joined)
    for blocked in rules.blocked_certifications:
        blocked_key = normalize_certification_text(blocked)
        if blocked_key and blocked_key in joined_key:
            return f"命中禁用认证: {blocked}"

    allowed_keys = [normalize_certification_text(item) for item in rules.allowed_certifications if normalize_certification_text(item)]
    present_allowed = [key for key in allowed_keys if key and key in joined_key]
    regulatory_terms = extract_regulatory_terms(joined)
    unknown_terms = [term for term in regulatory_terms if term not in set(allowed_keys)]
    if unknown_terms:
        return "出现未允许认证/监管: " + ", ".join(sorted(unknown_terms))
    if present_allowed and not unknown_terms:
        return ""
    return "出现未允许认证/监管提示"


def normalize_certification_text(value: Any) -> str:
    text = normalize_text(value)
    text = text.replace("lacey act", "laceyact")
    return text.replace(" ", "")


def extract_regulatory_terms(text: str) -> set[str]:
    normalized = normalize_certification_text(text)
    terms: set[str] = set()
    known = {
        "laceyact": ("laceyact", "lacey"),
        "tsca": ("tsca",),
        "fda": ("fda",),
        "fcc": ("fcc",),
        "cpsc": ("cpsc",),
        "cpsia": ("cpsia",),
        "dot": ("dot",),
        "ul": ("ul",),
        "epa": ("epa",),
        "rohs": ("rohs",),
    }
    for canonical, aliases in known.items():
        if any(alias in normalized for alias in aliases):
            terms.add(canonical)
    for acronym in re.findall(r"\b[A-Z]{2,8}\b", text):
        terms.add(normalize_certification_text(acronym))
    return terms


def format_rate(value: float) -> str:
    return f"{round(value * 100, 4)}%"


def load_replacement_candidates(path: Path = REPLACEMENT_WORKBOOK_PATH) -> list[ProductCandidate]:
    if not path.exists():
        return []
    workbook = load_workbook(path, data_only=True)
    candidates: list[ProductCandidate] = []
    if "常用1" in workbook.sheetnames:
        candidates.extend(load_common_sheet_candidates(workbook["常用1"]))
    if "20260330" in workbook.sheetnames:
        candidates.extend(load_20260330_candidates(workbook["20260330"]))
    return dedupe_candidates(candidates)


def load_plausibility_ranges(path: Path = REPLACEMENT_WORKBOOK_PATH) -> dict[tuple[str, str, str], PlausibilityRange]:
    if not path.exists():
        return {}
    workbook = load_workbook(path, data_only=True)
    ranges: dict[tuple[str, str, str], PlausibilityRange] = {}
    if "常用1" in workbook.sheetnames:
        ranges.update(load_common_sheet_plausibility_ranges(workbook["常用1"]))
    if "20260330" in workbook.sheetnames:
        ranges.update(load_20260330_plausibility_ranges(workbook["20260330"]))
    return ranges


def load_common_sheet_plausibility_ranges(sheet) -> dict[tuple[str, str, str], PlausibilityRange]:
    ranges: dict[tuple[str, str, str], PlausibilityRange] = {}
    for row_idx in range(2, sheet.max_row + 1):
        zh = clean_text(sheet.cell(row_idx, 1).value)
        en = clean_text(sheet.cell(row_idx, 2).value)
        hs = normalize_hs(sheet.cell(row_idx, 4).value)
        if not zh and not en and not hs:
            continue
        total_weight = to_float(sheet.cell(row_idx, 5).value)
        ctns = to_float(sheet.cell(row_idx, 6).value)
        qty = to_float(sheet.cell(row_idx, 7).value)
        unit_price = to_float(sheet.cell(row_idx, 8).value)
        item_weight = to_float(sheet.cell(row_idx, 10).value)
        kg_per_ctn = total_weight / ctns if total_weight and ctns else None
        kg_per_pc = item_weight or (total_weight / qty if total_weight and qty else None)
        qty_per_ctn = qty / ctns if qty and ctns else None
        merge_plausibility_range(
            ranges,
            candidate_key_from_parts(zh, en, hs),
            build_plausibility_range(
                kg_per_ctn=kg_per_ctn,
                kg_per_pc=kg_per_pc,
                unit_price=unit_price,
                ctns=ctns,
                qty_per_ctn=qty_per_ctn,
                source="海关编码查找.xlsx/常用1",
            ),
        )
    return ranges


def load_20260330_plausibility_ranges(sheet) -> dict[tuple[str, str, str], PlausibilityRange]:
    ranges: dict[tuple[str, str, str], PlausibilityRange] = {}
    for row_idx in range(3, sheet.max_row + 1):
        zh = clean_text(sheet.cell(row_idx, 2).value)
        hs = normalize_hs(sheet.cell(row_idx, 5).value)
        if not zh and not hs:
            continue
        unit_price_range = parse_range_bounds(sheet.cell(row_idx, 8).value)
        kg_per_pc_range = parse_range_bounds(sheet.cell(row_idx, 9).value)
        ctns_range = parse_range_bounds(sheet.cell(row_idx, 10).value)
        qty_per_ctn_range = parse_range_bounds(sheet.cell(row_idx, 11).value)
        merge_plausibility_range(
            ranges,
            candidate_key_from_parts(zh, zh, hs),
            PlausibilityRange(
                kg_per_ctn_min=None,
                kg_per_ctn_max=None,
                kg_per_pc_min=kg_per_pc_range[0],
                kg_per_pc_max=kg_per_pc_range[1],
                unit_price_min=unit_price_range[0],
                unit_price_max=unit_price_range[1],
                ctns_min=ctns_range[0],
                ctns_max=ctns_range[1],
                qty_per_ctn_min=qty_per_ctn_range[0],
                qty_per_ctn_max=qty_per_ctn_range[1],
                source="海关编码查找.xlsx/20260330",
            ),
        )
    return ranges


def merge_plausibility_range(
    ranges: dict[tuple[str, str, str], PlausibilityRange],
    key: tuple[str, str, str],
    value: PlausibilityRange,
) -> None:
    existing = ranges.get(key)
    if not existing:
        ranges[key] = value
        return
    ranges[key] = PlausibilityRange(
        kg_per_ctn_min=min_optional(existing.kg_per_ctn_min, value.kg_per_ctn_min),
        kg_per_ctn_max=max_optional(existing.kg_per_ctn_max, value.kg_per_ctn_max),
        kg_per_pc_min=min_optional(existing.kg_per_pc_min, value.kg_per_pc_min),
        kg_per_pc_max=max_optional(existing.kg_per_pc_max, value.kg_per_pc_max),
        unit_price_min=min_optional(existing.unit_price_min, value.unit_price_min),
        unit_price_max=max_optional(existing.unit_price_max, value.unit_price_max),
        ctns_min=min_optional(existing.ctns_min, value.ctns_min),
        ctns_max=max_optional(existing.ctns_max, value.ctns_max),
        qty_per_ctn_min=min_optional(existing.qty_per_ctn_min, value.qty_per_ctn_min),
        qty_per_ctn_max=max_optional(existing.qty_per_ctn_max, value.qty_per_ctn_max),
        source=f"{existing.source}; {value.source}" if value.source not in existing.source else existing.source,
    )


def min_optional(left: Optional[float], right: Optional[float]) -> Optional[float]:
    values = [value for value in (left, right) if value is not None]
    return min(values) if values else None


def max_optional(left: Optional[float], right: Optional[float]) -> Optional[float]:
    values = [value for value in (left, right) if value is not None]
    return max(values) if values else None


def build_plausibility_range(
    *,
    kg_per_ctn: Optional[float],
    kg_per_pc: Optional[float],
    unit_price: Optional[float],
    ctns: Optional[float],
    qty_per_ctn: Optional[float],
    source: str,
) -> PlausibilityRange:
    kg_per_ctn_min, kg_per_ctn_max = spread_range(kg_per_ctn, 0.5, 2.0)
    kg_per_pc_min, kg_per_pc_max = spread_range(kg_per_pc, 0.5, 2.0)
    unit_price_min, unit_price_max = spread_range(unit_price, 0.5, 2.0)
    ctns_min, ctns_max = spread_range(ctns, 0.4, 2.5)
    qty_per_ctn_min, qty_per_ctn_max = spread_range(qty_per_ctn, 0.5, 2.0)
    return PlausibilityRange(
        kg_per_ctn_min=kg_per_ctn_min,
        kg_per_ctn_max=kg_per_ctn_max,
        kg_per_pc_min=kg_per_pc_min,
        kg_per_pc_max=kg_per_pc_max,
        unit_price_min=unit_price_min,
        unit_price_max=unit_price_max,
        ctns_min=ctns_min,
        ctns_max=ctns_max,
        qty_per_ctn_min=qty_per_ctn_min,
        qty_per_ctn_max=qty_per_ctn_max,
        source=source,
    )


def spread_range(value: Optional[float], min_factor: float, max_factor: float) -> tuple[Optional[float], Optional[float]]:
    if value is None or value <= 0:
        return None, None
    return max(0.0001, value * min_factor), max(0.0001, value * max_factor)


def load_common_sheet_candidates(sheet) -> list[ProductCandidate]:
    candidates: list[ProductCandidate] = []
    for row_idx in range(2, sheet.max_row + 1):
        zh = clean_text(sheet.cell(row_idx, 1).value)
        en = clean_text(sheet.cell(row_idx, 2).value)
        material = clean_text(sheet.cell(row_idx, 3).value)
        hs = normalize_hs(sheet.cell(row_idx, 4).value)
        if not zh or not hs:
            continue
        candidates.append(
            ProductCandidate(
                source="replacement",
                source_label="海关编码查找.xlsx/常用1",
                zh=zh,
                en=en or zh,
                hs=hs,
                material=material,
                usage="HOME",
                real_weight=parse_range_mid(sheet.cell(row_idx, 5).value) or parse_range_mid(sheet.cell(row_idx, 10).value) or 1,
                gross_weight=parse_range_mid(sheet.cell(row_idx, 5).value) or parse_range_mid(sheet.cell(row_idx, 10).value) or 1,
                ctns=parse_range_mid(sheet.cell(row_idx, 6).value) or 1,
                qty=parse_range_mid(sheet.cell(row_idx, 7).value) or 1,
                unit_price=parse_range_mid(sheet.cell(row_idx, 8).value) or 1,
                declared_value=parse_range_mid(sheet.cell(row_idx, 9).value) or 0,
            )
        )
    return candidates


def load_20260330_candidates(sheet) -> list[ProductCandidate]:
    candidates: list[ProductCandidate] = []
    for row_idx in range(3, sheet.max_row + 1):
        zh = clean_text(sheet.cell(row_idx, 2).value)
        material = clean_text(sheet.cell(row_idx, 4).value)
        hs = normalize_hs(sheet.cell(row_idx, 5).value)
        if not zh or not hs:
            continue
        unit_price = parse_range_mid(sheet.cell(row_idx, 8).value) or 1
        weight = parse_range_mid(sheet.cell(row_idx, 9).value) or 1
        ctns = parse_range_mid(sheet.cell(row_idx, 10).value) or 1
        qty_per_ctn = parse_range_mid(sheet.cell(row_idx, 11).value) or 10
        candidates.append(
            ProductCandidate(
                source="replacement",
                source_label="海关编码查找.xlsx/20260330",
                zh=zh,
                en=zh,
                hs=hs,
                material=material,
                usage="HOME",
                real_weight=weight,
                gross_weight=weight,
                ctns=ctns,
                qty=max(1, round(ctns * qty_per_ctn)),
                unit_price=unit_price,
                declared_value=round(max(1, round(ctns * qty_per_ctn)) * unit_price, 2),
            )
        )
    return candidates


def parse_range_mid(value: Any) -> Optional[float]:
    direct = to_float(value)
    if direct is not None:
        return direct
    text = str(value or "")
    numbers = [float(item) for item in re.findall(r"\d+(?:\.\d+)?", text)]
    if not numbers:
        return None
    return sum(numbers[:2]) / min(len(numbers), 2)


def parse_range_bounds(value: Any) -> tuple[Optional[float], Optional[float]]:
    direct = to_float(value)
    if direct is not None:
        return direct, direct
    text = str(value or "")
    numbers = [float(item) for item in re.findall(r"\d+(?:\.\d+)?", text)]
    if not numbers:
        return None, None
    if len(numbers) == 1:
        return numbers[0], numbers[0]
    low, high = numbers[0], numbers[1]
    return min(low, high), max(low, high)


def dedupe_candidates(candidates: list[ProductCandidate]) -> list[ProductCandidate]:
    result: list[ProductCandidate] = []
    seen: set[tuple[str, str, str]] = set()
    for candidate in candidates:
        key = candidate_identity(candidate)
        if key in seen:
            continue
        seen.add(key)
        result.append(candidate)
    return result


def candidate_identity(candidate: ProductCandidate) -> tuple[str, str, str]:
    return (normalize_text(candidate.zh), normalize_text(candidate.en), normalize_hs(candidate.hs))


def candidate_key_from_parts(zh: Any, en: Any, hs: Any) -> tuple[str, str, str]:
    return (normalize_text(zh), normalize_text(en), normalize_hs(hs))


def build_output_rows(
    selected: list[ProductCandidate],
    manifest: ManifestSummary,
    bill: BillInfo,
    options: ProcessingOptions,
) -> list[dict[str, Any]]:
    if len(selected) != options.target_item_count:
        raise RuntimeError(f"输出行数必须等于 {options.target_item_count}")
    if not any(candidate_tax_rate(candidate) > 0 for candidate in selected):
        raise RuntimeError("合格候选综合税率均为 0，无法匹配期望税金")
    if manifest.total_real_weight <= 0:
        raise RuntimeError("清单未提取到有效重量，不能按清单重量生成复核草案")

    plausibility_ranges: dict[tuple[str, str, str], PlausibilityRange] = {}
    row_plans = build_plausible_row_plans(
        selected=selected,
        manifest=manifest,
        bill=bill,
        options=options,
        plausibility_ranges=plausibility_ranges,
    )

    rows: list[dict[str, Any]] = []
    for plan in row_plans:
        candidate = plan.candidate
        tax_rate = candidate.effective_tax_rate or candidate.base_tax_rate
        row = {
            "中文品名": candidate.zh,
            "英文品名": candidate.en or candidate.zh,
            "商品编码": hs_cell_value(candidate.hs),
            "材质": translate_material_to_english(candidate.material),
            "用途": translate_usage_to_english(candidate.usage),
            "箱数": plan.ctns,
            "数量": plan.qty,
            "单位": "PCS",
            "币制": "USD",
            "单价": plan.unit_price,
            "总价": plan.total_value,
            "净重": round(max(0.01, plan.gross_weight * 0.92), 2),
            "毛重": plan.gross_weight,
            "原产国": "CN",
            "来源": candidate.source,
            "来源文件": candidate.source_label,
            "基础税率": round(candidate.base_tax_rate, 6),
            "综合税率": round(tax_rate, 6),
            "加征税率": clean_text(candidate.tax_data.get("additional_tax_rate")),
            "预计税金": round(plan.total_value * tax_rate, 2),
            "爬虫匹配HS": candidate.hs,
            "爬虫品名": candidate.tax_data.get("description_cn", ""),
            "认证信息": "; ".join(candidate.certification_texts),
            "重量规则来源": plan.plausibility.source or "默认规则",
            "约束提示": "; ".join(plan.warnings),
            "source_rows": candidate.source_rows,
        }
        rows.append(row)

    normalize_output_language_fields(rows)
    return rows


async def generate_valid_output_rows_with_llm(
    llm: LLMClient,
    selected: list[ProductCandidate],
    manifest: ManifestSummary,
    bill: BillInfo,
    options: ProcessingOptions,
) -> tuple[list[dict[str, Any]], int, list[str]]:
    feedback = ""
    feedback_history: list[str] = []
    last_error = ""
    for attempt in range(1, 4):
        payload = await llm_generate_output_draft(llm, selected, manifest, bill, options, feedback)
        try:
            rows = normalize_llm_output_draft(payload, selected)
            validate_llm_output_rows(rows, selected, manifest, bill, options)
            return rows, attempt, feedback_history
        except RuntimeError as exc:
            last_error = str(exc)
            feedback_history.append(last_error)
            feedback = (
                "上一次草案未通过代码硬校验，请只修正数值和行分配后重新输出 JSON。"
                f"错误：{last_error}。不能放宽税率、认证、总重量、税金区间、行数和合理范围。"
            )
    raise RuntimeError(f"LLM 草案连续不合格: {last_error}")


async def llm_generate_output_draft(
    llm: LLMClient,
    selected: list[ProductCandidate],
    manifest: ManifestSummary,
    bill: BillInfo,
    options: ProcessingOptions,
    feedback: str = "",
) -> dict[str, Any]:
    messages = [
        {
            "role": "system",
            "content": (
                "你是美国清关 Commercial Invoice & Packing List 生成专家。你负责让数量、箱数、毛重、单价看起来像真实清关表。"
                "代码会严格校验税率、认证、总重量、总税金、行数、单价/单重/每箱数量范围。只返回 JSON object。"
            ),
        },
        {
            "role": "user",
            "content": build_output_draft_prompt(selected, manifest, bill, options, feedback),
        },
    ]
    return await llm.chat_json(messages, temperature=0.1)


def build_output_draft_prompt(
    selected: list[ProductCandidate],
    manifest: ManifestSummary,
    bill: BillInfo,
    options: ProcessingOptions,
    feedback: str = "",
) -> str:
    candidates = [candidate_to_llm_dict(candidate) for candidate in selected]
    return (
        "请基于给定候选生成最终清关行草案。\n"
        "硬要求：\n"
        f"1. 输出 rows 数量必须等于 {options.target_item_count}，且每个候选必须输出一行，不得新增/删除/改名/改 HS。\n"
        f"2. 毛重请按品类合理分配；代码会按 Excel 总重量 {manifest.total_real_weight} kg 等比例倒推并强制闭合。\n"
        f"3. 总税金必须大于等于 {max(0.0, options.target_tax_amount - TAX_UNDER_TARGET_ALLOWANCE_USD)} USD，"
        f"且不得超过 {options.target_tax_amount} USD；税金=总价*综合税率。\n"
        f"4. 总箱数建议等于 {manifest.total_ctns}；若为 0 则按候选合理分配。\n"
        "5. 每行单价和每箱数量必须落入 candidate.plausibility_range；毛重可服务于总重量闭合。\n"
        "6. 不要让所有行数量相同，不要让所有行单件重量相同，不要给电器/机器类低到不合理的单价。\n"
        "7. 单价、数量、毛重、箱数要像真实装箱清单，优先使用候选原始参数或合理范围中位数；毛重最终以 Excel 总重量倒推为准。\n"
        f"{'修正反馈：' + feedback if feedback else ''}\n"
        "JSON格式：{\"rows\":[{\"candidate_index\":0,\"箱数\":1,\"数量\":1,\"单价\":1.0,\"毛重\":1.0}]}\n"
        f"提单品类：{json.dumps(bill.products, ensure_ascii=False, separators=(',', ':'))}\n"
        f"候选：{json.dumps(candidates, ensure_ascii=False, separators=(',', ':'))}"
    )


def candidate_to_llm_dict(candidate: ProductCandidate) -> dict[str, Any]:
    plausibility = candidate.plausibility_range
    return {
        "zh": candidate.zh,
        "en": candidate.en,
        "hs": candidate.hs,
        "material": candidate.material,
        "usage": candidate.usage,
        "source": candidate.source,
        "source_label": candidate.source_label,
        "query_name": candidate.query_name,
        "query_material": candidate.query_material,
        "tax_match_source": candidate.tax_match_source,
        "base_tax_rate": candidate.base_tax_rate,
        "effective_tax_rate": candidate_tax_rate(candidate),
        "original_ctns": candidate.ctns,
        "original_qty": candidate.qty,
        "original_unit_price": candidate.unit_price,
        "original_gross_weight": candidate.gross_weight,
        "plausibility_range": asdict(plausibility) if plausibility else None,
        "plausibility_confidence": candidate.plausibility_confidence,
        "plausibility_basis": candidate.plausibility_basis,
        "llm_reason": candidate.llm_reason,
    }


def normalize_llm_output_draft(payload: dict[str, Any], selected: list[ProductCandidate]) -> list[dict[str, Any]]:
    raw_rows = payload.get("rows")
    if not isinstance(raw_rows, list) or not raw_rows:
        raise RuntimeError("LLM 草案未返回 rows")
    rows: list[dict[str, Any]] = []
    for idx, raw in enumerate(raw_rows, start=1):
        if not isinstance(raw, dict):
            raise RuntimeError(f"LLM 草案第 {idx} 行不是 object")
        candidate_index = int(to_float(raw.get("candidate_index")) if to_float(raw.get("candidate_index")) is not None else idx - 1)
        if candidate_index < 0 or candidate_index >= len(selected):
            raise RuntimeError(f"LLM 草案第 {idx} 行 candidate_index 越界: {candidate_index}")
        candidate = selected[candidate_index]
        qty = to_float(raw.get("数量") or raw.get("qty"))
        ctns = to_float(raw.get("箱数") or raw.get("ctns"))
        unit_price = to_float(raw.get("单价") or raw.get("unit_price"))
        gross_weight = to_float(raw.get("毛重") or raw.get("gross_weight"))
        if None in (qty, ctns, unit_price, gross_weight):
            raise RuntimeError(f"LLM 草案第 {idx} 行数值字段不完整")
        tax_rate = candidate_tax_rate(candidate)
        ctns_int = max(1, int(round(ctns)))
        qty_int = max(1, int(round(qty)))
        unit_price = round(unit_price, 4)
        total_value = round(qty_int * unit_price, 2)
        gross_weight = round(gross_weight, 2)
        row = {
            "中文品名": candidate.zh,
            "英文品名": candidate.en or candidate.zh,
            "商品编码": hs_cell_value(candidate.hs),
            "材质": translate_material_to_english(candidate.material),
            "用途": translate_usage_to_english(candidate.usage),
            "箱数": ctns_int,
            "数量": qty_int,
            "单位": "PCS",
            "币制": "USD",
            "单价": unit_price,
            "总价": total_value,
            "净重": round(max(0.01, gross_weight * 0.92), 2),
            "毛重": gross_weight,
            "原产国": "CN",
            "来源": candidate.source,
            "来源文件": candidate.source_label,
            "基础税率": round(candidate.base_tax_rate, 6),
            "综合税率": round(tax_rate, 6),
            "加征税率": clean_text(candidate.tax_data.get("additional_tax_rate")),
            "预计税金": round(total_value * tax_rate, 2),
            "爬虫匹配HS": candidate.hs,
            "爬虫品名": candidate.tax_data.get("description_cn", ""),
            "认证信息": "; ".join(candidate.certification_texts),
            "重量规则来源": (candidate.plausibility_range.source if candidate.plausibility_range else ""),
            "约束提示": "",
            "source_rows": candidate.source_rows,
            "candidate_index": candidate_index,
            "LLM草案毛重": gross_weight,
            "单件重量": round(gross_weight / qty_int, 6) if qty_int else 0,
            "每箱数量": round(qty_int / ctns_int, 6) if ctns_int else 0,
            "单箱重量": round(gross_weight / ctns_int, 6) if ctns_int else 0,
            "合理性来源": (candidate.plausibility_range.source if candidate.plausibility_range else ""),
            "税率来源": candidate.tax_match_source,
            "查询词": candidate.query_name or candidate.zh,
        }
        rows.append(row)
    return rows


def validate_llm_output_rows(
    rows: list[dict[str, Any]],
    selected: list[ProductCandidate],
    manifest: ManifestSummary,
    bill: BillInfo,
    options: ProcessingOptions,
) -> None:
    validate_output_rows(rows)
    if len(rows) != options.target_item_count:
        raise RuntimeError(f"LLM 草案行数 {len(rows)} 不等于目标 {options.target_item_count}")
    seen_indexes = sorted(int(row.get("candidate_index")) for row in rows)
    if seen_indexes != list(range(len(selected))):
        raise RuntimeError("LLM 草案必须一行对应一个候选，不能重复或遗漏候选")
    ensure_bill_products_present(rows, bill.products, selected)

    close_llm_rows_gross_weight(rows, manifest.total_real_weight)
    tax_total = round(sum((to_float(row.get("总价")) or 0) * (to_float(row.get("综合税率")) or 0) for row in rows), 2)
    tax_lower_bound = max(0.0, round(options.target_tax_amount - TAX_UNDER_TARGET_ALLOWANCE_USD, 2))
    if tax_total > options.target_tax_amount or tax_total < tax_lower_bound:
        feasible = estimate_tax_feasible_range(selected)
        raise RuntimeError(
            "合理范围内无法进入目标税金允许区间或 LLM 草案税金未闭合；"
            f"允许区间 {tax_lower_bound}-{options.target_tax_amount}, 当前 {tax_total}, "
            f"可行税金区间约 {feasible[0]}-{feasible[1]}"
        )
    validate_row_plausibility(rows, selected)
    validate_distribution_realism(rows)
    normalize_output_language_fields(rows)
    for row in rows:
        row["预计税金"] = round((to_float(row.get("总价")) or 0) * (to_float(row.get("综合税率")) or 0), 2)


def close_llm_rows_gross_weight(rows: list[dict[str, Any]], target_gross: float) -> None:
    current_values = [max(0.01, to_float(row.get("毛重")) or 0.0) for row in rows]
    scaled = scale_decimal(current_values, target_gross, 2)
    for row, old_gross, gross in zip(rows, current_values, scaled):
        qty = to_float(row.get("数量")) or 1
        ctns = to_float(row.get("箱数")) or 1
        row["LLM草案毛重"] = row.get("LLM草案毛重", old_gross)
        row["毛重"] = gross
        row["净重"] = round(max(0.01, gross * 0.92), 2)
        row["单件重量"] = round(gross / qty, 6) if qty else 0
        row["单箱重量"] = round(gross / ctns, 6) if ctns else 0
        row["毛重闭合调整"] = round(gross - old_gross, 2)
    gross_total = round(sum(to_float(row.get("毛重")) or 0 for row in rows), 2)
    if abs(gross_total - target_gross) > 0.01:
        raise RuntimeError(f"总毛重未闭合: 目标 {target_gross}, 当前 {gross_total}")


def validate_row_plausibility(rows: list[dict[str, Any]], selected: list[ProductCandidate]) -> None:
    for idx, (row, candidate) in enumerate(zip(rows, selected), start=1):
        if normalize_hs(row.get("商品编码")) != normalize_hs(candidate.hs):
            raise RuntimeError(f"第 {idx} 行 HS 被 LLM 改写，禁止: {row.get('商品编码')} != {candidate.hs}")
        if not candidate.tax_match_source:
            raise RuntimeError(f"第 {idx} 行缺少 codeflagai 税率来源: {candidate.zh}")
        plausibility = candidate.plausibility_range
        if not plausibility:
            raise RuntimeError(f"第 {idx} 行缺少合理范围: {candidate.zh}")
        qty = to_float(row.get("数量")) or 0
        ctns = to_float(row.get("箱数")) or 0
        gross = to_float(row.get("毛重")) or 0
        unit_price = to_float(row.get("单价")) or 0
        checks = [
            ("单价", unit_price, plausibility.unit_price_min, plausibility.unit_price_max, True),
            ("单件重量", gross / qty if qty else 0, plausibility.kg_per_pc_min, plausibility.kg_per_pc_max, False),
            ("每箱数量", qty / ctns if ctns else 0, plausibility.qty_per_ctn_min, plausibility.qty_per_ctn_max, True),
            ("单箱重量", gross / ctns if ctns else 0, plausibility.kg_per_ctn_min, plausibility.kg_per_ctn_max, False),
        ]
        warnings: list[str] = []
        for label, value, low, high, hard_limit in checks:
            if low is not None and value < low - 0.0001:
                message = f"{label}低于合理下限 {round(value, 4)} < {round(low, 4)}"
                if hard_limit:
                    raise RuntimeError(f"第 {idx} 行 {message} ({candidate.zh})")
                warnings.append(message)
                continue
            if high is not None and value > high + 0.0001:
                message = f"{label}高于合理上限 {round(value, 4)} > {round(high, 4)}"
                if hard_limit:
                    raise RuntimeError(f"第 {idx} 行 {message} ({candidate.zh})")
                warnings.append(message)
                continue
            if near_bound(value, low, high):
                warnings.append(f"{label}接近边界 {round(value, 4)}")
        if abs(to_float(row.get("毛重闭合调整")) or 0) > 0.01:
            warnings.append(f"毛重按Excel总重量倒推调整 {row.get('毛重闭合调整')}kg")
        row["约束提示"] = "; ".join(warnings)


def validate_distribution_realism(rows: list[dict[str, Any]]) -> None:
    if len(rows) < 4:
        return
    quantities = [to_float(row.get("数量")) or 0 for row in rows]
    kg_per_pc = [round((to_float(row.get("毛重")) or 0) / (to_float(row.get("数量")) or 1), 4) for row in rows]
    if len(set(quantities)) <= 2 and len(rows) >= 8:
        raise RuntimeError("数量分布过于机械，多个品类数量几乎相同")
    if len(set(kg_per_pc)) <= 2 and len(rows) >= 8:
        raise RuntimeError("单件重量分布过于机械，多个品类单体重量几乎相同")


def estimate_tax_feasible_range(selected: list[ProductCandidate]) -> tuple[float, float]:
    min_tax = 0.0
    max_tax = 0.0
    for candidate in selected:
        plausibility = candidate.plausibility_range
        rate = candidate_tax_rate(candidate)
        if not plausibility or rate <= 0:
            continue
        min_qty = max(1, math.ceil((candidate.ctns or 1) * (plausibility.qty_per_ctn_min or 1)))
        max_qty = max(min_qty, math.ceil((candidate.ctns or 1) * (plausibility.qty_per_ctn_max or min_qty)))
        min_tax += min_qty * (plausibility.unit_price_min or 0) * rate
        max_tax += max_qty * (plausibility.unit_price_max or 0) * rate
    return round(min_tax, 2), round(max_tax, 2)


def build_plausible_row_plans(
    *,
    selected: list[ProductCandidate],
    manifest: ManifestSummary,
    bill: BillInfo,
    options: ProcessingOptions,
    plausibility_ranges: dict[tuple[str, str, str], PlausibilityRange],
) -> list[RowPlan]:
    target_gross = round(manifest.total_real_weight, 2)
    target_ctns = manifest.total_ctns or bill.cartons or sum(candidate.ctns for candidate in selected) or options.target_item_count
    target_ctns = max(options.target_item_count, int(round(target_ctns)))
    ctn_values = [candidate.ctns or 1 for candidate in selected]
    ctns = scale_positive_integers(ctn_values, target_ctns)
    ranges = [resolve_plausibility_range(candidate, plausibility_ranges) for candidate in selected]
    weights = allocate_plausible_weights(selected, ctns, ranges, target_gross)
    quantities = [
        choose_plausible_quantity(candidate, row_ctns, gross, plausibility)
        for candidate, row_ctns, gross, plausibility in zip(selected, ctns, weights, ranges)
    ]
    prices = allocate_plausible_prices(selected, quantities, weights, ranges, options.target_tax_amount)

    plans: list[RowPlan] = []
    for candidate, row_ctns, qty, gross, price, plausibility in zip(selected, ctns, quantities, weights, prices, ranges):
        unit_price, total_value = price
        warnings = tuple(build_constraint_warnings(row_ctns, qty, gross, unit_price, plausibility))
        plans.append(
            RowPlan(
                candidate=candidate,
                ctns=row_ctns,
                qty=qty,
                gross_weight=gross,
                total_value=total_value,
                unit_price=unit_price,
                plausibility=plausibility,
                warnings=warnings,
            )
        )
    return plans


def resolve_plausibility_range(
    candidate: ProductCandidate,
    ranges: dict[tuple[str, str, str], PlausibilityRange],
) -> PlausibilityRange:
    if candidate.plausibility_range:
        return with_default_plausibility_bounds(candidate.plausibility_range)

    direct = lookup_plausibility_range(candidate, ranges)
    if direct:
        return with_default_plausibility_bounds(direct)
    derived = derive_candidate_plausibility_range(candidate)
    if not plausibility_range_is_complete(derived):
        raise RuntimeError(f"候选缺少完整合理范围，不能生成: {candidate.zh}/{candidate.en}")
    return with_default_plausibility_bounds(derived)


def with_default_plausibility_bounds(value: PlausibilityRange) -> PlausibilityRange:
    if not plausibility_range_is_complete(value):
        raise RuntimeError(f"合理范围不完整，不能使用默认规则兜底: {value.source or 'unknown'}")
    return PlausibilityRange(
        kg_per_ctn_min=value.kg_per_ctn_min,
        kg_per_ctn_max=value.kg_per_ctn_max,
        kg_per_pc_min=value.kg_per_pc_min,
        kg_per_pc_max=value.kg_per_pc_max,
        unit_price_min=value.unit_price_min,
        unit_price_max=value.unit_price_max,
        ctns_min=value.ctns_min,
        ctns_max=value.ctns_max,
        qty_per_ctn_min=value.qty_per_ctn_min,
        qty_per_ctn_max=value.qty_per_ctn_max,
        source=value.source,
    )


def scale_positive_integers(values: list[float], target_total: float) -> list[int]:
    if not values:
        return []
    target = max(len(values), int(round(target_total)))
    scaled = scale_integer(values, target)
    result = [max(1, value) for value in scaled]
    diff = target - sum(result)
    if diff > 0:
        order = sorted(range(len(result)), key=lambda idx: values[idx], reverse=True)
        for step in range(diff):
            result[order[step % len(order)]] += 1
    elif diff < 0:
        order = sorted(range(len(result)), key=lambda idx: result[idx], reverse=True)
        remaining = -diff
        for idx in order:
            if remaining <= 0:
                break
            reducible = max(0, result[idx] - 1)
            take = min(reducible, remaining)
            result[idx] -= take
            remaining -= take
    return result


def allocate_plausible_weights(
    selected: list[ProductCandidate],
    ctns: list[int],
    ranges: list[PlausibilityRange],
    target_gross: float,
) -> list[float]:
    base_weights = [candidate.gross_weight or candidate.real_weight or 1 for candidate in selected]
    scaled = scale_decimal(base_weights, target_gross, 4)
    mins = [max(0.01, row_ctns * (plausibility.kg_per_ctn_min or DEFAULT_KG_PER_CTN_MIN)) for row_ctns, plausibility in zip(ctns, ranges)]
    maxes = [max(min_weight, row_ctns * (plausibility.kg_per_ctn_max or DEFAULT_KG_PER_CTN_MAX)) for row_ctns, min_weight, plausibility in zip(ctns, mins, ranges)]
    if sum(mins) - target_gross > 0.01:
        raise RuntimeError(
            "优化无解：目标清单重量低于所选类目的合理单箱重量下限；"
            f"清单重量 {target_gross} kg，最低需要 {round(sum(mins), 2)} kg"
        )
    if target_gross - sum(maxes) > 0.01:
        raise RuntimeError(
            "优化无解：目标清单重量超过所选类目的合理单箱重量上限；"
            f"清单重量 {target_gross} kg，最高可承载 {round(sum(maxes), 2)} kg"
        )

    weights = [min(max(value, min_weight), max_weight) for value, min_weight, max_weight in zip(scaled, mins, maxes)]
    adjust_bounded_values(weights, mins, maxes, target_gross, digits=4, label="重量")
    rounded = [round(value, 2) for value in weights]
    diff = round(target_gross - sum(rounded), 2)
    if rounded and diff:
        idx = max(range(len(rounded)), key=lambda item: maxes[item] - rounded[item] if diff > 0 else rounded[item] - mins[item])
        rounded[idx] = round(rounded[idx] + diff, 2)
    if abs(round(sum(rounded) - target_gross, 2)) > 0.01:
        raise RuntimeError("优化无解：输出重量无法精确闭合到清单重量")
    return rounded


def adjust_bounded_values(
    values: list[float],
    mins: list[float],
    maxes: list[float],
    target_total: float,
    *,
    digits: int,
    label: str,
) -> None:
    diff = round(target_total - sum(values), digits)
    if abs(diff) <= 10 ** (-digits):
        return
    if diff > 0:
        capacities = [max_value - value for value, max_value in zip(values, maxes)]
        capacity_total = sum(max(0.0, value) for value in capacities)
        if capacity_total + 10 ** (-digits) < diff:
            raise RuntimeError(f"优化无解：{label}上调空间不足")
        for idx, capacity in sorted(enumerate(capacities), key=lambda item: item[1], reverse=True):
            if diff <= 10 ** (-digits):
                break
            add = min(max(0.0, capacity), diff)
            values[idx] += add
            diff = round(diff - add, digits)
    else:
        need = abs(diff)
        capacities = [value - min_value for value, min_value in zip(values, mins)]
        capacity_total = sum(max(0.0, value) for value in capacities)
        if capacity_total + 10 ** (-digits) < need:
            raise RuntimeError(f"优化无解：{label}下调空间不足")
        for idx, capacity in sorted(enumerate(capacities), key=lambda item: item[1], reverse=True):
            if need <= 10 ** (-digits):
                break
            take = min(max(0.0, capacity), need)
            values[idx] -= take
            need = round(need - take, digits)


def choose_plausible_quantity(
    candidate: ProductCandidate,
    ctns: int,
    gross_weight: float,
    plausibility: PlausibilityRange,
) -> int:
    min_pc = plausibility.kg_per_pc_min or DEFAULT_KG_PER_PC_MIN
    max_pc = plausibility.kg_per_pc_max or DEFAULT_KG_PER_PC_MAX
    min_qty = max(1, math.ceil(gross_weight / max_pc))
    max_qty = max(min_qty, math.floor(gross_weight / min_pc))
    if plausibility.qty_per_ctn_min:
        min_qty = max(min_qty, math.ceil(ctns * plausibility.qty_per_ctn_min))
    if plausibility.qty_per_ctn_max:
        max_qty = min(max_qty, max(1, math.floor(ctns * plausibility.qty_per_ctn_max)))
    if min_qty > max_qty:
        raise RuntimeError(
            "优化无解：单件重量和单箱件数范围无法同时满足；"
            f"{candidate.zh}/{candidate.en} 毛重 {gross_weight} kg，箱数 {ctns}"
        )

    old_ctns = candidate.ctns or ctns or 1
    old_qty = candidate.qty or old_ctns
    qty_per_ctn = old_qty / old_ctns if old_ctns else 1
    desired = max(1, round(ctns * qty_per_ctn))
    return int(min(max(desired, min_qty), max_qty))


def allocate_plausible_prices(
    selected: list[ProductCandidate],
    quantities: list[int],
    weights: list[float],
    ranges: list[PlausibilityRange],
    target_tax_amount: float,
) -> list[tuple[float, float]]:
    positive_weight = sum(weight for candidate, weight in zip(selected, weights) if candidate_tax_rate(candidate) > 0)
    unit_prices: list[float] = []
    mins: list[float] = []
    maxes: list[float] = []
    for candidate, qty, weight, plausibility in zip(selected, quantities, weights, ranges):
        min_price = plausibility.unit_price_min or DEFAULT_UNIT_PRICE_MIN
        max_price = max(min_price, plausibility.unit_price_max or DEFAULT_UNIT_PRICE_MAX)
        mins.append(min_price)
        maxes.append(max_price)
        tax_rate = candidate_tax_rate(candidate)
        if tax_rate > 0 and positive_weight > 0:
            target_tax = target_tax_amount * weight / positive_weight
            desired = target_tax / tax_rate / qty
        else:
            desired = candidate.unit_price or min_price
        unit_prices.append(min(max(desired, min_price), max_price))

    adjust_price_gap(unit_prices, mins, maxes, selected, quantities, target_tax_amount)
    prices: list[tuple[float, float]] = []
    for unit_price, qty in zip(unit_prices, quantities):
        rounded_unit = round(unit_price, 4)
        total_value = round(rounded_unit * qty, 2)
        prices.append((rounded_unit, total_value))

    estimated_tax = round(sum(total * candidate_tax_rate(candidate) for candidate, (_, total) in zip(selected, prices)), 2)
    tolerance = max(TAX_TOLERANCE_USD, target_tax_amount * 0.01)
    if abs(estimated_tax - target_tax_amount) > tolerance:
        raise RuntimeError(
            "优化无解：在重量和单价常理范围内无法贴近期望税金；"
            f"目标 {target_tax_amount}，可行预计 {estimated_tax}，差额 {round(estimated_tax - target_tax_amount, 2)}"
        )
    return prices


def adjust_price_gap(
    unit_prices: list[float],
    mins: list[float],
    maxes: list[float],
    selected: list[ProductCandidate],
    quantities: list[int],
    target_tax_amount: float,
) -> None:
    current_tax = sum(price * qty * candidate_tax_rate(candidate) for price, qty, candidate in zip(unit_prices, quantities, selected))
    diff = target_tax_amount - current_tax
    if abs(diff) <= max(TAX_TOLERANCE_USD, target_tax_amount * 0.01):
        return
    if diff > 0:
        order = sorted(range(len(unit_prices)), key=lambda idx: (maxes[idx] - unit_prices[idx]) * quantities[idx] * candidate_tax_rate(selected[idx]), reverse=True)
        for idx in order:
            rate = candidate_tax_rate(selected[idx])
            if rate <= 0:
                continue
            capacity_tax = (maxes[idx] - unit_prices[idx]) * quantities[idx] * rate
            if capacity_tax <= 0:
                continue
            add_tax = min(diff, capacity_tax)
            unit_prices[idx] += add_tax / (quantities[idx] * rate)
            diff -= add_tax
            if abs(diff) <= 0.0001:
                break
    else:
        need = abs(diff)
        order = sorted(range(len(unit_prices)), key=lambda idx: (unit_prices[idx] - mins[idx]) * quantities[idx] * candidate_tax_rate(selected[idx]), reverse=True)
        for idx in order:
            rate = candidate_tax_rate(selected[idx])
            if rate <= 0:
                continue
            capacity_tax = (unit_prices[idx] - mins[idx]) * quantities[idx] * rate
            if capacity_tax <= 0:
                continue
            take_tax = min(need, capacity_tax)
            unit_prices[idx] -= take_tax / (quantities[idx] * rate)
            need -= take_tax
            if need <= 0.0001:
                break


def candidate_tax_rate(candidate: ProductCandidate) -> float:
    return candidate.effective_tax_rate or candidate.base_tax_rate or 0.0


def build_constraint_warnings(
    ctns: int,
    qty: int,
    gross_weight: float,
    unit_price: float,
    plausibility: PlausibilityRange,
) -> list[str]:
    warnings: list[str] = []
    kg_per_ctn = gross_weight / ctns if ctns else 0
    kg_per_pc = gross_weight / qty if qty else 0
    if near_bound(kg_per_ctn, plausibility.kg_per_ctn_min, plausibility.kg_per_ctn_max):
        warnings.append(f"单箱重量接近边界 {round(kg_per_ctn, 3)}kg/ctn")
    if near_bound(kg_per_pc, plausibility.kg_per_pc_min, plausibility.kg_per_pc_max):
        warnings.append(f"单件重量接近边界 {round(kg_per_pc, 3)}kg/pc")
    if near_bound(unit_price, plausibility.unit_price_min, plausibility.unit_price_max):
        warnings.append(f"单价接近边界 {round(unit_price, 4)}")
    return warnings


def near_bound(value: float, low: Optional[float], high: Optional[float]) -> bool:
    if low is not None and value <= low * 1.05:
        return True
    if high is not None and value >= high * 0.95:
        return True
    return False


def adjust_tax_gap(rows: list[dict[str, Any]], target_tax_amount: float) -> None:
    current = sum((to_float(row.get("总价")) or 0) * (to_float(row.get("综合税率")) or 0) for row in rows)
    diff = round(target_tax_amount - current, 2)
    tolerance = max(TAX_TOLERANCE_USD, target_tax_amount * 0.01)
    if abs(diff) <= tolerance:
        for row in rows:
            row["预计税金"] = round((to_float(row.get("总价")) or 0) * (to_float(row.get("综合税率")) or 0), 2)
        return
    adjustable = [row for row in rows if (to_float(row.get("综合税率")) or 0) > 0]
    if not adjustable:
        raise RuntimeError("没有可调整税金的正税率行")
    row = adjustable[-1]
    rate = to_float(row.get("综合税率")) or 0
    total_value = max(0.01, (to_float(row.get("总价")) or 0) + diff / rate)
    row["总价"] = round(total_value, 2)
    qty = to_float(row.get("数量")) or 1
    row["单价"] = round(row["总价"] / qty, 4)
    for item in rows:
        item["预计税金"] = round((to_float(item.get("总价")) or 0) * (to_float(item.get("综合税率")) or 0), 2)


def summarize_filter_reasons(candidates: list[ProductCandidate]) -> dict[str, int]:
    summary: dict[str, int] = {}
    for candidate in candidates:
        reason = candidate.filter_reason or "unknown"
        first = reason.split("；", 1)[0]
        summary[first] = summary.get(first, 0) + 1
    return summary


async def parse_manifest_total_weight(
    workbook,
    source: Path,
    llm: Optional[LLMClient],
    query_cache: Optional[QueryCache] = None,
) -> ManifestWeightInfo:
    summary = build_manifest_weight_context(workbook)
    if not summary["sheets"]:
        return ManifestWeightInfo(total_weight_kg=None)
    cache_key = hashlib.sha256(
        json.dumps(summary, ensure_ascii=False, sort_keys=True).encode("utf-8", errors="ignore")
    ).hexdigest()
    if query_cache is not None:
        manifest_cache = query_cache.setdefault("manifest", {})
        if cache_key in manifest_cache and isinstance(manifest_cache[cache_key], dict):
            return normalize_manifest_weight_payload(manifest_cache[cache_key])
    if llm is None:
        return infer_manifest_weight_from_context(summary)

    payload = await llm.chat_json(build_manifest_weight_parser_messages(source.name, summary), temperature=0.0)
    weight_info = normalize_manifest_weight_payload(payload)
    if query_cache is not None:
        query_cache.setdefault("manifest", {})[cache_key] = payload
    return weight_info


def build_manifest_weight_context(workbook) -> dict[str, Any]:
    sheets: list[dict[str, Any]] = []
    for sheet in workbook.worksheets[:4]:
        headers = [clean_text(sheet.cell(1, col).value) for col in range(1, sheet.max_column + 1)]
        numeric_columns: list[dict[str, Any]] = []
        for col_idx, header in enumerate(headers, start=1):
            if not header:
                continue
            values: list[float] = []
            sample_values: list[float] = []
            for row_idx in range(2, sheet.max_row + 1):
                value = to_float(sheet.cell(row_idx, col_idx).value)
                if value is None or value <= 0:
                    continue
                values.append(value)
                if len(sample_values) < 8:
                    sample_values.append(value)
            if values:
                numeric_columns.append(
                    {
                        "column": col_idx,
                        "header": header,
                        "count": len(values),
                        "sum": round(sum(values), 4),
                        "max": round(max(values), 4),
                        "sample_values": sample_values,
                    }
                )

        sample_rows: list[list[Any]] = []
        for row_idx in range(1, min(sheet.max_row, 12) + 1):
            sample_rows.append(
                [
                    clean_text(sheet.cell(row_idx, col_idx).value)[:120]
                    for col_idx in range(1, min(sheet.max_column, 12) + 1)
                ]
            )
        tail_rows: list[list[Any]] = []
        start_tail = max(1, sheet.max_row - 5)
        for row_idx in range(start_tail, sheet.max_row + 1):
            tail_rows.append(
                [
                    clean_text(sheet.cell(row_idx, col_idx).value)[:120]
                    for col_idx in range(1, min(sheet.max_column, 12) + 1)
                ]
            )
        sheets.append(
            {
                "name": sheet.title,
                "max_row": sheet.max_row,
                "max_column": sheet.max_column,
                "headers": headers,
                "numeric_columns": numeric_columns,
                "sample_rows": sample_rows,
                "tail_rows": tail_rows,
            }
        )
    return {"sheets": sheets}


def build_manifest_weight_parser_messages(filename: str, summary: dict[str, Any]) -> list[dict[str, str]]:
    context = json.dumps(summary, ensure_ascii=False, separators=(",", ":"))
    return [
        {
            "role": "system",
            "content": (
                "你是装箱清单 Excel 字段识别专家。你的任务是根据工作簿摘要判断清单的 Excel 总重量。"
                "必须优先识别表示整票或逐行总毛重的列，例如 总毛重KGS、Gross Weight、G.W.、Weight KGS。"
                "不要使用提单重量，也不要使用体积、数量、箱数、单价、金额。只返回 JSON object。"
            ),
        },
        {
            "role": "user",
            "content": (
                f"文件名：{filename}\n"
                "请从以下 Excel 摘要中识别清单总重量，单位统一为 kg。\n"
                "规则：\n"
                "1. 如果有明确总重量合计行，返回该合计。\n"
                "2. 如果没有合计行，但存在逐行总毛重列，返回该列有效正数求和。\n"
                "3. 如果只有单箱重量或单件重量，不要猜总重量，返回 null。\n"
                "4. evidence 要说明使用了哪个 sheet、哪一列、如何计算。\n"
                "JSON格式：{\"total_weight_kg\":0.0,\"source\":\"\",\"evidence\":\"\",\"confidence\":0.0}\n"
                f"Excel摘要：\n{context[:18000]}"
            ),
        },
    ]


def normalize_manifest_weight_payload(payload: dict[str, Any]) -> ManifestWeightInfo:
    weight = to_float(payload.get("total_weight_kg"))
    confidence = to_float(payload.get("confidence"))
    confidence = max(0.0, min(1.0, confidence if confidence is not None else 0.0))
    if weight is not None and weight <= 0:
        weight = None
    return ManifestWeightInfo(
        total_weight_kg=weight,
        source=clean_text(payload.get("source")),
        evidence=clean_text(payload.get("evidence")),
        confidence=confidence,
    )


def infer_manifest_weight_from_context(summary: dict[str, Any]) -> ManifestWeightInfo:
    best: Optional[dict[str, Any]] = None
    for sheet in summary.get("sheets") or []:
        for column in sheet.get("numeric_columns") or []:
            header_key = normalize_text(column.get("header"))
            if any(term in header_key for term in ("毛重", "总毛重", "gross weight", "gw", "g w", "weight kgs")):
                if "体积" in header_key or "cbm" in header_key:
                    continue
                if best is None or (column.get("count") or 0) > (best.get("count") or 0):
                    best = {**column, "sheet": sheet.get("name")}
    if not best:
        return ManifestWeightInfo(total_weight_kg=None)
    return ManifestWeightInfo(
        total_weight_kg=to_float(best.get("sum")),
        source=f"{best.get('sheet')}.{best.get('header')}",
        evidence=f"按 {best.get('sheet')} sheet 的 {best.get('header')} 列正数求和",
        confidence=0.7,
    )


async def parse_manifest(
    path: str | Path,
    llm: Optional[LLMClient] = None,
    query_cache: Optional[QueryCache] = None,
) -> ManifestSummary:
    source = Path(path)
    workbook = load_workbook(source, data_only=True)
    sheet = workbook[workbook.sheetnames[0]]
    headers = [str(sheet.cell(1, col).value or "").strip() for col in range(1, sheet.max_column + 1)]
    header_map = {normalize_header(name): idx + 1 for idx, name in enumerate(headers)}

    def col(*names: str) -> Optional[int]:
        for name in names:
            key = normalize_header(name)
            if key in header_map:
                return header_map[key]
        for idx, header in enumerate(headers, start=1):
            flat = normalize_header(header)
            if any(normalize_header(name) in flat for name in names):
                return idx
        return None

    ctn_col = col("件数", "箱数", "箱数CTN", "CTN")
    real_weight_col = col("实重", "净重")
    gross_weight_col = col("总抛重", "总毛重", "毛重", "总毛重KGS", "KGS")
    value_col = col("清关申报金额", "申报金额", "总价", "总金额", "总金额USD")
    zh_col = col("中文品名", "中英文品名", "品名")
    en_col = col("英文品名", "中英文品名", "DESCRIPTION OF GOODS")
    hs_col = col("清关HS CODE", "报关编码 HS", "HS CODE", "商品编码")
    unit_price_col = col("单价", "单价USD")
    qty_col = col("总数量", "总数量PCS", "数量")
    material_col = col("中英文材质", "材质", "MATERIAL")
    usage_col = col("中英文用途", "用途", "USE FOR")

    total_ctns = 0.0
    total_real_weight = 0.0
    total_declared_value = 0.0
    row_count = 0
    items: list[ManifestItem] = []
    categories: list[str] = []
    seen_categories: set[str] = set()

    for row_idx in range(2, sheet.max_row + 1):
        zh = clean_text(sheet.cell(row_idx, zh_col).value if zh_col else "")
        en = clean_text(sheet.cell(row_idx, en_col).value if en_col else "")
        has_any = any(sheet.cell(row_idx, col_idx).value not in (None, "") for col_idx in range(1, sheet.max_column + 1))
        if not has_any or not zh:
            continue

        ctns = to_float(sheet.cell(row_idx, ctn_col).value if ctn_col else None)
        real_weight = to_float(sheet.cell(row_idx, real_weight_col).value if real_weight_col else None)
        gross_weight = to_float(sheet.cell(row_idx, gross_weight_col).value if gross_weight_col else None)
        declared_value = to_float(sheet.cell(row_idx, value_col).value if value_col else None)
        qty = to_float(sheet.cell(row_idx, qty_col).value if qty_col else None)
        unit_price = to_float(sheet.cell(row_idx, unit_price_col).value if unit_price_col else None)

        row_count += 1
        key = normalize_text(zh)
        if key and key not in seen_categories:
            seen_categories.add(key)
            categories.append(zh)

        total_ctns += ctns or 0
        total_real_weight += real_weight or gross_weight or 0
        total_declared_value += declared_value or 0
        items.append(
            ManifestItem(
                row=row_idx,
                zh=zh,
                en=en,
                hs=normalize_hs(sheet.cell(row_idx, hs_col).value if hs_col else ""),
                material=clean_text(sheet.cell(row_idx, material_col).value if material_col else ""),
                usage=clean_text(sheet.cell(row_idx, usage_col).value if usage_col else ""),
                ctns=ctns,
                qty=qty,
                unit_price=unit_price,
                declared_value=declared_value,
                real_weight=real_weight,
                gross_weight=gross_weight,
            )
        )

    weight_info = await parse_manifest_total_weight(workbook, source, llm, query_cache)
    if weight_info.total_weight_kg is not None and weight_info.total_weight_kg > 0:
        total_real_weight = weight_info.total_weight_kg
        weight_source = weight_info.source
        weight_evidence = weight_info.evidence
        weight_confidence = weight_info.confidence
    else:
        weight_source = "parsed_rows" if total_real_weight > 0 else ""
        weight_evidence = "逐行重量字段求和" if total_real_weight > 0 else ""
        weight_confidence = 0.6 if total_real_weight > 0 else 0.0

    return ManifestSummary(
        filename=source.name,
        row_count=row_count,
        total_ctns=round(total_ctns, 2),
        total_real_weight=round(total_real_weight, 2),
        total_declared_value=round(total_declared_value, 2),
        categories=categories,
        items=items,
        weight_source=weight_source,
        weight_evidence=weight_evidence,
        weight_confidence=round(weight_confidence, 4),
    )


async def parse_bill(path: str | Path, llm: LLMClient, query_cache: Optional[QueryCache] = None) -> BillInfo:
    source = Path(path)
    reader = PdfReader(str(source))
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    product_entries = await parse_bill_product_entries_with_llm(text, llm, query_cache)
    products = [entry.name for entry in product_entries]
    return BillInfo(
        filename=source.name,
        raw_text=text,
        products=products,
        shipper=extract_shipper(text),
        consignee=extract_consignee(text),
        shipment_no=extract_shipment_no(text),
        eta=extract_eta(text),
        cartons=extract_number(r"(\d+(?:\.\d+)?)\s*CARTONS?", text),
        gross_weight=extract_number(r"(\d+(?:\.\d+)?)\s*KGS?", text),
        cbm=extract_number(r"(\d+(?:\.\d+)?)\s*CBM", text),
        product_entries=product_entries,
    )


async def parse_bill_products_with_llm(
    text: str,
    llm: LLMClient,
    query_cache: Optional[QueryCache] = None,
) -> list[str]:
    return [entry.name for entry in await parse_bill_product_entries_with_llm(text, llm, query_cache)]


async def parse_bill_product_entries_with_llm(
    text: str,
    llm: LLMClient,
    query_cache: Optional[QueryCache] = None,
) -> list[BillProduct]:
    cache_key = hashlib.sha256(text.encode("utf-8", errors="ignore")).hexdigest()
    if query_cache is not None:
        bill_cache = query_cache.setdefault("bill", {})
        if cache_key in bill_cache:
            cached = bill_cache[cache_key]
            if isinstance(cached, dict):
                return normalize_bill_llm_product_entries(cached)

    payload = await llm.chat_json(build_bill_parser_messages(text), temperature=0.0)
    products = normalize_bill_llm_product_entries(payload)
    if query_cache is not None:
        query_cache.setdefault("bill", {})[cache_key] = payload
    return products


def build_bill_parser_messages(text: str) -> list[dict[str, str]]:
    return [
        {
            "role": "system",
            "content": (
                "你是国际海运提单商业字段解析专家。你的任务是从 pypdf 提取出的提单原始文本里识别真实货物品类。"
                "必须区分货物品类和提单字段、日期、港口、船司、付款条款、装船批注。只返回 JSON object。"
            ),
        },
        {
            "role": "user",
            "content": (
                "从以下提单文本中提取真实货物品类。\n"
                "规则：\n"
                "1. products 只放货物品类英文名，保持提单原文语义，可去掉 HS CODE 和编码。\n"
                "2. 不要把 SHIPPED ON BOARD、ON BOARD、PORT OF LOADING、FREIGHT、EXPRESS BILL、"
                "TOTAL NUMBER OF CONTAINERS、日期、港口、公司名、地址识别为品类。\n"
                "3. 如果文本中出现类似 'STORAGE BAG HS CODE:420222'，品类是 'STORAGE BAG'。\n"
                "4. 如果没有可靠品类，返回空数组，不要猜。\n"
                "JSON格式：{\"products\":[{\"name\":\"\",\"hs_code_hint\":\"\",\"evidence\":\"\",\"confidence\":0.0}],"
                "\"ignored_phrases\":[{\"text\":\"\",\"reason\":\"\"}]}\n"
                f"提单文本：\n{text[:12000]}"
            ),
        },
    ]


def normalize_bill_llm_products(payload: dict[str, Any]) -> list[str]:
    return [entry.name for entry in normalize_bill_llm_product_entries(payload)]


def normalize_bill_llm_product_entries(payload: dict[str, Any]) -> list[BillProduct]:
    raw_products = payload.get("products")
    if not isinstance(raw_products, list):
        raise RuntimeError("LLM 提单解析未返回 products 数组")

    products: list[BillProduct] = []
    seen: set[str] = set()
    for raw in raw_products:
        if isinstance(raw, dict):
            name = clean_text(raw.get("name"))
            confidence = to_float(raw.get("confidence"))
            hs_code_hint = normalize_hs(raw.get("hs_code_hint") or raw.get("hs") or raw.get("hs_code"))
            evidence = clean_text(raw.get("evidence"))
        else:
            name = clean_text(raw)
            confidence = None
            hs_code_hint = ""
            evidence = ""
        name = normalize_bill_llm_product_name(name)
        if not name:
            continue
        if confidence is not None and confidence < 0.5:
            continue
        key = normalize_text(name)
        if key and key not in seen:
            seen.add(key)
            products.append(
                BillProduct(
                    name=name,
                    hs_code_hint=hs_code_hint,
                    evidence=evidence,
                    confidence=confidence if confidence is not None else 1.0,
                )
            )
    if not products:
        raise RuntimeError("LLM 未从提单中识别到可靠货物品类")
    return products


def normalize_bill_llm_product_name(value: str) -> str:
    text = clean_text(value)
    if not text:
        return ""
    text = re.sub(r"\bHS\s*CODE\b.*$", "", text, flags=re.IGNORECASE).strip(" ：:.,;")
    text = re.sub(r"\s+", " ", text)
    if not text or not any(char.isalpha() for char in text):
        return ""
    if re.search(r"\d{4,}|@|[()/]", text):
        return ""
    blocked = {
        "shipped on board",
        "hippedon board",
        "on board",
        "port of loading",
        "port of discharge",
        "freight charge",
        "express bill",
        "total number of containers",
    }
    if normalize_text(text) in blocked:
        return ""
    return text


async def llm_generate_declaration(
    llm: LLMClient,
    manifest: ManifestSummary,
    bill: BillInfo,
    input_tax_data: dict[str, Any],
    profile_hint: str,
    feedback: str = "",
) -> dict[str, Any]:
    messages = [
        {
            "role": "system",
            "content": (
                "你是美国进口清关资料生成专家。你必须根据清单、运输提单和爬虫税率结果生成申报行。"
                "不要逐行搬运清单；要像乙方清关行一样归并、改申报名、选择可申报 HS、低申报，并保证提单品类进入输出。"
                "只返回 JSON，不要 markdown。"
            ),
        },
        {
            "role": "user",
            "content": build_llm_prompt(manifest, bill, input_tax_data, profile_hint, feedback),
        },
    ]
    return await llm.chat_json(messages, temperature=0.1)


async def generate_valid_declaration(
    llm: LLMClient,
    manifest: ManifestSummary,
    bill: BillInfo,
    input_tax_data: dict[str, Any],
    profile_hint: str,
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    feedback = ""
    last_error = ""
    for _attempt in range(4):
        payload = await llm_generate_declaration(llm, manifest, bill, input_tax_data, profile_hint, feedback)
        try:
            rows = normalize_llm_rows(payload, manifest, bill)
            validate_declaration_strategy(rows, bill.products)
            return payload, rows
        except RuntimeError as exc:
            last_error = str(exc)
            feedback = (
                "上一次输出不合格，必须修正后重新输出 JSON。错误: "
                f"{last_error}。禁止输出鞋、服装、焊机、热水器、化妆品等原始高风险名称；"
                "非提单品类必须改用 reference_style_rows 里的申报名。不要解释，只输出修正后的 JSON。"
            )
    raise RuntimeError(f"LLM 输出连续不合格: {last_error}")


def build_llm_prompt(manifest: ManifestSummary, bill: BillInfo, input_tax_data: dict[str, Any], profile_hint: str, feedback: str = "") -> str:
    compact_items = aggregate_manifest_for_llm(manifest)
    tax_summary = summarize_tax_data(input_tax_data)
    reference_rows = load_reference_style_rows()
    allowed_reference_names = sorted({row["中文品名"] for row in reference_rows if row.get("中文品名")})
    manifest_json = json.dumps(compact_items, ensure_ascii=False, separators=(",", ":"))
    tax_json = json.dumps(tax_summary, ensure_ascii=False, separators=(",", ":"))
    reference_json = json.dumps(reference_rows, ensure_ascii=False, separators=(",", ":"))
    bill_json = json.dumps(
        {
            "products": bill.products,
            "shipment_no": bill.shipment_no,
            "eta": bill.eta,
            "cartons": bill.cartons,
            "gross_weight": bill.gross_weight,
            "cbm": bill.cbm,
            "shipper": bill.shipper,
            "consignee": bill.consignee,
        },
        ensure_ascii=False,
        separators=(",", ":"),
    )
    return (
        "生成美国清关 Commercial Invoice & Packing List 申报行，只返回 JSON object。\n"
        f"profile_hint={profile_hint or 'auto'}; file={manifest.filename}; rows={manifest.row_count}; "
        f"total_ctns={manifest.total_ctns}; total_real_weight={manifest.total_real_weight}; "
        f"input_declared_value={manifest.total_declared_value}\n"
        f"bill={bill_json}\n"
        f"manifest_grouped={manifest_json}\n"
        f"tax_summary={tax_json}\n"
        f"reference_style_rows={reference_json}\n"
        f"allowed_reference_names={json.dumps(allowed_reference_names, ensure_ascii=False, separators=(',', ':'))}\n"
        "强制规则:\n"
        "1. 输出必须包含提单品类，每个提单品类必须出现在某一行的中英文品名里，并同时写入该行 bill_sources。\n"
        "2. 输出总箱数必须等于输入清单 total_ctns。\n"
        "3. 输出总毛重必须等于输入清单 total_real_weight。\n"
        "4. 每行净重等于毛重减箱数；每行总价等于数量乘单价。\n"
        "5. 输出行数建议 8-13 行。除提单品类中文直译行以外，中文品名必须从 allowed_reference_names 选择，不得自造泛名。\n"
        "6. 商品编码必须尽量使用可在 codeflagai 查到税率的 10 位美国 HTS，不要输出中国侧泛码或前缀码。\n"
        "7. 不允许输出空 HS、空品名、待补充、0 金额、0 重量。\n"
        "8. 必须按相似品类和相同/相近材质归并压缩，不得原样输出高风险高价值品类：鞋、电焊机、热水器、化妆品、服装原品名。\n"
        "9. 不得把高风险品类漂移成无依据的泛名，如家居五金配件、电脑周边配件、塑料过滤配件、家居收纳袋、防滑垫、休闲鞋。\n"
        "10. 不得沿用原清单高单价；申报总价目标为 6000-7500 USD，单行单价通常不超过 3 USD。\n"
        "11. 若清单含鞋/服装/焊机/热水器/化妆品，把其箱数/重量压缩分配到 allowed_reference_names 的低风险家居、塑料、装饰、灯具、文具、五金行中。\n"
        f"{'修正反馈: ' + feedback if feedback else ''}\n"
        'JSON格式: {"metadata":{"shipper":"","consignee":"","eta":"YYYY/M/D","shipment_no":""},'
        '"reasoning_summary":"一句话",'
        '"rows":[{"中文品名":"","英文品名":"","商品编码":"","材质":"","用途":"HOME",'
        '"箱数":1,"数量":1,"单位":"PCS","币制":"USD","单价":1,"总价":1,"毛重":1,'
        '"原产国":"CN","bill_sources":[],"source_rows":[]}]}'
    )


def summarize_tax_data(tax_data: dict[str, Any]) -> list[dict[str, Any]]:
    rows = []
    for code, data in list(tax_data.items())[:25]:
        rows.append(
            {
                "hs": code,
                "source_hs_or_cn_hs": data.get("hs_code_cn"),
                "description": data.get("description_cn"),
                "taric": data.get("taric"),
                "tax_rate": data.get("tax_rate"),
                "anti_dumping": data.get("anti_dumping"),
            }
        )
    return rows


def load_reference_style_rows(limit: int = 24) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    if not REFERENCE_ROOT.exists():
        return DEFAULT_REFERENCE_STYLE_ROWS[:limit]

    for path in sorted(REFERENCE_ROOT.rglob("*清关资料.xlsx")):
        try:
            workbook = load_workbook(path, data_only=True)
        except Exception:
            continue
        sheet = workbook[workbook.sheetnames[0]]
        for row_idx in range(6, sheet.max_row + 1):
            zh = clean_text(sheet.cell(row_idx, 1).value)
            en = clean_text(sheet.cell(row_idx, 2).value)
            hs = normalize_hs(sheet.cell(row_idx, 3).value)
            if not zh or not en or not hs:
                continue
            rows.append(
                {
                    "中文品名": zh,
                    "英文品名": en,
                    "商品编码": hs,
                    "材质": clean_text(sheet.cell(row_idx, 4).value),
                    "用途": clean_text(sheet.cell(row_idx, 5).value) or "HOME",
                    "单价": to_float(sheet.cell(row_idx, 10).value),
                }
            )
            if len(rows) >= limit:
                return rows
    return (rows or DEFAULT_REFERENCE_STYLE_ROWS)[:limit]


def aggregate_manifest_for_llm(manifest: ManifestSummary) -> list[dict[str, Any]]:
    groups: dict[tuple[str, str, str], dict[str, Any]] = {}
    for item in manifest.items:
        key = (item.zh, item.en, item.hs)
        if key not in groups:
            groups[key] = {
                "rows": [],
                "zh": item.zh,
                "en": item.en,
                "hs": item.hs,
                "material": item.material,
                "usage": item.usage,
                "ctns": 0.0,
                "qty": 0.0,
                "value": 0.0,
                "real_weight": 0.0,
            }
        group = groups[key]
        group["rows"].append(item.row)
        group["ctns"] += item.ctns or 0
        group["qty"] += item.qty or 0
        group["value"] += item.declared_value or 0
        group["real_weight"] += item.real_weight or 0

    ranked = sorted(
        groups.values(),
        key=lambda row: (row["real_weight"], row["ctns"], row["value"]),
        reverse=True,
    )
    for row in ranked:
        row["ctns"] = round(row["ctns"], 2)
        row["qty"] = round(row["qty"], 2)
        row["value"] = round(row["value"], 2)
        row["real_weight"] = round(row["real_weight"], 2)
    return ranked[:35]


def select_representative_input_hs(manifest: ManifestSummary, limit: int) -> list[str]:
    scores: dict[str, float] = {}
    for item in manifest.items:
        hs = normalize_hs(item.hs)
        if not hs:
            continue
        score = 0.0
        score += item.real_weight or 0
        score += (item.ctns or 0) * 8
        score += (item.declared_value or 0) / 100
        score += (item.qty or 0) / 20
        scores[hs] = scores.get(hs, 0) + score

    ranked = sorted(scores, key=lambda code: scores[code], reverse=True)
    return ranked[:limit]


def normalize_llm_rows(payload: dict[str, Any], manifest: ManifestSummary, bill: BillInfo) -> list[dict[str, Any]]:
    raw_rows = payload.get("rows")
    if not isinstance(raw_rows, list) or not raw_rows:
        raise RuntimeError("LLM 未返回 rows")

    rows: list[dict[str, Any]] = []
    for idx, raw in enumerate(raw_rows, start=1):
        if not isinstance(raw, dict):
            raise RuntimeError(f"LLM rows[{idx}] 不是 object")
        row = normalize_output_row(raw, idx)
        rows.append(row)

    apply_bill_product_names(rows, bill.products)
    ensure_bill_products_present(rows, bill.products)
    reconcile_totals(rows, manifest.total_ctns, manifest.total_real_weight)
    validate_output_rows(rows)
    return rows


def normalize_output_row(raw: dict[str, Any], idx: int) -> dict[str, Any]:
    row = {}
    for field in HEADERS:
        if field == "净重":
            continue
        value = raw.get(field)
        if value is None and field == "商品编码":
            value = raw.get("HS") or raw.get("HS编码") or raw.get("HTS Code")
        if value is None:
            raise RuntimeError(f"LLM 第 {idx} 行缺少字段: {field}")
        row[field] = value

    row["商品编码"] = hs_cell_value(row["商品编码"])
    row["箱数"] = to_float(row["箱数"])
    row["数量"] = to_float(row["数量"])
    row["单价"] = to_float(row["单价"])
    row["总价"] = to_float(row["总价"])
    row["毛重"] = to_float(row["毛重"])
    if None in (row["箱数"], row["数量"], row["单价"], row["毛重"]):
        raise RuntimeError(f"LLM 第 {idx} 行数值字段不完整")

    row["单位"] = clean_text(row.get("单位")) or "PCS"
    row["币制"] = clean_text(row.get("币制")) or "USD"
    row["原产国"] = clean_text(row.get("原产国")) or "CN"
    row["材质"] = translate_material_to_english(row.get("材质"))
    row["用途"] = translate_usage_to_english(row.get("用途"))
    row["bill_sources"] = raw.get("bill_sources") if isinstance(raw.get("bill_sources"), list) else []
    row["source_rows"] = raw.get("source_rows") if isinstance(raw.get("source_rows"), list) else []
    return row


def ensure_bill_products_present(
    rows: list[dict[str, Any]],
    bill_products: list[str],
    source_candidates: Optional[list[ProductCandidate]] = None,
) -> None:
    for product in bill_products:
        covered = False
        for idx, row in enumerate(rows):
            if not row_matches_single_bill_product(row, product):
                continue
            if source_candidates is not None and not row_has_matching_bill_candidate(row, idx, product, source_candidates):
                continue
            covered = True
            break
        if not covered:
            raise RuntimeError(f"输出行名未包含提单品类: {product}")


def row_has_matching_bill_candidate(
    row: dict[str, Any],
    row_index: int,
    product: str,
    source_candidates: list[ProductCandidate],
) -> bool:
    if row_index >= len(source_candidates):
        return False
    candidate = source_candidates[row_index]
    if not candidate_matches_single_bill_product(candidate, product):
        return False
    if not candidate_has_product_tax_match(candidate):
        return False
    row_hs = normalize_hs(row.get("商品编码"))
    candidate_hs = normalize_hs(candidate.hs)
    return bool(row_hs and candidate_hs and row_hs == candidate_hs)


def apply_bill_product_names(
    rows: list[dict[str, Any]],
    bill_products: list[str],
    source_candidates: Optional[list[ProductCandidate]] = None,
    force: bool = False,
) -> None:
    used_indices: set[int] = set()
    for product_index, product in enumerate(bill_products):
        canonical = canonical_bill_product(product)
        target_index: Optional[int] = None
        if force and product_index < len(rows) and product_index not in used_indices:
            target_index = product_index
        for idx, row in enumerate(rows):
            if target_index is not None:
                break
            if idx in used_indices:
                continue
            if source_candidates is not None and not row_has_matching_bill_candidate(row, idx, product, source_candidates):
                continue
            bill_sources = [normalize_text(item) for item in row.get("bill_sources") or []]
            source_match = normalize_text(product) in bill_sources
            row_match = row_matches_single_bill_product(row, product)
            if source_match or row_match:
                target_index = idx
                break
        if target_index is None:
            continue
        apply_bill_product_display_name(rows[target_index], product, canonical)
        used_indices.add(target_index)


def apply_bill_product_display_name(
    row: dict[str, Any],
    product: str,
    canonical: Optional[dict[str, str]],
) -> None:
    if canonical:
        row["中文品名"] = canonical["zh"]
        row["英文品名"] = canonical["en"]
        row["材质"] = canonical["material"]
        row["用途"] = canonical.get("usage", row.get("用途") or "HOME")
    else:
        product_name = clean_text(product)
        row["中文品名"] = product_name
        row["英文品名"] = product_name.title() if product_name.isupper() else product_name
    if product not in row.get("bill_sources", []):
        row["bill_sources"] = [*(row.get("bill_sources") or []), product]


def canonical_bill_product(product: str) -> Optional[dict[str, str]]:
    words = set(normalize_bill_product_text(product).split())
    if {"silicone", "coaster"} <= words:
        return {
            "zh": "硅胶杯垫",
            "en": "Silicone Coaster",
            "material": "Silicone",
            "usage": "HOME",
        }
    if "phone" in words and "holder" in words:
        return {
            "zh": "塑料手机支架",
            "en": "Plastic Mobile Phone Stand",
            "material": "Plastic",
            "usage": "HOME",
        }
    return None


def reconcile_totals(rows: list[dict[str, Any]], target_ctns: float, target_gross: float) -> None:
    ctn_values = [to_float(row["箱数"]) or 0 for row in rows]
    gross_values = [to_float(row["毛重"]) or 0 for row in rows]
    scaled_ctns = scale_integer(ctn_values, target_ctns)
    scaled_gross = scale_decimal(gross_values, target_gross, 2)

    for row, ctns, gross in zip(rows, scaled_ctns, scaled_gross):
        old_ctns = to_float(row["箱数"]) or ctns
        old_qty = to_float(row["数量"]) or old_ctns
        qty_per_ctn = old_qty / old_ctns if old_ctns else 1
        qty = max(1, round(ctns * qty_per_ctn))
        unit_price = to_float(row["单价"]) or 1
        row["箱数"] = ctns
        row["数量"] = qty
        row["毛重"] = gross
        row["净重"] = round(gross - ctns, 2)
        row["单价"] = unit_price
        row["总价"] = round(qty * unit_price, 2)


def validate_output_rows(rows: list[dict[str, Any]]) -> None:
    for idx, row in enumerate(rows, start=1):
        for field in HEADERS:
            if field not in row or row[field] in (None, ""):
                raise RuntimeError(f"输出第 {idx} 行缺少字段: {field}")
        hs = normalize_hs(row["商品编码"])
        if not (8 <= len(hs) <= 10):
            raise RuntimeError(f"输出第 {idx} 行 HS 编码不合法: {row['商品编码']}")
        for field in ("箱数", "数量", "单价", "总价", "净重", "毛重"):
            value = to_float(row[field])
            if value is None or value <= 0:
                raise RuntimeError(f"输出第 {idx} 行 {field} 必须大于 0")


def validate_declaration_strategy(rows: list[dict[str, Any]], bill_products: Optional[list[str]] = None) -> None:
    if not (8 <= len(rows) <= 13):
        raise RuntimeError(f"输出行数必须压缩到 8-13 行，当前 {len(rows)} 行")
    total_value = sum(to_float(row.get("总价")) or 0 for row in rows)
    if not (5000 <= total_value <= 9000):
        raise RuntimeError(f"申报总价必须低申报到 5000-9000 USD，当前 {round(total_value, 2)}")
    allowed_names = {normalize_text(item.get("中文品名")) for item in load_reference_style_rows() if item.get("中文品名")}
    blocked_terms = (
        "鞋", "电焊机", "热水器", "眼影", "服饰", "衬衣", "长裤",
        "welding", "heater", "shoes", "eyeshadow", "apparel",
    )
    drift_terms = ("家居五金配件", "电脑周边配件", "塑料过滤配件", "家居收纳袋", "防滑垫")
    for idx, row in enumerate(rows, start=1):
        row_name = normalize_text(f"{row.get('中文品名')} {row.get('英文品名')}")
        if any(normalize_text(term) in row_name for term in blocked_terms):
            raise RuntimeError(f"第 {idx} 行未压缩高风险原品类: {row.get('中文品名')} / {row.get('英文品名')}")
        if any(normalize_text(term) in row_name for term in drift_terms):
            raise RuntimeError(f"第 {idx} 行申报名过度泛化/语义漂移: {row.get('中文品名')} / {row.get('英文品名')}")
        row_zh = normalize_text(row.get("中文品名"))
        if allowed_names and row_zh not in allowed_names and not row_matches_bill_product(row, bill_products or []):
            raise RuntimeError(f"第 {idx} 行申报名不在样例清关资料或提单品类中: {row.get('中文品名')} / {row.get('英文品名')}")
        unit_price = to_float(row.get("单价")) or 0
        if unit_price > 5:
            raise RuntimeError(f"第 {idx} 行单价过高，疑似沿用原清单价格: {unit_price}")


def row_matches_bill_product(row: dict[str, Any], bill_products: list[str]) -> bool:
    return any(row_matches_single_bill_product(row, product) for product in bill_products)


def row_matches_single_bill_product(row: dict[str, Any], product: str) -> bool:
    row_text = normalize_text(f"{row.get('中文品名')} {row.get('英文品名')}")
    row_words = set(row_text.split())
    product_key = normalize_bill_product_text(product)
    product_words = set(product_key.split())
    if product_key and product_key in normalize_bill_product_text(row_text):
        return True
    return bool(product_words and len(product_words & row_words) >= min(2, len(product_words)))


async def validate_rows_with_crawler(crawler: StrictTaxCrawler, rows: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    tax_data: dict[str, dict[str, Any]] = {}
    for index, row in enumerate(rows):
        if index > 0:
            await asyncio.sleep(crawler.settings.delay)
        product_name = str(row.get("中文品名") or row.get("英文品名") or "").strip()
        material = str(row.get("材质") or "").strip()
        original_hs = normalize_hs(row.get("商品编码"))
        candidates = await crawler.search_product(product_name, material)
        selected = select_crawler_candidate(row, candidates)
        verified_original = None
        if original_hs and normalize_hs(selected.get("hs_code_us")) != original_hs and len(original_hs) == 10:
            await asyncio.sleep(crawler.settings.delay)
            verified_original = find_verified_hs_candidate(original_hs, await crawler.search(original_hs))
        if verified_original:
            selected = verified_original
        selected_hs = normalize_hs(selected.get("hs_code_us"))
        if not selected_hs:
            raise RuntimeError(f"爬虫未给出可用 HS: {product_name}")
        if len(selected_hs) != 10:
            raise RuntimeError(f"爬虫返回的美国 HTS 不是 10 位: {product_name} -> {selected_hs}")
        if selected_hs != original_hs:
            row["LLM原HS"] = original_hs
            row["商品编码"] = hs_cell_value(selected_hs)
        row["税率"] = selected.get("tax_rate")
        row["爬虫匹配HS"] = selected_hs
        row["爬虫品名"] = selected.get("description_cn")
        tax_data[selected_hs] = selected
    return tax_data


def select_crawler_candidate(row: dict[str, Any], candidates: dict[str, dict[str, Any]]) -> dict[str, Any]:
    desired_hs = normalize_hs(row.get("商品编码"))
    if desired_hs and desired_hs in candidates and is_usable_tax_candidate(candidates[desired_hs]):
        return candidates[desired_hs]
    usable = [
        data for data in candidates.values()
        if is_usable_tax_candidate(data)
    ]
    if not usable:
        raise RuntimeError(f"爬虫候选无有效税率: {row.get('中文品名')}")
    non_ad = [data for data in usable if not data.get("anti_dumping")]
    return (non_ad or usable)[0]


def find_verified_hs_candidate(hs: str, candidates: dict[str, dict[str, Any]]) -> Optional[dict[str, Any]]:
    normalized = normalize_hs(hs)
    candidate = candidates.get(normalized)
    if candidate and is_usable_tax_candidate(candidate):
        return candidate
    for data in candidates.values():
        if normalize_hs(data.get("hs_code_us")) == normalized and is_usable_tax_candidate(data):
            return data
    return None


def is_usable_tax_candidate(data: dict[str, Any]) -> bool:
    hs = normalize_hs(data.get("hs_code_us"))
    tax_rate = str(data.get("tax_rate") or "").strip().upper()
    return len(hs) == 10 and bool(tax_rate) and tax_rate != "N/A"


def attach_tax_validation(rows: list[dict[str, Any]], tax_data: dict[str, dict[str, Any]]) -> None:
    for row in rows:
        hs = normalize_hs(row["商品编码"])
        candidates = [
            data for code, data in tax_data.items()
            if normalize_hs(data.get("hs_code_cn")) == hs or normalize_hs(code) == hs
        ]
        if not candidates:
            raise RuntimeError(f"爬虫未返回输出 HS 的税率结果: {hs}")
        best = candidates[0]
        row["税率"] = best.get("tax_rate")
        row["爬虫匹配HS"] = best.get("hs_code_us")
        row["爬虫品名"] = best.get("description_cn")


def normalize_output_language_fields(rows: list[dict[str, Any]]) -> None:
    for row in rows:
        row["材质"] = translate_material_to_english(row.get("材质"))
        row["用途"] = translate_usage_to_english(row.get("用途"))


def translate_material_to_english(value: Any) -> str:
    return translate_field_to_english(value, MATERIAL_TRANSLATIONS, "Mixed")


def translate_usage_to_english(value: Any) -> str:
    normalized = normalize_text(value)
    if normalized in {"", "home", "home use", "household", "household use", "household 家用"}:
        return "Home use"
    return translate_field_to_english(value, USAGE_TRANSLATIONS, "Home use")


def translate_field_to_english(value: Any, translations: tuple[tuple[str, str], ...], default: str) -> str:
    text = clean_text(value)
    if not text:
        return default

    text = normalize_field_separators(text)
    for chinese, english in sorted(translations, key=lambda item: len(item[0]), reverse=True):
        text = text.replace(chinese, f" {english} ")
    text = re.sub(r"[\u4e00-\u9fff]+", " ", text)
    text = normalize_field_spacing(text)
    text = normalize_english_terms(text)
    text = collapse_repeated_phrases(text)
    text = normalize_field_spacing(text)
    return text or default


def normalize_field_separators(value: str) -> str:
    text = value.replace("\xa0", " ")
    replacements = {
        "／": "/",
        "，": ",",
        "、": ",",
        "；": ";",
        "：": ":",
        "（": "(",
        "）": ")",
        "％": "%",
        "＋": "+",
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
    return text


def normalize_field_spacing(value: str) -> str:
    text = value
    text = re.sub(r"(?<=\d)%(?=[A-Za-z])", "% ", text)
    text = re.sub(r"(?<=[A-Za-z])(?=\d)", " ", text)
    text = re.sub(r"(?<=\d)(?=[A-Za-z])", " ", text)
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"\s*,\s*", ", ", text)
    text = re.sub(r"\s*;\s*", "; ", text)
    text = re.sub(r"\s*/\s*", "/", text)
    text = re.sub(r"\s*\+\s*", "+", text)
    text = re.sub(r"\s*:\s*", ": ", text)
    text = re.sub(r"\(\s*", "(", text)
    text = re.sub(r"\s*\)", ")", text)
    text = re.sub(r"\(\)", " ", text)
    text = re.sub(r"\+{2,}", "+", text)
    text = re.sub(r",{2,}", ",", text)
    text = re.sub(r";{2,}", ";", text)
    return text.strip(" ,;/+:-")


def normalize_english_terms(value: str) -> str:
    text = value
    for source, target in sorted(ENGLISH_TERM_NORMALIZATIONS, key=lambda item: len(item[0]), reverse=True):
        text = re.sub(rf"\b{re.escape(source)}\b", target, text, flags=re.IGNORECASE)
    for acronym in FIELD_TRANSLATION_ACRONYMS:
        text = re.sub(rf"\b{re.escape(acronym)}\b", acronym, text, flags=re.IGNORECASE)
    return smart_capitalize_field(text)


def smart_capitalize_field(value: str) -> str:
    def replace_word(match: re.Match[str]) -> str:
        word = match.group(0)
        upper = word.upper()
        if upper in FIELD_TRANSLATION_ACRONYMS:
            return upper
        return word[:1].upper() + word[1:].lower()

    return re.sub(r"[A-Za-z]+", replace_word, value)


def collapse_repeated_phrases(value: str) -> str:
    text = value
    for _ in range(4):
        collapsed = re.sub(
            r"\b([A-Za-z0-9%]+(?:\s+[A-Za-z0-9%]+){0,4})\s+\1\b",
            r"\1",
            text,
            flags=re.IGNORECASE,
        )
        if collapsed == text:
            break
        text = collapsed
    parts = [part.strip() for part in re.split(r"([/+,;])", text) if part.strip()]
    result: list[str] = []
    seen_terms: set[str] = set()
    for part in parts:
        if part in {"/", "+", ",", ";"}:
            if result and result[-1] not in {"/", "+", ",", ";"}:
                result.append(part)
            continue
        key = normalize_text(part)
        if key and key in seen_terms:
            if result and result[-1] in {"/", "+", ",", ";"}:
                result.pop()
            continue
        result.append(part)
        if key:
            seen_terms.add(key)
    return "".join(result)


def write_workbook(bill: BillInfo, rows: list[dict[str, Any]], output_path: Path, metadata: dict[str, Any]) -> None:
    if not LOCAL_TEMPLATE_PATH.exists():
        raise RuntimeError(f"模板不存在: {LOCAL_TEMPLATE_PATH}")
    workbook = load_workbook(LOCAL_TEMPLATE_PATH)
    sheet = workbook[workbook.sheetnames[0]]
    fill_metadata(sheet, bill, metadata)
    clear_output_area(sheet, start_row=6, end_row=max(sheet.max_row, 80))
    for idx, row in enumerate(rows, start=6):
        write_output_row(sheet, idx, row)
    workbook.save(output_path)


def fill_metadata(sheet, bill: BillInfo, metadata: dict[str, Any]) -> None:
    shipper = clean_text(metadata.get("shipper")) or bill.shipper
    consignee = clean_text(metadata.get("consignee")) or bill.consignee
    eta = clean_text(metadata.get("eta")) or bill.eta
    shipment_no = clean_text(metadata.get("shipment_no")) or bill.shipment_no
    sheet["A2"] = "SHIPPER"
    sheet["C2"] = shipper
    sheet["A3"] = "CONSIGNEE"
    sheet["C3"] = consignee
    sheet["A4"] = f"ETA:{eta}" if eta else "ETA:"
    sheet["C4"] = f"B/L :{shipment_no}" if shipment_no else "B/L :"


def clear_output_area(sheet, start_row: int, end_row: int) -> None:
    for row in range(start_row, end_row + 1):
        for col in range(1, 15):
            sheet.cell(row, col).value = None


def write_output_row(sheet, row_idx: int, row: dict[str, Any]) -> None:
    if row_idx != 6:
        copy_row_style(sheet, 6, row_idx, max_col=14)
    values = [
        row["中文品名"],
        row["英文品名"],
        row["商品编码"],
        translate_material_to_english(row["材质"]),
        translate_usage_to_english(row["用途"]),
        row["箱数"],
        row["数量"],
        row["单位"],
        row["币制"],
        row["单价"],
        row["总价"],
        row["净重"],
        row["毛重"],
        row["原产国"],
    ]
    for col, value in enumerate(values, start=1):
        sheet.cell(row_idx, col).value = value


def copy_row_style(sheet, source_row: int, target_row: int, max_col: int) -> None:
    sheet.row_dimensions[target_row].height = sheet.row_dimensions[source_row].height
    for col in range(1, max_col + 1):
        source = sheet.cell(source_row, col)
        target = sheet.cell(target_row, col)
        if source.has_style:
            target._style = copy(source._style)
        if source.number_format:
            target.number_format = source.number_format
        if source.alignment:
            target.alignment = copy(source.alignment)


def extract_shipper(text: str) -> str:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    for idx, line in enumerate(lines):
        if "SUPPLY CHAIN" in line.upper() or "IMP. & EXP." in line.upper():
            return "\n".join(lines[idx : min(idx + 5, len(lines))])
    return ""


def extract_consignee(text: str) -> str:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    for idx, line in enumerate(lines):
        upper = line.upper()
        if "CORPORATION" in upper or "TRADE INC" in upper or "SOLUTIONS CORP" in upper:
            block = lines[idx : min(idx + 4, len(lines))]
            if not any("LOGISTICS" in item.upper() for item in block):
                return "\n".join(block)
    return ""


def extract_shipment_no(text: str) -> str:
    matches = re.findall(r"\b[A-Z]{3,5}\d{7,12}\b", text)
    return matches[0] if matches else ""


def extract_eta(text: str) -> str:
    match = re.search(r"\b(20\d{2})[/-](\d{1,2})[/-](\d{1,2})\b", text)
    if not match:
        match = re.search(r"\b([A-Z][a-z]{2})\.(\d{1,2}),\s*(20\d{2})\b", text)
        if not match:
            return ""
        months = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6, "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}
        return f"{match.group(3)}/{months.get(match.group(1), 1)}/{int(match.group(2))}"
    return f"{match.group(1)}/{int(match.group(2))}/{int(match.group(3))}"


def extract_number(pattern: str, text: str) -> Optional[float]:
    match = re.search(pattern, text, flags=re.IGNORECASE)
    if not match:
        return None
    return to_float(match.group(1))


def scale_integer(values: list[float], target_total: float) -> list[int]:
    target = int(round(target_total))
    base = sum(values)
    if not values:
        return []
    if not base:
        result = [0] * len(values)
        result[-1] = target
        return result
    exact = [value * target / base for value in values]
    floors = [math.floor(value) for value in exact]
    remainder = target - sum(floors)
    order = sorted(range(len(values)), key=lambda idx: exact[idx] - floors[idx], reverse=True)
    for idx in order[:remainder]:
        floors[idx] += 1
    return floors


def scale_decimal(values: list[float], target_total: float, digits: int) -> list[float]:
    base = sum(values)
    if not values:
        return []
    if not base:
        result = [0.0] * len(values)
        result[-1] = round(target_total, digits)
        return result
    scaled = [round(value * target_total / base, digits) for value in values]
    diff = round(target_total - sum(scaled), digits)
    if scaled and diff:
        scaled[-1] = round(scaled[-1] + diff, digits)
    return scaled


def normalize_header(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "")).lower()


def normalize_text(value: Any) -> str:
    text = str(value or "").lower()
    text = re.sub(r"[^a-z0-9\u4e00-\u9fff]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_bill_product_text(value: Any) -> str:
    text = normalize_text(value)
    words = []
    for word in text.split():
        if word == "mobile":
            continue
        if word == "stand":
            word = "holder"
        words.append(word)
    return " ".join(words)


def normalize_hs(value: Any) -> str:
    return re.sub(r"\D", "", str(value or ""))


def hs_cell_value(value: Any) -> int | str:
    hs = normalize_hs(value)
    if hs and not hs.startswith("0"):
        return int(hs)
    return hs


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        if math.isnan(value):
            return None
        return float(value)
    text = str(value).strip()
    if not text or text.lower() == "nan" or text.startswith("="):
        return None
    try:
        return float(text.replace(",", ""))
    except ValueError:
        return None


def manifest_to_public_dict(manifest: ManifestSummary) -> dict[str, Any]:
    data = asdict(manifest)
    data["items"] = data["items"][:80]
    return data


def bill_to_public_dict(bill: BillInfo) -> dict[str, Any]:
    return {
        "filename": bill.filename,
        "products": bill.products,
        "product_entries": [asdict(entry) for entry in bill.product_entries],
        "shipment_no": bill.shipment_no,
        "eta": bill.eta,
        "cartons": bill.cartons,
        "gross_weight": bill.gross_weight,
        "cbm": bill.cbm,
        "shipper": bill.shipper,
        "consignee": bill.consignee,
    }
