"""
bparty-pipeline CLI 入口

用法:
    cd bparty-pipeline
    python pipeline.py --input 清单.xlsx --output result.xlsx --mode conservative
    python pipeline.py --input 清单.xlsx --output result.xlsx --mode aggressive --no-crawl
    python pipeline.py --input 清单.xlsx --dry-run

流水线步骤 (7步):
    1. DocumentParser.parse_files()     → 解析表格
    2. FieldMapper.map_fields()          → 映射到14列标准字段
    3. TaxRateCrawler.batch_search()     → 获取税率数据
    4. ProductClassifier.classify()      → 品名归类+合并
    5. CrossChapterOptimizer.optimize()  → 跨章HS优化
    6. DevaluationStrategy.apply()       → 低申报调整
    7. FileGenerator.generate()          → 生成输出xlsx
"""

import argparse
import asyncio
import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

# --- sys.path 注入: 复用 backend/ 代码 ---
# backend 内部使用 `from app.core.config import settings` 等相对导入，
# 需要将 backend/ 目录加入 sys.path 使 `app` 解析为 `backend/app/`。
# _BACKEND_ROOT 必须在 _HERE 之前，否则 bparty-pipeline/app.py 遮蔽 backend/app/ 包。
_HERE = Path(__file__).resolve().parent
_PROJECT_ROOT = _HERE.parent
_BACKEND_ROOT = _PROJECT_ROOT / "backend"

sys.path.insert(0, str(_HERE))           # strategy.*, pipeline (后搜索)
sys.path.insert(0, str(_BACKEND_ROOT))   # app.* → backend/app/* (先搜索)

import yaml

from app.services.parser import DocumentParser
from app.services.mapper import FieldMapper
from app.services.crawler import TaxRateCrawler
from app.services.generator import FileGenerator
from strategy.classifier import ProductClassifier
from strategy.cross_optimizer import CrossChapterOptimizer
from strategy.devaluation import DevaluationStrategy

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("bpipeline")


def _load_config(config_path: str) -> dict:
    """加载 config.yaml，返回完整配置字典"""
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def _resolve_mode_config(full_config: dict, mode: str) -> dict:
    """
    将 base + mode_overlay 合并为最终配置。
    conservative 使用顶层键；aggressive 在顶层基础上叠加 aggressive 键覆盖。
    """
    # 深拷贝基础配置（排除 aggressive 键）
    base = {}
    for k, v in full_config.items():
        if k == "aggressive":
            continue
        if isinstance(v, dict):
            base[k] = v.copy() if not any(isinstance(x, dict) for x in v.values()) else _deep_merge({}, v)
        else:
            base[k] = v

    if mode == "aggressive" and "aggressive" in full_config:
        return _deep_merge(base, full_config["aggressive"])
    return base


def _deep_merge(base: dict, overlay: dict) -> dict:
    """递归合并字典，overlay 覆盖 base"""
    result = base.copy()
    for k, v in overlay.items():
        if k in result and isinstance(result[k], dict) and isinstance(v, dict):
            result[k] = _deep_merge(result[k], v)
        else:
            result[k] = v
    return result


async def run_pipeline(
    input_path: str,
    output_path: str,
    mode_cfg: dict,
    *,
    use_crawler: bool = True,
    bl_products: list = None,
) -> dict:
    """
    执行完整流水线（async，兼容 FastAPI 和 CLI）。

    Args:
        bl_products: 提单上的品名列表，这些品名会在归类/合并时被保护为独立行

    返回:
        {
            "stats": {"total_items": N, "optimized_items": K, ...},
            "output_file": str,       # 生成的 xlsx 路径
            "optimization_logs": [...],
        }
    """
    stats = {}

    # ── Step 1: 解析 ──
    logger.info(f"[1/7] 解析文件: {input_path}")
    parser = DocumentParser()
    parsed = parser.parse_files([input_path])
    logger.info(f"    解析完成: {len(parsed['files'])} 个文件")

    # ── Step 2: 映射 ──
    logger.info("[2/7] 字段映射到14列标准字段")
    mapper = FieldMapper()
    mapping_result = mapper.map_fields(parsed)
    items = mapping_result["items"]
    logger.info(f"    映射完成: {len(items)} 条数据")

    if not items:
        raise ValueError("未能从输入文件中提取到有效数据（检查品名/HS编码等字段）")

    # ── Step 3: 税率查询 ──
    tax_data = {}
    if use_crawler:
        logger.info("[3/7] 查询HS编码税率")
        crawler = TaxRateCrawler()
        tax_data = await crawler.batch_search(items)
        logger.info(f"    税率查询完成: {len(tax_data)} 个编码")
    else:
        logger.info("[3/7] 跳过税率查询 (--no-crawl)")

    # ── Step 4: 品名归类 ──
    logger.info("[4/7] 品名归类与合并")
    classifier_cfg = mode_cfg.get("classifier", {})
    classifier = ProductClassifier(classifier_cfg)
    if bl_products:
        classifier.set_bl_products(bl_products)
        logger.info(f"    提单品名保护: {len(bl_products)} 个 → {bl_products}")
    items = classifier.classify(items)
    logger.info(f"    归类/合并后: {len(items)} 条数据")

    # ── Step 5: 跨章HS优化 ──
    logger.info("[5/7] 跨章HS编码优化")
    opt_cfg = mode_cfg.get("cross_optimizer", {})
    cross_opt = CrossChapterOptimizer(opt_cfg)
    opt_result = cross_opt.optimize(items, tax_data)
    items = opt_result["items"]
    optimization_logs = opt_result["optimization_logs"]
    stats["optimized_items"] = len(optimization_logs)
    logger.info(f"    优化完成: {stats['optimized_items']} 个商品重新归类")

    # ── Step 6: 低申报调整 ──
    logger.info("[6/7] 低申报策略调整")
    dev_cfg = mode_cfg.get("devaluation", {})
    deval = DevaluationStrategy(dev_cfg)
    items = deval.apply(items)
    logger.info("    低申报调整完成")

    # ── Step 7: 生成输出 ──
    logger.info("[7/7] 生成清关文件")
    output_dir = str(Path(output_path).parent) or "."
    template_dir = Path(__file__).resolve().parent / "templates"
    template_path = str(template_dir / "清关模板.xlsx")

    generator = FileGenerator()
    output_file = generator.generate(
        optimized_data=items,
        template_path=template_path,
        output_dir=output_dir,
        optimization_logs=optimization_logs,
    )
    logger.info(f"    输出文件: {output_file}")

    # ── 统计 ──
    stats["total_items"] = len(items)
    stats["output_file"] = output_file

    # 计算平均税率降幅
    reductions = []
    for log in optimization_logs:
        orig = _parse_rate(log.get("original_tax_rate", ""))
        new = _parse_rate(log.get("new_tax_rate", ""))
        if orig is not None and new is not None and orig > new:
            reductions.append(orig - new)
    stats["avg_tax_reduction"] = sum(reductions) / len(reductions) if reductions else 0
    stats["max_tax_reduction"] = max(reductions) if reductions else 0

    # 总金额
    total_value = sum(
        float(i.get("总价", 0) or 0)
        for i in items
    )
    stats["total_value_usd"] = round(total_value, 2)

    logger.info("=" * 50)
    logger.info(f"流水线完成!")
    logger.info(f"  总商品数:     {stats['total_items']}")
    logger.info(f"  优化商品数:   {stats['optimized_items']}")
    logger.info(f"  平均降税:     {stats['avg_tax_reduction']:.1f}%")
    logger.info(f"  最大降税:     {stats['max_tax_reduction']:.1f}%")
    logger.info(f"  申报总价:     ${stats['total_value_usd']:,.2f}")
    logger.info(f"  输出文件:     {output_file}")

    return {
        "stats": stats,
        "output_file": output_file,
        "optimization_logs": optimization_logs,
        "items": items,
    }


def _parse_rate(rate_str: str) -> Optional[float]:
    """解析税率字符串为数值"""
    if not rate_str:
        return None
    s = str(rate_str).lower().strip()
    if s in ("free", "0", "0%", ""):
        return 0.0
    try:
        return float(s.replace("%", ""))
    except (ValueError, TypeError):
        return None


# ── CLI ──────────────────────────────────────────────────────────────────


def main():
    ap = argparse.ArgumentParser(
        description="bparty-pipeline: 乙方清关流水线自动化",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python pipeline.py --input 清单.xlsx --output result.xlsx --mode conservative
  python pipeline.py --input 清单.xlsx --output result.xlsx --mode aggressive --no-crawl
  python pipeline.py --input 清单.xlsx --dry-run
        """,
    )
    ap.add_argument("--input", required=True, help="输入 Excel 文件路径")
    ap.add_argument("--output", default="./output.xlsx", help="输出 Excel 文件路径 (默认: ./output.xlsx)")
    ap.add_argument("--mode", default="conservative", choices=["conservative", "aggressive"])
    ap.add_argument("--config", default=None, help="config.yaml 路径 (默认: 同目录下的 config.yaml)")
    ap.add_argument("--no-crawl", action="store_true", help="跳过税率爬取（使用内置税率库）")
    ap.add_argument("--dry-run", action="store_true", help="仅输出统计，不生成文件")
    args = ap.parse_args()

    # 配置路径
    config_path = args.config or str(Path(__file__).resolve().parent / "config.yaml")
    if not Path(config_path).exists():
        logger.error(f"配置文件不存在: {config_path}")
        sys.exit(1)

    # 加载配置
    full_config = _load_config(config_path)
    mode_cfg = _resolve_mode_config(full_config, args.mode)
    logger.info(f"模式: {args.mode}")

    # 输入校验
    if not Path(args.input).exists():
        logger.error(f"输入文件不存在: {args.input}")
        sys.exit(1)

    try:
        result = asyncio.run(run_pipeline(
            input_path=args.input,
            output_path=args.output if not args.dry_run else "/tmp/bpipeline_dryrun.xlsx",
            mode_cfg=mode_cfg,
            use_crawler=not args.no_crawl,
        ))

        if args.dry_run:
            logger.info("--dry-run 模式: 跳过文件生成")
            # 仍然输出 stats
            print("\n统计摘要:")
            for k, v in result["stats"].items():
                print(f"  {k}: {v}")
    except Exception as e:
        logger.exception(f"流水线执行失败: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
