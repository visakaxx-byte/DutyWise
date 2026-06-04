from __future__ import annotations

import asyncio
import argparse
from pathlib import Path

from engine import build_clearance, target_tax_upper_bound


HERE = Path(__file__).resolve().parent


async def main() -> None:
    parser = argparse.ArgumentParser(description="Run a real LLM/crawler validation")
    parser.add_argument("--manifest", required=True, help="清单 Excel 路径")
    parser.add_argument("--bill", required=True, help="运输提单 PDF 路径")
    parser.add_argument("--target-tax-amount", required=True, type=float, help="期望总税金金额")
    parser.add_argument("--target-item-count", required=True, type=int, help="最终生成条目数")
    parser.add_argument("--output-dir", default=str(HERE / "outputs" / "validation"))
    args = parser.parse_args()

    result = await build_clearance(
        Path(args.manifest),
        Path(args.bill),
        Path(args.output_dir),
        "auto",
        target_tax_amount=args.target_tax_amount,
        target_item_count=args.target_item_count,
    )
    stats = result["stats"]
    assert stats["llm_used"] is False
    assert stats["crawler_used"] is True
    assert stats["output_rows"] == args.target_item_count
    assert abs(stats["bill_gross_weight"] - stats["output_gross_weight"]) < 0.01
    assert stats["estimated_tax_amount"] <= target_tax_upper_bound(args.target_tax_amount)
    print(result["output_file"])
    print(stats)
    print(result["flow"])


if __name__ == "__main__":
    asyncio.run(main())
