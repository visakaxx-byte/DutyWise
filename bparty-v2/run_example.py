from __future__ import annotations

import argparse
import asyncio
from pathlib import Path

from engine import build_clearance


async def async_main() -> None:
    parser = argparse.ArgumentParser(description="Run BParty V2 clearance generation")
    parser.add_argument("--manifest", required=True, help="清单 Excel 路径")
    parser.add_argument("--bill", required=True, help="运输提单 PDF 路径")
    parser.add_argument("--target-tax-amount", required=True, type=float, help="期望总税金金额")
    parser.add_argument("--target-item-count", required=True, type=int, help="最终生成条目数")
    parser.add_argument("--output-dir", default=str(Path(__file__).resolve().parent / "outputs"))
    args = parser.parse_args()

    result = await build_clearance(
        args.manifest,
        args.bill,
        args.output_dir,
        "auto",
        target_tax_amount=args.target_tax_amount,
        target_item_count=args.target_item_count,
    )
    print(result["output_file"])
    print(result["stats"])
    print("flow=", result["flow"])
    print("bill_categories=", result["bill_categories"])


if __name__ == "__main__":
    asyncio.run(async_main())
