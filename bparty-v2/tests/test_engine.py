from __future__ import annotations

import unittest
from pathlib import Path
from types import SimpleNamespace

from engine import (
    BillInfo,
    BillProduct,
    ManifestWeightInfo,
    ManifestSummary,
    ProductCandidate,
    ProcessingOptions,
    SelectionRules,
    apply_bill_product_names,
    build_precheck,
    build_manifest_weight_context,
    build_output_rows,
    build_plausible_row_plans,
    certification_filter_reason,
    effective_tax_rate,
    ensure_candidate_plausibility_ranges,
    ensure_bill_products_present,
    estimate_precheck_summary,
    generate_valid_output_rows_with_llm,
    load_plausibility_ranges,
    load_replacement_candidates,
    normalize_manifest_weight_payload,
    normalize_bill_llm_products,
    normalize_bill_llm_product_entries,
    parse_non_exempt_additional_tax_rate,
    parse_tax_rate,
    parse_bill_products_with_llm,
    PlausibilityRange,
    qualify_single_candidate,
    qualify_bill_product_candidates,
    select_initial_candidates,
    tax_filter_reason,
    translate_material_to_english,
    translate_usage_to_english,
    normalize_generated_candidate,
    validate_llm_output_rows,
)
from crawler_client import parse_classification_results


def tax_result(rate: str = "3.4%", hs: str = "3924104000", certifications: list[str] | None = None) -> dict[str, dict]:
    return {
        hs: {
            "hs_code_us": hs,
            "tax_rate": rate,
            "certification_texts": certifications or [],
            "description_cn": "测试品名",
        }
    }


class FakeCrawler:
    def __init__(self, result: dict[str, dict]):
        self.result = result
        self.settings = SimpleNamespace(delay=0.0)

    async def search_product(self, product_name: str, material: str = "") -> dict[str, dict]:
        return self.result

    async def search(self, hs_code: str) -> dict[str, dict]:
        normalized = "".join(ch for ch in str(hs_code or "") if ch.isdigit())
        if len(normalized) == 10 and len(self.result) == 1:
            data = dict(next(iter(self.result.values())))
            data["hs_code_us"] = normalized
            return {normalized: data}
        return self.result


class RoutedFakeCrawler:
    def __init__(self, product_results: dict[str, dict[str, dict]], hs_results: dict[str, dict[str, dict]] | None = None):
        self.product_results = product_results
        self.hs_results = hs_results or {}
        self.settings = SimpleNamespace(delay=0.0)
        self.product_calls: list[tuple[str, str]] = []
        self.hs_calls: list[str] = []

    async def search_product(self, product_name: str, material: str = "") -> dict[str, dict]:
        self.product_calls.append((product_name, material))
        key = product_name.lower()
        if key not in self.product_results:
            raise RuntimeError(f"missing product result: {product_name}")
        return self.product_results[key]

    async def search(self, hs_code: str) -> dict[str, dict]:
        self.hs_calls.append(hs_code)
        if hs_code not in self.hs_results:
            raise RuntimeError(f"missing hs result: {hs_code}")
        return self.hs_results[hs_code]


class FakeBillParser:
    def __init__(self, payload: dict):
        self.payload = payload
        self.calls = 0

    async def chat_json(self, messages: list[dict[str, str]], *, temperature: float = 0.1) -> dict:
        self.calls += 1
        return self.payload


class QueueFakeLLM:
    def __init__(self, payloads: list[dict]):
        self.payloads = list(payloads)
        self.calls = 0

    async def chat_json(self, messages: list[dict[str, str]], *, temperature: float = 0.1) -> dict:
        self.calls += 1
        if not self.payloads:
            raise RuntimeError("no fake LLM payload left")
        return self.payloads.pop(0)


class RoutedFakeParser:
    def __init__(self, bill_payload: dict, manifest_payload: dict):
        self.bill_payload = bill_payload
        self.manifest_payload = manifest_payload

    async def chat_json(self, messages: list[dict[str, str]], *, temperature: float = 0.1) -> dict:
        content = "\n".join(message.get("content", "") for message in messages)
        if "total_weight_kg" in content:
            return self.manifest_payload
        return self.bill_payload


class TaxRateTests(unittest.TestCase):
    def test_parse_tax_rate_formats(self) -> None:
        self.assertEqual(parse_tax_rate("N/A"), None)
        self.assertAlmostEqual(parse_tax_rate(0.034), 0.034)
        self.assertAlmostEqual(parse_tax_rate("3.4%"), 0.034)
        self.assertAlmostEqual(parse_tax_rate("25%+10%"), 0.35)
        self.assertAlmostEqual(parse_tax_rate("20%"), 0.2)
        self.assertAlmostEqual(parse_tax_rate("20"), 0.2)
        self.assertAlmostEqual(parse_tax_rate("Free"), 0.0)

    def test_effective_tax_rate_adds_non_exempt_additional_tax(self) -> None:
        self.assertAlmostEqual(
            effective_tax_rate({"tax_rate": "3%", "additional_tax_rate": "7%"}),
            0.10,
        )
        self.assertAlmostEqual(
            effective_tax_rate({"tax_rate": "3%", "additional_tax_rate": "7%", "applicable_tax_rate": "12%"}),
            0.10,
        )
        self.assertAlmostEqual(
            effective_tax_rate({"additional_tax_rate": "7%", "applicable_tax_rate": "12%"}),
            0.07,
        )

    def test_additional_tax_rate_treats_conditional_exemption_as_exempt(self) -> None:
        self.assertAlmostEqual(
            parse_non_exempt_additional_tax_rate("无豁免7.5%；条件豁免10%"),
            0.075,
        )
        self.assertAlmostEqual(
            effective_tax_rate({"tax_rate": "3.4%", "additional_tax_rate": "无豁免7.5%；条件豁免10%"}),
            0.109,
        )
        self.assertAlmostEqual(
            effective_tax_rate(
                {
                    "tax_rate": "3.4%",
                    "additional_tax_details": [
                        {"label": "通用税率", "rate": "3.4%"},
                        {"label": "无豁免", "rate": "7.5%"},
                        {"label": "条件豁免", "rate": "10%"},
                    ],
                }
            ),
            0.109,
        )

    def test_parse_codeflagai_additional_tax_and_authentication(self) -> None:
        parsed = parse_classification_results(
            "9017800000",
            {
                "classificationResultList": [{"hsCode": "9017800000", "gName": "其他手用测量长度的器具"}],
                "classificationCodeList": [
                    {
                        "hsCode": "9017800000",
                        "importTariffRate": "5.3%",
                        "applicableTaxRate": "5.3%",
                        "additionalTaxRate": "25%+10%",
                        "authentication": "FD1: FDA data MAY BE required",
                    }
                ],
            },
        )
        item = parsed["9017800000"]
        self.assertEqual(item["additional_tax_rate"], "25%+10%")
        self.assertEqual(item["certification_texts"], ["FD1: FDA data MAY BE required"])


class CertificationRuleTests(unittest.TestCase):
    def setUp(self) -> None:
        self.rules = SelectionRules(allowed_certifications=["Lacey Act", "TSCA"])

    def test_allowed_certification_combinations(self) -> None:
        self.assertEqual(certification_filter_reason([], self.rules), "")
        self.assertEqual(certification_filter_reason(["May need Lacey Act"], self.rules), "")
        self.assertEqual(certification_filter_reason(["TSCA"], self.rules), "")
        self.assertEqual(certification_filter_reason(["Lacey Act and TSCA"], self.rules), "")

    def test_unknown_or_blocked_certification_is_filtered(self) -> None:
        self.assertIn("未允许", certification_filter_reason(["FDA"], self.rules))
        blocked = SelectionRules(
            allowed_certifications=["Lacey Act", "TSCA"],
            blocked_certifications=["FDA"],
        )
        self.assertIn("禁用认证", certification_filter_reason(["FDA and TSCA"], blocked))

    def test_tax_filter_combines_rate_and_certification_rules(self) -> None:
        good = {"hs_code_us": "3924104000", "tax_rate": "3.4%", "certification_texts": ["TSCA"]}
        self.assertEqual(tax_filter_reason(good, self.rules), "")
        high = {"hs_code_us": "3924104000", "tax_rate": "20%", "certification_texts": []}
        self.assertIn("不小于 20%", tax_filter_reason(high, self.rules))
        unknown = {"hs_code_us": "3924104000", "tax_rate": "3.4%", "certification_texts": ["FDA"]}
        self.assertIn("未允许", tax_filter_reason(unknown, self.rules))


class ReplacementWorkbookTests(unittest.TestCase):
    def test_load_replacement_candidates_from_known_sheets(self) -> None:
        path = Path(__file__).resolve().parents[1] / "docs/海关编码查找.xlsx"
        candidates = load_replacement_candidates(path)
        self.assertGreater(len(candidates), 50)
        common = next(item for item in candidates if item.zh == "托盘")
        self.assertEqual(common.en, "Pallet")
        self.assertEqual(common.hs, "3924104000")
        self.assertGreater(common.unit_price, 0)
        self.assertTrue(any(item.source_label.endswith("/20260330") for item in candidates))

    def test_load_plausibility_ranges_from_known_sheets(self) -> None:
        path = Path(__file__).resolve().parents[1] / "docs/海关编码查找.xlsx"
        ranges = load_plausibility_ranges(path)

        towel = ranges[("毛巾", "毛巾", "6302932000")]
        self.assertAlmostEqual(towel.unit_price_min, 0.5)
        self.assertAlmostEqual(towel.unit_price_max, 0.8)
        self.assertAlmostEqual(towel.kg_per_pc_min, 0.5)
        self.assertAlmostEqual(towel.kg_per_pc_max, 0.9)
        self.assertAlmostEqual(towel.ctns_min, 40)
        self.assertAlmostEqual(towel.ctns_max, 60)
        self.assertAlmostEqual(towel.qty_per_ctn_min, 6)
        self.assertAlmostEqual(towel.qty_per_ctn_max, 10)

        pallet = ranges[("托盘", "pallet", "3924104000")]
        self.assertAlmostEqual(pallet.kg_per_ctn_min, 5.1746, places=3)
        self.assertAlmostEqual(pallet.kg_per_ctn_max, 32.0)
        self.assertAlmostEqual(pallet.kg_per_pc_min, 0.265, places=3)
        self.assertAlmostEqual(pallet.kg_per_pc_max, 1.88, places=2)


class BillProductParsingTests(unittest.IsolatedAsyncioTestCase):
    async def test_llm_parser_extracts_products_and_ignores_broken_shipped_on_board_phrase(self) -> None:
        text = """
        STORAGE BAG HS CODE:420222
        PLASTIC SHELL HS CODE:392690
        S
        HIPPEDON BOARD:
        MAY.06 2026
        """
        parser = FakeBillParser(
            {
                "products": [
                    {"name": "STORAGE BAG", "hs_code_hint": "420222", "evidence": "STORAGE BAG HS CODE:420222", "confidence": 0.98},
                    {"name": "PLASTIC SHELL", "hs_code_hint": "392690", "evidence": "PLASTIC SHELL HS CODE:392690", "confidence": 0.97},
                ],
                "ignored_phrases": [{"text": "HIPPEDON BOARD", "reason": "broken shipped on board field"}],
            }
        )
        cache = {"product": {}, "hs": {}, "bill": {}}
        self.assertEqual(await parse_bill_products_with_llm(text, parser, cache), ["STORAGE BAG", "PLASTIC SHELL"])
        self.assertEqual(await parse_bill_products_with_llm(text, parser, cache), ["STORAGE BAG", "PLASTIC SHELL"])
        self.assertEqual(parser.calls, 1)

    def test_llm_products_are_validated_by_code(self) -> None:
        payload = {
            "products": [
                {"name": "STORAGE BAG HS CODE:420222", "confidence": 0.99},
                {"name": "HIPPEDON BOARD", "confidence": 0.99},
                {"name": "LOW CONFIDENCE ITEM", "confidence": 0.2},
            ]
        }
        self.assertEqual(normalize_bill_llm_products(payload), ["STORAGE BAG"])

    def test_llm_product_entries_preserve_hs_hint(self) -> None:
        payload = {
            "products": [
                {
                    "name": "STORAGE BAG",
                    "hs_code_hint": "420222",
                    "evidence": "STORAGE BAG HS CODE:420222",
                    "confidence": 0.99,
                }
            ]
        }
        entries = normalize_bill_llm_product_entries(payload)
        self.assertEqual(entries[0].name, "STORAGE BAG")
        self.assertEqual(entries[0].hs_code_hint, "420222")
        self.assertEqual(entries[0].evidence, "STORAGE BAG HS CODE:420222")


class ManifestWeightParsingTests(unittest.TestCase):
    def test_manifest_weight_payload_is_normalized(self) -> None:
        payload = {
            "total_weight_kg": "3077.7",
            "source": "Sheet1.总毛重KGS",
            "evidence": "总毛重KGS 列求和",
            "confidence": 0.93,
        }
        info = normalize_manifest_weight_payload(payload)
        self.assertEqual(info.total_weight_kg, 3077.7)
        self.assertEqual(info.source, "Sheet1.总毛重KGS")
        self.assertEqual(info.evidence, "总毛重KGS 列求和")
        self.assertEqual(info.confidence, 0.93)


class PrecheckTests(unittest.IsolatedAsyncioTestCase):
    async def test_precheck_parses_files_and_returns_green_summary(self) -> None:
        root = Path(__file__).resolve().parents[1]
        result = await build_precheck(
            root / "docs/example04/NJPE5C277600 清单.xlsx",
            root / "docs/提单/海运MATS1895441000.pdf",
            target_tax_amount=100,
            target_item_count=3,
            crawler=FakeCrawler(tax_result("3.4%")),
            bill_parser=FakeBillParser({"products": [{"name": "SILICONE COASTER", "confidence": 0.99}]}),
            manifest_parser=FakeBillParser(
                {
                    "total_weight_kg": 280,
                    "source": "test",
                    "evidence": "test weight",
                    "confidence": 0.99,
                }
            ),
            manifest_sample_limit=2,
            replacement_sample_limit=1,
        )
        precheck = result["precheck"]
        self.assertEqual(precheck["status"], "green")
        self.assertTrue(precheck["can_start"])
        self.assertGreater(precheck["manifest_rows"], 0)
        self.assertGreater(precheck["estimated_seconds"], 0)
        self.assertEqual(precheck["qualified_bill_products"], 1)
        self.assertIn("product", result["query_cache"])

    async def test_precheck_low_pass_rate_is_red(self) -> None:
        root = Path(__file__).resolve().parents[1]
        result = await build_precheck(
            root / "docs/example04/NJPE5C277600 清单.xlsx",
            root / "docs/提单/海运MATS1895441000.pdf",
            target_tax_amount=100,
            target_item_count=8,
            crawler=FakeCrawler(tax_result("20%")),
            bill_parser=FakeBillParser({"products": [{"name": "SILICONE COASTER", "confidence": 0.99}]}),
            manifest_parser=FakeBillParser(
                {
                    "total_weight_kg": 280,
                    "source": "test",
                    "evidence": "test weight",
                    "confidence": 0.99,
                }
            ),
            manifest_sample_limit=2,
            replacement_sample_limit=1,
        )
        precheck = result["precheck"]
        self.assertEqual(precheck["status"], "red")
        self.assertFalse(precheck["can_start"])
        self.assertIn("预计合格品名", "；".join(precheck["risk_reasons"]))

    def test_precheck_missing_manifest_weight_is_blocked(self) -> None:
        summary = estimate_precheck_summary(
            options=ProcessingOptions(target_tax_amount=100, target_item_count=3),
            manifest=ManifestSummary("input.xlsx", 3, 3, 0, 100, []),
            bill=BillInfo("bill.pdf", "", [], gross_weight=30),
            manifest_candidate_count=10,
            replacement_candidate_count=10,
            manifest_sample_count=0,
            replacement_sample_count=0,
            sampled_manifest=[],
            sampled_replacements=[],
            manifest_filtered=[],
            replacement_filtered=[],
            bill_required=[],
            bill_filtered=[],
            sample_elapsed=0,
            crawler_delay=0,
            early_reasons=["清单 Excel 未识别到有效总重量"],
        )
        self.assertEqual(summary["status"], "red")
        self.assertFalse(summary["can_start"])

    async def test_precheck_uses_excel_weight_and_ignores_missing_bill_weight(self) -> None:
        root = Path(__file__).resolve().parents[1]
        result = await build_precheck(
            root / "docs/example04/NJPE5C277600 清单.xlsx",
            root / "docs/提单/海运MATS1895441000.pdf",
            target_tax_amount=100,
            target_item_count=3,
            crawler=FakeCrawler(tax_result("3.4%")),
            bill_parser=FakeBillParser({"products": [{"name": "SILICONE COASTER", "confidence": 0.99}]}),
            manifest_parser=FakeBillParser(
                {
                    "total_weight_kg": 3077.7,
                    "source": "Sheet1.总毛重KGS",
                    "evidence": "总毛重KGS 列求和",
                    "confidence": 0.99,
                }
            ),
            manifest_sample_limit=1,
            replacement_sample_limit=1,
        )
        precheck = result["precheck"]
        self.assertTrue(precheck["can_start"])
        self.assertEqual(precheck["manifest_total_weight"], 3077.7)
        self.assertEqual(result["manifest"]["weight_source"], "Sheet1.总毛重KGS")

    def test_precheck_target_over_candidate_capacity_is_blocked(self) -> None:
        summary = estimate_precheck_summary(
            options=ProcessingOptions(target_tax_amount=100, target_item_count=30),
            manifest=ManifestSummary("input.xlsx", 3, 3, 30, 100, []),
            bill=BillInfo("bill.pdf", "", [], gross_weight=30),
            manifest_candidate_count=2,
            replacement_candidate_count=3,
            manifest_sample_count=0,
            replacement_sample_count=0,
            sampled_manifest=[],
            sampled_replacements=[],
            manifest_filtered=[],
            replacement_filtered=[],
            bill_required=[],
            bill_filtered=[],
            sample_elapsed=0,
            crawler_delay=0,
            early_reasons=["目标条目数 30 超过客户候选和替换候选总数 5"],
        )
        self.assertEqual(summary["status"], "red")
        self.assertFalse(summary["can_start"])

    async def test_precheck_blocks_when_bill_products_exceed_target_count(self) -> None:
        root = Path(__file__).resolve().parents[1]
        result = await build_precheck(
            root / "docs/example04/NJPE5C277600 清单.xlsx",
            root / "docs/提单/海运MATS1895441000.pdf",
            target_tax_amount=100,
            target_item_count=1,
            crawler=FakeCrawler(tax_result("3.4%")),
            bill_parser=FakeBillParser(
                {
                    "products": [
                        {"name": "STORAGE BAG", "confidence": 0.99},
                        {"name": "PLASTIC SHELL", "confidence": 0.99},
                    ]
                }
            ),
            manifest_parser=FakeBillParser(
                {
                    "total_weight_kg": 280,
                    "source": "test",
                    "evidence": "test weight",
                    "confidence": 0.99,
                }
            ),
            manifest_sample_limit=0,
            replacement_sample_limit=0,
        )
        precheck = result["precheck"]
        self.assertEqual(precheck["status"], "red")
        self.assertFalse(precheck["can_start"])
        self.assertIn("提单品类 2 个超过目标输出行数 1", "；".join(precheck["risk_reasons"]))


class OutputOptimizationTests(unittest.TestCase):
    def test_material_and_usage_are_translated_to_english(self) -> None:
        material = translate_material_to_english("铁+塑料/Iron + Plastic")
        usage = translate_usage_to_english("测量戒指大小、ring sizer")
        self.assertNotRegex(material, r"[\u4e00-\u9fff]")
        self.assertIn("Iron", material)
        self.assertIn("Plastic", material)
        self.assertNotRegex(usage, r"[\u4e00-\u9fff]")
        self.assertIn("Ring Size Measurement", usage)
        self.assertIn("Ring Sizer", usage)

    def test_output_rows_match_count_weight_and_tax_target(self) -> None:
        candidates = [
            ProductCandidate(
                source="manifest",
                source_label="input.xlsx",
                zh=f"品名{i}",
                en=f"Item {i}",
                hs="3924104000",
                material="Plastic",
                usage="HOME",
                ctns=10,
                qty=100,
                unit_price=2,
                real_weight=100,
                gross_weight=100,
                base_tax_rate=0.034,
                effective_tax_rate=0.234,
                plausibility_range=PlausibilityRange(
                    kg_per_ctn_min=5,
                    kg_per_ctn_max=20,
                    kg_per_pc_min=0.5,
                    kg_per_pc_max=2,
                    unit_price_min=0.5,
                    unit_price_max=10,
                    qty_per_ctn_min=5,
                    qty_per_ctn_max=20,
                    source="test range",
                ),
            )
            for i in range(3)
        ]
        candidates[0].material = "塑料 Plastic"
        candidates[0].usage = "装饰/Decoration"
        bill = BillInfo(filename="bill.pdf", raw_text="", products=[], gross_weight=300, cartons=30)
        manifest = ManifestSummary(
            filename="input.xlsx",
            row_count=3,
            total_ctns=30,
            total_real_weight=280,
            total_declared_value=1000,
            categories=[],
        )
        options = ProcessingOptions(target_tax_amount=120, target_item_count=3)

        rows = build_output_rows(candidates, manifest, bill, options)
        self.assertEqual(len(rows), 3)
        self.assertEqual(rows[0]["材质"], "Plastic")
        self.assertEqual(rows[0]["用途"], "Decoration")
        self.assertAlmostEqual(sum(row["毛重"] for row in rows), 280, places=2)
        total_tax = sum(row["总价"] * row["综合税率"] for row in rows)
        self.assertLessEqual(abs(total_tax - 120), max(1, 120 * 0.01))

    def test_generated_candidates_do_not_inherit_historical_weight_constraints(self) -> None:
        candidate = ProductCandidate(
            source="replacement",
            source_label="海关编码查找.xlsx/常用1",
            zh="托盘",
            en="Pallet",
            hs="3924104000",
            material="Plastic",
            usage="HOME",
            ctns=12,
            qty=360,
            unit_price=0.8,
            real_weight=192,
            gross_weight=192,
            base_tax_rate=0.034,
            effective_tax_rate=0.034,
        )
        self.assertEqual(candidate.source, "replacement")
        self.assertEqual(candidate.unit_price, 0.8)
        self.assertEqual(candidate.qty, 360)
        self.assertEqual(candidate.gross_weight, 192)

    def test_output_rows_require_target_item_count(self) -> None:
        candidates = [
            ProductCandidate(
                source="manifest",
                source_label="input.xlsx",
                zh="品名",
                en="Item",
                hs="3924104000",
                material="Plastic",
                usage="HOME",
                ctns=10,
                qty=100,
                real_weight=100,
                gross_weight=100,
                base_tax_rate=0.034,
            )
        ]
        manifest = ManifestSummary("input.xlsx", 1, 10, 100, 100, [])
        bill = BillInfo("bill.pdf", "", [], gross_weight=300, cartons=10)
        options = ProcessingOptions(target_tax_amount=30, target_item_count=2)
        with self.assertRaisesRegex(RuntimeError, "输出行数必须等于 2"):
            build_output_rows(candidates, manifest, bill, options)

    def test_plausible_row_plans_obey_weight_and_price_ranges(self) -> None:
        candidates = [
            ProductCandidate(
                source="manifest",
                source_label="input.xlsx",
                zh="托盘",
                en="Pallet",
                hs="3924104000",
                material="Plastic",
                usage="HOME",
                ctns=12,
                qty=360,
                unit_price=0.8,
                real_weight=192,
                gross_weight=192,
                base_tax_rate=0.034,
            ),
            ProductCandidate(
                source="manifest",
                source_label="input.xlsx",
                zh="垫子",
                en="Mat",
                hs="3924901050",
                material="Plastic",
                usage="HOME",
                ctns=13,
                qty=220,
                unit_price=1.2,
                real_weight=300,
                gross_weight=300,
                base_tax_rate=0.033,
            ),
        ]
        manifest = ManifestSummary("input.xlsx", 2, 25, 492, 912, [])
        bill = BillInfo("bill.pdf", "", [], gross_weight=1200, cartons=25)
        options = ProcessingOptions(target_tax_amount=37, target_item_count=2)
        ranges = load_plausibility_ranges(Path(__file__).resolve().parents[1] / "docs/海关编码查找.xlsx")

        plans = build_plausible_row_plans(
            selected=candidates,
            manifest=manifest,
            bill=bill,
            options=options,
            plausibility_ranges=ranges,
        )

        self.assertAlmostEqual(sum(plan.gross_weight for plan in plans), 492, places=2)
        total_tax = sum(plan.total_value * (plan.candidate.effective_tax_rate or plan.candidate.base_tax_rate) for plan in plans)
        self.assertLessEqual(abs(total_tax - 37), max(1, 37 * 0.01))
        for plan in plans:
            kg_per_ctn = plan.gross_weight / plan.ctns
            kg_per_pc = plan.gross_weight / plan.qty
            self.assertGreaterEqual(kg_per_ctn, plan.plausibility.kg_per_ctn_min)
            self.assertLessEqual(kg_per_ctn, plan.plausibility.kg_per_ctn_max)
            self.assertGreaterEqual(kg_per_pc, plan.plausibility.kg_per_pc_min)
            self.assertLessEqual(kg_per_pc, plan.plausibility.kg_per_pc_max)
            self.assertGreaterEqual(plan.unit_price, plan.plausibility.unit_price_min)
            self.assertLessEqual(plan.unit_price, plan.plausibility.unit_price_max)

    def test_plausible_row_plans_block_impossible_weight(self) -> None:
        candidates = [
            ProductCandidate(
                source="manifest",
                source_label="input.xlsx",
                zh="塑料花",
                en="Plastic flowers",
                hs="6702104000",
                material="Plastic",
                usage="HOME",
                ctns=1,
                qty=100,
                unit_price=0.15,
                real_weight=1,
                gross_weight=1,
                base_tax_rate=0.034,
            )
        ]
        manifest = ManifestSummary("input.xlsx", 1, 1, 500, 100, [])
        bill = BillInfo("bill.pdf", "", [], gross_weight=500, cartons=1)
        options = ProcessingOptions(target_tax_amount=10, target_item_count=1)
        with self.assertRaisesRegex(RuntimeError, "合理单箱重量上限"):
            build_plausible_row_plans(
                selected=candidates,
                manifest=manifest,
                bill=bill,
                options=options,
                plausibility_ranges={},
            )

    def test_plausible_row_plans_block_impossible_tax(self) -> None:
        candidates = [
            ProductCandidate(
                source="manifest",
                source_label="input.xlsx",
                zh="托盘",
                en="Pallet",
                hs="3924104000",
                material="Plastic",
                usage="HOME",
                ctns=12,
                qty=360,
                unit_price=0.8,
                real_weight=192,
                gross_weight=192,
                base_tax_rate=0.034,
            )
        ]
        manifest = ManifestSummary("input.xlsx", 1, 12, 192, 288, [])
        bill = BillInfo("bill.pdf", "", [], gross_weight=192, cartons=12)
        options = ProcessingOptions(target_tax_amount=1000, target_item_count=1)
        ranges = load_plausibility_ranges(Path(__file__).resolve().parents[1] / "docs/海关编码查找.xlsx")
        with self.assertRaisesRegex(RuntimeError, "无法贴近期望税金"):
            build_plausible_row_plans(
                selected=candidates,
                manifest=manifest,
                bill=bill,
                options=options,
                plausibility_ranges=ranges,
            )


class BillProductCoverageTests(unittest.IsolatedAsyncioTestCase):
    async def test_replacement_candidate_keeps_original_hs_when_product_query_returns_unrelated_low_tax_hs(self) -> None:
        candidate = ProductCandidate(
            source="replacement",
            source_label="海关编码查找.xlsx/常用1",
            zh="电动钻",
            en="Electric drill",
            hs="8467210010",
            material="Plastic",
            usage="HOME",
            ctns=40,
            qty=120,
            unit_price=2.2,
            gross_weight=285.3,
        )
        crawler = RoutedFakeCrawler(
            product_results={
                "电动钻": tax_result("Free", hs="8508110000"),
            },
            hs_results={
                "8467210010": tax_result("1.7%", hs="8467210010"),
            },
        )

        result = await qualify_single_candidate(
            crawler,
            candidate,
            SelectionRules(allowed_certifications=["Lacey Act", "TSCA"]),
            query_cache={},
        )

        self.assertFalse(result.filter_reason)
        self.assertEqual(result.hs, "8467210010")
        self.assertEqual(result.tax_match_source, "replacement_hs")

    async def test_replacement_candidate_rejects_product_query_when_original_hs_has_no_result(self) -> None:
        candidate = ProductCandidate(
            source="replacement",
            source_label="海关编码查找.xlsx/常用1",
            zh="电动钻",
            en="Electric drill",
            hs="8467210010",
            material="Plastic",
            usage="HOME",
            ctns=40,
            qty=120,
            unit_price=2.2,
            gross_weight=285.3,
        )
        crawler = RoutedFakeCrawler(
            product_results={
                "电动钻": tax_result("Free", hs="8508110000"),
            },
            hs_results={
                "8467210010": {},
            },
        )

        result = await qualify_single_candidate(
            crawler,
            candidate,
            SelectionRules(allowed_certifications=["Lacey Act", "TSCA"]),
            query_cache={},
        )

        self.assertIn("返回 HS 与替换清单原始 HS 不一致", result.filter_reason)
        self.assertEqual(result.hs, "8467210010")

    async def test_bill_product_from_manifest_is_selected_before_other_items(self) -> None:
        bill = BillInfo(
            filename="bill.pdf",
            raw_text="",
            products=["PLASTIC MOBILE PHONE STAND"],
            gross_weight=150,
            cartons=15,
        )
        options = ProcessingOptions(target_tax_amount=51, target_item_count=3)
        rules = SelectionRules(allowed_certifications=["Lacey Act", "TSCA"])
        other_candidates = [
            ProductCandidate(
                source="manifest",
                source_label="input.xlsx",
                zh=f"其他品名{i}",
                en=f"Other Item {i}",
                hs="3924104000",
                material="Plastic",
                usage="HOME",
                ctns=50,
                qty=500,
                real_weight=500,
                gross_weight=500,
                unit_price=2,
                base_tax_rate=0.034,
                effective_tax_rate=0.034,
                tax_match_source="product",
                plausibility_range=PlausibilityRange(
                    kg_per_ctn_min=2,
                    kg_per_ctn_max=20,
                    kg_per_pc_min=0.2,
                    kg_per_pc_max=2,
                    unit_price_min=0.5,
                    unit_price_max=10,
                    qty_per_ctn_min=5,
                    qty_per_ctn_max=30,
                    source="test range",
                ),
            )
            for i in range(3)
        ]
        phone_candidate = ProductCandidate(
            source="manifest",
            source_label="input.xlsx",
            zh="手机支架",
            en="Phone holder",
            hs="3926100000",
            material="Plastic",
            usage="HOME",
            ctns=1,
            qty=20,
            unit_price=2,
            real_weight=10,
            gross_weight=10,
            base_tax_rate=0.034,
            effective_tax_rate=0.034,
            tax_match_source="product",
            plausibility_range=PlausibilityRange(
                kg_per_ctn_min=2,
                kg_per_ctn_max=20,
                kg_per_pc_min=0.2,
                kg_per_pc_max=2,
                unit_price_min=0.5,
                unit_price_max=10,
                qty_per_ctn_min=5,
                qty_per_ctn_max=30,
                source="test range",
            ),
        )
        manifest_candidates = [*other_candidates, phone_candidate]

        bill_required, bill_filtered = await qualify_bill_product_candidates(
            FakeCrawler(tax_result("3.4%", hs="3926100000")),
            bill,
            manifest_candidates,
            [],
            rules,
            options,
            query_cache={},
        )
        self.assertEqual(bill_filtered, [])
        self.assertEqual(bill_required[0].zh, "手机支架")
        self.assertEqual(bill_required[0].hs, "3926100000")

        selected = select_initial_candidates(manifest_candidates, bill_required, rules, options.target_item_count)
        self.assertEqual(len(selected), 3)
        self.assertEqual(selected[0].zh, "手机支架")

        manifest = ManifestSummary(
            filename="input.xlsx",
            row_count=3,
            total_ctns=15,
            total_real_weight=150,
            total_declared_value=1000,
            categories=[],
        )
        rows = build_output_rows(selected, manifest, bill, options)
        apply_bill_product_names(rows, bill.products, selected)
        ensure_bill_products_present(rows, bill.products, selected)
        self.assertTrue(any(row["中文品名"] == "塑料手机支架" for row in rows))

    async def test_bill_product_without_matching_hs_candidate_uses_llm_query_terms(self) -> None:
        bill = BillInfo(
            filename="bill.pdf",
            raw_text="",
            products=["PLASTIC SHELL"],
            gross_weight=150,
            cartons=15,
        )
        options = ProcessingOptions(target_tax_amount=90, target_item_count=2)
        crawler = RoutedFakeCrawler(
            product_results={
                "plastic shell": {},
                "plastic enclosure": tax_result("3.4%", hs="3926909985"),
            }
        )
        llm = QueueFakeLLM(
            [
                {
                    "queries": [
                        {"product_name": "Plastic enclosure", "material": "Plastic", "reason": "same product category"},
                    ]
                }
            ]
        )
        rules = SelectionRules(allowed_certifications=["Lacey Act", "TSCA"])

        bill_required, bill_filtered = await qualify_bill_product_candidates(
            crawler,
            bill,
            [],
            [],
            rules,
            options,
            query_cache={},
            llm=llm,
        )

        self.assertEqual(len(bill_required), 1)
        self.assertEqual(bill_required[0].zh, "PLASTIC SHELL")
        self.assertEqual(bill_required[0].hs, "3926909985")
        self.assertEqual(bill_required[0].tax_match_source, "llm_query")
        self.assertGreaterEqual(len(bill_filtered), 1)

    async def test_bill_product_with_hs_hint_is_qualified_from_codeflagai_hs_query(self) -> None:
        bill = BillInfo(
            filename="bill.pdf",
            raw_text="",
            products=["STORAGE BAG"],
            gross_weight=150,
            cartons=15,
            product_entries=[
                BillProduct(
                    name="STORAGE BAG",
                    hs_code_hint="420222",
                    evidence="STORAGE BAG HS CODE:420222",
                    confidence=0.99,
                )
            ],
        )
        options = ProcessingOptions(target_tax_amount=90, target_item_count=3)
        rules = SelectionRules(allowed_certifications=["Lacey Act", "TSCA"])

        bill_required, bill_filtered = await qualify_bill_product_candidates(
            FakeCrawler(tax_result("Free", hs="4202221000")),
            bill,
            [],
            [],
            rules,
            options,
            query_cache={},
        )

        self.assertEqual(bill_filtered, [])
        self.assertEqual(len(bill_required), 1)
        self.assertEqual(bill_required[0].source, "bill")
        self.assertEqual(bill_required[0].zh, "STORAGE BAG")
        self.assertEqual(bill_required[0].hs, "4202221000")
        self.assertEqual(bill_required[0].tax_match_source, "bill_hs")
        self.assertEqual(bill_required[0].base_tax_rate, 0.0)

    async def test_llm_plausibility_range_is_used_when_candidate_lacks_ranges(self) -> None:
        candidate = ProductCandidate(
            source="bill",
            source_label="bill.pdf",
            zh="HANDHELD GARMENT STEAMER",
            en="Handheld Garment Steamer",
            hs="8451300000",
            material="Plastic",
            usage="HOME",
            ctns=1,
            qty=1,
            gross_weight=1,
            base_tax_rate=0.034,
            effective_tax_rate=0.034,
            tax_match_source="bill_product",
        )
        llm = QueueFakeLLM(
            [
                {
                    "unit_price_min": 8,
                    "unit_price_max": 18,
                    "kg_per_pc_min": 0.6,
                    "kg_per_pc_max": 1.5,
                    "qty_per_ctn_min": 4,
                    "qty_per_ctn_max": 12,
                    "kg_per_ctn_min": 4,
                    "kg_per_ctn_max": 18,
                    "confidence": 0.9,
                    "basis": "small appliance",
                }
            ]
        )
        selected, used = await ensure_candidate_plausibility_ranges(
            [candidate],
            ManifestSummary("input.xlsx", 1, 10, 100, 100, []),
            BillInfo("bill.pdf", "", ["HANDHELD GARMENT STEAMER"]),
            {},
            llm,
            {},
        )
        self.assertTrue(used)
        self.assertEqual(selected[0].plausibility_range.unit_price_min, 8)
        self.assertIn("LLM合理性估算", selected[0].plausibility_range.source)

    async def test_incomplete_replacement_range_uses_llm_instead_of_placeholder_fields(self) -> None:
        candidate = ProductCandidate(
            source="replacement",
            source_label="海关编码查找.xlsx/20260330",
            zh="加湿器",
            en="Humidifier",
            hs="8509809000",
            material="Plastic",
            usage="HOME",
            ctns=50,
            qty=1000,
            unit_price=1,
            gross_weight=1,
            base_tax_rate=0.034,
            effective_tax_rate=0.034,
            tax_match_source="product",
        )
        known = {
            ("加湿器", "humidifier", "8509809000"): PlausibilityRange(
                kg_per_ctn_min=None,
                kg_per_ctn_max=None,
                kg_per_pc_min=0.5,
                kg_per_pc_max=1.5,
                unit_price_min=8,
                unit_price_max=20,
                qty_per_ctn_min=4,
                qty_per_ctn_max=12,
                source="incomplete test range",
            )
        }
        llm = QueueFakeLLM(
            [
                {
                    "unit_price_min": 8,
                    "unit_price_max": 20,
                    "kg_per_pc_min": 0.5,
                    "kg_per_pc_max": 1.5,
                    "qty_per_ctn_min": 4,
                    "qty_per_ctn_max": 12,
                    "kg_per_ctn_min": 4,
                    "kg_per_ctn_max": 18,
                    "confidence": 0.92,
                    "basis": "small electric appliance",
                }
            ]
        )

        selected, used = await ensure_candidate_plausibility_ranges(
            [candidate],
            ManifestSummary("input.xlsx", 1, 50, 500, 1000, []),
            BillInfo("bill.pdf", "", []),
            known,
            llm,
            {},
        )

        self.assertTrue(used)
        self.assertEqual(selected[0].plausibility_range.kg_per_ctn_min, 4)
        self.assertIn("small electric appliance", selected[0].plausibility_range.source)

    async def test_llm_output_rejects_mechanical_quantity_and_weight_distribution(self) -> None:
        candidates = []
        for idx in range(8):
            candidates.append(
                ProductCandidate(
                    source="replacement",
                    source_label="test",
                    zh=f"品名{idx}",
                    en=f"Item {idx}",
                    hs=f"39269099{idx:02d}",
                    material="Plastic",
                    usage="HOME",
                    ctns=10,
                    qty=100,
                    unit_price=2,
                    gross_weight=50,
                    base_tax_rate=0.1,
                    effective_tax_rate=0.1,
                    tax_match_source="product",
                    plausibility_range=PlausibilityRange(
                        kg_per_ctn_min=2,
                        kg_per_ctn_max=20,
                        kg_per_pc_min=0.2,
                        kg_per_pc_max=2,
                        unit_price_min=1,
                        unit_price_max=5,
                        qty_per_ctn_min=5,
                        qty_per_ctn_max=20,
                        source="test range",
                    ),
                )
            )
        llm = QueueFakeLLM(
            [
                {
                    "rows": [
                        {"candidate_index": idx, "箱数": 10, "数量": 100, "单价": 1.25, "毛重": 47}
                        for idx in range(8)
                    ]
                },
                {
                    "rows": [
                        {"candidate_index": 0, "箱数": 8, "数量": 50, "单价": 1.1, "毛重": 20},
                        {"candidate_index": 1, "箱数": 9, "数量": 60, "单价": 1.2, "毛重": 26},
                        {"candidate_index": 2, "箱数": 10, "数量": 70, "单价": 1.3, "毛重": 32},
                        {"candidate_index": 3, "箱数": 11, "数量": 80, "单价": 1.4, "毛重": 39},
                        {"candidate_index": 4, "箱数": 12, "数量": 90, "单价": 1.5, "毛重": 48},
                        {"candidate_index": 5, "箱数": 13, "数量": 100, "单价": 1.6, "毛重": 58},
                        {"candidate_index": 6, "箱数": 14, "数量": 110, "单价": 1.7, "毛重": 70},
                        {"candidate_index": 7, "箱数": 15, "数量": 120, "单价": 1.5667, "毛重": 83},
                    ]
                },
            ]
        )
        manifest = ManifestSummary("input.xlsx", 8, 92, 376, 1000, [])
        bill = BillInfo("bill.pdf", "", [])
        options = ProcessingOptions(target_tax_amount=100, target_item_count=8)

        rows, attempts, feedback = await generate_valid_output_rows_with_llm(llm, candidates, manifest, bill, options)

        self.assertEqual(attempts, 2)
        self.assertIn("数量分布过于机械", feedback[0])
        self.assertEqual(len(rows), 8)
        self.assertLessEqual(abs(sum(row["总价"] * row["综合税率"] for row in rows) - 100), 1)

    def test_llm_output_accepts_tax_up_to_100_under_target_without_exceeding(self) -> None:
        candidates = [
            ProductCandidate(
                source="replacement",
                source_label="test",
                zh="品名",
                en="Item",
                hs="3926909985",
                material="Plastic",
                usage="HOME",
                ctns=10,
                qty=100,
                unit_price=85.133,
                gross_weight=100,
                base_tax_rate=0.1,
                effective_tax_rate=0.1,
                tax_match_source="product",
                plausibility_range=PlausibilityRange(
                    kg_per_ctn_min=5,
                    kg_per_ctn_max=20,
                    kg_per_pc_min=0.5,
                    kg_per_pc_max=2,
                    unit_price_min=50,
                    unit_price_max=100,
                    qty_per_ctn_min=5,
                    qty_per_ctn_max=20,
                    source="test range",
                ),
            )
        ]
        rows = [
            {
                "中文品名": "品名",
                "英文品名": "Item",
                "商品编码": "3926909985",
                "材质": "Plastic",
                "用途": "Home use",
                "箱数": 10,
                "数量": 100,
                "单位": "PCS",
                "币制": "USD",
                "单价": 85.133,
                "总价": 8513.3,
                "净重": 92,
                "毛重": 100,
                "原产国": "CN",
                "综合税率": 0.1,
                "candidate_index": 0,
            }
        ]
        validate_llm_output_rows(
            rows,
            candidates,
            ManifestSummary("input.xlsx", 1, 10, 100, 1000, []),
            BillInfo("bill.pdf", "", []),
            ProcessingOptions(target_tax_amount=900, target_item_count=1),
        )

    def test_llm_output_closes_gross_weight_by_scaling_draft(self) -> None:
        candidates = [
            ProductCandidate(
                source="replacement",
                source_label="test",
                zh="品名",
                en="Item",
                hs="3926909985",
                material="Plastic",
                usage="HOME",
                ctns=10,
                qty=100,
                unit_price=85.133,
                gross_weight=100,
                base_tax_rate=0.1,
                effective_tax_rate=0.1,
                tax_match_source="product",
                plausibility_range=PlausibilityRange(
                    kg_per_ctn_min=5,
                    kg_per_ctn_max=20,
                    kg_per_pc_min=0.5,
                    kg_per_pc_max=2,
                    unit_price_min=50,
                    unit_price_max=100,
                    qty_per_ctn_min=5,
                    qty_per_ctn_max=20,
                    source="test range",
                ),
            )
        ]
        rows = [
            {
                "中文品名": "品名",
                "英文品名": "Item",
                "商品编码": "3926909985",
                "材质": "Plastic",
                "用途": "Home use",
                "箱数": 10,
                "数量": 100,
                "单位": "PCS",
                "币制": "USD",
                "单价": 85.133,
                "总价": 8513.3,
                "净重": 91.99,
                "毛重": 100,
                "原产国": "CN",
                "综合税率": 0.1,
                "candidate_index": 0,
            }
        ]

        validate_llm_output_rows(
            rows,
            candidates,
            ManifestSummary("input.xlsx", 1, 10, 120, 1000, []),
            BillInfo("bill.pdf", "", []),
            ProcessingOptions(target_tax_amount=900, target_item_count=1),
        )

        self.assertEqual(rows[0]["毛重"], 120)
        self.assertEqual(rows[0]["LLM草案毛重"], 100)
        self.assertEqual(rows[0]["毛重闭合调整"], 20)
        self.assertIn("毛重按Excel总重量倒推调整", rows[0]["约束提示"])

    def test_llm_output_rejects_tax_above_target(self) -> None:
        candidates = [
            ProductCandidate(
                source="replacement",
                source_label="test",
                zh="品名",
                en="Item",
                hs="3926909985",
                material="Plastic",
                usage="HOME",
                ctns=10,
                qty=100,
                unit_price=90.1,
                gross_weight=100,
                base_tax_rate=0.1,
                effective_tax_rate=0.1,
                tax_match_source="product",
                plausibility_range=PlausibilityRange(
                    kg_per_ctn_min=5,
                    kg_per_ctn_max=20,
                    kg_per_pc_min=0.5,
                    kg_per_pc_max=2,
                    unit_price_min=50,
                    unit_price_max=100,
                    qty_per_ctn_min=5,
                    qty_per_ctn_max=20,
                    source="test range",
                ),
            )
        ]
        rows = [
            {
                "中文品名": "品名",
                "英文品名": "Item",
                "商品编码": "3926909985",
                "材质": "Plastic",
                "用途": "Home use",
                "箱数": 10,
                "数量": 100,
                "单位": "PCS",
                "币制": "USD",
                "单价": 90.1,
                "总价": 9010,
                "净重": 92,
                "毛重": 100,
                "原产国": "CN",
                "综合税率": 0.1,
                "candidate_index": 0,
            }
        ]
        with self.assertRaisesRegex(RuntimeError, "允许区间 800.0-900"):
            validate_llm_output_rows(
                rows,
                candidates,
                ManifestSummary("input.xlsx", 1, 10, 100, 1000, []),
                BillInfo("bill.pdf", "", []),
                ProcessingOptions(target_tax_amount=900, target_item_count=1),
            )


if __name__ == "__main__":
    unittest.main()
