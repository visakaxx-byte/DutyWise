from __future__ import annotations

import asyncio
import unittest
import tempfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import load_workbook

from engine import (
    BillInfo,
    BillLLMFields,
    BillProduct,
    ManifestWeightInfo,
    ManifestSummary,
    ProductCandidate,
    ProcessingOptions,
    SelectionRules,
    apply_bill_product_names,
    build_manifest_weight_context,
    build_output_rows,
    write_workbook,
    build_plausible_row_plans,
    certification_filter_reason,
    effective_tax_rate,
    ensure_candidate_plausibility_ranges,
    ensure_bill_products_present,
    generate_valid_output_rows_with_llm,
    load_plausibility_ranges,
    load_replacement_candidates,
    normalize_manifest_weight_payload,
    normalize_bill_llm_products,
    normalize_bill_llm_fields,
    normalize_bill_llm_product_entries,
    parse_bill,
    parse_bill_product_entries_from_images,
    parse_bill_product_entries_from_text,
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
    translate_output_chinese_names_with_llm,
    normalize_generated_candidate,
    validate_llm_output_rows,
    validate_qty_ctn_relationship,
    infer_bill_material_from_entry,
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
        self.messages = []
        self.temperatures = []

    async def chat_json(self, messages: list[dict[str, str]], *, temperature: float = 0.1) -> dict:
        self.calls += 1
        self.messages.append(messages)
        self.temperatures.append(temperature)
        return self.payload


class FakeVisionBillParser(FakeBillParser):
    def __init__(self, payload: dict):
        super().__init__(payload)
        self.image_calls = 0
        self.image_prompts: list[str] = []
        self.image_data_urls: list[list[str]] = []

    async def chat_json_with_images(
        self,
        text: str,
        image_data_urls: list[str],
        *,
        temperature: float = 0.0,
        model: str | None = None,
        max_tokens: int = 4096,
    ) -> dict:
        self.image_calls += 1
        self.image_prompts.append(text)
        self.image_data_urls.append(image_data_urls)
        self.temperatures.append(temperature)
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


class DictFakeLLM:
    def __init__(self, payload: dict):
        self.payload = payload
        self.calls = 0

    async def chat_json(self, messages: list[dict[str, str]], *, temperature: float = 0.1) -> dict:
        self.calls += 1
        return self.payload


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
        self.assertAlmostEqual(
            effective_tax_rate({"tax_rate": "5%", "additional_tax_rate": "7.5%+10%", "additional_tax_details": []}),
            0.225,
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
        high_effective = {
            "hs_code_us": "7020006000",
            "tax_rate": "5%",
            "additional_tax_rate": "25%+10%",
            "certification_texts": [],
        }
        self.assertIn("综合税率", tax_filter_reason(high_effective, self.rules))
        self.assertEqual(tax_filter_reason(high_effective, self.rules, enforce_tax_limit=False), "")
        unknown = {"hs_code_us": "3924104000", "tax_rate": "3.4%", "certification_texts": ["FDA"]}
        self.assertIn("未允许", tax_filter_reason(unknown, self.rules))
        self.assertIn("未允许", tax_filter_reason(unknown, self.rules, enforce_tax_limit=False))


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
        self.assertEqual(parser.temperatures, [0.0])
        prompt = "\n".join(message.get("content", "") for message in parser.messages[0])
        self.assertIn("ignored_phrases", prompt)

    async def test_text_bill_cache_key_is_namespaced_and_preserves_product_entry_details(self) -> None:
        text = "STORAGE BAG HS CODE:420222"
        parser = FakeBillParser(
            {
                "shipper": "ACME EXPORT LTD\n1 SHIPPER ROAD",
                "consignee": "BUYER INC\n9 CONSIGNEE AVE",
                "products": [
                    {"name": "STORAGE BAG", "hs": "HS CODE: 4202.22", "evidence": text, "confidence": 0.98},
                ]
            }
        )
        cache = {"bill": {}}

        first = await parse_bill_product_entries_from_text(text, parser, cache)
        second = await parse_bill_product_entries_from_text(text, parser, cache)

        self.assertEqual(parser.calls, 1)
        self.assertTrue(list(cache["bill"].keys())[0].startswith("text:v3:"))
        self.assertEqual(second[0].name, first[0].name)
        self.assertEqual(second[0].hs_code_hint, "420222")
        self.assertEqual(second[0].evidence, text)

    async def test_parse_text_bill_uses_llm_shipper_and_consignee(self) -> None:
        parser = FakeBillParser(
            {
                "shipper": "ACME EXPORT LTD\n1 SHIPPER ROAD",
                "consignee": "BUYER INC\n9 CONSIGNEE AVE",
                "carton_count": 930,
                "carton_evidence": "SAY NINE HUNDRED AND THIRTY CARTONS ONLY",
                "products": [
                    {"name": "STORAGE BAG", "hs_code_hint": "420222", "evidence": "text evidence", "confidence": 0.96},
                ],
            }
        )
        text = "STORAGE BAG HS CODE:420222 " * 5
        with tempfile.TemporaryDirectory() as tmp:
            pdf = Path(tmp) / "text.pdf"
            pdf.write_bytes(b"fake pdf bytes")
            with patch("engine.extract_bill_text", return_value=text):
                bill = await parse_bill(pdf, parser, {"bill": {}})

        self.assertEqual(bill.shipper, "ACME EXPORT LTD\n1 SHIPPER ROAD")
        self.assertEqual(bill.consignee, "BUYER INC\n9 CONSIGNEE AVE")
        self.assertEqual(bill.cartons, 930)

    async def test_image_bill_parser_uses_vision_cache_key_and_image_llm(self) -> None:
        parser = FakeVisionBillParser(
            {
                "shipper": "VISION SHIPPER LTD\n2 EXPORT ST",
                "consignee": "VISION BUYER INC\n3 IMPORT AVE",
                "products": [
                    {"name": "STORAGE BAG", "hs_code_hint": "420222", "evidence": "image evidence", "confidence": 0.96},
                ]
            }
        )
        cache = {"bill": {}}
        with tempfile.TemporaryDirectory() as tmp:
            pdf = Path(tmp) / "scan.pdf"
            pdf.write_bytes(b"fake pdf bytes")
            first = await parse_bill_product_entries_from_images(pdf, ["data:image/jpeg;base64,abc"], parser, cache)
            second = await parse_bill_product_entries_from_images(pdf, ["data:image/jpeg;base64,abc"], parser, cache)

        self.assertEqual(parser.image_calls, 1)
        self.assertEqual(first[0].name, "STORAGE BAG")
        self.assertEqual(second[0].hs_code_hint, "420222")
        self.assertTrue(next(iter(cache["bill"])).startswith("vision:v3:"))
        self.assertIn("扫描图", parser.image_prompts[0])

    async def test_parse_bill_falls_back_to_vision_when_pdf_text_is_empty(self) -> None:
        parser = FakeVisionBillParser(
            {
                "shipper": "VISION SHIPPER LTD\n2 EXPORT ST",
                "consignee": "VISION BUYER INC\n3 IMPORT AVE",
                "carton_count": 661,
                "carton_evidence": "40'HC/661CARTONS/16536.71KGS/68.00CBM",
                "products": [
                    {"name": "PLASTIC SHELL", "hs_code_hint": "392690", "evidence": "PLASTIC SHELL HS CODE:392690", "confidence": 0.97},
                ]
            }
        )
        with tempfile.TemporaryDirectory() as tmp:
            pdf = Path(tmp) / "scan.pdf"
            pdf.write_bytes(b"fake pdf bytes")
            with (
                patch("engine.extract_bill_text", return_value=""),
                patch("engine.render_bill_pdf_pages", return_value=["data:image/jpeg;base64,abc"]),
            ):
                bill = await parse_bill(pdf, parser, {"bill": {}})

        self.assertEqual(bill.products, ["PLASTIC SHELL"])
        self.assertEqual(bill.product_entries[0].hs_code_hint, "392690")
        self.assertEqual(bill.parse_source, "vision")
        self.assertEqual(bill.text_chars, 0)
        self.assertEqual(bill.vision_pages, 1)
        self.assertEqual(bill.shipper, "VISION SHIPPER LTD\n2 EXPORT ST")
        self.assertEqual(bill.consignee, "VISION BUYER INC\n3 IMPORT AVE")
        self.assertEqual(bill.cartons, 661)

    async def test_parse_bill_keeps_text_path_when_pdf_text_is_long_enough(self) -> None:
        parser = FakeVisionBillParser(
            {
                "carton_count": 748,
                "products": [
                    {"name": "STORAGE BAG", "hs_code_hint": "420222", "evidence": "text evidence", "confidence": 0.96},
                ]
            }
        )
        text = "STORAGE BAG HS CODE:420222 " * 5
        with tempfile.TemporaryDirectory() as tmp:
            pdf = Path(tmp) / "text.pdf"
            pdf.write_bytes(b"fake pdf bytes")
            with (
                patch("engine.extract_bill_text", return_value=text),
                patch("engine.render_bill_pdf_pages", side_effect=AssertionError("vision path should not render")),
            ):
                bill = await parse_bill(pdf, parser, {"bill": {}})

        self.assertEqual(bill.products, ["STORAGE BAG"])
        self.assertEqual(bill.parse_source, "text")
        self.assertEqual(parser.image_calls, 0)
        self.assertEqual(bill.cartons, 748)

    async def test_parse_bill_rejects_missing_carton_count_without_manifest_fallback(self) -> None:
        parser = FakeVisionBillParser(
            {
                "products": [
                    {"name": "PLASTIC SHELL", "hs_code_hint": "392690", "evidence": "PLASTIC SHELL HS CODE:392690", "confidence": 0.97},
                ]
            }
        )
        with tempfile.TemporaryDirectory() as tmp:
            pdf = Path(tmp) / "scan.pdf"
            pdf.write_bytes(b"fake pdf bytes")
            with (
                patch("engine.extract_bill_text", return_value=""),
                patch("engine.render_bill_pdf_pages", return_value=["data:image/jpeg;base64,abc"]),
                self.assertRaisesRegex(RuntimeError, "提单未识别到有效总箱数"),
            ):
                await parse_bill(pdf, parser, {"bill": {}})

    async def test_parse_bill_scanned_empty_vision_result_has_clear_error(self) -> None:
        parser = FakeVisionBillParser({"products": []})
        with tempfile.TemporaryDirectory() as tmp:
            pdf = Path(tmp) / "scan.pdf"
            pdf.write_bytes(b"fake pdf bytes")
            with (
                patch("engine.extract_bill_text", return_value=""),
                patch("engine.render_bill_pdf_pages", return_value=["data:image/jpeg;base64,abc"]),
                self.assertRaisesRegex(RuntimeError, "提单为扫描件，视觉模型未识别到可靠货物品类"),
            ):
                await parse_bill(pdf, parser, {"bill": {}})

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

    def test_llm_product_entries_deduplicate_and_support_hs_aliases(self) -> None:
        payload = {
            "products": [
                {"name": "STORAGE BAG", "hs_code": "HS CODE: 4202.22", "evidence": "first", "confidence": 0.99},
                {"name": "storage   bag", "hs_code_hint": "999999", "evidence": "duplicate", "confidence": 0.99},
            ]
        }
        entries = normalize_bill_llm_product_entries(payload)
        self.assertEqual(len(entries), 1)
        self.assertEqual(entries[0].hs_code_hint, "420222")
        self.assertEqual(entries[0].evidence, "first")

    def test_llm_product_entries_split_comma_joined_products_and_hs_hints(self) -> None:
        payload = {
            "products": [
                {
                    "name": "PLASTIC ORNAMENTS,NECKLACE",
                    "hs_code_hint": "392640711790",
                    "evidence": "PLASTIC ORNAMENTS,NECKLACE HS:392640,711790",
                    "confidence": 1.0,
                }
            ]
        }
        entries = normalize_bill_llm_product_entries(payload)
        self.assertEqual([entry.name for entry in entries], ["PLASTIC ORNAMENTS", "NECKLACE"])
        self.assertEqual([entry.hs_code_hint for entry in entries], ["392640", "711790"])

    def test_llm_product_entries_reject_bad_structure_and_empty_products(self) -> None:
        with self.assertRaisesRegex(RuntimeError, "LLM 提单解析未返回 products 数组"):
            normalize_bill_llm_product_entries({"items": []})
        with self.assertRaisesRegex(RuntimeError, "LLM 未从提单中识别到可靠货物品类"):
            normalize_bill_llm_product_entries({"products": [{"name": "LOW CONFIDENCE ITEM", "confidence": 0.1}]})

    def test_bill_llm_fields_normalize_party_blocks(self) -> None:
        fields = normalize_bill_llm_fields(
            {
                "shipper": [" ACME EXPORT LTD  ", " 1 SHIPPER ROAD "],
                "consignee": " BUYER INC \n 9 CONSIGNEE AVE ",
                "carton_count": "1,234 CARTONS",
                "carton_evidence": "TOTAL 1,234 CARTONS",
                "products": [{"name": "STORAGE BAG", "confidence": 0.99}],
            }
        )

        self.assertEqual(fields.shipper, "ACME EXPORT LTD\n1 SHIPPER ROAD")
        self.assertEqual(fields.consignee, "BUYER INC\n9 CONSIGNEE AVE")
        self.assertEqual(fields.product_entries[0].name, "STORAGE BAG")
        self.assertEqual(fields.carton_count, 1234)
        self.assertEqual(fields.carton_evidence, "TOTAL 1,234 CARTONS")

    def test_bill_material_inference_prefers_modifier_and_does_not_overmatch_pe_pp(self) -> None:
        self.assertEqual(
            "Plastic",
            infer_bill_material_from_entry(
                BillProduct(name="PLASTIC ORNAMENTS", hs_code_hint="392640", evidence="PLASTIC ORNAMENTS HS:392640")
            ),
        )
        self.assertEqual(
            "",
            infer_bill_material_from_entry(BillProduct(name="PIECE", hs_code_hint="392640", evidence="PIECE HS:392640")),
        )


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
        self.assertLessEqual(abs(total_tax - 120), 20)
        self.assertEqual(rows[0]["税率"], "23.4%")
        self.assertAlmostEqual(rows[0]["税金"], round(rows[0]["总价"] * rows[0]["综合税率"], 2))

    def test_workbook_writes_tax_rate_and_tax_amount_as_last_columns(self) -> None:
        row = {
            "中文品名": "塑料杯",
            "英文品名": "Plastic Cup",
            "商品编码": "3924104000",
            "材质": "Plastic",
            "用途": "HOME",
            "箱数": 10,
            "数量": 100,
            "单位": "PCS",
            "币制": "USD",
            "单价": 2.5,
            "总价": 250,
            "净重": 90,
            "毛重": 100,
            "原产国": "CN",
            "综合税率": 0.109,
        }
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "clearance.xlsx"
            write_workbook(BillInfo("bill.pdf", "", []), [row], output, metadata={})
            workbook = load_workbook(output, data_only=True)
            sheet = workbook[workbook.sheetnames[0]]

        self.assertEqual(sheet.cell(5, 15).value, "税率")
        self.assertEqual(sheet.cell(5, 16).value, "税金")
        self.assertEqual(sheet.cell(6, 15).value, "10.9%")
        self.assertEqual(sheet.cell(6, 16).value, 27.25)
        self.assertEqual(row["税率"], "10.9%")
        self.assertEqual(row["税金"], 27.25)

    def test_workbook_writes_bill_shipper_and_consignee(self) -> None:
        row = {
            "中文品名": "塑料杯",
            "英文品名": "Plastic Cup",
            "商品编码": "3924104000",
            "材质": "Plastic",
            "用途": "HOME",
            "箱数": 10,
            "数量": 100,
            "单位": "PCS",
            "币制": "USD",
            "单价": 2.5,
            "总价": 250,
            "净重": 90,
            "毛重": 100,
            "原产国": "CN",
            "综合税率": 0.109,
        }
        bill = BillInfo(
            "bill.pdf",
            "",
            [],
            shipper="ACME EXPORT LTD\n1 SHIPPER ROAD",
            consignee="BUYER INC\n9 CONSIGNEE AVE",
        )
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "clearance.xlsx"
            write_workbook(bill, [dict(row)], output, metadata={})
            workbook = load_workbook(output, data_only=True)
            sheet = workbook[workbook.sheetnames[0]]

        self.assertEqual(sheet["C2"].value, "ACME EXPORT LTD\n1 SHIPPER ROAD")
        self.assertEqual(sheet["C3"].value, "BUYER INC\n9 CONSIGNEE AVE")

    def test_workbook_metadata_overrides_bill_parties(self) -> None:
        row = {
            "中文品名": "塑料杯",
            "英文品名": "Plastic Cup",
            "商品编码": "3924104000",
            "材质": "Plastic",
            "用途": "HOME",
            "箱数": 10,
            "数量": 100,
            "单位": "PCS",
            "币制": "USD",
            "单价": 2.5,
            "总价": 250,
            "净重": 90,
            "毛重": 100,
            "原产国": "CN",
            "综合税率": 0.109,
        }
        bill = BillInfo("bill.pdf", "", [], shipper="BILL SHIPPER", consignee="BILL CONSIGNEE")
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "clearance.xlsx"
            write_workbook(
                bill,
                [dict(row)],
                output,
                metadata={"shipper": "META SHIPPER", "consignee": "META CONSIGNEE"},
            )
            workbook = load_workbook(output, data_only=True)
            sheet = workbook[workbook.sheetnames[0]]

        self.assertEqual(sheet["C2"].value, "META SHIPPER")
        self.assertEqual(sheet["C3"].value, "META CONSIGNEE")

    def test_chinese_name_column_is_translated_to_simplified_chinese(self) -> None:
        rows = [
            {
                "中文品名": "PLASTIC ORNAMENTS",
                "英文品名": "Plastic Ornaments",
                "商品编码": "3926400090",
                "材质": "Plastic",
            }
        ]
        llm = DictFakeLLM({"translations": [{"index": 0, "translation": "塑料装饰品"}]})

        asyncio.run(translate_output_chinese_names_with_llm(llm, rows))

        self.assertEqual(rows[0]["中文品名"], "塑料装饰品")
        self.assertEqual(llm.calls, 1)

    def test_chinese_name_column_rejects_latin_translation(self) -> None:
        rows = [
            {
                "中文品名": "STORAGE BAG",
                "英文品名": "Storage Bag",
                "商品编码": "4202220000",
                "材质": "Textile",
            }
        ]
        llm = DictFakeLLM({"translations": [{"index": 0, "translation": "Storage Bag"}]})

        with self.assertRaisesRegex(RuntimeError, "有效简体中文品名"):
            asyncio.run(translate_output_chinese_names_with_llm(llm, rows))

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
                    base_tax_rate=0.1,
                    effective_tax_rate=0.1,
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
                qty=900,
                unit_price=0.8,
                real_weight=192,
                gross_weight=192,
                base_tax_rate=0.1,
                effective_tax_rate=0.1,
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
                qty=500,
                unit_price=1.2,
                real_weight=300,
                gross_weight=300,
                base_tax_rate=0.1,
                effective_tax_rate=0.1,
            ),
        ]
        manifest = ManifestSummary("input.xlsx", 2, 25, 492, 912, [])
        bill = BillInfo("bill.pdf", "", [], gross_weight=1200, cartons=25)
        options = ProcessingOptions(target_tax_amount=80, target_item_count=2)
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
        self.assertLessEqual(abs(total_tax - 80), 20)
        for plan in plans:
            self.assertGreaterEqual(plan.total_value * (plan.candidate.effective_tax_rate or plan.candidate.base_tax_rate), 30)
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
        with self.assertRaisesRegex(RuntimeError, "税金无法达到最低 30.0 USD|无法贴近期望税金"):
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
        options = ProcessingOptions(target_tax_amount=120, target_item_count=3)
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
                base_tax_rate=0.1,
                effective_tax_rate=0.1,
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
                qty=100,
                unit_price=2,
                real_weight=10,
                gross_weight=50,
            base_tax_rate=0.1,
            effective_tax_rate=0.1,
            tax_match_source="product",
            plausibility_range=PlausibilityRange(
                kg_per_ctn_min=2,
                kg_per_ctn_max=20,
                kg_per_pc_min=0.2,
                kg_per_pc_max=2,
                unit_price_min=0.5,
                unit_price_max=10,
                qty_per_ctn_min=5,
                qty_per_ctn_max=100,
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
        self.assertEqual(bill_required[0].tax_data["hs_code_us"], "4202221000")

    async def test_bill_product_with_hs_hint_accepts_actual_tax_over_twenty_percent(self) -> None:
        bill = BillInfo(
            filename="bill.pdf",
            raw_text="",
            products=["POLYESTER PILLOWCASE"],
            gross_weight=150,
            cartons=15,
            product_entries=[
                BillProduct(
                    name="POLYESTER PILLOWCASE",
                    hs_code_hint="630232",
                    evidence="POLYESTER PILLOWCASE HS CODE:630232",
                    confidence=0.99,
                )
            ],
        )
        options = ProcessingOptions(target_tax_amount=90, target_item_count=3)
        rules = SelectionRules(allowed_certifications=["Lacey Act", "TSCA"])
        crawler = FakeCrawler(
            {
                "6302322010": {
                    "hs_code_us": "6302322010",
                    "tax_rate": "11.4%",
                    "additional_tax_rate": "10%",
                    "certification_texts": [],
                    "description_cn": "化纤制床上用品",
                }
            }
        )

        bill_required, bill_filtered = await qualify_bill_product_candidates(
            crawler,
            bill,
            [],
            [],
            rules,
            options,
            query_cache={},
        )

        self.assertEqual(bill_filtered, [])
        self.assertEqual(len(bill_required), 1)
        self.assertEqual(bill_required[0].hs, "6302322010")
        self.assertAlmostEqual(bill_required[0].effective_tax_rate, 0.214)
        self.assertEqual(bill_required[0].tax_match_source, "bill_hs")

    async def test_bill_product_query_accepts_actual_tax_over_twenty_percent(self) -> None:
        bill = BillInfo(
            filename="bill.pdf",
            raw_text="",
            products=["POLYESTER PILLOWCASE"],
            gross_weight=150,
            cartons=15,
        )
        options = ProcessingOptions(target_tax_amount=90, target_item_count=3)
        rules = SelectionRules(allowed_certifications=["Lacey Act", "TSCA"])
        crawler = RoutedFakeCrawler(
            product_results={
                "polyester pillowcase": {
                    "6302322010": {
                        "hs_code_us": "6302322010",
                        "tax_rate": "11.4%",
                        "additional_tax_rate": "10%",
                        "certification_texts": [],
                        "description_cn": "化纤制床上用品",
                    }
                }
            }
        )

        bill_required, bill_filtered = await qualify_bill_product_candidates(
            crawler,
            bill,
            [],
            [],
            rules,
            options,
            query_cache={},
        )

        self.assertEqual(bill_filtered, [])
        self.assertEqual(len(bill_required), 1)
        self.assertEqual(bill_required[0].hs, "6302322010")
        self.assertAlmostEqual(bill_required[0].effective_tax_rate, 0.214)
        self.assertEqual(bill_required[0].tax_match_source, "bill_product")

    async def test_bill_product_with_six_digit_hs_hint_outputs_qualified_ten_digit_hts(self) -> None:
        bill = BillInfo(
            filename="bill.pdf",
            raw_text="",
            products=["PLASTIC ORNAMENTS"],
            gross_weight=220,
            cartons=22,
            product_entries=[
                BillProduct(
                    name="PLASTIC ORNAMENTS",
                    hs_code_hint="392640",
                    evidence="PLASTIC ORNAMENTS HS CODE:392640",
                    confidence=0.99,
                )
            ],
        )
        options = ProcessingOptions(target_tax_amount=58.3, target_item_count=1)
        rules = SelectionRules(allowed_certifications=["Lacey Act", "TSCA"])

        bill_required, bill_filtered = await qualify_bill_product_candidates(
            FakeCrawler(tax_result("5.3%", hs="3926400090")),
            bill,
            [],
            [],
            rules,
            options,
            query_cache={},
        )

        self.assertEqual(bill_filtered, [])
        self.assertEqual(len(bill_required), 1)
        self.assertEqual(bill_required[0].hs, "3926400090")
        self.assertEqual(bill_required[0].tax_data["hs_code_us"], "3926400090")
        self.assertEqual(bill_required[0].tax_match_source, "bill_hs")

    async def test_bill_material_prefers_plastic_modifier(self) -> None:
        bill = BillInfo(
            filename="bill.pdf",
            raw_text="",
            products=["PLASTIC ORNAMENTS"],
            gross_weight=220,
            cartons=22,
            product_entries=[
                BillProduct(
                    name="PLASTIC ORNAMENTS",
                    hs_code_hint="392640",
                    evidence="PLASTIC ORNAMENTS HS CODE:392640",
                    confidence=0.99,
                )
            ],
        )
        options = ProcessingOptions(target_tax_amount=58.3, target_item_count=1)
        rules = SelectionRules(allowed_certifications=["Lacey Act", "TSCA"])

        bill_required, _ = await qualify_bill_product_candidates(
            FakeCrawler(tax_result("5.3%", hs="3926400090")),
            bill,
            [],
            [],
            rules,
            options,
            query_cache={},
        )

        self.assertEqual(bill_required[0].material, "Plastic")
        self.assertEqual(bill_required[0].tax_match_source, "bill_hs")

    async def test_bill_material_can_be_inferred_from_hs_description(self) -> None:
        bill = BillInfo(
            filename="bill.pdf",
            raw_text="",
            products=["ORNAMENT"],
            gross_weight=100,
            cartons=10,
            product_entries=[
                BillProduct(
                    name="ORNAMENT",
                    hs_code_hint="392640",
                    evidence="ORNAMENT HS CODE:392640",
                    confidence=0.99,
                )
            ],
        )
        options = ProcessingOptions(target_tax_amount=90, target_item_count=2)
        rules = SelectionRules(allowed_certifications=["Lacey Act", "TSCA"])

        crawler = RoutedFakeCrawler(
            product_results={},
            hs_results={
                "392640": {
                    "3926400090": {
                        "hs_code_us": "3926400090",
                        "tax_rate": "5.3%",
                        "certification_texts": [],
                        "description_cn": "Other articles of plastics and articles of other materials of headings 3901 to 3914: Statuettes and other ornamental articles",
                    }
                }
            },
        )
        llm = DictFakeLLM({"material": "Plastic"})

        bill_required, _ = await qualify_bill_product_candidates(
            crawler,
            bill,
            [],
            [],
            rules,
            options,
            query_cache={},
            llm=llm,
        )

        self.assertEqual(bill_required[0].material, "Plastic")
        self.assertEqual(llm.calls, 1)

    async def test_bill_product_fallback_uses_inferred_material_without_material_field(self) -> None:
        bill = BillInfo(
            filename="bill.pdf",
            raw_text="",
            products=["PLASTIC ORNAMENTS"],
            gross_weight=100,
            cartons=10,
            product_entries=[
                BillProduct(
                    name="PLASTIC ORNAMENTS",
                    hs_code_hint="",
                    evidence="PLASTIC ORNAMENTS",
                    confidence=0.99,
                )
            ],
        )
        options = ProcessingOptions(target_tax_amount=30, target_item_count=1)
        rules = SelectionRules(allowed_certifications=["Lacey Act", "TSCA"])
        crawler = RoutedFakeCrawler(
            product_results={
                "plastic ornaments": {
                    "3926400090": {
                        "hs_code_us": "3926400090",
                        "tax_rate": "5.3%",
                        "certification_texts": [],
                        "description_cn": "Other articles of plastics and articles of other materials of headings 3901 to 3914: Statuettes and other ornamental articles",
                    }
                }
            },
            hs_results={},
        )

        bill_required, bill_filtered = await qualify_bill_product_candidates(
            crawler,
            bill,
            [],
            [],
            rules,
            options,
            query_cache={},
        )

        self.assertEqual(bill_filtered, [])
        self.assertEqual(len(bill_required), 1)
        self.assertEqual(bill_required[0].material, "Plastic")
        self.assertEqual(bill_required[0].tax_match_source, "bill_product")
        self.assertEqual(bill_required[0].hs, "3926400090")

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
            BillInfo("bill.pdf", "", ["HANDHELD GARMENT STEAMER"], cartons=10),
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
            BillInfo("bill.pdf", "", [], cartons=50),
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
                        {"candidate_index": 0, "箱数": 8, "数量": 60, "单价": 5.0, "毛重": 20},
                        {"candidate_index": 1, "箱数": 9, "数量": 60, "单价": 5.0, "毛重": 26},
                        {"candidate_index": 2, "箱数": 10, "数量": 80, "单价": 4.0, "毛重": 32},
                        {"candidate_index": 3, "箱数": 11, "数量": 80, "单价": 4.0, "毛重": 39},
                        {"candidate_index": 4, "箱数": 12, "数量": 90, "单价": 4.0, "毛重": 48},
                        {"candidate_index": 5, "箱数": 13, "数量": 100, "单价": 3.5, "毛重": 58},
                        {"candidate_index": 6, "箱数": 14, "数量": 110, "单价": 3.5, "毛重": 70},
                        {"candidate_index": 7, "箱数": 15, "数量": 120, "单价": 3.0, "毛重": 83},
                    ]
                },
            ]
        )
        manifest = ManifestSummary("input.xlsx", 8, 92, 376, 1000, [])
        bill = BillInfo("bill.pdf", "", [], cartons=92)
        options = ProcessingOptions(target_tax_amount=350, target_item_count=8)

        rows, attempts, feedback = await generate_valid_output_rows_with_llm(llm, candidates, manifest, bill, options)

        self.assertEqual(attempts, 2)
        self.assertIn("数量分布过于机械", feedback[0])
        self.assertEqual(len(rows), 8)
        self.assertLessEqual(abs(sum(row["总价"] * row["综合税率"] for row in rows) - 350), 20)
        self.assertTrue(all(row["税金"] >= 30 for row in rows))

    def test_llm_output_closes_tax_gap_within_20_usd(self) -> None:
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
                    unit_price_min=90,
                    unit_price_max=91,
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
            BillInfo("bill.pdf", "", [], cartons=10),
            ProcessingOptions(target_tax_amount=900, target_item_count=1),
        )
        self.assertEqual(rows[0]["单价"], 90.0)
        self.assertEqual(rows[0]["税金"], 900.0)
        self.assertLessEqual(abs(sum(row["总价"] * row["综合税率"] for row in rows) - 900), 20)

    def test_llm_output_closes_bill_cartons(self) -> None:
        candidates = [
            ProductCandidate(
                source="replacement",
                source_label="test",
                zh="品名A",
                en="Item A",
                hs="3926909985",
                material="Plastic",
                usage="HOME",
                ctns=10,
                qty=4000,
                unit_price=0.2,
                gross_weight=400,
                base_tax_rate=0.1,
                effective_tax_rate=0.1,
                tax_match_source="product",
                plausibility_range=PlausibilityRange(
                    kg_per_ctn_min=5,
                    kg_per_ctn_max=20,
                    kg_per_pc_min=0.5,
                    kg_per_pc_max=2,
                    unit_price_min=0.1,
                    unit_price_max=1.0,
                    qty_per_ctn_min=5,
                    qty_per_ctn_max=20,
                    source="test range",
                ),
            ),
            ProductCandidate(
                source="replacement",
                source_label="test",
                zh="品名B",
                en="Item B",
                hs="3926909986",
                material="Plastic",
                usage="HOME",
                ctns=10,
                qty=4170,
                unit_price=0.2,
                gross_weight=417,
                base_tax_rate=0.1,
                effective_tax_rate=0.1,
                tax_match_source="product",
                plausibility_range=PlausibilityRange(
                    kg_per_ctn_min=5,
                    kg_per_ctn_max=20,
                    kg_per_pc_min=0.5,
                    kg_per_pc_max=2,
                    unit_price_min=0.1,
                    unit_price_max=1.0,
                    qty_per_ctn_min=5,
                    qty_per_ctn_max=20,
                    source="test range",
                ),
            ),
        ]
        rows = [
            {
                "中文品名": "品名A",
                "英文品名": "Item A",
                "商品编码": "3926909985",
                "材质": "Plastic",
                "用途": "Home use",
                "箱数": 400,
                "数量": 4000,
                "单位": "PCS",
                "币制": "USD",
                "单价": 0.2,
                "总价": 800.0,
                "净重": 368,
                "毛重": 400,
                "原产国": "CN",
                "综合税率": 0.1,
                "candidate_index": 0,
            },
            {
                "中文品名": "品名B",
                "英文品名": "Item B",
                "商品编码": "3926909986",
                "材质": "Plastic",
                "用途": "Home use",
                "箱数": 417,
                "数量": 4170,
                "单位": "PCS",
                "币制": "USD",
                "单价": 0.2,
                "总价": 834.0,
                "净重": 384.36,
                "毛重": 417,
                "原产国": "CN",
                "综合税率": 0.1,
                "candidate_index": 1,
            },
        ]

        validate_llm_output_rows(
            rows,
            candidates,
            ManifestSummary("input.xlsx", 2, 0, 0, 1000, []),
            BillInfo("bill.pdf", "", [], cartons=817),
            ProcessingOptions(target_tax_amount=160, target_item_count=2),
        )

        self.assertEqual(sum(row["箱数"] for row in rows), 817)

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
                    unit_price_min=90,
                    unit_price_max=91,
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
            BillInfo("bill.pdf", "", [], cartons=10),
            ProcessingOptions(target_tax_amount=900, target_item_count=1),
        )

        self.assertEqual(rows[0]["毛重"], 120)
        self.assertEqual(rows[0]["LLM草案毛重"], 100)
        self.assertEqual(rows[0]["毛重闭合调整"], 20)
        self.assertIn("毛重按Excel总重量倒推调整", rows[0]["约束提示"])

    def test_llm_output_requires_ten_digit_hs_after_bill_hint_lookup(self) -> None:
        candidates = [
            ProductCandidate(
                source="bill_product",
                source_label="bill.pdf",
                zh="塑料装饰品",
                en="Plastic Ornaments",
                hs="3926400090",
                material="Plastic",
                usage="HOME",
                ctns=10,
                qty=100,
                unit_price=95,
                gross_weight=100,
                base_tax_rate=0.1,
                effective_tax_rate=0.1,
                tax_match_source="bill_hs",
                tax_data={
                    "hs_code_us": "3926400090",
                    "tax_rate": "10%",
                    "description_cn": "Statuettes and other ornamental articles",
                },
                plausibility_range=PlausibilityRange(
                    kg_per_ctn_min=5,
                    kg_per_ctn_max=20,
                    kg_per_pc_min=0.5,
                    kg_per_pc_max=2,
                    unit_price_min=90,
                    unit_price_max=100,
                    qty_per_ctn_min=5,
                    qty_per_ctn_max=20,
                    source="test range",
                ),
            )
        ]
        rows = [
            {
                "中文品名": "塑料装饰品",
                "英文品名": "Plastic Ornaments",
                "商品编码": "3926400090",
                "材质": "Plastic",
                "用途": "Home use",
                "箱数": 10,
                "数量": 100,
                "单位": "PCS",
                "币制": "USD",
                "单价": 95,
                "总价": 9500,
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
            BillInfo("bill.pdf", "", ["PLASTIC ORNAMENTS"], cartons=10),
            ProcessingOptions(target_tax_amount=950, target_item_count=1),
        )

        self.assertEqual(str(rows[0]["商品编码"]), "3926400090")
        self.assertEqual(rows[0]["税金"], 950.0)

    def test_qty_must_not_be_less_than_ctns_and_must_be_multiple(self) -> None:
        with self.assertRaisesRegex(RuntimeError, "数量不能小于箱数"):
            validate_qty_ctn_relationship([{"箱数": 50, "数量": 25}])
        with self.assertRaisesRegex(RuntimeError, "整数倍"):
            validate_qty_ctn_relationship([{"箱数": 25, "数量": 51}])
        validate_qty_ctn_relationship([{"箱数": 25, "数量": 50}])

    def test_llm_output_enforces_minimum_tax_per_row(self) -> None:
        candidates = []
        rows = []
        for idx, (rate, unit_price) in enumerate(((0.1, 10.0), (0.1, 1.0)), start=1):
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
                    qty=10,
                    unit_price=unit_price,
                    gross_weight=100,
                    base_tax_rate=rate,
                    effective_tax_rate=rate,
                    tax_match_source="product",
                    plausibility_range=PlausibilityRange(
                        kg_per_ctn_min=5,
                        kg_per_ctn_max=20,
                        kg_per_pc_min=0.5,
                        kg_per_pc_max=20,
                        unit_price_min=1,
                        unit_price_max=200,
                        qty_per_ctn_min=1,
                        qty_per_ctn_max=1,
                        source="test range",
                    ),
                )
            )
            rows.append(
                {
                    "中文品名": f"品名{idx}",
                    "英文品名": f"Item {idx}",
                    "商品编码": f"39269099{idx:02d}",
                    "材质": "Plastic",
                    "用途": "Home use",
                    "箱数": 10,
                    "数量": 10,
                    "单位": "PCS",
                    "币制": "USD",
                    "单价": unit_price,
                    "总价": unit_price * 10,
                    "净重": 92,
                    "毛重": 100,
                    "原产国": "CN",
                    "综合税率": rate,
                    "candidate_index": idx - 1,
                }
            )

        validate_llm_output_rows(
            rows,
            candidates,
            ManifestSummary("input.xlsx", 2, 20, 200, 1000, []),
            BillInfo("bill.pdf", "", [], cartons=20),
            ProcessingOptions(target_tax_amount=120, target_item_count=2),
        )

        self.assertGreaterEqual(rows[0]["税金"], 30)
        self.assertGreaterEqual(rows[1]["税金"], 30)
        self.assertLessEqual(abs(sum(row["税金"] for row in rows) - 120), 20)

    def test_llm_output_allows_up_to_two_zero_tax_rows(self) -> None:
        candidates = []
        rows = []
        for idx, rate in enumerate((0.0, 0.0, 0.1), start=1):
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
                    qty=10,
                    unit_price=10,
                    gross_weight=100,
                    base_tax_rate=rate,
                    effective_tax_rate=rate,
                    tax_match_source="product",
                    plausibility_range=PlausibilityRange(
                        kg_per_ctn_min=5,
                        kg_per_ctn_max=20,
                        kg_per_pc_min=0.5,
                        kg_per_pc_max=20,
                        unit_price_min=1,
                        unit_price_max=200,
                        qty_per_ctn_min=1,
                        qty_per_ctn_max=1,
                        source="test range",
                    ),
                )
            )
            rows.append(
                {
                    "中文品名": f"品名{idx}",
                    "英文品名": f"Item {idx}",
                    "商品编码": f"39269099{idx:02d}",
                    "材质": "Plastic",
                    "用途": "Home use",
                    "箱数": 10,
                    "数量": 10,
                    "单位": "PCS",
                    "币制": "USD",
                    "单价": 10,
                    "总价": 100,
                    "净重": 92,
                    "毛重": 100,
                    "原产国": "CN",
                    "综合税率": rate,
                    "candidate_index": idx - 1,
                }
            )

        validate_llm_output_rows(
            rows,
            candidates,
            ManifestSummary("input.xlsx", 3, 30, 300, 1000, []),
            BillInfo("bill.pdf", "", [], cartons=30),
            ProcessingOptions(target_tax_amount=100, target_item_count=3),
        )

        self.assertEqual([row["税金"] for row in rows[:2]], [0.0, 0.0])
        self.assertGreaterEqual(rows[2]["税金"], 30)

    def test_llm_output_rejects_three_zero_tax_rows(self) -> None:
        candidates = []
        rows = []
        for idx in range(3):
            candidates.append(
                ProductCandidate(
                    source="replacement",
                    source_label="test",
                    zh=f"免税品名{idx}",
                    en=f"Free Item {idx}",
                    hs=f"39269099{idx:02d}",
                    material="Plastic",
                    usage="HOME",
                    ctns=10,
                    qty=10,
                    unit_price=10,
                    gross_weight=100,
                    base_tax_rate=0.0,
                    effective_tax_rate=0.0,
                    tax_match_source="product",
                    plausibility_range=PlausibilityRange(
                        kg_per_ctn_min=5,
                        kg_per_ctn_max=20,
                        kg_per_pc_min=0.5,
                        kg_per_pc_max=20,
                        unit_price_min=1,
                        unit_price_max=200,
                        qty_per_ctn_min=1,
                        qty_per_ctn_max=1,
                        source="test range",
                    ),
                )
            )
            rows.append(
                {
                    "中文品名": f"免税品名{idx}",
                    "英文品名": f"Free Item {idx}",
                    "商品编码": f"39269099{idx:02d}",
                    "材质": "Plastic",
                    "用途": "Home use",
                    "箱数": 10,
                    "数量": 10,
                    "单位": "PCS",
                    "币制": "USD",
                    "单价": 10,
                    "总价": 100,
                    "净重": 92,
                    "毛重": 100,
                    "原产国": "CN",
                    "综合税率": 0.0,
                    "candidate_index": idx,
                }
            )

        with self.assertRaisesRegex(RuntimeError, "税金为 0 的行数不能超过 2 行"):
            validate_llm_output_rows(
                rows,
                candidates,
                ManifestSummary("input.xlsx", 3, 30, 300, 1000, []),
                BillInfo("bill.pdf", "", [], cartons=30),
                ProcessingOptions(target_tax_amount=1, target_item_count=3),
            )

    def test_llm_output_rejects_when_row_cannot_reach_minimum_tax(self) -> None:
        candidates = [
            ProductCandidate(
                source="replacement",
                source_label="test",
                zh="低税金品名",
                en="Low Tax Item",
                hs="3926909985",
                material="Plastic",
                usage="HOME",
                ctns=10,
                qty=10,
                unit_price=1,
                gross_weight=100,
                base_tax_rate=0.1,
                effective_tax_rate=0.1,
                tax_match_source="product",
                plausibility_range=PlausibilityRange(
                    kg_per_ctn_min=5,
                    kg_per_ctn_max=20,
                    kg_per_pc_min=0.5,
                    kg_per_pc_max=20,
                    unit_price_min=1,
                    unit_price_max=2,
                    qty_per_ctn_min=1,
                    qty_per_ctn_max=1,
                    source="test range",
                ),
            )
        ]
        rows = [
            {
                "中文品名": "低税金品名",
                "英文品名": "Low Tax Item",
                "商品编码": "3926909985",
                "材质": "Plastic",
                "用途": "Home use",
                "箱数": 10,
                "数量": 10,
                "单位": "PCS",
                "币制": "USD",
                "单价": 1,
                "总价": 10,
                "净重": 92,
                "毛重": 100,
                "原产国": "CN",
                "综合税率": 0.1,
                "candidate_index": 0,
            }
        ]

        with self.assertRaisesRegex(RuntimeError, "税金无法达到最低 30.0 USD"):
            validate_llm_output_rows(
                rows,
                candidates,
                ManifestSummary("input.xlsx", 1, 10, 100, 1000, []),
                BillInfo("bill.pdf", "", [], cartons=10),
                ProcessingOptions(target_tax_amount=30, target_item_count=1),
            )

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
                    unit_price_min=90,
                    unit_price_max=91,
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
        with self.assertRaisesRegex(RuntimeError, "允许区间 780.0-820"):
            validate_llm_output_rows(
                rows,
                candidates,
                ManifestSummary("input.xlsx", 1, 10, 100, 1000, []),
                BillInfo("bill.pdf", "", [], cartons=10),
                ProcessingOptions(target_tax_amount=800, target_item_count=1),
            )


if __name__ == "__main__":
    unittest.main()
