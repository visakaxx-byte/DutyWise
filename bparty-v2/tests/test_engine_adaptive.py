from __future__ import annotations

import asyncio
import json
import tempfile
import threading
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from engine import (
    BillInfo,
    ManifestSummary,
    PlausibilityRange,
    ProductCandidate,
    SelectionRules,
    apply_llm_review_constraints,
    attach_price_evidence_to_candidates_concurrently,
    ensure_candidate_plausibility_ranges,
    parse_input_documents_concurrently,
    qualify_manual_invoice_replacement_batch,
)


def tax_result(rate: str, hs: str) -> dict[str, dict]:
    return {
        hs: {
            "hs_code_us": hs,
            "tax_rate": rate,
            "certification_texts": [],
            "description_cn": "测试品名",
        }
    }


class DictFakeLLM:
    def __init__(self, payload: dict):
        self.payload = payload

    async def chat_json(self, messages, **kwargs):
        return self.payload


class QueueFakeLLM:
    def __init__(self, payloads: list[dict]):
        self.payloads = list(payloads)
        self.calls = 0

    async def chat_json(self, messages, **kwargs):
        self.calls += 1
        return self.payloads.pop(0)


class QueueVisionPackageLLM:
    def __init__(self, payloads: list[dict | BaseException]):
        self.payloads = list(payloads)
        self.image_calls = 0
        self.prompts: list[str] = []
        self.temperatures: list[float] = []

    async def chat_json_with_images(
        self,
        text,
        image_data_urls,
        *,
        temperature=0.0,
        **kwargs,
    ):
        self.image_calls += 1
        self.prompts.append(text)
        self.temperatures.append(temperature)
        payload = self.payloads.pop(0)
        if isinstance(payload, BaseException):
            raise payload
        return payload


class RoutedFakeCrawler:
    def __init__(self, hs_results: dict[str, dict[str, dict]]):
        self.hs_results = hs_results
        self.settings = SimpleNamespace(delay=0.0)
        self.hs_calls: list[str] = []

    async def search(self, hs_code: str) -> dict[str, dict]:
        self.hs_calls.append(hs_code)
        return self.hs_results[hs_code]

    async def search_product(self, product_name: str, material: str = "") -> dict[str, dict]:
        raise AssertionError("HS replacement should not use product search")


class AdaptiveWorkflowTests(unittest.IsolatedAsyncioTestCase):
    async def test_manifest_and_bill_parsers_run_with_two_way_concurrency(self) -> None:
        active = 0
        max_active = 0

        async def fake_manifest(*args, **kwargs):
            nonlocal active, max_active
            active += 1
            max_active = max(max_active, active)
            await asyncio.sleep(0.03)
            active -= 1
            return ManifestSummary("input.xlsx", 1, 1, 1, 0, ["item"])

        async def fake_bill(*args, **kwargs):
            nonlocal active, max_active
            active += 1
            max_active = max(max_active, active)
            await asyncio.sleep(0.03)
            active -= 1
            return BillInfo("bill.pdf", "", [], cartons=1)

        with (
            patch("engine.parse_manifest", side_effect=fake_manifest),
            patch("engine.parse_bill", side_effect=fake_bill),
        ):
            manifest, bill = await parse_input_documents_concurrently(
                "input.xlsx", "bill.pdf", DictFakeLLM({}), DictFakeLLM({}), {}
            )

        self.assertEqual(manifest.filename, "input.xlsx")
        self.assertEqual(bill.filename, "bill.pdf")
        self.assertEqual(max_active, 2)

    async def test_parser_concurrency_retries_only_failed_side_serially(self) -> None:
        manifest_calls = 0
        bill_calls = 0

        async def fake_manifest(*args, **kwargs):
            nonlocal manifest_calls
            manifest_calls += 1
            if manifest_calls == 1:
                raise RuntimeError("temporary provider concurrency rejection")
            return ManifestSummary("input.xlsx", 1, 1, 1, 0, ["item"])

        async def fake_bill(*args, **kwargs):
            nonlocal bill_calls
            bill_calls += 1
            return BillInfo("bill.pdf", "", [], cartons=1)

        metrics: dict = {}
        with (
            patch("engine.parse_manifest", side_effect=fake_manifest),
            patch("engine.parse_bill", side_effect=fake_bill),
        ):
            await parse_input_documents_concurrently(
                "input.xlsx",
                "bill.pdf",
                DictFakeLLM({}),
                DictFakeLLM({}),
                {},
                metrics=metrics,
            )

        self.assertEqual(manifest_calls, 2)
        self.assertEqual(bill_calls, 1)
        self.assertEqual(metrics["sequential_fallbacks"], 1)

    async def test_missing_bill_package_count_is_reidentified_three_times_with_manifest_context(self) -> None:
        manifest = ManifestSummary(
            "input.xlsx",
            1,
            153,
            1898,
            0,
            ["Curtain"],
        )
        bill = BillInfo(
            "bill.pdf",
            "",
            ["Curtain"],
            cartons=None,
            parse_source="vision",
            carton_resolution_history=[
                {
                    "attempt": 1,
                    "status": "missing",
                    "package_count": None,
                }
            ],
        )
        parser = QueueVisionPackageLLM(
            [
                {
                    "package_count": None,
                    "package_unit": "PIECES",
                    "source_document": "air_waybill",
                    "evidence": "RCP maybe 153",
                    "reasoning_summary": "uncertain candidate 153",
                    "confidence": 0.55,
                },
                {
                    "package_count": 153,
                    "package_unit": "PIECES",
                    "source_document": "air_waybill",
                    "field_label": "No. of Pieces RCP",
                    "evidence": "No. of Pieces RCP: 153",
                    "manifest_comparison": "提单与清单均为 153",
                    "reasoning_summary": "航空运单整票交运件数",
                    "confidence": 0.98,
                    "inferred": True,
                },
            ]
        )
        cache: dict = {}
        metrics: dict = {}
        with tempfile.TemporaryDirectory() as tmp:
            bill_path = Path(tmp) / "bill.pdf"
            bill_path.write_bytes(b"fake")
            with (
                patch("engine.parse_manifest", return_value=manifest),
                patch("engine.parse_bill", return_value=bill),
                patch(
                    "engine.render_bill_pdf_pages",
                    return_value=["data:image/jpeg;base64,abc"],
                ) as render_pages,
            ):
                _, resolved = await parse_input_documents_concurrently(
                    "input.xlsx",
                    bill_path,
                    DictFakeLLM({}),
                    parser,
                    cache,
                    metrics=metrics,
                )

        self.assertEqual(resolved.cartons, 153)
        self.assertEqual(resolved.carton_unit, "PIECES")
        self.assertEqual(resolved.carton_source, "air_waybill")
        self.assertEqual(resolved.carton_recognition_attempts, 3)
        self.assertEqual(parser.image_calls, 2)
        self.assertEqual(parser.temperatures, [0.1, 0.2])
        self.assertIn("No. of Pieces RCP", parser.prompts[0])
        self.assertIn('"total_ctns":153', parser.prompts[0])
        self.assertIn("RCP maybe 153", parser.prompts[1])
        self.assertIn("uncertain candidate 153", parser.prompts[1])
        self.assertEqual(render_pages.call_args_list[0].kwargs["max_pages"], 2)
        self.assertEqual(render_pages.call_args_list[1].kwargs["max_pages"], 6)
        self.assertEqual(metrics["carton_recognition_attempts"], 3)
        audit = next(iter(cache["bill_package_resolution"].values()))
        self.assertEqual(audit["status"], "accepted")
        self.assertEqual(len(audit["attempts"]), 3)

    async def test_package_recognition_aborts_after_three_unsuccessful_attempts(self) -> None:
        manifest = ManifestSummary("input.xlsx", 1, 0, 1898, 0, ["Curtain"])
        bill = BillInfo(
            "bill.pdf",
            "",
            ["Curtain"],
            cartons=None,
            parse_source="vision",
        )
        parser = QueueVisionPackageLLM(
            [
                {"package_count": None},
                {"package_count": None},
            ]
        )
        cache: dict = {}
        with tempfile.TemporaryDirectory() as tmp:
            bill_path = Path(tmp) / "bill.pdf"
            bill_path.write_bytes(b"fake")
            with (
                patch("engine.parse_manifest", return_value=manifest),
                patch("engine.parse_bill", return_value=bill),
                patch(
                    "engine.render_bill_pdf_pages",
                    return_value=["data:image/jpeg;base64,abc"],
                ),
                self.assertRaisesRegex(RuntimeError, "模型已识别 3 次"),
            ):
                await parse_input_documents_concurrently(
                    "input.xlsx",
                    bill_path,
                    DictFakeLLM({}),
                    parser,
                    cache,
                )

        self.assertEqual(parser.image_calls, 2)
        audit = next(iter(cache["bill_package_resolution"].values()))
        self.assertEqual(audit["status"], "failed")
        self.assertEqual(len(audit["attempts"]), 3)

    async def test_manifest_mismatch_is_returned_to_model_for_final_selection(self) -> None:
        manifest = ManifestSummary("input.xlsx", 1, 160, 1898, 0, ["Curtain"])
        bill = BillInfo(
            "bill.pdf",
            "",
            ["Curtain"],
            cartons=153,
            carton_unit="PIECES",
            carton_evidence="No. of Pieces RCP: 153",
            parse_source="vision",
        )
        parser = QueueVisionPackageLLM(
            [
                {
                    "package_count": 153,
                    "package_unit": "PIECES",
                    "source_document": "air_waybill",
                    "evidence": "No. of Pieces RCP: 153",
                    "manifest_comparison": "清单 160 是商品行箱数汇总；采用运单整票交运件数 153",
                    "reasoning_summary": "以航空运单标准汇总栏为准",
                    "confidence": 0.93,
                }
            ]
        )
        with tempfile.TemporaryDirectory() as tmp:
            bill_path = Path(tmp) / "bill.pdf"
            bill_path.write_bytes(b"fake")
            with (
                patch("engine.parse_manifest", return_value=manifest),
                patch("engine.parse_bill", return_value=bill),
                patch(
                    "engine.render_bill_pdf_pages",
                    return_value=["data:image/jpeg;base64,abc"],
                ),
            ):
                _, resolved = await parse_input_documents_concurrently(
                    "input.xlsx",
                    bill_path,
                    DictFakeLLM({}),
                    parser,
                    {},
                )

        self.assertEqual(resolved.cartons, 153)
        self.assertEqual(resolved.carton_recognition_attempts, 2)
        self.assertIn("清单 160", resolved.carton_manifest_comparison)
        self.assertIn("不一致", parser.prompts[0])

    async def test_failed_initial_bill_parse_counts_toward_three_attempt_limit(self) -> None:
        manifest = ManifestSummary("input.xlsx", 1, 153, 1898, 0, ["Curtain"])
        bill_calls = 0

        async def fake_bill(*args, **kwargs):
            nonlocal bill_calls
            bill_calls += 1
            if bill_calls == 1:
                raise RuntimeError("sensitive provider body")
            return BillInfo(
                "bill.pdf",
                "",
                ["Curtain"],
                cartons=None,
                parse_source="vision",
            )

        parser = QueueVisionPackageLLM(
            [
                {
                    "package_count": 153,
                    "package_unit": "PIECES",
                    "source_document": "air_waybill",
                    "evidence": "No. of Pieces RCP: 153",
                    "reasoning_summary": "航空运单整票交运件数",
                }
            ]
        )
        with tempfile.TemporaryDirectory() as tmp:
            bill_path = Path(tmp) / "bill.pdf"
            bill_path.write_bytes(b"fake")
            with (
                patch("engine.parse_manifest", return_value=manifest),
                patch("engine.parse_bill", side_effect=fake_bill),
                patch(
                    "engine.render_bill_pdf_pages",
                    return_value=["data:image/jpeg;base64,abc"],
                ),
            ):
                _, resolved = await parse_input_documents_concurrently(
                    "input.xlsx",
                    bill_path,
                    DictFakeLLM({}),
                    parser,
                    {},
                )

        self.assertEqual(bill_calls, 2)
        self.assertEqual(parser.image_calls, 1)
        self.assertEqual(resolved.carton_recognition_attempts, 3)
        self.assertEqual(resolved.carton_resolution_history[0]["error"], "RuntimeError")
        self.assertNotIn(
            "sensitive provider body",
            json.dumps(resolved.carton_resolution_history),
        )

    async def test_three_initial_bill_parse_failures_are_audited_and_sanitized(self) -> None:
        manifest = ManifestSummary("input.xlsx", 1, 153, 1898, 0, ["Curtain"])
        bill_calls: list[dict] = []

        async def fake_bill(*args, **kwargs):
            bill_calls.append(dict(kwargs))
            raise RuntimeError("sensitive provider body api-key-like-value")

        cache: dict = {}
        with tempfile.TemporaryDirectory() as tmp:
            bill_path = Path(tmp) / "bill.pdf"
            bill_path.write_bytes(b"fake")
            with (
                patch("engine.parse_manifest", return_value=manifest),
                patch("engine.parse_bill", side_effect=fake_bill),
            ):
                with self.assertRaisesRegex(RuntimeError, "模型已识别 3 次") as caught:
                    await parse_input_documents_concurrently(
                        "input.xlsx",
                        bill_path,
                        DictFakeLLM({}),
                        DictFakeLLM({}),
                        cache,
                    )

        self.assertEqual(len(bill_calls), 3)
        self.assertEqual(bill_calls[1]["text_max_chars"], 12000)
        self.assertEqual(bill_calls[1]["vision_max_pages"], 2)
        self.assertEqual(bill_calls[2]["text_max_chars"], 24000)
        self.assertEqual(bill_calls[2]["vision_max_pages"], 6)
        self.assertNotIn("sensitive provider body", str(caught.exception))
        audit = next(iter(cache["bill_package_resolution"].values()))
        self.assertEqual(audit["status"], "failed")
        self.assertEqual(len(audit["attempts"]), 3)
        self.assertEqual(
            [entry["error"] for entry in audit["attempts"]],
            ["RuntimeError", "RuntimeError", "RuntimeError"],
        )
        self.assertNotIn("api-key-like-value", json.dumps(audit))

    async def test_plausibility_estimation_batches_candidates_with_two_way_concurrency(self) -> None:
        class ConcurrentBatchLLM:
            def __init__(self):
                self.calls = 0
                self.active = 0
                self.max_active = 0

            async def chat_json(self, messages, **kwargs):
                self.calls += 1
                self.active += 1
                self.max_active = max(self.max_active, self.active)
                await asyncio.sleep(0.03)
                context = json.loads(messages[-1]["content"].split("上下文：", 1)[1])
                self.active -= 1
                return {
                    "candidates": [
                        {
                            "candidate_id": item["candidate_id"],
                            **valid_plausibility_payload("batch test"),
                        }
                        for item in context["candidates"]
                    ]
                }

        candidates = build_bill_candidates(8)
        llm = ConcurrentBatchLLM()
        metrics: dict = {}

        resolved, used = await ensure_candidate_plausibility_ranges(
            candidates,
            ManifestSummary("input.xlsx", 8, 80, 400, 0, []),
            BillInfo("bill.pdf", "", [], cartons=80),
            {},
            llm,
            {},
            metrics=metrics,
        )

        self.assertTrue(used)
        self.assertEqual(len(resolved), 8)
        self.assertEqual(llm.calls, 2)
        self.assertEqual(llm.max_active, 2)
        self.assertEqual(metrics["batch_calls"], 2)
        self.assertEqual(metrics["individual_fallback_calls"], 0)

    async def test_price_evidence_network_work_runs_with_two_way_concurrency(self) -> None:
        candidates = build_bill_candidates(4)
        active = 0
        max_active = 0
        lock = threading.Lock()

        def fake_attach(items, **kwargs):
            nonlocal active, max_active
            with lock:
                active += 1
                max_active = max(max_active, active)
            time.sleep(0.03)
            with lock:
                active -= 1
            return items

        with patch("engine.attach_price_evidence_to_candidates", side_effect=fake_attach):
            resolved = await attach_price_evidence_to_candidates_concurrently(
                candidates, max_concurrency=2
            )

        self.assertEqual(resolved, candidates)
        self.assertEqual(max_active, 2)

    async def test_malformed_plausibility_batch_falls_back_per_candidate(self) -> None:
        candidates = build_bill_candidates(2)
        valid = valid_plausibility_payload("fallback")
        llm = QueueFakeLLM([{}, valid, valid])
        metrics: dict = {}

        resolved, used = await ensure_candidate_plausibility_ranges(
            candidates,
            ManifestSummary("input.xlsx", 2, 2, 2, 0, []),
            BillInfo("bill.pdf", "", [], cartons=2),
            {},
            llm,
            {},
            metrics=metrics,
        )

        self.assertTrue(used)
        self.assertEqual(len(resolved), 2)
        self.assertEqual(llm.calls, 3)
        self.assertEqual(metrics["batch_calls"], 1)
        self.assertEqual(metrics["individual_fallback_calls"], 2)

    def test_high_severity_review_can_only_tighten_existing_ranges(self) -> None:
        candidate = ProductCandidate(
            source="manifest_group",
            source_label="input.xlsx/HS归并",
            zh="水杯",
            en="Water cup",
            hs="3924104000",
            material="Plastic",
            usage="HOME",
            plausibility_range=PlausibilityRange(
                kg_per_ctn_min=2,
                kg_per_ctn_max=20,
                kg_per_pc_min=0.1,
                kg_per_pc_max=1,
                unit_price_min=0.5,
                unit_price_max=3,
                qty_per_ctn_min=5,
                qty_per_ctn_max=20,
                source="test",
            ),
        )
        review = {
            "issues": [
                {
                    "row_index": 1,
                    "severity": "high",
                    "suggested_bounds": {"unit_price_min": 1.2},
                }
            ]
        }

        adjusted, applied = apply_llm_review_constraints(
            [candidate],
            [{"单价": 0.8, "箱数": 10, "数量": 100, "毛重": 50}],
            review,
        )

        self.assertEqual(applied, 1)
        self.assertEqual(adjusted[0].plausibility_range.unit_price_min, 1.2)
        self.assertEqual(adjusted[0].plausibility_range.unit_price_max, 3)

    async def test_adaptive_replacement_query_stops_after_one_batch(self) -> None:
        seeds = [
            ProductCandidate(
                source="replacement",
                source_label="replacement.xlsx",
                zh=f"候选{index}",
                en=f"Candidate {index}",
                hs=f"39269099{index:02d}",
                material="Plastic",
                usage="HOME",
            )
            for index in range(6)
        ]
        crawler = RoutedFakeCrawler(
            {seed.hs: tax_result("5%", seed.hs) for seed in seeds}
        )

        qualified, filtered, cursor, attempts = await qualify_manual_invoice_replacement_batch(
            crawler,
            seeds,
            0,
            [],
            SelectionRules(),
            batch_size=4,
            query_cache={},
        )

        self.assertEqual(len(qualified), 4)
        self.assertEqual(filtered, [])
        self.assertEqual(cursor, 4)
        self.assertEqual(attempts, 4)
        self.assertEqual(len(crawler.hs_calls), 4)


def build_bill_candidates(count: int) -> list[ProductCandidate]:
    return [
        ProductCandidate(
            source="bill",
            source_label="bill.pdf",
            zh=f"品名{index}",
            en=f"Item {index}",
            hs=f"39269099{index:02d}",
            material="Plastic",
            usage="HOME",
            ctns=10,
            qty=100,
            gross_weight=50,
            effective_tax_rate=0.1,
            tax_match_source="bill_product",
        )
        for index in range(count)
    ]


def valid_plausibility_payload(basis: str) -> dict:
    return {
        "unit_price_min": 1,
        "unit_price_max": 3,
        "kg_per_pc_min": 0.1,
        "kg_per_pc_max": 1,
        "qty_per_ctn_min": 5,
        "qty_per_ctn_max": 20,
        "kg_per_ctn_min": 2,
        "kg_per_ctn_max": 20,
        "confidence": 0.9,
        "basis": basis,
    }


if __name__ == "__main__":
    unittest.main()
