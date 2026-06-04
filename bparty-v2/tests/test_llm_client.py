from __future__ import annotations

import httpx

from llm_client import should_retry_without_json_mode


def test_retry_without_json_mode_for_response_format_errors() -> None:
    response = httpx.Response(400, text="unknown field: response_format")

    assert should_retry_without_json_mode(response, {"response_format": {"type": "json_object"}})


def test_retry_without_json_mode_ignores_unrelated_errors() -> None:
    response = httpx.Response(401, text="invalid api key")

    assert not should_retry_without_json_mode(response, {"response_format": {"type": "json_object"}})
