from __future__ import annotations

import hashlib
import json
import re
import urllib.parse
import urllib.request
from dataclasses import asdict, dataclass, field
from pathlib import Path
from statistics import median
from typing import Any


DEFAULT_DECLARATION_RATIO = 0.3
PRICE_SEARCH_CACHE_VERSION = "search-v2"
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0 Safari/537.36"
)


@dataclass
class PriceEvidence:
    query: str
    samples: list[dict[str, Any]] = field(default_factory=list)
    retail_unit_price: float = 0.0
    declared_unit_price: float = 0.0
    basis: str = ""
    confidence: float = 0.0
    source: str = "web_search"

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def estimate_declared_unit_price_from_web(
    query: str,
    *,
    cache_dir: str | Path | None = None,
    timeout: float = 8.0,
    max_pages: int = 4,
    declaration_ratio: float = DEFAULT_DECLARATION_RATIO,
) -> PriceEvidence:
    query = clean_query(query)
    if not query:
        return empty_price_evidence("", "empty query")

    cache_path = price_cache_path(query, cache_dir)
    if cache_path and cache_path.exists():
        try:
            payload = json.loads(cache_path.read_text(encoding="utf-8"))
            return PriceEvidence(**payload)
        except Exception:
            pass

    html_blobs = fetch_search_pages(query, timeout=timeout, max_pages=max_pages)
    samples = extract_price_samples(html_blobs)
    evidence = build_price_evidence(query, samples, declaration_ratio=declaration_ratio)
    if cache_path:
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        cache_path.write_text(json.dumps(evidence.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
    return evidence


def build_price_evidence(
    query: str,
    samples: list[dict[str, Any]],
    *,
    declaration_ratio: float = DEFAULT_DECLARATION_RATIO,
) -> PriceEvidence:
    cleaned = filter_price_samples(samples)
    if not cleaned:
        return empty_price_evidence(query, "no usable public web price samples")

    unit_prices = sorted(float(item["unit_price_usd"]) for item in cleaned if item.get("unit_price_usd"))
    trimmed = trim_outliers(unit_prices)
    retail_unit = round(median(trimmed), 4)
    declared = round(max(0.0001, retail_unit * declaration_ratio), 4)
    confidence = min(0.9, 0.35 + min(len(cleaned), 12) * 0.04)
    return PriceEvidence(
        query=query,
        samples=cleaned[:12],
        retail_unit_price=retail_unit,
        declared_unit_price=declared,
        basis=f"public web search median unit retail price * {declaration_ratio:.0%}",
        confidence=round(confidence, 2),
        source="web_search",
    )


def extract_price_samples(html_blobs: list[str] | str) -> list[dict[str, Any]]:
    samples: list[dict[str, Any]] = []
    if isinstance(html_blobs, str):
        html_blobs = [html_blobs]
    for html in html_blobs:
        text = html_to_text(html)
        for match in re.finditer(r"\$\s?(\d{1,4}(?:,\d{3})*(?:\.\d{1,2})?)", text):
            price = parse_money(match.group(1))
            if price is None:
                continue
            window_start = max(0, match.start() - 220)
            window_end = min(len(text), match.end() + 220)
            context = normalize_space(text[window_start:window_end])
            if price <= 0:
                continue
            pack_qty = parse_pack_qty(context)
            samples.append(
                {
                    "title": context[:220],
                    "url": "",
                    "price_usd": round(price, 2),
                    "pack_qty": pack_qty,
                    "unit_price_usd": round(price / pack_qty, 4),
                    "source_domain": "",
                }
            )
    return samples


def filter_price_samples(samples: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    seen: set[tuple[float, int, str]] = set()
    blocked_terms = ("used", "pre-owned", "refurbished", "wholesale lot", "bulk lot")
    for sample in samples:
        title = str(sample.get("title") or "")
        normalized_title = title.lower()
        if any(term in normalized_title for term in blocked_terms):
            continue
        price = to_float(sample.get("price_usd"))
        unit = to_float(sample.get("unit_price_usd"))
        pack_qty = int(to_float(sample.get("pack_qty")) or 1)
        if price is None or unit is None or price <= 0 or unit <= 0:
            continue
        if price > 5000 or unit > 1000:
            continue
        if unit < 0.001:
            continue
        key = (round(unit, 4), pack_qty, normalized_title[:80])
        if key in seen:
            continue
        seen.add(key)
        result.append(
            {
                **sample,
                "price_usd": round(price, 2),
                "pack_qty": pack_qty,
                "unit_price_usd": round(unit, 4),
            }
        )
    return sorted(result, key=lambda item: item["unit_price_usd"])


def parse_pack_qty(text: str) -> int:
    normalized = text.lower().replace(",", "")
    patterns = (
        r"(\d{1,5})\s*[- ]\s*(?:pack|packs|pk)\b",
        r"(\d{1,5})\s*(?:pcs|pieces|piece|count|ct|pack|packs|pk)\b",
        r"(?:pack of|set of|lot of)\s*(\d{1,5})\b",
        r"(\d{1,5})\s*[- ]?(?:piece|pc)\s*(?:set|pack)",
    )
    for pattern in patterns:
        match = re.search(pattern, normalized)
        if not match:
            continue
        value = int(match.group(1))
        if 1 <= value <= 100000:
            return value
    return 1


def trim_outliers(values: list[float]) -> list[float]:
    if len(values) < 5:
        return values
    lower = int(len(values) * 0.15)
    upper = len(values) - lower
    trimmed = values[lower:upper]
    return trimmed or values


def fetch_search_pages(query: str, *, timeout: float, max_pages: int) -> list[str]:
    encoded = urllib.parse.quote_plus(query)
    urls = [
        f"https://search.aol.com/aol/search?q={encoded}",
        f"https://search.brave.com/search?q={encoded}",
        f"https://www.mojeek.com/search?q={encoded}",
        f"https://search.yahoo.com/search?p={encoded}",
        f"https://www.google.com/search?q={encoded}",
        f"https://www.bing.com/search?q={encoded}",
        f"https://duckduckgo.com/html/?q={encoded}",
        f"https://www.google.com/search?q={urllib.parse.quote_plus('site:walmart.com ' + query)}",
        f"https://www.google.com/search?q={urllib.parse.quote_plus('site:target.com ' + query)}",
        f"https://www.google.com/search?q={urllib.parse.quote_plus('site:ebay.com ' + query)}",
    ]
    blobs: list[str] = []
    for url in urls[:max_pages]:
        try:
            request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
            with urllib.request.urlopen(request, timeout=timeout) as response:
                content = response.read(1_000_000)
            blobs.append(content.decode("utf-8", errors="ignore"))
        except Exception:
            continue
    return blobs


def price_cache_path(query: str, cache_dir: str | Path | None) -> Path | None:
    if cache_dir is None:
        return None
    cache_key = f"{PRICE_SEARCH_CACHE_VERSION}\n{query}"
    digest = hashlib.sha256(cache_key.encode("utf-8", errors="ignore")).hexdigest()
    return Path(cache_dir) / f"{digest}.json"


def empty_price_evidence(query: str, basis: str) -> PriceEvidence:
    return PriceEvidence(query=query, basis=basis, confidence=0.0, source="fallback")


def clean_query(value: Any) -> str:
    return normalize_space(str(value or ""))[:180]


def html_to_text(value: str) -> str:
    text = re.sub(r"<script\b[^<]*(?:(?!</script>)<[^<]*)*</script>", " ", value, flags=re.I)
    text = re.sub(r"<style\b[^<]*(?:(?!</style>)<[^<]*)*</style>", " ", text, flags=re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    return normalize_space(urllib.parse.unquote(text))


def normalize_space(value: str) -> str:
    return re.sub(r"\s+", " ", value.replace("\xa0", " ")).strip()


def parse_money(value: Any) -> float | None:
    try:
        return float(str(value).replace(",", "").strip())
    except Exception:
        return None


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    try:
        return float(value)
    except Exception:
        return None
