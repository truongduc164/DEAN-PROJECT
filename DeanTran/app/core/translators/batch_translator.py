"""
GeminiBatchTranslator – high-performance batched translation via Gemini.

Collects cells into batches (max 100 items / 15000 chars), sends one
Gemini request per batch returning strict JSON. Includes retry,
split-on-failure, translation cache, and full traceback logging.
"""
from __future__ import annotations

import hashlib
import json
import logging
import threading
import time
import traceback
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

from app.settings.settings_manager import settings
from app.core.translators.protection import TextProtector

logger = logging.getLogger("DeanTran.batch_translator")


def _extract_retry_delay(exc_str: str) -> float | None:
    """Parse retryDelay value from API error message.

    The Gemini API returns strings like:
        retryDelay: 35s
        retry_delay { seconds: 35 }
    Returns seconds as float, or None if not found.
    """
    import re
    # Pattern: retryDelay: 35s
    m = re.search(r'retryDelay\s*[:=]\s*(\d+\.?\d*)\s*s', exc_str, re.IGNORECASE)
    if m:
        return float(m.group(1))
    # Pattern: seconds: 35
    m = re.search(r'seconds\s*[:=]\s*(\d+\.?\d*)', exc_str, re.IGNORECASE)
    if m:
        return float(m.group(1))
    # Pattern: retry after 35 seconds
    m = re.search(r'retry\s+after\s+(\d+\.?\d*)\s*s', exc_str, re.IGNORECASE)
    if m:
        return float(m.group(1))
    return None

# ── Batch config ─────────────────────────────────────────────────────

MAX_ITEMS_PER_BATCH = 200
MAX_CHARS_PER_BATCH = 30000
DEFAULT_MIN_INTERVAL = 0.2

# Fallback models when primary model hits rate limit
FALLBACK_MODELS = [
    "gemini-2.5-flash",
    "gemini-2.5-lite",
]


@dataclass
class CellItem:
    """One translatable cell."""
    cell_id: str          # e.g. "Sheet1!A1"
    original: str         # original text
    translated: str = ""  # filled after translation
    error: str = ""       # filled on failure


@dataclass
class BatchResult:
    """Result of processing all batches."""
    items: list[CellItem] = field(default_factory=list)
    api_calls: int = 0
    cache_hits: int = 0
    errors: int = 0
    elapsed: float = 0.0


class TranslationCache:
    """In-memory + optional disk cache for translations."""

    def __init__(self, cache_path: Optional[Path] = None) -> None:
        self._mem: dict[str, str] = {}
        self._path = cache_path
        self._load_disk()

    def _cache_key(self, mode: str, text: str) -> str:
        return hashlib.md5(f"{mode}|{text}".encode()).hexdigest()

    def get(self, mode: str, text: str) -> Optional[str]:
        return self._mem.get(self._cache_key(mode, text))

    def put(self, mode: str, text: str, translation: str) -> None:
        self._mem[self._cache_key(mode, text)] = translation

    def save_disk(self) -> None:
        if self._path:
            try:
                self._path.parent.mkdir(parents=True, exist_ok=True)
                self._path.write_text(
                    json.dumps(self._mem, ensure_ascii=False),
                    encoding="utf-8",
                )
            except OSError as exc:
                logger.debug("Cache save failed: %s", exc)

    def _load_disk(self) -> None:
        if self._path and self._path.exists():
            try:
                self._mem = json.loads(self._path.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, OSError):
                self._mem = {}

    @property
    def size(self) -> int:
        return len(self._mem)


def _normalize_model(model: str) -> str:
    """Ensure model string has 'models/' prefix if needed by the SDK."""
    return model


class GeminiBatchTranslator:
    """
    High-performance batch translator using google.genai SDK.
    Supports pause/cancel, adaptive batch sizing, and batch completeness validation.
    """

    def __init__(
        self,
        key_provider,
        model_name: str = "gemini-3.1-flash-lite",
        source_lang: str = "Chinese",
        target_lang: str = "Vietnamese",
        mode: str = "SOP",
        prompt: str = "",
        min_interval: float = DEFAULT_MIN_INTERVAL,
        log_fn: Optional[Callable] = None,
        cache_dir: Optional[Path] = None,
        pause_event: Optional[threading.Event] = None,
        cancel_event: Optional[threading.Event] = None,
        progress_fn: Optional[Callable[[int, int], None]] = None,
    ) -> None:
        self._kp = key_provider
        self._model = _normalize_model(model_name)
        self._source = source_lang
        self._target = target_lang
        self._mode = mode
        self._prompt = prompt
        self._min_interval = min_interval
        self._log = log_fn or (lambda lvl, msg: None)
        self._progress = progress_fn or (lambda cur, tot: None)
        self._last_call_time: float = 0

        # Fallback models (exclude primary to avoid duplicate)
        self._fallback_models = [m for m in FALLBACK_MODELS if m != self._model]
        self._exhausted_models: set[str] = set()

        # Cache
        cache_path = (cache_dir / "cache.json") if cache_dir else None
        self._cache = TranslationCache(cache_path)

        # Pause / Cancel events
        self._pause_event = pause_event or threading.Event()
        if pause_event is None:
            self._pause_event.set()  # default = running
        self._cancel_event = cancel_event or threading.Event()

        # Dynamic Configs
        self._max_items_per_batch = settings.get("limits_settings.max_items_per_batch", 200)
        self._max_chars_per_batch = settings.get("limits_settings.max_chars_per_batch", 30000)
        self._retry_limit = settings.get("limits_settings.retry_limit", 3)
        self._api_timeout = settings.get("limits_settings.api_timeout", 60)
        self._cb_max_failures = settings.get("limits_settings.circuit_breaker_failures", 5)

        # Protection
        self._protector = TextProtector()
        
        # Adaptive batch sizing
        self._current_batch_size = self._max_items_per_batch
        self._consecutive_successes = 0
        self._MIN_BATCH_SIZE = 5
        self._circuit_breaker_failures = 0

    # ── Source-language filter ────────────────────────────────

    @staticmethod
    def _contains_source_lang(text: str, source_lang: str) -> bool:
        """Check if text contains characters of the source language."""
        if not text or not text.strip():
            return False
        
        import re
        if source_lang == "Chinese":
            # CJK Unified Ideographs + Extension A
            return bool(re.search(r'[一-鿿㐀-䶿]', text))
        elif source_lang == "Vietnamese":
            # Vietnamese-specific chars with diacritics
            return bool(re.search(r'[À-ïĂăĐđĨĩŨũƠơƯưẠ-ỹ]', text))
        elif source_lang == "English":
            # Contains Latin letters but NO Chinese/Vietnamese-specific chars
            has_latin = bool(re.search(r'[a-zA-Z]', text))
            has_cjk = bool(re.search(r'[一-鿿]', text))
            has_vi = bool(re.search(r'[À-ïĂăĐđẠ-ỹ]', text))
            return has_latin and not has_cjk and not has_vi
        else:
            # Unknown source language - translate everything
            return True

    def translate_batch(self, items: list[CellItem]) -> BatchResult:
        """Translate a list of CellItems using batching."""
        t0 = time.time()
        result = BatchResult(items=items)

        # Check cache and protection
        uncached: list[CellItem] = []
        self._item_tokens = {} # cell_id -> tokens
        total_items = len(items)
        done_items = 0

        for item in items:
            if item.translated:
                continue
            # Source-language filter: skip non-source-language text
            if settings.get("processing_options.translate_source_lang_only", False):
                if not self._contains_source_lang(item.original, self._source):
                    item.translated = item.original
                    result.cache_hits += 1
                    done_items += 1
                    continue
            cached = self._cache.get(self._mode, item.original)
            if cached:
                item.translated = cached
                result.cache_hits += 1
                continue
                
            modified_text, tokens, fully_protected = self._protector.protect(item.original)
            self._item_tokens[item.cell_id] = tokens
            
            if fully_protected:
                item.translated = self._protector.restore(modified_text, tokens)
                # Count as cache hit functionally, avoiding API
                result.cache_hits += 1
                done_items += 1
                self._cache.put(self._mode, item.original, item.translated)
            else:
                # API Call needed, use modified text and let system know original is temporarily modified
                item.original_pre_protect = item.original
                item.original = modified_text 
                uncached.append(item)

        if result.cache_hits > 0:
            self._log("INFO", f"Cache/Protected hits: {result.cache_hits}/{total_items}")
            self._progress(done_items, total_items)

        if not uncached:
            result.elapsed = time.time() - t0
            return result

        # Build batches (using adaptive size)
        batches = self._build_batches(uncached)
        self._log("INFO", f"batch_count={len(batches)} uncached_cells={len(uncached)} batch_size={self._current_batch_size}")

        # Process each batch
        for i, batch in enumerate(batches, 1):
            # Check cancel
            if self._cancel_event.is_set():
                self._log("WARN", "Cancel requested – stopping batch processing.")
                for remaining_item in batch:
                    if not remaining_item.translated:
                        remaining_item.error = "Cancelled by user"
                        result.errors += 1
                break

            # Check pause (blocks without busy loop)
            self._pause_event.wait()

            chars = sum(len(it.original) for it in batch)
            self._log(
                "INFO",
                f"[batch_started] batch={i}/{len(batches)} items={len(batch)} "
                f"chars={chars} model={self._model}",
            )
            api_calls, errors = self._process_batch(batch, batch_idx=i, total_batches=len(batches))
            result.api_calls += api_calls
            result.errors += errors
            
            done_items += len(batch)
            self._progress(done_items, total_items)

            self._log(
                "INFO",
                f"[batch_completed] batch={i}/{len(batches)} "
                f"errors={errors} model={self._model}",
            )

            if errors == 0:
                self._circuit_breaker_failures = 0
                self._consecutive_successes += 1
                if self._consecutive_successes >= 3 and self._current_batch_size < self._max_items_per_batch:
                    old = self._current_batch_size
                    self._current_batch_size = min(
                        self._max_items_per_batch,
                        int(self._current_batch_size * 1.25),
                    )
                    self._log("INFO", f"Adaptive: batch size {old} → {self._current_batch_size}")
                    self._consecutive_successes = 0
            else:
                self._consecutive_successes = 0
                self._circuit_breaker_failures += 1
                if self._circuit_breaker_failures >= self._cb_max_failures:
                    self._log("ERROR", "CIRCUIT BREAKER TRIGGERED: Too many consecutive failures. Halting.")
                    self._cancel_event.set()

        # Save cache and restore originals
        for item in uncached:
            if hasattr(item, "original_pre_protect"):
                # Restore tokens
                if item.translated and not item.error:
                    tokens = self._item_tokens.get(item.cell_id, [])
                    item.translated = self._protector.restore(item.translated, tokens)
                    self._cache.put(self._mode, item.original_pre_protect, item.translated)
                    
                item.original = item.original_pre_protect

        self._cache.save_disk()

        result.elapsed = time.time() - t0
        return result

    def _build_batches(self, items: list[CellItem]) -> list[list[CellItem]]:
        batches: list[list[CellItem]] = []
        current: list[CellItem] = []
        current_chars = 0
        max_items = self._current_batch_size

        for item in items:
            text_len = len(item.original)
            if (len(current) >= max_items
                    or (current_chars + text_len > self._max_chars_per_batch and current)):
                batches.append(current)
                current = []
                current_chars = 0
            current.append(item)
            current_chars += text_len

        if current:
            batches.append(current)
        return batches

    def _process_batch(
        self,
        batch: list[CellItem],
        depth: int = 0,
        batch_idx: int = 0,
        total_batches: int = 0,
    ) -> tuple[int, int]:
        """
        Process one batch. Returns (api_calls, errors).
        On parse failure: retry once, then split in half recursively.
        """
        if depth > 3:
            err_msg = f"Max recursion depth reached, {len(batch)} cells untranslated"
            self._log("ERROR", err_msg)
            for item in batch:
                item.error = err_msg
            return 0, len(batch)

        # Rate limit
        self._throttle()

        # Build prompt
        payload = [{"id": it.cell_id, "text": it.original} for it in batch]
        system_prompt = self._prompt or (
            f"Translate from {self._source} to {self._target}. "
            f"Return ONLY the translated text."
        )
        user_prompt = (
            f"{system_prompt}\n\n"
            f"Translate ALL {len(batch)} text items from {self._source} "
            f"to {self._target}. Do NOT skip any item. "
            f"Return a JSON array containing exactly {len(batch)} objects, "
            f"each with the original 'id' and a new 'translated' field. "
            f"Return ONLY valid JSON, no markdown, no explanation.\n\n"
            f"Input:\n{json.dumps(payload, ensure_ascii=False)}"
        )
        # Check active translation tool
        tool = settings.get("translation_tool", "Gemini")

        if tool == "DeepSeek":
            # Use settings-based max batch size; split oversized batches instead of truncating
            deepseek_max_items = self._max_items_per_batch
            if len(batch) > deepseek_max_items:
                self._log(
                    "INFO",
                    f"DeepSeek batch oversized ({len(batch)} > {deepseek_max_items}), "
                    f"splitting into sub-batches"
                )
                total_calls = 0
                total_errors = 0
                for i in range(0, len(batch), deepseek_max_items):
                    sub_batch = batch[i:i + deepseek_max_items]
                    calls, errs = self._process_batch(
                        sub_batch, depth + 1, batch_idx, total_batches
                    )
                    total_calls += calls
                    total_errors += errs
                return total_calls, total_errors
            import os
            from app.core.secure_storage import SecureStorage
            key = SecureStorage().load_deepseek_key()
            if not key:
                key = os.environ.get("DEEPSEEK_API_KEY", "")
            if not key:
                err_msg = "No DeepSeek API key available"
                self._log("ERROR", err_msg)
                for item in batch:
                    item.error = err_msg
                return 0, len(batch)

            key_prefix = key[:6] + "***" if len(key) >= 6 else "(short)"
            deepseek_model = settings.get("selected_models.deepseek", "deepseek-chat")
            self._log(
                "INFO",
                f"DEEPSEEK_CALL batch_idx={batch_idx} items={len(batch)} "
                f"chars={sum(len(it.original) for it in batch)} "
                f"model={deepseek_model} key_loaded=True key_prefix={key_prefix}",
            )

            import requests
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {key}"
            }
            ds_payload = {
                "model": deepseek_model,
                "messages": [
                    {"role": "user", "content": user_prompt}
                ],
                "temperature": 1.0,
                "max_tokens": 65536,
            }
            # Disable thinking mode for V4 models (faster + cheaper for translation)
            if "v4" in deepseek_model:
                ds_payload["thinking"] = {"type": "disabled"}
                ds_payload["temperature"] = 0.2

            max_attempts = self._retry_limit
            for attempt in range(max_attempts):
                try:
                    self._throttle()
                    t0 = time.time()
                    resp = requests.post(
                        "https://api.deepseek.com/chat/completions",
                        headers=headers,
                        json=ds_payload,
                        timeout=600,
                        verify=False
                    )
                    elapsed = time.time() - t0
                    
                    if resp.status_code == 429:
                        wait_time = 2 ** (attempt + 2)
                        self._log("WARN", f"DeepSeek Rate Limit 429. Waiting {wait_time}s...")
                        if self._cancel_aware_sleep(wait_time):
                            return 0, len(batch)
                        continue
                        
                    resp.raise_for_status()
                    res_data = resp.json()
                    raw = res_data["choices"][0]["message"]["content"] or ""
                    raw = raw.strip()
                    
                    self._log("INFO", f"response_text_len={len(raw)}")
                    
                    if not raw:
                        raise ValueError("Empty response from DeepSeek")

                    # Strip think tags if present
                    if "<think>" in raw:
                        end_think = raw.find("</think>")
                        if end_think != -1:
                            raw = raw[end_think + len("</think>"):].strip()

                    # Strip markdown fences if present
                    if raw.startswith("```"):
                        lines = raw.split("\n")
                        raw = "\n".join(lines[1:-1] if lines[-1].strip() == "```" else lines[1:])
                        raw = raw.strip()

                    translations = json.loads(raw)
                    if isinstance(translations, dict):
                        for k, v in translations.items():
                            if isinstance(v, list):
                                translations = v
                                break
                    if not isinstance(translations, list):
                        raise ValueError("Response content is not a JSON array")

                    # Map translations back – use "translated" field
                    trans_map = {}
                    for t in translations:
                        if not isinstance(t, dict):
                            continue
                        tid = t.get("id", "")
                        tval = t.get("translated", "") or t.get("vi", "") or t.get("translation", "")
                        if tid and tval:
                            trans_map[tid] = tval

                    matched = 0
                    missing_items: list[CellItem] = []
                    for item in batch:
                        if item.cell_id in trans_map and trans_map[item.cell_id]:
                            item.translated = trans_map[item.cell_id]
                            self._cache.put(self._mode, item.original, item.translated)
                            matched += 1
                            
                            orig_disp = item.original.replace('\n', ' ')[:40]
                            trans_disp = item.translated.replace('\n', ' ')[:40]
                            self._log("INFO", f"[{item.cell_id}] {orig_disp} ➔ {trans_disp}")
                        elif not item.translated:
                            item.error = "No translation in response"
                            missing_items.append(item)

                    self._log("INFO", f"DeepSeek call ok in {elapsed:.1f}s matched={matched}/{len(batch)}")

                    # Batch completeness validation
                    if missing_items:
                        retry_calls, retry_errs = self._retry_missing(missing_items, depth)
                        return 1 + retry_calls, retry_errs

                    return 1, len(batch) - matched

                except json.JSONDecodeError as exc:
                    self._log("WARN", f"JSON parse failed (attempt {attempt+1}): {exc}")
                    if attempt < max_attempts - 1:
                        time.sleep(2 ** attempt)
                        continue
                    return self._split_and_retry(batch, depth)
                except Exception as exc:
                    self._log("ERROR", f"DeepSeek error (attempt {attempt+1}): {exc}")
                    if attempt < max_attempts - 1:
                        time.sleep(2 ** attempt)
                        continue
                    err_msg = f"[api_error] {type(exc).__name__}: {exc}"
                    for item in batch:
                        item.error = err_msg
                    return 1, len(batch)
                    
            err_msg = "DeepSeek retries exhausted"
            for item in batch:
                item.error = err_msg
            return max_attempts, len(batch)

        # Get key
        key = self._kp.get_key()
        if not key:
            err_msg = "No API key available"
            self._log("ERROR", err_msg)
            for item in batch:
                item.error = err_msg
            return 0, len(batch)

        key_prefix = key[:6] + "***" if len(key) >= 6 else "(short)"
        is_single_key = self._kp.key_count <= 1

        self._log(
            "INFO",
            f"GEMINI_CALL batch_idx={batch_idx} items={len(batch)} "
            f"chars={sum(len(it.original) for it in batch)} "
            f"model={self._model} key_loaded=True key_prefix={key_prefix}",
        )

        max_attempts = max(3, int(getattr(self._kp, "key_count", 1) or 1))
        for attempt in range(max_attempts):
            try:
                from google import genai  # type: ignore
                client = genai.Client(api_key=key)
                t0 = time.time()
                response = client.models.generate_content(
                    model=self._model,
                    contents=user_prompt,
                )
                elapsed = time.time() - t0

                # Parse JSON response
                raw_text = getattr(response, "text", "") or ""
                raw = raw_text.strip()
                self._log(
                    "INFO",
                    f"response_text_len={len(raw)}",
                )

                if not raw:
                    raise ValueError("Empty response from Gemini")

                # Strip markdown fences if present
                if raw.startswith("```"):
                    lines = raw.split("\n")
                    raw = "\n".join(lines[1:-1] if lines[-1].strip() == "```" else lines[1:])
                    raw = raw.strip()

                translations = json.loads(raw)
                if not isinstance(translations, list):
                    raise ValueError("Response is not a JSON array")

                # Map translations back – use "translated" field
                trans_map = {}
                for t in translations:
                    tid = t.get("id", "")
                    tval = t.get("translated", "") or t.get("vi", "") or t.get("translation", "")
                    if tid and tval:
                        trans_map[tid] = tval

                matched = 0
                missing_items: list[CellItem] = []
                for item in batch:
                    if item.cell_id in trans_map and trans_map[item.cell_id]:
                        item.translated = trans_map[item.cell_id]
                        self._cache.put(self._mode, item.original, item.translated)
                        matched += 1
                        
                        orig_disp = item.original.replace('\n', ' ')[:40]
                        trans_disp = item.translated.replace('\n', ' ')[:40]
                        self._log("INFO", f"[{item.cell_id}] {orig_disp} ➔ {trans_disp}")
                    elif not item.translated:
                        item.error = "No translation in response"
                        missing_items.append(item)

                self._kp.report_success(key)

                self._log(
                    "INFO",
                    f"gemini call ok in {elapsed:.1f}s matched={matched}/{len(batch)}"
                )

                # Batch completeness validation: retry missing items
                if missing_items:
                    retry_calls, retry_errs = self._retry_missing(missing_items, depth)
                    return 1 + retry_calls, retry_errs

                return 1, len(batch) - matched

            except json.JSONDecodeError as exc:
                self._log("WARN", f"JSON parse failed (attempt {attempt+1}): {exc}")
                # Adaptive: reduce batch size on parse failure
                old_size = self._current_batch_size
                self._current_batch_size = max(self._MIN_BATCH_SIZE, self._current_batch_size // 2)
                if old_size != self._current_batch_size:
                    self._log("INFO", f"Adaptive: batch size {old_size} → {self._current_batch_size} (parse fail)")
                self._consecutive_successes = 0
                if attempt < 1:
                    continue
                return self._split_and_retry(batch, depth)

            except Exception as exc:
                exc_str = str(exc).lower()
                tb = traceback.format_exc()

                self._log(
                    "ERROR",
                    f"GEMINI_ERROR type={type(exc).__name__} msg={exc}",
                )
                self._log("ERROR", f"Traceback:\n{tb}")

                retryable = any(k in exc_str for k in (
                    "429", "503", "resource_exhausted", "overloaded",
                    "unavailable", "capacity", "deadline",
                ))
                auth_error = any(k in exc_str for k in ("401", "403", "invalid", "api_key"))
                model_not_found = (
                    "404" in exc_str
                    and (
                        "not_found" in exc_str
                        or "is not found for api version" in exc_str
                        or "not supported for generatecontent" in exc_str
                    )
                )

                if retryable:
                    # Try to parse retryDelay from API response
                    parsed_delay = _extract_retry_delay(exc_str)
                    if parsed_delay:
                        # Use parsed delay with exponential backoff
                        multipliers = [1.0, 1.5, 2.0]
                        mult = multipliers[min(attempt, len(multipliers) - 1)]
                        wait_time = parsed_delay * mult
                        self._log(
                            "WARN",
                            f"[retry_wait] seconds={wait_time:.0f} "
                            f"retryDelay={parsed_delay}s multiplier={mult} "
                            f"attempt={attempt+1}/3 model={self._model}",
                        )
                    else:
                        # Fallback: exponential backoff 4s, 8s, 16s
                        wait_time = 2 ** (attempt + 2)
                        self._log(
                            "WARN",
                            f"[retry_wait] seconds={wait_time} "
                            f"(no retryDelay parsed, using fallback) "
                            f"attempt={attempt+1}/3 model={self._model}",
                        )

                    self._log(
                        "WARN",
                        f"⚠ Quota exceeded. Waiting {wait_time:.0f} seconds to resume…",
                    )

                    if is_single_key:
                        if self._cancel_aware_sleep(wait_time):
                            self._log("WARN", "Cancel during quota wait – stopping immediately.")
                            for item in batch:
                                if not item.translated:
                                    item.error = "Cancelled during quota wait"
                            return 1, len([i for i in batch if not i.translated])
                        continue
                    else:
                        error_code = 429 if "429" in exc_str else 503
                        
                        # Get current index before rotating
                        old_idx = getattr(self._kp, "current_index", 0)
                        
                        self._kp.report_error(key, error_code)
                        if self._cancel_aware_sleep(wait_time):
                            self._log("WARN", "Cancel during quota wait – stopping immediately.")
                            for item in batch:
                                if not item.translated:
                                    item.error = "Cancelled during quota wait"
                            return 1, len([i for i in batch if not i.translated])
                        key = self._kp.get_key()
                        if key:
                            new_idx = getattr(self._kp, "current_index", 0)
                            self._log("WARN", f"quota exceeded on key index={old_idx}")
                            self._log("INFO", f"switched to key index={new_idx}")
                            self._log("INFO", "retrying same batch with next key")
                        else:
                            err_msg = "All Gemini API keys are exhausted."
                            self._log("ERROR", f"[api_error] {err_msg}")
                            for item in batch:
                                item.error = err_msg
                            return 1, len(batch)
                        continue

                if auth_error:
                    if not is_single_key:
                        old_idx = getattr(self._kp, "current_index", 0)
                        self._kp.report_error(key, 401)
                        key = self._kp.get_key()
                        if key:
                            new_idx = getattr(self._kp, "current_index", 0)
                            self._log("WARN", f"auth failed on key index={old_idx}")
                            self._log("INFO", f"switched to key index={new_idx}")
                            self._log("INFO", "retrying same batch with next key")
                            continue
                        err_msg = "All Gemini API keys are invalid or exhausted."
                        self._log("ERROR", f"[api_error] {err_msg}")
                        for item in batch:
                            item.error = err_msg
                        return 1, len(batch)

                    if "expired" in exc_str:
                        err_msg = (
                            "[api_error] API key expired. "
                            "Please renew/update Gemini API key in Admin > API Configuration."
                        )
                    else:
                        err_msg = (
                            "[api_error] API key invalid/unauthorized. "
                            "Please update Gemini API key in Admin > API Configuration."
                        )
                    self._log("ERROR", err_msg)
                    for item in batch:
                        item.error = err_msg
                    return 1, len(batch)

                if model_not_found:
                    self._log(
                        "WARN",
                        f"Model {self._model} not available for this API key/version; trying fallback models.",
                    )
                    break

                # Non-retryable error
                err_msg = f"[api_error] {type(exc).__name__}: {exc}"
                for item in batch:
                    item.error = err_msg
                return 1, len(batch)

        # All retry attempts exhausted on current model — try fallback models
        self._exhausted_models.add(self._model)
        for fb_model in self._fallback_models:
            if fb_model in self._exhausted_models:
                continue
            self._log(
                "WARN",
                f"Model {self._model} exhausted → falling back to {fb_model}",
            )
            old_model = self._model
            self._model = fb_model
            calls, errs = self._process_batch(
                batch, depth=depth, batch_idx=batch_idx,
                total_batches=total_batches,
            )
            if errs < len(batch):
                return calls, errs
            self._model = old_model
            self._exhausted_models.add(fb_model)

        # All models exhausted
        err_msg = f"All models exhausted ({', '.join(self._exhausted_models)})"
        self._log("ERROR", err_msg)
        for item in batch:
            if not item.translated:
                item.error = err_msg
        return 1, len(batch)

    def _split_and_retry(self, batch: list[CellItem], depth: int) -> tuple[int, int]:
        if len(batch) <= 1:
            self._log("WARN", f"Single-item batch failed: {batch[0].cell_id}")
            batch[0].error = "Translation failed after split"
            return 0, 1

        mid = len(batch) // 2
        self._log("INFO", f"Splitting batch ({len(batch)}) into {mid} + {len(batch)-mid}")
        calls1, err1 = self._process_batch(batch[:mid], depth + 1)
        calls2, err2 = self._process_batch(batch[mid:], depth + 1)
        return calls1 + calls2, err1 + err2

    def _retry_missing(
        self,
        missing_items: list[CellItem],
        depth: int,
        max_retries: int = 3,
    ) -> tuple[int, int]:
        """Retry ONLY the missing items from a batch, up to max_retries times."""
        total_calls = 0
        for retry in range(1, max_retries + 1):
            # Clear errors for retry
            for item in missing_items:
                item.error = ""

            self._log(
                "INFO",
                f"Retry missing items: attempt {retry}/{max_retries}, "
                f"items={len(missing_items)}",
            )
            calls, errors = self._process_batch(
                missing_items, depth=depth + 1,
            )
            total_calls += calls

            # Check which are still missing
            still_missing = [it for it in missing_items if not it.translated]
            if not still_missing:
                self._log("INFO", f"All missing items resolved on retry {retry}")
                return total_calls, 0

            missing_items = still_missing

        # Max retries exhausted
        self._log(
            "ERROR",
            f"{len(missing_items)} items still missing after {max_retries} retries",
        )
        for item in missing_items:
            item.error = f"Still missing after {max_retries} retries"
        return total_calls, len(missing_items)

    def _throttle(self) -> None:
        now = time.time()
        elapsed = now - self._last_call_time
        if elapsed < self._min_interval:
            time.sleep(self._min_interval - elapsed)
        self._last_call_time = time.time()

    def _cancel_aware_sleep(self, seconds: float) -> bool:
        """Sleep for `seconds` but wake up immediately if cancel_event is set.

        Returns True if cancelled during the sleep.
        """
        end = time.time() + seconds
        while time.time() < end:
            if self._cancel_event.is_set():
                return True
            time.sleep(min(0.5, end - time.time()))
        return self._cancel_event.is_set()
