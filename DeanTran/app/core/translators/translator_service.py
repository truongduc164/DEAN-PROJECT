"""
TranslatorService – base interface, MockTranslator, and GeminiTranslator.

Usage
-----
- **Tests / dry-run**: always use ``MockTranslator`` (no API key needed).
- **Production**: call ``create_translator()`` which returns ``GeminiTranslator``
  only if an API key is available, otherwise falls back to ``MockTranslator``.
"""
from __future__ import annotations

import logging
import os
import time
from abc import ABC, abstractmethod

logger = logging.getLogger("DeanTran.translator")

DEFAULT_GEMINI_MODEL = "gemini-3.1-flash-lite"
GEMINI_FALLBACK_MODELS = [
    "gemini-2.5-flash",
    "gemini-2.5-lite",
]


class BaseTranslator(ABC):
    """Abstract interface for all translator backends."""

    @abstractmethod
    def translate(self, text: str, source_lang: str, target_lang: str) -> str:
        """Translate *text* from *source_lang* to *target_lang*."""
        ...


class MockTranslator(BaseTranslator):
    """
    Deterministic translator for testing & dry-run mode.
    Returns ``[<target_lang>] <text>`` so assertions are predictable.
    **Never** requires an API key.

    Parameters
    ----------
    delay : float
        Optional per-call sleep (seconds). Defaults to 0 for fast tests.
    """

    def __init__(self, delay: float = 0.0) -> None:
        self.delay = delay
        self.call_count: int = 0

    def translate(self, text: str, source_lang: str, target_lang: str) -> str:
        if self.delay:
            time.sleep(self.delay)
        self.call_count += 1
        return f"[{target_lang}] {text}"


class GeminiTranslator(BaseTranslator):
    """
    Translator backed by Google Gemini API using the new ``google-genai`` SDK.

    Features:
    - Exponential backoff retry on 503 / capacity exhausted / timeout
    - Automatic fallback across free-tier models if primary is unavailable

    Parameters
    ----------
    api_key : str
        Google AI API key.
    model_name : str
        Gemini model identifier (default ``gemini-3.1-flash-lite``).
    """

    _RETRY_DELAYS = [30, 60, 120]  # seconds
    _FALLBACK_MODELS = GEMINI_FALLBACK_MODELS

    def __init__(
        self,
        api_key: str,
        model_name: str = DEFAULT_GEMINI_MODEL,
    ) -> None:
        from google import genai  # type: ignore[import-untyped]

        self._client = genai.Client(api_key=api_key)
        self._model_name = model_name
        logger.info(
            "GeminiTranslator created: model=%s key_prefix=%s****",
            model_name, api_key[:6] if len(api_key) >= 6 else "***",
        )

    def translate(self, text: str, source_lang: str, target_lang: str) -> str:
        prompt = (
            f"Translate the following text from {source_lang} to {target_lang}. "
            f"Return ONLY the exact translated text without any explanation, conversational filler, or 'THINK' blocks.\n\n{text}"
        )
        return self._call_with_retry(prompt)

    def _call_with_retry(self, prompt: str) -> str:
        """Call Gemini with retries and model fallback."""
        last_exc: Exception | None = None
        models_to_try = [self._model_name]
        for fallback_model in self._FALLBACK_MODELS:
            if fallback_model not in models_to_try:
                models_to_try.append(fallback_model)

        for model_name in models_to_try:
            for attempt, delay in enumerate(self._RETRY_DELAYS, 1):
                try:
                    response = self._client.models.generate_content(
                        model=model_name,
                        contents=prompt,
                    )
                    out_text = getattr(response, 'text', '') or ''
                    return out_text.strip()
                except Exception as exc:
                    last_exc = exc
                    exc_str = str(exc).lower()
                    retryable = any(
                        kw in exc_str
                        for kw in ("503", "capacity", "overloaded", "timeout",
                                   "resource_exhausted", "unavailable", "deadline")
                    )
                    if not retryable:
                        logger.warning(
                            "Non-retryable error on model=%s: %s: %s",
                            model_name, type(exc).__name__, exc,
                        )
                        break  # try next model
                    logger.warning(
                        "Retry %d/%d on model=%s (wait %ds): %s",
                        attempt, len(self._RETRY_DELAYS), model_name, delay, exc,
                    )
                    time.sleep(delay)

            logger.warning("All retries exhausted for model=%s", model_name)

        # If we get here, all models failed — raise so factory can fallback
        raise RuntimeError(
            f"Gemini failed after all retries and model fallback: {last_exc}"
        )


# ── DeepSeek Translator ──────────────────────────────────────────────

class DeepSeekTranslator(BaseTranslator):
    """
    Translator backed by DeepSeek API using requests.
    """

    def __init__(
        self,
        api_key: str,
        model_name: str = "deepseek-chat",
    ) -> None:
        self._api_key = api_key
        self._model_name = model_name
        logger.info(
            "DeepSeekTranslator created: model=%s key_prefix=%s****",
            model_name, api_key[:6] if len(api_key) >= 6 else "***",
        )

    def translate(self, text: str, source_lang: str, target_lang: str) -> str:
        prompt = (
            f"Translate the following text from {source_lang} to {target_lang}. "
            f"Return ONLY the exact translated text without any explanation, conversational filler, or 'THINK' blocks.\n\n{text}"
        )
        return self._call_api(prompt)

    def _call_api(self, prompt: str) -> str:
        import requests
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self._api_key}"
        }
        payload = {
            "model": self._model_name,
            "messages": [
                {"role": "user", "content": prompt}
            ],
            "temperature": 1.0
        }
        # Disable thinking mode for V4 models (faster + cheaper for translation)
        if "v4" in self._model_name:
            payload["thinking"] = {"type": "disabled"}
            payload["temperature"] = 0.2
        
        last_exc = None
        for attempt in range(3):
            try:
                response = requests.post(
                    "https://api.deepseek.com/chat/completions",
                    headers=headers,
                    json=payload,
                    timeout=60,
                    verify=False
                )
                if response.status_code == 429:
                    time.sleep(2 ** (attempt + 2))
                    continue
                response.raise_for_status()
                res_data = response.json()
                out_text = res_data["choices"][0]["message"]["content"] or ""
                
                # Strip think tags if present
                if "<think>" in out_text:
                    end_think = out_text.find("</think>")
                    if end_think != -1:
                        out_text = out_text[end_think + len("</think>"):].strip()
                        
                return out_text.strip()
            except Exception as exc:
                last_exc = exc
                time.sleep(2 ** attempt)

        raise RuntimeError(f"DeepSeek translation failed: {last_exc}")


# ── Factory ──────────────────────────────────────────────────────────

def create_translator(
    *,
    api_key: str | None = None,
    model_name: str | None = None,
    force_mock: bool = False,
    use_secure_storage: bool = True,
) -> BaseTranslator:
    """
    Create the best available translator.

    Model name comes from SettingsManager.
    API key comes from SecureStorage (keyring) — never from settings JSON.
    """
    if force_mock:
        logger.info("force_mock=True → MockTranslator")
        return MockTranslator()

    try:
        from app.settings.settings_manager import settings
        tool = settings.get("translation_tool", "Gemini")
    except Exception:
        tool = "Gemini"

    if tool == "DeepSeek":
        # Resolve model from settings if not explicitly given
        if not model_name:
            try:
                from app.settings.settings_manager import settings
                model_name = settings.get("selected_models.deepseek", "deepseek-chat")
            except Exception:
                model_name = "deepseek-chat"

        # Explicit key
        key = api_key or ""

        # Try secure storage (keyring)
        if not key and use_secure_storage:
            try:
                from app.core.secure_storage import SecureStorage
                ss = SecureStorage()
                key = ss.load_deepseek_key() or ""
            except Exception as exc:
                logger.warning("Failed to load DeepSeek key from secure storage: %s", exc)

        # Try env var
        if not key:
            key = os.environ.get("DEEPSEEK_API_KEY", "")

        key_loaded = bool(key)
        logger.info(
            "DeepSeek key_loaded=%s key_prefix=%s model=%s",
            key_loaded,
            (key[:6] + "****") if len(key) >= 6 else "(none)",
            model_name,
        )

        if key:
            try:
                return DeepSeekTranslator(api_key=key, model_name=model_name)
            except Exception as exc:
                logger.error(
                    "DeepSeekTranslator creation failed, falling back to MockTranslator: %s", exc
                )
                return MockTranslator()

        logger.info("No DeepSeek API key found → MockTranslator")
        return MockTranslator()

    # Default: Gemini
    # Resolve model from settings if not explicitly given
    if not model_name:
        try:
            from app.settings.settings_manager import settings
            model_name = settings.get("selected_models.gemini", DEFAULT_GEMINI_MODEL)
        except Exception:
            model_name = DEFAULT_GEMINI_MODEL

    # Explicit key
    key = api_key or ""

    # Try secure storage (keyring) — key ONLY, not model
    if not key and use_secure_storage:
        try:
            from app.core.secure_storage import SecureStorage
            ss = SecureStorage()
            key = ss.load_api_key() or ""
        except Exception as exc:
            logger.warning("Failed to load key from secure storage: %s", exc)

    # Try env var
    if not key:
        key = os.environ.get("GEMINI_API_KEY", "")

    key_loaded = bool(key)
    logger.info(
        "key_loaded=%s key_prefix=%s model=%s",
        key_loaded,
        (key[:6] + "****") if len(key) >= 6 else "(none)",
        model_name,
    )

    if key:
        try:
            return GeminiTranslator(api_key=key, model_name=model_name)
        except Exception as exc:
            logger.error(
                "GeminiTranslator creation failed, falling back to MockTranslator: %s", exc
            )
            return MockTranslator()

    logger.info("No API key found → MockTranslator")
    return MockTranslator()

