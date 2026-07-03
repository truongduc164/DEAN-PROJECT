"""
ImageTranslationPipeline – Coordinates OCR detection, filtering, batch
translation via Gemini, and hallucination rejection.

Key design: OCR all images first, collect translatable segments,
send ONE batch Gemini request (JSON), parse, map back, render.
"""
import json
import copy
import hashlib
import logging
import re
import time
import uuid

from app.core.ocr.models import ImageOcrResult, OcrSegment
from app.core.ocr.image_ocr_engine import PaddleOcrEngine, GoogleVisionOcrEngine
from app.core.translators.protection import TextProtector
from app.settings.settings_manager import settings

logger = logging.getLogger("DeanTran.ocr.pipeline")


def normalize_text(value) -> str:
    """Null-safe text normalizer. None -> '', always str."""
    if value is None:
        return ""
    return str(value).strip()


def _is_suspicious(text: str) -> bool:
    """Detect hallucinated or reasoning output from Gemini."""
    if not text:
        return False
    lower = text.lower()
    if "think:" in lower or "the user wants" in lower:
        return True
    if "analyze" in lower and "conclusion" in lower:
        return True
    if text.startswith("**") and text.endswith("**"):
        return True
    # Reject if output has markdown bullets or numbered reasoning
    if re.search(r"^\d+\.\s", text, re.MULTILINE) and len(text) > 200:
        return True
    return False


class ImageTranslationPipeline:
    """
    Coordinates OCR detection, Protection padding, Translation, and records results safely.
    Supports both single-image and batch-image translation modes.
    """
    def __init__(self, translator_service, event_manager):
        self.em = event_manager
        
        engine_choice = settings.get("ocr_settings.engine", "paddle")
        if engine_choice == "google_vision":
            self.engine = GoogleVisionOcrEngine(event_manager=self.em)
        else:
            self.engine = PaddleOcrEngine(event_manager=self.em)
            
        self.protector = TextProtector()
        self.translator = translator_service

    def _classify_block(self, block: OcrSegment) -> tuple[bool, str]:
        """Determine if block is meaningful enough to translate."""
        text = str(block.text).strip()
        if not text:
            return False, "empty"

        conf_threshold = float(settings.get("ocr_settings.min_confidence", 0.0))
        if block.confidence < conf_threshold:
            return False, f"low_confidence_({block.confidence:.1f}<{conf_threshold})"

        # Too small bbox?
        _, _, w, h = block.bbox
        if w < 5 or h < 5:
            return False, "bbox_too_small"

        has_chinese = bool(re.search(r'[\u4e00-\u9fff]', text))

        if len(text) == 1 and not has_chinese:
            if text.isalpha():
                return False, "single_letter"
            if text.isdigit():
                return False, "single_digit"
            return False, "single_symbol"

        if text.isdigit() and len(text) <= 4:
            return False, "short_numeric"

        if re.fullmatch(r"[\W_\s]+", text):
            return False, "punctuation_only"

        # mostly uppercase Latin letters with length <= 3
        if len(text) <= 3 and text.isupper() and re.fullmatch(r"[A-Z0-9\W]+", text) and not has_chinese:
            return False, "short_acronym_or_noise"

        # mostly numeric/symbolic and has no Chinese characters
        if not has_chinese and len(re.sub(r'[\d\W_]', '', text)) < 2:
            return False, "mostly_numeric_or_symbolic"

        return True, "meaningful"

    # ── Single image processing (legacy / fallback) ──────────────────

    def process_image(self, image_bytes: bytes, source_lang: str, target_lang: str, image_id: str = "") -> ImageOcrResult:
        """Process a single image: OCR -> filter -> translate -> return."""
        result = self.engine.process_image_bytes(image_bytes, source_lang, image_id)
        if result.error or not result.blocks:
            return result

        preserve_special = settings.get("ocr_settings.preserve_special_tokens_in_images", True)
        translate_with_api = settings.get("ocr_settings.translate_with_api", True)

        blocks_to_translate = self._filter_blocks(result, preserve_special)

        if not translate_with_api:
            self.em.log("INFO", f"[{result.image_id}] translate_with_api=False, bypassing API...")
            for block in blocks_to_translate:
                block.translated_text = block.text
        elif blocks_to_translate:
            # Single-image mode: use individual translate calls as fallback
            self._translate_blocks_individually(blocks_to_translate, result, source_lang, target_lang, preserve_special)

        self.em.log("INFO", f"[{result.image_id}] image_translation_completed ({len(result.blocks)} blocks)")
        return result

    # ── Batch OCR: Step 1 – OCR all images, filter, return segments ──

    def ocr_and_filter_images(self, images: list, source_lang: str) -> list:
        """
        OCR all images and return list of (image_meta, ocr_result, blocks_to_translate).
        images: list of (slide, shape, image_bytes, slide_idx)
        Returns: list of dict with 'slide', 'shape', 'result', 'translatable_blocks'
        """
        preserve_special = settings.get("ocr_settings.preserve_special_tokens_in_images", True)
        dedupe_enabled = settings.get("ocr_settings.google_vision_dedupe_images", True)
        vision_batch_enabled = settings.get("ocr_settings.google_vision_batch_enabled", True)
        vision_canvas_enabled = settings.get("ocr_settings.google_vision_canvas_enabled", False)

        image_entries = []
        for slide, shape, image_bytes, slide_idx in images:
            # Word documents pass shape=None (they carry run_elem/para instead).
            # Fall back to a simple index-based id when shape is missing.
            if shape is not None and hasattr(shape, 'shape_id'):
                image_id = f"s{slide_idx}_{shape.shape_id}"
            else:
                image_id = f"img_{slide_idx}_{uuid.uuid4().hex[:8]}"

            image_key = hashlib.sha1(image_bytes).hexdigest() if dedupe_enabled else image_id
            image_entries.append({
                "slide": slide,
                "shape": shape,
                "image_bytes": image_bytes,
                "image_id": image_id,
                "image_key": image_key,
            })

        unique_entries = []
        key_to_entry = {}
        for entry in image_entries:
            if entry["image_key"] in key_to_entry:
                continue
            key_to_entry[entry["image_key"]] = entry
            unique_entries.append(entry)

        unique_count = len(unique_entries)
        deduped_count = max(0, len(image_entries) - unique_count)
        supports_vision_batch = (
            isinstance(self.engine, GoogleVisionOcrEngine)
            and hasattr(self.engine, "process_images_batch")
        )
        supports_vision_canvas = (
            isinstance(self.engine, GoogleVisionOcrEngine)
            and hasattr(self.engine, "process_images_canvas_batch")
        )
        use_vision_canvas = supports_vision_canvas and vision_canvas_enabled and unique_count > 0
        use_vision_batch = supports_vision_batch and vision_batch_enabled and unique_count > 0 and not use_vision_canvas

        if use_vision_canvas:
            self.em.log(
                "INFO",
                f"ocr_vision_mode=canvas unique_images={unique_count} deduped={deduped_count}",
            )
        elif use_vision_batch:
            self.em.log(
                "INFO",
                f"ocr_vision_mode=batch unique_images={unique_count} deduped={deduped_count}",
            )
        else:
            self.em.log(
                "INFO",
                f"ocr_vision_mode=single unique_images={unique_count} deduped={deduped_count}",
            )

        cache_by_key = {}
        if use_vision_canvas:
            batch_items = [(entry["image_id"], entry["image_bytes"]) for entry in unique_entries]
            batch_map = self.engine.process_images_canvas_batch(batch_items, source_lang)
            for entry in unique_entries:
                result = batch_map.get(entry["image_id"])
                if result is None:
                    result = self.engine.process_image_bytes(
                        entry["image_bytes"], source_lang, entry["image_id"]
                    )
                cache_by_key[entry["image_key"]] = result
        elif use_vision_batch:
            batch_items = [(entry["image_id"], entry["image_bytes"]) for entry in unique_entries]
            batch_map = self.engine.process_images_batch(batch_items, source_lang)
            for entry in unique_entries:
                result = batch_map.get(entry["image_id"])
                if result is None:
                    result = self.engine.process_image_bytes(
                        entry["image_bytes"], source_lang, entry["image_id"]
                    )
                cache_by_key[entry["image_key"]] = result
        else:
            for entry in unique_entries:
                cache_by_key[entry["image_key"]] = self.engine.process_image_bytes(
                    entry["image_bytes"], source_lang, entry["image_id"]
                )

        all_results = []
        total_segments = 0
        total_skipped = 0

        for entry in image_entries:
            result_template = cache_by_key.get(entry["image_key"])
            if result_template is None:
                result_template = self.engine.process_image_bytes(
                    entry["image_bytes"], source_lang, entry["image_id"]
                )

            result = copy.deepcopy(result_template)
            result.image_id = entry["image_id"]
            for idx, block in enumerate(result.blocks):
                block.id = f"{entry['image_id']}_b{idx}_{uuid.uuid4().hex[:8]}"

            if result.error or not result.blocks:
                all_results.append({
                    'slide': entry["slide"], 'shape': entry["shape"], 'result': result,
                    'translatable_blocks': [],
                    'image_bytes': entry.get("image_bytes"),
                })
                continue

            blocks_to_translate = self._filter_blocks(result, preserve_special)
            skipped = len(result.blocks) - len(blocks_to_translate)
            total_segments += len(blocks_to_translate)
            total_skipped += skipped

            all_results.append({
                'slide': entry["slide"], 'shape': entry["shape"], 'result': result,
                'translatable_blocks': blocks_to_translate,
                'image_bytes': entry.get("image_bytes"),
            })

        self.em.log("INFO",
            f"ocr_batch_summary: images={len(images)} "
            f"total_segments={total_segments} skipped={total_skipped}"
        )
        return all_results

    # ── Batch OCR: Step 2 – Batch translate via Gemini JSON ──────────

    def batch_translate_segments(self, ocr_results: list, source_lang: str, target_lang: str):
        """
        Collect all translatable segments from all images, send as ONE batch
        Gemini request with JSON format, parse, and map translations back.
        """
        translate_with_api = True # Forced
        preserve_special = settings.get("ocr_settings.preserve_special_tokens_in_images", True)

        if not translate_with_api:
            self.em.log("INFO", "translate_with_api=False, bypassing API for all OCR segments...")
            for item in ocr_results:
                for block in item['translatable_blocks']:
                    block.translated_text = block.text
            return

        # Collect all segments with unique IDs
        segment_map = {}  # segment_id -> block
        batch_items = []   # list of {"id": ..., "text": ...}

        for item in ocr_results:
            for block in item['translatable_blocks']:
                seg_id = block.id
                segment_map[seg_id] = block
                batch_items.append({
                    "id": seg_id,
                    "text": block._modified_text or block.text,
                })

        if not batch_items:
            self.em.log("INFO", "No translatable OCR segments to send to API.")
            return

        total_chars = sum(len(b["text"]) for b in batch_items)
        self.em.log("INFO",
            f"ocr_batch_translate: segments={len(batch_items)} "
            f"total_chars={total_chars}"
        )

        # Send batch request(s)
        max_per_batch = settings.get("limits_settings.max_items_per_batch", 200)
        max_chars = settings.get("limits_settings.max_chars_per_batch", 30000)

        # Split into sub-batches if needed
        sub_batches = self._split_into_batches(batch_items, max_per_batch, max_chars)

        self.em.log("INFO", f"ocr_batch_count={len(sub_batches)}")

        api_calls = 0
        translated_count = 0
        failed_count = 0

        for batch_idx, batch in enumerate(sub_batches):
            batch_chars = sum(len(b["text"]) for b in batch)
            self.em.log("INFO",
                f"ocr_batch[{batch_idx+1}/{len(sub_batches)}]: "
                f"items={len(batch)} chars={batch_chars}"
            )

            # Try vision-context translation first (per image)
            translations = None
            if settings.get("ocr_settings.vision_context_enabled", False):
                # Find the image_bytes for this batch's segments
                first_block = segment_map.get(batch[0]["id"]) if batch else None
                if first_block:
                    for item in ocr_results:
                        for blk in item.get('translatable_blocks', []):
                            if blk.id == first_block.id:
                                img_bytes = item.get('image_bytes')
                                if img_bytes:
                                    image_blocks = item['translatable_blocks']
                                    translations = self._call_vision_context_translate(
                                        img_bytes, image_blocks, source_lang, target_lang
                                    )
                                break
                        if translations:
                            break

            if not translations:
                translations = self._call_gemini_batch(batch, source_lang, target_lang)
            api_calls += 1

            if translations is None:
                # Retry with smaller batches
                self.em.log("WARN", f"Batch {batch_idx+1} failed. Trying split...")
                if len(batch) > 1:
                    half = len(batch) // 2
                    t1 = self._call_gemini_batch(batch[:half], source_lang, target_lang)
                    api_calls += 1
                    t2 = self._call_gemini_batch(batch[half:], source_lang, target_lang)
                    api_calls += 1
                    translations = {}
                    if t1:
                        translations.update(t1)
                    if t2:
                        translations.update(t2)

            if not translations:
                translations = {}

            # Map translations back to segments
            for seg_item in batch:
                seg_id = seg_item["id"]
                block = segment_map.get(seg_id)
                if not block:
                    continue

                raw_translation = normalize_text(translations.get(seg_id, ""))

                if not raw_translation:
                    block.translated_text = block.text
                    failed_count += 1
                    self.em.log("WARN", f"[{seg_id}] Empty translation, keeping original: '{block.text}'")
                    continue

                # Restore protected tokens
                if preserve_special and block._tokens:
                    raw_translation = self.protector.restore(raw_translation, block._tokens)

                # Hallucination check
                if _is_suspicious(raw_translation):
                    self.em.log("WARN",
                        f"[{seg_id}] Suspicious output rejected. Keeping original: '{block.text}'"
                    )
                    block.translated_text = block.text
                    failed_count += 1
                    continue

                # Length sanity: reject if output is 5x longer than input
                if len(raw_translation) > len(block.text) * 5 and len(block.text) < 20:
                    self.em.log("WARN",
                        f"[{seg_id}] Output too long ({len(raw_translation)} vs {len(block.text)}). Keeping original."
                    )
                    block.translated_text = block.text
                    failed_count += 1
                    continue

                block.translated_text = raw_translation
                translated_count += 1
                self.em.log("INFO", f"[{seg_id}] '{block.text}' -> '{raw_translation}'")

        self.em.log("INFO",
            f"ocr_batch_result: api_calls={api_calls} "
            f"translated={translated_count} failed={failed_count}"
        )

    def _call_vision_context_translate(
        self, image_bytes: bytes, blocks: list, source_lang: str, target_lang: str
    ) -> dict | None:
        """Send OCR text fragments to AI. AI groups them into coherent sentences
        based on meaning, translates, and maps back to block IDs.

        No coordinate analysis - AI figures out grouping from content alone.
        """
        import requests

        use_vision = settings.get("ocr_settings.vision_context_enabled", False)
        if not use_vision:
            return None

        # Collect all text fragments
        texts = []
        for b in blocks:
            texts.append("[id={}] {}".format(b.id, b._modified_text or b.text))
        all_text = "\n".join(texts)

        prompt = (
            "You received OCR text fragments from an image. Each line has format:\n"
            "[id=xxx] text content\n\n"
            "=== TEXT FRAGMENTS ===\n"
            "{}\n\n"
            "=== YOUR TASK ===\n"
            "1. Read ALL fragments and understand what the image is about.\n"
            "2. GROUP fragments that belong together as coherent sentences/meanings.\n"
            "   - Use content/meaning to decide grouping, NOT positions.\n"
            "   - Numbers and their labels go together (e.g. 'width' + '120mm')\n"
            "   - Multiple lines of same paragraph go together\n"
            "   - Standalone items stay separate\n"
            "3. FORM each group into a natural {} sentence.\n"
            "4. TRANSLATE to {}.\n"
            "5. MAP each translation to ALL block IDs in that group.\n\n"
            "=== OUTPUT ===\n"
            'JSON only: {{"block_id_1": "translation_1", "block_id_2": "translation_2", ...}}\n'
            "- Every block ID must appear exactly once\n"
            "- No markdown, no explanation"
        ).format(all_text, source_lang, target_lang)

        retry_limit = settings.get("limits_settings.retry_limit", 3)

        for attempt in range(retry_limit):
            try:
                tool = settings.get("translation_tool", "Gemini")
                if tool == "DeepSeek":
                    import os
                    from app.core.secure_storage import SecureStorage
                    key = SecureStorage().load_deepseek_key()
                    if not key:
                        key = os.environ.get("DEEPSEEK_API_KEY", "")
                    if not key:
                        self.em.log("ERROR", "No DeepSeek key")
                        return None

                    model = settings.get("selected_models.deepseek", "deepseek-chat")
                    headers = {
                        "Content-Type": "application/json",
                        "Authorization": "Bearer {}".format(key)
                    }
                    payload = {
                        "model": model,
                        "messages": [{"role": "user", "content": prompt}],
                        "temperature": 0.2,
                        "max_tokens": 4096
                    }
                    resp = requests.post(
                        "https://api.deepseek.com/chat/completions",
                        headers=headers, json=payload, timeout=90
                    )
                    if resp.status_code == 429:
                        delay = 2.0 ** (attempt + 2)
                        self.em.log("WARN", "Rate limited, waiting {:.1f}s".format(delay))
                        time.sleep(delay)
                        continue
                    resp.raise_for_status()
                    raw = normalize_text(resp.json()["choices"][0]["message"]["content"])
                    result = self._parse_vision_response(raw, blocks)
                    if result:
                        self.em.log("INFO", "DeepSeek grouped {} blocks into {} translations".format(
                            len(blocks), len(set(result.values()))))
                    return result
                else:
                    from app.core.key_provider import KeyProvider
                    from google import genai
                    kp = KeyProvider()
                    key = kp.get_key()
                    if not key:
                        return None

                    client = genai.Client(api_key=key)
                    model_name = settings.get("selected_models.gemini", "gemini-2.5-flash")

                    import PIL.Image, io
                    pil_img = PIL.Image.open(io.BytesIO(image_bytes))
                    response = client.models.generate_content(
                        model=model_name,
                        contents=[prompt, pil_img],
                    )
                    raw = normalize_text(getattr(response, 'text', ''))
                    return self._parse_vision_response(raw, blocks)

            except Exception as e:
                self.em.log("WARN", "Vision attempt {} failed: {}".format(attempt+1, e))
                if attempt < retry_limit - 1:
                    time.sleep(2 ** attempt)
                continue

        return None


    def _parse_vision_response(self, raw: str, blocks: list) -> dict:
        """Parse vision AI JSON response into block_id -> translation mapping."""
        import json as _json
        try:
            # Strip markdown code fences
            cleaned = raw.strip()
            if cleaned.startswith("```"):
                lines = cleaned.split("\n")
                lines = [l for l in lines if not l.startswith("```")]
                cleaned = "\n".join(lines)
            # Find JSON object
            start_pos = cleaned.find("{")
            end_pos = cleaned.rfind("}")
            if start_pos >= 0 and end_pos > start_pos:
                cleaned = cleaned[start_pos:end_pos+1]
            result = _json.loads(cleaned)
            if isinstance(result, dict):
                return {str(k): str(v) for k, v in result.items()}
        except (_json.JSONDecodeError, ValueError) as e:
            self.em.log("WARN", "Failed to parse vision response: {}\nRaw: {}".format(e, raw[:200]))
        return {}



    def _filter_blocks(self, result, preserve_special: bool) -> list:
        """Filter OCR blocks: skip noise, protect tokens, return translatable list."""
        blocks_to_translate = []
        skip_noise = settings.get("ocr_settings.skip_noise_blocks", True)

        for block in result.blocks:
            is_meaningful, reason = self._classify_block(block)

            if not is_meaningful and skip_noise:
                block.classification = "noise"
                block.skip_reason = reason
                block.skipped = True
                block.translated_text = block.text
                self.em.log("INFO", "OCR block skipped reason={} text=\"{}\"".format(reason, block.text))
                continue

            block.classification = "meaningful"

            if preserve_special:
                modified_text, tokens, fully_protected = self.protector.protect(block.text)
                if fully_protected:
                    block.translated_text = self.protector.restore(modified_text, tokens)
                    block.skipped = True
                    continue

                block._modified_text = modified_text
                block._tokens = tokens
            else:
                block._modified_text = block.text
                block._tokens = []

            blocks_to_translate.append(block)

        return blocks_to_translate

    def _translate_blocks_individually(self, blocks, result, source_lang, target_lang, preserve_special):
        """Fallback: translate blocks one-by-one (non-batch)."""
        try:
            for block in blocks:
                translated = normalize_text(
                    self.translator.translate(block._modified_text, source_lang, target_lang)
                )
                if not translated:
                    block.translated_text = block.text
                    self.em.log("WARN", "[{}] Empty API response, keeping original: '{}'".format(result.image_id, block.text))
                    continue
                if preserve_special and block._tokens:
                    translated = self.protector.restore(translated, block._tokens)
                if _is_suspicious(translated):
                    self.em.log("WARN", "[{}] Suspicious output rejected.".format(result.image_id))
                    translated = block.text
                block.translated_text = translated
                self.em.log("INFO", "[{}] '{}' -> '{}'".format(result.image_id, block.text, translated))
                time.sleep(0.1)
        except Exception as e:
            self.em.log("ERROR", "[{}] Translation failed: {}".format(result.image_id, e))
            for block in blocks:
                if not block.translated_text:
                    block.translated_text = block.text
                    block.error = str(e)

    def _call_gemini_batch(self, batch_items: list, source_lang: str, target_lang: str) -> dict | None:
        """Send a batch of OCR segments to Gemini for translation.

        Returns dict of {segment_id: translated_text} or None on failure.
        """
        from app.core.key_provider import KeyProvider

        input_json = json.dumps(
            [{"id": b["id"], "text": b["text"]} for b in batch_items],
            ensure_ascii=False
        )

        prompt = (
            "Translate the following text segments from {} to {}.\n".format(source_lang, target_lang) +
            "Return ONLY a valid JSON array. Each element must have 'id' and 'translation' fields.\n"
            "Rules:\n"
            "- Return ONLY the translated text in 'translation' field\n"
            "- Do NOT explain, reason, or add notes\n"
            "- Do NOT use markdown\n"
            "- If a segment is a code, acronym, number, or technical term, keep it as-is\n"
            "- If a segment is unrecognizable OCR noise, keep it as-is\n"
            "- Do NOT expand short text into long sentences\n\n"
            "Input:\n{}\n\n".format(input_json) +
            "Output (JSON array only):"
        )

        retry_limit = settings.get("limits_settings.retry_limit", 3)
        kp = KeyProvider()

        tool = settings.get("translation_tool", "Gemini")

        for attempt in range(retry_limit):
            try:
                if tool == "DeepSeek":
                    import os
                    from app.core.secure_storage import SecureStorage
                    key = SecureStorage().load_deepseek_key()
                    if not key:
                        key = os.environ.get("DEEPSEEK_API_KEY", "")
                    if not key:
                        self.em.log("ERROR", "No DeepSeek API key available for OCR batch translation")
                        return None

                    model = settings.get("selected_models.deepseek", "deepseek-chat")
                    import requests
                    headers = {
                        "Content-Type": "application/json",
                        "Authorization": "Bearer {}".format(key)
                    }
                    payload = {
                        "model": model,
                        "messages": [
                            {"role": "user", "content": prompt}
                        ],
                        "temperature": 0.2
                    }
                    t0 = time.time()
                    resp = requests.post(
                        "https://api.deepseek.com/chat/completions",
                        headers=headers,
                        json=payload,
                        timeout=60
                    )
                    elapsed = time.time() - t0
                    if resp.status_code == 429:
                        delay = 2.0 ** (attempt + 2)
                        self.em.log("WARN", "DeepSeek Rate Limit 429 on OCR (attempt {}, wait {:.1f}s)".format(attempt+1, delay))
                        time.sleep(delay)
                        continue
                    resp.raise_for_status()
                    res_data = resp.json()
                    raw = normalize_text(res_data["choices"][0]["message"]["content"])
                else:
                    key = kp.get_key()
                    if not key:
                        self.em.log("ERROR", "No API key available for OCR batch translation")
                        return None

                    from google import genai
                    client = genai.Client(api_key=key)
                    model = settings.get("selected_models.gemini", "gemini-3.1-flash-lite")

                    response = client.models.generate_content(
                        model=model,
                        contents=prompt,
                    )
                    raw = normalize_text(getattr(response, 'text', ''))

                if not raw:
                    self.em.log("WARN", "Empty response (attempt {})".format(attempt+1))
                    time.sleep(2 ** attempt)
                    continue

                # Strip think tags if present
                if "<think>" in raw:
                    end_think = raw.find("</think>")
                    if end_think != -1:
                        raw = raw[end_think + len("</think>"):].strip()

                # Find JSON array or object
                start = raw.find("[")
                end = raw.rfind("]")
                if start >= 0 and end > start:
                    raw = raw[start:end+1]
                else:
                    # Try object format
                    start = raw.find("{")
                    end = raw.rfind("}")
                    if start >= 0 and end > start:
                        raw = raw[start:end+1]

                translations_list = json.loads(raw)

                # Map back: id -> translation
                if isinstance(translations_list, list):
                    return {
                        str(item["id"]): str(item["translation"])
                        for item in translations_list
                        if "id" in item and "translation" in item
                    }
                elif isinstance(translations_list, dict):
                    # Handle object format: {"id1": "trans1", ...}
                    return {str(k): str(v) for k, v in translations_list.items()}

                self.em.log("WARN", "Unexpected JSON format: list={}, dict={}".format(
                    isinstance(translations_list, list), isinstance(translations_list, dict)))

            except Exception as e:
                self.em.log("WARN", "Batch attempt {} failed: {}".format(attempt+1, e))
                if attempt < retry_limit - 1:
                    time.sleep(2 ** attempt)
                continue

        return None


    def _split_into_batches(self, items: list, max_items: int, max_chars: int) -> list:
        """Split items into sub-batches respecting item and char limits."""
        batches = []
        current = []
        current_chars = 0

        for item in items:
            item_chars = len(item["text"])
            if current and (len(current) >= max_items or current_chars + item_chars > max_chars):
                batches.append(current)
                current = []
                current_chars = 0
            current.append(item)
            current_chars += item_chars

        if current:
            batches.append(current)
        return batches