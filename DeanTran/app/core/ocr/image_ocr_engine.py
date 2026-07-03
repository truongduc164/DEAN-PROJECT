import hashlib
import json
import logging
import uuid
from typing import List, Optional

from app.core.ocr.base_engine import BaseOcrEngine
from app.core.ocr.models import ImageOcrResult, OcrSegment, OcrRegion
from app.settings.settings_manager import settings

logger = logging.getLogger("DeanTran.ocr.engine")


class PaddleOcrEngine(BaseOcrEngine):
    def __init__(self, event_manager=None):
        super().__init__(event_manager)
        self._ocr = None

    def process_image_bytes(
        self, 
        image_bytes: bytes, 
        source_lang: str = "", 
        image_id: str = "", 
        regions: Optional[List[OcrRegion]] = None
    ) -> ImageOcrResult:
        image_id = image_id or str(uuid.uuid4())
        mode = settings.get("ocr_settings.image_ocr_mode", "off")
        if mode == "off":
            logger.info(f"[{image_id}] OCR mapping disabled. Skipping.")
            return ImageOcrResult(image_id=image_id, error="OCR disabled")
            
        try:
            import io
            from PIL import Image
            import numpy as np
            from paddleocr import PaddleOCR
            import logging
            
            # Subdue paddleocr logging
            logging.getLogger("ppocr").setLevel(logging.ERROR)
            
            image = Image.open(io.BytesIO(image_bytes))
            width, height = image.size
            if image.mode != "RGB":
                image = image.convert("RGB")
            
            # Map language to paddleocr lang
            lang_param = "en"
            if source_lang:
                sl = source_lang.lower()
                if "chinese" in sl:
                    lang_param = "ch"
                elif "vietnamese" in sl:
                    lang_param = "vi"
                    
            if self._ocr is None or getattr(self, '_lang_param', None) != lang_param:
                # PaddleOCR 3.x removed `use_angle_cls` and `show_log`.
                # `use_textline_orientation` serves a similar purpose.
                self._ocr = PaddleOCR(lang=lang_param, use_textline_orientation=True)
                self._lang_param = lang_param
            
            img_np = np.array(image)
            # PaddleOCR 3.x: ocr() is deprecated; use predict() instead.
            # predict() returns an iterator of result objects.
            result = list(self._ocr.predict(img_np))
            
            segments = []
            if result:
                for pred in result:
                    # PaddleOCR 3.x result objects have .text, .score, .bbox attributes
                    text = str(pred.text).strip() if hasattr(pred, 'text') else ''
                    conf = float(getattr(pred, 'score', 0)) * 100.0
                    
                    bbox = getattr(pred, 'bbox', None)
                    if bbox and len(bbox) >= 4:
                        xs = [p[0] for p in bbox]
                        ys = [p[1] for p in bbox]
                        left, top = min(xs), min(ys)
                        right, bottom = max(xs), max(ys)
                        w, h = right - left, bottom - top
                    else:
                        left, top, w, h = 0, 0, 0, 0
                    
                    segments.append(OcrSegment(
                        id=str(uuid.uuid4()),
                        text=text.strip(),
                        confidence=conf,
                        bbox=(int(left), int(top), int(w), int(h))
                    ))
            
            return ImageOcrResult(image_id=image_id, blocks=segments, width=width, height=height)
            
        except Exception as e:
            msg = f"PaddleOCR exception: {e}"
            if self.em: self.em.log("ERROR", f"[{image_id}] {msg}")
            return ImageOcrResult(image_id=image_id, error=msg)

class GoogleVisionOcrEngine(BaseOcrEngine):
    def __init__(self, event_manager=None):
        super().__init__(event_manager)
        self._client = None
        self._client_signature = ""

    def _build_client(self):
        from google.cloud import vision
        from google.oauth2 import service_account

        key = settings.get("ocr_settings.google_vision_key", "").strip()
        signature = hashlib.sha1(key.encode("utf-8")).hexdigest() if key else ""

        if self._client is not None and self._client_signature == signature:
            return self._client

        if key:
            if "{" in key:
                creds = service_account.Credentials.from_service_account_info(json.loads(key))
                self._client = vision.ImageAnnotatorClient(credentials=creds)
            else:
                self._client = vision.ImageAnnotatorClient(client_options={"api_key": key})
        else:
            self._client = vision.ImageAnnotatorClient()

        self._client_signature = signature
        return self._client

    def _extract_image_size(self, image_bytes: bytes) -> tuple[int, int]:
        import io
        from PIL import Image

        image = Image.open(io.BytesIO(image_bytes))
        return image.size

    def _parse_response(self, response, image_id: str, width: int, height: int) -> ImageOcrResult:
        if getattr(response, "error", None) and response.error.message:
            return ImageOcrResult(image_id=image_id, width=width, height=height, error=response.error.message)

        segments = []
        full_text = getattr(response, "full_text_annotation", None)
        pages = getattr(full_text, "pages", []) if full_text else []

        for page in pages:
            for block in page.blocks:
                for paragraph in block.paragraphs:
                    text_parts = []
                    for word in paragraph.words:
                        for symbol in word.symbols:
                            text_parts.append(symbol.text)
                        text_parts.append(" ")

                    text = "".join(text_parts).strip()
                    if not text:
                        continue

                    vertices = getattr(paragraph.bounding_box, "vertices", [])
                    xs = [v.x for v in vertices if v.x is not None]
                    ys = [v.y for v in vertices if v.y is not None]
                    if not xs or not ys:
                        continue

                    left, top = min(xs), min(ys)
                    right, bottom = max(xs), max(ys)
                    w, h = right - left, bottom - top

                    segments.append(OcrSegment(
                        id=str(uuid.uuid4()),
                        text=text,
                        confidence=paragraph.confidence * 100.0 if hasattr(paragraph, "confidence") else 100.0,
                        bbox=(int(left), int(top), int(w), int(h)),
                    ))

        return ImageOcrResult(image_id=image_id, blocks=segments, width=width, height=height)

    def _parse_rest_response(self, response_dict: dict, image_id: str, width: int, height: int) -> ImageOcrResult:
        err = response_dict.get("error")
        if err and err.get("message"):
            return ImageOcrResult(image_id=image_id, width=width, height=height, error=err.get("message"))

        segments = []
        full_text = response_dict.get("fullTextAnnotation")
        pages = full_text.get("pages", []) if full_text else []

        for page in pages:
            for block in page.get("blocks", []):
                for paragraph in block.get("paragraphs", []):
                    text_parts = []
                    for word in paragraph.get("words", []):
                        for symbol in word.get("symbols", []):
                            text_parts.append(symbol.get("text", ""))
                        text_parts.append(" ")

                    text = "".join(text_parts).strip()
                    if not text:
                        continue

                    vertices = paragraph.get("boundingBox", {}).get("vertices", [])
                    xs = [v.get("x") for v in vertices if v.get("x") is not None]
                    ys = [v.get("y") for v in vertices if v.get("y") is not None]
                    if not xs or not ys:
                        continue

                    left, top = min(xs), min(ys)
                    right, bottom = max(xs), max(ys)
                    w, h = right - left, bottom - top

                    segments.append(OcrSegment(
                        id=str(uuid.uuid4()),
                        text=text,
                        confidence=paragraph.get("confidence", 1.0) * 100.0,
                        bbox=(int(left), int(top), int(w), int(h)),
                    ))

        return ImageOcrResult(image_id=image_id, blocks=segments, width=width, height=height)

    def _process_image_bytes_rest(self, image_bytes: bytes, image_id: str) -> ImageOcrResult:
        import base64
        import requests
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

        key = settings.get("ocr_settings.google_vision_key", "").strip()
        if not key:
            return ImageOcrResult(image_id=image_id, error="Google Vision Key is empty")

        url = f"https://vision.googleapis.com/v1/images:annotate?key={key}"
        img_b64 = base64.b64encode(image_bytes).decode('utf-8')
        payload = {
            "requests": [
                {
                    "image": {"content": img_b64},
                    "features": [{"type": "DOCUMENT_TEXT_DETECTION"}]
                }
            ]
        }
        
        try:
            width, height = self._extract_image_size(image_bytes)
        except Exception:
            width, height = 0, 0

        try:
            res = requests.post(url, json=payload, verify=False, timeout=30)
            if res.status_code != 200:
                return ImageOcrResult(image_id=image_id, width=width, height=height, error=f"REST API error status: {res.status_code}")
            
            res_json = res.json()
            responses = res_json.get("responses", [])
            if not responses:
                return ImageOcrResult(image_id=image_id, width=width, height=height, error="Empty REST responses")
            
            return self._parse_rest_response(responses[0], image_id, width, height)
        except Exception as e:
            return ImageOcrResult(image_id=image_id, width=width, height=height, error=f"REST fallback failed: {e}")

    def _process_images_batch_rest(self, image_items: list[tuple[str, bytes]]) -> dict[str, ImageOcrResult]:
        import base64
        import requests
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

        results: dict[str, ImageOcrResult] = {}
        key = settings.get("ocr_settings.google_vision_key", "").strip()
        if not key:
            for image_id, _ in image_items:
                results[image_id] = ImageOcrResult(image_id=image_id, error="Google Vision Key is empty")
            return results

        url = f"https://vision.googleapis.com/v1/images:annotate?key={key}"
        
        # Build batch requests
        req_list = []
        meta_list = []
        for image_id, image_bytes in image_items:
            try:
                width, height = self._extract_image_size(image_bytes)
            except Exception:
                width, height = 0, 0
            
            img_b64 = base64.b64encode(image_bytes).decode('utf-8')
            req_list.append({
                "image": {"content": img_b64},
                "features": [{"type": "DOCUMENT_TEXT_DETECTION"}]
            })
            meta_list.append((image_id, width, height))

        payload = {"requests": req_list}

        try:
            res = requests.post(url, json=payload, verify=False, timeout=60)
            if res.status_code != 200:
                err_msg = f"REST API batch error status: {res.status_code}"
                for image_id, w, h in meta_list:
                    results[image_id] = ImageOcrResult(image_id=image_id, width=w, height=h, error=err_msg)
                return results

            res_json = res.json()
            responses = res_json.get("responses", [])
            for idx, (image_id, w, h) in enumerate(meta_list):
                if idx >= len(responses):
                    results[image_id] = ImageOcrResult(image_id=image_id, width=w, height=h, error="Missing batch REST response")
                    continue
                results[image_id] = self._parse_rest_response(responses[idx], image_id, w, h)
        except Exception as e:
            err_msg = f"REST batch fallback failed: {e}"
            for image_id, w, h in meta_list:
                results[image_id] = ImageOcrResult(image_id=image_id, width=w, height=h, error=err_msg)
        
        return results

    def process_image_bytes(
        self, 
        image_bytes: bytes, 
        source_lang: str = "", 
        image_id: str = "", 
        regions: Optional[List[OcrRegion]] = None
    ) -> ImageOcrResult:
        image_id = image_id or str(uuid.uuid4())
            
        try:
            from google.cloud import vision

            width, height = self._extract_image_size(image_bytes)
            client = self._build_client()
            image_vision = vision.Image(content=image_bytes)
            response = client.document_text_detection(image=image_vision)
            if getattr(response, "error", None) and response.error.message:
                raise Exception(response.error.message)
            return self._parse_response(response, image_id, width, height)
            
        except Exception as e:
            if self.em: self.em.log("WARN", f"[{image_id}] Google Vision gRPC failed, falling back to REST: {e}")
            return self._process_image_bytes_rest(image_bytes, image_id)

    def process_images_batch(
        self,
        image_items: list[tuple[str, bytes]],
        source_lang: str = "",
    ) -> dict[str, ImageOcrResult]:
        results: dict[str, ImageOcrResult] = {}
        if not image_items:
            return results

        try:
            from google.cloud import vision

            client = self._build_client()
            max_images = int(settings.get("ocr_settings.google_vision_max_images_per_request", 16))
            max_images = max(1, min(max_images, 16))

            prepared: list[tuple[str, bytes, int, int]] = []
            for image_id, image_bytes in image_items:
                try:
                    width, height = self._extract_image_size(image_bytes)
                except Exception:
                    width, height = 0, 0
                prepared.append((image_id, image_bytes, width, height))

            request_calls = 0
            for start in range(0, len(prepared), max_images):
                chunk = prepared[start:start + max_images]
                requests = []
                for _, image_bytes, _, _ in chunk:
                    requests.append(
                        vision.AnnotateImageRequest(
                            image=vision.Image(content=image_bytes),
                            features=[vision.Feature(type_=vision.Feature.Type.DOCUMENT_TEXT_DETECTION)],
                        )
                    )

                request_calls += 1
                batch_response = client.batch_annotate_images(requests=requests)
                responses = list(getattr(batch_response, "responses", []))

                # If any response has an error, trigger REST fallback
                for r in responses:
                    if getattr(r, "error", None) and r.error.message:
                        raise Exception(r.error.message)

                for idx, (image_id, _bytes, width, height) in enumerate(chunk):
                    if idx >= len(responses):
                        results[image_id] = ImageOcrResult(
                            image_id=image_id,
                            width=width,
                            height=height,
                            error="Missing OCR response in batch result",
                        )
                        continue
                    results[image_id] = self._parse_response(responses[idx], image_id, width, height)

            if self.em:
                self.em.log(
                    "INFO",
                    f"google_vision_batch_summary: images={len(image_items)} request_calls={request_calls}",
                )
            return results

        except Exception as exc:
            if self.em:
                self.em.log("WARN", f"Google Vision batch gRPC failed, falling back to REST: {exc}")
            return self._process_images_batch_rest(image_items)

    def process_images_canvas_batch(
        self,
        image_items: list[tuple[str, bytes]],
        source_lang: str = "",
    ) -> dict[str, ImageOcrResult]:
        """
        Merge many images into canvas images, OCR canvas(es), then map text back
        to original image coordinates.
        """
        results: dict[str, ImageOcrResult] = {}
        if not image_items:
            return results

        try:
            import io
            from PIL import Image
        except ImportError:
            msg = "Pillow not installed for canvas OCR."
            for image_id, _image_bytes in image_items:
                results[image_id] = ImageOcrResult(image_id=image_id, error=msg)
            return results

        images_per_canvas = int(settings.get("ocr_settings.google_vision_canvas_images_per_canvas", 4))
        images_per_canvas = max(1, min(images_per_canvas, 16))
        padding = int(settings.get("ocr_settings.google_vision_canvas_padding", 24))
        padding = max(0, padding)
        max_canvas_width = int(settings.get("ocr_settings.google_vision_canvas_max_width", 8192))
        max_canvas_height = int(settings.get("ocr_settings.google_vision_canvas_max_height", 8192))

        decoded = []
        raw_bytes_map = {}
        for image_id, image_bytes in image_items:
            raw_bytes_map[image_id] = image_bytes
            try:
                image = Image.open(io.BytesIO(image_bytes))
                if image.mode != "RGB":
                    image = image.convert("RGB")
                width, height = image.size
                decoded.append({
                    "image_id": image_id,
                    "image": image,
                    "width": width,
                    "height": height,
                })
            except Exception as exc:
                results[image_id] = ImageOcrResult(image_id=image_id, error=f"Invalid image for canvas OCR: {exc}")

        if not decoded:
            return results

        # Build vertical canvases with limit by count and max size.
        canvas_groups = []
        idx = 0
        while idx < len(decoded):
            group = []
            used_height = padding
            used_width = 0

            while idx < len(decoded) and len(group) < images_per_canvas:
                item = decoded[idx]
                est_width = max(used_width, item["width"]) + (padding * 2)
                est_height = used_height + item["height"] + padding

                if group and (est_width > max_canvas_width or est_height > max_canvas_height):
                    break

                group.append(item)
                used_height = est_height
                used_width = max(used_width, item["width"])
                idx += 1

            if not group:
                # Force include at least one image to avoid infinite loop.
                group = [decoded[idx]]
                idx += 1

            canvas_groups.append(group)

        canvas_requests = []
        canvas_meta = {}

        for canvas_idx, group in enumerate(canvas_groups, 1):
            canvas_id = f"canvas_{canvas_idx}"
            canvas_width = max(item["width"] for item in group) + (padding * 2)
            canvas_height = sum(item["height"] for item in group) + (padding * (len(group) + 1))

            canvas = Image.new("RGB", (canvas_width, canvas_height), color=(255, 255, 255))
            y_cursor = padding
            regions = []

            for item in group:
                x = padding
                y = y_cursor
                canvas.paste(item["image"], (x, y))
                regions.append({
                    "image_id": item["image_id"],
                    "x": x,
                    "y": y,
                    "w": item["width"],
                    "h": item["height"],
                })
                y_cursor += item["height"] + padding

                if item["image_id"] not in results:
                    results[item["image_id"]] = ImageOcrResult(
                        image_id=item["image_id"],
                        width=item["width"],
                        height=item["height"],
                    )

            buffer = io.BytesIO()
            canvas.save(buffer, format="PNG")
            canvas_requests.append((canvas_id, buffer.getvalue()))
            canvas_meta[canvas_id] = {"regions": regions}

        canvas_ocr_map = self.process_images_batch(canvas_requests, source_lang)

        for canvas_id, meta in canvas_meta.items():
            canvas_result = canvas_ocr_map.get(canvas_id)
            if canvas_result is None or canvas_result.error:
                if self.em:
                    self.em.log("WARN", f"{canvas_id}: canvas OCR failed, fallback to single-image OCR")
                for region in meta["regions"]:
                    image_id = region["image_id"]
                    if image_id in raw_bytes_map:
                        results[image_id] = self.process_image_bytes(raw_bytes_map[image_id], source_lang, image_id)
                continue

            for block in canvas_result.blocks:
                bx, by, bw, bh = block.bbox
                cx = bx + (bw / 2.0)
                cy = by + (bh / 2.0)

                owner_region = None
                for region in meta["regions"]:
                    if (region["x"] <= cx <= region["x"] + region["w"]) and (region["y"] <= cy <= region["y"] + region["h"]):
                        owner_region = region
                        break

                if not owner_region:
                    continue

                image_id = owner_region["image_id"]
                result = results.get(image_id)
                if result is None:
                    continue

                rel_x = max(0, int(bx - owner_region["x"]))
                rel_y = max(0, int(by - owner_region["y"]))
                rel_w = max(1, int(min(bw, owner_region["w"] - rel_x)))
                rel_h = max(1, int(min(bh, owner_region["h"] - rel_y)))

                result.blocks.append(OcrSegment(
                    id=str(uuid.uuid4()),
                    text=block.text,
                    confidence=block.confidence,
                    bbox=(rel_x, rel_y, rel_w, rel_h),
                ))

        if self.em:
            self.em.log(
                "INFO",
                f"google_vision_canvas_summary: images={len(decoded)} canvases={len(canvas_groups)}",
            )

        return results
