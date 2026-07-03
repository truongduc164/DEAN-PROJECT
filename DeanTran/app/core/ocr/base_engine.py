from abc import ABC, abstractmethod
from typing import List, Optional
import logging

from app.core.ocr.models import ImageOcrResult, OcrRegion

class BaseOcrEngine(ABC):
    """
    Abstract Base Class for OCR Engines.
    Future-proofs the pipeline so we can swap Tesseract with Vision APIs seamlessly.
    """
    def __init__(self, event_manager=None):
        self.em = event_manager

    @abstractmethod
    def process_image_bytes(self, image_bytes: bytes, source_lang: str = "", image_id: str = "", regions: Optional[List[OcrRegion]] = None) -> ImageOcrResult:
        """
        Process the image and return extracted OcrSegments.
        If `regions` are provided, the engine should restrict OCR to those areas.
        """
        pass
