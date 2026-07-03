from dataclasses import dataclass, field
from typing import List, Tuple, Optional

@dataclass
class OcrRegion:
    """Represents a specific region to perform OCR on (e.g. drawn by user)."""
    id: str
    bbox: Tuple[int, int, int, int]  # (x, y, w, h)
    bbox_normalized: Tuple[float, float, float, float] = (0.0, 0.0, 0.0, 0.0) # 0 to 1
    region_type: str = "paragraph"   # "title", "paragraph", "note", "ignore"
    order_index: int = 0
    label: str = ""

@dataclass
class OcrSegment:
    """Represents a detected chunk of text (word, line, or paragraph) to be translated."""
    id: str
    text: str
    confidence: float
    bbox: Tuple[int, int, int, int]  # (x, y, w, h)
    translated_text: str = ""
    error: str = ""
    skipped: bool = False
    
    classification: str = "unknown"  # "meaningful", "noise", "skipped"
    skip_reason: str = ""
    overlay_created: bool = False
    
    # Internal usage for protection padding
    _modified_text: str = ""
    _tokens: list = field(default_factory=list)

# Alias for backward compatibility with older pipeline code
TextBlock = OcrSegment

@dataclass
class ImageOcrResult:
    image_id: str
    blocks: List[OcrSegment] = field(default_factory=list)
    width: int = 0
    height: int = 0
    error: str = ""
