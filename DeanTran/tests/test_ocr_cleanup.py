import pytest
from app.core.ocr.models import OcrSegment, ImageOcrResult
from app.core.ocr.image_translation_pipeline import ImageTranslationPipeline
from app.settings.settings_manager import settings

@pytest.fixture(autouse=True)
def setup_settings():
    settings.set("ocr_settings.min_confidence", 40.0)
    settings.set("ocr_settings.skip_noise_blocks", True)
    yield

class DummyProtector:
    def protect(self, text): return text, [], False
    def restore(self, text, tokens): return text

class MockEventMgr:
    def log(self, lvl, msg): pass

def test_classify_meaningful_chinese():
    pipeline = ImageTranslationPipeline(None, MockEventMgr())
    
    # 1. meaningful Chinese
    block = OcrSegment(id="1", text="不合格", confidence=90.0, bbox=(0,0,100,20))
    is_meaningful, reason = pipeline._classify_block(block)
    assert is_meaningful is True

def test_classify_noise_digits():
    pipeline = ImageTranslationPipeline(None, MockEventMgr())
    
    # single digit
    block1 = OcrSegment(id="1", text="9", confidence=90.0, bbox=(0,0,100,20))
    is_meaningful, reason = pipeline._classify_block(block1)
    assert is_meaningful is False
    assert reason == "single_digit"
    
    # short numeric string
    block2 = OcrSegment(id="2", text="28", confidence=90.0, bbox=(0,0,100,20))
    is_meaningful, reason = pipeline._classify_block(block2)
    assert is_meaningful is False
    assert reason == "short_numeric"

def test_classify_single_latin_letter():
    pipeline = ImageTranslationPipeline(None, MockEventMgr())
    block = OcrSegment(id="1", text="W", confidence=90.0, bbox=(0,0,100,20))
    is_meaningful, reason = pipeline._classify_block(block)
    assert is_meaningful is False
    assert reason == "single_letter"

def test_filter_blocks_applies_classification():
    pipeline = ImageTranslationPipeline(None, MockEventMgr())
    result = ImageOcrResult(image_id="img1")
    result.blocks = [
        OcrSegment(id="1", text="W", confidence=90.0, bbox=(0,0,100,20)),
        OcrSegment(id="2", text="This is valid text", confidence=90.0, bbox=(0,0,100,20))
    ]
    
    filtered = pipeline._filter_blocks(result, preserve_special=False)
    
    # Only 1 block should be translated
    assert len(filtered) == 1
    assert filtered[0].text == "This is valid text"
    
    # The first block should be classified as noise and skipped
    assert result.blocks[0].classification == "noise"
    assert result.blocks[0].skipped is True
