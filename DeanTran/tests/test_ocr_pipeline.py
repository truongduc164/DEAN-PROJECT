"""Tests for the upgraded OCR pipeline, text styling, and Word processor."""
import pytest
from app.core.ocr.models import ImageOcrResult, OcrSegment, TextBlock
from app.core.ocr.image_translation_pipeline import ImageTranslationPipeline
from app.core.ocr.image_text_renderer import _hex_to_rgb, _compute_font_size


# ── Helpers ──────────────────────────────────────────────────────────

class DummyTranslator:
    def translate(self, text: str, source: str, target: str) -> str:
        if "__PROTECT_0__" in text:
            return f"Translated {text}"
        return f"Translated {text}"


class DummyEventManager:
    def __init__(self):
        self.logs = []
    def log(self, level, msg):
        self.logs.append((level, msg))
    def progress(self, done, total):
        pass


class DummyEngine:
    def process_image_bytes(self, image_bytes, source_lang="", image_id=""):
        return ImageOcrResult(
            image_id="img1",
            blocks=[
                OcrSegment("1", "2025-09-01", 99.0, (0, 0, 10, 10)),
                OcrSegment("2", "BTN-8", 99.0, (0, 0, 10, 10)),
                OcrSegment("3", "Part EPS-0929", 99.0, (0, 0, 10, 10)),
            ],
        )


# ── Tests ────────────────────────────────────────────────────────────

def test_ocr_skip_special_content(monkeypatch):
    import app.settings.settings_manager as sm
    sm.settings.set("ocr_settings.preserve_special_tokens_in_images", True)
    sm.settings.set("ocr_settings.translate_with_api", True)

    translator = DummyTranslator()
    em = DummyEventManager()
    pipeline = ImageTranslationPipeline(translator, em)
    pipeline.engine = DummyEngine()

    result = pipeline.process_image(b"dummy_bytes", "English", "Vietnamese")

    # 2025-09-01 is special content -> should be skipped
    assert result.blocks[0].skipped is True
    assert result.blocks[0].translated_text == "2025-09-01"

    # BTN-8 is token -> skipped
    assert result.blocks[1].skipped is True
    assert result.blocks[1].translated_text == "BTN-8"

    # "Part EPS-0929" mixed -> translated but token restored
    assert result.blocks[2].skipped is False
    assert "EPS-0929" in result.blocks[2].translated_text
    assert "Translated" in result.blocks[2].translated_text


def test_skip_noise_tokens():
    """Noise like |, 28, A9, REE should be skipped."""
    import app.settings.settings_manager as sm
    sm.settings.set("ocr_settings.translate_with_api", True)

    translator = DummyTranslator()
    em = DummyEventManager()
    pipeline = ImageTranslationPipeline(translator, em)

    class NoiseEngine:
        def process_image_bytes(self, image_bytes, source_lang="", image_id=""):
            return ImageOcrResult(
                image_id="noise_test",
                blocks=[
                    OcrSegment("n1", "|", 99.0, (0, 0, 10, 10)),
                    OcrSegment("n2", "28", 99.0, (0, 0, 10, 10)),
                    OcrSegment("n3", "REE", 99.0, (0, 0, 10, 10)),
                    OcrSegment("n4", "A9", 99.0, (0, 0, 10, 10)),
                    OcrSegment("n5", "BOE", 99.0, (0, 0, 10, 10)),
                ],
            )

    pipeline.engine = NoiseEngine()
    result = pipeline.process_image(b"dummy", "Chinese", "Vietnamese")

    for block in result.blocks:
        assert block.skipped is True
        assert block.translated_text == block.text


def test_translate_with_api_disabled():
    """When translate_with_api=False, text should pass through unchanged."""
    import app.settings.settings_manager as sm
    sm.settings.set("ocr_settings.translate_with_api", False)
    sm.settings.set("ocr_settings.preserve_special_tokens_in_images", False)

    translator = DummyTranslator()
    em = DummyEventManager()
    pipeline = ImageTranslationPipeline(translator, em)

    class SimpleEngine:
        def process_image_bytes(self, image_bytes, source_lang="", image_id=""):
            return ImageOcrResult(
                image_id="api_off",
                blocks=[
                    OcrSegment("s1", "Hello World Long Text Here", 99.0, (0, 0, 100, 20)),
                ],
            )

    pipeline.engine = SimpleEngine()
    result = pipeline.process_image(b"dummy", "Chinese", "Vietnamese")

    assert result.blocks[0].translated_text == "Hello World Long Text Here"
    # Reset
    sm.settings.set("ocr_settings.translate_with_api", True)


def test_none_translation_no_crash():
    """If translator returns None, pipeline should not crash."""
    import app.settings.settings_manager as sm
    sm.settings.set("ocr_settings.translate_with_api", True)
    sm.settings.set("ocr_settings.preserve_special_tokens_in_images", False)

    class NoneTranslator:
        def translate(self, text, source, target):
            return None

    em = DummyEventManager()
    pipeline = ImageTranslationPipeline(NoneTranslator(), em)

    class SimpleEngine:
        def process_image_bytes(self, image_bytes, source_lang="", image_id=""):
            return ImageOcrResult(
                image_id="none_test",
                blocks=[
                    OcrSegment("x1", "Some Chinese text here", 99.0, (0, 0, 100, 20)),
                ],
            )

    pipeline.engine = SimpleEngine()
    # Should not raise
    result = pipeline.process_image(b"dummy", "Chinese", "Vietnamese")
    # None was normalized to "" -> block should get original text back (null-safe)
    assert result.blocks[0].translated_text == "Some Chinese text here"


def test_suspicious_output_rejected():
    """If translator returns THINK: output, it should be rejected."""
    import app.settings.settings_manager as sm
    sm.settings.set("ocr_settings.translate_with_api", True)
    sm.settings.set("ocr_settings.preserve_special_tokens_in_images", False)

    class ThinkTranslator:
        def translate(self, text, source, target):
            return "THINK: The user wants to translate..."

    em = DummyEventManager()
    pipeline = ImageTranslationPipeline(ThinkTranslator(), em)

    class SimpleEngine:
        def process_image_bytes(self, image_bytes, source_lang="", image_id=""):
            return ImageOcrResult(
                image_id="think_test",
                blocks=[
                    OcrSegment("t1", "Hello", 99.0, (0, 0, 100, 20)),
                ],
            )

    pipeline.engine = SimpleEngine()
    result = pipeline.process_image(b"dummy", "Chinese", "Vietnamese")

    # The THINK output should be rejected; original text kept
    assert result.blocks[0].translated_text == "Hello"
    assert any("suspicious" in msg.lower() for _, msg in em.logs)


# ── Text Styling Tests ───────────────────────────────────────────────

def test_hex_to_rgb():
    assert _hex_to_rgb("#FF0000") == (255, 0, 0)
    assert _hex_to_rgb("#00FF00") == (0, 255, 0)
    assert _hex_to_rgb("#333333") == (51, 51, 51)
    assert _hex_to_rgb("#1A1A2E") == (26, 26, 46)
    assert _hex_to_rgb("invalid") == (0, 0, 0)


def test_compute_font_size_fixed():
    size = _compute_font_size("Hello", 200, 50, 14, 10, 48, "fixed", 100)
    assert size == 14


def test_compute_font_size_auto_fit():
    size = _compute_font_size("Short", 200, 50, 24, 10, 48, "auto_fit", 100)
    assert 10 <= size <= 48


def test_compute_font_size_min_respected():
    size = _compute_font_size("Very long text " * 50, 50, 10, 24, 12, 48, "auto_fit", 100)
    assert size >= 12


def test_compute_font_size_scale():
    size = _compute_font_size("Hi", 200, 50, 20, 10, 48, "scale_from_original", 150)
    assert size == 30  # 20 * 150%


# ── Backward Compatibility ───────────────────────────────────────────

def test_textblock_alias():
    """TextBlock alias should still work."""
    block = TextBlock(id="1", text="test", confidence=99.0, bbox=(0, 0, 10, 10))
    assert block.text == "test"
    assert isinstance(block, OcrSegment)
