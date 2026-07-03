"""
ImageTextRenderer – Renders translated text blocks back as native PPT textboxes
with configurable font, size, color, fit strategy, and background.
"""
import logging
from typing import Optional
from app.core.ocr.models import ImageOcrResult
from app.settings.settings_manager import settings

logger = logging.getLogger("DeanTran.ocr.renderer")


def _hex_to_rgb(hex_color: str):
    """Convert '#RRGGBB' to (R, G, B) tuple."""
    hex_color = hex_color.lstrip("#")
    if len(hex_color) != 6:
        return (0, 0, 0)
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))


def _compute_font_size(text: str, bbox_w_pt: float, bbox_h_pt: float,
                        target_size_pt: int, min_size: int, max_size: int,
                        mode: str, scale_pct: int) -> int:
    """
    Compute the optimal font size based on the fit strategy.

    Parameters
    ----------
    text : str          – The text to fit
    bbox_w_pt : float   – Textbox width in points
    bbox_h_pt : float   – Textbox height in points
    target_size_pt : int – User-configured target size
    min_size : int       – Minimum font size
    max_size : int       – Maximum font size
    mode : str           – 'auto_fit', 'fixed', or 'scale_from_original'
    scale_pct : int      – Scale % (used when mode == 'scale_from_original')
    """
    if mode == "fixed":
        return max(min_size, min(target_size_pt, max_size))

    if mode == "scale_from_original":
        scaled = int(target_size_pt * scale_pct / 100)
        return max(min_size, min(scaled, max_size))

    # Auto Fit: estimate based on bbox area
    if bbox_h_pt <= 0 or bbox_w_pt <= 0:
        return max(min_size, min(target_size_pt, max_size))

    # Rough estimate: how many chars fit per line, how many lines
    text_len = max(len(text), 1)
    chars_per_line = max(int(bbox_w_pt / (target_size_pt * 0.6)), 1)
    lines_needed = max((text_len + chars_per_line - 1) // chars_per_line, 1)
    available_height = bbox_h_pt * 0.9  # 90% usable
    size_by_height = int(available_height / (lines_needed * 1.3))

    # Clamp
    result = max(min_size, min(size_by_height, target_size_pt, max_size))
    return result


class ImageTextRenderer:
    """
    Renders translated text blocks back into the image or host document (e.g. PPT/Excel).
    """
    def __init__(self, event_manager):
        self.em = event_manager

    def render_overlay(self, image_bytes: bytes, result: ImageOcrResult, fallback_font: str = "Arial") -> Optional[bytes]:
        """
        Renders translated blocks as an overlay on the original image using PIL.
        Returns the modified image bytes, or None if disabled or failed.
        """
        if not settings.get("ocr_settings.render_textbox_overlay", True):
            return None

        if not result.blocks or result.error:
            return None

        try:
            from PIL import Image, ImageDraw, ImageFont
            import io

            img = Image.open(io.BytesIO(image_bytes)).convert("RGBA")
            overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
            draw = ImageDraw.Draw(overlay)

            body_font_family = settings.get("text_style_settings.body_font_family", "Arial")
            body_size = settings.get("text_style_settings.body_font_size", 14)
            body_color = settings.get("text_style_settings.body_color", "#333333")
            bg_mode = settings.get("text_style_settings.background_mode", "semi_transparent")
            bg_color_hex = settings.get("text_style_settings.background_color", "#FFFFFF")
            bg_opacity = settings.get("text_style_settings.background_opacity", 180)
            min_size = settings.get("text_style_settings.min_font_size", 10)

            r, g, b = _hex_to_rgb(body_color)
            bg_r, bg_g, bg_b = _hex_to_rgb(bg_color_hex)

            try:
                font = ImageFont.truetype(f"{body_font_family.lower()}.ttf", body_size)
            except IOError:
                try:
                    font = ImageFont.truetype("arial.ttf", body_size)
                except IOError:
                    font = ImageFont.load_default()

            for block in result.blocks:
                rendered_text = block.translated_text or block.text
                if not rendered_text or not rendered_text.strip():
                    continue

                x, y, w, h = block.bbox

                # Background
                if bg_mode == "solid":
                    draw.rectangle([x, y, x + w, y + h], fill=(bg_r, bg_g, bg_b, 255))
                elif bg_mode == "semi_transparent":
                    draw.rectangle([x, y, x + w, y + h], fill=(bg_r, bg_g, bg_b, bg_opacity))

                draw.text((x + 2, y + 2), rendered_text, fill=(r, g, b, 255), font=font)

            result_img = Image.alpha_composite(img, overlay).convert("RGB")
            out_io = io.BytesIO()
            result_img.save(out_io, format="PNG")

            self.em.log("INFO", f"[{result.image_id}] render_overlay_success")
            return out_io.getvalue()

        except ImportError:
            self.em.log("WARN", f"[{result.image_id}] render_overlay_failed: PIL not installed.")
            return None
        except Exception as e:
            self.em.log("ERROR", f"[{result.image_id}] render_overlay_failed: {e}")
            return None
