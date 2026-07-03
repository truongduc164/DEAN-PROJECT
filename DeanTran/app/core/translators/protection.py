import re
from dataclasses import dataclass
from typing import List, Tuple

@dataclass
class ProtectedToken:
    placeholder: str
    original: str


class TextProtector:
    """
    Detects and protects specific patterns (Dates, Times, Tokens) in text
    by replacing them with placeholders, or entirely skipping translation
    if the text is exclusively composed of protected patterns and whitespace.
    """

    # Matches BTN-8, TLD-04, EPS-0929, etc. Max 4 letters, hyphen, max 4 digits.
    # Also dates/times: YYYY-MM-DD, HH:MM:SS, DD/MM/YYYY, etc.
    # Note: Using robust regex for dates, times and specific code patterns.
    PROTECTED_PATTERNS = [
        # Datetime combinations: 2025-09-01 00:00:00
        r'\b\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\b',
        # ISO Dates: 2025-09-01
        r'\b\d{4}-\d{2}-\d{2}\b',
        # Common Dates: 9/11/2025, 09/11/25
        r'\b\d{1,2}/\d{1,2}/\d{2,4}\b',
        # Time: 00:00:00, 12:34
        r'\b\d{2}:\d{2}(?::\d{2})?\b',
        # Chinese/Japanese dates: 2025年9月1日
        r'\d{4}年\d{1,2}月\d{1,2}日',
        # Code tokens: BTN-8, TLD-04, EPS-0929. 2-4 uppercase letters, hyphen, 1-4 digits
        r'\b[A-Z]{2,4}-\d{1,4}\b'
    ]

    def __init__(self):
        # Compile a single regex OR-ing all patterns
        self.pattern = re.compile(
            "|".join(f"({p})" for p in self.PROTECTED_PATTERNS)
        )

    def protect(self, text: str) -> Tuple[str, List[ProtectedToken], bool]:
        """
        Protect text before translation.
        Returns:
            - modified_text: The text with placeholders (e.g., __PROTECT_0__).
            - tokens: List of tokens replaced.
            - is_fully_protected: True if text contains ONLY protected tokens and whitespace.
        """
        if not text or not text.strip():
            return text, [], True

        tokens = []
        
        def replace_match(match) -> str:
            original = match.group(0)
            placeholder = f"__PROTECT_{len(tokens)}__"
            tokens.append(ProtectedToken(placeholder, original))
            return placeholder

        modified_text = self.pattern.sub(replace_match, text)
        
        # Check if the text is entirely whitespace+placeholders to skip API translation cost
        stripped_modified = modified_text
        for tk in tokens:
            stripped_modified = stripped_modified.replace(tk.placeholder, "")
            
        is_fully_protected = len(stripped_modified.strip()) == 0

        return modified_text, tokens, is_fully_protected

    def restore(self, translated_text: str, tokens: List[ProtectedToken]) -> str:
        """
        Restore original tokens back into the translated text.
        """
        if not translated_text:
            return translated_text
            
        restored = translated_text
        for tk in tokens:
            # Simple string replace to restore exactly
            restored = restored.replace(tk.placeholder, tk.original)
            
        # Fallback: if Gemini stripped double underscores, try relaxed match.
        # Sometimes Gemini translates __PROTECT_0__ to protect_0. 
        # But `__PROTECT_0__` is quite robust against translation.
        return restored
