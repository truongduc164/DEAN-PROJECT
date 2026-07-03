import pytest
from app.core.translators.protection import TextProtector

def test_protect_dates():
    protector = TextProtector()
    text = "Meeting on 2025-09-01 at 00:00:00 is important."
    
    modified, tokens, fully = protector.protect(text)
    assert not fully
    assert "2025-09-01 00:00:00" not in modified
    assert "__PROTECT_0__" in modified
    
    # Simulate translation
    translated = modified.replace("Meeting", "Cuộc họp").replace("important", "quan trọng")
    restored = protector.restore(translated, tokens)
    
    assert "Cuộc họp" in restored
    assert "2025-09-01 at 00:00:00" in restored
    assert "__PROTECT_0__" not in restored

def test_protect_japan_dates():
    protector = TextProtector()
    text = "Date is 2025年9月1日."
    
    modified, tokens, fully = protector.protect(text)
    assert not fully
    assert "__PROTECT_0__" in modified
    
    restored = protector.restore(modified, tokens)
    assert "2025年9月1日" in restored

def test_protect_tokens():
    protector = TextProtector()
    text = "Replace BTN-8 and TLD-04 in EPS-0929."
    
    modified, tokens, fully = protector.protect(text)
    assert not fully
    assert "BTN-8" not in modified
    assert "EPS-0929" not in modified
    assert "__PROTECT_0__" in modified
    assert "__PROTECT_2__" in modified
    
    restored = protector.restore(modified, tokens)
    assert "BTN-8" in restored
    assert "EPS-0929" in restored

def test_fully_protected():
    protector = TextProtector()
    text = " 2025-09-01   00:00:00 "
    modified, tokens, fully = protector.protect(text)
    assert fully  # True because it's only protected codes and whitespace
    
    # Should restore to original exactly
    restored = protector.restore(modified, tokens)
    assert restored == text
    
def test_fully_protected_tokens():
    protector = TextProtector()
    text = "BTN-8"
    modified, tokens, fully = protector.protect(text)
    assert fully
    
    restored = protector.restore(modified, tokens)
    assert restored == text
