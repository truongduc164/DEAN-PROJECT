# Implementation Plan - Phase 1: Core Logic (Excel Focus)

## Goal Description
Build the backend "engine" capable of replicating the core features seen in the **TransLeo** reference images.
The immediate goal is to handle **Excel files (`.xlsx`)** programmatically, enabling:
1.  Reading workbook structure (sheets).
2.  Extracting text (Cells & shapes/textboxes).
3.  Sending text to a Translator Service (Mock/API).
4.  Writing results back (Overwrite or New File).
5.  **Crucial**: Emitting detailed events (Logs, Progress) so the future UI can display them.

## User Review Required
> [!IMPORTANT]
> **Dependency on APIs**: The reference app uses "Gemini | Model: models/gemini-2.5-flash". We will design the `TranslatorEngine` to be modular so we can plug in Gemini, OpenAI, or Google Translate later. For now, we will use a **Mock Translator** (simulated delay) to test the flow without burning API keys.

## Architecture Design (based on References)

The "TransLeo" app relies heavily on real-time feedback (Logs, Progress Bars). Our Core Logic must support this via an Event System.

### 1. Core Modules
-   **`EventManager`**: The central nervous system. Allows deep logic to send "signals" (Log Message, Progress Update, Error) to the UI without knowing about the UI.
-   **`ExcelProcessor`**:
    -   Uses `openpyxl`.
    -   Functions: `load_file`, `get_sheet_names`, `extract_text(sheet)`, `write_text(sheet)`.
    -   *Feature from Image*: Support for "Textboxes" (Shape processing).
-   **`TranslatorEngine`**:
    -   Interface: `translate(text, source_lang, target_lang)`.
    -   Supports batching (as seen in "Số lượng ở mỗi Batch tự động").

## Proposed Changes

### `app/core`

#### [NEW] [event_manager.py](file:///D:/0. Lập trình/1.DEANTRANS/DeanTran/app/core/event_manager.py)
-   Singleton pattern or global instance.
-   Basic signals: `log(message, level)`, `progress(current, total)`, `status(text)`.

#### [NEW] [excel_processor.py](file:///D:/0. Lập trình/1.DEANTRANS/DeanTran/app/core/excel_processor.py)
-   Class `ExcelProcessor`.
-   Method `analyze_file(path)` -> returns list of sheets (for that "Choose sheet" popup).
-   Method `process_sheet(sheet_name, translator_callback)`.

#### [NEW] [translator_service.py](file:///D:/0. Lập trình/1.DEANTRANS/DeanTran/app/core/translator_service.py)
-   Base class `BaseTranslator`.
-   Implementation `MockTranslator` (returns "Translated [Text]").

## Verification Plan

### Automated Tests
1.  **Test Script**: Create `tests/test_flow_phase1.py`
    -   Create a dummy Excel file.
    -   Run `ExcelProcessor` on it.
    -   Assert that logs are printed to console (simulating UI log window).
    -   Assert that a new file is created with "Translated" text.

### Manual Verification
-   Run the test script via CMD.
-   Check the output folder for the result file.
