# Handwrite Ledger AI

Handwrite Ledger AI is a local desktop/browser tool for converting photographed handwritten ledgers and fixed-format tables into structured Excel outputs.

It combines repeated OCR extraction, rule-based field normalization, optional vision-language-model review, row-level visual audit, and template-preserving Excel export. The workflow is useful for handwritten forms where key fields such as name, ID number, address, phone, and amount must be extracted into a fixed spreadsheet format.

## What It Does

- Upload multiple ledger/table images and process them in one task.
- Run OCR three times per image and choose the most reliable result by consistency and confidence.
- Parse row-level fields such as name, ID number, address, phone, and amount.
- Optionally call an OpenAI-compatible vision LLM for row-by-row visual review.
- Generate an Excel workbook from an existing `.xls` template while preserving layout and formatting.
- Calculate derived numeric fields as static values, not formulas.
- Provide an audit screen with row crops beside editable extracted fields.
- Save audit edits and re-export the final workbook.
- Keep task history and downloadable artifacts locally.

## Why Add LLM_VL

Traditional OCR is weak on handwritten names, mixed Chinese/numeric fields, and column boundaries. This project adds an optional LLM_VL review layer that sees both:

- the original cropped row image
- the OCR text and coordinates

In testing on handwritten ledger photos, this layer improved extraction quality noticeably, especially for names and row alignment. The improvement comes from visual cross-checking: the model can correct OCR character confusions, ignore row numbers or margin marks, avoid merging adjacent rows, and choose the correct amount column.

The tool still works without LLM_VL, but best accuracy is expected when `VISION_LLM_*` is configured with a capable vision model.

## Architecture

```text
Images
  -> repeated OCR
  -> row grouping and field parsing
  -> optional LLM_VL row review
  -> audit UI
  -> template-preserving Excel export
```

## Configuration

No API keys are stored in this repository. Copy `.env.example` or set environment variables directly.

Required:

```powershell
$env:PADDLEOCR_TOKEN="your-ocr-token"
```

Optional LLM_VL:

```powershell
$env:VISION_LLM_API_URL="https://your-openai-compatible-endpoint/v1/chat/completions"
$env:VISION_LLM_API_KEY="your-vision-llm-key"
$env:VISION_LLM_MODEL="your-vision-model"
```

Optional default template:

```powershell
$env:LEDGER_TEMPLATE_PATH="C:\path\to\template.xls"
```

You can also upload a `.xls` template in the UI for each task.

## Run In Browser Mode

```powershell
python -m pip install -r requirements.txt
python app.py
```

Open:

```text
http://127.0.0.1:8765
```

## Run As Desktop App

```powershell
.\run_desktop.ps1
```

## Build Desktop Executable

```powershell
.\build_desktop.ps1
```

The generated executable is placed under:

```text
dist\LedgerOCRDesktop\LedgerOCRDesktop.exe
```

Move the whole `dist\LedgerOCRDesktop` folder when distributing the desktop build.

## Notes

- Output files and audit data are stored locally under `runs/`.
- `runs/`, `dist/`, `build/`, logs, and generated workbooks are intentionally ignored by Git.
- For public deployments, rotate any keys that were ever used during local development before sharing builds or logs.
