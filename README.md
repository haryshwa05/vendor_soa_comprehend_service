# Vendor SOA Extractor

This utility converts Outlook `.msg` emails into a standardized Excel workbook for your reconciliation flow.

The output workbook is always a `Vendor SOA` sheet with exactly these columns:

- `INV Number`
- `INV Date`
- `Outstanding Amount`

It supports three extraction sources in priority order:

1. Email body HTML tables
2. Excel-like attachments: `.xlsx`, `.xlsm`, `.csv`
3. PDF attachments

## Files

- [handle-msg.py](./handle-msg.py): local desktop UI for manual use
- [msg_to_excel.py](./msg_to_excel.py): shared extraction and workbook logic
- [msg_api.py](./msg_api.py): FastAPI service for recon-engine integration
- [smoke_check.py](./smoke_check.py): lightweight validation checks

## How It Works

### 1. Message open

The system opens the `.msg` with `extract-msg`. If the file cannot be opened or is not a valid `.msg`, the request fails immediately.

### 2. Source selection order

The extractor tries sources in this exact order:

1. Body HTML tables
2. Excel attachments
3. PDF attachments

The first source that produces at least one valid canonical row wins. It does not merge rows across sources in this version. This keeps behavior deterministic for your recon engine.

### 3. Body extraction logic

- Reads `message.htmlBody`
- Parses all HTML `<table>` elements with BeautifulSoup
- Reconstructs row/column layout, including `rowspan` and `colspan`
- Picks usable rows beneath the detected header row
- Maps source headers to canonical fields using alias scoring

If a body table can map at least:

- invoice number
- outstanding amount

then rows are normalized and returned.

### 4. Excel attachment logic

Supported attachment types:

- `.xlsx`
- `.xlsm`
- `.csv`

For each supported attachment:

- Read each worksheet or CSV row matrix
- Remove fully blank rows
- Detect the most likely header row
- Map the headers to canonical fields
- Normalize the rows

The attachment/worksheet producing the most valid rows wins.

### 5. PDF attachment logic

For each PDF attachment, two strategies are attempted:

1. Table extraction with `pdfplumber.extract_tables()`
2. Plain-text fallback with `pdfplumber.extract_text()`

The PDF strategy producing the most valid rows wins.

The text fallback looks for:

- a date token such as `02/16/2026` or `2025-12-09`
- one or more money tokens like `$575.00`
- a likely invoice identifier before the date

The last money value on the line is treated as the outstanding amount.

### 6. Header mapping logic

The extractor maps arbitrary vendor headers into the three canonical fields using alias scoring.

Recognized invoice number aliases include:

- `Invoice Number`
- `Invoice No`
- `Invoice#`
- `Inv No`
- `Document Number`

Recognized invoice date aliases include:

- `Invoice Date`
- `Inv Date`
- `Date`
- `Bill Date`

Recognized amount aliases include:

- `Outstanding Amount`
- `Open Balance`
- `Balance Due`
- `Amount Due`
- `Total Due`

The best matching column is selected per field. If invoice number or amount cannot be mapped, that table is rejected.

### 7. Row normalization logic

For every accepted row:

- invoice number is trimmed and `.0` suffixes are removed
- invoice date is normalized to `YYYY-MM-DD` when it matches a supported format
- outstanding amount is normalized to a plain decimal string like `211.00`

Rows are dropped if:

- invoice number is missing
- outstanding amount is missing

Duplicate canonical rows are removed using the tuple:

- invoice number
- invoice date
- outstanding amount

### 8. Workbook output

The generated workbook contains one sheet named `Vendor SOA` with the exact three required columns.

When using the desktop UI, the file is written into the same folder as the script with a name like:

`original_email_vendor_soa.xlsx`

If the name already exists, the system appends `_1`, `_2`, and so on.

## Running Locally

### Desktop UI

```powershell
python handle-msg.py
```

### API server

```powershell
uvicorn msg_api:app --reload
```

Health check:

```powershell
curl http://127.0.0.1:8000/health
```

### API endpoints

#### `POST /extract`

Accepts one multipart `.msg` upload and returns JSON:

```json
{
  "filename": "sample.msg",
  "source_type": "pdf_attachment",
  "source_name": "StatementOfAccount.pdf",
  "source_details": "PDF table extraction from page 1, table 1.",
  "row_count": 2,
  "rows": [
    {
      "invoice_number": "7349700",
      "invoice_date": "2026-02-16",
      "outstanding_amount": "340.00"
    }
  ]
}
```

#### `POST /extract/xlsx`

Accepts one multipart `.msg` upload and returns the generated Excel workbook directly.

Response headers also include:

- `X-Source-Type`
- `X-Source-Name`
- `X-Source-Details`
- `X-Row-Count`

## Recon Integration

For your actual recon engine, the recommended integration is:

1. Upload `.msg` to `POST /extract`
2. Use the returned JSON rows directly if your backend already accepts canonical invoice data
3. Or call `POST /extract/xlsx` if you want a physical workbook artifact first

This keeps the extractor independent from the reconciliation engine while still producing the exact three fields your SOA flow needs.

## Validation

Run the built-in smoke checks:

```powershell
python smoke_check.py
```

These tests cover:

- canonical header mapping
- Excel attachment extraction
- PDF text fallback extraction

## Current Limits

- `.xls` attachments are not supported yet
- OCR/scanned image PDFs are not supported yet
- Plain-text email body parsing is not implemented yet
- Rows are taken from one winning source, not merged across body and attachments
