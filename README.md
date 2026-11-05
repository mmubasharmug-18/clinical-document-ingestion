# Clinical Document Ingestion

A Streamlit app that reads clinical files in different formats (PDF, DOCX, TXT, HTML, scanned PDF and images) and turns every one of them into the same clean, structured output. It uses the same look as the Guideline PDF Parser.

## Quick notes

- **Why one pipeline:** clinical notes arrive as Word files, web pages, plain text, PDFs and phone photos. Later steps (search, summaries, coding) should not care which one it was, so everything is converted to one shape first.
- **File type detection:** the app looks at the first bytes of the file, then at the extension. A PDF named `.txt` is still read as a PDF.
- **One reader per format:** PyMuPDF for PDF, python-docx for DOCX, BeautifulSoup for HTML, a plain reader for TXT and Tesseract for images.
- **Scanned PDFs:** if a PDF page has almost no text, it is treated as a scan. The page is turned into an image and sent to Tesseract.
- **Blocks:** the text is saved as blocks. A block is a heading, a paragraph, a list item or a table.
- **Metadata:** the title, author and dates are kept when the file has them. Source details are always added: file name, size, SHA-256 and the time of extraction.
- **OCR confidence:** Tesseract gives each piece of text a score from 0 to 1. Low scores mean the text should be checked by eye.

## What the app shows

| Tab | What you get |
|---|---|
| Summary | One row per file with counts, plus a list of problems found |
| Text | The full extracted text, or one page at a time, with a download button |
| Structure | Every block with its page, kind, heading level, source and confidence |
| Metadata | Source details and the properties stored inside the file |
| Export | The standard output as JSON, for one file or all files |

The numbers:

- **Pages:** page count. Only PDFs and images have one.
- **Blocks:** how many headings, paragraphs, list items and tables were found.
- **Words:** word count of the extracted text.
- **OCR pages:** pages that had to be read with OCR.
- **OCR confidence:** average score from Tesseract, 0 to 1.
- **Seconds:** time taken to read the file.
- **Warnings:** problems found, such as a scanned page with OCR turned off.

## The standard output

Every format ends up like this:

```
source     file_name, extension, detected_type, size_bytes, sha256, extracted_at
type       pdf, docx, html, txt or image
method     which library read it
metadata   title, author, dates and other properties the file had
stats      pages, blocks, words, characters, ocr_pages, avg_ocr_confidence, seconds
warnings   list of problems
blocks     kind, text, page, level, source (text or ocr), confidence, rows (tables)
text       all block text joined together
```

## Folder structure

```
clinical_ingestion/
  app.py                      the Streamlit app
  extractors.py               file type detection and all the readers
  make_samples.py             builds the made-up sample files
  .streamlit/
    config.toml               light theme and colours
  samples/
    practice_guideline.pdf
    scanned_progress_note.pdf
    referral_letter.docx
    discharge_summary.txt
    lab_report.html
    clinic_note_photo.png
  requirements.txt
  README.md
```

`.streamlit` must be a folder with `config.toml` inside it. All sample files use made-up patient data.

## Set up (Windows, VS Code)

Install Tesseract first. On Windows, use the installer from the UB Mannheim page and keep the default folder. Then open the project folder in VS Code, open a terminal (Terminal > New Terminal) and run:

```
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

If PowerShell blocks the activate step, run `Set-ExecutionPolicy -Scope Process Bypass` once and try again.

If the app says Tesseract was not found, set the path once in the same terminal:

```
set TESSERACT_CMD=C:\Program Files\Tesseract-OCR\tesseract.exe
```

On Mac use `brew install tesseract`. On Linux use `sudo apt install tesseract-ocr`.

## Run

```
streamlit run app.py
```

It opens at http://localhost:8501.

## How to use it

1. Pick the files in the sidebar: the sample files, or upload your own.
2. Set the OCR options: turn OCR on or off for scanned PDF pages, scan quality (DPI) and language.
3. Look at the **Summary** tab to compare all files.
4. Use **View > Open file** to pick one file, then check **Text**, **Structure** and **Metadata**.
5. Open **Export** to download the standard output as JSON.

## What I analyze

| File | Type | Method | Blocks | Words | OCR confidence |
|---|---|---|---|---|---|
| discharge_summary.txt | txt | Plain text reader | 13 | 96 | - |
| lab_report.html | html | BeautifulSoup | 8 | 55 | - |
| referral_letter.docx | docx | python-docx | 10 | 64 | - |
| clinic_note_photo.png | image | Tesseract OCR | 3 | 52 | 0.95 |

- The TXT file gave 5 headings, 3 paragraphs and 5 list items.
- The HTML file gave 3 headings, 2 paragraphs, 2 list items and 1 table. The script and style tags were ignored.
- The DOCX file gave 3 headings, 3 paragraphs, 3 list items and 1 table. The author, title and the header text were kept as metadata.
- The photo is slightly tilted and noisy, and OCR still read it with a score of 0.95.

## Limits

- Heading detection in PDFs and TXT files is rule based. In a PDF, a short line with a bigger or bold font counts as a heading. In a TXT file, a short line in capitals or ending with a colon counts as one. It can get this wrong.
- OCR quality depends on the scan. Blurry, tilted or low resolution pages give low confidence and wrong words.
- Handwriting is not read well by Tesseract.
- Tables in scanned pages come out as plain text, not as table rows.
- Images inside DOCX and HTML files are not read. The app shows a warning for DOCX.
- Old `.doc` files, password protected PDFs and unknown formats are not supported. The app shows a warning instead of crashing.
- Page count is only available for PDFs and images.
- The OCR language needs its Tesseract language file installed. The default is `eng`.