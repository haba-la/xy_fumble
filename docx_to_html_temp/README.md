# DOCX Upload to Temporary HTML

This module provides a simple upload flow for converting `.docx` files into temporary `.html` files.
The generated HTML can be used as the next step of your formatting/editing pipeline.

## Features

- Upload DOCX file from browser
- Convert DOCX to HTML
- Save converted HTML into runtime temp folder
- Open the generated HTML by URL for later edits

## Install

```bash
pip install flask mammoth
```

## Run Web Upload Service

```bash
cd /root/all_rule_and_consistent/docx_to_html_temp
python web_upload.py --host 0.0.0.0 --port 8060
```

Open `http://127.0.0.1:8060` in your browser.

## CLI Conversion (without web)

```bash
python docx_to_html.py --docx /path/to/input.docx --temp-dir /root/all_rule_and_consistent/docx_to_html_temp/runtime/html_temp
```

Optional fixed output path:

```bash
python docx_to_html.py --docx /path/to/input.docx --output /path/to/output.html
```

## Output Paths

- Uploaded DOCX: `runtime/uploads/`
- Generated temp HTML: `runtime/html_temp/`
