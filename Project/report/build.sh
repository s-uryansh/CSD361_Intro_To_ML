#!/usr/bin/env bash
# Builds both deliverables from source:
#   1. report.html  -> CSD361_Preliminary_Report_StructX.pdf   (Playwright / Chromium, IEEE 2-column)
#   2. build_docx.py-> CSD361_Preliminary_Report_StructX.docx  (direct OOXML, Word-native styles)
set -euo pipefail
cd "$(dirname "$0")"

node extract_figure.js     # rasterise Fig. 1 for the Word build
node render.js             # PDF
python3 build_docx.py      # DOCX

echo
echo "Deliverables:"
ls -lh CSD361_Preliminary_Report_StructX.pdf CSD361_Preliminary_Report_StructX.docx
