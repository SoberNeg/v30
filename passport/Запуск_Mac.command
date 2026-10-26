#!/bin/bash
cd "$(dirname "$0")"
python3 -m pip install -q --user python-docx pillow pypdf 2>/dev/null
python3 app.py
