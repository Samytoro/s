#!/bin/bash
cd "$(dirname "$0")"
source venv_gui/bin/activate
python3 gui_merge.py
