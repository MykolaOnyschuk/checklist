#!/bin/bash
# Double-click this file in Finder to launch the Report Generator
cd "$(dirname "$0")"
python3 report_generator.py
echo ""
echo "Press Enter to close this window..."
read
