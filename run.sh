#!/bin/bash
# Transformation Office — Block & Gantt Creator Tool
# This script installs dependencies and launches the application.

echo ""
echo "  ╔═══════════════════════════════════════════════╗"
echo "  ║   Transformation Office                       ║"
echo "  ║   Block & Gantt Creator Tool  ·  v0.9 Beta    ║"
echo "  ╚═══════════════════════════════════════════════╝"
echo ""

# Check Python
if ! command -v python3 &> /dev/null; then
    echo "  [ERROR] Python 3 is required."
    echo "  Please install it from https://www.python.org/downloads/"
    exit 1
fi

# Install dependencies
echo "  Installing dependencies..."
pip3 install -q -r requirements.txt 2>/dev/null || pip install -q -r requirements.txt 2>/dev/null

echo ""
echo "  Starting application..."
echo "  The tool will open in your browser at http://localhost:8501"
echo ""

streamlit run app.py --server.headless true
