#!/bin/bash
# Roadmap Pro — Quick Start
# This script installs dependencies and launches the app.

echo ""
echo "  ╔══════════════════════════════════════╗"
echo "  ║         Roadmap Pro v1.0             ║"
echo "  ║  Beautiful project visualizations    ║"
echo "  ╚══════════════════════════════════════╝"
echo ""

# Check Python
if ! command -v python3 &> /dev/null; then
    echo "Python 3 is required. Please install it from python.org"
    exit 1
fi

# Install dependencies
echo "Installing dependencies..."
pip3 install -q -r requirements.txt 2>/dev/null || pip install -q -r requirements.txt 2>/dev/null

echo ""
echo "Starting Roadmap Pro..."
echo "The app will open in your browser at http://localhost:8501"
echo ""

streamlit run app.py --server.headless true
