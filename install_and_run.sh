#!/bin/bash
echo "============================================"
echo "  Salary Slip Generator - Installer"
echo "============================================"
echo ""

# Check Python3
if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python 3 is not installed."
    echo "Install it via: sudo apt install python3 python3-pip python3-tk (Linux)"
    echo "Or download from https://www.python.org/downloads/ (Mac)"
    exit 1
fi

echo "[1/2] Installing Python dependencies..."
python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt

if [ $? -ne 0 ]; then
    echo "ERROR: Failed to install dependencies."
    echo "Try: pip3 install openpyxl pandas Pillow reportlab"
    exit 1
fi

echo ""
echo "[2/2] Done!"
echo ""
echo "============================================"
echo "  Installation Complete!"
echo "  Default password: admin123"
echo "============================================"
echo ""
python3 salary_slip_app.py
