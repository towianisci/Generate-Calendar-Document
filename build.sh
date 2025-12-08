#!/bin/bash

# Build script for generating the calendar executable using PyInstaller

# Define paths (using system Python and tools)
PYTHON_PATH=python3
PIP_PATH=pip3
PYINSTALLER_PATH=~/Library/Python/3.9/bin/pyinstaller

# Check and install python-docx if not present
if ! $PYTHON_PATH -c "import docx" 2>/dev/null; then
    echo "python-docx not found. Installing..."
    $PYTHON_PATH -m pip install python-docx
fi

# Check and install pyinstaller if not present
if ! [ -x "$PYINSTALLER_PATH" ]; then
    echo "PyInstaller not found. Installing..."
    $PYTHON_PATH -m pip install pyinstaller
fi

# Build the executable
$PYINSTALLER_PATH --onefile generate_calendar.py