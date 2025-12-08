# Build script for generating the calendar executable using PyInstaller on Windows

# Check and install python-docx if not present
try {
    python -c "import docx" 2>$null
} catch {
    Write-Host "python-docx not found. Installing..."
    python -m pip install python-docx
}

# Check and install pyinstaller if not present
if (!(Get-Command pyinstaller -ErrorAction SilentlyContinue)) {
    Write-Host "PyInstaller not found. Installing..."
    python -m pip install pyinstaller
}

# Build the executable
pyinstaller --onefile generate_calendar.py