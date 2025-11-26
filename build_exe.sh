#!/bin/bash
# Build script for creating Audentes Automation Tool executable (Linux/Mac)
# For Windows, use build_exe.bat instead

echo "========================================"
echo "Building Audentes Automation Tool"
echo "========================================"
echo ""

# Check if PyInstaller is installed
if ! python -m pip show pyinstaller > /dev/null 2>&1; then
    echo "Installing PyInstaller..."
    python -m pip install pyinstaller==6.16.0
fi

echo ""
echo "Building executable..."
echo ""

# Build the executable
pyinstaller --onefile \
    --noconsole \
    --name "AudentesAutomationTool" \
    --add-data "config.json:." \
    --hidden-import=pandas \
    --hidden-import=openpyxl \
    --hidden-import=selenium \
    --hidden-import=tkinter \
    main.py

if [ $? -ne 0 ]; then
    echo ""
    echo "Build failed! Check the error messages above."
    exit 1
fi

echo ""
echo "========================================"
echo "Build Successful!"
echo "========================================"
echo ""
echo "Executable location: dist/AudentesAutomationTool"
echo ""
echo "Next steps:"
echo "1. Copy AudentesAutomationTool to deployment folder"
echo "2. Copy config.json to the same folder"
echo "3. Test the executable"
echo ""

