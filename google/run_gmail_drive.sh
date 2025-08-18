#!/bin/bash

# Gmail and Google Drive Automation Runner
# This script runs the email tracking automation

echo
echo "========================================"
echo "  Gmail and Google Drive Automation"
echo "========================================"
echo

# Check if virtual environment exists
if [ ! -d ".venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv .venv
    if [ $? -ne 0 ]; then
        echo "Error: Failed to create virtual environment"
        exit 1
    fi
fi

# Activate virtual environment
echo "Activating virtual environment..."
source .venv/bin/activate
if [ $? -ne 0 ]; then
    echo "Error: Failed to activate virtual environment"
    exit 1
fi

# Check if requirements are installed
echo "Checking dependencies..."
pip show google-auth >/dev/null 2>&1
if [ $? -ne 0 ]; then
    echo "Installing requirements..."
    pip install -r requirements_gmail_drive.txt
    if [ $? -ne 0 ]; then
        echo "Error: Failed to install requirements"
        exit 1
    fi
fi

# Run the automation
echo
echo "Starting Gmail and Drive automation..."
echo
python gmail_drive_automation.py

# Check if script ran successfully
if [ $? -ne 0 ]; then
    echo
    echo "Error: Automation failed"
    exit 1
else
    echo
    echo "Automation completed successfully!"
fi

# Deactivate virtual environment
deactivate

echo
echo "Press Enter to exit..."
read