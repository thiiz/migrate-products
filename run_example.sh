#!/bin/bash

echo "CSV to XLS Template Converter"
echo "--------------------------"
echo

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "Python 3 not found. Please install Python 3.6 or higher."
    exit 1
fi

# Install dependencies if needed
echo "Installing required dependencies..."
pip3 install -r requirements.txt

# Run the conversion
echo
echo "Converting products_export_1.csv to XLS format..."
python3 csv_to_xls.py "products_export_1 (1).csv"

echo
if [ $? -eq 0 ]; then
    echo "Conversion completed successfully!"
else
    echo "Conversion failed. Please check the error messages above."
fi

echo
read -p "Press Enter to continue..."