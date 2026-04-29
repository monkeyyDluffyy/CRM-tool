#!/bin/bash
# NexusCRM - Start Script
#
# Usage:
#   ./start.sh                                    (uses default CRM_Sample_Data_Template.xlsx)
#   ./start.sh /path/to/your/excel-file.xlsx      (uses your custom Excel file)

cd "$(dirname "$0")"
source venv/bin/activate

if [ -n "$1" ]; then
    echo ""
    echo "  ⚡ Starting NexusCRM..."
    echo "  📊 Data source: $1"
    echo "  🌐 Open http://localhost:5000 in your browser"
    echo ""
    python3 app.py "$1"
else
    echo ""
    echo "  ⚡ Starting NexusCRM..."
    echo "  📊 Data source: CRM_Sample_Data_Template.xlsx"
    echo "  🌐 Open http://localhost:5000 in your browser"
    echo ""
    python3 app.py
fi
