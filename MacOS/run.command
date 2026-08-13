#!/bin/bash

cd "$(dirname "$0")/.."

python3 Scripts/run_reconciliation.py

echo
echo "======================================"
echo "Completed."
echo "Press any key and close..."
read -n 1 -s
exit 0