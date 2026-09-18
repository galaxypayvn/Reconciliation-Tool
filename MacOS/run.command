#!/bin/bash

cd "$(dirname "$0")/.."

python3 run_reconciliation.py

echo
echo "======================================"
echo "Completed."

exit 0