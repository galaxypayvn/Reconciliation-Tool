#!/usr/bin/env python3

import sys
import traceback

from convert_data_to_json import main as convert_main
from get_data_4 import main as get_data_main
from export_upload_file_1 import main as export_main


def run_step(name, func):
    print(f"\n{'='*60}")
    print(name)
    print("="*60)

    try:
        func()
        print(f"✅ {name} completed")
        return True
    except Exception as e:
        print(f"❌ {name} failed")
        print(f"Error: {e}")
        traceback.print_exc()
        return False


def main():
    steps = [
        ("STEP 1 - Convert Excel -> CSV/JSON", convert_main),
        ("STEP 2 - Generate recon_summary.csv", get_data_main),
        ("STEP 3 - Generate upload_file.csv", export_main),
    ]

    for name, func in steps:
        if not run_step(name, func):
            # Dừng ngay: các step sau sẽ đọc file cũ / thiếu
            print("\n⛔ Stopped. Do NOT use files in csv_recon_report from this run.")
            input("\n⏸ Press Enter to close the Error...")
            sys.exit(1)

    print("\n🎉 All steps completed successfully.")
    input("\n Press Enter to close the program...")

if __name__ == "__main__":
    main()
