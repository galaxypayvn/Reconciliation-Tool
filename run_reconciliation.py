#!/usr/bin/env python3

import sys

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
    except Exception as e:
        print(f"❌ {name} failed")
        print(e)
        sys.exit(1)


def main():
    run_step("STEP 1 - Convert Excel -> CSV/JSON", convert_main)
    run_step("STEP 2 - Generate recon_summary.csv", get_data_main)
    run_step("STEP 3 - Generate upload_file.csv", export_main)

    print("\n🎉 All steps completed successfully.")


if __name__ == "__main__":
    main()