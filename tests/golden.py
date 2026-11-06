#!/usr/bin/env python3
"""
Golden (snapshot) test cho toàn bộ pipeline.

Chạy 3 step trên bộ file mẫu trong 1 thư mục tạm
(không đụng tới DAILY_RECON / csv_recon_report thật), rồi so sánh
recon_summary.csv và upload_file*.csv với snapshot đã lưu.

2 bộ file:
    DAILY_RECON/EXAMPLE  bộ nhỏ (~1 phút), chạy sau mỗi lần sửa code
    DAILY_RECON/FILES    bộ đầy đủ (~4 phút), chạy trước khi build exe

Usage:
    python3 tests/golden.py                   # bộ nhỏ: so sánh với snapshot
    python3 tests/golden.py --full            # bộ đầy đủ
    python3 tests/golden.py [--full] --update # ghi lại snapshot

Snapshot lưu ở tests/golden_data/<example|full>/ (gitignored vì chứa dữ liệu giao dịch thật).
Khi fail, output thực tế được ghi vào tests/golden_data/<example|full>/actual/.
"""

import difflib
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parent.parent
GOLDEN_DATA_DIR = Path(__file__).resolve().parent / "golden_data"

# tên bộ -> thư mục file mẫu
FIXTURE_SETS = {
    "example": PROJECT_DIR / "DAILY_RECON" / "EXAMPLE",
    "full": PROJECT_DIR / "DAILY_RECON" / "FILES",
}

SCRIPTS = (
    "convert_data_to_json.py",
    "get_data_4.py",
    "export_upload_file_1.py",
    "helper.py",
)

# Thư mục đích trong DAILY_RECON theo prefix tên file
FIXTURE_ROUTES = (
    ("RECON_REPORT_INPUT", "UPC"),
    ("RECON_REPORT_WALLET", "WALLET"),
    ("DS_Daily", "WALLET"),
    ("HDBANK", "STATEMENT"),
    ("SAOKE_BVBank", "STATEMENT"),
)

PIPELINE = (
    "import convert_data_to_json, get_data_4, export_upload_file_1\n"
    "convert_data_to_json.main()\n"
    "get_data_4.main()\n"
    "export_upload_file_1.main()\n"
)


def prepare_sandbox(sandbox: Path, fixture_dir: Path):
    for script in SCRIPTS:
        shutil.copy2(PROJECT_DIR / script, sandbox / script)

    for _, folder in FIXTURE_ROUTES:
        (sandbox / "DAILY_RECON" / folder).mkdir(parents=True, exist_ok=True)

    copied = 0
    for file in sorted(fixture_dir.iterdir()):
        for prefix, folder in FIXTURE_ROUTES:
            if file.name.startswith(prefix):
                shutil.copy2(file, sandbox / "DAILY_RECON" / folder / file.name)
                copied += 1
                break

    if copied == 0:
        sys.exit(f"[ERROR] No example files found in {fixture_dir}")


def run_pipeline(sandbox: Path) -> dict[str, str]:
    result = subprocess.run(
        [sys.executable, "-c", PIPELINE],
        cwd=sandbox,
        capture_output=True,
        text=True,
    )

    if result.returncode != 0:
        print(result.stdout[-5000:])
        print(result.stderr[-5000:])
        sys.exit("[ERROR] Pipeline failed")

    output_dir = sandbox / "csv_recon_report"
    upload_files = sorted(output_dir.glob("upload_file_*.csv"))

    if len(upload_files) != 1:
        sys.exit(f"[ERROR] Expected 1 upload file, found {len(upload_files)}")

    return {
        "recon_summary.csv": (output_dir / "recon_summary.csv").read_text(encoding="utf-8-sig"),
        "upload_file.csv": upload_files[0].read_text(encoding="utf-8-sig"),
    }


def main():
    update = "--update" in sys.argv
    fixture_set = "full" if "--full" in sys.argv else "example"
    fixture_dir = FIXTURE_SETS[fixture_set]
    snapshot_dir = GOLDEN_DATA_DIR / fixture_set

    print(f"[INFO] Fixture set: {fixture_set} ({fixture_dir})")

    with tempfile.TemporaryDirectory() as tmp:
        sandbox = Path(tmp)
        prepare_sandbox(sandbox, fixture_dir)
        outputs = run_pipeline(sandbox)

    if update:
        snapshot_dir.mkdir(parents=True, exist_ok=True)
        for name, content in outputs.items():
            (snapshot_dir / name).write_text(content, encoding="utf-8")
        print(f"[INFO] Snapshot updated in {snapshot_dir}")
        return

    failed = False

    for name, actual in outputs.items():
        snapshot = snapshot_dir / name

        if not snapshot.exists():
            sys.exit(f"[ERROR] Missing snapshot {snapshot}, run with --update first")

        expected = snapshot.read_text(encoding="utf-8")

        if actual == expected:
            print(f"[PASS] {name}")
            continue

        failed = True
        actual_file = snapshot_dir / "actual" / name
        actual_file.parent.mkdir(exist_ok=True)
        actual_file.write_text(actual, encoding="utf-8")
        print(f"[FAIL] {name} (actual output: {actual_file})")
        sys.stdout.writelines(
            difflib.unified_diff(
                expected.splitlines(keepends=True),
                actual.splitlines(keepends=True),
                fromfile=f"expected/{name}",
                tofile=f"actual/{name}",
            )
        )

    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
