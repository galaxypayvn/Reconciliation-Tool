#!/usr/bin/env python3

from pathlib import Path
from openpyxl import load_workbook
import csv
import json
import shutil
from typing import Dict, Any, List
import pandas as pd

INPUT_FOLDER = "DAILY_RECON/UPC"
FILE_PREFIX = "RECON_REPORT_INPUT"
WALLET_INPUT_FOLDER = "DAILY_RECON/WALLET"
WALLET_FILE_PREFIX = "DS_Daily_gui"
WALLET_NEW_FILE_PREFIX = "RECON_REPORT_WALLET"
WALLET_NEW_FILE_PREFIX_INPUT = "RECON_REPORT_WALLET_INPUT"
WALLET_NEW_FILE_PREFIX_OUTPUT = "RECON_REPORT_WALLET_OUTPUT"
CSV_FOLDER = "csv_recon_report"
OUTPUT_JSON = f"{CSV_FOLDER}/recon_export.json"
BANK_STATEMENT_FOLDER = "DAILY_RECON/STATEMENT"
BANK_STATEMENT_FILE_PREFIX = "BANK"

def convert_xls_to_csv(source_folder, file_prefix, output_folder):
    """
    Convert all .xls files matching file_prefix into CSV files.

    Each sheet in the XLS file becomes one CSV file.

    Example:
        source_folder = "./statement"
        file_prefix = "STATEMENT"
        output_folder = "./csv_recon_report"
    """

    source_folder = Path(source_folder)
    output_folder = Path(output_folder)

    output_folder.mkdir(parents=True, exist_ok=True)

    xls_files = sorted(
        file for file in source_folder.glob("*.xls")
        if file_prefix in file.name.upper()
    )

    if not xls_files:
        print(
            f"[WARNING] No XLS files found "
            f"in {source_folder} with prefix {file_prefix}"
        )
        return

    for xls_file in xls_files:

        print(f"[INFO] Processing: {xls_file.name}")

        try:
            excel_file = pd.ExcelFile(
                xls_file,
                engine="xlrd"
            )

            for sheet_name in excel_file.sheet_names:

                safe_sheet_name = (
                    sheet_name
                    .replace("/", "_")
                    .replace("\\", "_")
                    .replace(" ", "_")
                )

                csv_file = (
                    output_folder
                    / f"{xls_file.stem}_{safe_sheet_name}.csv"
                )

                df = pd.read_excel(
                    xls_file,
                    sheet_name=sheet_name,
                    engine="xlrd"
                )

                df.to_csv(
                    csv_file,
                    index=False,
                    encoding="utf-8-sig"
                )

                print(
                    f"[INFO] Converted "
                    f"{xls_file.name} / {sheet_name}"
                    f" -> {csv_file.name}"
                )
                

        except Exception as e:
            print(
                f"[ERROR] Failed to process "
                f"{xls_file.name}: {e}"
            )

def clear_csv_folder(folder: str):
    """
    Remove existing csv folder and recreate it.
    """
    path = Path(folder)

    if path.exists():
        shutil.rmtree(path)

    path.mkdir(parents=True, exist_ok=True)

    print(f"[INFO] Created clean folder: {folder}")


def convert_xlsx_to_csv(source_folder, file_prefix):
    """
    Convert all RECON_REPORT_INPUT*.xlsx files
    into CSV files under csv_recon_report.

    Each sheet becomes one CSV.
    """

    # source_folder = Path(INPUT_FOLDER)
    output_folder = Path(CSV_FOLDER)
    
    xlsx_files = sorted(
        source_folder.glob(f"{file_prefix}*.xlsx")
    )

    if not xlsx_files:
        print("[WARNING] No xlsx files found.")
        return

    for xlsx_file in xlsx_files:

        print(f"[INFO] Processing: {xlsx_file.name}")

        wb = load_workbook(
            filename=str(xlsx_file),
            data_only=True
        )

        for sheet in wb.worksheets:

            safe_sheet_name = (
                sheet.title
                .replace("/", "_")
                .replace("\\", "_")
                .replace(" ", "_")
            )

            csv_file = (
                output_folder
                / f"{xlsx_file.stem}_{safe_sheet_name}.csv"
            )

            with open(
                csv_file,
                mode="w",
                newline="",
                encoding="utf-8"
            ) as f:

                writer = csv.writer(f)

                for row in sheet.iter_rows(values_only=True):

                    cleaned_row = []

                    for cell in row:

                        if cell is None:
                            cleaned_row.append("")
                        else:
                            cleaned_row.append(str(cell))

                    writer.writerow(cleaned_row)

            print(
                f"[INFO] Converted "
                f"{xlsx_file.name} / {sheet.title}"
                f" -> {csv_file.name}"
            )


# Các cột dùng để nhận diện dòng header của từng loại file
HEADER_COLUMNS = {
    "DS_DAILY": ("TT", "Ngày giao dịch", "Mã giao dịch GPC (Bank)"),
    "RECON_REPORT_INPUT": ("TT", "Ngày giao dịch", "Transaction Id", "Kết quả đối soát"),
    "RECON_REPORT_WALLET": ("TT", "Transaction Date", "Reconciliation Status"),
    "HDBANK": ("Số tham chiếu", "Thông tin giao dịch"),
    "BVBANK": ("SỐ GIAO DỊCH (Reference Number)", "DIỄN GIẢI (Narrative)"),
}


def find_header_index(all_rows, required, file_name) -> int:
    """
    Quét từ trên xuống, trả về index của dòng đầu tiên chứa đủ các cột required.
    """

    for index, row in enumerate(all_rows):
        # Gộp khoảng trắng / xuống dòng trong ô (vd "SỐ GIAO DỊCH\n(Reference Number)")
        cells = {" ".join(str(col).split()) for col in row}

        if all(col in cells for col in required):
            return index

    raise ValueError(
        f"Cannot find header row with columns {list(required)}: {file_name}"
    )


def parse_csv_file(csv_file: Path) -> List[Dict[str, Any]]:
    """
    Parse CSV using first row as header.
    """

    rows = []

    with open(
        csv_file,
        mode="r",
        encoding="utf-8",
        errors="replace",
        newline=""
    ) as f:

        reader = csv.reader(f)

        all_rows = list(reader)
        
        # Determine header row: dòng header không cố định tuỳ file
        # -> lấy dòng đầu tiên có đủ các cột nhận diện của loại file
        if "DS_Daily" in csv_file.stem:
            required = HEADER_COLUMNS["DS_DAILY"]
        elif "RECON_REPORT_INPUT" in csv_file.stem.upper():
            required = HEADER_COLUMNS["RECON_REPORT_INPUT"]
        elif "RECON_REPORT_WALLET" in csv_file.stem.upper():
            required = HEADER_COLUMNS["RECON_REPORT_WALLET"]
        elif "HDBANK_068704070165977" in csv_file.stem.upper() and "PrintPage" in csv_file.stem:
            required = HEADER_COLUMNS["HDBANK"]
        elif "SAOKE_BVBank_0697041105669" in csv_file.stem:
            required = HEADER_COLUMNS["BVBANK"]
        else:
            raise ValueError(
                f"Unknown file type, cannot determine header row: {csv_file.name}"
            )

        header_index = find_header_index(all_rows, required, csv_file.name)

        headers = [
        str(col).strip() if str(col).strip() else f"col{i}"
        for i, col in enumerate(all_rows[header_index])
        ]

        # headers = [
        #     str(col).strip() if str(col).strip()
        #     else f"col{i}"
        #     for i, col in enumerate(all_rows[0])
        # ]

        for row in all_rows[header_index + 1:]:

            row_dict = {}

            for i, header in enumerate(headers):
                row_dict[header] = (
                    row[i]
                    if i < len(row)
                    else ""
                )

            rows.append(row_dict)

    return rows


def scan_csv_folder() -> Dict[str, Any]:
    """
    Scan all csv files and build JSON structure.
    """

    folder = Path(CSV_FOLDER)

    result = {}
    failed_files = []

    # csv_files = sorted(folder.glob("*.csv"))
    keywords = (
    "chi_tiết_giao_dịch",
    "chi_tiet_vi",
    "chi_tiết_ví",
    "bank"
    )

    csv_files = sorted(
        f for f in folder.glob("*.csv")
        if any(keyword in f.stem.lower() for keyword in keywords)
    )

    print(f"[INFO] Found {len(csv_files)} csv files")

    for csv_file in csv_files:

        # Sao kê HDBANK chỉ dùng sheet PrintPage, các sheet khác (Sheet2, Sheet3...) bỏ qua
        if (
            "HDBANK_068704070165977" in csv_file.stem.upper()
            and "PrintPage" not in csv_file.stem
        ):
            print(f"[SKIP] Unused HDBANK sheet: {csv_file.name}")
            continue

        try:

            data = parse_csv_file(csv_file)

            result[csv_file.name] = {
                "rows": len(data),
                "data": data
            }

            print(                                                                                                                                                                                                                                                                        
                f"[INFO] Parsed "
                f"{csv_file.name} "
                f"({len(data)} rows)"
            )

        except Exception as exc:

            failed_files.append(f"{csv_file.name}: {exc}")

            print(
                f"[ERROR] Failed "
                f"{csv_file.name}: {exc}"
            )

    # Dừng pipeline thay vì bỏ qua file lỗi (giữ lại CSV để kiểm tra)
    if failed_files:
        raise RuntimeError(
            "Cannot parse CSV files:\n  " + "\n  ".join(failed_files)
        )

    return result


def export_json(data: Dict[str, Any]):
    """
    Export final JSON file.
    """

    output_file = Path(OUTPUT_JSON)

    with open(
        output_file,
        mode="w",
        encoding="utf-8"
    ) as f:

        json.dump(
            data,
            f,
            ensure_ascii=False,
            indent=2
        )

    print(
        f"[INFO] JSON exported to "
        f"{output_file}"
    )


def main():

    print("=" * 80)
    print("STEP 1 - CLEAR CSV FOLDER")
    print("=" * 80)

    clear_csv_folder(CSV_FOLDER)

    print("\n" + "=" * 80)
    print("STEP 2 - CONVERT XLSX TO CSV")
    print("=" * 80)

    convert_xlsx_to_csv(Path(INPUT_FOLDER), FILE_PREFIX)
    convert_xlsx_to_csv(Path(WALLET_INPUT_FOLDER), WALLET_FILE_PREFIX)
    convert_xlsx_to_csv(Path(WALLET_INPUT_FOLDER), WALLET_NEW_FILE_PREFIX)
    convert_xls_to_csv(Path(BANK_STATEMENT_FOLDER), BANK_STATEMENT_FILE_PREFIX, CSV_FOLDER)
    
    print("\n" + "=" * 80)
    print("STEP 3 - SCAN CSV FILES")
    print("=" * 80)

    results = scan_csv_folder()

    print("\n" + "=" * 80)
    print("STEP 4 - EXPORT JSON")
    print("=" * 80)

    export_json(results)

    print("\n[SUCCESS] Process completed.")


if __name__ == "__main__":
    main()