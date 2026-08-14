#!/usr/bin/env python3

import calendar
from datetime import datetime 
import json
import csv
import re
from collections import defaultdict
from export_upload_file_1 import get_date_range_from_filename, parse_posting_date_from_filename
from pathlib import Path
import sys
import os


INPUT_FILE = "csv_recon_report/recon_export.json"
SUMMARY_JSON_FILE = "csv_recon_report/recon_summary.json"
SUMMARY_CSV_FILE = "csv_recon_report/recon_summary.csv"

MERCHANT_ID_KEY = "Merchant ID"
PAYMENT_TYPE_KEY = "Loại giao dịch"
CURRENCY_KEY = "Loại tiền tệ"
MERCHANT_NAME_KEY = "Đối tác"
AMOUNT_KEY = "Số tiền"
ORIGIN_PAYMENT_DATE_KEY = "Ngày giao dịch gốc"
RECON_RESULT_KEY = "Kết quả đối soát"
NOTE_KEY = "Ghi chú"
POSTING_DATE_KEY = "Ngày giao dịch"
IS_CHARGE_KEY = "IsChargeFee"


def get_posting_date_from_range(cell_value, file_name):
    """
    Ví dụ file:
        DS_Daily_30.08-02.09.xlsx

    Kết quả:
        30/08/2026 -> 31/08/2026
        31/08/2026 -> 31/08/2026
        01/09/2026 -> 02/09/2026
        02/09/2026 -> 02/09/2026

    Các ngày ngoài range sẽ giữ nguyên.
    """

    if not cell_value or not file_name:
        return cell_value

    # ==========================================
    # Lấy range từ filename
    # Ví dụ: 30.08-02.09
    # ==========================================
    range_match = re.search(
        r"(\d{1,2})\.(\d{1,2})-(\d{1,2})\.(\d{1,2})",
        file_name
    )

    if not range_match:
        return cell_value

    start_day = int(range_match.group(1))
    start_month = int(range_match.group(2))
    end_day = int(range_match.group(3))
    end_month = int(range_match.group(4))

    # ==========================================
    # Parse cell date
    # ==========================================
    try:
        cell_date = datetime.strptime(
            cell_value.strip(),
            "%d/%m/%Y"
        )
    except ValueError:
        return cell_value

    year = cell_date.year

    # ==========================================
    # CASE 1:
    # Range cùng tháng
    #
    # 01.09-05.09
    # ==========================================
    if start_month == end_month:

        if (
            cell_date.month == start_month
            and start_day <= cell_date.day <= end_day
        ):
            return f"{end_day:02d}/{end_month:02d}/{year}"

        return cell_value

    # ==========================================
    # CASE 2:
    # Range qua 2 tháng
    #
    # 30.08-02.09
    # ==========================================

    # Tháng bắt đầu:
    # 30/08, 31/08
    # -> 31/08
    if (
        cell_date.month == start_month
        and cell_date.day >= start_day
    ):
        last_day = calendar.monthrange(
            year,
            start_month
        )[1]

        return f"{last_day:02d}/{start_month:02d}/{year}"

    # Tháng kết thúc:
    # 01/09, 02/09
    # -> 02/09
    elif (
        cell_date.month == end_month
        and cell_date.day <= end_day
    ):
        return f"{end_day:02d}/{end_month:02d}/{year}"

    # Ngoài range
    return cell_value

def format_money(value: int) -> str:
    return f"{value:,}"

def export_summary_to_json(data, output_file=SUMMARY_JSON_FILE):
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def process_file_data(file_name, rows):

    print("Processing:", file_name)

    grouped_data = defaultdict(lambda: {
        "total_amount": 0,
        "txn_count": 0,
        "currencies": set(),
        "note_raw": "",
        "origin_date": "",
        POSTING_DATE_KEY: ""
    })

    is_recon_input = "RECON_REPORT_INPUT" in file_name
    is_daily_wallet = "DS_Daily" in file_name

    # ======================================================
    # HELPER: GET POSTING DATE
    # ======================================================

    def get_posting_date(row):
        pattern_days = r"\((\d{1,2})-(\d{1,2})\)"
        days_match = re.search(pattern_days, file_name)

        pattern_months = r"\((\d{1,2})\.(\d{1,2})-(\d{1,2})\.(\d{1,2})\)"
        month_match = re.search(pattern_months, file_name)

        posting_date = row.get(
            POSTING_DATE_KEY, ""
        ).split(" ")[0]

        if days_match:
            posting_date = parse_posting_date_from_filename(file_name)

        if month_match:
            posting_date = get_posting_date_from_range(
                posting_date,
                file_name
            )

        return posting_date, days_match

    # ======================================================
    # RECON_REPORT_INPUT
    # ======================================================

    if is_recon_input:

        for row in rows:

            posting_date, days_match = get_posting_date(row)

            merchant_name = row.get(
                MERCHANT_NAME_KEY, ""
            ).strip()

            merchant_id = row.get(
                MERCHANT_ID_KEY, ""
            ).strip()

            payment_type = row.get(
                PAYMENT_TYPE_KEY, ""
            ).strip()

            recon_result = row.get(
                RECON_RESULT_KEY, ""
            ).strip()

            currency = row.get(
                CURRENCY_KEY, ""
            ).strip()

            note_raw = row.get(
                NOTE_KEY, ""
            ).strip()

            origin_date = row.get(
                ORIGIN_PAYMENT_DATE_KEY, ""
            )

            is_charge = row.get(
                IS_CHARGE_KEY, ""
            )

            # ==================================================
            # GOM MERCHANT VIETJET
            # ==================================================

            if (
                "VIETJET" in merchant_id.upper()
                or "VJ" in merchant_id.upper()
            ):
                if (
                    "KIOSK" not in merchant_id.upper()
                    and "SKYSHOP" not in merchant_id.upper()
                ):
                    merchant_id = "VIETJET"

            # ==================================================
            # AMOUNT
            # ==================================================

            try:
                amount = int(float(
                    row.get(AMOUNT_KEY, 0)
                ))
            except:
                amount = 0

            note_group = ""

            # ==================================================
            # MPGS
            # ==================================================

            if (
                "MPGS" in merchant_name.upper()
                and origin_date
                and " " in origin_date
                and recon_result == "MATCH_SUCCESS"
                and payment_type == "PAYMENT"
            ):

                posting_date = origin_date.split(" ")[0]

                if days_match:
                    posting_date = parse_posting_date_from_filename(
                        file_name
                    )

            # ==================================================
            # MISS GPAY
            # ==================================================

            if recon_result == "MISS_GPAY":

                note_lower = note_raw.lower()

                if "lệch ngày" in note_lower:
                    note_group = "lệch ngày"

                elif (
                    "galaxy pay" in note_lower
                    and "thất bại" in note_lower
                ):
                    note_group = "Galaxy Pay: thất bại"

                else:
                    note_group = "other_miss_gpay"

            # ==================================================
            # REFUND
            # ==================================================

            if payment_type == "REFUND":

                if "IBFT không thành công" in note_raw:
                    note_group = "IBFT không thành công"

                elif is_charge == "false":
                    note_group = "MISS_REFUND"

                else:
                    note_group = "REFUND"

            # ==================================================
            # MISS PARTNER / TRANSFER
            # ==================================================

            if (
                recon_result == "MISS_PARTNER"
                or "TRANSFER" in payment_type
            ):
                recon_result = ""
                note_group = "MISS_PARTNER"

            # ==================================================
            # GROUP
            # ==================================================

            group_key = (
                merchant_name,
                merchant_id,
                payment_type,
                recon_result,
                posting_date,
                note_group
            )

            info = grouped_data[group_key]

            info["total_amount"] += amount
            info["txn_count"] += 1
            info["currencies"].add(currency)

            if not info["note_raw"]:
                info["note_raw"] = note_raw

            if not info["origin_date"]:
                info["origin_date"] = origin_date

            info[POSTING_DATE_KEY] = posting_date

            # ==================================================
            # REFUND NOTE
            # ==================================================

            if (
                payment_type == "REFUND"
                and recon_result == "MATCH_SUCCESS"
            ):
                if note_group == "MISS_REFUND":
                    info["note_raw"] = "MISS_REFUND_FALSE"

    # ======================================================
    # DS_Daily
    # ======================================================

    elif is_daily_wallet:

        # ==================================================
        # 1. BANK SIDE
        # ==================================================

        print("[INFO] Processing BANK side")

        for row in rows:

            bank_partner = row.get(
                "Đối tác (Bank)",
                ""
            ).strip()

            # Không phải Bank row -> skip
            if not bank_partner:
                continue

            posting_date, days_match = get_posting_date(row)

            note_raw = row.get(
                "Ghi chú",
                ""
            ).strip()

            currency = row.get(
                CURRENCY_KEY,
                ""
            ).strip()

            origin_date = ""

            merchant_name = bank_partner
            merchant_id = bank_partner

            raw_payment_type = row.get(
                "Loại giao dịch",
                ""
            ).strip()

            recon_text = row.get(
                "Kết quả đối soát\n(Bank)",
                ""
            ).strip()

            # ==================================================
            # PAYMENT TYPE
            # ==================================================

            if raw_payment_type == "PAYMENT-CASHIN":
                payment_type = "CASHIN"
            else:
                payment_type = raw_payment_type

            # ==================================================
            # RECON RESULT
            # ==================================================

            if recon_text == "Khớp":
                recon_result = "MATCH_SUCCESS"

            elif recon_text in [
                "Lệch (Đối tác nhiều hơn)",
                "Xử lý lệch (Đối tác nhiều hơn)"
            ]:
                recon_result = "MISS_PARTNER"

            elif recon_text == "Lệch (Galaxy Pay nhiều hơn)":
                recon_result = "MISS_GPAY"

            elif recon_text == "Xử lý lệch (Galaxy Pay nhiều hơn)":
                recon_result = "GPAY_MORE"

            else:
                recon_result = ""

            # ==================================================
            # HDBANK / BIDV / HDB-VA
            # ==================================================

            if bank_partner.upper() in [
                "HDBANK",
                "BIDV",
                "HDB-VA"
            ]:

                if (
                    recon_text == "Lệch (Đối tác nhiều hơn)"
                    and payment_type == "CASHOUT"
                ):
                    recon_result = "MISS_BANK"

                elif (
                    recon_text == "Xử lý lệch (Galaxy Pay nhiều hơn)"
                    and payment_type == "CASHIN"
                ):
                    recon_result = "GPAY_MORE"

                elif (
                    recon_text == "Xử lý lệch (Galaxy Pay nhiều hơn)"
                    and payment_type == "CASHOUT"
                ):
                    recon_result = "GPAY_MORE"

                elif (
                    recon_text == "Lệch (Galaxy Pay nhiều hơn)"
                    and payment_type == "CASHOUT"
                ):
                    recon_result = "MISS_GPAY"

                else:
                    recon_result = ""

            # ==================================================
            # NAPAS
            # ==================================================

            elif bank_partner.upper() == "NAPAS":

                if (
                    recon_text == "Lệch (Đối tác nhiều hơn)"
                    and payment_type == "CASHIN"
                ):
                    recon_result = "MISS_PARTNER"

                elif (
                    recon_text == "Xử lý lệch (Galaxy Pay nhiều hơn)"
                    and payment_type == "CASHIN"
                ):
                    recon_result = "GPAY_MORE"

                elif (
                    recon_text == "Xử lý lệch (Đối tác nhiều hơn)"
                ):
                    recon_result = "PARTNER_MORE"
                    posting_date = parse_posting_date_from_filename(file_name)

            # ==================================================
            # TRANSFER
            # ==================================================

            if "TRANSFER" in payment_type:
                recon_result = ""

            # ==================================================
            # AMOUNT BANK
            # ==================================================

            try:
                amount = int(float(
                    row.get(
                        "Số tiền\n(Bank)",
                        0
                    )
                ))
            except:
                amount = 0

            note_group = ""

            # ==================================================
            # GROUP BANK
            # ==================================================

            group_key = (
                merchant_name,
                merchant_id,
                payment_type,
                recon_result,
                posting_date,
                note_group
            )

            info = grouped_data[group_key]

            info["total_amount"] += amount
            info["txn_count"] += 1
            info["currencies"].add(currency)
            info[POSTING_DATE_KEY] = posting_date

            if not info["note_raw"]:
                info["note_raw"] = note_raw

            if not info["origin_date"]:
                info["origin_date"] = origin_date

        # ==================================================
        # 2. PAYMENT SIDE
        # ==================================================

        print("[INFO] Processing PAYMENT side")

        for row in rows:

            payment_partner = row.get(
                "Đối tác (Payment)",
                ""
            ).strip()

            # Không phải Payment row -> skip
            if not payment_partner:
                continue

            posting_date, days_match = get_posting_date(row)

            note_raw = row.get(
                "Ghi chú",
                ""
            ).strip()

            currency = row.get(
                CURRENCY_KEY,
                ""
            ).strip()

            origin_date = ""

            merchant_name = payment_partner

            merchant_id = row.get(
                "ĐVCNTT",
                ""
            ).strip()

            raw_payment_type = row.get(
                "Loại giao dịch",
                ""
            ).strip()

            # ==================================================
            # PAYMENT TYPE
            # ==================================================

            if raw_payment_type == "PAYMENT-CASHIN":
                payment_type = "PAYMENT"
            else:
                payment_type = raw_payment_type

            recon_text = row.get(
                "Kết quả đối soát\n(Payment)",
                ""
            ).strip()

            # ==================================================
            # PARTNER MORE
            # ==================================================

            if recon_text == "Lệch (Đối tác nhiều hơn)":
                continue

            # ==================================================
            # RECON RESULT
            # ==================================================

            if recon_text == "Khớp":

                recon_result = "MATCH_SUCCESS"
                note_group = ""

            elif (
                recon_text
                == "Xử lý lệch (Galaxy Pay nhiều hơn)"
            ):

                recon_result = "GPAY_MORE"
                note_group = "GPAY_MORE"

            elif (
                recon_text
                == "Lệch (Galaxy Pay nhiều hơn)"
            ):

                note_lower = note_raw.lower()

                if (
                    "thất bại" in note_lower
                    or "không thành công" in note_lower
                ):
                    recon_result = "MISS_GPAY"
                    note_group = "Galaxy Pay: thất bại"

                else:
                    recon_result = "MATCH_SUCCESS"
                    note_group = "GPAY_MORE"

            else:

                recon_result = ""
                note_group = ""

            # ==================================================
            # AMOUNT PAYMENT
            # ==================================================

            try:
                amount = int(float(
                    row.get(
                        "Số tiền\n(Payment)",
                        0
                    )
                ))
            except:
                amount = 0

            # ==================================================
            # GROUP PAYMENT
            # ==================================================

            group_key = (
                merchant_name,
                merchant_id,
                payment_type,
                recon_result,
                posting_date,
                note_group
            )

            info = grouped_data[group_key]

            info["total_amount"] += amount
            info["txn_count"] += 1
            info["currencies"].add(currency)
            info[POSTING_DATE_KEY] = posting_date

            if not info["note_raw"]:
                info["note_raw"] = note_raw

            if not info["origin_date"]:
                info["origin_date"] = origin_date

    # ======================================================
    # BUILD JSON
    # ======================================================

    nested_summary = {}

    for (
        merchant_name,
        merchant_id,
        payment_type,
        recon_result,
        posting_date,
        note_group
    ), info in grouped_data.items():

        nested_summary.setdefault(
            merchant_name,
            {}
        )

        nested_summary[
            merchant_name
        ].setdefault(
            merchant_id,
            {}
        )

        nested_summary[
            merchant_name
        ][
            merchant_id
        ].setdefault(
            payment_type,
            {}
        )

        nested_summary[
            merchant_name
        ][
            merchant_id
        ][
            payment_type
        ].setdefault(
            posting_date,
            {}
        )

        nested_summary[
            merchant_name
        ][
            merchant_id
        ][
            payment_type
        ][
            posting_date
        ].setdefault(
            recon_result,
            []
        )

        nested_summary[
            merchant_name
        ][
            merchant_id
        ][
            payment_type
        ][
            posting_date
        ][
            recon_result
        ].append({
            "total_amount": info["total_amount"],
            "txn_count": info["txn_count"],
            POSTING_DATE_KEY: posting_date,
            CURRENCY_KEY: list(info["currencies"]),
            ORIGIN_PAYMENT_DATE_KEY: info["origin_date"],
            NOTE_KEY: info["note_raw"]
        })

    return nested_summary

def export_to_csv(json_file, output_csv):

    with open(json_file, "r", encoding="utf-8") as f:
        data = json.load(f)

    rows = []

    for file_name, merchants in data.items():
        for merchant_name, merchant_ids in merchants.items():
            for merchant_id, payment_types in merchant_ids.items():
                for payment_type, posting_date_dict in payment_types.items():
                    for posting_date, recon_results in posting_date_dict.items():
                        for recon_result, infos in recon_results.items():
                            for info in infos:
                                if recon_result != "":  # Exclude empty result rows
                                    rows.append({
                                        "File": file_name,
                                        POSTING_DATE_KEY: posting_date,
                                        MERCHANT_NAME_KEY: merchant_name,
                                        MERCHANT_ID_KEY: merchant_id,
                                        PAYMENT_TYPE_KEY: payment_type,
                                        RECON_RESULT_KEY: recon_result,
                                        "Tổng số tiền": format_money(info["total_amount"]),
                                        "Số lượng giao dịch": info["txn_count"],
                                        CURRENCY_KEY: ",".join(info.get(CURRENCY_KEY, [])),
                                        ORIGIN_PAYMENT_DATE_KEY: info.get(ORIGIN_PAYMENT_DATE_KEY, ""),
                                        NOTE_KEY: info.get(NOTE_KEY, "")
                                    })

    fieldnames = [
        "File",
        POSTING_DATE_KEY,
        MERCHANT_NAME_KEY,
        MERCHANT_ID_KEY,
        PAYMENT_TYPE_KEY,
        RECON_RESULT_KEY,
        "Tổng số tiền",
        "Số lượng giao dịch",
        CURRENCY_KEY,
        ORIGIN_PAYMENT_DATE_KEY,
        NOTE_KEY
    ]

    with open(output_csv, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main():
    print("cwd =", os.getcwd())
    print("__file__ =", __file__)
    print("resolved =", Path(__file__).resolve())
    with open(INPUT_FILE, "r", encoding="utf-8") as f:
        recon_data = json.load(f)

    final = {}

    for file_name, file_content in recon_data.items():
        final[file_name] = process_file_data(file_name, file_content.get("data", []))

    export_summary_to_json(final)
    export_to_csv(SUMMARY_JSON_FILE, SUMMARY_CSV_FILE)


if __name__ == "__main__":
    main()