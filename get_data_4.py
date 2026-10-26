#!/usr/bin/env python3

import calendar
from datetime import datetime, timedelta 
import json
import csv
import re
from collections import defaultdict
from pathlib import Path
import sys
import os
from convert_data_to_json import FILE_PREFIX, WALLET_FILE_PREFIX, WALLET_NEW_FILE_PREFIX, WALLET_NEW_FILE_PREFIX_INPUT, WALLET_NEW_FILE_PREFIX_OUTPUT
from export_upload_file_1 import get_date_range_from_filename, parse_posting_date_from_filename, extract_hdbank_merchant, extract_bvbank_merchant
from helper import normalize_posting_date

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
REFER_NUM_KEY = "Số tham chiếu"
POSTING_DATE_KEY = "Ngày giao dịch"
POSTING_DATE_NEW_WALLET_KEY = "Transaction Date"
GP_TRANS_DATE_KEY = "GpTransDate"
IS_CHARGE_KEY = "IsChargeFee"

# Reconciliation Status của file RECON_REPORT_WALLET_INPUT/OUTPUT -> recon_result.
# Phần lớn đã là mã chuẩn, một số dòng vẫn là text tiếng Việt như file DS_Daily cũ.
NEW_WALLET_RECON_STATUS = {
    "MATCH_SUCCESS": "MATCH_SUCCESS",
    "MISS_GPAY": "MISS_GPAY",
    "MISS_PARTNER": "MISS_PARTNER",
    "GPAY_MORE": "GPAY_MORE",
    "PARTNER_MORE": "PARTNER_MORE",
    "Khớp": "MATCH_SUCCESS",
    "Lệch (Galaxy Pay nhiều hơn)": "MISS_GPAY",
    "Lệch (Đối tác nhiều hơn)": "MISS_PARTNER",
    "Xử lý lệch (Galaxy Pay nhiều hơn)": "GPAY_MORE",
    "Xử lý lệch (Đối tác nhiều hơn)": "PARTNER_MORE",
}


def normalize_new_wallet_recon_status(recon_text, file_name):
    if recon_text not in NEW_WALLET_RECON_STATUS:
        raise ValueError(
            f"Unknown Reconciliation Status '{recon_text}' in file {file_name}"
        )

    return NEW_WALLET_RECON_STATUS[recon_text]


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
    refer_num = ""
    grouped_data = defaultdict(lambda: {
        "total_amount": 0,
        "txn_count": 0,
        "currencies": set(),
        "note_raw": "",
        "origin_date": "",
        POSTING_DATE_KEY: "",
        REFER_NUM_KEY: "",
    })

    is_recon_input = FILE_PREFIX in file_name
    is_daily_wallet = WALLET_FILE_PREFIX in file_name
    is_new_daily_wallet = WALLET_NEW_FILE_PREFIX in file_name
    is_new_daily_wallet_input = WALLET_NEW_FILE_PREFIX_INPUT in file_name
    is_new_daily_wallet_output = WALLET_NEW_FILE_PREFIX_OUTPUT in file_name
    is_HDBANK = "HDBANK_068704070165977" in file_name
    is_BVB = "BVBank_0697041105669" in file_name
    # ======================================================
    # HELPER: GET POSTING DATE
    # ======================================================

    def get_posting_date(row):
        pattern_days = r"\((\d{1,2})-(\d{1,2})\)"
        days_match = re.search(pattern_days, file_name)

        pattern_months = r"\((\d{1,2})\.(\d{1,2})-(\d{1,2})\.(\d{1,2})\)"
        month_match = re.search(pattern_months, file_name)

        posting_date = row.get(
            POSTING_DATE_KEY, row.get(
            POSTING_DATE_NEW_WALLET_KEY, ""
        )
        ).split(" ")[0]

        # Ví mới: GD đối tác không ghi nhận (MISS_PARTNER...) thì Transaction Date
        # trống -> lấy ngày phía Galaxy Pay
        if not posting_date and row.get(GP_TRANS_DATE_KEY):
            posting_date = normalize_posting_date(
                row[GP_TRANS_DATE_KEY].split(" ")[0]
            )

        if days_match:
            posting_date = parse_posting_date_from_filename(file_name)

        if month_match:
            posting_date = get_posting_date_from_range(
                posting_date,
                file_name
            )

        return posting_date, days_match
    def parse_amount(value):
            try:
                if value in (None, ""):
                    return 0
                return int(float(str(value).replace(",", "")))
            except:
                return 0
    def add_group(
            merchant_name,
            merchant_id,
            payment_type,
            posting_date,
            amount,
            note_raw,
            refer_num="",
            # Sao kê HDBANK / BVBANK không có kết quả đối soát -> MATCH_SUCCESS
            recon_result="MATCH_SUCCESS"
        ):
            if not posting_date or amount == 0:
                return

            note_group = note_raw
            
            group_key = (
                merchant_name,
                merchant_id,
                payment_type,
                recon_result,
                posting_date,
                note_group,
                refer_num
            )

            info = grouped_data[group_key]

            info["total_amount"] += amount
            info["txn_count"] += 1
            info["currencies"].add(currency)
            info[POSTING_DATE_KEY] = posting_date
            info[REFER_NUM_KEY] = refer_num

            if not info["note_raw"]:
                info["note_raw"] = note_raw
    # ======================================================
    # RECON_REPORT_INPUT
    # ======================================================

    if is_recon_input:
        mpgs_daily = defaultdict(int)
        mpgs_daily_info = {}
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
            refer_num = str(
                row.get(
                    REFER_NUM_KEY,
                    ""
                ) or ""
            ).strip()
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
                note_group,
                ""
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
            refer_num = str(
                row.get(
                    REFER_NUM_KEY,
                    ""
                ) or ""
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
            refer_num = str(
                row.get(
                    REFER_NUM_KEY,
                    ""
                ) or ""
            ).strip()
            group_key = (
                merchant_name,
                merchant_id,
                payment_type,
                recon_result,
                posting_date,
                note_group,
                refer_num
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
                
            info[REFER_NUM_KEY] = refer_num

        # ==================================================
        # 2. PAYMENT SIDE
        # ==================================================

        print("[INFO] Processing PAYMENT side")

        for row in rows:
            refer_num = str(
                row.get(
                    REFER_NUM_KEY,
                    ""
                ) or ""
            ).strip()
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
                note_group,
                refer_num
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
            info[REFER_NUM_KEY] = refer_num
    elif is_new_daily_wallet:

        # ==================================================
        # 1. BANK SIDE
        # ==================================================
        if is_new_daily_wallet_input:
            print("[INFO] Processing BANK side")

            for row in rows:

                bank_partner = row.get(
                    "Provider",
                    ""
                ).strip()

                # Không phải Bank row -> skip
                if not bank_partner:
                    continue

                posting_date, days_match = get_posting_date(row)

                note_raw = row.get(
                    "Note",
                    ""
                ).strip()

                currency = row.get(
                    "Currency",
                    ""
                ).strip()

                origin_date = row.get(
                    "GpTransDate",
                    ""
                ).strip()
                refer_num = str(
                    row.get(
                    REFER_NUM_KEY,
                    ""
                    ) or ""
                    ).strip()
                origin_date = ""
                merchant_name = bank_partner
                merchant_id = bank_partner

                raw_payment_type = row.get(
                    "TransType",
                    ""
                ).strip()

                recon_text = row.get(
                    "Reconciliation Status",
                    ""
                ).strip()
                # refer_num = str(
                #     row.get(
                #         "GPC TransactionCode",
                #         ""
                #     ) or ""
                # ).strip()
                # ==================================================
                # PAYMENT TYPE
                # ==================================================

                if "CASHIN" in raw_payment_type or raw_payment_type == "PAYMENT":
                    payment_type = "CASHIN"
                else:
                    payment_type = raw_payment_type

                # ==================================================
                # RECON RESULT
                # ==================================================

                recon_result = normalize_new_wallet_recon_status(
                    recon_text,
                    file_name
                )

                
                # ==================================================
                # NAPAS
                # ==================================================

                # if bank_partner.upper() == "NAPAS":

                #     if (
                #         recon_text == "Lệch (Đối tác nhiều hơn)"
                #         and payment_type == "CASHIN"
                #     ):
                #         recon_result = "MISS_PARTNER"

                #     elif (
                #         recon_text == "Xử lý lệch (Galaxy Pay nhiều hơn)"
                #         and payment_type == "CASHIN"
                #     ):
                #         recon_result = "GPAY_MORE"

                #     elif (
                #         recon_text == "Xử lý lệch (Đối tác nhiều hơn)"
                #     ):
                #         recon_result = "PARTNER_MORE"
                #         posting_date = parse_posting_date_from_filename(file_name)

                # ==================================================
                # TRANSFER
                # ==================================================

                if "TRANSFER" in payment_type:
                    recon_result = ""
                elif (
                    recon_text == "Xử lý lệch (Đối tác nhiều hơn)"
                ):
                    recon_result = "PARTNER_MORE"
                    posting_date = parse_posting_date_from_filename(file_name)
                # ==================================================
                # AMOUNT BANK
                # ==================================================

                try:
                    amount = int(float(
                        row.get(
                            "Trans Amount",
                            0
                        )
                    ))
                except:
                    amount = 0

                note_group = ""

                # ==================================================
                # GROUP BANK
                # ==================================================
                # refer_num = str(
                #     row.get(
                #         "GPC TransactionCode",
                #         ""
                #     ) or ""
                # ).strip()
                add_group(
                    merchant_name=merchant_name,
                    merchant_id=merchant_id,
                    payment_type=payment_type,
                    posting_date=posting_date,
                    amount=amount,
                    note_raw="",
                    refer_num = refer_num,
                    recon_result=recon_result
                )
                # group_key = (
                #     merchant_name,
                #     merchant_id,
                #     payment_type,
                #     recon_result,
                #     posting_date,
                #     note_group,
                #     refer_num
                # )

                # info = grouped_data[group_key]

                # info["total_amount"] += amount
                # info["txn_count"] += 1
                # info["currencies"].add(currency)
                # info[POSTING_DATE_KEY] = posting_date

                # if not info["note_raw"]:
                #     info["note_raw"] = note_raw

                # if not info["origin_date"]:
                #     info["origin_date"] = origin_date
                    
                # info[REFER_NUM_KEY] = refer_num

        # ==================================================
        # 2. PAYMENT SIDE
        # ==================================================
        if is_new_daily_wallet_output:
            print("[INFO] Processing PAYMENT side")

            for row in rows:
                # refer_num = str(
                #     row.get(
                #         "GPC TransactionCode",
                #         ""
                #     ) or ""
                # ).strip()
                payment_partner = row.get(
                    "Merchant Name",
                    ""
                ).strip()

                # Không phải Payment row -> skip
                if not payment_partner:
                    continue

                posting_date, days_match = get_posting_date(row)

                note_raw = row.get(
                    "Note",
                    ""
                ).strip()

                currency = row.get(
                    "Currency",
                    ""
                ).strip()

                origin_date = row.get(
                    "GpTransDate",
                    ""
                ).strip()
                refer_num = str(
                    row.get(
                    REFER_NUM_KEY,
                    ""
                    ) or ""
                    ).strip()
                origin_date = ""
                merchant_name = payment_partner

                merchant_id = row.get(
                    "Merchant Name",
                    ""
                ).strip()
                tran_service = row.get(
                    "Trans Service",
                    ""
                ).strip()
                if "IMEDIA" in merchant_id:
                    if "BILL" in tran_service:
                        merchant_name = "IMEDIABILL"
                        merchant_id = "IMEDIABILL"
                    elif "TOPUP" in tran_service:
                        merchant_name = "IMEDIATOPUP"
                        merchant_id = "IMEDIATOPUP"
                
                raw_payment_type = row.get(
                    "TransType",
                    ""
                ).strip()

                # ==================================================
                # PAYMENT TYPE
                # ==================================================

                if "CASHIN" in raw_payment_type:
                    payment_type = "PAYMENT"
                else:
                    payment_type = raw_payment_type

                recon_text = row.get(
                    "Reconciliation Status",
                    ""
                ).strip()

                # ==================================================
                # PARTNER MORE
                # ==================================================

                # if recon_text == "MISS_PARTNER":
                #     continue

                # ==================================================
                # RECON RESULT
                # ==================================================

                recon_result = normalize_new_wallet_recon_status(
                    recon_text,
                    file_name
                )
                # ==================================================
                # AMOUNT PAYMENT
                # ==================================================

                try:
                    amount = int(float(
                        row.get(
                            "Trans Amount",
                            0
                        )
                    ))
                except:
                    amount = 0

                # ==================================================
                # GROUP PAYMENT
                # ==================================================
                add_group(
                    merchant_name=merchant_name,
                    merchant_id=merchant_id,
                    payment_type=payment_type,
                    posting_date=posting_date,
                    amount=amount,
                    note_raw="",
                    refer_num = refer_num,
                    recon_result=recon_result
                )
                # note_group = ""
                # group_key = (
                #     merchant_name,
                #     merchant_id,
                #     payment_type,
                #     recon_result,
                #     posting_date,
                #     note_group,
                #     refer_num
                # )

                # info = grouped_data[group_key]

                # info["total_amount"] += amount
                # info["txn_count"] += 1
                # info["currencies"].add(currency)
                # info[POSTING_DATE_KEY] = posting_date

                # if not info["note_raw"]:
                #     info["note_raw"] = note_raw

                # if not info["origin_date"]:
                #     info["origin_date"] = origin_date
                # info[REFER_NUM_KEY] = refer_num
    
    # ======================================================
    # HDBANK_INPUT
    # ======================================================
    elif is_HDBANK:
        merchant_name = "HDBANK_STATEMENT"
        currency = "VND"
        def extract_merchant_id(payment_info):
                search_pattern =  r"BAO\s+CO\s+TID\s*:\s*(\d+)"

                match = re.search(
                    search_pattern,
                    payment_info,
                    re.IGNORECASE
                )

                return match.group(1) if match else payment_info
        # ========================================================
        # SORT HDBANK ROWS
        #
        # Thứ tự:
        #   1. Số tham chiếu tăng dần
        #   2. Trong cùng Số tham chiếu:
        #      - Phát sinh có trước
        #      - Phát sinh nợ sau
        # ========================================================

        for index, row in enumerate(rows):

            posting_date = row.get(
                POSTING_DATE_KEY,
                ""
            ).strip()
            refer_num = str(
                row.get(
                    REFER_NUM_KEY,
                    ""
                )
            ).strip()
            if not posting_date:
                print(
                    "[WARN] HDBANK: "
                    "Skip row because posting_date is empty"
                )
                continue

            payment_info = row.get(
                "Thông tin giao dịch",
                ""
            ).strip()
            
            
            
            
            # ========================================================
            # A. GALAXY PAY TRANSFER
            # ========================================================

            if ("Galaxy Pay chuyen tien" in payment_info ):

                merchant_id = extract_hdbank_merchant(
                    payment_info
                )

                amount = parse_amount(
                    row.get("Phát sinh nợ", 0)
                )

                if amount == 0:
                    continue

                add_group(
                    merchant_name=merchant_name,
                    merchant_id=merchant_id,
                    payment_type="TRANSFER",
                    posting_date=posting_date,
                    amount=amount,
                    note_raw=f"Phát sinh nợ {refer_num} {payment_info}",
                    refer_num = refer_num
                    
                )

                continue

            # ========================================================
            # B. FEE
            # ========================================================
            
            elif (
                "Loai giao dich: VAT-Merchant Charges-2015"
                in payment_info
            ):

                if index + 2 >= len(rows):
                    continue

                row_n1 = rows[index + 1]
                row_n2 = rows[index + 2]

                amount_n = parse_amount(
                    row.get("Phát sinh nợ", 0)
                )

                amount_n1 = parse_amount(
                    row_n1.get("Phát sinh nợ", 0)
                )

                amount_n2_credit = parse_amount(
                    row_n2.get("Phát sinh có", 0)
                )

                fee_amount = amount_n + amount_n1

                merchant_info_n2 = row_n2.get(
                    "Thông tin giao dịch",
                    ""
                ).strip()

                merchant_id = extract_merchant_id(
                    merchant_info_n2
                )

                add_group(
                    merchant_name=merchant_name,
                    merchant_id=merchant_id,
                    payment_type="FEE",
                    posting_date=posting_date,
                    amount=fee_amount,
                    note_raw=f"Phát sinh nợ {refer_num} {merchant_info_n2}",
                    refer_num = refer_num
                )

                # amount_n2_credit chưa tạo group
                # sẽ xử lý ở bước tiếp theo

                continue

            # ========================================================
            # C. CÁC BÁO CÓ
            # ========================================================
            elif "Loai giao dich: Merchant payment for on-u-1001" in payment_info:

                amount = parse_amount(
                    row.get("Phát sinh có", 0)
                )

                merchant_id = extract_merchant_id(
                    payment_info
                )

                add_group(
                    merchant_name=merchant_name,
                    merchant_id=merchant_id,
                    payment_type="PAYMENT_RECEIVED",
                    posting_date=posting_date,
                    amount=amount,
                    note_raw=f"Phát sinh có {payment_info}",
                    refer_num = refer_num
                )

                continue
            # MOMO
            elif "MOMO TT GALAXY PAY" in payment_info:

                amount = parse_amount(
                    row.get("Phát sinh có", 0)
                )

                add_group(
                    merchant_name="MOMO",
                    merchant_id="MOMO",
                    payment_type="PAYMENT_RECEIVED",
                    posting_date=posting_date,
                    amount=amount,
                    note_raw= f"Phát sinh có {refer_num}",
                    refer_num = refer_num
                )

                continue

            # NAPAS
            elif "NAPAS BAO CO GD THANH TOAN" in payment_info:

                amount = parse_amount(
                    row.get("Phát sinh có", 0)
                )

                add_group(
                    merchant_name="NAPAS",
                    merchant_id="NAPAS",
                    payment_type="PAYMENT_RECEIVED",
                    posting_date=posting_date,
                    amount=amount,
                    note_raw= f"Phát sinh có {refer_num}",
                    refer_num = refer_num
                )

                continue

            # ZALOPAY
            elif "TT GD QUA VI ZALOPAY" in payment_info:

                amount = parse_amount(
                    row.get("Phát sinh có", 0)
                )

                add_group(
                    merchant_name="ZALOPAY",
                    merchant_id="ZALOPAY",
                    payment_type="PAYMENT_RECEIVED",
                    posting_date=posting_date,
                    amount=amount,
                    note_raw= f"Phát sinh có {refer_num}",
                    refer_num = refer_num
                )

                continue

            # LÃI
            elif "TRA LAI TIEN GUI" in payment_info:

                amount = parse_amount(
                    row.get("Phát sinh có", 0)
                )

                add_group(
                    merchant_name=merchant_name,
                    merchant_id="GPay",
                    payment_type="DEPOSITS_INTEREST",
                    posting_date=posting_date,
                    amount=amount,
                    note_raw= f"Phát sinh có {refer_num} {payment_info}",
                    refer_num = refer_num
                )

                continue
            elif "vietjet air" in payment_info.lower():
                amount = parse_amount(
                    row.get("Phát sinh có", 0)
                )

                add_group(
                    merchant_name=merchant_name,
                    merchant_id="VIETJET",
                    payment_type="PAYMENT_RECEIVED",
                    posting_date=posting_date,
                    amount=amount,
                    note_raw=f"Phát sinh có {refer_num} {payment_info}",
                    refer_num = refer_num
                )
                continue
            elif "So giao dich" in payment_info:
                amount = parse_amount(
                    row.get("Phát sinh có", 0)
                )

                add_group(
                    merchant_name=merchant_name,
                    merchant_id="HDBANK",
                    payment_type="PAYMENT_RECEIVED",
                    posting_date=posting_date,
                    amount=amount,
                    note_raw=f"Phát sinh có GD-QR {posting_date}",
                    refer_num = ""
                )
                continue
            else:
                amount = parse_amount(
                    row.get("Phát sinh có", 0)
                )

                add_group(
                    merchant_name=merchant_name,
                    merchant_id="HDBANK",
                    payment_type="PAYMENT_INTERNAL",
                    posting_date=posting_date,
                    amount=amount,
                    note_raw= f"Phát sinh có {refer_num} {posting_date}  {payment_info}",
                    refer_num = refer_num
                )

                
                
    elif is_BVB:
        # ============================================================
        # PREPARE ROW DATA
        # ============================================================
        merchant_name = "BVBANK"
        currency = "VND"
        processed_rows = []

        for index, row in enumerate(rows):
            refer_num = str(
                row.get(
                    REFER_NUM_KEY,
                    ""
                ) or ""
            ).strip()
            posting_date = row.get(
                "NGÀY GIAO DỊCH\n(Transaction Date)",
                ""
            )

            posting_date = (
                str(posting_date).strip()
                if posting_date is not None
                else ""
            )

            payment_info = row.get(
                "DIỄN GIẢI\n(Narrative)",
                ""
            )

            payment_info = (
                str(payment_info).strip()
                if payment_info is not None
                else ""
            )

            debit_amount = parse_amount(
                row.get(
                    "PHÁT SINH NỢ\n(Debit Amount)",
                    0
                )
            )

            credit_amount = parse_amount(
                row.get(
                    "PHÁT SINH CÓ\n(Credit Amount)",
                    0
                )
            )

            processed_rows.append({
                "index": index,
                "row": row,
                "posting_date": posting_date,
                "payment_info": payment_info,
                "debit_amount": debit_amount,
                "credit_amount": credit_amount,
                "refer_num": refer_num
            })

        # ============================================================
        # 1. PHÁT SINH NỢ
        # ============================================================

        # Index các dòng phí đã được hạch toán cùng dòng chuyển tiền
        consumed_fee_indexes = set()

        for item in processed_rows:

            if item["index"] in consumed_fee_indexes:
                continue

            posting_date = item["posting_date"]
            payment_info = item["payment_info"]
            debit_amount = item["debit_amount"]
            refer_num = item["refer_num"]

            if not posting_date:
                print(
                    "[WARN] BVBANK: "
                    "Skip debit row because posting_date is empty"
                )
                continue

            if debit_amount == 0:
                continue

            # --------------------------------------------------------
            # 1.1 GALAXY PAY CHUYEN TIEN
            #
            # Nếu diễn giải giống nhau:
            #   dòng trên = chuyển tiền
            #   dòng dưới = phí
            #
            # Mỗi dòng hạch toán riêng
            # --------------------------------------------------------

            if "Galaxy Pay chuyen tien" in payment_info:

                # ----------------------------------------------------
                # Kiểm tra dòng tiếp theo
                # ----------------------------------------------------

                current_index = item["index"]

                next_item = None

                if current_index + 1 < len(processed_rows):
                    next_item = processed_rows[
                        current_index + 1
                    ]

                payment_type = "TRANSFER"
                merchant_id = extract_bvbank_merchant(payment_info)

                add_group(
                    merchant_name=merchant_name,
                    merchant_id=merchant_id,
                    payment_type=payment_type,
                    posting_date=posting_date,
                    amount=debit_amount,
                    note_raw=f"Phát sinh nợ {payment_info}",
                    refer_num = refer_num
                )

                # Nếu dòng tiếp theo có cùng diễn giải -> dòng phí
                if (
                    next_item
                    and next_item["payment_info"] == payment_info
                ):
                    add_group(
                        merchant_name=merchant_name,
                        merchant_id="GPay",
                        payment_type="FEE",
                        posting_date=posting_date,
                        amount=next_item["debit_amount"],
                        note_raw=f"Phát sinh nợ {payment_info}",
                        refer_num = next_item["refer_num"]
                    )
                    consumed_fee_indexes.add(next_item["index"])

                continue

            # --------------------------------------------------------
            # 1.2 GALAXY PAY YEU CAU HOAN TRA
            # --------------------------------------------------------

            if "GALAXY PAY YEU CAU HOAN TRA" in payment_info:

                add_group(
                    merchant_name=merchant_name,
                    merchant_id="GPay",
                    payment_type="REFUND",
                    posting_date=posting_date,
                    amount=debit_amount,
                    note_raw="Phát sinh nợ GALAXY PAY YEU CAU HOAN TRA",
                    refer_num = refer_num
                )

                continue
        
        # ============================================================
        # 2. PHÁT SINH CÓ
        # ============================================================

        # ------------------------------------------------------------
        # 2.1 LÃI TIỀN GỬI
        #
        # Cộng toàn bộ "TRA LAI TIEN GUI"
        # trong toàn bộ file thành 1 group
        # ------------------------------------------------------------

        total_interest = 0

        interest_posting_date = ""

        for item in processed_rows:

            payment_info = item["payment_info"]
            credit_amount = item["credit_amount"]
            posting_date = item["posting_date"]

            if "TRA LAI TIEN GUI" in payment_info:

                total_interest += credit_amount

                if not interest_posting_date and posting_date:
                    interest_posting_date = posting_date

        if total_interest > 0:

            add_group(
                merchant_name=merchant_name,
                merchant_id="GPay",
                payment_type="PAYMENT_INTEREST",
                posting_date=interest_posting_date,
                amount=total_interest,
                note_raw="Phát sinh có TRA LAI TIEN GUI",
                # Gom nhiều dòng -> không có 1 số tham chiếu duy nhất
                refer_num = ""
            )

        # ============================================================
        # 2.2 PHẦN CÒN LẠI
        # ============================================================

        # Loại các dòng lãi
        credit_rows = [
            item
            for item in processed_rows
            if (
                item["credit_amount"] != 0
                and "TRA LAI TIEN GUI" not in item["payment_info"]
            )
        ]

        # ============================================================
        # TÌM CÁC HEADER 23:00 - 23:59
        #
        # Header 23:00 - 23:59 là GIAO DỊCH CUỐI NGÀY
        #
        # Ví dụ:
        #
        # 01/06 transaction A
        # 01/06 transaction B
        # 01/06 23:00 - 23:59   <-- kết thúc ngày 01/06
        #
        # 02/06 transaction C
        # 02/06 transaction D
        # 02/06 23:00 - 23:59   <-- kết thúc ngày 02/06
        #
        # => Amount của mỗi ngày được cộng từ:
        #    previous header + 1
        #    đến chính header hiện tại
        # ============================================================

        daily_amounts = defaultdict(int)

        current_amount = 0
        current_date = None

        for item in credit_rows:

            payment_info = item["payment_info"]
            credit_amount = item["credit_amount"]

            # ========================================================
            # HEADER 23:00 - 23:59
            #
            # Header này nói rõ:
            # "ngay 31/05/2026"
            #
            # => amount trước đó thuộc về 31/05
            # ========================================================

            if "Giao dich tu 23:00:00 den 23:59:59" in payment_info:

                match = re.search(
                    r"ngay\s+(\d{2}/\d{2}/\d{4})",
                    payment_info
                )

                if match:

                    header_date = datetime.strptime(
                        match.group(1),
                        "%d/%m/%Y"
                    ).date()

                    current_amount += credit_amount

                    if current_amount != 0:
                        daily_amounts[header_date] += current_amount

                    current_amount = 0
                    current_date = None

                continue

            # ========================================================
            # Transaction bình thường
            # ========================================================

            posting_date = item["posting_date"]

            if not posting_date:
                continue

            try:
                item_date = datetime.strptime(
                    posting_date,
                    "%d/%m/%Y"
                ).date()

            except ValueError:
                continue

            # Nếu bắt đầu một ngày mới
            if current_date is None:
                current_date = item_date

            # Nếu ngày thay đổi mà chưa gặp header
            elif item_date != current_date:

                if current_amount != 0:
                    daily_amounts[current_date] += current_amount

                current_date = item_date
                current_amount = 0

            current_amount += credit_amount


        # ============================================================
        # PHẦN CUỐI FILE
        # ============================================================

        if current_date is not None and current_amount != 0:
            daily_amounts[current_date] += current_amount

        # ============================================================
        # 2.3 XÁC ĐỊNH CÁCH GOM THEO THÁNG
        # ============================================================

        monthly_amounts = defaultdict(int)

        daily_items = {}

        for date_obj, amount in daily_amounts.items():

            month_key = (
                date_obj.year,
                date_obj.month
            )

            daily_items[date_obj] = amount

            monthly_amounts[month_key] += amount

        # ============================================================
        # GROUP WEEKEND
        #
        # Thứ 6 + Thứ 7 + Chủ nhật
        # => Gom thành 1 dòng, posting_date = Chủ nhật
        #
        # Nếu weekend crossing month:
        #
        # 31/07 (Fri)
        # 01/08 (Sat)
        # 02/08 (Sun)
        #
        # => Không gom chung
        # ============================================================

        processed_dates = set()

        sorted_dates = sorted(daily_items.keys())

        for friday in sorted_dates:

            # Chỉ bắt đầu từ Thứ 6
            if friday.weekday() != 4:
                continue

            saturday = friday + timedelta(days=1)
            sunday = friday + timedelta(days=2)

            # ========================================================
            # Phải có đủ Fri + Sat + Sun
            # ========================================================

            if (
                saturday not in daily_items
                or sunday not in daily_items
            ):
                continue

            # ========================================================
            # Không cho phép weekend crossing month
            # ========================================================

            if not (
                friday.year == saturday.year == sunday.year
                and friday.month == saturday.month == sunday.month
            ):
                continue

            # ========================================================
            # Gom Fri + Sat + Sun
            # Posting date = Sunday
            # ========================================================

            weekend_amount = (
                daily_items[friday]
                + daily_items[saturday]
                + daily_items[sunday]
            )

            if weekend_amount == 0:
                continue

            add_group(
                merchant_name=merchant_name,
                merchant_id="GPay",
                payment_type="PAYMENT",
                posting_date=sunday.strftime("%d/%m/%Y"),
                amount=weekend_amount,
                note_raw="Phát sinh có BVBANK WEEKEND",
                refer_num = ""
            )

            # Đánh dấu đã xử lý
            processed_dates.update([
                friday,
                saturday,
                sunday
            ])


        # ============================================================
        # CÁC NGÀY CÒN LẠI
        # ============================================================

        for date_obj in sorted(daily_items.keys()):

            if date_obj in processed_dates:
                continue

            amount = daily_items[date_obj]

            if amount == 0:
                continue

            add_group(
                merchant_name=merchant_name,
                merchant_id="GPay",
                payment_type="PAYMENT",
                posting_date=date_obj.strftime("%d/%m/%Y"),
                amount=amount,
                note_raw="Phát sinh có BVBANK PAYMENT",
                refer_num = ""
            )
    #     # ============================================================
    #     # 2.4 HẠCH TOÁN
    #     # ============================================================

    #     processed_dates = set()

    #     for friday in sorted(daily_items.keys()):

    #         if friday.weekday() != 4:
    #             continue

    #         saturday = friday + timedelta(days=1)
    #         sunday = friday + timedelta(days=2)

    #         # Phải đủ Fri + Sat + Sun
    #         if (
    #             saturday not in daily_items
    #             or sunday not in daily_items
    #         ):
    #             continue

    #         # Không cho phép crossing month
    #         if not (
    #             friday.year == saturday.year == sunday.year
    #             and friday.month == saturday.month == sunday.month
    #         ):
    #             continue

    #         weekend_amount = (
    #             daily_items[friday]
    #             + daily_items[saturday]
    #             + daily_items[sunday]
    #         )

    #         add_group(
    #             merchant_name=merchant_name,
    #             merchant_id="GPay",
    #             payment_type="PAYMENT",
    #             posting_date=sunday.strftime("%d/%m/%Y"),
    #             amount=weekend_amount,
    #             note_raw="BVBANK WEEKEND"
    #         )

    #         processed_dates.update([
    #     friday,
    #     saturday,
    #     sunday
    # ])
    #     # ------------------------------------------------------------
    #     # Các ngày còn lại
    #     # ------------------------------------------------------------

    #     for date_obj in sorted(daily_items.keys()):

    #         if date_obj in processed_dates:
    #             continue

    #         amount = daily_items[date_obj]

    #         add_group(
    #             merchant_name=merchant_name,
    #             merchant_id="GPay",
    #             payment_type="PAYMENT",
    #             posting_date=date_obj.strftime("%d/%m/%Y"),
    #             amount=amount,
    #             note_raw="BVBANK PAYMENT"
    #         )
            
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
        note_group,
        refer_num
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
            NOTE_KEY: info["note_raw"],
            REFER_NUM_KEY: info[REFER_NUM_KEY]
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
                                        NOTE_KEY: info.get(NOTE_KEY, ""),
                                        REFER_NUM_KEY: info.get(
                                            REFER_NUM_KEY,
                                            ""
                                        )
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
        NOTE_KEY,
        REFER_NUM_KEY
    ]
    
    def parse_csv_date(value):
        try:
            return datetime.strptime(
                str(value).strip(),
                "%d/%m/%Y"
            )
        except (ValueError, TypeError):
            return datetime.max


    def get_flow_order(row):
        note = str(
            row.get(NOTE_KEY, "") or ""
        ).strip().lower()

        if note.startswith("phát sinh có"):
            return 0

        if note.startswith("phát sinh nợ"):
            return 1

        return 2


    def get_reference_sort_key(row):
        reference = str(
            row.get(REFER_NUM_KEY, "") or ""
        ).strip()

        if reference.isdigit():
            return int(reference)

        return float("inf")
    print("\n========== BEFORE SORT ==========")

    for row in rows:
            print(
                row.get("Ngày chứng từ"),
                "|",
                row.get("Ghi chú"),
                "|",
                repr(row.get("Số tham chiếu"))
            )

    print("================================\n")

    rows.sort(
        key=lambda row: (
            row.get("File", ""),
            parse_csv_date(
                row.get(POSTING_DATE_KEY, "")
            ),
            get_flow_order(row),
            get_reference_sort_key(row),
        )
    )
    print("\n========== AFTER SORT ==========")

    for row in rows:
            print(
                row.get("Ngày chứng từ"),
                "|",
                row.get("Ghi chú"),
                "|",
                repr(row.get("Số tham chiếu"))
            )

    print("================================\n")
    
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