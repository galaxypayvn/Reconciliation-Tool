#!/usr/bin/env python3

import csv
from datetime import datetime, time, timedelta
import re
import calendar
from pathlib import Path

# from trio import Path
INPUT_FILE = "csv_recon_report/recon_summary.csv"
OUTPUT_FILE = "csv_recon_report/upload_file.csv"


DEFAULT_CURRENCY = "VND"
DEFAULT_PROFIT_CENTER = "10800001"
from openpyxl import Workbook
from openpyxl.styles import numbers

def get_date_range_from_months_match(posting_date, file_name):
    """
    Lấy khoảng ngày hạch toán từ filename dạng:
        29.07-03.08
        31.07-02.08

    posting_date:
        dd/mm/yyyy

    Ví dụ:

        file_name = "...29.07-03.08..."
        posting_date = "29/07/2026"
        -> "29-31/07/2026"

        posting_date = "01/08/2026"
        -> "01-03/08/2026"

        file_name = "...31.07-02.08..."
        posting_date = "31/07/2026"
        -> "31/07/2026"

        posting_date = "01/08/2026"
        -> "01-02/08/2026"
    """

    if not posting_date or not file_name:
        return posting_date

    # ==========================================
    # Lấy range từ filename
    # Ví dụ: 29.07-03.08
    # ==========================================
    pattern = r"\((\d{1,2})\.(\d{1,2})-(\d{1,2})\.(\d{1,2})\)"
    match = re.search(pattern, file_name)

    if not match:
        return posting_date

    start_day = int(match.group(1))
    start_month = int(match.group(2))
    end_day = int(match.group(3))
    end_month = int(match.group(4))

    # ==========================================
    # Parse posting_date
    # ==========================================
    try:
        posting_dt = datetime.strptime(
            posting_date.strip(),
            "%d/%m/%Y"
        )
    except ValueError:
        return posting_date

    year = posting_dt.year
    current_month = posting_dt.month

    # ==========================================
    # THÁNG START
    # ==========================================
    if current_month == start_month:

        last_day = calendar.monthrange(
            year,
            start_month
        )[1]

        # Ví dụ:
        # 29.07-03.08
        # -> 29-31/07/2026
        if start_day != last_day:
            return f"{start_day:02d}-{last_day:02d}/{start_month:02d}/{year}"

        # Ví dụ:
        # 31.07-02.08
        # -> 31/07/2026
        return f"{start_day:02d}/{start_month:02d}/{year}"

    # ==========================================
    # THÁNG END
    # ==========================================
    if current_month == end_month:

        # Nếu start_month == end_month thì không phải
        # trường hợp chuyển tháng
        if start_month == end_month:
            return posting_date

        # Ví dụ:
        # 29.07-03.08
        # -> 01-03/08/2026
        if end_day != 1:
            return f"01-{end_day:02d}/{end_month:02d}/{year}"

        # Ví dụ range kết thúc ngay ngày 01
        # -> 01/08/2026
        return f"{end_day:02d}/{end_month:02d}/{year}"

    # Không thuộc range
    return posting_date

def csv_to_xlsx(input_file, output_file=None):
    input_file = Path(input_file)

    if not input_file.exists():
        print(f"[ERROR] File not found: {input_file}")
        return

    if output_file is None:
        output_file = input_file.with_suffix(".xlsx")
    else:
        output_file = Path(output_file)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Sheet1"

    # Đọc CSV
    with open(
        input_file,
        "r",
        encoding="utf-8-sig",
        newline=""
    ) as f:

        reader = csv.reader(f)

        for row_index, row in enumerate(reader, start=1):

            for col_index, value in enumerate(row, start=1):

                cell = worksheet.cell(
                    row=row_index,
                    column=col_index
                )

                # Ép toàn bộ dữ liệu thành TEXT
                cell.value = str(value)

                # Excel Text format
                cell.number_format = "@"

    workbook.save(output_file)

    print(f"[INFO] Converted:")
    print(f"       CSV : {input_file}")
    print(f"       XLSX: {output_file}")
    
def normalize_posting_date(posting_date):
    posting_date = posting_date.strip()

    for date_format in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(
                posting_date,
                date_format
            ).strftime("%d/%m/%Y")
        except ValueError:
            continue

    return posting_date

def mapping_merchant(merchant, merchant_id):
    merchant_name_map = {
        "VCCB": "BVB",
        "MOMO": "Momo",
        "ZALOPAY": "ZaloPay",
        "NAPAS": "Napas",
        "HDB_MPGS": "MPGS",
        "GPAY": "GPay",
        "IMEDIA": "Imedia",
        "BIDV": "BIDV",
        "HDBANK": "HDB",
        "HDB-VA": "HDB",
        "IMEDIA JSC_TOPUP": "Imedia",
        "IMediaBill": "Imedia",
        "IMediaTopup": "Imedia",
        "SKY JOY": "GPay",
        "SKY FI": "GPay",
        "SKYFI": "GPay",
        "VIETJET AIR": "GPay",
        "VNPAY": "VNPay"
  
    }

    merchant_name = merchant_name_map.get(
        merchant,
        merchant.title()
    )

    # ======================================================
    # Merchant mapping
    # ======================================================

    business_mapping = {

        "VJ APP TICKET": "VIETJET",
        "VJ WEB TICKET": "VIETJET",
        "VIETJET WEBSITE": "VIETJET",
        "VIETJET MOBILE": "VIETJET",
        "VJ APP LATER": "VIETJET",
        "VJ WEB LATER": "VIETJET",
        "VJ PAYMENT LINK": "VIETJET",
        "VIETJET AIR": "VIETJET",
        "SKYFI": "SKYFI",
        "SKY JOY": "GJoy",
        "ONAIR": "OnAir",
        "VJ WEB SKYSHOP": "SKYSHOP",
        "VJ KIOSK": "KIOSK",
        "RONG VIET TRAVEL": "Rồng Việt",

        "BAO HIEM HD": "HDI",
        "BAO HIEM HD - PORTAL": "HDI",
        "CÔNG TY TNHH BẢO HIỂM HD": "HDI",

        "VICTORIA SCHOOL": "VSSS",
        "VICTORIA SCHOOL RIVERSIDE": "KHoan-Riverside",
        "VICTORIA SAIGON SOUTH": "Khải Hoàn-NSG",

        "VINCLUB": "VinApp",
        "MENAS STORE": "Menas",
        
    }

    business = business_mapping.get(
        merchant_id,
        merchant_id
    )
    return {"merchant_map": merchant_name, "business_map": business}

def get_date_range_from_filename(file_name: str) -> str:
    """
    VD:
    DS_Daily_gui FA_20260706-Vi (3-5)_Chi_tiet_Vi.csv
        -> 03-05/07/2026

    DS_Daily_gui FA_20260706-Vi (9-10)_Chi_tiet_Vi.csv
        -> 09-10/07/2026
    """

    # Lấy tháng/năm từ ngày trong tên file (20260706)
    date_match = re.search(r"(\d{8})", file_name)
    if not date_match:
        return ""

    base_date = datetime.strptime(date_match.group(1), "%Y%m%d")

    # Lấy khoảng ngày (3-5)
    pattern = r"\((\d{1,2})-(\d{1,2})\)"
    days_match = re.search(pattern, file_name)

    if not days_match:
        return ""

    start_day = int(days_match.group(1))
    end_day = int(days_match.group(2))

    return f"{start_day:02d}-{end_day:02d}/{base_date.month:02d}/{base_date.year}"

def parse_posting_date_from_filename(file_name):
    date_match = re.search(r"(\d{8})", file_name)
    day_match = re.search(r"\((?:\d{1,2}-)?(\d{1,2})\)", file_name)

    if not date_match or not day_match:
        return ""

    try:
        file_date = datetime.strptime(date_match.group(1), "%Y%m%d")

        posting_date = datetime(
            year=file_date.year,
            month=file_date.month,
            day=int(day_match.group(1))
        )

        return posting_date.strftime("%d/%m/%Y")

    except ValueError:
        return ""

def build_posting_info(title, dr_cr, merchant, merchant_id, note):

    title_upper = title.upper()
    merchant = merchant.strip().upper()
    merchant_id = merchant_id.strip().upper()
    merchant_mapping = {
        "VIETJET AIR": "VIETJET",
        "VJ APP TICKET": "VIETJET",
        "VIETJET WEBSITE": "VIETJET",
        "VIETJET MOBILE": "VIETJET",
        "VJ WEB TICKET": "VIETJET",
        "VJ APP LATER": "VIETJET",
        "VJ APP SKYPOS": "VIETJET",
        "VJ PAYMENT LINK": "VIETJET",
        "VJ WEB LATER": "VIETJET",
        "VJ WEB SKYSHOP": "VIETJET",
        "VJ KIOSK": "VIETJET",
        "SKYFI": "SKYFI",
        "SKY JOY": "GJoy",
        "GALAXY PAY": "GPAY",
        "IMEDIA JSC_TOPUP": "IMEDIA",
        "IMEDIABILL": "IMEDIA",
        "IMEDIATOPUP": "IMEDIA",
        "BAO HIEM HD": "HDI",
        "BAO HIEM HD - PORTAL": "HDI",
        "CÔNG TY TNHH BẢO HIỂM HD": "HDI",

        "ONAIR": "OnAir",
        "RONG VIET TRAVEL": "Rồng Việt",

        "VICTORIA SCHOOL": "VSSS",
        "VICTORIA SCHOOL RIVERSIDE": "KHoan-Riverside",
        "VICTORIA SAIGON SOUTH": "Khải Hoàn-NSG",
        "SO BAN HANG": "SBH",
        "VINCLUB": "VinApp",
        "NAPAS" : "GPay",
        "MENAS STORE": "Menas"
    }
    merchant_name_map = {
        "VCCB": "BVB",
        "MOMO": "Momo",
        "ZALOPAY": "ZaloPay",
        "NAPAS": "Napas",
        "HDB_MPGS": "MPGS",
        "GPAY": "GPay",
        "BIDV": "BIDV",
        "HDBANK": "HDB",
        "HDB-VA": "HDB",
        "IMEDIA JSC_TOPUP": "GPay",
        "CÔNG TY TNHH BẢO HIỂM HD": "GPay",
        "IMEDIABILL": "GPay",
        "IMEDIATOPUP": "GPay",
        "SKY JOY": "GPay",
        "SKY FI": "GPay",
        "SKYFI": "GPay",
        "VIETJET AIR": "GPay",
        "VNPAY": "VNPay"
  
    }

    merchant_name = merchant_name_map.get(
        merchant,
        merchant.title()
    ).upper()

    business = merchant_mapping.get(merchant_id, merchant_id).upper()
    # business_mapping = {
    #     "VIETJET-SKYSHOP": "VIETJET",
    #     "VJ KIOSK": "VIETJET",
    #     "VJ APP TICKET": "VIETJET",
    #     "VJ WEB TICKET": "VIETJET",
    #     "VIETJET WEBSITE": "VIETJET",
    #     "VIETJET MOBILE": "VIETJET",
    #     "VJ APP LATER": "VIETJET",
    #     "VJ WEB LATER": "VIETJET",
    #     "VJ PAYMENT LINK": "VIETJET",
    #     "VIETJET AIR": "VIETJET",
    #     "SKYFI": "SKYFI",
    #     "SKY JOY": "GJoy",
    #     "ONAIR": "OnAir",

    #     "RONG VIET TRAVEL": "Rồng Việt",

    #     "BAO HIEM HD": "HDI",
    #     "BAO HIEM HD - PORTAL": "HDI",

    #     "VICTORIA SCHOOL": "VSSS",
    #     "VICTORIA SCHOOL RIVERSIDE": "Khải Hoàn-NSG",
    #     "VICTORIA SAIGON SOUTH": "KHoan-Riverside",

    #     "VINCLUB": "VinApp",
    #     "MENAS STORE": "Menas",
        
    # }

    # business = business_mapping.get(
    #     merchant_id,
    #     merchant_id
    # ).upper()
    # ======================================================
    # DEFAULT
    # ======================================================

    # document_type = "DK"

    if dr_cr == "Nợ" and "HLệch" not in title:
        account = "1388404200"
    else:
        account = "3388403100"

    # ======================================================
    # BUSINESS UNIT MAPPING
    # ======================================================

    BUSINESS_UNIT = {
        # Merchant (Partner)
        # "VIETJET AIR": "0037360001",
        "VIETJET": "1010",
        "SKYFI": "1000003120",
        "GJOY": "0037888888",
        "SKY JOY": "0037888888",
        "SKYJOY": "0037888888",
        "RỒNG VIỆT": "0037486431",
        "RONG VIET": "0037486431",
        "HDI": "1000000860",
        "ONAIR": "1000002978",
        "VSSS": "1000002177",
        "KHẢI HOÀN-NSG": "1000003029",
        "KHOAN-RIVERSIDE": "1000003028",
        "SBH": "1000002630",
        "VINAPP": "1000002977",
        "MENAS": "1000000265",

        # Đối tác
        "MOMO": "0037385423",
        "ZALOPAY": "0037482684",
        "NAPAS": "1000002622",
        "BVB": "1000002626",
        "VCCB": "1000002626",
        "VNPAY": "0037405944",
        "IMEDIA": "1000002623",
        "GPAY": "0037360001",
        "GALAXY PAY": "0037360001",
        "HDB": "0037360001",
        "HDB-VA": "0037360001",
        "HDBANK": "0037360001",
        "BIDV": "0037360001",
        "MPGS": "0037360001",
        "GIAC MO BAY RE": "0037483314"
    }
    
    ACCOUNT_MAPPING = {
        "VIETJET": "0037360001",
        "SKYFI": "0037360001",
        "GJOY": "3388401000",
        "RỒNG VIỆT": "0037360001",
        "HDI": "0037360001",
        "ONAIR": "0037360001",
        "VSSS": "0037360001",
        "KHẢI HOÀN-NSG": "0037360001",
        "KHOAN-RIVERSIDE": "0037360001",
        "SBH": "0037360001",
        "VINAPP": "0037360001",
        "MENAS": "0037360001",

        # Đối tác
        "MOMO": "0037385423",
        "ZALOPAY": "0037482684",
        "NAPAS": "1388404300",
        "BVB": "1000002626",
        "VCCB": "1000002626",
        "VNPAY": "0037405944",
        "IMEDIA": "1000002623",
        "GPAY": "0037360001",
        "HDB": "0037360001",
        "BIDV": "0037360001",
    }

    partner_bu = BUSINESS_UNIT.get(merchant_name, merchant_name)
    merchant_bu = BUSINESS_UNIT.get(business, business)
    partner_acc = ACCOUNT_MAPPING.get(merchant_name, merchant_name)
    merchant_acc = ACCOUNT_MAPPING.get(business, business)
    # ======================================================
    # BUSINESS UNIT
    # Không Hoàn:
    #   Nợ -> Đối tác
    #   Có  -> Merchant
    #
    # Hoàn:
    #   Nợ -> Merchant
    #   Có  -> Đối tác
    # ======================================================

    if "HOÀN" in title_upper and "KHẢI HOÀN" not in title_upper:
        business_unit = merchant_bu if dr_cr == "Nợ" else partner_bu
        # account = merchant_acc if dr_cr == "Nợ" else partner_acc
    else:
        business_unit = partner_bu if dr_cr == "Nợ" else merchant_bu
        # account = partner_acc if dr_cr == "Nợ" else merchant_acc
    if "HLỆCH" in title_upper:
        business_unit = partner_bu
    if "MPGS" in title_upper and "HOÀN" in title_upper:
        
        if dr_cr == "Nợ":
            account = "3388403100"
        else:
            account = "1388404200"
    # ======================================================
    # MOMO, ZALO
    # ======================================================

    

    if ("HOÀN" in title_upper and "KHẢI HOÀN" not in title_upper) or "HLỆCH" in title_upper:
            
        if dr_cr == "Nợ":
            if "HLỆCH" in title_upper:
                account = "3388403200"
            else:
                account = "3388403100"
        else:
            account = "1388404200"
            
    elif "LỆCH" in title_upper:

        if dr_cr == "Nợ":
            account = "1388404200"
        else:
            account = "3388403200"
            business_unit = partner_bu

 
    # ======================================================
    # BVB
    # ======================================================

    if "BVB" in title_upper:

        if "IBFT" in title_upper:
            if note == "MISS_REFUND_FALSE":
                account = "3388403200" if dr_cr == "Nợ" else "3388403100"
                business_unit = partner_bu
            else:
                account = "3388403100"

    # ======================================================
    # NAPAS
    # ======================================================

    elif "NAPAS" in title_upper and "KHẢI HOÀN" not in title_upper:
        account = "1388404300"
            
        if "HOÀN" in title_upper:

            if dr_cr == "Nợ":
                account = "3388403100"
            else:
                account = "1388404300"
        
        elif "NẠP" in title_upper:
            if "XỬ LÝ" in title_upper or "XLLỆCH" in title_upper:

                # document_type = "KZ"
                business_unit = "0037360001"
                if dr_cr == "Nợ":
                    account = "3388402000"
                else:
                    account = "3388401000"
            elif "HLỆCH" in title_upper or "Hoàn LỆCH" in title_upper:
                business_unit = merchant_bu if dr_cr == "Nợ" else partner_bu
                if dr_cr == "Nợ":
                    account = "3388402000"
                else:
                    account = "1388404300"       
            elif "LỆCH NẠP" in title_upper:

                if dr_cr == "Nợ":
                    account = "1388404300"
                else:
                    account = "3388402000"
                    business_unit = merchant_bu                    

            else:

                if dr_cr == "Nợ":
                    account = "1388404300"
                else:
                    account = "3388401000"
        else:
            if dr_cr == "Nợ":
                account = "1388404300"
            else:
                account = "3388403100"
    # ======================================================
    # GPAY
    # ======================================================

    elif "GPAY" in title_upper:

        # document_type = "CL"

        if "HOÀN" in title_upper:

            if dr_cr == "Nợ":
                account = "3388403100"
            else:
                account = "3388401000"

        else:

            if dr_cr == "Nợ":
                account = "3388401000"
                # business_unit = merchant_bu
            else:
                account = "3388403100"

    # ======================================================
    # IMEDIA
    # ======================================================

    elif "IMEDIA" in title_upper:

        if "HLỆCH" in title_upper:

            # document_type = "KZ"
            business_unit = "0037360001"
            if dr_cr == "Nợ":
                account = "3388402000"
            else:
                account = "3388401000"
        elif "LỆCH" in title_upper:
            business_unit = partner_bu
            # document_type = "KZ"

            if dr_cr == "Nợ":
                account = "3388401000"
            else:
                account = "3388402000"
        else:

            if dr_cr == "Nợ":
                account = "3388401000"
            else:
                account = "1388404000"

    # ======================================================
    # HDB / BIDV
    # ======================================================
    # "HDB" in title_upper or
    elif "HDB" in title_upper or "BIDV" in title_upper:

        # document_type = "KZ"

        if "XỬ LÝ" in title_upper or "XLÝ" in title_upper:

            if dr_cr == "Nợ":
                account = "3388402000"
            else:
                account = "3388401000"
        elif "GD QR" in title_upper and "LỆCH" not in title_upper and "Hoàn" not in title:
            if dr_cr == "Nợ":
                account = "1388404200"
            else:
                account = "3388403100"
        elif "TEST" in title_upper:
            if dr_cr == "Nợ":
                account = "1388404200"
            else:
                account = "3388403300"
                business_unit = "1080"    
        elif "HDB-Lệch GD QR" in title :
            if dr_cr == "Nợ":
                account = "1388404200"
            else:
                account = "3388403200"
        elif "Hoàn" not in title:

            if dr_cr == "Nợ":
                account = "3388401000"
            else:
                account = "3388402000"

    # ======================================================
    # VNPAY
    # ======================================================

    elif "VNPAY" in title_upper:

        # document_type = "KZ"

        if dr_cr == "Nợ":
            account = "3388401000"
        else:
            account = "3388403100"

    # ======================================================
    # ĐÓNG TK VÍ
    # ======================================================

    elif "ĐÓNG TK VÍ" in title_upper:

        # document_type = "KZ"

        if dr_cr == "Nợ":
            account = "3388401000"
        else:
            account = "7110990000"

        business_unit = ""
    
    return {
        # "document_type": document_type,
        "account": account,
        "business_unit": business_unit
    }

def parse_date_from_filename(file_name: str):
    match = re.search(r"\d{8}", file_name)
    if not match:
        return ""

    try:
        dt = datetime.strptime(match.group(), "%Y%m%d")
        dt = dt - timedelta(days=1)
        return dt.strftime("%d/%m/%Y")
    except:
        return ""

def build_title_wallet(
        merchant,
        payment_type,
        merchant_id,
        recon_result,
        note
):
    merchant = merchant.strip().upper()
    payment_type = payment_type.strip().upper()
    recon_result = recon_result.strip().upper()
    note_lower = note.lower().strip()

    # ======================================================
    # Merchant name
    # ======================================================

    merchant_name_map = {
        "VCCB": "BVB",
        "MOMO": "Momo",
        "ZALOPAY": "ZaloPay",
        "NAPAS": "Napas",
        "HDB_MPGS": "MPGS",
        "GPAY": "GPay",
        "IMEDIA": "Imedia",
        "BIDV": "BIDV",
        "HDBANK": "HDB",
        "HDB-VA": "HDB",
        "CÔNG TY TNHH BẢO HIỂM HD": "GPay",
        "IMEDIA JSC_TOPUP": "Imedia",
        "IMediaBill": "Imedia",
        "IMediaTopup": "Imedia",
        "SKY JOY": "GPay",
        "SKY FI": "GPay",
        "SKYFI": "GPay",
        "VIETJET AIR": "GPay",
        "VNPAY": "VNPay"
  
    }

    merchant_name = merchant_name_map.get(
        merchant,
        merchant.title()
    )

    # ======================================================
    # Merchant mapping
    # ======================================================

    business_mapping = {

        "VJ APP TICKET": "VIETJET",
        "VJ WEB TICKET": "VIETJET",
        "VIETJET WEBSITE": "VIETJET",
        "VIETJET MOBILE": "VIETJET",
        "VJ APP LATER": "VIETJET",
        "VJ WEB LATER": "VIETJET",
        "VJ PAYMENT LINK": "VIETJET",
        "VIETJET AIR": "VIETJET",
        "SKYFI": "SKYFI",
        "SKY JOY": "GJoy",
        "ONAIR": "OnAir",
        "VJ WEB SKYSHOP": "VIETJET",
        "VJ KIOSK": "VIETJET",
        "RONG VIET TRAVEL": "Rồng Việt",

        "BAO HIEM HD": "HDI",
        "CÔNG TY TNHH BẢO HIỂM HD": "HDI",
        "BAO HIEM HD - PORTAL": "HDI",

        "VICTORIA SCHOOL": "VSSS",
        "VICTORIA SCHOOL RIVERSIDE": "KHoan-Riverside",
        "VICTORIA SAIGON SOUTH": "Khải Hoàn-NSG",

        "VINCLUB": "VinApp",
        "MENAS STORE": "Menas",
        
    }

    business = business_mapping.get(
        merchant_id,
        merchant_id
    )

    # ======================================================
    # Cấn trừ phí
    # ======================================================

    if "cấn trừ" in note_lower:

        if business:
            return f"{business}-Cấn trừ Phí"

        return "Cấn trừ Phí"

    # ======================================================
    # Đóng ví
    # ======================================================

    if "đóng tk ví" in note_lower:
        return "Xử lý số dư KH đóng TK Ví"

    # ======================================================
    # HDB / BIDV Wallet
    # ======================================================

    if merchant_name in ["HDB", "BIDV"]:

        bank = merchant_name

        if payment_type == "CASHOUT":

            if recon_result == "MATCH_SUCCESS":
                return f"{bank}-Rút tiền từ TK Ví"

            elif recon_result == "MISS_GPAY":
                return f"{bank}-Lệch Rút tiền Ví(GP)"

            elif recon_result == "GPAY_MORE":
                return f"{bank}-HLệch Rút tiền Ví(GP)"

        elif payment_type in [
            "CASHIN",
            "PAYMENT-CASHIN"
        ]:

            if recon_result == "MATCH_SUCCESS":
                return f"{bank}-GD QR {business}"

            elif recon_result == "MISS_GPAY" or recon_result == "GPAY_MORE":
                return f"{bank}-XLý Lệch Nạp tiền Ví"
            
    # ======================================================
    # NAPAS Wallet
    # ======================================================

    if merchant_name == "Napas":

        if payment_type in [
            "CASHIN",
            "PAYMENT-CASHIN"
        ]:

            if recon_result == "MATCH_SUCCESS":
                return "Napas-Nạp tiền TK Ví"

            elif recon_result == "MISS_PARTNER":
                return "Napas-Lệch Nạp Ví (ĐT)"

            elif recon_result == "MISS_GPAY" or recon_result == "PARTNER_MORE" or recon_result == "GPAY_MORE":

                if "lệch ngày" in note_lower:
                    return "Napas-Nạp tiền TK Ví"

                if "xử lý" in note_lower:
                    return "Napas-XLLệch Nạp tiền Ví"

                return "Napas-HLệch Nạp Ví (ĐT)"

        if payment_type == "PAYMENT":

            if recon_result == "MATCH_SUCCESS":
                if merchant_name.upper() == business.upper():
                    return "Napas-Nạp tiền TK Ví"
                else:
                    return f"Napas-GD UPC {business}"

            elif recon_result == "MISS_GPAY" or recon_result == "MISS_PARTNER":

                if "lệch ngày" in note_lower:
                    return f"Napas-GD UPC {business}"

                return f"Napas-Lệch GD {business}(ĐT)"
            elif recon_result == "GPAY_MORE":
                return f"Napas-XLLệch Nạp tiền Ví"
        if payment_type == "REFUND":

            if recon_result == "MATCH_SUCCESS":
                return f"Napas-Hoàn GD {business}"

            return f"Napas-HLệch {business} (ĐT)"

    # ======================================================
    # Momo
    # ======================================================

    if merchant == "MOMO":

        if payment_type == "PAYMENT":

            if recon_result == "MATCH_SUCCESS":
                return f"Momo-GD UPC {business}"

            if recon_result == "MISS_GPAY":

                if "lệch ngày" in note_lower:
                    return f"Momo-GD UPC {business}(ĐT)"

                return f"Momo-Lệch GD {business}(ĐT)"

        elif payment_type == "REFUND":

            if recon_result == "MATCH_SUCCESS":
                return f"Momo-Hoàn GD UPC {business}"

            return f"Momo-HLệch {business} (ĐT)"

    # ======================================================
    # ZaloPay
    # ======================================================

    if merchant == "ZALOPAY":

        if payment_type == "PAYMENT":

            if recon_result == "MATCH_SUCCESS":
                return f"ZaloPay-GD {business}"

            return f"ZaloPay-Lệch {business} (ĐT)"

        elif payment_type == "REFUND":

            if recon_result == "MATCH_SUCCESS":
                return f"ZaloPay-Hoàn GD {business}"

            return f"ZaloPay-HLệch {business}(ĐT)"

    # ======================================================
    # BVB
    # ======================================================

    if merchant == "VCCB":

        if "ibft" in note_lower:
            return "BVB-DV IBFT Hoàn GD QR"

        if payment_type == "PAYMENT":

            if recon_result == "MATCH_SUCCESS":
                return f"BVB-GD QR {business}"

            return f"BVB-Lệch QR {business} (ĐT)"

        elif payment_type == "REFUND":

            if recon_result == "MATCH_SUCCESS":
                return f"BVB-DV IBFT Hoàn GD QR"

            return f"BVB-HLệch QR {business} (ĐT)"

    # ======================================================
    # GPay
    # ======================================================

    if merchant == "GPAY":
        
        if payment_type == "PAYMENT":

            if recon_result == "MATCH_SUCCESS":
                return f"GPay-GD TT {business}"

            return f"GPay-Lệch {business} (GP)"

        elif payment_type == "REFUND":

            if recon_result == "MATCH_SUCCESS":
                return f"GPay-Hoàn GD {business}"
            
            return f"GPay-HLệch {business} (GP)"

    # ======================================================
    # Imedia
    # ======================================================

    if "IMEDIA" in merchant.upper():

        if payment_type == "PAYMENT":

            if "TOPUP" in merchant.upper():
                if recon_result == "MATCH_SUCCESS":
                    return "Imedia-GD Topup"
                return "Imedia-Lệch GD topup (GP)"

            if "BILL" in merchant.upper():
                if recon_result == "MATCH_SUCCESS":
                    return "Imedia-GD Billing"
                return "Imedia-Lệch GD Bill (GP)"

        elif payment_type == "REFUND":
            if "TOPUP" in merchant.upper():
                return "Imedia-HLệch topup (GP)"
            elif "BILL" in merchant.upper():
                return "Imedia-HLệch Billing (GP)"
            else:
                return f"Imedia-HLệch {merchant} (GP)"

    # ======================================================
    # VNPAY
    # ======================================================

    if merchant == "VNPAY":
        return "VNPay-GD thanh toán"

    # ======================================================
    # Default
    # ======================================================
    # GPay
    # ======================================================

    if merchant_name == "GPay":

        if payment_type == "PAYMENT":

            if recon_result == "MATCH_SUCCESS":
                return f"GPay-GD TT {business}"

            return f"GPay-Lệch {business} (GP)"

        elif payment_type == "REFUND":

            if recon_result == "MATCH_SUCCESS":
                return f"GPay-Hoàn GD {business}"

            return f"GPay-HLệch {business} (GP)"

    # ======================================================
    return f"{merchant_name}-{payment_type}-{business}-{recon_result}"

def build_title_upc(
        merchant,
        payment_type,
        merchant_id,
        recon_result,
        note
):
    merchant = merchant.strip()
    payment_type = payment_type.strip().upper()
    recon_result = recon_result.strip().upper()
    note_lower = note.lower().strip()
    merchant_name = merchant
    
    match merchant:
        case "VCCB":
            merchant_name = "BVB"
        case "MOMO":
            merchant_name = "Momo"
        case "ZALOPAY":
            merchant_name = "ZaloPay"
        case "NAPAS":
            merchant_name = "Napas" 
        case "HDB_MPGS":
            merchant_name = "MPGS"
        case "HDBANK":
            merchant_name = "HDB"
    # =======================
    # Mapping Merchant ID
    # =======================
    merchant_mapping = {
        "VJ APP TICKET": "VIETJET",
        "VIETJET WEBSITE": "VIETJET",
        "VIETJET MOBILE": "VIETJET",
        "VJ WEB TICKET": "VIETJET",
        "VJ APP LATER": "VIETJET",
        "VJ APP SKYPOS": "VIETJET",
        "VJ PAYMENT LINK": "VIETJET",
        "VJ WEB LATER": "VIETJET",
        "VJ WEB SKYSHOP": "VIETJET",
        "VJ KIOSK": "VIETJET",
        "SKYFI": "SKYFI",
        "SKY JOY": "GJoy",
        "SKYJOY": "GJoy",

        "BAO HIEM HD": "HDI",
        "BAO HIEM HD - PORTAL": "HDI",
        "CÔNG TY TNHH BẢO HIỂM HD": "HDI",

        "ONAIR": "OnAir",
        "RONG VIET TRAVEL": "Rồng Việt",

        "VICTORIA SCHOOL": "VSSS",
        "VICTORIA SCHOOL RIVERSIDE": "KHoan-Riverside",
        "VICTORIA SAIGON SOUTH": "Khải Hoàn-NSG",
        "SO BAN HANG": "SBH",
        "VINCLUB": "VinApp",
        "Galaxy Pay": "test",
        "MENAS STORE": "Menas",
        "GIAC MO BAY RE": "Giấc Mơ Bay Rẻ"
    }

    business = merchant_mapping.get(merchant_id, merchant_id)

    # =======================
    # Bỏ qua cấn trừ phí
    # =======================
    if "cấn trừ" in note_lower:
        return "CÁN TRỪ"

    # =======================
    # Bỏ qua BVB-HLệch QR VIETJET (ĐT)
    # =======================
    
    
    # if (
    #     merchant.upper() == "BVB"
    #     and payment_type == "REFUND"
    #     and recon_result == "MISS_GPAY"
    # ):
    #     return None


    # =======================
    # Channel mapping
    # =======================
    channel_mapping = {
        "MOMO": "UPC",
        "NAPAS": "UPC",
        "HDB_MPGS": "UPC",
        "GPAY": "TT",

        "VCCB": "QR",
        "HDB": "QR",
        "SO BAN HANG": "UPC",
        "ZALOPAY": "",
        "VNPAY": "thanh toán"
    }

    channel = channel_mapping.get(
        merchant.upper(),
        ""
    )

    channel_text = f" {channel}" if channel else ""

    source = "(GP)" if merchant.upper() == "GPAY" else "(ĐT)"

    # ============================================
    # PAYMENT
    # ============================================
    if payment_type == "PAYMENT":

        # Payment + Match
        if recon_result == "MATCH_SUCCESS":
            if merchant == "VICTORIA SAIGON SOUTH":
                return f"{merchant_name}-GD {business}"
            elif merchant == "HDB_MPGS" and business == "KHoan-Riverside":
                return f"{merchant_name}-GD {business}"
            return f"{merchant_name}-GD{channel_text} {business}"

        # Payment + MissGpay
        if recon_result == "MISS_GPAY":

            # Miss nhưng lệch ngày => như MATCH
            if "lệch ngày" in note_lower:
                return f"{merchant_name}-GD{channel_text} {business}"

            elif merchant == "MOMO":
                return f"{merchant_name}-Lệch GD {business}{source}"
            elif merchant_name == "HDB":
                return f"{merchant_name}-Lệch GD{channel_text} {business} {source}"
            elif business == "VIETJET":
                if merchant_name == "Napas":
                    return f"{merchant_name}-Lệch GD {business}{source}"
                elif merchant_name == "ZaloPay":
                    return f"{merchant_name}-Lệch {business} {source}"
                return f"{merchant_name}-Lệch QR {business} {source}"
            # Miss thất bại => dòng lệch
            return f"{merchant_name}-Lệch{channel_text} {business} {source}"


    # ============================================
    # REFUND
    # ============================================

    elif payment_type == "REFUND":
        
        if merchant == "ZALOPAY":
            if recon_result == "MATCH_SUCCESS" and note != "MISS_REFUND_FALSE":
                return f"ZaloPay-Hoàn GD {business}"

            return f"ZaloPay-HLệch {business}(ĐT)"
        # Refund bình thường
        elif recon_result == "MATCH_SUCCESS" and note != "MISS_REFUND_FALSE":
            if merchant in ["HDB_MPGS", "VCCB"] or business == "VIETJET":
                if merchant_name == "BVB":
                    if "IBFT không thành công, hoàn bảng kê cho KH" in note:
                        return f"{merchant_name}-Hoàn GD{channel_text} {business}"
                    return f"{merchant_name}-DV IBFT Hoàn GD QR"
                elif merchant_name == "Momo":
                    return f"{merchant_name}-Hoàn GD{channel_text} {business}"
                return f"{merchant_name}-Hoàn GD {business}"
            return f"{merchant_name}-Hoàn GD{channel_text} {business}"

        # Refund lệch
        else:
            if merchant_name == "BVB":
                if "IBFT không thành công, hoàn bảng kê cho KH" in note:
                    return f"{merchant_name}-Hoàn GD{channel_text} {business}"
                return f"{merchant_name}-DV IBFT Hoàn GD QR"
            return f"{merchant_name}-HLệch {business} {source}"

    return f"{merchant_name}-{payment_type} {merchant_id} {recon_result} {note}"


def build_description(title, posting_date, accounting_date, merchant_id, merchant_name, file_name, note):
    
    try:
        posting_date = datetime.strptime(posting_date, "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        # Không phải định dạng YYYY-MM-DD thì giữ nguyên
        pass
    
    pattern = r"\((\d{1,2})-(\d{1,2})\)"
    days_match = re.search(pattern, file_name)
    
    if days_match:
        posting_date = get_date_range_from_filename(file_name)
    
    pattern_months = r"\((\d{1,2})\.(\d{1,2})-(\d{1,2})\.(\d{1,2})\)"
    month_match = re.search(pattern_months, file_name)
    if month_match:
        posting_date = get_date_range_from_months_match(posting_date, file_name)
    
    if merchant_id == "VJ WEB SKYSHOP":
        title = f"{title}-SKYSHOP"
    elif merchant_id == "VJ KIOSK":
        title = f"{title}-KIOSK"
        
    if "KHoan" in title :
        title = title.replace("KHoan", "Khải Hoàn")
    elif "Rồng Việt" in title:
        title = title.replace("Rồng Việt", "Rồng Việt Travel")
        
    if "XLLệch" in title:
        title = title.replace("XLLệch ", "Xử lý Lệch ")
    elif "XLý Lệch" in title:
        title = title.replace("XLý Lệch ", "Xử lý Lệch ")
    elif "HLệch" in title:
        title = title.replace("HLệch ", "Hoàn Lệch ")
        #print(title + "+"*60)
        
        
    if "(ĐT)" in title:
        rtitle = title.split(" (ĐT)")[0] if " (ĐT)" in title else title.split("(ĐT)")[0]
        rtitle = f"{rtitle} N{posting_date} (ĐT)"
        if "Nạp Ví" in title:
            rtitle = f"{rtitle.replace(" Nạp Ví ", " Nạp tiền vào TK ví ")}"
            #print(rtitle + "="*60)
        elif "BVB" in title:
            rtitle = f"{rtitle.replace("QR", "GD QR code")}"
        elif "Zalo" in title:
            rtitle = f"{rtitle.replace("Lệch ", "Lệch GD ")}"
    
    elif "(GP)" in title:
        rtitle = title.split(" (GP)")[0] if " (GP)" in title else title.split("(GP)")[0]
        if "Lệch GD Bill" in title:
            rtitle = f"{rtitle}ing N{posting_date} (GP)"
        else:
            rtitle = f"{rtitle} N{posting_date} (GP)"
        
    else:
        if "BVB" in title:
            if "GJoy" in title or "SKYFI" in title or "VSSS" in title:
                rtitle = f"{title.replace(" QR ", " Thanh toán QR code ")} N{posting_date}"
            else:   
                rtitle = f"{title.replace(" QR ", " QR code ")} N{posting_date}"
            
        elif "TT" in title:
            rtitle = f"{title.replace(" TT ", " Thanh toán ")} N{posting_date}"
        elif "TK Ví" in title:
            rtitle = f"{title.replace(" TK Ví", " vào tài khoản ví")} N{posting_date}"
        elif merchant_id == "BAO HIEM HD - PORTAL":
            rtitle = f"{title.replace("HDI", "HDI-Portal")} N{posting_date}"
        else:
            rtitle = f"{title} N{posting_date}"


    if "IBFT" in title and "BVB" in title:
        mapping = mapping_merchant(merchant_name, merchant_id)
        merchant_id_map = mapping["business_map"]
        if note == "MISS_REFUND_FALSE":
            rtitle = f"BVB-Dịch vụ IBFT Hoàn Lệch QR code {merchant_id_map}"
        else:
            rtitle = f"BVB-Dịch vụ IBFT Hoàn GD QR code {merchant_id_map}"
    elif any(keyword in title for keyword in (
        "Hoàn",
        "Xử lý Lệch",
        "Hoàn Lệch",
        "HLệch",
        "XLý Lệch",
    )):
        if "GPay" in title:
            rtitle = f"{title.replace("GD ", "GD thanh toán ")}"
            # print(f"đổi tên GPAY GD Thanh Toán ============++++++++")
        elif "BVB-Hoàn" in title:
            rtitle = f"{title.replace("GD QR ", "GD Thanh toán QR code ")}"
        elif "topup" in title and "Hoàn Lệch" in title:
            rtitle = f"{title.replace("Hoàn Lệch ", "Hoàn Lệch GD ")}"
            # print(f"đổi tên IMEDIA Hoàn Lệch GD ============++++++++")
        elif "Khải Hoàn" in title:
            rtitle = f"{title.replace(" QR ", " QR code ")} N{posting_date}"
        elif "Hoàn Lệch Rút tiền Ví(GP)" in title:
            rtitle = title.replace("Hoàn Lệch Rút tiền Ví(GP)", "Xử lý Lệch Rút tiền TK Ví (GP)")
        elif "Xử lý Lệch" in title or "Napas-Hoàn Lệch" in title:
            if "Napas" in title:
                if "Nạp Ví" in title:
                    print("z"*60)
                    rtitle = title.replace("Nạp Ví", "Nạp tiền vào tài khoản ví")
                else:
                    rtitle = title.replace("Nạp tiền Ví", "Nạp tiền TK Ví (ĐT)")
            elif "HDB" in title:
                rtitle = title.replace("Nạp tiền Ví", "Nạp tiền vào TK Ví (ĐT)")
        elif "VIETJET" in title:
            rtitle = title.replace("VIETJET", "GD VIETJET")
        else:
            rtitle = title
    # if posting_date != accounting_date:
    #     return f"{title} N{accounting_date}"
    
    return rtitle

def build_document_type(
                title,
                debit_info,
                credit_info,
                merchant,
                payment_type,
                recon_result,
                note
            ):
    """
    Xác định SAP Document Type.

    Input:
        debit_info["account"]
        credit_info["account"]

    Output:
        SA / KA / DZ / KR / ...
    """
    document_type = ""
    debit = debit_info["account"]
    credit = credit_info["account"]
    SA = ["112", "511", "711", "811", "642", "632"] 
    #print(debit[0:2], credit[0:2])
    # ===========================
    # Rule 1
    if debit[0:2] in ["13", "33"] and credit[0:2] in ["13", "33"]:
        document_type = "DK"
    # ===========================

    elif debit.startswith("33") and credit.startswith("33"):
        document_type = "KZ"
    # ===========================
    # Rule 2
    # ===========================

    elif debit.startswith("131") and credit.startswith("511"):
        document_type = "RV"

    elif debit[0:2] in SA and credit[0:2] in SA:
        document_type = "SA"
    else:
        document_type = "CL"
    # ===========================
    # Rule 3
    # ===========================

    # if "HOÀN" in title.upper():
    #     document_type = "KA"

    # ===========================
    # Default
    # ===========================
    if "DV IBFT Hoàn" in title and note == "MISS_REFUND_FALSE":
        document_type = "CL"
    return document_type

def main():

    rows_out = []

    with open(INPUT_FILE, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)

        stt = 1

        for row in reader:

            file_name = row.get("File", "")
            merchant = row.get("Đối tác", "")
            merchant_id = row.get("Merchant ID", "")
            payment_type = row.get("Loại giao dịch", "")
            amount = row.get("Tổng số tiền", "0").replace(",", "")
            origin_payment_date = row.get("Ngày giao dịch gốc", "")
            note = row.get("Ghi chú", "")
            recon_result = row.get("Kết quả đối soát", "")
            posting_date = row.get("Ngày giao dịch", "").split(" ")[0]
            try:
                posting_date = normalize_posting_date(posting_date)
                # if "1/8" in posting_date or "01/08" in posting_date:
                #     print(f"[WARN] Posting date is 1/8, check file: {posting_date}")
            except ValueError:
                pass
            accounting_date = posting_date
            
            origin_payment_date = row.get("Ngày giao dịch gốc", "").strip()

            is_after_23 = False

            if origin_payment_date:
                dt = datetime.strptime(origin_payment_date, "%d/%m/%Y %H:%M:%S")
                is_after_23 = dt.time() >= time(23, 0, 0)

            if "MPGS" in merchant.upper() and (is_after_23 or recon_result == "MISS_GPAY"):
                #print(f"before plus 1 day {accounting_date}")
                accounting_date = (
                    datetime.strptime(posting_date, "%d/%m/%Y") + timedelta(days=1)
                ).strftime("%d/%m/%Y")
                #print(f"after plus 1 day {accounting_date}")   
            try:
                amount = float(amount)
            except:
                print(f"[WARN] Cannot parse amount: {amount} in file: {file_name}")
                # amount = 0

            # ===== PARSE DATE =====
            # posting_date = parse_date_from_filename(file_name)
            # if not posting_date:
            #     print("[WARN] Cannot parse date from:", file_name)
            # YYYYMMDD cho asset value date
            if len(posting_date) < 10:
                # print(f"[WARN] Posting date is too short: {posting_date} in file: {file_name}")
                posting_date = parse_date_from_filename(file_name)
                # print(f"[INFO] Using parsed posting date: {posting_date} for file: {file_name}")
            # posting_date_ymd = posting_date.replace("/", "")

            # ===== TITLE =====
            if "RECON" in file_name.upper():
                title = build_title_upc(merchant,payment_type,merchant_id,recon_result,note)
            else:
                title = build_title_wallet(merchant,payment_type,merchant_id,recon_result,note)

            if "XLý Lệch Nạp tiền Ví" in title or "HLệch Nạp Ví (ĐT)" in title or "XLLệch Nạp tiền Ví" in title or "HLệch Rút tiền Ví" in title:
                accounting_date = parse_posting_date_from_filename(file_name)
            
            # ===== DIỄN GIẢI (FIXED) =====
            description = build_description(title, posting_date, accounting_date, merchant_id, merchant, file_name, note)
            
            debit_info = build_posting_info(title, "Nợ", merchant, merchant_id, note)
            credit_info = build_posting_info(title, "Có", merchant, merchant_id, note)
            document_type = build_document_type(
                title,
                debit_info,
                credit_info,
                merchant,
                payment_type,
                recon_result,
                note
            )
            # CREATE 2 ROWS (NỢ / CÓ)
            # =======================================================

            for dr_cr in ["Nợ", "Có"]:
                posting_info = build_posting_info(title, dr_cr, merchant, merchant_id, note)
                if amount == 0:
                    print(f"[WARN] Amount is 0, skipping row: {row}")
                    continue
                else:
                    rows_out.append({
                        "STT": "",
    
                        # ===== FIXED DATES =====
                        "Ngày chứng từ": accounting_date,
                        "Ngày hạch toán": accounting_date,
                        "Period (Kì hạch toán)": "",

                        "Tiêu đề chứng từ": title,
                        "Mẫu hóa đơn": "",
                        "Số hoá đơn (Reference)": "",
                        "Loại chứng từ": document_type,
                        "Mã công ty": "1080",

                        "Đ.vị tiền tệ": DEFAULT_CURRENCY,
                        "Ngày lấy tỉ giá trong hệ thống": "",

                        "Reference Key 1 HD (Không cần điền)": "",
                        "Reference Key 2 HD (Không cần điền)": "",

                        "Nợ/Có": dr_cr,
                        "Tài khoản": posting_info["account"],
                        "Mã BP": posting_info["business_unit"],
                        "Mã tài sản": "",
                        "Transaction type": "",

                        "Asset value date (Để định dạng YYYYMMDD)": "",

                        "Special GL indicator": "",

                        "Số tiền": f"{amount:,.0f}",

                        "Số tiền quy đổi (Khi currency là VND thì bỏ trống cột này, KHÔNG ĐƯỢC ĐỂ SỐ 0)": "",

                        "Base Amount": "",
                        "Thuế suất": "",
                        "Số tiền thuế": "",
                        "Số tiền thuế quy đổi (Khi currency là VND thì bỏ trống cột này, KHÔNG ĐƯỢC ĐỂ SỐ 0)": "",

                        "Chi cục thuế (Không cần điền)": "",
                        "Chi nhánh (Không cần điền)": "",

                        "TT Chi phí (Cost center)": "",
                        "TT lợi nhuận (Profit center)": DEFAULT_PROFIT_CENTER,

                        "Payment cur.": "",
                        "Pmnt/c amnt": "",

                        # ===== FIXED DIỄN GIẢI =====
                        "Diễn giải": description,

                        "Nội dung gán (Không cần điền)": "",
                        "MSN": "",
                        "Serial Number": "",

                        "Order (Nhập lessor cho tài khoản 335* hoặc nhập mã chương trình sự kiện cho TK CP)": "",
                        "Máy bay (Nhập cho TK Doanh thu)": "",
                        "Chặng bay (Nhập cho TK Doanh thu)": "",

                        "Điều khoản thanh toán": "",
                        "Ngày tính nợ": "",
                        "Khóa thanh toán": "",
                        "Phương thức thanh toán": "",
                        "Mã tài khoản ngân hàng giao dịch": "",
                        "Đánh dấu ghi âm": "",

                        "Reference Key 1 Line (Không cần điền)": "",
                        "Reference Key 2 Line (Với một số loại chứng từ như KA thì điền Profit Center vào Ref key 2)": DEFAULT_PROFIT_CENTER,
                        "Reference Key 3 Line (Không cần điền)": "",

                        "Tên (Khai báo hóa đơn cho NCC vãng lai)": "",
                        "Địa chỉ (Khai báo hóa đơn cho NCC vãng lai)": "",
                        "Tỉnh/ thành phố (Khai báo hóa đơn cho NCC vãng lai)": "",
                        "Mã số thuế (Khai báo hóa đơn cho NCC vãng lai)": "",
                        "Bank country (Khai báo hóa đơn cho NCC vãng lai)": "",
                        "Bank key (Khai báo hóa đơn cho NCC vãng lai)": "",
                        "Bank account (Khai báo hóa đơn cho NCC vãng lai)": "",
                        "CN ngân hàng (Khai báo hóa đơn cho NCC vãng lai)": "",
                        "Type of Bussiness (Khai báo hóa đơn cho NCC vãng lai)": "",

                        "Reference key 4": "",
                        "Purchasing Doc.": "",
                        "Trading.Partner ( Phục vụ hợp nhất BCTC)": ""
                    })
                    

        rows_out.sort(
            key=lambda row: (
                datetime.strptime(row["Ngày chứng từ"], "%d/%m/%Y")
                )
        )
        
        for i in range(0, len(rows_out), 2):
            stt = i // 2 + 1
            rows_out[i]["STT"] = stt
            rows_out[i + 1]["STT"] = stt
# Tạo timestamp: 16-50_31-07-2026
    timestamp = datetime.now().strftime("%H-%M_%d-%m-%Y")

# upload_file_16-50_31-07-2026.csv
    output_file = OUTPUT_FILE.replace(".csv", f"_{timestamp}.csv")
    # write CSV
    with open(output_file, "w", newline="", encoding="utf-8-sig") as f:
        fieldnames = list(rows_out[0].keys())
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows_out)

    print(f"[INFO] Exported upload file: {output_file}")
    xlsx_file = output_file.replace(".csv", ".xlsx")
    csv_to_xlsx(output_file, xlsx_file)

if __name__ == "__main__":
    main()