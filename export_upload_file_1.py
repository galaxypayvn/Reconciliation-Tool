#!/usr/bin/env python3

import csv
from datetime import datetime, time, timedelta
import re
import calendar
from pathlib import Path
from openpyxl import Workbook
from convert_data_to_json import FILE_PREFIX, WALLET_FILE_PREFIX, WALLET_NEW_FILE_PREFIX, WALLET_NEW_FILE_PREFIX_INPUT, WALLET_NEW_FILE_PREFIX_OUTPUT
from helper import (extract_bvbank_merchant, extract_hdbank_merchant, get_date_range_from_months_match,
    get_posting_date_minus_1, convert_payment_info_to_output, normalize_year, csv_to_xlsx, 
    parse_date_from_filename, parse_posting_date_from_filename, normalize_posting_date)
# from trio import Path
INPUT_FILE = "csv_recon_report/recon_summary.csv"
OUTPUT_FILE = "csv_recon_report/upload_file.csv"


DEFAULT_CURRENCY = "VND"
DEFAULT_PROFIT_CENTER = "10800001"

BUSINESS_UNIT = {
        # Merchant (Partner)
        "VIETJET": "1010",
        "UPC VIETJET AIR": "1010",
        "UPC MPGS VIETJET AIR": "1010",
        "VIETJET KIOSK": "1010",
        "VIETJET SKYSHOP": "1010",
        "GALAXY PAY": "1080",
        "GPAY": "1080",
        "SKYFI": "1000003120",
        "GALAXY TELECOM": "1000003120",
        "GJOY": "0037888888",
        "GALAXY JOY": "0037888888",
        "SKY JOY": "0037888888",
        "SKYJOY": "0037888888",
        "RỒNG VIỆT": "0037486431",
        "RONG VIET": "0037486431",
        "HDI": "1000000860",
        "HDI PORTAL": "1000000860",
        "ONAIR": "1000002978",
        "VSSS": "1000002177",
        "VSSS VICTORIA INTERNATIONAL SCHOOL": "1000002177",
        "VSSS (VICTORIA INTERNATIONAL SCHOOL)": "1000002177",
        "VSSS VICTORIA SCHOOL": "1000002177",
        "TRUONG KHAI HOAN-NAM SG": "1000003029",
        "KHẢI HOÀN-NSG": "1000003029",
        "TRUONG KHAI HOANNSG": "1000003029",
        "KHOAN-RIVERSIDE": "1000003028",
        "TRUONG KHAI HOAN-RIVERSIDE": "1000003028",
        "TRUONG KHAI HOANRIVERSIDE": "1000003028",
        "SBH": "1000002630",
        "UPC SBH": "1000002630",
        "VINAPP": "1000002977",
        "VINCLUB": "1000002977",
        "VIKKAFE": "1000003587",
        "QR CODE VIKKAFE": "1000003587",
        "MENAS": "1000000265",

        # Đối tác
        "MOMO": "0037385423",
        "ZALOPAY": "0037482684",
        "NAPAS": "1000002622",
        "BVB": "1000002626",
        "BVBANK": "1000002626",
        "VCCB": "1000002626",
        "VNPAY": "0037405944",
        "IMEDIA": "1000002623",
        "HDB": "0037360001",
        "HDB-VA": "0037360001",
        "HDBANK": "0037360001",
        "HDBANK_STATEMENT": "0037360001",
        "BIDV": "0037360001",
        "MPGS": "0037360001",
        "GIAC MO BAY RE": "0037483314",
        "GD GIAC MO BAY RE": "0037483314",
        "SML": "1000002847",
        "MAILAND": "1000003027",
        "BMGĐ": "1000000932",
        "BMI - ABBANK": "1000000932",
        "BẢO MINH": "1000003702",
        "GOTRUST": "1000003386",
        "VNS": "1000002624",
    }







    

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
        "SKY JOY": "GJoy",
        "Galaxy Joy": "GJoy",
        "SKY FI": "SKYFI",
        "SKYFI": "SKYFI",
        "Galaxy Telecom": "SKYFI",
        "VIETJET AIR": "VIETJET",
        "VIETJET Air": "VIETJET",
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
        "Galaxy Telecom": "SKYFI",
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


def build_posting_info(title, dr_cr, merchant, merchant_id, note, payment_type, recon_result, file_name):

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
        "IMEDIA JSC": "IMEDIA",
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
        "NAPAS" : "HDB",
        "MENAS STORE": "Menas"
    }
    merchant_name_map = {
        "VCCB": "BVB",
        "BVBANK": "BVB",
        "MOMO": "Momo",
        "ZALOPAY": "ZaloPay",
        "NAPAS": "Napas",
        "HDB_MPGS": "MPGS",
        "GPAY": "GPay",
        "BIDV": "BIDV",
        "HDBANK": "HDB",
        "HDB-VA": "HDB",
        "IMEDIA JSC_TOPUP": "HDB",
        "IMEDIA JSC": "HDB",
        "CÔNG TY TNHH BẢO HIỂM HD": "HDB",
        "IMEDIABILL": "HDB",
        "IMEDIATOPUP": "HDB",
        "SKY JOY": "HDB",
        "SKY FI": "GPHDBay",
        "SKYFI": "HDB",
        "VIETJET AIR": "HDB",
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
    #account = partner_acc if dr_cr == "Nợ" else merchant_acc
    business_unit = partner_bu if dr_cr == "Nợ" else merchant_bu
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

    
    
    account_hd_bvb = "1121121305" if "HDBANK" in file_name else "1121122302"
    if "BC GD" in title_upper or "BÁO CÓ GD" in title_upper:
     
        if dr_cr == "Nợ":
            account = account_hd_bvb
            business_unit = ""
        else:
            account = "1388404200"
            business_unit = merchant_bu  
    # ======================================================
    # BVB
    # ======================================================

    elif "BVB" in title_upper:

        if "IBFT" in title_upper:
            if note == "MISS_REFUND_FALSE":
                account = "3388403200" if dr_cr == "Nợ" else "3388403100"
                business_unit = partner_bu
            else:
                account = "3388403100"
        
      
    elif "LÃI TIỀN GỬI" in title_upper:
        business_unit = ""
        if dr_cr == "Nợ":
            account = account_hd_bvb
        else:
            account = "5151510100"
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
            elif "HLỆCH" in title_upper or "HOÀN LỆCH" in title_upper:
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
        elif "HLỆCH" in title_upper:
            business_unit = "0037360001"
            if dr_cr == "Nợ":
                account = "3388402000"
            else:
                account = "3388401000"
        elif "LỆCH" in title_upper:
            business_unit = "0037360001"
            if dr_cr == "Nợ":
                account = "3388401000"
            else:
                account = "3388402000"
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

        if "XỬ LÝ" in title_upper or "XLÝ" in title_upper or "HLệch Rút tiền Ví" in title:

            if dr_cr == "Nợ":
                account = "3388402000"
            else:
                account = "3388401000"
                
        elif "GD QR" in title_upper and "LỆCH" not in title_upper and "Hoàn" not in title and "TEST" not in title_upper:
            if dr_cr == "Nợ":
                account = "1388404200"
            else:
                account = "3388403100"
        
        elif "HLỆCH TEST" in title_upper:
            if dr_cr == "Nợ":
                account = "3388403300"
                business_unit = "1080"
            else:
                account = "1388404200"
                business_unit = "0037360001"    
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
        elif "BIDV-Lệch Rút tiền Ví" in title :
            if dr_cr == "Nợ":
                account = "3388401000"
            else:
                account = "3388402000" 
                
        
        elif "Hoàn" not in title or "HDB-HLệch" in title:

            if dr_cr == "Nợ":
                account = "3388403200"
            else:
                account = "1388404200"
        elif "LỆCH" in title_upper:

            if dr_cr == "Nợ":
                account = "1388404200"
            else:
                account = "3388403200"
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
    
    # ======================================================
    # NEW_WALLET_INPUT_FILE
    # ======================================================
    if WALLET_NEW_FILE_PREFIX_INPUT in file_name:
        if merchant == "HDB":
            if ("nạp tiền ví" in title.lower() or "qr nạp tk ví" in title.lower()) and "lệch" not in title.lower():
                account = "1121121302" if dr_cr == "Nợ" else "3388401000"
                business_unit = "" if dr_cr == "Nợ" else "0037360001"
            elif "lệch qr nạp tk ví" in title.lower():
                account = "1121121302" if dr_cr == "Nợ" else "3388402000"
                business_unit = "" if dr_cr == "Nợ" else "0037360001"
        elif merchant == "HDBANK":
            if "nạp tiền tk ví" in title.lower():
                account = "1121121302" if dr_cr == "Nợ" else "3388401000"
                business_unit = "" if dr_cr == "Nợ" else "0037360001"
            elif "lệch qr nạp tk ví" in title.lower():
                account = "1121121302" if dr_cr == "Nợ" else "3388402000"
                business_unit = "" if dr_cr == "Nợ" else "0037360001"
            elif "lệch rút tiền" in title.lower():
                account = "3388401000" if dr_cr == "Nợ" else "3388402000"
                business_unit = "0037360001"
            elif "rút tiền" in title.lower():
                account = "3388401000" if dr_cr == "Nợ" else "1121121302"
                business_unit = "0037360001" if dr_cr == "Nợ" else ""
        elif merchant == "BIDV":
            if "nạp tiền" in title.lower():
                account = "1121123302" if dr_cr == "Nợ" else "3388401000"
                business_unit = "" if dr_cr == "Nợ" else "0037360001"
            elif "lệch nạp tiền" in title.lower():
                account = "1121123302" if dr_cr == "Nợ" else "3388402000"
                business_unit = "" if dr_cr == "Nợ" else "0037360001"
            elif "lệch rút tiền" in title.lower():
                account = "3388401000" if dr_cr == "Nợ" else "3388402000"
                business_unit = "0037360001"
            elif "rút tiền" in title.lower():
                account = "3388401000" if dr_cr == "Nợ" else "1121123302"
                business_unit = "0037360001" if dr_cr == "Nợ" else ""
            
    
    return {
        # "document_type": document_type,
        "account": account,
        "business_unit": business_unit
    }
def build_posting_info_mpgs(title, dr_cr, merchant, merchant_id, note, payment_type, recon_result, file_name, amount):

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
        "GALAXYPAY": "GPAY",
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
        "NAPAS" : "HDB",
        "MENAS STORE": "Menas",
        "SMILE LIVING": "SML",
        "BAO MINH GIA DINH": "BMGĐ",
        "TONG CONG TY BAO MINH": "BẢO MINH",
    }
    merchant_name_map = {
        "VCCB": "BVB",
        "BVBANK": "BVB",
        "MOMO": "Momo",
        "ZALOPAY": "ZaloPay",
        "NAPAS": "Napas",
        "HDB_MPGS": "MPGS",
        "GPAY": "GPay",
        "BIDV": "BIDV",
        "HDBANK": "HDB",
        "HDB-VA": "HDB",
        "IMEDIA JSC_TOPUP": "HDB",
        "CÔNG TY TNHH BẢO HIỂM HD": "HDB",
        "IMEDIABILL": "HDB",
        "IMEDIATOPUP": "HDB",
        "SKY JOY": "HDB",
        "SKY FI": "HDB",
        "SKYFI": "HDB",
        "VIETJET AIR": "HDB",
        "VNPAY": "VNPay"
  
    }

    merchant_name = merchant_name_map.get(
        merchant,
        merchant.title()
    ).upper()

    business = merchant_mapping.get(merchant_id, merchant_id).upper()
    
    account_hd_bvb = "1121121305" if "HDBANK" in file_name else "1121122302"
    # ======================================================
    # BUSINESS UNIT MAPPING
    # ======================================================
    
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
    
    if dr_cr == "Nợ" and "HLệch" not in title:
        account = "1388404200"
    else:
        account = "3388403100"
    partner_bu = BUSINESS_UNIT.get(merchant_name, merchant_name)
    merchant_bu = BUSINESS_UNIT.get(business, business)
    partner_acc = ACCOUNT_MAPPING.get(merchant_name, merchant_name)
    merchant_acc = ACCOUNT_MAPPING.get(business, business)
    #account = partner_acc if dr_cr == "Nợ" else merchant_acc
    business_unit = partner_bu if dr_cr == "Nợ" else merchant_bu
    # ======================================================
    if payment_type == "PAYMENT" and recon_result == "MATCH_SUCCESS" and business == "GPAY" and "TEST" in title_upper:
        account = "1388404200" if dr_cr == "Nợ" else "3388403300"
        business_unit = "0037360001" if dr_cr == "Nợ" else "1080"
    elif payment_type == "PAYMENT" and recon_result == "MATCH_SUCCESS":
        account = "1388404200" if dr_cr == "Nợ" else "3388403100"
        business_unit = "0037360001" if dr_cr == "Nợ" else merchant_bu
    elif payment_type == "REFUND" and recon_result == "MATCH_SUCCESS" and note != "MISS_REFUND_FALSE":
        account = "3388403100" if dr_cr == "Nợ" else "1388404200"
        business_unit = merchant_bu if dr_cr == "Nợ" else "0037360001"
        if "TEST" in title_upper:
            account = "3388403300" if dr_cr == "Nợ" else "1388404200"
            business_unit = "1080" if dr_cr == "Nợ" else "0037360001"
    elif payment_type == "PAYMENT" and recon_result == "MISS_PARTNER":
        #không ghi nhận hạch toán
        account = "NONE"
        business_unit = "NONE"
    elif payment_type == "PAYMENT" and recon_result == "MISS_GPAY":
        business_unit = "0037360001" if dr_cr == "Nợ" else merchant_bu
        if "lệch ngày" in note: 
            account = "1388404200" if dr_cr == "Nợ" else "3388403100"
        elif "Galaxy Pay: thất bại" in note:
            account = "1388404200" if dr_cr == "Nợ" else "3388403200"
    elif payment_type == "REFUND" and recon_result == "MATCH_SUCCESS" and note == "MISS_REFUND_FALSE":
        account = "3388403200" if dr_cr == "Nợ" else "1388404200"
        business_unit = merchant_bu if dr_cr == "Nợ" else "0037360001"
    elif "LÃI TIỀN GỬI" in title_upper:
        business_unit = ""
        if dr_cr == "Nợ":
            account = account_hd_bvb
        else:
            account = "5151510100"
    else:
        #account = partner_acc if dr_cr == "Nợ" else merchant_acc
        business_unit = partner_bu if dr_cr == "Nợ" else merchant_bu
    
    #HDBANK_STATMENT
    if merchant == "HDBANK_STATEMENT" or "HDBANK" in file_name or "SAOKE_BVBANK" in file_name.upper():
        if "Phát sinh có" in note:
            account = account_hd_bvb if dr_cr == "Nợ" else "1388404200"
            if "NAPAS-BC" in title.upper() and dr_cr == "Có":
                account = "1388404300"
            business_unit = "" if dr_cr == "Nợ" else partner_bu
        elif "Phí" in title:
            account = "3311401000" if dr_cr == "Nợ" else account_hd_bvb
            business_unit = partner_bu if dr_cr == "Nợ" else ""
            if "BVB-Phí" in title and amount < 200000:
                account = "6321510200" if dr_cr == "Nợ" else account_hd_bvb
                if (dr_cr == "Nợ" and "669" in file_name):
                    business_unit = "1000002626" 
        elif "Chuyển tiền" in title or "Ch.tien" in title:
            account = "3388403100" if dr_cr == "Nợ" else account_hd_bvb
            business_unit = merchant_bu if dr_cr == "Nợ" else ""
        elif "Phát sinh nợ" in note:
            account = "3311401000" if dr_cr == "Nợ" else account_hd_bvb
            
            business_unit = partner_bu if dr_cr == "Nợ" else ""
        
            
    
    return {
        # "document_type": document_type,
        "account": account,
        "business_unit": business_unit
    }

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
        "HDBANK": "HDBANK",
        "HDB": "HDB",
        "HDB-VA": "HDB",
        "BMI - ABBANK": "BMI-ABBANK",
        "CÔNG TY TNHH BẢO HIỂM HD": "GPay",
        "IMEDIA JSC_TOPUP": "Imedia",
        "IMEDIA JSC": "Imedia",
        "IMEDIABILL": "Imedia",
        "IMEDIATOPUP": "Imedia",
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
        "HDBANK": "",
        "HDB": "",
        "BIDV": "",
        
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

    if merchant_name in ["HDB", "BIDV", "HDBANK"]:

        bank = merchant_name

        if payment_type == "CASHOUT":
            
            if recon_result == "MATCH_SUCCESS":
                bank = "HDB" if "HDB" in bank else bank
                if bank == "BIDV":
                    return f"{bank}-Rút tiền TK Ví GPay"
                return f"{bank}-Rút tiền TK Ví"

            elif recon_result == "MISS_GPAY":
                bank = "HDB" if "HDB" in bank else bank
                return f"{bank}-Lệch Rút tiền Ví(GP)"

            elif recon_result == "GPAY_MORE":
                bank = "HDB" if "HDB" in bank else bank
                return f"{bank}-HLệch Rút tiền Ví(GP)"
            elif recon_result == "MISS_PARTNER":
                bank = "HDB" if "HDB" in bank else bank
                return f"{bank}-Lệch Rút tiền Ví(GP)"
            
        elif payment_type in [
            "CASHIN",
            "PAYMENT-CASHIN",
            "PAYMENT"
        ]:

            if recon_result == "MATCH_SUCCESS":
                if merchant_name == "HDB":
                    return f"{bank}-QR Nạp TK Ví"
                if merchant_name == "HDBANK":
                    return f"HDB-Nạp tiền TK Ví"
                if merchant_name == "BIDV":
                    return f"{bank}-Nạp tiền TK Ví GPay"
                return f"{bank}-GD QR {business}"

            elif recon_result == "MISS_GPAY" or recon_result == "GPAY_MORE":
                if merchant_name == "HDB":
                    return f"{bank}-Lệch QR Nạp TK Ví"
                return f"{bank}-XLý Lệch Nạp tiền Ví"
            elif recon_result == "MISS_PARTNER":
                if bank == "BIDV":
                    return f"{bank}-Lệch nạp tiền Ví(ĐT)"
                elif bank == "HDBANK":
                    return "HDB-Lệch QR Nạp TK Ví(ĐT)"
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
                elif recon_result == "MISS_GPAY":
                    return "Napas-Lệch Nạp Ví (ĐT)"
                elif "xử lý" in note_lower or recon_result == "PARTNER_MORE":
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
        "BMI - ABBANK": "BMI-ABBANK",
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
        "GIAC MO BAY RE": "GMBR"
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
    if merchant == "HDB_MPGS" and merchant_id == "GIAC MO BAY RE":
        business = "GMBR"
    if merchant == "HDB_MPGS" and merchant_id == "VICTORIA SCHOOL RIVERSIDE":
        business = "KH-Riverside"
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
            elif merchant == "HDB_MPGS" and merchant_id == "GIAC MO BAY RE":
                return f"{merchant_name}-GD UPC GMBR"
            elif merchant == "HDB_MPGS" and merchant_id == "VICTORIA SCHOOL RIVERSIDE":
                return f"{merchant_name}-GD UPC KH-Riverside"
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
                elif merchant_name == "MPGS":
                    return f"{merchant_name}-Lệch GD {business} {source}"
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
def build_title_hdbank(
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
    merchant_name = merchant.upper()
    
    if merchant == "HDBANK_STATEMENT" and payment_type == "PAYMENT_RECEIVED":
        merchant_name = "MPGS"
    elif merchant_id == "GPay" and payment_type == "PAYMENT_RECEIVED":
        match merchant_id:
            case "MOMO":
                merchant_id = "UPC"
            case "ZALOPAY":
                merchant_id = "VIETJET"
            case "NAPAS":
                merchant_id = "Cổng" 
    elif "VSSS" in merchant_id:
        merchant_id = "VSSS"
    else:
        merchant_name = "HDB"
    match merchant_id:
        case "MOMO":
            merchant_name = "Momo"
        case "ZALOPAY":
            merchant_name = "ZaloPay"
        case "NAPAS":
            merchant_name = "Napas"
            
    # =======================
    # Mapping Merchant ID
    # =======================
    merchant_mapping = {
        "32400458": "SKYFI",
        "06800680": "GJOY",
        "06800948": "VIETJET",
        "32400035": "KH-RVS",
        "32400036": "KH-NSG",
        "32400629": "MENAS",
        "32403748": "GOTRUST",
        "32402744": "Giấc Mơ Bay Rẻ",
        "GD Giac Mo Bay Re": "Giấc Mơ Bay Rẻ",
        "UPC Vietjet Air": "VIETJET",
        "VIETJET Air": "VIETJET",
        "Galaxy Telecom": "SKYFI",
        "Galaxy Joy": "-GJoy",
        "UPC MPGS Vietjet Air": "VIETJET",
        "HDI": "-HDI",
        "HDI Portal": "-HDI",
        "Rong Viet Travel": "-RồngViệt",
        "32400034": "ONAIR",
        "06800950": "SBH",
        "06800902": "HDI",
        "06800891": "VSSS",
        "06800681": "RỒNG VIỆT",
        "Truong Khai HoanRIVERSIDE": "KH-RVS",
        "Truong Khai Hoan-RIVERSIDE": "KH-RVS",
        "Truong Khai HoanNSG": "-NSG",
        "Menas Store": "-Menas",
        "ONAIR": "-OnAir",
        "SBH": "-SBH",
        "UPC SBH": "SBH",
        "Napas": "Cổng",
        "NAPAS": "Cổng",
        "QR Code Vikkafe": "Vikkafe",
        "MOMO": "UPC",
        "ZALOPAY": "UPC",
    }

    business = merchant_mapping.get(merchant_id, merchant_id)

    # =======================
    # Channel mapping
    # =======================
    channel_mapping = {
        "MOMO": "BC DV",
        "UPC Vietjet Air": "UPC",
        "UPC MPGS Vietjet Air": "UPC",
        "Galaxy Telecom": "UPC",
        "Galaxy Joy": "GD UPC",
        "GJoy": "GD UPC",
        "HDI": "GD UPC",
        "HDI Portal": "GD UPC",
        "Truong Khai HoanNSG": "GD UPC",
        "Menas Store": "GD UPC",
        "ONAIR": "GD UPC",
        "SBH": "GD UPC",
        "UPC SBH": "GD UPC",
        "Rong Viet Travel": "UPC",
        "VSSS": "UPC",
        "SKYFI": "UPC",
        "32400458": "UPC",
        "06800948": "UPC",
        "UPC MPGS Vietjet Air": "UPC",
        "UPC Vietjet Air": "UPC",
        "Galaxy Telecom": "UPC",

    }

    channel = channel_mapping.get(
        merchant,
            "BC GD" if payment_type == "PAYMENT_RECEIVED" else "GD UPC"
        )
    channel_text = f"{channel}" if channel else ""   
    if not business.startswith("-"):
            business = f" {business}"

    # ============================================
    # PAYMENT_RECEIVED
    # ============================================
    if payment_type == "PAYMENT_RECEIVED":
        if merchant_id == "HDBANK":
            return f"HDB-Báo có GD QR"
        return f"{merchant_name}-{channel_text}{business}"
            
    elif payment_type == "FEE":
        return f"HDB-MPGS Phí GD{business}"

    elif payment_type == "TRANSFER":
        if business == " KH-RVS":
            return f"Ch.tien UPC-KH Riverside"
        
        return f"Chuyển tiền {channel_text}{business}".replace("GD UPC", "UPC")
    elif payment_type == "DEPOSITS_INTEREST":
        return f"Lãi tiền gửi"
    elif payment_type == "PAYMENT_INTERNAL":
        return f"Chuyển tiền nội bộ {note}"
    return f"{merchant_name}-{payment_type} {merchant_id} {recon_result} {note}"

def build_title_bvb(
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
    
    if "VIETJET" in merchant_id:
        merchant_id = "VIETJET"
    elif "VSSS" in merchant_id:
        merchant_id = "VSSS"
    match merchant_name:
        case "BVBANK":
            merchant_name = "BVB"
        
    # =======================
    # Mapping Merchant ID
    # =======================
    merchant_mapping = {
        
        "UPC Vietjet Air": "VIETJET",
        "Galaxy Telecom": "SKYFI",
        "Galaxy Joy": "GJoy",
        "UPC MPGS Vietjet Air": "VIETJET",
        "HDI": "HDI",
        "Rong Viet Travel": "Rồng Việt",
        "GPay": "QR",
        "ONAIR": "OnAir",
        "VIETJET": "Vietjet",
        "Truong Khai Hoan-NAM SG": "Khải Hoàn-NSG",
        "Truong Khai HoanRIVERSIDE": "KH-RVS",
        "Truong Khai Hoan-RIVERSIDE": "KH-RVS",
        "VSSS": "VSSS",
    }

    business = merchant_mapping.get(merchant_id, merchant_id)

    # =======================
    # Channel mapping
    # =======================
    channel_mapping = {
        "MOMO": "BC DV",
        "UPC Vietjet Air": "UPC",
        "UPC MPGS Vietjet Air": "UPC",
        "Galaxy Telecom": "UPC",
        "Galaxy Joy": "GD UPC",
        "HDI": "GD UPC",
        "Rong Viet Travel": "-RồngViệt",

    }

    channel = channel_mapping.get(
        merchant,
            "BC GD"
        )
    channel_text = f"{channel}" if channel else ""   

    if not business.startswith("-"):
            business = f" {business}"
    # ============================================
    # PAYMENT_RECEIVED
    # ============================================
    if payment_type == "PAYMENT":
        
        return f"{merchant_name}-{channel_text}{business}"
            
    elif payment_type == "FEE":
        return f"BVB-Phí chuyển tiền"

    elif payment_type == "TRANSFER":
        if business == " KH-RVS":
            return f"Ch.tiền QR-KH Riverside"
        return f"Chuyển tiền QR{business}"
    
    elif payment_type == "PAYMENT_INTEREST":
        return f"Lãi tiền gửi"
    elif payment_type == "REFUND":
        return f"BVB-Cắt tiền H.Lệch QR VJ"
    
    return f"{merchant_name}-{payment_type} {merchant_id} {recon_result} {note}"


def build_description(title, posting_date, accounting_date, merchant_id, merchant_name, file_name, note, payment_type):
    
    #posting_date
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
    
    posting_date_1 = get_posting_date_minus_1(posting_date)
    
    #description
    
    if "MPGS-GD UPC" in title:
        title = title.replace("UPC ", "")
    # if "BMI - ABBANK" in title:
    #     title = title.replace("BMGĐ", "BAO MINH GIA DINH")
    if "GMBR" in title:
        title = title.replace("GMBR", "Giấc Mơ Bay Rẻ")
    
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
        elif "Lệch QR Nạp TK Ví" in title:
            rtitle = f"{rtitle.replace("Lệch QR Nạp TK Ví ", "Lệch Nạp tiền vào tài khoản Ví ")}"
        elif "Lệch Nạp tiền Ví" in title:
            rtitle = f"{rtitle.replace("Lệch Nạp tiền Ví ", "Lệch Nạp tiền vào tài khoản Ví ")}"
    
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
            if "QR Nạp" in title:
                rtitle = f"{title.replace("QR Nạp TK Ví", "Nạp tiền vào Ví qua QR Code")} N{posting_date}"
            elif "Nạp tiền TK Ví GPay" in title:
                rtitle = f"{title.replace("Nạp tiền TK Ví GPay", "Nạp tiền vào Ví")} N{posting_date}"
            else:
                # Trước đây là `elif "Rút tiền":` (luôn True) nên nhánh
                # " TK Ví" -> " vào tài khoản ví" chưa bao giờ chạy.
                # Giữ nguyên output đang dùng: "Napas-Nạp tiền TK Ví", "BIDV-Rút tiền từ TK Ví GPay"
                rtitle = f"{title.replace("Rút tiền TK Ví", "Rút tiền từ TK Ví")} N{posting_date}"
        elif merchant_id == "BAO HIEM HD - PORTAL":
            rtitle = f"{title.replace("HDI", "HDI-Portal")} N{posting_date}"
        else:
            rtitle = f"{title} N{posting_date}"
    
    
    #=============HDBANK_STATEMENT=============
    if "HDB-Báo có GD QR" in title:
        rtitle = f"{title} N{posting_date}" 
    elif ("BC GD" in title or "BC DV UPC" in title) and payment_type == "PAYMENT_RECEIVED":
        rtitle = f"{title} N{posting_date_1}"
    elif "MPGS Phí" in title:
        rtitle = title
    elif "Chuyển tiền" in title or "Ch.tien" in title:
        if "HDB" in merchant_name:
            merchant = extract_hdbank_merchant(note.upper())

        elif "BVB" in merchant_name:
            merchant = extract_bvbank_merchant(note.upper())

        else:
            merchant = None

        merchant = mapping_merchant(merchant, merchant_id)["merchant_map"]
        if merchant:
            note_upper = note.upper()

            merchant_pos = note_upper.find(merchant.upper())

            if merchant_pos != -1:
                # Chỉ lấy phần phía sau merchant
                rtitle= convert_payment_info_to_output(note)
               #rtitle = f"{title} {date_map}"


            else:
                rtitle = "ERROR_DATE_EXTRACT"

        else:
            rtitle = "ERROR_DATE_EXTRACT"
        
        
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
        elif ("topup" in title or "Billing" in title) and "Hoàn Lệch" in title:
            rtitle = f"{title.replace("Hoàn Lệch ", "Hoàn Lệch GD ")}"
            # print(f"đổi tên IMEDIA Hoàn Lệch GD ============++++++++")
        elif "Khải Hoàn" in title:
            rtitle = f"{title.replace(" QR ", " QR code ")} N{posting_date}"
        elif "Hoàn Lệch Rút tiền Ví(GP)" in title:
            rtitle = title.replace("Hoàn Lệch Rút tiền Ví(GP)", "Xử lý Lệch Rút tiền TK Ví (GP)")
        elif "Xử lý Lệch" in title or "Napas-Hoàn Lệch" in title:
            if "Napas" in title:
                if "Nạp Ví" in title:
                    # print("z"*60)
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
    
    if "Lãi tiền gửi" in title:
        rtitle = title
    
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
    
    if "Phát sinh có" in note:
        document_type = "DZ"
    elif "Phát sinh nợ" in note:
        document_type = "KZ"
    else:
        document_type = "CL"
    # ===========================
    # Default
    # ===========================
    if "DV IBFT Hoàn" in title and note == "MISS_REFUND_FALSE":
        document_type = "CL"
        
    #print(debit[0:2], credit[0:2])
    # ===========================
    # Rule 1
    if debit == "NONE" or credit == "NONE":
        return "CL"
    if debit[0:2] in ["13", "33"] and credit[0:2] in ["13", "33"]:
        document_type = "DK"
    # ===========================

    elif debit.startswith("33") and credit.startswith("33"):
        document_type = "KZ"
    elif debit.startswith("11") and credit.startswith("33"):
        document_type = "KZ"
    # ===========================
    # Rule 2
    # ===========================

    elif debit.startswith("131") and credit.startswith("511"):
        document_type = "RV"

    elif debit.startswith("6") or credit.startswith("6"):
        document_type = "SA"
   
    # ===========================
    # Rule 3
    # ===========================

    # if "HOÀN" in title.upper():
    #     document_type = "KA"
    
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
            except ValueError:
                raise ValueError(
                    f"Cannot parse amount '{amount}' in file {file_name}, row: {row}"
                )

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
            if "RECON_REPORT_INPUT" in file_name.upper():
                title = build_title_upc(merchant,payment_type,merchant_id,recon_result,note)
            elif "DAILY" in file_name.upper() or WALLET_NEW_FILE_PREFIX in file_name.upper():
                title = build_title_wallet(merchant,payment_type,merchant_id,recon_result,note)
            elif "HDBANK" in file_name.upper():
                title = build_title_hdbank(merchant,payment_type,merchant_id,recon_result,note)
            elif "SAOKE" in file_name.upper():
                title = build_title_bvb(merchant,payment_type,merchant_id,recon_result,note)
            else:
                # Không fallback: tránh dùng lại title của dòng trước
                raise ValueError(f"Unknown file type, cannot build title: {file_name}")
            if "XLý Lệch Nạp tiền Ví" in title or "HLệch Nạp Ví (ĐT)" in title or "XLLệch Nạp tiền Ví" in title or "HLệch Rút tiền Ví" in title:
                accounting_date = parse_posting_date_from_filename(file_name)
            
            # ===== DIỄN GIẢI (FIXED) =====
            description = build_description(title, posting_date, accounting_date, merchant_id, merchant, file_name, note, payment_type)
            if "MPGS" in merchant.upper() or merchant == "HDBANK_STATEMENT" or "HDBANK" in file_name.upper() or "SAOKE_BVBANK" in file_name.upper():
                debit_info = build_posting_info_mpgs(title, "Nợ", merchant, merchant_id, note, payment_type, recon_result, file_name, amount)
                credit_info = build_posting_info_mpgs(title, "Có", merchant, merchant_id, note, payment_type, recon_result, file_name, amount)
            else:
                debit_info = build_posting_info(title, "Nợ", merchant, merchant_id, note, payment_type, recon_result, file_name)
                credit_info = build_posting_info(title, "Có", merchant, merchant_id, note, payment_type, recon_result, file_name)

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
                if "MPGS" in merchant.upper() or merchant == "HDBANK_STATEMENT" or "HDBANK" in file_name.upper() or "SAOKE_BVBANK" in file_name.upper():
                    posting_info = build_posting_info_mpgs(title, dr_cr, merchant, merchant_id, note, payment_type, recon_result, file_name, amount)
                else:
                    posting_info = build_posting_info(title, dr_cr, merchant, merchant_id, note, payment_type, recon_result, file_name)
                
                if title == "Lãi tiền gửi" and dr_cr == "Có":
                    dummy = "DUMMY"
                else:
                    dummy = ""
                
                if posting_info["account"].startswith("63") or posting_info["account"].startswith("644"):
                    cost_center = "8001601000"
                else:
                    cost_center = ""
                
                
                if amount == 0 or payment_type == "PAYMENT_INTERNAL" or "BVB-Cắt tiền" in title:
                    print(f"[WARN] Amount is 0 or PAYMENT_INTERNAL, skipping row: {row}")
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

                        "TT Chi phí (Cost center)": cost_center,
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
                        "Chặng bay (Nhập cho TK Doanh thu)": dummy,

                        "Điều khoản thanh toán": dummy,
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
                        "Country(Khai báo hóa đơn cho NCC vãng lai/ Payee)": "",
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
                    

    if not rows_out:
        raise ValueError(f"No rows to export from {INPUT_FILE}")

    # Sort theo ngày thật (dd/mm/yyyy sort theo chuỗi sẽ sai thứ tự giữa các tháng).
    # Sort ổn định nên cặp Nợ/Có của cùng 1 chứng từ vẫn đứng liền nhau.
    def parse_accounting_date(row):
        try:
            return datetime.strptime(row.get("Ngày hạch toán", "").strip(), "%d/%m/%Y")
        except ValueError:
            return datetime.max

    rows_out.sort(key=parse_accounting_date)

    # Đánh STT sau khi sort: mỗi cặp Nợ/Có chung 1 STT
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