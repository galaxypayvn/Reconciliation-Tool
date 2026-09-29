
import re
import csv
from datetime import datetime, time, timedelta
import calendar
from pathlib import Path
from openpyxl import Workbook

def extract_hdbank_merchant(payment_info):
    """
    Extract merchant name from Galaxy Pay transfer descriptions.

    Supported examples:
        Galaxy Pay chuyen tien cac giao dich UPC Vietjet Air N15/06/2026 (GD ND)
            -> UPC Vietjet Air

        Galaxy Pay chuyen tien cac giao dich HDI N12/06 den N14/06/2026
            -> HDI

        Galaxy Pay chuyen tien cac GD HDI N01/06/2026
            -> HDI

        Galaxy Pay chuyen tien VSSS (Victoria International School) N12/06 den N14/06/2026
            -> VSSS (Victoria International School)

        Galaxy Pay chuyen tien cac giao dich Galaxy Telecom tu N05/06 den N07/06/2026
            -> Galaxy Telecom
    """

    patterns = [
        # ============================================================
        # 1. Galaxy Pay chuyen tien cac giao dich <merchant>
        #    Galaxy Pay chuyen tien cac GD <merchant>
        #
        # Examples:
        #   Galaxy Pay chuyen tien cac giao dich HDI N12/06 den N14/06/2026
        #   Galaxy Pay chuyen tien cac giao dich Galaxy Telecom N15/06/2026
        #   Galaxy Pay chuyen tien cac giao dich Galaxy Telecom tu N05/06 den N07/06/2026
        #   Galaxy Pay chuyen tien cac giao dich UPC MPGS Vietjet Air
        #       tu N07/08 den 09/08/2026 (GD QT)
        # ============================================================
        (
            r"Galaxy\s+Pay\s+chuyen\s+tien\s+cac\s+"
            r"(?:giao\s+dich|GD)"
            r"\s+(.*?)\s+"
            r"(?:"
                # tu N07/08 den 09/08/2026
                r"tu\s+N\d{1,4}/\d{1,2}"
                r"\s+den\s+N?\d{1,2}/\d{1,2}"
                r"|"
                # N07/08 den N09/08/2026
                r"N\d{1,4}/\d{1,2}"
                r"(?:\s+(?:den|va)\s+N?\d{1,2}/\d{1,2})?"
                r"(?:\s+(?:den|va)\s+N?\d{1,2}/\d{1,2})?"
            r")"
            r"/\d{4}"
        ),

        # ============================================================
        # 2. Galaxy Pay chuyen tien <merchant>
        #
        # Examples:
        #   Galaxy Pay chuyen tien VSSS (Victoria International School)
        #       N12/06 den N14/06/2026
        #
        #   Galaxy Pay chuyen tien Truong Khai HoanNSG N16/06/2026
        # ============================================================
        (
            r"Galaxy\s+Pay\s+chuyen\s+tien\s+"
            r"(.*?)\s+"
            r"N\d{1,4}/\d{1,2}"
            r"(?:\s+(?:den|va)\s+N\d{1,2}/\d{1,2})?"
            r"(?:\s+(?:den|va)\s+N\d{1,2}/\d{1,2})?"
            r"/\d{4}"
        ),
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            payment_info,
            re.IGNORECASE
        )

        if match:
            return match.group(1).strip()

    return "CAN_NOT_EXTRACT_MERCHANT_ID"
def extract_bvbank_merchant(payment_info):
    """
    Extract merchant name from Galaxy Pay transfer descriptions.

    Supported examples:
        Galaxy Pay chuyen tien cac GD QR Code VIETJET N31/05/2026
            -> VIETJET

        Galaxy Pay chuyen tien cac GD QR Code VIETJET N29-30/05/2026
            -> VIETJET

       Galaxy Pay chuyen tien VSSS Victoria International School N29-31/05/2026 GD QR va bu tru phi T04/2026
            -> VSSS Victoria International School

        Galaxy Pay chuyen tien cac giao dich ONAIR N11/06/2026 GD QR
            -> ONAIR

        Galaxy Pay chuyen tien cac  GD QR Code Galaxy Telecom N29-31/05/2026
            -> Galaxy Telecom
        
        Galaxy Pay chuyen tien cac GD QR CODE Truong Khai Hoan-NAM SG N10/062026
            -> Truong Khai Hoan-NAM SG

    """

    patterns = [
        # Galaxy Pay chuyen tien cac giao dich <merchant> Ndd/mm/yyyy
        # Example:
        # Galaxy Pay chuyen tien cac giao dich ONAIR N11/06/2026
        (
            r"Galaxy\s+Pay\s+chuyen\s+tien\s+cac\s+"
            r"giao\s+dich\s+"
            r"(.*?)\s+"
            r"N\d{1,2}(?:-\d{1,2})?"
            r"/\d{1,2}/?\d{4}"
        ),

        # Galaxy Pay chuyen tien cac GD QR Code <merchant> Ndd/mm/yyyy
        # Example:
        # Galaxy Pay chuyen tien cac GD QR Code VIETJET N31/05/2026
        #
        # QR Code is optional
        # so it also supports:
        # Galaxy Pay chuyen tien cac GD Galaxy Telecom N29-31/05/2026
        (
            r"Galaxy\s+Pay\s+chuyen\s+tien\s+cac\s+"
            r"GD(?:\s+QR\s+Code)?\s+"
            r"(.*?)\s+"
            r"N\d{1,2}(?:-\d{1,2})?"
            r"/\d{1,2}/?\d{4}"
        ),

        # Galaxy Pay chuyen tien <merchant> Ndd/mm/yyyy
        # Example:
        # Galaxy Pay chuyen tien VSSS Victoria International School N29-31/05/2026
        (
            r"Galaxy\s+Pay\s+chuyen\s+tien\s+"
            r"(.*?)\s+"
            r"N\d{1,2}(?:-\d{1,2})?"
            r"/\d{1,2}/?\d{4}"
        ),
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            payment_info,
            re.IGNORECASE
        )

        if match:
            return match.group(1).strip()

    return "CAN_NOT_EXTRACT_MERCHANT_ID"

def get_posting_date_minus_1(posting_date: str) -> str:
    """
    Tạo posting_date_1 từ posting_date.

    Rule:
        - posting_date là 1 ngày T
            -> posting_date_1 = T-1

        - posting_date là khoảng x-T
            -> posting_date_1 = x-(T-1)

    Examples:
        "05/08/2026"
            -> "04/08/2026"

        "01/08/2026"
            -> "31/07/2026"

        "03-05/08/2026"
            -> "03-04/08/2026"

        "29-31/07/2026"
            -> "29-30/07/2026"

        "01-03/08/2026"
            -> "01-02/08/2026"
    """

    if not posting_date:
        return posting_date

    posting_date = posting_date.strip()

    # ==========================================
    # CASE 1: Khoảng ngày
    # Ví dụ:
    # 03-05/08/2026
    # ==========================================
    range_match = re.fullmatch(
        r"(\d{1,2})-(\d{1,2})/(\d{1,2})/(\d{4})",
        posting_date
    )

    if range_match:
        start_day = int(range_match.group(1))
        end_day = int(range_match.group(2))
        month = int(range_match.group(3))
        year = int(range_match.group(4))

        try:
            end_date = datetime(
                year,
                month,
                end_day
            )

            new_end_date = end_date - timedelta(days=1)

            # Nếu sau khi trừ 1 ngày mà về tháng trước
            if new_end_date.month != month:
                return (
                    f"{start_day:02d}/{month:02d}/{year}"
                )

            return (
                f"{start_day:02d}-"
                f"{new_end_date.day:02d}/"
                f"{month:02d}/"
                f"{year}"
            )

        except ValueError:
            return posting_date

    # ==========================================
    # CASE 2: Một ngày
    # Ví dụ:
    # 05/08/2026
    # ==========================================
    single_match = re.fullmatch(
        r"(\d{1,2})/(\d{1,2})/(\d{4})",
        posting_date
    )

    if single_match:
        day = int(single_match.group(1))
        month = int(single_match.group(2))
        year = int(single_match.group(3))

        try:
            posting_dt = datetime(
                year,
                month,
                day
            )

            previous_date = posting_dt - timedelta(days=1)

            return previous_date.strftime("%d/%m/%Y")

        except ValueError:
            return posting_date

    # Không đúng format
    return posting_date


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

def convert_payment_info_to_output(payment_info: str) -> str:
    """
    Convert payment_info từ format ngân hàng sang format output.

    INPUT examples:

        Ch.tien UPC-KH Riverside N03/09/2026 (GD QT)

        Chuyển tiền BC GD SKYFI N01/09/2026 (GD QT)
        va N01/09 den N02/09/2026 (GD ND)

        Chuyển tiền BC GD-GJoy N03/09/2026 (GD QT)
        va N0406/09/2026 (GD QR, ND)
        va can tru GD BVB theo yeu cau hoan N28/08/2026

    OUTPUT:

        Chuyển tiền Khải Hoàn-Riverside N03/09/2026(QT)

        Chuyển tiền SKYFI N01/09/2026(QT);N01-02/09/2026(NĐ)

        Chuyển tiền GJoy N03/09/2026(QT);N04-06/09/2026(QR,NĐ)

    Không phụ thuộc vào ngày cụ thể.
    """

    if not payment_info:
        return ""

    # ==========================================================
    # 1. Normalize whitespace
    # ==========================================================
    text = re.sub(r"\s+", " ", payment_info).strip()

    # ==========================================================
    # 2. Xác định merchant
    # ==========================================================
    merchant = None

    merchant_patterns = [
        # Riverside
        (
            r"(?:UPC-KH\s+)?Riverside",
            "Khải Hoàn-Riverside"
        ),

        # Khải Hoàn-Riverside
        (
            r"Khải Hoàn-Riverside",
            "Khải Hoàn-Riverside"
        ),

        # Khải Hoàn-NSG
        (
            r"Khải Hoàn-NSG",
            "Khải Hoàn-NSG"
        ),

        # GJoy
        (
            r"Galaxy Joy",
            "GJoy"
        ),

        # HDI
        (
            r"HDI(?:-Portal)?",
            None
        ),

        # SKYFI
        (
            r"SKYFI",
            "SKYFI"
        ),

        # Vikkafe
        (
            r"Vikkafe",
            "Vikkafe"
        ),

        # VIETJET
        (
            r"VIETJET Air",
            "VIETJET"
        ),
        (
            r"VIETJET",
            "VIETJET"
        ),
        

        # VSSS
        (
            r"VSSS",
            "VSSS"
        ),

        # Menas
        (
            r"Menas",
            "Menas"
        ),

        # OnAir
        (
            r"OnAir",
            "OnAir"
        ),

        # Giấc Mơ Bay Rẻ
        (
            r"Giấc Mơ Bay Rẻ",
            "Giấc Mơ Bay Rẻ"
        ),
        # Giấc Mơ Bay Rẻ
        (
            r"Giac Mo Bay Re",
            "Giac Mo Bay Re"
        ),

        # Galaxy Telecom
        (
            r"Galaxy Telecom",
            "Skyfi"
        ),

        # NSG
        (
            r"NSG",
            "Khải Hoàn-NSG"
        ),

        # RồngViệt
        (
            r"RồngViệt",
            "RồngViệt"
        ),
        # RồngViệt
        (
            r"Rong Viet Travel",
            "Rong Viet Travel"
        ),
    ]

    for pattern, output_name in merchant_patterns:
        match = re.search(pattern, text, re.IGNORECASE)

        if match:
            merchant = output_name or match.group(0)
            break

    if not merchant:
        return f"Chuyển tiền {text}"

    # ==========================================================
    # 3. Lấy phần phía sau merchant
    # ==========================================================
    merchant_match = re.search(
        re.escape(match.group(0)),
        text,
        re.IGNORECASE
    )

    if not merchant_match:
        return f"Chuyển tiền {merchant}"

    date_part = text[merchant_match.end():].strip()

    # ==========================================================
    # 4. Xử lý các range ngày
    #
    # N04/09 den N06/09/2026
    # N04/09 đến N06/09/2026
    # N04/09 den N06/09/26
    #
    # -> N04-06/09/2026
    # ==========================================================
    date_part = re.sub(
        r"N(\d{1,2})/(\d{1,2})"
        r"\s+(?:den|đến)\s+"
        r"N(\d{1,2})/(\d{1,2})/(\d{2,4})",
        lambda m: (
            f"N{int(m.group(1)):02d}-{int(m.group(3)):02d}"
            f"/{int(m.group(2)):02d}/{normalize_year(m.group(5))}"
        ),
        date_part,
        flags=re.IGNORECASE
    )

    # ==========================================================
    # 5. Xử lý trường hợp:
    #
    # N01/09 den N02/09/2026
    #
    # đã xử lý ở trên.
    #
    # Xử lý:
    # N27/08 den N31/08/2026
    # ==========================================================

    # ==========================================================
    # 6. Chuẩn hóa ngày dạng N0406/09/2026
    #
    # INPUT:
    # N0406/09/2026
    #
    # OUTPUT:
    # N04-06/09/2026
    # ==========================================================
    date_part = re.sub(
        r"N(\d{2})(\d{2})/(\d{1,2})/(\d{2,4})",
        lambda m: (
            f"N{m.group(1)}-{m.group(2)}/"
            f"{int(m.group(3)):02d}/{normalize_year(m.group(4))}"
        ),
        date_part
    )

    # ==========================================================
    # 7. Chuẩn hóa ngày có year 2 digits
    #
    # N04/06/26 -> N04/06/2026
    # N18/06/26 -> N18/06/2026
    # ==========================================================
    date_part = re.sub(
        r"N(\d{1,2})/(\d{1,2})/(\d{2})",
        lambda m: (
            f"N{int(m.group(1)):02d}/"
            f"{int(m.group(2)):02d}/"
            f"{normalize_year(m.group(3))}"
        ),
        date_part
    )

    # ==========================================================
    # 8. Chuẩn hóa ngày bình thường
    #
    # N3/9/2026 -> N03/09/2026
    # ==========================================================
    date_part = re.sub(
        r"N(\d{1,2})/(\d{1,2})/(\d{4})",
        lambda m: (
            f"N{int(m.group(1)):02d}/"
            f"{int(m.group(2)):02d}/{m.group(3)}"
        ),
        date_part
    )

    # ==========================================================
    # 9. Chuẩn hóa "(GD QT)"
    #
    # (GD QT) -> (QT)
    # (GD ND) -> (NĐ)
    # (GD QR) -> (QR)
    # ==========================================================
    date_part = re.sub(
        r"\(\s*GD\s+QT\s*\)",
        "(QT)",
        date_part,
        flags=re.IGNORECASE
    )

    date_part = re.sub(
        r"\(\s*GD\s+ND\s*\)",
        "(NĐ)",
        date_part,
        flags=re.IGNORECASE
    )

    date_part = re.sub(
        r"\(\s*GD\s+QR\s*\)",
        "(QR)",
        date_part,
        flags=re.IGNORECASE
    )

    # ==========================================================
    # 10. Chuẩn hóa "(GD QR, ND)"
    #
    # (GD QR, ND) -> (QR,NĐ)
    # ==========================================================
    date_part = re.sub(
        r"\(\s*GD\s+QR\s*,\s*ND\s*\)",
        "(QR,NĐ)",
        date_part,
        flags=re.IGNORECASE
    )

    date_part = re.sub(
        r"\(\s*GD\s+QR\s*,\s*ND\s*\)",
        "(QR,NĐ)",
        date_part,
        flags=re.IGNORECASE
    )

    # ==========================================================
    # 11. Xử lý "va" giữa các transaction
    #
    # N01/06/2026(QT) va N02/06/2026(NĐ)
    #
    # ->
    # N01/06/2026(QT);N02/06/2026(NĐ)
    # ==========================================================
    date_part = re.sub(
        r"\)\s+(?:va|và)\s+(?=N)",
        ");",
        date_part,
        flags=re.IGNORECASE
    )

    # ==========================================================
    # 12. Xử lý trường hợp:
    #
    # N01/09/2026 (GD QT)
    # va tu N11/09 den N13/09/2026 (GD ND)
    #
    # ==========================================================
    date_part = re.sub(
        r"\)\s+(?:va|và)\s+(?:tu|từ)\s+(?=N)",
        ");",
        date_part,
        flags=re.IGNORECASE
    )

    # ==========================================================
    # 13. Chuẩn hóa "tu N07/08 den N09/08/2026"
    #
    # Trường hợp này có thể chưa được bắt nếu phía trước
    # không có year.
    # ==========================================================
    date_part = re.sub(
        r"(?:tu|từ)\s+"
        r"N(\d{1,2})/(\d{1,2})"
        r"\s+(?:den|đến)\s+"
        r"N(\d{1,2})/(\d{1,2})/(\d{2,4})",
        lambda m: (
            f"N{int(m.group(1)):02d}-{int(m.group(3)):02d}"
            f"/{int(m.group(2)):02d}/{normalize_year(m.group(5))}"
        ),
        date_part,
        flags=re.IGNORECASE
    )

    # ==========================================================
    # 14. Loại bỏ khoảng trắng thừa
    # ==========================================================
    date_part = re.sub(r"\s+", " ", date_part).strip()

    # ==========================================================
    # 15. Loại bỏ khoảng trắng trước "(QT)", "(NĐ)", "(QR)"
    #
    # N01/06/2026 (QT)
    # -> N01/06/2026(QT)
    # ==========================================================
    date_part = re.sub(
        r"\s+\((QT|NĐ|QR)(?:,NĐ)?\)",
        r"(\1)",
        date_part
    )

    # ==========================================================
    # 16. Xử lý các phần diễn giải không phải transaction
    #
    # Ví dụ:
    #   va can tru GD phi dich vu T07/2026
    #
    # Output:
    #   CT VSSS N08/09/2026(QT);trừ phí T07/2026
    # ==========================================================

    has_non_transaction_note = False

    # ----------------------------------------------------------
    # 16.1. Cấn trừ phí dịch vụ
    # ----------------------------------------------------------
    date_part = re.sub(
        r"\s+(?:va|và)\s+can\s+tru\s+GD\s+phi\s+dich\s+vu\s+T(\d{1,2})/(\d{4})",
        lambda m: f";trừ phí T{m.group(1)}/{m.group(2)}",
        date_part,
        flags=re.IGNORECASE
    )

    if re.search(r";trừ phí T\d{1,2}/\d{4}", date_part, flags=re.IGNORECASE):
        has_non_transaction_note = True


    # ----------------------------------------------------------
    # 16.2. Nếu có diễn giải "cấn trừ" khác
    #     Ví dụ:
    #     va can tru GD BVB theo yeu cau hoan N28/08/2026
    #
    #     Có thể giữ nguyên phần diễn giải sau khi transaction.
    # ----------------------------------------------------------
    date_part = re.sub(
        r"\s+(?:va|và)\s+(?=can\s+tru|cấn\s+trừ)",
        ";",
        date_part,
        flags=re.IGNORECASE
    )


    # ----------------------------------------------------------
    # 16.3. Đổi prefix nếu có phần diễn giải không phải transaction
    # ----------------------------------------------------------
    if has_non_transaction_note:
        output_prefix = "CT"
    else:
        output_prefix = "Chuyển tiền"


    # ==========================================================
    # 17. Final
    # ==========================================================
    date_part = date_part.strip(" ;")

    return f"Chuyển tiền {merchant} {date_part}".strip()


def normalize_year(year: str) -> str:
    """
    26 -> 2026
    25 -> 2025
    2026 -> 2026
    """

    if len(year) == 2:
        return f"20{year}"

    return year

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

    # Các column được xem là ngày
    DATE_COLUMNS = {
        "Ngày giao dịch",
        "Ngày hạch toán",
        "Ngày giao dịch gốc",
        "Ngày chứng từ",
    }

    # Các column được xem là số tiền
    AMOUNT_COLUMNS = {
        "Số tiền",
        "Số dư",
        "Debit",
        "Credit",
        "Balance",
        "Số tiền (Bank)",
        "Số tiền (Payment)",
    }

    # ============================================================
    # Đọc CSV
    # ============================================================

    with open(
        input_file,
        "r",
        encoding="utf-8-sig",
        newline=""
    ) as f:

        reader = csv.reader(f)

        # --------------------------------------------------------
        # Đọc header
        # --------------------------------------------------------

        try:
            headers = next(reader)
        except StopIteration:
            print(f"[ERROR] CSV file is empty: {input_file}")
            return

        # Ghi header
        for col_index, header in enumerate(headers, start=1):

            cell = worksheet.cell(
                row=1,
                column=col_index
            )

            cell.value = header
            cell.number_format = "@"

        # --------------------------------------------------------
        # Đọc data
        # --------------------------------------------------------

        for row_index, row in enumerate(reader, start=2):

            for col_index, value in enumerate(row, start=1):

                header = headers[col_index - 1].strip()

                cell = worksheet.cell(
                    row=row_index,
                    column=col_index
                )

                value = value.strip()

                # ====================================================
                # COLUMN DATE
                # ====================================================

                if header in DATE_COLUMNS:

                    try:

                        # Hỗ trợ:
                        # 31/05/2026
                        # 31/05/2026 00:00:00
                        # 2026-05-31

                        date_value = None

                        for date_format in (
                            "%d/%m/%Y",
                            "%d/%m/%Y %H:%M:%S",
                            "%Y-%m-%d",
                            "%Y-%m-%d %H:%M:%S",
                        ):

                            try:
                                date_value = datetime.strptime(
                                    value,
                                    date_format
                                )
                                break

                            except ValueError:
                                continue

                        if date_value:

                            cell.value = date_value
                            cell.number_format = "dd/mm/yyyy"

                        else:
                            cell.value = value
                            cell.number_format = "@"

                    except Exception:
                        cell.value = value
                        cell.number_format = "@"

                # ====================================================
                # COLUMN AMOUNT
                # ====================================================

                elif header in AMOUNT_COLUMNS:

                    try:

                        # Remove thousand separator
                        numeric_value = value.replace(",", "")

                        if numeric_value == "":
                            cell.value = None

                        else:
                            cell.value = float(numeric_value)

                        cell.number_format = "#,##0.00"

                    except (ValueError, TypeError):

                        # Nếu không convert được thì giữ nguyên text
                        cell.value = value
                        cell.number_format = "@"

                # ====================================================
                # OTHER COLUMN
                # ====================================================

                else:

                    cell.value = value
                    cell.number_format = "@"

    # ============================================================
    # Auto width
    # ============================================================

    for column in worksheet.columns:

        max_length = 0
        column_letter = column[0].column_letter

        for cell in column:

            if cell.value is not None:

                cell_length = len(str(cell.value))

                if cell_length > max_length:
                    max_length = cell_length

        worksheet.column_dimensions[column_letter].width = min(
            max_length + 2,
            50
        )

    # ============================================================
    # Save
    # ============================================================

    workbook.save(output_file)

    print("[INFO] Converted:")
    print(f"       CSV : {input_file}")
    print(f"       XLSX: {output_file}")
    

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