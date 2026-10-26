#!/usr/bin/env python3
"""
Unit test cho các bug đã sửa.

Usage:
    python3 -m unittest discover -s tests -v
"""

import csv
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import convert_data_to_json
import export_upload_file_1
import run_reconciliation
from get_data_4 import process_file_data
from helper import convert_payment_info_to_output


class ConvertPaymentInfoTest(unittest.TestCase):

    def test_four_digit_year_is_not_corrupted(self):
        self.assertEqual(
            convert_payment_info_to_output(
                "Galaxy Pay chuyen tien cac giao dich UPC Vietjet Air N15/06/2026 (GD ND)"
            ),
            "Chuyển tiền VIETJET N15/06/2026(NĐ)",
        )

    def test_two_digit_year_is_normalized(self):
        self.assertEqual(
            convert_payment_info_to_output(
                "Galaxy Pay chuyen tien cac giao dich HDI N04/06/26"
            ),
            "Chuyển tiền HDI N04/06/2026",
        )

    def test_multiple_transactions(self):
        self.assertEqual(
            convert_payment_info_to_output(
                "Galaxy Pay chuyen tien cac giao dich Galaxy Telecom N01/09/2026 (GD QT) "
                "va tu N11/09 den N13/09/2026 (GD ND)"
            ),
            "Chuyển tiền Skyfi N01/09/2026(QT);N11-13/09/2026(NĐ)",
        )

    def test_qr_nd_suffix_is_kept(self):
        self.assertEqual(
            convert_payment_info_to_output(
                "Galaxy Pay chuyen tien cac giao dich Galaxy Joy N03/09/2026 (GD QT) "
                "va N0406/09/2026 (GD QR, ND)"
            ),
            "Chuyển tiền GJoy N03/09/2026(QT);N04-06/09/2026(QR,NĐ)",
        )

    def test_fee_offset_uses_ct_prefix(self):
        self.assertEqual(
            convert_payment_info_to_output(
                "Galaxy Pay chuyen tien VSSS N08/09/2026 (GD QT) "
                "va can tru GD phi dich vu T07/2026"
            ),
            "CT VSSS N08/09/2026(QT);trừ phí T07/2026",
        )


BVB_FILE = "SAOKE_BVBank_0697041105669_T06_Sheet1.csv"
BVB_DATE = "NGÀY GIAO DỊCH\n(Transaction Date)"
BVB_INFO = "DIỄN GIẢI\n(Narrative)"
BVB_DEBIT = "PHÁT SINH NỢ\n(Debit Amount)"
BVB_CREDIT = "PHÁT SINH CÓ\n(Credit Amount)"


def bvb_row(info, debit=0, credit=0, refer="", date="10/06/2026"):
    return {
        BVB_DATE: date,
        BVB_INFO: info,
        BVB_DEBIT: str(debit),
        BVB_CREDIT: str(credit),
        "Số tham chiếu": refer,
    }


def flatten(summary):
    rows = []
    for merchant_ids in summary.values():
        for merchant_id, payment_types in merchant_ids.items():
            for payment_type, dates in payment_types.items():
                for results in dates.values():
                    for infos in results.values():
                        for info in infos:
                            rows.append((merchant_id, payment_type, info))
    return rows


class BvbankTest(unittest.TestCase):

    TRANSFER = "Galaxy Pay chuyen tien cac GD QR Code VIETJET N09/06/2026"

    def run_bvb(self, rows):
        with mock.patch("builtins.print"):
            return flatten(process_file_data(BVB_FILE, rows))

    def test_transfer_with_fee_row(self):
        result = self.run_bvb([
            bvb_row(self.TRANSFER, debit=1_000_000, refer="111"),
            bvb_row(self.TRANSFER, debit=5_500, refer="112"),
            bvb_row("Other", credit=1, refer="999"),
        ])

        transfers = [r for r in result if r[1] == "TRANSFER"]
        fees = [r for r in result if r[1] == "FEE"]

        self.assertEqual(len(transfers), 1)
        self.assertEqual(transfers[0][2]["total_amount"], 1_000_000)
        self.assertEqual(transfers[0][2]["Số tham chiếu"], "111")

        self.assertEqual(len(fees), 1)
        self.assertEqual(fees[0][2]["total_amount"], 5_500)
        self.assertEqual(fees[0][2]["Số tham chiếu"], "112")

    def test_transfer_without_fee_row_is_not_dropped(self):
        result = self.run_bvb([
            bvb_row(self.TRANSFER, debit=1_000_000, refer="111"),
            bvb_row("Other", credit=1, refer="999"),
        ])

        transfers = [r for r in result if r[1] == "TRANSFER"]

        self.assertEqual(len(transfers), 1)
        self.assertEqual(transfers[0][2]["total_amount"], 1_000_000)
        self.assertEqual(transfers[0][2]["Số tham chiếu"], "111")
        self.assertFalse([r for r in result if r[1] == "FEE"])

    # BUG đã biết, CHƯA SỬA: dòng kế tiếp cùng diễn giải luôn bị coi là phí,
    # bất kể số tiền -> khoản chuyển tiền thứ 2 (300tr) bị hạch toán thành
    # "BVB-Phí chuyển tiền" (Nợ 3311401000) thay vì "Chuyển tiền QR Vietjet" (Nợ 3388403100).
    # Bỏ @expectedFailure khi đã sửa.
    @unittest.expectedFailure
    def test_two_large_transfers_with_same_narrative(self):
        result = self.run_bvb([
            bvb_row(self.TRANSFER, debit=500_000_000, refer="111"),
            bvb_row(self.TRANSFER, debit=300_000_000, refer="112"),
            bvb_row("Other", credit=1, refer="999"),
        ])

        transfers = sorted(
            (r[2]["Số tham chiếu"], r[2]["total_amount"])
            for r in result if r[1] == "TRANSFER"
        )

        self.assertEqual(transfers, [("111", 500_000_000), ("112", 300_000_000)])
        self.assertFalse([r for r in result if r[1] == "FEE"])

    def test_consecutive_same_narrative_debits_are_not_double_counted(self):
        # Code cũ: 3 dòng 500tr/300tr/200tr -> TRANSFER 800tr + FEE 500tr = 1.3 tỷ
        # (dòng 300tr bị cộng 2 lần). Tổng phát sinh nợ phải đúng bằng 1 tỷ.
        result = self.run_bvb([
            bvb_row(self.TRANSFER, debit=500_000_000, refer="111"),
            bvb_row(self.TRANSFER, debit=300_000_000, refer="112"),
            bvb_row(self.TRANSFER, debit=200_000_000, refer="113"),
            bvb_row("Other", credit=1, refer="999"),
        ])

        total_debit = sum(
            r[2]["total_amount"] for r in result if r[1] in ("TRANSFER", "FEE")
        )
        self.assertEqual(total_debit, 1_000_000_000)

    def test_refund_uses_its_own_reference(self):
        result = self.run_bvb([
            bvb_row("GALAXY PAY YEU CAU HOAN TRA", debit=200, refer="555"),
            bvb_row("Other", credit=1, refer="999"),
        ])

        refunds = [r for r in result if r[1] == "REFUND"]
        self.assertEqual(refunds[0][2]["Số tham chiếu"], "555")


WALLET_INPUT_FILE = "RECON_REPORT_WALLET_INPUT_20261005 (5)_Chi_tiết_giao_dịch.csv"
WALLET_OUTPUT_FILE = "RECON_REPORT_WALLET_OUTPUT_20261005 (5)_Chi_tiết_giao_dịch.csv"


def wallet_input_row(status, amount=100_000, trans_type="CASHIN"):
    return {
        "Provider": "NAPAS",
        "TransType": trans_type,
        "Reconciliation Status": status,
        "Trans Amount": str(amount),
        "Transaction Date": "05/10/2026 10:00:00",
        "Currency": "VND",
    }


def wallet_output_row(status, amount=100_000, trans_type="PAYMENT"):
    return {
        "Merchant Name": "VIETJET AIR",
        "TransType": trans_type,
        "Reconciliation Status": status,
        "Trans Amount": str(amount),
        "Transaction Date": "05/10/2026 10:00:00",
        "Currency": "VND",
    }


class NewWalletReconResultTest(unittest.TestCase):

    def run_wallet(self, file_name, rows):
        with mock.patch("builtins.print"):
            summary = process_file_data(file_name, rows)

        return sorted(
            (payment_type, recon_result, info["total_amount"])
            for merchant_ids in summary.values()
            for payment_types in merchant_ids.values()
            for payment_type, dates in payment_types.items()
            for results in dates.values()
            for recon_result, infos in results.items()
            for info in infos
        )

    def test_input_keeps_recon_result(self):
        self.assertEqual(
            self.run_wallet(WALLET_INPUT_FILE, [
                wallet_input_row("MATCH_SUCCESS", 100_000),
                wallet_input_row("MISS_GPAY", 20_000),
                wallet_input_row("MATCH_SUCCESS", 50_000),
            ]),
            [("CASHIN", "MATCH_SUCCESS", 150_000), ("CASHIN", "MISS_GPAY", 20_000)],
        )

    def test_input_vietnamese_status_is_normalized(self):
        self.assertEqual(
            self.run_wallet(WALLET_INPUT_FILE, [
                wallet_input_row("Xử lý lệch (Galaxy Pay nhiều hơn)", 30_000),
            ]),
            [("CASHIN", "GPAY_MORE", 30_000)],
        )

    def test_output_miss_partner_is_not_counted_as_match(self):
        self.assertEqual(
            self.run_wallet(WALLET_OUTPUT_FILE, [
                wallet_output_row("MATCH_SUCCESS", 100_000),
                wallet_output_row("MISS_PARTNER", 40_000),
            ]),
            [("PAYMENT", "MATCH_SUCCESS", 100_000), ("PAYMENT", "MISS_PARTNER", 40_000)],
        )

    def test_unknown_status_raises(self):
        with self.assertRaises(ValueError):
            self.run_wallet(WALLET_INPUT_FILE, [wallet_input_row("SOMETHING_NEW")])

    def test_empty_transaction_date_falls_back_to_gp_trans_date(self):
        row = wallet_input_row("MISS_PARTNER", 9_900_000, trans_type="CASHOUT")
        row["Transaction Date"] = ""
        row["GpTransDate"] = "01/10/2026 13:44:05"

        with mock.patch("builtins.print"):
            summary = process_file_data(WALLET_INPUT_FILE, [row])

        self.assertEqual(
            list(summary["NAPAS"]["NAPAS"]["CASHOUT"].keys()),
            ["01/10/2026"],
        )

    def test_transaction_date_has_priority_over_gp_trans_date(self):
        row = wallet_input_row("MATCH_SUCCESS")
        row["GpTransDate"] = "04/10/2026 23:59:00"

        with mock.patch("builtins.print"):
            summary = process_file_data(WALLET_INPUT_FILE, [row])

        self.assertEqual(
            list(summary["NAPAS"]["NAPAS"]["CASHIN"].keys()),
            ["05/10/2026"],
        )


class ParseCsvFileTest(unittest.TestCase):

    def test_unknown_file_type_raises(self):
        with tempfile.TemporaryDirectory() as tmp:
            csv_file = Path(tmp) / "UNKNOWN_bank.csv"
            csv_file.write_text("a,b\n1,2\n", encoding="utf-8")

            with self.assertRaises(ValueError):
                convert_data_to_json.parse_csv_file(csv_file)

    def test_scan_keeps_invalid_csv_and_raises(self):
        with tempfile.TemporaryDirectory() as tmp:
            csv_file = Path(tmp) / "UNKNOWN_bank.csv"
            csv_file.write_text("a,b\n1,2\n", encoding="utf-8")

            with mock.patch.object(convert_data_to_json, "CSV_FOLDER", tmp), \
                    mock.patch("builtins.print"):
                with self.assertRaises(RuntimeError):
                    convert_data_to_json.scan_csv_folder()

            self.assertTrue(csv_file.exists())


SUMMARY_HEADER = [
    "File", "Ngày giao dịch", "Đối tác", "Merchant ID", "Loại giao dịch",
    "Kết quả đối soát", "Tổng số tiền", "Số lượng giao dịch", "Loại tiền tệ",
    "Ngày giao dịch gốc", "Ghi chú", "Số tham chiếu",
]


def summary_row(file_name, posting_date, amount="100,000"):
    return [
        file_name, posting_date, "MOMO", "VJ APP TICKET", "PAYMENT",
        "MATCH_SUCCESS", amount, "1", "VND", "", "", "",
    ]


class ExportMainTest(unittest.TestCase):

    def run_export(self, rows):
        with tempfile.TemporaryDirectory() as tmp:
            input_file = Path(tmp) / "recon_summary.csv"
            output_file = Path(tmp) / "upload_file.csv"

            with open(input_file, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(SUMMARY_HEADER)
                writer.writerows(rows)

            with mock.patch.object(export_upload_file_1, "INPUT_FILE", str(input_file)), \
                    mock.patch.object(export_upload_file_1, "OUTPUT_FILE", str(output_file)), \
                    mock.patch("builtins.print"):
                export_upload_file_1.main()

            [upload] = Path(tmp).glob("upload_file_*.csv")
            with open(upload, encoding="utf-8-sig") as f:
                return list(csv.DictReader(f))

    def test_sorted_by_real_date_and_stt_sequential(self):
        rows = self.run_export([
            summary_row("RECON_REPORT_INPUT_20260901.xlsx", "01/09/2026"),
            summary_row("RECON_REPORT_INPUT_20260831.xlsx", "31/08/2026"),
        ])

        self.assertEqual(
            [r["Ngày hạch toán"] for r in rows],
            ["31/08/2026", "31/08/2026", "01/09/2026", "01/09/2026"],
        )
        self.assertEqual([r["STT"] for r in rows], ["1", "1", "2", "2"])
        self.assertEqual([r["Nợ/Có"] for r in rows], ["Nợ", "Có", "Nợ", "Có"])

    def test_unknown_file_type_raises(self):
        with self.assertRaises(ValueError):
            self.run_export([summary_row("SOMETHING_ELSE.xlsx", "01/09/2026")])

    def test_invalid_amount_raises(self):
        with self.assertRaises(ValueError):
            self.run_export([
                summary_row("RECON_REPORT_INPUT_20260901.xlsx", "01/09/2026", amount="abc")
            ])

    def test_empty_input_raises(self):
        with self.assertRaises(ValueError):
            self.run_export([])


class BuildDescriptionTest(unittest.TestCase):

    def describe(self, title):
        return export_upload_file_1.build_description(
            title, "05/10/2026", "05/10/2026", "", "BIDV",
            "RECON_REPORT_WALLET_INPUT_20261006.xlsx", "", "CASHOUT",
        )

    def test_withdraw(self):
        self.assertEqual(
            self.describe("BIDV-Rút tiền TK Ví GPay"),
            "BIDV-Rút tiền từ TK Ví GPay N05/10/2026",
        )

    def test_deposit(self):
        self.assertEqual(
            self.describe("HDB-Nạp tiền TK Ví"),
            "HDB-Nạp tiền TK Ví N05/10/2026",
        )


class RunReconciliationTest(unittest.TestCase):

    def test_stops_after_failed_step(self):
        calls = []

        def ok():
            calls.append("ok")

        def fail():
            raise RuntimeError("boom")

        with mock.patch.object(run_reconciliation, "convert_main", fail), \
                mock.patch.object(run_reconciliation, "get_data_main", ok), \
                mock.patch.object(run_reconciliation, "export_main", ok), \
                mock.patch("builtins.input"), \
                mock.patch("builtins.print"), \
                mock.patch("traceback.print_exc"):
            with self.assertRaises(SystemExit) as ctx:
                run_reconciliation.main()

        self.assertEqual(ctx.exception.code, 1)
        self.assertEqual(calls, [])


if __name__ == "__main__":
    unittest.main()
