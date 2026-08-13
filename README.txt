# Reconciliation Tool

Tool dùng để chuyển đổi các file Daily Reconciliation thành file upload kế toán.

---

# Cấu trúc thư mục

```
ReconciliationTool/
│
├── csv_recon_report/
│   ├── recon_export.json
│   ├── recon_summary.json
│   ├── recon_summary.csv
│   └── upload_file.csv
│
├── DAILY_RECON/
│   └── (Đặt các file Excel cần xử lý vào đây)
│
├── logs/
│
├── Scripts/
│   ├── convert_data_to_json.py
│   ├── get_data_4.py
│   ├── export_upload_file_1.py
│   └── run_reconciliation.py
│
├── requirements.txt
│
├── run.command        (Mac)
├── run.bat            (Windows)
│
├── Reconciliation      (Mac executable)
└── Reconciliation.exe  (Windows executable)
```

---

# Yêu cầu

## Người dùng

Không cần cài Python.

Chỉ cần:

- Reconciliation (Mac) hoặc Reconciliation.exe (Windows)
- run.command (Mac) hoặc run.bat (Windows)

---

## Developer

Yêu cầu:

- Python 3.10+
- pip

Cài package:

```bash
python -m pip install -r requirements.txt
```

---

# Cách sử dụng

## Bước 1

Copy toàn bộ file Daily Recon vào

```
DAILY_RECON/
```

---

## Bước 2

Mac:

Double click

```
run.command
```

Windows:

Double click

```
run.bat
```

---

## Bước 3

Đợi chương trình chạy hoàn tất.

Terminal sẽ hiển thị:

```
STEP 1 - Convert Excel -> CSV/JSON
STEP 2 - Generate recon_summary.csv
STEP 3 - Generate upload_file.csv

All steps completed successfully.
```

---

## Bước 4

File kết quả:

```
csv_recon_report/upload_file.csv
```

Đây là file dùng để upload lên hệ thống.

---

# Build executable

## Mac

Cài package

```bash
python3 -m pip install -r requirements.txt
```

Build

```bash
python3 -m PyInstaller \
--onefile \
--name Reconciliation \
Scripts/run_reconciliation.py
```

Executable sẽ nằm tại

```
dist/Reconciliation
```

Copy file này ra thư mục gốc của project.

---

## Windows

Cài package

```cmd
python -m pip install -r requirements.txt
```

Build

```cmd
pyinstaller --onefile --name Reconciliation Scripts\run_reconciliation.py
```

Executable sẽ nằm tại

```
dist\Reconciliation.exe
```

Copy file này ra thư mục gốc của project.

---

# Các bước xử lý

Tool thực hiện 3 bước:

### STEP 1

```
convert_data_to_json.py
```

- Đọc file Excel trong DAILY_RECON
- Xuất CSV
- Xuất recon_export.json

---

### STEP 2

```
get_data_4.py
```

- Đọc recon_export.json
- Tổng hợp dữ liệu
- Xuất recon_summary.json
- Xuất recon_summary.csv

---

### STEP 3

```
export_upload_file_1.py
```

- Đọc recon_summary.csv
- Sinh upload_file.csv

---

# Lưu ý

- Không đổi tên các thư mục:
  - DAILY_RECON
  - csv_recon_report
  - logs

- Chỉ đặt file Excel đầu vào trong thư mục DAILY_RECON.

- File upload cuối cùng luôn nằm tại:

```
csv_recon_report/upload_file.csv
```

---

# Version

Current Version: 1.0