# -*- coding: utf-8 -*-
"""
Hệ thống Quản lý & Tra cứu Hoàn tiền Hoa hồng Shopee
Giao diện UI/UX Hiện đại - Tone màu Cam Shopee (#ee4d2d)
Hỗ trợ lưu trữ trực tiếp vào Google Sheets theo thời gian thực (Real-time)
Tích hợp Bảng Mapping_Data, Import CSV Shopee & Đồng bộ Tên Khách Hàng
Cơ chế tự động dự phòng (Fallback) file CSV cục bộ
"""

import streamlit as st
import pandas as pd
import datetime
import os
import json
import io

# Thử import gspread để kết nối Google Sheets
try:
    import gspread
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False

# =============================================================================
# 1. CẤU HÌNH HỆ THỐNG
# =============================================================================
CSV_FILE = "shopee_cashback.csv"
MAPPING_CSV_FILE = "mapping_data.csv"
CREDENTIALS_FILE = "credentials.json"
DEFAULT_SHEET_NAME = "Shopee_Cashback"
MAPPING_SHEET_NAME = "Mapping_Data"
DEFAULT_ADMIN_PASS = "Kunhan1996@"
STATUS_OPTIONS = ["Chờ Shopee duyệt", "Chờ Bank", "Đã bank", "Không hợp lệ"]
REQUIRED_COLS = ["Tên khách", "Ngày đặt", "Mã đơn", "Tổng HH thực tế", "Trạng thái"]
MAPPING_COLS = ["Mã Zalo (Sub_id2)", "Tên Khách Hàng"]


# =============================================================================
# 2. XỬ LÝ KẾT NỐI GOOGLE SHEETS & SECRETS
# =============================================================================
def get_credentials_dict():
    """
    Lấy thông tin xác thực Service Account:
    1. Ưu tiên lấy từ st.secrets["google_json"] khi chạy trên Streamlit Cloud / Hosting.
    2. Dự phòng đọc từ file credentials.json cục bộ nếu có.
    """
    try:
        if "google_json" in st.secrets:
            secret_data = st.secrets["google_json"]
            if isinstance(secret_data, str):
                return json.loads(secret_data)
            elif isinstance(secret_data, dict):
                return secret_data
    except Exception:
        pass

    if os.path.exists(CREDENTIALS_FILE):
        try:
            with open(CREDENTIALS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None

    return None

def get_service_account_email():
    """Đọc email của Service Account từ st.secrets hoặc file credentials.json."""
    creds = get_credentials_dict()
    if creds and isinstance(creds, dict):
        return creds.get("client_email", "")
    return ""

def get_gsheet_spreadsheet(sheet_identifier=DEFAULT_SHEET_NAME):
    """Mở Spreadsheet Google Sheets."""
    if not GSPREAD_AVAILABLE:
        return None, "Thư viện `gspread` chưa được cài đặt."
    
    creds_dict = get_credentials_dict()
    if not creds_dict:
        return None, "Chưa tìm thấy cấu hình xác thực (st.secrets['google_json'] hoặc file credentials.json)."
    
    try:
        gc = gspread.service_account_from_dict(creds_dict)
        if str(sheet_identifier).startswith("https://"):
            spreadsheet = gc.open_by_url(sheet_identifier)
        else:
            try:
                spreadsheet = gc.open(sheet_identifier)
            except Exception:
                spreadsheet = gc.open_by_key(sheet_identifier)
        return spreadsheet, None
    except Exception as e:
        return None, str(e)

def get_gsheet_worksheet(sheet_identifier=DEFAULT_SHEET_NAME, worksheet_name=None):
    """
    Lấy Worksheet từ Google Sheets.
    Nếu worksheet_name=None -> Lấy sheet1 (Shopee_Cashback).
    Nếu worksheet_name được chỉ định (như Mapping_Data) -> Lấy sheet đó, nếu chưa có thì tự động tạo mới.
    """
    sh, err = get_gsheet_spreadsheet(sheet_identifier)
    if sh is None:
        return None, err
    
    try:
        if not worksheet_name:
            return sh.sheet1, None
        
        # Tìm worksheet theo tên
        titles = [ws.title for ws in sh.worksheets()]
        if worksheet_name in titles:
            return sh.worksheet(worksheet_name), None
        else:
            # Tự động tạo worksheet mới nếu chưa tồn tại
            cols_init = MAPPING_COLS if worksheet_name == MAPPING_SHEET_NAME else REQUIRED_COLS
            new_ws = sh.add_worksheet(title=worksheet_name, rows=1000, cols=10)
            new_ws.append_row(cols_init)
            return new_ws, None
    except Exception as e:
        return None, str(e)


# =============================================================================
# 3. QUẢN LÝ DỮ LIỆU ĐƠN HÀNG (SHOPEE_CASHBACK)
# =============================================================================
def normalize_df(df):
    """Chuẩn hóa cấu trúc DataFrame đơn hàng đồng nhất."""
    df_clean = df.copy()
    for col in REQUIRED_COLS:
        if col not in df_clean.columns:
            df_clean[col] = "" if col != "Tổng HH thực tế" else 0
            
    df_clean["Mã đơn"] = df_clean["Mã đơn"].fillna("").astype(str).str.strip()
    df_clean["Tên khách"] = df_clean["Tên khách"].fillna("").astype(str).str.strip()
    df_clean["Ngày đặt"] = df_clean["Ngày đặt"].fillna("").astype(str).str.strip()
    df_clean["Tổng HH thực tế"] = pd.to_numeric(df_clean["Tổng HH thực tế"], errors="coerce").fillna(0).astype(int)
    df_clean["Trạng thái"] = df_clean["Trạng thái"].fillna("Chờ Shopee duyệt").astype(str).str.strip()
    return df_clean[REQUIRED_COLS]

def load_local_csv():
    """Đọc dữ liệu đơn hàng từ file CSV cục bộ."""
    if not os.path.exists(CSV_FILE):
        df = pd.DataFrame(columns=REQUIRED_COLS)
        save_local_csv(df)
        return df
    try:
        df = pd.read_csv(
            CSV_FILE,
            dtype={"Mã đơn": str, "Tên khách": str, "Ngày đặt": str, "Trạng thái": str},
            encoding="utf-8-sig"
        )
        return normalize_df(df)
    except Exception:
        return pd.DataFrame(columns=REQUIRED_COLS)

def save_local_csv(df):
    """Lưu backup đơn hàng ra file CSV cục bộ."""
    try:
        df_save = df.copy()
        df_save["Mã đơn"] = df_save["Mã đơn"].astype(str)
        df_save.to_csv(CSV_FILE, index=False, encoding="utf-8-sig")
    except Exception:
        pass

def load_data(sheet_target=DEFAULT_SHEET_NAME):
    """
    Đọc dữ liệu đơn hàng: Ưu tiên từ Google Sheets.
    Fallback về file CSV cục bộ nếu chưa cấu hình hoặc mất mạng.
    """
    ws, err = get_gsheet_worksheet(sheet_target)
    if ws is not None:
        try:
            records = ws.get_all_values()
            if len(records) > 0:
                headers = records[0]
                rows = records[1:] if len(records) > 1 else []
                df = pd.DataFrame(rows, columns=headers)
                df = normalize_df(df)
            else:
                ws.append_row(REQUIRED_COLS)
                df = pd.DataFrame(columns=REQUIRED_COLS)
                
            save_local_csv(df)
            return df, "gsheets", None
        except Exception as e:
            fallback_df = load_local_csv()
            return fallback_df, "csv_fallback", f"Lỗi đọc Google Sheets: {e}"
            
    fallback_df = load_local_csv()
    return fallback_df, "csv", err

def add_order_to_storage(new_row_dict, sheet_target=DEFAULT_SHEET_NAME):
    """Thêm 1 đơn hàng mới vào Google Sheets & CSV."""
    cust_name = new_row_dict.get("Tên khách", "")
    order_date = new_row_dict.get("Ngày đặt", "")
    order_code = str(new_row_dict.get("Mã đơn", "")).strip()
    commission = int(new_row_dict.get("Tổng HH thực tế", 0))
    status = new_row_dict.get("Trạng thái", "Chờ Shopee duyệt")
    
    ws, err = get_gsheet_worksheet(sheet_target)
    saved_to_gsheet = False
    
    if ws is not None:
        try:
            code_formatted = f"'{order_code}" if order_code.startswith("0") else order_code
            row_data = [cust_name, order_date, code_formatted, commission, status]
            ws.append_row(row_data, value_input_option="USER_ENTERED")
            saved_to_gsheet = True
        except Exception as e:
            st.error(f"Lỗi ghi vào Google Sheets: {e}")
            
    current_df = load_local_csv()
    new_df = pd.DataFrame([{
        "Tên khách": cust_name,
        "Ngày đặt": order_date,
        "Mã đơn": order_code,
        "Tổng HH thực tế": commission,
        "Trạng thái": status
    }])
    updated_df = pd.concat([current_df, new_df], ignore_index=True)
    save_local_csv(updated_df)
    
    return saved_to_gsheet

def batch_add_orders_to_storage(orders_list, sheet_target=DEFAULT_SHEET_NAME):
    """
    Thêm danh sách nhiều đơn hàng cùng lúc (Dùng khi Import CSV).
    Sử dụng append_rows để tối ưu tốc độ.
    """
    if not orders_list:
        return True, 0
    
    ws, err = get_gsheet_worksheet(sheet_target)
    rows_to_append = []
    for item in orders_list:
        code_str = str(item.get("Mã đơn", "")).strip()
        code_fmt = f"'{code_str}" if code_str.startswith("0") else code_str
        rows_to_append.append([
            str(item.get("Tên khách", "")).strip(),
            str(item.get("Ngày đặt", "")).strip(),
            code_fmt,
            int(item.get("Tổng HH thực tế", 0)),
            str(item.get("Trạng thái", "Chờ Shopee duyệt")).strip()
        ])
    
    # 1. Ghi lên Google Sheets
    if ws is not None:
        try:
            ws.append_rows(rows_to_append, value_input_option="USER_ENTERED")
        except Exception as e:
            st.warning(f"Lưu ý: Không thể ghi trực tiếp lên Google Sheets ({e}). Đã lưu dự phòng vào CSV.")
            
    # 2. Đồng bộ vào CSV cục bộ
    current_df = load_local_csv()
    new_records_df = pd.DataFrame([{
        "Tên khách": str(item.get("Tên khách", "")).strip(),
        "Ngày đặt": str(item.get("Ngày đặt", "")).strip(),
        "Mã đơn": str(item.get("Mã đơn", "")).strip(),
        "Tổng HH thực tế": int(item.get("Tổng HH thực tế", 0)),
        "Trạng thái": str(item.get("Trạng thái", "Chờ Shopee duyệt")).strip()
    } for item in orders_list])
    
    updated_df = pd.concat([current_df, new_records_df], ignore_index=True)
    save_local_csv(updated_df)
    return True, len(orders_list)

def save_all_to_storage(df, sheet_target=DEFAULT_SHEET_NAME):
    """Ghi đè toàn bộ dữ liệu đơn hàng (dùng cho Tab 2 Data Editor)."""
    df_clean = normalize_df(df)
    save_local_csv(df_clean)
    
    ws, err = get_gsheet_worksheet(sheet_target)
    if ws is not None:
        try:
            rows = [REQUIRED_COLS]
            for _, r in df_clean.iterrows():
                code_val = str(r["Mã đơn"]).strip()
                code_fmt = f"'{code_val}" if code_val.startswith("0") else code_val
                rows.append([
                    str(r["Tên khách"]),
                    str(r["Ngày đặt"]),
                    code_fmt,
                    int(r["Tổng HH thực tế"]),
                    str(r["Trạng thái"])
                ])
            ws.clear()
            ws.update(rows, value_input_option="USER_ENTERED")
            return True, None
        except Exception as e:
            return False, str(e)
    return False, err


# =============================================================================
# 4. QUẢN LÝ BẢNG MAPPING (MAPPING_DATA) & ĐỒNG BỘ TÊN KHÁCH HÀNG
# =============================================================================
def normalize_mapping_df(df):
    """Chuẩn hóa cấu trúc DataFrame Mapping Khách Hàng."""
    df_clean = df.copy()
    for col in MAPPING_COLS:
        if col not in df_clean.columns:
            df_clean[col] = ""
    df_clean["Mã Zalo (Sub_id2)"] = df_clean["Mã Zalo (Sub_id2)"].fillna("").astype(str).str.strip()
    df_clean["Tên Khách Hàng"] = df_clean["Tên Khách Hàng"].fillna("").astype(str).str.strip()
    # Loại bỏ dòng hoàn toàn rỗng
    df_clean = df_clean[df_clean["Mã Zalo (Sub_id2)"] != ""]
    return df_clean[MAPPING_COLS].drop_duplicates(subset=["Mã Zalo (Sub_id2)"], keep="last")

def load_local_mapping_csv():
    """Đọc dữ liệu mapping từ file CSV cục bộ."""
    if not os.path.exists(MAPPING_CSV_FILE):
        df = pd.DataFrame(columns=MAPPING_COLS)
        save_local_mapping_csv(df)
        return df
    try:
        df = pd.read_csv(MAPPING_CSV_FILE, dtype={"Mã Zalo (Sub_id2)": str, "Tên Khách Hàng": str}, encoding="utf-8-sig")
        return normalize_mapping_df(df)
    except Exception:
        return pd.DataFrame(columns=MAPPING_COLS)

def save_local_mapping_csv(df_mapping):
    """Lưu backup mapping ra file CSV cục bộ."""
    try:
        df_save = df_mapping.copy()
        df_save["Mã Zalo (Sub_id2)"] = df_save["Mã Zalo (Sub_id2)"].astype(str)
        df_save.to_csv(MAPPING_CSV_FILE, index=False, encoding="utf-8-sig")
    except Exception:
        pass

def load_mapping_data(sheet_target=DEFAULT_SHEET_NAME):
    """
    Đọc dữ liệu bảng Mapping: Ưu tiên đọc từ Google Sheets (Sheet: Mapping_Data).
    Fallback về file mapping_data.csv nếu offline.
    """
    ws, err = get_gsheet_worksheet(sheet_target, worksheet_name=MAPPING_SHEET_NAME)
    if ws is not None:
        try:
            records = ws.get_all_values()
            if len(records) > 0:
                headers = records[0]
                rows = records[1:] if len(records) > 1 else []
                df = pd.DataFrame(rows, columns=headers)
                df = normalize_mapping_df(df)
            else:
                ws.append_row(MAPPING_COLS)
                df = pd.DataFrame(columns=MAPPING_COLS)
                
            save_local_mapping_csv(df)
            return df, "gsheets", None
        except Exception as e:
            fallback_df = load_local_mapping_csv()
            return fallback_df, "csv_fallback", f"Lỗi đọc Mapping_Data: {e}"
            
    fallback_df = load_local_mapping_csv()
    return fallback_df, "csv", err

def save_mapping_data(df_mapping, sheet_target=DEFAULT_SHEET_NAME):
    """Lưu toàn bộ bảng Mapping_Data lên Google Sheets & CSV."""
    df_clean = normalize_mapping_df(df_mapping)
    save_local_mapping_csv(df_clean)
    
    ws, err = get_gsheet_worksheet(sheet_target, worksheet_name=MAPPING_SHEET_NAME)
    if ws is not None:
        try:
            rows = [MAPPING_COLS]
            for _, r in df_clean.iterrows():
                zalo_val = str(r["Mã Zalo (Sub_id2)"]).strip()
                zalo_fmt = f"'{zalo_val}" if zalo_val.startswith("0") else zalo_val
                rows.append([zalo_fmt, str(r["Tên Khách Hàng"]).strip()])
            ws.clear()
            ws.update(rows, value_input_option="USER_ENTERED")
            return True, None
        except Exception as e:
            return False, str(e)
    return False, err

def batch_add_mapping_records(new_mappings_dict, sheet_target=DEFAULT_SHEET_NAME):
    """
    Thêm danh sách khách hàng mới vào bảng Mapping_Data.
    new_mappings_dict: {zalo_id: customer_name}
    """
    if not new_mappings_dict:
        return
    
    current_mapping, _, _ = load_mapping_data(sheet_target)
    new_rows = []
    existing_zalos = set(current_mapping["Mã Zalo (Sub_id2)"].astype(str).str.strip())
    
    for zalo, name in new_mappings_dict.items():
        zalo_str = str(zalo).strip()
        if zalo_str and (zalo_str not in existing_zalos):
            new_rows.append({"Mã Zalo (Sub_id2)": zalo_str, "Tên Khách Hàng": str(name).strip()})
            existing_zalos.add(zalo_str)
            
    if new_rows:
        updated_mapping = pd.concat([current_mapping, pd.DataFrame(new_rows)], ignore_index=True)
        save_mapping_data(updated_mapping, sheet_target)

def sync_customer_name(old_name, new_name, target_sub2=None, sheet_target=DEFAULT_SHEET_NAME):
    """
    TÍNH NĂNG 3: ĐỒNG BỘ TÊN KHÁCH HÀNG TRỰC TIẾP TRÊN WEB.
    1. Cập nhật tên mới vào bảng Mapping_Data (theo target_sub2 hoặc old_name).
    2. Tự động quét toàn bộ bảng Shopee_Cashback, tìm mọi dòng mang tên old_name -> đổi thành new_name.
    3. Lưu đồng bộ cả 2 bảng lên Google Sheets & CSV.
    """
    old_name_clean = str(old_name).strip()
    new_name_clean = str(new_name).strip()
    
    if not old_name_clean or not new_name_clean:
        return False, 0, "Tên không được để trống!"
        
    # 1. Cập nhật trong bảng Mapping_Data
    df_mapping, _, _ = load_mapping_data(sheet_target)
    if target_sub2:
        mask_map = df_mapping["Mã Zalo (Sub_id2)"] == str(target_sub2).strip()
        df_mapping.loc[mask_map, "Tên Khách Hàng"] = new_name_clean
    else:
        mask_map = df_mapping["Tên Khách Hàng"] == old_name_clean
        df_mapping.loc[mask_map, "Tên Khách Hàng"] = new_name_clean
    save_mapping_data(df_mapping, sheet_target)
    
    # 2. Cập nhật hàng loạt trong bảng Shopee_Cashback
    df_orders, _, _ = load_data(sheet_target)
    mask_orders = df_orders["Tên khách"] == old_name_clean
    matched_count = int(mask_orders.sum())
    
    if matched_count > 0:
        df_orders.loc[mask_orders, "Tên khách"] = new_name_clean
        save_all_to_storage(df_orders, sheet_target)
        
    return True, matched_count, None

def auto_sync_names_from_mapping(sheet_target=DEFAULT_SHEET_NAME):
    """
    Tự động quét và chuẩn hóa toàn bộ tên khách hàng trong Shopee_Cashback
    dựa theo bảng Mapping_Data (chuyển các tên không dấu / viết liền sang tên chuẩn có dấu).
    """
    import unicodedata
    def clean_key(s):
        s = unicodedata.normalize('NFD', str(s))
        s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
        return s.replace('đ', 'd').replace('Đ', 'D').lower().replace(' ', '').strip()
        
    df_mapping, _, _ = load_mapping_data(sheet_target)
    df_orders, _, _ = load_data(sheet_target)
    if df_mapping.empty or df_orders.empty:
        return 0, []
        
    clean_map = {}
    for _, r in df_mapping.iterrows():
        std_name = str(r["Tên Khách Hàng"]).strip()
        if std_name:
            clean_map[clean_key(std_name)] = std_name
            
    changes = []
    updated_orders = 0
    for idx, row in df_orders.iterrows():
        cur_name = str(row["Tên khách"]).strip()
        ck = clean_key(cur_name)
        if ck in clean_map:
            std_name = clean_map[ck]
            if cur_name != std_name:
                df_orders.at[idx, "Tên khách"] = std_name
                updated_orders += 1
                changes.append(f"{cur_name} ➔ {std_name}")
                
    if updated_orders > 0:
        save_all_to_storage(df_orders, sheet_target=sheet_target)
        
    return updated_orders, list(set(changes))


# =============================================================================
# 5. BỘ LỌC & XỬ LÝ FILE CSV SHOPEE AFFILIATE (TÍNH NĂNG 2)
# =============================================================================
def detect_shopee_csv_columns(df):
    """Tự động nhận diện linh hoạt các cột trong file CSV Shopee Affiliate."""
    detected = {
        "order_id": None,
        "order_time": None,
        "complete_time": None,
        "status": None,
        "commission": None,
        "sub1": None,
        "sub2": None
    }
    
    col_mapping_rules = {
        "order_id": ["id đơn hàng", "mã đơn hàng", "mã đơn", "order id", "order sn", "mã đơn shopee", "order_id", "order_sn", "mã đặt hàng"],
        "order_time": ["thời gian đặt hàng", "thời gian tạo đơn", "order time", "purchase time", "thời gian đặt"],
        "complete_time": ["thời gian hoàn thành đơn hàng", "thời gian hoàn thành đơn", "thời gian hoàn thành", "complete time", "completed time"],
        "status": ["trạng thái đặt hàng", "trạng thái đơn hàng", "trạng thái đơn", "order status", "trạng thái"],
        "commission": [
            "tổng hoa hồng sản phẩm", "hoa hồng ròng tiếp thị liên kết", "tổng hoa hồng đơn hàng",
            "hoa hồng thực tế", "tổng hh thực tế", "hoa hồng thuần", "tổng hoa hồng",
            "net commission", "commission amount", "actual commission"
        ],
        "sub1": ["sub_id1", "sub id 1", "sub id1", "sub1", "sub_1", "tên tạm", "subid1"],
        "sub2": ["sub_id2", "sub id 2", "sub id2", "sub2", "sub_2", "mã zalo", "zalo", "subid2"]
    }
    
    for actual_col in df.columns:
        norm_col = str(actual_col).lower().strip()
        for field, keywords in col_mapping_rules.items():
            if detected[field] is None:
                # Bỏ qua các cột chứa từ khóa gây nhiễu cho commission
                if field == "commission" and any(bad in norm_col for bad in ["loại", "tỷ lệ", "phí", "đối tác"]):
                    continue
                for kw in keywords:
                    if kw in norm_col:
                        detected[field] = actual_col
                        break
                        
    return detected

def parse_flexible_date(date_val):
    """Parse chuỗi ngày tháng sang datetime.date bất kể định dạng."""
    if pd.isna(date_val) or date_val is None or str(date_val).strip() == "":
        return None
    val_str = str(date_val).strip()
    
    # Thử parse ngày tháng phổ biến
    for fmt in [
        "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%d-%m-%Y %H:%M:%S",
        "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y",
        "%Y/%m/%d %H:%M:%S", "%Y/%m/%d",
        "%m/%d/%Y %H:%M:%S", "%m/%d/%Y"
    ]:
        try:
            return datetime.datetime.strptime(val_str, fmt).date()
        except Exception:
            pass
            
    # Thử pandas to_datetime
    try:
        dt = pd.to_datetime(val_str, errors="coerce")
        if pd.notna(dt):
            return dt.date()
    except Exception:
        pass
        
    return None

def parse_commission_number(val):
    """Chuyển chuỗi số tiền hoa hồng sang số nguyên VNĐ an toàn."""
    if pd.isna(val) or val is None:
        return 0
    val_str = str(val).strip()
    if not val_str:
        return 0
    clean = val_str.replace("₫", "").replace("VND", "").replace("VNĐ", "").replace(" ", "").strip()
    
    # Có cả chấm và phẩy (vd: 1,234.56 hoặc 1.234,56)
    if "," in clean and "." in clean:
        if clean.rfind(".") > clean.rfind(","):
            clean = clean.replace(",", "")
        else:
            clean = clean.replace(".", "").replace(",", ".")
    elif "." in clean:
        parts = clean.split(".")
        if len(parts) > 2:
            clean = clean.replace(".", "")
        elif len(parts) == 2:
            if len(parts[1]) == 3:
                clean = clean.replace(".", "")
    elif "," in clean:
        parts = clean.split(",")
        if len(parts) > 2:
            clean = clean.replace(",", "")
        elif len(parts) == 2:
            if len(parts[1]) == 3:
                clean = clean.replace(",", "")
            else:
                clean = clean.replace(",", ".")
                
    try:
        return int(round(float(clean)))
    except Exception:
        return 0


# =============================================================================
# 6. GIAO DIỆN CHÍNH (STREAMLIT APP)
# =============================================================================
def format_vnd(amount):
    """Định dạng tiền tệ VNĐ chuẩn có dấu chấm phân cách hàng nghìn."""
    return f"{amount:,.0f} ₫".replace(",", ".")

def main():
    st.set_page_config(
        page_title="Tra cứu Hoàn tiền Hoa hồng Shopee",
        page_icon="🛍️",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # Custom CSS chuẩn Shopee hiện đại
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
        
        html, body, [class*="css"] {
            font-family: 'Plus Jakarta Sans', sans-serif;
        }
        
        :root {
            --shopee-orange: #ee4d2d;
            --shopee-orange-hover: #d73211;
            --shopee-orange-light: #fff5f2;
            --shopee-orange-border: #ffc5b8;
            --text-dark: #1e293b;
            --text-muted: #64748b;
        }
        
        .shopee-header-box {
            text-align: center;
            padding: 32px 20px 28px;
            background: linear-gradient(135deg, #fff5f2 0%, #ffffff 100%);
            border-radius: 20px;
            border: 1px solid var(--shopee-orange-border);
            box-shadow: 0 10px 30px rgba(238, 77, 45, 0.07);
            margin-bottom: 28px;
        }
        
        .shopee-badge {
            display: inline-block;
            background: linear-gradient(135deg, #ee4d2d 0%, #ff6b45 100%);
            color: white;
            font-size: 12px;
            font-weight: 700;
            letter-spacing: 1px;
            padding: 6px 16px;
            border-radius: 50px;
            text-transform: uppercase;
            margin-bottom: 12px;
            box-shadow: 0 4px 12px rgba(238, 77, 45, 0.25);
        }
        
        .shopee-title {
            color: #ee4d2d;
            font-size: 32px;
            font-weight: 800;
            margin: 6px 0 10px;
            letter-spacing: -0.5px;
        }
        
        .shopee-desc {
            color: var(--text-muted);
            font-size: 15px;
            max-width: 600px;
            margin: 0 auto;
            line-height: 1.5;
        }
        
        .metrics-container {
            display: flex;
            gap: 16px;
            margin: 20px 0 24px;
            justify-content: center;
            flex-wrap: wrap;
        }
        
        .metric-card {
            flex: 1;
            min-width: 220px;
            background: #ffffff;
            border-radius: 16px;
            padding: 20px 22px;
            border: 1px solid #e2e8f0;
            box-shadow: 0 4px 15px rgba(0,0,0,0.04);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        
        .metric-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 24px rgba(0,0,0,0.08);
        }
        
        .metric-card.pending {
            border-top: 4px solid #f59e0b;
            background: linear-gradient(180deg, #fffbf0 0%, #ffffff 70%);
        }
        
        .metric-card.success {
            border-top: 4px solid #10b981;
            background: linear-gradient(180deg, #f0fdf4 0%, #ffffff 70%);
        }

        .metric-card.info {
            border-top: 4px solid #3b82f6;
            background: linear-gradient(180deg, #eff6ff 0%, #ffffff 70%);
        }

        .metric-card.warning {
            border-top: 4px solid #ee4d2d;
            background: linear-gradient(180deg, #fff5f2 0%, #ffffff 70%);
        }
        
        .metric-title {
            font-size: 13px;
            font-weight: 600;
            color: var(--text-muted);
            margin-bottom: 6px;
            display: flex;
            align-items: center;
            gap: 6px;
        }
        
        .metric-amount {
            font-size: 26px;
            font-weight: 800;
            line-height: 1.2;
        }
        
        .metric-card.pending .metric-amount { color: #d97706; }
        .metric-card.success .metric-amount { color: #059669; }
        .metric-card.info .metric-amount { color: #2563eb; }
        .metric-card.warning .metric-amount { color: #ee4d2d; }
        
        .metric-count {
            font-size: 12px;
            color: #94a3b8;
            margin-top: 6px;
            font-weight: 500;
        }

        .stButton > button {
            background: linear-gradient(135deg, #ee4d2d 0%, #ff5722 100%) !important;
            color: white !important;
            border: none !important;
            border-radius: 10px !important;
            font-weight: 700 !important;
            font-size: 15px !important;
            padding: 10px 24px !important;
            box-shadow: 0 4px 14px rgba(238, 77, 45, 0.25) !important;
            transition: all 0.2s ease !important;
        }
        
        .stButton > button:hover {
            transform: translateY(-1px) !important;
            box-shadow: 0 6px 20px rgba(238, 77, 45, 0.4) !important;
            background: linear-gradient(135deg, #d73211 0%, #e64a19 100%) !important;
        }
        
        div[data-testid="stTextInput"] input {
            border-radius: 12px !important;
            border: 2px solid #e2e8f0 !important;
            padding: 12px 16px !important;
            font-size: 15px !important;
            transition: all 0.2s ease !important;
        }
        
        div[data-testid="stTextInput"] input:focus {
            border-color: #ee4d2d !important;
            box-shadow: 0 0 0 3px rgba(238, 77, 45, 0.15) !important;
        }

        .guide-box {
            background: #ffffff;
            border: 1px dashed #cbd5e1;
            border-radius: 16px;
            padding: 40px 25px;
            text-align: center;
            margin-top: 15px;
        }
        
        .guide-step {
            display: inline-block;
            margin: 10px 12px;
            text-align: left;
            background: #f8fafc;
            border-radius: 12px;
            padding: 16px 20px;
            border: 1px solid #e2e8f0;
            max-width: 250px;
            vertical-align: top;
        }
        
        .guide-step-num {
            background: #ee4d2d;
            color: white;
            font-weight: 700;
            font-size: 12px;
            border-radius: 50%;
            width: 24px;
            height: 24px;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            margin-right: 8px;
        }

        .sidebar-admin-badge {
            padding: 10px 14px;
            border-radius: 10px;
            font-weight: 700;
            font-size: 13px;
            text-align: center;
            margin-bottom: 16px;
        }
        
        .sidebar-admin-active {
            background: #ecfdf5;
            color: #065f46;
            border: 1px solid #a7f3d0;
        }
        
        .sidebar-admin-locked {
            background: #fff7ed;
            color: #9a3412;
            border: 1px solid #fed7aa;
        }

        .storage-status-badge {
            font-size: 12px;
            padding: 4px 10px;
            border-radius: 20px;
            font-weight: 600;
            display: inline-block;
            margin-top: 6px;
        }
        .storage-gsheets {
            background: #e8f5e9;
            color: #2e7d32;
            border: 1px solid #a5d6a7;
        }
        .storage-csv {
            background: #fff3e0;
            color: #e65100;
            border: 1px solid #ffcc80;
        }
        
        .cust-tag-new {
            background: #fff7ed;
            color: #c2410c;
            padding: 2px 8px;
            border-radius: 6px;
            font-size: 11px;
            font-weight: 700;
            border: 1px solid #fed7aa;
        }
        .cust-tag-old {
            background: #f0fdf4;
            color: #15803d;
            padding: 2px 8px;
            border-radius: 6px;
            font-size: 11px;
            font-weight: 700;
            border: 1px solid #bbf7d0;
        }
    </style>
    """, unsafe_allow_html=True)

    # Quản lý Session State
    if "admin_logged_in" not in st.session_state:
        st.session_state["admin_logged_in"] = False
    if "sheet_name" not in st.session_state:
        st.session_state["sheet_name"] = DEFAULT_SHEET_NAME

    # Cấu hình Sidebar
    with st.sidebar:
        st.markdown("### ⚙️ Quản Trị Hệ Thống")
        
        if not st.session_state["admin_logged_in"]:
            st.markdown("""
            <div class="sidebar-admin-badge sidebar-admin-locked">
                🔒 Chưa đăng nhập quyền Admin
            </div>
            """, unsafe_allow_html=True)
            
            admin_pass = st.text_input("Mật khẩu Admin:", type="password", placeholder="Nhập mật khẩu...", key="admin_password_input")
            if st.button("🔑 Đăng nhập Admin", use_container_width=True):
                if admin_pass == DEFAULT_ADMIN_PASS:
                    st.session_state["admin_logged_in"] = True
                    st.toast("Đăng nhập Admin thành công!", icon="🎉")
                    st.rerun()
                else:
                    st.error("Mật khẩu không đúng! Vui lòng thử lại.")
        else:
            st.markdown("""
            <div class="sidebar-admin-badge sidebar-admin-active">
                🟢 Đang đăng nhập quyền Admin
            </div>
            """, unsafe_allow_html=True)
            
            view_mode = st.radio(
                "Chọn chế độ xem:",
                ["👑 Bảng Quản trị (Admin)", "👁️ Xem Giao diện Khách hàng"],
                index=0,
                key="admin_view_mode"
            )
            
            # Cấu hình Google Sheets trong Sidebar Admin
            with st.expander("☁️ Cấu hình Google Sheets", expanded=False):
                cfg_sheet = st.text_input(
                    "Tên hoặc Link Google Sheet:",
                    value=st.session_state["sheet_name"],
                    help="Nhập tên bảng tính Google Sheets (mặc định 'Shopee_Cashback') hoặc link URL đầy đủ."
                )
                if cfg_sheet != st.session_state["sheet_name"]:
                    st.session_state["sheet_name"] = cfg_sheet
                    st.rerun()
                
                sa_email = get_service_account_email()
                if sa_email:
                    try:
                        has_secrets = "google_json" in st.secrets
                    except Exception:
                        has_secrets = False
                    if has_secrets:
                        st.success("✅ Đã kết nối qua `st.secrets['google_json']`")
                    else:
                        st.success("✅ Đã tìm thấy `credentials.json` cục bộ")
                    st.text_area("Email Service Account (chia sẻ quyền Editor cho email này):", value=sa_email, height=70)
                else:
                    st.warning("⚠️ Chưa cấu hình `st.secrets['google_json']` hoặc file `credentials.json`.")
                    st.caption("Ứng dụng đang tự động chạy ở chế độ **CSV cục bộ**.")

            if st.button("🚪 Đăng xuất", use_container_width=True):
                st.session_state["admin_logged_in"] = False
                st.toast("Đã đăng xuất khỏi quyền Admin!", icon="👋")
                st.rerun()
                
            st.divider()

        # Hiển thị trạng thái nguồn dữ liệu
        df_data, storage_mode, storage_err = load_data(st.session_state["sheet_name"])
        df_mapping, mapping_mode, mapping_err = load_mapping_data(st.session_state["sheet_name"])
        
        if storage_mode == "gsheets":
            st.markdown("""
            <div class="storage-status-badge storage-gsheets">
                🟢 Đang kết nối: Google Sheets
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div class="storage-status-badge storage-csv">
                📁 Đang lưu trữ: CSV cục bộ
            </div>
            """, unsafe_allow_html=True)
            if storage_err:
                st.caption(f"ℹ️ *{storage_err}*")

    # Điều hướng giao diện
    show_admin = st.session_state["admin_logged_in"] and (view_mode == "👑 Bảng Quản trị (Admin)")

    if not show_admin:
        # =====================================================================
        # A. GIAO DIỆN KHÁCH HÀNG (MẶC ĐỊNH) - GIỮ NGUYÊN 100% GIAO DIỆN
        # =====================================================================
        _, center_col, _ = st.columns([1, 8, 1])
        
        with center_col:
            st.markdown("""
            <div class="shopee-header-box">
                <span class="shopee-badge">SHOPEE AFFILIATE CASHBACK</span>
                <h1 class="shopee-title">🔍 Tra cứu Hoàn tiền Hoa hồng Shopee</h1>
                <p class="shopee-desc">
                    Nhập Tên của bạn hoặc Mã đơn hàng Shopee để kiểm tra tiến độ duyệt và hoàn tiền hoa hồng nhanh chóng.
                </p>
            </div>
            """, unsafe_allow_html=True)
            
            search_query = st.text_input(
                label="Nhập Tên hoặc Mã đơn hàng của bạn...",
                placeholder="🔍 Nhập Tên hoặc Mã đơn hàng của bạn... (Ví dụ: Nguyễn Văn An, 024891028421)",
                label_visibility="collapsed",
                key="client_search_box"
            ).strip()
            
            if search_query:
                mask = (
                    df_data["Tên khách"].str.contains(search_query, case=False, na=False) |
                    df_data["Mã đơn"].str.contains(search_query, case=False, na=False)
                )
                matched_df = df_data[mask].copy()
                if len(matched_df) > 0:
                    pending_df = matched_df[matched_df["Trạng thái"].isin(["Chờ Shopee duyệt", "Chờ Bank"])]
                    banked_df = matched_df[matched_df["Trạng thái"] == "Đã bank"]
                    
                    total_pending = pending_df["Tổng HH thực tế"].sum()
                    count_pending = len(pending_df)
                    
                    total_banked = banked_df["Tổng HH thực tế"].sum()
                    count_banked = len(banked_df)
                    
                    st.markdown(f"""
                    <div class="metrics-container">
                        <div class="metric-card pending">
                            <div class="metric-title">⏳ Tổng tiền chờ duyệt / Chờ bank</div>
                            <div class="metric-amount">{format_vnd(total_pending)}</div>
                            <div class="metric-count">{count_pending} đơn hàng đang chờ thanh toán</div>
                        </div>
                        <div class="metric-card success">
                            <div class="metric-title">✅ Tổng tiền đã chuyển khoản</div>
                            <div class="metric-amount">{format_vnd(total_banked)}</div>
                            <div class="metric-count">{count_banked} đơn hàng đã thanh toán thành công</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.markdown("##### 📋 Bảng dữ liệu đơn hàng:")
                    
                    st.dataframe(
                        matched_df,
                        hide_index=True,
                        use_container_width=True,
                        column_config={
                            "Tên khách": st.column_config.TextColumn("Tên khách", width="medium"),
                            "Ngày đặt": st.column_config.TextColumn("Ngày đặt (dd/mm/yyyy)", width="small"),
                            "Mã đơn": st.column_config.TextColumn("Mã đơn", width="medium"),
                            "Tổng HH thực tế": st.column_config.NumberColumn(
                                "Tổng HH thực tế",
                                format="%,d ₫",
                                width="medium"
                            ),
                            "Trạng thái": st.column_config.TextColumn("Trạng thái", width="medium")
                        }
                    )
                else:
                    st.warning(f"🔎 Không tìm thấy đơn hàng nào khớp với: **\"{search_query}\"**. Vui lòng kiểm tra lại chính xác Tên hoặc Mã đơn!")
            else:
                st.markdown("""
                <div class="guide-box">
                    <div style="font-size: 38px; margin-bottom: 12px;">🔎</div>
                    <h3 style="color: #1e293b; margin-bottom: 8px; font-weight: 700;">Nhập thông tin để bắt đầu tra cứu</h3>
                    <p style="color: #64748b; font-size: 14px; max-width: 520px; margin: 0 auto 20px;">
                        Hệ thống bảo vệ quyền riêng tư của bạn. Vui lòng nhập Tên hoặc Mã đơn Shopee vào ô tìm kiếm phía trên để hiển thị trạng thái hoàn tiền.
                    </p>
                    <div>
                        <div class="guide-step">
                            <div style="font-weight: 700; color: #1e293b; margin-bottom: 4px;">
                                <span class="guide-step-num">1</span> Lấy Mã Đơn Shopee
                            </div>
                            <div style="font-size: 13px; color: #64748b;">Vào Shopee -> Đơn Mua -> Chi tiết đơn hàng để sao chép Mã đơn.</div>
                        </div>
                        <div class="guide-step">
                            <div style="font-weight: 700; color: #1e293b; margin-bottom: 4px;">
                                <span class="guide-step-num">2</span> Tra Cứu Nhanh
                            </div>
                            <div style="font-size: 13px; color: #64748b;">Nhập Mã đơn hoặc Tên khách hàng vào ô tìm kiếm ở trên.</div>
                        </div>
                        <div class="guide-step">
                            <div style="font-weight: 700; color: #1e293b; margin-bottom: 4px;">
                                <span class="guide-step-num">3</span> Xem Trạng Thái Hoàn Tiền
                            </div>
                            <div style="font-size: 13px; color: #64748b;">Xem tiến độ duyệt và số tiền đã chuyển khoản thực tế.</div>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

    else:
        # =====================================================================
        # B. GIAO DIỆN ADMIN (QUẢN TRỊ VIÊN) - 4 TABS CHUYÊN BIỆT
        # =====================================================================
        st.markdown("""
        <div style="margin-bottom: 24px;">
            <h2 style="color: #ee4d2d; font-weight: 800; margin: 0 0 6px;">👑 Bảng Quản Trị Hệ Thống Shopee Cashback</h2>
            <p style="color: #64748b; margin: 0; font-size: 14px;">Quản lý đơn hàng, đối soát file CSV báo cáo Shopee, quản lý khách hàng và bảng Mapping ngầm.</p>
        </div>
        """, unsafe_allow_html=True)
        
        tab1, tab2, tab3, tab4 = st.tabs([
            "💰 Phần 1: Đối Soát & Chi Trả (Chờ Bank ➔ Đã Bank)",
            "📥 Phần 2: Nạp Đơn Hàng Từ CSV (Hàng Ngày)",
            "📊 Phần 3: Quản Lý & Chỉnh Sửa Dữ Liệu Hàng Loạt",
            "👥 Phần 4: Quản Lý Khách Hàng (Mapping & Đồng Bộ)"
        ])
        
        # ---------------------------------------------------------------------
        # TAB 1: PHẦN 1 - ĐỐI SOÁT HOA HỒNG & CHI TRẢ (CHỜ BANK ➔ ĐÃ BANK)
        # ---------------------------------------------------------------------
        with tab1:
            st.markdown("##### 💰 Đối Soát Hoa Hồng & Quản Lý Chi Trả (Chờ Bank ➔ Đã Bank)")
            st.caption("💡 *Quy trình khép kín: Đối soát file CSV Shopee, tự động cập nhật số tiền Thực nhận chuẩn vào từng Mã đơn hàng, đổi trạng thái sang 'Chờ Bank' (tuyệt đối không tạo dòng tổng), tổng hợp danh sách chuyển khoản theo khách hàng và cập nhật hàng loạt 'Đã Bank'.*")
            
            # -----------------------------------------------------------------
            # MODULE 1: KHU VỰC TẢI FILE CSV & ĐỐI SOÁT (CẬP NHẬT DATABASE)
            # -----------------------------------------------------------------
            with st.expander("📥 1. Tải Lên File Báo Cáo Shopee & Bắt Đầu Đối Soát", expanded=True):
                st.markdown("###### 🗓️ Chọn Khoảng Ngày Hoàn Thành Đơn Hàng:")
                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    default_start = datetime.date.today().replace(day=1)
                    start_date = st.date_input("Từ ngày hoàn thành (start_date):", value=default_start, format="DD/MM/YYYY", key="rec_start_date")
                with col_d2:
                    end_date = st.date_input("Đến ngày hoàn thành (end_date):", value=datetime.date.today(), format="DD/MM/YYYY", key="rec_end_date")
                    
                st.markdown("###### 📁 Tải Lên File Báo Cáo Chuyển Đổi Shopee (.csv):")
                uploaded_file = st.file_uploader(
                    "Chọn file .csv xuất từ Shopee Affiliate:",
                    type=["csv"],
                    help="File CSV báo cáo chuyển đổi Shopee (chứa cột Thời gian hoàn thành đơn hàng, Hoa hồng ròng, Sub_id2...)",
                    key="rec_csv_uploader"
                )
                
                st.markdown("<br>", unsafe_allow_html=True)
                col_btn, _ = st.columns([3, 7])
                with col_btn:
                    btn_start_rec = st.button("🚀 Bắt đầu đối soát & Cập nhật", use_container_width=True, type="primary")
                    
                # XỬ LÝ ĐỐI SOÁT KHI BẤM NÚT
                if btn_start_rec:
                    if uploaded_file is None:
                        st.error("⚠️ Vui lòng tải lên file báo cáo .csv trước khi bấm đối soát!")
                    elif start_date > end_date:
                        st.error("⚠️ Khoảng ngày không hợp lệ! 'Từ ngày' phải nhỏ hơn hoặc bằng 'Đến ngày'.")
                    else:
                        raw_bytes = uploaded_file.getvalue()
                        df_raw = None
                        for enc in ["utf-8-sig", "utf-8", "cp1258", "latin-1"]:
                            try:
                                df_raw = pd.read_csv(io.BytesIO(raw_bytes), encoding=enc, dtype=str)
                                break
                            except Exception:
                                continue
                                
                        if df_raw is None or df_raw.empty:
                            st.error("❌ Không thể đọc nội dung file CSV. Vui lòng kiểm tra lại định dạng file!")
                        else:
                            # Nhận diện cột theo quy tắc nghiêm ngặt
                            col_order_id = None
                            col_status = None
                            col_complete = None
                            col_comm = None
                            col_sub2 = None
                            col_sub1 = None
                            
                            for col in df_raw.columns:
                                c_clean = str(col).strip()
                                c_low = c_clean.lower()
                                if col_order_id is None and any(kw in c_low for kw in ["id đơn hàng", "mã đơn hàng", "mã đơn", "order id", "order sn", "mã đơn shopee"]):
                                    col_order_id = c_clean
                                if col_status is None and ("trạng thái đặt hàng" in c_low or "trạng thái đơn" in c_low or "order status" in c_low or c_low == "trạng thái"):
                                    col_status = c_clean
                                if col_complete is None and ("thời gian hoàn thành" in c_low or "complete time" in c_low or "completed time" in c_low):
                                    col_complete = c_clean
                                if col_comm is None and ("hoa hồng ròng tiếp thị liên kết" in c_low or "hoa hồng ròng" in c_low or "net commission" in c_low or "tổng hoa hồng sản phẩm" in c_low):
                                    if not any(bad in c_low for bad in ["loại", "tỷ lệ", "phí"]):
                                        col_comm = c_clean
                                if col_sub2 is None and ("sub_id2" in c_low or "sub id 2" in c_low or "sub2" in c_low or "mã zalo" in c_low):
                                    col_sub2 = c_clean
                                if col_sub1 is None and ("sub_id1" in c_low or "sub id 1" in c_low or "sub1" in c_low or "tên tạm" in c_low):
                                    col_sub1 = c_clean
                                    
                            missing_cols = []
                            if not col_order_id: missing_cols.append("Mã đơn hàng (Order ID)")
                            if not col_status: missing_cols.append("Trạng thái đặt hàng")
                            if not col_complete: missing_cols.append("Thời gian hoàn thành")
                            if not col_comm: missing_cols.append("Hoa hồng ròng tiếp thị liên kết(₫)")
                            if not col_sub2: missing_cols.append("Sub_id2 (Mã Zalo)")
                            
                            if missing_cols:
                                st.error(f"⚠️ File CSV thiếu các cột bắt buộc: **{', '.join(missing_cols)}**. Vui lòng kiểm tra lại cấu trúc file xuất từ Shopee!")
                                st.write("Các cột hiện có trong file CSV:", list(df_raw.columns))
                            else:
                                # 1. Lọc theo trạng thái 'Hoàn thành' (Bỏ qua hoàn toàn thời gian đặt & click)
                                mask_status = df_raw[col_status].astype(str).str.strip().str.lower() == "hoàn thành"
                                df_completed = df_raw[mask_status].copy()
                                
                                if df_completed.empty:
                                    st.warning("⚠️ Không tìm thấy đơn hàng nào có Trạng thái đặt hàng là 'Hoàn thành' trong file CSV!")
                                else:
                                    # 2. Ép kiểu thời gian hoàn thành sang Date (cắt bỏ giờ/phút/giây)
                                    s_time = df_completed[col_complete].astype(str).str.strip()
                                    df_completed["complete_date"] = pd.to_datetime(s_time, errors="coerce").dt.date
                                    nat_mask = df_completed["complete_date"].isna() & (s_time != "") & (s_time != "nan")
                                    if nat_mask.any():
                                        fallback_dt = pd.to_datetime(s_time[nat_mask], errors="coerce", format="mixed").dt.date
                                        df_completed.loc[nat_mask, "complete_date"] = fallback_dt
                                        
                                    # 3. Lọc start_date <= complete_date <= end_date
                                    mask_date = (df_completed["complete_date"] >= start_date) & (df_completed["complete_date"] <= end_date)
                                    df_valid_reconcile = df_completed[mask_date].copy()
                                    
                                    if df_valid_reconcile.empty:
                                        st.warning(f"⚠️ Không có đơn 'Hoàn thành' nào có Thời gian hoàn thành từ **{start_date.strftime('%d/%m/%Y')}** đến **{end_date.strftime('%d/%m/%Y')}**!")
                                    else:
                                        # 4. Tính Thực nhận = (Hoa hồng ròng * 0.9) * 0.6
                                        raw_comm = df_valid_reconcile[col_comm].astype(str)\
                                            .str.replace("₫", "", regex=False)\
                                            .str.replace("VND", "", regex=False)\
                                            .str.replace("VNĐ", "", regex=False)\
                                            .str.replace(" ", "", regex=False)\
                                            .str.strip()
                                            
                                        def parse_comm_to_float(v):
                                            if not v or v == "nan": return 0.0
                                            if "," in v and "." in v:
                                                v = v.replace(",", "")
                                            elif "," in v:
                                                parts = v.split(",")
                                                if len(parts) == 2 and len(parts[1]) != 3:
                                                    v = v.replace(",", ".")
                                                else:
                                                    v = v.replace(",", "")
                                            try:
                                                return float(v)
                                            except Exception:
                                                return 0.0
                                                
                                        comm_numeric = raw_comm.apply(parse_comm_to_float)
                                        df_valid_reconcile["Thực nhận"] = ((comm_numeric * 0.9) * 0.6).round().astype(int)
                                        
                                        # Làm sạch Order ID, Sub_id2, Sub_id1
                                        df_valid_reconcile["clean_order_id"] = df_valid_reconcile[col_order_id].astype(str).str.strip()
                                        df_valid_reconcile["clean_sub2"] = df_valid_reconcile[col_sub2].fillna("").astype(str).str.strip()
                                        df_valid_reconcile["clean_sub1"] = df_valid_reconcile[col_sub1].fillna("").astype(str).str.strip() if col_sub1 else ""
                                        
                                        # Gộp các dòng cùng Mã đơn (đơn nhiều sản phẩm) và cộng dồn Thực nhận
                                        df_order_grouped = df_valid_reconcile.groupby("clean_order_id", as_index=False).agg(
                                            Thuc_Nhan=("Thực nhận", "sum"),
                                            Complete_Date=("complete_date", "first"),
                                            Sub_id2=("clean_sub2", "first"),
                                            Sub_id1=("clean_sub1", "first")
                                        )
                                        
                                        with st.spinner("Đang cập nhật dữ liệu vào Google Sheets (UPDATE theo Mã đơn)..."):
                                            # Load Mapping Data
                                            df_mapping_curr, _, _ = load_mapping_data(st.session_state["sheet_name"])
                                            mapping_dict = dict(zip(
                                                df_mapping_curr["Mã Zalo (Sub_id2)"].astype(str).str.strip(),
                                                df_mapping_curr["Tên Khách Hàng"].astype(str).str.strip()
                                            ))
                                            
                                            # A. Auto-Mapping khách mới: Quét Sub_id2 mới
                                            new_mapping_rows = []
                                            for _, r in df_order_grouped.iterrows():
                                                z = str(r["Sub_id2"]).strip()
                                                s1 = str(r["Sub_id1"]).strip()
                                                if z and z != "(Không có Sub_id2)" and (z not in mapping_dict):
                                                    temp_name = s1 if s1 else z
                                                    new_mapping_rows.append({"Mã Zalo (Sub_id2)": z, "Tên Khách Hàng": temp_name})
                                                    mapping_dict[z] = temp_name
                                                    
                                            if new_mapping_rows:
                                                df_new_map_df = pd.DataFrame(new_mapping_rows).drop_duplicates(subset=["Mã Zalo (Sub_id2)"])
                                                updated_mapping_all = pd.concat([df_mapping_curr, df_new_map_df], ignore_index=True)
                                                save_mapping_data(updated_mapping_all, sheet_target=st.session_state["sheet_name"])
                                                
                                            # B. Cập nhật Shopee_Cashback theo Mã đơn hàng (UPDATE, KHÔNG INSERT dòng tổng)
                                            df_orders_curr, _, _ = load_data(st.session_state["sheet_name"])
                                            
                                            count_orders_updated = 0
                                            count_orders_added = 0
                                            total_payout_reconciled = 0
                                            
                                            # Tạo index mã đơn hiện tại trong Shopee_Cashback để dò tìm
                                            existing_code_indices = {}
                                            for idx, r_ord in df_orders_curr.iterrows():
                                                c_code = str(r_ord["Mã đơn"]).strip()
                                                if c_code not in existing_code_indices:
                                                    existing_code_indices[c_code] = []
                                                existing_code_indices[c_code].append(idx)
                                                
                                            new_orders_to_insert = []
                                            
                                            for _, r in df_order_grouped.iterrows():
                                                order_code = str(r["clean_order_id"]).strip()
                                                thuc_nhan_amt = int(r["Thuc_Nhan"])
                                                order_comp_date = r["Complete_Date"]
                                                z_code = str(r["Sub_id2"]).strip()
                                                s1_code = str(r["Sub_id1"]).strip()
                                                
                                                resolved_cust_name = mapping_dict.get(z_code, s1_code if s1_code else (z_code if z_code else "Khách mới"))
                                                
                                                if order_code in existing_code_indices:
                                                    # DÒ TÌM KHỚP MÃ ĐƠN HÀNG:
                                                    # Hành động 1: Ghi đè số tiền Thực nhận chuẩn
                                                    # Hành động 2: Cập nhật Trạng thái thành 'Chờ Bank'
                                                    target_idx = existing_code_indices[order_code][0]
                                                    df_orders_curr.at[target_idx, "Tổng HH thực tế"] = thuc_nhan_amt
                                                    df_orders_curr.at[target_idx, "Trạng thái"] = "Chờ Bank"
                                                    if resolved_cust_name and resolved_cust_name != "Khách mới":
                                                        df_orders_curr.at[target_idx, "Tên khách"] = resolved_cust_name
                                                    count_orders_updated += 1
                                                else:
                                                    # Đơn hoàn thành chưa có trong hệ thống: Thêm đơn lẻ đó vào với trạng thái 'Chờ Bank' (Tuyệt đối không tạo dòng tổng)
                                                    new_orders_to_insert.append({
                                                        "Tên khách": resolved_cust_name,
                                                        "Ngày đặt": order_comp_date.strftime("%d/%m/%Y"),
                                                        "Mã đơn": order_code,
                                                        "Tổng HH thực tế": thuc_nhan_amt,
                                                        "Trạng thái": "Chờ Bank"
                                                    })
                                                    count_orders_added += 1
                                                    
                                                total_payout_reconciled += thuc_nhan_amt
                                                
                                            if new_orders_to_insert:
                                                df_orders_curr = pd.concat([df_orders_curr, pd.DataFrame(new_orders_to_insert)], ignore_index=True)
                                                
                                            # Lưu toàn bộ vào Shopee_Cashback (Google Sheets + CSV cục bộ)
                                            save_all_to_storage(df_orders_curr, sheet_target=st.session_state["sheet_name"])
                                            
                                        st.session_state["reconcile_success_msg"] = (
                                            f"Đã đối soát hoàn tất! Cập nhật **{count_orders_updated}** đơn hàng cũ (ghi đè Thực nhận chuẩn & đổi sang 'Chờ Bank'), "
                                            f"bổ sung **{count_orders_added}** đơn mới hợp lệ. Tự động mapping **{len(new_mapping_rows)}** khách mới. "
                                            f"Tổng tiền Thực nhận chốt kỳ: **{format_vnd(total_payout_reconciled)}**."
                                        )
                                        st.toast("Đối soát & Cập nhật Google Sheets thành công!", icon="🎉")
                                        st.rerun()

            # Thông báo đối soát vừa thực hiện xong
            if "reconcile_success_msg" in st.session_state:
                st.success(f"🎉 **KẾT QUẢ ĐỐI SOÁT:** {st.session_state['reconcile_success_msg']}")
                del st.session_state["reconcile_success_msg"]
                
            st.divider()

            # -----------------------------------------------------------------
            # MODULE 2: DASHBOARD CHI TRẢ & BULK UPDATE 'ĐÃ BANK'
            # -----------------------------------------------------------------
            # Lấy toàn bộ danh sách đơn hàng đang ở trạng thái 'Chờ Bank'
            df_pending_bank = df_data[df_data["Trạng thái"] == "Chờ Bank"].copy()
            
            if df_pending_bank.empty:
                st.info("ℹ️ **Hiện tại không có đơn hàng nào đang ở trạng thái 'Chờ Bank'.**\n\n*(Tất cả đơn hàng đã được chuyển khoản 'Đã bank' hoặc đang ở trạng thái 'Chờ Shopee duyệt'). Bạn có thể tải file CSV ở trên để bắt đầu đối soát kỳ mới.*")
            else:
                total_pending_amount = int(df_pending_bank["Tổng HH thực tế"].sum())
                total_pending_orders = len(df_pending_bank)
                total_pending_custs = df_pending_bank["Tên khách"].nunique()
                
                # Thẻ thống kê KPI
                st.markdown(f"""
                <div class="metrics-container">
                    <div class="metric-card warning">
                        <div class="metric-title">💳 Tổng Tiền Cần Bank Trong Kỳ</div>
                        <div class="metric-amount">{format_vnd(total_pending_amount)}</div>
                        <div class="metric-count">Đã chốt thực nhận (60%) cho các đơn Hoàn thành</div>
                    </div>
                    <div class="metric-card info">
                        <div class="metric-title">👥 Khách Hàng Cần Chuyển Khoản</div>
                        <div class="metric-amount">{total_pending_custs} người</div>
                        <div class="metric-count">Danh sách chi tiết ở Bảng Khu vực 1 bên dưới</div>
                    </div>
                    <div class="metric-card pending">
                        <div class="metric-title">📦 Tổng Số Đơn Hàng 'Chờ Bank'</div>
                        <div class="metric-amount">{total_pending_orders} đơn</div>
                        <div class="metric-count">Cần thanh toán hoa hồng cho khách</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                # -------------------------------------------------------------
                # KHU VỰC 1: BẢNG TỔNG HỢP CHUYỂN KHOẢN (GROUP BY KHÁCH HÀNG)
                # -------------------------------------------------------------
                st.markdown("#### 🏢 Khu Vực 1: Bảng Tổng Hợp Chuyển Khoản (Theo Khách Hàng)")
                st.caption("📱 *Admin mở App ngân hàng, nhìn vào danh sách dưới đây để chuyển khoản cho từng khách hàng.*")
                
                # Gom nhóm theo Tên khách
                grouped_bank = df_pending_bank.groupby("Tên khách", as_index=False).agg(
                    So_Don=("Mã đơn", "count"),
                    Tong_Tien=("Tổng HH thực tế", "sum")
                )
                
                # Tra mã Zalo (Sub_id2) từ df_mapping
                map_zalo_dict = dict(zip(
                    df_mapping["Tên Khách Hàng"].astype(str).str.strip(),
                    df_mapping["Mã Zalo (Sub_id2)"].astype(str).str.strip()
                ))
                grouped_bank["Mã Zalo (Sub_id2)"] = grouped_bank["Tên khách"].map(lambda x: map_zalo_dict.get(str(x).strip(), "(Chưa có)"))
                grouped_bank = grouped_bank.sort_values(by="Tong_Tien", ascending=False)
                
                grouped_display_df = grouped_bank[["Tên khách", "Mã Zalo (Sub_id2)", "So_Don", "Tong_Tien"]].rename(columns={
                    "Tên khách": "Tên Khách Hàng",
                    "So_Don": "Số lượng đơn",
                    "Tong_Tien": "Tổng số tiền cần Bank trong kỳ này"
                })
                
                st.dataframe(
                    grouped_display_df,
                    hide_index=True,
                    use_container_width=True,
                    column_config={
                        "Tên Khách Hàng": st.column_config.TextColumn("Tên Khách Hàng", width="medium"),
                        "Mã Zalo (Sub_id2)": st.column_config.TextColumn("Mã Zalo (Sub_id2)", width="medium"),
                        "Số lượng đơn": st.column_config.NumberColumn("Số lượng đơn", width="small"),
                        "Tổng số tiền cần Bank trong kỳ này": st.column_config.NumberColumn(
                            "Tổng số tiền cần Bank trong kỳ này",
                            format="%,d ₫",
                            width="medium"
                        )
                    }
                )
                
                st.divider()
                
                # -------------------------------------------------------------
                # KHU VỰC 2: THAO TÁC HÀNG LOẠT (BULK UPDATE "ĐÃ BANK")
                # -------------------------------------------------------------
                st.markdown("#### ⚡ Khu Vực 2: Thao Tác Hàng Loạt (Cập Nhật 'Đã Bank')")
                st.caption("💡 *Sau khi Admin đã chuyển khoản xong, tích chọn các đơn hàng lẻ bên dưới và bấm **CẬP NHẬT ĐÃ BANK** để đổi trạng thái trên Google Sheets.*")
                
                # Điều khiển Check-all và Lọc theo khách hàng
                ctrl_col1, ctrl_col2 = st.columns([1, 2])
                with ctrl_col1:
                    check_all_active = st.checkbox("✅ Chọn tất cả các đơn", value=False, key="chk_all_pending_bank")
                with ctrl_col2:
                    list_pending_custs = ["Tất cả khách hàng"] + sorted(df_pending_bank["Tên khách"].unique().tolist())
                    sel_cust_view = st.selectbox("Lọc chi tiết theo từng khách:", list_pending_custs, index=0, key="filter_pending_cust")
                    
                df_detail_view = df_pending_bank.copy()
                if sel_cust_view != "Tất cả khách hàng":
                    df_detail_view = df_detail_view[df_detail_view["Tên khách"] == sel_cust_view]
                    
                df_detail_view.insert(0, "Chọn", check_all_active)
                
                edited_pending_table = st.data_editor(
                    df_detail_view[["Chọn", "Tên khách", "Mã đơn", "Ngày đặt", "Tổng HH thực tế", "Trạng thái"]],
                    hide_index=True,
                    use_container_width=True,
                    disabled=["Tên khách", "Mã đơn", "Ngày đặt", "Tổng HH thực tế", "Trạng thái"],
                    column_config={
                        "Chọn": st.column_config.CheckboxColumn("Chọn", help="Tích chọn để cập nhật Đã Bank", default=check_all_active),
                        "Tên khách": st.column_config.TextColumn("Tên khách", width="medium"),
                        "Mã đơn": st.column_config.TextColumn("Mã đơn", width="medium"),
                        "Ngày đặt": st.column_config.TextColumn("Ngày đặt", width="small"),
                        "Tổng HH thực tế": st.column_config.NumberColumn("Thực nhận (VNĐ)", format="%,d ₫", width="medium"),
                        "Trạng thái": st.column_config.TextColumn("Trạng thái", width="small")
                    },
                    key=f"editor_bulk_bank_{check_all_active}_{sel_cust_view}"
                )
                
                st.markdown("<br>", unsafe_allow_html=True)
                col_btn_paid, _ = st.columns([3, 7])
                with col_btn_paid:
                    btn_bulk_paid = st.button("💸 CẬP NHẬT ĐÃ BANK CHO CÁC ĐƠN ĐÃ CHỌN", use_container_width=True, type="primary")
                    
                if btn_bulk_paid:
                    selected_rows = edited_pending_table[edited_pending_table["Chọn"] == True]
                    if selected_rows.empty:
                        st.warning("⚠️ Vui lòng tích chọn ít nhất 1 đơn hàng để cập nhật trạng thái 'Đã bank'!")
                    else:
                        selected_codes = set(selected_rows["Mã đơn"].astype(str).str.strip())
                        mask_orders_to_paid = df_data["Mã đơn"].astype(str).str.strip().isin(selected_codes) & (df_data["Trạng thái"] == "Chờ Bank")
                        num_paid = mask_orders_to_paid.sum()
                        
                        df_data.loc[mask_orders_to_paid, "Trạng thái"] = "Đã bank"
                        ok, err = save_all_to_storage(df_data, sheet_target=st.session_state["sheet_name"])
                        
                        if ok:
                            st.toast(f"Đã cập nhật {num_paid} đơn hàng sang 'Đã bank'!", icon="💸")
                            st.success(f"🎉 **Thành công!** Đã cập nhật **{num_paid}** đơn hàng sang trạng thái **'Đã bank'** trên Google Sheets!")
                        else:
                            st.toast("Đã cập nhật vào CSV cục bộ!", icon="💾")
                            st.warning(f"✅ Đã cập nhật vào CSV cục bộ. (Google Sheets: {err})")
                        st.rerun()

        # ---------------------------------------------------------------------
        # TAB 2: PHẦN 2 - NẠP ĐƠN HÀNG TỪ CSV (HÀNG NGÀY)
        # ---------------------------------------------------------------------
        with tab2:
            st.markdown("##### 📥 Phần 2: Nạp Đơn Hàng Mới Từ File CSV Báo Cáo Shopee")
            st.caption("💡 *Tải lên file CSV báo cáo chuyển đổi Shopee hàng ngày. Tự động gom sản phẩm theo mã đơn, lọc đơn hợp lệ, nhận diện khách cũ từ Mapping_Data hoặc thêm khách mới, và nạp nhanh vào Shopee_Cashback để khách tra cứu.*")
            
            st.markdown("###### 🗓️ 1. Thiết Lập Điều Kiện Nạp Đơn:")
            col_d1, col_d2, col_d3 = st.columns([2, 2, 2])
            with col_d1:
                default_start = datetime.date.today() - datetime.timedelta(days=7)
                daily_start_date = st.date_input("Từ ngày đặt:", value=default_start, format="DD/MM/YYYY", key="daily_start_date")
            with col_d2:
                daily_end_date = st.date_input("Đến ngày đặt:", value=datetime.date.today(), format="DD/MM/YYYY", key="daily_end_date")
            with col_d3:
                daily_import_status = st.selectbox("Trạng thái ghi nhận:", options=STATUS_OPTIONS, index=0, key="daily_import_status")
                
            daily_ignore_date = st.checkbox("Nạp toàn bộ đơn trong file (không lọc theo khoảng ngày ở trên)", value=False, key="daily_ignore_date")
            
            st.markdown("###### 📁 2. Tải Lên File Báo Cáo Shopee (.csv):")
            uploaded_daily_file = st.file_uploader(
                "Chọn file CSV báo cáo từ Shopee Affiliate:",
                type=["csv"],
                help="Chọn file CSV báo cáo chuyển đổi / đơn hàng xuất từ Shopee Affiliate",
                key="daily_csv_uploader"
            )
            
            if uploaded_daily_file is not None:
                raw_bytes = uploaded_daily_file.getvalue()
                df_raw_csv = None
                for enc in ["utf-8-sig", "utf-8", "cp1258", "latin-1"]:
                    try:
                        df_raw_csv = pd.read_csv(io.BytesIO(raw_bytes), encoding=enc, dtype=str)
                        break
                    except Exception:
                        continue
                        
                if df_raw_csv is None or df_raw_csv.empty:
                    st.error("❌ Không thể đọc nội dung file CSV. Vui lòng kiểm tra lại định dạng file!")
                else:
                    col_map = detect_shopee_csv_columns(df_raw_csv)
                    
                    time_col_detected = None
                    if col_map["order_time"] and df_raw_csv[col_map["order_time"]].dropna().str.strip().ne("").any():
                        time_col_detected = col_map["order_time"]
                        time_col_type_label = "Thời gian đặt hàng"
                    elif col_map["complete_time"]:
                        time_col_detected = col_map["complete_time"]
                        time_col_type_label = "Thời gian hoàn thành"
                    else:
                        time_col_detected = col_map["order_time"]
                        time_col_type_label = "Thời gian đặt hàng"
                    
                    missing_cols = []
                    if not col_map["order_id"]: missing_cols.append("Mã đơn hàng (ID đơn hàng / Order ID)")
                    if not time_col_detected and not daily_ignore_date: missing_cols.append("Thời gian đặt hàng / hoàn thành")
                    if not col_map["status"]: missing_cols.append("Trạng thái (Order Status)")
                    if not col_map["commission"]: missing_cols.append("Hoa hồng thực tế (Commission)")
                    
                    if missing_cols:
                        st.error(f"⚠️ File CSV thiếu các cột bắt buộc: {', '.join(missing_cols)}. Vui lòng kiểm tra lại file xuất từ Shopee!")
                        st.write("Các cột hiện có trong file CSV của bạn:", list(df_raw_csv.columns))
                    else:
                        with st.expander("🔍 Chi tiết nhận diện các cột Shopee:", expanded=False):
                            st.write(f"- Cột Mã đơn: `{col_map['order_id']}`")
                            st.write(f"- Cột Thời gian: `{time_col_detected}` ({time_col_type_label})")
                            st.write(f"- Cột Trạng thái: `{col_map['status']}`")
                            st.write(f"- Cột Hoa hồng: `{col_map['commission']}`")
                            st.write(f"- Cột Sub_id1 (Tên tạm): `{col_map['sub1'] or 'Không tìm thấy'}`")
                            st.write(f"- Cột Sub_id2 (Mã Zalo): `{col_map['sub2'] or 'Không tìm thấy'}`")
                            
                        order_id_col = col_map["order_id"]
                        status_col = col_map["status"]
                        commission_col = col_map["commission"]
                        sub1_col = col_map["sub1"]
                        sub2_col = col_map["sub2"]
                        
                        mapping_dict = dict(zip(
                            df_mapping["Mã Zalo (Sub_id2)"].astype(str).str.strip(),
                            df_mapping["Tên Khách Hàng"].astype(str).str.strip()
                        ))
                        
                        existing_order_ids = set(df_data["Mã đơn"].astype(str).str.strip())
                        
                        valid_orders = []
                        new_mappings_discovered = {}
                        count_skipped_status = 0
                        count_skipped_date = 0
                        count_skipped_duplicate = 0
                        
                        for _, row in df_raw_csv.iterrows():
                            status_val = str(row.get(status_col, "")).lower().strip()
                            if any(bad in status_val for bad in ["hủy", "huy", "cancel", "bị hủy"]):
                                count_skipped_status += 1
                                continue
                            if not any(good in status_val for good in ["chờ xử lý", "cho xu ly", "hoàn thành", "hoan thanh", "complete", "pending", "thành công"]):
                                count_skipped_status += 1
                                continue
                                
                            c_date = None
                            if time_col_detected:
                                c_date = parse_flexible_date(row.get(time_col_detected))
                                
                            if not daily_ignore_date:
                                if not c_date:
                                    count_skipped_date += 1
                                    continue
                                if not (daily_start_date <= c_date <= daily_end_date):
                                    count_skipped_date += 1
                                    continue
                                    
                            code = str(row.get(order_id_col, "")).strip()
                            if not code:
                                continue
                                
                            comm = parse_commission_number(row.get(commission_col, 0))
                            
                            if code in existing_order_ids:
                                count_skipped_duplicate += 1
                                continue
                                
                            zalo_id = str(row.get(sub2_col, "")).strip() if sub2_col else ""
                            sub1_val = str(row.get(sub1_col, "")).strip() if sub1_col else ""
                            
                            if zalo_id and (zalo_id in mapping_dict) and mapping_dict[zalo_id]:
                                customer_assigned_name = mapping_dict[zalo_id]
                                customer_type_label = "Khách cũ"
                            else:
                                temp_name = sub1_val if sub1_val else (zalo_id if zalo_id else "Khách mới")
                                customer_assigned_name = temp_name
                                customer_type_label = "Khách mới (Tên tạm)"
                                if zalo_id:
                                    new_mappings_discovered[zalo_id] = temp_name
                                    mapping_dict[zalo_id] = temp_name
                                    
                            order_date_str = c_date.strftime("%d/%m/%Y") if c_date else datetime.date.today().strftime("%d/%m/%Y")
                            
                            valid_orders.append({
                                "Mã đơn": code,
                                "Tên khách": customer_assigned_name,
                                "Loại khách": customer_type_label,
                                "Mã Zalo (Sub_id2)": zalo_id if zalo_id else "(Không có)",
                                "Ngày đặt": order_date_str,
                                "Tổng HH thực tế": comm,
                                "Trạng thái": daily_import_status
                            })
                            
                        if valid_orders:
                            df_valid = pd.DataFrame(valid_orders)
                            df_grouped = df_valid.groupby("Mã đơn", as_index=False).agg({
                                "Tên khách": "first",
                                "Loại khách": "first",
                                "Mã Zalo (Sub_id2)": "first",
                                "Ngày đặt": "first",
                                "Tổng HH thực tế": "sum",
                                "Trạng thái": "first"
                            })
                        else:
                            df_grouped = pd.DataFrame()
                            
                        st.divider()
                        st.markdown("#### 📊 Kết Quả Lọc Đơn Hàng")
                        
                        total_valid_count = len(df_grouped)
                        total_valid_comm = int(df_grouped["Tổng HH thực tế"].sum()) if total_valid_count > 0 else 0
                        count_new_cust = len(new_mappings_discovered)
                        count_old_orders = len(df_grouped[df_grouped["Loại khách"] == "Khách cũ"]) if total_valid_count > 0 else 0
                        
                        st.markdown(f"""
                        <div class="metrics-container">
                            <div class="metric-card success">
                                <div class="metric-title">📦 Đơn Hợp Lệ Sẵn Sàng Nạp</div>
                                <div class="metric-amount">{total_valid_count} đơn</div>
                                <div class="metric-count">Đã gộp theo mã đơn duy nhất</div>
                            </div>
                            <div class="metric-card warning">
                                <div class="metric-title">💰 Tổng Hoa Hồng Đơn Hàng</div>
                                <div class="metric-amount">{format_vnd(total_valid_comm)}</div>
                                <div class="metric-count">Hoa hồng đơn hàng mới</div>
                            </div>
                            <div class="metric-card info">
                                <div class="metric-title">👤 Đơn Khách Cũ (Đã Khớp Tên)</div>
                                <div class="metric-amount">{count_old_orders} đơn</div>
                                <div class="metric-count">Khớp mã Zalo từ Mapping_Data</div>
                            </div>
                            <div class="metric-card pending">
                                <div class="metric-title">🆕 Khách Hàng Mới Phát Hiện</div>
                                <div class="metric-amount">{count_new_cust} khách</div>
                                <div class="metric-count">Sẽ tự thêm vào bảng Mapping</div>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        filter_notes = []
                        if count_skipped_status > 0:
                            filter_notes.append(f"Đã bỏ qua **{count_skipped_status}** dòng đơn bị hủy")
                        if count_skipped_date > 0:
                            filter_notes.append(f"Đã bỏ qua **{count_skipped_date}** dòng ngoài khoảng ngày")
                        if count_skipped_duplicate > 0:
                            filter_notes.append(f"Đã bỏ qua **{count_skipped_duplicate}** đơn đã có trong hệ thống (chống trùng)")
                        if filter_notes:
                            st.info("ℹ️ " + " | ".join(filter_notes))
                            
                        if total_valid_count > 0:
                            if new_mappings_discovered:
                                with st.expander(f"🆕 Danh sách {len(new_mappings_discovered)} khách hàng mới sẽ tự thêm vào Mapping_Data:", expanded=True):
                                    df_new_cust_preview = pd.DataFrame([
                                        {"Mã Zalo (Sub_id2)": z, "Tên tạm gán (Sub_id1)": n}
                                        for z, n in new_mappings_discovered.items()
                                    ])
                                    st.dataframe(df_new_cust_preview, hide_index=True, use_container_width=True)
                                    st.caption("💡 *Sau khi nạp, bạn có thể chuyển sang **Phần 5: Quản Lý Khách Hàng** để đổi tên chính xác và đồng bộ hàng loạt.*")
                                    
                            st.markdown("##### 📋 Xem trước danh sách đơn hàng sẽ nạp vào hệ thống:")
                            st.dataframe(
                                df_grouped,
                                hide_index=True,
                                use_container_width=True,
                                column_config={
                                    "Mã đơn": st.column_config.TextColumn("Mã đơn", width="medium"),
                                    "Tên khách": st.column_config.TextColumn("Tên khách ghi nhận", width="medium"),
                                    "Loại khách": st.column_config.TextColumn("Loại khách", width="small"),
                                    "Mã Zalo (Sub_id2)": st.column_config.TextColumn("Mã Zalo", width="small"),
                                    "Ngày đặt": st.column_config.TextColumn("Ngày đặt", width="small"),
                                    "Tổng HH thực tế": st.column_config.NumberColumn("Hoa hồng thực tế", format="%,d ₫", width="medium"),
                                    "Trạng thái": st.column_config.TextColumn("Trạng thái", width="small")
                                }
                            )
                            
                            st.markdown("<br>", unsafe_allow_html=True)
                            col_confirm, _ = st.columns([3, 7])
                            with col_confirm:
                                confirm_daily_import_btn = st.button("🚀 XÁC NHẬN NẠP VÀO HỆ THỐNG", use_container_width=True, type="primary")
                                
                            if confirm_daily_import_btn:
                                with st.spinner("Đang đồng bộ dữ liệu lên Google Sheets..."):
                                    if new_mappings_discovered:
                                        batch_add_mapping_records(new_mappings_discovered, sheet_target=st.session_state["sheet_name"])
                                        
                                    orders_to_add = []
                                    for _, r in df_grouped.iterrows():
                                        orders_to_add.append({
                                            "Tên khách": r["Tên khách"],
                                            "Ngày đặt": r["Ngày đặt"],
                                            "Mã đơn": r["Mã đơn"],
                                            "Tổng HH thực tế": int(r["Tổng HH thực tế"]),
                                            "Trạng thái": r["Trạng thái"]
                                        })
                                    batch_add_orders_to_storage(orders_to_add, sheet_target=st.session_state["sheet_name"])
                                    
                                st.toast(f"Đã nạp thành công {total_valid_count} đơn hàng vào hệ thống!", icon="🎉")
                                st.success(f"🎉 Hoàn tất nạp đơn! Đã nạp thành công **{total_valid_count}** đơn hàng ({format_vnd(total_valid_comm)}) vào `Shopee_Cashback` và cập nhật **{count_new_cust}** khách hàng mới vào `Mapping_Data`!")
                                st.rerun()
                        else:
                            st.warning("⚠️ Không tìm thấy đơn hàng nào hợp lệ để nạp!")

        # ---------------------------------------------------------------------
        # TAB 3: PHẦN 3 - CẬP NHẬT DỮ LIỆU HÀNG LOẠT (DATA EDITOR)
        # ---------------------------------------------------------------------
        with tab3:
            st.markdown("##### 📊 Phần 3: Quản Lý & Chỉnh Sửa Dữ Liệu Hàng Loạt")
            st.caption("💡 *Chỉnh sửa trực tiếp trên bảng Data Editor bên dưới, sau đó bấm **LƯU THAY ĐỔI** để cập nhật lên Google Sheets.*")
            
            c_search, c_status = st.columns([3, 2])
            with c_search:
                filter_text = st.text_input(
                    "Tra nhanh đơn cần cập nhật:",
                    placeholder="🔍 Nhập Tên khách hoặc Mã đơn cần tìm...",
                    key="admin_search_bar"
                ).strip()
                
            with c_status:
                filter_status = st.selectbox(
                    "Lọc theo trạng thái:",
                    ["Tất cả"] + STATUS_OPTIONS,
                    index=0,
                    key="admin_status_filter"
                )
            
            display_df = df_data.copy()
            if filter_text:
                display_df = display_df[
                    display_df["Tên khách"].str.contains(filter_text, case=False, na=False) |
                    display_df["Mã đơn"].str.contains(filter_text, case=False, na=False)
                ]
            if filter_status != "Tất cả":
                display_df = display_df[display_df["Trạng thái"] == filter_status]
            
            st.markdown(f"**Tổng số đơn hiển thị:** `{len(display_df)}` / `{len(df_data)}` đơn hàng.")
            
            edited_table = st.data_editor(
                display_df,
                hide_index=True,
                use_container_width=True,
                num_rows="dynamic",
                column_config={
                    "Tên khách": st.column_config.TextColumn("Tên khách", required=True, width="medium"),
                    "Ngày đặt": st.column_config.TextColumn("Ngày đặt (dd/mm/yyyy)", required=True, width="small"),
                    "Mã đơn": st.column_config.TextColumn("Mã đơn (Text)", required=True, width="medium"),
                    "Tổng HH thực tế": st.column_config.NumberColumn(
                        "Tổng HH thực tế (VNĐ)",
                        format="%,d ₫",
                        required=True,
                        width="medium"
                    ),
                    "Trạng thái": st.column_config.SelectboxColumn(
                        "Trạng thái",
                        options=STATUS_OPTIONS,
                        required=True,
                        width="medium"
                    )
                },
                key="shopee_batch_data_editor"
            )
            
            st.markdown("<br>", unsafe_allow_html=True)
            btn_col, _ = st.columns([3, 7])
            with btn_col:
                save_batch_btn = st.button("💾 LƯU THAY ĐỔI", use_container_width=True)
                
            if save_batch_btn:
                if filter_text or filter_status != "Tất cả":
                    merged_df = df_data.copy()
                    for idx in edited_table.index:
                        if idx in merged_df.index:
                            merged_df.loc[idx] = edited_table.loc[idx]
                    save_target_df = merged_df
                else:
                    save_target_df = edited_table
                    
                ok, err = save_all_to_storage(save_target_df, sheet_target=st.session_state["sheet_name"])
                if ok:
                    st.toast("Đã lưu các thay đổi lên Google Sheets & CSV thành công!", icon="💾")
                    st.success("✅ Dữ liệu đã được cập nhật thành công lên **Google Sheets**!")
                else:
                    st.toast("Đã lưu vào CSV cục bộ!", icon="💾")
                    st.warning(f"✅ Dữ liệu đã lưu vào CSV cục bộ. (Google Sheets: {err})")
                st.rerun()

        # ---------------------------------------------------------------------
        # TAB 4: QUẢN LÝ KHÁCH HÀNG & ĐỒNG BỘ TÊN (PHẦN 4)
        # ---------------------------------------------------------------------
        with tab4:
            st.markdown("##### 👥 Phần 4: Quản Lý Khách Hàng & Đồng Bộ Tên Tự Động")
            st.caption("💡 *Đọc dữ liệu từ bảng **Mapping_Data**. Khi bạn sửa tên khách hàng tại đây, hệ thống sẽ **tự động cập nhật vào bảng Mapping** và **quét sửa toàn bộ các đơn hàng cũ mang tên cũ trong Shopee_Cashback**.*")
            
            # Tính toán thống kê đơn hàng cho từng khách trong Mapping
            customer_stats = {}
            for _, r in df_data.iterrows():
                c_name = str(r["Tên khách"]).strip()
                c_amt = int(r["Tổng HH thực tế"])
                if c_name not in customer_stats:
                    customer_stats[c_name] = {"count": 0, "total": 0}
                customer_stats[c_name]["count"] += 1
                customer_stats[c_name]["total"] += c_amt
                
            # Tạo DataFrame Mapping có thêm cột thống kê
            mapping_display_list = []
            for _, r in df_mapping.iterrows():
                zalo = str(r["Mã Zalo (Sub_id2)"]).strip()
                name = str(r["Tên Khách Hàng"]).strip()
                stats = customer_stats.get(name, {"count": 0, "total": 0})
                mapping_display_list.append({
                    "Mã Zalo (Sub_id2)": zalo,
                    "Tên Khách Hàng": name,
                    "Số đơn hàng": stats["count"],
                    "Tổng tiền hoa hồng (VNĐ)": stats["total"]
                })
            df_mapping_stats = pd.DataFrame(mapping_display_list)
            
            # PHẦN A: CÔNG CỤ SỬA VÀ ĐỒNG BỘ TÊN NHANH
            st.markdown("###### 🔄 Đổi Tên Chuẩn & Đồng Bộ Toàn Bộ Đơn Hàng Cũ:")
            
            if not df_mapping.empty:
                # Tạo danh sách lựa chọn khách hàng
                cust_options = []
                for _, r in df_mapping.iterrows():
                    z = str(r["Mã Zalo (Sub_id2)"]).strip()
                    n = str(r["Tên Khách Hàng"]).strip()
                    s = customer_stats.get(n, {"count": 0, "total": 0})
                    cust_options.append(f"{n} (Zalo: {z}) - [{s['count']} đơn, {format_vnd(s['total'])}]")
                    
                selected_option = st.selectbox(
                    "Chọn khách hàng cần chuẩn hóa tên:",
                    options=cust_options,
                    help="Chọn khách hàng từ bảng Mapping để đổi tên và đồng bộ đơn hàng"
                )
                
                # Trích xuất Zalo và Tên hiện tại từ option đã chọn
                idx_sel = cust_options.index(selected_option)
                selected_zalo = df_mapping.iloc[idx_sel]["Mã Zalo (Sub_id2)"]
                current_selected_name = df_mapping.iloc[idx_sel]["Tên Khách Hàng"]
                curr_orders_count = customer_stats.get(current_selected_name, {"count": 0})["count"]
                
                col_rename1, col_rename2, col_rename3 = st.columns([3, 3, 2])
                with col_rename1:
                    st.text_input("Tên hiện tại trong hệ thống:", value=current_selected_name, disabled=True)
                with col_rename2:
                    new_correct_name = st.text_input("Nhập tên chuẩn mới:", placeholder="Ví dụ: Hương Rosy, Nguyễn Văn A...", key="input_new_cust_name")
                with col_rename3:
                    st.markdown("<br>", unsafe_allow_html=True)
                    sync_name_btn = st.button("🔄 LƯU & ĐỒNG BỘ", use_container_width=True)
                    
                if sync_name_btn:
                    if not new_correct_name.strip():
                        st.error("⚠️ Vui lòng nhập tên chuẩn mới!")
                    elif new_correct_name.strip() == current_selected_name.strip():
                        st.warning("⚠️ Tên mới trùng với tên cũ. Không có thay đổi nào được thực hiện.")
                    else:
                        with st.spinner("Đang đồng bộ tên trên Mapping_Data và Shopee_Cashback..."):
                            ok, num_synced, err = sync_customer_name(
                                old_name=current_selected_name,
                                new_name=new_correct_name,
                                target_sub2=selected_zalo,
                                sheet_target=st.session_state["sheet_name"]
                            )
                        if ok:
                            st.toast(f"Đã đổi '{current_selected_name}' -> '{new_correct_name}' và đồng bộ {num_synced} đơn hàng!", icon="✨")
                            st.success(f"🎉 **Thành công!** Đã đổi tên thành công trong bảng `Mapping_Data` và tự động cập nhật **{num_synced}** đơn hàng cũ trong bảng `Shopee_Cashback` sang tên mới **'{new_correct_name}'**!")
                            st.rerun()
                        else:
                            st.error(f"Lỗi khi đồng bộ tên: {err}")
            else:
                st.info("Bảng Mapping hiện chưa có dữ liệu. Bạn có thể thêm khách mới bên dưới hoặc Import từ file CSV Shopee.")
                
            st.divider()
            
            # PHẦN B: BẢNG DỮ LIỆU MAPPING & CHỈNH SỬA TRỰC TIẾP
            st.markdown("###### 📋 Bảng Dữ Liệu Mapping Khách Hàng (Mã Zalo ↔ Tên Chuẩn):")
            st.caption("💡 *Bạn cũng có thể sửa trực tiếp trên bảng dưới đây hoặc thêm khách mới, sau đó nhấn **LƯU BẢNG MAPPING**.*")
            
            edited_mapping_table = st.data_editor(
                df_mapping_stats if not df_mapping_stats.empty else pd.DataFrame(columns=["Mã Zalo (Sub_id2)", "Tên Khách Hàng", "Số đơn hàng", "Tổng tiền hoa hồng (VNĐ)"]),
                hide_index=True,
                use_container_width=True,
                num_rows="dynamic",
                column_config={
                    "Mã Zalo (Sub_id2)": st.column_config.TextColumn("Mã Zalo (Sub_id2)", required=True, width="medium"),
                    "Tên Khách Hàng": st.column_config.TextColumn("Tên Khách Hàng Chuẩn", required=True, width="medium"),
                    "Số đơn hàng": st.column_config.NumberColumn("Số đơn hiện có", disabled=True, width="small"),
                    "Tổng tiền hoa hồng (VNĐ)": st.column_config.NumberColumn("Tổng tiền tích lũy", format="%,d ₫", disabled=True, width="medium")
                },
                key="mapping_data_editor"
            )
            
            col_save_map, _ = st.columns([3, 7])
            with col_save_map:
                save_mapping_btn = st.button("💾 LƯU BẢNG MAPPING & ĐỒNG BỘ", use_container_width=True)
                
            if save_mapping_btn:
                # Kiểm tra các thay đổi tên để đồng bộ sang Shopee_Cashback
                df_edited_clean = edited_mapping_table[["Mã Zalo (Sub_id2)", "Tên Khách Hàng"]].copy()
                
                # So sánh tên cũ và mới
                old_map_dict = dict(zip(df_mapping["Mã Zalo (Sub_id2)"].astype(str).str.strip(), df_mapping["Tên Khách Hàng"].astype(str).str.strip()))
                synced_total_orders = 0
                
                with st.spinner("Đang lưu bảng Mapping và quét đồng bộ đơn hàng..."):
                    for _, r in df_edited_clean.iterrows():
                        z = str(r["Mã Zalo (Sub_id2)"]).strip()
                        new_n = str(r["Tên Khách Hàng"]).strip()
                        if z in old_map_dict:
                            old_n = old_map_dict[z]
                            if old_n and new_n and (old_n != new_n):
                                _, n_synced, _ = sync_customer_name(old_n, new_n, target_sub2=z, sheet_target=st.session_state["sheet_name"])
                                synced_total_orders += n_synced
                                
                    save_mapping_data(df_edited_clean, sheet_target=st.session_state["sheet_name"])
                    
                st.toast("Đã lưu bảng Mapping và đồng bộ thành công!", icon="💾")
                st.success(f"✅ Bảng `Mapping_Data` đã được cập nhật thành công! (Đã quét đồng bộ {synced_total_orders} đơn hàng trong `Shopee_Cashback`).")
                st.rerun()

            st.divider()
            # PHẦN C: ĐỒNG BỘ KHI SỬA TRỰC TIẾP TRÊN GOOGLE SHEETS
            st.markdown("###### ⚡ Đồng Bộ Khi Chỉnh Sửa Trên Google Sheets:")
            st.caption("💡 *Nếu bạn vừa gõ sửa tên trực tiếp trên website Google Sheets (ở sheet `Mapping_Data`), bấm nút dưới đây để hệ thống tự động quét và cập nhật toàn bộ đơn hàng trong `Shopee_Cashback` theo tên chuẩn mới.*")
            col_auto_sync, _ = st.columns([4, 6])
            with col_auto_sync:
                auto_sync_btn = st.button("⚡ QUÉT & ĐỒNG BỘ TÊN TỪ GOOGLE SHEETS SANG ĐƠN HÀNG", use_container_width=True)
            if auto_sync_btn:
                with st.spinner("Đang đối soát và đồng bộ đơn hàng theo bảng Mapping..."):
                    cnt_updated, list_changes = auto_sync_names_from_mapping(sheet_target=st.session_state["sheet_name"])
                if cnt_updated > 0:
                    st.toast(f"Đã cập nhật {cnt_updated} đơn hàng!", icon="✨")
                    st.success(f"🎉 Đã quét và cập nhật thành công **{cnt_updated}** đơn hàng theo tên chuẩn trong `Mapping_Data`: {', '.join(list_changes)}")
                    st.rerun()
                else:
                    st.info("✅ Toàn bộ đơn hàng trong `Shopee_Cashback` hiện tại đã khớp 100% với tên chuẩn trong `Mapping_Data`!")

if __name__ == "__main__":
    main()
