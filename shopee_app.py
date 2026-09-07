# -*- coding: utf-8 -*-
"""
Hệ thống Quản lý & Tra cứu Hoàn tiền Hoa hồng Shopee
Giao diện UI/UX Hiện đại - Tone màu Cam Shopee (#ee4d2d)
Hỗ trợ lưu trữ trực tiếp vào Google Sheets theo thời gian thực (Real-time)
Cơ chế tự động dự phòng (Fallback) file CSV cục bộ
"""

import streamlit as st
import pandas as pd
import datetime
import os
import json

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
CREDENTIALS_FILE = "credentials.json"
DEFAULT_SHEET_NAME = "Shopee_Cashback"
DEFAULT_ADMIN_PASS = "Kunhan1996@"
STATUS_OPTIONS = ["Chờ Shopee duyệt", "Đã bank", "Không hợp lệ"]
REQUIRED_COLS = ["Tên khách", "Ngày đặt", "Mã đơn", "Tổng HH thực tế", "Trạng thái"]


# =============================================================================
# 2. XỬ LÝ KẾT NỐI GOOGLE SHEETS & CSV
# =============================================================================
def get_service_account_email():
    """Đọc email của Service Account từ file credentials.json để tiện cấp quyền chia sẻ."""
    if os.path.exists(CREDENTIALS_FILE):
        try:
            with open(CREDENTIALS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("client_email", "")
        except Exception:
            return ""
    return ""

def get_gsheet_worksheet(sheet_identifier=DEFAULT_SHEET_NAME):
    """
    Kết nối tới Google Sheets qua Service Account.
    Hỗ trợ mở bằng Tên Sheet, Spreadsheet ID hoặc Link URL.
    """
    if not GSPREAD_AVAILABLE:
        return None, "Thư viện `gspread` chưa được cài đặt."
    
    if not os.path.exists(CREDENTIALS_FILE):
        return None, f"Chưa tìm thấy file `{CREDENTIALS_FILE}` trong thư mục ứng dụng."
    
    try:
        gc = gspread.service_account(filename=CREDENTIALS_FILE)
        
        # Kiểm tra nếu là URL
        if sheet_identifier.startswith("https://"):
            spreadsheet = gc.open_by_url(sheet_identifier)
        else:
            try:
                spreadsheet = gc.open(sheet_identifier)
            except Exception:
                spreadsheet = gc.open_by_key(sheet_identifier)
                
        worksheet = spreadsheet.sheet1
        return worksheet, None
    except Exception as e:
        return None, str(e)

def normalize_df(df):
    """Chuẩn hóa cấu trúc DataFrame đồng nhất."""
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
    """Đọc dữ liệu từ file CSV cục bộ."""
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
    """Lưu backup ra file CSV cục bộ."""
    try:
        df_save = df.copy()
        df_save["Mã đơn"] = df_save["Mã đơn"].astype(str)
        df_save.to_csv(CSV_FILE, index=False, encoding="utf-8-sig")
    except Exception:
        pass

def load_data(sheet_target=DEFAULT_SHEET_NAME):
    """
    Đọc dữ liệu: Ưu tiên đọc trực tiếp từ Google Sheets.
    Nếu chưa cấu hình hoặc mất mạng -> Fallback về file CSV cục bộ.
    """
    ws, err = get_gsheet_worksheet(sheet_target)
    if ws is not None:
        try:
            # get_all_values() trả về mảng chuỗi nguyên bản, bảo toàn số 0 đầu của Mã đơn
            records = ws.get_all_values()
            if len(records) > 0:
                headers = records[0]
                rows = records[1:] if len(records) > 1 else []
                df = pd.DataFrame(rows, columns=headers)
                df = normalize_df(df)
            else:
                # Nếu trang tính trống, tự động ghi dòng tiêu đề
                ws.append_row(REQUIRED_COLS)
                df = pd.DataFrame(columns=REQUIRED_COLS)
                
            # Lưu backup cục bộ
            save_local_csv(df)
            return df, "gsheets", None
        except Exception as e:
            fallback_df = load_local_csv()
            return fallback_df, "csv_fallback", f"Lỗi đọc Google Sheets: {e}"
            
    # Chưa kết nối Google Sheets -> Dùng CSV cục bộ
    fallback_df = load_local_csv()
    return fallback_df, "csv", err

def add_order_to_storage(new_row_dict, sheet_target=DEFAULT_SHEET_NAME):
    """
    Thêm 1 đơn hàng mới:
    Nếu có Google Sheets -> gọi append_row cực nhanh.
    Đồng thời cập nhật file CSV cục bộ.
    """
    cust_name = new_row_dict.get("Tên khách", "")
    order_date = new_row_dict.get("Ngày đặt", "")
    order_code = str(new_row_dict.get("Mã đơn", "")).strip()
    commission = int(new_row_dict.get("Tổng HH thực tế", 0))
    status = new_row_dict.get("Trạng thái", "Chờ Shopee duyệt")
    
    ws, err = get_gsheet_worksheet(sheet_target)
    saved_to_gsheet = False
    
    if ws is not None:
        try:
            # Nếu mã đơn có số 0 ở đầu, thêm dấu ' để Google Sheets hiểu là dạng Text
            code_formatted = f"'{order_code}" if order_code.startswith("0") else order_code
            row_data = [cust_name, order_date, code_formatted, commission, status]
            ws.append_row(row_data, value_input_option="USER_ENTERED")
            saved_to_gsheet = True
        except Exception as e:
            st.error(f"Lỗi ghi vào Google Sheets: {e}")
            
    # Luôn đồng bộ vào file CSV cục bộ
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

def save_all_to_storage(df, sheet_target=DEFAULT_SHEET_NAME):
    """
    Ghi đè toàn bộ dữ liệu (dùng cho Tab 2 Data Editor):
    Cập nhật đồng thời lên Google Sheets và file CSV.
    """
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

def format_vnd(amount):
    """Định dạng tiền tệ VNĐ chuẩn có dấu chấm phân cách hàng nghìn."""
    return f"{amount:,.0f} ₫".replace(",", ".")


# =============================================================================
# 3. HÀM GIAO DIỆN CHÍNH (STREAMLIT)
# =============================================================================
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
        
        /* Container tiêu đề căn giữa */
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
        
        /* Thẻ thống kê Metric Cards bo tròn */
        .metrics-container {
            display: flex;
            gap: 20px;
            margin: 20px 0 28px;
            justify-content: center;
        }
        
        .metric-card {
            flex: 1;
            background: #ffffff;
            border-radius: 16px;
            padding: 22px 24px;
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
        
        .metric-title {
            font-size: 14px;
            font-weight: 600;
            color: var(--text-muted);
            margin-bottom: 8px;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .metric-amount {
            font-size: 28px;
            font-weight: 800;
            line-height: 1.2;
        }
        
        .metric-card.pending .metric-amount {
            color: #d97706;
        }
        
        .metric-card.success .metric-amount {
            color: #059669;
        }
        
        .metric-count {
            font-size: 13px;
            color: #94a3b8;
            margin-top: 6px;
            font-weight: 500;
        }

        /* Nút chính Shopee */
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
        
        /* Ô input tìm kiếm */
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

        /* Hướng dẫn tra cứu */
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

        /* Sidebar Badge */
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
                    st.success("✅ Đã tìm thấy `credentials.json`")
                    st.text_area("Email Service Account (chia sẻ quyền Editor cho email này):", value=sa_email, height=70)
                else:
                    st.warning("⚠️ Chưa có file `credentials.json` trong thư mục `D:\\CheckShopee`.")
                    st.caption("Ứng dụng đang tự động chạy ở chế độ **CSV cục bộ**.")

            if st.button("🚪 Đăng xuất", use_container_width=True):
                st.session_state["admin_logged_in"] = False
                st.toast("Đã đăng xuất khỏi quyền Admin!", icon="👋")
                st.rerun()
                
            st.divider()

        # Hiển thị trạng thái nguồn dữ liệu
        df_data, storage_mode, storage_err = load_data(st.session_state["sheet_name"])
        
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
                    pending_df = matched_df[matched_df["Trạng thái"] == "Chờ Shopee duyệt"]
                    banked_df = matched_df[matched_df["Trạng thái"] == "Đã bank"]
                    
                    total_pending = pending_df["Tổng HH thực tế"].sum()
                    count_pending = len(pending_df)
                    
                    total_banked = banked_df["Tổng HH thực tế"].sum()
                    count_banked = len(banked_df)
                    
                    st.markdown(f"""
                    <div class="metrics-container">
                        <div class="metric-card pending">
                            <div class="metric-title">⏳ Tổng tiền chờ duyệt</div>
                            <div class="metric-amount">{format_vnd(total_pending)}</div>
                            <div class="metric-count">{count_pending} đơn hàng đang đợi đối soát</div>
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
        # B. GIAO DIỆN ADMIN (QUẢN TRỊ VIÊN) - 2 TABS GIỮ NGUYÊN 100%
        # =====================================================================
        st.markdown("""
        <div style="margin-bottom: 24px;">
            <h2 style="color: #ee4d2d; font-weight: 800; margin: 0 0 6px;">👑 Bảng Quản Trị Hệ Thống Shopee Cashback</h2>
            <p style="color: #64748b; margin: 0; font-size: 14px;">Quản lý và cập nhật đơn hàng hoàn tiền hoa hồng. Dữ liệu đồng bộ trực tiếp với Google Sheets / CSV.</p>
        </div>
        """, unsafe_allow_html=True)
        
        tab1, tab2 = st.tabs([
            "⚡ Phần 1: Form Nhập Đơn Mới (Smart Input)",
            "📊 Phần 2: Cập Nhật Dữ Liệu (Quản lý hàng loạt)"
        ])
        
        # ---------------------------------------------------------------------
        # TAB 1: FORM NHẬP ĐƠN MỚI
        # ---------------------------------------------------------------------
        with tab1:
            st.markdown("##### 📝 Thêm Đơn Hàng Mới Tốc Độ Cao")
            st.caption("💡 *Sử dụng Form cho phép bạn dùng phím **Tab** để chuyển nhanh giữa các ô và nhấn **Enter** để thêm đơn tức thì.*")
            
            existing_customers = sorted([c for c in df_data["Tên khách"].dropna().unique().tolist() if str(c).strip() != ""])
            
            customer_mode = st.radio(
                "Loại khách hàng:",
                ["Khách hàng cũ", "Khách hàng mới"],
                horizontal=True,
                key="customer_mode_radio"
            )
            
            with st.form(key="form_add_new_order", clear_on_submit=True):
                f_col1, f_col2 = st.columns(2)
                
                with f_col1:
                    if customer_mode == "Khách hàng cũ":
                        if existing_customers:
                            customer_name = st.selectbox(
                                "Tên khách hàng cũ (gõ tìm kiếm nhanh):",
                                options=existing_customers,
                                help="Danh sách khách hàng không trùng lặp đã có sẵn"
                            )
                        else:
                            st.info("Chưa có khách hàng nào trong hệ thống. Vui lòng chuyển sang 'Khách hàng mới'.")
                            customer_name = st.text_input("Tên khách hàng:", placeholder="Nhập tên khách...")
                    else:
                        customer_name = st.text_input(
                            "Tên khách hàng mới:",
                            placeholder="Ví dụ: Nguyễn Văn An, Trần Thị Bích...",
                            help="Nhập họ và tên khách hàng mới"
                        )
                    
                    order_code = st.text_input(
                        "Mã đơn (Bắt buộc):",
                        placeholder="Ví dụ: 081928472910... (giữ nguyên số 0 ở đầu)",
                        help="Mã đơn hàng Shopee. Luôn được lưu dạng Text để giữ nguyên số 0 ở đầu."
                    )
                    
                    yesterday_date = datetime.date.today() - datetime.timedelta(days=1)
                    order_date = st.date_input(
                        "Ngày đặt:",
                        value=yesterday_date,
                        format="DD/MM/YYYY",
                        help="Mặc định tự động lấy trước ngày hiện tại 1 ngày (định dạng dd/mm/yyyy)"
                    )
                    
                with f_col2:
                    commission_val = st.number_input(
                        "Tổng HH thực tế (VNĐ):",
                        min_value=0,
                        step=1000,
                        value=0,
                        format="%d",
                        help="Số tiền hoa hồng thực tế nhận được (VNĐ)"
                    )
                    
                    order_status = st.selectbox(
                        "Trạng thái:",
                        options=STATUS_OPTIONS,
                        index=0,
                        help="Mặc định là 'Chờ Shopee duyệt'"
                    )
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    submit_btn = st.form_submit_button("➕ THÊM ĐƠN HÀNG", use_container_width=True)
                
                if submit_btn:
                    clean_name = str(customer_name).strip() if customer_name else ""
                    clean_code = str(order_code).strip() if order_code else ""
                    
                    if not clean_name:
                        st.error("⚠️ Vui lòng cung cấp Tên khách hàng!")
                    elif not clean_code:
                        st.error("⚠️ Mã đơn là bắt buộc, không được để trống!")
                    else:
                        formatted_date = order_date.strftime("%d/%m/%Y")
                        
                        saved_to_gs = add_order_to_storage({
                            "Tên khách": clean_name,
                            "Ngày đặt": formatted_date,
                            "Mã đơn": clean_code,
                            "Tổng HH thực tế": int(commission_val),
                            "Trạng thái": order_status
                        }, sheet_target=st.session_state["sheet_name"])
                        
                        target_msg = "Google Sheets & CSV" if saved_to_gs else "CSV cục bộ"
                        st.toast(f"Đã thêm đơn '{clean_code}' vào {target_msg}!", icon="✅")
                        st.success(f"🎉 Đã thêm thành công đơn hàng **{clean_code}** ({format_vnd(commission_val)}) vào **{target_msg}**!")
                        st.rerun()

        # ---------------------------------------------------------------------
        # TAB 2: CẬP NHẬT DỮ LIỆU (QUẢN LÝ HÀNG LOẠT)
        # ---------------------------------------------------------------------
        with tab2:
            st.markdown("##### ✏️ Quản Lý & Chỉnh Sửa Dữ Liệu Hàng Loạt")
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

if __name__ == "__main__":
    main()
