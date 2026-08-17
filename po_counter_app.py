import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Đếm PO theo Mã khách hàng", layout="wide")
st.title("📊 Đếm PO theo Mã khách hàng")

st.markdown(
    "Upload file Excel export (có các cột: Ngày, Mã KH, Tên KH, Mã số thuế...). "
    "App sẽ đếm số PO theo tổ hợp **Mã khách hàng + Ngày/Tháng**. "
    "Nếu 1 Mã KH trong cùng ngày xuất nhiều hóa đơn, vẫn tính là **1 PO**."
)

uploaded_files = st.file_uploader(
    "Chọn file Excel (có thể chọn nhiều file)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

if not uploaded_files:
    st.info("👆 Upload ít nhất 1 file Excel để bắt đầu.")
    st.stop()

# ---------- Đọc và gộp dữ liệu ----------
dfs = []
for f in uploaded_files:
    try:
        df = pd.read_excel(f)
        df["__nguon_file__"] = f.name
        dfs.append(df)
    except Exception as e:
        st.error(f"Lỗi đọc file {f.name}: {e}")

if not dfs:
    st.stop()

data = pd.concat(dfs, ignore_index=True)
# pandas tự đổi tên cột trùng thành Mã, Mã.1, Mã.2 ...
all_cols = list(data.columns)

st.success(f"Đã đọc {len(data)} dòng từ {len(uploaded_files)} file.")
with st.expander("📄 Xem trước dữ liệu"):
    st.dataframe(data.head(50), use_container_width=True)

# ---------- Chọn cột ----------
st.subheader("⚙️ Chọn cột dữ liệu")

col1, col2, col3 = st.columns(3)

def guess_col(keywords, cols):
    for c in cols:
        for kw in keywords:
            if kw.lower() in str(c).lower():
                return c
    return cols[0]

with col1:
    col_makh = st.selectbox(
        "Cột Mã khách hàng",
        all_cols,
        index=all_cols.index(guess_col(["mã kh", "makh", "ma kh"], all_cols)),
    )
with col2:
    col_ngay = st.selectbox(
        "Cột Ngày",
        all_cols,
        index=all_cols.index(guess_col(["ngày", "ngay"], all_cols)),
    )
with col3:
    col_mst = st.selectbox(
        "Cột Mã số thuế",
        all_cols,
        index=all_cols.index(guess_col(["mã số thuế", "mst"], all_cols)),
    )

# ---------- Xử lý ----------
work = data.copy()
work["_MAKH_"] = work[col_makh].astype(str).str.strip()
work["_MST_"] = work[col_mst].astype(str).str.strip()

# Parse ngày - hỗ trợ cả 2 dạng: text ("01/08/26") và serial number Excel (46235)
raw_ngay = work[col_ngay]
if pd.api.types.is_numeric_dtype(raw_ngay):
    # Ngày lưu dạng số serial của Excel (VD: 46235)
    work["_NGAY_"] = pd.to_datetime(raw_ngay, unit="D", origin="1899-12-30", errors="coerce")
else:
    work["_NGAY_"] = pd.to_datetime(raw_ngay, dayfirst=True, errors="coerce")
    # Nếu vẫn lỗi nhiều, thử lại coi có phải chuỗi số serial không
    if work["_NGAY_"].isna().mean() > 0.5:
        as_num = pd.to_numeric(raw_ngay, errors="coerce")
        work["_NGAY_"] = pd.to_datetime(as_num, unit="D", origin="1899-12-30", errors="coerce")

work["_THANG_"] = work["_NGAY_"].dt.to_period("M").astype(str)
work["_NGAY_STR_"] = work["_NGAY_"].dt.strftime("%d/%m/%Y")

n_loi_ngay = work["_NGAY_"].isna().sum()
if n_loi_ngay > 0:
    st.warning(f"⚠️ Có {n_loi_ngay} dòng không đọc được ngày (sẽ bị bỏ qua khi group theo ngày/tháng).")

st.subheader("🔎 Cách gộp nhóm để đếm PO")
group_mode = st.radio(
    "Đếm PO theo tổ hợp Mã KH +:",
    ["Ngày (từng ngày riêng)", "Tháng (gộp cả tháng)"],
    horizontal=True,
)

group_cols = ["_MAKH_"]
if group_mode.startswith("Ngày"):
    group_cols.append("_NGAY_STR_")
    time_label = "Ngày"
else:
    group_cols.append("_THANG_")
    time_label = "Tháng"

valid = work.dropna(subset=["_NGAY_"]) if True else work

# Mỗi nhóm (Mã KH + Ngày/Tháng) = 1 PO, dù có bao nhiêu dòng hóa đơn bên trong
summary = (
    valid.groupby(group_cols)
    .agg(So_hoa_don=("_MAKH_", "size"), MST=("_MST_", "first"))
    .reset_index()
    .rename(columns={"_MAKH_": "Ma_KH", group_cols[-1]: time_label})
)
summary["So_luong_PO"] = 1  # mỗi nhóm luôn tính là 1 PO
# Sắp cột: Ma_KH, MST, thời gian, số hóa đơn, số PO
cols_order = ["Ma_KH", "MST", time_label, "So_hoa_don", "So_luong_PO"]
summary = summary[[c for c in cols_order if c in summary.columns]]
summary = summary.sort_values("So_hoa_don", ascending=False).reset_index(drop=True)

# Tổng theo Mã KH: đếm SỐ NHÓM (= số PO thật sự), không cộng dồn số dòng hóa đơn
tong_theo_makh = (
    summary.groupby(["Ma_KH", "MST"])
    .agg(
        Tong_So_PO=("So_luong_PO", "sum"),      # số PO thật (đã loại trùng hóa đơn)
        Tong_So_Hoa_Don=("So_hoa_don", "sum"),  # tổng số dòng hóa đơn gốc
    )
    .reset_index()
    .sort_values("Tong_So_PO", ascending=False)
    .reset_index(drop=True)
)

st.caption(
    f"ℹ️ Quy tắc: nếu cùng 1 Mã KH trong cùng {time_label.lower()} xuất nhiều hóa đơn "
    "thì vẫn chỉ tính là **1 PO**."
)

st.subheader(f"📋 Danh sách PO (mỗi dòng = 1 PO, theo Mã KH + {time_label})")
st.dataframe(summary, use_container_width=True)

st.subheader("📈 Tổng số PO theo từng Mã KH")
st.dataframe(tong_theo_makh, use_container_width=True)

# ---------- Xuất Excel ----------
def to_excel_bytes(summary_df, tong_df, raw_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Chi tiet dem PO", index=False)
        tong_df.to_excel(writer, sheet_name="Tong theo Ma KH", index=False)
        raw_df.drop(columns=["__nguon_file__"], errors="ignore").to_excel(
            writer, sheet_name="Du lieu goc", index=False
        )
    return output.getvalue()

excel_bytes = to_excel_bytes(summary, tong_theo_makh, data)

st.download_button(
    label="⬇️ Tải kết quả Excel",
    data=excel_bytes,
    file_name=f"dem_PO_theo_MaKH_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
