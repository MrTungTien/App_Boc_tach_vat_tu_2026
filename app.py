import io
import math
import unicodedata

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Bóc tách & gộp vật tư cơ khí", layout="wide")
st.title("HỆ THỐNG TỰ ĐỘNG BÓC TÁCH & GỘP VẬT TƯ CƠ KHÍ")

CAY_MM = 6000  # chiều dài 1 cây thép hình (mm)


# ---------- Hàm hỗ trợ ----------
def bo_dau(s: str) -> str:
    s = unicodedata.normalize("NFD", str(s))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.replace("đ", "d").replace("Đ", "D").lower().strip()


def chuan_hoa_text(x) -> str:
    """In hoa, xóa khoảng trắng thừa."""
    if pd.isna(x):
        return ""
    return " ".join(str(x).split()).upper()


def doc_file(f) -> pd.DataFrame:
    if f.name.lower().endswith(".csv"):
        for enc in ("utf-8-sig", "utf-8", "cp1258", "latin-1"):
            try:
                f.seek(0)
                return pd.read_csv(f, encoding=enc)
            except UnicodeDecodeError:
                continue
        raise ValueError("Không đọc được mã hóa của file CSV.")
    return pd.read_excel(f, engine="openpyxl")


def doan_cot(cols, tu_khoa, mac_dinh=0):
    for i, c in enumerate(cols):
        k = bo_dau(c)
        if any(t in k for t in tu_khoa):
            return i
    return mac_dinh


def xu_ly(df, c_nd, c_dvt, c_sl, c_qd):
    d = pd.DataFrame(
        {
            "Nội dung": df[c_nd].map(chuan_hoa_text),
            "ĐVT": df[c_dvt].map(chuan_hoa_text),
            "Số lượng": pd.to_numeric(df[c_sl], errors="coerce").fillna(0),
            "Quy đổi (mm)": pd.to_numeric(df[c_qd], errors="coerce") if c_qd else float("nan"),
        }
    )

    # Quét từ trên xuống: dòng ĐVT = CỤM đặt hệ số nhân cho các dòng con bên dưới
    he_so, cum_hien_tai, hs_list, cum_list = 1.0, "", [], []
    for _, r in d.iterrows():
        if r["ĐVT"] == "CỤM":
            he_so, cum_hien_tai = r["Số lượng"], r["Nội dung"]
        hs_list.append(he_so)
        cum_list.append(cum_hien_tai)
    d["Thuộc cụm"] = cum_list
    d["Hệ số nhân của Cụm"] = hs_list
    d["Tổng nhu cầu thực tế"] = d["Số lượng"] * d["Hệ số nhân của Cụm"]

    is_cum = d["ĐVT"] == "CỤM"
    ct = d[~is_cum & (d["Nội dung"] != "")].copy()  # chi tiết con (bỏ dòng Cụm & dòng trống)

    # Thép hình: ĐVT = thanh
    is_thanh = ct["ĐVT"] == "THANH"
    hinh = ct[is_thanh]
    hinh = (
        hinh.groupby(["Nội dung", "Quy đổi (mm)"], dropna=False, as_index=False)["Tổng nhu cầu thực tế"]
        .sum()
        .rename(columns={"Tổng nhu cầu thực tế": "Tổng số thanh"})
    )
    hinh["Tổng chiều dài (mm)"] = hinh["Quy đổi (mm)"].fillna(0) * hinh["Tổng số thanh"]
    hinh["Số cây 6m"] = (hinh["Tổng chiều dài (mm)"] / CAY_MM).round(3)

    # Tổng hợp theo mác thép: cộng chiều dài các đoạn rồi làm tròn lên số cây
    tonghop = hinh.groupby("Nội dung", as_index=False)["Tổng chiều dài (mm)"].sum()
    tonghop["Số cây 6m"] = (tonghop["Tổng chiều dài (mm)"] / CAY_MM).round(3)
    tonghop["Số cây cần mua (làm tròn lên)"] = tonghop["Số cây 6m"].map(lambda v: math.ceil(v - 1e-9))

    # Thép tấm / vật tư khác
    khac = (
        ct[~is_thanh]
        .groupby(["Nội dung", "ĐVT"], as_index=False)["Tổng nhu cầu thực tế"]
        .sum()
        .rename(columns={"Tổng nhu cầu thực tế": "Tổng số lượng"})
    )
    return d, hinh, tonghop, khac


def xuat_excel(chi_tiet, hinh, tonghop, khac) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        hinh.to_excel(w, sheet_name="Thep_hinh", index=False)
        tonghop.to_excel(w, sheet_name="Thep_hinh_tong_hop", index=False)
        khac.to_excel(w, sheet_name="Thep_tam_vat_tu_khac", index=False)
        chi_tiet.to_excel(w, sheet_name="Chi_tiet_sau_nhan_he_so", index=False)
        for ws in w.book.worksheets:
            for col in ws.columns:
                n = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(12, n + 2), 60)
    return buf.getvalue()


# ---------- Giao diện ----------
file = st.file_uploader("Kéo thả file Excel (.xlsx) hoặc CSV dữ liệu thô", type=["xlsx", "csv"])

if file:
    try:
        raw = doc_file(file)
    except Exception as e:
        st.error(f"Không đọc được file: {e}")
        st.stop()

    raw = raw.dropna(how="all").reset_index(drop=True)
    st.subheader("Dữ liệu thô")
    st.dataframe(raw, use_container_width=True)

    cols = list(raw.columns)
    st.subheader("Chọn cột dữ liệu")
    a, b, c, d_ = st.columns(4)
    c_nd = a.selectbox("Cột Nội dung", cols, index=doan_cot(cols, ["noi dung", "ten", "vat tu"], 0))
    c_dvt = b.selectbox("Cột ĐVT", cols, index=doan_cot(cols, ["dvt", "don vi"], min(1, len(cols) - 1)))
    c_sl = c.selectbox("Cột Số lượng", cols, index=doan_cot(cols, ["so luong", "sl"], min(2, len(cols) - 1)))
    opts = ["(không có)"] + cols
    i_qd = doan_cot(cols, ["quy doi"], -1) + 1
    c_qd = d_.selectbox("Cột Số lượng quy đổi (chiều dài mm)", opts, index=i_qd)
    c_qd = None if c_qd == "(không có)" else c_qd

    if st.button("Xử lý & Gom nhóm vật tư", type="primary"):
        st.session_state["kq"] = xu_ly(raw, c_nd, c_dvt, c_sl, c_qd)

    if "kq" in st.session_state:
        chi_tiet, hinh, tonghop, khac = st.session_state["kq"]
        st.success("Đã xử lý xong.")
        st.subheader("1. Thép hình (gộp theo Nội dung + chiều dài)")
        st.dataframe(hinh, use_container_width=True)
        st.subheader("   Tổng hợp thép hình theo Nội dung (số cây 6m)")
        st.dataframe(tonghop, use_container_width=True)
        st.subheader("2. Thép tấm / vật tư khác (gộp theo Nội dung)")
        st.dataframe(khac, use_container_width=True)
        with st.expander("Chi tiết từng dòng sau khi nhân hệ số Cụm"):
            st.dataframe(chi_tiet, use_container_width=True)

        st.download_button(
            "Tải về file Excel kết quả",
            data=xuat_excel(chi_tiet, hinh, tonghop, khac),
            file_name="ket_qua_gom_vat_tu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Hãy tải file dữ liệu lên để bắt đầu.")
