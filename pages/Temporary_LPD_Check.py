# pages/Temporary_LPD_Check_fixed.py — Temporary LPD Check (robust SO matching)
# -----------------------------------------------------------------------------
# Upload 2 file:
#   1) Temporary Tracking.xlsx  (kolom minimal: SO, Remark 2)
#   2) PGD Comparison Tracking Report - <tanggal>.xlsx  (kolom: SO, Result_LPD)
# Aturan:
#   - Cari SO duplikat di Temporary Tracking dengan Remark 2 kosong
#   - Semua baris di PGD Report yang memiliki SO tsb -> set Result_LPD = "TEMP"
# Perbaikan: normalisasi SO (hapus non-digit, buang leading zero) agar match stabil
# terhadap perbedaan format (angka -> "110...0", "110...0.0", notasi ilmiah, dll).
# -----------------------------------------------------------------------------

from __future__ import annotations
import pandas as pd
import streamlit as st

from utils_pgd import read_excel_file, _export_excel_styled

# --------------------------------- Page Setup --------------------------------
st.set_page_config(page_title="🕒 Temporary LPD Check", layout="wide")
st.title("🕒 Temporary LPD Check — Merge ke PGD Report (Fixed)")

st.caption(
    "Upload **Temporary Tracking.xlsx** dan **PGD Comparison Tracking Report - <tanggal>.xlsx**. "
    "SO distandarkan (hanya digit, tanpa leading zero) untuk menghindari mismatch."
)

# --------------------------------- Sidebar -----------------------------------
with st.sidebar:
    st.header("📤 Upload Files")
    temp_file = st.file_uploader("Temporary Tracking (.xlsx)", type=["xlsx"], key="temp_lpd_file_fixed")
    pgd_file = st.file_uploader("PGD Comparison Report (.xlsx)", type=["xlsx"], key="pgd_report_file_fixed")

# ------------------------------- Helpers -------------------------------------

def _normname(s: str) -> str:
    return (
        str(s)
        .strip()
        .lower()
        .replace(".", " ")
        .replace("_", " ")
        .replace("-", " ")
    )


def _find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    m = {col: _normname(col) for col in df.columns}
    for want in candidates:
        wantn = _normname(want)
        for col, normed in m.items():
            if normed == wantn:
                return col
    return None


def _is_empty_series(s: pd.Series) -> pd.Series:
    if s is None:
        return pd.Series([False] * 0)
    return s.isna() | s.astype(str).str.strip().eq("")


def _normalize_so_series(s: pd.Series) -> pd.Series:
    return (
        s.astype(str)
        .str.replace(r"\D+", "", regex=True)  # hanya digit
        .str.lstrip("0")                         # buang leading zero
    )

# --------------------------------- Main --------------------------------------
if temp_file and pgd_file:
    try:
        temp_df = read_excel_file(temp_file)
        pgd_df = read_excel_file(pgd_file)
        st.success(f"Temporary Tracking dibaca: {temp_df.shape[0]} baris, {temp_df.shape[1]} kolom")
        st.success(f"PGD Report dibaca: {pgd_df.shape[0]} baris, {pgd_df.shape[1]} kolom")

        # Kolom penting
        temp_so_col = _find_col(temp_df, ["SO"]) or "SO"
        temp_remark2_col = _find_col(temp_df, ["Remark 2", "Remark2", "Remark-2"]) or "Remark 2"
        pgd_so_col = _find_col(pgd_df, ["SO"]) or "SO"
        pgd_result_lpd_col = _find_col(pgd_df, ["Result LPD", "Result_LPD", "Result-LPD"]) or "Result_LPD"

        # Validasi
        miss_temp = [n for n, c in {"SO": temp_so_col, "Remark 2": temp_remark2_col}.items() if c not in temp_df.columns]
        if miss_temp:
            st.error("Temporary Tracking: kolom wajib tidak ditemukan: " + ", ".join(miss_temp))
            st.stop()
        if pgd_so_col not in pgd_df.columns:
            st.error("PGD Report: kolom 'SO' tidak ditemukan.")
            st.stop()
        if pgd_result_lpd_col not in pgd_df.columns:
            pgd_df[pgd_result_lpd_col] = ""

        # Normalisasi SO kedua file
        temp_df["__SO_norm__"] = _normalize_so_series(temp_df[temp_so_col])
        pgd_df["__SO_norm__"] = _normalize_so_series(pgd_df[pgd_so_col])

        # SO target: duplikat + remark2 kosong
        tmp_rem2_empty = _is_empty_series(temp_df[temp_remark2_col])
        dup_mask = temp_df["__SO_norm__"].duplicated(keep=False)
        so_target = set(temp_df.loc[dup_mask & tmp_rem2_empty, "__SO_norm__"].unique().tolist())

        # Terapkan ke PGD
        before = pgd_df[pgd_result_lpd_col].copy()
        match_mask = pgd_df["__SO_norm__"].isin(so_target)
        pgd_df.loc[match_mask, pgd_result_lpd_col] = "TEMP"

        # Ringkasan
        st.divider()
        st.subheader("📊 Ringkasan Perubahan")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total baris PGD", len(pgd_df))
        c2.metric("Baris diubah → TEMP", int(match_mask.sum()))
        c3.metric("SO terpengaruh (unik)", len(so_target))

        # Pratinjau
        st.subheader("🔎 Pratinjau Hasil (baris berubah saja)")
        show_cols = st.multiselect(
            "Kolom yang ditampilkan",
            options=[c for c in pgd_df.columns if not c.startswith("__")],
            default=[col for col in ["PO No.(Full)", pgd_so_col, "LPD", "Infor LPD", pgd_result_lpd_col] if col in pgd_df.columns],
        )
        view_df = pgd_df.loc[match_mask, [c for c in show_cols]]
        if pgd_result_lpd_col in view_df.columns:
            view_df = view_df.copy()
            view_df["Result_LPD_before"] = before.loc[view_df.index]
        st.dataframe(view_df.head(2000), use_container_width=True)

        # Unduhan
        st.subheader("⬇️ Unduh PGD Report (hasil diperbarui)")
        out_xlsx_name = "PGD_Comparison_Updated_Temporary_LPD.xlsx"
        out_csv_name = "PGD_Comparison_Updated_Temporary_LPD.csv"

        xbio = _export_excel_styled(pgd_df.drop(columns=[c for c in pgd_df.columns if c.startswith("__")]), sheet_name="Report")
        st.download_button(
            "Download Excel (styled)", data=xbio, file_name=out_xlsx_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True
        )
        st.download_button(
            "Download CSV", data=pgd_df.drop(columns=[c for c in pgd_df.columns if c.startswith("__")]).to_csv(index=False).encode("utf-8"),
            file_name=out_csv_name, mime="text/csv", use_container_width=True
        )

    except Exception as e:
        st.error("Terjadi error saat memproses file.")
        st.exception(e)
else:
    st.info("Unggah kedua file di sidebar untuk mulai (Temporary Tracking & PGD Report).")
