# pages/Temporary_LPD_Check.py — Temporary LPD Check (merge ke PGD Report)
# -----------------------------------------------------------------------------
# Fitur: User upload dua file:
#   1) Temporary Tracking.xlsx  (berisi kolom minimal: SO, Remark 2)
#   2) PGD Comparison Tracking Report - <tanggal>.xlsx (sheet hasil: Report)
#
# Aturan:
# - Cari SO yang sama (duplikat) di Temporary Tracking DAN kolom "Remark 2"-nya kosong
# - Untuk semua baris pada PGD Report dengan SO tersebut → set kolom "Result_LPD" menjadi "TEMP"
#
# Output:
# - Ringkasan perubahan
# - Pratinjau baris yang berubah
# - Unduh Excel (styled) & CSV untuk PGD Report yang sudah diperbarui
# -----------------------------------------------------------------------------

from __future__ import annotations
import pandas as pd
import streamlit as st

from utils_pgd import read_excel_file, _export_excel_styled

# --------------------------------- Page Setup --------------------------------
st.set_page_config(page_title="🕒 Temporary LPD Check", layout="wide")
st.title("🕒 Temporary LPD Check — Merge ke PGD Report")

st.caption(
    "Upload **Temporary Tracking.xlsx** dan **PGD Comparison Tracking Report - <tanggal>.xlsx**. "
    "Jika ada SO duplikat di Temporary Tracking dan `Remark 2` kosong, maka semua baris di PGD Report "
    "dengan SO tersebut akan di-set `Result_LPD = ""TEMP""`."
)

# --------------------------------- Sidebar -----------------------------------
with st.sidebar:
    st.header("📤 Upload Files")
    temp_file = st.file_uploader("Temporary Tracking (.xlsx)", type=["xlsx"], key="temp_lpd_file")
    pgd_file = st.file_uploader("PGD Comparison Report (.xlsx)", type=["xlsx"], key="pgd_report_file")
    st.markdown(
        """
**Minimal header:**
- Temporary Tracking: `SO`, `Remark 2`
- PGD Report: `SO`, `Result_LPD` (bila tidak ada akan dibuat)
        """
    )

# ------------------------------- Helpers -------------------------------------

def _norm(s: str) -> str:
    return (
        str(s)
        .strip()
        .lower()
        .replace(".", " ")
        .replace("_", " ")
        .replace("-", " ")
    )


def _find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    m = {col: _norm(col) for col in df.columns}
    for want in candidates:
        wantn = _norm(want)
        for col, normed in m.items():
            if normed == wantn:
                return col
    return None


def _is_empty_series(s: pd.Series) -> pd.Series:
    if s is None:
        return pd.Series([False] * 0)
    return s.isna() | s.astype(str).str.strip().eq("")


# --------------------------------- Main --------------------------------------
if temp_file and pgd_file:
    try:
        # Baca kedua file (sheet pertama)
        temp_df = read_excel_file(temp_file)
        pgd_df = read_excel_file(pgd_file)

        st.success(f"Temporary Tracking dibaca: {temp_df.shape[0]} baris, {temp_df.shape[1]} kolom")
        st.success(f"PGD Report dibaca: {pgd_df.shape[0]} baris, {pgd_df.shape[1]} kolom")

        # Deteksi kolom penting
        temp_so_col = _find_col(temp_df, ["SO"]) or "SO"
        temp_remark2_col = _find_col(temp_df, ["Remark 2", "Remark2", "Remark-2"]) or "Remark 2"

        pgd_so_col = _find_col(pgd_df, ["SO"]) or "SO"
        pgd_result_lpd_col = _find_col(pgd_df, ["Result LPD", "Result_LPD", "Result-LPD"]) or "Result_LPD"

        # Validasi kolom wajib di Temporary
        miss_temp = [n for n, c in {"SO": temp_so_col, "Remark 2": temp_remark2_col}.items() if c not in temp_df.columns]
        if miss_temp:
            st.error("Temporary Tracking: kolom wajib tidak ditemukan: " + ", ".join(miss_temp))
            st.stop()
        # Validasi kolom SO di PGD
        if pgd_so_col not in pgd_df.columns:
            st.error("PGD Report: kolom 'SO' tidak ditemukan.")
            st.stop()
        # Pastikan kolom Result_LPD ada di PGD
        if pgd_result_lpd_col not in pgd_df.columns:
            pgd_df[pgd_result_lpd_col] = ""

        # Identifikasi SO duplikat di Temporary dengan Remark 2 kosong
        tmp_so = temp_df[temp_so_col].astype(str).str.strip()
        tmp_rem2_empty = _is_empty_series(temp_df[temp_remark2_col])
        dup_mask = tmp_so.duplicated(keep=False)
        so_target = set(tmp_so[dup_mask & tmp_rem2_empty].unique().tolist())

        # Terapkan ke PGD Report
        before = pgd_df[pgd_result_lpd_col].copy()
        match_mask = pgd_df[pgd_so_col].astype(str).str.strip().isin(so_target)
        pgd_df.loc[match_mask, pgd_result_lpd_col] = "TEMP"

        # Ringkasan
        total_pgd = len(pgd_df)
        affected_rows = int(match_mask.sum())
        affected_sos = len(so_target)

        st.divider()
        st.subheader("📊 Ringkasan Perubahan")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total baris PGD", total_pgd)
        c2.metric("Baris diubah → TEMP", affected_rows)
        c3.metric("SO terpengaruh (unik)", affected_sos)

        # Pratinjau
        st.subheader("🔎 Pratinjau Hasil (baris berubah saja)")
        show_cols = st.multiselect(
            "Kolom yang ditampilkan",
            options=list(pgd_df.columns),
            default=[col for col in ["PO No.(Full)", pgd_so_col, "LPD", "Infor LPD", pgd_result_lpd_col] if col in pgd_df.columns],
        )
        view_df = pgd_df[match_mask]
        if show_cols:
            view_df = view_df[show_cols]
        view_df = view_df.copy()
        if pgd_result_lpd_col in view_df.columns:
            view_df["Result_LPD_before"] = before.loc[view_df.index]
        st.dataframe(view_df.head(2000), use_container_width=True)

        # Unduhan hasil
        st.subheader("⬇️ Unduh PGD Report (hasil diperbarui)")
        out_xlsx_name = "PGD_Comparison_Updated_Temporary_LPD.xlsx"
        out_csv_name = "PGD_Comparison_Updated_Temporary_LPD.csv"

        xbio = _export_excel_styled(pgd_df, sheet_name="Report")
        st.download_button(
            "Download Excel (styled)", data=xbio, file_name=out_xlsx_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True
        )
        st.download_button(
            "Download CSV", data=pgd_df.to_csv(index=False).encode("utf-8"), file_name=out_csv_name,
            mime="text/csv", use_container_width=True
        )

        # Info tambahan
        with st.expander("ℹ️ Catatan Teknis"):
            st.write(
                "SO target diambil dari Temporary Tracking dengan kondisi: duplikat (ada lebih dari 1 baris dengan SO yang sama) "
                "dan `Remark 2` kosong. Semua baris pada PGD Report dengan SO tersebut ditandai `TEMP` di kolom Result_LPD."
            )
    except Exception as e:
        st.error("Terjadi error saat memproses file.")
        st.exception(e)
else:
    st.info("Unggah kedua file di sidebar untuk mulai (Temporary Tracking & PGD Report).")
