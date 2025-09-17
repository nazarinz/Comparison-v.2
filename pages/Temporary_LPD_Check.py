# pages/Temporary_LPD_Check.py — Temporary LPD Check
# -----------------------------------------------------------------------------
# Fitur: User upload "Temporary Tracking.xlsx". Jika ada SO yang sama (duplikat)
# dan kolom "Remark 2"-nya kosong, maka kolom "Result_LPD" untuk baris tsb
# diisi/ganti menjadi "TEMP".
#
# Langkah umum:
# 1) Upload file XLSX (sheet pertama)
# 2) Deteksi kolom penting secara fleksibel (case-insensitive): SO, Remark 2, Result_LPD
# 3) Tandai baris dengan SO duplikat & Remark 2 kosong → set Result_LPD = "TEMP"
# 4) Tampilkan ringkasan & pratinjau, sediakan unduhan Excel/CSV hasil
# -----------------------------------------------------------------------------

from __future__ import annotations
import pandas as pd
import streamlit as st

from utils_pgd import read_excel_file, _export_excel_styled

# --------------------------------- Page Setup --------------------------------
st.set_page_config(page_title="🕒 Temporary LPD Check", layout="wide")
st.title("🕒 Temporary LPD Check")

st.caption(
    "Upload **Temporary Tracking.xlsx**. Aturan: *Jika ada SO yang sama (duplikat) "
    "dan `Remark 2` kosong, maka kolom `Result_LPD` baris tsb diisi `TEMP`*."
)

# --------------------------------- Sidebar -----------------------------------
with st.sidebar:
    st.header("📤 Upload File")
    xlsx = st.file_uploader("Temporary Tracking (.xlsx)", type=["xlsx"], key="temp_lpd_file")
    st.markdown(
        """
**Header contoh minimal:**
- `SO`
- `Remark 2`
- `Result_LPD` *(opsional, bila tidak ada akan dibuat)*

Kolom lain (Requester, Order date, PO, Article, LPD original, dll) akan dipertahankan.
        """
    )

# ------------------------------- Helpers -------------------------------------

def _normalize_columns(df: pd.DataFrame) -> dict:
    """Buat peta {key: actual_colname} dengan pencocokan case-insensitive/longgar."""
    # siapkan versi normal dari nama kolom untuk pencarian fleksibel
    def norm(s: str) -> str:
        return (
            str(s)
            .strip()
            .lower()
            .replace(".", " ")
            .replace("_", " ")
            .replace("-", " ")
        )

    wanted = {
        "so": ["so"],
        "remark2": ["remark 2", "remark2", "remark-2", "keterangan 2"],
        "result_lpd": ["result lpd", "result_lpd", "result-lpd"],
        "lpd_original": ["lpd original", "lpd", "original lpd"],  # fallback bila result_lpd tidak ada
    }

    norm_map = {col: norm(col) for col in df.columns}
    found = {}
    for key, aliases in wanted.items():
        actual = None
        for col, normed in norm_map.items():
            for alias in aliases:
                if normed == norm(alias):
                    actual = col
                    break
            if actual:
                break
        found[key] = actual
    return found


def _is_empty_series(s: pd.Series) -> pd.Series:
    if s is None:
        return pd.Series([False] * 0)
    return s.isna() | s.astype(str).str.strip().eq("")


# --------------------------------- Main --------------------------------------
if xlsx:
    try:
        df = read_excel_file(xlsx)  # sheet pertama
        st.success(f"File dibaca: {df.shape[0]} baris, {df.shape[1]} kolom")

        # deteksi kolom penting
        cols = _normalize_columns(df)
        so_col = cols.get("so")
        remark2_col = cols.get("remark2")
        result_lpd_col = cols.get("result_lpd")
        lpd_orig_col = cols.get("lpd_original")

        # validasi kolom minimal
        missing_critical = [n for n, c in {"SO": so_col, "Remark 2": remark2_col}.items() if c is None]
        if missing_critical:
            st.error(f"Kolom wajib tidak ditemukan: {', '.join(missing_critical)}. Cek header file kamu.")
            st.stop()

        # siapkan kolom Result_LPD (buat bila tidak ada)
        if result_lpd_col is None:
            result_lpd_col = "Result_LPD"
            if lpd_orig_col and lpd_orig_col in df.columns:
                df[result_lpd_col] = df[lpd_orig_col]
            else:
                df[result_lpd_col] = ""  # isi kosong dahulu

        # snapshot sebelum perubahan untuk perbandingan
        before = df[result_lpd_col].copy()

        # mask SO duplikat
        so_series = df[so_col].astype(str).str.strip()
        mask_dup_so = so_series.duplicated(keep=False)

        # mask remark2 kosong
        remark2_series = df[remark2_col]
        mask_remark_empty = _is_empty_series(remark2_series)

        # baris yang perlu di-TEMP
        mask_temp = mask_dup_so & mask_remark_empty

        # update nilai Result_LPD
        df.loc[mask_temp, result_lpd_col] = "TEMP"

        # ringkasan
        total_rows = len(df)
        affected = int(mask_temp.sum())
        dup_groups = int(so_series[mask_dup_so].nunique())

        st.divider()
        st.subheader("📊 Ringkasan Perubahan")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total baris", total_rows)
        c2.metric("Baris diubah → TEMP", affected)
        c3.metric("SO duplikat (grup)", dup_groups)

        # opsi tampil
        st.subheader("🔎 Pratinjau Hasil")
        show_only_changed = st.checkbox("Tampilkan hanya baris yang berubah (TEMP)", value=True)
        show_cols = st.multiselect(
            "Kolom yang ditampilkan",
            options=list(df.columns),
            default=[col for col in ["Requester", "Order date", "PO", so_col, "Article No.", result_lpd_col, remark2_col] if col in df.columns],
        )

        view_df = df.copy()
        if show_only_changed:
            view_df = view_df[mask_temp]
        if show_cols:
            view_df = view_df[show_cols]

        # tampilkan perubahan (before/after) jika kolom ada di pratinjau
        if result_lpd_col in view_df.columns:
            view_df = view_df.copy()
            view_df["Result_LPD_before"] = before.loc[view_df.index]

        st.dataframe(view_df.head(2000), use_container_width=True)

        # unduhan
        st.subheader("⬇️ Unduh Hasil")
        out_xlsx_name = "Temporary_LPD_Checked.xlsx"
        out_csv_name = "Temporary_LPD_Checked.csv"

        xbio = _export_excel_styled(df, sheet_name="Temporary LPD")
        st.download_button(
            "Download Excel (styled)", data=xbio, file_name=out_xlsx_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True
        )
        st.download_button(
            "Download CSV", data=df.to_csv(index=False).encode("utf-8"), file_name=out_csv_name,
            mime="text/csv", use_container_width=True
        )

    except Exception as e:
        st.error("Terjadi error saat memproses file.")
        st.exception(e)
else:
    st.info("Unggah file **Temporary Tracking.xlsx** di sidebar untuk mulai.")
