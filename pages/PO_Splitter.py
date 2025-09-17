# pages/PGD_Comparison.py — Rewritten
# -----------------------------------------------------------------------------
# Halaman utama untuk perbandingan SAP vs Infor.
# Fitur:
# - Upload 1 file SAP (.xlsx) dan multi-file Infor (.csv)
# - Merge, cleaning, compare kolom kunci (Result_*)
# - Filter interaktif (status, PO, hasil perbandingan)
# - Tiga mode tampilan (Semua, Analisis LPD/PODD, Analisis FPD/PSDD)
# - Ringkasan TRUE/FALSE + bar chart
# - Unduhan Excel (styled) & CSV (filtered)
# -----------------------------------------------------------------------------

from __future__ import annotations
import pandas as pd
import streamlit as st

from utils_pgd import (
    read_excel_file,
    read_csv_file,
    load_infor_from_many_csv,
    build_report,
    _blank_delay_columns,
    _export_excel_styled,
    today_str_id,
)

# ---------------------------- Page Setup -------------------------------------
st.set_page_config(page_title="📦 PGD Comparison", layout="wide")
st.title("📦 PGD Comparison Tracking — SAP vs Infor")

st.caption(
    "Upload 1 SAP Excel (*.xlsx) dan satu atau lebih Infor CSV (*.csv). "
    "Aplikasi akan merge, cleaning, comparison, visualisasi, filter (Execute), dan unduhan laporan (Excel/CSV)."
)

# ---------------------------- Sidebar: Upload --------------------------------
with st.sidebar:
    st.header("📤 Upload Files (PGD)")
    sap_file = st.file_uploader("SAP Excel (.xlsx)", type=["xlsx"], key="pgd_comp_sap")
    infor_files = st.file_uploader(
        "Infor CSV (boleh multi-file)", type=["csv"], accept_multiple_files=True, key="pgd_comp_infor"
    )
    st.markdown(
        """
**Tips:**
- SAP minimal punya `PO No.(Full)` & `Quantity`.
- Infor CSV minimal punya `PSDD`, `CRD`, dan `Line Aggregator`.
        """
    )

# ---------------------------- Helper (UI) ------------------------------------
def _uniq_vals(df: pd.DataFrame, col: str) -> list[str]:
    if col in df.columns:
        vals = df[col].dropna().astype(str).unique().tolist()
        return sorted(vals)
    return []

# ---------------------------- Main Logic -------------------------------------
if sap_file and infor_files:
    with st.status("Membaca & menggabungkan file...", expanded=True) as status:
        try:
            sap_df = read_excel_file(sap_file)
            st.write("SAP dibaca:", sap_df.shape)

            infor_csv_dfs = [read_csv_file(f) for f in infor_files]
            infor_all = load_infor_from_many_csv(
                infor_csv_dfs,
                on_info=lambda m: st.success(m),
                on_warn=lambda m: st.warning(m),
            )
            st.write("Total Infor (gabungan CSV):", infor_all.shape)

            if infor_all.empty:
                status.update(label="Gagal: tidak ada CSV Infor yang valid.", state="error")
            else:
                status.update(label="Sukses membaca semua file. Lanjut proses...", state="running")
                final_df = build_report(sap_df, infor_all)

                if final_df.empty:
                    status.update(label="Gagal membuat report — periksa kolom wajib.", state="error")
                else:
                    status.update(label="Report siap! ✅", state="complete")

                    # -------------------- Sidebar: Filters + Mode --------------
                    with st.sidebar.form("pgd_comp_filters_form"):
                        st.header("🔎 Filters & Mode")
                        status_opts = _uniq_vals(final_df, "Order Status Infor")
                        selected_status = st.multiselect(
                            "Order Status Infor", options=status_opts, default=status_opts, key="pgd_comp_status"
                        )
                        po_opts = _uniq_vals(final_df, "PO No.(Full)")
                        selected_pos = st.multiselect(
                            "PO No.(Full)", options=po_opts, placeholder="Pilih PO (opsional)", key="pgd_comp_po"
                        )

                        result_cols = [
                            "Result_Quantity",
                            "Result_FPD",
                            "Result_LPD",
                            "Result_CRD",
                            "Result_PSDD",
                            "Result_PODD",
                            "Result_PD",
                        ]
                        result_selections: dict[str, list[str]] = {}
                        for col in result_cols:
                            opts = _uniq_vals(final_df, col)
                            if opts:
                                result_selections[col] = st.multiselect(
                                    col, options=opts, default=opts, key=f"pgd_comp_{col}"
                                )

                        mode = st.radio(
                            "Mode tampilan data",
                            ["Semua Kolom", "Analisis LPD PODD", "Analisis FPD PSDD"],
                            horizontal=False,
                            key="pgd_comp_mode",
                        )
                        submitted = st.form_submit_button("🔄 Execute / Terapkan")

                    # -------------------- Apply filters after Execute -----------
                    if submitted or "pgd_comp_df_view" in st.session_state:
                        if submitted:
                            st.session_state["pgd_comp_selected_status"] = selected_status
                            st.session_state["pgd_comp_selected_pos"] = selected_pos
                            st.session_state["pgd_comp_result_selections"] = result_selections
                            st.session_state["pgd_comp_mode_val"] = mode

                        selected_status = st.session_state.get("pgd_comp_selected_status", status_opts)
                        selected_pos = st.session_state.get("pgd_comp_selected_pos", [])
                        result_selections = st.session_state.get("pgd_comp_result_selections", {})
                        mode = st.session_state.get("pgd_comp_mode_val", "Semua Kolom")

                        df_view = final_df.copy()
                        if selected_status:
                            df_view = df_view[df_view["Order Status Infor"].astype(str).isin(selected_status)]
                        if selected_pos:
                            df_view = df_view[df_view["PO No.(Full)"].astype(str).isin(selected_pos)]
                        for col, sel in result_selections.items():
                            base_opts = _uniq_vals(final_df, col)
                            if sel and set(sel) != set(base_opts):
                                df_view = df_view[df_view[col].astype(str).isin(sel)]

                        st.session_state["pgd_comp_df_view"] = df_view
                        st.session_state["pgd_comp_final_df"] = final_df

                        # -------------------- Preview sesuai mode ----------------
                        st.subheader("🔎 Preview Hasil (After Execute)")

                        def _subset(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
                            existing = [c for c in cols if c in df.columns]
                            missing = [c for c in cols if c not in df.columns]
                            if missing:
                                st.caption(f"Kolom tidak ditemukan & di-skip: {missing}")
                            if not existing:
                                st.warning("Tidak ada kolom yang cocok untuk mode ini.")
                                return pd.DataFrame()
                            return df[existing]

                        if mode == "Semua Kolom":
                            st.dataframe(df_view.head(100), use_container_width=True)
                        elif mode == "Analisis LPD PODD":
                            cols_lpd = [
                                "PO No.(Full)",
                                "Order Status Infor",
                                "DRC",
                                "Delay/Early - Confirmation PD",
                                "Delay/Early - Confirmation CRD",
                                "Infor Delay/Early - Confirmation CRD",
                                "Result_Delay_CRD",
                                "Delay - PO PSDD Update",
                                "Infor Delay - PO PSDD Update",
                                "Result_Delay_PSDD",
                                "Delay - PO PD Update",
                                "LPD",
                                "Infor LPD",
                                "Result_LPD",
                                "PODD",
                                "Infor PODD",
                                "Result_PODD",
                            ]
                            st.dataframe(_subset(df_view, cols_lpd).head(2000), use_container_width=True)
                        elif mode == "Analisis FPD PSDD":
                            cols_fpd_psdd = [
                                "PO No.(Full)",
                                "Order Status Infor",
                                "DRC",
                                "Delay/Early - Confirmation PD",
                                "Delay/Early - Confirmation CRD",
                                "Infor Delay/Early - Confirmation CRD",
                                "Result_Delay_CRD",
                                "Delay - PO PSDD Update",
                                "Infor Delay - PO PSDD Update",
                                "Result_Delay_PSDD",
                                "Delay - PO PD Update",
                                "FPD",
                                "Infor FPD",
                                "Result_FPD",
                                "PSDD",
                                "Infor PSDD",
                                "Result_PSDD",
                            ]
                            st.dataframe(_subset(df_view, cols_fpd_psdd).head(2000), use_container_width=True)

                        # -------------------- Summary TRUE/FALSE -----------------
                        st.subheader("📊 Comparison Summary (TRUE vs FALSE)")
                        existing_results = [
                            c
                            for c in [
                                "Result_Quantity",
                                "Result_FPD",
                                "Result_LPD",
                                "Result_CRD",
                                "Result_PSDD",
                                "Result_PODD",
                                "Result_PD",
                            ]
                            if c in df_view.columns
                        ]
                        if existing_results:
                            true_counts = [int(df_view[c].eq("TRUE").sum()) for c in existing_results]
                            false_counts = [int(df_view[c].eq("FALSE").sum()) for c in existing_results]
                            totals = [int(df_view[c].isin(["TRUE", "FALSE"]).sum()) for c in existing_results]
                            acc = [(t / tot * 100.0) if tot > 0 else 0.0 for t, tot in zip(true_counts, totals)]

                            summary_df = pd.DataFrame(
                                {
                                    "Metric": existing_results,
                                    "TRUE": true_counts,
                                    "FALSE": false_counts,
                                    "Total (TRUE+FALSE)": totals,
                                    "TRUE %": [round(a, 2) for a in acc],
                                }
                            )
                            st.dataframe(summary_df, use_container_width=True)
                            st.bar_chart(summary_df.set_index("Metric")[ ["TRUE", "FALSE"] ])

                            false_df_sorted = (
                                pd.DataFrame({"Metric": existing_results, "FALSE": false_counts})
                                .sort_values("FALSE", ascending=False)
                                .reset_index(drop=True)
                            )
                            st.markdown("**Distribusi FALSE (descending)**")
                            st.bar_chart(false_df_sorted.set_index("Metric")["FALSE"])  # satu seri
                            st.markdown("**🏆 TOP FALSE terbanyak**")
                            st.dataframe(false_df_sorted.head(min(5, len(false_df_sorted))), use_container_width=True)
                        else:
                            st.info("Kolom hasil perbandingan (Result_*) belum tersedia di data final.")

                        # -------------------- Downloads -------------------------
                        out_name_xlsx = f"PGD Comparison Tracking Report - {today_str_id()}.xlsx"
                        out_name_csv = f"PGD Comparison Tracking Report - {today_str_id()}.csv"

                        df_export = _blank_delay_columns(df_view)
                        excel_bytes = _export_excel_styled(df_export, sheet_name="Report")

                        st.download_button(
                            label="⬇️ Download Excel (Filtered, styled)",
                            data=excel_bytes,
                            file_name=out_name_xlsx,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )
                        st.download_button(
                            label="⬇️ Download CSV (Filtered)",
                            data=df_export.to_csv(index=False).encode("utf-8"),
                            file_name=out_name_csv,
                            mime="text/csv",
                            use_container_width=True,
                        )
                    else:
                        st.info("Atur filter/mode di sidebar, lalu klik **🔄 Execute / Terapkan**.")
        except Exception as e:
            status.update(label="Terjadi error saat menjalankan aplikasi.", state="error")
            st.error("Terjadi error saat menjalankan proses. Detail di bawah ini:")
            st.exception(e)
else:
    st.info("Unggah file SAP & Infor di sidebar untuk mulai.")
