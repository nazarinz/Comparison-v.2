
# Home.py — PGD Multi-Page App
import streamlit as st

st.set_page_config(page_title="PGD Apps — Home", page_icon="🧰", layout="wide")

st.title("🧰 PGD Apps — Home")
st.markdown("""
Selamat datang di **PGD Multi-Page App**.

**Halaman yang tersedia:**
1. **📦 PGD Comparison** — Merge SAP vs Infor, cleaning, comparison, visualisasi, filter, dan unduh laporan.
2. **🧩 PO Splitter** — Bagi daftar PO menjadi beberapa bagian (mis. 5.000/baris).

Gunakan menu **Pages** di kiri untuk berpindah halaman.
""")
