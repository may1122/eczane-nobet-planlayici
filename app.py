import streamlit as st
from nobet_engine import run_schedule
import datetime

st.set_page_config(page_title="Eczane Nöbet Planı", layout="wide")

st.title("💊 Eczane Nöbet Planlayıcı")

yil = st.number_input("Yıl", value=datetime.datetime.now().year)
ay = st.number_input("Başlangıç Ayı", min_value=1, max_value=12, value=1)
kac_ay = st.number_input("Kaç Ay Planlansın", min_value=1, max_value=12, value=3)

if st.button("Plan Oluştur"):

    with st.spinner("Plan hazırlanıyor..."):
        file1, file2 = run_schedule(yil, ay, kac_ay)

    st.success("Plan hazır")

    with open(file1, "rb") as f:
        st.download_button(
            "Excel indir",
            f,
            file_name="nobet_plani.xlsx"
        )

    with open(file2, "rb") as f:
        st.download_button(
            "Aylık istatistik indir",
            f,
            file_name="aylik_detay.xlsx"
        )
