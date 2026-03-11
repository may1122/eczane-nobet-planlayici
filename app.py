import streamlit as st
from nobet_engine import run_schedule
import datetime

st.title("💊 Eczane Nöbet Planlayıcı")

yil = st.number_input("Yıl", value=datetime.datetime.now().year)
ay = st.number_input("Başlangıç Ayı", 1,12,1)
kac_ay = st.number_input("Kaç Ay",1,12,3)

if st.button("Plan Oluştur"):

    file1,file2 = run_schedule(yil,ay,kac_ay)

    with open(file1,"rb") as f:
        st.session_state.plan_data = f.read()

    with open(file2,"rb") as f:
        st.session_state.aylik_data = f.read()


if "plan_data" in st.session_state and "aylik_data" in st.session_state:

    st.download_button(
        "Excel indir",
        st.session_state.plan_data,
        "nobet_plani.xlsx"
    )

    st.download_button(
        "Aylık istatistik indir",
        st.session_state.aylik_data,
        "aylik_detay.xlsx"
    )
