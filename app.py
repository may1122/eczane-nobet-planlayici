import streamlit as st
from nobet_engine import run_schedule
import datetime

st.title("💊 Eczane Nöbet Planlayıcı")

# ================================
# ECZANE EKLE / ÇIKAR PANELİ
# ================================

st.subheader("Eczane Değişikliği")

degisim = st.radio(
    "Eczane ekle / çıkar yapmak istiyor musunuz?",
    ("Hayır","Evet")
)

eklenme = {}
cikma = {}

if degisim == "Evet":

    # ====================
    # ECZANE EKLE
    # ====================

    st.markdown("### ➕ Eczane Ekle")

    eczane_ekle = st.text_input("Eklenecek Eczane İsmi")

    eczane_grup = st.selectbox(
        "Eklenecek Grup",
        ["A1","A2","A3",
         "B1","B2","B3",
         "C1","C2","C3",
         "D1","D2","D3",
         "E1","E2","E3",
         "F1","F2","F3",
         "G1","G2","G3"]
    )

    ekleme_tarihi = st.date_input(
        "Eczane Eklenme Tarihi",
        value=datetime.date.today()
    )

    if eczane_ekle:
        eklenme[eczane_ekle.upper()] = {
            "tarih": ekleme_tarihi,
            "grup": eczane_grup
        }

    # ====================
    # ECZANE ÇIKAR
    # ====================

    st.markdown("### ➖ Eczane Çıkar")

    eczane_cikar = st.text_input("Çıkarılacak Eczane İsmi")

    cikis_tarihi = st.date_input(
        "Eczane Çıkış Tarihi",
        value=datetime.date.today(),
        key="cikis"
    )

    if eczane_cikar:
        cikma[eczane_cikar.upper()] = cikis_tarihi


st.divider()

# ================================
# PLAN PARAMETRELERİ
# ================================

yil = st.number_input("Yıl", value=datetime.datetime.now().year)

ay = st.number_input(
    "Başlangıç Ayı",
    min_value=1,
    max_value=12,
    value=1
)

kac_ay = st.number_input(
    "Kaç Ay",
    min_value=1,
    max_value=12,
    value=3
)

# ================================
# PLAN OLUŞTUR
# ================================

if st.button("Plan Oluştur"):

    file1, file2 = run_schedule(
        yil,
        ay,
        kac_ay,
        eklenme,
        cikma
    )

    with open(file1,"rb") as f:
        st.session_state.plan_data = f.read()

    with open(file2,"rb") as f:
        st.session_state.aylik_data = f.read()


# ================================
# DOWNLOAD BUTONLARI
# ================================

if "plan_data" in st.session_state and "aylik_data" in st.session_state:

    st.download_button(
        "📥 Nöbet Planını İndir",
        st.session_state.plan_data,
        "nobet_plani.xlsx"
    )

    st.download_button(
        "📊 Aylık İstatistik İndir",
        st.session_state.aylik_data,
        "aylik_detay.xlsx"
    )
