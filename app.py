import streamlit as st
from nobet_engine import run_schedule
import datetime

# ==============================
# SAYFA AYARI
# ==============================

st.set_page_config(
    page_title="Eczane Nöbet Planlayıcı",
    page_icon="💊",
    layout="wide"
)

st.title("💊 Kahramanmaraş Eczane Nöbet Planlayıcı")
st.caption("Akıllı nöbet planlama sistemi")

# ==============================
# SIDEBAR - PLAN AYARLARI
# ==============================

with st.sidebar:

    st.header("📅 Plan Parametreleri")

    yil = st.number_input(
        "Yıl",
        value=datetime.datetime.now().year
    )

    ay = st.number_input(
        "Başlangıç Ayı",
        min_value=1,
        max_value=12,
        value=1
    )

    kac_ay = st.number_input(
        "Kaç Ay Planlansın",
        min_value=1,
        max_value=12,
        value=3
    )

    st.divider()

    planla = st.button("🚀 Plan Oluştur", use_container_width=True)


# ==============================
# ECZANE DEĞİŞİKLİK PANELİ
# ==============================

st.subheader("🔧 Eczane Değişiklikleri")

degisim = st.toggle("Eczane ekleme / çıkarma yapılacak mı?")

eklenme = {}
cikma = {}

if degisim:

    col1, col2 = st.columns(2)

    # ======================
    # ECZANE EKLE
    # ======================

    with col1:

        st.markdown("### ➕ Eczane Ekle")

        eczane_ekle = st.text_input("Eczane İsmi")

        eczane_grup = st.selectbox(
            "Grup",
            ["A1","A2","A3",
             "B1","B2","B3",
             "C1","C2","C3",
             "D1","D2","D3",
             "E1","E2","E3",
             "F1","F2","F3",
             "G1","G2","G3"]
        )

        ekleme_tarihi = st.date_input(
            "Eklenme Tarihi",
            value=datetime.date.today()
        )

        if eczane_ekle:
            eklenme[eczane_ekle.upper()] = {
                "tarih": ekleme_tarihi,
                "grup": eczane_grup
            }

    # ======================
    # ECZANE ÇIKAR
    # ======================

    with col2:

        st.markdown("### ➖ Eczane Çıkar")

        eczane_cikar = st.text_input("Eczane İsmi", key="cikar")

        cikis_tarihi = st.date_input(
            "Çıkış Tarihi",
            value=datetime.date.today()
        )

        if eczane_cikar:
            cikma[eczane_cikar.upper()] = cikis_tarihi


st.divider()

# ==============================
# PLAN OLUŞTUR
# ==============================

if planla:

    with st.spinner("Plan oluşturuluyor..."):

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

    st.success("Plan başarıyla oluşturuldu!")

# ==============================
# DOWNLOAD PANELİ
# ==============================

if "plan_data" in st.session_state:

    st.subheader("📥 Dosyalar")

    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            "📄 Nöbet Planını İndir",
            st.session_state.plan_data,
            "nobet_plani.xlsx",
            use_container_width=True
        )

    with col2:
        st.download_button(
            "📊 Aylık İstatistik İndir",
            st.session_state.aylik_data,
            "aylik_detay.xlsx",
            use_container_width=True
        )
