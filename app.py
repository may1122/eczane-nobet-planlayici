import streamlit as st
from nobet_engine import run_schedule
import datetime
import pandas as pd
import plotly.express as px

st.set_page_config(
    page_title="Eczane Nöbet Sistemi",
    page_icon="💊",
    layout="wide"
)

st.title("💊 Eczane Nöbet Yönetim Sistemi")

# =========================
# TAB YAPISI
# =========================

tab1, tab2 = st.tabs([
    "📅 Nöbet Planlayıcı",
    "📊 Nöbet Analiz Sistemi"
])

# =====================================================
# TAB 1 → NÖBET PLANLAYICI
# =====================================================

with tab1:

    st.caption("Akıllı nöbet planlama sistemi")

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

    st.subheader("🔧 Eczane Değişiklikleri")

    degisim = st.toggle("Eczane ekleme / çıkarma yapılacak mı?")

    eklenme = {}
    cikma = {}

    if degisim:

        col1, col2 = st.columns(2)

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


# =====================================================
# TAB 2 → NÖBET ANALİZ DASHBOARD
# =====================================================

with tab2:

    st.title("💊 Eczane Nöbet Takip Sistemi")

    @st.cache_data
    def load_excel(file):

        xls = pd.ExcelFile(file)

        all_data = []
        genel = None

        for sheet in xls.sheet_names:

            df = pd.read_excel(file, sheet_name=sheet)
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

            if "GENEL" in sheet.upper():

                genel = df[[
                    "Eczane",
                    "Grup",
                    "Geçmiş Katsayı",
                    "Geçmiş Bayram",
                    "Toplam Nöbet",
                    "Toplam Katsayı",
                    "Bayram"
                ]]

                continue

            if "Tarih" not in df.columns:
                continue

            df_long = df.melt(
                id_vars=["Tarih", "Gün"],
                var_name="Grup",
                value_name="Eczane"
            )

            df_long = df_long.dropna(subset=["Eczane"])
            df_long["Ay"] = sheet

            all_data.append(df_long)

        df = pd.concat(all_data, ignore_index=True)

        return df, genel


    file = st.file_uploader("Excel dosyasını yükleyin", type=["xlsx"])

    if not file:
        st.info("Başlamak için Excel dosyasını yükleyin.")
        st.stop()

    df, genel = load_excel(file)

    menu = st.sidebar.radio("Menü", [
        "Genel Özet",
        "Tarih Seç",
        "Aylık Takvim",
        "Grup Analizi",
        "Eczane Analizi"
    ])

    if menu == "Genel Özet":

        col1, col2, col3, col4 = st.columns(4)

        col1.metric("Toplam Nöbet", len(df))
        col2.metric("Toplam Eczane", df["Eczane"].nunique())
        col3.metric("Toplam Ay", df["Ay"].nunique())
        col4.metric("Ortalama Nöbet", round(len(df) / df["Eczane"].nunique(), 2))

        gun_sayim = df["Gün"].value_counts().reset_index()
        gun_sayim.columns = ["Gün", "Sayı"]

        fig = px.pie(
            gun_sayim,
            names="Gün",
            values="Sayı",
            hole=0.4
        )

        st.plotly_chart(fig, use_container_width=True)

    elif menu == "Tarih Seç":

        tarih = st.selectbox("Tarih", sorted(df["Tarih"].unique()))
        sonuc = df[df["Tarih"] == tarih]
        st.dataframe(sonuc)

    elif menu == "Aylık Takvim":

        ay = st.selectbox("Ay seç", sorted(df["Ay"].unique()))
        sonuc = df[df["Ay"] == ay]

        pivot = sonuc.pivot(
            index="Tarih",
            columns="Grup",
            values="Eczane"
        )

        pivot = pivot.fillna("")
        st.dataframe(pivot, use_container_width=True)

    elif menu == "Grup Analizi":

        grup = st.selectbox(
            "Grup seç",
            sorted(genel["Grup"].unique())
        )

        sonuc = df[df["Grup"] == grup]

        sayim = (
            sonuc
            .groupby(["Gün","Eczane"])
            .size()
            .reset_index(name="Nöbet Sayısı")
        )

        fig = px.bar(
            sayim,
            x="Gün",
            y="Nöbet Sayısı",
            color="Eczane",
            barmode="group"
        )

        st.plotly_chart(fig, use_container_width=True)

    elif menu == "Eczane Analizi":

        eczane = st.selectbox(
            "Eczane",
            sorted(df["Eczane"].unique())
        )

        sonuc = df[df["Eczane"] == eczane]

        st.metric("Toplam Nöbet", len(sonuc))
        st.dataframe(sonuc.sort_values("Tarih"))
