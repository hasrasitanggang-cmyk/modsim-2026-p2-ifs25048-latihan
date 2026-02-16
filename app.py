import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import os

# =====================================
# KONFIGURASI APLIKASI
# =====================================
st.set_page_config(
    page_title="🎓 Analisis Kuesioner Pembelajaran",
    layout="wide",
    page_icon="📊"
)

# =====================================
# MEMUAT DATA KUESIONER (OTOMATIS)
# =====================================
@st.cache_data
def muat_data_kuesioner():
    """Memuat dan memproses data kuesioner dari Excel"""
    file_path = "data_kuesioner.xlsx"
    
    # Validasi keberadaan file
    if not os.path.exists(file_path):
        st.error(f"❌ File '{file_path}' tidak ditemukan di direktori saat ini: `{os.getcwd()}`")
        st.info("Pastikan file 'data_kuesioner.xlsx' berada di folder yang sama dengan file app.py")
        return None, None, None, None
    
    try:
        # Baca ketiga sheet yang diperlukan
        df_kuesioner = pd.read_excel(file_path, sheet_name="Kuesioner")
        df_keterangan = pd.read_excel(file_path, sheet_name="Keterangan")
        df_pertanyaan = pd.read_excel(file_path, sheet_name="Pertanyaan")
        
        # Validasi struktur data
        kolom_wajib_kuesioner = ["Partisipan"] + [f"Q{i}" for i in range(1, 18)]
        if not all(col in df_kuesioner.columns for col in kolom_wajib_kuesioner):
            st.error("❌ Struktur kolom pada sheet 'Kuesioner' tidak sesuai. Pastikan ada kolom Q1-Q17 dan Partisipan")
            return None, None, None, None
        
        if not all(col in df_keterangan.columns for col in ["Singkatan", "Deskripsi", "Point"]):
            st.error("❌ Sheet 'Keterangan' harus memiliki kolom: Singkatan, Deskripsi, Point")
            return None, None, None, None
        
        if not all(col in df_pertanyaan.columns for col in ["Kode", "Pertanyaan"]):
            st.error("❌ Sheet 'Pertanyaan' harus memiliki kolom: Kode, Pertanyaan")
            return None, None, None, None
        
        # Buat mapping skor
        mapping_skor = dict(zip(df_keterangan["Singkatan"], df_keterangan["Point"]))
        mapping_pertanyaan = dict(zip(df_pertanyaan["Kode"], df_pertanyaan["Pertanyaan"]))
        
        # Konversi respons ke skor numerik
        kolom_pertanyaan = [f"Q{i}" for i in range(1, 18)]
        df_skor = df_kuesioner[["Partisipan"] + kolom_pertanyaan].copy()
        
        # Handle nilai kosong/NaN
        for kolom in kolom_pertanyaan:
            if kolom in df_skor.columns:
                df_skor[kolom] = df_skor[kolom].map(mapping_skor).fillna(0)
        
        # Hitung total skor dan rata-rata
        df_skor["Total_Skor"] = df_skor[kolom_pertanyaan].sum(axis=1)
        df_skor["Rata_Rata"] = df_skor[kolom_pertanyaan].mean(axis=1)
        
        return df_skor, df_kuesioner, mapping_pertanyaan, df_keterangan
    
    except Exception as e:
        st.error(f"❌ Gagal memuat data kuesioner: {str(e)}")
        st.info("Pastikan struktur file Excel sesuai dengan format yang diharapkan")
        return None, None, None, None

# Load data secara otomatis saat aplikasi dibuka
df_skor, df_raw, mapping_pertanyaan, df_keterangan = muat_data_kuesioner()

# Validasi data sebelum melanjutkan
if df_skor is None or df_skor.empty:
    st.stop()

# =====================================
# DASHBOARD UTAMA (LANGSUNG TAMPIL)
# =====================================
st.title("🎓 Dashboard Analisis Kuesioner Pembelajaran")
st.markdown("Analisis otomatis dari **113 partisipan** - Data dimuat langsung tanpa input manual")
st.markdown("---")

# KPI Cards
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Total Responden", f"{len(df_skor):,}")
with col2:
    st.metric("Rata-rata Skor", f"{df_skor['Rata_Rata'].mean():.2f}/6.00")
with col3:
    st.metric("Skor Tertinggi", f"{df_skor['Total_Skor'].max():.0f}/102")
with col4:
    st.metric("Skor Terendah", f"{df_skor['Total_Skor'].min():.0f}/102")

st.markdown("---")

# Visualisasi Utama
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Ringkasan Skor", 
    "📈 Distribusi Jawaban", 
    "❓ Analisis Pertanyaan", 
    "🔍 Insight Kritis"
])

with tab1:
    st.subheader("Distribusi Skor Total Responden")
    fig = px.histogram(
        df_skor, 
        x="Total_Skor",
        nbins=15,
        color_discrete_sequence=["#1f77b4"],
        labels={"Total_Skor": "Total Skor (Maks: 102)"},
        marginal="box"
    )
    fig.update_layout(height=400)
    st.plotly_chart(fig, use_container_width=True)
    
    # Tabel ringkasan
    st.subheader("Statistik Skor per Pertanyaan")
    kolom_pertanyaan = [f"Q{i}" for i in range(1, 18)]
    stats = df_skor[kolom_pertanyaan].agg(["mean", "min", "max", "std"]).round(2).T
    stats.columns = ["Rata-rata", "Min", "Maks", "Std Dev"]
    stats["Pertanyaan"] = stats.index.map(mapping_pertanyaan)
    stats = stats[["Pertanyaan", "Rata-rata", "Min", "Maks", "Std Dev"]].sort_values("Rata-rata", ascending=False)
    st.dataframe(stats, use_container_width=True)

with tab2:
    st.subheader("Distribusi Jenis Respons")
    # Hitung frekuensi semua respons
    kolom_pertanyaan = [f"Q{i}" for i in range(1, 18)]
    semua_respons = df_raw[kolom_pertanyaan].stack().value_counts().reset_index()
    semua_respons.columns = ["Respons", "Frekuensi"]
    
    # Tambahkan deskripsi dari sheet Keterangan
    semua_respons = semua_respons.merge(
        df_keterangan[["Singkatan", "Deskripsi"]], 
        left_on="Respons", 
        right_on="Singkatan",
        how="left"
    )
    semua_respons["Deskripsi"] = semua_respons["Deskripsi"].fillna(semua_respons["Respons"])
    semua_respons["Persentase"] = (semua_respons["Frekuensi"] / semua_respons["Frekuensi"].sum() * 100).round(1)
    
    fig = px.pie(
        semua_respons,
        values="Frekuensi",
        names="Deskripsi",
        hole=0.4,
        color_discrete_sequence=px.colors.sequential.RdBu_r
    )
    fig.update_traces(textposition="inside", textinfo="percent+label")
    st.plotly_chart(fig, use_container_width=True)
    
    st.dataframe(semua_respons[["Deskripsi", "Frekuensi", "Persentase"]].sort_values("Frekuensi", ascending=False), 
                 use_container_width=True)

with tab3:
    st.subheader("Analisis Detail per Pertanyaan")
    pertanyaan_pilih = st.selectbox(
        "Pilih pertanyaan untuk analisis mendalam:",
        options=[f"Q{i}" for i in range(1, 18)],
        format_func=lambda x: f"{x}: {mapping_pertanyaan[x][:70]}..."
    )
    
    col_a, col_b = st.columns(2)
    with col_a:
        # Distribusi skor untuk pertanyaan terpilih
        fig = px.histogram(
            df_skor,
            x=pertanyaan_pilih,
            nbins=6,
            labels={pertanyaan_pilih: "Skor"},
            color_discrete_sequence=["#2ca02c"]
        )
        fig.update_layout(title=f"Distribusi Skor: {pertanyaan_pilih}", height=350)
        st.plotly_chart(fig, use_container_width=True)
    
    with col_b:
        # Box plot perbandingan
        fig = px.box(
            df_skor,
            y=pertanyaan_pilih,
            labels={pertanyaan_pilih: "Skor"},
            color_discrete_sequence=["#ff7f0e"]
        )
        fig.update_layout(title="Sebaran Respons", height=350, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    
    st.info(f"**Pertanyaan Lengkap:** {mapping_pertanyaan[pertanyaan_pilih]}")

with tab4:
    st.subheader("🔍 Insight Kritis")
    
    # Identifikasi area perbaikan
    kolom_pertanyaan = [f"Q{i}" for i in range(1, 18)]
    rata_pertanyaan = df_skor[kolom_pertanyaan].mean().sort_values().reset_index()
    rata_pertanyaan.columns = ["Kode", "Skor_Rata"]
    rata_pertanyaan["Pertanyaan"] = rata_pertanyaan["Kode"].map(mapping_pertanyaan)
    
    # Area perlu perbaikan (skor terendah)
    perlu_perbaikan = rata_pertanyaan.nsmallest(3, "Skor_Rata")
    # Area kekuatan (skor tertinggi)
    kekuatan = rata_pertanyaan.nlargest(3, "Skor_Rata")
    
    col_x, col_y = st.columns(2)
    
    with col_x:
        st.warning("⚠️ **Area Perlu Perbaikan**")
        for _, row in perlu_perbaikan.iterrows():
            st.markdown(f"- **{row['Kode']}**: {row['Pertanyaan'][:70]}... (Skor: {row['Skor_Rata']:.2f}/6.00)")
    
    with col_y:
        st.success("✅ **Kekuatan Utama**")
        for _, row in kekuatan.iterrows():
            st.markdown(f"- **{row['Kode']}**: {row['Pertanyaan'][:70]}... (Skor: {row['Skor_Rata']:.2f}/6.00)")
    
    # Rekomendasi berdasarkan Q16 (Kepuasan) dan Q17 (Pengetahuan)
    st.subheader("🎯 Rekomendasi Strategis")
    kepuasan = df_skor["Q16"].mean()
    pengetahuan = df_skor["Q17"].mean()
    
    col_r1, col_r2 = st.columns(2)
    
    with col_r1:
        if kepuasan < 4.5:
            st.error(f"❗ Tingkat kepuasan ({kepuasan:.2f}/6.00) perlu ditingkatkan")
            st.markdown("""
            **Fokus perbaikan:**
            - Q3: Rencana perkuliahan dan evaluasi
            - Q4: Keterstrukturan materi
            - Q6: Penjelasan dosen/TA
            - Q15: Transparansi penilaian
            """)
        else:
            st.success(f"✅ Tingkat kepuasan baik ({kepuasan:.2f}/6.00)")
    
    with col_r2:
        if pengetahuan < 4.7:
            st.error(f"❗ Pencapaian pengetahuan ({pengetahuan:.2f}/6.00) perlu dioptimalkan")
            st.markdown("""
            **Fokus perbaikan:**
            - Q9: Keragaman sumber belajar
            - Q12: Kecukupan waktu ujian
            - Q13: Kesesuaian contoh soal & materi
            """)
        else:
            st.success(f"✅ Pencapaian pengetahuan baik ({pengetahuan:.2f}/6.00)")

# =====================================
# FOOTER
# =====================================
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: #666;'>
    📊 Analisis Kuesioner Pembelajaran | © 2026<br>
    <small>Data dimuat otomatis dari <code>data_kuesioner.xlsx</code> - Tidak diperlukan input manual</small>
    </div>
    """,
    unsafe_allow_html=True
)

# =====================================
# STYLE TAMBAHAN
# =====================================
st.markdown("""
<style>
    .stMetric {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 10px;
        border-left: 4px solid #4e79a7;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    h1, h2, h3 {
        color: #2c3e50 !important;
    }
    .css-1d391kg {
        padding-top: 1.5rem;
    }
    div[data-testid="stTabs"] button {
        font-weight: 500;
    }
    .stAlert {
        padding: 1rem;
        border-radius: 8px;
    }
</style>
""", unsafe_allow_html=True)
