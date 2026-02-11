import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np

# =====================================================
# FUNGSI BANTU (HELPER FUNCTIONS)
# =====================================================
def format_angka_otomatis(value):
    """
    Format angka dengan skala otomatis (Rb, Jt, M)
    
    Parameters:
    - value: Angka yang akan diformat
    
    Returns:
    - String yang sudah diformat
    """
    if value is None or pd.isna(value):
        return "0"
    
    if isinstance(value, str):
        try:
            value = float(value.replace(".", "").replace(",", "."))
        except:
            return value
    
    try:
        v = abs(float(value))
    except (ValueError, TypeError):
        return str(value)
    
    if v >= 1_000_000_000:
        return f"Rp {value/1_000_000_000:.2f} M"
    elif v >= 1_000_000:
        return f"Rp {value/1_000_000:.2f} Jt"
    elif v >= 1_000:
        return f"Rp {value/1_000:.2f} Rb"
    else:
        return f"Rp {value:,.0f}".replace(",", ".")

def format_angka_tanpa_rp(value):
    """
    Format angka tanpa 'Rp' di depan (untuk hover)
    
    Parameters:
    - value: Angka yang akan diformat
    
    Returns:
    - String yang sudah diformat tanpa 'Rp'
    """
    result = format_angka_otomatis(value)
    if result.startswith("Rp "):
        return result[3:]
    return result

def tambahkan_hover_uang(fig, df, kolom, tipe="bar"):
    """
    Tambahkan hover template untuk visualisasi uang
    
    Parameters:
    - fig: Figure Plotly
    - df: DataFrame
    - kolom: Nama kolom yang berisi data uang
    - tipe: Jenis plot ('bar', 'line', 'hbar', 'scatter')
    """
    if tipe == "hbar":
        fig.update_traces(
            customdata=df[kolom].apply(format_angka_tanpa_rp),
            hovertemplate="<b>%{y}</b><br>Rp%{customdata}<extra></extra>"
        )
    elif tipe == "line":
        fig.update_traces(
            customdata=df[kolom].apply(format_angka_tanpa_rp),
            hovertemplate="%{x}<br>Rp%{customdata}<extra></extra>"
        )
    else:  # bar, scatter, dll
        fig.update_traces(
            customdata=df[kolom].apply(format_angka_tanpa_rp),
            hovertemplate="<b>%{x}</b><br>Rp%{customdata}<extra></extra>"
        )
    return fig

# =====================================================
# KONFIGURASI APLIKASI
# =====================================================
st.set_page_config(
    page_title="ITDel Tech Analytics 2025",
    layout="wide",
    page_icon="📊"
)

DATA_FILE = "itdeltech_2025.csv"

# =====================================================
# MEMUAT DATA (DATA LOADING)
# =====================================================
@st.cache_data
def muat_data():
    """
    Memuat dan mempersiapkan data dari file CSV
    """
    try:
        df = pd.read_csv(DATA_FILE, parse_dates=["date"])
        
        # Kolom yang berisi data uang
        kolom_uang = ["unit_price", "sales_qty", "revenue", "cost", "profit"]
        
        # Konversi kolom uang ke numerik
        for kolom in kolom_uang:
            if kolom in df.columns:
                if df[kolom].dtype == "object":
                    # Bersihkan format string
                    df[kolom] = (
                        df[kolom].astype(str)
                        .str.replace("Rp", "", regex=False)
                        .str.replace(".", "", regex=False)
                        .str.replace(",", ".", regex=False)
                    )
                df[kolom] = pd.to_numeric(df[kolom], errors="coerce")
        
        # Ekstrak informasi bulan
        df["bulan"] = df["date"].dt.month
        df["nama_bulan"] = df["date"].dt.strftime("%B")
        df["tahun"] = df["date"].dt.year
        df["hari"] = df["date"].dt.day
        df["minggu"] = df["date"].dt.isocalendar().week
        
        # Optimasi memori dengan tipe data kategori
        for kolom in ["city", "category", "channel", "product_name", "customer_type"]:
            if kolom in df.columns:
                df[kolom] = df[kolom].astype("category")
        
        # Hitung profit margin jika tidak ada
        if "profit_margin" not in df.columns:
            df["profit_margin"] = (df["profit"] / df["revenue"] * 100).round(2)
        
        return df
    except Exception as e:
        st.error(f"❌ Gagal memuat data: {str(e)}")
        return pd.DataFrame()

df = muat_data()

# Validasi data
if df.empty:
    st.error("❌ Data tidak ditemukan atau file kosong")
    st.stop()

if "revenue" not in df.columns:
    st.error("❌ Kolom 'revenue' tidak ditemukan dalam data")
    st.stop()

# =====================================================
# SIDEBAR (NAVIGASI)
# =====================================================
st.sidebar.title("📊 ITDel Tech")
st.sidebar.markdown("---")

# Filter tanggal
st.sidebar.subheader("🗓️ Filter Tanggal")
if not df.empty:
    tanggal_min = df["date"].min().date()
    tanggal_max = df["date"].max().date()
    rentang_tanggal = st.sidebar.date_input(
        "Pilih rentang tanggal",
        [tanggal_min, tanggal_max],
        min_value=tanggal_min,
        max_value=tanggal_max
    )
    
    if len(rentang_tanggal) == 2:
        df = df[(df["date"].dt.date >= rentang_tanggal[0]) & 
                (df["date"].dt.date <= rentang_tanggal[1])]

# Filter kategori
st.sidebar.subheader("🏷️ Filter Kategori")
if "category" in df.columns:
    semua_kategori = ["Semua"] + sorted(df["category"].unique().tolist())
    kategori_terpilih = st.sidebar.multiselect(
        "Pilih kategori produk",
        semua_kategori,
        default=["Semua"]
    )
    
    if "Semua" not in kategori_terpilih and kategori_terpilih:
        df = df[df["category"].isin(kategori_terpilih)]

# Filter kota
st.sidebar.subheader("🏙️ Filter Kota")
if "city" in df.columns:
    semua_kota = ["Semua"] + sorted(df["city"].unique().tolist())
    kota_terpilih = st.sidebar.multiselect(
        "Pilih kota",
        semua_kota,
        default=["Semua"]
    )
    
    if "Semua" not in kota_terpilih and kota_terpilih:
        df = df[df["city"].isin(kota_terpilih)]

st.sidebar.markdown("---")

# Menu navigasi
menu = st.sidebar.radio(
    "📋 PILIH VISUALISASI",
    [
        "📊 Dashboard Utama",
        "📈 Tren Pendapatan",
        "📊 Performa Produk",
        "🏙️ Performa Kota",
        "📦 Analisis Kategori",
        "🛒 Analisis Channel",
        "💰 Analisis Profitabilitas",
        "📉 Analisis Diskonting",
        "📅 Analisis Waktu",
        "📱 Analisis Pelanggan",
        "📋 Tabel Data Lengkap"
    ]
)

# =====================================================
# DASHBOARD UTAMA
# =====================================================
if menu == "📊 Dashboard Utama":
    st.title("📊 ITDel Tech - Dashboard")
    st.markdown("---")
    
    # KPI Cards - Baris 1
    col1, col2 = st.columns(2)

    with col1:
        total_pendapatan = df["revenue"].sum()
        st.metric(
            label="💰 Total Pendapatan",
            value=format_angka_otomatis(total_pendapatan),
            delta=f"{len(df):,} transaksi".replace(",", ".")
        )

    with col2:
        total_keuntungan = df["profit"].sum()
        margin_profit = (total_keuntungan / total_pendapatan * 100) if total_pendapatan > 0 else 0
        st.metric(
            label="📈 Total Keuntungan",
            value=format_angka_otomatis(total_keuntungan),
            delta=f"{margin_profit:.1f}% margin"
        )

    col3, col4 = st.columns(2)

    with col3:
        total_unit = df["sales_qty"].sum()
        rata_harga = df["unit_price"].mean() if not df["unit_price"].isna().all() else 0
        st.metric(
            label="📦 Unit Terjual",
            value=f"{total_unit:,.0f}".replace(",", "."),
            delta=f"Rp {rata_harga:,.0f}/unit".replace(",", ".")
        )

    with col4:
        kota_aktif = df["city"].nunique()
        produk_aktif = df["product_name"].nunique()
        st.metric(
            label="📍 Jangkauan",
            value=f"{kota_aktif} Kota",
            delta=f"{produk_aktif} Produk"
        )

    col5, col6 = st.columns(2)

    with col5:
        pelanggan_unik = df["customer_type"].nunique() if "customer_type" in df.columns else 0
        st.metric(
            label="👥 Tipe Pelanggan",
            value=pelanggan_unik,
            delta="Segmentasi"
        )

    with col6:
        channel_aktif = df["channel"].nunique() if "channel" in df.columns else 0
        st.metric(
            label="🛒 Channel Aktif",
            value=channel_aktif,
            delta="Distribusi"
        )

    col7, col8 = st.columns(2)

    with col7:
        rata_transaksi = total_pendapatan / len(df) if len(df) > 0 else 0
        st.metric(
            label="💵 Rata-rata Transaksi",
            value=format_angka_otomatis(rata_transaksi),
            delta="Per transaksi"
        )

    with col8:
        pertumbuhan = (
            df.groupby("bulan")["revenue"].sum().pct_change().iloc[-1] * 100
            if "bulan" in df.columns and len(df) > 1
            else 0
        )
        st.metric(
            label="📊 Growth Bulanan",
            value=f"{pertumbuhan:.1f}%" if not pd.isna(pertumbuhan) else "0%",
            delta="MoM"
        )

    st.markdown("---")
        
    # Visualisasi Utama - Baris 1
    col_a, col_b = st.columns(2)
    
    with col_a:
        st.subheader("📈 Tren Pendapatan Bulanan")
        df_bulanan = df.groupby(["bulan", "nama_bulan"], as_index=False).agg({
            "revenue": "sum",
            "profit": "sum"
        }).sort_values("bulan")
        
        df_bulanan["revenue_miliar"] = df_bulanan["revenue"] / 1_000_000_000
        df_bulanan["profit_miliar"] = df_bulanan["profit"] / 1_000_000_000
        
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=df_bulanan["nama_bulan"],
            y=df_bulanan["revenue_miliar"],
            mode="lines+markers",
            name="Pendapatan",
            line=dict(color="#1f77b4", width=3),
            marker=dict(size=8)
        ))
        fig.add_trace(go.Bar(
            x=df_bulanan["nama_bulan"],
            y=df_bulanan["profit_miliar"],
            name="Keuntungan",
            marker_color="#ff7f0e",
            opacity=0.6
        ))
        
        fig.update_layout(
            title="Pendapatan vs Keuntungan per Bulan",
            xaxis_title="Bulan",
            yaxis_title="Nilai (Miliar Rupiah)",
            hovermode="x unified",
            height=400
        )
        
        # Tambahkan hover
        fig.update_traces(
            customdata=df_bulanan[["revenue", "profit"]].applymap(format_angka_tanpa_rp),
            hovertemplate=(
                "<b>%{x}</b><br>"
                "Pendapatan: Rp%{customdata[0]}<br>"
                "Keuntungan: Rp%{customdata[1]}"
                "<extra></extra>"
            )
        )
        
        st.plotly_chart(fig, use_container_width=True)
    
    with col_b:
        st.subheader("📊 Top 5 Produk Terlaris")
        df_produk = df.groupby("product_name", as_index=False).agg({
            "revenue": "sum",
            "sales_qty": "sum",
            "profit": "sum"
        }).nlargest(5, "revenue")
        
        df_produk["revenue_miliar"] = df_produk["revenue"] / 1_000_000_000
        
        fig = px.bar(
            df_produk,
            x="revenue_miliar",
            y="product_name",
            orientation="h",
            color="revenue_miliar",
            color_continuous_scale="Blues",
            text=df_produk["revenue"].apply(lambda x: f"Rp{format_angka_tanpa_rp(x)}")
        )
        
        fig.update_layout(
            title="Pendapatan per Produk",
            xaxis_title="Pendapatan (Miliar Rupiah)",
            yaxis_title="Produk",
            height=400,
            showlegend=False
        )
        
        fig = tambahkan_hover_uang(fig, df_produk, "revenue", "hbar")
        st.plotly_chart(fig, use_container_width=True)
    
    # Visualisasi Utama - Baris 2
    col_c, col_d = st.columns(2)
    
    with col_c:
        st.subheader("🏙️ Distribusi Pendapatan per Kota")
        df_kota = df.groupby("city", as_index=False)["revenue"].sum().nlargest(10, "revenue")
        df_kota["revenue_miliar"] = df_kota["revenue"] / 1_000_000_000
        
        fig = px.pie(
            df_kota,
            values="revenue_miliar",
            names="city",
            hole=0.4,
            color_discrete_sequence=px.colors.sequential.RdBu
        )
        
        fig.update_traces(
            textposition="inside",
            textinfo="percent+label",
            hovertemplate="<b>%{label}</b><br>Pendapatan: Rp%{customdata}<extra></extra>",
            customdata=df_kota["revenue"].apply(format_angka_tanpa_rp)
        )
        
        fig.update_layout(
            title="Kontribusi 10 Kota Teratas",
            height=400,
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5)
        )
        
        st.plotly_chart(fig, use_container_width=True)
    
    with col_d:
        st.subheader("📊 Performa Kategori Produk")
        if "category" in df.columns:
            df_kategori = df.groupby("category", as_index=False).agg({
                "revenue": "sum",
                "profit": "sum",
                "sales_qty": "sum"
            })
            
            fig = make_subplots(rows=1, cols=2, subplot_titles=("Pendapatan", "Keuntungan"))
            
            # Pendapatan per kategori
            fig.add_trace(
                go.Bar(
                    x=df_kategori["category"],
                    y=df_kategori["revenue"] / 1_000_000_000,
                    name="Pendapatan",
                    marker_color="#2ca02c"
                ),
                row=1, col=1
            )
            
            # Keuntungan per kategori
            fig.add_trace(
                go.Bar(
                    x=df_kategori["category"],
                    y=df_kategori["profit"] / 1_000_000_000,
                    name="Keuntungan",
                    marker_color="#d62728"
                ),
                row=1, col=2
            )
            
            fig.update_layout(
                height=400,
                showlegend=False,
                xaxis_title="Kategori",
                yaxis_title="Pendapatan (Miliar Rupiah)",
                xaxis2_title="Kategori",
                yaxis2_title="Keuntungan (Miliar Rupiah)"
            )
            
            st.plotly_chart(fig, use_container_width=True)
    
    # Insight Otomatis
    st.markdown("---")
    st.subheader("🔍 Insight Analisis Otomatis")
    
    # Hitung insight
    bulan_tertinggi = df_bulanan.loc[df_bulanan["revenue"].idxmax(), "nama_bulan"] if not df_bulanan.empty else "Tidak ada data"
    bulan_terendah = df_bulanan.loc[df_bulanan["revenue"].idxmin(), "nama_bulan"] if not df_bulanan.empty else "Tidak ada data"
    produk_terlaris = df_produk.iloc[0]["product_name"] if not df_produk.empty else "Tidak ada data"
    kota_tertinggi = df_kota.iloc[0]["city"] if not df_kota.empty else "Tidak ada data"
    margin_rata = df["profit_margin"].mean() if "profit_margin" in df.columns else 0
    
    col_insight1, col_insight2 = st.columns(2)
    
    with col_insight1:
        st.info("📈 **Perform Tertinggi**")
        st.markdown(f"""
        - **Bulan Terbaik**: {bulan_tertinggi}
        - **Produk Unggulan**: {produk_terlaris}
        - **Kota Teratas**: {kota_tertinggi}
        - **Margin Rata-rata**: {margin_rata:.1f}%
        """)
    
    with col_insight2:
        st.warning("📉 **Area Perbaikan**")
        st.markdown(f"""
        - **Bulan Terendah**: {bulan_terendah}
        - **Pertumbuhan MoM**: {pertumbuhan:.1f}%
        - **Jumlah Kota**: {kota_aktif} dari target
        - **Transaksi/Hari**: {(len(df) / df['date'].nunique()):.0f}
        """)

# =====================================================
# TREN PENDAPATAN
# =====================================================
elif menu == "📈 Tren Pendapatan":
    st.title("📈 Analisis Tren Pendapatan")
    st.markdown("---")
    
    tab1, tab2, tab3 = st.tabs(["📅 Harian", "📆 Bulanan", "📊 Mingguan"])
    
    with tab1:
        st.subheader("Tren Pendapatan Harian")
        df_harian = df.groupby("date", as_index=False).agg({
            "revenue": "sum",
            "profit": "sum",
            "sales_qty": "sum"
        }).sort_values("date")
        
        df_harian["revenue_miliar"] = df_harian["revenue"] / 1_000_000_000
        
        fig = px.line(
            df_harian,
            x="date",
            y="revenue_miliar",
            markers=True,
            line_shape="spline"
        )
        
        fig.update_layout(
            title="Pendapatan Harian",
            xaxis_title="Tanggal",
            yaxis_title="Pendapatan (Miliar Rupiah)",
            hovermode="x unified"
        )
        
        fig = tambahkan_hover_uang(fig, df_harian, "revenue", "line")
        st.plotly_chart(fig, use_container_width=True)
        
        # Statistik harian
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Rata-rata Harian", format_angka_otomatis(df_harian["revenue"].mean()))
        with col2:
            st.metric("Maksimum Harian", format_angka_otomatis(df_harian["revenue"].max()))
        with col3:
            st.metric("Minimum Harian", format_angka_otomatis(df_harian["revenue"].min()))
    
    with tab2:
        st.subheader("Tren Pendapatan Bulanan")
        df_bulanan = df.groupby(["bulan", "nama_bulan"], as_index=False).agg({
            "revenue": "sum",
            "profit": "sum",
            "sales_qty": "sum",
            "profit_margin": "mean"
        }).sort_values("bulan")
        
        fig = make_subplots(
            rows=2, cols=1,
            subplot_titles=("Pendapatan Bulanan", "Margin Keuntungan"),
            vertical_spacing=0.15
        )
        
        # Pendapatan
        fig.add_trace(
            go.Bar(
                x=df_bulanan["nama_bulan"],
                y=df_bulanan["revenue"] / 1_000_000_000,
                name="Pendapatan",
                marker_color="#1f77b4"
            ),
            row=1, col=1
        )
        
        # Margin
        fig.add_trace(
            go.Scatter(
                x=df_bulanan["nama_bulan"],
                y=df_bulanan["profit_margin"],
                name="Margin",
                mode="lines+markers",
                line=dict(color="#d62728", width=3),
                marker=dict(size=8)
            ),
            row=2, col=1
        )
        
        fig.update_layout(
            height=600,
            showlegend=True,
            xaxis_title="Bulan",
            yaxis_title="Pendapatan (Miliar Rupiah)",
            yaxis2_title="Margin (%)"
        )
        
        st.plotly_chart(fig, use_container_width=True)
    
    with tab3:
        st.subheader("Tren Pendapatan Mingguan")
        df_mingguan = df.groupby("minggu", as_index=False).agg({
            "revenue": "sum",
            "profit": "sum"
        }).sort_values("minggu")
        
        col1, col2 = st.columns(2)
        
        with col1:
            fig = px.bar(
                df_mingguan,
                x="minggu",
                y="revenue",
                title="Pendapatan Mingguan"
            )
            fig.update_layout(yaxis_title="Pendapatan (Rupiah)")
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            fig = px.line(
                df_mingguan,
                x="minggu",
                y="profit",
                title="Keuntungan Mingguan",
                markers=True
            )
            fig.update_layout(yaxis_title="Keuntungan (Rupiah)")
            st.plotly_chart(fig, use_container_width=True)

# =====================================================
# PERFORMANSI PRODUK
# =====================================================
elif menu == "📊 Performa Produk":
    st.title("📊 Analisis Performa Produk")
    st.markdown("---")
    
    # Filter jumlah produk
    jumlah_produk = st.slider("Jumlah produk teratas yang ditampilkan:", 5, 20, 10)
    
    tab1, tab2, tab3 = st.tabs(["📈 Pendapatan", "📦 Volume", "💰 Profitabilitas"])
    
    with tab1:
        st.subheader(f"Top {jumlah_produk} Produk Berdasarkan Pendapatan")
        df_produk = df.groupby("product_name", as_index=False).agg({
            "revenue": "sum",
            "sales_qty": "sum",
            "profit": "sum"
        }).nlargest(jumlah_produk, "revenue")
        
        fig = px.bar(
            df_produk,
            x="revenue",
            y="product_name",
            orientation="h",
            color="revenue",
            color_continuous_scale="Viridis",
            text=df_produk["revenue"].apply(lambda x: f"Rp{format_angka_tanpa_rp(x)}")
        )
        
        fig.update_layout(
            title=f"Top {jumlah_produk} Produk - Pendapatan",
            xaxis_title="Pendapatan (Rupiah)",
            yaxis_title="Produk",
            height=500
        )
        
        fig = tambahkan_hover_uang(fig, df_produk, "revenue", "hbar")
        st.plotly_chart(fig, use_container_width=True)
    
    with tab2:
        st.subheader(f"Top {jumlah_produk} Produk Berdasarkan Volume Penjualan")
        df_volume = df.groupby("product_name", as_index=False).agg({
            "sales_qty": "sum",
            "revenue": "sum",
            "unit_price": "mean"
        }).nlargest(jumlah_produk, "sales_qty")
        
        fig = px.bar(
            df_volume,
            x="sales_qty",
            y="product_name",
            orientation="h",
            color="sales_qty",
            color_continuous_scale="Plasma",
            text="sales_qty"
        )
        
        fig.update_layout(
            title=f"Top {jumlah_produk} Produk - Volume Penjualan",
            xaxis_title="Jumlah Unit Terjual",
            yaxis_title="Produk",
            height=500
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        # Informasi tambahan
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Total Unit Terjual", f"{df_volume['sales_qty'].sum():,.0f}".replace(",", "."))
        with col2:
            st.metric("Rata-rata Harga Unit", format_angka_otomatis(df_volume['unit_price'].mean()))
    
    with tab3:
        st.subheader(f"Top {jumlah_produk} Produk Berdasarkan Keuntungan")
        df_profit = df.groupby("product_name", as_index=False).agg({
            "profit": "sum",
            "revenue": "sum",
            "profit_margin": "mean"
        }).nlargest(jumlah_produk, "profit")
        
        # Scatter plot profit vs revenue
        fig = px.scatter(
            df_profit,
            x="revenue",
            y="profit",
            size="profit_margin",
            color="profit_margin",
            hover_name="product_name",
            size_max=50,
            color_continuous_scale="RdYlGn"
        )
        
        fig.update_layout(
            title="Profitabilitas Produk",
            xaxis_title="Pendapatan (Rupiah)",
            yaxis_title="Keuntungan (Rupiah)",
            height=500
        )
        
        # Format hover
        fig.update_traces(
            customdata=df_profit[["revenue", "profit", "profit_margin"]].applymap(format_angka_tanpa_rp),
            hovertemplate=(
                "<b>%{hovertext}</b><br>"
                "Pendapatan: Rp%{customdata[0]}<br>"
                "Keuntungan: Rp%{customdata[1]}<br>"
                "Margin: %{customdata[2]}%"
                "<extra></extra>"
            )
        )
        
        st.plotly_chart(fig, use_container_width=True)

# =====================================================
# PERFORMANSI KOTA
# =====================================================
elif menu == "🏙️ Performa Kota":
    st.title("🏙️ Analisis Performa Kota")
    st.markdown("---")
    
    tab1, tab2, tab3 = st.tabs(["📊 Pendapatan", "📈 Pertumbuhan", "🗺️ Geografis"])
    
    with tab1:
        st.subheader("Distribusi Pendapatan per Kota")
        
        jumlah_kota = st.slider("Jumlah kota teratas:", 5, 30, 15)
        df_kota = df.groupby("city", as_index=False).agg({
            "revenue": "sum",
            "profit": "sum",
            "sales_qty": "sum",
            "profit_margin": "mean"
        }).nlargest(jumlah_kota, "revenue")
        
        fig = px.bar(
            df_kota,
            y="city",
            x="revenue",
            orientation="h",
            color="profit_margin",
            color_continuous_scale="RdYlGn",
            text=df_kota["revenue"].apply(lambda x: f"Rp{format_angka_tanpa_rp(x)}")
        )
        
        fig.update_layout(
            title=f"Top {jumlah_kota} Kota Berdasarkan Pendapatan",
            xaxis_title="Pendapatan (Rupiah)",
            yaxis_title="Kota",
            height=500 + (jumlah_kota * 10)
        )
        
        st.plotly_chart(fig, use_container_width=True)
    
    with tab2:
        st.subheader("Pertumbuhan Pendapatan per Kota")
        
        if "bulan" in df.columns:
            # Hitung pertumbuhan bulanan per kota
            df_kota_bulan = df.groupby(["city", "bulan"], as_index=False)["revenue"].sum()
            pivot_kota = df_kota_bulan.pivot(index="city", columns="bulan", values="revenue").fillna(0)
            
            # Hitung pertumbuhan bulan terakhir
            if pivot_kota.shape[1] > 1:
                pivot_kota["pertumbuhan"] = (pivot_kota.iloc[:, -1] - pivot_kota.iloc[:, -2]) / pivot_kota.iloc[:, -2] * 100
                
                df_pertumbuhan = pivot_kota["pertumbuhan"].reset_index()
                df_pertumbuhan = df_pertumbuhan.sort_values("pertumbuhan", ascending=False)
                
                fig = px.bar(
                    df_pertumbuhan.head(15),
                    x="pertumbuhan",
                    y="city",
                    orientation="h",
                    color="pertumbuhan",
                    color_continuous_scale="RdYlGn",
                    text=df_pertumbuhan["pertumbuhan"].apply(lambda x: f"{x:.1f}%")
                )
                
                fig.update_layout(
                    title="Pertumbuhan Pendapatan Bulanan",
                    xaxis_title="Pertumbuhan (%)",
                    yaxis_title="Kota",
                    height=500
                )
                
                st.plotly_chart(fig, use_container_width=True)
    
    with tab3:
        st.subheader("Visualisasi Geografis Pendapatan")
        
        # Data contoh untuk peta (biasanya akan menggunakan koordinat GPS)
        # Di sini kita gunakan chart alternatif karena tidak ada data koordinat
        
        df_kota_map = df.groupby("city", as_index=False).agg({
            "revenue": "sum",
            "profit": "sum"
        }).nlargest(20, "revenue")
        
        # Bubble chart sebagai alternatif peta
        fig = px.scatter(
            df_kota_map,
            x="revenue",
            y="profit",
            size="revenue",
            color="profit",
            hover_name="city",
            size_max=60,
            color_continuous_scale="Rainbow"
        )
        
        fig.update_layout(
            title="Distribusi Pendapatan vs Keuntungan per Kota",
            xaxis_title="Pendapatan (Rupiah)",
            yaxis_title="Keuntungan (Rupiah)",
            height=600
        )
        
        st.plotly_chart(fig, use_container_width=True)

# =====================================================
# ANALISIS KATEGORI
# =====================================================
elif menu == "📦 Analisis Kategori":
    st.title("📦 Analisis Kategori Produk")
    st.markdown("---")
    
    if "category" in df.columns:
        tab1, tab2, tab3 = st.tabs(["📊 Overview", "📈 Tren", "📊 Perbandingan"])
        
        with tab1:
            st.subheader("Performa Kategori Produk")
            
            df_kategori = df.groupby("category", as_index=False).agg({
                "revenue": "sum",
                "profit": "sum",
                "sales_qty": "sum",
                "profit_margin": "mean"
            }).sort_values("revenue", ascending=False)
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Pie chart kontribusi pendapatan
                fig = px.pie(
                    df_kategori,
                    values="revenue",
                    names="category",
                    title="Kontribusi Pendapatan per Kategori",
                    hole=0.3,
                    color_discrete_sequence=px.colors.qualitative.Set3
                )
                
                fig.update_traces(
                    textposition="inside",
                    textinfo="percent+label",
                    hovertemplate="<b>%{label}</b><br>Pendapatan: Rp%{customdata}<extra></extra>",
                    customdata=df_kategori["revenue"].apply(format_angka_tanpa_rp)
                )
                
                st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                # Bar chart margin
                fig = px.bar(
                    df_kategori,
                    x="category",
                    y="profit_margin",
                    title="Margin Keuntungan per Kategori",
                    color="profit_margin",
                    color_continuous_scale="RdYlGn",
                    text=df_kategori["profit_margin"].apply(lambda x: f"{x:.1f}%")
                )
                
                fig.update_layout(
                    xaxis_title="Kategori",
                    yaxis_title="Margin (%)",
                    height=400
                )
                
                st.plotly_chart(fig, use_container_width=True)
        
        with tab2:
            st.subheader("Tren Kategori per Bulan")
            
            df_kategori_bulan = df.groupby(["category", "bulan"], as_index=False).agg({
                "revenue": "sum",
                "profit": "sum"
            })
            
            fig = px.line(
                df_kategori_bulan,
                x="bulan",
                y="revenue",
                color="category",
                markers=True,
                line_shape="spline"
            )
            
            fig.update_layout(
                title="Tren Pendapatan Kategori per Bulan",
                xaxis_title="Bulan",
                yaxis_title="Pendapatan (Rupiah)",
                height=500
            )
            
            st.plotly_chart(fig, use_container_width=True)
        
        with tab3:
            st.subheader("Perbandingan Kategori")
            
            kategori_pilihan = st.multiselect(
                "Pilih kategori untuk dibandingkan:",
                options=df["category"].unique(),
                default=df["category"].unique()[:3] if len(df["category"].unique()) >= 3 else df["category"].unique()
            )
            
            if kategori_pilihan:
                df_filtered = df[df["category"].isin(kategori_pilihan)]
                df_perbandingan = df_filtered.groupby("category", as_index=False).agg({
                    "revenue": "sum",
                    "profit": "sum",
                    "sales_qty": "sum",
                    "unit_price": "mean"
                })
                
                fig = go.Figure()
                
                # Tambahkan trace untuk setiap metrik
                fig.add_trace(go.Bar(
                    name="Pendapatan",
                    x=df_perbandingan["category"],
                    y=df_perbandingan["revenue"] / 1_000_000_000,
                    marker_color="#1f77b4"
                ))
                
                fig.add_trace(go.Bar(
                    name="Keuntungan",
                    x=df_perbandingan["category"],
                    y=df_perbandingan["profit"] / 1_000_000_000,
                    marker_color="#ff7f0e"
                ))
                
                fig.update_layout(
                    title="Perbandingan Kategori",
                    xaxis_title="Kategori",
                    yaxis_title="Nilai (Miliar Rupiah)",
                    barmode="group",
                    height=500
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
                # Tabel perbandingan
                st.subheader("Tabel Perbandingan")
                df_tabel = df_perbandingan.copy()
                df_tabel["revenue"] = df_tabel["revenue"].apply(format_angka_otomatis)
                df_tabel["profit"] = df_tabel["profit"].apply(format_angka_otomatis)
                df_tabel["unit_price"] = df_tabel["unit_price"].apply(format_angka_otomatis)
                
                st.dataframe(
                    df_tabel.style.format({
                        "sales_qty": "{:,.0f}"
                    }),
                    use_container_width=True
                )

# =====================================================
# ANALISIS CHANNEL
# =====================================================
elif menu == "🛒 Analisis Channel":
    st.title("🛒 Analisis Channel Penjualan")
    st.markdown("---")
    
    if "channel" in df.columns:
        tab1, tab2, tab3 = st.tabs(["📊 Distribusi", "📈 Performa", "📱 Customer Insights"])
        
        with tab1:
            st.subheader("Distribusi Channel")
            
            df_channel = df.groupby("channel", as_index=False).agg({
                "revenue": "sum",
                "profit": "sum",
                "sales_qty": "sum"
            })
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Donut chart
                fig = px.pie(
                    df_channel,
                    values="revenue",
                    names="channel",
                    title="Distribusi Pendapatan per Channel",
                    hole=0.4,
                    color_discrete_sequence=px.colors.sequential.Rainbow
                )
                
                fig.update_traces(
                    textposition="inside",
                    textinfo="percent+label"
                )
                
                st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                # Bar chart perbandingan
                fig = px.bar(
                    df_channel,
                    x="channel",
                    y=["revenue", "profit"],
                    title="Perbandingan Pendapatan vs Keuntungan",
                    barmode="group"
                )
                
                fig.update_layout(
                    xaxis_title="Channel",
                    yaxis_title="Nilai (Rupiah)",
                    height=400
                )
                
                st.plotly_chart(fig, use_container_width=True)
        
        with tab2:
            st.subheader("Performa Channel per Bulan")
            
            df_channel_bulan = df.groupby(["channel", "bulan"], as_index=False).agg({
                "revenue": "sum",
                "profit_margin": "mean"
            })
            
            # Pilih channel untuk dianalisis
            channel_pilihan = st.multiselect(
                "Pilih channel:",
                options=df["channel"].unique(),
                default=df["channel"].unique()[:3] if len(df["channel"].unique()) >= 3 else df["channel"].unique()
            )
            
            if channel_pilihan:
                df_filtered = df_channel_bulan[df_channel_bulan["channel"].isin(channel_pilihan)]
                
                col1, col2 = st.columns(2)
                
                with col1:
                    # Line chart revenue
                    fig = px.line(
                        df_filtered,
                        x="bulan",
                        y="revenue",
                        color="channel",
                        markers=True,
                        title="Tren Pendapatan per Channel"
                    )
                    
                    fig.update_layout(
                        xaxis_title="Bulan",
                        yaxis_title="Pendapatan (Rupiah)"
                    )
                    
                    st.plotly_chart(fig, use_container_width=True)
                
                with col2:
                    # Line chart margin
                    fig = px.line(
                        df_filtered,
                        x="bulan",
                        y="profit_margin",
                        color="channel",
                        markers=True,
                        title="Tren Margin per Channel"
                    )
                    
                    fig.update_layout(
                        xaxis_title="Bulan",
                        yaxis_title="Margin (%)"
                    )
                    
                    st.plotly_chart(fig, use_container_width=True)
        
        with tab3:
            st.subheader("Analisis Customer per Channel")
            
            if "customer_type" in df.columns:
                df_customer_channel = df.groupby(["channel", "customer_type"], as_index=False).agg({
                    "revenue": "sum",
                    "sales_qty": "sum"
                })
                
                fig = px.sunburst(
                    df_customer_channel,
                    path=["channel", "customer_type"],
                    values="revenue",
                    title="Distribusi Customer per Channel",
                    color="revenue",
                    color_continuous_scale="RdBu"
                )
                
                fig.update_layout(height=600)
                st.plotly_chart(fig, use_container_width=True)

# =====================================================
# ANALISIS PROFITABILITAS
# =====================================================
elif menu == "💰 Analisis Profitabilitas":
    st.title("💰 Analisis Profitabilitas")
    st.markdown("---")
    
    tab1, tab2, tab3 = st.tabs(["📊 Margin", "📈 Cost Analysis", "📉 Profit Drivers"])
    
    with tab1:
        st.subheader("Analisis Margin Keuntungan")
        
        # Histogram margin
        df_margin = df.dropna(subset=["profit_margin"])
        
        fig = px.histogram(
            df_margin,
            x="profit_margin",
            nbins=30,
            title="Distribusi Margin Keuntungan",
            color_discrete_sequence=["#2ca02c"]
        )
        
        fig.update_layout(
            xaxis_title="Margin Keuntungan (%)",
            yaxis_title="Frekuensi",
            height=400
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        # Statistik margin
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Rata-rata Margin", f"{df_margin['profit_margin'].mean():.1f}%")
        with col2:
            st.metric("Median Margin", f"{df_margin['profit_margin'].median():.1f}%")
        with col3:
            st.metric("Margin Maks", f"{df_margin['profit_margin'].max():.1f}%")
        with col4:
            st.metric("Margin Min", f"{df_margin['profit_margin'].min():.1f}%")
    
    with tab2:
        st.subheader("Analisis Biaya vs Pendapatan")
        
        df_scatter = df.dropna(subset=["cost", "revenue", "profit_margin"])
        
        fig = px.scatter(
            df_scatter,
            x="cost",
            y="revenue",
            color="profit_margin",
            size="sales_qty",
            hover_name="product_name",
            color_continuous_scale="RdYlGn",
            title="Hubungan Biaya vs Pendapatan"
        )
        
        fig.update_layout(
            xaxis_title="Biaya (Rupiah)",
            yaxis_title="Pendapatan (Rupiah)",
            height=500
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        # Regresi linier
        st.subheader("Hubungan Linear Biaya-Pendapatan")
        
        # Hitung koefisien korelasi
        correlation = df_scatter["cost"].corr(df_scatter["revenue"])
        
        st.info(f"**Koefisien Korelasi:** {correlation:.3f}")
        if correlation > 0.7:
            st.success("✅ Korelasi kuat positif: Biaya yang lebih tinggi cenderung menghasilkan pendapatan yang lebih tinggi")
        elif correlation > 0.3:
            st.warning("⚠️ Korelasi moderat: Ada hubungan antara biaya dan pendapatan")
        else:
            st.error("❌ Korelasi lemah: Biaya tidak berkorelasi kuat dengan pendapatan")
    
    with tab3:
        st.subheader("Driver Keuntungan")
        
        # Analisis faktor yang mempengaruhi profit
        col1, col2 = st.columns(2)
        
        with col1:
            # By category
            if "category" in df.columns:
                df_profit_category = df.groupby("category", as_index=False).agg({
                    "profit": "sum",
                    "profit_margin": "mean"
                }).sort_values("profit", ascending=False)
                
                fig = px.bar(
                    df_profit_category,
                    x="category",
                    y="profit",
                    title="Keuntungan per Kategori",
                    color="profit_margin",
                    color_continuous_scale="Viridis"
                )
                
                st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            # By city
            if "city" in df.columns:
                df_profit_city = df.groupby("city", as_index=False).agg({
                    "profit": "sum",
                    "profit_margin": "mean"
                }).nlargest(10, "profit")
                
                fig = px.bar(
                    df_profit_city,
                    x="city",
                    y="profit",
                    title="Top 10 Kota Berdasarkan Keuntungan",
                    color="profit_margin",
                    color_continuous_scale="Plasma"
                )
                
                st.plotly_chart(fig, use_container_width=True)

# =====================================================
# ANALISIS DISKONTING
# =====================================================
elif menu == "📉 Analisis Diskonting":
    st.title("📉 Analisis Dampak Diskon")
    st.markdown("---")
    
    if "discount" in df.columns:
        tab1, tab2, tab3 = st.tabs(["📊 Overview", "📈 Impact", "📉 Optimization"])
        
        with tab1:
            st.subheader("Distribusi Diskon")
            
            df_diskon = df.dropna(subset=["discount"])
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Histogram diskon
                fig = px.histogram(
                    df_diskon,
                    x="discount",
                    nbins=20,
                    title="Distribusi Tingkat Diskon",
                    color_discrete_sequence=["#d62728"]
                )
                
                fig.update_layout(
                    xaxis_title="Diskon (%)",
                    yaxis_title="Jumlah Transaksi",
                    height=400
                )
                
                st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                # Box plot diskon per kategori
                if "category" in df.columns:
                    fig = px.box(
                        df_diskon,
                        x="category",
                        y="discount",
                        title="Distribusi Diskon per Kategori",
                        color="category"
                    )
                    
                    fig.update_layout(
                        xaxis_title="Kategori",
                        yaxis_title="Diskon (%)",
                        height=400,
                        showlegend=False
                    )
                    
                    st.plotly_chart(fig, use_container_width=True)
        
        with tab2:
            st.subheader("Dampak Diskon terhadap Penjualan")
            
            # Scatter plot diskon vs volume
            fig = px.scatter(
                df_diskon,
                x="discount",
                y="sales_qty",
                color="revenue",
                size="revenue",
                hover_name="product_name",
                title="Hubungan Diskon vs Volume Penjualan",
                trendline="ols"
            )
            
            fig.update_layout(
                xaxis_title="Diskon (%)",
                yaxis_title="Volume Penjualan (Unit)",
                height=500
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            # Analisis efektivitas diskon
            st.subheader("Analisis Efektivitas Diskon")
            
            # Group by discount range
            df_diskon["discount_range"] = pd.cut(
                df_diskon["discount"],
                bins=[0, 5, 10, 15, 20, 30, 50, 100],
                labels=["0-5%", "5-10%", "10-15%", "15-20%", "20-30%", "30-50%", "50+%"]
            )
            
            df_effectiveness = df_diskon.groupby("discount_range", as_index=False).agg({
                "sales_qty": "mean",
                "revenue": "mean",
                "profit_margin": "mean"
            })
            
            fig = make_subplots(
                rows=2, cols=1,
                subplot_titles=("Rata-rata Volume Penjualan", "Rata-rata Margin"),
                vertical_spacing=0.15
            )
            
            fig.add_trace(
                go.Bar(
                    x=df_effectiveness["discount_range"],
                    y=df_effectiveness["sales_qty"],
                    name="Volume Penjualan",
                    marker_color="#1f77b4"
                ),
                row=1, col=1
            )
            
            fig.add_trace(
                go.Scatter(
                    x=df_effectiveness["discount_range"],
                    y=df_effectiveness["profit_margin"],
                    name="Margin",
                    mode="lines+markers",
                    line=dict(color="#d62728", width=3),
                    marker=dict(size=8)
                ),
                row=2, col=1
            )
            
            fig.update_layout(
                height=600,
                showlegend=True,
                xaxis_title="Rentang Diskon",
                yaxis_title="Volume Rata-rata",
                yaxis2_title="Margin Rata-rata (%)"
            )
            
            st.plotly_chart(fig, use_container_width=True)
        
        with tab3:
            st.subheader("Optimisasi Strategi Diskon")
            
            # Rekomendasi diskon optimal
            st.info("""
            **Rekomendasi Strategi Diskon:**
            
            1. **Diskon Rendah (0-10%):** Ideal untuk produk dengan permintaan elastis
            2. **Diskon Sedang (10-20%):** Efektif untuk meningkatkan volume tanpa mengurangi margin signifikan
            3. **Diskon Tinggi (>20%):** Hanya untuk produk dengan stok berlebih atau musiman
            """)
            
            # Analisis produk dengan diskon
            df_produk_diskon = df_diskon.groupby("product_name", as_index=False).agg({
                "discount": "mean",
                "sales_qty": "sum",
                "profit_margin": "mean"
            })
            
            # Identifikasi produk dengan diskon tinggi tapi margin rendah
            produk_inefisien = df_produk_diskon[
                (df_produk_diskon["discount"] > 15) & 
                (df_produk_diskon["profit_margin"] < 10)
            ]
            
            if not produk_inefisien.empty:
                st.warning("⚠️ **Produk dengan Diskon Tinggi dan Margin Rendah:**")
                st.dataframe(
                    produk_inefisien[["product_name", "discount", "profit_margin"]]
                    .sort_values("discount", ascending=False)
                    .head(10),
                    use_container_width=True
                )
            
            # Identifikasi produk dengan potensi diskon
            produk_potensial = df_produk_diskon[
                (df_produk_diskon["discount"] < 5) & 
                (df_produk_diskon["profit_margin"] > 20)
            ]
            
            if not produk_potensial.empty:
                st.success("✅ **Produk dengan Potensi Diskon:**")
                st.dataframe(
                    produk_potensial[["product_name", "discount", "profit_margin", "sales_qty"]]
                    .sort_values("profit_margin", ascending=False)
                    .head(10),
                    use_container_width=True
                )

# =====================================================
# ANALISIS WAKTU
# =====================================================
elif menu == "📅 Analisis Waktu":
    st.title("📅 Analisis Temporal")
    st.markdown("---")
    
    tab1, tab2, tab3 = st.tabs(["📅 Musiman", "📆 Hari", "⏰ Jam"])
    
    with tab1:
        st.subheader("Analisis Musiman")
        
        if "bulan" in df.columns:
            df_musiman = df.groupby("bulan", as_index=False).agg({
                "revenue": "sum",
                "profit": "sum",
                "sales_qty": "sum"
            })
            
            fig = make_subplots(
                rows=3, cols=1,
                subplot_titles=("Pendapatan", "Keuntungan", "Volume Penjualan"),
                vertical_spacing=0.2
            )
            
            fig.add_trace(
                go.Scatter(
                    x=df_musiman["bulan"],
                    y=df_musiman["revenue"],
                    mode="lines+markers",
                    name="Pendapatan",
                    line=dict(color="#1f77b4", width=3)
                ),
                row=1, col=1
            )
            
            fig.add_trace(
                go.Scatter(
                    x=df_musiman["bulan"],
                    y=df_musiman["profit"],
                    mode="lines+markers",
                    name="Keuntungan",
                    line=dict(color="#2ca02c", width=3)
                ),
                row=2, col=1
            )
            
            fig.add_trace(
                go.Bar(
                    x=df_musiman["bulan"],
                    y=df_musiman["sales_qty"],
                    name="Volume",
                    marker_color="#ff7f0e"
                ),
                row=3, col=1
            )
            
            fig.update_layout(
                height=700,
                showlegend=True,
                xaxis_title="Bulan",
                yaxis_title="Pendapatan (Rupiah)",
                yaxis2_title="Keuntungan (Rupiah)",
                yaxis3_title="Volume (Unit)"
            )
            
            st.plotly_chart(fig, use_container_width=True)
    
    with tab2:
        st.subheader("Analisis Performa Hari")
        
        if "date" in df.columns:
            df["hari_dalam_minggu"] = df["date"].dt.day_name()
            df["hari_dalam_minggu_num"] = df["date"].dt.dayofweek
            
            df_hari = df.groupby(["hari_dalam_minggu", "hari_dalam_minggu_num"], as_index=False).agg({
                "revenue": "mean",
                "sales_qty": "mean"
            }).sort_values("hari_dalam_minggu_num")
            
            fig = px.bar(
                df_hari,
                x="hari_dalam_minggu",
                y="revenue",
                title="Rata-rata Pendapatan per Hari",
                color="revenue",
                color_continuous_scale="Blues",
                text=df_hari["revenue"].apply(lambda x: f"Rp{format_angka_tanpa_rp(x)}")
            )
            
            fig.update_layout(
                xaxis_title="Hari",
                yaxis_title="Rata-rata Pendapatan (Rupiah)",
                height=400
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            # Heatmap hari vs bulan
            st.subheader("Heatmap: Hari vs Performa")
            
            df_hari_bulan = df.groupby(["hari_dalam_minggu_num", "bulan"], as_index=False)["revenue"].mean()
            pivot = df_hari_bulan.pivot(index="hari_dalam_minggu_num", columns="bulan", values="revenue")
            
            # Map hari angka ke nama
            hari_mapping = {
                0: "Senin", 1: "Selasa", 2: "Rabu", 3: "Kamis",
                4: "Jumat", 5: "Sabtu", 6: "Minggu"
            }
            pivot.index = pivot.index.map(hari_mapping)
            
            fig = px.imshow(
                pivot,
                labels=dict(x="Bulan", y="Hari", color="Pendapatan"),
                color_continuous_scale="Viridis",
                aspect="auto"
            )
            
            fig.update_layout(
                title="Heatmap Rata-rata Pendapatan: Hari vs Bulan",
                height=400
            )
            
            st.plotly_chart(fig, use_container_width=True)
    
    with tab3:
        st.subheader("Analisis Pola Waktu")
        
        # Jika ada data waktu spesifik
        st.info("""
        **Insight Temporal:**
        
        - **Akhir Bulan:** Biasanya terjadi peningkatan penjualan
        - **Hari Kerja vs Akhir Pekan:** Pola berbeda berdasarkan produk
        - **Musim:** Produk tertentu memiliki pola musiman
        
        Gunakan filter tanggal di sidebar untuk analisis lebih detail.
        """)
        
        # Time series decomposition (simplified)
        if "date" in df.columns and "revenue" in df.columns:
            df_timeseries = df.groupby("date")["revenue"].sum().reset_index()
            df_timeseries = df_timeseries.set_index("date").asfreq("D").fillna(0)
            
            # Plot rolling average
            df_timeseries["rolling_7"] = df_timeseries["revenue"].rolling(window=7).mean()
            df_timeseries["rolling_30"] = df_timeseries["revenue"].rolling(window=30).mean()
            
            fig = go.Figure()
            
            fig.add_trace(go.Scatter(
                x=df_timeseries.index,
                y=df_timeseries["revenue"],
                name="Aktual",
                line=dict(color="lightgray", width=1),
                opacity=0.5
            ))
            
            fig.add_trace(go.Scatter(
                x=df_timeseries.index,
                y=df_timeseries["rolling_7"],
                name="Rata-rata 7 Hari",
                line=dict(color="#1f77b4", width=2)
            ))
            
            fig.add_trace(go.Scatter(
                x=df_timeseries.index,
                y=df_timeseries["rolling_30"],
                name="Rata-rata 30 Hari",
                line=dict(color="#d62728", width=3)
            ))
            
            fig.update_layout(
                title="Tren Pendapatan dengan Moving Average",
                xaxis_title="Tanggal",
                yaxis_title="Pendapatan (Rupiah)",
                height=500,
                hovermode="x unified"
            )
            
            st.plotly_chart(fig, use_container_width=True)

# =====================================================
# ANALISIS PELANGGAN
# =====================================================
elif menu == "📱 Analisis Pelanggan":
    st.title("📱 Analisis Segmentasi Pelanggan")
    st.markdown("---")
    
    if "customer_type" in df.columns:
        tab1, tab2 = st.tabs(["👥 Segmentasi", "📊 Value"])
        
        with tab1:
            st.subheader("Segmentasi Pelanggan")
            
            df_customer = df.groupby("customer_type", as_index=False).agg({
                "revenue": "sum",
                "profit": "sum",
                "sales_qty": "sum",
                "date": "nunique"  # jumlah hari transaksi
            })
            
            df_customer = df_customer.rename(columns={"date": "hari_aktif"})
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Pie chart distribusi pelanggan
                fig = px.pie(
                    df_customer,
                    values="revenue",
                    names="customer_type",
                    title="Distribusi Pendapatan per Tipe Pelanggan",
                    hole=0.3,
                    color_discrete_sequence=px.colors.qualitative.Pastel
                )
                
                fig.update_traces(
                    textposition="inside",
                    textinfo="percent+label"
                )
                
                st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                # Bar chart perbandingan
                fig = px.bar(
                    df_customer,
                    x="customer_type",
                    y=["revenue", "profit"],
                    title="Perbandingan Pendapatan vs Keuntungan",
                    barmode="group"
                )
                
                fig.update_layout(
                    xaxis_title="Tipe Pelanggan",
                    yaxis_title="Nilai (Rupiah)",
                    height=400
                )
                
                st.plotly_chart(fig, use_container_width=True)
        
        with tab2:
            st.subheader("Customer Value Analysis")
            
            # Hitung metrics per pelanggan
            df_customer_value = df.groupby("customer_type", as_index=False).agg({
                "revenue": ["sum", "mean", "std"],
                "sales_qty": ["sum", "mean"],
                "profit_margin": "mean"
            })
            
            # Flatten multi-index columns
            df_customer_value.columns = [
                "customer_type",
                "total_revenue",
                "avg_transaction",
                "std_transaction",
                "total_quantity",
                "avg_quantity",
                "avg_margin"
            ]
            
            # Visualisasi
            fig = make_subplots(
                rows=2, cols=2,
                subplot_titles=(
                    "Total Pendapatan",
                    "Rata-rata Transaksi",
                    "Rata-rata Kuantitas",
                    "Rata-rata Margin"
                ),
                vertical_spacing=0.15,
                horizontal_spacing=0.15
            )
            
            fig.add_trace(
                go.Bar(
                    x=df_customer_value["customer_type"],
                    y=df_customer_value["total_revenue"] / 1_000_000_000,
                    name="Total Pendapatan",
                    marker_color="#1f77b4"
                ),
                row=1, col=1
            )
            
            fig.add_trace(
                go.Bar(
                    x=df_customer_value["customer_type"],
                    y=df_customer_value["avg_transaction"],
                    name="Rata-rata Transaksi",
                    marker_color="#ff7f0e"
                ),
                row=1, col=2
            )
            
            fig.add_trace(
                go.Bar(
                    x=df_customer_value["customer_type"],
                    y=df_customer_value["avg_quantity"],
                    name="Rata-rata Kuantitas",
                    marker_color="#2ca02c"
                ),
                row=2, col=1
            )
            
            fig.add_trace(
                go.Bar(
                    x=df_customer_value["customer_type"],
                    y=df_customer_value["avg_margin"],
                    name="Rata-rata Margin",
                    marker_color="#d62728"
                ),
                row=2, col=2
            )
            
            fig.update_layout(
                height=600,
                showlegend=False
            )
            
            fig.update_yaxes(title_text="Miliar Rupiah", row=1, col=1)
            fig.update_yaxes(title_text="Rupiah", row=1, col=2)
            fig.update_yaxes(title_text="Unit", row=2, col=1)
            fig.update_yaxes(title_text="%", row=2, col=2)
            
            st.plotly_chart(fig, use_container_width=True)

# =====================================================
# TABEL DATA LENGKAP
# =====================================================
elif menu == "📋 Tabel Data Lengkap":
    st.title("📋 Data Lengkap Transaksi")
    st.markdown("---")
    
    # Tampilkan jumlah baris
    st.write(f"**Total Transaksi:** {len(df):,} baris".replace(",", "."))
    
    # Pilih kolom untuk ditampilkan
    semua_kolom = df.columns.tolist()
    kolom_terpilih = st.multiselect(
        "Pilih kolom yang akan ditampilkan:",
        semua_kolom,
        default=semua_kolom[:10] if len(semua_kolom) > 10 else semua_kolom
    )
    
    # Filter data
    if kolom_terpilih:
        df_tampil = df[kolom_terpilih]
        
        # Pagination
        baris_per_halaman = st.slider("Baris per halaman:", 10, 100, 20)
        total_halaman = len(df_tampil) // baris_per_halaman + 1
        
        halaman = st.number_input(
            "Halaman:",
            min_value=1,
            max_value=total_halaman,
            value=1
        )
        
        mulai = (halaman - 1) * baris_per_halaman
        selesai = mulai + baris_per_halaman
        
        # Tampilkan data
        st.dataframe(
            df_tampil.iloc[mulai:selesai],
            use_container_width=True,
            height=400
        )
        
        # Download button
        csv = df_tampil.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download Data sebagai CSV",
            data=csv,
            file_name="deltech_sumut_2025_filtered.csv",
            mime="text/csv"
        )
        
        # Statistik sederhana
        st.subheader("📊 Statistik Data")
        
        if st.checkbox("Tampilkan statistik deskriptif"):
            st.dataframe(
                df_tampil.describe(),
                use_container_width=True
            )
        
        if st.checkbox("Tampilkan info data types"):
            info_data = pd.DataFrame({
                "Kolom": df_tampil.columns,
                "Tipe Data": df_tampil.dtypes.astype(str),
                "Nilai Unik": df_tampil.nunique(),
                "Nilai Null": df_tampil.isnull().sum()
            })
            st.dataframe(info_data, use_container_width=True)

# =====================================================
# FOOTER
# =====================================================
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: gray;'>
    <p>📊 ITDel Tech | © 2026</p>
    <p>Dashboard ini dibuat untuk simulasi analisis data penjualan tahun 2025</p>
    </div>
    """,
    unsafe_allow_html=True
)

# =====================================================
# STYLE TAMBAHAN
# =====================================================
st.markdown("""
<style>
    .stMetric {
        background-color: #f0f2f6;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #1f77b4;
    }
    
    .stMetric label {
        font-weight: bold;
        color: #2c3e50;
    }
    
    .stMetric div {
        font-size: 1.5rem;
        color: #1f77b4;
    }
    
    .css-1d391kg {
        padding-top: 2rem;
    }
    
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    
    h1, h2, h3 {
        color: #2c3e50;
    }
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
    }
    
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #f0f2f6;
        border-radius: 5px 5px 0px 0px;
        gap: 1px;
        # padding-top: 10px;
        # padding-bottom: 10px;
        padding: 8px 12px;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: #1f77b4;
        color: white;
    }
        
</style>
""", unsafe_allow_html=True)