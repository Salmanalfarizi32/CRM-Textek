import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import os

# --- Path Excel (relative path untuk Streamlit Cloud) ---
file_path = os.path.join(os.getcwd(), "CRM Analyst.xlsx")

# Cek file ada nggak
if not os.path.exists(file_path):
    st.error(f"File CRM Analyst.xlsx tidak ditemukan. Pastikan file ada di repo yang sama dengan CRM.py.")
    st.stop()

# --- Pilih Sheet ---
sheet_options = ["VIP BUYER", "Kategori Buyer", "Marketing Ads", "Pertumbuhan Pelanggan", "Produk Populer", "Produk Favorit Customer"]
current_sheet = st.selectbox("Pilih Sheet CRM", sheet_options)

# =========================
# SHEET 1: VIP BUYER
# =========================
if current_sheet == "VIP BUYER":
    df = pd.read_excel(file_path, sheet_name="VIP BUYER", engine='openpyxl')
    df.columns = df.columns.str.strip()
    
    # Bersihkan Total Transaksi
    df['Total Transaksi'] = df['Total Transaksi'].replace(['-', '', ' '], 0)
    df['Total Transaksi'] = df['Total Transaksi'].replace('[Rp$,]', '', regex=True)
    df['Total Transaksi'] = pd.to_numeric(df['Total Transaksi'], errors='coerce').fillna(0)

    # Session state
    if 'vip_data' not in st.session_state:
        st.session_state.vip_data = df.copy()

    # --- Form Tambah Data ---
    st.subheader("Tambah Data Baru")
    with st.form("add_form"):
        nama = st.text_input("Nama Pelanggan")
        jumlah = st.number_input("Jumlah Transaksi", min_value=0, step=1)
        total = st.number_input("Total Transaksi (Rp)", min_value=0, step=1000)
        submitted = st.form_submit_button("Tambah")
        if submitted:
            new_row = pd.DataFrame({
                'Nama Pelanggan': [nama],
                'Jumlah Transaksi': [jumlah],
                'Total Transaksi': [total]
            })
            st.session_state.vip_data = pd.concat([st.session_state.vip_data, new_row], ignore_index=True)
            st.success(f"{nama} berhasil ditambahkan!")

    # --- Form Hapus Data ---
    st.subheader("Hapus Data")
    if len(st.session_state.vip_data) > 0:
        delete_name = st.selectbox("Pilih Nama Pelanggan yang akan dihapus", st.session_state.vip_data['Nama Pelanggan'])
        if st.button("Hapus"):
            st.session_state.vip_data = st.session_state.vip_data[st.session_state.vip_data['Nama Pelanggan'] != delete_name]
            st.success(f"{delete_name} berhasil dihapus!")

    # --- Tampilkan Tabel VIP BUYER ---
    df_display = st.session_state.vip_data.sort_values(by='Total Transaksi', ascending=False).reset_index(drop=True)
    df_display.index = range(1, len(df_display)+1)

    # Highlight Top 10
    top10_idx = df_display.index[:10]
    def highlight_top10_all(row):
        if row.name in top10_idx[:3]:
            return ['background-color: lightgreen']*len(row)
        elif row.name in top10_idx[3:10]:
            return ['background-color: lightyellow']*len(row)
        else:
            return ['']*len(row)

    st.subheader("Tabel VIP BUYER")
    st.dataframe(df_display.style.format({'Total Transaksi': 'Rp {:,.0f}'}).apply(highlight_top10_all, axis=1))

# =========================
# SHEET 2: KATEGORI BUYER
# =========================
elif current_sheet == "Kategori Buyer":
    df = pd.read_excel(file_path, sheet_name="Kategori Buyer", engine='openpyxl')
    df.columns = df.columns.str.strip()
    df = df[['Nama Customer','Repeat Status','Buyer Status']]
    df.index = range(1, len(df)+1)

    # Highlight Buyer Status
    buyer_color = ['#FFC0CB','#ADD8E6','#90EE90','#FFFFE0','#FFA07A','#D8BFD8','#F0E68C']
    buyer_categories = ['Active Buyer','Cooling Buyer','Dormant Buyer','Inactive Buyer','Lost Buyer','Very Inactive Buyer','Warm Buyer']
    buyer_map = {v: buyer_color[i] for i,v in enumerate(buyer_categories)}
    def highlight_buyer(x):
        return [f'background-color: {buyer_map.get(v,"")}' if col=='Buyer Status' else '' for col,v in zip(x.index,x)]

    st.subheader("Tabel Kategori Buyer")
    st.dataframe(df.style.apply(highlight_buyer, axis=1))

# =========================
# SHEET 3: MARKETING ADS
# =========================
elif current_sheet == "Marketing Ads":
    df_marketing = pd.read_excel(file_path, sheet_name="Marketing Ads", engine='openpyxl')
    df_marketing.index = range(1, len(df_marketing)+1)

    # Pie Chart
    st.subheader("Marketing Ads")
    st.pyplot(df_marketing.plot.pie(y='Jumlah', labels=df_marketing['Channel'], autopct='%1.1f%%').figure)

    # Ringkasan insight simple
    top_channel = df_marketing.loc[df_marketing['Jumlah'].idxmax(), 'Channel']
    st.subheader("Ringkasan Insight")
    st.write(f"Dari data marketing ads, **{top_channel}** adalah yang paling efektif menarik customer karena memiliki persentase tertinggi")

# =========================
# SHEET 4: PERTUMBUHAN PELANGGAN
# =========================
elif current_sheet == "Pertumbuhan Pelanggan":
    df = pd.read_excel(file_path, sheet_name="Pertumbuhan Pelanggan", engine='openpyxl')
    df.columns = df.columns.str.strip()
    df['Jumlah Pelanggan'] = pd.to_numeric(df['Jumlah Pelanggan'], errors='coerce').fillna(0)

    # Hapus baris kosong / None
    df = df.dropna(subset=['Bulan','Jumlah Pelanggan']).reset_index(drop=True)

    # Session state
    if 'growth_data' not in st.session_state:
        st.session_state.growth_data = df.copy()

    # --- Form Tambah Data ---
    st.subheader("Tambah Data Pertumbuhan")
    with st.form("add_growth"):
        bulan = st.text_input("Bulan")
        jumlah = st.number_input("Jumlah Pelanggan", min_value=0, step=1)
        submitted = st.form_submit_button("Tambah")
        if submitted:
            new_row = pd.DataFrame({'Bulan':[bulan],'Jumlah Pelanggan':[jumlah]})
            st.session_state.growth_data = pd.concat([st.session_state.growth_data, new_row], ignore_index=True)
            st.success(f"{bulan} berhasil ditambahkan!")

    # --- Form Hapus Data ---
    st.subheader("Hapus Data Pertumbuhan")
    if len(st.session_state.growth_data) > 0:
        delete_bulan = st.selectbox("Pilih Bulan yang akan dihapus", st.session_state.growth_data['Bulan'])
        if st.button("Hapus Bulan"):
            st.session_state.growth_data = st.session_state.growth_data[st.session_state.growth_data['Bulan'] != delete_bulan]
            st.success(f"{delete_bulan} berhasil dihapus!")

    # --- Tampilkan Tabel ---
    df_display = st.session_state.growth_data.sort_values(by='Jumlah Pelanggan', ascending=False).reset_index(drop=True)
    df_display.index = range(len(df_display))  # pakai 0-based supaya logika trend bener

    # Trend: 🔼 top 3 hijau, nomor 3 strip, sisanya 🔽 merah
    df_display['Trend'] = ''
    for i in range(len(df_display)):
        if i < 3:
            df_display.loc[i, 'Trend'] = '🔼'
        elif i == 3:
            df_display.loc[i, 'Trend'] = '—'
        else:
            df_display.loc[i, 'Trend'] = '🔽'

    st.subheader("Tabel Pertumbuhan Pelanggan")
    st.dataframe(df_display[['Bulan','Jumlah Pelanggan','Trend']])

# =========================
# SHEET 5: PRODUK POPULER
# =========================
elif current_sheet == "Produk Populer":
    df = pd.read_excel(file_path, sheet_name="Produk Populer", engine='openpyxl')
    df.columns = df.columns.str.strip()
    df['Jumlah Pembelian'] = pd.to_numeric(df['Jumlah Pembelian'], errors='coerce').fillna(0)
    df_display = df.sort_values(by='Jumlah Pembelian', ascending=False).reset_index(drop=True)
    df_display.index = range(1, len(df_display)+1)

    # Highlight Top 5
    top5_idx = df_display.index[:5]
    def highlight_top5(row):
        if row.name in top5_idx:
            return ['background-color: lightgreen']*len(row)
        else:
            return ['']*len(row)

    st.subheader("Tabel Produk Populer")
    st.dataframe(df_display.style.apply(highlight_top5, axis=1))

# =========================
# SHEET 6: PRODUK FAVORIT CUSTOMER
# =========================
elif current_sheet == "Produk Favorit Customer":
    df = pd.read_excel(file_path, sheet_name="Produk Favorit Customer", engine='openpyxl')
    df.columns = df.columns.str.strip()
    df['Jumlah Dibeli'] = pd.to_numeric(df['Jumlah Dibeli'], errors='coerce').fillna(0)
    df_display = df.reset_index(drop=True)
    df_display.index = range(1, len(df_display)+1)

    st.subheader("Tabel Produk Favorit Customer")
    st.dataframe(df_display)
