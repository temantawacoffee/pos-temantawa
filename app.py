import streamlit as st
import pandas as pd
from datetime import datetime
import os
import io

# Konfigurasi Halaman
st.set_page_config(page_title="POS TEMAN TAWA COFFEE", page_icon="☕", layout="wide")
DB_FILE = "Sistem_Inventaris_TemanTawa.xlsx"

# Inisialisasi Database Excel Lengkap
def init_database():
    if not os.path.exists(DB_FILE):
        with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
            # Master Bahan Baku
            df_bahan = pd.DataFrame([
                {"Kode": "B01", "Nama Bahan": "Biji Kopi Espresso", "Satuan": "gram", "Harga/Satuan": 180, "Stok": 2000},
                {"Kode": "B02", "Nama Bahan": "Fresh Milk 1L", "Satuan": "ml", "Harga/Satuan": 22, "Stok": 5000},
                {"Kode": "B03", "Nama Bahan": "Gula Aren Cair", "Satuan": "ml", "Harga/Satuan": 15, "Stok": 1000}
            ])
            df_bahan.to_excel(writer, sheet_name="Master Bahan", index=False)

            # Master Menu
            df_menu = pd.DataFrame([
                {"Kode Menu": "M01", "Nama Menu": "Kopi Susu Gula Aren", "Kategori": "Coffee", "Harga Offline": 22000, "Harga Online": 26000},
                {"Kode Menu": "M02", "Nama Menu": "Iced Americano", "Kategori": "Coffee", "Harga Offline": 18000, "Harga Online": 21000}
            ])
            df_menu.to_excel(writer, sheet_name="Master Menu", index=False)

            # Resep Detail
            df_resep = pd.DataFrame([
                {"Kode Menu": "M01", "Kode Bahan": "B01", "Takaran": 18},
                {"Kode Menu": "M01", "Kode Bahan": "B02", "Takaran": 150},
                {"Kode Menu": "M01", "Kode Bahan": "B03", "Takaran": 20},
                {"Kode Menu": "M02", "Kode Bahan": "B01", "Takaran": 18}
            ])
            df_resep.to_excel(writer, sheet_name="Resep", index=False)

            # Log Transaksi Penjualan
            df_sales = pd.DataFrame(columns=[
                "ID Transaksi", "Tanggal", "Kanal", "Kode Menu", "Nama Menu", 
                "Qty", "Harga Satuan", "Diskon Member", "Komisi Aplikator", "Total Omzet Bersih", "Metode Bayar", "Kasir"
            ])
            df_sales.to_excel(writer, sheet_name="Penjualan", index=False)

            # Log Waste Bahan Baku
            df_waste_bhn = pd.DataFrame(columns=["Tanggal", "Kode Bahan", "Nama Bahan", "Qty Waste", "Keterangan"])
            df_waste_bhn.to_excel(writer, sheet_name="Waste Bahan", index=False)

            # Log Waste Menu (Konsumsi Barista / Trial)
            df_waste_menu = pd.DataFrame(columns=["Tanggal", "Kode Menu", "Nama Menu", "Qty Waste", "Keterangan"])
            df_waste_menu.to_excel(writer, sheet_name="Waste Menu", index=False)

init_database()

def load_data(sheet_name):
    return pd.read_excel(DB_FILE, sheet_name=sheet_name)

def save_data(df, sheet_name):
    with pd.ExcelWriter(DB_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

# Session State Login & Cart
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
if 'cart' not in st.session_state:
    st.session_state['cart'] = []
if 'last_trx' not in st.session_state:
    st.session_state['last_trx'] = None

# Form Login
def login():
    st.markdown("<h2 style='text-align: center;'>☕ LOGIN SYSTEM - TEMAN TAWA COFFEE</h2>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        with st.container(border=True):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            if st.button("Masuk Sistem", use_container_width=True):
                if username == "admin" and password == "admin123":
                    st.session_state['logged_in'] = True
                    st.session_state['username'] = "admin"
                    st.session_state['role'] = "Admin"
                    st.success("Login Admin Berhasil!")
                    st.rerun()
                elif username == "kasir" and password == "kasir123":
                    st.session_state['logged_in'] = True
                    st.session_state['username'] = "kasir"
                    st.session_state['role'] = "Kasir"
                    st.success("Login Kasir Berhasil!")
                    st.rerun()
                else:
                    st.error("Username atau Password salah!")
            st.caption("<center><small>POS TEMAN TAWA COFFEE<br>by Clemens - Business Developer & System</small></center>", unsafe_allow_html=True)

if not st.session_state['logged_in']:
    login()
    st.stop()

# Sidebar Navigasi Berdasarkan Role
st.sidebar.title("☕ TEMAN TAWA COFFEE")
st.sidebar.write(f"👤 User: **{st.session_state['username'].upper()}**")
st.sidebar.write(f"Hak Akses: **{st.session_state['role']}**")
st.sidebar.markdown("---")

if st.session_state['role'] == "Admin":
    nav_options = [
        "🛒 Kasir POS (Penjualan)", 
        "🍔 Kelola Menu & Resep", 
        "📦 Kelola Stok & Waste", 
        "📊 Laporan Keuangan & P&L"
    ]
else:
    nav_options = [
        "🛒 Kasir POS (Penjualan)", 
        "📦 Catat Waste & Konsumsi Barista"
    ]

menu_option = st.sidebar.radio("Navigasi Menu:", nav_options)

st.sidebar.markdown("---")
if st.sidebar.button("Logout", type="primary", use_container_width=True):
    st.session_state['logged_in'] = False
    st.session_state['cart'] = []
    st.rerun()

st.sidebar.markdown("<br><br><br><center><small><b>POS TEMAN TAWA COFFEE</b><br>by Clemens - Business Developer & System</small></center>", unsafe_allow_html=True)

# ==========================================
# 1. MODUL KASIR POS (PENJUALAN + MEMBER + STRUK)
# ==========================================
if menu_option == "🛒 Kasir POS (Penjualan)":
    st.header("🛒 Kasir POS Operasional - Teman Tawa Coffee")
    
    df_menu = load_data("Master Menu")
    df_bahan = load_data("Master Bahan")
    df_resep = load_data("Resep")
    
    col_left, col_right = st.columns([3, 2], gap="medium")
    
    with col_left:
        st.subheader("Daftar Menu & Katalog")
        channel = st.selectbox("Kanal Penjualan:", ["Offline / Dine-in", "GoFood", "GrabFood", "ShopeeFood"])
        komisi_rate = 0.20 if channel != "Offline / Dine-in" else 0.0
        
        search_query = st.text_input("🔍 Cari Menu:", "").lower()
        filtered_menu = df_menu[df_menu['Nama Menu'].str.lower().str.contains(search_query)]
        
        if filtered_menu.empty:
            st.warning("Menu tidak ditemukan.")
        else:
            cols = st.columns(2)
            for idx, row in filtered_menu.iterrows():
                harga = row['Harga Offline'] if channel == "Offline / Dine-in" else row['Harga Online']
                with cols[idx % 2]:
                    with st.container(border=True):
                        st.markdown(f"**{row['Nama Menu']}**")
                        st.caption(f"Kategori: {row['Kategori']}")
                        st.write(f"Harga: **Rp {harga:,.0f}**")
                        if st.button(f"➕ Tambah", key=f"add_{row['Kode Menu']}"):
                            st.session_state['cart'].append({
                                "Kode Menu": row['Kode Menu'],
                                "Nama Menu": row['Nama Menu'],
                                "Harga": harga,
                                "Kanal": channel,
                                "Komisi Rate": komisi_rate
                            })
                            st.toast(f"{row['Nama Menu']} ditambahkan!", icon="✅")

    with col_right:
        st.subheader("📋 Keranjang & Pembayaran")
        if len(st.session_state['cart']) == 0:
            st.info("Keranjang belanja masih kosong.")
        else:
            cart_df = pd.DataFrame(st.session_state['cart'])
            summary_cart = cart_df.groupby(['Kode Menu', 'Nama Menu', 'Harga', 'Kanal', 'Komisi Rate']).size().reset_index(name='Qty')
            summary_cart['Subtotal'] = summary_cart['Harga'] * summary_cart['Qty']
            
            st.dataframe(summary_cart[['Nama Menu', 'Harga', 'Qty', 'Subtotal']], use_container_width=True)
            
            subtotal_gross = summary_cart['Subtotal'].sum()
            
            # Fitur Diskon Member (10%)
            is_member = st.checkbox("⭐ Terapkan Diskon Member VIP (10%)")
            diskon_nominal = subtotal_gross * 0.10 if is_member else 0.0
            total_gross_after_discount = subtotal_gross - diskon_nominal
            
            total_komisi = total_gross_after_discount * summary_cart['Komisi Rate'].iloc[0]
            total_net = total_gross_after_discount - total_komisi
            
            st.markdown(f"Subtotal: Rp {subtotal_gross:,.0f}")
            if is_member:
                st.markdown(f"Diskon Member (10%): -Rp {diskon_nominal:,.0f}")
            st.markdown(f"### Total Bayar: **Rp {total_gross_after_discount:,.0f}**")
            
            if channel != "Offline / Dine-in":
                st.caption(f"*Komisi Aplikator: Rp {total_komisi:,.0f}")
                st.caption(f"*Omzet Bersih Toko: **Rp {total_net:,.0f}**")
            
            metode_bayar = st.radio("Metode Pembayaran:", ["Cash", "QRIS", "Transfer"], horizontal=True)
            
            c1, c2 = st.columns(2)
            with c1:
                if st.button("🔴 Reset", use_container_width=True):
                    st.session_state['cart'] = []
                    st.rerun()
            with c2:
                if st.button("🟢 Bayar & Cetak", type="primary", use_container_width=True):
                    df_sales = load_data("Penjualan")
                    trans_id = f"TRX-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
                    
                    new_sales = []
                    for _, row_item in summary_cart.iterrows():
                        item_proporsi = row_item['Subtotal'] / subtotal_gross if subtotal_gross > 0 else 0
                        item_diskon = diskon_nominal * item_proporsi
                        item_gross_final = row_item['Subtotal'] - item_diskon
                        net_item = item_gross_final * (1 - row_item['Komisi Rate'])
                        
                        new_sales.append({
                            "ID Transaksi": trans_id,
                            "Tanggal": now_str,
                            "Kanal": row_item['Kanal'],
                            "Kode Menu": row_item['Kode Menu'],
                            "Nama Menu": row_item['Nama Menu'],
                            "Qty": row_item['Qty'],
                            "Harga Satuan": row_item['Harga'],
                            "Diskon Member": item_diskon,
                            "Komisi Aplikator": item_gross_final * row_item['Komisi Rate'],
                            "Total Omzet Bersih": net_item,
                            "Metode Bayar": metode_bayar,
                            "Kasir": st.session_state['username']
                        })
                    
                    df_sales = pd.concat([df_sales, pd.DataFrame(new_sales)], ignore_index=False)
                    save_data(df_sales, "Penjualan")
                    
                    # Potong Stok Bahan Baku Otomatis
                    for _, row_item in summary_cart.iterrows():
                        resep_items = df_resep[df_resep['Kode Menu'] == row_item['Kode Menu']]
                        for _, r in resep_items.iterrows():
                            total_deduct = r['Takaran'] * row_item['Qty']
                            df_bahan.loc[df_bahan['Kode'] == r['Kode Bahan'], 'Stok'] -= total_deduct
                    
                    save_data(df_bahan, "Master Bahan")
                    
                    st.session_state['last_trx'] = {
                        "id": trans_id, "tanggal": now_str, "items": summary_cart, 
                        "subtotal": subtotal_gross, "diskon": diskon_nominal, 
                        "total": total_gross_after_discount, "bayar": metode_bayar
                    }
                    st.session_state['cart'] = []
                    st.success("✅ Transaksi Sukses! Stok Terpotong.")
                    st.rerun()

        # Preview & Tombol Cetak Struk
        if st.session_state['last_trx']:
            st.markdown("---")
            st.subheader("🧾 Preview Struk Kasir (Thermal 58mm/80mm)")
            t = st.session_state['last_trx']
            
            receipt_html = f"""
            <div style="font-family: monospace; width: 280px; background: white; color: black; padding: 10px; border: 1px solid #ccc;">
                <center><b>TEMAN TAWA COFFEE</b><br>Enjoy Your Coffee & Smile<br>--------------------------------</center>
                ID: {t['id']}<br>Waktu: {t['tanggal']}<br>--------------------------------<br>
            """
            for _, item in t['items'].iterrows():
                receipt_html += f"{item['Nama Menu']} x{item['Qty']}<br>&nbsp;&nbsp;Rp {item['Subtotal']:,.0f}<br>"
            receipt_html += f"""
                --------------------------------<br>
                Subtotal: Rp {t['subtotal']:,.0f}<br>
                Diskon Member: -Rp {t['diskon']:,.0f}<br>
                <b>TOTAL: Rp {t['total']:,.0f}</b><br>
                Metode: {t['bayar']}<br>
                --------------------------------<br>
                <center>Terima Kasih Atas Kunjungan Anda!<br><small>by Clemens - Business Dev</small></center>
            </div>
            """
            st.markdown(receipt_html, unsafe_allow_html=True)
            if st.button("🖨️ Cetak Struk (Kirim ke Printer Bluetooth/Thermal)"):
                st.components.v1.html(f"<script>window.print();</script>", height=0)

# ==========================================
# 2. MODUL KELOLA MENU & RESEP (ADMIN ONLY)
# ==========================================
elif menu_option == "🍔 Kelola Menu & Resep":
    st.header("🍔 Manajemen Menu & Pengaturan Resep")
    
    df_menu = load_data("Master Menu")
    df_bahan = load_data("Master Bahan")
    df_resep = load_data("Resep")
    
    tab1, tab2, tab3, tab4 = st.tabs([
        "Tambah Menu Baru", 
        "✏️ Edit Harga / Nama Menu", 
        "🛠️ Edit / Update Resep Menu", 
        "Daftar Menu & Resep Aktif"
    ])
    
    with tab1:
        st.subheader("Form Tambah Menu Baru")
        with st.form("form_menu"):
            c_code, c_cat = st.columns(2)
            with c_code:
                kode_menu = st.text_input("Kode Menu Unik (Contoh: M03)")
            with c_cat:
                kategori = st.selectbox("Kategori", ["Coffee", "Non-Coffee", "Snack", "Manual Brew"])
            
            nama_menu = st.text_input("Nama Menu Baru")
            
            c_off, c_on = st.columns(2)
            with c_off:
                harga_off = st.number_input("Harga Offline / Dine-in (Rp)", min_value=0, value=20000, step=1000)
            with c_on:
                harga_on = st.number_input("Harga Online / GoFood (Rp)", min_value=0, value=25000, step=1000)
            
            submitted = st.form_submit_button("Simpan Menu Baru")
            if submitted:
                if kode_menu and nama_menu:
                    new_m = pd.DataFrame([{
                        "Kode Menu": kode_menu, "Nama Menu": nama_menu, 
                        "Kategori": kategori, "Harga Offline": harga_off, "Harga Online": harga_on
                    }])
                    df_menu = pd.concat([df_menu, new_m], ignore_index=True)
                    save_data(df_menu, "Master Menu")
                    st.success(f"Menu {nama_menu} berhasil ditambahkan!")
                    st.rerun()
                else:
                    st.error("Kode Menu dan Nama Menu wajib diisi!")

    with tab2:
        st.subheader("✏️ Edit Harga Jual Menu (Kenaikan Bahan Baku)")
        st.info("Gunakan form ini untuk menyesuaikan harga jual menu saat terjadi kenaikan harga modal bahan baku.")
        
        with st.form("form_edit_harga_menu"):
            m_edit = st.selectbox("Pilih Menu yang Ingin Diubah Harganya:", df_menu['Kode Menu'] + " - " + df_menu['Nama Menu'])
            km_selected = m_edit.split(" - ")[0]
            menu_row = df_menu[df_menu['Kode Menu'] == km_selected].iloc[0]
            
            h_off_lama = int(menu_row['Harga Offline'])
            h_on_lama = int(menu_row['Harga Online'])
            nama_lama = str(menu_row['Nama Menu'])
            
            st.caption(f"Harga Offline Saat Ini: **Rp {h_off_lama:,.0f}** | Harga Online Saat Ini: **Rp {h_on_lama:,.0f}**")
            
            c_nm, c_cat_edit = st.columns(2)
            with c_nm:
                nama_menu_baru = st.text_input("Nama Menu:", value=nama_lama)
            with c_cat_edit:
                kategori_baru = st.selectbox("Kategori:", ["Coffee", "Non-Coffee", "Snack", "Manual Brew"], index=["Coffee", "Non-Coffee", "Snack", "Manual Brew"].index(menu_row['Kategori']) if menu_row['Kategori'] in ["Coffee", "Non-Coffee", "Snack", "Manual Brew"] else 0)
                
            c_ho, c_hn = st.columns(2)
            with c_ho:
                harga_off_baru = st.number_input("Harga Offline Baru (Rp):", min_value=0, value=h_off_lama, step=1000)
            with c_hn:
                harga_on_baru = st.number_input("Harga Online Baru (Rp):", min_value=0, value=h_on_lama, step=1000)
                
            if st.form_submit_button("🔄 Update Harga / Data Menu"):
                df_menu.loc[df_menu['Kode Menu'] == km_selected, 'Nama Menu'] = nama_menu_baru
                df_menu.loc[df_menu['Kode Menu'] == km_selected, 'Kategori'] = kategori_baru
                df_menu.loc[df_menu['Kode Menu'] == km_selected, 'Harga Offline'] = harga_off_baru
                df_menu.loc[df_menu['Kode Menu'] == km_selected, 'Harga Online'] = harga_on_baru
                
                save_data(df_menu, "Master Menu")
                st.success(f"Menu '{nama_menu_baru}' berhasil diperbarui! Harga Offline: Rp {harga_off_baru:,.0f}, Online: Rp {harga_on_baru:,.0f}")
                st.rerun()

    with tab3:
        st.subheader("🛠️ Tambah / Edit Takaran Resep Bahan")
        st.info("Pilih menu dan bahan baku untuk menambahkan atau mengubah takaran gramasi/mililiter. Perubahan ini hanya berpengaruh pada transaksi mendatang.")
        
        with st.form("form_update_resep"):
            pilih_menu = st.selectbox("Pilih Menu:", df_menu['Kode Menu'] + " - " + df_menu['Nama Menu'])
            pilih_bahan = st.selectbox("Pilih Bahan Baku:", df_bahan['Kode'] + " - " + df_bahan['Nama Bahan'])
            takaran_baru = st.number_input("Takaran Baru per Cup (sesuai satuan bahan):", min_value=0.1, value=18.0, step=1.0)
            
            submit_update_resep = st.form_submit_button("🔄 Simpan / Update Resep")
            if submit_update_resep:
                k_menu = pilih_menu.split(" - ")[0]
                k_bahan = pilih_bahan.split(" - ")[0]
                
                mask = (df_resep['Kode Menu'] == k_menu) & (df_resep['Kode Bahan'] == k_bahan)
                if mask.any():
                    df_resep.loc[mask, 'Takaran'] = takaran_baru
                    st.success(f"Takaran resep untuk {pilih_menu} berhasil DIPERBARUI menjadi {takaran_baru}!")
                else:
                    new_r = pd.DataFrame([{"Kode Menu": k_menu, "Kode Bahan": k_bahan, "Takaran": takaran_baru}])
                    df_resep = pd.concat([df_resep, new_r], ignore_index=True)
                    st.success(f"Bahan baru berhasil DITAMBAHKAN ke dalam resep {pilih_menu}!")
                
                save_data(df_resep, "Resep")
                st.rerun()

        st.markdown("---")
        st.subheader("🗑️ Hapus Bahan dari Resep Menu")
        with st.form("form_hapus_resep"):
            c_m, c_b = st.columns(2)
            with c_m:
                del_menu = st.selectbox("Pilih Menu:", df_menu['Kode Menu'] + " - " + df_menu['Nama Menu'], key="del_m")
            with c_b:
                del_bahan = st.selectbox("Pilih Bahan yang Ingin Dihapus:", df_bahan['Kode'] + " - " + df_bahan['Nama Bahan'], key="del_b")
            
            submit_del_resep = st.form_submit_button("❌ Hapus Bahan dari Resep")
            if submit_del_resep:
                km = del_menu.split(" - ")[0]
                kb = del_bahan.split(" - ")[0]
                df_resep = df_resep[~((df_resep['Kode Menu'] == km) & (df_resep['Kode Bahan'] == kb))]
                save_data(df_resep, "Resep")
                st.warning("Bahan berhasil dihapus dari resep menu tersebut!")
                st.rerun()

    with tab4:
        st.subheader("Daftar Menu Aktif")
        st.dataframe(df_menu, use_container_width=True)
        st.subheader("Komposisi Resep Bahan Saat Ini")
        st.dataframe(df_resep, use_container_width=True)

# ==========================================
# 3. MODUL KELOLA STOK & EDIT HARGA BAHAN (ADMIN & KASIR)
# ==========================================
elif menu_option in ["📦 Kelola Stok & Waste", "📦 Catat Waste & Konsumsi Barista"]:
    st.header("📦 Manajemen Stok, Waste Bahan & Konsumsi Barista")
    
    df_bahan = load_data("Master Bahan")
    df_menu = load_data("Master Menu")
    df_resep = load_data("Resep")
    
    if st.session_state['role'] == "Admin":
        tab1, tab2, tab3, tab4 = st.tabs(["Stok & Restock", "➕ Tambah Bahan Baru", "Waste Bahan Baku", "Waste Menu (Konsumsi Barista)"])
    else:
        tab1, tab3, tab4 = st.tabs(["Cek Stok Bahan", "Waste Bahan Baku", "Waste Menu (Konsumsi Barista)"])
    
    with tab1:
        st.dataframe(df_bahan, use_container_width=True)
        if st.session_state['role'] == "Admin":
            c1, c2 = st.columns(2, gap="large")
            with c1:
                st.subheader("➕ Restock (Tambah Stok Masuk)")
                with st.form("form_restock"):
                    b_pilih = st.selectbox("Pilih Bahan Baku:", df_bahan['Nama Bahan'].tolist())
                    q_tambah = st.number_input("Jumlah Stok Masuk:", min_value=1.0, value=1000.0, step=100.0)
                    if st.form_submit_button("➕ Tambah Stok"):
                        df_bahan.loc[df_bahan['Nama Bahan'] == b_pilih, 'Stok'] += q_tambah
                        save_data(df_bahan, "Master Bahan")
                        st.success("Stok berhasil diperbarui!")
                        st.rerun()
            
            with c2:
                st.subheader("✏️ Edit Harga Satuan / Modal Bahan")
                with st.form("form_edit_harga_bahan"):
                    b_edit = st.selectbox("Pilih Bahan untuk Diubah Harganya:", df_bahan['Nama Bahan'].tolist())
                    bahan_row = df_bahan[df_bahan['Nama Bahan'] == b_edit].iloc[0]
                    
                    harga_lama = float(bahan_row['Harga/Satuan'])
                    satuan_lama = str(bahan_row['Satuan'])
                    
                    st.caption(f"Harga Satuan Saat Ini: **Rp {harga_lama:,.2f} / {satuan_lama}**")
                    harga_baru = st.number_input("Harga Satuan Baru (Rp):", min_value=0.0, value=harga_lama, step=1.0)
                    
                    if st.form_submit_button("🔄 Update Harga Satuan"):
                        df_bahan.loc[df_bahan['Nama Bahan'] == b_edit, 'Harga/Satuan'] = harga_baru
                        save_data(df_bahan, "Master Bahan")
                        st.success(f"Harga satuan untuk '{b_edit}' berhasil diubah menjadi Rp {harga_baru:,.2f}!")
                        st.rerun()

    if st.session_state['role'] == "Admin":
        with tab2:
            st.subheader("Form Pendaftaran Bahan Baku Baru ke Gudang")
            with st.form("form_bahan_baru"):
                c_kb, c_nb = st.columns(2)
                with c_kb:
                    kode_bhn_baru = st.text_input("Kode Bahan Unik (Contoh: B04)")
                with c_nb:
                    nama_bhn_baru = st.text_input("Nama Bahan Baku (Contoh: Matcha Powder)")
                
                c_st, c_hp = st.columns(2)
                with c_st:
                    satuan_bhn = st.selectbox("Satuan:", ["gram", "ml", "pcs", "kg", "liter", "pack", "sachet"])
                with c_hp:
                    harga_satuan = st.number_input("Biaya / Harga per Satuan (Rp):", min_value=0.0, value=100.0, step=10.0)
                
                stok_awal = st.number_input("Stok Awal Saat Ini:", min_value=0.0, value=1000.0)
                
                submit_bahan_baru = st.form_submit_button("Simpan Bahan Baku Baru")
                if submit_bahan_baru:
                    if kode_bhn_baru and nama_bhn_baru:
                        if kode_bhn_baru in df_bahan['Kode'].values:
                            st.error("Kode bahan tersebut sudah digunakan! Gunakan kode lain.")
                        else:
                            new_bhn_row = pd.DataFrame([{
                                "Kode": kode_bhn_baru,
                                "Nama Bahan": nama_bhn_baru,
                                "Satuan": satuan_bhn,
                                "Harga/Satuan": harga_satuan,
                                "Stok": stok_awal
                            }])
                            df_bahan = pd.concat([df_bahan, new_bhn_row], ignore_index=True)
                            save_data(df_bahan, "Master Bahan")
                            st.success(f"Bahan baku baru '{nama_bhn_baru}' berhasil ditambahkan ke sistem!")
                            st.rerun()
                    else:
                        st.error("Kode dan Nama Bahan wajib diisi!")

    with tab3:
        st.subheader("Pencatatan Waste Bahan Baku (Susu Tumpah / Kalibrasi)")
        with st.form("waste_bhn_form"):
            w_bahan = st.selectbox("Bahan Terbuang:", df_bahan['Nama Bahan'].tolist())
            w_qty = st.number_input("Jumlah Qty Terbuang:", min_value=1.0, value=18.0)
            w_ket = st.text_input("Keterangan:", value="Kalibrasi Pagi / Susu Tumpah")
            
            if st.form_submit_button("Simpan Waste Bahan"):
                df_wb = load_data("Waste Bahan")
                kode_b = df_bahan.loc[df_bahan['Nama Bahan'] == w_bahan, 'Kode'].values[0]
                new_wb = pd.DataFrame([{
                    "Tanggal": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "Kode Bahan": kode_b, "Nama Bahan": w_bahan,
                    "Qty Waste": w_qty, "Keterangan": w_ket
                }])
                df_wb = pd.concat([df_wb, new_wb], ignore_index=True)
                save_data(df_wb, "Waste Bahan")
                
                df_bahan.loc[df_bahan['Nama Bahan'] == w_bahan, 'Stok'] -= w_qty
                save_data(df_bahan, "Master Bahan")
                st.warning("Waste bahan tercatat & stok dipotong!")
                st.rerun()

    with tab4:
        st.subheader("Pencatatan Waste Menu / Konsumsi Barista (Staff Meal / Trial)")
        with st.form("waste_menu_form"):
            w_menu_pilih = st.selectbox("Menu Terbuang / Dikonsumsi:", df_menu['Kode Menu'] + " - " + df_menu['Nama Menu'])
            w_menu_qty = st.number_input("Jumlah Cup:", min_value=1, value=1)
            w_menu_ket = st.text_input("Keterangan:", value="Konsumsi Shift Barista")
            
            if st.form_submit_button("Simpan Waste Menu & Potong Bahan"):
                k_menu = w_menu_pilih.split(" - ")[0]
                nm_menu = w_menu_pilih.split(" - ")[1]
                
                df_wm = load_data("Waste Menu")
                new_wm = pd.DataFrame([{
                    "Tanggal": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "Kode Menu": k_menu, "Nama Menu": nm_menu,
                    "Qty Waste": w_menu_qty, "Keterangan": w_menu_ket
                }])
                df_wm = pd.concat([df_wm, new_wm], ignore_index=True)
                save_data(df_wm, "Waste Menu")
                
                resep_terkait = df_resep[df_resep['Kode Menu'] == k_menu]
                for _, r in resep_terkait.iterrows():
                    total_potong = r['Takaran'] * w_menu_qty
                    df_bahan.loc[df_bahan['Kode'] == r['Kode Bahan'], 'Stok'] -= total_potong
                save_data(df_bahan, "Master Bahan")
                
                st.warning("Waste menu berhasil dicatat & seluruh bahan baku resepnya terpotong otomatis!")
                st.rerun()

# ==========================================
# 4. MODUL LAPORAN KEUANGAN P&L & TUTUP BUKU (ADMIN ONLY)
# ==========================================
elif menu_option == "📊 Laporan Keuangan & P&L":
    st.header("📊 Dashboard Laporan Penjualan & Ekspor Excel")
    
    df_sales = load_data("Penjualan")
    df_wb = load_data("Waste Bahan")
    df_wm = load_data("Waste Menu")
    
    if df_sales.empty:
        st.info("Belum ada data transaksi penjualan.")
    else:
        c1, c2, c3, c4 = st.columns(4)
        omzet_kotor = (df_sales['Qty'] * df_sales['Harga Satuan']).sum()
        total_diskon = df_sales['Diskon Member'].sum()
        total_komisi = df_sales['Komisi Aplikator'].sum()
        omzet_bersih = df_sales['Total Omzet Bersih'].sum()
        
        c1.metric("Omzet Kotor", f"Rp {omzet_kotor:,.0f}")
        c2.metric("Total Diskon Member", f"Rp {total_diskon:,.0f}")
        c3.metric("Potongan Komisi", f"Rp {total_komisi:,.0f}")
        c4.metric("Omzet Bersih Toko", f"Rp {omzet_bersih:,.0f}")
        
        st.markdown("---")
        st.subheader("📥 Unduh Laporan Harian / Bulanan (.xlsx)")
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_sales.to_excel(writer, sheet_name="Penjualan", index=False)
            df_wb.to_excel(writer, sheet_name="Waste Bahan", index=False)
            df_wm.to_excel(writer, sheet_name="Waste Menu", index=False)
            load_data("Master Bahan").to_excel(writer, sheet_name="Stok Gudang", index=False)
        
        excel_bytes = output.getvalue()
        st.download_button(
            label="📥 Download Laporan Teman Tawa (.xlsx)",
            data=excel_bytes,
            file_name=f"Laporan_TemanTawaCoffee_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        st.markdown("---")
        st.subheader("Riwayat Transaksi Penjualan")
        st.dataframe(df_sales, use_container_width=True)
        
        if not df_wm.empty:
            st.subheader("Riwayat Waste Menu / Konsumsi Barista")
            st.dataframe(df_wm, use_container_width=True)

    # Fitur Tutup Buku / Reset Harian
    st.markdown("---")
    st.subheader("🔄 Tutup Buku & Reset Transaksi Harian")
    st.warning("⚠️ **Perhatian:** Pastikan Anda sudah mengklik tombol **'Download Laporan Teman Tawa'** di atas sebelum melakukan reset harian agar data omzet tidak hilang!")
    
    if st.button("🗑️ Reset Riwayat Transaksi & Waste Hari Ini", type="primary"):
        with pd.ExcelWriter(DB_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            # Kosongkan Penjualan
            pd.DataFrame(columns=[
                "ID Transaksi", "Tanggal", "Kanal", "Kode Menu", "Nama Menu", 
                "Qty", "Harga Satuan", "Diskon Member", "Komisi Aplikator", "Total Omzet Bersih", "Metode Bayar", "Kasir"
            ]).to_excel(writer, sheet_name="Penjualan", index=False)
            
            # Kosongkan Waste Bahan
            pd.DataFrame(columns=["Tanggal", "Kode Bahan", "Nama Bahan", "Qty Waste", "Keterangan"]).to_excel(writer, sheet_name="Waste Bahan", index=False)
            
            # Kosongkan Waste Menu
            pd.DataFrame(columns=["Tanggal", "Kode Menu", "Nama Menu", "Qty Waste", "Keterangan"]).to_excel(writer, sheet_name="Waste Menu", index=False)
            
            # Master Bahan, Master Menu, & Resep dibiarkan utuh
            load_data("Master Bahan").to_excel(writer, sheet_name="Master Bahan", index=False)
            load_data("Master Menu").to_excel(writer, sheet_name="Master Menu", index=False)
            load_data("Resep").to_excel(writer, sheet_name="Resep", index=False)
            
        st.success("✅ Tutup Buku Berhasil! Riwayat transaksi dan waste harian telah direset untuk hari baru. Stok gudang dan menu tetap aman.")
        st.rerun()