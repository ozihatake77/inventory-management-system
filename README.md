# 📦 Stok Toko Elektronik

> **Sistem Manajemen Stok & Penjualan Toko Elektronik Rumah Tangga**

Aplikasi web untuk mengelola inventaris, penjualan, dan keuangan toko elektronik rumah tangga. Dirancang untuk memudahkan pemilik toko dan karyawan dalam mengelola stok barang, mencatat transaksi, dan memantau keuntungan.

## ✨ Fitur Utama

### 📊 Dashboard
- Ringkasan penjualan harian, mingguan, dan bulanan
- Grafik penjualan 7 hari terakhir
- Produk terlaris bulan ini
- Notifikasi stok menipis
- Mutasi stok terbaru

### 📱 Manajemen Produk
- CRUD produk lengkap (tambah, edit, hapus)
- Kategori produk (TV, Kulkas, Mesin Cuci, AC, dll)
- Harga modal & harga jual
- Stok minimum & barcode
- Riwayat perubahan harga

### 📥 Stok Masuk & Keluar
- Catat barang masuk dari supplier
- Catat barang keluar (rusak, hilang, retur)
- Otomatis update stok saat penjualan
- Riwayat mutasi stok lengkap

### 💰 Penjualan & Nota
- Input penjualan dengan pilih pelanggan
- Otomatis hitung keuntungan
- Cetak nota/struk penjualan
- Support metode bayar tunai & hutang

### 👥 Pelanggan
- Data pelanggan lengkap (nama, alamat, telepon)
- Riwayat transaksi per pelanggan
- Piutang pelanggan

### 💳 Hutang & Piutang
- Catat penjualan dengan metode hutang
- Jatuh tempo otomatis (30 hari)
- Bayar hutang sebagian/seluruhnya
- Notifikasi hutang jatuh tempo

### 📈 Laporan Keuntungan
- Laporan harian, mingguan, bulanan
- Grafik penjualan & keuntungan
- Export ke Excel (.xlsx)

### 🔐 Autentikasi & Role
- Login dengan username & password
- Role **Bos** (akses penuh) dan **Karyawan** (terbatas)
- Kelola user (tambah, edit, hapus)

### 💾 Backup & Restore
- Backup database otomatis
- Restore dari file backup
- Download backup

### 📱 Responsive Design
- Tampilan responsif (desktop & mobile)
- Sidebar navigasi
- Dark mode ready

## 🛠️ Tech Stack

| Komponen | Teknologi |
|----------|-----------|
| **Backend** | Python FastAPI |
| **Database** | SQLite |
| **Frontend** | HTML, TailwindCSS, Alpine.js |
| **Chart** | Chart.js |
| **Export** | OpenPyXL (Excel) |
| **Server** | Uvicorn |

## 🚀 Cara Install

### 1. Clone Repository
```bash
git clone https://github.com/ozihatake77/App-penjualan.git
cd App-penjualan
```

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```

### 3. Jalankan Aplikasi
```bash
python app.py
```

### 4. Buka Browser
```
http://localhost:8000
```

## 🔑 Login Default

| Role | Username | Password |
|------|----------|----------|
| **Bos** | `admin` | `admin123` |

## 📸 Screenshots

### Dashboard
![Dashboard](https://via.placeholder.com/800x400?text=Dashboard+Screenshot)

### Data Produk
![Produk](https://via.placeholder.com/800x400?text=Produk+Screenshot)

### Penjualan
![Penjualan](https://via.placeholder.com/800x400?text=Penjualan+Screenshot)

## 📁 Struktur Project

```
App-penjualan/
├── app.py                 # Backend FastAPI (semua route & logic)
├── requirements.txt       # Dependencies Python
├── Procfile              # Config deployment Railway
├── .gitignore            # File yang di-ignore Git
├── static/               # File statis (CSS, JS, gambar)
└── templates/            # Template HTML (Jinja2)
    ├── base.html         # Layout utama (sidebar + header)
    ├── login.html        # Halaman login
    ├── dashboard.html    # Dashboard utama
    ├── produk.html       # Manajemen produk
    ├── kategori.html     # Manajemen kategori
    ├── pelanggan.html    # Manajemen pelanggan
    ├── stok_masuk.html   # Input stok masuk
    ├── stok_keluar.html  # Input stok keluar
    ├── opname.html       # Stok opname
    ├── penjualan.html    # Input penjualan
    ├── nota.html         # Cetak nota
    ├── hutang.html       # Hutang & piutang
    ├── laporan.html      # Laporan keuntungan
    ├── users.html        # Kelola user
    ├── backup.html       # Backup & restore
    ├── notifikasi.html   # Notifikasi
    └── riwayat_harga.html # Riwayat harga
```

## 🎯 Target Pengguna

- Toko elektronik rumah tangga
- Usaha kecil-menengah (UKM)
- Freelance developer yang butuh app inventaris

## 💡 Keunggulan

- ✅ **Gratis & Open Source** - Tidak perlu bayar lisensi
- ✅ **Mudah Digunakan** - Interface intuitif
- ✅ **Responsif** - Bisa diakses dari HP/Tablet
- ✅ **Offline Ready** - Database SQLite lokal
- ✅ **Customizable** - Mudah dimodifikasi sesuai kebutuhan
- ✅ **Deploy Mudah** - Bisa di Railway, VPS, atau hosting biasa

## 📞 Kontak

**Developer:** Ozihatake  
**GitHub:** [@ozihatake77](https://github.com/ozihatake77)  
**Email:** ozihatake77@gmail.com

## 📄 License

MIT License - Silakan digunakan dan dimodifikasi sesuai kebutuhan.

---

**Built with ❤️ by Yamada for Ozihatake**
