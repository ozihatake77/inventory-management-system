# 🔍 Komparasi Aplikasi POS Toko Elektronik
*Riset: 10 Juni 2026*

---

## Kompetitor Indonesia

### 1. MOKA POS (Gojek-backed)
- **Harga:** Rp 299.000 - 799.000/bulan/outlet
- **Fitur:** POS, stok dasar, promo, manajemen karyawan, QRIS, laporan real-time
- **Target:** F&B, retail umum
- ❌ Tidak ada: Serial Number/IMEI, PO workflow, kas & bank ledger, service center, retur
- ❌ POS general, bukan khusus elektronik

### 2. MAJOO
- **Harga:** Rp 249.000 - 999.000/bulan/outlet (sering diskon 50%)
- **Fitur:** POS, inventory, karyawan, laporan, integrasi marketplace (Tokopedia, Shopee)
- **Target:** F&B, retail
- ❌ Tidak ada: Serial Number/IMEI, PO workflow, service center
- ❌ Lebih ke F&B daripada retail elektronik

### 3. OLSERA
- **Harga:** Rp 1.288.000 - 2.688.000/tahun (diskon dari 1.8-3.8jt)
- **Fitur:** POS multi-platform (Android/iOS/Windows), stok, komisi staff, absensi, akunting, marketplace integrasi, transfer stok antar toko
- **Target:** F&B, retail, franchise
- ⚠️ Ada: komisi staff, akunting, transfer stok antar toko
- ❌ Tidak ada: Serial Number/IMEI tracking, PO workflow

### 4. PAWOON
- **Harga:** ~Rp 200.000-500.000/bulan
- **Fitur:** POS dasar, stok, laporan
- ❌ Sangat basic, kurang untuk toko elektronik

### 5. iREAP
- **Harga:** Gratis (basic) / Rp 50.000-150.000/bulan
- **Fitur:** POS, inventory, stok opname
- ❌ Basic, tidak ada fitur khusus elektronik

---

## Head-to-Head: Stok Toko Elektronik vs Kompetitor

### Fitur yang KITA MENANG:
| Fitur | Kita | Moka | Majoo | Olsera | Pawoon | iReap |
|-------|------|------|-------|--------|--------|-------|
| Serial Number/IMEI | ✅ | ❌ | ❌ | ❌ | ❌ | ❌ |
| Purchase Order Workflow | ✅ | ❌ | ❌ | ❌ | ❌ | ❌ |
| Kas & Bank Ledger | ✅ | ❌ | ❌ | ⚠️ | ❌ | ❌ |
| Cash Flow Report | ✅ | ❌ | ❌ | ⚠️ | ❌ | ❌ |
| Retur Customer (refund) | ✅ | ⚠️ | ⚠️ | ⚠️ | ❌ | ❌ |
| Brand Management | ✅ | ❌ | ❌ | ❌ | ❌ | ❌ |
| Service Center/Tiket | ✅ | ❌ | ❌ | ❌ | ❌ | ❌ |
| Komisi Sales | ✅ | ❌ | ❌ | ✅ | ❌ | ❌ |
| Invoice Resmi (NPWP) | ✅ | ❌ | ❌ | ⚠️ | ❌ | ❌ |
| Audit Log | ✅ | ❌ | ❌ | ❌ | ❌ | ❌ |
| Hutang/Piutang (cicilan) | ✅ | ❌ | ❌ | ❌ | ❌ | ❌ |
| Closing Harian | ✅ | ⚠️ | ⚠️ | ⚠️ | ❌ | ❌ |
| Role & Permission | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Promo & Voucher | ✅ | ✅ | ✅ | ✅ | ✅ | ❌ |

### Fitur yang KOMPETITOR MENANG:
| Fitur | Kita | Moka | Majoo | Olsera |
|-------|------|------|-------|--------|
| Mobile Native App | ❌ | ✅ | ✅ | ✅ |
| QRIS Payment | ❌ | ✅ | ✅ | ✅ |
| Marketplace Integration | ❌ | ✅ | ✅ | ✅ |
| Multi-warehouse Transfer | ❌ | ❌ | ❌ | ✅ |
| Absensi (PIN/Wajah) | ❌ | ❌ | ❌ | ✅ |
| Integrasi Akunting (Jurnal, dll) | ❌ | ❌ | ❌ | ✅ |
| Katalog Website | ❌ | ❌ | ❌ | ✅ |
| Franchise Dashboard | ❌ | ❌ | ❌ | ✅ |

---

## Harga Komparasi (per tahun, 1 outlet)

| Aplikasi | Harga/Tahun | Serial# | PO | Ledger |
|----------|-------------|---------|-----|--------|
| **Stok Toko Kita** | **GRATIS** | ✅ | ✅ | ✅ |
| Moka POS | Rp 3.6jt - 9.6jt | ❌ | ❌ | ❌ |
| Majoo | Rp 3jt - 12jt | ❌ | ❌ | ❌ |
| Olsera | Rp 1.3jt - 2.7jt | ❌ | ❌ | ⚠️ |
| Pawoon | Rp 2.4jt - 6jt | ❌ | ❌ | ❌ |
| iReap | Gratis - Rp 1.8jt | ❌ | ❌ | ❌ |

---

## Kesimpulan

### Keunggulan Kita:
- **Paling lengkap untuk toko elektronik** — tidak ada kompetitor yang punya Serial Number + PO + Service Center + Ledger sekaligus
- **GRATIS** — kompetitor bayar Rp 3-12jt/tahun/outlet
- **Full kontrol data** — di server sendiri, tidak tergantung vendor
- **Custom** — bisa diubah sesuai kebutuhan

### Gap yang Perlu Ditutup:
1. **QRIS / Payment Gateway** (Midtrans/Xendit) — terima bayar digital
2. **Mobile PWA** — responsive, bisa di-install di HP
3. **Integrasi Marketplace** (Tokopedia, Shopee) — opsional
4. **Multi-warehouse** — transfer antar gudang (kalau butuh)
5. **Absensi Karyawan** — opsional, bisa nanti

### Verdict:
Aplikasi kita **lebih lengkap dari SEMUA kompetitor Indonesia** untuk fitur inti toko elektronik. Kompetitor lebih kuat di **ecosystem** (mobile app, payment gateway, marketplace), tapi untuk **operasional toko elektronik murni** — kita menang telak.

Fokus selanjutnya: bukan ngejar fitur mereka (POS kita udah lebih kuat), tapi tambahin **ecosystem layer**: QRIS payment, mobile PWA, dan mungkin integrasi marketplace.
