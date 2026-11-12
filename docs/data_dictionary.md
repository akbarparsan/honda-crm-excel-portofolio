# Kamus Data

Seluruh data bersifat sintetis. Harga model di tabel master hanya ilustrasi.

## Dataset A: penjualan dan servis motor

File: `data/Honda_Motor_Sales_Service_Dataset.xlsx` (juga tersedia sebagai CSV di `data/csv/`).

### Penjualan (1.000 baris, 1 baris = 1 transaksi)

| Kolom | Tipe | Keterangan |
|---|---|---|
| No_Transaksi | Teks | Kode transaksi, format TRX-00001 |
| Tanggal_Transaksi | Tanggal | Tanggal pemesanan |
| Nama_Konsumen | Teks | Nama konsumen (penulisan tidak seragam) |
| No_HP | Teks | Nomor telepon, disimpan sebagai teks agar angka 0 di depan tidak hilang |
| Kota | Teks | Kota domisili |
| Cabang | Teks | Cabang dealer (Pusat, Utara, Selatan, Timur, Barat) |
| Salesperson | Teks | Sales yang menjual (acuan: Master_Sales) |
| Model_Motor | Teks | Model yang dibeli (acuan: Master_Model) |
| No_Rangka | Teks | 15 karakter: `MH1` + Kode_Model (3) + kode tahun (1; S = 2025, T = 2026) + nomor urut (8) |
| Warna | Teks | Warna unit |
| Metode_Bayar | Teks | Cash atau Kredit Leasing |
| Leasing | Teks | Perusahaan leasing (kosong untuk cash) |
| Tenor_Bulan | Angka | Tenor kredit (12, 18, 24, 30, 35; kosong untuk cash) |
| DP_Rp | Angka | Uang muka (kosong untuk cash) |
| Diskon_Rp | Angka | Potongan harga |
| Status_Kredit | Teks | Disetujui, Ditolak, atau Dalam Proses (kosong untuk cash) |
| Status_Transaksi | Teks | Selesai, Batal, atau Dalam Proses |
| Sumber_Prospek | Teks | Walk-in, Pameran, Media sosial, Referral, Telemarketing |
| Tanggal_Serah_Terima | Tanggal | Tanggal unit diserahkan (kosong bila belum selesai) |

Kolom bantu pada `Honda_Motor_KPI_Dashboard.xlsx` (sheet Penjualan):

| Kolom | Rumus ringkas | Keterangan |
|---|---|---|
| T Kategori | `INDEX/MATCH` ke Master_Model | Matic, Bebek, Sport, atau Trail |
| U Nilai_Bersih | harga dari Master_Model dikurangi Diskon_Rp | Nilai per transaksi |
| V Bulan | `DATE(YEAR(B),MONTH(B),1)` | Tanggal pertama bulan transaksi |
| W Rangka_Bersih | `UPPER(TRIM(I))` | Kunci penghubung ke Servis_AHASS |
| X Eligible_KPB1 | `IF(AND(Selesai, ada tanggal serah, serah lebih dari 60 hari lalu),"Ya","Tidak")` | Layak dihitung kepatuhan KPB 1 |
| Y Sudah_KPB1 | `COUNTIFS` ke Servis_AHASS | "Sudah" atau "Belum" (kosong bila tidak eligible) |

### Servis_AHASS (1.270 baris, 1 baris = 1 kunjungan)

| Kolom | Tipe | Keterangan |
|---|---|---|
| No_Servis | Teks | Kode servis, format SVC-00001 |
| Tanggal_Servis | Tanggal | Tanggal kunjungan |
| No_Rangka | Teks | Format bersih (huruf besar, tanpa spasi) |
| Jenis_Servis | Teks | KPB 1, KPB 2, KPB 3, KPB 4, Servis Rutin, Ganti Oli, Perbaikan |
| KM_Servis | Angka | Odometer saat servis |
| Biaya_Jasa | Angka | Biaya jasa (Rp); KPB bernilai 0 |
| Biaya_Sparepart | Angka | Biaya suku cadang (Rp) |
| Mekanik | Teks | Nama mekanik |
| Durasi_Menit | Angka | Lama pengerjaan |
| Rating_Layanan | Angka | Skala 1 sampai 5 (bisa kosong) |

Kolom bantu pada dashboard: **K Cabang**, diambil dari Penjualan lewat `Rangka_Bersih`.

### Tabel master

| Tabel | Kolom |
|---|---|
| Master_Model | Model_Motor, Kode_Model, Kategori, Kapasitas_CC, Harga_OTR (Rp, ilustrasi) |
| Master_Sales | Salesperson, Cabang, Target_Unit_Bulanan |
| Master_Bonus | Unit_Minimal, Bonus_per_Unit (Rp); dipakai dengan pencocokan perkiraan (`VLOOKUP` dengan `TRUE`) |

## Dataset B: CRM pelanggan mobil

File: `data/Honda_Car_CRM_Dataset.xlsx`, sheet `Data_Pelanggan` (1.000 baris, 985 pelanggan unik).

| Kolom | Tipe | Keterangan |
|---|---|---|
| ID_Pelanggan | Teks | Kode pelanggan (duplikat pada 15 baris) |
| Nama_Pelanggan | Teks | Nama (penulisan tidak seragam) |
| No_HP | Teks | Nomor telepon (format tidak seragam) |
| Kota | Teks | Kota domisili |
| No_Polisi | Teks | Nomor polisi kendaraan |
| Model_Mobil | Teks | Brio, Mobilio, BR-V, WR-V, HR-V, CR-V, City Hatchback, Civic |
| Tanggal_Beli | Tanggal | Tanggal pembelian |
| Tipe_Pembayaran | Teks | Cash atau Kredit Leasing |
| Salesperson | Teks | Sales yang menjual |
| Tanggal_Servis_Terakhir | Tanggal | Kosong bila belum pernah servis |
| KM_Terakhir | Angka | Odometer saat servis terakhir |
| Jenis_Servis_Terakhir | Teks | Servis Berkala, Ganti Oli, General Repair, Body Repair |
| Total_Biaya_Servis | Angka | Biaya servis terakhir (Rp) |
| Jumlah_Kunjungan_Servis | Angka | Total kunjungan servis sejak pembelian |
| Status_Follow_Up | Teks | Sudah Dihubungi, Belum Dihubungi, Tidak Diangkat, Janji Servis |
| Rating_Kepuasan | Angka | Skala 1 sampai 5 (bisa kosong) |

Segmentasi follow-up (berdasarkan hari sejak servis terakhir terhadap tanggal acuan 19 Oktober 2026): lebih dari 180 hari = Segera Hubungi, 90 sampai 180 hari = Pantau, kurang dari 90 hari = Aman, tanggal kosong = Belum Pernah Servis.
