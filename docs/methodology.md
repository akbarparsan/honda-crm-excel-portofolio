# Metodologi

Dokumen ini menjelaskan model data, langkah re-modeling, rumus utama, dan cara validasi. Rumus ditulis dengan pemisah koma. Jika Excel Anda memakai pengaturan regional Indonesia, ganti koma dengan titik koma.

## 1. Model data

| Jenis | Tabel | Peran |
|---|---|---|
| Fakta | Penjualan | 1 baris = 1 transaksi |
| Fakta | Servis_AHASS | 1 baris = 1 kunjungan servis |
| Dimensi | Master_Model | Kode, kategori, dan harga per model |
| Dimensi | Master_Sales, Master_Bonus | Cabang, target, dan tangga bonus |

Relasi:

- `Penjualan.Model_Motor` ke `Master_Model.Model_Motor` (banyak ke satu).
- `Servis_AHASS.No_Rangka` ke `Penjualan.No_Rangka` (banyak ke satu, setelah nomor rangka dirapikan).

Satu motor bisa memiliki beberapa baris servis. Karena itu status "sudah KPB 1" dicek dengan `COUNTIFS` yang hanya menandai ada atau tidaknya baris KPB 1, sehingga jumlah transaksi tidak berlipat ganda.

## 2. Langkah re-modeling

### Sheet Penjualan (kolom bantu T sampai Y)

| Kolom | Rumus |
|---|---|
| Kategori | `=INDEX(Master_Model!$C$2:$C$13,MATCH(H2,Master_Model!$A$2:$A$13,0))` |
| Nilai_Bersih | `=INDEX(Master_Model!$E$2:$E$13,MATCH(H2,Master_Model!$A$2:$A$13,0))-O2` |
| Bulan | `=DATE(YEAR(B2),MONTH(B2),1)` |
| Rangka_Bersih | `=UPPER(TRIM(I2))` |
| Eligible_KPB1 | `=IF(AND(Q2="Selesai",ISNUMBER(S2),S2<Petunjuk!$B$3-60),"Ya","Tidak")` |
| Sudah_KPB1 | `=IF(X2="Ya",IF(COUNTIFS(Servis_AHASS!$C$2:$C$1271,W2,Servis_AHASS!$D$2:$D$1271,"KPB 1")>0,"Sudah","Belum"),"")` |

### Sheet Servis_AHASS (kolom bantu K)

| Kolom | Rumus |
|---|---|
| Cabang | `=IFERROR(INDEX(Penjualan!$F$2:$F$1001,MATCH(C2,Penjualan!$W$2:$W$1001,0)),"")` |

### Lapisan parameter (sheet KPI_Data)

| Sel | Fungsi | Rumus |
|---|---|---|
| B4 | Bulan terpilih | Mengambil nilai dropdown di KPI_Dashboard |
| B5 | Cabang terpilih | Mengambil nilai dropdown di KPI_Dashboard |
| B6 | Kriteria cabang | `=IF(B5="Semua Cabang","*",B5)` |
| B7 | Awal YTD | `=DATE(YEAR(B4),1,1)` |
| B8 | Awal YTD tahun lalu | `=EDATE(B7,-12)` |
| B9 | Bulan sama tahun lalu | `=EDATE(B4,-12)` |
| B10 | Bulan sebelumnya | `=EDATE(B4,-1)` |

Tanda bintang `*` pada kriteria `COUNTIFS` cocok dengan teks apa pun, sehingga satu rumus melayani pilihan "Semua Cabang" maupun satu cabang.

## 3. Perhitungan KPI

| KPI | Rumus (kategori di A14, kriteria cabang di B6) |
|---|---|
| Unit YTD | `=COUNTIFS(Penjualan!$T$2:$T$1001,$A14,Penjualan!$F$2:$F$1001,$B$6,Penjualan!$Q$2:$Q$1001,"Selesai",Penjualan!$V$2:$V$1001,">="&$B$7,Penjualan!$V$2:$V$1001,"<="&$B$4)` |
| Omzet YTD | `SUMIFS` pada `Penjualan!$U$2:$U$1001` dengan kriteria yang sama |
| Porsi unit | `=IF(B$18=0,0,B14/B$18)` |
| YTD vs PY | `=IF(C14=0,"n/a",B14/C14-1)` |
| Unit bulan ini | `COUNTIFS` dengan `Penjualan!$V$2:$V$1001` sama dengan `$B$4` |
| Unit tahun lalu dan bulan lalu | Sama, dengan `$B$9` dan `$B$10` |
| Kepatuhan KPB 1 | `=IF(B44=0,0,B45/B44)` |

Eligible dan sudah KPB 1 (mengikuti filter cabang):

```
=COUNTIFS(Penjualan!$X$2:$X$1001,"Ya",Penjualan!$F$2:$F$1001,$B$6)
=COUNTIFS(Penjualan!$Y$2:$Y$1001,"Sudah",Penjualan!$F$2:$F$1001,$B$6)
```

Tahap KPB:

```
=COUNTIFS(Servis_AHASS!$D$2:$D$1271,A39,Servis_AHASS!$K$2:$K$1271,$B$6)
```

## 4. Rumus penting lainnya

| Tujuan | Rumus |
|---|---|
| Pelanggan unik dari data berduplikat | `=ROUND(SUMPRODUCT(1/COUNTIF(A2:A1001,A2:A1001)),0)` |
| Segmentasi follow-up | `=IF(J2="","Belum Pernah Servis",IF(Q2>180,"Segera Hubungi",IF(Q2>=90,"Pantau","Aman")))` |
| Rata-rata lead time valid | `=AVERAGEIFS(AB2:AB1001,AB2:AB1001,">=0")` |
| Bonus bertingkat | `=VLOOKUP(C4,Master_Bonus!$A$2:$B$6,2,TRUE)` |
| Peringkat | `=RANK(C4,$C$4:$C$13,0)` |
| Kode model dari rangka | `=UPPER(MID(TRIM(I2),4,3))` |
| Judul dinamis | `="Dari "&B39&" unit ber-KPB 1, tinggal "&B42&" tuntas KPB 4"` |

## 5. Validasi

1. **Pencocokan silang.** Semua KPI utama dihitung ulang secara terpisah dengan Python (pandas) dari data sumber, lalu dibandingkan dengan nilai di Excel. Hasilnya sama untuk tampilan awal (September 2026, semua cabang), misalnya 284, 38, 24, dan 23 unit YTD pada empat kategori, serta kepatuhan KPB 1 sebesar 430 dari 797 unit.
2. **Uji filter.** Dashboard diuji dengan Cabang Pusat dan bulan Maret 2026. Unit per kategori, jumlah tiap tahap KPB (125, 65, 29, 7), dan kepatuhan KPB 1 (121 dari 226 unit) sama dengan perhitungan pembanding.
3. **Uji rumus deteksi.** Seluruh rumus deteksi masalah data di `data_quality_log.md` dihitung ulang dan menghasilkan jumlah yang tercantum.
4. **Pemeriksaan galat.** Tidak ada nilai `#N/A`, `#VALUE!`, `#REF!`, atau `#NAME?` pada seluruh sheet.

## 6. Kompatibilitas

- Dashboard hanya memakai fungsi yang tersedia di Excel 2016 atau yang lebih baru (`COUNTIFS`, `SUMIFS`, `INDEX`, `MATCH`, `EDATE`, `EOMONTH`, `SUMPRODUCT`, `IFERROR`). Tidak ada fungsi khusus Microsoft 365.
- Tampilan terbaik di Microsoft Excel desktop. Piramida 3-D, tangkai lollipop dari error bars, dan label gelembung dapat berbeda di Google Sheets atau LibreOffice.
- Dashboard memuat sekitar 7.700 rumus. Jika terasa lambat, ubah **Calculation Options** ke Manual sementara waktu dan tekan F9 untuk menghitung ulang.
