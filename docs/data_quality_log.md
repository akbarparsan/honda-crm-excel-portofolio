# Catatan Kualitas Data

Seluruh masalah di bawah sengaja disisipkan pada data sintetis untuk mensimulasikan kondisi nyata. Setiap rumus deteksi sudah diuji dan menghasilkan jumlah yang tertera. Rumus ditulis dengan pemisah koma. Jika Excel Anda memakai pengaturan regional Indonesia, ganti koma dengan titik koma.

## Dataset A: penjualan dan servis motor

Rumus mengacu ke sheet `Penjualan` kecuali disebut lain.

### A1. Nomor rangka tidak seragam

- **Temuan:** 41 berhuruf kecil, 30 dengan spasi di akhir (panjang 16 karakter).
- **Deteksi:**
  - Huruf kecil: `=SUMPRODUCT(--EXACT(I2:I1001,LOWER(I2:I1001)))`
  - Spasi: `=SUMPRODUCT(--(LEN(I2:I1001)<>15))`
- **Dampak:** nomor rangka adalah kunci penghubung ke Servis_AHASS. Spasi membuat pencocokan gagal pada 28 baris servis. Kepatuhan KPB 1 terhitung 52,6% (419 unit) padahal seharusnya 54,0% (430 unit). Pencocokan di Excel tidak membedakan huruf besar dan kecil, sehingga hanya spasi yang merusak hasil, tetapi huruf kecil tetap perlu diseragamkan.
- **Penanganan:** kolom `Rangka_Bersih` = `UPPER(TRIM(I2))`.
- **Saran di sumber:** ubah huruf menjadi kapital otomatis, batasi panjang 15 karakter, dan gunakan pemindai atau dropdown.

### A2. Tanggal serah terima sebelum tanggal transaksi

- **Temuan:** 10 transaksi, selisih -5 sampai -2 hari.
- **Deteksi:** `=SUMPRODUCT((S2:S1001<>"")*(S2:S1001<B2:B1001))`
- **Dampak:** rata-rata lead time menjadi 7,44 hari, bukan 7,56 hari yang valid (862 transaksi).
- **Penanganan:** kolom `Hari_Serah` memakai `IF(S2="","",S2-B2)`, lalu rata-rata dengan `AVERAGEIFS(..., ">=0")`. Baris ditandai agar bisa dikonfirmasi.
- **Saran di sumber:** validasi tanggal serah tidak boleh lebih awal dari tanggal transaksi.

### A3. Diskon ekstrem

- **Temuan:** 8 transaksi dengan diskon Rp 9 sampai 15 juta (total Rp 96 juta). Enam di antaranya berstatus Selesai (Rp 75 juta). Diskon wajar paling besar Rp 1,5 juta, dan 393 transaksi (39,3%) tanpa diskon.
- **Deteksi:** batas atas IQR sebesar Rp 2,375 juta.
  `=COUNTIF(O2:O1001,">"&(QUARTILE(O2:O1001,3)+1.5*(QUARTILE(O2:O1001,3)-QUARTILE(O2:O1001,1))))`
- **Dampak:** omzet bersih terdistorsi. Ini bisa berupa salah input (kelebihan nol) atau diskon khusus tanpa pencatatan persetujuan.
- **Penanganan:** ditandai dengan conditional formatting (aturan: Diskon_Rp lebih besar dari batas atas IQR), tetap dihitung dalam omzet dengan catatan.
- **Saran di sumber:** persetujuan berjenjang di atas ambang dan audit transaksi yang ditandai.

### A4. Nama konsumen tidak seragam

- **Temuan:** 62 berawalan Bpk./Ibu/Sdr., 53 berspasi di akhir, 17 huruf kapital semua.
- **Deteksi:**
  - Awalan: `=SUMPRODUCT(--((LEFT(C2:C1001,4)="Bpk.")+(LEFT(C2:C1001,4)="Ibu ")+(LEFT(C2:C1001,4)="Sdr.")>0))`
  - Spasi: `=SUMPRODUCT(--(RIGHT(C2:C1001,1)=" "))`
  - Kapital: `=SUMPRODUCT(--EXACT(C2:C1001,UPPER(C2:C1001)))`
- **Penanganan:** `=TRIM(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(PROPER(C2),"Bpk. ",""),"Ibu ",""),"Sdr. ",""))`. `PROPER` harus paling dalam, karena `SUBSTITUTE` membedakan huruf besar dan kecil.

### A5. Transaksi belum final dan tanggal serah kosong

- **Temuan:** 128 baris tanpa tanggal serah: 110 Batal dan 18 Dalam Proses.
- **Deteksi:** `=COUNTBLANK(S2:S1001)`
- **Dampak:** sel kosong dibaca sebagai 0. Jika ikut dihitung, lead time dan kepatuhan menjadi salah.
- **Penanganan:** hanya status Selesai yang dihitung, dan sel kosong tetap kosong (`""`), bukan 0.

### A6. Bulan berjalan belum lengkap

- **Temuan:** Oktober 2026 hanya 8 unit Selesai (data sampai 20 Oktober, sebagian transaksi belum selesai).
- **Deteksi:** `=COUNTIFS(B2:B1001,">="&DATE(2026,10,1),Q2:Q1001,"Selesai")`
- **Dampak:** tren tampak turun tajam padahal bulannya belum lengkap.
- **Penanganan:** Oktober dikeluarkan dari tren dan dari daftar bulan di dropdown.

### A7. Kategori kecil dengan satu model

- **Temuan:** Sport (CB150R) 24 unit dan Trail (CRF150L) 23 unit pada YTD September 2026.
- **Dampak:** persentase perubahan ekstrem (Trail +187,5% dari 8 menjadi 23 unit) dan mudah disalahartikan.
- **Penanganan:** angka dasar selalu ditampilkan bersama persentase, dan badge menampilkan "n/a" bila pembanding bernilai 0.

### A8. Rating layanan tidak lengkap

- **Temuan:** 255 dari 1.270 kunjungan (20,1%) tanpa rating. Rata-rata 4,30 dari 1.015 responden.
- **Deteksi:** `=COUNTBLANK(Servis_AHASS!J2:J1271)`
- **Penanganan:** `AVERAGE` mengabaikan sel kosong, dan jumlah responden dilaporkan bersama rata-rata.

## Dataset B: CRM pelanggan mobil

Rumus mengacu ke sheet `Data_Pelanggan`.

### B1. Baris duplikat

- **Temuan:** 15 baris duplikat penuh (985 pelanggan unik dari 1.000 baris).
- **Deteksi:** `=COUNTA(A2:A1001)-ROUND(SUMPRODUCT(1/COUNTIF(A2:A1001,A2:A1001)),0)`
- **Dampak:** jumlah pelanggan dan total biaya terlalu besar sekitar 1,5%.
- **Penanganan:** Remove Duplicates, atau penghitungan unik dengan `SUMPRODUCT(1/COUNTIF())`.

### B2. Format nomor HP tidak seragam

- **Temuan:** 63 berawalan +62, 127 memakai tanda hubung, 40 tanpa angka 0 di depan (kategori bisa tumpang tindih).
- **Deteksi:**
  - `=COUNTIF(C2:C1001,"+62*")`
  - `=COUNTIF(C2:C1001,"*-*")`
  - `=SUMPRODUCT((LEFT(C2:C1001,1)<>"0")*(LEFT(C2:C1001,1)<>"+"))`
- **Penanganan:** buang spasi, tanda hubung, dan plus dengan `SUBSTITUTE`, lalu seragamkan awalan dengan `IF(LEFT(...,2)="62","0"&MID(...,3,100),...)`.

### B3. Nama pelanggan tidak seragam

- **Temuan:** 64 berspasi di akhir, 44 huruf kecil semua, 28 huruf kapital semua, 20 berspasi ganda.
- **Deteksi:**
  - Spasi akhir: `=SUMPRODUCT(--(RIGHT(B2:B1001,1)=" "))`
  - Kecil: `=SUMPRODUCT(--EXACT(B2:B1001,LOWER(B2:B1001)))`
  - Kapital: `=SUMPRODUCT(--EXACT(B2:B1001,UPPER(B2:B1001)))`
  - Spasi ganda: `=SUMPRODUCT(--ISNUMBER(SEARCH("  ",B2:B1001)))`
- **Penanganan:** `=TRIM(PROPER(B2))`.

### B4. Pelanggan belum pernah servis

- **Temuan:** 129 pelanggan unik (13,1%) tanpa tanggal servis.
- **Deteksi:** `=ROUND(SUMPRODUCT((J2:J1001="")/COUNTIF(A2:A1001,A2:A1001)),0)`
- **Dampak:** sel kosong dibaca sebagai 0 hari dan salah masuk segmen "Aman".
- **Penanganan:** segmen khusus "Belum Pernah Servis", dan kondisi sel kosong diperiksa paling awal dalam rumus segmentasi.

### B5. Rating tidak lengkap

- **Temuan:** 484 dari 985 pelanggan (49,1%) tanpa rating. Rata-rata 4,20 dari 501 responden.
- **Deteksi:** `=ROUND(SUMPRODUCT((P2:P1001="")/COUNTIF(A2:A1001,A2:A1001)),0)`
- **Penanganan:** rata-rata dari responden, dengan jumlah responden dilaporkan.
