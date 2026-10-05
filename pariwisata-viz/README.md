# Pintu Masuk Nusantara — kode awal

Visualisasi interaktif pariwisata Indonesia dari data BPS. Tahap ini memakai 3 berkas yang sudah ada; peta dan topik lain menyusul.

## Isi proyek

```
pariwisata-viz/
├── data/raw/                  # taruh 3 berkas xlsx BPS di sini
├── preprocess/preprocess.py   # xlsx -> JSON kecil
└── web/
    ├── index.html             # halaman utama (D3.js)
    └── data/*.json            # hasil pra-proses (sudah disertakan)
```

## Cara menjalankan

1. Pasang Python 3.9+ dan paket `openpyxl`:
   ```bash
   pip install openpyxl
   ```
2. Salin tiga berkas xlsx ke `data/raw/`.
3. (Opsional, hanya jika data berubah) jalankan pra-proses dari folder proyek:
   ```bash
   python preprocess/preprocess.py
   ```
   Skrip mencetak hasil cek konsistensi (jumlah rincian vs total BPS).
4. Jalankan server lokal. Klik dua kali `index.html` **tidak bisa**, karena browser memblokir pemuatan JSON dari `file://`:
   ```bash
   cd web
   python -m http.server 8000
   ```
5. Buka `http://localhost:8000`. Perlu internet saat pertama kali karena D3.js dimuat dari cdnjs.

Sebelum deploy, unduh `d3.min.js` ke `web/lib/` dan ubah tag `<script src=...>` agar tidak bergantung pada CDN.

## Tampilan yang sudah ada

| No | Tampilan | Topik | Interaksi |
|---|---|---|---|
| 1 | Tren wisman per pintu masuk (garis) | — | pilih dua seri, tooltip |
| 2 | Treemap kawasan → negara, atau jenis pintu → pintu | Hierarki | drill-down, breadcrumb, filter tahun dan dimensi; ukuran = kunjungan, warna = perubahan tahunan |
| 3 | Selisih perjalanan wisnus tujuan − asal per provinsi | Aliran (bersih) | filter tahun dan ukuran, tooltip, sorot provinsi |
| 4 | TPK vs lama menginap (sebar) | Multivariat (awal) | filter tahun dan jenis tamu, tooltip, sorot provinsi tertaut ke tampilan 3 |

## Catatan kualitas data (hasil cek otomatis)

- **Pintu masuk dan wisnus:** rincian cocok 100% dengan total.
- **Paspor:** untuk kawasan selain ASEAN, jumlah negara kadang berbeda dari total kawasan di tabel BPS: seluruh kawasan pada **Juli 2026** (rincian lebih besar sekitar 7-16%) dan kawasan **Asia di luar ASEAN** pada banyak bulan sebelumnya. Bulan 2026-01 dan 2026-08 cocok persis. Treemap memakai jumlah rincian, jadi bisa sedikit berbeda dari angka total BPS. Cek ulang tabel sumber atau rilis BPS untuk bulan tersebut sebelum dikutip di laporan.
- Baris "Other ..." pada tabel paspor adalah agregat negara kecil; negara di bawahnya dilewati supaya tidak terhitung dua kali.
- Baris "Indonesia" pada paspor ikut terhitung di ASEAN; tentukan apakah akan dikeluarkan dari analisis wisman.
- Data bulanan 2026 baru sampai Agustus (September-Desember kosong, otomatis diabaikan). TPK dan lama menginap hanya tahunan 2021-2025.
- Tahun 2021 dipengaruhi pandemi, sehingga pertumbuhan 2022 sangat besar dan warnanya jenuh (dibatasi ±100%).

## Data yang masih kurang

| Kebutuhan | Status | Keterangan |
|---|---|---|
| **Batas provinsi (GeoJSON)** | Belum ada | Untuk peta provinsi (38). Perlu kolom kode wilayah BPS/Kemendagri agar bisa digabung dengan data; nama provinsi di data sekarang huruf kapital dan harus dipetakan ke kode |
| **Batas kabupaten/kota (GeoJSON, ± 500 unit)** | Belum ada | Syarat peta kab/kota. Sederhanakan geometri (mapshaper/TopoJSON) agar ringan |
| **Indikator pariwisata kab/kota** | Belum ada | Mis. jumlah akomodasi, dari publikasi *Kabupaten/Kota dalam Angka* atau PODES. Tanpa ini peta kab/kota hanya bisa memakai indikator umum |
| **Penduduk dan luas wilayah** | Belum ada | Untuk rasio (wisnus per kapita, akomodasi per 100.000 penduduk, kepadatan). Peta koroplet tidak boleh memakai angka absolut |
| **Variabel tambahan multivariat** | Sebagian | Baru ada TPK dan lama menginap (3 jenis) serta wisnus. Syarat minimal 8 variabel: tambah IPM, kemiskinan, TPT, PDRB per kapita, pangsa PDRB akomodasi dan makan minum, kepadatan, pengeluaran per kapita |
| **Silang kebangsaan × pintu masuk** | Tidak ada | Walau nama berkasnya begitu, sheet 1 hanya per pintu masuk dan sheet 2 hanya per paspor. Tanpa tabel silang, diagram Sankey/chord asal-tujuan tidak bisa dibuat dengan jujur |
| **Matriks asal-tujuan wisnus** | Tidak ada | Data hanya total per provinsi asal dan per provinsi tujuan, belum pasangan asal-tujuan. Karena itu tampilan 3 hanya menunjukkan selisih |

## Langkah berikutnya yang disarankan

1. Cari tabel silang kebangsaan × pintu masuk (atau matriks OD wisnus) untuk topik aliran.
2. Unduh GeoJSON provinsi + kab/kota dan siapkan tabel kode wilayah.
3. Kumpulkan indikator provinsi untuk PCA, parallel coordinates, dan heatmap terklaster.
4. Tambahkan representasi hierarki kedua (sunburst) pada tampilan 2.
5. Bungkus dalam scrollytelling dan deploy ke GitHub Pages.
