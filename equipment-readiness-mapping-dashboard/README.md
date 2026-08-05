# Equipment Readiness Mapping Dashboard

Dashboard interaktif untuk memantau kesiapan alat (Shore Crane, Yard Crane, dan Internal Transfer Vehicle) di seluruh terminal Pelindo Terminal Petikemas (TPK) secara real-time, terpusat, dan otomatis — menggantikan proses pemantauan manual berbasis laporan harian.

🔗 **Live Dashboard:** [Equipment Readiness Mapping](https://datastudio.google.com/reporting/8e481f29-61ce-4e2a-88bc-0fc7b9632665)

## Latar Belakang

Sebelumnya, pemantauan kesiapan alat seperti:
- **SC** — Shore Crane
- **YC** — Yard Crane
- **ITV** — Internal Transfer Vehicle

dilakukan secara manual melalui laporan harian tiap terminal, sehingga rawan *human error* dan ketidaksamaan data antar terminal.

Dashboard ini dibangun untuk mengatasi hal tersebut dengan alur otomatis:
1. **Scraping** data dari *Monthly Terminal Report* menggunakan aplikasi scraping (dibuat oleh Tim Magang Divisi PPO)
2. **Import & penggabungan data** melalui Google Apps Script
3. **Visualisasi interaktif** di Google Looker Studio

## Fitur Utama

- **Peta status terminal** (Ready / Struggling / Crisis) di seluruh Indonesia
- **Filter** berdasarkan Terminal dan Tanggal (Day)
- **Status kondisi alat R/C/B**
  - 🟢 **R — Ready**: alat siap beroperasi tanpa kendala
  - 🟡 **C — Dengan Catatan**: alat masih bisa beroperasi namun ada catatan
  - 🔴 **B — Breakdown**: alat tidak dapat dioperasikan
- **Equipment Ratio (Condition)** — rasio kesiapan alat berbasis unit R+C, dibandingkan terhadap jumlah SC
- **Equipment Ratio (Terminal)** — rasio seluruh alat (R+C+B) per kategori
- **Equipment Condition Proportion** — persentase %R, %C, %B per terminal/hari
- **Terminal Status** — klasifikasi otomatis berdasarkan threshold rasio SC:YC:ITV = 1:3:5 (per November 2025)
- Halaman detail per kategori alat: **SC Readiness Status**, **YC Readiness Status**, **ITV Readiness Status**

## Klasifikasi Status Terminal

| Status | Kriteria |
|---|---|
| 🟢 Ready | Rasio alat R+C memenuhi/melampaui SC:YC:ITV ≥ 1:3:5 |
| 🟡 Struggling | Rasio mendekati threshold, namun salah satu jenis alat di bawah rasio minimal |
| 🔴 Crisis | Rasio R+C jauh di bawah threshold 1:3:5 |

## Alur Update Data

1. **Scraping Data** — jalankan aplikasi scraping "Ambil Data [Internet ON]", input bulan & tahun, generate file Excel
2. **Import Data** — upload hasil scraping ke spreadsheet `[UPLOAD DATA]`
3. **Penggabungan Data** — jalankan tombol `UPDATE` di spreadsheet `[UPDATE DASHBOARD]` untuk memproses via Google Apps Script
4. **Segarkan Data** — refresh data di Looker Studio agar seluruh visual otomatis terbarui

## Tech Stack

- Google Apps Script (pemrosesan & penggabungan data)
- Google Sheets (staging data: upload & update)
- Google Looker Studio (visualisasi dashboard)
- Aplikasi scraping internal (Python/desktop app — Tim Magang Divisi PPO)

## Dokumentasi

Panduan lengkap penggunaan dashboard (navigasi, filter, rumus rasio, dan langkah update data) tersedia di [`Guidebook_Mapping_Kesiapan_Alat.pdf`](./Guidebook_Mapping_Kesiapan_Alat.pdf).

---
*Project ini dikembangkan sebagai bagian dari magang di PT Pelindo Terminal Petikemas (Divisi PPO).*
