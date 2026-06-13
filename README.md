# IAMS - Internal Audit Management System

Aplikasi management audit internal berbasis Serverless Web App memanfaatkan Google Workspace Ecosystem.

## 🚀 Fitur Utama
- **Dashboard Metrik:** Grafik proporsi kasus (Open, Progress, Closed) via Chart.js.
- **LPJ Generator:** Pengklusteran otomatis temuan berdasarkan Entitas Cabang Outlet untuk kemudahan laporan manajemen.
- **Enterprise Security:** Enkripsi kata sandi menggunakan SHA-256 Hashing via Google Cloud Digest.
- **High Performance:** Batch write operations (`setValues`) untuk menghemat kuota I/O Google Sheets.

## 🛠️ Teknologi
- **Frontend:** React.js (SPA), Tailwind CSS, Chart.js, ExcelJS.
- **Backend:** Google Apps Script (GAS) Engine.
- **Database:** Google Sheets DB.

## 📋 Cara Instalasi
1. Salin kode dari folder `backend/Code.js` ke Apps Script Editor.
2. Salin kode dari folder `frontend/Index.html` ke file HTML Apps Script.
3. Jalankan fungsi `setupDatabase` pada GAS Editor untuk inisialisasi tabel.
4. Deploy sebagai Web App dengan akses *Anyone*.
