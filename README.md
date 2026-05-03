# Dashboard Monitoring SE2026 – BPS Provinsi Lampung

## Struktur File
```
dashboard-se2026/
├── index.html                  ← Entry point utama
├── README.md                   ← Dokumentasi ini
└── assets/
    ├── css/
    │   ├── main.css            ← Variables, reset, layout, grid
    │   ├── sidebar.css         ← Navigasi sidebar
    │   ├── cards.css           ← Card components, KPI, tabel
    │   └── charts.css          ← Bar chart, donut, progress bar
    └── js/
        ├── config.js           ← ⚙️ EDIT DI SINI: setting, kabkot list, warna
        ├── data.js             ← 🔌 EDIT DI SINI: koneksi API / Supabase
        ├── utils.js            ← Fungsi helper (format angka, rupiah, dsb)
        ├── map.js              ← Leaflet map helper
        ├── charts.js           ← Chart helper (placeholder)
        ├── app.js              ← 🚀 Router & bootstrap utama
        └── pages/
            ├── overview.js     ← Halaman Overview
            ├── progres.js      ← 1. Progres Pendataan
            ├── kinerja.js      ← 2. Kinerja Petugas
            └── other-pages.js  ← 3-8. Kualitas, Anggaran, Komunikasi,
                                        Laporan, KKD, Kontrak Kinerja
```

## Cara Menggunakan
1. Buka `index.html` di browser (atau deploy ke GitHub Pages / server)
2. Data otomatis dimuat dari `assets/js/data.js`

## Cara Menghubungkan ke Data Nyata (Supabase)
Edit file `assets/js/data.js`, ganti setiap fungsi `fetchXxx()` dengan query Supabase:

```js
// Contoh dengan Supabase:
async fetchProgres() {
  const { data, error } = await supabase
    .from('se2026_progres')
    .select('*')
    .eq('provinsi', 'LAMPUNG');
  return data;
}
```

## Cara Menambah Kabupaten/Kota
Edit `assets/js/config.js` bagian `KABKOT`:
```js
{ id: 'lp16', name: 'Nama Baru', lat: -4.xxx, lng: 105.xxx },
```

## Cara Menambah Halaman Baru
1. Tambahkan `<section class="page" id="page-namahalaman">` di `index.html`
2. Tambahkan nav item di sidebar
3. Buat file `assets/js/pages/namahalaman.js` dengan struktur:
   ```js
   const PageNamaHalaman = {
     async render() { ... }
   };
   ```
4. Daftarkan di `App._pages` di `app.js`

## Auto-Refresh
Atur interval di `config.js`:
```js
REFRESH_INTERVAL: 5 * 60 * 1000, // 5 menit (set 0 untuk nonaktif)
```
