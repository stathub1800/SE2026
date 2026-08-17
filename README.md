# Struktur Progres SE2026 — BPS Provinsi Lampung

Halaman hirarki kartu untuk memantau progres pendataan SE2026, dibangun ulang setiap hari dari
dua berkas ekspor FASIH.

- Repositori : https://github.com/stathub1800/SE2026
- Folder lokal : `C:\Users\supoy\hirarki_se2026`
- Alamat publik : https://stathub1800.github.io/SE2026/

---

## Isi folder

| Berkas | Fungsi |
|---|---|
| `data\` | tempat menaruh dua berkas ekspor FASIH terbaru |
| `build.py` | membaca berkas ekspor, menguji mutunya, lalu menulis `index.html` |
| `template.html` | rangka halaman (HTML, CSS, JavaScript) dengan penanda `__PAYLOAD__` |
| `index.html` | hasil akhir yang ditayangkan GitHub Pages |
| `update.bat` | jalankan `build.py` lalu commit dan push ke GitHub |
| `pasang_jadwal.bat` | mendaftarkan `update.bat` ke Task Scheduler, sekali saja |
| `log_update.txt` | catatan setiap kali pembaruan berjalan (tidak ikut di-commit) |

---

## Pemasangan pertama kali

Cukup dikerjakan satu kali.

**1. Siapkan Python.** Pastikan Python 3 sudah terpasang dan tercentang *Add Python to PATH*.
Uji di Command Prompt:

```
python --version
pip install pandas openpyxl
```

**2. Siapkan folder dan repositori.**

```
cd C:\Users\supoy
git clone https://github.com/stathub1800/SE2026.git hirarki_se2026
cd hirarki_se2026
```

Kalau folder `hirarki_se2026` sudah berisi berkas dari paket ini dan belum terhubung ke GitHub:

```
cd C:\Users\supoy\hirarki_se2026
git init
git branch -M main
git remote add origin https://github.com/stathub1800/SE2026.git
git add -A
git commit -m "Halaman hirarki SE2026"
git push -u origin main
```

**3. Simpan kredensial Git** supaya `update.bat` tidak meminta kata sandi saat berjalan otomatis.
Cara paling sederhana, sekali push manual dengan Git Credential Manager aktif:

```
git config --global credential.helper manager
git push origin main
```

Masukkan username GitHub dan **Personal Access Token** (bukan kata sandi akun). Token dibuat di
GitHub → Settings → Developer settings → Personal access tokens → Tokens (classic), centang
lingkup `repo`. Setelah sekali berhasil, kredensial tersimpan dan tugas terjadwal bisa push sendiri.

**4. Nyalakan GitHub Pages.** Di halaman repositori: Settings → Pages → Source `Deploy from a branch`
→ Branch `main`, folder `/ (root)` → Save.

**5. Pasang jadwal harian.** Klik kanan `pasang_jadwal.bat` → **Run as administrator**.
Tugas akan berjalan setiap hari pukul 08.15.

---

## Pemakaian harian

Hanya satu langkah manual: **timpa dua berkas di folder `data\`** dengan ekspor FASIH terbaru.

Nama berkas tidak harus persis sama. `build.py` mencari dengan pola:

- `*Pendataan*.xlsx` → Export Progres Pendataan Kabupaten/Kota
- `*Pemutakhiran*Keluarga*.xlsx` → Export Progres Pemutakhiran Keluarga Kabupaten/Kota

Kalau ada lebih dari satu berkas yang cocok, yang dipakai adalah **yang paling baru**.

Sisanya berjalan sendiri pukul 08.15. Untuk memicu langsung tanpa menunggu, klik dua kali
`update.bat`, atau:

```
schtasks /Run /TN "SE2026 Hirarki - Update Harian"
```

Waktu tarikan data yang tampil di pojok kanan atas halaman diambil dari **waktu ubah berkas Excel**,
jadi otomatis ikut menyesuaikan.

---

## Uji mutu otomatis

Setiap kali dibangun, `build.py` memeriksa dan menampilkan hasilnya di layar dan di `log_update.txt`:

1. **Enam identitas penjumlahan** — memastikan angka anak benar-benar menjumlah ke angka induk
   di provinsi maupun 15 kabupaten/kota.
2. **Tiga belas persentase resmi FASIH** — memastikan persen yang disalin dari berkas sama dengan
   hasil hitung ulang nilai dibagi penyebutnya.

Kalau ada yang tidak lulus, halaman tetap dibangun tetapi muncul baris `PERINGATAN`. Periksa dulu
berkas ekspornya sebelum dipakai rapat.

---

## Mengubah isi halaman

**Mengganti nama kartu.** Buka `build.py`, cari daftar `N = [` di bagian bawah, ubah nilai `label`:

```python
dict(id='capaian', parent='assign', label='Verifikasi', ...)
```

Jangan ubah `id` — itu kunci yang menyambungkan kartu satu sama lain. Kalau nama kartu itu disebut
di dalam teks `rumus`, `ket`, atau `turunan` kartu lain, sesuaikan juga di situ.

**Mengubah tampilan.** Warna, ukuran kartu, dan tata letak ada di blok `<style>` pada
`template.html`. Ukuran kartu mengikuti lebar layar lewat variabel `--wk` (kartu induk) dan
`--wm` (kartu rincian) pada beberapa `@media`.

Setelah mengubah apa pun, jalankan `python build.py` untuk melihat hasilnya di `index.html`.

---

## Kalau bermasalah

| Gejala | Penyebab yang paling sering |
|---|---|
| `GAGAL: berkas ... tidak ditemukan` | berkas belum ditaruh di `data\`, atau namanya tidak memuat kata *Pendataan* / *Pemutakhiran Keluarga* |
| `GAGAL push ke GitHub` | kredensial belum tersimpan, atau token sudah kedaluwarsa |
| Halaman GitHub Pages belum berubah | tunggu satu sampai dua menit, lalu muat ulang dengan Ctrl+F5 |
| Angka tidak berubah padahal berkas baru | berkas lama masih ada di `data\` dan waktu ubahnya lebih baru; hapus yang lama |
| Muncul `PERINGATAN` di log | ada identitas atau persentase yang tidak cocok, periksa berkas ekspornya |

Riwayat lengkap ada di `log_update.txt`.
