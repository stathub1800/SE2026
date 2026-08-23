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
| `data\` | tempat menaruh berkas ekspor FASIH terbaru |
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

Hanya satu langkah manual: **taruh berkas ekspor FASIH terbaru di folder `data\`**.

Nama berkas tidak harus persis sama. `build.py` mencari dengan pola:

| Pola | Berkas | Wajib |
|---|---|---|
| `*Pendataan*.xlsx` | Export Progres Pendataan Kabupaten/Kota | ya |
| `*Pemutakhiran*Keluarga*.xlsx` | Export Progres Pemutakhiran Keluarga Kabupaten/Kota | ya |
| `*Radar*Anomali*Usaha*.xlsx` | Radar Anomali Usaha | tidak |
| `*Radar*Anomali*Keluarga*.xlsx` | Radar Anomali Keluarga | tidak |
| `*Monitoring*SLS*.xlsx` | Export Monitoring SLS Kabupaten/Kota | tidak |

Kalau ada lebih dari satu berkas yang cocok, yang dipakai adalah **yang paling baru**, jadi berkas
radar yang namanya berawalan tanggal boleh menumpuk. Rapikan sesekali agar folder tidak membengkak.

Berkas yang tidak wajib boleh tidak ada. Tab yang bersangkutan akan disembunyikan sendiri.

## Tiga tab

| Tab | Isi | Bentuk |
|---|---|---|
| Progres Pendataan | hirarki assignment sampai rincian usaha dan keluarga | bagan hirarki kartu |
| Radar Anomali | 8 anomali usaha dan 7 anomali keluarga | papan tindak lanjut, diurutkan menurut sisa terbanyak |
| Monitoring SLS | target dan penyelesaian SLS | ringkasan tiga angka dan peringkat 15 kabupaten/kota |

Pilihan wilayah di bagian atas berlaku untuk ketiga tab sekaligus.

**Menamai anomali.** Nama dan keterangan tiap anomali diatur pada `NAMA_ANOMALI` di dalam
`build.py`. Anomali usaha sudah terisi berdasarkan daftar anomali yang disusun Pokja TI;
**cocokkan urutannya dengan radar FASIH** dan perbaiki bila berbeda. Anomali keluarga belum
bernama, jadi masih tampil sebagai "Anomali 1" sampai "Anomali 7". Setiap kali dibangun,
`build.py` mengingatkan anomali mana saja yang belum bernama.

Sisanya berjalan sendiri pukul 08.15. Untuk memicu langsung tanpa menunggu, klik dua kali
`update.bat`, atau:

```
schtasks /Run /TN "SE2026 Hirarki - Update Harian"
```

Waktu tarikan data yang tampil di pojok kanan atas halaman diambil dari **waktu ubah berkas Excel**,
jadi otomatis ikut menyesuaikan.

---

## Tahan terhadap perubahan struktur ekspor

Kolom dicari berdasarkan **judul kolom**, bukan nomor urutnya. Kalau FASIH menyisipkan kolom baru
di tengah tabel, halaman tetap terbangun dengan benar. Setiap kali dijalankan, `build.py`
menampilkan **Peta kolom yang terbaca** — daftar posisi setiap kolom yang dipakai. Bandingkan
dengan log kemarin kalau curiga ada yang berubah.

Kalau sebuah kolom benar-benar tidak ditemukan atau namanya jadi ambigu, `build.py` **berhenti**
dan mencetak seluruh judul kolom yang terbaca, sehingga `index.html` lama tidak tertimpa oleh
angka yang salah.

Kolom `Jumlah Prelist Belum Didata (Open)` yang muncul pada ekspor Agustus 2026 dipakai langsung
bila tersedia, lengkap dengan rincian UB, UM, UMK, dan keluarga. Bila berkas belum memuatnya,
Open dihitung sebagai sisa seperti sebelumnya.

## Uji mutu otomatis

Setiap kali dibangun, `build.py` memeriksa dan menampilkan hasilnya di layar dan di `log_update.txt`:

1. **Tujuh identitas penjumlahan** — memastikan angka anak benar-benar menjumlah ke angka induk
   di provinsi maupun 15 kabupaten/kota.
2. **Tiga belas persentase resmi FASIH** — memastikan persen yang disalin dari berkas sama dengan
   hasil hitung ulang nilai dibagi penyebutnya.
3. **Cek silang Total Usaha BKU antar lembar** — angka ini muncul di lembar Skala Usaha dan
   lembar Usaha/Perusahaan. Bila keduanya berbeda, biasanya karena kedua lembar ditarik pada waktu
   berbeda. Halaman memakai angka **lembar Skala Usaha** karena itu yang tampil di layar FASIH,
   dan selisihnya diumumkan di kaki halaman agar tidak menjadi pertanyaan saat rapat.
4. **Catatan Open negatif** — FASIH kadang mengeluarkan Open bernilai negatif untuk kabupaten yang
   assignment terkerjakannya melampaui prelist. Angkanya ditampilkan apa adanya, tetapi kabupaten
   mana saja yang terkena dicatat di log agar siap dijelaskan bila ditanya pimpinan.

Kalau ada yang tidak lulus, halaman tetap dibangun tetapi muncul baris `PERINGATAN`. Periksa dulu
berkas ekspornya sebelum dipakai rapat.

---

## Konsep dan definisi

Sejak ekspor 23 Agustus 2026, FASIH menuliskan konsep dan definisi di bawah tabel pada berkas
ekspornya. `build.py` **membaca blok itu otomatis** dan menampilkannya di popup kartu terkait,
sehingga bila BPS memperbarui kondef, halaman ikut memperbarui diri tanpa perlu diedit.

Keterangan "Apa yang dihitung" yang ditulis sendiri tetap mengikuti kondef resmi:

| Istilah | Inti definisi |
|---|---|
| Prelist berjalan | Jumlah prelist terkini di FASIH, dinamis mengikuti perpindahan wilayah assignment |
| Hasil Verifikasi | Assignment usaha dan keluarga yang tidak berstatus Open atau Draft, mencakup seluruh status keberadaan |
| Responden Hasil Pendataan | Responden usaha (BKU) dan keluarga yang berhasil didata dengan keberadaan ditemukan, baru, atau force submit |
| UB, UM, UMK, Usaha Keluarga | Usaha yang berhasil didata dari responden berstatus keberadaan ditemukan atau baru |
| UMK* | UMK ditambah unit pembantu atau penunjang yang nilai pendapatannya nol sehingga skalanya tidak dapat diklasifikasikan |
| Total Keluarga Hasil Pendataan | Keluarga ditemukan ditambah keluarga baru |
| Keluarga Khusus | Keluarga yang tinggal di barak, pondok pesantren, hunian sementara, pengungsian, dan sejenisnya |

Kalau BPS memperbarui kondef, ubah teks `ket=` pada kartu yang bersangkutan di daftar `N`
di dalam `build.py`.

## Kolom rincian pada tabel popup

Tabel di setiap kartu diperkaya dengan kolom rincian dari kedua berkas ekspor, misalnya status
Meninggal, Tidak Eligible, Tidak Ditemukan, dan Nonrespon pada kartu Keluarga, atau status UMKM
Ditemukan, Baru, Tutup, Ganda, dan Tidak Ditemukan pada kartu UMK. Daftarnya diatur pada `EKSTRA`
di dalam `build.py` — cukup tambah atau kurangi pasangan `("Judul kolom", "nama_field")`.
Bila tabel lebih lebar dari layar, tabel dapat digeser mendatar.

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
