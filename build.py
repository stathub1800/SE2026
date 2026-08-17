# -*- coding: utf-8 -*-
"""
Pembangun halaman Struktur Progres SE2026 - BPS Provinsi Lampung.

Cara pakai:
    1. Taruh dua berkas ekspor FASIH terbaru di folder  data\
    2. Jalankan:  python build.py
    3. Hasilnya:  index.html  (siap di-commit ke GitHub Pages)

Nama kartu, definisi, dan rumus ada di daftar N di bawah. Ubah "label" untuk
mengganti nama kartu; jangan ubah "id" karena itu kunci penghubung antar kartu.
"""
import sys, glob, json, datetime, pathlib
import pandas as pd

AKAR = pathlib.Path(__file__).resolve().parent
DATA = AKAR / "data"

def cari(pola, keterangan):
    hit = sorted(DATA.glob(pola))
    if not hit:
        sys.exit(f"GAGAL: berkas {keterangan} tidak ditemukan di {DATA}\n"
                 f"       Dicari dengan pola: {pola}")
    return max(hit, key=lambda p: p.stat().st_mtime)

F1 = cari("*Pendataan*.xlsx", "Progres Pendataan")
F2 = cari("*Pemutakhiran*Keluarga*.xlsx", "Pemutakhiran Keluarga")
print(f"  Pendataan   : {F1.name}")
print(f"  Pemutakhiran: {F2.name}")

def grab(f, s, skip):
    d = pd.read_excel(f, sheet_name=s, header=None, skiprows=skip)
    d = d[pd.to_numeric(d[0], errors="coerce").notna()].copy()
    d[0] = d[0].astype(int).astype(str)
    return d.set_index(0)

def I(df, k, c):
    v = df.loc[k, c]
    return 0 if pd.isna(v) else int(v)

def PC(df, k, c):
    v = df.loc[k, c]
    if pd.isna(v):
        return None
    return float(v.replace(".", "").replace(",", ".")) if isinstance(v, str) else float(v)

pp  = grab(F1, "PROGRES PENDATAAN", 6)
up  = grab(F1, "USAHA PERUSAHAAN", 7)
sk  = grab(F1, "SKALA USAHA", 6)
uk  = grab(F1, "USAHA KELUARGA", 6)
ku  = grab(F1, "KESELURUHAN USAHA", 6)
kel = grab(F2, "KELUARGA", 5)
ak  = grab(F2, "ANGGOTA KELUARGA", 5)
kk  = grab(F2, "KELUARGA KHUSUS", 6)

kodes_semua = list(pp.index)
rows = {}
for k in kodes_semua:
    r = {"nama": str(pp.loc[k, 1])}
    r["pre_assign"] = I(pp,k,2); r["baru_assign"] = I(pp,k,3)
    r["tot_assign"] = r["pre_assign"] + r["baru_assign"]
    r["verif"] = I(pp,k,4); r["responden"] = I(pp,k,6); r["draft"] = I(pp,k,8)
    r["open"]  = max(r["tot_assign"] - r["verif"] - r["draft"], 0)
    r["nihil"] = r["verif"] - r["responden"]
    r["pre_usaha"] = I(up,k,2); r["bku"] = I(up,k,35); r["force"] = I(up,k,33)
    r["usaha_unit"] = r["bku"] + r["force"]
    r["ub"] = I(sk,k,6); r["um"] = I(sk,k,8); r["umk"] = I(sk,k,10); r["unk"] = I(sk,k,12)
    r["pre_ub"] = I(sk,k,2); r["pre_um"] = I(sk,k,3); r["pre_umk"] = I(sk,k,4)
    r["pre_uk"] = I(uk,k,2); r["uk_ok"] = I(uk,k,3); r["uk_baru"] = I(uk,k,11); r["uk"] = I(uk,k,15)
    r["usaha_all"] = I(ku,k,17); r["pre_usaha_all"] = I(ku,k,4)
    r["pre_kel"] = I(kel,k,2); r["kel_ok"] = I(kel,k,3); r["kel_baru"] = I(kel,k,5); r["kel"] = I(kel,k,14)
    r["ak_bersama"] = I(ak,k,2); r["ak_baru"] = I(ak,k,3); r["ak"] = I(ak,k,9)
    r["kk_bangunan"] = I(kk,k,2); r["kk_didata"] = I(kk,k,3)
    r["p_verif"] = PC(pp,k,5); r["p_responden"] = PC(pp,k,7); r["p_draft"] = PC(pp,k,9)
    r["p_bku"] = PC(up,k,36); r["p_force"] = PC(up,k,34)
    r["p_ub"] = PC(sk,k,7); r["p_um"] = PC(sk,k,9); r["p_umk"] = PC(sk,k,11)
    r["p_uk"] = PC(uk,k,16); r["p_usaha_all"] = PC(ku,k,18)
    r["p_kel"] = PC(kel,k,15); r["p_kelok"] = PC(kel,k,4); r["p_kk"] = PC(kk,k,4)
    rows[k] = r

# ---- uji mutu: identitas penjumlahan dan persentase resmi ----
def uji(nama, fn):
    salah = [k for k in kodes_semua if fn(rows[k])[0] != fn(rows[k])[1]]
    print(("  OK    " if not salah else "  GAGAL ") + nama + ("" if not salah else "  -> " + ", ".join(salah)))
    return not salah

print("\nUji identitas penjumlahan:")
lulus = all([
    uji("prelist keluarga + prelist usaha = prelist assignment", lambda r: (r["pre_kel"]+r["pre_usaha"], r["pre_assign"])),
    uji("responden + hasil nihil = verifikasi",                  lambda r: (r["responden"]+r["nihil"], r["verif"])),
    uji("keluarga + usaha BKU + force submit = responden",       lambda r: (r["kel"]+r["bku"]+r["force"], r["responden"])),
    uji("UB + UM + UMK + lainnya = usaha BKU",                   lambda r: (r["ub"]+r["um"]+r["umk"]+r["unk"], r["bku"])),
    uji("keluarga ditemukan + baru = keluarga",                  lambda r: (r["kel_ok"]+r["kel_baru"], r["kel"])),
    uji("usaha BKU + usaha dalam keluarga = total usaha",        lambda r: (r["bku"]+r["uk"], r["usaha_all"])),
])

print("\nUji persentase resmi FASIH terhadap hitung ulang:")
pasangan = [("p_verif","verif","pre_assign"),("p_responden","responden","pre_assign"),("p_draft","draft","pre_assign"),
            ("p_bku","bku","pre_usaha"),("p_force","force","pre_usaha"),("p_ub","ub","pre_ub"),("p_um","um","pre_um"),
            ("p_umk","umk","pre_umk"),("p_uk","uk","pre_uk"),("p_kel","kel","pre_kel"),("p_kelok","kel_ok","pre_kel"),
            ("p_kk","kk_didata","kk_bangunan"),("p_usaha_all","usaha_all","pre_usaha_all")]
beda = 0
for pf, vf, bf in pasangan:
    for k in kodes_semua:
        r = rows[k]
        if r[pf] is None or not r[bf]:
            continue
        if abs(round(r[vf]/r[bf]*100, 2) - r[pf]) > 0.02:
            beda += 1
            print(f"  selisih {k} {pf}: berkas {r[pf]} vs hitung {round(r[vf]/r[bf]*100,2)}")
print(f"  {'OK    semua cocok' if beda == 0 else f'PERIKSA {beda} selisih'}")

kodes = sorted(k for k in rows if k != "18")

N=[
 dict(id='assign',parent=None,label='Assignment SE2026',field='tot_assign',base=None,unit='assignment',tipe='akar',warna='maroon',
   ket='Seluruh tugas pencacahan di FASIH: prelist awal dari basis data pusat ditambah assignment baru yang dibentuk petugas di lapangan.',
   rumus='prelist awal + baru',parts=[('Prelist awal','pre_assign'),('Baru dari lapangan','baru_assign')],
   sumber='Progres Pendataan kol. (3) + (4)',turunan='Penjumlahan dua kolom. Tidak ada satu kolom tunggal di ekspor yang memuat angka ini.'),

 dict(id='capaian',parent='assign',label='Verifikasi',field='verif',base='tot_assign',unit='assignment',tipe='bagi',warna='hijau',
   ket='Assignment yang sudah disubmit petugas dan lolos verifikasi pengawas. Inilah pekerjaan yang benar-benar tuntas.',
   rumus='hasil verifikasi \u00f7 total assignment',sumber='Progres Pendataan kol. (5)',
   resmi=dict(p='p_verif',b='pre_assign',bl='prelist assignment',kol='Progres Pendataan kol. (6)')),
 dict(id='draft',parent='assign',label='Draft',field='draft',base='tot_assign',unit='assignment',tipe='bagi',warna='kuning',neg=1,
   ket='Assignment yang sudah dibuka dan sedang diisi tetapi belum disubmit. Pekerjaan menggantung yang perlu didorong tuntas.',
   rumus='responden sedang didata \u00f7 total assignment',sumber='Progres Pendataan kol. (9)',
   resmi=dict(p='p_draft',b='pre_assign',bl='prelist assignment',kol='Progres Pendataan kol. (10)')),
 dict(id='open',parent='assign',label='Open',field='open',base='tot_assign',unit='assignment',tipe='bagi',warna='merah',neg=1,
   ket='Assignment yang belum disentuh sama sekali. Sisa beban murni yang masih harus dikerjakan.',
   rumus='total assignment \u2212 verifikasi \u2212 draft',sumber='hitungan turunan',
   turunan='Ekspor FASIH tidak memuat kolom Open. Angka ini sisa dari total assignment setelah dikurangi verifikasi dan draft.'),

 dict(id='responden',parent='capaian',label='Responden Didata',field='responden',base='verif',unit='responden',tipe='bagi',warna='hijau',
   ket='Assignment terverifikasi yang menghasilkan responden nyata: keluarga ditemukan atau baru, dan usaha yang berhasil dicacah.',
   rumus='responden didata \u00f7 verifikasi',sumber='Progres Pendataan kol. (7)',
   resmi=dict(p='p_responden',b='pre_assign',bl='prelist assignment',kol='Progres Pendataan kol. (8)')),
 dict(id='nihil',parent='capaian',label='Hasil Nihil',field='nihil',base='verif',unit='assignment',tipe='bagi',warna='abu',neg=1,
   ket='Sudah diverifikasi tetapi tidak menghasilkan responden: usaha tutup, pindah tak tertelusuri, ganda, tidak ditemukan, keluarga meninggal atau tidak eligible. Pekerjaan tetap dihitung selesai.',
   rumus='verifikasi \u2212 responden didata',sumber='hitungan turunan',
   turunan='Ekspor FASIH tidak memuat kolom ini. Angka ini selisih antara hasil verifikasi dan responden didata.'),

 dict(id='usaha',parent='responden',label='Usaha',field='usaha_unit',base='responden',unit='usaha',tipe='bagi',warna='biru',
   ket='Responden berupa unit usaha yang dicacah lewat assignment usaha, mencakup usaha BKU ditambah usaha yang diklaim sebagai usaha keluarga (force submit).',
   rumus='usaha BKU + force submit',parts=[('Usaha BKU','bku'),('Force submit','force')],
   sumber='Usaha/Perusahaan kol. (35) + (36)',
   turunan='Penjumlahan dua kolom status. Ekspor tidak memuat gabungan ini dalam satu kolom.',
   resmi=dict(p='p_bku',f='bku',fl='Usaha BKU',b='pre_usaha',bl='prelist usaha',kol='Usaha/Perusahaan kol. (37)')),
 dict(id='keluarga',parent='responden',label='Keluarga',field='kel',base='responden',unit='keluarga',tipe='bagi',warna='oranye',
   ket='Responden berupa keluarga hasil pemutakhiran: keluarga prelist yang ditemukan ditambah keluarga baru yang dijaring petugas.',
   rumus='keluarga ditemukan + keluarga baru',parts=[('Ditemukan','kel_ok'),('Baru','kel_baru')],
   sumber='Pemutakhiran Keluarga kol. (15)',
   resmi=dict(p='p_kel',b='pre_kel',bl='prelist awal keluarga',kol='Keluarga kol. (16)')),

 dict(id='ub',parent='usaha',label='UB',sub='Usaha Besar',field='ub',base='usaha_unit',unit='usaha',tipe='bagi',warna='maroon',
   ket='Usaha Besar hasil pencacahan. Jumlahnya bisa melampaui prelist karena ada usaha yang naik kelas setelah dicacah.',
   rumus='UB tercacah \u00f7 usaha',sumber='Skala Usaha kol. (7)',
   resmi=dict(p='p_ub',b='pre_ub',bl='prelist UB',kol='Skala Usaha kol. (8)')),
 dict(id='um',parent='usaha',label='UM',sub='Usaha Menengah',field='um',base='usaha_unit',unit='usaha',tipe='bagi',warna='biru',
   ket='Usaha Menengah hasil pencacahan.',rumus='UM tercacah \u00f7 usaha',sumber='Skala Usaha kol. (9)',
   resmi=dict(p='p_um',b='pre_um',bl='prelist UM',kol='Skala Usaha kol. (10)')),
 dict(id='umk',parent='usaha',label='UMK',sub='Usaha Mikro Kecil',field='umk',base='usaha_unit',unit='usaha',tipe='bagi',warna='biru',
   ket='Usaha Mikro dan Kecil hasil pencacahan. Tulang punggung jumlah usaha di Lampung.',
   rumus='UMK tercacah \u00f7 usaha',sumber='Skala Usaha kol. (11)',
   resmi=dict(p='p_umk',b='pre_umk',bl='prelist UMK',kol='Skala Usaha kol. (12)')),
 dict(id='unk',parent='usaha',label='Usaha Lainnya',sub='Skala belum terklasifikasi',field='unk',base='usaha_unit',unit='usaha',tipe='bagi',warna='abu',neg=1,
   ket='Usaha yang sudah tercacah tetapi skalanya belum bisa ditentukan karena isian omzet, aset, atau tenaga kerja belum lengkap. Sasaran perbaikan kualitas.',
   rumus='tidak dapat diklasifikasikan \u00f7 usaha',sumber='Skala Usaha kol. (13)',
   turunan='Ekspor memuat jumlahnya tetapi tidak memuat kolom persentase untuk rincian ini.'),
 dict(id='force',parent='usaha',label='Usaha Keluarga',sub='Force submit dari assignment usaha',field='force',base='usaha_unit',unit='usaha',tipe='bagi',warna='kuning',
   ket='Assignment usaha yang setelah dikunjungi ternyata berskala usaha keluarga, lalu disubmit paksa dan dialihkan ke jalur keluarga.',
   rumus='force submit \u00f7 usaha',sumber='Usaha/Perusahaan kol. (34)',
   resmi=dict(p='p_force',b='pre_usaha',bl='prelist usaha',kol='Usaha/Perusahaan kol. (35)')),
 dict(id='uk',parent='usaha',label='Usaha dalam Keluarga',field='uk',base=None,unit='usaha',tipe='tautan',warna='biru',
   ket='Usaha yang dicatat sebagai roster di dalam assignment keluarga, bukan assignment usaha tersendiri. Angka ini tidak menambah jumlah assignment, tetapi menambah jumlah usaha.',
   rumus='usaha dalam keluarga ditemukan + baru',parts=[('Ditemukan','uk_ok'),('Baru','uk_baru')],
   tautan_ket='Tercatat di dalam assignment keluarga',sumber='Usaha Keluarga kol. (16)',
   resmi=dict(p='p_uk',b='pre_uk',bl='prelist usaha keluarga',kol='Usaha Keluarga kol. (17)')),

 dict(id='kelok',parent='keluarga',label='Keluarga Ditemukan',field='kel_ok',base='kel',unit='keluarga',tipe='bagi',warna='oranye',
   ket='Keluarga prelist yang berhasil ditemukan dan dimutakhirkan di lapangan.',
   rumus='ditemukan \u00f7 keluarga',sumber='Pemutakhiran Keluarga kol. (4)',
   resmi=dict(p='p_kelok',b='pre_kel',bl='prelist awal keluarga',kol='Keluarga kol. (5)')),
 dict(id='kelbaru',parent='keluarga',label='Keluarga Baru',field='kel_baru',base='kel',unit='keluarga',tipe='bagi',warna='kuning',
   ket='Keluarga yang belum ada di prelist dan ditambahkan petugas saat pemutakhiran.',
   rumus='keluarga baru \u00f7 keluarga',sumber='Pemutakhiran Keluarga kol. (6)',
   turunan='Ekspor memuat jumlahnya tetapi tidak memuat kolom persentase untuk rincian ini.'),
 dict(id='ak',parent='keluarga',label='Anggota Keluarga',field='ak',base=None,unit='jiwa',tipe='tautan',warna='oranye',
   ket='Jumlah jiwa di dalam keluarga yang terdata. Satuannya orang, bukan keluarga, sehingga wajar jauh lebih besar daripada jumlah keluarga.',
   rumus='tinggal bersama + anggota baru',parts=[('Tinggal bersama','ak_bersama'),('Anggota baru','ak_baru')],
   tautan_ket='Isi dari setiap keluarga',rasio=('kel','jiwa per keluarga'),sumber='Anggota Keluarga kol. (10)',
   turunan='Ekspor tidak memuat kolom persentase untuk anggota keluarga.'),
 dict(id='kkhusus',parent='keluarga',label='Keluarga Khusus',field='kk_didata',base=None,unit='bangunan',tipe='tautan',warna='merah',
   ket='Bangunan tempat tinggal khusus seperti asrama, panti, dan lembaga pemasyarakatan. Didata terpisah lewat dokumen K1 dengan satuan bangunan.',
   rumus='bangunan didata \u00f7 bangunan hasil pendataan PPL',tautan_ket='Jalur K1, satuan bangunan',
   sumber='Keluarga Khusus kol. (4)',
   resmi=dict(p='p_kk',b='kk_bangunan',bl='bangunan hasil listing PPL',kol='Keluarga Khusus kol. (5)')),
]

def jam(p):
    return datetime.datetime.fromtimestamp(p.stat().st_mtime).strftime("%d %b %Y %H.%M")

stempel = ("Sumber FASIH Pendataan SE2026<br>Pendataan " + jam(F1)
           + " &middot; Keluarga " + jam(F2)
           + "<br>Halaman disusun " + datetime.datetime.now().strftime("%d %b %Y %H.%M"))

muatan = json.dumps({"rows": rows, "kodes": kodes, "nodes": N, "stempel": stempel},
                    ensure_ascii=False, separators=(",", ":"))

tpl = (AKAR / "template.html").read_text(encoding="utf-8")
if "__PAYLOAD__" not in tpl:
    sys.exit("GAGAL: template.html tidak memuat penanda __PAYLOAD__")
(AKAR / "index.html").write_text(tpl.replace("__PAYLOAD__", muatan), encoding="utf-8")

print(f"\nindex.html selesai ditulis ({len(rows)} wilayah, {len(N)} kartu).")
if not lulus or beda:
    print("PERINGATAN: ada uji yang tidak lulus. Periksa berkas ekspor sebelum dipublikasikan.")
