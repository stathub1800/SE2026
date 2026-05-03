/* ============================================================
   data.js – Data Layer
   Ganti fungsi fetchXxx() dengan pemanggilan API nyata.
   Untuk demo, data dummy digunakan.
   ============================================================ */

const DATA = {

  /* ---- OVERVIEW / AGREGAT ---- */
  async fetchOverview() {
    // TODO: ganti dengan fetch('/api/overview') atau Supabase query
    return {
      progress: {
        sls_tercacah: 0,
        sls_target: 0,
        pct: 0.0,
        target_harian: 0,
        hari_tersisa: 0,
      },
      anggaran: {
        pagu_revisi:    1_266_887_499_000,
        pagu_realisasi: 23_668_739_202,
        pct_serapan:    4.82,
        realisasi_tertinggi: 'Pengadaan Instrumen SE2026',
      },
      kinerja: {
        petugas_aktif: 0,
        pcl: 0,
        pml: 0,
        sangat_produktif: 0,
        produktif: 0,
        kurang_produktif: 0,
      },
      kualitas: {
        total_usaha: 0,
        total_omzet: 0,
        nilai_tambah: 0,
      },
      komunikasi: {
        berita: 2601,
        posting: 6744,
        media_tv: 0,
        media_cetak: 177,
        media_online: 2407,
        running_text: 17,
        radio: 0,
      },
      laporan: {
        target: 1104,
        diterima: 316,
        belum: 788,
        pct: 28.6,
      },
      kkd: {
        surat_dikirim: 849,
        surat_diterima: 817,
        tingkat_respons: 96.3,
        kkd_pusat_pct: 10.9,
      },
      kontrak: {
        surat_dukungan_pct: 80.33,
        sosialisasi_pct: 36.49,
        ngibar_pct: 5.44,
      },
    };
  },

  /* ---- PROGRES PENDATAAN ---- */
  async fetchProgres() {
    return {
      sls: { selesai: 0, total: 0, perubahan: 0, rasio_perubahan: 0.0 },
      petugas: { nilai_rata2: 0.0 },
      cawi: { min: 0, realisasi: 0 },
      usaha_per_sls: 0,
      kendala: { menolak: 0, sulit_ditemui: 0 },
      ub_nonrespons: { usaha_besar: 0, umkm: 0 },
      cakupan: { tercacah: 0, pct: 0.0 },
      geotagging: { sudah: 0, belum: 0, target: 0, pct: 0.0 },
      pendataan: { cawi: 0, capi: 0, dopu: 0 },
      moda: {
        surat_cawi_terkirim: 0,
        alamat_tidak_ditemukan: 0,
        gagal_bertemu: 0,
        switch: { cawi_capi: 0, cawi_dopu: 0, capi_cawi: 0, capi_dopu: 0, dopu_cawi: 0, dopu_capi: 0 },
      },
      virtual_office: 0,

      /* Per kabkot – isi array untuk peta */
      per_kabkot: CONFIG.KABKOT.map(k => ({
        ...k,
        tercacah: 0,
        target: 0,
        pct: 0,
      })),
    };
  },

  /* ---- KINERJA PETUGAS ---- */
  async fetchKinerja() {
    return {
      total_rekrutmen: 556,
      organik: 55,
      mitra_umum: 404,
      mitra_afirmasi: 97,
      mahasiswa: 0,
      pelatihan: { terlatih_penuh: 0, sedang: 0, belum: 125, pct_belum: 35.1 },
      total_petugas_training: 356,
      target_terlatih: 'Agustus 2026',
      risiko: { mengundurkan_diri: 0, kecelakaan: 0, meninggal: 0 },
      produktivitas: {
        sangat_produktif: 0,
        produktif: 0,
        kurang_produktif: 0,
      },
      per_kabkot: CONFIG.KABKOT.map(k => ({
        ...k,
        jumlah_petugas: Math.floor(Math.random() * 30) + 5, // demo
      })),
    };
  },

  /* ---- KUALITAS DATA ---- */
  async fetchKualitas() {
    return {
      total_usaha: 0,
      total_omzet: 0,
      nilai_tambah: 0,
      per_kabkot: CONFIG.KABKOT.map(k => ({
        ...k, usaha: 0, omzet: 0,
      })),
    };
  },

  /* ---- SERAPAN ANGGARAN ---- */
  async fetchAnggaran() {
    return {
      pagu_revisi:    1_266_887_499_000,
      pagu_realisasi: 23_668_739_202,
      pct_serapan:    4.82,
      realisasi_tertinggi: 'Pengadaan Instrumen SE2026',
      rincian: [
        { nama: 'Pengadaan Instrumen', pagu: 450_000_000, realisasi: 23_668_739_202, pct: 52.6 },
        { nama: 'Honor Petugas',       pagu: 350_000_000, realisasi: 0,              pct: 0 },
        { nama: 'Operasional',         pagu: 200_000_000, realisasi: 0,              pct: 0 },
        { nama: 'Pelatihan',           pagu: 150_000_000, realisasi: 0,              pct: 0 },
        { nama: 'Komunikasi Publik',   pagu: 100_000_000, realisasi: 0,              pct: 0 },
      ],
    };
  },

  /* ---- KOMUNIKASI PUBLIK ---- */
  async fetchKomunikasi() {
    return {
      berita: 2601,
      posting: 6744,
      media: [
        { nama: 'Media Online', jumlah: 2407 },
        { nama: 'Media Cetak',  jumlah: 177  },
        { nama: 'Running Text', jumlah: 17   },
        { nama: 'Televisi',     jumlah: 0    },
        { nama: 'Radio',        jumlah: 0    },
      ],
    };
  },

  /* ---- LAPORAN KEGIATAN ---- */
  async fetchLaporan() {
    return {
      target: 1104,
      diterima: 316,
      belum: 788,
      pct: 28.6,
      per_kabkot: CONFIG.KABKOT.map((k, i) => ({
        ...k,
        target: Math.floor(Math.random() * 50) + 20,
        diterima: Math.floor(Math.random() * 20),
        pct: Math.floor(Math.random() * 60),
      })),
    };
  },

  /* ---- KKD ---- */
  async fetchKKD() {
    return {
      surat_dikirim: 849,
      surat_diterima: 817,
      tingkat_respons: 96.3,
      kkd_pusat_pct: 10.9,
      per_kabkot: CONFIG.KABKOT.map(k => ({
        ...k,
        dikirim: Math.floor(Math.random() * 60) + 10,
        diterima: Math.floor(Math.random() * 55) + 10,
      })),
    };
  },

  /* ---- KONTRAK KINERJA ---- */
  async fetchKontrak() {
    return {
      surat_dukungan_pct: 80.33,
      sosialisasi_pct: 36.49,
      ngibar_pct: 5.44,
      per_kabkot: CONFIG.KABKOT.map(k => ({
        ...k,
        surat_dukungan_pct: Math.floor(Math.random() * 100),
        sosialisasi_pct: Math.floor(Math.random() * 80),
        ngibar_pct: Math.floor(Math.random() * 30),
      })),
    };
  },
};
