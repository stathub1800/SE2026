/* ============================================================
   config.js – Central Configuration
   Edit this file to change settings without touching logic.
   ============================================================ */

const CONFIG = {

  /* App */
  APP_NAME:    'Dashboard Monitoring SE2026',
  PROVINCE:    'Provinsi Lampung',
  VERSION:     '1.0.0',

  /* Auto-refresh interval (milliseconds). Set 0 to disable. */
  REFRESH_INTERVAL: 5 * 60 * 1000, // 5 menit

  /* Map center & zoom for Lampung Province */
  MAP: {
    CENTER: [-4.8, 105.2],
    ZOOM:   8,
    ZOOM_MOBILE: 7,
    MIN_ZOOM: 6,
    MAX_ZOOM: 14,
    TILE_URL: 'https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',
    TILE_ATTR: '© OpenStreetMap contributors',
  },

  /* Colors */
  COLORS: {
    PRIMARY:  '#E8500A',
    SUCCESS:  '#22C55E',
    WARNING:  '#EAB308',
    DANGER:   '#EF4444',
    INFO:     '#3B82F6',
    GRAY:     '#A3A3A3',
  },

  /* Kabupaten/Kota in Lampung (15 daerah) */
  KABKOT: [
    { id: 'lp01', name: 'Kab. Lampung Barat',   lat: -4.970, lng: 104.072 },
    { id: 'lp02', name: 'Kab. Tanggamus',        lat: -5.361, lng: 104.627 },
    { id: 'lp03', name: 'Kab. Lampung Selatan',  lat: -5.468, lng: 105.397 },
    { id: 'lp04', name: 'Kab. Lampung Timur',    lat: -4.950, lng: 105.680 },
    { id: 'lp05', name: 'Kab. Lampung Tengah',   lat: -4.815, lng: 105.069 },
    { id: 'lp06', name: 'Kab. Lampung Utara',    lat: -4.518, lng: 104.921 },
    { id: 'lp07', name: 'Kab. Way Kanan',        lat: -4.302, lng: 104.538 },
    { id: 'lp08', name: 'Kab. Tulang Bawang',    lat: -4.349, lng: 105.637 },
    { id: 'lp09', name: 'Kab. Pesawaran',        lat: -5.264, lng: 105.127 },
    { id: 'lp10', name: 'Kab. Pringsewu',        lat: -5.356, lng: 104.973 },
    { id: 'lp11', name: 'Kab. Mesuji',           lat: -3.870, lng: 105.625 },
    { id: 'lp12', name: 'Kab. Tulang Bawang Barat', lat: -4.200, lng: 105.090 },
    { id: 'lp13', name: 'Kab. Pesisir Barat',    lat: -4.908, lng: 103.780 },
    { id: 'lp14', name: 'Kota Bandar Lampung',   lat: -5.396, lng: 105.261 },
    { id: 'lp15', name: 'Kota Metro',            lat: -5.109, lng: 105.307 },
  ],

  /* Data source labels */
  SOURCES: {
    FASIH:     'FASIH Pendataan SE 2026',
    BACKOFFICE:'Backoffice',
    MEDIA:     'Dashboard Kurasi Media',
  },
};
