// ============================================================
// Portal Administrasi SE2026 — bagian: templatesConst
// ============================================================

import { upperText } from "../lib/helpers";

// ─── FILTER GROUPS ───────────────────────────────────────────────────────────

export const DAFTAR_HADIR_GROUPS = {
  "pml-ppl":      { label: "PML & PPL",      roles: ["PML", "PPL"] },
  "panitia-inda": { label: "Panitia & Inda", roles: ["PANITIA", "INDA", "KEPALA BPS JAKARTA TIMUR"] },
};

export function jabatanMasukGroup(jabatan, groupKey) {
  const group = DAFTAR_HADIR_GROUPS[groupKey];
  return group ? group.roles.includes(upperText(jabatan)) : false;
}

// ─── TEMPLATE URLS ───────────────────────────────────────────────────────────

export const DAFTAR_HADIR_TEMPLATE_URL               = "/templates/1. Daftar Hadir Pelatihan SE2026.docx";
export const TANDA_TERIMA_TEMPLATE_URL               = "/templates/2. Tanda Terima Perlengkapan SE2026.docx";
export const TANDA_TERIMA_LAPANGAN_TEMPLATE_URL      = "/templates/2. Tanda Terima Perlengkapan SE2026 - Copy.docx";
export const SURAT_PERNYATAAN_KENDARAAN_TEMPLATE_URL = "/templates/3. Super Kendis Pelatihan SE2026.docx";
export const PENGELUARAN_RIIL_TEMPLATE_URL           = "/templates/4. DPR_Pelatihan SE 2026.docx";
export const SPJ_TEMPLATE_URL                        = "/templates/5. SPJ Pelatihan_SE26.docx";
export const SPD_TEMPLATE_URL                        = "/templates/6. SPD.docx";
export const SPD_LAMPIRAN_TEMPLATE_URL               = "/templates/6. Lampiran SPD.docx";
export const SURAT_TUGAS_TEMPLATE_URL                = "/templates/6. Surat Tugas.docx";
export const BAPP_PML_TEMPLATE_URL                   = "/templates/BAPP PML.docx";
export const BAPP_PPL_TEMPLATE_URL                   = "/templates/BAPP PPL.docx";
export const SURAT_PERNYATAAN_PENYELESAIAN_LAPANGAN_TEMPLATE_URL = "/templates/Dasar Pembayaran.docx";
export const LAMPIRAN_PML_TEMPLATE_URL              = "/templates/LAMPIRAN PML.docx";
export const LAMPIRAN_PPL_TEMPLATE_URL              = "/templates/LAMPIRAN PPL.docx";
export const BAST_PML_TEMPLATE_URL = "/templates/BAST PML.docx";
export const BAST_PPL_TEMPLATE_URL = "/templates/BAST PPL.docx";
export const BERKAS_PEMBAYARAN_PML_TEMPLATE_URL = "/templates/BERKAS PEMBAYARAN TERMIN II PML (1).docx";
export const BERKAS_PEMBAYARAN_PPL_TEMPLATE_URL = "/templates/BERKAS PEMBAYARAN PPL TERMIN II (1).docx";

export const SURAT_KEPALA_TEMPLATE_URL = "/templates/SURAT PERNYATAAN KEPALA BPS.docx";


// ─── TEMPLATE PER DURASI — Gabungan Administrasi Pembayaran (khusus PPL) ───
// Sebagian PPL dibayar dengan durasi kontrak yang beda (1 bulan / 1,5 bulan /
// 2 bulan), jadi butuh file .docx yang berbeda juga. Orangnya dicocokkan
// lewat EMAIL (daftar di GABUNGAN_PEMBAYARAN_PPL_EMAIL_DURASI di bawah).
//
// ⚠️ FILE-NYA BELUM DIUPLOAD. Kalau sudah siap file-nya:
// 1. Upload/taruh file .docx ke folder `public/templates/`.
// 2. Ganti isi 3 baris URL di bawah ini (BERKAS_PEMBAYARAN_PPL_2_BULAN_...
//    dst) supaya PERSIS SAMA dengan nama file yang diupload — termasuk
//    spasi, huruf besar/kecil, dan tanda baca. Formatnya:
//    "/templates/NAMA FILE PERSIS.docx".
// Sebelum diganti, kalau salah satu email di bawah kebetulan dipilih untuk
// digenerate, akan muncul error "Gagal memuat template..." — itu tandanya
// filenya belum ada/namanya belum cocok, bukan bug.
export const BERKAS_PEMBAYARAN_PPL_2_BULAN_TEMPLATE_URL   = "/templates/BERKAS PEMBAYARAN PPL TERMIN II - 2 BULAN (1).docx";
export const BERKAS_PEMBAYARAN_PPL_1_5_BULAN_TEMPLATE_URL = "/templates/BERKAS PEMBAYARAN PPL TERMIN II - 1,5 BULAN (1).docx";
export const BERKAS_PEMBAYARAN_PPL_1_BULAN_TEMPLATE_URL   = "/templates/BERKAS PEMBAYARAN PPL TERMIN II - 1 BULAN (1).docx";

// Peta kunci durasi -> URL template. Kalau nanti ada durasi baru lagi
// (mis. "3 BULAN"), tinggal tambah baris URL di atas + satu baris di sini +
// satu key baru di GABUNGAN_PEMBAYARAN_PPL_EMAIL_DURASI di bawah.
export const GABUNGAN_PEMBAYARAN_PPL_DURASI_TEMPLATE_URL = {
  "2_BULAN": BERKAS_PEMBAYARAN_PPL_2_BULAN_TEMPLATE_URL,
  "1_5_BULAN": BERKAS_PEMBAYARAN_PPL_1_5_BULAN_TEMPLATE_URL,
  "1_BULAN": BERKAS_PEMBAYARAN_PPL_1_BULAN_TEMPLATE_URL,
};

// Daftar email per durasi. Huruf besar/kecil bebas (dicocokkan otomatis).
// Email yang TIDAK ada di daftar manapun tetap pakai template PPL default
// (BERKAS_PEMBAYARAN_PPL_TEMPLATE_URL di atas).
export const GABUNGAN_PEMBAYARAN_PPL_EMAIL_DURASI = {
  "2_BULAN": [
    "aqellaprametaps@gmail.com",
    "sistantopurnomo@gmail.com",
    "armildasatriyaardiyanti@gmail.com",
    "adr.andara16@gmail.com",
    "nurainine80@gmail.com",
    "aldinurcajyo4@gmail.com",
    "novitatiolinas27@gmail.com",
    "Mulyadi.silent@gmail.com",
    "azzahranurzelika27@gmail.com",
    "ariakbarmudzakir@gmail.com",
    "Mboot13@gmail.com",
    "mooguumooguu02@gmail.com",
    "mamanikivia@gmail.com",
    "atikalisati@gmail.com",
    "riaaqi@gmail.com",
    "Nanygeulies@gmail.com",
  ],
  "1_5_BULAN": [
    "tyaarnia15@gmail.com",
    "yasnirohanalbs@gmail.com",
  ],
  "1_BULAN": [
    "salwadnsyh@gmail.com",
  ],
};

// ─── EXPORT EXCEL (kop surat, per halaman → per sheet) ─────────────────────
// Word tidak punya "jumlah baris per halaman" yang eksplisit (auto-flow sesuai
// margin/font), jadi untuk versi Excel dipakai angka perkiraan yang bisa
// disesuaikan di sini tanpa perlu bongkar kode generator.
export const KOP_SURAT_LOGO_URL = "/Logo BPS - Horizontal.png";
export const DAFTAR_HADIR_ROWS_PER_PAGE = 20;
export const SPJ_ROWS_PER_PAGE           = 15;
export const DPR_ROWS_PER_PAGE           = 15;

