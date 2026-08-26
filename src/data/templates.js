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
export const BERKAS_PEMBAYARAN_PML_TEMPLATE_URL = "/templates/BERKAS PEMBAYARAN PML.docx";
export const BERKAS_PEMBAYARAN_PPL_TEMPLATE_URL = "/templates/BERKAS PEMBAYARAN PPL.docx";
export const SURAT_KEPALA_TEMPLATE_URL = "/templates/SURAT PERNYATAAN KEPALA BPS.docx";

// ─── EXPORT EXCEL (kop surat, per halaman → per sheet) ─────────────────────
// Word tidak punya "jumlah baris per halaman" yang eksplisit (auto-flow sesuai
// margin/font), jadi untuk versi Excel dipakai angka perkiraan yang bisa
// disesuaikan di sini tanpa perlu bongkar kode generator.
export const KOP_SURAT_LOGO_URL = "/Logo BPS - Horizontal.png";
export const DAFTAR_HADIR_ROWS_PER_PAGE = 20;
export const SPJ_ROWS_PER_PAGE           = 15;
export const DPR_ROWS_PER_PAGE           = 15;

