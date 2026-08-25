// ============================================================
// Portal Administrasi SE2026 — bagian: docTypes
// ============================================================

import { Briefcase, Car, ClipboardList, FileText, Map as MapIcon, Receipt, Users } from "lucide-react";


// ─── DATA ────────────────────────────────────────────────────────────────────

export const DOC_TYPES = [
  { id: "daftar-hadir",  icon: <ClipboardList />, label: "Daftar Hadir", desc: "Daftar Hadir Pelatihan SE2026", color: "orange", disabled: false, lockedMessage: "Fitur Daftar Hadir dikunci" },
  { id: "tanda-terima",  icon: <Briefcase />,     label: "Tanda Terima", desc: "Fitur dikunci", color: "amber", disabled: true, lockedMessage: "Fitur Tanda Terima dikunci" },
  { id: "surat-pernyataan-kendaraan", icon: <Car />, label: "Super Kendis", desc: "Fitur dikunci", color: "orange", disabled: true, lockedMessage: "Fitur Super Kendis dikunci" },
  { id: "pengeluaran-riil", icon: <Receipt />,    label: "DPR", desc: "Daftar Pengeluaran Riil", color: "amber", disabled: false, lockedMessage: "Fitur DPR dikunci" },
  { id: "spj",           icon: <FileText />,      label: "SPJ", desc: "SPJ Pelatihan SE2026", color: "orange", disabled: false, lockedMessage: "Fitur SPJ dikunci" },
  { id: "spd",           icon: <MapIcon />,           label: "SPD", desc: "Fitur dikunci", color: "amber", disabled: true, lockedMessage: "Fitur SPD dikunci" },
  { id: "surat-tugas",   icon: <Users />,         label: "Surtug", desc: "Fitur dikunci", color: "orange", disabled: true, lockedMessage: "Fitur Surat Tugas dikunci" },
  { id: "bapp",          icon: <FileText />,      label: "BAPP", desc: "Fitur dikunci", color: "amber", disabled: true, lockedMessage: "Fitur BAPP dikunci" },
  { id: "bast",          icon: <FileText />,      label: "BAST", desc: "Fitur dikunci", color: "amber", disabled: true, lockedMessage: "Fitur BAST dikunci" },
  { id: "surat-pernyataan-penyelesaian-lapangan", icon: <FileText />, label: "Surat Pernyataan Penyelesaian Lapangan", desc: "Fitur dikunci", color: "amber", disabled: true, lockedMessage: "Fitur Surat Pernyataan Penyelesaian Lapangan dikunci" },
  { id: "lampiran",      icon: <FileText />,      label: "Lampiran", desc: "Fitur dikunci", color: "amber", disabled: true, lockedMessage: "Fitur Lampiran dikunci" },
  { id: "gabungan-pembayaran", icon: <Receipt />, label: "Gabungan Administrasi Pembayaran", desc: "Generate satu berkas pembayaran lengkap per PML atau PPL", color: "amber" },
  {
  id: "surat-kepala",
  icon: <FileText />,
  label: "Surat Kepala",
  desc: "SPEPL Kepala BPS (PML & PPL)",
  color: "amber",
},
];

