// ============================================================
// Portal Administrasi SE2026 — bagian: helpers
// ============================================================

import * as XLSX from "xlsx";

// ─── HELPERS ─────────────────────────────────────────────────────────────────

export function cleanText(value) { return String(value ?? "").trim(); }
export function upperText(value) { return cleanText(value).toUpperCase(); }

// ─── SELECTION XLSX (fitur "Download Beberapa") ──────────────────────────────
// Memungkinkan pengguna mengunggah file Excel berisi kolom Nama dan/atau Email
// untuk memilih sekumpulan orang yang mau di-generate sekaligus (batch terpilih),
// tanpa harus memilih satu-per-satu lewat dropdown "Download Terpilih".
// Pencocokan diprioritaskan lewat Email (lebih unik), fallback ke Nama (uppercase).

export function normalizeSelectionRow(row) {
  const normalized = {};
  Object.entries(row).forEach(([k, v]) => {
    const key = String(k ?? "").trim().toLowerCase().replace(/\s+/g, " ");
    normalized[key] = String(v ?? "").trim();
  });
  return {
    nama: normalized["nama"] || normalized["nama lengkap"] || normalized["nama_lengkap"] || normalized["nama-lengkap"] || "",
    email: normalized["email"] || normalized["mail"] || normalized["e-mail"] || "",
  };
}

export function parseSelectionXlsx(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) return [];
  const raw = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  return raw.map(normalizeSelectionRow).filter((r) => r.nama || r.email);
}

export function buildSelectionKeySet(selectionRows = []) {
  const set = new Set();
  for (const r of selectionRows) {
    const email = cleanText(r.email);
    const nama = cleanText(r.nama);
    if (email) set.add(`EMAIL::${upperText(email)}`);
    if (nama) set.add(`NAME::${upperText(nama)}`);
  }
  return set;
}

export function rowMatchesSelection(selectionKeySet, nama, email) {
  if (!selectionKeySet || selectionKeySet.size === 0) return false;
  const emailClean = cleanText(email);
  if (emailClean && selectionKeySet.has(`EMAIL::${upperText(emailClean)}`)) return true;
  const namaClean = cleanText(nama);
  if (namaClean && selectionKeySet.has(`NAME::${upperText(namaClean)}`)) return true;
  return false;
}

// Versi khusus email-only, dipakai HANYA di fitur Gabungan Administrasi Pembayaran.
export function buildEmailOnlySelectionKeySet(selectionRows = []) {
  const set = new Set();
  for (const r of selectionRows) {
    const email = cleanText(r.email);
    if (email) set.add(upperText(email));
  }
  return set;
}

export function rowMatchesSelectionByEmail(selectionKeySet, email) {
  if (!selectionKeySet || selectionKeySet.size === 0) return false;
  const emailClean = cleanText(email);
  if (!emailClean) return false;
  return selectionKeySet.has(upperText(emailClean));
}

export function uniqueSorted(values) {
  return [...new Set(values.map(cleanText).filter(Boolean))].sort((a, b) =>
    a.localeCompare(b, "id-ID", { numeric: true, sensitivity: "base" })
  );
}

export function formatTanggalIndonesia(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  if (Number.isNaN(d.getTime())) return dateStr;
  return d.toLocaleDateString("id-ID", { day: "numeric", month: "long", year: "numeric" });
}

export function formatTanggalLengkapIndonesia(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  if (Number.isNaN(d.getTime())) return dateStr;
  return d.toLocaleDateString("id-ID", { weekday: "long", day: "numeric", month: "long", year: "numeric" });
}

export function formatTanggalBulanIndonesia(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  if (Number.isNaN(d.getTime())) return dateStr;
  return d.toLocaleDateString("id-ID", { day: "numeric", month: "long" });
}

export function formatHariIndonesia(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  if (Number.isNaN(d.getTime())) return dateStr;
  return d.toLocaleDateString("id-ID", { weekday: "long" });
}

export function calcDurationDays(startDate, endDate) {
  if (!startDate || !endDate) return "";
  const start = new Date(startDate);
  const end = new Date(endDate);
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return "";
  const diff = Math.floor((end.getTime() - start.getTime()) / (1000 * 60 * 60 * 24)) + 1;
  return diff > 0 ? diff : 0;
}

export function normalizeJamIndonesia(value, fallback = "") {
  if (!value) return fallback;
  return String(value).trim().replace(":", ".");
}

export function formatRupiah(value) {
  if (value == null || value === "") return "";
  const number = Number(String(value).replace(/[^0-9-]/g, ""));
  if (Number.isNaN(number)) return String(value);
  return number.toLocaleString("id-ID");
}

export function capitalizeWords(text) {
  return String(text || "").split(/\s+/).filter(Boolean)
    .map((w) => w[0].toUpperCase() + w.slice(1)).join(" ");
}

export function spellTerbilang(value) {
  const units = ["", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan"];
  const teens = ["sepuluh", "sebelas", "dua belas", "tiga belas", "empat belas", "lima belas",
    "enam belas", "tujuh belas", "delapan belas", "sembilan belas"];
  const toWords = (n) => {
    if (n < 10) return units[n];
    if (n < 20) return teens[n - 10];
    if (n < 100) return `${toWords(Math.floor(n / 10))} puluh${n % 10 ? ` ${toWords(n % 10)}` : ""}`.trim();
    if (n < 200) return `seratus${n % 100 ? ` ${toWords(n % 100)}` : ""}`.trim();
    if (n < 1000) return `${toWords(Math.floor(n / 100))} ratus${n % 100 ? ` ${toWords(n % 100)}` : ""}`.trim();
    if (n < 2000) return `seribu${n % 1000 ? ` ${toWords(n % 1000)}` : ""}`.trim();
    if (n < 1000000) return `${toWords(Math.floor(n / 1000))} ribu${n % 1000 ? ` ${toWords(n % 1000)}` : ""}`.trim();
    if (n < 1000000000) return `${toWords(Math.floor(n / 1000000))} juta${n % 1000000 ? ` ${toWords(n % 1000000)}` : ""}`.trim();
    if (n < 1000000000000) return `${toWords(Math.floor(n / 1000000000))} miliar${n % 1000000000 ? ` ${toWords(n % 1000000000)}` : ""}`.trim();
    return String(n);
  };
  const number = Number(String(value).replace(/[^0-9]/g, ""));
  if (Number.isNaN(number) || number === 0) return capitalizeWords("nol");
  return capitalizeWords(toWords(number));
}

export const PARTICIPANT_ROLE_ORDER = { INDA: 0, PANITIA: 1, PML: 2, PPL: 3 };
export function pesertaRoleOrder(jabatan) { return PARTICIPANT_ROLE_ORDER[upperText(jabatan)] ?? 99; }
export function sortPesertaByJabatanOrder(peserta = []) {
  return [...peserta].sort((a, b) => {
    const diff = pesertaRoleOrder(a.jabatan) - pesertaRoleOrder(b.jabatan);
    if (diff !== 0) return diff;
    return cleanText(a.nama).localeCompare(cleanText(b.nama), "id-ID", { sensitivity: "base" });
  });
}

export const DAFTAR_HADIR_PESERTA_ROLE_ORDER = { "KEPALA BPS JAKARTA TIMUR": -1, INDA: 0, PANITIA: 1, PML: 2, PPL: 3 };
export function pesertaRoleOrderDaftarHadir(jabatan) { return DAFTAR_HADIR_PESERTA_ROLE_ORDER[upperText(jabatan)] ?? 99; }
export function sortDaftarHadirPeserta(peserta = []) {
  return [...peserta].sort((a, b) => {
    const diff = pesertaRoleOrderDaftarHadir(a.jabatan) - pesertaRoleOrderDaftarHadir(b.jabatan);
    if (diff !== 0) return diff;
    return cleanText(a.nama).localeCompare(cleanText(b.nama), "id-ID", { sensitivity: "base" });
  });
}

