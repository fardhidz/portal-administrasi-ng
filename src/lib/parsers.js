// ============================================================
// Portal Administrasi SE2026 — bagian: parsers
// ============================================================

import * as XLSX from "xlsx";
import ImageModule from "docxtemplater-image-module-free";

import { cleanText, upperText } from "./helpers.js";

// ─── XLSX PARSER ─────────────────────────────────────────────────────────────

export function normalizeRowHeaders(row) {
  const normalized = {};
  Object.entries(row).forEach(([k, v]) => {
    const key = String(k ?? "").trim().toLowerCase().replace(/\s+/g, " ");
    normalized[key] = String(v ?? "").trim();
  });
  return {
    no:           normalized["no"] ?? "",
    nama:         normalized["nama"] ?? normalized["nama lengkap"] ?? normalized["nama_lengkap"] ?? normalized["nama-lengkap"] ?? "",
    nik:          normalized["nik"] ?? normalized["nip"] ?? "",
    asal:         normalized["asal"] ?? "",
    wilTugas:     (normalized["wil. tugas"] ?? normalized["wil tugas"] ?? normalized["wil.tugas"] ?? normalized["wilayah tugas"] ?? "").toUpperCase(),
    jabatan:      (normalized["jabatan"] ?? normalized["posisi"] ?? "").toUpperCase(),
    pangkatGol:   normalized["pangkat/gol"] ?? normalized["pangkat gol"] ?? normalized["pangkatgol"] ?? "",
    kelas:        String(normalized["kelas"] ?? "").trim(),
    hotel:        (normalized["tc"] ?? normalized["hotel"] ?? "").toUpperCase(),
    gelombang:    String(normalized["gelombang"] ?? "").trim(),
    tc:           (normalized["tc"] ?? "").toUpperCase(),
    sobatId:      normalized["sobat id"] ?? normalized["sobatid"] ?? "",
    email:        normalized["email"] ?? "",
    jenisKelamin: normalized["jenis kelamin"] ?? normalized["jeniskelamin"] ?? "",
  };
}

export function parseXlsxData(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  
  // Cari sheet yang berisi data administrasi secara otomatis
  const adminSheetInfo = findAdministrasiSheet(workbook);
  if (!adminSheetInfo) {
    console.warn("Sheet data administrasi tidak ditemukan. Mencoba sheet pertama...");
    // Fallback ke sheet pertama jika tidak ditemukan
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    if (!firstSheet) throw new Error("File XLSX kosong");
    const raw = XLSX.utils.sheet_to_json(firstSheet, { defval: "", raw: false });
    return raw.map(normalizeRowHeaders).filter(r => r.nama !== "" || r.nik !== "" || r.sobatId !== "");
  }
  
  const raw = adminSheetInfo.data;
  return raw.map(normalizeRowHeaders).filter(r => r.nama !== "" || r.nik !== "" || r.sobatId !== "");
}

export function normalizeBappRow(row) {
  const normalized = {};
  Object.entries(row).forEach(([k, v]) => {
    const key = String(k ?? "").trim().toLowerCase().replace(/\s+/g, " ");
    normalized[key] = String(v ?? "").trim();
  });

  const get = (...keys) => {
    for (const key of keys) {
      const value = normalized[key];
      const text = String(value ?? "").trim();
      // Sel formula Excel/Google Sheets seperti #N/A jangan dianggap sebagai data.
      if (text && !/^#(?:N\/A|VALUE!|REF!|DIV\/0!|NAME\?|NUM!|NULL!)$/i.test(text)) return text;
    }
    return "";
  };

  const nama = get("nama", "nama lengkap", "nama_lengkap", "nama petugas", "nama peserta", "nama_petugas");
  const jabatan = get("jabatan", "posisi", "jenis petugas", "role", "kategori");
  const jabatanUpper = jabatan.toUpperCase();
  const isPml = /PML|PENGAWAS/.test(jabatanUpper);
  const isPpl = /PPL|PENCACAH/.test(jabatanUpper);
  const email = get("email", "email petugas", "email peserta", "mail");
  const emailPengawas = get("email pengawas", "email pml", "mail pengawas", "mail pml");
  const namaPengawas = get("nama pengawas", "nama pml", "pengawas");
  const namaPplKolom = get("nama ppl", "pencacah", "nama pencacah");
  const nomorSpk = get("nomor spk", "nomor_spk", "nomor kontrak", "nomor_kontrak", "spk");
  const slsOngoing = get(
    "sls (selesai + sedang dikerjakan)",
    "sls selesai + sedang dikerjakan",
    "sls (selesai dan sedang dikerjakan)",
    "sls selesai dan sedang dikerjakan",
    "sls ongoing",
    "sls_ongoing"
  );
  const prelistTotal = get("prelist total", "prelist_total", "target prelist", "target_prelist");
  const realisasiTotal = get(
    "realisasi total",
    "realisasi tot",
    "realisasi_total",
    "realisasi hasil pendataan",
    "realisasi_hasil",
    "realisasi",
    "realisasi_hasil_pendataan"
  );
  const persentasePrelist = get(
    "persentase prelist",
    "persentase pendataan",
    "persentase realisasi",
    "persentase",
    "persentase p",
    "persentase_p",
    "persentase_prelist"
  );

  return {
    no: get("no", "nomor"),
    nama,
    jabatan,
    jabatan_raw: jabatan,
    wilayah: get("wilayah", "wil tugas", "wil. tugas", "wilayah tugas", "kecamatan", "asal"),
    kecamatan: get("kecamatan", "wilayah", "wil tugas", "wil. tugas", "wilayah tugas", "asal"),
    email,
    username_sobat: get("username sobat", "username_sobat", "username", "sobat id", "sobatid"),
    sobat_id: get("sobat id", "sobatid", "username sobat", "username_sobat", "username"),
    kelas: get("kelas"),
    gelombang: get("gelombang"),
    tempat: get("tempat", "hotel", "tc"),
    telp: get("telp", "no hp", "nomor hp"),

    // Untuk baris PML, identitas PML adalah Nama Lengkap/Email baris itu sendiri.
    // Untuk baris PPL, nama dan email pengawas diambil dari kolom relasi pada sheet yang sama.
    nama_pengawas: namaPengawas || (isPml ? nama : ""),
    email_pengawas: emailPengawas || (isPml ? email : ""),
    nama_pml: namaPengawas || (isPml ? nama : ""),
    email_pml: emailPengawas || (isPml ? email : ""),
    nama_ppl: namaPplKolom || (isPpl ? nama : ""),
    email_ppl: isPpl ? email : get("email ppl", "email pencacah"),

    nik: get("nik", "nik petugas", "nik peserta"),
    nomor_spk: nomorSpk,
    nomor_kontrak: nomorSpk,

    sls_total: get("sls total", "sls_total"),
    sls_40: get("sls 40%", "sls 40", "sls_40", "sls40", "sls 40 persen"),
    sls_60: get("sls 60%", "sls 60", "sls_60", "sls60", "sls 60 persen"),
    sls_ongoing: slsOngoing,
    sls_selesai_sedang_dikerjakan: slsOngoing,
    persentase_sls: get("persentase sls", "persentase_sls", "% sls"),
    tanggal_screenshot: get("tanggal screenshot", "tanggal_screenshot"),

    prelist_total: prelistTotal,
    target_prelist: prelistTotal,
    realisasi_total: realisasiTotal,
    realisasi_hasil_pendataan: realisasiTotal,
    persentase_prelist: persentasePrelist,
    persentase_pendataan: persentasePrelist,
    flag: get("flag", "status flag"),
  };
}

export function parseBappData(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = workbook.SheetNames.find(
    (name) => String(name ?? "").trim().toLowerCase() === "pembayaran"
  );

  if (!sheetName) {
    console.warn("Sheet bernama 'Pembayaran' tidak ditemukan; BAPP tidak akan memuat data.");
    return [];
  }

  const sheet = workbook.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  return raw.map(normalizeBappRow).filter((r) => r.nama || r.email || r.jabatan);
}

// ─── STATUS SLS PARSER ───────────────────────────────────────────────────────
// Sheet "Status SLS" dipakai khusus untuk menyaring tabel pada BERKAS
// PEMBAYARAN gabungan. Hanya SLS dengan status "Selesai" atau
// "Sedang Dikerjakan" yang akan dihitung dan ditampilkan.
export function normalizeStatusSlsCode(value, width = 0) {
  const digits = String(value ?? "").replace(/[^0-9]/g, "");
  if (!digits) return "";
  if (!width) return digits;
  return digits.padStart(width, "0").slice(-width);
}

export function normalizeStatusSlsLabel(value) {
  return String(value ?? "")
    .replace(/^\s*\[[^\]]*\]\s*/g, "")
    .replace(/^\s*[-:–—]+\s*/g, "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
}

export function isAllowedStatusSls(value) {
  const status = normalizeStatusSlsLabel(value);
  return status === "SELESAI" || status === "SEDANG DIKERJAKAN";
}

export function normalizeStatusSlsRow(row) {
  const normalized = {};
  Object.entries(row || {}).forEach(([key, value]) => {
    normalized[String(key ?? "").trim().toLowerCase().replace(/\s+/g, " ")] = String(value ?? "").trim();
  });

  const get = (...keys) => {
    for (const key of keys) {
      const value = normalized[key];
      if (value != null && String(value).trim() !== "") return String(value).trim();
    }
    return "";
  };

  const statusRaw = get("status", "status sls", "status_sls", "keterangan status", "keterangan");
  const kodeKecamatan = get("kode kecamatan", "kdkec", "kode_kecamatan", "kec");
  const kodeKelurahan = get("kode kelurahan", "kddesa", "kode_kelurahan", "desa");
  const kodeSls = get("kode sls", "kode_sls", "kdsls", "sls");

  return {
    nama_pml: get("nama pml", "nama_pml", "pengawas", "pml"),
    email_pml: get("email pml", "email_pml", "email pengawas", "mail pml", "mail pengawas"),
    nama_ppl: get("nama ppl", "nama_ppl", "pencacah", "ppl"),
    email_ppl: get("email ppl", "email_ppl", "email pencacah", "mail ppl", "mail pencacah"),

    kdkec: normalizeStatusSlsCode(kodeKecamatan, 3),
    kddesa: normalizeStatusSlsCode(kodeKelurahan, 3),
    kode_sls: normalizeStatusSlsCode(kodeSls, 6),

    jumlah_prelist: get("jumlah prelist", "prelist total", "prelist_total", "target prelist"),
    jumlah_realisasi: get("jumlah realisasi", "realisasi total", "realisasi_total", "realisasi"),
    persentase: get("persentase", "persentase pendataan", "persentase_pendataan"),
    status: normalizeStatusSlsLabel(statusRaw),
    status_raw: statusRaw,
    status_diperbolehkan: isAllowedStatusSls(statusRaw),
  };
}

export function parseStatusSlsData(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = workbook.SheetNames.find(
    (name) => String(name ?? "").trim().toLowerCase() === "status sls"
  );

  if (!sheetName) {
    console.warn("Sheet bernama 'Status SLS' tidak ditemukan; tabel gabungan tidak akan memakai filter status.");
    return [];
  }

  const sheet = workbook.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  return raw
    .map(normalizeStatusSlsRow)
    .filter((row) => row.nama_pml || row.nama_ppl || row.email_pml || row.email_ppl || row.kode_sls || row.status);
}

// ─── DATA PER SLS PARSER ─────────────────────────────────────────────────────
// Sheet "Data per SLS" memiliki tiga baris header bertingkat. Karena itu parser
// membaca nilai berdasarkan posisi kolom A:T, bukan berdasarkan nama header.
export function cleanDataPerSlsCell(value) {
  const text = String(value ?? "").trim();
  if (!text || /^#(?:N\/A|VALUE!|REF!|DIV\/0!|NAME\?|NUM!|NULL!)$/i.test(text)) return "";
  return text;
}

export function normalizeDataPerSlsRowFromArray(row = [], fallbackNo = 0) {
  const get = (index) => cleanDataPerSlsCell(row?.[index]);
  const keterangan = get(19);

  return {
    no_sumber: get(0) || String(fallbackNo || ""),
    nama_pml: get(1),
    username_pml: get(2),
    nama_ppl: get(3),
    username_ppl: get(4),

    kdkec: normalizeStatusSlsCode(get(5), 3),
    kddesa: normalizeStatusSlsCode(get(6), 3),
    kode_sls: normalizeStatusSlsCode(get(7), 6),

    target_keluarga: get(8),
    target_usaha: get(9),
    target_jumlah: get(10),

    realisasi_dengan_tidak_ditemukan_keluarga: get(11),
    realisasi_dengan_tidak_ditemukan_usaha: get(12),
    realisasi_dengan_tidak_ditemukan_jumlah: get(13),
    persentase_dengan_tidak_ditemukan: get(14),

    realisasi_tanpa_tidak_ditemukan_keluarga: get(15),
    realisasi_tanpa_tidak_ditemukan_usaha: get(16),
    realisasi_tanpa_tidak_ditemukan_jumlah: get(17),
    persentase_tanpa_tidak_ditemukan: get(18),

    keterangan,
    status: normalizeStatusSlsLabel(keterangan),
  };
}

export function parseDataPerSlsValues(values = []) {
  const rows = Array.isArray(values) ? values : [];

  // Cari baris data pertama. Baris nomor kolom seperti -1, -2, dan seterusnya
  // tidak dianggap data karena tidak memuat nama petugas.
  const firstDataIndex = rows.findIndex((row) => {
    const no = cleanDataPerSlsCell(row?.[0]);
    const namaPml = cleanDataPerSlsCell(row?.[1]);
    const namaPpl = cleanDataPerSlsCell(row?.[3]);
    const hasPersonName = /[A-Za-z]/.test(`${namaPml} ${namaPpl}`);
    const hasCode = [5, 6, 7].some((index) => /\d/.test(cleanDataPerSlsCell(row?.[index])));
    return /^\d+(?:[.,]0+)?$/.test(no) && hasPersonName && hasCode;
  });

  if (firstDataIndex < 0) return [];

  return rows
    .slice(firstDataIndex)
    .map((row, index) => normalizeDataPerSlsRowFromArray(row, index + 1))
    .filter((row) =>
      row.nama_pml || row.nama_ppl || row.username_pml || row.username_ppl ||
      row.kdkec || row.kddesa || row.kode_sls
    );
}

export function parseDataPerSlsData(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = workbook.SheetNames.find(
    (name) => String(name ?? "").trim().toLowerCase() === "data per sls"
  );

  if (!sheetName) {
    console.warn("Sheet bernama 'Data per SLS' tidak ditemukan; tabel beban kerja pada berkas pembayaran akan kosong.");
    return [];
  }

  const sheet = workbook.Sheets[sheetName];
  const values = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false,
    blankrows: false,
  });
  return parseDataPerSlsValues(values);
}

// ─── DATA PML PROGRESS PARSER ────────────────────────────────────────────────
// Sheet "Data PML Progress" juga punya header bertingkat seperti "Data per SLS",
// jadi dibaca berdasarkan posisi kolom (array), bukan nama header.
// Kolom L (index 11, 0-based) = "Realisasi Jumlah (Dengan Tidak Ditemukan) > Jumlah"
// itulah nilai yang dipakai sebagai realisasi PML untuk Surat Kepala.
export function normalizeDataPmlProgressRowFromArray(row = [], fallbackNo = 0) {
  const get = (index) => cleanDataPerSlsCell(row?.[index]);
  return {
    no: get(0) || String(fallbackNo || ""),
    nama_pml: get(1),
    username_sobat_pml: get(2),
    beban_sls_1: get(3),
    beban_sls_2: get(4),
    persentase_sls: get(5),
    target_prelist_awal: get(6),
    kolom_h: get(7),
    jumlah_target: get(8),      // kolom I
    target_keluarga: get(9),    // kolom J
    target_usaha: get(10),      // kolom K
    realisasi_jumlah: get(11),  // kolom L <-- REALISASI YANG DIPAKAI
    persentase_realisasi: get(12), // kolom M
    keterangan: get(16),        // kolom Q
  };
}

export function parseDataPmlProgressValues(values = []) {
  const rows = Array.isArray(values) ? values : [];

  // Baris data pertama = baris yang kolom No-nya angka DAN kolom Nama berisi huruf.
  const firstDataIndex = rows.findIndex((row) => {
    const no = cleanDataPerSlsCell(row?.[0]);
    const nama = cleanDataPerSlsCell(row?.[1]);
    const hasPersonName = /[A-Za-z]/.test(nama);
    return /^\d+(?:[.,]0+)?$/.test(no) && hasPersonName;
  });

  if (firstDataIndex < 0) return [];

  return rows
    .slice(firstDataIndex)
    .map((row, index) => normalizeDataPmlProgressRowFromArray(row, index + 1))
    .filter((row) => row.nama_pml);
}

export function parseDataPmlProgressData(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = workbook.SheetNames.find(
    (name) => String(name ?? "").trim().toLowerCase() === "data pml progress"
  );

  if (!sheetName) {
    console.warn("Sheet bernama 'Data PML Progress' tidak ditemukan; realisasi PML pada Surat Kepala akan pakai fallback lama.");
    return [];
  }

  const sheet = workbook.Sheets[sheetName];
  const values = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false,
    blankrows: false,
  });
  return parseDataPmlProgressValues(values);
}

// ─── APPROVE BY PML PARSER ──────────────────────────────────────────────────
// Untuk berkas PML, nilai realisasi/jumlah pemeriksaan bersumber dari kolom
// "Jumlah Approve PML" pada sheet "Approve by PML".
export function normalizeApproveByPmlRow(row = {}) {
  const normalized = {};
  Object.entries(row || {}).forEach(([key, value]) => {
    const normalizedKey = String(key ?? "")
      .trim()
      .toLowerCase()
      .replace(/[_-]+/g, " ")
      .replace(/\s+/g, " ");
    normalized[normalizedKey] = cleanDataPerSlsCell(value);
  });

  const get = (...keys) => {
    for (const key of keys) {
      const value = normalized[String(key ?? "").trim().toLowerCase().replace(/[_-]+/g, " ").replace(/\s+/g, " ")];
      if (cleanDataPerSlsCell(value)) return cleanDataPerSlsCell(value);
    }
    return "";
  };

  const namaSls = get("nama sls", "sls", "nama_sls");
  const kodeDalamKurung = String(namaSls).match(/\[\s*(\d{1,6})\s*\]/);
  const kodeSlsLangsung = get("kode sls", "kode_sls", "kdsls");
  const kodeSls = normalizeStatusSlsCode(
    kodeSlsLangsung || (kodeDalamKurung ? kodeDalamKurung[1] : ""),
    6
  );

  return {
    nama_sls: namaSls,
    kode_sls: kodeSls,
    email_ppl: get("email ppl", "email pencacah", "email_ppl"),
    nama_ppl: get("nama ppl", "nama pencacah", "nama_ppl"),
    email_pml: get("email pml", "email pengawas", "email_pml"),
    nama_pml: get("nama pml", "nama pengawas", "nama_pml"),
    selesai: get("selesai", "status selesai", "status"),
    tanggal_screen: get("tanggal screen", "tanggal screenshot", "tanggal"),
    jumlah_submit: get("jumlah submit", "jumlah_submit"),
    jumlah_approve_pml: get(
      "jumlah approve pml",
      "jumlah approve by pml",
      "jumlah_approve_pml",
      "approve pml"
    ),
    submitted_by: get("submitted by", "submitted_by", "pengirim"),
    submitted_by_role: get("submitted by 1", "submitted by_1", "role submitted by", "submitted by role"),
    waktu_submit: get("waktu submit", "waktu_submit"),
    catatan: get("catatan", "keterangan"),
    // 🔥 BARU: kolom "Foto Bukti" bisa berisi lebih dari satu link (dipisah baris
    // baru/koma/titik koma). Disimpan sebagai array URL, siap dipakai loop gambar.
    foto_bukti: splitFotoBuktiUrls(get("foto bukti", "foto_bukti", "link foto", "foto")),
  };
}

// ─── FOTO BUKTI (GOOGLE DRIVE) ───────────────────────────────────────────────
// ─── FOTO BUKTI DARI SPREADSHEET TERPISAH (Database SLS) ────────────────────
// Link foto tidak ada di sheet "Approve by PML" yang dipakai utama, tapi ada di
// spreadsheet lain bernama "Database SLS [JANGAN DIUBAH]", tab "Submission-V2".
// 🔥 BARU: tab ini memisahkan foto jadi 2 kolom sendiri-sendiri, "Foto Bukti PPL"
// dan "Foto Bukti PML" (sebelumnya cuma satu kolom "Foto Bukti" gabungan).
// Data ini diambil terpisah lalu digabungkan ke approveByPmlRows berdasarkan
// Email PML + Email PPL.
export const FOTO_BUKTI_SPREADSHEET_ID = "1U694SejnIYezDRgy6Ao_1Moik4ckW7iMBWJeOmgpkcI";
export const FOTO_BUKTI_SPREADSHEET_EXPORT_URL = `https://docs.google.com/spreadsheets/d/${FOTO_BUKTI_SPREADSHEET_ID}/export?format=xlsx`;
export const FOTO_BUKTI_SHEET_NAME = "Submission-V2-Testing";

// Cache supaya spreadsheet foto tidak di-fetch berulang kali dalam satu sesi.
export let fotoBuktiDatabaseSlsCache = null;

export function normalizeFotoBuktiRow(row = {}) {
  const normalized = {};
  Object.entries(row || {}).forEach(([key, value]) => {
    const normalizedKey = String(key ?? "").trim().toLowerCase().replace(/[_-]+/g, " ").replace(/\s+/g, " ");
    normalized[normalizedKey] = cleanDataPerSlsCell(value);
  });
  const get = (...keys) => {
    for (const key of keys) {
      const value = normalized[String(key ?? "").trim().toLowerCase().replace(/[_-]+/g, " ").replace(/\s+/g, " ")];
      if (cleanDataPerSlsCell(value)) return cleanDataPerSlsCell(value);
    }
    return "";
  };
  return {
    email_ppl: get("email ppl", "email_ppl"),
    email_pml: get("email pml", "email_pml"),
    // 🔥 BARU: dipisah per role. foto_bukti_ppl dari kolom "Foto Bukti PPL",
    // foto_bukti_pml dari kolom "Foto Bukti PML".
    foto_bukti_ppl: splitFotoBuktiUrls(get("foto bukti ppl", "foto_bukti_ppl")),
    foto_bukti_pml: splitFotoBuktiUrls(get("foto bukti pml", "foto_bukti_pml")),
  };
}

export async function fetchFotoBuktiRowsFromDatabaseSls() {
  if (fotoBuktiDatabaseSlsCache) return fotoBuktiDatabaseSlsCache;

  const response = await fetch(FOTO_BUKTI_SPREADSHEET_EXPORT_URL);
  if (!response.ok) throw new Error(`Gagal memuat Database SLS: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });

  const sheetName = workbook.SheetNames.find(
    (name) => String(name ?? "").trim().toLowerCase() === FOTO_BUKTI_SHEET_NAME.toLowerCase()
  );
  if (!sheetName) {
    console.warn(`Sheet '${FOTO_BUKTI_SHEET_NAME}' tidak ditemukan di Database SLS. Sheet tersedia:`, workbook.SheetNames);
    fotoBuktiDatabaseSlsCache = [];
    return fotoBuktiDatabaseSlsCache;
  }

  const sheet = workbook.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  const rows = raw.map(normalizeFotoBuktiRow).filter((row) => row.email_pml || row.email_ppl);

  // 🔥 DIAGNOSTIK: supaya kelihatan di console kalau Database SLS ternyata
  // ke-fetch tapi isinya nol baris yang punya email (header kolom beda nama,
  // sheet kosong, dll) — sebelumnya kegagalan ini "diam" dan baru ketahuan
  // belakangan lewat foto_bukti_pml/ppl yang undefined di approveByPmlRows.
  console.log(
    `Database SLS: ${raw.length} baris mentah, ${rows.length} baris punya email PML/PPL.`
  );

  fotoBuktiDatabaseSlsCache = rows;
  return rows;
}

// Map utama: "EMAIL_PML::EMAIL_PPL" -> { fotoPml: Set, fotoPpl: Set } (satu
// pasangan PML+PPL bisa punya beberapa foto dari beberapa baris SLS berbeda).
// Map cadangan (fallback): per email PPL saja dan per email PML saja.
// 🔥 FIX: sebelumnya HANYA memakai kunci gabungan "EMAIL_PML::EMAIL_PPL", yang
// mengharuskan KEDUA email cocok persis pada baris Database SLS yang sama.
// Di praktiknya banyak baris submission di Database SLS hanya mengisi salah
// satu email (mis. Email PPL terisi tapi Email PML kosong/beda ejaan dengan
// sheet Approve by PML), sehingga hampir semua baris gagal cocok dan foto
// selalu kosong walau datanya sebenarnya ada. Sekarang dicoba pasangan persis
// dulu, lalu fallback ke pencocokan per-email saja.
function normalizePhotoKey(value) {
  return String(value ?? "").trim().toLowerCase();
}

function mergePhotoArrays(...values) {
  const merged = new Set();
  for (const list of values) {
    for (const url of list || []) {
      const cleaned = String(url ?? "").trim();
      if (cleaned) merged.add(cleaned);
    }
  }
  return [...merged];
}

export function buildFotoBuktiMapByPmlPpl(fotoBuktiRows = []) {
  const pairMap = new Map();
  const byPplEmail = new Map();
  const byPmlEmail = new Map();
  const byPplName = new Map();
  const byPmlName = new Map();
  // 🔧 FIX: banyak baris "Database SLS" hanya mengisi salah satu email (paling
  // sering Email PPL terisi, Email PML kosong/beda ejaan — lihat komentar di
  // mergeFotoBuktiIntoApproveByPmlRows). Sebelumnya foto_bukti_pml HANYA bisa
  // ditemukan lewat byPmlEmail, jadi kalau Email PML pada baris itu kosong,
  // foto PML-nya hilang sama sekali walau Email PPL pada baris yang sama valid.
  // Map tambahan ini mengaitkan foto_bukti_pml lewat Email PPL pada baris yang
  // sama, supaya foto PML tetap ketemu lewat PPL yang dia bawahi.
  const byPplEmailToFotoPml = new Map();
  // Simetris: foto_bukti_ppl juga bisa dicari lewat Email PML pada baris yang sama.
  const byPmlEmailToFotoPpl = new Map();

  for (const row of fotoBuktiRows || []) {
    const emailPml = normalizePhotoKey(row.email_pml);
    const emailPpl = normalizePhotoKey(row.email_ppl);
    const namePml = normalizePhotoKey(row.nama_pml || row.nama_pengawas || row.nama || "");
    const namePpl = normalizePhotoKey(row.nama_ppl || row.nama || "");
    if (!emailPml && !emailPpl && !namePml && !namePpl) continue;

    const pairKey = `${emailPml}::${emailPpl}`;
    if (!pairMap.has(pairKey)) pairMap.set(pairKey, { fotoPml: new Set(), fotoPpl: new Set() });
    const pairEntry = pairMap.get(pairKey);
    for (const url of row.foto_bukti_pml || row.foto_bukti || []) pairEntry.fotoPml.add(String(url).trim());
    for (const url of row.foto_bukti_ppl || row.foto_bukti || []) pairEntry.fotoPpl.add(String(url).trim());

    if (emailPpl) {
      if (!byPplEmail.has(emailPpl)) byPplEmail.set(emailPpl, new Set());
      const entry = byPplEmail.get(emailPpl);
      for (const url of row.foto_bukti_ppl || row.foto_bukti || []) entry.add(String(url).trim());

      if (!byPplEmailToFotoPml.has(emailPpl)) byPplEmailToFotoPml.set(emailPpl, new Set());
      const fotoPmlEntry = byPplEmailToFotoPml.get(emailPpl);
      for (const url of row.foto_bukti_pml || row.foto_bukti || []) fotoPmlEntry.add(String(url).trim());
    }
    if (emailPml) {
      if (!byPmlEmail.has(emailPml)) byPmlEmail.set(emailPml, new Set());
      const entry = byPmlEmail.get(emailPml);
      for (const url of row.foto_bukti_pml || row.foto_bukti || []) entry.add(String(url).trim());

      if (!byPmlEmailToFotoPpl.has(emailPml)) byPmlEmailToFotoPpl.set(emailPml, new Set());
      const fotoPplEntry = byPmlEmailToFotoPpl.get(emailPml);
      for (const url of row.foto_bukti_ppl || row.foto_bukti || []) fotoPplEntry.add(String(url).trim());
    }
    if (namePpl) {
      if (!byPplName.has(namePpl)) byPplName.set(namePpl, new Set());
      const entry = byPplName.get(namePpl);
      for (const url of row.foto_bukti_ppl || row.foto_bukti || []) entry.add(String(url).trim());
    }
    if (namePml) {
      if (!byPmlName.has(namePml)) byPmlName.set(namePml, new Set());
      const entry = byPmlName.get(namePml);
      for (const url of row.foto_bukti_pml || row.foto_bukti || []) entry.add(String(url).trim());
    }
  }

  return { pairMap, byPplEmail, byPmlEmail, byPplName, byPmlName, byPplEmailToFotoPml, byPmlEmailToFotoPpl };
}

export function mergeFotoBuktiIntoApproveByPmlRows(approveByPmlRows = [], fotoBuktiMap) {
  const isEmptyMap =
    !fotoBuktiMap ||
    ((fotoBuktiMap.pairMap?.size || 0) === 0 &&
      (fotoBuktiMap.byPplEmail?.size || 0) === 0 &&
      (fotoBuktiMap.byPmlEmail?.size || 0) === 0 &&
      (fotoBuktiMap.byPplName?.size || 0) === 0 &&
      (fotoBuktiMap.byPmlName?.size || 0) === 0 &&
      (fotoBuktiMap.byPplEmailToFotoPml?.size || 0) === 0 &&
      (fotoBuktiMap.byPmlEmailToFotoPpl?.size || 0) === 0);
  if (isEmptyMap) return approveByPmlRows;

  return (approveByPmlRows || []).map((row) => {
    const emailPml = normalizePhotoKey(row.email_pml);
    const emailPpl = normalizePhotoKey(row.email_ppl);
    const namePml = normalizePhotoKey(row.nama_pml || row.nama_pengawas || row.nama || "");
    const namePpl = normalizePhotoKey(row.nama_ppl || row.nama || "");
    const pairEntry = fotoBuktiMap.pairMap.get(`${emailPml}::${emailPpl}`);

    const fotoPplSource =
      pairEntry?.fotoPpl?.size ? pairEntry.fotoPpl
      : (emailPpl && fotoBuktiMap.byPplEmail.get(emailPpl))
      || (namePpl && fotoBuktiMap.byPplName.get(namePpl))
      // 🔧 FIX: fallback terakhir lewat Email PML pada baris yang sama, untuk
      // kasus Email PPL kosong/beda ejaan di Database SLS tapi Email PML valid.
      || (emailPml && fotoBuktiMap.byPmlEmailToFotoPpl.get(emailPml))
      || null;
    const fotoPmlSource =
      pairEntry?.fotoPml?.size ? pairEntry.fotoPml
      : (emailPml && fotoBuktiMap.byPmlEmail.get(emailPml))
      || (namePml && fotoBuktiMap.byPmlName.get(namePml))
      // 🔧 FIX UTAMA: fallback lewat Email PPL pada baris yang sama. Ini yang
      // paling sering dipakai karena di Database SLS, Email PPL jauh lebih
      // sering terisi benar dibanding Email PML — sebelumnya foto PML selalu
      // gagal ditemukan kalau Email PML pada baris submission kosong/salah,
      // padahal Email PPL-nya (dan foto_bukti_pml di baris yang sama) valid.
      || (emailPpl && fotoBuktiMap.byPplEmailToFotoPml.get(emailPpl))
      || null;

    const mergedPml = mergePhotoArrays(row.foto_bukti_pml || row.foto_bukti || [], fotoPmlSource || []);
    const mergedPpl = mergePhotoArrays(row.foto_bukti_ppl || row.foto_bukti || [], fotoPplSource || []);

    const result = { ...row };
    if (mergedPml.length > 0) result.foto_bukti_pml = mergedPml;
    else if (row.foto_bukti) result.foto_bukti_pml = row.foto_bukti;
    if (mergedPpl.length > 0) result.foto_bukti_ppl = mergedPpl;
    else if (row.foto_bukti) result.foto_bukti_ppl = row.foto_bukti;
    return result;
  });
}

export async function enrichApproveByPmlWithFotoBukti(approveByPmlRows = []) {
  try {
    const fotoBuktiRows = await fetchFotoBuktiRowsFromDatabaseSls();
    const fotoBuktiMap = buildFotoBuktiMapByPmlPpl(fotoBuktiRows);
    const merged = mergeFotoBuktiIntoApproveByPmlRows(approveByPmlRows, fotoBuktiMap);
    const rowsWithFoto = merged.filter(
      (row) => (row.foto_bukti_pml || []).length > 0 || (row.foto_bukti_ppl || []).length > 0
    ).length;
    console.log(
      `Enrich foto bukti: ${fotoBuktiRows.length} baris Database SLS, ${rowsWithFoto} dari ${merged.length} baris Approve by PML kebagian foto.`
    );
    return merged;
  } catch (err) {
    console.warn("Gagal memuat foto dari Database SLS:", err.message, err);
    return approveByPmlRows;
  }
}

export function splitFotoBuktiUrls(value) {
  return String(value ?? "")
    .split(/[\n,;]+/)
    .map((item) => item.trim())
    .filter((item) => /^https?:\/\//i.test(item));
}

export function chunkFotoBuktiIntoRows(fotoEntries = [], perRow = 3) {
  // Satu object = satu baris tabel Word dengan 3 kolom tetap.
  // Contoh 5 foto => baris 1: foto1-3, baris 2: foto4-5 + satu slot kosong.
  const entries = Array.isArray(fotoEntries)
    ? fotoEntries.filter(Boolean)
    : [];

  const rows = [];

  for (let i = 0; i < entries.length; i += perRow) {
    const slice = entries.slice(i, i + perRow);

    rows.push({
      foto1: getFotoUrlFromTagValue(slice[0]),
      foto2: getFotoUrlFromTagValue(slice[1]),
      foto3: getFotoUrlFromTagValue(slice[2]),
    });
  }

  return rows;
}

export function extractGoogleDriveFileId(url) {
  const text = String(url ?? "");
  const patterns = [/\/d\/([a-zA-Z0-9_-]{15,})/, /[?&]id=([a-zA-Z0-9_-]{15,})/];
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) return match[1];
  }
  return "";
}

// Beberapa bentuk URL Google Drive dicoba berurutan karena tidak semua endpoint
// selalu mengizinkan akses langsung (CORS) dari browser.
export function buildGoogleDriveImageUrlCandidates(url) {
  const fileId = extractGoogleDriveFileId(url);
  const candidates = [];
  if (fileId) {
    candidates.push(`https://lh3.googleusercontent.com/d/${fileId}`);
    candidates.push(`https://drive.google.com/uc?export=view&id=${fileId}`);
    candidates.push(`https://drive.google.com/uc?export=download&id=${fileId}`);
  }
  if (url) candidates.push(url);
  return candidates;
}

// Cache supaya foto yang sama tidak diunduh berulang kali.
export const fotoBuktiArrayBufferCache = new Map();

// PNG transparan 1x1 sebagai fallback, dipakai bila sebuah foto gagal diunduh
// (link rusak/tidak publik) supaya proses generate dokumen tidak gagal total.
export const FALLBACK_FOTO_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";
export function base64ToArrayBuffer(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

export async function fetchFotoBuktiArrayBuffer(url) {
  const key = String(url ?? "").trim();
  if (!key) return null;
  if (fotoBuktiArrayBufferCache.has(key)) return fotoBuktiArrayBufferCache.get(key);

  let result = null;
  for (const candidateUrl of buildGoogleDriveImageUrlCandidates(key)) {
    try {
      const response = await fetch(candidateUrl);
      if (!response.ok) continue;
      const buffer = await response.arrayBuffer();
      if (buffer && buffer.byteLength > 0) {
        result = buffer;
        break;
      }
    } catch (err) {
      console.warn(`Gagal memuat foto bukti dari ${candidateUrl}:`, err);
    }
  }

  fotoBuktiArrayBufferCache.set(key, result);
  return result;
}


// ---------------------------------------------------------------------------
// FIX PPL: NORMALISASI TAG FOTO + PREFETCH SEBELUM DOCXTEMPLATER RENDER
// ---------------------------------------------------------------------------
// docxtemplater-image-module-free dapat bermasalah ketika getImage() async
// dipanggil di dalam loop gambar. Karena itu semua foto diunduh lebih dulu,
// disimpan di cache, kemudian image module membaca cache secara synchronous.
export function getFotoUrlFromTagValue(tagValue) {
  if (!tagValue) return "";

  if (typeof tagValue === "string") {
    return tagValue.trim();
  }

  if (typeof tagValue === "object" && tagValue.url) {
    return String(tagValue.url).trim();
  }

  return "";
}

export function collectFotoUrlsFromTemplateData(templateData = {}) {
  const urls = new Set();

  const addUrl = (value) => {
    const url = getFotoUrlFromTagValue(value);
    if (url && /^https?:\/\//i.test(url)) {
      urls.add(url);
    }
  };

  // Kompatibilitas dengan struktur foto lama.
  for (const item of templateData?.foto || []) {
    addUrl(item);
  }

  for (const item of templateData?.foto_bukti || []) {
    addUrl(item);
  }

  // Struktur grid baru: 3 foto per baris.
  for (const row of templateData?.foto_rows || []) {
    addUrl(row?.foto1);
    addUrl(row?.foto2);
    addUrl(row?.foto3);

    // Fallback untuk data lama yang masih memakai row.slot[].
    if (Array.isArray(row?.slot)) {
      for (const slot of row.slot) {
        addUrl(slot);
      }
    }
  }

  return [...urls];
}

export async function prefetchFotoBuktiForTemplate(templateData = {}) {
  const urls = collectFotoUrlsFromTemplateData(templateData);

  console.log(`Prefetch ${urls.length} foto bukti sebelum render DOCX`);

  if (urls.length === 0) return;

  await Promise.all(
    urls.map(async (url) => {
      try {
        const buffer = await fetchFotoBuktiArrayBuffer(url);
        if (!buffer) {
          console.warn("Foto gagal dimuat, akan memakai fallback transparan:", url);
        }
      } catch (err) {
        console.warn("Prefetch foto gagal, akan memakai fallback transparan:", url, err);
      }
    })
  );
}

export const DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

// 🔥 FIX: setelah render lewat renderAsync + modul gambar, doc.getZip().generate({type:"blob"})
// kadang menghasilkan objek yang gagal dikenali sebagai Blob asli oleh browser
// (URL.createObjectURL melempar "Overload resolution failed"). Solusinya, ambil hasilnya
// sebagai arraybuffer lalu bungkus manual pakai konstruktor Blob bawaan browser.
export function zipToDocxBlob(zip) {
  const arrayBuffer = zip.generate({ type: "arraybuffer" });
  return new Blob([arrayBuffer], { type: DOCX_MIME_TYPE });
}

// Modul gambar docxtemplater — dibuat lewat fungsi (bukan instance tunggal)
// supaya aman dipakai berulang untuk banyak dokumen dalam satu batch download.
export function createFotoBuktiImageModule() {
  return new ImageModule({
    centered: false,

    // PENTING: getImage synchronous.
    // Semua foto sudah di-prefetch sebelum doc.render().
    getImage: (tagValue) => {
      const url = getFotoUrlFromTagValue(tagValue);

      if (!url) {
        return base64ToArrayBuffer(FALLBACK_FOTO_BASE64);
      }

      const cached = fotoBuktiArrayBufferCache.get(url);

      if (cached instanceof ArrayBuffer && cached.byteLength > 0) {
        return cached;
      }

      if (ArrayBuffer.isView(cached) && cached.byteLength > 0) {
        return cached;
      }

      console.warn("Foto belum tersedia di cache, memakai fallback transparan:", url);
      return base64ToArrayBuffer(FALLBACK_FOTO_BASE64);
    },

    getSize: (img, tagValue) => {
      const url = getFotoUrlFromTagValue(tagValue);
      if (!url) return [1, 1];

      // Ukuran seragam agar tiga foto muat dan rapi dalam satu baris tabel Word.
      return [180, 135];
    },
  });
}

// PML => dicocokkan lewat Email/Nama PML (SELURUH foto PPL yang dia bawahi).
// PPL => dicocokkan lewat Email/Nama PPL (foto khusus miliknya sendiri).
export function matchApproveByPmlRowForRole(approveRow, row, role) {
  const isPml = upperText(role) === "PML";
  const targetEmail = upperText(cleanText(row?.email || ""));
  const targetName = upperText(cleanText(row?.nama || ""));
  const approveEmail = upperText(cleanText(isPml ? approveRow?.email_pml : approveRow?.email_ppl));
  const approveName = upperText(cleanText(isPml ? approveRow?.nama_pml : approveRow?.nama_ppl));
  if (targetEmail && approveEmail) return targetEmail === approveEmail;
  if (targetName && approveName) return targetName === approveName;
  return false;
}

export function filterApproveByPmlRowsForBappRow(approveByPmlRows = [], row = {}, role = "PML") {
  return (approveByPmlRows || []).filter((approveRow) => matchApproveByPmlRowForRole(approveRow, row, role));
}

// Kumpulkan URL unik. Template Word memakai {#foto_rows}{%foto1} | {%foto2} | {%foto3}{/foto_rows}.
// 🔥 BARU: role menentukan kolom foto yang dipakai — generate PML memakai
// foto_bukti_pml (foto dari SEMUA PPL di bawah PML tsb, karena approveRows yang
// dikirim ke sini sudah difilter per-PML sebelumnya), generate PPL memakai
// foto_bukti_ppl (foto milik PPL itu sendiri saja).
export function collectFotoBuktiFromApproveRows(approveRows = [], role = "PML") {
  const isPml = upperText(role) === "PML";
  const seen = new Set();
  const entries = [];
  for (const approveRow of approveRows || []) {
    const urls = isPml ? (approveRow?.foto_bukti_pml || approveRow?.foto_bukti || []) : (approveRow?.foto_bukti_ppl || approveRow?.foto_bukti || []);
    for (const url of urls || []) {
      const key = String(url ?? "").trim();
      if (!key || seen.has(key)) continue;
      seen.add(key);
      entries.push({ url: key });
    }
  }
  return entries;
}

export function parseApproveByPmlData(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = workbook.SheetNames.find(
    (name) => String(name ?? "").trim().toLowerCase() === "approve by pml"
  );

  if (!sheetName) {
    console.warn("Sheet bernama 'Approve by PML' tidak ditemukan; jumlah PML akan memakai fallback Data per SLS.");
    return [];
  }

  const sheet = workbook.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  return raw
    .map(normalizeApproveByPmlRow)
    .filter((row) =>
      row.nama_sls || row.kode_sls || row.nama_pml || row.email_pml ||
      row.nama_ppl || row.email_ppl || row.jumlah_approve_pml
    );
}

export function isBappRowForRole(row, role = "PML") {
  const jabatan = upperText(row?.jabatan || row?.jabatan_raw || "");
  if (role === "PML") return /PML|PENGAWAS/.test(jabatan);
  if (role === "PPL") return /PPL|PENCACAH/.test(jabatan);
  return true;
}

export function extractNomorPrefix(value) {
  const text = String(value ?? "").trim();
  if (!text) return "";
  const match = text.match(/(\d+)/);
  return match ? match[1] : "";
}

export function parseDateInput(dateStr) {
  if (!dateStr) return null;
  const parts = String(dateStr).split("-");
  if (parts.length === 3) {
    const [year, month, day] = parts.map(Number);
    if ([year, month, day].every((n) => Number.isFinite(n))) {
      return new Date(year, month - 1, day);
    }
  }
  const parsed = new Date(dateStr);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

export function numberToIndonesianWords(value) {
  const number = Number(value);
  if (!Number.isFinite(number) || number < 1 || number > 31) return "";
  const words = [
    "", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan",
    "sepuluh", "sebelas", "dua belas", "tiga belas", "empat belas", "lima belas", "enam belas",
    "tujuh belas", "delapan belas", "sembilan belas", "dua puluh", "dua puluh satu",
    "dua puluh dua", "dua puluh tiga", "dua puluh empat", "dua puluh lima", "dua puluh enam",
    "dua puluh tujuh", "dua puluh delapan", "dua puluh sembilan", "tiga puluh", "tiga puluh satu"
  ];
  return words[number];
}

export function getBappDateParts(dateStr) {
  const parsed = parseDateInput(dateStr);
  if (!parsed) return { hari_terbilang: "", tanggal_terbilang: "", tanggal: "", bulan: "", bulan_terbilang: "" };
  const monthNumber = String(parsed.getMonth() + 1);
  const monthNames = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
  ];
  return {
    hari_terbilang: parsed.toLocaleDateString("id-ID", { weekday: "long" }),
    tanggal_terbilang: numberToIndonesianWords(parsed.getDate()),
    tanggal: String(parsed.getDate()),
    bulan: monthNumber,
    bulan_terbilang: monthNames[parsed.getMonth()] || "",
  };
}

export function getBappIdentityKey(row, role = "PML") {
  const email = cleanText(row?.email || "");
  if (email) return `EMAIL::${upperText(email)}`;
  const name = cleanText(row?.nama || "");
  const jabatan = cleanText(row?.jabatan_raw || row?.jabatan || role || "");
  return `NAME::${upperText(name)}::${upperText(jabatan)}`;
}

export function dedupeBappRows(rows = []) {
  const seen = new Set();
  return (rows || []).filter((row) => {
    const key = getBappIdentityKey(row);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

export function normalizeLampiranRow(row) {
  const normalized = {};

  Object.entries(row).forEach(([k, v]) => {
    const key = String(k ?? "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, " ");

    normalized[key] = String(v ?? "").trim();
  });

  const get = (...keys) => {
    for (const key of keys) {
      const value = normalized[key];
      if (value != null && String(value).trim() !== "") return String(value).trim();
    }
    return "";
  };

  // 🔥 FIX: jangan andalkan exact-match nama header untuk nomor kontrak — variasi
  // penulisan header di Google Sheet (titik, spasi, urutan kata, dll) gampang
  // membuat exact-match get() gagal dan hasilnya jadi kosong/undefined.
  // Cari secara fuzzy: kolom apa pun yang namanya mengandung "kontrak" DAN
  // mengandung "pml" (atau "ppl") akan dipakai, berapa pun variasi penulisannya.
  const findKontrakFuzzy = (token) => {
    for (const [key, value] of Object.entries(normalized)) {
      if (key.includes("kontrak") && key.includes(token) && String(value ?? "").trim() !== "") {
        return String(value).trim();
      }
    }
    return "";
  };

  return {
    no: get("no"),

    // Struktur sheet Lampiran dari Google Sheet:
    // PENGAWAS = PML, PENCACAH = PPL.
    nama_pml: get("pengawas", "nama pml", "nama_pml", "pml"),
    nama_ppl: get("pencacah", "nama ppl", "nama_ppl", "ppl", "nama petugas lapangan sensus"),

    email_pengawas: get("email pengawas", "mail pengawas", "email pml"),
    email_pencacah: get("email pencacah", "mail pencacah", "email ppl"),

    kdprov: get("kdprov"),
    kdkab: get("kdkab"),
    kdkec: get("kdkec"),
    kddesa: get("kddesa"),
    kdsls: get("kdsls"),
    kdsubsls: get("kdsubsls"),
    kdsubslspanjang: get("kdsubsls_25_2", "kdsubsls_25", "kdsubsls panjang"),

    nmprov: get("nmprov"),
    nmkab: get("nmkab"),
    kecamatan: get("nmkec", "kecamatan", "kecamatan/distrik", "kecamatan / distrik").toUpperCase(),
    kelurahan: get("nmdesa", "kelurahan", "desa/kampung/nagari", "desa / kampung / nagari").toUpperCase(),
    sls: get("nmsls", "sls"),
    subsls: get("nmsubsls", "sub-sls", "sub sls"),

    // Kolom ini opsional. Kalau tidak ada, jumlah dihitung dari banyaknya baris.
    jumlah: get("jumlah", "jumlah sls/sub-sls", "jumlah sls/sub sls", "jumlah sls / sub-sls", "jumlah sls"),

    jabatan: get("jabatan").toUpperCase(),
    kelas: get("kelas"),
    gelombang: get("gelombang"),
    hotel: get("tc", "hotel", "tempat").toUpperCase(),
    tc: get("tc", "hotel", "tempat").toUpperCase(),

    // Nomor kontrak per jenis petugas. Dipakai untuk variabel {nomor_kontrak} di
    // template, sesuai jenis dokumen yang sedang digenerate (PML atau PPL).
    // Coba exact-match alias dulu, kalau tidak ketemu baru fallback ke fuzzy search.
    nomor_kontrak_pml:
      get("no kontrak pml", "nomor kontrak pml", "no_kontrak_pml", "kontrak pml", "no. kontrak pml") ||
      findKontrakFuzzy("pml"),
    nomor_kontrak_ppl:
      get("no kontrak ppl", "nomor kontrak ppl", "no_kontrak_ppl", "kontrak ppl", "no. kontrak ppl") ||
      findKontrakFuzzy("ppl"),
  };
}

// 🔥 FIX: Google Sheet sumber Lampiran biasanya pakai "merged cell" secara visual —
// nama Pengawas/Pencacah, kecamatan, kelurahan, dst hanya diisi SEKALI di baris pertama
// tiap blok, lalu baris-baris SLS berikutnya di bawahnya dikosongkan.
// XLSX.utils.sheet_to_json TIDAK menurunkan nilai merged cell, jadi baris-baris itu
// terbaca kosong dan akhirnya DIBUANG oleh generateLampiran() (karena identity-nya kosong).
// Ini sebabnya 1 PML yang sebenarnya membawahi 7-8 SLS/PPL hanya muncul 2-3 baris saja.
//
// Solusinya: forward-fill — isi sel kosong dengan nilai terakhir yang valid di kolom
// yang sama, KHUSUS untuk kolom yang memang lazim merged (nama petugas, lokasi, dll).
// Kolom kode SLS/Sub-SLS sengaja TIDAK di-forward-fill karena itu harus unik per baris.
export const LAMPIRAN_FORWARD_FILL_KEYS = [
  "nama_pml", "nama_ppl",
  "email_pengawas", "email_pencacah",
  "kdprov", "kdkab", "kdkec", "kddesa",
  "nmprov", "nmkab", "kecamatan", "kelurahan",
  "jabatan", "kelas", "gelombang", "hotel", "tc",
  "nomor_kontrak_pml", "nomor_kontrak_ppl",
];

export function forwardFillLampiranRows(rows) {
  const lastValue = {};
  return rows.map((row) => {
    const filled = { ...row };
    for (const key of LAMPIRAN_FORWARD_FILL_KEYS) {
      const value = cleanText(filled[key]);
      if (value) {
        lastValue[key] = value;
      } else if (lastValue[key]) {
        // Sel kosong karena merged cell -> turunkan nilai dari baris di atasnya
        filled[key] = lastValue[key];
      }
    }
    return filled;
  });
}

export function parseLampiranXlsxData(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });

  // Cari tab bernama Lampiran secara toleran:
  // - tidak sensitif kapital
  // - mengabaikan spasi di awal/akhir
  // Ini mencegah kasus tab "Lampiran " terbaca di console tetapi dianggap kosong di frontend.
  const sheetName = workbook.SheetNames.find(
    (name) => String(name ?? "").trim().toLowerCase() === "lampiran"
  );

  console.log("Daftar sheet terbaca:", workbook.SheetNames);

  if (!sheetName) {
    console.warn("Sheet bernama 'Lampiran' tidak ditemukan. Sheet tersedia:", workbook.SheetNames);
    return [];
  }

  const sheet = workbook.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });

  console.log("Raw Lampiran:", raw);

  const normalizedRows = raw.map(normalizeLampiranRow);

  // 🔥 Diagnostik: tampilkan contoh hasil pembacaan No Kontrak PML/PPL dari 3 baris
  // pertama, supaya kalau masih kosong/undefined, gampang dicek header apa saja yang
  // terbaca dari sheet vs nilai yang berhasil diambil.
  if (normalizedRows.length > 0) {
    console.log(
      "Lampiran: header mentah baris pertama ->",
      raw.length ? Object.keys(raw[0]) : []
    );
    console.log(
      "Lampiran: contoh hasil No Kontrak PML/PPL (3 baris pertama) ->",
      normalizedRows.slice(0, 3).map((r) => ({
        nama_pml: r.nama_pml,
        nama_ppl: r.nama_ppl,
        nomor_kontrak_pml: r.nomor_kontrak_pml,
        nomor_kontrak_ppl: r.nomor_kontrak_ppl,
      }))
    );
  }

  // 🔥 FIX: Google Sheet biasanya punya banyak baris kosong tambahan di ekor sheet
  // (range default jauh lebih panjang dari data aslinya). Kalau forward-fill langsung
  // dijalankan ke SEMUA baris (termasuk baris kosong di ekor), baris-baris kosong itu
  // akan "ketarik" nilai PENCACAH/kecamatan/kelurahan terakhir yang valid, sehingga
  // jumlah baris yang diproses bisa membengkak jadi ribuan baris palsu milik PPL
  // terakhir. Ini yang menyebabkan error "Array buffer allocation failed" saat
  // generate Lampiran PPL.
  //
  // Solusinya: saring dulu baris yang BENAR-BENAR baris data, berdasarkan kolom yang
  // TIDAK PERNAH di-forward-fill (kdsls/kdkec/kddesa/sls/subsls) — kolom ini aman
  // dipakai sebagai penanda baris asli, karena nilainya selalu apa adanya dari sheet,
  // bukan hasil "tebakan" forward-fill.
  const candidateRows = normalizedRows.filter((r) =>
    cleanText(r.kdsls) || cleanText(r.kdkec) || cleanText(r.kddesa) || cleanText(r.sls) || cleanText(r.subsls)
  );

  if (normalizedRows.length !== candidateRows.length) {
    console.log(
      `Lampiran: membuang ${normalizedRows.length - candidateRows.length} baris kosong/bukan-data sebelum forward-fill (dari ${normalizedRows.length} baris mentah).`
    );
  }

  // 🔥 FIX: turunkan nilai dari merged cell SEBELUM difilter,
  // supaya baris dengan nama_pml/nama_ppl kosong tidak ikut terbuang.
  const filledRows = forwardFillLampiranRows(candidateRows);

  const parsed = filledRows
    .filter((r) => r.nama_pml || r.nama_ppl || r.kecamatan || r.kelurahan || r.sls || r.subsls);

  console.log("Lampiran parsed rows (setelah forward-fill):", parsed);

  return parsed;
}

export function normalizeGoogleSheetUrl(url) {
  if (!url) return null;
  try {
    const parsed = new URL(url.trim());
    const sheetIdMatch = parsed.pathname.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!sheetIdMatch) return null;
    return {
      spreadsheetId: sheetIdMatch[1],
      exportUrl: `https://docs.google.com/spreadsheets/d/${sheetIdMatch[1]}/export?format=xlsx`,
    };
  } catch { return null; }
}

export function buildGoogleSheetsApiMetadataUrl(spreadsheetId, apiKey) {
  return `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?key=${encodeURIComponent(apiKey)}`;
}

export function buildGoogleSheetsApiValuesUrl(spreadsheetId, apiKey, sheetName) {
  return `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(sheetName)}?key=${encodeURIComponent(apiKey)}`;
}

export function parseGoogleSheetApiRows(values = []) {
  const headers = (values[0] || []).map((value) => String(value ?? "").trim());
  return values.slice(1).map((row) => {
    const record = {};
    headers.forEach((header, index) => {
      record[header] = String(row?.[index] ?? "");
    });
    return record;
  });
}

export function findAdministrasiSheetFromRows(sheetRows = []) {
  for (const sheet of sheetRows) {
    const rows = Array.isArray(sheet.rows) ? sheet.rows : [];
    if (rows.length === 0) continue;
    const headers = Object.keys(rows[0]).map((h) => String(h ?? "").trim());
    if (isAdministrasiSheet(headers)) {
      return { sheet, data: rows, headers, sheetName: sheet.sheetName };
    }
  }
  return null;
}

export function findLampiranSheetFromRows(sheetRows = []) {
  for (const sheet of sheetRows) {
    const rows = Array.isArray(sheet.rows) ? sheet.rows : [];
    if (rows.length === 0) continue;
    const headers = Object.keys(rows[0]).map((h) => String(h ?? "").trim());
    const normalizedHeaders = headers.map((h) => String(h ?? "").trim().toLowerCase());
    const hasPml = normalizedHeaders.some((h) => ["nama pml", "nama_pml", "pengawas", "pml"].includes(h));
    const hasPpl = normalizedHeaders.some((h) => ["nama ppl", "nama_ppl", "pencacah", "ppl"].includes(h));
    const hasKecamatan = normalizedHeaders.some((h) => ["kecamatan", "kecamatan/distrik", "kecamatan / distrik", "nmkec"].includes(h));
    const hasSls = normalizedHeaders.some((h) => ["sls", "nmsls"].includes(h));
    if (hasPml && hasPpl && hasKecamatan && hasSls) {
      return { sheet, data: rows, headers, sheetName: sheet.sheetName };
    }
  }
  return null;
}

export function findBappSheetFromRows(sheetRows = []) {
  for (const sheet of sheetRows) {
    const rows = Array.isArray(sheet.rows) ? sheet.rows : [];
    if (rows.length === 0) continue;
    const headers = Object.keys(rows[0]).map((h) => String(h ?? "").trim());
    const normalizedHeaders = headers.map((h) => String(h ?? "").trim().toLowerCase());
    const hasName = normalizedHeaders.some((h) => ["nama", "nama lengkap", "nama_lengkap", "nama petugas", "nama_petugas"].includes(h));
    const hasEmail = normalizedHeaders.some((h) => ["email", "email petugas", "email peserta", "mail"].includes(h));
    const hasJabatan = normalizedHeaders.some((h) => ["jabatan", "posisi", "jenis petugas", "role", "kategori"].includes(h));
    const hasNomorSpk = normalizedHeaders.some((h) => ["nomor spk", "nomor_spk", "nomor kontrak", "nomor_kontrak", "spk"].includes(h));
    if (hasName && hasEmail && hasJabatan && hasNomorSpk) {
      return { sheet, data: rows, headers, sheetName: sheet.sheetName };
    }
  }
  return null;
}

export function findStatusSlsSheetFromRows(sheetRows = []) {
  // Prioritaskan nama tab persis agar tidak tertukar dengan sheet lain.
  const exact = (sheetRows || []).find(
    (sheet) => String(sheet?.sheetName ?? "").trim().toLowerCase() === "status sls"
  );
  if (exact && Array.isArray(exact.rows) && exact.rows.length > 0) {
    const headers = Object.keys(exact.rows[0]).map((h) => String(h ?? "").trim());
    return { sheet: exact, data: exact.rows, headers, sheetName: exact.sheetName };
  }

  // Fallback berdasarkan struktur header.
  for (const sheet of sheetRows || []) {
    const rows = Array.isArray(sheet.rows) ? sheet.rows : [];
    if (rows.length === 0) continue;
    const headers = Object.keys(rows[0]).map((h) => String(h ?? "").trim());
    const keys = headers.map((h) => h.toLowerCase());
    const hasPml = keys.some((h) => ["nama pml", "nama_pml", "pengawas"].includes(h));
    const hasPpl = keys.some((h) => ["nama ppl", "nama_ppl", "pencacah"].includes(h));
    const hasKodeSls = keys.some((h) => ["kode sls", "kode_sls", "kdsls"].includes(h));
    const hasStatus = keys.some((h) => ["status", "status sls", "status_sls"].includes(h));
    if (hasPml && hasPpl && hasKodeSls && hasStatus) {
      return { sheet, data: rows, headers, sheetName: sheet.sheetName };
    }
  }
  return null;
}

export function findDataPerSlsSheetFromRows(sheetRows = []) {
  const exact = (sheetRows || []).find(
    (sheet) => String(sheet?.sheetName ?? "").trim().toLowerCase() === "data per sls"
  );
  if (exact && Array.isArray(exact.values) && exact.values.length > 0) return exact;

  // Fallback bila nama tab sedikit berubah, tetapi struktur header tetap sama.
  return (sheetRows || []).find((sheet) => {
    const headerText = (sheet?.values || [])
      .slice(0, 4)
      .flat()
      .map((value) => String(value ?? "").trim().toLowerCase())
      .join(" ");
    return headerText.includes("target prelist awal") &&
      headerText.includes("username sobat") &&
      headerText.includes("sls/sub-sls");
  }) || null;
}

export function findDataPmlProgressSheetFromRows(sheetRows = []) {
  const exact = (sheetRows || []).find(
    (sheet) => String(sheet?.sheetName ?? "").trim().toLowerCase() === "data pml progress"
  );
  if (exact && Array.isArray(exact.values) && exact.values.length > 0) return exact;

  return (sheetRows || []).find((sheet) => {
    const headerText = (sheet?.values || [])
      .slice(0, 4)
      .flat()
      .map((value) => String(value ?? "").trim().toLowerCase())
      .join(" ");
    return headerText.includes("username sobat") &&
      headerText.includes("realisasi jumlah") &&
      headerText.includes("dengan tidak ditemukan");
  }) || null;
}

export function findApproveByPmlSheetFromRows(sheetRows = []) {
  const exact = (sheetRows || []).find(
    (sheet) => String(sheet?.sheetName ?? "").trim().toLowerCase() === "approve by pml"
  );
  if (exact && Array.isArray(exact.rows) && exact.rows.length > 0) return exact;

  return (sheetRows || []).find((sheet) => {
    const rows = Array.isArray(sheet?.rows) ? sheet.rows : [];
    if (rows.length === 0) return false;
    const headers = Object.keys(rows[0] || {})
      .map((value) => String(value ?? "").trim().toLowerCase().replace(/[_-]+/g, " ").replace(/\s+/g, " "));
    const hasPml = headers.includes("nama pml") || headers.includes("email pml");
    const hasPpl = headers.includes("nama ppl") || headers.includes("email ppl");
    const hasApprove = headers.some((header) =>
      header === "jumlah approve pml" ||
      header === "jumlah approve by pml" ||
      header === "jumlah approve"
    );
    return hasPml && hasPpl && hasApprove;
  }) || null;
}

export function normalizeHeaderKeys(headers) {
  return headers.map((h) => String(h ?? "").trim().toLowerCase());
}

export function isAdministrasiSheet(headers) {
  const keys = normalizeHeaderKeys(headers);
  const hasName = keys.some((h) => ["nama", "nama lengkap", "nama_lengkap", "nama-lengkap"].includes(h));
  const hasJabatan = keys.some((h) => ["jabatan", "posisi"].includes(h));
  const hasKelas = keys.some((h) => ["kelas"].includes(h));
  const hasGelombang = keys.some((h) => ["gelombang"].includes(h));
  // "tc" dapat berupa kolom terpisah atau bagian dari header lain (seperti "Hotel")
  const hasTc = keys.some((h) => ["tc", "hotel", "tempat"].includes(h));
  return hasName && hasJabatan && hasKelas && hasGelombang && hasTc;
}

// Fungsi untuk menemukan sheet yang berisi data administrasi
export function findAdministrasiSheet(workbook) {
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
    if (data.length > 0) {
      const headers = Object.keys(data[0]).map((h) => String(h ?? "").trim());
      if (isAdministrasiSheet(headers)) {
        return { sheet, data, headers, sheetName };
      }
    }
  }
  return null;
}

export async function loadGoogleSheet(source, apiKey = "") {
  if (source?.spreadsheetId && apiKey) {
    const metadataResponse = await fetch(buildGoogleSheetsApiMetadataUrl(source.spreadsheetId, apiKey));
    if (!metadataResponse.ok) throw new Error(`Google Sheets API error: ${metadataResponse.status}`);

    const metadata = await metadataResponse.json();
    const sheets = Array.isArray(metadata.sheets) ? metadata.sheets : [];
    const sheetRows = [];

    for (const sheet of sheets) {
      const sheetName = sheet?.properties?.title;
      if (!sheetName) continue;
      const valuesResponse = await fetch(buildGoogleSheetsApiValuesUrl(source.spreadsheetId, apiKey, sheetName));
      if (!valuesResponse.ok) continue;
      const valuesData = await valuesResponse.json();
      const rows = parseGoogleSheetApiRows(valuesData.values || []);
      sheetRows.push({ sheetName, rows, values: valuesData.values || [] });
    }

    const adminSheetInfo = findAdministrasiSheetFromRows(sheetRows);
    let data = [];
    let rawHeaders = [];

    if (adminSheetInfo) {
      const { data: raw, headers: foundHeaders, sheetName: foundSheetName } = adminSheetInfo;
      console.log(`✓ Sheet data administrasi ditemukan via API: "${foundSheetName}"`);
      data = raw.map(normalizeRowHeaders).filter((r) => r.nama !== "" || r.nik !== "" || r.sobatId !== "");
      rawHeaders = foundHeaders;
    } else {
      const availableSheets = sheetRows.map((sheet) => sheet.sheetName).join(", ");
      console.warn(`Sheet data administrasi tidak ditemukan via API. Sheet tersedia: ${availableSheets}`);
    }

    const lampiranSheetInfo = findLampiranSheetFromRows(sheetRows);
    const lampiran = lampiranSheetInfo ? lampiranSheetInfo.data.map(normalizeLampiranRow) : [];

    const bappSheetInfo = findBappSheetFromRows(sheetRows);
    const bappData = bappSheetInfo ? bappSheetInfo.data.map(normalizeBappRow) : [];

    const statusSlsSheetInfo = findStatusSlsSheetFromRows(sheetRows);
    const statusSls = statusSlsSheetInfo
      ? statusSlsSheetInfo.data.map(normalizeStatusSlsRow).filter((row) => row.kode_sls || row.status || row.nama_ppl || row.email_ppl)
      : [];

    const dataPerSlsSheetInfo = findDataPerSlsSheetFromRows(sheetRows);
    const dataPerSls = dataPerSlsSheetInfo
      ? parseDataPerSlsValues(dataPerSlsSheetInfo.values || [])
      : [];

    const approveByPmlSheetInfo = findApproveByPmlSheetFromRows(sheetRows);
    const approveByPml = approveByPmlSheetInfo
      ? (approveByPmlSheetInfo.rows || []).map(normalizeApproveByPmlRow).filter((row) =>
          row.nama_sls || row.kode_sls || row.nama_pml || row.email_pml ||
          row.nama_ppl || row.email_ppl || row.jumlah_approve_pml
        )
      : [];

    const dataPmlProgressSheetInfo = findDataPmlProgressSheetFromRows(sheetRows);
    const dataPmlProgress = dataPmlProgressSheetInfo
      ? parseDataPmlProgressValues(dataPmlProgressSheetInfo.values || [])
      : [];

    console.log("Data PML Progress parsed rows (API):", dataPmlProgress);

    return { data, lampiran, bappData, statusSls, dataPerSls, approveByPml, dataPmlProgress, rawHeaders };
  }

  const response = await fetch(source?.exportUrl || source);
  if (!response.ok) throw new Error(`HTTP ${response.status}`);

  const arrayBuffer = await response.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });

  const adminSheetInfo = findAdministrasiSheet(workbook);
  let data = [];
  let rawHeaders = [];

  if (adminSheetInfo) {
    const { data: raw, headers: foundHeaders, sheetName: foundSheetName } = adminSheetInfo;
    console.log(`✓ Sheet data administrasi ditemukan: "${foundSheetName}"`);
    data = raw.map(normalizeRowHeaders).filter((r) => r.nama !== "" || r.nik !== "" || r.sobatId !== "");
    rawHeaders = foundHeaders;
  } else {
    const availableSheets = workbook.SheetNames.map((name, idx) => `[${idx}] ${name}`).join(", ");
    console.warn(`Sheet data administrasi tidak ditemukan. Lanjutkan hanya dengan data lampiran. Sheet tersedia: ${availableSheets}`);
  }

  const lampiran = parseLampiranXlsxData(arrayBuffer);
  const bappData = parseBappData(arrayBuffer);
  const statusSls = parseStatusSlsData(arrayBuffer);
  const dataPerSls = parseDataPerSlsData(arrayBuffer);
  const approveByPml = parseApproveByPmlData(arrayBuffer);
  const dataPmlProgress = parseDataPmlProgressData(arrayBuffer);

  return { data, lampiran, bappData, statusSls, dataPerSls, approveByPml, dataPmlProgress, rawHeaders };
}

