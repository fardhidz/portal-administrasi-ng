// ============================================================
// Portal Administrasi SE2026 — bagian: berkasPembayaran
// ============================================================

import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { saveAs } from "file-saver";

import { buildBastTemplateData } from "./bast";
import { buildBappTemplateData, downloadMultipleAsZip, formatKodeNama, groupLampiranRows, sanitizeFileName } from "./docGenerators";
import { cleanText, upperText } from "./helpers";
import {
  chunkFotoBuktiIntoRows,
  collectFotoBuktiFromApproveRows,
  createFotoBuktiImageModule,
  extractNomorPrefix,
  fetchHalamanDepanMapByEmail,
  findHalamanDepanUrl,
  isAllowedStatusSls,
  isBappRowForRole,
  normalizeStatusSlsCode,
  normalizeStatusSlsLabel,
  prefetchFotoBuktiForTemplate,
  zipToDocxBlob,
} from "./parsers";

// ─── BERKAS PEMBAYARAN PML/PPL ───────────────────────────────────────────────
// Satu record menyatukan identitas dari sheet Pembayaran dengan semua baris
// wilayah/SLS milik orang yang sama dari sheet Lampiran.
export function getBerkasIdentity(nama, email) {
  const emailClean = cleanText(email);
  if (emailClean) return `EMAIL::${upperText(emailClean)}`;
  return `NAME::${upperText(nama)}`;
}

export function buildBerkasPembayaranRecords(
  bappRows = [],
  lampiranRows = [],
  role = "PML",
  statusSlsRows = [],
  dataPerSlsRows = [],
  approveByPmlRows = []
) {
  const isPml = upperText(role) === "PML";
  const records = [];
  const byEmail = new Map();
  const byName = new Map();
  const ambiguousNames = new Set();

  const registerName = (record) => {
    const nameKey = upperText(record.displayName);
    if (!nameKey) return;
    if (byName.has(nameKey) && byName.get(nameKey) !== record) {
      ambiguousNames.add(nameKey);
      byName.delete(nameKey);
      return;
    }
    if (!ambiguousNames.has(nameKey)) byName.set(nameKey, record);
  };

  const ensureRecord = (displayName, email, bappRow = null) => {
    const identity = getBerkasIdentity(displayName, email);
    let record = records.find((item) => item.identity === identity);
    if (record) {
      if (!record.bappRow && bappRow) record.bappRow = bappRow;
      return record;
    }

    record = {
      identity,
      displayName: displayName || email || "Tanpa Nama",
      email,
      bappRow,
      lampiranRows: [],
      // Baris sheet Pembayaran yang terkait dengan record ini. Untuk PML,
      // isinya adalah baris-baris PPL di bawah pengawas tersebut. Untuk PPL,
      // isinya adalah baris Pembayaran milik PPL itu sendiri.
      pembayaranRows: [],
      // Baris dari sheet Status SLS milik PML/PPL ini. Dipakai untuk menyaring
      // baris Lampiran sebelum tabel gabungan dibentuk.
      statusSlsRows: [],
      // Baris sheet Data per SLS untuk tabel beban kerja pada halaman terakhir.
      dataPerSlsRows: [],
      // Baris sheet Approve by PML. Khusus dokumen PML, nilai ini menjadi sumber
      // realisasi/jumlah pemeriksaan, baik per SLS, per PPL, maupun total PML.
      approveByPmlRows: [],
      usernameSobat: cleanText(bappRow?.username_sobat || bappRow?.sobat_id || ""),
    };
    records.push(record);
    if (email) byEmail.set(upperText(email), record);
    registerName(record);
    return record;
  };

  // Record utama tetap dibuat berdasarkan role yang dipilih.
  const roleBappRows = (bappRows || []).filter((row) => isBappRowForRole(row, role));
  for (const row of roleBappRows) {
    const displayName = cleanText(row?.nama);
    const email = cleanText(row?.email);
    if (!displayName && !email) continue;
    ensureRecord(displayName, email, row);
  }

  // Data Lampiran masih dipertahankan untuk bagian BAST/lampiran dokumen gabungan.
  for (const row of lampiranRows || []) {
    const displayName = cleanText(isPml ? row?.nama_pml : row?.nama_ppl);
    const email = cleanText(isPml ? row?.email_pengawas : row?.email_pencacah);
    if (!displayName && !email) continue;

    const emailKey = upperText(email);
    const nameKey = upperText(displayName);
    let record = emailKey ? byEmail.get(emailKey) : null;
    if (!record && nameKey && !ambiguousNames.has(nameKey)) record = byName.get(nameKey);
    if (!record) record = ensureRecord(displayName, email, null);

    record.lampiranRows.push(row);
    if (!record.email && email) record.email = email;
    if ((!record.displayName || record.displayName === "Tanpa Nama") && displayName) record.displayName = displayName;
  }

  // Hubungkan metrik progres dari sheet Pembayaran ke record.
  //
  // FIX v5: jangan bergantung hanya pada kolom Email Pengawas/Nama Pengawas di
  // sheet Pembayaran. Pada banyak file, kolom relasi itu kosong atau berisi #N/A,
  // sehingga pembayaranRows menjadi kosong dan tag {prelist_total}, {realisasi},
  // {persentase}, serta {average_persentase} ikut kosong.
  //
  // Sumber relasi yang paling stabil adalah daftar PPL yang SUDAH menempel pada
  // masing-masing PML di sheet Lampiran. Jadi setiap PPL di Lampiran dicari langsung
  // ke sheet Pembayaran lewat Email Pencacah, lalu fallback ke Nama Pencacah.
  const pplPaymentRows = (bappRows || []).filter((row) => isBappRowForRole(row, "PPL"));
  const paymentByPplEmail = new Map();
  const paymentByPplName = new Map();
  const ambiguousPaymentNames = new Set();

  for (const row of pplPaymentRows) {
    const pplEmail = upperText(row?.email_ppl || row?.email || "");
    const pplName = upperText(row?.nama_ppl || row?.nama || "");
    if (pplEmail && !paymentByPplEmail.has(pplEmail)) paymentByPplEmail.set(pplEmail, row);
    if (pplName) {
      if (paymentByPplName.has(pplName) && paymentByPplName.get(pplName) !== row) {
        ambiguousPaymentNames.add(pplName);
        paymentByPplName.delete(pplName);
      } else if (!ambiguousPaymentNames.has(pplName)) {
        paymentByPplName.set(pplName, row);
      }
    }
  }

  const addPaymentRow = (record, row) => {
    if (!record || !row) return;
    const nama = cleanText(row?.nama_ppl || row?.nama);
    const email = cleanText(row?.email_ppl || row?.email);
    if (!nama && !email) return;
    const paymentIdentity = getBerkasIdentity(nama, email);
    const alreadyAdded = record.pembayaranRows.some((item) =>
      getBerkasIdentity(
        cleanText(item?.nama_ppl || item?.nama),
        cleanText(item?.email_ppl || item?.email)
      ) === paymentIdentity
    );
    if (!alreadyAdded) record.pembayaranRows.push(row);
  };

  // Jalur utama: Lampiran -> identitas PPL -> baris Pembayaran.
  for (const record of records) {
    if (!isPml) {
      addPaymentRow(record, record?.bappRow);
      continue;
    }

    for (const lampiranRow of record.lampiranRows || []) {
      const pplEmail = upperText(lampiranRow?.email_pencacah || lampiranRow?.email_ppl || "");
      const pplName = upperText(lampiranRow?.nama_ppl || "");
      let paymentRow = pplEmail ? paymentByPplEmail.get(pplEmail) : null;
      if (!paymentRow && pplName && !ambiguousPaymentNames.has(pplName)) {
        paymentRow = paymentByPplName.get(pplName) || null;
      }
      addPaymentRow(record, paymentRow);
    }
  }

  // Jalur cadangan: tetap manfaatkan kolom relasi Pengawas bila tersedia.
  // Ini berguna untuk PPL yang ada di Pembayaran tetapi belum tercantum di Lampiran.
  for (const row of pplPaymentRows) {
    const ownerName = cleanText(isPml ? (row?.nama_pengawas || row?.nama_pml) : row?.nama);
    const ownerEmail = cleanText(isPml ? (row?.email_pengawas || row?.email_pml) : row?.email);
    const emailKey = upperText(ownerEmail);
    const nameKey = upperText(ownerName);

    let record = emailKey ? byEmail.get(emailKey) : null;
    if (!record && nameKey && !ambiguousNames.has(nameKey)) record = byName.get(nameKey);
    if (!record) continue;
    addPaymentRow(record, row);
  }

  // Kaitkan sheet Status SLS ke pemilik dokumen. Untuk role PML, pemiliknya
  // ditentukan dari Nama/Email PML; untuk role PPL dari Nama/Email PPL.
  for (const statusRow of statusSlsRows || []) {
    const ownerName = cleanText(isPml ? statusRow?.nama_pml : statusRow?.nama_ppl);
    const ownerEmail = cleanText(isPml ? statusRow?.email_pml : statusRow?.email_ppl);
    const emailKey = upperText(ownerEmail);
    const nameKey = upperText(ownerName);

    let record = emailKey ? byEmail.get(emailKey) : null;
    if (!record && nameKey && !ambiguousNames.has(nameKey)) record = byName.get(nameKey);

    // Fallback paling kuat untuk PML: cari record yang Lampirannya memuat PPL
    // yang sama dengan baris Status SLS. Ini menolong ketika Nama/Email PML pada
    // sheet Status SLS kosong atau penulisannya berbeda.
    if (!record) {
      const statusPplEmail = upperText(statusRow?.email_ppl || "");
      const statusPplName = upperText(statusRow?.nama_ppl || "");
      record = records.find((candidate) =>
        (candidate?.lampiranRows || []).some((lampiranRow) => {
          const lampiranPplEmail = upperText(lampiranRow?.email_pencacah || lampiranRow?.email_ppl || "");
          const lampiranPplName = upperText(lampiranRow?.nama_ppl || "");
          return (statusPplEmail && lampiranPplEmail === statusPplEmail) ||
            (statusPplName && lampiranPplName === statusPplName);
        })
      ) || null;
    }

    if (!record) continue;
    record.statusSlsRows.push(statusRow);
  }

  // Kaitkan sheet Data per SLS ke dokumen PML/PPL. Nama menjadi kunci utama.
  // Username Sobat dipakai untuk membantu pencocokan bila nama tidak unik.
  const normalizePersonKey = (value) => upperText(value).replace(/\s+/g, " ");
  const dataPerSlsRowKey = (row) => [
    cleanText(row?.no_sumber),
    normalizePersonKey(row?.nama_pml),
    normalizePersonKey(row?.username_pml),
    normalizePersonKey(row?.nama_ppl),
    normalizePersonKey(row?.username_ppl),
    cleanText(row?.kdkec),
    cleanText(row?.kddesa),
    cleanText(row?.kode_sls),
  ].join("|");

  for (const dataRow of dataPerSlsRows || []) {
    const ownerName = cleanText(isPml ? dataRow?.nama_pml : dataRow?.nama_ppl);
    const ownerUsername = cleanText(isPml ? dataRow?.username_pml : dataRow?.username_ppl);
    const nameKey = normalizePersonKey(ownerName);
    const usernameKey = normalizePersonKey(ownerUsername);

    let record = null;

    if (usernameKey) {
      const usernameMatches = records.filter((candidate) => {
        const candidateUsername = normalizePersonKey(
          candidate?.usernameSobat ||
          candidate?.bappRow?.username_sobat ||
          candidate?.bappRow?.sobat_id ||
          ""
        );
        const candidateEmail = normalizePersonKey(candidate?.email || candidate?.bappRow?.email || "");
        const candidateEmailLocal = candidateEmail.includes("@") ? candidateEmail.split("@")[0] : candidateEmail;
        return candidateUsername === usernameKey ||
          candidateEmail === usernameKey ||
          candidateEmailLocal === usernameKey;
      });
      if (usernameMatches.length === 1) record = usernameMatches[0];
    }

    if (!record && nameKey) {
      const nameMatches = records.filter(
        (candidate) => normalizePersonKey(candidate?.displayName || "") === nameKey
      );
      if (nameMatches.length === 1) record = nameMatches[0];
    }

    if (!record && nameKey && !ambiguousNames.has(nameKey)) {
      record = byName.get(nameKey) || null;
    }

    // Data per SLS tetap dapat membentuk pilihan dokumen ketika orang tersebut
    // belum muncul pada sheet Pembayaran atau Lampiran.
    if (!record && ownerName) record = ensureRecord(ownerName, "", null);
    if (!record) continue;

    if (!record.usernameSobat && ownerUsername) record.usernameSobat = ownerUsername;
    const rowKey = dataPerSlsRowKey(dataRow);
    const alreadyAdded = record.dataPerSlsRows.some((item) => dataPerSlsRowKey(item) === rowKey);
    if (!alreadyAdded) record.dataPerSlsRows.push(dataRow);
  }

  // Kaitkan sheet Approve by PML ke pemilik dokumen. Untuk role PML, relasi
  // utama memakai Email PML lalu Nama PML. Untuk role PPL data ini tidak mengubah
  // hasil, tetapi tetap dapat ditempel untuk kebutuhan diagnostik.
  const approveRowKey = (row) => [
    upperText(row?.email_pml || ""),
    upperText(row?.nama_pml || ""),
    upperText(row?.email_ppl || ""),
    upperText(row?.nama_ppl || ""),
    normalizeStatusSlsCode(row?.kode_sls || "", 6),
    cleanText(row?.jumlah_approve_pml || ""),
    cleanText(row?.waktu_submit || ""),
  ].join("|");

  for (const approveRow of approveByPmlRows || []) {
    const ownerName = cleanText(isPml ? approveRow?.nama_pml : approveRow?.nama_ppl);
    const ownerEmail = cleanText(isPml ? approveRow?.email_pml : approveRow?.email_ppl);
    const emailKey = upperText(ownerEmail);
    const nameKey = upperText(ownerName);

    let record = emailKey ? byEmail.get(emailKey) : null;
    if (!record && nameKey && !ambiguousNames.has(nameKey)) record = byName.get(nameKey);

    if (!record && isPml) {
      const approvePplEmail = upperText(approveRow?.email_ppl || "");
      const approvePplName = upperText(approveRow?.nama_ppl || "");
      record = records.find((candidate) =>
        (candidate?.lampiranRows || []).some((lampiranRow) => {
          const lampiranPplEmail = upperText(lampiranRow?.email_pencacah || lampiranRow?.email_ppl || "");
          const lampiranPplName = upperText(lampiranRow?.nama_ppl || "");
          return (approvePplEmail && approvePplEmail === lampiranPplEmail) ||
            (approvePplName && approvePplName === lampiranPplName);
        })
      ) || null;
    }

    if (!record && ownerName) record = ensureRecord(ownerName, ownerEmail, null);
    if (!record) continue;

    const rowKey = approveRowKey(approveRow);
    const alreadyAdded = record.approveByPmlRows.some((item) => approveRowKey(item) === rowKey);
    if (!alreadyAdded) record.approveByPmlRows.push(approveRow);
  }

  return records.sort((a, b) =>
    cleanText(a.displayName).localeCompare(cleanText(b.displayName), "id-ID", { sensitivity: "base" })
  );
}

export function pembayaranPersonIdentity(row = {}) {
  return getBerkasIdentity(
    cleanText(row?.nama_ppl || row?.nama),
    cleanText(row?.email_ppl || row?.email)
  );
}

export function findPembayaranRowForLampiran(groupedLampiranRow = {}, pembayaranRows = []) {
  const emailPpl = upperText(groupedLampiranRow?.email_pencacah || groupedLampiranRow?.email_ppl || "");
  const namaPpl = upperText(groupedLampiranRow?.nama_ppl || "");

  if (emailPpl) {
    const byEmail = (pembayaranRows || []).find((row) =>
      upperText(row?.email_ppl || row?.email || "") === emailPpl
    );
    if (byEmail) return byEmail;
  }

  if (namaPpl) {
    return (pembayaranRows || []).find((row) =>
      upperText(row?.nama_ppl || row?.nama || "") === namaPpl
    ) || null;
  }

  return null;
}

export function parsePercentageNumber(value) {
  let text = cleanText(value).replace(/%/g, "").replace(/\s+/g, "");
  if (!text || /^#(?:N\/A|VALUE!|REF!|DIV\/0!|NAME\?|NUM!|NULL!)$/i.test(text)) return null;

  // Toleran terhadap format Indonesia (47,69) maupun format spreadsheet (47.69).
  if (text.includes(",") && text.includes(".")) {
    if (text.lastIndexOf(",") > text.lastIndexOf(".")) {
      text = text.replace(/\./g, "").replace(",", ".");
    } else {
      text = text.replace(/,/g, "");
    }
  } else if (text.includes(",")) {
    text = text.replace(",", ".");
  }

  const parsed = Number(text.replace(/[^0-9.+-]/g, ""));
  return Number.isFinite(parsed) ? parsed : null;
}

export function formatPercentageNumber(value, maximumFractionDigits = 2) {
  if (!Number.isFinite(value)) return "";
  return value.toLocaleString("id-ID", {
    minimumFractionDigits: maximumFractionDigits,
    maximumFractionDigits,
  });
}

export function calculateAveragePersentase(pembayaranRows = []) {
  const uniqueRows = new Map();
  for (const row of pembayaranRows || []) {
    const key = pembayaranPersonIdentity(row);
    if (!uniqueRows.has(key)) uniqueRows.set(key, row);
  }

  const values = [...uniqueRows.values()]
    .map((row) => parsePercentageNumber(row?.persentase_pendataan || row?.persentase_prelist))
    .filter((value) => value != null);

  if (values.length === 0) return { raw: null, formatted: "", count: 0 };
  const raw = values.reduce((sum, value) => sum + value, 0) / values.length;
  return { raw, formatted: formatPercentageNumber(raw, 2), count: values.length };
}

export function parseDataPerSlsNumber(value) {
  let text = cleanText(value).replace(/%/g, "").replace(/\s+/g, "");
  if (!text || /^#(?:N\/A|VALUE!|REF!|DIV\/0!|NAME\?|NUM!|NULL!)$/i.test(text)) return null;

  if (text.includes(",") && text.includes(".")) {
    if (text.lastIndexOf(",") > text.lastIndexOf(".")) {
      text = text.replace(/\./g, "").replace(",", ".");
    } else {
      text = text.replace(/,/g, "");
    }
  } else if (text.includes(",")) {
    text = text.replace(",", ".");
  } else if (/^[+-]?\d{1,3}(?:\.\d{3})+$/.test(text)) {
    // Untuk angka hitungan, titik berulang tiga digit dianggap pemisah ribuan.
    text = text.replace(/\./g, "");
  }

  const parsed = Number(text.replace(/[^0-9.+-]/g, ""));
  return Number.isFinite(parsed) ? parsed : null;
}

export function formatDataPerSlsAggregate(value) {
  if (!Number.isFinite(value)) return "";
  if (Math.abs(value - Math.round(value)) < 1e-9) return String(Math.round(value));
  return value.toLocaleString("id-ID", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

export function normalizeApprovePersonKey(value) {
  return upperText(value)
    .replace(/^@/, "")
    .replace(/\s+/g, " ");
}

export function approveEmailLocalPart(value) {
  const text = normalizeApprovePersonKey(value);
  return text.includes("@") ? text.split("@")[0] : text;
}

export function sumJumlahApprovePml(rows = []) {
  let total = 0;
  let count = 0;
  for (const row of rows || []) {
    const value = parseDataPerSlsNumber(row?.jumlah_approve_pml);
    if (value == null) continue;
    total += value;
    count += 1;
  }
  return {
    raw: count > 0 ? total : null,
    formatted: count > 0 ? formatDataPerSlsAggregate(total) : "",
    count,
  };
}

export function approveRowMatchesPpl(approveRow = {}, target = {}) {
  const approveEmail = normalizeApprovePersonKey(approveRow?.email_ppl || "");
  const approveEmailLocal = approveEmailLocalPart(approveRow?.email_ppl || "");
  const approveName = normalizeApprovePersonKey(approveRow?.nama_ppl || "");

  const targetEmail = normalizeApprovePersonKey(
    target?.email_ppl || target?.email_pencacah || target?.email || ""
  );
  const targetEmailLocal = approveEmailLocalPart(
    target?.email_ppl || target?.email_pencacah || target?.email || ""
  );
  const targetUsername = normalizeApprovePersonKey(
    target?.username_ppl || target?.username_sobat_ppl || ""
  );
  const targetName = normalizeApprovePersonKey(target?.nama_ppl || target?.nama || "");

  return Boolean(
    (approveEmail && targetEmail && approveEmail === targetEmail) ||
    (approveEmailLocal && targetEmailLocal && approveEmailLocal === targetEmailLocal) ||
    (approveEmailLocal && targetUsername && approveEmailLocal === targetUsername) ||
    (approveName && targetName && approveName === targetName)
  );
}

export function findApproveRowsForWorkloadRow(workloadRow = {}, approveRows = []) {
  const kodeSls = normalizeStatusSlsCode(workloadRow?.kode_sls || "", 6);
  const byCode = (approveRows || []).filter((row) =>
    normalizeStatusSlsCode(row?.kode_sls || "", 6) === kodeSls
  );
  if (byCode.length === 0) return [];

  const byPpl = byCode.filter((row) => approveRowMatchesPpl(row, workloadRow));
  if (byPpl.length > 0) return byPpl;

  const workloadPmlName = normalizeApprovePersonKey(workloadRow?.nama_pml || "");
  const workloadPmlUsername = normalizeApprovePersonKey(
    workloadRow?.username_pml || workloadRow?.username_sobat_pml || ""
  );
  const byPml = byCode.filter((row) => {
    const approvePmlName = normalizeApprovePersonKey(row?.nama_pml || "");
    const approvePmlEmailLocal = approveEmailLocalPart(row?.email_pml || "");
    return (workloadPmlName && approvePmlName === workloadPmlName) ||
      (workloadPmlUsername && approvePmlEmailLocal === workloadPmlUsername);
  });
  if (byPml.length > 0) return byPml;

  return byCode.length === 1 ? byCode : [];
}

export function applyApproveByPmlToWorkload(workload = {}, approveRows = [], role = "PML") {
  if (upperText(role) !== "PML" || !Array.isArray(approveRows) || approveRows.length === 0) {
    return workload;
  }

  const rows = (workload?.rows || []).map((row) => {
    const matchingApproveRows = findApproveRowsForWorkloadRow(row, approveRows);
    const approved = sumJumlahApprovePml(matchingApproveRows);
    const approvedRaw = approved.raw == null ? 0 : approved.raw;
    const approvedFormatted = approved.raw == null ? "0" : approved.formatted;

    const target = parseDataPerSlsNumber(row?.target_jumlah);
    const percentageRaw = target && target > 0 ? (approvedRaw / target) * 100 : null;
    const percentage = percentageRaw == null ? "" : formatPercentageNumber(percentageRaw, 2);

    return {
      ...row,
      realisasi_jumlah: approvedFormatted,
      realisasi_dengan_tidak_ditemukan_jumlah: approvedFormatted,
      jumlah_approve_pml: approvedFormatted,
      jumlah_approve_pml_raw: approvedRaw,
      jumlah_baris_approve_pml: approved.count,
      approve_pml_ditemukan: approved.raw != null,
      persentase: percentage,
      persentase_dengan_tidak_ditemukan: percentage,
      persentase_raw: percentageRaw == null ? "" : percentageRaw,
      sumber_realisasi: "Approve by PML",
    };
  });

  const targetKeluargaRaw = rows.reduce((sum, row) => sum + (parseDataPerSlsNumber(row?.target_keluarga) || 0), 0);
  const targetUsahaRaw = rows.reduce((sum, row) => sum + (parseDataPerSlsNumber(row?.target_usaha) || 0), 0);
  const targetJumlahRaw = rows.reduce((sum, row) => sum + (parseDataPerSlsNumber(row?.target_jumlah) || 0), 0);
  const realisasiKeluargaRaw = rows.reduce((sum, row) => sum + (parseDataPerSlsNumber(row?.realisasi_keluarga) || 0), 0);
  const realisasiUsahaRaw = rows.reduce((sum, row) => sum + (parseDataPerSlsNumber(row?.realisasi_usaha) || 0), 0);
  const realisasiJumlahDariBarisRaw = rows.reduce(
    (sum, row) => sum + (parseDataPerSlsNumber(row?.realisasi_jumlah) || 0),
    0
  );
  // Total PML harus sama persis dengan penjumlahan seluruh kolom
  // "Jumlah Approve PML" pada record PML, walaupun ada kode SLS yang tidak cocok.
  const directApproveTotal = sumJumlahApprovePml(approveRows);
  const realisasiJumlahRaw = directApproveTotal.raw == null
    ? realisasiJumlahDariBarisRaw
    : directApproveTotal.raw;
  const percentageRaw = targetJumlahRaw > 0 ? (realisasiJumlahRaw / targetJumlahRaw) * 100 : null;

  const total = {
    ...(workload?.total || {}),
    target_keluarga: rows.length ? formatDataPerSlsAggregate(targetKeluargaRaw) : "",
    target_usaha: rows.length ? formatDataPerSlsAggregate(targetUsahaRaw) : "",
    target_jumlah: rows.length ? formatDataPerSlsAggregate(targetJumlahRaw) : "",
    realisasi_keluarga: rows.length ? formatDataPerSlsAggregate(realisasiKeluargaRaw) : "",
    realisasi_usaha: rows.length ? formatDataPerSlsAggregate(realisasiUsahaRaw) : "",
    realisasi_jumlah: rows.length ? formatDataPerSlsAggregate(realisasiJumlahRaw) : "",
    jumlah_approve_pml: rows.length ? formatDataPerSlsAggregate(realisasiJumlahRaw) : "",
    persentase: percentageRaw == null ? "" : formatPercentageNumber(percentageRaw, 2),
    persentase_raw: percentageRaw == null ? "" : percentageRaw,
    // Keterangan PML sengaja dikosongkan agar tidak ditulis pada pemberkasan.
    keterangan: "",
    sumber_realisasi: "Approve by PML",
  };

  return { rows, total };
}

export function getDataPerSlsValueOrSum(primaryValue, firstValue, secondValue) {
  const primary = cleanText(primaryValue);
  if (primary) return primary;
  const first = parseDataPerSlsNumber(firstValue);
  const second = parseDataPerSlsNumber(secondValue);
  if (first == null && second == null) return "";
  return formatDataPerSlsAggregate((first || 0) + (second || 0));
}

export function buildDataPerSlsWorkloadRows(sourceRows = [], role = "PML") {
  const sortedSource = [...(sourceRows || [])].sort((a, b) => {
    const noA = parseDataPerSlsNumber(a?.no_sumber);
    const noB = parseDataPerSlsNumber(b?.no_sumber);
    if (noA != null && noB != null && noA !== noB) return noA - noB;

    const kecDiff = cleanText(a?.kdkec).localeCompare(cleanText(b?.kdkec), "id-ID", { numeric: true });
    if (kecDiff !== 0) return kecDiff;
    const desaDiff = cleanText(a?.kddesa).localeCompare(cleanText(b?.kddesa), "id-ID", { numeric: true });
    if (desaDiff !== 0) return desaDiff;
    return cleanText(a?.kode_sls).localeCompare(cleanText(b?.kode_sls), "id-ID", { numeric: true });
  });

  const rows = sortedSource.map((row, index) => {
    const targetJumlah = getDataPerSlsValueOrSum(
      row?.target_jumlah,
      row?.target_keluarga,
      row?.target_usaha
    );
    const realisasiJumlah = getDataPerSlsValueOrSum(
      row?.realisasi_dengan_tidak_ditemukan_jumlah,
      row?.realisasi_dengan_tidak_ditemukan_keluarga,
      row?.realisasi_dengan_tidak_ditemukan_usaha
    );

    const targetNumber = parseDataPerSlsNumber(targetJumlah);
    const realisasiNumber = parseDataPerSlsNumber(realisasiJumlah);
    const percentageSource = cleanText(row?.persentase_dengan_tidak_ditemukan || "");
    const percentageNumber = parsePercentageNumber(percentageSource);
    const computedPercentage = percentageNumber != null
      ? percentageNumber
      : targetNumber && realisasiNumber != null
        ? (realisasiNumber / targetNumber) * 100
        : null;
    const percentageFormatted = computedPercentage == null
      ? percentageSource
      : formatPercentageNumber(computedPercentage, 2);

    return {
      no: index + 1,
      no_sumber: cleanText(row?.no_sumber || ""),
      nama_pml: cleanText(row?.nama_pml || ""),
      username_pml: cleanText(row?.username_pml || ""),
      username_sobat_pml: cleanText(row?.username_pml || ""),
      nama_ppl: cleanText(row?.nama_ppl || ""),
      username_ppl: cleanText(row?.username_ppl || ""),
      username_sobat_ppl: cleanText(row?.username_ppl || ""),

      kdkec: normalizeStatusSlsCode(row?.kdkec || "", 3),
      kode_kecamatan: normalizeStatusSlsCode(row?.kdkec || "", 3),
      kddesa: normalizeStatusSlsCode(row?.kddesa || "", 3),
      kode_kelurahan: normalizeStatusSlsCode(row?.kddesa || "", 3),
      kode_sls: normalizeStatusSlsCode(row?.kode_sls || "", 6),

      target_keluarga: cleanText(row?.target_keluarga || ""),
      target_usaha: cleanText(row?.target_usaha || ""),
      target_jumlah: targetJumlah,
      target_prelist_keluarga: cleanText(row?.target_keluarga || ""),
      target_prelist_usaha: cleanText(row?.target_usaha || ""),
      target_prelist_jumlah: targetJumlah,

      realisasi_keluarga: cleanText(row?.realisasi_dengan_tidak_ditemukan_keluarga || ""),
      realisasi_usaha: cleanText(row?.realisasi_dengan_tidak_ditemukan_usaha || ""),
      realisasi_jumlah: realisasiJumlah,
      realisasi_dengan_tidak_ditemukan_keluarga: cleanText(row?.realisasi_dengan_tidak_ditemukan_keluarga || ""),
      realisasi_dengan_tidak_ditemukan_usaha: cleanText(row?.realisasi_dengan_tidak_ditemukan_usaha || ""),
      realisasi_dengan_tidak_ditemukan_jumlah: realisasiJumlah,

      persentase: percentageFormatted,
      persentase_dengan_tidak_ditemukan: percentageFormatted,
      persentase_raw: percentageSource,

      realisasi_tanpa_tidak_ditemukan_keluarga: cleanText(row?.realisasi_tanpa_tidak_ditemukan_keluarga || ""),
      realisasi_tanpa_tidak_ditemukan_usaha: cleanText(row?.realisasi_tanpa_tidak_ditemukan_usaha || ""),
      realisasi_tanpa_tidak_ditemukan_jumlah: cleanText(row?.realisasi_tanpa_tidak_ditemukan_jumlah || ""),
      persentase_tanpa_tidak_ditemukan: cleanText(row?.persentase_tanpa_tidak_ditemukan || ""),

      keterangan: cleanText(row?.keterangan || ""),
      status: cleanText(row?.status || ""),
      role,
    };
  });

  const sum = (key) => rows.reduce((total, row) => {
    const value = parseDataPerSlsNumber(row?.[key]);
    return total + (value == null ? 0 : value);
  }, 0);

  const targetKeluargaRaw = sum("target_keluarga");
  const targetUsahaRaw = sum("target_usaha");
  const targetJumlahRaw = sum("target_jumlah");
  const realisasiKeluargaRaw = sum("realisasi_keluarga");
  const realisasiUsahaRaw = sum("realisasi_usaha");
  const realisasiJumlahRaw = sum("realisasi_jumlah");
  const percentageRaw = targetJumlahRaw > 0
    ? (realisasiJumlahRaw / targetJumlahRaw) * 100
    : null;

  const total = {
    target_keluarga: rows.length ? formatDataPerSlsAggregate(targetKeluargaRaw) : "",
    target_usaha: rows.length ? formatDataPerSlsAggregate(targetUsahaRaw) : "",
    target_jumlah: rows.length ? formatDataPerSlsAggregate(targetJumlahRaw) : "",
    realisasi_keluarga: rows.length ? formatDataPerSlsAggregate(realisasiKeluargaRaw) : "",
    realisasi_usaha: rows.length ? formatDataPerSlsAggregate(realisasiUsahaRaw) : "",
    realisasi_jumlah: rows.length ? formatDataPerSlsAggregate(realisasiJumlahRaw) : "",
    persentase: percentageRaw == null ? "" : formatPercentageNumber(percentageRaw, 2),
    persentase_raw: percentageRaw == null ? "" : percentageRaw,
    keterangan: percentageRaw == null
      ? ""
      : percentageRaw >= 40
        ? "Bisa Dibayar karena lebih dari 40%"
        : "Belum Bisa Dibayar karena kurang dari 40%",
  };

  return { rows, total };
}

export function buildLampiranStatusSlsCode(row = {}) {
  // Sheet Lampiran menyimpan kode SLS sebagai kdsls (4 digit) + kdsubsls (2 digit).
  // Jangan memakai kdsubslspanjang karena kolom itu dapat memuat kode wilayah panjang
  // dan enam digit terakhirnya tidak selalu identik dengan Kode SLS pada sheet Status SLS.
  const direct = normalizeStatusSlsCode(row?.kode_sls || "", 6);
  if (direct) return direct;

  const kdsls = normalizeStatusSlsCode(row?.kdsls || "", 4);
  const kdsubsls = normalizeStatusSlsCode(row?.kdsubsls || "", 2);
  return kdsls ? `${kdsls}${kdsubsls || "00"}` : "";
}

export function statusSlsRowsHaveUsableFilter(statusSlsRows = []) {
  return (statusSlsRows || []).some((row) =>
    cleanText(row?.kode_sls) || cleanText(row?.status) || cleanText(row?.status_raw)
  );
}

export function findStatusSlsRowForLampiran(lampiranRow = {}, statusSlsRows = []) {
  const kodeKec = normalizeStatusSlsCode(lampiranRow?.kdkec || "", 3);
  const kodeDesa = normalizeStatusSlsCode(lampiranRow?.kddesa || "", 3);
  const kodeSls = buildLampiranStatusSlsCode(lampiranRow);
  const emailPpl = upperText(lampiranRow?.email_pencacah || lampiranRow?.email_ppl || "");
  const namaPpl = upperText(lampiranRow?.nama_ppl || "");
  const emailPml = upperText(lampiranRow?.email_pengawas || lampiranRow?.email_pml || "");
  const namaPml = upperText(lampiranRow?.nama_pml || "");

  const sameCodes = (row) => {
    const rowKec = normalizeStatusSlsCode(row?.kdkec || "", 3);
    const rowDesa = normalizeStatusSlsCode(row?.kddesa || "", 3);
    const rowSls = normalizeStatusSlsCode(row?.kode_sls || "", 6);
    return (!kodeKec || rowKec === kodeKec) &&
      (!kodeDesa || rowDesa === kodeDesa) &&
      (!kodeSls || rowSls === kodeSls);
  };

  const candidates = (statusSlsRows || []).filter(sameCodes);
  if (candidates.length === 0) return null;

  if (emailPpl) {
    const match = candidates.find((row) => upperText(row?.email_ppl || "") === emailPpl);
    if (match) return match;
  }
  if (namaPpl) {
    const match = candidates.find((row) => upperText(row?.nama_ppl || "") === namaPpl);
    if (match) return match;
  }
  if (emailPml) {
    const match = candidates.find((row) => upperText(row?.email_pml || "") === emailPml);
    if (match) return match;
  }
  if (namaPml) {
    const match = candidates.find((row) => upperText(row?.nama_pml || "") === namaPml);
    if (match) return match;
  }

  // Fallback berdasarkan kode hanya boleh dipakai bila hasilnya tunggal, agar SLS
  // dengan kode sama milik petugas berbeda tidak salah ditempelkan.
  return candidates.length === 1 ? candidates[0] : null;
}

export function filterLampiranRowsByStatusSls(lampiranRows = [], statusSlsRows = []) {
  if (!statusSlsRowsHaveUsableFilter(statusSlsRows)) return [...(lampiranRows || [])];

  return (lampiranRows || [])
    .map((row) => {
      const statusRow = findStatusSlsRowForLampiran(row, statusSlsRows);
      if (!statusRow || !isAllowedStatusSls(statusRow?.status || statusRow?.status_raw)) return null;
      return {
        ...row,
        status_sls: normalizeStatusSlsLabel(statusRow?.status || statusRow?.status_raw),
        status_sls_raw: cleanText(statusRow?.status_raw || statusRow?.status || ""),
        status_jumlah_prelist: cleanText(statusRow?.jumlah_prelist || ""),
        status_jumlah_realisasi: cleanText(statusRow?.jumlah_realisasi || ""),
        status_persentase: cleanText(statusRow?.persentase || ""),
      };
    })
    .filter(Boolean);
}

export function buildBerkasLampiranTableRows(lampiranRows = [], role = "PML", pembayaranRows = [], statusSlsRows = []) {
  if (!Array.isArray(lampiranRows) || lampiranRows.length === 0) return [];

  // Bila sheet Status SLS tersedia, hanya baris dengan status Selesai atau
  // Sedang Dikerjakan yang masuk ke pengelompokan dan perhitungan {jumlah}.
  const filteredLampiranRows = filterLampiranRowsByStatusSls(lampiranRows, statusSlsRows);
  const grouped = groupLampiranRows(filteredLampiranRows, role);

  return grouped.map((row, index) => {
    const jumlah = Number(row?.jumlah || 0);
    const jumlah40 = Math.ceil(jumlah * 0.4);
    const jumlah60 = jumlah - jumlah40;
    const namaPml = cleanText(row?.nama_pml || "");
    const namaPpl = cleanText(row?.nama_ppl || "");

    // Setiap baris tabel PML mewakili PPL, sehingga nomor generik pada baris
    // diarahkan ke kontrak PPL. Nomor kontrak PML tetap tersedia lewat tag khusus.
    const nomorKontrakPml = cleanText(row?.nomor_kontrak_pml || "");
    const nomorKontrakPpl = cleanText(row?.nomor_kontrak_ppl || "");
    const nomorKontrakBaris = nomorKontrakPpl;

    const pembayaranRow = findPembayaranRowForLampiran(row, pembayaranRows);
    const prelistTotal = cleanText(pembayaranRow?.prelist_total || pembayaranRow?.target_prelist || "");
    const realisasi = cleanText(pembayaranRow?.realisasi_total || pembayaranRow?.realisasi_hasil_pendataan || "");
    const persentaseRaw = cleanText(pembayaranRow?.persentase_pendataan || pembayaranRow?.persentase_prelist || "");
    const persentaseNumber = parsePercentageNumber(persentaseRaw);
    // Semua tag tampilan persentase memakai dua angka di belakang koma.
    // Nilai asli tetap tersedia melalui {persentase_raw}.
    const persentaseFormat = persentaseNumber == null
      ? persentaseRaw
      : formatPercentageNumber(persentaseNumber, 2);
    const persentase = persentaseFormat;

    return {
      no: index + 1,
      nama: namaPpl,
      nama_petugas: namaPpl,
      nama_pml: namaPml,
      nama_pengawas: namaPml,
      nama_ppl: namaPpl,
      email_pengawas: cleanText(row?.email_pengawas || ""),
      email_pml: cleanText(row?.email_pengawas || ""),
      email_pencacah: cleanText(row?.email_pencacah || ""),
      email_ppl: cleanText(row?.email_pencacah || ""),
      email: cleanText(row?.email_pencacah || ""),

      kdprov: cleanText(row?.kdprov || ""),
      kdkab: cleanText(row?.kdkab || ""),
      kdkec: cleanText(row?.kdkec || ""),
      kddesa: cleanText(row?.kddesa || ""),
      nmprov: cleanText(row?.nmprov || ""),
      nmkab: cleanText(row?.nmkab || ""),
      kecamatan: formatKodeNama(row?.kdkec, row?.kecamatan),
      kelurahan: formatKodeNama(row?.kddesa, row?.kelurahan),

      jumlah,
      total_jumlah: jumlah,
      jumlah_40: jumlah40,
      jumlah_60: jumlah60,
      sls_total: jumlah,
      sls_40: jumlah40,
      sls_60: jumlah60,

      nomor_spk: nomorKontrakBaris,
      nomor_kontrak: nomorKontrakBaris,
      nomor_kontrak_pml: nomorKontrakPml,
      nomor_kontrak_ppl: nomorKontrakPpl,

      // Struktur/identitas tabel tetap dari sheet Lampiran. Tiga metrik berikut
      // diambil dari baris PPL yang cocok pada sheet Pembayaran.
      prelist_total: prelistTotal,
      target_prelist: prelistTotal,
      realisasi,
      realisasi_total: realisasi,
      realisasi_hasil_pendataan: realisasi,
      persentase,
      persentase_pendataan: persentase,
      persentase_prelist: persentase,
      persentase_raw: persentaseRaw,
      persentase_format: persentaseFormat,

      // Ringkasan status SLS yang lolos filter untuk baris kelompok ini.
      // Karena satu baris tabel dapat berisi beberapa SLS, status digabung unik.
      status_sls: (() => {
        const statuses = filteredLampiranRows
          .filter((item) =>
            upperText(item?.nama_ppl || "") === upperText(row?.nama_ppl || "") &&
            normalizeStatusSlsCode(item?.kdkec || "", 3) === normalizeStatusSlsCode(row?.kdkec || "", 3) &&
            normalizeStatusSlsCode(item?.kddesa || "", 3) === normalizeStatusSlsCode(row?.kddesa || "", 3)
          )
          .map((item) => normalizeStatusSlsLabel(item?.status_sls || ""))
          .filter(Boolean);
        return [...new Set(statuses)].join(" & ");
      })(),

      // Alias lama tetap disediakan agar template sebelumnya tidak rusak.
      sls_ongoing: "",
      sls_selesai_sedang_dikerjakan: "",
      persentase_sls: "",
      tanggal_screenshot: "",
      flag: "",
    };
  });
}

export function buildPembayaranTableRow(row = {}, index = 0) {
  const jabatan = upperText(row?.jabatan || row?.jabatan_raw || "");
  const namaPpl = cleanText(row?.nama_ppl || (/PPL|PENCACAH/.test(jabatan) ? row?.nama : ""));
  const namaPml = cleanText(row?.nama_pengawas || row?.nama_pml || (/PML|PENGAWAS/.test(jabatan) ? row?.nama : ""));
  const targetPrelist = cleanText(row?.target_prelist || row?.prelist_total || "");
  const realisasiHasil = cleanText(row?.realisasi_hasil_pendataan || row?.realisasi_total || "");
  const slsOngoing = cleanText(row?.sls_ongoing || row?.sls_selesai_sedang_dikerjakan || "");
  const persentasePrelist = cleanText(row?.persentase_prelist || row?.persentase_pendataan || "");
  const nomorSpk = cleanText(row?.nomor_spk || row?.nomor_kontrak || "");

  return {
    no: index + 1,
    nama: namaPpl || cleanText(row?.nama || ""),
    nama_petugas: namaPpl || cleanText(row?.nama || ""),
    nama_ppl: namaPpl,
    email_ppl: cleanText(row?.email_ppl || row?.email || ""),
    email: cleanText(row?.email || row?.email_ppl || ""),
    nik: cleanText(row?.nik || ""),
    jabatan: cleanText(row?.jabatan || row?.jabatan_raw || ""),

    nama_pml: namaPml,
    nama_pengawas: namaPml,
    email_pengawas: cleanText(row?.email_pengawas || row?.email_pml || ""),

    nomor_spk: nomorSpk,
    nomor_kontrak: nomorSpk,
    kecamatan: cleanText(row?.kecamatan || row?.wilayah || ""),
    sls_total: cleanText(row?.sls_total || ""),
    sls_40: cleanText(row?.sls_40 || ""),
    sls_60: cleanText(row?.sls_60 || ""),
    sls_ongoing: slsOngoing,
    sls_selesai_sedang_dikerjakan: slsOngoing,
    persentase_sls: cleanText(row?.persentase_sls || ""),
    tanggal_screenshot: cleanText(row?.tanggal_screenshot || ""),

    target_prelist: targetPrelist,
    prelist_total: targetPrelist,
    realisasi_hasil_pendataan: realisasiHasil,
    realisasi_total: realisasiHasil,
    persentase_prelist: persentasePrelist,
    persentase_pendataan: persentasePrelist,
    flag: cleanText(row?.flag || ""),
  };
}

export function prefixTemplateData(prefix, data = {}) {
  const prefixed = {};
  for (const [key, value] of Object.entries(data || {})) {
    prefixed[`${prefix}_${key}`] = value;
  }
  return prefixed;
}

export function buildBerkasPembayaranTemplateData(formValues, record, role, nikLookup, halamanDepanMap) {
  const isPml = upperText(role) === "PML";
  const lampiranRows = Array.isArray(record?.lampiranRows) ? record.lampiranRows : [];
  const pembayaranRows = Array.isArray(record?.pembayaranRows) ? record.pembayaranRows : [];
  const statusSlsRows = Array.isArray(record?.statusSlsRows) ? record.statusSlsRows : [];
  const dataPerSlsRows = Array.isArray(record?.dataPerSlsRows) ? record.dataPerSlsRows : [];
  const approveByPmlRows = Array.isArray(record?.approveByPmlRows) ? record.approveByPmlRows : [];
  const firstLampiran = lampiranRows[0] || {};
  const displayName = cleanText(
    record?.displayName || record?.bappRow?.nama || (isPml ? firstLampiran?.nama_pml : firstLampiran?.nama_ppl)
  );
  const email = cleanText(
    record?.email || record?.bappRow?.email || (isPml ? firstLampiran?.email_pengawas : firstLampiran?.email_pencacah)
  );
  const nomorKontrakLampiran = cleanText(
    isPml ? firstLampiran?.nomor_kontrak_pml : firstLampiran?.nomor_kontrak_ppl
  );
  // Isi tabel mengikuti sheet Lampiran, memakai pengelompokan yang sama dengan
  // dokumen Lampiran/BAST terpisah. Metrik target dan realisasi kemudian
  // diperkaya dari sheet Data per SLS, khusus kolom "Dengan Tidak Ditemukan".
  const lampiranTableRowsBase = buildBerkasLampiranTableRows(
    lampiranRows,
    role,
    pembayaranRows,
    statusSlsRows
  );
  const bebanKerjaBase = buildDataPerSlsWorkloadRows(dataPerSlsRows, role);
  const bebanKerja = applyApproveByPmlToWorkload(
    bebanKerjaBase,
    approveByPmlRows,
    role
  );

  const normalizeMatchKey = (value) => upperText(value).replace(/^@/, "").replace(/\s+/g, " ");
  const emailLocalPart = (value) => {
    const text = normalizeMatchKey(value);
    return text.includes("@") ? text.split("@")[0] : text;
  };

  // Untuk dokumen PML, setiap baris tabel peserta mewakili satu PPL. Target,
  // realisasi, dan persentase dijumlahkan dari seluruh SLS milik PPL tersebut.
  // Untuk dokumen PPL, seluruh baris Data per SLS pada record adalah milik PPL itu.
  const lampiranTableRows = lampiranTableRowsBase.map((tableRow) => {
    const pplName = normalizeMatchKey(tableRow?.nama_ppl || tableRow?.nama || "");
    const pplEmailLocal = emailLocalPart(tableRow?.email_ppl || tableRow?.email_pencacah || tableRow?.email || "");

    const matchedSourceRows = isPml
      ? dataPerSlsRows.filter((sourceRow) => {
          const sourceName = normalizeMatchKey(sourceRow?.nama_ppl || "");
          const sourceUsername = normalizeMatchKey(sourceRow?.username_ppl || "");
          return (pplName && sourceName === pplName) ||
            (pplEmailLocal && sourceUsername === pplEmailLocal);
        })
      : dataPerSlsRows;

    const matchedApproveRows = isPml
      ? approveByPmlRows.filter((approveRow) => approveRowMatchesPpl(approveRow, tableRow))
      : [];

    if (matchedSourceRows.length === 0 && matchedApproveRows.length === 0) return tableRow;

    const pplWorkloadBase = buildDataPerSlsWorkloadRows(matchedSourceRows, "PPL");
    const pplWorkload = isPml
      ? applyApproveByPmlToWorkload(pplWorkloadBase, matchedApproveRows, "PML")
      : pplWorkloadBase;
    const total = pplWorkload.total || {};
    const approveTotal = sumJumlahApprovePml(matchedApproveRows);
    const targetForPercentage = parseDataPerSlsNumber(total.target_jumlah || tableRow?.jumlah_pre || tableRow?.prelist_total);
    const useApproveSource = isPml && matchedApproveRows.length > 0;
    const approveValueRaw = useApproveSource
      ? (approveTotal.raw == null ? 0 : approveTotal.raw)
      : null;
    const approvePercentageRaw = useApproveSource && targetForPercentage
      ? (approveValueRaw / targetForPercentage) * 100
      : null;
    const approvePercentage = useApproveSource
      ? (approvePercentageRaw == null ? "" : formatPercentageNumber(approvePercentageRaw, 2))
      : total.persentase;
    const rowRealisasi = useApproveSource
      ? formatDataPerSlsAggregate(approveValueRaw)
      : total.realisasi_jumlah;

    return {
      ...tableRow,
      pre_keluarga: total.target_keluarga,
      pre_usaha: total.target_usaha,
      jumlah_pre: total.target_jumlah,
      target_keluarga: total.target_keluarga,
      target_usaha: total.target_usaha,
      target_jumlah: total.target_jumlah,
      prelist_total: total.target_jumlah,
      target_prelist: total.target_jumlah,

      realisasi_keluarga: total.realisasi_keluarga,
      realisasi_usaha: total.realisasi_usaha,
      jumlah_realisasi: rowRealisasi,
      realisasi: rowRealisasi,
      realisasi_total: rowRealisasi,
      realisasi_hasil_pendataan: rowRealisasi,
      jumlah_approve_pml: useApproveSource ? formatDataPerSlsAggregate(approveValueRaw) : "",

      persentase: approvePercentage,
      pesentase: approvePercentage,
      persentase_pendataan: approvePercentage,
      persentase_prelist: approvePercentage,
      persentase_raw: approvePercentageRaw == null ? total.persentase_raw : approvePercentageRaw,
      persentase_format: approvePercentage,
      // Pemberkasan PML dan PPL tidak menampilkan keterangan/status pembayaran.
      status: "",
      keterangan: "",
    };
  });

  // Hitung rata-rata dari baris yang benar-benar tampil pada tabel. Bila Data per
  // SLS tersedia, nilai yang dirata-ratakan sudah memakai realisasi "Dengan Tidak
  // Ditemukan", bukan lagi nilai lama dari sheet Pembayaran.
  const averagePersentase = calculateAveragePersentase(lampiranTableRows);

  const bappRow = record?.bappRow || {
    nama: displayName,
    email,
    jabatan: role,
    jabatan_raw: role,
    nik: nikLookup?.get(upperText(displayName)) || "",
    nomor_spk: nomorKontrakLampiran,
    nomor_kontrak: nomorKontrakLampiran,
  };

  const syntheticLampiranRow = {
    ...firstLampiran,
    nama_pml: isPml ? displayName : cleanText(firstLampiran?.nama_pml),
    nama_ppl: isPml ? cleanText(firstLampiran?.nama_ppl) : displayName,
    email_pengawas: isPml ? email : cleanText(firstLampiran?.email_pengawas),
    email_pencacah: isPml ? cleanText(firstLampiran?.email_pencacah) : email,
    nomor_kontrak_pml: isPml ? cleanText(bappRow?.nomor_spk || bappRow?.nomor_kontrak || nomorKontrakLampiran) : cleanText(firstLampiran?.nomor_kontrak_pml),
    nomor_kontrak_ppl: isPml ? cleanText(firstLampiran?.nomor_kontrak_ppl) : cleanText(bappRow?.nomor_spk || bappRow?.nomor_kontrak || nomorKontrakLampiran),
  };

  const rowsForBast = lampiranRows.length > 0 ? lampiranRows : [syntheticLampiranRow];
  console.log("Record approveByPmlRows:", record?.approveByPmlRows?.length, record?.approveByPmlRows);
  console.log("Record displayName/email:", record?.displayName, record?.email);
  
  const bappRaw = buildBappTemplateData(formValues || {}, bappRow, role, approveByPmlRows);
  const bastRaw = buildBastTemplateData(formValues || {}, rowsForBast, role, nikLookup);
  const nik = cleanText(bappRaw.nik || bastRaw.nik || nikLookup?.get(upperText(displayName)) || "");
  const nomorKontrak = cleanText(bappRaw.nomor_kontrak || bastRaw.nomor_perjanjian || nomorKontrakLampiran);
  const bapp = {
    ...bappRaw,
    nama: displayName,
    nama_peserta: displayName,
    nama_petugas: displayName,
    email,
    nik,
    nomor_kontrak: nomorKontrak,
    nomor_prefix: extractNomorPrefix(nomorKontrak),
  };
  const bast = {
    ...bastRaw,
    nama: displayName,
    nik,
    nomor_perjanjian: bastRaw.nomor_perjanjian || nomorKontrak,
  };
  const suratPernyataan = {
    nomor_prefix: extractNomorPrefix(nomorKontrak),
    nama_petugas: displayName,
    nama: displayName,
    nik,
    nomor_spk: nomorKontrak,
    nomor_kontrak: nomorKontrak,
  };
  const pembayaranRingkasan = buildPembayaranTableRow(bappRow, 0);
  const firstLampiranTableRow = lampiranTableRows[0] || {};
  const hasDataPerSls = bebanKerja.rows.length > 0;
  const dataPerSlsSummary = bebanKerja.total || {};
  const summaryTargetKeluarga = hasDataPerSls ? dataPerSlsSummary.target_keluarga : "";
  const summaryTargetUsaha = hasDataPerSls ? dataPerSlsSummary.target_usaha : "";
  const summaryTargetJumlah = hasDataPerSls
    ? dataPerSlsSummary.target_jumlah
    : pembayaranRingkasan.prelist_total;
  const summaryRealisasiKeluarga = hasDataPerSls ? dataPerSlsSummary.realisasi_keluarga : "";
  const summaryRealisasiUsaha = hasDataPerSls ? dataPerSlsSummary.realisasi_usaha : "";
  const summaryRealisasiJumlah = hasDataPerSls
    ? dataPerSlsSummary.realisasi_jumlah
    : pembayaranRingkasan.realisasi_total;
  const summaryPersentase = hasDataPerSls
    ? dataPerSlsSummary.persentase
    : pembayaranRingkasan.persentase_pendataan;
  const fotoBukti = collectFotoBuktiFromApproveRows(approveByPmlRows, role);
  const fotoRows = chunkFotoBuktiIntoRows(fotoBukti, 3);
  // Foto halaman depan FASIH (tag {%halaman_depan}), khusus Berkas Pembayaran
  // Termin II. Dicari dari sheet "Fasih-Cover" berdasarkan Email PML/PPL yang
  // sedang diproses (email = variabel yang sama dipakai untuk bapp/bast di atas).
  const halamanDepan = findHalamanDepanUrl(halamanDepanMap, email);

  // Variabel tanpa prefix dipertahankan untuk kompatibilitas dengan template lama.
  // Variabel berprefix memberi ruang bagi template gabungan saat ada nama tag yang
  // bentrok, misalnya {bapp_nomor_surat} dan {bast_nomor_surat}.
  return {
    ...bapp,
    ...bast,
    ...suratPernyataan,
    ...prefixTemplateData("bapp", bapp),
    ...prefixTemplateData("bast", bast),
    ...prefixTemplateData("surat_pernyataan", suratPernyataan),
    role,
    jenis: role,
    jenis_petugas: role,
    nama: displayName,
    nama_petugas: displayName,
    nama_pml: isPml ? displayName : cleanText(pembayaranRingkasan.nama_pml || firstLampiran?.nama_pml || ""),
    nama_pengawas: isPml ? displayName : cleanText(pembayaranRingkasan.nama_pengawas || firstLampiran?.nama_pml || ""),
    email_pengawas: isPml ? email : cleanText(pembayaranRingkasan.email_pengawas || firstLampiran?.email_pengawas || ""),
    nama_ppl: isPml ? cleanText(firstLampiranTableRow.nama_ppl || firstLampiran?.nama_ppl || "") : displayName,
    email,
    nik,
    nomor_spk: cleanText(pembayaranRingkasan.nomor_spk || nomorKontrak),
    nomor_kontrak: nomorKontrak,
    nomor_perjanjian: bast.nomor_perjanjian || nomorKontrak,
    nomor_surat: bast.nomor_surat || bapp.nomor_surat || "",
    nomor_prefix: extractNomorPrefix(nomorKontrak),

    // Ringkasan PML/PPL langsung dari kolom sheet Pembayaran.
    sls_total: pembayaranRingkasan.sls_total,
    sls_40: pembayaranRingkasan.sls_40,
    sls_60: pembayaranRingkasan.sls_60,
    sls_ongoing: pembayaranRingkasan.sls_ongoing,
    sls_selesai_sedang_dikerjakan: pembayaranRingkasan.sls_ongoing,
    persentase_sls: pembayaranRingkasan.persentase_sls,
    tanggal_screenshot: pembayaranRingkasan.tanggal_screenshot,
    // Ringkasan target/realisasi memakai agregat sheet Data per SLS bila tersedia.
    // Realisasi yang dipakai adalah kolom "Dengan Tidak Ditemukan".
    pre_keluarga: summaryTargetKeluarga,
    pre_usaha: summaryTargetUsaha,
    jumlah_pre: summaryTargetJumlah,
    target_keluarga: summaryTargetKeluarga,
    target_usaha: summaryTargetUsaha,
    target_jumlah: summaryTargetJumlah,
    target_prelist: summaryTargetJumlah,
    prelist_total: summaryTargetJumlah,

    realisasi_keluarga: summaryRealisasiKeluarga,
    realisasi_usaha: summaryRealisasiUsaha,
    jumlah_realisasi: summaryRealisasiJumlah,
    realisasi: summaryRealisasiJumlah,
    realisasi_hasil_pendataan: summaryRealisasiJumlah,
    realisasi_total: summaryRealisasiJumlah,

    persentase: summaryPersentase,
    pesentase: summaryPersentase,
    persentase_prelist: summaryPersentase,
    persentase_pendataan: summaryPersentase,
    // Keterangan/status tidak ditulis pada pemberkasan PML maupun PPL.
    status: "",
    keterangan: "",
    flag: pembayaranRingkasan.flag,
    kecamatan: pembayaranRingkasan.kecamatan,

    // Rata-rata aritmetika Persentase Pendataan seluruh PPL yang masuk ke tabel.
    // Setiap PPL dihitung sekali walaupun memiliki beberapa baris wilayah di Lampiran.
    average_persentase: averagePersentase.formatted,
    rata_rata_persentase: averagePersentase.formatted,
    average_persentase_raw: averagePersentase.raw == null ? "" : averagePersentase.raw,
    jumlah_persentase_dihitung: averagePersentase.count,

    // Diagnostik/filter Status SLS.
    status_filter_aktif: statusSlsRowsHaveUsableFilter(statusSlsRows),
    jumlah_status_sls_sumber: statusSlsRows.length,
    jumlah_baris_lampiran_sebelum_filter: lampiranRows.length,
    jumlah_baris_tabel_setelah_filter: lampiranTableRows.length,

    // Data tabel beban kerja pada halaman terakhir, langsung dari sheet Data per SLS.
    jumlah_baris_data_per_sls: bebanKerja.rows.length,
    beban_kerja_rows: bebanKerja.rows,
    data_per_sls_rows: bebanKerja.rows,
    rincian_data_per_sls: bebanKerja.rows,
    beban_kerja_total: bebanKerja.total,
    total_target_keluarga: bebanKerja.total.target_keluarga,
    total_target_usaha: bebanKerja.total.target_usaha,
    total_target_jumlah: bebanKerja.total.target_jumlah,
    total_realisasi_keluarga: bebanKerja.total.realisasi_keluarga,
    total_realisasi_usaha: bebanKerja.total.realisasi_usaha,
    total_realisasi_jumlah: bebanKerja.total.realisasi_jumlah,
    jumlah_approve_pml: isPml ? bebanKerja.total.realisasi_jumlah : "",
    total_jumlah_approve_pml: isPml ? bebanKerja.total.realisasi_jumlah : "",
    total_persentase_data_per_sls: bebanKerja.total.persentase,
    keterangan_pembayaran_data_per_sls: "",
    sumber_realisasi_pml: isPml && approveByPmlRows.length > 0 ? "Approve by PML" : "Data per SLS",
    jumlah_baris_approve_by_pml: approveByPmlRows.length,

    // Seluruh alias loop tabel memakai struktur sheet Lampiran. Kolom progres pada
    // tiap peserta diperkaya dari sheet Pembayaran melalui Email Pencacah/Nama PPL.
    peserta: lampiranTableRows,
    pembayaran: lampiranTableRows,
    pembayaran_rows: lampiranTableRows,
    rincian_pembayaran: lampiranTableRows,
    tabel_ppl: lampiranTableRows,
    ppl_rows: lampiranTableRows,
    lampiran_peserta: lampiranTableRows,
    bapp_peserta: bapp.peserta || [],
    bast_peserta: bast.peserta || [],
    foto: fotoBukti,
    foto_bukti: fotoBukti,
    jumlah_foto: fotoBukti.length,
    jumlah_foto_bukti: fotoBukti.length,
    foto_rows: fotoRows,
    halaman_depan: halamanDepan,
    tampil_surat_pernyataan: isPml,
    is_pml: isPml,
    is_ppl: !isPml,
  };
}

export const WORDPROCESSING_ML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const XML_NS = "http://www.w3.org/XML/1998/namespace";

export function wordLocalName(node) {
  return String(node?.localName || node?.nodeName || "").split(":").pop();
}

export function getDirectWordChildren(node, localName) {
  return Array.from(node?.childNodes || []).filter(
    (child) => child?.nodeType === 1 && wordLocalName(child) === localName
  );
}

export function getWordNodeText(node) {
  const textNodes = node?.getElementsByTagNameNS
    ? Array.from(node.getElementsByTagNameNS(WORDPROCESSING_ML_NS, "t"))
    : Array.from(node?.getElementsByTagName?.("w:t") || []);
  return textNodes.map((item) => item.textContent || "").join("");
}

export function createWordElement(xmlDoc, localName) {
  return xmlDoc.createElementNS(WORDPROCESSING_ML_NS, `w:${localName}`);
}

export function setWordCellText(xmlDoc, cell, value) {
  if (!cell) return;

  const text = String(value ?? "");
  const paragraphs = Array.from(
    cell.getElementsByTagNameNS(WORDPROCESSING_ML_NS, "p")
  );
  let paragraph = paragraphs[0];

  if (!paragraph) {
    paragraph = createWordElement(xmlDoc, "p");
    cell.appendChild(paragraph);
  }

  const firstRun = Array.from(
    paragraph.getElementsByTagNameNS(WORDPROCESSING_ML_NS, "r")
  )[0];
  const firstRunProperties = firstRun
    ? Array.from(firstRun.childNodes || []).find(
        (child) => child?.nodeType === 1 && wordLocalName(child) === "rPr"
      )
    : null;

  // Pertahankan properti paragraf, lalu ganti isi sel dengan satu run baru.
  for (const child of Array.from(paragraph.childNodes || [])) {
    if (!(child?.nodeType === 1 && wordLocalName(child) === "pPr")) {
      paragraph.removeChild(child);
    }
  }

  // Hapus paragraf tambahan dari sel template agar tinggi baris tidak membengkak.
  for (const extraParagraph of paragraphs.slice(1)) {
    extraParagraph.parentNode?.removeChild(extraParagraph);
  }

  const run = createWordElement(xmlDoc, "r");
  if (firstRunProperties) run.appendChild(firstRunProperties.cloneNode(true));

  const textNode = createWordElement(xmlDoc, "t");
  if (/^\s|\s$/.test(text)) textNode.setAttributeNS(XML_NS, "xml:space", "preserve");
  textNode.textContent = text;
  run.appendChild(textNode);
  paragraph.appendChild(run);
}

export function setWordRowValues(xmlDoc, rowNode, values = []) {
  const cells = getDirectWordChildren(rowNode, "tc");
  values.forEach((value, index) => {
    if (cells[index]) setWordCellText(xmlDoc, cells[index], value);
  });
}

export function fillBebanKerjaTableInDocxZip(zip, workloadRows = [], workloadTotal = {}) {
  const documentFile = zip?.file?.("word/document.xml");
  if (!documentFile || typeof DOMParser === "undefined" || typeof XMLSerializer === "undefined") {
    console.warn("Tabel beban kerja tidak dapat diisi karena document.xml atau XML DOM tidak tersedia.");
    return;
  }

  const xmlText = documentFile.asText();
  const xmlDoc = new DOMParser().parseFromString(xmlText, "application/xml");
  const parserErrors = xmlDoc.getElementsByTagName("parsererror");
  if (parserErrors.length > 0) {
    console.warn("Gagal membaca document.xml untuk mengisi tabel beban kerja.");
    return;
  }

  const tables = Array.from(
    xmlDoc.getElementsByTagNameNS(WORDPROCESSING_ML_NS, "tbl")
  );
  const targetTable = tables.find((table) => {
    const tableText = getWordNodeText(table).replace(/\s+/g, " ").toUpperCase();
    return tableText.includes("TARGET PRELIST AWAL") &&
      tableText.includes("USERNAME SOBAT") &&
      tableText.includes("SLS/SUB-SLS");
  });

  if (!targetTable) {
    console.warn("Tabel Beban Kerja pada halaman terakhir template tidak ditemukan.");
    return;
  }

  const tableRows = getDirectWordChildren(targetTable, "tr");
  if (tableRows.length < 5) {
    console.warn("Struktur tabel Beban Kerja tidak sesuai template.");
    return;
  }

  // Tiga baris pertama adalah header. Baris dengan teks "Jumlah" adalah footer.
  // Template PPL memiliki 16 kolom detail, sedangkan template PML yang dikirim
  // memiliki 14 kolom detail. Keduanya ditangani tanpa mengubah format Word.
  const headerRowCount = 3;
  let totalRow = tableRows.find((row, index) => {
    if (index < headerRowCount) return false;
    const cells = getDirectWordChildren(row, "tc");
    const firstCellText = getWordNodeText(cells[0] || "")
      .replace(/\s+/g, " ")
      .trim()
      .toUpperCase();
    return firstCellText === "JUMLAH";
  });
  if (!totalRow) totalRow = tableRows[tableRows.length - 1];

  const totalRowIndex = tableRows.indexOf(totalRow);
  const candidateDetailRows = tableRows.slice(headerRowCount, totalRowIndex);
  const templateRow = candidateDetailRows
    .map((row) => ({ row, count: getDirectWordChildren(row, "tc").length }))
    .sort((a, b) => b.count - a.count)[0]?.row;

  if (!templateRow) {
    console.warn("Baris template kosong pada tabel Beban Kerja tidak ditemukan.");
    return;
  }

  const detailColumnCount = getDirectWordChildren(templateRow, "tc").length;
  if (![14, 16].includes(detailColumnCount)) {
    console.warn(`Jumlah kolom tabel Beban Kerja tidak didukung: ${detailColumnCount}.`);
    return;
  }

  const templateClone = templateRow.cloneNode(true);
  const minimumTemplateRows = Math.max(totalRowIndex - headerRowCount, 1);
  for (const row of candidateDetailRows) {
    targetTable.removeChild(row);
  }

  const buildDetailValues = (row) => {
    const common = [
      row.no,
      row.nama_pml,
      row.username_pml,
      row.nama_ppl,
      row.username_ppl,
      row.kdkec,
      row.kddesa,
      row.kode_sls,
      row.target_keluarga,
      row.target_usaha,
      row.target_jumlah,
    ];

    // PPL: Realisasi Keluarga, Usaha, Jumlah.
    if (detailColumnCount === 16) {
      return [
        ...common,
        row.realisasi_keluarga,
        row.realisasi_usaha,
        row.realisasi_jumlah,
        row.persentase,
        "", // Kolom Keterangan PPL dikosongkan.
      ];
    }

    // PML: template hanya menyediakan satu kolom Realisasi, sehingga dipakai
    // nilai Jumlah dari blok "Dengan Tidak Ditemukan".
    return [
      ...common,
      row.realisasi_jumlah,
      row.persentase,
      "", // Kolom Keterangan PML dikosongkan.
    ];
  };

  // Pertahankan minimal jumlah baris bawaan template agar posisi footer stabil.
  const rowCountToInsert = Math.max(workloadRows.length, minimumTemplateRows);
  for (let rowIndex = 0; rowIndex < rowCountToInsert; rowIndex++) {
    const row = workloadRows[rowIndex] || null;
    const clonedRow = templateClone.cloneNode(true);
    const values = row
      ? buildDetailValues(row)
      : Array(detailColumnCount).fill("");

    setWordRowValues(xmlDoc, clonedRow, values);
    targetTable.insertBefore(clonedRow, totalRow);
  }

  const totalCellCount = getDirectWordChildren(totalRow, "tc").length;
  const hasRows = workloadRows.length > 0;
  const valueOrBlank = (value) => hasRows ? value : "";

  // Footer PPL memiliki 9 sel langsung: satu sel gabungan + 8 nilai.
  // Footer PML memiliki 7 sel langsung: satu sel gabungan + 6 nilai.
  const totalValues = detailColumnCount === 16
    ? [
        "Jumlah",
        valueOrBlank(workloadTotal?.target_keluarga),
        valueOrBlank(workloadTotal?.target_usaha),
        valueOrBlank(workloadTotal?.target_jumlah),
        valueOrBlank(workloadTotal?.realisasi_keluarga),
        valueOrBlank(workloadTotal?.realisasi_usaha),
        valueOrBlank(workloadTotal?.realisasi_jumlah),
        valueOrBlank(workloadTotal?.persentase),
        "", // Footer Keterangan PPL dikosongkan.
      ]
    : [
        "Jumlah",
        valueOrBlank(workloadTotal?.target_keluarga),
        valueOrBlank(workloadTotal?.target_usaha),
        valueOrBlank(workloadTotal?.target_jumlah),
        valueOrBlank(workloadTotal?.realisasi_jumlah),
        valueOrBlank(workloadTotal?.persentase),
        "", // Footer Keterangan PML dikosongkan.
      ];

  setWordRowValues(xmlDoc, totalRow, totalValues.slice(0, totalCellCount));

  const serialized = new XMLSerializer().serializeToString(xmlDoc);
  zip.file("word/document.xml", serialized);
}

export async function createBerkasPembayaranBlobFromTemplateBuffer(
  templateArrayBuffer,
  formValues,
  record,
  role,
  nikLookup
) {
  // Peta Email -> foto halaman depan FASIH (sheet "Fasih-Cover"). Fetch
  // di-cache di dalam parsers.js, jadi aman dipanggil per-record/per-batch.
  const halamanDepanMap = await fetchHalamanDepanMapByEmail();

  // Bangun data sebelum membuat Docxtemplater agar URL foto bisa di-prefetch.
  const templateData = buildBerkasPembayaranTemplateData(
    formValues || {},
    record || {},
    role,
    nikLookup,
    halamanDepanMap
  );

  console.log("Generate berkas pembayaran:", {
    role,
    nama: templateData?.nama_petugas,
    jumlahFoto: templateData?.jumlah_foto,
    jumlahBarisFoto: templateData?.foto_rows?.length || 0,
    halamanDepan: templateData?.halaman_depan ? "ada" : "kosong",
  });

  // FIX PPL:
  // Download seluruh foto lebih dulu. Setelah tahap ini ImageModule tidak perlu
  // mengembalikan Promise saat Docxtemplater memproses loop gambar.
  await prefetchFotoBuktiForTemplate(templateData);

  const zip = new PizZip(templateArrayBuffer);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    modules: [createFotoBuktiImageModule()],
  });

  // Jangan gunakan await doc.renderAsync(templateData) untuk image module gratis
  // pada loop foto. Foto sudah tersedia di cache sehingga render sinkron aman.
  doc.render(templateData);

  console.log("Jumlah foto:", templateData.jumlah_foto, templateData.foto_bukti);

  // Template tetap dipakai apa adanya. Kode hanya mengganti baris kosong pada tabel
  // Beban Kerja halaman terakhir dengan data dari sheet Data per SLS.
  fillBebanKerjaTableInDocxZip(
    doc.getZip(),
    templateData.beban_kerja_rows || [],
    templateData.beban_kerja_total || {}
  );

  return zipToDocxBlob(doc.getZip());
}

export async function generateSingleBerkasPembayaran(templateUrl, formValues, record, role, nikLookup) {
  if (!record) throw new Error("Data berkas pembayaran tidak ditemukan.");
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template berkas pembayaran: ${response.status} ${response.statusText}`);
  const templateArrayBuffer = await response.arrayBuffer();
  const blob = await createBerkasPembayaranBlobFromTemplateBuffer(templateArrayBuffer, formValues, record, role, nikLookup); // ⬅️ tambah await
  saveAs(blob, `BERKAS PEMBAYARAN ${role} - ${sanitizeFileName(record.displayName)}.docx`);
}

export const BERKAS_PEMBAYARAN_ZIP_BATCH_SIZE = 100;

export async function generateBerkasPembayaran(templateUrl, formValues, records, role, nikLookup, onProgress) {
  const entries = Array.isArray(records) ? records.filter(Boolean) : [];
  if (entries.length === 0) throw new Error(`Tidak ada data berkas pembayaran ${role}.`);

  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template berkas pembayaran: ${response.status} ${response.statusText}`);
  const templateArrayBuffer = await response.arrayBuffer();
  const totalBatches = Math.ceil(entries.length / BERKAS_PEMBAYARAN_ZIP_BATCH_SIZE);

  for (let batchIndex = 0; batchIndex < totalBatches; batchIndex++) {
    const batchEntries = entries.slice(
      batchIndex * BERKAS_PEMBAYARAN_ZIP_BATCH_SIZE,
      (batchIndex + 1) * BERKAS_PEMBAYARAN_ZIP_BATCH_SIZE
    );
    const files = [];
    const fileNameCounts = new Map();

    for (const record of batchEntries) {
      const blob = await createBerkasPembayaranBlobFromTemplateBuffer(   // ⬅️ tambah await
        templateArrayBuffer, formValues || {}, record, role, nikLookup
      );
      const baseName = sanitizeFileName(record.displayName);
      const count = (fileNameCounts.get(upperText(baseName)) || 0) + 1;
      fileNameCounts.set(upperText(baseName), count);
      const duplicateSuffix = count > 1 ? ` (${count})` : "";
      files.push({
        name: `BERKAS PEMBAYARAN ${role} - ${baseName}${duplicateSuffix}.docx`,
        blob,
      });
    }

    if (typeof onProgress === "function") {
      onProgress({ batchIndex: batchIndex + 1, totalBatches, totalRows: entries.length });
    }

    if (files.length === 1 && totalBatches === 1) {
      saveAs(files[0].blob, files[0].name);
      continue;
    }

    const batchSuffix = totalBatches > 1 ? ` - Bagian ${batchIndex + 1} dari ${totalBatches}` : "";
    await downloadMultipleAsZip(
      files,
      `BERKAS PEMBAYARAN ${role} ${cleanText(formValues?.tanggal_surat || "SE2026")}${batchSuffix}.zip`
    );
  }
}

export function normalizePmlProgressKey(value) {
  return upperText(value).replace(/^@/, "").replace(/\s+/g, " ");
}

export function findDataPmlProgressRow(dataPmlProgressData = [], nama = "", email = "") {
  const namaKey = normalizePmlProgressKey(nama);
  const emailLocal = normalizePmlProgressKey(email).includes("@")
    ? normalizePmlProgressKey(email).split("@")[0]
    : normalizePmlProgressKey(email);

  // Prioritas: cocokkan lewat Username Sobat (biasanya = bagian depan email), lalu Nama.
  let row = null;
  if (emailLocal) {
    row = dataPmlProgressData.find(
      (r) => normalizePmlProgressKey(r.username_sobat_pml) === emailLocal
    );
  }
  if (!row && namaKey) {
    row = dataPmlProgressData.find(
      (r) => normalizePmlProgressKey(r.nama_pml) === namaKey
    );
  }
  return row || null;
}

