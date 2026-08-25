// ============================================================
// Portal Administrasi SE2026 — bagian: bast
// ============================================================

import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { saveAs } from "file-saver";

import { downloadMultipleAsZip, formatKodeNama, groupLampiranRows, sanitizeFileName } from "./docGenerators";
import { cleanText, upperText } from "./helpers";
import { extractNomorPrefix, getBappDateParts } from "./parsers";

// ─── BAST ─────────────────────────────────────────────────
export const BAST_NILAI_PERJANJIAN = {
  PML: { nilai: "Rp5.831.000,00", terbilang: "Lima juta delapan ratus tiga puluh satu ribu rupiah" },
  PPL: { nilai: "Rp5.534.000,00", terbilang: "Lima juta lima ratus tiga puluh empat ribu rupiah" },
};

// Nomor surat: pakai nomor kontrak petugas sebagai sumber nomor urut
// (pola yang sama dipakai buildBappTemplateData -> nomor_prefix).
export function buildBastNomorSurat(nomorKontrak, jenis) {
  const prefix = extractNomorPrefix(nomorKontrak) || "...";
  const suffix = jenis === "PML"
    ? "BAST-I-SE2026/PML/3172/SS.340/2026"
    : "BAST-I-SE2026/PPL/3172/SS.330/2026";
  return `B-${prefix}/${suffix}`;
}

export function buildBastTemplateData(formValues, personRows, jenis, nikLookup) {
  const isPml = upperText(jenis) === "PML";
  const grouped = groupLampiranRows(personRows || [], jenis);

  const namaPetugas = isPml
    ? cleanText(personRows?.[0]?.nama_pml)
    : cleanText(personRows?.[0]?.nama_ppl);

  const nomorKontrak = cleanText(
    isPml ? personRows?.[0]?.nomor_kontrak_pml : personRows?.[0]?.nomor_kontrak_ppl
  );

  const totalJumlah = grouped.reduce((sum, r) => sum + (r.jumlah || 0), 0);
  const dateParts = getBappDateParts(formValues?.tanggal_surat || "");
  const nilai = BAST_NILAI_PERJANJIAN[isPml ? "PML" : "PPL"];
  const nik = nikLookup?.get(upperText(namaPetugas)) || "";

  return {
    nomor_surat: buildBastNomorSurat(nomorKontrak, isPml ? "PML" : "PPL"),
    nomor_perjanjian: nomorKontrak || "...",
    hari: dateParts.hari_terbilang,
    tanggal_terbilang: dateParts.tanggal_terbilang,
    tanggal: dateParts.tanggal,
    bulan: dateParts.bulan,
    bulan_terbilang: dateParts.bulan_terbilang,
    nama: namaPetugas,
    nik,
    jumlah: totalJumlah,
    nilai_perjanjian: nilai.nilai,
    nilai_perjanjian_terbilang: nilai.terbilang,
    peserta: grouped.map((r, idx) => ({
      no: idx + 1,
      nama_petugas: isPml
          ? cleanText(r.nama_ppl)
          : namaPetugas,
      kecamatan: formatKodeNama(r.kdkec, r.kecamatan),
      kelurahan: formatKodeNama(r.kddesa, r.kelurahan),
      jumlah: r.jumlah || 0,
    })),
  };
}

// ── 5) CREATE / DOWNLOAD FUNCTIONS ───────────────────────────
// Mengikuti pola createLampiranBlobFromTemplateBuffer + generateLampiran,
// supaya fetch template hanya sekali lalu dipakai berulang untuk tiap orang.

export function createBastBlobFromTemplateBuffer(templateArrayBuffer, formValues, personRows, jenis, nikLookup) {
  const zip = new PizZip(templateArrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(buildBastTemplateData(formValues || {}, personRows || [], jenis, nikLookup));
  return doc.getZip().generate({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

export async function generateSingleBast(templateUrl, formValues, personRows, jenis, nikLookup, displayName) {
  if (!personRows || personRows.length === 0) throw new Error("Tidak ada data untuk BAST yang dipilih");
  const templateResponse = await fetch(templateUrl);
  if (!templateResponse.ok) throw new Error(`Gagal memuat template BAST: ${templateResponse.status} ${templateResponse.statusText}`);
  const templateArrayBuffer = await templateResponse.arrayBuffer();
  const blob = createBastBlobFromTemplateBuffer(templateArrayBuffer, formValues || {}, personRows, jenis, nikLookup);
  const safeName = sanitizeFileName(displayName || `${jenis}-bast`);
  saveAs(blob, `BAST ${jenis} - ${safeName}.docx`);
}

export const BAST_ZIP_BATCH_SIZE = 150;

export async function generateBast(templateUrl, formValues, lampiranRows, jenis, nikLookup, onProgress) {
  const sourceRows = lampiranRows || [];
  const isPml = upperText(jenis) === "PML";

  // Kelompokkan baris per-orang, sama persis dengan logika di generateLampiran()
  // supaya email dipakai sebagai kunci utama (menghindari salah gabung nama kembar).
  const groups = new Map();
  for (const row of sourceRows) {
    const namaPml = cleanText(row.nama_pml || "");
    const namaPpl = cleanText(row.nama_ppl || "");
    const emailPml = cleanText(row.email_pengawas || "");
    const emailPpl = cleanText(row.email_pencacah || "");
    const displayName = isPml ? namaPml : namaPpl;
    if (!displayName) continue;
    const emailKey = upperText(isPml ? emailPml : emailPpl);
    const identity = emailKey || `NAMA::${upperText(displayName)}`;
    if (!groups.has(identity)) groups.set(identity, { displayName, rows: [] });
    groups.get(identity).rows.push(row);
  }

  if (groups.size === 0) throw new Error(`Tidak ada data ${jenis}`);

  const templateResponse = await fetch(templateUrl);
  if (!templateResponse.ok) throw new Error(`Gagal memuat template BAST: ${templateResponse.status} ${templateResponse.statusText}`);
  const templateArrayBuffer = await templateResponse.arrayBuffer();

  const entries = [...groups.values()];
  const totalBatches = Math.ceil(entries.length / BAST_ZIP_BATCH_SIZE);

  for (let batchIndex = 0; batchIndex < totalBatches; batchIndex++) {
    const batchEntries = entries.slice(batchIndex * BAST_ZIP_BATCH_SIZE, (batchIndex + 1) * BAST_ZIP_BATCH_SIZE);
    const zipFiles = [];
    for (const { displayName, rows } of batchEntries) {
      const blob = createBastBlobFromTemplateBuffer(templateArrayBuffer, formValues || {}, rows, jenis, nikLookup);
      zipFiles.push({ name: `BAST ${jenis} - ${sanitizeFileName(displayName)}.docx`, blob });
    }
    if (typeof onProgress === "function") {
      onProgress({ batchIndex: batchIndex + 1, totalBatches, totalRows: entries.length });
    }
    const batchSuffix = totalBatches > 1 ? ` - Bagian ${batchIndex + 1} dari ${totalBatches}` : "";
    await downloadMultipleAsZip(zipFiles, `BAST ${jenis} ${formValues?.tanggal_surat || "SE2026"}${batchSuffix}.zip`);
  }
}

