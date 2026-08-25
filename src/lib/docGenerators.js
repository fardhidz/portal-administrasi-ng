// ============================================================
// Portal Administrasi SE2026 — bagian: docGenerators
// ============================================================

import PizZip from "pizzip";
import JSZip from "jszip";
import Docxtemplater from "docxtemplater";
import { saveAs } from "file-saver";

import { calcDurationDays, cleanText, formatRupiah, formatTanggalBulanIndonesia, formatTanggalIndonesia, formatTanggalLengkapIndonesia, normalizeJamIndonesia, sortDaftarHadirPeserta, sortPesertaByJabatanOrder, spellTerbilang, upperText } from "./helpers";
import { chunkFotoBuktiIntoRows, collectFotoBuktiFromApproveRows, createFotoBuktiImageModule, dedupeBappRows, extractNomorPrefix, filterApproveByPmlRowsForBappRow, getBappDateParts, zipToDocxBlob } from "./parsers";
import { TANDA_TERIMA_LAPANGAN_TEMPLATE_URL, TANDA_TERIMA_TEMPLATE_URL } from "../data/templates";

// ─── TEMPLATE DATA BUILDERS ───────────────────────────────────────────────────

export function buildBappTemplateData(formValues, row = {}, role = "PML", approveByPmlRows = []) {
  const tanggalSurat = cleanText(formValues?.tanggal_surat || "");
  const nama = cleanText(row?.nama || row?.nama_pml || row?.nama_ppl || "");
  const jabatan = cleanText(row?.jabatan_raw || row?.jabatan || "");
  const wilayah = cleanText(row?.wilayah || row?.tempat || row?.asal || "");
  const dateParts = getBappDateParts(tanggalSurat);
  const nomorKontrak = cleanText(row?.nomor_spk || row?.nomor_kontrak || formValues?.nomor_kontrak || "");
  const fotoBukti = collectFotoBuktiFromApproveRows(approveByPmlRows, role);
  const fotoRows = chunkFotoBuktiIntoRows(fotoBukti, 3); // ganti 3 -> 2 kalau mau 2 foto/baris
  return {
    tanggal_surat: tanggalSurat,
    tanggal_surat_fmt: formatTanggalIndonesia(tanggalSurat),
    tanggal_surat_lengkap: formatTanggalLengkapIndonesia(tanggalSurat),
    hari_terbilang: dateParts.hari_terbilang,
    hari: dateParts.hari_terbilang,
    tanggal_terbilang: dateParts.tanggal_terbilang,
    tanggal: dateParts.tanggal,
    bulan: dateParts.bulan,
    bulan_terbilang: dateParts.bulan_terbilang,
    nama,
    nama_peserta: nama,
    nama_petugas: nama,
    jabatan,
    jabatan_peserta: jabatan,
    jabatan_petugas: jabatan,
    role,
    jenis: role,
    jenis_dokumen: "BAPP",
    wilayah,
    wilayah_tugas: wilayah,
    tempat: cleanText(formValues?.tempat || ""),
    email: cleanText(row?.email || ""),
    kelas: cleanText(row?.kelas || ""),
    gelombang: cleanText(row?.gelombang || ""),
    nomor_surat: cleanText(formValues?.nomor_surat || ""),
    nomor_dokumen: cleanText(formValues?.nomor_surat || ""),
    nomor_prefix: extractNomorPrefix(nomorKontrak),
    nomor_spk: nomorKontrak,
    nomor_kontrak: nomorKontrak,
    nik: cleanText(row?.nik || ""),
    nama_pml: cleanText(row?.nama_pml || row?.nama_pengawas || (role === "PML" ? nama : "")),
    nama_pengawas: cleanText(row?.nama_pengawas || row?.nama_pml || (role === "PML" ? nama : "")),
    email_pengawas: cleanText(row?.email_pengawas || row?.email_pml || (role === "PML" ? row?.email : "")),
    nama_ppl: cleanText(row?.nama_ppl || (role === "PPL" ? nama : "")),
    sls_total: cleanText(row?.sls_total || ""),
    sls_40: cleanText(row?.sls_40 || ""),
    sls_60: cleanText(row?.sls_60 || ""),
    sls_ongoing: cleanText(row?.sls_ongoing || row?.sls_selesai_sedang_dikerjakan || ""),
    sls_selesai_sedang_dikerjakan: cleanText(row?.sls_ongoing || row?.sls_selesai_sedang_dikerjakan || ""),
    persentase_sls: cleanText(row?.persentase_sls || ""),
    tanggal_screenshot: cleanText(row?.tanggal_screenshot || ""),
    target_prelist: cleanText(row?.target_prelist || row?.prelist_total || ""),
    prelist_total: cleanText(row?.prelist_total || row?.target_prelist || ""),
    realisasi_hasil_pendataan: cleanText(row?.realisasi_hasil_pendataan || row?.realisasi_total || ""),
    realisasi_total: cleanText(row?.realisasi_total || row?.realisasi_hasil_pendataan || ""),
    persentase_prelist: cleanText(row?.persentase_prelist || row?.persentase_pendataan || ""),
    persentase_pendataan: cleanText(row?.persentase_pendataan || row?.persentase_prelist || ""),
    flag: cleanText(row?.flag || ""),
    foto: fotoBukti,
    foto_bukti: fotoBukti,
    jumlah_foto: fotoBukti.length,
    jumlah_foto_bukti: fotoBukti.length,
    foto_rows: fotoRows,
  };
}

export async function createBappBlob(templateUrl, formValues, row, role, approveByPmlRows = []) {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template BAPP: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    modules: [createFotoBuktiImageModule()],
  });
  const fotoRows = filterApproveByPmlRowsForBappRow(approveByPmlRows, row || {}, role);
  // renderAsync wajib dipakai karena foto diambil lewat fetch() (asinkron).
  await doc.renderAsync(buildBappTemplateData(formValues || {}, row || {}, role, fotoRows));
  return zipToDocxBlob(doc.getZip());
}

export async function generateSingleBapp(templateUrl, formValues, row, role, approveByPmlRows = []) {
  const blob = await createBappBlob(templateUrl, formValues || {}, row || {}, role, approveByPmlRows);
  const safeName = sanitizeFileName(cleanText(row?.nama || `${role}-bapp`));
  saveAs(blob, `BAPP ${role} - ${safeName}.docx`);
}

export const BAPP_ZIP_BATCH_SIZE = 150;

export async function generateBapp(templateUrl, formValues, rows, role, onProgress, approveByPmlRows = []) {
  if (!rows || rows.length === 0) throw new Error("Tidak ada data BAPP untuk role yang dipilih.");
  const uniqueRows = dedupeBappRows(rows);
  const totalBatches = Math.ceil(uniqueRows.length / BAPP_ZIP_BATCH_SIZE);

  for (let batchIndex = 0; batchIndex < totalBatches; batchIndex++) {
    const batchRows = uniqueRows.slice(
      batchIndex * BAPP_ZIP_BATCH_SIZE,
      (batchIndex + 1) * BAPP_ZIP_BATCH_SIZE
    );

    const files = [];
    for (const row of batchRows) {
      const blob = await createBappBlob(templateUrl, formValues || {}, row || {}, role, approveByPmlRows);
      files.push({
        name: `BAPP ${role} - ${sanitizeFileName(cleanText(row?.nama || "Tanpa Nama"))}.docx`,
        blob,
      });
    }

    if (typeof onProgress === "function") {
      onProgress({
        batchIndex: batchIndex + 1,
        totalBatches,
        totalRows: uniqueRows.length,
      });
    }

    if (files.length === 1 && totalBatches === 1) {
      saveAs(files[0].blob, files[0].name);
      continue;
    }

    const batchSuffix = totalBatches > 1 ? ` - Bagian ${batchIndex + 1} dari ${totalBatches}` : "";
    await downloadMultipleAsZip(files, `BAPP ${role} ${cleanText(formValues?.tanggal_surat || "SE2026")}${batchSuffix}.zip`);
  }
}

export async function createSuratPernyataanPenyelesaianLapanganBlob(templateUrl, formValues, row) {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template surat: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render({
    nomor_prefix: cleanText(row?.nomor_prefix || extractNomorPrefix(row?.nomor_spk || row?.nomor_kontrak || "")),
    nama_petugas: cleanText(row?.nama || row?.nama_pml || row?.nama_ppl || ""),
    nik: cleanText(row?.nik || ""),
    nomor_kontrak: cleanText(row?.nomor_spk || row?.nomor_kontrak || ""),
  });
  return doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

export async function generateSingleSuratPernyataanPenyelesaianLapangan(templateUrl, row, onProgress) {
  const blob = await createSuratPernyataanPenyelesaianLapanganBlob(templateUrl, {}, row);
  const safeName = sanitizeFileName(cleanText(row?.nama || "Tanpa Nama"));
  if (typeof onProgress === "function") {
    onProgress({ batchIndex: 1, totalBatches: 1, totalRows: 1 });
  }
  saveAs(blob, `Surat Pernyataan Penyelesaian Lapangan - ${safeName}.docx`);
}

export async function generateSuratPernyataanPenyelesaianLapangan(templateUrl, rows, onProgress) {
  if (!rows || rows.length === 0) throw new Error("Tidak ada data untuk Surat Pernyataan Penyelesaian Lapangan.");
  const uniqueRows = dedupeBappRows(rows);
  const totalBatches = Math.ceil(uniqueRows.length / BAPP_ZIP_BATCH_SIZE);

  for (let batchIndex = 0; batchIndex < totalBatches; batchIndex++) {
    const batchRows = uniqueRows.slice(batchIndex * BAPP_ZIP_BATCH_SIZE, (batchIndex + 1) * BAPP_ZIP_BATCH_SIZE);
    const files = [];
    for (const row of batchRows) {
      const blob = await createSuratPernyataanPenyelesaianLapanganBlob(templateUrl, {}, row);
      files.push({
        name: `Surat Pernyataan Penyelesaian Lapangan - ${sanitizeFileName(cleanText(row?.nama || "Tanpa Nama"))}.docx`,
        blob,
      });
    }

    if (typeof onProgress === "function") {
      onProgress({ batchIndex: batchIndex + 1, totalBatches, totalRows: uniqueRows.length });
    }

    const batchSuffix = totalBatches > 1 ? ` - Bagian ${batchIndex + 1} dari ${totalBatches}` : "";
    await downloadMultipleAsZip(files, `Surat Pernyataan Penyelesaian Lapangan${batchSuffix}.zip`);
  }
}

// DAFTAR HADIR
export const DAFTAR_HADIR_KEPALA_BPS = {
  nama:      "Widiastuti",
  jabatan:   "Penanggung Jawab",
  wilTugas:  "BPS Kota Jakarta Timur",
};

export function buildDaftarHadirTemplateData(formValues, peserta, namaInda, selectedFilterGroup = "") {
  const jamMulai        = normalizeJamIndonesia(formValues.jamMulai,   "07.30");
  const jamSelesai      = normalizeJamIndonesia(formValues.jamSelesai, "18.00");
  const tanggalFmt      = formatTanggalIndonesia(formValues.tanggal);
  const tanggalKegiatan = formatTanggalLengkapIndonesia(formValues.tanggal);
  const jamFmt          = `${jamMulai} - ${jamSelesai}`;
  const isPanitiaInda   = selectedFilterGroup === "panitia-inda";
  const isPmlPpl        = selectedFilterGroup === "pml-ppl";

  const hotelValue = cleanText(formValues.hotel || formValues.tempat || "").toLowerCase();
  const isHotelBwp = hotelValue.includes("bwp");
  const gelombangValue = cleanText(formValues.gelombang || "");

  let kepalaBpsEntry = {
    no:        1,
    nama:      "Widiastuti",
    jabatan:   "Penanggung Jawab",
    kecamatan: "BPS Kota Jakarta Timur",
    wil_tugas: "BPS Kota Jakarta Timur",
    wilTugas:  "BPS Kota Jakarta Timur",
  };

  if (gelombangValue === "4") {
    if (hotelValue.includes("bwp")) {
      kepalaBpsEntry = {
        no:        1,
        nama:      "Budi Utami",
        jabatan:   "Penanggung Jawab",
        kecamatan: "BPS Kota Jakarta Timur",
        wil_tugas: "BPS Kota Jakarta Timur",
        wilTugas:  "BPS Kota Jakarta Timur",
      };
    } else if (hotelValue.includes("stis")) {
      kepalaBpsEntry = {
        no:        1,
        nama:      "Widiastuti",
        jabatan:   "Penanggung Jawab",
        kecamatan: "BPS Kota Jakarta Timur",
        wil_tugas: "BPS Kota Jakarta Timur",
        wilTugas:  "BPS Kota Jakarta Timur",
      };
    }
  }

  const includeKepalaBps = (isHotelBwp || (hotelValue.includes("stis") && gelombangValue === "4")) && !isPmlPpl;

  return {
    tanggal_kegiatan: tanggalKegiatan,
    tanggal_aja:      tanggalFmt,
    hari_tanggal:     tanggalFmt,
    tanggal:          tanggalFmt,
    jam_mulai:        jamMulai,
    jam_selesai:      jamSelesai,
    jam:              jamFmt,
    jam_kegiatan:     jamFmt,
    tempat:           formValues.tempat || formValues.hotel || "",
    tempat_kegiatan:  formValues.tempat || formValues.hotel || "",
    gelombang:        formValues.gelombang || "",
    kelas:            isPanitiaInda ? "-" : (formValues.kelas || ""),
    nama_inda:        isPanitiaInda ? "Ir. Tristiati, MA" : (namaInda || ""),
    keterangan_ttd:   isPanitiaInda ? "Kepala Sub Bagian Umum" : (isPmlPpl ? "Instruktur Daerah" : ""),
    peserta: includeKepalaBps ? [
      kepalaBpsEntry,
      ...sortDaftarHadirPeserta(peserta || []).map((p, idx) => ({
        no:        idx + 2,
        nama:      p.nama     || "",
        jabatan:   p.jabatan  || "",
        kecamatan: p.wilTugas || "",
        wil_tugas: p.wilTugas || "",
        wilTugas:  p.wilTugas || "",
      })),
    ] : sortDaftarHadirPeserta(peserta || []).map((p, idx) => ({
      no:        idx + 1,
      nama:      p.nama     || "",
      jabatan:   p.jabatan  || "",
      kecamatan: p.wilTugas || "",
      wil_tugas: p.wilTugas || "",
      wilTugas:  p.wilTugas || "",
    })),
  };
}

export async function createDaftarHadirBlob(templateUrl, formValues, peserta, namaInda, selectedFilterGroup = "") {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(buildDaftarHadirTemplateData(formValues || {}, peserta || [], namaInda || "", selectedFilterGroup));
  return doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

export async function generateDaftarHadir(templateUrl, formValues, peserta, namaInda, selectedFilterGroup = "") {
  const blob = await createDaftarHadirBlob(templateUrl, formValues, peserta, namaInda, selectedFilterGroup);
  const safeGelombang = formValues?.gelombang || "X";
  const safeKelas     = selectedFilterGroup === "panitia-inda" ? "-" : (formValues?.kelas || "X");
  saveAs(blob, `Daftar Hadir Gelombang ${safeGelombang} Kelas ${safeKelas}.docx`);
}

// TANDA TERIMA
export function buildTandaTerimaTemplateData(formValues, peserta) {
  const tanggalFmt = formatTanggalIndonesia(formValues.tanggal);
  const tanggalKegiatan = formatTanggalLengkapIndonesia(formValues.tanggal);
  const filtered   = sortPesertaByJabatanOrder((peserta || []).filter(p => ["PML", "PPL"].includes(upperText(p.jabatan))));
  return {
    tanggal_kegiatan: tanggalKegiatan,
    tanggal_aja:      tanggalFmt,
    tanggal:          tanggalFmt,
    tempat:           formValues.tempat || formValues.hotel || "",
    gelombang:        formValues.gelombang || "",
    kelas:            formValues.kelas || "",
    peserta: filtered.map((p, idx) => ({
      no: idx + 1, nama: p.nama || "", jabatan: p.jabatan || "",
      kecamatan: p.wilTugas || "", wil_tugas: p.wilTugas || "", wilTugas: p.wilTugas || "",
    })),
  };
}

export async function createTandaTerimaBlob(templateUrl, formValues, peserta) {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(buildTandaTerimaTemplateData(formValues || {}, peserta || []));
  return doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

export async function generateTandaTerimaSmart(formValues, peserta, tandaTerimaType) {
  const hotelValue = cleanText(formValues?.tempat || formValues?.hotel || "").toLowerCase();

  if (tandaTerimaType === "mitra-umum") {
    const wilTugas = peserta?.[0]?.wilTugas ? cleanText(peserta[0].wilTugas) : "MITRA UMUM";
    const blob = await createTandaTerimaBlob(TANDA_TERIMA_LAPANGAN_TEMPLATE_URL, formValues || {}, peserta || []);
    saveAs(blob, `Tanda Terima Perlengkapan Lapangan - ${wilTugas}.docx`);
    return;
  }

  if (tandaTerimaType === "lapangan") {
    if (hotelValue === "stis") {
      const blob = await createTandaTerimaBlob(TANDA_TERIMA_LAPANGAN_TEMPLATE_URL, formValues || {}, peserta || []);
      saveAs(blob, `Tanda Terima Perlengkapan STIS Gelombang ${formValues?.gelombang || "X"}.docx`);
      return;
    }

    const groups = (peserta || []).reduce((acc, p) => {
      const key = cleanText(p.wilTugas) || "LAINNYA";
      (acc[key] = acc[key] || []).push(p);
      return acc;
    }, {});

    for (const [kec, list] of Object.entries(groups)) {
      const blob = await createTandaTerimaBlob(TANDA_TERIMA_LAPANGAN_TEMPLATE_URL, formValues || {}, list);
      const safeKec = kec || "LAINNYA";
      saveAs(blob, `Tanda Terima Perlengkapan ${safeKec} Gelombang ${formValues?.gelombang || "X"} Kelas ${formValues?.kelas || "X"}.docx`);
    }
    return;
  }

  const blob = await createTandaTerimaBlob(TANDA_TERIMA_TEMPLATE_URL, formValues || {}, peserta || []);
  saveAs(blob, `Tanda Terima Perlengkapan Gelombang ${formValues?.gelombang || "X"} Kelas ${formValues?.kelas || "X"}.docx`);
}

// SURAT PERNYATAAN KENDARAAN / SUPER KENDIS
export function buildSuratPernyataanKendaraanTemplateData(formValues, peserta) {
  const tanggalFmt = formatTanggalIndonesia(formValues.tanggal_surat || formValues.tanggal);
  const sorted     = sortPesertaByJabatanOrder(peserta || []);
  const hotelValue = cleanText(formValues.tempat || formValues.hotel || "").toLowerCase();

  let nomor_surtug_val = formValues.nomor_surat || formValues.nomor || "";
  let tanggal_kegiatan_val = formatTanggalLengkapIndonesia(formValues.tanggal) || "";
  let tanggal_surtug_val = tanggalFmt;

  if (hotelValue.includes("bwp")) {
    nomor_surtug_val = "B-999.1/3172/SS.220/2026";
    tanggal_kegiatan_val = "1 Juni - 3 Juni 2026";
    tanggal_surtug_val = "29 Mei 2026";
  } else if (hotelValue.includes("park")) {
    nomor_surtug_val = "B-999.3/3172/SS.220/2026";
    tanggal_kegiatan_val = "1 Juni - 3 Juni 2026";
    tanggal_surtug_val = "29 Mei 2026";
  } else if (hotelValue.includes("harper")) {
    nomor_surtug_val = "B-999.2/3172/SS.220/2026";
    tanggal_kegiatan_val = "1 Juni - 3 Juni 2026";
    tanggal_surtug_val = "29 Mei 2026";
  }

  return {
    tanggal_kegiatan: tanggal_kegiatan_val,
    nomor_surtug:     nomor_surtug_val,
    tanggal_surtug:   tanggal_surtug_val,
    tanggal_aja:      tanggalFmt,
    tanggal_surat:    tanggalFmt,
    tanggal:          tanggalFmt,
    tempat:           formValues.tempat || formValues.hotel || "",
    gelombang:        formValues.gelombang || "",
    kelas:            formValues.kelas || "-",
    peserta: sorted.map((p, idx) => {
      const pangkatGolRaw = cleanText(p.pangkatGol);
      const pangkatGolValue = (!pangkatGolRaw || upperText(pangkatGolRaw) === "#N/A") ? "-" : pangkatGolRaw;
      return {
        no:         idx + 1,
        nama:       p.nama || "",
        nik:        p.nik || "",
        jabatan:    p.jabatan || "",
        pangkat:    pangkatGolValue,
        pangkatGol: pangkatGolValue,
        pangkat_gol:pangkatGolValue,
        kecamatan:  p.wilTugas || "",
        wil_tugas:  p.wilTugas || "",
        wilTugas:   p.wilTugas || "",
      };
    }),
  };
}

export async function createSuratPernyataanKendaraanBlob(templateUrl, formValues, peserta) {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(buildSuratPernyataanKendaraanTemplateData(formValues || {}, peserta || []));
  return doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

export async function generateSuratPernyataanKendaraan(templateUrl, formValues, peserta) {
  const blob = await createSuratPernyataanKendaraanBlob(templateUrl, formValues, peserta);
  saveAs(blob, `Surat Pernyataan Kendaraan ${formValues?.tempat || "SE2026"} Gelombang ${formValues?.gelombang || "X"}.docx`);
}

// PENGELUARAN RIIL
export function buildPengeluaranRiilTemplateData(formValues, peserta = []) {
  const sorted    = sortPesertaByJabatanOrder(peserta || []);
  const biayaTotal = 510000 * sorted.length;
  return {
    tanggal_aja:      formatTanggalIndonesia(formValues.tanggal_surat || ""),
    tanggal_surat:    formatTanggalIndonesia(formValues.tanggal_surat || ""),
    biaya_total:      formatRupiah(biayaTotal),
    biaya_terbilang:  `${spellTerbilang(biayaTotal)} rupiah`,
    peserta: sorted.map((p, idx) => {
      const isPanitiaInda = ["PANITIA", "INDA"].includes(upperText(p.jabatan));
      const pangkatGolRaw = cleanText(p.pangkatGol);
      const pangkatGolValue = (!pangkatGolRaw || upperText(pangkatGolRaw) === "#N/A") ? "-" : pangkatGolRaw;
      return { no: idx + 1, nama: cleanText(p.nama), nik: cleanText(p.nik), jabatan: cleanText(p.jabatan), pangkat: isPanitiaInda ? pangkatGolValue : pangkatGolRaw };
    }),
  };
}

export async function createPengeluaranRiilBlob(templateUrl, formValues, peserta) {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(buildPengeluaranRiilTemplateData(formValues || {}, peserta || []));
  return doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

export async function generatePengeluaranRiil(templateUrl, formValues, peserta) {
  const blob = await createPengeluaranRiilBlob(templateUrl, formValues, peserta);
  saveAs(blob, `Daftar Pengeluaran Riil ${formValues.no || "SE2026"}.docx`);
}

// SPJ
export function buildSpjTemplateData(formValues, peserta = []) {
  const sorted = sortPesertaByJabatanOrder(peserta || []);
  const n = sorted.length;
  const total = 510000 * n;
  const tanggalAwal = formatTanggalBulanIndonesia(formValues.tanggal_awal_kegiatan || "");
  const tanggalAkhir = formatTanggalBulanIndonesia(formValues.tanggal_akhir_kegiatan || "");
  return {
    tanggal_aja:       formatTanggalIndonesia(formValues.tanggal_pelunasan || ""),
    tanggal_pelunasan: formatTanggalIndonesia(formValues.tanggal_pelunasan || ""),
    tanggal_awal:      tanggalAwal,
    tanggal_akhir:     tanggalAkhir,
    tanggal_awal_kegiatan: tanggalAwal,
    tanggal_akhir_kegiatan: tanggalAkhir,
    tempat: formValues.tempat || formValues.hotel || "",
    gelombang: formValues.gelombang || "",
    kelas: formValues.kelas || "-",
    kelompok_peserta: formValues.kelompokPeserta || "",
    kelompokPeserta: formValues.kelompokPeserta || "",
    jumlah_ok: n * 3,
    jumlah_uang: formatRupiah(170000 * n),
    total_jumlah_kotor: formatRupiah(total),
    total_jumlah_bersih: formatRupiah(total),
    total_terbilang: `${spellTerbilang(total)} rupiah`,
    ttd_kiri: formValues.ttd_kiri || "",
    ttd_kanan: formValues.ttd_kanan || "",
    peserta: sorted.map((p, idx) => ({
      no: idx + 1,
      nama: p.nama || "",
      nik: p.nik || "",
      jabatan: p.jabatan || "",
      kecamatan: p.wilTugas || "",
      wil_tugas: p.wilTugas || "",
      wilTugas: p.wilTugas || "",
    })),
  };
}

export async function createSpjBlob(templateUrl, formValues, peserta) {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(buildSpjTemplateData(formValues || {}, peserta || []));
  return doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

export async function generateSpj(templateUrl, formValues, peserta) {
  const blob = await createSpjBlob(templateUrl, formValues || {}, peserta || []);
  const safeKelompok = formValues?.kelompokPeserta || "SPJ";
  const safeTempat = formValues?.tempat || "SE2026";
  const safeGelombang = formValues?.gelombang || "X";
  const safeKelas = formValues?.kelas || "-";
  saveAs(blob, `SPJ ${safeKelompok} ${safeTempat} Gelombang ${safeGelombang} Kelas ${safeKelas}.docx`);
}

// SPD
export function buildSpdTemplateData(formValues, peserta = []) {
  const sorted       = sortPesertaByJabatanOrder(peserta || []);
  const tanggalAwal  = formatTanggalIndonesia(formValues.tanggal_awal_kegiatan || "");
  const tanggalAkhir = formatTanggalIndonesia(formValues.tanggal_akhir_kegiatan || "");
  const lamaHari     = calcDurationDays(formValues.tanggal_awal_kegiatan, formValues.tanggal_akhir_kegiatan);
  const tanggalSurat = formatTanggalIndonesia(formValues.tanggal_surat || formValues.tanggal || "");

  return {
    nomor_dokumen: formValues.nomor_dokumen || formValues.nomor || "",
    nomor: formValues.nomor_dokumen || formValues.nomor || "",
    tanggal_aja: tanggalSurat,
    tanggal_surat: tanggalSurat,
    tanggal: tanggalSurat,
    tempat: formValues.tempat || formValues.hotel || "",
    lokasi: formValues.lokasi || formValues.tempat || formValues.hotel || "",
    gelombang: formValues.gelombang || "",
    kelas: formValues.kelas || "-",
    tanggal_awal_kegiatan: tanggalAwal,
    tanggal_akhir_kegiatan: tanggalAkhir,
    lama_hari: lamaHari,
    lama: lamaHari,
    jumlah_peserta: sorted.length || 0,
    ttd_nama: formValues.namaKabps || "",
    ttd_nip: formValues.nipKabps || "",
    peserta: sorted.map((p, idx) => {
      const jabatan = cleanText(p.jabatan);
      const pangkatGolRaw = cleanText(p.pangkatGol);
      const isPanitia = upperText(jabatan) === "PANITIA";
      const pangkatValue = isPanitia && (!pangkatGolRaw || upperText(pangkatGolRaw) === "#N/A") ? "" : pangkatGolRaw;

      return {
        no: idx + 1,
        nama: cleanText(p.nama),
        nik: cleanText(p.nik),
        jabatan,
        pangkat: pangkatValue,
        pangkatGol: pangkatValue,
        pangkat_gol: pangkatValue,
        kecamatan: cleanText(p.wilTugas),
        wil_tugas: cleanText(p.wilTugas),
        wilTugas: cleanText(p.wilTugas),
        tanggal_awal_kegiatan: tanggalAwal,
        tanggal_akhir_kegiatan: tanggalAkhir,
        lama: lamaHari,
      };
    }),
  };
}

export function makeSlsKey(row) {
  return `${cleanText(row.kdkec)}|${cleanText(row.kddesa)}|${cleanText(row.kdsls)}|${cleanText(row.kdsubsls)}`;
}

export async function createSpdBlob(templateUrl, formValues, peserta) {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(buildSpdTemplateData(formValues || {}, peserta || []));
  return doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

export async function generateSpd(mainUrl, attachmentUrl, formValues, peserta) {
  const [blobMain, blobLampiran] = await Promise.all([
    createSpdBlob(mainUrl, formValues || {}, peserta || []),
    createSpdBlob(attachmentUrl, formValues || {}, peserta || []),
  ]);

  const safeNomor = formValues?.nomor_dokumen || formValues?.nomor || "SE2026";
  const safeTempat = formValues?.tempat || "SE2026";
  const safeGelombang = formValues?.gelombang || "X";
  const safeTanggal = formValues?.tanggal_surat || formValues?.tanggal || "";

  saveAs(blobMain, `SPD ${safeNomor} ${safeTempat} Gelombang ${safeGelombang} ${safeTanggal}.docx`);
  saveAs(blobLampiran, `Lampiran SPD ${safeNomor} ${safeTempat} Gelombang ${safeGelombang} ${safeTanggal}.docx`);
}

// SURAT TUGAS
export function buildSuratTugasTemplateData(formValues, peserta = []) {
  const sorted       = sortPesertaByJabatanOrder(peserta || []);
  const tanggalAwal  = formatTanggalIndonesia(formValues.tanggal_awal_kegiatan || "");
  const tanggalAkhir = formatTanggalIndonesia(formValues.tanggal_akhir_kegiatan || "");
  const lamaHari     = calcDurationDays(formValues.tanggal_awal_kegiatan, formValues.tanggal_akhir_kegiatan) || "";
  const tanggalSurat = formatTanggalIndonesia(formValues.tanggal_surat || "");

  return {
    nomor_surat:            formValues.nomor_surat || "",
    nomor:                  formValues.nomor_surat || "",
    tanggal_aja:            tanggalSurat,
    tanggal_surat:          tanggalSurat,
    tanggal:                tanggalSurat,
    tempat:                 formValues.tempat || formValues.hotel || "",
    lokasi:                 formValues.lokasi || formValues.tempat || formValues.hotel || "",
    gelombang:              formValues.gelombang || "",
    kelas:                  formValues.kelas || "-",
    tanggal_awal_kegiatan:  tanggalAwal,
    tanggal_akhir_kegiatan: tanggalAkhir,
    lama:                   lamaHari,
    lama_hari:              lamaHari,
    jumlah_peserta:         sorted.length || 0,
    peserta: sorted.map((p, idx) => ({
      no:          idx + 1,
      nama:        cleanText(p.nama),
      nik:         cleanText(p.nik),
      jabatan:     cleanText(p.jabatan),
      pangkat:     cleanText(p.pangkatGol),
      pangkatGol:  cleanText(p.pangkatGol),
      pangkat_gol: cleanText(p.pangkatGol),
      kecamatan:   cleanText(p.wilTugas),
      wil_tugas:   cleanText(p.wilTugas),
      wilTugas:    cleanText(p.wilTugas),
    })),
  };
}

export async function createSuratTugasBlob(templateUrl, formValues, peserta) {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(buildSuratTugasTemplateData(formValues || {}, peserta || []));
  return doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

export async function generateSuratTugas(templateUrl, formValues, peserta) {
  const blob = await createSuratTugasBlob(templateUrl, formValues || {}, peserta || []);
  const safeNomor = formValues?.nomor_surat || "SE2026";
  const safeTempat = formValues?.tempat || "SE2026";
  const safeGelombang = formValues?.gelombang || "X";
  const safeTanggal = formValues?.tanggal_surat || "";
  saveAs(blob, `Surat Tugas ${safeNomor} ${safeTempat} Gelombang ${safeGelombang} ${safeTanggal}.docx`);
}

// LAMPIRAN
export function formatKodeNama(kode, nama) {
  const kodeText = cleanText(kode);
  const namaText = cleanText(nama).toUpperCase();
  if (kodeText && namaText) return `${kodeText} ${namaText}`;
  return namaText || kodeText;
}

export function groupLampiranRows(lampiranRows = [], jenis = "PML") {
  const isPml = upperText(jenis) === "PML";

  const map = new Map();

  for (const r of lampiranRows || []) {
    const namaPml = cleanText(r.nama_pml);
    const namaPpl = cleanText(r.nama_ppl);

    const kec = cleanText(r.kdkec);
    const desa = cleanText(r.kddesa);

    // 🔥 FIX: untuk Lampiran PML, satu PML bisa membawahi banyak PPL yang
    // kebetulan bertugas di kecamatan/kelurahan yang SAMA. Kalau key grouping
    // hanya namaPml|kec|desa, semua PPL berbeda dengan kec/desa sama akan
    // ditumpuk jadi SATU baris saja (nama PPL pertama menang, PPL lain hilang).
    // Maka nama_ppl WAJIB ikut jadi bagian key supaya setiap PPL tetap dapat
    // baris sendiri-sendiri.
    const keyBase = isPml
      ? `${namaPml}|${namaPpl}`
      : `${kec}|${desa}`;

    if (!map.has(keyBase)) {
      map.set(keyBase, {
        nama_pml: namaPml,
        nama_ppl: namaPpl,
        email_pengawas: cleanText(r.email_pengawas),
        email_pencacah: cleanText(r.email_pencacah),
        kdprov: cleanText(r.kdprov),
        kdkab: cleanText(r.kdkab),
        kdkec: kec,
        kddesa: desa,
        nmprov: cleanText(r.nmprov).toUpperCase(),
        nmkab: cleanText(r.nmkab).toUpperCase(),
        kecamatan: cleanText(r.kecamatan).toUpperCase(),
        kelurahan: cleanText(r.kelurahan).toUpperCase(),

        // Nomor kontrak dibawa dari baris pertama tiap grup. Diasumsikan konsisten
        // untuk satu petugas yang sama (PML/PPL yang sama harus punya 1 nomor kontrak).
        nomor_kontrak_pml: cleanText(r.nomor_kontrak_pml),
        nomor_kontrak_ppl: cleanText(r.nomor_kontrak_ppl),

        // 🔥 FIX: pakai SET untuk UNIQUE SLS
        slsSet: new Set(),
      });
    }

    const item = map.get(keyBase);

    const slsKey = makeSlsKey(r);
    item.slsSet.add(slsKey);
  }

  return [...map.values()]
    .map(v => ({
      ...v,
      jumlah: v.slsSet.size, // 🔥 FIX: UNIQUE SLS/SUBSLS
    }))
    .sort((a, b) => {
      // PML: urutkan per nama PPL (alfabetis) supaya tidak acak sesuai urutan baris sheet.
      // PPL: urutkan per kecamatan lalu kelurahan.
      if (isPml) {
        return cleanText(a.nama_ppl).localeCompare(cleanText(b.nama_ppl), "id-ID", { sensitivity: "base" });
      }
      const kecDiff = cleanText(a.kecamatan).localeCompare(cleanText(b.kecamatan), "id-ID", { sensitivity: "base" });
      if (kecDiff !== 0) return kecDiff;
      return cleanText(a.kelurahan).localeCompare(cleanText(b.kelurahan), "id-ID", { sensitivity: "base" });
    });
}

export function buildLampiranTemplateData(formValues, lampiranRows = [], jenis = "PML") {
  // Fungsi ini menerima rows yang SUDAH dipisah per orang oleh generateLampiran().
  // Jadi output DOCX akan berisi satu petugas saja:
  // - PML: satu file per PENGAWAS
  // - PPL: satu file per PENCACAH
  const isPml = upperText(jenis) === "PML";
  const namaPetugas = isPml
    ? cleanText(lampiranRows?.[0]?.nama_pml)
    : cleanText(lampiranRows?.[0]?.nama_ppl);

  const grouped = groupLampiranRows(lampiranRows || [], jenis);

  // 🔥 BARU: variabel "sekali saja" untuk ditaruh di ATAS surat (di luar blok
  // {#peserta}...{/peserta}), bukan per baris tabel.
  // - nomor_kontrak: nomor kontrak orang ini (sama untuk semua barisnya, ambil satu saja
  //   dari baris pertama).
  // - total_jumlah: total SEMUA SLS/Sub-SLS milik orang ini, digabung dari semua
  //   kecamatan/kelurahan yang dia kerjakan (bukan cuma satu baris).
  // - total_jumlah_40 / total_jumlah_60: pembagian 40%/60% dari total_jumlah, dengan
  //   total_jumlah_60 = total_jumlah - total_jumlah_40 (supaya jumlahnya pas, tidak ada
  //   selisih pembulatan).
  const nomorKontrakDokumen = cleanText(
    isPml ? lampiranRows?.[0]?.nomor_kontrak_pml : lampiranRows?.[0]?.nomor_kontrak_ppl
  );
  const totalJumlahDokumen = grouped.reduce((sum, r) => sum + (r.jumlah || 0), 0);
  const totalJumlah40Dokumen = Math.ceil(totalJumlahDokumen * 0.4);
  const totalJumlah60Dokumen = totalJumlahDokumen - totalJumlah40Dokumen;

  return {
    jenis_lampiran: jenis,
    nama_petugas: namaPetugas,
    nama_pml: isPml ? namaPetugas : cleanText(lampiranRows?.[0]?.nama_pml),
    nama_ppl: !isPml ? namaPetugas : cleanText(lampiranRows?.[0]?.nama_ppl),
    tempat: formValues.tempat || formValues.hotel || "",
    hotel: formValues.hotel || formValues.tempat || "",
    gelombang: formValues.gelombang || "",
    kelas: formValues.kelas || "",
    jumlah_baris: grouped.length,

    // Variabel "sekali saja" di atas surat:
    nomor_kontrak: nomorKontrakDokumen,
    total_jumlah: totalJumlahDokumen,
    total_jumlah_40: totalJumlah40Dokumen,
    total_jumlah_60: totalJumlah60Dokumen,

    peserta: grouped.map((r, idx) => {
      // 🔥 FIX BARU: variabel tambahan untuk template.
      // - nomor_kontrak: ambil dari "No Kontrak PML" kalau jenisnya PML, atau
      //   "No Kontrak PPL" kalau jenisnya PPL.
      // - jumlah_40 / jumlah_60: pembagian 40%/60% dari jumlah SLS/Sub-SLS baris ini.
      //   jumlah_60 dihitung sebagai (total - jumlah_40), BUKAN dibulatkan sendiri-sendiri,
      //   supaya jumlah_40 + jumlah_60 selalu pas balik ke jumlah total (tidak ada selisih
      //   pembulatan kalau dijumlahkan manual).
      const totalJumlah = r.jumlah || 0;
      const jumlah40 = Math.ceil(totalJumlah * 0.4);
      const jumlah60 = totalJumlah - jumlah40;
      const nomorKontrak = isPml ? (r.nomor_kontrak_pml || "") : (r.nomor_kontrak_ppl || "");

      return {
        no: idx + 1,

        // Kompatibel dengan template lama:
        // - {nama_petugas} untuk nama utama
        // - {nama_pml} untuk PENGAWAS
        // - {nama_ppl} untuk PENCACAH
        nama_petugas: namaPetugas,
        nama_pml: isPml ? namaPetugas : r.nama_pml || "",
        nama_ppl: !isPml ? namaPetugas : r.nama_ppl || "",

        kecamatan: formatKodeNama(r.kdkec, r.kecamatan),
        kelurahan: formatKodeNama(r.kddesa, r.kelurahan),
        sls: r.sls || "",
        subsls: r.subsls || "",
        jumlah: totalJumlah,

        nomor_kontrak: nomorKontrak,
        jumlah_40: jumlah40,
        jumlah_60: jumlah60,
      };
    }),
  };
}

export async function createLampiranBlob(templateUrl, formValues, lampiranRows, jenis) {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template lampiran: ${response.status} ${response.statusText}`);

  const arrayBuffer = await response.arrayBuffer();
  return createLampiranBlobFromTemplateBuffer(arrayBuffer, formValues, lampiranRows, jenis);
}

// 🔥 FIX: dipisah dari createLampiranBlob() supaya template HANYA di-fetch SEKALI
// (lewat generateLampiran()), bukan di-fetch ulang dari network untuk SETIAP petugas.
// Untuk ratusan/ribuan petugas, fetch berulang ini sangat lambat dan boros memori.
// PizZip dibuat baru dari arrayBuffer yang sama setiap kali dipanggil — ini sesuai
// rekomendasi docxtemplater untuk batch generation (instance Docxtemplater TIDAK
// boleh dipakai ulang untuk render() berkali-kali, tapi arrayBuffer template-nya aman
// dipakai ulang berkali-kali untuk membuat PizZip baru).
export function createLampiranBlobFromTemplateBuffer(templateArrayBuffer, formValues, lampiranRows, jenis) {
  const zip = new PizZip(templateArrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

  doc.render(buildLampiranTemplateData(formValues || {}, lampiranRows || [], jenis));

  return doc.getZip().generate({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

export function sanitizeFileName(value) {
  return cleanText(value || "Tanpa Nama")
    .replace(/[/:*?"<>|]/g, "-")
    .replace(/\s+/g, " ")
    .trim();
}

export async function downloadMultipleAsZip(files, zipName = "dokumen.zip") {
  const zip = new JSZip();

  for (const file of files) {
    // file: { name, blob }
    zip.file(file.name, file.blob);
  }

  const content = await zip.generateAsync({ type: "blob" });
  saveAs(content, zipName);
}



// 🔥 FIX: jumlah PPL/PML bisa mencapai ribuan. Kalau semua dokumen ditahan di memori
// lalu di-zip jadi SATU file raksasa sekaligus, browser bisa kehabisan memori
// ("Array buffer allocation failed"). Maka proses zip dipecah per-batch — setiap
// batch jadi satu file .zip terpisah, sehingga beban memori di setiap tahap kecil
// dan terkendali. Bisa diturunkan lagi (misal 50-100) kalau template lampiran-nya berat.
export const LAMPIRAN_ZIP_BATCH_SIZE = 150;

export async function generateLampiran(templateUrl, formValues, lampiranRows, jenis) {
  const sourceRows = lampiranRows || [];
  const isPml = upperText(jenis) === "PML";

  // 🔥 FIX: jangan kelompokkan HANYA berdasarkan teks nama. Kalau ada 2 petugas
  // berbeda dengan nama yang sama persis (nama kembar), pengelompokan by-nama-saja
  // akan keliru menyatukan data SLS milik 2 orang berbeda jadi 1 dokumen — salah
  // satu identitas aslinya akan "tertelan" oleh yang lain.
  //
  // Solusinya: kunci pengelompokan diutamakan pakai EMAIL (kolom Email Pengawas /
  // Email Pencacah), karena email jauh lebih unik per-orang dibanding nama. Nama
  // tetap dipakai untuk ditampilkan di dokumen & sebagai dasar nama file. Kalau email
  // kosong di data, fallback ke nama saja (risiko nama kembar tetap ada untuk kasus
  // ini, tapi sudah jauh lebih baik daripada selalu mengandalkan nama).
  const groups = new Map(); // identity -> { displayName, rows: [] }

  for (const row of sourceRows) {
    const namaPml = cleanText(row.nama_pml || row.pengawas || row.nama || "");
    const namaPpl = cleanText(row.nama_ppl || row.pencacah || row.nama || "");
    const emailPml = cleanText(row.email_pengawas || "");
    const emailPpl = cleanText(row.email_pencacah || "");

    const displayName = isPml ? namaPml : namaPpl;
    if (!displayName) continue;

    const emailKey = upperText(isPml ? emailPml : emailPpl);
    const identity = emailKey || `NAMA::${upperText(displayName)}`;

    if (!groups.has(identity)) groups.set(identity, { displayName, rows: [] });
    groups.get(identity).rows.push(row);
  }

  if (groups.size === 0) {
    throw new Error(`Tidak ada data ${jenis}`);
  }

  // 🔥 FIX: deteksi nama kembar (identity berbeda tapi displayName sama persis) supaya
  // nama file tidak saling menimpa. Kalau ketemu, file dibedakan jadi "Nama (1).docx",
  // "Nama (2).docx", dst sesuai urutan kemunculan di data.
  const nameOccurrences = new Map();
  for (const { displayName } of groups.values()) {
    const key = upperText(displayName);
    nameOccurrences.set(key, (nameOccurrences.get(key) || 0) + 1);
  }
  const nameRunningIndex = new Map();
  const buildFileBaseName = (displayName) => {
    const key = upperText(displayName);
    const total = nameOccurrences.get(key) || 1;
    if (total <= 1) return sanitizeFileName(displayName);
    const idx = (nameRunningIndex.get(key) || 0) + 1;
    nameRunningIndex.set(key, idx);
    console.warn(`Lampiran ${jenis}: nama kembar terdeteksi -> "${displayName}" (salinan ke-${idx} dari ${total}, dibedakan via email).`);
    return `${sanitizeFileName(displayName)} (${idx})`;
  };

  // 🔥 FIX: fetch template SEKALI saja di sini, lalu arrayBuffer-nya dipakai ulang
  // untuk membuat setiap dokumen. Sebelumnya template di-fetch dari network untuk
  // SETIAP petugas (bisa 1000+ kali fetch untuk file yang sama) — sangat lambat.
  const templateResponse = await fetch(templateUrl);
  if (!templateResponse.ok) {
    throw new Error(`Gagal memuat template lampiran: ${templateResponse.status} ${templateResponse.statusText}`);
  }
  const templateArrayBuffer = await templateResponse.arrayBuffer();

  const entries = [...groups.values()]; // [{ displayName, rows }, ...]
  const totalBatches = Math.ceil(entries.length / LAMPIRAN_ZIP_BATCH_SIZE);

  console.log(
    `Lampiran ${jenis}: ${entries.length} petugas ditemukan, dipecah jadi ${totalBatches} file zip ` +
    `(maks ${LAMPIRAN_ZIP_BATCH_SIZE} dokumen/zip) untuk menghindari kehabisan memori.`
  );

  for (let batchIndex = 0; batchIndex < totalBatches; batchIndex++) {
    const batchEntries = entries.slice(
      batchIndex * LAMPIRAN_ZIP_BATCH_SIZE,
      (batchIndex + 1) * LAMPIRAN_ZIP_BATCH_SIZE
    );

    const zipFiles = [];
    for (const { displayName, rows } of batchEntries) {
      const blob = createLampiranBlobFromTemplateBuffer(templateArrayBuffer, formValues || {}, rows, jenis);
      zipFiles.push({
        name: `Lampiran ${jenis} - ${buildFileBaseName(displayName)}.docx`,
        blob,
      });
    }

    const batchSuffix = totalBatches > 1 ? ` - Bagian ${batchIndex + 1} dari ${totalBatches}` : "";
    await downloadMultipleAsZip(
      zipFiles,
      `Lampiran ${jenis} - ${formValues?.gelombang || "SE2026"}${batchSuffix}.zip`
    );
  }
}

// Generate single lampiran (one person) — fetch template once and render one doc
export async function generateSingleLampiran(templateUrl, formValues, lampiranRows, jenis, displayName) {
  if (!lampiranRows || lampiranRows.length === 0) throw new Error("Tidak ada data untuk lampiran yang dipilih");
  const templateResponse = await fetch(templateUrl);
  if (!templateResponse.ok) throw new Error(`Gagal memuat template lampiran: ${templateResponse.status} ${templateResponse.statusText}`);
  const templateArrayBuffer = await templateResponse.arrayBuffer();
  const blob = createLampiranBlobFromTemplateBuffer(templateArrayBuffer, formValues || {}, lampiranRows, jenis);
  const safeName = sanitizeFileName(displayName || (jenis + "-lampiran"));
  saveAs(blob, `Lampiran ${jenis} - ${safeName}.docx`);
}
