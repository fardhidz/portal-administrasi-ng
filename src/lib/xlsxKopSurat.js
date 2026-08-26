// ============================================================
// Portal Administrasi SE2026 — bagian: xlsxKopSurat
// Helper bersama untuk export Excel "resmi" (kop surat + border + per
// halaman jadi per sheet), dipakai oleh Daftar Hadir, SPJ, dan DPR.
// Pakai ExcelJS (bukan SheetJS/xlsx) karena butuh insert gambar logo
// dan styling cell (border, merge, bold) yang tidak didukung xlsx versi free.
// ============================================================

import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { KOP_SURAT_LOGO_URL } from "../data/templates";

const KOP_SURAT_LINES = [
  "BADAN PUSAT STATISTIK",
  "KOTA JAKARTA TIMUR",
];

const THIN_BORDER = { style: "thin", color: { argb: "FF000000" } };
const ALL_BORDERS = { top: THIN_BORDER, left: THIN_BORDER, bottom: THIN_BORDER, right: THIN_BORDER };

let cachedLogoBase64 = null;

// ─── LOGO ────────────────────────────────────────────────────────────────────
async function getLogoBase64() {
  if (cachedLogoBase64) return cachedLogoBase64;
  const response = await fetch(KOP_SURAT_LOGO_URL);
  if (!response.ok) return null;
  const buffer = await response.arrayBuffer();
  let binary = "";
  const bytes = new Uint8Array(buffer);
  for (let i = 0; i < bytes.length; i += 1) binary += String.fromCharCode(bytes[i]);
  cachedLogoBase64 = btoa(binary);
  return cachedLogoBase64;
}

// ─── PAGINATION ──────────────────────────────────────────────────────────────
// Pecah array baris menjadi beberapa "halaman" (nantinya 1 halaman = 1 sheet).
export function chunkRows(rows, rowsPerPage) {
  const list = rows || [];
  if (list.length === 0) return [[]];
  const size = Math.max(1, rowsPerPage || 20);
  const chunks = [];
  for (let i = 0; i < list.length; i += size) chunks.push(list.slice(i, i + size));
  return chunks;
}

// ─── KOP SURAT ──────────────────────────────────────────────────────────────
// Menggambar blok kop surat (logo + nama instansi + judul dokumen) di baris
// paling atas sheet. Mengembalikan nomor baris berikutnya yang masih kosong.
export async function drawKopSurat(worksheet, { judul, totalKolom = 6 }) {
  const logoBase64 = await getLogoBase64();
  let row = 1;

  if (logoBase64) {
    const imageId = worksheet.workbook.addImage({ base64: logoBase64, extension: "png" });
    worksheet.addImage(imageId, { tl: { col: 0.15, row: 0.15 }, ext: { width: 46, height: 56 } });
  }

  worksheet.mergeCells(row, 2, row, totalKolom);
  worksheet.getCell(row, 2).value = KOP_SURAT_LINES[0];
  worksheet.getCell(row, 2).font = { bold: true, size: 13 };
  row += 1;

  worksheet.mergeCells(row, 2, row, totalKolom);
  worksheet.getCell(row, 2).value = KOP_SURAT_LINES[1];
  worksheet.getCell(row, 2).font = { bold: true, size: 13 };
  row += 1;

  worksheet.getRow(row).height = 6;
  row += 1;

  worksheet.mergeCells(row, 1, row, totalKolom);
  const garis = worksheet.getCell(row, 1);
  garis.border = { bottom: { style: "medium", color: { argb: "FF000000" } } };
  row += 1;
  row += 1; // jarak setelah garis kop surat

  worksheet.mergeCells(row, 1, row, totalKolom);
  worksheet.getCell(row, 1).value = judul;
  worksheet.getCell(row, 1).font = { bold: true, size: 12 };
  worksheet.getCell(row, 1).alignment = { horizontal: "center" };
  row += 1;
  row += 1; // baris kosong pemisah

  return row;
}

// ─── PARAGRAF BEBAS (isi surat) ─────────────────────────────────────────────
export function drawParagraph(worksheet, startRow, text, { bold = false, align = "left", totalKolom = 6, italic = false } = {}) {
  const row = startRow;
  worksheet.mergeCells(row, 1, row, totalKolom);
  const cell = worksheet.getCell(row, 1);
  cell.value = text;
  cell.font = { bold, italic };
  cell.alignment = { horizontal: align, vertical: "middle", wrapText: true };
  return row + 1;
}

export function drawBlankRow(worksheet, startRow) {
  return startRow + 1;
}

// ─── BLOK LABEL: NILAI (mis. "Nama : (terlampir)") ─────────────────────────
export function drawLabelValue(worksheet, startRow, label, value, { totalKolom = 6 } = {}) {
  const row = startRow;
  worksheet.getCell(row, 1).value = label;
  worksheet.getCell(row, 1).font = { bold: true };
  worksheet.mergeCells(row, 2, row, totalKolom);
  worksheet.getCell(row, 2).value = `: ${value ?? ""}`;
  return row + 1;
}


export function drawInfoLines(worksheet, startRow, infoLines = []) {
  let row = startRow;
  infoLines.forEach(([label, value]) => {
    worksheet.getCell(row, 1).value = label;
    worksheet.getCell(row, 1).font = { bold: true };
    worksheet.getCell(row, 2).value = ": " + (value ?? "");
    row += 1;
  });
  row += 1;
  return row;
}

// ─── TABEL PESERTA ───────────────────────────────────────────────────────────
export function drawTable(worksheet, startRow, columns, rows) {
  let row = startRow;
  columns.forEach((col, idx) => {
    const cell = worksheet.getCell(row, idx + 1);
    cell.value = col.header;
    cell.font = { bold: true };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border = ALL_BORDERS;
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE8E8E8" } };
    if (col.width) worksheet.getColumn(idx + 1).width = col.width;
  });
  row += 1;

  rows.forEach((dataRow) => {
    columns.forEach((col, idx) => {
      const cell = worksheet.getCell(row, idx + 1);
      cell.value = col.value(dataRow);
      cell.border = ALL_BORDERS;
      cell.alignment = { vertical: "middle", wrapText: true, horizontal: col.align || "left" };
    });
    row += 1;
  });

  return row;
}

// ─── FOOTER (total / tanda tangan, biasanya cuma di halaman terakhir) ───────
export function drawFooterLines(worksheet, startRow, footerLines = []) {
  let row = startRow + 1;
  footerLines.forEach((line) => {
    if (line === null) { row += 1; return; }
    const [label, value, boldValue] = line;
    worksheet.getCell(row, 1).value = label;
    worksheet.getCell(row, 1).font = { bold: true };
    if (value !== undefined) {
      worksheet.getCell(row, 2).value = value;
      if (boldValue) worksheet.getCell(row, 2).font = { bold: true };
    }
    row += 1;
  });
  return row;
}

// ─── ORKESTRASI: 1 sheet = 1 halaman, kop surat diulang tiap sheet ─────────
// pages: array of array-of-rows (hasil chunkRows)
export async function buildPaginatedWorkbook({
  fileName,
  sheetLabel = "Halaman",
  judul,
  infoLines = [],
  columns,
  pages,
  footerLines = [],
  colWidths,
}) {
  const workbook = new ExcelJS.Workbook();
  const totalKolom = columns.length;

  for (let i = 0; i < pages.length; i += 1) {
    const worksheet = workbook.addWorksheet(`${sheetLabel} ${i + 1}`);
    if (colWidths) worksheet.columns = colWidths.map((width) => ({ width }));

    let row = await drawKopSurat(worksheet, { judul, totalKolom });
    row = drawInfoLines(worksheet, row, infoLines);
    row = drawTable(worksheet, row, columns, pages[i]);

    const isLastPage = i === pages.length - 1;
    if (isLastPage && footerLines.length > 0) {
      drawFooterLines(worksheet, row, footerLines);
    }
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  saveAs(blob, fileName);
}
