// ============================================================
// Portal Administrasi SE2026 — bagian: dprSpjExcel
// Mereplikasi ISI LENGKAP template docx DPR & SPJ (bukan cuma tabel
// peserta) ke Excel: judul, paragraf pernyataan, tabel rincian biaya,
// blok tanda tangan — semuanya tetap ada.
//
// PERUBAHAN vs versi lama:
// - DPR: tetap 2 sheet tetap ("Halaman 1" dan "Halaman 2"), TIDAK
//   bertambah lagi mengikuti jumlah peserta:
//     - Halaman 1 = surat pernyataan + tabel biaya + tanda tangan SAJA.
//     - Halaman 2 = lampiran peserta, SEMUA peserta dalam satu sheet
//                   (tidak dipecah per rowsPerPage lagi).
// - SPJ dan Daftar Hadir: tidak dipecah sama sekali, satu sheet saja,
//   tabel mengalir apa adanya (Excel auto-paginate sendiri saat dicetak).
// ============================================================

import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { drawKopSurat, drawLabelValue, drawParagraph } from "./xlsxKopSurat";

const THIN_BORDER = { style: "thin", color: { argb: "FF000000" } };
const ALL_BORDERS = { top: THIN_BORDER, left: THIN_BORDER, bottom: THIN_BORDER, right: THIN_BORDER };

function styledCell(worksheet, r, c, value, { bold = false, align = "left", fill } = {}) {
  const cell = worksheet.getCell(r, c);
  cell.value = value;
  cell.border = ALL_BORDERS;
  cell.font = { bold };
  cell.alignment = { horizontal: align, vertical: "middle", wrapText: true };
  if (fill) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fill } };
  return cell;
}

async function newWorkbook() {
  return new ExcelJS.Workbook();
}

function saveWorkbook(workbook, fileName) {
  return workbook.xlsx.writeBuffer().then((buffer) => {
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    saveAs(blob, fileName);
  });
}

// ============================================================
// DAFTAR PENGELUARAN RIIL (DPR)
// Dua sheet tetap (tidak lebih):
//   - "Halaman 1" : surat pernyataan + tabel biaya + tanda tangan saja
//   - "Halaman 2" : lampiran peserta, SEMUA peserta dalam satu sheet
//                   (tidak dipecah lagi per rowsPerPage)
// ============================================================
const DPR_COLUMNS = [{ width: 5 }, { width: 26 }, { width: 20 }, { width: 20 }, { width: 20 }, { width: 20 }];

export async function generateDprExcel(data, fileName) {
  const workbook = await newWorkbook();
  const totalKolom = 6;

  // ---------------- Halaman 1: surat + biaya + tanda tangan ----------------
  const wsSurat = workbook.addWorksheet("Halaman 1");
  wsSurat.columns = DPR_COLUMNS;

  let row = await drawKopSurat(wsSurat, { judul: "DAFTAR PENGELUARAN RIIL", totalKolom });

  row = drawParagraph(wsSurat, row, "Yang bertanda tangan di bawah ini :", { totalKolom });
  row = drawLabelValue(wsSurat, row, "Nama", "(terlampir)", { totalKolom });
  row = drawLabelValue(wsSurat, row, "NIP/NIK", "(terlampir)", { totalKolom });
  row = drawLabelValue(wsSurat, row, "Pangkat/Golongan", "(terlampir)", { totalKolom });
  row += 1;
  row = drawParagraph(wsSurat, row, "Berdasarkan surat Perjalanan Dinas (SPD)", { totalKolom });
  row = drawLabelValue(wsSurat, row, "Tanggal", "19 Januari 2026", { totalKolom });
  row = drawLabelValue(wsSurat, row, "Nomor", "001/539184-92800/TRANSLOK-2903/01/2026", { totalKolom });
  row += 1;
  row = drawParagraph(
    wsSurat, row,
    "Biaya transport pegawai dan/atau biaya penginapan selama 3 hari efektif pelaksanaan pelatihan SE2026 di bawah ini yang tidak diperoleh bukti-bukti pengeluarannya meliputi :",
    { totalKolom }
  );
  row += 1;

  // Tabel rincian biaya
  const biayaHeader = ["No", "Uraian", "Jumlah", "Keterangan"];
  biayaHeader.forEach((h, idx) => styledCell(wsSurat, row, idx + 1, h, { bold: true, align: "center", fill: "FFE8E8E8" }));
  row += 1;
  styledCell(wsSurat, row, 1, "1", { align: "center" });
  styledCell(wsSurat, row, 2, "Biaya Transportasi");
  styledCell(wsSurat, row, 3, `Rp ${data.biaya_total}`, { align: "right" });
  styledCell(wsSurat, row, 4, "-", { align: "center" });
  row += 1;
  wsSurat.mergeCells(row, 1, row, 2);
  styledCell(wsSurat, row, 1, "Jumlah", { bold: true });
  styledCell(wsSurat, row, 3, `Rp ${data.biaya_total}`, { bold: true, align: "right" });
  styledCell(wsSurat, row, 4, "");
  row += 1;
  wsSurat.mergeCells(row, 1, row, totalKolom);
  styledCell(wsSurat, row, 1, `Terbilang: ${data.biaya_terbilang}`, { bold: true });
  row += 2;

  row = drawParagraph(
    wsSurat, row,
    "Jumlah uang tersebut pada angka 1 di atas benar-benar dikeluarkan untuk pelaksanaan Perjalanan dinas dimaksud dan apabila di kemudian hari terdapat kelebihan atas pembayaran, kami bersedia untuk menyetorkan kelebihan tersebut ke Kas Negara.",
    { totalKolom }
  );
  row = drawParagraph(wsSurat, row, "Demikian pernyataan ini kami buat dengan sebenarnya, untuk dipergunakan sebagaimana mestinya.", { totalKolom });
  row += 1;

  // Blok tanda tangan (kanan) — ini akhir dari Halaman 1
  row = drawParagraph(wsSurat, row, `Jakarta, ${data.tanggal_surat}`, { align: "right", totalKolom });
  row = drawParagraph(wsSurat, row, "Mengetahui/menyetujui", { align: "right", totalKolom });
  row = drawParagraph(wsSurat, row, "Pejabat Pembuat Komitmen", { align: "right", totalKolom });
  row += 3;
  row = drawParagraph(wsSurat, row, "Ragil Hermanto", { bold: true, align: "right", totalKolom });
  row = drawParagraph(wsSurat, row, "NIP. 199406212017011001", { align: "right", totalKolom });

  // ---------------- Halaman 2: lampiran peserta, satu sheet, semua peserta ----------------
  const wsLampiran = workbook.addWorksheet("Halaman 2");
  wsLampiran.columns = DPR_COLUMNS;

  let rowLampiran = await drawKopSurat(wsLampiran, { judul: "DAFTAR PENGELUARAN RIIL", totalKolom });

  rowLampiran = drawParagraph(wsLampiran, rowLampiran, "LAMPIRAN PELAKSANA PERJALANAN DINAS", { bold: true, align: "center", totalKolom });
  rowLampiran += 1;

  const lampiranHeader = ["No", "Nama", "NIP/NIK", "Pangkat/Golongan", "Jabatan", "Tanda Tangan"];
  lampiranHeader.forEach((h, idx) => styledCell(wsLampiran, rowLampiran, idx + 1, h, { bold: true, align: "center", fill: "FFE8E8E8" }));
  const lampiranHeaderRow = rowLampiran;
  rowLampiran += 1;

  data.peserta.forEach((p) => {
    styledCell(wsLampiran, rowLampiran, 1, p.no, { align: "center" });
    styledCell(wsLampiran, rowLampiran, 2, p.nama);
    styledCell(wsLampiran, rowLampiran, 3, p.nik);
    styledCell(wsLampiran, rowLampiran, 4, p.pangkat);
    styledCell(wsLampiran, rowLampiran, 5, p.jabatan);
    styledCell(wsLampiran, rowLampiran, 6, "");
    rowLampiran += 1;
  });

  // header tabel diulang otomatis kalau Excel auto-paginate saat dicetak
  wsLampiran.pageSetup.printTitlesRow = `${lampiranHeaderRow}:${lampiranHeaderRow}`;

  await saveWorkbook(workbook, fileName);
}

// ============================================================
// SPJ PELATIHAN SE2026
// Satu worksheet, tanpa page break manual — mengalir apa adanya.
// ============================================================
const SPJ_OK_PER_ORANG = 3;
const SPJ_UANG_PER_OK = 170000;
const SPJ_KOTOR_PER_ORANG = SPJ_UANG_PER_OK * SPJ_OK_PER_ORANG;

export async function generateSpjExcel(data, fileName) {
  const workbook = await newWorkbook();
  const totalKolom = 8;
  const ws = workbook.addWorksheet("SPJ");
  ws.columns = [
    { width: 5 }, { width: 26 }, { width: 10 }, { width: 16 },
    { width: 14 }, { width: 12 }, { width: 14 }, { width: 16 },
  ];

  let row = await drawKopSurat(ws, { judul: "UANG TRANSPORT PESERTA PELATIHAN PETUGAS SENSUS EKONOMI 2026", totalKolom });

  row = drawLabelValue(ws, row, "Satuan kerja", "BPS Kota Jakarta Timur", { totalKolom });
  row = drawLabelValue(ws, row, "Program", "Program Penyediaan dan Pelayanan Informasi Statistik ( 054.01.GG )", { totalKolom });
  row = drawLabelValue(ws, row, "Kode Kegiatan", "Penyediaan dan Pengembangan Statistik Distribusi (2902)", { totalKolom });
  row = drawLabelValue(ws, row, "Output (KRO)", "Data dan Informasi Publik (2902)", { totalKolom });
  row = drawLabelValue(ws, row, "Rincian Output (RO)", "Publikasi/Laporan SENSUS EKONOMI (BMA.006)", { totalKolom });
  row = drawLabelValue(ws, row, "Komponen", "Pelaksanaan SE2026 (530)", { totalKolom });
  row = drawLabelValue(ws, row, "Tanggal/Bulan", `${data.tanggal_awal} s.d ${data.tanggal_akhir}`, { totalKolom });
  row = drawLabelValue(ws, row, "Lokasi", data.tempat, { totalKolom });
  row += 1;

  const header = ["No", "Nama", "Jumlah O-K", "Uang Transport/OK (Rp)", "Jumlah Kotor (Rp)", "PPh Ps 21 (Rp)", "Jumlah Bersih (Rp)", "Tanda Tangan"];
  header.forEach((h, idx) => styledCell(ws, row, idx + 1, h, { bold: true, align: "center", fill: "FFE8E8E8" }));
  const headerRow = row;
  row += 1;

  data.peserta.forEach((p) => {
    styledCell(ws, row, 1, p.no, { align: "center" });
    styledCell(ws, row, 2, p.nama);
    styledCell(ws, row, 3, SPJ_OK_PER_ORANG, { align: "center" });
    styledCell(ws, row, 4, SPJ_UANG_PER_OK.toLocaleString("id-ID"), { align: "right" });
    styledCell(ws, row, 5, SPJ_KOTOR_PER_ORANG.toLocaleString("id-ID"), { align: "right" });
    styledCell(ws, row, 6, "-", { align: "center" });
    styledCell(ws, row, 7, SPJ_KOTOR_PER_ORANG.toLocaleString("id-ID"), { align: "right" });
    styledCell(ws, row, 8, "");
    row += 1;
  });

  ws.mergeCells(row, 1, row, 2);
  styledCell(ws, row, 1, "Jumlah", { bold: true, align: "center" });
  styledCell(ws, row, 3, data.jumlah_ok, { bold: true, align: "center" });
  styledCell(ws, row, 4, data.jumlah_uang, { bold: true, align: "right" });
  styledCell(ws, row, 5, data.total_jumlah_kotor, { bold: true, align: "right" });
  styledCell(ws, row, 6, "-", { bold: true, align: "center" });
  styledCell(ws, row, 7, data.total_jumlah_bersih, { bold: true, align: "right" });
  styledCell(ws, row, 8, "");
  row += 2;

  row = drawParagraph(ws, row, `Terbilang : ${data.total_terbilang}`, { bold: true, totalKolom });
  row += 2;

  // Blok tanda tangan (3 kolom: PPK, Bendahara, Pembuat Daftar)
  styledCell(ws, row, 1, "Setuju dibayar :");
  ws.mergeCells(row, 6, row, totalKolom);
  styledCell(ws, row, 6, `Lunas pada tanggal : ${data.tanggal_pelunasan}`);
  row += 1;

  styledCell(ws, row, 1, "Pejabat Pembuat Komitmen,");
  styledCell(ws, row, 4, "Bendahara Pengeluaran,");
  styledCell(ws, row, 6, "Pembuat Daftar,");
  row += 4;

  styledCell(ws, row, 1, "Ragil Hermanto, SST", { bold: true });
  styledCell(ws, row, 4, "Ade Yotifali, A.Md", { bold: true });
  styledCell(ws, row, 6, "Mujiono", { bold: true });
  row += 1;
  styledCell(ws, row, 1, "NIP. 199406212017011001");
  styledCell(ws, row, 4, "NIP. 198107062011011011");
  styledCell(ws, row, 6, "NIP. 196904161989031003");
  row += 1;

  // header tabel diulang otomatis kalau Excel auto-paginate saat dicetak
  ws.pageSetup.printTitlesRow = `${headerRow}:${headerRow}`;

  await saveWorkbook(workbook, fileName);
}

// ============================================================
// DAFTAR HADIR PELATIHAN SE2026
// Sudah dari dulu 1 sheet saja — tidak berubah.
// ============================================================
export async function generateDaftarHadirExcel(data, fileName) {
  const workbook = await newWorkbook();
  const totalKolom = 5;
  const ws = workbook.addWorksheet("Daftar Hadir");
  ws.columns = [{ width: 5 }, { width: 30 }, { width: 22 }, { width: 26 }, { width: 20 }];

  let row = await drawKopSurat(ws, { judul: "DAFTAR HADIR PESERTA PELATIHAN PETUGAS SENSUS EKONOMI 2026", totalKolom });

  row = drawLabelValue(ws, row, "Hari / Tanggal", data.tanggal_kegiatan, { totalKolom });
  row = drawLabelValue(ws, row, "Jam", `${data.jam_mulai} s.d ${data.jam_selesai} WIB`, { totalKolom });
  row = drawLabelValue(ws, row, "Tempat", data.tempat, { totalKolom });
  row = drawLabelValue(ws, row, "Gelombang/Kelas", `${data.gelombang}/${data.kelas}`, { totalKolom });
  row += 1;

  const header = ["NO", "NAMA", "JABATAN", "WILAYAH TUGAS", "TANDA TANGAN"];
  header.forEach((h, idx) => styledCell(ws, row, idx + 1, h, { bold: true, align: "center", fill: "FFE8E8E8" }));
  row += 1;

  data.peserta.forEach((p) => {
    styledCell(ws, row, 1, p.no, { align: "center" });
    styledCell(ws, row, 2, p.nama);
    styledCell(ws, row, 3, p.jabatan);
    styledCell(ws, row, 4, p.wil_tugas);
    styledCell(ws, row, 5, "");
    row += 1;
  });
  row += 1;

  row = drawParagraph(ws, row, "Mengetahui", { align: "center", totalKolom });
  row = drawParagraph(ws, row, data.keterangan_ttd || "", { align: "center", totalKolom });
  row += 3;
  row = drawParagraph(ws, row, data.nama_inda || "", { bold: true, align: "center", totalKolom });

  await saveWorkbook(workbook, fileName);
}