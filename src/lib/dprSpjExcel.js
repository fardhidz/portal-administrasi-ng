// ============================================================
// Portal Administrasi SE2026 — bagian: dprSpjExcel
// Mereplikasi ISI LENGKAP template docx DPR & SPJ (bukan cuma tabel
// peserta) ke Excel: judul, paragraf pernyataan, tabel rincian biaya,
// blok tanda tangan — semuanya tetap ada. Tabel peserta yang panjang
// dipecah per halaman → per sheet, sisanya (surat + tanda tangan)
// hanya muncul di sheet pertama (DPR) / diulang di tiap sheet (SPJ,
// karena tabelnya menyatu dengan info kegiatan).
// ============================================================

import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { drawKopSurat, drawLabelValue, drawParagraph, chunkRows } from "./xlsxKopSurat";

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
// ============================================================
export async function generateDprExcel(data, fileName, rowsPerPage = 15) {
  const workbook = await newWorkbook();
  const totalKolom = 6;
  const pages = chunkRows(data.peserta, rowsPerPage);

  for (let i = 0; i < pages.length; i += 1) {
    const isFirstPage = i === 0;
    const isLastPage = i === pages.length - 1;
    const ws = workbook.addWorksheet(`Halaman ${i + 1}`);
    ws.columns = [{ width: 5 }, { width: 26 }, { width: 20 }, { width: 20 }, { width: 20 }, { width: 20 }];

    let row = await drawKopSurat(ws, { judul: "DAFTAR PENGELUARAN RIIL", totalKolom });

    if (isFirstPage) {
      row = drawParagraph(ws, row, "Yang bertanda tangan di bawah ini :", { totalKolom });
      row = drawLabelValue(ws, row, "Nama", "(terlampir)", { totalKolom });
      row = drawLabelValue(ws, row, "NIP/NIK", "(terlampir)", { totalKolom });
      row = drawLabelValue(ws, row, "Pangkat/Golongan", "(terlampir)", { totalKolom });
      row += 1;
      row = drawParagraph(ws, row, "Berdasarkan surat Perjalanan Dinas (SPD)", { totalKolom });
      row = drawLabelValue(ws, row, "Tanggal", "19 Januari 2026", { totalKolom });
      row = drawLabelValue(ws, row, "Nomor", "001/539184-92800/TRANSLOK-2903/01/2026", { totalKolom });
      row += 1;
      row = drawParagraph(
        ws, row,
        "Biaya transport pegawai dan/atau biaya penginapan selama 3 hari efektif pelaksanaan pelatihan SE2026 di bawah ini yang tidak diperoleh bukti-bukti pengeluarannya meliputi :",
        { totalKolom }
      );
      row += 1;

      // Tabel rincian biaya
      const biayaHeader = ["No", "Uraian", "Jumlah", "Keterangan"];
      biayaHeader.forEach((h, idx) => styledCell(ws, row, idx + 1, h, { bold: true, align: "center", fill: "FFE8E8E8" }));
      row += 1;
      styledCell(ws, row, 1, "1", { align: "center" });
      styledCell(ws, row, 2, "Biaya Transportasi");
      styledCell(ws, row, 3, `Rp ${data.biaya_total}`, { align: "right" });
      styledCell(ws, row, 4, "-", { align: "center" });
      row += 1;
      ws.mergeCells(row, 1, row, 2);
      styledCell(ws, row, 1, "Jumlah", { bold: true });
      styledCell(ws, row, 3, `Rp ${data.biaya_total}`, { bold: true, align: "right" });
      styledCell(ws, row, 4, "");
      row += 1;
      ws.mergeCells(row, 1, row, totalKolom);
      styledCell(ws, row, 1, `Terbilang: ${data.biaya_terbilang}`, { bold: true });
      row += 2;

      row = drawParagraph(
        ws, row,
        "Jumlah uang tersebut pada angka 1 di atas benar-benar dikeluarkan untuk pelaksanaan Perjalanan dinas dimaksud dan apabila di kemudian hari terdapat kelebihan atas pembayaran, kami bersedia untuk menyetorkan kelebihan tersebut ke Kas Negara.",
        { totalKolom }
      );
      row = drawParagraph(ws, row, "Demikian pernyataan ini kami buat dengan sebenarnya, untuk dipergunakan sebagaimana mestinya.", { totalKolom });
      row += 1;

      // Blok tanda tangan (kanan)
      row = drawParagraph(ws, row, `Jakarta, ${data.tanggal_surat}`, { align: "right", totalKolom });
      row = drawParagraph(ws, row, "Mengetahui/menyetujui", { align: "right", totalKolom });
      row = drawParagraph(ws, row, "Pejabat Pembuat Komitmen", { align: "right", totalKolom });
      row += 3;
      row = drawParagraph(ws, row, "Ragil Hermanto", { bold: true, align: "right", totalKolom });
      row = drawParagraph(ws, row, "NIP. 199406212017011001", { align: "right", totalKolom });
      row += 2;
    }

    // Lampiran pelaksana perjalanan dinas (tabel peserta, per halaman)
    row = drawParagraph(ws, row, isFirstPage ? "LAMPIRAN PELAKSANA PERJALANAN DINAS" : "LAMPIRAN PELAKSANA PERJALANAN DINAS (Lanjutan)", { bold: true, align: "center", totalKolom });
    row += 1;

    const lampiranHeader = ["No", "Nama", "NIP/NIK", "Pangkat/Golongan", "Jabatan", "Tanda Tangan"];
    lampiranHeader.forEach((h, idx) => styledCell(ws, row, idx + 1, h, { bold: true, align: "center", fill: "FFE8E8E8" }));
    row += 1;

    pages[i].forEach((p) => {
      styledCell(ws, row, 1, p.no, { align: "center" });
      styledCell(ws, row, 2, p.nama);
      styledCell(ws, row, 3, p.nik);
      styledCell(ws, row, 4, p.pangkat);
      styledCell(ws, row, 5, p.jabatan);
      styledCell(ws, row, 6, "");
      row += 1;
    });

    if (isLastPage) {
      row += 1;
      drawParagraph(ws, row, `Halaman ${i + 1} dari ${pages.length}`, { italic: true, align: "right", totalKolom });
    }
  }

  await saveWorkbook(workbook, fileName);
}

// ============================================================
// SPJ PELATIHAN SE2026
// ============================================================
const SPJ_OK_PER_ORANG = 3;
const SPJ_UANG_PER_OK = 170000;
const SPJ_KOTOR_PER_ORANG = SPJ_UANG_PER_OK * SPJ_OK_PER_ORANG;

export async function generateSpjExcel(data, fileName, rowsPerPage = 20) {
  const workbook = await newWorkbook();
  const totalKolom = 8;
  const pages = chunkRows(data.peserta, rowsPerPage);

  for (let i = 0; i < pages.length; i += 1) {
    const isLastPage = i === pages.length - 1;
    const ws = workbook.addWorksheet(`Halaman ${i + 1}`);
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
    row += 1;

    pages[i].forEach((p) => {
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

    if (isLastPage) {
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
      const kolLabelStart = [1, 4, 6];
      const labelRow = row;
      styledCell(ws, labelRow, 1, "Setuju dibayar :");
      ws.mergeCells(labelRow, 6, labelRow, totalKolom);
      styledCell(ws, labelRow, 6, `Lunas pada tanggal : ${data.tanggal_pelunasan}`);
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
    } else {
      row += 1;
      drawParagraph(ws, row, `Bersambung ke Halaman ${i + 2}...`, { italic: true, align: "right", totalKolom });
    }
  }

  await saveWorkbook(workbook, fileName);
}

// ============================================================
// DAFTAR HADIR PELATIHAN SE2026
// Selalu 1 sheet saja (tidak dipecah per halaman), berapa pun jumlah peserta.
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
