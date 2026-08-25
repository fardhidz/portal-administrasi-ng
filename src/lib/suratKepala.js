// ============================================================
// Portal Administrasi SE2026 — bagian: suratKepala
// ============================================================

import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { saveAs } from "file-saver";

import { applyApproveByPmlToWorkload, buildBerkasPembayaranRecords, buildDataPerSlsWorkloadRows, findDataPmlProgressRow, formatPercentageNumber, getBerkasIdentity, parseDataPerSlsNumber } from "./berkasPembayaran";
import { buildSelectionKeySet, cleanText, formatRupiah, rowMatchesSelection, upperText } from "./helpers";

export function buildSuratKepalaRows(
  selectionRows = [],
  bappData = [],
  dataPerSlsData = [],
  lampiranData = [],
  statusSlsData = [],
  approveByPmlData = [],
  dataPmlProgressData = []   // ⬅️ parameter baru: sheet "Data PML Progress"
) {
  const keySet = buildSelectionKeySet(selectionRows);
  const rows = [];
  const skipped = [];

  const pmlRecords = buildBerkasPembayaranRecords(
    bappData || [],
    lampiranData || [],
    "PML",
    statusSlsData || [],
    dataPerSlsData || [],
    approveByPmlData || []
  );

  for (const sel of selectionRows) {
    const bappRow = (bappData || []).find((r) => rowMatchesSelection(keySet, r.nama, r.email) &&
      (upperText(r.jabatan_raw || r.jabatan) === "PML" || upperText(r.jabatan_raw || r.jabatan) === "PPL") &&
      (cleanText(sel.email) ? upperText(cleanText(sel.email)) === upperText(cleanText(r.email)) : upperText(cleanText(sel.nama)) === upperText(cleanText(r.nama)))
    );

    if (!bappRow) {
      skipped.push({ nama: sel.nama, email: sel.email, alasan: "Tidak ditemukan di sheet Pembayaran (atau jabatan bukan PML/PPL)" });
      continue;
    }

    const jabatan = upperText(bappRow.jabatan_raw || bappRow.jabatan);
    let target = parseDataPerSlsNumber(bappRow.prelist_total) || 0;
    let realisasi = 0;

    if (jabatan === "PML") {
      // 🔥 DISAMAKAN dengan fitur Gabungan Administrasi Pembayaran:
      // realisasi PML diambil dari total kolom "Jumlah Approve PML" pada sheet
      // "Approve by PML" (via applyApproveByPmlToWorkload), BUKAN lagi dari sheet
      // "Data PML Progress". Target tetap dijumlahkan dari sheet "Data per SLS".
      const identity = getBerkasIdentity(cleanText(bappRow.nama), cleanText(bappRow.email));
      const record = pmlRecords.find((r) => r.identity === identity) ||
        pmlRecords.find((r) => upperText(r.displayName) === upperText(cleanText(bappRow.nama)));

      const dataPerSlsRows = record?.dataPerSlsRows || [];
      const approveByPmlRows = record?.approveByPmlRows || [];
      const bebanKerjaBase = buildDataPerSlsWorkloadRows(dataPerSlsRows, "PML");
      const bebanKerja = applyApproveByPmlToWorkload(bebanKerjaBase, approveByPmlRows, "PML");
      const hasApproveData = approveByPmlRows.length > 0;
      const hasDataPerSls = bebanKerja.rows.length > 0;

      if (hasApproveData || hasDataPerSls) {
        // total.target_jumlah = jumlah target dari Data per SLS
        // total.realisasi_jumlah = jumlah "Jumlah Approve PML" dari Approve by PML
        // (lihat applyApproveByPmlToWorkload -> directApproveTotal), sama persis
        // dengan yang dipakai pada Gabungan Administrasi Pembayaran PML.
        target = parseDataPerSlsNumber(bebanKerja.total.target_jumlah) || target;
        realisasi = parseDataPerSlsNumber(bebanKerja.total.realisasi_jumlah) || 0;
      } else {
        // Fallback terakhir kalau PML ini sama sekali tidak punya baris di
        // Data per SLS maupun Approve by PML.
        const progressRow = findDataPmlProgressRow(dataPmlProgressData, bappRow.nama, bappRow.email);
        if (progressRow) {
          realisasi = parseDataPerSlsNumber(progressRow.realisasi_jumlah) || 0;
          const progressTarget = parseDataPerSlsNumber(progressRow.jumlah_target);
          if (progressTarget) target = progressTarget;
        } else {
          realisasi = parseDataPerSlsNumber(bappRow.realisasi_total) || 0;
        }
      }
    } else {
      // PPL: tetap seperti sebelumnya, jumlahkan semua baris Data per SLS miliknya.
      const namaKey = upperText(bappRow.nama);
      const emailLocal = upperText(cleanText(bappRow.email)).split("@")[0];
      const matchedSlsRows = (dataPerSlsData || []).filter((r) => {
        const rowName = upperText(r.nama_ppl || "");
        const rowUsername = upperText(r.username_ppl || "");
        return (namaKey && rowName === namaKey) || (emailLocal && rowUsername === emailLocal);
      });
      if (matchedSlsRows.length === 0) {
        skipped.push({ nama: bappRow.nama, email: bappRow.email, alasan: "PPL tidak ditemukan di sheet Data per SLS" });
        continue;
      }
      realisasi = matchedSlsRows.reduce(
        (sum, r) => sum + (parseDataPerSlsNumber(r.realisasi_dengan_tidak_ditemukan_jumlah) || 0),
        0
      );
    }

    rows.push({
      nama: cleanText(bappRow.nama),
      jabatan,
      target,
      realisasi,
      persentase: target ? (realisasi / target) * 100 : 0,
    });
  }

  const pml = rows.filter((r) => r.jabatan === "PML")
    .sort((a, b) => a.nama.localeCompare(b.nama, "id-ID", { sensitivity: "base" }));
  const ppl = rows.filter((r) => r.jabatan === "PPL")
    .sort((a, b) => a.nama.localeCompare(b.nama, "id-ID", { sensitivity: "base" }));

  return { rows: [...pml, ...ppl], skipped };
}
 
export function buildSuratKepalaTemplateData(rows = []) {
  const totalTarget = rows.reduce((s, r) => s + r.target, 0);
  const totalRealisasi = rows.reduce((s, r) => s + r.realisasi, 0);
  const totalPersentase = totalTarget ? (totalRealisasi / totalTarget) * 100 : 0;

  return {
    peserta: rows.map((r, idx) => ({
      no: idx + 1,                         // ✅ tag {no}
      nama: r.nama,
      nama_petugas: r.nama,                // ✅ tag {nama_petugas}
      jabatan: r.jabatan,
      target_prelist: formatRupiah(r.target),
      realisasi: formatRupiah(r.realisasi),
      persentase: `${formatPercentageNumber(r.persentase, 2)}%`,
    })),
    total_target_prelist: formatRupiah(totalTarget),
    total_realisasi: formatRupiah(totalRealisasi),
    total_persentase: `${formatPercentageNumber(totalPersentase, 2)}%`,
    average_persentase: `${formatPercentageNumber(totalPersentase, 2)}%`,  // ✅ tag {average_persentase}
  };
}
 
export async function generateSuratKepala(templateUrl, rows) {
  if (!rows || rows.length === 0) throw new Error("Tidak ada data untuk Surat Kepala.");
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template Surat Kepala: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(buildSuratKepalaTemplateData(rows));
  const blob = doc.getZip().generate({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
  saveAs(blob, `Surat Pernyataan Kepala BPS - ${rows.length} Petugas.docx`);
}

