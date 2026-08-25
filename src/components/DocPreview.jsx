// ============================================================
// Portal Administrasi SE2026 — bagian: docPreview
// ============================================================

import { DaftarHadirDocxPreview, LampiranDocxPreview, PengeluaranRiilDocxPreview, SpdDocxPreview, SpjDocxPreview, SuratPernyataanKendaraanDocxPreview, SuratTugasDocxPreview, TandaTerimaDocxPreview } from "./DocxPreviews";

// ─── DOC PREVIEW (dispatcher) ─────────────────────────────────────────────────

export function DocPreview({ docType, data }) {
  const renderDoc = () => {
    switch (docType.id) {
      case "daftar-hadir":            return <DaftarHadirDocxPreview formValues={data.formValues || {}} peserta={data.peserta || []} namaInda={data.namaInda || ""} selectedFilterGroup={data.selectedFilterGroup || ""} />;
      case "tanda-terima":            return <TandaTerimaDocxPreview formValues={data.formValues || {}} peserta={data.peserta || []} tandaTerimaType={data.tandaTerimaType || "pelatihan"} />;
      case "surat-pernyataan-kendaraan": return <SuratPernyataanKendaraanDocxPreview formValues={data.formValues || {}} peserta={data.peserta || []} />;
      case "pengeluaran-riil":        return <PengeluaranRiilDocxPreview formValues={data.formValues || {}} peserta={data.peserta || []} />;
      case "spj":                     return <SpjDocxPreview formValues={data.formValues || {}} peserta={data.peserta || []} />;
      case "spd":                     return <SpdDocxPreview formValues={data.formValues || data} peserta={data.peserta || []} />;
      case "surat-tugas":             return <SuratTugasDocxPreview formValues={data.formValues || {}} peserta={data.peserta || []} />;
      case "lampiran":                return <LampiranDocxPreview formValues={data.formValues || {}} lampiranRows={data.lampiranRows || []} />;
      default:                        return <p className="text-sm text-slate-500">Dokumen tidak dikenali.</p>;
    }
  };
  return (
    <div className="overflow-hidden rounded-[2.5rem] border border-orange-100 bg-white/80 shadow-xl shadow-orange-900/5 backdrop-blur">
      <div className="border-b border-orange-100 bg-orange-500 px-8 py-4">
        <div className="flex items-center gap-3">
          <div className="flex gap-1.5">
            <div className="h-3 w-3 rounded-full bg-white/40" /><div className="h-3 w-3 rounded-full bg-white/40" /><div className="h-3 w-3 rounded-full bg-white/40" />
          </div>
          <p className="text-sm font-black text-white">Pratinjau: {docType.label}</p>
        </div>
      </div>
      <div className="p-8 md:p-12 print:p-0">
        <div className="mx-auto w-full max-w-none font-serif text-slate-900">{renderDoc()}</div>
      </div>
    </div>
  );
}
