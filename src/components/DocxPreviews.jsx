// ============================================================
// Portal Administrasi SE2026 — bagian: docxPreviews
// ============================================================

import React, { useRef, useState } from "react";
import { renderAsync } from "docx-preview";

import { createDaftarHadirBlob, createLampiranBlob, createPengeluaranRiilBlob, createSpdBlob, createSpjBlob, createSuratPernyataanKendaraanBlob, createSuratTugasBlob, createTandaTerimaBlob } from "../lib/docGenerators";
import { DAFTAR_HADIR_TEMPLATE_URL, LAMPIRAN_PML_TEMPLATE_URL, LAMPIRAN_PPL_TEMPLATE_URL, PENGELUARAN_RIIL_TEMPLATE_URL, SPD_LAMPIRAN_TEMPLATE_URL, SPD_TEMPLATE_URL, SPJ_TEMPLATE_URL, SURAT_PERNYATAAN_KENDARAAN_TEMPLATE_URL, SURAT_TUGAS_TEMPLATE_URL, TANDA_TERIMA_LAPANGAN_TEMPLATE_URL, TANDA_TERIMA_TEMPLATE_URL } from "../data/templates";

// ─── DOCX PREVIEW COMPONENTS ─────────────────────────────────────────────────

const docxPreviewStyle = `
  .docx-wrapper { background: transparent !important; padding: 0 !important; }
  .docx-wrapper > section.docx { margin: 0 auto 24px auto !important; box-shadow: 0 20px 45px rgba(15,23,42,0.12) !important; }
  @media print {
    body * { visibility: hidden; }
    .docx-wrapper, .docx-wrapper * { visibility: visible; }
    .docx-wrapper { position: absolute; left: 0; top: 0; width: 100%; }
    .docx-wrapper > section.docx { box-shadow: none !important; margin: 0 !important; }
  }
`;

const RENDER_OPTS = { className: "docx-preview", inWrapper: true, ignoreWidth: false, ignoreHeight: false, ignoreFonts: false, breakPages: true, renderHeaders: true, renderFooters: true, useBase64URL: true };

export function useSingleDocxPreview(createBlob, deps) {
  const containerRef = useRef(null);
  const [loading, setLoading] = useState(true);
  const [error, setError]     = useState("");
  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        setLoading(true); setError("");
        const blob = await createBlob();
        if (cancelled || !containerRef.current) return;
        containerRef.current.innerHTML = "";
        await renderAsync(blob, containerRef.current, null, RENDER_OPTS);
      } catch (err) {
        if (!cancelled) setError(err?.message || "Gagal memuat pratinjau DOCX.");
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, deps);
  return { containerRef, loading, error };
}

export function DocxPreviewShell({ loading, error, templateName, children }) {
  return (
    <div className="rounded-[2rem] border border-orange-100 bg-white p-4 shadow-xl shadow-orange-900/5">
      <style>{docxPreviewStyle}</style>
      {loading && (
        <div className="rounded-2xl border border-orange-100 bg-orange-50/60 p-8 text-center">
          <p className="text-sm font-black text-orange-700">Memuat pratinjau dari template DOCX...</p>
          <p className="mt-1 text-xs font-semibold text-slate-500">Template: {templateName}</p>
        </div>
      )}
      {error && <div className="rounded-2xl border border-red-200 bg-red-50 p-4 text-sm font-bold text-red-600">{error}</div>}
      {children}
    </div>
  );
}

export function DaftarHadirDocxPreview({ formValues, peserta, namaInda, selectedFilterGroup }) {
  const { containerRef, loading, error } = useSingleDocxPreview(
    () => createDaftarHadirBlob(DAFTAR_HADIR_TEMPLATE_URL, formValues || {}, peserta || [], namaInda || "", selectedFilterGroup || ""),
    [formValues, peserta, namaInda, selectedFilterGroup]
  );
  return <DocxPreviewShell loading={loading} error={error} templateName="1. Daftar Hadir Pelatihan SE2026.docx"><div ref={containerRef} className="overflow-x-auto" /></DocxPreviewShell>;
}

export function TandaTerimaDocxPreview({ formValues, peserta, tandaTerimaType }) {
  const templateUrl = (tandaTerimaType === "lapangan" || tandaTerimaType === "mitra-umum") ? TANDA_TERIMA_LAPANGAN_TEMPLATE_URL : TANDA_TERIMA_TEMPLATE_URL;
  const templateName = (tandaTerimaType === "lapangan" || tandaTerimaType === "mitra-umum") ? "2. Tanda Terima Perlengkapan SE2026 - Copy.docx" : "2. Tanda Terima Perlengkapan SE2026.docx";
  const { containerRef, loading, error } = useSingleDocxPreview(
    () => createTandaTerimaBlob(templateUrl, formValues || {}, peserta || []), [formValues, peserta, tandaTerimaType]
  );
  return <DocxPreviewShell loading={loading} error={error} templateName={templateName}><div ref={containerRef} className="overflow-x-auto" /></DocxPreviewShell>;
}

export function SuratPernyataanKendaraanDocxPreview({ formValues, peserta }) {
  const { containerRef, loading, error } = useSingleDocxPreview(
    () => createSuratPernyataanKendaraanBlob(SURAT_PERNYATAAN_KENDARAAN_TEMPLATE_URL, formValues || {}, peserta || []), [formValues, peserta]
  );
  return <DocxPreviewShell loading={loading} error={error} templateName="3. Super Kendis Pelatihan SE2026.docx"><div ref={containerRef} className="overflow-x-auto" /></DocxPreviewShell>;
}

export function PengeluaranRiilDocxPreview({ formValues, peserta }) {
  const { containerRef, loading, error } = useSingleDocxPreview(
    () => createPengeluaranRiilBlob(PENGELUARAN_RIIL_TEMPLATE_URL, formValues || {}, peserta || []), [formValues, peserta]
  );
  return <DocxPreviewShell loading={loading} error={error} templateName="4. DPR_Pelatihan SE 2026.docx"><div ref={containerRef} className="overflow-x-auto" /></DocxPreviewShell>;
}

export function SpjDocxPreview({ formValues, peserta }) {
  const { containerRef, loading, error } = useSingleDocxPreview(
    () => createSpjBlob(SPJ_TEMPLATE_URL, formValues || {}, peserta || []), [formValues, peserta]
  );
  return <DocxPreviewShell loading={loading} error={error} templateName="5. SPJ Pelatihan_SE26.docx"><div ref={containerRef} className="overflow-x-auto" /></DocxPreviewShell>;
}

export function SuratTugasDocxPreview({ formValues, peserta }) {
  const { containerRef, loading, error } = useSingleDocxPreview(
    () => createSuratTugasBlob(SURAT_TUGAS_TEMPLATE_URL, formValues || {}, peserta || []), [formValues, peserta]
  );
  return <DocxPreviewShell loading={loading} error={error} templateName="6. Surat Tugas.docx"><div ref={containerRef} className="overflow-x-auto" /></DocxPreviewShell>;
}

export function SpdDocxPreview({ formValues, peserta }) {
  const mainRef     = useRef(null);
  const lampiranRef = useRef(null);
  const [loading, setLoading]             = useState(true);
  const [error, setError]                 = useState("");
  const [activePreview, setActivePreview] = useState("spd");

  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        setLoading(true); setError("");
        const [blobMain, blobLampiran] = await Promise.all([
          createSpdBlob(SPD_TEMPLATE_URL, formValues || {}, peserta || []),
          createSpdBlob(SPD_LAMPIRAN_TEMPLATE_URL, formValues || {}, peserta || []),
        ]);
        if (cancelled) return;
        if (mainRef.current)     { mainRef.current.innerHTML     = ""; await renderAsync(blobMain,     mainRef.current,     null, RENDER_OPTS); }
        if (lampiranRef.current) { lampiranRef.current.innerHTML = ""; await renderAsync(blobLampiran, lampiranRef.current, null, RENDER_OPTS); }
      } catch (err) {
        if (!cancelled) setError(err?.message || "Gagal memuat pratinjau DOCX.");
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [formValues, peserta]);

  return (
    <div className="space-y-4 rounded-[2rem] border border-orange-100 bg-white p-4 shadow-xl shadow-orange-900/5">
      <style>{docxPreviewStyle}</style>
      {loading && <div className="rounded-2xl border border-orange-100 bg-orange-50/60 p-8 text-center"><p className="text-sm font-black text-orange-700">Memuat pratinjau SPD &amp; Lampiran...</p></div>}
      {error && <div className="rounded-2xl border border-red-200 bg-red-50 p-4 text-sm font-bold text-red-600">{error}</div>}
      <div className="flex gap-3">
        {[{ key: "spd", label: "Preview SPD" }, { key: "lampiran", label: "Preview Lampiran SPD" }].map(({ key, label }) => (
          <button key={key} type="button" onClick={() => setActivePreview(key)}
            className={`rounded-2xl px-5 py-2.5 text-sm font-black transition ${activePreview === key ? "bg-orange-600 text-white" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
            {label}
          </button>
        ))}
      </div>
      <div className={activePreview === "spd" ? "block" : "hidden"}><div ref={mainRef} className="overflow-x-auto rounded-2xl border border-orange-100 bg-white" /></div>
      <div className={activePreview === "lampiran" ? "block" : "hidden"}><div ref={lampiranRef} className="overflow-x-auto rounded-2xl border border-orange-100 bg-white" /></div>
    </div>
  );
}

export function LampiranDocxPreview({ formValues, lampiranRows }) {
  const pmlRef = useRef(null);
  const pplRef = useRef(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [activePreview, setActivePreview] = useState("pml");

  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        setLoading(true); setError("");
        const [blobPml, blobPpl] = await Promise.all([
          createLampiranBlob(LAMPIRAN_PML_TEMPLATE_URL, formValues || {}, lampiranRows || [], "PML"),
          createLampiranBlob(LAMPIRAN_PPL_TEMPLATE_URL, formValues || {}, lampiranRows || [], "PPL"),
        ]);
        if (cancelled) return;
        if (pmlRef.current) { pmlRef.current.innerHTML = ""; await renderAsync(blobPml, pmlRef.current, null, RENDER_OPTS); }
        if (pplRef.current) { pplRef.current.innerHTML = ""; await renderAsync(blobPpl, pplRef.current, null, RENDER_OPTS); }
      } catch (err) {
        if (!cancelled) setError(err?.message || "Gagal memuat pratinjau Lampiran.");
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [formValues, lampiranRows]);

  return (
    <div className="space-y-4 rounded-[2rem] border border-orange-100 bg-white p-4 shadow-xl shadow-orange-900/5">
      <style>{docxPreviewStyle}</style>
      {loading && <div className="rounded-2xl border border-orange-100 bg-orange-50/60 p-8 text-center"><p className="text-sm font-black text-orange-700">Memuat pratinjau Lampiran PML &amp; PPL...</p></div>}
      {error && <div className="rounded-2xl border border-red-200 bg-red-50 p-4 text-sm font-bold text-red-600">{error}</div>}
      {/* Preview-only: removed manual controls (moved to form) */}
      <div className="flex gap-3">
        {[{ key: "pml", label: "Preview Lampiran PML" }, { key: "ppl", label: "Preview Lampiran PPL" }].map(({ key, label }) => (
          <button key={key} type="button" onClick={() => setActivePreview(key)}
            className={`rounded-2xl px-5 py-2.5 text-sm font-black transition ${activePreview === key ? "bg-orange-600 text-white" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
            {label}
          </button>
        ))}
      </div>
      <div className={activePreview === "pml" ? "block" : "hidden"}><div ref={pmlRef} className="overflow-x-auto rounded-2xl border border-orange-100 bg-white" /></div>
      <div className={activePreview === "ppl" ? "block" : "hidden"}><div ref={pplRef} className="overflow-x-auto rounded-2xl border border-orange-100 bg-white" /></div>
    </div>
  );
}

