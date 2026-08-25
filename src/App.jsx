// ============================================================
// Portal Administrasi SE2026 — BPS Kota Jakarta Timur
// Dependencies: npm install xlsx docxtemplater pizzip file-saver docx-preview
// ============================================================

import React, { useCallback, useState } from "react";
import { AnimatePresence, motion } from "framer-motion";
import { ArrowLeft, CheckCircle, FileText, LayoutDashboard, Menu, Printer, X } from "lucide-react";

import { DocForm } from "./components/DocForm";
import { generateDaftarHadir, generateDaftarHadirXlsx, generateLampiran, generatePengeluaranRiil, generatePengeluaranRiilXlsx, generateSingleLampiran, generateSpj, generateSpjXlsx } from "./lib/docGenerators";
import { DocPreview } from "./components/DocPreview";
import { DOC_TYPES } from "./data/docTypes";
import { DocCard, GenerateDocxButton, GoogleSheetCard, StatCard, XlsxUploadCard } from "./components/FormPanels";
import { cleanText, upperText } from "./lib/helpers";
import { enrichApproveByPmlWithFotoBukti, loadGoogleSheet, normalizeGoogleSheetUrl, parseApproveByPmlData, parseBappData, parseDataPerSlsData, parseDataPmlProgressData, parseLampiranXlsxData, parseStatusSlsData, parseXlsxData } from "./lib/parsers";
import { DAFTAR_HADIR_TEMPLATE_URL, LAMPIRAN_PML_TEMPLATE_URL, LAMPIRAN_PPL_TEMPLATE_URL, PENGELUARAN_RIIL_TEMPLATE_URL, SPJ_TEMPLATE_URL } from "./data/templates";

// ─── MAIN APP ─────────────────────────────────────────────────────────────────

export default function PortalAdministrasiSE2026() {
  const [view,        setView]        = useState("dashboard");
  const [selectedDoc, setSelectedDoc] = useState(null);
  const [formData,    setFormData]    = useState({});
  const [previewData, setPreviewData] = useState(null);
  const [sidebarOpen, setSidebarOpen] = useState(false);

  const [petugasData,        setPetugasData]        = useState([]);
  const [lampiranData,       setLampiranData]       = useState([]);
  const [bappData,           setBappData]           = useState([]);
  const [statusSlsData,      setStatusSlsData]      = useState([]);
  const [dataPerSlsData,     setDataPerSlsData]     = useState([]);
  const [approveByPmlData,   setApproveByPmlData]   = useState([]);
  const [xlsxLoaded,         setXlsxLoaded]         = useState(false);
  const [xlsxFileName,       setXlsxFileName]       = useState("data-petugas.xlsx");
  const [googleSheetUrl,     setGoogleSheetUrl]     = useState("https://docs.google.com/spreadsheets/d/10jA_NOMNn5pBuy1OPrSdHstscRrUOUlEDElk-jOmXLQ/edit?gid=1095810027#gid=1095810027");
  const [googleSheetApiKey,  setGoogleSheetApiKey]  = useState("");
  const [googleSheetError,   setGoogleSheetError]   = useState(null);
  const [googleSheetLoading, setGoogleSheetLoading] = useState(false);
  const [dataPmlProgressData, setDataPmlProgressData] = useState([]);
  // Lampiran preview controls: pilih jenis (PML/PPL) dan pilih identity (email/name) untuk generate satu-per-orang
  const [lampiranPreviewJenis, setLampiranPreviewJenis] = useState("PML");
  const [lampiranPreviewIdentity, setLampiranPreviewIdentity] = useState("__ALL__");

  // 🔥 BARU: status data cache dari /api/data (hasil sinkronisasi harian
  // via Vercel Cron), supaya beda dengan proses "load langsung dari
  // Google Sheets" yang lama (masih dipertahankan sebagai fallback).
  const [dataSyncedAt, setDataSyncedAt] = useState(null);
  const [syncingNow, setSyncingNow] = useState(false);

  // Terapkan hasil parsing (baik dari cache /api/data, dari sync manual,
  // maupun dari upload xlsx/Google Sheets langsung) ke semua state form
  // secara konsisten di satu tempat.
  const applyLoadedData = useCallback((payload, sourceLabel) => {
    const {
      data = [],
      lampiran = [],
      bappData: loadedBappData = [],
      statusSls: loadedStatusSls = [],
      dataPerSls: loadedDataPerSls = [],
      approveByPml: loadedApproveByPml = [],
      dataPmlProgress: loadedDataPmlProgress = [],
    } = payload || {};

    setPetugasData(data);
    setLampiranData(lampiran);
    setBappData(loadedBappData);
    setStatusSlsData(loadedStatusSls);
    setDataPerSlsData(loadedDataPerSls);
    setApproveByPmlData(loadedApproveByPml);
    setDataPmlProgressData(loadedDataPmlProgress);
    setXlsxLoaded(Boolean(
      data.length || lampiran.length || loadedBappData.length ||
      loadedStatusSls.length || loadedDataPerSls.length || loadedApproveByPml.length
    ));
    setXlsxFileName(
      sourceLabel ||
      `${data.length} petugas, ${lampiran.length} lampiran, ${loadedBappData.length} pembayaran, ${loadedStatusSls.length} status SLS, ${loadedDataPerSls.length} data per SLS, ${loadedApproveByPml.length} approve PML`
    );
  }, []);

  // Coba baca cache cepat dari /api/data (hasil sync harian di server).
  // Kalau belum ada (belum pernah sync, atau lagi dev lokal tanpa Vercel
  // functions), balikan false supaya App.jsx fallback ke cara lama.
  const loadFromApiCache = useCallback(async () => {
    try {
      const response = await fetch("/api/data");
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const payload = await response.json();
      applyLoadedData(payload);
      setDataSyncedAt(payload.syncedAt || null);
      return true;
    } catch (err) {
      console.warn("Cache /api/data belum tersedia, fallback ke sumber lama:", err.message);
      return false;
    }
  }, [applyLoadedData]);

  // Tombol "Sinkronkan Sekarang": trigger /api/sync (fetch ulang Google
  // Sheet + parse di server + simpan ke Blob), lalu baca ulang /api/data.
  // Dipakai kalau Google Sheet baru saja diupdate dan tidak mau menunggu
  // jadwal cron harian.
  const handleSyncNow = async () => {
    const token = window.prompt(
      "Masukkan kode sinkronisasi (CRON_SECRET). Kosongkan kalau belum di-set di project ini:"
    );
    if (token === null) return; // dibatalkan
    setSyncingNow(true);
    try {
      const query = token ? `?token=${encodeURIComponent(token)}` : "";
      const syncRes = await fetch(`/api/sync${query}`);
      const syncJson = await syncRes.json();
      if (!syncRes.ok) throw new Error(syncJson.error || `HTTP ${syncRes.status}`);
      await loadFromApiCache();
      alert(`Sinkronisasi berhasil. Data per ${new Date(syncJson.syncedAt).toLocaleString("id-ID")}.`);
    } catch (err) {
      alert("Gagal sinkronisasi: " + err.message);
    } finally {
      setSyncingNow(false);
    }
  };

  const loadGoogleSheetData = async () => {
    const normalized = normalizeGoogleSheetUrl(googleSheetUrl);
    if (!normalized) { setGoogleSheetError("URL Google Sheets tidak valid."); return; }
    setGoogleSheetError(null);
    setGoogleSheetLoading(true);
    try {
      const {
        data,
        lampiran,
        bappData: loadedBappData,
        statusSls: loadedStatusSls,
        dataPerSls: loadedDataPerSls,
        approveByPml: loadedApproveByPml,
        dataPmlProgress: loadedDataPmlProgress,   // ⬅️ baru
      } = await loadGoogleSheet(normalized, googleSheetApiKey);
      const enrichedApproveByPml = await enrichApproveByPmlWithFotoBukti(loadedApproveByPml || []);
      if (
        data.length === 0 &&
        (!lampiran || lampiran.length === 0) &&
        (!loadedBappData || loadedBappData.length === 0) &&
        (!loadedStatusSls || loadedStatusSls.length === 0) &&
        (!loadedDataPerSls || loadedDataPerSls.length === 0) &&
        (!loadedApproveByPml || loadedApproveByPml.length === 0)
      ) {
        setGoogleSheetError("Tidak ada data petugas, Lampiran, Pembayaran, Status SLS, Data per SLS, maupun Approve by PML yang terbaca.");
        return;
      }

      setPetugasData(data || []);
      setLampiranData(lampiran || []);
      setBappData(loadedBappData || []);
      setStatusSlsData(loadedStatusSls || []);
      setDataPerSlsData(loadedDataPerSls || []);
      setApproveByPmlData(enrichedApproveByPml || []);
      setXlsxLoaded(true);
      setXlsxFileName(`${data.length} petugas, ${lampiran?.length || 0} lampiran, ${loadedBappData?.length || 0} pembayaran, ${loadedStatusSls?.length || 0} status SLS, ${loadedDataPerSls?.length || 0} data per SLS, ${loadedApproveByPml?.length || 0} approve PML`);
    } catch (err) {
      setGoogleSheetError(`Gagal memuat Google Sheet: ${err.message}`);
    } finally {
      setGoogleSheetLoading(false);
    }
  };

  React.useEffect(() => {
    const loadLocal = async () => {
      try {
        const response = await fetch("/data/data-petugas.xlsx");
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const buffer = await response.arrayBuffer();
        const data = parseXlsxData(buffer);
        const lampiran = parseLampiranXlsxData(buffer);
        const bappRows = parseBappData(buffer);
        const statusSlsRows = parseStatusSlsData(buffer);
        const dataPerSlsRows = parseDataPerSlsData(buffer);
        let approveByPmlRows = parseApproveByPmlData(buffer);
        approveByPmlRows = await enrichApproveByPmlWithFotoBukti(approveByPmlRows);
        const dataPmlProgressRows = parseDataPmlProgressData(buffer);
        setDataPmlProgressData(dataPmlProgressRows || []);
        setPetugasData(data || []);
        setLampiranData(lampiran || []);
        setBappData(bappRows || []);
        setStatusSlsData(statusSlsRows || []);
        setDataPerSlsData(dataPerSlsRows || []);
        setApproveByPmlData(approveByPmlRows || []);
        setXlsxLoaded(Boolean(
          lampiran?.length || data?.length || bappRows?.length ||
          statusSlsRows?.length || dataPerSlsRows?.length || approveByPmlRows?.length
        ));
      } catch (err) {
        console.warn("Tidak dapat memuat data-petugas.xlsx:", err.message);
      }
    };

    (async () => {
      // 🔥 BARU: coba cache cepat dulu (/api/data, hasil sync harian).
      // Kalau belum tersedia (misal deploy pertama sebelum /api/sync
      // pernah jalan, atau lagi npm run dev tanpa Vercel functions),
      // baru fallback ke perilaku lama: load langsung dari Google
      // Sheets / file lokal (lambat, tapi tetap berfungsi).
      const gotCache = await loadFromApiCache();
      if (!gotCache) {
        googleSheetUrl ? loadGoogleSheetData() : loadLocal();
      }
    })();
  }, []);

  const openForm = (docType) => {
    if (docType?.disabled) {
      alert(docType.lockedMessage || "Fitur ini dikunci. Saat ini hanya Lampiran yang aktif.");
      return;
    }
    setSelectedDoc(docType);
    setFormData({});
    setPreviewData(null);
    setView("form");
  };
  const handleBack = () => {
    if (view === "preview") { setView("form"); setPreviewData(null); }
    else { setView("dashboard"); setSelectedDoc(null); }
  };
  const handleXlsxUpload = useCallback((file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const buffer = e.target.result;
        const data = parseXlsxData(buffer);
        const lampiran = parseLampiranXlsxData(buffer);
        const bappRows = parseBappData(buffer);
        const statusSlsRows = parseStatusSlsData(buffer);
        const dataPerSlsRows = parseDataPerSlsData(buffer);
        let approveByPmlRows = parseApproveByPmlData(buffer);
        approveByPmlRows = await enrichApproveByPmlWithFotoBukti(approveByPmlRows);
        setPetugasData(data || []);
        setLampiranData(lampiran || []);
        setBappData(bappRows || []);
        setStatusSlsData(statusSlsRows || []);
        setDataPerSlsData(dataPerSlsRows || []);
        setApproveByPmlData(approveByPmlRows || []);
        setXlsxLoaded(Boolean(
          lampiran?.length || data?.length || bappRows?.length ||
          statusSlsRows?.length || dataPerSlsRows?.length || approveByPmlRows?.length
        ));
        setXlsxFileName(file.name);
      } catch (err) { alert("Gagal membaca file xlsx: " + err.message); }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  return (
    <div className="min-h-screen overflow-x-hidden bg-[#fff8f0] text-slate-950 selection:bg-orange-200 selection:text-orange-950">
      <div className="pointer-events-none fixed inset-0 z-0 opacity-70">
        <div className="absolute left-[-12rem] top-[-10rem] h-[28rem] w-[28rem] rounded-full bg-orange-300/40 blur-3xl" />
        <div className="absolute right-[-10rem] top-[12rem] h-[30rem] w-[30rem] rounded-full bg-amber-300/30 blur-3xl" />
        <div className="absolute bottom-[-12rem] left-1/2 h-[30rem] w-[30rem] -translate-x-1/2 rounded-full bg-orange-400/20 blur-3xl" />
      </div>

      <header className="fixed left-0 right-0 top-0 z-50 border-b border-orange-100/80 bg-white/75 backdrop-blur-2xl">
        <div className="mx-auto flex max-w-7xl items-center justify-between px-5 py-3.5 lg:px-8">
          <div className="flex items-center gap-3">
            <button className="rounded-xl border border-orange-100 bg-white p-2 text-slate-700 shadow-sm lg:hidden" onClick={() => setSidebarOpen(!sidebarOpen)}>
              <Menu size={20} />
            </button>
            <div className="flex items-center gap-3">
              <div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-orange-500 shadow-lg shadow-orange-500/25">
                <FileText size={22} className="text-white" />
              </div>
              <div>
                <p className="text-sm font-black leading-none tracking-tight text-slate-950">BPS Kota Jakarta Timur</p>
                <p className="mt-1 text-xs font-semibold text-orange-600">Portal Administrasi SE2026</p>
              </div>
            </div>
          </div>
          <div className="hidden items-center gap-2 lg:flex">
            {xlsxLoaded && (
              <div className="flex items-center gap-2 rounded-full border border-green-200 bg-green-50 px-3 py-1.5">
                <CheckCircle size={14} className="text-green-600" />
                <span className="text-xs font-bold text-green-700">
                  Peserta {petugasData.length || 0} • SLS {lampiranData.length || 0} • Petugas {bappData.length || 0}
                </span>
              </div>
            )}
            {dataSyncedAt && (
              <span className="text-xs font-semibold text-slate-500">
                Sinkron terakhir: {new Date(dataSyncedAt).toLocaleString("id-ID")}
              </span>
            )}
            <button
              type="button"
              onClick={handleSyncNow}
              disabled={syncingNow}
              className="rounded-full border border-orange-200 bg-white px-3 py-1.5 text-xs font-black text-orange-600 shadow-sm transition hover:-translate-y-0.5 hover:bg-orange-50 disabled:cursor-not-allowed disabled:opacity-60"
            >
              {syncingNow ? "Menyinkronkan…" : "Sinkronkan Sekarang"}
            </button>
          </div>
        </div>
      </header>

      <AnimatePresence>
        {sidebarOpen && (
          <>
            <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}
              className="fixed inset-0 z-40 bg-black/30 backdrop-blur-sm lg:hidden" onClick={() => setSidebarOpen(false)} />
            <motion.div initial={{ x: -280 }} animate={{ x: 0 }} exit={{ x: -280 }}
              transition={{ type: "spring", damping: 28, stiffness: 300 }}
              className="fixed left-0 top-0 z-50 flex h-full w-72 flex-col border-r border-orange-100 bg-white shadow-2xl lg:hidden">
              <div className="flex items-center justify-between border-b border-orange-100 p-5">
                <p className="font-black text-slate-950">Menu Dokumen</p>
                <button onClick={() => setSidebarOpen(false)} className="rounded-xl bg-orange-50 p-2 text-orange-600"><X size={18} /></button>
              </div>
              <div className="flex-1 overflow-y-auto p-4">
                <button onClick={() => { setView("dashboard"); setSelectedDoc(null); setSidebarOpen(false); }}
                  className={`mb-2 flex w-full items-center gap-3 rounded-2xl px-4 py-3 text-sm font-bold transition ${view === "dashboard" && !selectedDoc ? "bg-orange-500 text-white" : "text-slate-700 hover:bg-orange-50"}`}>
                  <LayoutDashboard size={18} /> Dashboard
                </button>
                {DOC_TYPES.map((d) => (
                  <button key={d.id} onClick={() => { if (d.disabled) { alert(d.lockedMessage || "Fitur ini dikunci."); return; } openForm(d); setSidebarOpen(false); }}
                    className={`mb-2 flex w-full items-center gap-3 rounded-2xl px-4 py-3 text-sm font-bold transition ${selectedDoc?.id === d.id ? "bg-orange-500 text-white" : "text-slate-700 hover:bg-orange-50"} ${d.disabled ? "cursor-not-allowed opacity-70" : ""}`}>
                    {React.cloneElement(d.icon, { size: 18 })} {d.label}
                    {d.disabled && <span className="ml-auto rounded-full bg-amber-100 px-2 py-0.5 text-[10px] font-black uppercase tracking-wider text-amber-700">Kunci</span>}
                  </button>
                ))}
              </div>
            </motion.div>
          </>
        )}
      </AnimatePresence>

      <main className="relative z-10 pt-20">
        <AnimatePresence mode="wait">
          {view === "dashboard" && (
            <motion.div key="dashboard" initial={{ opacity: 0, y: 18 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }} transition={{ duration: 0.4 }} className="px-5 py-10 lg:px-8">
              <div className="mx-auto max-w-7xl">
                <div className="mb-10">
                  <p className="text-sm font-black uppercase tracking-[0.25em] text-orange-600">Sensus Ekonomi 2026</p>
                  <h1 className="mt-3 text-5xl font-black tracking-[-0.04em] text-slate-950 lg:text-6xl">
                    Portal Administrasi<br />
                    <span className="bg-gradient-to-r from-orange-600 via-orange-500 to-amber-500 bg-clip-text text-transparent">Pelatihan Petugas</span>
                  </h1>
                  <p className="mt-5 max-w-2xl text-base leading-8 text-slate-600">Terbitkan dokumen administrasi secara cepat dan terstandar.</p>
                </div>
                {!xlsxLoaded && (
                  <>
                    <XlsxUploadCard loaded={xlsxLoaded} fileName={xlsxFileName} petugasCount={petugasData.length} onUpload={handleXlsxUpload} />
                    <GoogleSheetCard url={googleSheetUrl} apiKey={googleSheetApiKey} onUrlChange={setGoogleSheetUrl} onApiKeyChange={setGoogleSheetApiKey} onLoad={loadGoogleSheetData} loading={googleSheetLoading} error={googleSheetError} />
                  </>
                )}
                {xlsxLoaded && (
                  <motion.div initial={{ opacity: 0, y: -10 }} animate={{ opacity: 1, y: 0 }}
                    className="mb-8 flex items-center gap-4 rounded-3xl border border-blue-200 bg-blue-50 px-6 py-4 shadow-sm">
                    <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-blue-500 shadow-lg shadow-blue-500/25">
                      <CheckCircle size={22} className="text-white" />
                    </div>
                    <div>
                      <p className="font-black text-blue-900">Data Siap Digunakan</p>
                      <p className="text-sm font-semibold text-blue-600">Peserta {petugasData.length || 0} • SLS {lampiranData.length || 0} • Petugas {bappData.length || 0}</p>
                    </div>
                  </motion.div>
                )}
                <div className="mb-10 grid grid-cols-3 gap-4">
                  <StatCard value="4" label="Jenis Dokumen Aktif" />
                  <StatCard value="SE2026" label="Kegiatan" />
                  <StatCard value={bappData.length || "—"} label="Petugas Aktif" highlight={bappData.length > 0} />
                </div>
                <div className="grid gap-5 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4">
                  {DOC_TYPES.map((doc, i) => <DocCard key={doc.id} doc={doc} index={i} onSelect={() => openForm(doc)} />)}
                </div>
              </div>
            </motion.div>
          )}

          {view === "form" && selectedDoc && (
            <motion.div key="form" initial={{ opacity: 0, y: 18 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }} transition={{ duration: 0.4 }} className="px-5 py-10 lg:px-8">
              <div className="mx-auto max-w-3xl">
                <button onClick={handleBack} className="mb-6 inline-flex items-center gap-2 rounded-2xl border border-orange-100 bg-white/80 px-4 py-2 text-sm font-bold text-slate-700 shadow-sm backdrop-blur transition hover:border-orange-200 hover:bg-white">
                  <ArrowLeft size={16} /> Kembali
                </button>
                <div className="mb-8 flex items-center gap-4">
                  <div className="flex h-16 w-16 items-center justify-center rounded-3xl bg-orange-500 text-white shadow-xl shadow-orange-500/25">
                    {React.cloneElement(selectedDoc.icon, { size: 30 })}
                  </div>
                  <div>
                    <p className="text-sm font-black uppercase tracking-[0.2em] text-orange-600">Formulir Penerbitan</p>
                    <h2 className="text-3xl font-black tracking-tight text-slate-950">{selectedDoc.label}</h2>
                  </div>
                </div>
                <DocForm key={selectedDoc.id} docType={selectedDoc} formData={formData} setFormData={setFormData} onPreview={(data) => { setPreviewData(data); setView("preview"); }} petugasData={petugasData} lampiranData={lampiranData} bappData={bappData} statusSlsData={statusSlsData} dataPerSlsData={dataPerSlsData} approveByPmlData={approveByPmlData} dataPmlProgressData={dataPmlProgressData} xlsxLoaded={xlsxLoaded} />
              </div>
            </motion.div>
          )}

          {view === "preview" && previewData && (
            <motion.div key="preview" initial={{ opacity: 0, y: 18 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }} transition={{ duration: 0.4 }} className="px-5 py-10 lg:px-8">
              <div className="mx-auto max-w-4xl">
                <div className="mb-6 flex items-center justify-between">
                  <button onClick={handleBack} className="inline-flex items-center gap-2 rounded-2xl border border-orange-100 bg-white/80 px-4 py-2 text-sm font-bold text-slate-700 shadow-sm backdrop-blur transition hover:border-orange-200 hover:bg-white">
                    <ArrowLeft size={16} /> Kembali ke Formulir
                  </button>
                  <div className="flex gap-3">
                    <button onClick={() => window.print()} className="inline-flex items-center gap-2 rounded-2xl border border-orange-200 bg-white px-5 py-2.5 text-sm font-black text-orange-600 shadow-sm transition hover:-translate-y-0.5 hover:bg-orange-50">
                      <Printer size={16} /> Cetak
                    </button>
                    {selectedDoc?.id === "daftar-hadir" && (
                      <div className="flex flex-col gap-1 sm:flex-row sm:items-center">
                        <GenerateDocxButton label="Unduh Word" onGenerate={() => generateDaftarHadir(DAFTAR_HADIR_TEMPLATE_URL, previewData.formValues, previewData.peserta, previewData.namaInda, previewData.selectedFilterGroup)} />
                        <GenerateDocxButton label="Unduh Excel" onGenerate={() => generateDaftarHadirXlsx(previewData.formValues, previewData.peserta, previewData.namaInda, previewData.selectedFilterGroup)} />
                      </div>
                    )}
                    {selectedDoc?.id === "pengeluaran-riil" && (
                      <div className="flex flex-col gap-1 sm:flex-row sm:items-center">
                        <GenerateDocxButton label="Unduh Word" onGenerate={() => generatePengeluaranRiil(PENGELUARAN_RIIL_TEMPLATE_URL, previewData.formValues, previewData.peserta)} />
                        <GenerateDocxButton label="Unduh Excel" onGenerate={() => generatePengeluaranRiilXlsx(previewData.formValues, previewData.peserta)} />
                      </div>
                    )}
                    {selectedDoc?.id === "spj" && (
                      <div className="flex flex-col gap-1 sm:flex-row sm:items-center">
                        <GenerateDocxButton label="Unduh Word" onGenerate={() => generateSpj(SPJ_TEMPLATE_URL, previewData.formValues, previewData.peserta)} />
                        <GenerateDocxButton label="Unduh Excel" onGenerate={() => generateSpjXlsx(previewData.formValues, previewData.peserta)} />
                      </div>
                    )}
                    {selectedDoc?.id === "lampiran" && (
                      <div className="flex items-center gap-3">
                        <div className="flex flex-col sm:flex-row sm:items-center gap-2">
                            <select value={lampiranPreviewJenis} onChange={(e) => { setLampiranPreviewJenis(e.target.value); setLampiranPreviewIdentity("__ALL__"); }}
                              className="w-full sm:w-40 rounded-2xl border border-orange-100 bg-white/80 px-3 py-2 text-sm font-semibold text-slate-800 shadow-sm outline-none">
                            <option value="PML">PML</option>
                            <option value="PPL">PPL</option>
                          </select>

                            <select value={lampiranPreviewIdentity} onChange={(e) => setLampiranPreviewIdentity(e.target.value)}
                              className="w-full sm:w-72 rounded-2xl border border-orange-100 bg-white/80 px-3 py-2 text-sm font-semibold text-slate-800 shadow-sm outline-none">
                            <option value="__ALL__">— Semua —</option>
                            {(() => {
                              const rows = previewData?.lampiranRows || [];
                              const isPml = upperText(lampiranPreviewJenis) === "PML";
                              const seen = new Set();
                              return rows.map((r) => {
                                const displayName = isPml ? cleanText(r.nama_pml) : cleanText(r.nama_ppl);
                                if (!displayName) return null;
                                const email = cleanText(isPml ? r.email_pengawas : r.email_pencacah) || "";
                                const emailKey = upperText(email);
                                const identity = emailKey || `NAMA::${upperText(displayName)}`;
                                if (seen.has(identity)) return null;
                                seen.add(identity);
                                return <option key={identity} value={identity}>{displayName}</option>;
                              });
                            })()}
                          </select>

                            <button onClick={async () => {
                            try {
                              if (!previewData?.lampiranRows) throw new Error("Data lampiran belum tersedia");
                              const rows = previewData.lampiranRows || [];
                              const isPml = upperText(lampiranPreviewJenis) === "PML";
                              if (lampiranPreviewIdentity === "__ALL__") {
                                // generate all for this jenis as zip
                                await generateLampiran(isPml ? LAMPIRAN_PML_TEMPLATE_URL : LAMPIRAN_PPL_TEMPLATE_URL, previewData.formValues, rows, lampiranPreviewJenis);
                                return;
                              }
                              const filtered = [];
                              for (const r of rows) {
                                const displayName = isPml ? cleanText(r.nama_pml) : cleanText(r.nama_ppl);
                                const email = cleanText(isPml ? r.email_pengawas : r.email_pencacah) || "";
                                const emailKey = upperText(email);
                                const identity = emailKey || `NAMA::${upperText(displayName)}`;
                                if (identity === lampiranPreviewIdentity) filtered.push(r);
                              }
                              if (filtered.length === 0) throw new Error("Tidak ada data untuk identity yang dipilih");
                              const displayName = (isPml ? filtered[0].nama_pml : filtered[0].nama_ppl) || "Tanpa Nama";
                              await generateSingleLampiran(isPml ? LAMPIRAN_PML_TEMPLATE_URL : LAMPIRAN_PPL_TEMPLATE_URL, previewData.formValues, filtered, lampiranPreviewJenis, displayName);
                            } catch (err) { alert(err.message || err); }
                            }} type="button" className="w-full sm:w-auto inline-flex items-center justify-center gap-2 rounded-2xl bg-orange-500 px-4 py-2 text-sm font-black text-white shadow transition hover:bg-orange-600">Generate Terpilih</button>
                          </div>

                        <div className="flex flex-col sm:flex-row sm:items-center gap-2">
                          <GenerateDocxButton label="Unduh Semua Lampiran PML" onGenerate={() => generateLampiran(LAMPIRAN_PML_TEMPLATE_URL, previewData.formValues, previewData.lampiranRows, "PML")} />
                          <GenerateDocxButton label="Unduh Semua Lampiran PPL" onGenerate={() => generateLampiran(LAMPIRAN_PPL_TEMPLATE_URL, previewData.formValues, previewData.lampiranRows, "PPL")} />
                        </div>
                      </div>
                    )}
                  </div>
                </div>
                <DocPreview docType={selectedDoc} data={previewData} />
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      <footer className="relative z-10 mt-10 border-t border-orange-100 bg-white/90 backdrop-blur-xl">
        <div className="mx-auto max-w-7xl px-5 py-6 lg:px-8">
          <div className="flex flex-col items-center justify-between gap-2 text-center text-xs font-semibold text-slate-500 sm:flex-row sm:text-left">
            <p>© 2026 BPS Kota Jakarta Timur — Portal Administrasi SE2026</p>
            <p>Sistem Penerbitan Dokumen Pelatihan Petugas</p>
          </div>
        </div>
      </footer>
    </div>
  );
}

