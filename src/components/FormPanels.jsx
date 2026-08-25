// ============================================================
// Portal Administrasi SE2026 — bagian: formPanels
// ============================================================

import React, { useRef, useState } from "react";
import { motion } from "framer-motion";
import { AlertCircle, CheckCircle, ChevronRight, Download, Filter, MapPin, Upload, Users, X } from "lucide-react";
import { saveAs } from "file-saver";

import { cleanText, parseSelectionXlsx, sortPesertaByJabatanOrder, uniqueSorted, upperText } from "../lib/helpers";
import { DAFTAR_HADIR_GROUPS, jabatanMasukGroup } from "../data/templates";

// ─── XLSX UPLOAD CARD ─────────────────────────────────────────────────────────

export function XlsxUploadCard({ loaded, fileName, petugasCount, onUpload }) {
  const fileRef = useRef(null);
  const [dragging, setDragging] = useState(false);
  const handleDrop = (e) => {
    e.preventDefault(); setDragging(false);
    const file = e.dataTransfer.files[0];
    if (file && file.name.endsWith(".xlsx")) onUpload(file);
  };
  return (
    <div className="mb-8">
      {loaded ? (
        <motion.div initial={{ opacity: 0, scale: 0.97 }} animate={{ opacity: 1, scale: 1 }}
          className="flex items-center justify-between rounded-3xl border border-green-200 bg-green-50 px-6 py-4 shadow-sm">
          <div className="flex items-center gap-4">
            <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-green-500 shadow-lg shadow-green-500/25"><CheckCircle size={22} className="text-white" /></div>
            <div>
              <p className="font-black text-green-800">Data Petugas Berhasil Dimuat</p>
              <p className="text-sm font-semibold text-green-600">{fileName} — {petugasCount} petugas</p>
            </div>
          </div>
          <button onClick={() => fileRef.current?.click()} className="rounded-2xl border border-green-200 bg-white px-4 py-2 text-sm font-bold text-green-700 transition hover:bg-green-50">Ganti File</button>
          <input ref={fileRef} type="file" accept=".xlsx" className="hidden" onChange={(e) => onUpload(e.target.files[0])} />
        </motion.div>
      ) : (
        <div onDragOver={(e) => { e.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={handleDrop} onClick={() => fileRef.current?.click()}
          className={`cursor-pointer rounded-3xl border-2 border-dashed p-8 text-center transition ${dragging ? "border-orange-400 bg-orange-50" : "border-orange-200 bg-white/60 hover:border-orange-300 hover:bg-orange-50/50"}`}>
          <div className="mx-auto mb-4 flex h-14 w-14 items-center justify-center rounded-2xl bg-orange-100"><Upload size={24} className="text-orange-600" /></div>
          <p className="font-black text-slate-800">Unggah Data Petugas (Opsional)</p>
          <p className="mt-1 text-sm font-semibold text-slate-500">Drag & drop file <span className="text-orange-600">data-petugas.xlsx</span> di sini untuk mengganti.</p>
          <input ref={fileRef} type="file" accept=".xlsx" className="hidden" onChange={(e) => onUpload(e.target.files[0])} />
        </div>
      )}
    </div>
  );
}

export function GoogleSheetCard({ url, apiKey, onUrlChange, onApiKeyChange, onLoad, loading, error }) {
  return (
    <motion.div initial={{ opacity: 0, y: -6 }} animate={{ opacity: 1, y: 0 }} className="mb-8 rounded-3xl border border-orange-200 bg-white p-6 shadow-sm">
      <div className="flex items-center gap-4">
        <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-orange-100"><MapPin size={22} className="text-orange-600" /></div>
        <div>
          <p className="font-black text-slate-900">Sedang Load Data, Harap Tunggu!</p>
        </div>
      </div>
      {/* <div className="mt-4 grid gap-3 sm:grid-cols-[1fr_auto]">
        <input value={url} onChange={(e) => onUrlChange(e.target.value)} placeholder="https://docs.google.com/spreadsheets/d/ID_SHEET/edit#gid=0"
          className="w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm text-slate-800 outline-none transition focus:border-orange-300 focus:bg-white" />
        <button type="button" onClick={onLoad} disabled={loading || !url}
          className="rounded-2xl bg-orange-600 px-5 py-3 text-sm font-bold text-white transition enabled:hover:bg-orange-700 disabled:cursor-not-allowed disabled:bg-orange-200">
          {loading ? "Memuat..." : "Muat"}
        </button>
      </div> */}
    </motion.div>
  );
}

// ─── GENERATE DOCX BUTTON ────────────────────────────────────────────────────

export function GenerateDocxButton({ onGenerate, label = "Unduh .docx" }) {
  const [loading, setLoading] = useState(false);
  const [error, setError]     = useState(null);
  const handleGenerate = async () => {
    setLoading(true); setError(null);
    try { await onGenerate(); } catch (err) { setError(err.message); } finally { setLoading(false); }
  };
  return (
    <div className="flex flex-col items-end gap-1">
      <button type="button" onClick={handleGenerate} disabled={loading}
        className="inline-flex items-center gap-2 rounded-2xl bg-orange-500 px-5 py-2.5 text-sm font-black text-white shadow-xl shadow-orange-500/20 transition hover:-translate-y-0.5 hover:bg-orange-600 disabled:opacity-60">
        <Download size={16} />
        {loading ? "Membuat..." : label}
      </button>
      {error && <p className="flex items-center gap-1 text-xs font-semibold text-red-500"><AlertCircle size={12} /> {error}</p>}
    </div>
  );
}

// ─── UNGGAH EXCEL NAMA/EMAIL (fitur "Download Beberapa") ────────────────────
export const SELECTION_TEMPLATE_URL = "/templates/Template Download Beberapa (Nama-Email).xlsx";

export async function downloadSelectionTemplate() {
  const response = await fetch(SELECTION_TEMPLATE_URL);
  if (!response.ok) throw new Error(`Gagal memuat template: ${response.status} ${response.statusText}`);
  const blob = await response.blob();
  saveAs(blob, "Template Download Beberapa (Nama-Email).xlsx");
}

export function SelectionUploadPanel({ selectionRows, onSelectionLoaded, onClear, hint }) {
  const fileRef = useRef(null);
  const [error, setError] = useState("");
  const [templateDownloading, setTemplateDownloading] = useState(false);

  const handleFile = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const rows = parseSelectionXlsx(e.target.result);
        if (rows.length === 0) throw new Error("Tidak ada baris Nama/Email yang terbaca dari file.");
        setError("");
        onSelectionLoaded(rows);
      } catch (err) {
        setError(err.message || "Gagal membaca file Excel.");
      } finally {
        if (fileRef.current) fileRef.current.value = "";
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleDownloadTemplate = async () => {
    setTemplateDownloading(true);
    try {
      await downloadSelectionTemplate();
    } catch (err) {
      setError(err.message || "Gagal mengunduh template.");
    } finally {
      setTemplateDownloading(false);
    }
  };

  return (
    <div className="rounded-2xl border border-dashed border-orange-200 bg-orange-50/50 p-4 space-y-2">
      <div className="flex flex-wrap items-center gap-2">
        <button type="button" onClick={() => fileRef.current?.click()}
          className="inline-flex items-center gap-2 rounded-2xl border border-orange-200 bg-white px-4 py-2 text-sm font-bold text-orange-700 shadow-sm transition hover:bg-orange-50">
          <Upload size={14} /> Unggah Excel Nama/Email
        </button>
        <input ref={fileRef} type="file" accept=".xlsx" className="hidden" onChange={(e) => handleFile(e.target.files[0])} />
        {selectionRows.length > 0 && (
          <span className="inline-flex items-center gap-2 rounded-full border border-green-200 bg-green-50 px-3 py-1 text-xs font-bold text-green-700">
            <CheckCircle size={12} /> {selectionRows.length} baris dimuat
            <button type="button" onClick={onClear} className="ml-1 text-green-700 hover:text-green-900"><X size={12} /></button>
          </span>
        )}
      </div>
      {error && <p className="flex items-center gap-1 text-xs font-semibold text-red-500"><AlertCircle size={12} /> {error}</p>}
      <p className="text-xs font-semibold text-slate-400">
        {hint || "Kolom yang dibaca: Nama dan/atau Email. Baris yang tidak cocok dengan data akan diabaikan."}{" "}
        <button
          type="button"
          onClick={handleDownloadTemplate}
          disabled={templateDownloading}
          className="font-bold text-orange-600 underline underline-offset-2 hover:text-orange-700 disabled:opacity-50"
        >
          {templateDownloading ? "Mengunduh..." : "disini"}
        </button>
      </p>
    </div>
  );
}

// ─── STAT + DOC CARDS ────────────────────────────────────────────────────────

export function StatCard({ value, label, highlight }) {
  return (
    <div className={`rounded-3xl border p-5 shadow-sm backdrop-blur ${highlight ? "border-green-200 bg-green-50/80" : "border-orange-100 bg-white/75"}`}>
      <p className={`text-3xl font-black ${highlight ? "text-green-700" : "text-slate-950"}`}>{value}</p>
      <p className="mt-1 text-xs font-bold uppercase tracking-wider text-slate-500">{label}</p>
    </div>
  );
}

export function DocCard({ doc, index, onSelect }) {
  return (
    <motion.button initial={{ opacity: 0, y: 24 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: 0.5, delay: index * 0.07 }}
      onClick={doc.disabled ? undefined : onSelect}
      disabled={doc.disabled}
      className={`group relative overflow-hidden rounded-[2rem] border border-orange-100 bg-white/85 p-6 text-left shadow-lg shadow-orange-900/5 backdrop-blur transition ${doc.disabled ? "cursor-not-allowed opacity-70" : "hover:-translate-y-2 hover:shadow-2xl hover:shadow-orange-500/15"}`}>
      <div className="absolute right-0 top-0 h-24 w-24 rounded-bl-[4rem] bg-orange-50 transition group-hover:bg-orange-100" />
      <div className="relative flex h-14 w-14 items-center justify-center rounded-2xl bg-orange-500 text-white shadow-xl shadow-orange-500/25">
        {React.cloneElement(doc.icon, { size: 26 })}
      </div>
      <h3 className="relative mt-5 text-lg font-black tracking-tight text-slate-950">{doc.label}</h3>
      <p className="relative mt-2 text-sm leading-6 text-slate-500">{doc.desc}</p>
      <div className="relative mt-5 flex items-center gap-1 text-sm font-black text-orange-500">
        {doc.disabled ? "Terkunci" : "Buat Dokumen"} <ChevronRight size={16} className={`transition ${doc.disabled ? "" : "group-hover:translate-x-1"}`} />
      </div>
    </motion.button>
  );
}

// ─── PESERTA TABLE PREVIEW ────────────────────────────────────────────────────

export function PesertaTablePreview({ peserta }) {
  if (!peserta || peserta.length === 0) return null;
  return (
    <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }}>
      <div className="overflow-hidden rounded-2xl border border-orange-100 bg-white/80">
        <div className="grid grid-cols-[2rem_1fr_1fr_1fr] border-b border-orange-100 bg-orange-500 px-4 py-2 text-xs font-black uppercase text-white">
          <span>No</span><span>Nama</span><span>Jabatan</span><span>Wil. Tugas</span>
        </div>
        <div className="max-h-64 divide-y divide-orange-50 overflow-y-auto">
          {peserta.map((p, i) => (
            <div key={i} className="grid grid-cols-[2rem_1fr_1fr_1fr] px-4 py-2.5 text-xs font-semibold text-slate-700 hover:bg-orange-50/50">
              <span className="text-slate-400">{i + 1}</span>
              <span>{p.nama}</span>
              <span>{p.jabatan}</span>
              <span>{p.wilTugas}</span>
            </div>
          ))}
        </div>
      </div>
    </motion.div>
  );
}

// ─── FILTER PESERTA PANEL ─────────────────────────────────────────────────────

export function FilterPesertaPanel({ xlsxLoaded, formData, setFormData, petugasData, mode = "grouped", selectedGroup = "", onFilterResult, prependRow = null }) {
  const inputCls = "w-full rounded-2xl border border-orange-100 bg-white/80 px-4 py-3 text-sm font-semibold text-slate-800 shadow-sm outline-none focus:border-orange-300 focus:ring-2 focus:ring-orange-100 transition";
  const labelCls = "mb-2 block text-xs font-black uppercase tracking-[0.2em] text-slate-500";

  const [filtered,        setFiltered]        = useState(false);
  const [filteredPeserta, setFilteredPeserta] = useState([]);
  const [namaInda,        setNamaInda]        = useState("");

  const showKelas = selectedGroup !== "panitia-inda";

  const selectedHotel     = cleanText(formData.hotel);
  const selectedKelas     = cleanText(formData.kelas);
  const selectedGelombang = cleanText(formData.gelombang);

  const hotelOptions = React.useMemo(() => uniqueSorted([...petugasData.map((p) => p.hotel), "STIS"]), [petugasData]);
  const gelombangOptions = React.useMemo(() => {
    if (cleanText(selectedHotel).toLowerCase() === "stis") return ["4"];
    return uniqueSorted(petugasData.filter((p) => !selectedHotel || cleanText(p.hotel) === selectedHotel).map((p) => p.gelombang));
  }, [petugasData, selectedHotel]);
  const kelasOptions = React.useMemo(() =>
    uniqueSorted(petugasData
      .filter((p) => !selectedHotel || cleanText(p.hotel) === selectedHotel)
      .filter((p) => !selectedGelombang || cleanText(p.gelombang) === selectedGelombang)
      .map((p) => p.kelas)),
    [petugasData, selectedHotel, selectedGelombang]
  );

  const resetFilterResult = () => {
    setFilteredPeserta([]); setNamaInda(""); setFiltered(false);
    onFilterResult([], "", "");
  };

  const runFilter = () => {
    if (!xlsxLoaded) { alert("Data petugas (.xlsx) belum berhasil dimuat."); return; }
    if (!selectedHotel || !selectedGelombang) { alert("Pilih Tempat dan Gelombang terlebih dahulu."); return; }
    if (showKelas && !selectedKelas) { alert("Pilih Kelas terlebih dahulu."); return; }

    const baseRows = petugasData.filter((p) =>
      cleanText(p.hotel)     === selectedHotel &&
      cleanText(p.gelombang) === selectedGelombang &&
      (showKelas ? cleanText(p.kelas) === selectedKelas : true)
    );

    const inda = baseRows.find((p) => upperText(p.jabatan) === "INDA");
    let hasil;
    if (mode === "all") {
      hasil = sortPesertaByJabatanOrder(baseRows.filter((p) => ["PML", "PPL"].includes(upperText(p.jabatan))));
    } else if (selectedGroup) {
      hasil = sortPesertaByJabatanOrder(baseRows.filter((p) => jabatanMasukGroup(p.jabatan, selectedGroup)));
    } else {
      hasil = sortPesertaByJabatanOrder(baseRows);
    }

    const displayPeserta = hasil.length > 0 && prependRow ? [prependRow, ...hasil] : hasil;
    setNamaInda(inda?.nama || "");
    setFilteredPeserta(displayPeserta);
    setFiltered(true);
    onFilterResult(hasil, inda?.nama || "", selectedGroup);
  };

  return (
    <div className="rounded-3xl border border-orange-100 bg-orange-50/60 p-5 space-y-4">
      <p className="flex items-center gap-2 text-xs font-black uppercase tracking-[0.2em] text-orange-700">
        <Filter size={14} /> Parameter Filter Peserta
      </p>

      <div>
        <label className={labelCls}>Tempat</label>
        <select className={inputCls} value={formData.hotel || ""}
          onChange={(e) => { const hotel = e.target.value; setFormData((prev) => ({ ...prev, hotel, tempat: hotel, kelas: "", gelombang: "" })); resetFilterResult(); }}
          disabled={!xlsxLoaded || hotelOptions.length === 0}>
          <option value="">Pilih tempat</option>
          {hotelOptions.map((h) => <option key={h} value={h}>{h}</option>)}
        </select>
        <p className="mt-1 text-xs font-semibold text-slate-400">Daftar tempat diambil dari kolom TC pada data XLSX.</p>
      </div>

      <div className={`grid gap-4 ${showKelas ? "sm:grid-cols-2" : ""}`}>
        <div>
          <label className={labelCls}>Gelombang</label>
          <select className={inputCls} value={formData.gelombang || ""}
            onChange={(e) => { setFormData((prev) => ({ ...prev, gelombang: e.target.value, kelas: "" })); resetFilterResult(); }}
            disabled={!selectedHotel}>
            <option value="">Pilih gelombang</option>
            {gelombangOptions.map((g) => <option key={g} value={g}>{g}</option>)}
          </select>
        </div>
        {showKelas && (
          <div>
            <label className={labelCls}>Kelas</label>
            <select className={inputCls} value={formData.kelas || ""}
              onChange={(e) => { setFormData((prev) => ({ ...prev, kelas: e.target.value })); resetFilterResult(); }}
              disabled={!selectedHotel || !selectedGelombang}>
              <option value="">Pilih kelas</option>
              {kelasOptions.map((k) => <option key={k} value={k}>{k}</option>)}
            </select>
          </div>
        )}
      </div>

      {!showKelas && (
        <div className="flex items-center gap-2 rounded-2xl border border-orange-100 bg-white/70 px-4 py-3 text-xs font-semibold text-slate-500">
          <AlertCircle size={14} className="shrink-0 text-orange-400" />
          Untuk Panitia &amp; Inda, kelas tidak diperlukan. Variabel kelas akan otomatis diisi "-" pada dokumen.
        </div>
      )}

      <button type="button" onClick={runFilter}
        className={`w-full inline-flex items-center justify-center gap-2 rounded-2xl px-5 py-3 text-sm font-black shadow-lg transition hover:-translate-y-0.5 ${filtered ? "bg-orange-600 text-white shadow-orange-500/25" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
        <Users size={16} /> Tampilkan Peserta
      </button>

      {filtered && (
        <div className="flex items-center gap-2 rounded-2xl border border-green-200 bg-green-50 px-4 py-3">
          <CheckCircle size={16} className="text-green-600" />
          <span className="text-sm font-bold text-green-700">
            {filteredPeserta.length} peserta ditemukan{selectedGroup && DAFTAR_HADIR_GROUPS[selectedGroup] ? ` — ${DAFTAR_HADIR_GROUPS[selectedGroup].label}` : ""}
          </span>
        </div>
      )}

      {filtered && filteredPeserta.length > 0 && <PesertaTablePreview peserta={filteredPeserta} />}
      {filtered && filteredPeserta.length === 0 && (
        <div className="rounded-2xl border border-dashed border-orange-200 bg-orange-50/50 p-6 text-center text-sm font-semibold text-slate-400">
          Tidak ada peserta untuk Tempat {formData.hotel}{showKelas ? `, Kelas ${formData.kelas}` : ""}, Gelombang {formData.gelombang}
        </div>
      )}
    </div>
  );
}

// ─── FILTER HOTEL PANEL ───────────────────────────────────────────────────────

export function FilterPesertaHotelPanel({ xlsxLoaded, formData, setFormData, petugasData, onFilterResult }) {
  const inputCls = "w-full rounded-2xl border border-orange-100 bg-white/80 px-4 py-3 text-sm font-semibold text-slate-800 shadow-sm outline-none focus:border-orange-300 focus:ring-2 focus:ring-orange-100 transition";
  const labelCls = "mb-2 block text-xs font-black uppercase tracking-[0.2em] text-slate-500";
  const [filtered, setFiltered]               = useState(false);
  const [filteredPeserta, setFilteredPeserta] = useState([]);
  const selectedHotel = cleanText(formData.hotel);
  const hotelOptions  = React.useMemo(() => uniqueSorted([...petugasData.map((p) => p.hotel), "STIS"]), [petugasData]);

  const runFilter = () => {
    if (!xlsxLoaded) { alert("Data petugas belum dimuat."); return; }
    if (!selectedHotel) { alert("Pilih tempat terlebih dahulu."); return; }
    const hasil = sortPesertaByJabatanOrder(petugasData.filter((p) => cleanText(p.hotel) === selectedHotel));
    setFilteredPeserta(hasil); setFiltered(true); onFilterResult(hasil);
  };
  const resetFilter = () => { setFiltered(false); setFilteredPeserta([]); onFilterResult([]); };

  return (
    <div className="rounded-3xl border border-orange-100 bg-orange-50/60 p-5 space-y-4">
      <p className="flex items-center gap-2 text-xs font-black uppercase tracking-[0.2em] text-orange-700"><Filter size={14} /> Filter Peserta Berdasarkan Tempat</p>
      <div>
        <label className={labelCls}>Tempat</label>
        <select className={inputCls} value={formData.hotel || ""}
          onChange={(e) => { setFormData((prev) => ({ ...prev, hotel: e.target.value, tempat: e.target.value })); resetFilter(); }}
          disabled={!xlsxLoaded || hotelOptions.length === 0}>
          <option value="">Pilih tempat</option>
          {hotelOptions.map((h) => <option key={h} value={h}>{h}</option>)}
        </select>
      </div>
      <button type="button" onClick={runFilter}
        className={`w-full inline-flex items-center justify-center gap-2 rounded-2xl px-5 py-3 text-sm font-black shadow-lg transition hover:-translate-y-0.5 ${filtered ? "bg-orange-600 text-white shadow-orange-500/25" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
        <Users size={16} /> Tampilkan Peserta
      </button>
      {filtered && <div className="rounded-2xl border border-green-200 bg-green-50 px-4 py-3 text-sm font-bold text-green-700">{filteredPeserta.length} peserta ditemukan untuk tempat {selectedHotel}</div>}
      {filtered && filteredPeserta.length > 0 && <PesertaTablePreview peserta={filteredPeserta} />}
    </div>
  );
}

// ─── FILTER HOTEL + GELOMBANG PANEL ──────────────────────────────────────────

export function FilterPesertaHotelGelombangPanel({ xlsxLoaded, formData, setFormData, petugasData, onFilterResult }) {
  const inputCls = "w-full rounded-2xl border border-orange-100 bg-white/80 px-4 py-3 text-sm font-semibold text-slate-800 shadow-sm outline-none focus:border-orange-300 focus:ring-2 focus:ring-orange-100 transition";
  const labelCls = "mb-2 block text-xs font-black uppercase tracking-[0.2em] text-slate-500";
  const [filtered, setFiltered]               = useState(false);
  const [filteredPeserta, setFilteredPeserta] = useState([]);
  const selectedHotel     = cleanText(formData.hotel);
  const selectedGelombang = cleanText(formData.gelombang);
  const hotelOptions      = React.useMemo(() => uniqueSorted([...petugasData.map((p) => p.hotel), "STIS"]), [petugasData]);
  const gelombangOptions  = React.useMemo(() => {
    if (cleanText(selectedHotel).toLowerCase() === "stis") return ["4"];
    return uniqueSorted(petugasData.filter((p) => !selectedHotel || cleanText(p.hotel) === selectedHotel).map((p) => p.gelombang));
  }, [petugasData, selectedHotel]);

  const runFilter = () => {
    if (!xlsxLoaded) { alert("Data petugas belum dimuat."); return; }
    if (!selectedHotel) { alert("Pilih tempat."); return; }
    if (!selectedGelombang) { alert("Pilih gelombang."); return; }
    const hasil = sortPesertaByJabatanOrder(petugasData.filter((p) => cleanText(p.hotel) === selectedHotel && cleanText(p.gelombang) === selectedGelombang));
    setFilteredPeserta(hasil); setFiltered(true); onFilterResult(hasil);
  };
  const resetFilter = () => { setFiltered(false); setFilteredPeserta([]); onFilterResult([]); };

  return (
    <div className="rounded-3xl border border-orange-100 bg-orange-50/60 p-5 space-y-4">
      <p className="flex items-center gap-2 text-xs font-black uppercase tracking-[0.2em] text-orange-700"><Filter size={14} /> Filter Peserta — Tempat &amp; Gelombang</p>
      <div>
        <label className={labelCls}>Tempat</label>
        <select className={inputCls} value={formData.hotel || ""}
          onChange={(e) => { setFormData((prev) => ({ ...prev, hotel: e.target.value, tempat: e.target.value, gelombang: "" })); resetFilter(); }}
          disabled={!xlsxLoaded || hotelOptions.length === 0}>
          <option value="">Pilih tempat</option>
          {hotelOptions.map((h) => <option key={h} value={h}>{h}</option>)}
        </select>
      </div>
      <div>
        <label className={labelCls}>Gelombang</label>
        <select className={inputCls} value={formData.gelombang || ""}
          onChange={(e) => { setFormData((prev) => ({ ...prev, gelombang: e.target.value })); resetFilter(); }}
          disabled={!selectedHotel}>
          <option value="">Pilih gelombang</option>
          {gelombangOptions.map((g) => <option key={g} value={g}>{g}</option>)}
        </select>
      </div>
      <button type="button" onClick={runFilter}
        className={`w-full inline-flex items-center justify-center gap-2 rounded-2xl px-5 py-3 text-sm font-black shadow-lg transition hover:-translate-y-0.5 ${filtered ? "bg-orange-600 text-white shadow-orange-500/25" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
        <Users size={16} /> Tampilkan Peserta
      </button>
      {filtered && <div className="rounded-2xl border border-green-200 bg-green-50 px-4 py-3 text-sm font-bold text-green-700">{filteredPeserta.length} peserta — {selectedHotel}, Gelombang {selectedGelombang}</div>}
      {filtered && filteredPeserta.length > 0 && <PesertaTablePreview peserta={filteredPeserta} />}
    </div>
  );
}

