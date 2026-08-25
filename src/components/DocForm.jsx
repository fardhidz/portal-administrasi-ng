// ============================================================
// Portal Administrasi SE2026 — bagian: docForm
// ============================================================

import React, { useState } from "react";
import { AnimatePresence, motion } from "framer-motion";
import { AlertCircle, Briefcase, Check, ChevronRight, Download, FileText, LoaderCircle, Users } from "lucide-react";

import { generateBast, generateSingleBast } from "../lib/bast";
import { buildBerkasPembayaranRecords, generateBerkasPembayaran, generateSingleBerkasPembayaran } from "../lib/berkasPembayaran";
import { generateBapp, generateLampiran, generateSingleBapp, generateSingleLampiran, generateSingleSuratPernyataanPenyelesaianLapangan, generateSuratPernyataanPenyelesaianLapangan } from "../lib/docGenerators";
import { FilterPesertaHotelGelombangPanel, FilterPesertaPanel, GenerateDocxButton, SelectionUploadPanel } from "./FormPanels";
import { buildEmailOnlySelectionKeySet, buildSelectionKeySet, cleanText, normalizeJamIndonesia, rowMatchesSelection, rowMatchesSelectionByEmail, uniqueSorted, upperText } from "../lib/helpers";
import { getBappIdentityKey, isBappRowForRole } from "../lib/parsers";
import { buildSuratKepalaRows, generateSuratKepala } from "../lib/suratKepala";
import { BAPP_PML_TEMPLATE_URL, BAPP_PPL_TEMPLATE_URL, BAST_PML_TEMPLATE_URL, BAST_PPL_TEMPLATE_URL, BERKAS_PEMBAYARAN_PML_TEMPLATE_URL, BERKAS_PEMBAYARAN_PPL_TEMPLATE_URL, DAFTAR_HADIR_GROUPS, LAMPIRAN_PML_TEMPLATE_URL, LAMPIRAN_PPL_TEMPLATE_URL, SURAT_KEPALA_TEMPLATE_URL, SURAT_PERNYATAAN_PENYELESAIAN_LAPANGAN_TEMPLATE_URL } from "../data/templates";

// ─── DOC FORM ─────────────────────────────────────────────────────────────────

export function DocForm({ docType, formData, setFormData, onPreview, petugasData, lampiranData = [], bappData = [], statusSlsData = [], dataPerSlsData = [], dataPmlProgressData = [], approveByPmlData = [], xlsxLoaded }) {
  const update = (key, val) => setFormData((p) => ({ ...p, [key]: val }));

const [gabunganRole, setGabunganRole] = useState(""); // "PML" | "PPL"
const [gabunganManualSelect, setGabunganManualSelect] = useState("");
const [gabunganGenerating, setGabunganGenerating] = useState(false);
const [gabunganProgressText, setGabunganProgressText] = useState("");
const [gabunganSelectionRows, setGabunganSelectionRows] = useState([]);

  const [daftarHadirPeserta,     setDaftarHadirPeserta]     = useState([]);
  const [daftarHadirNamaInda,    setDaftarHadirNamaInda]    = useState("");
  const [daftarHadirFiltered,    setDaftarHadirFiltered]    = useState(false);
  const [daftarHadirFilterGroup, setDaftarHadirFilterGroup] = useState("");

  const [tandaTerimaPeserta,      setTandaTerimaPeserta]      = useState([]);
  const [tandaTerimaFiltered,     setTandaTerimaFiltered]     = useState(false);
  const [tandaTerimaType,         setTandaTerimaType]         = useState("");
  const [tandaTerimaWilTugas,     setTandaTerimaWilTugas]     = useState("");
  const [tandaTerimaLokasi,       setTandaTerimaLokasi]       = useState("BPS Kota Jakarta Timur");
  const [suratPernyataanPeserta,  setSuratPernyataanPeserta]  = useState([]);
  const [suratPernyataanFiltered, setSuratPernyataanFiltered] = useState(false);
  const [suratPernyataanFilterGroup, setSuratPernyataanFilterGroup] = useState("");
  const [suratTugasPeserta,       setSuratTugasPeserta]       = useState([]);
  const [suratTugasFiltered,      setSuratTugasFiltered]      = useState(false);
  const [spjPeserta,              setSpjPeserta]              = useState([]);
  const [spjFiltered,             setSpjFiltered]             = useState(false);
  const [spjFilterGroup,          setSpjFilterGroup]          = useState("");
  const [spdPeserta,              setSpdPeserta]              = useState([]);
  const [spdFiltered,             setSpdFiltered]             = useState(false);
  const [pengeluaranPeserta,      setPengeluaranPeserta]      = useState([]);
  const [pengeluaranFiltered,     setPengeluaranFiltered]     = useState(false);
  const [pengeluaranFilterGroup,  setPengeluaranFilterGroup]  = useState("");
  const [bappRole,                setBappRole]                = useState("");
  const [bappManualSelect,        setBappManualSelect]        = useState("");
  const [bappGenerating,          setBappGenerating]          = useState(false);
  const [bappProgressText,        setBappProgressText]        = useState("");
  const [bappSelectionRows,       setBappSelectionRows]       = useState([]);
  const [bastRole,                setBastRole]                = useState("");
  const [bastManualSelect,        setBastManualSelect]        = useState("");
  const [bastGenerating,          setBastGenerating]          = useState(false);
  const [bastProgressText,        setBastProgressText]        = useState("");
  const [bastSelectionRows,       setBastSelectionRows]       = useState([]);
  const [suratKepalaSelectionRows, setSuratKepalaSelectionRows] = useState([]);
  const [suratKepalaGenerating, setSuratKepalaGenerating] = useState(false);
  const [suratKepalaProgressText, setSuratKepalaProgressText] = useState("");
  const [suratPenyelesaianLapanganSelect, setSuratPenyelesaianLapanganSelect] = useState("");
  const [suratPenyelesaianLapanganGenerating, setSuratPenyelesaianLapanganGenerating] = useState(false);
  const [suratPenyelesaianLapanganProgressText, setSuratPenyelesaianLapanganProgressText] = useState("");
  const [suratPenyelesaianLapanganSelectionRows, setSuratPenyelesaianLapanganSelectionRows] = useState([]);

  // Lampiran form controls: combined manual select for PML + PPL
  const [lampiranManualSelect, setLampiranManualSelect] = useState("");
  const [lampiranSelectionRows, setLampiranSelectionRows] = useState([]);
  const [lampiranBeberapaGenerating, setLampiranBeberapaGenerating] = useState(false);
  const [lampiranBeberapaProgressText, setLampiranBeberapaProgressText] = useState("");

  React.useEffect(() => {
    if (docType.id === "daftar-hadir" && !formData.jamMulai && !formData.jamSelesai) {
      setFormData((prev) => ({ ...prev, jamMulai: "07.30", jamSelesai: "18.00" }));
    }
  }, [docType.id]);

  const inputCls = "w-full rounded-2xl border border-orange-100 bg-white/80 px-4 py-3 text-sm font-semibold text-slate-800 shadow-sm outline-none focus:border-orange-300 focus:ring-2 focus:ring-orange-100 transition";
  const labelCls = "mb-2 block text-xs font-black uppercase tracking-[0.2em] text-slate-500";

  const handleSubmit = (e) => {
    e.preventDefault();
    
    if (docType.id === "daftar-hadir") {
      if (!daftarHadirFilterGroup) {
        alert("Pilih kelompok peserta (PML & PPL atau Panitia & Inda) terlebih dahulu.");
        return;
      }
      if (!daftarHadirFiltered || daftarHadirPeserta.length === 0) {
        alert("Tampilkan peserta terlebih dahulu dengan mengklik tombol filter.");
        return;
      }
      const isPanitiaInda = daftarHadirFilterGroup === "panitia-inda";
      onPreview({
        formValues: {
          ...formData,
          tempat:          formData.tempat || formData.hotel || "",
          kelompokPeserta: DAFTAR_HADIR_GROUPS[daftarHadirFilterGroup]?.label || "",
          kelas:           isPanitiaInda ? "-" : (formData.kelas || ""),
        },
        peserta:             daftarHadirPeserta,
        namaInda:            daftarHadirNamaInda,
        selectedFilterGroup: daftarHadirFilterGroup,
      });
      return;
    }

    if (docType.id === "tanda-terima") {
      if (!tandaTerimaType) { alert("Pilih jenis tanda terima terlebih dahulu."); return; }
      if (!tandaTerimaFiltered || tandaTerimaPeserta.length === 0) { alert("Tampilkan peserta terlebih dahulu."); return; }
      onPreview({ formValues: { ...formData, tempat: tandaTerimaType === "mitra-umum" ? tandaTerimaLokasi : (formData.tempat || formData.hotel || "") }, peserta: tandaTerimaPeserta, tandaTerimaType: tandaTerimaType });
      return;
    }

    if (docType.id === "surat-pernyataan-kendaraan") {
      if (!suratPernyataanFilterGroup) {
        alert("Pilih kelompok peserta (PML & PPL atau Panitia & Inda) terlebih dahulu.");
        return;
      }
      if (!suratPernyataanFiltered || suratPernyataanPeserta.length === 0) { alert("Tampilkan peserta terlebih dahulu."); return; }
      const isPanitiaInda = suratPernyataanFilterGroup === "panitia-inda";
      onPreview({
        formValues: {
          tanggal_surat: formData.tanggal_surat || "",
          tempat:        formData.tempat || formData.hotel || "",
          hotel:         formData.hotel || "",
          gelombang:     formData.gelombang || "",
          kelompokPeserta: suratPernyataanFilterGroup === "panitia-inda" ? "Panitia & Inda" : "Petugas PML & PPL",
          kelas:         isPanitiaInda ? "-" : (formData.kelas || "-"),
        },
        peserta: suratPernyataanPeserta,
      });
      return;
    }

    if (docType.id === "pengeluaran-riil") {
      if (!pengeluaranFilterGroup) {
        alert("Pilih kelompok peserta (PML & PPL atau Panitia & Inda) terlebih dahulu.");
        return;
      }
      if (!pengeluaranFiltered || pengeluaranPeserta.length === 0) { alert("Tampilkan peserta terlebih dahulu."); return; }
      const isPanitiaInda = pengeluaranFilterGroup === "panitia-inda";
      onPreview({
        formValues: {
          tanggal_surat: formData.tanggal_surat || "",
          hotel:         formData.hotel || "",
          kelompokPeserta: pengeluaranFilterGroup === "panitia-inda" ? "Panitia & Inda" : "Petugas PML & PPL",
          kelas:         isPanitiaInda ? "-" : (formData.kelas || "-"),
        },
        peserta: pengeluaranPeserta,
      });
      return;
    }

    if (docType.id === "spj") {
      if (!spjFilterGroup) {
        alert("Pilih kelompok peserta SPJ (PML & PPL atau Panitia & Inda) terlebih dahulu.");
        return;
      }
      if (!spjFiltered || spjPeserta.length === 0) {
        alert("Tampilkan peserta terlebih dahulu.");
        return;
      }
      const isPanitiaInda = spjFilterGroup === "panitia-inda";
      onPreview({
        formValues: {
          ...formData,
          tanggal_pelunasan: formData.tanggal_pelunasan || "",
          tempat:            formData.tempat || formData.hotel || "",
          kelompokPeserta:   spjFilterGroup === "panitia-inda" ? "Panitia & Inda" : "Petugas PML & PPL",
          kelas:             isPanitiaInda ? "-" : (formData.kelas || ""),
          ttd_kiri:          formData.ttd_kiri || "",
          ttd_kanan:         formData.ttd_kanan || "",
        },
        peserta: spjPeserta,
        selectedFilterGroup: spjFilterGroup,
      });
      return;
    }

    if (docType.id === "spd") {
      if (!spdFiltered || spdPeserta.length === 0) {
        alert("Tampilkan peserta terlebih dahulu.");
        return;
      }

      onPreview({
        formValues: {
          nomor_dokumen:           formData.nomor_dokumen || formData.nomor || "",
          nomor:                   formData.nomor_dokumen || formData.nomor || "",
          tanggal_surat:           formData.tanggal_surat || "",
          tanggal:                 formData.tanggal_surat || "",
          tanggal_awal_kegiatan:   formData.tanggal_awal_kegiatan || "",
          tanggal_akhir_kegiatan:  formData.tanggal_akhir_kegiatan || "",
          tempat:                  formData.tempat || formData.hotel || "",
          hotel:                   formData.hotel || "",
          gelombang:               formData.gelombang || "",
          kelas:                   "-",
          lokasi:                  formData.tempat || formData.hotel || "",
          namaKabps:               formData.namaKabps || "",
          nipKabps:                formData.nipKabps || "",
        },
        peserta: spdPeserta,
      });
      return;
    }

    if (docType.id === "surat-tugas") {
      if (!suratTugasFiltered || suratTugasPeserta.length === 0) { alert("Tampilkan peserta terlebih dahulu."); return; }
      onPreview({
        formValues: {
          nomor_surat:            formData.nomor_surat || "",
          tanggal_surat:          formData.tanggal_surat || "",
          tanggal_awal_kegiatan:  formData.tanggal_awal_kegiatan || "",
          tanggal_akhir_kegiatan: formData.tanggal_akhir_kegiatan || "",
          tempat:                 formData.tempat || formData.hotel || "",
          hotel:                  formData.hotel || "",
          gelombang:              formData.gelombang || "",
          kelas:                  "-",
        },
        peserta: suratTugasPeserta,
      });
      return;
    }

    if (docType.id === "lampiran") {
      // Lampiran tidak memakai input Tempat/Gelombang/Kelas.
      // Data langsung diambil dari sheet bernama "Lampiran" dan tombol PML/PPL akan generate template masing-masing.
      onPreview({
        formValues: {},
        lampiranRows: lampiranData || [],
      });
      return;
    }

    if (docType.id === "bapp") {
      e.preventDefault();
      return;
    }

    if (docType.id === "bast") {
      e.preventDefault();
      return;
    }

    if (docType.id === "gabungan-pembayaran") {
      e.preventDefault();
      return;
    }

    if (docType.id === "surat-kepala") {
      e.preventDefault();
      return;
    }

    onPreview({ ...formData });
  };

  const renderFields = () => {
    switch (docType.id) {

      // ── ADMINISTRASI PEMBAYARAN LENGKAP PER PML/PPL ──────────────────────
      case "gabungan-pembayaran": {
        const nikLookup = React.useMemo(() => {
          const map = new Map();
          for (const p of petugasData || []) {
            const nama = upperText(p.nama);
            if (nama && !map.has(nama)) map.set(nama, cleanText(p.nik));
          }
          return map;
        }, [petugasData]);

        const berkasRecords = React.useMemo(() => {
          if (!gabunganRole) return [];
          return buildBerkasPembayaranRecords(
            bappData || [],
            lampiranData || [],
            gabunganRole,
            statusSlsData || [],
            dataPerSlsData || [],
            approveByPmlData || []
          );
        }, [bappData, lampiranData, gabunganRole, statusSlsData, dataPerSlsData, approveByPmlData]);

        const berkasOptions = React.useMemo(() => {
          return berkasRecords.map((record) => ({
            value: record.identity,
            label: `${record.displayName || "Tanpa Nama"}${record.email ? ` — ${record.email}` : ""}`,
            record,
          }));
        }, [berkasRecords]);

        const validateTanggal = () => {
          if (!formData.tanggal_surat) throw new Error("Isi tanggal surat terlebih dahulu.");
          const selectedDate = new Date(formData.tanggal_surat);
          const minDate = new Date("2026-07-15T00:00:00");
          const maxDate = new Date("2026-08-31T23:59:59");
          if (selectedDate < minDate || selectedDate > maxDate) {
            throw new Error("Tanggal surat hanya boleh 15 Juli 2026 sampai 31 Agustus 2026.");
          }
        };

        const getTemplateUrl = () => gabunganRole === "PML"
          ? BERKAS_PEMBAYARAN_PML_TEMPLATE_URL
          : BERKAS_PEMBAYARAN_PPL_TEMPLATE_URL;

        return (
          <div className="space-y-5">
            <div className="rounded-3xl border border-orange-100 bg-orange-50/70 p-5">
              <p className="text-xs font-black uppercase tracking-[0.2em] text-orange-700">Gabungan Administrasi Pembayaran</p>
              <p className="mt-2 text-sm font-semibold leading-6 text-slate-600">
                Pilih PML atau PPL, lalu generate satu berkas pembayaran lengkap per orang. Pilihan tersedia untuk satu nama, beberapa nama melalui Excel, atau semua nama.
              </p>
              <p className="mt-3 text-xs font-bold text-slate-500">
                Template: {gabunganRole === "PML" ? "BERKAS PEMBAYARAN PML.docx" : gabunganRole === "PPL" ? "BERKAS PEMBAYARAN PPL.docx" : "pilih role terlebih dahulu"}
              </p>
            </div>

            <div>
              <p className={labelCls}>Pilih Role</p>
              <div className="flex gap-3">
                <button type="button" onClick={() => {
                  setGabunganRole("PML");
                  setGabunganManualSelect("");
                  setGabunganSelectionRows([]);
                }} className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${gabunganRole === "PML" ? "bg-orange-600 text-white" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
                  PML
                </button>
                <button type="button" onClick={() => {
                  setGabunganRole("PPL");
                  setGabunganManualSelect("");
                  setGabunganSelectionRows([]);
                }} className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${gabunganRole === "PPL" ? "bg-orange-600 text-white" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
                  PPL
                </button>
              </div>
            </div>

            {gabunganRole && (
              <>
                <div>
                  <label className={labelCls}>Tanggal Surat</label>
                  <input
                    type="date"
                    min="2026-07-15"
                    max="2026-08-31"
                    className={inputCls}
                    value={formData.tanggal_surat || ""}
                    onChange={(e) => update("tanggal_surat", e.target.value)}
                  />
                  <p className="mt-1 text-xs font-semibold text-slate-400">Rentang tanggal yang diizinkan: 15–31 Agustus 2026.</p>
                </div>

                <div>
                  <label className={labelCls}>Pilih Sendiri</label>
                  <select value={gabunganManualSelect} onChange={(e) => setGabunganManualSelect(e.target.value)} className={inputCls}>
                    <option value="">— Pilih Nama {gabunganRole} —</option>
                    {berkasOptions.map((option) => (
                      <option key={option.value} value={option.value}>{option.label}</option>
                    ))}
                  </select>
                  <p className="mt-1 text-xs font-semibold text-slate-400">
                    {berkasRecords.length} orang terdeteksi. Filter Lampiran memakai {statusSlsData.length} baris Status SLS. Tabel beban kerja memakai {dataPerSlsData.length} baris Data per SLS. Jumlah PML memakai {approveByPmlData.length} baris Approve by PML.
                  </p>
                </div>

                <SelectionUploadPanel
                  selectionRows={gabunganSelectionRows}
                  onSelectionLoaded={setGabunganSelectionRows}
                  onClear={() => setGabunganSelectionRows([])}
                  hint={`Unggah Excel berisi kolom Nama dan/atau Email untuk memilih beberapa ${gabunganRole}.`}
                />

                {gabunganGenerating && (
                  <div className="flex items-center gap-3 rounded-2xl border border-orange-200 bg-orange-50 px-4 py-3 text-sm font-semibold text-orange-700">
                    <LoaderCircle size={18} className="animate-spin" />
                    <span>{gabunganProgressText || `Sedang membuat berkas pembayaran ${gabunganRole}...`}</span>
                  </div>
                )}

                <div className="grid gap-4 sm:grid-cols-3">
                  <button type="button" onClick={async () => {
                    try {
                      validateTanggal();
                      if (!gabunganManualSelect) throw new Error("Pilih nama terlebih dahulu.");
                      const chosen = berkasOptions.find((option) => option.value === gabunganManualSelect)?.record;
                      if (!chosen) throw new Error("Data nama yang dipilih tidak ditemukan.");
                      setGabunganGenerating(true);
                      setGabunganProgressText("Membuat satu berkas pembayaran...");
                      await generateSingleBerkasPembayaran(
                        getTemplateUrl(), formData, chosen, gabunganRole, nikLookup
                      );
                    } catch (err) { alert(err.message || err); }
                    finally { setGabunganGenerating(false); setGabunganProgressText(""); }
                  }} disabled={!gabunganManualSelect || berkasRecords.length === 0 || gabunganGenerating}
                    className="inline-flex items-center justify-center rounded-2xl bg-orange-500 px-5 py-3 text-sm font-black text-white shadow transition hover:bg-orange-600 disabled:cursor-not-allowed disabled:bg-orange-200">
                    Generate Terpilih
                  </button>

                  <button type="button" onClick={async () => {
                    try {
                      validateTanggal();
                      if (gabunganSelectionRows.length === 0) throw new Error("Unggah file Excel Email terlebih dahulu.");
                      const keySet = buildEmailOnlySelectionKeySet(gabunganSelectionRows);
                      const matchedRecords = berkasRecords.filter((record) =>
                        rowMatchesSelectionByEmail(keySet, record.email)
                      );
                      if (matchedRecords.length === 0) throw new Error("Tidak ada data yang cocok dengan file Excel.");
                      setGabunganGenerating(true);
                      setGabunganProgressText(`Menyiapkan ${matchedRecords.length} berkas pembayaran...`);
                      await generateBerkasPembayaran(
                        getTemplateUrl(), formData, matchedRecords, gabunganRole, nikLookup,
                        ({ batchIndex, totalBatches }) => setGabunganProgressText(`Membuat batch ${batchIndex} dari ${totalBatches}...`)
                      );
                    } catch (err) { alert(err.message || err); }
                    finally { setGabunganGenerating(false); setGabunganProgressText(""); }
                  }} disabled={berkasRecords.length === 0 || gabunganGenerating || gabunganSelectionRows.length === 0}
                    className="inline-flex items-center justify-center rounded-2xl border border-orange-200 bg-white px-5 py-3 text-sm font-black text-orange-700 shadow transition hover:bg-orange-50 disabled:cursor-not-allowed disabled:opacity-60">
                    Generate Beberapa
                  </button>

                  <button type="button" onClick={async () => {
                    try {
                      validateTanggal();
                      if (berkasRecords.length === 0) throw new Error(`Tidak ada data ${gabunganRole}.`);
                      setGabunganGenerating(true);
                      setGabunganProgressText("Mempersiapkan semua berkas pembayaran...");
                      await generateBerkasPembayaran(
                        getTemplateUrl(), formData, berkasRecords, gabunganRole, nikLookup,
                        ({ batchIndex, totalBatches }) => setGabunganProgressText(`Membuat batch ${batchIndex} dari ${totalBatches}...`)
                      );
                    } catch (err) { alert(err.message || err); }
                    finally { setGabunganGenerating(false); setGabunganProgressText(""); }
                  }} disabled={berkasRecords.length === 0 || gabunganGenerating}
                    className="inline-flex items-center justify-center rounded-2xl border border-orange-200 bg-white px-5 py-3 text-sm font-black text-orange-700 shadow transition hover:bg-orange-50 disabled:cursor-not-allowed disabled:opacity-60">
                    Generate Semua
                  </button>
                </div>
              </>
            )}
          </div>
        );
      }

      // ── DAFTAR HADIR ────────────────────────────────────────────────────────
      case "daftar-hadir":
        return (
          <>
            {!xlsxLoaded && (
              <div className="flex items-center gap-3 rounded-2xl border border-amber-200 bg-amber-50 px-4 py-3">
                <AlertCircle size={18} className="shrink-0 text-amber-500" />
                <p className="text-sm font-semibold text-amber-700">Data petugas belum terdeteksi. Unggah file dari dashboard.</p>
              </div>
            )}

            <div>
              <label className={labelCls}>Tanggal Kegiatan</label>
              <input type="date" className={inputCls} value={formData.tanggal || ""} onChange={(e) => update("tanggal", e.target.value)} />
            </div>

            <div className="grid gap-4 sm:grid-cols-2">
              <div>
                <label className={labelCls}>Jam Mulai</label>
                <input className={inputCls} placeholder="07.30" value={formData.jamMulai || "07.30"} onChange={(e) => update("jamMulai", normalizeJamIndonesia(e.target.value))} />
              </div>
              <div>
                <label className={labelCls}>Jam Selesai</label>
                <input className={inputCls} placeholder="18.00" value={formData.jamSelesai || "18.00"} onChange={(e) => update("jamSelesai", normalizeJamIndonesia(e.target.value))} />
              </div>
            </div>

            <div>
              <p className={labelCls}>Kelompok Peserta</p>
              <div className="flex gap-3">
                {Object.entries(DAFTAR_HADIR_GROUPS).map(([key, grp]) => (
                  <button key={key} type="button"
                    onClick={() => {
                      setDaftarHadirFilterGroup(key);
                      setDaftarHadirFiltered(false);
                      setDaftarHadirPeserta([]);
                      setDaftarHadirNamaInda("");
                      if (key === "panitia-inda") {
                        setFormData((prev) => ({ ...prev, kelas: "" }));
                      }
                    }}
                    className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${daftarHadirFilterGroup === key ? "bg-orange-600 text-white shadow-lg shadow-orange-500/20" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
                    {key === "pml-ppl" ? <Users size={15} /> : <Briefcase size={15} />}
                    {grp.label}
                  </button>
                ))}
              </div>
              {!daftarHadirFilterGroup && (
                <p className="mt-2 text-xs font-semibold text-slate-400">Pilih kelompok untuk menampilkan filter peserta.</p>
              )}
            </div>

            <AnimatePresence mode="wait">
              {daftarHadirFilterGroup && (
                <motion.div key={daftarHadirFilterGroup} initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -8 }} transition={{ duration: 0.2 }}>
                  <FilterPesertaPanel
                    xlsxLoaded={xlsxLoaded}
                    formData={formData}
                    setFormData={setFormData}
                    petugasData={petugasData}
                    mode="grouped"
                    selectedGroup={daftarHadirFilterGroup}
                    prependRow={(() => {
                      const hotelLower = cleanText(formData.hotel || formData.tempat || "").toLowerCase();
                      const gelombangVal = cleanText(formData.gelombang || "");
                      if (daftarHadirFilterGroup === "panitia-inda" && gelombangVal === "4") {
                        if (hotelLower.includes("bwp")) return { nama: "Budi Utami", jabatan: "Penanggung Jawab", wilTugas: "BPS Kota Jakarta Timur" };
                        if (hotelLower.includes("stis")) return { nama: "Widiastuti", jabatan: "Kepala BPS Kota Jakarta Timur", wilTugas: "BPS Kota Jakarta Timur" };
                      }
                      return null;
                    })()}
                    onFilterResult={(peserta, namaInda) => {
                      setDaftarHadirPeserta(peserta);
                      setDaftarHadirNamaInda(namaInda);
                      setDaftarHadirFiltered(peserta.length > 0);
                    }}
                  />
                </motion.div>
              )}
            </AnimatePresence>

            {daftarHadirFiltered && daftarHadirFilterGroup && (
              <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }}
                className="flex items-center justify-between rounded-2xl border border-orange-100 bg-white/80 px-4 py-3">
                <div>
                  <p className="text-xs font-black uppercase tracking-wider text-slate-500">Tanda Tangan</p>
                  <p className="mt-0.5 font-bold text-slate-800">
                    {daftarHadirFilterGroup === "panitia-inda" ? "Kepala Sub Bagian Umum" : "Instruktur Daerah"}
                  </p>
                </div>
                <div className="flex items-center gap-3">
                  <div className="text-right">
                    <p className="text-xs font-semibold text-slate-400">Kelas di dokumen</p>
                    <p className="font-black text-orange-600">
                      {daftarHadirFilterGroup === "panitia-inda" ? "-" : (formData.kelas || "—")}
                    </p>
                  </div>
                  <Check size={18} className="text-green-500" />
                </div>
              </motion.div>
            )}
          </>
        );

      // ── TANDA TERIMA ────────────────────────────────────────────────────────
      case "tanda-terima": {
        // ✅ FIX: wilTugasOptions hanya dari petugas non-STIS
        const wilTugasOptions = React.useMemo(
          () =>
            uniqueSorted(
              (petugasData || [])
                .filter((p) => cleanText(p.hotel).toUpperCase() !== "STIS")
                .map((p) => cleanText(p.wilTugas))
                .filter(Boolean)
            ),
          [petugasData]
        );

        return (
          <>
            <div>
              <p className={labelCls}>Jenis Tanda Terima</p>
              <div className="flex flex-wrap gap-3">
                <button type="button"
                  onClick={() => { setTandaTerimaType("pelatihan"); setTandaTerimaFiltered(false); setTandaTerimaPeserta([]); setTandaTerimaWilTugas(""); }}
                  className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${tandaTerimaType === "pelatihan" ? "bg-orange-600 text-white shadow-lg shadow-orange-500/20" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
                  <Briefcase size={15} /> Tanda Terima Perlengkapan Pelatihan
                </button>
                <button type="button"
                  onClick={() => { setTandaTerimaType("lapangan"); setTandaTerimaFiltered(false); setTandaTerimaPeserta([]); setTandaTerimaWilTugas(""); }}
                  className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${tandaTerimaType === "lapangan" ? "bg-orange-600 text-white shadow-lg shadow-orange-500/20" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
                  <Briefcase size={15} /> Tanda Terima Perlengkapan Lapangan
                </button>
                <button type="button"
                  onClick={() => { setTandaTerimaType("mitra-umum"); setTandaTerimaFiltered(false); setTandaTerimaPeserta([]); setTandaTerimaWilTugas(""); setTandaTerimaLokasi("BPS Kota Jakarta Timur"); }}
                  className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${tandaTerimaType === "mitra-umum" ? "bg-orange-600 text-white shadow-lg shadow-orange-500/20" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
                  <Briefcase size={15} /> Tanda Terima Perlengkapan Lapangan - Mitra Umum
                </button>
              </div>
              {!tandaTerimaType && (
                <p className="mt-2 text-xs font-semibold text-slate-400">Pilih jenis tanda terima terlebih dahulu.</p>
              )}
            </div>
            {tandaTerimaType && (tandaTerimaType === "mitra-umum" ? (
              <>
                <div>
                  <label className={labelCls}>Tanggal Kegiatan</label>
                  <input type="date" className={inputCls} value={formData.tanggal || ""} onChange={(e) => update("tanggal", e.target.value)} />
                </div>
                <div>
                  <label className={labelCls}>Pilih Wilayah Tugas</label>
                  {/* ✅ FIX: dropdown wilTugas sudah exclude TC = STIS, filter data juga exclude STIS */}
                  <select className={inputCls} value={tandaTerimaWilTugas} onChange={(e) => {
                    const newWilTugas = e.target.value;
                    setTandaTerimaWilTugas(newWilTugas);
                    const filtered = newWilTugas
                      ? (petugasData || []).filter(
                          (p) =>
                            cleanText(p.wilTugas).toUpperCase() === newWilTugas.toUpperCase() &&
                            cleanText(p.hotel).toUpperCase() !== "STIS"
                        )
                      : [];
                    setTandaTerimaPeserta(filtered);
                    setTandaTerimaFiltered(filtered.length > 0);
                  }}>
                    <option value="">-- Pilih Wilayah Tugas --</option>
                    {wilTugasOptions.map((wil) => (
                      <option key={wil} value={wil}>{wil}</option>
                    ))}
                  </select>
                  {tandaTerimaWilTugas && (
                    <p className="mt-1 text-xs font-semibold text-slate-400">
                      {tandaTerimaPeserta.length} petugas non-STIS ditemukan untuk wilayah ini.
                    </p>
                  )}
                </div>
                <div>
                  <label className={labelCls}>Lokasi</label>
                  <input type="text" className={inputCls} value={tandaTerimaLokasi} onChange={(e) => setTandaTerimaLokasi(e.target.value)} />
                </div>
              </>
            ) : (
              <>
                <div>
                  <label className={labelCls}>Tanggal Kegiatan</label>
                  <input type="date" className={inputCls} value={formData.tanggal || ""} onChange={(e) => update("tanggal", e.target.value)} />
                </div>
                <FilterPesertaPanel xlsxLoaded={xlsxLoaded} formData={formData} setFormData={setFormData} petugasData={petugasData} mode="all" selectedGroup="" onFilterResult={(peserta) => { setTandaTerimaPeserta(peserta); setTandaTerimaFiltered(peserta.length > 0); }} />
              </>
            ))}
          </>
        );
      }

      // ── SURAT PERNYATAAN KENDARAAN / SUPER KENDIS ──────────────────────────
      case "surat-pernyataan-kendaraan":
        return (
          <>
            <div>
              <label className={labelCls}>Tanggal Surat</label>
              <input type="date" className={inputCls} value={formData.tanggal_surat || ""} onChange={(e) => update("tanggal_surat", e.target.value)} />
            </div>

            <div>
              <p className={labelCls}>Kelompok Peserta</p>
              <div className="flex gap-3">
                {Object.entries(DAFTAR_HADIR_GROUPS).map(([key, grp]) => (
                  <button key={key} type="button"
                    onClick={() => {
                      setSuratPernyataanFilterGroup(key);
                      setSuratPernyataanFiltered(false);
                      setSuratPernyataanPeserta([]);
                      if (key === "panitia-inda") setFormData((prev) => ({ ...prev, kelas: "" }));
                    }}
                    className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${suratPernyataanFilterGroup === key ? "bg-orange-600 text-white shadow-lg shadow-orange-500/20" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
                    {key === "pml-ppl" ? <Users size={15} /> : <Briefcase size={15} />}
                    {key === "pml-ppl" ? "Petugas PML & PPL" : grp.label}
                  </button>
                ))}
              </div>
              {!suratPernyataanFilterGroup && (
                <p className="mt-2 text-xs font-semibold text-slate-400">Pilih kelompok peserta agar Super Kendis dapat dibagi menjadi PML/PPL atau Panitia/Inda.</p>
              )}
            </div>

            <AnimatePresence mode="wait">
              {suratPernyataanFilterGroup && (
                <motion.div key={suratPernyataanFilterGroup} initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -8 }} transition={{ duration: 0.2 }}>
                  <FilterPesertaPanel
                    xlsxLoaded={xlsxLoaded}
                    formData={formData}
                    setFormData={setFormData}
                    petugasData={petugasData}
                    mode="grouped"
                    selectedGroup={suratPernyataanFilterGroup}
                    onFilterResult={(peserta) => {
                      setSuratPernyataanPeserta(peserta);
                      setSuratPernyataanFiltered(peserta.length > 0);
                    }}
                  />
                </motion.div>
              )}
            </AnimatePresence>

            {suratPernyataanFiltered && suratPernyataanFilterGroup && (
              <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }}
                className="flex items-center justify-between rounded-2xl border border-orange-100 bg-white/80 px-4 py-3">
                <div>
                  <p className="text-xs font-black uppercase tracking-wider text-slate-500">Kelompok Super Kendis</p>
                  <p className="mt-0.5 font-bold text-slate-800">
                    {suratPernyataanFilterGroup === "panitia-inda" ? "Panitia & Inda" : "Petugas PML & PPL"}
                  </p>
                </div>
                <div className="flex items-center gap-3">
                  <div className="text-right">
                    <p className="text-xs font-semibold text-slate-400">Kelas di dokumen</p>
                    <p className="font-black text-orange-600">
                      {suratPernyataanFilterGroup === "panitia-inda" ? "-" : (formData.kelas || "—")}
                    </p>
                  </div>
                  <Check size={18} className="text-green-500" />
                </div>
              </motion.div>
            )}
          </>
        );

      // ── PENGELUARAN RIIL ────────────────────────────────────────────────────
      case "pengeluaran-riil":
        return (
          <>
            <div>
              <label className={labelCls}>Tanggal Surat</label>
              <input type="date" className={inputCls} value={formData.tanggal_surat || ""} onChange={(e) => update("tanggal_surat", e.target.value)} />
            </div>
            <div>
              <p className={labelCls}>Kelompok Peserta</p>
              <div className="flex gap-3">
                {Object.entries(DAFTAR_HADIR_GROUPS).map(([key, grp]) => (
                  <button key={key} type="button"
                    onClick={() => {
                      setPengeluaranFilterGroup(key);
                      setPengeluaranFiltered(false);
                      setPengeluaranPeserta([]);
                      if (key === "panitia-inda") setFormData((prev) => ({ ...prev, kelas: "" }));
                    }}
                    className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${pengeluaranFilterGroup === key ? "bg-orange-600 text-white shadow-lg shadow-orange-500/20" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
                    {key === "pml-ppl" ? <Users size={15} /> : <Briefcase size={15} />}
                    {key === "pml-ppl" ? "Petugas PML & PPL" : grp.label}
                  </button>
                ))}
              </div>
              {!pengeluaranFilterGroup && (
                <p className="mt-2 text-xs font-semibold text-slate-400">Pilih kelompok peserta agar DPR dapat dipisah menjadi PML/PPL atau Panitia/Inda.</p>
              )}
            </div>

            <AnimatePresence mode="wait">
              {pengeluaranFilterGroup && (
                <motion.div key={pengeluaranFilterGroup} initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -8 }} transition={{ duration: 0.2 }}>
                  <FilterPesertaPanel
                    xlsxLoaded={xlsxLoaded}
                    formData={formData}
                    setFormData={setFormData}
                    petugasData={petugasData}
                    mode="grouped"
                    selectedGroup={pengeluaranFilterGroup}
                    onFilterResult={(peserta) => {
                      setPengeluaranPeserta(peserta);
                      setPengeluaranFiltered(peserta.length > 0);
                    }}
                  />
                </motion.div>
              )}
            </AnimatePresence>

            {pengeluaranFiltered && pengeluaranFilterGroup && (
              <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }}
                className="flex items-center justify-between rounded-2xl border border-orange-100 bg-white/80 px-4 py-3">
                <div>
                  <p className="text-xs font-black uppercase tracking-wider text-slate-500">Kelompok DPR</p>
                  <p className="mt-0.5 font-bold text-slate-800">
                    {pengeluaranFilterGroup === "panitia-inda" ? "Panitia & Inda" : "Petugas PML & PPL"}
                  </p>
                </div>
                <div className="flex items-center gap-3">
                  <div className="text-right">
                    <p className="text-xs font-semibold text-slate-400">Kelas di dokumen</p>
                    <p className="font-black text-orange-600">
                      {pengeluaranFilterGroup === "panitia-inda" ? "-" : (formData.kelas || "—")}
                    </p>
                  </div>
                  <Check size={18} className="text-green-500" />
                </div>
              </motion.div>
            )}
          </>
        );

      // ── SPJ ─────────────────────────────────────────────────────────────────
      case "spj":
        return (
          <>
            <div>
              <label className={labelCls}>Tanggal Pelunasan</label>
              <input type="date" className={inputCls} value={formData.tanggal_pelunasan || ""} onChange={(e) => update("tanggal_pelunasan", e.target.value)} />
            </div>

            <div className="grid gap-5 sm:grid-cols-2">
              <div>
                <label className={labelCls}>Tanggal Awal</label>
                <input type="date" className={inputCls} value={formData.tanggal_awal_kegiatan || ""} onChange={(e) => update("tanggal_awal_kegiatan", e.target.value)} />
              </div>
              <div>
                <label className={labelCls}>Tanggal Akhir</label>
                <input type="date" className={inputCls} value={formData.tanggal_akhir_kegiatan || ""} onChange={(e) => update("tanggal_akhir_kegiatan", e.target.value)} />
              </div>
            </div>

            <div>
              <p className={labelCls}>Kelompok Peserta SPJ</p>
              <div className="flex gap-3">
                {Object.entries(DAFTAR_HADIR_GROUPS).map(([key, grp]) => (
                  <button key={key} type="button"
                    onClick={() => {
                      setSpjFilterGroup(key);
                      setSpjFiltered(false);
                      setSpjPeserta([]);
                      if (key === "panitia-inda") {
                        setFormData((prev) => ({ ...prev, kelas: "" }));
                      }
                    }}
                    className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${spjFilterGroup === key ? "bg-orange-600 text-white shadow-lg shadow-orange-500/20" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
                    {key === "pml-ppl" ? <Users size={15} /> : <Briefcase size={15} />}
                    {key === "pml-ppl" ? "Petugas PML & PPL" : grp.label}
                  </button>
                ))}
              </div>
              {!spjFilterGroup && (
                <p className="mt-2 text-xs font-semibold text-slate-400">Pilih kelompok agar SPJ terpisah antara Panitia & Inda dan Petugas PML & PPL.</p>
              )}
            </div>

            <AnimatePresence mode="wait">
              {spjFilterGroup && (
                <motion.div key={spjFilterGroup} initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -8 }} transition={{ duration: 0.2 }}>
                  <FilterPesertaPanel
                    xlsxLoaded={xlsxLoaded}
                    formData={formData}
                    setFormData={setFormData}
                    petugasData={petugasData}
                    mode="grouped"
                    selectedGroup={spjFilterGroup}
                    onFilterResult={(peserta) => {
                      setSpjPeserta(peserta);
                      setSpjFiltered(peserta.length > 0);
                    }}
                  />
                </motion.div>
              )}
            </AnimatePresence>

            {spjFiltered && spjFilterGroup && (
              <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }}
                className="flex items-center justify-between rounded-2xl border border-orange-100 bg-white/80 px-4 py-3">
                <div>
                  <p className="text-xs font-black uppercase tracking-wider text-slate-500">Kelompok SPJ</p>
                  <p className="mt-0.5 font-bold text-slate-800">
                    {spjFilterGroup === "panitia-inda" ? "Panitia & Inda" : "Petugas PML & PPL"}
                  </p>
                </div>
                <div className="flex items-center gap-3">
                  <div className="text-right">
                    <p className="text-xs font-semibold text-slate-400">Kelas di dokumen</p>
                    <p className="font-black text-orange-600">
                      {spjFilterGroup === "panitia-inda" ? "-" : (formData.kelas || "—")}
                    </p>
                  </div>
                  <Check size={18} className="text-green-500" />
                </div>
              </motion.div>
            )}
          </>
        );

      // ── SPD ─────────────────────────────────────────────────────────────────
      case "spd":
        return (
          <>
            <div className="grid gap-5 sm:grid-cols-2">
              <div>
                <label className={labelCls}>Nomor Dokumen</label>
                <input
                  className={inputCls}
                  placeholder="Contoh: 100/BPS-3171/2026"
                  value={formData.nomor_dokumen || formData.nomor || ""}
                  onChange={(e) => {
                    update("nomor_dokumen", e.target.value);
                    update("nomor", e.target.value);
                  }}
                />
              </div>
              <div>
                <label className={labelCls}>Tanggal Surat</label>
                <input
                  type="date"
                  className={inputCls}
                  value={formData.tanggal_surat || ""}
                  onChange={(e) => update("tanggal_surat", e.target.value)}
                />
              </div>
            </div>

            <div className="grid gap-5 sm:grid-cols-2">
              <div>
                <label className={labelCls}>Tanggal Awal Kegiatan</label>
                <input
                  type="date"
                  className={inputCls}
                  value={formData.tanggal_awal_kegiatan || ""}
                  onChange={(e) => update("tanggal_awal_kegiatan", e.target.value)}
                />
              </div>
              <div>
                <label className={labelCls}>Tanggal Akhir Kegiatan</label>
                <input
                  type="date"
                  className={inputCls}
                  value={formData.tanggal_akhir_kegiatan || ""}
                  onChange={(e) => update("tanggal_akhir_kegiatan", e.target.value)}
                />
              </div>
            </div>

            <FilterPesertaHotelGelombangPanel
              xlsxLoaded={xlsxLoaded}
              formData={formData}
              setFormData={setFormData}
              petugasData={petugasData}
              onFilterResult={(peserta) => {
                setSpdPeserta(peserta);
                setSpdFiltered(peserta.length > 0);
              }}
            />
          </>
        );

      // ── SURAT TUGAS ─────────────────────────────────────────────────────────
      case "surat-tugas":
        return (
          <>
            <div className="grid gap-5 sm:grid-cols-2">
              <div><label className={labelCls}>Nomor Surat</label><input className={inputCls} placeholder="Contoh: 100/BPS-3171/2026" value={formData.nomor_surat || ""} onChange={(e) => update("nomor_surat", e.target.value)} /></div>
              <div><label className={labelCls}>Tanggal Surat</label><input type="date" className={inputCls} value={formData.tanggal_surat || ""} onChange={(e) => update("tanggal_surat", e.target.value)} /></div>
            </div>
            <div className="grid gap-5 sm:grid-cols-2">
              <div><label className={labelCls}>Tanggal Awal Penyelenggaraan</label><input type="date" className={inputCls} value={formData.tanggal_awal_kegiatan || ""} onChange={(e) => update("tanggal_awal_kegiatan", e.target.value)} /></div>
              <div><label className={labelCls}>Tanggal Akhir Penyelenggaraan</label><input type="date" className={inputCls} value={formData.tanggal_akhir_kegiatan || ""} onChange={(e) => update("tanggal_akhir_kegiatan", e.target.value)} /></div>
            </div>
            <FilterPesertaHotelGelombangPanel
              xlsxLoaded={xlsxLoaded}
              formData={formData}
              setFormData={setFormData}
              petugasData={petugasData}
              onFilterResult={(peserta) => {
                setSuratTugasPeserta(peserta);
                setSuratTugasFiltered(peserta.length > 0);
              }}
            />
          </>
        );

      // ── BAPP ────────────────────────────────────────────────────────────────
      case "bapp": {
        const filteredBappRows = React.useMemo(() => {
          const rows = Array.isArray(bappData) ? bappData : [];
          return rows.filter((row) => isBappRowForRole(row, bappRole));
        }, [bappData, bappRole]);

        const bappOptions = React.useMemo(() => {
          const seen = new Set();
          return filteredBappRows
            .map((row) => {
              const identity = getBappIdentityKey(row, bappRole);
              if (seen.has(identity)) return null;
              seen.add(identity);
              return {
                value: identity,
                label: `${cleanText(row.nama) || "Tanpa Nama"} — ${cleanText(row.jabatan_raw || row.jabatan || bappRole)}`,
                row,
              };
            })
            .filter(Boolean)
            .sort((a, b) => a.label.localeCompare(b.label, "id-ID", { sensitivity: "base" }));
        }, [filteredBappRows, bappRole]);

        return (
          <div className="space-y-5">
            <div className="rounded-3xl border border-orange-100 bg-orange-50/70 p-5">
              <p className="text-xs font-black uppercase tracking-[0.2em] text-orange-700">BAPP PML/PPL</p>
              <p className="mt-2 text-sm font-semibold leading-6 text-slate-600">
                Pilih role, tentukan tanggal surat, lalu unduh dokumen manual atau semua.
              </p>
              <p className="mt-3 text-xs font-bold text-slate-500">
                Data terbaca: {bappData.length} baris dari sheet Pembayaran
              </p>
            </div>

            <div>
              <p className={labelCls}>Pilih Role</p>
              <div className="flex gap-3">
                <button type="button" onClick={() => { setBappRole("PML"); setBappManualSelect(""); }} className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${bappRole === "PML" ? "bg-orange-600 text-white" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>PML</button>
                <button type="button" onClick={() => { setBappRole("PPL"); setBappManualSelect(""); }} className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${bappRole === "PPL" ? "bg-orange-600 text-white" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>PPL</button>
              </div>
            </div>

            {bappRole && (
              <>
                <div>
                  <label className={labelCls}>Tanggal Surat</label>
                  <input
                    type="date"
                    min="2026-07-15"
                    max="2026-08-31"
                    className={inputCls}
                    value={formData.tanggal_surat || ""}
                    onChange={(e) => update("tanggal_surat", e.target.value)}
                  />
                  <p className="mt-1 text-xs font-semibold text-slate-400">Rentang tanggal yang diizinkan: 15 Juli 2026 sampai 31 Agustus 2026.</p>
                </div>

                <div>
                  <label className={labelCls}>Unduh Manual</label>
                  <select value={bappManualSelect} onChange={(e) => setBappManualSelect(e.target.value)} className={inputCls}>
                    <option value="">— Pilih Nama {bappRole} —</option>
                    {bappOptions.map((option) => (
                      <option key={option.value} value={option.value}>{option.label}</option>
                    ))}
                  </select>
                  <p className="mt-1 text-xs font-semibold text-slate-400">Daftar nama diurutkan berdasarkan abjad.</p>
                </div>

                <SelectionUploadPanel
                  selectionRows={bappSelectionRows}
                  onSelectionLoaded={setBappSelectionRows}
                  onClear={() => setBappSelectionRows([])}
                  hint="Pastikan file dan isian file benar. Download template"
                />

                {bappGenerating && (
                  <div className="flex items-center gap-3 rounded-2xl border border-orange-200 bg-orange-50 px-4 py-3 text-sm font-semibold text-orange-700">
                    <LoaderCircle size={18} className="animate-spin" />
                    <span>{bappProgressText || "Sedang menyiapkan file BAPP..."}</span>
                  </div>
                )}

                <div className="grid gap-4 sm:grid-cols-3">
                  <button type="button" onClick={async () => {
                    try {
                      if (!formData.tanggal_surat) throw new Error("Isi tanggal surat terlebih dahulu.");
                      const selectedDate = new Date(formData.tanggal_surat);
                      const minDate = new Date("2026-07-15T00:00:00");
                      const maxDate = new Date("2026-08-31T23:59:59");
                      if (selectedDate < minDate || selectedDate > maxDate) throw new Error("Tanggal surat hanya boleh 15 Juli 2026 sampai 31 Agustus 2026.");
                      if (!bappManualSelect) throw new Error("Pilih nama terlebih dahulu.");
                      const chosenRow = bappOptions.find((option) => option.value === bappManualSelect)?.row;
                      if (!chosenRow) throw new Error("Data nama yang dipilih tidak ditemukan.");
                      setBappGenerating(true);
                      setBappProgressText("Membuat dokumen terpilih...");
                      await generateSingleBapp(bappRole === "PML" ? BAPP_PML_TEMPLATE_URL : BAPP_PPL_TEMPLATE_URL, formData, chosenRow, bappRole, approveByPmlData);
                    } catch (err) { alert(err.message || err); }
                    finally { setBappGenerating(false); setBappProgressText(""); }
                  }} disabled={!bappManualSelect || filteredBappRows.length === 0 || bappGenerating} className="inline-flex items-center justify-center rounded-2xl bg-orange-500 px-5 py-3 text-sm font-black text-white shadow transition hover:bg-orange-600 disabled:cursor-not-allowed disabled:bg-orange-200">
                    Download Terpilih
                  </button>

                  <button type="button" onClick={async () => {
                    try {
                      if (!formData.tanggal_surat) throw new Error("Isi tanggal surat terlebih dahulu.");
                      const selectedDate = new Date(formData.tanggal_surat);
                      const minDate = new Date("2026-07-15T00:00:00");
                      const maxDate = new Date("2026-08-31T23:59:59");
                      if (selectedDate < minDate || selectedDate > maxDate) throw new Error("Tanggal surat hanya boleh 15 Juli 2026 sampai 31 Agustus 2026.");
                      if (bappSelectionRows.length === 0) throw new Error("Unggah file Excel Nama/Email terlebih dahulu.");
                      const keySet = buildSelectionKeySet(bappSelectionRows);
                      const matchedRows = filteredBappRows.filter((row) => rowMatchesSelection(keySet, row.nama, row.email));
                      if (matchedRows.length === 0) throw new Error("Tidak ada data yang cocok dengan file yang diunggah.");
                      setBappGenerating(true);
                      setBappProgressText(`Menyiapkan ${matchedRows.length} dokumen terpilih...`);
                      await generateBapp(bappRole === "PML" ? BAPP_PML_TEMPLATE_URL : BAPP_PPL_TEMPLATE_URL, formData, matchedRows, bappRole, ({ batchIndex, totalBatches }) => {
                        setBappProgressText(`Membuat batch ${batchIndex} dari ${totalBatches}...`);
                      }, approveByPmlData);
                    } catch (err) { alert(err.message || err); }
                    finally { setBappGenerating(false); setBappProgressText(""); }
                  }} disabled={filteredBappRows.length === 0 || bappGenerating || bappSelectionRows.length === 0} className="inline-flex items-center justify-center rounded-2xl border border-orange-200 bg-white px-5 py-3 text-sm font-black text-orange-700 shadow transition hover:bg-orange-50 disabled:cursor-not-allowed disabled:opacity-60">
                    Download Beberapa
                  </button>

                  <button type="button" onClick={async () => {
                    try {
                      if (!formData.tanggal_surat) throw new Error("Isi tanggal surat terlebih dahulu.");
                      const selectedDate = new Date(formData.tanggal_surat);
                      const minDate = new Date("2026-07-15T00:00:00");
                      const maxDate = new Date("2026-08-31T23:59:59");
                      if (selectedDate < minDate || selectedDate > maxDate) throw new Error("Tanggal surat hanya boleh 15 Juli 2026 sampai 31 Agustus 2026.");
                      setBappGenerating(true);
                      setBappProgressText("Mempersiapkan batch download...");
                      await generateBapp(bappRole === "PML" ? BAPP_PML_TEMPLATE_URL : BAPP_PPL_TEMPLATE_URL, formData, filteredBappRows, bappRole, ({ batchIndex, totalBatches }) => {
                        setBappProgressText(`Membuat batch ${batchIndex} dari ${totalBatches}...`);
                      }, approveByPmlData);
                    } catch (err) { alert(err.message || err); }
                    finally { setBappGenerating(false); setBappProgressText(""); }
                  }} disabled={filteredBappRows.length === 0 || bappGenerating} className="inline-flex items-center justify-center rounded-2xl border border-orange-200 bg-white px-5 py-3 text-sm font-black text-orange-700 shadow transition hover:bg-orange-50 disabled:cursor-not-allowed disabled:opacity-60">
                    Download Semua
                  </button>
                </div>
              </>
            )}
          </div>
        );
      }

      // ── BAST ────────────────────────────────────────────────────────────────
      case "bast": {
        const rows = lampiranData || [];
        const isPml = bastRole === "PML";

        const nikLookup = React.useMemo(() => {
          const map = new Map();
          for (const p of petugasData || []) {
            const nama = upperText(p.nama);
            if (nama && !map.has(nama)) map.set(nama, cleanText(p.nik));
          }
          return map;
        }, [petugasData]);

        const bastOptions = React.useMemo(() => {
          if (!bastRole) return [];
          const map = new Map();
          for (const r of rows) {
            const nama = cleanText(isPml ? r.nama_pml : r.nama_ppl);
            if (!nama) continue;
            const email = cleanText(isPml ? r.email_pengawas : r.email_pencacah) || "";
            const identity = email ? upperText(email) : `NAMA::${upperText(nama)}`;
            if (!map.has(identity)) map.set(identity, { value: identity, label: nama, name: nama, rows: [] });
            map.get(identity).rows.push(r);
          }
          return [...map.values()].sort((a, b) => a.name.localeCompare(b.name, "id-ID", { sensitivity: "base" }));
        }, [rows, bastRole, isPml]);

        const filteredBastRows = React.useMemo(() => {
          if (!bastRole) return [];
          return rows.filter((r) => cleanText(isPml ? r.nama_pml : r.nama_ppl));
        }, [rows, bastRole, isPml]);

        return (
          <div className="space-y-5">
            <div className="rounded-3xl border border-orange-100 bg-orange-50/70 p-5">
              <p className="text-xs font-black uppercase tracking-[0.2em] text-orange-700">BAST PML/PPL</p>
              <p className="mt-2 text-sm font-semibold leading-6 text-slate-600">
                Pilih role, tentukan tanggal surat, lalu unduh dokumen manual atau semua.
              </p>
              <p className="mt-3 text-xs font-bold text-slate-500">
                Data terbaca: {rows.length} baris dari sheet Lampiran
              </p>
            </div>

            <div>
              <p className={labelCls}>Pilih Role</p>
              <div className="flex gap-3">
                <button type="button" onClick={() => { setBastRole("PML"); setBastManualSelect(""); }} className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${bastRole === "PML" ? "bg-orange-600 text-white" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>PML</button>
                <button type="button" onClick={() => { setBastRole("PPL"); setBastManualSelect(""); }} className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${bastRole === "PPL" ? "bg-orange-600 text-white" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>PPL</button>
              </div>
            </div>

            {bastRole && (
              <>
                <div>
                  <label className={labelCls}>Tanggal Surat</label>
                  <input
                    type="date"
                    min="2026-07-15"
                    max="2026-08-31"
                    className={inputCls}
                    value={formData.tanggal_surat || ""}
                    onChange={(e) => update("tanggal_surat", e.target.value)}
                  />
                  <p className="mt-1 text-xs font-semibold text-slate-400">Rentang tanggal yang diizinkan: 15 Juli 2026 sampai 31 Agustus 2026.</p>
                </div>

                <div>
                  <label className={labelCls}>Unduh Manual</label>
                  <select value={bastManualSelect} onChange={(e) => setBastManualSelect(e.target.value)} className={inputCls}>
                    <option value="">— Pilih Nama {bastRole} —</option>
                    {bastOptions.map((option) => (
                      <option key={option.value} value={option.value}>{option.label}</option>
                    ))}
                  </select>
                  <p className="mt-1 text-xs font-semibold text-slate-400">Daftar nama diurutkan berdasarkan abjad.</p>
                </div>

                <SelectionUploadPanel
                  selectionRows={bastSelectionRows}
                  onSelectionLoaded={setBastSelectionRows}
                  onClear={() => setBastSelectionRows([])}
                  hint={`Kolom Nama dan/atau Email dicocokkan dengan nama ${isPml ? "Pengawas & Email Pengawas" : "Pencacah & Email Pencacah"}.`}
                />

                {bastGenerating && (
                  <div className="flex items-center gap-3 rounded-2xl border border-orange-200 bg-orange-50 px-4 py-3 text-sm font-semibold text-orange-700">
                    <LoaderCircle size={18} className="animate-spin" />
                    <span>{bastProgressText || "Sedang menyiapkan file BAST..."}</span>
                  </div>
                )}

                <div className="grid gap-4 sm:grid-cols-3">
                  <button type="button" onClick={async () => {
                    try {
                      if (!formData.tanggal_surat) throw new Error("Isi tanggal surat terlebih dahulu.");
                      const selectedDate = new Date(formData.tanggal_surat);
                      const minDate = new Date("2026-07-15T00:00:00");
                      const maxDate = new Date("2026-08-31T23:59:59");
                      if (selectedDate < minDate || selectedDate > maxDate) throw new Error("Tanggal surat hanya boleh 15 Juli 2026 sampai 31 Agustus 2026.");
                      if (!bastManualSelect) throw new Error("Pilih nama terlebih dahulu.");
                      const chosen = bastOptions.find((option) => option.value === bastManualSelect);
                      if (!chosen) throw new Error("Data nama yang dipilih tidak ditemukan.");
                      setBastGenerating(true);
                      setBastProgressText("Membuat dokumen terpilih...");
                      await generateSingleBast(bastRole === "PML" ? BAST_PML_TEMPLATE_URL : BAST_PPL_TEMPLATE_URL, formData, chosen.rows, bastRole, nikLookup, chosen.name);
                    } catch (err) { alert(err.message || err); }
                    finally { setBastGenerating(false); setBastProgressText(""); }
                  }} disabled={!bastManualSelect || filteredBastRows.length === 0 || bastGenerating} className="inline-flex items-center justify-center rounded-2xl bg-orange-500 px-5 py-3 text-sm font-black text-white shadow transition hover:bg-orange-600 disabled:cursor-not-allowed disabled:bg-orange-200">
                    Download Terpilih
                  </button>

                  <button type="button" onClick={async () => {
                    try {
                      if (!formData.tanggal_surat) throw new Error("Isi tanggal surat terlebih dahulu.");
                      const selectedDate = new Date(formData.tanggal_surat);
                      const minDate = new Date("2026-07-15T00:00:00");
                      const maxDate = new Date("2026-08-31T23:59:59");
                      if (selectedDate < minDate || selectedDate > maxDate) throw new Error("Tanggal surat hanya boleh 15 Juli 2026 sampai 31 Agustus 2026.");
                      if (bastSelectionRows.length === 0) throw new Error("Unggah file Excel Nama/Email terlebih dahulu.");
                      const keySet = buildSelectionKeySet(bastSelectionRows);
                      const matchedRows = filteredBastRows.filter((row) => rowMatchesSelection(keySet, isPml ? row.nama_pml : row.nama_ppl, isPml ? row.email_pengawas : row.email_pencacah));
                      if (matchedRows.length === 0) throw new Error("Tidak ada data yang cocok dengan file yang diunggah.");
                      setBastGenerating(true);
                      setBastProgressText(`Menyiapkan dokumen terpilih dari file...`);
                      await generateBast(bastRole === "PML" ? BAST_PML_TEMPLATE_URL : BAST_PPL_TEMPLATE_URL, formData, matchedRows, bastRole, nikLookup, ({ batchIndex, totalBatches }) => {
                        setBastProgressText(`Membuat batch ${batchIndex} dari ${totalBatches}...`);
                      });
                    } catch (err) { alert(err.message || err); }
                    finally { setBastGenerating(false); setBastProgressText(""); }
                  }} disabled={filteredBastRows.length === 0 || bastGenerating || bastSelectionRows.length === 0} className="inline-flex items-center justify-center rounded-2xl border border-orange-200 bg-white px-5 py-3 text-sm font-black text-orange-700 shadow transition hover:bg-orange-50 disabled:cursor-not-allowed disabled:opacity-60">
                    Download Beberapa
                  </button>

                  <button type="button" onClick={async () => {
                    try {
                      if (!formData.tanggal_surat) throw new Error("Isi tanggal surat terlebih dahulu.");
                      const selectedDate = new Date(formData.tanggal_surat);
                      const minDate = new Date("2026-07-15T00:00:00");
                      const maxDate = new Date("2026-08-31T23:59:59");
                      if (selectedDate < minDate || selectedDate > maxDate) throw new Error("Tanggal surat hanya boleh 15 Juli 2026 sampai 31 Agustus 2026.");
                      setBastGenerating(true);
                      setBastProgressText("Mempersiapkan batch download...");
                      await generateBast(bastRole === "PML" ? BAST_PML_TEMPLATE_URL : BAST_PPL_TEMPLATE_URL, formData, filteredBastRows, bastRole, nikLookup, ({ batchIndex, totalBatches }) => {
                        setBastProgressText(`Membuat batch ${batchIndex} dari ${totalBatches}...`);
                      });
                    } catch (err) { alert(err.message || err); }
                    finally { setBastGenerating(false); setBastProgressText(""); }
                  }} disabled={filteredBastRows.length === 0 || bastGenerating} className="inline-flex items-center justify-center rounded-2xl border border-orange-200 bg-white px-5 py-3 text-sm font-black text-orange-700 shadow transition hover:bg-orange-50 disabled:cursor-not-allowed disabled:opacity-60">
                    Download Semua
                  </button>
                </div>
              </>
            )}
          </div>
        );
      }

      // ── SURAT PERNYATAAN PENYELESAIAN LAPANGAN ───────────────────────────
      case "surat-pernyataan-penyelesaian-lapangan": {
        const filteredSuratPenyelesaianLapanganRows = React.useMemo(() => {
          const rows = Array.isArray(bappData) ? bappData : [];
          return rows.filter((row) => isBappRowForRole(row, "PML"));
        }, [bappData]);

        const suratPenyelesaianLapanganOptions = React.useMemo(() => {
          const seen = new Set();
          return filteredSuratPenyelesaianLapanganRows
            .map((row) => {
              const identity = getBappIdentityKey(row, "PML");
              if (seen.has(identity)) return null;
              seen.add(identity);
              return {
                value: identity,
                label: `${cleanText(row.nama) || "Tanpa Nama"} — ${cleanText(row.jabatan_raw || row.jabatan || "PML")}`,
                row,
              };
            })
            .filter(Boolean)
            .sort((a, b) => a.label.localeCompare(b.label, "id-ID", { sensitivity: "base" }));
        }, [filteredSuratPenyelesaianLapanganRows]);

        return (
          <div className="space-y-5">
            <div className="rounded-3xl border border-orange-100 bg-orange-50/70 p-5">
              <p className="text-xs font-black uppercase tracking-[0.2em] text-orange-700">Surat Pernyataan Penyelesaian Lapangan</p>
              <p className="mt-2 text-sm font-semibold leading-6 text-slate-600">
                Pilih nama untuk unduh manual, atau unduh semua dalam batch.
              </p>
              <p className="mt-3 text-xs font-bold text-slate-500">
                Data terbaca: {bappData.length} baris dari sheet Pembayaran
              </p>
            </div>

            <div>
              <label className={labelCls}>Unduh Manual</label>
              <select value={suratPenyelesaianLapanganSelect} onChange={(e) => setSuratPenyelesaianLapanganSelect(e.target.value)} className={inputCls}>
                <option value="">— Pilih Nama PML —</option>
                {suratPenyelesaianLapanganOptions.map((option) => (
                  <option key={option.value} value={option.value}>{option.label}</option>
                ))}
              </select>
              <p className="mt-1 text-xs font-semibold text-slate-400">Daftar nama diurutkan berdasarkan abjad dan memakai email sebagai kunci unik.</p>
            </div>

            <SelectionUploadPanel
              selectionRows={suratPenyelesaianLapanganSelectionRows}
              onSelectionLoaded={setSuratPenyelesaianLapanganSelectionRows}
              onClear={() => setSuratPenyelesaianLapanganSelectionRows([])}
              hint="Kolom Nama dan/atau Email dicocokkan dengan data PML pada sheet Pembayaran."
            />

            {suratPenyelesaianLapanganGenerating && (
              <div className="flex items-center gap-3 rounded-2xl border border-orange-200 bg-orange-50 px-4 py-3 text-sm font-semibold text-orange-700">
                <LoaderCircle size={18} className="animate-spin" />
                <span>{suratPenyelesaianLapanganProgressText || "Sedang menyiapkan dokumen..."}</span>
              </div>
            )}

            <div className="grid gap-4 sm:grid-cols-3">
              <button type="button" onClick={async () => {
                try {
                  if (!suratPenyelesaianLapanganSelect) throw new Error("Pilih nama terlebih dahulu.");
                  const chosenRow = suratPenyelesaianLapanganOptions.find((option) => option.value === suratPenyelesaianLapanganSelect)?.row;
                  if (!chosenRow) throw new Error("Data nama yang dipilih tidak ditemukan.");
                  setSuratPenyelesaianLapanganGenerating(true);
                  setSuratPenyelesaianLapanganProgressText("Membuat dokumen terpilih...");
                  await generateSingleSuratPernyataanPenyelesaianLapangan(SURAT_PERNYATAAN_PENYELESAIAN_LAPANGAN_TEMPLATE_URL, chosenRow);
                } catch (err) { alert(err.message || err); }
                finally { setSuratPenyelesaianLapanganGenerating(false); setSuratPenyelesaianLapanganProgressText(""); }
              }} disabled={!suratPenyelesaianLapanganSelect || filteredSuratPenyelesaianLapanganRows.length === 0 || suratPenyelesaianLapanganGenerating} className="inline-flex items-center justify-center rounded-2xl bg-orange-500 px-5 py-3 text-sm font-black text-white shadow transition hover:bg-orange-600 disabled:cursor-not-allowed disabled:bg-orange-200">
                Download Terpilih
              </button>

              <button type="button" onClick={async () => {
                try {
                  if (suratPenyelesaianLapanganSelectionRows.length === 0) throw new Error("Unggah file Excel Nama/Email terlebih dahulu.");
                  const keySet = buildSelectionKeySet(suratPenyelesaianLapanganSelectionRows);
                  const matchedRows = filteredSuratPenyelesaianLapanganRows.filter((row) => rowMatchesSelection(keySet, row.nama, row.email));
                  if (matchedRows.length === 0) throw new Error("Tidak ada data yang cocok dengan file yang diunggah.");
                  setSuratPenyelesaianLapanganGenerating(true);
                  setSuratPenyelesaianLapanganProgressText(`Menyiapkan ${matchedRows.length} dokumen terpilih...`);
                  await generateSuratPernyataanPenyelesaianLapangan(SURAT_PERNYATAAN_PENYELESAIAN_LAPANGAN_TEMPLATE_URL, matchedRows, ({ batchIndex, totalBatches }) => {
                    setSuratPenyelesaianLapanganProgressText(`Membuat batch ${batchIndex} dari ${totalBatches}...`);
                  });
                } catch (err) { alert(err.message || err); }
                finally { setSuratPenyelesaianLapanganGenerating(false); setSuratPenyelesaianLapanganProgressText(""); }
              }} disabled={filteredSuratPenyelesaianLapanganRows.length === 0 || suratPenyelesaianLapanganGenerating || suratPenyelesaianLapanganSelectionRows.length === 0} className="inline-flex items-center justify-center rounded-2xl border border-orange-200 bg-white px-5 py-3 text-sm font-black text-orange-700 shadow transition hover:bg-orange-50 disabled:cursor-not-allowed disabled:opacity-60">
                Download Beberapa
              </button>

              <button type="button" onClick={async () => {
                try {
                  setSuratPenyelesaianLapanganGenerating(true);
                  setSuratPenyelesaianLapanganProgressText("Mempersiapkan batch download...");
                  await generateSuratPernyataanPenyelesaianLapangan(SURAT_PERNYATAAN_PENYELESAIAN_LAPANGAN_TEMPLATE_URL, filteredSuratPenyelesaianLapanganRows, ({ batchIndex, totalBatches }) => {
                    setSuratPenyelesaianLapanganProgressText(`Membuat batch ${batchIndex} dari ${totalBatches}...`);
                  });
                } catch (err) { alert(err.message || err); }
                finally { setSuratPenyelesaianLapanganGenerating(false); setSuratPenyelesaianLapanganProgressText(""); }
              }} disabled={filteredSuratPenyelesaianLapanganRows.length === 0 || suratPenyelesaianLapanganGenerating} className="inline-flex items-center justify-center rounded-2xl border border-orange-200 bg-white px-5 py-3 text-sm font-black text-orange-700 shadow transition hover:bg-orange-50 disabled:cursor-not-allowed disabled:opacity-60">
                Download Semua
              </button>
            </div>
          </div>
        );
      }

      // ── LAMPIRAN ────────────────────────────────────────────────────────────
      case "lampiran": {
        const rows = lampiranData || [];

        const generateLampiranPml = async () => {
          if (!xlsxLoaded) {
            throw new Error("Data XLSX/Google Sheet belum dimuat.");
          }
          if (rows.length === 0) {
            throw new Error("Sheet Lampiran belum terbaca atau kosong. Pastikan Google Sheet dibaca sebagai XLSX dan nama tab adalah Lampiran.");
          }
          await generateLampiran(LAMPIRAN_PML_TEMPLATE_URL, {}, rows, "PML");
        };

        const generateLampiranPpl = async () => {
          if (!xlsxLoaded) {
            throw new Error("Data XLSX/Google Sheet belum dimuat.");
          }
          if (rows.length === 0) {
            throw new Error("Sheet Lampiran belum terbaca atau kosong. Pastikan Google Sheet dibaca sebagai XLSX dan nama tab adalah Lampiran.");
          }
          await generateLampiran(LAMPIRAN_PPL_TEMPLATE_URL, {}, rows, "PPL");
        };

        return (
          <div className="space-y-5">
            <div className="rounded-3xl border border-orange-100 bg-orange-50/70 p-5">
              <p className="text-xs font-black uppercase tracking-[0.2em] text-orange-700">Generate Lampiran</p>
              <p className="mt-2 text-sm font-semibold leading-6 text-slate-600">
                Pilih jenis lampiran
              </p>
              <p className="mt-3 text-xs font-bold text-slate-500">
                Status data: {xlsxLoaded ? `${rows.length} baris Lampiran terbaca` : "data belum dimuat"}
              </p>
            </div>

            {/* Manual generate: combined dropdown PML + PPL */}
            <div className="rounded-2xl border border-orange-100 bg-white p-4">
              <p className="mb-2 text-xs font-black uppercase tracking-[0.2em] text-orange-700">Generate Manual (PML &amp; PPL)</p>
              <div className="flex flex-col sm:flex-row sm:items-center gap-2">
                <select value={lampiranManualSelect} onChange={(e) => setLampiranManualSelect(e.target.value)}
                  className="w-full sm:w-96 rounded-2xl border border-orange-100 bg-white/80 px-3 py-2 text-sm font-semibold text-slate-800 shadow-sm outline-none">
                  <option value="">— Pilih Petugas (PML / PPL) —</option>
                  {(() => {
                    const map = new Map();
                    for (const r of rows) {
                      // PML
                      const namePml = cleanText(r.nama_pml);
                      if (namePml) {
                        const email = cleanText(r.email_pengawas) || "";
                        const id = email ? upperText(email) : `NAMA::${upperText(namePml)}`;
                        const key = `PML::${id}`;
                        if (!map.has(key)) map.set(key, { value: key, label: `${namePml} (PML)` , name: namePml });
                      }
                      // PPL
                      const namePpl = cleanText(r.nama_ppl);
                      if (namePpl) {
                        const email = cleanText(r.email_pencacah) || "";
                        const id = email ? upperText(email) : `NAMA::${upperText(namePpl)}`;
                        const key = `PPL::${id}`;
                        if (!map.has(key)) map.set(key, { value: key, label: `${namePpl} (PPL)`, name: namePpl });
                      }
                    }
                    return Array.from(map.values()).sort((a, b) => a.name.localeCompare(b.name, 'id-ID', { sensitivity: 'base' })).map(o => (
                      <option key={o.value} value={o.value}>{o.label}</option>
                    ));
                  })()}
                </select>

                <button onClick={async () => {
                  try {
                    if (!lampiranManualSelect) { alert('Pilih petugas terlebih dahulu.'); return; }
                    if (!xlsxLoaded) throw new Error('Data belum dimuat');
                    const sepIndex = lampiranManualSelect.indexOf('::');
                    const role = lampiranManualSelect.slice(0, sepIndex);
                    const id = lampiranManualSelect.slice(sepIndex + 2);
                    const isPml = role === 'PML';
                    const filtered = rows.filter((r) => {
                      const name = isPml ? cleanText(r.nama_pml) : cleanText(r.nama_ppl);
                      const email = isPml ? cleanText(r.email_pengawas) || "" : cleanText(r.email_pencacah) || "";
                      const key = email ? upperText(email) : `NAMA::${upperText(name)}`;
                      return key === id;
                    });
                    if (filtered.length === 0) throw new Error('Tidak ada data untuk pilihan ini');
                    const displayName = isPml ? filtered[0].nama_pml : filtered[0].nama_ppl;
                    await generateSingleLampiran(isPml ? LAMPIRAN_PML_TEMPLATE_URL : LAMPIRAN_PPL_TEMPLATE_URL, {}, filtered, isPml ? 'PML' : 'PPL', displayName || 'Tanpa Nama');
                  } catch (err) { alert(err.message || err); }
                }} type="button" className="w-full sm:w-auto inline-flex items-center justify-center gap-2 rounded-2xl bg-orange-500 px-4 py-2 text-sm font-black text-white shadow transition hover:bg-orange-600">Generate Terpilih</button>
              </div>
            </div>

            {/* Download Beberapa: unggah Excel Nama/Email lalu generate untuk PML dan/atau PPL sekaligus */}
            <div className="rounded-2xl border border-orange-100 bg-white p-4 space-y-3">
              <p className="text-xs font-black uppercase tracking-[0.2em] text-orange-700">Download Beberapa (dari file Excel Nama/Email)</p>
              <SelectionUploadPanel
                selectionRows={lampiranSelectionRows}
                onSelectionLoaded={setLampiranSelectionRows}
                onClear={() => setLampiranSelectionRows([])}
                hint="Silahkan upload file excel sesuai template. Download template disini"
              />

              {lampiranBeberapaGenerating && (
                <div className="flex items-center gap-3 rounded-2xl border border-orange-200 bg-orange-50 px-4 py-3 text-sm font-semibold text-orange-700">
                  <LoaderCircle size={18} className="animate-spin" />
                  <span>{lampiranBeberapaProgressText || "Sedang menyiapkan dokumen..."}</span>
                </div>
              )}

              <div className="grid gap-4 sm:grid-cols-2">
                <button type="button" disabled={lampiranSelectionRows.length === 0 || lampiranBeberapaGenerating} onClick={async () => {
                  try {
                    if (!xlsxLoaded) throw new Error("Data XLSX/Google Sheet belum dimuat.");
                    if (lampiranSelectionRows.length === 0) throw new Error("Unggah file Excel Nama/Email terlebih dahulu.");
                    const keySet = buildSelectionKeySet(lampiranSelectionRows);
                    const matchedRows = rows.filter((r) => rowMatchesSelection(keySet, r.nama_pml, r.email_pengawas));
                    if (matchedRows.length === 0) throw new Error("Tidak ada PML yang cocok dengan file yang diunggah.");
                    setLampiranBeberapaGenerating(true);
                    setLampiranBeberapaProgressText("Menyiapkan Lampiran PML terpilih...");
                    await generateLampiran(LAMPIRAN_PML_TEMPLATE_URL, {}, matchedRows, "PML");
                  } catch (err) { alert(err.message || err); }
                  finally { setLampiranBeberapaGenerating(false); setLampiranBeberapaProgressText(""); }
                }} className="inline-flex items-center justify-center rounded-2xl bg-orange-500 px-5 py-3 text-sm font-black text-white shadow transition hover:bg-orange-600 disabled:cursor-not-allowed disabled:bg-orange-200">
                  Download Beberapa PML
                </button>

                <button type="button" disabled={lampiranSelectionRows.length === 0 || lampiranBeberapaGenerating} onClick={async () => {
                  try {
                    if (!xlsxLoaded) throw new Error("Data XLSX/Google Sheet belum dimuat.");
                    if (lampiranSelectionRows.length === 0) throw new Error("Unggah file Excel Nama/Email terlebih dahulu.");
                    const keySet = buildSelectionKeySet(lampiranSelectionRows);
                    const matchedRows = rows.filter((r) => rowMatchesSelection(keySet, r.nama_ppl, r.email_pencacah));
                    if (matchedRows.length === 0) throw new Error("Tidak ada PPL yang cocok dengan file yang diunggah.");
                    setLampiranBeberapaGenerating(true);
                    setLampiranBeberapaProgressText("Menyiapkan Lampiran PPL terpilih...");
                    await generateLampiran(LAMPIRAN_PPL_TEMPLATE_URL, {}, matchedRows, "PPL");
                  } catch (err) { alert(err.message || err); }
                  finally { setLampiranBeberapaGenerating(false); setLampiranBeberapaProgressText(""); }
                }} className="inline-flex items-center justify-center rounded-2xl border border-orange-200 bg-white px-5 py-3 text-sm font-black text-orange-700 shadow transition hover:bg-orange-50 disabled:cursor-not-allowed disabled:opacity-60">
                  Download Beberapa PPL
                </button>
              </div>
            </div>

            <div className="grid gap-4 sm:grid-cols-2">
              <div className="rounded-3xl border border-orange-100 bg-white p-5 shadow-sm">
                <div className="mb-4 flex h-12 w-12 items-center justify-center rounded-2xl bg-orange-100 text-orange-600">
                  <Users size={22} />
                </div>
                <h3 className="text-lg font-black text-slate-900">Lampiran PML</h3>
                <p className="mt-1 text-sm font-semibold leading-6 text-slate-500">
                  Format lampiran dengan kolom No, Nama Petugas Lapangan Sensus, Kecamatan/Distrik, Desa/Kampung/Nagari, dan Jumlah SLS/Sub-SLS.
                </p>
                <div className="mt-5">
                  <GenerateDocxButton label="Generate Lampiran PML" onGenerate={generateLampiranPml} />
                </div>
              </div>

              <div className="rounded-3xl border border-orange-100 bg-white p-5 shadow-sm">
                <div className="mb-4 flex h-12 w-12 items-center justify-center rounded-2xl bg-orange-100 text-orange-600">
                  <FileText size={22} />
                </div>
                <h3 className="text-lg font-black text-slate-900">Lampiran PPL</h3>
                <p className="mt-1 text-sm font-semibold leading-6 text-slate-500">
                  Format lampiran dengan kolom No, Kecamatan/Distrik, Desa/Kampung/Nagari, dan Jumlah SLS/Sub-SLS.
                </p>
                <div className="mt-5">
                  <GenerateDocxButton label="Generate Lampiran PPL" onGenerate={generateLampiranPpl} />
                </div>
              </div>
            </div>

            <div className="rounded-2xl border border-orange-100 bg-white/75 p-4 text-sm font-semibold text-slate-600">
              <p className="font-bold text-slate-800">Statistik unik email</p>
              <p className="mt-1">PML unik (email): {new Set(rows.map(r => upperText(cleanText(r.email_pengawas))).filter(Boolean)).size}</p>
              <p className="mt-1">PPL unik (email): {new Set(rows.map(r => upperText(cleanText(r.email_pencacah))).filter(Boolean)).size}</p>
            </div>
          </div>
        );
      }

      case "surat-kepala": {
        return (
          <div className="space-y-5">
            <div className="rounded-3xl border border-orange-100 bg-orange-50/70 p-5">
              <p className="text-xs font-black uppercase tracking-[0.2em] text-orange-700">
                Surat Pernyataan Evaluasi Pelaksanaan Lapangan (Kepala BPS)
              </p>
              <p className="mt-2 text-sm font-semibold leading-6 text-slate-600">
                Unggah Excel berisi kolom Nama dan/atau Email untuk mengisi tabel lampiran.
                Satu file akan berisi semua nama yang cocok: PML tampil lebih dulu (A-Z),
                lalu PPL (A-Z). Target Prelist &amp; Jabatan diambil dari sheet Pembayaran;
                Realisasi PML dari kolom Realisasi Total, Realisasi PPL dijumlah dari
                sheet Data per SLS.
              </p>
            </div>
      
            <SelectionUploadPanel
              selectionRows={suratKepalaSelectionRows}
              onSelectionLoaded={setSuratKepalaSelectionRows}
              onClear={() => setSuratKepalaSelectionRows([])}
              hint="Kolom Nama dan/atau Email dicocokkan ke sheet Pembayaran."
            />
      
            {suratKepalaGenerating && (
              <div className="flex items-center gap-3 rounded-2xl border border-orange-200 bg-orange-50 px-4 py-3 text-sm font-semibold text-orange-700">
                <LoaderCircle size={18} className="animate-spin" />
                <span>{suratKepalaProgressText || "Sedang membuat surat..."}</span>
              </div>
            )}
      
            <button
              type="button"
              onClick={async () => {
                try {
                  if (suratKepalaSelectionRows.length === 0) {
                    throw new Error("Unggah file Excel Nama/Email terlebih dahulu.");
                  }
                  const { rows, skipped } = buildSuratKepalaRows(
                    suratKepalaSelectionRows,
                    bappData,
                    dataPerSlsData,
                    lampiranData,
                    statusSlsData,
                    approveByPmlData,
                    dataPmlProgressData
                  );
                  if (rows.length === 0) {
                    throw new Error("Tidak ada baris yang cocok untuk dimasukkan ke lampiran.");
                  }
                  setSuratKepalaGenerating(true);
                  setSuratKepalaProgressText(`Menyiapkan surat untuk ${rows.length} petugas...`);
                  await generateSuratKepala(SURAT_KEPALA_TEMPLATE_URL, rows);
                  if (skipped.length > 0) {
                    alert(
                      `Surat berhasil dibuat. ${skipped.length} baris dilewati:\n\n` +
                      skipped.map((s) => `- ${s.nama || s.email}: ${s.alasan}`).join("\n")
                    );
                  }
                } catch (err) {
                  alert(err.message || err);
                } finally {
                  setSuratKepalaGenerating(false);
                  setSuratKepalaProgressText("");
                }
              }}
              disabled={suratKepalaSelectionRows.length === 0 || suratKepalaGenerating}
              className="inline-flex items-center justify-center gap-2 rounded-2xl bg-orange-500 px-5 py-3 text-sm font-black text-white shadow transition hover:bg-orange-600 disabled:cursor-not-allowed disabled:bg-orange-200"
            >
              <Download size={16} /> Generate Surat Kepala
            </button>
          </div>
        );
      }

      default:
        return null;
    }
  };

  return (
    <form onSubmit={handleSubmit} className="space-y-5 rounded-[2.5rem] border border-orange-100 bg-white/80 p-6 shadow-xl shadow-orange-900/5 backdrop-blur md:p-10">
      {renderFields()}
      {docType.id !== "lampiran" && docType.id !== "bapp" && docType.id !== "bast" && docType.id !== "surat-pernyataan-penyelesaian-lapangan" && docType.id !== "gabungan-pembayaran" && (
        <div className="border-t border-orange-100 pt-5">
          <button type="submit" className="group inline-flex items-center gap-2 rounded-2xl bg-orange-500 px-7 py-4 font-black text-white shadow-2xl shadow-orange-500/25 transition hover:-translate-y-1 hover:bg-orange-600">
            Pratinjau Dokumen <ChevronRight className="transition group-hover:translate-x-1" size={18} />
          </button>
        </div>
      )}
    </form>
  );
}

// BAST

