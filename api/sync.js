// ============================================================
// /api/sync — dipanggil oleh Vercel Cron (harian) ATAU manual
// (tombol "Sinkronkan Sekarang" di dashboard).
//
// Tugasnya: ambil Google Sheet sumber (yang selama ini di-fetch
// langsung dari browser), parse SEMUA sheet di server (Node, bukan
// browser), lalu simpan hasilnya sebagai satu baris JSONB ke tabel
// Supabase (portal_cache). Client (App.jsx) tinggal baca /api/data
// yang jauh lebih kecil & cepat dibanding download+parse xlsx
// 9-10MB di browser.
//
// Auth: pakai env var CRON_SECRET.
//  - Kalau dipanggil oleh Vercel Cron, Vercel OTOMATIS mengirim
//    header "Authorization: Bearer <CRON_SECRET>" -> otomatis lolos.
//  - Kalau dipanggil manual dari tombol di dashboard, kirim
//    query "?token=<CRON_SECRET>".
//  - Kalau CRON_SECRET belum di-set sama sekali di Vercel project
//    settings, endpoint ini dibiarkan terbuka (supaya gampang testing
//    lokal) — tapi sebaiknya SELALU di-set CRON_SECRET di production.
// ============================================================

import { getSupabaseAdmin, CACHE_TABLE, CACHE_KEY } from "./_supabaseAdmin.js";
import {
  normalizeGoogleSheetUrl,
  parseXlsxData,
  parseLampiranXlsxData,
  parseBappData,
  parseStatusSlsData,
  parseDataPerSlsData,
  parseApproveByPmlData,
  parseDataPmlProgressData,
  enrichApproveByPmlWithFotoBukti,
} from "../src/lib/parsers.js";

// URL Google Sheet sumber utama. Bisa dioverride lewat env var
// GOOGLE_SHEET_URL kalau suatu saat sheet-nya pindah/ganti tanpa
// perlu ubah kode.
const DEFAULT_GOOGLE_SHEET_URL =
  "https://docs.google.com/spreadsheets/d/1_tHblm6GA_oO_k2iI8VE14-NDsIt5Lx8t80BqWsS7Ls/edit?usp=sharing";

function isAuthorized(req) {
  const configured = process.env.CRON_SECRET;
  if (!configured) return true; // belum di-set -> terbuka (mode dev)

  const authHeader = req.headers?.authorization || req.headers?.Authorization || "";
  if (authHeader === `Bearer ${configured}`) return true;

  const queryToken = req.query?.token;
  if (queryToken && queryToken === configured) return true;

  return false;
}

export default async function handler(req, res) {
  if (!isAuthorized(req)) {
    res.status(401).json({ error: "Unauthorized. Sertakan ?token=<CRON_SECRET> yang benar." });
    return;
  }

  const startedAt = Date.now();

  try {
    const sourceUrl = process.env.GOOGLE_SHEET_URL || DEFAULT_GOOGLE_SHEET_URL;
    const normalized = normalizeGoogleSheetUrl(sourceUrl);
    if (!normalized) throw new Error("GOOGLE_SHEET_URL tidak valid.");

    const response = await fetch(normalized.exportUrl);
    if (!response.ok) {
      throw new Error(`Gagal mengambil Google Sheet: HTTP ${response.status} ${response.statusText}`);
    }
    const arrayBuffer = await response.arrayBuffer();

    // Parsing berat ini sekarang jalan di server, sekali per sync,
    // bukan di browser tiap kali user buka halaman.
    const data = parseXlsxData(arrayBuffer);
    const lampiran = parseLampiranXlsxData(arrayBuffer);
    const bappData = parseBappData(arrayBuffer);
    const statusSls = parseStatusSlsData(arrayBuffer);
    const dataPerSls = parseDataPerSlsData(arrayBuffer);
    let approveByPml = parseApproveByPmlData(arrayBuffer);
    approveByPml = await enrichApproveByPmlWithFotoBukti(approveByPml);
    const dataPmlProgress = parseDataPmlProgressData(arrayBuffer);

    const syncedAt = new Date().toISOString();
    const payload = {
      data,
      lampiran,
      bappData,
      statusSls,
      dataPerSls,
      approveByPml,
      dataPmlProgress,
      syncedAt,
      sourceUrl,
    };

    const supabase = getSupabaseAdmin();
    // Satu baris per key, ditimpa (upsert) tiap kali sync -> selalu
    // berisi snapshot data yang paling baru, mirip cara kerja Blob
    // sebelumnya tapi sekarang di database.
    const { error: upsertError } = await supabase
      .from(CACHE_TABLE)
      .upsert({ key: CACHE_KEY, payload, synced_at: syncedAt }, { onConflict: "key" });

    if (upsertError) {
      throw new Error(`Gagal menyimpan ke Supabase: ${upsertError.message}`);
    }

    res.status(200).json({
      ok: true,
      syncedAt,
      durationMs: Date.now() - startedAt,
      counts: {
        petugas: data.length,
        lampiran: lampiran.length,
        pembayaran: bappData.length,
        statusSls: statusSls.length,
        dataPerSls: dataPerSls.length,
        approveByPml: approveByPml.length,
        dataPmlProgress: dataPmlProgress.length,
      },
    });
  } catch (err) {
    console.error("Sync gagal:", err);
    res.status(500).json({ ok: false, error: err.message });
  }
}
