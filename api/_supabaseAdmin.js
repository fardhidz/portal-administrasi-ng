// ============================================================
// Client Supabase khusus server (dipakai di api/sync.js & api/data.js).
//
// PENTING: file ini memakai SUPABASE_SERVICE_ROLE_KEY, yang punya akses
// penuh (bypass Row Level Security). JANGAN PERNAH mengimpor file ini
// dari kode yang jalan di browser (src/**) — hanya boleh dipakai di
// dalam folder api/ yang jalan sebagai Vercel Function di server.
// ============================================================

import { createClient } from "@supabase/supabase-js";

let cachedClient = null;

export function getSupabaseAdmin() {
  if (cachedClient) return cachedClient;

  // .trim() untuk jaga-jaga kalau ada spasi/newline tak sengaja ikut
  // ke-paste ke .env (sering terjadi di Windows / copy-paste dari
  // dashboard), yang bikin error "Invalid path specified in request URL".
  const url = process.env.SUPABASE_URL?.trim();
  const serviceRoleKey = process.env.SUPABASE_SERVICE_ROLE_KEY?.trim();

  if (!url || !serviceRoleKey) {
    throw new Error(
      "SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY belum di-set. " +
      "Set di Vercel Dashboard > Settings > Environment Variables (atau .env untuk dev lokal)."
    );
  }

  if (!/^https:\/\/[a-z0-9-]+\.supabase\.co\/?$/i.test(url)) {
    throw new Error(
      `SUPABASE_URL tidak valid: "${url}". Formatnya harus persis ` +
      `https://xxxxxxxx.supabase.co (Project URL dari Settings > API atau ` +
      `Settings > General di dashboard Supabase), tanpa path tambahan.`
    );
  }

  cachedClient = createClient(url, serviceRoleKey, {
    auth: { persistSession: false },
  });
  return cachedClient;
}

// Nama tabel cache di Supabase. Struktur: satu baris = satu snapshot
// hasil sync, disimpan sebagai JSONB supaya tidak perlu redesign skema
// tiap kali ada sheet/kolom baru di Google Sheets sumber.
export const CACHE_TABLE = "portal_cache";
export const CACHE_KEY = "portal-administrasi";
