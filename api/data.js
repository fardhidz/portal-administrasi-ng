// ============================================================
// /api/data — dibaca oleh App.jsx saat halaman pertama kali dibuka.
// Isinya JSON kecil hasil sinkronisasi terakhir (lihat /api/sync),
// dibaca langsung dari Supabase (tabel portal_cache), jadi client
// TIDAK perlu lagi download+parse xlsx 9-10MB tiap load.
// ============================================================

import { getSupabaseAdmin, CACHE_TABLE, CACHE_KEY } from "./_supabaseAdmin.js";

export default async function handler(req, res) {
  try {
    const supabase = getSupabaseAdmin();
    const { data: row, error } = await supabase
      .from(CACHE_TABLE)
      .select("payload, synced_at")
      .eq("key", CACHE_KEY)
      .maybeSingle();

    if (error) throw new Error(error.message);

    if (!row) {
      res.status(404).json({
        error: "Belum ada data cache. Jalankan /api/sync dulu (otomatis via cron harian, atau tombol 'Sinkronkan Sekarang').",
      });
      return;
    }

    // Cache di edge/CDN selama 5 menit, boleh sajikan versi basi sampai
    // 10 menit sambil revalidate di background — data ini toh cuma
    // berubah lewat sync harian, tidak perlu selalu-fresh per request.
    res.setHeader("Cache-Control", "s-maxage=300, stale-while-revalidate=600");
    res.status(200).json(row.payload);
  } catch (err) {
    console.error("Gagal memuat /api/data:", err);
    res.status(500).json({ error: err.message });
  }
}
