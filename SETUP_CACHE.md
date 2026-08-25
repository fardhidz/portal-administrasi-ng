# Setup: Cache Data Server-Side (Supabase + Cron)

Perubahan ini menambahkan lapisan cache di server supaya halaman tidak
lagi harus download + parse file Excel 9-10MB dari Google Sheets setiap
kali dibuka di browser. Alurnya sekarang:

```
Google Sheets (xlsx, ~10MB, banyak sheet)
        │
        │  1x per hari (cron) ATAU saat klik "Sinkronkan Sekarang"
        ▼
/api/sync   → fetch + parse SEMUA sheet di server (Node)
        │      lalu simpan hasilnya sebagai 1 baris JSONB
        ▼
Supabase (tabel portal_cache)
        │
        │  dibaca tiap kali user buka halaman (cepat, KB bukan MB)
        ▼
/api/data   → App.jsx (useEffect awal)
```

Kalau `/api/data` belum ada isinya (misal baru pertama kali deploy dan
`/api/sync` belum pernah jalan), aplikasi otomatis **fallback** ke cara
lama (load langsung dari Google Sheets di browser), jadi tidak akan
blank/error — cuma lebih lambat, seperti sebelumnya.

## Langkah setup di Supabase + Vercel

1. **Buat project Supabase**
   Daftar/login di [supabase.com](https://supabase.com) → **New
   project** (pilih region terdekat, misal Singapore). Tunggu sampai
   project selesai di-provision (~2 menit).

2. **Buat tabel `portal_cache`**
   Di dashboard Supabase → **SQL Editor** → **New query** → tempel isi
   file `supabase/migration.sql` (ada di root repo ini) → **Run**.
   Ini hanya perlu dijalankan **sekali**.

3. **Ambil kredensial API**
   Dashboard Supabase → **Project Settings** → **API**. Catat dua
   nilai ini:
   - **Project URL** → jadi env var `SUPABASE_URL`
   - **service_role key** (bukan `anon` key!) → jadi env var
     `SUPABASE_SERVICE_ROLE_KEY`

   ⚠️ `service_role` key punya akses penuh ke database (bypass Row
   Level Security). Jangan pernah ditaruh di kode frontend (`src/**`)
   atau di-commit ke git — hanya dipakai di `api/**` sebagai
   environment variable server-side di Vercel.

4. **Set environment variables di Vercel**
   Dashboard proyek Vercel → **Settings** → **Environment Variables**
   → tambahkan:
   - `SUPABASE_URL`
   - `SUPABASE_SERVICE_ROLE_KEY`
   - `CRON_SECRET` — string acak minimal 16 karakter, untuk melindungi
     `/api/sync` supaya tidak bisa dipicu sembarang orang dari luar.

5. **(Opsional) Set `GOOGLE_SHEET_URL`**
   Kalau suatu saat sheet sumbernya pindah/ganti, cukup update env var
   ini tanpa perlu redeploy kode.

6. **Deploy.** File `vercel.json` di root sudah berisi jadwal cron:
   ```json
   { "crons": [{ "path": "/api/sync", "schedule": "0 23 * * *" }] }
   ```
   `0 23 * * *` = 23:00 UTC = **06:00 WIB**, dijalankan otomatis oleh
   Vercel Cron 1x sehari (jam pastinya bisa meleset sampai 1 jam — ini
   perilaku normal plan Hobby, bukan bug).

7. **Sync pertama kali secara manual**, karena cron baru jalan besok
   pagi. Buka:
   ```
   https://<domain-kamu>.vercel.app/api/sync?token=<CRON_SECRET>
   ```
   di browser, atau klik tombol **"Sinkronkan Sekarang"** di pojok
   kanan atas dashboard aplikasi (akan minta kode `CRON_SECRET`).

## Kalau Google Sheet diupdate mendadak

Tidak perlu menunggu jadwal cron besok pagi — klik tombol
**"Sinkronkan Sekarang"** di header aplikasi kapan saja. Prosesnya sama
persis dengan yang dijalankan cron, cuma dipicu manual.

## Dev lokal (`npm run dev`)

Vite dev server tidak menjalankan `/api/*` (itu fitur Vercel Functions).
Jadi saat dev lokal, `fetch('/api/data')` akan gagal secara wajar, dan
aplikasi otomatis fallback ke perilaku lama (load langsung dari Google
Sheets). Tidak perlu setup tambahan untuk development.

Untuk test `/api/*` secara lokal, pakai `vercel dev` (perlu install
Vercel CLI: `npm i -g vercel`, lalu `vercel dev`, pastikan `.env` lokal
sudah berisi `SUPABASE_URL` & `SUPABASE_SERVICE_ROLE_KEY`) — ini
menjalankan serverless functions secara lokal juga.

## Kenapa Supabase, bukan bikin tabel per jenis data?

Data dari Google Sheets (petugas, lampiran, pembayaran, status SLS,
dst) disimpan sebagai **satu kolom JSONB** di tabel `portal_cache`,
bukan didesain jadi banyak tabel relasional. Ini sengaja: skema sheet
sumbernya masih sering berubah (kolom nambah/ganti), jadi menyimpan
apa adanya sebagai JSON menghindari perlu migrasi skema database tiap
kali ada perubahan di sheet. Kalau ke depannya mau query/filter data
langsung dari SQL (bukan cuma baca semuanya lewat `/api/data`), tabel
relasional per jenis data bisa dibuat menyusul.
