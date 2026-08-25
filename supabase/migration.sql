-- Jalankan sekali di Supabase Dashboard > SQL Editor > New query.
-- Membuat tabel penyimpanan hasil sync (pengganti Vercel Blob).

create table if not exists portal_cache (
  key        text primary key,
  payload    jsonb not null,
  synced_at  timestamptz not null default now()
);

-- Row Level Security: kita akses tabel ini HANYA lewat service_role key
-- di server (api/sync.js, api/data.js), tidak pernah langsung dari
-- browser. Jadi RLS diaktifkan tanpa policy sama sekali -> otomatis
-- menolak semua akses lewat anon/public key, dan service_role key
-- selalu bypass RLS.
alter table portal_cache enable row level security;
