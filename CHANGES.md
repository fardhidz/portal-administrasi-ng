# Ringkasan Restrukturisasi

`src/App.jsx` (semula ±7.140 baris, semua logic ditumpuk dalam 1 file) dipecah
menjadi modul-modul berikut. **Semua fungsi dipindahkan apa adanya** (isinya
tidak diubah), kecuali 3 perbaikan bug yang tercantum di bawah.

```
src/
  data/
    docTypes.jsx        — daftar jenis dokumen (DOC_TYPES) + ikon
    templates.js         — URL template .docx & filter group jabatan
  lib/
    parsers.js            — semua parser XLSX/Google Sheets, foto bukti
    helpers.js             — util umum (format tanggal, rupiah, dsb.)
    docGenerators.js       — builder data template + generator .docx
    bast.js                 — logic BAST
    berkasPembayaran.js     — logic Berkas Pembayaran PML/PPL
    suratKepala.js           — logic Surat Kepala BPS
  components/
    FormPanels.jsx           — kartu upload, filter peserta, dsb.
    DocForm.jsx               — form utama tiap jenis dokumen
    DocxPreviews.jsx           — komponen pratinjau .docx
    DocPreview.jsx              — dispatcher pratinjau
  App.jsx                      — komponen utama (default export)
```

Cara verifikasi: dependency antar-fungsi dianalisis otomatis dengan Babel
(bukan manual) supaya tidak ada import yang salah/hilang, lalu proyek
di-build (`npm run build`) dan di-lint (`npm run lint`) sampai bersih dari
error struktural. Hasil pemecahan juga di-diff ulang terhadap file asli untuk
memastikan tidak ada baris logic yang hilang atau terduplikasi.

## Bug yang diperbaiki

1. **Tombol "Download Semua" pada fitur BAPP** memanggil variabel
   `matchedRows` yang tidak pernah didefinisikan di scope itu (sisa
   copy-paste dari tombol "Download Beberapa" di sebelahnya). Kalau diklik,
   akan langsung error `ReferenceError` di browser dan gagal generate.
   Diperbaiki jadi memakai `filteredBappRows`, sesuai konteks tombol
   "download semua data". (Tombol serupa di fitur BAST dan Surat Pernyataan
   Penyelesaian Lapangan sudah dicek dan aman, tidak ada bug yang sama.)

2. **Risiko crash saat pindah jenis dokumen lewat sidebar.** Di dalam
   `DocForm`, fungsi `renderFields()` memanggil `React.useMemo` secara
   kondisional (jumlah hook yang terpanggil berbeda tergantung
   `docType.id`). Karena `<DocForm>` sebelumnya tidak diberi `key`, saat
   pengguna berpindah jenis dokumen langsung dari sidebar (tanpa balik ke
   dashboard dulu), React bisa memakai instance komponen yang sama dengan
   jumlah/urutan hook yang berbeda dari render sebelumnya — berpotensi memicu
   error React "Rendered more hooks than during the previous render".
   Diperbaiki dengan menambahkan `key={selectedDoc.id}` pada `<DocForm>` di
   `App.jsx`, sehingga React me-remount komponen secara bersih setiap kali
   jenis dokumen berganti.

3. Rapihan kecil: karakter escape regex yang tidak perlu di
   `sanitizeFileName` (`docGenerators.js`), dan dua konstanta
   (`docxPreviewStyle`, `RENDER_OPTS`) yang hanya dipakai internal di
   `DocxPreviews.jsx` tidak lagi di-export (supaya fast-refresh React bersih
   dari warning, tidak mengubah perilaku).

## Belum diubah (bukan bug, disengaja tidak disentuh)

Lint masih menandai beberapa hal berikut sebagai warning/style, tapi TIDAK
diubah karena berisiko mengubah perilaku tanpa konfirmasi eksplisit:

- Beberapa `useEffect`/`useMemo` dengan dependency array yang tidak lengkap
  (pola ini sudah ada sejak file aslinya).
- Parameter yang tidak dipakai pada `GoogleSheetCard` (komponennya memang
  tampak belum aktif/dipakai penuh).
- Variabel `loadedDataPmlProgress` di `App.jsx` yang di-assign tapi belum
  dipakai lebih lanjut (kelihatannya fitur yang belum selesai diintegrasikan).

Kalau mau saya lanjutkan membenahi hal-hal di atas juga (terutama pola
`useMemo` di dalam `renderFields` yang idealnya dipindah ke atas jadi hook di
level komponen, bukan dipanggil lewat fungsi biasa), tinggal bilang saja.
