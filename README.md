# Peminjaman Ruangan

Aplikasi peminjaman ruangan dengan tampilan Bootstrap 5 dan Google Sheets sebagai database.

## File

- \`index.html\` — tampilan aplikasi dan mode pratinjau.
- \`Code.gs\` — penghubung Google Apps Script ke 3 tab Google Sheets.

## Struktur Google Sheets

Nama tab harus persis:

1. \`Pengguna\`
2. \`Ruangan\`
3. \`Peminjaman\`

Header dan urutan kolom mengikuti rancangan aplikasi.

## Memasang penghubung Google Sheets

1. Buka file Google Sheets.
2. Pilih **Ekstensi → Apps Script**.
3. Buat file kode bernama \`Code.gs\`, lalu salin isi file \`Code.gs\` dari repository ini.
4. Ganti nilai \`SPREADSHEET_ID\` dengan ID Google Sheets milik kamu.
5. Buat file HTML bernama \`index\`, lalu salin isi \`index.html\`.
6. Jalankan fungsi \`setupDatabase\` satu kali untuk memeriksa tab dan mengisi ruangan R002–R006 jika tab Ruangan masih kosong.
7. Pilih **Deploy → New deployment → Web app**.

## Catatan

- Jangan menghapus atau mengubah nama kolom.
- Jangan menaruh password atau token di repository.
- Google Sign-In, data live, dan deployment akan diaktifkan setelah penghubung Apps Script selesai dipasang.
- Setiap perubahan kode perlu dibuat sebagai deployment baru.
