# Peminjaman Ruangan

Aplikasi peminjaman ruangan berbasis Google Apps Script dengan Google Sheets sebagai database dan notifikasi di dalam aplikasi.

## File

- `index.html` — tampilan aplikasi live dengan Bootstrap 5.
- `Code.gs` — penghubung aplikasi ke 3 tab Google Sheets.

## Struktur Google Sheets

Nama tab dan urutan kolom harus persis:

1. `Pengguna`: Email | Nama | Role | Status | TanggalDaftar | Dept
2. `Ruangan`: ID | NamaRuangan | Status | Deskripsi
3. `Peminjaman`: ID | EmailPeminjam | NamaPeminjam | IDRuangan | NamaRuangan | Tanggal | WaktuMulai | WaktuSelesai | Tujuan | JumlahPeserta | PenyetujuEmail | PenyetujuNama | Status | CatatanPenyetuju | TanggalPengajuan

## Cara memasang

1. Buka Google Sheets, lalu pilih **Ekstensi → Apps Script**.
2. Salin isi `Code.gs` dari repository ini ke file `Code.gs`.
3. Salin isi `index.html` ke file HTML bernama `index`.
4. Pastikan `SPREADSHEET_ID` berisi ID Google Sheets yang benar.
5. Jalankan fungsi `setupDatabase` satu kali.
6. Pilih **Deploy → New deployment → Web app**.
7. Untuk login Google, buka URL Web App Apps Script, bukan halaman GitHub.

## Catatan

- Aplikasi tidak memakai data contoh dan tidak mengirim email.
- Pengguna baru masuk lewat Google Sign-In; email menjadi identitas tetap.
- Pengguna langsung aktif. Penyetuju dan Admin menunggu persetujuan Admin.
- Setiap perubahan kode harus disimpan sebagai **versi deployment baru**.
