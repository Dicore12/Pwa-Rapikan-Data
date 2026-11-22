# Rapikan Data (PWA)

Aplikasi web untuk merapikan file Excel/CSV, mengurutkan, memfilter, mengelompokkan, dan membuat pivot.
Semua diproses di browser (tidak ada server), dan bisa dipasang seperti aplikasi serta dipakai offline.

## Menjalankan lokal
```
cd rapikan-data
python3 -m http.server 8080
```
Buka http://localhost:8080 (service worker hanya aktif di `localhost` atau HTTPS, bukan lewat file://).

## Memasang untuk dipakai sehari-hari
Unggah seluruh folder ke hosting statis HTTPS (Cloudflare Pages, Netlify, GitHub Pages, dll.).
Setelah dibuka di Chrome/Edge/Android akan muncul tombol "Pasang aplikasi"; di iPhone gunakan Bagikan → Tambah ke Layar Utama.

## Fitur
- **Data**: cari, filter bertingkat, urut beberapa kolom, ekspor Excel/CSV.
- **Rapikan**: hapus spasi berlebih, baris/kolom kosong, duplikat; ubah "Rp 1.250.000" jadi angka; atur huruf besar/kecil; isi sel merge yang kosong.
- **Kelompok**: group by banyak kolom dengan jumlah, hitung, rata-rata, min, maks, hitung unik.
- **Pivot**: baris (bisa lebih dari satu), kolom, nilai, fungsi, dan total.

## Untuk laporan hasil pivot/merge Excel
- Jika satu sheet berisi beberapa laporan berdampingan (dipisah kolom kosong), pilih **Blok kolom** di bagian atas.
- Tombol **⚡ Rapikan otomatis** (tab Rapikan): menghapus baris Total/Subtotal, mengisi sel gabungan yang kosong, mengubah teks tanggal dan angka, serta merapikan spasi. Hasilnya bisa diurungkan.
- Teks tanggal diubah jadi tanggal sungguhan (urutan hari/bulan dideteksi otomatis; jika tidak bisa dipastikan, dipakai hari/bulan/tahun dan Anda diberi peringatan).

## Catatan
- Angka berformat Indonesia (`1.250,50`) dikenali. `1.250` dibaca sebagai 1250. Kode berawalan 0 (mis. nomor telepon) dan angka ≥12 digit tidak diubah.
- CSV diekspor dengan pemisah `;` dan koma desimal agar langsung rapi di Excel versi Indonesia.
- Tabel menampilkan 500 baris pertama agar tetap cepat; ekspor selalu berisi semua baris.
- Pustaka SheetJS (`lib/xlsx.full.min.js`, versi 0.18.5) ikut dibundel agar bisa offline.
