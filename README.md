# Wonderland Samarinda — PWA

Frontend statis (GitHub Pages) + backend Google Apps Script/Spreadsheet.

## Isi folder
- `index.html` aplikasi (sudah responsif HP, sidebar jadi drawer)
- `config.js` URL Web App Apps Script (WAJIB diisi)
- `gas-bridge.js` pengganti `google.script.run` via fetch
- `manifest.webmanifest`, `sw.js`, `icons/` komponen PWA
- `Code.gs` backend (sudah ditambah `doPost` + whitelist) → ditempel ke Apps Script, BUKAN di-upload ke GitHub

## Langkah
1. **Apps Script**: ganti isi `Code.gs` dengan file ini → Deploy → Manage deployments → Edit (ikon pensil) → Version: *New version* → Execute as: **Me**, Who has access: **Anyone** → Deploy. Salin URL `/exec`.
2. **config.js**: tempel URL `/exec` ke `window.GAS_URL`.
3. **GitHub**: upload semua file KECUALI `Code.gs` dan README (opsional) ke repo → Settings → Pages → Deploy from branch `main` / root.
4. Buka `https://USERNAME.github.io/REPO/` di HP (Chrome: menu ⋮ → *Install app / Add to Home screen*; iOS Safari: Share → *Add to Home Screen*).

## Catatan
- Setiap mengubah `Code.gs`, buat **New version** deployment (URL tetap sama).
- Akses data dilindungi session token + `checkPermission_` di tiap fungsi; hanya fungsi di `API_WHITELIST_` yang bisa dipanggil. Jika menambah fungsi baru yang dipanggil klien, tambahkan namanya ke whitelist.
- Setelah update `index.html`, tutup & buka ulang aplikasi (service worker network-first untuk html/js). Naikkan `CACHE` di `sw.js` bila perlu.
- Mode offline hanya untuk tampilan aplikasi; transaksi tetap perlu internet.
- Upload logo (`uploadLogo`) mengirim data base64 via POST; ukuran besar bisa lambat.
