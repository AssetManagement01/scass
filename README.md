# SCASS

Backup kode **Google Apps Script** untuk:

- `pos/` — kasir Angkringan Gondrong (`Index.html` + `Code.gs`)
- `scass/` — ERP NewStar Kancing (`Scass.html` + `Code-Scass.gs`)

Repo ini **bukan** tempat menjalankan aplikasi. `google.script.run` hanya hidup di web app Apps Script.

## Cara pakai

1. Buka [script.google.com](https://script.google.com) → project yang sudah terhubung ke spreadsheet.
2. Tempel file:
   - Kasir: `pos/Code.gs` → `Kode.gs`, `pos/Index.html` → `Index.html`
   - ERP: `scass/Code-Scass.gs` → `Kode.gs` (project terpisah)
3. Deploy → **New version** → buka URL `/exec`.

## Login

- POS kasir: `AG` / `AG`
- SCASS ERP: `USER` / `USER`

Jangan host di GitHub Pages jika ingin Bayar / Simpan / Inventory tetap jalan.
