# Changelog

Semua perubahan penting pada proyek ini didokumentasikan di file ini.

Format mengikuti pedoman [Keep a Changelog](https://keepachangelog.com/en/1.0.0/)  
dan penomoran versi mengikuti prinsip [Semantic Versioning](https://semver.org/).

---

## [1.1.0] - 2025-12-04

### Added

- Penambahan fitur **Header Row**, memungkinkan pengguna memilih baris mana yang menjadi header kolom.
- Penambahan kontrol **TextBox: Header Row** untuk fleksibilitas struktur data.
- Penambahan fitur **Reset**, sebagai pengganti tombol Close untuk membersihkan highlight, font, dan summary tanpa menutup form.
- Penambahan full support Dark/Light mode pada kontrol baru (`lblHeaderRow`, `txtHeaderRow`).

### Changed

- Logika pencarian kolom (colStart) sekarang membaca header sesuai nilai `Header Row`.
- Validasi `Start Row` diperbarui agar selalu dimulai setelah baris header.
- UI diperbarui agar lebih konsisten dan responsif setelah penambahan fitur baru.

### Fixed

- Masalah di mana header tidak terbaca ketika berada di baris selain baris 1.
- Perilaku label yang tidak mengikuti styling Dark/Light mode setelah penambahan kontrol baru.

---

## [1.0.0] - 2025-12-04

### Added

- Fitur pemindaian duplikasi berdasarkan kolom yang dipilih.
- Sistem highlight otomatis untuk baris duplikat.
- Fitur penghapusan duplikasi dengan **backup sheet otomatis**.
- Fitur restore data menggunakan sheet backup tersembunyi.
- Antarmuka mendukung **Dark/Light Mode**.
- Shortcut global:
  - `Ctrl + Shift + D` — buka Duplicate Manager
  - `Ctrl + Shift + R` — restore data
  - `Ctrl + Shift + M` — toggle Dark/Light mode
