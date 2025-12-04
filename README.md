# Excel Duplicate Manager

![License](https://img.shields.io/badge/license-MIT-green.svg)
![Excel](https://img.shields.io/badge/Excel-Add--In-217346?logo=microsoft-excel&logoColor=white)
![Status](https://img.shields.io/badge/status-stable-brightgreen)

**Excel Duplicate Manager** adalah Add-In Excel yang memudahkan Anda untuk:

- Mendeteksi duplikat
- Menghapus duplikat
- Membuat backup otomatis
- Melakukan restore data
- Menggunakan tampilan Dark Mode / Light Mode
- Menggunakan keyboard shortcut global

Add-In ini dibuat menggunakan **VBA** dan dapat dipakai pada semua versi Excel for Windows
(Office 2010 – Office 365).

---

## ✨ Fitur Utama

### 🔍 Scan Duplicates

- Memindai kolom tertentu berdasarkan header
- Highlight otomatis pada baris yang terduplikasi
- Real-time summary informasi duplikat

### 🗑 Delete + Backup

- Menghapus duplikat secara otomatis
- Membuat sheet backup (`BK_yyyymmdd_hhmmss`)
- Backup tersembunyi untuk opsi *restore last state*

### ♻ Restore Last State

- Mengembalikan data ke kondisi sebelum penghapusan
- Menggunakan hidden sheet `__DM_BACKUP`

### 🌓 Dark / Light Mode

- Tampilan UserForm bisa diubah dengan shortcut
- Warna tombol, background, dan teks menyesuaikan tema

### ⌨ Global Keyboard Shortcuts

Shortcut dapat digunakan di semua workbook:

| Shortcut             | Fungsi                 |
| -------------------- | ---------------------- |
| **Ctrl + Shift + D** | Buka Duplicate Manager |
| **Ctrl + Shift + R** | Restore Latest Backup  |
| **Ctrl + Shift + M** | Toggle Dark/Light Mode |

### 🧩 Add-In (XLAM)

- Dapat digunakan di semua file Excel tanpa perlu import ulang
- Aman untuk dibagikan atau digunakan secara offline
- Tidak membutuhkan library eksternal

---

## 📥 Instalasi Add-In

1. Download file:  
   **`DuplicateManager.xlam`**  
   dari folder `/build` atau **GitHub Releases**.

2. Buka Excel → **File → Options → Add-ins**

3. Pada bagian bawah, pilih:
   Manage: Excel Add-ins → Go...

4. Klik **Browse** → pilih file:
   `DuplicateManager.xlam`

5. Centang add-in tersebut → **OK**

6. Selesai!  
   Shortcut dan fitur akan aktif di semua workbook.

---

## 📚 Cara Penggunaan

### 1. Buka Duplicate Manager

Gunakan shortcut: `CTRL + SHIFT + D`

Atau via menu Add-ins → *Duplicate Manager*.

### 2. Pilih Header Kolom

Pilih header yang ingin dipindai duplikatnya.

### 3. Pilih Starting Row

Masukkan baris awal data (default: 2).

### 4. Klik **Scan Duplicates**

Baris duplikat akan di-highlight warna merah.

### 5. Klik **Delete + Backup**

- Semua duplikat dihapus
- Backup sheet dibuat otomatis

### 6. Klik **Restore Last State**

Untuk mengembalikan kondisi sebelum penghapusan.

---

## 📺 Demo Video

<video width="600" controls>
  <source src="https://raw.githubusercontent.com/ilhamrzr/Excel-Duplicate-Manager/main/demo/Duplicate%20Manager.mp4" type="video/mp4">
  Your browser does not support the video tag.
</video>

---

## 📂 Struktur Project

Excel-Duplicate-Manager/  
│  
├── src/  
│ ├── modDuplicateManager.bas  
│ ├── frmDuplicateManager.frm  
│ ├── frmDuplicateManager.frx  
│ └── ThisWorkbook.cls  
│  
├── build/  
│ └── DuplicateManager.xlam  
│  
├── README.md  
└── LICENSE

---

## 🧑‍💻 Kontribusi

Kontribusi sangat diterima!  
Anda dapat membantu dengan:

- Melaporkan bug  
- Mengajukan fitur baru  
- Mengoptimalkan performa  
- Menambahkan dokumentasi  

Submit melalui **Issues** atau **Pull Request**.

---

## 📄 Lisensi

Proyek ini dilisensikan di bawah **MIT License**.  
Silakan gunakan, modifikasi, dan distribusikan bebas.

---

## ⭐ Support Projek Ini

Jika Add-In ini membantu Anda, jangan lupa:

- ⭐ **Star repo ini di GitHub**
- 🔄 Share ke pengguna Excel lainnya

Terima kasih telah menggunakan Excel Duplicate Manager! 🙌
