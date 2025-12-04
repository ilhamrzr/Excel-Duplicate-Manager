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

### 🔹 1. Pemindaian Duplikasi Berdasarkan Kolom

- Pilih kolom target melalui ComboBox.
- Baris duplikat otomatis diberi highlight merah.
- Ringkasan hasil ditampilkan setelah pemindaian.

### 🔹 2. Dukungan **Header Row** Fleksibel (Baru pada v1.1.0)

- Header tidak lagi terbatas pada baris pertama.
- Anda dapat menentukan header berada di baris mana saja (misal: row 3, row 5, row 7, dll).
- ComboBox kolom otomatis mengikuti header yang dipilih.

### 🔹 3. Penghapusan Duplikasi + Backup Otomatis

- Sebelum menghapus duplikasi, Add-In membuat sheet backup dengan nama: BK_YYMMDD_HHMMSS
- Backup ini berisi seluruh isi sheet sebelum duplikasi dihapus.
- Backup hidden digunakan sebagai sumber data untuk fitur restore.

### 🔹 4. Fitur Restore Data

- Mengembalikan sheet ke kondisi sebelum penghapusan duplikasi.
- Aman, cepat, dan tidak merusak sheet lainnya.

### 🔹 5. Tema Dark/Light Mode

- Add-In menyediakan dua mode tampilan.
- Mode dapat ditukar menggunakan shortcut.

### 🔹 6. Tombol Reset (Baru pada v1.1.0)

- Menghapus semua highlight duplikasi.
- Mereset status scan dan summary.
- Tidak menutup jendela Duplicate Manager (berbeda dari tombol X di pojok kanan atas).

---

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
   dari folder `/dist` atau **GitHub Releases**.

2. Buka Excel → **File → Options → Add-ins**

3. Pada bagian bawah, pilih:
   Manage: Excel Add-ins → Go...

4. Klik **Browse** → pilih file:
   `DuplicateManager.xlam`

5. Centang add-in tersebut → **OK**

6. Selesai!  
   Shortcut dan fitur akan aktif di semua workbook.

---

## # 📚 Cara Penggunaan (Versi Final)

### **1. Buka Duplicate Manager**

**Shortcut:** CTRL + SHIFT + D

### **2. Tentukan Header Row**

Masukkan nomor baris tempat header kolom berada.  
Contoh:

- Header berada di baris 1 → isi **1**

- Header berada di baris 5 → isi **5**

ComboBox kolom akan menyesuaikan berdasarkan header row ini.

### **3. Pilih Header Kolom**

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
│ ├── ThisWorkbook.cls
│
├── dist/
│ ├── DuplicateManager.xlam
│ ├── DuplicateManager_v1.1.0.zip
│
├── CHANGELOG.md
├── README.md
├── LICENSE

---

🆕 Changelog

Semua catatan perubahan dapat dilihat di sini:

👉 **[CHANGELOG.md](./CHANGELOG.md)**

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
