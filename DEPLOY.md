# 🚀 Panduan Deploy MANPRO ke Internet (GRATIS)

## Overview
- **Database Cloud**: TiDB Cloud (MySQL-compatible, 5GB gratis)
- **Hosting Cloud**: **Railway.app** (tanpa kartu kredit) atau Render.com
- **Total Biaya**: GRATIS

---

## ⭐ CARA CEPAT: Deploy ke Railway.app (TANPA KARTU KREDIT)

### STEP 1: Buat Database di TiDB Cloud

1. Buka **https://tidbcloud.com** → Sign Up (Google/GitHub)
2. Pilih **Serverless** → Region **AWS Singapore** → Klik **Create**
3. Tunggu 1-2 menit → Klik nama cluster → tab **Connect**
4. Copy connection string (format: `mysql://user:pass@host:4000/db_proyek`)
5. Catat **Host**, **User**, **Password** untuk digunakan nanti

### STEP 2: Upload ke GitHub

```bash
cd D:\Software\BelajarPython\daily_report_proyek_web
git init
git add .
git commit -m "First commit - MANPRO"
git remote add origin https://github.com/USERNAME/manpro-webapp.git
git branch -M main
git push -u origin main
```

> **Catatan**: Jangan push `db_proyek.db` (sudah di-.gitignore)

### STEP 3: Deploy ke Railway

1. Buka **https://railway.app** → Login pakai **GitHub**
2. Dashboard → klik **New Project** → **Deploy from GitHub repo**
3. Pilih repository **manpro-webapp**
4. Railway akan otomatis detect Python app → klik **Deploy**

### STEP 4: Tambah MySQL Database

1. Di dashboard Railway project → klik **New** → **Database** → **MySQL**
2. Tunggu MySQL selesai provision (1-2 menit)
3. Klik service MySQL → tab **Variables** → copy `MYSQL_URL`

### STEP 5: Set Environment Variables

1. Klik service **manpro-webapp** → tab **Variables**
2. Tambahkan environment variables berikut:

| Key | Value |
|-----|-------|
| `SECRET_KEY` | `manprosecretkey2026xyz` (bebas) |
| `MYSQL_URL` | *(paste dari MySQL service di Railway)* |

> **Tips**: Railway juga otomatis set `MYSQL_URL` ke app jika MySQL dan app di same project. Cek apakah sudah ada.

### STEP 6: Verifikasi

1. Klik tab **Settings** → scroll ke **Networking** → klik **Generate Domain**
2. URL akan seperti: `manpro-webapp-production-xxxx.up.railway.app`
3. Buka URL → Login `admin` / `admin123`
4. ✅ Aplikasi berjalan!

---

## CARA ALTERNATIF: Deploy ke Render.com

> ⚠️ Render memerlukan kartu kredit untuk verifikasi (tidak dikenakan biaya)

### STEP 1: Buat Database di TiDB Cloud
(Sama seperti di atas)

### STEP 2: Upload ke GitHub
(Sama seperti di atas)

### STEP 3: Deploy ke Render

1. Buka **https://render.com** → Login pakai GitHub
2. **New +** → **Web Service** → pilih repository
3. Isi form:
   - **Name**: `manpro-app`
   - **Runtime**: `Python 3`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn daily:app --bind 0.0.0.0:$PORT`
   - **Plan**: Free

4. Environment Variables:
   - `SECRET_KEY` → `manprosecretkey2026xyz`
   - `DATABASE_URL` → `mysql://user:pass@host:4000/db_proyek`

5. Klik **Create Web Service** → Tunggu 3-5 menit

---

## ⚠️ Tips Penting

### Auto-Migrate
Aplikasi `daily.py` sudah memiliki **auto-migrate** — saat pertama kali dijalankan, ia akan otomatis:
- Membuat semua tabel (`users`, `master_proyek`, `laporan_harian`, dll)
- Seed default permissions (`menu_permissions`)
- Buat admin default (`admin` / `admin123`)

**Tidak perlu jalankan SQL manual!**

### Railway Free Tier
- Mendapat **$5 credit/bulan** (gratis)
- App tidak sleep seperti Render — selalu aktif
- Sangat cocok untuk Flask app

### Keamanan
- **Jangan commit** file `.env` atau password ke GitHub
- Gunakan environment variables untuk semua secret
- Database cloud sudah pakai SSL otomatis

---

## 🔧 Troubleshooting

### Error: "Table doesn't exist"
- Auto-migrate akan membuat tabel otomatis
- Jika masih error, cek `MYSQL_URL` format benar

### Error: "Access denied"
- Cek username/password di environment variables
- Pastikan MySQL service di Railway sudah **Running**

### App Slow / Timeout
- **Railway**: Pertama kali akses mungkin agak lambat (cold start ~10-20 detik)
- **Render**: Free tier sleep setelah 15 menit idle, butuh ~30-60 detik wake up

---

## 📋 Checklist

### Railway (Recommended)
- [ ] TiDB Cloud account dibuat + cluster serverless
- [ ] GitHub repository dibuat & file di-push
- [ ] Railway account dibuat (pakai GitHub login)
- [ ] Project baru + deploy dari GitHub
- [ ] MySQL database ditambahkan di Railway
- [ ] Environment variables diisi (`MYSQL_URL`, `SECRET_KEY`)
- [ ] Deploy berhasil
- [ ] Aplikasi bisa diakses via internet

### Render (Alternatif)
- [ ] TiDB Cloud account dibuat + cluster serverless
- [ ] GitHub repository dibuat & file di-push
- [ ] Render account dibuat + kartu kredit (verifikasi)
- [ ] Web service dibuat
- [ ] Environment variables diisi (`DATABASE_URL`, `SECRET_KEY`)
- [ ] Deploy berhasil
- [ ] Aplikasi bisa diakses via internet
