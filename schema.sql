-- ============================================
-- SKEMA DATABASE: Sistem Manajemen Proyek Terpadu
-- Database: db_proyek
-- Engine: InnoDB (support foreign keys)
-- ============================================

CREATE DATABASE IF NOT EXISTS `db_proyek`
  DEFAULT CHARACTER SET utf8mb4
  COLLATE utf8mb4_unicode_ci;

USE `db_proyek`;

-- ============================================
-- 0. TABEL USERS (Autentikasi)
-- ============================================
CREATE TABLE IF NOT EXISTS `users` (
    `id`           INT AUTO_INCREMENT PRIMARY KEY,
    `nama_lengkap` VARCHAR(150) NOT NULL,
    `username`     VARCHAR(80)  NOT NULL UNIQUE,
    `password_hash` VARCHAR(255) NOT NULL COMMENT 'Hashed password (werkzeug)',
    `role`         ENUM('admin','manager','user') NOT NULL DEFAULT 'user',
    `status`       ENUM('pending','approved','rejected') NOT NULL DEFAULT 'approved' COMMENT 'Status persetujuan akun',
    `created_at`   DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB;

-- ============================================
-- 1. MASTER PROYEK
-- ============================================
CREATE TABLE IF NOT EXISTS `master_proyek` (
    `id`                  INT AUTO_INCREMENT PRIMARY KEY,
    `nama_kegiatan`       VARCHAR(255) NOT NULL,
    `nama_proyek`         VARCHAR(255) NOT NULL,
    `penyedia_jasa`       VARCHAR(255) DEFAULT NULL,
    `konsultan`           VARCHAR(255) DEFAULT NULL,
    `pemilik_proyek`      VARCHAR(255) DEFAULT NULL,
    `lokasi`              VARCHAR(500) DEFAULT NULL,
    `no_kontrak`          VARCHAR(100) DEFAULT NULL,
    `tgl_mulai`           DATE DEFAULT NULL,
    `pagu_kontrak_total`  DECIMAL(20,2) DEFAULT 0.00,
    `alamat_kontraktor`   TEXT DEFAULT NULL,
    `telp_kontraktor`     VARCHAR(50) DEFAULT NULL,
    `logo_perusahaan`     VARCHAR(255) DEFAULT NULL COMMENT 'File logo kontraktor',
    `logo_pemilik`        VARCHAR(255) DEFAULT NULL COMMENT 'File logo owner/proyek',
    `created_at`          DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB;

-- ============================================
-- 2. LAPORAN HARIAN (Header)
-- ============================================
CREATE TABLE IF NOT EXISTS `laporan_harian` (
    `id`              INT AUTO_INCREMENT PRIMARY KEY,
    `proyek_id`       INT NOT NULL,
    `tanggal_laporan` DATE NOT NULL,
    `cuaca`           VARCHAR(100) DEFAULT NULL,
    `dokumen_upload`  VARCHAR(255) DEFAULT NULL,
    `created_at`      DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (`proyek_id`) REFERENCES `master_proyek`(`id`) ON DELETE CASCADE
) ENGINE=InnoDB;

-- ============================================
-- 3. TENAGA KERJA
-- ============================================
CREATE TABLE IF NOT EXISTS `tenaga_kerja` (
    `id`             INT AUTO_INCREMENT PRIMARY KEY,
    `laporan_id`     INT NOT NULL,
    `jenis_pekerja`  VARCHAR(150) DEFAULT NULL,
    `jumlah`         INT DEFAULT 0,
    `hadir`          INT DEFAULT 0,
    `tidak_hadir`    INT DEFAULT 0,
    `keterangan`     VARCHAR(255) DEFAULT NULL,
    FOREIGN KEY (`laporan_id`) REFERENCES `laporan_harian`(`id`) ON DELETE CASCADE
) ENGINE=InnoDB;

-- ============================================
-- 4. PERALATAN / ALAT BERAT
-- ============================================
CREATE TABLE IF NOT EXISTS `peralatan` (
    `id`            INT AUTO_INCREMENT PRIMARY KEY,
    `laporan_id`    INT NOT NULL,
    `nama_alat`     VARCHAR(200) DEFAULT NULL,
    `jumlah_alat`   INT DEFAULT 0,
    `kondisi`       VARCHAR(100) DEFAULT NULL,
    `keterangan`    VARCHAR(255) DEFAULT NULL,
    FOREIGN KEY (`laporan_id`) REFERENCES `laporan_harian`(`id`) ON DELETE CASCADE
) ENGINE=InnoDB;

-- ============================================
-- 5. MATERIAL / BAHAN DATANG
-- ============================================
CREATE TABLE IF NOT EXISTS `material` (
    `id`               INT AUTO_INCREMENT PRIMARY KEY,
    `laporan_id`       INT NOT NULL,
    `nama_material`    VARCHAR(200) DEFAULT NULL,
    `volume_datang`    DECIMAL(12,2) DEFAULT 0,
    `satuan`           VARCHAR(50) DEFAULT NULL,
    `keterangan`       VARCHAR(255) DEFAULT NULL,
    FOREIGN KEY (`laporan_id`) REFERENCES `laporan_harian`(`id`) ON DELETE CASCADE
) ENGINE=InnoDB;

-- ============================================
-- 6. PEKERJAAN / PROGRES HARIAN
-- ============================================
CREATE TABLE IF NOT EXISTS `pekerjaan` (
    `id`                 INT AUTO_INCREMENT PRIMARY KEY,
    `laporan_id`         INT NOT NULL,
    `jenis_pekerjaan`    VARCHAR(255) DEFAULT NULL,
    `lokasi`             VARCHAR(255) DEFAULT NULL,
    `vol_harian`         VARCHAR(100) DEFAULT NULL,
    `proses_kumulatif`   VARCHAR(50) DEFAULT NULL,
    `target_harian`      VARCHAR(50) DEFAULT NULL,
    `keterangan`         VARCHAR(255) DEFAULT NULL,
    FOREIGN KEY (`laporan_id`) REFERENCES `laporan_harian`(`id`) ON DELETE CASCADE
) ENGINE=InnoDB;

-- ============================================
-- 7. KONDISI LAPANGAN
-- ============================================
CREATE TABLE IF NOT EXISTS `kondisi_lapangan` (
    `id`                 INT AUTO_INCREMENT PRIMARY KEY,
    `laporan_id`         INT NOT NULL UNIQUE,
    `akses`              VARCHAR(255) DEFAULT NULL,
    `k3`                 VARCHAR(255) DEFAULT NULL,
    `kondisi_fisik`      VARCHAR(255) DEFAULT NULL,
    `hambatan`           VARCHAR(255) DEFAULT NULL,
    `tak_terencana`      TEXT DEFAULT NULL,
    FOREIGN KEY (`laporan_id`) REFERENCES `laporan_harian`(`id`) ON DELETE CASCADE
) ENGINE=InnoDB;

-- ============================================
-- 8. PENGESAHAN
-- ============================================
CREATE TABLE IF NOT EXISTS `pengesahan` (
    `id`               INT AUTO_INCREMENT PRIMARY KEY,
    `laporan_id`       INT NOT NULL UNIQUE,
    `nama_pembuat`     VARCHAR(200) DEFAULT NULL,
    `nama_penyetuju`   VARCHAR(200) DEFAULT NULL,
    FOREIGN KEY (`laporan_id`) REFERENCES `laporan_harian`(`id`) ON DELETE CASCADE
) ENGINE=InnoDB;

-- ============================================
-- 9. MASTER WBS (Work Breakdown Structure)
-- ============================================
CREATE TABLE IF NOT EXISTS `master_wbs` (
    `id`              INT AUTO_INCREMENT PRIMARY KEY,
    `proyek_id`       INT NOT NULL,
    `kode_wbs`        VARCHAR(50) NOT NULL,
    `nama_pekerjaan`  VARCHAR(255) NOT NULL,
    `bobot_persen`    DECIMAL(6,2) DEFAULT 0.00,
    FOREIGN KEY (`proyek_id`) REFERENCES `master_proyek`(`id`) ON DELETE CASCADE
) ENGINE=InnoDB;

-- ============================================
-- 10. KATEGORI BUDGET (per proyek)
-- ============================================
CREATE TABLE IF NOT EXISTS `kategori_budget` (
    `id`             INT AUTO_INCREMENT PRIMARY KEY,
    `proyek_id`      INT DEFAULT NULL COMMENT 'ID proyek terkait (NULL untuk kategori global)',
    `nama_kategori`  VARCHAR(100) NOT NULL,
    `anggaran_pagu`  DECIMAL(20,2) DEFAULT 0.00,
    UNIQUE KEY `uniq_kategori_per_proyek` (`proyek_id`, `nama_kategori`),
    FOREIGN KEY (`proyek_id`) REFERENCES `master_proyek`(`id`) ON DELETE CASCADE
) ENGINE=InnoDB;

-- ============================================
-- 11. MASTER BUDGET / TRANSAKSI KAS
-- ============================================
CREATE TABLE IF NOT EXISTS `master_budget` (
    `id`                 INT AUTO_INCREMENT PRIMARY KEY,
    `proyek_id`          INT NOT NULL,
    `kategori`           VARCHAR(100) DEFAULT NULL,
    `nama_item`          VARCHAR(255) DEFAULT NULL,
    `anggaran_pagu`      DECIMAL(20,2) DEFAULT 0.00,
    `realisasi_biaya`    DECIMAL(20,2) DEFAULT 0.00,
    `no_transaksi`       VARCHAR(50) DEFAULT NULL,
    `tanggal_transaksi`  DATE DEFAULT NULL,
    `jenis_kas`          ENUM('masuk','keluar') DEFAULT 'keluar',
    `pembuat`            VARCHAR(100) DEFAULT NULL,
    `kuantitas`          DECIMAL(12,2) DEFAULT 1.00,
    `harga_satuan`       DECIMAL(20,2) DEFAULT 0.00,
    `created_at`         DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (`proyek_id`) REFERENCES `master_proyek`(`id`) ON DELETE CASCADE
) ENGINE=InnoDB;

-- ============================================
-- 12. MENU PERMISSIONS (HAK AKSES PER ROLE)
-- ============================================
CREATE TABLE IF NOT EXISTS `menu_permissions` (
    `id`         INT AUTO_INCREMENT PRIMARY KEY,
    `role`       ENUM('admin','manager','user') NOT NULL,
    `menu_key`   VARCHAR(50) NOT NULL COMMENT 'Identifier unik menu',
    `can_access` TINYINT(1) NOT NULL DEFAULT 1 COMMENT '1=Boleh akses, 0=Tidak boleh',
    UNIQUE KEY `uniq_role_menu` (`role`, `menu_key`)
) ENGINE=InnoDB;

-- ============================================
-- 13. ACTIVE SESSIONS (SESI AKTIF PENGGUNA)
-- ============================================
CREATE TABLE IF NOT EXISTS `active_sessions` (
    `id`            INT AUTO_INCREMENT PRIMARY KEY,
    `user_id`       INT NOT NULL,
    `session_token` VARCHAR(255) NOT NULL COMMENT 'Session ID dari Flask',
    `ip_address`    VARCHAR(50) DEFAULT NULL,
    `user_agent`    TEXT DEFAULT NULL COMMENT 'Raw User-Agent string',
    `device_type`   VARCHAR(50) DEFAULT NULL COMMENT 'Desktop/Mobile/Tablet',
    `browser`       VARCHAR(100) DEFAULT NULL COMMENT 'Nama browser',
    `os`            VARCHAR(100) DEFAULT NULL COMMENT 'Sistem operasi',
    `login_at`      DATETIME DEFAULT CURRENT_TIMESTAMP,
    `last_active`   DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    `is_active`     TINYINT(1) NOT NULL DEFAULT 1,
    KEY `idx_user_id` (`user_id`),
    KEY `idx_session_token` (`session_token`),
    KEY `idx_is_active` (`is_active`)
) ENGINE=InnoDB;

-- ============================================
-- USER DEFAULT (Admin)
-- Password: admin123 (akan di-hash oleh seed_data.py)
-- ============================================
-- Catatan: User default dibuat oleh seed_data.py karena password perlu di-hash
