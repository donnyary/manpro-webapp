-- ============================================
-- MIGRASI: TABEL HAK AKSES (MENU PERMISSIONS)
-- Menambahkan kontrol akses per role per menu
-- ============================================

CREATE TABLE IF NOT EXISTS `menu_permissions` (
    `id`         INT AUTO_INCREMENT PRIMARY KEY,
    `role`       ENUM('admin','manager','user') NOT NULL,
    `menu_key`   VARCHAR(50) NOT NULL COMMENT 'Identifier unik menu',
    `can_access` TINYINT(1) NOT NULL DEFAULT 1 COMMENT '1=Boleh akses, 0=Tidak boleh',
    UNIQUE KEY `uniq_role_menu` (`role`, `menu_key`)
) ENGINE=InnoDB;

-- ============================================
-- SEED: DEFAULT PERMISSIONS PER ROLE
-- admin   = Full akses semua menu
-- manager = Akses semua kecuali Kelola User
-- user    = Akses terbatas (hanya lihat laporan & proyek)
-- ============================================

-- ADMIN: Full akses semua menu
INSERT IGNORE INTO `menu_permissions` (`role`, `menu_key`, `can_access`) VALUES
('admin', 'dashboard',        1),
('admin', 'pengaturan',       1),
('admin', 'users',            1),
('admin', 'active_sessions',  1),
('admin', 'daily_report',     1),
('admin', 'weekly_report',    1),
('admin', 'monthly_report',   1),
('admin', 'gantt',            1),
('admin', 'wbs',              1),
('admin', 'budget',           1);

-- MANAGER: Akses semua kecuali Kelola User & Sesi Aktif
INSERT IGNORE INTO `menu_permissions` (`role`, `menu_key`, `can_access`) VALUES
('manager', 'dashboard',        1),
('manager', 'pengaturan',       1),
('manager', 'users',            0),
('manager', 'active_sessions',  0),
('manager', 'daily_report',     1),
('manager', 'weekly_report',    1),
('manager', 'monthly_report',   1),
('manager', 'gantt',            1),
('manager', 'wbs',              1),
('manager', 'budget',           1);

-- USER: Akses terbatas - hanya laporan dan monitoring
INSERT IGNORE INTO `menu_permissions` (`role`, `menu_key`, `can_access`) VALUES
('user', 'dashboard',        1),
('user', 'pengaturan',       0),
('user', 'users',            0),
('user', 'active_sessions',  0),
('user', 'daily_report',     1),
('user', 'weekly_report',    1),
('user', 'monthly_report',   1),
('user', 'gantt',            1),
('user', 'wbs',              0),
('user', 'budget',           0);
