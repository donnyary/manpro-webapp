-- ============================================
-- MIGRASI: TABEL SESI AKTIF PENGGUNA
-- Melacak siapa yang sedang login + IP + Device
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
