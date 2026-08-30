"""
Seed Data - Sistem Manajemen Proyek Terpadu
Membuat schema, admin user, dan data contoh 3 proyek.
"""

import mysql.connector
from werkzeug.security import generate_password_hash


def create_schema(cursor):
    """Buat semua tabel jika belum ada."""
    cursor.execute("""CREATE TABLE IF NOT EXISTS users (
        id INT AUTO_INCREMENT PRIMARY KEY,
        nama_lengkap VARCHAR(150) NOT NULL,
        username VARCHAR(80) NOT NULL UNIQUE,
        password_hash VARCHAR(255) NOT NULL,
        role ENUM('admin','manager','user') NOT NULL DEFAULT 'user',
        status ENUM('pending','approved','rejected') NOT NULL DEFAULT 'approved',
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    ) ENGINE=InnoDB""")

    cursor.execute("""CREATE TABLE IF NOT EXISTS master_proyek (
        id INT AUTO_INCREMENT PRIMARY KEY,
        nama_kegiatan VARCHAR(255) NOT NULL,
        nama_proyek VARCHAR(255) NOT NULL,
        penyedia_jasa VARCHAR(255) DEFAULT NULL,
        konsultan VARCHAR(255) DEFAULT NULL,
        pemilik_proyek VARCHAR(255) DEFAULT NULL,
        lokasi VARCHAR(500) DEFAULT NULL,
        no_kontrak VARCHAR(100) DEFAULT NULL,
        tgl_mulai DATE DEFAULT NULL,
        pagu_kontrak_total DECIMAL(20,2) DEFAULT 0.00,
        alamat_kontraktor TEXT DEFAULT NULL,
        telp_kontraktor VARCHAR(50) DEFAULT NULL,
        logo_perusahaan VARCHAR(255) DEFAULT NULL,
        logo_pemilik VARCHAR(255) DEFAULT NULL,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    ) ENGINE=InnoDB""")

    cursor.execute("""CREATE TABLE IF NOT EXISTS laporan_harian (
        id INT AUTO_INCREMENT PRIMARY KEY,
        proyek_id INT NOT NULL,
        tanggal_laporan DATE NOT NULL,
        cuaca VARCHAR(100) DEFAULT NULL,
        dokumen_upload VARCHAR(255) DEFAULT NULL,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (proyek_id) REFERENCES master_proyek(id) ON DELETE CASCADE
    ) ENGINE=InnoDB""")

    cursor.execute("""CREATE TABLE IF NOT EXISTS tenaga_kerja (
        id INT AUTO_INCREMENT PRIMARY KEY,
        laporan_id INT NOT NULL,
        jenis_pekerja VARCHAR(150) DEFAULT NULL,
        jumlah INT DEFAULT 0,
        hadir INT DEFAULT 0,
        tidak_hadir INT DEFAULT 0,
        keterangan VARCHAR(255) DEFAULT NULL,
        FOREIGN KEY (laporan_id) REFERENCES laporan_harian(id) ON DELETE CASCADE
    ) ENGINE=InnoDB""")

    cursor.execute("""CREATE TABLE IF NOT EXISTS peralatan (
        id INT AUTO_INCREMENT PRIMARY KEY,
        laporan_id INT NOT NULL,
        nama_alat VARCHAR(200) DEFAULT NULL,
        jumlah_alat INT DEFAULT 0,
        kondisi VARCHAR(100) DEFAULT NULL,
        keterangan VARCHAR(255) DEFAULT NULL,
        FOREIGN KEY (laporan_id) REFERENCES laporan_harian(id) ON DELETE CASCADE
    ) ENGINE=InnoDB""")

    cursor.execute("""CREATE TABLE IF NOT EXISTS material (
        id INT AUTO_INCREMENT PRIMARY KEY,
        laporan_id INT NOT NULL,
        nama_material VARCHAR(200) DEFAULT NULL,
        volume_datang DECIMAL(12,2) DEFAULT 0,
        satuan VARCHAR(50) DEFAULT NULL,
        keterangan VARCHAR(255) DEFAULT NULL,
        FOREIGN KEY (laporan_id) REFERENCES laporan_harian(id) ON DELETE CASCADE
    ) ENGINE=InnoDB""")

    cursor.execute("""CREATE TABLE IF NOT EXISTS pekerjaan (
        id INT AUTO_INCREMENT PRIMARY KEY,
        laporan_id INT NOT NULL,
        jenis_pekerjaan VARCHAR(255) DEFAULT NULL,
        lokasi VARCHAR(255) DEFAULT NULL,
        vol_harian VARCHAR(100) DEFAULT NULL,
        proses_kumulatif VARCHAR(50) DEFAULT NULL,
        target_harian VARCHAR(50) DEFAULT NULL,
        keterangan VARCHAR(255) DEFAULT NULL,
        FOREIGN KEY (laporan_id) REFERENCES laporan_harian(id) ON DELETE CASCADE
    ) ENGINE=InnoDB""")

    cursor.execute("""CREATE TABLE IF NOT EXISTS kondisi_lapangan (
        id INT AUTO_INCREMENT PRIMARY KEY,
        laporan_id INT NOT NULL UNIQUE,
        akses VARCHAR(255) DEFAULT NULL,
        k3 VARCHAR(255) DEFAULT NULL,
        kondisi_fisik VARCHAR(255) DEFAULT NULL,
        hambatan VARCHAR(255) DEFAULT NULL,
        tak_terencana TEXT DEFAULT NULL,
        FOREIGN KEY (laporan_id) REFERENCES laporan_harian(id) ON DELETE CASCADE
    ) ENGINE=InnoDB""")

    cursor.execute("""CREATE TABLE IF NOT EXISTS pengesahan (
        id INT AUTO_INCREMENT PRIMARY KEY,
        laporan_id INT NOT NULL UNIQUE,
        nama_pembuat VARCHAR(200) DEFAULT NULL,
        nama_penyetuju VARCHAR(200) DEFAULT NULL,
        FOREIGN KEY (laporan_id) REFERENCES laporan_harian(id) ON DELETE CASCADE
    ) ENGINE=InnoDB""")

    cursor.execute("""CREATE TABLE IF NOT EXISTS master_wbs (
        id INT AUTO_INCREMENT PRIMARY KEY,
        proyek_id INT NOT NULL,
        kode_wbs VARCHAR(50) NOT NULL,
        nama_pekerjaan VARCHAR(255) NOT NULL,
        bobot_persen DECIMAL(6,2) DEFAULT 0.00,
        FOREIGN KEY (proyek_id) REFERENCES master_proyek(id) ON DELETE CASCADE
    ) ENGINE=InnoDB""")

    cursor.execute("""CREATE TABLE IF NOT EXISTS kategori_budget (
        id INT AUTO_INCREMENT PRIMARY KEY,
        proyek_id INT DEFAULT NULL,
        nama_kategori VARCHAR(100) NOT NULL,
        anggaran_pagu DECIMAL(20,2) DEFAULT 0.00,
        UNIQUE KEY uniq_kategori_per_proyek (proyek_id, nama_kategori),
        FOREIGN KEY (proyek_id) REFERENCES master_proyek(id) ON DELETE CASCADE
    ) ENGINE=InnoDB""")

    cursor.execute("""CREATE TABLE IF NOT EXISTS master_budget (
        id INT AUTO_INCREMENT PRIMARY KEY,
        proyek_id INT NOT NULL,
        kategori VARCHAR(100) DEFAULT NULL,
        nama_item VARCHAR(255) DEFAULT NULL,
        anggaran_pagu DECIMAL(20,2) DEFAULT 0.00,
        realisasi_biaya DECIMAL(20,2) DEFAULT 0.00,
        no_transaksi VARCHAR(50) DEFAULT NULL,
        tanggal_transaksi DATE DEFAULT NULL,
        jenis_kas ENUM('masuk','keluar') DEFAULT 'keluar',
        pembuat VARCHAR(100) DEFAULT NULL,
        kuantitas DECIMAL(12,2) DEFAULT 1.00,
        harga_satuan DECIMAL(20,2) DEFAULT 0.00,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (proyek_id) REFERENCES master_proyek(id) ON DELETE CASCADE
    ) ENGINE=InnoDB""")

    # Kolom tambahan yang mungkin belum ada pada tabel lama
    try:
        cursor.execute("ALTER TABLE users ADD COLUMN `status` ENUM('pending','approved','rejected') NOT NULL DEFAULT 'approved' AFTER `role`")
        print("Kolom 'status' ditambahkan ke tabel users.")
    except Exception:
        pass

    try:
        cursor.execute("ALTER TABLE master_budget ADD COLUMN `kuantitas` DECIMAL(12,2) DEFAULT 1.00 AFTER `pembuat`")
        print("Kolom 'kuantitas' ditambahkan ke tabel master_budget.")
    except Exception:
        pass

    try:
        cursor.execute("ALTER TABLE master_budget ADD COLUMN `harga_satuan` DECIMAL(20,2) DEFAULT 0.00 AFTER `kuantitas`")
        print("Kolom 'harga_satuan' ditambahkan ke tabel master_budget.")
    except Exception:
        pass

    # Tambah kolom proyek_id & anggaran_pagu ke kategori_budget jika belum ada
    try:
        cursor.execute("ALTER TABLE kategori_budget ADD COLUMN `proyek_id` INT DEFAULT NULL AFTER `id`")
        print("Kolom 'proyek_id' ditambahkan ke tabel kategori_budget.")
    except Exception:
        pass

    try:
        cursor.execute("ALTER TABLE kategori_budget ADD COLUMN `anggaran_pagu` DECIMAL(20,2) DEFAULT 0.00 AFTER `nama_kategori`")
        print("Kolom 'anggaran_pagu' ditambahkan ke tabel kategori_budget.")
    except Exception:
        pass

    print("Schema database berhasil dibuat/dipastikan.")


def seed_admin_user(cursor):
    """Buat user admin default jika belum ada."""
    cursor.execute("SELECT COUNT(*) FROM users WHERE username = 'admin'")
    cnt = cursor.fetchone()[0]
    if cnt == 0:
        pw = generate_password_hash('admin123')
        cursor.execute(
            "INSERT INTO users (nama_lengkap, username, password_hash, role, status) VALUES (%s, %s, %s, %s, %s)",
            ('Administrator', 'admin', pw, 'admin', 'approved')
        )
        print("User admin dibuat -> Username: admin | Password: admin123")
    else:
        print("User admin sudah ada, skip.")


def seed_database():
    conn = mysql.connector.connect(
        host='127.0.0.1', port=3306, user='root', password='', database='db_proyek'
    )
    cur = conn.cursor()

    create_schema(cur)
    conn.commit()
    seed_admin_user(cur)
    conn.commit()

    # Bersihkan data lama
    print("Membersihkan data lama...")
    for t in ['master_budget','master_wbs','pengesahan','kondisi_lapangan',
              'material','pekerjaan','tenaga_kerja','peralatan',
              'laporan_harian','master_proyek']:
        cur.execute(f'DELETE FROM {t}')
    conn.commit()

    print("Memasukkan data contoh 3 proyek...")

    # === PROYEK 1 ===
    cur.execute(
        "INSERT INTO master_proyek (nama_kegiatan,nama_proyek,penyedia_jasa,konsultan,"
        "pemilik_proyek,lokasi,no_kontrak,tgl_mulai,pagu_kontrak_total) "
        "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)",
        ("Pembangunan Infrastruktur Pendidikan Tinggi",
         "Gedung Kuliah 5 Lantai Fakultas Teknik",
         "PT. Adhi Karya (Persero) Tbk", "PT. Virama Karya (Persero)",
         "Yayasan Pendidikan Cendekia Utama",
         "Jl. Boulevard Universitas No. 45, Depok",
         "042/KONTRAK-FTEK/VIII/2026", "2026-03-01", 45000000000))
    p1 = cur.lastrowid
    cur.executemany("INSERT INTO master_wbs (proyek_id,kode_wbs,nama_pekerjaan,bobot_persen) VALUES (%s,%s,%s,%s)",
                    [(p1,"1.1","Struktur Bawah (Pondasi & Basement)",20),
                     (p1,"1.2","Struktur Atas (Lantai 1 s.d Roof)",35),
                     (p1,"1.3","Arsitektur & Finishing",30),
                     (p1,"1.4","MEP (Mechanical, Electrical, Plumbing)",15)])
    cur.executemany("INSERT INTO master_budget (proyek_id,kategori,nama_item,anggaran_pagu,realisasi_biaya) VALUES (%s,%s,%s,%s,%s)",
                    [(p1,"JASA SUBCONT","Jasa Struktur & Beton Ready Mix",18000000000,16500000000),
                     (p1,"MATERIAL","Besi Beton U-39 & Semen",12000000000,11800000000),
                     (p1,"UPAH KERJA","Tukang & Pekerja Lapangan",8000000000,7900000000),
                     (p1,"PERALATAN","Sewa Tower Crane & Concrete Pump",7000000000,6800000000)])
    cur.execute("INSERT INTO laporan_harian (proyek_id,tanggal_laporan,cuaca) VALUES (%s,%s,%s)",
                (p1,"2026-03-15","Cerah Berawan (30 C)"))
    l1 = cur.lastrowid
    cur.execute("INSERT INTO tenaga_kerja (laporan_id,jenis_pekerja,jumlah,hadir,tidak_hadir,keterangan) VALUES (%s,%s,%s,%s,%s,%s)",
                (l1,"Tukang Batu & Besi",55,53,2,"Lengkap dan produktif"))
    cur.execute("INSERT INTO peralatan (laporan_id,nama_alat,jumlah_alat,kondisi,keterangan) VALUES (%s,%s,%s,%s,%s)",
                (l1,"Tower Crane TC-6015",1,"Baik","Beroperasi normal"))
    cur.execute("INSERT INTO material (laporan_id,nama_material,volume_datang,satuan,keterangan) VALUES (%s,%s,%s,%s,%s)",
                (l1,"Beton Ready Mix K-350",45.0,"m3","Sesuai jadwal"))
    cur.execute("INSERT INTO pekerjaan (laporan_id,jenis_pekerjaan,lokasi,vol_harian,proses_kumulatif,target_harian,keterangan) VALUES (%s,%s,%s,%s,%s,%s,%s)",
                (l1,"Pengecoran Kolom Lantai 3","Zona A","45 m3","18%","40 m3","Selesai tepat waktu"))
    cur.execute("INSERT INTO kondisi_lapangan (laporan_id,akses,k3,kondisi_fisik,hambatan,tak_terencana) VALUES (%s,%s,%s,%s,%s,%s)",
                (l1,"Lancar","Lengkap (APD Standar)","Baik","Nihil","-"))
    cur.execute("INSERT INTO pengesahan (laporan_id,nama_pembuat,nama_penyetuju) VALUES (%s,%s,%s)",
                (l1,"Ir. Budi Santoso","Ir. Hartono, M.T."))

    # === PROYEK 2 ===
    cur.execute(
        "INSERT INTO master_proyek (nama_kegiatan,nama_proyek,penyedia_jasa,konsultan,"
        "pemilik_proyek,lokasi,no_kontrak,tgl_mulai,pagu_kontrak_total) "
        "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)",
        ("Peningkatan Infrastruktur Transportasi Wilayah Selatan",
         "Revitalisasi Jembatan Citarik & Jalan Penghubung",
         "PT. Pembangunan Perumahan (PP)", "PT. Indah Karya (Persero)",
         "Kementerian PUPR",
         "Ruas Jalan Raya Citarik KM 14, Kabupaten Bekasi",
         "PK.02.01/BBPJN-JABAR/2026/12", "2026-02-15", 75000000000))
    p2 = cur.lastrowid
    cur.executemany("INSERT INTO master_wbs (proyek_id,kode_wbs,nama_pekerjaan,bobot_persen) VALUES (%s,%s,%s,%s)",
                    [(p2,"2.1","Pekerjaan Tanah & Drainase",15),
                     (p2,"2.2","Struktur Pondasi Borepile & Abutment",25),
                     (p2,"2.3","Erection Girder & Plat Lantai Jembatan",35),
                     (p2,"2.4","Perkerasan Jalan Oprit & Finishing Aspal",25)])
    cur.executemany("INSERT INTO master_budget (proyek_id,kategori,nama_item,anggaran_pagu,realisasi_biaya) VALUES (%s,%s,%s,%s,%s)",
                    [(p2,"MATERIAL","Baja Profil Wide Flange & Beton K-500",30000000000,28500000000),
                     (p2,"PERALATAN","Sewa Crawler Crane 150 Ton & Trailer",15000000000,14200000000),
                     (p2,"JASA SUBCONT","Pengujian Pile Load Test & NDT Sambungan Las",10000000000,9800000000),
                     (p2,"UPAH KERJA","Mandor, Welder Bersertifikat, & Pekerja",20000000000,19500000000)])
    cur.execute("INSERT INTO laporan_harian (proyek_id,tanggal_laporan,cuaca) VALUES (%s,%s,%s)",
                (p2,"2026-03-27","Hujan Ringan (26 C)"))
    l2 = cur.lastrowid
    cur.execute("INSERT INTO tenaga_kerja (laporan_id,jenis_pekerja,jumlah,hadir,tidak_hadir,keterangan) VALUES (%s,%s,%s,%s,%s,%s)",
                (l2,"Operator & Welder",34,32,2,"Aman terkendali"))
    cur.execute("INSERT INTO peralatan (laporan_id,nama_alat,jumlah_alat,kondisi,keterangan) VALUES (%s,%s,%s,%s,%s)",
                (l2,"Crawler Crane 150T",1,"Baik","Standby di lokasi"))
    cur.execute("INSERT INTO material (laporan_id,nama_material,volume_datang,satuan,keterangan) VALUES (%s,%s,%s,%s,%s)",
                (l2,"Baja Wide Flange WF 600",8.0,"buah","Untuk erection girder"))
    cur.execute("INSERT INTO pekerjaan (laporan_id,jenis_pekerjaan,lokasi,vol_harian,proses_kumulatif,target_harian,keterangan) VALUES (%s,%s,%s,%s,%s,%s,%s)",
                (l2,"Pemasangan Girder Baja Bentang Tengah","Abutment 1 ke 2","2 buah","42%","2 buah","Berjalan lancar"))
    cur.execute("INSERT INTO kondisi_lapangan (laporan_id,akses,k3,kondisi_fisik,hambatan,tak_terencana) VALUES (%s,%s,%s,%s,%s,%s)",
                (l2,"Cukup Licin","Full APD & Safety Briefing","Basah","Hujan pagi hari","Penambahan sirtu jalan kerja"))
    cur.execute("INSERT INTO pengesahan (laporan_id,nama_pembuat,nama_penyetuju) VALUES (%s,%s,%s)",
                (l2,"Dian Pratama, S.T.","Ir. Suhendra"))

    # === PROYEK 3 ===
    cur.execute(
        "INSERT INTO master_proyek (nama_kegiatan,nama_proyek,penyedia_jasa,konsultan,"
        "pemilik_proyek,lokasi,no_kontrak,tgl_mulai,pagu_kontrak_total) "
        "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)",
        ("Penyediaan Infrastruktur Sanitasi & Lingkungan Modern",
         "IPAL Terpadu Kawasan Industri Surya Cipta",
         "PT. Wijaya Karya (Wika) Tirta", "PT. Yodya Karya (Persero)",
         "PT. Surya Cipta Swadaya",
         "Kawasan Industri Karawang Internasional",
         "089/SCS-IPAL/I/2026", "2026-01-10", 30000000000))
    p3 = cur.lastrowid
    cur.executemany("INSERT INTO master_wbs (proyek_id,kode_wbs,nama_pekerjaan,bobot_persen) VALUES (%s,%s,%s,%s)",
                    [(p3,"3.1","Pekerjaan Sipil Bak & Reservoir",40),
                     (p3,"3.2","Instalasi Pervious & Piping System",20),
                     (p3,"3.3","Pengadaan & Pemasangan Mesin M/E",30),
                     (p3,"3.4","Commissioning & Uji Laboratorium",10)])
    cur.executemany("INSERT INTO master_budget (proyek_id,kategori,nama_item,anggaran_pagu,realisasi_biaya) VALUES (%s,%s,%s,%s,%s)",
                    [(p3,"MATERIAL","Pipa HDPE, Valve, & Bahan Kimia Treatment",10000000000,9600000000),
                     (p3,"JASA SUBCONT","Spesialis Panel Kontrol & Software SCADA",8000000000,7700000000),
                     (p3,"PERALATAN","Pompa Dewatering, Generator Set, & Crane",6000000000,5900000000),
                     (p3,"UPAH KERJA","Tenaga Ahli M/E, Plumber, & Tukang Sipil",6000000000,5800000000)])
    cur.execute("INSERT INTO laporan_harian (proyek_id,tanggal_laporan,cuaca) VALUES (%s,%s,%s)",
                (p3,"2026-03-10","Cerah (32 C)"))
    l3 = cur.lastrowid
    cur.execute("INSERT INTO tenaga_kerja (laporan_id,jenis_pekerja,jumlah,hadir,tidak_hadir,keterangan) VALUES (%s,%s,%s,%s,%s,%s)",
                (l3,"Teknisi M/E & Plumber",52,51,1,"Disiplin tinggi"))
    cur.execute("INSERT INTO peralatan (laporan_id,nama_alat,jumlah_alat,kondisi,keterangan) VALUES (%s,%s,%s,%s,%s)",
                (l3,"Pompa Dewatering 4 inch",3,"Baik","Beroperasi 24 jam"))
    cur.execute("INSERT INTO material (laporan_id,nama_material,volume_datang,satuan,keterangan) VALUES (%s,%s,%s,%s,%s)",
                (l3,"Pipa HDPE 6 inch",120.0,"meter","SNI standar"))
    cur.execute("INSERT INTO pekerjaan (laporan_id,jenis_pekerjaan,lokasi,vol_harian,proses_kumulatif,target_harian,keterangan) VALUES (%s,%s,%s,%s,%s,%s,%s)",
                (l3,"Instalasi Pipa HDPE & Water Retaining Test","Zona Pengolahan Awal","120 meter","65%","100 meter","Lolos uji kedap air"))
    cur.execute("INSERT INTO kondisi_lapangan (laporan_id,akses,k3,kondisi_fisik,hambatan,tak_terencana) VALUES (%s,%s,%s,%s,%s,%s)",
                (l3,"Sangat Baik","Standar K3 Confined Space Ketat","Kering","Nihil","-"))
    cur.execute("INSERT INTO pengesahan (laporan_id,nama_pembuat,nama_penyetuju) VALUES (%s,%s,%s)",
                (l3,"Rian Hidayat, S.T.","Ir. Bambang Pamungkas"))

    conn.commit()
    cur.close()
    conn.close()
    print("")
    print("=" * 60)
    print("SUKSES: 3 proyek + data contoh berhasil dimasukkan!")
    print("Login default -> Username: admin | Password: admin123")
    print("=" * 60)


if __name__ == '__main__':
    seed_database()
