import mysql.connector
from datetime import datetime, timedelta
import random

def seed_1year_complex_data():
    # Koneksi ke database MySQL lokal (XAMPP / MySQL Server)
    conn = mysql.connector.connect(
        host="127.0.0.1", port=3306, user="root", password="", database="db_proyek"
    )
    cursor = conn.cursor()
    
    print("Memulai seeding data komprehensif 1 tahun penuh untuk 3 proyek besar...")
    
    # 1. Bersihkan data lama terkait laporan dan transaksi agar bersih dari awal
    tables = [
        'tenaga_kerja', 'peralatan', 'material', 'pekerjaan', 
        'kondisi_lapangan', 'pengesahan', 'laporan_harian', 
        'master_budget', 'master_wbs', 'master_proyek'
    ]
    for t in tables:
        try:
            cursor.execute(f"DELETE FROM {t}")
        except Exception as e:
            print(f"Info tabel {t}: {e}")
    conn.commit()

    # 2. Masukkan 3 Proyek Besar
    proyek_data = [
        (
            "Pembangunan Infrastruktur Pendidikan Tinggi",
            "Gedung Kuliah 5 Lantai Fakultas Teknik",
            "PT. Adhi Karya (Persero) Tbk",
            "PT. Virama Karya (Persero)",
            "Yayasan Pendidikan Cendekia Utama",
            "Jl. Boulevard Universitas No. 45, Depok",
            "042/KONTRAK-FTEK/VIII/2026",
            "2026-01-05",
            45000000000.00
        ),
        (
            "Peningkatan Infrastruktur Transportasi Wilayah Selatan",
            "Revitalisasi Jembatan Citarik & Jalan Penghubung",
            "PT. Pembangunan Perumahan (PP) Tbk",
            "PT. Indah Karya (Persero)",
            "Kementerian PUPR",
            "Ruas Jalan Raya Citarik KM 14, Kabupaten Bekasi",
            "PK.02.01/BBPJN-JABAR/2026/12",
            "2026-01-05",
            75000000000.00
        ),
        (
            "Penyediaan Infrastruktur Sanitasi & Lingkungan Modern",
            "IPAL Terpadu Kawasan Industri Surya Cipta",
            "PT. Wijaya Karya (Wika) Tirta",
            "PT. Yodya Karya (Persero)",
            "PT. Surya Cipta Swadaya",
            "Kawasan Industri Karawang Internasional",
            "089/SCS-IPAL/I/2026",
            "2026-01-05",
            30000000000.00
        )
    ]

    proyek_ids = []
    for p in proyek_data:
        cursor.execute("""
            INSERT INTO master_proyek (nama_kegiatan, nama_proyek, penyedia_jasa, konsultan, pemilik_proyek, lokasi, no_kontrak, tgl_mulai, pagu_kontrak_total)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, p)
        proyek_ids.append(cursor.lastrowid)
    conn.commit()
    print(f"Berhasil memasukkan 3 Proyek Master dengan ID: {proyek_ids}")

    # 3. Alokasi Kategori Budget & WBS per Proyek
    kategori_list = ['MATERIAL', 'UPAH KERJA', 'PERALATAN', 'JASA SUBCONT', 'OPERASIONAL', 'LAINNYA']
    
    for pid in proyek_ids:
        # WBS (Work Breakdown Structure)
        wbs_items = [
            (pid, "1.1", "Pekerjaan Persiapan & Pondasi", 20.00),
            (pid, "1.2", "Pekerjaan Struktur Utama", 35.00),
            (pid, "1.3", "Pekerjaan Arsitektur & Finishing", 30.00),
            (pid, "1.4", "Pekerjaan MEP (Mechanical, Electrical, Plumbing)", 15.00)
        ]
        cursor.executemany("INSERT INTO master_wbs (proyek_id, kode_wbs, nama_pekerjaan, bobot_persen) VALUES (%s, %s, %s, %s)", wbs_items)
        
        # Anggaran Pagu per Kategori
        for kat in kategori_list:
            pagu_val = random.randint(4000000000, 12000000000)
            cursor.execute("UPDATE kategori_budget SET anggaran_pagu = %s WHERE nama_kategori = %s AND proyek_id = %s", (pagu_val, kat, pid))
            if cursor.rowcount == 0:
                cursor.execute("INSERT INTO kategori_budget (proyek_id, nama_kategori, anggaran_pagu) VALUES (%s, %s, %s)", (pid, kat, pagu_val))
    conn.commit()

    # 4. Generate Laporan Harian 1 Tahun Penuh (Hari Kerja Aktif Senin s.d. Jumat Tahun 2026)
    start_date = datetime(2026, 1, 5)
    end_date = datetime(2026, 12, 31)
    
    current = start_date
    working_days = []
    while current <= end_date:
        if current.weekday() < 5: # Senin sampai Jumat
            working_days.append(current)
        current += timedelta(days=1)
        
    print(f"Total hari kerja aktif tahun 2026 yang akan digenerate: {len(working_days)} hari per proyek.")

    cuaca_options = ["Cerah", "Berawan", "Hujan Ringan", "Cerah Berawan"]
    progres_tracker = {proyek_ids[0]: 1.0, proyek_ids[1]: 1.0, proyek_ids[2]: 1.0}

    print("Memasukkan laporan harian detail untuk setiap hari kerja (1 tahun penuh)...")
    
    for idx, day in enumerate(working_days):
        tgl_str = day.strftime('%Y-%m-%d')
        
        for pid in proyek_ids:
            cuaca = random.choice(cuaca_options)
            cursor.execute("INSERT INTO laporan_harian (proyek_id, tanggal_laporan, cuaca) VALUES (%s, %s, %s)", (pid, tgl_str, cuaca))
            laporan_id = cursor.lastrowid
            
            # Tenaga Kerja
            jml_pekerja = random.randint(50, 90)
            hadir = jml_pekerja - random.randint(0, 2)
            tidak_hadir = jml_pekerja - hadir
            cursor.execute("INSERT INTO tenaga_kerja (laporan_id, jenis_pekerja, jumlah, hadir, tidak_hadir, keterangan) VALUES (%s, %s, %s, %s, %s, %s)",
                           (laporan_id, "Tukang, Mandor & Pekerja Terampil", jml_pekerja, hadir, tidak_hadir, "Personil lengkap dan produktif"))
            
            # Peralatan
            cursor.execute("INSERT INTO peralatan (laporan_id, nama_alat, jumlah_alat, kondisi, keterangan) VALUES (%s, %s, %s, %s, %s)",
                           (laporan_id, "Tower Crane, Excavator & Concrete Pump", random.randint(3, 6), "Baik", "Beroperasi normal"))
            
            # Material
            cursor.execute("INSERT INTO material (laporan_id, nama_material, volume_datang, satuan, keterangan) VALUES (%s, %s, %s, %s, %s)",
                           (laporan_id, "Beton Ready Mix / Besi Tulangan / Pipa", round(random.uniform(25.0, 180.0), 2), "m3 / Ton", "Sesuai spesifikasi teknis kontrak"))
            
            # Pekerjaan & Progres Kumulatif bertahap hingga 100%
            progres_tracker[pid] = min(100.0, progres_tracker[pid] + round(random.uniform(0.32, 0.38), 2))
            vol_h = f"{random.randint(20, 60)} m3 / unit"
            kum = f"{progres_tracker[pid]:.1f}%"
            target = f"{(progres_tracker[pid] + 0.35):.1f}%"
            
            cursor.execute("INSERT INTO pekerjaan (laporan_id, jenis_pekerjaan, lokasi, vol_harian, proses_kumulatif, target_harian, keterangan) VALUES (%s, %s, %s, %s, %s, %s, %s)",
                           (laporan_id, f"Pekerjaan Struktur & Finishing Zona {pid}", "Area Proyek Utama", vol_h, kum, target, "Sesuai schedule kurva S"))
            
            # Kondisi Lapangan
            cursor.execute("INSERT INTO kondisi_lapangan (laporan_id, akses, k3, kondisi_fisik, hambatan, tak_terencana) VALUES (%s, %s, %s, %s, %s, %s)",
                           (laporan_id, "Lancar dan aman untuk mobilitas material", "Full APD & Safety Briefing Pagi rutin", "Kering & Siap Kerja", "Nihil", "-"))
            
            # Pengesahan
            cursor.execute("INSERT INTO pengesahan (laporan_id, nama_pembuat, nama_penyetuju) VALUES (%s, %s, %s)",
                           (laporan_id, f"Ir. Pelaksana Lapangan {pid}", f"Ir. Konsultan Pengawas {pid}"))

        if (idx + 1) % 50 == 0:
            conn.commit()
            print(f"Progress Laporan Harian: {idx + 1}/{len(working_days)} hari selesai...")
            
    conn.commit()
    print("Selesai memasukkan laporan harian 1 tahun penuh.")

    # 5. Generate Ratusan Transaksi Keuangan Rinci (Petty Cash & Realisasi Biaya)
    print("Memasukkan ratusan transaksi keuangan rinci (kas masuk dan kas keluar)...")
    
    # Kas Masuk Awal / Termin untuk masing-masing proyek
    for pid in proyek_ids:
        cursor.execute("""
            INSERT INTO master_budget (proyek_id, kategori, nama_item, anggaran_pagu, realisasi_biaya, no_transaksi, tanggal_transaksi, jenis_kas, pembuat, kuantitas, harga_satuan)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (pid, "LAINNYA", "Termin Pembayaran 1 / Modal Kerja Proyek", 20000000000.00, 20000000000.00, f"TRX-IN-{pid}-001", "2026-01-02", "masuk", "Direktur Keuangan", 1, 20000000000.00))

    # Daftar master transaksi pengeluaran realistis
    item_descs = [
        ("Pembelian Semen Gresik / Tiga Roda (Sak)", "MATERIAL", 600, 125000),
        ("Pembelian Besi Beton Polos & Ulir SNI", "MATERIAL", 200, 9000000),
        ("Pengadaan Pasir Pasang & Batu Split", "MATERIAL", 120, 375000),
        ("Upah Borongan Pekerja & Tukang Mingguan", "UPAH KERJA", 75, 180000),
        ("Gaji Bulanan Mandor & Ahli K3", "UPAH KERJA", 6, 5000000),
        ("Sewa Tower Crane Bulanan", "PERALATAN", 1, 80000000),
        ("Sewa Excavator PC200 & Solar Operasional", "PERALATAN", 2, 28000000),
        ("Sewa Pompa Beton & Generator Set", "PERALATAN", 1, 20000000),
        ("Jasa Subcon Pengujian Soil Test & NDT", "JASA SUBCONT", 1, 40000000),
        ("Jasa Subcon Pemasangan Struktur Rangka", "JASA SUBCONT", 1, 150000000),
        ("Biaya Operasional Kantor Lapangan & Konsumsi", "OPERASIONAL", 35, 600000),
        ("Biaya Perizinan, Koordinasi & Keamanan", "OPERASIONAL", 1, 18000000),
        ("Pembelian Alat Pelindung Diri (APD) K3", "LAINNYA", 60, 160000),
        ("Alat Tulis Kantor & Dokumentasi Drone Proyek", "LAINNYA", 1, 3000000)
    ]

    trx_counter = 100
    for pid in proyek_ids:
        # Buat 130 transaksi keuangan rinci per proyek sepanjang tahun 2026
        current_d = datetime(2026, 1, 10)
        for i in range(130):
            current_d += timedelta(days=random.choice([1, 2, 3]))
            if current_d > datetime(2026, 12, 28):
                current_d = datetime(2026, 12, 28)
            
            tgl_str = current_d.strftime('%Y-%m-%d')
            item = random.choice(item_descs)
            nama_item = item[0]
            kategori = item[1]
            qty = item[2]
            harga = item[3]
            subtotal = qty * harga
            no_trx = f"TRX-2026-{trx_counter:04d}"
            trx_counter += 1
            
            cursor.execute("""
                INSERT INTO master_budget (proyek_id, kategori, nama_item, anggaran_pagu, realisasi_biaya, no_transaksi, tanggal_transaksi, jenis_kas, pembuat, kuantitas, harga_satuan)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (pid, kategori, nama_item, 0.00, subtotal, no_trx, tgl_str, "keluar", "Petty Cash Officer", qty, harga))

    conn.commit()
    cursor.close()
    conn.close()
    
    print("=" * 70)
    print("SUKSES: Seluruh data laporan harian 1 tahun penuh & ratusan transaksi keuangan untuk 3 proyek berhasil dimasukkan!")
    print("=" * 70)

if __name__ == '__main__':
    seed_1year_complex_data()