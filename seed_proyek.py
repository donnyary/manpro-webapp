import mysql.connector

def seed_rumah_1_lantai():
    # Sesuaikan dengan kredensial database Anda
    conn = mysql.connector.connect(
        host="127.0.0.1", port=3306, user="root", password="", database="db_proyek"
    )
    cursor = conn.cursor()

    print("Memasukkan Master Proyek Baru...")
    cursor.execute("""
        INSERT INTO master_proyek 
        (nama_kegiatan, nama_proyek, penyedia_jasa, konsultan, pemilik_proyek, lokasi, no_kontrak, pagu_kontrak_total) 
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
    """, (
        "Pembangunan Tempat Tinggal", 
        "Proyek Pembangunan Rumah 1 Lantai", 
        "PT. Bina Konstruksi", 
        "CV. Arsitektur Mandiri", 
        "Bapak Budi Santoso", 
        "Jl. Kenanga No.10, Jakarta", 
        "001/SPK-RUMAH/2026", 
        500000000.00
    ))
    
    proyek_id = cursor.lastrowid
    
    print("Memasukkan WBS / Rekapitulasi Bobot Pekerjaan...")
    wbs_data = [
        (proyek_id, '1.1', 'Pekerjaan Persiapan', 3.00),
        (proyek_id, '1.2', 'Pekerjaan Pondasi', 13.00),
        (proyek_id, '1.3', 'Pekerjaan Dinding', 5.79),
        (proyek_id, '1.4', 'Pekerjaan Beton Bertulang', 17.00),
        (proyek_id, '1.5', 'Pekerjaan Atap', 12.00),
        (proyek_id, '1.6', 'Pekerjaan Plafond', 8.00),
        (proyek_id, '1.7', 'Pekerjaan Lantai', 11.00),
        (proyek_id, '1.8', 'Pekerjaan Landscape', 9.00),
        (proyek_id, '1.9', 'Pekerjaan Pintu dan Jendela', 9.00),
        (proyek_id, '1.10', 'Pekerjaan MEP', 12.21)
    ]
    
    cursor.executemany("""
        INSERT INTO master_wbs (proyek_id, kode_wbs, nama_pekerjaan, bobot_persen)
        VALUES (%s, %s, %s, %s)
    """, wbs_data)
    
    # Memasukkan data simulasi realisasi lapangan (Dummy data)
    print("Memasukkan Simulasi Data Laporan Harian...")
    cursor.execute("INSERT INTO laporan_harian (proyek_id, tanggal_laporan) VALUES (%s, %s)", (proyek_id, '2026-08-01'))
    laporan_id_1 = cursor.lastrowid
    cursor.execute("INSERT INTO pekerjaan (laporan_id, proses_kumulatif) VALUES (%s, %s)", (laporan_id_1, '1.50%'))
    
    cursor.execute("INSERT INTO laporan_harian (proyek_id, tanggal_laporan) VALUES (%s, %s)", (proyek_id, '2026-08-08'))
    laporan_id_2 = cursor.lastrowid
    cursor.execute("INSERT INTO pekerjaan (laporan_id, proses_kumulatif) VALUES (%s, %s)", (laporan_id_2, '3.00%'))

    conn.commit()
    cursor.close()
    conn.close()
    
    print(f"SUKSES! Data berhasil dimasukkan dengan ID Proyek: {proyek_id}")

if __name__ == '__main__':
    seed_rumah_1_lantai()