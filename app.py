from flask import Flask, render_template, request, redirect, url_for, flash, session
import mysql.connector
import os
import secrets
from functools import wraps

app = Flask(__name__)
app.secret_key = secrets.token_hex(32)

def get_db_connection():
    return mysql.connector.connect(
        host="127.0.0.1",
        port=3306,
        user="root",        # Ganti sesuai user MySQL Anda
        password="",        # Ganti sesuai password MySQL Anda
        database="db_proyek"
    )

# Bypass Login (Hanya untuk keperluan demonstrasi)
@app.route('/')
def index():
    session['logged_in'] = True
    session['username'] = 'admin'
    return redirect(url_for('gantt_chart', proyek_id=1))

# ========================================================
# ROUTE TIME SCHEDULE & KURVA S (GANTT)
# ========================================================
@app.route('/gantt/<int:proyek_id>')
def gantt_chart(proyek_id):
    if not session.get('logged_in'):
        return "Harap login terlebih dahulu!"

    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    
    # 1. Ambil detail proyek
    cursor.execute("SELECT * FROM master_proyek WHERE id = %s", (proyek_id,))
    proyek = cursor.fetchone()
    if not proyek:
        return "Data proyek tidak ditemukan!"
    
    # 2. Ambil data rencana RAB / WBS
    cursor.execute("SELECT * FROM master_wbs WHERE proyek_id = %s ORDER BY id ASC", (proyek_id,))
    wbs_data = cursor.fetchall()
    
    # 3. Setup Variabel Kalkulasi Time Schedule
    minggu_total = 12 # Proyek dijadwalkan selesai dalam 12 minggu
    rencana_mingguan = [0] * minggu_total
    
    # Penjadwalan: Kapan mulai (minggu ke-) dan berapa lama durasinya (minggu)
    # Mapping Format: 'kode_wbs': (mulai_minggu, durasi_minggu)
    schedule_map = {
        '1.1': (1, 2),  # Persiapan: M1-M2
        '1.2': (2, 3),  # Pondasi: M2-M4
        '1.3': (3, 6),  # Dinding: M3-M8
        '1.4': (4, 5),  # Beton: M4-M8
        '1.5': (7, 3),  # Atap: M7-M9
        '1.6': (8, 3),  # Plafond: M8-M10
        '1.7': (9, 3),  # Lantai: M9-M11
        '1.8': (10, 3), # Landscape: M10-M12
        '1.9': (8, 4),  # Pintu & Jendela: M8-M11
        '1.10': (6, 6)  # MEP: M6-M11
    }
    
    # Hitung bobot mingguan yang didistribusikan
    for item in wbs_data:
        kode = item.get('kode_wbs')
        bobot = float(item.get('bobot_persen', 0))
        
        # Simpan property durasi & mulai_minggu ke dict agar bisa dibaca di HTML
        item['start_week'] = 0
        item['duration'] = 0

        if kode in schedule_map:
            start_wk, duration = schedule_map[kode]
            item['start_week'] = start_wk
            item['duration'] = duration
            weekly_val = bobot / duration
            
            for w in range(start_wk, start_wk + duration):
                if 1 <= w <= minggu_total:
                    rencana_mingguan[w-1] += weekly_val

    # 4. Hitung Rencana Kumulatif (Kurva S Rencana)
    rencana_kumulatif = []
    cum_val = 0
    for val in rencana_mingguan:
        cum_val += val
        rencana_kumulatif.append(round(cum_val, 2))
        
    # Membulatkan nilai akhir proyek (koreksi presisi desimal agar 100%)
    if rencana_kumulatif and rencana_kumulatif[-1] > 99.0:
        rencana_kumulatif[-1] = 100.0

    # 5. Ambil Realisasi dari Laporan Harian (Kurva S Aktual)
    sql_kurva = """
        SELECT YEARWEEK(lh.tanggal_laporan, 1) as minggu, 
               MAX(CAST(REPLACE(p.proses_kumulatif, '%', '') AS DECIMAL(5,2))) as progres
        FROM laporan_harian lh 
        JOIN pekerjaan p ON lh.id = p.laporan_id
        WHERE lh.proyek_id = %s 
        GROUP BY minggu 
        ORDER BY minggu ASC
    """
    cursor.execute(sql_kurva, (proyek_id,))
    data_harian = cursor.fetchall()
    
    labels = [f"Minggu {i+1}" for i in range(minggu_total)]
    
    # Proses realisasi disesuaikan dengan indeks minggu
    realisasi = []
    current_realisasi = 0
    for idx in range(len(data_harian)):
        try:
            current_realisasi = float(data_harian[idx]['progres'])
        except:
            pass
        realisasi.append(current_realisasi)

    conn.close()
    
    return render_template('gantt.html', 
                           proyek=proyek, 
                           labels=labels, 
                           realisasi=realisasi, 
                           rencana=rencana_kumulatif,
                           wbs_data=wbs_data,
                           rencana_mingguan=rencana_mingguan,
                           minggu_total=minggu_total)

if __name__ == '__main__':
    app.run(debug=True, port=5001)