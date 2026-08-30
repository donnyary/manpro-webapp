#!/usr/bin/env python3
"""
Export Data dari TiDB Cloud ke File SQL Lokal
Jalankan: python export_data.py
"""

import mysql.connector
import os
import sys
from datetime import datetime

# Fix Windows console encoding
sys.stdout.reconfigure(encoding='utf-8')

# Koneksi ke TiDB Cloud (sama seperti daily.py)
DB_CONFIG = {
    'host': 'gateway01.ap-southeast-1.prod.aws.tidbcloud.com',
    'port': 4000,
    'user': 'vqpHmv4RwShLMQa.root',
    'password': 'R2HE6uHS62HSOtMS',
    'database': 'db_proyek',
    'ssl_disabled': False,
    'ssl_verify_cert': False,
    'ssl_verify_identity': False,
}

# Tabel yang akan di-export
TABLES = [
    'master_proyek', 'master_wbs', 'kategori_budget', 'master_budget',
    'users', 'menu_permissions',
    'laporan_harian', 'tenaga_kerja', 'peralatan', 'material',
    'pekerjaan', 'kondisi_lapangan', 'pengesahan'
]

def export():
    print("=" * 60)
    print("  EXPORT DATA DARI TiDB CLOUD")
    print("=" * 60)
    
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor(dictionary=True)
        print(f"✅ Berhasil terhubung ke TiDB Cloud!")
    except Exception as e:
        print(f"❌ Gagal koneksi: {e}")
        return
    
    # Buat file output
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'backup_tidb_{timestamp}.sql'
    
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(f"-- Backup dari TiDB Cloud - {datetime.now()}\n")
        f.write(f"-- Database: db_proyek\n")
        f.write(f"-- Import: mysql -h <host> -P 4000 -u <user> -p db_proyek < {filename}\n\n")
        
        for table in TABLES:
            print(f"📦 Export table: {table}...", end=" ")
            try:
                # Get count
                cursor.execute(f"SELECT COUNT(*) as cnt FROM `{table}`")
                count = cursor.fetchone()['cnt']
                
                # Get columns
                cursor.execute(f"DESCRIBE `{table}`")
                columns = [col['Field'] for col in cursor.fetchall()]
                
                # Get data
                cursor.execute(f"SELECT * FROM `{table}` ORDER BY id")
                rows = cursor.fetchall()
                
                # Write SQL
                f.write(f"\n-- Table: {table} ({count} rows)\n")
                f.write(f"TRUNCATE TABLE `{table}`;\n")
                
                for row in rows:
                    cols_str = ', '.join([f'`{c}`' for c in columns])
                    vals = []
                    for c in columns:
                        v = row.get(c)
                        if v is None:
                            vals.append('NULL')
                        elif isinstance(v, (int, float)):
                            vals.append(str(v))
                        else:
                            vals.append(f"'{str(v).replace(chr(39), chr(39)+chr(39))}'")
                    vals_str = ', '.join(vals)
                    f.write(f"INSERT INTO `{table}` ({cols_str}) VALUES ({vals_str});\n")
                
                print(f"✅ {count} baris")
            except Exception as e:
                print(f"❌ Error: {e}")
    
    cursor.close()
    conn.close()
    
    # Summary
    size = os.path.getsize(filename)
    print(f"\n{'=' * 60}")
    print(f"📁 File: {filename}")
    print(f"📊 Ukuran: {size / 1024:.1f} KB")
    print(f"✅ Export selesai!")
    print(f"\nCara import ke MySQL lokal:")
    print(f"  mysql -u root -p db_proyek < {filename}")
    print(f"{'=' * 60}")

if __name__ == '__main__':
    export()
