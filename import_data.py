#!/usr/bin/env python3
"""
Import Data dari File SQL Backup ke MySQL Lokal
Jalankan: python import_data.py backup_tidb_YYYYMMDD_HHMMSS.sql
"""

import mysql.connector
import sys
import os

# Koneksi ke MySQL Lokal
DB_CONFIG = {
    'host': '127.0.0.1',
    'port': 3306,
    'user': 'root',
    'password': '',
    'database': 'db_proyek',
}

def import_sql(sql_file):
    print("=" * 60)
    print("  IMPORT DATA KE MySQL LOKAL")
    print("=" * 60)
    
    if not os.path.exists(sql_file):
        print(f"❌ File tidak ditemukan: {sql_file}")
        return
    
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor()
        print(f"✅ Berhasil terhubung ke MySQL Lokal!")
    except Exception as e:
        print(f"❌ Gagal koneksi: {e}")
        return
    
    print(f"📄 Membaca file: {sql_file}")
    
    with open(sql_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Split per statement
    statements = content.split(';')
    total = len([s for s in statements if s.strip() and not s.strip().startswith('--')])
    done = 0
    errors = 0
    
    for stmt in statements:
        stmt = stmt.strip()
        if not stmt or stmt.startswith('--'):
            continue
        
        try:
            cursor.execute(stmt)
            done += 1
            if done % 50 == 0:
                print(f"  ⏳ Progress: {done}/{total} statements...")
        except Exception as e:
            errors += 1
            if errors <= 3:
                print(f"  ⚠️ Skip: {str(e)[:80]}")
    
    conn.commit()
    cursor.close()
    conn.close()
    
    print(f"\n{'=' * 60}")
    print(f"✅ Import selesai!")
    print(f"   ✅ Berhasil: {done} statements")
    print(f"   ⚠️ Error/Skip: {errors} statements")
    print(f"{'=' * 60}")

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Cara pakai:")
        print("  python import_data.py backup_tidb_20260830_180000.sql")
        print("\nFile backup ada di folder ini:")
        for f in sorted(os.listdir('.')):
            if f.startswith('backup_tidb_') and f.endswith('.sql'):
                print(f"  📄 {f}")
    else:
        import_sql(sys.argv[1])
