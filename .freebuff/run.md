# MANPRO - Preview Run Doc

## Reproduce Artifacts
1. Ensure MySQL is running on localhost:3306 (root, no password)
2. Create database if not exists: `CREATE DATABASE IF NOT EXISTS db_proyek`
3. Install dependencies: `pip install Flask mysql-connector-python openpyxl Werkzeug gunicorn`

## Run Server
```bash
cd D:\Software\BelajarPython\daily_report_proyek_web
set PORT=5050
set SECRET_KEY=previewkey123
python daily.py
```

The app auto-migrates tables on first run. Default admin: `admin` / `admin123`
