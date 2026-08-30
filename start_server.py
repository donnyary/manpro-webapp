import sys, os, time, subprocess, urllib.request

# Gunakan direktori script ini sebagai working directory
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(SCRIPT_DIR)
sys.path.insert(0, SCRIPT_DIR)

LOG_DIR = os.path.join(SCRIPT_DIR, '.freebuff')
os.makedirs(LOG_DIR, exist_ok=True)

LOG_FILE = os.path.join(LOG_DIR, 'preview.log')
LOG_ERR = os.path.join(LOG_DIR, 'preview.log.err')

# Start Flask on port 5001
proc = subprocess.Popen(
    [sys.executable, '-c', f'import sys; sys.path.insert(0, r"{SCRIPT_DIR}"); from daily import app; app.run(debug=False, port=5001, host="127.0.0.1")'],
    stdout=open(LOG_FILE, 'w'),
    stderr=open(LOG_ERR, 'w'),
    creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0),
    cwd=SCRIPT_DIR
)

print(f'PID: {proc.pid}')

# Wait and verify
time.sleep(4)
try:
    resp = urllib.request.urlopen('http://127.0.0.1:5001/login', timeout=5)
    print(f'Server OK: {resp.status}')
except Exception as e:
    print(f'Server ERROR: {e}')
