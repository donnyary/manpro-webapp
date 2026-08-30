#!/usr/bin/env python3
"""
MANPRO Auto Deploy Watcher
==========================
Monitor perubahan file dan otomatis push ke GitHub.

Cara pakai:
    python auto_deploy.py
    
Tekan Ctrl+C untuk berhenti.
"""

import time
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path

# Konfigurasi
WATCH_EXTENSIONS = {'.py', '.html', '.css', '.js', '.sql', '.txt', '.md'}
IGNORE_FOLDERS = {'__pycache__', '.git', '.freebuff', 'node_modules', 'uploads'}
DEBOUNCE_SECONDS = 5  # Tunggu 5 detik setelah perubahan terakhir
COMMIT_PREFIX = "Auto"

# Warna
class Colors:
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    CYAN = '\033[96m'
    BLUE = '\033[94m'
    BOLD = '\033[1m'
    END = '\033[0m'

def print_header():
    print(f"""
{Colors.CYAN}╔══════════════════════════════════════════════════════════════╗
║                                                              ║
║   🔄 MANPRO AUTO DEPLOY WATCHER                             ║
║   Monitor perubahan & auto-push ke GitHub                    ║
║                                                              ║
║   Tekan Ctrl+C untuk berhenti                                ║
║                                                              ║
╚══════════════════════════════════════════════════════════════╝{Colors.END}
""")

def run_cmd(cmd, cwd=None):
    """Jalankan command."""
    try:
        result = subprocess.run(
            cmd, shell=True, capture_output=True, text=True, cwd=cwd,
            encoding='utf-8', errors='replace'
        )
        return result.returncode == 0, result.stdout.strip(), result.stderr.strip()
    except Exception as e:
        return False, '', str(e)

def get_changed_files(project_dir):
    """Ambil daftar file yang berubah."""
    success, output, _ = run_cmd('git status --short', cwd=project_dir)
    if not success:
        return []
    
    files = []
    for line in output.split('\n'):
        line = line.strip()
        if line and len(line) > 3:
            # Ambil nama file (hapus status di depan)
            file_path = line[3:].strip().strip('"')
            files.append(file_path)
    return files

def should_ignore(file_path):
    """Cek apakah file harus di-ignore."""
    for folder in IGNORE_FOLDERS:
        if folder in file_path:
            return True
    
    # Cek ekstensi
    ext = Path(file_path).suffix.lower()
    if ext not in WATCH_EXTENSIONS:
        return True
    
    return False

def auto_commit_and_push(project_dir, changed_files):
    """Auto commit dan push."""
    timestamp = datetime.now().strftime('%H:%M:%S')
    
    # Buat commit message
    if len(changed_files) == 1:
        msg = f"{COMMIT_PREFIX}: {changed_files[0]}"
    else:
        msg = f"{COMMIT_PREFIX}: {len(changed_files)} files updated"
    
    print(f"\n{Colors.YELLOW}{'='*60}{Colors.END}")
    print(f"{Colors.BOLD}🚀 AUTO DEPLOY - {timestamp}{Colors.END}")
    print(f"{Colors.YELLOW}{'='*60}{Colors.END}")
    
    # Tampilkan file yang berubah
    print(f"\n{Colors.BLUE}📋 File yang berubah:{Colors.END}")
    for f in changed_files[:10]:  # Tampilkan max 10 file
        print(f"   • {f}")
    if len(changed_files) > 10:
        print(f"   ... dan {len(changed_files) - 10} file lainnya")
    
    # Git add
    print(f"\n{Colors.YELLOW}📦 Git add...{Colors.END}")
    success, _, err = run_cmd('git add .', cwd=project_dir)
    if not success:
        print(f"{Colors.RED}   ❌ Error: {err}{Colors.END}")
        return False
    print(f"{Colors.GREEN}   ✅ OK{Colors.END}")
    
    # Git commit
    print(f"{Colors.YELLOW}💾 Git commit...{Colors.END}")
    success, _, err = run_cmd(f'git commit -m "{msg}"', cwd=project_dir)
    if not success:
        if 'nothing to commit' in err:
            print(f"{Colors.YELLOW}   ⚠️  Tidak ada perubahan{Colors.END}")
            return False
        print(f"{Colors.RED}   ❌ Error: {err}{Colors.END}")
        return False
    print(f"{Colors.GREEN}   ✅ OK{Colors.END}")
    
    # Git push
    print(f"{Colors.YELLOW}⬆️  Git push...{Colors.END}")
    success, output, err = run_cmd('git push', cwd=project_dir)
    if not success:
        print(f"{Colors.RED}   ❌ Error: {err}{Colors.END}")
        return False
    
    if 'Everything up-to-date' in output:
        print(f"{Colors.YELLOW}   ⚠️  Sudah up-to-date{Colors.END}")
    else:
        print(f"{Colors.GREEN}   ✅ OK{Colors.END}")
        print(f"\n{Colors.GREEN}🎉 BERHASIL! Railway akan auto-deploy dalam 1-2 menit{Colors.END}")
    
    return True

def main():
    print_header()
    
    project_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"{Colors.BLUE}📁 Monitoring:{Colors.END} {project_dir}\n")
    
    # Cek git
    success, _, _ = run_cmd('git --version', cwd=project_dir)
    if not success:
        print(f"{Colors.RED}❌ Git belum terinstall!{Colors.END}")
        return
    
    print(f"{Colors.GREEN}✅ Git OK{Colors.END}")
    print(f"{Colors.GREEN}✅ Auto-deploy watcher aktif{Colors.END}")
    print(f"\n{Colors.CYAN}👀 Menunggu perubahan file...{Colors.END}")
    print(f"{Colors.YELLOW}(Tekan Ctrl+C untuk berhenti){Colors.END}\n")
    
    last_deploy_time = 0
    last_file_state = {}
    
    try:
        while True:
            # Ambil state file saat ini
            current_state = {}
            for root, dirs, files in os.walk(project_dir):
                # Skip ignored folders
                dirs[:] = [d for d in dirs if d not in IGNORE_FOLDERS]
                
                for file in files:
                    file_path = os.path.join(root, file)
                    rel_path = os.path.relpath(file_path, project_dir)
                    
                    if not should_ignore(rel_path):
                        try:
                            mtime = os.path.getmtime(file_path)
                            current_state[rel_path] = mtime
                        except:
                            pass
            
            # Deteksi perubahan
            changed_files = []
            for file_path, mtime in current_state.items():
                if file_path not in last_file_state or last_file_state[file_path] != mtime:
                    changed_files.append(file_path)
            
            # Jika ada perubahan dan sudah lewat debounce time
            current_time = time.time()
            if changed_files and (current_time - last_deploy_time) > DEBOUNCE_SECONDS:
                # Filter file yang benar-benar berubah
                real_changes = [f for f in changed_files if not should_ignore(f)]
                
                if real_changes:
                    if auto_commit_and_push(project_dir, real_changes):
                        last_deploy_time = current_time
                    
                    print(f"\n{Colors.CYAN}👀 Menunggu perubahan file berikutnya...{Colors.END}\n")
            
            last_file_state = current_state
            time.sleep(2)  # Check setiap 2 detik
            
    except KeyboardInterrupt:
        print(f"\n\n{Colors.YELLOW}👋 Auto-deploy watcher dihentikan.{Colors.END}\n")

if __name__ == '__main__':
    main()
