#!/usr/bin/env python3
"""
MANPRO Auto Deploy Script
=========================
Satu klik untuk push kode ke GitHub dan auto-deploy ke Railway.

Cara pakai:
    python deploy.py
    
Atau klik 2x file deploy.py
"""

import subprocess
import sys
import os
from datetime import datetime

# Warna untuk terminal
class Colors:
    HEADER = '\033[95m'
    BLUE = '\033[94m'
    CYAN = '\033[96m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    END = '\033[0m'
    BOLD = '\033[1m'

def print_header():
    print(f"""
{Colors.CYAN}╔══════════════════════════════════════════════════════════════╗
║                                                              ║
║   🚀 MANPRO AUTO DEPLOY                                     ║
║   Push kode ke GitHub & auto-deploy ke Railway               ║
║                                                              ║
╚══════════════════════════════════════════════════════════════╝{Colors.END}
""")

def run_command(cmd, cwd=None):
    """Jalankan command dan return output."""
    try:
        result = subprocess.run(
            cmd, 
            shell=True, 
            capture_output=True, 
            text=True, 
            cwd=cwd,
            encoding='utf-8',
            errors='replace'
        )
        return result.returncode == 0, result.stdout, result.stderr
    except Exception as e:
        return False, '', str(e)

def check_git():
    """Cek apakah git sudah terinstall."""
    success, _, _ = run_command('git --version')
    return success

def check_remote():
    """Cek apakah remote sudah di-set."""
    success, output, _ = run_command('git remote -v')
    return success and 'origin' in output

def main():
    print_header()
    
    # Folder project
    project_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"{Colors.BLUE}📁 Folder Project:{Colors.END} {project_dir}\n")
    
    # Cek git
    print(f"{Colors.YELLOW}🔍 Cek Git...{Colors.END}")
    if not check_git():
        print(f"{Colors.RED}❌ Git belum terinstall!{Colors.END}")
        print(f"   Download: https://git-scm.com/download/win")
        input("\nTekan Enter untuk keluar...")
        return
    print(f"{Colors.GREEN}✅ Git OK{Colors.END}\n")
    
    # Cek remote
    print(f"{Colors.YELLOW}🔍 Cek Remote...{Colors.END}")
    if not check_remote():
        print(f"{Colors.YELLOW}⚠️  Remote belum di-set. Setting remote...{Colors.END}")
        success, _, err = run_command(
            'git remote add origin https://github.com/donnyary/manpro-webapp.git',
            cwd=project_dir
        )
        if not success and 'already exists' not in str(err):
            print(f"{Colors.RED}❌ Gagal set remote: {err}{Colors.END}")
            input("\nTekan Enter untuk keluar...")
            return
    print(f"{Colors.GREEN}✅ Remote OK{Colors.END}\n")
    
    # Tampilkan perubahan
    print(f"{Colors.YELLOW}📋 Perubahan yang terdeteksi:{Colors.END}")
    success, output, _ = run_command('git status --short', cwd=project_dir)
    if output.strip():
        print(output)
    else:
        print(f"   {Colors.GREEN}Tidak ada perubahan{Colors.END}\n")
        
        # Tanya tetap push
        answer = input(f"{Colors.YELLOW}Tetap push? (y/n): {Colors.END}").lower()
        if answer != 'y':
            print(f"\n{Colors.CYAN}Dibatalkan.{Colors.END}")
            return
    
    # Konfirmasi
    print(f"\n{Colors.YELLOW}═══════════════════════════════════════════════════════════════{Colors.END}")
    print(f"{Colors.BOLD}🚀 Siap deploy ke production?{Colors.END}")
    print(f"{Colors.YELLOW}═══════════════════════════════════════════════════════════════{Colors.END}\n")
    
    answer = input(f"{Colors.YELLOW}Ketik 'DEPLOY' untuk konfirmasi: {Colors.END}")
    if answer != 'DEPLOY':
        print(f"\n{Colors.CYAN}Dibatalkan.{Colors.END}")
        return
    
    # Mulai deploy
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"\n{Colors.CYAN}{'='*60}{Colors.END}")
    print(f"{Colors.BOLD}🚀 DEPLOY DIMULAI - {timestamp}{Colors.END}")
    print(f"{Colors.CYAN}{'='*60}{Colors.END}\n")
    
    # Step 1: Git Add
    print(f"{Colors.YELLOW}📦 Step 1/4: Menambahkan file...{Colors.END}")
    success, _, err = run_command('git add .', cwd=project_dir)
    if success:
        print(f"{Colors.GREEN}   ✅ File ditambahkan{Colors.END}\n")
    else:
        print(f"{Colors.RED}   ❌ Error: {err}{Colors.END}\n")
        input("Tekan Enter untuk keluar...")
        return
    
    # Step 2: Git Commit
    print(f"{Colors.YELLOW}💾 Step 2/4: Commit perubahan...{Colors.END}")
    commit_msg = f"Deploy: {timestamp}"
    success, _, err = run_command(f'git commit -m "{commit_msg}"', cwd=project_dir)
    if success:
        print(f"{Colors.GREEN}   ✅ Commit berhasil{Colors.END}\n")
    else:
        if 'nothing to commit' in err or 'nothing to commit' in _:
            print(f"{Colors.YELLOW}   ⚠️  Tidak ada perubahan untuk di-commit{Colors.END}\n")
        else:
            print(f"{Colors.RED}   ❌ Error: {err}{Colors.END}\n")
            input("Tekan Enter untuk keluar...")
            return
    
    # Step 3: Git Push
    print(f"{Colors.YELLOW}⬆️  Step 3/4: Push ke GitHub...{Colors.END}")
    success, output, err = run_command('git push', cwd=project_dir)
    if success:
        print(f"{Colors.GREEN}   ✅ Push berhasil!{Colors.END}\n")
    else:
        if 'Everything up-to-date' in output:
            print(f"{Colors.YELLOW}   ⚠️  Sudah up-to-date{Colors.END}\n")
        else:
            print(f"{Colors.RED}   ❌ Error: {err}{Colors.END}\n")
            input("Tekan Enter untuk keluar...")
            return
    
    # Step 4: Info
    print(f"{Colors.YELLOW}🔄 Step 4/4: Railway auto-deploy...{Colors.END}")
    print(f"{Colors.GREEN}   ✅ Railway akan auto-deploy dalam 1-2 menit{Colors.END}\n")
    
    # Selesai
    print(f"{Colors.CYAN}{'='*60}{Colors.END}")
    print(f"""
{Colors.GREEN}🎉 DEPLOY BERHASIL!

📋 Yang terjadi:
   ✅ Kode di-push ke GitHub
   ✅ Railway akan auto-deploy dalam 1-2 menit

🌐 URL Aplikasi:
   https://manpro-webapp-production.up.railway.app

📊 Cek Status Deploy:
   https://railway.com → project → manpro-webapp → Deploy Logs
{Colors.END}""")
    
    print(f"{Colors.CYAN}{'='*60}{Colors.END}\n")
    
    input(f"{Colors.YELLOW}Tekan Enter untuk keluar...{Colors.END}")

if __name__ == '__main__':
    main()
