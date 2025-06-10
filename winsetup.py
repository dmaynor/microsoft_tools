#!/usr/bin/env python3
"""
winsetup.py

Fully automated Windows 11 VM setup with QEMU on Windows.
Includes Chocolatey and QEMU installation, ISO download, and VM provisioning.
"""

import argparse
import ctypes
import os
import re
import shutil
import subprocess
import sys
import winreg
from pathlib import Path

# Constants
DEFAULT_DISK_PATH = Path.cwd() / "win11_vm.qcow2"
QEMU_EXE = "qemu-system-x86_64.exe"
QEMU_IMG = "qemu-img.exe"
QEMU_CHOCOLATEY_NAME = "qemu"
CHOCOLATEY_PATH = Path("C:/ProgramData/chocolatey/bin")
WINDOWS_ISO_DOWNLOAD_PAGE = "https://www.microsoft.com/en-us/software-download/windows11"

# ----------------------------- Core Checks ---------------------------------- #

def ensure_admin():
    if not ctypes.windll.shell32.IsUserAnAdmin():
        sys.exit("[!] This script must be run as Administrator. Right-click → Run as Administrator.")

def ensure_requests():
    try:
        import requests
    except ImportError:
        print("[*] Installing 'requests' via pip...")
        try:
            subprocess.run([sys.executable, "-m", "pip", "install", "requests"], check=True)
            import requests
        except Exception as e:
            sys.exit(f"[!] Failed to install requests: {e}")

# ----------------------------- Chocolatey Logic ---------------------------- #

def get_choco_path_from_registry():
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Chocolatey") as key:
            install_dir, _ = winreg.QueryValueEx(key, "ChocolateyInstall")
            choco_path = Path(install_dir) / "bin" / "choco.exe"
            if choco_path.exists():
                return str(choco_path)
    except Exception:
        return None

def detect_choco():
    try:
        test = subprocess.run(["choco", "-v"], capture_output=True, text=True)
        if test.returncode == 0:
            return "choco"
    except FileNotFoundError:
        pass

    reg_path = get_choco_path_from_registry()
    if reg_path:
        return reg_path

    fallback_paths = [
        Path("C:/ProgramData/chocolatey/bin/choco.exe"),
        Path(os.environ.get("LocalAppData", "")) / "choco" / "bin" / "choco.exe",
        Path(os.environ.get("SystemDrive", "C:")) / "choco" / "bin" / "choco.exe"
    ]
    for path in fallback_paths:
        if path.exists():
            return str(path)

    return None

def install_chocolatey():
    print("[*] Installing Chocolatey...")
    choco_cmd = (
        'Set-ExecutionPolicy Bypass -Scope Process -Force; '
        '[System.Net.ServicePointManager]::SecurityProtocol = '
        '[System.Net.SecurityProtocolType]::Tls12; '
        'iex ((New-Object System.Net.WebClient).DownloadString('
        "'https://community.chocolatey.org/install.ps1'))"
    )
    subprocess.run([
        "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-Command", choco_cmd
    ], check=True)
    os.environ["PATH"] = str(CHOCOLATEY_PATH) + os.pathsep + os.environ["PATH"]

def install_qemu_choco():
    print("[*] Installing QEMU via Chocolatey...")

    choco_bin = detect_choco()
    if not choco_bin:
        sys.exit("[!] Could not locate Chocolatey binary. Run 'choco' manually to verify install.")

    print(f"[+] Using Chocolatey at: {choco_bin}")
    os.environ["PATH"] = str(Path(choco_bin).parent) + os.pathsep + os.environ["PATH"]

    subprocess.run([choco_bin, "install", QEMU_CHOCOLATEY_NAME, "-y"], check=True)
    print("[+] QEMU installed via Chocolatey.")

def ensure_toolchain():
    if not detect_choco():
        install_chocolatey()
    else:
        print("[+] Chocolatey is already installed.")

    if not (shutil.which(QEMU_EXE) and shutil.which(QEMU_IMG)):
        install_qemu_choco()
    else:
        print("[+] QEMU tools already in PATH.")

# ----------------------------- Windows ISO Logic --------------------------- #

def fetch_latest_windows_iso():
    import requests

    print("[*] Fetching the latest Windows 11 ISO...")
    session = requests.Session()
    headers = {"User-Agent": "Mozilla/5.0"}

    resp = session.get(WINDOWS_ISO_DOWNLOAD_PAGE, headers=headers)
    if resp.status_code != 200:
        sys.exit("[!] Failed to reach Microsoft ISO page.")

    iso_url = None
    for line in resp.text.splitlines():
        if "software-download" in line and "iso" in line and "href" in line:
            match = re.search(r'href="(https://software-download\.microsoft\.com/[^"]+\.iso)"', line)
            if match:
                iso_url = match.group(1)
                break

    if not iso_url:
        sys.exit("[!] Unable to parse ISO link. Microsoft may have changed layout.")

    iso_name = Path(iso_url).name
    iso_path = Path.cwd() / iso_name
    print(f"[+] Downloading Windows 11 ISO: {iso_name}")

    with session.get(iso_url, headers=headers, stream=True) as r:
        r.raise_for_status()
        with open(iso_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)

    print(f"[+] ISO saved to {iso_path}")
    return str(iso_path)

# ----------------------------- VM Setup Logic ------------------------------ #

def create_disk_image(path: Path, size_gb: int):
    if path.exists():
        print(f"[+] Disk image already exists: {path}")
        return
    print(f"[+] Creating disk image {path} ({size_gb}G)...")
    subprocess.run([QEMU_IMG, "create", "-f", "qcow2", str(path), f"{size_gb}G"], check=True)

def build_qemu_command(args, disk_path):
    netdev = (
        ["-netdev", "user,id=net0", "-device", "e1000,netdev=net0"]
        if args.net == "user" else
        ["-netdev", "bridge,id=net0", "-device", "e1000,netdev=net0"]
    )
    return [
        QEMU_EXE,
        "-m", args.ram,
        "-smp", str(args.cpus),
        "-drive", f"file={disk_path},format=qcow2",
        "-cdrom", args.iso,
        "-boot", "order=d",
        "-cpu", "host,hv_relaxed,hv_vapic,hv_spinlocks=0x1fff",
        "-machine", "type=pc,accel=tcg",
        "-vga", "std",
        "-usb", "-device", "usb-tablet",
        *netdev
    ]

# ----------------------------- CLI & Execution ----------------------------- #

def parse_args():
    parser = argparse.ArgumentParser(description="Create and run a Windows 11 VM using QEMU.")
    parser.add_argument("--ram", type=str, required=True, help="RAM size (e.g. 8G or 4096M)")
    parser.add_argument("--disk-size", type=int, default=60, help="Disk size in GB (default: 60)")
    parser.add_argument("--cpus", type=int, default=4, help="CPU cores (default: 4)")
    parser.add_argument("--iso", type=str, help="Path to Windows 11 ISO (auto-downloads if not set)")
    parser.add_argument("--disk-image", type=str, help="Optional qcow2 path (default: win11_vm.qcow2)")
    parser.add_argument("--net", choices=["user", "bridge"], default="user", help="Networking mode (default: user)")
    return parser.parse_args()

def main():
    ensure_admin()
    ensure_requests()
    args = parse_args()
    ensure_toolchain()

    iso_path = args.iso or fetch_latest_windows_iso()
    args.iso = iso_path

    disk_path = Path(args.disk_image) if args.disk_image else DEFAULT_DISK_PATH
    create_disk_image(disk_path, args.disk_size)

    cmd = build_qemu_command(args, disk_path)
    print(f"[+] Launching Windows 11 VM:\n{' '.join(cmd)}")
    subprocess.run(cmd)

if __name__ == "__main__":
    main()
