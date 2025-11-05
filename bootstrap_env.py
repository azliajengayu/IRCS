#!/usr/bin/env python3
import sys
import os
import subprocess
import shutil
import time


def resource_path(relative_path):
    """
    Get absolute path to resource, works for dev and for PyInstaller bundle.
    """
    base_path = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base_path, relative_path)


def run(cmd):
    print(f"▶ {' '.join(cmd)}")
    subprocess.check_call(cmd)


def main():
    start_time = time.time()
    root = os.getcwd()
    env_dir = os.path.join(root, '.venv')
    modules_dir = resource_path('modules')

    # 1) Locate host Python
    if os.name == 'nt':
        host_py = shutil.which('py') or shutil.which('python')
    else:
        host_py = shutil.which('python')

    if not host_py:
        print("❌ No Python interpreter found. Please install Python and ensure it's on your PATH.")
        sys.exit(1)

    # 2) Check if pip is already available
    try:
        subprocess.check_call([host_py, '-m', 'pip', '--version'],
                              stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        have_pip = True
    except Exception:
        have_pip = False

    # 3) Bootstrap pip offline if missing
    if not have_pip:
        get_pip = os.path.join(modules_dir, 'get-pip.py')
        if os.path.isfile(get_pip):
            run([host_py, get_pip, '--no-index', '--find-links', modules_dir])
        else:
            print("❌ get-pip.py not found; cannot bootstrap pip.")
            sys.exit(1)

    # 4) Create virtual environment if missing
    if not os.path.isdir(env_dir):
        run([host_py, '-m', 'venv', env_dir])

    # 5) Determine venv's python & pip
    if os.name == 'nt':
        venv_py = os.path.join(env_dir, 'Scripts', 'python.exe')
    else:
        venv_py = os.path.join(env_dir, 'bin', 'python')
    pip_cmd = [venv_py, '-m', 'pip']

    # 6) Install wheels in safe order (skip already-installed)
    skip_prefixes = [
        'pyinstaller', 'altgraph', 'pefile',
        'packaging', 'pyinstaller_hooks_contrib', 'pywin32_ctypes'
    ]
    ordered_prefixes = [
        'wheel', 'setuptools', 'tzdata', 'six',
        'python_dateutil', 'pytz', 'et_xmlfile', 'openpyxl', 'xlsxwriter', 'numpy', 'pandas'
    ]

    installed = set()
    # Pass 1: core and prerequisites
    for prefix in ordered_prefixes:
        pkg_name = prefix.replace('_', '-')
        try:
            subprocess.check_call([venv_py, '-m', 'pip', 'show', pkg_name],
                                  stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            # already installed
            continue
        except Exception:
            pass
        # install matching wheel files
        for fname in sorted(os.listdir(modules_dir)):
            lower = fname.lower()
            if not lower.endswith('.whl') or not lower.startswith(prefix):
                continue
            path = os.path.join(modules_dir, fname)
            run(pip_cmd + ['install', '--no-index', path])
            installed.add(fname)

    # Pass 2: remaining runtime wheels
    for fname in sorted(os.listdir(modules_dir)):
        lower = fname.lower()
        if fname in installed or not lower.endswith('.whl'):
            continue
        if any(lower.startswith(pref) for pref in skip_prefixes):
            continue
        # derive package from wheel filename (before first dash)
        pkg_name = lower.split('-')[0]
        try:
            subprocess.check_call([venv_py, '-m', 'pip', 'show', pkg_name],
                                  stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            continue
        except Exception:
            pass
        path = os.path.join(modules_dir, fname)
        run(pip_cmd + ['install', '--no-index', path])

    # 7) Success message
    end_time = time.time()
    if round((end_time - start_time),0) > 60:
        print(f"RUNTIME: {round((end_time - start_time), 2) / 60} minutes")
    else:
        print(f"RUNTIME: {round((end_time - start_time), 2)} second")
    print("\n✅ Environment ready! Activate with:")
    if os.name == 'nt':
        print("   source .venv/Scripts/activate   (Git Bash)")
    else:
        print("   source .venv/bin/activate      (Linux/Mac)")


if __name__ == '__main__':
    main()
