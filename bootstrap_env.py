#!/usr/bin/env python3
import os
import sys
import time
import subprocess
import shutil

def resource_path(relative_path):
    """
    Get absolute path to resource, works for dev and for PyInstaller bundle.
    """
    base_path = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base_path, relative_path)

def run(cmd):
    print(f"▶ {' '.join(cmd)}")
    subprocess.check_call(cmd)

def install_wheel_if_needed(venv_py, pip_cmd, wheel_path, pkg_name, installed):
    """
    Install wheel at wheel_path into venv if package pkg_name not already installed.
    installed is set of installed wheel filenames (strings).
    """
    if os.path.basename(wheel_path) in installed:
        return True
    try:
        # check if package already installed in venv
        subprocess.check_call([venv_py, '-m', 'pip', 'show', pkg_name],
                              stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        installed.add(os.path.basename(wheel_path))
        print(f"Skipping install; {pkg_name} already installed in venv")
        return True
    except Exception:
        pass

    try:
        run(pip_cmd + ['install', '--no-index', wheel_path])
        installed.add(os.path.basename(wheel_path))
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error installing {wheel_path}: {e}")
        return False

def find_wheels_for_prefix(modules_dir, prefix):
    """
    Return sorted list of wheel file paths in modules_dir that start with prefix (case-insensitive).
    """
    out = []
    pref = prefix.replace('_', '-').lower()
    for fname in os.listdir(modules_dir):
        if not fname.lower().endswith('.whl'):
            continue
        if fname.lower().startswith(pref):
            out.append(os.path.join(modules_dir, fname))
    return sorted(out)

def main():
    start_time = time.time()
    root = os.getcwd()
    env_dir = os.path.join(root, '.venv')

    # look for either "modules" or "modules13" in the bundle/working dir
    modules_dir = None
    for _name in ('modules', 'modules13'):
        candidate = resource_path(_name)
        if os.path.isdir(candidate):
            modules_dir = candidate
            break
    if modules_dir is None:
        # fallback to 'modules' path (will error later if not present)
        modules_dir = resource_path('modules')

    # 1) Locate host Python
    if os.name == 'nt':
        host_py = shutil.which('py') or shutil.which('python')
    else:
        host_py = shutil.which('python')

    if not host_py:
        print("❌ No Python interpreter found. Please install Python and ensure it's on your PATH.")
        sys.exit(1)

    # 2) Check if pip is already available on host Python
    try:
        subprocess.check_call([host_py, '-m', 'pip', '--version'],
                              stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        have_pip = True
    except Exception:
        have_pip = False

    # 3) Bootstrap pip offline if missing (use get-pip.py from modules_dir)
    if not have_pip:
        get_pip = os.path.join(modules_dir, 'get-pip.py')
        if os.path.isfile(get_pip):
            print("Bootstrapping pip on host Python using bundled get-pip.py...")
            try:
                run([host_py, get_pip])
            except Exception as e:
                print("Failed to bootstrap pip on host Python:", e)
                sys.exit(1)
        else:
            print("❌ pip not available and get-pip.py not found in modules/modules13.")
            print("Please provide get-pip.py in the modules folder or install pip on the host.")
            sys.exit(1)

    # 4) Create virtual environment if missing
    if not os.path.isdir(env_dir):
        run([host_py, '-m', 'venv', env_dir])
    else:
        print(f"Virtualenv already exists at {env_dir}")

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
        'python_dateutil', 'pytz', 'et_xmlfile', 'openpyxl', 'xlsxwriter',
        'numpy', 'pandas',
        # ensure these Windows/native packages are installed early
        'pywin32', 'psutil', 'xlwings'
    ]

    installed = set()

    # Pass 1: install in ordered sequence (if matching wheel exists)
    for prefix in ordered_prefixes:
        wheels = find_wheels_for_prefix(modules_dir, prefix)
        for wheel in wheels:
            fname = os.path.basename(wheel)
            pkg_name = fname.split('-')[0].lower()
            try:
                install_wheel_if_needed(venv_py, pip_cmd, wheel, pkg_name, installed)
            except Exception as e:
                print(f"Warning installing {fname}: {e}")

    # Pass 2: remaining runtime wheels
    for fname in sorted(os.listdir(modules_dir)):
        if not fname.lower().endswith('.whl'):
            continue
        if fname in installed:
            continue
        if any(fname.lower().startswith(pref) for pref in skip_prefixes):
            continue
        wheel_path = os.path.join(modules_dir, fname)
        pkg_name = fname.split('-')[0].lower()
        try:
            install_wheel_if_needed(venv_py, pip_cmd, wheel_path, pkg_name, installed)
        except Exception as e:
            print(f"Warning installing {fname}: {e}")

    # Post-install steps (Windows-specific)
    try:
        subprocess.check_call([venv_py, '-m', 'pywin32_postinstall', '-install'])
    except Exception:
        print('Warning: pywin32_postinstall failed or not present (non-fatal).')

    # copy xlwings add-in from site-packages into venv for easy access
    addin_candidates = [
        os.path.join(env_dir, 'Lib', 'site-packages', 'xlwings', 'xlwings.xlam'),
        os.path.join(env_dir, 'Lib', 'site-packages', 'xlwings.xlam'),
    ]
    copied_addin = False
    for addin_src in addin_candidates:
        if os.path.isfile(addin_src):
            dest_dir = os.path.join(env_dir, 'xlwings_addin')
            os.makedirs(dest_dir, exist_ok=True)
            shutil.copy2(addin_src, dest_dir)
            print('Copied xlwings add-in to', dest_dir)
            copied_addin = True
            break
    if not copied_addin:
        print('xlwings.xlam not found in venv site-packages (ok if not needed).')

    # 7) Success message
    end_time = time.time()
    elapsed = round(end_time - start_time, 0)
    if elapsed > 60:
        print(f"\n✅ Environment ready! Took {int(elapsed)} seconds.")
    else:
        print(f"\n✅ Environment ready! Took {int(elapsed)} seconds.")

    print("\nActivate virtualenv:")
    # show activation commands for common shells
    if os.name == 'nt':
        print("  Command Prompt:  .venv\\Scripts\\activate")
        print("  PowerShell:      .venv\\Scripts\\Activate.ps1")
        print("  Git Bash:        source .venv/Scripts/activate")
    else:
        print("  Bash/zsh:        source .venv/bin/activate")

if __name__ == '__main__':
    main()