"""
Unified builder for the harmonized Vault Toolkit.

Run this on Windows (PyInstaller produces a binary for the OS it runs on):

    python build_vault.py

It builds from Vault_Toolkit.spec, which is configured for a FAST-launch
onedir/no-UPX build. Output lands in dist/Vault_Toolkit/Vault_Toolkit.exe
(ship the whole dist/Vault_Toolkit folder, or zip it).
"""
import os
import shutil
import subprocess
import sys

SPEC_FILE = "Vault_Toolkit.spec"


def clean_registry():
    """Cleans this toolkit's own config registry keys (Windows only)."""
    if not sys.platform.startswith("win"):
        print("Not on Windows - skipping registry cleanup.")
        return
    import winreg
    print("Cleaning associated registry keys...")
    for hive_path in (r"Software\OfficeVaultToolkit", r"Software\TextVaultToolkit"):
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, hive_path)
            print(f"Removed HKCU\\{hive_path}")
        except FileNotFoundError:
            pass
        except Exception as e:
            print(f"Error cleaning {hive_path}: {e}")


def clean_build_dirs():
    """Removes previous PyInstaller build artifacts."""
    print("Cleaning old build directories...")
    for dir_name in ['build', 'dist', '__pycache__']:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"Removed {dir_name}/")


def build_executable():
    """Builds the executable from the optimized spec file."""
    print("Installing/Upgrading PyInstaller...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pyinstaller"])

    print("Building executable from spec (onedir, no UPX)...")
    subprocess.check_call([sys.executable, "-m", "PyInstaller", "--noconfirm", SPEC_FILE])
    print("\nBuild complete.")
    print("Executable: dist/Vault_Toolkit/Vault_Toolkit.exe")
    print("Ship the entire dist/Vault_Toolkit/ folder (zip it for distribution).")


if __name__ == "__main__":
    clean_registry()
    clean_build_dirs()
    build_executable()
