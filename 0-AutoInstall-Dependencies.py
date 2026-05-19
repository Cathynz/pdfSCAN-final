import sys
import subprocess
import importlib.util
import traceback
import tkinter as tk
from tkinter import messagebox

# ============================================================
# Auto Installer for McConnell Dowell Document Management data extraction tools :-)
# This script:
# 1. Checks Python version
# 2. Checks whether tkinter exists
# 3. Installs only missing external packages
# 4. Skips packages already installed
# 5. Shows clear success/failure messages
# ============================================================

MIN_PYTHON = (3, 9)

# Required packages for all 4 scripts
# (module_name, pip_name)
REQUIRED_PACKAGES = [
    ("pandas", "pandas"),           # Excel file handling
    ("openpyxl", "openpyxl"),       # .xlsx file support
    ("pdfplumber", "pdfplumber"),   # PDF text extraction
    ("xlrd", "xlrd"),               # .xls file support (legacy)
    ("pyxlsb", "pyxlsb"),           # .xlsb file support
]

# Standard library modules (check only, don't install)
CHECK_ONLY_MODULES = [
    "tkinter",  # GUI framework (built-in to Python)
]


def module_exists(module_name: str) -> bool:
    """Return True if module can be found."""
    return importlib.util.find_spec(module_name) is not None


def ensure_pip() -> bool:
    """Return True if pip is available."""
    try:
        result = subprocess.run(
            [sys.executable, "-m", "pip", "--version"],
            capture_output=True,
            text=True,
            timeout=10
        )
        return result.returncode == 0
    except Exception:
        return False


def install_package(pip_name: str) -> bool:
    """Install a package using the current Python interpreter."""
    try:
        print(f"Installing {pip_name}...")
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", pip_name, "--upgrade"],
            capture_output=True,
            text=True,
            timeout=300  # 5 minutes max per package
        )
        success = result.returncode == 0
        if success:
            print(f"✓ {pip_name} installed successfully")
        else:
            print(f"✗ {pip_name} failed to install")
            print(f"Error: {result.stderr}")
        return success
    except Exception as e:
        print(f"✗ {pip_name} failed: {str(e)}")
        return False


def build_report(title: str, lines: list) -> str:
    """Build formatted report text."""
    separator = "=" * 60
    return f"{separator}\n{title}\n{separator}\n\n" + "\n".join(lines)


def main():
    """Main installer logic."""
    # Create hidden tkinter window for message boxes
    root = tk.Tk()
    root.withdraw()
    
    print("\n" + "=" * 60)
    print("  TW Document Management System - Dependency Installer")
    print("=" * 60 + "\n")
    
    # System info
    info_lines = []
    info_lines.append(f"Python executable: {sys.executable}")
    info_lines.append(f"Python version: {sys.version.split()[0]}")
    
    print(f"Python version: {sys.version.split()[0]}")
    print(f"Python path: {sys.executable}\n")
    
    # Check Python version
    if sys.version_info < MIN_PYTHON:
        warning_msg = (
            f"⚠️ Your Python version is {sys.version.split()[0]}\n\n"
            f"Recommended version: Python {MIN_PYTHON[0]}.{MIN_PYTHON[1]} or newer\n\n"
            "The scripts may still work, but upgrading Python is recommended.\n\n"
            "Continue anyway?"
        )
        if not messagebox.askyesno("Python Version Warning", warning_msg):
            print("Installation cancelled by user.")
            return
    
    # Check pip availability
    print("Checking pip availability...")
    if not ensure_pip():
        error_msg = (
            "❌ pip is not available in this Python installation.\n\n"
            "pip is required to install packages.\n\n"
            "Please reinstall Python and ensure 'pip' is included during installation."
        )
        messagebox.showerror("pip Missing", error_msg)
        print("ERROR: pip not found!")
        return
    
    print("✓ pip is available\n")
    
    # Check standard library modules (like tkinter)
    print("Checking built-in Python components...")
    missing_std = []
    for module_name in CHECK_ONLY_MODULES:
        if not module_exists(module_name):
            missing_std.append(module_name)
            print(f"✗ {module_name} - MISSING")
        else:
            print(f"✓ {module_name} - found")
    
    if missing_std:
        error_msg = (
            "❌ The following built-in Python component is missing:\n\n"
            + "\n".join(f"  • {name}" for name in missing_std) +
            "\n\ntkinter is required for the GUI tools.\n\n"
            "Please reinstall Python and ensure 'tcl/tk and IDLE' is checked during installation."
        )
        messagebox.showerror("Missing Python Components", error_msg)
        print(f"\nERROR: Missing components: {', '.join(missing_std)}")
        return
    
    print()
    
    # Check required external packages
    print("Checking required packages...")
    already_installed = []
    missing_packages = []
    
    for module_name, pip_name in REQUIRED_PACKAGES:
        if module_exists(module_name):
            already_installed.append(module_name)
            print(f"✓ {module_name} - already installed")
        else:
            missing_packages.append((module_name, pip_name))
            print(f"✗ {module_name} - MISSING")
    
    print()
    
    # If everything is already installed
    if not missing_packages:
        success_msg = build_report(
            "✅ All Ready!",
            info_lines + 
            ["", "All required packages are already installed:", ""] +
            [f"  ✓ {name}" for name in already_installed] +
            ["", "You can now run the TW Document Management tools!"]
        )
        messagebox.showinfo("Installation Check Complete", success_msg)
        print("✅ All packages already installed!")
        print("\nYou can now run:")
        print("  • 1_PDF_List_Generator_GUI.py")
        print("  • 2_PDF_Batch_Scanner_UPDATED.py")
        print("  • 3_SharePoint_URL_Generator_GUI.py")
        return
    
    # Ask user to proceed with installation
    install_msg_lines = []
    install_msg_lines += info_lines
    install_msg_lines += [""]
    
    if already_installed:
        install_msg_lines.append("Already installed:")
        install_msg_lines += [f"  ✓ {name}" for name in already_installed]
        install_msg_lines.append("")
    
    install_msg_lines.append("The following packages need to be installed:")
    install_msg_lines += [f"  • {module_name}" for module_name, _ in missing_packages]
    install_msg_lines.append("")
    install_msg_lines.append("This may take 1-3 minutes depending on your internet speed.")
    install_msg_lines.append("")
    install_msg_lines.append("Click OK to install now.")
    
    proceed = messagebox.askokcancel(
        "Install Missing Packages",
        build_report("📦 Package Installation Required", install_msg_lines)
    )
    
    if not proceed:
        print("Installation cancelled by user.")
        return
    
    # Install missing packages
    print("=" * 60)
    print("Installing missing packages...")
    print("=" * 60 + "\n")
    
    installed_now = []
    failed = []
    
    for module_name, pip_name in missing_packages:
        print(f"Installing {module_name} ({pip_name})...")
        
        ok = install_package(pip_name)
        
        # Verify installation
        if ok and module_exists(module_name):
            installed_now.append(module_name)
            print(f"✓ {module_name} installed and verified\n")
        else:
            failed.append((module_name, pip_name))
            print(f"✗ {module_name} installation failed\n")
    
    print("=" * 60)
    print("Installation Complete")
    print("=" * 60 + "\n")
    
    # Build summary report
    summary_lines = []
    summary_lines += info_lines
    summary_lines += [""]
    
    if already_installed:
        summary_lines.append("Already installed:")
        summary_lines += [f"  ✓ {name}" for name in already_installed]
        summary_lines.append("")
    
    if installed_now:
        summary_lines.append("Installed successfully:")
        summary_lines += [f"  ✓ {name}" for name in installed_now]
        summary_lines.append("")
    
    if failed:
        summary_lines.append("❌ Failed to install:")
        summary_lines += [f"  ✗ {module_name} (pip install {pip_name})" 
                         for module_name, pip_name in failed]
        summary_lines.append("")
        summary_lines.append("To install failed packages manually, run:")
        for _, pip_name in failed:
            summary_lines.append(f"  {sys.executable} -m pip install {pip_name}")
        summary_lines.append("")
        summary_lines.append("Or copy and paste this command:")
        failed_packages = " ".join([pip_name for _, pip_name in failed])
        summary_lines.append(f"  {sys.executable} -m pip install {failed_packages}")
    
    # Show final message
    if failed:
        messagebox.showwarning(
            "Installation Finished with Problems",
            build_report("⚠️ Some packages could not be installed", summary_lines)
        )
        print("⚠️ Installation completed with errors")
        print("\nFailed packages:", ", ".join([m for m, _ in failed]))
    else:
        summary_lines.append("✅ All packages are now installed!")
        summary_lines.append("")
        summary_lines.append("You can now run the McConnell Dowell Document Management data extraction tools:")
        summary_lines.append("  • 1_PDF_List_Generator_GUI.py")
        summary_lines.append("  • 2_PDF_Batch_Scanner_UPDATED.py")
        summary_lines.append("  • 3_SharePoint_URL_Generator_GUI.py")
        
        messagebox.showinfo(
            "Installation Complete",
            build_report("✅ All packages installed successfully!", summary_lines)
        )
        print("✅ All packages installed successfully!")
        print("\nYou can now run:")
        print("  • 1_PDF_List_Generator_GUI.py")
        print("  • 2_PDF_Batch_Scanner_UPDATED.py")
        print("  • 3_SharePoint_URL_Generator_GUI.py")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nInstallation cancelled by user (Ctrl+C)")
    except Exception as e:
        error_text = f"Unexpected error occurred:\n\n{str(e)}\n\n{traceback.format_exc()}"
        print("\n" + "=" * 60)
        print("ERROR")
        print("=" * 60)
        print(error_text)
        
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Unexpected Error", error_text)
        except:
            pass