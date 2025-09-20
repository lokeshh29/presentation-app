#!/usr/bin/env python3
"""
Simple installer script for AI-PPT Agent dependencies.
This script will check for and install missing dependencies.
"""

import subprocess
import sys
import importlib

def install_package(package):
    """Install a package using pip."""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        print(f"✅ Installed {package}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to install {package}: {e}")
        return False

def check_and_install_dependencies():
    """Check for required dependencies and install if missing."""
    
    # Core FastAPI dependencies (required for API server)
    api_dependencies = [
        ("fastapi", "fastapi"),
        ("uvicorn", "uvicorn[standard]"),
        ("python-multipart", "python-multipart"),
    ]
    
    # Optional system monitoring
    optional_dependencies = [
        ("psutil", "psutil"),
    ]
    
    print("🔍 Checking API dependencies...")
    
    missing_core = []
    for module, package in api_dependencies:
        try:
            importlib.import_module(module)
            print(f"✅ {module} - available")
        except ImportError:
            print(f"❌ {module} - missing")
            missing_core.append(package)
    
    if missing_core:
        print(f"\n📦 Installing {len(missing_core)} missing core dependencies...")
        for package in missing_core:
            install_package(package)
    
    print(f"\n🔍 Checking optional dependencies...")
    for module, package in optional_dependencies:
        try:
            importlib.import_module(module)
            print(f"✅ {module} - available")
        except ImportError:
            print(f"⚠️  {module} - missing (optional)")
            choice = input(f"Install {package}? (y/N): ").lower().strip()
            if choice == 'y':
                install_package(package)
    
    print(f"\n✅ Dependency check completed!")
    print(f"🚀 You can now run: python app.py")

if __name__ == "__main__":
    print("🔧 AI-PPT Agent Dependency Installer")
    print("=" * 40)
    
    try:
        check_and_install_dependencies()
    except KeyboardInterrupt:
        print(f"\n⚠️  Installation cancelled by user")
    except Exception as e:
        print(f"\n❌ Installation failed: {e}")
        sys.exit(1)