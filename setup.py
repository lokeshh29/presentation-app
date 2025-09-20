"""
Setup script for Voice-to-Text PowerPoint Generation Pipeline

This script helps install required packages and test the system setup.
"""

import subprocess
import sys
import os


def install_package(package):
    """Install a package using pip."""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        return True
    except subprocess.CalledProcessError:
        return False


def main():
    """Main setup function."""
    print("🔧 Voice-to-Text PowerPoint Generation Setup")
    print("=" * 50)
    
    # Required packages
    packages = [
        "python-pptx",
        "speechrecognition", 
        "pyttsx3",
        "pyaudio",
        "Pillow"
    ]
    
    print("Installing required packages...")
    
    failed_packages = []
    
    for package in packages:
        print(f"Installing {package}...")
        if install_package(package):
            print(f"✅ {package} installed successfully")
        else:
            print(f"❌ Failed to install {package}")
            failed_packages.append(package)
    
    if failed_packages:
        print(f"\n⚠️ Failed to install: {', '.join(failed_packages)}")
        print("\nNote: PyAudio might require system-level dependencies:")
        print("Windows: pip install pyaudio should work")
        print("macOS: brew install portaudio")
        print("Linux: sudo apt-get install python3-pyaudio")
    else:
        print("\n✅ All packages installed successfully!")
    
    print("\n🎤 Testing voice recognition...")
    
    try:
        import speech_recognition as sr
        import pyttsx3
        
        # Test basic imports
        recognizer = sr.Recognizer()
        microphone = sr.Microphone()
        tts = pyttsx3.init()
        
        print("✅ Voice recognition components loaded successfully")
        print("\n🚀 Setup complete! You can now run:")
        print("   python demo_voice_pipeline.py")
        
    except ImportError as e:
        print(f"❌ Import error: {e}")
        print("Some packages may not have installed correctly.")
    
    except Exception as e:
        print(f"❌ Setup error: {e}")


if __name__ == "__main__":
    main()