"""
Voice-to-Text Pipeline Demo for PowerPoint Generation

This demo script showcases the complete voice-to-text pipeline integrated
with PowerPoint generation capabilities.

Features demonstrated:
- Voice recognition using speechrecognition library
- Voice command parsing and interpretation
- PowerPoint operations through voice commands
- Text-to-speech feedback
- Continuous listening mode

Usage:
    python demo_voice_pipeline.py

Requirements:
    pip install speechrecognition pyaudio pyttsx3 python-pptx
"""

import sys
import time
from voice_controlled_ppt import VoiceControlledPPT
from voice_to_text import VoiceToText
from voice_command_parser import VoiceCommandParser


def check_dependencies():
    """Check if all required packages are installed."""
    required_packages = [
        ('speech_recognition', 'speechrecognition'),
        ('pyttsx3', 'pyttsx3'),
        ('pptx', 'python-pptx'),
    ]
    
    missing_packages = []
    
    for module_name, package_name in required_packages:
        try:
            __import__(module_name)
            print(f"✅ {package_name} is installed")
        except ImportError:
            print(f"❌ {package_name} is missing")
            missing_packages.append(package_name)
    
    if missing_packages:
        print(f"\nPlease install missing packages:")
        print(f"pip install {' '.join(missing_packages)}")
        return False
    
    return True


def demo_basic_voice_recognition():
    """Demonstrate basic voice recognition capabilities."""
    print("\n=== Basic Voice Recognition Demo ===")
    
    voice_system = VoiceToText(enable_voice_feedback=True)
    
    print("Testing microphone...")
    if not voice_system.test_microphone():
        print("❌ Microphone test failed. Please check your microphone setup.")
        return False
    
    print("\n🎤 Say something for 5 seconds:")
    result = voice_system.listen_once(timeout=5)
    
    if result:
        print(f"✅ Recognition successful: '{result}'")
        voice_system.speak(f"You said: {result}")
    else:
        print("❌ No speech recognized")
    
    return True


def demo_command_parsing():
    """Demonstrate voice command parsing."""
    print("\n=== Voice Command Parsing Demo ===")
    
    parser = VoiceCommandParser()
    
    test_commands = [
        "create new slide with title My Presentation",
        "add column chart",
        "change background to blue",
        "insert image from sample.png",
        "update title to New Title",
        "delete slide number 2",
        "save presentation as my_demo",
        "help",
        "stop listening"
    ]
    
    print("Testing command parsing:")
    for cmd_text in test_commands:
        command = parser.parse_command(cmd_text)
        if command:
            print(f"  ✅ '{cmd_text}'")
            print(f"     → Action: {command.action}")
            print(f"     → Parameters: {command.parameters}")
        else:
            print(f"  ❌ '{cmd_text}' → No match found")
        print()
    
    return True


def demo_interactive_voice_commands():
    """Demonstrate interactive voice command processing."""
    print("\n=== Interactive Voice Commands Demo ===")
    print("This demo will process individual voice commands.")
    print("You can test commands like:")
    print("• 'create new slide'")
    print("• 'add column chart'") 
    print("• 'change background to blue'")
    print("• 'help' to see all commands")
    print("• 'save presentation' to save")
    print()
    
    voice_ppt = VoiceControlledPPT(enable_voice_feedback=True)
    
    for i in range(5):  # Process up to 5 commands
        print(f"\nCommand {i+1}/5 - Speak now (10 second timeout):")
        success = voice_ppt.process_single_command(timeout=10)
        
        if not success:
            print("No command recognized or command failed")
        
        # Show current status
        status = voice_ppt.get_current_status()
        print(f"Status: Slide {status['current_slide']}/{status['total_slides']}")
        
        time.sleep(1)
    
    # Save the presentation
    voice_ppt.ppt_generator.save("interactive_demo.pptx")
    print("\n💾 Saved presentation as 'interactive_demo.pptx'")
    
    return True


def demo_continuous_voice_control():
    """Demonstrate continuous voice control mode."""
    print("\n=== Continuous Voice Control Demo ===")
    print("This will start continuous listening mode.")
    print("You can speak multiple commands in sequence.")
    print("Say 'stop listening' to end the demo.")
    print()
    
    print("Starting in 3 seconds...")
    time.sleep(3)
    
    voice_ppt = VoiceControlledPPT(enable_voice_feedback=True)
    
    # This will run until user says "stop listening"
    voice_ppt.start_voice_control()
    
    return True


def main():
    """Main demo function."""
    print("🎤 Voice-to-Text Pipeline Demo for PowerPoint Generation")
    print("=" * 60)
    
    # Check dependencies
    if not check_dependencies():
        print("\n❌ Missing required dependencies. Please install them first.")
        return
    
    print("\nChoose a demo to run:")
    print("1. Basic Voice Recognition Test")
    print("2. Command Parsing Demo")
    print("3. Interactive Voice Commands (5 commands)")
    print("4. Continuous Voice Control (full system)")
    print("5. Run All Demos")
    print("0. Exit")
    
    while True:
        try:
            choice = input("\nEnter your choice (0-5): ").strip()
            
            if choice == "0":
                print("👋 Goodbye!")
                break
                
            elif choice == "1":
                demo_basic_voice_recognition()
                
            elif choice == "2":
                demo_command_parsing()
                
            elif choice == "3":
                demo_interactive_voice_commands()
                
            elif choice == "4":
                demo_continuous_voice_control()
                
            elif choice == "5":
                print("\n🔄 Running all demos...")
                demo_basic_voice_recognition()
                demo_command_parsing()
                
                run_interactive = input("\nRun interactive demo? (y/n): ").lower().startswith('y')
                if run_interactive:
                    demo_interactive_voice_commands()
                
                run_continuous = input("\nRun continuous demo? (y/n): ").lower().startswith('y')
                if run_continuous:
                    demo_continuous_voice_control()
                    
            else:
                print("Invalid choice. Please enter 0-5.")
                continue
                
            input("\nPress Enter to continue...")
            
        except KeyboardInterrupt:
            print("\n\n⌨️ Demo interrupted by user")
            break
        except Exception as e:
            print(f"\n❌ Error: {str(e)}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n👋 Demo terminated by user")
    except Exception as e:
        print(f"\n❌ Unexpected error: {str(e)}")
        sys.exit(1)