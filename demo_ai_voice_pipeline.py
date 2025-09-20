"""
AI-Enhanced Voice-to-Text Pipeline Demo

This demo showcases the complete voice-controlled PowerPoint generation
system with Gemini AI integration for intelligent content creation.

Features:
- Voice recognition with speechrecognition
- Natural language command processing  
- AI-powered content generation with Gemini
- Intelligent slide creation and enhancement
- Professional logging system
- Environment-based configuration

Usage:
    python demo_ai_voice_pipeline.py

Requirements:
    1. Install packages: pip install -r requirements.txt
    2. Copy .env.template to .env 
    3. Add your Gemini API key to .env file
"""

import sys
import time
from pathlib import Path
from voice_controlled_ppt import VoiceControlledPPT
from gemini_ai import get_gemini_ai
from config import get_config, get_logger


def check_setup():
    """Check system setup and configuration."""
    logger = get_logger()
    config = get_config()
    
    logger.info("Checking system setup...")
    
    # Check .env file
    env_file = Path('.env')
    if not env_file.exists():
        logger.error(".env file not found!")
        logger.info("Please copy .env.template to .env and configure your settings")
        return False
    
    # Check Gemini API key
    if not config.gemini_api_key or config.gemini_api_key == "your_gemini_api_key_here":
        logger.error("Gemini API key not configured!")
        logger.info("Please set GEMINI_API_KEY in your .env file")
        logger.info("Get your API key from: https://makersuite.google.com/app/apikey")
        return False
    
    # Test Gemini AI
    logger.info("Testing Gemini AI connection...")
    ai = get_gemini_ai()
    if not ai.is_available():
        logger.error("Gemini AI not available")
        return False
    
    logger.info("System setup verified successfully")
    return True


def demo_ai_content_generation():
    """Demonstrate AI content generation capabilities."""
    logger = get_logger()
    logger.info("=== AI Content Generation Demo ===")
    
    ai = get_gemini_ai()
    
    # Test slide content generation
    logger.info("Generating AI slide content...")
    content = ai.generate_slide_content("Machine Learning Overview", "content")
    if content:
        logger.info(f"Generated Title: {content.title}")
        logger.info(f"Generated Content: {content.content[:100]}...")
        logger.info(f"Suggested Layout: {content.suggested_layout}")
    
    # Test presentation outline
    logger.info("Generating presentation outline...")
    outline = ai.generate_presentation_outline("Artificial Intelligence in Business", 5)
    if outline:
        logger.info("Generated Outline:")
        for i, title in enumerate(outline, 1):
            logger.info(f"  {i}. {title}")
    
    # Test chart data generation
    logger.info("Generating chart data...")
    chart_data = ai.suggest_chart_data("quarterly sales performance", "column")
    if chart_data:
        logger.info(f"Chart Categories: {chart_data['categories']}")
        for series in chart_data['series']:
            logger.info(f"Series: {series['name']} - Values: {series['values']}")
    
    return True


def demo_voice_command_enhancement():
    """Demonstrate AI voice command enhancement."""
    logger = get_logger()
    logger.info("=== Voice Command Enhancement Demo ===")
    
    ai = get_gemini_ai()
    
    test_commands = [
        "add slide about sales",
        "make chart",
        "change color blue",
        "insert picture",
        "create presentation about AI",
    ]
    
    logger.info("Testing voice command enhancement:")
    for cmd in test_commands:
        enhanced = ai.enhance_voice_command(cmd)
        if enhanced:
            logger.info(f"Original: '{cmd}'")
            logger.info(f"Enhanced: '{enhanced}'")
            logger.info("")
    
    return True


def demo_interactive_ai_commands():
    """Interactive demo with AI-powered voice commands."""
    logger = get_logger()
    logger.info("=== Interactive AI Voice Commands Demo ===")
    
    logger.info("This demo processes AI-enhanced voice commands")
    logger.info("Available AI commands:")
    logger.info("• 'create presentation about [topic]'")
    logger.info("• 'generate slide about [topic]'")
    logger.info("• 'enhance this slide with AI'")
    logger.info("• 'add smart chart about [topic]'")
    logger.info("• Regular voice commands (create slide, add chart, etc.)")
    
    voice_ppt = VoiceControlledPPT(enable_voice_feedback=True, enable_ai=True)
    
    for i in range(3):  # Process 3 commands
        logger.info(f"AI Command {i+1}/3 - Speak now (15 second timeout):")
        success = voice_ppt.process_single_command(timeout=15)
        
        if not success:
            logger.warning("No command recognized or command failed")
        
        # Show current status
        status = voice_ppt.get_current_status()
        logger.info(f"Status: Slide {status['current_slide']}/{status['total_slides']}, AI Generations: {status['ai_generations']}")
        
        time.sleep(1)
    
    # Save the presentation
    voice_ppt.ppt_generator.save("ai_interactive_demo.pptx")
    logger.info("💾 Saved AI-enhanced presentation as 'ai_interactive_demo.pptx'")
    
    return True


def demo_full_ai_voice_system():
    """Demonstrate the complete AI-enhanced voice system."""
    logger = get_logger()
    logger.info("=== Complete AI Voice System Demo ===")
    logger.info("Starting full AI-enhanced continuous voice control...")
    logger.info("AI Features Available:")
    logger.info("• Intelligent content generation")
    logger.info("• Voice command enhancement")
    logger.info("• Smart chart data creation")
    logger.info("• Automatic presentation generation")
    
    logger.info("Say 'stop listening' to end the demo")
    logger.info("Starting in 3 seconds...")
    time.sleep(3)
    
    voice_ppt = VoiceControlledPPT(
        enable_voice_feedback=True, 
        enable_ai=True,
        auto_save=True
    )
    
    # This will run until user says "stop listening"
    voice_ppt.start_voice_control()
    
    return True


def main():
    """Main demo function."""
    logger = get_logger()
    logger.info("AI-Enhanced Voice-to-Text Pipeline Demo")
    logger.info("=" * 60)
    
    # Check setup first
    if not check_setup():
        logger.error("Setup check failed. Please fix the issues above.")
        return
    
    logger.info("Choose a demo to run:")
    logger.info("1. AI Content Generation Test")
    logger.info("2. Voice Command Enhancement Demo") 
    logger.info("3. Interactive AI Voice Commands (3 commands)")
    logger.info("4. Complete AI Voice System (continuous)")
    logger.info("5. Run All Demos")
    logger.info("0. Exit")
    
    while True:
        try:
            choice = input("\nEnter your choice (0-5): ").strip()
            
            if choice == "0":
                logger.info("👋 Goodbye!")
                break
                
            elif choice == "1":
                demo_ai_content_generation()
                
            elif choice == "2":
                demo_voice_command_enhancement()
                
            elif choice == "3":
                demo_interactive_ai_commands()
                
            elif choice == "4":
                demo_full_ai_voice_system()
                
            elif choice == "5":
                logger.info("🔄 Running all demos...")
                demo_ai_content_generation()
                demo_voice_command_enhancement()
                
                run_interactive = input("\nRun interactive AI demo? (y/n): ").lower().startswith('y')
                if run_interactive:
                    demo_interactive_ai_commands()
                
                run_full = input("\nRun complete AI voice system? (y/n): ").lower().startswith('y')
                if run_full:
                    demo_full_ai_voice_system()
                    
            else:
                logger.warning("Invalid choice. Please enter 0-5.")
                continue
                
            input("\nPress Enter to continue...")
            
        except KeyboardInterrupt:
            logger.info("Demo interrupted by user")
            break
        except Exception as e:
            logger.error(f"Demo error: {str(e)}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n👋 Demo terminated by user")
    except Exception as e:
        print(f"Unexpected error: {str(e)}")
        sys.exit(1)