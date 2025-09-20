"""
Voice-Controlled PowerPoint Generator with AI Integration

This module integrates voice recognition, command parsing, PowerPoint
generation, and Gemini AI for intelligent content creation.

Usage:
    voice_ppt = VoiceControlledPPT()
    voice_ppt.start_voice_control()
"""

from typing import Optional, Dict, Any
import os
import time
from ppt_generator import PPTGenerator
from voice_to_text import VoiceToText
from voice_command_parser import VoiceCommandParser, VoiceCommand
from gemini_ai import get_gemini_ai, AIGeneratedContent
from config import get_config, get_logger


class VoiceControlledPPT:
    """
    Main class that integrates voice recognition with PowerPoint generation and AI content creation.
    Provides a complete voice-controlled presentation creation system with intelligent content generation.
    """
    
    def __init__(self, 
                 voice_engine: str = "google",
                 enable_voice_feedback: bool = True,
                 auto_save: bool = True,
                 enable_ai: bool = True):
        """
        Initialize the voice-controlled PowerPoint system.
        
        Args:
            voice_engine (str): Speech recognition engine to use
            enable_voice_feedback (bool): Enable text-to-speech feedback
            auto_save (bool): Automatically save presentation after changes
            enable_ai (bool): Enable AI content generation features
        """
        # Initialize configuration and logging
        self.config = get_config()
        self.logger = get_logger()
        
        # Initialize components
        self.ppt_generator = PPTGenerator()
        self.voice_to_text = VoiceToText(
            engine=voice_engine,
            enable_voice_feedback=enable_voice_feedback
        )
        self.command_parser = VoiceCommandParser()
        
        # Initialize AI if enabled
        self.enable_ai = enable_ai
        if enable_ai:
            self.ai = get_gemini_ai()
            if not self.ai.is_available():
                self.logger.warning("Gemini AI not available - disabling AI features")
                self.enable_ai = False
        else:
            self.ai = None
        
        # Configuration
        self.auto_save = auto_save
        self.current_slide_index = 0
        self.presentation_name = self.config.default_presentation_name
        self.is_running = False
        
        # Statistics
        self.commands_processed = 0
        self.successful_commands = 0
        self.ai_generations = 0
        
        self.logger.info("Voice-Controlled PowerPoint system initialized")
        if enable_voice_feedback:
            self.voice_to_text.speak("Voice controlled PowerPoint system ready")
    
    def start_voice_control(self):
        """Start the voice control system with continuous listening."""
        if self.is_running:
            self.logger.warning("Voice control is already running")
            return
        
        self.logger.info("Starting Voice-Controlled PowerPoint system")
        self.logger.info("Available commands: Say 'help' to see all options")
        self.logger.info("Say 'stop listening' to quit")
        
        self.is_running = True
        
        # Add initial title slide
        self.ppt_generator.add_slide(
            layout_index=0, 
            title="AI-Powered Voice Presentation",
            subtitle=f"Created on {time.strftime('%Y-%m-%d %H:%M:%S')}"
        )
        
        if self.voice_to_text.enable_voice_feedback:
            self.voice_to_text.speak("Starting voice control. Say help for commands.")
        
        # Start continuous listening
        self.voice_to_text.start_continuous_listening(self._process_voice_command)
        
        try:
            # Keep the main thread alive
            while self.is_running:
                time.sleep(1)
        except KeyboardInterrupt:
            self.logger.info("Keyboard interrupt received")
        finally:
            self._shutdown()
    
    def process_single_command(self, timeout: int = 10) -> bool:
        """
        Process a single voice command with timeout.
        
        Args:
            timeout (int): Timeout in seconds for listening
            
        Returns:
            bool: True if command was processed successfully
        """
        self.logger.info("Listening for a single voice command...")
        text = self.voice_to_text.listen_once(timeout=timeout)
        
        if text:
            return self._process_voice_command(text)
        
        return False
    
    def _process_voice_command(self, text: str) -> bool:
        """
        Process a voice command text with AI enhancement.
        
        Args:
            text (str): Recognized voice command text
            
        Returns:
            bool: True if command was processed successfully
        """
        self.commands_processed += 1
        
        # Enhance command with AI if available
        if self.enable_ai and self.ai.is_available():
            enhanced_text = self.ai.enhance_voice_command(text)
            if enhanced_text and enhanced_text != text:
                self.logger.info(f"AI enhanced command: '{text}' → '{enhanced_text}'")
                text = enhanced_text
        
        # Parse the command
        command = self.command_parser.parse_command(text)
        
        if not command:
            self.logger.warning(f"Unknown command: '{text}'")
            if self.voice_to_text.enable_voice_feedback:
                self.voice_to_text.speak("I didn't understand that command. Say help for available commands.")
            return False
        
        self.logger.info(f"Processing command: {command.action}")
        
        # Validate command
        is_valid, error_msg = self.command_parser.validate_command(command)
        if not is_valid:
            self.logger.error(f"Command validation failed: {error_msg}")
            if self.voice_to_text.enable_voice_feedback:
                self.voice_to_text.speak(f"Command error: {error_msg}")
            return False
        
        # Execute command
        success = self._execute_command(command)
        
        if success:
            self.successful_commands += 1
            self.logger.info("Command executed successfully")
        else:
            self.logger.error("Command execution failed")
        
        return success
    
    def _execute_command(self, command: VoiceCommand) -> bool:
        """
        Execute a parsed voice command.
        
        Args:
            command (VoiceCommand): The command to execute
            
        Returns:
            bool: True if execution was successful
        """
        try:
            action = command.action
            params = command.parameters
            
            if action == "add_slide":
                title = params.get("title", "New Slide")
                layout = params.get("layout", 1)
                
                # Use AI to generate content if no specific title provided
                if self.enable_ai and title == "New Slide":
                    ai_content = self.ai.generate_slide_content("presentation slide", "content")
                    if ai_content:
                        title = ai_content.title
                        layout = ai_content.suggested_layout
                        self.ai_generations += 1
                        self.logger.info(f"AI generated slide title: {title}")
                
                slide_index = self.ppt_generator.add_slide(
                    layout_index=layout,
                    title=title
                )
                
                if slide_index >= 0:
                    self.current_slide_index = slide_index
                    if self.voice_to_text.enable_voice_feedback:
                        self.voice_to_text.speak(f"Added slide: {title}")
                    return True
                    
            elif action == "delete_slide":
                slide_number = params.get("slide_number", 0)
                success = self.ppt_generator.delete_slide(slide_number)
                
                if success and self.voice_to_text.enable_voice_feedback:
                    self.voice_to_text.speak(f"Deleted slide {slide_number + 1}")
                return success
                
            elif action == "update_text":
                updates = {}
                
                if "title" in params:
                    updates["title"] = params["title"]
                if "content" in params:
                    updates["content"] = params["content"]
                
                # Use AI to enhance content if available
                if self.enable_ai and self.ai.is_available():
                    if "title" in updates:
                        ai_content = self.ai.generate_slide_content(updates["title"], "content")
                        if ai_content and not updates.get("content"):
                            updates["content"] = ai_content.content
                            self.ai_generations += 1
                            self.logger.info("AI generated slide content")
                
                success = self.ppt_generator.update_text(
                    slide_index=self.current_slide_index,
                    text_updates=updates,
                    font_size=16,
                    bold=True
                )
                
                if success and self.voice_to_text.enable_voice_feedback:
                    self.voice_to_text.speak("Updated slide text")
                return success
                
            elif action == "insert_chart":
                chart_type = params.get("chart_type", "column")
                
                # Use AI to generate relevant chart data
                chart_data = None
                if self.enable_ai and self.ai.is_available():
                    chart_data = self.ai.suggest_chart_data("sample data visualization", chart_type)
                    if chart_data:
                        self.ai_generations += 1
                        self.logger.info("AI generated chart data")
                
                # Fallback to sample data if AI not available or failed
                if not chart_data:
                    chart_data = {
                        'categories': ['Q1', 'Q2', 'Q3', 'Q4'],
                        'series': [
                            {'name': 'Sales', 'values': [100, 120, 110, 140]},
                            {'name': 'Profit', 'values': [20, 25, 22, 30]}
                        ]
                    }
                
                success = self.ppt_generator.insert_chart(
                    slide_index=self.current_slide_index,
                    chart_type=chart_type,
                    data=chart_data
                )
                
                if success and self.voice_to_text.enable_voice_feedback:
                    self.voice_to_text.speak(f"Added {chart_type} chart")
                return success
                
            elif action == "insert_image":
                image_path = params.get("image_path", "")
                
                if not os.path.exists(image_path):
                    print(f"Image not found: {image_path}")
                    if self.voice_to_text.enable_voice_feedback:
                        self.voice_to_text.speak("Image file not found")
                    return False
                
                success = self.ppt_generator.insert_image(
                    slide_index=self.current_slide_index,
                    image_path=image_path
                )
                
                if success and self.voice_to_text.enable_voice_feedback:
                    self.voice_to_text.speak("Added image to slide")
                return success
                
            elif action == "change_background":
                color = params.get("color", (255, 255, 255))
                
                success = self.ppt_generator.change_background(
                    slide_index=self.current_slide_index,
                    background_type="solid",
                    color=color
                )
                
                if success and self.voice_to_text.enable_voice_feedback:
                    self.voice_to_text.speak("Changed background color")
                return success
                
            elif action == "modify_layout":
                layout_index = params.get("layout", 1)
                
                success = self.ppt_generator.modify_layout(
                    slide_index=self.current_slide_index,
                    new_layout_index=layout_index
                )
                
                if success and self.voice_to_text.enable_voice_feedback:
                    self.voice_to_text.speak("Changed slide layout")
                return success
                
            elif action == "save":
                filename = params.get("filename", self.presentation_name)
                
                if not filename.endswith('.pptx'):
                    filename += '.pptx'
                
                self.ppt_generator.save(filename)
                
                if self.voice_to_text.enable_voice_feedback:
                    self.voice_to_text.speak(f"Presentation saved as {filename}")
                return True
                
            elif action == "go_to_slide":
                slide_number = params.get("slide_number", 0)
                max_slides = self.ppt_generator.get_slide_count()
                
                if 0 <= slide_number < max_slides:
                    self.current_slide_index = slide_number
                    if self.voice_to_text.enable_voice_feedback:
                        self.voice_to_text.speak(f"Moved to slide {slide_number + 1}")
                    return True
                else:
                    if self.voice_to_text.enable_voice_feedback:
                        self.voice_to_text.speak("Invalid slide number")
                    return False
                    
            elif action == "help":
                help_text = self.command_parser.get_help_text()
                print(help_text)
                
                if self.voice_to_text.enable_voice_feedback:
                    self.voice_to_text.speak("Available commands printed to console")
                return True
                
            elif action == "stop_listening":
                self.logger.info("Stopping voice control...")
                if self.voice_to_text.enable_voice_feedback:
                    self.voice_to_text.speak("Stopping voice control")
                
                self.is_running = False
                return True
            
            # New AI-powered commands
            elif action == "generate_presentation":
                topic = params.get("topic", "presentation")
                slide_count = params.get("slide_count", 5)
                
                if self.enable_ai and self.ai.is_available():
                    outline = self.ai.generate_presentation_outline(topic, slide_count)
                    if outline:
                        self._create_ai_presentation(outline, topic)
                        self.ai_generations += 1
                        if self.voice_to_text.enable_voice_feedback:
                            self.voice_to_text.speak(f"Generated {len(outline)} slide presentation about {topic}")
                        return True
                
                if self.voice_to_text.enable_voice_feedback:
                    self.voice_to_text.speak("AI presentation generation not available")
                return False
            
            elif action == "ai_enhance_slide":
                if self.enable_ai and self.ai.is_available():
                    # Get current slide title to use as context
                    current_slides = self.ppt_generator.presentation.slides
                    if self.current_slide_index < len(current_slides):
                        current_slide = current_slides[self.current_slide_index]
                        title = current_slide.shapes.title.text if current_slide.shapes.title else "slide content"
                        
                        ai_content = self.ai.generate_slide_content(title, "content")
                        if ai_content:
                            updates = {
                                "title": ai_content.title,
                                "content": ai_content.content
                            }
                            
                            success = self.ppt_generator.update_text(
                                slide_index=self.current_slide_index,
                                text_updates=updates,
                                font_size=16,
                                bold=True
                            )
                            
                            if success:
                                self.ai_generations += 1
                                if self.voice_to_text.enable_voice_feedback:
                                    self.voice_to_text.speak("Enhanced slide with AI content")
                                return True
                
                if self.voice_to_text.enable_voice_feedback:
                    self.voice_to_text.speak("AI enhancement not available")
                return False
                
            else:
                self.logger.warning(f"Unknown action: {action}")
                return False
                
        except Exception as e:
            self.logger.error(f"Error executing command: {str(e)}")
            if self.voice_to_text.enable_voice_feedback:
                self.voice_to_text.speak("Command execution error")
            return False
    
    def _create_ai_presentation(self, outline: list, topic: str):
        """Create a full presentation from AI-generated outline."""
        try:
            # Clear existing slides (except first one)
            slides_count = self.ppt_generator.get_slide_count()
            for i in range(slides_count - 1, 0, -1):
                self.ppt_generator.delete_slide(i)
            
            # Update first slide as title slide
            title_content = self.ai.generate_slide_content(topic, "title") if self.ai else None
            if title_content:
                self.ppt_generator.update_text(0, {
                    "title": title_content.title,
                    "content": title_content.content
                })
            
            # Generate slides from outline
            for i, slide_title in enumerate(outline[1:], 1):  # Skip first title
                ai_content = self.ai.generate_slide_content(slide_title, "content")
                if ai_content:
                    slide_index = self.ppt_generator.add_slide(
                        layout_index=ai_content.suggested_layout,
                        title=ai_content.title
                    )
                    
                    if slide_index >= 0:
                        self.ppt_generator.update_text(slide_index, {
                            "content": ai_content.content
                        })
                        
            self.logger.info(f"Created AI presentation with {len(outline)} slides")
            
        except Exception as e:
            self.logger.error(f"Error creating AI presentation: {str(e)}")
    
    def _shutdown(self):
        """Shutdown the voice control system."""
        self.logger.info("Shutting down voice control system...")
        
        # Stop voice recognition
        self.voice_to_text.stop_continuous_listening()
        
        # Auto-save if enabled
        if self.auto_save:
            filename = f"{self.presentation_name}_{int(time.time())}.pptx"
            self.ppt_generator.save(filename)
            self.logger.info(f"Auto-saved presentation as: {filename}")
        
        # Print statistics
        self.logger.info("Session Statistics:")
        self.logger.info(f"  Commands processed: {self.commands_processed}")
        self.logger.info(f"  Successful commands: {self.successful_commands}")
        self.logger.info(f"  Success rate: {(self.successful_commands/max(1, self.commands_processed))*100:.1f}%")
        self.logger.info(f"  Total slides: {self.ppt_generator.get_slide_count()}")
        if self.enable_ai:
            self.logger.info(f"  AI generations: {self.ai_generations}")
        
        if self.voice_to_text.enable_voice_feedback:
            self.voice_to_text.speak("Voice control system stopped")
        
        self.logger.info("Voice control system shutdown complete")
    
    def test_system(self):
        """Test the voice control system components."""
        print("\n=== Voice Control System Test ===")
        
        # Test microphone
        print("1. Testing microphone...")
        mic_test = self.voice_to_text.test_microphone()
        
        if not mic_test:
            print("❌ Microphone test failed")
            return False
        
        # Test command parsing
        print("\n2. Testing command parsing...")
        test_commands = [
            "create new slide with title Test Slide",
            "add column chart",
            "change background to blue",
            "save presentation as test"
        ]
        
        for cmd_text in test_commands:
            command = self.command_parser.parse_command(cmd_text)
            if command:
                print(f"   ✅ '{cmd_text}' → {command.action}")
            else:
                print(f"   ❌ '{cmd_text}' → No match")
        
        print("\n✅ System test complete")
        return True
    
    def get_current_status(self) -> Dict[str, Any]:
        """Get current system status."""
        return {
            "is_running": self.is_running,
            "current_slide": self.current_slide_index + 1,
            "total_slides": self.ppt_generator.get_slide_count(),
            "commands_processed": self.commands_processed,
            "successful_commands": self.successful_commands,
            "presentation_name": self.presentation_name,
            "voice_feedback_enabled": self.voice_to_text.enable_voice_feedback,
            "ai_enabled": self.enable_ai,
            "ai_generations": getattr(self, 'ai_generations', 0)
        }