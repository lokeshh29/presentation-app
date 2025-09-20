"""
Voice to Text Pipeline using SpeechRecognition

This module provides voice recognition capabilities for controlling PowerPoint
generation through voice commands. It uses the speechrecognition library
with various speech recognition engines.

Required packages:
- speechrecognition
- pyaudio (for microphone access)
- pyttsx3 (for text-to-speech feedback)

Install with: pip install speechrecognition pyaudio pyttsx3
"""

import speech_recognition as sr
import pyttsx3
import time
import threading
from typing import Optional, Callable, Dict, List
import re
import queue
from config import get_config, get_logger


class VoiceToText:
    """
    Voice to Text pipeline for converting speech to text commands.
    Supports multiple speech recognition engines and provides voice feedback.
    """
    
    def __init__(self, 
                 engine: str = "google", 
                 language: str = "en-US",
                 enable_voice_feedback: bool = True):
        """
        Initialize the Voice to Text system.
        
        Args:
            engine (str): Speech recognition engine ('google', 'sphinx', 'bing', 'azure')
            language (str): Language for speech recognition (default: 'en-US')
            enable_voice_feedback (bool): Enable text-to-speech feedback
        """
        self.config = get_config()
        self.logger = get_logger()
        
        self.recognizer = sr.Recognizer()
        self.microphone = sr.Microphone()
        self.engine = engine.lower()
        self.language = language
        self.is_listening = False
        self.command_queue = queue.Queue()
        self.listening_thread = None
        
        # Voice feedback setup
        self.enable_voice_feedback = enable_voice_feedback
        if enable_voice_feedback:
            try:
                self.tts_engine = pyttsx3.init()
                self.tts_engine.setProperty('rate', 150)  # Speaking rate
                voices = self.tts_engine.getProperty('voices')
                if voices:
                    self.tts_engine.setProperty('voice', voices[0].id)
            except Exception as e:
                self.logger.warning(f"Could not initialize text-to-speech: {e}")
                self.enable_voice_feedback = False
        
        # Calibrate microphone for ambient noise
        self._calibrate_microphone()
        
        self.logger.info(f"Voice recognition initialized with {engine} engine")
        if self.enable_voice_feedback:
            self.speak("Voice recognition system ready")
    
    def _calibrate_microphone(self):
        """Calibrate microphone for ambient noise."""
        self.logger.info("Calibrating microphone for ambient noise...")
        try:
            with self.microphone as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=1)
            self.logger.info("Microphone calibrated successfully")
        except Exception as e:
            self.logger.warning(f"Could not calibrate microphone: {e}")
    
    def speak(self, text: str):
        """
        Convert text to speech for voice feedback.
        
        Args:
            text (str): Text to speak
        """
        if not self.enable_voice_feedback:
            return
            
        try:
            self.logger.debug(f"Speaking: {text}")
            self.tts_engine.say(text)
            self.tts_engine.runAndWait()
        except Exception as e:
            self.logger.error(f"Error in text-to-speech: {e}")
    
    def listen_once(self, timeout: int = 5) -> Optional[str]:
        """
        Listen for a single voice command.
        
        Args:
            timeout (int): Timeout in seconds for listening
            
        Returns:
            str: Recognized text or None if no speech detected
        """
        try:
            self.logger.info("Listening for voice input...")
            if self.enable_voice_feedback:
                self.speak("Listening")
            
            with self.microphone as source:
                # Listen for audio with timeout
                audio = self.recognizer.listen(source, timeout=timeout, phrase_time_limit=10)
                
            self.logger.debug("Processing speech recognition...")
            
            # Recognize speech using selected engine
            text = self._recognize_speech(audio)
            
            if text:
                self.logger.info(f"Speech recognized: '{text}'")
                if self.enable_voice_feedback:
                    self.speak(f"I heard: {text}")
                return text.lower()
            else:
                self.logger.warning("No speech recognized")
                return None
                
        except sr.WaitTimeoutError:
            self.logger.info("Listening timeout - no speech detected")
            return None
        except sr.UnknownValueError:
            self.logger.warning("Could not understand speech")
            if self.enable_voice_feedback:
                self.speak("Sorry, I couldn't understand that")
            return None
        except sr.RequestError as e:
            self.logger.error(f"Recognition service error: {e}")
            return None
        except Exception as e:
            self.logger.error(f"Unexpected error in speech recognition: {e}")
            return None
    
    def _recognize_speech(self, audio) -> Optional[str]:
        """
        Recognize speech using the selected engine.
        
        Args:
            audio: Audio data from microphone
            
        Returns:
            str: Recognized text or None
        """
        try:
            if self.engine == "google":
                return self.recognizer.recognize_google(audio, language=self.language)
            elif self.engine == "sphinx":
                return self.recognizer.recognize_sphinx(audio)
            elif self.engine == "bing":
                # Requires Bing API key
                return self.recognizer.recognize_bing(audio, language=self.language)
            elif self.engine == "azure":
                # Requires Azure API key
                return self.recognizer.recognize_azure(audio, language=self.language)
            else:
                # Default to Google
                return self.recognizer.recognize_google(audio, language=self.language)
        except Exception as e:
            self.logger.error(f"Recognition error with {self.engine}: {e}")
            return None
    
    def start_continuous_listening(self, callback: Callable[[str], None]):
        """
        Start continuous listening for voice commands.
        
        Args:
            callback: Function to call with recognized text
        """
        if self.is_listening:
            self.logger.warning("Already listening - cannot start continuous mode")
            return
        
        self.is_listening = True
        self.listening_thread = threading.Thread(
            target=self._continuous_listen_worker, 
            args=(callback,)
        )
        self.listening_thread.daemon = True
        self.listening_thread.start()
        
        self.logger.info("Started continuous listening mode")
        if self.enable_voice_feedback:
            self.speak("Started continuous listening")
    
    def stop_continuous_listening(self):
        """Stop continuous listening."""
        if not self.is_listening:
            self.logger.warning("Not currently listening - cannot stop")
            return
        
        self.is_listening = False
        if self.listening_thread:
            self.listening_thread.join(timeout=2)
        
        self.logger.info("Stopped continuous listening mode")
        if self.enable_voice_feedback:
            self.speak("Stopped listening")
    
    def _continuous_listen_worker(self, callback: Callable[[str], None]):
        """Worker thread for continuous listening."""
        self.logger.debug("Continuous listening worker thread started")
        
        while self.is_listening:
            try:
                with self.microphone as source:
                    # Listen for audio with shorter timeout for responsiveness
                    audio = self.recognizer.listen(source, timeout=1, phrase_time_limit=10)
                
                if not self.is_listening:
                    break
                
                # Recognize speech
                text = self._recognize_speech(audio)
                
                if text and self.is_listening:
                    self.logger.info(f"Continuous mode recognized: '{text}'")
                    callback(text.lower())
                    
            except sr.WaitTimeoutError:
                # Normal timeout, continue listening
                continue
            except sr.UnknownValueError:
                # Could not understand, continue listening
                continue
            except Exception as e:
                self.logger.error(f"Error in continuous listening: {e}")
                time.sleep(0.5)  # Brief pause before retrying
        
        self.logger.debug("Continuous listening worker thread stopped")
    
    def test_microphone(self):
        """Test microphone and recognition system."""
        self.logger.info("Starting microphone test")
        self.logger.info("Available microphones:")
        
        for i, mic_name in enumerate(sr.Microphone.list_microphone_names()):
            self.logger.info(f"  {i}: {mic_name}")
        
        self.logger.info("Using default microphone for test...")
        self.logger.info("Speak something for 3 seconds...")
        
        result = self.listen_once(timeout=3)
        if result:
            self.logger.info(f"Microphone test successful! Recognized: '{result}'")
            return True
        else:
            self.logger.error("Microphone test failed - no speech recognized")
            return False
    
    def get_available_engines(self) -> List[str]:
        """Get list of available speech recognition engines."""
        engines = ["google"]  # Google is usually always available
        
        # Test other engines
        try:
            sr.Recognizer().recognize_sphinx
            engines.append("sphinx")
        except AttributeError:
            pass
        
        return engines