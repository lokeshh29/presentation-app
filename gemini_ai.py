"""
Gemini AI Integration for Voice-Controlled PowerPoint Generation

This module integrates Google's Gemini AI for intelligent content generation
in presentations. It provides AI-powered slide content, titles, and formatting
suggestions based on voice commands and context.
"""

import asyncio
import time
from typing import Optional, List, Dict, Any, Tuple
from dataclasses import dataclass
from config import get_config, get_logger

try:
    import google.generativeai as genai
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False


@dataclass
class AIGeneratedContent:
    """Container for AI-generated presentation content."""
    title: str
    content: str
    bullet_points: List[str]
    suggested_layout: int = 1
    confidence: float = 1.0
    metadata: Dict[str, Any] = None


class GeminiAI:
    """
    Gemini AI integration for intelligent presentation content generation.
    """
    
    def __init__(self):
        """Initialize Gemini AI with configuration from environment."""
        self.config = get_config()
        self.logger = get_logger()
        self.model = None
        self.is_initialized = False
        
        if not GEMINI_AVAILABLE:
            self.logger.error("google-generativeai package not installed. Install with: pip install google-generativeai")
            return
        
        self._initialize_gemini()
    
    def _initialize_gemini(self) -> bool:
        """
        Initialize Gemini AI with API key from configuration.
        
        Returns:
            bool: True if initialization successful
        """
        try:
            api_key = self.config.gemini_api_key
            
            if not api_key or api_key == "your_gemini_api_key_here":
                self.logger.error("Gemini API key not configured. Please set GEMINI_API_KEY in .env file")
                return False
            
            # Configure Gemini
            genai.configure(api_key=api_key)
            
            # Initialize the model
            self.model = genai.GenerativeModel('gemini-1.5-flash')
            
            # Test the connection
            test_response = self.model.generate_content("Hello")
            if test_response:
                self.logger.info("Gemini AI initialized successfully")
                self.is_initialized = True
                return True
            else:
                self.logger.error("Failed to get response from Gemini AI")
                return False
                
        except Exception as e:
            self.logger.error(f"Failed to initialize Gemini AI: {str(e)}")
            return False
    
    def is_available(self) -> bool:
        """Check if Gemini AI is available and initialized."""
        return GEMINI_AVAILABLE and self.is_initialized
    
    def generate_slide_content(self, topic: str, slide_type: str = "content", 
                             context: str = "") -> Optional[AIGeneratedContent]:
        """
        Generate slide content based on topic and type.
        
        Args:
            topic (str): Main topic for the slide
            slide_type (str): Type of slide ("title", "content", "conclusion", "intro")
            context (str): Additional context or previous slides content
            
        Returns:
            AIGeneratedContent: Generated content or None if failed
        """
        if not self.is_available():
            self.logger.warning("Gemini AI not available for content generation")
            return None
        
        try:
            prompt = self._build_content_prompt(topic, slide_type, context)
            self.logger.debug(f"Generating content with prompt: {prompt[:100]}...")
            
            response = self.model.generate_content(prompt)
            
            if response and response.text:
                content = self._parse_content_response(response.text, slide_type)
                self.logger.info(f"Generated content for topic: {topic}")
                return content
            else:
                self.logger.warning("Empty response from Gemini AI")
                return None
                
        except Exception as e:
            self.logger.error(f"Error generating slide content: {str(e)}")
            return None
    
    def generate_presentation_outline(self, topic: str, slide_count: int = 5) -> Optional[List[str]]:
        """
        Generate a presentation outline with slide titles.
        
        Args:
            topic (str): Main presentation topic
            slide_count (int): Number of slides to generate
            
        Returns:
            List[str]: List of slide titles or None if failed
        """
        if not self.is_available():
            self.logger.warning("Gemini AI not available for outline generation")
            return None
        
        try:
            prompt = f"""Create a presentation outline for the topic: "{topic}"

Generate exactly {slide_count} slide titles that would make a comprehensive presentation.
Format your response as a numbered list:
1. [First slide title]
2. [Second slide title]
...

Guidelines:
- Start with an introduction/title slide
- Include 2-3 main content slides
- End with a conclusion if appropriate
- Keep titles concise and engaging
- Focus on key aspects of the topic"""

            self.logger.debug(f"Generating outline for: {topic}")
            
            response = self.model.generate_content(prompt)
            
            if response and response.text:
                outline = self._parse_outline_response(response.text)
                self.logger.info(f"Generated outline with {len(outline)} slides for: {topic}")
                return outline
            else:
                self.logger.warning("Empty response from Gemini AI for outline")
                return None
                
        except Exception as e:
            self.logger.error(f"Error generating presentation outline: {str(e)}")
            return None
    
    def enhance_voice_command(self, voice_text: str) -> Optional[str]:
        """
        Enhance and clarify voice commands using AI.
        
        Args:
            voice_text (str): Raw voice recognition text
            
        Returns:
            str: Enhanced/clarified command or None if failed
        """
        if not self.is_available():
            return voice_text  # Return original if AI not available
        
        try:
            prompt = f"""Analyze this voice command for presentation control and enhance it if needed:
Voice input: "{voice_text}"

If this appears to be a presentation command, enhance it for clarity and completeness.
If it's unclear, suggest the most likely intended command.
If it's not a presentation command, return the original text.

Examples:
"add slide about sales" → "create new slide with title Sales Report"
"make chart" → "add column chart"
"change color blue" → "change background to blue"

Enhanced command:"""

            response = self.model.generate_content(prompt)
            
            if response and response.text:
                enhanced = response.text.strip()
                if enhanced != voice_text:
                    self.logger.info(f"Enhanced voice command: '{voice_text}' → '{enhanced}'")
                return enhanced
            else:
                return voice_text
                
        except Exception as e:
            self.logger.error(f"Error enhancing voice command: {str(e)}")
            return voice_text
    
    def suggest_chart_data(self, topic: str, chart_type: str = "column") -> Optional[Dict[str, Any]]:
        """
        Generate sample chart data relevant to the topic.
        
        Args:
            topic (str): Chart topic or context
            chart_type (str): Type of chart to generate data for
            
        Returns:
            Dict: Chart data structure or None if failed
        """
        if not self.is_available():
            self.logger.warning("Gemini AI not available for chart data generation")
            return None
        
        try:
            prompt = f"""Generate realistic sample data for a {chart_type} chart about: "{topic}"

Create data that would be appropriate for this topic with:
- 4-6 categories/time periods
- 2-3 data series with realistic values
- Data that tells a meaningful story

Format the response as:
Categories: [list of categories]
Series 1 Name: [name]
Series 1 Values: [list of values]
Series 2 Name: [name] 
Series 2 Values: [list of values]"""

            response = self.model.generate_content(prompt)
            
            if response and response.text:
                chart_data = self._parse_chart_data_response(response.text)
                self.logger.info(f"Generated chart data for: {topic}")
                return chart_data
            else:
                self.logger.warning("Empty response from Gemini AI for chart data")
                return None
                
        except Exception as e:
            self.logger.error(f"Error generating chart data: {str(e)}")
            return None
    
    def _build_content_prompt(self, topic: str, slide_type: str, context: str) -> str:
        """Build prompt for content generation."""
        base_prompt = f"""Generate content for a presentation slide about: "{topic}"

Slide type: {slide_type}
Language: {self.config.ai_content_language}
"""
        
        if context:
            base_prompt += f"Previous context: {context}\n"
        
        if slide_type == "title":
            base_prompt += """
Create a compelling title slide with:
Title: [Main presentation title]
Subtitle: [Brief subtitle or tagline]
"""
        elif slide_type == "content":
            base_prompt += """
Create engaging content with:
Title: [Slide title]
Content: [3-5 bullet points or main content]
"""
        elif slide_type == "conclusion":
            base_prompt += """
Create a conclusion slide with:
Title: [Conclusion title like "Key Takeaways" or "Summary"]
Content: [3-4 key points or conclusions]
"""
        else:
            base_prompt += """
Create appropriate content with:
Title: [Relevant slide title]
Content: [Relevant content for the slide type]
"""
        
        base_prompt += """
Guidelines:
- Keep content concise and presentation-friendly
- Use bullet points where appropriate
- Make titles engaging and descriptive
- Ensure content is professional and informative
"""
        
        return base_prompt
    
    def _parse_content_response(self, response_text: str, slide_type: str) -> AIGeneratedContent:
        """Parse AI response into structured content."""
        lines = response_text.strip().split('\n')
        
        title = ""
        content = ""
        bullet_points = []
        
        current_section = None
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            if line.lower().startswith('title:'):
                title = line[6:].strip()
                current_section = 'title'
            elif line.lower().startswith('subtitle:'):
                if not content:
                    content = line[9:].strip()
                current_section = 'content'
            elif line.lower().startswith('content:'):
                content = line[8:].strip()
                current_section = 'content'
            elif line.startswith('•') or line.startswith('-') or line.startswith('*'):
                bullet_point = line[1:].strip()
                bullet_points.append(bullet_point)
            elif current_section == 'content' and not line.lower().startswith(('title:', 'content:')):
                if content:
                    content += '\n' + line
                else:
                    content = line
        
        # If no title found, use first line or generate from topic
        if not title and lines:
            title = lines[0].strip()
        
        # Combine bullet points into content if content is empty
        if not content and bullet_points:
            content = '\n'.join(f"• {point}" for point in bullet_points)
        
        # Determine suggested layout based on content
        suggested_layout = 1  # Default to content layout
        if slide_type == "title":
            suggested_layout = 0
        elif len(bullet_points) > 5:
            suggested_layout = 3  # Two content layout
        
        return AIGeneratedContent(
            title=title,
            content=content,
            bullet_points=bullet_points,
            suggested_layout=suggested_layout,
            confidence=0.8,
            metadata={"slide_type": slide_type}
        )
    
    def _parse_outline_response(self, response_text: str) -> List[str]:
        """Parse outline response into list of slide titles."""
        lines = response_text.strip().split('\n')
        titles = []
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            # Look for numbered list items
            if line[0].isdigit() and '.' in line:
                # Extract title after number
                parts = line.split('.', 1)
                if len(parts) > 1:
                    title = parts[1].strip()
                    titles.append(title)
            elif line.startswith('-') or line.startswith('•'):
                # Handle bullet points
                title = line[1:].strip()
                titles.append(title)
        
        return titles
    
    def _parse_chart_data_response(self, response_text: str) -> Optional[Dict[str, Any]]:
        """Parse chart data response into structured format."""
        try:
            lines = response_text.strip().split('\n')
            
            categories = []
            series = []
            
            current_series = None
            
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                    
                if line.lower().startswith('categories:'):
                    # Extract categories
                    cat_text = line[11:].strip()
                    categories = [cat.strip() for cat in cat_text.replace('[', '').replace(']', '').split(',')]
                elif line.lower().startswith('series') and 'name:' in line.lower():
                    # Start new series
                    name = line.split(':', 1)[1].strip()
                    current_series = {'name': name, 'values': []}
                elif line.lower().startswith('series') and 'values:' in line.lower():
                    # Add values to current series
                    if current_series:
                        values_text = line.split(':', 1)[1].strip()
                        try:
                            values = [float(val.strip()) for val in values_text.replace('[', '').replace(']', '').split(',')]
                            current_series['values'] = values
                            series.append(current_series)
                            current_series = None
                        except ValueError:
                            continue
            
            if categories and series:
                return {
                    'categories': categories,
                    'series': series
                }
            else:
                self.logger.warning("Failed to parse chart data from AI response")
                return None
                
        except Exception as e:
            self.logger.error(f"Error parsing chart data response: {str(e)}")
            return None


# Global AI instance
_gemini_ai: Optional[GeminiAI] = None


def get_gemini_ai() -> GeminiAI:
    """Get the global Gemini AI instance."""
    global _gemini_ai
    if _gemini_ai is None:
        _gemini_ai = GeminiAI()
    return _gemini_ai