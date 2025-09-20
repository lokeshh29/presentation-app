"""
Voice Command Parser for PowerPoint Operations

This module parses voice commands and converts them into PowerPoint
operations that can be executed by the PPTGenerator class.

Supported voice commands:
- "create new slide with title [title]"
- "add slide"
- "delete slide number [n]"
- "change background to [color]"
- "insert chart with data [data]"
- "add image from [path]"
- "update title to [text]"
- "save presentation as [filename]"
"""

import re
from typing import Dict, List, Optional, Tuple, Any
from dataclasses import dataclass


@dataclass
class VoiceCommand:
    """Represents a parsed voice command."""
    action: str
    parameters: Dict[str, Any]
    confidence: float = 1.0
    raw_text: str = ""


class VoiceCommandParser:
    """
    Parser for converting voice commands into structured PowerPoint operations.
    """
    
    def __init__(self):
        """Initialize the command parser with predefined patterns."""
        self.command_patterns = self._initialize_patterns()
        self.color_map = self._initialize_color_map()
        self.layout_map = self._initialize_layout_map()
    
    def _initialize_patterns(self) -> Dict[str, Dict[str, Any]]:
        """Initialize voice command patterns and their mappings."""
        return {
            # Slide management commands
            "add_slide": {
                "patterns": [
                    r"(?:create|add|make|new)\s+(?:new\s+)?slide",
                    r"(?:create|add|make)\s+(?:a\s+)?(?:new\s+)?slide\s+with\s+title\s+(.+)",
                    r"(?:create|add|make)\s+(?:a\s+)?(.+)\s+slide",
                ],
                "action": "add_slide",
                "extract_params": ["title", "layout"]
            },
            
            "delete_slide": {
                "patterns": [
                    r"(?:delete|remove)\s+slide\s+(?:number\s+)?(\d+)",
                    r"(?:delete|remove)\s+(?:the\s+)?(\d+)(?:st|nd|rd|th)?\s+slide",
                ],
                "action": "delete_slide",
                "extract_params": ["slide_number"]
            },
            
            # Content commands
            "update_title": {
                "patterns": [
                    r"(?:change|update|set)\s+(?:the\s+)?title\s+to\s+(.+)",
                    r"(?:make|set)\s+(?:the\s+)?title\s+(.+)",
                    r"title\s+should\s+be\s+(.+)",
                ],
                "action": "update_text",
                "extract_params": ["title"]
            },
            
            "update_content": {
                "patterns": [
                    r"(?:add|set|change)\s+content\s+(.+)",
                    r"(?:add|set|change)\s+(?:the\s+)?text\s+(.+)",
                    r"(?:write|type)\s+(.+)",
                ],
                "action": "update_text",
                "extract_params": ["content"]
            },
            
            # Chart commands
            "insert_chart": {
                "patterns": [
                    r"(?:add|insert|create)\s+(?:a\s+)?(\w+)\s+chart",
                    r"(?:add|insert|create)\s+(?:a\s+)?chart\s+(?:of\s+type\s+)?(\w+)",
                    r"(?:make|create)\s+(?:a\s+)?chart",
                ],
                "action": "insert_chart",
                "extract_params": ["chart_type"]
            },
            
            # Image commands
            "insert_image": {
                "patterns": [
                    r"(?:add|insert)\s+(?:an?\s+)?image\s+(?:from\s+)?(.+)",
                    r"(?:add|insert)\s+(?:a\s+)?picture\s+(?:from\s+)?(.+)",
                    r"(?:load|open)\s+image\s+(.+)",
                ],
                "action": "insert_image",
                "extract_params": ["image_path"]
            },
            
            # Background commands
            "change_background": {
                "patterns": [
                    r"(?:change|set)\s+(?:the\s+)?background\s+(?:to\s+)?(\w+)",
                    r"(?:make|set)\s+background\s+(\w+)",
                    r"background\s+(?:should\s+be\s+)?(\w+)",
                ],
                "action": "change_background",
                "extract_params": ["color"]
            },
            
            # Layout commands
            "change_layout": {
                "patterns": [
                    r"(?:change|set)\s+(?:the\s+)?layout\s+to\s+(.+)",
                    r"(?:use|apply)\s+(.+)\s+layout",
                    r"(?:make|set)\s+(?:this\s+)?(?:a\s+)?(.+)\s+slide",
                ],
                "action": "modify_layout",
                "extract_params": ["layout"]
            },
            
            # File operations
            "save_presentation": {
                "patterns": [
                    r"(?:save|export)\s+(?:the\s+)?presentation\s+(?:as\s+)?(.+)",
                    r"(?:save|export)\s+(?:as\s+)?(.+)",
                    r"(?:save|export)\s+(?:the\s+)?(?:file|presentation)",
                ],
                "action": "save",
                "extract_params": ["filename"]
            },
            
            # Navigation commands
            "go_to_slide": {
                "patterns": [
                    r"(?:go\s+to|open|show)\s+slide\s+(?:number\s+)?(\d+)",
                    r"(?:go\s+to|open|show)\s+(?:the\s+)?(\d+)(?:st|nd|rd|th)?\s+slide",
                ],
                "action": "go_to_slide",
                "extract_params": ["slide_number"]
            },
            
            # System commands
            "help": {
                "patterns": [
                    r"(?:help|what\s+can\s+you\s+do)",
                    r"(?:show|list)\s+(?:available\s+)?commands",
                ],
                "action": "help",
                "extract_params": []
            },
            
            "stop_listening": {
                "patterns": [
                    r"(?:stop|quit|exit)\s+(?:listening|recognition)",
                    r"(?:stop|quit|exit)",
                ],
                "action": "stop_listening",
                "extract_params": []
            }
        }
    
    def _initialize_color_map(self) -> Dict[str, Tuple[int, int, int]]:
        """Initialize color name to RGB mapping."""
        return {
            "red": (255, 0, 0),
            "green": (0, 255, 0),
            "blue": (0, 0, 255),
            "yellow": (255, 255, 0),
            "orange": (255, 165, 0),
            "purple": (128, 0, 128),
            "pink": (255, 192, 203),
            "black": (0, 0, 0),
            "white": (255, 255, 255),
            "gray": (128, 128, 128),
            "grey": (128, 128, 128),
            "light blue": (173, 216, 230),
            "dark blue": (0, 0, 139),
            "light green": (144, 238, 144),
            "dark green": (0, 100, 0),
        }
    
    def _initialize_layout_map(self) -> Dict[str, int]:
        """Initialize layout name to index mapping."""
        return {
            "title": 0,
            "title slide": 0,
            "content": 1,
            "title and content": 1,
            "section": 2,
            "section header": 2,
            "two content": 3,
            "comparison": 4,
            "title only": 5,
            "blank": 6,
            "caption": 7,
            "content with caption": 7,
            "picture": 8,
            "picture with caption": 8,
        }
    
    def parse_command(self, text: str) -> Optional[VoiceCommand]:
        """
        Parse a voice command text into a structured command.
        
        Args:
            text (str): Raw voice command text
            
        Returns:
            VoiceCommand: Parsed command or None if no match found
        """
        if not text or not isinstance(text, str):
            return None
        
        text = text.lower().strip()
        
        # Try to match against all command patterns
        for command_name, command_info in self.command_patterns.items():
            for pattern in command_info["patterns"]:
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    # Extract parameters based on the pattern groups
                    parameters = self._extract_parameters(
                        match, command_info["extract_params"], text
                    )
                    
                    return VoiceCommand(
                        action=command_info["action"],
                        parameters=parameters,
                        confidence=0.8,
                        raw_text=text
                    )
        
        # If no exact match, try fuzzy matching for common commands
        fuzzy_command = self._fuzzy_match(text)
        if fuzzy_command:
            return fuzzy_command
        
        return None
    
    def _extract_parameters(self, match, param_names: List[str], text: str) -> Dict[str, Any]:
        """Extract parameters from regex match groups."""
        parameters = {}
        
        for i, param_name in enumerate(param_names):
            if i + 1 <= len(match.groups()) and match.group(i + 1):
                value = match.group(i + 1).strip()
                
                # Process specific parameter types
                if param_name == "slide_number":
                    try:
                        parameters[param_name] = int(value) - 1  # Convert to 0-based index
                    except ValueError:
                        parameters[param_name] = 0
                        
                elif param_name == "color":
                    parameters[param_name] = self._resolve_color(value)
                    
                elif param_name == "layout":
                    parameters[param_name] = self._resolve_layout(value)
                    
                elif param_name == "chart_type":
                    parameters[param_name] = self._resolve_chart_type(value)
                    
                else:
                    parameters[param_name] = value
        
        return parameters
    
    def _resolve_color(self, color_name: str) -> Tuple[int, int, int]:
        """Resolve color name to RGB tuple."""
        color_name = color_name.lower().strip()
        return self.color_map.get(color_name, (255, 255, 255))  # Default to white
    
    def _resolve_layout(self, layout_name: str) -> int:
        """Resolve layout name to layout index."""
        layout_name = layout_name.lower().strip()
        return self.layout_map.get(layout_name, 1)  # Default to content layout
    
    def _resolve_chart_type(self, chart_type: str) -> str:
        """Resolve chart type variations."""
        chart_type = chart_type.lower().strip()
        
        chart_map = {
            "bar": "bar",
            "column": "column",
            "line": "line",
            "pie": "pie",
            "graph": "column",  # Default graph to column
            "chart": "column",  # Default chart to column
        }
        
        return chart_map.get(chart_type, "column")
    
    def _fuzzy_match(self, text: str) -> Optional[VoiceCommand]:
        """Attempt fuzzy matching for commands that don't match exactly."""
        # Simple keyword-based fuzzy matching
        keywords = {
            "slide": "add_slide",
            "delete": "delete_slide",
            "title": "update_title",
            "chart": "insert_chart",
            "image": "insert_image",
            "picture": "insert_image",
            "background": "change_background",
            "save": "save_presentation",
            "help": "help",
        }
        
        for keyword, action in keywords.items():
            if keyword in text:
                return VoiceCommand(
                    action=action,
                    parameters={},
                    confidence=0.3,
                    raw_text=text
                )
        
        return None
    
    def get_help_text(self) -> str:
        """Get help text with available commands."""
        help_text = """
Available Voice Commands:

📄 Slide Management:
• "Create new slide" / "Add slide"
• "Create slide with title [your title]"
• "Delete slide number [n]"
• "Change layout to [layout name]"

✏️ Content:
• "Change title to [your title]"
• "Add content [your text]"
• "Update text [your text]"

📊 Charts:
• "Add column chart" / "Insert bar chart"
• "Create line chart" / "Add pie chart"

🖼️ Images:
• "Add image from [path]"
• "Insert picture [path]"

🎨 Styling:
• "Change background to [color]"
• "Set background [color name]"

💾 File Operations:
• "Save presentation as [filename]"
• "Save as [filename]"

🎤 System:
• "Help" - Show this help
• "Stop listening" - Stop voice recognition

Example: "Create new slide with title My Presentation"
Example: "Add column chart"
Example: "Change background to blue"
        """
        return help_text.strip()
    
    def validate_command(self, command: VoiceCommand) -> Tuple[bool, str]:
        """
        Validate if a command can be executed.
        
        Args:
            command: The voice command to validate
            
        Returns:
            Tuple of (is_valid, error_message)
        """
        if not command:
            return False, "No command provided"
        
        # Check required parameters for each action
        required_params = {
            "delete_slide": ["slide_number"],
            "insert_image": ["image_path"],
        }
        
        if command.action in required_params:
            for param in required_params[command.action]:
                if param not in command.parameters:
                    return False, f"Missing required parameter: {param}"
        
        return True, ""