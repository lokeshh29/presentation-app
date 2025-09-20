# AI-PPT Agent 🎤📊

An AI-powered voice-controlled PowerPoint generation system with a comprehensive REST API.

## Features

- 🎤 **Voice-to-Text**: Convert voice commands to actionable PPT operations
- 🤖 **AI Integration**: Gemini AI for intelligent content generation
- 📊 **PPT Manipulation**: Complete suite of PowerPoint editing functions
- 🌐 **REST API**: FastAPI-based endpoints for web integration
- 🔄 **Real-time Processing**: Live voice command processing
- 📁 **File Management**: Upload, download, and manage presentations

## Project Structure

```
ai-ppt-agent/
├── app.py                  # Main FastAPI application with all endpoints
├── config.py              # Configuration and logging setup
├── ppt_generator.py        # PowerPoint manipulation functions
├── voice_to_text.py        # Voice recognition pipeline
├── voice_command_parser.py # Natural language command parsing
├── gemini_ai.py           # AI integration for content generation
├── voice_controlled_ppt.py # Integrated voice control system
├── requirements.txt        # Python dependencies
├── .env                   # Environment configuration (API keys, settings)
├── .gitignore            # Git ignore rules
├── outputs/              # Generated PowerPoint files
├── uploads/              # Uploaded images/documents
├── logs/                 # Application logs
├── frontend/             # (Future) Web interface
│   ├── index.html
│   ├── app.js
│   └── components/
└── demos/                # Demo scripts and examples
```

## Quick Start

### 1. Installation

```bash
# Clone the repository
git clone <repository-url>
cd presentation-app

# Install dependencies
pip install -r requirements.txt
```

### 2. Configuration

```bash
# Copy environment template
cp .env.template .env

# Edit .env and add your Gemini API key
GEMINI_API_KEY=your_gemini_api_key_here
```

### 3. Run the Server

```bash
# Start the FastAPI server
python app.py

# Or use uvicorn directly
uvicorn app:app --host 0.0.0.0 --port 8000 --reload
```

The API will be available at:
- **Server**: http://localhost:8000
- **API Documentation**: http://localhost:8000/api/docs
- **Interactive API**: http://localhost:8000/api/redoc

## API Endpoints

### Core Endpoints

#### 🏠 Root & Health
- `GET /` - Root endpoint (serves frontend or API info)
- `GET /api/health` - System health check
- `GET /api/config` - Current configuration
- `GET /api/system-info` - System information

#### 🎤 Voice Processing
- `POST /api/voice-command` - Process voice commands
- `POST /api/voice-test` - Test voice recognition

#### 📊 PowerPoint Operations
- `POST /api/ppt-action` - Execute PPT actions (add slide, delete slide, etc.)
- `GET /api/presentations` - List all presentations
- `GET /api/presentations/{filename}` - Download presentation
- `DELETE /api/presentations/{filename}` - Delete presentation

#### 🤖 AI Content Generation
- `POST /api/ai-content` - Generate AI-powered content

#### 📁 File Management
- `POST /api/upload` - Upload files (images, documents)

### **Core Voice Recognition**
- **Speech Recognition Engine**: Uses `speechrecognition` library with multiple backend options
- **Continuous Listening**: Real-time voice command processing
- **Voice Feedback**: Text-to-speech responses using `pyttsx3`
- **Multi-Engine Support**: Google (default), Sphinx, Bing, Azure speech recognition

### **Natural Language Command Processing**
- **Smart Command Parsing**: Interprets natural language voice commands
- **Context-Aware**: Understands variations in command phrasing
- **Parameter Extraction**: Automatically extracts titles, colors, numbers from speech
- **Fuzzy Matching**: Handles imperfect voice recognition

### **PowerPoint Operations via Voice**

#### 📄 **Slide Management Commands**
```
"Create new slide"
"Add slide with title My Presentation"
"Delete slide number 2"
"Change layout to title and content"
```

#### ✏️ **Content Commands** 
```
"Change title to Project Overview"
"Add content This is my presentation"
"Update text Welcome to our meeting"
```

#### 📊 **Chart Commands**
```
"Add column chart"
"Insert bar chart"
"Create pie chart"
"Add line graph"
```

#### 🖼️ **Image Commands**
```
"Add image from sample.png"
"Insert picture from my_image.jpg"
```

#### 🎨 **Styling Commands**
```
"Change background to blue"
"Set background red"
"Make background light green"
```

#### 💾 **File Operations**
```
"Save presentation as my_demo"
"Save as quarterly_report"
"Export presentation"
```

## 🚀 Quick Start

### Installation

```bash
# Install required packages
pip install speechrecognition pyaudio pyttsx3 python-pptx Pillow

# Or run setup script
python setup.py
```

### Basic Usage

#### **Option 1: Interactive Single Commands**
```python
from voice_controlled_ppt import VoiceControlledPPT

# Initialize the system
voice_ppt = VoiceControlledPPT(enable_voice_feedback=True)

# Process single voice command
voice_ppt.process_single_command(timeout=10)
```

#### **Option 2: Continuous Voice Control**
```python
from voice_controlled_ppt import VoiceControlledPPT

# Start continuous listening mode
voice_ppt = VoiceControlledPPT()
voice_ppt.start_voice_control()  # Runs until you say "stop listening"
```

#### **Option 3: Run Complete Demo**
```bash
python demo_voice_pipeline.py
```

## 📁 Project Structure

### **Core Voice-to-Text Files**
- **`voice_to_text.py`** - Speech recognition and audio processing
- **`voice_command_parser.py`** - Natural language command interpretation  
- **`voice_controlled_ppt.py`** - Integration layer connecting voice to PowerPoint
- **`demo_voice_pipeline.py`** - Complete demonstration system

### **PowerPoint Generation Files**  
- **`ppt_generator.py`** - PowerPoint creation and manipulation
- **`demo_ppt_generation.py`** - PowerPoint generation demo

### **Setup & Configuration**
- **`setup.py`** - Automated package installation and testing
- **`requirements.txt`** - Python package dependencies

## 🎯 Voice Recognition Features

### **Multiple Recognition Engines**
```python
# Google Speech Recognition (default, free)
voice_system = VoiceToText(engine="google")

# Offline recognition with Sphinx
voice_system = VoiceToText(engine="sphinx") 

# Microsoft Bing (requires API key)
voice_system = VoiceToText(engine="bing")
```

### **Voice Feedback System**
- Real-time spoken confirmation of commands
- Error notifications via speech
- Status updates and help via voice

### **Robust Audio Processing**
- Automatic noise cancellation calibration
- Timeout handling for responsive interaction
- Multi-threaded continuous listening
- Microphone testing and validation

## 🔧 Advanced Configuration

### **Customizing Voice Recognition**
```python
voice_system = VoiceToText(
    engine="google",           # Recognition engine
    language="en-US",          # Language setting  
    enable_voice_feedback=True # Text-to-speech feedback
)

# Test microphone setup
voice_system.test_microphone()

# Single command recognition
result = voice_system.listen_once(timeout=5)
```

### **Command Parser Customization**
```python
parser = VoiceCommandParser()

# Parse natural language
command = parser.parse_command("create slide with title My Demo")

# Get available commands
help_text = parser.get_help_text()
```

## 🎮 Demo Modes

Run `python demo_voice_pipeline.py` and choose from:

1. **Basic Voice Recognition Test** - Test microphone and recognition
2. **Command Parsing Demo** - See how commands are interpreted
3. **Interactive Voice Commands** - Process 5 individual commands
4. **Continuous Voice Control** - Full hands-free operation
5. **Run All Demos** - Complete system demonstration

## 🔊 Voice Commands Reference

### **System Commands**
- `"Help"` - Show available commands
- `"Stop listening"` - Exit continuous mode

### **Slide Navigation**
- `"Go to slide number 3"` - Navigate to specific slide
- `"Show slide 1"` - Display slide

### **Advanced Operations**
- `"Change background to light blue"` - Custom colors
- `"Create title slide with title Company Overview"` - Specific layouts
- `"Add chart with quarterly data"` - Data-driven charts

## 🛠️ System Requirements

### **Audio Requirements**
- **Microphone**: Any system microphone or headset
- **Audio Drivers**: Up-to-date audio drivers
- **Network**: Internet connection for Google Speech Recognition (recommended)

### **Python Packages**
- `speechrecognition>=3.10.0` - Speech recognition engine
- `pyaudio>=0.2.11` - Microphone access
- `pyttsx3>=2.90` - Text-to-speech feedback  
- `python-pptx>=0.6.21` - PowerPoint generation
- `Pillow>=8.0.0` - Image processing (optional)

### **Platform-Specific Notes**
- **Windows**: PyAudio should install directly with pip
- **macOS**: May need `brew install portaudio`
- **Linux**: May need `sudo apt-get install python3-pyaudio`

## 🚨 Troubleshooting

### **Common Issues**

#### **Microphone Not Working**
```python
# Test microphone setup
from voice_to_text import VoiceToText
voice = VoiceToText()
voice.test_microphone()
```

#### **Recognition Accuracy Issues**
- Speak clearly and at normal pace
- Ensure quiet environment
- Check microphone positioning
- Try different recognition engines

#### **Package Installation Issues**
```bash
# Run setup script for automated installation
python setup.py

# Manual installation troubleshooting
pip install --upgrade pip
pip install pyaudio --force-reinstall
```

## 🎯 Next Steps & Integration

The voice-to-text pipeline is ready for integration with:

### **AI Content Generation**
```python
# Integration with OpenAI/Ollama for AI-generated content
voice_ppt.add_ai_content_generation(api_key="your_key")
```

### **Web Framework Integration**
- **FastAPI**: RESTful voice-controlled presentation API
- **Streamlit**: Interactive web interface for voice commands

### **Advanced Features**
- Custom vocabulary training
- Multiple language support
- Voice biometrics for user identification
- Real-time collaboration via voice commands

## 📊 Performance & Statistics

The system tracks:
- Command recognition accuracy
- Processing speed and response time
- Voice command success rates  
- Presentation generation metrics

## 🤝 Contributing

The voice-to-text pipeline is modular and extensible:
- Add new voice commands in `voice_command_parser.py`
- Extend PowerPoint operations in `ppt_generator.py`  
- Customize voice feedback in `voice_to_text.py`

---

**🎤 Ready to create presentations with your voice? Run `python demo_voice_pipeline.py` to get started!**