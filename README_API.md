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

### API Request Examples

#### Voice Command Processing
```json
POST /api/voice-command
{
  "text_command": "Create a new slide with title 'Introduction to AI'",
  "enhance_with_ai": true,
  "session_id": "user123"
}
```

#### PPT Action Execution
```json
POST /api/ppt-action
{
  "action": "add_slide",
  "parameters": {
    "title": "My New Slide",
    "content": "This is the slide content",
    "layout": 1
  },
  "auto_save": true
}
```

#### AI Content Generation
```json
POST /api/ai-content
{
  "topic": "Machine Learning Basics",
  "content_type": "presentation_outline",
  "slide_count": 5,
  "tone": "professional"
}
```

## Available PPT Actions

| Action | Description | Parameters |
|--------|-------------|------------|
| `add_slide` | Add a new slide | `title`, `content`, `layout` |
| `delete_slide` | Delete a slide | `slide_number` |
| `modify_layout` | Change slide layout | `slide_number`, `layout` |
| `insert_image` | Add image to slide | `slide_number`, `image_path` |
| `insert_chart` | Add chart to slide | `slide_number`, `chart_data`, `chart_type` |
| `update_text` | Update slide text | `slide_number`, `text` |
| `change_background` | Change slide background | `slide_number`, `color` |

## Voice Commands

The system understands natural language commands like:

- "Create a new slide about machine learning"
- "Add an image to slide 2"
- "Delete the last slide"
- "Change the background color to blue"
- "Generate content about artificial intelligence"
- "Insert a bar chart showing sales data"

## Environment Variables

```env
# Required
GEMINI_API_KEY=your_gemini_api_key_here

# Optional (with defaults)
VOICE_RECOGNITION_ENGINE=google
VOICE_LANGUAGE=en-US
ENABLE_VOICE_FEEDBACK=true
LOG_LEVEL=INFO
AUTO_SAVE=true
DEFAULT_PRESENTATION_NAME=ai_voice_presentation
MAX_SLIDE_COUNT=50
AI_TEMPERATURE=0.7
AI_MAX_TOKENS=1000
```

## Development

### Running in Development Mode

```bash
# Install development dependencies
pip install -r requirements.txt

# Run with auto-reload
uvicorn app:app --reload --log-level debug

# Run tests
python -m pytest tests/

# Check API documentation
open http://localhost:8000/api/docs
```

### API Testing

Use the interactive API documentation at `/api/docs` or tools like curl/Postman:

```bash
# Health check
curl http://localhost:8000/api/health

# Process text command
curl -X POST "http://localhost:8000/api/voice-command" \
  -H "Content-Type: application/json" \
  -d '{"text_command": "Create a slide about AI"}'

# Upload file
curl -X POST "http://localhost:8000/api/upload" \
  -F "file=@image.jpg"
```

## Demo Scripts

Run the included demo scripts to test functionality:

```bash
# Basic PPT generation demo
python demo_ppt_generator.py

# Voice recognition demo
python demo_voice_recognition.py

# AI integration demo
python demo_ai_voice_pipeline.py

# Complete voice-controlled PPT system
python demo_voice_controlled_ppt.py
```

## Troubleshooting

### Common Issues

1. **Microphone Access**: Ensure PyAudio is properly installed and microphone permissions are granted
2. **API Key**: Make sure your Gemini API key is valid and set in `.env`
3. **Dependencies**: Install all requirements with `pip install -r requirements.txt`
4. **Port Conflicts**: Change the port in `app.py` if 8000 is already in use

### Logging

Logs are written to `logs/presentation_app.log` with configurable levels:
- DEBUG: Detailed information
- INFO: General information (default)
- WARNING: Warning messages
- ERROR: Error messages

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For issues and questions:
1. Check the API documentation at `/api/docs`
2. Review the logs in `logs/presentation_app.log`
3. Test individual components with the demo scripts
4. Open an issue on GitHub

---

**Made with ❤️ for voice-controlled productivity**