"""
AI-PPT Agent - FastAPI Application

A comprehensive REST API for AI-powered voice-controlled PowerPoint generation.
Provides endpoints for voice processing, PPT manipulation, AI content generation,
and file management.
"""

import os
import sys
import json
import base64
from pathlib import Path
from datetime import datetime
from typing import List, Optional, Dict, Any, Union

# Try to import FastAPI - if not available, provide helpful error
try:
    from fastapi import FastAPI, HTTPException, UploadFile, File, Form, BackgroundTasks
    from fastapi.middleware.cors import CORSMiddleware
    from fastapi.staticfiles import StaticFiles
    from fastapi.responses import FileResponse, JSONResponse
    from pydantic import BaseModel, Field
    import uvicorn
    FASTAPI_AVAILABLE = True
except ImportError as e:
    print(f"⚠️  FastAPI not installed: {e}")
    print("📦 Install with: pip install fastapi uvicorn python-multipart")
    FASTAPI_AVAILABLE = False

# Add project root to path for imports
project_root = Path(__file__).parent
sys.path.append(str(project_root))

# Import existing modules with error handling
try:
    from ppt_generator import PPTGenerator
    PPT_AVAILABLE = True
except ImportError as e:
    print(f"⚠️  PPT Generator not available: {e}")
    PPT_AVAILABLE = False

try:
    from voice_to_text import VoiceToText
    VOICE_AVAILABLE = True
except ImportError as e:
    print(f"⚠️  Voice processor not available: {e}")
    VOICE_AVAILABLE = False

try:
    from voice_command_parser import VoiceCommandParser
    PARSER_AVAILABLE = True
except ImportError as e:
    print(f"⚠️  Command parser not available: {e}")
    PARSER_AVAILABLE = False

try:
    from gemini_ai import GeminiAI
    AI_AVAILABLE = True
except ImportError as e:
    print(f"⚠️  AI agent not available: {e}")
    AI_AVAILABLE = False

try:
    from config import get_config, get_logger
    CONFIG_AVAILABLE = True
except ImportError as e:
    print(f"⚠️  Config not available: {e}")
    CONFIG_AVAILABLE = False

# Exit if FastAPI is not available
if not FASTAPI_AVAILABLE:
    print("❌ Cannot start server without FastAPI. Please install required dependencies.")
    sys.exit(1)

# Initialize configuration and logging
if CONFIG_AVAILABLE:
    config = get_config()
    logger = get_logger()
else:
    import logging
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger(__name__)
    config = type('Config', (), {
        'voice_recognition_engine': 'google',
        'voice_language': 'en-US',
        'enable_voice_feedback': True,
        'log_level': 'INFO',
        'gemini_api_key': '',
        'ai_temperature': 0.7,
        'ai_max_tokens': 1000,
        'auto_save': True,
        'max_slide_count': 50
    })()

# Initialize FastAPI app
app = FastAPI(
    title="AI-PPT Agent",
    description="AI-powered voice-controlled PowerPoint generation system",
    version="1.0.0",
    docs_url="/api/docs",
    redoc_url="/api/redoc"
)

# Configure CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Configure appropriately for production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Mount static directories
try:
    frontend_path = project_root / "frontend"
    if frontend_path.exists():
        app.mount("/static", StaticFiles(directory=str(frontend_path)), name="static")
    
    outputs_path = project_root / "outputs"
    if outputs_path.exists():
        app.mount("/outputs", StaticFiles(directory=str(outputs_path)), name="outputs")
        
    uploads_path = project_root / "uploads"
    if uploads_path.exists():
        app.mount("/uploads", StaticFiles(directory=str(uploads_path)), name="uploads")
        
except Exception as e:
    logger.warning(f"Could not mount static directories: {e}")

# Initialize core components with error handling
ppt_generator = None
if PPT_AVAILABLE:
    try:
        ppt_generator = PPTGenerator()
    except Exception as e:
        logger.error(f"Failed to initialize PPT Generator: {e}")

voice_processor = None
if VOICE_AVAILABLE:
    try:
        voice_processor = VoiceToText()
    except Exception as e:
        logger.error(f"Failed to initialize Voice Processor: {e}")

command_parser = None
if PARSER_AVAILABLE:
    try:
        command_parser = VoiceCommandParser()
    except Exception as e:
        logger.error(f"Failed to initialize Command Parser: {e}")

ai_agent = None
if AI_AVAILABLE:
    try:
        ai_agent = GeminiAI()
    except Exception as e:
        logger.error(f"Failed to initialize AI Agent: {e}")

# Pydantic models for request/response validation
class APIResponse(BaseModel):
    """Standard API response model."""
    success: bool
    message: str
    data: Optional[Dict[str, Any]] = None
    error: Optional[str] = None
    timestamp: str = Field(default_factory=lambda: datetime.now().isoformat())


class VoiceCommandRequest(BaseModel):
    """Request model for voice command processing."""
    audio_data: Optional[str] = Field(None, description="Base64 encoded audio data")
    text_command: Optional[str] = Field(None, description="Direct text command")
    session_id: str = Field(default="default", description="Session identifier")
    enhance_with_ai: bool = Field(default=True, description="Use AI to enhance command")


class PPTActionRequest(BaseModel):
    """Request model for PPT manipulation actions."""
    action: str = Field(..., description="Action to perform (add_slide, delete_slide, etc.)")
    parameters: Dict[str, Any] = Field(default_factory=dict, description="Action parameters")
    presentation_path: Optional[str] = Field(None, description="Path to existing presentation")
    auto_save: bool = Field(default=True, description="Auto-save after action")


class AIContentRequest(BaseModel):
    """Request model for AI content generation."""
    topic: str = Field(..., description="Topic for content generation")
    content_type: str = Field(default="slide_content", description="Type of content to generate")
    slide_count: Optional[int] = Field(None, description="Number of slides to generate")
    additional_context: Optional[str] = Field(None, description="Additional context for AI")
    tone: str = Field(default="professional", description="Content tone")


# Root endpoint
@app.get("/", response_model=APIResponse)
async def root():
    """Root endpoint - serves frontend or API info."""
    try:
        frontend_file = project_root / "frontend" / "index.html"
        if frontend_file.exists():
            return FileResponse(str(frontend_file))
        
        return APIResponse(
            success=True,
            message="AI-PPT Agent API is running",
            data={
                "version": "1.0.0",
                "docs": "/api/docs",
                "status": "healthy",
                "components": {
                    "ppt_generator": PPT_AVAILABLE and ppt_generator is not None,
                    "voice_processor": VOICE_AVAILABLE and voice_processor is not None,
                    "ai_agent": AI_AVAILABLE and ai_agent is not None,
                    "command_parser": PARSER_AVAILABLE and command_parser is not None
                },
                "endpoints": {
                    "health": "/api/health",
                    "voice_command": "/api/voice-command",
                    "ppt_action": "/api/ppt-action",
                    "ai_content": "/api/ai-content",
                    "presentations": "/api/presentations"
                }
            }
        )
    except Exception as e:
        logger.error(f"Root endpoint error: {e}")
        raise HTTPException(status_code=500, detail=str(e))


# Health check endpoint
@app.get("/api/health", response_model=APIResponse)
async def health_check():
    """Check the health status of all system components."""
    try:
        # Check component availability
        components_status = {
            "ppt_generator": PPT_AVAILABLE and ppt_generator is not None,
            "voice_processor": VOICE_AVAILABLE and voice_processor is not None,
            "ai_agent": AI_AVAILABLE and ai_agent is not None and (ai_agent.is_available() if hasattr(ai_agent, 'is_available') else True),
            "command_parser": PARSER_AVAILABLE and command_parser is not None
        }
        
        # Check directories
        directories_status = {
            "outputs": (project_root / "outputs").exists(),
            "uploads": (project_root / "uploads").exists(),
            "logs": (project_root / "logs").exists()
        }
        
        all_healthy = any(components_status.values())  # At least one component should work
        
        return APIResponse(
            success=all_healthy,
            message="Health check completed",
            data={
                "status": "healthy" if all_healthy else "degraded",
                "components": components_status,
                "directories": directories_status,
                "config": {
                    "ai_available": bool(getattr(config, 'gemini_api_key', '')),
                    "voice_engine": getattr(config, 'voice_recognition_engine', 'unknown'),
                    "log_level": getattr(config, 'log_level', 'INFO')
                }
            }
        )
    except Exception as e:
        logger.error(f"Health check error: {e}")
        return APIResponse(
            success=False,
            message="Health check failed",
            error=str(e)
        )


# Voice command processing endpoint
@app.post("/api/voice-command", response_model=APIResponse)
async def process_voice_command(request: VoiceCommandRequest):
    """Process voice command and execute PPT actions."""
    try:
        logger.info(f"Processing voice command for session: {request.session_id}")
        
        if not voice_processor:
            raise HTTPException(status_code=503, detail="Voice processor not available")
        
        if not command_parser:
            raise HTTPException(status_code=503, detail="Command parser not available")
        
        # Convert voice to text if audio data provided
        recognized_text = ""
        if request.audio_data:
            try:
                # For now, we'll simulate audio processing
                raise HTTPException(status_code=501, detail="Audio processing not yet implemented - use text_command instead")
            except Exception as e:
                logger.error(f"Audio processing failed: {e}")
                raise HTTPException(status_code=400, detail=f"Audio processing failed: {str(e)}")
        elif request.text_command:
            recognized_text = request.text_command
        else:
            raise HTTPException(status_code=400, detail="Either audio_data or text_command is required")
        
        if not recognized_text:
            return APIResponse(
                success=False,
                message="No text could be extracted from the input",
                error="Empty recognition result"
            )
        
        logger.info(f"Recognized text: {recognized_text}")
        
        # Parse the command
        parsed_command = command_parser.parse_command(recognized_text)
        
        if not parsed_command:
            return APIResponse(
                success=False,
                message="Could not parse the command",
                error="Command parsing failed",
                data={"recognized_text": recognized_text}
            )
        
        # Enhance with AI if requested and available
        enhanced_command = parsed_command
        ai_enhancement_used = False
        if request.enhance_with_ai and ai_agent:
            try:
                # For now, just log that AI enhancement was requested
                logger.info("AI enhancement requested but not yet fully implemented")
                ai_enhancement_used = False
            except Exception as e:
                logger.warning(f"AI enhancement failed, using parsed command: {e}")
        
        # Execute the command using PPT generator if available
        execution_result = {}
        if ppt_generator and parsed_command:
            try:
                # Simple command execution based on command type
                if hasattr(parsed_command, 'command') and parsed_command.command == 'add_slide':
                    title = getattr(parsed_command, 'title', 'New Slide')
                    slide_num = ppt_generator.add_slide(title=title)
                    execution_result = {"action": "add_slide", "slide_number": slide_num, "title": title}
                else:
                    execution_result = {"message": "Command parsed but execution not yet fully implemented"}
            except Exception as e:
                logger.error(f"Command execution failed: {e}")
                execution_result = {"error": str(e)}
        
        return APIResponse(
            success=True,
            message="Voice command processed successfully",
            data={
                "session_id": request.session_id,
                "recognized_text": recognized_text,
                "parsed_command": str(parsed_command),
                "execution_result": execution_result,
                "ai_enhanced": ai_enhancement_used
            }
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Voice command processing error: {e}")
        return APIResponse(
            success=False,
            message="Failed to process voice command",
            error=str(e)
        )


# PPT action execution endpoint
@app.post("/api/ppt-action", response_model=APIResponse)
async def execute_ppt_action(request: PPTActionRequest):
    """Execute specific PowerPoint manipulation action."""
    try:
        logger.info(f"Executing PPT action: {request.action}")
        
        if not ppt_generator:
            raise HTTPException(status_code=503, detail="PPT generator not available")
        
        # Execute the action based on type
        result = None
        
        if request.action == "add_slide":
            title = request.parameters.get("title", "New Slide")
            content = request.parameters.get("content", "")
            layout_index = request.parameters.get("layout", 0)
            result = ppt_generator.add_slide(layout_index=layout_index, title=title, subtitle=content)
            
        elif request.action == "delete_slide":
            slide_number = request.parameters.get("slide_number")
            if slide_number is None:
                raise HTTPException(status_code=400, detail="slide_number parameter required")
            result = ppt_generator.delete_slide(slide_number)
            
        elif request.action == "update_text":
            slide_number = request.parameters.get("slide_number")
            new_text = request.parameters.get("text")
            if slide_number is None or new_text is None:
                raise HTTPException(status_code=400, detail="slide_number and text parameters required")
            result = ppt_generator.update_text(slide_number, new_text)
            
        else:
            raise HTTPException(status_code=400, detail=f"Action '{request.action}' not yet implemented")
        
        # Auto-save if requested
        save_path = None
        if request.auto_save:
            save_path = project_root / "outputs" / f"presentation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
            ppt_generator.save_presentation(str(save_path))
            result = result or {}
            if isinstance(result, dict):
                result["saved_to"] = str(save_path)
        
        return APIResponse(
            success=True,
            message=f"Action '{request.action}' executed successfully",
            data={
                "action": request.action,
                "parameters": request.parameters,
                "result": result,
                "saved_to": str(save_path) if save_path else None
            }
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"PPT action execution error: {e}")
        return APIResponse(
            success=False,
            message=f"Failed to execute action '{request.action}'",
            error=str(e)
        )


# AI content generation endpoint
@app.post("/api/ai-content", response_model=APIResponse)
async def generate_ai_content(request: AIContentRequest):
    """Generate AI-powered content for presentations."""
    try:
        logger.info(f"Generating AI content for topic: {request.topic}")
        
        if not ai_agent:
            raise HTTPException(status_code=503, detail="AI agent not available")
        
        # Generate content based on type
        content = None
        
        if request.content_type == "slide_content":
            content = ai_agent.generate_slide_content(
                topic=request.topic,
                context=request.additional_context or ""
            )
        elif request.content_type == "presentation_outline":
            content = ai_agent.generate_presentation_outline(
                topic=request.topic,
                slide_count=request.slide_count or 5,
                context=request.additional_context or ""
            )
        else:
            raise HTTPException(status_code=400, detail=f"Content type '{request.content_type}' not yet implemented")
        
        return APIResponse(
            success=True,
            message="AI content generated successfully",
            data={
                "topic": request.topic,
                "content_type": request.content_type,
                "generated_content": content.to_dict() if hasattr(content, 'to_dict') else str(content),
                "tone": request.tone
            }
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"AI content generation error: {e}")
        return APIResponse(
            success=False,
            message="Failed to generate AI content",
            error=str(e)
        )


# File upload endpoint
@app.post("/api/upload", response_model=APIResponse)
async def upload_file(file: UploadFile = File(...)):
    """Upload images or documents for use in presentations."""
    try:
        logger.info(f"Uploading file: {file.filename}")
        
        # Validate file type
        allowed_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.svg', '.pdf', '.docx', '.txt'}
        file_ext = Path(file.filename or 'unknown').suffix.lower()
        
        if file_ext not in allowed_extensions:
            raise HTTPException(
                status_code=400, 
                detail=f"File type {file_ext} not allowed. Supported: {', '.join(allowed_extensions)}"
            )
        
        # Create upload directory if it doesn't exist
        upload_dir = project_root / "uploads"
        upload_dir.mkdir(exist_ok=True)
        
        # Generate unique filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{timestamp}_{file.filename}"
        file_path = upload_dir / filename
        
        # Save file
        content = await file.read()
        with open(file_path, "wb") as f:
            f.write(content)
        
        return APIResponse(
            success=True,
            message="File uploaded successfully",
            data={
                "filename": filename,
                "original_filename": file.filename,
                "path": str(file_path),
                "size": len(content),
                "type": file.content_type,
                "url": f"/uploads/{filename}"
            }
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"File upload error: {e}")
        return APIResponse(
            success=False,
            message="Failed to upload file",
            error=str(e)
        )


# List presentations endpoint
@app.get("/api/presentations", response_model=APIResponse)
async def list_presentations():
    """List all generated presentations."""
    try:
        outputs_dir = project_root / "outputs"
        presentations = []
        
        if outputs_dir.exists():
            for ppt_file in outputs_dir.glob("*.pptx"):
                stat = ppt_file.stat()
                presentations.append({
                    "name": ppt_file.name,
                    "path": str(ppt_file),
                    "size": stat.st_size,
                    "created": datetime.fromtimestamp(stat.st_ctime).isoformat(),
                    "modified": datetime.fromtimestamp(stat.st_mtime).isoformat(),
                    "url": f"/outputs/{ppt_file.name}"
                })
        
        # Sort by creation time (newest first)
        presentations.sort(key=lambda x: x["created"], reverse=True)
        
        return APIResponse(
            success=True,
            message="Presentations listed successfully",
            data={
                "presentations": presentations,
                "total_count": len(presentations)
            }
        )
        
    except Exception as e:
        logger.error(f"List presentations error: {e}")
        return APIResponse(
            success=False,
            message="Failed to list presentations",
            error=str(e)
        )


# Configuration endpoint
@app.get("/api/config", response_model=APIResponse)
async def get_configuration():
    """Get current system configuration."""
    try:
        config_data = {
            "voice_recognition_engine": getattr(config, 'voice_recognition_engine', 'google'),
            "voice_language": getattr(config, 'voice_language', 'en-US'),
            "enable_voice_feedback": getattr(config, 'enable_voice_feedback', True),
            "log_level": getattr(config, 'log_level', 'INFO'),
            "ai_available": bool(getattr(config, 'gemini_api_key', '')),
            "ai_temperature": getattr(config, 'ai_temperature', 0.7),
            "ai_max_tokens": getattr(config, 'ai_max_tokens', 1000),
            "auto_save": getattr(config, 'auto_save', True),
            "max_slide_count": getattr(config, 'max_slide_count', 50)
        }
        
        return APIResponse(
            success=True,
            message="Configuration retrieved successfully",
            data=config_data
        )
        
    except Exception as e:
        logger.error(f"Configuration retrieval error: {e}")
        return APIResponse(
            success=False,
            message="Failed to retrieve configuration",
            error=str(e)
        )


# System information endpoint
@app.get("/api/system-info", response_model=APIResponse)
async def get_system_info():
    """Get system information and statistics."""
    try:
        try:
            import psutil
            import platform
            
            # Get disk usage for outputs directory
            outputs_dir = project_root / "outputs"
            disk_usage = psutil.disk_usage(str(outputs_dir.parent))
            
            system_info = {
                "platform": platform.system(),
                "python_version": platform.python_version(),
                "cpu_count": psutil.cpu_count(),
                "memory_total_gb": round(psutil.virtual_memory().total / (1024**3), 2),
                "disk_usage": {
                    "total_gb": round(disk_usage.total / (1024**3), 2),
                    "used_gb": round(disk_usage.used / (1024**3), 2),
                    "free_gb": round(disk_usage.free / (1024**3), 2)
                }
            }
        except ImportError:
            import platform
            system_info = {
                "platform": platform.system(),
                "python_version": platform.python_version(),
                "psutil_available": False
            }
        
        # Add directory info
        outputs_dir = project_root / "outputs"
        uploads_dir = project_root / "uploads"
        system_info["directories"] = {
            "outputs": len(list(outputs_dir.glob("*.pptx"))) if outputs_dir.exists() else 0,
            "uploads": len(list(uploads_dir.iterdir())) if uploads_dir.exists() else 0
        }
        
        return APIResponse(
            success=True,
            message="System information retrieved successfully",
            data=system_info
        )
        
    except Exception as e:
        logger.error(f"System info error: {e}")
        return APIResponse(
            success=False,
            message="Failed to retrieve system information",
            error=str(e)
        )


# Error handlers
@app.exception_handler(404)
async def not_found_handler(request, exc):
    return JSONResponse(
        status_code=404,
        content=APIResponse(
            success=False,
            message="Endpoint not found",
            error="The requested resource was not found"
        ).dict()
    )


@app.exception_handler(500)
async def internal_error_handler(request, exc):
    logger.error(f"Internal server error: {exc}")
    return JSONResponse(
        status_code=500,
        content=APIResponse(
            success=False,
            message="Internal server error",
            error="An unexpected error occurred"
        ).dict()
    )


# Startup event
@app.on_event("startup")
async def startup_event():
    """Initialize system on startup."""
    logger.info("AI-PPT Agent starting up...")
    
    # Create necessary directories
    directories = ["outputs", "uploads", "logs"]
    for dir_name in directories:
        dir_path = project_root / dir_name
        dir_path.mkdir(exist_ok=True)
        logger.info(f"Ensured directory exists: {dir_path}")
    
    # Log component status
    logger.info(f"Components initialized:")
    logger.info(f"  - PPT Generator: {'✓' if ppt_generator else '✗'}")
    logger.info(f"  - Voice Processor: {'✓' if voice_processor else '✗'}")
    logger.info(f"  - Command Parser: {'✓' if command_parser else '✗'}")
    logger.info(f"  - AI Agent: {'✓' if ai_agent else '✗'}")
    
    logger.info("AI-PPT Agent startup completed successfully")


# Shutdown event
@app.on_event("shutdown")
async def shutdown_event():
    """Cleanup on shutdown."""
    logger.info("AI-PPT Agent shutting down...")
    
    # Stop any continuous voice processing
    if hasattr(voice_processor, 'stop_continuous_listening'):
        voice_processor.stop_continuous_listening()
    
    logger.info("AI-PPT Agent shutdown completed")


# Main entry point
if __name__ == "__main__":
    logger.info("Starting AI-PPT Agent server...")
    
    try:
        uvicorn.run(
            "app:app",
            host="0.0.0.0",
            port=8000,
            reload=True,
            log_level="info",
            access_log=True
        )
    except Exception as e:
        print(f"❌ Failed to start server: {e}")
        print("💡 Make sure you have installed: pip install fastapi uvicorn")
        sys.exit(1)


# PPT action execution endpoint
@app.post("/api/ppt-action", response_model=APIResponse)
async def execute_ppt_action(request: PPTActionRequest):
    """Execute specific PowerPoint manipulation action."""
    try:
        logger.info(f"Executing PPT action: {request.action}")
        
        # Load presentation if path provided
        if request.presentation_path and os.path.exists(request.presentation_path):
            ppt_generator.load_presentation(request.presentation_path)
        
        # Execute the action based on type
        result = None
        
        if request.action == "add_slide":
            title = request.parameters.get("title", "New Slide")
            content = request.parameters.get("content", "")
            layout_index = request.parameters.get("layout", 1)
            result = ppt_generator.add_slide(title, content, layout_index)
            
        elif request.action == "delete_slide":
            slide_number = request.parameters.get("slide_number")
            if slide_number is None:
                raise HTTPException(status_code=400, detail="slide_number parameter required")
            result = ppt_generator.delete_slide(slide_number)
            
        elif request.action == "modify_layout":
            slide_number = request.parameters.get("slide_number")
            new_layout = request.parameters.get("layout")
            if slide_number is None or new_layout is None:
                raise HTTPException(status_code=400, detail="slide_number and layout parameters required")
            result = ppt_generator.modify_layout(slide_number, new_layout)
            
        elif request.action == "insert_image":
            slide_number = request.parameters.get("slide_number")
            image_path = request.parameters.get("image_path")
            if slide_number is None or image_path is None:
                raise HTTPException(status_code=400, detail="slide_number and image_path parameters required")
            result = ppt_generator.insert_image(slide_number, image_path)
            
        elif request.action == "insert_chart":
            slide_number = request.parameters.get("slide_number")
            chart_data = request.parameters.get("chart_data")
            chart_type = request.parameters.get("chart_type", "column")
            if slide_number is None or chart_data is None:
                raise HTTPException(status_code=400, detail="slide_number and chart_data parameters required")
            result = ppt_generator.insert_chart(slide_number, chart_data, chart_type)
            
        elif request.action == "update_text":
            slide_number = request.parameters.get("slide_number")
            new_text = request.parameters.get("text")
            if slide_number is None or new_text is None:
                raise HTTPException(status_code=400, detail="slide_number and text parameters required")
            result = ppt_generator.update_text(slide_number, new_text)
            
        elif request.action == "change_background":
            slide_number = request.parameters.get("slide_number")
            background_color = request.parameters.get("color")
            if slide_number is None or background_color is None:
                raise HTTPException(status_code=400, detail="slide_number and color parameters required")
            result = ppt_generator.change_background(slide_number, background_color)
            
        else:
            raise HTTPException(status_code=400, detail=f"Unsupported action: {request.action}")
        
        # Auto-save if requested
        if request.auto_save:
            save_path = project_root / "outputs" / f"presentation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
            ppt_generator.save_presentation(str(save_path))
            result = result or {}
            result["saved_to"] = str(save_path)
        
        return APIResponse(
            success=True,
            message=f"Action '{request.action}' executed successfully",
            data={
                "action": request.action,
                "parameters": request.parameters,
                "result": result
            }
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"PPT action execution error: {e}")
        return APIResponse(
            success=False,
            message=f"Failed to execute action '{request.action}'",
            error=str(e)
        )


# AI content generation endpoint
@app.post("/api/ai-content", response_model=APIResponse)
async def generate_ai_content(request: AIContentRequest):
    """Generate AI-powered content for presentations."""
    try:
        logger.info(f"Generating AI content for topic: {request.topic}")
        
        if not ai_agent:
            raise HTTPException(status_code=503, detail="AI agent not available")
        
        # Generate content based on type
        content = None
        
        if request.content_type == "slide_content":
            content = ai_agent.generate_slide_content(
                topic=request.topic,
                context=request.additional_context,
                tone=request.tone
            )
        elif request.content_type == "presentation_outline":
            content = ai_agent.generate_presentation_outline(
                topic=request.topic,
                slide_count=request.slide_count or 5,
                context=request.additional_context
            )
        elif request.content_type == "chart_data":
            content = ai_agent.suggest_chart_data(
                topic=request.topic,
                context=request.additional_context
            )
        else:
            raise HTTPException(status_code=400, detail=f"Unsupported content type: {request.content_type}")
        
        return APIResponse(
            success=True,
            message="AI content generated successfully",
            data={
                "topic": request.topic,
                "content_type": request.content_type,
                "generated_content": content.to_dict() if hasattr(content, 'to_dict') else content,
                "tone": request.tone
            }
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"AI content generation error: {e}")
        return APIResponse(
            success=False,
            message="Failed to generate AI content",
            error=str(e)
        )


# File upload endpoint
@app.post("/api/upload", response_model=APIResponse)
async def upload_file(file: UploadFile = File(...)):
    """Upload images or documents for use in presentations."""
    try:
        logger.info(f"Uploading file: {file.filename}")
        
        # Validate file type
        allowed_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.svg', '.pdf', '.docx', '.txt'}
        file_ext = Path(file.filename).suffix.lower()
        
        if file_ext not in allowed_extensions:
            raise HTTPException(
                status_code=400, 
                detail=f"File type {file_ext} not allowed. Supported: {', '.join(allowed_extensions)}"
            )
        
        # Create upload directory if it doesn't exist
        upload_dir = project_root / "uploads"
        upload_dir.mkdir(exist_ok=True)
        
        # Generate unique filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{timestamp}_{file.filename}"
        file_path = upload_dir / filename
        
        # Save file
        content = await file.read()
        with open(file_path, "wb") as f:
            f.write(content)
        
        return APIResponse(
            success=True,
            message="File uploaded successfully",
            data={
                "filename": filename,
                "original_filename": file.filename,
                "path": str(file_path),
                "size": len(content),
                "type": file.content_type,
                "url": f"/uploads/{filename}"
            }
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"File upload error: {e}")
        return APIResponse(
            success=False,
            message="Failed to upload file",
            error=str(e)
        )


# List presentations endpoint
@app.get("/api/presentations", response_model=APIResponse)
async def list_presentations():
    """List all generated presentations."""
    try:
        outputs_dir = project_root / "outputs"
        presentations = []
        
        if outputs_dir.exists():
            for ppt_file in outputs_dir.glob("*.pptx"):
                stat = ppt_file.stat()
                presentations.append({
                    "name": ppt_file.name,
                    "path": str(ppt_file),
                    "size": stat.st_size,
                    "created": datetime.fromtimestamp(stat.st_ctime).isoformat(),
                    "modified": datetime.fromtimestamp(stat.st_mtime).isoformat(),
                    "url": f"/outputs/{ppt_file.name}"
                })
        
        # Sort by creation time (newest first)
        presentations.sort(key=lambda x: x["created"], reverse=True)
        
        return APIResponse(
            success=True,
            message="Presentations listed successfully",
            data={
                "presentations": presentations,
                "total_count": len(presentations)
            }
        )
        
    except Exception as e:
        logger.error(f"List presentations error: {e}")
        return APIResponse(
            success=False,
            message="Failed to list presentations",
            error=str(e)
        )


# Download presentation endpoint
@app.get("/api/presentations/{filename}")
async def download_presentation(filename: str):
    """Download a specific presentation file."""
    try:
        file_path = project_root / "outputs" / filename
        
        if not file_path.exists():
            raise HTTPException(status_code=404, detail="Presentation not found")
        
        if not filename.endswith('.pptx'):
            raise HTTPException(status_code=400, detail="Invalid file type")
        
        return FileResponse(
            path=str(file_path),
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Download presentation error: {e}")
        raise HTTPException(status_code=500, detail=str(e))


# Delete presentation endpoint
@app.delete("/api/presentations/{filename}", response_model=APIResponse)
async def delete_presentation(filename: str):
    """Delete a specific presentation file."""
    try:
        file_path = project_root / "outputs" / filename
        
        if not file_path.exists():
            raise HTTPException(status_code=404, detail="Presentation not found")
        
        if not filename.endswith('.pptx'):
            raise HTTPException(status_code=400, detail="Invalid file type")
        
        file_path.unlink()  # Delete the file
        
        return APIResponse(
            success=True,
            message="Presentation deleted successfully",
            data={"filename": filename}
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Delete presentation error: {e}")
        return APIResponse(
            success=False,
            message="Failed to delete presentation",
            error=str(e)
        )


# Voice test endpoint
@app.post("/api/voice-test", response_model=APIResponse)
async def test_voice_recognition():
    """Test voice recognition capabilities."""
    try:
        if not voice_processor:
            raise HTTPException(status_code=503, detail="Voice processor not available")
        
        test_results = voice_processor.test_speech_recognition()
        
        return APIResponse(
            success=test_results.get("test_recognition", False),
            message="Voice recognition test completed",
            data=test_results
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Voice test error: {e}")
        return APIResponse(
            success=False,
            message="Voice recognition test failed",
            error=str(e)
        )


# Configuration endpoint
@app.get("/api/config", response_model=APIResponse)
async def get_configuration():
    """Get current system configuration."""
    try:
        config_data = {
            "voice_recognition_engine": config.voice_recognition_engine,
            "voice_language": config.voice_language,
            "enable_voice_feedback": config.enable_voice_feedback,
            "log_level": config.log_level,
            "ai_available": bool(config.gemini_api_key),
            "ai_temperature": config.ai_temperature,
            "ai_max_tokens": config.ai_max_tokens,
            "auto_save": config.auto_save,
            "max_slide_count": config.max_slide_count
        }
        
        return APIResponse(
            success=True,
            message="Configuration retrieved successfully",
            data=config_data
        )
        
    except Exception as e:
        logger.error(f"Configuration retrieval error: {e}")
        return APIResponse(
            success=False,
            message="Failed to retrieve configuration",
            error=str(e)
        )


# System information endpoint
@app.get("/api/system-info", response_model=APIResponse)
async def get_system_info():
    """Get system information and statistics."""
    try:
        import psutil
        import platform
        
        # Get disk usage for outputs directory
        outputs_dir = project_root / "outputs"
        disk_usage = psutil.disk_usage(str(outputs_dir.parent))
        
        system_info = {
            "platform": platform.system(),
            "python_version": platform.python_version(),
            "cpu_count": psutil.cpu_count(),
            "memory_total_gb": round(psutil.virtual_memory().total / (1024**3), 2),
            "disk_usage": {
                "total_gb": round(disk_usage.total / (1024**3), 2),
                "used_gb": round(disk_usage.used / (1024**3), 2),
                "free_gb": round(disk_usage.free / (1024**3), 2)
            },
            "directories": {
                "outputs": len(list(outputs_dir.glob("*.pptx"))) if outputs_dir.exists() else 0,
                "uploads": len(list((project_root / "uploads").iterdir())) if (project_root / "uploads").exists() else 0
            }
        }
        
        return APIResponse(
            success=True,
            message="System information retrieved successfully",
            data=system_info
        )
        
    except Exception as e:
        logger.error(f"System info error: {e}")
        return APIResponse(
            success=False,
            message="Failed to retrieve system information",
            error=str(e)
        )


# Error handlers
@app.exception_handler(404)
async def not_found_handler(request, exc):
    return JSONResponse(
        status_code=404,
        content=APIResponse(
            success=False,
            message="Endpoint not found",
            error="The requested resource was not found"
        ).dict()
    )


@app.exception_handler(500)
async def internal_error_handler(request, exc):
    logger.error(f"Internal server error: {exc}")
    return JSONResponse(
        status_code=500,
        content=APIResponse(
            success=False,
            message="Internal server error",
            error="An unexpected error occurred"
        ).dict()
    )


# Startup event
@app.on_event("startup")
async def startup_event():
    """Initialize system on startup."""
    logger.info("AI-PPT Agent starting up...")
    
    # Create necessary directories
    directories = ["outputs", "uploads", "logs"]
    for dir_name in directories:
        dir_path = project_root / dir_name
        dir_path.mkdir(exist_ok=True)
        logger.info(f"Ensured directory exists: {dir_path}")
    
    logger.info("AI-PPT Agent startup completed successfully")


# Shutdown event
@app.on_event("shutdown")
async def shutdown_event():
    """Cleanup on shutdown."""
    logger.info("AI-PPT Agent shutting down...")
    
    # Stop any continuous voice processing
    if hasattr(voice_processor, 'stop_continuous_listening'):
        voice_processor.stop_continuous_listening()
    
    logger.info("AI-PPT Agent shutdown completed")


# Main entry point
if __name__ == "__main__":
    logger.info("Starting AI-PPT Agent server...")
    
    uvicorn.run(
        "app:app",
        host="0.0.0.0",
        port=8000,
        reload=True,
        log_level="info",
        access_log=True
    )