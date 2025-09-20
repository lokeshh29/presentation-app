"""
Simple test script to verify that the app.py API server works correctly.
Run this after starting the server with 'python app.py'
"""

import requests
import json
from pathlib import Path

# Server configuration
BASE_URL = "http://localhost:8000"

def test_health():
    """Test the health endpoint."""
    print("🔍 Testing health endpoint...")
    try:
        response = requests.get(f"{BASE_URL}/api/health")
        print(f"Status: {response.status_code}")
        print(f"Response: {json.dumps(response.json(), indent=2)}")
        return response.status_code == 200
    except Exception as e:
        print(f"❌ Health test failed: {e}")
        return False

def test_config():
    """Test the configuration endpoint."""
    print("\n🔍 Testing configuration endpoint...")
    try:
        response = requests.get(f"{BASE_URL}/api/config")
        print(f"Status: {response.status_code}")
        print(f"Response: {json.dumps(response.json(), indent=2)}")
        return response.status_code == 200
    except Exception as e:
        print(f"❌ Config test failed: {e}")
        return False

def test_voice_command():
    """Test voice command processing with text input."""
    print("\n🎤 Testing voice command endpoint...")
    try:
        payload = {
            "text_command": "Create a new slide with title 'Test Slide from API'",
            "enhance_with_ai": False,  # Set to False to avoid needing AI
            "session_id": "test_session"
        }
        
        response = requests.post(
            f"{BASE_URL}/api/voice-command",
            json=payload,
            headers={"Content-Type": "application/json"}
        )
        print(f"Status: {response.status_code}")
        print(f"Response: {json.dumps(response.json(), indent=2)}")
        return response.status_code == 200
    except Exception as e:
        print(f"❌ Voice command test failed: {e}")
        return False

def test_ppt_action():
    """Test PPT action execution."""
    print("\n📊 Testing PPT action endpoint...")
    try:
        payload = {
            "action": "add_slide",
            "parameters": {
                "title": "API Test Slide",
                "content": "This slide was created via the API",
                "layout": 1
            },
            "auto_save": True
        }
        
        response = requests.post(
            f"{BASE_URL}/api/ppt-action",
            json=payload,
            headers={"Content-Type": "application/json"}
        )
        print(f"Status: {response.status_code}")
        print(f"Response: {json.dumps(response.json(), indent=2)}")
        return response.status_code == 200
    except Exception as e:
        print(f"❌ PPT action test failed: {e}")
        return False

def test_presentations_list():
    """Test listing presentations."""
    print("\n📋 Testing presentations list endpoint...")
    try:
        response = requests.get(f"{BASE_URL}/api/presentations")
        print(f"Status: {response.status_code}")
        print(f"Response: {json.dumps(response.json(), indent=2)}")
        return response.status_code == 200
    except Exception as e:
        print(f"❌ Presentations list test failed: {e}")
        return False

def test_ai_content():
    """Test AI content generation (will fail if no API key)."""
    print("\n🤖 Testing AI content generation...")
    try:
        payload = {
            "topic": "Artificial Intelligence Basics",
            "content_type": "slide_content",
            "tone": "professional"
        }
        
        response = requests.post(
            f"{BASE_URL}/api/ai-content",
            json=payload,
            headers={"Content-Type": "application/json"}
        )
        print(f"Status: {response.status_code}")
        if response.status_code == 200:
            print(f"Response: {json.dumps(response.json(), indent=2)}")
        else:
            print(f"Response: {response.text}")
        return response.status_code == 200
    except Exception as e:
        print(f"❌ AI content test failed: {e}")
        return False

def main():
    """Run all API tests."""
    print("🚀 Starting API Tests for AI-PPT Agent")
    print("="*50)
    
    # Check if server is running
    try:
        requests.get(f"{BASE_URL}/", timeout=5)
        print("✅ Server is running at", BASE_URL)
    except Exception as e:
        print(f"❌ Server not accessible at {BASE_URL}")
        print("Please start the server first: python app.py")
        return
    
    # Run tests
    tests = [
        ("Health Check", test_health),
        ("Configuration", test_config),
        ("Voice Command", test_voice_command),
        ("PPT Action", test_ppt_action),
        ("Presentations List", test_presentations_list),
        ("AI Content", test_ai_content)
    ]
    
    results = []
    for test_name, test_func in tests:
        try:
            result = test_func()
            results.append((test_name, result))
        except Exception as e:
            print(f"❌ {test_name} test error: {e}")
            results.append((test_name, False))
    
    # Summary
    print("\n" + "="*50)
    print("📊 Test Results Summary:")
    passed = 0
    for test_name, result in results:
        status = "✅ PASS" if result else "❌ FAIL"
        print(f"{status} - {test_name}")
        if result:
            passed += 1
    
    print(f"\n🎯 {passed}/{len(results)} tests passed")
    
    if passed == len(results):
        print("🎉 All tests passed! Your API is working correctly.")
        print(f"📋 Visit {BASE_URL}/api/docs for interactive API documentation")
    else:
        print("⚠️  Some tests failed. Check the logs and configuration.")
        print("💡 Common issues:")
        print("   - Gemini API key not configured (AI tests will fail)")
        print("   - Missing dependencies (run: pip install -r requirements.txt)")
        print("   - Microphone not available (voice tests may fail)")

if __name__ == "__main__":
    main()