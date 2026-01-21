# run.py - SIMPLE SERVER STARTER
import uvicorn

if __name__ == "__main__":
    print("🚀 Starting Report Analyzer AI Server...")
    print("📡 http://localhost:8000")
    print("📚 http://localhost:8000/docs")
    
    uvicorn.run(
        "app_deepseek:app",
        host="0.0.0.0",
        port=8000,
        reload=True,
        log_level="info"
    )