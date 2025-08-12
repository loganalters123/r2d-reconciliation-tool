#!/usr/bin/env python3
"""
Simple launcher for the R2D Reconciliation Streamlit app
Double-click this file to start the web interface
"""

import subprocess
import sys
import os
import webbrowser
import time
from pathlib import Path

def main():
    """Launch the Streamlit app"""
    print("🚀 Starting R2D Reconciliation Tool...")
    print("="*50)
    
    # Get the directory of this script
    script_dir = Path(__file__).parent
    app_file = script_dir / "streamlit_app.py"
    venv_python = script_dir / ".venv" / "bin" / "python"
    
    # Check if virtual environment exists
    if not venv_python.exists():
        print("❌ Virtual environment not found!")
        print("Please run the following commands first:")
        print()
        print("python3 -m venv .venv")
        print("source .venv/bin/activate")
        print("pip install -r requirements.txt")
        print()
        input("Press Enter to exit...")
        return
    
    # Check if streamlit app exists
    if not app_file.exists():
        print(f"❌ Streamlit app not found: {app_file}")
        input("Press Enter to exit...")
        return
    
    try:
        print("📦 Checking dependencies...")
        
        # Check if streamlit is installed
        result = subprocess.run([str(venv_python), "-c", "import streamlit"], 
                              capture_output=True, text=True)
        if result.returncode != 0:
            print("❌ Streamlit not installed!")
            print("Installing dependencies...")
            subprocess.run([str(venv_python), "-m", "pip", "install", "-r", "requirements.txt"], 
                          cwd=script_dir)
        
        print("🌐 Starting web server...")
        
        # Start streamlit
        cmd = [
            str(venv_python), "-m", "streamlit", "run", str(app_file),
            "--browser.gatherUsageStats", "false",
            "--server.headless", "false",
            "--server.port", "8501"
        ]
        
        print("🔗 Opening browser...")
        print("📍 App will be available at: http://localhost:8501")
        print()
        print("💡 To stop the server, press Ctrl+C in the terminal")
        print("="*50)
        
        # Start the process
        subprocess.run(cmd, cwd=script_dir)
        
    except KeyboardInterrupt:
        print("\n👋 Stopping server...")
    except Exception as e:
        print(f"❌ Error starting app: {e}")
        input("Press Enter to exit...")

if __name__ == "__main__":
    main()
import time
import threading

def main():
    print("🚀 Starting R2D Reconciliation Tool...")
    print("=" * 50)
    
    # Change to the script directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    
    print(f"📁 Working directory: {script_dir}")
    
    # Check if streamlit is installed
    try:
        import streamlit
        print("✅ Streamlit is installed")
    except ImportError:
        print("❌ Streamlit not found. Installing...")
        subprocess.run([sys.executable, "-m", "pip", "install", "streamlit"])
        print("✅ Streamlit installed")
    
    # Start browser after a delay
    def open_browser():
        time.sleep(3)
        print("🌐 Opening web browser...")
        webbrowser.open("http://localhost:8501")
    
    browser_thread = threading.Thread(target=open_browser)
    browser_thread.daemon = True
    browser_thread.start()
    
    print("🌐 Starting web interface...")
    print("📱 The app will open in your browser automatically")
    print("🔄 If it doesn't open, go to: http://localhost:8501")
    print()
    print("💡 To stop the app, press Ctrl+C in this window")
    print("=" * 50)
    
    try:
        # Run streamlit
        subprocess.run([
            sys.executable, "-m", "streamlit", "run", "streamlit_app.py",
            "--browser.gatherUsageStats", "false",
            "--server.address", "localhost",
            "--server.port", "8501"
        ], check=True)
    except KeyboardInterrupt:
        print("\n👋 App stopped by user")
    except subprocess.CalledProcessError as e:
        print(f"\n❌ Error running app: {e}")
        input("Press Enter to exit...")
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        input("Press Enter to exit...")

if __name__ == "__main__":
    main()
