"""
StimuPop Launcher - Standalone executable entry point.
This script launches the Streamlit app programmatically.
"""

import sys
import os
import time
import threading
import webbrowser

# Determine paths based on whether we're frozen or running as script
if getattr(sys, 'frozen', False):
    # Running as compiled executable - PyInstaller extracts to _MEIPASS
    bundle_dir = sys._MEIPASS
    exe_dir = os.path.dirname(sys.executable)

    # Force Streamlit to use production mode (serve static files, not dev server)
    os.environ['STREAMLIT_SERVER_ENABLE_STATIC_SERVING'] = 'true'
    os.environ['STREAMLIT_GLOBAL_DEVELOPMENT_MODE'] = 'false'
else:
    # Running as script
    bundle_dir = os.path.dirname(os.path.abspath(__file__))
    exe_dir = bundle_dir

# Set working directory to exe location (for output files)
os.chdir(exe_dir)

# Add bundle directory to path (where app.py and src/ are)
sys.path.insert(0, bundle_dir)

# Launch Streamlit
from streamlit.web import cli as stcli


def open_browser_when_ready(url: str, delay: float = 3.0):
    """
    Wait for server startup, then open browser.
    Uses a simple time-based delay to avoid dependency on requests library.
    """
    print("Starting StimuPop server...")
    time.sleep(delay)
    print("Server ready! Opening browser...")
    webbrowser.open(url)


if __name__ == "__main__":
    app_path = os.path.join(bundle_dir, "app.py")

    # Default Streamlit port
    port = 8501
    url = f"http://localhost:{port}"

    # Start browser opener in background thread
    browser_thread = threading.Thread(
        target=open_browser_when_ready,
        args=(url,),
        daemon=True
    )
    browser_thread.start()

    sys.argv = [
        "streamlit",
        "run",
        app_path,
        "--server.headless", "true",
        "--browser.gatherUsageStats", "false",
        "--browser.serverAddress", "localhost",
        "--browser.serverPort", str(port),
        "--global.developmentMode", "false",
    ]
    sys.exit(stcli.main())
