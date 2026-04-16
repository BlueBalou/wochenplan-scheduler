import os, subprocess, sys

if __name__ == "__main__":
    port = os.environ.get("PORT", "8501")
    sys.exit(subprocess.call([
        sys.executable, "-m", "streamlit", "run", "streamlit_app.py",
        f"--server.port={port}",
        "--server.address=0.0.0.0",
    ]))
