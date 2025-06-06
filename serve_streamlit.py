import shlex
import subprocess
from pathlib import Path
import modal

# Define the local and remote paths for the Streamlit app script
streamlit_script_local_path = Path(__file__).parent / "app.py"
streamlit_script_remote_path_str = "/root/app.py" # Use string for remote path

# Ensure app.py exists locally before defining the image
if not streamlit_script_local_path.exists():
    raise RuntimeError(
        f"app.py not found at {streamlit_script_local_path}! "
        "Place the app.py script (your Streamlit app) in the same directory as serve_streamlit.py."
    )

# Define container dependencies and add the Streamlit app script to the image
image = (
    modal.Image.debian_slim(python_version="3.11")
    .pip_install(
        "streamlit~=1.45.0",
        "pandas~=2.3.0",
        "openpyxl~=3.1.5",
    )
    .add_local_file(local_path=streamlit_script_local_path, remote_path=streamlit_script_remote_path_str)
)

app = modal.App(name="dci-mbox-converter-webapp", image=image)

# Define the web server function
@app.function() # Removed allow_concurrent_inputs
@modal.concurrent(max_inputs=100) # Added concurrent decorator
@modal.web_server(8000) # Modal will expose this port
def run():
    target = shlex.quote(streamlit_script_remote_path_str)
    
    # Explicitly set server address to 0.0.0.0
    cmd = f"streamlit run {target} --server.port 8000 --server.address=0.0.0.0 --server.headless=true --server.enableCORS=false --server.enableXsrfProtection=false"
    
    subprocess.Popen(cmd, shell=True)

if __name__ == "__main__":
    app.serve()

