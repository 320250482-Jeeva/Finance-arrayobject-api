
#!/bin/bash
# Install dependencies (optional if already installed)
pip install -r requirements.txt

# Start FastAPI app with Gunicorn and Uvicorn workers
gunicorn -w 4 -k uvicorn.workers.UvicornWorker app:app
