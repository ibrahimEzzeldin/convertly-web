"""
Production server entry point for Windows (uses waitress).
Linux/cloud deployments use gunicorn via the Procfile.

Usage:
    python run.py
"""
from dotenv import load_dotenv
load_dotenv(override=True)

import os
from waitress import serve
from app import app

host = os.getenv("HOST", "0.0.0.0")
port = int(os.getenv("FLASK_PORT", 5000))

if os.getenv("FLASK_DEBUG", "False").lower() == "true":
    print(f"Starting Convertly in DEBUG mode on http://{host}:{port}")
    app.run(host=host, port=port, debug=True, threaded=True)
else:
    print(f"Starting Convertly on http://{host}:{port}")
    serve(app, host=host, port=port, threads=4)
