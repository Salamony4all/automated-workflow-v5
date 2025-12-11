"""
Vercel serverless function entry point for Flask application
"""
import sys
import os

# Add the parent directory to the path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Import Flask app
from app import app

# Export for Vercel
app = app

