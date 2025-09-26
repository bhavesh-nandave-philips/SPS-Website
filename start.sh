#!/bin/bash

echo "Installing Playwright browsers..."
playwright install

echo "Starting Gunicorn server..."
gunicorn --bind 0.0.0.0:$PORT app:app
