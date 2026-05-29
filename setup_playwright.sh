#!/bin/bash

# Streamlit Cloud post-deployment setup script
# This script installs Playwright browsers after dependencies are installed

echo "========================================="
echo "Setting up Playwright browsers..."
echo "========================================="

python -m playwright install chromium

echo "========================================="
echo "✅ Setup complete!"
echo "========================================="
