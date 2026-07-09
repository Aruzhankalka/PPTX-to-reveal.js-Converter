#!/usr/bin/env bash
set -e

if ! command -v node > /dev/null 2>&1; then
  echo "Node.js is not installed. Get Node 18 LTS from nodejs.org and run this again."
  exit 1
fi

NODE_MAJOR=$(node --version | sed 's/^v//' | cut -d. -f1)
if [ "$NODE_MAJOR" -lt 18 ]; then
  echo "Found Node $(node --version), but this project needs Node 18 or later."
  exit 1
fi

echo "Node $(node --version) - ok"
echo "Installing dependencies..."
npm install

read -r -p "Install Playwright browsers too? (only needed for the browser tests) [y/N] " answer
if [ "$answer" = "y" ] || [ "$answer" = "Y" ]; then
  npx playwright install chromium firefox
fi

echo ""
echo "Done. Start the app with: npm run dev"
echo "Then open http://localhost:3000"
echo "Sample file to try: tests/fixtures/sample.pptx"
echo ""

read -r -p "Start it now? [Y/n] " start
if [ "$start" != "n" ] && [ "$start" != "N" ]; then
  npm run dev
fi
