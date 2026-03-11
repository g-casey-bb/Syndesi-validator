@echo off
cd /d "%~dp0"
echo Starting Syndesi frontend on http://localhost:4201
npx ng serve --port 4201
