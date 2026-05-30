@echo off
SET "NODE_PATH=%~dp0node-portable\node-v20.12.2-win-x64"
SET "PATH=%NODE_PATH%;%PATH%"

echo [Node.js Portable] Checking version...
node -v

echo [Node.js Portable] Installing dependencies (please wait)...
call npm install

echo [Node.js Portable] Starting the app...
call npm start

pause
