@echo off
setlocal
cd /d "%~dp0"
echo [*] Installing Aero Host Suite dependencies...
call npm install --no-audit --no-fund
echo [*] Running self-test suite...
call node tests/run_all_tests.mjs
