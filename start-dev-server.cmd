@echo off
cd /d "%~dp0"
:: Windows Command Processor Wrapper
powershell -ExecutionPolicy RemoteSigned -Command ".\start-dev-server.ps1 -daemon"