@echo off
title Stage Panic Switch - Force Launch
echo Force launching presentation and terminating background timer...

:: 1. Force PowerPoint into full screen using the relative folder path
start "" "powerpnt.exe" /s "%~dp0Test.pptx"

:: 2. Kill any running wscript.exe instances so the VBScript background timer stops polling
taskkill /f /im wscript.exe >nul 2>&1

echo Presentation launched successfully.
timeout /t 2 >nul
exit
