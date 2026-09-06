' ==============================================================================
' LAUNCHER.VBS - Air-Gapped PowerPoint Stage Timer & Auto-Launcher
' ==============================================================================
Option Explicit

Dim fso, shell, pptApp, presentation
Dim targetTimeStr, targetTime, checkIntervalMs
Dim overrideFile, presentationPath

' ------------------------------------------------------------------------------
' CONFIGURATION
' ------------------------------------------------------------------------------
' Set target time (YYYY-MM-DD HH:MM:SS)
targetTimeStr = "2026-09-06 16:30:00"

' Relative path to PowerPoint deck (or .ppsx)
presentationPath = "presentation.pptx"

' Emergency override trigger file name
overrideFile = "launch.now"

' Loop delay in milliseconds (1000ms = 1 second)
checkIntervalMs = 500

' ------------------------------------------------------------------------------
' INITIALIZATION
' ------------------------------------------------------------------------------
Set fso   = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' Convert string to Date object
targetTime = CDate(targetTimeStr)

' Resolve full absolute path to the presentation file
If Not fso.FileExists(presentationPath) Then
    presentationPath = fso.GetAbsolutePathName(".") & "\" & presentationPath
End If

' ------------------------------------------------------------------------------
' MAIN COUNTDOWN LOOP
' ------------------------------------------------------------------------------
Do While Now < targetTime
    ' Emergency override check: If launch.now exists, break immediately
    If fso.FileExists(overrideFile) Then
        On Error Resume Next
        fso.DeleteFile overrideFile
        On Error GoTo 0
        Exit Do
    End If
    
    WScript.Sleep checkIntervalMs
Loop

' ------------------------------------------------------------------------------
' POWERPOINT FULL-SCREEN EXECUTION (COM)
' ------------------------------------------------------------------------------
On Error Resume Next
Set pptApp = CreateObject("PowerPoint.Application")
If Err.Number <> 0 Then
    ' Fallback to CLI launch if COM instantiation is blocked or fails
    shell.Run "POWERPNT.EXE /S """ & presentationPath & """", 1, False
    WScript.Quit
End If
On Error GoTo 0

' Make PowerPoint window visible
pptApp.Visible = True

' Open presentation and launch full-screen slide show
Set presentation = pptApp.Presentations.Open(presentationPath)
presentation.SlideShowSettings.Run()

' Clean up objects
Set presentation = Nothing
Set pptApp = Nothing
Set fso = Nothing
Set shell = Nothing