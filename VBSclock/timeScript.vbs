' Configuration: Set target time and presentation filename
TargetTime = "2026-09-04 14:00:00" 
PresentationFile = "Test.pptx"

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Get absolute path of the current folder
CurrentDir = fso.GetAbsolutePathName(".")
FilePath = CurrentDir & "\" & PresentationFile

Do
    CurrentTime = Now()
    If CDate(CurrentTime) >= CDate(TargetTime) Then
        ' Command line to force maximized presentation mode
        CmdLine = "powerpnt.exe /s """ & FilePath & """"
        WshShell.Run CmdLine, 1, False
        Exit Do
    End If
    WScript.Sleep 1000 ' Poll every 1 second
Loop