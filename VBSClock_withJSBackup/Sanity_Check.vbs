Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Get current folder path
CurrentDir = fso.GetAbsolutePathName(".")
FilePath = CurrentDir & "\Test.pptx"

' Wait 5 seconds
WScript.Sleep 5000

' Force launch PowerPoint in full screen
WshShell.Run "powerpnt.exe /s """ & FilePath & """", 1, False