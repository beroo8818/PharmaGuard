Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

folder = fso.GetParentFolderName(WScript.ScriptFullName)
batFile = folder & "\START_PHARMAGUARD.bat"
logFile = folder & "\startup_log.txt"

shell.CurrentDirectory = folder

If Not fso.FileExists(batFile) Then
    MsgBox "START_PHARMAGUARD.bat was not found in:" & vbCrLf & folder, 16, "PharmaGuard startup error"
    WScript.Quit 1
End If

cmd = "cmd.exe /c """ & batFile & """ > """ & logFile & """ 2>&1"
shell.Run cmd, 0, False
