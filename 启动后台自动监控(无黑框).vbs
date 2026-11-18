Set fso = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("WScript.Shell")
rootDir = fso.GetParentFolderName(WScript.ScriptFullName)
batPath = rootDir & "\system\run_engine.bat"
ws.Run Chr(34) & batPath & Chr(34) & " watch", 0, False
