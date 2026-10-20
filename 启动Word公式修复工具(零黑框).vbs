Set ws = CreateObject("Wscript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
currentDir = fso.GetParentFolderName(WScript.ScriptFullName)
appExe = currentDir & "\node_modules\electron\dist\electron.exe"
ws.CurrentDirectory = currentDir
ws.Run """" & appExe & """ . --no-sandbox", 0, False
