Dim shell
Dim fso
Dim scriptRoot
Dim ps1Path
Dim cmd

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptRoot = fso.GetParentFolderName(WScript.ScriptFullName)
ps1Path = fso.BuildPath(scriptRoot, "Start-ReconcileTool.ps1")

cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & ps1Path & """"
shell.Run cmd, 0, False
