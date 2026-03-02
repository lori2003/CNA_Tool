Set WshShell = CreateObject("WScript.Shell")
Dim dir : dir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
WshShell.Run "cmd /c """ & dir & "\core\avvia_toolbox.bat""", 0, False
