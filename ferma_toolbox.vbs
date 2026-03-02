Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c taskkill /f /im python.exe", 0, True
WScript.Echo "Toolbox fermata."
