Set objShell = WScript.CreateObject("WScript.Shell")
' Run the batch file completely hidden (0 means hidden window)
objShell.Run "cmd /c start.bat", 0, False
