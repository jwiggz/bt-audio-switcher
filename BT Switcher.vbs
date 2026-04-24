Dim dir
dir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
CreateObject("WScript.Shell").Run "pythonw.exe """ & dir & "bluetooth_switcher.py""", 0, False
