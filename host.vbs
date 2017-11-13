Dim oShell
Set oShell = WScript.CreateObject ("WScript.Shell")
oShell.run '"C:\Program Files (x86)\Python36-32\Scripts\jupyter.exe" notebook', 0
Set oShell = Nothing