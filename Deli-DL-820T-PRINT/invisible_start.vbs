Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & "e:\github\InterOfHardware\Deli-DL-820T-PRINT\start_print_service_enhanced.bat" & Chr(34), 0, False
Set WshShell = Nothing