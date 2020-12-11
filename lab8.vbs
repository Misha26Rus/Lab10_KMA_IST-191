' Сценарий WMI, 2 вариант. Работа выполнена Куценко М.А. ИСТ-191
On Error Resume Next
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("D:\WMI.txt")
Set objService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
If Err.Number <> 0 Then
	WScript.Echo Err.Number & ": " & Err.Description
	WScript.Quit
End If
For Each objObject In objService.ExecQuery("SELECT * FROM Win32_CacheMemory")
	info = info & "Processor description " & objProc.Description & chr(10)
    info = info & "Current size of the installed cache memory - "  & objObject.InstalledSize & chr(10)
	info = info & "Type of cache - "  & objObject.CacheType & chr(10) ' Other (1), Unknown (2), Instruction (3)
' Data (4), Unified (5)
	info = info & "Level of the cache - "  & objObject.Level & chr(10)
	info = info & "Max cache size - " &  objObject.MaxCacheSize & chr(10)
Next
For Each objProc In objService.ExecQuery("SELECT * FROM Win32_Processor")
	info = info & "Number of logical processors - " & objProc.NumberOfLogicalProcessors & chr(10)
	info = info & "External clock frequency in MHz - " & objProc.ExtClock & chr(10)
	info = info & "Max Clock Speed in MHz - " & objProc.MaxClockSpeed & chr(10)
	info = info & "L2CacheSize - " & objProc.L2CacheSize & chr(10)
Next
objFile.WriteLine info
WScript.echo info