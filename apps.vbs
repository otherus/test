c="Software\Microsoft\Windows\CurrentVersion\Uninstall":a="":i=&H80000002:j=1:k=vbCrLf:Set d=GetObject("winmgmts:\\.\root\default:StdRegProv"):d.EnumKey i,c,h
For Each g In h
d.GetStringValue i,c&"\"&g,"DisplayName",f
If Not IsNull(f)Then a=a&f&k:j=j+1
Next
CreateObject("Wscript.Shell").Popup j&" aplikasi ditemukan"&k&k&a