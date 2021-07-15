Dim iPPort
Dim iPPort_Len
Dim dns
Dim dName
Dim dIndex_en
Dim dIndex_ch
Dim devicePort
Dim deviceIP
Dim adbPath
Dim resultsMsg
dns = "223.5.5.5"
Dim regEx
Set regEx = New RegExp
Function getDomainIp(dns, domain)
	'Begin lookup domain IP by command nslookup
	Dim objShell, objExecObject, strText
	Set objShell = CreateObject("WScript.Shell")
	Set objExecObject = objShell.Exec("%comspec% /c nslookup " & domain & " " & dns)
	Do While Not objExecObject.StdOut.AtEndOfStream
		strText = objExecObject.StdOut.ReadAll()
	Loop
	'search target IP
	Dim Matches, Match, domainIp
	regEx.Pattern = "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}"
	regEx.IgnoreCase = False
	regEx.Global = True
	Set Matches = regEx.Execute(strText)
	If Matches.Count > 1 Then
		domainIp = Matches.Item(1).Value
	End If
	getDomainIp = domainIp
End Function

Function FindCount(Str,toSearch)
Dim Times,WordLen
Times = 0
WordLen = Len(toSearch)
For i = 1 To Len(Str)+1-WordLen
If Mid(Str,i,WordLen) = toSearch Then
Times = Times + 1
End If
Next
FindCount = Times
end function

iPPort=InputBox("请输入需要连接 IP（域名）+端口号 " & Chr(10) & "示例：192.168.0.1:5555/domain.com:5555" & Chr(10) & "使用默认端口【5555】时可以只输入IP（域名）" & Chr(10) & "Please enter IP (domain name) + port number to connect" & Chr(10) & "EG.：192.168.0.1:5555/domain.com:5555" & Chr(10) & "When using the default port [5555], you can only enter the IP (domain name)","输入ADB连接信息 - Enter ADB connection information")
if iPPort = "" then
wscript.quit
end if

iPPort_Len = len(iPPort)
dIndex_en = instr(iPPort,":")
if dIndex_en =0 then
dIndex_ch = instr(iPPort,"：")

if dIndex_ch = 0 then
dName = iPPort
devicePort = "5555"
else
dName=left(iPPort,dIndex_ch-1)
devicePort = right(iPPort,iPPort_Len-dIndex_ch)
end if

else
dName=left(iPPort,dIndex_en-1)
devicePort = right(iPPort,iPPort_Len-dIndex_en)
end if

if FindCount(dName,".") = 3 then
deviceIP = dName
else
deviceIP = getDomainIp(dns,dName)
end if

adbPath = deviceIP & ":" & devicePort
Set oShell = wscript.createobject("WScript.Shell")
oShell.run "%COMSPEC% /C cd %cd%", 0, false
wscript.sleep 500
oShell.run "%COMSPEC% /C adb kill-server", 0, false
wscript.sleep 1000
oShell.run "%COMSPEC% /C adb start-server", 0, false
wscript.sleep 1000 
Set results = oShell.Exec("%COMSPEC% /C adb connect " & adbPath)
Do While Not results.StdOut.AtEndOfStream
    resultsMsg = results.StdOut.ReadAll()
Loop

if FindCount(resultsMsg,"connected") = 0 then
msgBox resultsMsg,vbExclamation,"连接失败 - The connection fails"
else
yesOrNo = msgbox ("是否立即运行 Scrcpy ？" & Chr(10) & "Do you want to run Scrcpy immediately?",32+4+256,"运行 Scrcpy - Run Scrcpy")
if yesOrNo = 6 Then
CreateObject("Wscript.Shell").Run "StartScrcpy.vbs", 0, false
end if
end if
