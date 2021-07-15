Dim tcpIPPort
tcpIPPort=InputBox("请输入需要TCPIP监听的端口号 " & Chr(10) & "PS：请先开启USB调试并通过USB线连接设备" & Chr(10) & "默认：5555" & Chr(10) & "Please enter the port number for which you want TCPIP listening" & Chr(10) & "PS: Please start USB debugging and connect the device through USB cable" & Chr(10) & "Default：5555" & Chr(10),"输入TCPID端口信息-Enter TCPID port information")
if tcpIPPort = "" then
tcpIPPort = "5555"
end if
strCommand = "cmd /c adb.exe tcpip " & tcpIPPort
For Each Arg In WScript.Arguments
    strCommand = strCommand & " """ & replace(Arg, """", """""""""") & """"
Next
CreateObject("Wscript.Shell").Run strCommand, 0, false