Dim tcpIPPort
tcpIPPort=InputBox("��������ҪTCPIP�����Ķ˿ں� " & Chr(10) & "PS�����ȿ���USB���Բ�ͨ��USB�������豸" & Chr(10) & "Ĭ�ϣ�5555" & Chr(10) & "Please enter the port number for which you want TCPIP listening" & Chr(10) & "PS: Please start USB debugging and connect the device through USB cable" & Chr(10) & "Default��5555" & Chr(10),"����TCPID�˿���Ϣ-Enter TCPID port information")
if tcpIPPort = "" then
tcpIPPort = "5555"
end if
strCommand = "cmd /c adb.exe tcpip " & tcpIPPort
For Each Arg In WScript.Arguments
    strCommand = strCommand & " """ & replace(Arg, """", """""""""") & """"
Next
CreateObject("Wscript.Shell").Run strCommand, 0, false