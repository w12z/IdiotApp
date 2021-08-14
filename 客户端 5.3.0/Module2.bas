Attribute VB_Name = "Module2"
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Function FindService() As String
On Error Resume Next
frmRegister.Winsock1.LocalPort = 2003
Dim ip, sip As String
Dim i, j, Length As Integer
ip = frmRegister.Winsock1.LocalIP
For i = 1 To Len(ip)
If Mid(ip, i, 1) = "." Then
j = j + 1
End If
sip = sip + Mid(ip, i, 1)
If j = 3 Then
Exit For
End If
Next i
Length = Len(sip)
For i = 2 To 10
sip = Mid(sip, 1, Length) & i
frmRegister.Text1.Text = frmRegister.Text1.Text & vbCrLf + "正在尝试IP:" & sip
frmRegister.Winsock1.RemoteHost = sip
frmRegister.Winsock1.RemotePort = 32767  ' 32767为探测端口
frmRegister.Winsock1.Connect
Sleep 100
If frmRegister.Winsock1.State = 7 Then
frmRegister.Text1.Text = frmRegister.Text1.Text & vbCrLf + "寻找成功"
FindService = sip
Exit Function
Else
frmRegister.Winsock1.Close
End If
Next i
frmRegister.Text1.Text = frmRegister.Text1.Text & vbCrLf + "未找到主机，请重试"
End Function
