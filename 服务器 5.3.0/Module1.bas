Attribute VB_Name = "Module1"
Dim yip As String
Dim i, j, now, strlen As Integer
Dim CDKEY, made As String

Function ip() As String
Dim URL As String, NP, Data, Here As String, Hers As String
On Error Resume Next
URL = "http://members.3322.org/dyndns/getip"                  'ip地址查询网站
Set NP = CreateObject("Microsoft.XMLHTTP")
NP.Open "GET", URL, True
NP.send
Data = StrConv(NP.responseBody, vbUnicode)
ip = Data
End Function

Function make() As String
now = 0
made = ""
yip = ""
yip = ip
strlen = Len(yip)
For i = 1 To strlen  'rand公式:Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
If Mid(yip, i, 1) = "." Then
For j = 1 To 4
made = made & Chr(Int(Rnd(now) * 26) + Asc("A"))
Next j
now = 0
Else
now = now * 10 + Int(Asc(Mid(yip, i, 1))) - 48
End If
Next i
make = made
End Function
Function SendPort(Port As Integer)
Form1.WinsockGet.SendData "Port:" & Port
End Function
