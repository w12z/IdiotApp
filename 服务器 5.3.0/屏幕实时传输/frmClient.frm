VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   AutoRedraw      =   -1  'True
   Caption         =   "Client"
   ClientHeight    =   11625
   ClientLeft      =   2760
   ClientTop       =   3885
   ClientWidth     =   20505
   LinkTopic       =   "Form1"
   ScaleHeight     =   11625
   ScaleWidth      =   20505
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tmrReSend 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6120
      Top             =   4200
   End
   Begin VB.CommandButton cmdQuality 
      Caption         =   "Quality"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect To Server"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
   Begin MSWinsockLib.Winsock wsCommand 
      Left            =   6720
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsImage 
      Left            =   7320
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Command Port = 11000
'Image Port = 12000

Dim Temp() As Byte '接收到的二进制数据```1
Public iSize As Long '图片大小
Public bReceived As Boolean                 '图片接收完成标记
Public had As Boolean
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub SC(Command As String)                   'Send Command
    Me.wsCommand.SendData Command
End Sub

Private Sub cmdConnect_Click()
    Dim a As String
    Dim port As Integer
    a = InputBox("Input server's IP adress:", "Input")
    If Trim(a) <> "" Then
        Me.wsCommand.Connect a, 11000
        Me.wsImage.Connect a, 12000
    End If
End Sub

Private Sub cmdQuality_Click()
    On Error Resume Next
    Dim a As String
    a = InputBox("Input quality(1-255):", "Input")
    If Trim(a) <> "" And a > 0 And a < 256 Then
        Me.wsCommand.SendData "QUA " & a
    End If
End Sub

Private Sub Form_Load()
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
wsImage.Close
wsCommand.Close
End Sub

Private Sub wsCommand_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim a As String
    Me.wsCommand.GetData a
    Select Case Left(a, 3)
        Case "CNT"                  '接收到 连接成功 消息
        Me.Caption = "Client - Connected to server."
        bReceived = False
        SC "BEG"                          '发送 可以开始发送图片 消息
        
        Case "STA" '接收到 开始发送图片 消息
        If Dir(App.Path + "TempRecImage.jpg") = "" Then
        Open App.Path + "\TempRecImage.jpg" For Binary As #2        '创建临时文件
        Me.tmrReSend.Enabled = False
        End If
        Case "SIZ"                  '接收到 当前图片大小 消息
        iSize = CDbl(Right(a, Len(a) - 4))
        
    End Select
End Sub
Private Sub wsImage_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim i As Integer
    Me.wsImage.GetData Temp
    Put #2, LOF(2) + 1, Temp
    If LOF(2) = iSize Then
        had = True
        Close #2
        Set pic = LoadPicture(App.Path + "\TempRecImage.jpg")
        Me.Picture = pic
        had = False
        SC "BEG"
        Me.tmrReSend.Enabled = True
        bReceived = True
         Kill App.Path + "\TempRecImage.jpg"
         pic = Nothing
         Exit Sub
    End If
    If LOF(2) > iSize Then
        Close #2
        Kill App.Path + "\TempRecImage.jpg"
        SC "BEG"
        Exit Sub
    End If
End Sub

