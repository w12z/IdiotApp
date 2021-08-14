VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmRegister 
   Caption         =   "Form2"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5940
   LinkTopic       =   "Form2"
   ScaleHeight     =   4755
   ScaleWidth      =   5940
   StartUpPosition =   3  '窗口缺省
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "取消探测"
      Height          =   420
      Left            =   4200
      TabIndex        =   2
      Top             =   4200
      Width           =   1695
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "重试"
      Height          =   1215
      Left            =   1680
      TabIndex        =   1
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   2055
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmRegister.frx":0000
      Top             =   360
      Width           =   4935
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ip As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Cancel_Click()
Me.Hide
Load main
main.Show
Unload Me
End Sub

Private Sub Command1_Click()
FindService
Command1.Caption = "重试"
End Sub

Private Sub Form_Load()
'frmRegister.Hide
Text1.FontSize = 10
Winsock2.Close
Winsock2.LocalPort = 2001
Winsock2.Listen
Winsock3.Close
Winsock3.LocalPort = 2002
Winsock3.Listen
Text1.Text = "请点击下方按钮开始"
Command1.Caption = "寻找"
Command1.FontSize = 20
End Sub

Private Sub Text1_Change()
Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)
Winsock2.Close
Winsock2.Accept (requestID)
Winsock2.SendData "$OK"
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
Dim s As String
MsgBox bytesTotal
Winsock2.GetData s
ip = Winsock2.RemoteHostIP
main.Winsock1.RemoteHost = ip
main.Winsock1.RemotePort = CInt(s)
main.Winsock1.Connect
MsgBox "OK"
Winsock2.Close
End Sub

Private Sub Winsock3_ConnectionRequest(ByVal requestID As Long)
Winsock3.Close
Winsock3.Accept requestID
End Sub

Private Sub Winsock3_DataArrival(ByVal bytesTotal As Long)
Dim s As Integer
Winsock3.GetData s
main.Winsock1.Connect Winsock3.RemoteHost, s
Me.Hide
Unload Me
Load Form1
End Sub

