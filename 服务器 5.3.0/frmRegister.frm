VERSION 5.00
Begin VB.Form frmRegister 
   Caption         =   "注册"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7620
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   7620
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   1920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "注册"
      Height          =   855
      Left            =   1920
      TabIndex        =   2
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   2400
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wsh As Object
Dim a As String
Private Sub Command1_Click()
If a = Text1.Text Then
MsgBox ("CDKEY 正确!")
wsh.regwrite "HKEY_CURRENT_USER\SOFTWARE\test\isok", 1, "REG_DWORD"
wsh.regwrite "HKEY_CURRENT_USER\SOFTWARE\test\user", InputBox("创建用户名:"), "REG_SZ"
wsh.regwrite "HKEY_CURRENT_USER\SOFTWARE\test\password", InputBox("用户密码"), "REG_SZ"
Load frmLogin
frmLogin.Show
Unload Me
Else
MsgBox ("CDKEY 错误! 请联系管理员")
Text1.Text = a
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Set wsh = CreateObject("Wscript.Shell")
Dim reg As Integer
reg = wsh.regread("HKEY_CURRENT_USER\SOFTWARE\test\isok")
If reg = 1 Then
Load frmLogin
frmLogin.Show
Unload frmRegister
Exit Sub
End If
Label1.Caption = "CDKEY:"
Label1.FontSize = 16
Text1.Text = ""
Text1.FontSize = 13
a = make
End Sub
