VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��¼"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   390
      Left            =   480
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      Height          =   390
      Left            =   1920
      TabIndex        =   5
      Top             =   960
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "�û�����(&U):"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "����(&P):"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean
Dim password As String
Dim user As String
Dim tmp As String
Private Sub cmdCancel_Click()
    '����ȫ�ֱ���Ϊ false
    '����ʾʧ�ܵĵ�¼
    LoginSucceeded = False
    Unload Me
    Me.hide
End Sub

Private Sub cmdOK_Click()
    '�����ȷ������
    If txtPassword.Text = password And txtUserName.Text = user Then
        '������������ﴫ��
        '�ɹ��� calling ����
        '����ȫ�ֱ���ʱ�����׵�
        LoginSucceeded = True
        Load Form1
        Form1.Show
        Me.hide
        Unload Me
    Else
        MsgBox "��Ч�����룬������!", , "��¼"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub Form_Load()
Dim wsh As Object
Set wsh = CreateObject("Wscript.Shell")
user = wsh.regread("HKEY_CURRENT_USER\SOFTWARE\test\user")
password = wsh.regread("HKEY_CURRENT_USER\SOFTWARE\test\password")
End Sub
