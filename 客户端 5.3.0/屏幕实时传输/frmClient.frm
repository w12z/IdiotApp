VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Client"
   ClientHeight    =   11625
   ClientLeft      =   2640
   ClientTop       =   3435
   ClientWidth     =   20505
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   11625
   ScaleWidth      =   20505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
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

Dim Temp() As Byte '���յ��Ķ���������```1
Public iSize As Long 'ͼƬ��С
Public bReceived As Boolean                 'ͼƬ������ɱ��
Public had As Boolean
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub SC(Command As String)                   'Send Command
    Me.wsCommand.SendData Command
End Sub

Private Function Begin()
    Dim a As String
    Dim port As Integer
    a = main.Winsock1.RemoteHostIP
    If Trim(a) <> "" Then
        Me.wsCommand.Connect a, main.Winsock1.RemotePort - 2009 + 10000
        Me.wsImage.Connect a, main.Winsock1.RemotePort - 2009 + 20000
    End If
End Function

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
    HooK
    Me.Left = 0: Me.Top = 0
    Me.Width = Screen.Width: Me.Height = Screen.Height
    Begin
End Sub

Private Sub Form_Unload(Cancel As Integer)
wsImage.Close
wsCommand.Close
UnHooK
End Sub

Private Sub wsCommand_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim a As String
    Me.wsCommand.GetData a
    Select Case Left(a, 3)
        Case "CNT"                  '���յ� ���ӳɹ� ��Ϣ
        Me.Caption = "Client - Connected to server."
        bReceived = False
        SC "BEG"                          '���� ���Կ�ʼ����ͼƬ ��Ϣ
        
        Case "STA" '���յ� ��ʼ����ͼƬ ��Ϣ
        If Dir(App.Path + "TempRecImage.jpg") = "" Then
        Open App.Path + "\TempRecImage.jpg" For Binary As #2        '������ʱ�ļ�
        Me.tmrReSend.Enabled = False
        End If
        Case "SIZ"                  '���յ� ��ǰͼƬ��С ��Ϣ
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

