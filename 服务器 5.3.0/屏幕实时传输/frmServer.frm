VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server"
   ClientHeight    =   1365
   ClientLeft      =   5070
   ClientTop       =   6465
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   3150
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tmrKeepSend 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1440
      Top             =   840
   End
   Begin MSWinsockLib.Winsock wsImage 
      Index           =   1
      Left            =   2640
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   12000
   End
   Begin MSWinsockLib.Winsock wsCommand 
      Index           =   1
      Left            =   2040
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   11000
   End
   Begin MSWinsockLib.Winsock wsCommand 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   11000
   End
   Begin MSWinsockLib.Winsock wsCommand 
      Index           =   2
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   11000
   End
   Begin MSWinsockLib.Winsock wsCommand 
      Index           =   3
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   11000
   End
   Begin MSWinsockLib.Winsock wsCommand 
      Index           =   4
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   11000
   End
   Begin MSWinsockLib.Winsock wsCommand 
      Index           =   5
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   11000
   End
   Begin MSWinsockLib.Winsock wsCommand 
      Index           =   6
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   11000
   End
   Begin MSWinsockLib.Winsock wsCommand 
      Index           =   7
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   11000
   End
   Begin MSWinsockLib.Winsock wsCommand 
      Index           =   8
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   11000
   End
   Begin MSWinsockLib.Winsock wsCommand 
      Index           =   9
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   11000
   End
   Begin MSWinsockLib.Winsock wsCommand 
      Index           =   10
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   11000
   End
   Begin MSWinsockLib.Winsock wsImage 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   12000
   End
   Begin MSWinsockLib.Winsock wsImage 
      Index           =   2
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   12000
   End
   Begin MSWinsockLib.Winsock wsImage 
      Index           =   3
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   12000
   End
   Begin MSWinsockLib.Winsock wsImage 
      Index           =   4
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   12000
   End
   Begin MSWinsockLib.Winsock wsImage 
      Index           =   5
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   12000
   End
   Begin MSWinsockLib.Winsock wsImage 
      Index           =   6
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   12000
   End
   Begin MSWinsockLib.Winsock wsImage 
      Index           =   7
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   12000
   End
   Begin MSWinsockLib.Winsock wsImage 
      Index           =   8
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   12000
   End
   Begin MSWinsockLib.Winsock wsImage 
      Index           =   9
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   12000
   End
   Begin MSWinsockLib.Winsock wsImage 
      Index           =   10
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   12000
   End
   Begin VB.Label labState 
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public iQuality As Integer                  '图片质量
Public bSend As Boolean                     '是否可以继续发送图片

Private Type GUID
 Data1 As Long
 Data2 As Integer
 Data3 As Integer
 Data4(0 To 7) As Byte
End Type

Private Type GdiplusStartupInput
 GdiplusVersion As Long
 DebugEventCallback As Long
 SuppressBackgroundThread As Long
 SuppressExternalCodecs As Long
End Type

Private Type EncoderParameter
 GUID As GUID
 NumberOfValues As Long
 type As Long
 Value As Long
End Type

Private Type EncoderParameters
 Count As Long
 Parameter As EncoderParameter
End Type

Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, ByVal outputbuf As Long) As Long
Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hpal As Long, Bitmap As Long) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal FileName As Long, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, id As GUID) As Long
Private Declare Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal FileName As Long, Bitmap As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function AbortDoc Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Function HBmp2JPG(ByVal hBmp As Long, ByVal FileName As String, Optional ByVal quality As Byte = 80) As Boolean
    Dim tSI As GdiplusStartupInput
    Dim lRes As Long
    Dim lGDIP As Long
    Dim lBitmap As Long
    
    '初始化 GDI+
    tSI.GdiplusVersion = 1
    lRes = GdiplusStartup(lGDIP, tSI, 0)
     
    If lRes = 0 Then
        '从句柄创建 GDI+ 图像
        lRes = GdipCreateBitmapFromHBITMAP(hBmp, 0, lBitmap)
        If lRes = 0 Then
            Dim tJpgEncoder As GUID
            Dim tParams As EncoderParameters
            '初始化解码器的GUID标识
            CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
             
            '设置解码器参数
            tParams.Count = 1
            With tParams.Parameter ' Quality
               '得到Quality参数的GUID标识
               CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
               .NumberOfValues = 1
               .type = 4
               .Value = VarPtr(quality)
            End With
            '保存图像
            lRes = GdipSaveImageToFile(lBitmap, StrPtr(FileName), tJpgEncoder, tParams)
            '销毁GDI+图像
            GdipDisposeImage lBitmap
        End If
        '销毁 GDI+
        GdiplusShutdown lGDIP
    End If
    HBmp2JPG = IIf(lRes, False, True)
End Function

Sub CaptureScreen(FileName As String)
    Dim hDC As Long
    Dim hDCmem As Long
    Dim hBmp As Long
    Dim hBmpPrev As Long
    hDC = GetDC(0)
    hDCmem = CreateCompatibleDC(hDC)
    hBmp = CreateCompatibleBitmap(hDC, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY)
    hBmpPrev = SelectObject(hDCmem, hBmp)
    BitBlt hDCmem, 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, hDC, 0, 0, vbSrcCopy
    SelectObject hDCmem, hBmpPrev
    DeleteDC hDCmem
    ReleaseDC 0, hDC
    HBmp2JPG hBmp, FileName, iQuality
    DeleteObject hBmp
End Sub

'Command Port = 11000
'Image Port = 12000

Sub SC(index As Integer, Command As String)                  'Send Command
    Me.wsCommand(index).SendData Command
End Sub

Sub SendImage(index As Integer)                             '发送图片过程
    Dim Temp() As Byte
    bSend = False
    CaptureScreen App.Path + "\TempImage.jpg"
    SC index, "SIZ " & FileLen(App.Path + "\TempImage.jpg")
    Open App.Path + "\TempImage.jpg" For Binary As #1
        ReDim Temp(LOF(1) - 1)
        Get #1, , Temp
    Close #1
    Me.wsImage(index).SendData Temp
End Sub


Private Sub Form_Load()
    iQuality = 1
    iQuailty = Int(InputBox("设置发送的图片的清晰度(1-255,根据网络水平设置)"))
    Me.tmrKeepSend.Interval = 1000 / Int(InputBox("设置发送帧率(1-20帧/秒,根据网络水平设置)"))
    Dim i As Integer
    For i = 1 To 10
    Me.wsCommand(i).LocalPort = 9999 + i
    Me.wsImage(i).LocalPort = 19999 + i
    Me.wsCommand(i).Bind
    Me.wsImage(i).Bind
    Me.wsCommand(i).Listen
    Me.wsImage(i).Listen
    Next i
    Me.labState.Caption = "状态:Listening."
End Sub

Private Sub tmrKeepSend_Timer()
'    If bSend = True Then
        Dim j As Integer
        For j = 1 To 10
        If wsImage(j).State = 7 Then
        SendImage j
        End If
        Next j
'    End If
End Sub

Private Sub wsCommand_ConnectionRequest(index As Integer, ByVal requestID As Long)
    Me.wsCommand(index).Close
    Me.wsCommand(index).Accept requestID
    SC index, "CNT"
End Sub

Private Sub wsCommand_DataArrival(index As Integer, ByVal bytesTotal As Long)
    Dim a As String
    Me.wsCommand(index).GetData a
    Select Case Left(a, 3)
        Case "BEG"                  '接收到 可以开始发送图片 消息
        SC index, "STA"
        Me.tmrKeepSend.Enabled = True
        bSend = True
        
        Case "QUA"                  '接收到 更改图片质量 消息
        iQuality = CDbl(Right(a, Len(a) - 4))
        
    End Select
End Sub

Private Sub wsImage_ConnectionRequest(index As Integer, ByVal requestID As Long)
    Me.wsImage(index).Close
    Me.wsImage(index).Accept requestID
End Sub
