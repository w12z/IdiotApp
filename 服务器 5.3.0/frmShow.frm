VERSION 5.00
Begin VB.Form frmShow 
   Caption         =   "���"
   ClientHeight    =   720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3285
   LinkTopic       =   "Form2"
   ScaleHeight     =   720
   ScaleWidth      =   3285
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton reload 
      Caption         =   "�ָ�����"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub reload_Click()
Form1.Show
Unload Me
End Sub
