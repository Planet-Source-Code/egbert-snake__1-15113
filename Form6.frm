VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FR As Long

Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub Timer1_Timer()
Label1.Caption = FR
FR = 0
End Sub

Function GetData()
FR = FR + 1
SendData 1
End Function

Function SendData(data As Long)
Form4.W1.SendData data
End Function

