VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Playing - Snake"
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3870
   LinkTopic       =   "Form3"
   Moveable        =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Begin Snake.Button Button6 
      Height          =   390
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "MultiPlayer"
   End
   Begin Snake.Button Button5 
      Height          =   390
      Left            =   480
      TabIndex        =   4
      Top             =   3840
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "Exit"
   End
   Begin Snake.Button Button4 
      Height          =   390
      Left            =   480
      TabIndex        =   3
      Top             =   2400
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "Option"
   End
   Begin Snake.Button Button3 
      Height          =   390
      Left            =   480
      TabIndex        =   2
      Top             =   3120
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "About"
   End
   Begin Snake.Button Button2 
      Height          =   390
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "Hi-Score"
   End
   Begin Snake.Button Button1 
      Height          =   390
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "New Game"
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button1_Click(Button As Integer)
DStayOnTop Form3
StayOnTop Form2
Form2.Show
Me.Hide
End Sub

Private Sub Button1_MouseMove()
'Button1.Off
Button2.Off
Button3.Off
Button4.Off
Button5.Off
Button6.Off
End Sub

Private Sub Button2_MouseMove()
Button1.Off
'Button2.Off
Button3.Off
Button4.Off
Button5.Off
Button6.Off
End Sub

Private Sub Button3_MouseMove()
Button1.Off
Button2.Off
'Button3.Off
Button4.Off
Button5.Off
Button6.Off
End Sub

Private Sub Button4_MouseMove()
Button1.Off
Button2.Off
Button3.Off
'Button4.Off
Button5.Off
Button6.Off
End Sub

Private Sub Button5_Click(Button As Integer)
Stop_Playing
DStayOnTop Form3
End
End Sub

Private Sub Button5_MouseMove()
Button1.Off
Button2.Off
Button3.Off
Button4.Off
'Button5.Off
Button6.Off
End Sub

Private Sub Button6_Click(Button As Integer)
Me.Hide
Form4.Show
End Sub

Private Sub Form_Load()
Load
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Button1.Off
Button2.Off
Button3.Off
Button4.Off
Button5.Off
Button6.Off
End Sub

Function Load()
Const ShowOpen = 1
ChDir App.Path
If Not WAVMIX_InitMixer() Then
MsgBox "Unable to Initialize WaveMix DLL" & Chr(13) & "If you insalled all correctly then it will some times help to restart your pc.", vbExclamation + vbOKOnly, "WaveMix Error"
End
End If
SetFile2
PlayBackGround App.Path & "\sound\MenuBackGround.wav"
StayOnTop Me
Button1.Off
Button2.Off
Button3.Off
Button4.Off
Button5.Off
Button6.Off
End Function
