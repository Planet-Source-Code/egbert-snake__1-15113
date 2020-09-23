VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Mode"
   ClientHeight    =   4185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4140
   LinkTopic       =   "Form2"
   Moveable        =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   4140
   StartUpPosition =   2  'CenterScreen
   Begin Snake.Button Button5 
      Height          =   390
      Left            =   600
      TabIndex        =   4
      Top             =   3360
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "Back"
   End
   Begin Snake.Button Button4 
      Height          =   390
      Left            =   600
      TabIndex        =   3
      Top             =   2520
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "NASTY"
   End
   Begin Snake.Button Button3 
      Height          =   390
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "Hard"
   End
   Begin Snake.Button Button2 
      Height          =   390
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "Normal"
   End
   Begin Snake.Button Button1 
      Height          =   390
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "Easy"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button1_Click(Button As Integer)
Form1.SetSpeed 15
Form1.Show
Me.Hide
End Sub

Private Sub Button1_MouseMove()
'Button1.Off
Button2.Off
Button3.Off
Button4.Off
Button5.Off
End Sub

Private Sub Button2_Click(Button As Integer)
Form1.SetSpeed 10
Form1.Show
Me.Hide
End Sub

Private Sub Button2_MouseMove()
Button1.Off
'Button2.Off
Button3.Off
Button4.Off
Button5.Off
End Sub

Private Sub Button3_Click(Button As Integer)
Form1.SetSpeed 5
Form1.Show
Me.Hide
End Sub

Private Sub Button3_MouseMove()
Button1.Off
Button2.Off
'Button3.Off
Button4.Off
Button5.Off
End Sub

Private Sub Button4_Click(Button As Integer)
Form1.SetSpeed 2
Form1.Show
Me.Hide
End Sub

Private Sub Button4_MouseMove()
Button1.Off
Button2.Off
Button3.Off
'Button4.Off
Button5.Off
End Sub

Private Sub Button5_Click(Button As Integer)
Form3.Show
Me.Hide
End Sub

Private Sub Button5_MouseMove()
Button1.Off
Button2.Off
Button3.Off
Button4.Off
'Button5.Off
End Sub

Private Sub Form_Load()
StayOnTop Me
Button1.Off
Button2.Off
Button3.Off
Button4.Off
Button5.Off
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Button1.Off
Button2.Off
Button3.Off
Button4.Off
Button5.Off
End Sub
