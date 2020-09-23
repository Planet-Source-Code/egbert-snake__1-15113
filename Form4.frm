VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "multiplayer"
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6870
   LinkTopic       =   "Form4"
   ScaleHeight     =   2670
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2655
      ScaleWidth      =   6855
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   120
         Top             =   1080
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   120
         Top             =   600
      End
      Begin Snake.Button Button5 
         Height          =   390
         Left            =   1920
         TabIndex        =   10
         Top             =   1680
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   688
         Caption         =   "Cancel"
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   120
         Top             =   120
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Trying to connect..."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   6855
      End
   End
   Begin Snake.Button Button4 
      Height          =   390
      Left            =   3720
      TabIndex        =   7
      Top             =   240
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "Join"
   End
   Begin MSWinsockLib.Winsock W1 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin Snake.Button Button3 
      Height          =   390
      Left            =   3720
      TabIndex        =   6
      Top             =   1440
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "Refresh"
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   1680
      Width           =   2295
   End
   Begin Snake.Button Button2 
      Height          =   390
      Left            =   3720
      TabIndex        =   3
      Top             =   2040
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "Back/Cancel"
   End
   Begin Snake.Button Button1 
      Height          =   390
      Left            =   3720
      TabIndex        =   2
      Top             =   840
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   688
      Caption         =   "Create"
   End
   Begin VB.ComboBox List1 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Hostanme/IP :"
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Founded hostname's    :"
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1425
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Host As Boolean

Private Sub Button1_Click(Button As Integer)
On Error GoTo 1
Label3.Caption = "Creating a game.."
Timer2.Enabled = True
Picture1.Visible = True
W1.LocalPort = 2332
W1.Listen
Timer2.Enabled = True
Timer3.Enabled = True
Host = True
Exit Sub
1:
Button5_Click (0)
MsgBox Error, vbInformation + vbOKOnly, "Error"
End Sub

Private Sub Button1_MouseMove()
'Button1.Off
Button2.Off
Button3.Off
Button4.Off
End Sub

Private Sub Button2_Click(Button As Integer)
Me.Hide
Form3.Show
End Sub

Private Sub Button2_MouseMove()
Button1.Off
'Button2.Off
Button3.Off
Button4.Off
End Sub

Private Sub Button3_Click(Button As Integer)
List1.Clear
Call DoNetEnum(List1)
If List1.ListCount > 0 Then List1.ListIndex = 0
End Sub

Private Sub Button3_MouseMove()
Button1.Off
Button2.Off
'Button3.Off
Button4.Off
End Sub

Private Sub Button4_Click(Button As Integer)
On Error GoTo 1
W1.Close
W1.RemotePort = 2332
W1.RemoteHost = Text1
W1.Connect Text1, 2332
Label3.Caption = "Trying to connect"
Picture1.Visible = True
Timer1.Enabled = True
Host = False
Exit Sub
1:
Button5_Click (0)
MsgBox Error, vbInformation + vbOKOnly, "Error"
End Sub

Private Sub Button4_MouseMove()
Button1.Off
Button2.Off
Button3.Off
'Button4.Off
End Sub

Private Sub Button5_Click(Button As Integer)
W1.Close
Timer1.Enabled = False
Timer2.Enabled = False
Picture1.Visible = False
Me.Enabled = True
Host = False
End Sub

Private Sub Form_Load()
Dim Pos, I As Integer
List1.Clear
Button3_Click (1)
Call DoNetEnum(List1)
If List1.ListCount > 0 Then List1.ListIndex = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Button1.Off
Button2.Off
Button3.Off
Button4.Off
Button5.Off
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = -1
End Sub

Private Sub List1_Click()
Text1.Text = List1.Text
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Button5.Off
End Sub

Private Sub Timer1_Timer()
Host = False
If W1.State = 6 Then Exit Sub
If W1.State = 7 Then
Timer1.Enabled = False
Me.Enabled = True
Picture1.Visible = False
Me.Hide
Form6.Show
Exit Sub
End If
MsgBox "Failt to Join", vbInformation + vbOKOnly, "Failt"
Picture1.Visible = False
Me.Enabled = True
Timer1.Enabled = False
W1.Close
End Sub

Private Sub Timer2_Timer()
If W1.State = 7 Then
Timer2.Enabled = False
Me.Enabled = True
Picture1.Visible = False
Me.Hide
Form5.SetSpeed 10
End If
End Sub

Private Sub Timer3_Timer()
Timer3.Enabled = False
Label3.Caption = "Waiting for other players"
End Sub

Private Sub W1_ConnectionRequest(ByVal requestID As Long)
If W1.State <> sckClosed Then
W1.Close
End If
W1.Accept requestID
End Sub

Private Sub W1_DataArrival(ByVal bytesTotal As Long)
'W1.GetData data
If Host = True Then
Form5.GetData
Else
Form6.GetData
End If
End Sub
