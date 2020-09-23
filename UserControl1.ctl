VERSION 5.00
Begin VB.UserControl Button 
   ClientHeight    =   1290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4065
   ScaleHeight     =   1290
   ScaleWidth      =   4065
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      Picture         =   "UserControl1.ctx":0000
      ScaleHeight     =   390
      ScaleWidth      =   2880
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      Picture         =   "UserControl1.ctx":3AC2
      ScaleHeight     =   390
      ScaleWidth      =   2880
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      Picture         =   "UserControl1.ctx":7584
      ScaleHeight     =   390
      ScaleWidth      =   2880
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   25
      Width           =   2895
   End
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event Click(Button As Integer)
Public Event MouseMove()

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label1.ForeColor <> vbWhite Then Label1.ForeColor = vbWhite
UserControl.Picture = Picture2.Picture
DoEvents
Play 2
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 0 Then
If Label1.ForeColor <> vbWhite Then
Label1.ForeColor = vbWhite
If UserControl.Picture <> Picture2.Picture Then Play 1
End If
If UserControl.Picture <> Picture3.Picture Then UserControl.Picture = Picture3.Picture
RaiseEvent MouseMove
End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label1.ForeColor <> vbBlack Then Label1.ForeColor = vbBlack
UserControl.Picture = Picture1.Picture
RaiseEvent Click(Button)
End Sub

Function Off()
If Label1.ForeColor <> vbBlack Then Label1.ForeColor = vbBlack
If UserControl.Picture <> Picture1.Picture Then UserControl.Picture = Picture1.Picture
End Function

Private Sub UserControl_Initialize()
If Label1.ForeColor <> vbBlack Then Label1.ForeColor = vbBlack
If UserControl.Picture <> Picture3.Picture Then UserControl.Picture = Picture3.Picture
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label1.ForeColor <> vbWhite Then Label1.ForeColor = vbWhite
UserControl.Picture = Picture2.Picture
DoEvents
Play 2
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 0 Then
If Label1.ForeColor <> vbWhite Then
Label1.ForeColor = vbWhite
If UserControl.Picture <> Picture2.Picture Then Play 1
End If
If UserControl.Picture <> Picture3.Picture Then UserControl.Picture = Picture3.Picture
RaiseEvent MouseMove
End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label1.ForeColor <> vbBlack Then Label1.ForeColor = vbBlack
UserControl.Picture = Picture1.Picture
RaiseEvent Click(Button)
End Sub

Private Sub UserControl_Resize()
UserControl.Width = Picture1.Width
UserControl.Height = Picture1.Height
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Label1.Caption = PropBag.ReadProperty("Caption", Nothing)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Caption", Label1.Caption, Nothing)
End Sub

Public Property Get Caption() As String
Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New2 As String)
Label1.Caption = New2
PropertyChanged "Caption"
End Property


