VERSION 5.00
Object = "{58635701-4313-11D1-9D7F-CD6975009A1F}#1.0#0"; "RD.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Launcher"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   DrawWidth       =   10
   FillColor       =   &H000000FF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   48
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   550
   ScaleMode       =   0  'User
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Wurk2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   3480
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   64
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Wurk1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   2400
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   63
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox SPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   3
      Left            =   4440
      Picture         =   "Form1.frx":0442
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox SPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   2
      Left            =   6960
      Picture         =   "Form1.frx":1F3C
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox SPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Index           =   1
      Left            =   3720
      Picture         =   "Form1.frx":3A36
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox SPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Index           =   0
      Left            =   1560
      Picture         =   "Form1.frx":556C
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox SMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   3
      Left            =   5520
      Picture         =   "Form1.frx":70A2
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox SMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   2
      Left            =   8160
      Picture         =   "Form1.frx":8B9C
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox SMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Index           =   1
      Left            =   3720
      Picture         =   "Form1.frx":A696
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox SMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Index           =   0
      Left            =   1560
      Picture         =   "Form1.frx":C1CC
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox HeadPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   6
      Left            =   6960
      Picture         =   "Form1.frx":DD02
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox HeadPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   11
      Left            =   4440
      Picture         =   "Form1.frx":F7FC
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox HeadMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   9
      Left            =   5520
      Picture         =   "Form1.frx":112F6
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox HeadPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   10
      Left            =   4440
      Picture         =   "Form1.frx":12DF0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox HeadPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   9
      Left            =   4440
      Picture         =   "Form1.frx":148EA
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox HeadMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   10
      Left            =   5520
      Picture         =   "Form1.frx":163E4
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox HeadMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   11
      Left            =   5520
      Picture         =   "Form1.frx":17EDE
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox HeadPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   8
      Left            =   6960
      Picture         =   "Form1.frx":199D8
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox HeadMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   6
      Left            =   8160
      Picture         =   "Form1.frx":1B4D2
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox HeadPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   7
      Left            =   6960
      Picture         =   "Form1.frx":1CFCC
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox HeadMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   7
      Left            =   8160
      Picture         =   "Form1.frx":1EAC6
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox HeadMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   8
      Left            =   8160
      Picture         =   "Form1.frx":205C0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox HeadPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Index           =   5
      Left            =   3240
      Picture         =   "Form1.frx":220BA
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox HeadMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Index           =   3
      Left            =   2280
      Picture         =   "Form1.frx":23BF0
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox HeadPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Index           =   4
      Left            =   2760
      Picture         =   "Form1.frx":25726
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox HeadPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Index           =   3
      Left            =   2280
      Picture         =   "Form1.frx":2725C
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox HeadMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Index           =   4
      Left            =   2760
      Picture         =   "Form1.frx":28D92
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox HeadMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Index           =   5
      Left            =   3240
      Picture         =   "Form1.frx":2A8C8
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox KersMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1200
      Picture         =   "Form1.frx":2C3FE
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox Kers1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   360
      Picture         =   "Form1.frx":2CF08
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox HeadMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Index           =   2
      Left            =   1080
      Picture         =   "Form1.frx":2DA12
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox HeadMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Index           =   1
      Left            =   600
      Picture         =   "Form1.frx":2F548
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox HeadPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Index           =   2
      Left            =   1080
      Picture         =   "Form1.frx":3107E
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox HeadPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Index           =   1
      Left            =   600
      Picture         =   "Form1.frx":32BB4
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox B2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   5400
      Picture         =   "Form1.frx":346EA
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox BackGound 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   5400
      ScaleHeight     =   104
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   104
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox B1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Index           =   6
      Left            =   5400
      Picture         =   "Form1.frx":34DE9
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox B1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Index           =   5
      Left            =   5400
      Picture         =   "Form1.frx":35EE4
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox B1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Index           =   4
      Left            =   5400
      Picture         =   "Form1.frx":36807
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox B1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Index           =   3
      Left            =   5400
      Picture         =   "Form1.frx":376F5
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox B1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Index           =   2
      Left            =   5520
      Picture         =   "Form1.frx":37C88
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox B1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Index           =   1
      Left            =   5400
      Picture         =   "Form1.frx":38B61
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.PictureBox B1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Index           =   0
      Left            =   5520
      Picture         =   "Form1.frx":39F9E
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   0
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   7
      Top             =   8280
      Width           =   12000
      Begin REALDIGITSLib.RD Lenght1 
         Height          =   225
         Left            =   360
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2858
         _ExtentY        =   397
         _StockProps     =   0
         Digits          =   "0"
         ThreeDView      =   0   'False
      End
      Begin REALDIGITSLib.RD Time1 
         Height          =   225
         Left            =   2760
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   360
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2858
         _ExtentY        =   397
         _StockProps     =   0
         Digits          =   " 00:00"
         ThreeDView      =   0   'False
      End
      Begin REALDIGITSLib.RD Frames 
         Height          =   225
         Left            =   5160
         TabIndex        =   10
         Top             =   360
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2858
         _ExtentY        =   397
         _StockProps     =   0
         Digits          =   "0"
         ThreeDView      =   0   'False
      End
      Begin REALDIGITSLib.RD Score 
         Height          =   225
         Left            =   7560
         TabIndex        =   9
         Top             =   360
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2858
         _ExtentY        =   397
         _StockProps     =   0
         Digits          =   "0"
         ThreeDView      =   0   'False
      End
      Begin REALDIGITSLib.RD Level1 
         Height          =   225
         Left            =   9960
         TabIndex        =   8
         Top             =   360
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2858
         _ExtentY        =   397
         _StockProps     =   0
         Digits          =   "0"
         ThreeDView      =   0   'False
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   10200
         Top             =   120
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   9960
         TabIndex        =   17
         Top             =   120
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2760
         TabIndex        =   16
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lenght :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   360
         TabIndex        =   15
         Top             =   120
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frame's sec :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   5160
         TabIndex        =   14
         Top             =   120
         Width           =   1560
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Score :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   7560
         TabIndex        =   13
         Top             =   120
         Width           =   840
      End
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   6480
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   120
      Top             =   6960
   End
   Begin VB.PictureBox Apple 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   360
      Picture         =   "Form1.frx":3A6C2
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox AppleMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1200
      Picture         =   "Form1.frx":3B578
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox HeadMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":3C42E
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox HeadPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":3DF64
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox BolMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1200
      Picture         =   "Form1.frx":3FA9A
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox BolPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   360
      Picture         =   "Form1.frx":405A4
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label7 
      Caption         =   "Right :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4440
      TabIndex        =   54
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Left :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6960
      TabIndex        =   53
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Down :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   52
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "UP :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   51
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      Height          =   1035
      Left            =   0
      TabIndex        =   6
      Top             =   6360
      Visible         =   0   'False
      Width           =   3420
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Level As Long
Dim Break As Boolean
Dim Bol(1 To 10000) As Worm
Dim Hoi As Long
Dim Lenght As Long
Dim M As Integer
Dim X4, Y4 As Long
Dim Speed As Long
Dim MoveBaby As Boolean
Dim Height_P As Long
Dim Widht_P As Long
Dim Height_a As Long
Dim Widht_a As Long
Dim Height_b As Long
Dim Widht_b As Long
Dim AddNew As Long
Dim Select1 As Boolean
Dim I As Integer
Dim Dead As Boolean
Dim Pick As Ap
Dim Count1, Count2, Count3 As Long
Dim Pause As Boolean
Dim TimeCount As Long
Dim FR As Long
Dim Next_Level As Boolean
Dim A, B, C, p As Long
Dim Sp As Long
Dim ToetsenBlok, ToetsenBlok2 As Boolean
Dim KJ As Long
Dim UP As Boolean
Dim Kers As Boolean
Dim KersX, KersY As Long
Dim UDLR As Long
Dim Explode As Boolean
Dim R1(1 To 1000) As Smoke
Dim Count4 As Long

Private Type Smoke
X As Long
Y As Long
Speed1 As Long
Speed2 As Long
Range As Long
Color As Integer
Ber As Long
Visible As Boolean
End Type
Private Type Ap
X As Long
Y As Long
End Type

Private Type Worm
X As Long
Y As Long
Up1 As Boolean
Down1 As Boolean
Right1 As Boolean
Left1 As Boolean
End Type

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If ToetsenBlok = True Or ToetsenBlok2 = True Then Exit Sub
Select Case KeyCode
Case vbKeyUp
If Bol(1).Down1 = True Then Exit Sub
UDLR = 0
Bol(1).Left1 = False
Bol(1).Right1 = False
Bol(1).Down1 = False
Bol(1).Up1 = True
ToetsenBlok2 = True
Height_b = 0 - HeadPic(5).ScaleHeight + 30
Widht_b = 0 'HeadPic(5).Width
Case vbKeyDown
If Bol(1).Up1 = True Then Exit Sub
UDLR = 3
Bol(1).Left1 = False
Bol(1).Right1 = False
Bol(1).Up1 = False
Bol(1).Down1 = True
ToetsenBlok2 = True
Height_b = 0
Widht_b = 0
Case vbKeyLeft
If Bol(1).Right1 = True Then Exit Sub
UDLR = 6
Bol(1).Right1 = False
Bol(1).Up1 = False
Bol(1).Down1 = False
Bol(1).Left1 = True
ToetsenBlok2 = True
Height_b = 0 'HeadPic(5).ScaleHeight
Widht_b = 0 - HeadPic(5).ScaleWidth - 15
Case vbKeyRight
If Bol(1).Left1 = True Then Exit Sub
UDLR = 9
Bol(1).Left1 = False
Bol(1).Up1 = False
Bol(1).Down1 = False
Bol(1).Right1 = True
ToetsenBlok2 = True
Height_b = 0
Widht_b = 0
Case vbKeyEscape
DStayOnTop Me
Pause = True
DoEvents
C = MsgBox("Are you sure you want to quit this game?", vbYesNo + vbInformation, "?>>Quit<<?")
If C = vbYes Then Break = True
StayOnTop Me
Pause = False
Case vbKeyPause
If Pause = True Then Pause = False Else Pause = True
End Select
'UDLR
'1 = Up '0
'2 = Down ' 3
'3 = Left ' 6
'4 = Right ' 9
End Sub

Function Main()
MoveBaby = True
Height_P = BolMask.ScaleHeight - 6
Widht_P = BolMask.ScaleWidth - 6
GreatBackGround2
GreatBackGround
Count2 = 1
Bol(1).Down1 = False
Bol(1).Up1 = False
Bol(1).Right1 = True
Bol(1).Left1 = False
Score.Digits = 0
Time1.Digits = "00:00"
Stop_Playing
DStayOnTop Form2
Form2.Hide
StayOnTop Me
Timer1.Enabled = True
'Timer2.Enabled = True
Const ShowOpen = 1
ChDir App.Path
If Not WAVMIX_InitMixer() Then
MsgBox "Unable to Initialize WaveMix DLL" & Chr(13) & "If you insalled all correctly then it will some times help to restart your pc.", vbExclamation + vbOKOnly, "WaveMix Error"
End
End If
Bol(1).X = 10
Bol(1).Y = Me.ScaleHeight / 2
UDLR = 9
Bol(1).Left1 = False
Bol(1).Up1 = False
Bol(1).Down1 = False
Bol(1).Right1 = True
ToetsenBlok2 = True
Height_b = 0
Widht_b = 0
Lenght = 10
Count1 = -50
Level = 0
Break = False
Pause = False
Dead = False
SetFile False
PlayBackGround App.Path & "\Sound\BackGround.wav"
Pick.X = Me.ScaleWidth - 100
Pick.Y = Me.ScaleHeight - 100
Pick.X = Int((Pick.X * Rnd) + 50)
Pick.Y = Int((Pick.Y * Rnd) + 50)
Lenght = 1
Play 5
Lenght1.Digits = 1
Level1.Digits = 1
Do
DoEvents
Cls
Sleep Sp

''Toetsenblok for to fast players
If ToetsenBlok2 = True Then
If Count2 > 6 Then
ToetsenBlok2 = False
Count2 = 0
Else
Count2 = Count2 + 1
End If
End If

''Draw the background
BitBlt Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, BackGound.hdc, 0, 0, SRCCOPY

FR = FR + 1

If MoveBaby Then
TimeCount = TimeCount + 1
Time1.Digits = Format(TimeCount, "00:00")
End If
If Count1 < 31 Then Count1 = Count1 + 1

''Paint the explosion
If Explode = True Then
For I = 1 To 200
If R1(I).Visible Then
SetPixel Me.hdc, R1(I).X, R1(I).Y, vbWhite
R1(I).Y = R1(I).Y - R1(I).Speed1
R1(I).X = R1(I).X - R1(I).Speed2
R1(I).Color = R1(I).Color + R1(I).Ber
If R1(I).Color = 255 Then R1(I).Visible = False
End If
Next I
Count4 = Count4 + 1
If Count4 > 100 Then Explode = False
Else
For I = 1 To 200
SetPixel Me.hdc, -1, -1, vbBlack
Next I
End If

'' Move mouth
If Count3 > 10 Then
If KJ = 0 And UP = True Then
UP = False
End If
If KJ = 2 And UP = False Then
UP = True
End If
If UP = True Then
KJ = KJ - 1
Else
KJ = KJ + 1
End If
Count3 = 0
Else
Count3 = Count3 + 1
End If

''add an new bol
If AddNew <> 0 Then
If ToetsenBlok Then
Bol(1).Left1 = False
Bol(1).Up1 = False
Bol(1).Down1 = False
Bol(1).Right1 = True
ToetsenBlok = False
End If
AddNew = AddNew - 1
Bol(Lenght + 1) = Bol(Lenght)
Score.Digits = Score.Digits + Int((100 * Rnd) + 30)
Lenght = Lenght + 1
End If

'' check out of she hit somthing
If Count1 > 29 And MoveBaby Then
For I = 2 To Lenght
If Bol(1).X + Widht_P - 10 > Bol(I).X And Bol(1).X < Bol(I).X + Widht_P - 10 Then
If Bol(1).Y + Height_P - 10 > Bol(I).Y And Bol(1).Y < Bol(I).Y + Height_P - 10 Then
Dead = True
End If
End If
If Bol(1).X < 0 Or Bol(1).X > Me.ScaleWidth - Widht_P Or Bol(1).Y < 0 Or Bol(1).Y > Me.ScaleHeight - Height_P Then
Dead = True
End If
Next I
End If

''check if she pick's up the apple
If MoveBaby Then
If Bol(1).X + Apple.ScaleWidth > Pick.X And Bol(1).X < Pick.X + Apple.ScaleWidth Then
If Bol(1).Y + Apple.ScaleHeight > Pick.Y And Bol(1).Y < Pick.Y + Apple.ScaleHeight Then
Lenght1.Digits = Lenght
Paint Pick.X - Apple.ScaleWidth / 2, Pick.Y - Apple.ScaleHeight / 2
Play 2
Play 4
AddNew = AddNew + Int((4 * Rnd) + 1)
Pick.X = Me.ScaleWidth - 100
Pick.Y = Me.ScaleHeight - 100
Pick.X = Int((Pick.X * Rnd) + 50)
Pick.Y = Int((Pick.Y * Rnd) + 50)
End If
End If
End If

'' check if she pick's up the kers
If Kers = True And MoveBaby Then
If Bol(1).X + Kers1.ScaleWidth > KersX And Bol(1).X < KersX + Kers1.ScaleWidth Then
If Bol(1).Y + Kers1.ScaleHeight > KersY And Bol(1).Y < KersY + Kers1.ScaleHeight Then
Lenght1.Digits = Lenght
Paint KersX - Kers1.ScaleWidth / 2, KersY - Kers1.ScaleHeight / 2
Play 2
Play 4
AddNew = AddNew + 10
Kers = False
End If
End If
End If

''Check out of he needs to go left, right, up or down
For I = 2 To Lenght
Select1 = True
If Bol(I).Left1 = True Then
GoTo 1
End If
If Bol(I).Down1 = True Then
GoTo 2
End If
If Bol(I).Right1 = True Then
GoTo 3
End If
If Bol(I).Up1 = True Then
GoTo 4
End If

5:
Select1 = False
Bol(I).Left1 = False
Bol(I).Up1 = False
Bol(I).Down1 = False
Bol(I).Right1 = False


2:
If Bol(I).Up1 Or Bol(I).Down1 Then
Widht_a = 0
Else
Widht_a = Widht_P
End If
If Bol(I - 1).Y > Bol(I).Y + 1 + Widht_a Then
Bol(I).Down1 = True
GoTo Nexti
Else
If Select1 = True Then GoTo 5
End If

4:
If Bol(I).Up1 Or Bol(I).Down1 Then
Widht_a = 0
Else
Widht_a = Widht_P
End If
If Bol(I - 1).Y < Bol(I).Y - 1 - Widht_a Then
Bol(I).Up1 = True
GoTo Nexti
Else
If Select1 = True Then GoTo 5
End If

3:
If Bol(I - 1).Up1 Or Bol(I - 1).Down1 Then
Height_a = 0
Else
Height_a = Height_P
End If
If Bol(I - 1).X > Bol(I).X + 1 + Height_a Then
Bol(I).Right1 = True
GoTo Nexti
Else
If Select1 = True Then GoTo 5
End If

1:
If Bol(I - 1).Up1 Or Bol(I - 1).Down1 Then
Height_a = 0
Else
Height_a = Height_P
End If
If Bol(I - 1).X < Bol(I).X - 1 - Height_a Then
Bol(I).Left1 = True
GoTo Nexti
Else
If Select1 = True Then GoTo 5
End If
Nexti:
Next I

'' Move the baby
If MoveBaby Then
For I = 1 To Lenght
If I = 1 Then Speed = 4
If I <> 1 Then Speed = 4
If Bol(I).Left1 = True Then
Bol(I).X = Bol(I).X - Speed
End If
If Bol(I).Right1 = True Then
Bol(I).X = Bol(I).X + Speed
End If
If Bol(I).Up1 = True Then
Bol(I).Y = Bol(I).Y - Speed
End If
If Bol(I).Down1 = True Then
Bol(I).Y = Bol(I).Y + Speed
End If
Next I
End If

''Draw him
For I = 1 To Lenght
If I <> 1 Then
If I = Lenght Then
If MoveBaby = False Then
Bol(I).Right1 = True
Bol(I).Left1 = False
Bol(I).Down1 = False
Bol(I).Up1 = False
End If
If Bol(I).Up1 = True Then p = 0: X4 = 0: Y4 = 0
If Bol(I).Down1 = True Then
p = 1
X4 = 0
Y4 = SMask(p).ScaleHeight / 2
Y4 = Y4 + 6
End If
If Bol(I).Right1 = True Then
p = 3
X4 = SMask(p).ScaleWidth / 2  ''Super bug vb5.0
X4 = X4 + 6
Y4 = 0
End If
If Bol(I).Left1 = True Then p = 2: X4 = 0: Y4 = 0

BitBlt Me.hdc, Bol(I).X - X4, Bol(I).Y - Y4, SMask(p).ScaleWidth, SMask(p).ScaleHeight, SMask(p).hdc, 0, 0, vbSrcAnd
BitBlt Me.hdc, Bol(I).X - X4, Bol(I).Y - Y4, SMask(p).ScaleWidth, SMask(p).ScaleHeight, SPic(p).hdc, 0, 0, vbSrcPaint
Else
BitBlt Me.hdc, Bol(I).X, Bol(I).Y, BolMask.ScaleWidth, BolMask.ScaleHeight, BolMask.hdc, 0, 0, vbSrcAnd
BitBlt Me.hdc, Bol(I).X, Bol(I).Y, BolMask.ScaleWidth, BolMask.ScaleHeight, BolPic.hdc, 0, 0, vbSrcPaint
End If
Else
BitBlt Me.hdc, Bol(I).X + Widht_b, Bol(I).Y + Height_b, HeadMask(KJ + UDLR).ScaleWidth, HeadMask(KJ + UDLR).ScaleHeight, HeadMask(KJ + UDLR).hdc, 0, 0, vbSrcAnd
BitBlt Me.hdc, Bol(I).X + Widht_b, Bol(I).Y + Height_b, HeadMask(KJ + UDLR).ScaleWidth, HeadMask(KJ + UDLR).ScaleHeight, HeadPic(KJ + UDLR).hdc, 0, 0, vbSrcPaint
End If
Next I

''Draw apple
BitBlt Me.hdc, Pick.X, Pick.Y, Apple.ScaleWidth, Apple.ScaleHeight, AppleMask.hdc, 0, 0, vbSrcAnd
BitBlt Me.hdc, Pick.X, Pick.Y, Apple.ScaleWidth, Apple.ScaleHeight, Apple.hdc, 0, 0, vbSrcPaint

'' Draw the kers
If Kers = True Then
BitBlt Me.hdc, KersX, KersY, KersMask.ScaleWidth, KersMask.ScaleHeight, KersMask.hdc, 0, 0, vbSrcAnd
BitBlt Me.hdc, KersX, KersY, KersMask.ScaleWidth, KersMask.ScaleHeight, Kers1.hdc, 0, 0, vbSrcPaint
End If

'' Kers activater
If Int((1000 * Rnd) + 1) = 3 And MoveBaby Then
Kers = True
KersX = Me.ScaleWidth - 100
KersX = Int((KersX * Rnd) + 30)
KersY = Me.ScaleHeight - 100
KersY = Int((KersY * Rnd) + 30)
End If

If Dead = True Then
Play 3
DStayOnTop Me
MsgBox "Game over", vbInformation + vbOKOnly, "Dead"
Break = True
End If

''if you made the level
If Level * 30 < Lenght And MoveBaby Then
If Level <> 0 Then Play 1
ToetsenBlok = True
Count2 = 1
Next_Level = True
End If

''Pause the system
If Pause = True Then
Me.Label4.Caption = "PAUSE"
PainL Me.Label4.Caption, Me.ScaleWidth / 2 - Label4.Width / 2, Me.ScaleHeight / 2 - Label4.Height / 2
Do
DoEvents
Loop Until Pause = False
End If

'' Show new level
If Next_Level = True Then
If MoveBaby Then Level = Level + 1: Next_Level = True
ToetsenBlok = True
MoveBaby = False
Level1.Digits = Level
B = B + 1
A = A + 1
Select Case A
Case Is < 10
Me.Label4.Caption = ">Level " & Level & "<"
PainL Me.Label4.Caption, Me.ScaleWidth / 2 - Label4.Width / 2, Me.ScaleHeight / 2 - Label4.Height / 2
Case Is < 20
Me.Label4.Caption = ">>Level " & Level & "<<"
PainL Me.Label4.Caption, Me.ScaleWidth / 2 - Label4.Width / 2, Me.ScaleHeight / 2 - Label4.Height / 2
Case Is < 30
Me.Label4.Caption = ">>>Level " & Level & "<<<"
PainL Me.Label4.Caption, Me.ScaleWidth / 2 - Label4.Width / 2, Me.ScaleHeight / 2 - Label4.Height / 2
Case Is < 40
Me.Label4.Caption = "Level " & Level & ""
PainL Me.Label4.Caption, Me.ScaleWidth / 2 - Label4.Width / 2, Me.ScaleHeight / 2 - Label4.Height / 2
Case Is < 50
Me.Label4.Caption = ">Level " & Level & "<"
PainL Me.Label4.Caption, Me.ScaleWidth / 2 - Label4.Width / 2, Me.ScaleHeight / 2 - Label4.Height / 2
A = 0
End Select
If B = 230 Then
Call Timer2_Timer
Height_b = 0
Widht_b = 0
Bol(1).X = 10
Bol(1).Y = Me.ScaleHeight / 2
Next_Level = False
AddNew = Lenght
Lenght = 1
Count1 = 0
B = 0
A = 0
UDLR = 9
Lenght1.Digits = Lenght
If Level <> 1 Then GreatBackGround
MoveBaby = True
For I = 1 To 300
Bol(I).Left1 = False
Bol(I).Up1 = False
Bol(I).Down1 = False
Bol(I).Right1 = True
Next I
Play 7
Explode = False
ToetsenBlok = False
End If
End If

''Dead or presst escape
Loop Until Break = True
Stop_Playing
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = True
End Function

Private Sub Form_Unload(Cancel As Integer)
Cancel = -1
DStayOnTop Form1
StayOnTop Form3
Form3.Load
Form3.Show
Me.Hide
End Sub

Private Sub Timer1_Timer()
Frames.Digits = FR
FR = 0
End Sub

Private Sub Timer2_Timer()
SetFile True
End Sub

Function SetSpeed(M As Long)
Me.Cls
ToetsenBlok = True
Sp = M
Me.Show
AddNew = 10
Main
End Function

Private Sub Timer3_Timer()
Timer3.Enabled = False
Unload Me
End Sub

Function CreatPicture(P1 As PictureBox, P2 As PictureBox, Width As Long, Height As Long)
P1.Cls
P2.Cls
Dim D, E
For D = 0 To Height Step P1.ScaleHeight
For E = 0 To Width Step P1.ScaleWidth
BitBlt P2.hdc, E, D, P1.ScaleWidth, P1.ScaleHeight, P1.hdc, 0, 0, SRCCOPY
Next E
Next D
End Function

Function GreatBackGround()
Dim F
F = Int((6 * Rnd) + 0)
BackGound.Width = Me.ScaleWidth
BackGound.Height = Me.ScaleHeight
CreatPicture B1(F), BackGound, Me.ScaleWidth, Me.ScaleHeight
End Function

Function GreatBackGround2()
CreatPicture B2, Picture1, Picture1.ScaleWidth, Picture1.ScaleHeight
End Function

Function PainL(STR As String, X2 As Single, Y2 As Single)
Dim X1, Y1, Co, G
X1 = X2
Y1 = Y2
Co = 0
For G = 1 To 5
Co = Co + 255 / 5
Y1 = Y1 + 1
X1 = X1 + 1
Me.ForeColor = RGB(Co, 0, 0)
Me.CurrentX = X1
Me.CurrentY = Y1
Me.Print STR
Next G
End Function

Function Paint(X3 As Long, Y3 As Long)
Dim M, N, I
For I = 1 To 200
M = Int((50 * Rnd) + X3)
N = Int((50 * Rnd) + Y3)
R1(I).X = M
R1(I).Y = N
R1(I).Range = Int((100 * Rnd) + 20)
R1(I).Color = 100
R1(I).Ber = 255 / R1(I).Range
R1(I).Visible = True
Select Case Int((2 * Rnd) + 1)
Case 1
R1(I).Speed1 = Int((-10 * Rnd) - 1)
Case 2
R1(I).Speed1 = Int((10 * Rnd) + 1)
End Select
Select Case Int((2 * Rnd) + 1)
Case 1
R1(I).Speed2 = Int((-10 * Rnd) - 1)
Case 2
R1(I).Speed2 = Int((10 * Rnd) + 1)
End Select
Next I
Count4 = 0
Explode = True
End Function
