VERSION 5.00
Begin VB.Form Frm6 
   Caption         =   "Error!"
   ClientHeight    =   10845
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17055
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm6.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10845
   ScaleWidth      =   17055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CMD1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      MouseIcon       =   "Frm6.frx":0ECA
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   10200
      Width           =   2055
   End
   Begin VB.Timer Tmr1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Pic1 
      BorderStyle     =   0  'None
      FontTransparent =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   9255
      Left            =   120
      ScaleHeight     =   9255
      ScaleWidth      =   16815
      TabIndex        =   0
      Top             =   480
      Width           =   16815
   End
   Begin VB.Label Label35 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Powered By"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14640
      TabIndex        =   4
      Top             =   9960
      Width           =   1335
   End
   Begin VB.Label Label36 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Sankyu System"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   14640
      TabIndex        =   3
      Top             =   10200
      Width           =   2055
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   3
      Height          =   735
      Left            =   14520
      Shape           =   4  'Rounded Rectangle
      Top             =   9840
      Width           =   2295
   End
   Begin VB.Label Label44 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "sankyusystem@gmail.com / 010 - 900 4788"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9120
      TabIndex        =   2
      Top             =   10320
      Width           =   5265
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kesilapan pada data seperti di bawah telah dikesan.Sila periksa data anda."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   7815
   End
   Begin VB.Label Label37 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14520
      TabIndex        =   5
      Top             =   9840
      Width           =   2295
   End
End
Attribute VB_Name = "Frm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD1_Click()
'On Error Resume Next
Unload Me
End Sub
Private Sub Form_Load()
'On Error Resume Next
Frm6.Picture = MDI_frm1.Picture
End Sub
Private Sub Tmr1_Timer()
'On Error Resume Next
Frm6.Caption = "Sankyu System     " & DateTime.Date & "     " & DateTime.Time$
End Sub
