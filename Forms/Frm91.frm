VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm91 
   Caption         =   "Image"
   ClientHeight    =   12930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23760
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
   Icon            =   "Frm91.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Frm91.frx":0ECA
   ScaleHeight     =   12930
   ScaleWidth      =   23760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Tmr1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton CMD1 
      BackColor       =   &H000080FF&
      Caption         =   "Kembali Ke Menu Sebelum"
      Height          =   360
      Left            =   21120
      MaskColor       =   &H00400000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   11520
      Width           =   2475
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   0
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kedai Emas Sri Harmoni"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   4680
      TabIndex        =   10
      Top             =   480
      Width           =   14295
   End
   Begin VB.Label Label67 
      BackColor       =   &H00000000&
      Caption         =   "Label7"
      ForeColor       =   &H00FFFFFF&
      Height          =   45
      Left            =   0
      TabIndex        =   9
      Top             =   2040
      Width           =   24000
   End
   Begin VB.Label L1_Text 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "88/88/8888"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   21720
      TabIndex        =   7
      Top             =   1320
      Width           =   2100
   End
   Begin VB.Label L2_Text 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "88/88/8888"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   21705
      TabIndex        =   6
      Top             =   1635
      Width           =   2100
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
      Left            =   16320
      TabIndex        =   5
      Top             =   12480
      Width           =   4905
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   3
      Height          =   735
      Left            =   21360
      Shape           =   4  'Rounded Rectangle
      Top             =   12000
      Width           =   2295
   End
   Begin VB.Label Label36 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Sankyu System"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   21480
      TabIndex        =   4
      Top             =   12360
      Width           =   2055
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
      Left            =   21480
      TabIndex        =   3
      Top             =   12120
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   10215
      Left            =   120
      ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
      Top             =   2400
      Width           =   23415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sila double click di bawah atau pada image untuk upload atau memadamkan image."
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   14400
   End
   Begin VB.Label L3_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L3_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   22560
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   945
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
      Left            =   21360
      TabIndex        =   8
      Top             =   12000
      Width           =   2295
   End
   Begin VB.Label Label85 
      BackColor       =   &H00FFFFFF&
      Height          =   2070
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   24000
   End
End
Attribute VB_Name = "Frm91"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD1_Click()
'on error resume next
Frm85.Show
Unload Frm91
End Sub

Private Sub Tmr1_Timer()
'on error resume next
Frm91.L1_Text = DateTime.Date
Frm91.L2_Text = DateTime.Time$
End Sub
