VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm28 
   BackColor       =   &H80000003&
   Caption         =   "Maklumat Pembeli (Pelanggan Berdaftar)"
   ClientHeight    =   11970
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13005
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
   Icon            =   "Frm28.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11970
   ScaleWidth      =   13005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CMD3 
      BackColor       =   &H80000004&
      Caption         =   "Kembali Ke Menu Sebelum"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4920
      MaskColor       =   &H00400000&
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Kembali Ke Menu Sebelum"
      Top             =   11520
      Width           =   2145
   End
   Begin VB.CommandButton CMD2 
      BackColor       =   &H80000004&
      Caption         =   "Padam Maklumat Pembeli"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   8400
      MaskColor       =   &H00400000&
      Picture         =   "Frm28.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Padam maklumat pembeli"
      Top             =   7200
      Width           =   4305
   End
   Begin VB.TextBox TB2 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2040
      TabIndex        =   6
      Text            =   "TB2"
      Top             =   9840
      Width           =   5265
   End
   Begin VB.TextBox TB3 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2040
      TabIndex        =   7
      Text            =   "TB3"
      Top             =   10200
      Width           =   5265
   End
   Begin VB.TextBox TB4 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2040
      TabIndex        =   8
      Text            =   "TB4"
      Top             =   10560
      Width           =   5265
   End
   Begin VB.CommandButton CMD21 
      Caption         =   "Back"
      Height          =   810
      Left            =   10440
      MouseIcon       =   "Frm28.frx":1874
      MousePointer    =   99  'Custom
      Picture         =   "Frm28.frx":1B7E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Tutup senarai ini."
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton CMD22 
      Caption         =   "Next"
      Height          =   810
      Left            =   11640
      MouseIcon       =   "Frm28.frx":2C48
      MousePointer    =   99  'Custom
      Picture         =   "Frm28.frx":2F52
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Tutup senarai ini."
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simpan Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7440
      MouseIcon       =   "Frm28.frx":401C
      MousePointer    =   99  'Custom
      Picture         =   "Frm28.frx":4326
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9840
      Width           =   1815
   End
   Begin VB.CommandButton CMD4 
      BackColor       =   &H80000004&
      Caption         =   "Carian / Pendaftaran"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   18360
      MaskColor       =   &H00400000&
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Carian maklumat pembeli / Pendaftaran maklumat pembeli."
      Top             =   5640
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Timer Tmr1 
      Interval        =   100
      Left            =   12600
      Top             =   0
   End
   Begin VB.CheckBox CB1 
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   13200
      TabIndex        =   17
      Top             =   1440
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.CheckBox CB2 
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   13200
      TabIndex        =   0
      Top             =   1140
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.CheckBox CB3 
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   15840
      TabIndex        =   11
      Top             =   1140
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.TextBox TB1 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2145
      TabIndex        =   1
      Text            =   "TB1"
      Top             =   1560
      Width           =   5700
   End
   Begin VB.CommandButton CMD1 
      BackColor       =   &H80000004&
      Caption         =   "Carian"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   7920
      MaskColor       =   &H00400000&
      Picture         =   "Frm28.frx":68F0
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Carian Maklumat Pembeli"
      Top             =   1440
      Width           =   2145
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   3540
      Left            =   120
      TabIndex        =   21
      Top             =   2520
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   6244
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Paparan Muka  :          / "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   8040
      TabIndex        =   47
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Shape Shape2 
      Height          =   1695
      Left            =   240
      Top             =   7080
      Width           =   12615
   End
   Begin VB.Label L4_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L4_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2280
      TabIndex        =   46
      Top             =   8160
      Width           =   8835
   End
   Begin VB.Label L3_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L3_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2280
      TabIndex        =   45
      Top             =   7920
      Width           =   8835
   End
   Begin VB.Label L2_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L2_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2280
      TabIndex        =   44
      Top             =   7680
      Width           =   8835
   End
   Begin VB.Label L1_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L1_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2280
      TabIndex        =   43
      Top             =   7440
      Width           =   8835
   End
   Begin VB.Label L5_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L5_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2280
      TabIndex        =   42
      Top             =   8400
      Width           =   8835
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "** Anda perlu menutup menu ini dahulu sebelum boleh ke menu seterusnya."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2280
      TabIndex        =   41
      Top             =   11280
      Width           =   7095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Pendaftaran data pelanggan."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   40
      Top             =   9000
      Width           =   6555
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nama * :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      TabIndex        =   39
      Top             =   9870
      Width           =   1755
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "** Ruangan ini hanya digunakan untuk pendaftaran pelanggan biasa. Bagi pendaftaran ahli sila ke menu ""Maklumat Pelanggan"""
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   525
      Left            =   480
      TabIndex        =   38
      Top             =   9240
      Width           =   7785
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Kad Pengenalan * :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      TabIndex        =   37
      Top             =   10230
      Width           =   1755
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Telefon :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      TabIndex        =   36
      Top             =   10590
      Width           =   1755
   End
   Begin VB.Shape Shape3 
      Height          =   2295
      Left            =   240
      Top             =   8880
      Width           =   12615
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Kad Pengenalan :"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      TabIndex        =   35
      Top             =   7680
      Width           =   2010
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nama :"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      TabIndex        =   34
      Top             =   7440
      Width           =   2010
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Telefon :"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      TabIndex        =   33
      Top             =   7920
      Width           =   2010
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail :"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      TabIndex        =   32
      Top             =   8160
      Width           =   2010
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Pelanggan :"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      TabIndex        =   31
      Top             =   8400
      Width           =   2010
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai Pelanggan."
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
      Height          =   315
      Left            =   240
      TabIndex        =   30
      Top             =   2280
      Width           =   4650
   End
   Begin VB.Label L70_Text 
      BackColor       =   &H8000000C&
      Caption         =   "L70_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   6120
      TabIndex        =   29
      Top             =   6240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label L69_Text 
      BackColor       =   &H8000000C&
      Caption         =   "L69_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   7080
      TabIndex        =   28
      Top             =   6240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label L67_Text 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "L67_Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9360
      TabIndex        =   27
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label L68_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L68_Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9960
      TabIndex        =   26
      Top             =   6120
      Width           =   615
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Maklumat pelanggan yang telah dipilih."
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
      Height          =   315
      Left            =   360
      TabIndex        =   25
      Top             =   7200
      Width           =   4650
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Bilangan :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   240
      TabIndex        =   24
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label L71_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L71_Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   1080
      TabIndex        =   23
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label L72_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L72_Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   360
      TabIndex        =   22
      Top             =   6360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Carian data pelanggan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   20
      Top             =   240
      Width           =   6555
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Scanner Mode (Hanya boleh digunakan bagi carian [No. Keahlian] sahaja."
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   13485
      TabIndex        =   18
      Top             =   1425
      Visible         =   0   'False
      Width           =   6930
   End
   Begin VB.Label L6_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L6_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   13320
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label Label45 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Keyword Carian :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   14
      Top             =   1590
      Width           =   1785
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"Frm28.frx":729A
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
      Height          =   525
      Left            =   360
      TabIndex        =   13
      Top             =   600
      Width           =   10515
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "**Hanya maklumat ahli yang sudah berdaftar dengan sistem sahaja akan dipaparkan."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   12
      Top             =   1080
      Width           =   7395
   End
   Begin VB.Shape Shape1 
      Height          =   2175
      Left            =   120
      Top             =   120
      Width           =   12735
   End
   Begin VB.Label Label90 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Kad Pengenalan               No. Keahlian"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   13470
      TabIndex        =   15
      Top             =   1110
      Visible         =   0   'False
      Width           =   4650
   End
   Begin VB.Menu frm28_pm_menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu frm28_sm_pilih 
         Caption         =   "Pilih Maklumat Pelanggan Ini"
      End
   End
End
Attribute VB_Name = "Frm28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB2_Click()
'On Error Resume Next
If Frm28.CB2 = 1 Then
    Frm28.CB3 = 0
    Frm28.L6_Text = "No. Kad Pengenalan"
End If
End Sub
Private Sub CB3_Click()
'On Error Resume Next
If Frm28.CB3 = 1 Then
    Frm28.CB2 = 0
    Frm28.L6_Text = "No. Keahlian"
End If
End Sub
Private Sub CMD1_Click()
'on error resume next
Call frm28_periksa_carian
End Sub
Private Sub CMD2_Click()
'on error resume next
Note = "Padamkan maklumat pembeli Iini ?" & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then

    If MDI_frm1.L5_Text = 4 Then frm130.L26_Text = Format(0, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
    If MDI_frm1.L5_Text = 7 Then Frm87.L27_Text = Format(0, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
    If MDI_frm1.L5_Text = 10 Then frm130.L26_Text = Format(0, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
    If MDI_frm1.L5_Text = 8 Then frm130.L26_Text = Format(0, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
    If MDI_frm1.L5_Text = 9 Then frm130.L26_Text = Format(0, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
    
    Frm28.L1_Text = vbNullString
    Frm28.L2_Text = vbNullString
    Frm28.L3_Text = vbNullString
    Frm28.L4_Text = vbNullString
    Frm28.L5_Text = vbNullString
    
End If
End Sub

Private Sub CMD21_Click()
'on error resume next
Dim frm28_LM_CURR_PAGE As Double
Dim frm28_LM_TOTAL_PAGE As Double

frm28_LM_CURR_PAGE = 0
frm28_LM_TOTAL_PAGE = 0

If Frm28.L67_Text <> vbNullString And IsNumeric(Frm28.L67_Text) Then
    If Frm28.L68_Text <> vbNullString And IsNumeric(Frm28.L68_Text) Then
        frm28_LM_CURR_PAGE = Frm28.L67_Text
        frm28_LM_TOTAL_PAGE = Frm28.L68_Text
        
        If frm28_LM_CURR_PAGE <> 1 And frm28_LM_CURR_PAGE <> 0 Then
        
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
                    
            Call frm28_senarai_pelanggan_header
            Call frm28_senarai_pelanggan
                    
        End If

    End If
End If
End Sub
Private Sub CMD22_Click()
'on error resume next
Dim frm28_LM_CURR_PAGE As Double
Dim frm28_LM_TOTAL_PAGE As Double

frm28_LM_CURR_PAGE = 0
frm28_LM_TOTAL_PAGE = 0

If Frm28.L67_Text <> vbNullString And IsNumeric(Frm28.L67_Text) Then
    If Frm28.L68_Text <> vbNullString And IsNumeric(Frm28.L68_Text) Then
        frm28_LM_CURR_PAGE = Frm28.L67_Text
        frm28_LM_TOTAL_PAGE = Frm28.L68_Text
        
        If frm28_LM_CURR_PAGE < frm28_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call frm28_senarai_pelanggan_header
            Call frm28_senarai_pelanggan
            
        End If
    End If
End If
End Sub
Private Sub CMD3_Click()
'on error resume next
Frm28.Hide
End Sub
Private Sub CMD4_Click()
'on error resume next
'Data Pelanggan
'---------------------
'0:  Pendaftaran Biasa
'1 : Jualan Gold Bar
'2 : Buyback Gold Bar
'3:  Jualan BK
'4:  Buyback BK
'5:  Ansuran
'6:  Servis
'7:  Tempahan
'8:  Jualan kepada agen

If Frm84.Visible = True Then
    Frm68.L15_Text = 3
    Frm68.L37_Text = Frm84.L45_Text
End If
If Frm83.Visible = True Then
    Frm68.L15_Text = 4
    Frm68.L37_Text = Frm83.L39_Text
End If
If Frm87.Visible = True Then
    Frm68.L15_Text = 5
    Frm68.L37_Text = Frm87.L40_Text
End If
If Frm93.Visible = True Then
    Frm68.L15_Text = 7
    Frm68.L37_Text = Frm93.L37_Text
End If
If Frm92.Visible = True Then
    Frm68.L15_Text = 6
    Frm68.L37_Text = Frm92.L54_Text
End If
If Frm102.Visible = True Then
    Frm68.L15_Text = 8
    Frm68.L37_Text = 3
End If
'Data Agen Drophip
'--------------------
'20 : Jualan

Frm68.L36_Text = 1 '0 : Terus dari menu data pelanggan , 1 : Data pembeli , 2 : Data agen dropship

If Frm28.L1_Text <> vbNullString Then

    Note = "Maklumat penjual akan dipadamkan jika ada meneruskan menu ini." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan?"
            
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then
        Unload Frm28
        
        If MDI_frm1.L5_Text = 3 Then Frm83.Hide
        
        Frm68.Show
    End If
    
Else
    
    Unload Frm28
    If MDI_frm1.L5_Text = 3 Then Frm83.Hide
    
    Frm68.Show
End If
End Sub

Private Sub CMD5_Click()
'on error resume next
Dim myAt As Integer
Dim myDot As Integer
Dim myDotDot As Integer
Dim Err(5)

DATA_WRITE = 0 '0 : Tiada Data Disimpan , 1 : Data Telah Disimpan
Frm28_LM_No_PELANGGAN = 0 'No. Giliran Pelanggan
Frm28_LM_INVOICE_AHLI = 1
Frm28_LM_ACTIVE = 0

If Frm28.TB2 = vbNullString Then

    MsgBox "Sila masukkan nama", vbExclamation, "Info"
    
    Exit Sub
End If
If Frm28.TB3 = vbNullString Then

    MsgBox "Sila masukkan no. kad pengenalan", vbExclamation, "Info"
    
    Exit Sub
End If
    
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from senarai_pelanggan where no_ic='" & UCase(Frm28.TB3) & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    Frm28_LM_KATEGORI = "Pelanggan Biasa"
    If Not IsNull(rs!Nama) Then Frm28_LM_NAMA = rs!Nama 'Nama
    If Not IsNull(rs!no_ic) Then Frm28_LM_No_IC = rs!no_ic 'No. IC
    If Not IsNull(rs!no_pelanggan) Then Frm28_LM_No_PELANGGAN = rs!no_pelanggan 'No. Pelanggan
    If Not IsNull(rs!kategori_pelanggan) Then
        If rs!kategori_pelanggan = 1 Then Frm28_LM_KATEGORI = "Pelanggan Biasa"
        If rs!kategori_pelanggan = 2 Then Frm28_LM_KATEGORI = "Ahli Biasa"
        If rs!kategori_pelanggan = 3 Then Frm28_LM_KATEGORI = "Silver"
        If rs!kategori_pelanggan = 4 Then Frm28_LM_KATEGORI = "Gold"
        If rs!kategori_pelanggan = 5 Then Frm28_LM_KATEGORI = "Platinum"
    End If
    
    MsgBox "Pelanggan dengan No. Kad Pengenalan [" & Frm28_LM_No_IC & "] telah didaftarkan sebelum ini." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Maklumat yang telah didaftarkan adalah seperti berikut :" & vbCrLf & _
            "Nama : " & Frm28_LM_NAMA & vbCrLf & _
            "No. Kad Pengenalan : " & Frm28_LM_No_IC & vbCrLf & _
            "No. Pelanggan : " & Frm28_LM_No_PELANGGAN & vbCrLf & _
            "Kategori : " & Frm28_LM_KATEGORI & vbCrLf & _
            vbNullString, vbExclamation, "Info"
    
    rs.Close
    Set rs = Nothing
    
    Exit Sub
    
End If

rs.Close
Set rs = Nothing

'### Periksa apakah jenis kategori pendaftaran ### - Start

Frm28_LM_KATEGORI = "Pelanggan Biasa"
Frm28_LM_CODE = "C"

Note = "Adakah anda ingin simpan data ini?"

'### Periksa apakah jenis kategori pendaftaran ### - End
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then

'### Carian No. Pelanggan ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If rs!Default1 = "Default" Then
            
            If Not IsNull(rs!no_customer) Then Frm28_LM_No_PELANGGAN = rs!no_customer 'No. Giliran Pelanggan
            
            If Not IsNull(rs!kod_customer) Then
                Frm28_LM_KOD_KEDAI = rs!kod_customer 'Kod Kedai
            Else
            
                MsgBox "Tiada maklumat tentang Kod Kedai" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Sila hubungi Sankyu System untuk langkah seterusnya.", vbExclamation, "Error"
                
                rs.Close
                Set rs = Nothing
                
                Exit Sub
            End If
        End If
    End If
    
    rs.Close
    Set rs = Nothing
'### Carian No. Pelanggan ### - End

'### Periksa samada No. Pelanggan telah digunakan atau tidak ### - Start
re_gen_number:
    LM_NO_AHLI = Frm28_LM_KOD_KEDAI & Frm28_LM_CODE & Format(Frm28_LM_No_PELANGGAN, "00000")
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & LM_NO_AHLI & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        Frm28_LM_No_PELANGGAN = Frm28_LM_No_PELANGGAN + 1
        
        rs.Close
        Set rs = Nothing
        
        GoTo re_gen_number:
    End If
    
    rs.Close
    Set rs = Nothing

'### Periksa samada No. Pelanggan telah digunakan atau tidak ### - End

'### Simpan data pelanggan ke dalam database ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from senarai_pelanggan where no_ic='" & UCase(Frm28.TB2) & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
    
    If rs.EOF Then
        rs.AddNew
        rs!kategori_pelanggan = 1
        If Frm28.TB2 <> vbNullString Then
            rs!Nama = UCase(Frm28.TB2) 'Nama
        Else
            rs!Nama = Null
        End If
        If Frm28.TB3 <> vbNullString Then
            rs!no_ic = UCase(Frm28.TB3) 'No. IC
        Else
            rs!no_ic = Null
        End If
        If LM_NO_AHLI <> vbNullString Then
            rs!no_pelanggan = UCase(LM_NO_AHLI) 'No. Pelanggan
        Else
            rs!no_pelanggan = Null
        End If
        If Frm28.TB4 <> vbNullString Then
            rs!no_tel = UCase(Frm28.TB4) 'No. Tel
        Else
            rs!no_tel = Null
        End If
        rs!dropship = 0 '0 : Bukan agen dropship , 1 : Agen dropship
        rs!baki_simpanan = "0.00" 'Baki Simpan Di Kedai
        rs!membership_card = 0 '0 : Tiada kad keahlian , 1 : Ada kad keahlian
        rs!yuran_flag = 0 'Flag samada ada bayaran yang dikenakan bagi pendaftaran ini atau tidak (0 : Tiada bayaran , 1 : Ada bayaran)
        rs!jumlah_yuran = "0.00"
        rs!no_invoice = Null
        rs!tarikh = DateTime.Date 'Tarikh pendaftaran
        rs!Status = 1 '0 : Sudah dipadamkan , 1 : Aktif , 2 : Tidak aktif
        rs!write_timestamp = Now 'Tarikh Data Dimasukkan
        DATA_WRITE = 1 '0 : Tiada Data Disimpan , 1 : Data Telah Disimpan , 2 : Data Telah Diedit
        rs.Update
        
    Else
    
        MsgBox "Pengguna dengan No. Kad Pengenalan " & UCase(Frm28.TB3) & " telah didaftarkan sebelum ini." & vbCrLf & _
                "Sila periksa senarai pelanggan ini.", vbInformation, "Info"
        
    End If
    
    rs.Close
    Set rs = Nothing
'### Simpan data pelanggan ke dalam database ### - End
    
    If DATA_WRITE = 1 Then  '0 : Tiada Data Disimpan , 1 : Data Customer Baru Telah Disimpan , 2 : Data Belian Telah Disimpan
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from default_setting where Default1='" & "Default" & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then

            rs!no_customer = Frm28_LM_No_PELANGGAN + 1 'No. Giliran Pelanggan

            rs.Update
        End If
        
        rs.Close
        Set rs = Nothing
        
        user = MDI_frm1.L3_Text
        LogAct_Memory = "[" & user & "] Pendaftaran data pelanggan. IC [" & UCase(Frm28.TB3) & "] , No. Pelanggan [" & UCase(LM_NO_AHLI) & "]"
        LogDate_Memory = DateTime.Date & " " & DateTime.Time$
        Call UpdateLog_Database
        
        Frm28.TB2 = vbNullString
        Frm28.TB3 = vbNullString
        Frm28.TB4 = vbNullString
        
        MsgBox "Data pelanggan telah berjaya disimpan.", vbInformation, "Info"

    End If
End If
End Sub

Private Sub Command1_Click()
'on error resume next
Dim myAt As Integer
Dim myDot As Integer
Dim myDotDot As Integer
Dim Err(10)

DATA_WRITE = 0 '0 : Tiada Data Disimpan , 1 : Data Telah Disimpan
Frm28_LM_No_PELANGGAN = 0 'No. Giliran Pelanggan
Frm28_LM_INVOICE_AHLI = 1
Frm28_LM_ACTIVE = 0

If Frm28.TB2 = vbNullString Then
    MsgBox "Sila Masukkan [Nama].", vbExclamation, "Info"
    Exit Sub
End If
If Frm28.TB2 <> vbNullString Then
    If InStr(1, Frm28.TB2, "*") <> 0 Or InStr(1, Frm28.TB2, "/") <> 0 Or InStr(1, Frm28.TB2, "\") <> 0 Or InStr(1, Frm28.TB2, "'") <> 0 Or InStr(1, Frm28.TB2, "`") <> 0 Then
        MsgBox "[Nama] Mengandungi Simbol Yang Tidak Dibenarkan.", vbExclamation, "Info"
        Exit Sub
    End If
End If
If Frm28.TB3 = vbNullString Then
    MsgBox "Sila Masukkan [No. Kad Pengenalan].", vbExclamation, "Info"
    Exit Sub
End If
If Frm28.TB3 <> vbNullString Then
    If InStr(1, Frm28.TB3, "*") <> 0 Or InStr(1, Frm28.TB3, "/") <> 0 Or InStr(1, Frm28.TB3, "\") <> 0 Or InStr(1, Frm28.TB3, "'") <> 0 Or InStr(1, Frm28.TB3, "`") <> 0 Then
        MsgBox "[No. Kad Pengenalan] Mengandungi Simbol Yang Tidak Dibenarkan.", vbExclamation, "Info"
        Exit Sub
    End If
End If
    
'### Periksa kewujudan NO KAD PENGENALAN ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from senarai_pelanggan where no_ic='" & UCase(Frm28.TB3) & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    Frm28_LM_KATEGORI = "Pelanggan Biasa"
    If Not IsNull(rs!Nama) Then Frm28_LM_NAMA = rs!Nama 'Nama
    If Not IsNull(rs!no_ic) Then Frm28_LM_No_IC = rs!no_ic 'No. IC
    If Not IsNull(rs!no_pelanggan) Then Frm28_LM_No_PELANGGAN = rs!no_pelanggan 'No. Pelanggan
    If Not IsNull(rs!kategori_pelanggan) Then
        If rs!kategori_pelanggan = 1 Then Frm28_LM_KATEGORI = "Pelanggan Biasa"
        If rs!kategori_pelanggan = 2 Then Frm28_LM_KATEGORI = "Ahli Biasa"
        If rs!kategori_pelanggan = 3 Then Frm28_LM_KATEGORI = "Silver"
        If rs!kategori_pelanggan = 4 Then Frm28_LM_KATEGORI = "Gold"
        If rs!kategori_pelanggan = 5 Then Frm28_LM_KATEGORI = "Platinum"
    End If
    
    MsgBox "Pelanggan dengan No. Kad Pengenalan [" & Frm28_LM_No_IC & "] telah didaftarkan sebelum ini." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Maklumat yang telah didaftarkan adalah seperti berikut :" & vbCrLf & _
            "Nama : " & Frm28_LM_NAMA & vbCrLf & _
            "No. Kad Pengenalan : " & Frm28_LM_No_IC & vbCrLf & _
            "No. Pelanggan : " & Frm28_LM_No_PELANGGAN & vbCrLf & _
            "Kategori : " & Frm28_LM_KATEGORI & vbCrLf & _
            vbNullString, vbExclamation, "Info"
    
    rs.Close
    Set rs = Nothing
    
    Exit Sub
    
End If

rs.Close
Set rs = Nothing
'### Periksa kewujudan NO KAD PENGENALAN ### - End

'### Periksa apakah jenis kategori pendaftaran ### - Start
Frm28_LM_KATEGORI = "Pelanggan Biasa"

Note = "Adakah anda ingin mendaftarkan pelanggan ini ?" & vbCrLf & _
        vbNullString & vbCrLf & _
        "Maklumat pelanggan adalah seperti di bawah :" & vbCrLf & _
        "Nama : " & UCase(Frm28.TB2) & vbCrLf & _
        "No. Kad Pengenalan : " & UCase(Frm28.TB3) & vbCrLf & _
        "Kategori : " & Frm28_LM_KATEGORI & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan ?"
    
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    
'---------------------------------------No. Invoice
    LM_NOW = Now
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
    strsql = "insert into 13_rujukan_customer(tarikh,terminal,write_timestamp,Status,nama_staff,cawangan)" & _
                "select '" & DateTime.Date & "','" & G_TERMINAL & "','" & LM_NOW & "',1,'" & MDI_frm1.L3_Text & "','" & G_CAWANGAN & "'"
    
    Set rs = cn2.Execute(strsql)
    Set rs = Nothing
        
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
    rs.Open "select * from 13_rujukan_customer where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND cawangan='" & G_CAWANGAN & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & DateTime.Date & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic

    If Not rs.EOF Then
        If Not IsNull(rs!ID) Then
            rs!no_rujukan = G_KOD_KEDAI & "-" & Format(rs!ID, "000000")
            LM_NO_CUSTOMER = G_KOD_KEDAI & "-" & Format(rs!ID, "000000")
            rs.Update
        End If
    Else
        MsgBox "Berlaku ralat semasa data cuba disimpan. Sila keluar dari menu ini dan cuba lagi.", vbCritical, "Error"
        
        rs.Close
        Set rs = Nothing
        
        Exit Sub
    End If
    
    rs.Close
    Set rs = Nothing
    
'### Simpan data pelanggan ke dalam database ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from senarai_pelanggan where no_ic='" & UCase(Frm28.TB2) & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
    
    If rs.EOF Then
    
        rs.AddNew
        rs!kategori_pelanggan = 1
        rs!no_pelanggan = LM_NO_CUSTOMER
        If Frm28.TB2 <> vbNullString Then
            rs!Nama = UCase(Frm28.TB2) 'Nama
        Else
            rs!Nama = Null
        End If
        If Frm28.TB3 <> vbNullString Then
            rs!no_ic = UCase(Frm28.TB3) 'No. IC
        Else
            rs!no_ic = Null
        End If
        If Frm28.TB4 <> vbNullString Then
            rs!no_tel = UCase(Frm28.TB4) 'No. Tel
        Else
            rs!no_tel = Null
        End If
        rs!dropship = 0 '0 : Bukan agen dropship , 1 : Agen dropship
        rs!baki_simpanan = "0.00" 'Baki Simpan Di Kedai
        rs!membership_card = 0 '0 : Tiada kad keahlian , 1 : Ada kad keahlian
        rs!yuran_flag = 0 'Flag samada ada bayaran yang dikenakan bagi pendaftaran ini atau tidak (0 : Tiada bayaran , 1 : Ada bayaran)
        rs!jumlah_yuran = "0.00"
        rs!no_invoice = Null
        rs!tarikh = DateTime.Date 'Tarikh pendaftaran
        rs!Status = 1 '0 : Sudah dipadamkan , 1 : Aktif , 2 : Tidak aktif
        rs!cawangan = G_CAWANGAN
        rs!write_timestamp = LM_NOW 'Tarikh Data Dimasukkan
        DATA_WRITE = 1 '0 : Tiada Data Disimpan , 1 : Data Telah Disimpan , 2 : Data Telah Diedit
        rs.Update
        
    Else
    
        MsgBox "Pengguna dengan No. Kad Pengenalan " & UCase(Frm28.TB3) & " telah didaftarkan sebelum ini." & vbCrLf & _
                "Sila periksa senarai pelanggan ini.", vbInformation, "Info"
        
    End If
    
    rs.Close
    Set rs = Nothing
'### Simpan data pelanggan ke dalam database ### - End
    
    If DATA_WRITE = 1 Then  '0 : Tiada Data Disimpan , 1 : Data Customer Baru Telah Disimpan , 2 : Data Belian Telah Disimpan
        
        user = MDI_frm1.L3_Text
        LogAct_Memory = "[" & user & "] Pendaftaran pelanggan baru. IC [" & UCase(Frm28.TB3) & "] , No. Pelanggan [" & LM_NO_CUSTOMER & "]"
        LogDate_Memory = LM_NOW
        Call UpdateLog_Database
        
        If MDI_frm1.L5_Text = 4 Or MDI_frm1.L5_Text = 5 Then
            Frm84.L79_Text = 0 '0 : Tiada kad , 1 : Ada kad
            Frm84.L77_Text = "0"
        End If
            
        Frm28.L1_Text = UCase(Frm28.TB2)
        Frm28.L2_Text = UCase(Frm28.TB3)
        Frm28.L3_Text = UCase(Frm28.TB4)
        Frm28.L5_Text = LM_NO_CUSTOMER
        
        Frm28.TB2 = vbNullString
        Frm28.TB3 = vbNullString
        Frm28.TB4 = vbNullString
        
        MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
        
        'Frm28.TB18.SetFocus
    End If
End If
End Sub

Private Sub Form_Load()
'on error resume next
Frm28.CB2 = 1
Frm28.CB3 = 0

'Set rs = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

'If Not rs.EOF Then
'    If rs!Default1 = "Default" Then
'        If Not IsNull(rs!ScannerMode) Then
'            If rs!ScannerMode = 1 Then
'                Frm28.CB1 = 1
'            Else
'                Frm28.CB1 = 0
'            End If
'        Else
'            Frm28.CB1 = 0
'        End If
'    End If
'End If

'rs.Close
'Set rs = Nothing
End Sub

Private Sub frm28_sm_pilih_Click()
'on error resume next
Dim rs11 As ADODB.Recordset
DATA_FOUND = 0

If IsNumeric(Frm28.LV1.SelectedItem.Index) Then
    
    frm28_LM_No_ID = Frm28.LV1.ListItems(Frm28.LV1.SelectedItem.Index)
    
    If frm28_LM_No_ID <> vbNullString Then
        
        Frm28.L1_Text = vbNullString
        Frm28.L2_Text = vbNullString
        Frm28.L3_Text = vbNullString
        Frm28.L4_Text = vbNullString
        Frm28.L5_Text = vbNullString
        
        Set rs11 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs11.Open "select * from senarai_pelanggan where ID='" & frm28_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
    
        If Not rs11.EOF Then

            If Not IsNull(rs11!Nama) Then Frm28.L1_Text = rs11!Nama 'Nama
            If Not IsNull(rs11!no_ic) Then Frm28.L2_Text = rs11!no_ic 'No. IC
            If Not IsNull(rs11!no_tel) Then Frm28.L3_Text = rs11!no_tel 'No. Telefon
            If Not IsNull(rs11!Email) Then Frm28.L4_Text = rs11!Email 'E-mail
            If Not IsNull(rs11!no_pelanggan) Then Frm28.L5_Text = rs11!no_pelanggan 'No. Customer
            If Not IsNull(rs11!baki_simpanan) Then frm130.L26_Text = Format(rs11!baki_simpanan, "#,##0.00")
            
            If MDI_frm1.L5_Text = 4 Or MDI_frm1.L5_Text = 5 Then
            
                If Not IsNull(rs11!membership_card) Then
                    If rs11!membership_card = 0 Then
                    
                        If MDI_frm1.L5_Text = 4 Or MDI_frm1.L5_Text = 5 Then
                            Frm84.L79_Text = 0 '0 : Tiada kad , 1 : Ada kad
                        End If
            
                        Frm84.L77_Text = "0"
            
                    ElseIf rs11!membership_card = 1 Then
                    
                        If MDI_frm1.L5_Text = 4 Or MDI_frm1.L5_Text = 5 Then
                            Frm84.L79_Text = 1 '0 : Tiada kad , 1 : Ada kad
                        End If
                        If Not IsNull(rs11!baki_point) Then
                            Frm84.L77_Text = rs11!baki_point
                        Else
                            Frm84.L77_Text = "0"
                        End If
                    End If
                Else
                
                    If MDI_frm1.L5_Text = 4 Or MDI_frm1.L5_Text = 5 Then
                        Frm84.L79_Text = 0 '0 : Tiada kad , 1 : Ada kad
                    End If
            
                    Frm84.L77_Text = "0"
            
                End If
                
            End If
        
        End If
        
        rs11.Close
        Set rs11 = Nothing
        
    Else
    
        MsgBox "Tiada Data.", vbInformation, "Info"
        
    End If
    
Else

    MsgBox "Tiada Data.", vbInformation, "Info"
    
End If
End Sub

Private Sub L1_Text_Change()
'on error resume next
If MDI_frm1.L5_Text = "3" Then Frm83.L37_Text = UCase(Frm28.L1_Text)  'Nama Pembeli
If MDI_frm1.L5_Text = "4" Or MDI_frm1.L5_Text = "5" Then
    Frm84.L28_Text = UCase(Frm28.L1_Text)  'Nama Pembeli
End If
If MDI_frm1.L5_Text = "6" Then Frm102.L46_Text = UCase(Frm28.L1_Text) 'Nama Pembeli
If MDI_frm1.L5_Text = "7" Then Frm87.L6_Text = UCase(Frm28.L1_Text) 'Nama Pembeli
If MDI_frm1.L5_Text = "10" Then Frm92.L52_Text = UCase(Frm28.L1_Text) 'Nama Pembeli
If MDI_frm1.L5_Text = "8" Then Frm93.L36_Text = UCase(Frm28.L1_Text) 'Nama Pembeli
'If Frm102.Visible = True Then Frm102.L46_Text = UCase(Frm28.L1_Text) 'Nama Pembeli
End Sub

Private Sub LV1_DblClick()
'on error resume next
frm28_LM_No_ID = vbNullString

If IsNumeric(Frm28.LV1.SelectedItem.Index) Then
    
    frm28_LM_No_ID = Frm28.LV1.SelectedItem.Index
    
    If frm28_LM_No_ID <> vbNullString Then

        PopupMenu frm28_pm_menu
    
    Else
    
        MsgBox "Tiada Data.", vbInformation, "Info"
        
    End If
    
Else

    MsgBox "Tiada Data.", vbInformation, "Info"
    
End If
End Sub

Private Sub TB1_Change()
'On Error Resume Next
If Frm28.CB1 = 1 And Frm28.CB3 = 1 And Frm28.TB1 <> vbNullString Then
    Frm28.Tmr1.Enabled = False
    Frm28.Tmr1.Enabled = True
    Frm28.Tmr1.Interval = 100
End If
End Sub
Private Sub Tmr1_Timer()
'On Error Resume Next
If Frm28.CB1 = 1 And Frm28.CB3 = 1 And Frm28.TB1 <> vbNullString And Frm28.Tmr1.Enabled = True And Frm28.Visible = True Then
    If Frm28.Tmr1.Interval = 100 Then
        If InStr(1, Frm28.TB1, "'") <> 0 Then
            MsgBox "No. Keahlian Mengandungi Simbol Yang Tidak Sah , ['].", vbInformation, "Info"
            Exit Sub
        End If
        
        Call Frm28_carian_ahli
    End If
End If
End Sub
Private Sub TB1_KeyPress(KeyAscii As Integer)
'on error resume next
If KeyAscii = 13 Then
    
    Call frm28_periksa_carian

End If
End Sub
