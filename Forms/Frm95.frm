VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm95 
   Caption         =   "Tetapan Asas Sistem"
   ClientHeight    =   13035
   ClientLeft      =   120
   ClientTop       =   -13770
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
   Icon            =   "Frm95.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13035
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Senarai Dulang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11775
      Left            =   6000
      TabIndex        =   101
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
      Begin MSComctlLib.ListView LV4 
         Height          =   10725
         Left            =   120
         TabIndex        =   102
         Top             =   360
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   18918
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
         NumItems        =   0
      End
      Begin VB.Label L30_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L30_Text"
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
         Left            =   2115
         TabIndex        =   105
         Top             =   11400
         Width           =   1215
      End
      Begin VB.Label L29_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L29_Text"
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
         Left            =   2115
         TabIndex        =   104
         Top             =   11175
         Width           =   1215
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan Aktif               :    Bilangan Tidak Aktif    :"
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
         Height          =   660
         Left            =   120
         TabIndex        =   103
         Top             =   11160
         Width           =   2295
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pendaftaran Dulang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   7200
      TabIndex        =   92
      Top             =   1440
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CommandButton CMD10 
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
         Height          =   855
         Left            =   2280
         MouseIcon       =   "Frm95.frx":0ECA
         MousePointer    =   99  'Custom
         Picture         =   "Frm95.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   100
         ToolTipText     =   "Simpan data"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton CMD12 
         BackColor       =   &H80000003&
         Caption         =   "Batal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3240
         MouseIcon       =   "Frm95.frx":1B7E
         MousePointer    =   99  'Custom
         Picture         =   "Frm95.frx":1E88
         Style           =   1  'Graphical
         TabIndex        =   99
         ToolTipText     =   "Batal"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton CMD11 
         BackColor       =   &H80000003&
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
         Height          =   855
         Left            =   1200
         MouseIcon       =   "Frm95.frx":2F52
         MousePointer    =   99  'Custom
         Picture         =   "Frm95.frx":325C
         Style           =   1  'Graphical
         TabIndex        =   98
         ToolTipText     =   "Simpan data"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox TB14 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1620
         TabIndex        =   93
         Text            =   "TB14"
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label L15_Text 
         Caption         =   "L15_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   97
         Top             =   1080
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "** Sila gunakan maksimum 2 abjad SAHAJA dalam ruangan ini."
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
         Height          =   525
         Left            =   3120
         OLEDropMode     =   1  'Manual
         TabIndex        =   96
         Top             =   795
         Width           =   3705
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan ruangan di bawah bagi pendaftaran dulang."
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   95
         Top             =   360
         Width           =   6705
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Dulang * :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   94
         Top             =   750
         Width           =   1905
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pendaftaran Kategori"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   10320
      TabIndex        =   81
      Top             =   6360
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CommandButton CMD8 
         BackColor       =   &H80000003&
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
         Height          =   855
         Left            =   1560
         MouseIcon       =   "Frm95.frx":3C06
         MousePointer    =   99  'Custom
         Picture         =   "Frm95.frx":3F10
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Simpan data"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton CMD9 
         BackColor       =   &H80000003&
         Caption         =   "Batal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3600
         MouseIcon       =   "Frm95.frx":48BA
         MousePointer    =   99  'Custom
         Picture         =   "Frm95.frx":4BC4
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Batal"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton CMD7 
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
         Height          =   855
         Left            =   2640
         MouseIcon       =   "Frm95.frx":5C8E
         MousePointer    =   99  'Custom
         Picture         =   "Frm95.frx":5F98
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Simpan data"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox TB12 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1620
         TabIndex        =   22
         Text            =   "TB12"
         Top             =   720
         Width           =   4500
      End
      Begin VB.TextBox TB13 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1620
         TabIndex        =   23
         Text            =   "TB13"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label L14_Text 
         Caption         =   "L14_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5280
         TabIndex        =   86
         Top             =   360
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "** Sila gunakan maksimum 2 abjad SAHAJA dalam ruangan ini."
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
         Height          =   525
         Left            =   3120
         OLEDropMode     =   1  'Manual
         TabIndex        =   85
         Top             =   1155
         Width           =   3705
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Produk * :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   84
         Top             =   750
         Width           =   1905
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Kod Produk *    :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   83
         Top             =   1110
         Width           =   1905
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan ruangan di bawah bagi pendaftaran produk."
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   82
         Top             =   360
         Width           =   6705
      End
   End
   Begin VB.PictureBox Pic10 
      BorderStyle     =   0  'None
      Height          =   10215
      Left            =   20160
      ScaleHeight     =   10215
      ScaleWidth      =   7440
      TabIndex        =   45
      Top             =   7080
      Visible         =   0   'False
      Width           =   7440
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid5 
         Height          =   9525
         Left            =   120
         TabIndex        =   47
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   480
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   16801
         _Version        =   393216
         Rows            =   1
         Cols            =   0
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   12648384
         BackColorSel    =   16777215
         ForeColorSel    =   16711680
         BackColorBkg    =   16777215
         GridColor       =   0
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai tukang emas."
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   46
         Top             =   240
         Width           =   5385
      End
   End
   Begin VB.PictureBox Pic9 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   1080
      ScaleHeight     =   1935
      ScaleWidth      =   6975
      TabIndex        =   41
      Top             =   9240
      Visible         =   0   'False
      Width           =   6975
      Begin VB.CommandButton CMD15 
         Caption         =   "Batal"
         Height          =   375
         Left            =   3360
         MouseIcon       =   "Frm95.frx":6942
         MousePointer    =   99  'Custom
         TabIndex        =   30
         ToolTipText     =   "Batal"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton CMD14 
         Caption         =   "Simpan Data"
         Height          =   375
         Left            =   1320
         MouseIcon       =   "Frm95.frx":6C4C
         MousePointer    =   99  'Custom
         TabIndex        =   29
         ToolTipText     =   "Simpan data"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton CMD13 
         Caption         =   "Simpan Data"
         Height          =   375
         Left            =   2280
         MouseIcon       =   "Frm95.frx":6F56
         MousePointer    =   99  'Custom
         TabIndex        =   28
         ToolTipText     =   "Simpan data"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox TB15 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1620
         TabIndex        =   27
         Text            =   "TB15"
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label L18_Text 
         Caption         =   "L18_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   44
         Top             =   1080
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan ruangan di bawah bagi pendaftaran tukang emas."
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   43
         Top             =   360
         Width           =   6705
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Tukang Emas * :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   42
         Top             =   750
         Width           =   1905
      End
   End
   Begin VB.Timer Tmr1 
      Interval        =   10
      Left            =   480
      Top             =   5040
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pendaftaran Supplier"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   13560
      TabIndex        =   48
      Top             =   240
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CommandButton CMD1 
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
         Height          =   855
         Left            =   3720
         MouseIcon       =   "Frm95.frx":7260
         MousePointer    =   99  'Custom
         Picture         =   "Frm95.frx":756A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Simpan data"
         Top             =   5640
         Width           =   1935
      End
      Begin VB.CommandButton CMD2 
         BackColor       =   &H80000003&
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
         Height          =   855
         Left            =   2640
         MouseIcon       =   "Frm95.frx":7F14
         MousePointer    =   99  'Custom
         Picture         =   "Frm95.frx":821E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Simpan data"
         Top             =   5640
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton CMD3 
         BackColor       =   &H80000003&
         Caption         =   "Batal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4680
         MouseIcon       =   "Frm95.frx":8BC8
         MousePointer    =   99  'Custom
         Picture         =   "Frm95.frx":8ED2
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Batal"
         Top             =   5640
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox TB3 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1830
         TabIndex        =   4
         Top             =   2040
         Width           =   6500
      End
      Begin VB.TextBox TB4 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1830
         TabIndex        =   5
         Top             =   2400
         Width           =   6500
      End
      Begin VB.TextBox TB5 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1830
         TabIndex        =   6
         Top             =   2760
         Width           =   6500
      End
      Begin VB.TextBox TB6 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1830
         TabIndex        =   7
         Top             =   3120
         Width           =   6500
      End
      Begin VB.TextBox TB7 
         BackColor       =   &H00FFFFFF&
         Height          =   1320
         Left            =   1830
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   3480
         Width           =   6500
      End
      Begin VB.TextBox TB8 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1830
         TabIndex        =   9
         Top             =   4800
         Width           =   6500
      End
      Begin VB.TextBox TB9 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1830
         TabIndex        =   10
         Top             =   5160
         Width           =   6500
      End
      Begin VB.TextBox TB1 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1830
         TabIndex        =   2
         Top             =   1320
         Width           =   6500
      End
      Begin VB.TextBox TB2 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1830
         TabIndex        =   3
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox CB1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   1800
         TabIndex        =   0
         Top             =   990
         Width           =   200
      End
      Begin VB.CheckBox CB2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   3000
         TabIndex        =   1
         Top             =   990
         Width           =   200
      End
      Begin VB.Label L12_Text 
         Caption         =   "L12_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   480
         TabIndex        =   61
         Top             =   4080
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. ID GST :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   60
         Top             =   2070
         Width           =   1665
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Pendaftaran :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   59
         Top             =   2430
         Width           =   1665
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Telefon (O) :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   58
         Top             =   2790
         Width           =   1665
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Telefon (HP) :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   57
         Top             =   3120
         Width           =   1665
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   56
         Top             =   3480
         Width           =   1665
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Bank :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   55
         Top             =   4830
         Width           =   1665
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Akaun :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   54
         Top             =   5190
         Width           =   1665
      End
      Begin VB.Label L19_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Supplier * :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   53
         Top             =   1350
         Width           =   1665
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kod Supplier * :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   52
         Top             =   1710
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "** Sila gunakan 2 abjad SAHAJA dalam ruangan ini."
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
         Height          =   285
         Left            =   3645
         TabIndex        =   51
         Top             =   1770
         Width           =   4650
      End
      Begin VB.Label Label54 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier        Agen / Kedai"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   2085
         TabIndex        =   50
         Top             =   960
         Width           =   3690
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila buat pilihan pendaftaran supplier atau agen. Kemudian masukkan ruangan di bawah bagi pendaftaran supplier/agen tersebut."
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   240
         TabIndex        =   49
         Top             =   360
         Width           =   8145
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pendaftaran Supplier"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   10800
      TabIndex        =   67
      Top             =   0
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CommandButton CMD4 
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
         Height          =   855
         Left            =   2520
         MouseIcon       =   "Frm95.frx":9F9C
         MousePointer    =   99  'Custom
         Picture         =   "Frm95.frx":A2A6
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Simpan data"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CommandButton CMD6 
         BackColor       =   &H80000003&
         Caption         =   "Batal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3480
         MouseIcon       =   "Frm95.frx":AC50
         MousePointer    =   99  'Custom
         Picture         =   "Frm95.frx":AF5A
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Batal"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton CMD5 
         BackColor       =   &H80000003&
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
         Height          =   855
         Left            =   1440
         MouseIcon       =   "Frm95.frx":C024
         MousePointer    =   99  'Custom
         Picture         =   "Frm95.frx":C32E
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Simpan data"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox TB16 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2760
         TabIndex        =   17
         Text            =   "TB16"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TB17 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3285
         MaxLength       =   10
         TabIndex        =   18
         Text            =   "TB17"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox TB11 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   15
         Text            =   "TB11"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox TB10 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1860
         TabIndex        =   14
         Text            =   "TB10"
         Top             =   720
         Width           =   3300
      End
      Begin VB.TextBox TB18 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   16
         Text            =   "TB18"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Kadar Tukaran Purity 999.9  * :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   0
         TabIndex        =   75
         Top             =   3030
         Visible         =   0   'False
         Width           =   3105
      End
      Begin VB.Label L13_Text 
         Caption         =   "L13_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5640
         TabIndex        =   74
         Top             =   360
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Kadar Pemalar Harga Trade In* :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   360
         TabIndex        =   73
         Top             =   1830
         Width           =   3105
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "** Sila gunakan maksimum 4 abjad SAHAJA dalam ruangan ini."
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
         Height          =   525
         Left            =   3480
         OLEDropMode     =   1  'Manual
         TabIndex        =   72
         Top             =   1080
         Width           =   3705
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kod Purity * :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   71
         Top             =   1110
         Width           =   1665
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Purity * :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   70
         Top             =   750
         Width           =   1665
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Assay * :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   69
         Top             =   1470
         Width           =   1665
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan ruangan di bawah bagi pendaftaran purity."
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   68
         Top             =   360
         Width           =   6705
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Senarai Kategori"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11775
      Left            =   6960
      TabIndex        =   87
      Top             =   600
      Visible         =   0   'False
      Width           =   11535
      Begin MSComctlLib.ListView LV3 
         Height          =   10725
         Left            =   120
         TabIndex        =   88
         Top             =   360
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   18918
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
         NumItems        =   0
      End
      Begin VB.Label L28_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L28_Text"
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
         Left            =   2235
         TabIndex        =   91
         Top             =   11400
         Width           =   1215
      End
      Begin VB.Label L27_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L27_Text"
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
         Left            =   2235
         TabIndex        =   90
         Top             =   11175
         Width           =   1215
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan Aktif               :    Bilangan Tidak Aktif    :"
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
         Height          =   660
         Left            =   240
         TabIndex        =   89
         Top             =   11160
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Senarai Supplier"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11775
      Left            =   13080
      TabIndex        =   62
      Top             =   120
      Visible         =   0   'False
      Width           =   17775
      Begin MSComctlLib.ListView LV1 
         Height          =   10725
         Left            =   120
         TabIndex        =   63
         Top             =   360
         Width           =   17475
         _ExtentX        =   30824
         _ExtentY        =   18918
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
         NumItems        =   0
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan Aktif               :    Bilangan Tidak Aktif    :"
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
         Height          =   660
         Left            =   240
         TabIndex        =   66
         Top             =   11280
         Width           =   2295
      End
      Begin VB.Label L23_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L23_Text"
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
         Left            =   2235
         TabIndex        =   65
         Top             =   11295
         Width           =   1215
      End
      Begin VB.Label L24_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L24_Text"
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
         Left            =   2235
         TabIndex        =   64
         Top             =   11490
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Senarai Purity"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11775
      Left            =   16080
      TabIndex        =   76
      Top             =   1200
      Visible         =   0   'False
      Width           =   11655
      Begin MSComctlLib.ListView LV2 
         Height          =   10725
         Left            =   240
         TabIndex        =   77
         Top             =   360
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   18918
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
         NumItems        =   0
      End
      Begin VB.Label L26_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L26_Text"
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
         Left            =   2235
         TabIndex        =   80
         Top             =   11490
         Width           =   1215
      End
      Begin VB.Label L25_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L25_Text"
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
         Left            =   2235
         TabIndex        =   79
         Top             =   11295
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan Aktif               :    Bilangan Tidak Aktif    :"
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
         Height          =   660
         Left            =   240
         TabIndex        =   78
         Top             =   11280
         Width           =   2295
      End
   End
   Begin VB.Label L16_Text 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pendaftaran Tukang Emas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      MouseIcon       =   "Frm95.frx":CCD8
      MousePointer    =   99  'Custom
      TabIndex        =   40
      Top             =   3000
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label L17_Text 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai Tukang Emas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      MouseIcon       =   "Frm95.frx":CFE2
      MousePointer    =   99  'Custom
      TabIndex        =   39
      Top             =   3240
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label L10_Text 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai Dulang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      MouseIcon       =   "Frm95.frx":D2EC
      MousePointer    =   99  'Custom
      TabIndex        =   38
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label L9_Text 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pendaftaran Dulang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      MouseIcon       =   "Frm95.frx":D5F6
      MousePointer    =   99  'Custom
      TabIndex        =   37
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label L8_Text 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai Kategori Produk"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      MouseIcon       =   "Frm95.frx":D900
      MousePointer    =   99  'Custom
      TabIndex        =   36
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label L7_Text 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pendaftaran Kategori Produk"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      MouseIcon       =   "Frm95.frx":DC0A
      MousePointer    =   99  'Custom
      TabIndex        =   35
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label L6_Text 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai Purity"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      MouseIcon       =   "Frm95.frx":DF14
      MousePointer    =   99  'Custom
      TabIndex        =   34
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label L5_Text 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pendaftaran Purity"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      MouseIcon       =   "Frm95.frx":E21E
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label L4_Text 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai Supplier"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      MouseIcon       =   "Frm95.frx":E528
      MousePointer    =   99  'Custom
      TabIndex        =   32
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label L3_Text 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pendaftaran Supplier/Agen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      MouseIcon       =   "Frm95.frx":E832
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   120
      Width           =   2775
   End
   Begin VB.Menu Frm95_PM_Menu 
      Caption         =   "Supplier"
      Visible         =   0   'False
      Begin VB.Menu Frm95_SM_edit 
         Caption         =   "Lihat / Edit Data"
      End
      Begin VB.Menu Frm95_SM_ubah_status1 
         Caption         =   "Ubah status"
         Begin VB.Menu Frm95_SSM_aktif 
            Caption         =   "Aktif"
         End
         Begin VB.Menu Frm95_SSM_tidak_aktif 
            Caption         =   "Tidak aktif"
         End
      End
   End
   Begin VB.Menu Frm95_PM_Menu2 
      Caption         =   "Purity"
      Visible         =   0   'False
      Begin VB.Menu Frm95_SM_edit_purity 
         Caption         =   "Lihat / Edit Data"
      End
      Begin VB.Menu Frm95_SM_ubah_status2 
         Caption         =   "Ubah status"
         Begin VB.Menu Frm95_SSM_aktif2 
            Caption         =   "Aktif"
         End
         Begin VB.Menu Frm95_SSM_tidak_aktif2 
            Caption         =   "Tidak aktif"
         End
      End
   End
   Begin VB.Menu Frm95_PM_Menu3 
      Caption         =   "Produk"
      Visible         =   0   'False
      Begin VB.Menu Frm95_SM_edit_produk 
         Caption         =   "Lihat / Edit Data"
      End
      Begin VB.Menu Frm95_SM_ubah_status3 
         Caption         =   "Ubah status"
         Begin VB.Menu Frm95_SSM_aktif3 
            Caption         =   "Aktif"
         End
         Begin VB.Menu Frm95_SSM_tidak_aktif3 
            Caption         =   "Tidak aktif"
         End
      End
   End
   Begin VB.Menu Frm95_PM_Menu4 
      Caption         =   "Dulang"
      Visible         =   0   'False
      Begin VB.Menu Frm95_SM_edit_dulang 
         Caption         =   "Lihat / Edit Data"
      End
      Begin VB.Menu Frm95_SM_ubah_status4 
         Caption         =   "Ubah status"
         Begin VB.Menu Frm95_SSM_aktif4 
            Caption         =   "Aktif"
         End
         Begin VB.Menu Frm95_SSM_tidak_aktif4 
            Caption         =   "Tidak aktif"
         End
      End
   End
   Begin VB.Menu Frm95_PM_Menu5 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm95_SM_edit_tukang 
         Caption         =   "Lihat / Edit Data"
      End
      Begin VB.Menu Frm95_SM_ubah_status5 
         Caption         =   "Ubah status"
         Begin VB.Menu Frm95_SSM_aktif5 
            Caption         =   "Aktif"
         End
         Begin VB.Menu Frm95_SSM_tidak_aktif5 
            Caption         =   "Tidak aktif"
         End
      End
   End
End
Attribute VB_Name = "Frm95"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB1_Click()
'On Error Resume Next
If Frm95.CB1 = 1 Then
    Frm95.CB2 = 0
    
    Frm95.L19_Text = "Nama Supplier * :"
    
    Frm95.TB2.Locked = False
    Frm95.TB2 = vbNullString
    Frm95.TB2.BackColor = &HFFFFFF
End If
End Sub
Private Sub CB2_Click()
'On Error Resume Next
If Frm95.CB2 = 1 Then
    Frm95.CB1 = 0
    
    Frm95.L19_Text = "Nama Agen * :"
    
    Frm95.TB2.Locked = True
    Frm95.TB2 = vbNullString
    Frm95.TB2.BackColor = &H8000000A
End If
End Sub
Private Sub CMD1_Click()
'On Error Resume Next
Dim Err(10)
DATA_SAVE = 0

If Frm95.CB1 = 0 And Frm95.CB2 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan samada pendaftaran bagi supplier atau agen/kedai"
End If
If Frm95.CB1 = 1 Then
    If Frm95.TB1 = vbNullString Then
        x = x + 1
        Err(x) = "Sila masukkan nama supplier."
    End If
    If InStr(1, Frm95.TB1, "*") <> 0 Or InStr(1, Frm95.TB1, "/") <> 0 Or InStr(1, Frm95.TB1, "\") <> 0 Or InStr(1, Frm95.TB1, "'") <> 0 Then
        x = x + 1
        Err(x) = "Nama supplier mempunyai simbol yang tidak sah."
    End If
    If Frm95.TB2 = vbNullString Then
        x = x + 1
        Err(x) = "Sila masukkan kod supplier."
    End If
    If InStr(1, Frm95.TB2, "*") <> 0 Or InStr(1, Frm95.TB2, "/") <> 0 Or InStr(1, Frm95.TB2, "\") <> 0 Or InStr(1, Frm95.TB2, "'") <> 0 Then
        x = x + 1
        Err(x) = "Kod supplier mempunyai simbol yang tidak sah."
    End If
End If
If Frm95.CB2 = 1 Then
    If Frm95.TB1 = vbNullString Then
        x = x + 1
        Err(x) = "Sila masukkan nama agen/kedai."
    End If
    If InStr(1, Frm95.TB1, "*") <> 0 Or InStr(1, Frm95.TB1, "/") <> 0 Or InStr(1, Frm95.TB1, "\") <> 0 Or InStr(1, Frm95.TB1, "'") <> 0 Then
        x = x + 1
        Err(x) = "Nama agen/kedai mempunyai simbol yang tidak sah."
    End If
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah anda ingin simpan data ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
        If UCase(Frm95.TB1) = "TRADE IN" Then
        
            MsgBox "Nama Supplier ini tidak dibenarkan untuk didaftarkan. [Trade In] dikhaskan untuk barang trade in sahaja.", vbExclamation, "Info"
            
            Frm95.TB1.SetFocus
            Exit Sub
            
        End If
        
        If UCase(Frm95.TB2) = "TI" Then
        
            MsgBox "Kod Supplier ini tidak dibenarkan untuk didaftarkan. [TI] dikhaskan untuk kod barang trade in sahaja.", vbExclamation, "Info"
            
            Frm95.TB1.SetFocus
            Exit Sub
        End If

'#### Periksa Samada Nama Supplier Wujud Atau Tidak #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where Supplier='" & UCase(Frm95.TB1) & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            rs.Close
            Set rs = Nothing
            
            MsgBox "Nama Supplier [" & UCase(Frm95.TB1) & "] telah didaftarkan sebelum ini." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila periksa data anda.", vbInformation, "Info"
             
            Frm95.TB1.SetFocus
            
            Exit Sub
            
        End If
        
        rs.Close
        Set rs = Nothing
'#### Periksa Samada Nama Supplier Wujud Atau Tidak #### - End

'#### Periksa Samada Kod Supplier Wujud Atau Tidak #### - Start
        If Frm95.CB1 = 1 Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from setting_database where kod_supplier='" & UCase(Frm95.TB2) & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
            
                rs.Close
                Set rs = Nothing
                
                MsgBox "Kod Supplier [" & UCase(Frm95.TB2) & "] telah didaftarkan sebelum ini." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Sila periksa data anda.", vbInformation, "Info"
                       
                Frm95.TB1.SetFocus
                
                Exit Sub
            End If
            
            rs.Close
            Set rs = Nothing
            
        End If
'#### Periksa Samada Kod Supplier Wujud Atau Tidak #### - End
        
        LM_NOW = Now
        
'#### Simpan Data Supplier #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database", cn, adOpenKeyset, adLockOptimistic

        rs.AddNew
        If Frm95.CB1 = 1 Then
            rs!jenis_supplier = "Supplier"
        ElseIf Frm95.CB2 = 1 Then
            rs!jenis_supplier = "Agen/Kedai"
        End If
        If Frm95.TB1 <> vbNullString Then 'Nama Supplier
            rs!supplier = UCase(Frm95.TB1)
        Else
            rs!supplier = Null
        End If
        If Frm95.TB2 <> vbNullString Then 'Kod Supplier
            rs!Kod_Supplier = UCase(Frm95.TB2)
        Else
            rs!Kod_Supplier = Null
        End If
        If Frm95.TB3 <> vbNullString Then 'No. ID GST
            rs!no_id_gst = UCase(Frm95.TB3)
        Else
            rs!no_id_gst = Null
        End If
        If Frm95.TB4 <> vbNullString Then 'No. Pendaftaran Syarikat
            rs!no_pendaftaran = UCase(Frm95.TB4)
        Else
            rs!no_pendaftaran = Null
        End If
        If Frm95.TB5 <> vbNullString Then 'No. Telefon Office
            rs!no_tel_off = UCase(Frm95.TB5)
        Else
            rs!no_tel_off = Null
        End If
        If Frm95.TB6 <> vbNullString Then 'No. Telefon HP
            rs!no_tel_hp = UCase(Frm95.TB6)
        Else
            rs!no_tel_hp = Null
        End If
        If Frm95.TB7 <> vbNullString Then 'Alamat
            rs!alamat = UCase(Frm95.TB7)
        Else
            rs!alamat = Null
        End If
        If Frm95.TB8 <> vbNullString Then 'Nama Bank
            rs!nama_bank = UCase(Frm95.TB8)
        Else
            rs!nama_bank = Null
        End If
        If Frm95.TB9 <> vbNullString Then 'No. Akaun
            rs!no_akaun = UCase(Frm95.TB9)
        Else
            rs!no_akaun = Null
        End If
        rs!Status = 1 '0 : Tidak aktif , 1 : Aktif
        rs.Update
        
        rs.Close
        Set rs = Nothing
'#### Simpan Data Supplier #### - End

'#### Update Log Aktiviti Sistem #### - Start
        user = MDI_frm1.L3_Text
        LogAct_Memory = "[" & user & "] Daftar supplier [" & UCase(Frm95.TB1) & "]."
        LogDate_Memory = LM_NOW
        Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End
        
        Call Frm95_initial
        
        MsgBox "Data telah berjaya disimpan.", vbInformation, "Info"
        
        Frm95.TB1.SetFocus
        
    End If
End If
End Sub
Private Sub CMD10_Click()
'On Error Resume Next
Dim Err(3)
DATA_SAVE = 0

If Frm95.TB14 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan Nama Dulang."
End If
If Frm95.TB14 <> vbNullString Then
    If InStr(1, Frm95.TB14, "*") <> 0 Or InStr(1, Frm95.TB14, "/") <> 0 Or InStr(1, Frm95.TB14, "\") <> 0 Or InStr(1, Frm95.TB14, "'") <> 0 Then
        x = x + 1
        Err(x) = "Dulang mempunyai simbol yang tidak sah."
    End If
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah Anda Ingin Simpan Data Ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then

'#### Periksa Samada Nama Dulang Wujud Atau Tidak #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where SenaraiDulang='" & UCase(Frm95.TB14) & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            rs.Close
            Set rs = Nothing
            
            MsgBox "Nama Dulang [" & UCase(Frm95.TB14) & "] Telah Didaftarkan Sebelum Ini." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila Periksa Data Anda.", vbInformation, "Info"
                    
            Exit Sub
        End If
        
        rs.Close
        Set rs = Nothing
'#### Periksa Samada Nama Dulang Wujud Atau Tidak #### - End

'#### Simpan Data Dulang #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database", cn, adOpenKeyset, adLockOptimistic

        rs.AddNew
        If Frm95.TB14 <> vbNullString Then 'Nama Dulang
            rs!SenaraiDulang = UCase(Frm95.TB14)
        Else
            rs!SenaraiDulang = Null
        End If
        rs!Status = 1 '0 : Tidak aktif , 1 : Aktif
        rs.Update
        
        rs.Close
        Set rs = Nothing
'#### Simpan Data Dulang #### - End

'#### Update Log Aktiviti Sistem #### - Start
        user = MDI_frm1.L3_Text
        LogAct_Memory = "[" & user & "] Daftar Dulang [" & UCase(Frm95.TB14) & "]."
        LogDate_Memory = DateTime.Date & " " & DateTime.Time$
        Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End
        
        Call Frm95_initial
        
        MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
        
        Frm95.TB14.SetFocus
    End If
End If
End Sub
Private Sub CMD11_Click()
'On Error Resume Next
Dim Err(3)
DATA_SAVE = 0
Frm95_LM_DULANG = vbNullString

If Frm95.TB14 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan Nama Dulang."
End If
If Frm95.TB14 <> vbNullString Then
    If InStr(1, Frm95.TB14, "*") <> 0 Or InStr(1, Frm95.TB14, "/") <> 0 Or InStr(1, Frm95.TB14, "\") <> 0 Or InStr(1, Frm95.TB14, "'") <> 0 Then
        x = x + 1
        Err(x) = "Dulang mempunyai simbol yang tidak sah."
    End If
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah Anda Ingin Ubah Data Ini ?" & vbCrLf & _
            vbNullString & vbCrLf & _
            "Sistem Akan Mengambil Sedikit Masa Untuk Update Semua Data Di Dalam Sistem." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then

'#### Periksa Samada Nama Dulang Wujud Atau Tidak #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where SenaraiDulang='" & UCase(Frm95.TB14) & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If rs!ID <> Frm95.L15_Text Then
                rs.Close
                Set rs = Nothing
                
                MsgBox "Nama Dulang [" & UCase(Frm95.TB14) & "] Telah Didaftarkan Sebelum Ini." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Sila Periksa Data Anda.", vbInformation, "Info"
                        
                Exit Sub
            End If
        End If
        
        rs.Close
        Set rs = Nothing
'#### Periksa Samada Nama Dulang Wujud Atau Tidak #### - End

'#### Memorize Data Asal #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where ID='" & Frm95.L15_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!SenaraiDulang) Then Frm95_LM_DULANG = rs!SenaraiDulang 'Nama Dulang

        End If
        
        rs.Close
        Set rs = Nothing
'#### Memorize Data Asal #### - End

'#### Simpan Data Produk #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where ID='" & Frm95.L15_Text & "'", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            If Frm95.TB14 <> vbNullString Then 'Nama Dulang
                rs!SenaraiDulang = UCase(Frm95.TB14)
            Else
                rs!SenaraiDulang = Null
            End If
            rs.Update
            
            DATA_SAVE = 1
        End If
        
        rs.Close
        Set rs = Nothing
'#### Simpan Data Produk #### - End
        
        If DATA_SAVE = 1 Then

'#### Update Nama Dulang #### - Start
            If Frm95_LM_DULANG <> UCase(Frm95.TB14) Then
                
                '#### Update Maklumat Dulang Dalam Table Data_Database #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE Data_Database set Dulang='" & UCase(Frm95.TB14) & "'" _
                & "WHERE Dulang='" & Frm95_LM_DULANG & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Dulang Dalam Table Data_Database #### - End
                
                '#### Update Maklumat Dulang Dalam Table 12_gold_bar_database #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE 12_gold_bar_database set dulang='" & UCase(Frm95.TB14) & "'" _
                & "WHERE dulang='" & Frm95_LM_DULANG & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Dulang Dalam Table 12_gold_bar_database #### - End
                
                '#### Update Maklumat Dulang Dalam Table 23_senarai_jualan #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE 23_senarai_jualan set dulang='" & UCase(Frm95.TB14) & "'" _
                & "WHERE dulang='" & Frm95_LM_DULANG & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Dulang Dalam Table 23_senarai_jualan #### - End
                
                '#### Update Maklumat Dulang Dalam Table 27_senarai_ansuran #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE 27_senarai_ansuran set dulang='" & UCase(Frm95.TB14) & "'" _
                & "WHERE dulang='" & Frm95_LM_DULANG & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Dulang Dalam Table 27_senarai_ansuran #### - End
                
                '#### Update Maklumat Dulang Dalam Table 42_tempahan_siap #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE 42_tempahan_siap set dulang='" & UCase(Frm95.TB14) & "'" _
                & "WHERE dulang='" & Frm95_LM_DULANG & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Dulang Dalam Table 42_tempahan_siap #### - End
                
            End If
'#### Update Nama & Kod Produk #### - End
        
        
'#### Update Log Aktiviti Sistem #### - Start
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & user & "] Edit Dulang [" & UCase(Frm95.TB14) & "]."
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End
            
            Call Frm95_initial
            
            MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
        End If
    End If
End If
End Sub
Private Sub CMD12_Click()
'on error resume next
Frm95.Frame7.Visible = False
Frm95.Frame8.Visible = True
End Sub
Private Sub CMD13_Click()
'On Error Resume Next
Dim Err(3)
DATA_SAVE = 0

If Frm95.TB15 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan Nama Tukang Emas."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah Anda Ingin Simpan Data Ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then

'#### Periksa Samada Nama Tukang Emas Wujud Atau Tidak #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where tukang_emas='" & UCase(Frm95.TB15) & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            rs.Close
            Set rs = Nothing
            
            MsgBox "Nama Tukang Emas [" & UCase(Frm95.TB15) & "] Telah Didaftarkan Sebelum Ini." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila Periksa Data Anda.", vbInformation, "Info"
                    
            Exit Sub
        End If
        
        rs.Close
        Set rs = Nothing
'#### Periksa Samada Nama Tukang Emas Wujud Atau Tidak #### - End

'#### Simpan Data Tukang Emas #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database", cn, adOpenKeyset, adLockOptimistic

        rs.AddNew
        If Frm95.TB15 <> vbNullString Then 'Nama Tukang Emas
            rs!tukang_emas = UCase(Frm95.TB15)
        Else
            rs!tukang_emas = Null
        End If
        rs.Update
        
        rs.Close
        Set rs = Nothing
'#### Simpan Data Tukang Emas #### - End

'#### Update Log Aktiviti Sistem #### - Start
        user = MDI_frm1.L3_Text
        LogAct_Memory = "[" & user & "] Daftar Tukang Emas [" & UCase(Frm95.TB15) & "]."
        LogDate_Memory = DateTime.Date & " " & DateTime.Time$
        Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End
        
        Call Frm95_initial
        
        MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
        
        Frm95.TB15.SetFocus
    End If
End If
End Sub
Private Sub CMD14_Click()
'On Error Resume Next
Dim Err(3)
DATA_SAVE = 0
Frm95_LM_TUKANG = vbNullString

If Frm95.TB15 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan Nama Tukang Emas."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah Anda Ingin Ubah Data Ini ?" & vbCrLf & _
            vbNullString & vbCrLf & _
            "Sistem Akan Mengambil Sedikit Masa Untuk Update Semua Data Di Dalam Sistem." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then

'#### Periksa Samada Nama Tukang Emas Wujud Atau Tidak #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where tukang_emas='" & UCase(Frm95.TB15) & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If rs!ID <> Frm95.L18_Text Then
                rs.Close
                Set rs = Nothing
                
                MsgBox "Nama Tukang Emas [" & UCase(Frm95.TB15) & "] Telah Didaftarkan Sebelum Ini." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Sila Periksa Data Anda.", vbInformation, "Info"
                        
                Exit Sub
            End If
        End If
        
        rs.Close
        Set rs = Nothing
'#### Periksa Samada Nama Tukang Emas Wujud Atau Tidak #### - End

'#### Memorize Data Asal #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where ID='" & Frm95.L18_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!tukang_emas) Then Frm95_LM_TUKANG = rs!tukang_emas 'Nama Tukang EMas

        End If
        
        rs.Close
        Set rs = Nothing
'#### Memorize Data Asal #### - End

'#### Simpan Data Nama Tukang Emas #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where ID='" & Frm95.L18_Text & "'", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            If Frm95.TB15 <> vbNullString Then 'Nama Tukang Emas
                rs!tukang_emas = UCase(Frm95.TB15)
            Else
                rs!tukang_emas = Null
            End If
            rs.Update
            
            DATA_SAVE = 1
        End If
        
        rs.Close
        Set rs = Nothing
'#### Simpan Data Produk #### - End
        
        If DATA_SAVE = 1 Then

'#### Update Nama Dulang #### - Start
            If Frm95_LM_TUKANG <> UCase(Frm95.TB15) Then
                
                '#### Update Maklumat Tukang Emas Dalam Table 46_tempahan_tukang_emas #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE 46_tempahan_tukang_emas set tukang_emas='" & UCase(Frm95.TB15) & "'" _
                & "WHERE tukang_emas='" & Frm95_LM_TUKANG & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Tukang Emas Dalam Table 46_tempahan_tukang_emas #### - End
                
                
            End If
'#### Update Nama Nama Tukang Emas #### - End
        
        
'#### Update Log Aktiviti Sistem #### - Start
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & user & "] Edit Tukang Emas [" & UCase(Frm95.TB15) & "]."
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End
            
            Call Frm95_initial
            
            MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
        End If
    End If
End If
End Sub
Private Sub CMD2_Click()
'On Error Resume Next
Dim Err(10)
Dim strsql As String

DATA_SAVE = 0
Frm95_LM_SUPPLIER = vbNullString
Frm95_LM_KOD_SUPPLIER = vbNullString
Frm95_LM_ID_GST = vbNullString

If Frm95.CB1 = 0 And Frm95.CB2 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan samada pendaftaran bagi supplier atau agen/kedai"
End If
If Frm95.CB1 = 1 Then
    If Frm95.TB1 = vbNullString Then
        x = x + 1
        Err(x) = "Sila masukkan nama supplier."
    End If
    If InStr(1, Frm95.TB1, "*") <> 0 Or InStr(1, Frm95.TB1, "/") <> 0 Or InStr(1, Frm95.TB1, "\") <> 0 Or InStr(1, Frm95.TB1, "'") <> 0 Then
        x = x + 1
        Err(x) = "Nama supplier mempunyai simbol yang tidak sah."
    End If
    If Frm95.TB2 = vbNullString Then
        x = x + 1
        Err(x) = "Sila masukkan kod supplier."
    End If
    If InStr(1, Frm95.TB2, "*") <> 0 Or InStr(1, Frm95.TB2, "/") <> 0 Or InStr(1, Frm95.TB2, "\") <> 0 Or InStr(1, Frm95.TB2, "'") <> 0 Then
        x = x + 1
        Err(x) = "Kod supplier mempunyai simbol yang tidak sah."
    End If
End If
If Frm95.CB2 = 1 Then
    If Frm95.TB1 = vbNullString Then
        x = x + 1
        Err(x) = "Sila masukkan nama agen/kedai."
    End If
    If InStr(1, Frm95.TB1, "*") <> 0 Or InStr(1, Frm95.TB1, "/") <> 0 Or InStr(1, Frm95.TB1, "\") <> 0 Or InStr(1, Frm95.TB1, "'") <> 0 Then
        x = x + 1
        Err(x) = "Nama agen/kedai mempunyai simbol yang tidak sah."
    End If
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah anda ingin ubah maklumat berkenaan supplier ini?" & vbCrLf & _
            vbNullString & vbCrLf & _
            "Sistem mungkin akan mengambil sedikit masa untuk update semua data di dalam sistem." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
        If UCase(Frm95.TB1) = "TRADE IN" Then
            MsgBox "Nama Supplier Ini Tidak Dibenarkan Untuk Didaftarkan. [Trade In] Dikhaskan Untuk Barang Trade In.", vbExclamation, "Info"
            Exit Sub
        End If
        
        If UCase(Frm95.TB2) = "TI" Then
            MsgBox "Kod Supplier Ini Tidak Dibenarkan Untuk Didaftarkan. [TI] Dikhaskan Untuk Kod Barang Trade In.", vbExclamation, "Info"
            Exit Sub
        End If

'#### Periksa Samada Nama Supplier Wujud Atau Tidak #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where Supplier='" & UCase(Frm95.TB1) & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If rs!ID <> Frm95.L12_Text Then
            
                rs.Close
                Set rs = Nothing
                
                MsgBox "Nama Supplier [" & UCase(Frm95.TB1) & "] Telah Didaftarkan Sebelum Ini." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Sila Periksa Data Anda.", vbInformation, "Info"
                        
                Exit Sub
                
            End If
        End If
        
        rs.Close
        Set rs = Nothing
'#### Periksa Samada Nama Supplier Wujud Atau Tidak #### - End

'#### Periksa Samada Kod Supplier Wujud Atau Tidak #### - Start
        If Frm95.CB1 = 1 Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from setting_database where kod_supplier='" & UCase(Frm95.TB2) & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If rs!ID <> Frm95.L12_Text Then
                
                    rs.Close
                    Set rs = Nothing
                    
                    MsgBox "Kod Supplier [" & UCase(Frm95.TB2) & "] Telah Didaftarkan Sebelum Ini." & vbCrLf & _
                            vbNullString & vbCrLf & _
                            "Sila Periksa Data Anda.", vbInformation, "Info"
                            
                    Exit Sub
                    
                End If
            End If
            
            rs.Close
            Set rs = Nothing
            
        End If
'#### Periksa Samada Kod Supplier Wujud Atau Tidak #### - End

'#### Memorize Data Asal #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where ID='" & Frm95.L12_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!supplier) Then Frm95_LM_SUPPLIER = rs!supplier 'Nama Supplier
            If Not IsNull(rs!Kod_Supplier) Then Frm95_LM_KOD_SUPPLIER = rs!Kod_Supplier 'Kod Supplier
            If Not IsNull(rs!no_id_gst) Then Frm95_LM_ID_GST = rs!no_id_gst 'No. ID GST

        End If
        
        rs.Close
        Set rs = Nothing
'#### Memorize Data Asal #### - End
        
        LM_NOW = Now
        
'#### Simpan Data Supplier #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where ID='" & Frm95.L12_Text & "'", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
        
            If Frm95.CB1 = 1 Then
                rs!jenis_supplier = "Supplier"
            ElseIf Frm95.CB2 = 1 Then
                rs!jenis_supplier = "Agen/Kedai"
            End If
            If Frm95.TB1 <> vbNullString Then 'Nama Supplier
                rs!supplier = UCase(Frm95.TB1)
            Else
                rs!supplier = Null
            End If
            If Frm95.TB2 <> vbNullString Then 'Kod Supplier
                rs!Kod_Supplier = UCase(Frm95.TB2)
            Else
                rs!Kod_Supplier = Null
            End If
            If Frm95.TB3 <> vbNullString Then 'No. ID GST
                rs!no_id_gst = UCase(Frm95.TB3)
            Else
                rs!no_id_gst = Null
            End If
            If Frm95.TB4 <> vbNullString Then 'No. Pendaftaran Syarikat
                rs!no_pendaftaran = UCase(Frm95.TB4)
            Else
                rs!no_pendaftaran = Null
            End If
            If Frm95.TB5 <> vbNullString Then 'No. Telefon Office
                rs!no_tel_off = UCase(Frm95.TB5)
            Else
                rs!no_tel_off = Null
            End If
            If Frm95.TB6 <> vbNullString Then 'No. Telefon HP
                rs!no_tel_hp = UCase(Frm95.TB6)
            Else
                rs!no_tel_hp = Null
            End If
            If Frm95.TB7 <> vbNullString Then 'Alamat
                rs!alamat = UCase(Frm95.TB7)
            Else
                rs!alamat = Null
            End If
            If Frm95.TB8 <> vbNullString Then 'Nama Bank
                rs!nama_bank = UCase(Frm95.TB8)
            Else
                rs!nama_bank = Null
            End If
            If Frm95.TB9 <> vbNullString Then 'No. Akaun
                rs!no_akaun = UCase(Frm95.TB9)
            Else
                rs!no_akaun = Null
            End If
            rs.Update
            DATA_SAVE = 1
        End If
        
        rs.Close
        Set rs = Nothing
'#### Simpan Data Supplier #### - End
        
        If DATA_SAVE = 1 Then
            
'#### Update Nama & Kod Supplier & ID GST #### - Start
            If Frm95_LM_SUPPLIER <> UCase(Frm95.TB1) Or Frm95_LM_KOD_SUPPLIER <> UCase(Frm95.TB2) Then
                
                '#### Update Maklumat Supplier Dalam Table Data_Database #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE Data_Database set nama_Supplier='" & UCase(Frm95.TB1) & "'," _
                & "Kod_Supplier='" & UCase(Frm95.TB2) & "'" _
                & "WHERE nama_Supplier='" & Frm95_LM_SUPPLIER & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Supplier Dalam Table Data_Database #### - End
                
                '#### Update Maklumat Supplier Dalam Table 12_gold_bar_database #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE 12_gold_bar_database set supplier='" & UCase(Frm95.TB1) & "'," _
                & "kod_supplier='" & UCase(Frm95.TB2) & "'" _
                & "WHERE supplier='" & Frm95_LM_SUPPLIER & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Supplier Dalam Table 12_gold_bar_database #### - End
                
                '#### Update Maklumat Supplier Dalam Table 77_gdn_grn #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE 77_gdn_grn set supplier_agen='" & UCase(Frm95.TB1) & "' WHERE supplier_agen='" & Frm95_LM_SUPPLIER & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Supplier Dalam Table 77_gdn_grn #### - End
                
                '#### Update Maklumat Supplier Dalam Table 12_gold_bar_database #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE 39_akaun_expense set nama_kedai='" & UCase(Frm95.TB1) & "' WHERE nama_kedai='" & Frm95_LM_SUPPLIER & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Supplier Dalam Table 12_gold_bar_database #### - End
                
            End If
            
            '#### Update ID GST #### - Start
            If Frm95_LM_ID_GST <> UCase(Frm95.TB3) Then
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE 16_gold_bar_belian set no_id_gst_supplier='" & UCase(Frm95.TB3) & "'" _
                & "WHERE kod_supplier='" & Frm95_LM_KOD_SUPPLIER & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                
                '#### Update Maklumat No ID GST Supplier Dalam Table Data_Database #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE Data_Database set no_id_gst='" & UCase(Frm95.TB3) & "'" _
                & "WHERE kod_supplier='" & Frm95_LM_KOD_SUPPLIER & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat No ID GST Supplier Dalam Table Data_Database #### - End
                
                '#### Update Maklumat Supplier Dalam Table 12_gold_bar_database #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE 39_akaun_expense set no_id_gst='" & UCase(Frm95.TB3) & "' WHERE no_id_gst='" & Frm95_LM_ID_GST & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Supplier Dalam Table 12_gold_bar_database #### - End
                
            End If
            '#### Update ID GST #### - End
            
            '### Update nama kedai / supplier dalam table forming out ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            
            strsql = "UPDATE 57_form_out set nama_kedai='" & UCase(Frm95.TB1) & "'" _
            & "WHERE id_kedai='" & Frm95.L12_Text & "'"
            
            Set rs = cn.Execute(strsql)
            Set rs = Nothing
            '### Update nama kedai / supplier dalam table forming out ### - End
            
'#### Update Nama & Kod Supplier #### - End

'#### Update ID GST #### - Start
        
'#### Update Log Aktiviti Sistem #### - Start
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & user & "] Edit data supplier [" & UCase(Frm95.TB1) & "]."
            LogDate_Memory = LM_NOW
            Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End
        
            Call Frm95_initial
            Call Frm95_senarai_supplier_header
            Call Frm95_senarai_supplier
            
            Frm95.Frame1.Visible = False
            Frm95.Frame2.Visible = True
            
            MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
            
        End If
        
    End If
End If

End Sub
Private Sub CMD3_Click()
'on error resume next
Frm95.Frame1.Visible = False
Frm95.Frame2.Visible = True
End Sub
Private Sub CMD4_Click()
'On Error Resume Next
Dim Err(10)
DATA_SAVE = 0

If Frm95.TB10 = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan nama purity."
End If
If Frm95.TB10 <> vbNullString Then
    If InStr(1, Frm95.TB10, "*") <> 0 Or InStr(1, Frm95.TB10, "/") <> 0 Or InStr(1, Frm95.TB10, "\") <> 0 Or InStr(1, Frm95.TB10, "'") <> 0 Then
        x = x + 1
        Err(x) = "Nama purity mempunyai simbol yang tidak sah."
    End If
End If
If Frm95.TB11 = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan kod purity."
End If
If Frm95.TB11 <> vbNullString Then
    If InStr(1, Frm95.TB11, "*") <> 0 Or InStr(1, Frm95.TB11, "/") <> 0 Or InStr(1, Frm95.TB11, "\") <> 0 Or InStr(1, Frm95.TB11, "'") <> 0 Then
        x = x + 1
        Err(x) = "Kod purity mempunyai simbol yang tidak sah."
    End If
End If
If Frm95.TB16 = vbNullString Or (Frm95.TB16 <> vbNullString And Not IsNumeric(Frm95.TB16)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar Tukaran Purity 999.9]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm95.TB17 = vbNullString Or (Frm95.TB17 <> vbNullString And Not IsNumeric(Frm95.TB17)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar Pemalar Harga Trade In]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm95.TB18 = vbNullString Or (Frm95.TB18 <> vbNullString And Not IsNumeric(Frm95.TB18)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Assay]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah anda ingin simpan data ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then

'#### Periksa Samada Nama Purity Wujud Atau Tidak #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where Metal_Purity='" & UCase(Frm95.TB10) & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            rs.Close
            Set rs = Nothing
            
            MsgBox "Nama Purity [" & UCase(Frm95.TB10) & "] Telah Didaftarkan Sebelum Ini." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila Periksa Data Anda.", vbInformation, "Info"
                    
            Exit Sub
        End If
        
        rs.Close
        Set rs = Nothing
'#### Periksa Samada Nama Purity Wujud Atau Tidak #### - End

'#### Periksa Samada Kod Purity Wujud Atau Tidak #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where Kod_Metal_Purity='" & UCase(Frm95.TB11) & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            rs.Close
            Set rs = Nothing
            
            MsgBox "Kod Purity [" & UCase(Frm95.TB11) & "] Telah Didaftarkan Sebelum Ini." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila Periksa Data Anda.", vbInformation, "Info"
                    
            Exit Sub
        End If
        
        rs.Close
        Set rs = Nothing
'#### Periksa Samada Kod Purity Wujud Atau Tidak #### - End

'#### Simpan Data Purity #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database", cn, adOpenKeyset, adLockOptimistic

        rs.AddNew
        If Frm95.TB10 <> vbNullString Then 'Nama Purity
            rs!Metal_Purity = UCase(Frm95.TB10)
        Else
            rs!Metal_Purity = Null
        End If
        If Frm95.TB11 <> vbNullString Then 'Kod Purity
            rs!Kod_Metal_Purity = UCase(Frm95.TB11)
        Else
            rs!Kod_Metal_Purity = Null
        End If
        If Frm95.TB16 <> vbNullString Then 'Kadar Tukaran Purity 999.9
            rs!kadar_tukaran = Format(Frm95.TB16, "0.00")
        Else
            rs!kadar_tukaran = "0.00"
        End If
        If Frm95.TB17 <> vbNullString Then 'Kadar pemalar harga trade in
            rs!trade_in = Frm95.TB17
        Else
            rs!trade_in = "0.00"
        End If
        If Frm95.TB18 <> vbNullString Then 'Assay
            rs!assay = Frm95.TB18
        Else
            rs!assay = "0.00"
        End If
        rs!Status = 1 '0 : Tidak aktif , 1 : Aktif
        rs.Update
        
        rs.Close
        Set rs = Nothing
'#### Simpan Data Supplier #### - End

'#### Update Log Aktiviti Sistem #### - Start
        user = MDI_frm1.L3_Text
        LogAct_Memory = "[" & user & "] Daftar Purity [" & UCase(Frm95.TB10) & "]."
        LogDate_Memory = DateTime.Date & " " & DateTime.Time$
        Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End
        
        Call Frm95_initial
        
        MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
        
        Frm95.TB10.SetFocus
    End If
End If
End Sub
Private Sub CMD5_Click()
'On Error Resume Next
Dim Err(10)

DATA_SAVE = 0
Frm95_LM_PURITY = vbNullString
Frm95_LM_KOD_PURITY = vbNullString

If Frm95.TB10 = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan nama purity."
End If
If InStr(1, Frm95.TB10, "*") <> 0 Or InStr(1, Frm95.TB10, "/") <> 0 Or InStr(1, Frm95.TB10, "\") <> 0 Or InStr(1, Frm95.TB10, "'") <> 0 Then
    x = x + 1
    Err(x) = "Nama purity mempunyai simbol yang tidak sah."
End If
If Frm95.TB11 = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan kod purity."
End If
If InStr(1, Frm95.TB11, "*") <> 0 Or InStr(1, Frm95.TB11, "/") <> 0 Or InStr(1, Frm95.TB11, "\") <> 0 Or InStr(1, Frm95.TB11, "'") <> 0 Then
    x = x + 1
    Err(x) = "Kod purity mempunyai simbol yang tidak sah."
End If
If Frm95.TB16 = vbNullString Or (Frm95.TB16 <> vbNullString And Not IsNumeric(Frm95.TB16)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar Tukaran Purity 999.9]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm95.TB17 = vbNullString Or (Frm95.TB17 <> vbNullString And Not IsNumeric(Frm95.TB17)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar Pemalar Harga Trade In]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm95.TB18 = vbNullString Or (Frm95.TB18 <> vbNullString And Not IsNumeric(Frm95.TB18)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Assay]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah Anda Ingin Ubah Data Ini ?" & vbCrLf & _
            vbNullString & vbCrLf & _
            "Sistem Akan Mengambil Sedikit Masa Untuk Update Semua Data Di Dalam Sistem." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then

'#### Periksa Samada Nama Purity Wujud Atau Tidak #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where Metal_Purity='" & UCase(Frm95.TB10) & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If rs!ID <> Frm95.L13_Text Then
                rs.Close
                Set rs = Nothing
                
                MsgBox "Nama Purity [" & UCase(Frm95.TB10) & "] Telah Didaftarkan Sebelum Ini." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Sila Periksa Data Anda.", vbInformation, "Info"
                        
                Exit Sub
            End If
        End If
        
        rs.Close
        Set rs = Nothing
'#### Periksa Samada Nama Purity Wujud Atau Tidak #### - End

'#### Periksa Samada Kod Purity Wujud Atau Tidak #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where Kod_Metal_Purity='" & UCase(Frm95.TB11) & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If rs!ID <> Frm95.L13_Text Then
                rs.Close
                Set rs = Nothing
                
                MsgBox "Kod Purity [" & UCase(Frm95.TB11) & "] Telah Didaftarkan Sebelum Ini." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Sila Periksa Data Anda.", vbInformation, "Info"
                        
                Exit Sub
            End If
        End If
        
        rs.Close
        Set rs = Nothing
'#### Periksa Samada Kod Purity Wujud Atau Tidak #### - End

'#### Memorize Data Asal #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where ID='" & Frm95.L13_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!Metal_Purity) Then Frm95_LM_PURITY = rs!Metal_Purity 'Nama Purity
            If Not IsNull(rs!Kod_Metal_Purity) Then Frm95_LM_KOD_PURITY = rs!Kod_Metal_Purity 'Kod Purity

        End If
        
        rs.Close
        Set rs = Nothing
'#### Memorize Data Asal #### - End

'#### Simpan Data Purity #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where Metal_Purity='" & Frm95_LM_PURITY & "'", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
        
            If Frm95.TB10 <> vbNullString Then 'Nama Purity
                rs!Metal_Purity = UCase(Frm95.TB10)
            Else
                rs!Metal_Purity = Null
            End If
            If Frm95.TB11 <> vbNullString Then 'Kod Purity
                rs!Kod_Metal_Purity = UCase(Frm95.TB11)
            Else
                rs!Kod_Metal_Purity = Null
            End If
            If Frm95.TB16 <> vbNullString Then 'Kadar Tukaran Purity 999.9
                rs!kadar_tukaran = Format(Frm95.TB16, "0.00")
            Else
                rs!kadar_tukaran = "0.00"
            End If
            If Frm95.TB17 <> vbNullString Then 'Kadar pemalar harga trade in
                rs!trade_in = Frm95.TB17
            Else
                rs!trade_in = "0.00"
            End If
            If Frm95.TB18 <> vbNullString Then 'Assay
                rs!assay = Frm95.TB18
            Else
                rs!assay = "0.00"
            End If

            rs.Update
            DATA_SAVE = 1
        End If
        
        rs.Close
        Set rs = Nothing
'#### Simpan Data Supplier #### - End

        If DATA_SAVE = 1 Then
'#### Update Nama & Kod Purity #### - Start
            If Frm95_LM_PURITY <> UCase(Frm95.TB10) Or Frm95_LM_KOD_PURITY <> UCase(Frm95.TB11) Then
            
                '#### Update Maklumat Purity Dalam Table hargaemas #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE hargaemas set Purity='" & UCase(Frm95.TB11) & "'" _
                & "WHERE Purity='" & Frm95_LM_KOD_PURITY & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Purity Dalam Table hargaemas #### - End
                
                '#### Update Maklumat Purity Dalam Table Data_Database #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE Data_Database set purity='" & UCase(Frm95.TB10) & "'," _
                & "kod_Purity='" & UCase(Frm95.TB11) & "'" _
                & "WHERE purity='" & Frm95_LM_PURITY & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Purity Dalam Table Data_Database #### - End
                
                '#### Update Maklumat Purity Dalam Table 12_gold_bar_database #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE 12_gold_bar_database set purity='" & UCase(Frm95.TB10) & "'," _
                & "kod_purity='" & UCase(Frm95.TB11) & "'" _
                & "WHERE purity='" & Frm95_LM_PURITY & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Purity Dalam Table 12_gold_bar_database #### - End
                
                '#### Update Maklumat Purity Dalam Table 23_senarai_jualan #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 23_senarai_jualan set purity='" & UCase(Frm95.TB11) & "'" _
                '& "WHERE purity='" & Frm95_LM_PURITY & "'"
                
                strsql = "UPDATE 23_senarai_jualan set purity='" & UCase(Frm95.TB11) & "'" _
                & "WHERE purity='" & Frm95_LM_KOD_PURITY & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Purity Dalam Table 23_senarai_jualan #### - End
                
                '#### Update Maklumat Purity Dalam Table 27_senarai_ansuran #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE 27_senarai_ansuran set purity='" & UCase(Frm95.TB11) & "'" _
                & "WHERE purity='" & Frm95_LM_PURITY & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Purity Dalam Table 27_senarai_ansuran #### - End
                
                '#### Update Maklumat Purity Dalam Table 40_tempahan_deposit #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE 40_tempahan_deposit set purity='" & UCase(Frm95.TB11) & "'" _
                & "WHERE purity='" & Frm95_LM_PURITY & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Purity Dalam Table 40_tempahan_deposit #### - End
                
                '#### Update Maklumat Purity Dalam Table 42_tempahan_siap #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE 42_tempahan_siap set purity='" & UCase(Frm95.TB11) & "'" _
                & "WHERE purity='" & Frm95_LM_PURITY & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Purity Dalam Table 42_tempahan_siap #### - End
                
                '#### Update Maklumat Purity Dalam Table 50_belian_emas_agen #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE 50_belian_emas_agen set purity='" & UCase(Frm95.TB10) & "'," _
                & "kod_purity='" & UCase(Frm95.TB11) & "'" _
                & "WHERE purity='" & Frm95_LM_PURITY & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Purity Dalam Table 50_belian_emas_agen #### - End
                
                '#### Update Maklumat Purity Dalam Table 50_belian_emas_agen #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE 79_grn set purity='" & UCase(Frm95.TB10) & "' WHERE purity='" & Frm95_LM_PURITY & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Purity Dalam Table 50_belian_emas_agen #### - End
                
                '#### Update Maklumat Purity Dalam Table 50_belian_emas_agen #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE 85_penggunaan_ti set purity='" & UCase(Frm95.TB10) & "' WHERE purity='" & Frm95_LM_PURITY & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Purity Dalam Table 50_belian_emas_agen #### - End
                
            End If
'#### Update Nama & Kod Purity #### - End

'#### Update Log Aktiviti Sistem #### - Start
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & user & "] Edit Purity [" & UCase(Frm95.TB10) & "]."
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End
            
            Call Frm95_initial
            
            MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
        End If

    End If
End If
End Sub
Private Sub CMD6_Click()
'on error resume next
Frm95.Frame3.Visible = False
Frm95.Frame4.Visible = True
End Sub
Private Sub CMD7_Click()
'On Error Resume Next
Dim Err(3)
DATA_SAVE = 0

If Frm95.TB12 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan Nama Produk."
End If
If Frm95.TB12 <> vbNullString Then
    If InStr(1, Frm95.TB12, "*") <> 0 Or InStr(1, Frm95.TB12, "/") <> 0 Or InStr(1, Frm95.TB12, "\") <> 0 Or InStr(1, Frm95.TB12, "'") <> 0 Then
        x = x + 1
        Err(x) = "Nama produk mempunyai simbol yang tidak sah."
    End If
End If
If Frm95.TB13 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan Kod Produk."
End If
If Frm95.TB13 <> vbNullString Then
    If InStr(1, Frm95.TB13, "*") <> 0 Or InStr(1, Frm95.TB13, "/") <> 0 Or InStr(1, Frm95.TB13, "\") <> 0 Or InStr(1, Frm95.TB13, "'") <> 0 Then
        x = x + 1
        Err(x) = "Kod produk mempunyai simbol yang tidak sah."
    End If
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah Anda Ingin Simpan Data Ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then

'#### Periksa Samada Nama Produk Wujud Atau Tidak #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where Kategori_Produk='" & UCase(Frm95.TB12) & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            rs.Close
            Set rs = Nothing
            
            MsgBox "Nama Produk [" & UCase(Frm95.TB12) & "] Telah Didaftarkan Sebelum Ini." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila Periksa Data Anda.", vbInformation, "Info"
                    
            Exit Sub
        End If
        
        rs.Close
        Set rs = Nothing
'#### Periksa Samada Nama Produk Wujud Atau Tidak #### - End

'#### Periksa Samada Kod Produk Wujud Atau Tidak #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where Kod_Kategori_Produk='" & UCase(Frm95.TB13) & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            rs.Close
            Set rs = Nothing
            
            MsgBox "Kod Produk [" & UCase(Frm95.TB13) & "] Telah Didaftarkan Sebelum Ini." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila Periksa Data Anda.", vbInformation, "Info"
                    
            Exit Sub
        End If
        
        rs.Close
        Set rs = Nothing
'#### Periksa Samada Kod Produk Wujud Atau Tidak #### - End

'#### Simpan Data Produk #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database", cn, adOpenKeyset, adLockOptimistic

        rs.AddNew
        If Frm95.TB12 <> vbNullString Then 'Nama Produk
            rs!kategori_Produk = UCase(Frm95.TB12)
        Else
            rs!kategori_Produk = Null
        End If
        If Frm95.TB13 <> vbNullString Then 'Kod Produk
            rs!Kod_Kategori_Produk = UCase(Frm95.TB13)
        Else
            rs!Kod_Kategori_Produk = Null
        End If
        rs!Status = 1 '0 : Tidak aktif , 1 : Aktif
        rs.Update
        
        rs.Close
        Set rs = Nothing
'#### Simpan Data Produk #### - End

'#### Update Log Aktiviti Sistem #### - Start
        user = MDI_frm1.L3_Text
        LogAct_Memory = "[" & user & "] Daftar Produk [" & UCase(Frm95.TB12) & "]."
        LogDate_Memory = DateTime.Date & " " & DateTime.Time$
        Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End
        
        Call Frm95_initial
        
        MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
        
        Frm95.TB12.SetFocus
    End If
End If
End Sub
Private Sub CMD8_Click()
'On Error Resume Next
Dim Err(3)
DATA_SAVE = 0
Frm95_LM_PRODUK = vbNullString
Frm95_LM_KOD_PRODUK = vbNullString

If Frm95.TB12 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan Nama Produk."
End If
If Frm95.TB13 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan Kod Produk."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah Anda Ingin Ubah Data Ini ?" & vbCrLf & _
            vbNullString & vbCrLf & _
            "Sistem Akan Mengambil Sedikit Masa Untuk Update Semua Data Di Dalam Sistem." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then

'#### Periksa Samada Nama Produk Wujud Atau Tidak #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where Kategori_Produk='" & UCase(Frm95.TB12) & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If rs!ID <> Frm95.L14_Text Then
                rs.Close
                Set rs = Nothing
                
                MsgBox "Nama Produk [" & UCase(Frm95.TB12) & "] Telah Didaftarkan Sebelum Ini." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Sila Periksa Data Anda.", vbInformation, "Info"
                        
                Exit Sub
            End If
        End If
        
        rs.Close
        Set rs = Nothing
'#### Periksa Samada Nama Produk Wujud Atau Tidak #### - End

'#### Periksa Samada Kod Produk Wujud Atau Tidak #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where Kod_Kategori_Produk='" & UCase(Frm95.TB13) & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If rs!ID <> Frm95.L14_Text Then
                rs.Close
                Set rs = Nothing
                
                MsgBox "Kod Produk [" & UCase(Frm95.TB13) & "] Telah Didaftarkan Sebelum Ini." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Sila Periksa Data Anda.", vbInformation, "Info"
                        
                Exit Sub
            End If
        End If
        
        rs.Close
        Set rs = Nothing
'#### Periksa Samada Kod Produk Wujud Atau Tidak #### - End

'#### Memorize Data Asal #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where ID='" & Frm95.L14_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!kategori_Produk) Then Frm95_LM_PRODUK = rs!kategori_Produk 'Nama Produk
            If Not IsNull(rs!Kod_Kategori_Produk) Then Frm95_LM_KOD_PRODUK = rs!Kod_Kategori_Produk 'Kod Produk

        End If
        
        rs.Close
        Set rs = Nothing
'#### Memorize Data Asal #### - End

'#### Simpan Data Produk #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where ID='" & Frm95.L14_Text & "'", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            If Frm95.TB12 <> vbNullString Then 'Nama Produk
                rs!kategori_Produk = UCase(Frm95.TB12)
            Else
                rs!kategori_Produk = Null
            End If
            If Frm95.TB13 <> vbNullString Then 'Kod Produk
                rs!Kod_Kategori_Produk = UCase(Frm95.TB13)
            Else
                rs!Kod_Kategori_Produk = Null
            End If
            rs.Update
            
            DATA_SAVE = 1
        End If
        
        rs.Close
        Set rs = Nothing
'#### Simpan Data Produk #### - End
        
        If DATA_SAVE = 1 Then

'#### Update Nama & Kod Produk #### - Start
            If Frm95_LM_PRODUK <> UCase(Frm95.TB12) Or Frm95_LM_KOD_PRODUK <> UCase(Frm95.TB13) Then
                
                '#### Update Maklumat Produk Dalam Table Data_Database #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE Data_Database set kategori_Produk='" & UCase(Frm95.TB12) & "'," _
                & "kod_Kategori_Produk='" & UCase(Frm95.TB13) & "'" _
                & "WHERE kategori_Produk='" & Frm95_LM_PRODUK & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Produk Dalam Table Data_Database #### - End
                
                '#### Update Maklumat Produk Dalam Table 23_senarai_jualan #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE 23_senarai_jualan set kategori_produk='" & UCase(Frm95.TB12) & "'" _
                & "WHERE kategori_produk='" & Frm95_LM_PRODUK & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Produk Dalam Table 23_senarai_jualan #### - End
                
                '#### Update Maklumat Produk Dalam Table 27_senarai_ansuran #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE 27_senarai_ansuran set kategori_produk='" & UCase(Frm95.TB12) & "'" _
                & "WHERE kategori_produk='" & Frm95_LM_PRODUK & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Produk Dalam Table 27_senarai_ansuran #### - End
                
                '#### Update Maklumat Produk Dalam Table 40_tempahan_deposit #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE 40_tempahan_deposit set kategori_produk='" & UCase(Frm95.TB12) & "'" _
                & "WHERE kategori_produk='" & Frm95_LM_PRODUK & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Produk Dalam Table 40_tempahan_deposit #### - End
                
                '#### Update Maklumat Produk Dalam Table 42_tempahan_siap #### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "UPDATE 42_tempahan_siap set kategori_produk='" & UCase(Frm95.TB12) & "'" _
                & "WHERE kategori_produk='" & Frm95_LM_PRODUK & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
                '#### Update Maklumat Produk Dalam Table 42_tempahan_siap #### - End
                
            End If
'#### Update Nama & Kod Produk #### - End
        
        
'#### Update Log Aktiviti Sistem #### - Start
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & user & "] Edit Produk [" & UCase(Frm95.TB12) & "]."
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End
            
            Call Frm95_initial
            
            MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
        End If
    End If
End If
End Sub
Private Sub CMD9_Click()
'on error resume next
Frm95.Frame5.Visible = False
Frm95.Frame6.Visible = True
End Sub
Private Sub Form_Load()
'on error resume next
Call Frm95_on_time_reset
End Sub
Private Sub Frm95_SM_edit_Click()
'on error resume next
Frm95_LM_ID = vbNullString
DATA_FOUND = 0

If IsNumeric(Frm95.LV1.SelectedItem.Index) Then
    
    Frm95_LM_ID = Frm95.LV1.ListItems(Frm95.LV1.SelectedItem.Index)
    
    If Frm95_LM_ID <> vbNullString Then
    
        Call Frm95_initial
        
'#### Carian Data Supplier #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where ID='" & Frm95_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            
            Frm95.L12_Text = Frm95_LM_ID 'No. ID Database
            If Not IsNull(rs!jenis_supplier) Then
                If rs!jenis_supplier = "Supplier" Then
                    Frm95.CB1 = 1
                    Frm95.CB2 = 0
                ElseIf rs!jenis_supplier = "Agen/Kedai" Then
                    Frm95.CB1 = 0
                    Frm95.CB2 = 1
                End If
            End If
            If Not IsNull(rs!supplier) Then Frm95.TB1 = rs!supplier 'Nama Supplier
            If Not IsNull(rs!Kod_Supplier) Then Frm95.TB2 = rs!Kod_Supplier 'Kod Supplier
            If Not IsNull(rs!no_id_gst) Then Frm95.TB3 = rs!no_id_gst 'No. ID GST
            If Not IsNull(rs!no_pendaftaran) Then Frm95.TB4 = rs!no_pendaftaran 'No. Pendaftaran Syarikat
            If Not IsNull(rs!no_tel_off) Then Frm95.TB5 = rs!no_tel_off 'No. Telefon Office
            If Not IsNull(rs!no_tel_hp) Then Frm95.TB6 = rs!no_tel_hp 'No. Telefon HP
            If Not IsNull(rs!alamat) Then Frm95.TB7 = rs!alamat 'Alamat
            If Not IsNull(rs!nama_bank) Then Frm95.TB8 = rs!nama_bank 'Nama Bank
            If Not IsNull(rs!no_akaun) Then Frm95.TB9 = rs!no_akaun 'No. Akaun
            DATA_FOUND = 1
            
        End If
        
        rs.Close
        Set rs = Nothing
'#### Carian Data Supplier #### - End

        If DATA_FOUND = 1 Then
            Frm95.CMD1.Visible = False
            Frm95.CMD2.Visible = True
            Frm95.CMD3.Visible = True
            
            Frm95.Frame1.Visible = True
            Frm95.Frame2.Visible = False
        End If
        
    End If
End If
End Sub
Private Sub Frm95_SM_edit_dulang_Click()
'on error resume next
DATA_FOUND = 0
Frm95_LM_ID = vbNullString

If IsNumeric(Frm95.LV4.SelectedItem.Index) Then
    
    Frm95_LM_ID = Frm95.LV4.ListItems(Frm95.LV4.SelectedItem.Index)
    
    If Frm95_LM_ID <> vbNullString Then
    
        Call Frm95_initial
        
'#### Carian Data Produk #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where ID='" & Frm95_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            
            Frm95.L15_Text = Frm95_LM_ID 'No. ID Database
            If Not IsNull(rs!SenaraiDulang) Then Frm95.TB14 = rs!SenaraiDulang 'Nama Dulang
            DATA_FOUND = 1
            
        End If
        
        rs.Close
        Set rs = Nothing
'#### Carian Data Produk #### - End

        If DATA_FOUND = 1 Then
            Frm95.CMD10.Visible = False
            Frm95.CMD11.Visible = True
            Frm95.CMD12.Visible = True
            
            Frm95.Frame7.Visible = True
            Frm95.Frame8.Visible = False
        End If
        
    End If
End If
End Sub
Private Sub Frm95_SM_edit_produk_Click()
'on error resume next
DATA_FOUND = 0
Frm95_LM_ID = vbNullString

If IsNumeric(Frm95.LV3.SelectedItem.Index) Then
    
    Frm95_LM_ID = Frm95.LV3.ListItems(Frm95.LV3.SelectedItem.Index)
    
    If Frm95_LM_ID <> vbNullString Then
    
        Call Frm95_initial
        
'#### Carian Data Produk #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where ID='" & Frm95_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            
            Frm95.L14_Text = Frm95_LM_ID 'No. ID Database
            If Not IsNull(rs!kategori_Produk) Then Frm95.TB12 = rs!kategori_Produk 'Nama Produk
            If Not IsNull(rs!Kod_Kategori_Produk) Then Frm95.TB13 = rs!Kod_Kategori_Produk 'Kod Produk
            DATA_FOUND = 1
            
        End If
        
        rs.Close
        Set rs = Nothing
'#### Carian Data Produk #### - End

        If DATA_FOUND = 1 Then
            Frm95.CMD7.Visible = False
            Frm95.CMD8.Visible = True
            Frm95.CMD9.Visible = True
            
            Frm95.Frame5.Visible = True
            Frm95.Frame6.Visible = False
        End If
        
    End If
End If
End Sub
Private Sub Frm95_SM_edit_purity_Click()
'on error resume next
DATA_FOUND = 0
Frm95_LM_ID = vbNullString

If IsNumeric(Frm95.LV2.SelectedItem.Index) Then
    
    Frm95_LM_ID = Frm95.LV2.ListItems(Frm95.LV2.SelectedItem.Index)
    
    If Frm95_LM_ID <> vbNullString Then
    
        Call Frm95_initial
        
'#### Carian Data Supplier #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where ID='" & Frm95_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            
            Frm95.L13_Text = Frm95_LM_ID 'No. ID Database
            If Not IsNull(rs!Metal_Purity) Then Frm95.TB10 = rs!Metal_Purity 'Nama Purity
            If Not IsNull(rs!Kod_Metal_Purity) Then Frm95.TB11 = rs!Kod_Metal_Purity 'Kod Purity
            If Not IsNull(rs!assay) Then Frm95.TB18 = rs!assay 'Assay
            If Not IsNull(rs!trade_in) Then Frm95.TB17 = rs!trade_in 'Kadar pemalar harga trade in
            DATA_FOUND = 1
            
        End If
        
        rs.Close
        Set rs = Nothing
'#### Carian Data Supplier #### - End

        If DATA_FOUND = 1 Then
            Frm95.CMD4.Visible = False
            Frm95.CMD5.Visible = True
            Frm95.CMD6.Visible = True
            
            Frm95.Frame3.Visible = True
            Frm95.Frame4.Visible = False
        End If
        
    End If
End If
End Sub
Private Sub Frm95_SM_edit_tukang_Click()
'on error resume next
DATA_FOUND = 0
Frm95_LM_ID = vbNullString

If Frm95.MSFlexGrid5 <> vbNullString Then
    Frm95_LM_ID = Frm95.MSFlexGrid5.TextMatrix(Frm95.MSFlexGrid5, 2) 'No. ID

    If Frm95_LM_ID <> vbNullString Then
        Call Frm95_initial
        
'#### Carian Data Produk #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where ID='" & Frm95_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            
            Frm95.L18_Text = Frm95_LM_ID 'No. ID Database
            If Not IsNull(rs!tukang_emas) Then Frm95.TB15 = rs!tukang_emas 'Nama Dulang
            DATA_FOUND = 1
            
        End If
        
        rs.Close
        Set rs = Nothing
'#### Carian Data Produk #### - End

        If DATA_FOUND = 1 Then
            Frm95.CMD13.Visible = False
            Frm95.CMD14.Visible = True
            Frm95.CMD15.Visible = True
            
            Frm95.Pic9.Visible = True
            Frm95.Pic10.Visible = False
        End If
        
    End If
End If
End Sub
Private Sub Frm95_SSM_aktif_Click()
'On Error Resume Next
Frm95_LM_ID = vbNullString
DATA_SAVE = 0
Frm95_LM_KOD_SUPPLIER = vbNullString

DATA_FOUND = 0

frm95_LM_No_ID = vbNullString

If IsNumeric(Frm95.LV1.SelectedItem.Index) Then
    
    Frm95_LM_ID = Frm95.LV1.ListItems(Frm95.LV1.SelectedItem.Index)
    
    If Frm95_LM_ID <> vbNullString Then
    
        
        Note = "Adakah anda ingin ubah status supplier ini kepada AKTIF ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then

            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from setting_database where ID='" & Frm95_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!Status) Then
                    
                    If rs!Status = 0 Then
                        
                        If Not IsNull(rs!Kod_Supplier) Then Frm95_LM_KOD_SUPPLIER = rs!Kod_Supplier
                        rs!Status = 1
                        rs.Update
                        DATA_SAVE = 1
                        
                    ElseIf rs!Status = 1 Then
                        
                        DATA_SAVE = 2
                        
                    End If
                    
                End If
                
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_SAVE = 0 Then
            
                MsgBox "Tiada data yang diupdate." & vbCrLf & _
                        "Sila keluar dari menu ini dan cuba sekali lagi.", vbExclamation, "Info"
                        
            ElseIf DATA_SAVE = 1 Then
                
                '#### Update Log Aktiviti Sistem #### - Start
                user = MDI_frm1.L3_Text
                LogAct_Memory = "[" & user & "] Tukar status supplier [" & Frm95_LM_KOD_SUPPLIER & "] kepada AKTIF."
                LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                Call UpdateLog_Database
                '#### Update Log Aktiviti Sistem #### - End
                
                Call Frm95_senarai_supplier_header
                Call Frm95_senarai_supplier
                                
                MsgBox "Status bagi supplier ini telah berjaya ditukar kepada AKTIF.", vbInformation, "Info"
            
            ElseIf DATA_SAVE = 2 Then
            
                MsgBox "Tiada data yang diupdate kerana status asal supplier ini telah AKTIF.", vbInformation, "Info"
            
            End If

            
        End If
        
    End If

End If
End Sub
Private Sub Frm95_SSM_aktif2_Click()
'On Error Resume Next
Frm95_LM_ID = vbNullString
DATA_SAVE = 0
Frm95_LM_KOD_PURITY = vbNullString

If IsNumeric(Frm95.LV2.SelectedItem.Index) Then
    
    Frm95_LM_ID = Frm95.LV2.ListItems(Frm95.LV2.SelectedItem.Index)
    
    If Frm95_LM_ID <> vbNullString Then
        
        Note = "Adakah anda ingin ubah status purity ini kepada AKTIF ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then

            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from setting_database where ID='" & Frm95_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!Status) Then
                    
                    If rs!Status = 0 Then
                        
                        If Not IsNull(rs!Kod_Metal_Purity) Then Frm95_LM_KOD_PURITY = rs!Kod_Metal_Purity
                        rs!Status = 1
                        rs.Update
                        DATA_SAVE = 1
                        
                    ElseIf rs!Status = 1 Then
                        
                        DATA_SAVE = 2
                        
                    End If
                    
                End If
                
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_SAVE = 0 Then
            
                MsgBox "Tiada data yang diupdate." & vbCrLf & _
                        "Sila keluar dari menu ini dan cuba sekali lagi.", vbExclamation, "Info"
                        
            ElseIf DATA_SAVE = 1 Then
                
                '#### Update Log Aktiviti Sistem #### - Start
                user = MDI_frm1.L3_Text
                LogAct_Memory = "[" & user & "] Tukar status purity [" & Frm95_LM_KOD_PURITY & "] kepada AKTIF."
                LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                Call UpdateLog_Database
                '#### Update Log Aktiviti Sistem #### - End
                
                Call Frm95_senarai_purity_header
                Call Frm95_senarai_purity
                                
                MsgBox "Status bagi purity ini telah berjaya ditukar kepada AKTIF.", vbInformation, "Info"
            
            ElseIf DATA_SAVE = 2 Then
            
                MsgBox "Tiada data yang diupdate kerana status asal purity ini telah AKTIF.", vbInformation, "Info"
            
            End If
            
        End If
            
    End If

End If
End Sub
Private Sub Frm95_SSM_aktif3_Click()
'On Error Resume Next
Frm95_LM_ID = vbNullString
DATA_SAVE = 0
Frm95_LM_KOD_PRODUK = vbNullString

If IsNumeric(Frm95.LV3.SelectedItem.Index) Then
    
    Frm95_LM_ID = Frm95.LV3.ListItems(Frm95.LV3.SelectedItem.Index)
    
    If Frm95_LM_ID <> vbNullString Then
        
        Note = "Adakah anda ingin ubah status produk ini kepada AKTIF ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then

            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from setting_database where ID='" & Frm95_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!Status) Then
                    
                    If rs!Status = 0 Then
                        
                        If Not IsNull(rs!Kod_Kategori_Produk) Then Frm95_LM_KOD_PRODUK = rs!Kod_Kategori_Produk
                        rs!Status = 1
                        rs.Update
                        DATA_SAVE = 1
                        
                    ElseIf rs!Status = 1 Then
                        
                        DATA_SAVE = 2
                        
                    End If
                    
                End If
                
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_SAVE = 0 Then
            
                MsgBox "Tiada data yang diupdate." & vbCrLf & _
                        "Sila keluar dari menu ini dan cuba sekali lagi.", vbExclamation, "Info"
                        
            ElseIf DATA_SAVE = 1 Then
                
                '#### Update Log Aktiviti Sistem #### - Start
                user = MDI_frm1.L3_Text
                LogAct_Memory = "[" & user & "] Tukar status produk [" & Frm95_LM_KOD_PRODUK & "] kepada AKTIF."
                LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                Call UpdateLog_Database
                '#### Update Log Aktiviti Sistem #### - End
                
                Call Frm95_senarai_produk_header
                Call Frm95_senarai_produk
                                
                MsgBox "Status bagi produk ini telah berjaya ditukar kepada AKTIF.", vbInformation, "Info"
            
            ElseIf DATA_SAVE = 2 Then
            
                MsgBox "Tiada data yang diupdate kerana status asal produk ini telah AKTIF.", vbInformation, "Info"
            
            End If
            
        End If
        
    End If
    
End If
End Sub
Private Sub Frm95_SSM_aktif4_Click()
'On Error Resume Next
Frm95_LM_ID = vbNullString
DATA_SAVE = 0
Frm95_LM_KOD_DULANG = vbNullString

If IsNumeric(Frm95.LV4.SelectedItem.Index) Then
    
    Frm95_LM_ID = Frm95.LV4.ListItems(Frm95.LV4.SelectedItem.Index)
    
    If Frm95_LM_ID <> vbNullString Then
        
        Note = "Adakah anda ingin ubah status dulang ini kepada AKTIF ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then

            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from setting_database where ID='" & Frm95_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!Status) Then
                    
                    If rs!Status = 0 Then
                        
                        If Not IsNull(rs!SenaraiDulang) Then Frm95_LM_KOD_DULANG = rs!SenaraiDulang
                        rs!Status = 1
                        rs.Update
                        DATA_SAVE = 1
                        
                    ElseIf rs!Status = 1 Then
                        
                        DATA_SAVE = 2
                        
                    End If
                    
                End If
                
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_SAVE = 0 Then
            
                MsgBox "Tiada data yang diupdate." & vbCrLf & _
                        "Sila keluar dari menu ini dan cuba sekali lagi.", vbExclamation, "Info"
                        
            ElseIf DATA_SAVE = 1 Then
                
                '#### Update Log Aktiviti Sistem #### - Start
                user = MDI_frm1.L3_Text
                LogAct_Memory = "[" & user & "] Tukar status dulang [" & Frm95_LM_KOD_DULANG & "] kepada AKTIF."
                LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                Call UpdateLog_Database
                '#### Update Log Aktiviti Sistem #### - End
                
                Call Frm95_senarai_dulang_header
                Call Frm95_senarai_dulang
                                
                MsgBox "Status bagi dulang ini telah berjaya ditukar kepada AKTIF.", vbInformation, "Info"
            
            ElseIf DATA_SAVE = 2 Then
            
                MsgBox "Tiada data yang diupdate kerana status asal dulang ini telah AKTIF.", vbInformation, "Info"
            
            End If
            
        End If
        
    End If

End If
End Sub
Private Sub Frm95_SSM_tidak_aktif_Click()
'On Error Resume Next
Frm95_LM_ID = vbNullString
DATA_SAVE = 0
Frm95_LM_KOD_SUPPLIER = vbNullString

DATA_FOUND = 0

frm95_LM_No_ID = vbNullString

If IsNumeric(Frm95.LV1.SelectedItem.Index) Then
    
    Frm95_LM_ID = Frm95.LV1.ListItems(Frm95.LV1.SelectedItem.Index)
    
    If Frm95_LM_ID <> vbNullString Then
        
        Note = "Adakah anda ingin ubah status supplier ini kepada TIDAK AKTIF ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then

            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from setting_database where ID='" & Frm95_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!Status) Then
                    
                    If rs!Status = 1 Then
                        
                        If Not IsNull(rs!Kod_Supplier) Then Frm95_LM_KOD_SUPPLIER = rs!Kod_Supplier
                        rs!Status = 0
                        rs.Update
                        DATA_SAVE = 1
                        
                    ElseIf rs!Status = 0 Then
                        
                        DATA_SAVE = 2
                        
                    End If
                    
                End If
                
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_SAVE = 0 Then
            
                MsgBox "Tiada data yang diupdate." & vbCrLf & _
                        "Sila keluar dari menu ini dan cuba sekali lagi.", vbExclamation, "Info"
                        
            ElseIf DATA_SAVE = 1 Then
                
                '#### Update Log Aktiviti Sistem #### - Start
                user = MDI_frm1.L3_Text
                LogAct_Memory = "[" & user & "] Tukar status supplier [" & Frm95_LM_KOD_SUPPLIER & "] kepada TIDAK AKTIF."
                LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                Call UpdateLog_Database
                '#### Update Log Aktiviti Sistem #### - End
                
                Call Frm95_senarai_supplier_header
                Call Frm95_senarai_supplier
                                
                MsgBox "Status bagi supplier ini telah berjaya ditukar kepada TIDAK AKTIF.", vbInformation, "Info"
            
            ElseIf DATA_SAVE = 2 Then
            
                MsgBox "Tiada data yang diupdate kerana status asal supplier ini telah TIDAK AKTIF.", vbInformation, "Info"
            
            End If
            
        End If
        
    End If

End If
End Sub
Private Sub Frm95_SSM_tidak_aktif2_Click()
'On Error Resume Next
Frm95_LM_ID = vbNullString
DATA_SAVE = 0
Frm95_LM_KOD_PURITY = vbNullString

If IsNumeric(Frm95.LV2.SelectedItem.Index) Then
    
    Frm95_LM_ID = Frm95.LV2.ListItems(Frm95.LV2.SelectedItem.Index)
    
    If Frm95_LM_ID <> vbNullString Then
        
        Note = "Adakah anda ingin ubah status purity ini kepada TIDAK AKTIF ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then

            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from setting_database where ID='" & Frm95_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!Status) Then
                    
                    If rs!Status = 1 Then
                        
                        If Not IsNull(rs!Kod_Metal_Purity) Then Frm95_LM_KOD_PURITY = rs!Kod_Metal_Purity
                        rs!Status = 0
                        rs.Update
                        DATA_SAVE = 1
                        
                    ElseIf rs!Status = 0 Then
                        
                        DATA_SAVE = 2
                        
                    End If
                    
                End If
                
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_SAVE = 0 Then
            
                MsgBox "Tiada data yang diupdate." & vbCrLf & _
                        "Sila keluar dari menu ini dan cuba sekali lagi.", vbExclamation, "Info"
                        
            ElseIf DATA_SAVE = 1 Then
                
                '#### Update Log Aktiviti Sistem #### - Start
                user = MDI_frm1.L3_Text
                LogAct_Memory = "[" & user & "] Tukar status purity [" & Frm95_LM_KOD_PURITY & "] kepada TIDAK AKTIF."
                LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                Call UpdateLog_Database
                '#### Update Log Aktiviti Sistem #### - End
                
                Call Frm95_senarai_purity_header
                Call Frm95_senarai_purity
                                
                MsgBox "Status bagi purity ini telah berjaya ditukar kepada TIDAK AKTIF.", vbInformation, "Info"
            
            ElseIf DATA_SAVE = 2 Then
            
                MsgBox "Tiada data yang diupdate kerana status asal purity ini telah TIDAK AKTIF.", vbInformation, "Info"
            
            End If
            
        End If
        
    End If

End If
End Sub
Private Sub Frm95_SSM_tidak_aktif3_Click()
'On Error Resume Next
Frm95_LM_ID = vbNullString
DATA_SAVE = 0
Frm95_LM_KOD_PRODUK = vbNullString

If IsNumeric(Frm95.LV3.SelectedItem.Index) Then
    
    Frm95_LM_ID = Frm95.LV3.ListItems(Frm95.LV3.SelectedItem.Index)
    
    If Frm95_LM_ID <> vbNullString Then
        
        Note = "Adakah anda ingin ubah status purity ini kepada TIDAK AKTIF ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then

            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from setting_database where ID='" & Frm95_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!Status) Then
                    
                    If rs!Status = 1 Then
                        
                        If Not IsNull(rs!Kod_Kategori_Produk) Then Frm95_LM_KOD_PRODUK = rs!Kod_Kategori_Produk
                        rs!Status = 0
                        rs.Update
                        DATA_SAVE = 1
                        
                    ElseIf rs!Status = 0 Then
                        
                        DATA_SAVE = 2
                        
                    End If
                    
                End If
                
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_SAVE = 0 Then
            
                MsgBox "Tiada data yang diupdate." & vbCrLf & _
                        "Sila keluar dari menu ini dan cuba sekali lagi.", vbExclamation, "Info"
                        
            ElseIf DATA_SAVE = 1 Then
                
                '#### Update Log Aktiviti Sistem #### - Start
                user = MDI_frm1.L3_Text
                LogAct_Memory = "[" & user & "] Tukar status produk [" & Frm95_LM_KOD_PRODUK & "] kepada TIDAK AKTIF."
                LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                Call UpdateLog_Database
                '#### Update Log Aktiviti Sistem #### - End
                
                Call Frm95_senarai_produk_header
                Call Frm95_senarai_produk
                                
                MsgBox "Status bagi produk ini telah berjaya ditukar kepada TIDAK AKTIF.", vbInformation, "Info"
            
            ElseIf DATA_SAVE = 2 Then
            
                MsgBox "Tiada data yang diupdate kerana status asal produk ini telah TIDAK AKTIF.", vbInformation, "Info"
            
            End If
            
        End If
        
    End If
    
End If
End Sub
Private Sub Frm95_SSM_tidak_aktif4_Click()
'On Error Resume Next
Frm95_LM_ID = vbNullString
DATA_SAVE = 0
Frm95_LM_KOD_DULANG = vbNullString

If IsNumeric(Frm95.LV4.SelectedItem.Index) Then
    
    Frm95_LM_ID = Frm95.LV4.ListItems(Frm95.LV4.SelectedItem.Index)
    
    If Frm95_LM_ID <> vbNullString Then
    
        Note = "Adakah anda ingin ubah status dulang ini kepada TIDAK AKTIF ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then

            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from setting_database where ID='" & Frm95_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!Status) Then
                    
                    If rs!Status = 1 Then
                        
                        If Not IsNull(rs!SenaraiDulang) Then Frm95_LM_KOD_DULANG = rs!SenaraiDulang
                        rs!Status = 0
                        rs.Update
                        DATA_SAVE = 1
                        
                    ElseIf rs!Status = 0 Then
                        
                        DATA_SAVE = 2
                        
                    End If
                    
                End If
                
            End If
            
            rs.Close
            Set rs = Nothing
        
            If DATA_SAVE = 0 Then
            
                MsgBox "Tiada data yang diupdate." & vbCrLf & _
                        "Sila keluar dari menu ini dan cuba sekali lagi.", vbExclamation, "Info"
                        
            ElseIf DATA_SAVE = 1 Then
                
                '#### Update Log Aktiviti Sistem #### - Start
                user = MDI_frm1.L3_Text
                LogAct_Memory = "[" & user & "] Tukar status dulang [" & Frm95_LM_KOD_DULANG & "] kepada TIDAK AKTIF."
                LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                Call UpdateLog_Database
                '#### Update Log Aktiviti Sistem #### - End
                
                Call Frm95_senarai_dulang_header
                Call Frm95_senarai_dulang
                                
                MsgBox "Status bagi dulang ini telah berjaya ditukar kepada TIDAK AKTIF.", vbInformation, "Info"
            
            ElseIf DATA_SAVE = 2 Then
            
                MsgBox "Tiada data yang diupdate kerana status asal dulang ini telah TIDAK AKTIF.", vbInformation, "Info"
            
            End If
            
        End If
        
    End If

End If
End Sub
Private Sub L10_Text_Click()
'on error resume next
If Frm95.Frame8.Visible = False Then

    Call Frm95_initial
    Call Frm95_invisible
    Call Frm95_senarai_dulang_header
    Call Frm95_senarai_dulang
    
    Frm95.Frame8.Visible = True
    
Else

    Frm95.Frame8.Visible = False
    
End If
End Sub
Private Sub L16_Text_Click()
'on error resume next
If Frm95.Pic9.Visible = False Then
    Call Frm95_initial
    
    Frm95.Pic9.Visible = True
    Frm95.TB15.SetFocus
Else
    Frm95.Pic9.Visible = False
End If
End Sub
Private Sub L17_Text_Click()
'on error resume next
If Frm95.Pic10.Visible = False Then
    Call Frm95_initial
    Call Frm95_senarai_tukang_header
    Call Frm95_senarai_tukang
    
    Frm95.Pic10.Visible = True
Else
    Frm95.Pic10.Visible = False
End If
End Sub

Private Sub L3_Text_Click()
'on error resume next
If Frm95.Frame1.Visible = False Then

    Call Frm95_initial
    Call Frm95_invisible
    
    Frm95.Frame1.Visible = True
    
Else

    Frm95.Frame1.Visible = False
    
End If
End Sub
Private Sub L4_Text_Click()
'on error resume next
If Frm95.Frame2.Visible = False Then

    Call Frm95_initial
    Call Frm95_invisible
    
    Call Frm95_senarai_supplier_header
    Call Frm95_senarai_supplier
    
    Frm95.Frame2.Visible = True
    
Else

    Frm95.Frame2.Visible = False
    
End If
End Sub
Private Sub L5_Text_Click()
'on error resume next
If Frm95.Frame3.Visible = False Then

    Call Frm95_initial
    Call Frm95_invisible
    
    Frm95.Frame3.Visible = True
    Frm95.TB10.SetFocus
    
Else

    Frm95.Frame3.Visible = False
    
End If
End Sub
Private Sub L6_Text_Click()
'on error resume next
If Frm95.Frame4.Visible = False Then

    Call Frm95_initial
    Call Frm95_invisible
    Call Frm95_senarai_purity_header
    Call Frm95_senarai_purity
    
    Frm95.Frame4.Visible = True
    
Else

    Frm95.Frame4.Visible = False
    
End If
End Sub
Private Sub L7_Text_Click()
'on error resume next
If Frm95.Frame5.Visible = False Then

    Call Frm95_initial
    Call Frm95_invisible
    
    Frm95.Frame5.Visible = True
    Frm95.TB12.SetFocus
    
Else

    Frm95.Frame5.Visible = False
    
End If
End Sub
Private Sub L8_Text_Click()
'on error resume next
If Frm95.Frame6.Visible = False Then

    Call Frm95_initial
    Call Frm95_invisible
    Call Frm95_senarai_produk_header
    Call Frm95_senarai_produk
    
    Frm95.Frame6.Visible = True
    
Else

    Frm95.Frame6.Visible = False
    
End If
End Sub
Private Sub L9_Text_Click()
'on error resume next
If Frm95.Frame7.Visible = False Then

    Call Frm95_initial
    Call Frm95_invisible
    
    Frm95.Frame7.Visible = True
    Frm95.TB14.SetFocus
    
Else

    Frm95.Frame7.Visible = False
    
End If
End Sub
Private Sub LV1_DblClick()
'On Error Resume Next
frm95_LM_No_ID = vbNullString

If IsNumeric(Frm95.LV1.SelectedItem.Index) Then
    
    frm95_LM_No_ID = Frm95.LV1.ListItems(Frm95.LV1.SelectedItem.Index)
    
    If frm95_LM_No_ID <> vbNullString Then

        PopupMenu Frm95_PM_Menu
        
    End If
    
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub LV2_DblClick()
'On Error Resume Next
frm95_LM_No_ID = vbNullString

If IsNumeric(Frm95.LV2.SelectedItem.Index) Then
    
    frm95_LM_No_ID = Frm95.LV2.ListItems(Frm95.LV2.SelectedItem.Index)
    
    If frm95_LM_No_ID <> vbNullString Then

        PopupMenu Frm95_PM_Menu2
        
    End If
    
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub



Private Sub LV3_DblClick()
'On Error Resume Next
frm95_LM_No_ID = vbNullString

If IsNumeric(Frm95.LV3.SelectedItem.Index) Then
    
    frm95_LM_No_ID = Frm95.LV3.ListItems(Frm95.LV3.SelectedItem.Index)
    
    If frm95_LM_No_ID <> vbNullString Then

        PopupMenu Frm95_PM_Menu3
        
    End If
    
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub LV4_DblClick()
'On Error Resume Next
frm95_LM_No_ID = vbNullString

If IsNumeric(Frm95.LV4.SelectedItem.Index) Then
    
    frm95_LM_No_ID = Frm95.LV4.ListItems(Frm95.LV4.SelectedItem.Index)
    
    If frm95_LM_No_ID <> vbNullString Then

        PopupMenu Frm95_PM_Menu4
        
    End If
    
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub MSFlexGrid5_DblClick()
'on error resume next
If Frm95.MSFlexGrid5 <> vbNullString Then
    PopupMenu Frm95_PM_Menu5
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
