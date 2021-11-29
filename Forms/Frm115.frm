VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm115 
   Caption         =   "Goods Despatch Note (GDN) - Per Item"
   ClientHeight    =   12735
   ClientLeft      =   120
   ClientTop       =   -36150
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
   Icon            =   "Frm115.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12735
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Senarai Stok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11415
      Left            =   1440
      TabIndex        =   65
      Top             =   1080
      Visible         =   0   'False
      Width           =   12135
      Begin VB.CommandButton CMD5 
         Caption         =   "Tutup Paparan Ini"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   240
         MouseIcon       =   "Frm115.frx":0ECA
         MousePointer    =   99  'Custom
         Picture         =   "Frm115.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   10440
         Width           =   2175
      End
      Begin VB.CommandButton CMD14 
         Caption         =   "Next"
         Height          =   810
         Left            =   10800
         MouseIcon       =   "Frm115.frx":165E
         MousePointer    =   99  'Custom
         Picture         =   "Frm115.frx":1968
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Paparan Seterusnya"
         Top             =   10200
         Width           =   1095
      End
      Begin VB.CommandButton CMD6 
         Caption         =   "Back"
         Height          =   810
         Left            =   9600
         MouseIcon       =   "Frm115.frx":2A32
         MousePointer    =   99  'Custom
         Picture         =   "Frm115.frx":2D3C
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Paparan Sebelum"
         Top             =   10200
         Width           =   1095
      End
      Begin VB.CommandButton CMD12 
         Caption         =   "Paparan Stok"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   6600
         MouseIcon       =   "Frm115.frx":3E06
         MousePointer    =   99  'Custom
         Picture         =   "Frm115.frx":4110
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox CBB6 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   960
         Width           =   4605
      End
      Begin VB.ComboBox CBB5 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   600
         Width           =   4605
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   8490
         Left            =   120
         TabIndex        =   78
         Top             =   1680
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   14975
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
      Begin VB.Label L60_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L60_Text"
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
         TabIndex        =   87
         Top             =   10215
         Width           =   1335
      End
      Begin VB.Label Label55 
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
         TabIndex        =   86
         Top             =   10200
         Width           =   975
      End
      Begin VB.Label L64_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L64_Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   6720
         TabIndex        =   84
         Top             =   10560
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L63_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L63_Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7800
         TabIndex        =   83
         Top             =   10560
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L61_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L61_Text"
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
         Left            =   8640
         TabIndex        =   82
         Top             =   10200
         Width           =   375
      End
      Begin VB.Label L62_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L62_Text"
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
         Left            =   9240
         TabIndex        =   81
         Top             =   10200
         Width           =   615
      End
      Begin VB.Label L59_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L59_Text"
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
         Left            =   9600
         TabIndex        =   77
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L28_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai stok"
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
         TabIndex        =   76
         Top             =   1440
         Width           =   10215
      End
      Begin VB.Label L58_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L58_Text"
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
         Left            =   9600
         TabIndex        =   75
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L57_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L57_Text"
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
         Left            =   8520
         TabIndex        =   74
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L56_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L56_Text"
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
         Left            =   8520
         TabIndex        =   73
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L55_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L55_Text"
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
         Left            =   8520
         TabIndex        =   72
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label51 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila buat pilihan di bawah untuk paparan data stok."
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
         TabIndex        =   70
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Purity * :"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   150
         TabIndex        =   69
         Top             =   1000
         Width           =   1695
      End
      Begin VB.Label Label53 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kategori Produk * :"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   150
         TabIndex        =   68
         Top             =   615
         Width           =   1695
      End
      Begin VB.Label Label67 
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
         Left            =   7320
         TabIndex        =   85
         Top             =   10200
         Width           =   2295
      End
   End
   Begin VB.CommandButton CMD8 
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
      Left            =   15600
      MouseIcon       =   "Frm115.frx":51DA
      MousePointer    =   99  'Custom
      Picture         =   "Frm115.frx":54E4
      Style           =   1  'Graphical
      TabIndex        =   155
      Top             =   10400
      Width           =   2775
   End
   Begin VB.CommandButton CMD9 
      Caption         =   "Keluar"
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
      Left            =   18480
      MouseIcon       =   "Frm115.frx":7AAE
      MousePointer    =   99  'Custom
      Picture         =   "Frm115.frx":7DB8
      Style           =   1  'Graphical
      TabIndex        =   154
      Top             =   10400
      Width           =   2775
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Maklumat Cukai GST"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   15360
      TabIndex        =   138
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton CMD13 
         Caption         =   "Tutup Paparan Ini"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2040
         MouseIcon       =   "Frm115.frx":A382
         MousePointer    =   99  'Custom
         Picture         =   "Frm115.frx":A68C
         Style           =   1  'Graphical
         TabIndex        =   139
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Line Line1 
         X1              =   2400
         X2              =   5680
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label L20_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         TabIndex        =   153
         Top             =   1680
         Width           =   1785
      End
      Begin VB.Label L19_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         TabIndex        =   152
         Top             =   1440
         Width           =   1785
      End
      Begin VB.Label L18_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   151
         Top             =   1680
         Width           =   1785
      End
      Begin VB.Label Label123 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah GST (RM)"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         TabIndex        =   150
         Top             =   1200
         Width           =   1800
      End
      Begin VB.Label Label122 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Harga (RM)"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   149
         Top             =   1200
         Width           =   1800
      End
      Begin VB.Label Label118 
         BackStyle       =   0  'Transparent
         Caption         =   "Standard Rated (SR)"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   148
         Top             =   1680
         Width           =   2145
      End
      Begin VB.Label Label121 
         BackStyle       =   0  'Transparent
         Caption         =   "Zero Rated (ZR)"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   147
         Top             =   1440
         Width           =   2145
      End
      Begin VB.Label L17_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   146
         Top             =   1440
         Width           =   1785
      End
      Begin VB.Label L16_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3075
         TabIndex        =   145
         Top             =   720
         Width           =   1545
      End
      Begin VB.Label Label115 
         BackStyle       =   0  'Transparent
         Caption         =   ": RM"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2640
         TabIndex        =   144
         Top             =   720
         Width           =   600
      End
      Begin VB.Label L15_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3075
         TabIndex        =   143
         Top             =   480
         Width           =   1545
      End
      Begin VB.Label Label111 
         BackStyle       =   0  'Transparent
         Caption         =   ": RM"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2640
         TabIndex        =   142
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label117 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Harga Dengan GST"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   141
         Top             =   720
         Width           =   2505
      End
      Begin VB.Label Label114 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Harga Tanpa GST"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   140
         Top             =   480
         Width           =   2505
      End
   End
   Begin VB.CommandButton CMD21 
      Caption         =   "Back"
      Height          =   810
      Left            =   12960
      MouseIcon       =   "Frm115.frx":AB16
      MousePointer    =   99  'Custom
      Picture         =   "Frm115.frx":AE20
      Style           =   1  'Graphical
      TabIndex        =   135
      ToolTipText     =   "Paparan Sebelum"
      Top             =   10680
      Width           =   1095
   End
   Begin VB.CommandButton CMD22 
      Caption         =   "Next"
      Height          =   810
      Left            =   14160
      MouseIcon       =   "Frm115.frx":BEEA
      MousePointer    =   99  'Custom
      Picture         =   "Frm115.frx":C1F4
      Style           =   1  'Graphical
      TabIndex        =   134
      ToolTipText     =   "Paparan Seterusnya"
      Top             =   10680
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Maklumat barang yang telah di scan."
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
      Left            =   5640
      TabIndex        =   99
      Top             =   0
      Width           =   8295
      Begin VB.CheckBox CB5 
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
         Left            =   4080
         TabIndex        =   137
         Top             =   450
         Width           =   200
      End
      Begin VB.CommandButton CMD2 
         BackColor       =   &H8000000C&
         Caption         =   "Masukkan Dalam Senarai Jualan"
         Height          =   360
         Left            =   720
         MouseIcon       =   "Frm115.frx":D2BE
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   132
         Top             =   3000
         Width           =   3135
      End
      Begin VB.CommandButton CMD1 
         Caption         =   "Masukkan Dalam Senarai Jualan"
         Height          =   360
         Left            =   2160
         MouseIcon       =   "Frm115.frx":D5C8
         MousePointer    =   99  'Custom
         TabIndex        =   131
         Top             =   3000
         Width           =   3135
      End
      Begin VB.CommandButton CMD3 
         BackColor       =   &H8000000C&
         Caption         =   "Batal Edit Data"
         Height          =   360
         Left            =   3960
         MouseIcon       =   "Frm115.frx":D8D2
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   130
         Top             =   3000
         Width           =   3135
      End
      Begin VB.TextBox TB4 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6480
         TabIndex        =   120
         Text            =   "0.00"
         Top             =   720
         Width           =   1260
      End
      Begin VB.CheckBox CB4 
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
         Left            =   4080
         TabIndex        =   119
         Top             =   2000
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
         Left            =   4080
         TabIndex        =   118
         Top             =   1535
         Width           =   200
      End
      Begin VB.CheckBox CB3 
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
         Left            =   4080
         TabIndex        =   117
         Top             =   1775
         Width           =   200
      End
      Begin VB.TextBox TB5 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   116
         Text            =   "0.00"
         Top             =   2325
         Width           =   1260
      End
      Begin VB.TextBox TB6 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   115
         Text            =   "0.00"
         Top             =   2625
         Width           =   1260
      End
      Begin VB.TextBox TB3 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   101
         Text            =   "0.00"
         Top             =   1800
         Width           =   1260
      End
      Begin VB.TextBox TB7 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2640
         MaxLength       =   5
         TabIndex        =   100
         Text            =   "0.00"
         Top             =   2130
         Width           =   1260
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanda di sini jika ada upah dikenakan."
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4320
         TabIndex        =   136
         Top             =   400
         Width           =   3405
      End
      Begin VB.Shape Shape2 
         Height          =   2655
         Left            =   3960
         Top             =   300
         Width           =   4250
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Upah Dengan GST"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4080
         TabIndex        =   128
         Top             =   2640
         Width           =   2265
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Upah (RM)"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4080
         TabIndex        =   127
         Top             =   720
         Width           =   1665
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6360
         TabIndex        =   126
         Top             =   720
         Width           =   150
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "Maklumat GST"
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
         Height          =   285
         Left            =   4125
         TabIndex        =   125
         Top             =   1005
         Width           =   1575
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "Zero Rated ZR(L)           Standard Rated SR      Standard Rated SR (Inclusive)"
         ForeColor       =   &H00000000&
         Height          =   765
         Left            =   4320
         TabIndex        =   124
         Top             =   1485
         Width           =   2730
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "** GST hanya dikenakan kepada upah SAHAJA."
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4080
         TabIndex        =   123
         Top             =   1245
         Width           =   4185
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "RM :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5850
         TabIndex        =   122
         Top             =   2325
         Width           =   600
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "RM :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5850
         TabIndex        =   121
         Top             =   2640
         Width           =   600
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Kadar Tukaran Purity 999.9"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   109
         Top             =   2145
         Width           =   2505
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Berat 999.9 (g) :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   108
         Top             =   2400
         Width           =   1515
      End
      Begin VB.Label L7_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1680
         TabIndex        =   107
         Top             =   2400
         Width           =   1305
      End
      Begin VB.Label L3_Text 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1680
         TabIndex        =   106
         Top             =   600
         Width           =   2745
      End
      Begin VB.Label L4_Text 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1680
         TabIndex        =   105
         Top             =   900
         Width           =   2745
      End
      Begin VB.Label L5_Text 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1680
         TabIndex        =   104
         Top             =   1200
         Width           =   3945
      End
      Begin VB.Label L6_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1680
         TabIndex        =   103
         Top             =   1515
         Width           =   2745
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2520
         TabIndex        =   102
         Top             =   2145
         Width           =   150
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Siri Produk :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   114
         Top             =   585
         Width           =   1515
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Purity :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   113
         Top             =   900
         Width           =   1515
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kategori Produk :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   112
         Top             =   1200
         Width           =   1515
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Berat Asal (g) :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   111
         Top             =   1515
         Width           =   1515
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Berat Jualan (g) :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   110
         Top             =   1830
         Width           =   1515
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah GST"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4080
         TabIndex        =   129
         Top             =   2355
         Width           =   2505
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
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
      Left            =   120
      TabIndex        =   89
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton CMD7 
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
         Height          =   800
         Left            =   3840
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm115.frx":DBDC
         MousePointer    =   99  'Custom
         Picture         =   "Frm115.frx":DEE6
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   880
         Width           =   1305
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
         Left            =   165
         TabIndex        =   92
         Top             =   270
         Width           =   200
      End
      Begin VB.TextBox TB1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1605
         TabIndex        =   91
         Top             =   1215
         Width           =   2100
      End
      Begin VB.TextBox TB2 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   90
         Text            =   "0.00"
         Top             =   2555
         Width           =   1380
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Harga semasa (RM/g) :"
         Height          =   285
         Left            =   165
         TabIndex        =   97
         Top             =   2565
         Width           =   2130
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Scanner Mode"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   450
         TabIndex        =   96
         Top             =   240
         Width           =   1530
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Siri Produk :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   165
         TabIndex        =   95
         Top             =   1245
         Width           =   1545
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila scan / masukkan data barang yang ingin dijual dalam ruangan di bawah."
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
         Left            =   165
         TabIndex        =   94
         Top             =   615
         Width           =   5145
      End
      Begin VB.Shape Shape1 
         Height          =   1245
         Left            =   120
         Top             =   495
         Width           =   5175
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Berikut adalah nilaian harga emas bagi purity 999.9 untuk pengiraan jualan."
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
         Left            =   240
         TabIndex        =   93
         Top             =   1920
         Width           =   3945
      End
      Begin VB.Shape Shape3 
         Height          =   1245
         Left            =   120
         Top             =   1800
         Width           =   5175
      End
   End
   Begin VB.CommandButton CMD4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SENARAI STOK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3240
      Left            =   14040
      MaskColor       =   &H00FFFF00&
      MouseIcon       =   "Frm115.frx":EFB0
      MousePointer    =   99  'Custom
      Picture         =   "Frm115.frx":F2BA
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox CBB2 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      ItemData        =   "Frm115.frx":11884
      Left            =   17400
      List            =   "Frm115.frx":11886
      Style           =   2  'Dropdown List
      TabIndex        =   57
      Top             =   5160
      Width           =   3975
   End
   Begin VB.CommandButton CDM13 
      Caption         =   "Papar Maklumat Terperinci GST"
      Height          =   360
      Left            =   15480
      MouseIcon       =   "Frm115.frx":11888
      MousePointer    =   99  'Custom
      TabIndex        =   35
      Top             =   3720
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0FF&
      Height          =   2175
      Left            =   18240
      ScaleHeight     =   2115
      ScaleWidth      =   5835
      TabIndex        =   12
      Top             =   7680
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Label L54_Text 
         Caption         =   "L54_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4560
         TabIndex        =   62
         Top             =   1200
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L22_Text 
         Caption         =   "L22_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   34
         Top             =   240
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L23_Text 
         Caption         =   "L23_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   33
         Top             =   480
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L24_Text 
         Caption         =   "L24_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   32
         Top             =   720
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L25_Text 
         Caption         =   "L25_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   31
         Top             =   960
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L21_Text 
         Caption         =   "L21_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L29_Text 
         Caption         =   "L29_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   29
         Top             =   1200
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L31_Text 
         Caption         =   "L31_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L30_Text 
         Caption         =   "L30_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L32_Text 
         Caption         =   "L32_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L33_Text 
         Caption         =   "L33_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         TabIndex        =   25
         Top             =   720
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L40_Text 
         Caption         =   "L40_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   24
         Top             =   1200
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L35_Text 
         Caption         =   "L35_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L39_Text 
         Caption         =   "L39_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   22
         Top             =   960
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L38_Text 
         Caption         =   "L38_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   21
         Top             =   720
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L37_Text 
         Caption         =   "L37_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L36_Text 
         Caption         =   "L36_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L41_Text 
         Caption         =   "L41_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   18
         Top             =   1440
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L42_Text 
         Caption         =   "L42_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   17
         Top             =   1680
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L45_Text 
         Caption         =   "L45_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3240
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L47_Text 
         Caption         =   "L47_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3240
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L49_Text 
         Caption         =   "L49_Text"
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
         Left            =   3240
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label L50_Text 
         Caption         =   "L50_Text"
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
         Left            =   3240
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   645
      End
   End
   Begin VB.Timer Tmr2 
      Interval        =   100
      Left            =   16200
      Top             =   -240
   End
   Begin VB.ComboBox CBB4 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   360
      ItemData        =   "Frm115.frx":11B92
      Left            =   17400
      List            =   "Frm115.frx":11B94
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   5715
      Width           =   3975
   End
   Begin VB.Timer Tmr1 
      Interval        =   10
      Left            =   16680
      Top             =   -240
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   360
      Left            =   17400
      TabIndex        =   7
      Top             =   6120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   16744576
      Format          =   416415744
      CurrentDate     =   41561
   End
   Begin VB.TextBox TB8 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   19680
      MaxLength       =   5
      TabIndex        =   61
      Text            =   "0.00"
      Top             =   1110
      Width           =   1620
   End
   Begin MSComctlLib.ListView LV2 
      Height          =   7050
      Left            =   120
      TabIndex        =   133
      Top             =   3600
      Width           =   15200
      _ExtentX        =   26802
      _ExtentY        =   12435
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
   Begin VB.CommandButton CMD11 
      BackColor       =   &H80000003&
      Caption         =   "Keluar"
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
      Left            =   18480
      MouseIcon       =   "Frm115.frx":11B96
      MousePointer    =   99  'Custom
      Picture         =   "Frm115.frx":11EA0
      Style           =   1  'Graphical
      TabIndex        =   156
      Top             =   10400
      Width           =   2775
   End
   Begin VB.CommandButton CMD10 
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
      Height          =   1095
      Left            =   15600
      MouseIcon       =   "Frm115.frx":1446A
      MousePointer    =   99  'Custom
      Picture         =   "Frm115.frx":14774
      Style           =   1  'Graphical
      TabIndex        =   157
      Top             =   10400
      Width           =   2775
   End
   Begin VB.Label L71_Text 
      BackColor       =   &H8000000C&
      Caption         =   "L71_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   1080
      TabIndex        =   63
      Top             =   10800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label46 
      BackStyle       =   0  'Transparent
      Caption         =   "Kadar Tukaran Mutu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   15360
      TabIndex        =   60
      Top             =   1080
      Width           =   4185
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   18840
      TabIndex        =   59
      Top             =   1080
      Width           =   825
   End
   Begin VB.Label Label43 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier/Agen:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   15480
      TabIndex        =   58
      Top             =   5160
      Width           =   2025
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Ini adalah nilaian harga emas jualan setelah ditukar mutu ke 999.9"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   15480
      TabIndex        =   56
      Top             =   4080
      Width           =   5865
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   15360
      X2              =   21480
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label48 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   18840
      TabIndex        =   55
      Top             =   2280
      Width           =   825
   End
   Begin VB.Label Label47 
      BackStyle       =   0  'Transparent
      Caption         =   "Upah (Tanpa GST)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   15360
      TabIndex        =   54
      Top             =   2280
      Width           =   4185
   End
   Begin VB.Label L51_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   19680
      TabIndex        =   53
      Top             =   2280
      Width           =   3675
   End
   Begin VB.Label Label45 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   18840
      TabIndex        =   52
      Top             =   2760
      Width           =   825
   End
   Begin VB.Label Label44 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah GST"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   15360
      TabIndex        =   51
      Top             =   2760
      Width           =   4185
   End
   Begin VB.Label L52_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   19680
      TabIndex        =   50
      Top             =   2760
      Width           =   3675
   End
   Begin VB.Label Label37 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   18840
      TabIndex        =   48
      Top             =   3240
      Width           =   825
   End
   Begin VB.Label L53_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   19680
      TabIndex        =   47
      Top             =   3240
      Width           =   3675
   End
   Begin VB.Label Label36 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   18840
      TabIndex        =   46
      Top             =   120
      Width           =   825
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Bilangan Barang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   15360
      TabIndex        =   45
      Top             =   120
      Width           =   4185
   End
   Begin VB.Label L43_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   19680
      TabIndex        =   44
      Top             =   120
      Width           =   3675
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(g) :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   18840
      TabIndex        =   43
      Top             =   600
      Width           =   825
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Berat Asal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   15360
      TabIndex        =   42
      Top             =   600
      Width           =   4185
   End
   Begin VB.Label L48_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   19680
      TabIndex        =   41
      Top             =   600
      Width           =   3675
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
      Left            =   12480
      TabIndex        =   39
      Top             =   10680
      Width           =   615
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
      Left            =   11880
      TabIndex        =   38
      Top             =   10680
      Width           =   375
   End
   Begin VB.Label L69_Text 
      BackColor       =   &H8000000C&
      Caption         =   "L69_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   8280
      TabIndex        =   37
      Top             =   10680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label L70_Text 
      BackColor       =   &H8000000C&
      Caption         =   "L70_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   8280
      TabIndex        =   36
      Top             =   10920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Maklumat supplier / agen."
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   15480
      TabIndex        =   11
      Top             =   4800
      Width           =   5865
   End
   Begin VB.Label Label110 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tarikh Jualan   :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   15240
      TabIndex        =   10
      Top             =   6160
      Width           =   2025
   End
   Begin VB.Label Label109 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Pekerja:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   15480
      TabIndex        =   9
      Top             =   5715
      Width           =   2025
   End
   Begin VB.Label Label63 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   18120
      TabIndex        =   5
      Top             =   4320
      Width           =   600
   End
   Begin VB.Label L12_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   18840
      TabIndex        =   4
      Top             =   4320
      Width           =   3225
   End
   Begin VB.Label L9_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   19680
      TabIndex        =   3
      Top             =   1560
      Width           =   3675
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai barang yang dijual."
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
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Width           =   5385
   End
   Begin VB.Label Label35 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(g) :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   18840
      TabIndex        =   1
      Top             =   1560
      Width           =   825
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Berat Jualan 999.9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   15360
      TabIndex        =   0
      Top             =   1560
      Width           =   4185
   End
   Begin VB.Label Label64 
      BackStyle       =   0  'Transparent
      Caption         =   "Nilaian Harga Emas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   15480
      TabIndex        =   6
      Top             =   4320
      Width           =   2505
   End
   Begin VB.Label Label14 
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
      Left            =   10560
      TabIndex        =   40
      Top             =   10680
      Width           =   2295
   End
   Begin VB.Label Label42 
      BackStyle       =   0  'Transparent
      Caption         =   "Upah(Dengan GST)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   15360
      TabIndex        =   49
      Top             =   3240
      Width           =   4185
   End
   Begin VB.Menu Frm115_PM_menu3 
      Caption         =   "Scan Mode (F2)"
      Begin VB.Menu Frm115_scan_mode 
         Caption         =   "Scan Mode"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu Frm115_PM_menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm115_SM_edit_data1 
         Caption         =   "Edit data"
      End
      Begin VB.Menu Frm115_SM_remove_jualan 
         Caption         =   "Keluarkan dari senarai jualan (Pulangkan ke stok kedai)"
      End
   End
   Begin VB.Menu Frm115_PM_menu2 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm115_SM_edit_data2 
         Caption         =   "Edit data"
      End
      Begin VB.Menu Frm115_SM_remove_belian 
         Caption         =   "Keluarkan dari senarai belian"
      End
   End
   Begin VB.Menu Frm115_PM_menu4 
      Caption         =   "select"
      Visible         =   0   'False
      Begin VB.Menu frm115_sm_select_item 
         Caption         =   "Pilih item ini"
      End
   End
End
Attribute VB_Name = "Frm115"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CB2_Click()
'On Error Resume Next
If Frm115.CB2 = 1 Then
    
    Frm115.CB3 = 0
    Frm115.CB4 = 0

End If

Call Frm115_calc2
End Sub
Private Sub CB3_Click()
'On Error Resume Next
If Frm115.CB3 = 1 Then
    
    Frm115.CB2 = 0
    Frm115.CB4 = 0
    
End If

Call Frm115_calc2
End Sub
Private Sub CB4_Click()
'On Error Resume Next
If Frm115.CB4 = 1 Then
    
    Frm115.CB2 = 0
    Frm115.CB3 = 0
    
End If

Call Frm115_calc2
End Sub
Private Sub CDM13_Click()
'On Error Resume Next
Frm115.Frame4.Visible = True
End Sub
Private Sub CMD1_Click()
'On Error Resume Next
Dim Err(30)
Dim Frm115_LM_BERAT_ASAL As Double
Dim Frm115_LM_BERAT_JUAL As Double
Dim Frm115_LM_HARGA_MODAL As Double
Dim Frm115_LM_HARGA_JUAL As Double
Dim Frm115_LM_HARGA_SEMASA_MODAL As Double
Dim Frm115_LM_TETAPANHARGA As Double
Dim Frm115_LM_LIMIT As Double
Dim Frm115_LM_HARGA_STAFF As Double 'Tetapan harga jualan kepada staff
Dim Frm115_LM_HARGA_PELANGGAN As Double 'Tetapan harga jualan kepada pelanggan
Dim Frm115_LM_HARGA_SEMASA As Double 'Harga semasa (jualan)
Dim Frm115_LM_BERAT_JUAL_ASAL As Double 'Berat Jualan (Purity Asal)
Dim Frm115_LM_HARGA_SEMASA_999 As Double 'Harga semasa (jualan) (Purity 999.9)
Dim Frm115_LM_HARGA_SUPPLIER As Double 'Harga per gram (harga semasa) dari supplier (modal)
Dim Frm115_LM_BERAT_999 As Double 'Berat Jualan (Purity Asal)
Dim Frm115_UPAH_MODAL As Double 'Upah modal
Dim Frm115_UPAH_JUAL As Double 'Upah jualan
Dim LM_KADAR_TUKARAN As Double

LM_KADAR_TUKARAN = 0
Frm115_UPAH_MODAL = 0 'Upah modal
Frm115_UPAH_JUAL = 0 'Upah jualan
Frm115_LM_BERAT_JUAL_ASAL = 0 'Berat Jualan (Purity Asal)
Frm115_LM_HARGA_SEMASA_999 = 0 'Harga semasa (jualan) (Purity 999.9)
Frm115_LM_HARGA_SUPPLIER = 0 'Harga per gram (harga semasa) dari supplier (modal)
Frm115_LM_BERAT_999 = 0 'Berat Jualan (Purity Asal)

Frm115_LM_HARGA_SEMASA = 0 'Harga semasa (jualan)
Frm115_LM_HARGA_SUPPLIER = 0 'Harga per gram (harga semasa) dari supplier (modal)

x = 0
Frm115_LM_BERAT_ASAL = 0
Frm115_LM_BERAT_JUAL = 0
Frm115_LM_DATA_SAVE = 0
Frm115_LM_HARGA_MODAL = 0
Frm115_LM_HARGA_JUAL = 0
Frm115_LM_HARGA_SEMASA_MODAL = 0
Frm115_LM_PRICE_CHECK = 0 '0 : Tidak Perlu Periksa Harga Semasa Jualan , 1 : Perlu Periksa Harga Semasa Jualan
Frm115_LM_TETAPANHARGA = 0
Frm115_LM_LIMIT = 0
Frm115_LM_HARGA_STAFF = 0
Frm115_LM_HARGA_PELANGGAN = 0

If Frm115.L3_Text = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan [No. Siri Produk]."
End If
If Frm115.L33_Text = vbNullString Then
    x = x + 1
    Err(x) = "Maklumat harga semasa modal belian item ini yang tidak sah. Sila keluar dari menu ini dan cuba sekali lagi."
End If
If Frm115.L50_Text = vbNullString Or (Frm115.L50_Text <> vbNullString And Not IsNumeric(Frm115.L50_Text)) Then
    x = x + 1
    Err(x) = "Maklumat upah modal yang tidak sah. Sila keluar dari menu ini dan cuba sekali lagi."
End If
If Frm115.L6_Text = vbNullString Or (Frm115.L6_Text <> vbNullString And Not IsNumeric(Frm115.L6_Text)) Then
    x = x + 1
    Err(x) = "Sila maklumat [Berat Asal]. Sila scan item sekali lagi."
End If
If Frm115.TB3 = vbNullString Or (Frm115.TB3 <> vbNullString And Not IsNumeric(Frm115.TB3)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Berat Jualan]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm115.TB2 = vbNullString Or (Frm115.TB2 <> vbNullString And Not IsNumeric(Frm115.TB2)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Harga Semasa Emas 999.9]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm115.TB2 <> vbNullString And IsNumeric(Frm115.TB2) Then

    If Format(Frm115.TB2, "0.00") = "0.00" Then
        x = x + 1
        Err(x) = "Harga emas semasa 999.9 yang tidak sah. Nilai 0.00 tidak dibenarkan."
    End If
    
End If
If Frm115.TB7 = vbNullString Or (Frm115.TB7 <> vbNullString And Not IsNumeric(Frm115.TB7)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar Tukaran Purity 999.9]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If (Frm115.TB7 <> vbNullString And IsNumeric(Frm115.TB7)) Then
    
    LM_KADAR_TUKARAN = Frm115.TB7
    
    If LM_KADAR_TUKARAN > 1 Then
    
        x = x + 1
        Err(x) = "[Kadar Tukaran Purity 999.9] tidak boleh lebih dari 1.00. Hanya NOMBOR dibenarkan dalam ruangan ini."
        
    End If
    
End If
If Frm115.L7_Text = vbNullString Or (Frm115.L7_Text <> vbNullString And Not IsNumeric(Frm115.L7_Text)) Then
    x = x + 1
    Err(x) = "[Berat 999.9] yang tidak sah. Sila scan item sekali lagi."
End If
If Frm115.TB4 = vbNullString Or (Frm115.TB4 <> vbNullString And Not IsNumeric(Frm115.TB4)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Upah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm115.CB2 = 0 And Frm115.CB3 = 0 And Frm115.CB4 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan jenis GST bagi upah."
End If
If Frm115.TB5 = vbNullString Or Frm115.TB6 = vbNullString Then
    x = x + 1
    Err(x) = "Maklumat berkenaan GST yang tidak sah. Sila keluar dari menu ini dan cuba sekali lagi."
End If

If (Frm115.L6_Text <> vbNullString And IsNumeric(Frm115.L6_Text)) And (Frm115.TB3 <> vbNullString And IsNumeric(Frm115.TB3)) Then
    Frm115_LM_BERAT_ASAL = Frm115.L6_Text 'Berat Asal
    Frm115_LM_BERAT_JUAL = Frm115.TB3 'Berat Jualan
    
    If Frm115_LM_BERAT_JUAL > Frm115_LM_BERAT_ASAL Then
        x = x + 1
        Err(x) = "Berat jualan melebihi berat asal."
    End If
End If
If Frm115.L49_Text = vbNullString Or (Frm115.L49_Text <> vbNullString And Not IsNumeric(Frm115.L49_Text)) Then
    x = x + 1
    Err(x) = "Tiada maklumat harga dari supplier bagi purity ini." & vbCrLf & _
                "Sila pastikan harga dari supplier bagi purity ini telah ditetapkan dalam TETAPAN HARIAN SISTEM."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah anda ingin masukkan item ini ke dalam senarai jualan ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
'### Periksa Data Dulang ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where no_siri_Produk='" & Frm115.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!dulang) Then Frm115_LM_DULANG = rs!dulang 'Dulang
        End If
        
        rs.Close
        Set rs = Nothing
'### Periksa Data Dulang ### - End
        
'### Masukkan Data Ke Dalam Temp Table ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_GDN_TEMP & " where no_siri_Produk='" & Frm115.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
        
            rs.AddNew
            If Frm115.L3_Text <> vbNullString Then
                rs!no_siri_Produk = Frm115.L3_Text 'No. Siri Produk
            Else
                rs!no_siri_Produk = Null 'No. Siri Produk
            End If
            If Frm115.L5_Text <> vbNullString Then
                rs!kategori_Produk = Frm115.L5_Text 'Kategori Produk
            Else
                rs!kategori_Produk = Null 'Kategori Produk
            End If
            If Frm115.L4_Text <> vbNullString Then
                rs!purity = Frm115.L4_Text 'Purity
            Else
                rs!purity = Null 'Purity
            End If
            If Frm115.L6_Text <> vbNullString Then
                rs!Berat_Asal = Format(Frm115.L6_Text, "0.00") 'Berat Asal (g)
            Else
                rs!Berat_Asal = Null 'Berat Asal (g)
            End If
            If Frm115.TB3 <> vbNullString Then
                rs!berat_jualan = Format(Frm115.TB3, "0.00") 'Berat Jualan (g)
            Else
                rs!berat_jualan = Null 'Berat Jualan (g)
            End If
            If Frm115.TB2 <> vbNullString Then
                rs!harga_Semasa = Format(Frm115.TB2, "0.00") 'Harga Semasa (RM/g)
            Else
                rs!harga_Semasa = Null 'Harga Semasa (RM/g)
            End If
            If Frm115.TB4 <> vbNullString Then
                rs!UPAH = Format(Frm115.TB4, "0.00") 'Upah (RM)
            Else
                rs!UPAH = Null 'Upah (RM)
            End If
            
            Frm115_LM_HARGA_SEMASA = Frm115.TB2 'Harga emas semasa 999.9 (Untuk tujuan jualan kepada pelanggan)
            Frm115_LM_BERAT_JUALAN_9999 = Frm115.L7_Text 'Berat jualan dalam purity 999.9
            Frm115_LM_UPAH_DAN_GST = Frm115.TB6 'Jumlah Upah + GST (Bagi jualan setiap item)

            If Frm115.TB6 <> vbNullString Then
                rs!harga_asal = Format((Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            Else
                rs!harga_asal = Null 'Harga Asal Item (RM)
            End If
            
            rs!diskaun = "0.00" 'Diskaun (%)
            rs!harga_lepas_diskaun = Format((Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            rs!adjustment = Format(0, "0.00") 'Adjustment (RM)
            rs!harga_jualan = Format((Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            rs!harga_jualan_dengan_gst = Format((Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            
            If Frm115.CB2 = 1 Then
                rs!gst_ari_nashi = "ZR (L)" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                rs!kadar_gst = Null 'Kadar Cukai GST (%)
                If Frm115.TB5 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm115.TB5, "0.00") 'Jumlah GST (Bagi jualan setiap item)
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah GST (Bagi jualan setiap item)
                End If
            ElseIf Frm115.CB3 = 1 Or Frm115.CB4 = 1 Then
                rs!gst_ari_nashi = "SR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                If Frm115.L21_Text <> vbNullString Then
                    rs!kadar_gst = Frm115.L21_Text 'Kadar Cukai GST (%)
                Else
                    rs!kadar_gst = "0" 'Jumlah Cukai GST (RM)
                End If
                If Frm115.TB5 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm115.TB5, "0.00") 'Jumlah GST (Bagi jualan setiap item)
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah GST (Bagi jualan setiap item)
                End If
                If Frm115.CB4 = 1 Then 'Jenis Cukai GST SR
                    rs!gst_include = 1 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                Else
                    rs!gst_include = 0 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                End If
            End If
            If Frm115.L30_Text <> vbNullString Then
                rs!harga_tanpa_gst = Format(Frm115.L30_Text, "0.00") 'Harga Jualan Tanpa GST (RM)
            Else
                rs!harga_tanpa_gst = Null 'Harga Jualan Tanpa GST (RM)
            End If
            If Frm115.TB6 <> vbNullString Then
                rs!harga_dengan_gst = Format(Frm115.TB6, "0.00") 'Harga Jualan Termasuk GST (RM)
            Else
                rs!harga_dengan_gst = Null 'Harga Jualan Termasuk GST (RM)
            End If
            rs!dropship = 0 '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
            rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g)
            rs!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini
            If Frm115.L32_Text = "0" Then '0 : Menu Data Baru , 1 : Menu Edit Data
                rs!Status = 1
            ElseIf Frm115.L32_Text = "1" Then
                rs!Status = 4
            End If
            rs!Type = 0 '0 : BK , 1 : Barang Permata
            If Frm115.L33_Text <> vbNullString Then
                rs!harga_per_gram_modal = Format(Frm115.L33_Text, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
                Frm115_LM_HARGA_SEMASA_MODAL = Frm115.L33_Text
            Else
                rs!harga_per_gram_modal = Format(0, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
            End If
            rs!modal = Format(Frm115_LM_HARGA_SEMASA_MODAL * Frm115_LM_BERAT_JUALAN_9999, "0.00") 'Harga Modal (RM)
            If IsNumeric(Frm115.TB6) And IsNumeric(Frm115.L33_Text) And IsNumeric(Frm115.TB3) Then
                Frm115_LM_HARGA_MODAL = Frm115.L33_Text * Frm115.TB3 'Harga modal
                Frm115_LM_HARGA_JUAL = (Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST 'Harga jualan
                
                rs!untung = Format(Frm115_LM_HARGA_JUAL - Frm115_LM_HARGA_MODAL, "0.00") 'Jumlah Keuntungan
            Else
                rs!untung = Format(0, "0.00") 'Jumlah Keuntungan
            End If
            
            If Frm115.L49_Text <> vbNullString Then 'Harga per gram (harga semasa) dari supplier (modal)
                rs!harga_per_gram_supplier = Frm115.L49_Text
            Else
                rs!harga_per_gram_supplier = 0
            End If
            
            If IsNumeric(Frm115.TB3) And IsNumeric(Frm115.TB2) And IsNumeric(Frm115.L49_Text) And IsNumeric(Frm115.L7_Text) And IsNumeric(Frm115.L6_Text) And IsNumeric(Frm115.L50_Text) And IsNumeric(Frm115.TB4) Then
                Frm115_LM_BERAT_JUAL_ASAL = Frm115.TB3 'Berat Jualan (Purity Asal)
                Frm115_LM_BERAT_ASAL = Frm115.L6_Text 'Berat Asal (Purity Asal)
                Frm115_UPAH_JUAL = Frm115.TB4 'Upah jualan
                Frm115_UPAH_MODAL = Frm115.L50_Text 'Upah modal
                Frm115_LM_HARGA_SEMASA_999 = Frm115.TB2 'Harga semasa (jualan) (Purity 999.9)
                Frm115_LM_HARGA_SUPPLIER = Frm115.L49_Text 'Harga per gram (harga semasa) dari supplier (modal)
                Frm115_LM_BERAT_999 = Frm115.L7_Text 'Berat emas dalam purity 999.9
                
                rs!upah_modal = Frm115.L50_Text 'Upah modal
                rs!harga_per_gram_supplier = Frm115.L49_Text 'Harga per gram (harga semasa) dari supplier (modal)
                rs!untung2 = Format(((Frm115_LM_BERAT_999 * Frm115_LM_HARGA_SEMASA_999) + Frm115_UPAH_JUAL) - ((Frm115_LM_BERAT_JUAL_ASAL * Frm115_LM_HARGA_SUPPLIER) + (Frm115_LM_BERAT_JUAL_ASAL / Frm115_LM_BERAT_ASAL) * Frm115_UPAH_MODAL), "0.00") 'Untung jika restok pada harga supplier ini
                
            Else
            
                rs!harga_per_gram_supplier = "0.00" 'Harga per gram (harga semasa) dari supplier (modal)
                rs!untung2 = "0.00" 'Untung jika restok pada harga supplier ini
                rs!upah_modal = "0.00" 'Upah modal
                
            End If
                
            If Format(Frm115.L6_Text, "0.00") = Format(Frm115.TB3, "0.00") Then
                rs!potong_flag = 0 '0 : Tiada Potong , 1 : Ada Potong
            Else
                rs!potong_flag = 1 '0 : Tiada Potong , 1 : Ada Potong
            End If
            rs!dulang = Frm115_LM_DULANG 'Dulang
            If Frm115.TB7 <> vbNullString Then
                rs!pemalar_tukaran_999 = Frm115.TB7 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
            Else
                rs!pemalar_tukaran_999 = Format(0, "0.00") 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
            End If
            If Frm115.L7_Text <> vbNullString Then
                rs!berat_999 = Format(Frm115.L7_Text, "0.00") 'Berat jualan dalam purity 999.9
            Else
                rs!berat_999 = Null 'Berat jualan dalam purity 999.9
            End If
            rs!gst_barang_atau_upah = 1 '0 : GST pada harga jualan , 1 : GST pada upah
            rs!gdn_temp = 1
            
            rs.Update
            Frm115_LM_DATA_SAVE = 1
        Else
            If Frm115.L3_Text <> vbNullString Then
                rs!no_siri_Produk = Frm115.L3_Text 'No. Siri Produk
            Else
                rs!no_siri_Produk = Null 'No. Siri Produk
            End If
            If Frm115.L5_Text <> vbNullString Then
                rs!kategori_Produk = Frm115.L5_Text 'Kategori Produk
            Else
                rs!kategori_Produk = Null 'Kategori Produk
            End If
            If Frm115.L4_Text <> vbNullString Then
                rs!purity = Frm115.L4_Text 'Purity
            Else
                rs!purity = Null 'Purity
            End If
            If Frm115.L6_Text <> vbNullString Then
                rs!Berat_Asal = Format(Frm115.L6_Text, "0.00") 'Berat Asal (g)
            Else
                rs!Berat_Asal = Null 'Berat Asal (g)
            End If
            If Frm115.TB3 <> vbNullString Then
                rs!berat_jualan = Format(Frm115.TB3, "0.00") 'Berat Jualan (g)
            Else
                rs!berat_jualan = Null 'Berat Jualan (g)
            End If
            If Frm115.TB2 <> vbNullString Then
                rs!harga_Semasa = Format(Frm115.TB2, "0.00") 'Harga Semasa (RM/g)
            Else
                rs!harga_Semasa = Null 'Harga Semasa (RM/g)
            End If
            If Frm115.TB4 <> vbNullString Then
                rs!UPAH = Format(Frm115.TB4, "0.00") 'Upah (RM)
            Else
                rs!UPAH = Null 'Upah (RM)
            End If
            
            Frm115_LM_HARGA_SEMASA = Frm115.TB2 'Harga emas semasa 999.9 (Untuk tujuan jualan kepada pelanggan)
            Frm115_LM_BERAT_JUALAN_9999 = Frm115.L7_Text 'Berat jualan dalam purity 999.9
            Frm115_LM_UPAH_DAN_GST = Frm115.TB6 'Jumlah Upah + GST (Bagi jualan setiap item)

            If Frm115.TB6 <> vbNullString Then
                rs!harga_asal = Format((Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            Else
                rs!harga_asal = Null 'Harga Asal Item (RM)
            End If
            
            rs!diskaun = "0.00" 'Diskaun (%)
            rs!harga_lepas_diskaun = Format((Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            rs!adjustment = Format(0, "0.00") 'Adjustment (RM)
            rs!harga_jualan = Format((Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            rs!harga_jualan_dengan_gst = Format((Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            
            If Frm115.CB2 = 1 Then
                rs!gst_ari_nashi = "ZR (L)" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                rs!kadar_gst = Null 'Kadar Cukai GST (%)
                If Frm115.TB5 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm115.TB5, "0.00") 'Jumlah GST (Bagi jualan setiap item)
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah GST (Bagi jualan setiap item)
                End If
            ElseIf Frm115.CB3 = 1 Then
                rs!gst_ari_nashi = "SR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                If Frm115.L21_Text <> vbNullString Then
                    rs!kadar_gst = Frm115.L21_Text 'Kadar Cukai GST (%)
                Else
                    rs!kadar_gst = "0" 'Jumlah Cukai GST (RM)
                End If
                If Frm115.TB5 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm115.TB5, "0.00") 'Jumlah GST (Bagi jualan setiap item)
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah GST (Bagi jualan setiap item)
                End If
                If Frm115.CB4 = 1 Then 'Jenis Cukai GST SR
                    rs!gst_include = 1 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                Else
                    rs!gst_include = 0 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                End If
            End If
            If Frm115.L30_Text <> vbNullString Then
                rs!harga_tanpa_gst = Format(Frm115.L30_Text, "0.00") 'Harga Jualan Tanpa GST (RM)
            Else
                rs!harga_tanpa_gst = Null 'Harga Jualan Tanpa GST (RM)
            End If
            If Frm115.TB6 <> vbNullString Then
                rs!harga_dengan_gst = Format(Frm115.TB6, "0.00") 'Harga Jualan Termasuk GST (RM)
            Else
                rs!harga_dengan_gst = Null 'Harga Jualan Termasuk GST (RM)
            End If
            rs!dropship = 0 '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
            rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g)
            rs!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini
            If Frm115.L32_Text = "0" Then '0 : Menu Data Baru , 1 : Menu Edit Data
                rs!Status = 1
            ElseIf Frm115.L32_Text = "1" Then
                rs!Status = 3
            End If
            rs!Type = 0 '0 : BK , 1 : Barang Permata
            If Frm115.L33_Text <> vbNullString Then
                rs!harga_per_gram_modal = Format(Frm115.L33_Text, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
                Frm115_LM_HARGA_SEMASA_MODAL = Frm115.L33_Text
            Else
                rs!harga_per_gram_modal = Format(0, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
            End If
            rs!modal = Format(Frm115_LM_HARGA_SEMASA_MODAL * Frm115_LM_BERAT_JUALAN_9999, "0.00") 'Harga Modal (RM)
            If IsNumeric(Frm115.TB6) And IsNumeric(Frm115.L33_Text) And IsNumeric(Frm115.TB3) Then
                Frm115_LM_HARGA_MODAL = Frm115.L33_Text * Frm115.TB3 'Harga modal
                Frm115_LM_HARGA_JUAL = (Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST 'Harga jualan
                
                rs!untung = Format(Frm115_LM_HARGA_JUAL - Frm115_LM_HARGA_MODAL, "0.00") 'Jumlah Keuntungan
            Else
                rs!untung = Format(0, "0.00") 'Jumlah Keuntungan
            End If
            If Frm115.L49_Text <> vbNullString Then 'Harga per gram (harga semasa) dari supplier (modal)
                rs!harga_per_gram_supplier = Frm115.L49_Text
            Else
                rs!harga_per_gram_supplier = 0
            End If
            
            If IsNumeric(Frm115.TB3) And IsNumeric(Frm115.TB2) And IsNumeric(Frm115.L49_Text) And IsNumeric(Frm115.L7_Text) And IsNumeric(Frm115.L6_Text) And IsNumeric(Frm115.L50_Text) And IsNumeric(Frm115.TB4) Then
                Frm115_LM_BERAT_JUAL_ASAL = Frm115.TB3 'Berat Jualan (Purity Asal)
                Frm115_LM_BERAT_ASAL = Frm115.L6_Text 'Berat Asal (Purity Asal)
                Frm115_UPAH_JUAL = Frm115.TB4 'Upah jualan
                Frm115_UPAH_MODAL = Frm115.L50_Text 'Upah modal
                Frm115_LM_HARGA_SEMASA_999 = Frm115.TB2 'Harga semasa (jualan) (Purity 999.9)
                Frm115_LM_HARGA_SUPPLIER = Frm115.L49_Text 'Harga per gram (harga semasa) dari supplier (modal)
                Frm115_LM_BERAT_999 = Frm115.L7_Text 'Berat emas dalam purity 999.9
                
                rs!upah_modal = Frm115.L50_Text 'Upah modal
                rs!harga_per_gram_supplier = Frm115.L49_Text 'Harga per gram (harga semasa) dari supplier (modal)
                rs!untung2 = Format(((Frm115_LM_BERAT_999 * Frm115_LM_HARGA_SEMASA_999) + Frm115_UPAH_JUAL) - ((Frm115_LM_BERAT_JUAL_ASAL * Frm115_LM_HARGA_SUPPLIER) + (Frm115_LM_BERAT_JUAL_ASAL / Frm115_LM_BERAT_ASAL) * Frm115_UPAH_MODAL), "0.00") 'Untung jika restok pada harga supplier ini
                
            Else
            
                rs!harga_per_gram_supplier = "0.00" 'Harga per gram (harga semasa) dari supplier (modal)
                rs!untung2 = "0.00" 'Untung jika restok pada harga supplier ini
                rs!upah_modal = "0.00" 'Upah modal
                
            End If
            
            If Format(Frm115.L6_Text, "0.00") = Format(Frm115.TB3, "0.00") Then
                rs!potong_flag = 0 '0 : Tiada Potong , 1 : Ada Potong
            Else
                rs!potong_flag = 1 '0 : Tiada Potong , 1 : Ada Potong
            End If
            rs!dulang = Frm115_LM_DULANG 'Dulang
            If Frm115.TB7 <> vbNullString Then
                rs!pemalar_tukaran_999 = Frm115.TB7 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
            Else
                rs!pemalar_tukaran_999 = Format(0, "0.00") 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
            End If
            If Frm115.L7_Text <> vbNullString Then
                rs!berat_999 = Format(Frm115.L7_Text, "0.00") 'Berat jualan dalam purity 999.9
            Else
                rs!berat_999 = Null 'Berat jualan dalam purity 999.9
            End If
            rs!gst_barang_atau_upah = 1 '0 : GST pada harga jualan , 1 : GST pada upah
            rs!gdn_temp = 1
            rs.Update
            Frm115_LM_DATA_SAVE = 1
        End If
        
        rs.Close
        Set rs = Nothing
'### Masukkan Data Ke Dalam Temp Table ### - End
        
        If Frm115_LM_DATA_SAVE = 1 Then
        
            GM_NEXT_PREV = 0
            
            Frm115.L69_Text = -1 'Titik Pencarian Data
            Frm115.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
            Frm115.L67_Text = 0 'Paparan Page ke-xxx

            Call Frm115_reset_1
            Call Frm115_Senarai_Jualan_Header
            Call Frm115_Senarai_Jualan
            
            MsgBox "Data telah berjaya dimasukkan ke dalam senarai jualan.", vbInformation, "Info"
            
            Frm115.TB1.SetFocus
        End If
    End If
End If
End Sub
Private Sub CMD10_Click()
'On Error Resume Next
Dim Err(30)
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim Frm115_LM_CUKAI_ZR As Double
Dim Frm115_LM_CUKAI_SR As Double
Dim Frm115_LM_BERAT_ASAL As Double
Dim Frm115_LM_BEZA_BERAT As Double
Dim Frm115_LM_BERAT_RETURN As Double
Dim Frm115_LM_BERAT_JUALAN As Double
Dim LM_KADAR_TUKARAN As Double
Dim Frm115_SUSUT_BERAT As Double

Frm115_SUSUT_BERAT = 0
LM_KADAR_TUKARAN = 0
Frm115_LM_BERAT_ASAL = 0 'Berat Asal (g)
Frm115_LM_BERAT_JUALAN = 0 'Berat Jualan (g)
Frm115_LM_CUKAI_ZR = 0 'Jumlah cukai GST ZR
Frm115_LM_CUKAI_SR = 0 'Jumlah cukai GST SR

If Frm115.L43_Text = "0" Then
    x = x + 1
    Err(x) = "Tiada senarai jualan."
End If
If Frm115.CBB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih supplier/agen yang membuat belian ini."
End If
If Frm115.CBB4 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih nama pekerja."
End If
If Frm115.TB8 = vbNullString Or (Frm115.TB8 <> vbNullString And Not IsNumeric(Frm115.TB8)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar Tukaran Purity 999.9]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If (Frm115.TB8 <> vbNullString And IsNumeric(Frm115.TB8)) Then
    
    LM_KADAR_TUKARAN = Frm115.TB8
    
    If LM_KADAR_TUKARAN > 1 Then
    
        x = x + 1
        Err(x) = "[Kadar Tukaran Purity 999.9] tidak boleh lebih dari 1.00. Hanya NOMBOR dibenarkan dalam ruangan ini."
        
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

    Note = "Adakah anda yakin untuk simpan data yang telah diedit?" & vbCrLf & _
            vbNullString & vbCrLf & _
            "Data jualan akan disimpan ke dalam sistem."

    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
        If Frm115.CBB4 <> vbNullString Then
        
            Frm115_LM_EMP_NAMA = Split(Frm115.CBB4, "  |  ")(0)
            Frm115_LM_EMP_NO = Split(Frm115.CBB4, "  |  ")(1)
            
        End If
    
'### Masukkan maklumat Good Delivery Note (GDN) ### - Start
        LM_NOW = Now
        LM_TARIKH = DateTime.Date$
        LM_MASA = DateTime.Time$

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 77_gdn_grn where no_rujukan='" & Frm115.L23_Text & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then

            G_ID = rs!ID
            Call recovery_77_gdn_grn
                    
            rs!tarikh = Frm115.DTPicker1
            rs!masa = LM_MASA
            'rs!write_timestamp = LM_NOW
            
            If Frm115.L48_Text <> vbNullString Then
                rs!Berat_Asal = Format(Frm115.L48_Text, "0.00")
            Else
                rs!Berat_Asal = "0.00"
            End If
            If Frm115.TB8 <> vbNullString Then
                rs!kadar_tukaran = Frm115.TB8
            Else
                rs!kadar_tukaran = "0.00"
            End If
            If Frm115.L9_Text <> vbNullString Then
                rs!berat_tukaran = Format(Frm115.L9_Text, "0.00")
            Else
                rs!berat_tukaran = Null
            End If
            If Frm115.L51_Text <> vbNullString Then
                rs!harga_tanpa_gst = Format(Frm115.L51_Text, "0.00")
            Else
                rs!harga_tanpa_gst = "0.00"
            End If
            If Frm115.L52_Text <> vbNullString Then
                rs!jumlah_gst = Format(Frm115.L52_Text, "0.00")
            Else
                rs!jumlah_gst = "0.00"
            End If
            If Frm115.L21_Text <> vbNullString Then
                rs!kadar_gst = Format(Frm115.L21_Text, "0.00")
            Else
                rs!kadar_gst = "0.00"
            End If
            If Frm115.L53_Text <> vbNullString Then
                rs!harga_dengan_gst = Format(Frm115.L53_Text, "0.00")
            Else
                rs!harga_dengan_gst = Null
            End If
            If Frm115.TB2 <> vbNullString Then
                rs!harga_999 = Format(Frm115.TB2, "0.00")
            Else
                rs!harga_999 = "0.00"
            End If
            If Frm115.L12_Text <> vbNullString Then
                rs!nilaian_harga_emas = Format(Frm115.L12_Text, "0.00")
            Else
                rs!nilaian_harga_emas = "0.00"
            End If
            If Frm115.L17_Text <> vbNullString Then
                rs!gst_zr_harga = Format(Frm115.L17_Text, "0.00")
            Else
                rs!gst_zr_harga = "0.00"
            End If
            If Frm115.L18_Text <> vbNullString Then
                rs!gst_sr_harga = Format(Frm115.L18_Text, "0.00")
            Else
                rs!gst_sr_harga = "0.00"
            End If
            If Frm115.L19_Text <> vbNullString Then
                rs!gst_zr_cukai = Format(Frm115.L19_Text, "0.00")
            Else
                rs!gst_zr_cukai = "0.00"
            End If
            If Frm115.L20_Text <> vbNullString Then
                rs!gst_sr_cukai = Format(Frm115.L20_Text, "0.00")
            Else
                rs!gst_sr_cukai = "0.00"
            End If
            If Frm115.L43_Text <> vbNullString Then
                rs!bil_barang = Frm115.L43_Text
            Else
                rs!bil_barang = 0
            End If
            rs!Status = 1
            rs!jenis_urusan = 0
            rs!terminal = G_TERMINAL
            If Frm115.CBB2 <> vbNullString Then
                rs!supplier_agen = Frm115.CBB2
            Else
                rs!supplier_agen = Null
            End If
            rs!user = Frm115_LM_EMP_NAMA 'Nama Pekerja
            rs.Update
            DATA_SAVE = 1
            
        End If
        
        rs.Close
        Set rs = Nothing
'### Masukkan data voucher / invoice bagi belian agen ini ### - End

'### Masukkan Data Jualan Ke Dalam Table Jualan ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_GDN_TEMP & "", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
        
            Frm115_LM_BERAT_ASAL = 0
            Frm115_LM_BEZA_BERAT = 0
            Frm115_LM_BERAT_RETURN = 0
            
'########### Kemasukan data baru dalam senarai ##############- Start
            If rs!Status = 4 Then

                Set rs1 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs1.Open "select * from 23_senarai_jualan", cn, adOpenKeyset, adLockOptimistic
                
                rs1.AddNew
                rs1!tarikh = Frm115.DTPicker1 'Tarikh Jualan
                rs1!no_resit = Frm115.L23_Text 'No. Invoice Jualan
                If Not IsNull(rs!no_siri_Produk) Then
                    rs1!no_siri_Produk = rs!no_siri_Produk 'No. Siri Produk
                Else
                    rs1!no_siri_Produk = Null 'No. Siri Produk
                End If
                
                If Not IsNull(rs!kategori_Produk) Then
                    rs1!kategori_Produk = rs!kategori_Produk 'Kategori Produk
                Else
                    rs1!no_siri_Produk = Null 'Kategori Produk
                End If
                If Not IsNull(rs!purity) Then
                    rs1!purity = rs!purity 'Purity
                Else
                    rs1!purity = Null 'Purity
                End If
                If Not IsNull(rs!Berat_Asal) Then
                    rs1!Berat_Asal = rs!Berat_Asal 'Berat Asal (g)
                Else
                    rs1!Berat_Asal = Null 'Berat Asal (g)
                End If
                If Not IsNull(rs!berat_jualan) Then
                    rs1!berat_jualan = Format(rs!berat_jualan, "0.00") 'Berat Jualan (g)
                Else
                    rs1!berat_jualan = Null 'Berat Jualan (g)
                End If
                If Not IsNull(rs!harga_Semasa) Then
                    rs1!harga_Semasa = Format(rs!harga_Semasa, "0.00") 'Harga Semasa (RM/g)
                Else
                    rs1!harga_Semasa = Null 'Harga Semasa (RM/g)
                End If
                If Not IsNull(rs!UPAH) Then
                    rs1!UPAH = Format(rs!UPAH, "0.00") 'Upah (RM)
                Else
                    rs1!UPAH = Null 'Upah (RM)
                End If
                If Not IsNull(rs!harga_asal) Then
                    rs1!harga_asal = Format(rs!harga_asal, "0.00") 'Harga Asal Item (RM)
                Else
                    rs1!harga_asal = Null 'Harga Asal Item (RM)
                End If
                If Not IsNull(rs!diskaun) Then
                    rs1!diskaun = Format(rs!diskaun, "0.00") 'Diskaun (%)
                Else
                    rs1!diskaun = Null 'Diskaun (%)
                End If
                If Not IsNull(rs!harga_lepas_diskaun) Then
                    rs1!harga_lepas_diskaun = Format(rs!harga_lepas_diskaun, "0.00") 'Harga Selepas Diskaun (RM)
                Else
                    rs1!harga_lepas_diskaun = Null 'Harga Selepas Diskaun (RM)
                End If
                If Not IsNull(rs!adjustment) Then
                    rs1!adjustment = Format(rs!adjustment, "0.00") 'Harga Selepas Diskaun (RM)
                Else
                    rs1!adjustment = Null 'Harga Selepas Diskaun (RM)
                End If
                If Not IsNull(rs!harga_jualan) Then
                    rs1!harga_jualan = Format(rs!harga_jualan, "0.00") 'Harga Jualan (RM)
                Else
                    rs1!harga_jualan = Null 'Harga Jualan (RM)
                End If
                If Not IsNull(rs!gst_ari_nashi) Then
                    rs1!gst_ari_nashi = rs!gst_ari_nashi '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                Else
                    rs1!gst_ari_nashi = Null '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                End If
                If Not IsNull(rs!kadar_gst) Then
                    rs1!kadar_gst = rs!kadar_gst 'Kadar Cukai GST (%)
                Else
                    rs1!kadar_gst = Null 'Kadar Cukai GST (%)
                End If
                If Not IsNull(rs!jumlah_gst) Then
                    rs1!jumlah_gst = Format(rs!jumlah_gst, "0.00") 'Jumlah Cukai GST (RM)
                Else
                    rs1!jumlah_gst = Null 'Jumlah Cukai GST (RM)
                End If
                If Not IsNull(rs!harga_dengan_gst) Then
                    rs1!harga_dengan_gst = Format(rs!harga_dengan_gst, "0.00") 'Harga Jualan Termasuk GST (RM)
                Else
                    rs1!harga_dengan_gst = Null 'Harga Jualan Termasuk GST (RM)
                End If
                If Not IsNull(rs!dropship) Then
                    rs1!dropship = rs!dropship '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
                Else
                    rs1!dropship = Null '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
                End If
                If Not IsNull(rs!komisyen_per_gram) Then
                    rs1!komisyen_per_gram = Format(rs!komisyen_per_gram, "0.00") 'Komisyen Per Gram Dropship (RM/g) : BK
                Else
                    rs1!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g) : BK
                End If
                If Not IsNull(rs!jumlah_komisyen) Then
                    rs1!jumlah_komisyen = Format(rs!jumlah_komisyen, "0.00") 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
                Else
                    rs1!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
                End If
                If Not IsNull(rs!harga_per_gram_modal) Then
                    rs1!harga_per_gram_modal = Format(rs!harga_per_gram_modal, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
                Else
                    rs1!harga_per_gram_modal = Null 'Harga Per Gram Bagi Modal (RM/g)
                End If
                If Not IsNull(rs!modal) Then
                    rs1!modal = Format(rs!modal, "0.00") 'Harga Modal (RM)
                Else
                    rs1!modal = Null 'Harga Modal (RM)
                End If
                If Not IsNull(rs!untung) Then
                    rs1!untung = Format(rs!untung, "0.00") 'Jumlah Keuntungan
                Else
                    rs1!untung = Null 'Jumlah Keuntungan
                End If
                If Not IsNull(rs!harga_per_gram_supplier) Then
                    rs1!harga_per_gram_supplier = Format(rs!harga_per_gram_supplier, "0.00") 'Harga per gram (harga semasa) dari supplier (modal)
                Else
                    rs1!harga_per_gram_supplier = Null 'Harga per gram (harga semasa) dari supplier (modal)
                End If
                If Not IsNull(rs!upah_modal) Then
                    rs1!upah_modal = Format(rs!upah_modal, "0.00") 'Upah modal
                Else
                    rs1!upah_modal = Null 'Upah modal
                End If
                If Not IsNull(rs!untung2) Then
                    rs1!untung2 = Format(rs!untung2, "0.00") 'Jumlah Keuntungan
                Else
                    rs1!untung2 = Null 'Jumlah Keuntungan
                End If
                If Not IsNull(rs!dulang) Then
                    rs1!dulang = rs!dulang 'Dulang
                Else
                    rs1!dulang = Null 'Dulang
                End If
                If Not IsNull(rs!potong_flag) Then
                    rs1!potong_flag = rs!potong_flag '0 : Tiada Potong , 1 : Ada Potong
                    If rs!potong_flag = 0 Then
                        rs1!Status = 0 '0 : Jualan Biasa , 1 : Jualan Secara Potong , 2 : Tempahan , 3 : Ansuran , 4 : ETA
                    Else
                        rs1!Status = 1 '0 : Jualan Biasa , 1 : Jualan Secara Potong , 2 : Tempahan , 3 : Ansuran , 4 : ETA
                    End If
                Else
                    rs1!potong_flag = Null '0 : Tiada Potong , 1 : Ada Potong
                End If

                If Not IsNull(rs!Type) Then
                    rs1!Type = rs!Type '0 : BK , 1 : Barang Permata
                Else
                    rs1!Type = Null '0 : BK , 1 : Barang Permata
                End If
                
                rs1!jualan_online = 0
                rs1!bil_rasmi = 1
                rs1!status_r = 0
                If Frm115.CBB4 <> vbNullString Then
                    Frm115_LM_EMP_NO = Split(Frm115.CBB4, "  |  ")(1)
                    rs1!no_pekerja = Frm115_LM_EMP_NO 'No. Pekerja
                End If
                'If Frm115.L46_Text <> vbNullString Then
                '    If Frm28.L5_Text <> vbNullString Then
                '        rs1!no_rujukan_pembeli = Frm28.L5_Text 'No. Rujukan Pembeli
                '    Else
                '        rs1!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
                '    End If
                'Else
                '    rs1!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
                'End If
                'If Frm27.L5_Text <> vbNullString Then
                '    rs1!no_rujukan_agen_dropship = Frm27.L5_Text 'No. Rujukan Agen Dropship
                'Else
                '    rs1!no_rujukan_agen_dropship = Null 'No. Rujukan Agen Dropship
                'End If
                
    '1:  Pelanggan
    '2:  Member
    '3:  RAF
    '4:  Pengedar
    '5:  Normal Dealer
    '6:  Master Dealer
    
                If Frm115_LM_KATEGORI <> vbNullString Then
                    rs1!kategori_pembeli = Frm115_LM_KATEGORI
                Else
                    rs1!kategori_pembeli = Null
                End If
                
                If Not IsNull(rs!gst_include) Then 'Pilihan Cukai GST (SR) Samada Pelanggan Bayar Atau Kedai Bayar
                    If rs!gst_include = 0 Then
                        rs1!gst_include = Null
                    ElseIf rs!gst_include = 1 Then
                        rs1!gst_include = "**Harga Termasuk GST"
                    End If
                Else
                    rs1!gst_include = Null
                End If
                If Not IsNull(rs!harga_tanpa_gst) Then
                    rs1!harga_tanpa_gst = Format(rs!harga_tanpa_gst, "0.00") 'Harga Semasa (RM/g)
                Else
                    rs1!harga_tanpa_gst = Null 'Harga Semasa (RM/g)
                End If

'### Maklumat tetapan harga jualan kepada staff ### - Start
                If Not IsNull(rs!kadar_penurunan_upah) Then 'Kadar peratusan penurunan harga upah kepada staff (%)
                    rs1!kadar_penurunan_upah = Format(rs!kadar_penurunan_upah, "0.00")
                Else
                    rs1!kadar_penurunan_upah = Null
                End If
                If Not IsNull(rs!harga_semasa_staff) Then 'Harga emas semasa yang dijual kepada staff
                    rs1!harga_semasa_staff = Format(rs!harga_semasa_staff, "0.00")
                Else
                    rs1!harga_semasa_staff = Null
                End If
                If Not IsNull(rs!kadar_penurunan_bp) Then 'Kadar peratusan penurunan harga barang permata kepada staff (%)
                    rs1!kadar_penurunan_bp = Format(rs!kadar_penurunan_bp, "0.00")
                Else
                    rs1!kadar_penurunan_bp = Null
                End If
                If Not IsNull(rs!harga_staff) Then 'Harga yang dijual kepada staff (RM)
                    rs1!harga_staff = Format(rs!harga_staff, "0.00")
                Else
                    rs1!harga_staff = Null
                End If
                If Not IsNull(rs!harga_bp_asal) Then 'Tetapan harga barang permata yang asal (RM)
                    rs1!harga_bp_asal = Format(rs!harga_bp_asal, "0.00")
                Else
                    rs1!harga_bp_asal = Null
                End If
                If Not IsNull(rs!upah_asal) Then 'Tetapan upah asal (RM)
                    rs1!upah_asal = Format(rs!upah_asal, "0.00")
                Else
                    rs1!upah_asal = Null
                End If
                If Not IsNull(rs!komisyen_staff) Then 'Tetapan upah asal (RM)
                    rs1!komisyen_staff = Format(rs!komisyen_staff, "0.00")
                Else
                    rs1!komisyen_staff = Null
                End If
'### Maklumat tetapan harga jualan kepada staff ### - End

                If Not IsNull(rs!pemalar_tukaran_999) Then 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
                    rs1!pemalar_tukaran_999 = rs!pemalar_tukaran_999
                Else
                    rs1!pemalar_tukaran_999 = Null
                End If
                If Not IsNull(rs!berat_999) Then 'Berat jualan dalam purity 999.9
                    rs1!berat_999 = Format(rs!berat_999, "0.00")
                Else
                    rs1!berat_999 = Null
                End If
                rs1!write_timestamp = LM_NOW
                rs1!jenis_jualan = 1 '0 : Jualan biasa kepada pelanggan , 1 : Jualan secara tukaran barang kepada agen
                If Not IsNull(rs!gst_barang_atau_upah) Then '0 : GST pada harga jualan , 1 : GST pada upah
                    rs1!gst_barang_atau_upah = rs!gst_barang_atau_upah
                Else
                    rs1!gst_barang_atau_upah = 0
                End If
                If Not IsNull(rs!harga_jualan_dengan_gst) Then
                    rs1!harga_jualan_dengan_gst = rs!harga_jualan_dengan_gst
                Else
                    rs1!harga_jualan_dengan_gst = 0
                End If
                rs1!status_rekod = 1
                rs1.Update
                
                rs1.Close
                Set rs1 = Nothing
            
'### Update Table Database Bagi Item Ini ### - Start
                Set rs2 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs2.Open "select * from Data_Database where no_siri_produk='" & rs!no_siri_Produk & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs2.EOF Then
                
                    G_ID = rs2!ID
                    Call recovery_data_database
                    
                    If rs!Type = 0 Then
                        Frm115_LM_BERAT_ASAL = 0
                        Frm115_LM_BERAT_JUALAN = 0
                        
                        Frm115_LM_BERAT_ASAL = rs2!beza_berat 'Berat Asal (g)
                        Frm115_LM_BERAT_JUALAN = rs!berat_jualan 'Berat Jualan (g)
                        
                        If Format(Frm115_LM_BERAT_JUALAN, "0.00") = Format(Frm115_LM_BERAT_ASAL, "0.00") Then
                            rs2!beza_berat = "0.00" 'Baki Berat
                            rs2!StatusItem = 27
                            rs2!tarikh_jualan1 = Null
                        Else
                            rs2!beza_berat = Format(Frm115_LM_BERAT_ASAL - Frm115_LM_BERAT_JUALAN, "0.00") 'Baki Berat
                            rs2!StatusItem = 28
                            rs2!tarikh_jualan1 = Frm115.DTPicker1
                        End If
                    Else
                        rs2!StatusItem = 27
                    End If

                    rs2!write_timestamp2 = LM_NOW
                    rs2!no_pekerja = Frm115_LM_EMP_NO
                    rs2!terminal = G_TERMINAL
                    rs2!Menu = 6

                    rs2.Update
                End If

                rs2.Close
                Set rs2 = Nothing
'### Update Table Database Bagi Item Ini ### - End
'########### Kemasukan data baru dalam senarai ##############- End
            
'########### Edit data sedia ada dalam senarai ##############- Start
            ElseIf rs!Status = 3 Then
            
                Set rs1 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs1.Open "select * from 23_senarai_jualan where ID='" & rs!id_database & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs1.EOF Then
                
                    G_ID = rs1!ID
                    Call recovery_23_senarai_jualan
                    
                    rs1!tarikh = Frm115.DTPicker1 'Tarikh Jualan
                    rs1!no_resit = Frm115.L23_Text 'No. Invoice Jualan
                    If Not IsNull(rs!no_siri_Produk) Then
                        rs1!no_siri_Produk = rs!no_siri_Produk 'No. Siri Produk
                    Else
                        rs1!no_siri_Produk = Null 'No. Siri Produk
                    End If
                    If Not IsNull(rs!kategori_Produk) Then
                        rs1!kategori_Produk = rs!kategori_Produk 'Kategori Produk
                    Else
                        rs1!no_siri_Produk = Null 'Kategori Produk
                    End If
                    If Not IsNull(rs!purity) Then
                        rs1!purity = rs!purity 'Purity
                    Else
                        rs1!purity = Null 'Purity
                    End If
                    If Not IsNull(rs!Berat_Asal) Then
                        rs1!Berat_Asal = rs!Berat_Asal 'Berat Asal (g)
                    Else
                        rs1!Berat_Asal = Null 'Berat Asal (g)
                    End If
                    If Not IsNull(rs!berat_jualan) Then
                        rs1!berat_jualan = Format(rs!berat_jualan, "0.00") 'Berat Jualan (g)
                    Else
                        rs1!berat_jualan = Null 'Berat Jualan (g)
                    End If
                    If Not IsNull(rs!harga_Semasa) Then
                        rs1!harga_Semasa = Format(rs!harga_Semasa, "0.00") 'Harga Semasa (RM/g)
                    Else
                        rs1!harga_Semasa = Null 'Harga Semasa (RM/g)
                    End If
                    If Not IsNull(rs!UPAH) Then
                        rs1!UPAH = Format(rs!UPAH, "0.00") 'Upah (RM)
                    Else
                        rs1!UPAH = Null 'Upah (RM)
                    End If
                    If Not IsNull(rs!harga_asal) Then
                        rs1!harga_asal = Format(rs!harga_asal, "0.00") 'Harga Asal Item (RM)
                    Else
                        rs1!harga_asal = Null 'Harga Asal Item (RM)
                    End If
                    If Not IsNull(rs!diskaun) Then
                        rs1!diskaun = Format(rs!diskaun, "0.00") 'Diskaun (%)
                    Else
                        rs1!diskaun = Null 'Diskaun (%)
                    End If
                    If Not IsNull(rs!harga_lepas_diskaun) Then
                        rs1!harga_lepas_diskaun = Format(rs!harga_lepas_diskaun, "0.00") 'Harga Selepas Diskaun (RM)
                    Else
                        rs1!harga_lepas_diskaun = Null 'Harga Selepas Diskaun (RM)
                    End If
                    If Not IsNull(rs!adjustment) Then
                        rs1!adjustment = Format(rs!adjustment, "0.00") 'Harga Selepas Diskaun (RM)
                    Else
                        rs1!adjustment = Null 'Harga Selepas Diskaun (RM)
                    End If
                    If Not IsNull(rs!harga_jualan) Then
                        rs1!harga_jualan = Format(rs!harga_jualan, "0.00") 'Harga Jualan (RM)
                    Else
                        rs1!harga_jualan = Null 'Harga Jualan (RM)
                    End If
                    If Not IsNull(rs!gst_ari_nashi) Then
                        rs1!gst_ari_nashi = rs!gst_ari_nashi '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                    Else
                        rs1!gst_ari_nashi = Null '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                    End If
                    If Not IsNull(rs!kadar_gst) Then
                        rs1!kadar_gst = rs!kadar_gst 'Kadar Cukai GST (%)
                    Else
                        rs1!kadar_gst = Null 'Kadar Cukai GST (%)
                    End If
                    If Not IsNull(rs!jumlah_gst) Then
                        rs1!jumlah_gst = Format(rs!jumlah_gst, "0.00") 'Jumlah Cukai GST (RM)
                    Else
                        rs1!jumlah_gst = Null 'Jumlah Cukai GST (RM)
                    End If
                    If Not IsNull(rs!harga_dengan_gst) Then
                        rs1!harga_dengan_gst = Format(rs!harga_dengan_gst, "0.00") 'Harga Jualan Termasuk GST (RM)
                    Else
                        rs1!harga_dengan_gst = Null 'Harga Jualan Termasuk GST (RM)
                    End If
                    If Not IsNull(rs!dropship) Then
                        rs1!dropship = rs!dropship '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
                    Else
                        rs1!dropship = Null '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
                    End If
                    If Not IsNull(rs!komisyen_per_gram) Then
                        rs1!komisyen_per_gram = Format(rs!komisyen_per_gram, "0.00") 'Komisyen Per Gram Dropship (RM/g) : BK
                    Else
                        rs1!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g) : BK
                    End If
                    If Not IsNull(rs!jumlah_komisyen) Then
                        rs1!jumlah_komisyen = Format(rs!jumlah_komisyen, "0.00") 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
                    Else
                        rs1!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
                    End If
                    If Not IsNull(rs!harga_per_gram_modal) Then
                        rs1!harga_per_gram_modal = Format(rs!harga_per_gram_modal, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
                    Else
                        rs1!harga_per_gram_modal = Null 'Harga Per Gram Bagi Modal (RM/g)
                    End If
                    If Not IsNull(rs!modal) Then
                        rs1!modal = Format(rs!modal, "0.00") 'Harga Modal (RM)
                    Else
                        rs1!modal = Null 'Harga Modal (RM)
                    End If
                    If Not IsNull(rs!untung) Then
                        rs1!untung = Format(rs!untung, "0.00") 'Jumlah Keuntungan
                    Else
                        rs1!untung = Null 'Jumlah Keuntungan
                    End If
                    If Not IsNull(rs!harga_per_gram_supplier) Then
                        rs1!harga_per_gram_supplier = Format(rs!harga_per_gram_supplier, "0.00") 'Harga per gram (harga semasa) dari supplier (modal)
                    Else
                        rs1!harga_per_gram_supplier = Null 'Harga per gram (harga semasa) dari supplier (modal)
                    End If
                    If Not IsNull(rs!upah_modal) Then
                        rs1!upah_modal = Format(rs!upah_modal, "0.00") 'Upah modal
                    Else
                        rs1!upah_modal = Null 'Upah modal
                    End If
                    If Not IsNull(rs!untung2) Then
                        rs1!untung2 = Format(rs!untung2, "0.00") 'Jumlah Keuntungan
                    Else
                        rs1!untung2 = Null 'Jumlah Keuntungan
                    End If
                    If Not IsNull(rs!dulang) Then
                        rs1!dulang = rs!dulang 'Dulang
                    Else
                        rs1!dulang = Null 'Dulang
                    End If
                    If Not IsNull(rs!potong_flag) Then
                        rs1!potong_flag = rs!potong_flag '0 : Tiada Potong , 1 : Ada Potong
                        If rs!potong_flag = 0 Then
                            rs1!Status = 0 '0 : Jualan Biasa , 1 : Jualan Secara Potong , 2 : Tempahan , 3 : Ansuran , 4 : ETA
                        Else
                            rs1!Status = 1 '0 : Jualan Biasa , 1 : Jualan Secara Potong , 2 : Tempahan , 3 : Ansuran , 4 : ETA
                        End If
                    Else
                        rs1!potong_flag = Null '0 : Tiada Potong , 1 : Ada Potong
                    End If
                    If Not IsNull(rs!Type) Then
                        rs1!Type = rs!Type '0 : BK , 1 : Barang Permata
                    Else
                        rs1!Type = Null '0 : BK , 1 : Barang Permata
                    End If
                    If Frm115.CBB4 <> vbNullString Then
                        Frm115_LM_EMP_NO = Split(Frm115.CBB4, "  |  ")(1)
                        rs1!no_pekerja = Frm115_LM_EMP_NO 'No. Pekerja
                    End If
                    rs1!jualan_online = 0
                    'If Frm115.L46_Text <> vbNullString Then
                    '    If Frm28.L5_Text <> vbNullString Then
                    '        rs1!no_rujukan_pembeli = Frm28.L5_Text 'No. Rujukan Pembeli
                    '    Else
                    '        rs1!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
                    '    End If
                    'Else
                    '    rs1!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
                    'End If
        '1:  Pelanggan
        '2:  Member
        '3:  RAF
        '4:  Pengedar
        '5:  Normal Dealer
        '6:  Master Dealer
        
                    If Frm115_LM_KATEGORI <> vbNullString Then
                        rs1!kategori_pembeli = Frm115_LM_KATEGORI
                    Else
                        rs1!kategori_pembeli = Null
                    End If

                    If Not IsNull(rs!gst_include) Then 'Pilihan Cukai GST (SR) Samada Pelanggan Bayar Atau Kedai Bayar
                        If rs!gst_include = 0 Then
                            rs1!gst_include = Null
                        ElseIf rs!gst_include = 1 Then
                            rs1!gst_include = "**Harga Termasuk GST"
                        End If
                    Else
                        rs1!gst_include = Null
                    End If
                    If Not IsNull(rs!harga_tanpa_gst) Then
                        rs1!harga_tanpa_gst = Format(rs!harga_tanpa_gst, "0.00") 'Harga Semasa (RM/g)
                    Else
                        rs1!harga_tanpa_gst = Null 'Harga Semasa (RM/g)
                    End If
    
    '### Maklumat tetapan harga jualan kepada staff ### - Start
                    If Not IsNull(rs!kadar_penurunan_upah) Then 'Kadar peratusan penurunan harga upah kepada staff (%)
                        rs1!kadar_penurunan_upah = Format(rs!kadar_penurunan_upah, "0.00")
                    Else
                        rs1!kadar_penurunan_upah = Null
                    End If
                    If Not IsNull(rs!harga_semasa_staff) Then 'Harga emas semasa yang dijual kepada staff
                        rs1!harga_semasa_staff = Format(rs!harga_semasa_staff, "0.00")
                    Else
                        rs1!harga_semasa_staff = Null
                    End If
                    If Not IsNull(rs!kadar_penurunan_bp) Then 'Kadar peratusan penurunan harga barang permata kepada staff (%)
                        rs1!kadar_penurunan_bp = Format(rs!kadar_penurunan_bp, "0.00")
                    Else
                        rs1!kadar_penurunan_bp = Null
                    End If
                    If Not IsNull(rs!harga_staff) Then 'Harga yang dijual kepada staff (RM)
                        rs1!harga_staff = Format(rs!harga_staff, "0.00")
                    Else
                        rs1!harga_staff = Null
                    End If
                    If Not IsNull(rs!harga_bp_asal) Then 'Tetapan harga barang permata yang asal (RM)
                        rs1!harga_bp_asal = Format(rs!harga_bp_asal, "0.00")
                    Else
                        rs1!harga_bp_asal = Null
                    End If
                    If Not IsNull(rs!upah_asal) Then 'Tetapan upah asal (RM)
                        rs1!upah_asal = Format(rs!upah_asal, "0.00")
                    Else
                        rs1!upah_asal = Null
                    End If
                    If Not IsNull(rs!komisyen_staff) Then 'Tetapan upah asal (RM)
                        rs1!komisyen_staff = Format(rs!komisyen_staff, "0.00")
                    Else
                        rs1!komisyen_staff = Null
                    End If
    '### Maklumat tetapan harga jualan kepada staff ### - End
    
                    If Not IsNull(rs!pemalar_tukaran_999) Then 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
                        rs1!pemalar_tukaran_999 = rs!pemalar_tukaran_999
                    Else
                        rs1!pemalar_tukaran_999 = Null
                    End If
                    If Not IsNull(rs!berat_999) Then 'Berat jualan dalam purity 999.9
                        rs1!berat_999 = Format(rs!berat_999, "0.00")
                    Else
                        rs1!berat_999 = Null
                    End If
                    If Not IsNull(rs!gst_barang_atau_upah) Then '0 : GST pada harga jualan , 1 : GST pada upah
                        rs1!gst_barang_atau_upah = rs!gst_barang_atau_upah
                    Else
                        rs1!gst_barang_atau_upah = 0
                    End If
                    If Not IsNull(rs!harga_jualan_dengan_gst) Then
                        rs1!harga_jualan_dengan_gst = rs!harga_jualan_dengan_gst
                    Else
                        rs1!harga_jualan_dengan_gst = 0
                    End If
                    rs1!write_timestamp2 = LM_NOW
                    rs1!no_staff = G_LOGIN_USER
                    rs1!status_rekod = 1
                    
                    rs1.Update
                End If

                rs1.Close
                Set rs1 = Nothing

                Set rs2 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs2.Open "select * from Data_Database where no_siri_produk='" & rs!no_siri_Produk & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs2.EOF Then

                    G_ID = rs2!ID
                    Call recovery_data_database
            
                    If rs!Type = 0 Then
                        If Not IsNull(rs2!Berat) Then Frm115_LM_BERAT_ASAL = Format(rs2!Berat, "0.00") 'Berat Asal (g)
                        If Not IsNull(rs2!beza_berat) Then Frm115_LM_BEZA_BERAT = Format(rs2!beza_berat, "0.00") 'Berat Asal (g)
                        If Not IsNull(rs!berat_jualan) Then Frm115_BERAT_JUALAN_BARU = Format(rs!berat_jualan, "0.00") 'Berat Jualan (g)
                        If Not IsNull(rs2!susut_berat) Then Frm115_SUSUT_BERAT = Format(rs2!susut_berat, "0.00") 'Susut berat
                        
                        Frm115_LM_BAKI_BERAT = Frm115_BERAT_JUALAN_BARU - Format((Frm115_LM_BERAT_JUALAN_ASAL + Frm115_LM_BEZA_BERAT), "0.00") - Frm115_SUSUT_BERAT
                        
                        If Format(Frm115_BERAT_JUALAN_BARU, "0.00") = Format(Frm115_LM_BERAT_RETURN + Frm115_LM_BEZA_BERAT + Frm115_SUSUT_BERAT, "0.00") Then
                            rs2!beza_berat = "0.00" 'Baki Berat
                            rs2!StatusItem = 27
                            rs2!tarikh_jualan1 = Null
                        Else
                            rs2!beza_berat = Format(Frm115_LM_BERAT_ASAL - Frm115_BERAT_JUALAN_BARU - Frm115_SUSUT_BERAT, "0.00") 'Baki Berat
                            rs2!StatusItem = 28
                            rs2!tarikh_jualan1 = Frm115.DTPicker1
                        End If
                    Else
                        rs2!StatusItem = 27
                    End If

                    rs2!write_timestamp2 = LM_NOW
                    rs2!no_pekerja = Frm115_LM_EMP_NO
                    rs2!terminal = G_TERMINAL
                    rs2!Menu = 6
                
                    rs2.Update
                    
                End If
                
                rs2.Close
                Set rs2 = Nothing

'########### Edit data sedia ada dalam senarai ##############- End
            
            ElseIf rs!Status = 5 Then

                Set rs1 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs1.Open "select * from 23_senarai_jualan where ID='" & rs!id_database & "'", cn, adOpenKeyset, adLockOptimistic
            
                If Not rs1.EOF Then
                
                    G_ID = rs1!ID
                    Call recovery_23_senarai_jualan
                    
                    If Not IsNull(rs1!berat_jualan) Then
                        Frm115_LM_BERAT_RETURN = rs1!berat_jualan
                    End If

                    rs1!no_staff = G_LOGIN_USER
                    rs1!write_timestamp2 = LM_NOW
                    rs1!status_rekod = 0
                    
                    rs1.Update
                End If
                
                rs1.Close
                Set rs1 = Nothing

'### Update Table Database Bagi Item Ini ### - Start
                Set rs2 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs2.Open "select * from Data_Database where no_siri_produk='" & rs!no_siri_Produk & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs2.EOF Then
                
                    G_ID = rs2!ID
                    Call recovery_data_database
                
                    If rs!Type = 0 Then
                        If Not IsNull(rs2!Berat) Then Frm115_LM_BERAT_ASAL = Format(rs2!Berat, "0.00") 'Berat Asal (g)
                        If Not IsNull(rs2!beza_berat) Then Frm115_LM_BEZA_BERAT = Format(rs2!beza_berat, "0.00") 'Berat Asal (g)
                        If Not IsNull(rs!berat_jualan) Then Frm115_BERAT_JUALAN_BARU = Format(rs!berat_jualan, "0.00") 'Berat Jualan (g)
                        If Not IsNull(rs2!susut_berat) Then Frm115_SUSUT_BERAT = Format(rs2!susut_berat, "0.00") 'Susut berat
                        
                        Frm115_LM_BAKI_BERAT = Frm115_BERAT_JUALAN_BARU - Format((Frm115_LM_BERAT_JUALAN_ASAL + Frm115_LM_BEZA_BERAT), "0.00") - Frm115_SUSUT_BERAT
                        
                        If Format(Frm115_LM_BERAT_ASAL, "0.00") = Format(Frm115_LM_BERAT_RETURN + Frm115_LM_BEZA_BERAT, "0.00") Then
                            rs2!beza_berat = Format(Frm115_LM_BERAT_RETURN + Frm115_LM_BEZA_BERAT, "0.00")  'Baki Berat
                            rs2!StatusItem = 10
                            rs2!tarikh_jualan1 = Null
                        Else
                            rs2!beza_berat = Format(Frm115_LM_BEZA_BERAT + Frm115_LM_BERAT_RETURN, "0.00")  'Baki Berat
                            rs2!StatusItem = 12
                            rs2!tarikh_jualan1 = Frm115.DTPicker1
                        End If
                                
                    Else
                        rs2!StatusItem = 10
                    End If
                    
                    rs2!write_timestamp2 = LM_NOW
                    rs2!no_pekerja = Frm115_LM_EMP_NO
                    rs2!terminal = G_TERMINAL
                    rs2!Menu = 6
                    
                    rs2.Update
                End If
                
                rs2.Close
                Set rs2 = Nothing
            
            End If

            
            rs.MoveNext
        Wend '2
        
        rs.Close
        Set rs = Nothing
    
        If DATA_SAVE = 1 Then
    '###Update No. Resit### - Start
            G_No_RESIT_JUALAN = Frm115.L23_Text
            
    '#### Update Log Aktiviti Sistem #### - Start
            If Frm115.CBB4 <> vbNullString Then
                Frm115_LM_EMP_NAME = Split(Frm115.CBB4, "  |  ")(0)
            End If
        
            'User = MDI_frm1.L3_Text
            LogAct_Memory = "[" & Frm115_LM_EMP_NAME & "] Edit data pengeluaran GDN kepada agen/supplier (Per Item). No. Rujukan [" & G_No_RESIT_JUALAN & "]."
            LogDate_Memory = LM_NOW
            Call UpdateLog_Database
    '#### Update Log Aktiviti Sistem #### - End
            
            If Frm115_LM_GENERATED = 1 Then '0 : Tiada No Voucher yang dihasilkan , 1 : Ada No. Voucher yang dihasilkan
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If rs!Default1 = "Default" Then

                            rs!no_trade_in_agen = Frm115.L22_Text + 1 'No. Voucher Trade In

                        rs.Update
                    End If
                End If
                
                rs.Close
                Set rs = Nothing
            End If
            
            Note = "Data Telah Berjaya Disimpan." & vbCrLf & _
                    "Refresh Data Anda ?"

            Answer = MsgBox(Note, vbQuestion + vbOK, "Confirmation")
            
            If Answer = vbOK Then
            
                GM_NEXT_PREV = 2
                
                If Frm115.L71_Text = "0" Then
                
                    If Frm101.L33_Text = 0 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                        If Frm101.CB3 = 1 Then 'Report Jualan
                            Call Frm85_Header_Report_Jualan
                            Call Frm85_Report_Jualan_page
                        End If
                    ElseIf Frm101.L33_Text = 2 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                        Call Frm85_Header_Report_Jualan
                        Call Frm85_carian_jualan_page
                    ElseIf Frm101.L33_Text = 5 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                        Call Frm85_Header_Report_Jualan
                        Call Frm85_Report_Jualan_barcode
                    End If
                    
                    Frm85.Show
                    Unload Frm115
                    MDI_frm1.L5_Text = 12
                    
                ElseIf Frm115.L71_Text = "1" Then
                
                    Call frm117_report_gdn_grn_header
                    Call frm117_report_gdn_grn

                    frm117.Show
                    Unload Frm115
                    
                End If
                
            Else
                
                If Frm115.L71_Text = "0" Then
                
                    Frm85.Show
                    Unload Frm115
                    MDI_frm1.L5_Text = 12
                    
                ElseIf Frm115.L71_Text = "1" Then
                
                    frm117.Show
                    Unload Frm115
                
                End If
                
            End If
            
            MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
    '###Update No. Resit### - End
        End If
        
    End If
End If
End Sub
Private Sub CMD11_Click()
'On Error Resume Next
Note = "Adakah anda ingin batalkan edit data ini?" & vbCrLf & _
        "Sistem tidak akan menyimpan data jika terdapat data yang diubah." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then

    If Frm115.L71_Text = "0" Then
    
        Frm85.Show
        Unload Frm115
        MDI_frm1.L5_Text = 12
        
    ElseIf Frm115.L71_Text = "1" Then
    
        frm117.Show
        Unload Frm115
        
    End If
    
End If
End Sub

Private Sub CMD12_Click()
'On Error Resume Next
If Frm115.CBB5 = vbNullString Then

    MsgBox "Sila buat pilihan kategori produk.", vbExclamation, "Info"
    
    Exit Sub
End If

If Frm115.CBB6 = vbNullString Then

    MsgBox "Sila buat pilihan purity.", vbExclamation, "Info"
    
    Exit Sub
End If

Frm115.L55_Text = Frm115.CBB5 'Kategori Produk
Frm115.L56_Text = Frm115.CBB6 'Purity

Call frm115_reset_gdn_list
End Sub

Private Sub CMD13_Click()
'On Error Resume Next
Frm115.Frame4.Visible = False
End Sub

Private Sub CMD14_Click()
'on error resume next
Dim frm115_LM_CURR_PAGE As Double
Dim frm115_LM_TOTAL_PAGE As Double

frm115_LM_CURR_PAGE = 0
frm115_LM_TOTAL_PAGE = 0

If Frm115.L61_Text <> vbNullString And IsNumeric(Frm115.L61_Text) Then
    If Frm115.L62_Text <> vbNullString And IsNumeric(Frm115.L62_Text) Then
        frm115_LM_CURR_PAGE = Frm115.L61_Text
        frm115_LM_TOTAL_PAGE = Frm115.L62_Text
        
        If frm115_LM_CURR_PAGE < frm115_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call frm115_gdn_list_header
            Call frm115_gdn_list
            
        End If
    End If
End If
End Sub

Private Sub CMD2_Click()
'On Error Resume Next
Dim Err(30)
Dim Frm115_LM_BERAT_ASAL As Double
Dim Frm115_LM_BERAT_JUAL As Double
Dim Frm115_LM_HARGA_MODAL As Double
Dim Frm115_LM_HARGA_JUAL As Double
Dim Frm115_LM_HARGA_SEMASA_MODAL As Double
Dim Frm115_LM_TETAPANHARGA As Double
Dim Frm115_LM_LIMIT As Double
Dim Frm115_LM_HARGA_STAFF As Double 'Tetapan harga jualan kepada staff
Dim Frm115_LM_HARGA_PELANGGAN As Double 'Tetapan harga jualan kepada pelanggan
Dim Frm115_LM_BERAT_JUAL_ASAL As Double 'Berat Jualan (Purity Asal)
Dim Frm115_LM_HARGA_SEMASA_999 As Double 'Harga semasa (jualan) (Purity 999.9)
Dim Frm115_LM_HARGA_SUPPLIER As Double 'Harga per gram (harga semasa) dari supplier (modal)
Dim Frm115_LM_BERAT_999 As Double 'Berat Jualan (Purity Asal)
Dim Frm115_UPAH_MODAL As Double 'Upah modal
Dim Frm115_UPAH_JUAL As Double 'Upah jualan
Dim LM_KADAR_TUKARAN As Double

LM_KADAR_TUKARAN = 0
Frm115_UPAH_MODAL = 0 'Upah modal
Frm115_UPAH_JUAL = 0 'Upah jualan
Frm115_LM_BERAT_JUAL_ASAL = 0 'Berat Jualan (Purity Asal)
Frm115_LM_HARGA_SEMASA_999 = 0 'Harga semasa (jualan) (Purity 999.9)
Frm115_LM_HARGA_SUPPLIER = 0 'Harga per gram (harga semasa) dari supplier (modal)
Frm115_LM_BERAT_999 = 0 'Berat Jualan (Purity Asal)
x = 0
Frm115_LM_BERAT_ASAL = 0
Frm115_LM_BERAT_JUAL = 0
Frm115_LM_DATA_SAVE = 0
Frm115_LM_HARGA_MODAL = 0
Frm115_LM_HARGA_JUAL = 0
Frm115_LM_HARGA_SEMASA_MODAL = 0
Frm115_LM_PRICE_CHECK = 0 '0 : Tidak Perlu Periksa Harga Semasa Jualan , 1 : Perlu Periksa Harga Semasa Jualan
Frm115_LM_TETAPANHARGA = 0
Frm115_LM_LIMIT = 0
Frm115_LM_HARGA_STAFF = 0
Frm115_LM_HARGA_PELANGGAN = 0

If Frm115.L3_Text = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan [No. Siri Produk]."
End If
If Frm115.L33_Text = vbNullString Then
    x = x + 1
    Err(x) = "Maklumat harga semasa modal belian item ini yang tidak sah. Sila keluar dari menu ini dan cuba sekali lagi."
End If
If Frm115.L50_Text = vbNullString Or (Frm115.L50_Text <> vbNullString And Not IsNumeric(Frm115.L50_Text)) Then
    x = x + 1
    Err(x) = "Maklumat upah modal yang tidak sah. Sila keluar dari menu ini dan cuba sekali lagi."
End If
If Frm115.L6_Text = vbNullString Or (Frm115.L6_Text <> vbNullString And Not IsNumeric(Frm115.L6_Text)) Then
    x = x + 1
    Err(x) = "Sila maklumat [Berat Asal]. Sila scan item sekali lagi."
End If
If Frm115.TB3 = vbNullString Or (Frm115.TB3 <> vbNullString And Not IsNumeric(Frm115.TB3)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Berat Jualan]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm115.TB2 = vbNullString Or (Frm115.TB2 <> vbNullString And Not IsNumeric(Frm115.TB2)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Harga Semasa Emas 999.9]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm115.TB2 <> vbNullString And IsNumeric(Frm115.TB2) Then

    If Format(Frm115.TB2, "0.00") = "0.00" Then
        x = x + 1
        Err(x) = "Harga emas semasa 999.9 yang tidak sah. Nilai 0.00 tidak dibenarkan."
    End If
    
End If
If Frm115.TB7 = vbNullString Or (Frm115.TB7 <> vbNullString And Not IsNumeric(Frm115.TB7)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar Tukaran Purity 999.9]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If (Frm115.TB7 <> vbNullString And IsNumeric(Frm115.TB7)) Then
    
    LM_KADAR_TUKARAN = Frm115.TB7
    
    If LM_KADAR_TUKARAN > 1 Then
    
        x = x + 1
        Err(x) = "[Kadar Tukaran Purity 999.9] tidak boleh lebih dari 1.00. Hanya NOMBOR dibenarkan dalam ruangan ini."
        
    End If
    
End If
If Frm115.L7_Text = vbNullString Or (Frm115.L7_Text <> vbNullString And Not IsNumeric(Frm115.L7_Text)) Then
    x = x + 1
    Err(x) = "[Berat 999.9] yang tidak sah. Sila scan item sekali lagi."
End If
If Frm115.TB4 = vbNullString Or (Frm115.TB4 <> vbNullString And Not IsNumeric(Frm115.TB4)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Upah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm115.CB2 = 0 And Frm115.CB3 = 0 And Frm115.CB4 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan jenis GST bagi upah."
End If
If Frm115.TB5 = vbNullString Or Frm115.TB6 = vbNullString Then
    x = x + 1
    Err(x) = "Maklumat berkenaan GST yang tidak sah. Sila keluar dari menu ini dan cuba sekali lagi."
End If

If (Frm115.L6_Text <> vbNullString And IsNumeric(Frm115.L6_Text)) And (Frm115.TB3 <> vbNullString And IsNumeric(Frm115.TB3)) Then
    Frm115_LM_BERAT_ASAL = Frm115.L6_Text 'Berat Asal
    Frm115_LM_BERAT_JUAL = Frm115.TB3 'Berat Jualan
    
    If Frm115_LM_BERAT_JUAL > Frm115_LM_BERAT_ASAL Then
        x = x + 1
        Err(x) = "Berat jualan melebihi berat asal."
    End If
End If
If Frm115.L49_Text = vbNullString Or (Frm115.L49_Text <> vbNullString And Not IsNumeric(Frm115.L49_Text)) Then
    x = x + 1
    Err(x) = "Tiada maklumat harga dari supplier bagi purity ini." & vbCrLf & _
                "Sila pastikan harga dari supplier bagi purity ini telah ditetapkan dalam TETAPAN HARIAN SISTEM."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah anda ingin masukkan item ini ke dalam senarai jualan ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
'### Periksa Data Dulang ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where no_siri_Produk='" & Frm115.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!dulang) Then Frm115_LM_DULANG = rs!dulang 'Dulang
        End If
        
        rs.Close
        Set rs = Nothing
'### Periksa Data Dulang ### - End
        
'### Masukkan Data Ke Dalam Temp Table ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_GDN_TEMP & " where ID='" & Frm115.L24_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Frm115.L3_Text <> vbNullString Then
                rs!no_siri_Produk = Frm115.L3_Text 'No. Siri Produk
            Else
                rs!no_siri_Produk = Null 'No. Siri Produk
            End If
            If Frm115.L5_Text <> vbNullString Then
                rs!kategori_Produk = Frm115.L5_Text 'Kategori Produk
            Else
                rs!kategori_Produk = Null 'Kategori Produk
            End If
            If Frm115.L4_Text <> vbNullString Then
                rs!purity = Frm115.L4_Text 'Purity
            Else
                rs!purity = Null 'Purity
            End If
            If Frm115.L6_Text <> vbNullString Then
                rs!Berat_Asal = Format(Frm115.L6_Text, "0.00") 'Berat Asal (g)
            Else
                rs!Berat_Asal = Null 'Berat Asal (g)
            End If
            If Frm115.TB3 <> vbNullString Then
                rs!berat_jualan = Format(Frm115.TB3, "0.00") 'Berat Jualan (g)
            Else
                rs!berat_jualan = Null 'Berat Jualan (g)
            End If
            If Frm115.TB2 <> vbNullString Then
                rs!harga_Semasa = Format(Frm115.TB2, "0.00") 'Harga Semasa (RM/g)
            Else
                rs!harga_Semasa = Null 'Harga Semasa (RM/g)
            End If
            If Frm115.TB4 <> vbNullString Then
                rs!UPAH = Format(Frm115.TB4, "0.00") 'Upah (RM)
            Else
                rs!UPAH = Null 'Upah (RM)
            End If
            
            Frm115_LM_HARGA_SEMASA = Frm115.TB2 'Harga emas semasa 999.9 (Untuk tujuan jualan kepada pelanggan)
            Frm115_LM_BERAT_JUALAN_9999 = Frm115.L7_Text 'Berat jualan dalam purity 999.9
            Frm115_LM_UPAH_DAN_GST = Frm115.TB6 'Jumlah Upah + GST (Bagi jualan setiap item)

            If Frm115.TB6 <> vbNullString Then
                rs!harga_asal = Format((Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            Else
                rs!harga_asal = Null 'Harga Asal Item (RM)
            End If
            
            rs!diskaun = "0.00" 'Diskaun (%)
            rs!harga_lepas_diskaun = Format((Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            rs!adjustment = Format(0, "0.00") 'Adjustment (RM)
            rs!harga_jualan = Format((Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            rs!harga_jualan_dengan_gst = Format((Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
            
            If Frm115.CB2 = 1 Then
                rs!gst_ari_nashi = "ZR (L)" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                rs!kadar_gst = Null 'Kadar Cukai GST (%)
                If Frm115.TB5 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm115.TB5, "0.00") 'Jumlah GST (Bagi jualan setiap item)
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah GST (Bagi jualan setiap item)
                End If
            ElseIf Frm115.CB3 = 1 Then
                rs!gst_ari_nashi = "SR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                If Frm115.L21_Text <> vbNullString Then
                    rs!kadar_gst = Frm115.L21_Text 'Kadar Cukai GST (%)
                Else
                    rs!kadar_gst = "0" 'Jumlah Cukai GST (RM)
                End If
                If Frm115.TB5 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm115.TB5, "0.00") 'Jumlah GST (Bagi jualan setiap item)
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah GST (Bagi jualan setiap item)
                End If
                If Frm115.CB4 = 1 Then 'Jenis Cukai GST SR
                    rs!gst_include = 1 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                Else
                    rs!gst_include = 0 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                End If
            End If
            If Frm115.L30_Text <> vbNullString Then
                rs!harga_tanpa_gst = Format(Frm115.L30_Text, "0.00") 'Harga Jualan Tanpa GST (RM)
            Else
                rs!harga_tanpa_gst = Null 'Harga Jualan Tanpa GST (RM)
            End If
            If Frm115.TB6 <> vbNullString Then
                rs!harga_dengan_gst = Format(Frm115.TB6, "0.00") 'Harga Jualan Termasuk GST (RM)
            Else
                rs!harga_dengan_gst = Null 'Harga Jualan Termasuk GST (RM)
            End If
            rs!dropship = 0 '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
            rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g)
            rs!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini
            
'Status
'0 : Keluarkan Dari Senarai
'1 : Data Baru (Fresh)
'2 : Data Baru Diedit (Fresh)
'3 : Data Baru Dari Menu Edit
'4 : Data Baru Dari Menu Edit Yang Telah Diedit

            If Frm115.L32_Text = "0" Then '0 : Menu Data Baru , 1 : Menu Edit Data
                rs!Status = 1
            ElseIf Frm115.L32_Text = "1" Then
                If rs!Status = "2" Then
                    rs!Status = 3
                End If
                If rs!Status = "4" Then
                    rs!Status = 4
                End If
            End If
            
            rs!Type = 0 '0 : BK , 1 : Barang Permata
            If Frm115.L33_Text <> vbNullString Then
                rs!harga_per_gram_modal = Format(Frm115.L33_Text, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
                Frm115_LM_HARGA_SEMASA_MODAL = Frm115.L33_Text
            Else
                rs!harga_per_gram_modal = Format(0, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
            End If
            rs!modal = Format(Frm115_LM_HARGA_SEMASA_MODAL * Frm115_LM_BERAT_JUALAN_9999, "0.00") 'Harga Modal (RM)
            If IsNumeric(Frm115.TB6) And IsNumeric(Frm115.L33_Text) And IsNumeric(Frm115.TB3) Then
                Frm115_LM_HARGA_MODAL = Frm115.L33_Text * Frm115.TB3 'Harga modal
                Frm115_LM_HARGA_JUAL = (Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST 'Harga jualan
                
                rs!untung = Format(Frm115_LM_HARGA_JUAL - Frm115_LM_HARGA_MODAL, "0.00") 'Jumlah Keuntungan
            Else
                rs!untung = Format(0, "0.00") 'Jumlah Keuntungan
            End If

            If Frm115.L49_Text <> vbNullString Then 'Harga per gram (harga semasa) dari supplier (modal)
                rs!harga_per_gram_supplier = Frm115.L49_Text
            Else
                rs!harga_per_gram_supplier = 0
            End If
            
            If Frm115.L49_Text <> vbNullString Then 'Harga per gram (harga semasa) dari supplier (modal)
                rs!harga_per_gram_supplier = Frm115.L49_Text
            Else
                rs!harga_per_gram_supplier = 0
            End If
            
            If IsNumeric(Frm115.TB3) And IsNumeric(Frm115.TB2) And IsNumeric(Frm115.L49_Text) And IsNumeric(Frm115.L7_Text) And IsNumeric(Frm115.L6_Text) And IsNumeric(Frm115.L50_Text) And IsNumeric(Frm115.TB4) Then
                Frm115_LM_BERAT_JUAL_ASAL = Frm115.TB3 'Berat Jualan (Purity Asal)
                Frm115_LM_BERAT_ASAL = Frm115.L6_Text 'Berat Asal (Purity Asal)
                Frm115_UPAH_JUAL = Frm115.TB4 'Upah jualan
                Frm115_UPAH_MODAL = Frm115.L50_Text 'Upah modal
                Frm115_LM_HARGA_SEMASA_999 = Frm115.TB2 'Harga semasa (jualan) (Purity 999.9)
                Frm115_LM_HARGA_SUPPLIER = Frm115.L49_Text 'Harga per gram (harga semasa) dari supplier (modal)
                Frm115_LM_BERAT_999 = Frm115.L7_Text 'Berat emas dalam purity 999.9
                
                rs!upah_modal = Frm115.L50_Text 'Upah modal
                rs!harga_per_gram_supplier = Frm115.L49_Text 'Harga per gram (harga semasa) dari supplier (modal)
                rs!untung2 = Format(((Frm115_LM_BERAT_999 * Frm115_LM_HARGA_SEMASA_999) + Frm115_UPAH_JUAL) - ((Frm115_LM_BERAT_JUAL_ASAL * Frm115_LM_HARGA_SUPPLIER) + (Frm115_LM_BERAT_JUAL_ASAL / Frm115_LM_BERAT_ASAL) * Frm115_UPAH_MODAL), "0.00") 'Untung jika restok pada harga supplier ini
                
            Else
            
                rs!harga_per_gram_supplier = "0.00" 'Harga per gram (harga semasa) dari supplier (modal)
                rs!untung2 = "0.00" 'Untung jika restok pada harga supplier ini
                rs!upah_modal = "0.00" 'Upah modal
                
            End If
            
            If Format(Frm115.L6_Text, "0.00") = Format(Frm115.TB3, "0.00") Then
                rs!potong_flag = 0 '0 : Tiada Potong , 1 : Ada Potong
            Else
                rs!potong_flag = 1 '0 : Tiada Potong , 1 : Ada Potong
            End If
            rs!dulang = Frm115_LM_DULANG 'Dulang
            If Frm115.TB7 <> vbNullString Then
                rs!pemalar_tukaran_999 = Frm115.TB7 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
            Else
                rs!pemalar_tukaran_999 = Null 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
            End If
            If Frm115.L7_Text <> vbNullString Then
                rs!berat_999 = Format(Frm115.L7_Text, "0.00") 'Berat jualan dalam purity 999.9
            Else
                rs!berat_999 = Null 'Berat jualan dalam purity 999.9
            End If
            
            rs.Update

            Frm115_LM_DATA_SAVE = 1
        End If
        
        rs.Close
        Set rs = Nothing
'### Masukkan Data Ke Dalam Temp Table ### - End
        
        If Frm115_LM_DATA_SAVE = 1 Then
            
            GM_NEXT_PREV = 2
            
            Call Frm115_reset_1
            Call Frm115_Senarai_Jualan_Header
            Call Frm115_Senarai_Jualan
            
            MsgBox "Data yang telah diedit telah berjaya dimasukkan ke dalam senarai jualan.", vbInformation, "Info"
            
            Frm115.TB1.SetFocus
        End If
    End If
End If
End Sub

Private Sub CMD21_Click()
'on error resume next
Dim frm115_LM_CURR_PAGE As Double
Dim frm115_LM_TOTAL_PAGE As Double

frm115_LM_CURR_PAGE = 0
frm115_LM_TOTAL_PAGE = 0

If Frm115.L67_Text <> vbNullString And IsNumeric(Frm115.L67_Text) Then
    If Frm115.L68_Text <> vbNullString And IsNumeric(Frm115.L68_Text) Then
        frm115_LM_CURR_PAGE = Frm115.L67_Text
        frm115_LM_TOTAL_PAGE = Frm115.L68_Text
        
        If frm115_LM_CURR_PAGE <> 1 And frm115_LM_CURR_PAGE <> 0 Then
        
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
        
            Call Frm115_Senarai_Jualan_Header
            Call Frm115_Senarai_Jualan
            
        End If

    End If
End If
End Sub
Private Sub CMD22_Click()
'on error resume next
Dim frm115_LM_CURR_PAGE As Double
Dim frm115_LM_TOTAL_PAGE As Double

frm115_LM_CURR_PAGE = 0
frm115_LM_TOTAL_PAGE = 0

If Frm115.L67_Text <> vbNullString And IsNumeric(Frm115.L67_Text) Then
    If Frm115.L68_Text <> vbNullString And IsNumeric(Frm115.L68_Text) Then
        frm115_LM_CURR_PAGE = Frm115.L67_Text
        frm115_LM_TOTAL_PAGE = Frm115.L68_Text
        
        If frm115_LM_CURR_PAGE < frm115_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm115_Senarai_Jualan_Header
            Call Frm115_Senarai_Jualan
            
        End If
    End If
End If
End Sub

Private Sub CMD3_Click()
'On Error Resume Next
Note = "Adakah anda ingin batalkan urusan edit data ini ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then

    Call Frm115_reset_1
    
    Frm115.CMD1.Visible = True
    Frm115.CMD2.Visible = False
    Frm115.CMD3.Visible = False
    
    Frm115.TB1.SetFocus
    
End If
End Sub

Private Sub CMD4_Click()
'On Error Resume Next
Dim LM_KADAR As Double

If Frm115.TB7 = vbNullString Then
    
    MsgBox "Sila masukkan kadar tukaran purity 999.9.", vbExclamation, "Info"
    
    Exit Sub
    
End If
If Frm115.TB7 <> vbNullString And Not IsNumeric(Frm115.TB7) Then
    
    MsgBox "Hanya NOMBOR dibenarkan dalam ruangan kadar tukaran purity 999.9.", vbExclamation, "Info"
    
    Exit Sub
    
End If
If Frm115.TB7 <> vbNullString And IsNumeric(Frm115.TB7) Then
    
    LM_KADAR = 0
    LM_KADAR = Frm115.TB7
    
    If LM_KADAR > 1 Then

        MsgBox "[Kadar Tukaran Purity 999.9] tidak boleh lebih dari 1.00.", vbExclamation, "Info"
    
        Exit Sub
        
    End If
    
End If

If Frm115.CB2 = 0 And Frm115.CB3 = 0 And Frm115.CB4 = 0 Then

    MsgBox "Sila buat pilihan jenis cukai GST.", vbExclamation, "Info"
    
    Exit Sub

End If

Frm115.Frame1.Visible = True
End Sub

Private Sub CMD5_Click()
'On Error Resume Next
Frm115.Frame1.Visible = False
End Sub

Private Sub CMD6_Click()
'on error resume next
Dim frm115_LM_CURR_PAGE As Double
Dim frm115_LM_TOTAL_PAGE As Double

frm115_LM_CURR_PAGE = 0
frm115_LM_TOTAL_PAGE = 0

If Frm115.L61_Text <> vbNullString And IsNumeric(Frm115.L61_Text) Then
    If Frm115.L62_Text <> vbNullString And IsNumeric(Frm115.L62_Text) Then
        frm115_LM_CURR_PAGE = Frm115.L61_Text
        frm115_LM_TOTAL_PAGE = Frm115.L62_Text
        
        If frm115_LM_CURR_PAGE <> 1 And frm115_LM_CURR_PAGE <> 0 Then
        
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
                    
            Call frm115_gdn_list_header
            Call frm115_gdn_list
            
        End If

    End If
End If
End Sub
Private Sub CMD7_Click()
'On Error Resume Next
If Frm115.TB1 = vbNullString Then

    MsgBox "Sila masukkan No. Siri Produk.", vbExclamation, "Error"
    
    Frm115.TB1.SetFocus
    Exit Sub

End If

If InStr(1, Frm115.TB1, "*") <> 0 Or InStr(1, Frm115.TB1, "/") <> 0 Or InStr(1, Frm115.TB1, "\") <> 0 Or InStr(1, Frm115.TB1, "'") <> 0 Then

    MsgBox "No. Siri Produk Mengandungi Simbol Yang Tidak Sah.", vbInformation, "Info"
    
    Frm115.TB1 = vbNullString
    
    Frm115.TB1.SetFocus
    Exit Sub
    
End If

Call Frm115_reset_1
Call Frm115_Call_Product_Detail
End Sub
Private Sub CMD8_Click()
'On Error Resume Next
Dim Err(10)
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim Frm115_LM_CUKAI_ZR As Double
Dim Frm115_LM_CUKAI_SR As Double
Dim Frm115_LM_BERAT_ASAL As Double
Dim Frm115_LM_BERAT_JUALAN As Double
Dim LM_KADAR_TUKARAN As Double

LM_KADAR_TUKARAN = 0
Frm115_LM_KATEGORI = 0
Frm115_LM_BERAT_ASAL = 0 'Berat Asal (g)
Frm115_LM_BERAT_JUALAN = 0 'Berat Jualan (g)
Frm115_LM_CUKAI_ZR = 0 'Jumlah cukai GST ZR
Frm115_LM_CUKAI_SR = 0 'Jumlah cukai GST SR

Frm115_LM_Flag_SIMPANAN = 0 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai

If Frm115.L43_Text = "0" Then
    x = x + 1
    Err(x) = "Tiada senarai jualan."
End If
If Frm115.CBB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih supplier/agen yang membuat belian ini."
End If
If Frm115.CBB4 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih nama pekerja."
End If
If Frm115.TB8 = vbNullString Or (Frm115.TB8 <> vbNullString And Not IsNumeric(Frm115.TB8)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar Tukaran Purity 999.9]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If (Frm115.TB8 <> vbNullString And IsNumeric(Frm115.TB8)) Then
    
    LM_KADAR_TUKARAN = Frm115.TB8
    
    If LM_KADAR_TUKARAN > 1 Then
    
        x = x + 1
        Err(x) = "[Kadar Tukaran Purity 999.9] tidak boleh lebih dari 1.00. Hanya NOMBOR dibenarkan dalam ruangan ini."
        
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

    Note = "Adakah anda ingin teruskan urusan jualan ini ?" & vbCrLf & _
            vbNullString & vbCrLf & _
            "Data jualan akan disimpan ke dalam sistem."
                
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
        
' ### Periksa No. GDN ### - Start
        GoTo a:
        
        LM_NO_GDN = 1
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from default_setting where Default1='" & "Default" & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!no_gdn) Then
                If IsNumeric(rs!no_gdn) Then LM_NO_GDN = rs!no_gdn 'No. GDN
            End If
            
        End If
        
        rs.Close
        Set rs = Nothing
        
Re_gen_no_resit:
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 77_gdn_grn where no_rujukan='" & "GDN" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_NO_GDN, "000000") & "' AND jenis_urusan = 0", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            LM_NO_GDN = LM_NO_GDN + 1
            
            rs.Close
            Set rs = Nothing
            
            GoTo Re_gen_no_resit:
            
        End If
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from default_setting where Default1='" & "Default" & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            rs!no_gdn = LM_NO_GDN + 1 'No. GDN
            rs.Update
        
        End If
        
        rs.Close
        Set rs = Nothing
' ### Periksa No. GDN ### - End

a:

'---------------------------------------No. Invoice
        LM_NOW = Now
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        rs.Open "select * from 6_gdn", cn2, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        rs!tarikh = Frm115.DTPicker1
        rs!terminal = G_TERMINAL
        rs!write_timestamp = LM_NOW
        rs!Status = 1
        rs!nama_staff = MDI_frm1.L3_Text
        rs.Update
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        rs.Open "select * from 6_gdn where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & Frm115.DTPicker1 & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!ID) Then
                
                LM_NO_GDN = rs!ID
                rs!no_rujukan = "GDN" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_NO_GDN, "000000")
                
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
'---------------------------------------No. Invoice

        If Frm115.CBB4 <> vbNullString Then
        
            Frm115_LM_EMP_NAMA = Split(Frm115.CBB4, "  |  ")(0)
            Frm115_LM_EMP_NO = Split(Frm115.CBB4, "  |  ")(1)
            
        End If
            
'### Masukkan maklumat Good Delivery Note (GDN) ### - Start
        LM_NOW = Now
        LM_TARIKH = DateTime.Date$
        LM_MASA = DateTime.Time$
        
        LM_GDN_RE_GEN = 0
        
Re_gen_no_resit2:

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 77_gdn_grn where no_rujukan='" & "GDN" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_NO_GDN, "000000") & "' AND jenis_urusan = 0", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
        
            rs.AddNew
            rs!tarikh = Frm115.DTPicker1
            rs!masa = LM_MASA
            rs!write_timestamp = LM_NOW
            
            rs!no_rujukan = "GDN" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_NO_GDN, "000000")
            G_No_RESIT_JUALAN = "GDN" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_NO_GDN, "000000")
            
            If Frm115.L48_Text <> vbNullString Then
                rs!Berat_Asal = Format(Frm115.L48_Text, "0.00")
            Else
                rs!Berat_Asal = "0.00"
            End If
            If Frm115.TB8 <> vbNullString Then
                rs!kadar_tukaran = Frm115.TB8
            Else
                rs!kadar_tukaran = "0.00"
            End If
            If Frm115.L9_Text <> vbNullString Then
                rs!berat_tukaran = Format(Frm115.L9_Text, "0.00")
            Else
                rs!berat_tukaran = Null
            End If
            If Frm115.L51_Text <> vbNullString Then
                rs!harga_tanpa_gst = Format(Frm115.L51_Text, "0.00")
            Else
                rs!harga_tanpa_gst = "0.00"
            End If
            If Frm115.L52_Text <> vbNullString Then
                rs!jumlah_gst = Format(Frm115.L52_Text, "0.00")
            Else
                rs!jumlah_gst = "0.00"
            End If
            If Frm115.L21_Text <> vbNullString Then
                rs!kadar_gst = Format(Frm115.L21_Text, "0.00")
            Else
                rs!kadar_gst = "0.00"
            End If
            If Frm115.L53_Text <> vbNullString Then
                rs!harga_dengan_gst = Format(Frm115.L53_Text, "0.00")
            Else
                rs!harga_dengan_gst = Null
            End If
            If Frm115.TB2 <> vbNullString Then
                rs!harga_999 = Format(Frm115.TB2, "0.00")
            Else
                rs!harga_999 = "0.00"
            End If
            If Frm115.L12_Text <> vbNullString Then
                rs!nilaian_harga_emas = Format(Frm115.L12_Text, "0.00")
            Else
                rs!nilaian_harga_emas = "0.00"
            End If
            If Frm115.L17_Text <> vbNullString Then
                rs!gst_zr_harga = Format(Frm115.L17_Text, "0.00")
            Else
                rs!gst_zr_harga = "0.00"
            End If
            If Frm115.L18_Text <> vbNullString Then
                rs!gst_sr_harga = Format(Frm115.L18_Text, "0.00")
            Else
                rs!gst_sr_harga = "0.00"
            End If
            If Frm115.L19_Text <> vbNullString Then
                rs!gst_zr_cukai = Format(Frm115.L19_Text, "0.00")
            Else
                rs!gst_zr_cukai = "0.00"
            End If
            If Frm115.L20_Text <> vbNullString Then
                rs!gst_sr_cukai = Format(Frm115.L20_Text, "0.00")
            Else
                rs!gst_sr_cukai = "0.00"
            End If
            If Frm115.L43_Text <> vbNullString Then
                rs!bil_barang = Frm115.L43_Text
            Else
                rs!bil_barang = 0
            End If
            rs!Status = 1
            rs!jenis_urusan = 0
            rs!jenis = "GDN"
            rs!terminal = G_TERMINAL
            If Frm115.CBB2 <> vbNullString Then
                rs!supplier_agen = Frm115.CBB2
            Else
                rs!supplier_agen = Null
            End If
            rs!user = Frm115_LM_EMP_NAMA 'Nama Pekerja
            rs!cawangan = G_CAWANGAN
            G_KEDAI = G_CAWANGAN
            rs.Update
            DATA_SAVE = 1
            
        Else
        
            LM_NO_GDN = LM_NO_GDN + 1
            LM_GDN_RE_GEN = 1
            
            rs.Close
            Set rs = Nothing
            
            GoTo Re_gen_no_resit2:
            
        End If
        
        rs.Close
        Set rs = Nothing
'### Masukkan data voucher / invoice bagi belian agen ini ### - End

        'If LM_GDN_RE_GEN = 1 Then

        '    Set rs = New ADODB.Recordset
        '    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        '    rs.Open "select * from default_setting where Default1='" & "Default" & "'", cn, adOpenKeyset, adLockOptimistic
            
        '    If Not rs.EOF Then
                
        '        rs!no_gdn = LM_NO_GDN + 1 'No. GDN
        '        rs.Update
            
        '    End If
            
        '    rs.Close
        '    Set rs = Nothing
        
        'End If

'### Masukkan Data Jualan Ke Dalam Table Jualan ### - Start

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "insert into 23_senarai_jualan(no_resit,jenis_jualan,tarikh,no_pekerja,no_rujukan_pembeli,no_rujukan_agen_dropship,kategori_pembeli,status_rekod,write_timestamp,jualan_online,bil_rasmi,status_r,no_siri_produk,kategori_produk,purity,berat_asal,berat_jualan,harga_semasa,upah,harga_asal,diskaun,harga_lepas_diskaun,adjustment,harga_jualan,gst_ari_nashi,kadar_gst,jumlah_gst,harga_dengan_gst,dropship,komisyen_per_gram,jumlah_komisyen,status,type,potong_flag,harga_per_gram_modal,modal,untung,harga_per_gram_supplier,untung2,dulang,gst_include,harga_tanpa_gst,harga_koperasi,kadar_penurunan_upah,kadar_penurunan_bp,harga_semasa_staff,harga_bp_asal,upah_asal,harga_staff,komisyen_staff,pemalar_tukaran_999,berat_999,upah_modal,gst_barang_atau_upah,harga_jualan_dengan_gst,kadar_komisyen_upah,komisyen_upah,jualan_per_gram,modal_per_gram,flag_upah,upah_per_gram,jenis_urusan,terminal)" & _
                    "select '" & G_No_RESIT_JUALAN & "',1,'" & Frm115.DTPicker1 & "','" & Frm115_LM_EMP_NO & "','" & LM_NO_PEMBELI & "','" _
                    & LM_NO_DROPSHIP & "','" & LM_KATEGORI & "',1,'" & LM_NOW & "',0,1,0,no_siri_produk,kategori_produk," _
                    & "purity,berat_asal,berat_jualan,harga_semasa,upah,harga_asal,diskaun,harga_lepas_diskaun,adjustment,harga_jualan," _
                    & "gst_ari_nashi,kadar_gst,jumlah_gst,harga_dengan_gst,dropship,komisyen_per_gram,jumlah_komisyen,status_jualan,type," _
                    & "potong_flag,harga_per_gram_modal,modal,untung,harga_per_gram_supplier,untung2,dulang,gst_include,harga_tanpa_gst," _
                    & "harga_koperasi,kadar_penurunan_upah,kadar_penurunan_bp,harga_semasa_staff,harga_bp_asal,upah_asal,harga_staff," _
                    & "komisyen_staff,pemalar_tukaran_999,berat_999,upah_modal,gst_barang_atau_upah,harga_jualan_dengan_gst," _
                    & "kadar_komisyen_upah,komisyen_upah,jualan_per_gram,modal_per_gram,flag_upah,upah_per_gram,'" & G_JENIS_URUSAN & "','" & G_TERMINAL & "'" _
                    & "from " & G_GDN_TEMP & " WHERE status='" & 1 & "'"
            
        Set rs = cn.Execute(strsql)
        Set rs = Nothing


'### Update status & info item yang terjual ### - Start

        Dim Frm84_LM_BERAT_ASAL As Double
        Dim Frm84_LM_BEZA_BERAT As Double
        Dim Frm84_LM_BERAT_JUALAN As Double
        Dim Frm84_LM_BERAT_SUSUT As Double
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_GDN_TEMP & " where status='" & "1" & "'", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
        
            Set rs2 = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs2.Open "select * from Data_Database where no_siri_produk='" & rs!no_siri_Produk & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs2.EOF Then
                
                G_ID = rs2!ID
                Call recovery_data_database
                
                If rs!Type = 0 Then
                
                    Frm84_LM_BERAT_ASAL = 0
                    Frm84_LM_BEZA_BERAT = 0
                    Frm84_LM_BERAT_JUALAN = 0
                    Frm84_LM_BERAT_SUSUT = 0
                    
                    'If Not IsNull(rs2!berat_asal) Then
                    '    If IsNumeric(rs2!berat_asal) Then Frm84_LM_BERAT_ASAL = rs2!berat_asal 'Berat Asal (g)
                    'End If
                    If Not IsNull(rs2!beza_berat) Then
                        If IsNumeric(rs2!beza_berat) Then Frm84_LM_BEZA_BERAT = rs2!beza_berat 'Beza Berat (g)
                    End If
                    If Not IsNull(rs!berat_jualan) Then
                        If IsNumeric(rs!berat_jualan) Then Frm84_LM_BERAT_JUALAN = rs!berat_jualan 'Berat Jualan (g)
                    End If
                    If Not IsNull(rs2!susut_berat) Then
                        If IsNumeric(rs2!susut_berat) Then Frm84_LM_BERAT_SUSUT = rs2!susut_berat 'Susut berat Jualan (g)
                    End If
                    
                    If Format(Frm84_LM_BERAT_JUALAN, "0.00") = Format(Frm84_LM_BEZA_BERAT, "0.00") Then
                        rs2!beza_berat = "0.00" 'Baki Berat
                        'rs2!susut_berat = "0.00" 'Susut berat
                        rs2!StatusItem = 27
                        rs2!tarikh_jualan1 = Null
                    Else
                        rs2!beza_berat = Format(Frm84_LM_BEZA_BERAT - Frm84_LM_BERAT_JUALAN, "0.00") 'Baki Berat
                        'rs2!susut_berat = "0.00" 'Susut berat
                        rs2!StatusItem = 28
                        rs2!tarikh_jualan1 = Frm115.DTPicker1
                    End If
                Else
                    rs2!StatusItem = 27
                End If
                
                rs2!write_timestamp2 = LM_NOW
                rs2!no_pekerja = Frm115_LM_EMP_NO
                rs2!terminal = G_TERMINAL
                rs2!Menu = 5

                rs2.Update
            End If
            
            rs2.Close
            Set rs2 = Nothing

            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
'### Update status & info item yang terjual ### - End
    
        If DATA_SAVE = 1 Then
    '#### Update Log Aktiviti Sistem #### - Start
            'User = MDI_frm1.L3_Text
            LogAct_Memory = "[" & Frm115_LM_EMP_NAMA & "] Pengeluaran GDN kepada agen/supplier (Per Item). No. Rujukan [" & G_No_RESIT_JUALAN & "]."
            LogDate_Memory = LM_NOW
            Call UpdateLog_Database
    '#### Update Log Aktiviti Sistem #### - End
            
            Call Frm115_reset_main
            Call Frm115_reset_1
            Call Frm115_reset_2
            Call Frm115_reset_3
            Call Frm115_Senarai_Jualan_Header
            Call Frm115_Senarai_Jualan
            Call frm115_reset_gdn_list
            
            Frm115.TB1.SetFocus
            
            Call Frm115_cetak_gdn

        End If
        
    End If
End If
End Sub
Private Sub CMD9_Click()
'On Error Resume Next
Unload Frm115
MDI_frm1.L5_Text = 0
End Sub

Private Sub Form_Load()
'On Error Resume Next
Frm115.L43_Text = 0 'Jumlah bilangan barang jualan

Frm115.L48_Text = "0.00" 'Jumlah berat (g)
'GLOBAL_DISABLE = 0
'Frm115.TB1 = vbNullString

'Call Frm115_reset_1
'Call Frm115_reset_2
'Call Frm115_reset_3
'Call Frm115_reset_main

'Frm115.L26_Text.BackStyle = 0
'Frm115.L27_Text.BackStyle = 0

'Frm115.DTPicker1 = DateTime.Date$
End Sub

Private Sub Frm115_scan_mode_Click()
'on error resume next
Frm115.TB1 = vbNullString
Frm115.TB1.SetFocus
End Sub

Private Sub Frm115_SM_edit_data1_Click()
'on error resume next
DATA_FOUND = 0

Frm115_LM_No_ID = vbNullString

If IsNumeric(Frm115.LV2.SelectedItem.Index) Then
    
    Frm115_LM_No_ID = Frm115.LV2.ListItems(Frm115.LV2.SelectedItem.Index)
    
    If Frm115_LM_No_ID <> vbNullString Then

        Call Frm115_reset_1 '!! Hati-hati dengan tempat letakkan command ini!!
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_GDN_TEMP & " where ID='" & Frm115_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!ID) Then Frm115.L24_Text = rs!ID 'No. ID Database
            If Not IsNull(rs!no_siri_Produk) Then Frm115.L3_Text = rs!no_siri_Produk 'No. Siri Produk
            If Not IsNull(rs!kategori_Produk) Then Frm115.L5_Text = rs!kategori_Produk 'Kategori Produk
            If Not IsNull(rs!purity) Then Frm115.L4_Text = rs!purity 'Purity
            If Not IsNull(rs!Berat_Asal) Then Frm115.L6_Text = Format(rs!Berat_Asal, "#,##0.00") 'Berat Asal (g)
            If Not IsNull(rs!berat_jualan) Then Frm115.TB3 = Format(rs!berat_jualan, "#,##0.00") 'Berat Jualan (g)
            If Not IsNull(rs!harga_Semasa) Then Frm115.TB2 = Format(rs!harga_Semasa, "#,##0.00") 'Harga Semasa (RM/g)
            If Not IsNull(rs!UPAH) Then Frm115.TB4 = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
            If Not IsNull(rs!gst_ari_nashi) Then 'Harga Jualan (RM)
                If rs!gst_ari_nashi = "ZR (L)" Then
                    Frm115.CB2 = 1 '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                    If Not IsNull(rs!jumlah_gst) Then
                        Frm115.TB5 = Format(rs!jumlah_gst, "#,##0.00") 'Jumlah Cukai GST (RM)
                    Else
                        Frm115.TB5 = "0.00"
                    End If
                ElseIf rs!gst_ari_nashi = "SR" Then
                    Frm115.CB3 = 1 '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                    If Not IsNull(rs!kadar_gst) Then
                        Frm115.L21_Text = rs!kadar_gst 'Kadar Cukai GST (%)
                    End If
                    If Not IsNull(rs!jumlah_gst) Then
                        Frm115.TB5 = rs!jumlah_gst 'Jumlah Cukai GST (RM)
                    Else
                        Frm115.TB5 = "0.00"
                    End If
                    If Not IsNull(rs!gst_include) Then
                        If rs!gst_include = 0 Then
                            Frm115.CB4 = 0
                        ElseIf rs!gst_include = 1 Then
                            Frm115.CB4 = 1
                        End If
                    Else
                        Frm115.CB4 = 0
                    End If
                End If
            End If
            If Not IsNull(rs!harga_tanpa_gst) Then Frm115.L30_Text = Format(rs!harga_tanpa_gst, "#,##0.00") 'Harga Jualan Tanpa GST (RM)
            If Not IsNull(rs!harga_dengan_gst) Then Frm115.TB6 = Format(rs!harga_dengan_gst, "#,##0.00") 'Harga Jualan Termasuk GST (RM)
            If Not IsNull(rs!harga_per_gram_modal) Then Frm115.L33_Text = Format(rs!harga_per_gram_modal, "#,##0.00") 'Harga Per Gram Bagi Modal (RM/g)
            If Not IsNull(rs!berat_999) Then Frm115.L7_Text = rs!berat_999 'Berat jualan dalam purity 999.9
            If Not IsNull(rs!harga_per_gram_supplier) Then Frm115.L49_Text = Format(rs!harga_per_gram_supplier, "#,##0.00") 'Harga per gram (harga semasa) dari supplier (modal)
            If Not IsNull(rs!upah_modal) Then Frm115.L50_Text = Format(rs!upah_modal, "#,##0.00") 'Upah modal
            If Not IsNull(rs!pemalar_tukaran_999) Then Frm115.TB7 = Format(rs!pemalar_tukaran_999, "#,##0.00")

            DATA_FOUND = 1
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then
            Frm115.TB1.Locked = True
            Frm115.TB1 = vbNullString
            Frm115.TB1.BackColor = &H8000000A
            
            Frm115.CMD1.Visible = False
            Frm115.CMD2.Visible = True
            Frm115.CMD3.Visible = True
        End If
        
    Else
    
        MsgBox "Tiada data.", vbExclamation, "Info"
        
    End If
    
End If
End Sub



Private Sub Frm115_SM_remove_jualan_Click()
'on error resume next
DATA_FOUND = 0
Frm115_LM_No_ID = vbNullString

If IsNumeric(Frm115.LV2.SelectedItem.Index) Then
    
    Frm115_LM_No_ID = Frm115.LV2.ListItems(Frm115.LV2.SelectedItem.Index)
    
    If Frm115_LM_No_ID <> vbNullString Then
        
        Note = "Adakah anda ingin keluarkan item ini dari senarai jualan dan pulangkan ke stok kedai ?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        If Answer = vbNo Then
            'Exit Sub
        End If
        If Answer = vbYes Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from " & G_GDN_TEMP & " where ID='" & Frm115_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!Status) Then
'### Kod Bagi Status ###
'==========================================
'0 : Tidak Aktif (Tidak Perlu Buat Apa-Apa)
'1 : Aktif - Kemasukkan Data Baru
'2 : Barang Sudah Terjual ****Tidak Dibenarkan Untuk Diedit Atau Dipadamkan
'3 : Kemasukkan Data Baru
'4 : Data Diedit
'5 : Keluarkan Data Dari Database (Bagi Item Asal Yang Dipadamkan)
'6 : Ignore Kemasukkan Data Ke Dalam Database
                    If rs!Status = 1 Or rs!Status = 4 Then
                        rs.Delete
                        rs.Update
                        
                        DATA_FOUND = 1
                    ElseIf rs!Status = 2 Or rs!Status = 3 Then
                        rs!Status = 5
                        rs.Update
                        
                        DATA_FOUND = 1
                    End If
                End If
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_FOUND = 1 Then
                
                GM_NEXT_PREV = 2 '0 : Next , 1 : Previous
                
                Call Frm115_reset_1
                
                Call Frm115_Senarai_Jualan_Header
                Call Frm115_Senarai_Jualan
                
                MsgBox "Item telah dikeluarkan dari senarai jualan.", vbInformation, "Info"
                
            End If
        End If

    Else
        
        MsgBox "Tiada data.", vbExclamation, "Info"
        
    End If
    
End If
End Sub
Private Sub L11_Text_Change()
'On Error Resume Next
Call Frm115_calc5
End Sub

Private Sub frm115_sm_select_item_Click()
'on error resume next
Dim LM_KADAR As Double
LM_DATA_FOUND = 0
Frm115_LM_No_ID = vbNullString
Frm115_LM_KOD_PURITY = vbNullString

Frm115_LM_KATEGORI = 1
'1:  Pelanggan Biasa
'2 : Member / Ahli
'3:  Pengedar
'4:  RAF
'5:  Normal Dealer(ND)
'6:  Master Dealer(MD)

Frm115_LM_No_ID = vbNullString

If IsNumeric(Frm115.LV1.SelectedItem.Index) Then
    
    Frm115_LM_No_ID = Frm115.LV1.ListItems(Frm115.LV1.SelectedItem.Index)
    
    If Frm115_LM_No_ID <> vbNullString Then
    
            
            If Frm115.TB7 = vbNullString Then
                
                MsgBox "Sila masukkan kadar tukaran purity 999.9.", vbExclamation, "Info"
                
                Exit Sub
                
            End If
            If Frm115.TB7 <> vbNullString And Not IsNumeric(Frm115.TB7) Then
                
                MsgBox "Hanya NOMBOR dibenarkan dalam ruangan kadar tukaran purity 999.9.", vbExclamation, "Info"
                
                Exit Sub
                
            End If
            If Frm115.TB7 <> vbNullString And IsNumeric(Frm115.TB7) Then
                
                LM_KADAR = 0
                LM_KADAR = Frm115.TB7
                
                If LM_KADAR > 1 Then
            
                    MsgBox "[Kadar Tukaran Purity 999.9] tidak boleh lebih dari 1.00.", vbExclamation, "Info"
                
                    Exit Sub
                    
                End If
                
            End If
            
            If Frm115.CB2 = 0 And Frm115.CB3 = 0 And Frm115.CB4 = 0 Then
            
                MsgBox "Sila buat pilihan jenis cukai GST.", vbExclamation, "Info"
                
                Exit Sub
            
            End If

'###Periksa Mode Upah### - Start
            'Set rs = New ADODB.Recordset
            'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            'rs.Open "select * from default_setting where Default1='" & "Default" & "'", cn, adOpenKeyset, adLockOptimistic
            
            'If Not rs.EOF Then
            '    If Not IsNull(rs!flag_upah) Then
            '        If rs!flag_upah = 1 Then
            '            LM_UPAH_MODE = 1
            '        Else
            '            LM_UPAH_MODE = 0
            '        End If
            '    End If
            'End If
            
            'rs.Close
            'Set rs = Nothing
            
            If G_UPAH_MODE = 1 Then
                LM_UPAH_MODE = 1
            Else
                LM_UPAH_MODE = 0
            End If
'###Periksa Mode Upah### - End
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from data_database where ID='" & Frm115_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                Note = "Adakah anda yakin ingin masukkan barang ini ke dalam senarai?" & vbCrLf & _
                        "No. Siri Produk : " & rs!no_siri_Produk & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Teruskan."
                        
                Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
                If Answer = vbYes Then
                    
                    If rs!gdn_temp = 0 Then
                    
                        If rs!StatusItem = "10" Or rs!StatusItem = "12" Or rs!StatusItem = "20" Or rs!StatusItem = "22" Or rs!StatusItem = "28" Then
  
                            If Not IsNull(rs!receiving_Status) Then
                            
                                If rs!receiving_Status = 0 Or rs!receiving_Status = 2 Or rs!receiving_Status = 4 Or rs!receiving_Status = 5 Then
                                    
                                    If Not IsNull(rs!no_siri_Produk) Then Frm115.L3_Text = rs!no_siri_Produk 'No. Siri Produk
                                    If Not IsNull(rs!beza_berat) Then Frm115.L6_Text = Format(rs!beza_berat, "0.00") 'Berat Asal (g)
                                    If Not IsNull(rs!beza_berat) Then Frm115.TB3 = Format(rs!beza_berat, "0.00") 'Berat Jualan (g)
                                    If Not IsNull(rs!harga_Per_Gram_Item) Then Frm115.L33_Text = Format(rs!harga_Per_Gram_Item, "0.00") 'Harga Per Gram Item (RM/g)
                                    If Not IsNull(rs!UPAH) Then Frm115.L50_Text = rs!UPAH 'Upah modal
                                    Frm115.TB4.Locked = False 'Upah
                                    Frm115.TB4.BackColor = &HFFFFFF 'Upah
                                    
                                    Frm115_LM_READY_TO_SAVE = 1 'Flag : Ready To Save
                                Else
                                    MsgBox "Barang yang ingin dimasukkan ke dalam senarai [" & rs!no_siri_Produk & "] adalah barang permata." & vbCrLf & _
                                            vbNullString & vbCrLf & _
                                            "Hanya barang kemas (yang mempunyai berat) SAHAJA dibenarkan dalam menu ini.", vbInformation, "Info"
                                            
                                    Frm115.TB1 = vbNullString
                                    Frm115.TB1.SetFocus
                                         
                                    rs.Close
                                    Set rs = Nothing
                                    
                                    Exit Sub
                                End If
                            'End If
            
                                If LM_UPAH_MODE = 1 And Frm115.CB5 = 1 Then
                                    If Frm115_LM_KATEGORI = 1 Then
                                        If Not IsNull(rs!Upah_Jualan) Then
                                            Frm115.TB4 = Format(rs!Upah_Jualan, "0.00") 'Upah Pelanggan
                                        End If
                                    ElseIf Frm115_LM_KATEGORI = 2 Then
                                        If Not IsNull(rs!Upah_Member) Then
                                            Frm115.TB4 = Format(rs!Upah_Member, "0.00") 'Upah Member
                                        End If
                                    ElseIf Frm115_LM_KATEGORI = 3 Then
                                        If Not IsNull(rs!Upah_RAF) Then
                                            Frm115.TB4 = Format(rs!Upah_RAF, "0.00") 'Upah RAF
                                        End If
                                    ElseIf Frm115_LM_KATEGORI = 4 Then
                                        If Not IsNull(rs!Upah_Pengedar) Then
                                            Frm115.TB4 = Format(rs!Upah_Pengedar, "0.00") 'Upah Pengedar
                                        End If
                                    ElseIf Frm115_LM_KATEGORI = 5 Then
                                        If Not IsNull(rs!upah_normal_dealer) Then
                                            Frm115.TB4 = Format(rs!upah_normal_dealer, "0.00") 'Upah Normal Dealer
                                        End If
                                    ElseIf Frm115_LM_KATEGORI = 6 Then
                                        If Not IsNull(rs!upah_master_dealer) Then
                                            Frm115.TB4 = Format(rs!upah_master_dealer, "0.00") 'Upah Master Dealer
                                        End If
                                    End If
                                Else
                                    Frm115.TB4 = Format(0, "0.00") 'Upah
                                End If
                            
                                If Not IsNull(rs!kategori_Produk) Then Frm115.L5_Text = rs!kategori_Produk 'Kategori Produk
                                If Not IsNull(rs!kod_Purity) Then
                                    Frm115_LM_KOD_PURITY = rs!kod_Purity 'Kod Purity
                                    Frm115.L4_Text = rs!kod_Purity 'Kod Purity
                                End If
    
                                rs!gdn_temp = 1
                                rs.Update
                                
                                LM_DATA_FOUND = 1
                                
                            Else
                                
                                MsgBox "Barang ini tidak dibenarkan untuk dimasukkan ke dalam senarai.", vbExclamation, "Info"
                            
                            End If
                            
                        ElseIf rs!StatusItem = "11" Then
                            MsgBox "Item ini telah dijual. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
    
                        'ElseIf rs!statusitem = "12" Then
                        '    MsgBox "Item ini telah dijual secara potong. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
    
                        ElseIf rs!StatusItem = "13" Then
                            MsgBox "Item ini telah dijual secara potong. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
    
                        ElseIf rs!StatusItem = "14" Or rs!StatusItem = "21" Then
                            MsgBox "Item ini telah ditempah oleh pelanggan. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
    
                        ElseIf rs!StatusItem = "15" Or rs!StatusItem = "19" Or rs!StatusItem = "20" Then
                            MsgBox "Item ini telah dibeli secara ansuran. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
    
                        ElseIf rs!StatusItem = "16" Then
                            MsgBox "Item ini telah dihantar ke Ar-Rahnu. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
    
                        ElseIf rs!StatusItem = "17" Then
                            MsgBox "Item ini telah dijual secara ETA. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
    
                        ElseIf rs!StatusItem = "23" Then
                            MsgBox "Item Ini Telah Dihantar Ke Supplier/Kilang. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
    
                        ElseIf rs!StatusItem = "24" Then
                            MsgBox "Item Ini Telah Dihantar Ke Supplier/Kilang. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
    
                        ElseIf rs!StatusItem = "25" Then
                            MsgBox "Item Ini Telah Diagihkan Ke Cawangan. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
    
                        ElseIf rs!StatusItem = "26" Then
                            MsgBox "Item Ini Telah Dijual Oleh Cawangan. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
    
                        ElseIf rs!StatusItem = "0" Then
                            MsgBox "Item Ini Telah Dipadamkan Dari Sistem. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
    
                        ElseIf rs!StatusItem = "30" Then
                            MsgBox "Item ini telah ditransfer ke kedai " & rs!nama_kedai & ".", vbExclamation, "Info"
    
                        ElseIf rs!StatusItem = "27" Then
                            MsgBox "Item Ini Telah Dijual Dari Menu GDN.", vbExclamation, "Info"

                        ElseIf rs!StatusItem = "29" Then
                            MsgBox "Item Ini Telah Diubah Status Kepada Hilang , Dicuri Dan Sebagainya.", vbExclamation, "Info"

                        End If
                        
                    ElseIf rs!gdn_temp = 1 Then
                        
                        LM_DATA_FOUND = 2
                        
                        'MsgBox "Item ini telah dimasukkan ke dalam senarai migrasi sebelum ini.", vbExclamation, "Info"
                        
                    End If
                
    
                End If
                
            End If
            
            rs.Close
            Set rs = Nothing
            
            If LM_DATA_FOUND = 1 Then
            
                If Frm115_LM_KOD_PURITY <> vbNullString Then
            
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from hargaemas where Purity='" & Frm115_LM_KOD_PURITY & "' AND cawangan='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
                    If Not rs.EOF Then
                        If Not IsNull(rs!HargaDariSupplier) Then
                            If IsNumeric(rs!HargaDariSupplier) Then
                                Frm115.L49_Text = rs!HargaDariSupplier
                            Else
                                Frm115.L49_Text = 0
                            End If
                        Else
                            Frm115.L49_Text = 0
                        End If
                    End If
                    
                    rs.Close
                    Set rs = Nothing
                    
                End If
                
                Call frm115_insert_data
                
                GM_NEXT_PREV = 2 '0 : Next , 1 : Previous
                
                Call Frm115_Senarai_Jualan_Header
                Call Frm115_Senarai_Jualan
    
                Call frm115_gdn_list_header
                Call frm115_gdn_list
                
                'MsgBox "Barang ini telah berjaya dimasukkan ke dalam senarai.", vbInformation, "Info"
                
            Else
                
                If LM_DATA_FOUND = 0 Then
                
                    MsgBox "Tiada data dijumpai/dibatalkan.", vbExclamation, "Info"
                
                ElseIf LM_DATA_FOUND = 2 Then
                
                    MsgBox "Item ini telah dimasukkan ke dalam senarai sebelum ini.", vbExclamation, "Info"
                    
                End If
                
            End If

        
    End If
    
End If
End Sub
Private Sub L30_Text_Change()
'On Error Resume Next
Call Frm115_calc3
End Sub

Private Sub L35_Text_Change()
'On Error Resume Next
Call Frm115_calc10
End Sub

Private Sub L36_Text_Change()
'On Error Resume Next
Call Frm115_calc10
End Sub
Private Sub L37_Text_Change()
'On Error Resume Next
Call Frm115_calc10
End Sub
Private Sub L38_Text_Change()
'On Error Resume Next
Call Frm115_calc10
End Sub
Private Sub L39_Text_Change()
'On Error Resume Next
Call Frm115_calc10
End Sub
Private Sub L40_Text_Change()
'On Error Resume Next
Call Frm115_calc10
End Sub
Private Sub L41_Text_Change()
'On Error Resume Next
Call Frm115_calc10
End Sub
Private Sub L42_Text_Change()
'On Error Resume Next
Call Frm115_calc10
End Sub
Private Sub L48_Text_Change()
'On Error Resume Next
Call Frm115_calc4
End Sub
Private Sub L51_Text_Change()
'On Error Resume Next
Call Frm115_calc10
End Sub

Private Sub L52_Text_Change()
'On Error Resume Next
Call Frm115_calc10
End Sub

Private Sub L9_Text_Change()
'On Error Resume Next
Call Frm115_calc5
End Sub

Private Sub LV1_DblClick()
'on error resume next
Frm115_LM_No_ID = vbNullString

If IsNumeric(Frm115.LV1.SelectedItem.Index) Then
    
    Frm115_LM_No_ID = Frm115.LV1.SelectedItem.Index
    
    If Frm115_LM_No_ID <> vbNullString Then
    
        PopupMenu Frm115_PM_menu4
    
    Else
    
        MsgBox "Tiada Data.", vbInformation, "Info"
        
    End If
    
Else

    MsgBox "Tiada Data.", vbInformation, "Info"
    
End If
End Sub



Private Sub LV2_DblClick()
'on error resume next
Frm115_LM_No_ID = vbNullString

If IsNumeric(Frm115.LV2.SelectedItem.Index) Then
    
    Frm115_LM_No_ID = Frm115.LV2.SelectedItem.Index
    
    If Frm115_LM_No_ID <> vbNullString Then
    
        Call Frm115_reset_1
        
        Frm115.CMD1.Visible = True
        Frm115.CMD2.Visible = False
        Frm115.CMD3.Visible = False
            
        PopupMenu Frm115_PM_menu
    
    Else
    
        MsgBox "Tiada Data.", vbInformation, "Info"
        
    End If
    
Else

    MsgBox "Tiada Data.", vbInformation, "Info"
    
End If
End Sub




Private Sub TB1_Change()
'on error resume next
If Frm115.CB1 = 1 And Frm115.TB1 <> vbNullString Then
    Frm115.Tmr2.Enabled = False
    Frm115.Tmr2.Enabled = True
    Frm115.Tmr2.Interval = 100
End If
End Sub
Private Sub TB2_Change()
'On Error Resume Next
Call Frm115_calc5
End Sub
Private Sub TB3_Change()
'On Error Resume Next
Call Frm115_calc1
End Sub
Private Sub TB4_Change()
'On Error Resume Next
Call Frm115_calc2
End Sub
Private Sub TB5_Change()
'On Error Resume Next
Call Frm115_calc3
End Sub
Private Sub TB7_Change()
'On Error Resume Next
Call Frm115_calc1
End Sub
Private Sub TB8_Change()
'On Error Resume Next
Call Frm115_calc4
End Sub
Private Sub Tmr2_Timer()
'On Error Resume Next
Dim Frm115_LM_LIMIT As Integer
Dim Frm115_LM_BIL As Integer

If Frm115.CB1 = 1 And Frm115.TB1 <> vbNullString And Frm115.Tmr2.Enabled = True Then

    If Frm115.Tmr2.Interval = 100 Then
    
        If InStr(1, Frm115.TB1, "*") <> 0 Or InStr(1, Frm115.TB1, "/") <> 0 Or InStr(1, Frm115.TB1, "\") <> 0 Or InStr(1, Frm115.TB1, "'") <> 0 Then
        
            MsgBox "No. Siri Produk Mengandungi Simbol Yang Tidak Sah.", vbInformation, "Info"
            
            Frm115.TB1 = vbNullString
            Exit Sub
            
        End If
        
        Call Frm115_reset_1
        Call Frm115_Call_Product_Detail
        
    End If
End If
End Sub


