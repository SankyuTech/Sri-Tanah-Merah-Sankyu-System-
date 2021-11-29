VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm123 
   Caption         =   "Goods Despatch Note (GDN) - Bulk"
   ClientHeight    =   13035
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13035
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.TextBox TB10 
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
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   126
      Text            =   "TB10"
      Top             =   1050
      Width           =   1500
   End
   Begin VB.PictureBox Pic1 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   17400
      ScaleHeight     =   3735
      ScaleWidth      =   6855
      TabIndex        =   100
      Top             =   960
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CommandButton CMD13 
         Caption         =   "Tutup paparan ini"
         Height          =   360
         Left            =   1680
         MouseIcon       =   "frm123.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   101
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Shape Shape4 
         Height          =   1575
         Left            =   120
         Top             =   480
         Width           =   5655
      End
      Begin VB.Line Line1 
         X1              =   2355
         X2              =   5635
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label L20_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L20_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         TabIndex        =   116
         Top             =   1680
         Width           =   1785
      End
      Begin VB.Label L19_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L19_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         TabIndex        =   115
         Top             =   1440
         Width           =   1785
      End
      Begin VB.Label L18_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L18_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   114
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
         TabIndex        =   113
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
         TabIndex        =   112
         Top             =   1200
         Width           =   1800
      End
      Begin VB.Label Label118 
         BackStyle       =   0  'Transparent
         Caption         =   "Standard Rated (SR)"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   111
         Top             =   1680
         Width           =   2145
      End
      Begin VB.Label Label121 
         BackStyle       =   0  'Transparent
         Caption         =   "Zero Rated (ZR)"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   110
         Top             =   1440
         Width           =   2145
      End
      Begin VB.Label L17_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L17_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   109
         Top             =   1440
         Width           =   1785
      End
      Begin VB.Label L16_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L16_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3075
         TabIndex        =   108
         Top             =   840
         Width           =   1545
      End
      Begin VB.Label Label115 
         BackStyle       =   0  'Transparent
         Caption         =   ": RM"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2640
         TabIndex        =   107
         Top             =   840
         Width           =   600
      End
      Begin VB.Label L15_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L15_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3075
         TabIndex        =   106
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label Label111 
         BackStyle       =   0  'Transparent
         Caption         =   ": RM"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2640
         TabIndex        =   105
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label117 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Harga Dengan GST"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   104
         Top             =   840
         Width           =   2505
      End
      Begin VB.Label Label114 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Harga Tanpa GST"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   103
         Top             =   600
         Width           =   2505
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Maklumat cukai GST"
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
         TabIndex        =   102
         Top             =   240
         Width           =   5385
      End
   End
   Begin VB.TextBox TB9 
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
      Left            =   12240
      TabIndex        =   10
      Text            =   "TB9"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1500
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
      Left            =   14520
      MaxLength       =   5
      TabIndex        =   8
      Text            =   "TB8"
      Top             =   1260
      Width           =   1620
   End
   Begin VB.CommandButton CDM13 
      Caption         =   "Papar Maklumat Terperinci GST"
      Height          =   360
      Left            =   10200
      MouseIcon       =   "frm123.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   117
      Top             =   4080
      Width           =   2895
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
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   97
      Text            =   "TB6"
      Top             =   5160
      Width           =   1380
   End
   Begin VB.CommandButton CMD22 
      BackColor       =   &H00FFFFFF&
      Height          =   650
      Left            =   8760
      MaskColor       =   &H00400000&
      MouseIcon       =   "frm123.frx":0614
      MousePointer    =   99  'Custom
      Picture         =   "frm123.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   91
      ToolTipText     =   "Paparan seterusnya"
      Top             =   9960
      Width           =   1100
   End
   Begin VB.CommandButton CMD21 
      BackColor       =   &H00FFFFFF&
      Height          =   650
      Left            =   7560
      MaskColor       =   &H00400000&
      MouseIcon       =   "frm123.frx":1244
      MousePointer    =   99  'Custom
      Picture         =   "frm123.frx":154E
      Style           =   1  'Graphical
      TabIndex        =   90
      ToolTipText     =   "Paparan sebelumnya"
      Top             =   9960
      Width           =   1100
   End
   Begin VB.ComboBox CBB1 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   680
      Width           =   3285
   End
   Begin VB.TextBox TB1 
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
      Left            =   2160
      TabIndex        =   2
      Text            =   "TB1"
      Top             =   1360
      Width           =   1500
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
      Left            =   8040
      TabIndex        =   4
      Text            =   "TB3"
      Top             =   435
      Width           =   1260
   End
   Begin VB.CheckBox CB4 
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
      Left            =   5640
      TabIndex        =   68
      Top             =   1680
      Width           =   200
   End
   Begin VB.CheckBox CB2 
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
      Left            =   5640
      TabIndex        =   67
      Top             =   1200
      Width           =   200
   End
   Begin VB.CheckBox CB3 
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
      Left            =   5640
      TabIndex        =   66
      Top             =   1440
      Width           =   200
   End
   Begin VB.TextBox TB4 
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
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   65
      Text            =   "TB4"
      Top             =   2040
      Width           =   1260
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
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   64
      Text            =   "TB5"
      Top             =   2340
      Width           =   1260
   End
   Begin VB.CommandButton CMD2 
      BackColor       =   &H8000000C&
      Caption         =   "Masukkan Dalam Senarai Belian"
      Height          =   360
      Left            =   120
      MouseIcon       =   "frm123.frx":1E8D
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   2900
   End
   Begin VB.CommandButton CMD1 
      Caption         =   "Masukkan Dalam Senarai Belian"
      Height          =   360
      Left            =   1320
      MouseIcon       =   "frm123.frx":2197
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2280
      Width           =   3135
   End
   Begin VB.CommandButton CMD3 
      BackColor       =   &H8000000C&
      Caption         =   "Batal Edit Data"
      Height          =   360
      Left            =   3120
      MouseIcon       =   "frm123.frx":24A1
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox TB2 
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
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   3
      Text            =   "TB2"
      Top             =   1660
      Width           =   1500
   End
   Begin VB.ComboBox CBB4 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   360
      ItemData        =   "frm123.frx":27AB
      Left            =   12240
      List            =   "frm123.frx":27AD
      Style           =   2  'Dropdown List
      TabIndex        =   54
      Top             =   7755
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0FF&
      Height          =   2175
      Left            =   19440
      ScaleHeight     =   2115
      ScaleWidth      =   5835
      TabIndex        =   30
      Top             =   4920
      Visible         =   0   'False
      Width           =   5895
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
         TabIndex        =   53
         Top             =   1080
         Visible         =   0   'False
         Width           =   645
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
         TabIndex        =   52
         Top             =   720
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label L47_Text 
         Caption         =   "L47_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3240
         TabIndex        =   51
         Top             =   360
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L45_Text 
         Caption         =   "L45_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3240
         TabIndex        =   50
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L42_Text 
         Caption         =   "L42_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   49
         Top             =   1680
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L41_Text 
         Caption         =   "L41_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   48
         Top             =   1440
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L36_Text 
         Caption         =   "L36_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   47
         Top             =   240
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L37_Text 
         Caption         =   "L37_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   46
         Top             =   480
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L38_Text 
         Caption         =   "L38_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   45
         Top             =   720
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L39_Text 
         Caption         =   "L39_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   44
         Top             =   960
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L35_Text 
         Caption         =   "L35_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L40_Text 
         Caption         =   "L40_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   42
         Top             =   1200
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L33_Text 
         Caption         =   "L33_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         TabIndex        =   41
         Top             =   720
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L32_Text 
         Caption         =   "L32_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         TabIndex        =   40
         Top             =   480
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L30_Text 
         Caption         =   "L30_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         TabIndex        =   39
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L31_Text 
         Caption         =   "L31_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         TabIndex        =   38
         Top             =   240
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L29_Text 
         Caption         =   "L29_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   37
         Top             =   1200
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L2_Text 
         Caption         =   "L2_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L25_Text 
         Caption         =   "L25_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   35
         Top             =   960
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L24_Text 
         Caption         =   "L24_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   34
         Top             =   720
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
      Begin VB.Label L22_Text 
         Caption         =   "L22_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   32
         Top             =   240
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label L54_Text 
         Caption         =   "L54_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4560
         TabIndex        =   31
         Top             =   1200
         Visible         =   0   'False
         Width           =   795
      End
   End
   Begin VB.ComboBox CBB2 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      ItemData        =   "frm123.frx":27AF
      Left            =   12240
      List            =   "frm123.frx":27B1
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   6720
      Width           =   4215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   350
      Left            =   12240
      TabIndex        =   55
      Top             =   8160
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   609
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
      Format          =   415760384
      CurrentDate     =   41561
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   7035
      Left            =   120
      TabIndex        =   88
      ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
      Top             =   2880
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   12409
      _Version        =   393216
      Rows            =   1
      Cols            =   0
      FixedCols       =   0
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   12648384
      ForeColorFixed  =   0
      BackColorSel    =   16777215
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      GridColor       =   0
      WordWrap        =   -1  'True
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
   Begin VB.CommandButton CMD9 
      Caption         =   "Keluar"
      Height          =   360
      Left            =   13200
      MouseIcon       =   "frm123.frx":27B3
      MousePointer    =   99  'Custom
      TabIndex        =   12
      ToolTipText     =   "Sila pastikan semua data telah dimasukkan dengan betul sebelum masukkan data ini ke dalam senarai jualan."
      Top             =   10200
      Width           =   3015
   End
   Begin VB.CommandButton CMD8 
      Caption         =   "Simpan Data"
      Height          =   360
      Left            =   10080
      MouseIcon       =   "frm123.frx":2ABD
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   10200
      Width           =   3015
   End
   Begin VB.CommandButton CMD10 
      BackColor       =   &H8000000C&
      Caption         =   "Simpan Data"
      Height          =   360
      Left            =   10080
      MouseIcon       =   "frm123.frx":2DC7
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   119
      Top             =   10200
      Width           =   3015
   End
   Begin VB.CommandButton CMD11 
      BackColor       =   &H8000000C&
      Caption         =   "Batal"
      Height          =   360
      Left            =   13200
      MouseIcon       =   "frm123.frx":30D1
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   118
      ToolTipText     =   "Sila pastikan semua data telah dimasukkan dengan betul sebelum masukkan data ini ke dalam senarai jualan."
      Top             =   10200
      Width           =   3015
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2040
      TabIndex        =   128
      Top             =   1080
      Width           =   150
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Baki Berat"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   127
      Top             =   1080
      Width           =   1305
   End
   Begin VB.Label L71_Text 
      BackColor       =   &H8000000C&
      Caption         =   "L71_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   10440
      TabIndex        =   125
      Top             =   9000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "No. rujukan bil/GDN dari supplier/agen jika ada."
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   13800
      TabIndex        =   124
      Top             =   7080
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   12135
      TabIndex        =   123
      Top             =   7080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Rujukan"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10320
      TabIndex        =   122
      Top             =   7110
      Visible         =   0   'False
      Width           =   1665
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
      Left            =   13680
      TabIndex        =   121
      Top             =   1230
      Width           =   825
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
      Left            =   10200
      TabIndex        =   120
      Top             =   1230
      Width           =   4185
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga semasa (RM/g)"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10440
      TabIndex        =   99
      Top             =   5190
      Width           =   1905
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "Berikut adalah nilaian harga belian emas bagi purity 999.9 untuk pengiraan belian."
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
      Left            =   10320
      TabIndex        =   98
      Top             =   4680
      Width           =   6345
   End
   Begin VB.Shape Shape3 
      Height          =   945
      Left            =   10200
      Top             =   4560
      Width           =   6735
   End
   Begin VB.Label L70_Text 
      BackColor       =   &H8000000C&
      Caption         =   "L70_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3600
      TabIndex        =   95
      Top             =   10200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label L69_Text 
      BackColor       =   &H8000000C&
      Caption         =   "L69_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3600
      TabIndex        =   94
      Top             =   9960
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
      Left            =   6480
      TabIndex        =   93
      Top             =   9960
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
      Left            =   7080
      TabIndex        =   92
      Top             =   9960
      Width           =   615
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai GDN."
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
      TabIndex        =   89
      Top             =   2640
      Width           =   5385
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah GST"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5640
      TabIndex        =   86
      Top             =   2070
      Width           =   2505
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Upah Dengan GST"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5640
      TabIndex        =   85
      Top             =   2355
      Width           =   2265
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sila masukkan maklumat terperinci barang yang dihantar secara GDN."
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
      Height          =   645
      Left            =   120
      TabIndex        =   84
      Top             =   120
      Width           =   5385
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Berat GDN * (g)"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   83
      Top             =   1390
      Width           =   1665
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2055
      TabIndex        =   80
      Top             =   720
      Width           =   150
   End
   Begin VB.Label L1_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L1_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2145
      TabIndex        =   79
      Top             =   1965
      Width           =   1545
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2055
      TabIndex        =   78
      Top             =   1390
      Width           =   150
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Upah (RM)"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5640
      TabIndex        =   77
      Top             =   435
      Width           =   1665
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7920
      TabIndex        =   76
      Top             =   435
      Width           =   150
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2055
      TabIndex        =   75
      Top             =   1675
      Width           =   150
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2055
      TabIndex        =   74
      Top             =   1965
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
      Left            =   5685
      TabIndex        =   73
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label50 
      BackStyle       =   0  'Transparent
      Caption         =   "Zero Rated ZR(L)           Standard Rated SR      Standard Rated SR (Inclusive)"
      ForeColor       =   &H00000000&
      Height          =   765
      Left            =   5880
      TabIndex        =   72
      Top             =   1200
      Width           =   2730
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "** GST hanya dikenakan kepada upah SAHAJA."
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5640
      TabIndex        =   71
      Top             =   960
      Width           =   4185
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7410
      TabIndex        =   70
      Top             =   2040
      Width           =   600
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RM :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7410
      TabIndex        =   69
      Top             =   2355
      Width           =   600
   End
   Begin VB.Shape Shape2 
      Height          =   2535
      Left            =   5520
      Top             =   240
      Width           =   4335
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
      Left            =   10320
      TabIndex        =   63
      Top             =   5760
      Width           =   2505
   End
   Begin VB.Label L12_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L12_Text"
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
      Left            =   13680
      TabIndex        =   62
      Top             =   5760
      Width           =   3225
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
      Left            =   12960
      TabIndex        =   61
      Top             =   5760
      Width           =   600
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
      Left            =   10320
      TabIndex        =   60
      Top             =   7755
      Width           =   2025
   End
   Begin VB.Label Label110 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tarikh  :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10200
      TabIndex        =   59
      Top             =   8200
      Width           =   2025
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Maklumat supplier / agen."
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10320
      TabIndex        =   58
      Top             =   6360
      Width           =   5865
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Ini adalah nilaian harga emas belian setelah ditukar mutu ke 999.9"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10320
      TabIndex        =   57
      Top             =   5520
      Width           =   5865
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
      Left            =   10320
      TabIndex        =   56
      Top             =   6720
      Width           =   2025
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
      Left            =   10200
      TabIndex        =   29
      Top             =   3480
      Width           =   4185
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Berat 999.9"
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
      Left            =   10200
      TabIndex        =   28
      Top             =   1800
      Width           =   4185
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
      Left            =   13680
      TabIndex        =   27
      Top             =   1800
      Width           =   825
   End
   Begin VB.Label L9_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L9_Text"
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
      Left            =   14520
      TabIndex        =   26
      Top             =   1800
      Width           =   3675
   End
   Begin VB.Label L48_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L48_Text"
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
      Left            =   14520
      TabIndex        =   25
      Top             =   720
      Width           =   3675
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
      Left            =   10200
      TabIndex        =   24
      Top             =   720
      Width           =   4185
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
      Left            =   13680
      TabIndex        =   23
      Top             =   720
      Width           =   825
   End
   Begin VB.Label L43_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L43_Text"
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
      Left            =   14520
      TabIndex        =   22
      Top             =   240
      Width           =   3675
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Bilangan"
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
      Left            =   10200
      TabIndex        =   21
      Top             =   240
      Width           =   4185
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
      Left            =   13680
      TabIndex        =   20
      Top             =   240
      Width           =   825
   End
   Begin VB.Label L53_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L53_Text"
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
      Left            =   14520
      TabIndex        =   19
      Top             =   3480
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
      Left            =   13680
      TabIndex        =   18
      Top             =   3480
      Width           =   825
   End
   Begin VB.Label L52_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L52_Text"
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
      Left            =   14520
      TabIndex        =   17
      Top             =   3000
      Width           =   3675
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
      Left            =   10200
      TabIndex        =   16
      Top             =   3000
      Width           =   4185
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
      Left            =   13680
      TabIndex        =   15
      Top             =   3000
      Width           =   825
   End
   Begin VB.Label L51_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L51_Text"
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
      Left            =   14520
      TabIndex        =   14
      Top             =   2520
      Width           =   3675
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
      Left            =   10200
      TabIndex        =   13
      Top             =   2520
      Width           =   4185
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
      Left            =   13680
      TabIndex        =   0
      Top             =   2520
      Width           =   825
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   10200
      X2              =   16560
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Kadar Tukaran Mutu *"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   82
      Top             =   1675
      Width           =   2025
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Berat 999.9 (g)"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   81
      Top             =   1965
      Width           =   2505
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Purity *"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   120
      TabIndex        =   87
      Top             =   680
      Width           =   2295
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
      Left            =   5160
      TabIndex        =   96
      Top             =   9960
      Width           =   2295
   End
   Begin VB.Menu frm123_pm_menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu frm123_sm_edit_data 
         Caption         =   "Edit data"
      End
      Begin VB.Menu frm123_sm_spacer1 
         Caption         =   "-"
      End
      Begin VB.Menu frm123_sm_remove 
         Caption         =   "Keluarkan dari senarai / Padam"
      End
   End
End
Attribute VB_Name = "frm123"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB2_Click()
'on error resume next
If frm123.CB2 = 1 Then
    frm123.CB3 = 0
    frm123.CB4 = 0
End If

Call Frm123_calc2
End Sub
Private Sub CB3_Click()
'on error resume next
If frm123.CB3 = 1 Then
    frm123.CB2 = 0
    frm123.CB4 = 0
End If

Call Frm123_calc2
End Sub
Private Sub CB4_Click()
'on error resume next
If frm123.CB4 = 1 Then
    frm123.CB3 = 0
    frm123.CB2 = 0
End If

Call Frm123_calc2
End Sub

Private Sub CBB1_Click()
'On Error Resume Next
Dim LM_BERAT_ASAL As Double
Dim LM_BERAT_GUNA As Double

Call frm123_periksa_baki_berat

If frm123.L2_Text <> vbNullString Then
    Call frm123_berat_edit
End If
End Sub

Private Sub CDM13_Click()
'On Error Resume Next
frm123.Pic1.Visible = True
End Sub

Private Sub CMD1_Click()
'On Error Resume Next
Dim Err(15)
Dim LM_KADAR_TUKARAN As Double
Dim Frm123_LM_HARGA_SEMASA As Double
Dim Frm123_LM_BERAT_BELIAN_9999 As Double

Dim LM_BAKI_BERAT As Double
Dim LM_BERAT As Double

LM_BAKI_BERAT = 0
LM_BERAT = 0

Frm123_LM_HARGA_SEMASA = 0
Frm123_LM_BERAT_BELIAN_9999 = 0
LM_KADAR_TUKARAN = 0
x = 0

If frm123.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila buat pilihan [Purity]."
End If
If frm123.TB1 = vbNullString Or (frm123.TB1 <> vbNullString And Not IsNumeric(frm123.TB1)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Berat Asal (g)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If frm123.TB10 = vbNullString Or (frm123.TB10 <> vbNullString And Not IsNumeric(frm123.TB10)) Then
    x = x + 1
    Err(x) = "Tiada maklumat [Baki Berat]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If frm123.TB2 = vbNullString Or (frm123.TB2 <> vbNullString And Not IsNumeric(frm123.TB2)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar Tukaran Mutu]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If frm123.TB6 = vbNullString Or (frm123.TB6 <> vbNullString And Not IsNumeric(frm123.TB6)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Harga Semasa Emas 999.9]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If (frm123.TB2 <> vbNullString And IsNumeric(frm123.TB2)) Then
    
    LM_KADAR_TUKARAN = frm123.TB2
    
    If LM_KADAR_TUKARAN > 1 Then
    
        x = x + 1
        Err(x) = "[Kadar Tukaran Purity 999.9] tidak boleh lebih dari 1.00. Hanya NOMBOR dibenarkan dalam ruangan ini."
        
    End If
    
End If
If frm123.TB3 = vbNullString Or (frm123.TB3 <> vbNullString And Not IsNumeric(frm123.TB3)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Upah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If frm123.CB2 = 0 And frm123.CB3 = 0 And frm123.CB4 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan jenis GST bagi upah."
End If
If frm123.TB4 = vbNullString Or frm123.TB5 = vbNullString Then
    x = x + 1
    Err(x) = "Maklumat berkenaan GST yang tidak sah. Sila keluar dari menu ini dan cuba sekali lagi."
End If

If (frm123.TB1 <> vbNullString And IsNumeric(frm123.TB1)) And (frm123.TB10 <> vbNullString And IsNumeric(frm123.TB10)) Then
    
    LM_BAKI_BERAT = frm123.TB10
    LM_BERAT = frm123.TB1
    
    If LM_BERAT > LM_BAKI_BERAT Then
    
        x = x + 1
        Err(x) = "Berat GDN melebihi baki berat asal."
        
    End If
    
End If

If (frm123.TB1 <> vbNullString And IsNumeric(frm123.TB1)) Then

    LM_BERAT = frm123.TB1
    
    If Format(LM_BERAT, "0.00") = "0.00" Then
    
        x = x + 1
        Err(x) = "Berat GDN yang tidak sah. Berat 0 tidak dibenarkan."
        
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
    Note = "Adakah anda ingin masukkan item ini ke dalam senarai penerimaan barang ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
'### Masukkan Data Ke Dalam Temp Table ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_GRN_TEMP & "", cn, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        If frm123.CBB1 <> vbNullString Then 'Purity
            rs!purity = frm123.CBB1
        Else
            rs!purity = Null
        End If
        If frm123.TB1 <> vbNullString Then 'Berat Asal
            rs!Berat_Asal = Format(frm123.TB1, "0.00")
        Else
            rs!Berat_Asal = Null
        End If
        If frm123.TB2 <> vbNullString Then 'Kadar Tukaran
            rs!kadar_tukaran = frm123.TB2
        Else
            rs!kadar_tukaran = Null
        End If
        If frm123.L1_Text <> vbNullString Then 'Berat Selepas Tukaran
            rs!berat_tukaran_grn = Format(frm123.L1_Text, "0.00")
        Else
            rs!berat_tukaran_grn = Null
        End If
        If frm123.TB3 <> vbNullString Then
            rs!UPAH = Format(frm123.TB3, "0.00") 'Upah (RM)
        Else
            rs!UPAH = Null 'Upah (RM)
        End If
        If frm123.L30_Text <> vbNullString Then 'Harga upah tanpa GST
            rs!harga_tanpa_gst_grn = Format(frm123.L30_Text, "0.00")
        Else
            rs!harga_tanpa_gst_grn = Null
        End If
        If frm123.TB4 <> vbNullString Then 'Jumlah GST
            rs!jumlah_gst = Format(frm123.TB4, "0.00")
        Else
            rs!jumlah_gst = Null
        End If
        If frm123.L22_Text <> vbNullString Then 'Kadar GST
            rs!kadar_gst = Format(frm123.L22_Text, "0.00")
        Else
            rs!kadar_gst = Null
        End If
        If frm123.TB5 <> vbNullString Then 'Jumlah Upah + GST
            rs!harga_dengan_gst_grn = Format(frm123.TB5, "0.00")
        Else
            rs!harga_dengan_gst_grn = Null
        End If
        If frm123.CB2 = 1 Then
            rs!gst_ari_nashi = "ZR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
        ElseIf frm123.CB3 = 1 Or frm123.CB4 = 1 Then
            rs!gst_ari_nashi = "SR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
            If frm123.CB4 = 1 Then 'Jenis Cukai GST SR
                rs!gst_include = 1 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
            Else
                rs!gst_include = 0 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
            End If
        End If
        
        Frm123_LM_HARGA_SEMASA = frm123.TB6 'Harga emas semasa 999.9 (Untuk tujuan belian dari agen)
        Frm123_LM_BERAT_BELIAN_9999 = frm123.L1_Text 'Berat belian dalam purity 999.9
        
        rs!nilaian_harga_emas = Format(Frm123_LM_HARGA_SEMASA * Frm123_LM_BERAT_BELIAN_9999, "0.00")
        rs!Status = 1
        rs.Update
        Frm123_LM_DATA_SAVE = 1
        
        rs.Close
        Set rs = Nothing
'### Masukkan Data Ke Dalam Temp Table ### - End

        If Frm123_LM_DATA_SAVE = 1 Then
        
            GM_NEXT_PREV = 0
            
            frm123.L69_Text = -1 'Titik Pencarian Data
            frm123.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
            frm123.L67_Text = 0 'Paparan Page ke-xxx

            Call Frm123_reset_1
            Call Frm123_Senarai_Belian_Header
            Call Frm123_Senarai_Belian
            Call frm123_periksa_baki_berat
            
            MsgBox "Data telah berjaya dimasukkan ke dalam senarai jualan.", vbInformation, "Info"
            
            frm123.TB1.SetFocus
        End If
    
    End If
    
End If
End Sub
Private Sub CMD10_Click()
'On Error Resume Next
Dim Err(10)
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim frm123_LM_CUKAI_ZR As Double
Dim frm123_LM_CUKAI_SR As Double
Dim frm123_LM_BERAT_ASAL As Double
Dim frm123_LM_BERAT_JUALAN As Double
Dim LM_KADAR_TUKARAN As Double

LM_KADAR_TUKARAN = 0
frm123_LM_KATEGORI = 0
frm123_LM_BERAT_ASAL = 0 'Berat Asal (g)
frm123_LM_BERAT_JUALAN = 0 'Berat Jualan (g)
frm123_LM_CUKAI_ZR = 0 'Jumlah cukai GST ZR
frm123_LM_CUKAI_SR = 0 'Jumlah cukai GST SR

frm123_LM_Flag_SIMPANAN = 0 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai

If frm123.L71_Text = vbNullString Then
    x = x + 1
    Err(x) = "Tiada maklumat ID data yang ingin diedit. Sila keluar dari menu ini dan cuba sekali lagi."
End If
If frm123.L43_Text = "0" Then
    x = x + 1
    Err(x) = "Tiada senarai belian/penerimaan barang."
End If
If frm123.CBB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih supplier/agen yang membuat belian ini."
End If
If frm123.CBB4 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih nama pekerja."
End If
If frm123.TB8 = vbNullString Or (frm123.TB8 <> vbNullString And Not IsNumeric(frm123.TB8)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar Tukaran Purity 999.9]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If (frm123.TB8 <> vbNullString And IsNumeric(frm123.TB8)) Then
    
    LM_KADAR_TUKARAN = frm123.TB8
    
    If LM_KADAR_TUKARAN > 1 Then
    
        x = x + 1
        Err(x) = "[Kadar Tukaran Mutu 999.9] tidak boleh lebih dari 1.00. Hanya NOMBOR dibenarkan dalam ruangan ini."
        
    End If
    
End If
If frm123.TB9 <> vbNullString Then

    If InStr(1, frm123.TB9, "*") <> 0 Or InStr(1, frm123.TB9, "/") <> 0 Or InStr(1, frm123.TB9, "\") <> 0 Or InStr(1, frm123.TB9, "'") <> 0 Then

        x = x + 1
        Err(x) = "No. rujukan dari supplier/agen mengandungi simbol yang tidak sah."
        
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
        
        If frm123.CBB4 <> vbNullString Then
        
            frm123_LM_EMP_NAMA = Split(frm123.CBB4, "  |  ")(0)
            frm123_LM_EMP_NO = Split(frm123.CBB4, "  |  ")(1)
            
        End If
            
'### Masukkan maklumat Good Delivery Note (GRN) ### - Start
        LM_NOW = Now
        LM_TARIKH = DateTime.Date$
        LM_MASA = DateTime.Time$
        LM_NO_RUJUKAN = vbNullString
        
        Dim LM_NO_REVISION As Single
        
        LM_NO_REVISION = 0
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 77_gdn_grn where ID='" & frm123.L71_Text & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            G_ID = rs!ID
            Call recovery_77_gdn_grn
            
            rs!tarikh = frm123.DTPicker1
            rs!masa = LM_MASA
            
            If Not IsNull(rs!no_rujukan) Then LM_NO_RUJUKAN = rs!no_rujukan
            If Not IsNull(rs!Revision) Then LM_NO_REVISION = rs!Revision
            
            If LM_NO_RUJUKAN <> vbNullString Then
                
                LM_NO_REVISION = LM_NO_REVISION + 1
                rs!no_rujukan_rev = LM_NO_RUJUKAN & "-" & LM_NO_REVISION
                LM_NO_RUJUKAN_REV = LM_NO_RUJUKAN & "-" & LM_NO_REVISION
                
                rs!Revision = LM_NO_REVISION
                
            End If
            
            If frm123.L48_Text <> vbNullString Then 'Berat asal sebelum tukaran mutu
                rs!Berat_Asal = Format(frm123.L48_Text, "0.00")
            Else
                rs!Berat_Asal = "0.00"
            End If
            If frm123.TB8 <> vbNullString Then
                rs!kadar_tukaran = frm123.TB8
            Else
                rs!kadar_tukaran = "0.00"
            End If
            If frm123.L9_Text <> vbNullString Then
                rs!berat_tukaran = Format(frm123.L9_Text, "0.00")
            Else
                rs!berat_tukaran = Null
            End If
            If frm123.L51_Text <> vbNullString Then
                rs!harga_tanpa_gst = Format(frm123.L51_Text, "0.00")
            Else
                rs!harga_tanpa_gst = "0.00"
            End If
            If frm123.L52_Text <> vbNullString Then
                rs!jumlah_gst = Format(frm123.L52_Text, "0.00")
            Else
                rs!jumlah_gst = "0.00"
            End If
            If frm123.L22_Text <> vbNullString Then
                rs!kadar_gst = Format(frm123.L22_Text, "0.00")
            Else
                rs!kadar_gst = "0.00"
            End If
            If frm123.L53_Text <> vbNullString Then
                rs!harga_dengan_gst = Format(frm123.L53_Text, "0.00")
            Else
                rs!harga_dengan_gst = Null
            End If
            If frm123.TB6 <> vbNullString Then
                rs!harga_999 = Format(frm123.TB6, "0.00")
            Else
                rs!harga_999 = "0.00"
            End If
            If frm123.L12_Text <> vbNullString Then
                rs!nilaian_harga_emas = Format(frm123.L12_Text, "0.00")
            Else
                rs!nilaian_harga_emas = "0.00"
            End If
            If frm123.L17_Text <> vbNullString Then
                rs!gst_zr_harga = Format(frm123.L17_Text, "0.00")
            Else
                rs!gst_zr_harga = "0.00"
            End If
            If frm123.L18_Text <> vbNullString Then
                rs!gst_sr_harga = Format(frm123.L18_Text, "0.00")
            Else
                rs!gst_sr_harga = "0.00"
            End If
            If frm123.L19_Text <> vbNullString Then
                rs!gst_zr_cukai = Format(frm123.L19_Text, "0.00")
            Else
                rs!gst_zr_cukai = "0.00"
            End If
            If frm123.L20_Text <> vbNullString Then
                rs!gst_sr_cukai = Format(frm123.L20_Text, "0.00")
            Else
                rs!gst_sr_cukai = "0.00"
            End If
            If frm123.L43_Text <> vbNullString Then
                rs!bil_barang = frm123.L43_Text
            Else
                rs!bil_barang = 0
            End If
            If frm123.TB9 <> vbNullString Then
                rs!no_rujukan_supplier = UCase(frm123.TB9)
            Else
                rs!no_rujukan_supplier = Null
            End If
            
            rs!Status = 1
            rs!jenis_urusan = 4
            rs!jenis = "GDN"
            rs!terminal = G_TERMINAL
            If frm123.CBB2 <> vbNullString Then
                rs!supplier_agen = frm123.CBB2
            Else
                rs!supplier_agen = Null
            End If
            rs!user = frm123_LM_EMP_NAMA 'Nama Pekerja
            If Not IsNull(rs!cawangan) Then LM_CAWANGAN = rs!cawangan
            rs.Update
            DATA_SAVE = 1
            
        End If
        
        rs.Close
        Set rs = Nothing
'### Masukkan data voucher / invoice bagi belian agen ini ### - End

'### Masukkan Data Jualan Ke Dalam Table Jualan ### - Start

        If DATA_SAVE = 1 Then
        
            '### Transfer data kepada recovery database ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

            strsql = "insert into " & G_RECOVERY_DATABASE & ".85_penggunaan_ti(id_asal,tarikh,no_rujukan,purity,berat,write_timestamp,terminal,Status,menu)" & _
                        "select ID,tarikh,no_rujukan,purity,berat,write_timestamp,terminal,Status,menu " _
                        & "from " & G_SERVER_DATABASE & ".85_penggunaan_ti WHERE no_rujukan='" & LM_NO_RUJUKAN & "' AND status = 1"
                        
            Set rs = cn.Execute(strsql)
            Set rs = Nothing
            '### Transfer data kepada recovery database ### - End
            
            '### Padam data asal (85_penggunaan_ti) ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

            strsql = "DELETE FROM 85_penggunaan_ti WHERE no_rujukan='" & LM_NO_RUJUKAN & "' AND status = 1"

            Set rs = cn.Execute(strsql)
            Set rs = Nothing
            '### Padam data asal (85_penggunaan_ti) ### - End
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

            strsql = "insert into 85_penggunaan_ti(tarikh,no_rujukan,cawangan,purity,berat,write_timestamp,terminal,Status,menu)" & _
                        "select '" & frm123.DTPicker1 & "','" & G_No_RESIT_JUALAN & "','" & LM_CAWANGAN & "',purity,berat_asal,'" & LM_NOW & "','" & G_TERMINAL & "',1,0 from " & G_GRN_TEMP & " WHERE (status = 1 OR status = 2 OR status = 3 OR status = 4)"

            Set rs = cn.Execute(strsql)
            Set rs = Nothing
            
            '### Transfer data kepada recovery database ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

            strsql = "insert into " & G_RECOVERY_DATABASE & ".79_grn(id_asal,tarikh,masa,write_timestamp,no_rujukan,no_rujukan_rev,purity,berat_asal," _
                        & "berat_tukaran_grn,upah,harga_tanpa_gst_grn,gst_ari_nashi," _
                        & "gst_include,jumlah_gst,kadar_gst,harga_dengan_gst_grn,nilaian_harga_emas," _
                        & "kadar_tukaran,Status,flag_urusan,terminal,user)" & _
                        "select ID,tarikh,masa,write_timestamp,no_rujukan,no_rujukan_rev,purity,berat_asal," _
                        & "berat_tukaran_grn,upah,harga_tanpa_gst_grn,gst_ari_nashi," _
                        & "gst_include,jumlah_gst,kadar_gst,harga_dengan_gst_grn,nilaian_harga_emas," _
                        & "kadar_tukaran,Status,flag_urusan,terminal,user " _
                        & "from " & G_SERVER_DATABASE & ".79_grn WHERE no_rujukan='" & LM_NO_RUJUKAN & "' AND status = 1"
                        
            Set rs = cn.Execute(strsql)
            Set rs = Nothing
            '### Transfer data kepada recovery database ### - End
            
            '### Padam data asal (senarai GRN) ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

            strsql = "DELETE FROM 79_grn WHERE no_rujukan='" & LM_NO_RUJUKAN & "' AND status = 1"

            Set rs = cn.Execute(strsql)
            Set rs = Nothing
            '### Padam data asal (senarai GRN) ### - End
            
            '### Transfer data dari temp table kepada 79_grn ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

            'strsql = "insert into 79_grn SELECT * FROM " & G_GRN_TEMP & " WHERE status = 1"
            
            strsql = "insert into 79_grn(tarikh,masa,write_timestamp,no_rujukan,no_rujukan_rev,purity,berat_asal," _
                        & "berat_tukaran_grn,upah,harga_tanpa_gst_grn,gst_ari_nashi," _
                        & "gst_include,jumlah_gst,kadar_gst,harga_dengan_gst_grn,nilaian_harga_emas," _
                        & "kadar_tukaran,Status,flag_urusan,terminal,cawangan,user)" & _
                        "select '" & frm123.DTPicker1 & "','" & LM_MASA & "','" & LM_NOW & "','" & LM_NO_RUJUKAN & "','" & LM_NO_RUJUKAN_REV & "',purity,berat_asal," _
                        & "berat_tukaran_grn,upah,harga_tanpa_gst_grn,gst_ari_nashi," _
                        & "gst_include,jumlah_gst,kadar_gst,harga_dengan_gst_grn,nilaian_harga_emas," _
                        & "kadar_tukaran,1,0,'" & G_TERMINAL & "','" & LM_CAWANGAN & "','" & frm123_LM_EMP_NAMA & "'" _
                        & "from " & G_GRN_TEMP & " where (status = 1 OR status = 2 OR status = 3 OR status = 4)"

            Set rs = cn.Execute(strsql)
            Set rs = Nothing
            '### Transfer data kepada recovery database ### - End

    '#### Update Log Aktiviti Sistem #### - Start
            'User = MDI_frm1.L3_Text
            LogAct_Memory = "[" & frm123_LM_EMP_NAMA & "] Edit/revise data GDN kepada agen/supplier (bulk). No. Rujukan [" & LM_NO_RUJUKAN & "][" & LM_NO_RUJUKAN_REV & "]."
            LogDate_Memory = LM_NOW
            Call UpdateLog_Database
    '#### Update Log Aktiviti Sistem #### - End
            
            frm117.Show
            Unload frm123
            
            GM_NEXT_PREV = 2
            
            Call frm117_report_gdn_grn_header
            Call frm117_report_gdn_grn

            MsgBox "Data GDN telah berjaya diedit/revise.", vbInformation, "Info"
        
        End If
    
    End If
    
End If
End Sub

Private Sub CMD11_Click()
'on error resume next
Note = "Adakah anda ingin batalkan edit data ini?" & vbCrLf & _
        "Sistem tidak akan menyimpan data jika terdapat data yang diubah." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then
    
    frm117.Show
    Unload frm123

End If
End Sub

Private Sub CMD13_Click()
'On Error Resume Next
frm123.Pic1.Visible = False
End Sub

Private Sub CMD2_Click()
'On Error Resume Next
Dim Err(15)
Dim LM_KADAR_TUKARAN As Double
Dim Frm123_LM_HARGA_SEMASA As Double
Dim Frm123_LM_BERAT_BELIAN_9999 As Double
Dim LM_BAKI_BERAT As Double
Dim LM_BERAT As Double

LM_BAKI_BERAT = 0
LM_BERAT = 0
Frm123_LM_HARGA_SEMASA = 0
Frm123_LM_BERAT_BELIAN_9999 = 0
LM_KADAR_TUKARAN = 0
x = 0

If frm123.L24_Text = vbNullString Or (frm123.L24_Text <> vbNullString And frm123.L24_Text = 0) Then
    x = x + 1
    Err(x) = "Tiada maklumat mengenai ID yang ingin diedit. Sila keluar dari menu ini dan cuba sekali lagi."
End If
If frm123.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila buat pilihan [Purity]."
End If
If frm123.TB1 = vbNullString Or (frm123.TB1 <> vbNullString And Not IsNumeric(frm123.TB1)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Berat Asal (g)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If frm123.TB2 = vbNullString Or (frm123.TB2 <> vbNullString And Not IsNumeric(frm123.TB2)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar Tukaran Mutu]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If frm123.TB6 = vbNullString Or (frm123.TB6 <> vbNullString And Not IsNumeric(frm123.TB6)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Harga Semasa Emas 999.9]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If (frm123.TB2 <> vbNullString And IsNumeric(frm123.TB2)) Then
    
    LM_KADAR_TUKARAN = frm123.TB2
    
    If LM_KADAR_TUKARAN > 1 Then
    
        x = x + 1
        Err(x) = "[Kadar Tukaran Purity 999.9] tidak boleh lebih dari 1.00. Hanya NOMBOR dibenarkan dalam ruangan ini."
        
    End If
    
End If
If frm123.TB3 = vbNullString Or (frm123.TB3 <> vbNullString And Not IsNumeric(frm123.TB3)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Upah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If frm123.CB2 = 0 And frm123.CB3 = 0 And frm123.CB4 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan jenis GST bagi upah."
End If
If frm123.TB4 = vbNullString Or frm123.TB5 = vbNullString Then
    x = x + 1
    Err(x) = "Maklumat berkenaan GST yang tidak sah. Sila keluar dari menu ini dan cuba sekali lagi."
End If
If (frm123.TB1 <> vbNullString And IsNumeric(frm123.TB1)) And (frm123.TB10 <> vbNullString And IsNumeric(frm123.TB10)) Then
    
    LM_BAKI_BERAT = frm123.TB10
    LM_BERAT = frm123.TB1
    
    If LM_BERAT > LM_BAKI_BERAT Then
    
        x = x + 1
        Err(x) = "Berat GDN melebihi baki berat asal."
        
    End If
    
End If

If (frm123.TB1 <> vbNullString And IsNumeric(frm123.TB1)) Then

    LM_BERAT = frm123.TB1
    
    If Format(LM_BERAT, "0.00") = "0.00" Then
    
        x = x + 1
        Err(x) = "Berat GDN yang tidak sah. Berat 0 tidak dibenarkan."
        
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
    Note = "Adakah anda ingin masukkan item ini ke dalam senarai penerimaan barang ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
'### Masukkan Data Ke Dalam Temp Table ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_GRN_TEMP & " where ID='" & frm123.L24_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If frm123.CBB1 <> vbNullString Then 'Purity
                rs!purity = frm123.CBB1
            Else
                rs!purity = Null
            End If
            If frm123.TB1 <> vbNullString Then 'Berat Asal
                rs!Berat_Asal = Format(frm123.TB1, "0.00")
            Else
                rs!Berat_Asal = Null
            End If
            If frm123.TB2 <> vbNullString Then 'Kadar Tukaran
                rs!kadar_tukaran = frm123.TB2
            Else
                rs!kadar_tukaran = Null
            End If
            If frm123.L1_Text <> vbNullString Then 'Berat Selepas Tukaran
                rs!berat_tukaran_grn = Format(frm123.L1_Text, "0.00")
            Else
                rs!berat_tukaran_grn = Null
            End If
            If frm123.TB3 <> vbNullString Then
                rs!UPAH = Format(frm123.TB3, "0.00") 'Upah (RM)
            Else
                rs!UPAH = Null 'Upah (RM)
            End If
            If frm123.L30_Text <> vbNullString Then 'Harga upah tanpa GST
                rs!harga_tanpa_gst_grn = Format(frm123.L30_Text, "0.00")
            Else
                rs!harga_tanpa_gst_grn = Null
            End If
            If frm123.TB4 <> vbNullString Then 'Jumlah GST
                rs!jumlah_gst = Format(frm123.TB4, "0.00")
            Else
                rs!jumlah_gst = Null
            End If
            If frm123.L22_Text <> vbNullString Then 'Kadar GST
                rs!kadar_gst = Format(frm123.L22_Text, "0.00")
            Else
                rs!kadar_gst = Null
            End If
            If frm123.TB5 <> vbNullString Then 'Jumlah Upah + GST
                rs!harga_dengan_gst_grn = Format(frm123.TB5, "0.00")
            Else
                rs!harga_dengan_gst_grn = Null
            End If
            If frm123.CB2 = 1 Then
                rs!gst_ari_nashi = "ZR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
            ElseIf frm123.CB3 = 1 Or frm123.CB4 = 1 Then
                rs!gst_ari_nashi = "SR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                If frm123.CB4 = 1 Then 'Jenis Cukai GST SR
                    rs!gst_include = 1 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                Else
                    rs!gst_include = 0 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                End If
            End If
            
            Frm123_LM_HARGA_SEMASA = frm123.TB6 'Harga emas semasa 999.9 (Untuk tujuan belian dari agen)
            Frm123_LM_BERAT_BELIAN_9999 = frm123.L1_Text 'Berat belian dalam purity 999.9
            
            rs!nilaian_harga_emas = Format(Frm123_LM_HARGA_SEMASA * Frm123_LM_BERAT_BELIAN_9999, "0.00")
            If frm123.L32_Text = "0" Then '0 : Menu Data Baru , 1 : Menu Edit Data
                rs!Status = 1
            ElseIf frm123.L32_Text = "1" Then
                If rs!Status = "2" Then
                    rs!Status = 3
                End If
                If rs!Status = "4" Then
                    rs!Status = 4
                End If
            End If
            rs.Update
            Frm123_LM_DATA_SAVE = 1
        
        End If
        
        rs.Close
        Set rs = Nothing
'### Masukkan Data Ke Dalam Temp Table ### - End

        If Frm123_LM_DATA_SAVE = 1 Then
        
            GM_NEXT_PREV = 2

            Call Frm123_reset_1
            Call Frm123_Senarai_Belian_Header
            Call Frm123_Senarai_Belian
            Call frm123_periksa_baki_berat
            
            MsgBox "Data telah berjaya dimasukkan ke dalam senarai jualan.", vbInformation, "Info"
            
            frm123.TB1.SetFocus
            
        End If
    
    End If
    
End If
End Sub

Private Sub CMD21_Click()
'on error resume next
Dim Frm123_LM_CURR_PAGE As Double
Dim Frm123_LM_TOTAL_PAGE As Double

Frm123_LM_CURR_PAGE = 0
Frm123_LM_TOTAL_PAGE = 0

If frm123.L67_Text <> vbNullString And IsNumeric(frm123.L67_Text) Then
    If frm123.L68_Text <> vbNullString And IsNumeric(frm123.L68_Text) Then
        Frm123_LM_CURR_PAGE = frm123.L67_Text
        Frm123_LM_TOTAL_PAGE = frm123.L68_Text
        
        If Frm123_LM_CURR_PAGE <> 1 And Frm123_LM_CURR_PAGE <> 0 Then
        
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
        
            Call Frm123_Senarai_Belian_Header
            Call Frm123_Senarai_Belian
            
        End If

    End If
End If
End Sub
Private Sub CMD22_Click()
'on error resume next
Dim Frm123_LM_CURR_PAGE As Double
Dim Frm123_LM_TOTAL_PAGE As Double

Frm123_LM_CURR_PAGE = 0
Frm123_LM_TOTAL_PAGE = 0

If frm123.L67_Text <> vbNullString And IsNumeric(frm123.L67_Text) Then
    If frm123.L68_Text <> vbNullString And IsNumeric(frm123.L68_Text) Then
        Frm123_LM_CURR_PAGE = frm123.L67_Text
        Frm123_LM_TOTAL_PAGE = frm123.L68_Text
        
        If Frm123_LM_CURR_PAGE < Frm123_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm123_Senarai_Belian_Header
            Call Frm123_Senarai_Belian
            
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

    Call Frm123_reset_1
    
    frm123.CMD1.Visible = True
    frm123.CMD2.Visible = False
    frm123.CMD3.Visible = False
    
End If
End Sub
Private Sub CMD8_Click()
'On Error Resume Next
Dim Err(10)
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim frm123_LM_CUKAI_ZR As Double
Dim frm123_LM_CUKAI_SR As Double
Dim frm123_LM_BERAT_ASAL As Double
Dim frm123_LM_BERAT_JUALAN As Double
Dim LM_KADAR_TUKARAN As Double

LM_KADAR_TUKARAN = 0
frm123_LM_KATEGORI = 0
frm123_LM_BERAT_ASAL = 0 'Berat Asal (g)
frm123_LM_BERAT_JUALAN = 0 'Berat Jualan (g)
frm123_LM_CUKAI_ZR = 0 'Jumlah cukai GST ZR
frm123_LM_CUKAI_SR = 0 'Jumlah cukai GST SR

frm123_LM_Flag_SIMPANAN = 0 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai

If frm123.L43_Text = "0" Then
    x = x + 1
    Err(x) = "Tiada senarai belian/penerimaan barang."
End If
If frm123.CBB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih supplier/agen yang membuat belian ini."
End If
If frm123.CBB4 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih nama pekerja."
End If
If frm123.TB8 = vbNullString Or (frm123.TB8 <> vbNullString And Not IsNumeric(frm123.TB8)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar Tukaran Purity 999.9]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If (frm123.TB8 <> vbNullString And IsNumeric(frm123.TB8)) Then
    
    LM_KADAR_TUKARAN = frm123.TB8
    
    If LM_KADAR_TUKARAN > 1 Then
    
        x = x + 1
        Err(x) = "[Kadar Tukaran Mutu 999.9] tidak boleh lebih dari 1.00. Hanya NOMBOR dibenarkan dalam ruangan ini."
        
    End If
    
End If
If frm123.TB9 <> vbNullString Then

    If InStr(1, frm123.TB9, "*") <> 0 Or InStr(1, frm123.TB9, "/") <> 0 Or InStr(1, frm123.TB9, "\") <> 0 Or InStr(1, frm123.TB9, "'") <> 0 Then

        x = x + 1
        Err(x) = "No. rujukan dari supplier/agen mengandungi simbol yang tidak sah."
        
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
        
        GoTo aaaa:
        
' ### Periksa No. GDN ### - Start
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
        rs.Open "select * from 77_gdn_grn where no_rujukan='" & "GDN" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_NO_GDN, "000000") & "' AND jenis_urusan = 1", cn, adOpenKeyset, adLockOptimistic
        
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

aaaa:

'---------------------------------------No. Invoice
        LM_NOW = Now
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        rs.Open "select * from 6_gdn", cn2, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        rs!tarikh = frm123.DTPicker1
        rs!terminal = G_TERMINAL
        rs!write_timestamp = LM_NOW
        rs!Status = 1
        rs!nama_staff = MDI_frm1.L3_Text
        rs.Update
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        rs.Open "select * from 6_gdn where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & frm123.DTPicker1 & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic
        
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

        If frm123.CBB4 <> vbNullString Then
        
            frm123_LM_EMP_NAMA = Split(frm123.CBB4, "  |  ")(0)
            frm123_LM_EMP_NO = Split(frm123.CBB4, "  |  ")(1)
            
        End If
            
'### Masukkan maklumat Good Delivery Note (GRN) ### - Start
        'LM_NOW = Now
        LM_TARIKH = DateTime.Date$
        LM_MASA = DateTime.Time$
        
        LM_GDN_RE_GEN = 0
        
Re_gen_no_resit2:

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 77_gdn_grn where no_rujukan='" & "GDN" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_NO_GDN, "000000") & "' AND jenis_urusan = 1", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
        
            rs.AddNew
            rs!tarikh = frm123.DTPicker1
            rs!masa = LM_MASA
            rs!write_timestamp = LM_NOW
            
            rs!no_rujukan = "GDN" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_NO_GDN, "000000")
            rs!Revision = 0
            rs!no_rujukan_rev = "GDN" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_NO_GDN, "000000")
            G_No_RESIT_JUALAN = "GDN" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_NO_GDN, "000000")
            
            If frm123.L48_Text <> vbNullString Then 'Berat asal sebelum tukaran mutu
                rs!Berat_Asal = Format(frm123.L48_Text, "0.00")
            Else
                rs!Berat_Asal = "0.00"
            End If
            If frm123.TB8 <> vbNullString Then
                rs!kadar_tukaran = frm123.TB8
            Else
                rs!kadar_tukaran = "0.00"
            End If
            If frm123.L9_Text <> vbNullString Then
                rs!berat_tukaran = Format(frm123.L9_Text, "0.00")
            Else
                rs!berat_tukaran = Null
            End If
            If frm123.L51_Text <> vbNullString Then
                rs!harga_tanpa_gst = Format(frm123.L51_Text, "0.00")
            Else
                rs!harga_tanpa_gst = "0.00"
            End If
            If frm123.L52_Text <> vbNullString Then
                rs!jumlah_gst = Format(frm123.L52_Text, "0.00")
            Else
                rs!jumlah_gst = "0.00"
            End If
            If frm123.L22_Text <> vbNullString Then
                rs!kadar_gst = Format(frm123.L22_Text, "0.00")
            Else
                rs!kadar_gst = "0.00"
            End If
            If frm123.L53_Text <> vbNullString Then
                rs!harga_dengan_gst = Format(frm123.L53_Text, "0.00")
            Else
                rs!harga_dengan_gst = Null
            End If
            If frm123.TB6 <> vbNullString Then
                rs!harga_999 = Format(frm123.TB6, "0.00")
            Else
                rs!harga_999 = "0.00"
            End If
            If frm123.L12_Text <> vbNullString Then
                rs!nilaian_harga_emas = Format(frm123.L12_Text, "0.00")
            Else
                rs!nilaian_harga_emas = "0.00"
            End If
            If frm123.L17_Text <> vbNullString Then
                rs!gst_zr_harga = Format(frm123.L17_Text, "0.00")
            Else
                rs!gst_zr_harga = "0.00"
            End If
            If frm123.L18_Text <> vbNullString Then
                rs!gst_sr_harga = Format(frm123.L18_Text, "0.00")
            Else
                rs!gst_sr_harga = "0.00"
            End If
            If frm123.L19_Text <> vbNullString Then
                rs!gst_zr_cukai = Format(frm123.L19_Text, "0.00")
            Else
                rs!gst_zr_cukai = "0.00"
            End If
            If frm123.L20_Text <> vbNullString Then
                rs!gst_sr_cukai = Format(frm123.L20_Text, "0.00")
            Else
                rs!gst_sr_cukai = "0.00"
            End If
            If frm123.L43_Text <> vbNullString Then
                rs!bil_barang = frm123.L43_Text
            Else
                rs!bil_barang = 0
            End If
            If frm123.TB9 <> vbNullString Then
                rs!no_rujukan_supplier = UCase(frm123.TB9)
            Else
                rs!no_rujukan_supplier = Null
            End If
            
            rs!Status = 1
            rs!jenis_urusan = 4
            rs!jenis = "GDN"
            rs!terminal = G_TERMINAL
            If frm123.CBB2 <> vbNullString Then
                rs!supplier_agen = frm123.CBB2
            Else
                rs!supplier_agen = Null
            End If
            rs!user = frm123_LM_EMP_NAMA 'Nama Pekerja
            rs!cawangan = G_KEDAI
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

            'Set rs = New ADODB.Recordset
            'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            'rs.Open "select * from default_setting where Default1='" & "Default" & "'", cn, adOpenKeyset, adLockOptimistic
           
            'If Not rs.EOF Then
                
            '    rs!no_gdn = LM_NO_GDN + 1 'No. GRN
            '    rs.Update
            
            'End If
            
            'rs.Close
            'Set rs = Nothing
        
        'End If

'### Masukkan Data Jualan Ke Dalam Table Jualan ### - Start

        If DATA_SAVE = 1 Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

            strsql = "insert into 79_grn(tarikh,masa,write_timestamp,no_rujukan,no_rujukan_rev,purity,berat_asal," _
                        & "berat_tukaran_grn,upah,harga_tanpa_gst_grn,gst_ari_nashi," _
                        & "gst_include,jumlah_gst,kadar_gst,harga_dengan_gst_grn,nilaian_harga_emas," _
                        & "kadar_tukaran,Status,flag_urusan,terminal,user,cawangan)" & _
                        "select '" & frm123.DTPicker1 & "','" & LM_MASA & "','" & LM_NOW & "','" & G_No_RESIT_JUALAN & "','" & G_No_RESIT_JUALAN & "',purity,berat_asal," _
                        & "berat_tukaran_grn,upah,harga_tanpa_gst_grn,gst_ari_nashi," _
                        & "gst_include,jumlah_gst,kadar_gst,harga_dengan_gst_grn,nilaian_harga_emas," _
                        & "kadar_tukaran,1,1,'" & G_TERMINAL & "','" & frm123_LM_EMP_NAMA & "','" & G_KEDAI & "'" _
                        & "from " & G_GRN_TEMP & " WHERE status='" & 1 & "'"

            Set rs = cn.Execute(strsql)
            Set rs = Nothing
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

            strsql = "insert into 85_penggunaan_ti(tarikh,no_rujukan,purity,berat,write_timestamp,terminal,cawangan,Status,menu)" & _
                        "select '" & frm123.DTPicker1 & "','" & G_No_RESIT_JUALAN & "',purity,berat_asal,'" & LM_NOW & "','" & G_TERMINAL & "','" & G_KEDAI & "',1,0 from " & G_GRN_TEMP & " WHERE status='" & 1 & "'"

            Set rs = cn.Execute(strsql)
            Set rs = Nothing

    '#### Update Log Aktiviti Sistem #### - Start
            'User = MDI_frm1.L3_Text
            LogAct_Memory = "[" & frm123_LM_EMP_NAMA & "] Pengeluaran GDN kepada agen/supplier (Bulk). No. Rujukan [" & G_No_RESIT_JUALAN & "]."
            LogDate_Memory = LM_NOW
            Call UpdateLog_Database
    '#### Update Log Aktiviti Sistem #### - End
            
            Call Frm123_one_time_reset
            Call Frm123_reset_1
            Call Frm123_reset_3
            Call Frm123_Senarai_Belian_Header
            Call Frm123_Senarai_Belian
            
            Call frm123_cetak_gdn
        
        End If
    
    End If
    
End If
End Sub

Private Sub CMD9_Click()
'on error resume next
Unload frm123
End Sub

Private Sub Frm123_sm_edit_data_Click()
'on error resume next
DATA_FOUND = 0
Frm123_LM_No_ID = vbNullString

If frm123.MSFlexGrid1 <> vbNullString Then

    If IsNumeric(frm123.MSFlexGrid1) Then
    
        Frm123_LM_No_ID = frm123.MSFlexGrid1.TextMatrix(frm123.MSFlexGrid1, 2) 'No. ID
        
        If Frm123_LM_No_ID <> vbNullString Then
        
            Call Frm123_reset_1 '!! Hati-hati dengan tempat letakkan command ini!!
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from " & G_GRN_TEMP & " where ID='" & Frm123_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                GLOBAL_DISABLE = 1

                If Not IsNull(rs!ID) Then frm123.L24_Text = rs!ID 'No. ID Database
                If Not IsNull(rs!ID) Then frm123.L2_Text = rs!ID 'No. ID Database
                If Not IsNull(rs!Berat_Asal) Then frm123.TB1 = Format(rs!Berat_Asal, "#,##0.00") 'Berat Asal
                If Not IsNull(rs!kadar_tukaran) Then frm123.TB2 = rs!kadar_tukaran 'Kadar Tukaran
                If Not IsNull(rs!berat_tukaran_grn) Then frm123.L1_Text = Format(rs!berat_tukaran_grn, "#,##0.00") 'Berat Selepas Tukaran
                If Not IsNull(rs!kadar_gst) Then frm123.L22_Text = rs!kadar_gst 'Kadar GST
                If Not IsNull(rs!UPAH) Then frm123.TB3 = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
                If Not IsNull(rs!harga_tanpa_gst_grn) Then frm123.L30_Text = Format(rs!harga_tanpa_gst_grn, "#,##0.00") 'Harga upah tanpa GST
                If Not IsNull(rs!jumlah_gst) Then frm123.TB4 = Format(rs!jumlah_gst, "#,##0.00") 'Jumlah GST
                If Not IsNull(rs!harga_dengan_gst_grn) Then frm123.TB5 = Format(rs!harga_dengan_gst_grn, "#,##0.00") 'Jumlah Upah + GST

                If Not IsNull(rs!gst_ari_nashi) Then 'Harga Jualan (RM)
                    If rs!gst_ari_nashi = "ZR" Then
                        frm123.CB2 = 1 '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                    ElseIf rs!gst_ari_nashi = "SR" Then
                        frm123.CB3 = 1 '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                        If Not IsNull(rs!gst_include) Then
                            If rs!gst_include = 0 Then
                                frm123.CB4 = 0
                            ElseIf rs!gst_include = 1 Then
                                frm123.CB4 = 1
                            End If
                        Else
                            frm123.CB4 = 0
                        End If
                    End If
                End If
                If Not IsNull(rs!purity) Then
                
                    LM_SUPPLIER = rs!purity
                    'On Error GoTo Err_A:
                    frm123.CBB1 = LM_SUPPLIER
                    
Restore_A:

                End If
                DATA_FOUND = 1
                GLOBAL_DISABLE = 0
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_FOUND = 1 Then
                frm123.CMD1.Visible = False
                frm123.CMD2.Visible = True
                frm123.CMD3.Visible = True
                
                'If frm123.L71_Text <> vbNullString Then
                    
                    frm123.CBB1.Enabled = False
                    frm123.CBB1.BackColor = &H8000000A
                
                'End If
                
            End If
            
        Else
            
            MsgBox "Tiada data.", vbExclamation, "Info"
            
        End If
        
    Else
    
        MsgBox "Tiada data.", vbExclamation, "Info"
        
    End If
    
End If

Exit Sub
Err_A:
frm123.CBB1.AddItem Frm115_LM_MAKLUMAT_PEKERJA
frm123.CBB1 = Frm115_LM_MAKLUMAT_PEKERJA
Resume Restore_A:
End Sub
Private Sub Frm123_sm_remove_Click()
'on error resume next
DATA_FOUND = 0
Frm123_LM_No_ID = vbNullString

If frm123.MSFlexGrid1 <> vbNullString Then

    If IsNumeric(frm123.MSFlexGrid1) Then
    
        Frm123_LM_No_ID = frm123.MSFlexGrid1.TextMatrix(frm123.MSFlexGrid1, 2) 'No. ID
        
        If Frm123_LM_No_ID <> vbNullString Then
        
            Note = "Adakah anda ingin keluarkan item ini dari senarai penerimaan barang?"
                    
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            If Answer = vbNo Then
                'Exit Sub
            End If
            If Answer = vbYes Then
            
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from " & G_GRN_TEMP & " where ID='" & Frm123_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
                
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
                    
                    Call Frm123_reset_1
                    Call Frm123_Senarai_Belian_Header
                    Call Frm123_Senarai_Belian
                    
                    MsgBox "Item telah dikeluarkan dari senarai penerimaan.", vbInformation, "Info"
                    
                End If
            End If

        Else
            
            MsgBox "Tiada data.", vbExclamation, "Info"
            
        End If
        
    Else
    
        MsgBox "Tiada data.", vbExclamation, "Info"
        
    End If
    
End If
End Sub
Private Sub L30_Text_Change()
'on error resume next
Call Frm123_calc3
End Sub
Private Sub L35_Text_Change()
'On Error Resume Next
Call Frm123_calc10
End Sub
Private Sub L36_Text_Change()
'On Error Resume Next
Call Frm123_calc10
End Sub
Private Sub L37_Text_Change()
'On Error Resume Next
Call Frm123_calc10
End Sub
Private Sub L38_Text_Change()
'On Error Resume Next
Call Frm123_calc10
End Sub
Private Sub L39_Text_Change()
'On Error Resume Next
Call Frm123_calc10
End Sub
Private Sub L40_Text_Change()
'On Error Resume Next
Call Frm123_calc10
End Sub
Private Sub L41_Text_Change()
'On Error Resume Next
Call Frm123_calc10
End Sub
Private Sub L42_Text_Change()
'On Error Resume Next
Call Frm123_calc10
End Sub
Private Sub L48_Text_Change()
'On Error Resume Next
Call Frm123_calc11
End Sub
Private Sub L51_Text_Change()
'On Error Resume Next
Call Frm123_calc10
End Sub
Private Sub L52_Text_Change()
'On Error Resume Next
Call Frm123_calc10
End Sub
Private Sub L9_Text_Change()
'on error resume next
Call Frm123_calc5
End Sub
Private Sub MSFlexGrid1_DblClick()
'On Error Resume Next
Frm123_LM_No_ID = vbNullString

If frm123.MSFlexGrid1 <> vbNullString Then

    If IsNumeric(frm123.MSFlexGrid1) Then
    
        Frm123_LM_No_ID = frm123.MSFlexGrid1.TextMatrix(frm123.MSFlexGrid1, 2) 'No. ID
        
        If Frm123_LM_No_ID <> vbNullString Then
        
            Call Frm123_reset_1
            
            frm123.CMD1.Visible = True
            frm123.CMD2.Visible = False
            frm123.CMD3.Visible = False
    
            PopupMenu frm123_pm_menu
            
        Else
            
            MsgBox "Tiada data.", vbExclamation, "Info"
            
        End If
        
    Else
    
        MsgBox "Tiada data.", vbExclamation, "Info"
        
    End If
    
End If
End Sub

Private Sub TB1_Change()
'on error resume next
Call Frm123_calc1
End Sub
Private Sub TB2_Change()
'on error resume next
Call Frm123_calc1
End Sub
Private Sub TB3_Change()
'on error resume next
Call Frm123_calc2
End Sub

Private Sub TB4_Change()
'on error resume next
Call Frm123_calc3
End Sub

Private Sub TB6_Change()
'on error resume next
Call Frm123_calc5
End Sub

Private Sub TB8_Change()
'On Error Resume Next
Call Frm123_calc11
End Sub


