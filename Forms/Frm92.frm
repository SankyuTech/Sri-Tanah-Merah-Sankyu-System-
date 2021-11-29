VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm92 
   Caption         =   "Servis Kepada Pelanggan"
   ClientHeight    =   13035
   ClientLeft      =   120
   ClientTop       =   -24315
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
   ForeColor       =   &H00000000&
   Icon            =   "Frm92.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13035
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Senarai Perbelanjaan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10935
      Left            =   11280
      TabIndex        =   146
      Top             =   4200
      Visible         =   0   'False
      Width           =   20535
      Begin VB.CommandButton CMD21 
         Caption         =   "Back"
         Height          =   810
         Left            =   18000
         MouseIcon       =   "Frm92.frx":0ECA
         MousePointer    =   99  'Custom
         Picture         =   "Frm92.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   168
         ToolTipText     =   "Tutup senarai ini."
         Top             =   9960
         Width           =   1095
      End
      Begin VB.CommandButton CMD23 
         Caption         =   "Next"
         Height          =   810
         Left            =   19200
         MouseIcon       =   "Frm92.frx":229E
         MousePointer    =   99  'Custom
         Picture         =   "Frm92.frx":25A8
         Style           =   1  'Graphical
         TabIndex        =   167
         ToolTipText     =   "Tutup senarai ini."
         Top             =   9960
         Width           =   1095
      End
      Begin VB.ComboBox CBB9 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   12600
         Style           =   2  'Dropdown List
         TabIndex        =   156
         Top             =   720
         Width           =   4260
      End
      Begin VB.ComboBox CBB8 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   12600
         Style           =   2  'Dropdown List
         TabIndex        =   154
         Top             =   360
         Width           =   4260
      End
      Begin VB.CommandButton CMD20 
         Caption         =   "Carian Senarai Belanja"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   16920
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm92.frx":3672
         MousePointer    =   99  'Custom
         Picture         =   "Frm92.frx":397C
         Style           =   1  'Graphical
         TabIndex        =   153
         Top             =   360
         Width           =   3465
      End
      Begin VB.CheckBox CB17 
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
         Left            =   405
         TabIndex        =   148
         Top             =   525
         Width           =   200
      End
      Begin MSComctlLib.ListView LV3 
         Height          =   7980
         Left            =   120
         TabIndex        =   147
         Top             =   1920
         Width           =   20235
         _ExtentX        =   35692
         _ExtentY        =   14076
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
      Begin MSComCtl2.DTPicker DTPicker5 
         Height          =   360
         Left            =   1560
         TabIndex        =   150
         Top             =   1080
         Width           =   3645
         _ExtentX        =   6429
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
         Format          =   415825920
         CurrentDate     =   41561
      End
      Begin MSComCtl2.DTPicker DTPicker6 
         Height          =   360
         Left            =   6600
         TabIndex        =   152
         Top             =   1080
         Width           =   3645
         _ExtentX        =   6429
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
         Format          =   415825920
         CurrentDate     =   41561
      End
      Begin VB.Label L80_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah   : RM 0.00"
         Height          =   255
         Left            =   240
         TabIndex        =   172
         Top             =   10200
         Width           =   2895
      End
      Begin VB.Label L79_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan : 0"
         Height          =   255
         Left            =   240
         TabIndex        =   171
         Top             =   9960
         Width           =   2895
      End
      Begin VB.Label L83_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L83_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   14520
         TabIndex        =   170
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L82_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L82_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   13560
         TabIndex        =   169
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L75_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L75_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   13080
         TabIndex        =   165
         Top             =   10320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L69_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L69_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   12120
         TabIndex        =   164
         Top             =   10320
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
         Left            =   16920
         TabIndex        =   163
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
         Left            =   17520
         TabIndex        =   162
         Top             =   9960
         Width           =   615
      End
      Begin VB.Label L46_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai perbelanjaan kedai."
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   161
         Top             =   1680
         Width           =   17385
      End
      Begin VB.Label L78_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L78_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   12600
         TabIndex        =   160
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L77_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L77_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   11640
         TabIndex        =   159
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L76_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L76_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   10680
         TabIndex        =   158
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Cawangan *             :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   10560
         TabIndex        =   157
         Top             =   735
         Width           =   2145
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Perbelanjaan * :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   10560
         TabIndex        =   155
         Top             =   375
         Width           =   2145
      End
      Begin VB.Shape Shape2 
         Height          =   1215
         Left            =   240
         Top             =   360
         Width           =   10215
      End
      Begin VB.Label Label62 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Mula :                                                                Tarikh Akhir :"
         Height          =   255
         Left            =   360
         TabIndex        =   151
         Top             =   1125
         Width           =   9015
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm92.frx":4A46
         ForeColor       =   &H00000000&
         Height          =   600
         Left            =   690
         TabIndex        =   149
         Top             =   480
         Width           =   8610
      End
      Begin VB.Label Label22 
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
         Left            =   15600
         TabIndex        =   166
         Top             =   9960
         Width           =   2295
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Perbelanjaan Kedai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10935
      Left            =   6840
      TabIndex        =   121
      Top             =   360
      Visible         =   0   'False
      Width           =   9615
      Begin VB.CommandButton CMD28 
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
         Left            =   3240
         MouseIcon       =   "Frm92.frx":4AD5
         MousePointer    =   99  'Custom
         Picture         =   "Frm92.frx":4DDF
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   9480
         Width           =   2775
      End
      Begin VB.CommandButton CMD30 
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
         Height          =   1095
         Left            =   4680
         MouseIcon       =   "Frm92.frx":73A9
         MousePointer    =   99  'Custom
         Picture         =   "Frm92.frx":76B3
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   9480
         Width           =   2775
      End
      Begin VB.CommandButton CMD29 
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
         Left            =   1800
         MouseIcon       =   "Frm92.frx":9C7D
         MousePointer    =   99  'Custom
         Picture         =   "Frm92.frx":9F87
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   9480
         Width           =   2775
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
         Left            =   240
         TabIndex        =   27
         Top             =   8520
         Width           =   200
      End
      Begin VB.CheckBox CB10 
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
         Left            =   240
         TabIndex        =   25
         Top             =   8040
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
         Left            =   240
         TabIndex        =   26
         Top             =   8280
         Width           =   200
      End
      Begin VB.TextBox TB45 
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   6765
         Locked          =   -1  'True
         TabIndex        =   138
         Text            =   "0.00"
         Top             =   7200
         Width           =   1620
      End
      Begin VB.ComboBox CBB7 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2250
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   5400
         Width           =   6660
      End
      Begin VB.TextBox TB49 
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   6765
         Locked          =   -1  'True
         TabIndex        =   131
         Text            =   "0.00"
         Top             =   6720
         Width           =   1620
      End
      Begin VB.TextBox TB48 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2520
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   6720
         Width           =   1620
      End
      Begin VB.TextBox TB47 
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   6765
         Locked          =   -1  'True
         TabIndex        =   130
         Text            =   "0.00"
         Top             =   6360
         Width           =   1620
      End
      Begin VB.TextBox TB46 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2520
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   6360
         Width           =   1620
      End
      Begin VB.TextBox TB44 
         BackColor       =   &H00FFFFFF&
         Height          =   1200
         Left            =   2250
         MultiLine       =   -1  'True
         TabIndex        =   21
         Text            =   "Frm92.frx":C551
         Top             =   4200
         Width           =   6660
      End
      Begin VB.ComboBox CBB2 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2370
         Style           =   2  'Dropdown List
         TabIndex        =   122
         Top             =   2835
         Width           =   4000
      End
      Begin VB.TextBox TB43 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2370
         TabIndex        =   18
         Top             =   1515
         Width           =   6735
      End
      Begin VB.TextBox TB42 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2370
         TabIndex        =   19
         Top             =   2115
         Width           =   6735
      End
      Begin VB.TextBox TB41 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2370
         TabIndex        =   17
         Top             =   1155
         Width           =   6735
      End
      Begin VB.ComboBox CBB5 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2370
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   795
         Width           =   6765
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   360
         Left            =   2370
         TabIndex        =   20
         Top             =   2475
         Width           =   4005
         _ExtentX        =   7064
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
         Format          =   142409728
         CurrentDate     =   41561
      End
      Begin VB.Label L45_Text 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Caption         =   "L45_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8280
         TabIndex        =   145
         Top             =   3240
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label L43_Text 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "L43_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8160
         TabIndex        =   144
         Top             =   2760
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Shape Shape17 
         Height          =   6975
         Left            =   120
         Top             =   3720
         Width           =   9255
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Tunai        Bank In      Cek"
         ForeColor       =   &H00000000&
         Height          =   885
         Left            =   480
         TabIndex        =   143
         Top             =   7995
         Width           =   945
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Cara Bayaran"
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
         Height          =   405
         Left            =   240
         TabIndex        =   142
         Top             =   7680
         Width           =   4335
      End
      Begin VB.Label L42_Text 
         Caption         =   "6"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4395
         TabIndex        =   141
         Top             =   7200
         Width           =   840
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah  : RM"
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
         Height          =   405
         Left            =   5460
         TabIndex        =   140
         Top             =   7230
         Width           =   1425
      End
      Begin VB.Label Label54 
         BackStyle       =   0  'Transparent
         Caption         =   "** Pengiraan GST adalah berdasarkan kadar @       %"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   139
         Top             =   7200
         Width           =   4440
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Perbelanjaan * :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   137
         Top             =   5420
         Width           =   2145
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Harga ZR *   : RM                             Jumlah Cukai GST ZR * : RM"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   136
         Top             =   6750
         Width           =   6585
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Harga SR *   : RM                             Jumlah Cukai GST SR * : RM"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   135
         Top             =   6390
         Width           =   6585
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Invoice"
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
         Height          =   405
         Left            =   240
         TabIndex        =   134
         Top             =   6000
         Width           =   4335
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Maklumat Tambahan Invoice"
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
         Height          =   405
         Left            =   240
         TabIndex        =   133
         Top             =   3840
         Width           =   4335
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Tujuan Perbelanjaan*:"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   132
         Top             =   4230
         Width           =   2385
      End
      Begin VB.Shape Shape12 
         Height          =   3135
         Left            =   120
         Top             =   360
         Width           =   9255
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh  *                     :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   129
         Top             =   2475
         Width           =   2385
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pekerja *           :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   128
         Top             =   2835
         Width           =   2295
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "No. ID GST *               :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   127
         Top             =   1545
         Width           =   2145
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Invoice *               :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   126
         Top             =   2145
         Width           =   2145
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan maklumat terperinci dari invoice bagi perbelanjaan kedai."
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   125
         Top             =   480
         Width           =   7560
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Kedai *              :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   124
         Top             =   825
         Width           =   2145
      End
      Begin VB.Label L44_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Kedai *              :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   123
         Top             =   1185
         Width           =   2145
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Filter Report Servis"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   2760
      TabIndex        =   81
      Top             =   240
      Visible         =   0   'False
      Width           =   7215
      Begin VB.ComboBox CBB6 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   360
         ItemData        =   "Frm92.frx":C556
         Left            =   1800
         List            =   "Frm92.frx":C558
         Style           =   2  'Dropdown List
         TabIndex        =   115
         Top             =   2880
         Width           =   4965
      End
      Begin VB.CommandButton CMD27 
         Caption         =   "Carian Senarai Servis"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   2400
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm92.frx":C55A
         MousePointer    =   99  'Custom
         Picture         =   "Frm92.frx":C864
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   3360
         Width           =   2865
      End
      Begin VB.ComboBox CBB4 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   360
         ItemData        =   "Frm92.frx":D92E
         Left            =   1800
         List            =   "Frm92.frx":D930
         Style           =   2  'Dropdown List
         TabIndex        =   84
         Top             =   2160
         Width           =   4965
      End
      Begin VB.TextBox TB14 
         Height          =   360
         Left            =   1800
         TabIndex        =   83
         Text            =   "TB14"
         Top             =   2520
         Width           =   4965
      End
      Begin VB.CheckBox CB16 
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
         Left            =   240
         TabIndex        =   82
         Top             =   615
         Width           =   200
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   360
         Left            =   1800
         TabIndex        =   85
         Top             =   1320
         Width           =   4995
         _ExtentX        =   8811
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
         Format          =   142409728
         CurrentDate     =   41561
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   360
         Left            =   1800
         TabIndex        =   86
         Top             =   1680
         Width           =   4995
         _ExtentX        =   8811
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
         Format          =   142409728
         CurrentDate     =   41561
      End
      Begin VB.Label L81_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L81_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         TabIndex        =   117
         Top             =   4200
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Cawangan *"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   116
         Top             =   2895
         Width           =   2295
      End
      Begin VB.Label L71_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L71_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   97
         Top             =   3840
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label L70_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L70_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   96
         Top             =   3480
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label L73_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L73_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         TabIndex        =   95
         Top             =   3480
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label L74_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L74_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         TabIndex        =   94
         Top             =   3840
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label L72_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L72_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   93
         Top             =   4200
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label64 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Mula "
         Height          =   375
         Left            =   240
         TabIndex        =   92
         Top             =   1350
         Width           =   2535
      End
      Begin VB.Label Label67 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm92.frx":D932
         ForeColor       =   &H00000000&
         Height          =   840
         Left            =   480
         TabIndex        =   91
         Top             =   600
         Width           =   6555
      End
      Begin VB.Label Label58 
         BackStyle       =   0  'Transparent
         Caption         =   "Krateria Carian *"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   90
         Top             =   2175
         Width           =   2295
      End
      Begin VB.Label Label57 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila buat pilihan krateria carian senarai servis."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   89
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   7455
      End
      Begin VB.Label L29_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Invoice *"
         Height          =   255
         Left            =   240
         TabIndex        =   88
         Top             =   2565
         UseMnemonic     =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Akhir "
         Height          =   255
         Left            =   240
         TabIndex        =   87
         Top             =   1725
         Width           =   2895
      End
   End
   Begin VB.Timer Tmr1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Senarai Servis"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10935
      Left            =   -240
      TabIndex        =   99
      Top             =   1320
      Visible         =   0   'False
      Width           =   18495
      Begin VB.CommandButton CMD26 
         Caption         =   "Next"
         Height          =   810
         Left            =   17160
         MouseIcon       =   "Frm92.frx":D9DA
         MousePointer    =   99  'Custom
         Picture         =   "Frm92.frx":DCE4
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "Tutup senarai ini."
         Top             =   10080
         Width           =   1095
      End
      Begin VB.CommandButton CMD25 
         Caption         =   "Back"
         Height          =   810
         Left            =   15960
         MouseIcon       =   "Frm92.frx":EDAE
         MousePointer    =   99  'Custom
         Picture         =   "Frm92.frx":F0B8
         Style           =   1  'Graphical
         TabIndex        =   106
         ToolTipText     =   "Tutup senarai ini."
         Top             =   10080
         Width           =   1095
      End
      Begin MSComctlLib.ListView LV2 
         Height          =   9420
         Left            =   120
         TabIndex        =   101
         Top             =   600
         Width           =   18195
         _ExtentX        =   32094
         _ExtentY        =   16616
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
      Begin VB.Label L24_Text 
         BackColor       =   &H00C0C0FF&
         Caption         =   "L24_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   12600
         TabIndex        =   120
         Top             =   10080
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label L25_Text 
         BackColor       =   &H00C0C0FF&
         Caption         =   "L25_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   12600
         TabIndex        =   119
         Top             =   10440
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label L26_Text 
         BackColor       =   &H00C0C0FF&
         Caption         =   "L26_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   12600
         TabIndex        =   118
         Top             =   10800
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label L64_Text 
         BackColor       =   &H008080FF&
         Caption         =   "L64_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7200
         TabIndex        =   114
         Top             =   10320
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label L65_Text 
         BackColor       =   &H008080FF&
         Caption         =   "L65_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7200
         TabIndex        =   113
         Top             =   10560
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label L63_Text 
         BackColor       =   &H008080FF&
         Caption         =   "L63_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5760
         TabIndex        =   112
         Top             =   10560
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label L62_Text 
         BackColor       =   &H008080FF&
         Caption         =   "L62_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5760
         TabIndex        =   111
         Top             =   10320
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label L61_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L61_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   15660
         TabIndex        =   109
         Top             =   10080
         Width           =   705
      End
      Begin VB.Label L60_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L60_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   15075
         TabIndex        =   108
         Top             =   10080
         Width           =   465
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah : RM"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2520
         TabIndex        =   105
         Top             =   10080
         Width           =   1065
      End
      Begin VB.Label L23_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L23_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3690
         TabIndex        =   104
         Top             =   10080
         Width           =   2145
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   103
         Top             =   10080
         Width           =   1065
      End
      Begin VB.Label L22_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L22_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         TabIndex        =   102
         Top             =   10080
         Width           =   945
      End
      Begin VB.Label L21_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai servis kepada pelanggan."
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
         Left            =   120
         TabIndex        =   100
         Top             =   360
         Width           =   12105
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Caption         =   "Paparan Muka :       /"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   13800
         TabIndex        =   110
         Top             =   10080
         Width           =   2505
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Servis Kepada Pelanggan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10935
      Left            =   6720
      TabIndex        =   37
      Top             =   1200
      Visible         =   0   'False
      Width           =   17535
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
         Height          =   1095
         Left            =   11040
         MouseIcon       =   "Frm92.frx":10182
         MousePointer    =   99  'Custom
         Picture         =   "Frm92.frx":1048C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   9600
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CommandButton CMD8 
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
         Height          =   1095
         Left            =   13920
         MouseIcon       =   "Frm92.frx":12A56
         MousePointer    =   99  'Custom
         Picture         =   "Frm92.frx":12D60
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   9600
         Visible         =   0   'False
         Width           =   2775
      End
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
         Height          =   1095
         Left            =   12600
         MouseIcon       =   "Frm92.frx":1532A
         MousePointer    =   99  'Custom
         Picture         =   "Frm92.frx":15634
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   9600
         Width           =   2775
      End
      Begin VB.CommandButton CMD6 
         Caption         =   "Cara Bayaran"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   10320
         Picture         =   "Frm92.frx":17BFE
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1800
         Width           =   3375
      End
      Begin VB.CheckBox CB9 
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
         Left            =   10320
         TabIndex        =   12
         Top             =   4590
         Width           =   200
      End
      Begin VB.ComboBox CBB1 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   360
         Left            =   12720
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   3840
         Width           =   4605
      End
      Begin VB.CommandButton CMD24 
         Caption         =   "Maklumat pelanggan - (Berdaftar)"
         Height          =   330
         Left            =   10320
         MouseIcon       =   "Frm92.frx":1A1C8
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   1440
         Width           =   3375
      End
      Begin VB.CommandButton CMD22 
         Caption         =   "Maklumat pembeli - (Tidak berdaftar)"
         Height          =   330
         Left            =   10320
         MouseIcon       =   "Frm92.frx":1A4D2
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1080
         Width           =   3375
      End
      Begin VB.CommandButton CMD1 
         Caption         =   "Masukkan Dalam Senarai"
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
         MouseIcon       =   "Frm92.frx":1A7DC
         MousePointer    =   99  'Custom
         Picture         =   "Frm92.frx":1AAE6
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2880
         Width           =   2895
      End
      Begin VB.Frame Frame2 
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
         Height          =   1240
         Left            =   240
         TabIndex        =   61
         Top             =   9600
         Width           =   6255
         Begin VB.Label L12_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3720
            TabIndex        =   68
            Top             =   480
            Width           =   1680
         End
         Begin VB.Label L14_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3720
            TabIndex        =   67
            Top             =   720
            Width           =   1680
         End
         Begin VB.Label L11_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1800
            TabIndex        =   66
            Top             =   480
            Width           =   2040
         End
         Begin VB.Label L13_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1800
            TabIndex        =   65
            Top             =   720
            Width           =   2040
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Harga (RM)    Cukai GST (RM)"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1920
            TabIndex        =   64
            Top             =   240
            Width           =   3480
         End
         Begin VB.Label Label122 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Zero Rated :"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   0
            TabIndex        =   63
            Top             =   480
            Width           =   1920
         End
         Begin VB.Label Label65 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Standard Rated SR:"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   0
            TabIndex        =   62
            Top             =   720
            Width           =   1920
         End
         Begin VB.Label L15_Text 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "6"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3720
            TabIndex        =   70
            Top             =   960
            Width           =   840
         End
         Begin VB.Label Label59 
            BackStyle       =   0  'Transparent
            Caption         =   "Pengiraan GST adalah berdasarkan kadar @       %"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            TabIndex        =   69
            Top             =   960
            Width           =   4440
         End
      End
      Begin VB.CommandButton CMD3 
         BackColor       =   &H8000000A&
         Caption         =   "Batal Edit Data"
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
         Left            =   5160
         MouseIcon       =   "Frm92.frx":1B490
         MousePointer    =   99  'Custom
         Picture         =   "Frm92.frx":1B79A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2880
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CommandButton CMD2 
         BackColor       =   &H8000000A&
         Caption         =   "Masukkan Dalam Senarai"
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
         Left            =   2160
         MouseIcon       =   "Frm92.frx":1C864
         MousePointer    =   99  'Custom
         Picture         =   "Frm92.frx":1CB6E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2880
         Visible         =   0   'False
         Width           =   2895
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
         Left            =   4275
         TabIndex        =   3
         Top             =   1600
         Width           =   200
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
         Left            =   2280
         TabIndex        =   2
         Top             =   1600
         Width           =   200
      End
      Begin VB.CheckBox CB8 
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
         Left            =   6540
         TabIndex        =   4
         Top             =   1600
         Width           =   200
      End
      Begin VB.TextBox TB1 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1725
         TabIndex        =   0
         Top             =   720
         Width           =   8460
      End
      Begin VB.TextBox TB2 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1725
         TabIndex        =   1
         Top             =   1065
         Width           =   1620
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   4980
         Left            =   120
         TabIndex        =   51
         Top             =   4320
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   8784
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
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   12720
         TabIndex        =   11
         Top             =   3480
         Width           =   4605
         _ExtentX        =   8123
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
         Format          =   142475264
         CurrentDate     =   41561
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm92.frx":1D518
         ForeColor       =   &H000000FF&
         Height          =   780
         Left            =   10605
         TabIndex        =   80
         Top             =   4560
         Width           =   6330
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Pilihan jenis invoice."
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
         Left            =   10320
         TabIndex        =   79
         Top             =   4320
         Width           =   2295
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh  *"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   10320
         TabIndex        =   78
         Top             =   3480
         Width           =   2385
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pekerja *"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   10320
         TabIndex        =   77
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "Maklumat Pelanggan"
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
         Height          =   405
         Left            =   10320
         TabIndex        =   75
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label L52_Text 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   14520
         TabIndex        =   74
         Top             =   1485
         Width           =   5505
      End
      Begin VB.Label Label75 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama  :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   13800
         TabIndex        =   73
         Top             =   1485
         Width           =   825
      End
      Begin VB.Label L51_Text 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   14520
         TabIndex        =   72
         Top             =   1110
         Width           =   5505
      End
      Begin VB.Label Label52 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama  :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   13800
         TabIndex        =   71
         Top             =   1110
         Width           =   825
      End
      Begin VB.Label L16_Text 
         Alignment       =   2  'Center
         Caption         =   "L16_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9000
         TabIndex        =   60
         Top             =   2880
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label L17_Text 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "L17_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9000
         TabIndex        =   59
         Top             =   3240
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label L53_Text 
         Alignment       =   2  'Center
         Caption         =   "L53_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9000
         TabIndex        =   58
         Top             =   3960
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label L28_Text 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "L28_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9000
         TabIndex        =   57
         Top             =   3600
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label L9_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3480
         TabIndex        =   55
         Top             =   9360
         Width           =   1545
      End
      Begin VB.Label L10_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9240
         TabIndex        =   54
         Top             =   9360
         Width           =   2025
      End
      Begin VB.Label L20_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   600
         TabIndex        =   53
         Top             =   9360
         Width           =   945
      End
      Begin VB.Label L7_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6075
         TabIndex        =   52
         Top             =   9360
         Width           =   1545
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai servis yang diberikan kepada pelanggan ini."
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   50
         Top             =   3960
         Width           =   7560
      End
      Begin VB.Label L18_Text 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm92.frx":1D603
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   795
         Left            =   3600
         TabIndex        =   49
         Top             =   1920
         Visible         =   0   'False
         Width           =   6315
      End
      Begin VB.Label L54_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L54_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9120
         TabIndex        =   48
         Top             =   1560
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Dengan GST : RM"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   47
         Top             =   2400
         Width           =   2265
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah GST : RM"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   46
         Top             =   2160
         Width           =   2265
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Tanpa GST : RM"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   45
         Top             =   1920
         Width           =   2265
      End
      Begin VB.Shape Shape1 
         Height          =   1335
         Left            =   120
         Top             =   1440
         Width           =   9975
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "Maklumat GST          :      Zero Rated ZR             Standard Rated SR          Standard  Rated Inclusive"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   240
         TabIndex        =   44
         Top             =   1560
         Width           =   9105
      End
      Begin VB.Label L8_Text 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   43
         Top             =   2160
         Width           =   1995
      End
      Begin VB.Label L50_Text 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   42
         Top             =   1920
         Width           =   1995
      End
      Begin VB.Label L55_Text 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   41
         Top             =   2400
         Width           =   1995
      End
      Begin VB.Label Label91 
         BackStyle       =   0  'Transparent
         Caption         =   "Maklumat Servis :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   40
         Top             =   750
         Width           =   1905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah         RM :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   39
         Top             =   1110
         Width           =   1905
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan maklumat terperinci servis yang telah diberikan kepada pelanggan."
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
         TabIndex        =   38
         Top             =   360
         Width           =   7560
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Bil :            Jumlah Tanpa GST RM :                   Jumlah GST RM :                 Jumlah Dengan GST RM :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   56
         Top             =   9360
         Width           =   10185
      End
   End
   Begin VB.Label L6_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai Perbelanjaan Kedai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8400
      MouseIcon       =   "Frm92.frx":1D6AC
      MousePointer    =   99  'Custom
      TabIndex        =   36
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label L5_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Perbelanjaan Kedai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      MouseIcon       =   "Frm92.frx":1D9B6
      MousePointer    =   99  'Custom
      TabIndex        =   35
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label L4_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai Servis Kepada Pelanggan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      MouseIcon       =   "Frm92.frx":1DCC0
      MousePointer    =   99  'Custom
      TabIndex        =   34
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label L3_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Servis Kepada Pelanggan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      MouseIcon       =   "Frm92.frx":1DFCA
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Top             =   0
      Width           =   2535
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
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   19185
      TabIndex        =   32
      Top             =   435
      Visible         =   0   'False
      Width           =   2100
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
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   19200
      TabIndex        =   31
      Top             =   120
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Menu Frm92_PM_Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm92_SM_edit 
         Caption         =   "Edit data"
      End
      Begin VB.Menu frm92_sm_bar1 
         Caption         =   "-"
      End
      Begin VB.Menu Frm92_SM_padam 
         Caption         =   "Padam / Keluarkan dari senarai"
      End
   End
   Begin VB.Menu Frm92_PM_Menu2 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm92_SM_edit2 
         Caption         =   "Lihat / Edit Data"
      End
      Begin VB.Menu frm92_sm_bar2 
         Caption         =   "-"
      End
      Begin VB.Menu Frm92_SM_padam2 
         Caption         =   "Padam Data Servis"
      End
      Begin VB.Menu frm92_sm_bar3 
         Caption         =   "-"
      End
      Begin VB.Menu Frm92_SM_cetak_resit 
         Caption         =   "Cetak Invoice"
      End
      Begin VB.Menu frm92_sm_excel 
         Caption         =   "Report Excel"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Frm92_PM_Menu3 
      Caption         =   "Belanja"
      Visible         =   0   'False
      Begin VB.Menu frm92_sm_cetak_pv 
         Caption         =   "Cetak Payment Voucher"
      End
      Begin VB.Menu frm92_sm_excel1 
         Caption         =   "Report Excel"
      End
      Begin VB.Menu frm92_sm_bar6 
         Caption         =   "-"
      End
      Begin VB.Menu Frm92_SM_edit3 
         Caption         =   "Edit Data"
      End
      Begin VB.Menu frm92_sm_bar4 
         Caption         =   "-"
      End
      Begin VB.Menu Frm92_SM_padam3 
         Caption         =   "Padam / Keluarkan Dari Senarai"
      End
   End
   Begin VB.Menu Frm92_PM_Menu4 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm92_SM_edit4 
         Caption         =   "Edit Data"
      End
      Begin VB.Menu frm92_sm_bar5 
         Caption         =   "-"
      End
      Begin VB.Menu Frm92_SM_padam4 
         Caption         =   "Padam / Keluarkan Dari Senarai"
      End
   End
End
Attribute VB_Name = "Frm92"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CB1_Click()
'on error resume next
If Frm92.CB1 = 1 Then
    Frm92.CB2 = 0
    Frm92.CB8 = 0
End If

Call frm92_kiraan_gst
End Sub

Private Sub CB10_Click()
'on error resume next
If Frm92.CB10 = 1 Then
    Frm92.CB3 = 0
    Frm92.CB4 = 0
End If
End Sub

Private Sub CB2_Click()
'on error resume next
If Frm92.CB2 = 1 Then
    Frm92.CB1 = 0
    Frm92.CB8 = 0
End If

Call frm92_kiraan_gst
End Sub

Private Sub CB3_Click()
'on error resume next
If Frm92.CB3 = 1 Then
    Frm92.CB10 = 0
    Frm92.CB4 = 0
End If
End Sub

Private Sub CB4_Click()
'on error resume next
If Frm92.CB4 = 1 Then
    Frm92.CB3 = 0
    Frm92.CB10 = 0
End If
End Sub




Private Sub CB8_Click()
'on error resume next
If Frm92.CB8 = 1 Then
    Frm92.CB1 = 0
    Frm92.CB2 = 0
End If

Call frm92_kiraan_gst
End Sub

Private Sub CBB4_Change()
'On Error Resume Next
If Frm92.CBB4 = "No. invoice" Then
    Frm92.TB14 = vbNullString
    Frm92.TB14.Visible = True
    Frm92.L29_Text.Visible = True
    Frm92.TB14.SetFocus
Else
    Frm92.TB14 = vbNullString
    Frm92.TB14.Visible = False
    Frm92.L29_Text.Visible = False
End If
End Sub

Private Sub CBB4_Click()
'On Error Resume Next
If Frm92.CBB4 = "No. invoice" Then
    Frm92.TB14 = vbNullString
    Frm92.TB14.Visible = True
    Frm92.L29_Text.Visible = True
    Frm92.TB14.SetFocus
Else
    Frm92.TB14 = vbNullString
    Frm92.TB14.Visible = False
    Frm92.L29_Text.Visible = False
End If
End Sub



Private Sub CBB5_Click()
'On Error Resume Next
If Frm92.CBB5 = "Lain-lain" Then
    
    Frm92.L44_Text.Visible = True
    Frm92.TB41.Visible = True
    Frm92.TB43 = vbNullString
    
    'Frm92.TB41 = vbNullString
    If Frm92.Frame5.Visible = True Then Frm92.TB41.SetFocus
    
    Frm92.TB43.Locked = False
    Frm92.TB43.BackColor = &HFFFFFF
    
Else
    
    'Frm92.TB41 = vbNullString
    Frm92.L44_Text.Visible = False
    Frm92.TB41.Visible = False
    
    Frm92.TB43.Locked = True
    Frm92.TB43.BackColor = &H8000000A
    
End If

If Frm92.L43_Text = "" Then
    
    If Frm92.CBB5 <> vbNullString Then
    
        Set rs2 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs2.Open "select * from setting_database where supplier='" & Frm92.CBB5 & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
    
        If Not rs2.EOF Then
            
            If Not IsNull(rs2!no_id_gst) Then Frm92.TB43 = rs2!no_id_gst
        
        End If
        
        rs2.Close
        Set rs2 = Nothing
    
    Else
    
        Frm92.TB43 = vbNullString
    
    End If
    
End If
End Sub

Private Sub CMD1_Click()
'On Error Resume Next
Dim Err(5)
x = 0
DATA_SAVE = 0

If Frm92.TB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan maklumat servis."
End If
If Frm92.TB2 = vbNullString Or (Frm92.TB2 <> vbNullString And Not IsNumeric(Frm92.TB2)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Jumlah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm92.CB1 = 0 And Frm92.CB2 = 0 And Frm92.CB8 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan cukai GST."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Masukkan ke dalam senarai servis ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_SERVICE_TEMP & "", cn, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        If Frm92.TB1 <> vbNullString Then 'Details
            rs!Detail = Frm92.TB1
        Else
            rs!Detail = vbNullString
        End If
        
        If Frm92.TB2 <> vbNullString Then 'Jumlah
            rs!jumlah = Format(Frm92.TB2, "0.00")
        Else
            rs!jumlah = Format(0, "0.00")
        End If
        
        If Frm92.CB1 = 1 Then 'Jenis GST
            rs!jenis_gst = "ZR(L)"
            rs!kod_gst = 0
        ElseIf Frm92.CB2 = 1 Then
            rs!jenis_gst = "SR"
            rs!kod_gst = 1
        ElseIf Frm92.CB8 = 1 Then
            rs!jenis_gst = "SR"
            rs!kod_gst = 2
        End If
        If Frm92.L50_Text <> vbNullString Then
            rs!harga_tanpa_gst = Format(Frm92.L50_Text, "0.00") 'Harga Keseluruhan Tanpa GST (RM)
        Else
            rs!harga_tanpa_gst = "0.00" 'Harga Keseluruhan Tanpa GST (RM)
        End If
        If Frm92.L8_Text <> vbNullString Then
            rs!jumlah_gst = Format(Frm92.L8_Text, "0.00") 'Jumlah cukai GST (RM)
        Else
            rs!jumlah_gst = "0.00" 'Jumlah cukai GST (RM)
        End If
        If Frm92.L55_Text <> vbNullString Then
            rs!harga_dengan_gst = Format(Frm92.L55_Text, "0.00") 'Harga keseluruhan dengan GST (RM)
        Else
            rs!harga_dengan_gst = "0.00" 'Harga keseluruhan dengan GST (RM)
        End If
        DATA_SAVE = 1
        rs.Update
        
        rs.Close
        Set rs = Nothing
        
        Frm92.TB1 = vbNullString
        Frm92.TB2 = "0.00"
        Frm92.L18_Text.Visible = False
        
        Call frm92_senarai_service_header
        Call frm92_senarai_service
        
        If DATA_SAVE = 1 Then MsgBox "Senarai servis telah berjaya diupdate.", vbInformation, "Info"
        
        Frm92.TB1.SetFocus
        
    End If
End If
End Sub



Private Sub CMD2_Click()
'On Error Resume Next
Dim Err(5)
x = 0
DATA_SAVE = 0

If Frm92.TB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan maklumat servis."
End If
If Frm92.TB2 = vbNullString Or (Frm92.TB2 <> vbNullString And Not IsNumeric(Frm92.TB2)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Jumlah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm92.CB1 = 0 And Frm92.CB2 = 0 And Frm92.CB8 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan cukai GST."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Masukkan ke dalam senarai servis ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_SERVICE_TEMP & " where ID='" & Frm92.L16_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            rs!Detail = Frm92.TB1 'Details
            rs!jumlah = Format(Frm92.TB2, "0.00") 'Jumlah
            If Frm92.CB1 = 1 Then 'Jenis GST
                rs!jenis_gst = "ZR(L)"
                rs!kod_gst = 0
            ElseIf Frm92.CB2 = 1 Then
                rs!jenis_gst = "SR"
                rs!kod_gst = 1
            ElseIf Frm92.CB8 = 1 Then
                rs!jenis_gst = "SR"
                rs!kod_gst = 2
            End If
            If Frm92.L50_Text <> vbNullString Then
                rs!harga_tanpa_gst = Format(Frm92.L50_Text, "0.00") 'Harga Keseluruhan Tanpa GST (RM)
            Else
                rs!harga_tanpa_gst = "0.00" 'Harga Keseluruhan Tanpa GST (RM)
            End If
            If Frm92.L8_Text <> vbNullString Then
                rs!jumlah_gst = Format(Frm92.L8_Text, "0.00") 'Jumlah cukai GST (RM)
            Else
                rs!jumlah_gst = "0.00" 'Jumlah cukai GST (RM)
            End If
            If Frm92.L55_Text <> vbNullString Then
                rs!harga_dengan_gst = Format(Frm92.L55_Text, "0.00") 'Harga keseluruhan dengan GST (RM)
            Else
                rs!harga_dengan_gst = "0.00" 'Harga keseluruhan dengan GST (RM)
            End If
            rs.Update
        End If
        
        rs.Close
        Set rs = Nothing
        
        Frm92.TB1 = vbNullString
        Frm92.TB2 = "0.00"
        
        Frm92.TB1 = vbNullString
        Frm92.TB2 = "0.00"
        Frm92.CMD1.Visible = True
        Frm92.CMD2.Visible = False
        Frm92.CMD3.Visible = False
        Frm92.L16_Text = 0
        Frm92.L18_Text.Visible = False
        
        Call frm92_senarai_service_header
        Call frm92_senarai_service
    End If
End If
End Sub

Private Sub CMD20_Click()
'on error resume next
If Frm92.CB17 = 1 Then
    Frm92.L78_Text = 1 '0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh
Else
    Frm92.L78_Text = 0 '0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh
End If

Frm92.L76_Text = Frm92.DTPicker5 'Tarikh mula
Frm92.L77_Text = Frm92.DTPicker6 'Tarikh akhir

Frm92.L69_Text = -1 'Titik Pencarian Data
Frm92.L75_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Frm92.L67_Text = 0 'Paparan Page ke-xxx
Frm92.L68_Text = 0

Frm92.L82_Text = Frm92.CBB8
Frm92.L83_Text = Frm92.CBB9

GM_NEXT_PREV = 0

Call Frm92_report_expenses_header
Call Frm92_report_expenses
End Sub

Private Sub CMD21_Click()
'on error resume next
Dim frm92_LM_CURR_PAGE As Double
Dim frm92_LM_TOTAL_PAGE As Double

frm92_LM_CURR_PAGE = 0
frm92_LM_TOTAL_PAGE = 0

If Frm92.L67_Text <> vbNullString And IsNumeric(Frm92.L67_Text) Then
    If Frm92.L68_Text <> vbNullString And IsNumeric(Frm92.L68_Text) Then
        frm92_LM_CURR_PAGE = Frm92.L67_Text
        frm92_LM_TOTAL_PAGE = Frm92.L68_Text
        
        If frm92_LM_CURR_PAGE <> 1 And frm92_LM_CURR_PAGE <> 0 Then
        
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
                    
            Call Frm92_report_expenses_header
            Call Frm92_report_expenses
            
        End If

    End If
End If
End Sub
Private Sub CMD23_Click()
'on error resume next
Dim frm92_LM_CURR_PAGE As Double
Dim frm92_LM_TOTAL_PAGE As Double

frm92_LM_CURR_PAGE = 0
frm92_LM_TOTAL_PAGE = 0

If Frm92.L67_Text <> vbNullString And IsNumeric(Frm92.L67_Text) Then
    If Frm92.L68_Text <> vbNullString And IsNumeric(Frm92.L68_Text) Then
        frm92_LM_CURR_PAGE = Frm92.L67_Text
        frm92_LM_TOTAL_PAGE = Frm92.L68_Text
        
        If frm92_LM_CURR_PAGE < frm92_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm92_report_expenses_header
            Call Frm92_report_expenses
            
        End If
    End If
End If
End Sub

Private Sub CMD22_Click()
'On Error Resume Next
If Frm92.L51_Text = vbNullString Then
    
    If Frm92.L52_Text <> vbNullString Then
    
        Note = "Anda telah mengisi data pembeli barangan ini di dalam ruangan pelanggan yang berdaftar dengan kedai." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Jika ada meneruskan menu ini , semua data pembeli di dalam ruangan pelanggan berdaftar akan dipadamkan." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
        
            Unload Frm27
            Unload Frm28
            Call Frm26_initial
            
            Frm92.L52_Text = vbNullString 'Nama pembeli : Berdaftar
            
            Frm26.Show 1
        End If
        
    Else
    
        Unload Frm27
        Unload Frm28
        Call Frm26_initial
        
        Frm26.Show 1
                
    End If
    
Else

    Frm26.Show 1
    
End If
End Sub
Private Sub CMD24_Click()
'On Error Resume Next
If Frm92.L52_Text = vbNullString Then
    
    If Frm92.L51_Text <> vbNullString Then
    
        Note = "Anda telah mengisi data pembeli barangan ini di dalam ruangan pelanggan yang TIDAK berdaftar dengan kedai." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Jika ada meneruskan menu ini , semua data pembeli di dalam ruangan pelanggan TIDAK berdaftar akan dipadamkan." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
            
            Unload Frm26
            Unload Frm27
            Call Frm28_initial
            
            Frm92.L51_Text = vbNullString 'Nama pembeli : Tidak berdaftar
            
            Frm28.Show 1
        End If
        
    Else
        
        Unload Frm26
        Unload Frm27
        Call Frm28_initial
        
        Frm28.Show 1
                
    End If
    
Else

    Frm28.Show 1
    
End If
End Sub


Private Sub CMD25_Click()
'on error resume next
Dim frm92_LM_CURR_PAGE As Double
Dim frm92_LM_TOTAL_PAGE As Double

frm92_LM_CURR_PAGE = 0
frm92_LM_TOTAL_PAGE = 0

If Frm92.L60_Text <> vbNullString And IsNumeric(Frm92.L60_Text) Then
    If Frm92.L61_Text <> vbNullString And IsNumeric(Frm92.L61_Text) Then
        frm92_LM_CURR_PAGE = Frm92.L60_Text
        frm92_LM_TOTAL_PAGE = Frm92.L61_Text
        
        If frm92_LM_CURR_PAGE <> 1 And frm92_LM_CURR_PAGE <> 0 Then
        
        GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
        
        Call frm92_senarai_servis_header
        Call frm92_senarai_servis
            
        End If
    End If
End If
End Sub

Private Sub CMD26_Click()
'on error resume next
Dim frm92_LM_CURR_PAGE As Double
Dim frm92_LM_TOTAL_PAGE As Double

frm92_LM_CURR_PAGE = 0
frm92_LM_TOTAL_PAGE = 0

If Frm92.L60_Text <> vbNullString And IsNumeric(Frm92.L60_Text) Then
    If Frm92.L61_Text <> vbNullString And IsNumeric(Frm92.L61_Text) Then
        frm92_LM_CURR_PAGE = Frm92.L60_Text
        frm92_LM_TOTAL_PAGE = Frm92.L61_Text
        
        If frm92_LM_CURR_PAGE < frm92_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call frm92_senarai_servis_header
            Call frm92_senarai_servis
            
        End If
    End If
End If
End Sub

Private Sub CMD27_Click()
'on error resume next
If Frm92.CBB4 = vbNullString Then
    MsgBox "Sila buat pilihan krateria carian.", vbExclamation, "Info"
    
    Exit Sub
End If

If Frm92.TB14.Visible = True Then

    If Frm92.TB14 = vbNullString Then
        MsgBox "Sila masukkan No. invoice", vbExclamation, "Info"
        
        Frm92.TB14.SetFocus
        Exit Sub
    End If
    
    If InStr(1, Frm92.TB14, "*") <> 0 Or InStr(1, Frm92.TB14, "/") <> 0 Or InStr(1, Frm92.TB14, "\") <> 0 Or InStr(1, Frm92.TB14, "'") <> 0 Then
    
        MsgBox "No. invoice mengandungi simbol yang tidak sah.", vbExclamation, "Info"
        
        Frm92.TB14.SetFocus
        Exit Sub
    End If
    
End If

If Frm92.CB16 = 0 Then
    Frm92.L70_Text = 0 '0 : Tiada carian mengikut tarikh , 1 : Carian mengikut tarikh
Else
    Frm92.L70_Text = 1 '0 : Tiada carian mengikut tarikh , 1 : Carian mengikut tarikh
    Frm92.L71_Text = Frm92.DTPicker2 'Tarik mula
    Frm92.L72_Text = Frm92.DTPicker3 'Tarikh akhir
End If
Frm92.L81_Text = Frm92.CBB6


Frm92.L73_Text = Frm92.CBB4 'Krateria carian
Frm92.L74_Text = UCase(Frm92.TB14) 'No. invoice

Frm92.L62_Text = -1 'Start Point
Frm92.L60_Text = 0 'Current Page
Frm92.L63_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
GM_NEXT_PREV = 0

Call frm92_senarai_servis_header
Call frm92_senarai_servis
End Sub

Private Sub CMD28_Click()
'On Error Resume Next
Dim Err(15)
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim LM_HARGA_SR As Double
Dim LM_CUKAI_SR As Double
Dim LM_HARGA_ZR As Double
Dim LM_CUKAI_ZR As Double
        
DATA_SAVE = 0
x = 0

If Frm92.CB10 = 0 And Frm92.CB3 = 0 And Frm92.CB4 = 0 Then
    x = x + 1
    Err(x) = "Sila buat cara bayaran dibuat."
End If
If Frm92.CBB5 = "Lain-lain" Then
    If Frm92.TB41 = vbNullString Then
        x = x + 1
        Err(x) = "Tiada maklumat nama kedai."
    End If
Else
    If Frm92.CBB5 = vbNullString Then
        x = x + 1
        Err(x) = "Tiada maklumat nama kedai."
    End If
End If
If Frm92.CBB7 = vbNullString Then
    x = x + 1
    Err(x) = "Tiada maklumat jenis perbelanjaan."
End If
If Frm92.TB42 = vbNullString Then
    x = x + 1
    Err(x) = "Tiada no. invoice."
End If
If Frm92.CBB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih nama pekerja."
End If
If Frm92.TB47 <> "0.00" Then
    If Frm92.TB43 = vbNullString Then
        x = x + 1
        Err(x) = "Tiada No. ID GST."
    End If
End If
If Frm92.TB44 = vbnullstirng Then
    x = x + 1
    Err(x) = "Tiada maklumat tujuan pembelanjaan."
End If
If Frm92.TB46 = vbNullString Or (Frm92.TB46 <> vbNullString And Not IsNumeric(Frm92.TB46)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Jumlah Harga SR]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm92.TB47 = vbNullString Or (Frm92.TB47 <> vbNullString And Not IsNumeric(Frm92.TB47)) Then
    x = x + 1
    Err(x) = "Tiada maklumat [Jumlah Cukai GST SR]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm92.TB48 = vbNullString Or (Frm92.TB48 <> vbNullString And Not IsNumeric(Frm92.TB48)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Jumlah Harga ZR]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm92.TB49 = vbNullString Or (Frm92.TB49 <> vbNullString And Not IsNumeric(Frm92.TB49)) Then
    x = x + 1
    Err(x) = "Tiada maklumat [Jumlah Cukai GST ZR]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm92.TB45 = vbNullString Or (Frm92.TB45 <> vbNullString And Not IsNumeric(Frm92.TB45)) Then
    x = x + 1
    Err(x) = "Tiada maklumat [Jumlah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
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
        
        LM_FOUND = 0
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 39_akaun_expense where nama_kedai='" & UCase(Frm92.TB41) & "' AND no_resit='" & UCase(Frm92.TB42) & "' AND menu = 1 AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            LM_FOUND = 1
        End If
        
        rs.Close
        Set rs = Nothing
        
        If LM_FOUND = 1 Then
        
            Note = "Invoice " & UCase(Frm92.TB42) & " dari kedai " & UCase(Frm92.TB41) & " telah disimpan sebelum ini." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Teruskan?"
            
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbNo Then
                Exit Sub
            End If
            
        End If
        
        LM_NOW = Now

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        rs.Open "select * from 14_senarai_voucher", cn2, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        rs!tarikh = Frm92.DTPicker4
        rs!terminal = G_TERMINAL
        rs!write_timestamp = LM_NOW
        rs!Status = 1
        rs!nama_staff = MDI_frm1.L3_Text
        rs.Update
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        rs.Open "select * from 14_senarai_voucher where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & Frm92.DTPicker4 & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!ID) Then
            
                rs!no_voucher = "PV" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(rs!ID, "000000")
                G_No_RESIT_JUALAN = "PV" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(rs!ID, "000000")
                
            End If
            
            rs.Update
            
        Else
        
            MsgBox "Berlaku ralat semasa data cuba disimpan. Sila keluar dari menu ini dan cuba lagi.", vbCritical, "Error"
            
            rs.Close
            Set rs = Nothing
            
            Exit Sub
            
        End If
        
        rs.Close
        Set rs = Nothing
        
'### Update Akaun Bagi Expense ### - Start
        LM_HARGA_SR = 0
        LM_CUKAI_SR = 0
        LM_HARGA_ZR = 0
        LM_CUKAI_ZR = 0

        If Frm92.TB46 <> vbNullString And IsNumeric(Frm92.TB46) Then LM_HARGA_SR = Frm92.TB46
        If Frm92.TB47 <> vbNullString And IsNumeric(Frm92.TB47) Then LM_CUKAI_SR = Frm92.TB47
        If Frm92.TB48 <> vbNullString And IsNumeric(Frm92.TB48) Then LM_HARGA_ZR = Frm92.TB48
        If Frm92.TB49 <> vbNullString And IsNumeric(Frm92.TB49) Then LM_CUKAI_ZR = Frm92.TB49

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 39_akaun_expense", cn, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        
        If Frm92.CBB5 = "Lain-lain" Then
            
            rs!flag_kedai = 0 '0 : Supplier tidak berdaftar , 1 : Supplier berdaftar
            
            If Frm92.TB41 <> vbNullString Then 'Nama Kedai / Supplier
                rs!nama_kedai = UCase(Frm92.TB41)
            Else
                rs!nama_kedai = Null
            End If
        
        Else
            
            rs!flag_kedai = 1 '0 : Supplier tidak berdaftar , 1 : Supplier berdaftar
            
            If Frm92.CBB5 <> vbNullString Then
                rs!nama_kedai = Frm92.CBB5
            Else
                rs!nama_kedai = Null
            End If
            
        End If
        rs!no_voucher = G_No_RESIT_JUALAN
        If Frm92.TB42 <> vbNullString Then 'No. Invoice
            rs!no_resit = UCase(Frm92.TB42)
        Else
            rs!no_resit = Null
        End If
        If Frm92.TB44 <> vbNullString Then 'Tujuan
            rs!tujuan = UCase(Frm92.TB44)
        Else
            rs!tujuan = Null
        End If
        If Frm92.TB43 <> vbNullString Then 'No. ID GST
            rs!no_id_gst = UCase(Frm92.TB43)
        Else
            rs!no_id_gst = Null
        End If
        If Frm92.DTPicker4 <> vbNullString Then 'Tarikh
            rs!tarikh = Frm92.DTPicker4
        Else
            rs!tarikh = Null
        End If
        rs!jumlah_tanpa_gst = Format(LM_HARGA_SR + LM_CUKAI_SR, "0.00") 'Jumlah Tanpa GST (RM)
        If Frm92.TB45 <> vbNullString Then 'Jumlah Dengan GST (RM)
            rs!harga_dengan_gst = Format(Frm92.TB45, "0.00")
        Else
            rs!harga_dengan_gst = Null
        End If
        If Frm92.TB48 <> vbNullString Then 'Harga Keseluruhan Bagi Barang ZR
            rs!gst_zr_harga = Format(Frm92.TB48, "0.00")
        Else
            rs!gst_zr_harga = Null
        End If
        If Frm92.TB49 <> vbNullString Then 'Jumlah Cukai Bagi ZR
            rs!gst_zr_cukai = Format(Frm92.TB49, "0.00")
        Else
            rs!gst_zr_cukai = Null
        End If
        If Frm92.TB46 <> vbNullString Then 'Harga Keseluruhan Bagi Barang SR
            rs!gst_sr_harga = Format(Frm92.TB46, "0.00")
        Else
            rs!gst_sr_harga = Null
        End If
        If Frm92.TB47 <> vbNullString Then 'Jumlah Cukai Bagi SR
            rs!gst_sr_cukai = Format(Frm92.TB47, "0.00")
        Else
            rs!gst_sr_cukai = Null
        End If
        If Frm92.L42_Text <> vbNullString Then '% Cukai GST
            rs!gst_value = Frm92.L42_Text
        Else
            rs!gst_value = Null
        End If
        If Frm92.CBB2 <> vbNullString Then
            Frm92_LM_EMP_NO = Split(Frm92.CBB2, "  |  ")(1)
            rs!no_pekerja = Frm92_LM_EMP_NO 'No. Pekerja
        End If
        rs!jenis_expense = Frm92.CBB7
        rs!cawangan = G_CAWANGAN
        G_KEDAI = G_CAWANGAN
        rs!write_timestamp = LM_NOW
        rs!terminal = G_TERMINAL
        rs!Menu = 1
        rs!Status = 1
        If Frm92.CB10 = 1 Then
            rs!cara_bayaran = 0
        ElseIf Frm92.CB3 = 1 Then
            rs!cara_bayaran = 1
        ElseIf Frm92.CB4 = 1 Then
            rs!cara_bayaran = 2
        End If
        rs.Update

        rs.Close
        Set rs = Nothing
'### Update Akaun Bagi Expense ### - End

'### Update Log ### - Start
        user = MDI_frm1.L3_Text
        LogAct_Memory = "[" & user & "] Perbelanjaan kedai. No. Invoice [" & UCase(Frm92.TB42) & "]"
        LogDate_Memory = LM_NOW
        Call UpdateLog_Database
'### Update Log ### - End
        
        Call Frm92_Initial_Setting
        
        G_PREVIEW = 1
        Call frm92_cetak_pv
        
        MsgBox "Data telah berjaya disimpan.", vbInformation, "Info"
        
    End If
End If
End Sub

Private Sub CMD29_Click()
'On Error Resume Next
Dim Err(12)
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim LM_HARGA_SR As Double
Dim LM_CUKAI_SR As Double
Dim LM_HARGA_ZR As Double
Dim LM_CUKAI_ZR As Double
        
DATA_SAVE = 0
x = 0

If Frm92.L43_Text = vbnullstirng Then
    x = x + 1
    Err(x) = "Telah berlaku ralat. Sila keluar dari menu ini dan cuba lagi."
End If
If Frm92.CB10 = 0 And Frm92.CB3 = 0 And Frm92.CB4 = 0 Then
    x = x + 1
    Err(x) = "Sila buat cara bayaran dibuat."
End If
If Frm92.CBB5 = "Lain-lain" Then
    If Frm92.TB41 = vbNullString Then
        x = x + 1
        Err(x) = "Tiada maklumat nama kedai."
    End If
Else
    If Frm92.CBB5 = vbNullString Then
        x = x + 1
        Err(x) = "Tiada maklumat nama kedai."
    End If
End If
If Frm92.TB42 = vbNullString Then
    x = x + 1
    Err(x) = "Tiada no. invoice."
End If
If Frm92.CBB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih nama pekerja."
End If
If Frm92.TB47 <> "0.00" Then
    If Frm92.TB43 = vbNullString Then
        x = x + 1
        Err(x) = "Tiada No. ID GST."
    End If
End If
If Frm92.TB44 = vbnullstirng Then
    x = x + 1
    Err(x) = "Tiada maklumat tujuan pembelanjaan."
End If
If Frm92.TB46 = vbNullString Or (Frm92.TB46 <> vbNullString And Not IsNumeric(Frm92.TB46)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Jumlah Harga SR]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm92.TB47 = vbNullString Or (Frm92.TB47 <> vbNullString And Not IsNumeric(Frm92.TB47)) Then
    x = x + 1
    Err(x) = "Tiada maklumat [Jumlah Cukai GST SR]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm92.TB48 = vbNullString Or (Frm92.TB48 <> vbNullString And Not IsNumeric(Frm92.TB48)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Jumlah Harga ZR]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm92.TB49 = vbNullString Or (Frm92.TB49 <> vbNullString And Not IsNumeric(Frm92.TB49)) Then
    x = x + 1
    Err(x) = "Tiada maklumat [Jumlah Cukai GST ZR]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm92.TB45 = vbNullString Or (Frm92.TB45 <> vbNullString And Not IsNumeric(Frm92.TB45)) Then
    x = x + 1
    Err(x) = "Tiada maklumat [Jumlah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
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
        
        LM_FOUND = 0
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 39_akaun_expense where ID='" & Frm92.L43_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            If Not IsNull(rs!Status) Then
                
                If rs!Status = 0 Then
                    
                    MsgBox "Status bagi perbelanjaan ini telah dipadamkan dari sistem. Sila refresh data anda dan periksa status terbaru.", vbExclamation, "Info"
                    
                    rs.Close
                    Set rs = Nothing
                    
                    Exit Sub
                    
                End If
            
            End If
            
        End If
        
        rs.Close
        Set rs = Nothing
        
        LM_NOW = Now
        
'### Update Akaun Bagi Expense ### - Start
        LM_HARGA_SR = 0
        LM_CUKAI_SR = 0
        LM_HARGA_ZR = 0
        LM_CUKAI_ZR = 0

        If Frm92.TB46 <> vbNullString And IsNumeric(Frm92.TB46) Then LM_HARGA_SR = Frm92.TB46
        If Frm92.TB47 <> vbNullString And IsNumeric(Frm92.TB47) Then LM_CUKAI_SR = Frm92.TB47
        If Frm92.TB48 <> vbNullString And IsNumeric(Frm92.TB48) Then LM_HARGA_ZR = Frm92.TB48
        If Frm92.TB49 <> vbNullString And IsNumeric(Frm92.TB49) Then LM_CUKAI_ZR = Frm92.TB49

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 39_akaun_expense where ID='" & Frm92.L43_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Frm92.CBB5 = "Lain-lain" Then
                
                rs!flag_kedai = 0 '0 : Supplier tidak berdaftar , 1 : Supplier berdaftar
                
                If Frm92.TB41 <> vbNullString Then 'Nama Kedai / Supplier
                    rs!nama_kedai = UCase(Frm92.TB41)
                Else
                    rs!nama_kedai = Null
                End If
            
            Else
                
                rs!flag_kedai = 1 '0 : Supplier tidak berdaftar , 1 : Supplier berdaftar
                
                If Frm92.CBB5 <> vbNullString Then
                    rs!nama_kedai = Frm92.CBB5
                Else
                    rs!nama_kedai = Null
                End If
                
            End If
            If Frm92.TB42 <> vbNullString Then 'No. Invoice
                rs!no_resit = UCase(Frm92.TB42)
            Else
                rs!no_resit = Null
            End If
            If Frm92.TB44 <> vbNullString Then 'Tujuan
                rs!tujuan = UCase(Frm92.TB44)
            Else
                rs!tujuan = Null
            End If
            If Frm92.TB43 <> vbNullString Then 'No. ID GST
                rs!no_id_gst = UCase(Frm92.TB43)
            Else
                rs!no_id_gst = Null
            End If
            If Frm92.DTPicker4 <> vbNullString Then 'Tarikh
                rs!tarikh = Frm92.DTPicker4
            Else
                rs!tarikh = Null
            End If
            rs!jumlah_tanpa_gst = Format(LM_HARGA_SR + LM_CUKAI_SR, "0.00") 'Jumlah Tanpa GST (RM)
            If Frm92.TB45 <> vbNullString Then 'Jumlah Dengan GST (RM)
                rs!harga_dengan_gst = Format(Frm92.TB45, "0.00")
            Else
                rs!harga_dengan_gst = Null
            End If
            If Frm92.TB48 <> vbNullString Then 'Harga Keseluruhan Bagi Barang ZR
                rs!gst_zr_harga = Format(Frm92.TB48, "0.00")
            Else
                rs!gst_zr_harga = Null
            End If
            If Frm92.TB49 <> vbNullString Then 'Jumlah Cukai Bagi ZR
                rs!gst_zr_cukai = Format(Frm92.TB49, "0.00")
            Else
                rs!gst_zr_cukai = Null
            End If
            If Frm92.TB46 <> vbNullString Then 'Harga Keseluruhan Bagi Barang SR
                rs!gst_sr_harga = Format(Frm92.TB46, "0.00")
            Else
                rs!gst_sr_harga = Null
            End If
            If Frm92.TB47 <> vbNullString Then 'Jumlah Cukai Bagi SR
                rs!gst_sr_cukai = Format(Frm92.TB47, "0.00")
            Else
                rs!gst_sr_cukai = Null
            End If
            If Frm92.L42_Text <> vbNullString Then '% Cukai GST
                rs!gst_value = Frm92.L42_Text
            Else
                rs!gst_value = Null
            End If
            If Frm92.CBB2 <> vbNullString Then
                Frm92_LM_EMP_NO = Split(Frm92.CBB2, "  |  ")(1)
                rs!no_pekerja = Frm92_LM_EMP_NO 'No. Pekerja
            End If
            rs!write_timestamp = LM_NOW
            rs!terminal = G_TERMINAL
            rs!jenis_expense = Frm92.CBB7
            
            rs!Menu = 1
            rs!Status = 1
            If Frm92.CB10 = 1 Then
                rs!cara_bayaran = 0
            ElseIf Frm92.CB3 = 1 Then
                rs!cara_bayaran = 1
            ElseIf Frm92.CB4 = 1 Then
                rs!cara_bayaran = 2
            End If
            rs.Update
            
        End If

        rs.Close
        Set rs = Nothing
'### Update Akaun Bagi Expense ### - End

'### Update Log ### - Start
        user = MDI_frm1.L3_Text
        LogAct_Memory = "[" & user & "] Edit perbelanjaan kedai. ID [" & Frm92.L43_Text & "]"
        LogDate_Memory = LM_NOW
        Call UpdateLog_Database
'### Update Log ### - End
        
        GM_NEXT_PREV = 2
        
        Call Frm92_report_expenses_header
        Call Frm92_report_expenses

        Frm92.Frame5.Visible = False
        Frm92.Frame6.Visible = True
        
        MsgBox "Data telah berjaya disimpan.", vbInformation, "Info"
        
    End If
End If
End Sub

Private Sub CMD3_Click()
'on error resume next
Frm92.TB1 = vbNullString
Frm92.TB2 = "0.00"
Frm92.CMD1.Visible = True
Frm92.CMD2.Visible = False
Frm92.CMD3.Visible = False
Frm92.L16_Text = 0
Frm92.L18_Text.Visible = False
End Sub

Private Sub CMD30_Click()
'On Error Resume Next
Frm92.Frame5.Visible = False
Frm92.Frame6.Visible = True
End Sub

Private Sub CMD4_Click()
'On Error Resume Next
Dim Err(30)
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim Frm92_LM_JUMLAH_BAYARAN As Double
Dim Frm92_LM_HARGA As Double
Dim Frm92_LM_JUMLAH_SIMPANAN As Double
Dim Frm92_LM_GUNA_SIMPAN As Double

DATA_SAVE = 0
x = 0
Frm92_LM_JUMLAH_BAYARAN = 0 'Jumlah Bayaran
Frm92_LM_HARGA = 0 'Harga Keseluruhan
Frm92_LM_JUMLAH_SIMPANAN = 0 'Jumlah Simpanan Yang Ada
Frm92_LM_GUNA_SIMPAN = 0 'Jumlah Simpanan Yang Hendak Digunakan
Frm92_LM_Flag_SIMPANAN = 0 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
Frm92_LM_KATEGORI = 1

If Frm92.L20_Text = 0 Then
    x = x + 1
    Err(x) = "Tiada senarai servis."
End If
If Frm92.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih nama pekerja."
End If
If frm130.TB27 = vbNullString Or (frm130.TB27 <> vbNullString And Not IsNumeric(frm130.TB27)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara TUNAI. Sila masukkan 0 jika tiada bayaran secara tunai."
End If
If frm130.TB28 = vbNullString Or (frm130.TB28 <> vbNullString And Not IsNumeric(frm130.TB28)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara BANK IN. Sila masukkan 0 jika tiada bayaran secara bank in."
End If
If frm130.TB29 = vbNullString Or (frm130.TB29 <> vbNullString And Not IsNumeric(frm130.TB29)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara KAD KREDIT. Sila masukkan 0 jika tiada bayaran secara kad kredit."
End If
If frm130.TB21 = vbNullString Or (frm130.TB21 <> vbNullString And Not IsNumeric(frm130.TB21)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara duit simpanan di kedai. Sila masukkan 0 jika tiada bayaran secara simpanan di kedai."
End If
If (frm130.TB32 <> vbNullString And IsNumeric(frm130.TB32)) And (frm130.TB33 <> vbNullString And IsNumeric(frm130.TB33)) Then
    Frm92_LM_JUMLAH_BAYARAN = frm130.TB32 'Jumlah Bayaran
    Frm92_LM_HARGA = frm130.TB33 'Harga Keseluruhan
    
    If Frm92_LM_JUMLAH_BAYARAN <> Frm92_LM_HARGA Then
        x = x + 1
        Err(x) = "Jumlah bayaran tidak sama dengan jumlah harga barang."
    End If
End If

If (frm130.TB21 <> vbNullString And IsNumeric(frm130.TB21)) And (frm130.L26_Text <> vbNullString And IsNumeric(frm130.L26_Text)) Then
    Frm92_LM_JUMLAH_SIMPANAN = frm130.L26_Text  'Jumlah Simpanan Yang Ada
    Frm92_LM_GUNA_SIMPAN = frm130.TB21  'Jumlah Simpanan Yang Hendak Digunakan
    
    If Frm92_LM_GUNA_SIMPAN > Frm92_LM_JUMLAH_SIMPANAN Then
        x = x + 1
        Err(x) = "Jumlah simpanan yang hendak digunakan melebihi simpanan terkumpul yang ada."
    End If
End If

'### Periksa Samada Maklumat Pembeli Diisi Dalam Kedua-dua Ruangan Berdaftar Dan Tidak Berdaftar ### - Start
If x = 0 Then
    If Frm92.L51_Text <> vbNullString And Frm92.L52_Text <> vbNullString Then
    
        MsgBox "Data bagi pelanggan telah diisi bagi kedua-dua ruangan pembeli berdaftar dan tidak berdaftar." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Sila padam salah satu yang tidak berkenaan.", vbExclamation, "Info"
                    
        Exit Sub
          
    End If
End If
'### Periksa Samada Maklumat Pembeli Diisi Dalam Kedua-dua Ruangan Berdaftar Dan Tidak Berdaftar ### - End

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    
    If Frm92.L51_Text <> vbNullString And Frm92.L52_Text = vbNullString Then
        Note = "Adakah anda yakin untuk teruskan urusan jualan ini ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Data jualan akan disimpan ke dalam sistem." & vbCrLf & _
                vbNullString & vbCrLf & _
                "*** Oleh ini pembeli ini tidak berdaftar dengan sistem, oleh itu data dan rekod pembelian pembeli ini tidak akan disimpan di dalam sistem ***"
    End If
    If Frm92.L51_Text = vbNullString And Frm92.L52_Text <> vbNullString Then
        Note = "Adakah anda yakin untuk teruskan urusan jualan ini ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Data jualan akan disimpan ke dalam sistem."
    End If
    If Frm92.L51_Text = vbNullString And Frm92.L52_Text = vbNullString Then
    
        Note = "TIADA maklumat bagi pembeli telah diisi." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Maklumat pembeli tidak akan dicetak di dalam invoice pembeli." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Adakah anda yakin untuk teruskan urusan jualan ini ?"
        
    End If
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
        '### Pop up confirmation bagi jualan bagi invoice tidak rasmi
        If Frm92.CB9 = 1 Then
        
            Note = "Jualan ini dibuat dengan pilihan INVOICE TIDAK RASMI." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Anda TIDAK BOLEH mengubah jenis invoice jika data ini telah disimpan." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Teruskan?"
                
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbNo Then
            
                Exit Sub
            
            End If
            
        End If
    
' ### Periksa kategori pembeli ### - Start
        If Frm92.L52_Text <> vbNullString Then
        
            If Frm28.L5_Text <> vbNullString Then
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
                If Not rs.EOF Then
                
                    If Not IsNull(rs!kategori_pelanggan) Then Frm92_LM_KATEGORI = rs!kategori_pelanggan
                    
                End If
                
                rs.Close
                Set rs = Nothing
                
            Else
            
                Frm92_LM_KATEGORI = 1
                
            End If
        
        Else
        
            'Frm92_LM_KATEGORI = 1
        
        End If
' ### Periksa kategori pembeli ### - End

        G_JENIS_URUSAN = 4

        '$$$ No. staff $$$ - Start
        If InStr(1, Frm92.CBB1, "  |  ") <> 0 Then
            Frm92_LM_EMP_NO = Split(Frm92.CBB1, "  |  ")(1)
            Frm92_LM_EMP_NAMA = Split(Frm92.CBB1, "  |  ")(0)
        Else
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!NoPekerja) Then Frm92_LM_EMP_NO = rs!NoPekerja
    
            End If
            
            rs.Close
            Set rs = Nothing
        
        End If
    
'### Update Akaun Bagi Servis ### - Start
        'If Frm92.CB9 = 0 Then Frm92_LM_No_RESIT_SERVIS = Frm92.L17_Text 'Turutan no invoice (rasmi)
        'If Frm92.CB9 = 1 Then Frm92_LM_No_RESIT_SERVIS = Frm92.L28_Text 'Turutan no invoice (Tidak rasmi)
        
'---------------------------------------No. Invoice
        LM_NOW = Now
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        If Frm92.CB9 = 0 Then rs.Open "select * from 4_senarai_invoice_rasmi", cn2, adOpenKeyset, adLockOptimistic
        If Frm92.CB9 = 1 Then rs.Open "select * from 5_senarai_invoice_tidak_rasmi", cn2, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        rs!tarikh = Frm92.DTPicker1
        rs!terminal = G_TERMINAL
        rs!write_timestamp = LM_NOW
        rs!Status = 1
        rs!nama_staff = MDI_frm1.L3_Text
        rs.Update
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        If Frm92.CB9 = 0 Then rs.Open "select * from 4_senarai_invoice_rasmi where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & Frm92.DTPicker1 & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic
        If Frm92.CB9 = 1 Then rs.Open "select * from 5_senarai_invoice_tidak_rasmi where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & Frm92.DTPicker1 & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!ID) Then
                
                If Frm92.CB9 = 0 Then rs!no_invoice = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(rs!ID, "000000")
                If Frm92.CB9 = 1 Then rs!no_invoice = "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(rs!ID, "000000")
                Frm92_LM_No_RESIT_SERVIS = rs!ID 'No. Rujukan Belian
                rs.Update
                
            End If
            
            'rs.Update
            
        Else
        
            MsgBox "Berlaku ralat semasa data cuba disimpan. Sila keluar dari menu ini dan cuba lagi.", vbCritical, "Error"
            
            rs.Close
            Set rs = Nothing
            
            Exit Sub
            
        End If
        
        rs.Close
        Set rs = Nothing
'---------------------------------------No. Invoice

Re_gen_no_resit2:
'###Masukkan Jualan Ke Dalam Table Akaun Jualan### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm92.CB9 = 0 Then rs.Open "select * from 22_jualan where no_resit='" & "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm92_LM_No_RESIT_SERVIS, "000000") & "' AND bil_rasmi = 1 AND status = 1", cn, adOpenKeyset, adLockOptimistic
        If Frm92.CB9 = 1 Then rs.Open "select * from 22_jualan where no_resit='" & "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm92_LM_No_RESIT_SERVIS, "000000") & "' AND bil_rasmi = 0 AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
            rs.AddNew
            If Frm92.CB9 = 0 Then
            
                If Frm92.L17_Text <> vbNullString Then
                    rs!no_resit = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm92_LM_No_RESIT_SERVIS, "000000") 'No. invoice rasmi
                    G_No_RESIT_SERVIS = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm92_LM_No_RESIT_SERVIS, "000000") 'No. invoice rasmi
                Else
                    rs!no_resit = Null 'No. invoice rasmi
                End If
                rs!bil_rasmi = 1 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
                
            Else
            
                If Frm92.L28_Text <> vbNullString Then
                    rs!no_resit = "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm92_LM_No_RESIT_SERVIS, "000000") 'No. invoice tidak rasmi
                    G_No_RESIT_SERVIS = "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm92_LM_No_RESIT_SERVIS, "000000") 'No. invoice tidak rasmi
                Else
                    rs!no_resit = Null 'No. invoice tidak rasmi
                End If
                rs!bil_rasmi = 0 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
            
            End If
            
            rs!tarikh = Frm92.DTPicker1
            rs!status_r = 0
            rs!tunai = Format(0, "0.00")
            rs!bank_in = Format(0, "0.00") 'Cara Bayaran : Bank In
            rs!kad_kredit = Format(0, "0.00") 'Cara Bayaran : Kad Kredit
            rs!duit_simpanan_kedai = Format(0, "0.00") 'Cara Bayaran : Simpanan Duit Di Kedai
            rs!cek = Format(0, "0.00") 'Cara Bayaran : Cek
            
            If frm130.TB27 <> vbNullString Then
                rs!tunai = Format(frm130.TB27, "0.00") 'Cara Bayaran : Tunai
            Else
                rs!tunai = "0.00" 'Cara Bayaran : Tunai
            End If
            If frm130.TB28 <> vbNullString Then
                rs!bank_in = Format(frm130.TB28, "0.00") 'Cara Bayaran : Bank In
            Else
                rs!bank_in = "0.00" 'Cara Bayaran : Bank In
            End If
            
            If frm130.TB29 <> vbNullString Then
                rs!kad_kredit = Format(frm130.TB29, "0.00") 'Cara Bayaran : Kad Kredit
                If Format(frm130.TB29, "0.00") <> "0.00" Then
                    
                    If frm130.CBB2 <> vbNullString Then
                        rs!jenis_kad = frm130.CBB2
                    Else
                        rs!jenis_kad = Null
                    End If
                    If frm130.L31_Text <> vbNullString Then
                        rs!cas_Kad_Kredit = Format(frm130.L31_Text, "0.00") 'Cara Bayaran : Cas Kad Kredit (%)
                    Else
                        rs!cas_Kad_Kredit = "0.00" 'Cara Bayaran : Cas Kad Kredit (%)
                    End If
                    If frm130.L32_Text <> vbNullString Then
                        rs!jumlah_cas_kad_kredit = Format(frm130.L32_Text, "0.00") 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
                    Else
                        rs!jumlah_cas_kad_kredit = "0.00" 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
                    End If
                    If frm130.L81_Text <> vbNullString Then
                        rs!gst_kad_kredit = Format(frm130.L81_Text, "0.00") 'Cara Bayaran : Jumlah GST kad kredit (RM)
                    Else
                        rs!gst_kad_kredit = "0.00" 'Cara Bayaran : Jumlah GST kad kredit (RM)
                    End If
                    If frm130.L82_Text <> vbNullString Then
                        rs!jumlah_potongan_kad_kredit = Format(frm130.L82_Text, "0.00") 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
                    Else
                        rs!jumlah_potongan_kad_kredit = "0.00" 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
                    End If
                    If frm130.L8_Text <> vbNullString Then
                        rs!kadar_gst_kad_kredit = Format(frm130.L8_Text, "0.00") 'Cara Bayaran : Kadar GST bagi kad kredit
                    Else
                        rs!kadar_gst_kad_kredit = "0.00" 'Cara Bayaran : Kadar GST bagi kad kredit
                    End If
                    rs!epp = 0 '0 : Bayaran selain dari EPP , 1 : Bayaran secara EPP
                    rs!approval_code_epp = Null 'Approval Code (EPP)
                    
                End If
            End If

            If frm130.TB21 <> vbNullString Then
                If Format(frm130.TB21, "0.00") <> "0.00" Then
                    Frm92_LM_Flag_SIMPANAN = 1 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
                End If
                rs!duit_simpanan_kedai = Format(frm130.TB21, "0.00") 'Cara Bayaran : Simpanan Duit Di Kedai
            Else
                rs!duit_simpanan_kedai = "0.00" 'Cara Bayaran : Simpanan Duit Di Kedai
            End If
            If frm130.TB32 <> vbNullString Then
                rs!jumlah_bayaran = Format(frm130.TB32, "0.00") 'Cara Bayaran : Jumlah Bayaran
            Else
                rs!jumlah_bayaran = Null 'Cara Bayaran : Jumlah Bayaran
            End If
            If Frm92.L9_Text <> vbNullString Then
                rs!harga_barang = Format(Frm92.L9_Text, "0.00") 'Jumlah Harga Barang Tanpa GST (RM)
            Else
                rs!harga_barang = Null 'Jumlah Harga Barang Tanpa GST (RM)
            End If
            If Frm92.L15_Text <> vbNullString Then 'Kadar cukai GST
                rs!kadar_gst = Format(Frm92.L15_Text, "0.00")
            Else
                rs!kadar_gst = Null
            End If
            If Frm92.L7_Text <> vbNullString Then
                rs!jumlah_cukai_gst = Format(Frm92.L7_Text, "0.00") 'Jumlah Cukai GST (ZR + SR)
            Else
                rs!jumlah_cukai_gst = Null 'Jumlah Cukai GST (ZR + SR)
            End If
            If Frm92.L10_Text <> vbNullString Then
                rs!harga_barang_dengan_gst = Format(Frm92.L10_Text, "0.00") 'Jumlah Harga Barang Dengan GST (RM)
                rs!harga_lepas_diskaun = Format(Frm92.L10_Text, "0.00") 'Harga Selepas Diskaun (RM)
                rs!harga_jualan = Format(Frm92.L10_Text, "0.00") 'Jumlah Harga Jualan (RM)
                rs!jumlah_perlu_bayar = Format(Frm92.L10_Text, "0.00") 'Jumlah Bayaran Yang Perlu Dibuat (RM)
            Else
                rs!harga_barang_dengan_gst = Null 'Jumlah Harga Barang Dengan GST (RM)
                rs!harga_lepas_diskaun = Null 'Harga Selepas Diskaun (RM)
                rs!harga_jualan = Null 'Jumlah Harga Jualan (RM)
                rs!jumlah_perlu_bayar = Null 'Jumlah Bayaran Yang Perlu Dibuat (RM)
            End If
            rs!diskaun = Null 'Jumlah Diskaun (%)
            rs!adjustment = "0.00" 'Adjustment (RM)
            rs!loss_trade_in = Null 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (%)
            rs!loss_trade_in_rm = Null 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (RM)
            rs!flag_bayaran = 0 '0 : Pembeli Bayar , 1 : Kedai Bayar
            If Frm92.L20_Text <> vbNullString Then 'Kuantiti servis
                rs!kuantiti_barang = Frm92.L20_Text
            Else
                rs!kuantiti_barang = Null
            End If
            rs!JUMLAH_BERAT = Null 'Jumlah Berat Barang Yang Dijual
            If Frm92.L11_Text <> vbNullString Then
                rs!gst_zr_harga = Format(Frm92.L11_Text, "0.00") 'Harga Keseluruhan Bagi Barang ZR
            Else
                rs!gst_zr_harga = Null 'Harga Keseluruhan Bagi Barang ZR
            End If
            If Frm92.L12_Text <> vbNullString Then
                rs!gst_zr_cukai = Format(Frm92.L12_Text, "0.00") 'Jumlah Cukai Bagi ZR
            Else
                rs!gst_zr_cukai = Null 'Jumlah Cukai Bagi ZR
            End If
            If Frm92.L13_Text <> vbNullString Then
                rs!gst_sr_harga = Format(Frm92.L13_Text, "0.00") 'Harga Keseluruhan Bagi Barang SR
            Else
                rs!gst_sr_harga = Null 'Harga Keseluruhan Bagi Barang SR
            End If
            If Frm92.L14_Text <> vbNullString Then
                rs!gst_sr_cukai = Format(Frm92.L14_Text, "0.00") 'Jumlah Cukai Bagi SR
            Else
                rs!gst_sr_cukai = Null 'Jumlah Cukai Bagi SR
            End If
            rs!caj_pos = "0.00"
            rs!no_tracking = Null

            rs!no_pekerja = Frm92_LM_EMP_NO 'No. Pekerja
            rs!nama_pekerja = Frm92_LM_EMP_NAMA
            If Frm92.L52_Text <> vbNullString Then
                If Frm28.L5_Text <> vbNullString Then
                    rs!no_rujukan_pembeli = Frm28.L5_Text 'No. Rujukan Pembeli
                    LM_NO_PELANGGAN = Frm28.L5_Text 'No. Rujukan Pembeli
                Else
                    rs!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
                End If
            Else
                rs!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
            End If
            rs!no_rujukan_agen_dropship = Null 'No. Rujukan Agen Dropship
            rs!flag_trade_in = 0 '0 : Tiada Urusan Trade in , 1 : Ada Urusan Trade In
            rs!no_resit_trade_in = Null 'No. Resit Trade In
            rs!jumlah_trade_in = Null 'Jumlah Resit Trade In (RM)
            rs!jenis_trade_in = Null '1 : Trade in (Voucher) , 2 : Belian dengan trade in
            rs!invoice_type = 0 '0 : Unlimited , Selain 0 (Limited : Mengikut nombor yang dimasukkan)

'Zakaria&Sons
'1 : Pembeli biasa
'2 : Ahli biasa
'3 : Silver
'4 : Gold
'5 : Platinum

            rs!kategori_pembeli = Frm92_LM_KATEGORI
            rs!jualan_online = 0
            rs!kupon_diskaun = "0.00"
            rs!Status = 1
            rs!terminal = G_TERMINAL
            rs!cawangan = G_CAWANGAN
            G_KEDAI = G_CAWANGAN
            rs!write_timestamp = LM_NOW
            rs!Menu = 1
            
            DATA_SAVE = 1
            rs.Update
        Else
        
            Frm92_LM_No_RESIT_SERVIS = Frm92_LM_No_RESIT_SERVIS + 1
            If Frm92.CB9 = 0 Then Frm92.L17_Text = Frm92_LM_No_RESIT_SERVIS
            If Frm92.CB9 = 1 Then Frm92.L28_Text = Frm92_LM_No_RESIT_SERVIS
            
            rs.Close
            Set rs = Nothing
            GoTo Re_gen_no_resit2:
            
        End If
        
        rs.Close
        Set rs = Nothing
'###Masukkan Jualan Ke Dalam Table Akaun Jualan### - End
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "insert into 35_senarai_servis(tarikh,no_resit_servis,no_pelanggan,write_timestamp,no_pekerja,terminal,cawangan,detail,jumlah,jenis_gst,jumlah_gst,kod_gst,harga_tanpa_gst,harga_dengan_gst)" & _
                    "select '" & Frm92.DTPicker1 & "','" & G_No_RESIT_SERVIS & "','" & LM_NO_PELANGGAN & "','" & LM_NOW & "','" & Frm92_LM_EMP_NO & "','" & G_TERMINAL & "','" & G_CAWANGAN & "',detail,jumlah,jenis_gst,jumlah_gst,kod_gst,harga_tanpa_gst,harga_dengan_gst from " & G_SERVICE_TEMP & ""
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing

'### Maklumat Pembeli Yang Tidak Berdaftar Dengan Kedai ### - Start (08-07-2015)
        If Frm92.L51_Text <> vbNullString Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 44_senarai_pelanggan", cn, adOpenKeyset, adLockOptimistic
            
            rs.AddNew
            rs!tarikh = Frm92.DTPicker1 'Tarikh
            If Frm92.CB9 = 0 Then rs!no_resit = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm92_LM_No_RESIT_SERVIS, "000000")
            If Frm92.CB9 = 1 Then rs!no_resit = "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm92_LM_No_RESIT_SERVIS, "000000")
            If Frm26.TB1 <> vbNullString Then 'Nama
                rs!Nama = UCase(Frm26.TB1)
            Else
                rs!Nama = Null
            End If
            If Frm26.TB2 <> vbNullString Then 'No. Telefon
                rs!no_tel = UCase(Frm26.TB2)
            Else
                rs!no_tel = Null
            End If
            rs!write_timestamp = LM_NOW
            rs!no_staff = Frm92_LM_EMP_NO 'No. Pekerja
            rs!terminal = G_TERMINAL
            rs!jenis_urusan = G_JENIS_URUSAN
            rs!cawangan = G_CAWANGAN
            rs.Update
            
            rs.Close
            Set rs = Nothing
            
        End If
'### Maklumat Pembeli Yang Tidak Berdaftar Dengan Kedai ### - End

'### Update Simpanan ### - Start
        If Frm92_LM_Flag_SIMPANAN = 1 Then '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai

            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
            
                G_ID = rs!ID
                Call recovery_senarai_pelanggan
                
                Frm92_LM_JUMLAH_SIMPANAN = frm130.L26_Text  'Jumlah Simpanan Yang Ada
                Frm92_LM_GUNA_SIMPAN = frm130.TB21  'Jumlah Simpanan Yang Hendak Digunakan
                
                rs!baki_simpanan = Format(Frm92_LM_JUMLAH_SIMPANAN - Frm92_LM_GUNA_SIMPAN, "0.00") 'Baki Simpanan
                rs!write_timestamp2 = LM_NOW
                rs!no_staff = Frm92_LM_EMP_NO 'No. Pekerja
                rs!terminal = G_TERMINAL
                rs!jenis_urusan = G_JENIS_URUSAN
                
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 24_rekod_kewangan_pelanggan", cn, adOpenKeyset, adLockOptimistic
            
            rs.AddNew
            rs!tarikh = Frm92.DTPicker1 'Tarikh
            rs!jenis = 1 '0 : Simpanan , 1 : Penggunaan Duit
            rs!no_rujukan_pelanggan = Frm28.L5_Text 'No. Rujukan Pelanggan
            
            If Frm92.CB9 = 0 Then rs!no_resit = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm92_LM_No_RESIT_SERVIS, "000000")
            If Frm92.CB9 = 1 Then rs!no_resit = "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm92_LM_No_RESIT_SERVIS, "000000")
            
            rs!jumlah = Format(frm130.TB21, "0.00") 'Jumlah Simpanan Yang Digunakan (RM)
            rs!jenis_penggunaan = 3 '0 : Belian Barangan Kemas , 1 : Ansuran , 2 : Tempahan (Deposit) , 3 : Servis , 4 : Tempahan (Ambilan Barang)
            rs!no_rujukan_pekerja = Frm92_LM_EMP_NO 'No. Pekerja
            rs!terminal = G_TERMINAL
            rs!write_timestamp = LM_NOW
            rs!jenis_urusan = G_JENIS_URUSAN
            rs!cawangan = G_CAWANGAN
            rs!Status = 1
            rs.Update
            
            rs.Close
            Set rs = Nothing
        
        End If
'### Update Simpanan ### - End

'### Update Log ### - Start
        user = MDI_frm1.L3_Text
        LogAct_Memory = "[" & G_LOGIN_USER & "] Servis kepada pelanggan. No. Invoice [" & G_No_RESIT_SERVIS & "]"
        LogDate_Memory = LM_NOW
        Call UpdateLog_Database
'### Update Log ### - End

'### Update No. Resit Servis ### - Start
        'Set rs = New ADODB.Recordset
        'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        'rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
        
        'If Not rs.EOF Then
        '    If rs!Default1 = "Default" Then
            
        '        If Frm92.CB9 = 0 Then rs!ResitNo = Frm92_LM_No_RESIT_SERVIS + 1 'No. invoice rasmi
        '        If Frm92.CB9 = 1 Then rs!no_rujukan_tak_rasmi = Frm92_LM_No_RESIT_SERVIS + 1 'No. invoice tidak rasmi
                
        '        rs.Update
        '    End If
        'End If
        
        'rs.Close
        'Set rs = Nothing
'### Update No. Resit Servis ### - End
        
        Call Frm92_Initial_Setting
        Call frm92_senarai_service_header
        Call frm130_reset
        
        G_PREVIEW = 1
        
        Note = "Data telah berjaya disimpan." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Cetak invoice ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
            Call Frm92_Resit_Servis
        End If
        
    End If
End If
End Sub
Private Sub CMD5_Click()
'On Error Resume Next
Dim Err(30)

Dim Frm92_LM_JUMLAH_BAYARAN As Double
Dim Frm92_LM_HARGA As Double
Dim Frm92_LM_JUMLAH_SIMPANAN As Double
Dim Frm92_LM_GUNA_SIMPAN As Double
Dim Frm92_LM_JUMLAH As Double
Dim Frm92_LM_BAKI_ASAL As Double

DATA_SAVE = 0
x = 0
Frm92_LM_JUMLAH = 0
Frm92_LM_JUMLAH_BAYARAN = 0 'Jumlah Bayaran
Frm92_LM_HARGA = 0 'Harga Keseluruhan
Frm92_LM_JUMLAH_SIMPANAN = 0 'Jumlah Simpanan Yang Ada
Frm92_LM_GUNA_SIMPAN = 0 'Jumlah Simpanan Yang Hendak Digunakan
Frm92_LM_Flag_SIMPANAN = 0 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
Frm92_LM_FLAG_SIMPANAN_ASAL = 0
Frm92_LM_BAKI_ASAL = 0
Frm92_LM_KATEGORI = 1

G_JENIS_URUSAN = 5

If Frm92.L20_Text = 0 Then
    x = x + 1
    Err(x) = "Tiada senarai servis."
End If
If Frm92.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih nama pekerja."
End If
If frm130.TB27 = vbNullString Or (frm130.TB27 <> vbNullString And Not IsNumeric(frm130.TB27)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara TUNAI. Sila masukkan 0 jika tiada bayaran secara tunai."
End If
If frm130.TB28 = vbNullString Or (frm130.TB28 <> vbNullString And Not IsNumeric(frm130.TB28)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara BANK IN. Sila masukkan 0 jika tiada bayaran secara bank in."
End If
If frm130.TB29 = vbNullString Or (frm130.TB29 <> vbNullString And Not IsNumeric(frm130.TB29)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara KAD KREDIT. Sila masukkan 0 jika tiada bayaran secara kad kredit."
End If
If frm130.TB21 = vbNullString Or (frm130.TB21 <> vbNullString And Not IsNumeric(frm130.TB21)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara duit simpanan di kedai. Sila masukkan 0 jika tiada bayaran secara simpanan di kedai."
End If
If (frm130.TB32 <> vbNullString And IsNumeric(frm130.TB32)) And (frm130.TB33 <> vbNullString And IsNumeric(frm130.TB33)) Then
    Frm92_LM_JUMLAH_BAYARAN = frm130.TB32 'Jumlah Bayaran
    Frm92_LM_HARGA = frm130.TB33 'Harga Keseluruhan
    
    If Frm92_LM_JUMLAH_BAYARAN <> Frm92_LM_HARGA Then
        x = x + 1
        Err(x) = "Jumlah bayaran tidak sama dengan jumlah harga barang."
    End If
End If

If (frm130.TB21 <> vbNullString And IsNumeric(frm130.TB21)) And (frm130.L26_Text <> vbNullString And IsNumeric(frm130.L26_Text)) Then
    Frm92_LM_JUMLAH_SIMPANAN = frm130.L26_Text  'Jumlah Simpanan Yang Ada
    Frm92_LM_GUNA_SIMPAN = frm130.TB21  'Jumlah Simpanan Yang Hendak Digunakan
    
    If Frm92_LM_GUNA_SIMPAN > Frm92_LM_JUMLAH_SIMPANAN Then
        x = x + 1
        Err(x) = "Jumlah simpanan yang hendak digunakan melebihi simpanan terkumpul yang ada."
    End If
End If

'### Periksa Samada Maklumat Pembeli Diisi Dalam Kedua-dua Ruangan Berdaftar Dan Tidak Berdaftar ### - Start
If x = 0 Then
    If Frm92.L51_Text <> vbNullString And Frm92.L52_Text <> vbNullString Then
    
        MsgBox "Data bagi pelanggan telah diisi bagi kedua-dua ruangan pembeli berdaftar dan tidak berdaftar." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Sila padam salah satu yang tidak berkenaan.", vbExclamation, "Info"
                    
        Exit Sub
          
    End If
End If
'### Periksa Samada Maklumat Pembeli Diisi Dalam Kedua-dua Ruangan Berdaftar Dan Tidak Berdaftar ### - End

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else

    If Frm92.L51_Text <> vbNullString And Frm92.L52_Text = vbNullString Then
        Note = "Adakah anda yakin untuk teruskan urusan jualan ini ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Data jualan akan disimpan ke dalam sistem." & vbCrLf & _
                vbNullString & vbCrLf & _
                "*** Oleh ini pembeli ini tidak berdaftar dengan sistem, oleh itu data dan rekod pembelian pembeli ini tidak akan disimpan di dalam sistem ***"
    End If
    If Frm92.L51_Text = vbNullString And Frm92.L52_Text <> vbNullString Then
        Note = "Adakah anda yakin untuk teruskan urusan jualan ini ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Data jualan akan disimpan ke dalam sistem."
    End If
    If Frm92.L51_Text = vbNullString And Frm92.L52_Text = vbNullString Then
    
        Note = "TIADA maklumat bagi pembeli telah diisi." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Maklumat pembeli tidak akan dicetak di dalam invoice pembeli." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Adakah anda yakin untuk teruskan urusan jualan ini ?"
        
    End If
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then
    
        '$$$ No. staff $$$ - Start
        If InStr(1, Frm92.CBB1, "  |  ") <> 0 Then
            Frm92_LM_EMP_NO = Split(Frm92.CBB1, "  |  ")(1)
        Else
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!NoPekerja) Then Frm92_LM_EMP_NO = rs!NoPekerja
    
            End If
            
            rs.Close
            Set rs = Nothing
        End If
    
' ### Periksa kategori pembeli ### - Start
        If Frm92.L52_Text <> vbNullString Then
            If Frm28.L5_Text <> vbNullString Then
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
                If Not rs.EOF Then
                    If Not IsNull(rs!kategori_pelanggan) Then Frm92_LM_KATEGORI = rs!kategori_pelanggan
                End If
                
                rs.Close
                Set rs = Nothing
                
            Else
                Frm92_LM_KATEGORI = 1
            End If
        End If
' ### Periksa kategori pembeli ### - End
    
        LM_NOW = Now
        G_No_RESIT_SERVIS = Frm92.L17_Text
        
'### Recovery ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

        strsql = "insert into " & G_RECOVERY_DATABASE & ".35_senarai_servis" & "(id_asal,tarikh,no_resit_servis,no_pelanggan,write_timestamp,terminal,detail,jumlah,jenis_gst," _
                    & "write_writestamp2,terminal2,no_staff," _
                    & "jumlah_gst,kod_gst,harga_tanpa_gst,harga_dengan_gst)" & _
                    "select ID,tarikh,no_resit_servis,no_pelanggan,write_timestamp,terminal,detail,jumlah,jenis_gst," _
                    & "'" & LM_NOW & "','" & G_TERMINAL & "','" & G_LOGIN_USER & "'," _
                    & "jumlah_gst,kod_gst,harga_tanpa_gst,harga_dengan_gst " _
                    & "from " & G_SERVER_DATABASE & ".35_senarai_servis WHERE no_resit_servis='" & G_No_RESIT_SERVIS & "'"
                                
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Recovery ### - End

'### Padam Data Senarai Servis ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "DELETE from 35_senarai_servis WHERE no_resit_servis='" & G_No_RESIT_SERVIS & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing

'### Padam Data Senarai Servis ### - End

'### Padam Penggunaan Duit Pelanggan ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 24_rekod_kewangan_pelanggan where no_resit='" & G_No_RESIT_SERVIS & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            G_ID = rs!ID
            Call recovery_24_rekod_kewangan_pelanggan
                
            Frm92_LM_FLAG_SIMPANAN_ASAL = 1
            
            If Not IsNull(rs!no_rujukan_pelanggan) Then Frm92_LM_No_PELANGGAN = rs!no_rujukan_pelanggan 'No. Rujukan Pelanggan
            If Not IsNull(rs!jumlah) Then
                If IsNumeric(rs!jumlah) Then Frm92_LM_JUMLAH = rs!jumlah
            End If
            
            rs.Delete
            rs.Update
            
        End If
        
        rs.Close
        Set rs = Nothing
        
        
        If Frm92_LM_FLAG_SIMPANAN_ASAL = 1 Then

            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm92_LM_No_PELANGGAN & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                G_ID = rs!ID
                Call recovery_senarai_pelanggan
                
                If Not IsNull(rs!baki_simpanan) Then Frm92_LM_BAKI_ASAL = rs!baki_simpanan 'Baki Simpanan Asal
                
                rs!baki_simpanan = Format(Frm92_LM_BAKI_ASAL + Frm92_LM_JUMLAH, "0.00") 'Baki Simpanan
                
                rs!write_timestamp2 = LM_NOW
                rs!no_staff = Frm92_LM_EMP_NO 'No. Pekerja
                rs!terminal = G_TERMINAL
                rs!jenis_urusan = G_JENIS_URUSAN
                rs.Update
                
            End If
            
            rs.Close
            Set rs = Nothing
            
        End If
'### Padam Penggunaan Duit Pelanggan ### - End

'### Padam Rekod Belian Pembeli Yang Tidak Berdaftar Dengan Kedai ### - Start (08-07-2015)
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 44_senarai_pelanggan where no_resit='" & G_No_RESIT_SERVIS & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then

            G_ID = rs!ID
            Call recovery_44_senarai_pelanggan
            
            rs.Delete
            rs.Update
        
        End If
        
        rs.Close
        Set rs = Nothing
            
'### Padam Rekod Belian Pembeli Yang Tidak Berdaftar Dengan Kedai ### - End (08-07-2015)

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 22_jualan where no_resit='" & G_No_RESIT_SERVIS & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then

            G_ID = rs!ID
            Call recovery_22_jualan
            
            rs!tarikh = Frm92.DTPicker1
        
            rs!tunai = Format(0, "0.00")
            rs!bank_in = Format(0, "0.00") 'Cara Bayaran : Bank In
            rs!kad_kredit = Format(0, "0.00") 'Cara Bayaran : Kad Kredit
            rs!duit_simpanan_kedai = Format(0, "0.00") 'Cara Bayaran : Simpanan Duit Di Kedai
            rs!cek = Format(0, "0.00") 'Cara Bayaran : Cek
            
            If frm130.TB27 <> vbNullString Then
                rs!tunai = Format(frm130.TB27, "0.00") 'Cara Bayaran : Tunai
            Else
                rs!tunai = "0.00" 'Cara Bayaran : Tunai
            End If
            If frm130.TB28 <> vbNullString Then
                rs!bank_in = Format(frm130.TB28, "0.00") 'Cara Bayaran : Bank In
            Else
                rs!bank_in = "0.00" 'Cara Bayaran : Bank In
            End If
            
            If frm130.TB29 <> vbNullString Then
                rs!kad_kredit = Format(frm130.TB29, "0.00") 'Cara Bayaran : Kad Kredit
                If Format(frm130.TB29, "0.00") <> "0.00" Then
                    
                    If frm130.CBB2 <> vbNullString Then
                        rs!jenis_kad = Frm92.CBB2
                    Else
                        rs!jenis_kad = Null
                    End If
                    If frm130.L31_Text <> vbNullString Then
                        rs!cas_Kad_Kredit = Format(frm130.L31_Text, "0.00") 'Cara Bayaran : Cas Kad Kredit (%)
                    Else
                        rs!cas_Kad_Kredit = "0.00" 'Cara Bayaran : Cas Kad Kredit (%)
                    End If
                    If frm130.L32_Text <> vbNullString Then
                        rs!jumlah_cas_kad_kredit = Format(frm130.L32_Text, "0.00") 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
                    Else
                        rs!jumlah_cas_kad_kredit = "0.00" 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
                    End If
                    If frm130.L81_Text <> vbNullString Then
                        rs!gst_kad_kredit = Format(frm130.L81_Text, "0.00") 'Cara Bayaran : Jumlah GST kad kredit (RM)
                    Else
                        rs!gst_kad_kredit = "0.00" 'Cara Bayaran : Jumlah GST kad kredit (RM)
                    End If
                    If frm130.L82_Text <> vbNullString Then
                        rs!jumlah_potongan_kad_kredit = Format(frm130.L82_Text, "0.00") 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
                    Else
                        rs!jumlah_potongan_kad_kredit = "0.00" 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
                    End If
                    If frm130.L8_Text <> vbNullString Then
                        rs!kadar_gst_kad_kredit = Format(frm130.L8_Text, "0.00") 'Cara Bayaran : Kadar GST bagi kad kredit
                    Else
                        rs!kadar_gst_kad_kredit = "0.00" 'Cara Bayaran : Kadar GST bagi kad kredit
                    End If
                    rs!epp = 0 '0 : Bayaran selain dari EPP , 1 : Bayaran secara EPP
                    rs!approval_code_epp = Null 'Approval Code (EPP)
                    
                End If
            End If

            If frm130.TB21 <> vbNullString Then
                If Format(frm130.TB21, "0.00") <> "0.00" Then
                    Frm92_LM_Flag_SIMPANAN = 1 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
                End If
                rs!duit_simpanan_kedai = Format(frm130.TB21, "0.00") 'Cara Bayaran : Simpanan Duit Di Kedai
            Else
                rs!duit_simpanan_kedai = "0.00" 'Cara Bayaran : Simpanan Duit Di Kedai
            End If
            If frm130.TB32 <> vbNullString Then
                rs!jumlah_bayaran = Format(frm130.TB32, "0.00") 'Cara Bayaran : Jumlah Bayaran
            Else
                rs!jumlah_bayaran = Null 'Cara Bayaran : Jumlah Bayaran
            End If
            If Frm92.L9_Text <> vbNullString Then
                rs!harga_barang = Format(Frm92.L9_Text, "0.00") 'Jumlah Harga Barang Tanpa GST (RM)
            Else
                rs!harga_barang = Null 'Jumlah Harga Barang Tanpa GST (RM)
            End If
            If Frm92.L15_Text <> vbNullString Then 'Kadar cukai GST
                rs!kadar_gst = Format(Frm92.L15_Text, "0.00")
            Else
                rs!kadar_gst = Null
            End If
            If Frm92.L7_Text <> vbNullString Then
                rs!jumlah_cukai_gst = Format(Frm92.L7_Text, "0.00") 'Jumlah Cukai GST (ZR + SR)
            Else
                rs!jumlah_cukai_gst = Null 'Jumlah Cukai GST (ZR + SR)
            End If
            If Frm92.L10_Text <> vbNullString Then
                rs!harga_barang_dengan_gst = Format(Frm92.L10_Text, "0.00") 'Jumlah Harga Barang Dengan GST (RM)
                rs!harga_lepas_diskaun = Format(Frm92.L10_Text, "0.00") 'Harga Selepas Diskaun (RM)
                rs!harga_jualan = Format(Frm92.L10_Text, "0.00") 'Jumlah Harga Jualan (RM)
                rs!jumlah_perlu_bayar = Format(Frm92.L10_Text, "0.00") 'Jumlah Bayaran Yang Perlu Dibuat (RM)
            Else
                rs!harga_barang_dengan_gst = Null 'Jumlah Harga Barang Dengan GST (RM)
                rs!harga_lepas_diskaun = Null 'Harga Selepas Diskaun (RM)
                rs!harga_jualan = Null 'Jumlah Harga Jualan (RM)
                rs!jumlah_perlu_bayar = Null 'Jumlah Bayaran Yang Perlu Dibuat (RM)
            End If
            rs!diskaun = Null 'Jumlah Diskaun (%)
            rs!adjustment = "0.00" 'Adjustment (RM)
            rs!loss_trade_in = Null 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (%)
            rs!loss_trade_in_rm = Null 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (RM)
            rs!flag_bayaran = 0 '0 : Pembeli Bayar , 1 : Kedai Bayar
            If Frm92.L20_Text <> vbNullString Then 'Kuantiti servis
                rs!kuantiti_barang = Frm92.L20_Text
            Else
                rs!kuantiti_barang = Null
            End If
            rs!JUMLAH_BERAT = Null 'Jumlah Berat Barang Yang Dijual
            If Frm92.L11_Text <> vbNullString Then
                rs!gst_zr_harga = Format(Frm92.L11_Text, "0.00") 'Harga Keseluruhan Bagi Barang ZR
            Else
                rs!gst_zr_harga = Null 'Harga Keseluruhan Bagi Barang ZR
            End If
            If Frm92.L12_Text <> vbNullString Then
                rs!gst_zr_cukai = Format(Frm92.L12_Text, "0.00") 'Jumlah Cukai Bagi ZR
            Else
                rs!gst_zr_cukai = Null 'Jumlah Cukai Bagi ZR
            End If
            If Frm92.L13_Text <> vbNullString Then
                rs!gst_sr_harga = Format(Frm92.L13_Text, "0.00") 'Harga Keseluruhan Bagi Barang SR
            Else
                rs!gst_sr_harga = Null 'Harga Keseluruhan Bagi Barang SR
            End If
            If Frm92.L14_Text <> vbNullString Then
                rs!gst_sr_cukai = Format(Frm92.L14_Text, "0.00") 'Jumlah Cukai Bagi SR
            Else
                rs!gst_sr_cukai = Null 'Jumlah Cukai Bagi SR
            End If
            rs!caj_pos = "0.00"
            rs!no_tracking = Null

            rs!no_pekerja = Frm92_LM_EMP_NO 'No. Pekerja
            rs!nama_pekerja = Frm92_LM_EMP_NAMA
            
            If Frm92.L52_Text <> vbNullString Then
                If Frm28.L5_Text <> vbNullString Then
                    rs!no_rujukan_pembeli = Frm28.L5_Text 'No. Rujukan Pembeli
                    LM_NO_PELANGGAN = Frm28.L5_Text 'No. Rujukan Pembeli
                Else
                    rs!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
                End If
            Else
                rs!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
            End If
            rs!no_rujukan_agen_dropship = Null 'No. Rujukan Agen Dropship
            rs!flag_trade_in = 0 '0 : Tiada Urusan Trade in , 1 : Ada Urusan Trade In
            rs!no_resit_trade_in = Null 'No. Resit Trade In
            rs!jumlah_trade_in = Null 'Jumlah Resit Trade In (RM)
            rs!jenis_trade_in = Null '1 : Trade in (Voucher) , 2 : Belian dengan trade in
            rs!invoice_type = 0 '0 : Unlimited , Selain 0 (Limited : Mengikut nombor yang dimasukkan)

'Zakaria&Sons
'1 : Pembeli biasa
'2 : Ahli biasa
'3 : Silver
'4 : Gold
'5 : Platinum

            rs!kategori_pembeli = Frm92_LM_KATEGORI
            rs!jualan_online = 0
            rs!kupon_diskaun = "0.00"
            rs!terminal = G_TERMINAL
            rs!no_staff = G_LOGIN_USER
            rs!write_timestamp2 = LM_NOW
            
            DATA_SAVE = 1
            rs.Update
            
        End If
        
        rs.Close
        Set rs = Nothing
'###Masukkan Jualan Ke Dalam Table Akaun Jualan### - End
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "insert into 35_senarai_servis(tarikh,no_resit_servis,no_pelanggan,write_timestamp,terminal,no_pekerja,detail,jumlah,jenis_gst,jumlah_gst,kod_gst,harga_tanpa_gst,harga_dengan_gst)" & _
                    "select '" & Frm92.DTPicker1 & "','" & G_No_RESIT_SERVIS & "','" & LM_NO_PELANGGAN & "','" & LM_NOW & "','" & G_TERMINAL & "','" & Frm92_LM_EMP_NO & "',detail,jumlah,jenis_gst,jumlah_gst,kod_gst,harga_tanpa_gst,harga_dengan_gst from " & G_SERVICE_TEMP & ""
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing

'### Maklumat Pembeli Yang Tidak Berdaftar Dengan Kedai ### - Start (08-07-2015)
        If Frm92.L51_Text <> vbNullString Then
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 44_senarai_pelanggan", cn, adOpenKeyset, adLockOptimistic
        
            rs.AddNew
            rs!tarikh = Frm92.DTPicker1 'Tarikh
            rs!no_resit = G_No_RESIT_SERVIS 'No. Resit Trade In
            If Frm26.TB1 <> vbNullString Then 'Nama
                rs!Nama = UCase(Frm26.TB1)
            Else
                rs!Nama = Null
            End If
            If Frm26.TB2 <> vbNullString Then 'No. Telefon
                rs!no_tel = UCase(Frm26.TB2)
            Else
                rs!no_tel = Null
            End If
            rs!write_timestamp = LM_NOW
            rs!no_staff = Frm92_LM_EMP_NO 'No. Pekerja
            rs!terminal = G_TERMINAL
            rs!jenis_urusan = G_JENIS_URUSAN
            rs!cawangan = G_CAWANGAN
            rs.Update
            
            rs.Close
            Set rs = Nothing
            
        End If
'### Maklumat Pembeli Yang Tidak Berdaftar Dengan Kedai ### - End

'### Update Log ### - Start
        user = MDI_frm1.L3_Text
        LogAct_Memory = "[" & G_LOGIN_USER & "] Edit servis kepada pelanggan. No. Invoice [" & G_No_RESIT_SERVIS & "]"
        LogDate_Memory = DateTime.Date & " " & DateTime.Time$
        Call UpdateLog_Database
'### Update Log ### - End

'### Update Simpanan ### - Start
        If Frm92_LM_Flag_SIMPANAN = 1 Then '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                Frm92_LM_JUMLAH_SIMPANAN = frm130.L26_Text  'Jumlah Simpanan Yang Ada
                Frm92_LM_GUNA_SIMPAN = frm130.TB21  'Jumlah Simpanan Yang Hendak Digunakan
                
                rs!baki_simpanan = Format(Frm92_LM_JUMLAH_SIMPANAN - Frm92_LM_GUNA_SIMPAN, "0.00") 'Baki Simpanan
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 24_rekod_kewangan_pelanggan", cn, adOpenKeyset, adLockOptimistic
            
            rs.AddNew
            rs!tarikh = Frm92.DTPicker1 'Tarikh
            rs!jenis = 1 '0 : Simpanan , 1 : Penggunaan Duit
            rs!no_rujukan_pelanggan = Frm28.L5_Text 'No. Rujukan Pelanggan
            rs!no_resit = G_No_RESIT_SERVIS 'No. Resit Servis
            rs!jumlah = Format(frm130.TB21, "0.00") 'Jumlah Simpanan Yang Digunakan (RM)
            rs!jenis_penggunaan = 3 '0 : Belian Barangan Kemas , 1 : Ansuran , 2 : Tempahan (Deposit) , 3 : Servis , 4 : Tempahan (Ambilan Barang)
            rs!cawangan = G_CAWANGAN
            rs!Status = 1
            rs.Update
            
            rs.Close
            Set rs = Nothing
        
        End If
'### Update Simpanan ### - End
        
        Call Frm92_Initial_Setting
        Frm92.Frame1.Visible = False
        
        GM_NEXT_PREV = 2
        
        Call frm92_senarai_servis_header
        Call frm92_senarai_servis
        
        G_PREVIEW = 1
        
        Note = "Data telah berjaya disimpan." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Cetak resit servis ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
            Call Frm92_Resit_Servis
        End If
        
    End If
End If
End Sub

Private Sub CMD6_Click()
'on error resume next
frm130.TB33 = Format(Frm92.L10_Text, "#,##0.00")
frm130.Show vbModal
End Sub

Private Sub CMD8_Click()
'on error resume next
Note = "Adakah anda ingin batalkan urusan edit data ini?" & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then
    Frm92.Frame1.Visible = False
    Frm92.Frame4.Visible = True
End If
End Sub


Private Sub Form_Load()
'on error resume next
Frm92.L22_Text = 0
Frm92.L23_Text = "0.00"

Frm92.L24_Text = 0
Frm92.L60_Text = 0 'Current Page
Frm92.L61_Text = 0 'Total Page
Frm92.L64_Text = 0 'Jenis Report , 0 : Keseluruhan , 1 : Filter Ikut Tarikh
Frm92.L65_Text = 0 'Jenis Report , 0 : Keseluruhan , 1 : Details
Frm92.L54_Text = 1
Frm92.L25_Text = DateTime.Date
Frm92.L26_Text = DateTime.Date
Frm92.CB16 = 1

Frm92.L8_Text.BackStyle = 0
Frm92.L50_Text.BackStyle = 0
Frm92.L55_Text.BackStyle = 0

user = MDI_frm1.L3_Text

user_level = MDI_frm1.L4_Text

If user_level = "Admin" Or user_level = "HQ" Or user_level = "Developer" Then

    Frm92.Frm92_SM_edit.Enabled = True
    Frm92.Frm92_SM_edit2.Enabled = True
    Frm92.Frm92_SM_edit3.Enabled = True
    Frm92.Frm92_SM_edit4.Enabled = True
    
    Frm92.Frm92_SM_padam.Enabled = True
    Frm92.Frm92_SM_padam2.Enabled = True
    Frm92.Frm92_SM_padam3.Enabled = True
    Frm92.Frm92_SM_padam4.Enabled = True
    
ElseIf user_level = "Manager" Then

    Frm92.Frm92_SM_edit.Enabled = True
    Frm92.Frm92_SM_edit2.Enabled = True
    Frm92.Frm92_SM_edit3.Enabled = True
    Frm92.Frm92_SM_edit4.Enabled = True
    
    Frm92.Frm92_SM_padam.Enabled = False
    Frm92.Frm92_SM_padam2.Enabled = False
    Frm92.Frm92_SM_padam3.Enabled = False
    Frm92.Frm92_SM_padam4.Enabled = False

Else

    Frm92.Frm92_SM_edit.Enabled = False
    Frm92.Frm92_SM_edit2.Enabled = False
    Frm92.Frm92_SM_edit3.Enabled = False
    Frm92.Frm92_SM_edit4.Enabled = False
    
    Frm92.Frm92_SM_padam.Enabled = False
    Frm92.Frm92_SM_padam2.Enabled = False
    Frm92.Frm92_SM_padam3.Enabled = False
    Frm92.Frm92_SM_padam4.Enabled = False

End If

'Set rs = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from tblelogin where username='" & User & "'", cn, adOpenKeyset, adLockOptimistic

'If Not rs.EOF Then
'    If Not IsNull(rs!usertype) Then
'        If rs!usertype = "Developer" Or rs!usertype = "Admin" Then
'            Frm92.Frm92_SM_padam2.Enabled = True
'        Else
'            Frm92.Frm92_SM_padam2.Enabled = False
'        End If
'    End If
'End If

'rs.Close
'Set rs = Nothing
End Sub



Private Sub frm92_sm_cetak_pv_Click()
'on error resume next
LM_DATA_FOUND = 0
LM_FOUND = 0

frm92_LM_No_ID = vbNullString

If IsNumeric(Frm92.LV3.SelectedItem.Index) Then
    
    frm92_LM_No_ID = Frm92.LV3.ListItems(Frm92.LV3.SelectedItem.Index)
    
    If frm92_LM_No_ID <> vbNullString Then
    
        Set rs = New ADODB.Recordset
        Call Main
        rs.Open "select * from 39_akaun_expense where ID='" & frm92_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            If Not IsNull(rs!cawangan) Then G_KEDAI = rs!cawangan
            
            If Not IsNull(rs!no_voucher) Then
                
                G_No_RESIT_JUALAN = rs!no_voucher
                LM_DATA_FOUND = 1
            
            End If
            
        End If
        
        rs.Close
        Set rs = Nothing
            
        
        If LM_DATA_FOUND = 1 Then
            
            G_PREVIEW = 1
            Call frm92_cetak_pv
            
        End If
            
    End If
    
End If
End Sub

Private Sub Frm92_SM_cetak_resit_Click()
'on error resume next
DATA_FOUND = 0

frm92_LM_No_ID = vbNullString

If IsNumeric(Frm92.LV2.SelectedItem.Index) Then
    
    frm92_LM_No_ID = Frm92.LV2.ListItems(Frm92.LV2.SelectedItem.Index)
    
    If frm92_LM_No_ID <> vbNullString Then
        
        Note = "Adakah anda ingin cetak invoice ini ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
        
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 22_jualan where ID='" & frm92_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!no_resit) Then G_No_RESIT_SERVIS = rs!no_resit
                If Not IsNull(rs!cawangan) Then G_KEDAI = rs!cawangan
                DATA_FOUND = 1
                
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_FOUND = 1 Then
            
                G_PREVIEW = 1
    
                Call Frm92_Resit_Servis
                
            Else
                
                MsgBox "Tiada data invoice dijumpai.", vbExclamation, "Info"
                
            End If
            
        End If
        
    End If
End If
End Sub
Private Sub Frm92_SM_edit_Click()
'on error resume next
DATA_FOUND = 0

DATA_FOUND = 0

frm92_LM_No_ID = vbNullString

If IsNumeric(Frm92.LV1.SelectedItem.Index) Then
    
    frm92_LM_No_ID = Frm92.LV1.ListItems(Frm92.LV1.SelectedItem.Index)
    
    If frm92_LM_No_ID <> vbNullString Then
        
        Note = "Adakah anda ingin edit data ini ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
        
            Frm92.TB1 = vbNullString
            Frm92.TB2 = "0.00"
            Frm92.CB1 = 1
            Frm92.CB2 = 0
            Frm92.CB8 = 0
            Frm92.L8_Text = "0.00"
            Frm92.L50_Text = "0.00"
            Frm92.L55_Text = "0.00"
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from " & G_SERVICE_TEMP & " where ID='" & frm92_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
            
                If Not IsNull(rs!ID) Then Frm92.L16_Text = rs!ID
                If Not IsNull(rs!Detail) Then Frm92.TB1 = rs!Detail 'Details
                If Not IsNull(rs!jumlah) Then Frm92.TB2 = Format(rs!jumlah, "0.00") 'Jumlah (RM)
                If Not IsNull(rs!kod_gst) Then
                
                    If rs!kod_gst = 0 Then
                    
                        Frm92.CB1 = 1
                        Frm92.CB2 = 0
                        Frm92.CB8 = 0
                        
                    ElseIf rs!kod_gst = 1 Then
                    
                        Frm92.CB1 = 0
                        Frm92.CB2 = 1
                        Frm92.CB8 = 0
                        
                    ElseIf rs!kod_gst = 2 Then

                        Frm92.CB1 = 0
                        Frm92.CB2 = 0
                        Frm92.CB8 = 1

                    End If
                
                End If
                If Not IsNull(rs!harga_tanpa_gst) Then Frm92.L50_Text = Format(rs!harga_tanpa_gst, "0.00") 'Harga Keseluruhan Tanpa GST (RM)
                If Not IsNull(rs!jumlah_gst) Then Frm92.L8_Text = Format(rs!jumlah_gst, "0.00") 'Jumlah cukai GST (RM)
                If Not IsNull(rs!harga_dengan_gst) Then Frm92.L55_Text = Format(rs!harga_dengan_gst, "0.00") 'Harga keseluruhan dengan GST (RM)

            End If
            
            rs.Close
            Set rs = Nothing
            
            Frm92.CMD1.Visible = False
            Frm92.CMD2.Visible = True
            Frm92.CMD3.Visible = True
            Frm92.L18_Text.Visible = True
            
        End If
        
    End If
    
End If
End Sub
Private Sub Frm92_SM_edit2_Click()
'on error resume next
Dim Frm92_LM_SIMPANAN_ASAL As Double
Dim Frm92_LM_SIMPANAN_DIGUNAKAN As Double

Frm92_LM_No_RUJUKAN_PEMBELI = vbNullString
Frm92_LM_No_PEKERJA = vbNullString

Frm92_LM_No_RESIT_SERVIS = vbNullString
Frm92_LM_Flag_SIMPANAN = 0 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
Frm92_LM_SIMPANAN_ASAL = 0
Frm92_LM_SIMPANAN_DIGUNAKAN = 0
Frm92_LM_KATEGORI_PEMBELI = 0 '0 : Pembeli Tidak Berdaftar , 1 : Pembeli Berdaftar , 2 : Ahli
DATA_FOUND = 0

frm92_LM_No_ID = vbNullString

If IsNumeric(Frm92.LV2.SelectedItem.Index) Then
    
    frm92_LM_No_ID = Frm92.LV2.ListItems(Frm92.LV2.SelectedItem.Index)
    
    If frm92_LM_No_ID <> vbNullString Then
        
        Note = "Adakah anda ingin edit data ini ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
        
            Call Frm92_Initial_Setting
            Call frm130_initial_setting
            
            Frm92.L17_Text = Frm92_LM_No_RESIT_SERVIS 'No. Resit Servis
        
'### Update Akaun Bagi Servis ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 22_jualan where ID='" & frm92_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                GLOBAL_DISABLE = 1
                frm130.L41_Text = "1"
                
                If Not IsNull(rs!no_resit) Then
                
                    Frm92.L17_Text = rs!no_resit
                    Frm92_LM_No_RESIT_SERVIS = rs!no_resit
                    
                End If
                
                If Not IsNull(rs!bil_rasmi) Then
                
                    If rs!bil_rasmi = 0 Then
                        Frm92.CB9 = 1
                    Else
                        Frm92.CB9 = 0
                    End If
                
                End If
                If Not IsNull(rs!tunai) Then frm130.TB27 = Format(rs!tunai, "#,##0.00") 'Cara Bayaran : Tunai
                If Not IsNull(rs!bank_in) Then frm130.TB28 = Format(rs!bank_in, "#,##0.00") 'Cara Bayaran : Bank In
                
                If Not IsNull(rs!kad_kredit) Then 'Cara Bayaran : Kad Kredit
                    frm130.TB29 = Format(rs!kad_kredit, "#,##0.00")
                Else
                    frm130.TB29 = "0.00"
                End If
            
                'On Error GoTo Err_B:
                If Not IsNull(rs!jenis_kad) Then
                    Frm92_LM_JENIS_KAD = rs!jenis_kad
                    frm130.CBB2 = Frm92_LM_JENIS_KAD
                    
Restore_B:
                End If
                'on error resume next
                
                If Not IsNull(rs!cas_Kad_Kredit) Then 'Cara Bayaran : Cas Kad Kredit (%)
                    frm130.L31_Text = Format(rs!cas_Kad_Kredit, "#,##0.00")
                Else
                    frm130.L31_Text = Format(0, "#,##0.00")
                End If
                If Not IsNull(rs!jumlah_cas_kad_kredit) Then 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
                    frm130.L32_Text = Format(rs!jumlah_cas_kad_kredit, "#,##0.00")
                Else
                    frm130.L32_Text = Format(0, "#,##0.00")
                End If
                If Not IsNull(rs!gst_kad_kredit) Then 'Cara Bayaran : Jumlah GST kad kredit (RM)
                    frm130.L81_Text = Format(rs!gst_kad_kredit, "#,##0.00")
                Else
                    frm130.L81_Text = Format(0, "#,##0.00")
                End If
                If Not IsNull(rs!jumlah_potongan_kad_kredit) Then 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
                    frm130.L82_Text = Format(rs!jumlah_potongan_kad_kredit, "#,##0.00")
                Else
                    frm130.L82_Text = Format(0, "#,##0.00")
                End If

                If Not IsNull(rs!duit_simpanan_kedai) Then
                    frm130.TB21 = Format(rs!duit_simpanan_kedai, "#,##0.00") 'Cara Bayaran : Simpanan Duit Di Kedai
                    If IsNumeric(rs!duit_simpanan_kedai) Then Frm92_LM_SIMPANAN_DIGUNAKAN = rs!duit_simpanan_kedai 'Jumlah Simpanan Yang Digunakan
                    If Format(rs!duit_simpanan_kedai, "0.00") <> "0.00" Then
                        Frm92_LM_Flag_SIMPANAN = 1 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
                    End If
                End If
                If Not IsNull(rs!kadar_gst) Then
                    frm130.L8_Text = Format(rs!kadar_gst, "#,##0.00")
                    Frm92.L15_Text = Format(rs!kadar_gst, "#,##0.00")
                Else
                    frm130.L8_Text = Format(0, "#,##0.00")
                    Frm92.L15_Text = Format(0, "#,##0.00")
                End If
                If Not IsNull(rs!harga_barang) Then
                    Frm92.L9_Text = Format(rs!harga_barang, "#,##0.00")
                Else
                    Frm92.L9_Text = Format(0, "#,##0.00")
                End If
                If Not IsNull(rs!jumlah_cukai_gst) Then
                    Frm92.L7_Text = Format(rs!jumlah_cukai_gst, "#,##0.00")
                Else
                    Frm92.L7_Text = Format(0, "#,##0.00")
                End If
                If Not IsNull(rs!harga_barang_dengan_gst) Then
                    Frm92.L10_Text = Format(rs!harga_barang_dengan_gst, "#,##0.00")
                Else
                    Frm92.L10_Text = Format(0, "#,##0.00")
                End If
                
                If Not IsNull(rs!kuantiti_barang) Then 'Kuantiti servis
                    Frm92.L20_Text = rs!kuantiti_barang
                Else
                    Frm92.L20_Text = "0"
                End If
                If Not IsNull(rs!gst_zr_harga) Then Frm92.L11_Text = Format(rs!gst_zr_harga, "#,##0.00")  'Harga Keseluruhan Bagi Barang ZR
                If Not IsNull(rs!gst_zr_cukai) Then Frm92.L12_Text = Format(rs!gst_zr_cukai, "#,##0.00")  'Jumlah Cukai Bagi ZR
                If Not IsNull(rs!gst_sr_harga) Then Frm92.L13_Text = Format(rs!gst_sr_harga, "#,##0.00")  'Harga Keseluruhan Bagi Barang SR
                If Not IsNull(rs!gst_sr_cukai) Then Frm92.L14_Text = Format(rs!gst_sr_cukai, "#,##0.00")  'Jumlah Cukai Bagi SR

                If Not IsNull(rs!no_rujukan_pembeli) Then Frm92_LM_No_RUJUKAN_PEMBELI = rs!no_rujukan_pembeli 'No. Rujukan Pembeli
                If Not IsNull(rs!no_pekerja) Then Frm92_LM_No_PEKERJA = rs!no_pekerja 'No. Pekerja
                DATA_FOUND = 1
                
            End If
            
            rs.Close
            Set rs = Nothing
'### Update Akaun Bagi Servis ### - End

            If Frm92_LM_No_RUJUKAN_PEMBELI = vbNullString Then '0 : Pembeli Tidak Berdaftar , 1 : Pembeli Berdaftar , 2 : Ahli
            
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 44_senarai_pelanggan where no_resit='" & Frm92_LM_No_RESIT_SERVIS & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    
                    Call Frm26_initial
                    
                    If Not IsNull(rs!Nama) Then
                        Frm26.TB1 = rs!Nama 'Nama
                        Frm92.L51_Text = rs!Nama 'Nama
                    End If
                    If Not IsNull(rs!no_tel) Then Frm26.TB2 = rs!no_tel 'No. Telefon
            
                End If
                
                rs.Close
                Set rs = Nothing
            
            Else
            
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm92_LM_No_RUJUKAN_PEMBELI & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    
                    Call Frm28_initial
                    
                    If Not IsNull(rs!Nama) Then
                        Frm28.L1_Text = rs!Nama 'Nama
                        Frm92.L52_Text = rs!Nama 'Nama
                    End If
                    If Not IsNull(rs!no_ic) Then Frm28.L2_Text = rs!no_ic 'No. Kad Pengenalan
                    If Not IsNull(rs!no_tel) Then Frm28.L3_Text = rs!no_tel 'No. Telefon
                    If Not IsNull(rs!Email) Then Frm28.L4_Text = rs!Email 'E-mail
                    If Not IsNull(rs!no_pelanggan) Then Frm28.L5_Text = rs!no_pelanggan 'No. Pelanggan
                    
                    If Not IsNull(rs!baki_simpanan) Then
                        frm130.L26_Text = Format(rs!baki_simpanan, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
                        
                        If IsNumeric(rs!baki_simpanan) Then
                            Frm92_LM_SIMPANAN_ASAL = rs!baki_simpanan 'Jumlah Simpanan Asal Yang Ada (RM)
                            
                            frm130.L26_Text = Format(Frm92_LM_SIMPANAN_ASAL + Frm92_LM_SIMPANAN_DIGUNAKAN, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
                        End If
                    End If

                End If
                
                rs.Close
                Set rs = Nothing
    
            End If

'### Carian Maklumat Penjual (Data Pekerja) ### - Start
            If Frm92_LM_No_PEKERJA <> vbNullString Then
                
                DATA_PEKERJA_FOUND = 0
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from employee where NoPekerja='" & Frm92_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    Frm92_LM_MAKLUMAT_PEKERJA = rs!Samaran & "  |  " & rs!NoPekerja
                    DATA_PEKERJA_FOUND = 1
                End If
                
                rs.Close
                Set rs = Nothing
                
                If DATA_PEKERJA_FOUND = 1 Then
                    'On Error GoTo Err_A:
                    Frm92.CBB1 = Frm92_LM_MAKLUMAT_PEKERJA
Restore_A:
                End If
                
            End If
'### Carian Maklumat Penjual (Data Pekerja) ### - End

            If DATA_FOUND = 1 Then
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "insert into " & G_SERVICE_TEMP & "(id_database,detail,jumlah,jenis_gst,jumlah_gst,kod_gst,harga_tanpa_gst,harga_dengan_gst)" & _
                            "select ID,detail,jumlah,jenis_gst,jumlah_gst,kod_gst,harga_tanpa_gst,harga_dengan_gst from 35_senarai_servis where no_resit_servis='" & Frm92_LM_No_RESIT_SERVIS & "'"
                
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing

                Call frm92_senarai_service_header
                Call frm92_senarai_service
                
                GLOBAL_DISABLE = 0
        
                Frm92.CBB1.Enabled = True
                Frm92.CBB1.BackColor = &HFFFFFF
                
                Frm92.CMD4.Visible = False
                Frm92.CMD5.Visible = True
                Frm92.CMD8.Visible = True
                
                Frm92.CB9.Enabled = False
                Frm92.Frame4.Visible = False
                Frm92.Frame1.Visible = True
            End If
        End If
        
    End If
    
End If

Exit Sub
Err_A:
Frm92.CBB1.AddItem Frm92_LM_MAKLUMAT_PEKERJA
Frm92.CBB1 = Frm92_LM_MAKLUMAT_PEKERJA
Resume Restore_A:

Exit Sub
Err_B:
frm130.CBB2.AddItem Frm92_LM_JENIS_KAD
frm130.CBB2 = Frm92_LM_JENIS_KAD
Resume Restore_B:
End Sub
Private Sub Frm92_SM_edit3_Click()
'on error resume next
LM_DATA_FOUND = 0
LM_FOUND = 0

frm92_LM_No_ID = vbNullString

If IsNumeric(Frm92.LV3.SelectedItem.Index) Then
    
    frm92_LM_No_ID = Frm92.LV3.ListItems(Frm92.LV3.SelectedItem.Index)
    
    If frm92_LM_No_ID <> vbNullString Then
            
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 39_akaun_expense where ID='" & frm92_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            Note = "Adakah anda ingin edit data ini?"
                    
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbNo Then
            
                LM_DATA_FOUND = 2
                
            End If
            If Answer = vbYes Then
                
                LM_FOUND = 1
                
            End If
            
        End If
        
        rs.Close
        Set rs = Nothing
            
        If LM_FOUND = 1 Then
        
            Call Frm92_Initial_Setting
            Call frm92_initial_one_time
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 39_akaun_expense where ID='" & frm92_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!ID) Then Frm92.L43_Text = rs!ID
                
                If Not IsNull(rs!flag_kedai) Then
                
                    If rs!flag_kedai = 0 Then '0 : Supplier tidak berdaftar , 1 : Supplier berdaftar
                        
                        Frm92.CBB5 = "Lain-lain"
                        If Not IsNull(rs!nama_kedai) Then Frm92.TB41 = rs!nama_kedai 'Nama Kedai
                        
                    ElseIf rs!flag_kedai = 1 Then '0 : Supplier tidak berdaftar , 1 : Supplier berdaftar
                        
                        If Not IsNull(rs!nama_kedai) Then Frm92.CBB5 = rs!nama_kedai
                        
                    End If
                    
                End If
                
                If Not IsNull(rs!no_resit) Then Frm92.TB42 = rs!no_resit 'No. Invoice
                If Not IsNull(rs!tujuan) Then Frm92.TB44 = rs!tujuan 'Tujuan
                If Not IsNull(rs!no_id_gst) Then Frm92.TB43 = rs!no_id_gst 'No. ID GST
                If Not IsNull(rs!tarikh) Then Frm92.DTPicker4 = rs!tarikh 'Tarikh
                If Not IsNull(rs!harga_dengan_gst) Then Frm92.TB45 = Format(rs!harga_dengan_gst, "#,##0.00") 'Jumlah Dengan GST (RM)
                If Not IsNull(rs!gst_zr_harga) Then Frm92.TB48 = Format(rs!gst_zr_harga, "#,##0.00") 'Harga Keseluruhan Bagi Barang ZR
                If Not IsNull(rs!gst_zr_cukai) Then Frm92.TB49 = Format(rs!gst_zr_cukai, "#,##0.00") 'Jumlah Cukai Bagi ZR
                If Not IsNull(rs!gst_sr_harga) Then Frm92.TB46 = Format(rs!gst_sr_harga, "#,##0.00") 'Harga Keseluruhan Bagi Barang SR
                If Not IsNull(rs!gst_sr_cukai) Then Frm92.TB47 = Format(rs!gst_sr_cukai, "#,##0.00") 'Jumlah Cukai Bagi SR
                If Not IsNull(rs!gst_value) Then Frm92.L42_Text = rs!gst_value '% Cukai GST

                If Not IsNull(rs!cara_bayaran) Then
                    
                    If rs!cara_bayaran = 0 Then
                    
                        Frm92.CB10 = 1
                        Frm92.CB3 = 0
                        Frm92.CB4 = 0
                        
                    ElseIf rs!cara_bayaran = 1 Then

                        Frm92.CB10 = 0
                        Frm92.CB3 = 1
                        Frm92.CB4 = 0
                        
                    ElseIf rs!cara_bayaran = 2 Then
                    
                        Frm92.CB10 = 0
                        Frm92.CB3 = 0
                        Frm92.CB4 = 1
                        
                    End If
                
                End If
                
                If Not IsNull(rs!no_pekerja) Then Frm92_LM_No_PEKERJA = rs!no_pekerja 'No. Pekerja
                
                If Not IsNull(rs!jenis_expense) Then Frm92_LM_JENIS_EXP = rs!jenis_expense
                    'On Error GoTo Err_B:
                    Frm92.CBB7 = Frm92_LM_JENIS_EXP
Restore_B:
                
                LM_DATA_FOUND = 1
                
            End If
            
            rs.Close
            Set rs = Nothing
            
        End If
            
        If LM_DATA_FOUND = 1 Then
        
'### Carian Maklumat Penjual (Data Pekerja) ### - Start
            If Frm92_LM_No_PEKERJA <> vbNullString Then
                
                DATA_PEKERJA_FOUND = 0
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from employee where NoPekerja='" & Frm92_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    Frm92_LM_MAKLUMAT_PEKERJA = rs!Samaran & "  |  " & rs!NoPekerja
                    DATA_PEKERJA_FOUND = 1
                End If
                
                rs.Close
                Set rs = Nothing
                
                If DATA_PEKERJA_FOUND = 1 Then
                    'On Error GoTo Err_A:
                    Frm92.CBB2 = Frm92_LM_MAKLUMAT_PEKERJA
Restore_A:
                End If
                
            End If
'### Carian Maklumat Penjual (Data Pekerja) ### - End
            
            Frm92.CMD29.Visible = True
            Frm92.CMD30.Visible = True
            Frm92.CMD28.Visible = False
            
            Frm92.CBB2.Enabled = True
            Frm92.CBB2.BackColor = &HFFFFFF
            
            Frm92.Frame5.Visible = True
            Frm92.Frame6.Visible = False
            
        ElseIf LM_DATA_FOUND = 2 Then
            
            MsgBox "Urusan dibatalkan.", vbInformation, "Info"
            
        ElseIf LM_DATA_FOUND = 0 Then
        
            MsgBox "Tiada data dijumpai. Sila keluar dari menu ini dan cuba lagi.", vbExclamation, "Info"
            
        End If
            
    End If
    
End If

Exit Sub
Err_A:
Frm92.CBB2.AddItem Frm92_LM_MAKLUMAT_PEKERJA
Frm92.CBB2 = Frm92_LM_MAKLUMAT_PEKERJA
Resume Restore_A:

Exit Sub
Err_B:
Frm92.CBB7.AddItem Frm92_LM_JENIS_EXP
Frm92.CBB7 = Frm92_LM_JENIS_EXP
Resume Restore_B:
End Sub
Private Sub Frm92_SM_excel_Click()
'on error resume next
If Frm92.L65_Text = 0 Then 'Jenis Report , 0 : Keseluruhan , 1 : Details
    If Frm92.L64_Text = 0 Then 'Jenis Report , 0 : Keseluruhan , 1 : Filter Ikut Tarikh
        Call Frm92_excel_overall
    ElseIf Frm92.L64_Text = 1 Then 'Jenis Report , 0 : Keseluruhan , 1 : Filter Ikut Tarikh
        Call Frm92_excel_overall_tarikh
    End If
Else
    If Frm92.L64_Text = 0 Then 'Jenis Report , 0 : Keseluruhan , 1 : Filter Ikut Tarikh
        Call Frm92_excel_detail
    ElseIf Frm92.L64_Text = 1 Then 'Jenis Report , 0 : Keseluruhan , 1 : Filter Ikut Tarikh
        Call Frm92_excel_detail_tarikh
    End If
End If
End Sub

Private Sub frm92_sm_excel1_Click()
'on error resume next
LM_DATA_FOUND = 0
LM_FOUND = 0

frm92_LM_No_ID = vbNullString

If IsNumeric(Frm92.LV3.SelectedItem.Index) Then
    
    frm92_LM_No_ID = Frm92.LV3.ListItems(Frm92.LV3.SelectedItem.Index)
    
    If frm92_LM_No_ID <> vbNullString Then
    
        Note = "Adakah anda ingin export semua data ini ke excel?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Sistem mungkin mengambil masa untuk export semua data ini." & vbCrLf & _
                "Sila tunggu sehingga sistem selesai export data ini." & vbCrLf & _
                "Teruskan?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

        If Answer = vbYes Then


            Set xlObject = New Excel.Application
            Set xlWB = xlObject.Workbooks.Add
                       
            'xlObject.Visible = True
            With xlObject.ActiveWorkbook.ActiveSheet
            
                TM = Frm92.L76_Text 'Tarikh Mula
                TA = Frm92.L77_Text 'Tarikh Akhir

                .Cells.VerticalAlignment = xlCenter
                .Columns("A").ColumnWidth = 5 'No.
                .Columns("B").ColumnWidth = 15 'Tarikh
                .Columns("C").ColumnWidth = 20 'No. invoice
                .Columns("D").ColumnWidth = 25 'No. Voucher
                .Columns("E").ColumnWidth = 60 'Nama Kedai
                .Columns("F").ColumnWidth = 20 'No ID GST
                .Columns("G").ColumnWidth = 80 'Tujuan
                .Columns("H").ColumnWidth = 15 'Jumlah (RM)
                .Columns("I").ColumnWidth = 15 'Jumlah GST (RM)
                .Columns("J").ColumnWidth = 20
                .Columns("K").ColumnWidth = 30

                If Frm92.L82_Text = "Semua jenis" Then 'Jenis
                    frm92_LM_SEARCH_1 = Null
                    frm92_LM_SEARCH_1_LOGIC = "<>"
                Else
                    frm92_LM_SEARCH_1 = Frm92.L82_Text
                    frm92_LM_SEARCH_1_LOGIC = "="
                End If
                If Frm92.L83_Text = "Semua cawangan" Then
                    frm92_LM_SEARCH_2 = Null
                    frm92_LM_SEARCH_2_LOGIC = "<>"
                Else
                    frm92_LM_SEARCH_2 = Frm92.L83_Text
                    frm92_LM_SEARCH_2_LOGIC = "="
                End If

                '### Maklumat kedai ### - Start
                If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
                    
                    LM_NAMA_HEADER = "HQ"
                    
                Else
                    
                    LM_NAMA_HEADER = MDI_frm1.L20_Text
                    
                End If
                        
                '### Maklumat kedai ### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 56_maklumat_kedai where cawangan='" & LM_NAMA_HEADER & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Not IsNull(rs!nama_kedai) Then
                        .Cells(1, 5) = rs!nama_kedai
                        .Cells(1, 5).Font.Name = "Times New Roman"
                    End If
                    If Not IsNull(rs!no_pendaftaran) Then .Cells(2, 5) = rs!no_pendaftaran
                    If Not IsNull(rs!alamat) Then .Cells(3, 5) = rs!alamat
                    If Not IsNull(rs!no_tel) Then .Cells(4, 5) = rs!no_tel
                    If Not IsNull(rs!no_id_gst) Then .Cells(5, 5) = rs!no_id_gst
                End If
                
                rs.Close
                Set rs = Nothing
                '### Maklumat kedai ### - End
                
                x = 0
            
                .Cells(1, 5).Font.Bold = True
                .Cells(1, 5).Font.Size = 30
                
                For Row = 1 To 5
                    .Cells(Row, 5).HorizontalAlignment = xlCenter
                Next Row
                
                .Cells(7, 1) = Frm92.L46_Text
                
                .Cells(8, 1) = "No."
                .Cells(8, 2) = "Tarikh"
                .Cells(8, 3) = "No. invoice"
                .Cells(8, 4) = "No. voucher"
                .Cells(8, 5) = "Nama Kedai"
                .Cells(8, 6) = "No ID GST"
                .Cells(8, 7) = "Tujuan"
                .Cells(8, 8) = "Jumlah (RM)"
                .Cells(8, 9) = "Jumlah GST (RM)"
                .Cells(8, 10) = "Jenis"
                .Cells(8, 11) = "Cawangan"
                
                For i = 1 To 11
                    .Cells(8, i).HorizontalAlignment = xlCenter
                    .Cells(8, i).Interior.ColorIndex = 15
                    .Cells(8, i).WrapText = True
                    .Cells(8, i).Borders.LineStyle = xlContinuous
                Next i
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                If Frm92.L78_Text = 0 Then rs.Open "select * from 39_akaun_expense where cawangan " & frm92_LM_SEARCH_2_LOGIC & "'" & frm92_LM_SEARCH_2 & "' AND jenis_expense " & frm92_LM_SEARCH_1_LOGIC & "'" & frm92_LM_SEARCH_1 & "' AND menu = 1 AND status = 1 order by tarikh ASC , ID ASC", cn, adOpenKeyset, adLockOptimistic
                If Frm92.L78_Text = 1 Then rs.Open "select * from 39_akaun_expense where cawangan " & frm92_LM_SEARCH_2_LOGIC & "'" & frm92_LM_SEARCH_2 & "' AND jenis_expense " & frm92_LM_SEARCH_1_LOGIC & "'" & frm92_LM_SEARCH_1 & "'AND menu = 1 AND status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC , ID ASC", cn, adOpenKeyset, adLockOptimistic

                While rs.EOF = False
                
                    x = x + 1
                    .Cells(8 + x, 1) = x 'No.
                    .Cells(8 + x, 1).HorizontalAlignment = xlCenter

                    If Not IsNull(rs!tarikh) Then .Cells(8 + x, 2) = "'" & rs!tarikh 'Tarikh
                    .Cells(8 + x, 2).HorizontalAlignment = xlCenter
                
                    If Not IsNull(rs!no_resit) Then .Cells(8 + x, 3) = rs!no_resit 'No. invoice
                    .Cells(8 + x, 3).HorizontalAlignment = xlCenter
                    
                    If Not IsNull(rs!no_voucher) Then .Cells(8 + x, 4) = rs!no_voucher 'No. voucher
                    .Cells(8 + x, 4).HorizontalAlignment = xlCenter
                    
                    If Not IsNull(rs!nama_kedai) Then .Cells(8 + x, 5) = rs!nama_kedai 'Nama Kedai
                    
                    If Not IsNull(rs!no_id_gst) Then .Cells(8 + x, 6) = rs!no_id_gst 'No ID GST
                    
                    If Not IsNull(rs!tujuan) Then .Cells(8 + x, 7) = rs!tujuan 'Tujuan
                    
                    .Cells(8 + x, 8).HorizontalAlignment = xlRight
                    If Not IsNull(rs!harga_dengan_gst) Then
                        .Cells(8 + x, 8) = Format(rs!harga_dengan_gst, "#,##0.00") 'Jumlah (RM)
                        .Cells(8 + x, 8).NumberFormat = "#,##0.00"
                    End If
                    
                    .Cells(8 + x, 9).HorizontalAlignment = xlRight
                    If Not IsNull(rs!gst_sr_cukai) Then
                        .Cells(8 + x, 9) = Format(rs!gst_sr_cukai, "#,##0.00") 'Jumlah GST (RM)
                        .Cells(8 + x, 9).NumberFormat = "#,##0.00"
                    End If
                                  
                    If Not IsNull(rs!jenis_expense) Then .Cells(8 + x, 10) = rs!jenis_expense 'Jenis
                    If Not IsNull(rs!cawangan) Then .Cells(8 + x, 11) = rs!cawangan 'Cawangan
                    
                    For Col = 1 To 11
                        .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
                    Next Col
                        
                    rs.MoveNext
                    
                Wend
                
                rs.Close
                Set rs = Nothing
                
                Y = 1
                Y = x + 1
                
                .Cells(8 + Y, 1) = Frm92.L79_Text
                .Cells(8 + Y, 1).Font.Bold = True
                
                Y = Y + 1
                
                '.Cells(8 + Y, 1).HorizontalAlignment = xlRight 'Jumlah berat
                .Cells(8 + Y, 1) = Frm92.L80_Text
                .Cells(8 + Y, 1).Font.Bold = True
                
                Y = Y + 3
                .Cells(8 + Y, 1).Font.Bold = True
                .Cells(8 + Y, 1) = "Report Generated By Sankyu System" 'Watermark Sankyu System
                Y = Y + 1
                .Cells(8 + Y, 1).Font.Bold = True
                .Cells(8 + Y, 1) = "Sankyu System , +6010 - 900 4788 , sankyusystem@gmail.com" 'Watermark Sankyu System
            End With
                
            ' This makes Excel visible
            xlObject.Visible = True
            xlObject.EnableEvents = True
    
        End If
        
    End If
    
End If
End Sub

Private Sub Frm92_SM_padam2_Click()
'On Error Resume Next
Dim Frm92_LM_FLAG_SIMPANAN_ASAL As Double
Dim Frm92_LM_JUMLAH As Double
Dim Frm92_LM_BAKI_ASAL As Double

Frm92_LM_FLAG_SIMPANAN_ASAL = 0
Frm92_LM_JUMLAH = 0
Frm92_LM_BAKI_ASAL = 0
LM_DELETE = 0

Frm92_LM_No_RESIT_SERVIS = vbNullString
        
frm92_LM_No_ID = vbNullString

If IsNumeric(Frm92.LV2.SelectedItem.Index) Then
    
    frm92_LM_No_ID = Frm92.LV2.ListItems(Frm92.LV2.SelectedItem.Index)
    
    If frm92_LM_No_ID <> vbNullString Then
        
        Note = "Adakah anda ingin padam data ini ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
            
            LM_NOW = Now
            G_JENIS_URUSAN = 6
            
            GoTo skip_carian_user:
            
            If MDI_frm1.L3_Text <> vbNullString Then
            
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                
                    If Not IsNull(rs!NoPekerja) Then G_LOGIN_USER = rs!NoPekerja
    
                End If
                
                rs.Close
                Set rs = Nothing
                
            End If
            
skip_carian_user:
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 22_jualan where ID='" & frm92_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
            
                G_ID = rs!ID
                Call recovery_22_jualan
                
                If Not IsNull(rs!no_resit) Then G_No_RESIT_SERVIS = rs!no_resit
                
                rs!Status = 0
                rs!terminal = G_TERMINAL
                rs!no_staff = G_LOGIN_USER
                rs!write_timestamp2 = LM_NOW
                rs.Update
                
                LM_DELETE = 1
        
            End If
            
            rs.Close
            Set rs = Nothing
            
            If LM_DELETE = 1 Then
            
'### Recovery ### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
                strsql = "insert into " & G_RECOVERY_DATABASE & ".35_senarai_servis" & "(id_asal,tarikh,no_resit_servis,no_pelanggan,write_timestamp,terminal,detail,jumlah,jenis_gst," _
                            & "write_writestamp2,terminal2,no_staff," _
                            & "jumlah_gst,kod_gst,harga_tanpa_gst,harga_dengan_gst)" & _
                            "select ID,tarikh,no_resit_servis,no_pelanggan,write_timestamp,terminal,detail,jumlah,jenis_gst," _
                            & "'" & LM_NOW & "','" & G_TERMINAL & "','" & G_LOGIN_USER & "'," _
                            & "jumlah_gst,kod_gst,harga_tanpa_gst,harga_dengan_gst " _
                            & "from " & G_SERVER_DATABASE & ".35_senarai_servis WHERE no_resit_servis='" & G_No_RESIT_SERVIS & "'"
                                        
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
'### Recovery ### - End

'### Padam Data Senarai Servis ### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "DELETE from 35_senarai_servis WHERE no_resit_servis='" & G_No_RESIT_SERVIS & "'"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
'### Padam Data Senarai Servis ### - End
            
'### Padam Penggunaan Duit Pelanggan ### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 24_rekod_kewangan_pelanggan where no_resit='" & G_No_RESIT_SERVIS & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    
                    G_ID = rs!ID
                    Call recovery_24_rekod_kewangan_pelanggan
                        
                    Frm92_LM_FLAG_SIMPANAN_ASAL = 1
                    If Not IsNull(rs!no_rujukan_pelanggan) Then Frm92_LM_No_PELANGGAN = rs!no_rujukan_pelanggan 'No. Rujukan Pelanggan
                    If Not IsNull(rs!jumlah) Then
                        If IsNumeric(rs!jumlah) Then Frm92_LM_JUMLAH = rs!jumlah
                    End If
                    
                    rs.Delete
                    rs.Update
                    
                End If
                
                rs.Close
                Set rs = Nothing
        
                If Frm92_LM_FLAG_SIMPANAN_ASAL = 1 Then
                
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm92_LM_No_PELANGGAN & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                        
                        G_ID = rs!ID
                        Call recovery_senarai_pelanggan
                        
                        If Not IsNull(rs!baki_simpanan) Then Frm92_LM_BAKI_ASAL = rs!baki_simpanan 'Baki Simpanan Asal
                        
                        rs!baki_simpanan = Format(Frm92_LM_BAKI_ASAL + Frm92_LM_JUMLAH, "0.00") 'Baki Simpanan
                        
                        rs!write_timestamp2 = LM_NOW
                        rs!no_staff = G_LOGIN_USER 'No. Pekerja
                        rs!terminal = G_TERMINAL
                        rs!jenis_urusan = G_JENIS_URUSAN
                        rs.Update
                
                    End If
                    
                    rs.Close
                    Set rs = Nothing

                End If
'### Padam Penggunaan Duit Pelanggan ### - End

    '### Padam Rekod Belian Pembeli Yang Tidak Berdaftar Dengan Kedai ### - Start (08-07-2015)
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 44_senarai_pelanggan where no_resit='" & G_No_RESIT_SERVIS & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                            
                    G_ID = rs!ID
                    Call recovery_44_senarai_pelanggan
            
                    rs.Delete
                    rs.Update
                
                End If
                
                rs.Close
                Set rs = Nothing
    '### Padam Rekod Belian Pembeli Yang Tidak Berdaftar Dengan Kedai ### - End (08-07-2015)
    
    '### Update Log ### - Start
                user = MDI_frm1.L3_Text
                LogAct_Memory = "[" & G_LOGIN_USER & "] Padam servis kepada pelanggan. No. Invoice [" & G_No_RESIT_SERVIS & "]"
                LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                Call UpdateLog_Database
    '### Update Log ### - End
                
                GM_NEXT_PREV = 2
                
                Call frm92_senarai_servis_header
                Call frm92_senarai_servis
        
                MsgBox "Invoice bagi servis ini telah berjaya dipadam.", vbInformation, "Info"
        
            End If
            
        End If
            
        
    End If
    
End If
End Sub
Private Sub Frm92_SM_padam3_Click()
'on error resume next
LM_DATA_FOUND = 0
frm92_LM_No_ID = vbNullString

frm92_LM_No_ID = vbNullString

If IsNumeric(Frm92.LV3.SelectedItem.Index) Then
    
    frm92_LM_No_ID = Frm92.LV3.ListItems(Frm92.LV3.SelectedItem.Index)
    
    If frm92_LM_No_ID <> vbNullString Then
            
        LM_NOW = Now
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 39_akaun_expense where ID='" & frm92_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            If Not IsNull(rs!Status) Then
                
                If rs!Status = 1 Then
                
                    Note = "Adakah anda ingin padam data ini?"
                            
                    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
                    If Answer = vbYes Then
                        
                        rs!Status = 0
                        rs.Update
                        LM_DATA_FOUND = 1
                        
                    End If
            
                ElseIf rs!Status = 0 Then
                
                    MsgBox "Tiada perubahan berjaya dilakukan kerana status bagi perbelanjaan ini telah berubah. Sila periksa status terkini data ini.", vbExclamation, "Info"
                    
                End If
                
            End If
            
        End If
        
        rs.Close
        Set rs = Nothing
            
        If LM_DATA_FOUND = 1 Then
            
    '### Update Log ### - Start
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & user & "] Padam perbelanjaan kedai. ID [" & frm92_LM_No_ID & "]"
            LogDate_Memory = LM_NOW
            Call UpdateLog_Database
    '### Update Log ### - End
            
            GM_NEXT_PREV = 2
            
            Call Frm92_report_expenses_header
            Call Frm92_report_expenses
            
            MsgBox "Data perbelanjaan ini telah berjaya dipadamkan.", vbInformation, "Info"
            
        Else
        
            'MsgBox "Tiada data dijumpai. Sila keluar dari menu ini dan cuba lagi.", vbExclamation, "Info"
            
        End If
        
    End If
    
End If
End Sub
Private Sub L10_Text_Change()
'On Error Resume Next
frm130.TB33 = Format(Frm92.L10_Text, "#,##0.00")
End Sub

Private Sub L15_Text_Change()
'On Error Resume Next
Call frm92_kiraan_gst
'Call Frm92_kira_caj_gst_kad_kredit
End Sub



Private Sub L2_Text_Change()
'On Error Resume Next
If Frm92.CMD2.Visible = True Then
    If Frm92.L18_Text.Visible = True Then
        Frm92.L18_Text.Visible = False
    Else
        Frm92.L18_Text.Visible = True
    End If
End If
End Sub



Private Sub L3_Text_Click()
'on error resume next
If Frm92.Frame1.Visible = False Then

    If MDI_frm1.L20_Text = "Semua cawangan" Then
    
        Frm96.CMD2.Visible = True
        Frm96.CMD1.Visible = False
    
        Call Frm96_initial
        
        Frm96.Show vbModal
        
    End If

    Call frm92_pic_visible
    Call Frm92_Initial_Setting
    Call frm130_initial_setting
    
    GLOBAL_DISABLE = 0
    
    Frm92.Frame1.Visible = True
    
    Frm92.TB1.SetFocus
Else
    Frm92.Frame1.Visible = False
End If
End Sub
Private Sub L4_Text_Click()
'on error resume next
If Frm92.Frame3.Visible = False Then

    Call frm92_pic_visible
    Call Frm92_Initial_Setting
    'Call frm92_setting_report
    
    Frm92.CB16 = 0
    
    Frm92.L62_Text = -1 'Start Point
    Frm92.L60_Text = 0 'Current Page
    Frm92.L63_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    
    Frm92.L22_Text = 0
    Frm92.L23_Text = "0.00"

    Frm92.Frame3.Visible = True
    
Else

    Frm92.Frame3.Visible = False
    
End If
End Sub
Private Sub L42_Text_Change()
'on error resume next
Call frm92_kiraan_cukai_sr_belanja
Call frm92_kiraan_cukai_zr_belanja
End Sub
Private Sub L5_Text_Click()
'on error resume next
If Frm92.Frame5.Visible = False Then

    If MDI_frm1.L20_Text = "Semua cawangan" Then
    
        Frm96.CMD2.Visible = True
        Frm96.CMD1.Visible = False
    
        Call Frm96_initial
        
        Frm96.Show vbModal
        
    End If
    
    Call frm92_pic_visible
    Call Frm92_Initial_Setting
    Call frm92_initial_one_time
    
    Frm92.Frame5.Visible = True
    
Else

    Frm92.Frame5.Visible = False
    
End If
End Sub
Private Sub L6_Text_Click()
'on error resume next
If Frm92.Frame6.Visible = False Then

    Call frm92_pic_visible
    Call Frm92_Initial_Setting
    Call frm92_initial_one_time
    'Call Frm92_Header_Senarai_expense
    
    Frm92.L76_Text = vbNullString 'Tarikh mula
    Frm92.L77_Text = vbNullString 'Tarikh akhir
    Frm92.L78_Text = 0 '0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh
    
    Frm92.L69_Text = -1 'Titik Pencarian Data
    Frm92.L75_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm92.L67_Text = 0 'Paparan Page ke-xxx
    Frm92.L68_Text = 0
    
    GM_NEXT_PREV = 0
    
    Call Frm92_report_expenses_header

    Frm92.Frame6.Visible = True
    
Else

    Frm92.Frame6.Visible = False
    
End If
End Sub

Private Sub LV1_DblClick()
'on error resume next
frm92_LM_No_ID = vbNullString

If IsNumeric(Frm92.LV1.SelectedItem.Index) Then
    
    frm92_LM_No_ID = Frm92.LV1.SelectedItem.Index
    
    If frm92_LM_No_ID <> vbNullString Then
    
    
        PopupMenu Frm92_PM_Menu
    
    Else
    
        MsgBox "Tiada Data.", vbInformation, "Info"
        
    End If
    
Else

    MsgBox "Tiada Data.", vbInformation, "Info"
    
End If
End Sub
Private Sub LV2_DblClick()
'on error resume next
frm92_LM_No_ID = vbNullString

If IsNumeric(Frm92.LV2.SelectedItem.Index) Then
    
    frm92_LM_No_ID = Frm92.LV2.SelectedItem.Index
    
    If frm92_LM_No_ID <> vbNullString Then
    
        user_level = MDI_frm1.L4_Text
        
        If user_level = "Admin" Or user_level = "HQ" Or user_level = "Developer" Then
        
            Frm92.Frm92_SM_edit2.Enabled = True
            Frm92.Frm92_SM_padam2.Enabled = True
            
        ElseIf user_level = "Manager" Then
        
            Frm92.Frm92_SM_edit2.Enabled = True
            Frm92.Frm92_SM_padam2.Enabled = False
        
        Else
        
            Frm92.Frm92_SM_edit2.Enabled = False
            Frm92.Frm92_SM_padam2.Enabled = False
        
        End If
    
        PopupMenu Frm92_PM_Menu2
    
    Else
    
        MsgBox "Tiada Data.", vbInformation, "Info"
        
    End If
    
Else

    MsgBox "Tiada Data.", vbInformation, "Info"
    
End If
End Sub





Private Sub LV3_DblClick()
'on error resume next
frm92_LM_No_ID = vbNullString

If IsNumeric(Frm92.LV3.SelectedItem.Index) Then
    
    frm92_LM_No_ID = Frm92.LV3.SelectedItem.Index
    
    If frm92_LM_No_ID <> vbNullString Then
    
        user = MDI_frm1.L3_Text
        
        user_level = MDI_frm1.L4_Text
        
        If user_level = "Admin" Or user_level = "HQ" Or user_level = "Developer" Then
        
            Frm92.Frm92_SM_padam3.Enabled = True
            Frm92.Frm92_SM_edit3.Enabled = True
            
        ElseIf user_level = "Manager" Then
        
            Frm92.Frm92_SM_padam3.Enabled = False
            Frm92.Frm92_SM_edit3.Enabled = True
        
        Else
        
            Frm92.Frm92_SM_padam3.Enabled = False
            Frm92.Frm92_SM_edit3.Enabled = False
        
        End If
        
        PopupMenu Frm92_PM_Menu3
    
    Else
    
        MsgBox "Tiada Data.", vbInformation, "Info"
        
    End If
    
Else

    MsgBox "Tiada Data.", vbInformation, "Info"
    
End If
End Sub


Private Sub TB2_Change()
'on error resume next
Call frm92_kiraan_gst
End Sub



Private Sub TB46_Change()
'on error resume next
Call frm92_kiraan_cukai_sr_belanja
Call frm92_kiraan_harga_belanja
End Sub

Private Sub TB47_Change()
'on error resume next
Call frm92_kiraan_harga_belanja
End Sub

Private Sub TB48_Change()
'on error resume next
Call frm92_kiraan_cukai_zr_belanja
Call frm92_kiraan_harga_belanja
End Sub

Private Sub TB49_Change()
'on error resume next
Call frm92_kiraan_harga_belanja
End Sub


