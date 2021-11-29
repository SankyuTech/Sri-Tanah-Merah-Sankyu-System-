VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm94 
   Caption         =   "Tempahan Siap"
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
   Icon            =   "Frm94.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13035
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
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
      Left            =   10440
      MouseIcon       =   "Frm94.frx":0ECA
      MousePointer    =   99  'Custom
      Picture         =   "Frm94.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   9840
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
      Height          =   1215
      Left            =   8160
      Picture         =   "Frm94.frx":379E
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   6720
      Width           =   3615
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
      Left            =   8160
      TabIndex        =   73
      Top             =   9030
      Width           =   200
   End
   Begin VB.TextBox TB24 
      BackColor       =   &H8000000A&
      Height          =   360
      Left            =   3150
      Locked          =   -1  'True
      TabIndex        =   68
      Text            =   "TB24"
      Top             =   6480
      Width           =   1395
   End
   Begin VB.TextBox TB17 
      BackColor       =   &H8000000A&
      Height          =   360
      Left            =   3150
      Locked          =   -1  'True
      TabIndex        =   65
      Text            =   "TB17"
      Top             =   6120
      Width           =   1395
   End
   Begin VB.TextBox TB23 
      BackColor       =   &H8000000A&
      Height          =   360
      Left            =   3150
      Locked          =   -1  'True
      TabIndex        =   64
      Text            =   "TB23"
      Top             =   5760
      Width           =   1395
   End
   Begin VB.CommandButton CMD1 
      Caption         =   "Carian Data"
      Height          =   375
      Left            =   5280
      MouseIcon       =   "Frm94.frx":5D68
      MousePointer    =   99  'Custom
      TabIndex        =   63
      Top             =   1320
      Width           =   2505
   End
   Begin VB.CheckBox CB5 
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
      Left            =   4680
      TabIndex        =   61
      Top             =   6240
      Width           =   200
   End
   Begin VB.ComboBox CBB1 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   360
      Left            =   9615
      Style           =   2  'Dropdown List
      TabIndex        =   51
      Top             =   8400
      Width           =   5475
   End
   Begin VB.TextBox TB22 
      BackColor       =   &H8000000A&
      Height          =   360
      Left            =   3150
      Locked          =   -1  'True
      TabIndex        =   49
      Text            =   "TB22"
      Top             =   7200
      Width           =   1395
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
      Left            =   4680
      TabIndex        =   42
      Top             =   6000
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
      Left            =   4680
      TabIndex        =   41
      Top             =   5775
      Width           =   200
   End
   Begin VB.CheckBox CB1 
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
      Left            =   360
      TabIndex        =   39
      Top             =   570
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
      Left            =   240
      TabIndex        =   37
      Top             =   8070
      Width           =   200
   End
   Begin VB.PictureBox Pic1 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   240
      ScaleHeight     =   2175
      ScaleWidth      =   5265
      TabIndex        =   28
      Top             =   8400
      Width           =   5265
      Begin VB.CommandButton CMD3 
         Caption         =   "Batal"
         Height          =   375
         Left            =   1440
         MouseIcon       =   "Frm94.frx":6072
         MousePointer    =   99  'Custom
         TabIndex        =   71
         ToolTipText     =   "Batal bayaran menggunakan trade in"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton CMD2 
         Caption         =   "Carian"
         Height          =   375
         Left            =   3480
         MouseIcon       =   "Frm94.frx":637C
         MousePointer    =   99  'Custom
         TabIndex        =   70
         ToolTipText     =   "Carian Maklumat Terperinci Voucher Buyback / Trade In"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TB14 
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "TB14"
         Top             =   1275
         Width           =   1965
      End
      Begin VB.TextBox TB13 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "TB13"
         Top             =   460
         Width           =   2235
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Carian Voucher Buyback / Trade In"
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
         Left            =   240
         TabIndex        =   36
         Top             =   120
         Width           =   4275
      End
      Begin VB.Label L4_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L4_Text"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   1680
         TabIndex        =   35
         Top             =   1080
         Width           =   1785
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Voucher :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   34
         Top             =   1080
         Width           =   1425
      End
      Begin VB.Label Label51 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Nilaian Voucher:RM"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   33
         Top             =   1320
         Width           =   2715
      End
      Begin VB.Label Label52 
         BackStyle       =   0  'Transparent
         Caption         =   "No.Voucher:"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   0
         TabIndex        =   32
         Top             =   480
         Width           =   2265
      End
      Begin VB.Label Label53 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Maklumat Voucher Buyback / Trade In"
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
         Left            =   240
         TabIndex        =   31
         Top             =   840
         Width           =   4275
      End
   End
   Begin VB.TextBox TB10 
      BackColor       =   &H8000000A&
      Height          =   360
      Left            =   3150
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "TB10"
      Top             =   5400
      Width           =   1395
   End
   Begin VB.TextBox TB11 
      BackColor       =   &H8000000A&
      Height          =   360
      Left            =   3150
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "TB11"
      Top             =   6840
      Width           =   1395
   End
   Begin VB.TextBox TB12 
      BackColor       =   &H8000000A&
      Height          =   360
      Left            =   3150
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "TB12"
      Top             =   7560
      Width           =   1395
   End
   Begin VB.TextBox TB6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2190
      TabIndex        =   8
      Text            =   "TB6"
      Top             =   3705
      Width           =   2000
   End
   Begin VB.TextBox TB2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Height          =   360
      Left            =   2190
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "TB2"
      Top             =   2265
      Width           =   2000
   End
   Begin VB.TextBox TB3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Height          =   360
      Left            =   2190
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "TB3"
      Top             =   2625
      Width           =   2000
   End
   Begin VB.TextBox TB4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2190
      TabIndex        =   5
      Text            =   "TB4"
      Top             =   2985
      Width           =   2000
   End
   Begin VB.TextBox TB5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2190
      TabIndex        =   4
      Text            =   "TB5"
      Top             =   3345
      Width           =   2000
   End
   Begin VB.TextBox TB7 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Height          =   360
      Left            =   2190
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "TB7"
      Top             =   4065
      Width           =   2000
   End
   Begin VB.TextBox TB9 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Height          =   360
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "TB9"
      Top             =   2595
      Width           =   1875
   End
   Begin VB.TextBox TB8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   5880
      TabIndex        =   1
      Text            =   "TB8"
      Top             =   2235
      Width           =   1875
   End
   Begin VB.TextBox TB1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2265
      TabIndex        =   0
      Text            =   "TB1"
      Top             =   1335
      Width           =   2940
   End
   Begin VB.Timer Tmr1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Tmr3 
      Interval        =   100
      Left            =   0
      Top             =   960
   End
   Begin VB.Timer Tmr2 
      Interval        =   100
      Left            =   0
      Top             =   480
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   360
      Left            =   9615
      TabIndex        =   52
      Top             =   8040
      Width           =   5475
      _ExtentX        =   9657
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
      Format          =   416612352
      CurrentDate     =   41561
   End
   Begin VB.Label Label24 
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
      Left            =   8160
      TabIndex        =   75
      Top             =   8760
      Width           =   2295
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   $"Frm94.frx":6686
      ForeColor       =   &H000000FF&
      Height          =   780
      Left            =   8445
      TabIndex        =   74
      Top             =   9000
      Width           =   6330
   End
   Begin VB.Label L32_Text 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "L32_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   21000
      TabIndex        =   72
      Top             =   7680
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Dengan GST (RM)           :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   69
      Top             =   6510
      Width           =   3585
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Cukai GST (RM)            :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   67
      Top             =   6150
      Width           =   3585
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Tanpa GST (RM)             :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   66
      Top             =   5790
      Width           =   3585
   End
   Begin VB.Label L33_Text 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "L33_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   21000
      TabIndex        =   62
      Top             =   7200
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label L15_Text 
      BackColor       =   &H8000000A&
      Caption         =   "L15_Text"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   19320
      TabIndex        =   60
      Top             =   7680
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label L14_Text 
      BackColor       =   &H8000000A&
      Caption         =   "L14_Text"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   19320
      TabIndex        =   59
      Top             =   7320
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label L13_Text 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Ini adalah jumlah yang pihak kedai pulangkan kepada pelanggan kerana lebihan pada bayaran oleh pembeli."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1275
      Left            =   4560
      TabIndex        =   58
      Top             =   6960
      Visible         =   0   'False
      Width           =   3225
   End
   Begin VB.Label L7_Text 
      BackColor       =   &H8000000A&
      Caption         =   "L7_Text"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   19320
      TabIndex        =   57
      Top             =   6240
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label L10_Text 
      BackColor       =   &H8000000A&
      Caption         =   "L10_Text"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   19320
      TabIndex        =   56
      Top             =   6960
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label L9_Text 
      BackColor       =   &H8000000A&
      Caption         =   "L9_Text"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   19320
      TabIndex        =   55
      Top             =   6600
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label80 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Pekerja *"
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   8160
      TabIndex        =   54
      Top             =   8400
      Width           =   2295
   End
   Begin VB.Label Label88 
      BackStyle       =   0  'Transparent
      Caption         =   "Tarikh  *"
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   8160
      TabIndex        =   53
      Top             =   8040
      Width           =   2385
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Bayaran Dari Trade In (RM)       :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   50
      Top             =   7230
      Width           =   3585
   End
   Begin VB.Shape Shape2 
      Height          =   6225
      Left            =   120
      Top             =   4680
      Width           =   7845
   End
   Begin VB.Shape Shape1 
      Height          =   2625
      Left            =   120
      Top             =   1920
      Width           =   7845
   End
   Begin VB.Label L12_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kembali Ke Menu Sebelum"
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
      MouseIcon       =   "Frm94.frx":6772
      MousePointer    =   99  'Custom
      TabIndex        =   48
      ToolTipText     =   "Keluar Ke Menu Sebelum"
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label L5_Text 
      BackColor       =   &H8000000A&
      Caption         =   "L5_Text"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   19320
      TabIndex        =   47
      Top             =   5880
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Shape Shape13 
      Height          =   1515
      Left            =   4560
      Top             =   5400
      Width           =   2835
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Zero Rated ZR(L)        Standard Rated SR            Standard Rated Inclusive SR"
      ForeColor       =   &H00000000&
      Height          =   765
      Left            =   4920
      TabIndex        =   44
      Top             =   5760
      Width           =   2625
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Maklumat GST"
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
      Left            =   4680
      TabIndex        =   43
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Label Label50 
      BackStyle       =   0  'Transparent
      Caption         =   "Scanner Mode"
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   675
      TabIndex        =   40
      Top             =   525
      Width           =   2385
   End
   Begin VB.Label Label54 
      BackStyle       =   0  'Transparent
      Caption         =   "Bayaran Dari Barang Trade In"
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   525
      TabIndex        =   38
      Top             =   8040
      Width           =   3690
   End
   Begin VB.Label Label102 
      BackStyle       =   0  'Transparent
      Caption         =   "Maklumat Baki Bayaran."
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
      TabIndex        =   27
      Top             =   4800
      Width           =   4455
   End
   Begin VB.Label Label103 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Jualan (RM)                    :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   26
      Top             =   5430
      Width           =   3585
   End
   Begin VB.Label Label104 
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit (RM)                            :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   25
      Top             =   6870
      Width           =   3585
   End
   Begin VB.Label L6_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "Lebihan Kedai Perlu Bayar (RM)  :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   24
      Top             =   7590
      Width           =   3585
   End
   Begin VB.Label Label90 
      BackStyle       =   0  'Transparent
      Caption         =   "Upah                   RM"
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   240
      TabIndex        =   20
      Top             =   3750
      Width           =   2265
   End
   Begin VB.Label Label91 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Siri Produk     "
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   240
      TabIndex        =   19
      Top             =   2280
      Width           =   2265
   End
   Begin VB.Label Label92 
      BackStyle       =   0  'Transparent
      Caption         =   "Berat Asal               g"
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   240
      TabIndex        =   18
      Top             =   2655
      Width           =   2265
   End
   Begin VB.Label Label93 
      BackStyle       =   0  'Transparent
      Caption         =   "Berat Jualan            g"
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   240
      TabIndex        =   17
      Top             =   3015
      Width           =   2265
   End
   Begin VB.Label Label94 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Semasa   RM/g"
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   240
      TabIndex        =   16
      Top             =   3390
      Width           =   2265
   End
   Begin VB.Label Label95 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Asal           RM"
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   240
      TabIndex        =   15
      Top             =   4080
      Width           =   2265
   End
   Begin VB.Label Label96 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Jualan  RM"
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   4320
      TabIndex        =   14
      Top             =   2655
      Width           =   2265
   End
   Begin VB.Label Label97 
      BackStyle       =   0  'Transparent
      Caption         =   "Adjustment    RM"
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   4320
      TabIndex        =   13
      Top             =   2310
      Width           =   2265
   End
   Begin VB.Label L3_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L3_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2265
      TabIndex        =   12
      Top             =   1995
      Width           =   5835
   End
   Begin VB.Label Label99 
      BackStyle       =   0  'Transparent
      Caption         =   "Kategori Produk       :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Top             =   1995
      Width           =   2385
   End
   Begin VB.Label Label100 
      BackStyle       =   0  'Transparent
      Caption         =   "Sila scan barang tempahan yang telah siap."
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   480
      TabIndex        =   10
      Top             =   1035
      Width           =   5745
   End
   Begin VB.Label Label101 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Siri Produk      :"
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   480
      TabIndex        =   9
      Top             =   1365
      Width           =   1905
   End
   Begin VB.Shape Shape12 
      Height          =   900
      Left            =   120
      Top             =   960
      Width           =   7845
   End
   Begin VB.Label L8_Text 
      Alignment       =   1  'Right Justify
      Caption         =   "L8_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5950
      TabIndex        =   46
      Top             =   6600
      Width           =   720
   End
   Begin VB.Label Label58 
      BackStyle       =   0  'Transparent
      Caption         =   "Kadar cukai GST          %"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4680
      TabIndex        =   45
      Top             =   6600
      Width           =   2760
   End
End
Attribute VB_Name = "Frm94"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB1_Click()
'on error resume next
If Frm94.CB1 = 1 Then
    'Frm94.TB1.SetFocus
End If
End Sub
Private Sub CB2_Click()
'on error resume next
If Frm94.CB2 = 1 Then
    Frm94.TB13.Locked = False
    Frm94.TB13.BackColor = &HFFFFFF
    
    If GLOBAL_DISABLE = 0 Then
        Frm94.TB13.SetFocus
    End If
Else
    Frm94.TB13.Locked = True
    Frm94.TB13.BackColor = &H8000000A
    
    Frm94.TB13 = vbNullString
    Frm94.L4_Text = vbNullString
    Frm94.TB14 = "0.00"
    Frm94.TB22 = "0.00"
End If
End Sub
Private Sub CB3_Click()
'On Error Resume Next
If Frm94.CB3 = 1 Then
    Frm94.CB4 = 0
    Frm94.CB5 = 0
End If

Call frm94_kiraan_cukai_gst
End Sub
Private Sub CB4_Click()
'On Error Resume Next
If Frm94.CB4 = 1 Then
    Frm94.CB3 = 0
    Frm94.CB5 = 0
End If

Call frm94_kiraan_cukai_gst
End Sub
Private Sub CB5_Click()
'On Error Resume Next
If Frm94.CB5 = 1 Then
    Frm94.CB4 = 0
    Frm94.CB3 = 0
End If

Call frm94_kiraan_cukai_gst
End Sub



Private Sub CBB2_Click()
'on error resume next
If frm130.CBB2 <> vbNullString Then
    If GLOBAL_DISABLE = 0 Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 74_cas_kad_kredit where jenis_kad='" & frm130.CBB2 & "' AND status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!cas_kad) Then
                If IsNumeric(rs!cas_kad) Then
                    frm130.L31_Text = rs!cas_kad
                Else
                    frm130.L31_Text = "0.00"
                End If
            Else
                frm130.L31_Text = "0.00"
            End If
            
        End If
        
        rs.Close
        Set rs = Nothing
    
    End If
End If
End Sub

Private Sub CMD1_Click()
'on error resume next
If Frm94.TB1 = vbNullString Then
    MsgBox "Sila Masukkan No. Siri Produk.", vbInformation, "Info"
    Exit Sub
End If

If Frm94.TB1 <> vbNullString Then
    If InStr(1, Frm94.TB1, "*") <> 0 Or InStr(1, Frm94.TB1, "/") <> 0 Or InStr(1, Frm94.TB1, "\") <> 0 Or InStr(1, Frm94.TB1, "'") <> 0 Then
        MsgBox "No. Siri Produk mengandungi simbol yang tidak sah.", vbExclamation, "Info"
        
        Frm94.TB1.SetFocus
        Exit Sub
    End If
End If


Call Frm94_Call_Product_Detail
End Sub
Private Sub CMD2_Click()
'on error resume next
DATA_FOUND = 0

If Frm94.TB13 = vbNullString Then
    MsgBox "Sila Masukkan No. Voucher Buyback/Trade In.", vbInformation, "Info"
    Exit Sub
End If
If Frm93.TB13 <> vbNullString Then
    If InStr(1, Frm93.TB13, "*") <> 0 Or InStr(1, Frm93.TB13, "/") <> 0 Or InStr(1, Frm93.TB13, "\") <> 0 Or InStr(1, Frm93.TB13, "'") <> 0 Then
        MsgBox "No. Voucher trade in mengandungi simbol yang tidak sah.", vbExclamation, "Info"
        
        Frm93.TB13.SetFocus
        Exit Sub
    End If
End If

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & UCase(Frm94.TB13) & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!trade_in_status) Then
        If rs!trade_in_status = 0 Then
            If Not IsNull(rs!jumlah_tanpa_gst) Then
                Frm94.TB14 = Format(rs!jumlah_tanpa_gst, "#,##0.00") 'Jumlah Nilaian Resit Trade In
                Frm94.TB22 = Format(rs!jumlah_tanpa_gst, "#,##0.00") 'Jumlah Nilaian Resit Trade In
            End If
            
            Frm94.L4_Text = UCase(Frm94.TB13) 'No. Voucher Trade In
            Frm94.TB13 = vbNullString
            
            DATA_FOUND = 1
        ElseIf rs!trade_in_status = 1 Then
            MsgBox "No. Voucher Trade In ini telah digunakan untuk urusan belian sebelum ini.", vbInformation, "Info"
            
            Frm94.TB13 = vbNullString
            Frm94.TB13.SetFocus
        End If
    End If
Else
    MsgBox "No. Voucher tidak dijumpai.", vbInformation, "Info"
    
    Frm94.TB13 = vbNullString
    Frm94.TB13.SetFocus
End If

rs.Close
Set rs = Nothing
End Sub
Private Sub CMD3_Click()
'on error resume next
If Frm94.L4_Text <> vbNullString Then
    Note = "Adakah anda ingin batalkan No. Voucher ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
        Frm94.L4_Text = vbNullString
        Frm94.TB14 = "0.00"
        Frm94.TB22 = "0.00"
    End If
Else
    MsgBox "Tiada maklumat tentang voucher trade in.", vbInformation, "Info"
End If
End Sub
Private Sub CMD4_Click()
'On Error Resume Next
Dim Err(30)
Dim Frm94_LM_ERR_BERAT_ASAL As Double
Dim Frm94_LM_ERR_BERAT_JUALAN As Double
Dim Frm94_LM_ERR_JUMLAH_BAYARAN As Double
Dim Frm94_LM_ERR_HARGA As Double
Dim Frm94_LM_JUMLAH_SIMPANAN As Double
Dim Frm94_LM_GUNA_SIMPAN As Double

Frm94_LM_ERR_BERAT_ASAL = 0
Frm94_LM_ERR_BERAT_JUALAN = 0
Frm94_LM_ERR_JUMLAH_BAYARAN = 0 'Jumlah Bayaran
Frm94_LM_ERR_HARGA = 0 'Jumlah Perlu Bayar
Frm94_LM_Flag_TRADE_IN = 0 '0 : Tiada Urusan Trade In , 1 : Ada Urusan Trade In
Frm94_LM_Flag_SIMPANAN = 0 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
Frm94_LM_JUMLAH_SIMPANAN = 0  'Jumlah Simpanan Yang Ada
Frm94_LM_GUNA_SIMPAN = 0 'Jumlah Simpanan Yang Hendak Digunakan

Frm94_LM_NAMA = vbNullString
Frm94_LM_IC = vbNullString
Frm94_LM_No_TEL = vbNullString
Frm94_LM_No_RUJUKAN_CUST = vbNullString
Frm94_KOD_PURITY = vbNullString
Frm94_DULANG = vbNullString

If Frm94.TB2 = vbNullString Then
    x = x + 1
    Err(x) = "Tiada maklumat no. siri produk"
End If
If Frm94.L3_Text = vbNullString Then
    x = x + 1
    Err(x) = "Tiada maklumat kategori produk"
End If
If Frm94.L15_Text = 0 Then 'Jenis Barang : Barang Kemas
    If Frm94.TB4 = vbNullString Or (Frm94.TB4 <> vbNullString And Not IsNumeric(Frm94.TB4)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Berat Jualan (g)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm94.TB5 = vbNullString Or (Frm94.TB5 <> vbNullString And Not IsNumeric(Frm94.TB5)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Harga Semasa (RM/g)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm94.TB6 = vbNullString Or (Frm94.TB6 <> vbNullString And Not IsNumeric(Frm94.TB6)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Upah (RM)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If

    If (Frm94.TB3 <> vbNullString And IsNumeric(Frm94.TB3)) And (Frm94.TB4 <> vbNullString And IsNumeric(Frm94.TB4)) Then
        Frm94_LM_ERR_BERAT_ASAL = Frm94.TB3 'Berat Asal
        Frm94_LM_ERR_BERAT_JUALAN = Frm94.TB4 'Berat Jualan
        
        If Frm94_LM_ERR_BERAT_JUALAN > Frm94_LM_ERR_BERAT_ASAL Then
            x = x + 1
            Err(x) = "Berat jualan melebihi berat asal"
        End If
    End If
End If
If Frm94.L15_Text = 1 Then 'Jenis Barang : Barang Permata
    If Frm94.TB6 = vbNullString Or (Frm94.TB6 <> vbNullString And Not IsNumeric(Frm94.TB6)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Upah (RM)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm94.TB7 = vbNullString Or (Frm94.TB7 <> vbNullString And Not IsNumeric(Frm94.TB7)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Harga Asal (RM)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If
If Frm94.TB8 = vbNullString Or (Frm94.TB8 <> vbNullString And Not IsNumeric(Frm94.TB8)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Adjustment (RM)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm94.TB10 = vbNullString Or (Frm94.TB10 <> vbNullString And Not IsNumeric(Frm94.TB10)) Then
    x = x + 1
    Err(x) = "Tiada maklumat bagi [Jumlah Harga Jualan]."
End If
If Frm94.TB11 = vbNullString Or (Frm94.TB11 <> vbNullString And Not IsNumeric(Frm94.TB11)) Then
    x = x + 1
    Err(x) = "Tiada maklumat bagi [Bayaran Sudah Jelas]."
End If
If Frm94.TB12 = vbNullString Or (Frm94.TB12 <> vbNullString And Not IsNumeric(Frm94.TB12)) Then
    x = x + 1
    Err(x) = "Tiada maklumat bagi [Baki]."
End If

If Frm94.TB22 = vbNullString Or (Frm94.TB22 <> vbNullString And Not IsNumeric(Frm94.TB22)) Then
    x = x + 1
    Err(x) = "Tiada maklumat bagi [Bayaran Dari Trade In]."
End If

If Frm94.CB3 = 0 And Frm94.CB4 = 0 And Frm94.CB5 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan jenis cukai GST."
End If
If Frm94.CB2 = 1 Then 'Bayaran Deposit Dari Barang Trade In
    If Frm94.L4_Text = vbNullString Then
        x = x + 1
        Err(x) = "Tiada maklumat no. voucher trade in."
    End If
    If Frm94.TB14 = vbNullString Then
        x = x + 1
        Err(x) = "Tiada maklumat nilaian voucher trade in."
    End If
End If
If Frm94.L13_Text.Visible = False Then
    If frm130.TB27 = vbNullString Or (frm130.TB27 <> vbNullString And Not IsNumeric(frm130.TB27)) Then
        x = x + 1
        Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara tunai. Sila masukkan 0 jika tiada bayaran tunai."
    End If
    If frm130.TB28 = vbNullString Or (frm130.TB28 <> vbNullString And Not IsNumeric(frm130.TB28)) Then
        x = x + 1
        Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara bank in. Sila masukkan 0 jika tiada bayaran bank in."
    End If
    If frm130.TB29 = vbNullString Or (frm130.TB29 <> vbNullString And Not IsNumeric(frm130.TB29)) Then
        x = x + 1
        Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara kad kredit. Sila masukkan 0 jika tiada bayaran kad kredit."
    End If
    If frm130.TB21 = vbNullString Or (frm130.TB21 <> vbNullString And Not IsNumeric(frm130.TB21)) Then
        x = x + 1
        Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara duit simpanan di kedai. Sila masukkan 0 jika tiada bayaran simpanan di kedai."
    End If
    If (frm130.TB21 <> vbNullString And IsNumeric(frm130.TB21)) And (frm130.L26_Text <> vbNullString And IsNumeric(frm130.L26_Text)) Then
        Frm94_LM_JUMLAH_SIMPANAN = frm130.L26_Text  'Jumlah Simpanan Yang Ada
        Frm94_LM_GUNA_SIMPAN = frm130.TB21  'Jumlah Simpanan Yang Hendak Digunakan

        If Frm94_LM_GUNA_SIMPAN > Frm94_LM_JUMLAH_SIMPANAN Then
            x = x + 1
            Err(x) = "Jumlah simpanan yang hendak digunakan melebihi simpanan yang ada."
        End If
    End If
    If Frm94.L6_Text = "Baki (RM)                                 :" Then
        If (frm130.TB32 <> vbNullString And IsNumeric(frm130.TB32)) And (Frm94.TB12 <> vbNullString And IsNumeric(Frm94.TB12)) Then
            Frm94_LM_ERR_JUMLAH_BAYARAN = frm130.TB32 'Jumlah Bayaran
            Frm94_LM_ERR_HARGA = Frm94.TB12 'Jumlah Perlu Bayar
            
            If Frm94_LM_ERR_JUMLAH_BAYARAN <> Frm94_LM_ERR_HARGA Then
                x = x + 1
                Err(x) = "Jumlah bayaran tidak sama dengan jumlah perlu bayar."
            End If
        End If
    End If
End If
If Frm94.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih nama pekerja."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah anda ingin simpan data ini?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
        '$$$ No. staff $$$ - Start
        If InStr(1, Frm94.CBB1, "  |  ") <> 0 Then
        
            Frm94_LM_EMP_NO = Split(Frm94.CBB1, "  |  ")(1)
            
        Else
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!NoPekerja) Then Frm94_LM_EMP_NO = rs!NoPekerja
    
            End If
            
            rs.Close
            Set rs = Nothing
        
        End If
        
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
        '$$$ No. staff $$$ - End
        
skip_carian_user:
    
        '### Pop up confirmation bagi jualan bagi invoice tidak rasmi
        If Frm94.CB9 = 1 Then
        
            Note = "Jualan ini dibuat dengan pilihan INVOICE TIDAK RASMI." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Anda TIDAK BOLEH mengubah jenis invoice jika jualan ini telah dibuat." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Teruskan?"
                
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbNo Then
            
                Exit Sub
            
            End If
            
        End If
        
        G_JENIS_URUSAN = 8
        
        If Frm94.L9_Text <> vbNullString Then
            Frm94_LM_No_RUJUKAN_TEMPAHAN = Frm94.L9_Text 'No. Rujukan Tempahan
        Else
            Frm94_LM_No_RUJUKAN_TEMPAHAN = 1
        End If
        'If Frm94.CB9 = 0 Then
        '    If Frm94.L10_Text <> vbNullString Then
        '        Frm94_LM_No_RESIT_TEMPAHAN = Frm94.L10_Text 'No. invoice rasmi
        '    Else
        '        Frm94_LM_No_RESIT_TEMPAHAN = 1
        '    End If
        'Else
        '    If Frm94.L32_Text <> vbNullString Then
        '        Frm94_LM_No_RESIT_TEMPAHAN = Frm94.L32_Text 'No. invoice tidak rasmi
        '    Else
        '        Frm94_LM_No_RESIT_TEMPAHAN = 1
        '    End If
        'End If
        
        '###Carian Purity Item Ini ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where no_siri_Produk='" & Frm94.TB2 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If rs!StatusItem = "10" Or rs!StatusItem = "14" Then

                If Not IsNull(rs!kod_Purity) Then Frm94_KOD_PURITY = rs!kod_Purity 'Kod Purity
                If Not IsNull(rs!dulang) Then Frm94_DULANG = rs!dulang 'Dulang
                
            ElseIf rs!StatusItem = "11" Then
            
                MsgBox "Item ini telah terjual.", vbExclamation, "Info"
                
                Exit Sub
                
            ElseIf rs!StatusItem = "12" Then
            
                MsgBox "Item ini telah terjual secara potong.", vbExclamation, "Info"
                
                Exit Sub
                
            ElseIf rs!StatusItem = "13" Then
            
                MsgBox "Item ini telah terjual secara potong.", vbExclamation, "Info"
                
                Exit Sub
                
            ElseIf rs!StatusItem = "21" Or rs!StatusItem = "22" Then
                MsgBox "Item ini telah terjual secara tempahan.", vbExclamation, "Info"
                
                Exit Sub
                
            ElseIf rs!StatusItem = "15" Or rs!StatusItem = "19" Or rs!StatusItem = "20" Then
                MsgBox "Item ini telah terjual secara ansuran.", vbExclamation, "Info"
                
                Exit Sub
                
            ElseIf rs!StatusItem = "16" Then
            
                MsgBox "Item ini telah dihantar ke ar-rahnu.", vbExclamation, "Info"
                
                Exit Sub
                
            ElseIf rs!StatusItem = "17" Then
            
                MsgBox "Item ini telah terjual secara ETA.", vbExclamation, "Info"
                
                Exit Sub
                
            ElseIf rs!StatusItem = "23" Then
            
                MsgBox "Item ini telah dihantar kepada cawangan / agen / kilang.", vbExclamation, "Info"
                
                Exit Sub
                
            ElseIf rs!StatusItem = "24" Then
            
                MsgBox "Item ini telah dihantar kepada cawangan / agen / kilang.", vbExclamation, "Info"
                
                Exit Sub
                
            ElseIf rs!StatusItem = "25" Then
            
                MsgBox "Item ini telah diagihkan kepada cawangan / agen / kilang.", vbExclamation, "Info"
        
                Exit Sub
                
            ElseIf rs!StatusItem = "26" Then
            
                MsgBox "Item ini telah dijual cawangan atau agen.", vbExclamation, "Info"
        
                Exit Sub
                
            ElseIf rs!StatusItem = "0" Then
            
                MsgBox "Item ini telah dipadamkan dari database.", vbExclamation, "Info"
                
                Exit Sub
                
            ElseIf rs!StatusItem = "27" Or rs!StatusItem = "28" Then
                
                MsgBox "Item Ini Telah Dijual Dari Menu GDN.", vbExclamation, "Info"
        
                Exit Sub
                
            ElseIf rs!StatusItem = "29" Then
            
                MsgBox "Item Ini Telah Diubah Status Kepada Hilang , Dicuri Dan Sebagainya.", vbExclamation, "Info"
        
                Exit Sub
                
            End If
            
        Else
        
            MsgBox "Tiada data berkenaan item ini dijumpai. Sila periksa data stok dan status terkini stok ini.", vbExclamation, "Info"
            
            Exit Sub
            
        End If
        
        rs.Close
        Set rs = Nothing
        '###Carian Purity Item Ini ### - End
        
'---------------------------------------No. Invoice
        LM_NOW = Now
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        If Frm94.CB9 = 0 Then rs.Open "select * from 4_senarai_invoice_rasmi", cn2, adOpenKeyset, adLockOptimistic
        If Frm94.CB9 = 1 Then rs.Open "select * from 5_senarai_invoice_tidak_rasmi", cn2, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        rs!tarikh = Frm94.DTPicker1
        rs!terminal = G_TERMINAL
        rs!write_timestamp = LM_NOW
        rs!Status = 1
        rs!nama_staff = MDI_frm1.L3_Text
        rs.Update
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        If Frm94.CB9 = 0 Then rs.Open "select * from 4_senarai_invoice_rasmi where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & Frm94.DTPicker1 & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic
        If Frm94.CB9 = 1 Then rs.Open "select * from 5_senarai_invoice_tidak_rasmi where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & Frm94.DTPicker1 & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!ID) Then
                
                Frm94_LM_No_RESIT_TEMPAHAN = rs!ID 'No. Rujukan Belian
                If Frm94.CB9 = 0 Then rs!no_invoice = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(rs!ID, "000000")
                If Frm94.CB9 = 1 Then rs!no_invoice = "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(rs!ID, "000000")
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

'### Periksa NO INVOICE sebelum simpan data ke dalam database ### - Start
        GoTo a:
        
Re_gen_no_resit:
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm94.CB9 = 0 Then rs.Open "select * from 22_jualan where no_resit='" & "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm94_LM_No_RESIT_TEMPAHAN, "000000") & "' AND bil_rasmi = 1 AND status = 1", cn, adOpenKeyset, adLockOptimistic
        If Frm94.CB9 = 1 Then rs.Open "select * from 22_jualan where no_resit='" & "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm94_LM_No_RESIT_TEMPAHAN, "000000") & "' AND bil_rasmi = 0 AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            Frm94_LM_No_RESIT_TEMPAHAN = Frm94_LM_No_RESIT_TEMPAHAN + 1
            If Frm94.CB9 = 0 Then Frm94.L10_Text = Frm94_LM_No_RESIT_TEMPAHAN
            If Frm94.CB9 = 1 Then Frm94.L32_Text = Frm94_LM_No_RESIT_TEMPAHAN
            
            rs.Close
            Set rs = Nothing
            
            GoTo Re_gen_no_resit:
        End If
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm94.CB9 = 0 Then rs.Open "select * from 42_tempahan_siap where no_resit_tempahan='" & "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm94_LM_No_RESIT_TEMPAHAN, "000000") & "' AND bil_rasmi = 1 AND status_invoice = 1", cn, adOpenKeyset, adLockOptimistic
        If Frm94.CB9 = 1 Then rs.Open "select * from 42_tempahan_siap where no_resit_tempahan='" & "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm94_LM_No_RESIT_TEMPAHAN, "000000") & "' AND bil_rasmi = 0 AND status_invoice = 1", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            Frm94_LM_No_RESIT_TEMPAHAN = Frm94_LM_No_RESIT_TEMPAHAN + 1
            If Frm94.CB9 = 0 Then Frm94.L10_Text = Frm94_LM_No_RESIT_TEMPAHAN
            If Frm94.CB9 = 1 Then Frm94.L32_Text = Frm94_LM_No_RESIT_TEMPAHAN
            
            rs.Close
            Set rs = Nothing
            
            GoTo Re_gen_no_resit:
        End If
        
        rs.Close
        Set rs = Nothing
        
a:
'### Periksa NO INVOICE sebelum simpan data ke dalam database ### - End

'### Carian Maklumat Pembeli Ini ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 40_tempahan_deposit where no_rujukan_tempahan='" & Frm94_LM_No_RUJUKAN_TEMPAHAN & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!Nama) Then Frm94_LM_NAMA = rs!Nama
            If Not IsNull(rs!no_ic) Then Frm94_LM_IC = rs!no_ic
            If Not IsNull(rs!no_tel) Then Frm94_LM_No_TEL = rs!no_tel
            If Not IsNull(rs!no_rujukan_pelanggan) Then Frm94_LM_No_RUJUKAN_CUST = rs!no_rujukan_pelanggan
            If Not IsNull(rs!kategori_pembeli) Then Frm94_LM_KATEGORI = rs!kategori_pembeli
            
            rs!terminal = G_TERMINAL
            LM_NOW = Now
            rs!write_timestamp2 = LM_NOW
            rs!no_staff = G_LOGIN_USER 'No. staff yang memasukkan/edit data (Login)
    
            rs!Status = "Siap" 'Status

            If Frm94.CB9 = 0 Then
            
                rs!invoice_siap = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm94_LM_No_RESIT_TEMPAHAN, "000000") 'No. invoice rasmi
                    
            Else
            
                rs!invoice_siap = "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm94_LM_No_RESIT_TEMPAHAN, "000000") 'No. invoice tidak rasmi
            
            End If
            
            rs.Update
            
        End If
        
        rs.Close
        Set rs = Nothing
'### Carian Maklumat Pembeli Ini ### - End

LM_RE_GEN_INVOICE_NO = 0

Re_Gen_No_Rujukan:
'###Masukkan Jualan Ke Dalam Table Akaun Jualan### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm94.CB9 = 0 Then rs.Open "select * from 22_jualan where no_resit='" & "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm94_LM_No_RESIT_TEMPAHAN, "000000") & "' AND bil_rasmi = 1 AND status = 1", cn, adOpenKeyset, adLockOptimistic
        If Frm94.CB9 = 1 Then rs.Open "select * from 22_jualan where no_resit='" & "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm94_LM_No_RESIT_TEMPAHAN, "000000") & "' AND bil_rasmi = 0 AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
        
            rs.AddNew
            If Frm94.CB9 = 0 Then
            
                If Frm94.L10_Text <> vbNullString Then
                    rs!no_resit = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm94_LM_No_RESIT_TEMPAHAN, "000000") 'No. invoice rasmi
                    LM_NO_INVOICE = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm94_LM_No_RESIT_TEMPAHAN, "000000") 'No. invoice rasmi
                Else
                    rs!no_resit = Null 'No. invoice rasmi
                End If
                rs!bil_rasmi = 1 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
                
            Else
            
                If Frm94.L32_Text <> vbNullString Then
                    rs!no_resit = "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm94_LM_No_RESIT_TEMPAHAN, "000000") 'No. invoice tidak rasmi
                    LM_NO_INVOICE = "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm94_LM_No_RESIT_TEMPAHAN, "000000") 'No. invoice tidak rasmi
                Else
                    rs!no_resit = Null 'No. invoice tidak rasmi
                End If
                rs!bil_rasmi = 0 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
            
            End If
            rs!tarikh = Frm94.DTPicker1 'Tarikh Jualan
            
            rs!tunai = Format(0, "0.00")
            rs!bank_in = Format(0, "0.00") 'Cara Bayaran : Bank In
            rs!kad_kredit = Format(0, "0.00") 'Cara Bayaran : Kad Kredit
            rs!duit_simpanan_kedai = Format(0, "0.00") 'Cara Bayaran : Simpanan Duit Di Kedai
            rs!cek = Format(0, "0.00") 'Cara Bayaran : Cek
            
            If Frm94.L13_Text.Visible = False Then
            
                If frm130.TB27 <> vbNullString Then
                    rs!tunai = Format(frm130.TB27, "0.00") 'Cara Bayaran : Tunai
                Else
                    rs!tunai = Null 'Cara Bayaran : Tunai
                End If
                If frm130.TB28 <> vbNullString Then
                    rs!bank_in = Format(frm130.TB28, "0.00") 'Cara Bayaran : Bank In
                Else
                    rs!bank_in = Null 'Cara Bayaran : Bank In
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
                        If Frm94.L8_Text <> vbNullString Then
                            rs!kadar_gst_kad_kredit = Format(Frm94.L8_Text, "0.00") 'Cara Bayaran : Kadar GST bagi kad kredit
                        Else
                            rs!kadar_gst_kad_kredit = "0.00" 'Cara Bayaran : Kadar GST bagi kad kredit
                        End If
                        rs!epp = 0 '0 : Bayaran selain dari EPP , 1 : Bayaran secara EPP
                        rs!approval_code_epp = Null 'Approval Code (EPP)
                        
                    Else
                
                        rs!jenis_kad = Null
                        rs!cas_Kad_Kredit = Null 'Cara Bayaran : Cas Kad Kredit (%)
                        rs!jumlah_cas_kad_kredit = Null 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
                        rs!jumlah_potongan_kad_kredit = Null 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
                        rs!kadar_gst_kad_kredit = Null 'Cara Bayaran : Kadar GST bagi kad kredit
                        rs!gst_kad_kredit = Null 'Cara Bayaran : Jumlah GST kad kredit (RM)
                        
                        rs!epp = 0 '0 : Bayaran selain dari EPP , 1 : Bayaran secara EPP
                        rs!approval_code_epp = Null 'Approval Code (EPP)
                                        
                    End If
                
                Else
                    
                    'rs!kad_kredit = Null
                    rs!jenis_kad = Null
                    rs!cas_Kad_Kredit = Null 'Cara Bayaran : Cas Kad Kredit (%)
                    rs!jumlah_cas_kad_kredit = Null 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
                    rs!jumlah_potongan_kad_kredit = Null 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
                    rs!kadar_gst_kad_kredit = Null 'Cara Bayaran : Kadar GST bagi kad kredit
                    rs!gst_kad_kredit = Null 'Cara Bayaran : Jumlah GST kad kredit (RM)
                    
                    rs!epp = 0 '0 : Bayaran selain dari EPP , 1 : Bayaran secara EPP
                    rs!approval_code_epp = Null 'Approval Code (EPP)
                                        
                End If
                If frm130.TB21 <> vbNullString Then
                    If Format(frm130.TB21, "0.00") <> "0.00" Then
                        Frm94_LM_Flag_SIMPANAN = 1 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
                    End If
                    rs!duit_simpanan_kedai = Format(frm130.TB21, "0.00") 'Cara Bayaran : Simpanan Duit Di Kedai
                Else
                    rs!duit_simpanan_kedai = Null 'Cara Bayaran : Simpanan Duit Di Kedai
                End If
                If frm130.TB32 <> vbNullString Then
                    rs!jumlah_bayaran = Format(frm130.TB32, "0.00") 'Cara Bayaran : Jumlah Bayaran
                Else
                    rs!jumlah_bayaran = Null 'Cara Bayaran : Jumlah Bayaran
                End If

            End If

            If Frm94.TB23 <> vbNullString Then
                rs!harga_barang = Format(Frm94.TB23, "0.00") 'Jumlah Harga Barang Tanpa GST (RM)
            Else
                rs!harga_barang = Null 'Jumlah Harga Barang Tanpa GST (RM)
            End If
            If Frm94.TB17 <> vbNullString Then
                rs!jumlah_cukai_gst = Format(Frm94.TB17, "0.00") 'Jumlah Cukai GST (ZR + SR)
            Else
                rs!jumlah_cukai_gst = Null 'Jumlah Cukai GST (ZR + SR)
            End If
            If Frm94.TB24 <> vbNullString Then
                rs!harga_barang_dengan_gst = Format(Frm94.TB24, "0.00") 'Jumlah Harga Barang Dengan GST (RM)
            Else
                rs!harga_barang_dengan_gst = Null 'Jumlah Harga Barang Dengan GST (RM)
            End If
            rs!diskaun = Format(0, "0.00") 'Jumlah Diskaun (%)
            If Frm94.TB24 <> vbNullString Then
                rs!harga_lepas_diskaun = Format(Frm94.TB24, "0.00") 'Harga Selepas Diskaun (RM)
            Else
                rs!harga_lepas_diskaun = Null 'Harga Selepas Diskaun (RM)
            End If
            rs!adjustment = Format(0, "0.00") 'Adjustment (RM)
            If Frm94.TB24 <> vbNullString Then
                rs!harga_jualan = Format(Frm94.TB24, "0.00") 'Jumlah Harga Jualan (RM)
            Else
                rs!harga_jualan = Null 'Jumlah Harga Jualan (RM)
            End If
            rs!loss_trade_in = Format(0, "0.00") 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (%)
            rs!loss_trade_in_rm = Format(0, "0.00") 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (RM)
            
            If Frm94.L6_Text = "Baki (RM)                                 :" Then
                rs!flag_bayaran = 0 '0 : Pembeli Bayar , 1 : Kedai Bayar
            ElseIf Frm94.L6_Text = "Lebihan Kedai Perlu Bayar (RM)  :" Then
                rs!flag_bayaran = 1 '0 : Pembeli Bayar , 1 : Kedai Bayar
            End If
            If Frm94.TB12 <> vbNullString Then
                rs!jumlah_perlu_bayar = Format(Frm94.TB12, "0.00") 'Jumlah Bayaran Yang Perlu Dibuat (RM)
            Else
                rs!jumlah_perlu_bayar = Null 'Jumlah Bayaran Yang Perlu Dibuat (RM)
            End If
            rs!kuantiti_barang = 1 'Kuantiti Barang Yang Dijual
            If Frm94.TB4 <> vbNullString Then 'Jumlah Berat Barang Yang Dijual
                rs!JUMLAH_BERAT = Format(Frm94.TB4, "0.00")
            Else
                rs!JUMLAH_BERAT = Null
            End If
            
            If Frm94.CB3 = 1 Then
            
                If Frm94.TB23 <> vbNullString Then
                    rs!gst_zr_harga = Format(Frm94.TB23, "0.00") 'Harga Keseluruhan Bagi Barang ZR
                Else
                    rs!gst_zr_harga = Null 'Harga Keseluruhan Bagi Barang ZR
                End If
                If Frm94.TB24 <> vbNullString Then
                    rs!gst_zr_cukai = Format(Frm94.TB24, "0.00") 'Jumlah Cukai Bagi ZR
                Else
                    rs!gst_zr_cukai = Null 'Jumlah Cukai Bagi ZR
                End If
                rs!gst_sr_harga = Format(0, "0.00") 'Harga Keseluruhan Bagi Barang SR
                rs!gst_sr_cukai = Format(0, "0.00") 'Jumlah Cukai Bagi SR

            
            ElseIf Frm94.CB4 = 1 Or Frm94.CB5 = 1 Then

                rs!gst_zr_harga = Format(0, "0.00") 'Harga Keseluruhan Bagi Barang ZR
                rs!gst_zr_cukai = Format(0, "0.00") 'Jumlah Cukai Bagi ZR
                If Frm94.TB23 <> vbNullString Then
                    rs!gst_sr_harga = Format(Frm94.TB23, "0.00") 'Harga Keseluruhan Bagi Barang SR
                Else
                    rs!gst_sr_harga = Null 'Harga Keseluruhan Bagi Barang SR
                End If
                If Frm94.TB24 <> vbNullString Then
                    rs!gst_sr_cukai = Format(Frm94.TB24, "0.00") 'Jumlah Cukai Bagi SR
                Else
                    rs!gst_sr_cukai = Null 'Jumlah Cukai Bagi SR
                End If
                
            End If
            rs!caj_pos = "0.00"
            rs!no_tracking = Null
            rs!no_pekerja = Frm94_LM_EMP_NO 'No. Pekerja
            If Frm94_LM_No_RUJUKAN_CUST <> vbNullString Then
                rs!no_rujukan_pembeli = Frm94_LM_No_RUJUKAN_CUST 'No. Rujukan Pembeli
            Else
                rs!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
            End If
            rs!no_rujukan_agen_dropship = Null 'No. Rujukan Agen Dropship

            If Frm94.CB2 = 1 Then
            
                Frm94_LM_Flag_TRADE_IN = 1 '0 : Tiada Urusan Trade In , 1 : Ada Urusan Trade In
                rs!flag_trade_in = 1 '0 : Tiada Urusan Trade in , 1 : Ada Urusan Trade In
                rs!jenis_trade_in = 1 '1 : Trade in (Voucher) , 2 : Belian dengan trade in
                
                If Frm94.L4_Text <> vbNullString Then 'No. Resit Trade In
                    rs!no_resit_trade_in = Frm94.L4_Text
                Else
                    rs!no_resit_trade_in = Null
                End If

                If Frm94.TB14 <> vbNullString Then
                    rs!jumlah_trade_in = Format(Frm94.TB14, "0.00") 'No. Resit Trade In
                Else
                    rs!jumlah_trade_in = Null 'No. Resit Trade In
                End If
            Else
                rs!flag_trade_in = 0 '0 : Tiada Urusan Trade in , 1 : Ada Urusan Trade In
                rs!no_resit_trade_in = Null 'No. Resit Trade In
                rs!jumlah_trade_in = Null 'Jumlah Resit Trade In (RM)
                rs!jenis_trade_in = Null '1 : Trade in (Voucher) , 2 : Belian dengan trade in
            End If
            rs!invoice_type = 0 '0 : Unlimited , Selain 0 (Limited : Mengikut nombor yang dimasukkan)

'Zakaria&Sons
'1 : Pembeli biasa
'2 : Ahli biasa
'3 : Silver
'4 : Gold
'5 : Platinum

            rs!kategori_pembeli = Frm94_LM_KATEGORI
            rs!jualan_online = 0
            rs!point_ari_nashi = 0
            rs!jumlah_point = 0
            rs!kupon_diskaun = "0.00"
            rs!kadar_peroleh_point = 0
            rs!kadar_tebus_point = 0
            rs!kadar_diskaun = Format(0, "0.00") 'Kadar diskaun per gram
            rs!Status = 1
            rs!status_r = 0
            rs!terminal = G_TERMINAL
            rs!cawangan = G_KEDAI
            rs!write_timestamp = LM_NOW
            rs!Menu = 3
            
            DATA_SAVE = 1
            rs.Update
        Else
            
            LM_RE_GEN_INVOICE_NO = 1
            Frm94_LM_No_RESIT_TEMPAHAN = Frm94_LM_No_RESIT_TEMPAHAN + 1
            If Frm94.CB9 = 0 Then Frm94.L10_Text = Frm94_LM_No_RESIT_TEMPAHAN
            If Frm94.CB9 = 1 Then Frm94.L32_Text = Frm94_LM_No_RESIT_TEMPAHAN
            
            rs.Close
            Set rs = Nothing
            GoTo Re_Gen_No_Rujukan:
        End If
        
        rs.Close
        Set rs = Nothing
'###Masukkan Jualan Ke Dalam Table Akaun Jualan### - End
    
'Jika terdapat perubahan pada no. invoice semasa masukkan data ke dalam table 22_jualan ### - Start
        If LM_RE_GEN_INVOICE_NO = 1 Then

            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 40_tempahan_deposit where no_rujukan_tempahan='" & Frm94_LM_No_RUJUKAN_TEMPAHAN & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
            
                If Frm94.CB9 = 0 Then
                
                    rs!invoice_siap = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm94_LM_No_RESIT_TEMPAHAN, "000000") 'No. invoice rasmi
                        
                Else
                
                    rs!invoice_siap = "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm94_LM_No_RESIT_TEMPAHAN, "000000") 'No. invoice tidak rasmi
                
                End If
                
                rs.Update
                
            End If
            
            rs.Close
            Set rs = Nothing
            
        End If
'Jika terdapat perubahan pada no. invoice semasa masukkan data ke dalam table 22_jualan ### - End

'### Masukkan Data Ke Dalam Tempahan Siap ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 42_tempahan_siap", cn, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        If Frm94.L9_Text <> vbNullString Then
            rs!no_rujukan_tempahan = Frm94.L9_Text 'No. Rujukan Tempahan
        Else
            rs!no_rujukan_tempahan = Null
        End If
        rs!no_resit_tempahan = LM_NO_INVOICE 'No. Resit Tempahan
        If Frm94.L6_Text = "Baki (RM)                                 :" Then
            rs!flag_bayaran = 0 '0 : Pembeli Bayar , 1 : Kedai Bayar
        ElseIf Frm94.L6_Text = "Lebihan Kedai Perlu Bayar (RM)  :" Then
            rs!flag_bayaran = 1 '0 : Pembeli Bayar , 1 : Kedai Bayar
        End If
            
        If Frm94.L7_Text <> vbNullString Then
            rs!jenis_tempahan = Frm94.L7_Text 'Jenis Tempahan , 0 : Tempahan Barang Baru , 1 : Tempahan Barang Kedai
        Else
            rs!jenis_tempahan = 1
        End If
        If Frm94.L15_Text <> vbNullString Then
            rs!type_barang_kemas = Frm94.L15_Text 'Jenis Barang Kemas , 0 : Barang Kemas , 1 : Barang Permata
        Else
            rs!type_barang_kemas = 0
        End If
        If Frm94.TB2 <> vbNullString Then
            rs!no_siri_Produk = Frm94.TB2 'No. Siri Produk
        Else
            rs!no_siri_Produk = 0
        End If
        If Frm94.L3_Text <> vbNullString Then
            rs!kategori_Produk = Frm94.L3_Text 'Kategori Produk
        Else
            rs!kategori_Produk = 0
        End If
        If Frm94_KOD_PURITY <> vbNullString Then
            rs!purity = Frm94_KOD_PURITY 'Purity
        Else
            rs!purity = Null 'Purity
        End If
        If Frm94_DULANG <> vbNullString Then
            rs!dulang = Frm94_DULANG 'Dulang
        Else
            rs!dulang = Null 'Dulang
        End If
        If Frm94.L15_Text = 0 Then 'Jenis Barang Kemas , 0 : Barang Kemas , 1 : Barang Permata
            If Frm94.TB3 <> vbNullString Then
                rs!Berat_Asal = Format(Frm94.TB3, "0.00") 'Berat Asal
            Else
                rs!Berat_Asal = Null
            End If
            If Frm94.TB4 <> vbNullString Then
                rs!berat_jualan = Format(Frm94.TB4, "0.00") 'Berat Jualan
            Else
                rs!berat_jualan = Null
            End If
            If Frm94.TB5 <> vbNullString Then
                rs!harga_Semasa = Format(Frm94.TB5, "0.00") 'Harga Semasa
            Else
                rs!harga_Semasa = Null
            End If
        Else
            rs!Berat_Asal = Null
            rs!berat_jualan = Null
            rs!harga_Semasa = Null
        End If
        If Frm94.TB6 <> vbNullString Then
            rs!UPAH = Format(Frm94.TB6, "0.00") 'Upah
        Else
            rs!UPAH = Null
        End If
        If Frm94.TB7 <> vbNullString Then
            rs!harga_asal = Format(Frm94.TB7, "0.00") 'Harga Asal
        Else
            rs!harga_asal = Null
        End If
        If Frm94.TB8 <> vbNullString Then
            rs!adjustment = Format(Frm94.TB8, "0.00") 'Adjustment
        Else
            rs!adjustment = Null
        End If
        If Frm94.TB9 <> vbNullString Then
            rs!harga = Format(Frm94.TB9, "0.00") 'Harga Jualan
        Else
            rs!harga = Null
        End If
        
'1 : Pelanggan
'2:  Member
'3:  RAF
'4:  Pengedar
'5:  Normal Dealer
'6:  Master Dealer

        rs!kategori_pembeli = Frm94_LM_KATEGORI
        If Frm94_LM_NAMA <> vbNullString Then
            rs!Nama = Frm94_LM_NAMA 'Maklumat Pembeli : Nama Pembeli
        Else
            rs!Nama = Null
        End If
        If Frm94_LM_IC <> vbNullString Then
            rs!no_ic = Frm94_LM_IC 'Maklumat Pembeli : No. IC
        Else
            rs!no_ic = Null
        End If
        If Frm94_LM_No_TEL <> vbNullString Then
            rs!no_tel = Frm94_LM_No_TEL 'Maklumat Pembeli : No. Telefon
        Else
            rs!no_tel = Null
        End If
        If Frm94_LM_No_RUJUKAN_CUST <> vbNullString Then
            rs!no_rujukan_pelanggan = Frm94_LM_No_RUJUKAN_CUST 'Maklumat Pembeli : No. Rujukan Pembeli
        Else
            rs!no_rujukan_pelanggan = Null
        End If
        If Frm94.CB2 = 1 Then
            rs!flag_trade_in = 1 '0 : Tiada Bayaran Deposit Menggunakan Trade In , 1 : Ada Bayaran Deposit Menggunakan Trade In
            Frm94_LM_Flag_TRADE_IN = 1
            
            If Frm94.L4_Text <> vbNullString Then
                rs!no_resit_trade_in = Frm94.L4_Text 'No. Resit Trade In
            Else
                rs!no_resit_trade_in = Null
            End If
            If Frm94.TB14 <> vbNullString Then
                rs!nilaian_trade_in = Format(Frm94.TB14, "0.00") 'Jumlah Nilaian Trade In
            Else
                rs!nilaian_trade_in = Null
            End If
        Else
            Frm94_LM_Flag_TRADE_IN = 0
            
            rs!flag_trade_in = 0
            rs!no_resit_trade_in = Null
            rs!nilaian_trade_in = Null
        End If
        If Frm94.TB10 <> vbNullString Then
            rs!JUMLAH_HARGA_JUALAN = Format(Frm94.TB10, "0.00") 'Maklumat Baki Bayaran : Jumlah Harga Jualan (RM)
        Else
            rs!JUMLAH_HARGA_JUALAN = Null
        End If
        If Frm94.TB11 <> vbNullString Then
            rs!bayaran_sudah_jelas = Format(Frm94.TB11, "0.00") 'Maklumat Baki Bayaran : Bayaran Sudah Jelas Dari Deposit (RM)
        Else
            rs!bayaran_sudah_jelas = Null
        End If
        If Frm94.TB12 <> vbNullString Then
            rs!baki = Format(Frm94.TB12, "0.00") 'Maklumat Baki Bayaran : Baki Bayaran (RM)
        Else
            rs!baki = Null
        End If
        If Frm94.TB17 <> vbNullString Then
            rs!jumlah_gst = Format(Frm94.TB17, "0.00") 'Maklumat Baki Bayaran : Jumlah GST (RM)
        Else
            rs!jumlah_gst = Null
        End If
        If Frm94.TB24 <> vbNullString Then
            rs!harga_dengan_gst = Format(Frm94.TB24, "0.00") 'Jumlah harga barang dengan GST
        Else
            rs!harga_dengan_gst = Null
        End If
        'If Frm94.TB18 <> vbNullString Then
        '    rs!baki_adjustment = Format(Frm94.TB18, "0.00") 'Maklumat Baki Bayaran : Adjustment Bagi Baki (RM)
        'Else
        '    rs!baki_adjustment = Null
        'End If
        'If Frm94.TB19 <> vbNullString Then
        '    rs!jumlah_baki_terakhir = Format(Frm94.TB19, "0.00") 'Maklumat Baki Bayaran : Jumlah Baki Terakhir (RM)
        'Else
        '    rs!jumlah_baki_terakhir = Null
        'End If
        rs!tarikh = Frm94.DTPicker1 'Tarikh Tempahan
        rs!no_pekerja = Frm94_LM_EMP_NO 'No. Pekerja
        
        If Frm94.CB3 = 1 Then
        
            rs!gst_include = 0
        
        ElseIf Frm94.CB4 = 1 Then
            
            rs!gst_include = 1
            
        ElseIf Frm94.CB5 = 1 Then
            
            rs!gst_include = 2
            
        End If
            
        If Frm94.CB5 = 0 Then
            rs!gst_include = Null
        ElseIf Frm94.CB5 = 1 Then
            rs!gst_include = "**Harga Termasuk GST"
        End If
        If Frm94.TB23 <> vbNullString Then
            rs!harga_tanpa_gst = Format(Frm94.TB23, "0.00") 'Harga Keseluruhan Tanpa GST (RM)
        Else
            rs!harga_tanpa_gst = Null 'Harga Keseluruhan Tanpa GST (RM)
        End If

        rs!terminal = G_TERMINAL
        rs!write_timestamp = LM_NOW
        rs!status_invoice = 1 '0 : Tidak aktif (dibatalkan) , 1:  Aktif
        If Frm94.CB9 = 0 Then
            rs!bil_rasmi = 1 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
        Else
            rs!bil_rasmi = 0 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
        End If
        rs!no_staff = G_LOGIN_USER 'No. staff yang memasukkan/edit data (Login)
        rs!cawangan = G_KEDAI
        
        rs.Update
        
        rs.Close
        Set rs = Nothing

'### Masukkan Data Ke Dalam Tempahan Siap ### - End

'###Update Data Simpanan Duit Pelanggan### - Start
        If Frm94_LM_Flag_SIMPANAN = 1 Then '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm94_LM_No_RUJUKAN_CUST & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
            
                G_ID = rs!ID
                Call recovery_senarai_pelanggan
                
                Frm94_LM_JUMLAH_SIMPANAN = frm130.L26_Text  'Jumlah Simpanan Yang Ada
                Frm94_LM_GUNA_SIMPAN = frm130.TB21  'Jumlah Simpanan Yang Hendak Digunakan
                
                rs!baki_simpanan = Format(Frm94_LM_JUMLAH_SIMPANAN - Frm94_LM_GUNA_SIMPAN, "0.00") 'Baki Simpanan
                
                rs!write_timestamp2 = LM_NOW
                rs!no_staff = Frm94_LM_EMP_NO 'No. Pekerja
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
            rs!tarikh = Frm94.DTPicker1 'Tarikh
            rs!jenis = 1 '0 : Simpanan , 1 : Penggunaan Duit
            rs!no_rujukan_pelanggan = Frm94_LM_No_RUJUKAN_CUST 'No. Rujukan Pelanggan
            rs!no_resit = LM_NO_INVOICE 'No. Resit Tempahan
            rs!jumlah = Format(frm130.TB21, "0.00") 'Jumlah Simpanan Yang Digunakan (RM)
            rs!jenis_penggunaan = 4 '0 : Belian Barangan Kemas , 1 : Ansuran , 2 : Tempahan (Deposit) , 3 : Servis , 4 : Tempahan (Ambilan Barang)
            rs!no_rujukan_pekerja = Frm94_LM_EMP_NO 'No. Pekerja
            rs!terminal = G_TERMINAL
            rs!write_timestamp = LM_NOW
            rs!jenis_urusan = G_JENIS_URUSAN
            rs!cawangan = G_CAWANGAN
            rs!Status = 1
            rs.Update
            
            rs.Close
            Set rs = Nothing
           
        End If
'###Update Data Simpanan Duit Pelanggan### - End

'### Update Maklumat Trade In ### - Start
        If Frm94_LM_Flag_TRADE_IN = 1 Then '0 : Tiada Urusan Trade In , 1 : Ada Urusan Trade In
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & Frm94.L4_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
            
                G_ID = rs!ID
                Call recovery_16_gold_bar_belian
                
                rs!trade_in_status = 1
                rs!no_staff = Frm94_LM_EMP_NO 'No. Pekerja
                rs!terminal = G_TERMINAL
                rs!write_timestamp2 = LM_NOW
                rs!jenis_urusan = G_JENIS_URUSAN
                rs!remarks = "Ubah status flag trade in bagi ambilan tempahan"
                rs.Update
                
            End If
            
            rs.Close
            Set rs = Nothing
            
        End If
'### Update Maklumat Trade In ### - End

'### Update Table Database Bagi Item Ini ### - Start

'10 : In Stock
'11 :  Sold
'12 : In Stock - Potong
'13 :  Sold -Potong
'14 :  Tempahan
'15 :  Ansuran
'16 :  Ar -Rahnu
'17 :  ETA
'18 :  Pinjaman
'19 : Terjual Secara Ansuran (Jelas)
'20 : Terjual Secara Ansuran - Potong (Jelas)
'21 : Terjual Secara Tempahan (Siap)
'22 : Terjual Secara Tempahan - Potong (Siap)

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where no_siri_produk='" & Frm94.TB2 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then

            G_ID = rs!ID
            Call recovery_data_database
                
            If Frm94.L15_Text = 0 Then
                If Format(Frm94_LM_ERR_BERAT_ASAL, "0.00") = Format(Frm94_LM_ERR_BERAT_JUALAN, "0.00") Then
                    rs!StatusItem = 21
                    rs!beza_berat = "0.00"
                Else
                    rs!StatusItem = 22
                    rs!beza_berat = Format(Frm94_LM_ERR_BERAT_ASAL - Frm94_LM_ERR_BERAT_JUALAN, "0.00") 'Beza Berat (Baki)
                End If
            Else
                rs!StatusItem = 21
            End If
            
            rs!write_timestamp2 = LM_NOW
            rs!no_pekerja = Frm94_LM_EMP_NO
            rs!terminal = G_TERMINAL
            rs!Menu = 3
            'rs!cawangan = G_KEDAI
            rs.Update
            
        End If
        
        rs.Close
        Set rs = Nothing
'### Update Table Database Bagi Item Ini ### - End

'### Update Log ### - Start
        user = MDI_frm1.L3_Text
        LogAct_Memory = "[" & G_LOGIN_USER & "] Ambilan tempahan. No. Invoice [" & LM_NO_INVOICE & "]"
        LogDate_Memory = LM_NOW
        Call UpdateLog_Database
'### Update Log ### - End
        
        'Set rs = New ADODB.Recordset
        'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        'rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
        
        'If Not rs.EOF Then
        '    If rs!Default1 = "Default" Then
            
        '        rs!no_rujukan_book = Frm94_LM_No_RUJUKAN_TEMPAHAN + 1
                'If Frm94.CB9 = 0 Then rs!ResitNo = Frm94_LM_No_RESIT_TEMPAHAN + 1
                'If Frm94.CB9 = 1 Then rs!no_rujukan_tak_rasmi = Frm94_LM_No_RESIT_TEMPAHAN + 1
                
        '        rs.Update
                
        '    End If
        'End If
        
        'rs.Close
        'Set rs = Nothing
        
        Frm93.Frame1.Visible = False
        
        Frm93.Show
        Unload Frm94
        
        MDI_frm1.L5_Text = 8
        
        Note = "Data ambilan tempahan siap telah berjaya disimpan." & vbCrLf & _
                "Adakah Anda Ingin Cetak Invoice Bayaran ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
            G_PREVIEW = 1
            G_No_INV_BOOK = vbNullString
            G_No_INV_BOOK = LM_NO_INVOICE 'No. Invoice
            Call Frm94_invoice_siap_tempahan
        End If
    End If
End If
End Sub

Private Sub CMD6_Click()
'on error resume next
'frm130.TB33 = Format(Frm92.L10_Text, "#,##0.00")
frm130.Show vbModal
End Sub

Private Sub Form_Load()
'on error resume next
Frm94.Picture = MDI_frm1.Picture
Frm94.Pic1 = MDI_frm1.Picture

frm130.L31_Text.BackStyle = 0
frm130.L32_Text.BackStyle = 0
frm130.L81_Text.BackStyle = 0
frm130.L82_Text.BackStyle = 0
End Sub
Private Sub L12_Text_Click()
'on error resume next
Frm93.Show
Unload Frm94

MDI_frm1.L5_Text = 8
End Sub















Private Sub L8_Text_Change()
'on error resume next
Call frm94_kiraan_cukai_gst
Call Frm94_kira_caj_kad_kredit
End Sub
Private Sub TB1_Change()
'on error resume next
If Frm94.CB1 = 1 And Frm94.TB1 <> vbNullString And Frm94.L7_Text = 0 Then
    Frm94.Tmr2.Enabled = False
    Frm94.Tmr2.Enabled = True
    Frm94.Tmr2.Interval = 100
End If
End Sub
Private Sub TB10_Change()
'on error resume next
Call frm94_kiraan_cukai_gst
End Sub
Private Sub TB11_Change()
'on error resume next
Call frm94_kira_baki
End Sub



Private Sub TB22_Change()
'on error resume next
Call frm94_kira_baki
End Sub

Private Sub TB24_Change()
'On Error Resume Next
Call frm94_kira_baki
End Sub








Private Sub TB4_Change()
'on error resume next
Call frm94_kira_harga_emas
End Sub
Private Sub TB5_Change()
'on error resume next
Call frm94_kira_harga_emas
End Sub
Private Sub TB6_Change()
'on error resume next
Call frm94_kira_harga_emas
End Sub
Private Sub TB7_Change()
'on error resume next
Call frm94_kira_harga_bersih
End Sub
Private Sub TB8_Change()
'on error resume next
Call frm94_kira_harga_bersih
End Sub
Private Sub TB9_Change()
'on error resume next
If Frm94.TB9 <> vbNullString And IsNumeric(Frm94.TB9) Then
    Frm94.TB10 = Format(Frm94.TB9, "#,##0.00")
Else
    Frm94.TB10 = "0.00"
End If
End Sub
Private Sub Tmr2_Timer()
'On Error Resume Next
If Frm94.CB1 = 1 And Frm94.TB1 <> vbNullString And Frm94.L7_Text = 0 And Frm94.Tmr2.Enabled = True Then
    If Frm94.Tmr2.Interval = 100 Then
        If InStr(1, Frm94.TB1, "'") <> 0 Then
            MsgBox "No. Siri Produk Mengandungi Simbol Yang Tidak Sah , ['].", vbInformation, "Info"
            
            Frm94.TB1 = vbNullString
            Exit Sub
        End If
        
        Call Frm94_Call_Product_Detail
    End If
End If
End Sub
