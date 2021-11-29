VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm87 
   Caption         =   "Jualan Secara Ansuran"
   ClientHeight    =   12915
   ClientLeft      =   60
   ClientTop       =   -26220
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
   Icon            =   "Frm87.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic7 
      BorderStyle     =   0  'None
      Height          =   12015
      Left            =   11040
      ScaleHeight     =   12015
      ScaleWidth      =   21480
      TabIndex        =   90
      Top             =   240
      Visible         =   0   'False
      Width           =   21480
      Begin VB.CommandButton CMD15 
         Caption         =   "Menu Sebelum"
         Height          =   375
         Left            =   2160
         MouseIcon       =   "Frm87.frx":0ECA
         MousePointer    =   99  'Custom
         TabIndex        =   178
         ToolTipText     =   "Kembali ke menu sebelum."
         Top             =   10920
         Width           =   2055
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   10755
         Left            =   6480
         TabIndex        =   177
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   480
         Width           =   14385
         _ExtentX        =   25374
         _ExtentY        =   18971
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
      Begin VB.Label Label88 
         BackStyle       =   0  'Transparent
         Caption         =   "Maklumat pembeli ini."
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   93
         Top             =   240
         Width           =   6585
      End
      Begin VB.Label L29_Text 
         Caption         =   "L29_Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8295
         Left            =   120
         TabIndex        =   92
         Top             =   600
         Width           =   7095
      End
      Begin VB.Label Label85 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai rekod bayaran bagi pembeli ini."
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   6480
         TabIndex        =   91
         Top             =   240
         Width           =   6585
      End
   End
   Begin VB.PictureBox Pic5 
      BorderStyle     =   0  'None
      Height          =   9855
      Left            =   5400
      ScaleHeight     =   9855
      ScaleWidth      =   14775
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   14775
      Begin VB.CheckBox CB27 
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
         Left            =   2400
         TabIndex        =   116
         Top             =   3735
         Width           =   200
      End
      Begin VB.TextBox TB42 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   115
         Text            =   "TB42"
         Top             =   4440
         Width           =   1515
      End
      Begin VB.CommandButton CMD14 
         BackColor       =   &H000080FF&
         Caption         =   "Batal / Keluar"
         Height          =   400
         Left            =   12120
         MaskColor       =   &H00400000&
         Style           =   1  'Graphical
         TabIndex        =   112
         ToolTipText     =   "Batal urusan pembayaran ansuran ini."
         Top             =   1680
         Width           =   2385
      End
      Begin VB.TextBox TB27 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2760
         TabIndex        =   70
         Text            =   "TB27"
         Top             =   7035
         Width           =   1725
      End
      Begin VB.TextBox TB28 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2760
         TabIndex        =   69
         Text            =   "TB28"
         Top             =   7410
         Width           =   1725
      End
      Begin VB.TextBox TB29 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2760
         TabIndex        =   68
         Text            =   "TB29"
         Top             =   7800
         Width           =   1725
      End
      Begin VB.TextBox TB30 
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   67
         Text            =   "TB30"
         Top             =   8520
         Width           =   1725
      End
      Begin VB.TextBox TB31 
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   66
         Text            =   "TB31"
         Top             =   8880
         Width           =   1725
      End
      Begin VB.TextBox TB32 
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "TB32"
         Top             =   9000
         Width           =   1725
      End
      Begin VB.TextBox TB40 
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   64
         Text            =   "TB40"
         Top             =   8520
         Width           =   1725
      End
      Begin VB.TextBox TB39 
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   "TB39"
         Top             =   8115
         Width           =   1725
      End
      Begin VB.TextBox TB38 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   7200
         TabIndex        =   62
         Text            =   "TB38"
         Top             =   7440
         Width           =   1725
      End
      Begin VB.TextBox TB21 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   7200
         TabIndex        =   61
         Text            =   "TB21"
         Top             =   7080
         Width           =   1725
      End
      Begin VB.CheckBox CB22 
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
         TabIndex        =   47
         Top             =   3495
         Width           =   200
      End
      Begin VB.CheckBox CB23 
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
         Left            =   2280
         TabIndex        =   46
         Top             =   3495
         Width           =   200
      End
      Begin VB.TextBox TB20 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "TB20"
         Top             =   6120
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox TB19 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   6990
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "TB19"
         Top             =   1440
         Width           =   2000
      End
      Begin VB.CheckBox CB21 
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
         Left            =   5040
         TabIndex        =   38
         Top             =   735
         Width           =   200
      End
      Begin VB.CheckBox CB20 
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
         TabIndex        =   36
         Top             =   750
         Width           =   200
      End
      Begin VB.TextBox TB18 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "TB18"
         Top             =   5880
         Width           =   1515
      End
      Begin VB.TextBox TB17 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2310
         TabIndex        =   31
         Text            =   "TB17"
         Top             =   5520
         Width           =   1515
      End
      Begin VB.CommandButton CMD10 
         BackColor       =   &H000080FF&
         Caption         =   "Batal / Keluar"
         Height          =   400
         Left            =   12120
         MaskColor       =   &H00400000&
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Batal urusan pembayaran ansuran ini."
         Top             =   1680
         Width           =   2385
      End
      Begin VB.ComboBox CBB2 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   10995
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1080
         Width           =   3400
      End
      Begin VB.CheckBox CB19 
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
         Left            =   12000
         TabIndex        =   20
         Top             =   7335
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.CheckBox CB18 
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
         Left            =   9960
         TabIndex        =   19
         Top             =   7335
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.TextBox TB14 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "TB14"
         Top             =   5160
         Width           =   1515
      End
      Begin VB.TextBox TB13 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "TB13"
         Top             =   4800
         Width           =   1515
      End
      Begin VB.TextBox TB12 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   2190
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "TB12"
         Top             =   1440
         Width           =   2000
      End
      Begin VB.PictureBox Pic6 
         BorderStyle     =   0  'None
         Height          =   1275
         Left            =   240
         ScaleHeight     =   1275
         ScaleWidth      =   4095
         TabIndex        =   7
         Top             =   1920
         Visible         =   0   'False
         Width           =   4095
         Begin VB.TextBox TB16 
            Alignment       =   2  'Center
            BackColor       =   &H8000000A&
            Height          =   360
            Left            =   1950
            Locked          =   -1  'True
            TabIndex        =   10
            Text            =   "TB16"
            Top             =   480
            Width           =   2000
         End
         Begin VB.TextBox TB15 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   1950
            TabIndex        =   8
            Text            =   "TB15"
            Top             =   120
            Width           =   2000
         End
         Begin VB.Label L19_Text 
            Caption         =   "L19_Text"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   2280
            TabIndex        =   45
            Top             =   840
            Width           =   2265
         End
         Begin VB.Label Label58 
            BackStyle       =   0  'Transparent
            Caption         =   "**Baki berat adalah (g) :"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   0
            TabIndex        =   42
            Top             =   840
            Width           =   2265
         End
         Begin VB.Label Label44 
            BackStyle       =   0  'Transparent
            Caption         =   "Berat Diperolehi      g"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   0
            TabIndex        =   11
            Top             =   510
            Width           =   2265
         End
         Begin VB.Label Label43 
            BackStyle       =   0  'Transparent
            Caption         =   "Harga Semasa  RM/g"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   0
            TabIndex        =   9
            Top             =   150
            Width           =   2265
         End
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   360
         Left            =   10995
         TabIndex        =   26
         Top             =   720
         Width           =   3405
         _ExtentX        =   6006
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
      Begin VB.CommandButton CMD13 
         BackColor       =   &H000080FF&
         Caption         =   "Simpan Data"
         Height          =   400
         Left            =   9600
         MaskColor       =   &H00400000&
         Style           =   1  'Graphical
         TabIndex        =   113
         ToolTipText     =   "Simpan data pembayaran ansuran ini ke dalam sistem."
         Top             =   1680
         Width           =   2385
      End
      Begin VB.CommandButton CMD9 
         BackColor       =   &H000080FF&
         Caption         =   "Simpan Data"
         Height          =   400
         Left            =   9600
         MaskColor       =   &H00400000&
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Simpan data pembayaran ansuran ini ke dalam sistem."
         Top             =   1680
         Width           =   2385
      End
      Begin VB.Shape Shape18 
         Height          =   615
         Left            =   2220
         Top             =   3390
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Termasuk GST"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   2640
         TabIndex        =   117
         Top             =   3720
         Width           =   3825
      End
      Begin VB.Label L39_Text 
         Caption         =   "L39_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9600
         TabIndex        =   114
         Top             =   5880
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "**Baki upah adalah (RM) :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   5160
         TabIndex        =   108
         Top             =   1920
         Width           =   2385
      End
      Begin VB.Label L20_Text 
         Caption         =   "L20_Text"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   7560
         TabIndex        =   107
         Top             =   1920
         Width           =   1545
      End
      Begin VB.Label L33_Text 
         Caption         =   "L33_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9600
         TabIndex        =   95
         Top             =   5520
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label L30_Text 
         Caption         =   "L30_Text"
         ForeColor       =   &H00000000&
         Height          =   2445
         Left            =   9600
         TabIndex        =   94
         Top             =   2520
         Width           =   5025
      End
      Begin VB.Label Label86 
         BackStyle       =   0  'Transparent
         Caption         =   "**Baki ansuran adalah  (RM):"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   360
         TabIndex        =   89
         Top             =   2880
         Width           =   2745
      End
      Begin VB.Label L28_Text 
         Caption         =   "L28_Text"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   3000
         TabIndex        =   88
         Top             =   2880
         Width           =   1545
      End
      Begin VB.Label Label84 
         BackStyle       =   0  'Transparent
         Caption         =   "Tunai                        RM :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   360
         TabIndex        =   85
         Top             =   7065
         Width           =   2715
      End
      Begin VB.Label Label80 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank In                     RM :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   360
         TabIndex        =   84
         Top             =   7440
         Width           =   2715
      End
      Begin VB.Label Label72 
         BackStyle       =   0  'Transparent
         Caption         =   "Kad Kredit                 RM :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   360
         TabIndex        =   83
         Top             =   7785
         Width           =   2715
      End
      Begin VB.Label Label76 
         BackStyle       =   0  'Transparent
         Caption         =   "Maklumat Cara Bayaran"
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
         Left            =   360
         TabIndex        =   82
         Top             =   6720
         Width           =   3585
      End
      Begin VB.Label Label77 
         BackStyle       =   0  'Transparent
         Caption         =   "Cas Kredit Kad           RM :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   360
         TabIndex        =   81
         Top             =   8535
         Width           =   2715
      End
      Begin VB.Label Label79 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Potongan       RM :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   360
         TabIndex        =   80
         Top             =   8895
         Width           =   2715
      End
      Begin VB.Label Label81 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Bayaran          RM :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   4800
         TabIndex        =   79
         Top             =   9030
         Width           =   2715
      End
      Begin VB.Shape Shape16 
         Height          =   975
         Left            =   240
         Top             =   8400
         Width           =   4395
      End
      Begin VB.Label Label114 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Potongan        RM :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   4800
         TabIndex        =   78
         Top             =   8505
         Width           =   2715
      End
      Begin VB.Label Label116 
         BackStyle       =   0  'Transparent
         Caption         =   "Cas Debit Kad             RM :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   4800
         TabIndex        =   77
         Top             =   8130
         Width           =   2715
      End
      Begin VB.Shape Shape14 
         Height          =   900
         Left            =   4680
         Top             =   8040
         Width           =   4395
      End
      Begin VB.Label Label118 
         BackStyle       =   0  'Transparent
         Caption         =   "Kad Debit                   RM :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   4800
         TabIndex        =   76
         Top             =   7485
         Width           =   2715
      End
      Begin VB.Label L31_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2040
         TabIndex        =   75
         Top             =   8160
         Width           =   600
      End
      Begin VB.Label L32_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6480
         TabIndex        =   74
         Top             =   7800
         Width           =   600
      End
      Begin VB.Label Label82 
         BackStyle       =   0  'Transparent
         Caption         =   "Simpanan Duit Di Kedai Sebanyak : RM"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4800
         TabIndex        =   73
         Top             =   6840
         Width           =   3435
      End
      Begin VB.Label Label83 
         BackStyle       =   0  'Transparent
         Caption         =   "Duit Simpanan Di Kedai RM:"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   4800
         TabIndex        =   72
         Top             =   7125
         Width           =   2715
      End
      Begin VB.Label L27_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L27_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8160
         TabIndex        =   71
         Top             =   6840
         Width           =   1635
      End
      Begin VB.Label L24_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L24_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7680
         TabIndex        =   60
         Top             =   5280
         Width           =   1680
      End
      Begin VB.Label L23_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L23_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6120
         TabIndex        =   59
         Top             =   5280
         Width           =   2040
      End
      Begin VB.Label L26_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L26_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7680
         TabIndex        =   58
         Top             =   4920
         Width           =   1680
      End
      Begin VB.Label L25_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L25_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6120
         TabIndex        =   57
         Top             =   4920
         Width           =   2040
      End
      Begin VB.Label Label69 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cukai GST"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7680
         TabIndex        =   56
         Top             =   4560
         Width           =   1680
      End
      Begin VB.Label Label68 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Harga"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6120
         TabIndex        =   55
         Top             =   4560
         Width           =   2040
      End
      Begin VB.Label Label65 
         BackStyle       =   0  'Transparent
         Caption         =   "Standard Rated SR (RM):"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4080
         TabIndex        =   54
         Top             =   5310
         Width           =   2760
      End
      Begin VB.Label Label122 
         BackStyle       =   0  'Transparent
         Caption         =   "Zero Rated ZR(L)    (RM):"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4080
         TabIndex        =   53
         Top             =   4920
         Width           =   2640
      End
      Begin VB.Shape Shape13 
         Height          =   1815
         Left            =   4800
         Top             =   600
         Width           =   4575
      End
      Begin VB.Shape Shape12 
         Height          =   855
         Left            =   9825
         Top             =   7200
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Label Label66 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah GST : RM"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   10200
         TabIndex        =   52
         Top             =   7680
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label L22_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L22_Text"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   11760
         TabIndex        =   51
         Top             =   7680
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Shape Shape11 
         Height          =   975
         Left            =   120
         Top             =   3360
         Width           =   4575
      End
      Begin VB.Label L21_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L21_Text"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   2160
         TabIndex        =   50
         Top             =   3960
         Width           =   1545
      End
      Begin VB.Label Label64 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah GST : RM"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   600
         TabIndex        =   49
         Top             =   3960
         Width           =   1545
      End
      Begin VB.Shape Shape10 
         Height          =   2775
         Left            =   120
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label60 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah                 RM"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   360
         TabIndex        =   44
         Top             =   4470
         Width           =   2265
      End
      Begin VB.Shape Shape9 
         Height          =   1575
         Left            =   3960
         Top             =   4440
         Width           =   5295
      End
      Begin VB.Label Label56 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Upah        RM"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   5040
         TabIndex        =   41
         Top             =   1440
         Width           =   2265
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila tanda di sini jika bayaran adalah dibuat untuk bayaran UPAH ansuran emas."
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   5400
         TabIndex        =   39
         Top             =   720
         Width           =   3825
      End
      Begin VB.Label Label54 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila tanda di sini jika bayaran adalah dibuat untuk bayaran ansuran emas."
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   600
         TabIndex        =   37
         Top             =   735
         Width           =   3825
      End
      Begin VB.Label L18_Text 
         Alignment       =   2  'Center
         Caption         =   "L18_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   13560
         TabIndex        =   35
         Top             =   360
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label53 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Bayaran    RM"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   360
         TabIndex        =   34
         Top             =   5910
         Width           =   2265
      End
      Begin VB.Label Label52 
         BackStyle       =   0  'Transparent
         Caption         =   "Adjustment          RM"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   360
         TabIndex        =   32
         Top             =   5550
         Width           =   2265
      End
      Begin VB.Shape Shape8 
         Height          =   2085
         Left            =   120
         Top             =   4320
         Width           =   9255
      End
      Begin VB.Label Label51 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh *"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   9600
         TabIndex        =   28
         Top             =   720
         Width           =   2385
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pekerja*"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   9600
         TabIndex        =   27
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "Zero Rated ZR(L)         Standard Rated SR"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   10320
         TabIndex        =   22
         Top             =   7320
         Visible         =   0   'False
         Width           =   3945
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "Ringkasan GST"
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
         Left            =   4080
         TabIndex        =   21
         Top             =   4560
         Width           =   3495
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Asal          RM"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   360
         TabIndex        =   18
         Top             =   5190
         Width           =   2265
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah GST          RM"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   360
         TabIndex        =   16
         Top             =   4830
         Width           =   2265
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Ansuran    RM"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   2265
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "Maklumat Bayaran"
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
         TabIndex        =   12
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label L17_Text 
         Alignment       =   2  'Center
         Caption         =   "L17_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7680
         TabIndex        =   24
         Top             =   5640
         Width           =   840
      End
      Begin VB.Label Label59 
         BackStyle       =   0  'Transparent
         Caption         =   "Pengiraan GST adalah berdasarkan kadar @       %"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4080
         TabIndex        =   23
         Top             =   5640
         Width           =   4440
      End
      Begin VB.Shape Shape15 
         BorderWidth     =   2
         Height          =   2895
         Left            =   120
         Top             =   6600
         Width           =   9255
      End
      Begin VB.Label Label119 
         BackStyle       =   0  'Transparent
         Caption         =   "** Cas Kad Debit :         %"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4920
         TabIndex        =   86
         Top             =   7800
         Width           =   2760
      End
      Begin VB.Label Label120 
         BackStyle       =   0  'Transparent
         Caption         =   "** Cas Kad Kredit :         %"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         TabIndex        =   87
         Top             =   8160
         Width           =   2760
      End
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   "Zero Rated ZR(L)         Standard Rated (SR)"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   600
         TabIndex        =   48
         Top             =   3480
         Width           =   3825
      End
   End
   Begin VB.PictureBox Pic4 
      BorderStyle     =   0  'None
      Height          =   12015
      Left            =   600
      ScaleHeight     =   12015
      ScaleWidth      =   21480
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   21480
      Begin VB.CommandButton CMD12 
         BackColor       =   &H000080FF&
         Caption         =   "Paparan Semua Senarai"
         Height          =   400
         Left            =   7680
         MaskColor       =   &H00400000&
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   150
         Width           =   2625
      End
      Begin VB.CommandButton CMD11 
         BackColor       =   &H000080FF&
         Caption         =   "Carian"
         Height          =   400
         Left            =   17760
         MaskColor       =   &H00400000&
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   240
         Width           =   1785
      End
      Begin VB.TextBox TB41 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   14760
         TabIndex        =   102
         Text            =   "TB41"
         Top             =   240
         Width           =   2940
      End
      Begin VB.CheckBox CB26 
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
         Left            =   10800
         TabIndex        =   100
         Top             =   480
         Width           =   200
      End
      Begin VB.CheckBox CB25 
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
         Left            =   10800
         TabIndex        =   98
         Top             =   240
         Width           =   200
      End
      Begin VB.CheckBox CB24 
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
         Left            =   10800
         TabIndex        =   96
         Top             =   0
         Width           =   200
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   10635
         Left            =   120
         TabIndex        =   176
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   720
         Width           =   20685
         _ExtentX        =   36486
         _ExtentY        =   18759
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
      Begin VB.Label L35_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L35_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4560
         TabIndex        =   111
         Top             =   11400
         Width           =   1215
      End
      Begin VB.Label L34_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L34_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         TabIndex        =   110
         Top             =   11400
         Width           =   975
      End
      Begin VB.Label Label93 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan :                              Jumlah Berat (g) :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   109
         Top             =   11400
         Width           =   5535
      End
      Begin VB.Shape Shape17 
         Height          =   615
         Left            =   4440
         Top             =   75
         Width           =   6135
      End
      Begin VB.Label Label92 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila klik di sini untuk memaparkan semua senarai ansuran."
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   4680
         TabIndex        =   106
         Top             =   120
         Width           =   3225
      End
      Begin VB.Label Label91 
         BackStyle       =   0  'Transparent
         Caption         =   "Maklumat Carian :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   13080
         TabIndex        =   103
         Top             =   270
         Width           =   1905
      End
      Begin VB.Label Label90 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Siri Produk"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   11160
         TabIndex        =   101
         Top             =   450
         Width           =   2385
      End
      Begin VB.Label Label89 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Kad Pengenalan"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   11160
         TabIndex        =   99
         Top             =   225
         Width           =   2385
      End
      Begin VB.Label Label87 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   11160
         TabIndex        =   97
         Top             =   0
         Width           =   2385
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai pembeli barang kemas secara ansuran."
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   10785
      End
   End
   Begin VB.Timer Tmr3 
      Interval        =   100
      Left            =   0
      Top             =   1200
   End
   Begin VB.Timer Tmr2 
      Interval        =   100
      Left            =   0
      Top             =   600
   End
   Begin VB.Timer Tmr1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Pic2 
      BorderStyle     =   0  'None
      Height          =   11895
      Left            =   2160
      ScaleHeight     =   11895
      ScaleWidth      =   21480
      TabIndex        =   118
      Top             =   360
      Visible         =   0   'False
      Width           =   21480
      Begin VB.CommandButton CMD6 
         Caption         =   "Simpan Data"
         Height          =   375
         Left            =   9240
         MouseIcon       =   "Frm87.frx":11D4
         MousePointer    =   99  'Custom
         TabIndex        =   170
         ToolTipText     =   "Simpan data ke dalam sistem"
         Top             =   9480
         Width           =   2415
      End
      Begin VB.CommandButton CMD5 
         Caption         =   "Carian Data"
         Height          =   375
         Left            =   12480
         MouseIcon       =   "Frm87.frx":14DE
         MousePointer    =   99  'Custom
         TabIndex        =   175
         ToolTipText     =   "Carian data terperinci produk"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CommandButton CMD3 
         Caption         =   "Maklumat Pelanggan"
         Height          =   375
         Left            =   240
         MouseIcon       =   "Frm87.frx":17E8
         MousePointer    =   99  'Custom
         TabIndex        =   174
         ToolTipText     =   "Maklumat pembeli yang berdattar dengan kedai."
         Top             =   3960
         Width           =   2055
      End
      Begin VB.CommandButton CMD1 
         Caption         =   "Maklumat Pembeli"
         Height          =   375
         Left            =   240
         MouseIcon       =   "Frm87.frx":1AF2
         MousePointer    =   99  'Custom
         TabIndex        =   173
         ToolTipText     =   "Maklumat pembeli yang TIDAK berdattar dengan kedai."
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CommandButton CMD17 
         Caption         =   "Batal"
         Height          =   375
         Left            =   10680
         MouseIcon       =   "Frm87.frx":1DFC
         MousePointer    =   99  'Custom
         TabIndex        =   172
         ToolTipText     =   "Batal"
         Top             =   9480
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton CMD16 
         Caption         =   "Simpan Data"
         Height          =   375
         Left            =   8040
         MouseIcon       =   "Frm87.frx":2106
         MousePointer    =   99  'Custom
         TabIndex        =   171
         ToolTipText     =   "Simpan data ke dalam sistem"
         Top             =   9480
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox TB6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   9390
         TabIndex        =   131
         Text            =   "TB6"
         Top             =   3480
         Width           =   2000
      End
      Begin VB.CheckBox CB13 
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
         TabIndex        =   130
         Top             =   120
         Width           =   200
      End
      Begin VB.TextBox TB8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   9225
         TabIndex        =   129
         Text            =   "TB8"
         Top             =   1095
         Width           =   2940
      End
      Begin VB.TextBox TB2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   9390
         Locked          =   -1  'True
         TabIndex        =   128
         Text            =   "TB2"
         Top             =   2040
         Width           =   2000
      End
      Begin VB.TextBox TB3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   9390
         Locked          =   -1  'True
         TabIndex        =   127
         Text            =   "TB3"
         Top             =   2400
         Width           =   2000
      End
      Begin VB.TextBox TB4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   9390
         TabIndex        =   126
         Text            =   "TB4"
         Top             =   2760
         Width           =   2000
      End
      Begin VB.TextBox TB5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   9390
         TabIndex        =   125
         Text            =   "TB5"
         Top             =   3120
         Width           =   2000
      End
      Begin VB.TextBox TB7 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   9390
         Locked          =   -1  'True
         TabIndex        =   124
         Text            =   "TB7"
         Top             =   3840
         Width           =   2000
      End
      Begin VB.TextBox TB10 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   13305
         Locked          =   -1  'True
         TabIndex        =   123
         Text            =   "TB10"
         Top             =   2400
         Width           =   2000
      End
      Begin VB.TextBox TB9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   13305
         TabIndex        =   122
         Text            =   "TB9"
         Top             =   2040
         Width           =   2000
      End
      Begin VB.CheckBox CB14 
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
         Left            =   7320
         TabIndex        =   121
         Top             =   4920
         Width           =   200
      End
      Begin VB.CheckBox CB15 
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
         Left            =   7320
         TabIndex        =   120
         Top             =   6240
         Width           =   200
      End
      Begin VB.ComboBox CBB1 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   8835
         Style           =   2  'Dropdown List
         TabIndex        =   119
         Top             =   9000
         Width           =   4365
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   8835
         TabIndex        =   132
         Top             =   8640
         Width           =   4365
         _ExtentX        =   7699
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
         Format          =   416677888
         CurrentDate     =   41561
      End
      Begin VB.Label L40_Text 
         Caption         =   "L40_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5160
         TabIndex        =   169
         Top             =   7080
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "Maklumat Pembeli                                                                Maklumat Produk"
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
         Left            =   360
         TabIndex        =   168
         Top             =   360
         Width           =   15495
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Upah                 RM"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   7440
         TabIndex        =   167
         Top             =   3525
         Width           =   2265
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Scanner Mode"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   645
         TabIndex        =   166
         Top             =   105
         Width           =   2385
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Siri Produk      :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   7440
         TabIndex        =   165
         Top             =   1125
         Width           =   1905
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Siri Produk     "
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   7440
         TabIndex        =   164
         Top             =   2055
         Width           =   2265
      End
      Begin VB.Shape Shape2 
         Height          =   900
         Left            =   7200
         Top             =   720
         Width           =   8055
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Berat Asal               g"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   7440
         TabIndex        =   163
         Top             =   2430
         Width           =   2265
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Berat Jualan            g"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   7440
         TabIndex        =   162
         Top             =   2790
         Width           =   2265
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Semasa   RM/g"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   7440
         TabIndex        =   161
         Top             =   3165
         Width           =   2265
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Asal         RM"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   7440
         TabIndex        =   160
         Top             =   3855
         Width           =   2265
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Jualan      RM"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   11520
         TabIndex        =   159
         Top             =   2430
         Width           =   2265
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Adjustment        RM"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   11520
         TabIndex        =   158
         Top             =   2085
         Width           =   2265
      End
      Begin VB.Label L10_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L10_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9465
         TabIndex        =   157
         Top             =   1770
         Width           =   5835
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Kategori Produk       :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7440
         TabIndex        =   156
         Top             =   1770
         Width           =   2385
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Pilihan Ansuran"
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
         Left            =   7440
         TabIndex        =   155
         Top             =   4440
         Width           =   4545
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Pengiraan bayaran mengikut harga emas semasa."
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   7680
         TabIndex        =   154
         Top             =   4875
         Width           =   6105
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm87.frx":2410
         ForeColor       =   &H00000000&
         Height          =   765
         Left            =   7800
         TabIndex        =   153
         Top             =   5280
         Width           =   6345
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Harga bagi item atau emas ini adalah tetap mengikut tetapan harga yang dibuat semasa pendaftaran dilakukan."
         ForeColor       =   &H00000000&
         Height          =   765
         Left            =   7800
         TabIndex        =   152
         Top             =   6840
         Width           =   6345
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Pengiraan bayaran mengikut harga semasa pendaftaran ansuran dilakukan"
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   7680
         TabIndex        =   151
         Top             =   6195
         Width           =   6105
      End
      Begin VB.Shape Shape7 
         Height          =   4215
         Left            =   7200
         Top             =   4320
         Width           =   7215
      End
      Begin VB.Label L11_Text 
         Caption         =   "L11_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   150
         Top             =   9120
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label L12_Text 
         Caption         =   "L12_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         TabIndex        =   149
         Top             =   9120
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh *"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   7440
         TabIndex        =   148
         Top             =   8640
         Width           =   2385
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pekerja*"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   7440
         TabIndex        =   147
         Top             =   9000
         Width           =   2295
      End
      Begin VB.Label L13_Text 
         Caption         =   "L13_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   146
         Top             =   9120
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label95 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm87.frx":24B3
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
         Height          =   1005
         Left            =   7440
         TabIndex        =   145
         Top             =   7440
         Width           =   6825
      End
      Begin VB.Label L36_Text 
         Caption         =   "L36_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3360
         TabIndex        =   144
         Top             =   9120
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label L37_Text 
         Caption         =   "L37_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4440
         TabIndex        =   143
         Top             =   9120
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label L38_Text 
         Caption         =   "L38_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5520
         TabIndex        =   142
         Top             =   9120
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama          :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   141
         Top             =   2280
         Width           =   1665
      End
      Begin VB.Label L5_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L5_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         TabIndex        =   140
         Top             =   2280
         Width           =   5385
      End
      Begin VB.Shape Shape1 
         Height          =   1935
         Left            =   120
         Top             =   720
         Width           =   6975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm87.frx":258B
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
         Height          =   1005
         Left            =   2400
         TabIndex        =   139
         Top             =   1200
         Width           =   5010
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Klik di sini bagi memasukkan maklumat pembeli."
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
         Left            =   2400
         TabIndex        =   138
         Top             =   840
         Width           =   5010
      End
      Begin VB.Label Label75 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama          :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   137
         Top             =   4320
         Width           =   1665
      End
      Begin VB.Label L6_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L6_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         TabIndex        =   136
         Top             =   4320
         Width           =   4665
      End
      Begin VB.Shape Shape5 
         Height          =   1935
         Left            =   120
         Top             =   2760
         Width           =   6975
      End
      Begin VB.Label Label71 
         BackStyle       =   0  'Transparent
         Caption         =   "Klik di sini bagi memasukkan maklumat pembeli."
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
         Left            =   2400
         TabIndex        =   135
         Top             =   2880
         Width           =   5010
      End
      Begin VB.Label Label73 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm87.frx":266B
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
         Height          =   1245
         Left            =   2400
         TabIndex        =   134
         Top             =   3120
         Width           =   4650
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm87.frx":274E
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
         Height          =   1725
         Left            =   360
         TabIndex        =   133
         Top             =   4800
         Width           =   6570
      End
   End
   Begin VB.Label L15_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai Ansuran"
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
      Left            =   2640
      MouseIcon       =   "Frm87.frx":2924
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label L3_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pendaftaran"
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
      Left            =   120
      MouseIcon       =   "Frm87.frx":2C2E
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   0
      Width           =   2295
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
      Left            =   20400
      TabIndex        =   1
      Top             =   480
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
      Left            =   20400
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Menu Frm87_PM_Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm87_SM_Edit 
         Caption         =   "Lihat Data / Edit Data"
      End
      Begin VB.Menu Frm87_SM_Update 
         Caption         =   "Update Bayaran Ansuran"
      End
      Begin VB.Menu Frm87_SM_Rekod 
         Caption         =   "Rekod Bayaran Pelanggan Ini"
      End
      Begin VB.Menu Frm87_SM_Padam2 
         Caption         =   "Padam Data"
      End
      Begin VB.Menu Frm87_SM_Exccel 
         Caption         =   "Export Excel Report"
      End
   End
   Begin VB.Menu Frm87_PM_Menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm87_SM_Edit2 
         Caption         =   "Lihat Data / Edit Data"
      End
      Begin VB.Menu Frm87_SM_Padam 
         Caption         =   "Padam Resit"
      End
      Begin VB.Menu Frm87_LM_resit_ansuran 
         Caption         =   "Cetak Resit Bayaran Ansuran"
      End
   End
End
Attribute VB_Name = "Frm87"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB14_Click()
'on error resume next
If Frm87.CB14 = 1 Then
    Frm87.CB15 = 0
End If
End Sub
Private Sub CB15_Click()
'on error resume next
If Frm87.CB15 = 1 Then
    Frm87.CB14 = 0
End If
End Sub
Private Sub CB18_Click()
'On Error Resume Next
'on error resume next
Dim Frm87_LM_ANSURAN As Double
Dim Frm87_LM_UPAH As Double
Dim Frm87_LM_GST_ANSURAN As Double
Dim Frm87_LM_GST_UPAH As Double

Frm87_LM_ANSURAN = 0
Frm87_LM_UPAH = 0
Frm87_LM_GST_ANSURAN = 0
Frm87_LM_GST_UPAH = 0

If Frm87.TB12 <> vbNullString And IsNumeric(Frm87.TB12) Then
    Frm87_LM_ANSURAN = Frm87.TB12 'Jumlah Ansuran
End If
If Frm87.TB19 <> vbNullString And IsNumeric(Frm87.TB19) Then
    Frm87_LM_UPAH = Frm87.TB19 'Jumlah Upah
End If
If Frm87.L21_Text <> vbNullString And IsNumeric(Frm87.L21_Text) Then
    Frm87_LM_GST_ANSURAN = Frm87.L21_Text 'Jumlah GST Ansuran
End If
If Frm87.L22_Text <> vbNullString And IsNumeric(Frm87.L22_Text) Then
    Frm87_LM_GST_UPAH = Frm87.L22_Text 'Jumlah GST Upah
End If
    
If Frm87.CB18 = 1 Then
    Frm87.CB19 = 0
    Frm87.L22_Text = "0.00" 'Jumlah Cukai GST (RM)
    
    If Frm87.CB22 = 1 Then
        Frm87.L25_Text = Format(Frm87_LM_ANSURAN + Frm87_LM_UPAH, "#,##0.00")
        Frm87.L26_Text = Format(Frm87_LM_GST_ANSURAN + Frm87_LM_GST_UPAH, "#,##0.00")
    Else
        Frm87.L25_Text = Format(Frm87_LM_UPAH, "#,##0.00")
        Frm87.L26_Text = Format(Frm87_LM_GST_ANSURAN + Frm87_LM_GST_UPAH, "#,##0.00")
    End If
Else
    If Frm87.CB22 = 1 Then
        Frm87.L25_Text = Format(Frm87_LM_ANSURAN, "#,##0.00")
        Frm87.L26_Text = Format(Frm87_LM_GST_ANSURAN + Frm87_LM_GST_UPAH, "#,##0.00")
    Else
        Frm87.L25_Text = Format(0, "0.00")
        Frm87.L26_Text = Format(Frm87_LM_GST_ANSURAN + Frm87_LM_GST_UPAH, "#,##0.00")
    End If
End If

Call Frm87_LM_Detail_GST
End Sub
Private Sub CB19_Click()
'On Error Resume Next
Dim Frm87_LM_KADAR_GST As Double
Dim Frm87_LM_BAYARAN As Double

If Frm87.CB19 = 1 Then
    Frm87.CB18 = 0
End If

Frm87_LM_KADAR_GST = 0
Frm87_LM_BAYARAN = 0

If GLOBAL_DISABLE = 0 Then
    If Frm87.CB19 = 1 And (Frm87.TB19 <> vbNullString And IsNumeric(Frm87.TB19)) And (Frm87.L17_Text <> vbNullString And IsNumeric(Frm87.L17_Text)) Then
        Frm87_LM_KADAR_GST = Frm87.L17_Text 'Jumlah Kadar GST (%)
        Frm87_LM_BAYARAN = Frm87.TB19 'Jumlah (RM)
        
        Frm87.L22_Text = Format((Frm87_LM_KADAR_GST / 100) * Frm87_LM_BAYARAN, "#,##0.00") 'Jumlah Cukai GST (RM)
    Else
        Frm87.L22_Text = "0.00" 'Jumlah Cukai GST (RM)
    End If
End If

Call Frm87_LM_Detail_GST
End Sub

Private Sub CB20_Click()
'on error resume next
If Frm87.CB20 = 1 Then
    Frm87.TB12 = "0.00"
    Frm87.TB12.Locked = False
    Frm87.TB12.BackColor = &HFFFFFF
Else
    Frm87.TB12 = "0.00"
    Frm87.TB12.Locked = True
    Frm87.TB12.BackColor = &H8000000A
    
    Frm87.CB22 = 1
    Frm87.CB23 = 0
End If
End Sub
Private Sub CB21_Click()
'on error resume next
If Frm87.CB21 = 1 Then
    Frm87.TB19 = "0.00"
    Frm87.TB19.Locked = False
    Frm87.TB19.BackColor = &HFFFFFF
Else
    Frm87.TB19 = "0.00"
    Frm87.TB19.Locked = True
    Frm87.TB19.BackColor = &H8000000A
    
    Frm87.CB18 = 1
    Frm87.CB19 = 0
End If
End Sub
Private Sub CB22_Click()
'on error resume next
If Frm87.CB22 = 1 Then
    Frm87.CB23 = 0
    Frm87.L21_Text = "0.00" 'Jumlah Cukai GST (RM)
    
    If IsNumeric(Frm87.TB20) Then
        Frm87.TB42 = Format(Frm87.TB20, "#,##0.00")
    Else
        Frm87.TB42 = Format(0, "#,##0.00")
    End If
End If

Call Frm87_LM_Detail_GST
End Sub
Private Sub CB23_Click()
'on error resume next
Dim Frm87_LM_KADAR_GST As Double
Dim Frm87_LM_ANSURAN As Double

If Frm87.CB23 = 1 Then
    Frm87.CB22 = 0
End If
If Frm87.CB23 = 0 Then
    Frm87.CB27 = 0
End If

Frm87_LM_KADAR_GST = 0
Frm87_LM_ANSURAN = 0

If GLOBAL_DISABLE = 0 Then
    
    If Frm87.CB27 = 0 Then
        
        If Frm87.CB23 = 1 And (Frm87.TB20 <> vbNullString And IsNumeric(Frm87.TB20)) And (Frm87.L17_Text <> vbNullString And IsNumeric(Frm87.L17_Text)) Then
        
            If IsNumeric(Frm87.L17_Text) Then Frm87_LM_KADAR_GST = Frm87.L17_Text 'Jumlah Kadar GST (%)
            If IsNumeric(Frm87.TB20) Then Frm87_LM_ANSURAN = Frm87.TB20 'Jumlah Ansuran (RM)
            
            Frm87.TB13 = Format((Frm87_LM_KADAR_GST / 100) * Frm87_LM_ANSURAN, "#,##0.00") 'Jumlah Cukai GST (RM)
            Frm87.L21_Text = Format((Frm87_LM_KADAR_GST / 100) * Frm87_LM_ANSURAN, "#,##0.00") 'Jumlah Cukai GST (RM)
            Frm87.TB42 = Format(Frm87_LM_ANSURAN, "#,##0.00") 'Jumlah Harga Tanpa GST (RM)
            
        Else
        
            Frm87.TB13 = "0.00" 'Jumlah Cukai GST (RM)
            Frm87.L21_Text = "0.00" 'Jumlah Cukai GST (RM)
            
            If IsNumeric(Frm87.TB20) Then
                Frm87.TB42 = Format(Frm87.TB20, "#,##0.00")
            Else
                Frm87.TB42 = Format(0, "#,##0.00")
            End If
        
        End If
            
    ElseIf Frm87.CB27 = 1 Then
        
        If Frm87.CB23 = 1 And (Frm87.TB20 <> vbNullString And IsNumeric(Frm87.TB20)) And (Frm87.L17_Text <> vbNullString And IsNumeric(Frm87.L17_Text)) Then
        
            If IsNumeric(Frm87.L17_Text) Then Frm87_LM_KADAR_GST = Frm87.L17_Text 'Jumlah Kadar GST (%)
            If IsNumeric(Frm87.TB20) Then Frm87_LM_ANSURAN = Frm87.TB20 'Jumlah Ansuran (RM)
            
            Frm87.TB13 = Format(Frm87_LM_ANSURAN - (Frm87_LM_ANSURAN / (1 + Frm87_LM_KADAR_GST / 100)), "#,##0.00") 'Jumlah Cukai GST (RM)
            Frm87.L21_Text = Format(Frm87_LM_ANSURAN - (Frm87_LM_ANSURAN / (1 + Frm87_LM_KADAR_GST / 100)), "#,##0.00") 'Jumlah Cukai GST (RM)
            Frm87.TB42 = Format(Frm87_LM_ANSURAN / (1 + (Frm87_LM_KADAR_GST / 100)), "#,##0.00") 'Jumlah Harga Tanpa GST (RM)
            
        Else
        
            Frm87.TB13 = "0.00" 'Jumlah Cukai GST (RM)
            Frm87.L21_Text = "0.00" 'Jumlah Cukai GST (RM)
            
            If IsNumeric(Frm87.TB20) Then
                Frm87.TB42 = Format(Frm87.TB20, "#,##0.00")
            Else
                Frm87.TB42 = Format(0, "#,##0.00")
            End If
            
        End If
        
    End If
    
End If

Call Frm87_LM_Detail_GST
End Sub
Private Sub CB24_Click()
'on error resume next
If Frm87.CB24 = 1 Then
    Frm87.CB25 = 0
    Frm87.CB26 = 0
End If
End Sub
Private Sub CB25_Click()
'on error resume next
If Frm87.CB25 = 1 Then
    Frm87.CB24 = 0
    Frm87.CB26 = 0
End If
End Sub
Private Sub CB26_Click()
'on error resume next
If Frm87.CB26 = 1 Then
    Frm87.CB25 = 0
    Frm87.CB24 = 0
End If
End Sub
Private Sub CB27_Click()
'on error resume next
Dim Frm87_LM_KADAR_GST As Double
Dim Frm87_LM_ANSURAN As Double

If Frm87.CB23 = 1 Then
    Frm87.CB22 = 0
End If
If Frm87.CB23 = 0 Then
    Frm87.CB27 = 0
End If

Frm87_LM_KADAR_GST = 0
Frm87_LM_ANSURAN = 0

If GLOBAL_DISABLE = 0 Then
    If Frm87.CB27 = 0 Then
        If Frm87.CB23 = 1 And (Frm87.TB20 <> vbNullString And IsNumeric(Frm87.TB20)) And (Frm87.L17_Text <> vbNullString And IsNumeric(Frm87.L17_Text)) Then
        
            If IsNumeric(Frm87.L17_Text) Then Frm87_LM_KADAR_GST = Frm87.L17_Text 'Jumlah Kadar GST (%)
            If IsNumeric(Frm87.TB20) Then Frm87_LM_ANSURAN = Frm87.TB20 'Jumlah Ansuran (RM)
            
            Frm87.TB13 = Format((Frm87_LM_KADAR_GST / 100) * Frm87_LM_ANSURAN, "#,##0.00") 'Jumlah Cukai GST (RM)
            Frm87.L21_Text = Format((Frm87_LM_KADAR_GST / 100) * Frm87_LM_ANSURAN, "#,##0.00") 'Jumlah Cukai GST (RM)
            Frm87.TB42 = Format(Frm87_LM_ANSURAN, "#,##0.00") 'Jumlah Harga Tanpa GST (RM)
            
        Else
        
            Frm87.TB13 = "0.00" 'Jumlah Cukai GST (RM)
            Frm87.L21_Text = "0.00" 'Jumlah Cukai GST (RM)
            If IsNumeric(Frm87.TB20) Then
                Frm92.TB42 = Format(Frm87.TB20, "#,##0.00")
            Else
                Frm92.TB42 = Format(0, "#,##0.00")
            End If
        
        
        End If
            
    ElseIf Frm87.CB27 = 1 Then
        
        If Frm87.CB23 = 1 And (Frm87.TB20 <> vbNullString And IsNumeric(Frm87.TB20)) And (Frm87.L17_Text <> vbNullString And IsNumeric(Frm87.L17_Text)) Then
        
            If IsNumeric(Frm87.L17_Text) Then Frm87_LM_KADAR_GST = Frm87.L17_Text 'Jumlah Kadar GST (%)
            If IsNumeric(Frm87.TB20) Then Frm87_LM_ANSURAN = Frm87.TB20 'Jumlah Ansuran (RM)
            
            Frm87.TB13 = Format(Frm87_LM_ANSURAN - (Frm87_LM_ANSURAN / (1 + Frm87_LM_KADAR_GST / 100)), "#,##0.00") 'Jumlah Cukai GST (RM)
            Frm87.L21_Text = Format(Frm87_LM_ANSURAN - (Frm87_LM_ANSURAN / (1 + Frm87_LM_KADAR_GST / 100)), "#,##0.00") 'Jumlah Cukai GST (RM)
            Frm87.TB42 = Format(Frm87_LM_ANSURAN / (1 + (Frm87_LM_KADAR_GST / 100)), "#,##0.00") 'Jumlah Harga Tanpa GST (RM)
            
        End If
    Else
    
        Frm87.TB13 = "0.00" 'Jumlah Cukai GST (RM)
        Frm87.L21_Text = "0.00" 'Jumlah Cukai GST (RM)
        
        If IsNumeric(Frm87.TB20) Then
            Frm92.TB42 = Format(Frm87.TB20, "#,##0.00")
        Else
            Frm92.TB42 = Format(0, "#,##0.00")
        End If
        
    End If
End If

Call Frm87_LM_Detail_GST
End Sub



Private Sub CMD1_Click()
'On Error Resume Next
If Frm87.L5_Text = vbNullString Then
    
    If Frm87.L6_Text <> vbNullString Then
    
        Note = "Anda telah mengisi data pembeli ini di dalam ruangan pelanggan yang berdaftar dengan kedai." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Jika ada meneruskan menu ini , semua data pembeli di dalam ruangan pelanggan berdaftar akan dipadamkan." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
        
            Unload Frm27
            Unload Frm28
            Call Frm26_initial
            
            Frm87.L6_Text = vbNullString 'Nama pembeli : Berdaftar
            
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
Private Sub CMD10_Click()
'on error resume next
Note = "Batal Urusan Ini ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    Frm87.Pic4.Visible = True
    Frm87.Pic5.Visible = False
End If
End Sub
Private Sub CMD11_Click()
'on error resume next
Dim Err(30)
DATA_SAVE = 0

If Frm87.CB24 = 0 And Frm87.CB25 = 0 And Frm87.CB26 = 0 Then
    x = x + 1
    Err(x) = "Sila Buat Pilihan Krateria Carian."
End If
If Frm87.TB41 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan Maklumat Carian."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    If Frm87.CB24 = 1 Then
        Call Frm87_Senarai_Ansuran_Header
        Call Frm87_Carian_Ansuran2
    End If
    
    If Frm87.CB25 = 1 Or Frm87.CB26 = 1 Then
        Call Frm87_Senarai_Ansuran_Header
        Call Frm87_Carian_Ansuran
    End If
End If
End Sub
Private Sub CMD12_Click()
'on error resume next
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Senarai Ini. Teruskan ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    Call Frm87_Senarai_Ansuran_Header
    Call Frm87_Senarai_Ansuran
End If
End Sub
Private Sub CMD13_Click()
'On Error Resume Next
Dim Err(30)
Dim Frm87_LM_BERAT_BAYARAN As Double
Dim Frm87_BAKI_BERAT As Double
Dim Frm87_LM_UPAH As Double
Dim Frm87_LM_BAKI_UPAH As Double
Dim Frm87_LM_HARGA As Double
Dim Frm87_LM_JUMLAH_BAYARAN As Double
Dim Frm87_LM_JUMLAH_SIMPANAN As Double
Dim Frm87_LM_GUNA_SIMPAN As Double
Dim Frm87_LM_JUMLAH_BAYARAN_ASAL  As Double
Dim Frm87_LM_JUMLAH_BERAT_JUALAN_ASAL As Double
Dim Frm87_LM_JUMLAH_BERAT_ASAL As Double
Dim Frm87_LM_UPAH_ASAL As Double
Dim Frm87_LM_UPAH_JUALAN_ASAL As Double
Dim Frm87_LM_HARGA_JUALAN_ASAL As Double
Dim Frm87_LM_BERAT_POTONG As Double
Dim aaa As Double
Dim bbb As Double

DATA_SAVE = 0
Frm87_LM_BERAT_BAYARAN = 0
Frm87_BAKI_BERAT = 0
Frm87_LM_UPAH = 0
Frm87_LM_BAKI_UPAH = 0
Frm87_LM_HARGA = 0
Frm87_LM_JUMLAH_BAYARAN = 0
Frm87_LM_JUMLAH_SIMPANAN = 0  'Jumlah Simpanan Yang Ada
Frm87_LM_GUNA_SIMPAN = 0  'Jumlah Simpanan Yang Hendak Digunakan
Frm87_LM_BERAT_POTONG = 0

Frm87_LM_JUMLAH_BAYARAN_ASAL = 0
Frm87_LM_JUMLAH_BERAT_JUALAN_ASAL = 0
Frm87_LM_JUMLAH_BERAT_ASAL = 0
Frm87_LM_UPAH_ASAL = 0
Frm87_LM_UPAH_JUALAN_ASAL = 0
Frm87_LM_HARGA_JUALAN_ASAL = 0

aaa = 0
bbb = 0

If Frm87.CB20 = 0 And Frm87.CB21 = 0 Then
    x = x + 1
    Err(x) = "Sila Buat Pilihan Jenis Bayaran."
End If
If Frm87.CB20 = 1 Then
    If Frm87.TB12 = vbNullString Or (Frm87.TB12 <> vbNullString And Not IsNumeric(Frm87.TB12)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Bayaran Ansuran]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm87.CB22 = 0 And Frm87.CB23 = 0 Then
        x = x + 1
        Err(x) = "Sila Buat Pilihan Jenis GST Bagi Bayaran Ansuran."
    End If
    If Frm87.Pic6.Visible = True Then
        If Frm87.TB15 = vbNullString Or (Frm87.TB15 <> vbNullString And Not IsNumeric(Frm87.TB15)) Then
            x = x + 1
            Err(x) = "Sila Masukkan [Harga Emas Semasa]. Hanya NOMBOR dibenarkan dalam ruangan ini."
        End If
    Else
        If (Frm87.TB12 <> vbNullString And IsNumeric(Frm87.TB12)) And (Frm87.L28_Text <> vbNullString And IsNumeric(Frm87.L28_Text)) Then
            aaa = Frm87.TB12
            bbb = Frm87.L28_Text
            
            If aaa > bbb Then
                x = x + 1
                Err(x) = "Jumlah Bayaran Ansuran Melebihi Jumlah Baki."
            End If
        End If
    End If
    If (Frm87.TB16 <> vbNullString And IsNumeric(Frm87.TB16)) And (Frm87.L19_Text <> vbNullString And IsNumeric(Frm87.L19_Text)) Then
        Frm87_LM_BERAT_BAYARAN = Frm87.TB16 'Berat Bayaran Kali Ini
        Frm87_BAKI_BERAT = Frm87.L19_Text 'Baki Berat
        
        If Frm87_LM_BERAT_BAYARAN > Frm87_BAKI_BERAT Then
            x = x + 1
            Err(x) = "Berat Bagi Bayaran Adalah Melebihi Baki Berat Yang Tinggal."
        End If
    End If
End If

If Frm87.CB21 = 1 Then
    If Frm87.TB19 = vbNullString Or (Frm87.TB19 <> vbNullString And Not IsNumeric(Frm87.TB19)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Jumlah Upah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    'If Frm87.CB18 = 0 And Frm87.CB19 = 0 Then
    '    X = X + 1
    '    Err(X) = "Sila Buat Pilihan Jenis GST Bagi Bayaran Upah."
    'End If
    If (Frm87.TB19 <> vbNullString And IsNumeric(Frm87.TB19)) And (Frm87.L20_Text <> vbNullString And IsNumeric(Frm87.L20_Text)) Then
        Frm87_LM_UPAH = Frm87.TB19 'Bayaran Upah
        Frm87_LM_BAKI_UPAH = Frm87.L20_Text 'Baki Upah
        
        If Frm87_LM_UPAH > Frm87_LM_BAKI_UPAH Then
            x = x + 1
            Err(x) = "Upah Bagi Bayaran Adalah Melebihi Baki Upah Yang Tinggal."
        End If
    End If
End If
If Frm87.TB17 = vbNullString Or (Frm87.TB17 <> vbNullString And Not IsNumeric(Frm87.TB17)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Adjustment]. Hanya NOMBOR dibenarkan dalam ruangan ini. Masukkan [0.00] Jika Tiada Adjustment."
End If
If Frm87.TB27 = vbNullString Or (Frm87.TB27 <> vbNullString And Not IsNumeric(Frm87.TB27)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR Dibenarkan Dalam Ruangan Bayaran Secara Tunai. Sila Masukkan 0 Jika Tiada Bayaran Tunai."
End If
If Frm87.TB28 = vbNullString Or (Frm87.TB28 <> vbNullString And Not IsNumeric(Frm87.TB28)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR Dibenarkan Dalam Ruangan Bayaran Secara Bank In. Sila Masukkan 0 Jika Tiada Bayaran Bank In."
End If
If Frm87.TB29 = vbNullString Or (Frm87.TB29 <> vbNullString And Not IsNumeric(Frm87.TB29)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR Dibenarkan Dalam Ruangan Bayaran Secara Kad Kredit. Sila Masukkan 0 Jika Tiada Bayaran Kad Kredit."
End If
If Frm87.TB21 = vbNullString Or (Frm87.TB21 <> vbNullString And Not IsNumeric(Frm87.TB21)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR Dibenarkan Dalam Ruangan Bayaran Secara Duit Simpanan Di Kedai. Sila Masukkan 0 Jika Tiada Bayaran Simpanan Di Kedai."
End If
If Frm87.TB38 = vbNullString Or (Frm87.TB38 <> vbNullString And Not IsNumeric(Frm87.TB38)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR Dibenarkan Dalam Ruangan Bayaran Secara Debit Kad. Sila Masukkan 0 Jika Tiada Bayaran Debit Kad."
End If

If (Frm87.TB21 <> vbNullString And IsNumeric(Frm87.TB21)) And (Frm87.L27_Text <> vbNullString And IsNumeric(Frm87.L27_Text)) Then
    Frm87_LM_JUMLAH_SIMPANAN = Frm87.L27_Text  'Jumlah Simpanan Yang Ada
    Frm87_LM_GUNA_SIMPAN = Frm87.TB21  'Jumlah Simpanan Yang Hendak Digunakan
    
    If Frm87_LM_GUNA_SIMPAN > Frm87_LM_JUMLAH_SIMPANAN Then
        x = x + 1
        Err(x) = "Jumlah Simpanan Yang Hendak Digunakan Melebihi Simpanan Yang Ada."
    End If
End If

If (Frm87.TB32 <> vbNullString And IsNumeric(Frm87.TB32)) And (Frm87.TB18 <> vbNullString And IsNumeric(Frm87.TB18)) Then
    Frm87_LM_JUMLAH_BAYARAN = Frm87.TB32 'Jumlah Bayaran
    Frm87_LM_HARGA = Frm87.TB18 'Harga Keseluruhan
    
    If Frm87_LM_JUMLAH_BAYARAN <> Frm87_LM_HARGA Then
        x = x + 1
        Err(x) = "Jumlah Bayaran Tidak Sama Dengan Jumlah Harga Barang."
    End If
End If
If Frm87.CBB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih Nama Pekerja"
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
    
'### Padamkan Maklumat Resit Yang Lama### - Start
        Call padam_rekod_ansuran
'### Padamkan Maklumat Resit Yang Lama### - End
    
'### Carian Jenis Ansuran ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 27_senarai_ansuran where ID='" & Frm87.L18_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!no_siri_Produk) Then Frm87_LM_No_SIRI = rs!no_siri_Produk 'No. Siri Produk
            If Not IsNull(rs!jenis_produk) Then
                If rs!jenis_produk = 1 Then
                    Frm87_LM_JENIS_PRODUK = 1
                ElseIf rs!jenis_produk = 0 Then
                    Frm87_LM_JENIS_PRODUK = 0
                    
                    Frm87_LM_POTONG = 0
                    If rs!Berat_Asal <> rs!berat_jualan Then
                        Frm87_LM_BERAT_POTONG = rs!berat_jualan
                        Frm87_LM_POTONG = 1
                    End If
                End If
            End If
            If Not IsNull(rs!jenis_ansuran) Then
                If rs!jenis_ansuran = 0 Then '0 : Harga Semasa , 1 : Harga Tetap
                    Frm87_LM_JENIS = 0
                    
                    If Not IsNull(rs!jumlah_bayaran) Then
                        Frm87_LM_JUMLAH_BAYARAN_ASAL = rs!jumlah_bayaran 'Jumlah Bayaran Terkumpul - Asal (RM)
                        
                        rs!jumlah_bayaran = Format(Frm87_LM_JUMLAH_BAYARAN_ASAL + Frm87.TB20, "0.00")
                    End If
                    
                    If Frm87.CB20 = 1 Then
                        If Not IsNull(rs!berat_jualan) Then Frm87_LM_JUMLAH_BERAT_JUALAN_ASAL = rs!berat_jualan 'Jumlah Berat Jualan - Asal (RM)
                        
                        If Not IsNull(rs!JUMLAH_BERAT) Then
                            Frm87_LM_JUMLAH_BERAT_ASAL = rs!JUMLAH_BERAT 'Jumlah Berat Yang Telah Dijelaskan - Asal (RM)
                        
                            rs!JUMLAH_BERAT = Format(Frm87.TB16 + Frm87_LM_JUMLAH_BERAT_ASAL, "0.00") 'Jumlah Berat Yang Telah Dijelaskan (g)
                        End If
                    
                        rs!BAKI_BERAT = Format(Frm87_LM_JUMLAH_BERAT_JUALAN_ASAL - rs!JUMLAH_BERAT, "0.00") 'Baki Berat
                    End If
                    
                    If Frm87.CB21 = 1 Then
                        If Not IsNull(rs!JUMLAH_UPAH) Then
                            Frm87_LM_UPAH_ASAL = rs!JUMLAH_UPAH 'Jumlah Upah Yang Telah Dijelaskan - Asal (RM)
                            
                            rs!JUMLAH_UPAH = Format(Frm87.TB19 + Frm87_LM_UPAH_ASAL, "0.00")
                        End If
                        If Not IsNull(rs!UPAH) Then Frm87_LM_UPAH_JUALAN_ASAL = rs!UPAH 'Jumlah Tetapan Upah - Asal (RM)
                        
                        rs!baki_upah = Format(Frm87_LM_UPAH_JUALAN_ASAL - rs!JUMLAH_UPAH, "0.00") 'Baki Upah
                    End If
                ElseIf rs!jenis_ansuran = 1 Then
                    Frm87_LM_JENIS = 1
                    If Not IsNull(rs!jumlah_bayaran) Then Frm87_LM_JUMLAH_BAYARAN_ASAL = rs!jumlah_bayaran 'Jumlah Bayaran Terkumpul - Asal (RM)
                    If Not IsNull(rs!harga_jualan) Then Frm87_LM_HARGA_JUALAN_ASAL = rs!harga_jualan 'Jumlah Harga Jualan - Asal (RM)
                    
                    aaa = Frm87.TB12
                    rs!jumlah_bayaran = Format(Frm87_LM_JUMLAH_BAYARAN_ASAL + aaa, "0.00") 'Jumlah Terkumpul Yang Baru
                    rs!baki_bayaran = Format(Frm87_LM_HARGA_JUALAN_ASAL - (Frm87_LM_JUMLAH_BAYARAN_ASAL + aaa), "0.00") 'Baki Bayaran
                End If
            End If
            If Not IsNull(no_rujukan_pelanggan) Then Frm87_LM_No_PEMBELI = rs!no_rujukan_pelanggan 'No. Rujukan Pembeli
            If Not IsNull(rs!kategori_pembeli) Then Frm87_LM_KATEGORI = rs!kategori_pembeli 'Kategori Pembeli
            
            rs.Update
        End If
        
        rs.Close
        Set rs = Nothing
'### Carian Jenis Ansuran ### - End
    
'### Update Rekod Ansuran ### - Start
        Frm87_LM_No_RESIT_ANSURAN = Frm87.L12_Text 'No. Resit Ansuran
        
'###Update Bayaran Ansuran### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 28_rekod_ansuran where no_resit_ansuran='" & Frm87_LM_No_RESIT_ANSURAN & "'", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
            rs.AddNew
            If Frm87.L18_Text <> vbNullString Then
                rs!id_database_reg = Frm87.L18_Text 'No. ID Dari Database Senarai Pembeli Ansuran
            Else
                rs!id_database_reg = Null
            End If
            If Frm87.L12_Text <> vbNullString Then
                rs!no_resit_ansuran = Frm87_LM_No_RESIT_ANSURAN 'No. Resit Ansuran
            Else
                rs!no_resit_ansuran = Null
            End If
            If Frm87.CB20 = 1 Then
                rs!FLAG_ANSURAN = 1 'Flag samada ada bayaran ansuran atau tidak , 0 : Tiada bayaran ansuran , 1 : Ada bayaran ansuran
                If Frm87.TB12 <> vbNullString Then
                    rs!jumlah_ansuran = Format(Frm87.TB12, "0.00") 'Jumlah Bayran Ansuran
                Else
                    rs!jumlah_ansuran = Null 'Jumlah Bayran Ansuran
                End If
'                If Frm87.CB22 = 1 Then
'                    rs!flag_ansuran_zr = 1 'Flag samada ada zr bagi bayaran ansuran , 0 : Tiada ZR , 1 : Ada ZR
'                Else
'                    rs!flag_ansuran_zr = 0 'Flag samada ada zr bagi bayaran ansuran , 0 : Tiada ZR , 1 : Ada ZR
'                End If
'                If Frm87.CB23 = 1 Then
'                    rs!flag_ansuran_sr = 1 'Flag samada ada sr bagi bayaran ansuran , 0 : Tiada SR , 1 : Ada SR
'                Else
'                    rs!flag_ansuran_sr = 0 'Flag samada ada sr bagi bayaran ansuran , 0 : Tiada SR , 1 : Ada SR
'                End If
'                If Frm87.L21_Text <> vbNullString Then
'                    rs!ansuran_gst = Format(Frm87.L21_Text, "0.00") 'Jumlah GST bagi bayaran ansuran (RM)
'                Else
'                    rs!ansuran_gst = Null
'                End If
                If Frm87.Pic6.Visible = True Then
                    If Frm87.TB15 <> vbNullString Then
                        rs!harga_Semasa = Format(Frm87.TB15, "0.00") 'Harga Semasa
                    Else
                        rs!harga_Semasa = "0.00"
                    End If
                    If Frm87.TB16 <> vbNullString Then
                        rs!berat_diperoleh = Format(Frm87.TB16, "0.00") 'Berat Diperolehi
                    Else
                        rs!berat_diperoleh = "0.00"
                    End If
                Else
                    rs!harga_Semasa = Null
                    rs!berat_diperoleh = Null
                End If
            Else
                rs!FLAG_ANSURAN = 0 'Flag samada ada bayaran ansuran atau tidak , 0 : Tiada bayaran ansuran , 1 : Ada bayaran ansuran
                rs!jumlah_ansuran = Null 'Jumlah Bayran Ansuran
'                rs!flag_ansuran_zr = Null 'Flag samada ada zr bagi bayaran ansuran , 0 : Tiada ZR , 1 : Ada ZR
'                rs!flag_ansuran_sr = Null 'Flag samada ada sr bagi bayaran ansuran , 0 : Tiada SR , 1 : Ada SR
'                rs!ansuran_gst = Null 'Jumlah GST bagi bayaran ansuran (RM)
                rs!harga_Semasa = Null
                rs!berat_diperoleh = Null
            End If
            
            If Frm87.CB22 = 1 Then
                rs!flag_ansuran_zr = 1 'Flag samada ada zr bagi bayaran ansuran , 0 : Tiada ZR , 1 : Ada ZR
            Else
                rs!flag_ansuran_zr = 0 'Flag samada ada zr bagi bayaran ansuran , 0 : Tiada ZR , 1 : Ada ZR
            End If
            If Frm87.CB23 = 1 Then
                rs!flag_ansuran_sr = 1 'Flag samada ada sr bagi bayaran ansuran , 0 : Tiada SR , 1 : Ada SR
            Else
                rs!flag_ansuran_sr = 0 'Flag samada ada sr bagi bayaran ansuran , 0 : Tiada SR , 1 : Ada SR
            End If
            If Frm87.L21_Text <> vbNullString Then
                rs!ansuran_gst = Format(Frm87.L21_Text, "0.00") 'Jumlah GST bagi bayaran ansuran (RM)
            Else
                rs!ansuran_gst = Null
            End If
            
            If Frm87.CB21 = 1 Then
                rs!flag_upah = 1 'Flag samada ada bayaran upah atau tidak , 0 : Tiada Bayaran Upah , 1 : Ada Bayaran Upah
                If Frm87.TB19 <> vbNullString Then
                    rs!JUMLAH_UPAH = Format(Frm87.TB19, "0.00") 'Jumlah Bayran Upah
                Else
                    rs!JUMLAH_UPAH = Null 'Jumlah Bayran Ansuran
                End If
'                If Frm87.CB18 = 1 Then
'                    rs!flag_upah_zr = 1 'Flag samada ada zr bagi bayaran upah , 0 : Tiada ZR , 1 : Ada ZR
'                Else
'                    rs!flag_upah_zr = 0 'Flag samada ada zr bagi bayaran upah , 0 : Tiada ZR , 1 : Ada ZR
'                End If
'                If Frm87.CB19 = 1 Then
'                    rs!flag_upah_sr = 1 'Flag samada ada sr bagi bayaran upah , 0 : Tiada SR , 1 : Ada SR
'                Else
'                    rs!flag_upah_sr = 0 'Flag samada ada sr bagi bayaran upah , 0 : Tiada SR , 1 : Ada SR
'                End If
'                If Frm87.L22_Text <> vbNullString Then
'                    rs!upah_gst = Format(Frm87.L22_Text, "0.00") 'Jumlah GST Bagi Upah (RM)
'                Else
'                    rs!upah_gst = Null
'                End If
            Else
                rs!flag_upah = 0 'Flag samada ada bayaran upah atau tidak , 0 : Tiada Bayaran Upah , 1 : Ada Bayaran Upah
                rs!JUMLAH_UPAH = Null 'Jumlah Bayran Ansuran
 '               rs!flag_upah_zr = Null 'Flag samada ada zr bagi bayaran upah , 0 : Tiada ZR , 1 : Ada ZR
 '               rs!flag_upah_sr = Null 'Flag samada ada sr bagi bayaran upah , 0 : Tiada SR , 1 : Ada SR
 '               rs!upah_gst = Null 'Jumlah GST Bagi Upah (RM)
            End If
            rs!jenis_ansuran = Frm87_LM_JENIS 'Jenis Ansuran , 0 : Harga Semasa , 1 : Harga Tetap
            If Frm87.TB20 <> vbNullString Then
                rs!jumlah_bayaran = Format(Frm87.TB20, "0.00") 'Jumlah Ansuran + Jumlah Upah
            Else
                rs!jumlah_bayaran = Null
            End If
            If Frm87.L17_Text <> vbNullString Then
                rs!kadar_gst = Frm87.L17_Text 'Kadar GST (%)
            Else
                rs!kadar_gst = Null
            End If
            If Frm87.TB13 <> vbNullString Then
                rs!jumlah_gst = Format(Frm87.TB13, "0.00") 'Jumlah GST
            Else
                rs!jumlah_gst = Null
            End If
            If Frm87.TB14 <> vbNullString Then
                rs!jumlah_asal = Format(Frm87.TB14, "0.00") 'Jumlah Asal (Ansuran + Upah) + GST
            Else
                rs!jumlah_asal = Null
            End If
            If Frm87.TB17 <> vbNullString Then
                rs!adjustment = Format(Frm87.TB17, "0.00") 'Adjustment
            Else
                rs!adjustment = Null
            End If
            If Frm87.TB18 <> vbNullString Then
                rs!jumlah_keseluruhan = Format(Frm87.TB18, "0.00") 'Jumlah bayaran selepas adjustment
            Else
                rs!jumlah_keseluruhan = Null
            End If
            rs!tarikh = Frm87.DTPicker2 'Tarikh Bayaran
            If Frm87.CBB2 <> vbNullString Then
                Frm87_LM_EMP_NO = Split(Frm87.CBB2, "  |  ")(1)
                rs!no_rujukan_pekerja = Frm87_LM_EMP_NO 'No. Pekerja
            End If
            rs!write_timestamp = Now
            rs.Update
            DATA_SAVE = 1
        End If
        
        rs.Close
        Set rs = Nothing
'###Update Bayaran Ansuran### - End
        
        If DATA_SAVE = 1 Then
        
'###Update Akaun Ansuran### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 29_akaun_ansuran", cn, adOpenKeyset, adLockOptimistic
            
            rs.AddNew
            If Frm87.L12_Text <> vbNullString Then
                rs!no_resit = Frm87_LM_No_RESIT_ANSURAN 'No. Resit Ansuran
            Else
                rs!no_resit = Null
            End If
            rs!tarikh = Frm87.DTPicker2 'Tarikh Bayaran
            If Frm87.TB27 <> vbNullString Then
                rs!tunai = Format(Frm87.TB27, "0.00") 'Cara Bayaran : Tunai
            Else
                rs!tunai = Null
            End If
            If Frm87.TB28 <> vbNullString Then
                rs!bank_in = Format(Frm87.TB28, "0.00") 'Cara Bayaran : Bank In
            Else
                rs!bank_in = Null
            End If
            If Frm87.TB29 <> vbNullString Then
                rs!kad_kredit = Format(Frm87.TB29, "0.00") 'Cara Bayaran : Kad Kredit
            Else
                rs!kad_kredit = Null
            End If
            If Frm87.L31_Text <> vbNullString Then
                rs!cas_Kad_Kredit = Frm87.L31_Text 'Cara Bayaran : Cas Kad Kredit (%)
            Else
                rs!cas_Kad_Kredit = Null
            End If
            If Frm87.TB30 <> vbNullString Then
                rs!jumlah_cas_kad_kredit = Format(Frm87.TB30, "0.00") 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
            Else
                rs!jumlah_cas_kad_kredit = Null
            End If
            If Frm87.TB31 <> vbNullString Then
                rs!jumlah_potongan_kad_kredit = Format(Frm87.TB31, "0.00") 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
            Else
                rs!jumlah_potongan_kad_kredit = Null
            End If
            If Frm87.TB21 <> vbNullString Then
                rs!duit_simpanan_kedai = Format(Frm87.TB21, "0.00") 'Cara Bayaran : Simpanan Duit Di Kedai
            Else
                rs!duit_simpanan_kedai = Null
            End If
            If Frm87.TB38 <> vbNullString Then
                rs!kad_debit = Format(Frm87.TB38, "0.00") 'Cara Bayaran : Kad Debit
            Else
                rs!kad_debit = Null
            End If
            If Frm87.L32_Text <> vbNullString Then
                rs!cas_kad_debit = Frm87.L32_Text 'Cara Bayaran : Jumlah Cas Kad Debit (%)
            Else
                rs!cas_kad_debit = Null
            End If
            If Frm87.TB39 <> vbNullString Then
                rs!jumlah_cas_kad_debit = Format(Frm87.TB39, "0.00") 'Cara Bayaran : Jumlah Cas Kad Debit (RM)
            Else
                rs!jumlah_cas_kad_debit = Null
            End If
            If Frm87.TB40 <> vbNullString Then
                rs!jumlah_potongan_kad_debit = Format(Frm87.TB40, "0.00") 'Cara Bayaran : Jumah Potongan Kad Debit (RM)
            Else
                rs!jumlah_potongan_kad_debit = Null
            End If
            If Frm87.TB32 <> vbNullString Then
                rs!jumlah = Format(Frm87.TB32, "0.00") 'Jumlah Harga Barang Tanpa GST (RM)
            Else
                rs!jumlah = Null
            End If
            If Frm87.TB13 <> vbNullString Then
                rs!jumlah_cukai_gst = Format(Frm87.TB13, "0.00") 'Jumlah Cukai GST (ZR + SR)
            Else
                rs!jumlah_cukai_gst = Null
            End If
            If Frm87.TB14 <> vbNullString Then
                rs!harga_barang_dengan_gst = Format(Frm87.TB14, "0.00") 'Jumlah Harga Barang Dengan GST (RM)
            Else
                rs!harga_barang_dengan_gst = Null
            End If
            If Frm87.TB17 <> vbNullString Then
                rs!adjustment = Format(Frm87.TB17, "0.00") 'Adjustment (RM)
            Else
                rs!adjustment = Null
            End If
            If Frm87.TB18 <> vbNullString Then
                rs!jumlah_bayaran = Format(Frm87.TB18, "0.00") 'Jumlah Harga Jualan (RM)
            Else
                rs!jumlah_bayaran = Null
            End If
            If Frm87.TB18 <> vbNullString Then
                rs!jumlah_bayaran = Format(Frm87.TB18, "0.00") 'Jumlah Harga Jualan (RM)
            Else
                rs!jumlah_bayaran = Null
            End If
            rs!flag_bayaran = 0 ' 0 : Pembeli Bayar , 1 : Kedai Bayar
            If Frm87.L25_Text <> vbNullString Then
                rs!gst_zr_harga = Format(Frm87.L25_Text, "0.00") 'Harga Keseluruhan Bagi Barang ZR
            Else
                rs!gst_zr_harga = Null
            End If
            If Frm87.L26_Text <> vbNullString Then
                rs!gst_zr_cukai = Format(Frm87.L26_Text, "0.00") 'Jumlah Cukai Bagi ZR
            Else
                rs!gst_zr_cukai = Null
            End If
            If Frm87.L23_Text <> vbNullString Then
                rs!gst_sr_harga = Format(Frm87.L23_Text, "0.00") 'Harga Keseluruhan Bagi Barang SR
            Else
                rs!gst_sr_harga = Null
            End If
            If Frm87.L24_Text <> vbNullString Then
                rs!gst_sr_cukai = Format(Frm87.L24_Text, "0.00") 'Jumlah Cukai Bagi SR
            Else
                rs!gst_sr_cukai = Null
            End If
            If Frm87.CB27 = 0 Then
                rs!gst_include = Null
            ElseIf Frm87.CB27 = 1 Then
                rs!gst_include = "**Harga Termasuk GST"
            End If
            If Frm87.TB42 <> vbNullString Then
                rs!harga_tanpa_gst = Format(Frm87.TB42, "0.00") 'Harga Keseluruhan Tanpa GST (RM)
            Else
                rs!harga_tanpa_gst = Null 'Harga Keseluruhan Tanpa GST (RM)
            End If
            If Frm87.CBB2 <> vbNullString Then
                Frm87_LM_EMP_NO = Split(Frm87.CBB2, "  |  ")(1)
                rs!no_pekerja = Frm87_LM_EMP_NO 'No. Pekerja
            End If
            rs!no_rujukan_pembeli = Frm87_LM_No_PEMBELI 'No. Rujukan Pembeli
            rs!kategori_pembeli = Frm87_LM_KATEGORI 'Kategori Pembeli
            rs!write_timestamp = Now
            
            rs.Update
            
            rs.Close
            Set rs = Nothing
'###Update Akaun Ansuran### - End

'### Update Senarai Ansuran ### - Start
'### Carian Jenis Ansuran ### - Start
    
            Frm87_FLAG_UPAH = 0
            Frm87_FLAG_ANSURAN = 0
            Frm87_FLAG_JELAS = 0
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 27_senarai_ansuran where ID='" & Frm87.L18_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!Status) Then Frm87_LM_STATUS_ASAL = rs!Status 'Status Asal
                If Not IsNull(rs!jenis_ansuran) Then
                    If rs!jenis_ansuran = 0 Then '0 : Harga Semasa , 1 : Harga Tetap
                        If Not IsNull(rs!UPAH) And Not IsNull(rs!JUMLAH_UPAH) Then
                            If Format(rs!UPAH, "0.00") = Format(rs!JUMLAH_UPAH, "0.00") Then
                                Frm87_FLAG_UPAH = 1
                            End If
                        End If
                        If Not IsNull(rs!berat_jualan) And Not IsNull(rs!JUMLAH_BERAT) Then
                            If Format(rs!berat_jualan, "0.00") = Format(rs!JUMLAH_BERAT, "0.00") Then
                                Frm87_FLAG_ANSURAN = 1
                            End If
                        End If
                        If Frm87_FLAG_UPAH = 1 And Frm87_FLAG_ANSURAN = 1 Then
                            Frm87_FLAG_JELAS = 1
                            rs!Status = "Jelas" 'Status
                            rs!tarikh_jelas = Frm87.DTPicker2 'Tarikh Jelas
                            rs.Update
                        Else
                            rs!Status = "Belum Jelas" 'Status
                            rs!tarikh_jelas = Null 'Tarikh Jelas
                            rs.Update
                        End If
                    ElseIf rs!jenis_ansuran = 1 Then
                        If Not IsNull(rs!harga_jualan) And Not IsNull(rs!jumlah_bayaran) Then
                            If Format(rs!harga_jualan, "0.00") = Format(rs!jumlah_bayaran, "0.00") Then
                                Frm87_FLAG_ANSURAN = 1
                            End If
                        End If
                        If Frm87_FLAG_ANSURAN = 1 Then
                            Frm87_FLAG_JELAS = 1
                            rs!Status = "Jelas" 'Status
                            rs!tarikh_jelas = Frm87.DTPicker2 'Tarikh Jelas
                            rs.Update
                        Else
                            rs!Status = "Belum Jelas" 'Status
                            rs!tarikh_jelas = Null 'Tarikh Jelas
                            rs.Update
                        End If
                    End If
                End If
            End If
            
            rs.Close
            Set rs = Nothing
'### Carian Jenis Ansuran ### - End
'### Update Senarai Ansuran ### - End

'### Jika Status Dari Jelas Bertukar Kepada Tidak Jelas ### - Start
            If Frm87_LM_STATUS_ASAL = "Jelas" Then
                If Frm87_FLAG_JELAS = 0 Then
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from Data_Database where no_siri_produk='" & Frm87_LM_No_SIRI & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                        If Frm87_LM_JENIS_PRODUK = 0 Then
                            If Frm87_LM_POTONG = 0 Then
                                rs!StatusItem = 15
                                rs!beza_berat = Format(rs!Berat, "0.00")
                            ElseIf Frm87_LM_POTONG = 1 Then
                                rs!StatusItem = 15
                                rs!beza_berat = Format(rs!beza_berat + Frm87_LM_BERAT_POTONG, "0.00") 'Beza Berat
                            End If
                        ElseIf Frm87_LM_JENIS_PRODUK = 1 Then
                            rs!StatusItem = 15
                        End If
                        rs.Update
                    End If
                    
                    rs.Close
                    Set rs = Nothing
                End If
            End If
'### Jika Status Dari Jelas Bertukar Kepada Tidak Jelas ### - End

'### Update Database Utama Jika Sudah Terjual ### - Start
            If Frm87_FLAG_JELAS = 1 Then
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from Data_Database where no_siri_produk='" & Frm87_LM_No_SIRI & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Frm87_LM_JENIS_PRODUK = 0 Then
                        If Frm87_LM_POTONG = 0 Then
                            rs!StatusItem = 19
                            rs!beza_berat = "0.00"
                        ElseIf Frm87_LM_POTONG = 1 Then
                            rs!StatusItem = 20
                            rs!beza_berat = Format(rs!Berat - Frm87_LM_BERAT_POTONG, "0.00") 'Beza Berat
                        End If
                        
                    ElseIf Frm87_LM_JENIS_PRODUK = 1 Then
                        rs!StatusItem = 19
                    End If
                    
                    rs.Update
                End If
                
                rs.Close
                Set rs = Nothing
            End If
'### Update Database Utama Jika Sudah Terjual ### - End
      
'### Update Log ### - Start
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & user & "] Edit Bayaran Ansuran. No. Ansuran [" & Frm87_LM_No_RESIT_ANSURAN & "]"
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
'### Update Log ### - End

'###Update Data Simpanan Duit Pelanggan### - Start
            If Format(Frm87.TB21, "0.00") <> "0.00" Then
            
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm87_LM_No_PEMBELI & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    Frm87_LM_JUMLAH_SIMPANAN = Frm87.L27_Text  'Jumlah Simpanan Yang Ada
                    Frm87_LM_GUNA_SIMPAN = Frm87.TB21  'Jumlah Simpanan Yang Hendak Digunakan
                    
                    rs!baki_simpanan = Format(Frm87_LM_JUMLAH_SIMPANAN - Frm87_LM_GUNA_SIMPAN, "0.00") 'Baki Simpanan
                    rs.Update
                End If
                
                rs.Close
                Set rs = Nothing
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 24_rekod_kewangan_pelanggan", cn, adOpenKeyset, adLockOptimistic
                
                rs.AddNew
                rs!tarikh = Frm87.DTPicker2 'Tarikh
                rs!jenis = 1 '0 : Simpanan , 1 : Penggunaan Duit
                rs!no_rujukan_pelanggan = Frm87_LM_No_PEMBELI 'No. Rujukan Pelanggan
                rs!no_resit = Frm87_LM_No_RESIT_ANSURAN 'No. Resit Ansuran
                rs!jumlah = Format(Frm87.TB21, "0.00") 'Jumlah Simpanan Yang Digunakan (RM)
                rs!jenis_penggunaan = 1 '0 : Belian Barangan Kemas , 1 : Ansuran , 2 : Tempahan (Deposit) , 3 : Servis , 4 : Tempahan (Ambilan Barang)
                rs!cawangan = G_CAWANGAN
                rs!Status = 1
                rs.Update
                
                rs.Close
                Set rs = Nothing
               
            End If
'###Update Data Simpanan Duit Pelanggan### - End

            Call Frm87_Initial_Setting
            
            If Frm87_FLAG_JELAS = 0 Then
            
                Note = "Data ansuran telah berjaya disimpan." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Cetak resit bayaran ansuran ?"
                
                Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
                
                If Answer = vbNo Then
                    Exit Sub
                End If
                If Answer = vbYes Then
                    G_No_RESIT_ANSURAN = Frm87_LM_No_RESIT_ANSURAN 'No. Resit Ansuran
                    Call Frm87_Resit_Ansuran
                End If
                
            ElseIf Frm87_FLAG_JELAS = 1 Then
            
                Note = "Data ansuran telah berjaya disimpan." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Pembeli ini telah menjelaskan semua bayaran bagi barang ini." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Status barang berubah kepada [JELAS]." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Cetak resit bayaran ansuran ?"
                
                Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
                
                If Answer = vbNo Then
                    Exit Sub
                End If
                If Answer = vbYes Then
                    G_No_RESIT_ANSURAN = Frm87_LM_No_RESIT_ANSURAN 'No. Resit Ansuran
                    Call Frm87_Resit_Ansuran
                End If
                
            End If
        End If
        
    End If
End If
End Sub
Private Sub CMD14_Click()
'on error resume next
Note = "Batal Urusan Ini ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
Else
    Frm87.Pic7.Visible = True
    Frm87.Pic5.Visible = False
End If
End Sub
Private Sub CMD15_Click()
'on error resume next
Frm87.Pic7.Visible = False
Frm87.Pic4.Visible = True
End Sub
Private Sub CMD16_Click()
'On Error Resume Next
Dim Err(15)
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim e As Double
Dim f As Double
Dim g As Double
Dim h As Double
Dim i As Double
Dim j As Double
Dim k As Double
Dim l As Double

a = 0
b = 0
c = 0
d = 0
e = 0
f = 0
g = 0
h = 0
i = 0
j = 0
k = 0
l = 0

DATA_SAVE = 0
Frm87_KOD_PURITY = vbNullString
Frm87_DULANG = vbNullStringFrm87_KOD_PURITY = vbNullString
Frm87_DULANG = vbNullString
Frm87_LM_KATEGORI = 1
'1:  Pelanggan Biasa
'2 : Member / Ahli
'3:  Pengedar
'4:  RAF
'5:  Normal Dealer(ND)
'6:  Master Dealer(MD)

If Frm87.L5_Text = vbNullString And Frm87.L6_Text = vbNullString Then
    x = x + 1
    Err(x) = "Tiada maklumat pembeli."
End If
If Frm87.TB2 = vbNullString Then
    x = x + 1
    Err(x) = "Tiada maklumat barang yang hendak dijual secara ansuran."
End If
If Frm87.CB14 = 0 And Frm87.CB15 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan jenis ansuran."
End If
If Frm87.L13_Text = 0 Then 'Flag Kategori Produk , 0 : BK , 1 : Permata
    If Frm87.TB4 = vbNullString Or (Frm87.TB4 <> vbNullString And Not IsNumeric(Frm87.TB4)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Berat Jualan]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm87.TB5 = vbNullString Or (Frm87.TB5 <> vbNullString And Not IsNumeric(Frm87.TB5)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Harga Semasa]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm87.TB6 = vbNullString Or (Frm87.TB6 <> vbNullString And Not IsNumeric(Frm87.TB6)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Upah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
ElseIf Frm87.L13_Text = 1 Then 'Flag Kategori Produk , 0 : BK , 1 : Permata
    If Frm87.TB7 = vbNullString Or (Frm87.TB7 <> vbNullString And Not IsNumeric(Frm87.TB7)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Harga Asal]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If
If Frm87.TB9 = vbNullString Or (Frm87.TB9 <> vbNullString And Not IsNumeric(Frm87.TB9)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Adjustment]. Hanya NOMBOR dibenarkan dalam ruangan ini. Masukkan [0.00] Jika Tiada Adjustment."
End If
If Frm87.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih nama pekerja"
End If

If Frm87.CB14 = 1 Then
    If IsNumeric(Frm87.L38_Text) Then a = Frm87.L38_Text
    If IsNumeric(Frm87.TB4) Then b = Frm87.TB4
    If IsNumeric(Frm87.L37_Text) Then c = Frm87.L37_Text
    If IsNumeric(Frm87.TB6) Then d = Frm87.TB6
    
    If a > b Then
        x = x + 1
        Err(x) = "Berat Jualan Kurang Dari Jumlah Berat Yang Telah Dibayar Oleh Pelanggan Ini."
    End If
    If c > d Then
        x = x + 1
        Err(x) = "Tetapan Upah Kurang Dari Jumlah Upah Yang Telah Dibayar Oleh Pelanggan Ini."
    End If
ElseIf Frm87.CB15 = 1 Then
    If IsNumeric(Frm87.L36_Text) Then e = Frm87.L36_Text
    If IsNumeric(Frm87.TB10) Then f = Frm87.TB10
    
    If e > f Then
        x = x + 1
        Err(x) = "Tetapan Harga Kurang Dari Jumlah Bayaran Yang Telah Dijelaskan Oleh Pelanggan Ini."
    End If
End If
'### Periksa Data Rekod Bayaran ### - End

'### Periksa Samada Maklumat Pembeli Diisi Dalam Kedua-dua Ruangan Berdaftar Dan Tidak Berdaftar ### - Start
If x = 0 Then
    If Frm87.L5_Text <> vbNullString And Frm87.L6_Text <> vbNullString Then
    
        MsgBox "Data bagi pembeli telah diisi bagi kedua-dua ruangan pembeli berdaftar dan tidak berdaftar." & vbCrLf & _
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
    If Frm87.L5_Text <> vbNullString And Frm87.L6_Text = vbNullString Then
        Note = "Adakah anda yakin untuk teruskan urusan jualan ini ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Data jualan akan disimpan ke dalam sistem." & vbCrLf & _
                vbNullString & vbCrLf & _
                "*** Oleh ini pembeli ini tidak berdaftar dengan sistem, oleh itu data dan rekod pembelian pembeli ini tidak akan disimpan di dalam sistem ***"
    End If
    If Frm87.L5_Text = vbNullString And Frm87.L6_Text <> vbNullString Then
        Note = "Adakah anda yakin untuk teruskan urusan jualan ini ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Data jualan akan disimpan ke dalam sistem."
    End If
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
        '###Carian Purity Item Ini ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where no_siri_Produk='" & Frm87.TB2 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            If Not IsNull(rs!kod_Purity) Then Frm87_KOD_PURITY = rs!kod_Purity 'Kod Purity
            If Not IsNull(rs!dulang) Then Frm87_DULANG = rs!dulang 'Dulang
            
        End If
        
        rs.Close
        Set rs = Nothing
        '###Carian Purity Item Ini ### - End
        
' ### Periksa kategori pembeli ### - Start
        If Frm87.L6_Text <> vbNullString Then
            If Frm28.L5_Text <> vbNullString Then
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
                If Not rs.EOF Then
                
                    If Not IsNull(rs!kategori_pelanggan) Then Frm87_LM_KATEGORI = rs!kategori_pelanggan
                    
                End If
                
                rs.Close
                Set rs = Nothing
                
            End If
        End If
' ### Periksa kategori pembeli ### - End
        
'###Masukkan Data Belian Ke Dalam Database Akaun Belian### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 27_senarai_ansuran where no_rujukan='" & Frm87.L11_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Frm87.L11_Text <> vbNullString Then
                rs!no_rujukan = Frm87.L11_Text
            Else
                rs!no_rujukan = Null
            End If
            If Frm87.L5_Text <> vbNullString Then
  
                rs!no_rujukan_pelanggan = Null 'No. Rujukan Pembeli
                rs!no_ic = Null 'Maklumat Pembeli : No. Kad Pengenalan
                If Frm26.TB1 <> vbNullString Then
                    rs!Nama = UCase(Frm26.TB1) 'Maklumat Pembeli : Nama
                Else
                    rs!Nama = Null 'Maklumat Pembeli : Nama
                End If
                If Frm26.TB2 <> vbNullString Then
                    rs!no_tel = UCase(Frm26.TB2) 'No. Telefon
                Else
                    rs!no_tel = Null 'No. Telefon
                End If

            End If
            If Frm87.L6_Text <> vbNullString Then
                If Frm28.L5_Text <> vbNullString Then
                    rs!no_rujukan_pelanggan = Frm28.L5_Text 'No. Rujukan Pembeli
                Else
                    rs!no_rujukan_pelanggan = Null 'No. Rujukan Pembeli
                End If
                If Frm28.L1_Text <> vbNullString Then
                    rs!Nama = Frm28.L1_Text 'Maklumat Pembeli : Nama
                Else
                    rs!Nama = Null 'Maklumat Pembeli : Nama
                End If
                If Frm28.L2_Text <> vbNullString Then
                    rs!no_ic = Frm28.L2_Text 'Maklumat Pembeli : No. Kad Pengenalan
                Else
                    rs!no_ic = Null 'Maklumat Pembeli : No. Kad Pengenalan
                End If
                If Frm28.L3_Text <> vbNullString Then
                    rs!no_tel = Frm28.L3_Text 'No. Telefon
                Else
                    rs!no_tel = Null 'No. Telefon
                End If
            End If
            
'Kategori Pembeli
'=================
'1:  Pelanggan
'2 : Member / Ahli
'3:  RAF
'4:  Pengedar
'5:  Normal Dealer (ND)
'6:  Master Dealer (MD)

            rs!kategori_pembeli = Frm87_LM_KATEGORI 'Kategori Pembeli
            
            If Frm87.TB2 <> vbNullString Then
                rs!no_siri_Produk = Frm87.TB2 'No. Siri Produk
            Else
                rs!no_siri_Produk = Null 'No. Siri Produk
            End If
            If Frm87.L13_Text = 0 Then
                rs!jenis_produk = 0 'Flat Kategori Produk , 0 : BK , 1 : Permata
            Else
                rs!jenis_produk = 1 'Flat Kategori Produk , 0 : BK , 1 : Permata
            End If
            If Frm87.L10_Text <> vbNullString Then
                rs!kategori_Produk = Frm87.L10_Text 'Kategori Produk
            Else
                rs!kategori_Produk = Null 'Kategori Produk
            End If
            If Frm87_KOD_PURITY <> vbNullString Then
                rs!purity = Frm87_KOD_PURITY 'Purity
            Else
                rs!purity = Null 'Purity
            End If
            If Frm87_DULANG <> vbNullString Then
                rs!dulang = Frm87_DULANG 'Dulang
            Else
                rs!dulang = Null 'Dulang
            End If
            If Frm87.TB3 <> vbNullString Then
                rs!Berat_Asal = Format(Frm87.TB3, "0.00") 'Berat Asal (g)
            Else
                rs!Berat_Asal = Null 'Berat Asal (g)
            End If
            If Frm87.TB4 <> vbNullString Then
                rs!berat_jualan = Format(Frm87.TB4, "0.00") 'Berat Jualan (g)
            Else
                rs!berat_jualan = Null 'Berat Jualan (g)
            End If
            If Frm87.TB5 <> vbNullString Then
                rs!harga_Semasa = Format(Frm87.TB5, "0.00") 'Harga Emas Semasa Masa Tempahan Dibuat
            Else
                rs!harga_Semasa = Null 'Harga Emas Semasa Masa Tempahan Dibuat
            End If
            If Frm87.TB6 <> vbNullString Then
                rs!UPAH = Format(Frm87.TB6, "0.00") 'Upah
            Else
                rs!UPAH = Null 'Upah
            End If
            If Frm87.TB7 <> vbNullString Then
                rs!harga_asal = Format(Frm87.TB7, "0.00") 'Harga Asal Jualan
            Else
                rs!harga_asal = Null 'Harga Asal Jualan
            End If
            If Frm87.TB9 <> vbNullString Then
                rs!adjustment = Format(Frm87.TB9, "0.00") 'Adjustment
            Else
                rs!adjustment = Null 'Adjustment
            End If
            If Frm87.TB10 <> vbNullString Then
                rs!harga_jualan = Format(Frm87.TB10, "0.00") 'Harga Jualan (RM)
            Else
                rs!harga_jualan = Null 'Harga Jualan (RM)
            End If

            If Frm87.CB14 = 1 Then '0 : Pengiraan Mengikut Harga Semasa , 1 : Harga Tetap
                rs!jenis_ansuran = 0
                If IsNumeric(Frm87.TB4) Then g = Frm87.TB4
                If IsNumeric(Frm87.L38_Text) Then h = Frm87.L38_Text
                
                rs!BAKI_BERAT = Format(g - h, "0.00") 'Baki Berat (g)
                
                If IsNumeric(Frm87.TB6) Then i = Frm87.TB6
                If IsNumeric(Frm87.L37_Text) Then j = Frm87.L37_Text
                
                rs!baki_upah = Format(i - j, "0.00") 'Baki Upah
            ElseIf Frm87.CB15 = 1 Then
                rs!jenis_ansuran = 1
                
                If IsNumeric(Frm87.TB10) Then k = Frm87.TB10
                If IsNumeric(Frm87.L36_Text) Then l = Frm87.L36_Text

                rs!baki_bayaran = Format(k - l, "0.00") 'Baki Jualan (RM)
            End If

            rs!tarikh = Frm87.DTPicker1 'Tarikh Tempahan

            If Frm87.CBB1 <> vbNullString Then
                Frm87_LM_EMP_NO = Split(Frm87.CBB1, "  |  ")(1)
                rs!no_rujukan_pekerja = Frm87_LM_EMP_NO 'No. Pekerja
            End If
            rs!write_timestamp = Now
            rs.Update
            DATA_SAVE = 1
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_SAVE = 1 Then
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & user & "] Edit Data Pendaftaran Belian Secara Ansuran , No. Siri [" & Frm87.TB2 & "]."
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
            
            Call Frm87_Initial_Setting
            MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
        End If
    End If
End If
End Sub
Private Sub CMD17_Click()
'on error resume next
Frm87.Pic2.Visible = False
Frm87.Pic4.Visible = True
End Sub
Private Sub CMD3_Click()
'On Error Resume Next
If Frm87.L6_Text = vbNullString Then
    
    If Frm87.L5_Text <> vbNullString Then
    
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
            
            Frm87.L5_Text = vbNullString 'Nama pembeli : Tidak berdaftar
            
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
Private Sub CMD5_Click()
'on error resume next
If Frm87.TB8 = vbNullString Then
    MsgBox "Sila masukkan No. Siri Produk.", vbInformation, "Info"
    Exit Sub
End If

If InStr(1, Frm87.TB8, "'") <> 0 Then
    MsgBox "No. Siri Produk mengandungi simbol yang tidak sah , ['].", vbInformation, "Info"
    
    Frm87.TB8 = vbNullString
    Exit Sub
End If

Call Frm87_Call_Product_Detail
End Sub
Private Sub CMD6_Click()
'On Error Resume Next
Dim Err(15)
DATA_SAVE = 0

Frm87_KOD_PURITY = vbNullString
Frm87_DULANG = vbNullString
Frm87_LM_KATEGORI = 1
'1:  Pelanggan Biasa
'2 : Member / Ahli
'3:  Pengedar
'4:  RAF
'5:  Normal Dealer(ND)
'6:  Master Dealer(MD)

If Frm87.L5_Text = vbNullString And Frm87.L6_Text = vbNullString Then
    x = x + 1
    Err(x) = "Tiada maklumat pembeli."
End If
If Frm87.TB2 = vbNullString Then
    x = x + 1
    Err(x) = "Tiada maklumat barang yang hendak dijual secara ansuran."
End If
If Frm87.CB14 = 0 And Frm87.CB15 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan jenis ansuran."
End If
If Frm87.L13_Text = 0 Then 'Flag Kategori Produk , 0 : BK , 1 : Permata
    If Frm87.TB4 = vbNullString Or (Frm87.TB4 <> vbNullString And Not IsNumeric(Frm87.TB4)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Berat Jualan]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm87.TB5 = vbNullString Or (Frm87.TB5 <> vbNullString And Not IsNumeric(Frm87.TB5)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Harga Semasa]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm87.TB6 = vbNullString Or (Frm87.TB6 <> vbNullString And Not IsNumeric(Frm87.TB6)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Upah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
ElseIf Frm87.L13_Text = 1 Then 'Flag Kategori Produk , 0 : BK , 1 : Permata
    If Frm87.TB7 = vbNullString Or (Frm87.TB7 <> vbNullString And Not IsNumeric(Frm87.TB7)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Harga Asal]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If
If Frm87.TB9 = vbNullString Or (Frm87.TB9 <> vbNullString And Not IsNumeric(Frm87.TB9)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Adjustment]. Hanya NOMBOR dibenarkan dalam ruangan ini. Masukkan [0.00] Jika Tiada Adjustment."
End If
If Frm87.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih nama pekerja"
End If

'### Periksa Samada Maklumat Pembeli Diisi Dalam Kedua-dua Ruangan Berdaftar Dan Tidak Berdaftar ### - Start
If x = 0 Then
    If Frm87.L5_Text <> vbNullString And Frm87.L6_Text <> vbNullString Then
    
        MsgBox "Data bagi pembeli telah diisi bagi kedua-dua ruangan pembeli berdaftar dan tidak berdaftar." & vbCrLf & _
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
    If Frm87.L5_Text <> vbNullString And Frm87.L6_Text = vbNullString Then
        Note = "Adakah anda yakin untuk teruskan urusan jualan ini ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Data jualan akan disimpan ke dalam sistem." & vbCrLf & _
                vbNullString & vbCrLf & _
                "*** Oleh ini pembeli ini tidak berdaftar dengan sistem, oleh itu data dan rekod pembelian pembeli ini tidak akan disimpan di dalam sistem ***"
    End If
    If Frm87.L5_Text = vbNullString And Frm87.L6_Text <> vbNullString Then
        Note = "Adakah anda yakin untuk teruskan urusan jualan ini ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Data jualan akan disimpan ke dalam sistem."
    End If
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
' ### Periksa kategori pembeli ### - Start
        If Frm87.L6_Text <> vbNullString Then
            If Frm28.L5_Text <> vbNullString Then
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
                If Not rs.EOF Then
                
                    If Not IsNull(rs!kategori_pelanggan) Then Frm87_LM_KATEGORI = rs!kategori_pelanggan
                    
                End If
                
                rs.Close
                Set rs = Nothing
                
            End If
        End If
' ### Periksa kategori pembeli ### - End
    
'###Carian Purity Item Ini ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where no_siri_Produk='" & Frm87.TB2 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            If Not IsNull(rs!kod_Purity) Then Frm87_KOD_PURITY = rs!kod_Purity 'Kod Purity
            If Not IsNull(rs!dulang) Then Frm87_DULANG = rs!dulang 'Dulang
            
        End If
        
        rs.Close
        Set rs = Nothing
'###Carian Purity Item Ini ### - End

        Frm87_LM_No_RUJUKAN_ANSURAN = Frm87.L11_Text 'No. Rujukan Ansuran
        
Re_Gen_No_Rujukan:
'###Masukkan Data Belian Ke Dalam Database Akaun Belian### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 27_senarai_ansuran where no_rujukan='" & Format(Frm87_LM_No_RUJUKAN_ANSURAN, "000000") & "'", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
            rs.AddNew
            If Frm87.L11_Text <> vbNullString Then
                rs!no_rujukan = Format(Frm87_LM_No_RUJUKAN_ANSURAN, "000000")
            Else
                rs!no_rujukan = Null
            End If
            
            If Frm87.L5_Text <> vbNullString Then
  
                rs!no_rujukan_pelanggan = Null 'No. Rujukan Pembeli
                rs!no_ic = Null 'Maklumat Pembeli : No. Kad Pengenalan
                If Frm26.TB1 <> vbNullString Then
                    rs!Nama = UCase(Frm26.TB1) 'Maklumat Pembeli : Nama
                Else
                    rs!Nama = Null 'Maklumat Pembeli : Nama
                End If
                If Frm26.TB2 <> vbNullString Then
                    rs!no_tel = UCase(Frm26.TB2) 'No. Telefon
                Else
                    rs!no_tel = Null 'No. Telefon
                End If

            End If
            If Frm87.L6_Text <> vbNullString Then
                If Frm28.L5_Text <> vbNullString Then
                    rs!no_rujukan_pelanggan = Frm28.L5_Text 'No. Rujukan Pembeli
                Else
                    rs!no_rujukan_pelanggan = Null 'No. Rujukan Pembeli
                End If
                If Frm28.L1_Text <> vbNullString Then
                    rs!Nama = Frm28.L1_Text 'Maklumat Pembeli : Nama
                Else
                    rs!Nama = Null 'Maklumat Pembeli : Nama
                End If
                If Frm28.L2_Text <> vbNullString Then
                    rs!no_ic = Frm28.L2_Text 'Maklumat Pembeli : No. Kad Pengenalan
                Else
                    rs!no_ic = Null 'Maklumat Pembeli : No. Kad Pengenalan
                End If
                If Frm28.L3_Text <> vbNullString Then
                    rs!no_tel = Frm28.L3_Text 'No. Telefon
                Else
                    rs!no_tel = Null 'No. Telefon
                End If
            End If
            
'Kategori Pembeli
'=================
'1:  Pelanggan
'2 : Member / Ahli
'3:  RAF
'4:  Pengedar
'5:  Normal Dealer (ND)
'6:  Master Dealer (MD)

            rs!kategori_pembeli = Frm87_LM_KATEGORI 'Kategori Pembeli
            
            If Frm87.TB2 <> vbNullString Then
                rs!no_siri_Produk = Frm87.TB2 'No. Siri Produk
            Else
                rs!no_siri_Produk = Null 'No. Siri Produk
            End If
            If Frm87.L13_Text = 0 Then
                rs!jenis_produk = 0 'Flat Kategori Produk , 0 : BK , 1 : Permata
            Else
                rs!jenis_produk = 1 'Flat Kategori Produk , 0 : BK , 1 : Permata
            End If
            If Frm87.L10_Text <> vbNullString Then
                rs!kategori_Produk = Frm87.L10_Text 'Kategori Produk
            Else
                rs!kategori_Produk = Null 'Kategori Produk
            End If
            If Frm87_KOD_PURITY <> vbNullString Then
                rs!purity = Frm87_KOD_PURITY 'Purity
            Else
                rs!purity = Null 'Purity
            End If
            If Frm87_DULANG <> vbNullString Then
                rs!dulang = Frm87_DULANG 'Dulang
            Else
                rs!dulang = Null 'Dulang
            End If
            If Frm87.TB3 <> vbNullString Then
                rs!Berat_Asal = Format(Frm87.TB3, "0.00") 'Berat Asal (g)
            Else
                rs!Berat_Asal = Null 'Berat Asal (g)
            End If
            If Frm87.TB4 <> vbNullString Then
                rs!berat_jualan = Format(Frm87.TB4, "0.00") 'Berat Jualan (g)
            Else
                rs!berat_jualan = Null 'Berat Jualan (g)
            End If
            If Frm87.TB5 <> vbNullString Then
                rs!harga_Semasa = Format(Frm87.TB5, "0.00") 'Harga Emas Semasa Masa Tempahan Dibuat
            Else
                rs!harga_Semasa = Null 'Harga Emas Semasa Masa Tempahan Dibuat
            End If
            If Frm87.TB6 <> vbNullString Then
                rs!UPAH = Format(Frm87.TB6, "0.00") 'Upah
            Else
                rs!UPAH = Null 'Upah
            End If
            If Frm87.TB7 <> vbNullString Then
                rs!harga_asal = Format(Frm87.TB7, "0.00") 'Harga Asal Jualan
            Else
                rs!harga_asal = Null 'Harga Asal Jualan
            End If
            If Frm87.TB9 <> vbNullString Then
                rs!adjustment = Format(Frm87.TB9, "0.00") 'Adjustment
            Else
                rs!adjustment = Null 'Adjustment
            End If
            If Frm87.TB10 <> vbNullString Then
                rs!harga_jualan = Format(Frm87.TB10, "0.00") 'Harga Jualan (RM)
            Else
                rs!harga_jualan = Null 'Harga Jualan (RM)
            End If
            If Frm87.CB14 = 1 Then '0 : Pengiraan Mengikut Harga Semasa , 1 : Harga Tetap
                rs!jenis_ansuran = 0
                rs!JUMLAH_BERAT = "0.00" 'Jumlah Berat Yang Telah Dijelaskan
                If Frm87.TB4 <> vbNullString Then
                    rs!BAKI_BERAT = Format(Frm87.TB4, "0.00") 'Baki Berat (g)
                Else
                    rs!BAKI_BERAT = Null 'Baki Berat (g)
                End If
                If Frm87.TB6 <> vbNullString Then
                    rs!baki_upah = Format(Frm87.TB6, "0.00") 'Baki Upah
                Else
                    rs!baki_upah = Null 'Baki Upah
                End If
                rs!JUMLAH_UPAH = "0.00" 'Jumlah Upah Yang Telah Dijelaskan
            ElseIf Frm87.CB15 = 1 Then
                rs!jenis_ansuran = 1
                rs!JUMLAH_BERAT = Null 'Jumlah Berat Yang Telah Dijelaskan
                If Frm87.TB10 <> vbNullString Then
                    rs!baki_bayaran = Format(Frm87.TB10, "0.00") 'Baki Jualan (RM)
                Else
                    rs!baki_bayaran = Null 'Baki Jualan (RM)
                End If
                rs!baki_upah = Null 'Baki Upah
                rs!JUMLAH_UPAH = Null 'Jumlah Upah Yang Telah Dijelaskan
            End If
            rs!tarikh = Frm87.DTPicker1 'Tarikh Tempahan
            rs!Status = "Belum Jelas"
            rs!jumlah_bayaran = "0.00" 'Jumlah Bayaran Yang Telah Dijelaskan (RM)
            If Frm87.CBB1 <> vbNullString Then
                Frm87_LM_EMP_NO = Split(Frm87.CBB1, "  |  ")(1)
                rs!no_rujukan_pekerja = Frm87_LM_EMP_NO 'No. Pekerja
            End If
            rs!write_timestamp = Now
            rs.Update
            DATA_SAVE = 1
        Else
            Frm87_LM_No_RUJUKAN_ANSURAN = Frm87_LM_No_RUJUKAN_ANSURAN + 1
            Frm87.L11_Text = Frm87_LM_No_RUJUKAN_ANSURAN 'No. Rujukan Ansuran
            
            rs.Close
            Set rs = Nothing
            GoTo Re_Gen_No_Rujukan:
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_SAVE = 1 Then
'### Update Table Database Bagi Item Ini ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from Data_Database where no_siri_produk='" & Frm87.TB2 & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                rs!StatusItem = 15
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
'### Update Table Database Bagi Item Ini ### - End

            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If rs!Default1 = "Default" Then
                    If IsNumeric(Frm87.L11_Text) Then
                        rs!no_rujukan_ansuran = Frm87.L11_Text + 1
                        rs.Update
                    End If
                End If
            End If
            
            rs.Close
            Set rs = Nothing
        
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & user & "] Daftar Belian Secara Ansuran , No. Siri [" & Frm87.TB2 & "]."
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
            
            Call Frm87_Initial_Setting
            MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
        End If
    End If
End If
End Sub
Private Sub CMD9_Click()
'On Error Resume Next
Dim Err(30)
Dim Frm87_LM_BERAT_BAYARAN As Double
Dim Frm87_BAKI_BERAT As Double
Dim Frm87_LM_UPAH As Double
Dim Frm87_LM_BAKI_UPAH As Double
Dim Frm87_LM_HARGA As Double
Dim Frm87_LM_JUMLAH_BAYARAN As Double
Dim Frm87_LM_JUMLAH_SIMPANAN As Double
Dim Frm87_LM_GUNA_SIMPAN As Double
Dim Frm87_LM_JUMLAH_BAYARAN_ASAL  As Double
Dim Frm87_LM_JUMLAH_BERAT_JUALAN_ASAL As Double
Dim Frm87_LM_JUMLAH_BERAT_ASAL As Double
Dim Frm87_LM_UPAH_ASAL As Double
Dim Frm87_LM_UPAH_JUALAN_ASAL As Double
Dim Frm87_LM_HARGA_JUALAN_ASAL As Double
Dim Frm87_LM_BERAT_POTONG As Double
Dim aaa As Double
Dim bbb As Double

DATA_SAVE = 0
Frm87_LM_BERAT_BAYARAN = 0
Frm87_BAKI_BERAT = 0
Frm87_LM_UPAH = 0
Frm87_LM_BAKI_UPAH = 0
Frm87_LM_HARGA = 0
Frm87_LM_JUMLAH_BAYARAN = 0
Frm87_LM_JUMLAH_SIMPANAN = 0  'Jumlah Simpanan Yang Ada
Frm87_LM_GUNA_SIMPAN = 0  'Jumlah Simpanan Yang Hendak Digunakan
Frm87_LM_BERAT_POTONG = 0

Frm87_LM_JUMLAH_BAYARAN_ASAL = 0
Frm87_LM_JUMLAH_BERAT_JUALAN_ASAL = 0
Frm87_LM_JUMLAH_BERAT_ASAL = 0
Frm87_LM_UPAH_ASAL = 0
Frm87_LM_UPAH_JUALAN_ASAL = 0
Frm87_LM_HARGA_JUALAN_ASAL = 0

aaa = 0
bbb = 0

If Frm87.CB20 = 0 And Frm87.CB21 = 0 Then
    x = x + 1
    Err(x) = "Sila Buat Pilihan Jenis Bayaran."
End If
If Frm87.CB20 = 1 Then
    If Frm87.TB12 = vbNullString Or (Frm87.TB12 <> vbNullString And Not IsNumeric(Frm87.TB12)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Bayaran Ansuran]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm87.CB22 = 0 And Frm87.CB23 = 0 Then
        x = x + 1
        Err(x) = "Sila Buat Pilihan Jenis GST Bagi Bayaran Ansuran."
    End If
    If Frm87.Pic6.Visible = True Then
        If Frm87.TB15 = vbNullString Or (Frm87.TB15 <> vbNullString And Not IsNumeric(Frm87.TB15)) Then
            x = x + 1
            Err(x) = "Sila Masukkan [Harga Emas Semasa]. Hanya NOMBOR dibenarkan dalam ruangan ini."
        End If
    Else
        If (Frm87.TB12 <> vbNullString And IsNumeric(Frm87.TB12)) And (Frm87.L28_Text <> vbNullString And IsNumeric(Frm87.L28_Text)) Then
            aaa = Frm87.TB12
            bbb = Frm87.L28_Text
            
            If aaa > bbb Then
                x = x + 1
                Err(x) = "Jumlah Bayaran Ansuran Melebihi Jumlah Baki."
            End If
        End If
    End If
    If (Frm87.TB16 <> vbNullString And IsNumeric(Frm87.TB16)) And (Frm87.L19_Text <> vbNullString And IsNumeric(Frm87.L19_Text)) Then
        Frm87_LM_BERAT_BAYARAN = Frm87.TB16 'Berat Bayaran Kali Ini
        Frm87_BAKI_BERAT = Frm87.L19_Text 'Baki Berat
        
        If Frm87_LM_BERAT_BAYARAN > Frm87_BAKI_BERAT Then
            x = x + 1
            Err(x) = "Berat Bagi Bayaran Adalah Melebihi Baki Berat Yang Tinggal."
        End If
    End If
End If

If Frm87.CB21 = 1 Then
    If Frm87.TB19 = vbNullString Or (Frm87.TB19 <> vbNullString And Not IsNumeric(Frm87.TB19)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Jumlah Upah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    'If Frm87.CB18 = 0 And Frm87.CB19 = 0 Then
    '    X = X + 1
    '    Err(X) = "Sila Buat Pilihan Jenis GST Bagi Bayaran Upah."
    'End If
    If (Frm87.TB19 <> vbNullString And IsNumeric(Frm87.TB19)) And (Frm87.L20_Text <> vbNullString And IsNumeric(Frm87.L20_Text)) Then
        Frm87_LM_UPAH = Frm87.TB19 'Bayaran Upah
        Frm87_LM_BAKI_UPAH = Frm87.L20_Text 'Baki Upah
        
        If Frm87_LM_UPAH > Frm87_LM_BAKI_UPAH Then
            x = x + 1
            Err(x) = "Upah Bagi Bayaran Adalah Melebihi Baki Upah Yang Tinggal."
        End If
    End If
End If
If Frm87.TB17 = vbNullString Or (Frm87.TB17 <> vbNullString And Not IsNumeric(Frm87.TB17)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Adjustment]. Hanya NOMBOR dibenarkan dalam ruangan ini. Masukkan [0.00] Jika Tiada Adjustment."
End If
If Frm87.TB27 = vbNullString Or (Frm87.TB27 <> vbNullString And Not IsNumeric(Frm87.TB27)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR Dibenarkan Dalam Ruangan Bayaran Secara Tunai. Sila Masukkan 0 Jika Tiada Bayaran Tunai."
End If
If Frm87.TB28 = vbNullString Or (Frm87.TB28 <> vbNullString And Not IsNumeric(Frm87.TB28)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR Dibenarkan Dalam Ruangan Bayaran Secara Bank In. Sila Masukkan 0 Jika Tiada Bayaran Bank In."
End If
If Frm87.TB29 = vbNullString Or (Frm87.TB29 <> vbNullString And Not IsNumeric(Frm87.TB29)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR Dibenarkan Dalam Ruangan Bayaran Secara Kad Kredit. Sila Masukkan 0 Jika Tiada Bayaran Kad Kredit."
End If
If Frm87.TB21 = vbNullString Or (Frm87.TB21 <> vbNullString And Not IsNumeric(Frm87.TB21)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR Dibenarkan Dalam Ruangan Bayaran Secara Duit Simpanan Di Kedai. Sila Masukkan 0 Jika Tiada Bayaran Simpanan Di Kedai."
End If
If Frm87.TB38 = vbNullString Or (Frm87.TB38 <> vbNullString And Not IsNumeric(Frm87.TB38)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR Dibenarkan Dalam Ruangan Bayaran Secara Debit Kad. Sila Masukkan 0 Jika Tiada Bayaran Debit Kad."
End If

If (Frm87.TB21 <> vbNullString And IsNumeric(Frm87.TB21)) And (Frm87.L27_Text <> vbNullString And IsNumeric(Frm87.L27_Text)) Then
    Frm87_LM_JUMLAH_SIMPANAN = Frm87.L27_Text  'Jumlah Simpanan Yang Ada
    Frm87_LM_GUNA_SIMPAN = Frm87.TB21  'Jumlah Simpanan Yang Hendak Digunakan
    
    If Frm87_LM_GUNA_SIMPAN > Frm87_LM_JUMLAH_SIMPANAN Then
        x = x + 1
        Err(x) = "Jumlah Simpanan Yang Hendak Digunakan Melebihi Simpanan Yang Ada."
    End If
End If

If (Frm87.TB32 <> vbNullString And IsNumeric(Frm87.TB32)) And (Frm87.TB18 <> vbNullString And IsNumeric(Frm87.TB18)) Then
    Frm87_LM_JUMLAH_BAYARAN = Frm87.TB32 'Jumlah Bayaran
    Frm87_LM_HARGA = Frm87.TB18 'Harga Keseluruhan
    
    If Frm87_LM_JUMLAH_BAYARAN <> Frm87_LM_HARGA Then
        x = x + 1
        Err(x) = "Jumlah Bayaran Tidak Sama Dengan Jumlah Harga Barang."
    End If
End If
If Frm87.CBB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih Nama Pekerja"
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
    
'### Carian Jenis Ansuran ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 27_senarai_ansuran where ID='" & Frm87.L18_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!no_siri_Produk) Then Frm87_LM_No_SIRI = rs!no_siri_Produk 'No. Siri Produk
        If Not IsNull(rs!jenis_produk) Then
            If rs!jenis_produk = 1 Then
                Frm87_LM_JENIS_PRODUK = 1
            ElseIf rs!jenis_produk = 0 Then
                Frm87_LM_JENIS_PRODUK = 0
                
                Frm87_LM_POTONG = 0
                If rs!Berat_Asal <> rs!berat_jualan Then
                    Frm87_LM_BERAT_POTONG = rs!berat_jualan
                    Frm87_LM_POTONG = 1
                End If
            End If
        End If
        If Not IsNull(rs!jenis_ansuran) Then
            If rs!jenis_ansuran = 0 Then '0 : Harga Semasa , 1 : Harga Tetap
                Frm87_LM_JENIS = 0
                
                If Not IsNull(rs!jumlah_bayaran) Then
                    Frm87_LM_JUMLAH_BAYARAN_ASAL = rs!jumlah_bayaran 'Jumlah Bayaran Terkumpul - Asal (RM)
                    
                    rs!jumlah_bayaran = Format(Frm87_LM_JUMLAH_BAYARAN_ASAL + Frm87.TB20, "0.00")
                End If
                
                If Frm87.CB20 = 1 Then
                    If Not IsNull(rs!berat_jualan) Then Frm87_LM_JUMLAH_BERAT_JUALAN_ASAL = rs!berat_jualan 'Jumlah Berat Jualan - Asal (RM)
                    
                    If Not IsNull(rs!JUMLAH_BERAT) Then
                        Frm87_LM_JUMLAH_BERAT_ASAL = rs!JUMLAH_BERAT 'Jumlah Berat Yang Telah Dijelaskan - Asal (RM)
                    
                        rs!JUMLAH_BERAT = Format(Frm87.TB16 + Frm87_LM_JUMLAH_BERAT_ASAL, "0.00") 'Jumlah Berat Yang Telah Dijelaskan (g)
                    End If
                
                    rs!BAKI_BERAT = Format(Frm87_LM_JUMLAH_BERAT_JUALAN_ASAL - rs!JUMLAH_BERAT, "0.00") 'Baki Berat
                End If
                
                If Frm87.CB21 = 1 Then
                    If Not IsNull(rs!JUMLAH_UPAH) Then
                        Frm87_LM_UPAH_ASAL = rs!JUMLAH_UPAH 'Jumlah Upah Yang Telah Dijelaskan - Asal (RM)
                        
                        rs!JUMLAH_UPAH = Format(Frm87.TB19 + Frm87_LM_UPAH_ASAL, "0.00")
                    End If
                    If Not IsNull(rs!UPAH) Then Frm87_LM_UPAH_JUALAN_ASAL = rs!UPAH 'Jumlah Tetapan Upah - Asal (RM)
                    
                    rs!baki_upah = Format(Frm87_LM_UPAH_JUALAN_ASAL - rs!JUMLAH_UPAH, "0.00") 'Baki Upah
                End If
            ElseIf rs!jenis_ansuran = 1 Then
                Frm87_LM_JENIS = 1
                If Not IsNull(rs!jumlah_bayaran) Then Frm87_LM_JUMLAH_BAYARAN_ASAL = rs!jumlah_bayaran 'Jumlah Bayaran Terkumpul - Asal (RM)
                If Not IsNull(rs!harga_jualan) Then Frm87_LM_HARGA_JUALAN_ASAL = rs!harga_jualan 'Jumlah Harga Jualan - Asal (RM)
                
                aaa = Frm87.TB12
                rs!jumlah_bayaran = Format(Frm87_LM_JUMLAH_BAYARAN_ASAL + aaa, "0.00") 'Jumlah Terkumpul Yang Baru
                rs!baki_bayaran = Format(Frm87_LM_HARGA_JUALAN_ASAL - (Frm87_LM_JUMLAH_BAYARAN_ASAL + aaa), "0.00") 'Baki Bayaran
            End If
        End If
        If Not IsNull(no_rujukan_pelanggan) Then Frm87_LM_No_PEMBELI = rs!no_rujukan_pelanggan 'No. Rujukan Pembeli
        If Not IsNull(rs!kategori_pembeli) Then Frm87_LM_KATEGORI = rs!kategori_pembeli 'Kategori Pembeli
        
        rs.Update
    End If
    
    rs.Close
    Set rs = Nothing
'### Carian Jenis Ansuran ### - End
    
'### Update Rekod Ansuran ### - Start
        Frm87_LM_No_RESIT_ANSURAN = Frm87.L12_Text 'No. Resit Ansuran
        
Re_Gen_No_Rujukan:
'###Update Bayaran Ansuran### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 28_rekod_ansuran where no_resit_ansuran='" & "ANS" & Format(Frm87_LM_No_RESIT_ANSURAN, "000000") & "'", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
            rs.AddNew
            If Frm87.L18_Text <> vbNullString Then
                rs!id_database_reg = Frm87.L18_Text 'No. ID Dari Database Senarai Pembeli Ansuran
            Else
                rs!id_database_reg = Null
            End If
            If Frm87.L12_Text <> vbNullString Then
                rs!no_resit_ansuran = "ANS" & Format(Frm87_LM_No_RESIT_ANSURAN, "000000") 'No. Resit Ansuran
            Else
                rs!no_resit_ansuran = Null
            End If
            If Frm87.CB20 = 1 Then
                rs!FLAG_ANSURAN = 1 'Flag samada ada bayaran ansuran atau tidak , 0 : Tiada bayaran ansuran , 1 : Ada bayaran ansuran
                If Frm87.TB12 <> vbNullString Then
                    rs!jumlah_ansuran = Format(Frm87.TB12, "0.00") 'Jumlah Bayran Ansuran
                Else
                    rs!jumlah_ansuran = Null 'Jumlah Bayran Ansuran
                End If
'                If Frm87.CB22 = 1 Then
'                    rs!flag_ansuran_zr = 1 'Flag samada ada zr bagi bayaran ansuran , 0 : Tiada ZR , 1 : Ada ZR
'                Else
'                    rs!flag_ansuran_zr = 0 'Flag samada ada zr bagi bayaran ansuran , 0 : Tiada ZR , 1 : Ada ZR
'                End If
'                If Frm87.CB23 = 1 Then
'                    rs!flag_ansuran_sr = 1 'Flag samada ada sr bagi bayaran ansuran , 0 : Tiada SR , 1 : Ada SR
'                Else
'                    rs!flag_ansuran_sr = 0 'Flag samada ada sr bagi bayaran ansuran , 0 : Tiada SR , 1 : Ada SR
'                End If
'                If Frm87.L21_Text <> vbNullString Then
'                    rs!ansuran_gst = Format(Frm87.L21_Text, "0.00") 'Jumlah GST bagi bayaran ansuran (RM)
'                Else
'                    rs!ansuran_gst = Null
'                End If
                If Frm87.Pic6.Visible = True Then
                    If Frm87.TB15 <> vbNullString Then
                        rs!harga_Semasa = Format(Frm87.TB15, "0.00") 'Harga Semasa
                    Else
                        rs!harga_Semasa = "0.00"
                    End If
                    If Frm87.TB16 <> vbNullString Then
                        rs!berat_diperoleh = Format(Frm87.TB16, "0.00") 'Berat Diperolehi
                    Else
                        rs!berat_diperoleh = "0.00"
                    End If
                Else
                    rs!harga_Semasa = Null
                    rs!berat_diperoleh = Null
                End If
            Else
                rs!FLAG_ANSURAN = 0 'Flag samada ada bayaran ansuran atau tidak , 0 : Tiada bayaran ansuran , 1 : Ada bayaran ansuran
                rs!jumlah_ansuran = Null 'Jumlah Bayran Ansuran
'                rs!flag_ansuran_zr = Null 'Flag samada ada zr bagi bayaran ansuran , 0 : Tiada ZR , 1 : Ada ZR
'                rs!flag_ansuran_sr = Null 'Flag samada ada sr bagi bayaran ansuran , 0 : Tiada SR , 1 : Ada SR
'                rs!ansuran_gst = Null 'Jumlah GST bagi bayaran ansuran (RM)
                rs!harga_Semasa = Null
                rs!berat_diperoleh = Null
            End If
            
            If Frm87.CB22 = 1 Then
                rs!flag_ansuran_zr = 1 'Flag samada ada zr bagi bayaran ansuran , 0 : Tiada ZR , 1 : Ada ZR
            Else
                rs!flag_ansuran_zr = 0 'Flag samada ada zr bagi bayaran ansuran , 0 : Tiada ZR , 1 : Ada ZR
            End If
            If Frm87.CB23 = 1 Then
                rs!flag_ansuran_sr = 1 'Flag samada ada sr bagi bayaran ansuran , 0 : Tiada SR , 1 : Ada SR
            Else
                rs!flag_ansuran_sr = 0 'Flag samada ada sr bagi bayaran ansuran , 0 : Tiada SR , 1 : Ada SR
            End If
            If Frm87.L21_Text <> vbNullString Then
                rs!ansuran_gst = Format(Frm87.L21_Text, "0.00") 'Jumlah GST bagi bayaran ansuran (RM)
            Else
                rs!ansuran_gst = Null
            End If

            If Frm87.CB21 = 1 Then
                rs!flag_upah = 1 'Flag samada ada bayaran upah atau tidak , 0 : Tiada Bayaran Upah , 1 : Ada Bayaran Upah
                If Frm87.TB19 <> vbNullString Then
                    rs!JUMLAH_UPAH = Format(Frm87.TB19, "0.00") 'Jumlah Bayran Upah
                Else
                    rs!JUMLAH_UPAH = Null 'Jumlah Bayran Ansuran
                End If
'                If Frm87.CB18 = 1 Then
'                    rs!flag_upah_zr = 1 'Flag samada ada zr bagi bayaran upah , 0 : Tiada ZR , 1 : Ada ZR
'                Else
'                    rs!flag_upah_zr = 0 'Flag samada ada zr bagi bayaran upah , 0 : Tiada ZR , 1 : Ada ZR
'                End If
'                If Frm87.CB19 = 1 Then
'                    rs!flag_upah_sr = 1 'Flag samada ada sr bagi bayaran upah , 0 : Tiada SR , 1 : Ada SR
'                Else
'                    rs!flag_upah_sr = 0 'Flag samada ada sr bagi bayaran upah , 0 : Tiada SR , 1 : Ada SR
'                End If
'                If Frm87.L22_Text <> vbNullString Then
'                    rs!upah_gst = Format(Frm87.L22_Text, "0.00") 'Jumlah GST Bagi Upah (RM)
'                Else
'                    rs!upah_gst = Null
'                End If
'            Else
'                rs!flag_upah = 0 'Flag samada ada bayaran upah atau tidak , 0 : Tiada Bayaran Upah , 1 : Ada Bayaran Upah
'                rs!JUMLAH_UPAH = Null 'Jumlah Bayran Ansuran
'                rs!flag_upah_zr = Null 'Flag samada ada zr bagi bayaran upah , 0 : Tiada ZR , 1 : Ada ZR
'                rs!flag_upah_sr = Null 'Flag samada ada sr bagi bayaran upah , 0 : Tiada SR , 1 : Ada SR
'                rs!upah_gst = Null 'Jumlah GST Bagi Upah (RM)
            End If
            rs!jenis_ansuran = Frm87_LM_JENIS 'Jenis Ansuran , 0 : Harga Semasa , 1 : Harga Tetap
            If Frm87.TB20 <> vbNullString Then
                rs!jumlah_bayaran = Format(Frm87.TB20, "0.00") 'Jumlah Ansuran + Jumlah Upah
            Else
                rs!jumlah_bayaran = Null
            End If
            If Frm87.L17_Text <> vbNullString Then
                rs!kadar_gst = Frm87.L17_Text 'Kadar GST (%)
            Else
                rs!kadar_gst = Null
            End If
            If Frm87.TB13 <> vbNullString Then
                rs!jumlah_gst = Format(Frm87.TB13, "0.00") 'Jumlah GST
            Else
                rs!jumlah_gst = Null
            End If
            If Frm87.TB14 <> vbNullString Then
                rs!jumlah_asal = Format(Frm87.TB14, "0.00") 'Jumlah Asal (Ansuran + Upah) + GST
            Else
                rs!jumlah_asal = Null
            End If
            If Frm87.TB17 <> vbNullString Then
                rs!adjustment = Format(Frm87.TB17, "0.00") 'Adjustment
            Else
                rs!adjustment = Null
            End If
            If Frm87.TB18 <> vbNullString Then
                rs!jumlah_keseluruhan = Format(Frm87.TB18, "0.00") 'Jumlah bayaran selepas adjustment
            Else
                rs!jumlah_keseluruhan = Null
            End If
            rs!tarikh = Frm87.DTPicker2 'Tarikh Bayaran
            If Frm87.CBB2 <> vbNullString Then
                Frm87_LM_EMP_NO = Split(Frm87.CBB2, "  |  ")(1)
                rs!no_rujukan_pekerja = Frm87_LM_EMP_NO 'No. Pekerja
            End If
            rs!write_timestamp = Now
            rs.Update
            DATA_SAVE = 1
        Else
            Frm87_LM_No_RESIT_ANSURAN = Frm87_LM_No_RESIT_ANSURAN + 1
            Frm87.L12_Text = Frm87_LM_No_RESIT_ANSURAN 'No. Resit Ansuran
            
            rs.Close
            Set rs = Nothing
            GoTo Re_Gen_No_Rujukan:
        End If
        
        rs.Close
        Set rs = Nothing
'###Update Bayaran Ansuran### - End
        
        If DATA_SAVE = 1 Then
        
'###Update Akaun Ansuran### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 29_akaun_ansuran", cn, adOpenKeyset, adLockOptimistic
            
            rs.AddNew
            If Frm87.L12_Text <> vbNullString Then
                rs!no_resit = "ANS" & Format(Frm87_LM_No_RESIT_ANSURAN, "000000") 'No. Resit Ansuran
            Else
                rs!no_resit = Null
            End If
            rs!tarikh = Frm87.DTPicker2 'Tarikh Bayaran
            If Frm87.TB27 <> vbNullString Then
                rs!tunai = Format(Frm87.TB27, "0.00") 'Cara Bayaran : Tunai
            Else
                rs!tunai = Null
            End If
            If Frm87.TB28 <> vbNullString Then
                rs!bank_in = Format(Frm87.TB28, "0.00") 'Cara Bayaran : Bank In
            Else
                rs!bank_in = Null
            End If
            If Frm87.TB29 <> vbNullString Then
                rs!kad_kredit = Format(Frm87.TB29, "0.00") 'Cara Bayaran : Kad Kredit
            Else
                rs!kad_kredit = Null
            End If
            If Frm87.L31_Text <> vbNullString Then
                rs!cas_Kad_Kredit = Frm87.L31_Text 'Cara Bayaran : Cas Kad Kredit (%)
            Else
                rs!cas_Kad_Kredit = Null
            End If
            If Frm87.TB30 <> vbNullString Then
                rs!jumlah_cas_kad_kredit = Format(Frm87.TB30, "0.00") 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
            Else
                rs!jumlah_cas_kad_kredit = Null
            End If
            If Frm87.TB31 <> vbNullString Then
                rs!jumlah_potongan_kad_kredit = Format(Frm87.TB31, "0.00") 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
            Else
                rs!jumlah_potongan_kad_kredit = Null
            End If
            If Frm87.TB21 <> vbNullString Then
                rs!duit_simpanan_kedai = Format(Frm87.TB21, "0.00") 'Cara Bayaran : Simpanan Duit Di Kedai
            Else
                rs!duit_simpanan_kedai = Null
            End If
            If Frm87.TB38 <> vbNullString Then
                rs!kad_debit = Format(Frm87.TB38, "0.00") 'Cara Bayaran : Kad Debit
            Else
                rs!kad_debit = Null
            End If
            If Frm87.L32_Text <> vbNullString Then
                rs!cas_kad_debit = Frm87.L32_Text 'Cara Bayaran : Jumlah Cas Kad Debit (%)
            Else
                rs!cas_kad_debit = Null
            End If
            If Frm87.TB39 <> vbNullString Then
                rs!jumlah_cas_kad_debit = Format(Frm87.TB39, "0.00") 'Cara Bayaran : Jumlah Cas Kad Debit (RM)
            Else
                rs!jumlah_cas_kad_debit = Null
            End If
            If Frm87.TB40 <> vbNullString Then
                rs!jumlah_potongan_kad_debit = Format(Frm87.TB40, "0.00") 'Cara Bayaran : Jumah Potongan Kad Debit (RM)
            Else
                rs!jumlah_potongan_kad_debit = Null
            End If
            If Frm87.TB32 <> vbNullString Then
                rs!jumlah = Format(Frm87.TB32, "0.00") 'Jumlah Harga Barang Tanpa GST (RM)
            Else
                rs!jumlah = Null
            End If
            If Frm87.TB13 <> vbNullString Then
                rs!jumlah_cukai_gst = Format(Frm87.TB13, "0.00") 'Jumlah Cukai GST (ZR + SR)
            Else
                rs!jumlah_cukai_gst = Null
            End If
            If Frm87.TB14 <> vbNullString Then
                rs!harga_barang_dengan_gst = Format(Frm87.TB14, "0.00") 'Jumlah Harga Barang Dengan GST (RM)
            Else
                rs!harga_barang_dengan_gst = Null
            End If
            If Frm87.TB17 <> vbNullString Then
                rs!adjustment = Format(Frm87.TB17, "0.00") 'Adjustment (RM)
            Else
                rs!adjustment = Null
            End If
            If Frm87.TB18 <> vbNullString Then
                rs!jumlah_bayaran = Format(Frm87.TB18, "0.00") 'Jumlah Harga Jualan (RM)
            Else
                rs!jumlah_bayaran = Null
            End If
            If Frm87.TB18 <> vbNullString Then
                rs!jumlah_bayaran = Format(Frm87.TB18, "0.00") 'Jumlah Harga Jualan (RM)
            Else
                rs!jumlah_bayaran = Null
            End If
            rs!flag_bayaran = 0 ' 0 : Pembeli Bayar , 1 : Kedai Bayar
            If Frm87.L25_Text <> vbNullString Then
                rs!gst_zr_harga = Format(Frm87.L25_Text, "0.00") 'Harga Keseluruhan Bagi Barang ZR
            Else
                rs!gst_zr_harga = Null
            End If
            If Frm87.L26_Text <> vbNullString Then
                rs!gst_zr_cukai = Format(Frm87.L26_Text, "0.00") 'Jumlah Cukai Bagi ZR
            Else
                rs!gst_zr_cukai = Null
            End If
            If Frm87.L23_Text <> vbNullString Then
                rs!gst_sr_harga = Format(Frm87.L23_Text, "0.00") 'Harga Keseluruhan Bagi Barang SR
            Else
                rs!gst_sr_harga = Null
            End If
            If Frm87.L24_Text <> vbNullString Then
                rs!gst_sr_cukai = Format(Frm87.L24_Text, "0.00") 'Jumlah Cukai Bagi SR
            Else
                rs!gst_sr_cukai = Null
            End If
            If Frm87.CB27 = 0 Then
                rs!gst_include = Null
            ElseIf Frm87.CB27 = 1 Then
                rs!gst_include = "**Harga Termasuk GST"
            End If
            If Frm87.TB42 <> vbNullString Then
                rs!harga_tanpa_gst = Format(Frm87.TB42, "0.00") 'Harga Keseluruhan Tanpa GST (RM)
            Else
                rs!harga_tanpa_gst = Null 'Harga Keseluruhan Tanpa GST (RM)
            End If
            If Frm87.CBB2 <> vbNullString Then
                Frm87_LM_EMP_NO = Split(Frm87.CBB2, "  |  ")(1)
                rs!no_pekerja = Frm87_LM_EMP_NO 'No. Pekerja
            End If
            rs!no_rujukan_pembeli = Frm87_LM_No_PEMBELI 'No. Rujukan Pembeli
            rs!kategori_pembeli = Frm87_LM_KATEGORI 'Kategori Pembeli
            rs!write_timestamp = Now
            
            rs.Update
            
            rs.Close
            Set rs = Nothing
'###Update Akaun Ansuran### - End

'### Update Senarai Ansuran ### - Start
'### Carian Jenis Ansuran ### - Start
    
            Frm87_FLAG_UPAH = 0
            Frm87_FLAG_ANSURAN = 0
            Frm87_FLAG_JELAS = 0
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 27_senarai_ansuran where ID='" & Frm87.L18_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!jenis_ansuran) Then
                    If rs!jenis_ansuran = 0 Then '0 : Harga Semasa , 1 : Harga Tetap
                        If Not IsNull(rs!UPAH) And Not IsNull(rs!JUMLAH_UPAH) Then
                            If Format(rs!UPAH, "0.00") = Format(rs!JUMLAH_UPAH, "0.00") Then
                                Frm87_FLAG_UPAH = 1
                            End If
                        End If
                        If Not IsNull(rs!berat_jualan) And Not IsNull(rs!JUMLAH_BERAT) Then
                            If Format(rs!berat_jualan, "0.00") = Format(rs!JUMLAH_BERAT, "0.00") Then
                                Frm87_FLAG_ANSURAN = 1
                            End If
                        End If
                        If Frm87_FLAG_UPAH = 1 And Frm87_FLAG_ANSURAN = 1 Then
                            Frm87_FLAG_JELAS = 1
                            rs!Status = "Jelas" 'Status
                            rs!tarikh_jelas = Frm87.DTPicker2 'Tarikh Jelas
                            rs.Update
                        Else
                            rs!Status = "Belum Jelas" 'Status
                            rs!tarikh_jelas = Null 'Tarikh Jelas
                            rs.Update
                        End If
                    ElseIf rs!jenis_ansuran = 1 Then
                        If Not IsNull(rs!harga_jualan) And Not IsNull(rs!jumlah_bayaran) Then
                            If Format(rs!harga_jualan, "0.00") = Format(rs!jumlah_bayaran, "0.00") Then
                                Frm87_FLAG_ANSURAN = 1
                            End If
                        End If
                        If Frm87_FLAG_ANSURAN = 1 Then
                            Frm87_FLAG_JELAS = 1
                            rs!Status = "Jelas" 'Status
                            rs!tarikh_jelas = Frm87.DTPicker2 'Tarikh Jelas
                            rs.Update
                        Else
                            rs!Status = "Belum Jelas" 'Status
                            rs!tarikh_jelas = Null 'Tarikh Jelas
                            rs.Update
                        End If
                    End If
                End If
            End If
            
            rs.Close
            Set rs = Nothing
'### Carian Jenis Ansuran ### - End
'### Update Senarai Ansuran ### - End

'### Update Database Utama Jika Sudah Terjual ### - Start
            If Frm87_FLAG_JELAS = 1 Then
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from Data_Database where no_siri_produk='" & Frm87_LM_No_SIRI & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Frm87_LM_JENIS_PRODUK = 0 Then
                        If Frm87_LM_POTONG = 0 Then
                            rs!StatusItem = 19
                            rs!beza_berat = "0.00"
                        ElseIf Frm87_LM_POTONG = 1 Then
                            rs!StatusItem = 20
                            rs!beza_berat = Format(rs!Berat - Frm87_LM_BERAT_POTONG, "0.00") 'Beza Berat
                        End If
                    ElseIf Frm87_LM_JENIS_PRODUK = 1 Then
                        rs!StatusItem = 19
                    End If
                    rs.Update
                End If
                
                rs.Close
                Set rs = Nothing
            End If
'### Update Database Utama Jika Sudah Terjual ### - End

'### Update No. Resit ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If rs!Default1 = "Default" Then
                    rs!no_resit_ansuran = Frm87.L12_Text + 1 'No. Resit Ansuran
                    rs.Update
                End If
            End If
            
            rs.Close
            Set rs = Nothing
            
'### Update Log ### - Start
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & user & "] Update Bayaran Ansuran. No. Ansuran [" & "ANS" & Format(Frm87_LM_No_RESIT_ANSURAN, "000000") & "]"
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
'### Update Log ### - End

'###Update Data Simpanan Duit Pelanggan### - Start
            If Format(Frm87.TB21, "0.00") <> "0.00" Then
            
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm87.L33_Text & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    Frm87_LM_JUMLAH_SIMPANAN = Frm87.L27_Text  'Jumlah Simpanan Yang Ada
                    Frm87_LM_GUNA_SIMPAN = Frm87.TB21  'Jumlah Simpanan Yang Hendak Digunakan
                    
                    rs!baki_simpanan = Format(Frm87_LM_JUMLAH_SIMPANAN - Frm87_LM_GUNA_SIMPAN, "0.00") 'Baki Simpanan
                    rs.Update
                End If
                
                rs.Close
                Set rs = Nothing
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 24_rekod_kewangan_pelanggan", cn, adOpenKeyset, adLockOptimistic
                
                rs.AddNew
                rs!tarikh = Frm87.DTPicker2 'Tarikh
                rs!jenis = 1 '0 : Simpanan , 1 : Penggunaan Duit
                rs!no_rujukan_pelanggan = Frm87.L33_Text 'No. Rujukan Pelanggan
                rs!no_resit = "ANS" & Format(Frm87_LM_No_RESIT_ANSURAN, "000000") 'No. Resit Ansuran
                rs!jumlah = Format(Frm87.TB21, "0.00") 'Jumlah Simpanan Yang Digunakan (RM)
                rs!jenis_penggunaan = 1 '0 : Belian Barangan Kemas , 1 : Ansuran , 2 : Tempahan (Deposit) , 3 : Servis , 4 : Tempahan (Ambilan Barang)
                rs!cawangan = G_CAWANGAN
                rs!Status = 1
                rs.Update
                
                rs.Close
                Set rs = Nothing
               
            End If
'###Update Data Simpanan Duit Pelanggan### - End

            Call Frm87_Initial_Setting
            
            If Frm87_FLAG_JELAS = 0 Then
            
                Note = "Data ansuran telah berjaya disimpan." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Cetak resit bayaran ansuran ?"
                
                Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
                
                If Answer = vbNo Then
                    Exit Sub
                End If
                If Answer = vbYes Then
                    G_No_RESIT_ANSURAN = "ANS" & Format(Frm87_LM_No_RESIT_ANSURAN, "000000") 'No. Resit Ansuran
                    Call Frm87_Resit_Ansuran
                End If
                
            ElseIf Frm87_FLAG_JELAS = 1 Then
            
                Note = "Data ansuran telah berjaya disimpan." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Pembeli ini telah menjelaskan semua bayaran bagi barang ini." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Status barang berubah kepada [JELAS]." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Cetak resit bayaran ansuran ?"
                
                Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
                
                If Answer = vbNo Then
                    Exit Sub
                End If
                If Answer = vbYes Then
                    G_No_RESIT_ANSURAN = "ANS" & Format(Frm87_LM_No_RESIT_ANSURAN, "000000") 'No. Resit Ansuran
                    Call Frm87_Resit_Ansuran
                End If
                
            End If
        End If
        
    End If
End If
End Sub
Private Sub Form_Load()
'on error resume next
GLOBAL_DISABLE = 0
Frm87.L29_Text = vbNullString
Frm87.L34_Text = 0
Frm87.L35_Text = "0.00 g"
End Sub
Private Sub Frm87_LM_resit_ansuran_Click()
'on error resume next
Frm87_LM_ID = vbNullString
DATA_FOUND = 0

If Frm87.MSFlexGrid2 <> vbNullString Then
    Frm87_LM_ID = Frm87.MSFlexGrid2.TextMatrix(Frm87.MSFlexGrid2, 2) 'No. ID
    
    If Frm87_LM_ID <> vbNullString Then
        Note = "Cetak invoice bayaran ansuran ini ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
    
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 28_rekod_ansuran where ID='" & Frm87_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!no_resit_ansuran) Then
                    G_No_RESIT_ANSURAN = rs!no_resit_ansuran 'No. Resit Ansuran
                    Call Frm87_Resit_Ansuran
                End If
            End If
            
            'rs.Close
            Set rs = Nothing
            
        End If
    End If
End If
End Sub
Private Sub Frm87_SM_Edit_Click()
'on error resume next
DATA_FOUND = 0
Frm87_LM_No_PEKERJA = vbNullString

If Frm87.MSFlexGrid1 <> vbNullString Then
    Frm87_LM_ID = Frm87.MSFlexGrid1.TextMatrix(Frm87.MSFlexGrid1, 2) 'No. ID
    
    If Frm87_LM_ID <> vbNullString Then
        Call Frm87_Initial_Setting
        Unload Frm26
        Unload Frm27
        Unload Frm28

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 27_senarai_ansuran where ID='" & Frm87_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!Status) Then
                If rs!Status = "Jelas" Then
                    rs.Close
                    Set rs = Nothing
                    
                    GoTo End_Task:
                End If
            End If
            
            If Not IsNull(rs!no_rujukan) Then Frm87.L11_Text = rs!no_rujukan 'No. Rujukan Ansuran
            
            If Not IsNull(rs!jumlah_bayaran) Then Frm87.L36_Text = rs!jumlah_bayaran 'Jumlah Bayaran Yang Telah Dibuat
            If Not IsNull(rs!JUMLAH_UPAH) Then Frm87.L37_Text = rs!JUMLAH_UPAH 'Jumlah Bayaran Upah Yang Telah Dibuat
            If Not IsNull(rs!JUMLAH_BERAT) Then Frm87.L38_Text = rs!JUMLAH_BERAT 'Jumlah Bayaran Berat Yang Telah Dibuat
            
            If Not IsNull(rs!no_rujukan_pelanggan) Then
                Call Frm28_initial
                
                If Not IsNull(rs!no_rujukan_pelanggan) Then Frm28.L5_Text = rs!no_rujukan_pelanggan 'No. Rujukan Pembeli
                If Not IsNull(rs!Nama) Then Frm28.L1_Text = rs!Nama 'Maklumat Pembeli : Nama
                If Not IsNull(rs!no_ic) Then Frm28.L2_Text = rs!no_ic 'Maklumat Pembeli : No. Kad Pengenalan
                If Not IsNull(rs!no_tel) Then Frm28.L3_Text = rs!no_tel 'Maklumat Pembeli : No. Telefon
                
            Else
                Call Frm26_initial
                
                If Not IsNull(rs!Nama) Then Frm26.TB1 = rs!Nama 'Maklumat Pembeli : Nama
                If Not IsNull(rs!no_tel) Then Frm26.TB2 = rs!no_tel 'Maklumat Pembeli : No. Telefon
                
            End If
            
            If Not IsNull(rs!no_siri_Produk) Then Frm87.TB2 = rs!no_siri_Produk 'No. Siri Produk
            If Not IsNull(rs!jenis_produk) Then Frm87.L13_Text = rs!jenis_produk 'Flat Kategori Produk , 0 : BK , 1 : Permata
            If Not IsNull(rs!kategori_Produk) Then Frm87.L10_Text = rs!kategori_Produk 'Kategori Produk
            If Not IsNull(rs!Berat_Asal) Then Frm87.TB3 = rs!Berat_Asal 'Berat Asal (g)
            If Not IsNull(rs!berat_jualan) Then Frm87.TB4 = rs!berat_jualan 'Berat Jualan (g)
            If Not IsNull(rs!harga_Semasa) Then Frm87.TB5 = rs!harga_Semasa 'Harga Emas Semasa Masa Tempahan Dibuat
            If Not IsNull(rs!UPAH) Then Frm87.TB6 = rs!UPAH 'Upah
            If Not IsNull(rs!harga_asal) Then Frm87.TB7 = rs!harga_asal 'Harga Asal Jualan
            If Not IsNull(rs!adjustment) Then Frm87.TB9 = rs!adjustment 'Adjustment
            If Not IsNull(rs!harga_jualan) Then Frm87.TB10 = rs!harga_jualan 'Harga Jualan (RM)
            If Not IsNull(rs!jenis_ansuran) Then
                If rs!jenis_ansuran = 0 Then
                    Frm87.CB14 = 1
                    Frm87.CB15 = 0
                ElseIf rs!jenis_ansuran = 1 Then
                    Frm87.CB15 = 1
                    Frm87.CB14 = 0
                End If
            End If
            If Not IsNull(rs!tarikh) Then Frm87.DTPicker1 = rs!tarikh 'Tarikh Tempahan
            If Not IsNull(rs!no_rujukan_pekerja) Then Frm87_LM_No_PEKERJA = rs!no_rujukan_pekerja 'No. Pekerja
            DATA_FOUND = 1
        End If
        
        rs.Close
        Set rs = Nothing
        
'### Carian Maklumat Penjual (Data Pekerja) ### - Start
        If Frm87_LM_No_PEKERJA <> vbNullString Then
            DATA_PEKERJA_FOUND = 0
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from employee where NoPekerja='" & Frm87_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                Frm87_LM_MAKLUMAT_PEKERJA = rs!Samaran & "  |  " & rs!NoPekerja
                DATA_PEKERJA_FOUND = 1
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_PEKERJA_FOUND = 1 Then
                On Error GoTo Err_A:
                Frm87.CBB1 = Frm87_LM_MAKLUMAT_PEKERJA
Restore_A:
            End If
        End If
'### Carian Maklumat Penjual (Data Pekerja) ### - End

        If DATA_FOUND = 1 Then
            
            Frm87.TB8.Locked = True
            Frm87.TB8.BackColor = &H8000000A
            
            Frm87.CB14.Enabled = False
            Frm87.CB15.Enabled = False
            
            Frm87.CMD6.Visible = False
            Frm87.CMD16.Visible = True
            Frm87.CMD17.Visible = True

            Frm87.Pic2.Visible = True
            Frm87.Pic4.Visible = False
        End If
    End If
End If

Exit Sub
Err_A:
Frm87.CBB1.AddItem Frm87_LM_MAKLUMAT_PEKERJA
Frm87.CBB1 = Frm87_LM_MAKLUMAT_PEKERJA
Resume Restore_A:

Exit Sub
End_Task:

Frm87.Pic4.Visible = True

MsgBox "Status bagi belian item ini adalah [JELAS]" & vbCrLf & _
        vbNullString & vbCrLf & _
        "Anda tidak dibenarkan untuk edit data ini."
End Sub
Private Sub Frm87_SM_Edit2_Click()
'on error resume next
Dim Frm87_LM_JUMLAH_BAYARAN_ASAL As Double
Dim Frm87_LM_BAKI_BERAT_ASAL As Double
Dim Frm87_LM_BAKI_UPAH_ASAL As Double
Dim Frm87_LM_BAKI_BAYARAN_ASAL As Double
Dim Frm87_LM_BERAT_DIPEROLEHI As Double
Dim Frm87_LM_JUMLAH_BAYARAN_UPAH As Double
Dim Frm87_LM_JUMLAH_BAYARAN_ANSURAN As Double
Dim Frm87_LM_SIMPANAN_DIGUNAKAN As Double
Dim Frm87_LM_SIMPANAN_ASAL As Double

Frm87_LM_JUMLAH_BAYARAN_ASAL = 0
Frm87_LM_BAKI_BERAT_ASAL = 0
Frm87_LM_BAKI_UPAH_ASAL = 0
Frm87_LM_BAKI_BAYARAN_ASAL = 0
Frm87_LM_BERAT_DIPEROLEHI = 0
Frm87_LM_JUMLAH_BAYARAN_UPAH = 0
Frm87_LM_JUMLAH_BAYARAN_ANSURAN = 0
Frm87_LM_SIMPANAN_DIGUNAKAN = 0 'Jumlah Simpanan Yang Digunakan (RM)
Frm87_LM_SIMPANAN_ASAL = 0

Frm87_LM_Flag_SIMPANAN = 0 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai

Frm87_LM_ID = vbNullString
Frm87_LM_No_PEKERJA = vbNullString
Frm87_LM_No_PEMBELI = vbNullString
DATA_FOUND = 0

If Frm87.MSFlexGrid2 <> vbNullString Then
    Frm87_LM_ID = Frm87.MSFlexGrid2.TextMatrix(Frm87.MSFlexGrid2, 2) 'No. ID
    
    If Frm87_LM_ID <> vbNullString Then
        Note = "Lihat atau edit data ini ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
    
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 28_rekod_ansuran where ID='" & Frm87_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!no_resit_ansuran) Then
                    Frm87_LM_No_RESIT = rs!no_resit_ansuran 'No. Resit Ansuran
                    
                    DATA_FOUND = 1
                End If
                If Not IsNull(rs!id_database_reg) Then Frm87_LM_ID_ASAL = rs!id_database_reg 'No. ID Dari Database Senarai Pembeli Ansuran
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_FOUND = 1 Then
            
                Call Frm87_Initial_Setting
                Frm87.L12_Text = Frm87_LM_No_RESIT 'No. Resit Ansuran
                Frm87.L18_Text = Frm87_LM_ID_ASAL 'No. ID Dari Database Senarai Pembeli Ansuran
                
'### Carian Jenis Ansuran ### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 27_senarai_ansuran where ID='" & Frm87.L18_Text & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Not IsNull(rs!jenis_ansuran) Then
                        If rs!jenis_ansuran = 0 Then '0 : Harga Semasa , 1 : Harga Tetap
                            Frm87_LM_JENIS = 0
                            
                            If Not IsNull(rs!jumlah_bayaran) Then Frm87_LM_JUMLAH_BAYARAN_ASAL = rs!jumlah_bayaran 'Jumlah Bayaran Terkumpul Asal (RM)
                            If Not IsNull(rs!BAKI_BERAT) Then Frm87_LM_BAKI_BERAT_ASAL = rs!BAKI_BERAT 'Baki Berat Asal (g)
                            If Not IsNull(rs!baki_upah) Then Frm87_LM_BAKI_UPAH_ASAL = rs!baki_upah 'Baki Upah Asal (RM)
                        ElseIf rs!jenis_ansuran = 1 Then
                            Frm87_LM_JENIS = 1
                            
                            If Not IsNull(rs!baki_bayaran) Then Frm87_LM_BAKI_BAYARAN_ASAL = rs!baki_bayaran 'Baki Bayaran Asal (RM)
                        End If
                    End If
                    If Not IsNull(no_rujukan_pelanggan) Then Frm87_LM_No_PEMBELI = rs!no_rujukan_pelanggan 'No. Rujukan Pembeli
                End If
                
                rs.Close
                Set rs = Nothing
'### Carian Jenis Ansuran ### - End

'###Update Bayaran Ansuran### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 28_rekod_ansuran where no_resit_ansuran='" & Frm87.L12_Text & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Not IsNull(rs!id_database_reg) Then Frm87.L18_Text = rs!id_database_reg 'No. ID Dari Database Senarai Pembeli Ansuran
                    If Not IsNull(rs!FLAG_ANSURAN) Then 'Flag samada ada bayaran ansuran atau tidak , 0 : Tiada bayaran ansuran , 1 : Ada bayaran ansuran
                        If rs!FLAG_ANSURAN = 1 Then
                            Frm87.CB20 = 1
                            If Not IsNull(rs!jumlah_ansuran) Then Frm87.TB12 = rs!jumlah_ansuran 'Jumlah Bayran Ansuran
'                            If Not IsNull(rs!flag_ansuran_zr) Then 'Flag samada ada zr bagi bayaran ansuran , 0 : Tiada ZR , 1 : Ada ZR
'                                If rs!flag_ansuran_zr = 0 Then
'                                    Frm87.CB22 = 0
'                                ElseIf rs!flag_ansuran_zr = 1 Then
'                                    Frm87.CB22 = 1
'                                End If
'                            Else
'                                Frm87.CB22 = 0
'                            End If
'                            If Not IsNull(rs!flag_ansuran_sr) Then 'Flag samada ada sr bagi bayaran ansuran , 0 : Tiada SR , 1 : Ada SR
'                                If rs!flag_ansuran_sr = 0 Then
'                                    Frm87.CB23 = 0
'                                ElseIf rs!flag_ansuran_sr = 1 Then
'                                    Frm87.CB23 = 1
'                                End If
'                            Else
'                                Frm87.CB23 = 0
'                            End If
'                            If Not IsNull(rs!ansuran_gst) Then 'Jumlah GST bagi bayaran ansuran (RM)
'                                Frm87.L21_Text = rs!ansuran_gst
'                            Else
'                                Frm87.L21_Text = "0.00"
'                            End If
                            If Not IsNull(rs!harga_Semasa) Then 'Harga Semasa
                                Frm87.TB15 = rs!harga_Semasa
                                Frm87.Pic6.Visible = True
                            Else
                                Frm87.Pic6.Visible = False
                            End If
                            If Not IsNull(rs!berat_diperoleh) Then 'Berat Diperolehi
                                Frm87.TB16 = rs!berat_diperoleh
                                If IsNumeric(rs!berat_diperoleh) Then Frm87_LM_BERAT_DIPEROLEHI = rs!berat_diperoleh 'Berat DiPerolehi
                                Frm87.L19_Text = Format(Frm87_LM_BERAT_DIPEROLEHI + Frm87_LM_BAKI_BERAT_ASAL, "#,##0.00") 'Baki Berat
                            Else
                                If IsNumeric(rs!jumlah_ansuran) Then Frm87_LM_JUMLAH_BAYARAN_ANSURAN = rs!jumlah_ansuran 'Jumlah Bayran Ansuran
                                Frm87.L28_Text = Format(Frm87_LM_JUMLAH_BAYARAN_ANSURAN + Frm87_LM_BAKI_BAYARAN_ASAL, "#,##0.00") 'Baki Bayaran
                            End If
                        ElseIf rs!FLAG_ANSURAN = 0 Then
                            Frm87.CB20 = 0
                            Frm87.L21_Text = "0.00"
                        End If
                    End If
                    
                    If Not IsNull(rs!flag_ansuran_zr) Then 'Flag samada ada zr bagi bayaran ansuran , 0 : Tiada ZR , 1 : Ada ZR
                        If rs!flag_ansuran_zr = 0 Then
                            Frm87.CB22 = 0
                        ElseIf rs!flag_ansuran_zr = 1 Then
                            Frm87.CB22 = 1
                        End If
                    Else
                        Frm87.CB22 = 0
                    End If
                    
                    If Not IsNull(rs!flag_ansuran_sr) Then 'Flag samada ada sr bagi bayaran ansuran , 0 : Tiada SR , 1 : Ada SR
                        If rs!flag_ansuran_sr = 0 Then
                            Frm87.CB23 = 0
                        ElseIf rs!flag_ansuran_sr = 1 Then
                            Frm87.CB23 = 1
                        End If
                    Else
                        Frm87.CB23 = 0
                    End If
                    If Not IsNull(rs!ansuran_gst) Then 'Jumlah GST bagi bayaran ansuran (RM)
                        Frm87.L21_Text = rs!ansuran_gst
                    Else
                        Frm87.L21_Text = "0.00"
                    End If
                
                    If Not IsNull(rs!flag_upah) Then
                        If rs!flag_upah = 1 Then 'Flag samada ada bayaran upah atau tidak , 0 : Tiada Bayaran Upah , 1 : Ada Bayaran Upah
                            Frm87.CB21 = 1
                            If Not IsNull(rs!JUMLAH_UPAH) Then
                                Frm87.TB19 = rs!JUMLAH_UPAH 'Jumlah Bayran Upah
                                
                                If IsNumeric(rs!JUMLAH_UPAH) Then Frm87_LM_JUMLAH_BAYARAN_UPAH = rs!JUMLAH_UPAH
                                Frm87.L20_Text = Format(Frm87_LM_JUMLAH_BAYARAN_UPAH + Frm87_LM_BAKI_UPAH_ASAL, "#,##0.00") 'Baki Upah
                            End If
'                            If Not IsNull(rs!flag_upah_zr) Then 'Flag samada ada zr bagi bayaran upah , 0 : Tiada ZR , 1 : Ada ZR
'                                If rs!flag_upah_zr = 1 Then
'                                    Frm87.CB18 = 1
'                                ElseIf rs!flag_upah_zr = 0 Then
'                                    Frm87.CB18 = 0
'                                End If
'                            Else
'                                Frm87.CB18 = 0
'                            End If
'                            If Not IsNull(rs!flag_upah_sr) Then 'Flag samada ada sr bagi bayaran upah , 0 : Tiada SR , 1 : Ada SR
'                                If rs!flag_upah_sr = 1 Then
'                                    Frm87.CB19 = 1
'                                ElseIf rs!flag_upah_sr = 0 Then
'                                    Frm87.CB19 = 0
'                                End If
'                            Else
'                                Frm87.CB19 = 0
'                            End If
'                            If Not IsNull(rs!upah_gst) Then 'Jumlah GST Bagi Upah (RM)
'                                Frm87.L22_Text = rs!upah_gst
'                            Else
'                                Frm87.L22_Text = "0.00"
'                            End If
                        Else
                            Frm87.CB21 = 0
                            Frm87.L22_Text = "0.00"
                        End If
                    End If
                
                    If Not IsNull(rs!jumlah_bayaran) Then Frm87.TB20 = rs!jumlah_bayaran 'Jumlah Ansuran + Jumlah Upah
                    If Not IsNull(rs!kadar_gst) Then Frm87.L17_Text = rs!kadar_gst 'Kadar GST (%)
                    If Not IsNull(rs!jumlah_gst) Then Frm87.TB13 = rs!jumlah_gst 'Jumlah GST
                    If Not IsNull(rs!jumlah_asal) Then Frm87.TB14 = rs!jumlah_asal 'Jumlah Asal (Ansuran + Upah) + GST
                    If Not IsNull(rs!adjustment) Then Frm87.TB17 = rs!adjustment 'Adjustment
                    If Not IsNull(rs!jumlah_keseluruhan) Then Frm87.TB18 = rs!jumlah_keseluruhan 'Jumlah bayaran selepas adjustment
                    If Not IsNull(rs!tarikh) Then Frm87.DTPicker2 = rs!tarikh 'Tarikh Bayaran
                    If Not IsNull(rs!no_rujukan_pekerja) Then Frm87_LM_No_PEKERJA = rs!no_rujukan_pekerja 'No. Pekerja
                End If
                
                rs.Close
                Set rs = Nothing
'###Update Bayaran Ansuran### - End

'### Carian Maklumat Penjual (Data Pekerja) ### - Start
                If Frm87_LM_No_PEKERJA <> vbNullString Then
                    DATA_PEKERJA_FOUND = 0
                    
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from employee where NoPekerja='" & Frm87_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                        Frm86_LM_MAKLUMAT_PEKERJA = rs!Samaran & "  |  " & rs!NoPekerja
                        DATA_PEKERJA_FOUND = 1
                    End If
                    
                    rs.Close
                    Set rs = Nothing
                    
                    If DATA_PEKERJA_FOUND = 1 Then
                        On Error GoTo Err_A:
                        Frm87.CBB2 = Frm86_LM_MAKLUMAT_PEKERJA
Restore_A:
                    End If
                End If
'### Carian Maklumat Penjual (Data Pekerja) ### - End

'###Update Akaun Ansuran### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 29_akaun_ansuran where no_resit='" & Frm87.L12_Text & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Not IsNull(rs!tunai) Then 'Cara Bayaran : Tunai
                        Frm87.TB27 = rs!tunai
                    Else
                        Frm87.TB27 = "0.00"
                    End If
                    If Not IsNull(rs!bank_in) Then 'Cara Bayaran : Bank In
                        Frm87.TB28 = rs!bank_in
                    Else
                        Frm87.TB28 = "0.00"
                    End If
                    If Not IsNull(rs!kad_kredit) Then 'Cara Bayaran : Kad Kredit
                        Frm87.TB29 = rs!kad_kredit
                    Else
                        Frm87.TB29 = "0.00"
                    End If
                    If Not IsNull(rs!cas_Kad_Kredit) Then 'Cara Bayaran : Cas Kad Kredit (%)
                        Frm87.L31_Text = rs!cas_Kad_Kredit
                    Else
                        Frm87.L31_Text = 0
                    End If
                    If Not IsNull(rs!jumlah_cas_kad_kredit) Then 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
                        Frm87.TB30 = rs!jumlah_cas_kad_kredit
                    Else
                        Frm87.TB30 = "0.00"
                    End If
                    If Not IsNull(rs!jumlah_potongan_kad_kredit) Then 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
                        Frm87.TB31 = rs!jumlah_potongan_kad_kredit
                    Else
                        Frm87.TB31 = "0.00"
                    End If
                    If Not IsNull(rs!duit_simpanan_kedai) Then 'Cara Bayaran : Simpanan Duit Di Kedai
                        Frm87.TB21 = rs!duit_simpanan_kedai
                        
                        
                        If rs!duit_simpanan_kedai <> "0.00" Then
                            Frm87_LM_Flag_SIMPANAN = 1 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
                            Frm87_LM_SIMPANAN_DIGUNAKAN = rs!duit_simpanan_kedai 'Jumlah Simpanan Yang Digunakan (RM)
                        End If
                        
                        
                    Else
                        Frm87.TB21 = "0.00"
                    End If
                    If Not IsNull(rs!kad_debit) Then 'Cara Bayaran : Kad Debit
                        Frm87.TB38 = rs!kad_debit
                    Else
                        Frm87.TB38 = "0.00"
                    End If
                    If Not IsNull(rs!cas_kad_debit) Then 'Cara Bayaran : Jumlah Cas Kad Debit (%)
                        Frm87.L32_Text = rs!cas_kad_debit
                    Else
                        Frm87.L32_Text = 0
                    End If
                    If Not IsNull(rs!jumlah_cas_kad_debit) Then 'Cara Bayaran : Jumlah Cas Kad Debit (RM)
                        Frm87.TB39 = rs!jumlah_cas_kad_debit
                    Else
                        Frm87.TB39 = "0.00"
                    End If
                    If Not IsNull(rs!jumlah_potongan_kad_debit) Then 'Cara Bayaran : Jumah Potongan Kad Debit (RM)
                        Frm87.TB40 = rs!jumlah_potongan_kad_debit
                    Else
                        Frm87.TB40 = "0.00"
                    End If
                    If Not IsNull(rs!jumlah) Then 'Jumlah Harga Barang Tanpa GST (RM)
                        Frm87.TB32 = rs!jumlah
                    Else
                        Frm87.TB32 = "0.00"
                    End If
                    If Not IsNull(rs!jumlah_cukai_gst) Then 'Jumlah Cukai GST (ZR + SR)
                        Frm87.TB13 = rs!jumlah_cukai_gst
                    Else
                        Frm87.TB13 = "0.00"
                    End If
                    If Not IsNull(rs!harga_barang_dengan_gst) Then 'Jumlah Harga Barang Dengan GST (RM)
                        Frm87.TB14 = rs!harga_barang_dengan_gst
                    Else
                        Frm87.TB14 = "0.00"
                    End If
                    If Not IsNull(rs!adjustment) Then 'Adjustment (RM)
                        Frm87.TB17 = rs!adjustment
                    Else
                        Frm87.TB17 = "0.00"
                    End If
                    If Not IsNull(rs!jumlah_bayaran) Then 'Jumlah Harga Jualan (RM)
                        Frm87.TB18 = rs!jumlah_bayaran
                    Else
                        Frm87.TB18 = "0.00"
                    End If
                    If Not IsNull(rs!gst_zr_harga) Then 'Harga Keseluruhan Bagi Barang ZR
                        Frm87.L25_Text = rs!gst_zr_harga
                    Else
                        Frm87.L25_Text = "0.00"
                    End If
                    If Not IsNull(rs!gst_zr_cukai) Then 'Jumlah Cukai Bagi ZR
                        Frm87.L26_Text = rs!gst_zr_cukai
                    Else
                        Frm87.L26_Text = "0.00"
                    End If
                    If Not IsNull(rs!gst_sr_harga) Then 'Harga Keseluruhan Bagi Barang SR
                        Frm87.L23_Text = rs!gst_sr_harga
                    Else
                        Frm87.L23_Text = "0.00"
                    End If
                    If Not IsNull(rs!gst_sr_cukai) Then 'Jumlah Cukai Bagi SR
                        Frm87.L24_Text = rs!gst_sr_cukai
                    Else
                        Frm87.L24_Text = "0.00"
                    End If
                    If Not IsNull(rs!gst_include) Then
                        If rs!gst_include = "**Harga Termasuk GST" Then
                            Frm87.CB27 = 1
                        Else
                            Frm87.CB27 = 0
                        End If
                    Else
                        Frm87.CB27 = 0
                    End If
                    If Not IsNull(rs!harga_tanpa_gst) Then Frm87.TB42 = rs!harga_tanpa_gst 'Jumlah Tanpa GST
                End If
                
                rs.Close
                Set rs = Nothing
'###Update Akaun Ansuran### - End

'###Update Data Simpanan Duit Pelanggan### - Start
                If Frm87_LM_No_PEMBELI <> vbNullString Then
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm87_LM_No_PEMBELI & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                        If Not IsNull(rs!baki_simpanan) Then
                            Frm87.L27_Text = Format(rs!baki_simpanan, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
                            If IsNumeric(rs!baki_simpanan) Then
                                Frm87_LM_SIMPANAN_ASAL = rs!baki_simpanan 'Jumlah Simpanan Asal Yang Ada (RM)
                                
                                Frm87.L27_Text = Format(Frm87_LM_SIMPANAN_ASAL + Frm87_LM_SIMPANAN_DIGUNAKAN, "#,##0.00") 'Baki Simpanan Pelanggan Ini (RM)
                            End If
                        End If
                    End If
                    
                    rs.Close
                    Set rs = Nothing
                End If
'###Update Data Simpanan Duit Pelanggan### - End
                
                Frm87.CMD9.Visible = False
                Frm87.CMD10.Visible = False
                Frm87.CMD13.Visible = True
                Frm87.CMD14.Visible = True
                
                Frm87.Pic4.Visible = False
                Frm87.Pic5.Visible = True
                
            End If
        End If
    End If
End If

Exit Sub
Err_A:
Frm87.CBB2.AddItem Frm87_LM_MAKLUMAT_PEKERJA
Frm87.CBB2 = Frm87_LM_MAKLUMAT_PEKERJA
Resume Restore_A:
End Sub
Private Sub Frm87_SM_Exccel_Click()
'On Error Resume Next
Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh
        .Columns("C").ColumnWidth = 20 'Jenis Ansuran
        .Columns("D").ColumnWidth = 20 'Status
        .Columns("E").ColumnWidth = 50 'Nama
        .Columns("F").ColumnWidth = 20 'No. Kad Pengenalan
        .Columns("G").ColumnWidth = 20 'No. Telefon
        .Columns("H").ColumnWidth = 20 'No. Siri
        .Columns("I").ColumnWidth = 35 'Kategori Produk
        .Columns("J").ColumnWidth = 20 'Berat Asal (g)
        .Columns("K").ColumnWidth = 20 'Berat Jualan (g)
        .Columns("L").ColumnWidth = 20 'Harga Semasa (RM/g)
        .Columns("M").ColumnWidth = 20 'Upah (RM)
        .Columns("N").ColumnWidth = 20 'Harga Asal (RM)
        .Columns("O").ColumnWidth = 20 'Adjustment (RM)
        .Columns("P").ColumnWidth = 20 'Harga Jualan (RM)
        .Columns("Q").ColumnWidth = 20 'Kategori Pembeli
        .Columns("R").ColumnWidth = 20 '
    
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Then
            
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        .Cells(7, 1) = "Senarai Pembeli Ansuran." 'Header Report
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh"
        .Cells(8, 3) = "Jenis Ansuran"
        .Cells(8, 4) = "Status"
        .Cells(8, 5) = "Nama"
        .Cells(8, 6) = "No. Kad Pengenalan"
        .Cells(8, 7) = "No. Telefon"
        .Cells(8, 8) = "No. Siri"
        .Cells(8, 9) = "Kategori Produk"
        .Cells(8, 10) = "Berat Asal (g)"
        .Cells(8, 11) = "Berat Jualan (g)"
        .Cells(8, 12) = "Harga Semasa (RM/g)"
        .Cells(8, 13) = "Upah (RM)"
        .Cells(8, 14) = "Harga Asal (RM)"
        .Cells(8, 15) = "Adjustment (RM)"
        .Cells(8, 16) = "Harga Jualan (RM)"
        .Cells(8, 17) = "Kategori Pembeli"
        
        For i = 1 To 17
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
        
        Y = 0
        For x = 1 To Frm87.MSFlexGrid1.Rows - 1
            Y = Y + 1
            .Cells(8 + Y, 1) = Y 'No.
            .Cells(8 + Y, 1).HorizontalAlignment = xlCenter
            .Cells(8 + Y, 2) = "'" & Frm87.MSFlexGrid1.TextMatrix(x, 3) 'Tarikh
            .Cells(8 + Y, 2).HorizontalAlignment = xlCenter
            .Cells(8 + Y, 3) = Frm87.MSFlexGrid1.TextMatrix(x, 4) 'Jenis Ansuran
            .Cells(8 + Y, 3).HorizontalAlignment = xlCenter
            .Cells(8 + Y, 4) = Frm87.MSFlexGrid1.TextMatrix(x, 5) 'Status
            .Cells(8 + Y, 4).HorizontalAlignment = xlCenter
            .Cells(8 + Y, 5) = Frm87.MSFlexGrid1.TextMatrix(x, 6) 'Nama
            .Cells(8 + Y, 6) = "'" & Frm87.MSFlexGrid1.TextMatrix(x, 7) 'No. Kad Pengenalan
            .Cells(8 + Y, 7) = "'" & Frm87.MSFlexGrid1.TextMatrix(x, 8) 'No. Telefon
            .Cells(8 + Y, 8) = Frm87.MSFlexGrid1.TextMatrix(x, 9) 'No. Siri
            .Cells(8 + Y, 8).HorizontalAlignment = xlCenter
            .Cells(8 + Y, 9) = Frm87.MSFlexGrid1.TextMatrix(x, 10) 'Kategori Produk
            .Cells(8 + Y, 10).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 10) = Frm87.MSFlexGrid1.TextMatrix(x, 11) 'Berat Asal (g)
            .Cells(8 + Y, 10).HorizontalAlignment = xlCenter
            .Cells(8 + Y, 11).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 11) = Frm87.MSFlexGrid1.TextMatrix(x, 12) 'Berat Jualan (g)
            .Cells(8 + Y, 11).HorizontalAlignment = xlCenter
            .Cells(8 + Y, 12).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 12) = Frm87.MSFlexGrid1.TextMatrix(x, 13) 'Harga Semasa (RM/g)
            .Cells(8 + Y, 12).HorizontalAlignment = xlCenter
            .Cells(8 + Y, 13).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 13) = Frm87.MSFlexGrid1.TextMatrix(x, 14) 'Upah (RM)
            .Cells(8 + Y, 13).HorizontalAlignment = xlCenter
            .Cells(8 + Y, 14).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 14) = Frm87.MSFlexGrid1.TextMatrix(x, 15) 'Harga Asal (RM)
            .Cells(8 + Y, 14).HorizontalAlignment = xlCenter
            .Cells(8 + Y, 15).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 15) = Frm87.MSFlexGrid1.TextMatrix(x, 16) 'Adjustment (RM)
            .Cells(8 + Y, 15).HorizontalAlignment = xlCenter
            .Cells(8 + Y, 16).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 16) = Frm87.MSFlexGrid1.TextMatrix(x, 17) 'Harga Jualan (RM)
            .Cells(8 + Y, 16).HorizontalAlignment = xlCenter
            .Cells(8 + Y, 17) = Frm87.MSFlexGrid1.TextMatrix(x, 18) 'Kategori Pembeli

            For Col = 1 To 17
                .Cells(8 + Y, Col).Borders.LineStyle = xlContinuous
            Next Col
        Next x
    
        Y = Y + 2
        .Cells(8 + Y, 1).Font.Bold = True
        .Cells(8 + Y, 1) = "Report Generated By Sankyu System [ +6010 - 900 4788 , sankyusystem@gmail.com ]" 'Watermark Sankyu System
    End With
    
    ' This makes Excel visible
    xlObject.Visible = True
    xlObject.EnableEvents = True
End If
End Sub
Private Sub Frm87_SM_Padam_Click()
'on error resume next
Dim Frm87_LM_JUMLAH_BAYARAN_ASAL As Double
Dim Frm87_LM_BAKI_BERAT_ASAL As Double
Dim Frm87_LM_BAKI_UPAH_ASAL As Double
Dim Frm87_LM_BAKI_BAYARAN_ASAL As Double
Dim Frm87_LM_BERAT_DIPEROLEHI As Double
Dim Frm87_LM_JUMLAH_BAYARAN_UPAH As Double
Dim Frm87_LM_JUMLAH_BAYARAN_ANSURAN As Double
Dim Frm87_LM_JUMLAH_BERAT_ASAL As Double
Dim Frm87_LM_JUMLAH_UPAH_ASAL As Double
Dim Frm87_LM_SIMPANAN_DIGUNAKAN As Double
Dim Frm87_LM_SIMPANAN_ASAL As Double

Frm87_LM_SIMPANAN_ASAL = 0
Frm87_LM_SIMPANAN_DIGUNAKAN = 0
Frm87_LM_JUMLAH_BAYARAN_ASAL = 0
Frm87_LM_BAKI_BERAT_ASAL = 0
Frm87_LM_BAKI_UPAH_ASAL = 0
Frm87_LM_BAKI_BAYARAN_ASAL = 0
Frm87_LM_BERAT_DIPEROLEHI = 0
Frm87_LM_JUMLAH_BAYARAN_UPAH = 0
Frm87_LM_JUMLAH_BAYARAN_ANSURAN = 0
Frm87_LM_JUMLAH_BERAT_ASAL = 0
Frm87_LM_JUMLAH_UPAH_ASAL = 0

Frm87_LM_ID = vbNullString
Frm87_LM_No_PEKERJA = vbNullString
DATA_FOUND = 0

If Frm87.MSFlexGrid2 <> vbNullString Then
    Frm87_LM_ID = Frm87.MSFlexGrid2.TextMatrix(Frm87.MSFlexGrid2, 2) 'No. ID
    
    If Frm87_LM_ID <> vbNullString Then
        Note = "Padam resit ini ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Jika anda padam resit ini , rekod bayaran ansuran bagi resit ini juga akan dipadamkan." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
    
'###Padam Bayaran Ansuran### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 28_rekod_ansuran where ID='" & Frm87_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!id_database_reg) Then Frm87_LM_ID_ASAL = rs!id_database_reg 'ID Asal
                If Not IsNull(rs!no_resit_ansuran) Then Frm87_LM_No_RESIT = rs!no_resit_ansuran 'No. Resit Ansuran
                If Not IsNull(rs!id_database_reg) Then Frm87_LM_ID_ASAL = rs!id_database_reg 'No. ID Dari Database Senarai Pembeli Ansuran
                If Not IsNull(rs!jumlah_ansuran) Then
                    If IsNumeric(rs!jumlah_ansuran) Then Frm87_LM_JUMLAH_BAYARAN_ANSURAN = rs!jumlah_ansuran 'Jumlah Bayran Ansuran
                End If
                If Not IsNull(rs!berat_diperoleh) Then 'Berat Diperolehi
                    If IsNumeric(rs!berat_diperoleh) Then Frm87_LM_BERAT_DIPEROLEHI = rs!berat_diperoleh 'Berat DiPerolehi
                End If
                If Not IsNull(rs!JUMLAH_UPAH) Then
                    If IsNumeric(rs!JUMLAH_UPAH) Then Frm87_LM_JUMLAH_BAYARAN_UPAH = rs!JUMLAH_UPAH
                End If
                
                DATA_FOUND = 1
                
                rs.Delete
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
'###Padam Bayaran Ansuran### - End
            
            If DATA_FOUND = 1 Then
                
'### Update Senarai Ansuran ### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 27_senarai_ansuran where ID='" & Frm87_LM_ID_ASAL & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Not IsNull(rs!no_rujukan_pelanggan) Then Frm87_LM_No_CUST = rs!no_rujukan_pelanggan 'No. Rujukan Pelanggan
                    If Not IsNull(rs!jenis_ansuran) Then
                        If rs!jenis_ansuran = 0 Then '0 : Harga Semasa , 1 : Harga Tetap
                            If Not IsNull(rs!jumlah_bayaran) Then
                                If IsNumeric(rs!jumlah_bayaran) Then Frm87_LM_JUMLAH_BAYARAN_ASAL = rs!jumlah_bayaran 'Jumlah Bayaran Terkumpul Asal (RM)
                                rs!jumlah_bayaran = Format(Frm87_LM_JUMLAH_BAYARAN_ASAL - Frm87_LM_JUMLAH_BAYARAN_ANSURAN, "0.00") 'Jumlah Bayaran Terkumpul (RM)
                            End If
                            If Not IsNull(rs!JUMLAH_BERAT) Then
                                If IsNumeric(rs!JUMLAH_BERAT) Then Frm87_LM_JUMLAH_BERAT_ASAL = rs!JUMLAH_BERAT 'Baki Berat Asal (g)
                                rs!JUMLAH_BERAT = Format(Frm87_LM_JUMLAH_BERAT_ASAL - Frm87_LM_BERAT_DIPEROLEHI, "0.00") 'Jumlah Berat Terkumpul (g)
                            End If
                            If Not IsNull(rs!BAKI_BERAT) Then
                                If IsNumeric(rs!BAKI_BERAT) Then Frm87_LM_BAKI_BERAT_ASAL = rs!BAKI_BERAT 'Baki Berat Asal (g)
                                rs!BAKI_BERAT = Format(Frm87_LM_BAKI_BERAT_ASAL + Frm87_LM_BERAT_DIPEROLEHI, "0.00") 'Baki Berat (g)
                            End If
                            If Not IsNull(rs!JUMLAH_UPAH) Then
                                If IsNumeric(rs!JUMLAH_UPAH) Then Frm87_LM_JUMLAH_UPAH_ASAL = rs!JUMLAH_UPAH 'Jumlah Upah Asal (RM)
                                rs!JUMLAH_UPAH = Format(Frm87_LM_JUMLAH_UPAH_ASAL - Frm87_LM_JUMLAH_BAYARAN_UPAH, "0.00") 'Baki Upah (RM)
                            End If
                            If Not IsNull(rs!baki_upah) Then
                                If IsNumeric(rs!baki_upah) Then Frm87_LM_BAKI_UPAH_ASAL = rs!baki_upah 'Baki Upah Asal (RM)
                                rs!baki_upah = Format(Frm87_LM_BAKI_UPAH_ASAL + Frm87_LM_JUMLAH_BAYARAN_UPAH, "0.00") 'Baki Upah (RM)
                            End If
                            
                            rs.Update
                        ElseIf rs!jenis_ansuran = 1 Then
                            If Not IsNull(rs!jumlah_bayaran) Then
                                If IsNumeric(rs!jumlah_bayaran) Then Frm87_LM_JUMLAH_BAYARAN_ASAL = rs!jumlah_bayaran 'Jumlah Bayaran Terkumpul Asal (RM)
                                rs!jumlah_bayaran = Format(Frm87_LM_JUMLAH_BAYARAN_ASAL - Frm87_LM_JUMLAH_BAYARAN_ANSURAN, "0.00") 'Jumlah Bayaran Terkumpul (RM)
                            End If
                            If Not IsNull(rs!baki_bayaran) Then
                                If IsNumeric(rs!baki_bayaran) Then Frm87_LM_BAKI_BAYARAN_ASAL = rs!baki_bayaran 'Baki Bayaran Asal (RM)
                                rs!baki_bayaran = Format(Frm87_LM_JUMLAH_BAYARAN_ANSURAN + Frm87_LM_BAKI_BAYARAN_ASAL, "0.00") 'Baki Bayaran (RM)
                            End If
                            
                            rs.Update
                        End If
                    End If

                End If
                
                rs.Close
                Set rs = Nothing
'### Update Senarai Ansuran ### - End

'###Padam Akaun Ansuran### - Start
                Frm87_LM_FLAG_SAVING = 0
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 29_akaun_ansuran where no_resit='" & Frm87_LM_No_RESIT & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Not IsNull(rs!duit_simpanan_kedai) Then
                        If Format(rs!duit_simpanan_kedai, "0.00") <> "0.00" Then
                            If IsNumeric(rs!duit_simpanan_kedai) Then Frm87_LM_SIMPANAN_DIGUNAKAN = rs!duit_simpanan_kedai
                            Frm87_LM_FLAG_SAVING = 1
                        End If
                    End If
                    rs.Delete
                    rs.Update
                End If
                
                rs.Close
                Set rs = Nothing
'###Padam Akaun Ansuran### - End

'###Update Simpanan Duit Di Kedai### - Start
                If Frm87_LM_FLAG_SAVING = 1 Then
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm87_LM_No_CUST & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                        If Not IsNull(rs!baki_simpanan) Then
                            If IsNumeric(rs!baki_simpanan) Then Frm87_LM_SIMPANAN_ASAL = rs!baki_simpanan 'Baki Simpanan Pelanggan Ini (RM)
                        End If
                        
                        rs!baki_simpanan = Format(Frm87_LM_SIMPANAN_ASAL + Frm87_LM_SIMPANAN_DIGUNAKAN, "0.00") 'Baki Simpanan Terkini Pelanggan Ini (RM)
                        
                        rs.Update
                    End If
                    
                    rs.Close
                    Set rs = Nothing
                    
'###Padam Rekod Bayaran Dalam Table Simpanan### - Start
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from 24_rekod_kewangan_pelanggan where NO_RESIT='" & Frm87_LM_No_RESIT & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                        rs.Delete
                        rs.Update
                    End If
                    
                    rs.Close
                    Set rs = Nothing
'###Padam Rekod Bayaran Dalam Table Simpanan### - End
                    
                End If
'###Update Simpanan Duit Di Kedai### - End

'### Update Log ### - Start
                user = MDI_frm1.L3_Text
                LogAct_Memory = "[" & user & "] Padam Resit Ansuran. No. Resit [" & Frm87_LM_No_RESIT & "]"
                LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                Call UpdateLog_Database
'### Update Log ### - End

                Note = "Rekod Ansuran Ini Telah Berjaya Dipadamkan." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Sistem Akan Kemaskini Senarai Rekod Bayaran Pembeli Ini." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Teruskan ?"
                
                Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
                
                If Answer = vbNo Then
                    Exit Sub
                End If
                If Answer = vbYes Then
                    Call kemaskini_rekod_bayaran
                End If
            End If
        End If
    End If
End If
End Sub
Private Sub Frm87_SM_Padam2_Click()
'on error resume next
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim rs4 As ADODB.Recordset

Dim Frm87_LM_BERAT_JUALAN As Double
Dim Frm87_LM_BEZA_BERAT_ASAL As Double
Dim Frm87_LM_BERAT_ASAL As Double
Dim Frm87_LM_SIMPANAN_DIGUNAKAN As Double
Dim Frm87_LM_SIMPANAN_ASAL As Double

DATA_FOUND = 0
Frm87_LM_No_PEKERJA = vbNullString

Frm87_LM_BERAT_JUALAN = 0
Frm87_LM_BEZA_BERAT_ASAL = 0
Frm87_LM_BERAT_ASAL = 0

If Frm87.MSFlexGrid1 <> vbNullString Then
    Frm87_LM_ID = Frm87.MSFlexGrid1.TextMatrix(Frm87.MSFlexGrid1, 2) 'No. ID
    
    If Frm87_LM_ID <> vbNullString Then
        
        Note = "Padam Data Ansuran Ini ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Semua Rekod Ansuran Dari Belian Ini Akan Dipadamkan ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then

'### Padam Maklumat Pendaftaran Belian Ansuran ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 27_senarai_ansuran where ID='" & Frm87_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!Status) Then Frm87_LM_STATUS = rs!Status 'Status
                If Not IsNull(rs!berat_jualan) Then Frm87_LM_BERAT_JUALAN = rs!berat_jualan 'Berat Jualan
                If Not IsNull(rs!jenis_produk) Then Frm87_LM_JENIS = rs!jenis_produk 'Jenis Produk
                If Not IsNull(rs!no_siri_Produk) Then Frm87_LM_No_SIRI = rs!no_siri_Produk 'No. Siri Produk
                If Not IsNull(rs!no_rujukan_pelanggan) Then Frm87_LM_No_CUST = rs!no_rujukan_pelanggan 'No. Rujukan Pembeli
                
                rs.Delete
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
'### Padam Maklumat Pendaftaran Belian Ansuran ### - End

'### Tukar Status Stok Kedai ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from Data_Database where no_siri_produk='" & Frm87_LM_No_SIRI & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Frm87_LM_STATUS = "Jelas" Then
                    If Frm87_LM_JENIS = 0 Then
                        Frm87_LM_BEZA_BERAT_ASAL = rs!beza_berat 'Beza Berat Asal (g)
                        Frm87_LM_BERAT_ASAL = rs!Berat 'Berat Asal
                        
                        rs!beza_berat = Format(Frm87_LM_BEZA_BERAT_ASAL + Frm87_LM_BERAT_JUALAN, "0.00") 'Beza Berat
                        
                        If Frm87_LM_BERAT_ASAL = (Frm87_LM_BEZA_BERAT_ASAL + Frm87_LM_BERAT_JUALAN) Then
                            rs!StatusItem = 10
                        Else
                            rs!StatusItem = 12
                        End If
                        
                    ElseIf Frm87_LM_JENIS = 1 Then
                        rs!StatusItem = 10
                    End If
                Else
                    rs!StatusItem = 10
                End If
                
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
'### Tukar Status Stok Kedai ### - End

'###Padam Bayaran Ansuran### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 28_rekod_ansuran where id_database_reg='" & Frm87_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            While rs.EOF = False

                If Not IsNull(rs!no_resit_ansuran) Then Frm87_LM_No_RESIT = rs!no_resit_ansuran 'No. Resit Ansuran
                
'###Padam Akaun Ansuran### - Start
                Frm87_LM_FLAG_SAVING = 0
                Frm87_LM_SIMPANAN_DIGUNAKAN = 0
                Frm87_LM_SIMPANAN_ASAL = 0
                
                Set rs1 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs1.Open "select * from 29_akaun_ansuran where no_resit='" & Frm87_LM_No_RESIT & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs1.EOF Then
                    If Not IsNull(rs1!duit_simpanan_kedai) Then
                        If Format(rs1!duit_simpanan_kedai, "0.00") <> "0.00" Then
                            If IsNumeric(rs1!duit_simpanan_kedai) Then Frm87_LM_SIMPANAN_DIGUNAKAN = rs1!duit_simpanan_kedai
                            Frm87_LM_FLAG_SAVING = 1
                        End If
                    End If
                    
                    rs1.Delete
                    rs1.Update
                End If
                
                rs1.Close
                Set rs1 = Nothing
'###Padam Akaun Ansuran### - End

'###Update Simpanan Duit Di Kedai### - Start
                If Frm87_LM_FLAG_SAVING = 1 Then
                    Set rs2 = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs2.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm87_LM_No_CUST & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs2.EOF Then
                        If Not IsNull(rs2!baki_simpanan) Then
                            If IsNumeric(rs2!baki_simpanan) Then Frm87_LM_SIMPANAN_ASAL = rs2!baki_simpanan 'Baki Simpanan Pelanggan Ini (RM)
                        End If
                        
                        rs2!baki_simpanan = Format(Frm87_LM_SIMPANAN_ASAL + Frm87_LM_SIMPANAN_DIGUNAKAN, "0.00") 'Baki Simpanan Terkini Pelanggan Ini (RM)
                        
                        rs2.Update
                    End If
                    
                    rs2.Close
                    Set rs2 = Nothing
                    
'###Padam Rekod Bayaran Dalam Table Simpanan### - Start
                    Set rs3 = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs3.Open "select * from 24_rekod_kewangan_pelanggan where NO_RESIT='" & Frm87_LM_No_RESIT & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs3.EOF Then
                        rs3.Delete
                        rs3.Update
                    End If
                    
                    rs3.Close
                    Set rs3 = Nothing
'###Padam Rekod Bayaran Dalam Table Simpanan### - End
                    
                End If
'###Update Simpanan Duit Di Kedai### - End
                
                rs.Delete
                rs.Update
                
                rs.MoveNext
            Wend
            
            rs.Close
            Set rs = Nothing
'###Padam Bayaran Ansuran### - End

'### Update Log ### - Start
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & user & "] Padam Data Ansuran. No. Siri Produk [" & Frm87_LM_No_SIRI & "]"
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
'### Update Log ### - End
            
            Call Frm87_Initial_Setting
            MsgBox "Data Telah Berjaya Dipadamkan.", vbInformation, "Info"
        End If
    End If
End If
End Sub
Private Sub Frm87_SM_Rekod_Click()
'on error resume next
Frm87_LM_ID = vbNullString
DATA_FOUND = 0

If Frm87.MSFlexGrid1 <> vbNullString Then
    Frm87_LM_ID = Frm87.MSFlexGrid1.TextMatrix(Frm87.MSFlexGrid1, 2) 'No. ID
    
    If Frm87_LM_ID <> vbNullString Then
        Call Frm87_Initial_Setting

        Frm87.L18_Text = Frm87_LM_ID

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 27_senarai_ansuran where ID='" & Frm87_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!Nama) Then 'Nama
                Frm87_LM_NAMA = rs!Nama
            Else
                Frm87_LM_NAMA = "----"
            End If
            If Not IsNull(rs!no_ic) Then 'No. Kad Pengenalan
                Frm87_LM_IC = rs!no_ic
            Else
                Frm87_LM_IC = "----"
            End If
            If Not IsNull(rs!no_tel) Then 'No. Telefon
                Frm87_LM_HP = rs!no_tel
            Else
                Frm87_LM_HP = "----"
            End If
            If Not IsNull(rs!no_siri_Produk) Then 'No. Siri Produk
                Frm87_LM_No_SIRI = rs!no_siri_Produk
            Else
                Frm87_LM_No_SIRI = "----"
            End If
            If Not IsNull(rs!Status) Then 'Status
                Frm87_LM_STATUS = rs!Status
            Else
                Frm87_LM_STATUS = "----"
            End If
            If Not IsNull(rs!kategori_Produk) Then 'Kategori Produk
                Frm87_LM_PRODUK = rs!kategori_Produk
            Else
                Frm87_LM_PRODUK = "----"
            End If
            If Not IsNull(rs!berat_jualan) Then 'Berat
                Frm87_LM_BERAT = Format(rs!berat_jualan, "0.00 g")
            Else
                Frm87_LM_BERAT = "----"
            End If
            If Not IsNull(rs!harga_Semasa) Then 'Harga Semasa
                Frm87_LM_HARGA_SEMASA = "RM " & Format(rs!harga_Semasa, "0.00 / g")
            Else
                Frm87_LM_HARGA_SEMASA = "----"
            End If
            If Not IsNull(rs!UPAH) Then 'Upah
                Frm87_LM_UPAH = "RM " & Format(rs!UPAH, "0.00")
            Else
                Frm87_LM_UPAH = "----"
            End If
            If Not IsNull(rs!UPAH) Then 'Upah
                Frm87_LM_UPAH = "RM " & Format(rs!UPAH, "0.00")
            Else
                Frm87_LM_UPAH = "----"
            End If
            If Not IsNull(rs!jenis_ansuran) Then 'Jenis Ansuran
                If rs!jenis_ansuran = 0 Then
                    Frm87_LM_JENIS = "Ansuran Mengikut Harga Semasa"
                    Frm87_LM_HARGA = "----"
                    Frm87_LM_JUMLAH_BAKI = "----"
                    Frm87_LM_JUMLAH_JELAS = "----"
                    
                    If Not IsNull(rs!JUMLAH_BERAT) Then 'Jumlah Berat Sudah Jelas
                        Frm87_LM_BERAT_JELAS = Format(rs!JUMLAH_BERAT, "0.00 g")
                    Else
                        Frm87_LM_BERAT_JELAS = "----"
                    End If
                    If Not IsNull(rs!BAKI_BERAT) Then 'Baki Berat
                        Frm87_LM_BAKI_BERAT = Format(rs!BAKI_BERAT, "0.00 g")
                    Else
                        Frm87_LM_BAKI_BERAT = "----"
                    End If
                    If Not IsNull(rs!JUMLAH_UPAH) Then 'Jumlah Upah Sudah Jelas
                        Frm87_LM_UPAH_JELAS = "RM " & Format(rs!JUMLAH_UPAH, "#,##0.00")
                    Else
                        Frm87_LM_UPAH_JELAS = "----"
                    End If
                    If Not IsNull(rs!baki_upah) Then 'Baki Upah
                        Frm87_LM_BAKI_UPAH = "RM " & Format(rs!baki_upah, "#,##0.00")
                    Else
                        Frm87_LM_BAKI_UPAH = "----"
                    End If
                ElseIf rs!jenis_ansuran = 1 Then
                    Frm87_LM_JENIS = "Ansuran Tetap"
                    Frm87_LM_BERAT_JELAS = "----"
                    Frm87_LM_BAKI_BERAT = "----"
                    Frm87_LM_UPAH_JELAS = "----"
                    Frm87_LM_BAKI_UPAH = "----"
                    
                    If Not IsNull(rs!harga_jualan) Then 'Harga Jualan
                        Frm87_LM_HARGA = "RM " & Format(rs!harga_jualan, "#,##0.00")
                    Else
                        Frm87_LM_HARGA = "----"
                    End If

                    If Not IsNull(rs!jumlah_bayaran) Then 'Bayaran Sudah Jelas
                        Frm87_LM_JUMLAH_JELAS = "RM " & Format(rs!jumlah_bayaran, "#,##0.00")
                    Else
                        Frm87_LM_JUMLAH_JELAS = "----"
                    End If
                    
                    If Not IsNull(rs!baki_bayaran) Then 'Baki Bayaran
                        Frm87_LM_JUMLAH_BAKI = "RM " & Format(rs!baki_bayaran, "#,##0.00")
                    Else
                        Frm87_LM_JUMLAH_BAKI = "----"
                    End If
                End If
            Else
                Frm87_LM_JENIS = "----"
                Frm87_LM_HARGA = "----"
                Frm87_LM_BERAT_JELAS = "----"
                Frm87_LM_BAKI_BERAT = "----"
                Frm87_LM_UPAH_JELAS = "----"
                Frm87_LM_BAKI_UPAH = "----"
                Frm87_LM_JUMLAH_JELAS = "----"
                Frm87_LM_JUMLAH_BAKI = "----"
            End If
            
            DATA_FOUND = 1

        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then
        
            Frm87.L29_Text = "Nama : " & Frm87_LM_NAMA & vbCrLf & _
                            "No. Kad Pengenalan : " & Frm87_LM_IC & vbCrLf & _
                            "No. Telefon : " & Frm87_LM_HP & vbCrLf & _
                            "Jenis Ansuran : " & Frm87_LM_JENIS & vbCrLf & _
                            "Status : " & Frm87_LM_STATUS & vbCrLf & _
                            "============================================================" & vbCrLf & _
                            "No. Siri Produk : " & Frm87_LM_No_SIRI & vbCrLf & _
                            "Nama Produk : " & Frm87_LM_PRODUK & vbCrLf & _
                            "Berat : " & Frm87_LM_BERAT & vbCrLf & _
                            "Harga Semasa Pendaftaran Dibuat : " & Frm87_LM_HARGA_SEMASA & vbCrLf & _
                            "Upah : " & Frm87_LM_UPAH & vbCrLf & _
                            "Tetapan Harga (Jika Harga Tetap) : " & Frm87_LM_HARGA & vbCrLf & _
                            "Bayaran Sudah Jelas (Jika Harga Tetap) : " & Frm87_LM_JUMLAH_JELAS & vbCrLf & _
                            "Baki Bayaran (Jika Ansuran Harga Tetap) : " & Frm87_LM_JUMLAH_BAKI & vbCrLf & _
                            "Jumlah Berat Sudah Dibayar (Jika Ansuran Ikut Harga Semasa) : " & Frm87_LM_BERAT_JELAS & vbCrLf & _
                            "Baki Berat (Jika Ansuran Ikut Harga Semasa) : " & Frm87_LM_BAKI_BERAT & vbCrLf & _
                            "Jumlah Upah Sudah Jelas (Jika Ansuran Ikut Harga Semasa) : " & Frm87_LM_UPAH_JELAS & vbCrLf & _
                            "Baki Upah (Jika Ansuran Ikut Harga Semasa) : " & Frm87_LM_BAKI_UPAH

            Call Frm87_Rekod_Ansuran_Header
            Call Frm87_Rekod_Ansuran
            
            Frm87.Pic7.Visible = True
        End If
    End If
End If
End Sub
Private Sub Frm87_SM_Update_Click()
'on error resume next
Frm87_LM_ID = vbNullString
Frm87_LM_BARCODE = vbNullString
Frm87_LM_KOD_PURITY = vbNullString
DATA_FOUND = 0
Frm87_LM_KATEGORI = 0
Frm87_LM_SAVING = 0

If Frm87.MSFlexGrid1 <> vbNullString Then
    Frm87_LM_ID = Frm87.MSFlexGrid1.TextMatrix(Frm87.MSFlexGrid1, 2) 'No. ID
    
    If Frm87_LM_ID <> vbNullString Then
        Call Frm87_Initial_Setting
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 27_senarai_ansuran where ID='" & Frm87_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!Status) Then
                If rs!Status = "Jelas" Then
                    rs.Close
                    Set rs = Nothing
                    
                    GoTo End_Task:
                End If
            End If
            If Not IsNull(rs!ID) Then Frm87.L18_Text = rs!ID
            
            If Not IsNull(rs!baki_upah) Then
                Frm87.L20_Text = Format(rs!baki_upah, "#,##0.00") 'Baki Upah
            Else
                Frm87.L20_Text = Format(0, "0.00") 'Baki Upah
            End If
            If Not IsNull(rs!BAKI_BERAT) Then
                Frm87.L19_Text = Format(rs!BAKI_BERAT, "#,##0.00")
            Else
                Frm87.L19_Text = "0.00"
            End If
            
            If Not IsNull(rs!no_rujukan_pelanggan) Then Frm87.L33_Text = rs!no_rujukan_pelanggan 'No. Rujukan Pelanggan
            
            If Not IsNull(rs!jenis_ansuran) Then
                If rs!jenis_ansuran = 0 Then
                    Frm87.Pic6.Visible = True
                    Frm87.CB21.Enabled = True
                    If Not IsNull(rs!no_siri_Produk) Then Frm87_LM_BARCODE = rs!no_siri_Produk 'No. Siri Produk
                    
                    
                    Frm87.L30_Text.Visible = True
                ElseIf rs!jenis_ansuran = 1 Then
                    Frm87.L30_Text.Visible = False
                    
                    If Not IsNull(rs!baki_bayaran) Then
                        Frm87.L28_Text = Format(rs!baki_bayaran, "#,##0.00")
                    Else
                        Frm87.L28_Text = "0.00"
                    End If
                
                    Frm87.Pic6.Visible = False
                    Frm87.CB20 = 1
                    Frm87.CB21.Enabled = False
                End If
            End If
            If Not IsNull(rs!no_ic) Then Frm87_LM_IC = rs!no_ic 'No. IC Pembeli
            If Not IsNull(rs!kategori_pembeli) Then
                If rs!kategori_pembeli = 1 Then
                    Frm87_LM_KATEGORI = 1
                ElseIf rs!kategori_pembeli = 2 Then
                    Frm87_LM_KATEGORI = 2
                ElseIf rs!kategori_pembeli = 3 Then
                    Frm87_LM_KATEGORI = 3
                ElseIf rs!kategori_pembeli = 4 Then
                    Frm87_LM_KATEGORI = 4
                ElseIf rs!kategori_pembeli = 5 Then
                    Frm87_LM_KATEGORI = 5
                ElseIf rs!kategori_pembeli = 6 Then
                    Frm87_LM_KATEGORI = 6
                End If
            End If
            DATA_FOUND = 1
        End If
        
        rs.Close
        Set rs = Nothing
        
        '### Carian Baki Simpanan Di Kedai ### - Start
        'If Frm87_LM_SAVING = 1 Then
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from senarai_pelanggan where no_ic='" & Frm87_LM_IC & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!baki_simpanan) Then Frm87.L27_Text = Format(rs!baki_simpanan, "#,##0.00") 'Baki Simpanan Pelanggan Ini (RM)
            End If
            
            rs.Close
            Set rs = Nothing
        'End If
        
        If Frm87_LM_BARCODE <> vbNullString Then
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from Data_Database where no_siri_Produk='" & Frm87_LM_BARCODE & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!kod_Purity) Then Frm87_LM_KOD_PURITY = rs!kod_Purity 'Kod Purity
            End If
            
            rs.Close
            Set rs = Nothing
            
            If Frm87_LM_KOD_PURITY <> vbNullString Then
            
                Frm87.TB15 = "0.00" 'Harga Semasa
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from hargaemas where Purity='" & Frm87_LM_KOD_PURITY & "' AND cawangan='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Frm87_LM_KATEGORI = 1 Then
                        If IsNumeric(rs!Harga_Pelanggan) Then Frm87.TB15 = Format(rs!Harga_Pelanggan, "#,##0.00") 'Harga Pelanggan
                    ElseIf Frm87_LM_KATEGORI = 2 Then
                        If IsNumeric(rs!Harga_Member) Then Frm87.TB15 = Format(rs!Harga_Member, "#,##0.00") 'Harga Member
                    ElseIf Frm87_LM_KATEGORI = 4 Then
                        If IsNumeric(rs!Harga_Pengedar) Then Frm87.TB15 = Format(rs!Harga_Pengedar, "#,##0.00") 'Harga Pengedar
                    ElseIf Frm87_LM_KATEGORI = 3 Then
                        If IsNumeric(rs!Harga_RAF) Then Frm87.TB15 = Format(rs!Harga_RAF, "#,##0.00") 'Harga RAF
                    ElseIf Frm87_LM_KATEGORI = 5 Then
                        If IsNumeric(rs!harga_normal_dealer) Then Frm87.TB15 = Format(rs!harga_normal_dealer, "#,##0.00") 'Harga Normal Dealer
                    ElseIf Frm87_LM_KATEGORI = 6 Then
                        If IsNumeric(rs!harga_master_dealer) Then Frm87.TB15 = Format(rs!harga_master_dealer, "#,##0.00") 'Harga Master Dealer
                    End If
                End If
                
                rs.Close
                Set rs = Nothing
                
            End If
            
        End If
        
        If DATA_FOUND = 1 Then
            Frm87.Pic4.Visible = False
            Frm87.Pic5.Visible = True
        End If
    End If
End If

Exit Sub
End_Task:

Frm87.Pic4.Visible = True

MsgBox "Status bagi belian item ini adalah [JELAS]" & vbCrLf & _
        "Anda tidak dibenarkan untuk update bayaran belian ini."
        
End Sub
Private Sub L15_Text_Click()
'on error resume next
If Frm87.Pic4.Visible = False Then
    Call Frm87_Initial_Setting

    Call Frm87_Senarai_Ansuran_Header
    'Call Frm87_Senarai_Ansuran
    Frm87.MSFlexGrid1.Rows = 1
    'Frm87.MSFlexGrid1.Cols = 0
    Frm87.MSFlexGrid1.Refresh
    
    Frm87.Pic4.Visible = True
Else
    Frm87.Pic4.Visible = False
End If
End Sub
Private Sub L19_Text_Change()
'on error resume next
Dim Frm87_LM_BAYARAN As Double
Dim Frm87_LM_HARGA_SEMASA As Double
Dim Frm87_LM_BERAT As Double
Dim Frm87_LM_BAKI_UPAH As Double

If GLOBAL_DISABLE = 0 Then

    Frm87_LM_BAYARAN = 0
    Frm87_LM_HARGA_SEMASA = 0
    Frm87_LM_BAKI = 0
    Frm87_LM_BAKI_UPAH = 0
    
    If ((Frm87.TB15 <> vbNullString And IsNumeric(Frm87.TB15)) And (Frm87.L19_Text <> vbNullString And IsNumeric(Frm87.L19_Text))) Then
        Frm87_LM_BERAT = Frm87.L19_Text 'Berat
        Frm87_LM_HARGA_SEMASA = Frm87.TB15 'Harga Emas Semasa
        
        Frm87_LM_BAKI = Format(Frm87_LM_BERAT * Frm87_LM_HARGA_SEMASA, "#,##0.00") 'Kenyataan Baki Bayaran (Ansuran)
    Else
        'Frm87_LM_BAKI = vbNullString 'Kenyataan Baki Bayaran (Ansuran)
    End If
    
    If IsNumeric(Frm87.L20_Text) Then
        Frm87_LM_BAKI_UPAH = Frm87.L20_Text
    End If
    
    Frm87.L30_Text = "Baki bayaran jika hendak bayar penuh adalah :" & vbCrLf & _
                        "Bayaran Ansuran Barang : RM " & Format(Frm87_LM_BAKI, "#,##0.00") & vbCrLf & _
                        "Baki Upah : RM " & Format(Frm87_LM_BAKI_UPAH, "#,##0.00") & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Jumlah Keseluruhan : RM " & Format(Frm87_LM_BAKI + Frm87_LM_BAKI_UPAH, "#,##0.00")
End If
End Sub

Private Sub L21_Text_Change()
'On Error Resume Next
Dim Frm87_LM_GST_ANSURAN As Double
Dim Frm87_LM_GST_UPAH As Double

Frm87_LM_GST_ANSURAN = 0
Frm87_LM_GST_UPAH = 0

If GLOBAL_DISABLE = 0 Then
    
    If (Frm87.L21_Text <> vbNullString And IsNumeric(Frm87.L21_Text)) And (Frm87.L22_Text <> vbNullString And IsNumeric(Frm87.L22_Text)) Then
        Frm87_LM_GST_ANSURAN = Frm87.L21_Text 'Jumlah GST Ansuran(RM)
        Frm87_LM_GST_UPAH = Frm87.L22_Text 'Jumlah GST Upah(RM)
        
        Frm87.TB13 = Format(Frm87_LM_GST_ANSURAN + Frm87_LM_GST_UPAH, "#,##0.00") 'Jumlah Cukai GST (RM)
    Else
        Frm87.TB13 = "0.00" 'Jumlah Cukai GST (RM)
    End If
    
End If
End Sub
Private Sub L22_Text_Change()
'On Error Resume Next
Dim Frm87_LM_GST_ANSURAN As Double
Dim Frm87_LM_GST_UPAH As Double

Frm87_LM_GST_ANSURAN = 0
Frm87_LM_GST_UPAH = 0

If GLOBAL_DISABLE = 0 Then
    
    If (Frm87.L21_Text <> vbNullString And IsNumeric(Frm87.L21_Text)) And (Frm87.L22_Text <> vbNullString And IsNumeric(Frm87.L22_Text)) Then
        Frm87_LM_GST_ANSURAN = Frm87.L21_Text 'Jumlah GST Ansuran(RM)
        Frm87_LM_GST_UPAH = Frm87.L22_Text 'Jumlah GST Upah(RM)
        
        Frm87.TB13 = Format(Frm87_LM_GST_ANSURAN + Frm87_LM_GST_UPAH, "#,##0.00") 'Jumlah Cukai GST (RM)
    Else
        Frm87.TB13 = "0.00" 'Jumlah Cukai GST (RM)
    End If
    
End If
End Sub
Private Sub L3_Text_Click()
'on error resume next
If Frm87.Pic2.Visible = False Then
    Call Frm87_Initial_Setting
    Frm87.Pic2.Visible = True
    Unload Frm26
    Unload Frm27
    Unload Frm28
    
    Frm87.Pic2.Visible = True
    Frm87.TB8.SetFocus
Else
    Frm87.Pic2.Visible = False
End If
End Sub
Private Sub MSFlexGrid1_DblClick()
'on error resume next
If Frm87.MSFlexGrid1 <> vbNullString Then
    PopupMenu Frm87_PM_Menu
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub MSFlexGrid2_DblClick()
'On Error Resume Next
If Frm87.MSFlexGrid2 <> vbNullString Then
    PopupMenu Frm87_PM_Menu1
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub

Private Sub TB12_Change()
'On Error Resume Next
Dim Frm87_LM_UPAH As Double
Dim Frm87_LM_BAYARAN As Double
Dim Frm87_LM_HARGA_SEMASA As Double

Frm87_LM_UPAH = 0
Frm87_LM_BAYARAN = 0
Frm87_LM_HARGA_SEMASA = 0

If GLOBAL_DISABLE = 0 Then

    If (Frm87.TB12 <> vbNullString And IsNumeric(Frm87.TB12)) And (Frm87.TB19 <> vbNullString And IsNumeric(Frm87.TB19)) Then
        Frm87_LM_UPAH = Frm87.TB19 'Jumlah Upah (RM)
        Frm87_LM_BAYARAN = Frm87.TB12 'Jumlah Ansuran (RM)
        
        Frm87.TB20 = Format(Frm87_LM_BAYARAN + Frm87_LM_UPAH, "#,##0.00") 'Jumlah (RM)
    Else
        Frm87.TB20 = "0.00" 'Jumlah (RM)
    End If
    
    If Frm87.Pic6.Visible = True And ((Frm87.TB12 <> vbNullString And IsNumeric(Frm87.TB12)) And (Frm87.TB15 <> vbNullString And IsNumeric(Frm87.TB15))) Then
        Frm87_LM_BAYARAN = Frm87.TB12 'Jumlah Bayaran
        Frm87_LM_HARGA_SEMASA = Frm87.TB15 'Harga Emas Semasa
        
        Frm87.TB16 = Format(Frm87_LM_BAYARAN / Frm87_LM_HARGA_SEMASA, "#,##0.00") 'Berat Diperolehi
    Else
        Frm87.TB16 = "0.00" 'Berat Diperolehi
    End If
    
    Call Frm87_LM_Detail_GST
    
End If
End Sub
Private Sub TB13_Change()
'On Error Resume Next
Dim Frm87_LM_BAYARAN As Double
Dim Frm87_LM_GST As Double

Frm87_LM_BAYARAN = 0
Frm87_LM_GST = 0

If GLOBAL_DISABLE = 0 Then

    If (Frm87.TB42 <> vbNullString And IsNumeric(Frm87.TB42)) And (Frm87.TB13 <> vbNullString And IsNumeric(Frm87.TB13)) Then
        If IsNumeric(Frm87.TB13) Then Frm87_LM_GST = Frm87.TB13 'Jumlah GST (RM)
        If IsNumeric(Frm87.TB42) Then Frm87_LM_BAYARAN = Frm87.TB42 'Jumlah (RM)
        
        Frm87.TB14 = Format(Frm87_LM_GST + Frm87_LM_BAYARAN, "#,##0.00") 'Jumlah Cukai GST (RM)
    Else
        Frm87.TB14 = "0.00" 'Jumlah Cukai GST (RM)
    End If
    
End If
End Sub
Private Sub TB14_Change()
'on error resume next
Dim Frm87_HARGA_ASAL As Double
Dim Frm87_ADJUSTMENT As Double

If GLOBAL_DISABLE = 0 Then

    Frm87_HARGA_ASAL = 0
    Frm87_ADJUSTMENT = 0
    
    If ((Frm87.TB14 <> vbNullString And IsNumeric(Frm87.TB14)) And (Frm87.TB17 <> vbNullString And IsNumeric(Frm87.TB17))) Then
        Frm87_HARGA_ASAL = Frm87.TB14 'Harga Asal
        Frm87_ADJUSTMENT = Frm87.TB17 'Adjustment
        
        Frm87.TB18 = Format(Frm87_HARGA_ASAL - Frm87_ADJUSTMENT, "#,##0.00") 'Jumlah Bayaran
    Else
        Frm87.TB18 = "0.00" 'Jumlah Bayaran
    End If
    
End If
End Sub
Private Sub TB15_Change()
'on error resume next
Dim Frm87_LM_BAYARAN As Double
Dim Frm87_LM_HARGA_SEMASA As Double
Dim Frm87_LM_BERAT As Double
Dim Frm87_LM_BAKI_UPAH As Double

If GLOBAL_DISABLE = 0 Then

    Frm87_LM_BAYARAN = 0
    Frm87_LM_HARGA_SEMASA = 0
    Frm87_LM_BAKI = 0
    Frm87_LM_BAKI_UPAH = 0
    
    If Frm87.Pic6.Visible = True And ((Frm87.TB12 <> vbNullString And IsNumeric(Frm87.TB12)) And (Frm87.TB15 <> vbNullString And IsNumeric(Frm87.TB15))) Then
        Frm87_LM_BAYARAN = Frm87.TB12 'Jumlah Bayaran
        Frm87_LM_HARGA_SEMASA = Frm87.TB15 'Harga Emas Semasa
        
        Frm87.TB16 = Format(Frm87_LM_BAYARAN / Frm87_LM_HARGA_SEMASA, "#,##0.00") 'Berat Diperolehi
    Else
        Frm87.TB16 = "0.00" 'Berat Diperolehi
    End If
    
    If ((Frm87.TB15 <> vbNullString And IsNumeric(Frm87.TB15)) And (Frm87.L19_Text <> vbNullString And IsNumeric(Frm87.L19_Text))) Then
        Frm87_LM_BERAT = Frm87.L19_Text 'Berat
        Frm87_LM_HARGA_SEMASA = Frm87.TB15 'Harga Emas Semasa
        
        Frm87_LM_BAKI = Format(Frm87_LM_BERAT * Frm87_LM_HARGA_SEMASA, "#,##0.00") 'Kenyataan Baki Bayaran (Ansuran)
    Else
        'Frm87_LM_BAKI = vbNullString 'Kenyataan Baki Bayaran (Ansuran)
    End If
    
    If IsNumeric(Frm87.L20_Text) Then
        Frm87_LM_BAKI_UPAH = Format(Frm87.L20_Text, "#,##0.00")
    End If
    
    Frm87.L30_Text = "Baki bayaran jika hendak bayar penuh adalah :" & vbCrLf & _
                        "Bayaran Ansuran Barang : RM " & Format(Frm87_LM_BAKI, "#,##0.00") & vbCrLf & _
                        "Baki Upah : RM " & Format(Frm87_LM_BAKI_UPAH, "#,##0.00") & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Jumlah Keseluruhan : RM " & Format(Frm87_LM_BAKI + Frm87_LM_BAKI_UPAH, "#,##0.00")
End If
End Sub
Private Sub TB17_Change()
'on error resume next
Dim Frm87_HARGA_ASAL As Double
Dim Frm87_ADJUSTMENT As Double

If GLOBAL_DISABLE = 0 Then

    Frm87_HARGA_ASAL = 0
    Frm87_ADJUSTMENT = 0
    
    If ((Frm87.TB14 <> vbNullString And IsNumeric(Frm87.TB14)) And (Frm87.TB17 <> vbNullString And IsNumeric(Frm87.TB17))) Then
        Frm87_HARGA_ASAL = Frm87.TB14 'Harga Asal
        Frm87_ADJUSTMENT = Frm87.TB17 'Adjustment
        
        Frm87.TB18 = Format(Frm87_HARGA_ASAL - Frm87_ADJUSTMENT, "#,##0.00") 'Jumlah Bayaran
    Else
        Frm87.TB18 = "0.00" 'Jumlah Bayaran
    End If
    
End If
End Sub
Private Sub TB18_Change()
'On Error Resume Next
Frm87.TB27 = Frm87.TB18
End Sub
Private Sub TB19_Change()
'On Error Resume Next
Dim Frm87_LM_KADAR_GST As Double
Dim Frm87_LM_BAYARAN As Double
Dim Frm87_LM_UPAH As Double

Frm87_LM_KADAR_GST = 0
Frm87_LM_BAYARAN = 0
Frm87_LM_UPAH = 0

If GLOBAL_DISABLE = 0 Then
    If (Frm87.TB12 <> vbNullString And IsNumeric(Frm87.TB12)) And (Frm87.TB19 <> vbNullString And IsNumeric(Frm87.TB19)) Then
        Frm87_LM_UPAH = Frm87.TB19 'Jumlah Upah (RM)
        Frm87_LM_BAYARAN = Frm87.TB12 'Jumlah Ansuran (RM)
        
        Frm87.TB20 = Format(Frm87_LM_BAYARAN + Frm87_LM_UPAH, "#,##0.00") 'Jumlah (RM)
    Else
        Frm87.TB20 = "0.00" 'Jumlah (RM)
    End If
    
    Call Frm87_LM_Detail_GST
    
End If
End Sub
Private Sub TB20_Change()
'On Error Resume Next
Dim Frm87_LM_KADAR_GST As Double
Dim Frm87_LM_BAYARAN As Double
Dim Frm87_LM_GST As Double

Frm87_LM_KADAR_GST = 0
Frm87_LM_BAYARAN = 0
Frm87_LM_GST = 0

If GLOBAL_DISABLE = 0 Then
    If Frm87.CB27 = 0 Then
        If Frm87.CB23 = 1 And (Frm87.TB20 <> vbNullString And IsNumeric(Frm87.TB20)) And (Frm87.L17_Text <> vbNullString And IsNumeric(Frm87.L17_Text)) Then
        
            If IsNumeric(Frm87.L17_Text) Then Frm87_LM_KADAR_GST = Frm87.L17_Text 'Jumlah Kadar GST (%)
            If IsNumeric(Frm87.TB20) Then Frm87_LM_ANSURAN = Frm87.TB20 'Jumlah Ansuran (RM)
            
            Frm87.TB13 = Format((Frm87_LM_KADAR_GST / 100) * Frm87_LM_ANSURAN, "#,##0.00") 'Jumlah Cukai GST (RM)
            Frm87.L21_Text = Format((Frm87_LM_KADAR_GST / 100) * Frm87_LM_ANSURAN, "#,##0.00") 'Jumlah Cukai GST (RM)
            Frm87.TB42 = Format(Frm87_LM_ANSURAN, "#,##0.00") 'Jumlah Harga Tanpa GST (RM)
            
        Else
        
            Frm87.TB13 = "0.00" 'Jumlah Cukai GST (RM)
            Frm87.L21_Text = "0.00" 'Jumlah Cukai GST (RM)
            If IsNumeric(Frm87.TB20) Then
                Frm87.TB42 = Format(Frm87.TB20, "#,##0.00")
            Else
                Frm87.TB42 = Format(0, "#,##0.00")
            End If
        
        
        End If
            
    ElseIf Frm87.CB27 = 1 Then
        
        If Frm87.CB23 = 1 And (Frm87.TB20 <> vbNullString And IsNumeric(Frm87.TB20)) And (Frm87.L17_Text <> vbNullString And IsNumeric(Frm87.L17_Text)) Then
        
            If IsNumeric(Frm87.L17_Text) Then Frm87_LM_KADAR_GST = Frm87.L17_Text 'Jumlah Kadar GST (%)
            If IsNumeric(Frm87.TB20) Then Frm87_LM_ANSURAN = Frm87.TB20 'Jumlah Ansuran (RM)
            
            Frm87.TB13 = Format(Frm87_LM_ANSURAN - (Frm87_LM_ANSURAN / (1 + Frm87_LM_KADAR_GST / 100)), "#,##0.00") 'Jumlah Cukai GST (RM)
            Frm87.L21_Text = Format(Frm87_LM_ANSURAN - (Frm87_LM_ANSURAN / (1 + Frm87_LM_KADAR_GST / 100)), "#,##0.00") 'Jumlah Cukai GST (RM)
            Frm87.TB42 = Format(Frm87_LM_ANSURAN / (1 + (Frm87_LM_KADAR_GST / 100)), "#,##0.00") 'Jumlah Harga Tanpa GST (RM)
            
        Else
    
            Frm87.TB13 = "0.00" 'Jumlah Cukai GST (RM)
            Frm87.L21_Text = "0.00" 'Jumlah Cukai GST (RM)
            
            If IsNumeric(Frm87.TB20) Then
                Frm87.TB42 = Format(Frm87.TB20, "#,##0.00")
            Else
                Frm87.TB42 = Format(0, "#,##0.00")
            End If
            
        End If
        
    End If
    
    If Frm87.CB22 = 1 Then
        
        If (Frm87.TB20 <> vbNullString And IsNumeric(Frm87.TB20)) Then
        
            If IsNumeric(Frm87.TB20) Then Frm87_LM_ANSURAN = Frm87.TB20 'Jumlah Ansuran (RM)
            
            Frm87.TB42 = Format(Frm87_LM_ANSURAN, "#,##0.00") 'Jumlah Harga Tanpa GST (RM)
            
        Else
    
            Frm87.TB13 = "0.00" 'Jumlah Cukai GST (RM)
            Frm87.L21_Text = "0.00" 'Jumlah Cukai GST (RM)
            
            If IsNumeric(Frm87.TB20) Then
                Frm87.TB42 = Format(Frm87.TB20, "#,##0.00")
            Else
                Frm87.TB42 = Format(0, "#,##0.00")
            End If
            
        End If
        
    End If
End If

Call Frm87_LM_Detail_GST
End Sub
Private Sub TB21_Change()
'On Error Resume Next
Dim Frm87_LM_TUNAI As Double
Dim Frm87_LM_BANK As Double
Dim Frm87_LM_KREDIT As Double
Dim Frm87_LM_DEBIT As Double
Dim Frm87_LM_SIMPANAN As Double

Frm87_LM_TUNAI = 0
Frm87_LM_BANK = 0
Frm87_LM_KREDIT = 0
Frm87_LM_DEBIT = 0
Frm87_LM_SIMPANAN = 0

If IsNumeric(Frm87.TB27) Then
    Frm87_LM_TUNAI = Frm87.TB27
End If
If IsNumeric(Frm87.TB28) Then
    Frm87_LM_BANK = Frm87.TB28
End If
If IsNumeric(Frm87.TB29) Then
    Frm87_LM_KREDIT = Frm87.TB29
End If
If IsNumeric(Frm87.TB38) Then
    Frm87_LM_DEBIT = Frm87.TB38
End If
If IsNumeric(Frm87.TB21) Then
    Frm87_LM_SIMPANAN = Frm87.TB21
End If

Frm87.TB32 = Format(Frm87_LM_TUNAI + Frm87_LM_BANK + Frm87_LM_KREDIT + Frm87_LM_DEBIT + Frm87_LM_SIMPANAN, "#,##0.00") 'Jumlah Bayaran Keseluruhan
End Sub
Private Sub TB27_Change()
'On Error Resume Next
Dim Frm87_LM_TUNAI As Double
Dim Frm87_LM_BANK As Double
Dim Frm87_LM_KREDIT As Double
Dim Frm87_LM_DEBIT As Double
Dim Frm87_LM_SIMPANAN As Double

Frm87_LM_TUNAI = 0
Frm87_LM_BANK = 0
Frm87_LM_KREDIT = 0
Frm87_LM_DEBIT = 0
Frm87_LM_SIMPANAN = 0

If IsNumeric(Frm87.TB27) Then
    Frm87_LM_TUNAI = Frm87.TB27
End If
If IsNumeric(Frm87.TB28) Then
    Frm87_LM_BANK = Frm87.TB28
End If
If IsNumeric(Frm87.TB29) Then
    Frm87_LM_KREDIT = Frm87.TB29
End If
If IsNumeric(Frm87.TB38) Then
    Frm87_LM_DEBIT = Frm87.TB38
End If
If IsNumeric(Frm87.TB21) Then
    Frm87_LM_SIMPANAN = Frm87.TB21
End If

Frm87.TB32 = Format(Frm87_LM_TUNAI + Frm87_LM_BANK + Frm87_LM_KREDIT + Frm87_LM_DEBIT + Frm87_LM_SIMPANAN, "#,##0.00") 'Jumlah Bayaran Keseluruhan
End Sub
Private Sub TB28_Change()
'On Error Resume Next
Dim Frm87_LM_TUNAI As Double
Dim Frm87_LM_BANK As Double
Dim Frm87_LM_KREDIT As Double
Dim Frm87_LM_DEBIT As Double
Dim Frm87_LM_SIMPANAN As Double

Frm87_LM_TUNAI = 0
Frm87_LM_BANK = 0
Frm87_LM_KREDIT = 0
Frm87_LM_DEBIT = 0
Frm87_LM_SIMPANAN = 0

If IsNumeric(Frm87.TB27) Then
    Frm87_LM_TUNAI = Frm87.TB27
End If
If IsNumeric(Frm87.TB28) Then
    Frm87_LM_BANK = Frm87.TB28
End If
If IsNumeric(Frm87.TB29) Then
    Frm87_LM_KREDIT = Frm87.TB29
End If
If IsNumeric(Frm87.TB38) Then
    Frm87_LM_DEBIT = Frm87.TB38
End If
If IsNumeric(Frm87.TB21) Then
    Frm87_LM_SIMPANAN = Frm87.TB21
End If

Frm87.TB32 = Format(Frm87_LM_TUNAI + Frm87_LM_BANK + Frm87_LM_KREDIT + Frm87_LM_DEBIT + Frm87_LM_SIMPANAN, "#,##0.00") 'Jumlah Bayaran Keseluruhan
End Sub
Private Sub TB29_Change()
'On Error Resume Next
Dim Frm87_LM_TUNAI As Double
Dim Frm87_LM_BANK As Double
Dim Frm87_LM_KREDIT As Double
Dim Frm87_LM_DEBIT As Double
Dim Frm87_LM_SIMPANAN As Double

Frm87_LM_TUNAI = 0
Frm87_LM_BANK = 0
Frm87_LM_KREDIT = 0
Frm87_LM_DEBIT = 0
Frm87_LM_SIMPANAN = 0


If IsNumeric(Frm87.TB27) Then
    Frm87_LM_TUNAI = Frm87.TB27
End If
If IsNumeric(Frm87.TB28) Then
    Frm87_LM_BANK = Frm87.TB28
End If
If IsNumeric(Frm87.TB29) Then
    Frm87_LM_KREDIT = Frm87.TB29
End If
If IsNumeric(Frm87.TB38) Then
    Frm87_LM_DEBIT = Frm87.TB38
End If
If IsNumeric(Frm87.TB21) Then
    Frm87_LM_SIMPANAN = Frm87.TB21
End If

Frm87.TB32 = Format(Frm87_LM_TUNAI + Frm87_LM_BANK + Frm87_LM_KREDIT + Frm87_LM_DEBIT + Frm87_LM_SIMPANAN, "#,##0.00")  'Jumlah Bayaran Keseluruhan

If IsNumeric(Frm87.L31_Text) Then
    Frm87.TB30 = Format((Frm87.L31_Text / 100) * Frm87_LM_KREDIT, "#,##0.00") 'Cas Kad Kredit
    Frm87.TB31 = Format((Frm87_LM_KREDIT + (Frm87.L31_Text / 100) * Frm87_LM_KREDIT), "#,##0.00") 'Cas Kad Kredit
Else
    Frm87.TB30 = "0.00" 'Cas Kad Kredit
    Frm87.TB31 = "0.00" 'Jumlah Potongan Kad
End If
End Sub
Private Sub TB38_Change()
'On Error Resume Next
Dim Frm87_LM_TUNAI As Double
Dim Frm87_LM_BANK As Double
Dim Frm87_LM_KREDIT As Double
Dim Frm87_LM_DEBIT As Double
Dim Frm87_LM_SIMPANAN As Double

Frm87_LM_TUNAI = 0
Frm87_LM_BANK = 0
Frm87_LM_KREDIT = 0
Frm87_LM_DEBIT = 0
Frm87_LM_SIMPANAN = 0

If IsNumeric(Frm87.TB27) Then
    Frm87_LM_TUNAI = Frm87.TB27
End If
If IsNumeric(Frm87.TB28) Then
    Frm87_LM_BANK = Frm87.TB28
End If
If IsNumeric(Frm87.TB29) Then
    Frm87_LM_KREDIT = Frm87.TB29
End If
If IsNumeric(Frm87.TB38) Then
    Frm87_LM_DEBIT = Frm87.TB38
End If
If IsNumeric(Frm87.TB21) Then
    Frm87_LM_SIMPANAN = Frm87.TB21
End If

Frm87.TB32 = Format(Frm87_LM_TUNAI + Frm87_LM_BANK + Frm87_LM_KREDIT + Frm87_LM_DEBIT + Frm87_LM_SIMPANAN, "#,##0.00")  'Jumlah Bayaran Keseluruhan

If IsNumeric(Frm87.L32_Text) Then
    Frm87.TB39 = Format((Frm87.L32_Text / 100) * Frm87_LM_DEBIT, "#,##0.00") 'Cas Kad Debit
    Frm87.TB40 = Format((Frm87_LM_DEBIT + (Frm87.L32_Text / 100) * Frm87_LM_DEBIT), "#,##0.00") 'Jumlah Potong Debit Kad
Else
    Frm87.TB39 = "0.00" 'Cas Debit Kad
    Frm87.TB40 = "0.00" 'Jumlah Potongan Kad Debit
End If
End Sub
Private Sub TB4_Change()
'on error resume next
Dim Frm87_BERAT As Double
Dim Frm87_HARGA_PER_GRAM As Double
Dim Frm87_UPAH As Double

If GLOBAL_DISABLE = 0 Then

    Frm87_BERAT = 0
    Frm87_HARGA_PER_GRAM = 0
    Frm87_UPAH = 0
    
    If ((Frm87.TB4 <> vbNullString And IsNumeric(Frm87.TB4)) And (Frm87.TB5 <> vbNullString And IsNumeric(Frm87.TB5)) And (Frm87.TB6 <> vbNullString And IsNumeric(Frm87.TB6))) Then
        Frm87_BERAT = Frm87.TB4 'Berat
        Frm87_HARGA_PER_GRAM = Frm87.TB5 'Harga Per Gram
        Frm87_UPAH = Frm87.TB6 'Upah
        
        Frm87.TB7 = Format((Frm87_BERAT * Frm87_HARGA_PER_GRAM) + Frm87_UPAH, "#,##0.00") 'Harga Asal
    Else
        Frm87.TB7 = "0.00" 'Harga Asal
    End If
    
End If
End Sub
Private Sub TB42_Change()
'On Error Resume Next
Dim Frm87_LM_BAYARAN As Double
Dim Frm87_LM_GST As Double

Frm87_LM_BAYARAN = 0
Frm87_LM_GST = 0

If GLOBAL_DISABLE = 0 Then

    If (Frm87.TB42 <> vbNullString And IsNumeric(Frm87.TB42)) And (Frm87.TB13 <> vbNullString And IsNumeric(Frm87.TB13)) Then
        If IsNumeric(Frm87.TB13) Then Frm87_LM_GST = Frm87.TB13 'Jumlah GST (RM)
        If IsNumeric(Frm87.TB42) Then Frm87_LM_BAYARAN = Frm87.TB42 'Jumlah (RM)
        
        Frm87.TB14 = Format(Frm87_LM_GST + Frm87_LM_BAYARAN, "#,##0.00") 'Jumlah Cukai GST (RM)
    Else
        Frm87.TB14 = "0.00" 'Jumlah Cukai GST (RM)
    End If
    
End If
End Sub
Private Sub TB5_Change()
'on error resume next
Dim Frm87_BERAT As Double
Dim Frm87_HARGA_PER_GRAM As Double
Dim Frm87_UPAH As Double

If GLOBAL_DISABLE = 0 Then

    Frm87_BERAT = 0
    Frm87_HARGA_PER_GRAM = 0
    Frm87_UPAH = 0
    
    If ((Frm87.TB4 <> vbNullString And IsNumeric(Frm87.TB4)) And (Frm87.TB5 <> vbNullString And IsNumeric(Frm87.TB5)) And (Frm87.TB6 <> vbNullString And IsNumeric(Frm87.TB6))) Then
        Frm87_BERAT = Frm87.TB4 'Berat
        Frm87_HARGA_PER_GRAM = Frm87.TB5 'Harga Per Gram
        Frm87_UPAH = Frm87.TB6 'Upah
        
        Frm87.TB7 = Format((Frm87_BERAT * Frm87_HARGA_PER_GRAM) + Frm87_UPAH, "#,##0.00") 'Harga Asal
    Else
        Frm87.TB7 = "0.00" 'Harga Asal
    End If
    
End If
End Sub
Private Sub TB6_Change()
'on error resume next
Dim Frm87_BERAT As Double
Dim Frm87_HARGA_PER_GRAM As Double
Dim Frm87_UPAH As Double

If GLOBAL_DISABLE = 0 Then

    Frm87_BERAT = 0
    Frm87_HARGA_PER_GRAM = 0
    Frm87_UPAH = 0
    
    If ((Frm87.TB4 <> vbNullString And IsNumeric(Frm87.TB4)) And (Frm87.TB5 <> vbNullString And IsNumeric(Frm87.TB5)) And (Frm87.TB6 <> vbNullString And IsNumeric(Frm87.TB6))) Then
        Frm87_BERAT = Frm87.TB4 'Berat
        Frm87_HARGA_PER_GRAM = Frm87.TB5 'Harga Per Gram
        Frm87_UPAH = Frm87.TB6 'Upah
        
        Frm87.TB7 = Format((Frm87_BERAT * Frm87_HARGA_PER_GRAM) + Frm87_UPAH, "#,##0.00") 'Harga Asal
    Else
        Frm87.TB7 = "0.00" 'Harga Asal
    End If
    
End If
End Sub
Private Sub TB7_Change()
'on error resume next
Dim Frm87_HARGA_ASAL As Double
Dim Frm87_ADJUSTMENT As Double

If GLOBAL_DISABLE = 0 Then

    Frm87_HARGA_ASAL = 0
    Frm87_ADJUSTMENT = 0
    
    If ((Frm87.TB7 <> vbNullString And IsNumeric(Frm87.TB7)) And (Frm87.TB9 <> vbNullString And IsNumeric(Frm87.TB9))) Then
        Frm87_HARGA_ASAL = Frm87.TB7 'Harga Asal
        Frm87_ADJUSTMENT = Frm87.TB9 'Adjustment
        
        Frm87.TB10 = Format(Frm87_HARGA_ASAL - Frm87_ADJUSTMENT, "#,##0.00") 'Harga Jualan
    Else
        Frm87.TB10 = "0.00" 'Harga Jualan
    End If
    
End If
End Sub
Private Sub TB8_Change()
'on error resume next
If Frm87.CB13 = 1 And Frm87.TB8 <> vbNullString Then
    Frm87.Tmr2.Enabled = False
    Frm87.Tmr2.Enabled = True
    Frm87.Tmr2.Interval = 100
End If
End Sub
Private Sub TB9_Change()
'on error resume next
Dim Frm87_HARGA_ASAL As Double
Dim Frm87_ADJUSTMENT As Double

If GLOBAL_DISABLE = 0 Then

    Frm87_HARGA_ASAL = 0
    Frm87_ADJUSTMENT = 0
    
    If ((Frm87.TB7 <> vbNullString And IsNumeric(Frm87.TB7)) And (Frm87.TB9 <> vbNullString And IsNumeric(Frm87.TB9))) Then
        Frm87_HARGA_ASAL = Frm87.TB7 'Harga Lepas Diskaun
        Frm87_ADJUSTMENT = Frm87.TB9 'Adjustment
        
        Frm87.TB10 = Format(Frm87_HARGA_ASAL - Frm87_ADJUSTMENT, "#,##0.00") 'Harga Jualan
    Else
        Frm87.TB10 = "0.00" 'Harga Jualan
    End If
    
End If
End Sub
Private Sub Tmr1_Timer()
'on error resume next
Frm87.L1_Text = DateTime.Date
Frm87.L2_Text = DateTime.Time$
End Sub
Private Sub kemaskini_rekod_bayaran()
'on error resume next
Frm87_LM_ID = vbNullString
DATA_FOUND = 0

If Frm87.MSFlexGrid1 <> vbNullString Then
    Frm87_LM_ID = Frm87.MSFlexGrid1.TextMatrix(Frm87.MSFlexGrid1, 2) 'No. ID
    
    If Frm87_LM_ID <> vbNullString Then

        'Frm87.L18_Text = Frm87_LM_ID

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 27_senarai_ansuran where ID='" & Frm87_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!Nama) Then 'Nama
                Frm87_LM_NAMA = rs!Nama
            Else
                Frm87_LM_NAMA = "----"
            End If
            If Not IsNull(rs!no_ic) Then 'No. Kad Pengenalan
                Frm87_LM_IC = rs!no_ic
            Else
                Frm87_LM_IC = "----"
            End If
            If Not IsNull(rs!no_tel) Then 'No. Telefon
                Frm87_LM_HP = rs!no_tel
            Else
                Frm87_LM_HP = "----"
            End If
            If Not IsNull(rs!no_siri_Produk) Then 'No. Siri Produk
                Frm87_LM_No_SIRI = rs!no_siri_Produk
            Else
                Frm87_LM_No_SIRI = "----"
            End If
            If Not IsNull(rs!Status) Then 'Status
                Frm87_LM_STATUS = rs!Status
            Else
                Frm87_LM_STATUS = "----"
            End If
            If Not IsNull(rs!kategori_Produk) Then 'Kategori Produk
                Frm87_LM_PRODUK = rs!kategori_Produk
            Else
                Frm87_LM_PRODUK = "----"
            End If
            If Not IsNull(rs!berat_jualan) Then 'Berat
                Frm87_LM_BERAT = Format(rs!berat_jualan, "0.00 g")
            Else
                Frm87_LM_BERAT = "----"
            End If
            If Not IsNull(rs!harga_Semasa) Then 'Harga Semasa
                Frm87_LM_HARGA_SEMASA = "RM " & Format(rs!harga_Semasa, "0.00 / g")
            Else
                Frm87_LM_HARGA_SEMASA = "----"
            End If
            If Not IsNull(rs!UPAH) Then 'Upah
                Frm87_LM_UPAH = "RM " & Format(rs!UPAH, "#,##0.00")
            Else
                Frm87_LM_UPAH = "----"
            End If
            If Not IsNull(rs!UPAH) Then 'Upah
                Frm87_LM_UPAH = "RM " & Format(rs!UPAH, "#,##0.00")
            Else
                Frm87_LM_UPAH = "----"
            End If
            If Not IsNull(rs!jenis_ansuran) Then 'Jenis Ansuran
                If rs!jenis_ansuran = 0 Then
                    Frm87_LM_JENIS = "Ansuran Mengikut Harga Semasa"
                    Frm87_LM_HARGA = "----"
                    Frm87_LM_JUMLAH_BAKI = "----"
                    Frm87_LM_JUMLAH_JELAS = "----"
                    
                    If Not IsNull(rs!JUMLAH_BERAT) Then 'Jumlah Berat Sudah Jelas
                        Frm87_LM_BERAT_JELAS = Format(rs!JUMLAH_BERAT, "0.00 g")
                    Else
                        Frm87_LM_BERAT_JELAS = "----"
                    End If
                    If Not IsNull(rs!BAKI_BERAT) Then 'Baki Berat
                        Frm87_LM_BAKI_BERAT = Format(rs!BAKI_BERAT, "0.00 g")
                    Else
                        Frm87_LM_BAKI_BERAT = "----"
                    End If
                    If Not IsNull(rs!JUMLAH_UPAH) Then 'Jumlah Upah Sudah Jelas
                        Frm87_LM_UPAH_JELAS = "RM " & Format(rs!JUMLAH_UPAH, "#,##0.00")
                    Else
                        Frm87_LM_UPAH_JELAS = "----"
                    End If
                    If Not IsNull(rs!baki_upah) Then 'Baki Upah
                        Frm87_LM_BAKI_UPAH = "RM " & Format(rs!baki_upah, "#,##0.00")
                    Else
                        Frm87_LM_BAKI_UPAH = "----"
                    End If
                ElseIf rs!jenis_ansuran = 1 Then
                    Frm87_LM_JENIS = "Ansuran Tetap"
                    Frm87_LM_BERAT_JELAS = "----"
                    Frm87_LM_BAKI_BERAT = "----"
                    Frm87_LM_UPAH_JELAS = "----"
                    Frm87_LM_BAKI_UPAH = "----"
                    
                    If Not IsNull(rs!harga_jualan) Then 'Harga Jualan
                        Frm87_LM_HARGA = "RM " & Format(rs!harga_jualan, "#,##0.00")
                    Else
                        Frm87_LM_HARGA = "----"
                    End If

                    If Not IsNull(rs!jumlah_bayaran) Then 'Bayaran Sudah Jelas
                        Frm87_LM_JUMLAH_JELAS = "RM " & Format(rs!jumlah_bayaran, "#,##0.00")
                    Else
                        Frm87_LM_JUMLAH_JELAS = "----"
                    End If
                    
                    If Not IsNull(rs!baki_bayaran) Then 'Baki Bayaran
                        Frm87_LM_JUMLAH_BAKI = "RM " & Format(rs!baki_bayaran, "#,##0.00")
                    Else
                        Frm87_LM_JUMLAH_BAKI = "----"
                    End If
                End If
            Else
                Frm87_LM_JENIS = "----"
                Frm87_LM_HARGA = "----"
                Frm87_LM_BERAT_JELAS = "----"
                Frm87_LM_BAKI_BERAT = "----"
                Frm87_LM_UPAH_JELAS = "----"
                Frm87_LM_BAKI_UPAH = "----"
                Frm87_LM_JUMLAH_JELAS = "----"
                Frm87_LM_JUMLAH_BAKI = "----"
            End If
            
            DATA_FOUND = 1

        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then
        
            Frm87.L29_Text = "Nama : " & Frm87_LM_NAMA & vbCrLf & _
                            "No. Kad Pengenalan : " & Frm87_LM_IC & vbCrLf & _
                            "No. Telefon : " & Frm87_LM_HP & vbCrLf & _
                            "Jenis Ansuran : " & Frm87_LM_JENIS & vbCrLf & _
                            "Status : " & Frm87_LM_STATUS & vbCrLf & _
                            "============================================================" & vbCrLf & _
                            "No. Siri Produk : " & Frm87_LM_No_SIRI & vbCrLf & _
                            "Nama Produk : " & Frm87_LM_PRODUK & vbCrLf & _
                            "Berat : " & Frm87_LM_BERAT & vbCrLf & _
                            "Harga Semasa Pendaftaran Dibuat : " & Frm87_LM_HARGA_SEMASA & vbCrLf & _
                            "Upah : " & Frm87_LM_UPAH & vbCrLf & _
                            "Tetapan Harga (Jika Harga Tetap) : " & Frm87_LM_HARGA & vbCrLf & _
                            "Bayaran Sudah Jelas (Jika Harga Tetap) : " & Frm87_LM_JUMLAH_JELAS & vbCrLf & _
                            "Baki Bayaran (Jika Ansuran Harga Tetap) : " & Frm87_LM_JUMLAH_BAKI & vbCrLf & _
                            "Jumlah Berat Sudah Dibayar (Jika Ansuran Ikut Harga Semasa) : " & Frm87_LM_BERAT_JELAS & vbCrLf & _
                            "Baki Berat (Jika Ansuran Ikut Harga Semasa) : " & Frm87_LM_BAKI_BERAT & vbCrLf & _
                            "Jumlah Upah Sudah Jelas (Jika Ansuran Ikut Harga Semasa) : " & Frm87_LM_UPAH_JELAS & vbCrLf & _
                            "Baki Upah (Jika Ansuran Ikut Harga Semasa) : " & Frm87_LM_BAKI_UPAH

            Call Frm87_Rekod_Ansuran_Header
            Call Frm87_Rekod_Ansuran
        End If
    End If
End If
End Sub
Private Sub padam_rekod_ansuran()
'on error resume next
Dim Frm87_LM_JUMLAH_BAYARAN_ASAL As Double
Dim Frm87_LM_BAKI_BERAT_ASAL As Double
Dim Frm87_LM_BAKI_UPAH_ASAL As Double
Dim Frm87_LM_BAKI_BAYARAN_ASAL As Double
Dim Frm87_LM_BERAT_DIPEROLEHI As Double
Dim Frm87_LM_JUMLAH_BAYARAN_UPAH As Double
Dim Frm87_LM_JUMLAH_BAYARAN_ANSURAN As Double
Dim Frm87_LM_JUMLAH_BERAT_ASAL As Double
Dim Frm87_LM_JUMLAH_UPAH_ASAL As Double
Dim Frm87_LM_SIMPANAN_DIGUNAKAN As Double
Dim Frm87_LM_SIMPANAN_ASAL As Double

Frm87_LM_SIMPANAN_ASAL = 0
Frm87_LM_SIMPANAN_DIGUNAKAN = 0
Frm87_LM_JUMLAH_BAYARAN_ASAL = 0
Frm87_LM_BAKI_BERAT_ASAL = 0
Frm87_LM_BAKI_UPAH_ASAL = 0
Frm87_LM_BAKI_BAYARAN_ASAL = 0
Frm87_LM_BERAT_DIPEROLEHI = 0
Frm87_LM_JUMLAH_BAYARAN_UPAH = 0
Frm87_LM_JUMLAH_BAYARAN_ANSURAN = 0
Frm87_LM_JUMLAH_BERAT_ASAL = 0
Frm87_LM_JUMLAH_UPAH_ASAL = 0

Frm87_LM_ID = vbNullString
Frm87_LM_No_PEKERJA = vbNullString
DATA_FOUND = 0

    
'###Padam Bayaran Ansuran### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 28_rekod_ansuran where no_resit_ansuran='" & Frm87.L12_Text & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!id_database_reg) Then Frm87_LM_ID_ASAL = rs!id_database_reg 'ID Asal
    If Not IsNull(rs!no_resit_ansuran) Then Frm87_LM_No_RESIT = rs!no_resit_ansuran 'No. Resit Ansuran
    If Not IsNull(rs!id_database_reg) Then Frm87_LM_ID_ASAL = rs!id_database_reg 'No. ID Dari Database Senarai Pembeli Ansuran
    If Not IsNull(rs!jumlah_ansuran) Then
        If IsNumeric(rs!jumlah_ansuran) Then Frm87_LM_JUMLAH_BAYARAN_ANSURAN = rs!jumlah_ansuran 'Jumlah Bayran Ansuran
    End If
    If Not IsNull(rs!berat_diperoleh) Then 'Berat Diperolehi
        If IsNumeric(rs!berat_diperoleh) Then Frm87_LM_BERAT_DIPEROLEHI = rs!berat_diperoleh 'Berat DiPerolehi
    End If
    If Not IsNull(rs!JUMLAH_UPAH) Then
        If IsNumeric(rs!JUMLAH_UPAH) Then Frm87_LM_JUMLAH_BAYARAN_UPAH = rs!JUMLAH_UPAH
    End If
    
    DATA_FOUND = 1
    
    rs.Delete
    rs.Update
End If

rs.Close
Set rs = Nothing
'###Padam Bayaran Ansuran### - End

If DATA_FOUND = 1 Then
    
'### Update Senarai Ansuran ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 27_senarai_ansuran where ID='" & Frm87_LM_ID_ASAL & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!no_rujukan_pelanggan) Then Frm87_LM_No_CUST = rs!no_rujukan_pelanggan 'No. Rujukan Pelanggan
        If Not IsNull(rs!jenis_ansuran) Then
            If rs!jenis_ansuran = 0 Then '0 : Harga Semasa , 1 : Harga Tetap
                If Not IsNull(rs!jumlah_bayaran) Then
                    If IsNumeric(rs!jumlah_bayaran) Then Frm87_LM_JUMLAH_BAYARAN_ASAL = rs!jumlah_bayaran 'Jumlah Bayaran Terkumpul Asal (RM)
                    rs!jumlah_bayaran = Format(Frm87_LM_JUMLAH_BAYARAN_ASAL - Frm87_LM_JUMLAH_BAYARAN_ANSURAN, "0.00") 'Jumlah Bayaran Terkumpul (RM)
                End If
                If Not IsNull(rs!JUMLAH_BERAT) Then
                    If IsNumeric(rs!JUMLAH_BERAT) Then Frm87_LM_JUMLAH_BERAT_ASAL = rs!JUMLAH_BERAT 'Baki Berat Asal (g)
                    rs!JUMLAH_BERAT = Format(Frm87_LM_JUMLAH_BERAT_ASAL - Frm87_LM_BERAT_DIPEROLEHI, "0.00") 'Jumlah Berat Terkumpul (g)
                End If
                If Not IsNull(rs!BAKI_BERAT) Then
                    If IsNumeric(rs!BAKI_BERAT) Then Frm87_LM_BAKI_BERAT_ASAL = rs!BAKI_BERAT 'Baki Berat Asal (g)
                    rs!BAKI_BERAT = Format(Frm87_LM_BAKI_BERAT_ASAL + Frm87_LM_BERAT_DIPEROLEHI, "0.00") 'Baki Berat (g)
                End If
                If Not IsNull(rs!JUMLAH_UPAH) Then
                    If IsNumeric(rs!JUMLAH_UPAH) Then Frm87_LM_JUMLAH_UPAH_ASAL = rs!JUMLAH_UPAH 'Jumlah Upah Asal (RM)
                    rs!JUMLAH_UPAH = Format(Frm87_LM_JUMLAH_UPAH_ASAL - Frm87_LM_JUMLAH_BAYARAN_UPAH, "0.00") 'Baki Upah (RM)
                End If
                If Not IsNull(rs!baki_upah) Then
                    If IsNumeric(rs!baki_upah) Then Frm87_LM_BAKI_UPAH_ASAL = rs!baki_upah 'Baki Upah Asal (RM)
                    rs!baki_upah = Format(Frm87_LM_BAKI_UPAH_ASAL + Frm87_LM_JUMLAH_BAYARAN_UPAH, "0.00") 'Baki Upah (RM)
                End If
                
                rs.Update
            ElseIf rs!jenis_ansuran = 1 Then
                If Not IsNull(rs!jumlah_bayaran) Then
                    If IsNumeric(rs!jumlah_bayaran) Then Frm87_LM_JUMLAH_BAYARAN_ASAL = rs!jumlah_bayaran 'Jumlah Bayaran Terkumpul Asal (RM)
                    rs!jumlah_bayaran = Format(Frm87_LM_JUMLAH_BAYARAN_ASAL - Frm87_LM_JUMLAH_BAYARAN_ANSURAN, "0.00") 'Jumlah Bayaran Terkumpul (RM)
                End If
                If Not IsNull(rs!baki_bayaran) Then
                    If IsNumeric(rs!baki_bayaran) Then Frm87_LM_BAKI_BAYARAN_ASAL = rs!baki_bayaran 'Baki Bayaran Asal (RM)
                    rs!baki_bayaran = Format(Frm87_LM_JUMLAH_BAYARAN_ANSURAN + Frm87_LM_BAKI_BAYARAN_ASAL, "0.00") 'Baki Bayaran (RM)
                End If
                
                rs.Update
            End If
        End If

    End If
    
    rs.Close
    Set rs = Nothing
'### Update Senarai Ansuran ### - End

'###Padam Akaun Ansuran### - Start
    Frm87_LM_FLAG_SAVING = 0
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 29_akaun_ansuran where no_resit='" & Frm87_LM_No_RESIT & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!duit_simpanan_kedai) Then
            If Format(rs!duit_simpanan_kedai, "0.00") <> "0.00" Then
                If IsNumeric(rs!duit_simpanan_kedai) Then Frm87_LM_SIMPANAN_DIGUNAKAN = rs!duit_simpanan_kedai
                Frm87_LM_FLAG_SAVING = 1
            End If
        End If
        
        rs.Delete
        rs.Update
    End If
    
    rs.Close
    Set rs = Nothing
'###Padam Akaun Ansuran### - End

'###Update Simpanan Duit Di Kedai### - Start
    If Frm87_LM_FLAG_SAVING = 1 Then
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm87_LM_No_CUST & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!baki_simpanan) Then
                If IsNumeric(Frm87_LM_SIMPANAN_ASAL) Then Frm87_LM_SIMPANAN_ASAL = rs!baki_simpanan 'Baki Simpanan Pelanggan Ini (RM)
            End If
            
            rs!baki_simpanan = Format(Frm87_LM_SIMPANAN_ASAL + Frm87_LM_SIMPANAN_DIGUNAKAN, "0.00") 'Baki Simpanan Terkini Pelanggan Ini (RM)
            
            rs.Update
        End If
        
        rs.Close
        Set rs = Nothing
        
'###Padam Rekod Bayaran Dalam Table Simpanan### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 24_rekod_kewangan_pelanggan where NO_RESIT='" & Frm87_LM_No_RESIT & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            rs.Delete
            rs.Update
        End If
        
        rs.Close
        Set rs = Nothing
'###Padam Rekod Bayaran Dalam Table Simpanan### - End
        
    End If
'###Update Simpanan Duit Di Kedai### - End

End If
End Sub
Private Sub Tmr2_Timer()
'On Error Resume Next
If Frm87.CB13 = 1 And Frm87.TB8 <> vbNullString And Frm87.Tmr2.Enabled = True Then
    If Frm87.Tmr2.Interval = 100 Then
        If InStr(1, Frm87.TB8, "'") <> 0 Then
            MsgBox "No. Siri Produk Mengandungi Simbol Yang Tidak Sah , ['].", vbInformation, "Info"
            
            Frm87.TB8 = vbNullString
            Exit Sub
        End If
        
        Call Frm87_Call_Product_Detail
    End If
End If
End Sub

