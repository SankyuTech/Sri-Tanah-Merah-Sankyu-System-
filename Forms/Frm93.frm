VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm93 
   Caption         =   "Tempahan"
   ClientHeight    =   13035
   ClientLeft      =   120
   ClientTop       =   150
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
   Icon            =   "Frm93.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13035
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tempahan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10815
      Left            =   2160
      TabIndex        =   60
      Top             =   360
      Visible         =   0   'False
      Width           =   20175
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
         Left            =   7440
         Picture         =   "Frm93.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   156
         Top             =   4080
         Width           =   3615
      End
      Begin VB.ComboBox CBB3 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   360
         Left            =   8310
         Style           =   2  'Dropdown List
         TabIndex        =   152
         Top             =   5760
         Width           =   3795
      End
      Begin VB.CommandButton CMD12 
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
         MouseIcon       =   "Frm93.frx":3494
         MousePointer    =   99  'Custom
         Picture         =   "Frm93.frx":379E
         Style           =   1  'Graphical
         TabIndex        =   151
         Top             =   8040
         Width           =   2775
      End
      Begin VB.CommandButton CMD15 
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
         Left            =   12000
         MouseIcon       =   "Frm93.frx":5D68
         MousePointer    =   99  'Custom
         Picture         =   "Frm93.frx":6072
         Style           =   1  'Graphical
         TabIndex        =   150
         Top             =   8040
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CommandButton CMD14 
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
         Left            =   9120
         MouseIcon       =   "Frm93.frx":863C
         MousePointer    =   99  'Custom
         Picture         =   "Frm93.frx":8946
         Style           =   1  'Graphical
         TabIndex        =   149
         Top             =   8040
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CheckBox CB9 
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
         Left            =   7440
         TabIndex        =   146
         Top             =   6630
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Voucher Trade In"
         Height          =   2415
         Left            =   11400
         TabIndex        =   136
         Top             =   2760
         Width           =   4695
         Begin VB.CommandButton CMD10 
            Caption         =   "Batal"
            Height          =   375
            Left            =   1320
            MouseIcon       =   "Frm93.frx":AF10
            MousePointer    =   99  'Custom
            TabIndex        =   145
            ToolTipText     =   "Batalkan Urusan / Data Buyback (Trade In)"
            Top             =   1800
            Width           =   2055
         End
         Begin VB.TextBox TB17 
            BackColor       =   &H8000000A&
            Height          =   360
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   141
            Text            =   "TB17"
            Top             =   1395
            Width           =   1365
         End
         Begin VB.TextBox TB18 
            Alignment       =   2  'Center
            BackColor       =   &H8000000A&
            Height          =   360
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   138
            Text            =   "TB18"
            Top             =   375
            Width           =   1635
         End
         Begin VB.CommandButton CMD9 
            Caption         =   "Carian"
            Height          =   375
            Left            =   2880
            MouseIcon       =   "Frm93.frx":B21A
            MousePointer    =   99  'Custom
            TabIndex        =   137
            ToolTipText     =   "Carian Maklumat Terperinci Voucher Buyback / Trade In"
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label L15_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L15_Text"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   1440
            TabIndex        =   144
            Top             =   1200
            Width           =   1785
         End
         Begin VB.Label Label49 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Voucher :"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   120
            TabIndex        =   143
            Top             =   1200
            Width           =   1425
         End
         Begin VB.Label Label51 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Nilaian Voucher:RM"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   120
            TabIndex        =   142
            Top             =   1440
            Width           =   2715
         End
         Begin VB.Label Label53 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Maklumat Voucher Buyback / Trade In"
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
            Height          =   405
            Left            =   240
            TabIndex        =   140
            Top             =   840
            Width           =   3555
         End
         Begin VB.Label Label52 
            BackStyle       =   0  'Transparent
            Caption         =   "No.Voucher:"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   120
            TabIndex        =   139
            Top             =   390
            Width           =   2265
         End
      End
      Begin VB.CheckBox CB6 
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
         Left            =   11400
         TabIndex        =   134
         Top             =   2550
         Width           =   200
      End
      Begin VB.TextBox TB20 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   9360
         TabIndex        =   129
         Text            =   "TB20"
         Top             =   2880
         Width           =   1755
      End
      Begin VB.TextBox TB22 
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   128
         Text            =   "TB22"
         Top             =   3240
         Width           =   1755
      End
      Begin VB.TextBox TB23 
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   127
         Text            =   "TB23"
         Top             =   3600
         Width           =   1755
      End
      Begin VB.CommandButton CMD21 
         Caption         =   "Info Pembeli - (Berdaftar)"
         Enabled         =   0   'False
         Height          =   930
         Left            =   9240
         MouseIcon       =   "Frm93.frx":B524
         MousePointer    =   99  'Custom
         Picture         =   "Frm93.frx":B82E
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CommandButton CMD19 
         Caption         =   "Info Pembeli - (Tidak berdaftar)"
         Height          =   1050
         Left            =   9240
         MouseIcon       =   "Frm93.frx":DDF8
         MousePointer    =   99  'Custom
         Picture         =   "Frm93.frx":E102
         Style           =   1  'Graphical
         TabIndex        =   121
         Top             =   360
         Width           =   2415
      End
      Begin VB.Frame Frame6 
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
         Height          =   1575
         Left            =   7440
         TabIndex        =   114
         Top             =   240
         Width           =   1695
         Begin VB.CheckBox CB16 
            Enabled         =   0   'False
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
            Left            =   120
            TabIndex        =   119
            Top             =   975
            Width           =   200
         End
         Begin VB.CheckBox CB17 
            Enabled         =   0   'False
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
            Left            =   120
            TabIndex        =   118
            Top             =   1215
            Width           =   200
         End
         Begin VB.CheckBox CB13 
            Enabled         =   0   'False
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
            Left            =   120
            TabIndex        =   117
            Top             =   270
            Width           =   200
         End
         Begin VB.CheckBox CB14 
            Enabled         =   0   'False
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
            Left            =   120
            TabIndex        =   116
            Top             =   495
            Width           =   200
         End
         Begin VB.CheckBox CB15 
            Enabled         =   0   'False
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
            Left            =   120
            TabIndex        =   115
            Top             =   735
            Width           =   200
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Pelanggan                    Ahli Biasa                      Silver                           Gold                       Platinum"
            ForeColor       =   &H00000000&
            Height          =   1245
            Left            =   360
            TabIndex        =   120
            Top             =   240
            Width           =   2370
         End
      End
      Begin VB.Frame Frame3 
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
         Height          =   8895
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   7215
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFFFFF&
            Height          =   3735
            Left            =   600
            TabIndex        =   83
            Top             =   480
            Visible         =   0   'False
            Width           =   6850
            Begin VB.TextBox TB12 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Left            =   5385
               TabIndex        =   95
               Text            =   "TB12"
               Top             =   1440
               Width           =   1400
            End
            Begin VB.TextBox TB13 
               Alignment       =   2  'Center
               BackColor       =   &H8000000A&
               Height          =   360
               Left            =   5385
               Locked          =   -1  'True
               TabIndex        =   94
               Text            =   "TB13"
               Top             =   1800
               Width           =   1400
            End
            Begin VB.TextBox TB11 
               Alignment       =   2  'Center
               BackColor       =   &H8000000A&
               Height          =   360
               Left            =   2190
               Locked          =   -1  'True
               TabIndex        =   93
               Text            =   "TB11"
               Top             =   3270
               Width           =   1635
            End
            Begin VB.TextBox TB9 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Left            =   2190
               TabIndex        =   92
               Text            =   "TB9"
               Top             =   2550
               Width           =   1635
            End
            Begin VB.TextBox TB8 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Left            =   2190
               TabIndex        =   91
               Text            =   "TB8"
               Top             =   2190
               Width           =   1635
            End
            Begin VB.TextBox TB7 
               Alignment       =   2  'Center
               BackColor       =   &H8000000A&
               Height          =   360
               Left            =   2190
               Locked          =   -1  'True
               TabIndex        =   90
               Text            =   "TB7"
               Top             =   1830
               Width           =   1635
            End
            Begin VB.TextBox TB6 
               Alignment       =   2  'Center
               BackColor       =   &H8000000A&
               Height          =   360
               Left            =   2190
               Locked          =   -1  'True
               TabIndex        =   89
               Text            =   "TB6"
               Top             =   1470
               Width           =   1635
            End
            Begin VB.TextBox TB10 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Left            =   2190
               TabIndex        =   88
               Text            =   "TB10"
               Top             =   2910
               Width           =   1635
            End
            Begin VB.TextBox TB5 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Left            =   2145
               TabIndex        =   85
               Text            =   "TB5"
               Top             =   615
               Width           =   2460
            End
            Begin VB.CommandButton CMD5 
               Caption         =   "Carian Data"
               Height          =   375
               Left            =   4680
               MouseIcon       =   "Frm93.frx":F1CC
               MousePointer    =   99  'Custom
               TabIndex        =   84
               ToolTipText     =   "Carian Data Terperinci Produk"
               Top             =   600
               Width           =   1900
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "No. Siri Produk                                       Adjustment    RM"
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   240
               TabIndex        =   103
               Top             =   1485
               Width           =   5865
            End
            Begin VB.Label Label27 
               BackStyle       =   0  'Transparent
               Caption         =   "Kategori Produk       :"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   240
               TabIndex        =   102
               Top             =   1200
               Width           =   2385
            End
            Begin VB.Label L4_Text 
               BackStyle       =   0  'Transparent
               Caption         =   "L4_Text"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   2265
               TabIndex        =   101
               Top             =   1200
               Width           =   5835
            End
            Begin VB.Label Label23 
               BackStyle       =   0  'Transparent
               Caption         =   "Harga Asal           RM"
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   240
               TabIndex        =   100
               Top             =   3285
               Width           =   2265
            End
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "Harga Semasa   RM/g"
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   240
               TabIndex        =   99
               Top             =   2595
               Width           =   2265
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               Caption         =   "Berat Jualan            g"
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   240
               TabIndex        =   98
               Top             =   2220
               Width           =   2265
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "Berat Asal               g                             Harga Jualan  RM"
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   240
               TabIndex        =   97
               Top             =   1860
               Width           =   6465
            End
            Begin VB.Label Label21 
               BackStyle       =   0  'Transparent
               Caption         =   "Upah                   RM"
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   240
               TabIndex        =   96
               Top             =   2955
               Width           =   2265
            End
            Begin VB.Shape Shape1 
               Height          =   900
               Left            =   120
               Top             =   240
               Width           =   6645
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               Caption         =   "No. Siri Produk      :"
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   360
               TabIndex        =   87
               Top             =   645
               Width           =   1905
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Sila scan barang yang hendak ditempah oleh pembeli ini."
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   360
               TabIndex        =   86
               Top             =   315
               Width           =   5745
            End
         End
         Begin VB.TextBox TB33 
            BackColor       =   &H00FFFFFF&
            Height          =   1800
            Left            =   2280
            MultiLine       =   -1  'True
            TabIndex        =   104
            Text            =   "Frm93.frx":F4D6
            Top             =   4920
            Width           =   4485
         End
         Begin VB.Frame Frame4 
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
            Height          =   3735
            Left            =   240
            TabIndex        =   66
            Top             =   1080
            Visible         =   0   'False
            Width           =   6735
            Begin VB.ComboBox CBB1 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Left            =   2040
               Style           =   2  'Dropdown List
               TabIndex        =   76
               Top             =   840
               Width           =   4485
            End
            Begin VB.ComboBox CBB2 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Left            =   2040
               Style           =   2  'Dropdown List
               TabIndex        =   75
               Top             =   1200
               Width           =   4485
            End
            Begin VB.TextBox TB1 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Left            =   2040
               TabIndex        =   74
               Text            =   "TB1"
               Top             =   1560
               Width           =   4485
            End
            Begin VB.TextBox TB2 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Left            =   2040
               TabIndex        =   73
               Text            =   "TB2"
               Top             =   1920
               Width           =   4485
            End
            Begin VB.TextBox TB3 
               BackColor       =   &H00FFFFFF&
               Height          =   360
               Left            =   2040
               TabIndex        =   72
               Text            =   "TB3"
               Top             =   2280
               Width           =   4485
            End
            Begin VB.TextBox TB4 
               BackColor       =   &H8000000A&
               Height          =   360
               Left            =   2040
               Locked          =   -1  'True
               TabIndex        =   71
               Text            =   "TB4"
               Top             =   2640
               Width           =   4485
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
               Left            =   2040
               TabIndex        =   68
               Top             =   590
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
               Left            =   4080
               TabIndex        =   67
               Top             =   590
               Width           =   200
            End
            Begin VB.Label Label26 
               BackStyle       =   0  'Transparent
               Caption         =   "Kategori Produk *"
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   120
               TabIndex        =   82
               Top             =   870
               Width           =   2295
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Purity *"
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   120
               TabIndex        =   81
               Top             =   1230
               Width           =   2295
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Anggaran Berat g*"
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   120
               TabIndex        =   80
               Top             =   1590
               Width           =   1905
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Harga Semasa RM/g*"
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   120
               TabIndex        =   79
               Top             =   1950
               Width           =   1905
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Upah RM *"
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   120
               TabIndex        =   78
               Top             =   2310
               Width           =   1905
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Anggaran Harga RM *"
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   120
               TabIndex        =   77
               Top             =   2670
               Width           =   1905
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Sila masukkan maklumat terperinci tempahan barang baru."
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   120
               TabIndex        =   70
               Top             =   240
               Width           =   5865
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Barang Kemas             Barang Permata"
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   2300
               TabIndex        =   69
               Top             =   540
               Width           =   4305
            End
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
            Left            =   240
            TabIndex        =   63
            Top             =   600
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
            TabIndex        =   62
            Top             =   840
            Width           =   200
         End
         Begin VB.Label L17_Text 
            BackColor       =   &H8000000A&
            Caption         =   "L17_Text "
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            TabIndex        =   113
            Top             =   7680
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label L18_Text 
            BackColor       =   &H8000000A&
            Caption         =   "L18_Text"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            TabIndex        =   112
            Top             =   7920
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label L13_Text 
            BackColor       =   &H8000000A&
            Caption         =   "L13_Text "
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2040
            TabIndex        =   111
            Top             =   8280
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label L34_Text 
            Alignment       =   2  'Center
            BackColor       =   &H8000000A&
            Caption         =   "L34_Text"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2160
            TabIndex        =   110
            Top             =   6840
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Label L37_Text 
            BackColor       =   &H8000000A&
            Caption         =   "L37_Text"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3840
            TabIndex        =   109
            Top             =   7920
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label L19_Text 
            Alignment       =   2  'Center
            BackColor       =   &H8000000A&
            Caption         =   "L19_Text"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2160
            TabIndex        =   108
            Top             =   7080
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Label L20_Text 
            BackColor       =   &H8000000A&
            Caption         =   "L20_Text"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1320
            TabIndex        =   107
            Top             =   7560
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label L21_Text 
            BackColor       =   &H8000000A&
            Caption         =   "L21_Text"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1320
            TabIndex        =   106
            Top             =   7920
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label Label91 
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   360
            TabIndex        =   105
            Top             =   4950
            Width           =   1905
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Tempahan Barang Baru  Tempahan Barang Kedai"
            ForeColor       =   &H00000000&
            Height          =   525
            Left            =   480
            TabIndex        =   65
            Top             =   550
            Width           =   2385
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Sila buat pilihan jenis tempahan."
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
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   4785
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   8310
         TabIndex        =   153
         Top             =   5400
         Width           =   3795
         _ExtentX        =   6694
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
         Format          =   416743424
         CurrentDate     =   41561
      End
      Begin VB.Label Label88 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh  *"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   7440
         TabIndex        =   155
         Top             =   5400
         Width           =   1425
      End
      Begin VB.Label Label80 
         BackStyle       =   0  'Transparent
         Caption         =   "Pekerja *"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   7440
         TabIndex        =   154
         Top             =   5760
         Width           =   1575
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
         Left            =   7440
         TabIndex        =   148
         Top             =   6360
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm93.frx":F4DB
         ForeColor       =   &H000000FF&
         Height          =   780
         Left            =   7725
         TabIndex        =   147
         Top             =   6600
         Visible         =   0   'False
         Width           =   6330
      End
      Begin VB.Label Label54 
         BackStyle       =   0  'Transparent
         Caption         =   "Bayaran deposit dari barang trade in"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   11685
         TabIndex        =   135
         Top             =   2520
         Width           =   3690
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Maklumat Bayaran Deposit"
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
         TabIndex        =   133
         Top             =   2520
         Width           =   3615
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Caption         =   "Deposit *             RM"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7440
         TabIndex        =   132
         Top             =   2910
         Width           =   1905
      End
      Begin VB.Label Label56 
         BackStyle       =   0  'Transparent
         Caption         =   "Deposit Trade In* RM"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7440
         TabIndex        =   131
         Top             =   3270
         Width           =   1905
      End
      Begin VB.Label Label57 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Deposit *  RM"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7440
         TabIndex        =   130
         Top             =   3630
         Width           =   1905
      End
      Begin VB.Shape Shape4 
         Height          =   2175
         Left            =   9170
         Top             =   240
         Width           =   10455
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama   :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   11760
         TabIndex        =   126
         Top             =   315
         Width           =   825
      End
      Begin VB.Label L35_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L35_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   12600
         TabIndex        =   125
         Top             =   315
         Width           =   5625
      End
      Begin VB.Label Label75 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama   :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   11760
         TabIndex        =   124
         Top             =   1365
         Width           =   825
      End
      Begin VB.Label L36_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L36_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   12600
         TabIndex        =   123
         Top             =   1365
         Width           =   5625
      End
   End
   Begin VB.PictureBox Pic4 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   15000
      ScaleHeight     =   3615
      ScaleWidth      =   7695
      TabIndex        =   15
      Top             =   840
      Visible         =   0   'False
      Width           =   7695
      Begin VB.CommandButton CMD4 
         Caption         =   "Carian Maklumat pelanggan"
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
         MouseIcon       =   "Frm93.frx":F5C7
         MousePointer    =   99  'Custom
         Picture         =   "Frm93.frx":F8D1
         Style           =   1  'Graphical
         TabIndex        =   185
         Top             =   2520
         Width           =   2985
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
         Left            =   2280
         TabIndex        =   20
         Top             =   1680
         Width           =   200
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
         Left            =   360
         TabIndex        =   19
         Top             =   1680
         Width           =   200
      End
      Begin VB.TextBox TB41 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2400
         TabIndex        =   18
         Text            =   "TB41"
         Top             =   2130
         Width           =   2955
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Keahlian               No. Kad Pengenalan"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   600
         TabIndex        =   21
         Top             =   1640
         Width           =   6075
      End
      Begin VB.Label L5_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombor keahlian     :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   480
         TabIndex        =   17
         Top             =   2160
         Width           =   1875
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm93.frx":1099B
         ForeColor       =   &H00000000&
         Height          =   1485
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   6435
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tetapan Paparan Senarai Tempahan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   5040
      TabIndex        =   157
      Top             =   1560
      Visible         =   0   'False
      Width           =   8415
      Begin VB.ComboBox CBB9 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   182
         Top             =   3600
         Width           =   5445
      End
      Begin VB.CommandButton CMD13 
         Caption         =   "Carian Senarai Tempahan"
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
         Left            =   3000
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm93.frx":10B10
         MousePointer    =   99  'Custom
         Picture         =   "Frm93.frx":10E1A
         Style           =   1  'Graphical
         TabIndex        =   181
         Top             =   4080
         Width           =   2865
      End
      Begin VB.ComboBox CBB4 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   163
         Top             =   720
         Width           =   5445
      End
      Begin VB.ComboBox CBB5 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   162
         Top             =   2160
         Width           =   5445
      End
      Begin VB.ComboBox CBB7 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   161
         Top             =   2520
         Width           =   5445
      End
      Begin VB.ComboBox CBB8 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   160
         Top             =   2880
         Width           =   5445
      End
      Begin VB.TextBox TB30 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2160
         TabIndex        =   159
         Text            =   "TB30"
         Top             =   3240
         Width           =   5445
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   360
         Left            =   2160
         TabIndex        =   164
         Top             =   1080
         Width           =   5445
         _ExtentX        =   9604
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
         Format          =   416874496
         CurrentDate     =   41561
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   360
         Left            =   2160
         TabIndex        =   165
         Top             =   1440
         Width           =   5445
         _ExtentX        =   9604
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
         Format          =   416874496
         CurrentDate     =   41561
      End
      Begin VB.Label L45_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L45_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7680
         TabIndex        =   184
         Top             =   3240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Cawangan *"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   183
         Top             =   3630
         Width           =   2295
      End
      Begin VB.Label L27_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L27_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7680
         TabIndex        =   180
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label L28_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L28_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7680
         TabIndex        =   179
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label L29_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L29_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7680
         TabIndex        =   178
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label L32_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L32_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7680
         TabIndex        =   177
         Top             =   2520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label L31_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L31_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7680
         TabIndex        =   176
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label L30_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L30_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7680
         TabIndex        =   175
         Top             =   1800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label L38_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L38_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7680
         TabIndex        =   174
         Top             =   2880
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label90 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh *"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   173
         Top             =   750
         Width           =   2295
      End
      Begin VB.Label Label93 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Tempahan *"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   172
         Top             =   2190
         Width           =   2295
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Status *"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   171
         Top             =   2550
         Width           =   2295
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Lain-lain *"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   170
         Top             =   2910
         Width           =   2295
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh mula  *"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   169
         Top             =   1080
         Width           =   2385
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh akhir *"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   168
         Top             =   1440
         Width           =   2385
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh adalah tarikh bayaran deposit."
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
         Left            =   2160
         TabIndex        =   167
         Top             =   1800
         Width           =   5505
      End
      Begin VB.Label L22_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L22_Text"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   166
         Top             =   3240
         Width           =   2715
      End
      Begin VB.Label Label92 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila buat pilihan di bawah bagi paparan senarai tempahan."
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
         Left            =   120
         TabIndex        =   158
         Top             =   360
         Width           =   8145
      End
   End
   Begin VB.PictureBox Pic1 
      BorderStyle     =   0  'None
      Height          =   11055
      Left            =   11400
      ScaleHeight     =   11055
      ScaleWidth      =   23535
      TabIndex        =   4
      Top             =   -3960
      Visible         =   0   'False
      Width           =   23535
      Begin VB.PictureBox Pic3 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   15600
         ScaleHeight     =   3615
         ScaleWidth      =   6855
         TabIndex        =   9
         Top             =   3480
         Visible         =   0   'False
         Width           =   6855
      End
      Begin VB.PictureBox Pic6 
         BorderStyle     =   0  'None
         Height          =   4455
         Left            =   7200
         ScaleHeight     =   4455
         ScaleWidth      =   8055
         TabIndex        =   22
         Top             =   5520
         Width           =   8055
         Begin VB.ComboBox CBB6 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   1920
            Width           =   4965
         End
         Begin VB.TextBox TB21 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   6480
            TabIndex        =   27
            Text            =   "TB21"
            Top             =   840
            Width           =   1260
         End
         Begin VB.TextBox TB32 
            BackColor       =   &H80000000&
            Height          =   360
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   26
            Text            =   "TB32"
            Top             =   3960
            Width           =   2100
         End
         Begin VB.TextBox TB29 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   2760
            TabIndex        =   25
            Text            =   "TB29"
            Top             =   1545
            Width           =   1260
         End
         Begin VB.TextBox TB28 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   2760
            TabIndex        =   24
            Text            =   "TB28"
            Top             =   930
            Width           =   1260
         End
         Begin VB.TextBox TB27 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   2760
            TabIndex        =   23
            Text            =   "TB27"
            Top             =   555
            Width           =   1260
         End
         Begin VB.Shape Shape3 
            BorderWidth     =   2
            Height          =   4335
            Left            =   120
            Top             =   120
            Width           =   7815
         End
         Begin VB.Label L41_Text 
            BackColor       =   &H8000000A&
            Caption         =   "L41_Text"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3240
            TabIndex        =   45
            Top             =   2325
            Width           =   4080
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "Jenis kad                        :"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   360
            TabIndex        =   44
            Top             =   1965
            Width           =   2475
         End
         Begin VB.Label Label48 
            BackStyle       =   0  'Transparent
            Caption         =   "Caj perkhidmatan               :  %"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   360
            TabIndex        =   43
            Top             =   2325
            Width           =   2835
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah caj perkhidmatan     : RM"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   360
            TabIndex        =   42
            Top             =   2685
            Width           =   2955
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "Cukai GST caj perkhidmatan: RM"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   360
            TabIndex        =   41
            Top             =   3045
            Width           =   2955
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah potongan kad          : RM"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   360
            TabIndex        =   40
            Top             =   3405
            Width           =   2955
         End
         Begin VB.Label L44_Text 
            BackColor       =   &H8000000A&
            Caption         =   "L44_Text"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3240
            TabIndex        =   39
            Top             =   3405
            Width           =   4080
         End
         Begin VB.Label L43_Text 
            BackColor       =   &H8000000A&
            Caption         =   "L43_Text"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3240
            TabIndex        =   38
            Top             =   3045
            Width           =   4080
         End
         Begin VB.Label L42_Text 
            BackColor       =   &H8000000A&
            Caption         =   "L42_Text"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3240
            TabIndex        =   37
            Top             =   2685
            Width           =   4080
         End
         Begin VB.Label L26_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L26_Text"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5880
            TabIndex        =   35
            Top             =   600
            Width           =   1635
         End
         Begin VB.Label Label83 
            BackStyle       =   0  'Transparent
            Caption         =   "Duit Simpanan Di Kedai RM:"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   4080
            TabIndex        =   34
            Top             =   885
            Width           =   2715
         End
         Begin VB.Label Label81 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Bayaran          RM :"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   360
            TabIndex        =   32
            Top             =   3990
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
            TabIndex        =   31
            Top             =   240
            Width           =   3585
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Kad Kredit                 RM :"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   360
            TabIndex        =   30
            Top             =   1545
            Width           =   2715
         End
         Begin VB.Label Label72 
            BackStyle       =   0  'Transparent
            Caption         =   "Bank In                     RM :"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   360
            TabIndex        =   29
            Top             =   960
            Width           =   2715
         End
         Begin VB.Label Label70 
            BackStyle       =   0  'Transparent
            Caption         =   "Tunai                        RM :"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   360
            TabIndex        =   28
            Top             =   585
            Width           =   2715
         End
         Begin VB.Shape Shape10 
            BorderWidth     =   2
            Height          =   2415
            Left            =   240
            Top             =   1440
            Width           =   7575
         End
         Begin VB.Label Label82 
            BackStyle       =   0  'Transparent
            Caption         =   "Baki simpanan : RM"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4080
            TabIndex        =   33
            Top             =   600
            Width           =   3435
         End
      End
      Begin VB.PictureBox Pic5 
         BorderStyle     =   0  'None
         Height          =   2175
         Left            =   5880
         ScaleHeight     =   2175
         ScaleWidth      =   4065
         TabIndex        =   11
         Top             =   1440
         Width           =   4065
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
         Left            =   5280
         TabIndex        =   5
         Top             =   300
         Width           =   200
      End
      Begin VB.PictureBox Pic2 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   3255
         Left            =   16440
         ScaleHeight     =   3255
         ScaleWidth      =   6855
         TabIndex        =   8
         Top             =   7320
         Visible         =   0   'False
         Width           =   6855
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Kategori Pembeli"
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
         Left            =   7080
         TabIndex        =   14
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Maklumat Pembeli"
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
         Left            =   9480
         TabIndex        =   13
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "Scanner Mode"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   5595
         TabIndex        =   6
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Pendaftaran bagi tempahan barang kemas."
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
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   8145
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Senarai Tempahan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11295
      Left            =   4200
      TabIndex        =   47
      Top             =   1560
      Visible         =   0   'False
      Width           =   20535
      Begin VB.CommandButton CMD25 
         Caption         =   "Back"
         Height          =   810
         Left            =   18000
         MouseIcon       =   "Frm93.frx":11EE4
         MousePointer    =   99  'Custom
         Picture         =   "Frm93.frx":121EE
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Tutup senarai ini."
         Top             =   10440
         Width           =   1095
      End
      Begin VB.CommandButton CMD26 
         Caption         =   "Next"
         Height          =   810
         Left            =   19200
         MouseIcon       =   "Frm93.frx":132B8
         MousePointer    =   99  'Custom
         Picture         =   "Frm93.frx":135C2
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Tutup senarai ini."
         Top             =   10440
         Width           =   1095
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   9900
         Left            =   120
         TabIndex        =   48
         Top             =   480
         Width           =   20235
         _ExtentX        =   35692
         _ExtentY        =   17463
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
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L60_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   16995
         TabIndex        =   58
         Top             =   10440
         Width           =   465
      End
      Begin VB.Label L61_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L61_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   17580
         TabIndex        =   57
         Top             =   10440
         Width           =   705
      End
      Begin VB.Label L62_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L62_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   11760
         TabIndex        =   56
         Top             =   10440
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label L63_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L63_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   11760
         TabIndex        =   55
         Top             =   10800
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label L33_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L33_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2160
         TabIndex        =   54
         Top             =   10440
         Width           =   1185
      End
      Begin VB.Label L39_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L39_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   5040
         TabIndex        =   53
         Top             =   10440
         Width           =   1185
      End
      Begin VB.Label L40_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L40_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7920
         TabIndex        =   52
         Top             =   10440
         Width           =   1185
      End
      Begin VB.Label L25_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai tempahan yang belum siap."
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
         TabIndex        =   49
         Top             =   240
         Width           =   17625
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "Paparan Muka :       /"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   15720
         TabIndex        =   59
         Top             =   10440
         Width           =   2505
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan Tempahan :                      Berat Belum Siap :                              Berat Siap :"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   240
         TabIndex        =   186
         Top             =   10440
         Width           =   12705
      End
   End
   Begin VB.Timer Tmr2 
      Interval        =   100
      Left            =   0
      Top             =   480
   End
   Begin VB.Timer Tmr3 
      Interval        =   100
      Left            =   0
      Top             =   960
   End
   Begin VB.Timer Tmr1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Pic7 
      BorderStyle     =   0  'None
      Height          =   11715
      Left            =   1440
      ScaleHeight     =   11715
      ScaleWidth      =   21255
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   21255
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   10005
         Left            =   0
         TabIndex        =   46
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   360
         Width           =   21105
         _ExtentX        =   37227
         _ExtentY        =   17648
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
   End
   Begin VB.Label L12_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Keluar"
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
      Left            =   21360
      MouseIcon       =   "Frm93.frx":1468C
      MousePointer    =   99  'Custom
      TabIndex        =   10
      ToolTipText     =   "Keluar Ke Menu Sebelum"
      Top             =   2160
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label L10_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai Tempahan"
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
      Left            =   2760
      MouseIcon       =   "Frm93.frx":14996
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label L3_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pendaftaran Tempahan"
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
      Left            =   360
      MouseIcon       =   "Frm93.frx":14CA0
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   120
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
      Left            =   21705
      TabIndex        =   1
      Top             =   1635
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
      Left            =   21720
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Menu Frm93_PM_Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm93_SM_Edit 
         Caption         =   "Lihat Data / Edit Data"
      End
      Begin VB.Menu Frm93_SM_belum_siap 
         Caption         =   "Tukar Status Ke Belum Siap"
      End
      Begin VB.Menu frm93_sm_bar1 
         Caption         =   "-"
      End
      Begin VB.Menu Frm93_SM_padam 
         Caption         =   "Padam Data / Batal Tempahan"
      End
      Begin VB.Menu frm93_sm_bar2 
         Caption         =   "-"
      End
      Begin VB.Menu Frm93_SM_Cetak_Invoice_Deposit 
         Caption         =   "Cetak Invoice Deposit"
      End
      Begin VB.Menu Frm93_SM_Cetak_Invoice_Siap 
         Caption         =   "Cetak Invoice Jelas (Siap)"
      End
      Begin VB.Menu frm93_sm_bar3 
         Caption         =   "-"
      End
      Begin VB.Menu Frm93_SM_Siap 
         Caption         =   "Ambilan Barang / Barang Siap"
         Begin VB.Menu Frm93_SM_harga_tempahan 
            Caption         =   "Harga Emas Mengikut Harga Tempahan"
         End
         Begin VB.Menu Frm93_SM_harga_semasa 
            Caption         =   "Harga Emas Mengikut Harga Semasa"
         End
      End
   End
End
Attribute VB_Name = "Frm93"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CB13_Click()
'on error resume next
If Frm93.CB13 = 1 Then
    Frm93.L37_Text = 1
End If
End Sub
Private Sub CB14_Click()
'on error resume next
If Frm93.CB14 = 1 Then
    Frm93.L37_Text = 2
End If
End Sub
Private Sub CB15_Click()
'on error resume next
If Frm93.CB15 = 1 Then
    Frm93.L37_Text = 4
End If
End Sub
Private Sub CB16_Click()
'on error resume next
If Frm93.CB16 = 1 Then
    Frm93.L37_Text = 3
    
    'Frm93.CB21.Enabled = False
    'Frm93.CB22.Enabled = False
    'Frm93.CB25.Enabled = False
    
    'Frm93.CB21 = 1
End If
End Sub
Private Sub CB17_Click()
'on error resume next
If Frm93.CB17 = 1 Then
    Frm93.L37_Text = 5
End If
End Sub
Private Sub CB19_Click()
'on error resume next
If Frm93.CB19 = 1 And GLOBAL_DISABLE = 0 Then
    Frm93.CB20 = 0
    
    Frm93.TB1 = "0.00" 'Anggaran Berat
    Frm93.TB3 = "0.00" 'Upah
    Frm93.TB2 = "0.00" 'Harga Semasa

    Frm93.TB1.Locked = False
    Frm93.TB2.Locked = False
    'Frm93.TB3.Locked = False
    Frm93.TB4.Locked = True
    
    Frm93.TB1.BackColor = &HFFFFFF
    Frm93.TB2.BackColor = &HFFFFFF
    'Frm93.TB3.BackColor = &HFFFFFF
    Frm93.TB4.BackColor = &H8000000A
End If
End Sub
Private Sub CB2_Click()
'on error resume next
If Frm93.CB2 = 1 Then
    Frm93.CB3 = 0
    
    Frm93.CB19 = 1
    
    Frm93.Frame4.Visible = True
    Frm93.Frame5.Visible = False
End If
End Sub
Private Sub CB20_Click()
'on error resume next
If Frm93.CB20 = 1 And GLOBAL_DISABLE = 0 Then
    Frm93.CB19 = 0
    
    Frm93.TB1 = vbNullString 'Anggaran Berat
    Frm93.TB2 = vbNullString 'Harga Semasa
    Frm93.TB3 = "0.00" 'Upah
    
    Frm93.TB1.Locked = True
    Frm93.TB2.Locked = True
    'Frm93.TB3.Locked = True
    Frm93.TB4.Locked = False
    
    Frm93.TB1.BackColor = &H8000000A
    Frm93.TB2.BackColor = &H8000000A
    'Frm93.TB3.BackColor = &H8000000A
    Frm93.TB4.BackColor = &HFFFFFF
End If
End Sub
Private Sub CB3_Click()
'on error resume next
If Frm93.CB3 = 1 Then
    Frm93.CB2 = 0
    
    Frm93.Frame5.Visible = True
    Frm93.Frame4.Visible = False
    
    If GLOBAL_DISABLE = 0 Then Frm93.TB5.SetFocus
End If
End Sub
Private Sub CB4_Click()
'on error resume next
If Frm93.CB4 = 1 Then
    Frm93.CB5 = 0
    Frm93.L5_Text = "Nombor keahlian     :"
    Frm93.TB41.SetFocus
End If
End Sub
Private Sub CB5_Click()
'on error resume next
If Frm93.CB5 = 1 Then
    Frm93.CB4 = 0
    Frm93.L5_Text = "No. Kad Pengenalan :"
    Frm93.TB41.SetFocus
End If
End Sub
Private Sub CB6_Click()
'on error resume next
If Frm93.CB6 = 1 Then
    Frm93.TB18.Locked = False
    Frm93.TB18.BackColor = &HFFFFFF
    
    If GLOBAL_DISABLE = 0 Then
        Frm93.TB18.SetFocus
    End If
Else
    Frm93.TB18.Locked = True
    Frm93.TB18.BackColor = &H8000000A
    
    Frm93.TB18 = vbNullString
    Frm93.L15_Text = vbNullString
    Frm93.TB17 = "0.00"
End If
End Sub



Private Sub CBB2_Click()
'On Error Resume Next
If GLOBAL_DISABLE = 0 Then
    
    If Frm93.CB19 = 1 Then
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from hargaemas where Purity='" & Frm93.CBB2 & "' AND cawangan='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Frm93.CB13 = 1 Then
                If IsNumeric(rs!Harga_Pelanggan) Then Frm93.TB2 = Format(rs!Harga_Pelanggan, "0.00") 'Harga Semasa Bagi Pelanggan (RM/g)
            ElseIf Frm93.CB14 = 1 Then
                If IsNumeric(rs!Harga_Member) Then Frm93.TB2 = Format(rs!Harga_Member, "0.00") 'Harga Semasa Bagi Member (RM/g)
            ElseIf Frm93.CB15 = 1 Then
                If IsNumeric(rs!Harga_Pengedar) Then Frm93.TB2 = Format(rs!Harga_Pengedar, "0.00") 'Harga Semasa Bagi Pengedar (RM/g)
            ElseIf Frm93.CB16 = 1 Then
                If IsNumeric(rs!Harga_RAF) Then Frm93.TB2 = Format(rs!Harga_RAF, "0.00") 'Harga Semasa Bagi RAF (RM/g)
            ElseIf Frm93.CB17 = 1 Then
                If IsNumeric(rs!harga_nd) Then Frm93.TB2 = Format(rs!harga_nd, "0.00") 'Harga Semasa Bagi Normal Dealer (RM/g)
            'ElseIf Frm93.CB18 = 1 Then
            '    If IsNumeric(rs!harga_md) Then Frm93.TB2 = Format(rs!harga_md, "0.00") 'Harga Semasa Bagi Master Dealer (RM/g)
            End If
        End If
        
        rs.Close
        Set rs = Nothing
    Else
        Frm93.TB2 = vbNullString
    End If
End If
End Sub

Private Sub CBB4_Change()
'on error resume next
If Frm93.CBB4 = "Tiada filter tarikh" Then
    Frm93.DTPicker2.Enabled = False
    Frm93.DTPicker3.Enabled = False
Else
    Frm93.DTPicker2.Enabled = True
    Frm93.DTPicker3.Enabled = True
End If
End Sub

Private Sub CBB4_Click()
'on error resume next
If Frm93.CBB4 = "Tiada filter tarikh" Then
    Frm93.DTPicker2.Enabled = False
    Frm93.DTPicker3.Enabled = False
Else
    Frm93.DTPicker2.Enabled = True
    Frm93.DTPicker3.Enabled = True
End If
End Sub

Private Sub CBB6_Change()
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

Private Sub CBB6_Click()
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

Private Sub CBB8_Change()
'on error resume next
If Frm93.CBB8 = "-" Then

    Frm93.L22_Text.Visible = False
    Frm93.TB30.Visible = False
    
ElseIf Frm93.CBB8 = "No. siri produk" Then

    Frm93.L22_Text.Visible = True
    Frm93.L22_Text.Caption = "No. Siri *"
    Frm93.TB30.Visible = True
    Frm93.TB30.SetFocus
    Frm93.TB30 = vbNullString

ElseIf Frm93.CBB8 = "No. invoice deposit" Then

    Frm93.L22_Text.Visible = True
    Frm93.L22_Text.Caption = "Invoice deposit *"
    Frm93.TB30.Visible = True
    Frm93.TB30.SetFocus
    Frm93.TB30 = vbNullString

ElseIf Frm93.CBB8 = "No. invoice ambilan barang" Then

    Frm93.L22_Text.Visible = True
    Frm93.L22_Text.Caption = "Invoice siap *"
    Frm93.TB30.Visible = True
    Frm93.TB30.SetFocus
    Frm93.TB30 = vbNullString

End If
End Sub

Private Sub CBB8_Click()
'on error resume next
If Frm93.CBB8 = "-" Then

    Frm93.L22_Text.Visible = False
    Frm93.TB30.Visible = False
    
ElseIf Frm93.CBB8 = "No. siri produk" Then

    Frm93.L22_Text.Visible = True
    Frm93.L22_Text.Caption = "No. Siri *"
    Frm93.TB30.Visible = True
    Frm93.TB30.SetFocus
    Frm93.TB30 = vbNullString

ElseIf Frm93.CBB8 = "No. invoice deposit" Then

    Frm93.L22_Text.Visible = True
    Frm93.L22_Text.Caption = "Invoice deposit *"
    Frm93.TB30.Visible = True
    Frm93.TB30.SetFocus
    Frm93.TB30 = vbNullString

ElseIf Frm93.CBB8 = "No. invoice ambilan barang" Then

    Frm93.L22_Text.Visible = True
    Frm93.L22_Text.Caption = "Invoice siap *"
    Frm93.TB30.Visible = True
    Frm93.TB30.SetFocus
    Frm93.TB30 = vbNullString

End If
End Sub

Private Sub CMD10_Click()
'on error resume next
If Frm93.L15_Text <> vbNullString Then
    Note = "Adakah anda ingin batalkan deposit dengan trade in ini?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then

        Frm93.L15_Text = vbNullString
        Frm93.TB17 = "0.00"

        Frm93.L35_Text = vbNullString
        Frm93.L36_Text = vbNullString
    End If
Else
    MsgBox "Tiada maklumat berkenaan no. voucher ini.", vbInformation, "Info"
End If
End Sub
Private Sub CMD12_Click()
'On Error Resume Next
Dim Err(40)
Dim Frm93_LM_ERR_BERAT_ASAL As Double
Dim Frm93_LM_ERR_BERAT_JUALAN As Double
Dim Frm93_LM_ERR_JUMLAH_BAYARAN As Double
Dim Frm93_LM_ERR_HARGA As Double
Dim Frm93_LM_JUMLAH_SIMPANAN As Double
Dim Frm93_LM_GUNA_SIMPAN As Double

Frm93_LM_ERR_BERAT_ASAL = 0
Frm93_LM_ERR_BERAT_JUALAN = 0
Frm93_LM_ERR_JUMLAH_BAYARAN = 0 'Jumlah Bayaran
Frm93_LM_ERR_HARGA = 0 'Jumlah Perlu Bayar
Frm93_LM_Flag_TRADE_IN = 0 '0 : Tiada Urusan Trade In , 1 : Ada Urusan Trade In
Frm93_LM_Flag_SIMPANAN = 0 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
Frm93_LM_JUMLAH_SIMPANAN = 0  'Jumlah Simpanan Yang Ada
Frm93_LM_GUNA_SIMPAN = 0 'Jumlah Simpanan Yang Hendak Digunakan
Frm93_LM_KATEGORI = 0
'1:  Pelanggan Biasa
'2 : Member / Ahli
'3:  Pengedar
'4:  RAF
'5:  Normal Dealer(ND)
'6:  Master Dealer(MD)

If Frm93.CB2 = 0 And Frm93.CB3 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan jenis tempahan."
End If
If Frm93.CB2 = 1 Then 'Jenis Tempahan : Tempahan Barang Baru
    If Frm93.CB19 = 0 And Frm93.CB20 = 0 Then
        x = x + 1
        Err(x) = "Sila buat pilihan jenis barang samada barang kemas atau barang permata."
    End If
    If Frm93.CBB1 = vbNullString Then
        x = x + 1
        Err(x) = "Sila pilih kategori produk."
    End If
    If Frm93.CBB2 = vbNullString Then
        x = x + 1
        Err(x) = "Sila pilih purity."
    End If
    
    If Frm93.CB19 = 1 Then 'Jenis Barang : Barang Kemas
        If Frm93.TB1 = vbNullString Or (Frm93.TB1 <> vbNullString And Not IsNumeric(Frm93.TB1)) Then
            x = x + 1
            Err(x) = "Sila masukkan [Anggaran Berat (g)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
        End If
        If Frm93.TB2 = vbNullString Or (Frm93.TB2 <> vbNullString And Not IsNumeric(Frm93.TB2)) Then
            x = x + 1
            Err(x) = "Sila masukkan [Harga Semasa (RM/g)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
        End If
        If Frm93.TB3 = vbNullString Or (Frm93.TB3 <> vbNullString And Not IsNumeric(Frm93.TB3)) Then
            x = x + 1
            Err(x) = "Sila masukkan [Upah (RM)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
        End If
    End If
    If Frm93.CB20 = 1 Then 'Jenis Barang : Barang Permata
        If Frm93.TB3 = vbNullString Or (Frm93.TB3 <> vbNullString And Not IsNumeric(Frm93.TB3)) Then
            x = x + 1
            Err(x) = "Sila masukkan [Upah (RM)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
        End If
        If Frm93.TB4 = vbNullString Or (Frm93.TB4 <> vbNullString And Not IsNumeric(Frm93.TB4)) Then
            x = x + 1
            Err(x) = "Sila masukkan [Anggaran Harga (RM)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
        End If
    End If
End If
If Frm93.CB3 = 1 Then 'Jenis Tempahan : Tempahan Barang Kedai
    If Frm93.L4_Text = vbNullString Then
        x = x + 1
        Err(x) = "Tiada maklumat kategori produk"
    End If
    If Frm93.L13_Text = 0 Then 'Jenis Barang : Barang Kemas
    
        If (Frm93.TB7 <> vbNullString And IsNumeric(Frm93.TB7)) And (Frm93.TB8 <> vbNullString And IsNumeric(Frm93.TB8)) Then
            Frm93_LM_ERR_BERAT_ASAL = Frm93.TB7 'Berat Asal
            Frm93_LM_ERR_BERAT_JUALAN = Frm93.TB8 'Berat Jualan
            
            If Frm93_LM_ERR_BERAT_JUALAN > Frm93_LM_ERR_BERAT_ASAL Then
                x = x + 1
                Err(x) = "Berat jualan melebihi berat asal"
            End If
        End If
        
        If Frm93.TB8 = vbNullString Or (Frm93.TB8 <> vbNullString And Not IsNumeric(Frm93.TB8)) Then
            x = x + 1
            Err(x) = "Sila masukkan [Berat Jualan (g)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
        End If
        If Frm93.TB9 = vbNullString Or (Frm93.TB9 <> vbNullString And Not IsNumeric(Frm93.TB9)) Then
            x = x + 1
            Err(x) = "Sila masukkan [Harga Semasa (RM/g)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
        End If
    ElseIf Frm93.L13_Text = 1 Then 'Jenis Barang : Barang Permata
        If Frm93.TB11 = vbNullString Or (Frm93.TB11 <> vbNullString And Not IsNumeric(Frm93.TB11)) Then
            x = x + 1
            Err(x) = "Sila masukkan [Harga Asal (RM)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
        End If
    End If
    If Frm93.TB10 = vbNullString Or (Frm93.TB10 <> vbNullString And Not IsNumeric(Frm93.TB10)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Upah (RM)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm93.TB12 = vbNullString Or (Frm93.TB12 <> vbNullString And Not IsNumeric(Frm93.TB12)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Adjustment (RM)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If
If Frm93.L35_Text = vbNullString And Frm93.L36_Text = vbNullString Then
    x = x + 1
    Err(x) = "Tiada maklumat pembeli."
End If
If Frm93.TB20 = vbNullString Or (Frm93.TB20 <> vbNullString And Not IsNumeric(Frm93.TB20)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Deposit (RM)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm93.CBB3 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih nama pekerja."
End If
If Frm93.CB6 = 1 Then 'Bayaran Deposit Dari Barang Trade In
    If Frm93.L15_Text = vbNullString Then
        x = x + 1
        Err(x) = "Tiada maklumat No. Voucher Trade In."
    End If
    If Frm93.TB17 = vbNullString Then
        x = x + 1
        Err(x) = "Tiada maklumat nilaian Voucher Trade In."
    End If
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
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara Duit Simpanan Di Kedai. Sila masukkan 0 jika tiada bayaran secara simpanan di kedai."
End If
If (frm130.TB21 <> vbNullString And IsNumeric(frm130.TB21)) And (frm130.L26_Text <> vbNullString And IsNumeric(frm130.L26_Text)) Then
    Frm93_LM_JUMLAH_SIMPANAN = frm130.L26_Text  'Jumlah Simpanan Yang Ada
    Frm93_LM_GUNA_SIMPAN = frm130.TB21  'Jumlah Simpanan Yang Hendak Digunakan
    
    If Frm93_LM_GUNA_SIMPAN > Frm93_LM_JUMLAH_SIMPANAN Then
        x = x + 1
        Err(x) = "Jumlah simpanan yang hendak digunakan melebihi simpanan yang ada."
    End If
End If
If (frm130.TB32 <> vbNullString And IsNumeric(frm130.TB32)) And (Frm93.TB20 <> vbNullString And IsNumeric(Frm93.TB20)) Then

    Frm93_LM_ERR_JUMLAH_BAYARAN = frm130.TB32 'Jumlah Bayaran
    Frm93_LM_ERR_HARGA = Frm93.TB20 'Jumlah Perlu Bayar
    
    If Frm93_LM_ERR_JUMLAH_BAYARAN <> Frm93_LM_ERR_HARGA Then
        x = x + 1
        Err(x) = "Jumlah bayaran tidak sama dengan jumlah perlu bayar."
    End If
End If

'### Periksa Samada Maklumat Pembeli Diisi Dalam Kedua-dua Ruangan Berdaftar Dan Tidak Berdaftar ### - Start
If x = 0 Then

    If Frm93.L35_Text <> vbNullString And Frm93.L36_Text <> vbNullString Then
    
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
    
    If Frm93.L35_Text <> vbNullString And Frm93.L36_Text = vbNullString Then
        Note = "Adakah anda yakin untuk teruskan urusan jualan ini ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Data jualan akan disimpan ke dalam sistem." & vbCrLf & _
                vbNullString & vbCrLf & _
                "*** Oleh ini pembeli ini tidak berdaftar dengan sistem, oleh itu data dan rekod pembelian pembeli ini tidak akan disimpan di dalam sistem ***"
    End If
    If Frm93.L35_Text = vbNullString And Frm93.L36_Text <> vbNullString Then
        Note = "Adakah anda yakin untuk teruskan urusan jualan ini ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Data jualan akan disimpan ke dalam sistem."
    End If

    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then
        
        G_JENIS_URUSAN = 7
        
        '$$$ No. staff $$$ - Start
        If InStr(1, Frm93.CBB3, "  |  ") <> 0 Then
        
            Frm93_LM_EMP_NO = Split(Frm93.CBB3, "  |  ")(1)
            
        Else
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!NoPekerja) Then Frm93_LM_EMP_NO = rs!NoPekerja
    
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
        If Frm93.CB9 = 1 Then
        
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
        
' ### Periksa status barang samada masih dalam stok atau tidak ### - Start
        If Frm93.CB3 = 1 Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from Data_Database where no_siri_produk='" & Frm93.TB6 & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!StatusItem) Then
                
                    If rs!StatusItem <> 10 Then
                        
                        MsgBox "Status terkini bagi item ini telah berubah. Sila periksa status terbaru item ini.", vbExclamation, "Info"
                        
                        rs.Close
                        Set rs = Nothing
                        
                        Exit Sub
                    
                    End If
                
                End If
                
            Else
            
                MsgBox "Status terkini bagi item ini telah berubah. Sila periksa status terbaru item ini.", vbExclamation, "Info"
                
                rs.Close
                Set rs = Nothing
            
                Exit Sub
                
            End If
            
            rs.Close
            Set rs = Nothing
            
        End If
' ### Periksa status barang samada masih dalam stok atau tidak ### - End
    
' ### Periksa kategori pembeli ### - Start
        If Frm93.L36_Text <> vbNullString Then
        
            If Frm28.L5_Text <> vbNullString Then
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
                If Not rs.EOF Then
                
                    If Not IsNull(rs!kategori_pelanggan) Then Frm93_LM_KATEGORI = rs!kategori_pelanggan
                    
                End If
                
                rs.Close
                Set rs = Nothing
                
            End If
            
        End If
' ### Periksa kategori pembeli ### - End
        
        'Frm93_LM_No_RUJUKAN_TEMPAHAN = Frm93.L17_Text 'No. Rujukan Tempahan
        'If Frm93.CB9 = 0 Then Frm93_LM_No_RESIT_TEMPAHAN = Frm93.L18_Text 'No. Invoice Tempahan (Rasmi)
        'If Frm93.CB9 = 1 Then Frm93_LM_No_RESIT_TEMPAHAN = Frm93.L21_Text 'No. Invoice Tempahan (Tidak Rasmi)
        
'---------------------------------------No. Invoice
        LM_NOW = Now
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        rs.Open "select * from 12_rujukan_tempahan", cn2, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        rs!tarikh = Frm93.DTPicker1
        rs!terminal = G_TERMINAL
        rs!write_timestamp = LM_NOW
        rs!Status = 1
        rs!nama_staff = MDI_frm1.L3_Text
        rs.Update
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        rs.Open "select * from 12_rujukan_tempahan where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & Frm93.DTPicker1 & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!ID) Then
                
                Frm93_LM_No_RUJUKAN_TEMPAHAN = rs!ID 'No. Rujukan Belian

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
        
        GoTo skip_aaa:
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        If Frm93.CB9 = 0 Then rs.Open "select * from 4_senarai_invoice_rasmi", cn2, adOpenKeyset, adLockOptimistic
        If Frm93.CB9 = 1 Then rs.Open "select * from 5_senarai_invoice_tidak_rasmi", cn2, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        rs!tarikh = Frm93.DTPicker1
        rs!terminal = G_TERMINAL
        rs!write_timestamp = LM_NOW
        rs!Status = 1
        rs!nama_staff = MDI_frm1.L3_Text
        rs.Update
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        If Frm93.CB9 = 0 Then rs.Open "select * from 4_senarai_invoice_rasmi where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & Frm93.DTPicker1 & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic
        If Frm93.CB9 = 1 Then rs.Open "select * from 5_senarai_invoice_tidak_rasmi where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & Frm93.DTPicker1 & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!ID) Then
                
                Frm93_LM_No_RESIT_TEMPAHAN = rs!ID 'No. Rujukan Belian
                If Frm93.CB9 = 0 Then rs!no_invoice = "TMP" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(rs!ID, "000000")
                If Frm93.CB9 = 1 Then rs!no_invoice = "TMP" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(rs!ID, "000000")
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
skip_aaa:

        GoTo a:
        
'### Periksa NO INVOICE sebelum simpan data ke dalam database ### - Start
Re_gen_no_resit:
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm93.CB9 = 0 Then rs.Open "select * from 22_jualan where no_resit='" & "TMP" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm93_LM_No_RESIT_TEMPAHAN, "000000") & "' AND bil_rasmi = 1 AND status = 1", cn, adOpenKeyset, adLockOptimistic
        If Frm93.CB9 = 1 Then rs.Open "select * from 22_jualan where no_resit='" & "TMP" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm93_LM_No_RESIT_TEMPAHAN, "000000") & "' AND bil_rasmi = 0 AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            Frm93_LM_No_RESIT_TEMPAHAN = Frm93_LM_No_RESIT_TEMPAHAN + 1
            If Frm93.CB9 = 0 Then Frm93.L18_Text = Frm93_LM_No_RESIT_TEMPAHAN
            If Frm93.CB9 = 1 Then Frm93.L21_Text = Frm93_LM_No_RESIT_TEMPAHAN
            
            rs.Close
            Set rs = Nothing
            
            GoTo Re_gen_no_resit:
        End If
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm93.CB9 = 0 Then rs.Open "select * from 40_tempahan_deposit where no_resit_tempahan='" & "TMP" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm93_LM_No_RESIT_TEMPAHAN, "000000") & "' AND bil_rasmi = 1", cn, adOpenKeyset, adLockOptimistic
        If Frm93.CB9 = 1 Then rs.Open "select * from 40_tempahan_deposit where no_resit_tempahan='" & "TMP" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm93_LM_No_RESIT_TEMPAHAN, "000000") & "' AND bil_rasmi = 0", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            Frm93_LM_No_RESIT_TEMPAHAN = Frm93_LM_No_RESIT_TEMPAHAN + 1
            If Frm93.CB9 = 0 Then Frm93.L18_Text = Frm93_LM_No_RESIT_TEMPAHAN
            If Frm93.CB9 = 1 Then Frm93.L21_Text = Frm93_LM_No_RESIT_TEMPAHAN
            
            rs.Close
            Set rs = Nothing
            
            GoTo Re_gen_no_resit:
        End If
        
        rs.Close
        Set rs = Nothing
'### Periksa NO INVOICE sebelum simpan data ke dalam database ### - End

a:

Re_Gen_No_Rujukan:
'###Masukkan Jualan Ke Dalam Table Akaun Jualan### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm93.CB9 = 0 Then rs.Open "select * from 22_jualan where no_resit='" & "TMP" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm93_LM_No_RESIT_TEMPAHAN, "000000") & "' AND bil_rasmi = 1 AND status = 1", cn, adOpenKeyset, adLockOptimistic
        If Frm93.CB9 = 1 Then rs.Open "select * from 22_jualan where no_resit='" & "TMP" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm93_LM_No_RESIT_TEMPAHAN, "000000") & "' AND bil_rasmi = 0 AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
            rs.AddNew
            
            rs!no_resit = "TMP" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm93_LM_No_RUJUKAN_TEMPAHAN, "000000")
            LM_NO_INVOICE = "TMP" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm93_LM_No_RUJUKAN_TEMPAHAN, "000000")
            rs!bil_rasmi = 1 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
            
            'If Frm93.CB9 = 0 Then
            
            '    If Frm93.L18_Text <> vbNullString Then
            '        rs!no_resit = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm93_LM_No_RESIT_TEMPAHAN, "000000")  'No. invoice rasmi
            '        LM_NO_INVOICE = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm93_LM_No_RESIT_TEMPAHAN, "000000")  'No. invoice rasmi
            '    Else
            '        rs!no_resit = Null 'No. invoice rasmi
            '    End If
            '    rs!bil_rasmi = 1 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
                
            'Else
            
            '    If Frm93.L21_Text <> vbNullString Then
            '        rs!no_resit = "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm93_LM_No_RESIT_TEMPAHAN, "000000")  'No. invoice tidak rasmi
            '        LM_NO_INVOICE = "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm93_LM_No_RESIT_TEMPAHAN, "000000")  'No. invoice tidak rasmi
            '    Else
            '        rs!no_resit = Null 'No. invoice tidak rasmi
            '    End If
            '    rs!bil_rasmi = 0 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
            
            'End If
            rs!tarikh = Frm93.DTPicker1 'Tarikh Jualan
            
            rs!tunai = Format(0, "0.00")
            rs!bank_in = Format(0, "0.00") 'Cara Bayaran : Bank In
            rs!kad_kredit = Format(0, "0.00") 'Cara Bayaran : Kad Kredit
            rs!duit_simpanan_kedai = Format(0, "0.00") 'Cara Bayaran : Simpanan Duit Di Kedai
            rs!cek = Format(0, "0.00") 'Cara Bayaran : Cek

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
                    If Frm93.L19_Text <> vbNullString Then
                        rs!kadar_gst_kad_kredit = Format(Frm93.L19_Text, "0.00") 'Cara Bayaran : Kadar GST bagi kad kredit
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
                    Frm93_LM_Flag_SIMPANAN = 1 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
                End If
                rs!duit_simpanan_kedai = Format(frm130.TB21, "0.00") 'Cara Bayaran : Simpanan Duit Di Kedai
            Else
                rs!duit_simpanan_kedai = "0.00" 'Cara Bayaran : Simpanan Duit Di Kedai
            End If
            If Frm93.TB32 <> vbNullString Then
                rs!jumlah_bayaran = Format(Frm93.TB32, "0.00") 'Cara Bayaran : Jumlah Bayaran
            Else
                rs!jumlah_bayaran = Null 'Cara Bayaran : Jumlah Bayaran
            End If
            If Frm93.TB23 <> vbNullString Then
                rs!harga_barang = Format(Frm93.TB23, "0.00") 'Jumlah Harga Barang Tanpa GST (RM)
                rs!harga_barang_dengan_gst = Format(Frm93.TB23, "0.00") 'Jumlah Harga Barang Dengan GST (RM)
                rs!harga_lepas_diskaun = Format(Frm93.TB23, "0.00") 'Harga Selepas Diskaun (RM)
                rs!harga_jualan = Format(Frm93.TB23, "0.00") 'Jumlah Harga Jualan (RM)
                rs!gst_zr_harga = Format(Frm93.TB23, "0.00") 'Harga Keseluruhan Bagi Barang ZR
            Else
                rs!harga_barang = Null 'Jumlah Harga Barang Tanpa GST (RM)
            End If
            rs!jumlah_cukai_gst = Null 'Jumlah Cukai GST (ZR + SR)
            rs!diskaun = Format(0, "0.00") 'Jumlah Diskaun (%)
            rs!adjustment = Format(0, "0.00") 'Adjustment (RM)
            rs!loss_trade_in = Format(0, "0.00") 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (%)
            rs!loss_trade_in_rm = Format(0, "0.00") 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (RM)
            rs!flag_bayaran = 0 '0 : Pembeli Bayar , 1 : Kedai Bayar
            If Frm93.TB20 <> vbNullString Then
                rs!jumlah_perlu_bayar = Format(Frm93.TB20, "0.00")  'Jumlah Bayaran Yang Perlu Dibuat (RM)
            Else
                rs!jumlah_perlu_bayar = Null 'Jumlah Bayaran Yang Perlu Dibuat (RM)
            End If
            rs!kuantiti_barang = 1 'Kuantiti Barang Yang Dijual
            If Frm93.CB19 = 1 Then 'Barang kemas
                If Frm93.TB1 <> vbNullString Then
                    rs!JUMLAH_BERAT = Format(Frm93.TB1, "0.00") 'Jumlah Berat Barang Yang Dijual
                Else
                    rs!JUMLAH_BERAT = Null 'Jumlah Berat Barang Yang Dijual
                End If
            ElseIf Frm93.CB20 = 1 Then 'Barang kemas
                If Frm93.TB8 <> vbNullString Then
                    rs!JUMLAH_BERAT = Format(Frm93.TB8, "0.00") 'Jumlah Berat Barang Yang Dijual
                Else
                    rs!JUMLAH_BERAT = Null 'Jumlah Berat Barang Yang Dijual
                End If
            Else
                rs!JUMLAH_BERAT = Null 'Jumlah Berat Barang Yang Dijual
            End If
            rs!gst_zr_cukai = Format(0, "0.00") 'Jumlah Cukai Bagi ZR
            rs!gst_sr_harga = Format(0, "0.00") 'Harga Keseluruhan Bagi Barang SR
            rs!gst_sr_cukai = Format(0, "0.00") 'Jumlah Cukai Bagi SR
            rs!caj_pos = "0.00"
            rs!no_tracking = Null
            rs!no_pekerja = Frm93_LM_EMP_NO 'No. Pekerja
            If Frm93.L36_Text <> vbNullString Then
                If Frm28.L5_Text <> vbNullString Then
                    rs!no_rujukan_pembeli = Frm28.L5_Text 'No. Rujukan Pembeli
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

                 
'1:  Pelanggan
'2:  Member
'3:  RAF
'4:  Pengedar
'5:  Normal Dealer
'6:  Master Dealer

'Zakaria&Sons
'1 : Pembeli biasa
'2 : Ahli biasa
'3 : Silver
'4 : Gold
'5 : Platinum
            rs!kategori_pembeli = Frm93_LM_KATEGORI
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
            rs!write_timestamp = LM_NOW
            rs!Menu = 2
            rs!cawangan = G_CAWANGAN
            
            DATA_SAVE = 1
            rs.Update
        Else
            
            Frm93_LM_No_RUJUKAN_TEMPAHAN = Frm93_LM_No_RUJUKAN_TEMPAHAN + 1
            'Frm93_LM_No_RESIT_TEMPAHAN = Frm93_LM_No_RESIT_TEMPAHAN + 1
            'If Frm93.CB9 = 0 Then Frm93.L18_Text = Frm93_LM_No_RESIT_TEMPAHAN 'No. invoice rasmi
            'If Frm93.CB9 = 1 Then Frm93.L21_Text = Frm93_LM_No_RESIT_TEMPAHAN 'No. invoice tidak rasmi
            
            rs.Close
            Set rs = Nothing
            GoTo Re_Gen_No_Rujukan:
        End If
        
        rs.Close
        Set rs = Nothing
'###Masukkan Jualan Ke Dalam Table Akaun Jualan### - End
        
Re_Gen_No_Rujukan2:
'###Masukkan Data Tempahan### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 40_tempahan_deposit where no_rujukan_tempahan='" & G_KOD_KEDAI & "-" & Frm93_LM_No_RUJUKAN_TEMPAHAN & "' AND status_invoice = 1", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
            rs.AddNew
            'If Frm93.L17_Text <> vbNullString Then 'No. Rujukan Tempahan
                rs!no_rujukan_tempahan = G_KOD_KEDAI & "-" & Frm93_LM_No_RUJUKAN_TEMPAHAN
            'Else
            '    rs!no_rujukan_tempahan = Null
            'End If
            rs!no_resit_tempahan = LM_NO_INVOICE
            rs!Status = "Belum Siap"
            
            If Frm93.CB2 = 1 Then
                rs!jenis_tempahan = 0 'Jenis Tempahan , 0 : Tempahan Barang Baru , 1 : Tempahan Barang Kedai
                
                If Frm93.CBB1 <> vbNullString Then 'Kategori Produk
                    rs!kategori_Produk = Frm93.CBB1
                Else
                    rs!kategori_Produk = Null
                End If
                If Frm93.CBB2 <> vbNullString Then 'Purity
                    rs!purity = Frm93.CBB2
                Else
                    rs!purity = Null
                End If
                If Frm93.TB3 <> vbNullString Then 'Upah
                    rs!UPAH = Format(Frm93.TB3, "0.00")
                Else
                    rs!UPAH = Null
                End If
                If Frm93.TB4 <> vbNullString Then 'Anggaran Harga
                    rs!anggaran_harga = Format(Frm93.TB4, "0.00")
                Else
                    rs!anggaran_harga = Null
                End If
                    
                If Frm93.CB19 = 1 Then
                    rs!type_barang_kemas = 0 'Jenis Barang , 0 : Barang Kemas , 1 : Barang Permata
                    If Frm93.TB1 <> vbNullString Then 'Anggaran Berat
                        rs!anggaran_berat = Format(Frm93.TB1, "0.00")
                    Else
                        rs!anggaran_berat = Null
                    End If
                    If Frm93.TB2 <> vbNullString Then 'Harga Semasa
                        rs!harga_Semasa = Format(Frm93.TB2, "0.00")
                    Else
                        rs!harga_Semasa = Null
                    End If
                ElseIf Frm93.CB20 = 1 Then
                    rs!type_barang_kemas = 1 'Jenis Barang , 0 : Barang Kemas , 1 : Barang Permata
                End If
                'rs!status_tukang = "Belum Hantar" ' Belum Hantar , Belum Siap , Siap
                
                rs!Berat_Asal = Null
                rs!berat_jualan = Null
                rs!adjustment = Null
                rs!harga_asal = Null
            End If
            
            If Frm93.CB3 = 1 Then
                rs!jenis_tempahan = 1 'Jenis Tempahan , 0 : Tempahan Barang Baru , 1 : Tempahan Barang Kedai
                
                If Frm93.TB6 <> vbNullString Then 'No. Siri Produk
                    rs!no_siri_Produk = Frm93.TB6
                Else
                    rs!no_siri_Produk = Null
                End If
                If Frm93.L4_Text <> vbNullString Then 'Kategori Produk
                    rs!kategori_Produk = Frm93.L4_Text
                Else
                    rs!kategori_Produk = Null
                End If
                
                If Frm93.L13_Text = 0 Then
                    rs!type_barang_kemas = 0 'Jenis Barang , 0 : Barang Kemas , 1 : Barang Permata
                    If Frm93.TB7 <> vbNullString Then 'Berat Asal
                        rs!Berat_Asal = Format(Frm93.TB7, "0.00")
                    Else
                        rs!Berat_Asal = Null
                    End If
                    If Frm93.TB8 <> vbNullString Then 'Berat Jualan
                        rs!berat_jualan = Format(Frm93.TB8, "0.00")
                    Else
                        rs!berat_jualan = Null
                    End If
                    If Frm93.TB9 <> vbNullString Then 'Harga Semasa
                        rs!harga_Semasa = Format(Frm93.TB9, "0.00")
                    Else
                        rs!harga_Semasa = Null
                    End If
                ElseIf Frm93.L13_Text = 1 Then
                    rs!type_barang_kemas = 1 'Jenis Barang , 0 : Barang Kemas , 1 : Barang Permata
                    rs!Berat_Asal = Null
                    rs!berat_jualan = Null
                    rs!harga_Semasa = Null
                End If
                
                If Frm93.TB10 <> vbNullString Then 'Upah
                    rs!UPAH = Format(Frm93.TB10, "0.00")
                Else
                    rs!UPAH = Null
                End If
                If Frm93.TB11 <> vbNullString Then 'Harga Asal
                    rs!harga_asal = Format(Frm93.TB11, "0.00")
                Else
                    rs!harga_asal = Null
                End If
                If Frm93.TB12 <> vbNullString Then 'Adjustment
                    rs!adjustment = Format(Frm93.TB12, "0.00")
                Else
                    rs!adjustment = Null
                End If
                If Frm93.TB13 <> vbNullString Then 'Anggaran Harga
                    rs!anggaran_harga = Format(Frm93.TB13, "0.00")
                Else
                    rs!anggaran_harga = Null
                End If
                
                rs!anggaran_berat = Null
            End If
            
'Kategori Pembeli
'=================
'1:  Pelanggan
'2 : Member / Ahli
'3:  RAF
'4:  Pengedar
'5:  Normal Dealer (ND)
'6:  Master Dealer (MD)

            rs!kategori_pembeli = Frm93_LM_KATEGORI
            
            If Frm93.L36_Text <> vbNullString Then
                If Frm28.L5_Text <> vbNullString Then
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
            End If
            If Frm93.L35_Text <> vbNullString Then
                If Frm26.TB1 <> vbNullString Then
                    rs!no_rujukan_pelanggan = Null 'No. Rujukan Pembeli
                    rs!no_ic = Null 'Maklumat Pembeli : No. Kad Pengenalan
                    
                    If Frm26.TB1 <> vbNullString Then
                        rs!Nama = UCase(Frm26.TB1) 'Maklumat Pembeli : Nama
                    Else
                        rs!Nama = Null 'Maklumat Pembeli : Nama
                    End If
                    rs!no_ic = Null 'Maklumat Pembeli : No. Kad Pengenalan
                    If Frm26.TB2 <> vbNullString Then
                        rs!no_tel = UCase(Frm26.TB2) 'No. Telefon
                    Else
                        rs!no_tel = Null 'No. Telefon
                    End If
                End If
            End If
            If Frm93.CB6 = 1 Then 'Flag Trade In
                Frm93_LM_Flag_TRADE_IN = 1 '0 : Tiada Urusan Trade In , 1 : Ada Urusan Trade In
                
                rs!flag_trade_in = 1 '0 : Tiada Bayaran Deposit Menggunakan Trade In , 1 : Ada Bayaran Deposit Menggunakan Trade In
                If Frm93.L15_Text <> vbNullString Then 'No. Resit Trade In
                    rs!no_resit_trade_in = Frm93.L15_Text
                Else
                    rs!no_resit_trade_in = Null
                End If
                If Frm93.TB17 <> vbNullString Then 'Jumlah Nilaian Trade In
                    rs!nilaian_trade_in = Format(Frm93.TB17, "0.00")
                Else
                    rs!nilaian_trade_in = Null
                End If
            Else
                Frm93_LM_Flag_TRADE_IN = 0 '0 : Tiada Urusan Trade In , 1 : Ada Urusan Trade In
                
                rs!flag_trade_in = 0 '0 : Tiada Bayaran Deposit Menggunakan Trade In , 1 : Ada Bayaran Deposit Menggunakan Trade In
                rs!no_resit_trade_in = Null 'No. Resit Trade In
                rs!nilaian_trade_in = Null 'Jumlah Nilaian Trade In
            End If
            
            If Frm93.TB20 <> vbNullString Then 'Jumlah Deposit Yang Dibayar Secara Tunai
                rs!jumlah_deposit_tunai = Format(Frm93.TB20, "0.00")
            Else
                rs!jumlah_deposit_tunai = Null
            End If
            If Frm93.TB22 <> vbNullString Then 'Jumlah Deposit Dari Barangan Trade In
                rs!jumlah_deposit_trade_in = Format(Frm93.TB22, "0.00")
            Else
                rs!jumlah_deposit_trade_in = Null
            End If
            If Frm93.TB23 <> vbNullString Then 'Jumlah Deposit (Tanpa GST)
                rs!jumlah_tanpa_gst = Format(Frm93.TB23, "0.00")
                rs!jumlah_dengan_gst = Format(Frm93.TB23, "0.00")
            Else
                rs!jumlah_tanpa_gst = Null
            End If
            If Frm93.TB20 <> vbNullString Then
                rs!jumlah_perlu_bayar = Format(Frm93.TB20, "0.00") 'Jumlah Bayaran Sebelum Adjustment
                rs!jumlah_bayaran = Format(Frm93.TB20, "0.00") 'Jumlah Bayaran Deposit Selepas Adjustment
            Else
                rs!jumlah_perlu_bayar = Null
            End If
            rs!adjustment_bayaran = Format(0, "0.00") 'Adjustment Bagi Bayaran Keseluruhan
            If Frm93.TB33 <> vbNullString Then 'Remarks
                rs!remarks = Frm93.TB33
            Else
                rs!remarks = Null
            End If
            rs!tarikh = Frm93.DTPicker1 'Tarikh Tempahan
            rs!no_pekerja = Frm93_LM_EMP_NO 'No. Pekerja

            rs!terminal = G_TERMINAL
            rs!write_timestamp = LM_NOW
            rs!status_invoice = 1 '0 : Tidak aktif (dibatalkan) , 1:  Aktif
            If Frm93.CB9 = 0 Then
                rs!bil_rasmi = 1 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
            Else
                rs!bil_rasmi = 0 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
            End If
            rs!no_staff = G_LOGIN_USER 'No. staff yang memasukkan/edit data (Login)
            rs!cawangan = G_CAWANGAN
            
            rs.Update
        Else
            Frm93_LM_No_RUJUKAN_TEMPAHAN = Frm93_LM_No_RUJUKAN_TEMPAHAN + 1
            'Frm93.L17_Text = Frm93_LM_No_RUJUKAN_TEMPAHAN 'No. Rujukan Ansuran
            
            rs.Close
            Set rs = Nothing
            GoTo Re_Gen_No_Rujukan2:
        End If
        
        rs.Close
        Set rs = Nothing
'###Masukkan Data Tempahan### - End

'### Update Maklumat Trade In ### - Start
        If Frm93_LM_Flag_TRADE_IN = 1 Then '0 : Tiada Urusan Trade In , 1 : Ada Urusan Trade In

            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & Frm93.L15_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
            
                G_ID = rs!ID
                Call recovery_16_gold_bar_belian
                
                rs!trade_in_status = 1
                rs!no_staff = Frm93_LM_EMP_NO 'No. Pekerja
                rs!terminal = G_TERMINAL
                rs!write_timestamp2 = LM_NOW
                rs!jenis_urusan = G_JENIS_URUSAN
                rs!remarks = "Deposit bagi tempahan emas"
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
            
        End If
'### Update Maklumat Trade In ### - End

'### Update Table Database Bagi Item Ini ### - Start
        If Frm93.CB3 = 1 Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from Data_Database where no_siri_produk='" & Frm93.TB6 & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
            
                G_ID = rs!ID
                Call recovery_data_database
                
                rs!StatusItem = 14
                rs!write_timestamp2 = LM_NOW
                rs!no_pekerja = Frm93_LM_EMP_NO
                rs!terminal = G_TERMINAL
                rs!Menu = 2
                
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
        End If
'### Update Table Database Bagi Item Ini ### - End

'###Update Data Simpanan Duit Pelanggan### - Start
        If Frm93_LM_Flag_SIMPANAN = 1 Then '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai

            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
            
                G_ID = rs!ID
                Call recovery_senarai_pelanggan
                
                Frm93_LM_JUMLAH_SIMPANAN = frm130.L26_Text  'Jumlah Simpanan Yang Ada
                Frm93_LM_GUNA_SIMPAN = frm130.TB21  'Jumlah Simpanan Yang Hendak Digunakan
                
                rs!baki_simpanan = Format(Frm93_LM_JUMLAH_SIMPANAN - Frm93_LM_GUNA_SIMPAN, "0.00") 'Baki Simpanan
                
                rs!write_timestamp2 = LM_NOW
                rs!no_staff = Frm93_LM_EMP_NO 'No. Pekerja
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
            rs!tarikh = Frm93.DTPicker1 'Tarikh
            rs!jenis = 1 '0 : Simpanan , 1 : Penggunaan Duit
            rs!no_rujukan_pelanggan = Frm28.L5_Text 'No. Rujukan Pelanggan
            rs!no_resit = LM_NO_INVOICE 'No. Resit Tempahan
            rs!jumlah = Format(frm130.TB21, "0.00") 'Jumlah Simpanan Yang Digunakan (RM)
            rs!jenis_penggunaan = 2 '0 : Belian Barangan Kemas , 1 : Ansuran , 2 : Tempahan (Deposit) , 3 : Servis , 4 : Tempahan (Ambilan Barang)
            rs!no_rujukan_pekerja = Frm93_LM_EMP_NO 'No. Pekerja
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

'###Masukkan Data Penjual Yang Tidak Berdaftar Ke Dalam Database### - Start
        If Frm93.L35_Text <> vbNullString And Frm93.L36_Text = vbNullString Then

            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 44_senarai_pelanggan", cn, adOpenKeyset, adLockOptimistic
            
            rs.AddNew
            rs!tarikh = Frm93.DTPicker1 'Tarikh
            rs!no_resit = LM_NO_INVOICE 'No. Resit Tempahan
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
            rs!no_staff = Frm93_LM_EMP_NO 'No. Pekerja
            rs!terminal = G_TERMINAL
            rs!jenis_urusan = G_JENIS_URUSAN
            rs!cawangan = G_CAWANGAN
            rs.Update
            
            rs.Close
            Set rs = Nothing
            
        End If
'###Masukkan Data Penjual Yang Tidak Berdaftar Ke Dalam Database### - End
        
        WRITE_NO = 0
        
        'Set rs = New ADODB.Recordset
        'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        'rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

        'If Not rs.EOF Then
        '    If rs!Default1 = "Default" Then
                'If Frm93.CB9 = 0 Then
                '    rs!ResitNo = Frm93.L18_Text + 1
                '    WRITE_NO = 1
                'Else
                '    rs!no_rujukan_tak_rasmi = Frm93.L21_Text + 1
                '    WRITE_NO = 1
                'End If
        '        If IsNumeric(Frm93.L17_Text) Then
        '            rs!no_rujukan_book = Frm93.L17_Text + 1
        '            WRITE_NO = 1
        '        End If
        '        If WRITE_NO = 1 Then
        '            rs.Update
        '        End If
        '    End If
        'End If
        
        'rs.Close
        'Set rs = Nothing
    
        user = MDI_frm1.L3_Text
        LogAct_Memory = "[" & G_LOGIN_USER & "] Deposit tempahan [" & LM_NO_INVOICE & "]."
        LogDate_Memory = LM_NOW
        Call UpdateLog_Database
        
        Call Frm93_initial_setting
        
        Frm93.Frame2.Visible = False
        
        Note = "Data telah berjaya disimpan." & vbCrLf & _
                "Adakah anda ingin cetak invoice tempahan ini?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
            G_PREVIEW = 1
            G_KEDAI = G_CAWANGAN
            G_No_INV_BOOK = vbNullString
            G_No_INV_BOOK = LM_NO_INVOICE 'No. Invoice
            Call Frm94_invoice_deposit_tempahan
        End If
    End If
End If
End Sub

Private Sub CMD13_Click()
'On Error Resume Next
If Frm93.CBB4 = "Tiada filter tarikh" Then
    Frm93.L27_Text = 0 '0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh
Else
    Frm93.L27_Text = 1 '0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh
    Frm93.L28_Text = Frm93.DTPicker2 'Tarikh mula
    Frm93.L29_Text = Frm93.DTPicker3 'Tarikh akhir
End If

If Frm93.TB30.Visible = True Then

    If Frm93.TB30 = vbNullString Then
    
        MsgBox "Carian tidak boleh dikosongkan.", vbExclamation, "Info"
        
        Frm93.TB30.SetFocus
        Exit Sub
        
    End If
    
    If InStr(1, Frm93.TB30, "*") <> 0 Or InStr(1, Frm93.TB30, "/") <> 0 Or InStr(1, Frm93.TB30, "\") <> 0 Or InStr(1, Frm93.TB30, "'") <> 0 Then
        MsgBox "Carian mengandungi simbol yang tidak sah.", vbExclamation, "Info"
        
        Frm93.TB30.SetFocus
        Exit Sub
    End If
   
End If

Frm93.L30_Text = Frm93.CBB5 'Jenis tempahan
Frm93.L31_Text = Frm93.CBB7 'Status
Frm93.L32_Text = Frm93.CBB8 'Lain-lain
Frm93.L38_Text = UCase(Frm93.TB30) 'Lain-lain
Frm93.L45_Text = Frm93.CBB9 'Cawangan

Frm93.L62_Text = -1 'Start Point
Frm93.L60_Text = 0 'Current Page
Frm93.L63_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
GM_NEXT_PREV = 0

Call frm93_tempahan_header
Call frm93_tempahan
End Sub

Private Sub CMD14_Click()
'On Error Resume Next
Dim Err(40)
Dim Frm93_LM_ERR_BERAT_ASAL As Double
Dim Frm93_LM_ERR_BERAT_JUALAN As Double
Dim Frm93_LM_ERR_JUMLAH_BAYARAN As Double
Dim Frm93_LM_ERR_HARGA As Double
Dim Frm93_LM_JUMLAH_SIMPANAN As Double
Dim Frm93_LM_GUNA_SIMPAN As Double

Frm93_LM_ERR_BERAT_ASAL = 0
Frm93_LM_ERR_BERAT_JUALAN = 0
Frm93_LM_ERR_JUMLAH_BAYARAN = 0 'Jumlah Bayaran
Frm93_LM_ERR_HARGA = 0 'Jumlah Perlu Bayar
Frm93_LM_Flag_TRADE_IN = 0 '0 : Tiada Urusan Trade In , 1 : Ada Urusan Trade In
Frm93_LM_Flag_SIMPANAN = 0 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
Frm93_LM_JUMLAH_SIMPANAN = 0  'Jumlah Simpanan Yang Ada
Frm93_LM_GUNA_SIMPAN = 0 'Jumlah Simpanan Yang Hendak Digunakan
Frm93_LM_KATEGORI = 0

If Frm93.CB2 = 0 And Frm93.CB3 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan jenis tempahan."
End If
If Frm93.CB2 = 1 Then 'Jenis Tempahan : Tempahan Barang Baru
    If Frm93.CB19 = 0 And Frm93.CB20 = 0 Then
        x = x + 1
        Err(x) = "Sila buat pilihan jenis barang samada barang kemas atau barang permata."
    End If
    If Frm93.CBB1 = vbNullString Then
        x = x + 1
        Err(x) = "Sila pilih kategori produk."
    End If
    If Frm93.CBB2 = vbNullString Then
        x = x + 1
        Err(x) = "Sila pilih purity."
    End If
    
    If Frm93.CB19 = 1 Then 'Jenis Barang : Barang Kemas
        If Frm93.TB1 = vbNullString Or (Frm93.TB1 <> vbNullString And Not IsNumeric(Frm93.TB1)) Then
            x = x + 1
            Err(x) = "Sila masukkan [Anggaran Berat (g)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
        End If
        If Frm93.TB2 = vbNullString Or (Frm93.TB2 <> vbNullString And Not IsNumeric(Frm93.TB2)) Then
            x = x + 1
            Err(x) = "Sila masukkan [Harga Semasa (RM/g)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
        End If
        If Frm93.TB3 = vbNullString Or (Frm93.TB3 <> vbNullString And Not IsNumeric(Frm93.TB3)) Then
            x = x + 1
            Err(x) = "Sila masukkan [Upah (RM)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
        End If
    End If
    If Frm93.CB20 = 1 Then 'Jenis Barang : Barang Permata
        If Frm93.TB3 = vbNullString Or (Frm93.TB3 <> vbNullString And Not IsNumeric(Frm93.TB3)) Then
            x = x + 1
            Err(x) = "Sila masukkan [Upah (RM)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
        End If
        If Frm93.TB4 = vbNullString Or (Frm93.TB4 <> vbNullString And Not IsNumeric(Frm93.TB4)) Then
            x = x + 1
            Err(x) = "Sila masukkan [Anggaran Harga (RM)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
        End If
    End If
End If
If Frm93.CB3 = 1 Then 'Jenis Tempahan : Tempahan Barang Kedai
    If Frm93.L4_Text = vbNullString Then
        x = x + 1
        Err(x) = "Tiada maklumat kategori produk"
    End If
    If Frm93.L13_Text = 0 Then 'Jenis Barang : Barang Kemas
    
        If (Frm93.TB7 <> vbNullString And IsNumeric(Frm93.TB7)) And (Frm93.TB8 <> vbNullString And IsNumeric(Frm93.TB8)) Then
            Frm93_LM_ERR_BERAT_ASAL = Frm93.TB7 'Berat Asal
            Frm93_LM_ERR_BERAT_JUALAN = Frm93.TB8 'Berat Jualan
            
            If Frm93_LM_ERR_BERAT_JUALAN > Frm93_LM_ERR_BERAT_ASAL Then
                x = x + 1
                Err(x) = "Berat jualan melebihi berat asal"
            End If
        End If
        
        If Frm93.TB8 = vbNullString Or (Frm93.TB8 <> vbNullString And Not IsNumeric(Frm93.TB8)) Then
            x = x + 1
            Err(x) = "Sila masukkan [Berat Jualan (g)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
        End If
        If Frm93.TB9 = vbNullString Or (Frm93.TB9 <> vbNullString And Not IsNumeric(Frm93.TB9)) Then
            x = x + 1
            Err(x) = "Sila masukkan [Harga Semasa (RM/g)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
        End If
    ElseIf Frm93.L13_Text = 1 Then 'Jenis Barang : Barang Permata
        If Frm93.TB11 = vbNullString Or (Frm93.TB11 <> vbNullString And Not IsNumeric(Frm93.TB11)) Then
            x = x + 1
            Err(x) = "Sila masukkan [Harga Asal (RM)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
        End If
    End If
    If Frm93.TB10 = vbNullString Or (Frm93.TB10 <> vbNullString And Not IsNumeric(Frm93.TB10)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Upah (RM)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm93.TB12 = vbNullString Or (Frm93.TB12 <> vbNullString And Not IsNumeric(Frm93.TB12)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Adjustment (RM)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If
If Frm93.L35_Text = vbNullString And Frm93.L36_Text = vbNullString Then
    x = x + 1
    Err(x) = "Tiada maklumat pembeli."
End If
If Frm93.TB20 = vbNullString Or (Frm93.TB20 <> vbNullString And Not IsNumeric(Frm93.TB20)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Deposit (RM)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm93.CBB3 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih nama pekerja."
End If
If Frm93.CB6 = 1 Then 'Bayaran Deposit Dari Barang Trade In
    If Frm93.L15_Text = vbNullString Then
        x = x + 1
        Err(x) = "Tiada maklumat No. Voucher Trade In."
    End If
    If Frm93.TB17 = vbNullString Then
        x = x + 1
        Err(x) = "Tiada maklumat nilaian Voucher Trade In."
    End If
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
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara Duit Simpanan Di Kedai. Sila masukkan 0 jika tiada bayaran secara simpanan di kedai."
End If
If (frm130.TB21 <> vbNullString And IsNumeric(frm130.TB21)) And (frm130.L26_Text <> vbNullString And IsNumeric(frm130.L26_Text)) Then
    Frm93_LM_JUMLAH_SIMPANAN = frm130.L26_Text  'Jumlah Simpanan Yang Ada
    Frm93_LM_GUNA_SIMPAN = frm130.TB21  'Jumlah Simpanan Yang Hendak Digunakan
    
    If Frm93_LM_GUNA_SIMPAN > Frm93_LM_JUMLAH_SIMPANAN Then
        x = x + 1
        Err(x) = "Jumlah simpanan yang hendak digunakan melebihi simpanan yang ada."
    End If
End If
If (frm130.TB32 <> vbNullString And IsNumeric(frm130.TB32)) And (Frm93.TB20 <> vbNullString And IsNumeric(Frm93.TB20)) Then

    Frm93_LM_ERR_JUMLAH_BAYARAN = frm130.TB32 'Jumlah Bayaran
    Frm93_LM_ERR_HARGA = Frm93.TB20 'Jumlah Perlu Bayar
    
    If Frm93_LM_ERR_JUMLAH_BAYARAN <> Frm93_LM_ERR_HARGA Then
        x = x + 1
        Err(x) = "Jumlah bayaran tidak sama dengan jumlah perlu bayar."
    End If
End If

'### Periksa Samada Maklumat Pembeli Diisi Dalam Kedua-dua Ruangan Berdaftar Dan Tidak Berdaftar ### - Start
If x = 0 Then

    If Frm93.L35_Text <> vbNullString And Frm93.L36_Text <> vbNullString Then
    
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
    If Frm93.L35_Text <> vbNullString And Frm93.L36_Text = vbNullString Then
        Note = "Adakah anda yakin untuk teruskan urusan jualan ini ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Data jualan akan disimpan ke dalam sistem." & vbCrLf & _
                vbNullString & vbCrLf & _
                "*** Oleh ini pembeli ini tidak berdaftar dengan sistem, oleh itu data dan rekod pembelian pembeli ini tidak akan disimpan di dalam sistem ***"
    End If
    If Frm93.L35_Text = vbNullString And Frm93.L36_Text <> vbNullString Then
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
        If Frm93.L36_Text <> vbNullString Then
            If Frm28.L5_Text <> vbNullString Then
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
                If Not rs.EOF Then
                
                    If Not IsNull(rs!kategori_pelanggan) Then Frm93_LM_KATEGORI = rs!kategori_pelanggan
                    
                End If
                
                rs.Close
                Set rs = Nothing
                
            End If
        End If
' ### Periksa kategori pembeli ### - End

        G_JENIS_URUSAN = 10

        '$$$ No. staff $$$ - Start
        If InStr(1, Frm93.CBB3, "  |  ") <> 0 Then
        
            Frm93_LM_EMP_NO = Split(Frm93.CBB3, "  |  ")(1)
            
        Else
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!NoPekerja) Then Frm93_LM_EMP_NO = rs!NoPekerja
    
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
        
        LM_NOW = Now

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 22_jualan where no_resit='" & Frm93.L18_Text & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!status_r) Then LM_STATUS_R_ASAL = rs!status_r
        End If
        
        rs.Close
        Set rs = Nothing
        
        Call Frm93_padam_data_deposit
        
        Frm93_LM_No_RUJUKAN_TEMPAHAN = Frm93.L17_Text 'No. Rujukan Tempahan
        Frm93_LM_No_RESIT_TEMPAHAN = Frm93.L18_Text 'No. Resit Tempahan
        
'###Masukkan Jualan Ke Dalam Table Akaun Jualan### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm93.CB9 = 0 Then rs.Open "select * from 22_jualan where no_resit='" & Frm93_LM_No_RESIT_TEMPAHAN & "' AND bil_rasmi = 1 AND status = 1", cn, adOpenKeyset, adLockOptimistic
        If Frm93.CB9 = 1 Then rs.Open "select * from 22_jualan where no_resit='" & Frm93_LM_No_RESIT_TEMPAHAN & "' AND bil_rasmi = 0 AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
            rs.AddNew
            If Frm93.CB9 = 0 Then
            
                rs!bil_rasmi = 1 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
                
            Else
        
                rs!bil_rasmi = 0 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
            
            End If
            rs!no_resit = Frm93_LM_No_RESIT_TEMPAHAN
            LM_NO_INVOICE = Frm93_LM_No_RESIT_TEMPAHAN
            
            rs!tarikh = Frm93.DTPicker1 'Tarikh Jualan

            rs!tunai = Format(0, "0.00")
            rs!bank_in = Format(0, "0.00") 'Cara Bayaran : Bank In
            rs!kad_kredit = Format(0, "0.00") 'Cara Bayaran : Kad Kredit
            rs!duit_simpanan_kedai = Format(0, "0.00") 'Cara Bayaran : Simpanan Duit Di Kedai
            rs!cek = Format(0, "0.00") 'Cara Bayaran : Cek
            
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
                    If Frm93.L19_Text <> vbNullString Then
                        rs!kadar_gst_kad_kredit = Format(Frm93.L19_Text, "0.00") 'Cara Bayaran : Kadar GST bagi kad kredit
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
                    Frm93_LM_Flag_SIMPANAN = 1 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
                End If
                rs!duit_simpanan_kedai = Format(frm130.TB21, "0.00") 'Cara Bayaran : Simpanan Duit Di Kedai
            Else
                rs!duit_simpanan_kedai = "0.00" 'Cara Bayaran : Simpanan Duit Di Kedai
            End If
            If Frm93.TB32 <> vbNullString Then
                rs!jumlah_bayaran = Format(Frm93.TB32, "0.00") 'Cara Bayaran : Jumlah Bayaran
            Else
                rs!jumlah_bayaran = Null 'Cara Bayaran : Jumlah Bayaran
            End If
            If Frm93.TB23 <> vbNullString Then
                rs!harga_barang = Format(Frm93.TB23, "0.00") 'Jumlah Harga Barang Tanpa GST (RM)
                rs!harga_barang_dengan_gst = Format(Frm93.TB23, "0.00") 'Jumlah Harga Barang Dengan GST (RM)
                rs!harga_lepas_diskaun = Format(Frm93.TB23, "0.00") 'Harga Selepas Diskaun (RM)
                rs!harga_jualan = Format(Frm93.TB23, "0.00") 'Jumlah Harga Jualan (RM)
                rs!gst_zr_harga = Format(Frm93.TB23, "0.00") 'Harga Keseluruhan Bagi Barang ZR
            Else
                rs!harga_barang = Null 'Jumlah Harga Barang Tanpa GST (RM)
            End If
            rs!jumlah_cukai_gst = Null 'Jumlah Cukai GST (ZR + SR)
            rs!diskaun = Format(0, "0.00") 'Jumlah Diskaun (%)
            rs!adjustment = Format(0, "0.00") 'Adjustment (RM)
            rs!loss_trade_in = Format(0, "0.00") 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (%)
            rs!loss_trade_in_rm = Format(0, "0.00") 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (RM)
            rs!flag_bayaran = 0 '0 : Pembeli Bayar , 1 : Kedai Bayar
            If Frm93.TB20 <> vbNullString Then
                rs!jumlah_perlu_bayar = Format(Frm93.TB20, "0.00")  'Jumlah Bayaran Yang Perlu Dibuat (RM)
            Else
                rs!jumlah_perlu_bayar = Null 'Jumlah Bayaran Yang Perlu Dibuat (RM)
            End If
            rs!kuantiti_barang = 1 'Kuantiti Barang Yang Dijual
            If Frm93.CB19 = 1 Then 'Barang kemas
                If Frm93.TB1 <> vbNullString Then
                    rs!JUMLAH_BERAT = Format(Frm93.TB1, "0.00") 'Jumlah Berat Barang Yang Dijual
                Else
                    rs!JUMLAH_BERAT = Null 'Jumlah Berat Barang Yang Dijual
                End If
            ElseIf Frm93.CB20 = 1 Then 'Barang kemas
                If Frm93.TB8 <> vbNullString Then
                    rs!JUMLAH_BERAT = Format(Frm93.TB8, "0.00") 'Jumlah Berat Barang Yang Dijual
                Else
                    rs!JUMLAH_BERAT = Null 'Jumlah Berat Barang Yang Dijual
                End If
            Else
                rs!JUMLAH_BERAT = Null 'Jumlah Berat Barang Yang Dijual
            End If
            rs!gst_zr_cukai = Format(0, "0.00") 'Jumlah Cukai Bagi ZR
            rs!gst_sr_harga = Format(0, "0.00") 'Harga Keseluruhan Bagi Barang SR
            rs!gst_sr_cukai = Format(0, "0.00") 'Jumlah Cukai Bagi SR
            rs!caj_pos = "0.00"
            rs!no_tracking = Null
            rs!no_pekerja = Frm93_LM_EMP_NO 'No. Pekerja
            If Frm93.L36_Text <> vbNullString Then
                If Frm28.L5_Text <> vbNullString Then
                    rs!no_rujukan_pembeli = Frm28.L5_Text 'No. Rujukan Pembeli
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
        
                 
        '1:  Pelanggan
        '2:  Member
        '3:  RAF
        '4:  Pengedar
        '5:  Normal Dealer
        '6:  Master Dealer
        
        'Zakaria&Sons
        '1 : Pembeli biasa
        '2 : Ahli biasa
        '3 : Silver
        '4 : Gold
        '5 : Platinum
            rs!kategori_pembeli = Frm93_LM_KATEGORI
            rs!jualan_online = 0
            rs!point_ari_nashi = 0
            rs!jumlah_point = 0
            rs!kupon_diskaun = "0.00"
            rs!kadar_peroleh_point = 0
            rs!kadar_tebus_point = 0
            rs!kadar_diskaun = Format(0, "0.00") 'Kadar diskaun per gram
            rs!Status = 1
            rs!terminal = G_TERMINAL
            rs!cawangan = G_CAWANGAN
            rs!write_timestamp = LM_NOW
            rs!status_r = LM_STATUS_R_ASAL
            rs!Menu = 2
            
            DATA_SAVE = 1
            rs.Update
        
        End If
        
        rs.Close
        Set rs = Nothing
'###Masukkan Jualan Ke Dalam Table Akaun Jualan### - End

'###Masukkan Data Tempahan### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 40_tempahan_deposit where no_rujukan_tempahan='" & Frm93.L17_Text & "' AND status_invoice = 1", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
            rs.AddNew
            If Frm93.L17_Text <> vbNullString Then 'No. Rujukan Tempahan
                rs!no_rujukan_tempahan = Frm93_LM_No_RUJUKAN_TEMPAHAN
            Else
                rs!no_rujukan_tempahan = Null
            End If
            
            rs!no_resit_tempahan = LM_NO_INVOICE
            rs!Status = "Belum Siap"
            
            If Frm93.CB2 = 1 Then
                rs!jenis_tempahan = 0 'Jenis Tempahan , 0 : Tempahan Barang Baru , 1 : Tempahan Barang Kedai
                
                If Frm93.CBB1 <> vbNullString Then 'Kategori Produk
                    rs!kategori_Produk = Frm93.CBB1
                Else
                    rs!kategori_Produk = Null
                End If
                If Frm93.CBB2 <> vbNullString Then 'Purity
                    rs!purity = Frm93.CBB2
                Else
                    rs!purity = Null
                End If
                If Frm93.TB3 <> vbNullString Then 'Upah
                    rs!UPAH = Format(Frm93.TB3, "0.00")
                Else
                    rs!UPAH = Null
                End If
                If Frm93.TB4 <> vbNullString Then 'Anggaran Harga
                    rs!anggaran_harga = Format(Frm93.TB4, "0.00")
                Else
                    rs!anggaran_harga = Null
                End If
                    
                If Frm93.CB19 = 1 Then
                    rs!type_barang_kemas = 0 'Jenis Barang , 0 : Barang Kemas , 1 : Barang Permata
                    If Frm93.TB1 <> vbNullString Then 'Anggaran Berat
                        rs!anggaran_berat = Format(Frm93.TB1, "0.00")
                    Else
                        rs!anggaran_berat = Null
                    End If
                    If Frm93.TB2 <> vbNullString Then 'Harga Semasa
                        rs!harga_Semasa = Format(Frm93.TB2, "0.00")
                    Else
                        rs!harga_Semasa = Null
                    End If
                ElseIf Frm93.CB20 = 1 Then
                    rs!type_barang_kemas = 1 'Jenis Barang , 0 : Barang Kemas , 1 : Barang Permata
                End If
                'rs!status_tukang = "Belum Hantar" ' Belum Hantar , Belum Siap , Siap
                
                rs!Berat_Asal = Null
                rs!berat_jualan = Null
                rs!adjustment = Null
                rs!harga_asal = Null
            End If
            
            If Frm93.CB3 = 1 Then
                rs!jenis_tempahan = 1 'Jenis Tempahan , 0 : Tempahan Barang Baru , 1 : Tempahan Barang Kedai
                
                If Frm93.TB6 <> vbNullString Then 'No. Siri Produk
                    rs!no_siri_Produk = Frm93.TB6
                Else
                    rs!no_siri_Produk = Null
                End If
                If Frm93.L4_Text <> vbNullString Then 'Kategori Produk
                    rs!kategori_Produk = Frm93.L4_Text
                Else
                    rs!kategori_Produk = Null
                End If
                
                If Frm93.L13_Text = 0 Then
                    rs!type_barang_kemas = 0 'Jenis Barang , 0 : Barang Kemas , 1 : Barang Permata
                    If Frm93.TB7 <> vbNullString Then 'Berat Asal
                        rs!Berat_Asal = Format(Frm93.TB7, "0.00")
                    Else
                        rs!Berat_Asal = Null
                    End If
                    If Frm93.TB8 <> vbNullString Then 'Berat Jualan
                        rs!berat_jualan = Format(Frm93.TB8, "0.00")
                    Else
                        rs!berat_jualan = Null
                    End If
                    If Frm93.TB9 <> vbNullString Then 'Harga Semasa
                        rs!harga_Semasa = Format(Frm93.TB9, "0.00")
                    Else
                        rs!harga_Semasa = Null
                    End If
                ElseIf Frm93.L13_Text = 1 Then
                    rs!type_barang_kemas = 1 'Jenis Barang , 0 : Barang Kemas , 1 : Barang Permata
                    rs!Berat_Asal = Null
                    rs!berat_jualan = Null
                    rs!harga_Semasa = Null
                End If
                
                If Frm93.TB10 <> vbNullString Then 'Upah
                    rs!UPAH = Format(Frm93.TB10, "0.00")
                Else
                    rs!UPAH = Null
                End If
                If Frm93.TB11 <> vbNullString Then 'Harga Asal
                    rs!harga_asal = Format(Frm93.TB11, "0.00")
                Else
                    rs!harga_asal = Null
                End If
                If Frm93.TB12 <> vbNullString Then 'Adjustment
                    rs!adjustment = Format(Frm93.TB12, "0.00")
                Else
                    rs!adjustment = Null
                End If
                If Frm93.TB13 <> vbNullString Then 'Anggaran Harga
                    rs!anggaran_harga = Format(Frm93.TB13, "0.00")
                Else
                    rs!anggaran_harga = Null
                End If
                
                rs!anggaran_berat = Null
            End If
            
'Kategori Pembeli
'=================
'1:  Pelanggan
'2 : Member / Ahli
'3:  RAF
'4:  Pengedar
'5:  Normal Dealer (ND)
'6:  Master Dealer (MD)

            rs!kategori_pembeli = Frm93_LM_KATEGORI
            
            If Frm93.L36_Text <> vbNullString Then
            
                If Frm28.L5_Text <> vbNullString Then
                
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
                
            End If
            
            If Frm93.L35_Text <> vbNullString Then
                If Frm26.TB1 <> vbNullString Then
                    rs!no_rujukan_pelanggan = Null 'No. Rujukan Pembeli
                    rs!no_ic = Null 'Maklumat Pembeli : No. Kad Pengenalan
                    
                    If Frm26.TB1 <> vbNullString Then
                        rs!Nama = UCase(Frm26.TB1) 'Maklumat Pembeli : Nama
                    Else
                        rs!Nama = Null 'Maklumat Pembeli : Nama
                    End If
                    rs!no_ic = Null 'Maklumat Pembeli : No. Kad Pengenalan
                    If Frm26.TB2 <> vbNullString Then
                        rs!no_tel = UCase(Frm26.TB2) 'No. Telefon
                    Else
                        rs!no_tel = Null 'No. Telefon
                    End If
                End If
            End If
            If Frm93.CB6 = 1 Then 'Flag Trade In
                Frm93_LM_Flag_TRADE_IN = 1 '0 : Tiada Urusan Trade In , 1 : Ada Urusan Trade In
                
                rs!flag_trade_in = 1 '0 : Tiada Bayaran Deposit Menggunakan Trade In , 1 : Ada Bayaran Deposit Menggunakan Trade In
                If Frm93.L15_Text <> vbNullString Then 'No. Resit Trade In
                    rs!no_resit_trade_in = Frm93.L15_Text
                Else
                    rs!no_resit_trade_in = Null
                End If
                If Frm93.TB17 <> vbNullString Then 'Jumlah Nilaian Trade In
                    rs!nilaian_trade_in = Format(Frm93.TB17, "0.00")
                Else
                    rs!nilaian_trade_in = Null
                End If
            Else
                Frm93_LM_Flag_TRADE_IN = 0 '0 : Tiada Urusan Trade In , 1 : Ada Urusan Trade In
                
                rs!flag_trade_in = 0 '0 : Tiada Bayaran Deposit Menggunakan Trade In , 1 : Ada Bayaran Deposit Menggunakan Trade In
                rs!no_resit_trade_in = Null 'No. Resit Trade In
                rs!nilaian_trade_in = Null 'Jumlah Nilaian Trade In
            End If
            
            If Frm93.TB20 <> vbNullString Then 'Jumlah Deposit Yang Dibayar Secara Tunai
                rs!jumlah_deposit_tunai = Format(Frm93.TB20, "0.00")
            Else
                rs!jumlah_deposit_tunai = Null
            End If
            If Frm93.TB22 <> vbNullString Then 'Jumlah Deposit Dari Barangan Trade In
                rs!jumlah_deposit_trade_in = Format(Frm93.TB22, "0.00")
            Else
                rs!jumlah_deposit_trade_in = Null
            End If
            If Frm93.TB23 <> vbNullString Then 'Jumlah Deposit (Tanpa GST)
                rs!jumlah_tanpa_gst = Format(Frm93.TB23, "0.00")
                rs!jumlah_dengan_gst = Format(Frm93.TB23, "0.00")
            Else
                rs!jumlah_tanpa_gst = Null
            End If
            If Frm93.TB20 <> vbNullString Then
                rs!jumlah_perlu_bayar = Format(Frm93.TB20, "0.00") 'Jumlah Bayaran Sebelum Adjustment
                rs!jumlah_bayaran = Format(Frm93.TB20, "0.00") 'Jumlah Bayaran Deposit Selepas Adjustment
            Else
                rs!jumlah_perlu_bayar = Null
            End If
            rs!adjustment_bayaran = Format(0, "0.00") 'Adjustment Bagi Bayaran Keseluruhan
            If Frm93.TB33 <> vbNullString Then 'Remarks
                rs!remarks = Frm93.TB33
            Else
                rs!remarks = Null
            End If
            rs!tarikh = Frm93.DTPicker1 'Tarikh Tempahan
            rs!no_pekerja = Frm93_LM_EMP_NO 'No. Pekerja

            rs!terminal = G_TERMINAL
            rs!write_timestamp = LM_NOW
            rs!status_invoice = 1 '0 : Tidak aktif (dibatalkan) , 1:  Aktif
            If Frm93.CB9 = 0 Then
                rs!bil_rasmi = 1 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
            Else
                rs!bil_rasmi = 0 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
            End If
            rs!cawangan = G_CAWANGAN
            rs!no_staff = G_LOGIN_USER 'No. staff yang memasukkan/edit data (Login)
            rs.Update

        End If
        
        rs.Close
        Set rs = Nothing
'###Masukkan Data Tempahan### - End

'### Update Maklumat Trade In ### - Start
        If Frm93_LM_Flag_TRADE_IN = 1 Then '0 : Tiada Urusan Trade In , 1 : Ada Urusan Trade In
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & Frm93.L15_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
            
                G_ID = rs!ID
                Call recovery_16_gold_bar_belian
                
                rs!trade_in_status = 1
                rs!no_staff = Frm93_LM_EMP_NO 'No. Pekerja
                rs!terminal = G_TERMINAL
                rs!write_timestamp2 = LM_NOW
                rs!jenis_urusan = G_JENIS_URUSAN
                rs!remarks = "Edit deposit bagi tempahan emas"
                
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
        End If
'### Update Maklumat Trade In ### - End

'### Update Table Database Bagi Item Ini ### - Start
        If Frm93.CB3 = 1 Then
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from Data_Database where no_siri_produk='" & Frm93.TB6 & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
            
                G_ID = rs!ID
                Call recovery_data_database
                
                rs!StatusItem = 14
                rs!write_timestamp2 = LM_NOW
                rs!no_pekerja = Frm93_LM_EMP_NO
                rs!terminal = G_TERMINAL
                rs!Menu = 2
                
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
        End If
'### Update Table Database Bagi Item Ini ### - End

'###Masukkan Data Penjual Yang Tidak Berdaftar Ke Dalam Database### - Start
        'If Frm28.L5_Text = vbNullString And Frm26.TB1 <> vbNullString Then
        If Frm93.L35_Text <> vbNullString And Frm93.L36_Text = vbNullString Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 44_senarai_pelanggan", cn, adOpenKeyset, adLockOptimistic
            
            rs.AddNew
            rs!tarikh = Frm93.DTPicker1 'Tarikh
            rs!no_resit = Frm93_LM_No_RESIT_TEMPAHAN 'No. Resit Tempahan
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
            rs!no_staff = Frm93_LM_EMP_NO 'No. Pekerja
            rs!terminal = G_TERMINAL
            rs!jenis_urusan = G_JENIS_URUSAN
            rs!cawangan = G_CAWANGAN
            rs.Update
            
            rs.Close
            Set rs = Nothing
            
        End If
'###Masukkan Data Penjual Yang Tidak Berdaftar Ke Dalam Database### - End

'###Update Data Simpanan Duit Pelanggan### - Start
        If Frm93_LM_Flag_SIMPANAN = 1 Then '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
            
                G_ID = rs!ID
                Call recovery_senarai_pelanggan
                
                Frm93_LM_JUMLAH_SIMPANAN = frm130.L26_Text  'Jumlah Simpanan Yang Ada
                Frm93_LM_GUNA_SIMPAN = frm130.TB21  'Jumlah Simpanan Yang Hendak Digunakan
                
                rs!baki_simpanan = Format(Frm93_LM_JUMLAH_SIMPANAN - Frm93_LM_GUNA_SIMPAN, "0.00") 'Baki Simpanan

                rs!write_timestamp2 = LM_NOW
                rs!no_staff = Frm93_LM_EMP_NO 'No. Pekerja
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
            rs!tarikh = Frm93.DTPicker1 'Tarikh
            rs!jenis = 1 '0 : Simpanan , 1 : Penggunaan Duit
            rs!no_rujukan_pelanggan = Frm28.L5_Text 'No. Rujukan Pelanggan
            rs!no_resit = Frm93_LM_No_RESIT_TEMPAHAN 'No. Resit Tempahan
            rs!jumlah = Format(frm130.TB21, "0.00") 'Jumlah Simpanan Yang Digunakan (RM)
            rs!jenis_penggunaan = 2 '0 : Belian Barangan Kemas , 1 : Ansuran , 2 : Tempahan (Deposit) , 3 : Servis , 4 : Tempahan (Ambilan Barang)
            rs!no_rujukan_pekerja = Frm93_LM_EMP_NO 'No. Pekerja
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

        user = MDI_frm1.L3_Text
        LogAct_Memory = "[" & G_LOGIN_USER & "] Edit data tempahan , No. invoice [" & Frm93_LM_No_RESIT_TEMPAHAN & "]."
        LogDate_Memory = DateTime.Date & " " & DateTime.Time$
        Call UpdateLog_Database
        
        Call Frm93_initial_setting
        
        GM_NEXT_PREV = 2
        
        Call frm93_tempahan_header
        Call frm93_tempahan
        
        Frm93.Frame2.Visible = False
        Frm93.Frame1.Visible = True
                    
        MsgBox "Data telah berjaya disimpan.", vbInformation, "Info"
    End If
End If

End Sub
Private Sub CMD15_Click()
'on error resume next
Note = "Adakah anda ingin batalkan urusan edit data ini?" & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then
    Frm93.Frame2.Visible = False
    Frm93.Frame1.Visible = True
End If
End Sub
Private Sub CMD3_Click()
'on error resume next
Frm68.Show
Frm87.Hide
Frm68.L15_Text = 7

'0 : Pendaftaran Biasa
'1 : Jualan Gold Bar
'2 : Buyback Gold Bar
'3 : Jualan BK
'4 : Buyback BK
'5 : Ansuran
'6 : Servis
'7 : Tempahan
End Sub



Private Sub CMD19_Click()
'On Error Resume Next
If Frm93.L35_Text = vbNullString Then
    
    If Frm93.L36_Text <> vbNullString Then
    
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
            
            Frm93.L36_Text = vbNullString 'Nama pembeli : Berdaftar
            
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
Private Sub CMD20_Click()
'On Error Resume Next
Frm26.Hide
Frm28.Hide
Frm27.Show

Frm27.TB1.SetFocus
End Sub
Private Sub CMD21_Click()
'On Error Resume Next
If Frm93.L36_Text = vbNullString Then
    
    If Frm93.L35_Text <> vbNullString Then
    
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
            
            Frm93.L35_Text = vbNullString 'Nama pembeli : Tidak berdaftar
            
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
Dim Frm93_LM_CURR_PAGE As Double
Dim Frm93_LM_TOTAL_PAGE As Double

Frm93_LM_CURR_PAGE = 0
Frm93_LM_TOTAL_PAGE = 0

If Frm93.L60_Text <> vbNullString And IsNumeric(Frm93.L60_Text) Then
    If Frm93.L61_Text <> vbNullString And IsNumeric(Frm93.L61_Text) Then
        Frm93_LM_CURR_PAGE = Frm93.L60_Text
        Frm93_LM_TOTAL_PAGE = Frm93.L61_Text
        
        If Frm93_LM_CURR_PAGE <> 1 And Frm93_LM_CURR_PAGE <> 0 Then
        
        GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
        
        Call frm93_tempahan_header
        Call frm93_tempahan
            
        End If
    End If
End If
End Sub

Private Sub CMD26_Click()
'on error resume next
Dim Frm93_LM_CURR_PAGE As Double
Dim Frm93_LM_TOTAL_PAGE As Double

Frm93_LM_CURR_PAGE = 0
Frm93_LM_TOTAL_PAGE = 0

If Frm93.L60_Text <> vbNullString And IsNumeric(Frm93.L60_Text) Then
    If Frm93.L61_Text <> vbNullString And IsNumeric(Frm93.L61_Text) Then
        Frm93_LM_CURR_PAGE = Frm93.L60_Text
        Frm93_LM_TOTAL_PAGE = Frm93.L61_Text
        
        If Frm93_LM_CURR_PAGE < Frm93_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call frm93_tempahan_header
            Call frm93_tempahan
            
        End If
    End If
End If
End Sub

Private Sub CMD4_Click()
'on error resume next
Dim Frm93_LM_FIELD As String

If Frm93.CB5 = 1 Then
    Frm93_LM_FIELD = "no_ic"
ElseIf Frm93.CB4 = 1 Then
    Frm93_LM_FIELD = "no_pelanggan"
End If

LM_FOUND = 0

If Frm93.TB41 <> vbNullString Then

    If InStr(1, Frm93.TB41, "*") <> 0 Or InStr(1, Frm93.TB41, "/") <> 0 Or InStr(1, Frm93.TB41, "\") <> 0 Or InStr(1, Frm93.TB41, "'") <> 0 Then
        MsgBox "Nombor keahlian/No. kad pengenalan mengandungi simbol yang tidak sah.", vbExclamation, "Info"
        
        Frm93.TB41.SetFocus
        Exit Sub
    End If
    
End If

If Frm93.TB41 <> vbNullString Then

    Note = "Sistem akan mencari maklumat berkenaan dengan ahli ini." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Jika maklumat dijumpai , sistem akan mengisi maklumat ahli ini ke dalam ruangan maklumat ahli." & vbCrLf & _
            "Penetapan harga (Terutama pada harga emas semasa dan upah) akan mengikut kategori ahli ini." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan?."
            
Else

    Note = "Sistem akan meneruskan tempahan ini sebagai PELANGGAN BIASA." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan?."

End If
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then
    
    Call frm130_initial_setting
    
    If Frm93.TB41 <> vbNullString Then
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where " & Frm93_LM_FIELD & "='" & UCase(Frm93.TB41) & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
        
            Call Frm28_initial
            
            If Not IsNull(rs!Nama) Then Frm28.L1_Text = rs!Nama 'Nama
            If Not IsNull(rs!no_ic) Then Frm28.L2_Text = rs!no_ic 'No. Kad Pengenalan
            If Not IsNull(rs!no_tel) Then Frm28.L3_Text = rs!no_tel 'No. Telefon
            If Not IsNull(rs!Email) Then Frm28.L4_Text = rs!Email 'E-mail
            If Not IsNull(rs!no_pelanggan) Then Frm28.L5_Text = rs!no_pelanggan 'No. Ahli
            If Not IsNull(rs!baki_simpanan) Then
            
                If Not IsNull(rs!baki_simpanan) Then
                
                    frm130.L26_Text = Format(rs!baki_simpanan, "#,##0.00") 'Baki Simpanan Pelanggan Ini (RM)
                
                Else
                
                    frm130.L26_Text = Format(0, "#,##0.00") 'Baki Simpanan Pelanggan Ini (RM)
                
                End If
                
            End If
            If Not IsNull(rs!kategori_pelanggan) Then
                
                If rs!kategori_pelanggan = 1 Then
                    Frm93.CB13 = 1
                ElseIf rs!kategori_pelanggan = 2 Then
                    Frm93.CB14 = 1
                ElseIf rs!kategori_pelanggan = 3 Then
                    Frm93.CB15 = 1
                ElseIf rs!kategori_pelanggan = 4 Then
                    Frm93.CB16 = 1
                ElseIf rs!kategori_pelanggan = 5 Then
                    Frm93.CB17 = 1
                End If
                
            End If
            
            LM_FOUND = 1
        
        Else
        
            MsgBox "Tiada maklumat dijumpai ATAU data pelanggan ini sudah tidak aktif.", vbInformation, "Info"
            Frm93.TB41.SetFocus
            
        End If
        
        rs.Close
        Set rs = Nothing
        
    Else
    
        LM_FOUND = 2
    
    End If
    
    
    If LM_FOUND = 1 Or LM_FOUND = 2 Then
        
        If LM_FOUND = 1 Then
        
            Frm93.CMD19.Enabled = False
            'Frm93.CMD21.Enabled = True
        
        ElseIf LM_FOUND = 2 Then
            
            Frm93.CB13 = 1
            
            Frm93.CMD19.Enabled = True
            Frm93.CMD21.Enabled = False
            
        End If
        
        'Call frm130_initial_setting
        
        Frm93.Frame2.Visible = True
        Frm93.Pic4.Visible = False
        
        Frm93.CB3 = 1
        
    End If

End If
End Sub
Private Sub CMD5_Click()
'on error resume next
If Frm93.TB5 = vbNullString Then
    MsgBox "Sila masukkan No. Siri Produk.", vbInformation, "Info"
    Exit Sub
End If

If Frm93.TB5 <> vbNullString Then
    If InStr(1, Frm93.TB5, "*") <> 0 Or InStr(1, Frm93.TB5, "/") <> 0 Or InStr(1, Frm93.TB5, "\") <> 0 Or InStr(1, Frm93.TB5, "'") <> 0 Then
        MsgBox "No. Siri Produk mengandungi simbol yang tidak sah.", vbExclamation, "Info"
        
        Frm93.TB5.SetFocus
        Exit Sub
    End If
End If

Call Frm93_Call_Product_Detail
End Sub

Private Sub CMD6_Click()
'on error resume next
'frm130.TB33 = Format(Frm92.L10_Text, "#,##0.00")
frm130.Show vbModal
End Sub

Private Sub CMD9_Click()
'on error resume next
DATA_FOUND = 0

If Frm93.TB18 = vbNullString Then
    MsgBox "Sila masukkan No. Voucher trade in.", vbInformation, "Info"
    Exit Sub
End If
If Frm93.TB18 <> vbNullString Then
    If InStr(1, Frm93.TB18, "*") <> 0 Or InStr(1, Frm93.TB18, "/") <> 0 Or InStr(1, Frm93.TB18, "\") <> 0 Or InStr(1, Frm93.TB18, "'") <> 0 Then
        MsgBox "No. Voucher trade in mengandungi simbol yang tidak sah.", vbExclamation, "Info"
        
        Frm93.TB18.SetFocus
        Exit Sub
    End If
End If

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & UCase(Frm93.TB18) & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Not IsNull(rs!trade_in_status) Then
        If rs!trade_in_status = 0 Then
            
            If Not IsNull(rs!no_resit_trade_in) Then Frm93.L15_Text = rs!no_resit_trade_in
            If Not IsNull(rs!jumlah_tanpa_gst) Then Frm93.TB17 = Format(rs!jumlah_tanpa_gst, "#,##0.00") 'Jumlah Nilaian Resit Trade In
            'Frm93.L15_Text = UCase(Frm93.TB18) 'No. Resit Trade In
            
            DATA_FOUND = 1
            
        ElseIf rs!trade_in_status = 1 Then
        
            MsgBox "No. Voucher Trade In ini telah digunakan untuk urusan belian sebelum ini.", vbInformation, "Info"
            
            Frm93.TB18 = vbNullString
            Frm93.TB18.SetFocus
            
        End If
    End If
    
Else

    MsgBox "No. Voucher tidak dijumpai.", vbInformation, "Info"
    
    Frm93.TB18 = vbNullString
    Frm93.TB18.SetFocus
    
End If

rs.Close
Set rs = Nothing
'### Carian Maklumat Penjual Bagi Buyback ### - End
End Sub

Private Sub Form_Load()
'on error resume next
Frm93.L27_Text = 0 '0 : Jenis Tempahan , Status , 1:  No.Siri Produk , 2:  No.Invoice
Frm93.L13_Text = 2

frm130.L31_Text.BackStyle = 0
frm130.L32_Text.BackStyle = 0
frm130.L81_Text.BackStyle = 0
frm130.L82_Text.BackStyle = 0

user = MDI_frm1.L3_Text

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from tblelogin where username='" & user & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!usertype) Then
        If rs!usertype = "Developer" Or rs!usertype = "Admin" Then
            Frm93.Frm93_SM_padam.Enabled = True
        Else
            Frm93.Frm93_SM_padam.Enabled = False
        End If
    End If
End If

rs.Close
Set rs = Nothing

Call frm93_setting_report
End Sub



Private Sub Frm93_SM_belum_siap_Click()
'on error resume next
L_FOUND = 0

frm93_LM_No_ID = vbNullString

If IsNumeric(Frm93.LV1.SelectedItem.Index) Then
    
    frm93_LM_No_ID = Frm93.LV1.ListItems(Frm93.LV1.SelectedItem.Index)
    
    If frm93_LM_No_ID <> vbNullString Then
        
        Note = "Adakah anda ingin tukar status tempahan ini kepada belum siap?" & vbCrLf & _
                "Invoice tempahan siap akan dipadamkan dari sistem." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
        
'### Padam Data Dari Senarai Tempahan (Deposit) ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 40_tempahan_deposit where ID='" & frm93_LM_No_ID & "' AND status_invoice = 1 AND status='" & "Siap" & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                L_FOUND = 1
                
            End If
            
            rs.Close
            Set rs = Nothing
            
            If L_FOUND = 1 Then

                G_TEMPAHAN = 1 '0 : Padam data , 1 : Tukar status kepada belum siap
                Call Frm93_padam_data_tempahan
                
            Else
            
                MsgBox "Status barang ini adalah belum siap. Anda tidak dibenarkan untuk ubah status barang ini kepada belukm siap.", vbExclamation, "Info"
                
            End If
            
        End If
        
    End If
    
End If
End Sub
Private Sub Frm93_SM_Cetak_Invoice_Deposit_Click()
'on error resume next
DATA_FOUND = 0

Frm93_LM_ID = vbNullString

If IsNumeric(Frm93.LV1.SelectedItem.Index) Then
    
    Frm93_LM_ID = Frm93.LV1.ListItems(Frm93.LV1.SelectedItem.Index)
    
    If Frm93_LM_ID <> vbNullString Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 40_tempahan_deposit where ID='" & Frm93_LM_ID & "' AND status_invoice = 1", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!cawangan) Then G_KEDAI = rs!cawangan
            If Not IsNull(rs!no_resit_tempahan) Then
                G_No_INV_BOOK = vbNullString
                
                G_No_INV_BOOK = rs!no_resit_tempahan 'No. Invoice
                
                DATA_FOUND = 1
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then
            G_PREVIEW = 1
            Call Frm94_invoice_deposit_tempahan
        End If
        
    End If
End If
End Sub
Private Sub Frm93_SM_Cetak_Invoice_Siap_Click()
'on error resume next
DATA_FOUND = 0
Frm93_LM_STATUS = vbNullString

Frm93_LM_ID = vbNullString

If IsNumeric(Frm93.LV1.SelectedItem.Index) Then
    
    Frm93_LM_ID = Frm93.LV1.ListItems(Frm93.LV1.SelectedItem.Index)
    
    If Frm93_LM_ID <> vbNullString Then
            
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 40_tempahan_deposit where ID='" & Frm93_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!Status) Then
                
                If rs!Status = "Belum Siap" Then
                    
                    MsgBox "Tiada invoice tempahan siap kerana status tempahan ini adalah BELUM SIAP.", vbExclamation, "Info"
                    
                    rs.Close
                    Set rs = Nothing
                    
                    Exit Sub
                
                End If
            
            End If
            If Not IsNull(rs!no_rujukan_tempahan) Then Frm93_LM_No_RUJUKAN = rs!no_rujukan_tempahan
        End If
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 42_tempahan_siap where no_rujukan_tempahan='" & Frm93_LM_No_RUJUKAN & "' AND status_invoice = 1", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!cawangan) Then G_KEDAI = rs!cawangan
            If Not IsNull(rs!no_resit_tempahan) Then
                G_No_INV_BOOK = vbNullString
                
                G_No_INV_BOOK = rs!no_resit_tempahan 'No. Invoice
                
                DATA_FOUND = 1
            End If
            
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then
            G_PREVIEW = 1
            Call Frm94_invoice_siap_tempahan
        End If

    End If
End If
End Sub
Private Sub Frm93_SM_Edit_Click()
'on error resume next
Dim Frm93_LM_SIMPANAN_ASAL As Double
Dim Frm93_LM_SIMPANAN_DIGUNAKAN As Double

DATA_FOUND = 0
Frm93_LM_No_PEKERJA = vbNullString
Frm93_LM_ID = vbNullString
Frm93_LM_No_PEMBELI = vbNullString
Frm93_LM_STATUS = vbNullString
Frm93_LM_KATEGORI = 0
Frm93_LM_SIMPANAN_ASAL = 0
Frm93_LM_SIMPANAN_DIGUNAKAN = 0
Frm93_LM_KATEGORI_PEMBELI = 0 '0 : Data Tidak Dijumpai , 1 : Pembeli Tidak Berdaftar , 2 : Pembeli Berdaftar , 3 : Ahli
'Frm93_LM_KATEGORI = 1

Frm93_LM_ID = vbNullString

If IsNumeric(Frm93.LV1.SelectedItem.Index) Then
    
    Frm93_LM_ID = Frm93.LV1.ListItems(Frm93.LV1.SelectedItem.Index)
    
    If Frm93_LM_ID <> vbNullString Then
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 40_tempahan_deposit where ID='" & Frm93_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            If Not IsNull(rs!status_invoice) Then
                
                If rs!status_invoice = 0 Then
                    
                    MsgBox "Status deposit tempahan / data tempahan ini telah bertukar status kepada TIDAK AKTIF/TELAH DIPADAMKAN." & vbCrLf & _
                            vbNullString & vbCrLf & _
                            "Oleh itu anda tidak dibenarkan untuk edit data ini.", vbExclamation, "Info"
                            
                    rs.Close
                    Set rs = Nothing
                    
                    Exit Sub
                    
                End If
            
            End If
            
            If Not IsNull(rs!Status) Then
                
                If rs!Status = "Siap" Then
                    
                    MsgBox "Status deposit tempahan / data tempahan ini telah bertukar status kepada SIAP." & vbCrLf & _
                            vbNullString & vbCrLf & _
                            "Oleh itu anda tidak dibenarkan untuk edit data ini.", vbExclamation, "Info"
                            
                    rs.Close
                    Set rs = Nothing
                    
                    Exit Sub
                    
                End If
            
            End If
            
        End If
        
        rs.Close
        Set rs = Nothing

        Call Frm93_initial_setting
        Call frm130_initial_setting
        
        frm130.L41_Text = "1"
        Unload Frm26
        Unload Frm27
        Unload Frm28
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 40_tempahan_deposit where ID='" & Frm93_LM_ID & "' AND status_invoice = 1", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            GLOBAL_DISABLE = 1
            
            If Not IsNull(rs!no_rujukan_tempahan) Then 'No. Rujukan Tempahan
                Frm93.L17_Text = rs!no_rujukan_tempahan
            Else
                Frm93.L17_Text = vbNullString
            End If
            If Not IsNull(rs!no_resit_tempahan) Then 'No. Resit Tempahan
                Frm93.L18_Text = rs!no_resit_tempahan
            Else
                Frm93.L18_Text = vbNullString
            End If
                    
            If Not IsNull(rs!jenis_tempahan) Then
                If rs!jenis_tempahan = 0 Then 'Jenis Tempahan , 0 : Tempahan Barang Baru , 1 : Tempahan Barang Kedai
                    Frm93.CB2 = 1
                    
                    Frm93.Frame4.Visible = True
                    Frm93.Frame5.Visible = False
                    
                    If Not IsNull(rs!kategori_Produk) Then 'Kategori Produk
                        'On Error GoTo Err_A:
                        Frm93_LM_KATEGORI_PRODUK = rs!kategori_Produk
                        Frm93.CBB1 = Frm93_LM_KATEGORI_PRODUK
Restore_A:
                    End If
                    
                    If Not IsNull(rs!purity) Then 'Purity
                        'On Error GoTo Err_B:
                        Frm93_LM_PURITY = rs!purity
                        Frm93.CBB2 = Frm93_LM_PURITY
Restore_B:
                    End If
                    If Not IsNull(rs!UPAH) Then 'Upah
                        Frm93.TB3 = Format(rs!UPAH, "#,##0.00")
                    Else
                        Frm93.TB3 = vbNullString
                    End If
                    If Not IsNull(rs!anggaran_harga) Then 'Anggaran Harga
                        Frm93.TB4 = Format(rs!anggaran_harga, "#,##0.00")
                    Else
                        Frm93.TB4 = vbNullString
                    End If
                            
                    If Not IsNull(rs!type_barang_kemas) Then
                        If rs!type_barang_kemas = 0 Then
                            'Frm93.L13_Text = 0
                            Frm93.CB19 = 1
                            Frm93.CB20 = 0
                            If Not IsNull(rs!anggaran_berat) Then 'Anggaran Berat
                                Frm93.TB1 = Format(rs!anggaran_berat, "#,##0.00")
                            Else
                                Frm93.TB1 = vbNullString
                            End If
                            If Not IsNull(rs!harga_Semasa) Then 'Harga Semasa
                                Frm93.TB2 = Format(rs!harga_Semasa, "#,##0.00")
                            Else
                                Frm93.TB2 = vbNullString
                            End If
                        Else
                            'Frm93.L13_Text = 1
                            Frm93.CB19 = 0
                            Frm93.CB20 = 1
                            
                            Frm93.TB1 = vbNullString 'Anggaran Berat
                            Frm93.TB2 = vbNullString 'Harga Semasa
                            
                            Frm93.TB1.Locked = True
                            Frm93.TB2.Locked = True
                            Frm93.TB4.Locked = False
                            
                            Frm93.TB1.BackColor = &H8000000A
                            Frm93.TB2.BackColor = &H8000000A
                            Frm93.TB4.BackColor = &HFFFFFF
                        End If
                    End If
                ElseIf rs!jenis_tempahan = 1 Then 'Jenis Tempahan , 0 : Tempahan Barang Baru , 1 : Tempahan Barang Kedai
                    Frm93.CB3 = 1
                    
                    Frm93.Frame4.Visible = False
                    Frm93.Frame5.Visible = True
                    
                    If Not IsNull(rs!no_siri_Produk) Then 'No. Siri Produk
                        Frm93.TB6 = rs!no_siri_Produk
                    Else
                        Frm93.TB6 = vbNullString
                    End If
                    If Not IsNull(rs!kategori_Produk) Then 'Kategori Produk
                        Frm93.L4_Text = rs!kategori_Produk
                    Else
                        Frm93.L4_Text = vbNullString
                    End If
                                     
                    If Not IsNull(rs!type_barang_kemas) Then
                        If rs!type_barang_kemas = 0 Then 'Jenis Barang , 0 : Barang Kemas , 1 : Barang Permata
                            Frm93.L13_Text = 0
                            If Not IsNull(rs!Berat_Asal) Then 'Berat Asal
                                Frm93.TB7 = Format(rs!Berat_Asal, "#,##0.00")
                            Else
                                Frm93.TB7 = vbNullString
                            End If
                            If Not IsNull(rs!berat_jualan) Then 'Berat Jualan
                                Frm93.TB8 = Format(rs!berat_jualan, "#,##0.00")
                            Else
                                Frm93.TB8 = vbNullString
                            End If
                            If Not IsNull(rs!harga_Semasa) Then 'Harga Semasa
                                Frm93.TB9 = Format(rs!harga_Semasa, "#,##0.00")
                            Else
                                Frm93.TB9 = vbNullString
                            End If
                            
                            Frm93.TB8.Locked = False
                            Frm93.TB9.Locked = False
                            Frm93.TB10.Locked = False
                            Frm93.TB11.Locked = True
                            
                            Frm93.TB8.BackColor = &HFFFFFF
                            Frm93.TB9.BackColor = &HFFFFFF
                            Frm93.TB10.BackColor = &HFFFFFF
                            Frm93.TB11.BackColor = &H8000000A
                        Else
                            Frm93.L13_Text = 1
                            
                            'Frm93.L4_Text = vbNullString
                            Frm93.TB7 = vbNullString
                            Frm93.TB8 = vbNullString
                            Frm93.TB9 = vbNullString
                            
                            Frm93.TB7.Locked = True
                            Frm93.TB8.Locked = True
                            Frm93.TB9.Locked = True
                            Frm93.TB11.Locked = False
                            
                            Frm93.TB7.BackColor = &H8000000A
                            Frm93.TB8.BackColor = &H8000000A
                            Frm93.TB9.BackColor = &H8000000A
                            Frm93.TB11.BackColor = &HFFFFFF
                        End If
                    End If
                    If Not IsNull(rs!UPAH) Then 'Upah
                        Frm93.TB10 = Format(rs!UPAH, "#,##0.00")
                    Else
                        Frm93.TB10 = vbNullString
                    End If
                    If Not IsNull(rs!harga_asal) Then 'Harga Asal
                        Frm93.TB11 = Format(rs!harga_asal, "#,##0.00")
                    Else
                        Frm93.TB11 = vbNullString
                    End If
                    If Not IsNull(rs!adjustment) Then 'Adjustment
                        Frm93.TB12 = Format(rs!adjustment, "#,##0.00")
                    Else
                        Frm93.TB12 = vbNullString
                    End If
                    If Not IsNull(rs!anggaran_harga) Then 'Anggaran Harga
                        Frm93.TB13 = Format(rs!anggaran_harga, "#,##0.00")
                    Else
                        Frm93.TB13 = vbNullString
                    End If
                    
                    
                End If
            End If
                    
            If Not IsNull(rs!no_rujukan_pelanggan) Then Frm93_LM_No_PEMBELI = rs!no_rujukan_pelanggan 'No. Rujukan Pembeli
            If Not IsNull(rs!flag_trade_in) Then
                If rs!flag_trade_in = 0 Then '0 : Tiada Urusan Trade In , 1 : Ada Urusan Trade In
                    Frm93.CB6 = 0
                ElseIf rs!flag_trade_in = 1 Then
                    Frm93.CB6 = 1
                    If Not IsNull(rs!no_resit_trade_in) Then 'No. Resit Trade In
                        Frm93.L15_Text = rs!no_resit_trade_in
                    Else
                        Frm93.L15_Text = vbNullString
                    End If
                    If Not IsNull(rs!nilaian_trade_in) Then 'Jumlah Nilaian Trade In
                        Frm93.TB17 = Format(rs!nilaian_trade_in, "#,##0.00")
                    Else
                        Frm93.TB17 = "0.00"
                    End If
                End If
            End If

            If Not IsNull(rs!jumlah_deposit_tunai) Then 'Jumlah Deposit Yang Dibayar Secara Tunai
                Frm93.TB20 = Format(rs!jumlah_deposit_tunai, "#,##0.00")
            Else
                Frm93.TB20 = "0.00"
            End If
            If Not IsNull(rs!jumlah_deposit_trade_in) Then 'Jumlah Deposit Dari Barangan Trade In
                Frm93.TB22 = Format(rs!jumlah_deposit_trade_in, "#,##0.00")
            Else
                Frm93.TB22 = "0.00"
            End If
            If Not IsNull(rs!jumlah_tanpa_gst) Then 'Jumlah Deposit (Tanpa GST)
                Frm93.TB23 = Format(rs!jumlah_tanpa_gst, "#,##0.00")
            Else
                Frm93.TB23 = "0.00"
            End If
            'If Not IsNull(rs!jumlah_dengan_gst) Then 'Jumlah Deposit Dengan GST
            '    Frm93.TB19 = rs!jumlah_dengan_gst
            'Else
            '    Frm93.TB19 = "0.00"
            'End If
            'If Not IsNull(rs!jumlah_perlu_bayar) Then 'Jumlah Bayaran Sebelum Adjustment
            '    Frm93.L23_Text = rs!jumlah_perlu_bayar
            'Else
            '    Frm93.L23_Text = "0.00"
            'End If
            'If Not IsNull(rs!adjustment_bayaran) Then 'Adjustment Bagi Bayaran Keseluruhan
            '    Frm93.TB24 = rs!adjustment_bayaran
            'Else
            '    Frm93.TB24 = "0.00"
            'End If
            'If Not IsNull(rs!jumlah_bayaran) Then 'Jumlah Bayaran Deposit Selepas Adjustment
            '    Frm93.L24_Text = rs!jumlah_bayaran
            'Else
            '    Frm93.L24_Text = "0.00"
            'End If
            If Not IsNull(rs!remarks) Then 'Remarks
                Frm93.TB33 = rs!remarks
            Else
                Frm93.TB33 = vbNullString
            End If
            If Not IsNull(rs!tarikh) Then Frm93.DTPicker1 = rs!tarikh 'Tarikh Tempahan

            If Not IsNull(rs!no_pekerja) Then Frm93_LM_No_PEKERJA = rs!no_pekerja 'No. Pekerja
            
            DATA_FOUND = 1
            GLOBAL_DISABLE = 0
        End If
        
        rs.Close
        Set rs = Nothing
                
'### Carian Maklumat Penjual (Data Pekerja) ### - Start
        If Frm93_LM_No_PEKERJA <> vbNullString Then
            DATA_PEKERJA_FOUND = 0
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from employee where NoPekerja='" & Frm93_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                Frm93_LM_MAKLUMAT_PEKERJA = rs!Samaran & "  |  " & rs!NoPekerja
                DATA_PEKERJA_FOUND = 1
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_PEKERJA_FOUND = 1 Then
                'On Error GoTo Err_C:
                Frm93.CBB3 = Frm93_LM_MAKLUMAT_PEKERJA
Restore_C:
            End If
        End If
'### Carian Maklumat Penjual (Data Pekerja) ### - End

        '###Makluamt invoice### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 22_jualan where no_resit='" & Frm93.L18_Text & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
                
        If Not rs.EOF Then
            GLOBAL_DISABLE = 1
            If Not IsNull(rs!bil_rasmi) Then
                If rs!bil_rasmi = 1 Then
                    Frm93.CB9 = 0
                ElseIf rs!bil_rasmi = 0 Then
                    Frm93.CB9 = 1
                End If
            End If
            If Not IsNull(rs!tunai) Then 'Cara Bayaran : Tunai
                frm130.TB27 = Format(rs!tunai, "#,##0.00")
            Else
                frm130.TB27 = Format(0, "#,##0.00")
            End If
            If Not IsNull(rs!bank_in) Then 'Cara Bayaran : Bank In
                frm130.TB28 = Format(rs!bank_in, "#,##0.00")
            Else
                frm130.TB28 = Format(0, "#,##0.00")
            End If
            If Not IsNull(rs!kad_kredit) Then 'Cara Bayaran : Kad Kredit
                frm130.TB29 = rs!kad_kredit
            Else
                frm130.TB29 = "0.00"
            End If
        
            'On Error GoTo Err_D:
            If Not IsNull(rs!jenis_kad) Then
                Frm93_LM_JENIS_KAD = rs!jenis_kad
                frm130.CBB2 = Frm93_LM_JENIS_KAD
                
Restore_D:
            End If
            'on error resume next
        
            If Not IsNull(rs!cas_Kad_Kredit) Then 'Cara Bayaran : Cas Kad Kredit (%)
                frm130.L31_Text = Format(rs!cas_Kad_Kredit, "#,##0.00")
            Else
                frm130.L31_Text = "0.00"
            End If
            If Not IsNull(rs!jumlah_cas_kad_kredit) Then 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
                frm130.L32_Text = Format(rs!jumlah_cas_kad_kredit, "#,##0.00")
            Else
                frm130.L32_Text = "0.00"
            End If
            If Not IsNull(rs!gst_kad_kredit) Then 'Cara Bayaran : Jumlah GST kad kredit (RM)
                frm130.L81_Text = Format(rs!gst_kad_kredit, "#,##0.00")
            Else
                frm130.L81_Text = "0.00"
            End If
            If Not IsNull(rs!jumlah_potongan_kad_kredit) Then 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
                frm130.L82_Text = Format(rs!jumlah_potongan_kad_kredit, "#,##0.00")
            Else
                frm130.L82_Text = "0.00"
            End If
            If Not IsNull(rs!duit_simpanan_kedai) Then 'Cara Bayaran : Simpanan Duit Di Kedai
                frm130.TB21 = Format(rs!duit_simpanan_kedai, "#,##0.00")
            
                If rs!duit_simpanan_kedai <> "0.00" Then
                    Frm93_LM_Flag_SIMPANAN = 1 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
                    Frm93_LM_SIMPANAN_DIGUNAKAN = rs!duit_simpanan_kedai 'Jumlah Simpanan Yang Digunakan (RM)
                End If
            
            Else
                frm130.TB21 = "0.00"
            End If
            If Not IsNull(rs!jumlah_bayaran) Then 'Cara Bayaran : Jumlah Bayaran
                Frm93.TB32 = Format(rs!jumlah_bayaran, "#,##0.00")
            Else
                Frm93.TB32 = Format(0, "#,##0.00")
            End If
            If Not IsNull(rs!jumlah_perlu_bayar) Then 'Jumlah Bayaran Yang Perlu Dibuat (RM)
                Frm93.TB20 = Format(rs!jumlah_perlu_bayar, "#,##0.00")
            Else
                Frm93.TB20 = Format(0, "#,##0.00")
            End If
            If Not IsNull(rs!kategori_pembeli) Then Frm93_LM_KATEGORI = rs!kategori_pembeli
            GLOBAL_DISABLE = 0
        End If
        
        rs.Close
        Set rs = Nothing
        '###Makluamt invoice### - End
                
        If Frm93_LM_KATEGORI = 0 Then
            Frm93.CMD19.Enabled = True
            Frm93.CMD21.Enabled = False
            'Frm93.CMD19.Enabled = False
            'Frm93.CMD21.Enabled = True
            Frm93.CB13 = 1
            
        Else
            
            If Frm93_LM_KATEGORI = 1 Then
                Frm93.CB13 = 1
                Frm93.CMD19.Enabled = True
                Frm93.CMD21.Enabled = False
            ElseIf Frm93_LM_KATEGORI = 2 Then
                Frm93.CB14 = 1
            ElseIf Frm93_LM_KATEGORI = 3 Then
                Frm93.CB15 = 1
            ElseIf Frm93_LM_KATEGORI = 4 Then
                Frm93.CB16 = 1
            ElseIf Frm93_LM_KATEGORI = 5 Then
                Frm93.CB17 = 1
            End If
    
            Frm93.CMD19.Enabled = False
            'Frm93.CMD21.Enabled = True
        End If

        '### Maklumat Pembeli Yang Tidak Berdaftar Dengan Kedai ### - Start (08-07-2015)
        If Frm93_LM_KATEGORI = 0 And Frm93_LM_No_PEMBELI = vbNullString Then '0 : Data Tidak Dijumpai , 1 : Pembeli Tidak Berdaftar , 2 : Pembeli Berdaftar , 3 : Ahli
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 44_senarai_pelanggan where no_resit='" & Frm93.L18_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!Nama) Then 'Nama
                    Frm26.TB1 = rs!Nama
                    Frm93.L35_Text = rs!Nama
                Else
                    Frm26.TB1 = vbNullString
                End If
                If Not IsNull(rs!no_tel) Then 'No. Telefon
                    Frm26.TB2 = rs!no_tel
                Else
                    Frm26.TB2 = vbNullString
                End If
            End If
            
            rs.Close
            Set rs = Nothing
        End If
        '### Maklumat Pembeli Yang Tidak Berdaftar Dengan Kedai ### - End
                
        '###Update Data Simpanan Duit Pelanggan### - Start
        'If Frm93_LM_KATEGORI <> 1 And Frm93_LM_No_PEMBELI <> vbNullString Then
        If Frm93_LM_No_PEMBELI <> vbNullString Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm93_LM_No_PEMBELI & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                Call Frm28_initial
                
                If Not IsNull(rs!Nama) Then
                    Frm28.L1_Text = rs!Nama 'Nama
                    Frm93.L36_Text = rs!Nama
                End If
                If Not IsNull(rs!no_ic) Then Frm28.L2_Text = rs!no_ic 'No. Kad Pengenalan
                If Not IsNull(rs!no_tel) Then Frm28.L3_Text = rs!no_tel 'No. Telefon
                If Not IsNull(rs!Email) Then Frm28.L4_Text = rs!Email 'E-mail
                If Not IsNull(rs!no_pelanggan) Then Frm28.L5_Text = rs!no_pelanggan 'No. Pelanggan
                If Not IsNull(rs!baki_simpanan) Then
                    frm130.L26_Text = Format(rs!baki_simpanan, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
                    If IsNumeric(rs!baki_simpanan) Then
                        Frm93_LM_SIMPANAN_ASAL = rs!baki_simpanan 'Jumlah Simpanan Asal Yang Ada (RM)
                        
                        frm130.L26_Text = Format(Frm93_LM_SIMPANAN_ASAL + Frm93_LM_SIMPANAN_DIGUNAKAN, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
                    End If
                End If
            End If
            
            rs.Close
            Set rs = Nothing
            
        End If
                
'###Update Data Simpanan Duit Pelanggan### - End
   
'###Update Data Simpanan Duit Pelanggan### - Start
        'If Frm93_LM_No_PEMBELI <> vbNullString And Frm93_LM_Flag_SIMPANAN = 1 Then
        '    Set rs = New ADODB.Recordset
        '    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        '    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm93_LM_No_PEMBELI & "'", cn, adOpenKeyset, adLockOptimistic
            
        '    If Not rs.EOF Then
        '        If Not IsNull(rs!baki_simpanan) Then
        '            If IsNumeric(rs!baki_simpanan) Then
        '                Frm93_LM_SIMPANAN_ASAL = rs!baki_simpanan 'Jumlah Simpanan Asal Yang Ada (RM)
                        
        '                Frm130.L26_Text = Format(Frm93_LM_SIMPANAN_ASAL + Frm93_LM_SIMPANAN_DIGUNAKAN, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
        '            End If
        '        End If
        '    End If
            
        '    rs.Close
        '    Set rs = Nothing
        'End If
'###Update Data Simpanan Duit Pelanggan### - End

        If DATA_FOUND = 1 Then
        
            Frm93.CBB3.Enabled = True
            Frm93.CBB3.BackColor = &HFFFFFF
            
            Frm93.CMD12.Visible = False
            Frm93.CMD14.Visible = True
            Frm93.CMD15.Visible = True
            
            Frm93.CB9.Enabled = False
            
            Frm93.Frame2.Visible = True
            Frm93.Frame1.Visible = False
        End If
    End If
    
End If

Exit Sub
Err_A:
Frm93.CBB1.AddItem Frm93_LM_KATEGORI_PRODUK
Frm93.CBB1 = Frm93_LM_KATEGORI_PRODUK
Resume Restore_A:

Exit Sub
Err_B:
Frm93.CBB2.AddItem Frm93_LM_PURITY
Frm93.CBB2 = Frm93_LM_PURITY
Resume Restore_B:

Exit Sub
Err_C:
Frm93.CBB3.AddItem Frm93_LM_MAKLUMAT_PEKERJA
Frm93.CBB3 = Frm93_LM_MAKLUMAT_PEKERJA
Resume Restore_C:

Exit Sub
Err_D:
frm130.CBB2.AddItem Frm93_LM_JENIS_KAD
frm130.CBB2 = Frm93_LM_JENIS_KAD
Resume Restore_D:
End Sub
Private Sub Frm93_SM_harga_semasa_Click()
'on error resume next
DATA_FOUND = 0
Frm93_LM_Search_Price = 0
Frm93_LM_KATEGORI_PEMBELI = 1
Frm93_LM_KOD_PURITY = vbNullString
Frm93_LM_No_PELANGGAN = vbNullString
Frm93_LM_STATUS = vbNullString

Frm93_LM_ID = vbNullString

If IsNumeric(Frm93.LV1.SelectedItem.Index) Then
    
    Frm93_LM_ID = Frm93.LV1.ListItems(Frm93.LV1.SelectedItem.Index)
    
    If Frm93_LM_ID <> vbNullString Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 40_tempahan_deposit where ID='" & Frm93_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            If Not IsNull(rs!status_invoice) Then
                
                If rs!status_invoice = 0 Then
                    
                    MsgBox "Status deposit tempahan / data tempahan ini telah bertukar status kepada TIDAK AKTIF/TELAH DIPADAMKAN." & vbCrLf & _
                            vbNullString & vbCrLf & _
                            "Oleh itu anda tidak dibenarkan untuk teruskan menu ini.", vbExclamation, "Info"
                            
                    rs.Close
                    Set rs = Nothing
                    
                    Exit Sub
                    
                End If
            
            End If
            
            If Not IsNull(rs!Status) Then
                
                If rs!Status = "Siap" Then
                    
                    MsgBox "Status deposit tempahan / data tempahan ini telah bertukar status kepada SIAP." & vbCrLf & _
                            vbNullString & vbCrLf & _
                            "Oleh itu anda tidak dibenarkan untuk teruskan menu ini.", vbExclamation, "Info"
                            
                    rs.Close
                    Set rs = Nothing
                    
                    Exit Sub
                    
                End If
            
            End If
            
        End If
        
        rs.Close
        Set rs = Nothing
        
        Call frm130_initial_setting
        Call Frm94_initial_setting
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 40_tempahan_deposit where ID='" & Frm93_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!no_rujukan_pelanggan) Then Frm93_LM_No_PELANGGAN = rs!no_rujukan_pelanggan 'No. Rujukan Pelanggan
            If Not IsNull(rs!cawangan) Then G_KEDAI = rs!cawangan
            
            If Not IsNull(rs!kategori_pembeli) Then
                Frm93_LM_KATEGORI_PEMBELI = rs!kategori_pembeli
                Frm94.L14_Text = Frm93_LM_KATEGORI_PEMBELI
            End If
            
            If Not IsNull(rs!jenis_tempahan) Then
                If rs!jenis_tempahan = 0 Then 'Jenis Tempahan , 0 : Tempahan Barang Baru , 1 : Tempahan Barang Kedai
                    
                    Frm94_LM_JENIS_TEMPAHAN = 0 'Jenis Tempahan , 0 : Tempahan Barang Baru , 1 : Tempahan Barang Kedai
                
                    Frm94.TB1.Locked = False
                    Frm94.TB1.BackColor = &HFFFFFF
                    Frm94.CMD1.Enabled = True
                    
                    If Not IsNull(rs!type_barang_kemas) Then
                        If rs!type_barang_kemas = 0 Then 'Jenis Barang , 0 : Barang Kemas , 1 : Barang Permata
                            Frm93_LM_Search_Price = 1
                            If Not IsNull(rs!purity) Then Frm93_LM_KOD_PURITY = rs!purity 'Purity
                            
                            Frm94.L15_Text = 0 'Jenis Barang Tempahan , 0 : Barang Kemas , 1 : Barang Permata
                            
                            If Not IsNull(rs!harga_Semasa) Then 'Harga Semasa
                                Frm94.TB5 = rs!harga_Semasa
                            Else
                                Frm94.TB5 = vbNullString
                            End If
                            
                            Frm94.TB4.Locked = False
                            Frm94.TB5.Locked = False
                            Frm94.TB6.Locked = False
                            Frm94.TB7.Locked = True
                            
                            Frm94.TB4.BackColor = &HFFFFFF
                            Frm94.TB5.BackColor = &HFFFFFF
                            Frm94.TB6.BackColor = &HFFFFFF
                            Frm94.TB7.BackColor = &H8000000A
                        Else
                            Frm94.L15_Text = 1 'Jenis Barang Tempahan , 0 : Barang Kemas , 1 : Barang Permata
                            
                            Frm94.TB3 = vbNullString
                            
                            Frm94.TB4 = vbNullString
                            Frm94.TB5 = vbNullString
                            'Frm94.TB6 = vbNullString
                            
                            Frm94.TB4.Locked = True
                            Frm94.TB5.Locked = True
                            Frm94.TB6.Locked = False
                            Frm94.TB7.Locked = False
                            
                            Frm94.TB4.BackColor = &H8000000A
                            Frm94.TB5.BackColor = &H8000000A
                            Frm94.TB6.BackColor = &HFFFFFF
                            Frm94.TB7.BackColor = &HFFFFFF
                        End If
                    End If
                    
                    If Not IsNull(rs!UPAH) Then 'Upah
                        Frm94.TB6 = rs!UPAH
                    Else
                        Frm94.TB6 = vbNullString
                    End If
                    If Not IsNull(rs!anggaran_harga) Then 'Anggaran Harga
                        Frm94.TB7 = rs!anggaran_harga
                        Frm94.TB9 = rs!anggaran_harga
                    Else
                        Frm94.TB7 = "0.00"
                        Frm94.TB9 = "0.00"
                    End If
                    If Not IsNull(rs!jumlah_tanpa_gst) Then 'Jumlah Deposit
                        Frm94.TB11 = rs!jumlah_tanpa_gst
                    Else
                        Frm94.TB11 = vbNullString
                    End If

                    Frm94.L7_Text = 0 'Jenis Tempahan , 0 : Tempahan Barang Baru , 1 : Tempahan Barang Kedai
                    
                ElseIf rs!jenis_tempahan = 1 Then 'Jenis Tempahan , 0 : Tempahan Barang Baru , 1 : Tempahan Barang Kedai
                    Frm94.TB1.Locked = True
                    Frm94.TB1.BackColor = &H8000000A
                    Frm94.CMD1.Enabled = False
                    
                    Frm94_LM_JENIS_TEMPAHAN = 1 'Jenis Tempahan , 0 : Tempahan Barang Baru , 1 : Tempahan Barang Kedai
                    
                    Frm94.L7_Text = 1 'Jenis Tempahan , 0 : Tempahan Barang Baru , 1 : Tempahan Barang Kedai
                    
                    If Not IsNull(rs!no_rujukan_tempahan) Then 'No. Rujukan Tempahan
                        Frm94.L9_Text = rs!no_rujukan_tempahan
                    Else
                        Frm94.L9_Text = 1
                    End If
                    
                    If Not IsNull(rs!no_siri_Produk) Then 'No. Siri Produk
                        Frm94.TB2 = rs!no_siri_Produk
                    Else
                        Frm94.TB2 = vbNullString
                    End If
                    If Not IsNull(rs!kategori_Produk) Then 'Kategori Produk
                        Frm94.L3_Text = rs!kategori_Produk
                    Else
                        Frm94.L3_Text = vbNullString
                    End If

                    If Not IsNull(rs!type_barang_kemas) Then
                        If rs!type_barang_kemas = 0 Then 'Jenis Barang , 0 : Barang Kemas , 1 : Barang Permata
                            If Not IsNull(rs!Berat_Asal) Then 'Berat Asal
                                Frm94.TB3 = rs!Berat_Asal
                            Else
                                Frm94.TB3 = vbNullString
                            End If
                            If Not IsNull(rs!berat_jualan) Then 'Berat Jualan
                                Frm94.TB4 = rs!berat_jualan
                            Else
                                Frm94.TB4 = vbNullString
                            End If
                            If Not IsNull(rs!harga_Semasa) Then 'Harga Semasa
                                Frm94.TB5 = rs!harga_Semasa
                            Else
                                Frm94.TB5 = vbNullString
                            End If
                            
                            Frm94.L15_Text = 0 'Jenis Barang Tempahan , 0 : Barang Kemas , 1 : Barang Permata
                            
                            If Not IsNull(rs!no_siri_Produk) Then Frm93_LM_No_SIRI = rs!no_siri_Produk 'No. Siri Produk
                            Frm93_LM_Search_Price = 1
                            
                            Frm94.TB4.Locked = False
                            Frm94.TB5.Locked = False
                            Frm94.TB6.Locked = False
                            Frm94.TB7.Locked = True
                            
                            Frm94.TB4.BackColor = &HFFFFFF
                            Frm94.TB5.BackColor = &HFFFFFF
                            Frm94.TB6.BackColor = &HFFFFFF
                            Frm94.TB7.BackColor = &H8000000A
                        Else
                        
                            Frm94.L15_Text = 1 'Jenis Barang Tempahan , 0 : Barang Kemas , 1 : Barang Permata
                            
                            Frm94.TB3 = vbNullString
                            
                            Frm94.TB4 = vbNullString
                            Frm94.TB5 = vbNullString
                            'Frm94.TB6 = vbNullString
                            
                            Frm94.TB4.Locked = True
                            Frm94.TB5.Locked = True
                            Frm94.TB6.Locked = False
                            Frm94.TB7.Locked = False
                            
                            Frm94.TB4.BackColor = &H8000000A
                            Frm94.TB5.BackColor = &H8000000A
                            Frm94.TB6.BackColor = &HFFFFFF
                            Frm94.TB7.BackColor = &HFFFFFF
                        End If
                    End If
                    If Not IsNull(rs!UPAH) Then 'Upah
                        Frm94.TB6 = rs!UPAH
                    Else
                        Frm94.TB6 = vbNullString
                    End If
                    If Not IsNull(rs!harga_asal) Then 'Harga Asal
                        Frm94.TB7 = rs!harga_asal
                    Else
                        Frm94.TB7 = vbNullString
                    End If
                    If Not IsNull(rs!adjustment) Then 'Adjustment
                        Frm94.TB8 = rs!adjustment
                    Else
                        Frm94.TB8 = vbNullString
                    End If
                    If Not IsNull(rs!anggaran_harga) Then 'Anggaran Harga
                        Frm94.TB9 = rs!anggaran_harga
                    Else
                        Frm94.TB9 = vbNullString
                    End If
                    If Not IsNull(rs!jumlah_tanpa_gst) Then 'Jumlah Deposit
                        Frm94.TB11 = rs!jumlah_tanpa_gst
                    Else
                        Frm94.TB11 = vbNullString
                    End If
                End If
                
                If Not IsNull(rs!no_rujukan_tempahan) Then 'No. Rujukan Tempahan
                    Frm94.L9_Text = rs!no_rujukan_tempahan
                Else
                    Frm94.L9_Text = 1
                End If
                
                Frm94.L5_Text = 1 '0 : Harga Emas Ikut Harga Tempahan , 1 : Harga Emas Ikut Harga Semasa
            End If
            
            DATA_FOUND = 1
        End If
        
        rs.Close
        Set rs = Nothing
        
        If Frm93_LM_No_PELANGGAN <> vbNullString Then
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm93_LM_No_PELANGGAN & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!baki_simpanan) Then frm130.L26_Text = Format(rs!baki_simpanan, "#,##0.00") 'Baki Simpanan Pelanggan Ini (RM)
            End If
            
            rs.Close
            Set rs = Nothing
        End If
                
        If DATA_FOUND = 1 Then
        
'### Periksa Harga Semasa Emas ### - Start
            If Frm93_LM_Search_Price = 1 Then
                
                '###Carian Purity Bagi Item Ini### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from Data_Database where no_siri_Produk='" & Frm93_LM_No_SIRI & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Not IsNull(rs!kod_Purity) Then Frm93_LM_KOD_PURITY = rs!kod_Purity 'Kod Purity
                End If
                
                rs.Close
                Set rs = Nothing
            
        
                '###Periksa Data Produk### - Start
                If Frm93_LM_KOD_PURITY <> vbNullString Then
                    
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from hargaemas where Purity='" & Frm93_LM_KOD_PURITY & "' AND cawangan='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                        If Frm93_LM_KATEGORI_PEMBELI = 1 Or Frm93_LM_KATEGORI_PEMBELI = 0 Then
                            If IsNumeric(rs!Harga_Pelanggan) Then Frm94.TB5 = Format(rs!Harga_Pelanggan, "0.00") 'Harga Semasa Bagi Pelanggan (RM/g)
                        ElseIf Frm93_LM_KATEGORI_PEMBELI = 2 Then
                            If IsNumeric(rs!Harga_Member) Then Frm94.TB5 = Format(rs!Harga_Member, "0.00") 'Harga Semasa Bagi Member (RM/g)
                        ElseIf Frm93_LM_KATEGORI_PEMBELI = 3 Then
                            If IsNumeric(rs!Harga_Pengedar) Then Frm94.TB5 = Format(rs!Harga_Pengedar, "0.00") 'Harga Semasa Bagi Pengedar (RM/g)
                        ElseIf Frm93_LM_KATEGORI_PEMBELI = 4 Then
                            If IsNumeric(rs!Harga_RAF) Then Frm94.TB5 = Format(rs!Harga_RAF, "0.00") 'Harga Semasa Bagi RAF (RM/g)
                        ElseIf Frm93_LM_KATEGORI_PEMBELI = 5 Then
                            If IsNumeric(rs!harga_nd) Then Frm94.TB5 = Format(rs!harga_nd, "0.00") 'Harga Semasa Bagi Normal Dealer (RM/g)
                        ElseIf Frm93_LM_KATEGORI_PEMBELI = 6 Then
                            If IsNumeric(rs!harga_md) Then Frm94.TB5 = Format(rs!harga_md, "0.00") 'Harga Semasa Bagi Master Dealer (RM/g)
                        End If
                    End If
                    
                    rs.Close
                    Set rs = Nothing
                End If
            End If
'### Periksa Harga Semasa Emas ### - End
                    
            Call Frm94_jurujual
            
            Frm94.Show
            Frm93.Hide
            
            MDI_frm1.L5_Text = 9
            
            Frm94.TB1.SetFocus
        End If

    End If
End If
End Sub
Private Sub Frm93_SM_harga_tempahan_Click()
'on error resume next
DATA_FOUND = 0
Frm93_LM_No_PELANGGAN = vbNullString
Frm93_LM_STATUS = vbNullString

Frm93_LM_ID = vbNullString

If IsNumeric(Frm93.LV1.SelectedItem.Index) Then
    
    Frm93_LM_ID = Frm93.LV1.ListItems(Frm93.LV1.SelectedItem.Index)
    
    If Frm93_LM_ID <> vbNullString Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 40_tempahan_deposit where ID='" & Frm93_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            If Not IsNull(rs!status_invoice) Then
                
                If rs!status_invoice = 0 Then
                    
                    MsgBox "Status deposit tempahan / data tempahan ini telah bertukar status kepada TIDAK AKTIF/TELAH DIPADAMKAN." & vbCrLf & _
                            vbNullString & vbCrLf & _
                            "Oleh itu anda tidak dibenarkan untuk teruskan menu ini.", vbExclamation, "Info"
                            
                    rs.Close
                    Set rs = Nothing
                    
                    Exit Sub
                    
                End If
            
            End If
            
            If Not IsNull(rs!Status) Then
                
                If rs!Status = "Siap" Then
                    
                    MsgBox "Status deposit tempahan / data tempahan ini telah bertukar status kepada SIAP." & vbCrLf & _
                            vbNullString & vbCrLf & _
                            "Oleh itu anda tidak dibenarkan untuk teruskan menu ini.", vbExclamation, "Info"
                            
                    rs.Close
                    Set rs = Nothing
                    
                    Exit Sub
                    
                End If
            
            End If
            
        End If
        
        rs.Close
        Set rs = Nothing
        
        Call frm130_initial_setting
        Call Frm94_initial_setting
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 40_tempahan_deposit where ID='" & Frm93_LM_ID & "' AND status_invoice = 1", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!no_rujukan_pelanggan) Then Frm93_LM_No_PELANGGAN = rs!no_rujukan_pelanggan 'No. Rujukan Pelanggan
            If Not IsNull(rs!cawangan) Then G_KEDAI = rs!cawangan
            
            If Not IsNull(rs!kategori_pembeli) Then
                Frm93_LM_KATEGORI_PEMBELI = rs!kategori_pembeli
                Frm94.L14_Text = Frm93_LM_KATEGORI_PEMBELI
            End If
            
            If Not IsNull(rs!jenis_tempahan) Then
                If rs!jenis_tempahan = 0 Then 'Jenis Tempahan , 0 : Tempahan Barang Baru , 1 : Tempahan Barang Kedai
                    Frm94.TB1.Locked = False
                    Frm94.TB1.BackColor = &HFFFFFF
                    Frm94.CMD1.Enabled = True
                    
                    If Not IsNull(rs!type_barang_kemas) Then
                        If rs!type_barang_kemas = 0 Then 'Jenis Barang , 0 : Barang Kemas , 1 : Barang Permata
                            
                            Frm94.L15_Text = 0 'Jenis Barang Tempahan , 0 : Barang Kemas , 1 : Barang Permata
                            
                            If Not IsNull(rs!harga_Semasa) Then 'Harga Semasa
                                Frm94.TB5 = rs!harga_Semasa
                            Else
                                Frm94.TB5 = vbNullString
                            End If
                            
                            Frm94.TB4.Locked = False
                            Frm94.TB5.Locked = False
                            Frm94.TB6.Locked = False
                            Frm94.TB7.Locked = True
                            
                            Frm94.TB4.BackColor = &HFFFFFF
                            Frm94.TB5.BackColor = &HFFFFFF
                            Frm94.TB6.BackColor = &HFFFFFF
                            Frm94.TB7.BackColor = &H8000000A
                        Else
                            
                            Frm94.L15_Text = 1 'Jenis Barang Tempahan , 0 : Barang Kemas , 1 : Barang Permata
                            
                            Frm94.TB3 = vbNullString
                            
                            Frm94.TB4 = vbNullString
                            Frm94.TB5 = vbNullString
                            'Frm94.TB6 = vbNullString
                            
                            Frm94.TB4.Locked = True
                            Frm94.TB5.Locked = True
                            Frm94.TB6.Locked = False
                            Frm94.TB7.Locked = False
                                    
                            Frm94.TB4.BackColor = &H8000000A
                            Frm94.TB5.BackColor = &H8000000A
                            Frm94.TB6.BackColor = &HFFFFFF
                            Frm94.TB7.BackColor = &HFFFFFF
                        End If
                    End If
                    
                    If Not IsNull(rs!UPAH) Then 'Upah
                        Frm94.TB6 = rs!UPAH
                    Else
                        Frm94.TB6 = vbNullString
                    End If
                    If Not IsNull(rs!anggaran_harga) Then 'Anggaran Harga
                        Frm94.TB7 = rs!anggaran_harga
                        Frm94.TB9 = rs!anggaran_harga
                    Else
                        Frm94.TB7 = "0.00"
                        Frm94.TB9 = "0.00"
                    End If
                    If Not IsNull(rs!jumlah_tanpa_gst) Then 'Jumlah Deposit
                        Frm94.TB11 = rs!jumlah_tanpa_gst
                    Else
                        Frm94.TB11 = vbNullString
                    End If

                    Frm94.L7_Text = 0 'Jenis Tempahan , 0 : Tempahan Barang Baru , 1 : Tempahan Barang Kedai
                    
                ElseIf rs!jenis_tempahan = 1 Then 'Jenis Tempahan , 0 : Tempahan Barang Baru , 1 : Tempahan Barang Kedai
                    Frm94.TB1.Locked = True
                    Frm94.TB1.BackColor = &H8000000A
                    Frm94.CMD1.Enabled = False
                    
                    Frm94.L7_Text = 1 'Jenis Tempahan , 0 : Tempahan Barang Baru , 1 : Tempahan Barang Kedai
                    
                    If Not IsNull(rs!no_rujukan_tempahan) Then 'No. Rujukan Tempahan
                        Frm94.L9_Text = rs!no_rujukan_tempahan
                    Else
                        Frm94.L9_Text = 1
                    End If
                    
                    If Not IsNull(rs!no_siri_Produk) Then 'No. Siri Produk
                        Frm94.TB2 = rs!no_siri_Produk
                    Else
                        Frm94.TB2 = vbNullString
                    End If
                    If Not IsNull(rs!kategori_Produk) Then 'Kategori Produk
                        Frm94.L3_Text = rs!kategori_Produk
                    Else
                        Frm94.L3_Text = vbNullString
                    End If
        
                    If Not IsNull(rs!type_barang_kemas) Then
                        If rs!type_barang_kemas = 0 Then 'Jenis Barang , 0 : Barang Kemas , 1 : Barang Permata
                            Frm94.L15_Text = 0 'Jenis Barang Tempahan , 0 : Barang Kemas , 1 : Barang Permata
                            
                            If Not IsNull(rs!Berat_Asal) Then 'Berat Asal
                                Frm94.TB3 = rs!Berat_Asal
                            Else
                                Frm94.TB3 = vbNullString
                            End If
                            If Not IsNull(rs!berat_jualan) Then 'Berat Jualan
                                Frm94.TB4 = rs!berat_jualan
                            Else
                                Frm94.TB4 = vbNullString
                            End If
                            If Not IsNull(rs!harga_Semasa) Then 'Harga Semasa
                                Frm94.TB5 = rs!harga_Semasa
                            Else
                                Frm94.TB5 = vbNullString
                            End If
                            
                            Frm94.TB4.Locked = False
                            Frm94.TB5.Locked = False
                            Frm94.TB6.Locked = False
                            Frm94.TB7.Locked = True
                            
                            Frm94.TB4.BackColor = &HFFFFFF
                            Frm94.TB5.BackColor = &HFFFFFF
                            Frm94.TB6.BackColor = &HFFFFFF
                            Frm94.TB7.BackColor = &H8000000A
                        Else
                            Frm94.L15_Text = 1 'Jenis Barang Tempahan , 0 : Barang Kemas , 1 : Barang Permata
                            
                            Frm94.TB3 = vbNullString
                            
                            Frm94.TB4 = vbNullString
                            Frm94.TB5 = vbNullString
                            'Frm94.TB6 = vbNullString
                            
                            Frm94.TB4.Locked = True
                            Frm94.TB5.Locked = True
                            Frm94.TB6.Locked = False
                            Frm94.TB7.Locked = False
                            
                            Frm94.TB4.BackColor = &H8000000A
                            Frm94.TB5.BackColor = &H8000000A
                            Frm94.TB6.BackColor = &HFFFFFF
                            Frm94.TB7.BackColor = &HFFFFFF
                        End If
                    End If
                    If Not IsNull(rs!UPAH) Then 'Upah
                        Frm94.TB6 = rs!UPAH
                    Else
                        Frm94.TB6 = vbNullString
                    End If
                    If Not IsNull(rs!harga_asal) Then 'Harga Asal
                        Frm94.TB7 = rs!harga_asal
                    Else
                        Frm94.TB7 = vbNullString
                    End If
                    If Not IsNull(rs!adjustment) Then 'Adjustment
                        Frm94.TB8 = rs!adjustment
                    Else
                        Frm94.TB8 = vbNullString
                    End If
                    If Not IsNull(rs!anggaran_harga) Then 'Anggaran Harga
                        Frm94.TB9 = rs!anggaran_harga
                    Else
                        Frm94.TB9 = vbNullString
                    End If
                    If Not IsNull(rs!jumlah_tanpa_gst) Then 'Jumlah Deposit
                        Frm94.TB11 = rs!jumlah_tanpa_gst
                    Else
                        Frm94.TB11 = vbNullString
                    End If
                End If
                        
                If Not IsNull(rs!no_rujukan_tempahan) Then 'No. Rujukan Tempahan
                    Frm94.L9_Text = rs!no_rujukan_tempahan
                Else
                    Frm94.L9_Text = 1
                End If
                
                Frm94.L5_Text = 0 '0 : Harga Emas Ikut Harga Tempahan , 1 : Harga Emas Ikut Harga Semasa
            End If
            
            DATA_FOUND = 1
        End If
        
        rs.Close
        Set rs = Nothing
        
        If Frm93_LM_No_PELANGGAN <> vbNullString Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm93_LM_No_PELANGGAN & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!baki_simpanan) Then frm130.L26_Text = Format(rs!baki_simpanan, "#,##0.00") 'Baki Simpanan Pelanggan Ini (RM)
            End If
            
            rs.Close
            Set rs = Nothing
            
        End If
                
        If DATA_FOUND = 1 Then
            
            If MDI_frm1.L20_Text <> G_KEDAI Then
            
                MDI_frm1.L20_Text = G_KEDAI
                
                Call main_setting_kedai
                Call main_setting
                
                MsgBox "Branch telah ditukar kepada " & G_KEDAI & " bagi meneruskan urusan ini.", vbInformation, "Info"
                
            End If
        
            Call Frm94_jurujual
        
            Frm94.Show
            Frm93.Hide
            
            MDI_frm1.L5_Text = 9
            
            Frm94.TB1.SetFocus
        End If

    End If
End If
End Sub
Private Sub Frm93_SM_padam_Click()
'on error resume next
Dim Frm93_LM_SIMPANAN_DIGUNAKAN As Double
Dim Frm93_LM_SIMPANAN_ASAL As Double

Frm93_LM_SIMPANAN_DIGUNAKAN = 0
Frm93_LM_SIMPANAN_ASAL = 0
Frm93_LM_FLAG_TI = 0
Frm93_LM_FLAG_BARANG_KEDAI = 0 'Flag Barang Kedai

frm93_LM_No_ID = vbNullString

If IsNumeric(Frm93.LV1.SelectedItem.Index) Then
    
    frm93_LM_No_ID = Frm93.LV1.ListItems(Frm93.LV1.SelectedItem.Index)
    
    If frm93_LM_No_ID <> vbNullString Then
        
        Note = "Adakah anda ingin padam data tempahan ini?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
        
            G_JENIS_URUSAN = 11
            
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
            
            LM_STATUS = 0
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 40_tempahan_deposit where ID='" & Frm93_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!status_invoice) Then
                    
                    If rs!status_invoice = 0 Then
                        
                        MsgBox "Status deposit tempahan / data tempahan ini telah bertukar status kepada TIDAK AKTIF/TELAH DIPADAMKAN." & vbCrLf & _
                                vbNullString & vbCrLf & _
                                "Oleh itu anda tidak dibenarkan untuk padam/batal data ini.", vbExclamation, "Info"
                                
                        rs.Close
                        Set rs = Nothing
                        
                        Exit Sub
                        
                    End If
                
                End If
                
            End If
            
            rs.Close
            Set rs = Nothing
                
'### Padam Data Dari Senarai Tempahan (Deposit) ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 40_tempahan_deposit where ID='" & frm93_LM_No_ID & "' AND status_invoice = 1", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
            
                G_ID = rs!ID
                Call recovery_40_tempahan_deposit
                
                If Not IsNull(rs!Status) Then
                    If rs!Status = "Siap" Then
                        LM_STATUS = 1
                    End If
                End If
                If Not IsNull(rs!invoice_siap) Then G_No_RESIT_JUALAN = rs!invoice_siap
                If Not IsNull(rs!no_resit_tempahan) Then Frm93_LM_No_RESIT = rs!no_resit_tempahan 'No. Resit
                If Not IsNull(rs!flag_trade_in) Then
                    If rs!flag_trade_in = 1 Then
                        Frm93_LM_FLAG_TI = 1
                        If Not IsNull(rs!no_resit_trade_in) Then Frm93_LM_No_RESIT_TI = rs!no_resit_trade_in 'No. Resit Trade In
                    End If
                End If
                
                If Not IsNull(rs!jenis_tempahan) Then
                    If rs!jenis_tempahan = 1 Then
                        Frm93_LM_FLAG_BARANG_KEDAI = 1 'Flag Barang Kedai
                        If Not IsNull(rs!no_siri_Produk) Then Frm93_LM_No_SIRI = rs!no_siri_Produk 'No. Siri Produk
                    End If
                End If
                
                If Not IsNull(rs!no_rujukan_pelanggan) Then Frm93_LM_No_ID_PEMBELI = rs!no_rujukan_pelanggan 'No. Rujukan Pembeli
                
                rs!status_invoice = 0
                
                rs!terminal = G_TERMINAL
                LM_NOW = Now
                rs!write_timestamp2 = LM_NOW
                rs!no_staff = G_LOGIN_USER 'No. staff yang memasukkan/edit data (Login)
                rs.Update
                
            End If
            
            rs.Close
            Set rs = Nothing
'### Padam Data Dari Senarai Tempahan (Deposit) ### - End
    
'### Pulangkan Status Barang Trade In ### - Start
            If Frm93_LM_FLAG_TI = 1 Then '0 : Tiada Urusan Trade In , 1 : Ada Urusan Trade In
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & Frm93_LM_No_RESIT_TI & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                
                    G_ID = rs!ID
                    Call recovery_16_gold_bar_belian
    
                    rs!trade_in_status = 0
                    
                    rs!no_staff = G_LOGIN_USER 'No. Pekerja
                    rs!terminal = G_TERMINAL
                    rs!write_timestamp2 = LM_NOW
                    rs!jenis_urusan = G_JENIS_URUSAN
                    rs!remarks = "Kembalikan status trade in - padam data tempahan"
                    rs.Update

                End If
                
                rs.Close
                Set rs = Nothing
                
            End If
'### Pulangkan Status Barang Trade In ### - End
    
'### Pulangkan Status Item Dalam Database ### - Start
            If Frm93_LM_FLAG_BARANG_KEDAI = 1 Then 'Flag Barang Kedai
            
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from Data_Database where no_siri_produk='" & Frm93_LM_No_SIRI & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                
                    G_ID = rs!ID
                    Call recovery_data_database
    
                    rs!StatusItem = 10
                    
                    rs!write_timestamp2 = LM_NOW
                    rs!no_pekerja = G_LOGIN_USER
                    rs!terminal = G_TERMINAL
                    rs!Menu = 4
    
                    rs.Update
                End If
                
                rs.Close
                Set rs = Nothing
                
            End If
'### Pulangkan Status Item Dalam Database ### - End

'###Padam Akaun Tempahan### - Start
            Frm93_LM_FLAG_SAVING = 0
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 22_jualan where no_resit='" & Frm93_LM_No_RESIT & "' AND Status = 1", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                G_ID = rs!ID
                Call recovery_22_jualan
                
                If Not IsNull(rs!duit_simpanan_kedai) Then
                    If Format(rs!duit_simpanan_kedai, "0.00") <> "0.00" Then
                        If IsNumeric(rs!duit_simpanan_kedai) Then Frm93_LM_SIMPANAN_DIGUNAKAN = rs!duit_simpanan_kedai
                        Frm93_LM_FLAG_SAVING = 1
                    End If
                End If
                
                rs!Status = 0
                rs!terminal = G_TERMINAL
                rs!no_staff = G_LOGIN_USER
                rs!write_timestamp2 = LM_NOW
                rs!Menu = 4
                rs.Update
                
            End If
            
            rs.Close
            Set rs = Nothing
'###Padam Akaun Tempahan### - End

'###Update Simpanan Duit Di Kedai### - Start
            If Frm93_LM_FLAG_SAVING = 1 Then
            
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm93_LM_No_ID_PEMBELI & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                
                    G_ID = rs!ID
                    Call recovery_senarai_pelanggan
    
                    If Not IsNull(rs!baki_simpanan) Then
                        If IsNumeric(rs!baki_simpanan) Then Frm93_LM_SIMPANAN_ASAL = rs!baki_simpanan 'Baki Simpanan Pelanggan Ini (RM)
                    End If
                    
                    rs!baki_simpanan = Format(Frm93_LM_SIMPANAN_ASAL + Frm93_LM_SIMPANAN_DIGUNAKAN, "0.00") 'Baki Simpanan Terkini Pelanggan Ini (RM)
                    
                    rs!write_timestamp2 = LM_NOW
                    rs!no_staff = G_LOGIN_USER 'No. Pekerja
                    rs!terminal = G_TERMINAL
                    rs!jenis_urusan = G_JENIS_URUSAN
    
                    rs.Update
                End If
                
                rs.Close
                Set rs = Nothing
                    
'###Padam Rekod Bayaran Dalam Table Simpanan### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 24_rekod_kewangan_pelanggan where NO_RESIT='" & Frm93_LM_No_RESIT & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    
                    G_ID = rs!ID
                    Call recovery_24_rekod_kewangan_pelanggan
    
                    rs.Delete
                    rs.Update
                End If
                
                rs.Close
                Set rs = Nothing
'###Padam Rekod Bayaran Dalam Table Simpanan### - End
                    
            End If
'###Update Simpanan Duit Di Kedai### - End

            '### Padam Rekod Belian Pembeli Yang Tidak Berdaftar Dengan Kedai ### - Start (08-07-2015)
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 44_senarai_pelanggan where no_resit='" & Frm93_LM_No_RESIT & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                G_ID = rs!ID
                Call recovery_44_senarai_pelanggan

                rs.Delete
                rs.Update
            
            End If
            
            rs.Close
            Set rs = Nothing
            '### Padam Rekod Belian Pembeli Yang Tidak Berdaftar Dengan Kedai ### - End (08-07-2015)

            If LM_STATUS = 1 Then
                
                G_TEMPAHAN = 0 '0 : Padam data , 1 : Tukar status kepada belum siap
                Call Frm93_padam_data_tempahan

            End If
'### Update Log ### - Start
            'User = MDI_frm1.L3_Text
            LogAct_Memory = "[" & G_LOGIN_USER & "] Padam data tempahan. No. ID tempahan [" & frm93_LM_No_ID & "]"
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
'### Update Log ### - End

            GM_NEXT_PREV = 2
                
            Call frm93_tempahan_header
            Call frm93_tempahan
                
            MsgBox "Data tempahan telah berjaya dipadamkan", vbInformation, "Info"
            
        End If
    
    
    End If
    
End If
End Sub
Private Sub L10_Text_Click()
'on error resume next
If Frm93.Frame8.Visible = False Then

    Call frm93_initial_setting2
    
    Frm93.Frame8.Visible = True
    
Else

    Frm93.Frame8.Visible = False
    
End If
End Sub
Private Sub L12_Text_Click()
'on error resume next
Frm15.Show
Unload Frm93
Unload Frm26
Unload Frm27
Unload Frm28
End Sub



Private Sub L19_Text_Change()
'on error resume next
Call Frm93_kira_caj_gst_kad_kredit
End Sub


Private Sub L3_Text_Click()
'on error resume next
If Frm93.Pic4.Visible = False Then

    If MDI_frm1.L20_Text = "Semua cawangan" Then
    
        Frm96.CMD2.Visible = True
        Frm96.CMD1.Visible = False
    
        Call Frm96_initial
        
        Frm96.Show vbModal
        
    End If
    
    Call Frm93_initial_setting
    Call frm93_initial_setting2
    Call Frm93_jurujual
    
    'Frm93.Pic6.Left = 11280
    'Frm93.Pic6.Top = 7200
    
    Unload Frm26
    Unload Frm27
    Unload Frm28
    
    Frm93.Pic4.Visible = True
    
    Frm93.CB4 = 1
Else
    Frm93.Pic4.Visible = False
End If
End Sub

Private Sub L41_Text_Change()
'On Error Resume Next
Call Frm93_kira_caj_kad_kredit
End Sub

Private Sub L42_Text_Change()
'on error resume next
Call Frm93_kira_caj_gst_kad_kredit
Call Frm93_kira_potongan_kad_kredit
End Sub

Private Sub L43_Text_Change()
'on error resume next
Call Frm93_kira_potongan_kad_kredit
End Sub



Private Sub LV1_DblClick()
'on error resume next
frm93_LM_No_ID = vbNullString

If IsNumeric(Frm93.LV1.SelectedItem.Index) Then
    
    frm93_LM_No_ID = Frm93.LV1.SelectedItem.Index
    
    If frm93_LM_No_ID <> vbNullString Then
        
        user_level = MDI_frm1.L4_Text
        
        If user_level = "Admin" Or user_level = "HQ" Or user_level = "Developer" Then
        
            Frm93.Frm93_SM_Edit.Enabled = True
            Frm93.Frm93_SM_padam.Enabled = True
            Frm93.Frm93_SM_belum_siap.Enabled = True
                    
        ElseIf user_level = "Manager" Then
        
            Frm93.Frm93_SM_Edit.Enabled = True
            Frm93.Frm93_SM_padam.Enabled = False
            Frm93.Frm93_SM_belum_siap.Enabled = False
            
        Else
        
            Frm93.Frm93_SM_Edit.Enabled = False
            Frm93.Frm93_SM_padam.Enabled = False
            Frm93.Frm93_SM_belum_siap.Enabled = False
        
        End If

        GLOBAL_DISABLE = 0
        PopupMenu Frm93_PM_Menu
    
    Else
    
        MsgBox "Tiada Data.", vbInformation, "Info"
        
    End If
    
Else

    MsgBox "Tiada Data.", vbInformation, "Info"
    
End If
End Sub

Private Sub MSFlexGrid1_DblClick()
'on error resume next
frm93_LM_No_ID = vbNullString

If Frm93.MSFlexGrid1 <> vbNullString Then

    If IsNumeric(Frm93.MSFlexGrid1) Then
    
        frm93_LM_No_ID = Frm93.MSFlexGrid1.TextMatrix(Frm93.MSFlexGrid1, 2) 'No. ID
        
        If frm93_LM_No_ID <> vbNullString Then
        
            user_level = MDI_frm1.L4_Text
            
            If user_level = "Admin" Or user_level = "HQ" Or user_level = "Developer" Then
            
                Frm93.Frm93_SM_Edit.Enabled = True
                Frm93.Frm93_SM_padam.Enabled = True
                Frm93.Frm93_SM_belum_siap.Enabled = True
                        
            ElseIf user_level = "Manager" Then
            
                Frm93.Frm93_SM_Edit.Enabled = True
                Frm93.Frm93_SM_padam.Enabled = False
                Frm93.Frm93_SM_belum_siap.Enabled = False
                
            Else
            
                Frm93.Frm93_SM_Edit.Enabled = False
                Frm93.Frm93_SM_padam.Enabled = False
                Frm93.Frm93_SM_belum_siap.Enabled = False
            
            End If
    
            GLOBAL_DISABLE = 0
            PopupMenu Frm93_PM_Menu
          
        Else
        
            MsgBox "Tiada Data.", vbExclamation, "Info"
            
        End If
    
    Else
    
        MsgBox "Tiada Data.", vbExclamation, "Info"
        
    End If
    
End If
End Sub

Private Sub TB1_Change()
'on error resume next
Call frm93_kira_harga_tempahan1
End Sub
Private Sub TB10_Change()
'on error resume next
Call frm93_kira_harga_tempahan2
End Sub
Private Sub TB11_Change()
'on error resume next
Call frm93_harga_jualan
End Sub
Private Sub TB12_Change()
'on error resume next
Call frm93_harga_jualan
End Sub
Private Sub TB17_Change()
'on error resume next
If IsNumeric(Frm93.TB17) Then
    Frm93.TB22 = Format(Frm93.TB17, "#,##0.00")
Else
    Frm93.TB22 = "0.00"
End If
End Sub



Private Sub TB2_Change()
'on error resume next
Call frm93_kira_harga_tempahan1
End Sub
Private Sub TB20_Change()
'On Error Resume Next
Call frm93_jumlah_deposit

If GLOBAL_DISABLE = 0 Then
    If IsNumeric(Frm93.TB20) Then frm130.TB33 = Format(Frm93.TB20, "#,##0.00")
End If
End Sub
Private Sub TB21_Change()
'On Error Resume Next
Call Frm93_kira_jumlah_bayaran
End Sub
Private Sub TB22_Change()
'On Error Resume Next
Call frm93_jumlah_deposit
End Sub




Private Sub TB27_Change()
'On Error Resume Next
Call Frm93_kira_jumlah_bayaran
End Sub
Private Sub TB28_Change()
'On Error Resume Next
Call Frm93_kira_jumlah_bayaran
End Sub
Private Sub TB29_Change()
'On Error Resume Next
Call Frm93_kira_jumlah_bayaran
Call Frm93_kira_caj_kad_kredit
Call Frm93_kira_potongan_kad_kredit
End Sub
Private Sub TB3_Change()
'on error resume next
Call frm93_kira_harga_tempahan1
End Sub



Private Sub TB5_Change()
'on error resume next
If Frm93.CB1 = 1 And Frm93.TB5 <> vbNullString Then
    Frm93.Tmr2.Enabled = False
    Frm93.Tmr2.Enabled = True
    Frm93.Tmr2.Interval = 100
End If
End Sub
Private Sub TB8_Change()
'on error resume next
Call frm93_kira_harga_tempahan2
End Sub
Private Sub TB9_Change()
'on error resume next
Call frm93_kira_harga_tempahan2
End Sub
Private Sub Tmr1_Timer()
'on error resume next
Frm93.L1_Text = DateTime.Date
Frm93.L2_Text = DateTime.Time$
End Sub
Private Sub Tmr2_Timer()
'On Error Resume Next
If Frm93.CB1 = 1 And Frm93.TB5 <> vbNullString And Frm93.Tmr2.Enabled = True Then
    If Frm93.Tmr2.Interval = 100 Then
        If InStr(1, Frm93.TB5, "'") <> 0 Then
            MsgBox "No. Siri Produk Mengandungi Simbol Yang Tidak Sah , ['].", vbInformation, "Info"
            
            Frm93.TB5 = vbNullString
            Exit Sub
        End If
        
        Call Frm93_Call_Product_Detail
    End If
End If
End Sub





