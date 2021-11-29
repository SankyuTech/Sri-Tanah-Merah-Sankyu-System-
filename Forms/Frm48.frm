VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm48 
   Caption         =   "Payroll"
   ClientHeight    =   12915
   ClientLeft      =   120
   ClientTop       =   -34650
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
   Icon            =   "Frm48.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12915
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11175
      Left            =   -2520
      Picture         =   "Frm48.frx":0ECA
      ScaleHeight     =   11115
      ScaleWidth      =   1875
      TabIndex        =   84
      Top             =   -2160
      Width           =   1935
   End
   Begin VB.PictureBox Pic2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   10935
      Left            =   1200
      ScaleHeight     =   10935
      ScaleWidth      =   23475
      TabIndex        =   34
      Top             =   1560
      Visible         =   0   'False
      Width           =   23475
      Begin VB.CheckBox CB5 
         BackColor       =   &H8000000C&
         Caption         =   "Scanner Mode"
         Height          =   200
         Left            =   360
         TabIndex        =   159
         Top             =   4320
         Width           =   200
      End
      Begin VB.TextBox TB19 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   15195
         TabIndex        =   14
         Text            =   "TB19"
         Top             =   2520
         Width           =   1200
      End
      Begin VB.TextBox TB20 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   15195
         TabIndex        =   15
         Text            =   "TB20"
         Top             =   2880
         Width           =   1200
      End
      Begin VB.TextBox TB21 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   15195
         TabIndex        =   16
         Text            =   "TB21"
         Top             =   3240
         Width           =   1200
      End
      Begin VB.TextBox TB18 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   11595
         TabIndex        =   10
         Text            =   "TB18"
         Top             =   2160
         Width           =   1200
      End
      Begin VB.TextBox TB17 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   11595
         TabIndex        =   9
         Text            =   "TB17"
         Top             =   1800
         Width           =   1200
      End
      Begin VB.TextBox TB16 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   11595
         TabIndex        =   8
         Text            =   "TB16"
         Top             =   1440
         Width           =   1200
      End
      Begin VB.CheckBox CB3 
         BackColor       =   &H8000000C&
         Caption         =   "Scanner Mode"
         Height          =   200
         Left            =   360
         TabIndex        =   6
         Top             =   3840
         Width           =   200
      End
      Begin VB.CheckBox CB4 
         BackColor       =   &H8000000C&
         Caption         =   "Scanner Mode"
         Height          =   200
         Left            =   360
         TabIndex        =   7
         Top             =   4080
         Width           =   200
      End
      Begin VB.CommandButton CMD7 
         Caption         =   "Rekod Jualan Pekerja"
         Height          =   375
         Left            =   240
         MouseIcon       =   "Frm48.frx":236B
         MousePointer    =   99  'Custom
         TabIndex        =   132
         Top             =   10320
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.CommandButton CMD13 
         Caption         =   "Pengiraan Gaji"
         Height          =   375
         Left            =   2520
         MouseIcon       =   "Frm48.frx":2675
         MousePointer    =   99  'Custom
         TabIndex        =   131
         Top             =   4920
         Width           =   3375
      End
      Begin VB.CommandButton CMD9 
         Caption         =   "Simpan Data Dan Cetak Payslip"
         Height          =   375
         Left            =   10680
         MouseIcon       =   "Frm48.frx":297F
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   7800
         Width           =   3375
      End
      Begin VB.CheckBox CB2 
         BackColor       =   &H8000000C&
         Caption         =   "Scanner Mode"
         Height          =   200
         Left            =   12600
         TabIndex        =   18
         Top             =   6630
         Width           =   200
      End
      Begin VB.CheckBox CB1 
         BackColor       =   &H8000000C&
         Caption         =   "Scanner Mode"
         Height          =   200
         Left            =   12600
         TabIndex        =   17
         Top             =   6390
         Width           =   200
      End
      Begin VB.TextBox TB15 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   15195
         TabIndex        =   13
         Text            =   "TB15"
         Top             =   2160
         Width           =   1200
      End
      Begin VB.CommandButton CMD10 
         BackColor       =   &H000080FF&
         Caption         =   "Pengiraan Gaji"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2040
         MaskColor       =   &H00400000&
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Pengiraan Gaji Pekerja"
         Top             =   8520
         Visible         =   0   'False
         Width           =   4000
      End
      Begin VB.TextBox TB14 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   10515
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   9135
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.CommandButton CMD6 
         BackColor       =   &H000080FF&
         Caption         =   "Pengiraan Gaji"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2040
         MaskColor       =   &H00400000&
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Pengiraan Gaji Pekerja"
         Top             =   9240
         Visible         =   0   'False
         Width           =   4000
      End
      Begin VB.TextBox TB13 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Height          =   360
         Left            =   10515
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "TB13"
         Top             =   6960
         Width           =   2025
      End
      Begin VB.TextBox TB12 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Height          =   360
         Left            =   10515
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "TB12"
         Top             =   6585
         Width           =   2025
      End
      Begin VB.TextBox TB11 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Height          =   360
         Left            =   10515
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "TB11"
         Top             =   6210
         Width           =   2025
      End
      Begin VB.TextBox TB10 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   15195
         TabIndex        =   12
         Text            =   "TB10"
         Top             =   1800
         Width           =   1200
      End
      Begin VB.TextBox TB9 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   15195
         TabIndex        =   11
         Text            =   "TB9"
         Top             =   1440
         Width           =   1200
      End
      Begin VB.TextBox TB8 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   10515
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   8715
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.TextBox TB7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Height          =   360
         Left            =   19755
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   5175
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.TextBox TB6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Height          =   360
         Left            =   19755
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   4800
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.TextBox TB5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Height          =   360
         Left            =   19755
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   4440
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.TextBox TB4 
         BackColor       =   &H00C0C0FF&
         Height          =   360
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "TB4"
         Top             =   2830
         Width           =   5300
      End
      Begin VB.TextBox TB3 
         BackColor       =   &H00C0C0FF&
         Height          =   360
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "TB3"
         Top             =   2460
         Width           =   5300
      End
      Begin VB.TextBox TB2 
         BackColor       =   &H00C0C0FF&
         Height          =   360
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "TB2"
         Top             =   2080
         Width           =   5300
      End
      Begin VB.TextBox TB1 
         BackColor       =   &H00C0C0FF&
         Height          =   360
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "TB1"
         Top             =   1710
         Width           =   5300
      End
      Begin VB.ComboBox CBB4 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1330
         Width           =   5300
      End
      Begin VB.ComboBox CBB3 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   960
         Width           =   5300
      End
      Begin MSComCtl2.DTPicker DTPicker5 
         Height          =   360
         Left            =   10515
         TabIndex        =   20
         Top             =   7320
         Width           =   3765
         _ExtentX        =   6641
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
         Format          =   141754368
         CurrentDate     =   41561
      End
      Begin VB.Label L45_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L45_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   17880
         TabIndex        =   162
         Top             =   6600
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label L44_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L44_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   17880
         TabIndex        =   161
         Top             =   6360
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label L43_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L43_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   17880
         TabIndex        =   160
         Top             =   6120
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Shape Shape5 
         Height          =   855
         Left            =   8160
         Top             =   4480
         Width           =   4815
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "Komisen per gram                  : RM"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   8280
         TabIndex        =   158
         Top             =   4800
         Width           =   3435
      End
      Begin VB.Label L41_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L41_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   11400
         TabIndex        =   157
         Top             =   4815
         Width           =   2340
      End
      Begin VB.Label L40_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L40_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   11400
         TabIndex        =   156
         Top             =   4560
         Width           =   2220
      End
      Begin VB.Label Label56 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah berat jualan                : (g)"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   8280
         TabIndex        =   155
         Top             =   4560
         Width           =   3435
      End
      Begin VB.Label L42_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L42_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   11400
         TabIndex        =   154
         Top             =   5055
         Width           =   2340
      End
      Begin VB.Label Label67 
         BackStyle       =   0  'Transparent
         Caption         =   "Komisen berat jualan              : RM"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   8280
         TabIndex        =   153
         Top             =   5040
         Width           =   3435
      End
      Begin VB.Shape Shape4 
         Height          =   855
         Left            =   8160
         Top             =   3520
         Width           =   4815
      End
      Begin VB.Label L34_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L34_Text"
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   11400
         TabIndex        =   152
         Top             =   4080
         Width           =   1755
      End
      Begin VB.Label L33_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L33_Text"
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   11400
         TabIndex        =   151
         Top             =   3840
         Width           =   1755
      End
      Begin VB.Label L32_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L32_Text"
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   11400
         TabIndex        =   150
         Top             =   3600
         Width           =   1755
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Komisen B.Permata                 : %"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   8280
         TabIndex        =   149
         Top             =   3840
         Width           =   3435
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Komisen Jualan B.Permata       : RM"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   8280
         TabIndex        =   148
         Top             =   4080
         Width           =   3435
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Jualan B.Permata         : RM"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   8280
         TabIndex        =   147
         Top             =   3600
         Width           =   3435
      End
      Begin VB.Shape Shape3 
         Height          =   855
         Left            =   8160
         Top             =   2600
         Width           =   4815
      End
      Begin VB.Label L31_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L31_Text"
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   11400
         TabIndex        =   146
         Top             =   3120
         Width           =   1600
      End
      Begin VB.Label L30_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L30_Text"
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   11400
         TabIndex        =   145
         Top             =   2880
         Width           =   1600
      End
      Begin VB.Label L29_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L29_Text"
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   11400
         TabIndex        =   144
         Top             =   2640
         Width           =   1600
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Komisen Jualan B.Kemas         : RM"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   8280
         TabIndex        =   143
         Top             =   3120
         Width           =   3435
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Komisen B.Kemas                  : %"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   8280
         TabIndex        =   142
         Top             =   2880
         Width           =   3435
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Jualan B.Kemas          : RM"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   8280
         TabIndex        =   141
         Top             =   2640
         Width           =   3435
      End
      Begin VB.Label Label58 
         BackStyle       =   0  'Transparent
         Caption         =   "Zakat *                 : RM"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   13080
         TabIndex        =   140
         Top             =   2550
         Width           =   2355
      End
      Begin VB.Label Label57 
         BackStyle       =   0  'Transparent
         Caption         =   "Income Tax *        : RM"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   13080
         TabIndex        =   139
         Top             =   2925
         Width           =   2475
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "Advance *             : RM"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   13080
         TabIndex        =   138
         Top             =   3285
         Width           =   2475
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "Lain-lain *                                : RM "
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   8280
         TabIndex        =   137
         Top             =   2190
         Width           =   3555
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Elaun Perjalanan *                    : RM"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   8280
         TabIndex        =   136
         Top             =   1830
         Width           =   3555
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "Overtime *                               : RM"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   8280
         TabIndex        =   135
         Top             =   1470
         Width           =   3555
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah berat jualan                         Jumlah jualan barang kemas             Jumlah jualan barang permata"
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
         Height          =   975
         Left            =   600
         TabIndex        =   134
         Top             =   3795
         Width           =   3975
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila buat pilihan cara pengiraan komisen bagi pekerja ini."
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
         Height          =   255
         Left            =   360
         TabIndex        =   133
         Top             =   3480
         Width           =   7335
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Cara bayaran gaji dbuat :"
         Height          =   300
         Left            =   12600
         TabIndex        =   122
         Top             =   6120
         Width           =   3495
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank In"
         Height          =   300
         Left            =   12840
         TabIndex        =   121
         Top             =   6600
         Width           =   1695
      End
      Begin VB.Label Label51 
         BackStyle       =   0  'Transparent
         Caption         =   "Tunai"
         Height          =   300
         Left            =   12840
         TabIndex        =   120
         Top             =   6360
         Width           =   1695
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Bayaran               :"
         Height          =   255
         Left            =   8160
         TabIndex        =   119
         Top             =   7365
         Width           =   2415
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Komisen                     : RM"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   8280
         TabIndex        =   118
         Top             =   5480
         Width           =   3555
      End
      Begin VB.Label L36_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L36_Text"
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   11400
         TabIndex        =   117
         Top             =   5480
         Width           =   1755
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Lain-lain *            : RM"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   13080
         TabIndex        =   116
         Top             =   2200
         Width           =   2475
      End
      Begin VB.Label L35_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Pekerja ini LAYAK untuk mendapat komisen dari setiap jualan yang dilakukan."
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
         Height          =   360
         Left            =   8160
         TabIndex        =   115
         Top             =   5760
         Visible         =   0   'False
         Width           =   8115
      End
      Begin VB.Label L21_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "00.00"
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
         Height          =   375
         Left            =   14520
         TabIndex        =   83
         Top             =   9240
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label L20_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Kadar komisen adalah          % "
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
         Height          =   375
         Left            =   12000
         TabIndex        =   82
         Top             =   9240
         Visible         =   0   'False
         Width           =   3675
      End
      Begin VB.Label Label53 
         BackStyle       =   0  'Transparent
         Caption         =   "Komisen Investor (RM)"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   7920
         TabIndex        =   81
         Top             =   9165
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.Label L19_Text 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm48.frx":2C89
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
         Height          =   855
         Left            =   0
         TabIndex        =   80
         Top             =   7320
         Visible         =   0   'False
         Width           =   7695
      End
      Begin VB.Label L12_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "**Pekerja ini adalah layak untuk mendapat bonus profit investor (Big) **"
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
         Left            =   -120
         TabIndex        =   78
         Top             =   6720
         Visible         =   0   'False
         Width           =   7695
      End
      Begin VB.Label L11_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "**Pekerja ini adalah layak untuk mendapat bonus profit investor (Small) **"
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
         Left            =   -120
         TabIndex        =   77
         Top             =   6480
         Visible         =   0   'False
         Width           =   8415
      End
      Begin VB.Shape Shape2 
         Height          =   7695
         Left            =   7920
         Top             =   720
         Width           =   8895
      End
      Begin VB.Shape Shape1 
         Height          =   5295
         Left            =   8040
         Top             =   840
         Width           =   8655
      End
      Begin VB.Label L10_Text 
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
         Height          =   375
         Left            =   10080
         TabIndex        =   75
         Top             =   7320
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label L9_Text 
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
         Height          =   375
         Left            =   8280
         TabIndex        =   74
         Top             =   7320
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label L3_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "**Pekerja ini adalah layak untuk mendapat bonus profit**"
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
         Left            =   -120
         TabIndex        =   66
         Top             =   6240
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "Pendapatan Bersih     RM :"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   8160
         TabIndex        =   65
         Top             =   6975
         Width           =   2835
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "Penolakan                 RM :"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   8160
         TabIndex        =   64
         Top             =   6615
         Width           =   2835
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Pendapatan Kasar      RM :"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   8160
         TabIndex        =   63
         Top             =   6255
         Width           =   2835
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Penolakan"
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
         Left            =   12960
         TabIndex        =   62
         Top             =   960
         Width           =   3615
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Socso *                : RM"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   13080
         TabIndex        =   61
         Top             =   1840
         Width           =   2475
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "KWSP *                : RM"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   13080
         TabIndex        =   60
         Top             =   1470
         Width           =   2355
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Komisen Profit    (RM)"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   7920
         TabIndex        =   59
         Top             =   8745
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Komisen  : RM"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   17760
         TabIndex        =   58
         Top             =   5190
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate Komisen      : %"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   17760
         TabIndex        =   57
         Top             =   4800
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.Label Label54 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pendapatan"
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
         Left            =   8160
         TabIndex        =   55
         Top             =   960
         Width           =   4695
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Pengiraan gaji mengikut bulan payroll , Sila pilih ""Bulan Payroll"" dan isi ruangan yang wajib diisi."
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
         Left            =   240
         TabIndex        =   54
         Top             =   360
         Width           =   10335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan Payroll *          :"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   48
         Top             =   960
         Width           =   3555
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pekerja *        :"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   49
         Top             =   1350
         Width           =   3555
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Penuh *          :"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   50
         Top             =   1740
         Width           =   3555
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Kad Pengenalan*      :"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   51
         Top             =   2120
         Width           =   3555
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Gaji Pokok            : RM"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   52
         Top             =   2500
         Width           =   3555
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Elaun                   : RM"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   53
         Top             =   2880
         Width           =   3555
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Jualan     : RM"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   17760
         TabIndex        =   56
         Top             =   4440
         Visible         =   0   'False
         Width           =   2715
      End
   End
   Begin VB.PictureBox Pic5 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9375
      Left            =   11520
      ScaleHeight     =   9375
      ScaleWidth      =   21315
      TabIndex        =   86
      Top             =   2640
      Visible         =   0   'False
      Width           =   21315
      Begin VB.CommandButton CMD12 
         BackColor       =   &H000080FF&
         Caption         =   "Pengiraan Komisyen"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2760
         MaskColor       =   &H00400000&
         Style           =   1  'Graphical
         TabIndex        =   106
         ToolTipText     =   "Pengiraan Komisyen Pekerja"
         Top             =   2040
         Width           =   2805
      End
      Begin VB.ComboBox CBB5 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   360
         ItemData        =   "Frm48.frx":2D40
         Left            =   2205
         List            =   "Frm48.frx":2D42
         Style           =   2  'Dropdown List
         TabIndex        =   104
         Top             =   840
         Width           =   5085
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
         Height          =   8055
         Left            =   7320
         TabIndex        =   87
         Top             =   840
         Width           =   15975
         _ExtentX        =   28178
         _ExtentY        =   14208
         _Version        =   393216
         BackColor       =   16777088
         ForeColor       =   0
         BackColorFixed  =   8454016
         BackColorBkg    =   12640511
         GridColor       =   0
         WordWrap        =   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   1
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
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   360
         Left            =   2205
         TabIndex        =   100
         Top             =   1200
         Width           =   5085
         _ExtentX        =   8969
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
         Format          =   416808960
         CurrentDate     =   41561
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   360
         Left            =   2205
         TabIndex        =   101
         Top             =   1560
         Width           =   5085
         _ExtentX        =   8969
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
         Format          =   416808960
         CurrentDate     =   41561
      End
      Begin VB.Label L27_Text 
         Caption         =   "L27_Text"
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   360
         TabIndex        =   112
         Top             =   5760
         Width           =   6525
      End
      Begin VB.Label Label69 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Komisen                         :"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   111
         Top             =   5280
         Width           =   3300
      End
      Begin VB.Label L26_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L26_Text"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3480
         TabIndex        =   110
         Top             =   5280
         Width           =   4005
      End
      Begin VB.Label L23_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L23_Text"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3500
         TabIndex        =   109
         Top             =   4200
         Width           =   4005
      End
      Begin VB.Label Label66 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Berat Keseluruhan           :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   108
         Top             =   4200
         Width           =   3300
      End
      Begin VB.Label L25_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L25_Text"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   7440
         TabIndex        =   107
         Top             =   480
         Width           =   10335
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pekerja"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   105
         Top             =   855
         Width           =   2295
      End
      Begin VB.Label Label59 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Akhir "
         Height          =   255
         Left            =   240
         TabIndex        =   103
         Top             =   1605
         Width           =   2895
      End
      Begin VB.Label Label62 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Mula "
         Height          =   255
         Left            =   240
         TabIndex        =   102
         Top             =   1245
         Width           =   2535
      End
      Begin VB.Label Label82 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila pilih nama pekerja dan pilihan tarikh bagi pengiraan komisyen"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   99
         Top             =   480
         Width           =   10335
      End
      Begin VB.Label Label81 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan Barang Yang Dijual        :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   98
         Top             =   3840
         Width           =   3300
      End
      Begin VB.Label Label80 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Maklumat Jualan"
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
         Height          =   375
         Left            =   240
         TabIndex        =   97
         Top             =   3360
         Width           =   6735
      End
      Begin VB.Label L22_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L22_Text"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3500
         TabIndex        =   96
         Top             =   3840
         Width           =   4005
      End
      Begin VB.Label Label71 
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Left            =   19320
         TabIndex        =   95
         Top             =   1440
         Width           =   4000
      End
      Begin VB.Label Label70 
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Left            =   19320
         TabIndex        =   94
         Top             =   3480
         Visible         =   0   'False
         Width           =   4000
      End
      Begin VB.Label L24_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L24_Text"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3500
         TabIndex        =   93
         Top             =   4560
         Width           =   4005
      End
      Begin VB.Label Label68 
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Jualan Keseluruhan           :"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   92
         Top             =   4560
         Width           =   3300
      End
      Begin VB.Label Label65 
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Left            =   19320
         TabIndex        =   91
         Top             =   3840
         Visible         =   0   'False
         Width           =   4000
      End
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Left            =   19320
         TabIndex        =   90
         Top             =   4200
         Visible         =   0   'False
         Width           =   4000
      End
      Begin VB.Label Label60 
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Left            =   19305
         TabIndex        =   89
         Top             =   5040
         Visible         =   0   'False
         Width           =   4000
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Left            =   19305
         TabIndex        =   88
         Top             =   4560
         Visible         =   0   'False
         Width           =   4000
      End
   End
   Begin VB.PictureBox Pic1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10695
      Left            =   8400
      ScaleHeight     =   10695
      ScaleWidth      =   21315
      TabIndex        =   35
      Top             =   2520
      Visible         =   0   'False
      Width           =   21315
      Begin VB.CommandButton CMD5 
         Caption         =   "Simpan Data"
         Height          =   375
         Left            =   2520
         MouseIcon       =   "Frm48.frx":2D44
         MousePointer    =   99  'Custom
         TabIndex        =   126
         Top             =   3000
         Width           =   3375
      End
      Begin VB.ComboBox CBB1 
         Height          =   360
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1200
         Width           =   5500
      End
      Begin VB.ComboBox CBB2 
         Height          =   360
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1560
         Width           =   5500
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   405
         Left            =   2160
         TabIndex        =   32
         Top             =   1920
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   714
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
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   16744576
         Format          =   416808960
         CurrentDate     =   41572
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   405
         Left            =   2160
         TabIndex        =   33
         Top             =   2320
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   714
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
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   16744576
         Format          =   416808960
         CurrentDate     =   41572
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   9495
         Left            =   8040
         TabIndex        =   127
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   840
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   16748
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
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai tetapan yang telah disimpan ke dalam sistem."
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   8160
         TabIndex        =   47
         Top             =   600
         Width           =   7455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2040
         TabIndex        =   46
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2025
         TabIndex        =   44
         Top             =   1580
         Width           =   135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "**Sistem akan mengira bonus/komisen bagi setiap pekerja dalam tempoh tarikh yang dipilih."
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
         Height          =   615
         Left            =   480
         TabIndex        =   42
         Top             =   600
         Width           =   6615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila pilih tarikh bagi pengiraan payroll."
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   480
         TabIndex        =   41
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2025
         TabIndex        =   39
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label L4_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Mula"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   480
         TabIndex        =   38
         Top             =   1980
         Width           =   1395
      End
      Begin VB.Label L5_Text 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2040
         TabIndex        =   37
         Top             =   2360
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Akhir"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   480
         TabIndex        =   40
         Top             =   2360
         Width           =   1395
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   480
         TabIndex        =   45
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   480
         TabIndex        =   43
         Top             =   1580
         Width           =   1395
      End
   End
   Begin VB.PictureBox Pic3 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11295
      Left            =   1680
      ScaleHeight     =   11295
      ScaleWidth      =   21075
      TabIndex        =   36
      Top             =   720
      Visible         =   0   'False
      Width           =   21075
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   9975
         Left            =   120
         TabIndex        =   128
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   480
         Width           =   20805
         _ExtentX        =   36698
         _ExtentY        =   17595
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
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai rekod bayaran gaji kepada pekerja."
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   76
         Top             =   120
         Width           =   5055
      End
   End
   Begin VB.PictureBox Pic4 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11295
      Left            =   360
      ScaleHeight     =   11295
      ScaleWidth      =   21075
      TabIndex        =   67
      Top             =   1080
      Visible         =   0   'False
      Width           =   21075
      Begin VB.CommandButton CMD8 
         Caption         =   "Maklumat Gaji Pekerja"
         Height          =   375
         Left            =   240
         MouseIcon       =   "Frm48.frx":304E
         MousePointer    =   99  'Custom
         TabIndex        =   130
         Top             =   10320
         Width           =   3375
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   9855
         Left            =   120
         TabIndex        =   129
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   360
         Width           =   15975
         _ExtentX        =   28178
         _ExtentY        =   17383
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
      Begin VB.Label L8_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L8_Text"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   19200
         TabIndex        =   114
         Top             =   1800
         Width           =   4005
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Jualan Keseluruhan    :"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   16560
         TabIndex        =   113
         Top             =   1800
         Width           =   3240
      End
      Begin VB.Label L13_Text 
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Left            =   19320
         TabIndex        =   79
         Top             =   2520
         Width           =   4000
      End
      Begin VB.Label L7_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L7_Text"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   19200
         TabIndex        =   73
         Top             =   1440
         Width           =   4005
      End
      Begin VB.Label L6_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L6_Text"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   19200
         TabIndex        =   72
         Top             =   1080
         Width           =   4005
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Berat Keseluruhan    :"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   16575
         TabIndex        =   71
         Top             =   1440
         Width           =   3240
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Maklumat Jualan Pekerja"
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
         Left            =   16320
         TabIndex        =   70
         Top             =   480
         Width           =   6975
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan Barang Yang Dijual : "
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   16575
         TabIndex        =   69
         Top             =   1080
         Width           =   3120
      End
      Begin VB.Label L28_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L28_Text"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   68
         Top             =   120
         Width           =   14415
      End
   End
   Begin VB.Label L39_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai Gaji / Payslip"
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
      Left            =   4080
      MouseIcon       =   "Frm48.frx":3358
      MousePointer    =   99  'Custom
      TabIndex        =   125
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label L38_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pengiraan Gaji Pekerja"
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
      Left            =   1920
      MouseIcon       =   "Frm48.frx":3662
      MousePointer    =   99  'Custom
      TabIndex        =   124
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label L37_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tetapan Payroll"
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
      Left            =   0
      MouseIcon       =   "Frm48.frx":396C
      MousePointer    =   99  'Custom
      TabIndex        =   123
      Top             =   120
      Width           =   1815
   End
   Begin VB.Menu Frm_Menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm48_PadamTetapan 
         Caption         =   "Padam Tetapan"
      End
   End
   Begin VB.Menu Frm48_Menu2 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm48_PadamData 
         Caption         =   "Padam Data"
      End
      Begin VB.Menu Frm48_CetakPayslip 
         Caption         =   "Cetak Payslip"
      End
      Begin VB.Menu frm48_sm_excel 
         Caption         =   "Export Excel"
      End
   End
End
Attribute VB_Name = "Frm48"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CB1_Click()
'On Error Resume Next
If Frm48.CB1 = 1 Then
    Frm48.CB2 = 0
End If
End Sub
Private Sub CB2_Click()
'On Error Resume Next
If Frm48.CB2 = 1 Then
    Frm48.CB1 = 0
End If
End Sub

Private Sub CB3_Click()
'On Error Resume Next
'If Frm48.CB3 = 1 Then
    
    Frm48.L10_Text = 0
    'Frm48.CB4 = 0
    'Frm48.Pic7.Visible = True
    'Frm48.Pic6.Visible = False
    Frm48.L10_Text = "0"
'End If
End Sub
Private Sub CB4_Click()
'On Error Resume Next
'If Frm48.CB4 = 1 Then
    
    Frm48.L10_Text = 0
    'Frm48.CB3 = 0
    'Frm48.Pic6.Visible = True
    'Frm48.Pic7.Visible = False
    Frm48.L10_Text = "0"
'End If
End Sub

Private Sub CB5_Click()
'On Error Resume Next
'If Frm48.CB4 = 1 Then
    
    Frm48.L10_Text = 0
    'Frm48.CB3 = 0
    'Frm48.Pic6.Visible = True
    'Frm48.Pic7.Visible = False
    Frm48.L10_Text = "0"
'End If
End Sub

Private Sub CBB3_Change()
'On Error Resume Next
'Frm48.TB1 = vbNullString
'Frm48.TB2 = vbNullString
'Frm48.TB3 = vbNullString
'Frm48.TB4 = vbNullString
Frm48.TB5 = vbNullString
'Frm48.TB6 = vbNullString
Frm48.TB7 = vbNullString
Frm48.TB8 = vbNullString
'Frm48.TB9 = vbNullString
'Frm48.TB10 = vbNullString
'Frm48.TB11 = "0.00"
'Frm48.TB12 = "0.00"
'Frm48.TB13 = "0.00"
Frm48.L10_Text = 0
'Frm48.L20_Text.Visible = False
'Frm48.L21_Text.Visible = False
End Sub
Private Sub CBB3_Click()
'On Error Resume Next
'Frm48.TB1 = vbNullString
'Frm48.TB2 = vbNullString
'Frm48.TB3 = vbNullString
'Frm48.TB4 = vbNullString
Frm48.TB5 = vbNullString
'Frm48.TB6 = vbNullString
Frm48.TB7 = vbNullString
Frm48.TB8 = vbNullString
'Frm48.TB9 = vbNullString
'Frm48.TB10 = vbNullString
'Frm48.TB11 = "0.00"
'Frm48.TB12 = "0.00"
'Frm48.TB13 = "0.00"
Frm48.L10_Text = 0
'Frm48.L20_Text = vbNullString
'Frm48.L21_Text.Visible = False
End Sub
Private Sub CBB4_Change()
'On Error Resume Next
Frm48.TB1 = vbNullString
Frm48.TB2 = vbNullString
Frm48.TB3 = vbNullString
Frm48.TB4 = vbNullString
Frm48.TB5 = vbNullString
Frm48.TB7 = vbNullString
Frm48.TB8 = vbNullString

'%%%% TukangemaS %%%%
Frm48.L29_Text = "0.00"
Frm48.L31_Text = "0.00"
Frm48.L32_Text = "0.00"
Frm48.L34_Text = "0.00"
Frm48.L36_Text = "0.00"
'%%%% TukangemaS %%%%

'Frm48.TB11 = "0.00"
'Frm48.TB12 = "0.00"
'Frm48.TB13 = "0.00"
Frm48.L10_Text = 0
Frm48.L20_Text.Visible = False
Frm48.L21_Text.Visible = False

If Frm48.CBB4 <> vbNullString Then
    EMP_NO_PENJUAL = Split(Frm48.CBB4, "  |  ")(1)
    Nama_PENJUAL = Split(Frm48.CBB4, "  |  ")(0)
End If

DATA_INVESTOR = 0 '0 : Tiada Komisen Bagi Investor , 1 : Ada Komisen Bagi Investor

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where NoPekerja='" & EMP_NO_PENJUAL & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!Nama) Then Frm48.TB1 = rs!Nama 'Nama Pekerja
    If Not IsNull(rs!NoIC) Then Frm48.TB2 = rs!NoIC 'No IC
    If Not IsNull(rs!Gaji) Then Frm48.TB3 = Format(rs!Gaji, "0.00") 'Gaji
    If Not IsNull(rs!Elaun) Then Frm48.TB4 = Format(rs!Elaun, "0.00") 'Elaun
    If Not IsNull(rs!komisen) Then
        If rs!komisen = 1 Then
            Frm48.L35_Text.Visible = True
            'Frm48.CMD7.Visible = True
        Else
            Frm48.L35_Text.Visible = False
            'Frm48.CMD7.Visible = False
        End If
    Else
        Frm48.L35_Text.Visible = False
    End If
    
    
    If rs!ElaunProfit = 1 Then
        Frm48.L3_Text.Visible = True
    Else
        Frm48.L3_Text.Visible = False
    End If
    If rs!InvestorSmall = 1 Then
        Frm48.L11_Text.Visible = True
        Frm48.L19_Text.Visible = True
        INVESTOR_TYPE = 1 '1 : Investor (Small) , 2 : Investor (Big)
    Else
        Frm48.L11_Text.Visible = False
        Frm48.L19_Text.Visible = False
    End If
    If rs!InvestorBig = 1 Then
        Frm48.L12_Text.Visible = True
        Frm48.L19_Text.Visible = True
        INVESTOR_TYPE = 2 '1 : Investor (Small) , 2 : Investor (Big)
    Else
        Frm48.L12_Text.Visible = False
        Frm48.L19_Text.Visible = False
    End If
    If rs!InvestorSmall = 1 Or rs!InvestorBig = 1 Then
        Frm48.TB9.BackColor = vbWhite
        Frm48.TB10.BackColor = vbWhite
        Frm48.TB9 = 0
        Frm48.TB10 = 0
        Frm48.TB9.Locked = True
        Frm48.TB10.Locked = True
        Frm48.L19_Text.Visible = True
        DATA_INVESTOR = 1 '0 : Tiada Komisen Bagi Investor , 1 : Ada Komisen Bagi Investor
    End If
    'If rs!InvestorSmall = 0 And rs!InvestorBig = 0 Then
    '    Frm48.TB9.BackColor = &HC0C0FF
    '    Frm48.TB10.BackColor = &HC0C0FF
    '    'Frm48.TB9 = vbNullString
    '    'Frm48.TB10 = vbNullString
    '    Frm48.TB9.Locked = False
    '    Frm48.TB10.Locked = False
    '    Frm48.L19_Text.Visible = False
    'End If
End If

rs.Close
Set rs = Nothing

If DATA_INVESTOR = 1 Then '0 : Tiada Komisen Bagi Investor , 1 : Ada Komisen Bagi Investor
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from default_setting where Default1='" & "Default" & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If INVESTOR_TYPE = 1 Then '1 : Investor (Small) , 2 : Investor (Big)
            If Not IsNull(rs!KomisenSmall) Then
                Frm48.L21_Text = rs!KomisenSmall 'Komisen Investor (Small)
            Else
                Frm48.L21_Text = "0.00" 'Komisen Investor (Small)
            End If
        End If
        If INVESTOR_TYPE = 2 Then '1 : Investor (Small) , 2 : Investor (Big)
            If Not IsNull(rs!KomisenBig) Then
                Frm48.L21_Text = rs!KomisenBig 'Komisen Investor (Big)
            Else
                Frm48.L21_Text = "0.00" 'Komisen Investor (Big)
            End If
        End If
    End If
    
    Frm48.L20_Text.Visible = True
    Frm48.L21_Text.Visible = True
End If
End Sub
Private Sub CBB4_Click()
'On Error Resume Next
Frm48.TB1 = vbNullString
Frm48.TB2 = vbNullString
Frm48.TB3 = vbNullString
Frm48.TB4 = vbNullString
Frm48.TB5 = vbNullString
'Frm48.TB6 = vbNullString
Frm48.TB7 = vbNullString
Frm48.TB8 = vbNullString

'%%%% TukangemaS %%%%
Frm48.L29_Text = "0.00"
Frm48.L31_Text = "0.00"
Frm48.L32_Text = "0.00"
Frm48.L34_Text = "0.00"
Frm48.L36_Text = "0.00"
'%%%% TukangemaS %%%%

'Frm48.TB9 = vbNullString
'Frm48.TB10 = vbNullString
'Frm48.TB11 = "0.00"
'Frm48.TB12 = "0.00"
'Frm48.TB13 = "0.00"
Frm48.L10_Text = 0
Frm48.L20_Text.Visible = False
Frm48.L21_Text.Visible = False
'NAMA_PEKERJA = Frm48.CBB4
If Frm48.CBB4 <> vbNullString Then
    EMP_NO_PENJUAL = Split(Frm48.CBB4, "  |  ")(1)
    Nama_PENJUAL = Split(Frm48.CBB4, "  |  ")(0)
End If
DATA_INVESTOR = 0 '0 : Tiada Komisen Bagi Investor , 1 : Ada Komisen Bagi Investor

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where NoPekerja='" & EMP_NO_PENJUAL & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!Nama) Then Frm48.TB1 = rs!Nama 'Nama Pekerja
    If Not IsNull(rs!NoIC) Then Frm48.TB2 = rs!NoIC 'No IC
    If Not IsNull(rs!Gaji) Then Frm48.TB3 = Format(rs!Gaji, "0.00") 'Gaji
    If Not IsNull(rs!Elaun) Then Frm48.TB4 = Format(rs!Elaun, "0.00") 'Elaun
    If Not IsNull(rs!komisen) Then
        If rs!komisen = 1 Then
            Frm48.L35_Text.Visible = True
            'Frm48.CMD7.Visible = True
        Else
            Frm48.L35_Text.Visible = False
            'Frm48.CMD7.Visible = False
        End If
    Else
        Frm48.L35_Text.Visible = False
    End If
    
    
    If rs!ElaunProfit = 1 Then
        Frm48.L3_Text.Visible = True
    Else
        Frm48.L3_Text.Visible = False
    End If
    If rs!InvestorSmall = 1 Then
        Frm48.L11_Text.Visible = True
        Frm48.L19_Text.Visible = True
        INVESTOR_TYPE = 1 '1 : Investor (Small) , 2 : Investor (Big)
    Else
        Frm48.L11_Text.Visible = False
        Frm48.L19_Text.Visible = False
    End If
    If rs!InvestorBig = 1 Then
        Frm48.L12_Text.Visible = True
        Frm48.L19_Text.Visible = True
        INVESTOR_TYPE = 2 '1 : Investor (Small) , 2 : Investor (Big)
    Else
        Frm48.L12_Text.Visible = False
        Frm48.L19_Text.Visible = False
    End If
    If rs!InvestorSmall = 1 Or rs!InvestorBig = 1 Then
        Frm48.TB9.BackColor = vbWhite
        Frm48.TB10.BackColor = vbWhite
        Frm48.TB9 = 0
        Frm48.TB10 = 0
        Frm48.TB9.Locked = True
        Frm48.TB10.Locked = True
        Frm48.L19_Text.Visible = True
        DATA_INVESTOR = 1 '0 : Tiada Komisen Bagi Investor , 1 : Ada Komisen Bagi Investor
    End If
    'If rs!InvestorSmall = 0 And rs!InvestorBig = 0 Then
    '    Frm48.TB9.BackColor = &HC0C0FF
    '    Frm48.TB10.BackColor = &HC0C0FF
    '    'Frm48.TB9 = vbNullString
    '    'Frm48.TB10 = vbNullString
    '    Frm48.TB9.Locked = False
    '    Frm48.TB10.Locked = False
    '    Frm48.L19_Text.Visible = False
    'End If
End If

rs.Close
Set rs = Nothing

If DATA_INVESTOR = 1 Then '0 : Tiada Komisen Bagi Investor , 1 : Ada Komisen Bagi Investor
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from default_setting where Default1='" & "Default" & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If INVESTOR_TYPE = 1 Then '1 : Investor (Small) , 2 : Investor (Big)
            If Not IsNull(rs!KomisenSmall) Then
                Frm48.L21_Text = rs!KomisenSmall 'Komisen Investor (Small)
            Else
                Frm48.L21_Text = "0.00" 'Komisen Investor (Small)
            End If
        End If
        If INVESTOR_TYPE = 2 Then '1 : Investor (Small) , 2 : Investor (Big)
            If Not IsNull(rs!KomisenBig) Then
                Frm48.L21_Text = rs!KomisenBig 'Komisen Investor (Big)
            Else
                Frm48.L21_Text = "0.00" 'Komisen Investor (Big)
            End If
        End If
    End If
    
    rs.Close
    Set rs = Nothing

    Frm48.L20_Text.Visible = True
    Frm48.L21_Text.Visible = True
End If
End Sub
Private Sub CMD10_Click()
'On Error Resume Next
Dim TM As Date 'Tarikh Mula
Dim TA As Date 'Tarikh Akhir
Dim JUMLAH_BERAT As Double 'Total Berat Yang Dijual Oleh Staff
Dim JUMLAH_HARGA_JUALAN As Double 'Total Berat Yang Terjual Di Kedai
Dim Err(10)
Dim a As Double

x = 0
JUMLAH_HARGA_JUALAN = 0
JUMLAH_BERAT = 0

If Frm48.CBB3 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Bulan Payroll]."
End If
If Frm48.CBB4 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Nama Pekerja]."
End If
If Frm48.TB9 = vbNullString Or (Frm48.TB9 <> vbNullString And Not IsNumeric(Frm48.TB9)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [KWSP] , Hanya NOMBOR Yang Dibenarkan Dalam Ruangan Ini."
End If
If Frm48.TB10 = vbNullString Or (Frm48.TB10 <> vbNullString And Not IsNumeric(Frm48.TB10)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Socso] , Hanya NOMBOR Yang Dibenarkan Dalam Ruangan Ini."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Frm48.MSFlexGrid2.Clear
    Frm48.MSFlexGrid2.RowHeight(0) = 600
    Frm48.MSFlexGrid2.FormatString = "No.|<No.|<No. Siri|<Kategori|<Tarikh Jualan|<Berat Jualan (g)|<Harga Jualan (RM)"
    
    Frm48.MSFlexGrid2.Rows = 1
    Frm48.MSFlexGrid2.ColWidth(0) = 600
    Frm48.MSFlexGrid2.ColWidth(1) = 0
    Frm48.MSFlexGrid2.ColWidth(2) = 3500
    Frm48.MSFlexGrid2.ColWidth(3) = 4500
    Frm48.MSFlexGrid2.ColWidth(4) = 2400
    Frm48.MSFlexGrid2.ColWidth(5) = 2400
    Frm48.MSFlexGrid2.ColWidth(6) = 2400
    
    PAYROLL = Frm48.CBB3 'Tetapan Bulan Payroll
    If InStr(1, PAYROLL, " ") <> 0 Then
        PAY_BULAN = Split(PAYROLL, " ")(0)
        PAY_TAHUN = Split(PAYROLL, " ")(1)
    End If
    
    If Frm48.CBB4 <> vbNullString Then
        EMP_NO_PENJUAL = Split(Frm48.CBB4, "  |  ")(1)
        Nama_PENJUAL = Split(Frm48.CBB4, "  |  ")(0)
    End If
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from tetapan_Payslip where Bulan='" & PAY_BULAN & "' AND Tahun='" & PAY_TAHUN & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!TarikhMula) Then TM = rs!TarikhMula 'Tarikh Mula
        If Not IsNull(rs!TarikhAkhir) Then TA = rs!TarikhAkhir 'Tarikh Akhir
    End If
    
    rs.Close
    Set rs = Nothing
    
    '///////Pengiraan Jumlah Jualan Staff Dan Jualan Kedai//////////// ####TukangemaS Tidak Gunakan Cara Ini######## START
    Frm49_LM_EMP_NO_PENJUAL = Split(Frm48.CBB4, "  |  ")(1)
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 23_senarai_jualan where no_pekerja='" & Frm49_LM_EMP_NO_PENJUAL & "' AND tarikh BETWEEN '" & TM & "'  AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        Y = Y + 1
        Frm48.MSFlexGrid2.Rows = Y + 1
        Frm48.MSFlexGrid2.TextMatrix(Y, 0) = Y
        Frm48.MSFlexGrid2.TextMatrix(Y, 1) = Y
        If Not IsNull(rs!no_siri_Produk) Then Frm48.MSFlexGrid2.TextMatrix(Y, 2) = rs!no_siri_Produk 'No. Siri Produk
        If Not IsNull(rs!kategori_Produk) Then Frm48.MSFlexGrid2.TextMatrix(Y, 3) = rs!kategori_Produk 'Nama Produk
        If Not IsNull(rs!tarikh) Then Frm48.MSFlexGrid2.TextMatrix(Y, 4) = rs!tarikh 'Tarikh Jualan
        If Not IsNull(rs!berat_jualan) Then Frm48.MSFlexGrid2.TextMatrix(Y, 5) = Format(rs!berat_jualan, "0.00") 'Berat Jualan
        If Not IsNull(rs!harga_jualan) Then Frm48.MSFlexGrid2.TextMatrix(Y, 6) = Format(rs!harga_jualan, "0.00") 'Harga Jualan
        rs.MoveNext
    Wend
        
    rs.Close
    Set rs = Nothing
        
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 22_jualan where no_pekerja='" & Frm49_LM_EMP_NO_PENJUAL & "' AND tarikh BETWEEN '" & TM & "'  AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        If Not IsNull(rs!harga_barang) Then
            If IsNumeric(rs!harga_barang) Then JUMLAH_HARGA_JUALAN = JUMLAH_HARGA_JUALAN + rs!harga_barang 'Jumlah Harga
        End If
        If Not IsNull(rs!JUMLAH_BERAT) Then
            If IsNumeric(rs!JUMLAH_BERAT) Then JUMLAH_BERAT = JUMLAH_BERAT + rs!JUMLAH_BERAT 'Jumlah Berat
        End If
        rs.MoveNext
    Wend
        
    rs.Close
    Set rs = Nothing

    If Frm48.L3_Text.Visible = True And IsNumeric(JUMLAH_BERAT_OVERALL) And IsNumeric(Frm48.L9_Text) Then
        a = Frm48.L9_Text 'Format(Frm48.L9_Text / 100, "0.00")
        Frm48.TB8 = Format(JUMLAH_BERAT_OVERALL * a, "0.00") 'Jumlah Profit
    Else
        Frm48.TB8 = Format(0, "0.00") 'Jumlah Profit
    End If
    
    If Y = vbNullString Then Y = 0
    'Frm48.TB5 = Format(JUMLAH_BERAT, "0.00") 'Jumlah Berat Terjual
    Frm48.TB5 = JUMLAH_RESIT 'Format(JUMLAH_BERAT, "0.00") 'Jumlah Berat Terjual
    Frm48.L6_Text = ":   " & Y 'Bilangan Barang Terjual
    Frm48.L7_Text = ":   " & Format(JUMLAH_BERAT, "0.00 g") 'Jumlah Berat Terjual Oleh Staff
    Frm48.L13_Text = ":   " & Format(JUMLAH_BERAT_OVERALL, "0.00 g") 'Jumlah Berat Terjual Di Kedai
    'Frm48.L14_Text = ":   RM " & Format(JUMLAH_PROFIT, "0.00") 'Jumlah Keuntungan Dari Jualan Emas
    'Frm48.L15_Text = ":   RM " & Format(TOTAL_SERVICE, "0.00") 'Jumlah Servis
    'Frm48.L16_Text = ":   RM " & Format(TOTAL_EXPENSES, "0.00") 'Jumlah Perbelanjaan Kedai
    'Frm48.L18_Text = ":   RM " & Format(TOTAL_GAJI, "0.00") 'Jumlah Pembayaran Gaji
    'Frm48.L17_Text = ":   RM " & Format(JUMLAH_PROFIT + TOTAL_SERVICE - TOTAL_EXPENSES - TOTAL_GAJI, "0.00") 'Jumlah Keuntungan Bersih
    If Frm48.L21_Text.Visible = True Then
        If IsNumeric(JUMLAH_PROFIT) And IsNumeric(Frm48.L21_Text) Then
            b = Frm48.L21_Text / 100
            Frm48.TB14 = Format(b * JUMLAH_PROFIT, "0.00") 'Komisen Bagi Investor
        End If
    Else
        Frm48.TB14 = Format(0, "0.00") 'Komisen Bagi Investor
    End If
    Frm48.L10_Text = 1
    
    MsgBox "Pengiraan Gaji Bagi Pekerja Telah Selesai.", vbInformation, "Info"
End If
End Sub
Private Sub CMD12_Click()
'On Error Resume Next
Dim TM As Date
Dim TA As Date
Dim JUMLAH_HARGA_JUALAN As Double
Dim JUMLAH_BERAT As Double
Dim Frm48_LM_RATE_KOMISEN As Double

Y = 0
JUMLAH_HARGA_JUALAN = 0
JUMLAH_BERAT = 0
Frm48_LM_RATE_KOMISEN = 0

If Frm48.CBB5 = vbNullString Then
    MsgBox "Sila Pilih Nama Pekerja.", vbInformation, "Info"
    Exit Sub
End If

If Frm48.CBB5 <> vbNullString Then
    Frm48_LM_EMP_NO = Split(Frm48.CBB5, "  |  ")(1)
    Frm48_LM_Nama = Split(Frm48.CBB5, "  |  ")(0)
End If

TM = Frm48.DTPicker3
TA = Frm48.DTPicker4

Note = "Pengiraan Komisyen Bagi " & Frm48_LM_Nama & " Dari " & TM & " Hingga " & TA & " ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    
    Frm48.L25_Text = "Pengiraan Komisyen Bagi " & Frm48_LM_Nama & " Dari " & TM & " Hingga " & TA 'Header

    Frm48.MSFlexGrid4.Clear
    Frm48.MSFlexGrid4.RowHeight(0) = 600
    Frm48.MSFlexGrid4.FormatString = "No.|<No.|<No. Siri|<Kategori|<Tarikh Jualan|<Berat Jualan (g)|<Harga Jualan (RM)"
    
    Frm48.MSFlexGrid4.Rows = 1
    Frm48.MSFlexGrid4.ColWidth(0) = 600
    Frm48.MSFlexGrid4.ColWidth(1) = 0
    Frm48.MSFlexGrid4.ColWidth(2) = 3500
    Frm48.MSFlexGrid4.ColWidth(3) = 4500
    Frm48.MSFlexGrid4.ColWidth(4) = 2400
    Frm48.MSFlexGrid4.ColWidth(5) = 2400
    Frm48.MSFlexGrid4.ColWidth(6) = 2400
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 23_senarai_jualan where no_pekerja='" & Frm48_LM_EMP_NO & "' AND tarikh BETWEEN '" & TM & "'  AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        Y = Y + 1
        Frm48.MSFlexGrid4.Rows = Y + 1
        Frm48.MSFlexGrid4.TextMatrix(Y, 0) = Y
        Frm48.MSFlexGrid4.TextMatrix(Y, 1) = Y
        If Not IsNull(rs!no_siri_Produk) Then Frm48.MSFlexGrid4.TextMatrix(Y, 2) = rs!no_siri_Produk 'No. Siri Produk
        If Not IsNull(rs!kategori_Produk) Then Frm48.MSFlexGrid4.TextMatrix(Y, 3) = rs!kategori_Produk 'Nama Produk
        If Not IsNull(rs!tarikh) Then Frm48.MSFlexGrid4.TextMatrix(Y, 4) = rs!tarikh 'Tarikh Jualan
        If Not IsNull(rs!berat_jualan) Then Frm48.MSFlexGrid4.TextMatrix(Y, 5) = Format(rs!berat_jualan, "0.00") 'Berat Jualan
        If Not IsNull(rs!harga_jualan) Then Frm48.MSFlexGrid4.TextMatrix(Y, 6) = Format(rs!harga_jualan, "0.00") 'Harga Jualan
        rs.MoveNext
    Wend
        
    rs.Close
    Set rs = Nothing
        
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 22_jualan where no_pekerja='" & Frm48_LM_EMP_NO & "' AND tarikh BETWEEN '" & TM & "'  AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        If Not IsNull(rs!harga_barang) Then
            If IsNumeric(rs!harga_barang) Then JUMLAH_HARGA_JUALAN = JUMLAH_HARGA_JUALAN + rs!harga_barang 'Jumlah Harga
        End If
        If Not IsNull(rs!JUMLAH_BERAT) Then
            If IsNumeric(rs!JUMLAH_BERAT) Then JUMLAH_BERAT = JUMLAH_BERAT + rs!JUMLAH_BERAT 'Jumlah Berat
        End If
        rs.MoveNext
    Wend
        
    rs.Close
    Set rs = Nothing
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If rs!Default1 = "Default" Then
            Frm48_LM_RATE_KOMISEN = rs!komisen 'Kadar Komisen (%)
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    
    Frm48.L22_Text = Y
    Frm48.L23_Text = Format(JUMLAH_BERAT, "0.00 g")
    Frm48.L24_Text = "RM " & Format(JUMLAH_HARGA_JUALAN, "0.00")
    Frm48.L26_Text = "RM " & Format((Frm48_LM_RATE_KOMISEN / 100) * JUMLAH_HARGA_JUALAN, "0.00")
    
    Frm48.L27_Text = "Pengiraan komisen ini adalah mengikut ketetapan kadar komisen pada " & Frm48_LM_RATE_KOMISEN & "%"
End If
End Sub
Private Sub CMD13_Click()
'On Error Resume Next
Dim TM As Date
Dim TA As Date
Dim JUMLAH_HARGA_JUALAN_BK As Double
Dim JUMLAH_HARGA_JUALAN_PERMATA As Double
Dim JUMLAH_BERAT As Double
Dim Frm48_LM_RATE_KOMISEN As Double
Dim Frm48_LM_GAJI_BASIC As Double
Dim Frm48_LM_ELAUN As Double
Dim Frm48_LM_KWSP As Double
Dim Frm48_LM_SOCSO As Double
Dim Err(5)

Y = 0
JUMLAH_HARGA_JUALAN_BK = 0
JUMLAH_HARGA_JUALAN_PERMATA = 0
JUMLAH_BERAT = 0
Frm48_LM_RATE_KOMISEN = 0
Frm48_LM_GAJI_BASIC = 0
Frm48_LM_ELAUN = 0
Frm48_LM_KWSP = 0
Frm48_LM_SOCSO = 0

'If Frm48.CB3 = 0 And Frm48.CB4 = 0 Then
'    x = x + 1
'    Err(x) = "Sila buat pilihan cara pengiraan komisen pekerja."
'End If
If Frm48.CBB3 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Bulan Payroll]."
End If
If Frm48.CBB4 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Nama Pekerja]."
End If
If Frm48.TB9 = vbNullString Or (Frm48.TB9 <> vbNullString And Not IsNumeric(Frm48.TB9)) Then
    x = x + 1
    Err(x) = "Sila masukkan [KWSP] , Hanya NOMBOR yang dibenarkan dalam ruangan ini."
End If
If Frm48.TB10 = vbNullString Or (Frm48.TB10 <> vbNullString And Not IsNumeric(Frm48.TB10)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Socso] , Hanya NOMBOR yang dibenarkan dalam ruangan ini."
End If
If Frm48.TB15 = vbNullString Or (Frm48.TB15 <> vbNullString And Not IsNumeric(Frm48.TB15)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Potongan Lain-lain] , Hanya NOMBOR yang dibenarkan dalam ruangan ini."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    If Frm48.CBB4 <> vbNullString Then
        Frm48_LM_EMP_NO = Split(Frm48.CBB4, "  |  ")(1)
        Frm48_LM_Nama = Split(Frm48.CBB4, "  |  ")(0)
    End If
    
    Frm48_LM_PAYROLL = Frm48.CBB3 'Tetapan Bulan Payroll
    
    Note = "Pengiraan Gaji Bagi " & Frm48_LM_Nama & " Untuk " & Frm48_LM_PAYROLL & " ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
        
        If InStr(1, Frm48_LM_PAYROLL, " ") <> 0 Then
            Frm48_LM_PAY_BULAN = Split(Frm48_LM_PAYROLL, " ")(0)
            Frm48_LM_PAY_TAHUN = Split(Frm48_LM_PAYROLL, " ")(1)
        End If
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from tetapan_Payslip where Bulan='" & Frm48_LM_PAY_BULAN & "' AND Tahun='" & Frm48_LM_PAY_TAHUN & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!TarikhMula) Then TM = rs!TarikhMula 'Tarikh Mula
            If Not IsNull(rs!TarikhAkhir) Then TA = rs!TarikhAkhir 'Tarikh Akhir
        End If
        
        rs.Close
        Set rs = Nothing
        
        If TM = "00:00:00" Or TA = "00:00:00" Then
            
            MsgBox "Terdapat teknikal error pada penetapan tarikh bagi bulan payroll " & s & "." & vbCrLf & _
                    "Sila padam tetapan bulan payroll ini dan cuba lagi.", vbCritical, "Error"
                    
            Exit Sub
            
        End If
        
        
        Frm48.L28_Text = "Pengiraan Gaji Bagi " & Frm48_LM_Nama & " Bagi " & Frm48_LM_PAYROLL & "." 'Header
        
        Frm48.L29_Text = Format(0, "#,##0.00")
        'Frm48.L30_Text = Format(0, "#,##0.00")
        Frm48.L31_Text = Format(0, "#,##0.00")
        Frm48.L32_Text = Format(0, "#,##0.00")
        'Frm48.L33_Text = Format(0, "#,##0.00")
        Frm48.L34_Text = Format(0, "#,##0.00")
        Frm48.L36_Text = Format(0, "#,##0.00")
        Frm48.L40_Text = Format(0, "#,##0.00")
        
        Frm48.L43_Text = 0
        Frm48.L44_Text = 0
        Frm48.L45_Text = 0
        
        If Frm48.L35_Text.Visible = True Then
        
            If Frm48.CB3 = 1 Then
                
                Frm48.L43_Text = 1
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                'rs.Open "select SUM(jumlah_berat) from 22_jualan where no_pekerja='" & Frm48_LM_EMP_NO & "' AND status = 1 AND tarikh BETWEEN '" & TM & "'  AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic
                rs.Open "select SUM(jumlah_berat) from 22_jualan where (menu = 0 OR menu = 3) AND no_pekerja='" & Frm48_LM_EMP_NO & "' AND status = 1 AND tarikh BETWEEN '" & TM & "'  AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not IsNull(rs(0)) Then Frm48.L40_Text = Format(rs(0), "#,##0.00")
                
                rs.Close
                Set rs = Nothing
                
            Else
            
                Frm48.L43_Text = 0
            
            End If
            
            If Frm48.CB4 = 1 Then
                
                Frm48.L44_Text = 1
    'Jumlah jualan BK
                
                Dim LM_TOTAL_SALE_BK_1 As Double
                Dim LM_TOTAL_SALE_BK_2 As Double
            
                LM_TOTAL_SALE_BK_1 = 0
                LM_TOTAL_SALE_BK_2 = 0
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select SUM(harga_jualan_dengan_gst - jumlah_gst) from 23_senarai_jualan where type = 0 AND no_pekerja='" & Frm48_LM_EMP_NO & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND status_rekod = 1", cn, adOpenKeyset, adLockOptimistic
                
                If Not IsNull(rs(0)) Then LM_TOTAL_SALE_BK_1 = Format(rs(0), "#,##0.00")
                
                rs.Close
                Set rs = Nothing
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select SUM(harga_tanpa_gst) from 42_tempahan_siap where type_barang_kemas = 0 AND no_pekerja='" & Frm48_LM_EMP_NO & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND status_invoice = 1", cn, adOpenKeyset, adLockOptimistic
                
                If Not IsNull(rs(0)) Then LM_TOTAL_SALE_BK_2 = Format(rs(0), "#,##0.00")
                
                rs.Close
                Set rs = Nothing
                
                Frm48.L29_Text = Format(LM_TOTAL_SALE_BK_1 + LM_TOTAL_SALE_BK_2, "#,##0.00")
                
            Else
                
                Frm48.L44_Text = 0
            
            End If
            
            If Frm48.CB5 = 1 Then
    'Jumlah jualan barang permata
            
                Frm48.L45_Text = 1
                
                Dim LM_TOTAL_SALE_PERMATA_1 As Double
                Dim LM_TOTAL_SALE_PERMATA_2 As Double
                
                LM_TOTAL_SALE_PERMATA_1 = 0
                LM_TOTAL_SALE_PERMATA_2 = 0
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select SUM(harga_jualan_dengan_gst - jumlah_gst) from 23_senarai_jualan where type = 1 AND no_pekerja='" & Frm48_LM_EMP_NO & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND status_rekod = 1", cn, adOpenKeyset, adLockOptimistic
                
                If Not IsNull(rs(0)) Then LM_TOTAL_SALE_PERMATA_1 = Format(rs(0), "#,##0.00")
                
                rs.Close
                Set rs = Nothing
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select SUM(harga_tanpa_gst) from 42_tempahan_siap where type_barang_kemas = 1 AND no_pekerja='" & Frm48_LM_EMP_NO & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND status_invoice = 1", cn, adOpenKeyset, adLockOptimistic
                
                If Not IsNull(rs(0)) Then LM_TOTAL_SALE_PERMATA_2 = Format(rs(0), "#,##0.00")
                
                rs.Close
                Set rs = Nothing
                
                Frm48.L32_Text = Format(LM_TOTAL_SALE_PERMATA_1 + LM_TOTAL_SALE_PERMATA_2, "#,##0.00")
                
            Else
            
                Frm48.L45_Text = 0
                
            End If
            
        End If
        
        Frm48.L10_Text = "1"
        
    End If
End If
End Sub
Private Sub CMD5_Click()
Dim Err(10)
'On Error Resume Next
x = 0
DATA_SAVE = 0 '0 : Data Tidak Disimpan , 1 : Data Disimpan

If Frm48.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Bulan]."
End If
If Frm48.CBB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Tahun]."
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
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from tetapan_Payslip where Bulan='" & Frm48.CBB1 & "' AND Tahun='" & Frm48.CBB2 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            MsgBox "Tetapan Payslip Bagi Bulan [" & Frm48.CBB1 & "] Dan Tahun [" & Frm48.CBB2 & "] Telah Dilakukan Sebelum Ini.", vbExclamation, "Error"
        Else
            rs.AddNew
            rs!Bulan = Frm48.CBB1 'Bulan
            rs!Tahun = Frm48.CBB2 'Tahun
            rs!TarikhMula = Frm48.DTPicker1 'Tarikh Mula
            rs!TarikhAkhir = Frm48.DTPicker2 'Tarikh Akhir
            rs.Update
            DATA_SAVE = 1 '0 : Data Tidak Disimpan , 1 : Data Disimpan
        End If
        
        rs.Close
        Set rs = Nothing
    End If
    
    If DATA_SAVE = 1 Then '0 : Data Tidak Disimpan , 1 : Data Disimpan
        user = MDI_frm1.L3_Text
        LogAct_Memory = "[" & user & "] Tetapan Payroll Bulan [" & Frm48.CBB1 & "] Dan Tahun [" & Frm48.CBB2 & "]."
        LogDate_Memory = DateTime.Date & " " & DateTime.Time$
        Call UpdateLog_Database
        
        Call Frm48_Default
        Call Frm48_ListPayroll
        MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
    End If
End If
End Sub
Private Sub CMD6_Click()
'On Error Resume Next
Dim TM As Date 'Tarikh Mula
Dim TA As Date 'Tarikh Akhir
Dim JUMLAH_BERAT As Double 'Total Berat Yang Dijual Oleh Staff
Dim JUMLAH_BERAT_OVERALL As Double 'Total Berat Yang Terjual Di Kedai
Dim JUMLAH_PROFIT As Double
Dim TOTAL_SERVICE As Double 'Jumlah Kutipan Dari Servis Kepada Pelanggan
Dim TOTAL_EXPENSES As Double 'Jumlah Perbelanjaan Kedai
Dim TOTAL_GAJI As Double 'Jumlah Bayaran Gaji
Dim a As Double
Dim b As Double
Dim JUMLAH_RESIT As Integer
Dim Err(10)

x = 0
JUMLAH_RESIT = 0

If Frm48.CBB3 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Bulan Payroll]."
End If
If Frm48.CBB4 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Nama Pekerja]."
End If
If Frm48.TB9 = vbNullString Or (Frm48.TB9 <> vbNullString And Not IsNumeric(Frm48.TB9)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [KWSP] , Hanya NOMBOR Yang Dibenarkan Dalam Ruangan Ini."
End If
If Frm48.TB10 = vbNullString Or (Frm48.TB10 <> vbNullString And Not IsNumeric(Frm48.TB10)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Socso] , Hanya NOMBOR Yang Dibenarkan Dalam Ruangan Ini."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Frm48.MSFlexGrid2.Clear
    Frm48.MSFlexGrid2.RowHeight(0) = 600
    Frm48.MSFlexGrid2.FormatString = "No.|<No.|<No. Siri|<Kategori|<Tarikh Jualan|<Berat Jualan (g)|<Harga Jualan (RM)"
    
    Frm48.MSFlexGrid2.Rows = 1
    Frm48.MSFlexGrid2.ColWidth(0) = 600
    Frm48.MSFlexGrid2.ColWidth(1) = 0
    Frm48.MSFlexGrid2.ColWidth(2) = 3500
    Frm48.MSFlexGrid2.ColWidth(3) = 4500
    Frm48.MSFlexGrid2.ColWidth(4) = 2400
    Frm48.MSFlexGrid2.ColWidth(5) = 2400
    Frm48.MSFlexGrid2.ColWidth(6) = 2400
    
    PAYROLL = Frm48.CBB3 'Tetapan Bulan Payroll
    If InStr(1, PAYROLL, " ") <> 0 Then
        PAY_BULAN = Split(PAYROLL, " ")(0)
        PAY_TAHUN = Split(PAYROLL, " ")(1)
    End If
    
    If Frm48.CBB4 <> vbNullString Then
        EMP_NO_PENJUAL = Split(Frm48.CBB4, "  |  ")(1)
        Nama_PENJUAL = Split(Frm48.CBB4, "  |  ")(0)
    End If
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from tetapan_Payslip where Bulan='" & PAY_BULAN & "' AND Tahun='" & PAY_TAHUN & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!TarikhMula) Then TM = rs!TarikhMula 'Tarikh Mula
        If Not IsNull(rs!TarikhAkhir) Then TA = rs!TarikhAkhir 'Tarikh Akhir
    End If
    
    rs.Close
    Set rs = Nothing
    
    JUMLAH_BERAT = 0
    JUMLAH_PROFIT = 0
    TOTAL_SERVICE = 0
    JUMLAH_BERAT_OVERALL = 0
    TOTAL_EXPENSES = 0
    TOTAL_GAJI = 0
    
    'GoTo Skip_Kira_Gaji:
    'If TM <> vbNullString And TA <> vbNullString Then
        '///////Pengiraan Jumlah Jualan Staff Dan Jualan Kedai//////////// ####TukangemaS Tidak Gunakan Cara Ini######## START
        Frm49_LM_EMP_NO_PENJUAL = Split(Frm48.CBB4, "  |  ")(1)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        'rs.Open "select * from Maklumat_Jualan where Jurujual='" & Frm48.CBB4 & "'", cn, adOpenKeyset, adLockOptimistic
        rs.Open "select * from Maklumat_Jualan where Jurujual='" & Frm49_LM_EMP_NO_PENJUAL & "'", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
            If rs!tarikh_Jualan >= TM And rs!tarikh_Jualan <= TA Then
                If Not IsNull(rs!BeratJualan) Then
                    If IsNumeric(rs!BeratJualan) Then JUMLAH_BERAT = JUMLAH_BERAT + rs!BeratJualan
                End If
                If Not IsNull(rs!Keuntungan1) Then
                    'If IsNumeric(rs!Keuntungan1) Then JUMLAH_PROFIT = JUMLAH_PROFIT + rs!Keuntungan1
                End If
                If IsNumeric(rs!BeratJualan) Then
                    Y = Y + 1
                    Frm48.MSFlexGrid2.Rows = Y + 1
                    Frm48.MSFlexGrid2.TextMatrix(Y, 0) = Y
                    Frm48.MSFlexGrid2.TextMatrix(Y, 1) = Y
                    If Not IsNull(rs!NOSIRI) Then Frm48.MSFlexGrid2.TextMatrix(Y, 2) = rs!NOSIRI 'No. Siri Produk
                    If Not IsNull(rs!NamaProduk) Then Frm48.MSFlexGrid2.TextMatrix(Y, 3) = rs!NamaProduk 'Nama Produk
                    If Not IsNull(rs!tarikh_Jualan) Then Frm48.MSFlexGrid2.TextMatrix(Y, 4) = rs!tarikh_Jualan 'Tarikh Jualan
                    If Not IsNull(rs!BeratJualan) Then Frm48.MSFlexGrid2.TextMatrix(Y, 5) = Format(rs!BeratJualan, "0.00") 'Berat Jualan
                    If Not IsNull(rs!harga_jualan) Then Frm48.MSFlexGrid2.TextMatrix(Y, 6) = Format(rs!harga_jualan, "0.00") 'Harga Jualan
                End If
            End If
            rs.MoveNext
        Wend
            
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Maklumat_Jualan", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
            If Not IsNull(rs!BeratJualan) Then
                If IsNumeric(rs!BeratJualan) Then JUMLAH_BERAT_OVERALL = JUMLAH_BERAT_OVERALL + rs!BeratJualan
            End If
            If rs!tarikh_Jualan >= TM And rs!tarikh_Jualan <= TA Then
                If Not IsNull(rs!Keuntungan1) Then
                    If IsNumeric(rs!Keuntungan1) Then JUMLAH_PROFIT = JUMLAH_PROFIT + rs!Keuntungan1
                End If
            End If
            rs.MoveNext
        Wend
            
        rs.Close
        Set rs = Nothing
        
        '///////Pengiraan Jumlah Jualan Staff Dan Jualan Kedai//////////// ####TukangemaS Tidak Gunakan Cara Ini######## END
        
'Skip_Kira_Gaji:
        Frm49_LM_EMP_NO_PENJUAL = Split(Frm48.CBB4, "  |  ")(1)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Akaun where jurujual='" & Frm49_LM_EMP_NO_PENJUAL & "'", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
            If rs!tarikh >= TM And rs!tarikh <= TA Then
                If rs!Type = 4 Or rs!Type = 5 Or rs!Type = 8 Or rs!Type = 9 Then
                    JUMLAH_RESIT = JUMLAH_RESIT + 1
                End If
            End If
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing

        '/////////Pengiraan Servis Yang Diberikan Kepada Pelanggan Dan Perbelanjaan Kedai//////////
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from ServiceKedai", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
            If rs!JenisUrusan = 1 And rs!tarikh >= TM And rs!tarikh <= TA Then
                If Not IsNull(rs!jumlah) And IsNumeric(rs!jumlah) Then TOTAL_SERVICE = TOTAL_SERVICE + rs!jumlah 'Jumlah Servis
            End If
            If rs!JenisUrusan = 2 And rs!tarikh >= TM And rs!tarikh <= TA Then
                If Not IsNull(rs!jumlah) And IsNumeric(rs!jumlah) Then TOTAL_EXPENSES = TOTAL_EXPENSES + rs!jumlah 'Jumlah Perbelanjaan Kedai
            End If
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        '////////Pengiraan Bayaran Gaji Kepada Staff Dan Kakitangan Kedai/////////
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from payslip where payroll_bulan='" & Frm48.CBB3 & "'", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
            If Not IsNull(rs!payroll_bersih) And IsNumeric(rs!payroll_bersih) Then TOTAL_GAJI = TOTAL_GAJI + rs!payroll_bersih 'Jumlah Bayaran Gaji
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        

        If Frm48.L3_Text.Visible = True And IsNumeric(JUMLAH_BERAT_OVERALL) And IsNumeric(Frm48.L9_Text) Then
            a = Frm48.L9_Text 'Format(Frm48.L9_Text / 100, "0.00")
            Frm48.TB8 = Format(JUMLAH_BERAT_OVERALL * a, "0.00") 'Jumlah Profit
        Else
            Frm48.TB8 = Format(0, "0.00") 'Jumlah Profit
        End If
        
        If Y = vbNullString Then Y = 0
        'Frm48.TB5 = Format(JUMLAH_BERAT, "0.00") 'Jumlah Berat Terjual
        Frm48.TB5 = JUMLAH_RESIT 'Format(JUMLAH_BERAT, "0.00") 'Jumlah Berat Terjual
        Frm48.L6_Text = ":   " & Y 'Bilangan Barang Terjual
        Frm48.L7_Text = ":   " & Format(JUMLAH_BERAT, "0.00 g") 'Jumlah Berat Terjual Oleh Staff
        Frm48.L13_Text = ":   " & Format(JUMLAH_BERAT_OVERALL, "0.00 g") 'Jumlah Berat Terjual Di Kedai
        'Frm48.L14_Text = ":   RM " & Format(JUMLAH_PROFIT, "0.00") 'Jumlah Keuntungan Dari Jualan Emas
        'Frm48.L15_Text = ":   RM " & Format(TOTAL_SERVICE, "0.00") 'Jumlah Servis
        'Frm48.L16_Text = ":   RM " & Format(TOTAL_EXPENSES, "0.00") 'Jumlah Perbelanjaan Kedai
        'Frm48.L18_Text = ":   RM " & Format(TOTAL_GAJI, "0.00") 'Jumlah Pembayaran Gaji
        'Frm48.L17_Text = ":   RM " & Format(JUMLAH_PROFIT + TOTAL_SERVICE - TOTAL_EXPENSES - TOTAL_GAJI, "0.00") 'Jumlah Keuntungan Bersih
        If Frm48.L21_Text.Visible = True Then
            If IsNumeric(JUMLAH_PROFIT) And IsNumeric(Frm48.L21_Text) Then
                b = Frm48.L21_Text / 100
                Frm48.TB14 = Format(b * JUMLAH_PROFIT, "0.00") 'Komisen Bagi Investor
            End If
        Else
            Frm48.TB14 = Format(0, "0.00") 'Komisen Bagi Investor
        End If
        Frm48.L10_Text = 1
        
        MsgBox "Pengiraan Gaji Bagi Pekerja Telah Selesai.", vbInformation, "Info"
    'End If
End If
End Sub
Private Sub CMD7_Click()
'On Error Resume Next
Frm48.Pic4.Visible = True
Frm48.Pic2.Visible = False
End Sub
Private Sub CMD8_Click()
'On Error Resume Next
Frm48.Pic2.Visible = True
Frm48.Pic4.Visible = False
End Sub
Private Sub CMD9_Click()
'On Error Resume Next
Dim Err(2)

x = 0
DATA_SAVE = 0 '0 : Data Tidak Disimpan , 1 : Data Disimpan
If Frm48.L10_Text = "0" Then
    x = x + 1
    Err(x) = "Berlaku perubahan pada data / gaji pekerja ini. Sila buat pengiraan gaji sekali lagi."
End If
If Frm48.CB1 = 0 And Frm48.CB2 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan cara bayaran gaji dibuat."
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
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from payslip where payroll_bulan='" & Frm48.CBB3 & "' AND payroll_nama='" & Frm48.CBB4 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
            rs.AddNew
            If Frm48.CBB3 <> vbNullString Then 'Bulan Payroll
                rs!payroll_bulan = Frm48.CBB3
            Else
                rs!payroll_bulan = Null
            End If
            If Frm48.CBB4 <> vbNullString Then 'Nama Pekerja : Samaran
                rs!payroll_nama = Frm48.CBB4
            Else
                rs!payroll_nama = Null
            End If
            If Frm48.TB1 <> vbNullString Then 'Nama Pekerja : Penuh
                rs!payroll_namapenuh = Frm48.TB1
            Else
                rs!payroll_namapenuh = Null
            End If
            If Frm48.TB2 <> vbNullString Then 'No. Kad Pengenalan
                rs!payroll_ic = Frm48.TB2
                PAYSLIP_IC = Frm48.TB2
            Else
                rs!payroll_ic = Null
            End If
            If Frm48.TB3 <> vbNullString Then 'Gaji Pokok
                rs!payroll_gajipokok = Format(Frm48.TB3, "0.00")
            Else
                rs!payroll_gajipokok = Null
            End If
            If Frm48.TB4 <> vbNullString Then 'Elaun
                rs!payroll_elaun = Format(Frm48.TB4, "0.00")
            Else
                rs!payroll_elaun = Null
            End If
            If Frm48.TB16 <> vbNullString Then 'Overtime
                rs!overtime = Format(Frm48.TB16, "0.00")
            Else
                rs!overtime = Null
            End If
            If Frm48.TB17 <> vbNullString Then 'Elaun Perjalanan
                rs!elaun_perjalanan = Format(Frm48.TB17, "0.00")
            Else
                rs!elaun_perjalanan = Null
            End If
            If Frm48.TB18 <> vbNullString Then 'Pendapatan Lain-lain
                rs!pendapatan_lain = Format(Frm48.TB18, "0.00")
            Else
                rs!pendapatan_lain = Null
            End If
            
            If Frm48.L43_Text = "0" Then
                rs!setting_comm_berat = 0
            ElseIf Frm48.L43_Text = "1" Then
                rs!setting_comm_berat = 1
            End If
            If Frm48.L44_Text = "0" Then
                rs!setting_comm_bk = 0
            ElseIf Frm48.L44_Text = "1" Then
                rs!setting_comm_bk = 1
            End If
            If Frm48.L45_Text = "0" Then
                rs!setting_comm_permata = 0
            ElseIf Frm48.L45_Text = "1" Then
                rs!setting_comm_permata = 1
            End If
            
            If Frm48.CB3 = 1 Then
                
                rs!jenis_komisen = 0
                If Frm48.L40_Text <> vbNullString Then 'Berat jualan terkumpul
                    rs!berat_jualan = Format(Frm48.L40_Text, "0.00")
                Else
                    rs!berat_jualan = Null
                End If
                If Frm48.L41_Text <> vbNullString Then 'Rate Komisen Per Gram
                    rs!payroll_komisenrate = Format(Frm48.L41_Text, "0.00")
                Else
                    rs!payroll_komisenrate = Null
                End If
            
            Else
                
                rs!berat_jualan = Null
                rs!payroll_komisenrate = Null
                
            End If
            If Frm48.CB4 = 1 Then
                
                rs!jenis_komisen = 1
                If Frm48.L29_Text <> vbNullString Then 'Jumlah Jualan Barang Kemas (RM)
                    rs!payroll_jualan_bk = Format(Frm48.L29_Text, "0.00")
                Else
                    rs!payroll_jualan_bk = Null
                End If
                If Frm48.L30_Text <> vbNullString Then 'Jumlah Rate Komisen Barang Kemas (%)
                    rs!payroll_bk_komisen_rate = Format(Frm48.L30_Text, "0.00")
                Else
                    rs!payroll_bk_komisen_rate = Null
                End If
                If Frm48.L31_Text <> vbNullString Then 'Jumlah Komisen Jualan Barang Kemas (RM)
                    rs!payroll_komisen_bk = Format(Frm48.L31_Text, "0.00")
                Else
                    rs!payroll_komisen_bk = Null
                End If
                If Frm48.L32_Text <> vbNullString Then 'Jumlah Jualan Barang Permata (RM)
                    rs!payroll_jualan_permata = Format(Frm48.L32_Text, "0.00")
                Else
                    rs!payroll_jualan_permata = Null
                End If
                If Frm48.L33_Text <> vbNullString Then 'Jumlah Rate Komisen Barang Permata (%)
                    rs!payroll_permata_komisen_rate = Format(Frm48.L33_Text, "0.00")
                Else
                    rs!payroll_permata_komisen_rate = Null
                End If
                If Frm48.L34_Text <> vbNullString Then 'Jumlah Komisen Jualan Barang Permata (RM)
                    rs!payroll_komisen_permata = Format(Frm48.L34_Text, "0.00")
                Else
                    rs!payroll_komisen_permata = Null
                End If
                
            Else
            
                rs!payroll_jualan_bk = Null
                rs!payroll_bk_komisen_rate = Null
                rs!payroll_komisen_bk = Null
                rs!payroll_jualan_permata = Null
                rs!payroll_permata_komisen_rate = Null
                rs!payroll_komisen_permata = Null
                
            End If
            If Frm48.L36_Text <> vbNullString Then 'Jumlah Keseluruhan Komisen (RM)
                rs!payroll_jumlah_komisen = Format(Frm48.L36_Text, "0.00")
            Else
                rs!payroll_jumlah_komisen = Null
            End If

            If Frm48.TB9 <> vbNullString Then 'KWSP
                rs!payroll_kwsp = Format(Frm48.TB9, "0.00")
            Else
                rs!payroll_kwsp = Null
            End If
            If Frm48.TB10 <> vbNullString Then 'Socso
                rs!payroll_socso = Format(Frm48.TB10, "0.00")
            Else
                rs!payroll_socso = Null
            End If
            If Frm48.TB15 <> vbNullString Then 'Potongan lain-lain
                rs!payroll_lain = Format(Frm48.TB15, "0.00")
            Else
                rs!payroll_lain = Null
            End If
            If Frm48.TB19 <> vbNullString Then 'Potongan zakat
                rs!zakat = Format(Frm48.TB19, "0.00")
            Else
                rs!zakat = Null
            End If
            If Frm48.TB20 <> vbNullString Then 'Potongan income tax
                rs!tax = Format(Frm48.TB20, "0.00")
            Else
                rs!tax = Null
            End If
            If Frm48.TB21 <> vbNullString Then 'Potongan advance
                rs!advance = Format(Frm48.TB21, "0.00")
            Else
                rs!advance = Null
            End If
            If Frm48.TB11 <> vbNullString Then 'Pendapatan Kasar
                rs!payroll_kasar = Format(Frm48.TB11, "0.00")
            Else
                rs!payroll_kasar = Null
            End If
            If Frm48.TB12 <> vbNullString Then 'Penolakan
                rs!payroll_tolak = Format(Frm48.TB12, "0.00")
            Else
                rs!payroll_tolak = Null
            End If
            If Frm48.TB13 <> vbNullString Then 'Jumlah Pendapatan Bersih
                rs!payroll_bersih = Format(Frm48.TB13, "0.00")
            Else
                rs!payroll_bersih = Null
            End If
            rs!tarikh = Frm48.DTPicker5 'Tarikh bayaran dibuat
            If Frm48.CB1 = 1 Then 'Cara bayaran dibuat
                If Frm48.TB13 <> vbNullString Then
                    rs!tunai = Format(Frm48.TB13, "0.00")
                Else
                    rs!tunai = "0.00"
                End If
                rs!bank_in = "0.00"
            End If
            If Frm48.CB2 = 1 Then 'Cara bayaran dibuat
                If Frm48.TB13 <> vbNullString Then
                    rs!bank_in = Format(Frm48.TB13, "0.00")
                Else
                    rs!bank_in = "0.00"
                End If
                rs!tunai = "0.00"
            End If
            
            rs.Update
            DATA_SAVE = 1 '0 : Data Tidak Disimpan , 1 : Data Disimpan
        Else
            MsgBox "Rekod Payslip Telah Disimpan Sebelum Ini , Jika Anda Ingin Simpan Rekod Payslip Yang Baru Perlu Padam Dahulu Rekod Yang Lama", vbExclamation, "Info"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_SAVE = 1 Then '0 : Data Tidak Disimpan , 1 : Data Disimpan
        
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & user & "] Simpan Rekod Payslip [" & Frm48.CBB3 & "] , Nama [" & Frm48.CBB4 & "]."
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
            
            PAYSLIP_BULAN = Frm48.CBB3
            PAYSLIP_NAMA = Frm48.CBB4
            
            Call frm48_reset_gaji
            Call Frm48_CalcDefault
            
            G_PAYSLIP_BULAN = vbNullString
            G_PAYSLIP_IC = vbNullString
            
            G_PAYSLIP_BULAN = PAYSLIP_BULAN
            G_PAYSLIP_IC = PAYSLIP_IC
            
            If G_PAYSLIP_BULAN <> vbNullString And G_PAYSLIP_IC <> vbNullString Then
                Call Frm48_M_cetak_payslip
            End If

            MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
        End If
    End If
End If
End Sub
Private Sub Form_Load()
'On Error Resume Next
Frm48.DTPicker5 = DateTime.Date
End Sub
Private Sub Frm48_CetakPayslip_Click()
'On Error Resume Next
G_PAYSLIP_BULAN = vbNullString
G_PAYSLIP_IC = vbNullString

G_PAYSLIP_BULAN = Frm48.MSFlexGrid3.TextMatrix(Frm48.MSFlexGrid3, 2)
G_PAYSLIP_IC = Frm48.MSFlexGrid3.TextMatrix(Frm48.MSFlexGrid3, 4)

If G_PAYSLIP_BULAN <> vbNullString And G_PAYSLIP_IC <> vbNullString Then
    Call Frm48_M_cetak_payslip
End If
End Sub
Private Sub Frm48_PadamData_Click()
'On Error Resume Next
DATA_DELETE = 0 '0 : Data Tidak Dipadamkan , 1 : Data Dipadam
PAYSLIP_BULAN = Frm48.MSFlexGrid3.TextMatrix(Frm48.MSFlexGrid3, 2)
PAYSLIP_NAMA = Frm48.MSFlexGrid3.TextMatrix(Frm48.MSFlexGrid3, 3)
PAYSLIP_IC = Frm48.MSFlexGrid3.TextMatrix(Frm48.MSFlexGrid3, 4)

If PAYSLIP_BULAN <> vbNullString And PAYSLIP_IC <> vbNullString Then
    Note = "Adakah Anda Ingin Padam Data Ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from payslip where payroll_bulan='" & PAYSLIP_BULAN & "' AND payroll_ic='" & PAYSLIP_IC & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            rs.Delete
            rs.Update
            DATA_DELETE = 1 '0 : Data Tidak Dipadamkan , 1 : Data Dipadam
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_DELETE = 1 Then '0 : Data Tidak Dipadamkan , 1 : Data Dipadam
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & user & "] Padam Rekod Payslip [" & PAYSLIP_BULAN & "] , Nama [" & PAYSLIP_NAMA & "]."
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
        
            Call Frm48_RekodPayslip
            MsgBox "Data Telah Berjaya Dipadamkan.", vbInformation, "Info"
        End If
    End If
End If
End Sub
Private Sub Frm48_PadamTetapan_Click()
'On Error Resume Next
DATA_DELETE = 0 '0 : Data Tidak Dipadamkan , 1 : Data Dipadamkan
TETAPAN_BULAN = Frm48.MSFlexGrid1.TextMatrix(Frm48.MSFlexGrid1, 2)
TETAPAN_TAHUN = Frm48.MSFlexGrid1.TextMatrix(Frm48.MSFlexGrid1, 3)

If TETAPAN_BULAN <> vbNullString And TETAPAN_TAHUN <> vbNullString Then
    Note = "Adakah Anda Ingin Padam Data Ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from tetapan_Payslip where Bulan='" & TETAPAN_BULAN & "' AND Tahun='" & TETAPAN_TAHUN & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            rs.Delete
            rs.Update
            DATA_DELETE = 1 '0 : Data Tidak Dipadamkan , 1 : Data Dipadamkan
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_DELETE = 1 Then '0 : Data Tidak Dipadamkan , 1 : Data Dipadamkan
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & user & "] Padam Tetapan Payroll Bulan [" & TETAPAN_BULAN & "] Dan Tahun [" & TETAPAN_TAHUN & "]."
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
        
            Call Frm48_ListPayroll
            MsgBox "Data Telah Berjaya Dipadamkan.", vbInformation, "Info"
        End If
    End If
End If
End Sub

Private Sub frm48_sm_excel_Click()
'On Error Resume Next
Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem mungkin akan mengambil sedikit masa untuk mengeluarkan report ini. Teruskan ?"
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
        .Columns("B").ColumnWidth = 15 'Bulan Payroll
        .Columns("C").ColumnWidth = 60 'Nama
        .Columns("D").ColumnWidth = 25 'No. Kad Pengenalan
        .Columns("E").ColumnWidth = 15 'Gaji Pokok (RM)
        .Columns("F").ColumnWidth = 15 'Elaun (RM)
        .Columns("G").ColumnWidth = 15 'Overtime (RM)
        .Columns("H").ColumnWidth = 15 'Elaun Perjalanan (RM)
        .Columns("I").ColumnWidth = 15 'Lain-lain (RM)
        .Columns("J").ColumnWidth = 15 'Jumlah Komisen (RM)
        .Columns("K").ColumnWidth = 15 'KWSP (RM)
        .Columns("L").ColumnWidth = 15 'Socso (RM)
        .Columns("M").ColumnWidth = 15 'Lain-lain (RM)
        .Columns("N").ColumnWidth = 15 'Zakat (RM)
        .Columns("O").ColumnWidth = 15 'Income Tax (RM)
        .Columns("P").ColumnWidth = 15 'Advance (RM)
        .Columns("Q").ColumnWidth = 15 'Pendapatan Kasar (RM)
        .Columns("R").ColumnWidth = 15 'Jumlah Penolakan (RM)
        .Columns("S").ColumnWidth = 15 'Pendapatan Bersih (RM)
        .Columns("T").ColumnWidth = 10
        .Columns("U").ColumnWidth = 10

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
                .Cells(1, 6) = rs!nama_kedai
                .Cells(1, 6).Font.Name = "Times New Roman"
            End If
            If Not IsNull(rs!no_pendaftaran) Then .Cells(2, 6) = rs!no_pendaftaran
            If Not IsNull(rs!alamat) Then .Cells(3, 6) = rs!alamat
            If Not IsNull(rs!no_tel) Then .Cells(4, 6) = rs!no_tel
            If Not IsNull(rs!no_id_gst) Then .Cells(5, 6) = rs!no_id_gst
        End If
        
        rs.Close
        Set rs = Nothing
        '### Maklumat kedai ### - End
        
        .Cells(1, 6).Font.Bold = True
        .Cells(1, 6).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 6).HorizontalAlignment = xlCenter
        Next Row
        
        .Cells(7, 1) = Frm48.Label49 'Header Report
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Bulan Payroll"
        .Cells(8, 3) = "Nama"
        .Cells(8, 4) = "No. Kad Pengenalan"
        .Cells(8, 5) = "Gaji Pokok (RM)"
        .Cells(8, 6) = "Elaun (RM)"
        .Cells(8, 7) = "Overtime (RM)"
        .Cells(8, 8) = "Elaun Perjalanan (RM)"
        .Cells(8, 9) = "Lain-lain (RM)"
        .Cells(8, 10) = "Jumlah Komisen (RM)"
        .Cells(8, 11) = "KWSP (RM)"
        .Cells(8, 12) = "Socso (RM)"
        .Cells(8, 13) = "Lain-lain (RM)"
        .Cells(8, 14) = "Zakat (RM)"
        .Cells(8, 15) = "Income Tax (RM)"
        .Cells(8, 16) = "Advance (RM)"
        .Cells(8, 17) = "Pendapatan Kasar (RM)"
        .Cells(8, 18) = "Jumlah Penolakan (RM)"
        .Cells(8, 19) = "Pendapatan Bersih (RM)"
    
        For i = 1 To 19
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
        
        
        Y = 0
        For x = 1 To Frm48.MSFlexGrid3.Rows - 1
            Y = Y + 1
            .Cells(8 + Y, 1) = Y 'No.
            .Cells(8 + Y, 1).HorizontalAlignment = xlCenter
            
            .Cells(8 + Y, 2) = "'" & Frm48.MSFlexGrid3.TextMatrix(x, 2) 'Bulan Payroll
            
            .Cells(8 + Y, 3) = Frm48.MSFlexGrid3.TextMatrix(x, 3) 'Nama

            .Cells(8 + Y, 4) = "'" & Frm48.MSFlexGrid3.TextMatrix(x, 4) 'No. Kad Pengenalan

            .Cells(8 + Y, 5) = Frm48.MSFlexGrid3.TextMatrix(x, 5) 'Gaji Pokok (RM)
            .Cells(8 + Y, 5).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 5).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 6) = Frm48.MSFlexGrid3.TextMatrix(x, 6) 'Elaun (RM)
            .Cells(8 + Y, 6).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 6).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 7) = Frm48.MSFlexGrid3.TextMatrix(x, 7) 'Overtime (RM)
            .Cells(8 + Y, 7).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 7).HorizontalAlignment = xlRight

            .Cells(8 + Y, 8) = Frm48.MSFlexGrid3.TextMatrix(x, 8) 'Elaun Perjalanan (RM)
            .Cells(8 + Y, 8).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 8).HorizontalAlignment = xlRight

            .Cells(8 + Y, 9) = Frm48.MSFlexGrid3.TextMatrix(x, 9) 'Lain-lain (RM)
            .Cells(8 + Y, 9).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 9).HorizontalAlignment = xlRight

            .Cells(8 + Y, 10) = Frm48.MSFlexGrid3.TextMatrix(x, 10) 'Jumlah Komisen (RM)
            .Cells(8 + Y, 10).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 10).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 11) = Frm48.MSFlexGrid3.TextMatrix(x, 11) 'KWSP (RM)
            .Cells(8 + Y, 11).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 11).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 12) = Frm48.MSFlexGrid3.TextMatrix(x, 12) 'Socso (RM)
            .Cells(8 + Y, 12).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 12).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 13) = Frm48.MSFlexGrid3.TextMatrix(x, 13) 'Lain-lain (RM)
            .Cells(8 + Y, 13).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 13).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 14) = Frm48.MSFlexGrid3.TextMatrix(x, 14) 'Zakat (RM)
            .Cells(8 + Y, 14).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 14).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 15) = Frm48.MSFlexGrid3.TextMatrix(x, 15) 'Income Tax (RM)
            .Cells(8 + Y, 15).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 15).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 16) = Frm48.MSFlexGrid3.TextMatrix(x, 16) 'Advance (RM)
            .Cells(8 + Y, 16).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 16).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 17) = Frm48.MSFlexGrid3.TextMatrix(x, 17) 'Pendapatan Kasar (RM)
            .Cells(8 + Y, 17).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 17).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 18) = Frm48.MSFlexGrid3.TextMatrix(x, 18) 'Jumlah Penolakan (RM)
            .Cells(8 + Y, 18).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 18).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 19) = Frm48.MSFlexGrid3.TextMatrix(x, 19) 'Pendapatan Bersih (RM)
            .Cells(8 + Y, 19).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 19).HorizontalAlignment = xlRight
            
            For Col = 1 To 19
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

Private Sub L29_Text_Change()
'On Error Resume Next
Call frm48_kiraan_komisen_bk
End Sub
Private Sub L30_Text_Change()
'On Error Resume Next
Call frm48_kiraan_komisen_bk
End Sub
Private Sub L31_Text_Change()
'On Error Resume Next
Call frm48_kiraan_komisen
End Sub
Private Sub L32_Text_Change()
'On Error Resume Next
Call frm48_kiraan_komisen_permata
End Sub
Private Sub L33_Text_Change()
'On Error Resume Next
Call frm48_kiraan_komisen_permata
End Sub
Private Sub L34_Text_Change()
'On Error Resume Next
Call frm48_kiraan_komisen
End Sub
Private Sub L36_Text_Change()
'On Error Resume Next
Call frm48_kiraan_pendapatan
End Sub
Private Sub L37_Text_Click()
'On Error Resume Next
If Frm48.Pic1.Visible = False Then
    Call Frm48_Default
    Call Frm48_ListPayroll
    Call frm48_pic_enable
    
    Frm48.Pic1.Visible = True

Else
    Frm48.Pic1.Visible = False
End If
End Sub
Private Sub L38_Text_Click()
'On Error Resume Next
If Frm48.Pic2.Visible = False Then
    
    If MDI_frm1.L20_Text = "Semua cawangan" Then
    
        Frm96.CMD2.Visible = True
        Frm96.CMD1.Visible = False
    
        Call Frm96_initial
        
        Frm96.Show vbModal
        
    End If

    Call frm48_reset_gaji
    Call Frm48_CalcDefault
    Call frm48_pic_enable
    
    Frm48.L6_Text = 0
    Frm48.L7_Text = "0.00 g"
    Frm48.L8_Text = "RM 0.00"
    Frm48.L28_Text = vbNullString
    
    Frm48.Pic2.Visible = True
Else
    Frm48.Pic2.Visible = False
End If
End Sub
Private Sub L39_Text_Click()
'On Error Resume Next
If Frm48.Pic3.Visible = False Then
    
    Call frm48_pic_enable
    Call Frm48_RekodPayslip

    Frm48.Pic3.Visible = True
Else
    Frm48.Pic3.Visible = False
End If
End Sub
Private Sub L40_Text_Change()
'On Error Resume Next
Call frm48_kiraan_komisen_berat
End Sub
Private Sub L41_Text_Change()
'On Error Resume Next
Call frm48_kiraan_komisen_berat
End Sub

Private Sub L42_Text_Change()
'On Error Resume Next
Call frm48_kiraan_komisen
End Sub

Private Sub MSFlexGrid1_DblClick()
'On Error Resume Next
PopupMenu Frm_Menu1
End Sub
Private Sub MSFlexGrid3_DblClick()
'On Error Resume Next
PopupMenu Frm48_Menu2
End Sub

Private Sub TB10_Change()
'On Error Resume Next
Call frm48_kiraan_tolakan
Frm48.L10_Text = "0"
End Sub
Private Sub TB11_Change()
'On Error Resume Next
Call frm48_kiraan_bersih
End Sub
Private Sub TB12_Change()
'On Error Resume Next
Call frm48_kiraan_bersih
End Sub

Private Sub TB15_Change()
'On Error Resume Next
Call frm48_kiraan_tolakan
Frm48.L10_Text = "0"
End Sub
Private Sub TB16_Change()
'On Error Resume Next
Call frm48_kiraan_pendapatan
Frm48.L10_Text = "0"
End Sub
Private Sub TB17_Change()
'On Error Resume Next
Call frm48_kiraan_pendapatan
Frm48.L10_Text = "0"
End Sub
Private Sub TB18_Change()
'On Error Resume Next
Call frm48_kiraan_pendapatan
Frm48.L10_Text = "0"
End Sub

Private Sub TB19_Change()
'On Error Resume Next
Call frm48_kiraan_tolakan
Frm48.L10_Text = "0"
End Sub

Private Sub TB20_Change()
'On Error Resume Next
Call frm48_kiraan_tolakan
Frm48.L10_Text = "0"
End Sub

Private Sub TB21_Change()
'On Error Resume Next
Call frm48_kiraan_tolakan
Frm48.L10_Text = "0"
End Sub

Private Sub TB3_Change()
'On Error Resume Next
Call frm48_kiraan_pendapatan
End Sub
Private Sub TB4_Change()
'On Error Resume Next
Call frm48_kiraan_pendapatan
End Sub
Private Sub TB5_Change()
'On Error Resume Next
Dim a As Double
Dim b As Double

If Frm48.TB5 <> vbNullString And IsNumeric(Frm48.TB5) And Frm48.TB6 <> vbNullString And IsNumeric(Frm48.TB6) Then
    a = Frm48.TB5
    b = Frm48.TB6
    
    Frm48.TB7 = Format(a * b, "0.00") 'Komisen
Else
    Frm48.TB7 = "0.00"
End If
End Sub
Private Sub TB6_Change()
'On Error Resume Next
Dim a As Double
Dim b As Double

If Frm48.TB5 <> vbNullString And IsNumeric(Frm48.TB5) And Frm48.TB6 <> vbNullString And IsNumeric(Frm48.TB6) Then
    a = Frm48.TB5
    b = Frm48.TB6
    
    Frm48.TB7 = Format(a * b, "0.00") 'Komisen
Else
    Frm48.TB7 = "0.00"
End If
End Sub
Private Sub TB9_Change()
'On Error Resume Next
Call frm48_kiraan_tolakan
Frm48.L10_Text = "0"
End Sub
