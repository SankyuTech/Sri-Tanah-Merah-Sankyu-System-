VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm49 
   Caption         =   "Data Pekerja"
   ClientHeight    =   12915
   ClientLeft      =   120
   ClientTop       =   450
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
   Icon            =   "Frm49.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Maklumat Pekerja"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11175
      Left            =   3120
      TabIndex        =   49
      Top             =   360
      Width           =   19095
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cawangan"
         Height          =   1335
         Left            =   10200
         TabIndex        =   131
         Top             =   8400
         Width           =   8775
         Begin VB.ComboBox CBB2 
            BackColor       =   &H00FFFFFF&
            DataField       =   "Supplier"
            Height          =   360
            ItemData        =   "Frm49.frx":0ECA
            Left            =   3120
            List            =   "Frm49.frx":0ECC
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   840
            Width           =   5415
         End
         Begin VB.Label Label60 
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3000
            TabIndex        =   134
            Top             =   840
            Width           =   135
         End
         Begin VB.Label Label59 
            BackStyle       =   0  'Transparent
            Caption         =   "Cawangan"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   240
            TabIndex        =   133
            Top             =   840
            Width           =   2595
         End
         Begin VB.Label Label67 
            BackStyle       =   0  'Transparent
            Caption         =   "Cawangan perlu dipilih bagi pekerja dengan level/role ""Manager"" dan ""Staff"" sahaja. ""Admin"" pula akan didaftarkan di bawah HQ."
            ForeColor       =   &H00008080&
            Height          =   540
            Left            =   240
            TabIndex        =   132
            Top             =   240
            Width           =   8355
         End
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
         Left            =   8160
         MouseIcon       =   "Frm49.frx":0ECE
         MousePointer    =   99  'Custom
         Picture         =   "Frm49.frx":11D8
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   9840
         Width           =   2775
      End
      Begin VB.CommandButton CMD6 
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
         Left            =   6720
         MouseIcon       =   "Frm49.frx":37A2
         MousePointer    =   99  'Custom
         Picture         =   "Frm49.frx":3AAC
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   9840
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CommandButton CMD7 
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
         Left            =   9600
         MouseIcon       =   "Frm49.frx":6076
         MousePointer    =   99  'Custom
         Picture         =   "Frm49.frx":6380
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   9840
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox TB9 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   13200
         TabIndex        =   20
         Text            =   "TB9"
         Top             =   7110
         Width           =   5775
      End
      Begin VB.TextBox TB10 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   13200
         TabIndex        =   21
         Text            =   "TB10"
         Top             =   7470
         Width           =   5775
      End
      Begin VB.CheckBox CB4 
         BackColor       =   &H8000000C&
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
         Left            =   10320
         TabIndex        =   22
         Top             =   7920
         Width           =   200
      End
      Begin VB.CheckBox CB1 
         BackColor       =   &H8000000C&
         Height          =   200
         Left            =   13035
         TabIndex        =   17
         Top             =   5310
         Width           =   200
      End
      Begin VB.CheckBox CB2 
         BackColor       =   &H8000000C&
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
         Left            =   13995
         TabIndex        =   18
         Top             =   5310
         Width           =   200
      End
      Begin VB.CheckBox CB3 
         BackColor       =   &H8000000C&
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
         Left            =   15195
         TabIndex        =   19
         Top             =   5310
         Width           =   200
      End
      Begin VB.TextBox TB13 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   12960
         TabIndex        =   15
         Text            =   "TB13"
         Top             =   3375
         Width           =   5775
      End
      Begin VB.TextBox TB16 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   12960
         TabIndex        =   16
         Text            =   "TB16"
         Top             =   3735
         Width           =   5775
      End
      Begin VB.TextBox TB12 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   12960
         TabIndex        =   14
         Text            =   "TB12"
         Top             =   1800
         Width           =   5775
      End
      Begin VB.ComboBox CBB1 
         Height          =   360
         Left            =   12960
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   720
         Width           =   5775
      End
      Begin VB.TextBox TB15 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2475
         TabIndex        =   11
         Text            =   "TB15"
         Top             =   7215
         Width           =   4000
      End
      Begin VB.TextBox TB14 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2475
         TabIndex        =   10
         Text            =   "TB14"
         Top             =   6840
         Width           =   4000
      End
      Begin VB.TextBox TB6 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2475
         TabIndex        =   6
         Text            =   "TB6"
         Top             =   2520
         Width           =   4000
      End
      Begin VB.TextBox TB5 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2475
         TabIndex        =   5
         Text            =   "TB5"
         Top             =   2160
         Width           =   4000
      End
      Begin VB.TextBox TB4 
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   2475
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "TB4"
         Top             =   1800
         Width           =   4000
      End
      Begin VB.TextBox TB8 
         BackColor       =   &H00FFFFFF&
         Height          =   1740
         Left            =   2475
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "Frm49.frx":894A
         Top             =   3615
         Width           =   7500
      End
      Begin VB.TextBox TB7 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2475
         TabIndex        =   7
         Text            =   "TB7"
         Top             =   2880
         Width           =   4000
      End
      Begin VB.TextBox TB3 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2475
         TabIndex        =   3
         Text            =   "TB3"
         Top             =   1440
         Width           =   4000
      End
      Begin VB.TextBox TB2 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2475
         TabIndex        =   2
         Text            =   "TB2"
         Top             =   1080
         Width           =   4000
      End
      Begin VB.TextBox TB1 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2475
         TabIndex        =   1
         Text            =   "TB1"
         Top             =   720
         Width           =   7500
      End
      Begin VB.TextBox TB19 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2475
         TabIndex        =   8
         Text            =   "TB19"
         Top             =   3240
         Width           =   4000
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   12960
         TabIndex        =   135
         Top             =   1080
         Width           =   5775
         _ExtentX        =   10186
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
         Format          =   415432704
         CurrentDate     =   41561
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   360
         Left            =   12960
         TabIndex        =   136
         Top             =   1440
         Width           =   5775
         _ExtentX        =   10186
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
         Format          =   415432704
         CurrentDate     =   41561
      End
      Begin VB.Label Label66 
         BackStyle       =   0  'Transparent
         Caption         =   "** USERNAME tidak boleh diubah apabila sudah didaftarkan."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   10200
         TabIndex        =   130
         Top             =   3120
         Width           =   6675
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   13080
         TabIndex        =   102
         Top             =   7515
         Width           =   135
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Akaun"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10320
         TabIndex        =   101
         Top             =   7515
         Width           =   2595
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   13080
         TabIndex        =   100
         Top             =   7155
         Width           =   135
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Bank"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10320
         TabIndex        =   99
         Top             =   7155
         Width           =   2595
      End
      Begin VB.Label Label51 
         BackStyle       =   0  'Transparent
         Caption         =   "Komisen Jualan"
         Height          =   300
         Left            =   10590
         TabIndex        =   98
         Top             =   7890
         Width           =   1695
      End
      Begin VB.Label Label52 
         BackStyle       =   0  'Transparent
         Caption         =   "(Sila tanda di sini jika pekerja ini layak untuk mendapat komisen bagi setiap jualan yang dilakukan)"
         Height          =   540
         Left            =   12075
         TabIndex        =   97
         Top             =   7920
         Width           =   6855
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "ADMIN : Boleh masuk ke dalam semua menu yang ada di dalam sistem    Hanya User dengan level HQ layak mendaftar ADMIN."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   540
         Left            =   10320
         TabIndex        =   96
         Top             =   5520
         Width           =   6675
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "** Nama samaran && Password ini akan digunakan bagi pekerja ini untuk login ke dalam sistem."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   555
         Left            =   10275
         TabIndex        =   95
         Top             =   4200
         Width           =   6675
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "** Pihak ADMIN perlu menyediakan password (initial password) kepada pekerja bagi pendaftaran baru."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   555
         Left            =   10275
         TabIndex        =   94
         Top             =   4680
         Width           =   6675
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "User Level *"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10200
         TabIndex        =   93
         Top             =   5265
         Width           =   2595
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "Admin       Manager        Staff"
         Height          =   300
         Left            =   13275
         TabIndex        =   92
         Top             =   5265
         Width           =   5055
      End
      Begin VB.Label Label53 
         BackStyle       =   0  'Transparent
         Caption         =   "MANAGER : Tidak boleh padam data , boleh buat tetapan harga jualan tetapi tidak boleh masuk ke menu ADMIN."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   540
         Left            =   10275
         TabIndex        =   91
         Top             =   6000
         Width           =   6675
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm49.frx":894E
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   780
         Left            =   10275
         TabIndex        =   90
         Top             =   6480
         Width           =   6675
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "** Sila masukkan nama samaran anda , nama samaran ini akan digunakan di dalam sistem dalam apa jua urusan kedai."
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   10200
         TabIndex        =   89
         Top             =   2520
         Width           =   6675
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Samaran / Username *"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10080
         TabIndex        =   88
         Top             =   3405
         Width           =   2835
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   12840
         TabIndex        =   87
         Top             =   3405
         Width           =   135
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   12840
         TabIndex        =   86
         Top             =   3765
         Width           =   135
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Password *"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10080
         TabIndex        =   85
         Top             =   3765
         Width           =   2595
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   12840
         TabIndex        =   84
         Top             =   1845
         Width           =   135
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Jawatan *"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10080
         TabIndex        =   83
         Top             =   1845
         Width           =   1155
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   12840
         TabIndex        =   82
         Top             =   780
         Width           =   135
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Status *"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10080
         TabIndex        =   81
         Top             =   780
         Width           =   1155
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Mula Kerja"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10080
         TabIndex        =   80
         Top             =   1125
         Width           =   1995
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   12840
         TabIndex        =   79
         Top             =   1125
         Width           =   135
      End
      Begin VB.Label L1_Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Berhenti Kerja"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10080
         TabIndex        =   78
         Top             =   1470
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.Label L2_Label 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   12840
         TabIndex        =   77
         Top             =   1470
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L4_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L4_Text"
         Height          =   255
         Left            =   480
         TabIndex        =   76
         Top             =   9000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2370
         TabIndex        =   75
         Top             =   7245
         Width           =   135
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Elaun *                   RM"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   74
         Top             =   7245
         Width           =   2115
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2370
         TabIndex        =   73
         Top             =   6870
         Width           =   135
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Gaji *                     RM"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   72
         Top             =   6870
         Width           =   2115
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Maklumat Gaji"
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
         TabIndex        =   71
         Top             =   6360
         Width           =   6255
      End
      Begin VB.Label Label54 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Maklumat Pekerja"
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
         TabIndex        =   70
         Top             =   240
         Width           =   18495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Kad Pengenalan *"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   69
         Top             =   1125
         Width           =   2235
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Income Tax"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   68
         Top             =   2550
         Width           =   1995
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2370
         TabIndex        =   67
         Top             =   2550
         Width           =   135
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "No. EPF"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   66
         Top             =   2205
         Width           =   1995
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2370
         TabIndex        =   65
         Top             =   2205
         Width           =   135
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Pekerja *"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   64
         Top             =   1845
         Width           =   1995
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2370
         TabIndex        =   63
         Top             =   1845
         Width           =   135
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "(Sila masukkan tanpa ""-"")"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6600
         TabIndex        =   62
         Top             =   1200
         Width           =   3195
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2370
         TabIndex        =   61
         Top             =   3600
         Width           =   135
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2370
         TabIndex        =   60
         Top             =   2910
         Width           =   135
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2370
         TabIndex        =   59
         Top             =   1470
         Width           =   135
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Telefon"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   58
         Top             =   2910
         Width           =   1995
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   57
         Top             =   3600
         Width           =   1995
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Passport"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   56
         Top             =   1470
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2370
         TabIndex        =   55
         Top             =   1125
         Width           =   135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2370
         TabIndex        =   54
         Top             =   750
         Width           =   135
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pekerja *"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   53
         Top             =   750
         Width           =   1995
      End
      Begin VB.Label Label56 
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail *"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   52
         Top             =   3270
         Width           =   1995
      End
      Begin VB.Label Label57 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2370
         TabIndex        =   51
         Top             =   3270
         Width           =   135
      End
      Begin VB.Label Label58 
         BackStyle       =   0  'Transparent
         Caption         =   "** Username dan password akan dihantar ke email ini jika user terlupa username atau password."
         ForeColor       =   &H000000FF&
         Height          =   795
         Left            =   6480
         TabIndex        =   50
         Top             =   2850
         Width           =   3195
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
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
      Height          =   3255
      Left            =   1800
      TabIndex        =   105
      Top             =   1680
      Width           =   10575
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Carian mengikut cawangan"
         Height          =   1095
         Left            =   120
         TabIndex        =   114
         Top             =   1920
         Width           =   10215
         Begin VB.CommandButton CMD2 
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
            Picture         =   "Frm49.frx":89F3
            Style           =   1  'Graphical
            TabIndex        =   118
            ToolTipText     =   "Carian Maklumat Pembeli"
            Top             =   240
            Width           =   2145
         End
         Begin VB.ComboBox CBB3 
            BackColor       =   &H00FFFFFF&
            DataField       =   "Supplier"
            Height          =   360
            ItemData        =   "Frm49.frx":939D
            Left            =   2715
            List            =   "Frm49.frx":939F
            Style           =   2  'Dropdown List
            TabIndex        =   115
            Top             =   480
            Width           =   5055
         End
         Begin VB.Label Label63 
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2595
            TabIndex        =   117
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label62 
            BackStyle       =   0  'Transparent
            Caption         =   "Cawangan"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   240
            TabIndex        =   116
            Top             =   480
            Width           =   2595
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Carian"
         Height          =   1455
         Left            =   120
         TabIndex        =   106
         Top             =   360
         Width           =   10215
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
            Picture         =   "Frm49.frx":93A1
            Style           =   1  'Graphical
            TabIndex        =   112
            ToolTipText     =   "Carian Maklumat Pembeli"
            Top             =   480
            Width           =   2145
         End
         Begin VB.TextBox TB20 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   2715
            TabIndex        =   110
            Top             =   840
            Width           =   5100
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No. Pekerja"
            Height          =   375
            Left            =   3480
            TabIndex        =   109
            Top             =   360
            Width           =   2535
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No. Kad Pengenalan"
            Height          =   375
            Left            =   1200
            TabIndex        =   108
            Top             =   360
            Width           =   2175
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nama"
            Height          =   375
            Left            =   240
            TabIndex        =   107
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2600
            TabIndex        =   113
            Top             =   840
            Width           =   135
         End
         Begin VB.Label L5_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Pekerja *"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   480
            TabIndex        =   111
            Top             =   870
            Width           =   1995
         End
      End
      Begin VB.Label L7_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L7_Text"
         Height          =   255
         Left            =   1080
         TabIndex        =   120
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L6_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L6_Text"
         Height          =   255
         Left            =   0
         TabIndex        =   119
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Senarai Pekerja"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11175
      Left            =   5640
      TabIndex        =   103
      Top             =   1680
      Width           =   19095
      Begin VB.CommandButton CMD10 
         Caption         =   "Next"
         Height          =   810
         Left            =   17760
         MouseIcon       =   "Frm49.frx":9D4B
         MousePointer    =   99  'Custom
         Picture         =   "Frm49.frx":A055
         Style           =   1  'Graphical
         TabIndex        =   129
         ToolTipText     =   "Paparan Seterusnya"
         Top             =   10080
         Width           =   1095
      End
      Begin VB.CommandButton CMD9 
         Caption         =   "Back"
         Height          =   810
         Left            =   16560
         MouseIcon       =   "Frm49.frx":B11F
         MousePointer    =   99  'Custom
         Picture         =   "Frm49.frx":B429
         Style           =   1  'Graphical
         TabIndex        =   128
         ToolTipText     =   "Paparan Sebelum"
         Top             =   10080
         Width           =   1095
      End
      Begin MSComctlLib.ListView LV2 
         Height          =   9420
         Left            =   240
         TabIndex        =   104
         Top             =   600
         Width           =   18675
         _ExtentX        =   32941
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
         Left            =   1680
         TabIndex        =   127
         Top             =   10080
         Width           =   1335
      End
      Begin VB.Label Label65 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan Data :"
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
         Left            =   120
         TabIndex        =   126
         Top             =   10080
         Width           =   1455
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
         Left            =   15960
         TabIndex        =   125
         Top             =   10080
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
         Left            =   15360
         TabIndex        =   124
         Top             =   10080
         Width           =   375
      End
      Begin VB.Label L69_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L69_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3840
         TabIndex        =   123
         Top             =   10680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L70_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L70_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3840
         TabIndex        =   122
         Top             =   10320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label64 
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
         Left            =   14040
         TabIndex        =   121
         Top             =   10080
         Width           =   2295
      End
   End
   Begin VB.PictureBox Pic2 
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
      Height          =   11175
      Left            =   18000
      ScaleHeight     =   11175
      ScaleWidth      =   21195
      TabIndex        =   39
      Top             =   720
      Visible         =   0   'False
      Width           =   21195
      Begin VB.CommandButton CMD5 
         Caption         =   "Carian"
         Height          =   350
         Left            =   19080
         MouseIcon       =   "Frm49.frx":C4F3
         MousePointer    =   99  'Custom
         TabIndex        =   47
         Top             =   300
         Width           =   1815
      End
      Begin VB.TextBox TB18 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   16320
         TabIndex        =   40
         Top             =   320
         Width           =   2685
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   10155
         Left            =   120
         TabIndex        =   46
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   720
         Width           =   20925
         _ExtentX        =   36909
         _ExtentY        =   17912
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
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai pekerja yang telah didaftarkan ke dalam sistem."
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   42
         Top             =   480
         Width           =   6555
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan No. Kad Pengenalan pekerja bagi mencari data terperinci."
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   14280
         TabIndex        =   41
         Top             =   0
         Width           =   8235
      End
      Begin VB.Shape Shape1 
         Height          =   690
         Left            =   14160
         Top             =   0
         Width           =   6855
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Kad Pengenalan  :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   14400
         TabIndex        =   43
         Top             =   360
         Width           =   2235
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   10680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm49.frx":C7FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm49.frx":EDD7
            Key             =   ""
         EndProperty
      EndProperty
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
      Height          =   7575
      Left            =   23520
      ScaleHeight     =   7575
      ScaleWidth      =   16935
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   16935
      Begin VB.TextBox TB17 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2940
         TabIndex        =   32
         Top             =   650
         Width           =   3360
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Carian Mengikut No. Kad Pengenalan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   31
         Top             =   240
         Width           =   4455
      End
      Begin VB.OptionButton Opt2 
         Caption         =   "Carian Mengikut Tarikh Penyertaan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   30
         Top             =   240
         Width           =   4215
      End
      Begin VB.OptionButton Opt3 
         Caption         =   "Semua Ahli"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12600
         TabIndex        =   28
         Top             =   240
         Width           =   3735
      End
      Begin VB.CommandButton CMD19 
         BackColor       =   &H0080C0FF&
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   12360
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   2055
      End
      Begin VB.CommandButton CMD20 
         BackColor       =   &H0080C0FF&
         Caption         =   "Excel Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   14640
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   420
         Left            =   1100
         TabIndex        =   29
         Top             =   1000
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   741
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   16744576
         Format          =   140771328
         CurrentDate     =   41561
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   5895
         Left            =   120
         TabIndex        =   33
         ToolTipText     =   "Sila ""Double Click"" untuk Menu seterusnya."
         Top             =   1560
         Width           =   16635
         _ExtentX        =   29342
         _ExtentY        =   10398
         _Version        =   393216
         Rows            =   1
         BackColor       =   12648447
         BackColorBkg    =   8421631
         GridColor       =   12582912
         GridColorFixed  =   12582912
         WordWrap        =   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   420
         Left            =   3800
         TabIndex        =   34
         Top             =   1080
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   741
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   16744576
         Format          =   140771328
         CurrentDate     =   41561
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Height          =   405
         Left            =   2820
         TabIndex        =   38
         Top             =   660
         Width           =   135
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Kad Pengenalan"
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
         Height          =   405
         Left            =   360
         TabIndex        =   37
         Top             =   660
         Width           =   2400
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "Dari"
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
         Height          =   285
         Left            =   6600
         TabIndex        =   36
         Top             =   660
         Width           =   1080
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "Hingga"
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
         Height          =   285
         Left            =   6600
         TabIndex        =   35
         Top             =   1120
         Width           =   1080
      End
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   11175
      Left            =   120
      TabIndex        =   48
      Top             =   120
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   19711
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Scan Item"
         Object.Width           =   2540
         ImageIndex      =   1
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Maklumat Pembeli"
         Object.Width           =   2540
         ImageIndex      =   2
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bayaran"
         Object.Width           =   2540
         ImageIndex      =   3
      EndProperty
   End
   Begin VB.Label Label38 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai pekerja"
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
      Left            =   3480
      MouseIcon       =   "Frm49.frx":113B1
      MousePointer    =   99  'Custom
      TabIndex        =   45
      Top             =   12000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label L3_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pendaftaran data pekerja"
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
      Left            =   720
      MouseIcon       =   "Frm49.frx":116BB
      MousePointer    =   99  'Custom
      TabIndex        =   44
      Top             =   12000
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Menu Frm49_Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm49_LihatData 
         Caption         =   "Lihat Data / Edit Data"
      End
      Begin VB.Menu frm49_sm_spacer1 
         Caption         =   "-"
      End
      Begin VB.Menu Frm49_PadamData 
         Caption         =   "Padam Data"
         Visible         =   0   'False
      End
      Begin VB.Menu Frm49_Excel 
         Caption         =   "Export Excel Report"
      End
   End
End
Attribute VB_Name = "Frm49"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB1_Click()
'On Error Resume Next
If Frm49.CB1 = 1 Then
    Frm49.CB2 = 0
    Frm49.CB3 = 0
    Frm49.Frame6.Visible = True
End If
End Sub
Private Sub CB2_Click()
'On Error Resume Next
If Frm49.CB2 = 1 Then
    Frm49.CB1 = 0
    Frm49.CB3 = 0
    Frm49.Frame6.Visible = True
End If
End Sub
Private Sub CB3_Click()
'On Error Resume Next
If Frm49.CB3 = 1 Then
    Frm49.CB2 = 0
    Frm49.CB1 = 0
    Frm49.Frame6.Visible = True
End If
End Sub
Private Sub CBB1_Change()
'On Error Resume Next
If Frm49.CBB1 = "Berhenti" Then
    Frm49.L1_Label.Visible = True
    Frm49.L2_Label.Visible = True
    Frm49.DTPicker4.Visible = True
Else
    Frm49.L1_Label.Visible = False
    Frm49.L2_Label.Visible = False
    Frm49.DTPicker4.Visible = False
End If
End Sub
Private Sub CBB1_Click()
'On Error Resume Next
If Frm49.CBB1 = "Berhenti" Then
    Frm49.L1_Label.Visible = True
    Frm49.L2_Label.Visible = True
    Frm49.DTPicker4.Visible = True
Else
    Frm49.L1_Label.Visible = False
    Frm49.L2_Label.Visible = False
    Frm49.DTPicker4.Visible = False
End If
End Sub

Private Sub CMD1_Click()
'on error resume next
If Frm49.TB20 = vbNullString Then
    
    If Frm49.Option1 = True Then MsgBox "Sila masukkan no kad pengenalan.", vbExclamation, "Info"
    If Frm49.Option2 = True Then MsgBox "Sila masukkan nama.", vbExclamation, "Info"
    If Frm49.Option3 = True Then MsgBox "Sila masukkan no pekerja.", vbExclamation, "Info"
    
    Exit Sub

End If

If Frm49.TB20 <> vbNullString Then
    If InStr(1, Frm49.TB20, "&") <> 0 Or InStr(1, Frm49.TB20, "*") <> 0 Or InStr(1, Frm49.TB20, "/") <> 0 Or InStr(1, Frm49.TB20, "\") <> 0 Or InStr(1, Frm49.TB20, "'") <> 0 Then

        If Frm49.Option1 = True Then MsgBox "No kad pengenalan mempunyai simbol yang tidak dibenarkan.", vbExclamation, "Info"
        If Frm49.Option2 = True Then MsgBox "Nama mempunyai simbol yang tidak dibenarkan.", vbExclamation, "Info"
        If Frm49.Option3 = True Then MsgBox "No pekerja mempunyai simbol yang tidak dibenarkan.", vbExclamation, "Info"
    
        Exit Sub
        
    End If
End If

Frm49.L69_Text = -1 'Titik Pencarian Data
Frm49.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Frm49.L67_Text = 0 'Paparan Page ke-xxx
Frm49.L68_Text = 0
Frm49.L71_Text = 0

GM_NEXT_PREV = 0

Frm49.L7_Text = 1 '0 : Carian mengikut cawangan , 1 : Carian mengikut maklumat pekerja
Frm49.L6_Text = UCase(Frm49.TB20)

Call frm49_senarai_staff_header
Call frm49_senarai_staff

If Frm49.L71_Text = "0" Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End Sub

Private Sub CMD10_Click()
'on error resume next
Dim frm49_LM_CURR_PAGE As Double
Dim frm49_LM_TOTAL_PAGE As Double

frm49_LM_CURR_PAGE = 0
frm49_LM_TOTAL_PAGE = 0

If Frm49.L67_Text <> vbNullString And IsNumeric(Frm49.L67_Text) Then
    If Frm49.L68_Text <> vbNullString And IsNumeric(Frm49.L68_Text) Then
        frm49_LM_CURR_PAGE = Frm49.L67_Text
        frm49_LM_TOTAL_PAGE = Frm49.L68_Text
        
        If frm49_LM_CURR_PAGE < frm49_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call frm49_senarai_staff_header
            Call frm49_senarai_staff
                            
        End If
    End If
End If
End Sub

Private Sub CMD2_Click()
'on error resume next
If Frm49.CBB3 = vbNullString Then
    
    MsgBox "Sila pilih cawangan.", vbExclamation, "Info"
    
    Exit Sub

End If

Frm49.L7_Text = 0 '0 : Carian mengikut cawangan , 1 : Carian mengikut maklumat pekerja
Frm49.L6_Text = Frm49.CBB3

Frm49.L69_Text = -1 'Titik Pencarian Data
Frm49.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Frm49.L67_Text = 0 'Paparan Page ke-xxx
Frm49.L68_Text = 0
Frm49.L71_Text = 0

GM_NEXT_PREV = 0

Call frm49_senarai_staff_header
Call frm49_senarai_staff

If Frm49.L71_Text = "0" Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End Sub

Private Sub CMD4_Click()
'On Error Resume Next
Dim Err(15)
Dim rs1 As ADODB.Recordset

x = 0
DataUpdate = 0

If Frm49.TB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan [Nama Pekerja]."
End If
If Frm49.TB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan [No. Kad Pengenalan]."
End If
'If Frm49.TB4 = vbNullString Then
'    x = x + 1
'    Err(x) = "Tiada Data Bagi [No. Pekerja]."
'End If
If Frm49.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Status] Pekerja."
End If
If Frm49.TB12 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan [Jawatan] Pekerja."
End If
If Frm49.TB13 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan [Nama Samaran] Pekerja."
End If
If Frm49.TB13 <> vbNullString Then
    If InStr(1, Frm49.TB13, "&") <> 0 Or InStr(1, Frm49.TB13, "*") <> 0 Or InStr(1, Frm49.TB13, "/") <> 0 Or InStr(1, Frm49.TB13, "\") <> 0 Or InStr(1, Frm49.TB13, "'") <> 0 Then
        x = x + 1
        Err(x) = "Nama pekerja mempunyai simbol yang tidak dibenarkan."
    End If
End If
If Frm49.TB16 = vbNullString Then
    x = x + 1
    Err(x) = "Sila sediakan password kepada pekerja ini untuk login ke dalam sistem."
End If
If Frm49.TB16 <> vbNullString Then
    If InStr(1, Frm49.TB16, "&") <> 0 Or InStr(1, Frm49.TB16, "*") <> 0 Or InStr(1, Frm49.TB16, "/") <> 0 Or InStr(1, Frm49.TB16, "\") <> 0 Or InStr(1, Frm49.TB16, "'") <> 0 Then
        x = x + 1
        Err(x) = "Password mempunyai simbol yang tidak dibenarkan."
    End If
End If
If Frm49.TB14 = vbNullString Or (Frm49.TB14 <> vbNullString And Not IsNumeric(Frm49.TB14)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR yang dibenarkan dalam ruangan [Gaji Pokok]."
End If
If Frm49.TB15 = vbNullString Or (Frm49.TB15 <> vbNullString And Not IsNumeric(Frm49.TB15)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR yang dibenarkan dalam ruangan [Elaun]."
End If
If Frm49.CB1 = 0 And Frm49.CB2 = 0 And Frm49.CB3 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan [User Level] bagi pekerja ini."
End If
If Frm49.TB19 = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan E-mail."
End If
If Frm49.TB19 <> vbNullString Then
    myAt = InStr(1, Frm49.TB19, "@", vbTextCompare)
    myDot = InStr(myAt + 2, Frm49.TB19, ".", vbTextCompare)
    myDotDot = InStr(myAt + 2, Frm49.TB19, "..", vbTextCompare)
    
    If myAt = 0 Or myDot = 0 Or Not myDotDot = 0 Or Right(Frm49.TB19, 1) = "." Then
        x = x + 1
        Err(x) = "Email yang tidak sah."
    End If
End If
'If Frm49.CB1 = 0 Then
    
    If Frm49.CBB2 = vbNullString Then
    
        x = x + 1
        Err(x) = "Sila buat pilihan cawangan."
        
    End If

'End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah anda ingin mendaftarkan pekerja ini ?" & vbCrLf & _
            vbNullString & vbCrLf & _
            "Sila pastikan semua data adalah BETUL mengenai pekerja ini."
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then

'### Periksa kewujudan NAMA SAMARAN ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where Samaran='" & UCase(Frm49.TB13) & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            MsgBox "Nama [" & UCase(Frm49.TB13) & "] telah digunakan bagi pekerja yang bernama [" & rs!Nama & "]" & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila tukar nama samaran yang lain bagi tujuan pendaftaran pekerja ini.", vbExclamation, "Error"
                    
            
            rs.Close
            Set rs = Nothing
            
            Exit Sub
        End If
        
        rs.Close
        Set rs = Nothing
'### Periksa kewujudan NAMA SAMARAN ### - End

'### Periksa kewujudan NO. KAD PENGENALAN ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where NoIC='" & UCase(Frm49.TB2) & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            MsgBox "No. Kad Pengenalan [" & UCase(Frm49.TB2) & "] telah didaftarkan sebelum ini." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila periksa senarai pekerja anda dengan No. Kad Pengenalan ini.", vbExclamation, "Error"
                    
            
            rs.Close
            Set rs = Nothing
            
            Exit Sub
        End If
        
        rs.Close
        Set rs = Nothing
'### Periksa kewujudan NO. KAD PENGENALAN ### - End

'### Periksa kewujudan E-mail ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where mail='" & Frm49.TB19 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            MsgBox "E-mail [" & Frm49.TB19 & "] telah didaftarkan sebelum ini." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila periksa senarai pekerja anda dengan e-mail ini.", vbExclamation, "Error"
                    
            
            rs.Close
            Set rs = Nothing
            
            Exit Sub
        End If
        
        rs.Close
        Set rs = Nothing
'### Periksa kewujudan E-mail ### - End

        LM_NOW = Now
        
        LM_DATE = DateTime.Date
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        rs.Open "select * from 11_no_pekerja", cn2, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        rs!tarikh = LM_DATE
        rs!terminal = G_TERMINAL
        rs!write_timestamp = LM_NOW
        rs!Status = 1
        rs!nama_staff = MDI_frm1.L3_Text
        rs.Update
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        rs.Open "select * from 11_no_pekerja where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & LM_DATE & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!ID) Then
                
                LM_ID = Format(rs!ID, "00000")
                rs!no_pekerja = Format(rs!ID, "00000")
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

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee", cn, adOpenKeyset, adLockOptimistic

        rs.AddNew
        If Frm49.TB1 <> vbNullString Then 'Nama Pekerja
            rs!Nama = UCase(Frm49.TB1)
        Else
            rs!Nama = Null
        End If
        If Frm49.TB2 <> vbNullString Then 'No IC
            rs!NoIC = UCase(Frm49.TB2)
        Else
            rs!NoIC = Null
        End If
        If Frm49.TB3 <> vbNullString Then 'No Passport
            rs!NoPassport = UCase(Frm49.TB3)
        Else
            rs!NoPassport = Null
        End If
        rs!NoPekerja = LM_ID 'No Pekerja
        'If Frm49.TB4 <> vbNullString Then 'No Pekerja
        '    rs!NoPekerja = UCase(Frm49.TB4)
        'Else
        '    rs!NoPekerja = Null
        'End If
        If Frm49.TB5 <> vbNullString Then 'No KWSP
            rs!NoKWSP = UCase(Frm49.TB5)
        Else
            rs!NoKWSP = Null
        End If
        If Frm49.TB6 <> vbNullString Then 'No Socso
            rs!NoSocso = UCase(Frm49.TB6)
        Else
            rs!NoSocso = Null
        End If
        If Frm49.TB7 <> vbNullString Then 'No Tel
            rs!NoTel = Frm49.TB7
        Else
            rs!NoTel = Null
        End If
        If Frm49.TB8 <> vbNullString Then 'Alamat
            rs!Alamat1 = UCase(Frm49.TB8)
        Else
            rs!Alamat1 = Null
        End If
        If Frm49.TB9 <> vbNullString Then 'Nama Bank
            rs!alamat2 = UCase(Frm49.TB9)
        Else
            rs!alamat2 = Null
        End If
        If Frm49.TB10 <> vbNullString Then 'No. Akaun
            rs!alamat3 = UCase(Frm49.TB10)
        Else
            rs!alamat3 = Null
        End If
        rs!Status = Frm49.CBB1 'Status
        rs!TarikhMula = Frm49.DTPicker1 'Tarikh Mula Kerja
        If Frm49.CBB1 = "Berhenti" Then
            rs!TarikhBerhenti = Frm49.DTPicker4 'Tarikh Berhenti
        Else
            rs!TarikhBerhenti = Null
        End If
        If Frm49.TB12 <> vbNullString Then 'Jawatan
            rs!Jawatan = UCase(Frm49.TB12)
        Else
            rs!Jawatan = Null
        End If
        If Frm49.TB13 <> vbNullString Then 'Samaran
            rs!Samaran = UCase(Frm49.TB13)
        Else
            rs!Samaran = Null
        End If
        If Frm49.TB14 <> vbNullString Then 'Gaji
            rs!Gaji = Format(Frm49.TB14, "#,##0.00")
        Else
            rs!Gaji = "0.00"
        End If
        If Frm49.TB15 <> vbNullString Then 'Elaun
            rs!Elaun = Format(Frm49.TB15, "#,##0.00")
        Else
            rs!Elaun = "0.00"
        End If
        If Frm49.TB16 <> vbNullString Then 'Password
            rs!Password = Frm49.TB16
        Else
            rs!Password = Null
        End If
        If Frm49.CB1 = 1 Then
            rs!user_level = 1
        ElseIf Frm49.CB2 = 1 Then
            rs!user_level = 2
        ElseIf Frm49.CB3 = 1 Then
            rs!user_level = 3
        End If

'user_level
'1 : Admin
'2 : Manager
'3 : Staff
'4 : Guest/User -> Audit (Bagi menggunakan back end system 1 (external system)
'5 : Administration -> Audit (Bagi menggunakan back end system 2 (internal system)
'6 : HQ

        If Frm49.TB19 <> vbNullString Then 'E-mail
            rs!mail = Frm49.TB19
        Else
            rs!mail = Null
        End If
        
        rs!ElaunProfit = 0 '0 : Tiada Elaun Profit , 1 : Ada Elaun Profit
        rs!InvestorSmall = 0 '0 : Tiada Elaun Profit Investor (Small) , 1 : Ada Elaun Profit Investor (Small)
        rs!InvestorBig = 0 '0 : Tiada Elaun Profit Investor (Big) , 1 : Ada Elaun Profit Investor (Big)

        
        '%%%% TukangemaS - Komisen Pekerja %%%%
        If Frm49.CB4 = 1 Then
            rs!komisen = 1 'Pilihan samada pekerja ini layak untuk mendapat komisen dari jualan yang dilakukan. , 0:  Tidak Layak , 1:  Layak
        Else
            rs!komisen = 0 'Pilihan samada pekerja ini layak untuk mendapat komisen dari jualan yang dilakukan. , 0:  Tidak Layak , 1:  Layak
        End If
        '%%%% TukangemaS - Komisen Pekerja %%%%
        
        If Frm49.CBB2 <> vbNullString Then
            rs!cawangan = Frm49.CBB2
        Else
            rs!cawangan = Null
        End If
        
        rs.Update
    
        rs.Close
        Set rs = Nothing
            
        user = MDI_frm1.L3_Text
        LogAct_Memory = "[" & user & "] Pendaftaran pekerja baru. No. kad pengenalan [" & Frm49.TB2 & "]."
        LogDate_Memory = LM_NOW
        Call UpdateLog_Database
            
        'Call NewGenerateEmpNo
        Call frm49_Default
        MsgBox "Data telah berjaya disimpan.", vbInformation, "Info"
        
        Frm49.TB1.SetFocus
        
    End If
End If
End Sub
Private Sub CMD5_Click()
'On Error Resume Next
If Frm49.TB18 = vbNullString Then
    MsgBox "Sila Masukkan No. Kad Pengenalan Pekerja.", vbExclamation, "Error"
    Exit Sub
End If

Frm49.MSFlexGrid1.Clear
Frm49.MSFlexGrid1.RowHeight(0) = 800
Frm49.MSFlexGrid1.FormatString = "No.|<No.|<Nama|<No. Kad Pengenalan|<No. Pekerja|<No. EPF|<No. Income Tax|<No. Tel|<Jawatan|<Nama Samaran|<Password|<User Level|<Gaji (RM)|<Elaun (RM)|<Komisen Jualan|<Status"

Frm49.MSFlexGrid1.Rows = 1
Frm49.MSFlexGrid1.ColWidth(0) = 600
Frm49.MSFlexGrid1.ColWidth(1) = 0
Frm49.MSFlexGrid1.ColWidth(2) = 4800
Frm49.MSFlexGrid1.ColWidth(3) = 1700
Frm49.MSFlexGrid1.ColWidth(4) = 1200
Frm49.MSFlexGrid1.ColWidth(5) = 1200
Frm49.MSFlexGrid1.ColWidth(6) = 1200
Frm49.MSFlexGrid1.ColWidth(7) = 1200
Frm49.MSFlexGrid1.ColWidth(8) = 1200
Frm49.MSFlexGrid1.ColWidth(9) = 1200
Frm49.MSFlexGrid1.ColWidth(10) = 1200
Frm49.MSFlexGrid1.ColWidth(11) = 1200
Frm49.MSFlexGrid1.ColWidth(12) = 1000
Frm49.MSFlexGrid1.ColWidth(13) = 1000
Frm49.MSFlexGrid1.ColWidth(14) = 1000
Frm49.MSFlexGrid1.ColWidth(15) = 1000

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where NoIC='" & Frm49.TB18 & "' AND user_level <> 4 AND user_level <> 5", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm49.MSFlexGrid1.Rows = x + 1
    Frm49.MSFlexGrid1.TextMatrix(x, 0) = x
    Frm49.MSFlexGrid1.TextMatrix(x, 1) = x
    If Not IsNull(rs!Nama) Then Frm49.MSFlexGrid1.TextMatrix(x, 2) = rs!Nama 'Nama Pekerja
    If Not IsNull(rs!NoIC) Then Frm49.MSFlexGrid1.TextMatrix(x, 3) = rs!NoIC 'No. Kad Pengenalan
    If Not IsNull(rs!NoPekerja) Then Frm49.MSFlexGrid1.TextMatrix(x, 4) = rs!NoPekerja 'No Pekerja
    If Not IsNull(rs!NoKWSP) Then Frm49.MSFlexGrid1.TextMatrix(x, 5) = rs!NoKWSP 'No. KWSP
    If Not IsNull(rs!NoSocso) Then Frm49.MSFlexGrid1.TextMatrix(x, 6) = rs!NoSocso 'No. Socso
    If Not IsNull(rs!NoTel) Then Frm49.MSFlexGrid1.TextMatrix(x, 7) = rs!NoTel 'No. Tel
    If Not IsNull(rs!Jawatan) Then Frm49.MSFlexGrid1.TextMatrix(x, 8) = rs!Jawatan 'Jawatan
    If Not IsNull(rs!Samaran) Then Frm49.MSFlexGrid1.TextMatrix(x, 9) = rs!Samaran 'Nama Samaran
    If Not IsNull(rs!Password) Then Frm49.MSFlexGrid1.TextMatrix(x, 10) = rs!Password 'Password
    If Not IsNull(rs!user_level) Then
        If rs!user_level = 1 Then
            Frm49.MSFlexGrid1.TextMatrix(x, 11) = "Admin" 'User Level
        ElseIf rs!user_level = 2 Then
            Frm49.MSFlexGrid1.TextMatrix(x, 11) = "Manager" 'User Level
        ElseIf rs!user_level = 3 Then
            Frm49.MSFlexGrid1.TextMatrix(x, 11) = "Staff" 'User Level
        End If
    Else
        Frm49.MSFlexGrid1.TextMatrix(x, 11) = "Staff" 'User Level
    End If
    If Not IsNull(rs!Gaji) Then Frm49.MSFlexGrid1.TextMatrix(x, 12) = rs!Gaji 'Gaji
    If Not IsNull(rs!Elaun) Then Frm49.MSFlexGrid1.TextMatrix(x, 13) = rs!Elaun 'Elaun
    If Not IsNull(rs!komisen) Then
        If rs!komisen = 0 Then
            Frm49.MSFlexGrid1.TextMatrix(x, 14) = "Tidak"
        Else
            Frm49.MSFlexGrid1.TextMatrix(x, 14) = "Ya"
        End If
    End If
    If Not IsNull(rs!Status) Then
        Frm49.MSFlexGrid1.TextMatrix(x, 15) = rs!Status 'Status
    Else
        Frm49.MSFlexGrid1.TextMatrix(x, 15) = "Aktif" 'Status
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm49.TB18 = vbNullString
End Sub
Private Sub CMD6_Click()
'On Error Resume Next
Dim Err(15)
x = 0
DataUpdate = 0

If Frm49.TB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan [Nama Pekerja]."
End If
If Frm49.TB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan [No. Kad Pengenalan]."
End If
If Frm49.TB4 = vbNullString Then
    x = x + 1
    Err(x) = "Tiada Data Bagi [No. Pekerja]."
End If
If Frm49.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Status] Pekerja."
End If
If Frm49.TB12 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan [Jawatan] Pekerja."
End If
If Frm49.TB13 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan [Nama Samaran] Pekerja."
End If
If Frm49.TB14 = vbNullString Or (Frm49.TB14 <> vbNullString And Not IsNumeric(Frm49.TB14)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR yang dibenarkan dalam ruangan [Gaji Pokok]."
End If
If Frm49.TB15 = vbNullString Or (Frm49.TB15 <> vbNullString And Not IsNumeric(Frm49.TB15)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR yang dibenarkan dalam ruangan [Elaun]."
End If
If Frm49.TB16 = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan password."
End If
If Frm49.CB1 = 0 And Frm49.CB2 = 0 And Frm49.CB3 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan [User Level] bagi pekerja ini."
End If
If Frm49.TB19 = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan E-mail."
End If
If Frm49.TB19 <> vbNullString Then
    myAt = InStr(1, Frm49.TB19, "@", vbTextCompare)
    myDot = InStr(myAt + 2, Frm49.TB19, ".", vbTextCompare)
    myDotDot = InStr(myAt + 2, Frm49.TB19, "..", vbTextCompare)
    
    If myAt = 0 Or myDot = 0 Or Not myDotDot = 0 Or Right(Frm49.TB19, 1) = "." Then
        x = x + 1
        Err(x) = "Email yang tidak sah."
    End If
End If
'If Frm49.CB1 = 0 Then
    
    If Frm49.CBB2 = vbNullString Then
    
        x = x + 1
        Err(x) = "Sila buat pilihan cawangan."
        
    End If

'End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else

    Note = "Adakah anda ingin simpan data yang telah diedit bagi pekerja ini ?" & vbCrLf & _
            vbNullString & vbCrLf & _
            "Sila pastikan semua data adalah BETUL mengenai pekerja ini."
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where (Samaran='" & UCase(Frm49.TB13) & "' OR NoIC='" & UCase(Frm49.TB2) & "' OR mail='" & Frm49.TB19 & "')", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Frm49.L4_Text <> rs!ID Then
            
                'MsgBox "Nama [" & UCase(Frm49.TB13) & "] telah digunakan bagi pekerja yang bernama [" & rs!Nama & "]" & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Sila tukar nama samaran yang lain bagi tujuan pendaftaran pekerja ini.", vbExclamation, "Error"
                        
                
                MsgBox "Nama [" & UCase(Frm49.TB13) & "] atau No. Kad Pengenalan [" & UCase(Frm49.TB2) & "] atau E-mail [" & UCase(Frm49.TB19) & "] telah ada di dalam sistem." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Sila periksa data anda.", vbExclamation, "Error"
                        
                rs.Close
                Set rs = Nothing
                
                Exit Sub
                
            End If
            
        End If
        
        rs.Close
        Set rs = Nothing
        
        GoTo aaa:
        
        
    
'### Perika kewujudan NAMA SAMARAN ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where Samaran='" & UCase(Frm49.TB13) & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Frm49.L4_Text <> rs!ID Then
            
                MsgBox "Nama [" & UCase(Frm49.TB13) & "] telah digunakan bagi pekerja yang bernama [" & rs!Nama & "]" & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Sila tukar nama samaran yang lain bagi tujuan pendaftaran pekerja ini.", vbExclamation, "Error"
                        
                
                rs.Close
                Set rs = Nothing
                
                Exit Sub
                
            End If
        End If
        
        rs.Close
        Set rs = Nothing
'### Perika kewujudan NAMA SAMARAN ### - End

'### Perika kewujudan NO. KAD PENGENALAN ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where NoIC='" & UCase(Frm49.TB2) & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Frm49.L4_Text <> rs!ID Then
            
                MsgBox "No. Kad Pengenalan [" & UCase(Frm49.TB2) & "] telah didaftarkan sebelum ini." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Sila periksa senarai pekerja anda dengan No. Kad Pengenalan ini.", vbExclamation, "Error"
                        
                
                rs.Close
                Set rs = Nothing
                
                Exit Sub
                
            End If
        End If
        
        rs.Close
        Set rs = Nothing
'### Perika kewujudan NO. KAD PENGENALAN ### - End

'### Periksa kewujudan E-mail ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where mail='" & Frm49.TB19 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Frm49.L4_Text <> rs!ID Then
            
                MsgBox "E-mail [" & Frm49.TB19 & "] telah didaftarkan sebelum ini." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Sila periksa senarai pekerja anda dengan e-mail ini.", vbExclamation, "Error"
                        
                
                rs.Close
                Set rs = Nothing
                
                Exit Sub
                
            End If
            
        End If
        
        rs.Close
        Set rs = Nothing
'### Periksa kewujudan E-mail ### - End
        
aaa:
        
        LM_NOW = Now
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where ID='" & Frm49.L4_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
        If Not rs.EOF Then
        
            If Frm49.TB1 <> vbNullString Then 'Nama Pekerja
                rs!Nama = UCase(Frm49.TB1)
            Else
                rs!Nama = Null
            End If
            If Frm49.TB2 <> vbNullString Then 'No IC
                rs!NoIC = UCase(Frm49.TB2)
            Else
                rs!NoIC = Null
            End If
            If Frm49.TB3 <> vbNullString Then 'No Passport
                rs!NoPassport = UCase(Frm49.TB3)
            Else
                rs!NoPassport = Null
            End If
            If Frm49.TB4 <> vbNullString Then 'No Pekerja
                rs!NoPekerja = UCase(Frm49.TB4)
            Else
                rs!NoPekerja = Null
            End If
            If Frm49.TB5 <> vbNullString Then 'No KWSP
                rs!NoKWSP = UCase(Frm49.TB5)
            Else
                rs!NoKWSP = Null
            End If
            If Frm49.TB6 <> vbNullString Then 'No Socso
                rs!NoSocso = UCase(Frm49.TB6)
            Else
                rs!NoSocso = Null
            End If
            If Frm49.TB7 <> vbNullString Then 'No Tel
                rs!NoTel = Frm49.TB7
            Else
                rs!NoTel = Null
            End If
            If Frm49.TB8 <> vbNullString Then 'Alamat
                rs!Alamat1 = UCase(Frm49.TB8)
            Else
                rs!Alamat1 = Null
            End If
            If Frm49.TB9 <> vbNullString Then 'Nama Bank
                rs!alamat2 = UCase(Frm49.TB9)
            Else
                rs!alamat2 = Null
            End If
            If Frm49.TB10 <> vbNullString Then 'No. Akaun
                rs!alamat3 = UCase(Frm49.TB10)
            Else
                rs!alamat3 = Null
            End If
            rs!Status = Frm49.CBB1 'Status
            rs!TarikhMula = Frm49.DTPicker1 'Tarikh Mula Kerja
            If Frm49.CBB1 = "Berhenti" Then
                rs!TarikhBerhenti = Frm49.DTPicker4 'Tarikh Berhenti
            Else
                rs!TarikhBerhenti = Null
            End If
            If Frm49.TB12 <> vbNullString Then 'Jawatan
                rs!Jawatan = UCase(Frm49.TB12)
            Else
                rs!Jawatan = Null
            End If
            If Frm49.TB13 <> vbNullString Then 'Samaran
                rs!Samaran = UCase(Frm49.TB13)
            Else
                rs!Samaran = Null
            End If
            If Frm49.TB14 <> vbNullString Then 'Gaji
                rs!Gaji = Format(Frm49.TB14, "#,##0.00")
            Else
                rs!Gaji = "0.00"
            End If
            If Frm49.TB15 <> vbNullString Then 'Elaun
                rs!Elaun = Format(Frm49.TB15, "#,##0.00")
            Else
                rs!Elaun = "0.00"
            End If
            If Frm49.TB16 <> vbNullString Then 'Password
                rs!Password = Frm49.TB16
            Else
                rs!Password = Null
            End If
            If Frm49.CB1 = 1 Then
                rs!user_level = 1
            ElseIf Frm49.CB2 = 1 Then
                rs!user_level = 2
            ElseIf Frm49.CB3 = 1 Then
                rs!user_level = 3
            End If
            If Frm49.TB19 <> vbNullString Then 'E-mail
                rs!mail = Frm49.TB19
            Else
                rs!mail = Null
            End If
            
            rs!ElaunProfit = 0 '0 : Tiada Elaun Profit , 1 : Ada Elaun Profit
            rs!InvestorSmall = 0 '0 : Tiada Elaun Profit Investor (Small) , 1 : Ada Elaun Profit Investor (Small)
            rs!InvestorBig = 0 '0 : Tiada Elaun Profit Investor (Big) , 1 : Ada Elaun Profit Investor (Big)
    
            
            '%%%% TukangemaS - Komisen Pekerja %%%%
            If Frm49.CB4 = 1 Then
                rs!komisen = 1 'Pilihan samada pekerja ini layak untuk mendapat komisen dari jualan yang dilakukan. , 0:  Tidak Layak , 1:  Layak
            Else
                rs!komisen = 0 'Pilihan samada pekerja ini layak untuk mendapat komisen dari jualan yang dilakukan. , 0:  Tidak Layak , 1:  Layak
            End If
            '%%%% TukangemaS - Komisen Pekerja %%%%

            If Frm49.CBB2 <> vbNullString Then
                rs!cawangan = Frm49.CBB2
            Else
                rs!cawangan = Null
            End If
        
            rs.Update
        End If
        
        rs.Close
        Set rs = Nothing
        
        user = MDI_frm1.L3_Text
        LogAct_Memory = "[" & user & "] Edit data pekerja. No. Pekerja [" & Frm49.TB4 & "]."
        LogDate_Memory = LM_NOW
        Call UpdateLog_Database
        
        GM_NEXT_PREV = 2
        
        Call frm49_senarai_staff_header
        Call frm49_senarai_staff
        
        Frm49.Frame2.Visible = True
        Frm49.Frame1.Visible = False
        
        Call frm49_Default
        MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
    End If
End If
End Sub
Private Sub CMD7_Click()
'On Error Resume Next
Frm49.Frame1.Visible = False
Frm49.Frame2.Visible = True
End Sub



Private Sub CMD9_Click()
'on error resume next
Dim frm49_LM_CURR_PAGE As Double
Dim frm49_LM_TOTAL_PAGE As Double

frm49_LM_CURR_PAGE = 0
frm49_LM_TOTAL_PAGE = 0

If Frm49.L67_Text <> vbNullString And IsNumeric(Frm49.L67_Text) Then
    If Frm49.L68_Text <> vbNullString And IsNumeric(Frm49.L68_Text) Then
        frm49_LM_CURR_PAGE = Frm49.L67_Text
        frm49_LM_TOTAL_PAGE = Frm49.L68_Text
        
        If frm49_LM_CURR_PAGE <> 1 And frm49_LM_CURR_PAGE <> 0 Then
        
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
                                
                Call frm49_senarai_staff_header
                Call frm49_senarai_staff
                            
        End If

    End If
End If
End Sub

Private Sub Frm49_Excel_Click()
'On Error Resume Next
Dim frm49_field_1 As String
DATA_FOUND = 0

frm49_LM_No_ID = vbNullString

If IsNumeric(Frm49.LV2.SelectedItem.Index) Then
    
    frm49_LM_No_ID = Frm49.LV2.ListItems(Frm49.LV2.SelectedItem.Index)
    
    If frm49_LM_No_ID <> vbNullString Then
    
        Dim xlObject As Excel.Application
        Dim xlWB As Excel.Workbook
               
        Note = "Sistem akan mengambil masa untuk mengeluarkan report." & vbCrLf & _
                vbNullString & vbCrLf & _
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
                .Columns("B").ColumnWidth = 50 'Nama
                .Columns("C").ColumnWidth = 20 'No. Kad Pengenalan
                .Columns("D").ColumnWidth = 15 'No. Pekerja
                .Columns("E").ColumnWidth = 20 'No. Telefon
                .Columns("F").ColumnWidth = 20 'No. EPF
                .Columns("G").ColumnWidth = 20 'No. Income Tax
                .Columns("H").ColumnWidth = 20 'Jawatan
                .Columns("I").ColumnWidth = 20 'Status
                .Columns("J").ColumnWidth = 20 'Tarikh Masuk
                .Columns("K").ColumnWidth = 20 'Tarikh Berhenti
                .Columns("L").ColumnWidth = 20 'User Level
                .Columns("M").ColumnWidth = 20 'Username
                .Columns("N").ColumnWidth = 20 'Password
                .Columns("O").ColumnWidth = 20 'Cawangan
                .Columns("P").ColumnWidth = 50 'E-mail
                .Columns("Q").ColumnWidth = 20 'Komisen
                .Columns("R").ColumnWidth = 20 'Gaji (RM)
                .Columns("S").ColumnWidth = 20 'Elaun (RM)
                .Columns("T").ColumnWidth = 40 'Nama Bank
                .Columns("U").ColumnWidth = 30 'No. akaun
                
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
                
                .Cells(1, 5).Font.Bold = True
                .Cells(1, 5).Font.Size = 30
                
                For Row = 1 To 5
                    .Cells(Row, 5).HorizontalAlignment = xlCenter
                Next Row
                
                .Cells(7, 1) = "Senarai pekerja."
                
                .Cells(8, 1) = "No."
                .Cells(8, 2) = "Nama"
                .Cells(8, 3) = "No. Kad Pengenalan"
                .Cells(8, 4) = "No. Pekerja"
                .Cells(8, 5) = "No. Telefon"
                .Cells(8, 6) = "No. EPF"
                .Cells(8, 7) = "No. Income Tax"
                .Cells(8, 8) = "Jawatan"
                .Cells(8, 9) = "Status"
                .Cells(8, 10) = "Tarikh Masuk"
                .Cells(8, 11) = "Tarikh Berhenti"
                .Cells(8, 12) = "User Level"
                .Cells(8, 13) = "Username"
                .Cells(8, 14) = "Password"
                .Cells(8, 15) = "Cawangan"
                .Cells(8, 16) = "E-mail"
                .Cells(8, 17) = "Komisen"
                .Cells(8, 18) = "Gaji (RM)"
                .Cells(8, 19) = "Elaun (RM)"
                .Cells(8, 20) = "Nama Bank"
                .Cells(8, 21) = "No. Akaun"
            
                For i = 1 To 21
                    .Cells(8, i).HorizontalAlignment = xlCenter
                    .Cells(8, i).Interior.ColorIndex = 15
                    .Cells(8, i).WrapText = True
                    .Cells(8, i).Borders.LineStyle = xlContinuous
                Next i
                
                x = 0
                
                frm49_LM_SEARCH_1 = Frm49.L6_Text
                frm49_LM_SEARCH_1_LOGIC = "="
                        
                If Frm49.L7_Text = "0" Then '0 : Carian mengikut cawangan , 1 : Carian mengikut maklumat pekerja
                
                    frm49_field_1 = "cawangan"
                    
                    If Frm49.L6_Text = "Semua cawangan" Then
                        frm49_LM_SEARCH_1 = Null
                        frm49_LM_SEARCH_1_LOGIC = "<>"
                    Else
                        frm49_LM_SEARCH_1 = Frm49.L6_Text
                        frm49_LM_SEARCH_1_LOGIC = "="
                    End If
                    
                ElseIf Frm49.L7_Text = "1" Then
                
                    If Frm49.Option1 = True Then frm49_field_1 = "NoIC"
                    If Frm49.Option2 = True Then frm49_field_1 = "Nama"
                    If Frm49.Option3 = True Then frm49_field_1 = "NoPekerja"
                    
                    frm49_LM_SEARCH_1 = Frm49.L6_Text
                    frm49_LM_SEARCH_1_LOGIC = "="
                        
                End If
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from employee where " & frm49_field_1 & " " & frm49_LM_SEARCH_1_LOGIC & "'" & frm49_LM_SEARCH_1 & "' AND (user_level <> 4 AND user_level <> 5 AND user_level <> 6) order by nama ASC", cn, adOpenKeyset, adLockOptimistic
                
                While rs.EOF = False
                
                    x = x + 1
                
                    .Cells(8 + x, 1) = x 'No.
                    .Cells(8 + x, 1).HorizontalAlignment = xlCenter
                    
                    If Not IsNull(rs!Nama) Then .Cells(8 + x, 2) = rs!Nama 'Nama
                    
                    If Not IsNull(rs!NoIC) Then .Cells(8 + x, 3) = rs!NoIC 'No. Kad Pengenalan
                    
                    If Not IsNull(rs!NoPekerja) Then .Cells(8 + x, 4) = rs!NoPekerja 'No. Pekerja
                    
                    .Cells(8 + x, 5).NumberFormat = "@"
                    If Not IsNull(rs!NoTel) Then .Cells(8 + x, 5) = rs!NoTel 'No. Tel
                    .Cells(8 + x, 5).NumberFormat = "@"
                    
                    .Cells(8 + x, 6).NumberFormat = "@"
                    If Not IsNull(rs!NoKWSP) Then .Cells(8 + x, 6) = rs!NoKWSP 'No. EPF
                    .Cells(8 + x, 6).NumberFormat = "@"
                    
                    .Cells(8 + x, 7).NumberFormat = "@"
                    If Not IsNull(rs!NoSocso) Then .Cells(8 + x, 7) = rs!NoSocso 'No. Income Tax
                    .Cells(8 + x, 7).NumberFormat = "@"
                    
                    .Cells(8 + x, 8).NumberFormat = "@"
                    If Not IsNull(rs!Jawatan) Then .Cells(8 + x, 8) = rs!Jawatan 'Jawatan
                    .Cells(8 + x, 8).NumberFormat = "@"
                    
                    .Cells(8 + x, 9).NumberFormat = "@"
                    If Not IsNull(rs!Status) Then .Cells(8 + x, 9) = rs!Status 'Status
                    .Cells(8 + x, 9).NumberFormat = "@"
                    
                    If Not IsNull(rs!TarikhMula) Then .Cells(8 + x, 10) = "'" & rs!TarikhMula 'Tarikh masuk
                    .Cells(8 + x, 10).HorizontalAlignment = xlCenter
                
                    If Not IsNull(rs!TarikhBerhenti) Then .Cells(8 + x, 11) = "'" & rs!TarikhBerhenti 'Tarikh berhenti
                    .Cells(8 + x, 11).HorizontalAlignment = xlCenter
                
                    .Cells(8 + x, 12).NumberFormat = "@"
                    If Not IsNull(rs!user_level) Then
                
                        If rs!user_level = 1 Then
                            .Cells(8 + x, 12) = "Admin" 'User Level
                        ElseIf rs!user_level = 2 Then
                            .Cells(8 + x, 12) = "Manager" 'User Level
                        ElseIf rs!user_level = 3 Then
                            .Cells(8 + x, 12) = "Staff"  'User Level
                        End If
                    
                    Else
                    
                        .Cells(8 + x, 12) = "Staff"  'User Level
                        
                    End If
                    .Cells(8 + x, 12).NumberFormat = "@"
                
                    .Cells(8 + x, 13).NumberFormat = "@"
                    If Not IsNull(rs!Samaran) Then .Cells(8 + x, 13) = rs!Samaran 'username
                    .Cells(8 + x, 13).NumberFormat = "@"
                        
                    .Cells(8 + x, 14).NumberFormat = "@"
                    If Not IsNull(rs!Password) Then .Cells(8 + x, 14) = rs!Password 'Password
                    .Cells(8 + x, 14).NumberFormat = "@"
                
                    .Cells(8 + x, 15).NumberFormat = "@"
                    If Not IsNull(rs!cawangan) Then .Cells(8 + x, 15) = rs!cawangan 'Cawangan
                    .Cells(8 + x, 15).NumberFormat = "@"
                        
                    .Cells(8 + x, 16).NumberFormat = "@"
                    If Not IsNull(rs!mail) Then .Cells(8 + x, 16) = rs!mail 'E-mail
                    .Cells(8 + x, 16).NumberFormat = "@"
                
                    If Not IsNull(rs!komisen) Then
                        
                        If rs!komisen = 0 Then
                            .Cells(8 + x, 17) = "Tiada"
                        ElseIf rs!komisen = 1 Then
                            .Cells(8 + x, 17) = "Ada"
                        End If
                        
                    Else
                    
                        .Cells(8 + x, 17) = "Tiada"
                    
                    End If
                    
                    .Cells(8 + x, 18).NumberFormat = "#,##0.00"
                    .Cells(8 + x, 18).HorizontalAlignment = xlRight
                    If Not IsNull(rs!Gaji) Then 'Gaji (RM)
                        .Cells(8 + x, 18) = Format(rs!Gaji, "#,##0.00")
                    Else
                        .Cells(8 + x, 18) = "0.00"
                    End If
                    
                    .Cells(8 + x, 19).NumberFormat = "#,##0.00"
                    .Cells(8 + x, 19).HorizontalAlignment = xlRight
                    If Not IsNull(rs!Elaun) Then 'Elaun (RM)
                        .Cells(8 + x, 19) = Format(rs!Elaun, "#,##0.00")
                    Else
                        .Cells(8 + x, 19) = "0.00"
                    End If
                    
                    .Cells(8 + x, 20).NumberFormat = "@"
                    If Not IsNull(rs!alamat2) Then .Cells(8 + x, 20) = rs!alamat2 'Nama Bank
                    .Cells(8 + x, 20).NumberFormat = "@"
                    
                    .Cells(8 + x, 21).NumberFormat = "@"
                    If Not IsNull(rs!alamat3) Then .Cells(8 + x, 21) = rs!alamat3 'No. Akaun
                    .Cells(8 + x, 21).NumberFormat = "@"
                    
                    For Col = 1 To 21
                        .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
                    Next Col
                    
                    rs.MoveNext
                Wend
                
                rs.Close
                Set rs = Nothing
                
                
                Y = x + 1
                .Cells(8 + Y, 1) = "Bilangan : " & x
                
                Y = Y + 4
                
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
Private Sub Frm49_LihatData_Click()
'On Error Resume Next
DATA_FOUND = 0

frm49_LM_No_ID = vbNullString

If IsNumeric(Frm49.LV2.SelectedItem.Index) Then
    
    frm49_LM_No_ID = Frm49.LV2.ListItems(Frm49.LV2.SelectedItem.Index)
    
    If frm49_LM_No_ID <> vbNullString Then

        Call frm49_Default
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where ID='" & frm49_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!ID) Then Frm49.L4_Text = rs!ID
            If Not IsNull(rs!Nama) Then Frm49.TB1 = rs!Nama 'Nama Pekerja
            If Not IsNull(rs!NoIC) Then Frm49.TB2 = rs!NoIC 'No IC
            If Not IsNull(rs!NoPassport) Then Frm49.TB3 = rs!NoPassport 'No Passport
            If Not IsNull(rs!NoPekerja) Then Frm49.TB4 = rs!NoPekerja 'No Pekerja
            If Not IsNull(rs!NoKWSP) Then Frm49.TB5 = rs!NoKWSP 'No. KWSP
            If Not IsNull(rs!NoSocso) Then Frm49.TB6 = rs!NoSocso 'No. Socso
            If Not IsNull(rs!NoTel) Then Frm49.TB7 = rs!NoTel 'No. Tel
            If Not IsNull(rs!Alamat1) Then Frm49.TB8 = rs!Alamat1 'Alamat 1
            If Not IsNull(rs!alamat2) Then Frm49.TB9 = rs!alamat2 'Alamat 2
            If Not IsNull(rs!alamat3) Then Frm49.TB10 = rs!alamat3 'Alamat 3
            If Not IsNull(rs!Status) Then Frm49.CBB1 = rs!Status 'Status
            If Not IsNull(rs!TarikhMula) Then Frm49.DTPicker1 = rs!TarikhMula 'Tarikh Mula Kerja
            If Not IsNull(rs!TarikhBerhenti) Then Frm49.DTPicker4 = rs!TarikhBerhenti 'Tarikh Berhenti
            If Not IsNull(rs!Jawatan) Then Frm49.TB12 = rs!Jawatan 'Jawatan
            If Not IsNull(rs!Samaran) Then Frm49.TB13 = rs!Samaran 'Samaran
            If Not IsNull(rs!Gaji) Then Frm49.TB14 = rs!Gaji 'Gaji
            If Not IsNull(rs!Elaun) Then Frm49.TB15 = rs!Elaun 'Elaun
            If Not IsNull(rs!Password) Then Frm49.TB16 = rs!Password 'Password
            If Not IsNull(rs!mail) Then Frm49.TB19 = rs!mail 'E-mail
            If Not IsNull(rs!user_level) Then
                If rs!user_level = 1 Then
                    Frm49.CB1 = 1
                ElseIf rs!user_level = 2 Then
                    Frm49.CB2 = 1
                ElseIf rs!user_level = 3 Then
                    Frm49.CB3 = 1
                End If
            Else
                Frm49.CB3 = 1
            End If
    
            '%%%% TukangemaS - Komisen Pekerja %%%%
            If Not IsNull(rs!komisen) Then
                If rs!komisen = 0 Then
                    Frm49.CB4 = 0 'Pilihan samada pekerja ini layak untuk mendapat komisen dari jualan yang dilakukan. , 0:  Tidak Layak , 1:  Layak
                Else
                    Frm49.CB4 = 1 'Pilihan samada pekerja ini layak untuk mendapat komisen dari jualan yang dilakukan. , 0:  Tidak Layak , 1:  Layak
                End If
            End If
            '%%%% TukangemaS - Komisen Pekerja %%%%
             
            If Not IsNull(rs!cawangan) Then
            
            On Error GoTo Err_A:
            If Not IsNull(rs!cawangan) Then
                If rs!cawangan <> "HQ" Then
                
                    frm49_LM_CAWANGAN = rs!cawangan 'Cawangan
                    Frm49.CBB2 = frm49_LM_CAWANGAN 'Cawangan
                    
                End If
            End If
            
Restore_A:
            'on error resume next
            
            End If
             
            DATA_FOUND = 1 '0 : Data Not Found , 1 : Data Found
            
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then '0 : Data Not Found , 1 : Data Found
            
            Frm49.TB13.Locked = True
            Frm49.TB13.BackColor = &H8000000A
            
            Frm49.Frame1.Visible = True
            Frm49.Frame2.Visible = False
            Frm49.CMD4.Visible = False
            Frm49.CMD6.Visible = True
            Frm49.CMD7.Visible = True
            
        End If
    
    End If
    
End If

Exit Sub
Err_A:
Frm49.CBB2.AddItem frm49_LM_CAWANGAN
Frm49.CBB2 = frm49_LM_CAWANGAN
Resume Restore_A:
End Sub
Private Sub Frm49_PadamData_Click()
'On Error Resume Next
DATA_FOUND = 0 '0 : Data Not Found , 1 : Data Found
no_ic = Frm49.MSFlexGrid1.TextMatrix(Frm49.MSFlexGrid1, 3)

If no_ic <> vbNullString Then
    Note = "Adakah Anda Ingin Padam Data Ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where NoIC='" & no_ic & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            rs.Delete
            rs.Update
            DATA_FOUND = 1 '0 : Data Not Found , 1 : Data Found
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then '0 : Data Not Found , 1 : Data Found
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & user & "] Padam Data Pekerja. No. IC [" & Frm49.TB2 & "]."
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
            
            Call Frm49_EmpList

            Frm49.Pic2.Visible = True
            
            MsgBox "Data Telah Berjaya Dipadamkan.", vbInformation, "Info"
        End If
    End If
End If
End Sub

Private Sub LV1_Click()
'on error resume next
LM_KEY = Frm49.LV1.SelectedItem.Key

If LM_KEY = "Pendaftaran Pekerja" Then
    
    Call frm49_Default
    Call frm49_disable_form
    Frm49.Frame1.Visible = True
    
    Frm49.TB1.SetFocus
    
ElseIf LM_KEY = "Senarai Pekerja" Then

    Call frm49_disable_form
    Frm49.Frame3.Visible = True
    Frm49.Option2 = True
    
End If
End Sub
Private Sub LV2_DblClick()
'On Error Resume Next
frm49_LM_No_ID = vbNullString

If IsNumeric(Frm49.LV2.SelectedItem.Index) Then
    
    frm49_LM_No_ID = Frm49.LV2.ListItems(Frm49.LV2.SelectedItem.Index)
    
    If frm49_LM_No_ID <> vbNullString Then

        PopupMenu Frm49_Menu
        
    End If
    
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub Option1_Click()
'On Error Resume Next
Frm49.L5_Text = "No. Kad Pengenalan *"
End Sub

Private Sub Option2_Click()
'On Error Resume Next
Frm49.L5_Text = "Nama Pekerja *"
End Sub

Private Sub Option3_Click()
'On Error Resume Next
Frm49.L5_Text = "No. Pekerja *"
End Sub

