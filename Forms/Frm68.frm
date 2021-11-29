VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm68 
   Caption         =   "Maklumat Agen Dropship / Pelanggan & Maklumat Promosi"
   ClientHeight    =   13035
   ClientLeft      =   120
   ClientTop       =   -27585
   ClientWidth     =   23880
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
   Icon            =   "Frm68.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13035
   ScaleWidth      =   23880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pendaftaran Data Pelanggan / Data Pelanggan"
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
      Left            =   4920
      TabIndex        =   76
      Top             =   600
      Visible         =   0   'False
      Width           =   20535
      Begin VB.CheckBox CB13 
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
         Left            =   10680
         TabIndex        =   211
         Top             =   6270
         Visible         =   0   'False
         Width           =   200
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
         Height          =   1095
         Left            =   3000
         MouseIcon       =   "Frm68.frx":0ECA
         MousePointer    =   99  'Custom
         Picture         =   "Frm68.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   9480
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CommandButton CMD29 
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
         Left            =   5880
         MouseIcon       =   "Frm68.frx":379E
         MousePointer    =   99  'Custom
         Picture         =   "Frm68.frx":3AA8
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   9480
         Visible         =   0   'False
         Width           =   2775
      End
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
         Height          =   1095
         Left            =   4440
         MouseIcon       =   "Frm68.frx":6072
         MousePointer    =   99  'Custom
         Picture         =   "Frm68.frx":637C
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   9480
         Width           =   2775
      End
      Begin VB.CheckBox CB19 
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
         Left            =   11160
         TabIndex        =   22
         Top             =   5280
         Width           =   200
      End
      Begin VB.CheckBox CB20 
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
         Left            =   12240
         TabIndex        =   24
         Top             =   5280
         Width           =   200
      End
      Begin VB.TextBox TB19 
         Height          =   360
         Left            =   14040
         TabIndex        =   25
         Text            =   "TB19"
         Top             =   5640
         Width           =   2115
      End
      Begin VB.CheckBox CB14 
         BackColor       =   &H8000000C&
         Caption         =   "Scanner Mode"
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
         Left            =   11160
         TabIndex        =   21
         Top             =   4080
         Width           =   200
      End
      Begin VB.CheckBox CB17 
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
         Left            =   16440
         TabIndex        =   99
         Top             =   3720
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.CheckBox CB12 
         BackColor       =   &H8000000C&
         Caption         =   "Scanner Mode"
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
         Left            =   11280
         TabIndex        =   20
         Top             =   2325
         Width           =   200
      End
      Begin VB.CheckBox CB11 
         BackColor       =   &H8000000C&
         Caption         =   "Scanner Mode"
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
         Left            =   11280
         TabIndex        =   19
         Top             =   2085
         Width           =   200
      End
      Begin VB.CheckBox CB10 
         BackColor       =   &H8000000C&
         Caption         =   "Scanner Mode"
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
         Left            =   11280
         TabIndex        =   18
         Top             =   1845
         Width           =   200
      End
      Begin VB.CheckBox CB9 
         BackColor       =   &H8000000C&
         Caption         =   "Scanner Mode"
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
         Left            =   11280
         TabIndex        =   17
         Top             =   1605
         Width           =   200
      End
      Begin VB.TextBox TB11 
         Height          =   360
         Left            =   2280
         TabIndex        =   16
         Text            =   "TB11"
         Top             =   8040
         Width           =   8000
      End
      Begin VB.TextBox TB10 
         Height          =   360
         Left            =   2280
         TabIndex        =   15
         Text            =   "TB10"
         Top             =   7680
         Width           =   8000
      End
      Begin VB.TextBox TB9 
         Height          =   360
         Left            =   2280
         TabIndex        =   14
         Text            =   "TB9"
         Top             =   7320
         Width           =   8000
      End
      Begin VB.TextBox TB8 
         Height          =   1320
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "Frm68.frx":8946
         Top             =   5520
         Width           =   8000
      End
      Begin VB.TextBox TB7 
         Height          =   360
         Left            =   2280
         TabIndex        =   12
         Text            =   "TB7"
         Top             =   5160
         Width           =   8000
      End
      Begin VB.TextBox TB6 
         Height          =   360
         Left            =   2280
         TabIndex        =   11
         Text            =   "TB6"
         Top             =   4800
         Width           =   8000
      End
      Begin VB.TextBox TB5 
         Height          =   1320
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "Frm68.frx":894A
         Top             =   2880
         Width           =   8000
      End
      Begin VB.TextBox TB4 
         Height          =   360
         Left            =   2280
         TabIndex        =   9
         Text            =   "TB4"
         Top             =   2520
         Width           =   8000
      End
      Begin VB.TextBox TB3 
         Height          =   360
         Left            =   2280
         TabIndex        =   8
         Text            =   "TB3"
         Top             =   2160
         Width           =   3435
      End
      Begin VB.TextBox TB2 
         Height          =   360
         Left            =   2280
         TabIndex        =   6
         Text            =   "TB2"
         Top             =   1440
         Width           =   3435
      End
      Begin VB.TextBox TB1 
         Height          =   360
         Left            =   2280
         TabIndex        =   5
         Text            =   "TB1"
         Top             =   1080
         Width           =   8000
      End
      Begin VB.TextBox TB12 
         BackColor       =   &H8000000A&
         Height          =   360
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   79
         Text            =   "TB12"
         Top             =   1800
         Width           =   2955
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   360
         Left            =   2280
         TabIndex        =   7
         Top             =   1800
         Width           =   3435
         _ExtentX        =   6059
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
         Format          =   120455168
         CurrentDate     =   41561
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm68.frx":894E
         ForeColor       =   &H000000FF&
         Height          =   780
         Left            =   10920
         TabIndex        =   212
         Top             =   6240
         Visible         =   0   'False
         Width           =   5370
      End
      Begin VB.Label L66_Text 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "L66_Text"
         Height          =   255
         Left            =   17400
         TabIndex        =   105
         Top             =   1920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape Shape7 
         Height          =   1575
         Left            =   10560
         Top             =   4560
         Width           =   5775
      End
      Begin VB.Label Label73 
         BackStyle       =   0  'Transparent
         Caption         =   "Ya               Tidak"
         Height          =   255
         Left            =   11475
         TabIndex        =   104
         Top             =   5235
         Width           =   3495
      End
      Begin VB.Label Label82 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila isi ruang di bawah jika ada bayaran yang dikenakan bagi urusan pendaftaran ini."
         Height          =   495
         Left            =   10800
         TabIndex        =   103
         Top             =   4680
         UseMnemonic     =   0   'False
         Width           =   4935
      End
      Begin VB.Label L65_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah yuran pendaftaran *       :"
         Height          =   255
         Left            =   11040
         TabIndex        =   102
         Top             =   5685
         Width           =   3135
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila klik di bawah jika pelanggan ini adalah agen dropship bagi kedai."
         Height          =   495
         Left            =   10800
         TabIndex        =   101
         Top             =   3480
         UseMnemonic     =   0   'False
         Width           =   4935
      End
      Begin VB.Label Label53 
         BackStyle       =   0  'Transparent
         Caption         =   "Agen Dropship"
         Height          =   255
         Left            =   11475
         TabIndex        =   100
         Top             =   4035
         Width           =   2415
      End
      Begin VB.Shape Shape2 
         Height          =   1095
         Left            =   10560
         Top             =   3360
         Width           =   5775
      End
      Begin VB.Label L64_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "*** Tiada pilihan jenis pelanggan yang dibenarkan. Hanya pendaftaran PELANGGAN BIASA sahaja dibenarkan."
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
         Height          =   615
         Left            =   10920
         TabIndex        =   98
         Top             =   2640
         UseMnemonic     =   0   'False
         Width           =   4935
      End
      Begin VB.Label Label51 
         BackStyle       =   0  'Transparent
         Caption         =   "Ahli Biasa                   Silver                              Gold                        Platinum"
         ForeColor       =   &H00000000&
         Height          =   1485
         Left            =   11520
         TabIndex        =   97
         Top             =   1560
         Width           =   2385
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila buat pilihan kategori pelanggan."
         Height          =   255
         Left            =   10800
         TabIndex        =   96
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   4935
      End
      Begin VB.Label Label52 
         BackStyle       =   0  'Transparent
         Caption         =   "***jangan buat pilihan di bawah jika pelanggan ini adalah PELANGGAN BIASA."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10800
         TabIndex        =   95
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   4935
      End
      Begin VB.Shape Shape3 
         Height          =   2655
         Left            =   10560
         Top             =   600
         Width           =   5775
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Maklumat Bank"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   94
         Top             =   6960
         Width           =   10095
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Akaun :"
         Height          =   255
         Left            =   80
         TabIndex        =   93
         Top             =   8085
         Width           =   2175
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Akaun :"
         Height          =   255
         Left            =   80
         TabIndex        =   92
         Top             =   7725
         Width           =   2175
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Bank :"
         Height          =   255
         Left            =   80
         TabIndex        =   91
         Top             =   7365
         Width           =   2175
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat :"
         Height          =   255
         Left            =   80
         TabIndex        =   90
         Top             =   5565
         Width           =   2175
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Telefon :"
         Height          =   255
         Left            =   80
         TabIndex        =   89
         Top             =   5205
         Width           =   2175
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama :"
         Height          =   255
         Left            =   80
         TabIndex        =   88
         Top             =   4845
         Width           =   2175
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Maklumat Waris"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   87
         Top             =   4440
         Width           =   10095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat :"
         Height          =   255
         Left            =   80
         TabIndex        =   86
         Top             =   2925
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail :"
         Height          =   255
         Left            =   80
         TabIndex        =   85
         Top             =   2565
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Telefon :"
         Height          =   255
         Left            =   80
         TabIndex        =   84
         Top             =   2205
         Width           =   2175
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Pendaftaran * :"
         Height          =   255
         Left            =   80
         TabIndex        =   83
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Kad Pengenalan * :"
         Height          =   255
         Left            =   80
         TabIndex        =   82
         Top             =   1480
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama * :"
         Height          =   255
         Left            =   80
         TabIndex        =   81
         Top             =   1125
         Width           =   2175
      End
      Begin VB.Label L11_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Pelanggan :"
         Height          =   255
         Left            =   5760
         TabIndex        =   80
         Top             =   1845
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Maklumat Asas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   78
         Top             =   600
         Width           =   10095
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan data pelanggan di dalam ruangan di bawah."
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   300
         Width           =   7335
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Senarai Komisen Agen"
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
      Left            =   3360
      TabIndex        =   197
      Top             =   6360
      Visible         =   0   'False
      Width           =   20535
      Begin VB.CommandButton CMD9 
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
         Height          =   945
         Left            =   17760
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm68.frx":89F4
         MousePointer    =   99  'Custom
         Picture         =   "Frm68.frx":8CFE
         Style           =   1  'Graphical
         TabIndex        =   208
         Top             =   9840
         Width           =   2625
      End
      Begin VB.CommandButton CMD25 
         Caption         =   "Back"
         Height          =   810
         Left            =   12240
         MouseIcon       =   "Frm68.frx":9DC8
         MousePointer    =   99  'Custom
         Picture         =   "Frm68.frx":A0D2
         Style           =   1  'Graphical
         TabIndex        =   206
         ToolTipText     =   "Tutup senarai ini."
         Top             =   9600
         Width           =   1095
      End
      Begin VB.CommandButton CMD26 
         Caption         =   "Next"
         Height          =   810
         Left            =   13440
         MouseIcon       =   "Frm68.frx":B19C
         MousePointer    =   99  'Custom
         Picture         =   "Frm68.frx":B4A6
         Style           =   1  'Graphical
         TabIndex        =   205
         ToolTipText     =   "Tutup senarai ini."
         Top             =   9600
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   8955
         Left            =   120
         TabIndex        =   198
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   600
         Width           =   14445
         _ExtentX        =   25479
         _ExtentY        =   15796
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
      Begin VB.Label L10_Text 
         Height          =   8895
         Left            =   14760
         TabIndex        =   207
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Ringkasan bagi data keseluruhan."
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
         TabIndex        =   204
         Top             =   9600
         Width           =   4995
      End
      Begin VB.Label L62_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L62_Text"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   203
         Top             =   10080
         Width           =   2415
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Komisyen : "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   202
         Top             =   10080
         Width           =   1815
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan             :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   201
         Top             =   9840
         Width           =   1695
      End
      Begin VB.Label L61_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L61_Text"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1920
         TabIndex        =   200
         Top             =   9840
         Width           =   975
      End
      Begin VB.Label L4_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai komisyen bagi agen ini."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   199
         Top             =   360
         Width           =   9615
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Komisen Agen Dropship"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   11160
      TabIndex        =   186
      Top             =   240
      Visible         =   0   'False
      Width           =   7695
      Begin VB.CommandButton CMD4 
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
         Height          =   825
         Left            =   4560
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm68.frx":C570
         MousePointer    =   99  'Custom
         Picture         =   "Frm68.frx":C87A
         Style           =   1  'Graphical
         TabIndex        =   196
         Top             =   2160
         Width           =   1665
      End
      Begin VB.CommandButton CMD3 
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
         Height          =   825
         Left            =   2760
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm68.frx":D944
         MousePointer    =   99  'Custom
         Picture         =   "Frm68.frx":DC4E
         Style           =   1  'Graphical
         TabIndex        =   195
         Top             =   2160
         Width           =   1665
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   1920
         TabIndex        =   191
         Top             =   1320
         Width           =   5565
         _ExtentX        =   9816
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
         Format          =   120389632
         CurrentDate     =   41561
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   360
         Left            =   1920
         TabIndex        =   192
         Top             =   1680
         Width           =   5565
         _ExtentX        =   9816
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
         Format          =   120389632
         CurrentDate     =   41561
      End
      Begin VB.Label L7_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L7_Text"
         Height          =   255
         Left            =   360
         TabIndex        =   210
         Top             =   2520
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label L6_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L6_Text"
         Height          =   255
         Left            =   360
         TabIndex        =   209
         Top             =   2160
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label62 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Mula * :"
         Height          =   255
         Left            =   300
         TabIndex        =   194
         Top             =   1365
         Width           =   1575
      End
      Begin VB.Label Label63 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Akhir * :"
         Height          =   255
         Left            =   300
         TabIndex        =   193
         Top             =   1725
         Width           =   1575
      End
      Begin VB.Label L5_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L5_Text"
         Height          =   255
         Left            =   1920
         TabIndex        =   190
         Top             =   960
         Width           =   5655
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "        Nama Agen  :    No. Agen :"
         Height          =   855
         Left            =   180
         TabIndex        =   189
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label L42_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L42_Text"
         Height          =   255
         Left            =   1920
         TabIndex        =   188
         Top             =   720
         Width           =   5655
      End
      Begin VB.Label Label64 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila buat pilihan tarikh bagi melihat senarai komiyen bagi agen dropship ini."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   187
         Top             =   360
         Width           =   7335
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Database Pelanggan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11535
      Left            =   19200
      TabIndex        =   121
      Top             =   1440
      Visible         =   0   'False
      Width           =   20535
      Begin VB.CommandButton CMD21 
         Caption         =   "Back"
         Height          =   810
         Left            =   18120
         MouseIcon       =   "Frm68.frx":E5F8
         MousePointer    =   99  'Custom
         Picture         =   "Frm68.frx":E902
         Style           =   1  'Graphical
         TabIndex        =   124
         ToolTipText     =   "Tutup senarai ini."
         Top             =   10680
         Width           =   1095
      End
      Begin VB.CommandButton CMD22 
         Caption         =   "Next"
         Height          =   810
         Left            =   19320
         MouseIcon       =   "Frm68.frx":F9CC
         MousePointer    =   99  'Custom
         Picture         =   "Frm68.frx":FCD6
         Style           =   1  'Graphical
         TabIndex        =   123
         ToolTipText     =   "Tutup senarai ini."
         Top             =   10680
         Width           =   1095
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   10140
         Left            =   120
         TabIndex        =   122
         Top             =   480
         Width           =   20235
         _ExtentX        =   35692
         _ExtentY        =   17886
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
      Begin VB.Label L14_Text 
         BackColor       =   &H8000000A&
         Height          =   8775
         Left            =   19320
         TabIndex        =   135
         Top             =   6840
         Width           =   6975
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
         Left            =   17640
         TabIndex        =   133
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
         Left            =   17040
         TabIndex        =   132
         Top             =   10680
         Width           =   375
      End
      Begin VB.Label L69_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L69_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   6360
         TabIndex        =   131
         Top             =   10920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L70_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L70_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7320
         TabIndex        =   130
         Top             =   10920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L57_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L57_Text"
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
         Left            =   720
         TabIndex        =   128
         Top             =   10680
         Width           =   975
      End
      Begin VB.Label L72_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L72_Text"
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
         Left            =   4560
         TabIndex        =   127
         Top             =   10680
         Width           =   975
      End
      Begin VB.Label L73_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L73_Text"
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
         Left            =   8640
         TabIndex        =   126
         Top             =   10680
         Width           =   975
      End
      Begin VB.Label L71_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L71_Text"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   125
         Top             =   240
         Width           =   14175
      End
      Begin VB.Label Label65 
         BackStyle       =   0  'Transparent
         Caption         =   "Bil. :                                        Bilangan Aktif :                                  Bilangan Tidak Aktif : "
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
         Left            =   240
         TabIndex        =   129
         Top             =   10680
         Width           =   9135
      End
      Begin VB.Label Label31 
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
         Left            =   15720
         TabIndex        =   134
         Top             =   10680
         Width           =   2295
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Carian Maklumat Pelanggan "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3120
      TabIndex        =   113
      Top             =   840
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CommandButton CMD7 
         Caption         =   "Senarai Pelanggan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   2280
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm68.frx":10DA0
         MousePointer    =   99  'Custom
         Picture         =   "Frm68.frx":110AA
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   1200
         Width           =   2505
      End
      Begin VB.TextBox TB14 
         Height          =   360
         Left            =   1560
         TabIndex        =   115
         Text            =   "TB14"
         Top             =   720
         Width           =   4965
      End
      Begin VB.ComboBox CBB3 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   360
         ItemData        =   "Frm68.frx":12174
         Left            =   1560
         List            =   "Frm68.frx":12176
         Style           =   2  'Dropdown List
         TabIndex        =   114
         Top             =   360
         Width           =   4965
      End
      Begin VB.Label L39_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L39_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6600
         TabIndex        =   119
         Top             =   360
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label L40_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L40_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6600
         TabIndex        =   118
         Top             =   720
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label L38_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Carian * :"
         Height          =   255
         Left            =   240
         TabIndex        =   117
         Top             =   765
         UseMnemonic     =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label58 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Krateria * :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   116
         Top             =   375
         Width           =   1215
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Komisen Agen Dropship"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   720
      TabIndex        =   178
      Top             =   3360
      Visible         =   0   'False
      Width           =   8415
      Begin VB.CommandButton CMD30 
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
         Height          =   825
         Left            =   2280
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm68.frx":12178
         MousePointer    =   99  'Custom
         Picture         =   "Frm68.frx":12482
         Style           =   1  'Graphical
         TabIndex        =   185
         Top             =   1800
         Width           =   1665
      End
      Begin VB.CommandButton CMD31 
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
         Height          =   825
         Left            =   4080
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm68.frx":12E2C
         MousePointer    =   99  'Custom
         Picture         =   "Frm68.frx":13136
         Style           =   1  'Graphical
         TabIndex        =   184
         Top             =   1800
         Width           =   1665
      End
      Begin MSComCtl2.DTPicker DTPicker7 
         Height          =   360
         Left            =   1920
         TabIndex        =   180
         Top             =   960
         Width           =   5205
         _ExtentX        =   9181
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
         Format          =   166526976
         CurrentDate     =   41561
      End
      Begin MSComCtl2.DTPicker DTPicker8 
         Height          =   360
         Left            =   1920
         TabIndex        =   181
         Top             =   1320
         Width           =   5205
         _ExtentX        =   9181
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
         Format          =   166526976
         CurrentDate     =   41561
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Mula *    :"
         Height          =   255
         Left            =   360
         TabIndex        =   183
         Top             =   1005
         Width           =   1695
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Akhir *    :"
         Height          =   255
         Left            =   360
         TabIndex        =   182
         Top             =   1365
         Width           =   1695
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila buat pilihan tarikh bagi melihat senarai jualan yang dibuat oleh agen dropship."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   179
         Top             =   360
         Width           =   8535
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rekod Transaksi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   6360
      TabIndex        =   136
      Top             =   6600
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CommandButton CMD19 
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
         Height          =   825
         Left            =   4200
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm68.frx":14200
         MousePointer    =   99  'Custom
         Picture         =   "Frm68.frx":1450A
         Style           =   1  'Graphical
         TabIndex        =   143
         Top             =   1560
         Width           =   1665
      End
      Begin VB.CommandButton CMD18 
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
         Height          =   825
         Left            =   2400
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm68.frx":155D4
         MousePointer    =   99  'Custom
         Picture         =   "Frm68.frx":158DE
         Style           =   1  'Graphical
         TabIndex        =   142
         Top             =   1560
         Width           =   1665
      End
      Begin MSComCtl2.DTPicker DTPicker5 
         Height          =   360
         Left            =   1800
         TabIndex        =   138
         Top             =   720
         Width           =   5200
         _ExtentX        =   9181
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
         Format          =   166526976
         CurrentDate     =   41561
      End
      Begin MSComCtl2.DTPicker DTPicker6 
         Height          =   360
         Left            =   1800
         TabIndex        =   139
         Top             =   1080
         Width           =   5200
         _ExtentX        =   9181
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
         Format          =   166526976
         CurrentDate     =   41561
      End
      Begin VB.Label Label79 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Akhir * :"
         Height          =   255
         Left            =   360
         TabIndex        =   141
         Top             =   1125
         Width           =   1380
      End
      Begin VB.Label Label80 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Mula * :"
         Height          =   255
         Left            =   360
         TabIndex        =   140
         Top             =   765
         Width           =   1380
      End
      Begin VB.Label Label77 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila buat pilihan tarikh bagi melihat senarai rekod belian pelanggan ini."
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
         TabIndex        =   137
         Top             =   360
         Width           =   8295
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pendaftaran Pelanggan/Ahli"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   0
      TabIndex        =   106
      Top             =   360
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CommandButton CMD27 
         Caption         =   "Carian -> Teruskan ke menu pendaftaran"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   1440
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm68.frx":16288
         MousePointer    =   99  'Custom
         Picture         =   "Frm68.frx":16592
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2880
         Width           =   3465
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "No. Keahlian"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   110
         Top             =   1320
         Width           =   6135
         Begin VB.TextBox TB18 
            Height          =   360
            Left            =   1680
            MaxLength       =   9
            TabIndex        =   3
            Text            =   "TB18"
            Top             =   840
            Width           =   3075
         End
         Begin VB.Label L63_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Keahlian *   :"
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   885
            Width           =   1815
         End
         Begin VB.Label Label76 
            BackStyle       =   0  'Transparent
            Caption         =   "Sila masukkan No. Keahlian yang tercatat pada kad keahlian jika kedai menawarkan kad keahlian."
            Height          =   615
            Left            =   120
            TabIndex        =   111
            Top             =   240
            Width           =   5895
         End
      End
      Begin VB.CheckBox CB16 
         BackColor       =   &H8000000C&
         Caption         =   "Scanner Mode"
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
         Left            =   3720
         TabIndex        =   109
         Top             =   840
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.CheckBox CB15 
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
         Left            =   300
         TabIndex        =   1
         Top             =   765
         Width           =   200
      End
      Begin VB.CheckBox CB18 
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
         Left            =   300
         TabIndex        =   2
         Top             =   1005
         Width           =   200
      End
      Begin VB.Label Label69 
         BackStyle       =   0  'Transparent
         Caption         =   "Keahlian kedai                                    Pendaftaran pelanggan biasa"
         ForeColor       =   &H00000000&
         Height          =   765
         Left            =   555
         TabIndex        =   108
         Top             =   735
         Width           =   3345
      End
      Begin VB.Label Label68 
         BackStyle       =   0  'Transparent
         Caption         =   "Kad keahlian (Sila tanda di bawah jika kedai mempunyai kad keahlian)"
         Height          =   375
         Left            =   240
         TabIndex        =   107
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   6735
      End
   End
   Begin VB.PictureBox Pic8 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   9240
      ScaleHeight     =   5895
      ScaleWidth      =   9615
      TabIndex        =   31
      Top             =   240
      Visible         =   0   'False
      Width           =   9615
      Begin VB.CommandButton CMD16 
         Caption         =   "Simpan Data Simpanan"
         Height          =   375
         Left            =   1200
         MouseIcon       =   "Frm68.frx":1765C
         MousePointer    =   99  'Custom
         TabIndex        =   74
         Top             =   4080
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton CMD17 
         Caption         =   "Batal"
         Height          =   375
         Left            =   4800
         MouseIcon       =   "Frm68.frx":17966
         MousePointer    =   99  'Custom
         TabIndex        =   73
         Top             =   4080
         Visible         =   0   'False
         Width           =   1900
      End
      Begin VB.CommandButton CMD14 
         Caption         =   "Simpan Data Simpanan"
         Height          =   375
         Left            =   2520
         MouseIcon       =   "Frm68.frx":17C70
         MousePointer    =   99  'Custom
         TabIndex        =   72
         Top             =   4080
         Width           =   2175
      End
      Begin VB.CommandButton CMD15 
         Caption         =   "Batal"
         Height          =   375
         Left            =   4800
         MouseIcon       =   "Frm68.frx":17F7A
         MousePointer    =   99  'Custom
         TabIndex        =   71
         Top             =   4080
         Width           =   1900
      End
      Begin VB.ComboBox CBB1 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2700
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   3600
         Width           =   4000
      End
      Begin VB.TextBox TB17 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2700
         TabIndex        =   41
         Text            =   "TB17"
         Top             =   2880
         Width           =   4005
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   360
         Left            =   2700
         TabIndex        =   44
         Top             =   3240
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
         Format          =   166526976
         CurrentDate     =   41561
      End
      Begin VB.Label L28_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L28_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   49
         Top             =   5520
         Visible         =   0   'False
         Width           =   2000
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Maklumat Penyimpan"
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
         TabIndex        =   48
         Top             =   1080
         Width           =   6945
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Maklumat Simpanan"
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
         TabIndex        =   47
         Top             =   2520
         Width           =   6945
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pekerja *"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   46
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Simpanan  *"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   45
         Top             =   3240
         Width           =   2385
      End
      Begin VB.Label Label70 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Simpanan (RM)"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   42
         Top             =   2910
         Width           =   2715
      End
      Begin VB.Label L27_Text 
         Caption         =   "No. Rujukan :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5760
         TabIndex        =   40
         Top             =   840
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label L26_Text 
         Caption         =   "L26_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7080
         TabIndex        =   39
         Top             =   840
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label L19_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L19_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2880
         TabIndex        =   38
         Top             =   2160
         Width           =   8625
      End
      Begin VB.Label L18_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L18_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2880
         TabIndex        =   37
         Top             =   1920
         Width           =   8625
      End
      Begin VB.Label L17_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L17_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2880
         TabIndex        =   36
         Top             =   1680
         Width           =   8625
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Simpan Duit Di Kedai"
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
         Left            =   240
         TabIndex        =   35
         Top             =   480
         Width           =   7305
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm68.frx":18284
         ForeColor       =   &H00000000&
         Height          =   1005
         Left            =   240
         TabIndex        =   34
         Top             =   1440
         Width           =   2625
      End
      Begin VB.Label L16_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L16_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2880
         TabIndex        =   33
         Top             =   1440
         Width           =   8625
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   360
         Width           =   9615
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rekod Transaksi"
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
      Left            =   15240
      TabIndex        =   144
      Top             =   -2160
      Visible         =   0   'False
      Width           =   20535
      Begin VB.CommandButton CMD24 
         Caption         =   "Next"
         Height          =   810
         Left            =   19080
         MouseIcon       =   "Frm68.frx":1831C
         MousePointer    =   99  'Custom
         Picture         =   "Frm68.frx":18626
         Style           =   1  'Graphical
         TabIndex        =   177
         ToolTipText     =   "Tutup senarai ini."
         Top             =   9480
         Width           =   1095
      End
      Begin VB.CommandButton CMD23 
         Caption         =   "Back"
         Height          =   810
         Left            =   17880
         MouseIcon       =   "Frm68.frx":196F0
         MousePointer    =   99  'Custom
         Picture         =   "Frm68.frx":199FA
         Style           =   1  'Graphical
         TabIndex        =   176
         ToolTipText     =   "Tutup senarai ini."
         Top             =   9480
         Width           =   1095
      End
      Begin VB.CommandButton CMD20 
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
         Height          =   945
         Left            =   120
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm68.frx":1AAC4
         MousePointer    =   99  'Custom
         Picture         =   "Frm68.frx":1ADCE
         Style           =   1  'Graphical
         TabIndex        =   160
         Top             =   9840
         Width           =   2625
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid10 
         Height          =   8700
         Left            =   5280
         TabIndex        =   166
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   720
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   15346
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid9 
         Height          =   8700
         Left            =   5040
         TabIndex        =   167
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   720
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   15346
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid8 
         Height          =   8700
         Left            =   4800
         TabIndex        =   168
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   720
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   15346
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid7 
         Height          =   8700
         Left            =   4560
         TabIndex        =   169
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   720
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   15346
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid6 
         Height          =   8700
         Left            =   4200
         TabIndex        =   170
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   720
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   15346
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
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Bil. Data :                                                                                     Bil. Data :"
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
         Height          =   255
         Left            =   5280
         TabIndex        =   175
         Top             =   9720
         Width           =   8895
      End
      Begin VB.Label L59_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L59_Text"
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
         Height          =   255
         Left            =   10560
         TabIndex        =   174
         Top             =   9720
         Width           =   1575
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Ringkasan bagi data yang sedang dipaparkan.             Ringkasan bagi data keseluruhan."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   173
         Top             =   9480
         Width           =   9135
      End
      Begin VB.Label L58_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L58_Text"
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
         Height          =   255
         Left            =   6000
         TabIndex        =   172
         Top             =   9720
         Width           =   1575
      End
      Begin VB.Label L55_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L55_Text"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5280
         TabIndex        =   171
         Top             =   480
         Width           =   12255
      End
      Begin VB.Label L50_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Report Belian"
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
         Left            =   120
         MouseIcon       =   "Frm68.frx":1BE98
         MousePointer    =   99  'Custom
         TabIndex        =   165
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label L51_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Report Trade In /  Buyback"
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
         Left            =   120
         MouseIcon       =   "Frm68.frx":1C1A2
         MousePointer    =   99  'Custom
         TabIndex        =   164
         Top             =   3000
         Width           =   3135
      End
      Begin VB.Label L52_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Report Tempahan"
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
         Left            =   120
         MouseIcon       =   "Frm68.frx":1C4AC
         MousePointer    =   99  'Custom
         TabIndex        =   163
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label L53_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Report Ansuran"
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
         Left            =   120
         MouseIcon       =   "Frm68.frx":1C7B6
         MousePointer    =   99  'Custom
         TabIndex        =   162
         Top             =   4080
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label L54_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Report Servis"
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
         Left            =   120
         MouseIcon       =   "Frm68.frx":1CAC0
         MousePointer    =   99  'Custom
         TabIndex        =   161
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label L47_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L47_Text"
         Height          =   255
         Left            =   240
         TabIndex        =   159
         Top             =   8040
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label L48_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L48_Text"
         Height          =   255
         Left            =   240
         TabIndex        =   158
         Top             =   8400
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label L60_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L60_Text"
         Height          =   255
         Left            =   240
         TabIndex        =   157
         Top             =   8760
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label L56_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L56_Text"
         Height          =   255
         Left            =   240
         TabIndex        =   156
         Top             =   9120
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label66 
         BackStyle       =   0  'Transparent
         Caption         =   "Maklumat Pelanggan / Ahli."
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
         Left            =   120
         TabIndex        =   155
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label L43_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L43_Text"
         Height          =   255
         Left            =   2040
         TabIndex        =   154
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label Label71 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pelanggan      :"
         Height          =   255
         Left            =   120
         TabIndex        =   153
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label72 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Kad Pengenalan :"
         Height          =   255
         Left            =   120
         TabIndex        =   152
         Top             =   1125
         Width           =   1935
      End
      Begin VB.Label L44_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L44_Text"
         Height          =   255
         Left            =   2040
         TabIndex        =   151
         Top             =   1125
         Width           =   3855
      End
      Begin VB.Label Label74 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Telefon             :"
         Height          =   255
         Left            =   120
         TabIndex        =   150
         Top             =   1395
         Width           =   1935
      End
      Begin VB.Label L45_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L45_Text"
         Height          =   255
         Left            =   2040
         TabIndex        =   149
         Top             =   1395
         Width           =   3855
      End
      Begin VB.Label Label78 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Pelanggan         :"
         Height          =   255
         Left            =   120
         TabIndex        =   148
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label L46_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L46_Text"
         Height          =   255
         Left            =   2040
         TabIndex        =   147
         Top             =   1680
         Width           =   3855
      End
      Begin VB.Label Label60 
         BackStyle       =   0  'Transparent
         Caption         =   "Kategori Pelanggan  :"
         Height          =   255
         Left            =   120
         TabIndex        =   146
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label L49_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L49_Text"
         Height          =   255
         Left            =   2040
         TabIndex        =   145
         Top             =   840
         Width           =   3855
      End
   End
   Begin VB.PictureBox Pic9 
      BorderStyle     =   0  'None
      Height          =   9975
      Left            =   1200
      ScaleHeight     =   9975
      ScaleWidth      =   23535
      TabIndex        =   50
      Top             =   2400
      Visible         =   0   'False
      Width           =   23535
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
         Height          =   7875
         Left            =   240
         TabIndex        =   69
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   1680
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   13891
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid5 
         Height          =   7875
         Left            =   6600
         TabIndex        =   70
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   1680
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   13891
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
      Begin VB.Label L34_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L34_Text"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7680
         TabIndex        =   66
         Top             =   9600
         Width           =   2175
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Penggunaan : RM"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5400
         TabIndex        =   65
         Top             =   9600
         Width           =   2175
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Rekod perbelanjaan."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6840
         TabIndex        =   64
         Top             =   1320
         Width           =   9615
      End
      Begin VB.Label L32_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L32_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3120
         TabIndex        =   63
         Top             =   960
         Width           =   8625
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Pelanggan                     :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         TabIndex        =   62
         Top             =   960
         Width           =   2625
      End
      Begin VB.Label L29_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L29_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3120
         TabIndex        =   61
         Top             =   240
         Width           =   8625
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama                                 :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         TabIndex        =   60
         Top             =   240
         Width           =   2625
      End
      Begin VB.Label L30_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L30_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3120
         TabIndex        =   59
         Top             =   480
         Width           =   8625
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Kad Pengenalan             :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         TabIndex        =   58
         Top             =   480
         Width           =   2625
      End
      Begin VB.Label L31_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L31_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3120
         TabIndex        =   57
         Top             =   720
         Width           =   8625
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Telefon                         :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         TabIndex        =   56
         Top             =   720
         Width           =   2625
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "Rekod simpanan."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   53
         Top             =   1320
         Width           =   9615
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "Baki : RM"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10680
         TabIndex        =   52
         Top             =   9600
         Width           =   975
      End
      Begin VB.Label L35_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L35_Text"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   11520
         TabIndex        =   51
         Top             =   9600
         Width           =   2415
      End
      Begin VB.Label L33_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L33_Text"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2280
         TabIndex        =   55
         Top             =   9600
         Width           =   2175
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Simpanan : RM"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   9600
         Width           =   2175
      End
   End
   Begin VB.Label L13_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Komisen Agen Dropship"
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
      Left            =   5160
      MouseIcon       =   "Frm68.frx":1CDCA
      MousePointer    =   99  'Custom
      TabIndex        =   75
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label L37_Text 
      Alignment       =   2  'Center
      Caption         =   "L37_Text"
      Height          =   255
      Left            =   15360
      TabIndex        =   68
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label L36_Text 
      Alignment       =   2  'Center
      Caption         =   "L36_Text"
      Height          =   255
      Left            =   13920
      TabIndex        =   67
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label L15_Text 
      Alignment       =   2  'Center
      Caption         =   "L15_Text"
      Height          =   255
      Left            =   12360
      TabIndex        =   30
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label L12_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai Pelanggan"
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
      MouseIcon       =   "Frm68.frx":1D0D4
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label L22_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Menu Sebelum"
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
      Left            =   7560
      MouseIcon       =   "Frm68.frx":1D3DE
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label L20_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pendaftaran Data Pelanggan"
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
      MouseIcon       =   "Frm68.frx":1D6E8
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.Menu Frm68_PM_Menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm68_SM_LihatData 
         Caption         =   "Lihat data terperinci / Edit data"
      End
      Begin VB.Menu Frm68_SM_SenaraiKomisyen 
         Caption         =   "Lihat senarai komisyen Agen Dropship ini"
      End
   End
   Begin VB.Menu Frm68_PM_Menu2 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm68_DetailPaymentVoucher 
         Caption         =   "Lihat data terperinci berkenaan payment voucher ini"
         Visible         =   0   'False
      End
      Begin VB.Menu Frm68_SM_CetakPenyata 
         Caption         =   "Cetak penyata komisyen Agen Dropship ini"
      End
   End
   Begin VB.Menu Frm68_PM_Menu3 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm68_SM_LihatDataCust 
         Caption         =   "Lihat Data Terperinci Pelanggan"
         Visible         =   0   'False
      End
      Begin VB.Menu Frm68_SM_EditDataCust 
         Caption         =   "Edit Data Pelanggan"
      End
      Begin VB.Menu frm68_sm_spacer1 
         Caption         =   "-"
      End
      Begin VB.Menu Frm68_SM_padam_data 
         Caption         =   "Tukar Status Kepada Tidak Aktif"
      End
      Begin VB.Menu frm68_sm_spacer2 
         Caption         =   "-"
      End
      Begin VB.Menu Frm68_SM_DaftrarProgram 
         Caption         =   "Pendaftaran program baru bagi pelanggan ini"
         Visible         =   0   'False
      End
      Begin VB.Menu Frm68_SM_SenaraiProgram 
         Caption         =   "Senarai program yang disertai oleh pelanggan ini"
         Visible         =   0   'False
      End
      Begin VB.Menu Frm68_SM_SimpanDuit 
         Caption         =   "Simpanan / Pulangan / Rekod Simpanan && Penggunaan Duit"
      End
      Begin VB.Menu Frm68_SM_Rekod_Simpanan 
         Caption         =   "Rekod Simpanan Dan Penggunaan Duit Pelanngan Ini"
         Visible         =   0   'False
      End
      Begin VB.Menu Frm58_SM_rekod_belian 
         Caption         =   "Rekod Belian Pelanggan Ini"
      End
      Begin VB.Menu Frm68_SM_Select 
         Caption         =   "Pilih Data Pelanggan Ini (Pembeli)"
      End
      Begin VB.Menu Frm68_SM_Select_dropship 
         Caption         =   "Pilih Data Agen Dropship Ini"
      End
      Begin VB.Menu Frm68_SM_komisyen_dropship 
         Caption         =   "Maklumat Komisyen Agen Dropship"
      End
      Begin VB.Menu Frm68_SM_Reg_Belian_Hibah 
         Caption         =   "Pendaftaran Belian Berhibah"
         Visible         =   0   'False
      End
      Begin VB.Menu Frm68_SM_cetak_invoice2 
         Caption         =   "Cetak Invoice Yuran Pendaftaran"
      End
      Begin VB.Menu Frm68_SM_upgrade_pelanggan 
         Caption         =   "Upgrade Keahlian Kepada Berdaftar (Tidak Berdaftar -> Berdaftar)"
      End
      Begin VB.Menu Frm68_SM_tukar_no 
         Caption         =   "Tukar Nombor Keahlian"
      End
      Begin VB.Menu Frm68_SM_mata_ganjaran 
         Caption         =   "Lihat Mata Ganjaran Terperinci"
      End
      Begin VB.Menu frm68_sm_excel 
         Caption         =   "Export Excel"
      End
   End
   Begin VB.Menu Frm68_PM_Menu4 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm68_SM_Edit_Simpanan 
         Caption         =   "Edit Data Simpanan Ini"
      End
      Begin VB.Menu Frm68_SM_PadamDataIni 
         Caption         =   "Padam Data Simpanan Ini"
      End
   End
End
Attribute VB_Name = "Frm68"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CB10_Click()
'on error resume next
If Frm68.CB10 = 1 Then
    Frm68.CB9 = 0
    Frm68.CB11 = 0
    Frm68.CB12 = 0
End If
End Sub
Private Sub CB11_Click()
'on error resume next
If Frm68.CB11 = 1 Then
    Frm68.CB10 = 0
    Frm68.CB9 = 0
    Frm68.CB12 = 0
End If
End Sub
Private Sub CB12_Click()
'on error resume next
If Frm68.CB12 = 1 Then
    Frm68.CB10 = 0
    Frm68.CB11 = 0
    Frm68.CB9 = 0
End If
End Sub
Private Sub CB15_Click()
'on error resume next
If Frm68.CB15 = 1 Then
    Frm68.CB16 = 0
    Frm68.CB18 = 0
    Frm68.L63_Text = "No. Keahlian *   :"
    
    Frm68.Frame3.Visible = True
    
    If Frm68.Frame3.Visible = True Then Frm68.TB18.SetFocus
End If
End Sub
Private Sub CB16_Click()
'on error resume next
If Frm68.CB16 = 1 Then
    Frm68.CB15 = 0
    Frm68.CB18 = 0
    Frm68.L63_Text = "No. Keahlian     :"
End If
End Sub
Private Sub CB18_Click()
'on error resume next
If Frm68.CB18 = 1 Then
    Frm68.CB15 = 0
    Frm68.CB16 = 0
    Frm68.L63_Text = "No. Keahlian     :"
    Frm68.Frame3.Visible = False
End If
End Sub
Private Sub CB19_Click()
'on error resume next
If Frm68.CB19 = 1 Then
    Frm68.CB20 = 0
    Frm68.L65_Text = "Jumlah yuran pendaftaran *       :"
End If
End Sub
Private Sub CB20_Click()
'on error resume next
If Frm68.CB20 = 1 Then
    Frm68.CB19 = 0
    Frm68.L65_Text = "Jumlah yuran pendaftaran          :"
End If
End Sub
Private Sub CB9_Click()
'on error resume next
If Frm68.CB9 = 1 Then
    Frm68.CB10 = 0
    Frm68.CB11 = 0
    Frm68.CB12 = 0
End If
End Sub

Private Sub CBB3_Change()
'On Error Resume Next
If Frm68.CBB3 = "Semua senarai" Or Frm68.CBB3 = "Semua pelanggan biasa" Or Frm68.CBB3 = "Semua ahli biasa" Or Frm68.CBB3 = "Semua silver" Or Frm68.CBB3 = "Semua gold" Or Frm68.CBB3 = "Semua platinum" Or Frm68.CBB3 = "Semua agen dropship" Then
    Frm68.TB14.Visible = False
    Frm68.L38_Text.Visible = False
Else
    Frm68.TB14.Visible = True
    Frm68.L38_Text.Visible = True
    Frm68.TB14.SetFocus
End If
End Sub

Private Sub CBB3_Click()
'On Error Resume Next
If Frm68.CBB3 = "Semua senarai" Or Frm68.CBB3 = "Semua pelanggan biasa" Or Frm68.CBB3 = "Semua ahli biasa" Or Frm68.CBB3 = "Semua silver" Or Frm68.CBB3 = "Semua gold" Or Frm68.CBB3 = "Semua platinum" Or Frm68.CBB3 = "Semua agen dropship" Then
    Frm68.TB14.Visible = False
    Frm68.L38_Text.Visible = False
Else
    Frm68.TB14.Visible = True
    Frm68.L38_Text.Visible = True
    Frm68.TB14.SetFocus
End If
End Sub

Private Sub CMD1_Click()
'on error resume next
Dim myAt As Integer
Dim myDot As Integer
Dim myDotDot As Integer
Dim Err(10)

DATA_WRITE = 0 '0 : Tiada Data Disimpan , 1 : Data Telah Disimpan
Frm68_LM_No_PELANGGAN = 0 'No. Giliran Pelanggan
Frm68_LM_INVOICE_AHLI = 1
Frm68_LM_ACTIVE = 0

If Frm68.TB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan [Nama]."
End If
If Frm68.TB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan [No. Kad Pengenalan]."
End If
'If Frm68.TB12 = vbNullString Then
'    X = X + 1
'    Err(X) = "Tiada Maklumat [No. Agen / Staff]."
'End If
If Frm68.TB4 <> vbNullString Then
    myAt = InStr(1, Frm68.TB4, "@", vbTextCompare)
    myDot = InStr(myAt + 2, Frm68.TB4, ".", vbTextCompare)
    myDotDot = InStr(myAt + 2, Frm68.TB4, "..", vbTextCompare)
    
    If myAt = 0 Or myDot = 0 Or Not myDotDot = 0 Or Right(Frm68.TB4, 1) = "." Then
        x = x + 1
        Err(x) = "Email Yang Tidak Sah"
    End If
End If
If Frm68.CB19 = 0 And Frm68.CB20 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan samada ada bayaran bagi yuran pendaftaran atau tidak."
End If
If Frm68.CB19 = 1 Then
    If Frm68.TB19 = vbNullString Or (Frm68.TB19 <> vbNullString And Not IsNumeric(Frm68.TB19)) Then
        x = x + 1
        Err(x) = "Sila masukkkan [Yuran Pendaftaran]. Hanya NOMBOR dibenarkan di dalam ruangan ini."
    End If
End If
If Frm68.CB9.Enabled = True Then
    If Frm68.CB9 = 0 And Frm68.CB10 = 0 And Frm68.CB11 = 0 And Frm68.CB12 = 0 Then
        x = x + 1
        Err(x) = "Sila buat pilihan kategori pelanggan."
    End If
    If Frm68.TB12 = vbNullString Then
        x = x + 1
        Err(x) = "Tiada Maklumat [No. Agen / Staff]."
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
    
'### Periksa kewujudan NO KAD PENGENALAN ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from senarai_pelanggan where no_ic='" & UCase(Frm68.TB2) & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        Frm68_LM_KATEGORI = "Pelanggan Biasa"
        If Not IsNull(rs!Nama) Then Frm68_LM_NAMA = rs!Nama 'Nama
        If Not IsNull(rs!no_ic) Then Frm68_LM_No_IC = rs!no_ic 'No. IC
        If Not IsNull(rs!no_pelanggan) Then Frm68_LM_No_PELANGGAN = rs!no_pelanggan 'No. Pelanggan
        If Not IsNull(rs!kategori_pelanggan) Then
            If rs!kategori_pelanggan = 1 Then Frm68_LM_KATEGORI = "Pelanggan Biasa"
            If rs!kategori_pelanggan = 2 Then Frm68_LM_KATEGORI = "Ahli Biasa"
            If rs!kategori_pelanggan = 3 Then Frm68_LM_KATEGORI = "Silver"
            If rs!kategori_pelanggan = 4 Then Frm68_LM_KATEGORI = "Gold"
            If rs!kategori_pelanggan = 5 Then Frm68_LM_KATEGORI = "Platinum"
        End If
        
        MsgBox "Pelanggan dengan No. Kad Pengenalan [" & Frm68_LM_No_IC & "] telah didaftarkan sebelum ini." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Maklumat yang telah didaftarkan adalah seperti berikut :" & vbCrLf & _
                "Nama : " & Frm68_LM_NAMA & vbCrLf & _
                "No. Kad Pengenalan : " & Frm68_LM_No_IC & vbCrLf & _
                "No. Pelanggan : " & Frm68_LM_No_PELANGGAN & vbCrLf & _
                "Kategori : " & Frm68_LM_KATEGORI & vbCrLf & _
                vbNullString, vbExclamation, "Info"
        
        rs.Close
        Set rs = Nothing
        
        Exit Sub
        
    End If
    
    rs.Close
    Set rs = Nothing
'### Periksa kewujudan NO KAD PENGENALAN ### - End

'### Periksa No. Keahlian telah digunakan atau belum ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & UCase(Frm68.TB18) & "' AND membership_card = 1", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        If Not IsNull(rs!Nama) Then Frm68_LM_NAMA = rs!Nama 'Nama
        If Not IsNull(rs!no_pelanggan) Then Frm68_LM_No_KEAHLIAN = rs!no_pelanggan 'No. Keahlian
        If Not IsNull(rs!no_ic) Then Frm68_LM_No_IC = rs!no_ic 'No. Kad Pengenalan
    
        MsgBox "No. keahlian [" & Frm68_LM_No_KEAHLIAN & "] telah didaftarkan/digunakan sebelum ini!" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Maklumat bagi nombor keahlian ini adalah seperti di bawah :" & vbCrLf & _
                "Nama : " & Frm68_LM_NAMA & vbCrLf & _
                "No. Kad Pengenalan : " & Frm68_LM_No_IC, vbExclamation, "Info"
            
        rs.Close
        Set rs = Nothing
            
        Exit Sub
        
    End If
    
    rs.Close
    Set rs = Nothing
'### Periksa No. Keahlian telah digunakan atau belum ### - Start
    
'### Periksa apakah jenis kategori pendaftaran ### - Start

    If Frm68.CB9 = 0 And Frm68.CB10 = 0 And Frm68.CB11 = 0 And Frm68.CB12 = 0 Then
    
        Frm68_LM_KATEGORI = "Pelanggan Biasa"
        Frm68_LM_CODE = "C"
        
    Else
    
'Frm68_LM_CODE
'C : Pelanggan Biasa
'A : Member / Ahli
'D : Dealer / Pengedar
'R : RAF
'N : Normal Dealer
'M : Master Dealer

        If Frm68.CB9 = 1 Then
            Frm68_LM_KATEGORI = "Ahli Biasa"
        End If
        If Frm68.CB10 = 1 Then
            Frm68_LM_KATEGORI = "Silver"
        End If
        If Frm68.CB11 = 1 Then
            Frm68_LM_KATEGORI = "Gold"
        End If
        If Frm68.CB12 = 1 Then
            Frm68_LM_KATEGORI = "Platinum"
        End If
        
    End If
'### Periksa apakah jenis kategori pendaftaran ### - End
    
    If Frm68.CB17 = 0 Then
    
        'Note = "Adakah anda ingin mendaftarkan pelanggan ini ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Maklumat pelanggan adalah seperti di bawah :" & vbCrLf & _
                "Nama : " & UCase(Frm68.TB1) & vbCrLf & _
                "No. Kad Pengenalan : " & UCase(Frm68.TB2) & vbCrLf & _
                "Kategori : " & Frm68_LM_KATEGORI & vbCrLf & _
                "No. Pelanggan : " & Frm68_LM_KOD_KEDAI & Frm68_LM_CODE & Format(Frm68_LM_No_PELANGGAN, "00000") & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan ?"
                
        Note = "Adakah anda ingin mendaftarkan pelanggan ini ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Maklumat pelanggan adalah seperti di bawah :" & vbCrLf & _
                "Nama : " & UCase(Frm68.TB1) & vbCrLf & _
                "No. Kad Pengenalan : " & UCase(Frm68.TB2) & vbCrLf & _
                "Kategori : " & Frm68_LM_KATEGORI & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan ?"
                
    ElseIf Frm68.CB17 = 1 Then
    
        Note = "Adakah anda ingin mendaftarkan pelanggan ini ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Maklumat pelanggan adalah seperti di bawah :" & vbCrLf & _
                "Nama : " & UCase(Frm68.TB1) & vbCrLf & _
                "No. Kad Pengenalan : " & UCase(Frm68.TB2) & vbCrLf & _
                "Kategori : " & Frm68_LM_KATEGORI & vbCrLf & _
                "No. Keahlian : " & Frm68.TB12 & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan ?"
    
    End If
    
    If Frm68.CB14 = 1 Then 'Agen Dropship
    
        If Frm68.CB17 = 0 Then
        
            'Note = "Adakah anda ingin mendaftarkan pelanggan ini ?" & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Maklumat pelanggan adalah seperti di bawah :" & vbCrLf & _
                    "Nama : " & UCase(Frm68.TB1) & vbCrLf & _
                    "No. Kad Pengenalan : " & UCase(Frm68.TB2) & vbCrLf & _
                    "Kategori : " & Frm68_LM_KATEGORI & vbCrLf & _
                    "No. Pelanggan : " & Frm68_LM_KOD_KEDAI & Frm68_LM_CODE & Format(Frm68_LM_No_PELANGGAN, "00000") & vbCrLf & _
                    "***Pelanggan ini akan didaftarkan sebagai agen dropship***" & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Teruskan ?"
                    
            Note = "Adakah anda ingin mendaftarkan pelanggan ini ?" & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Maklumat pelanggan adalah seperti di bawah :" & vbCrLf & _
                    "Nama : " & UCase(Frm68.TB1) & vbCrLf & _
                    "No. Kad Pengenalan : " & UCase(Frm68.TB2) & vbCrLf & _
                    "Kategori : " & Frm68_LM_KATEGORI & vbCrLf & _
                    "***Pelanggan ini akan didaftarkan sebagai agen dropship***" & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Teruskan ?"
                    
        ElseIf Frm68.CB17 = 1 Then
        
            Note = "Adakah anda ingin mendaftarkan pelanggan ini ?" & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Maklumat pelanggan adalah seperti di bawah :" & vbCrLf & _
                    "Nama : " & UCase(Frm68.TB1) & vbCrLf & _
                    "No. Kad Pengenalan : " & UCase(Frm68.TB2) & vbCrLf & _
                    "Kategori : " & Frm68_LM_KATEGORI & vbCrLf & _
                    "No. Pelanggan : " & Frm68.TB12 & vbCrLf & _
                    "***Pelanggan ini akan didaftarkan sebagai agen dropship***" & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Teruskan ?"
                    
        End If
                
    End If
            
        
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
'---------------------------------------No. Invoice
        LM_NOW = Now
        
        If Frm68.CB9 = 0 And Frm68.CB10 = 0 And Frm68.CB11 = 0 And Frm68.CB12 = 0 Then
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
            strsql = "insert into 13_rujukan_customer(tarikh,terminal,write_timestamp,Status,nama_staff,cawangan)" & _
                        "select '" & Frm68.DTPicker4 & "','" & G_TERMINAL & "','" & LM_NOW & "',1,'" & MDI_frm1.L3_Text & "','" & G_CAWANGAN & "'"
            
            Set rs = cn2.Execute(strsql)
            Set rs = Nothing
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
            rs.Open "select * from 13_rujukan_customer where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND cawangan='" & G_CAWANGAN & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & Frm68.DTPicker4 & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic

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
'---------------------------------------No. Invoice
        End If
        
        If Frm68.CB19 = 1 Then
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
            strsql = "insert into 15_senarai_invoice_member(tarikh,terminal,write_timestamp,Status,nama_staff,cawangan)" & _
                        "select '" & Frm68.DTPicker4 & "','" & G_TERMINAL & "','" & LM_NOW & "',1,'" & MDI_frm1.L3_Text & "','" & G_CAWANGAN & "'"
            
            Set rs = cn2.Execute(strsql)
            Set rs = Nothing
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
            rs.Open "select * from 15_senarai_invoice_member where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND cawangan='" & G_CAWANGAN & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & Frm68.DTPicker4 & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic

            If Not rs.EOF Then
                If Not IsNull(rs!ID) Then
                    rs!no_invoice = "MEM" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(rs!ID, "000000")
                    Frm68_LM_No_RESIT_JUALAN = rs!ID 'No. Rujukan Belian
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
        End If
'---------------------------------------No. Invoice

'### Simpan data pelanggan ke dalam database ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where no_ic='" & UCase(Frm68.TB2) & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
        
            rs.AddNew
            If Frm68.CB9 = 0 And Frm68.CB10 = 0 And Frm68.CB11 = 0 And Frm68.CB12 = 0 Then
                rs!kategori_pelanggan = 1
                rs!no_pelanggan = LM_NO_CUSTOMER
            Else
                If Frm68.CB9 = 1 Then rs!kategori_pelanggan = 2 'Ahli Biasa
                If Frm68.CB10 = 1 Then rs!kategori_pelanggan = 3 'Silver
                If Frm68.CB11 = 1 Then rs!kategori_pelanggan = 4 'Gold
                If Frm68.CB12 = 1 Then rs!kategori_pelanggan = 5 'Platinum
                rs!no_pelanggan = UCase(Frm68.TB12) 'No. Pelanggan
                LM_NO_CUSTOMER = UCase(Frm68.TB12)
            End If
            If Frm68.TB1 <> vbNullString Then
                rs!Nama = UCase(Frm68.TB1) 'Nama
            Else
                rs!Nama = Null
            End If
            If Frm68.TB2 <> vbNullString Then
                rs!no_ic = UCase(Frm68.TB2) 'No. IC
            Else
                rs!no_ic = Null
            End If
            'If Frm68.TB12 <> vbNullString Then
            '    rs!no_pelanggan = UCase(Frm68.TB12) 'No. Pelanggan
            'Else
            '    rs!no_pelanggan = Null
            'End If
            If Frm68.TB3 <> vbNullString Then
                rs!no_tel = UCase(Frm68.TB3) 'No. Tel
            Else
                rs!no_tel = Null
            End If
            If Frm68.TB4 <> vbNullString Then
                rs!Email = Frm68.TB4 'E-mail
            Else
                rs!Email = Null
            End If
            If Frm68.TB5 <> vbNullString Then
                rs!alamat = UCase(Frm68.TB5) 'Alamat
            Else
                rs!alamat = Null
            End If
            If Frm68.TB6 <> vbNullString Then
                rs!Nama_Waris = UCase(Frm68.TB6) 'Nama Waris
            Else
                rs!Nama_Waris = Null
            End If
            If Frm68.TB7 <> vbNullString Then
                rs!No_Tel_Waris = UCase(Frm68.TB7) 'No. Tel Waris
            Else
                rs!No_Tel_Waris = Null
            End If
            If Frm68.TB8 <> vbNullString Then
                rs!alamat_waris = UCase(Frm68.TB8) 'Alamat Waris
            Else
                rs!alamat_waris = Null
            End If
            If Frm68.TB9 <> vbNullString Then
                rs!nama_bank = UCase(Frm68.TB9) 'Nama Waris
            Else
                rs!nama_bank = Null
            End If
            If Frm68.TB10 <> vbNullString Then
                rs!nama_akaun = UCase(Frm68.TB10) 'Nama Akaun
            Else
                rs!nama_akaun = Null
            End If
            If Frm68.TB11 <> vbNullString Then
                rs!no_akaun = UCase(Frm68.TB11) 'No. Akaun
            Else
                rs!no_akaun = Null
            End If
            If Frm68.CB14 = 1 Then
                rs!dropship = 1 '0 : Bukan agen dropship , 1 : Agen dropship
            Else
                rs!dropship = 0 '0 : Bukan agen dropship , 1 : Agen dropship
            End If
            rs!baki_simpanan = "0.00" 'Baki Simpan Di Kedai
            If Frm68.CB17 = 0 Then
                rs!membership_card = 0 '0 : Tiada kad keahlian , 1 : Ada kad keahlian
            Else
                rs!membership_card = 1 '0 : Tiada kad keahlian , 1 : Ada kad keahlian
            End If
            If Frm68.CB19 = 1 Then
                rs!yuran_flag = 1 'Flag samada ada bayaran yang dikenakan bagi pendaftaran ini atau tidak (0 : Tiada bayaran , 1 : Ada bayaran)
                
                If Frm68.TB19 <> vbNullString Then 'Jumlah bayaran yuran yang dikenakan (RM)
                    rs!jumlah_yuran = Format(Frm68.TB19, "0.00")
                Else
                    rs!jumlah_yuran = "0.00"
                End If
                'If Frm68_LM_INVOICE_AHLI <> vbNullString Then 'No. Invoice Yuran Keahlian
                    rs!no_invoice = "MEM" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm68_LM_No_RESIT_JUALAN, "000000")
                'End If
                'If Frm68.CB13 = 0 Then
                '    rs!no_invoice = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm68_LM_No_RESIT_JUALAN, "000000")
                '    rs!bil_rasmi = 0
                'End If
                'If Frm68.CB13 = 1 Then
                '    rs!no_invoice = "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm68_LM_No_RESIT_JUALAN, "000000")
                '    rs!bil_rasmi = 1
                'End If
            End If
            
            If Frm68.CB20 = 1 Then
                rs!yuran_flag = 0 'Flag samada ada bayaran yang dikenakan bagi pendaftaran ini atau tidak (0 : Tiada bayaran , 1 : Ada bayaran)
                rs!jumlah_yuran = "0.00"
                rs!no_invoice = Null
            End If
            If Frm68.DTPicker4 <> vbNullString Then
                rs!tarikh = Frm68.DTPicker4 'Tarikh pendaftaran
            Else
                rs!tarikh = Null
            End If
            rs!Status = 1 '0 : Sudah dipadamkan , 1 : Aktif , 2 : Tidak aktif
            rs!cawangan = G_CAWANGAN
            G_KEDAI = G_CAWANGAN
            rs!write_timestamp = LM_NOW 'Tarikh Data Dimasukkan
            DATA_WRITE = 1 '0 : Tiada Data Disimpan , 1 : Data Telah Disimpan , 2 : Data Telah Diedit
            rs.Update
            
        Else
        
            MsgBox "Pengguna dengan No. Kad Pengenalan " & UCase(Frm68.TB2) & " telah didaftarkan sebelum ini." & vbCrLf & _
                    "Sila periksa senarai pelanggan ini.", vbInformation, "Info"
            
        End If
        
        rs.Close
        Set rs = Nothing
'### Simpan data pelanggan ke dalam database ### - End
        
        If DATA_WRITE = 1 Then  '0 : Tiada Data Disimpan , 1 : Data Customer Baru Telah Disimpan , 2 : Data Belian Telah Disimpan
            
            DATA_UPDATE = 0

            G_INVOICE_AHLI = "MEM" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm68_LM_No_RESIT_JUALAN, "000000")
            
            Frm68.Frame1.Visible = True
            Frm68.CMD1.Visible = True
            
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & user & "] Pendaftaran pelanggan baru. IC [" & UCase(Frm68.TB2) & "] , No. Pelanggan [" & LM_NO_CUSTOMER & "]"
            LogDate_Memory = LM_NOW
            Call UpdateLog_Database
            
            If Frm68.CB19 = 1 Then
                Call Frm68_invoice_yuran_ahli
            End If
                    
            Call Frm68_Reset_All
        
            Frm68.Frame2.Visible = True
            
            MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
        End If
    End If
End If
End Sub
Private Sub CMD14_Click()
'On Error Resume Next
Dim Err(5)
Dim Frm68_LM_SIMPANAN_ASAL As Double
Dim Frm68_LM_SIMPANAN_BARU As Double

DATA_SAVE = 0
Frm68_LM_SIMPANAN_ASAL = 0
Frm68_LM_SIMPANAN_BARU = 0

If Frm68.L19_Text = vbNullString Then
    x = x + 1
    Err(x) = "Tiada Maklumat Pelanggan. Sila Keluar Dari Menu Ini Dan Cuba Sekali Lagi."
End If
If Frm68.TB17 = vbNullString Or (Frm68.TB17 <> vbNullString And Not IsNumeric(Frm68.TB17)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Jumlah Simpanan]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
End If
If Frm68.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Nama Pekerja]."
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
    
        Frm68_LM_No_RUJUKAN = Frm68.L26_Text 'No. Rujukan
        
Re_Gen_No_Rujukan:
'###Masukkan Data Belian Ke Dalam Database Akaun Belian### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 24_rekod_kewangan_pelanggan where no_resit='" & "SAV" & Format(Frm68_LM_No_RUJUKAN, "000000") & "' AND jenis='" & "0" & "'", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
            rs.AddNew
            If Frm68.DTPicker3 <> vbNullString Then
                rs!tarikh = Frm68.DTPicker3 'Tarikh
            Else
                rs!tarikh = Null 'Tarikh
            End If
            rs!jenis = 0 '0 : Simpanan , 1 : Penggunaan Duit
            If Frm68.L19_Text <> vbNullString Then
                rs!no_rujukan_pelanggan = Frm68.L19_Text 'No. Rujukan Pelanggan
            Else
                rs!no_rujukan_pelanggan = Null 'No. Rujukan Pelanggan
            End If
            rs!no_resit = "SAV" & Format(Frm68_LM_No_RUJUKAN, "000000") 'No. Rujukan
            If Frm68.TB17 <> vbNullString Then
                rs!jumlah = Format(Frm68.TB17, "0.00") 'Jumlah Simpanan / Penggunaan (RM)
                Frm68_LM_SIMPANAN_BARU = Frm68.TB17 'Simpanan Yang Baru (RM)
            Else
                rs!jumlah = Null 'Jumlah Simpanan / Penggunaan (RM)
            End If
            If Frm68.CBB1 <> vbNullString Then
                Frm68_LM_EMP_NO = Split(Frm68.CBB1, "  |  ")(1)
                Frm68_LM_EMP_NAMA = Split(Frm68.CBB1, "  |  ")(0)
                rs!no_rujukan_pekerja = Frm68_LM_EMP_NO 'No. Pekerja
            End If
            rs!cawangan = G_CAWANGAN
            DATA_SAVE = 1
            rs.Update
        Else
            Frm68_LM_No_RUJUKAN = Frm68_LM_No_RUJUKAN + 1
            Frm68.L26_Text = Frm68_LM_No_RUJUKAN 'No. Rujukan
            
            rs.Close
            Set rs = Nothing
            GoTo Re_Gen_No_Rujukan:
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_SAVE = 1 Then
            If IsNumeric(Frm68.L28_Text) Then Frm68_LM_SIMPANAN_ASAL = Frm68.L28_Text 'Baki Simpanan Asal (RM)
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm68.L19_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
            If Not rs.EOF Then
                rs!baki_simpanan = Format(Frm68_LM_SIMPANAN_ASAL + Frm68_LM_SIMPANAN_BARU, "0.00") 'Baki Simpanan Terbaru (RM)
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
        
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & Frm68_LM_EMP_NAMA & "] Simpanan Duit Di Kedai , No Rujukan [" & "SAV" & Format(Frm68_LM_No_RUJUKAN, "000000") & "]."
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If rs!Default1 = "Default" Then
                    If IsNumeric(Frm68.L26_Text) Then rs!no_resit_simpanan = Frm68.L26_Text + 1 'No. Resit Simpanan
                    rs.Update
                End If
            End If
            
            rs.Close
            Set rs = Nothing
            
            Frm68.Pic8.Visible = False
            MsgBox "Data Telah Berjaya Disimpan", vbInformation, "Info"
        End If
    End If
End If
End Sub
Private Sub CMD15_Click()
'on error resume next
Note = "Adakah Anda Ingin Batalkan Urusan Ini ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    Frm68.Pic8.Visible = False
End If
End Sub
Private Sub CMD16_Click()
'On Error Resume Next
Dim Err(5)
Dim Frm68_LM_SIMPANAN_ASAL As Double
Dim Frm68_LM_SIMPANAN_BARU As Double
Dim Frm68_LM_SIMPANAN_LAMA As Double

DATA_SAVE = 0
Frm68_LM_SIMPANAN_ASAL = 0
Frm68_LM_SIMPANAN_BARU = 0
Frm68_LM_SIMPANAN_LAMA = 0

If Frm68.L19_Text = vbNullString Then
    x = x + 1
    Err(x) = "Tiada Maklumat Pelanggan. Sila Keluar Dari Menu Ini Dan Cuba Sekali Lagi."
End If
If Frm68.TB17 = vbNullString Or (Frm68.TB17 <> vbNullString And Not IsNumeric(Frm68.TB17)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Jumlah Simpanan]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
End If
If Frm68.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Nama Pekerja]."
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
        
'###Masukkan Data Belian Ke Dalam Database Akaun Belian### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 24_rekod_kewangan_pelanggan where no_resit='" & Frm68.L26_Text & "' AND jenis='" & "0" & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!jumlah) Then
                Frm68_LM_SIMPANAN_LAMA = rs!jumlah
            End If
            If Frm68.DTPicker3 <> vbNullString Then
                rs!tarikh = Frm68.DTPicker3 'Tarikh
            Else
                rs!tarikh = Null 'Tarikh
            End If
            rs!jenis = 0 '0 : Simpanan , 1 : Penggunaan Duit
            If Frm68.L19_Text <> vbNullString Then
                rs!no_rujukan_pelanggan = Frm68.L19_Text 'No. Rujukan Pelanggan
            Else
                rs!no_rujukan_pelanggan = Null 'No. Rujukan Pelanggan
            End If
            rs!no_resit = Frm68.L26_Text 'No. Rujukan
            If Frm68.TB17 <> vbNullString Then
                rs!jumlah = Format(Frm68.TB17, "0.00") 'Jumlah Simpanan / Penggunaan (RM)
                Frm68_LM_SIMPANAN_BARU = Frm68.TB17 'Simpanan Yang Baru (RM)
            Else
                rs!jumlah = Null 'Jumlah Simpanan / Penggunaan (RM)
            End If
            If Frm68.CBB1 <> vbNullString Then
                Frm68_LM_EMP_NO = Split(Frm68.CBB1, "  |  ")(1)
                Frm68_LM_EMP_NAMA = Split(Frm68.CBB1, "  |  ")(0)
                rs!no_rujukan_pekerja = Frm68_LM_EMP_NO 'No. Pekerja
            End If
            DATA_SAVE = 1
            rs.Update
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_SAVE = 1 Then
            If IsNumeric(Frm68.L28_Text) Then Frm68_LM_SIMPANAN_ASAL = Frm68.L28_Text 'Baki Simpanan Asal (RM)
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm68.L19_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
            If Not rs.EOF Then
                rs!baki_simpanan = Format(Frm68_LM_SIMPANAN_ASAL - Frm68_LM_SIMPANAN_LAMA + Frm68_LM_SIMPANAN_BARU, "0.00") 'Baki Simpanan Terbaru (RM)
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
        
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & Frm68_LM_EMP_NAMA & "] Edit Data Simpanan Duit Di Kedai , No Rujukan [" & Frm68.L26_Text & "]."
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
            
            Frm68.Pic8.Visible = False
            MsgBox "Data Telah Berjaya Disimpan", vbInformation, "Info"
        End If
    End If
End If
End Sub
Private Sub CMD17_Click()
'on error resume next
Note = "Adakah Anda Ingin Batalkan Urusan Edit Ini ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    Frm68.Pic8.Visible = False
End If
End Sub
Private Sub CMD18_Click()
'on error resume next
Note = "Paparan rekod belian oleh " & Frm68.L43_Text & " dari " & Frm68.DTPicker5 & " hingga " & Frm68.DTPicker6 & vbCrLf & _
        vbNullString & vbCrLf & _
        "Sistem mungkin akan mengambil sedikit masa untuk memaparkan semua rekod ini." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then
    Frm68.L47_Text = Frm68.DTPicker5
    Frm68.L48_Text = Frm68.DTPicker6
    
    Frm68.L56_Text = -1
    Frm68.L60_Text = 0 '0 : Rekod belian , 1 : Rekod buyback , 2 : Rekod tempahan , 3 : Rekod ansuran , 4 : Rekod servis
    GM_NEXT_PREV = 0
    
    Call Frm68_report_belian_header
    Call Frm68_report_belian_page
    'Call Frm68_report_buyback
    'Call Frm68_report_tempahan
    'Call Frm68_report_ansuran
    'Call Frm68_rekod_servis
    
    Frm68.Frame7.Visible = True
    Frm68.Frame6.Visible = False
    
    Frm68.MSFlexGrid6.Visible = True
    Frm68.L55_Text = "Report belian dari " & Frm68.L47_Text & " hingga " & Frm68.L48_Text
End If
End Sub
Private Sub CMD19_Click()
'on error resume next
Frm68.Frame5.Visible = True
Frm68.Frame6.Visible = False
End Sub
Private Sub CMD2_Click()
'on error resume next
Dim myAt As Integer
Dim myDot As Integer
Dim myDotDot As Integer
Dim Err(10)

DATA_FOUND = 0
DATA_WRITE = 0 '0 : Tiada Data Disimpan , 1 : Data Telah Disimpan
Frm68_LM_KATEGORI = 0
Frm68_LM_No_AHLI_ASAL = vbNullString
Frm68_LM_YURAN = 0 '0: Tiada bayaran yuran , 1 : Ada bayaran yuran
Frm68_LM_INVOICE_AHLI = 1

If Frm68.TB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan [Nama]."
End If
If Frm68.TB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan [No. Kad Pengenalan]."
End If
If Frm68.TB12 = vbNullString Then
    x = x + 1
    Err(x) = "Tiada Maklumat [No. Agen / Staff]."
End If
If Frm68.TB4 <> vbNullString Then
    myAt = InStr(1, Frm68.TB4, "@", vbTextCompare)
    myDot = InStr(myAt + 2, Frm68.TB4, ".", vbTextCompare)
    myDotDot = InStr(myAt + 2, Frm68.TB4, "..", vbTextCompare)
    
    If myAt = 0 Or myDot = 0 Or Not myDotDot = 0 Or Right(Frm68.TB4, 1) = "." Then
        x = x + 1
        Err(x) = "Email Yang Tidak Sah"
    End If
End If
If Frm68.CB19 = 0 And Frm68.CB20 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan samada ada bayaran bagi yuran pendaftaran atau tidak."
End If
If Frm68.CB19 = 1 Then
    If Frm68.TB19 = vbNullString Or (Frm68.TB19 <> vbNullString And Not IsNumeric(Frm68.TB19)) Then
        x = x + 1
        Err(x) = "Sila masukkkan [Yuran Pendaftaran]. Hanya NOMBOR dibenarkan di dalam ruangan ini."
    End If
End If
If Frm68.CB9.Enabled = True Then
    If Frm68.CB9 = 0 And Frm68.CB10 = 0 And Frm68.CB11 = 0 And Frm68.CB12 = 0 Then
        x = x + 1
        Err(x) = "Sila buat pilihan kategori pelanggan."
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
    Note = "Adakah anda ingin simpan data yang telah diedit?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
        
        'If Frm68.TB2.Locked = False Then
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where NO_IC='" & Frm68.TB2 & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If rs!ID <> Frm68.L66_Text Then
                MsgBox "Pelanggan dengan No. Kad Pengenalan [" & UCase(Frm68.TB2) & "] telah didaftarkan sebelum ini!" & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "No. Keahlian/pelanggan yang telah didaftarkan adalah : " & rs!no_pelanggan & vbCrLf & _
                        "Sila periksa data anda" & vbCrLf & _
                        vbNullString, vbExclamation, "Info"
                
                rs.Close
                Set rs = Nothing
                
                Exit Sub
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        
        'End If
        
'### Periksa nombor keahlian asal pelanggan ini ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where ID='" & Frm68.L66_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!yuran_flag) Then
            
                If rs!yuran_flag = 0 Then
                    Frm68_LM_YURAN = 0 '0: Tiada bayaran yuran , 1 : Ada bayaran yuran
                ElseIf rs!yuran_flag = 1 Then
                    Frm68_LM_YURAN = 1 '0: Tiada bayaran yuran , 1 : Ada bayaran yuran
                End If
                
            Else
                Frm68_LM_YURAN = 0 '0: Tiada bayaran yuran , 1 : Ada bayaran yuran
            End If
            
            If Not IsNull(rs!no_pelanggan) Then
                Frm68_LM_No_AHLI_ASAL = rs!no_pelanggan 'No. Pelanggan
                DATA_FOUND = 1
            End If
            
        End If
        
        rs.Close
        Set rs = Nothing
'### Periksa nombor keahlian asal pelanggan ini ### - End
        
        LM_NOW = Now
        
'### No. invoice terbaru ### - Start
        If Frm68_LM_YURAN = 0 And Frm68.CB19 = 1 Then  '0: Tiada bayaran yuran , 1 : Ada bayaran yuran
            
            GoTo skip_this:
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If rs!Default1 = "Default" Then
                
                        If Not IsNull(rs!no_invoice_membership) Then
                            Frm68_LM_INVOICE_AHLI = rs!no_invoice_membership 'No. Invoice Yuran Keahlian
                        End If
                End If
            End If
            
            rs.Close
            Set rs = Nothing
            
'### Periksa no invoice bayaran keahlian ### - Start
            If Frm68.CB17 = 0 Then
            
re_gen_number2:
        
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from senarai_pelanggan where no_invoice='" & "MEM" & Format(Frm68_LM_INVOICE_AHLI, "000000") & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    Frm68_LM_INVOICE_AHLI = Frm68_LM_INVOICE_AHLI + 1
                    
                    rs.Close
                    Set rs = Nothing
                    
                    GoTo re_gen_number2:
                End If
                
                rs.Close
                Set rs = Nothing
                
            End If
'### Periksa no invoice bayaran keahlian ### - End

skip_this:

            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
            rs.Open "select * from 15_senarai_invoice_member", cn2, adOpenKeyset, adLockOptimistic

            rs.AddNew
            rs!tarikh = Frm68.DTPicker4
            rs!terminal = G_TERMINAL
            rs!write_timestamp = LM_NOW
            rs!Status = 1
            rs!nama_staff = MDI_frm1.L3_Text
            rs!cawangan = G_CAWANGAN
            rs.Update
            
            rs.Close
            Set rs = Nothing
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
            rs.Open "select * from 15_senarai_invoice_member where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND cawangan='" & G_CAWANGAN & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & Frm68.DTPicker4 & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic

            If Not rs.EOF Then
            
                If Not IsNull(rs!ID) Then
                    
                    rs!no_invoice = "MEM" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(rs!ID, "000000")
                    Frm68_LM_No_RESIT_JUALAN = rs!ID 'No. Rujukan Belian
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
            
        End If
'### No. invoice terbaru ### - End

        If DATA_FOUND = 1 Then
    
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from senarai_pelanggan where ID='" & Frm68.L66_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
            If Not rs.EOF Then

                If Frm68.CB9 = 0 And Frm68.CB10 = 0 And Frm68.CB11 = 0 And Frm68.CB12 = 0 Then
                    rs!kategori_pelanggan = 1
                Else
                    If Frm68.CB9 = 1 Then rs!kategori_pelanggan = 2 'Ahli Biasa
                    If Frm68.CB10 = 1 Then rs!kategori_pelanggan = 3 'Silver
                    If Frm68.CB11 = 1 Then rs!kategori_pelanggan = 4 'Gold
                    If Frm68.CB12 = 1 Then rs!kategori_pelanggan = 5 'Platinum
                End If
            
                If Frm68.TB1 <> vbNullString Then
                    rs!Nama = UCase(Frm68.TB1) 'Nama
                Else
                    rs!Nama = Null
                End If
                If Frm68.TB2 <> vbNullString Then
                    rs!no_ic = UCase(Frm68.TB2) 'No. IC
                Else
                    rs!no_ic = Null
                End If
                If Frm68.TB12 <> vbNullString Then
                    rs!no_pelanggan = UCase(Frm68.TB12) 'No. Pelanggan
                Else
                    rs!no_pelanggan = Null
                End If
                If Frm68.TB3 <> vbNullString Then
                    rs!no_tel = UCase(Frm68.TB3) 'No. Tel
                Else
                    rs!no_tel = Null
                End If
                If Frm68.TB4 <> vbNullString Then
                    rs!Email = Frm68.TB4 'E-mail
                Else
                    rs!Email = Null
                End If
                If Frm68.TB5 <> vbNullString Then
                    rs!alamat = UCase(Frm68.TB5) 'Alamat
                Else
                    rs!alamat = Null
                End If
                If Frm68.TB6 <> vbNullString Then
                    rs!Nama_Waris = UCase(Frm68.TB6) 'Nama Waris
                Else
                    rs!Nama_Waris = Null
                End If
                If Frm68.TB7 <> vbNullString Then
                    rs!No_Tel_Waris = UCase(Frm68.TB7) 'No. Tel Waris
                Else
                    rs!No_Tel_Waris = Null
                End If
                If Frm68.TB8 <> vbNullString Then
                    rs!alamat_waris = UCase(Frm68.TB8) 'Alamat Waris
                Else
                    rs!alamat_waris = Null
                End If
                If Frm68.TB9 <> vbNullString Then
                    rs!nama_bank = UCase(Frm68.TB9) 'Nama Waris
                Else
                    rs!nama_bank = Null
                End If
                If Frm68.TB10 <> vbNullString Then
                    rs!nama_akaun = UCase(Frm68.TB10) 'Nama Akaun
                Else
                    rs!nama_akaun = Null
                End If
                If Frm68.TB11 <> vbNullString Then
                    rs!no_akaun = UCase(Frm68.TB11) 'No. Akaun
                Else
                    rs!no_akaun = Null
                End If
                If Frm68.CB14 = 1 Then
                    rs!dropship = 1 '0 : Bukan agen dropship , 1 : Agen dropship
                Else
                    rs!dropship = 0 '0 : Bukan agen dropship , 1 : Agen dropship
                End If
                If Frm68.CB17 = 0 Then
                    rs!membership_card = 0 '0 : Tiada kad keahlian , 1 : Ada kad keahlian
                Else
                    rs!membership_card = 1 '0 : Tiada kad keahlian , 1 : Ada kad keahlian
                End If
                If Frm68.CB19 = 1 Then
                    rs!yuran_flag = 1 'Flag samada ada bayaran yang dikenakan bagi pendaftaran ini atau tidak (0 : Tiada bayaran , 1 : Ada bayaran)
                    
                    If Frm68.TB19 <> vbNullString Then 'Jumlah bayaran yuran yang dikenakan (RM)
                        rs!jumlah_yuran = Format(Frm68.TB19, "0.00")
                    Else
                        rs!jumlah_yuran = "0.00"
                    End If
                    If Frm68_LM_YURAN = 0 Then '0: Tiada bayaran yuran , 1 : Ada bayaran yuran
                    
                        'If Frm68_LM_INVOICE_AHLI <> vbNullString Then 'No. Invoice Yuran Keahlian
                            rs!no_invoice = "MEM" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm68_LM_No_RESIT_JUALAN, "000000")
                        'End If
                        
                        'If Frm68.CB13 = 0 Then
                        '    rs!no_invoice = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm68_LM_No_RESIT_JUALAN, "000000")
                        '    rs!bil_rasmi = 0
                        'End If
                        'If Frm68.CB13 = 1 Then
                        '    rs!no_invoice = "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm68_LM_No_RESIT_JUALAN, "000000")
                        '    rs!bil_rasmi = 1
                        'End If
                        
                    End If
                End If
                If Frm68.CB20 = 1 Then
                    rs!yuran_flag = 0 'Flag samada ada bayaran yang dikenakan bagi pendaftaran ini atau tidak (0 : Tiada bayaran , 1 : Ada bayaran)
                    rs!jumlah_yuran = "0.00"
                    rs!no_invoice = Null
                End If
                If Frm68.DTPicker4 <> vbNullString Then
                    rs!tarikh = Frm68.DTPicker4 'Tarikh pendaftaran
                Else
                    rs!tarikh = Null
                End If
                
                rs!write_timestamp2 = LM_NOW 'Tarikh Data Dimasukkan
                DATA_WRITE = 1 '0 : Tiada Data Disimpan , 1 : Data Telah Disimpan , 2 : Data Telah Diedit
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
        
            If DATA_WRITE = 1 Then  '0 : Tiada Data Disimpan , 1 : Data Customer Baru Telah Disimpan , 2 : Data Belian Telah Disimpan
            
'#### Update maklumat di bawah dalam setiap table dalam database #### - Start
    
                '#### Update maklumat pelanggan dalam #16_gold_bar_belian #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 16_gold_bar_belian set no_rujukan_pelanggan_buyback='" & UCase(Frm68.TB12) & "'," _
                & "kategori_penjual='" & Frm68_LM_KATEGORI & "'" _
                & "WHERE no_rujukan_pelanggan_buyback='" & Frm68_LM_No_AHLI_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat pelanggan dalam #16_gold_bar_belian #### - End
                
                '#### Update maklumat pelanggan dalam #22_jualan #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 22_jualan set no_rujukan_pembeli='" & UCase(Frm68.TB12) & "'," _
                & "kategori_pembeli='" & Frm68_LM_KATEGORI & "'" _
                & "WHERE no_rujukan_pembeli='" & Frm68_LM_No_AHLI_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat pelanggan dalam #22_jualan #### - End
                
                '#### Update maklumat pelanggan dalam #23_senarai_jualan #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 23_senarai_jualan set no_rujukan_pembeli='" & UCase(Frm68.TB12) & "'," _
                & "kategori_pembeli='" & Frm68_LM_KATEGORI & "'" _
                & "WHERE no_rujukan_pembeli='" & Frm68_LM_No_AHLI_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat pelanggan dalam #23_senarai_jualan #### - End
                
                '#### Update maklumat agen dropship dalam #23_senarai_jualan #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 23_senarai_jualan set no_rujukan_agen_dropship='" & UCase(Frm68.TB12) & "'" _
                & "WHERE no_rujukan_agen_dropship='" & Frm68_LM_No_AHLI_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat agen dropship dalam #23_senarai_jualan #### - End
                
                '#### Update maklumat agen dropship dalam #24_rekod_kewangan_pelanggan #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 24_rekod_kewangan_pelanggan set no_rujukan_pelanggan='" & UCase(Frm68.TB12) & "'" _
                & "WHERE no_rujukan_pelanggan='" & Frm68_LM_No_AHLI_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat agen dropship dalam #24_rekod_kewangan_pelanggan #### - End
                
                '#### Update maklumat agen dropship dalam #27_senarai_ansuran #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 27_senarai_ansuran set no_rujukan_pelanggan='" & UCase(Frm68.TB12) & "'," _
                & "kategori_pembeli='" & Frm68_LM_KATEGORI & "'" _
                & "WHERE no_rujukan_pelanggan='" & Frm68_LM_No_AHLI_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat agen dropship dalam #27_senarai_ansuran #### - End
                
                '#### Update maklumat agen dropship dalam #29_akaun_ansuran #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 29_akaun_ansuran set no_rujukan_pembeli='" & UCase(Frm68.TB12) & "'," _
                & "kategori_pembeli='" & Frm68_LM_KATEGORI & "'" _
                & "WHERE no_rujukan_pembeli='" & Frm68_LM_No_AHLI_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat agen dropship dalam #29_akaun_ansuran #### - End
                
                '#### Update maklumat agen dropship dalam #35_senarai_servis #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 35_senarai_servis set no_pelanggan='" & UCase(Frm68.TB12) & "'" _
                & "WHERE no_pelanggan='" & Frm68_LM_No_AHLI_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat agen dropship dalam #35_senarai_servis #### - End

                '#### Update maklumat pelanggan dalam #36_akaun_servis #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 36_akaun_servis set no_rujukan_pembeli='" & UCase(Frm68.TB12) & "'," _
                & "kategori_pembeli='" & Frm68_LM_KATEGORI & "'" _
                & "WHERE no_rujukan_pembeli='" & Frm68_LM_No_AHLI_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat pelanggan dalam #36_akaun_servis #### - End
                
                '#### Update maklumat pelanggan dalam #40_tempahan_deposit #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 40_tempahan_deposit set no_rujukan_pelanggan='" & UCase(Frm68.TB12) & "'," _
                & "kategori_pembeli='" & Frm68_LM_KATEGORI & "'" _
                & "WHERE no_rujukan_pelanggan='" & Frm68_LM_No_AHLI_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat pelanggan dalam #40_tempahan_deposit #### - End
                
                '#### Update maklumat pelanggan dalam #42_tempahan_siap #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 42_tempahan_siap set no_rujukan_pelanggan='" & UCase(Frm68.TB12) & "'," _
                & "kategori_pembeli='" & Frm68_LM_KATEGORI & "'" _
                & "WHERE no_rujukan_pelanggan='" & Frm68_LM_No_AHLI_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat pelanggan dalam #42_tempahan_siap #### - End
                
                '#### Update maklumat agen dropship dalam #43_bonus_ahli #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 43_bonus_ahli set no_ahli='" & UCase(Frm68.TB12) & "'" _
                '& "WHERE no_ahli='" & Frm68_LM_No_AHLI_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat agen dropship dalam #43_bonus_ahli #### - End
                
                '#### Update maklumat agen dropship dalam #data_database #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE data_database set no_rujukan_pelanggan_buyback='" & UCase(Frm68.TB12) & "'" _
                & "WHERE no_rujukan_pelanggan_buyback='" & Frm68_LM_No_AHLI_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat agen dropship dalam #data_database #### - End
                
'#### Update No. Invoice bayaran yuran #### - Start
                'If Frm68_LM_YURAN = 0 Then '0: Tiada bayaran yuran , 1 : Ada bayaran yuran
                '    If Frm68.CB19 = 1 Then
        
                '        Set rs = New ADODB.Recordset
                '        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                '        rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
                        
                '        If Not rs.EOF Then
                '            If rs!Default1 = "Default" Then
                '                If Not IsNull(rs!no_invoice_membership) Then
                '                    Frm68_LM_No_INVOICE_AHLI = rs!no_invoice_membership
                '                    rs!no_invoice_membership = Frm68_LM_No_INVOICE_AHLI + 1
                '                    rs.Update
                '                End If
                '            End If
                '        End If
                        
                '        rs.Close
                '        Set rs = Nothing
                        
                '    End If
                'End If
'#### Update No. Invoice bayaran yuran #### - End
                
'#### Update maklumat di bawah dalam setiap table dalam database #### - End
                'If Frm68.TB2.Locked = False Then
                
                    user = MDI_frm1.L3_Text
                    LogAct_Memory = "[" & user & "] Edit Data Pelanggan. IC [" & UCase(Frm68.TB2) & "] , No. Pelanggan [" & UCase(Frm68.TB12) & "]"
                    LogDate_Memory = LM_NOW
                    Call UpdateLog_Database
                    
                'End If
                
                'If Frm68.TB2.Locked = True Then

                '    User = MDI_frm1.L3_Text
                '    LogAct_Memory = "[" & User & "] Upgrade data pelanggan dari tidak berdaftar kepada berdaftar. IC [" & UCase(Frm68.TB2) & "] , No. Pelanggan [" & UCase(Frm68.TB12) & "]"
                '    LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                '    Call UpdateLog_Database
                    
                'End If
                
                If Frm68_LM_YURAN = 0 Then '0: Tiada bayaran yuran , 1 : Ada bayaran yuran
                    If Frm68.CB19 = 1 Then
                        'If G_INVOICE_AHLI <> vbNullString Then 'No. Invoice Yuran Keahlian
                            'G_INVOICE_AHLI = "MEM" & Format(Frm68_LM_INVOICE_AHLI, "000000") 'No. Invoice Yuran Keahlian
                            
                            G_INVOICE_AHLI = "MEM" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm68_LM_No_RESIT_JUALAN, "000000")
                            'If Frm68.CB13 = 1 Then G_INVOICE_AHLI = "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm68_LM_No_RESIT_JUALAN, "000000")
                        
                            Call Frm68_invoice_yuran_ahli
                        'End If
                    End If
                End If

                GM_NEXT_PREV = 2

                Call frm68_senarai_pelanggan_header
                Call frm68_senarai_pelanggan
                
                Frm68.Frame5.Visible = True
                Frm68.Frame1.Visible = False
                
                MsgBox "Data telah berjaya disimpan.", vbInformation, "Info"
                
            End If
            
        End If
        
    End If
End If
End Sub
Private Sub CMD20_Click()
'On Error Resume Next
Frm68.Frame7.Visible = False
Frm68.Frame6.Visible = True
End Sub
Private Sub CMD21_Click()
'On Error Resume Next
Dim Frm68_LM_CURR_PAGE As Double
Dim Frm68_LM_TOTAL_PAGE As Double

Frm68_LM_CURR_PAGE = 0
Frm68_LM_TOTAL_PAGE = 0

If Frm68.L67_Text <> vbNullString And IsNumeric(Frm68.L67_Text) Then
    If Frm68.L68_Text <> vbNullString And IsNumeric(Frm68.L68_Text) Then
        Frm68_LM_CURR_PAGE = Frm68.L67_Text
        Frm68_LM_TOTAL_PAGE = Frm68.L68_Text
        
        If Frm68_LM_CURR_PAGE <> 1 And Frm68_LM_CURR_PAGE <> 0 Then
        
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
                        
            Call frm68_senarai_pelanggan_header
            Call frm68_senarai_pelanggan
                        
        End If
    End If
End If
End Sub
Private Sub CMD22_Click()
'on error resume next
Dim Frm68_LM_CURR_PAGE As Double
Dim Frm68_LM_TOTAL_PAGE As Double

Frm68_LM_CURR_PAGE = 0
Frm68_LM_TOTAL_PAGE = 0

If Frm68.L67_Text <> vbNullString And IsNumeric(Frm68.L67_Text) Then
    If Frm68.L68_Text <> vbNullString And IsNumeric(Frm68.L68_Text) Then
        Frm68_LM_CURR_PAGE = Frm68.L67_Text
        Frm68_LM_TOTAL_PAGE = Frm68.L68_Text
        
        If Frm68_LM_CURR_PAGE < Frm68_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
              
            Call frm68_senarai_pelanggan_header
            Call frm68_senarai_pelanggan
            
        End If
    End If
End If
End Sub
Private Sub CMD23_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

If Frm68.L60_Text = 0 Then '0 : Rekod belian , 1 : Rekod buyback , 2 : Rekod tempahan , 3 : Rekod ansuran , 4 : Rekod servis
    Call Frm68_report_belian_header
    Call Frm68_report_belian_page
ElseIf Frm68.L60_Text = 1 Then '0 : Rekod belian , 1 : Rekod buyback , 2 : Rekod tempahan , 3 : Rekod ansuran , 4 : Rekod servis
    Call Frm68_report_buyback_header
    Call Frm68_report_buyback_page
ElseIf Frm68.L60_Text = 2 Then '0 : Rekod belian , 1 : Rekod buyback , 2 : Rekod tempahan , 3 : Rekod ansuran , 4 : Rekod servis
    Call Frm68_report_tempahan_header
    Call Frm68_report_tempahan_page
ElseIf Frm68.L60_Text = 3 Then '0 : Rekod belian , 1 : Rekod buyback , 2 : Rekod tempahan , 3 : Rekod ansuran , 4 : Rekod servis
    Call Frm68_report_ansuran_header
    Call Frm68_report_ansuran_page
ElseIf Frm68.L60_Text = 4 Then '0 : Rekod belian , 1 : Rekod buyback , 2 : Rekod tempahan , 3 : Rekod ansuran , 4 : Rekod servis
    Call Frm68_rekod_servis_header
    Call Frm68_rekod_servis_page
End If
End Sub
Private Sub CMD24_Click()
'on error resume next
GM_NEXT_PREV = 0 '0 : Next , 1 : Previous

If Frm68.L60_Text = 0 Then '0 : Rekod belian , 1 : Rekod buyback , 2 : Rekod tempahan , 3 : Rekod ansuran , 4 : Rekod servis
    Call Frm68_report_belian_header
    Call Frm68_report_belian_page
ElseIf Frm68.L60_Text = 1 Then '0 : Rekod belian , 1 : Rekod buyback , 2 : Rekod tempahan , 3 : Rekod ansuran , 4 : Rekod servis
    Call Frm68_report_buyback_header
    Call Frm68_report_buyback_page
ElseIf Frm68.L60_Text = 2 Then '0 : Rekod belian , 1 : Rekod buyback , 2 : Rekod tempahan , 3 : Rekod ansuran , 4 : Rekod servis
    Call Frm68_report_tempahan_header
    Call Frm68_report_tempahan_page
ElseIf Frm68.L60_Text = 3 Then '0 : Rekod belian , 1 : Rekod buyback , 2 : Rekod tempahan , 3 : Rekod ansuran , 4 : Rekod servis
    Call Frm68_report_ansuran_header
    Call Frm68_report_ansuran_page
ElseIf Frm68.L60_Text = 4 Then '0 : Rekod belian , 1 : Rekod buyback , 2 : Rekod tempahan , 3 : Rekod ansuran , 4 : Rekod servis
    Call Frm68_rekod_servis_header
    Call Frm68_rekod_servis_page
End If
End Sub
Private Sub CMD25_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm68_SenaraiKomisyen_Header
Call Frm68_SenaraiKomisyen_page
End Sub
Private Sub CMD26_Click()
'on error resume next
GM_NEXT_PREV = 0 '0 : Next , 1 : Previous

Call Frm68_SenaraiKomisyen_Header
Call Frm68_SenaraiKomisyen_page
End Sub
Private Sub CMD27_Click()
'On Error Resume Next
Dim Frm68_LM_LEN As String

If Frm68.TB18 <> vbNullString Then

    If InStr(1, Frm68.TB18, "*") <> 0 Or InStr(1, Frm68.TB18, ".") <> 0 Or InStr(1, Frm68.TB18, "/") <> 0 Or InStr(1, Frm68.TB18, "\") <> 0 Or InStr(1, Frm68.TB18, "'") <> 0 Then
        MsgBox "No. Keahlian mengandungi simbol yang tidak sah.", vbExclamation, "Info"
        Frm68.TB18.SetFocus
        
        Exit Sub
    End If
    
End If

If Frm68.CB15 = 1 Then
    
    If Frm68.TB18 = vbNullString Then
    
        MsgBox "Sila masukkan/scan No. Keahlian yang tercatat pada kad keahlian bagi pendaftaran ahli.", vbExclamation, "Info"
        
        Frm68.TB18.SetFocus
        
        Exit Sub
    End If

    'Set rs = New ADODB.Recordset
    'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    'rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
    
    'If Not rs.EOF Then
    '    If rs!Default1 = "Default" Then
    '        If Not IsNull(rs!membership_len) Then Frm68_LM_LEN = rs!membership_len 'Panjang nombor keahlian
    '    End If
    'End If
    
    'rs.Close
    'Set rs = Nothing
    
'If Len(G_MODE) = 0 Or Len(G_MIN_LEN) = 0 Or Len(G_MAX_LEN) = 0 Or Len(G_CODE) = 0 Then

'### Periksa panjang no keahlian
    Frm68_LM_LEN = Len(Frm68.TB18)
    
    If Frm68_LM_LEN < G_MIN_LEN Or Frm68_LM_LEN > G_MAX_LEN Then
    
        MsgBox "Sila periksa No. Keahlian. No. Keahlian tidak menepati panjang No. Keahlian yang ditetapkan." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Panjang bagi nombor keahlian yang ditetapkan adalah : " & vbCrLf & _
                "Minimum abjad   : " & G_MIN_LEN & vbCrLf & _
                "Maksimum abjad : " & G_MAX_LEN, vbExclamation, "Info"
        
        Frm68.TB18.SetFocus
        
        Exit Sub
        
    End If
    
'### Periksa kod kedai
    If InStr(1, UCase(Frm68.TB18), G_CODE) = 0 Then
    
        MsgBox "No. keahlian tidak mengandungi kod kedai." & vbCrLf & _
                "Kod kedai ialah : " & G_CODE, vbInformation, "Info"
                
        Frm68.TB18.SetFocus
        
        Exit Sub
        
    End If
    
'### Periksa kedudukan kod kedai
    L_CODE_COORDINATE = 0
    
    L_CODE_COORDINATE = InStr(1, UCase(Frm68.TB18), G_CODE)
    
    If L_CODE_COORDINATE <> 0 Then
        
        If L_CODE_COORDINATE > 1 Then
        
            MsgBox "No. keahlian tidak mengandungi kod kedai." & vbCrLf & _
                    "Kod kedai ialah : " & G_CODE & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Contoh : " & G_CODE & "002000", vbInformation, "Info"
                    
            Frm68.TB18.SetFocus
            
            Exit Sub
            
        End If
    End If
    
'### Periksa kategori keahlian
    'LM_KATEGORI = vbNullString

    'LM_KATEGORI = Right(UCase(Frm68.TB18), 1)
    
    'LM_KATEGORI_FOUND = 0
    'If LM_KATEGORI = "N" Then LM_KATEGORI_FOUND = 1
    'If LM_KATEGORI = "S" Then LM_KATEGORI_FOUND = 1
    'If LM_KATEGORI = "G" Then LM_KATEGORI_FOUND = 1
    'If LM_KATEGORI = "P" Then LM_KATEGORI_FOUND = 1
    
    'If LM_KATEGORI_FOUND = 0 Then
    
    '    MsgBox "Jenis kad yang tidak wujud." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Jenis kad yang ditawarkan adalah seperti di bawah :" & vbCrLf & _
                "1) Biasa " & vbCrLf & _
                "2) Silver " & vbCrLf & _
                "3) Gold " & vbCrLf & _
                "4) Platinum", vbExclamation, "Info"
        
    '    Exit Sub
        
    'End If
    
'### Periksa no keahlian
    LM_NO_GILIRAN_1 = Right(UCase(Frm68.TB18), 6)
    'LM_NO_GILIRAN = Left(LM_NO_GILIRAN_1, 6)
    
    If Not IsNumeric(LM_NO_GILIRAN_1) Then
        
        MsgBox "No. Keahlian yang tidak sah.", vbExclamation, "Info"
        
        Frm68.TB18.SetFocus
        
        Exit Sub
    
    End If
    
' ### Periksa No. Keahlian telah digunakan atau belum ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & UCase(Frm68.TB18) & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        If Not IsNull(rs!Nama) Then Frm68_LM_NAMA = rs!Nama 'Nama
        If Not IsNull(rs!no_pelanggan) Then Frm68_LM_No_KEAHLIAN = rs!no_pelanggan 'No. Keahlian
        If Not IsNull(rs!no_ic) Then Frm68_LM_No_IC = rs!no_ic 'No. Kad Pengenalan
    
        MsgBox "No. keahlian [" & Frm68_LM_No_KEAHLIAN & "] telah didaftarkan/digunakan sebelum ini!" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Maklumat bagi nombor keahlian ini adalah seperti di bawah :" & vbCrLf & _
                "Nama : " & Frm68_LM_NAMA & vbCrLf & _
                "No. Kad Pengenalan : " & Frm68_LM_No_IC, vbExclamation, "Info"
            
        rs.Close
        Set rs = Nothing
         
        Frm68.TB18.SetFocus
        
        Exit Sub
        
    End If
    
    rs.Close
    Set rs = Nothing
' ### Periksa No. Keahlian telah digunakan atau belum ### - Start

    Dim LM_NO_GILIRAN_AHLI As Long
    
    LM_NO_GILIRAN_AHLI = LM_NO_GILIRAN
    
    'Frm68.CB9.Enabled = False
    'Frm68.CB10.Enabled = False
    'Frm68.CB11.Enabled = False
    'Frm68.CB12.Enabled = False
    
'### Periksa no keahlian mengikut kategori - Start

    Frm68.CB9 = 1
    Frm68.CB10 = 0
    Frm68.CB11 = 0
    Frm68.CB12 = 0
    
    Frm68.CB9.Enabled = True
    Frm68.CB10.Enabled = True
    Frm68.CB11.Enabled = True
    Frm68.CB12.Enabled = True
                
    GoTo skip_kategori:
    If LM_KATEGORI = "N" Then
        If 1000 <= LM_NO_GILIRAN_AHLI And LM_NO_GILIRAN_AHLI <= 1999 Then
            
            Note = "Kategori pelanggan yang cuba didaftarkan adalah AHLI BIASA dengan nombor keahlian " & UCase(Frm68.TB18) & "." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Adakah anda yakin untuk teruskan pendaftaran ini?"
            
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbNo Then
                Exit Sub
            End If
            If Answer = vbYes Then
                Frm68.CB9 = 1
                Frm68.CB10 = 0
                Frm68.CB11 = 0
                Frm68.CB12 = 0
            End If
            
        Else
        
            MsgBox "No. Keahlian yang tidak sah bagi keahlian biasa." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Turutan nombor keahlian bagi AHLI BIASA ialah 1000 ~ 1999" & vbCrLf & _
            vbNullString & vbCrLf & _
            "Sila periksa nombor keahlian yang tercatat di belakang kad ini.", vbInformation, "Info"
            
            Frm68.TB18.SetFocus
            
            Exit Sub
            
        End If
        
    End If
    
    If LM_KATEGORI = "S" Then
    
        If 2000 <= LM_NO_GILIRAN_AHLI And LM_NO_GILIRAN_AHLI <= 2999 Then
            
            Note = "Kategori pelanggan yang cuba didaftarkan adalah SILVER dengan nombor keahlian " & UCase(Frm68.TB18) & "." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Adakah anda yakin untuk teruskan pendaftaran ini?"
            
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbNo Then
                Exit Sub
            End If
            If Answer = vbYes Then
                Frm68.CB9 = 0
                Frm68.CB10 = 1
                Frm68.CB11 = 0
                Frm68.CB12 = 0
            End If
            
        Else
        
            MsgBox "No. Keahlian yang tidak sah bagi keahlian biasa." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Turutan nombor keahlian bagi SILVER ialah 2000 ~ 2999" & vbCrLf & _
            vbNullString & vbCrLf & _
            "Sila periksa nombor keahlian yang tercatat di belakang kad ini.", vbInformation, "Info"
            
            Frm68.TB18.SetFocus
            
            Exit Sub
        End If
        
    End If
    
    If LM_KATEGORI = "G" Then
    
        If 3000 <= LM_NO_GILIRAN_AHLI And LM_NO_GILIRAN_AHLI <= 3999 Then
            
            Note = "Kategori pelanggan yang cuba didaftarkan adalah GOLD dengan nombor keahlian " & UCase(Frm68.TB18) & "." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Adakah anda yakin untuk teruskan pendaftaran ini?"
            
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbNo Then
                Exit Sub
            End If
            If Answer = vbYes Then
                Frm68.CB9 = 0
                Frm68.CB10 = 0
                Frm68.CB11 = 1
                Frm68.CB12 = 0
            End If
            
        Else
        
            MsgBox "No. Keahlian yang tidak sah bagi keahlian biasa." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Turutan nombor keahlian bagi GOLD ialah 3000 ~ 3999" & vbCrLf & _
            vbNullString & vbCrLf & _
            "Sila periksa nombor keahlian yang tercatat di belakang kad ini.", vbInformation, "Info"
            
            Frm68.TB18.SetFocus
            
            Exit Sub
            
        End If
        
    End If
    
    If LM_KATEGORI = "P" Then
    
        If 4000 <= LM_NO_GILIRAN_AHLI And LM_NO_GILIRAN_AHLI <= 4999 Then
            
            Note = "Kategori pelanggan yang cuba didaftarkan adalah PLATINUM dengan nombor keahlian " & UCase(Frm68.TB18) & "." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Adakah anda yakin untuk teruskan pendaftaran ini?"
            
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbNo Then
                Exit Sub
            End If
            If Answer = vbYes Then
                Frm68.CB9 = 0
                Frm68.CB10 = 0
                Frm68.CB11 = 0
                Frm68.CB12 = 1
            End If
            
        Else
        
            MsgBox "No. Keahlian yang tidak sah bagi keahlian biasa." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Turutan nombor keahlian bagi PLATINUM ialah 4000 ~ 4999" & vbCrLf & _
            vbNullString & vbCrLf & _
            "Sila periksa nombor keahlian yang tercatat di belakang kad ini.", vbInformation, "Info"
            
            Frm68.TB18.SetFocus
            
            Exit Sub
            
        End If
        
    End If
skip_kategori:
'### Periksa no keahlian mengikut kategori - End

    Frm68.CB17 = 1
    Frm68.CB19 = 1
    Frm68.CB20 = 0
    Frm68.TB12 = UCase(Frm68.TB18) 'No. Keahlian
    
Else

    Frm68.CB19 = 0
    Frm68.CB20 = 1
    
End If

If Frm68.CB18 = 1 Then

    Frm68.CB9.Enabled = False
    Frm68.CB10.Enabled = False
    Frm68.CB11.Enabled = False
    Frm68.CB12.Enabled = False
    
    Frm68.L64_Text.Visible = True
    
End If

If Frm68.CB15 = 1 Then
    
    If LM_KATEGORI = vbNullString Then
    
    End If
    
End If

Frm68.CMD1.Visible = True
Frm68.Frame1.Visible = True
Frm68.Frame2.Visible = False
End Sub


Private Sub CMD29_Click()
'On Error Resume Next
Note = "Adakah anda ingin batalkan urusan edit data pelanggan/ahli ini?" & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then
    Frm68.Frame5.Visible = True
    Frm68.Frame1.Visible = False
End If
End Sub

Private Sub CMD3_Click()
'On Error Resume Next
If Frm68.L5_Text = vbNullString Then

    MsgBox "Tiada maklumat bagi No. Agen / Staff.", vbInformation, "Info"
    Exit Sub
    
End If

Frm68.L6_Text = Frm68.DTPicker1
Frm68.L7_Text = Frm68.DTPicker2

Note = "Sistem mungkin akan mengambil sedikit masa untuk memaparkan maklumat senarai komisyen agen ini." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then
    
    'Frm68.L8_Text = 0
    Frm68.L61_Text = 0
    'Frm68.L9_Text = "RM 0.00"
    Frm68.L62_Text = "RM 0.00"
    
    Frm68.L56_Text = -1
    GM_NEXT_PREV = 0
    
    Call Frm68_SenaraiKomisyen_Header
    Call Frm68_SenaraiKomisyen_page
End If
End Sub
Private Sub CMD30_Click()
'on error resume next
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset

Dim TA As Date
Dim TM As Date

Note = "Sistem akan mengira dan memaparkan komisen yang dibuat oleh agen dari " & Frm68.DTPicker7 & " hingga " & Frm68.DTPicker8 & "." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Sistem mungkin akan mengambil masa untuk memaparkan semua rekod ini." & vbCrLf & _
        vbNullString & vbCrLf & _
        "**** SILA PASTIKAN ANDA HANYA MEMBUAT PENGIRAAN dari satu station (PC) sahaja dalam satu masa ****" & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then

    LM_SUSUNAN_RANKING = vbNullString
    
    LM_SUSUNAN_RANKING = InputBox("Sila pilih cara penetapan ranking." & _
            vbCrLf & "Sila masukkan mengikut nombor yang diwakili oleh penetapan tersebut seperti di bawah." & _
            vbCrLf & _
            vbCrLf & vbTab & "1 - Bilangan Jualan" & _
            vbCrLf & vbTab & "2 - Jumlah Berat" & _
            vbCrLf & vbTab & "3 - Jumlah Harga Jualan" & _
            vbCrLf & vbTab & "4 - Jumlah Komisen", "Pilihan susunan ranking")
             
    Select Case LM_SUSUNAN_RANKING
        Case "1"
    
            G_RANKING_FIELD = "bil_barang"
            LM_RANKING = "jumlah BILANGAN JUALAN yang dibuat."
            
        Case "2"
        
            G_RANKING_FIELD = "jumlah_berat"
            LM_RANKING = "jumlah BERAT JUALAN yang dibuat."
            
        Case "3"

            G_RANKING_FIELD = "jumlah_harga"
            LM_RANKING = "jumlah HARGA JUALAN yang dibuat."
            
        Case "4"
    
            G_RANKING_FIELD = "jumlah_komisen"
            LM_RANKING = "JUMLAH KOMISEN yang diperolehi."
            
        Case Else
        
            MsgBox "Tiada pilihan dibuat atau pilihan yang tidak sah.", vbInformation, "Info"
            
            Exit Sub
            
    End Select
    
    
    '###Padam Table jualan agen ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    
    strsql = "TRUNCATE TABLE 75_senarai_komisen_agen"
    
    Set rs = cn.Execute(strsql)
    Set rs = Nothing
    '###Padam Table jualan agen ### - End
    
    TM = Frm68.DTPicker7
    TA = Frm68.DTPicker8
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from senarai_pelanggan where dropship = 1 AND status = 1", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        
        LM_JUM_KOMISEN = 0
        LM_BIL = 0
        LM_JUM_BERAT = 0
        LM_HARGA = 0
        
        Set rs1 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs1.Open "select COUNT(ID) , SUM(berat_jualan) , SUM(harga_jualan_dengan_gst) , SUM(jumlah_komisyen) from 23_senarai_jualan where no_rujukan_agen_dropship='" & rs!no_pelanggan & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND status_rekod = 1", cn, adOpenKeyset, adLockOptimistic
        
        If Not IsNull(rs1(0)) Then LM_BIL = rs1(0)
        If Not IsNull(rs1(1)) Then LM_JUM_BERAT = rs1(1)
        If Not IsNull(rs1(2)) Then LM_HARGA = rs1(2)
        If Not IsNull(rs1(3)) Then LM_JUM_KOMISEN = rs1(3)
        
        rs1.Close
        Set rs1 = Nothing

        Set rs1 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "insert into 75_senarai_komisen_agen(nama,no_tel,no_agen,bil_barang,jumlah_berat,jumlah_harga,jumlah_komisen)" & _
                    "select nama,no_tel,no_pelanggan,'" & LM_BIL & "','" & Format(LM_JUM_BERAT, "0.00") & "','" & Format(LM_HARGA, "0.00") & "','" & Format(LM_JUM_KOMISEN, "0.00") & "' from senarai_pelanggan WHERE no_pelanggan='" & rs!no_pelanggan & "'"
        
        Set rs1 = cn.Execute(strsql)
        Set rs1 = Nothing
        
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
    x = 0
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 75_senarai_komisen_agen order by " & G_RANKING_FIELD & " DESC", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
    
        x = x + 1
        rs!ranking = x
        rs.Update
        
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
    If x <> 0 Then Call frm68_statement_komisen
    
    MsgBox "Pengiraan komisen bagi agen dropship telah selesai." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Sistem telah menyusun mengikut ranking jumlah jualan agen." & vbCrLf & _
            "Ranking dikira dari " & LM_RANKING, vbInformation, "Info"
            '"Ranking dikira dari jumlah berat jualan yang dibuat.", vbInformation, "Info"
    
End If
End Sub

Private Sub CMD31_Click()
'On Error Resume Next
Frm68.Frame8.Visible = False
End Sub

Private Sub CMD4_Click()
'On Error Resume Next
Frm68.Frame5.Visible = True
Frm68.Frame9.Visible = False
End Sub
Private Sub CMD7_Click()
'on error resume next
If Frm68.CBB3 = vbNullString Then

    MsgBox "Sila buat pilihan krateria.", vbInformation, "Info"
    
    Frm68.TB14.SetFocus
    
    Exit Sub
End If

If Frm68.TB14.Visible = True And Frm68.TB14 = vbNullString Then

    MsgBox "Sila masukkan carian.", vbInformation, "Info"
    
    Frm68.TB14.SetFocus
    
    Exit Sub
End If

If InStr(1, Frm68.TB14, "*") <> 0 Or InStr(1, Frm68.TB14, "/") <> 0 Or InStr(1, Frm68.TB14, "\") <> 0 Or InStr(1, Frm68.TB14, "'") <> 0 Then

    MsgBox "Carian mengadungi simbol yang tidak sah.", vbInformation, "Info"
    
    Frm68.TB14.SetFocus
    
    Exit Sub
    
End If

If Frm68.CBB3 <> vbNullString Then Frm68.L39_Text = Frm68.CBB3
If Frm68.TB14 <> vbNullString Then
    Frm68.L40_Text = UCase(Frm68.TB14)
Else
    Frm68.L40_Text = vbNullString
End If
    
GM_NEXT_PREV = 0

Frm68.L69_Text = -1 'Titik Pencarian Data
Frm68.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Frm68.L67_Text = 0 'Paparan Page ke-xxx

Note = "Sistem mungkin akan mengambil sedikit masa untuk memaparkan senarai pelanggan ini." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then

    Call frm68_senarai_pelanggan_header
    Call frm68_senarai_pelanggan
    
End If
End Sub
Private Sub CMD9_Click()
'on error resume next
Frm68.Frame10.Visible = False
Frm68.Frame9.Visible = True
End Sub

Private Sub Command2_Click()

End Sub



Private Sub Form_Load()
'on error resume next
GLOBAL_DISABLE = 0
Frm68.L15_Text = 0

Frm68.L72_Text = 0 'Bilangan ahli yang aktif
Frm68.L73_Text = 0 'Bilangan ahli yang tidak aktif

Frm68.CBB3.Clear

Frm68.CBB3.AddItem "Semua senarai"
Frm68.CBB3.AddItem "Semua pelanggan biasa"
Frm68.CBB3.AddItem "Semua ahli biasa"
Frm68.CBB3.AddItem "Semua silver"
Frm68.CBB3.AddItem "Semua gold"
Frm68.CBB3.AddItem "Semua platinum"
Frm68.CBB3.AddItem "Semua agen dropship"
Frm68.CBB3.AddItem "Nama"
Frm68.CBB3.AddItem "No. kad pengenalan"
Frm68.CBB3.AddItem "No. keahlian"
Frm68.CBB3.AddItem "No. telefon"

Frm68.CBB3 = "Semua senarai"

If Len(G_MODE) = 0 Or Len(G_MIN_LEN) = 0 Or Len(G_MAX_LEN) = 0 Or Len(G_CODE) = 0 Then

    Call sys_config_membership

End If
End Sub

Private Sub Frm58_SM_rekod_belian_Click()
'On Error Resume Next
DATA_FOUND = 0 '0 : Data Tidak Dijumpai , 1 : Data Telah Berjaya Dijumpai
frm68_LM_No_ID = vbNullString

If IsNumeric(Frm68.LV1.SelectedItem.Index) Then
    
    frm68_LM_No_ID = Frm68.LV1.ListItems(Frm68.LV1.SelectedItem.Index)
    
    If frm68_LM_No_ID <> vbNullString Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where ID='" & frm68_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
    
        If Not rs.EOF Then
        
            If Not IsNull(rs!kategori_pelanggan) Then
                If rs!kategori_pelanggan = 1 Then Frm68.L49_Text = "Pelanggan Biasa"
                If rs!kategori_pelanggan = 2 Then Frm68.L49_Text = "Ahli Biasa"
                If rs!kategori_pelanggan = 3 Then Frm68.L49_Text = "Silver"
                If rs!kategori_pelanggan = 4 Then Frm68.L49_Text = "Gold"
                If rs!kategori_pelanggan = 5 Then Frm68.L49_Text = "Platinum"
                'If rs!kategori_pelanggan = 6 Then Frm68.L49_Text = "Master Dealer"
            End If
            If Not IsNull(rs!Nama) Then Frm68.L43_Text = rs!Nama 'Nama
            If Not IsNull(rs!no_ic) Then Frm68.L44_Text = rs!no_ic 'No. IC
            If Not IsNull(rs!no_tel) Then Frm68.L45_Text = rs!no_tel 'No. Telefon
            If Not IsNull(rs!no_pelanggan) Then Frm68.L46_Text = rs!no_pelanggan 'No. Customer
            DATA_FOUND = 1 '0 : Data Tidak Dijumpai , 1 : Data Telah Berjaya Dijumpai
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then '0 : Data Tidak Dijumpai , 1 : Data Telah Berjaya Dijumpai
            Frm68.Frame5.Visible = False
            Frm68.Frame6.Visible = True
        End If
    
    End If
End If
End Sub
Private Sub Frm68_SM_cetak_invoice_Click()
'on error resume next
DATA_FOUND = 1

frm68_LM_No_ID = vbNullString

If IsNumeric(Frm68.LV1.SelectedItem.Index) Then
    
    frm68_LM_No_ID = Frm68.LV1.ListItems(Frm68.LV1.SelectedItem.Index)

    If frm68_LM_No_ID <> vbNullString Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where ID='" & frm68_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!no_pelanggan) Then Frm68.L5_Text = rs!no_pelanggan 'No. Pelanggan
            If Not IsNull(rs!Nama) Then Frm68.L42_Text = rs!Nama 'Tarikh
            
            DATA_FOUND = 1
        End If
    End If
    
    If DATA_FOUND = 1 Then
        Frm68.Frame5.Visible = False
        Frm68.Frame9.Visible = True
    End If
End If
End Sub
Private Sub Frm68_SM_cetak_invoice2_Click()
'on error resume next
DATA_FOUND = 0
Frm68_LM_INVOICE = 0

frm68_LM_No_ID = vbNullString

If IsNumeric(Frm68.LV1.SelectedItem.Index) Then
    
    frm68_LM_No_ID = Frm68.LV1.ListItems(Frm68.LV1.SelectedItem.Index)

    If frm68_LM_No_ID <> vbNullString Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where ID='" & frm68_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!yuran_flag) Then
                If rs!yuran_flag = 1 Then Frm68_LM_INVOICE = 1
            End If
            If Not IsNull(rs!no_invoice) Then
                G_INVOICE_AHLI = rs!no_invoice
                DATA_FOUND = 1
            End If
            If Not IsNull(rs!cawangan) Then G_KEDAI = rs!cawangan
            
        End If
        
        rs.Close
        Set rs = Nothing

        If Frm68_LM_INVOICE = 1 Then
            If DATA_FOUND = 1 Then
                Call Frm68_invoice_yuran_ahli
            Else
                MsgBox "Tiada data bagi invoice bayaran yuran pendaftaran bagi ahli/pelanggan ini.", vbExclamation, "Info"
            End If
        Else
            MsgBox "Tiada data bagi invoice bayaran yuran pendaftaran bagi ahli/pelanggan ini.", vbExclamation, "Info"
        End If
        
    End If
End If
End Sub
Private Sub Frm68_SM_CetakPenyata_Click()
'on error resume next
Dim TM As Date
Dim TA As Date

TM = Frm68.L6_Text
TA = Frm68.L7_Text

'Set rs = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from default_setting where default1='" & "default" & "'", cn, adOpenKeyset, adLockOptimistic

'If Not rs.EOF Then
    
'    If Not IsNull(rs!default_printer) Then LM_PRINTER = rs!default_printer
    
'End If

'rs.Close
'Set rs = Nothing

'For Each oPrn In Printers
'    If oPrn.DeviceName = LM_PRINTER Then
'        Set Printer = oPrn
        PRINTER_FOUND = 1 '0 : Printer Not Found , 1 : Printer Found
'        Exit For
'    End If
'Next

'### Reset maklumat kedai ### - Start
Report53.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report53.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report53.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report53.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report53.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report53.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report53.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report53.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report53.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report53.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

'### Reset Maklumat Penjual #### - Start
Report53.Sections("Section4").Controls("L1").Caption = vbNullString 'Maklumat Pembeli : Nama
Report53.Sections("Section4").Controls("L2").Caption = vbNullString 'Maklumat Pembeli : No. Kad Pengenalan
Report53.Sections("Section4").Controls("L3").Caption = vbNullString 'Maklumat Pembeli : No. Telefon
'### Reset Maklumat Penjual #### - End

'### Maklumat Agen ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm68.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!Nama) Then Report53.Sections("Section4").Controls("L1").Caption = rs!Nama 'Maklumat Agen : Nama
    If Not IsNull(rs!no_tel) Then Report53.Sections("Section4").Controls("L2").Caption = rs!no_tel 'Maklumat Agen : No. Kad Pengenalan
    If Not IsNull(rs!no_pelanggan) Then Report53.Sections("Section4").Controls("L3").Caption = rs!no_pelanggan 'Maklumat Agen : No. Telefon
End If

rs.Close
Set rs = Nothing
'### Maklumat Agen ### - End

Report53.Sections("Section4").Controls("L4").Caption = Frm68.L4_Text
Report53.Sections("Section5").Controls("L5").Caption = Frm68.L61_Text 'Bilangan Barang
Report53.Sections("Section5").Controls("L6").Caption = Frm68.L62_Text 'Jumlah Berat (g)

'###Senarai komisyen bagi Agen / Staff###
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 23_senarai_jualan where status_rekod = 1 AND no_rujukan_agen_dropship='" & Frm68.L5_Text & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report53.DataSource = rs
    Report53.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing
End Sub
Private Sub Frm68_SM_Edit_Simpanan_Click()
'on error resume next
DATA_FOUND = 0 '0 : Data Tidak Dijumpai , 1 : Data Telah Berjaya Dijumpai
DATA_PEKERJA_FOUND = 0

If Frm68.MSFlexGrid4 <> vbNullString Then
    frm68_LM_No_ID = Frm68.MSFlexGrid4.TextMatrix(Frm68.MSFlexGrid4, 2) 'No. ID
    
    If frm68_LM_No_ID <> vbNullString Then
    
        'Call Frm68_Reset_All
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm68.L32_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
        If Not rs.EOF Then
            If Not IsNull(rs!Nama) Then Frm68.L16_Text = rs!Nama 'Nama
            If Not IsNull(rs!no_ic) Then Frm68.L17_Text = rs!no_ic 'No. Kad Pengenalan
            If Not IsNull(rs!no_tel) Then Frm68.L18_Text = rs!no_tel 'No. Telefon
            If Not IsNull(rs!no_pelanggan) Then Frm68.L19_Text = rs!no_pelanggan 'No. Pelanggan
            If Not IsNull(rs!baki_simpanan) Then Frm68.L28_Text = rs!baki_simpanan 'Baki Simpanan Yang Ada (RM)
            DATA_FOUND = 1 '0 : Data Tidak Dijumpai , 1 : Data Telah Berjaya Dijumpai
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 24_rekod_kewangan_pelanggan where ID='" & frm68_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!tarikh) Then
                    Frm68.DTPicker3 = rs!tarikh 'Tarikh
                Else
                    Frm68.DTPicker3 = DateTime.Date 'Tarikh
                End If
                If Not IsNull(rs!no_resit) Then
                    Frm68.L26_Text = rs!no_resit 'No. Rujukan
                Else
                    Frm68.L26_Text = DateTime.Date 'No. Rujukan
                End If
                If Not IsNull(rs!jumlah) Then
                    Frm68.TB17 = rs!jumlah 'Jumlah Simpanan / Penggunaan (RM)
                Else
                    Frm68.TB17 = "0.00" 'Jumlah Simpanan / Penggunaan (RM)
                End If
                If Not IsNull(rs!no_rujukan_pekerja) Then
                    Frm68_LM_No_PEKERJA = rs!no_rujukan_pekerja  'No. Pekerja
                    DATA_PEKERJA_FOUND = 1
                End If
            End If
            
            rs.Close
            Set rs = Nothing
            
            '### Carian Maklumat Penjual (Data Pekerja) ### - Start
            DATA_PEKERJA_FOUND = 0
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from employee where NoPekerja='" & Frm68_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                Frm68_LM_MAKLUMAT_PEKERJA = rs!Samaran & "  |  " & rs!NoPekerja
                DATA_PEKERJA_FOUND = 1
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_PEKERJA_FOUND = 1 Then
                On Error GoTo Err_A:
                Frm68.CBB1 = Frm68_LM_MAKLUMAT_PEKERJA
Restore_A:
            End If
            '### Carian Maklumat Penjual (Data Pekerja) ### - End
        End If
    End If
End If

If DATA_FOUND = 1 Then
    Frm68.CMD14.Visible = False
    Frm68.CMD15.Visible = False
    Frm68.CMD16.Visible = True
    Frm68.CMD17.Visible = True

    Frm68.Pic9.Visible = False
    Frm68.Pic8.Visible = True
End If

Exit Sub
Err_A:
Frm68.CBB1.AddItem Frm68_LM_MAKLUMAT_PEKERJA
Frm68.CBB1 = Frm68_LM_MAKLUMAT_PEKERJA
Resume Restore_A:
End Sub
Private Sub Frm68_SM_EditDataCust_Click()
'On Error Resume Next
DATA_FOUND = 0 '0 : Data Tidak Dijumpai , 1 : Data Telah Berjaya Dijumpai
frm68_LM_No_ID = vbNullString

If IsNumeric(Frm68.LV1.SelectedItem.Index) Then
    
    frm68_LM_No_ID = Frm68.LV1.ListItems(Frm68.LV1.SelectedItem.Index)
    
    If frm68_LM_No_ID <> vbNullString Then
        
        Call Frm68_Reset_All
        
        LM_KATEGORI = 0
        
        Frm68.L66_Text = frm68_LM_No_ID 'No. ID
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where ID='" & frm68_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
    
        If Not rs.EOF Then
        
            GLOBAL_DISABLE = 1
            If Not IsNull(rs!kategori_pelanggan) Then
                If rs!kategori_pelanggan = 1 Then Frm68.L64_Text.Visible = True
                
                If rs!kategori_pelanggan = 2 Then Frm68.CB9 = 1
                If rs!kategori_pelanggan = 3 Then Frm68.CB10 = 1
                If rs!kategori_pelanggan = 4 Then Frm68.CB11 = 1
                If rs!kategori_pelanggan = 5 Then Frm68.CB12 = 1
                
                If rs!kategori_pelanggan = "2" Or rs!kategori_pelanggan = "3" Or rs!kategori_pelanggan = "4" Or rs!kategori_pelanggan = "5" Then
                
                    LM_KATEGORI = 1
                
                End If
                
            End If

            If Not IsNull(rs!Nama) Then Frm68.TB1 = rs!Nama 'Nama
            If Not IsNull(rs!no_ic) Then Frm68.TB2 = rs!no_ic 'No. IC
            If Not IsNull(rs!no_tel) Then Frm68.TB3 = rs!no_tel 'No. Telefon
            If Not IsNull(rs!Email) Then Frm68.TB4 = rs!Email 'E-mail
            If Not IsNull(rs!alamat) Then Frm68.TB5 = rs!alamat 'Alamat
            If Not IsNull(rs!Nama_Waris) Then Frm68.TB6 = rs!Nama_Waris 'Nama Waris
            If Not IsNull(rs!No_Tel_Waris) Then Frm68.TB7 = rs!No_Tel_Waris 'No. Tel Waris
            If Not IsNull(rs!alamat_waris) Then Frm68.TB8 = rs!alamat_waris 'Alamat Waris
            If Not IsNull(rs!nama_bank) Then Frm68.TB9 = rs!nama_bank 'Nama Bank
            If Not IsNull(rs!nama_akaun) Then Frm68.TB10 = rs!nama_akaun 'Nama Akaun
            If Not IsNull(rs!no_akaun) Then Frm68.TB11 = rs!no_akaun 'No. Akaun
            If Not IsNull(rs!no_pelanggan) Then Frm68.TB12 = rs!no_pelanggan 'No. Customer
            If Not IsNull(rs!dropship) Then
                If rs!dropship = 0 Then
                    Frm68.CB14 = 0
                ElseIf rs!dropship = 1 Then
                    Frm68.CB14 = 1
                End If
            Else
                Frm68.CB14 = 0
            End If
            If Not IsNull(rs!membership_card) Then '0 : Tiada kad keahlian , 1 : Ada kad keahlian
                If rs!membership_card = 0 Then
                    Frm68.CB17 = 0
                ElseIf rs!membership_card = 1 Then
                    Frm68.CB17 = 1
                End If
            End If
            If Not IsNull(rs!yuran_flag) Then
                If rs!yuran_flag = 0 Then
                    Frm68.CB20 = 1
                ElseIf rs!yuran_flag = 1 Then
                    Frm68.CB19 = 1
                End If
                If Not IsNull(rs!bil_rasmi) Then
                    
                    If rs!bil_rasmi = 0 Then
                        
                        Frm68.CB13 = 0
                        
                    ElseIf rs!bil_rasmi = 1 Then
                        
                        Frm68.CB13 = 1
                    
                    End If
                
                End If
            Else
                Frm68.CB20 = 1
            End If
            If Not IsNull(rs!jumlah_yuran) Then Frm68.TB19 = Format(rs!jumlah_yuran, "0.00") 'Jumlah bayaran yuran yang dikenakan (RM)
            If Not IsNull(rs!tarikh) Then Frm68.DTPicker4 = rs!tarikh 'Tarikh pendaftaran
            
            DATA_FOUND = 1 '0 : Data Tidak Dijumpai , 1 : Data Telah Berjaya Dijumpai
            GLOBAL_DISABLE = 0
            
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then  '0 : Data Tidak Dijumpai , 1 : Data Telah Berjaya Dijumpai
        
            If LM_KATEGORI = 0 Then
            
                Frm68.CB9.Enabled = False
                Frm68.CB10.Enabled = False
                Frm68.CB11.Enabled = False
                Frm68.CB12.Enabled = False
            Else
                Frm68.CB9.Enabled = True
                Frm68.CB10.Enabled = True
                Frm68.CB11.Enabled = True
                Frm68.CB12.Enabled = True
            End If
            
            Frm68.L11_Text.Visible = True
            Frm68.TB12.Visible = True
            Frm68.CMD2.Visible = True
            Frm68.CMD29.Visible = True
            Frm68.Frame1.Visible = True
            
            Frm68.TB1.SetFocus
            
        Else
        
            MsgBox "Maklumat tentang pelanggan ini tidak dijumpai.", vbInformation, "Info"
            
        End If
        
    End If
End If
End Sub
Private Sub frm68_sm_excel_Click()
'on error resume next
Dim TM As Date
Dim TA As Date

LM_FOUND = 0
frm68_LM_No_ID = vbNullString

frm68_LM_No_ID = vbNullString

If IsNumeric(Frm68.LV1.SelectedItem.Index) Then
    
    frm68_LM_No_ID = Frm68.LV1.ListItems(Frm68.LV1.SelectedItem.Index)
    
    If frm68_LM_No_ID <> vbNullString Then

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
            
                .Cells.VerticalAlignment = xlCenter
                .Columns("A").ColumnWidth = 5 'No.
                .Columns("B").ColumnWidth = 15 'Kategori
                .Columns("C").ColumnWidth = 15 'No. Pelanggan
                .Columns("D").ColumnWidth = 70 'Nama
                .Columns("E").ColumnWidth = 25 'No. Kad Pengenalan
                .Columns("F").ColumnWidth = 25 'No. Telefon
                .Columns("G").ColumnWidth = 35 'E-mail
                .Columns("H").ColumnWidth = 15 'Jumlah Simpanan (RM)
                .Columns("I").ColumnWidth = 10 'Agen Dropship
                .Columns("J").ColumnWidth = 10 'Kad Ahli
                .Columns("K").ColumnWidth = 10 'Mata Terkumpul
                .Columns("L").ColumnWidth = 10 'Status
                    
'No.
'Kategori
'No. Pelanggan
'Nama
'No. Kad Pengenalan
'No. Telefon
'E-mail
'Jumlah Simpanan (RM)
'Agen Dropship
'Kad Ahli
'Mata Terkumpul
'Status

                
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
                
                x = 0
            
                .Cells(1, 5).Font.Bold = True
                .Cells(1, 5).Font.Size = 30
                
                For Row = 1 To 5
                    .Cells(Row, 5).HorizontalAlignment = xlCenter
                Next Row
                
                .Cells(7, 1) = Frm68.L71_Text
                
                .Cells(8, 1) = "No."
                .Cells(8, 2) = "Kategori"
                .Cells(8, 3) = "No. Pelanggan"
                .Cells(8, 4) = "Nama"
                .Cells(8, 5) = "No. Kad Pengenalan"
                .Cells(8, 6) = "No. Telefon"
                .Cells(8, 7) = "E-mail"
                .Cells(8, 8) = "Jumlah Simpanan (RM)"
                .Cells(8, 9) = "Agen Dropship"
                .Cells(8, 10) = "Kad Ahli"
                .Cells(8, 11) = "Mata Terkumpul"
                .Cells(8, 12) = "Status"
                    
                For i = 1 To 12
                    .Cells(8, i).HorizontalAlignment = xlCenter
                    .Cells(8, i).Interior.ColorIndex = 15
                    .Cells(8, i).WrapText = True
                    .Cells(8, i).Borders.LineStyle = xlContinuous
                Next i
        
                If Frm68.L39_Text = "Semua senarai" Then
                    Frm68_LM_SEARCH_1 = Null
                    Frm68_LM_SEARCH_1_LOGIC = "<>"
                    Frm68_LM_FIELD = "kategori_pelanggan"
                    
                    Frm68.L71_Text = "Senarai semua pelanggan."
                End If
                If Frm68.L39_Text = "Semua pelanggan biasa" Then
                    Frm68_LM_SEARCH_1 = "1"
                    Frm68_LM_SEARCH_1_LOGIC = "="
                    Frm68_LM_FIELD = "kategori_pelanggan"
                    
                    Frm68.L71_Text = "Senarai semua pelanggan biasa sahaja."
                End If
                If Frm68.L39_Text = "Semua ahli biasa" Then
                    Frm68_LM_SEARCH_1 = "2"
                    Frm68_LM_SEARCH_1_LOGIC = "="
                    Frm68_LM_FIELD = "kategori_pelanggan"
                    
                    Frm68.L71_Text = "Senarai semua ahli biasa sahaja."
                End If
                If Frm68.L39_Text = "Semua silver" Then
                    Frm68_LM_SEARCH_1 = "3"
                    Frm68_LM_SEARCH_1_LOGIC = "="
                    Frm68_LM_FIELD = "kategori_pelanggan"
                    
                    Frm68.L71_Text = "Senarai semua silver sahaja."
                End If
                If Frm68.L39_Text = "Semua gold" Then
                    Frm68_LM_SEARCH_1 = "4"
                    Frm68_LM_SEARCH_1_LOGIC = "="
                    Frm68_LM_FIELD = "kategori_pelanggan"
                    
                    Frm68.L71_Text = "Senarai semua gold sahaja."
                End If
                If Frm68.L39_Text = "Semua platinum" Then
                    Frm68_LM_SEARCH_1 = "5"
                    Frm68_LM_SEARCH_1_LOGIC = "="
                    Frm68_LM_FIELD = "kategori_pelanggan"
                    
                    Frm68.L71_Text = "Senarai semua platinum sahaja."
                End If
                If Frm68.L39_Text = "Semua agen dropship" Then
                    Frm68_LM_SEARCH_1 = "1"
                    Frm68_LM_SEARCH_1_LOGIC = "="
                    Frm68_LM_FIELD = "dropship"
                    
                    Frm68.L71_Text = "Senarai semua agen dropship sahaja."
                End If
                If Frm68.L39_Text = "Nama" Then
                    Frm68_LM_SEARCH_1 = UCase(Frm68.L40_Text)
                    Frm68_LM_SEARCH_1_LOGIC = "="
                    Frm68_LM_FIELD = "nama"
                    
                    Frm68.L71_Text = "Senarai ahli dengan nama " & UCase(Frm68.L40_Text) & "."
                End If
                If Frm68.L39_Text = "No. kad pengenalan" Then
                    Frm68_LM_SEARCH_1 = UCase(Frm68.L40_Text)
                    Frm68_LM_SEARCH_1_LOGIC = "="
                    Frm68_LM_FIELD = "no_ic"
                    
                    Frm68.L71_Text = "Senarai ahli dengan no kad pengenalan " & UCase(Frm68.L40_Text) & "."
                End If
                If Frm68.L39_Text = "No. keahlian" Then
                    Frm68_LM_SEARCH_1 = UCase(Frm68.L40_Text)
                    Frm68_LM_SEARCH_1_LOGIC = "="
                    Frm68_LM_FIELD = "no_pelanggan"
                    
                    Frm68.L71_Text = "Senarai ahli dengan no keahlian " & UCase(Frm68.L40_Text) & "."
                End If
                If Frm68.L39_Text = "no_tel_hp" Then
                    Frm68_LM_SEARCH_1 = UCase(Frm68.L40_Text)
                    Frm68_LM_SEARCH_1_LOGIC = "="
                    Frm68_LM_FIELD = "no_tel"
                    
                    Frm68.L71_Text = "Senarai ahli dengan no telefon " & UCase(Frm68.L40_Text) & "."
                End If
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from senarai_pelanggan where (status = 0 OR status = 1 OR status = 2) AND " & Frm68_LM_FIELD & " " & Frm68_LM_SEARCH_1_LOGIC & " '" & Frm68_LM_SEARCH_1 & "' order by nama ASC", cn, adOpenKeyset, adLockOptimistic
                
                While rs.EOF = False
                
                    x = x + 1
                    .Cells(8 + x, 1) = x 'No.
                    .Cells(8 + x, 1).HorizontalAlignment = xlCenter
                    
                    If Not IsNull(rs!kategori_pelanggan) Then
                    
                        If rs!kategori_pelanggan = 1 Then .Cells(8 + x, 2) = "Pelanggan Biasa"
                        If rs!kategori_pelanggan = 2 Then .Cells(8 + x, 2) = "Ahli Biasa"
                        If rs!kategori_pelanggan = 3 Then .Cells(8 + x, 2) = "Silver"
                        If rs!kategori_pelanggan = 4 Then .Cells(8 + x, 2) = "Gold"
                        If rs!kategori_pelanggan = 5 Then .Cells(8 + x, 2) = "Platinum"
                        
                    End If
                    If Not IsNull(rs!no_pelanggan) Then .Cells(8 + x, 3) = rs!no_pelanggan 'No. Keahlian
                    If Not IsNull(rs!Nama) Then .Cells(8 + x, 4) = rs!Nama 'Nama Ahli
                    If Not IsNull(rs!no_ic) Then .Cells(8 + x, 5) = "'" & rs!no_ic 'No. Kad Pengenalan
                    If Not IsNull(rs!no_tel) Then .Cells(8 + x, 6) = "'" & rs!no_tel 'No. Tel
                    If Not IsNull(rs!Email) Then .Cells(8 + x, 7) = rs!Email 'E-mail
                    
                    .Cells(8 + x, 8).HorizontalAlignment = xlRight
                    If Not IsNull(rs!baki_simpanan) Then
                        .Cells(8 + x, 8) = Format(rs!baki_simpanan, "#,##0.00") 'Jumlah Simpanan (RM)
                    Else
                        .Cells(8 + x, 8) = Format(0, "#,##0.00") 'Jumlah Simpanan (RM)
                    End If
                    .Cells(8 + x, 8).NumberFormat = "#,##0.00"
                    
                    .Cells(8 + x, 9).HorizontalAlignment = xlCenter
                    If Not IsNull(rs!dropship) Then
                        If rs!dropship = 0 Then
                            .Cells(8 + x, 9) = "Tidak"
                        ElseIf rs!dropship = 1 Then
                            .Cells(8 + x, 9) = "Ya"
                        End If
                    Else
                        .Cells(8 + x, 9) = "Tidak"
                    End If
                    
                    .Cells(8 + x, 10).HorizontalAlignment = xlCenter
                    If Not IsNull(rs!membership_card) Then
                        If rs!membership_card = 0 Then
                            .Cells(8 + x, 10) = "Tidak"
                        ElseIf rs!membership_card = 1 Then
                            .Cells(8 + x, 10) = "Ya"
                        End If
                    Else
                        .Cells(8 + x, 10) = "Tidak"
                    End If
                        
                    .Cells(8 + x, 11).HorizontalAlignment = xlCenter
                    If Not IsNull(rs!baki_point) Then
                        .Cells(8 + x, 11) = rs!baki_point 'Mata Terkumpul
                    Else
                        .Cells(8 + x, 11) = 0 'Mata Terkumpul
                    End If
                    
                    .Cells(8 + x, 12).HorizontalAlignment = xlCenter
                    If Not IsNull(rs!Status) Then
                        If rs!Status = 1 Then
                            .Cells(8 + x, 12) = "Aktif" 'Status
                        ElseIf rs!Status = 0 Then
                            .Cells(8 + x, 12) = "Tidak Aktif" 'Status
                        End If
                    Else
                        .Cells(8 + x, 12) = "Tidak Aktif" 'Status
                    End If
                    
                    For Col = 1 To 12
                        .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
                    Next Col
        
                    rs.MoveNext
                Wend
                
                rs.Close
                Set rs = Nothing

                Y = 0
                Y = x + 1
                
                .Cells(8 + Y, 1) = "Bilangan pelanggan : " & Frm68.L57_Text & "  ,  Bilangan aktif : " & Frm68.L72_Text & "  ,  Bilangan tidak aktif : " & Frm68.L73_Text
                .Cells(8 + Y, 1).Font.Bold = True
                
                Y = Y + 2
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
Private Sub Frm68_SM_komisyen_dropship_Click()
'on error resume next
DATA_FOUND = 1
LM_AGEN = 0
frm68_LM_No_ID = vbNullString

If IsNumeric(Frm68.LV1.SelectedItem.Index) Then
    
    frm68_LM_No_ID = Frm68.LV1.ListItems(Frm68.LV1.SelectedItem.Index)
    
    If frm68_LM_No_ID <> vbNullString Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where ID='" & frm68_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            If Not IsNull(rs!dropship) Then LM_AGEN = rs!dropship
            
            If Not IsNull(rs!no_pelanggan) Then Frm68.L5_Text = rs!no_pelanggan 'No. Pelanggan
            If Not IsNull(rs!Nama) Then Frm68.L42_Text = rs!Nama 'Tarikh
            
            DATA_FOUND = 1
            
        End If
        
    End If
    
    If DATA_FOUND = 1 Then
        
        If LM_AGEN = 1 Then
        
            Frm68.Frame5.Visible = False
            Frm68.Frame9.Visible = True
            
        Else
            
            MsgBox "Pelanggan ini bukan agen dropship kedai.", vbInformation, "Info"
        
        End If
        
    End If
End If
End Sub
Private Sub Frm68_SM_LihatDataCust_Click()
'on error resume next
frm68_LM_No_ID = vbNullString

If IsNumeric(Frm68.LV1.SelectedItem.Index) Then
    
    frm68_LM_No_ID = Frm68.LV1.ListItems(Frm68.LV1.SelectedItem.Index)

    If frm68_LM_No_ID <> vbNullString Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where ID='" & frm68_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!kategori_pelanggan) Then

                If rs!kategori_pelanggan = 1 Then Frm68_LM_KATEGORI = "Pelanggan Biasa"
                If rs!kategori_pelanggan = 2 Then Frm68_LM_KATEGORI = "Ahli Biasa"
                If rs!kategori_pelanggan = 3 Then Frm68_LM_KATEGORI = "Silver"
                If rs!kategori_pelanggan = 4 Then Frm68_LM_KATEGORI = "Gold"
                If rs!kategori_pelanggan = 5 Then Frm68_LM_KATEGORI = "Platinum"
                'If rs!kategori_pelanggan = 6 Then Frm68.L49_Text = "Master Dealer"
            Else
                Frm68_LM_KATEGORI = "Pelanggan Biasa"
            End If
            If Not IsNull(rs!Nama) Then
                Frm68_LM_NAMA = rs!Nama 'Nama
            Else
                Frm68_LM_NAMA = "Tiada Maklumat"
            End If
            If Not IsNull(rs!no_ic) Then
                Frm68_LM_IC = rs!no_ic 'No. Kad Pengenalan
            Else
                Frm68_LM_IC = "Tiada Maklumat"
            End If
            If Not IsNull(rs!no_pelanggan) Then
                Frm68_LM_No_PELANGGAN = rs!no_pelanggan 'No. Pelanggan
            Else
                Frm68_LM_No_PELANGGAN = "Tiada Maklumat"
            End If
            If Not IsNull(rs!no_tel) Then
                Frm68_LM_TEL = rs!no_tel 'No. Telefon
            Else
                Frm68_LM_TEL = "Tiada Maklumat"
            End If
            If Not IsNull(rs!Email) Then
                Frm68_LM_MAIL = rs!Email 'E-mail
            Else
                Frm68_LM_MAIL = "Tiada Maklumat"
            End If
            If Not IsNull(rs!alamat) Then
                Frm68_LM_ADD = rs!alamat 'Alamat
            Else
                Frm68_LM_ADD = "Tiada Maklumat"
            End If
            If Not IsNull(rs!Nama_Waris) Then
                Frm68_LM_WARIS = rs!Nama_Waris 'Nama Waris
            Else
                Frm68_LM_WARIS = "Tiada Maklumat"
            End If
            If Not IsNull(rs!No_Tel_Waris) Then
                Frm68_LM_TEL_WARIS = rs!No_Tel_Waris 'No. Tel Waris
            Else
                Frm68_LM_TEL_WARIS = "Tiada Maklumat"
            End If
            If Not IsNull(rs!alamat_waris) Then
                Frm68_LM_ADD_WARIS = rs!alamat_waris 'Alamat Waris
            Else
                Frm68_LM_ADD_WARIS = "Tiada Maklumat"
            End If
            If Not IsNull(rs!nama_bank) Then
                Frm68_LM_BANK = rs!nama_bank 'Nama Bank
            Else
                Frm68_LM_BANK = "Tiada Maklumat"
            End If
            If Not IsNull(rs!nama_akaun) Then
                Frm68_LM_AKAUN = rs!nama_akaun 'Nama Akaun
            Else
                Frm68_LM_AKAUN = "Tiada Maklumat"
            End If
            If Not IsNull(rs!no_akaun) Then
                Frm68_LM_No_AKAUN = rs!no_akaun 'No. Akaun
            Else
                Frm68_LM_No_AKAUN = "Tiada Maklumat"
            End If
            If Not IsNull(rs!membership_card) Then
                If rs!membership_card = 0 Then '0 : Tiada kad keahlian , 1 : Ada kad keahlian
                    Frm68_LM_KAD_AHLI = "Tiada"
                ElseIf rs!membership_card = 1 Then '0 : Tiada kad keahlian , 1 : Ada kad keahlian
                    Frm68_LM_KAD_AHLI = "Ada"
                End If
            Else
                Frm68_LM_KAD_AHLI = "Tiada"
            End If
        End If
        
        rs.Close
        Set rs = Nothing

        Frm68.L14_Text = vbNullString
        Frm68.L14_Text = "                         Sankyu System                   " & vbCrLf & _
                        "=========================================================" & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Nama : " & Frm68_LM_NAMA & vbCrLf & _
                        "No. Kad Pengenalan : " & Frm68_LM_IC & vbCrLf & _
                        "No. Pelanggan : " & Frm68_LM_No_PELANGGAN & vbCrLf & _
                        "Kad Keahlian : " & Frm68_LM_KAD_AHLI & vbCrLf & _
                        "Kategori Pelanggan : " & Frm68_LM_KATEGORI & vbCrLf & _
                        "No. Telefon : " & Frm68_LM_TEL & vbCrLf & _
                        "E-mail " & Frm68_LM_MAIL & vbCrLf & _
                        "Alamat : " & Frm68_LM_ADD & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "=========================================================" & vbCrLf & _
                        "Maklumat Waris " & vbCrLf & _
                        "=========================================================" & vbCrLf & _
                        "Nama Waris : " & Frm68_LM_WARIS & vbCrLf & _
                        "No. Telefon Waris : " & Frm68_LM_TEL_WARIS & vbCrLf & _
                        "Alamat Waris : " & Frm68_LM_ADD_WARIS & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "=========================================================" & vbCrLf & _
                        "Maklumat Bank" & vbCrLf & _
                        "=========================================================" & vbCrLf & _
                        "Nama Bank : " & Frm68_LM_BANK & vbCrLf & _
                        "Nama Akaun : " & Frm68_LM_AKAUN & vbCrLf & _
                        "No. Akaun : " & Frm68_LM_No_AKAUN
    End If
End If
End Sub

Private Sub Frm68_SM_mata_ganjaran_Click()
'on error resume next
frm68_LM_No_ID = vbNullString
Frm68_LM_KAD = 0
DATA_FOUND = 0

frm68_LM_No_ID = vbNullString

If IsNumeric(Frm68.LV1.SelectedItem.Index) Then
    
    frm68_LM_No_ID = Frm68.LV1.ListItems(Frm68.LV1.SelectedItem.Index)

    If frm68_LM_No_ID <> vbNullString Then

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where ID='" & frm68_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!Nama) Then Frm68_LM_NAMA = rs!Nama 'Nama
            If Not IsNull(rs!no_pelanggan) Then Frm68_LM_No_PELANGGAN = rs!no_pelanggan 'No. keahlian

            If Not IsNull(rs!membership_card) Then
            
                If rs!membership_card = 1 Then
                    Frm68_LM_KAD = 1
                End If
                
            End If
            
            DATA_FOUND = 1
            
        End If
        
        rs.Close
        Set rs = Nothing
            
        If DATA_FOUND = 1 Then
            
            If Frm68_LM_KAD = 1 Then
            
                Call Frm113_initial_setting
                
                'GM_NEXT_PREV = 0
                'Frm113.L7_Text = -1 'Titik Pencarian Data
                'Frm113.L8_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
                'Frm113.L5_Text = 0 'Paparan Page ke-xxx

                Frm113.L4_Text = Frm68_LM_NAMA
                Frm113.L9_Text = Frm68_LM_No_PELANGGAN
                
                Frm68.Hide
                Frm113.Show
                
            Else
            
                MsgBox "Pelanggan ini tidak mempunyai kad keahlian kedai. Oleh itu tiada maklumat mata ganjaran bagi pelanggan ini.", vbExclamation, "Info"
            
            End If
            
        Else
        
            MsgBox "Data tidak dijumpai.", vbInformation, "Info"
            
        End If
        
    End If
    
End If
End Sub

Private Sub Frm68_SM_padam_data_Click()
'on error resume next
frm68_LM_No_ID = vbNullString
DATA_WRITE = 0 '0 : Data Berjaya Diubah , 1 : Tiada Perubahan Pada Data

If IsNumeric(Frm68.LV1.SelectedItem.Index) Then
    
    frm68_LM_No_ID = Frm68.LV1.ListItems(Frm68.LV1.SelectedItem.Index)
    
    If frm68_LM_No_ID <> vbNullString Then
        
        Note = "Adakah anda ingin tukar status pelanggan ini kepada TIDAK AKTIF?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "** Jika pelanggan ini adalah pemegang kad keahlian kedai , kad keahlian tersebut tidak akan dapat digunakan lagi atau digunakan bagi ahli lain. **" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
        
'### Hantar log file ke recovery database ### - Start
            'Set rs = New ADODB.Recordset
            'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            
            'strsql = "insert into spke5100_recovery.senarai_pelanggan(ID_asal,kategori_pelanggan,nama,no_ic,no_tel,email,alamat,nama_waris,no_tel_waris,alamat_waris,nama_bank,nama_akaun,no_akaun,write_timestamp,write_timestamp2,no_pelanggan,baki_simpanan,dropship,membership_card,yuran_flag,jumlah_yuran,tarikh,no_invoice,recovery_datetime,recovery_code)" & _
                        "select ID,kategori_pelanggan,nama,no_ic,no_tel,email,alamat,nama_waris,no_tel_waris,alamat_waris,nama_bank,nama_akaun,no_akaun,write_timestamp,write_timestamp2,no_pelanggan,baki_simpanan,dropship,membership_card,yuran_flag,jumlah_yuran,tarikh,no_invoice,Now(),2 from spke5100.senarai_pelanggan WHERE ID='" & Frm68_LM_No_ID & "'"
            
            'Set rs = cn.Execute(strsql)
            'Set rs = Nothing
'### Hantar log file ke recovery database ### - End
            LM_NOW = Now
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from senarai_pelanggan where ID='" & frm68_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
            If Not rs.EOF Then
                
                If Not IsNull(rs!Status) Then
                    
                    If rs!Status = 0 Then
                        
                        MsgBox "Data bagi pelanggan ini tidak boleh ditukarkan kepada tidak aktif kerana status terkini adalah TIDAK AKTIF.", vbExclamation, "Info"
                        
                        rs.Close
                        Set rs = Nothing
                        
                        Exit Sub
                        
                    End If
                
                End If
                
                If Not IsNull(rs!no_pelanggan) Then Frm68_LM_No_AHLI = rs!no_pelanggan
                If Not IsNull(rs!no_ic) Then Frm68_LM_IC = rs!no_ic


                rs!Status = 0 '0 : Tidak aktif , 1 : Aktif
                rs.Update
                
                DATA_WRITE = 1 '0 : Data Berjaya Diubah , 1 : Tiada Perubahan Pada Data
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_WRITE = 1 Then '0 : Data Berjaya Diubah , 1 : Tiada Perubahan Pada Data
            
                user = MDI_frm1.L3_Text
                LogAct_Memory = "[" & user & "] Tukar status pelanggan kepada tidak aktif.IC [" & Frm68_LM_IC & "] , No. Pelanggan [" & Frm68_LM_No_AHLI & "]"
                LogDate_Memory = LM_NOW
                Call UpdateLog_Database
                
                GM_NEXT_PREV = 2
                
                Call frm68_senarai_pelanggan_header
                Call frm68_senarai_pelanggan
                
                MsgBox "Status pelanggan telah berjaya ditukar.", vbInformation, "Info"
                
            End If
            
        End If
        
    End If
    
End If
End Sub
Private Sub Frm68_SM_PadamDataIni_Click()
'on error resume next
Dim Frm68_LM_JUMLAH_PADAM As Double
Dim Frm68_LM_SIMPANAN_ASAL As Double

DATA_FOUND = 0 '0 : Data Tidak Dijumpai , 1 : Data Telah Berjaya Dijumpai
DATA_PEKERJA_FOUND = 0
Frm68_LM_SIMPANAN_ASAL = 0
Frm68_LM_JUMLAH_PADAM = 0

If Frm68.MSFlexGrid4 <> vbNullString Then
    frm68_LM_No_ID = Frm68.MSFlexGrid4.TextMatrix(Frm68.MSFlexGrid4, 2) 'No. ID
    
    If frm68_LM_No_ID <> vbNullString Then
    
        Note = "Adakah Anda Ingin Padam Data Ini ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
            '### Carian Maklumat Penjual (Data Pekerja) ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 24_rekod_kewangan_pelanggan where ID='" & frm68_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!jumlah) Then Frm68_LM_JUMLAH_PADAM = rs!jumlah
                If Not IsNull(rs!no_resit) Then Frm68_LM_No_PELANGGAN = rs!no_resit 'No. Resit
                If Not IsNull(rs!no_rujukan_pelanggan) Then Frm68_LM_No_PELANGGAN = rs!no_rujukan_pelanggan
                DATA_FOUND = 1
                
                rs.Delete
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
            '### Carian Maklumat Penjual (Data Pekerja) ### - End
        End If
    End If
End If

If DATA_FOUND = 1 Then
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm68_LM_No_PELANGGAN & "'", cn, adOpenKeyset, adLockOptimistic

    If Not rs.EOF Then
        If Not IsNull(rs!baki_simpanan) Then Frm68_LM_SIMPANAN_ASAL = rs!baki_simpanan
        rs!baki_simpanan = Format(Frm68_LM_SIMPANAN_ASAL - Frm68_LM_JUMLAH_PADAM, "0.00")  'Baki Simpanan Terbaru (RM)
        Frm68.L35_Text = Format(Frm68_LM_SIMPANAN_ASAL - Frm68_LM_JUMLAH_PADAM, "0.00") 'Baki Simpanan Terbaru (RM)
        rs.Update
    End If
    
    rs.Close
    Set rs = Nothing
    
    user = MDI_frm1.L3_Text
    LogAct_Memory = "[" & user & "] Padam Data Simpanan Duit Di Kedai , No Rujukan [" & Frm68_LM_No_PELANGGAN & "]."
    LogDate_Memory = DateTime.Date & " " & DateTime.Time$
    Call UpdateLog_Database
    
    Call Frm68_ListSimpanan_Header
    Call Frm68_List_Simpanan
    Call Frm68_ListPenggunaan_Header
    Call Frm68_List_Penggunaan
End If
End Sub
Private Sub Frm68_SM_Rekod_Simpanan_Click()
'on error resume next
DATA_FOUND = 0 '0 : Data Tidak Dijumpai , 1 : Data Telah Berjaya Dijumpai

frm68_LM_No_ID = vbNullString

If IsNumeric(Frm68.LV1.SelectedItem.Index) Then
    
    frm68_LM_No_ID = Frm68.LV1.ListItems(Frm68.LV1.SelectedItem.Index)

    If frm68_LM_No_ID <> vbNullString Then
    
        Call Frm68_Reset_All
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where ID='" & frm68_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
    
        If Not rs.EOF Then
            If Not IsNull(rs!Nama) Then Frm68.L29_Text = rs!Nama 'Nama
            If Not IsNull(rs!no_ic) Then Frm68.L30_Text = rs!no_ic 'No. Kad Pengenalan
            If Not IsNull(rs!no_tel) Then Frm68.L31_Text = rs!no_tel 'No. Telefon
            If Not IsNull(rs!no_pelanggan) Then Frm68.L32_Text = rs!no_pelanggan 'No. Pelanggan
            If Not IsNull(rs!baki_simpanan) Then
                Frm68.L35_Text = Format(rs!baki_simpanan, "0.00") 'Baki Simpanan Asal (RM)
            Else
                Frm68.L35_Text = "0.00"
            End If
            DATA_FOUND = 1 '0 : Data Tidak Dijumpai , 1 : Data Telah Berjaya Dijumpai
        End If
        
        rs.Close
        Set rs = Nothing
    End If
End If

If DATA_FOUND = 1 And Frm68.L32_Text <> vbNullString Then
    Call Frm68_ListSimpanan_Header
    Call Frm68_List_Simpanan
    Call Frm68_ListPenggunaan_Header
    Call Frm68_List_Penggunaan
    
    Frm68.Pic9.Visible = True
End If
End Sub
Private Sub Frm68_SM_Select_Click()
'On Error Resume Next
DATA_FOUND = 0 '0 : Data Tidak Dijumpai , 1 : Data Telah Berjaya Dijumpai

frm68_LM_No_ID = vbNullString

If IsNumeric(Frm68.LV1.SelectedItem.Index) Then
    
    frm68_LM_No_ID = Frm68.LV1.ListItems(Frm68.LV1.SelectedItem.Index)

    If frm68_LM_No_ID <> vbNullString Then
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where ID='" & frm68_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
    
        If Not rs.EOF Then
            
            If rs!Status = 2 Then
            
                MsgBox "Anda tidak dibenarkan untuk memilih data pelanggan ini kerana status adalah TIDAK AKTIF.", vbInformation, "Info"
                
                rs.Close
                Set rs = Nothing
                
                Exit Sub
                
            End If
            
            Frm28_initial

            Note = "Pilih data pelanggan ini ?"
            
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
            If Answer = vbYes Then
                
                If Not IsNull(rs!Nama) Then Frm28.L1_Text = rs!Nama  'Nama
                If Not IsNull(rs!no_ic) Then Frm28.L2_Text = rs!no_ic 'No. Kad Pengenalan
                If Not IsNull(rs!no_tel) Then Frm28.L3_Text = rs!no_tel 'No. Telefon
                If Not IsNull(rs!Email) Then Frm28.L4_Text = rs!Email 'E-mail
                If Not IsNull(rs!no_pelanggan) Then Frm28.L5_Text = rs!no_pelanggan 'No. Pelanggan
                If Not IsNull(rs!baki_simpanan) Then
                    If MDI_frm1.L5_Text = "4" Or MDI_frm1.L5_Text = "5" Then
                        frm130.L26_Text = Format(rs!baki_simpanan, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
                    End If
                    If MDI_frm1.L5_Text = "7" Then Frm87.L27_Text = Format(rs!baki_simpanan, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
                    If MDI_frm1.L5_Text = "10" Then frm130.L26_Text = Format(rs!baki_simpanan, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
                    If MDI_frm1.L5_Text = "8" Then frm130.L26_Text = Format(rs!baki_simpanan, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
                    If MDI_frm1.L5_Text = "9" Then frm130.L26_Text = Format(rs!baki_simpanan, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
                Else
                    If MDI_frm1.L5_Text = "4" Or MDI_frm1.L5_Text = "5" Then
                        frm130.L26_Text = Format(0, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
                    End If
                    If MDI_frm1.L5_Text = "7" Then Frm87.L27_Text = Format(0, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
                    If MDI_frm1.L5_Text = "10" Then frm130.L26_Text = Format(0, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
                    If MDI_frm1.L5_Text = "8" Then frm130.L26_Text = Format(0, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
                    If MDI_frm1.L5_Text = "9" Then frm130.L26_Text = Format(0, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
                End If
                If Not IsNull(rs!membership_card) Then
                    If rs!membership_card = 0 Then
                        If MDI_frm1.L5_Text = "4" Or MDI_frm1.L5_Text = "5" Then
                            Frm84.L79_Text = 0 '0 : Tiada kad , 1 : Ada kad

                            Frm84.L77_Text = "0"

                        End If
                    ElseIf rs!membership_card = 1 Then
                        If MDI_frm1.L5_Text = "4" Or MDI_frm1.L5_Text = "5" Then
                            Frm84.L79_Text = 1 '0 : Tiada kad , 1 : Ada kad
                            If Not IsNull(rs!baki_point) Then
                                Frm84.L77_Text = rs!baki_point
                            Else
                                Frm84.L77_Text = "0"
                            End If
                        End If
                    End If
                Else
                    If MDI_frm1.L5_Text = "4" Or MDI_frm1.L5_Text = "5" Then
                        Frm84.L79_Text = 0 '0 : Tiada kad , 1 : Ada kad

                        Frm84.L77_Text = "0"

                    End If
                End If
                
                DATA_FOUND = 1 '0 : Data Tidak Dijumpai , 1 : Data Telah Berjaya Dijumpai
            End If
            
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then  '0 : Data Tidak Dijumpai , 1 : Data Telah Berjaya Dijumpai
            
            If MDI_frm1.L5_Text = "3" Then
            
                Frm83.Show
                Unload Frm68
                
            ElseIf MDI_frm1.L5_Text = "4" Or MDI_frm1.L5_Text = "5" Then
            
                Frm84.Show
                Unload Frm68
                
            ElseIf MDI_frm1.L5_Text = "6" Then
            
                Frm102.Show
                Unload Frm68
                
            ElseIf MDI_frm1.L5_Text = "7" Then
            
                Frm87.Show
                Unload Frm68
                
            ElseIf MDI_frm1.L5_Text = "10" Then
            
                Frm92.Show
                Unload Frm68
                
            ElseIf MDI_frm1.L5_Text = "8" Then
            
                Frm93.Show
                Unload Frm68
                
            End If
            
        End If
    End If
End If
End Sub
Private Sub Frm68_SM_Select_dropship_Click()
'On Error Resume Next
DATA_FOUND = 0 '0 : Data Tidak Dijumpai , 1 : Data Telah Berjaya Dijumpai

frm68_LM_No_ID = vbNullString

If IsNumeric(Frm68.LV1.SelectedItem.Index) Then
    
    frm68_LM_No_ID = Frm68.LV1.ListItems(Frm68.LV1.SelectedItem.Index)

    If frm68_LM_No_ID <> vbNullString Then
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where ID='" & frm68_LM_No_ID & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
    
        If Not rs.EOF Then
            If Not IsNull(rs!dropship) Then
                If rs!dropship = 1 Then
                    Note = "Pilih Data Agen Dropship Ini ?"
                    
                    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
                    
                    If Answer = vbNo Then
                        'Exit Sub
                    End If
                    If Answer = vbYes Then
                        Call Frm27_initial
                
                        If Not IsNull(rs!Nama) Then
                            Frm27.L1_Text = rs!Nama  'Nama
                            If Frm68.L15_Text = 20 Then Frm84.L29_Text = rs!Nama 'Nama Agen
                        End If
                
                        If Not IsNull(rs!no_ic) Then Frm27.L2_Text = rs!no_ic 'No. Kad Pengenalan
                        If Not IsNull(rs!no_tel) Then Frm27.L3_Text = rs!no_tel 'No. Telefon
                        If Not IsNull(rs!Email) Then Frm27.L4_Text = rs!Email 'E-mail
                        If Not IsNull(rs!no_pelanggan) Then Frm27.L5_Text = rs!no_pelanggan 'No. Pelanggan
                
                        DATA_FOUND = 1 '0 : Data Tidak Dijumpai , 1 : Data Telah Berjaya Dijumpai
                    End If
                Else
                    MsgBox "Pelanggan ini bukan agen dropship bagi kedai.", vbExclamation, "Info"
                End If
            Else
                MsgBox "Pelanggan ini bukan agen dropship bagi kedai.", vbExclamation, "Info"
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then  '0 : Data Tidak Dijumpai , 1 : Data Telah Berjaya Dijumpai
            If MDI_frm1.L5_Text = "5" Then
            
                Frm84.Show
                Unload Frm68
                
            End If
        End If
    End If
End If
End Sub
Private Sub Frm68_SM_SimpanDuit_Click()
'on error resume next
frm68_LM_No_ID = vbNullString
DATA_FOUND = 0 '0 : Data Tidak Dijumpai , 1 : Data Telah Berjaya Dijumpai
LM_NO_PELANGGAN = vbNullString

frm68_LM_No_ID = vbNullString

If IsNumeric(Frm68.LV1.SelectedItem.Index) Then
    
    frm68_LM_No_ID = Frm68.LV1.ListItems(Frm68.LV1.SelectedItem.Index)

    If frm68_LM_No_ID <> vbNullString Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where ID='" & frm68_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
    
        If Not rs.EOF Then
        
            If Not IsNull(rs!Nama) Then LM_NAMA = rs!Nama 'Nama
            If Not IsNull(rs!no_ic) Then LM_IC = rs!no_ic 'No. Kad Pengenalan
            If Not IsNull(rs!no_tel) Then LM_NO_TEL = rs!no_tel 'No. Telefon
            If Not IsNull(rs!no_pelanggan) Then LM_NO_PELANGGAN = rs!no_pelanggan 'No. Pelanggan

            DATA_FOUND = 1 '0 : Data Tidak Dijumpai , 1 : Data Telah Berjaya Dijumpai
            
        End If
        
        rs.Close
        Set rs = Nothing
        
    Else
        
        MsgBox "Tiada data.", vbExclamation, "Info"
        
    End If
    
Else

    MsgBox "Tiada data.", vbExclamation, "Info"
    
End If
    
If DATA_FOUND = 1 Then
    
    If LM_NO_PELANGGAN <> vbNullString Then
        
        Call frm128_reset_data_utama
        Call frm128_pic_ena_disable
        Call frm128_default_setting
        Call frm128_jurujual

        frm128.L1_Text = LM_NAMA
        frm128.L2_Text = LM_IC
        frm128.L3_Text = LM_NO_TEL
        frm128.L4_Text = LM_NO_PELANGGAN
    
    End If
    
End If
End Sub
Private Sub Frm68_SM_tukar_no_Click()
'on error resume next
Dim Frm68_LM_No_AHLI As String

Frm68_LM_NO_ASAL = vbNullString
frm68_LM_No_ID = vbNullString
Frm68_LM_LEN = 0

frm68_LM_No_ID = vbNullString

If IsNumeric(Frm68.LV1.SelectedItem.Index) Then
    
    frm68_LM_No_ID = Frm68.LV1.ListItems(Frm68.LV1.SelectedItem.Index)

    If frm68_LM_No_ID <> vbNullString Then
        
'### Periksa status keahlian samada berdaftar atau tidak ### - Start
        Frm68.L66_Text = frm68_LM_No_ID 'No. ID
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where ID='" & frm68_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!membership_card) Then
            
                If rs!membership_card = 0 Then
                
                    MsgBox "Pelanggan/ahli ini TIDAK berdaftar dengan kedai dan tidak dibenarkan untuk meneruskan menu ini.", vbExclamation, "Info"
                    
                    rs.Close
                    Set rs = Nothing
                    
                    Exit Sub
                    
                End If
                
            Else
            
                MsgBox "Tiada maklumat terperinci status keahlian pelanggan ini. Sila keluar dari menu ini dan cuba lagi.", vbExclamation, "Info"
                
                rs.Close
                Set rs = Nothing
                
                Exit Sub
                
            End If
            
        Else
        
            MsgBox "Tiada maklumat terperinci pelanggan ini. Sila keluar dari menu ini dan cuba lagi.", vbExclamation, "Info"
            
            rs.Close
            Set rs = Nothing
            
            Exit Sub
            
        End If
        
        rs.Close
        Set rs = Nothing
'### Periksa status keahlian samada berdaftar atau tidak ### - End

        Note = "Sila masukkan no. keahlian yang baru." & vbCrLf & _
                "Semua rekod belian yang dibuat oleh ahli ini akan ditukar dengan no. keahlian yang baru." & vbCrLf & _
                vbNullString & vbCrLf & _
                "*** No. Keahlian yang lama TIDAK BOLEH digunakan lagi selepas ini."
        
        Frm68_LM_No_AHLI = InputBox(Note, "Tukaran Nombor Keahlian", "")
        
        If StrPtr(Frm68_LM_No_AHLI) = 0 Then
            Exit Sub
        End If
        
        If StrPtr(Frm68_LM_No_AHLI) <> 0 Then
        
            If Frm68_LM_No_AHLI = "" Then
                MsgBox "Tiada no. keahlian di masukkan." & vbCrLf & _
                        "Urusan dibatalkan.", vbInformation, "Info"
            Else
            
                Frm68_LM_No_AHLI = UCase(Frm68_LM_No_AHLI)
                
                If Frm68_LM_No_AHLI <> vbNullString Then
                
                    If InStr(1, Frm68_LM_No_AHLI, "*") <> 0 Or InStr(1, Frm68_LM_No_AHLI, ".") <> 0 Or InStr(1, Frm68_LM_No_AHLI, "/") <> 0 Or InStr(1, Frm68_LM_No_AHLI, "\") <> 0 Or InStr(1, Frm68_LM_No_AHLI, "'") <> 0 Then
                        
                        MsgBox "No. Keahlian mengandungi simbol yang tidak sah.", vbExclamation, "Info"
                        
                        Exit Sub
                    End If
                    
                End If
                
                '### Periksa panjang no keahlian
                Frm68_LM_LEN = Len(Frm68_LM_No_AHLI)
                
                If Frm68_LM_LEN < G_MIN_LEN Or Frm68_LM_LEN > G_MAX_LEN Then
                    MsgBox "Sila periksa No. Keahlian. No. Keahlian tidak menepati panjang yang ditetapkan." & vbCrLf & _
                            vbNullString & vbCrLf & _
                            "Panjang bagi nombor keahlian yang ditetapkan adalah : " & vbCrLf & _
                            "Minimum abjad   : " & G_MIN_LEN & vbCrLf & _
                            "Maksimum abjad : " & G_MAX_LEN, vbExclamation, "Info"
                    
                    Exit Sub
                End If
                    
                '### Periksa kod kedai
                If InStr(1, UCase(Frm68_LM_No_AHLI), G_CODE) = 0 Then
                    MsgBox "No. keahlian tidak mengandungi kod kedai." & vbCrLf & _
                            "Kod kedai ialah : " & G_CODE, vbInformation, "Info"
                    Exit Sub
                End If
                
                '### Periksa kedudukan kod kedai
                L_CODE_COORDINATE = 0
                
                L_CODE_COORDINATE = InStr(1, UCase(Frm68_LM_No_AHLI), G_CODE)
                
                If L_CODE_COORDINATE <> 0 Then
                    
                    If L_CODE_COORDINATE > 1 Then
                    
                        MsgBox "No. keahlian tidak mengandungi kod kedai." & vbCrLf & _
                                "Kod kedai ialah : " & G_CODE & vbCrLf & _
                                vbNullString & vbCrLf & _
                                "Contoh : " & G_CODE & "002000", vbInformation, "Info"
                        Exit Sub
                        
                    End If
                End If
                
                '### Periksa no keahlian
                LM_NO_GILIRAN_1 = Right(UCase(Frm68_LM_No_AHLI), 6)
                'LM_NO_GILIRAN = Left(LM_NO_GILIRAN_1, 6)
                
                If Not IsNumeric(LM_NO_GILIRAN_1) Then
                    
                    MsgBox "No. Keahlian yang tidak sah.", vbExclamation, "Info"
                    
                    Exit Sub
                
                End If
                
                ' ### Periksa No. Keahlian telah digunakan atau belum ### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm68_LM_No_AHLI & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                
                    If Not IsNull(rs!Nama) Then Frm68_LM_NAMA = rs!Nama 'Nama
                    If Not IsNull(rs!no_pelanggan) Then Frm68_LM_No_KEAHLIAN = rs!no_pelanggan 'No. Keahlian
                    If Not IsNull(rs!no_ic) Then Frm68_LM_No_IC = rs!no_ic 'No. Kad Pengenalan
                
                    MsgBox "No. keahlian [" & Frm68_LM_No_KEAHLIAN & "] telah didaftarkan/digunakan sebelum ini!" & vbCrLf & _
                            vbNullString & vbCrLf & _
                            "Maklumat bagi nombor keahlian ini adalah seperti di bawah :" & vbCrLf & _
                            "Nama : " & Frm68_LM_NAMA & vbCrLf & _
                            "No. Kad Pengenalan : " & Frm68_LM_No_IC, vbExclamation, "Info"
                        
                    rs.Close
                    Set rs = Nothing
                        
                    Exit Sub
                End If
                
                rs.Close
                Set rs = Nothing
                ' ### Periksa No. Keahlian telah digunakan atau belum ### - Start
                
            End If
        
        End If

'#### Maklumat terperinci ahli #### - Start

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where ID='" & frm68_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
    
        If Not rs.EOF Then

            Frm68_LM_NO_ASAL = rs!no_pelanggan 'No. Customer
            
        End If
        
        rs.Close
        Set rs = Nothing


        Note = "Adakah anda ingin menukarkan nombor keahlian pelanggan ini ?" & vbCrLf & _
                "" & vbCrLf & _
                "Maklumat perubahan adalah seperti berikut :" & vbCrLf & _
                "No. Keahlian Lama : " & Frm68_LM_NO_ASAL & vbCrLf & _
                "No. Keahlian Baru : " & Frm68_LM_No_AHLI & vbCrLf & _
                "[" & Frm68_LM_NO_ASAL & "] -> [" & Frm68_LM_No_AHLI & "]" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan ?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then

            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from senarai_pelanggan where ID='" & frm68_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
            If Not rs.EOF Then
    
                G_ID = rs!ID
                Call recovery_senarai_pelanggan
                
                rs!no_pelanggan = UCase(Frm68_LM_No_AHLI) 'No. Customer
                rs.Update
                
                DATA_FOUND = 1 '0 : Data Tidak Dijumpai , 1 : Data Telah Berjaya Dijumpai
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_FOUND = 1 Then  '0 : Data Tidak Dijumpai , 1 : Data Telah Berjaya Dijumpai
    
                '#### Update maklumat pelanggan dalam #16_gold_bar_belian #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 16_gold_bar_belian set no_rujukan_pelanggan_buyback='" & UCase(Frm68_LM_No_AHLI) & "'" _
                & "WHERE no_rujukan_pelanggan_buyback='" & Frm68_LM_NO_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat pelanggan dalam #16_gold_bar_belian #### - End
                
                '#### Update maklumat pelanggan dalam #22_jualan #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 22_jualan set no_rujukan_pembeli='" & UCase(Frm68_LM_No_AHLI) & "'" _
                & "WHERE no_rujukan_pembeli='" & Frm68_LM_NO_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat pelanggan dalam #22_jualan #### - End
                
                '#### Update maklumat pelanggan dalam #23_senarai_jualan #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 23_senarai_jualan set no_rujukan_pembeli='" & UCase(Frm68_LM_No_AHLI) & "'" _
                & "WHERE no_rujukan_pembeli='" & Frm68_LM_NO_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat pelanggan dalam #23_senarai_jualan #### - End
                
                '#### Update maklumat agen dropship dalam #23_senarai_jualan #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 23_senarai_jualan set no_rujukan_agen_dropship='" & UCase(Frm68_LM_No_AHLI) & "'" _
                & "WHERE no_rujukan_agen_dropship='" & Frm68_LM_NO_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat agen dropship dalam #23_senarai_jualan #### - End
                
                '#### Update maklumat agen dropship dalam #24_rekod_kewangan_pelanggan #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 24_rekod_kewangan_pelanggan set no_rujukan_pelanggan='" & UCase(Frm68_LM_No_AHLI) & "'" _
                & "WHERE no_rujukan_pelanggan='" & Frm68_LM_NO_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat agen dropship dalam #24_rekod_kewangan_pelanggan #### - End
                
                '#### Update maklumat agen dropship dalam #27_senarai_ansuran #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 27_senarai_ansuran set no_rujukan_pelanggan='" & UCase(Frm68_LM_No_AHLI) & "'" _
                & "WHERE no_rujukan_pelanggan='" & Frm68_LM_NO_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat agen dropship dalam #27_senarai_ansuran #### - End
                
                '#### Update maklumat agen dropship dalam #29_akaun_ansuran #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 29_akaun_ansuran set no_rujukan_pembeli='" & UCase(Frm68_LM_No_AHLI) & "'" _
                & "WHERE no_rujukan_pembeli='" & Frm68_LM_NO_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat agen dropship dalam #29_akaun_ansuran #### - End
                
                '#### Update maklumat agen dropship dalam #35_senarai_servis #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 35_senarai_servis set no_pelanggan='" & UCase(Frm68_LM_No_AHLI) & "'" _
                & "WHERE no_pelanggan='" & Frm68_LM_NO_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat agen dropship dalam #35_senarai_servis #### - End
    
                '#### Update maklumat pelanggan dalam #36_akaun_servis #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 36_akaun_servis set no_rujukan_pembeli='" & UCase(Frm68_LM_No_AHLI) & "'" _
                & "WHERE no_rujukan_pembeli='" & Frm68_LM_NO_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat pelanggan dalam #36_akaun_servis #### - End
                
                '#### Update maklumat pelanggan dalam #40_tempahan_deposit #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 40_tempahan_deposit set no_rujukan_pelanggan='" & UCase(Frm68_LM_No_AHLI) & "'" _
                & "WHERE no_rujukan_pelanggan='" & Frm68_LM_NO_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat pelanggan dalam #40_tempahan_deposit #### - End
                
                '#### Update maklumat pelanggan dalam #42_tempahan_siap #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 42_tempahan_siap set no_rujukan_pelanggan='" & UCase(Frm68_LM_No_AHLI) & "'" _
                & "WHERE no_rujukan_pelanggan='" & Frm68_LM_NO_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat pelanggan dalam #42_tempahan_siap #### - End
                
                '#### Update maklumat agen dropship dalam #43_bonus_ahli #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE 43_bonus_ahli set no_ahli='" & UCase(Frm68_LM_No_AHLI) & "'" _
                & "WHERE no_ahli='" & Frm68_LM_NO_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat agen dropship dalam #43_bonus_ahli #### - End
                
                '#### Update maklumat agen dropship dalam #data_database #### - Start
                'Set rs = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                'strsql = "UPDATE data_database set no_rujukan_pelanggan_buyback='" & UCase(Frm68_LM_No_AHLI) & "'" _
                & "WHERE no_rujukan_pelanggan_buyback='" & Frm68_LM_NO_ASAL & "'"
                
                'Set rs = cn.Execute(strsql)
                'Set rs = Nothing
                '#### Update maklumat agen dropship dalam #data_database #### - End
                
                user = MDI_frm1.L3_Text
                LogAct_Memory = "[" & user & "] Ubah no. keahlian pelanggan [" & Frm68_LM_NO_ASAL & "] -> [" & Frm68_LM_No_AHLI & "]."
                LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                Call UpdateLog_Database

                GM_NEXT_PREV = 2
                
                'If Frm68.L39_Text = 1 Then '1 : Carian Ikut Krateria , 2 : Carian Ikut Kategori
                '    Call Frm68_ListCust_Header
                '    Call Frm68_senarai_cust_krateria_page
                'ElseIf Frm68.L39_Text = 2 Then '1 : Carian Ikut Krateria , 2 : Carian Ikut Kategori
                '    Call Frm68_ListCust_Header
                '    Call Frm68_senarai_cust_page
                'End If
                
                Call frm68_senarai_pelanggan_header
                Call frm68_senarai_pelanggan
                
                MsgBox "Data telah berjaya disimpan.", vbInformation, "Info"
    
            End If
    '#### Maklumat terperinci ahli #### - End

        End If
        
    End If
End If
End Sub
Private Sub Frm68_SM_upgrade_pelanggan_Click()
'on error resume next
Dim Frm68_LM_No_AHLI As String

frm68_LM_No_ID = vbNullString
Frm68_LM_LEN = 0

LM_DATA_FOUND = 0

frm68_LM_No_ID = vbNullString

If IsNumeric(Frm68.LV1.SelectedItem.Index) Then
    
    frm68_LM_No_ID = Frm68.LV1.ListItems(Frm68.LV1.SelectedItem.Index)

    If frm68_LM_No_ID <> vbNullString Then
        
'### Periksa status keahlian samada berdaftar atau tidak ### - Start
        Frm68.L66_Text = frm68_LM_No_ID 'No. ID
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where ID='" & frm68_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!membership_card) Then
            
                If rs!membership_card = 1 Then
                
                    MsgBox "Pelanggan/ahli ini sudah berdaftar dengan kedai dan sudah mempunyai kad keahlian kedai.", vbExclamation, "Info"
                    
                    rs.Close
                    Set rs = Nothing
                    
                    Exit Sub
                    
                End If
            
            End If
            
        End If
        
        rs.Close
        Set rs = Nothing
'### Periksa status keahlian samada berdaftar atau tidak ### - End
    
        Frm68_LM_No_AHLI = InputBox("Sila masukkan nombor keahlian yang tercatat pada kad keahlian bagi pelanggan ini.", "Upgrade keahlian", "")
        
        If StrPtr(Frm68_LM_No_AHLI) = 0 Then
            Exit Sub
        End If
        
        If StrPtr(Frm68_LM_No_AHLI) <> 0 Then
        
            If Frm68_LM_No_AHLI = "" Then
                MsgBox "Tiada no. keahlian di masukkan." & vbCrLf & _
                        "Urusan dibatalkan.", vbInformation, "Info"
            Else
            
                Frm68_LM_No_AHLI = UCase(Frm68_LM_No_AHLI)
                
                If Frm68_LM_No_AHLI <> vbNullString Then
                
                    If InStr(1, Frm68_LM_No_AHLI, "*") <> 0 Or InStr(1, Frm68_LM_No_AHLI, ".") <> 0 Or InStr(1, Frm68_LM_No_AHLI, "/") <> 0 Or InStr(1, Frm68_LM_No_AHLI, "\") <> 0 Or InStr(1, Frm68_LM_No_AHLI, "'") <> 0 Then
                        
                        MsgBox "No. Keahlian mengandungi simbol yang tidak sah.", vbExclamation, "Info"
                        
                        Exit Sub
                    End If
                    
                End If
                
                '### Periksa panjang no keahlian
                Frm68_LM_LEN = Len(Frm68_LM_No_AHLI)
                
                If Frm68_LM_LEN < G_MIN_LEN Or Frm68_LM_LEN > G_MAX_LEN Then
                    MsgBox "Sila periksa No. Keahlian. No. Keahlian tidak menepati panjang yang ditetapkan." & vbCrLf & _
                            vbNullString & vbCrLf & _
                            "Panjang bagi nombor keahlian yang ditetapkan adalah : " & vbCrLf & _
                            "Minimum abjad   : " & G_MIN_LEN & vbCrLf & _
                            "Maksimum abjad : " & G_MAX_LEN, vbExclamation, "Info"
                    
                    Exit Sub
                End If
                    
                '### Periksa kod kedai
                If InStr(1, UCase(Frm68_LM_No_AHLI), G_CODE) = 0 Then
                    MsgBox "No. keahlian tidak mengandungi kod kedai." & vbCrLf & _
                            "Kod kedai ialah : " & G_CODE, vbInformation, "Info"
                    Exit Sub
                End If
                    
                '### Periksa kedudukan kod kedai
                L_CODE_COORDINATE = 0
                
                L_CODE_COORDINATE = InStr(1, UCase(Frm68_LM_No_AHLI), G_CODE)
                
                If L_CODE_COORDINATE <> 0 Then
                    
                    If L_CODE_COORDINATE > 1 Then
                    
                        MsgBox "No. keahlian tidak mengandungi kod kedai." & vbCrLf & _
                                "Kod kedai ialah : " & G_CODE & vbCrLf & _
                                vbNullString & vbCrLf & _
                                "Contoh : " & G_CODE & "002000", vbInformation, "Info"
                        Exit Sub
                        
                    End If
                End If
                
                '### Periksa no keahlian
                LM_NO_GILIRAN_1 = Right(UCase(Frm68_LM_No_AHLI), 6)
                'LM_NO_GILIRAN = Left(LM_NO_GILIRAN_1, 6)
                
                If Not IsNumeric(LM_NO_GILIRAN_1) Then
                    
                    MsgBox "No. Keahlian yang tidak sah.", vbExclamation, "Info"
                    
                    Exit Sub
                
                End If
                
                ' ### Periksa No. Keahlian telah digunakan atau belum ### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm68_LM_No_AHLI & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                
                    If Not IsNull(rs!Nama) Then Frm68_LM_NAMA = rs!Nama 'Nama
                    If Not IsNull(rs!no_pelanggan) Then Frm68_LM_No_KEAHLIAN = rs!no_pelanggan 'No. Keahlian
                    If Not IsNull(rs!no_ic) Then Frm68_LM_No_IC = rs!no_ic 'No. Kad Pengenalan
                
                    MsgBox "No. keahlian [" & Frm68_LM_No_KEAHLIAN & "] telah didaftarkan/digunakan sebelum ini!" & vbCrLf & _
                            vbNullString & vbCrLf & _
                            "Maklumat bagi nombor keahlian ini adalah seperti di bawah :" & vbCrLf & _
                            "Nama : " & Frm68_LM_NAMA & vbCrLf & _
                            "No. Kad Pengenalan : " & Frm68_LM_No_IC, vbExclamation, "Info"
                        
                    rs.Close
                    Set rs = Nothing
                        
                    Exit Sub
                End If
                    
                rs.Close
                Set rs = Nothing
                ' ### Periksa No. Keahlian telah digunakan atau belum ### - Start
                
            End If

        End If

'#### Maklumat terperinci ahli #### - Start
        Call Frm68_Reset_All
        
        LM_KATEGORI = 0
        
        Frm68.L66_Text = frm68_LM_No_ID 'No. ID
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where ID='" & frm68_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
    
        If Not rs.EOF Then
            
            If Not IsNull(rs!Status) Then
                
                If rs!Status <> 1 Then
                    
                    MsgBox "Anda tidak dibenarkan untuk upgrade keahlian pelanggan ini kerana status data pelanggan ini adalah TIDAK AKTIF." & vbCrLf & _
                            vbNullString & vbCrLf & _
                            "Sila periksa data terkini pelanggan ini.", vbExclamation, "Info"
                            
                    rs.Close
                    Set rs = Nothing
                    
                    Exit Sub
                
                End If
            
            End If
            
            
            'G_ID = rs!ID
            'Call recovery_senarai_pelanggan
            
            GLOBAL_DISABLE = 1
            
            If Not IsNull(rs!Nama) Then Frm68.TB1 = rs!Nama 'Nama
            If Not IsNull(rs!no_ic) Then Frm68.TB2 = rs!no_ic 'No. IC
            If Not IsNull(rs!no_tel) Then Frm68.TB3 = rs!no_tel 'No. Telefon
            If Not IsNull(rs!Email) Then Frm68.TB4 = rs!Email 'E-mail
            If Not IsNull(rs!alamat) Then Frm68.TB5 = rs!alamat 'Alamat
            If Not IsNull(rs!Nama_Waris) Then Frm68.TB6 = rs!Nama_Waris 'Nama Waris
            If Not IsNull(rs!No_Tel_Waris) Then Frm68.TB7 = rs!No_Tel_Waris 'No. Tel Waris
            If Not IsNull(rs!alamat_waris) Then Frm68.TB8 = rs!alamat_waris 'Alamat Waris
            If Not IsNull(rs!nama_bank) Then Frm68.TB9 = rs!nama_bank 'Nama Bank
            If Not IsNull(rs!nama_akaun) Then Frm68.TB10 = rs!nama_akaun 'Nama Akaun
            If Not IsNull(rs!no_akaun) Then Frm68.TB11 = rs!no_akaun 'No. Akaun
            Frm68.TB12 = Frm68_LM_No_AHLI 'No. Customer
            If Not IsNull(rs!dropship) Then
                If rs!dropship = 0 Then
                    Frm68.CB14 = 0
                ElseIf rs!dropship = 1 Then
                    Frm68.CB14 = 1
                End If
            Else
                Frm68.CB14 = 0
            End If
            If Not IsNull(rs!membership_card) Then '0 : Tiada kad keahlian , 1 : Ada kad keahlian
                If rs!membership_card = 0 Then
                    Frm68.CB17 = 0
                ElseIf rs!membership_card = 1 Then
                    Frm68.CB17 = 1
                End If
            End If
            If Not IsNull(rs!yuran_flag) Then
                If rs!yuran_flag = 0 Then
                    Frm68.CB20 = 1
                ElseIf rs!yuran_flag = 1 Then
                    Frm68.CB19 = 1
                End If
            Else
                Frm68.CB20 = 1
            End If
            If Not IsNull(rs!jumlah_yuran) Then Frm68.TB19 = Format(rs!jumlah_yuran, "0.00") 'Jumlah bayaran yuran yang dikenakan (RM)
            If Not IsNull(rs!tarikh) Then Frm68.DTPicker4 = rs!tarikh 'Tarikh pendaftaran
            
            DATA_FOUND = 1 '0 : Data Tidak Dijumpai , 1 : Data Telah Berjaya Dijumpai
            GLOBAL_DISABLE = 0
            
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then  '0 : Data Tidak Dijumpai , 1 : Data Telah Berjaya Dijumpai
            
            Frm68.L11_Text.Visible = True
            Frm68.TB12.Visible = True
            Frm68.CMD2.Visible = True
            Frm68.CMD29.Visible = True
            Frm68.Frame1.Visible = True
            Frm68.CB17 = 1
            Frm68.CB9 = 1
            Frm68.CB19 = 1
            
            Frm68.CB9.Enabled = True
            Frm68.CB10.Enabled = True
            Frm68.CB11.Enabled = True
            Frm68.CB12.Enabled = True
            
            Frm68.TB2.Locked = True
            Frm68.TB2.BackColor = &H8000000A
            
        End If
'#### Maklumat terperinci ahli #### - End
    
    End If
    
End If
End Sub
Private Sub L12_Text_Click()
'on error resume next
If Frm68.Frame4.Visible = False Then

    Call Frm68_Reset_All
    Frm68.Frame4.Visible = True
    
Else

    Frm68.Frame4.Visible = False
    
End If
End Sub

Private Sub L13_Text_Click()
'on error resume next
If Frm68.Frame8.Visible = False Then

    Call Frm68_Reset_All
    Frm68.Frame8.Visible = True
    
Else

    Frm68.Frame8.Visible = False
    
End If
End Sub

Private Sub L20_Text_Click()
'on error resume next
If Frm68.Frame2.Visible = False Then

    If MDI_frm1.L20_Text = "Semua cawangan" Then
    
        Frm96.CMD2.Visible = True
        Frm96.CMD1.Visible = False
    
        Call Frm96_initial
        
        Frm96.Show vbModal
        
    End If
    
    Call Frm68_Reset_All

    Frm68.Frame2.Visible = True
    If Frm68.Frame3.Visible = True Then Frm68.TB18.SetFocus
    
Else

    Frm68.Frame2.Visible = False
    
End If
End Sub
Private Sub L22_Text_Click()
'on error resume next

'0 : Pendaftaran Biasa
'1 : Jualan Gold Bar
'2 : Buyback Gold Bar
'3 : Jualan BK
'4 : Buyback BK
'5 : Ansuran
'6 : Servis
'7 : Tempahan

If MDI_frm1.L5_Text = "0" Or MDI_frm1.L5_Text = "11" Then
    
    Unload Frm68
    MDI_frm1.L5_Text = 0
    
ElseIf MDI_frm1.L5_Text = "3" Then
    
    Frm83.Show
    Unload Frm68

ElseIf MDI_frm1.L5_Text = "4" Or MDI_frm1.L5_Text = "5" Then
    
    Frm84.Show
    Unload Frm68

ElseIf MDI_frm1.L5_Text = "6" Then
    
    Frm102.Show
    Unload Frm68

ElseIf MDI_frm1.L5_Text = "7" Then
    
    Frm87.Show
    Unload Frm68

ElseIf MDI_frm1.L5_Text = "8" Then
    
    Frm93.Show
    Unload Frm68
    
ElseIf MDI_frm1.L5_Text = "10" Then
    
    Frm92.Show
    Unload Frm68

End If

Exit Sub

If Frm68.L36_Text = 0 Then '0 : Terus dari menu data pelanggan , 1 : Data pembeli , 2 : Data agen dropship

    'Frm15.Show
    
ElseIf Frm68.L36_Text = 1 Then

    If Frm68.L15_Text = 1 Then
        Frm79.Show
        Frm28.Show
    ElseIf Frm68.L15_Text = 2 Then
        Frm78.Show
        Frm28.Show
    ElseIf Frm68.L15_Text = 3 Then
        Frm84.Show
        Frm28.Show
    ElseIf Frm68.L15_Text = 4 Then
        Frm83.Show
        Frm28.Show
    ElseIf Frm68.L15_Text = 5 Then
        Frm87.Show
        Frm28.Show
    ElseIf Frm68.L15_Text = 6 Then
        Frm92.Show
        Frm28.Show
    ElseIf Frm68.L15_Text = 7 Then
        Frm93.Show
        Frm28.Show
    ElseIf Frm68.L15_Text = 8 Then
        Frm102.Show
        Frm28.Show
    End If
    
ElseIf Frm68.L36_Text = 2 Then

'Data Agen Drophip
'--------------------
'20 : Jualan

    If Frm68.L15_Text = 20 Then
        Frm84.Show
        Frm27.Show
    End If

End If

Unload Frm68
End Sub

Private Sub L50_Text_Click()
'On Error Resume Next
Call Frm68_hide_report
Frm68.L56_Text = -1
Frm68.L60_Text = 0 '0 : Rekod belian , 1 : Rekod buyback , 2 : Rekod tempahan , 3 : Rekod ansuran , 4 : Rekod servis
GM_NEXT_PREV = 0

Call Frm68_report_belian_header
Call Frm68_report_belian_page

Frm68.MSFlexGrid6.Visible = True
Frm68.L55_Text = "Report belian dari " & Frm68.L47_Text & " hingga " & Frm68.L48_Text
End Sub
Private Sub L51_Text_Click()
'On Error Resume Next
Call Frm68_hide_report

Frm68.L56_Text = -1
Frm68.L60_Text = 1 '0 : Rekod belian , 1 : Rekod buyback , 2 : Rekod tempahan , 3 : Rekod ansuran , 4 : Rekod servis
GM_NEXT_PREV = 0

Call Frm68_report_buyback_header
Call Frm68_report_buyback_page

Frm68.MSFlexGrid7.Visible = True
Frm68.L55_Text = "Report trade in dari " & Frm68.L47_Text & " hingga " & Frm68.L48_Text
End Sub
Private Sub L52_Text_Click()
'On Error Resume Next
Call Frm68_hide_report

Frm68.L56_Text = -1
Frm68.L60_Text = 2 '0 : Rekod belian , 1 : Rekod buyback , 2 : Rekod tempahan , 3 : Rekod ansuran , 4 : Rekod servis
GM_NEXT_PREV = 0

Call Frm68_report_tempahan_header
Call Frm68_report_tempahan_page

Frm68.MSFlexGrid8.Visible = True
Frm68.L55_Text = "Report tempahan dari " & Frm68.L47_Text & " hingga " & Frm68.L48_Text
End Sub
Private Sub L53_Text_Click()
'On Error Resume Next
Call Frm68_hide_report

Frm68.L56_Text = -1
Frm68.L60_Text = 3 '0 : Rekod belian , 1 : Rekod buyback , 2 : Rekod tempahan , 3 : Rekod ansuran , 4 : Rekod servis
GM_NEXT_PREV = 0

Call Frm68_report_ansuran_header
Call Frm68_report_ansuran_page

Frm68.MSFlexGrid9.Visible = True
Frm68.L55_Text = "Report ansuran dari " & Frm68.L47_Text & " hingga " & Frm68.L48_Text
End Sub
Private Sub L54_Text_Click()
'On Error Resume Next
Call Frm68_hide_report

Frm68.L56_Text = -1
Frm68.L60_Text = 4 '0 : Rekod belian , 1 : Rekod buyback , 2 : Rekod tempahan , 3 : Rekod ansuran , 4 : Rekod servis
GM_NEXT_PREV = 0

Call Frm68_rekod_servis_header
Call Frm68_rekod_servis_page

Frm68.MSFlexGrid10.Visible = True
Frm68.L55_Text = "Report servis dari " & Frm68.L47_Text & " hingga " & Frm68.L48_Text
End Sub

Private Sub LV1_DblClick()
'on error resume next
frm68_LM_No_ID = vbNullString

If IsNumeric(Frm68.LV1.SelectedItem.Index) Then
    
    frm68_LM_No_ID = Frm68.LV1.SelectedItem.Index
    
    If frm68_LM_No_ID <> vbNullString Then

        If MDI_frm1.L5_Text = "5" Then
            Frm68.Frm68_SM_Select_dropship.Visible = True
        Else
            Frm68.Frm68_SM_Select_dropship.Visible = False
        End If
        
        If MDI_frm1.L5_Text = "3" Or MDI_frm1.L5_Text = "4" Or MDI_frm1.L5_Text = "5" Or MDI_frm1.L5_Text = "6" Or MDI_frm1.L5_Text = "7" Or MDI_frm1.L5_Text = "8" Or MDI_frm1.L5_Text = "9" Then
            Frm68.Frm68_SM_Select.Visible = True
        Else
            Frm68.Frm68_SM_Select.Visible = False
        End If
        
        'If Frm68.MSFlexGrid3.TextMatrix(Frm68.MSFlexGrid3, 10) = "Ya" Then
        '    Frm68.Frm68_SM_komisyen_dropship.Enabled = True
        'Else
        '    Frm68.Frm68_SM_komisyen_dropship.Enabled = False
        'End If
        
        user_level = MDI_frm1.L4_Text

        If user_level = "Admin" Or user_level = "HQ" Or user_level = "Developer" Then
        
            Frm68.Frm68_SM_tukar_no.Enabled = True
            Frm68.Frm68_SM_upgrade_pelanggan.Enabled = True
            Frm68.Frm68_SM_padam_data.Enabled = True
            Frm68.Frm68_SM_EditDataCust.Enabled = True
            
        ElseIf user_level = "Manager" Then
            
            Frm68.Frm68_SM_tukar_no.Enabled = True
            Frm68.Frm68_SM_upgrade_pelanggan.Enabled = True
            Frm68.Frm68_SM_padam_data.Enabled = False
            Frm68.Frm68_SM_EditDataCust.Enabled = True
        
        Else
        
            Frm68.Frm68_SM_tukar_no.Enabled = False
            Frm68.Frm68_SM_upgrade_pelanggan.Enabled = False
            Frm68.Frm68_SM_padam_data.Enabled = False
            Frm68.Frm68_SM_EditDataCust.Enabled = False
        
        End If
        
        If G_MODE = "YES" Then

            Frm68.Frm68_SM_mata_ganjaran.Enabled = True
            Frm68.Frm68_SM_tukar_no.Enabled = True
            Frm68.Frm68_SM_upgrade_pelanggan.Enabled = True
            Frm68.Frm68_SM_cetak_invoice2.Enabled = True
            
        Else

            Frm68.Frm68_SM_mata_ganjaran.Enabled = False
            Frm68.Frm68_SM_tukar_no.Enabled = False
            Frm68.Frm68_SM_upgrade_pelanggan.Enabled = False
            Frm68.Frm68_SM_cetak_invoice2.Enabled = False
            
        End If

        PopupMenu Frm68_PM_Menu3
    
    Else
    
        MsgBox "Tiada Data.", vbInformation, "Info"
        
    End If
    
Else

    MsgBox "Tiada Data.", vbInformation, "Info"
    
End If
End Sub
Private Sub MSFlexGrid2_DblClick()
'On Error Resume Next
If Frm68.MSFlexGrid2 <> vbNullString Then

    If IsNumeric(Frm68.MSFlexGrid2) Then
        PopupMenu Frm68_PM_Menu2
    Else
        MsgBox "Tiada Data.", vbExclamation, "Info"
    End If

End If
End Sub
Private Sub MSFlexGrid4_DblClick()
'On Error Resume Next
If Frm68.MSFlexGrid4 <> vbNullString Then

    If IsNumeric(Frm68.MSFlexGrid4) Then
        PopupMenu Frm68_PM_Menu4
    Else
        MsgBox "Tiada Data.", vbExclamation, "Info"
    End If
    
End If

End Sub

