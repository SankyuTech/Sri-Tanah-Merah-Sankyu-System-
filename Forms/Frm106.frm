VERSION 5.00
Begin VB.Form Frm106 
   Caption         =   "Report Kewangan"
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
   Icon            =   "Frm106.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   13035
   ScaleWidth      =   23760
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD1 
      BackColor       =   &H000080FF&
      Caption         =   "Report Excel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      MaskColor       =   &H00400000&
      MouseIcon       =   "Frm106.frx":0ECA
      MousePointer    =   99  'Custom
      TabIndex        =   84
      Top             =   10440
      Width           =   3225
   End
   Begin VB.Label L86_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L86_Text"
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
      Left            =   7920
      TabIndex        =   105
      Top             =   4320
      Width           =   1995
   End
   Begin VB.Label L84_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L84_Text"
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
      Left            =   7920
      TabIndex        =   104
      Top             =   5400
      Width           =   1995
   End
   Begin VB.Label L85_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L85_Text"
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
      Left            =   10320
      TabIndex        =   103
      Top             =   5400
      Width           =   1995
   End
   Begin VB.Label L83_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L83_Text"
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
      Left            =   5760
      TabIndex        =   102
      Top             =   5400
      Width           =   1995
   End
   Begin VB.Label L82_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L82_Text"
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
      Left            =   3600
      TabIndex        =   101
      Top             =   5400
      Width           =   1995
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Pulangan duit pelanggan"
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
      Left            =   120
      TabIndex        =   100
      Top             =   5400
      Width           =   3375
   End
   Begin VB.Label L81_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L81_Text"
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
      Left            =   7920
      TabIndex        =   99
      Top             =   2880
      Width           =   1995
   End
   Begin VB.Label L80_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L80_Text"
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
      Left            =   16200
      TabIndex        =   98
      Top             =   3720
      Width           =   1995
   End
   Begin VB.Label L79_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L79_Text"
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
      Left            =   16920
      TabIndex        =   97
      Top             =   11760
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice (GDN/GRN)"
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
      Left            =   840
      TabIndex        =   96
      Top             =   11760
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label L70_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L70_Text"
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
      Left            =   4320
      TabIndex        =   95
      Top             =   11760
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label L71_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L71_Text"
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
      Left            =   6480
      TabIndex        =   94
      Top             =   11760
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label L72_Text 
      Alignment       =   2  'Center
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
      Height          =   255
      Left            =   8640
      TabIndex        =   93
      Top             =   11760
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label L73_Text 
      Alignment       =   2  'Center
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
      Height          =   255
      Left            =   11040
      TabIndex        =   92
      Top             =   11760
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label L74_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L74_Text"
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
      Left            =   14040
      TabIndex        =   91
      Top             =   11760
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Voucher (GDN/GRN)"
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
      Left            =   840
      TabIndex        =   90
      Top             =   12480
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label L75_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L75_Text"
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
      Left            =   4320
      TabIndex        =   89
      Top             =   12480
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label L76_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L76_Text"
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
      Left            =   6480
      TabIndex        =   88
      Top             =   12480
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label L78_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L78_Text"
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
      Left            =   11040
      TabIndex        =   87
      Top             =   12480
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label L77_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L77_Text"
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
      Left            =   8640
      TabIndex        =   86
      Top             =   12480
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   3960
      X2              =   12240
      Y1              =   4245
      Y2              =   4245
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah (RM)                Tunai (RM)                 Bank In (RM)                     Cek  (RM) "
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
      Left            =   4080
      TabIndex        =   85
      Top             =   3960
      Width           =   8775
   End
   Begin VB.Label L67_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L67_Text"
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
      Left            =   10320
      TabIndex        =   83
      Top             =   6120
      Width           =   1995
   End
   Begin VB.Label L65_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L65_Text"
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
      Left            =   7920
      TabIndex        =   82
      Top             =   5040
      Width           =   1995
   End
   Begin VB.Label L66_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L66_Text"
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
      Left            =   10320
      TabIndex        =   81
      Top             =   5040
      Width           =   1995
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Yuran keahlian"
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
      Left            =   120
      TabIndex        =   80
      Top             =   3240
      Width           =   4095
   End
   Begin VB.Label L63_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L63_Text"
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
      Left            =   3600
      TabIndex        =   79
      Top             =   3240
      Width           =   1995
   End
   Begin VB.Label L64_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L64_Text"
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
      Left            =   5760
      TabIndex        =   78
      Top             =   3240
      Width           =   1995
   End
   Begin VB.Label L61_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L61_Text"
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
      Left            =   5880
      TabIndex        =   77
      Top             =   6720
      Width           =   3375
   End
   Begin VB.Label L62_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L62_Text"
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
      Left            =   120
      TabIndex        =   76
      Top             =   7200
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Bayaran belian emas terpakai dari pelanggan (Tunai) : RM "
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
      Left            =   120
      TabIndex        =   75
      Top             =   6720
      Width           =   6135
   End
   Begin VB.Label L3_Text 
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
      Left            =   5400
      MouseIcon       =   "Frm106.frx":11D4
      MousePointer    =   99  'Custom
      TabIndex        =   74
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   $"Frm106.frx":14DE
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   960
      Left            =   120
      TabIndex        =   73
      Top             =   9120
      Width           =   9090
   End
   Begin VB.Label L60_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L60_Text"
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
      Left            =   2520
      TabIndex        =   72
      Top             =   8760
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label L58_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L58_Text"
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
      Left            =   2520
      TabIndex        =   71
      Top             =   8520
      Width           =   1995
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
      Height          =   255
      Left            =   2520
      TabIndex        =   70
      Top             =   8280
      Width           =   1995
   End
   Begin VB.Label L59_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L59_Text"
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
      Left            =   4800
      TabIndex        =   69
      Top             =   8760
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label L56_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L56_Text"
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
      Left            =   2520
      TabIndex        =   68
      Top             =   8040
      Width           =   1995
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   ": RM   : RM  : RM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1920
      TabIndex        =   67
      Top             =   8040
      Width           =   615
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Tunai                           Bank in                          Kad kredit                       "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   66
      Top             =   8040
      Width           =   2535
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Kesimpulan"
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
      Left            =   120
      TabIndex        =   65
      Top             =   7680
      Width           =   3375
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Debit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   64
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kredit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   63
      Top             =   3840
      Width           =   3735
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   3960
      X2              =   12240
      Y1              =   6045
      Y2              =   6045
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   3960
      X2              =   18120
      Y1              =   3555
      Y2              =   3555
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   3960
      X2              =   18000
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label L55_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L55_Text"
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
      Left            =   7920
      TabIndex        =   62
      Top             =   6120
      Width           =   1995
   End
   Begin VB.Label L54_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L54_Text"
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
      Left            =   5760
      TabIndex        =   61
      Top             =   6120
      Width           =   1995
   End
   Begin VB.Label L53_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L53_Text"
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
      Left            =   3600
      TabIndex        =   60
      Top             =   6120
      Width           =   1995
   End
   Begin VB.Label L52_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L52_Text"
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
      Left            =   7920
      TabIndex        =   59
      Top             =   5760
      Width           =   1995
   End
   Begin VB.Label L51_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L51_Text"
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
      Left            =   5760
      TabIndex        =   58
      Top             =   5760
      Width           =   1995
   End
   Begin VB.Label L50_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L50_Text"
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
      Left            =   3600
      TabIndex        =   57
      Top             =   5760
      Width           =   1995
   End
   Begin VB.Label L49_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L49_Text"
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
      Left            =   5760
      TabIndex        =   56
      Top             =   5040
      Width           =   1995
   End
   Begin VB.Label L48_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L48_Text"
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
      Left            =   3600
      TabIndex        =   55
      Top             =   5040
      Width           =   1995
   End
   Begin VB.Label L47_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L47_Text"
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
      Left            =   5760
      TabIndex        =   54
      Top             =   4680
      Width           =   1995
   End
   Begin VB.Label L46_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L46_Text"
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
      Left            =   3600
      TabIndex        =   53
      Top             =   4680
      Width           =   1995
   End
   Begin VB.Label L45_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L45_Text"
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
      Left            =   20880
      TabIndex        =   52
      Top             =   11640
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label L44_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L44_Text"
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
      Left            =   16320
      TabIndex        =   51
      Top             =   10440
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label L43_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L43_Text"
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
      Left            =   5760
      TabIndex        =   50
      Top             =   4320
      Width           =   1995
   End
   Begin VB.Label L42_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L42_Text"
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
      Left            =   3600
      TabIndex        =   49
      Top             =   4320
      Width           =   1995
   End
   Begin VB.Label L41_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L41_Text"
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
      Left            =   13320
      TabIndex        =   48
      Top             =   3600
      Width           =   1995
   End
   Begin VB.Label L40_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L40_Text"
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
      Left            =   20760
      TabIndex        =   47
      Top             =   9600
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label L39_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L39_Text"
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
      Left            =   10320
      TabIndex        =   46
      Top             =   3600
      Width           =   1995
   End
   Begin VB.Label L38_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L38_Text"
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
      Left            =   7920
      TabIndex        =   45
      Top             =   3600
      Width           =   1995
   End
   Begin VB.Label L37_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L37_Text"
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
      Left            =   5760
      TabIndex        =   44
      Top             =   3600
      Width           =   1995
   End
   Begin VB.Label L36_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L36_Text"
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
      Left            =   3600
      TabIndex        =   43
      Top             =   3600
      Width           =   1995
   End
   Begin VB.Label L35_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L35_Text"
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
      Left            =   5760
      TabIndex        =   42
      Top             =   2880
      Width           =   1995
   End
   Begin VB.Label L34_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L34_Text"
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
      Left            =   3600
      TabIndex        =   41
      Top             =   2880
      Width           =   1995
   End
   Begin VB.Label L33_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L33_Text"
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
      Left            =   5760
      TabIndex        =   40
      Top             =   2520
      Width           =   1995
   End
   Begin VB.Label L32_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L32_Text"
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
      Left            =   3600
      TabIndex        =   39
      Top             =   2520
      Width           =   1995
   End
   Begin VB.Label L31_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L31_Text"
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
      Left            =   13320
      TabIndex        =   38
      Top             =   2160
      Width           =   1995
   End
   Begin VB.Label L30_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L30_Text"
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
      Left            =   20760
      TabIndex        =   37
      Top             =   8160
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label L29_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L29_Text"
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
      Left            =   10320
      TabIndex        =   36
      Top             =   2160
      Width           =   1995
   End
   Begin VB.Label L28_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L28_Text"
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
      Left            =   7920
      TabIndex        =   35
      Top             =   2160
      Width           =   1995
   End
   Begin VB.Label L27_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L27_Text"
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
      Left            =   5760
      TabIndex        =   34
      Top             =   2160
      Width           =   1995
   End
   Begin VB.Label L26_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L26_Text"
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
      Left            =   3600
      TabIndex        =   33
      Top             =   2160
      Width           =   1995
   End
   Begin VB.Label L25_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L25_Text"
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
      Left            =   16800
      TabIndex        =   32
      Top             =   11160
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label L24_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L24_Text"
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
      Left            =   13560
      TabIndex        =   31
      Top             =   11160
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label L23_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L23_Text"
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
      Left            =   11040
      TabIndex        =   30
      Top             =   11160
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label L22_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L22_Text"
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
      Left            =   8640
      TabIndex        =   29
      Top             =   11160
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label L21_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L21_Text"
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
      Left            =   6480
      TabIndex        =   28
      Top             =   11160
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label L20_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L20_Text"
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
      Left            =   4320
      TabIndex        =   27
      Top             =   11160
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label L19_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L19_Text"
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
      Left            =   13320
      TabIndex        =   26
      Top             =   1800
      Width           =   1995
   End
   Begin VB.Label L18_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L18_Text"
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
      Left            =   20760
      TabIndex        =   25
      Top             =   7800
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label L17_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L17_Text"
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
      Left            =   10320
      TabIndex        =   24
      Top             =   1800
      Width           =   1995
   End
   Begin VB.Label L16_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L16_Text"
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
      Left            =   7920
      TabIndex        =   23
      Top             =   1800
      Width           =   1995
   End
   Begin VB.Label L15_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L15_Text"
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
      Left            =   5760
      TabIndex        =   22
      Top             =   1800
      Width           =   1995
   End
   Begin VB.Label L14_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L14_Text"
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
      Left            =   3600
      TabIndex        =   21
      Top             =   1800
      Width           =   1995
   End
   Begin VB.Label L13_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L13_Text"
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
      Left            =   13320
      TabIndex        =   20
      Top             =   1440
      Width           =   1995
   End
   Begin VB.Label L12_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L12_Text"
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
      Left            =   20760
      TabIndex        =   19
      Top             =   7440
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label L11_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L11_Text"
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
      Left            =   10320
      TabIndex        =   18
      Top             =   1440
      Width           =   1995
   End
   Begin VB.Label L10_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L10_Text"
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
      Left            =   7920
      TabIndex        =   17
      Top             =   1440
      Width           =   1995
   End
   Begin VB.Label L9_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L9_Text"
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
      Left            =   5760
      TabIndex        =   16
      Top             =   1440
      Width           =   1995
   End
   Begin VB.Label L8_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L8_Text"
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
      Left            =   3600
      TabIndex        =   15
      Top             =   1440
      Width           =   1995
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   $"Frm106.frx":1596
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
      Left            =   4080
      TabIndex        =   14
      Top             =   960
      Width           =   15375
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Bayaran gaji"
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
      Left            =   120
      TabIndex        =   13
      Top             =   5760
      Width           =   3375
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Perbelanjaan kedai"
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
      Left            =   120
      TabIndex        =   12
      Top             =   5040
      Width           =   3375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Ambilan tunai dari kedai"
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
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   3375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Belian tukaran barang oleh agen"
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
      Left            =   12840
      TabIndex        =   10
      Top             =   10440
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Belian barang trade in"
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
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   3375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Simpanan duit di kedai oleh pelanggan"
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
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   4095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Kemasukkan tunai ke kedai"
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
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai bayaran tempahan"
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
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai bayaran ansuran"
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
      Left            =   840
      TabIndex        =   5
      Top             =   11160
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai servis"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai jualan"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label L4_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tetapan Report"
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
      MouseIcon       =   "Frm106.frx":1640
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label L5_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Debit"
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
      MouseIcon       =   "Frm106.frx":194A
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label L6_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kredit"
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
      Left            =   3600
      MouseIcon       =   "Frm106.frx":1C54
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Frm106"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD1_Click()
'on error resume next
Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook

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
        .Columns("A").ColumnWidth = 40 'Debit / Kredit (Perkara)
        .Columns("B").ColumnWidth = 20 'Jumlah (RM)
        .Columns("C").ColumnWidth = 20 'Tunai (RM)
        .Columns("D").ColumnWidth = 20 'Bank In (RM)
        .Columns("E").ColumnWidth = 20 'Kad Kredit (RM)
        .Columns("F").ColumnWidth = 20 'Simpanan Di Kedai (RM)
        .Columns("G").ColumnWidth = 20 'Cek (RM)
        
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
                .Cells(1, 3) = rs!nama_kedai
                .Cells(1, 3).Font.Name = "Times New Roman"
            End If
            If Not IsNull(rs!no_pendaftaran) Then .Cells(2, 3) = rs!no_pendaftaran
            If Not IsNull(rs!alamat) Then .Cells(3, 3) = rs!alamat
            If Not IsNull(rs!no_tel) Then .Cells(4, 3) = rs!no_tel
            If Not IsNull(rs!no_id_gst) Then .Cells(5, 3) = rs!no_id_gst
        End If
        
        rs.Close
        Set rs = Nothing
        '### Maklumat kedai ### - End
        
        x = 0
    
        .Cells(1, 3).Font.Bold = True
        .Cells(1, 3).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 3).HorizontalAlignment = xlCenter
        Next Row
        
        .Cells(7, 1) = "Rekod kewangan bagi tempoh dari " & Frm105.L5_Text & " hingga " & Frm105.L6_Text & "."
        
        .Cells(8, 1) = "Debit (Perkara)"
        .Cells(8, 2) = "Jumlah (RM)"
        .Cells(8, 3) = "Tunai (RM)"
        .Cells(8, 4) = "Bank In (RM)"
        .Cells(8, 5) = "Kad Kredit (RM)"
        .Cells(8, 6) = "Simpanan Di Kedai (RM)"
        .Cells(8, 7) = "Cek (RM)"
        
        For i = 1 To 7
            .Cells(8, i).HorizontalAlignment = xlRight
            .Cells(8, i).Interior.ColorIndex = 19
            .Cells(8, i).WrapText = True
            '.Cells(8, i).Borders.LineStyle = xlContinuous
            .Cells(8, i).Font.Size = 16
            .Cells(8, i).Font.Bold = True
        Next i
        
        'Senarai Jualan
        .Cells(9, 1) = "Senarai Jualan"
        .Cells(9, 1).HorizontalAlignment = xlRight
        
        .Cells(9, 2) = Format(Frm106.L8_Text, "#,##0.00")
        .Cells(9, 2).HorizontalAlignment = xlRight
        .Cells(9, 2).NumberFormat = "#,##0.00"
        
        .Cells(9, 3) = Format(Frm106.L9_Text, "#,##0.00")
        .Cells(9, 3).HorizontalAlignment = xlRight
        .Cells(9, 3).NumberFormat = "#,##0.00"
        
        .Cells(9, 4) = Format(Frm106.L10_Text, "#,##0.00")
        .Cells(9, 4).HorizontalAlignment = xlRight
        .Cells(9, 4).NumberFormat = "#,##0.00"
        
        .Cells(9, 5) = Format(Frm106.L11_Text, "#,##0.00")
        .Cells(9, 5).HorizontalAlignment = xlRight
        .Cells(9, 5).NumberFormat = "#,##0.00"
        
        .Cells(9, 6) = Format(Frm106.L13_Text, "#,##0.00")
        .Cells(9, 6).HorizontalAlignment = xlRight
        .Cells(9, 6).NumberFormat = "#,##0.00"
        
        'Senarai servis
        .Cells(10, 1) = "Senarai Servis"
        .Cells(10, 1).HorizontalAlignment = xlRight
        
        .Cells(10, 2) = Format(Frm106.L14_Text, "#,##0.00")
        .Cells(10, 2).HorizontalAlignment = xlRight
        .Cells(10, 2).NumberFormat = "#,##0.00"
        
        .Cells(10, 3) = Format(Frm106.L15_Text, "#,##0.00")
        .Cells(10, 3).HorizontalAlignment = xlRight
        .Cells(10, 3).NumberFormat = "#,##0.00"
        
        .Cells(10, 4) = Format(Frm106.L16_Text, "#,##0.00")
        .Cells(10, 4).HorizontalAlignment = xlRight
        .Cells(10, 4).NumberFormat = "#,##0.00"
        
        .Cells(10, 5) = Format(Frm106.L17_Text, "#,##0.00")
        .Cells(10, 5).HorizontalAlignment = xlRight
        .Cells(10, 5).NumberFormat = "#,##0.00"
        
        .Cells(10, 6) = Format(Frm106.L19_Text, "#,##0.00")
        .Cells(10, 6).HorizontalAlignment = xlRight
        .Cells(10, 6).NumberFormat = "#,##0.00"
        
        'Senarai Bayaran Tempahan
        .Cells(11, 1) = "Senarai Bayaran Tempahan"
        .Cells(11, 1).HorizontalAlignment = xlRight
        
        .Cells(11, 2) = Format(Frm106.L26_Text, "#,##0.00")
        .Cells(11, 2).HorizontalAlignment = xlRight
        .Cells(11, 2).NumberFormat = "#,##0.00"
        
        .Cells(11, 3) = Format(Frm106.L27_Text, "#,##0.00")
        .Cells(11, 3).HorizontalAlignment = xlRight
        .Cells(11, 3).NumberFormat = "#,##0.00"
        
        .Cells(11, 4) = Format(Frm106.L28_Text, "#,##0.00")
        .Cells(11, 4).HorizontalAlignment = xlRight
        .Cells(11, 4).NumberFormat = "#,##0.00"
        
        .Cells(11, 5) = Format(Frm106.L29_Text, "#,##0.00")
        .Cells(11, 5).HorizontalAlignment = xlRight
        .Cells(11, 5).NumberFormat = "#,##0.00"
        
        .Cells(11, 6) = Format(Frm106.L31_Text, "#,##0.00")
        .Cells(11, 6).HorizontalAlignment = xlRight
        .Cells(11, 6).NumberFormat = "#,##0.00"
        
        'Invoice (GDN/GRN)
        .Cells(12, 1) = "Invoice (GDN/GRN)"
        .Cells(12, 1).HorizontalAlignment = xlRight
        
        .Cells(12, 2) = Format(Frm106.L70_Text, "#,##0.00")
        .Cells(12, 2).HorizontalAlignment = xlRight
        .Cells(12, 2).NumberFormat = "#,##0.00"
        
        .Cells(12, 3) = Format(Frm106.L71_Text, "#,##0.00")
        .Cells(12, 3).HorizontalAlignment = xlRight
        .Cells(12, 3).NumberFormat = "#,##0.00"
        
        .Cells(12, 4) = Format(Frm106.L72_Text, "#,##0.00")
        .Cells(12, 4).HorizontalAlignment = xlRight
        .Cells(12, 4).NumberFormat = "#,##0.00"
        
        .Cells(12, 5) = Format(Frm106.L73_Text, "#,##0.00")
        .Cells(12, 5).HorizontalAlignment = xlRight
        .Cells(12, 5).NumberFormat = "#,##0.00"
        
        .Cells(12, 6) = Format(Frm106.L74_Text, "#,##0.00")
        .Cells(12, 6).HorizontalAlignment = xlRight
        .Cells(12, 6).NumberFormat = "#,##0.00"
        
        .Cells(12, 7) = Format(Frm106.L60_Text, "#,##0.00")
        .Cells(12, 7).HorizontalAlignment = xlRight
        .Cells(12, 7).NumberFormat = "#,##0.00"
        
        
        'Kemasukkan Tunai Ke Kedai
        .Cells(13, 1) = "Kemasukkan Tunai Ke Kedai"
        .Cells(13, 1).HorizontalAlignment = xlRight
        
        .Cells(13, 2) = Format(Frm106.L32_Text, "#,##0.00")
        .Cells(13, 2).HorizontalAlignment = xlRight
        .Cells(13, 2).NumberFormat = "#,##0.00"
        
        .Cells(13, 3) = Format(Frm106.L33_Text, "#,##0.00")
        .Cells(13, 3).HorizontalAlignment = xlRight
        .Cells(13, 3).NumberFormat = "#,##0.00"
        
        'Simpanan Duit Di Kedai Oleh Pelanggan
        .Cells(14, 1) = "Simpanan Duit Di Kedai Oleh Pelanggan"
        .Cells(14, 1).HorizontalAlignment = xlRight
        
        .Cells(14, 2) = Format(Frm106.L34_Text, "#,##0.00")
        .Cells(14, 2).HorizontalAlignment = xlRight
        .Cells(14, 2).NumberFormat = "#,##0.00"
        
        .Cells(14, 3) = Format(Frm106.L35_Text, "#,##0.00")
        .Cells(14, 3).HorizontalAlignment = xlRight
        .Cells(14, 3).NumberFormat = "#,##0.00"
        
        .Cells(14, 4) = Format(Frm106.L81_Text, "#,##0.00")
        .Cells(14, 4).HorizontalAlignment = xlRight
        .Cells(14, 4).NumberFormat = "#,##0.00"
        
        'Yuran Keahlian
        .Cells(15, 1) = "Yuran Keahlian"
        .Cells(15, 1).HorizontalAlignment = xlRight
        
        .Cells(15, 2) = Format(Frm106.L63_Text, "#,##0.00")
        .Cells(15, 2).HorizontalAlignment = xlRight
        .Cells(15, 2).NumberFormat = "#,##0.00"
        
        .Cells(15, 3) = Format(Frm106.L64_Text, "#,##0.00")
        .Cells(15, 3).HorizontalAlignment = xlRight
        .Cells(15, 3).NumberFormat = "#,##0.00"
        
        'Jumlah Keseluruhan (Debit)
        .Cells(16, 1) = "Jumlah Keseluruhan (Debit)"
        .Cells(16, 1).HorizontalAlignment = xlRight
        .Cells(16, 1).Font.Bold = True
        
        .Cells(16, 2) = Format(Frm106.L36_Text, "#,##0.00")
        .Cells(16, 2).HorizontalAlignment = xlRight
        .Cells(16, 2).Font.Bold = True
        .Cells(16, 2).NumberFormat = "#,##0.00"
        
        .Cells(16, 3) = Format(Frm106.L37_Text, "#,##0.00")
        .Cells(16, 3).HorizontalAlignment = xlRight
        .Cells(16, 3).Font.Bold = True
        .Cells(16, 3).NumberFormat = "#,##0.00"
        
        .Cells(16, 4) = Format(Frm106.L38_Text, "#,##0.00")
        .Cells(16, 4).HorizontalAlignment = xlRight
        .Cells(16, 4).Font.Bold = True
        .Cells(16, 4).NumberFormat = "#,##0.00"
        
        .Cells(16, 5) = Format(Frm106.L39_Text, "#,##0.00")
        .Cells(16, 5).HorizontalAlignment = xlRight
        .Cells(16, 5).Font.Bold = True
        .Cells(16, 5).NumberFormat = "#,##0.00"
        
        .Cells(16, 6) = Format(Frm106.L41_Text, "#,##0.00")
        .Cells(16, 6).HorizontalAlignment = xlRight
        .Cells(16, 6).Font.Bold = True
        .Cells(16, 6).NumberFormat = "#,##0.00"
        
        .Cells(16, 7) = Format(Frm106.L80_Text, "#,##0.00")
        .Cells(16, 7).HorizontalAlignment = xlRight
        .Cells(16, 7).Font.Bold = True
        .Cells(16, 7).NumberFormat = "#,##0.00"
        
        .Cells(18, 1) = "Kredit (Perkara)"
        .Cells(18, 2) = "Jumlah (RM)"
        .Cells(18, 3) = "Tunai (RM)"
        .Cells(18, 4) = "Bank In (RM)"
        .Cells(18, 5) = "Cek (RM)"
        
        For i = 1 To 5
            .Cells(18, i).HorizontalAlignment = xlRight
            .Cells(18, i).Interior.ColorIndex = 19
            .Cells(18, i).WrapText = True
            '.Cells(18, i).Borders.LineStyle = xlContinuous
            .Cells(18, i).Font.Size = 16
            .Cells(18, i).Font.Bold = True
        Next i
        
        'Belian Barang Trade In
        .Cells(19, 1) = "Belian Barang Trade In"
        .Cells(19, 1).HorizontalAlignment = xlRight
        
        .Cells(19, 2) = Format(Frm106.L42_Text, "#,##0.00")
        .Cells(19, 2).HorizontalAlignment = xlRight
        .Cells(19, 2).NumberFormat = "#,##0.00"
        
        .Cells(19, 3) = Format(Frm106.L43_Text, "#,##0.00")
        .Cells(19, 3).HorizontalAlignment = xlRight
        .Cells(19, 3).NumberFormat = "#,##0.00"
        
        .Cells(19, 4) = Format(Frm106.L86_Text, "#,##0.00")
        .Cells(19, 4).HorizontalAlignment = xlRight
        .Cells(19, 4).NumberFormat = "#,##0.00"
        
        .Cells(19, 5) = Format(0, "#,##0.00")
        .Cells(19, 5).HorizontalAlignment = xlRight
        .Cells(19, 5).NumberFormat = "#,##0.00"
        
        'Ambilan Tunai Dari Kedai
        .Cells(20, 1) = "Ambilan Tunai Dari Kedai"
        .Cells(20, 1).HorizontalAlignment = xlRight
        
        .Cells(20, 2) = Format(Frm106.L46_Text, "#,##0.00")
        .Cells(20, 2).HorizontalAlignment = xlRight
        .Cells(20, 2).NumberFormat = "#,##0.00"
        
        .Cells(20, 3) = Format(Frm106.L47_Text, "#,##0.00")
        .Cells(20, 3).HorizontalAlignment = xlRight
        .Cells(20, 3).NumberFormat = "#,##0.00"
        
        .Cells(20, 4) = Format(0, "#,##0.00")
        .Cells(20, 4).HorizontalAlignment = xlRight
        .Cells(20, 4).NumberFormat = "#,##0.00"
        
        .Cells(20, 5) = Format(0, "#,##0.00")
        .Cells(20, 5).HorizontalAlignment = xlRight
        .Cells(20, 5).NumberFormat = "#,##0.00"
        
        'Perbelanjaan Kedai
        .Cells(21, 1) = "Perbelanjaan Kedai"
        .Cells(21, 1).HorizontalAlignment = xlRight
        
        .Cells(21, 2) = Format(Frm106.L48_Text, "#,##0.00")
        .Cells(21, 2).HorizontalAlignment = xlRight
        .Cells(21, 2).NumberFormat = "#,##0.00"
        
        .Cells(21, 3) = Format(Frm106.L49_Text, "#,##0.00")
        .Cells(21, 3).HorizontalAlignment = xlRight
        .Cells(21, 3).NumberFormat = "#,##0.00"
        
        .Cells(21, 4) = Format(Frm106.L65_Text, "#,##0.00")
        .Cells(21, 4).HorizontalAlignment = xlRight
        .Cells(21, 4).NumberFormat = "#,##0.00"
        
        .Cells(21, 5) = Format(Frm106.L66_Text, "#,##0.00")
        .Cells(21, 5).HorizontalAlignment = xlRight
        .Cells(21, 5).NumberFormat = "#,##0.00"
        
        'Voucher (GDN/GRN)
        .Cells(22, 1) = "Voucher (GDN/GRN)"
        .Cells(22, 1).HorizontalAlignment = xlRight
        
        .Cells(22, 2) = Format(Frm106.L75_Text, "#,##0.00")
        .Cells(22, 2).HorizontalAlignment = xlRight
        .Cells(22, 2).NumberFormat = "#,##0.00"
        
        .Cells(22, 3) = Format(Frm106.L76_Text, "#,##0.00")
        .Cells(22, 3).HorizontalAlignment = xlRight
        .Cells(22, 3).NumberFormat = "#,##0.00"
        
        .Cells(22, 4) = Format(Frm106.L77_Text, "#,##0.00")
        .Cells(22, 4).HorizontalAlignment = xlRight
        .Cells(22, 4).NumberFormat = "#,##0.00"
        
        .Cells(22, 5) = Format(Frm106.L78_Text, "#,##0.00")
        .Cells(22, 5).HorizontalAlignment = xlRight
        .Cells(22, 5).NumberFormat = "#,##0.00"
        
        'Pulangan duit pelanggan
        .Cells(23, 1) = "Pulangan duit pelanggan"
        .Cells(23, 1).HorizontalAlignment = xlRight
        
        .Cells(23, 2) = Format(Frm106.L82_Text, "#,##0.00")
        .Cells(23, 2).HorizontalAlignment = xlRight
        .Cells(23, 2).NumberFormat = "#,##0.00"
        
        .Cells(23, 3) = Format(Frm106.L83_Text, "#,##0.00")
        .Cells(23, 3).HorizontalAlignment = xlRight
        .Cells(23, 3).NumberFormat = "#,##0.00"
        
        .Cells(23, 4) = Format(Frm106.L84_Text, "#,##0.00")
        .Cells(23, 4).HorizontalAlignment = xlRight
        .Cells(23, 4).NumberFormat = "#,##0.00"
        
        .Cells(23, 5) = Format(Frm106.L85_Text, "#,##0.00")
        .Cells(23, 5).HorizontalAlignment = xlRight
        .Cells(23, 5).NumberFormat = "#,##0.00"
        
        'Bayaran Gaji
        .Cells(24, 1) = "Bayaran Gaji"
        .Cells(24, 1).HorizontalAlignment = xlRight
        
        .Cells(24, 2) = Format(Frm106.L50_Text, "#,##0.00")
        .Cells(24, 2).HorizontalAlignment = xlRight
        .Cells(24, 2).NumberFormat = "#,##0.00"
        
        .Cells(24, 3) = Format(Frm106.L51_Text, "#,##0.00")
        .Cells(24, 3).HorizontalAlignment = xlRight
        .Cells(24, 3).NumberFormat = "#,##0.00"
        
        .Cells(24, 4) = Format(Frm106.L52_Text, "#,##0.00")
        .Cells(24, 4).HorizontalAlignment = xlRight
        .Cells(24, 4).NumberFormat = "#,##0.00"
        
        .Cells(24, 5) = Format(0, "#,##0.00")
        .Cells(24, 5).HorizontalAlignment = xlRight
        .Cells(24, 5).NumberFormat = "#,##0.00"
        
        
        'Jumlah Keseluruhan (Kredit)
        .Cells(25, 1) = "Jumlah Keseluruhan (Kredit)"
        .Cells(25, 1).HorizontalAlignment = xlRight
        .Cells(25, 1).Font.Bold = True
        
        .Cells(25, 2) = Format(Frm106.L53_Text, "#,##0.00")
        .Cells(25, 2).HorizontalAlignment = xlRight
        .Cells(25, 2).Font.Bold = True
        .Cells(25, 2).NumberFormat = "#,##0.00"
        
        .Cells(25, 3) = Format(Frm106.L54_Text, "#,##0.00")
        .Cells(25, 3).HorizontalAlignment = xlRight
        .Cells(25, 3).Font.Bold = True
        .Cells(25, 3).NumberFormat = "#,##0.00"
        
        .Cells(25, 4) = Format(Frm106.L55_Text, "#,##0.00")
        .Cells(25, 4).HorizontalAlignment = xlRight
        .Cells(25, 4).Font.Bold = True
        .Cells(25, 4).NumberFormat = "#,##0.00"
        
        .Cells(25, 5) = Format(Frm106.L67_Text, "#,##0.00")
        .Cells(25, 5).HorizontalAlignment = xlRight
        .Cells(25, 5).Font.Bold = True
        .Cells(25, 5).NumberFormat = "#,##0.00"
    
        .Cells(27, 1) = "Bayaran belian emas terpakai dari pelanggan (Tunai) : RM " & Format(Frm106.L61_Text, "#,##0.00")
        .Cells(27, 1).HorizontalAlignment = xlLeft
        '.Cells(26, 1).Interior.ColorIndex = 19
        '.Cells(26, 1).Font.Size = 16
        .Cells(27, 1).Font.Bold = True

        
        .Cells(29, 1) = "Kesimpulan"
        .Cells(29, 1).HorizontalAlignment = xlRight
        .Cells(29, 1).Interior.ColorIndex = 19
        .Cells(29, 1).Font.Size = 16
        .Cells(29, 1).Font.Bold = True
        
        .Cells(30, 1) = "Tunai : RM "
        .Cells(30, 1).HorizontalAlignment = xlRight
        .Cells(30, 1).Font.Bold = True
        
        .Cells(31, 1) = "Bank In : RM "
        .Cells(31, 1).HorizontalAlignment = xlRight
        .Cells(31, 1).Font.Bold = True
        
        .Cells(32, 1) = "Kad Kredit : RM "
        .Cells(32, 1).HorizontalAlignment = xlRight
        .Cells(32, 1).Font.Bold = True
        
        '.Cells(33, 1) = "Simpanan Di Kedai : RM "
        '.Cells(33, 1).HorizontalAlignment = xlRight
        '.Cells(33, 1).Font.Bold = True
        
        .Cells(30, 2) = Format(Frm106.L56_Text, "#,##0.00")
        .Cells(30, 2).HorizontalAlignment = xlLeft
        .Cells(30, 2).Font.Bold = True
        .Cells(30, 2).NumberFormat = "#,##0.00"
        
        .Cells(31, 2) = Format(Frm106.L57_Text, "#,##0.00")
        .Cells(31, 2).HorizontalAlignment = xlLeft
        .Cells(31, 2).Font.Bold = True
        .Cells(31, 2).NumberFormat = "#,##0.00"
        
        .Cells(32, 2) = Format(Frm106.L58_Text, "#,##0.00")
        .Cells(32, 2).HorizontalAlignment = xlLeft
        .Cells(32, 2).Font.Bold = True
        .Cells(32, 2).NumberFormat = "#,##0.00"
        
        '.Cells(32, 2) = Format(Frm106.L60_Text, "#,##0.00")
        '.Cells(32, 2).HorizontalAlignment = xlLeft
        '.Cells(32, 2).Font.Bold = True
        '.Cells(32, 2).NumberFormat = "#,##0.00"
        
        .Cells(34, 1).Font.Bold = True
        .Cells(34, 1) = "Report Generated By Sankyu System" 'Watermark Sankyu System

        .Cells(35, 1).Font.Bold = True
        .Cells(35, 1) = "Sankyu System , +6010 - 900 4788 , sankyusystem@gmail.com" 'Watermark Sankyu System
    End With
        
    ' This makes Excel visible
    xlObject.Visible = True
    xlObject.EnableEvents = True

End If
End Sub

Private Sub Form_Load()
'On Error Resume Next
Call Frm106_initial_setting
End Sub
Private Sub L3_Text_Click()
'On Error Resume Next
Frm105.Show
Frm106.Hide
End Sub
Private Sub L4_Text_Click()
'on error resume next
Call Frm105_initial_setting
Call Frm105_debit_setting

Frm105.L7_Text.Visible = False
Frm105.L8_Text.Visible = False
Frm105.L108_Text.Visible = False

Frm105.Pic1.Visible = True

Frm105.Show
Frm106.Hide
End Sub
Private Sub L5_Text_Click()
'on error resume next
Call Frm105_initial_setting

Frm105.Pic3.Visible = False
Frm105.Pic4.Visible = False
Frm105.Pic5.Visible = False
Frm105.Pic6.Visible = False
Frm105.Pic7.Visible = False
Frm105.Pic8.Visible = False

Frm105.Pic2.Visible = True

Frm105.Show
Frm106.Hide
End Sub
Private Sub L6_Text_Click()
'on error resume next
Call Frm105_initial_setting

Frm105.Pic10.Visible = False
Frm105.Pic11.Visible = False
Frm105.Pic12.Visible = False
Frm105.Pic13.Visible = False
Frm105.Pic14.Visible = False

Frm105.Pic9.Visible = True

Frm105.Show
Frm106.Hide
End Sub

