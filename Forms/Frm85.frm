VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm85 
   Caption         =   "Report Belian / Jualan / Stok / Inventori"
   ClientHeight    =   13035
   ClientLeft      =   60
   ClientTop       =   -68610
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
   Icon            =   "Frm85.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13035
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.Timer Tmr1 
      Interval        =   10
      Left            =   6480
      Top             =   0
   End
   Begin VB.PictureBox Pic13 
      BorderStyle     =   0  'None
      Height          =   11055
      Left            =   4320
      ScaleHeight     =   11055
      ScaleWidth      =   11385
      TabIndex        =   100
      Top             =   360
      Visible         =   0   'False
      Width           =   11385
      Begin VB.CommandButton CMD23 
         BackColor       =   &H00FFFFFF&
         Height          =   840
         Left            =   8040
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm85.frx":0ECA
         MousePointer    =   99  'Custom
         Picture         =   "Frm85.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   115
         ToolTipText     =   "Paparan sebelumnya"
         Top             =   10080
         Width           =   1500
      End
      Begin VB.CommandButton CMD24 
         BackColor       =   &H00FFFFFF&
         Height          =   840
         Left            =   9720
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm85.frx":1B13
         MousePointer    =   99  'Custom
         Picture         =   "Frm85.frx":1E1D
         Style           =   1  'Graphical
         TabIndex        =   114
         ToolTipText     =   "Paparan seterusnya"
         Top             =   10080
         Width           =   1500
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid10 
         Height          =   9525
         Left            =   120
         TabIndex        =   101
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   360
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   16801
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
      Begin VB.Label L83_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L83_Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   105
         Top             =   10260
         Width           =   1695
      End
      Begin VB.Label L82_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L82_Text"
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
         Height          =   300
         Left            =   240
         TabIndex        =   104
         Top             =   60
         Width           =   22650
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Ringkasan bagi data keseluruhan."
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
         Left            =   120
         TabIndex        =   103
         Top             =   9990
         Width           =   12135
      End
      Begin VB.Label L84_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L84_Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   102
         Top             =   10440
         Width           =   1935
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Berat (g)             :   Jumlah Harga (RM)         :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   106
         Top             =   10260
         Width           =   1935
      End
   End
   Begin VB.PictureBox Pic12 
      BorderStyle     =   0  'None
      Height          =   11055
      Left            =   4200
      ScaleHeight     =   11055
      ScaleWidth      =   21225
      TabIndex        =   83
      Top             =   120
      Visible         =   0   'False
      Width           =   21225
      Begin VB.CommandButton CMD21 
         BackColor       =   &H00FFFFFF&
         Height          =   840
         Left            =   17280
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm85.frx":2743
         MousePointer    =   99  'Custom
         Picture         =   "Frm85.frx":2A4D
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Paparan sebelumnya"
         Top             =   10080
         Width           =   1785
      End
      Begin VB.CommandButton CMD22 
         BackColor       =   &H00FFFFFF&
         Height          =   840
         Left            =   19200
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm85.frx":338C
         MousePointer    =   99  'Custom
         Picture         =   "Frm85.frx":3696
         Style           =   1  'Graphical
         TabIndex        =   116
         ToolTipText     =   "Paparan seterusnya"
         Top             =   10080
         Width           =   1785
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid9 
         Height          =   9525
         Left            =   120
         TabIndex        =   84
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   480
         Width           =   20865
         _ExtentX        =   36804
         _ExtentY        =   16801
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
      Begin VB.Label L75_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L75_Text"
         Height          =   255
         Left            =   2160
         TabIndex        =   95
         Top             =   10755
         Width           =   2175
      End
      Begin VB.Label L74_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L74_Text"
         Height          =   255
         Left            =   2160
         TabIndex        =   93
         Top             =   10515
         Width           =   2175
      End
      Begin VB.Label L73_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L73_Text"
         Height          =   255
         Left            =   2160
         TabIndex        =   91
         Top             =   10275
         Width           =   1455
      End
      Begin VB.Label L72_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L72_Text"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   89
         Top             =   120
         Width           =   22650
      End
      Begin VB.Label Label56 
         BackStyle       =   0  'Transparent
         Caption         =   "Ringkasan bagi data yang sedang dipaparkan.                                                     Ringkasan bagi data keseluruhan."
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
         TabIndex        =   88
         Top             =   9990
         Width           =   12135
      End
      Begin VB.Label L76_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L76_Text"
         Height          =   255
         Left            =   9840
         TabIndex        =   87
         Top             =   10275
         Width           =   1455
      End
      Begin VB.Label L77_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L77_Text"
         Height          =   255
         Left            =   9840
         TabIndex        =   86
         Top             =   10515
         Width           =   2175
      End
      Begin VB.Label L78_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L78_Text"
         Height          =   255
         Left            =   9840
         TabIndex        =   85
         Top             =   10755
         Width           =   2175
      End
      Begin VB.Label Label66 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm85.frx":3FBC
         Height          =   255
         Left            =   240
         TabIndex        =   94
         Top             =   10755
         Width           =   10335
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm85.frx":404F
         Height          =   255
         Left            =   240
         TabIndex        =   92
         Top             =   10515
         Width           =   10335
      End
      Begin VB.Label Label59 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm85.frx":40E4
         Height          =   255
         Left            =   240
         TabIndex        =   90
         Top             =   10275
         Width           =   10335
      End
   End
   Begin VB.PictureBox Pic9 
      BorderStyle     =   0  'None
      Height          =   11055
      Left            =   360
      ScaleHeight     =   11055
      ScaleWidth      =   17145
      TabIndex        =   28
      Top             =   480
      Visible         =   0   'False
      Width           =   17145
      Begin VB.CommandButton CMD15 
         BackColor       =   &H00FFFFFF&
         Height          =   840
         Left            =   13080
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm85.frx":4179
         MousePointer    =   99  'Custom
         Picture         =   "Frm85.frx":4483
         Style           =   1  'Graphical
         TabIndex        =   111
         ToolTipText     =   "Paparan sebelumnya"
         Top             =   10080
         Width           =   1785
      End
      Begin VB.CommandButton CMD16 
         BackColor       =   &H00FFFFFF&
         Height          =   840
         Left            =   15000
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm85.frx":4DC2
         MousePointer    =   99  'Custom
         Picture         =   "Frm85.frx":50CC
         Style           =   1  'Graphical
         TabIndex        =   110
         ToolTipText     =   "Paparan seterusnya"
         Top             =   10080
         Width           =   1785
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid6 
         Height          =   9525
         Left            =   120
         TabIndex        =   29
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   480
         Width           =   16665
         _ExtentX        =   29395
         _ExtentY        =   16801
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
      Begin VB.Label L61_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L61_Text"
         Height          =   255
         Left            =   9840
         TabIndex        =   65
         Top             =   10755
         Width           =   2175
      End
      Begin VB.Label L60_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L60_Text"
         Height          =   255
         Left            =   9840
         TabIndex        =   64
         Top             =   10515
         Width           =   2175
      End
      Begin VB.Label L59_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L59_Text"
         Height          =   255
         Left            =   9840
         TabIndex        =   63
         Top             =   10275
         Width           =   1455
      End
      Begin VB.Label Label72 
         BackStyle       =   0  'Transparent
         Caption         =   "Ringkasan bagi data yang sedang dipaparkan.                                                   Ringkasan bagi data keseluruhan."
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
         TabIndex        =   62
         Top             =   9990
         Width           =   12615
      End
      Begin VB.Label L38_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L38_Text"
         Height          =   255
         Left            =   2160
         TabIndex        =   36
         Top             =   10755
         Width           =   2175
      End
      Begin VB.Label L37_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L37_Text"
         Height          =   255
         Left            =   2160
         TabIndex        =   34
         Top             =   10515
         Width           =   2175
      End
      Begin VB.Label L36_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L36_Text"
         Height          =   255
         Left            =   2160
         TabIndex        =   32
         Top             =   10275
         Width           =   1455
      End
      Begin VB.Label L35_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L35_Text"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   30
         Top             =   120
         Width           =   22650
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm85.frx":59F2
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   10755
         Width           =   9615
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm85.frx":5A85
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   10515
         Width           =   10335
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm85.frx":5B1A
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   10275
         Width           =   10095
      End
   End
   Begin VB.PictureBox Pic3 
      BorderStyle     =   0  'None
      Height          =   11055
      Left            =   1200
      ScaleHeight     =   11055
      ScaleWidth      =   21225
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   21225
      Begin VB.CommandButton CMD7 
         Caption         =   "Back"
         Height          =   810
         Left            =   18240
         MouseIcon       =   "Frm85.frx":5BAF
         MousePointer    =   99  'Custom
         Picture         =   "Frm85.frx":5EB9
         Style           =   1  'Graphical
         TabIndex        =   147
         ToolTipText     =   "Paparan Sebelum"
         Top             =   10200
         Width           =   1095
      End
      Begin VB.CommandButton CMD8 
         Caption         =   "Next"
         Height          =   810
         Left            =   19440
         MouseIcon       =   "Frm85.frx":6F83
         MousePointer    =   99  'Custom
         Picture         =   "Frm85.frx":728D
         Style           =   1  'Graphical
         TabIndex        =   146
         ToolTipText     =   "Paparan Seterusnya"
         Top             =   10200
         Width           =   1095
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   9690
         Left            =   240
         TabIndex        =   130
         Top             =   360
         Width           =   20415
         _ExtentX        =   36010
         _ExtentY        =   17092
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
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "*** Pengiraan keuntungan ini adalah termasuk dengan GST."
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
         Left            =   11880
         TabIndex        =   109
         Top             =   10320
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label L88_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L88_Text"
         Height          =   255
         Left            =   6825
         TabIndex        =   107
         Top             =   10800
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label L47_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L47_Text"
         Height          =   255
         Left            =   2505
         TabIndex        =   50
         Top             =   10245
         Width           =   2655
      End
      Begin VB.Label L48_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L48_Text"
         Height          =   255
         Left            =   2505
         TabIndex        =   49
         Top             =   10485
         Width           =   2655
      End
      Begin VB.Label L49_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L49_Text"
         Height          =   255
         Left            =   2505
         TabIndex        =   48
         Top             =   10725
         Width           =   2655
      End
      Begin VB.Label L50_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L50_Text"
         Height          =   255
         Left            =   6825
         TabIndex        =   47
         Top             =   10575
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label45 
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
         TabIndex        =   46
         Top             =   10050
         Width           =   10935
      End
      Begin VB.Label L14_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L14_Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   9
         Top             =   0
         Width           =   20445
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Keuntungan        :  "
         Height          =   255
         Left            =   8280
         TabIndex        =   13
         Top             =   10680
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Harga                :  "
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   10725
         Width           =   2655
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Berat                 :  "
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   10485
         Width           =   2655
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan Item                :                                                               "
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   10245
         Width           =   2775
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Keuntungan 2     :  "
         Height          =   255
         Left            =   8760
         TabIndex        =   108
         Top             =   10440
         Visible         =   0   'False
         Width           =   3015
      End
   End
   Begin VB.PictureBox Pic10 
      BorderStyle     =   0  'None
      Height          =   11055
      Left            =   3720
      ScaleHeight     =   11055
      ScaleWidth      =   16305
      TabIndex        =   37
      Top             =   2160
      Visible         =   0   'False
      Width           =   16305
      Begin VB.CommandButton CMD17 
         Caption         =   "Back"
         Height          =   810
         Left            =   13680
         MouseIcon       =   "Frm85.frx":8357
         MousePointer    =   99  'Custom
         Picture         =   "Frm85.frx":8661
         Style           =   1  'Graphical
         TabIndex        =   145
         ToolTipText     =   "Paparan Sebelum"
         Top             =   10080
         Width           =   1095
      End
      Begin VB.CommandButton CMD18 
         Caption         =   "Next"
         Height          =   810
         Left            =   14880
         MouseIcon       =   "Frm85.frx":972B
         MousePointer    =   99  'Custom
         Picture         =   "Frm85.frx":9A35
         Style           =   1  'Graphical
         TabIndex        =   144
         ToolTipText     =   "Paparan Seterusnya"
         Top             =   10080
         Width           =   1095
      End
      Begin MSComctlLib.ListView LV6 
         Height          =   9525
         Left            =   120
         TabIndex        =   143
         Top             =   360
         Width           =   15975
         _ExtentX        =   28178
         _ExtentY        =   16801
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
      Begin VB.Label L62_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L62_Text"
         Height          =   255
         Left            =   2050
         TabIndex        =   69
         Top             =   10275
         Width           =   1455
      End
      Begin VB.Label L63_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L63_Text"
         Height          =   255
         Left            =   2050
         TabIndex        =   68
         Top             =   10515
         Width           =   2175
      End
      Begin VB.Label L64_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L64_Text"
         Height          =   255
         Left            =   2050
         TabIndex        =   67
         Top             =   10755
         Width           =   2175
      End
      Begin VB.Label Label75 
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
         TabIndex        =   66
         Top             =   9990
         Width           =   10815
      End
      Begin VB.Label L39_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L39_Text"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   41
         Top             =   120
         Width           =   22650
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan Item        :"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   10275
         Width           =   7575
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Berat         :"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   10515
         Width           =   8055
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Harga        :"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   10755
         Width           =   7575
      End
   End
   Begin VB.PictureBox Pic6 
      BorderStyle     =   0  'None
      Height          =   11055
      Left            =   4440
      ScaleHeight     =   11055
      ScaleWidth      =   20025
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   20025
      Begin VB.PictureBox Pic1 
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   5040
         ScaleHeight     =   3015
         ScaleWidth      =   7695
         TabIndex        =   118
         Top             =   3480
         Visible         =   0   'False
         Width           =   7695
         Begin VB.CommandButton CMD2 
            Caption         =   "Batal"
            Height          =   350
            Left            =   3720
            MouseIcon       =   "Frm85.frx":AAFF
            MousePointer    =   99  'Custom
            TabIndex        =   129
            ToolTipText     =   "Batal"
            Top             =   2400
            Width           =   1695
         End
         Begin VB.CommandButton CMD3 
            Caption         =   "Simpan Data"
            Height          =   350
            Left            =   1920
            MouseIcon       =   "Frm85.frx":AE09
            MousePointer    =   99  'Custom
            TabIndex        =   122
            ToolTipText     =   "Simpan Data"
            Top             =   2400
            Width           =   1695
         End
         Begin VB.TextBox TB1 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   1920
            TabIndex        =   119
            Text            =   "TB1"
            Top             =   1920
            Width           =   1140
         End
         Begin VB.Label L92_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L92_Text"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1920
            TabIndex        =   128
            Top             =   1680
            Width           =   4905
         End
         Begin VB.Label L91_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L91_Text"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1920
            TabIndex        =   127
            Top             =   1440
            Width           =   4905
         End
         Begin VB.Label L90_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L90_Text"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1920
            TabIndex        =   126
            Top             =   720
            Width           =   4905
         End
         Begin VB.Label L89_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L89_Text"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1920
            TabIndex        =   125
            Top             =   480
            Width           =   4905
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   ": :     : : :"
            ForeColor       =   &H00000000&
            Height          =   1845
            Left            =   1800
            TabIndex        =   124
            Top             =   480
            Width           =   105
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   $"Frm85.frx":B113
            ForeColor       =   &H00000000&
            Height          =   1845
            Left            =   240
            TabIndex        =   123
            Top             =   480
            Width           =   1545
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Susut berat selepas barang kemas dipotong."
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
            TabIndex        =   121
            Top             =   120
            Width           =   7185
         End
         Begin VB.Label L4_Text 
            Alignment       =   2  'Center
            BackColor       =   &H8000000A&
            Caption         =   "L4_Text"
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
            TabIndex        =   120
            Top             =   2520
            Visible         =   0   'False
            Width           =   1575
         End
      End
      Begin VB.CommandButton CMD14 
         Caption         =   "Next"
         Height          =   810
         Left            =   18600
         MouseIcon       =   "Frm85.frx":B1A9
         MousePointer    =   99  'Custom
         Picture         =   "Frm85.frx":B4B3
         Style           =   1  'Graphical
         TabIndex        =   142
         ToolTipText     =   "Paparan Seterusnya"
         Top             =   10080
         Width           =   1095
      End
      Begin VB.CommandButton CMD13 
         Caption         =   "Back"
         Height          =   810
         Left            =   17400
         MouseIcon       =   "Frm85.frx":C57D
         MousePointer    =   99  'Custom
         Picture         =   "Frm85.frx":C887
         Style           =   1  'Graphical
         TabIndex        =   141
         ToolTipText     =   "Paparan Sebelum"
         Top             =   10080
         Width           =   1095
      End
      Begin MSComctlLib.ListView LV5 
         Height          =   9525
         Left            =   120
         TabIndex        =   140
         Top             =   480
         Width           =   19665
         _ExtentX        =   34687
         _ExtentY        =   16801
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
      Begin VB.Label L58_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L58_Text"
         Height          =   255
         Left            =   2200
         TabIndex        =   61
         Top             =   10515
         Width           =   2175
      End
      Begin VB.Label L57_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L57_Text"
         Height          =   255
         Left            =   2200
         TabIndex        =   60
         Top             =   10275
         Width           =   1455
      End
      Begin VB.Label Label65 
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
         TabIndex        =   59
         Top             =   9990
         Width           =   11895
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Berat           :"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   10515
         Width           =   2055
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan Item          :"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   10275
         Width           =   2055
      End
      Begin VB.Label L27_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L27_Text"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   25
         Top             =   120
         Width           =   22650
      End
   End
   Begin VB.PictureBox Pic4 
      BorderStyle     =   0  'None
      Height          =   11055
      Left            =   3240
      ScaleHeight     =   11055
      ScaleWidth      =   20745
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   20745
      Begin VB.CommandButton CMD9 
         Caption         =   "Back"
         Height          =   810
         Left            =   18240
         MouseIcon       =   "Frm85.frx":D951
         MousePointer    =   99  'Custom
         Picture         =   "Frm85.frx":DC5B
         Style           =   1  'Graphical
         TabIndex        =   139
         ToolTipText     =   "Paparan Sebelum"
         Top             =   9960
         Width           =   1095
      End
      Begin VB.CommandButton CMD10 
         Caption         =   "Next"
         Height          =   810
         Left            =   19440
         MouseIcon       =   "Frm85.frx":ED25
         MousePointer    =   99  'Custom
         Picture         =   "Frm85.frx":F02F
         Style           =   1  'Graphical
         TabIndex        =   138
         ToolTipText     =   "Paparan Seterusnya"
         Top             =   9960
         Width           =   1095
      End
      Begin MSComctlLib.ListView LV3 
         Height          =   9525
         Left            =   120
         TabIndex        =   132
         Top             =   360
         Width           =   20415
         _ExtentX        =   36010
         _ExtentY        =   16801
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
      Begin VB.Label L19_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L19_Text"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Width           =   22650
      End
      Begin VB.Label L51_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L51_Text"
         Height          =   255
         Left            =   1680
         TabIndex        =   54
         Top             =   10275
         Width           =   2655
      End
      Begin VB.Label L52_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L52_Text"
         Height          =   255
         Left            =   1680
         TabIndex        =   53
         Top             =   10515
         Width           =   2655
      End
      Begin VB.Label L53_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L53_Text"
         Height          =   255
         Left            =   1680
         TabIndex        =   52
         Top             =   10755
         Width           =   2655
      End
      Begin VB.Label Label52 
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
         TabIndex        =   51
         Top             =   9990
         Width           =   12495
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Modal :"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   10755
         Width           =   1335
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Berat :"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   10515
         Width           =   1335
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan Item :"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   10275
         Width           =   1335
      End
   End
   Begin VB.PictureBox Pic11 
      BorderStyle     =   0  'None
      Height          =   11055
      Left            =   8760
      ScaleHeight     =   11055
      ScaleWidth      =   21585
      TabIndex        =   70
      Top             =   1680
      Visible         =   0   'False
      Width           =   21585
      Begin VB.CommandButton CMD19 
         BackColor       =   &H00FFFFFF&
         Height          =   840
         Left            =   17640
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm85.frx":100F9
         MousePointer    =   99  'Custom
         Picture         =   "Frm85.frx":10403
         Style           =   1  'Graphical
         TabIndex        =   113
         ToolTipText     =   "Paparan sebelumnya"
         Top             =   10200
         Width           =   1785
      End
      Begin VB.CommandButton CMD20 
         BackColor       =   &H00FFFFFF&
         Height          =   840
         Left            =   19560
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm85.frx":10D42
         MousePointer    =   99  'Custom
         Picture         =   "Frm85.frx":1104C
         Style           =   1  'Graphical
         TabIndex        =   112
         ToolTipText     =   "Paparan seterusnya"
         Top             =   10200
         Width           =   1785
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid8 
         Height          =   9525
         Left            =   120
         TabIndex        =   71
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   480
         Width           =   21225
         _ExtentX        =   37439
         _ExtentY        =   16801
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
      Begin VB.Label L67_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L67_Text"
         Height          =   255
         Left            =   2160
         TabIndex        =   82
         Top             =   10755
         Width           =   2175
      End
      Begin VB.Label L66_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L66_Text"
         Height          =   255
         Left            =   2160
         TabIndex        =   80
         Top             =   10515
         Width           =   2175
      End
      Begin VB.Label L65_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L65_Text"
         Height          =   255
         Left            =   2160
         TabIndex        =   78
         Top             =   10275
         Width           =   1455
      End
      Begin VB.Label L71_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L71_Text"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   76
         Top             =   120
         Width           =   22650
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Ringkasan bagi data yang sedang dipaparkan.                                                     Ringkasan bagi data keseluruhan."
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
         TabIndex        =   75
         Top             =   9990
         Width           =   12135
      End
      Begin VB.Label L68_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L68_Text"
         Height          =   255
         Left            =   9840
         TabIndex        =   74
         Top             =   10275
         Width           =   1455
      End
      Begin VB.Label L69_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L69_Text"
         Height          =   255
         Left            =   9840
         TabIndex        =   73
         Top             =   10515
         Width           =   2175
      End
      Begin VB.Label L70_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L70_Text"
         Height          =   255
         Left            =   9840
         TabIndex        =   72
         Top             =   10755
         Width           =   2175
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm85.frx":11972
         Height          =   255
         Left            =   240
         TabIndex        =   81
         Top             =   10755
         Width           =   10335
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm85.frx":11A05
         Height          =   255
         Left            =   240
         TabIndex        =   79
         Top             =   10515
         Width           =   10335
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm85.frx":11A9A
         Height          =   255
         Left            =   240
         TabIndex        =   77
         Top             =   10275
         Width           =   10335
      End
   End
   Begin VB.PictureBox Pic2 
      BorderStyle     =   0  'None
      Height          =   11055
      Left            =   240
      ScaleHeight     =   11055
      ScaleWidth      =   20745
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   20745
      Begin VB.CommandButton CMD5 
         Caption         =   "Next"
         Height          =   810
         Left            =   19440
         MouseIcon       =   "Frm85.frx":11B2F
         MousePointer    =   99  'Custom
         Picture         =   "Frm85.frx":11E39
         Style           =   1  'Graphical
         TabIndex        =   134
         ToolTipText     =   "Paparan Seterusnya"
         Top             =   10080
         Width           =   1095
      End
      Begin VB.CommandButton CMD6 
         Caption         =   "Back"
         Height          =   810
         Left            =   18240
         MouseIcon       =   "Frm85.frx":12F03
         MousePointer    =   99  'Custom
         Picture         =   "Frm85.frx":1320D
         Style           =   1  'Graphical
         TabIndex        =   133
         ToolTipText     =   "Paparan Sebelum"
         Top             =   10080
         Width           =   1095
      End
      Begin MSComctlLib.ListView LV2 
         Height          =   9525
         Left            =   120
         TabIndex        =   131
         Top             =   480
         Width           =   20415
         _ExtentX        =   36010
         _ExtentY        =   16801
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
      Begin VB.Label L46_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L46_Text"
         Height          =   255
         Left            =   1800
         TabIndex        =   45
         Top             =   10755
         Width           =   3255
      End
      Begin VB.Label L45_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L45_Text"
         Height          =   255
         Left            =   1800
         TabIndex        =   44
         Top             =   10515
         Width           =   3255
      End
      Begin VB.Label L44_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L44_Text"
         Height          =   255
         Left            =   1800
         TabIndex        =   43
         Top             =   10275
         Width           =   3255
      End
      Begin VB.Label Label37 
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
         TabIndex        =   42
         Top             =   9990
         Width           =   12150
      End
      Begin VB.Label L10_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L10_Text"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   22650
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan Item :"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   10275
         Width           =   1455
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Berat :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   10515
         Width           =   1455
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Modal :"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   10755
         Width           =   1455
      End
   End
   Begin VB.PictureBox Pic5 
      BorderStyle     =   0  'None
      Height          =   11055
      Left            =   1680
      ScaleHeight     =   11055
      ScaleWidth      =   20865
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   20865
      Begin VB.CommandButton CMD11 
         Caption         =   "Back"
         Height          =   810
         Left            =   18360
         MouseIcon       =   "Frm85.frx":142D7
         MousePointer    =   99  'Custom
         Picture         =   "Frm85.frx":145E1
         Style           =   1  'Graphical
         TabIndex        =   137
         ToolTipText     =   "Paparan Sebelum"
         Top             =   9960
         Width           =   1095
      End
      Begin VB.CommandButton CMD12 
         Caption         =   "Next"
         Height          =   810
         Left            =   19560
         MouseIcon       =   "Frm85.frx":156AB
         MousePointer    =   99  'Custom
         Picture         =   "Frm85.frx":159B5
         Style           =   1  'Graphical
         TabIndex        =   136
         ToolTipText     =   "Paparan Seterusnya"
         Top             =   9960
         Width           =   1095
      End
      Begin MSComctlLib.ListView LV4 
         Height          =   9525
         Left            =   240
         TabIndex        =   135
         Top             =   360
         Width           =   20415
         _ExtentX        =   36010
         _ExtentY        =   16801
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
      Begin VB.Label L23_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L23_Text"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   23
         Top             =   120
         Width           =   22650
      End
      Begin VB.Label L56_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L56_Text"
         Height          =   255
         Left            =   1680
         TabIndex        =   58
         Top             =   10755
         Width           =   3015
      End
      Begin VB.Label L55_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L55_Text"
         Height          =   255
         Left            =   1680
         TabIndex        =   57
         Top             =   10515
         Width           =   3015
      End
      Begin VB.Label L54_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L54_Text"
         Height          =   255
         Left            =   1680
         TabIndex        =   56
         Top             =   10275
         Width           =   3015
      End
      Begin VB.Label Label57 
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
         TabIndex        =   55
         Top             =   9990
         Width           =   11895
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan Item :"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   10275
         Width           =   1335
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Berat :"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   10515
         Width           =   1335
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Modal :"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   10755
         Width           =   1335
      End
   End
   Begin VB.Label L81_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L81_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   2820
      TabIndex        =   99
      Top             =   11400
      Width           =   420
   End
   Begin VB.Label L80_Text 
      Caption         =   "L80_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3840
      TabIndex        =   98
      Top             =   11520
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label L79_Text 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "L79_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   2280
      TabIndex        =   96
      Top             =   11400
      Width           =   420
   End
   Begin VB.Label L3_Text 
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
      MouseIcon       =   "Frm85.frx":16A7F
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   0
      Width           =   1815
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
      Left            =   21600
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   2100
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
      Left            =   21600
      TabIndex        =   0
      Top             =   435
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Paparan Muka         :         /"
      Height          =   255
      Left            =   360
      TabIndex        =   97
      Top             =   11400
      Width           =   3255
   End
   Begin VB.Menu Frm85_PM_Menu11 
      Caption         =   "Susut Nilai"
      Visible         =   0   'False
      Begin VB.Menu Frm85_SM_Export1_Overall_11 
         Caption         =   "Export Excel Report (Keseluruhan)"
      End
   End
   Begin VB.Menu Frm85_PM_Menu1 
      Caption         =   "Belian"
      Visible         =   0   'False
      Begin VB.Menu Frm85_SM_Edit_Data_Belian 
         Caption         =   "Edit Data Belian"
      End
      Begin VB.Menu frm85_sm_spacer_6 
         Caption         =   "-"
      End
      Begin VB.Menu Frm85_SM_Padam_Data2 
         Caption         =   "Padam Data Ini"
      End
      Begin VB.Menu frm85_sm_spacer_1 
         Caption         =   "-"
      End
      Begin VB.Menu Frm85_SM_edit_supplier 
         Caption         =   "Tukar Maklumat Supplier Bagi Belian Item Ini"
         Visible         =   0   'False
      End
      Begin VB.Menu Frm85_SM_Print_Barcode2 
         Caption         =   "Print Barcode Item Ini"
      End
      Begin VB.Menu frm85_sm_spacer_2 
         Caption         =   "-"
      End
      Begin VB.Menu Frm85_SM_Export1_Overall_1 
         Caption         =   "Export Excel Report (Keseluruhan)"
      End
   End
   Begin VB.Menu Frm85_PM_Menu2 
      Caption         =   "Jualan"
      Visible         =   0   'False
      Begin VB.Menu Frm85_SM_Edit_Data_Jualan 
         Caption         =   "Edit Data Jualan"
      End
      Begin VB.Menu frm85_sm_spacer_5 
         Caption         =   "-"
      End
      Begin VB.Menu Frm85_SM_Padam_Data 
         Caption         =   "Padam Invoice"
      End
      Begin VB.Menu frm85_sm_spacer_3 
         Caption         =   "-"
      End
      Begin VB.Menu frm85_sm_email_jualan 
         Caption         =   "Hantar E-mail Jualan"
      End
      Begin VB.Menu Frm85_SM_cetak_resit_Jualan 
         Caption         =   "Cetak Invoice Jualan"
      End
      Begin VB.Menu Frm85_SM_cetak_voucher 
         Caption         =   "Cetak Voucher / Statement (Bagi jualan kepada agen sahaja)"
         Visible         =   0   'False
      End
      Begin VB.Menu frm85_sm_spacer_4 
         Caption         =   "-"
      End
      Begin VB.Menu Frm85_SM_Export1_Overall_2 
         Caption         =   "Export Excel Report (Keseluruhan)"
      End
   End
   Begin VB.Menu Frm85_PM_Menu3 
      Caption         =   "Trade In"
      Visible         =   0   'False
      Begin VB.Menu Frm85_SM_Edit_Data_Buyback 
         Caption         =   "Edit Data Buyback / Trade In"
      End
      Begin VB.Menu frm85_sm_spacer_7 
         Caption         =   "-"
      End
      Begin VB.Menu Frm85_SM_Padam_Data3 
         Caption         =   "Padam Data Ini"
      End
      Begin VB.Menu frm85_sm_spacer_8 
         Caption         =   "-"
      End
      Begin VB.Menu Frm85_SM_Cetak_Resit_Buyback 
         Caption         =   "Cetak Voucher Buyback / Trade In"
      End
      Begin VB.Menu Frm85_SM_Print_Barcode_buyback 
         Caption         =   "Print Barcode Dari Semua Penerimaan Stok Ini"
      End
      Begin VB.Menu Frm85_SM_Print_Barcode2_buyback 
         Caption         =   "Print Barcode Item Ini"
      End
      Begin VB.Menu frm85_sm_spacer_9 
         Caption         =   "-"
      End
      Begin VB.Menu Frm85_SM_Export1_Overall_3 
         Caption         =   "Export Excel Report (Keseluruhan)"
      End
   End
   Begin VB.Menu Frm85_PM_Menu4 
      Caption         =   "Stok"
      Visible         =   0   'False
      Begin VB.Menu Frm85_SM_Export1_Overall_4 
         Caption         =   "Export Excel Report (Keseluruhan)"
      End
      Begin VB.Menu frm85_sm_spacer_10 
         Caption         =   "-"
      End
      Begin VB.Menu frm85_sm_hilang 
         Caption         =   "Hilang , kecurian etc"
      End
   End
   Begin VB.Menu Frm85_PM_Menu5 
      Caption         =   "Potong"
      Visible         =   0   'False
      Begin VB.Menu Frm85_SM_susut_berat 
         Caption         =   "Ubah data susut berat"
      End
      Begin VB.Menu frm85_sm_spacer_11 
         Caption         =   "-"
      End
      Begin VB.Menu Frm85_SM_Export1_Overall_7 
         Caption         =   "Export Excel Report (Keseluruhan)"
      End
   End
   Begin VB.Menu Frm85_PM_Menu6 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm85_SM_Edit_Data_Belian_gb 
         Caption         =   "Edit Data Belian"
      End
      Begin VB.Menu Frm85_SM_edit_supplier3 
         Caption         =   "Tukar Maklumat Supplier Bagi Belian Item Ini"
         Visible         =   0   'False
      End
      Begin VB.Menu Frm85_SM_Padam_Data4 
         Caption         =   "Padam Data Ini"
      End
      Begin VB.Menu Frm85_SM_Print_Barcode3 
         Caption         =   "Print Barcode Dari Semua Penerimaan Stok Ini"
      End
      Begin VB.Menu Frm85_SM_Print_Barcode4 
         Caption         =   "Print Barcode Item Ini"
      End
      Begin VB.Menu Frm85_SM_Export6 
         Caption         =   "Export Excel Report (Paparan Ini Sahaja)"
      End
      Begin VB.Menu Frm85_SM_Export1_Overall_5 
         Caption         =   "Export Excel Report (Keseluruhan)"
      End
   End
   Begin VB.Menu Frm85_PM_Menu7 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm85_SM_Edit_Data_buyback_gb 
         Caption         =   "Edit Data Buyback / Trade In"
      End
      Begin VB.Menu Frm85_SM_Padam_Data5 
         Caption         =   "Padam Data Ini"
      End
      Begin VB.Menu Frm85_SM_Cetak_Resit_Buyback_gb 
         Caption         =   "Cetak Voucher Buyback / Trade In"
      End
      Begin VB.Menu Frm85_SM_Print_Barcode_buyback_gb 
         Caption         =   "Print Barcode Dari Semua Penerimaan Stok Ini"
      End
      Begin VB.Menu Frm85_SM_Print_Barcode2_buyback_gb 
         Caption         =   "Print Barcode Item Ini"
      End
      Begin VB.Menu Frm85_SM_Export7 
         Caption         =   "Export Excel Report (Paparan Ini Sahaja)"
      End
      Begin VB.Menu Frm85_SM_Export1_Overall_6 
         Caption         =   "Export Excel Report (Keseluruhan)"
      End
   End
   Begin VB.Menu Frm85_PM_Menu8 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm85_SM_Export1_Overall_8 
         Caption         =   "Export Excel Report (Keseluruhan)"
      End
   End
   Begin VB.Menu Frm85_PM_Menu9 
      Caption         =   "Tempahan"
      Visible         =   0   'False
      Begin VB.Menu Frm85_SM_Export1_Overall_9 
         Caption         =   "Export Excel Report (Keseluruhan)"
      End
   End
   Begin VB.Menu Frm85_PM_Menu10 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm85_SM_Export1_Overall_10 
         Caption         =   "Export Excel Report (Keseluruhan)"
      End
   End
End
Attribute VB_Name = "Frm85"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB10_Click()
'on error resume next
If Frm101.CB10 = 1 Then
    Frm101.CB3 = 0
    Frm101.CB4 = 0
    Frm101.CB5 = 0
    Frm101.CB2 = 0
    Frm101.CB9 = 0
    Frm101.CB6 = 0
    Frm101.CB11 = 0
    Frm101.CB12 = 0
End If
End Sub
Private Sub CB11_Click()
'on error resume next
If Frm101.CB11 = 1 Then
    Frm101.CB3 = 0
    Frm101.CB4 = 0
    Frm101.CB5 = 0
    Frm101.CB2 = 0
    Frm101.CB9 = 0
    Frm101.CB10 = 0
    Frm101.CB6 = 0
    Frm101.CB12 = 0
End If
End Sub
Private Sub CB12_Click()
'on error resume next
If Frm101.CB12 = 1 Then
    Frm101.CB3 = 0
    Frm101.CB4 = 0
    Frm101.CB5 = 0
    Frm101.CB2 = 0
    Frm101.CB9 = 0
    Frm101.CB10 = 0
    Frm101.CB11 = 0
    Frm101.CB6 = 0
End If
End Sub
Private Sub CB2_Click()
'on error resume next
If Frm101.CB2 = 1 Then
    Frm101.CB3 = 0
    Frm101.CB4 = 0
    Frm101.CB5 = 0
    Frm101.CB6 = 0
    Frm101.CB9 = 0
    Frm101.CB10 = 0
    Frm101.CB11 = 0
    Frm101.CB12 = 0
End If
End Sub
Private Sub CB3_Click()
'on error resume next
If Frm101.CB3 = 1 Then
    Frm101.CB2 = 0
    Frm101.CB4 = 0
    Frm101.CB5 = 0
    Frm101.CB6 = 0
    Frm101.CB9 = 0
    Frm101.CB10 = 0
    Frm101.CB11 = 0
    Frm101.CB12 = 0
End If
End Sub
Private Sub CB4_Click()
'on error resume next
If Frm101.CB4 = 1 Then
    Frm101.CB3 = 0
    Frm101.CB2 = 0
    Frm101.CB5 = 0
    Frm101.CB6 = 0
    Frm101.CB9 = 0
    Frm101.CB10 = 0
    Frm101.CB11 = 0
    Frm101.CB12 = 0
End If
End Sub
Private Sub CB5_Click()
'on error resume next
If Frm101.CB5 = 1 Then
    Frm101.CB3 = 0
    Frm101.CB4 = 0
    Frm101.CB2 = 0
    Frm101.CB6 = 0
    Frm101.CB9 = 0
    Frm101.CB10 = 0
    Frm101.CB11 = 0
    Frm101.CB12 = 0
End If
End Sub
Private Sub CB6_Click()
'on error resume next
If Frm101.CB6 = 1 Then
    Frm101.CB3 = 0
    Frm101.CB4 = 0
    Frm101.CB5 = 0
    Frm101.CB2 = 0
    Frm101.CB9 = 0
    Frm101.CB10 = 0
    Frm101.CB11 = 0
    Frm101.CB12 = 0
End If
End Sub
Private Sub CB7_Click()
'on error resume next
If Frm101.CB7 = 1 Then
    Frm101.CB8 = 0
End If
End Sub
Private Sub CB8_Click()
'on error resume next
If Frm101.CB8 = 1 Then
    Frm101.CB7 = 0
End If
End Sub
Private Sub CB9_Click()
'on error resume next
If Frm101.CB9 = 1 Then
    Frm101.CB3 = 0
    Frm101.CB4 = 0
    Frm101.CB5 = 0
    Frm101.CB2 = 0
    Frm101.CB6 = 0
    Frm101.CB10 = 0
    Frm101.CB11 = 0
    Frm101.CB12 = 0
End If
End Sub
Private Sub CMD10_Click()
'on error resume next
Dim LM_CURRENT_PAGE As Single
Dim LM_PAGE_QTY As Single

If Frm85.L79_Text <> vbNullString And IsNumeric(Frm85.L79_Text) Then
    If Frm85.L81_Text <> vbNullString And IsNumeric(Frm85.L81_Text) Then
        LM_PAGE_QTY = Frm85.L81_Text
        LM_CURRENT_PAGE = Frm85.L79_Text
        
        If LM_CURRENT_PAGE < LM_PAGE_QTY Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            If GM_REPORT_MODE = 5 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_Buyback
                Call Frm85_report_buyback_page
            ElseIf GM_REPORT_MODE = 4 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_Buyback
                Call Frm85_carian_buyback_page
            ElseIf GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_Buyback
                Call Frm85_report_buyback_barcode
            End If
            
        End If
    End If
End If
End Sub
Private Sub CMD11_Click()
'on error resume next
Dim LM_CURRENT_PAGE As Single
Dim LM_PAGE_QTY As Single

LM_CURRENT_PAGE = 0
LM_PAGE_QTY = 0

If Frm85.L79_Text <> vbNullString And IsNumeric(Frm85.L79_Text) Then
    If Frm85.L81_Text <> vbNullString And IsNumeric(Frm85.L81_Text) Then
        LM_PAGE_QTY = Frm85.L81_Text
        LM_CURRENT_PAGE = Frm85.L79_Text
        
        If LM_CURRENT_PAGE <> 1 And LM_CURRENT_PAGE <> 0 Then
        
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            If GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_Stok
                Call Frm85_report_stok_barcode
            Else
                Call Frm85_Header_Report_Stok
                Call Frm85_report_stok_page
            End If
            
        End If
    End If
End If
End Sub
Private Sub CMD12_Click()
'on error resume next
Dim LM_CURRENT_PAGE As Single
Dim LM_PAGE_QTY As Single

If Frm85.L79_Text <> vbNullString And IsNumeric(Frm85.L79_Text) Then
    If Frm85.L81_Text <> vbNullString And IsNumeric(Frm85.L81_Text) Then
        LM_PAGE_QTY = Frm85.L81_Text
        LM_CURRENT_PAGE = Frm85.L79_Text
        
        If LM_CURRENT_PAGE < LM_PAGE_QTY Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            If GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_Stok
                Call Frm85_report_stok_barcode
            Else
                Call Frm85_Header_Report_Stok
                Call Frm85_report_stok_page
            End If
            
        End If
    End If
End If
End Sub
Private Sub CMD13_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Frm85.Pic1.Visible = False

If GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_Header_Report_Potong
    Call Frm85_report_potong_barcode
Else
    Call Frm85_Header_Report_Potong
    Call Frm85_report_potong_page
End If
End Sub
Private Sub CMD14_Click()
'on error resume next
Dim LM_CURRENT_PAGE As Single
Dim LM_PAGE_QTY As Single

Frm85.Pic1.Visible = False

If Frm85.L79_Text <> vbNullString And IsNumeric(Frm85.L79_Text) Then
    If Frm85.L81_Text <> vbNullString And IsNumeric(Frm85.L81_Text) Then
        LM_PAGE_QTY = Frm85.L81_Text
        LM_CURRENT_PAGE = Frm85.L79_Text
        
        If LM_CURRENT_PAGE < LM_PAGE_QTY Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            If GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_Potong
                Call Frm85_report_potong_barcode
            Else
                Call Frm85_Header_Report_Potong
                Call Frm85_report_potong_page
            End If
            
        End If
    End If
End If
End Sub
Private Sub CMD15_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

If GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_Header_Report_Ansuran
    Call Frm85_report_ansuran_barcode
Else
    Call Frm85_Header_Report_Ansuran
    Call Frm85_report_ansuran_page
End If
End Sub
Private Sub CMD16_Click()
'on error resume next
'on error resume next
Dim LM_CURRENT_PAGE As Single
Dim LM_PAGE_QTY As Single

If Frm85.L79_Text <> vbNullString And IsNumeric(Frm85.L79_Text) Then
    If Frm85.L81_Text <> vbNullString And IsNumeric(Frm85.L81_Text) Then
        LM_PAGE_QTY = Frm85.L81_Text
        LM_CURRENT_PAGE = Frm85.L79_Text
        
        If LM_CURRENT_PAGE < LM_PAGE_QTY Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            If GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_Ansuran
                Call Frm85_report_ansuran_barcode
            Else
                Call Frm85_Header_Report_Ansuran
                Call Frm85_report_ansuran_page
            End If
            
        End If
    End If
End If
End Sub
Private Sub CMD17_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

If GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_Header_Report_Tempahan
    Call Frm85_report_tempahan_barcode
Else
    Call Frm85_Header_Report_Tempahan
    Call Frm85_report_tempahan_page
End If
End Sub
Private Sub CMD18_Click()
'on error resume next
Dim LM_CURRENT_PAGE As Single
Dim LM_PAGE_QTY As Single

If Frm85.L79_Text <> vbNullString And IsNumeric(Frm85.L79_Text) Then
    If Frm85.L81_Text <> vbNullString And IsNumeric(Frm85.L81_Text) Then
        LM_PAGE_QTY = Frm85.L81_Text
        LM_CURRENT_PAGE = Frm85.L79_Text
        
        If LM_CURRENT_PAGE < LM_PAGE_QTY Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            If GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_Tempahan
                Call Frm85_report_tempahan_barcode
            Else
                Call Frm85_Header_Report_Tempahan
                Call Frm85_report_tempahan_page
            End If
            
        End If
    End If
End If
End Sub
Private Sub CMD19_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

If GM_REPORT_MODE = 6 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_Header_Report_belian_gb
    Call Frm85_report_belian_gb_page
ElseIf GM_REPORT_MODE = 7 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_Header_Report_belian_gb
    Call Frm85_carian_buyback_page
ElseIf GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_Header_Report_belian_gb
    Call Frm85_report_belian_gb_barcode
End If
End Sub
Private Sub CMD2_Click()
'on error resume next
Note = "Adakah anda ingin batalkan urusan ini?" & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then
    Frm85.Pic1.Visible = False
End If
End Sub
Private Sub CMD20_Click()
'on error resume next
Dim LM_CURRENT_PAGE As Single
Dim LM_PAGE_QTY As Single

If Frm85.L79_Text <> vbNullString And IsNumeric(Frm85.L79_Text) Then
    If Frm85.L81_Text <> vbNullString And IsNumeric(Frm85.L81_Text) Then
        LM_PAGE_QTY = Frm85.L81_Text
        LM_CURRENT_PAGE = Frm85.L79_Text
        
        If LM_CURRENT_PAGE < LM_PAGE_QTY Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            If GM_REPORT_MODE = 6 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_belian_gb
                Call Frm85_report_belian_gb_page
            ElseIf GM_REPORT_MODE = 7 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_belian_gb
                Call Frm85_carian_buyback_page
            ElseIf GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_belian_gb
                Call Frm85_report_belian_gb_barcode
            End If
            
        End If
    End If
End If
End Sub
Private Sub CMD21_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

If GM_REPORT_MODE = 6 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_Header_Report_belian_gb
    Call Frm85_report_belian_gb_page
ElseIf GM_REPORT_MODE = 7 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_Header_Report_buyback_gb
    Call Frm85_report_buyback_gb_page
ElseIf GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_Header_Report_buyback_gb
    Call Frm85_report_buyback_gb_barcode
End If
End Sub
Private Sub CMD22_Click()
'on error resume next
Dim LM_CURRENT_PAGE As Single
Dim LM_PAGE_QTY As Single

If Frm85.L79_Text <> vbNullString And IsNumeric(Frm85.L79_Text) Then
    If Frm85.L81_Text <> vbNullString And IsNumeric(Frm85.L81_Text) Then
        LM_PAGE_QTY = Frm85.L81_Text
        LM_CURRENT_PAGE = Frm85.L79_Text
        
        If LM_CURRENT_PAGE < LM_PAGE_QTY Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            If GM_REPORT_MODE = 6 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_belian_gb
                Call Frm85_report_belian_gb_page
            ElseIf GM_REPORT_MODE = 7 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_buyback_gb
                Call Frm85_report_buyback_gb_page
            ElseIf GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_buyback_gb
                Call Frm85_report_buyback_gb_barcode
            End If
            
        End If
    End If
End If
End Sub
Private Sub CMD23_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm85_header_report_trade_in_susut_nilai
Call Frm85_report_trade_in_susut_nilai
End Sub
Private Sub CMD24_Click()
'on error resume next
Dim LM_CURRENT_PAGE As Single
Dim LM_PAGE_QTY As Single

If Frm85.L79_Text <> vbNullString And IsNumeric(Frm85.L79_Text) Then
    If Frm85.L81_Text <> vbNullString And IsNumeric(Frm85.L81_Text) Then
        LM_PAGE_QTY = Frm85.L81_Text
        LM_CURRENT_PAGE = Frm85.L79_Text
        
        If LM_CURRENT_PAGE < LM_PAGE_QTY Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm85_header_report_trade_in_susut_nilai
            Call Frm85_report_trade_in_susut_nilai
            
        End If
    End If
End If
End Sub
Private Sub CMD3_Click()
'On Error Resume Next
Dim Err(4)
Dim Frm85_LM_BERAT_ASAL As Double
Dim Frm85_LM_BERAT_KINI As Double
Dim Frm85_LM_BERAT_SUSUT As Double
Dim Frm85_LM_BERAT_OVERALL As Double
Dim Frm85_LM_BERAT_KINI_BEFORE As Double

Frm85_LM_BERAT_KINI_BEFORE = 0
Frm85_LM_BERAT_ASAL = 0
Frm85_LM_BERAT_KINI = 0
Frm85_LM_BERAT_SUSUT = 0
Frm85_LM_BERAT_OVERALL = 0
            
DATA_SAVE = 0
DATA_UPDATE = 0

If Frm85.L4_Text = vbNullString Then
    x = x + 1
    Err(x) = "Kesilapan teknikal telah berlaku. Sila keluar dari menu ini dan cuba sekali lagi."
End If
If Frm85.L91_Text = vbNullString Or (Frm85.L91_Text <> vbNullString And Not IsNumeric(Frm85.L91_Text)) Then
    x = x + 1
    Err(x) = "Tiadak maklumat berat asal."
End If
If Frm85.L92_Text = vbNullString Or (Frm85.L92_Text <> vbNullString And Not IsNumeric(Frm85.L92_Text)) Then
    x = x + 1
    Err(x) = "Tiadak maklumat berat terkini."
End If
If Frm85.TB1 = vbNullString Or (Frm85.TB1 <> vbNullString And Not IsNumeric(Frm85.TB1)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Nilai susut berat]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
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
    
    If Answer = vbYes Then
    
        If Frm85.L91_Text <> vbNullString And Frm85.L92_Text <> vbNullString Then
            If IsNumeric(Frm85.L91_Text) Then Frm85_LM_BERAT_ASAL = Frm85.L91_Text 'Berat Asal
            If IsNumeric(Frm85.L92_Text) Then Frm85_LM_BERAT_KINI = Frm85.L92_Text 'Berat Terkini
            If IsNumeric(Frm85.TB1) Then Frm85_LM_BERAT_SUSUT = Frm85.TB1 'Susut Berat
            
            Frm85_LM_BERAT_OVERALL = Frm85_LM_BERAT_KINI + Frm85_LM_BERAT_SUSUT
            
            If Frm85_LM_BERAT_KINI < Frm85_LM_BERAT_SUSUT Then
            
                MsgBox "Susut berat yang dimasukkan melebihi berat yang dibenarkan." & vbCrLf & _
                        "Berat asal : " & Format(Frm85_LM_BERAT_ASAL, "#,##0.00 g") & vbCrLf & _
                        "Susut berat : " & Format(Frm85_LM_BERAT_SUSUT, "#,##0.00 g") & vbCrLf & _
                        "Susut berat maksimum yang dibenarkan adalah " & Format(Frm85_LM_BERAT_KINI, "#,##0.00 g"), vbInformation, "Info"
                        
                Exit Sub
                        
            End If
        End If
        
        LM_UBAH_STATUS = 0
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where ID='" & Frm85.L4_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!susut_berat) Then Frm85_LM_BERAT_KINI_BEFORE = rs!susut_berat
            If Not IsNull(rs!Berat) Then Frm85_LM_BERAT_ASAL = rs!Berat 'Berat Asal (g)
            If Not IsNull(rs!beza_berat) Then Frm85_LM_BERAT_KINI = rs!beza_berat
            If IsNumeric(Frm85.TB1) Then Frm85_LM_BERAT_SUSUT = Frm85.TB1 'Susut Berat
            
            Frm85_LM_BERAT_AFTER = Frm85_LM_BERAT_KINI + Frm85_LM_BERAT_KINI_BEFORE - Frm85_LM_BERAT_SUSUT
            
            If Format(Frm85_LM_BERAT_ASAL, "0.00") = Format(Frm85_LM_BERAT_AFTER, "0.00") Then
                rs!beza_berat = Format(Frm85_LM_BERAT_AFTER, "0.00") 'Baki Berat
                rs!StatusItem = 10
                LM_UBAH_STATUS = 1
            Else
                rs!beza_berat = Format(Frm85_LM_BERAT_AFTER, "0.00") 'Baki Berat
            End If
            
            rs!susut_berat = Format(Frm85.TB1, "0.00") 'Susut berat
            rs.Update
            DATA_UPDATE = 1
            
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_UPDATE = 1 Then
        
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & user & "] Susut berat " & Format(Frm85_LM_BERAT_SUSUT, "#,##0.00 g") & ". No siri produk [" & Frm85.L89_Text & "]"
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
            
            If LM_UBAH_STATUS = 1 Then
                MsgBox "Data telah berjaya diubah dan barang ini telah dikembalikan status kepada IN STOCK.", vbInformation, "Info"
            Else
                MsgBox "Data telah berjaya diubah.", vbInformation, "Info"
            End If
        
            GM_NEXT_PREV = 2 '0 : Next , 1 : Previous
            
            Frm85.Pic1.Visible = False
            
            If GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_Potong
                Call Frm85_report_potong_barcode
            Else
                Call Frm85_Header_Report_Potong
                Call Frm85_report_potong_page
            End If
            
        End If
        
    End If
    
End If
End Sub
Private Sub CMD5_Click()
'on error resume next
Dim LM_CURRENT_PAGE As Single
Dim LM_PAGE_QTY As Single

LM_CURRENT_PAGE = 0
LM_PAGE_QTY = 0

If Frm85.L79_Text <> vbNullString And IsNumeric(Frm85.L79_Text) Then
    If Frm85.L81_Text <> vbNullString And IsNumeric(Frm85.L81_Text) Then
        LM_PAGE_QTY = Frm85.L81_Text
        LM_CURRENT_PAGE = Frm85.L79_Text
        
        If LM_CURRENT_PAGE < LM_PAGE_QTY Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            If GM_REPORT_MODE = 0 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_Belian
                Call Frm85_report_belian_page
            ElseIf GM_REPORT_MODE = 1 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_Belian
                Call Frm85_search_berat_page
            ElseIf GM_REPORT_MODE = 8 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_Belian
                Call Frm85_search_invoice_supplier_page
            ElseIf GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_Belian
                Call Frm85_report_belian_barcode
            End If
            
        End If
    End If
End If
End Sub
Private Sub CMD6_Click()
'on error resume next
Dim LM_CURRENT_PAGE As Single
Dim LM_PAGE_QTY As Single

LM_CURRENT_PAGE = 0
LM_PAGE_QTY = 0

If Frm85.L79_Text <> vbNullString And IsNumeric(Frm85.L79_Text) Then
    If Frm85.L81_Text <> vbNullString And IsNumeric(Frm85.L81_Text) Then
        LM_PAGE_QTY = Frm85.L81_Text
        LM_CURRENT_PAGE = Frm85.L79_Text
        
        If LM_CURRENT_PAGE <> 1 And LM_CURRENT_PAGE <> 0 Then
        
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            If GM_REPORT_MODE = 0 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_Belian
                Call Frm85_report_belian_page
            ElseIf GM_REPORT_MODE = 1 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_Belian
                Call Frm85_search_berat_page
            ElseIf GM_REPORT_MODE = 8 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_Belian
                Call Frm85_search_invoice_supplier_page
            ElseIf GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_Belian
                Call Frm85_report_belian_barcode
            End If
            
        End If
    End If
End If
End Sub
Private Sub CMD7_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

If GM_REPORT_MODE = 2 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_Header_Report_Jualan
    Call Frm85_Report_Jualan_page
ElseIf GM_REPORT_MODE = 3 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_Header_Report_Jualan
    Call Frm85_carian_jualan_page
ElseIf GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_Header_Report_Jualan
    Call Frm85_Report_Jualan_barcode
End If
End Sub
Private Sub CMD8_Click()
'on error resume next
Dim LM_CURRENT_PAGE As Single
Dim LM_PAGE_QTY As Single

If Frm85.L79_Text <> vbNullString And IsNumeric(Frm85.L79_Text) Then
    If Frm85.L81_Text <> vbNullString And IsNumeric(Frm85.L81_Text) Then
        LM_PAGE_QTY = Frm85.L81_Text
        LM_CURRENT_PAGE = Frm85.L79_Text
        
        If LM_CURRENT_PAGE < LM_PAGE_QTY Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            If GM_REPORT_MODE = 2 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_Jualan
                Call Frm85_Report_Jualan_page
            ElseIf GM_REPORT_MODE = 3 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_Jualan
                Call Frm85_carian_jualan_page
            ElseIf GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_Jualan
                Call Frm85_Report_Jualan_barcode
            End If
            
        End If
    End If
End If
End Sub
Private Sub CMD9_Click()
'on error resume next
Dim LM_CURRENT_PAGE As Single
Dim LM_PAGE_QTY As Single

LM_CURRENT_PAGE = 0
LM_PAGE_QTY = 0

If Frm85.L79_Text <> vbNullString And IsNumeric(Frm85.L79_Text) Then
    If Frm85.L81_Text <> vbNullString And IsNumeric(Frm85.L81_Text) Then
        LM_PAGE_QTY = Frm85.L81_Text
        LM_CURRENT_PAGE = Frm85.L79_Text
        
        If LM_CURRENT_PAGE <> 1 And LM_CURRENT_PAGE <> 0 Then
        
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            If GM_REPORT_MODE = 5 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_Buyback
                Call Frm85_report_buyback_page
            ElseIf GM_REPORT_MODE = 4 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_Buyback
                Call Frm85_carian_buyback_page
            ElseIf GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                Call Frm85_Header_Report_Buyback
                Call Frm85_report_buyback_barcode
            End If
            
        End If
    End If
End If
End Sub

Private Sub Form_Load()
'on error resume next
Frm85.L10_Text = vbNullString

Frm85.L79_Text = 0
Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
End Sub
Private Sub Frm85_SM_Cetak_Resit_Buyback_Click()
'on error resume next
DATA_FOUND = 0

frm85_LM_No_ID = vbNullString

If IsNumeric(Frm85.LV3.SelectedItem.Index) Then
    
    frm85_LM_No_ID = Frm85.LV3.ListItems(Frm85.LV3.SelectedItem.Index)
    
    If frm85_LM_No_ID <> vbNullString Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where ID='" & frm85_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!bill_No_Trade_In) Then
                G_No_RESIT_JUALAN = rs!bill_No_Trade_In 'No. Resit Buyback
                DATA_FOUND = 1
            Else
                MsgBox "Tiada Data No. Voucher Bagi Item Ini.", vbExclamation, "Error"
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!cawangan) Then
                G_KEDAI = rs!cawangan
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then
            Call Frm84_Resit_Buyback
        End If
    End If
End If
End Sub
Private Sub Frm85_SM_Cetak_Resit_Buyback_gb_Click()
'on error resume next
DATA_FOUND = 0

If Frm85.MSFlexGrid9 <> vbNullString Then

    Frm85_LM_ID = Frm85.MSFlexGrid9.TextMatrix(Frm85.MSFlexGrid9, 2) 'No. ID
    
    If Frm85_LM_ID <> vbNullString Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where ID='" & Frm85_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!bill_No_Trade_In) Then
                G_No_RESIT_JUALAN = rs!bill_No_Trade_In 'No. Resit Buyback
                DATA_FOUND = 1
            Else
                MsgBox "Tiada Data No. Voucher Bagi Item Ini.", vbExclamation, "Error"
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!cawangan) Then
                G_KEDAI = rs!cawangan
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then
            Call Frm84_Resit_Buyback
        End If
    End If
End If
End Sub
Private Sub Frm85_SM_Cetak_Resit_Jualan_Click()
'on error resume next
DATA_FOUND = 0
Frm85_LM_INVOICE_TYPE = 0 'Unlimited , 1 : Limited

If IsNumeric(Frm85.LV1.SelectedItem.Index) Then
    
    frm85_LM_No_ID = Frm85.LV1.ListItems(Frm85.LV1.SelectedItem.Index)
    
    If frm85_LM_No_ID <> vbNullString Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 23_senarai_jualan where ID='" & frm85_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!jenis_jualan) Then
                If rs!jenis_jualan = 0 Then
                    Frm85_LM_JENIS = 0
                ElseIf rs!jenis_jualan = 1 Then
                    Frm85_LM_JENIS = 1
                End If
            Else
                Frm85_LM_JENIS = 0
            End If
            If Not IsNull(rs!cawangan) Then G_KEDAI = rs!cawangan
            If Not IsNull(rs!no_resit) Then
                G_No_RESIT_JUALAN = rs!no_resit 'No. Resit Jualan
                DATA_FOUND = 1
            Else
                MsgBox "Tiada Data No. Invoice Bagi Item Ini.", vbExclamation, "Error"
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 22_jualan where no_resit='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!invoice_type) Then
                If rs!invoice_type <> 0 Then Frm85_LM_INVOICE_TYPE = 1 'Unlimited , 1 : Limited
                If rs!invoice_type = 0 Then Frm85_LM_INVOICE_TYPE = 0 'Unlimited , 1 : Limited
            Else
                Frm85_LM_INVOICE_TYPE = 0 'Unlimited , 1 : Limited
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        
        G_PREVIEW = 1
        
        If DATA_FOUND = 1 Then
            If Frm85_LM_JENIS = 0 Then
                If Frm85_LM_INVOICE_TYPE = 0 Then 'Unlimited , 1 : Limited
                    If G_INVOICE_TYPE = 0 Then '0 : Invoice Dari Sistem , 2 : Invoice Pre-printed
                        Call Frm84_Resit_Jualan
                    ElseIf G_INVOICE_TYPE = 1 Then '0 : Invoice Dari Sistem , 2 : Invoice Pre-printed
                        Call cetak_invoice
                    End If
                ElseIf Frm85_LM_INVOICE_TYPE = 1 Then 'Unlimited , 1 : Limited
                    Call Frm84_cetak_invoice_rms
                End If
            ElseIf Frm85_LM_JENIS = 1 Then
                Call Frm115_cetak_gdn
            End If
        End If
    End If
End If
End Sub
Private Sub Frm85_SM_Edit_Data_Belian_Click()
'on error resume next
DATA_FOUND = 0

frm85_LM_No_ID = vbNullString

If IsNumeric(Frm85.LV2.SelectedItem.Index) Then
    
    frm85_LM_No_ID = Frm85.LV2.ListItems(Frm85.LV2.SelectedItem.Index)
    
    If frm85_LM_No_ID <> vbNullString Then
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where ID='" & frm85_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then

            If Not IsNull(rs!StatusItem) Then
                
                If rs!StatusItem = "10" Then
                    
                    G_ID = rs!ID
                    DATA_FOUND = 1
                    
                Else
                
                    MsgBox "Status barang ini telah berubah dan anda tidak dibenarkan untuk edit data barang ini." & vbCrLf & _
                            vbNullString & vbCrLf & _
                            "", vbExclamation, "Info"
                        
                End If
            
            Else
                
                MsgBox "Tiada maklumat status bagi barang ini. Sila keluar dari menu ini dan cuba lagi.", vbExclamation, "Info"
            
            End If
            
        Else
            
            MsgBox "Tiada maklumat bagi item ini.", vbExclamation, "Info"
            
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then
        
            Frm83.CB9 = 1
            Frm83.CB10 = 0
            
            Call Frm83_Initial_Setting
            Call Frm83_initial_setting2
            Call Frm26_initial
            Call Frm27_initial
            Call Frm28_initial
            
            Frm83.L69_Text = -1 'Titik Pencarian Data
            Frm83.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
            Frm83.L67_Text = 0 'Paparan Page ke-xxx
            Frm83.L68_Text = 0
            
            GM_NEXT_PREV = 0
        
            Frm83.L41_Text = 2 '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
            Call frm83_recall_data_penerimaan_stok
            Frm83.Frame8.Visible = True
            Frm83.Frame1.Visible = True
            
            Frm83.CMD24.Enabled = False
            Frm83.CMD25.Enabled = False
            
            Frm83.L100_Text.Visible = False
            Frm83.L101_Text.Visible = False
            Frm83.L102_Text.Visible = False
            Frm83.TB40.Visible = False
            Frm83.TB41.Visible = False
            Frm83.TB42.Visible = False
            
            Frm83.CMD10.Visible = True
            Frm83.CMD11.Visible = True
            Frm83.CMD20.Visible = False
            Frm83.CMD21.Visible = False
            
            Frm83.CMD1.Visible = False
            Frm83.CMD6.Visible = False
            Frm83.CMD7.Visible = False
            Frm83.CMD12.Visible = False
            Frm83.CMD13.Visible = False
            Frm83.CMD14.Visible = False
            Frm83.CMD2.Visible = False
            Frm83.CMD5.Visible = False
            Frm83.CMD10.Visible = False
            Frm83.CMD11.Visible = False
            
            Frm83.Frame1.Left = 120
            Frm83.Frame1.Top = 120
    
            If Frm83.TB28 = vbNullString Then
            
                Frm83.CB2 = 1
                Frm83.CB2.Enabled = False
                Frm83.CB3.Enabled = False
                Frm83.CB11.Enabled = False
                Frm83.CB12.Enabled = False
                
            Else
            
                Frm83.CB2.Enabled = True
                Frm83.CB3.Enabled = True
                Frm83.CB11.Enabled = True
                Frm83.CB12.Enabled = True
                
            End If
        End If
    End If
End If
End Sub
Private Sub Frm85_SM_Edit_Data_Belian_gb_Click()
'on error resume next
DATA_FOUND = 0

If Frm85.MSFlexGrid8 <> vbNullString Then
    Frm85_LM_ID = Frm85.MSFlexGrid8.TextMatrix(Frm85.MSFlexGrid8, 2) 'No. ID
    
    If Frm85_LM_ID <> vbNullString Then
    
        Frm83.CB9 = 0
        Frm83.CB10 = 1
        
        Call Frm83_Initial_Setting
        Call Frm83_initial_setting2
        Call Frm26_initial
        Call Frm27_initial
        Call Frm28_initial
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where ID='" & Frm85_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!NoRujukanSistem) Then
            
                GLOBAL_DISABLE = 1
                Frm83.L9_Text = rs!NoRujukanSistem 'No. Rujukan Belian

                On Error GoTo Err_A:
                If Not IsNull(rs!nama_Supplier) Then
                    Frm83_LM_Supplier = rs!nama_Supplier 'Nama Supplier
                    Frm83.CBB1 = Frm83_LM_Supplier 'Nama Supplier
                End If
                
Restore_A:
                If Not IsNull(rs!Kod_Supplier) Then Frm83.TB1 = rs!Kod_Supplier 'Kod Supplier
                
                GLOBAL_DISABLE = 0
                DATA_FOUND = 1
            Else
                MsgBox "Tiada Data Rujukan Belian.", vbExclamation, "Error"
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then
            Frm83.CB9 = 0
            Frm83.CB10 = 0
            Frm83.L41_Text = 2 '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
            Call Frm85_Recall_Data_Belian
            
            If Frm83.TB28 = vbNullString Then
                Frm83.CB2 = 1
                Frm83.CB2.Enabled = False
                Frm83.CB3.Enabled = False
                Frm83.CB11.Enabled = False
                Frm83.CB12.Enabled = False
            Else
                Frm83.CB2.Enabled = True
                Frm83.CB3.Enabled = True
                Frm83.CB11.Enabled = True
                Frm83.CB12.Enabled = True
            End If
        End If
    End If
End If

Exit Sub
Err_A:
Frm83.CBB1.AddItem Frm83_LM_Supplier
Frm83.CBB1 = Frm83_LM_Supplier
Resume Restore_A:
End Sub
Private Sub Frm85_SM_Edit_Data_Buyback_Click()
'on error resume next
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset

DATA_FOUND = 0
Frm85_LM_JENIS_BELIAN = 0 '1 : Belian secara trade in barang , 2 : Belian secara trade in (voucher)
Frm85_LM_LOCKED = 0 '0 : Dibenarkan utk edit / padam , 1 : Tidak dibenarkan untuk edit / padam

frm85_LM_No_ID = vbNullString

If IsNumeric(Frm85.LV3.SelectedItem.Index) Then
    
    frm85_LM_No_ID = Frm85.LV3.ListItems(Frm85.LV3.SelectedItem.Index)
    
    If frm85_LM_No_ID <> vbNullString Then
    
'### Periksa status invoice trade in ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where ID='" & frm85_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            If rs!jenis_trade_in = 1 Then
            
                Frm85_LM_JENIS_BELIAN = 2 '1 : Belian secara trade in barang , 2 : Belian secara trade in (voucher)
                
            
            ElseIf rs!jenis_trade_in = 0 Then
                
                Frm85_LM_JENIS_BELIAN = 1 '1 : Belian secara trade in barang , 2 : Belian secara trade in (voucher)
            
            End If
            
            Set rs1 = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs1.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & rs!bill_No_Trade_In & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs1.EOF Then
                
                If Not IsNull(rs1!trade_in_status) Then
                
                    If rs1!trade_in_status = 1 Then
                    
                        Set rs2 = New ADODB.Recordset
                        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                        rs2.Open "select * from 22_jualan where no_resit_trade_in='" & rs!bill_No_Trade_In & "'", cn, adOpenKeyset, adLockOptimistic
                        
                        If Not rs2.EOF Then
                            If Not IsNull(rs2!no_resit) Then
                                Frm85_LM_No_INVOICE = rs2!no_resit 'No. invoice jualan
                                Frm85_LM_LOCKED = 1 '0 : Dibenarkan utk edit / padam , 1 : Tidak dibenarkan untuk edit / padam
                            End If
                        End If
                        
                        rs2.Close
                        Set rs2 = Nothing
                        
                    End If
                
                End If
            End If
            
            rs1.Close
            Set rs1 = Nothing
        
        End If
        
        rs.Close
        Set rs = Nothing
        
        If Frm85_LM_LOCKED = 1 Then '0 : Dibenarkan utk edit / padam , 1 : Tidak dibenarkan untuk edit / padam
            
            If Frm85_LM_JENIS_BELIAN = 1 Or Frm85_LM_JENIS_BELIAN = 2 Then '1 : Belian secara trade in barang , 2 : Belian secara trade in (voucher)
                
                MsgBox "Data bagi barang kemas ini tidak dibenarkan untuk diedit atau dipadamkan." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Sebab :" & vbCrLf & _
                        "Voucher atau data bagi data ini telah digunakan untuk jualan barang kepada pelanggan." & vbCrLf & _
                        "Untuk edit data ini perlu rujuk kepada invoice jualan [" & Frm85_LM_No_INVOICE & "] untuk edit atau padam data.", vbExclamation, "Info"
                
                Exit Sub
                
            End If
    
        End If
'### Periksa status invoice trade in ### - End
    
    
        Frm83.CB9 = 1
        Frm83.CB10 = 0
        
        Call Frm83_Initial_Setting '!! Hati-hati dengan tempat letakkan command ini!!
        Call Frm83_initial_setting2
        Call Frm26_initial
        Call Frm27_initial
        Call Frm28_initial
        
        Frm83.L69_Text = -1 'Titik Pencarian Data
        Frm83.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
        Frm83.L67_Text = 0 'Paparan Page ke-xxx
        Frm83.L68_Text = 0
        
        GM_NEXT_PREV = 0
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where ID='" & frm85_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!NoRujukanSistem) Then
            
                GLOBAL_DISABLE = 1
                
                Frm83.L9_Text = rs!NoRujukanSistem 'No. Rujukan Belian
                
                On Error GoTo Err_A:
                If Not IsNull(rs!nama_Supplier) Then
                    Frm83_LM_Supplier = rs!nama_Supplier 'Nama Supplier
                    Frm83.CBB1 = Frm83_LM_Supplier 'Nama Supplier
                End If
                
Restore_A:
                If Not IsNull(rs!Kod_Supplier) Then Frm83.TB1 = rs!Kod_Supplier 'Kod Supplier
                
                DATA_FOUND = 1
                GLOBAL_DISABLE = 0
                
            Else
                MsgBox "Tiada Data Rujukan Belian.", vbExclamation, "Error"
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then
        
            Frm83.CB14 = 0
            Frm83.CB15 = 0
            Frm83.TB35 = vbNullString
            
            Frm83.CB14.Enabled = False
            Frm83.CB15.Enabled = False
            
            Frm83.TB35.BackColor = &H8000000A
            Frm83.TB35.Locked = True
            
            Frm83.ListView1.ListItems.Clear
            
            With Frm83.ListView1
                Set .SmallIcons = Frm83.ImageList1
                Set .Icons = Frm83.ImageList1
        
                .ListItems.Add , "Data Item", "Data Item", 1
                .ListItems.Add , "Senarai Item", "Senarai Item", 2
                
            End With

            Frm83.CB9 = 0
            Frm83.CB10 = 0
            Frm83.L41_Text = 0 '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
            'GLOBAL_DISABLE = 1
            Call Frm85_Recall_Data_Belian
            'GLOBAL_DISABLE = 0
                    
            Frm83.CMD24.Enabled = True
            Frm83.CMD25.Enabled = True

            Frm83.L100_Text.Visible = True
            Frm83.L101_Text.Visible = True
            Frm83.L102_Text.Visible = True
            Frm83.TB40.Visible = True
            Frm83.TB41.Visible = True
            Frm83.TB42.Visible = True
            'Frm83.CMD5.Visible = True
            'Frm83.CMD2.Visible = True
            'Frm83.CMD22.Visible = False
            'Frm83.CMD23.Visible = False

            If G_PRINTER_TI_MODE = 0 Then
                Frm83.CB13 = 0
            ElseIf G_PRINTER_TI_MODE = 1 Then
                Frm83.CB13 = 1
            End If

            Frm83.Label40.Visible = True
            Frm83.L10_Text.Visible = True
    
            Frm83.Frame1.Left = 1680
            Frm83.Frame1.Top = 120
            
            Frm83.Frame9.Left = 1680
            Frm83.Frame9.Top = 120
            
            Frm83.ListView1.Left = 120
            Frm83.ListView1.Top = 120
            
            Frm83.Frame1.Visible = True
            Frm83.ListView1.Visible = True
            
            Frm83.CBB1.Enabled = False
            Frm83.CBB1.BackColor = &H8000000A
            
            Frm83.Frame1.Visible = False
            Frm83.Frame9.Visible = True
            
            If Frm83.TB28 = vbNullString Then
                Frm83.CB2 = 1
                Frm83.CB2.Enabled = False
                Frm83.CB3.Enabled = False
                Frm83.CB11.Enabled = False
                Frm83.CB12.Enabled = False
            Else
                Frm83.CB2.Enabled = True
                Frm83.CB3.Enabled = True
                Frm83.CB11.Enabled = True
                Frm83.CB12.Enabled = True
            End If
        End If
    End If
End If

Exit Sub
Err_A:
Frm83.CBB1.AddItem Frm83_LM_Supplier
Frm83.CBB1 = Frm83_LM_Supplier
Resume Restore_A:
End Sub
Private Sub Frm85_SM_Edit_Data_buyback_gb_Click()
'on error resume next
DATA_FOUND = 0
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset

Frm85_LM_JENIS_BELIAN = 0 '1 : Belian secara trade in barang , 2 : Belian secara trade in (voucher)
Frm85_LM_LOCKED = 0 '0 : Dibenarkan utk edit / padam , 1 : Tidak dibenarkan untuk edit / padam

If Frm85.MSFlexGrid9 <> vbNullString Then
    Frm85_LM_ID = Frm85.MSFlexGrid9.TextMatrix(Frm85.MSFlexGrid9, 2) 'No. ID
    
    If Frm85_LM_ID <> vbNullString Then
    
'### Periksa status invoice trade in ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where ID='" & Frm85_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            If rs!jenis_trade_in = 1 Then
            
                Frm85_LM_JENIS_BELIAN = 2 '1 : Belian secara trade in barang , 2 : Belian secara trade in (voucher)
                
            
            ElseIf rs!jenis_trade_in = 0 Then
                
                Frm85_LM_JENIS_BELIAN = 1 '1 : Belian secara trade in barang , 2 : Belian secara trade in (voucher)
            
            End If
            
            Set rs1 = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs1.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & rs!bill_No_Trade_In & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs1.EOF Then
                
                If Not IsNull(rs1!trade_in_status) Then
                
                    If rs1!trade_in_status = 1 Then
                    
                        Set rs2 = New ADODB.Recordset
                        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                        rs2.Open "select * from 22_jualan where no_resit_trade_in='" & rs!bill_No_Trade_In & "'", cn, adOpenKeyset, adLockOptimistic
                        
                        If Not rs2.EOF Then
                            If Not IsNull(rs2!no_resit) Then
                                Frm85_LM_No_INVOICE = rs2!no_resit 'No. invoice jualan
                                Frm85_LM_LOCKED = 1 '0 : Dibenarkan utk edit / padam , 1 : Tidak dibenarkan untuk edit / padam
                            End If
                        End If
                        
                        rs2.Close
                        Set rs2 = Nothing
                        
                    End If
                
                End If
            End If
            
            rs1.Close
            Set rs1 = Nothing
        
        End If
        
        rs.Close
        Set rs = Nothing
        
        If Frm85_LM_LOCKED = 1 Then '0 : Dibenarkan utk edit / padam , 1 : Tidak dibenarkan untuk edit / padam
            
            If Frm85_LM_JENIS_BELIAN = 1 Or Frm85_LM_JENIS_BELIAN = 2 Then '1 : Belian secara trade in barang , 2 : Belian secara trade in (voucher)
                
                MsgBox "Data bagi barang kemas ini tidak dibenarkan untuk diedit atau dipadamkan." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Sebab :" & vbCrLf & _
                        "Voucher atau data bagi data ini telah digunakan untuk jualan barang kepada pelanggan." & vbCrLf & _
                        "Untuk edit data ini perlu rujuk kepada invoice jualan [" & Frm85_LM_No_INVOICE & "] untuk edit atau padam data.", vbExclamation, "Info"
                
                Exit Sub
                
            End If
    
        End If
'### Periksa status invoice trade in ### - End
    
        Frm83.CB9 = 0
        Frm83.CB10 = 1
        
        Call Frm83_Initial_Setting
        Call Frm83_initial_setting2
        Call Frm26_initial
        Call Frm27_initial
        Call Frm28_initial
        
        Frm83.L69_Text = -1 'Titik Pencarian Data
        Frm83.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
        Frm83.L67_Text = 0 'Paparan Page ke-xxx
        Frm83.L68_Text = 0
        
        GM_NEXT_PREV = 0
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where ID='" & Frm85_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!NoRujukanSistem) Then
            
                GLOBAL_DISABLE = 1
                Frm83.L9_Text = rs!NoRujukanSistem 'No. Rujukan Belian

                On Error GoTo Err_A:
                If Not IsNull(rs!nama_Supplier) Then
                    Frm83_LM_Supplier = rs!nama_Supplier 'Nama Supplier
                    Frm83.CBB1 = Frm83_LM_Supplier 'Nama Supplier
                End If
                
Restore_A:
                If Not IsNull(rs!Kod_Supplier) Then Frm83.TB1 = rs!Kod_Supplier 'Kod Supplier
                
                GLOBAL_DISABLE = 0
                DATA_FOUND = 1
            Else
                MsgBox "Tiada Data Rujukan Belian.", vbExclamation, "Error"
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then
            Frm83.CB9 = 0
            Frm83.CB10 = 0
            Frm83.L41_Text = 0 '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
            Call Frm85_Recall_Data_Belian
            
            If Frm83.TB28 = vbNullString Then
                Frm83.CB2 = 1
                Frm83.CB2.Enabled = False
                Frm83.CB3.Enabled = False
                Frm83.CB11.Enabled = False
                Frm83.CB12.Enabled = False
            Else
                Frm83.CB2.Enabled = True
                Frm83.CB3.Enabled = True
                Frm83.CB11.Enabled = True
                Frm83.CB12.Enabled = True
            End If
        End If
    End If
End If

Exit Sub
Err_A:
Frm83.CBB1.AddItem Frm83_LM_Supplier
Frm83.CBB1 = Frm83_LM_Supplier
Resume Restore_A:
End Sub
Private Sub Frm85_SM_Edit_Data_Jualan_Click()
'on error resume next
DATA_FOUND = 0

frm85_LM_No_ID = vbNullString

If IsNumeric(Frm85.LV1.SelectedItem.Index) Then
    
    frm85_LM_No_ID = Frm85.LV1.ListItems(Frm85.LV1.SelectedItem.Index)
    
    If frm85_LM_No_ID <> vbNullString Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 23_senarai_jualan where ID='" & frm85_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!jenis_jualan) Then
                If rs!jenis_jualan = 0 Then
                    Frm85_LM_JENIS = 0
                ElseIf rs!jenis_jualan = 1 Then
                    Frm85_LM_JENIS = 1
                End If
            Else
                Frm85_LM_JENIS = 0
            End If
            If Not IsNull(rs!no_resit) Then
                Frm85_LM_No_INVOICE = rs!no_resit 'No. Resit Jualan
                DATA_FOUND = 1
            Else
                MsgBox "Tiada Data No. Resit Bagi Item Ini.", vbExclamation, "Error"
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_FOUND = 1 Then
        
            If Frm85_LM_JENIS = 0 Then
            
                Call Frm84_Reset
                Call Frm84_Load_Form '!! Hati-hati dengan tempat letakkan command ini!!
                Call frm130_initial_setting
                Unload Frm26
                Unload Frm27
                Unload Frm28
                'Call Frm26_initial
                'Call Frm27_initial
                'Call Frm28_initial
                MDI_frm1.L5_Text = 4
                
                '### Periksa jenis invoice jualan samada RASMI atau TIDAK RASMI ###
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 22_jualan where no_resit='" & Frm85_LM_No_INVOICE & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    
                    If Not IsNull(rs!ID) Then LM_ID = rs!ID
                    
                    If Not IsNull(rs!bil_rasmi) Then
                        If rs!bil_rasmi = 0 Then
                        
                            'Frm84.L66_Text = Frm85_LM_No_INVOICE
                            Frm84.CB13 = 1
                            
                        ElseIf rs!bil_rasmi = 1 Then
                        
                            'Frm84.L3_Text = Frm85_LM_No_INVOICE
                            Frm84.CB13 = 0
                            
                        End If
                    End If
                    
                End If
                
                rs.Close
                Set rs = Nothing
                
                Frm84.CB13.Enabled = False
                
                'GLOBAL_DISABLE = 1
                Frm84.L3_Text = Frm85_LM_No_INVOICE
                Call Frm85_Recall_Data_Jualan
                Call Frm84_background_color
                Call frm_kiraan_harga_selepas_ti
                'GLOBAL_DISABLE = 0
                
            ElseIf Frm85_LM_JENIS = 1 Then
                
                GoTo skip_a:
                
                GLOBAL_DISABLE = 0
                Frm102.TB1 = vbNullString
                
                'Call Frm28_initial
                Unload Frm28
                Call frm102_reset_1
                Call frm102_reset_2
                Call frm102_reset_3
                Call frm102_reset_main
                
                Frm102.L26_Text.BackStyle = 0
                Frm102.L27_Text.BackStyle = 0
                
                Frm102.DTPicker1 = DateTime.Date$
                
                Frm102.L32_Text = 1 '0 : Data Baru , 1 : Edit Data
                
                Frm102.CMD8.Visible = False
                Frm102.CMD9.Visible = False
                Frm102.CMD10.Visible = True
                Frm102.CMD11.Visible = True
                
                Frm102.L23_Text = Frm85_LM_No_INVOICE
                
                MDI_frm1.L5_Text = 6
                Call Frm102_recall_edit_jualan
                Call Frm102_background_color
                
skip_a:
          
                GLOBAL_DISABLE = 0
                Frm115.TB1 = vbNullString
                
                Call Frm115_reset_1
                Call Frm115_reset_2
                Call Frm115_reset_3
                Call Frm115_reset_main
                Call Frm115_reset_main2
                
                Call frm115_initial_setting_stok
                Call frm115_reset_gdn_list
                
                Frm115.DTPicker1 = DateTime.Date$
                
                Frm115.L32_Text = 1 '0 : Data Baru , 1 : Edit Data
                Frm115.L54_Text = LM_ID
                
                Frm115.CMD8.Visible = False
                Frm115.CMD9.Visible = False
                Frm115.CMD10.Visible = True
                Frm115.CMD11.Visible = True
                
                Frm115.L23_Text = Frm85_LM_No_INVOICE
                
                MDI_frm1.L5_Text = 16
                Call Frm115_recall_edit_jualan
                Call Frm115_background_color
                
                Frm115.L71_Text = "0"
                
                Frm115.TB1.SetFocus
                
            End If
            
        End If
    End If
End If
End Sub


Private Sub frm85_sm_email_jualan_Click()
'on error resume next
DATA_FOUND = 0
Frm85_LM_INVOICE_TYPE = 0 'Unlimited , 1 : Limited

If IsNumeric(Frm85.LV1.SelectedItem.Index) Then
    
    frm85_LM_No_ID = Frm85.LV1.ListItems(Frm85.LV1.SelectedItem.Index)
    
    If frm85_LM_No_ID <> vbNullString Then
    
        Note = "Hantar email berkenaan invoice jualan ini kepada pihak pengurusan?" & vbCrLf & _
                "" & vbCrLf & _
                "Teruskan ?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 23_senarai_jualan where ID='" & frm85_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!no_resit) Then
                    G_No_RESIT_JUALAN = rs!no_resit 'No. Resit Jualan
                    DATA_FOUND = 1
                Else
                    MsgBox "Tiada Data No. Invoice Bagi Item Ini.", vbExclamation, "Error"
                End If
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_FOUND = 1 Then
                
                LM_NOW = Now
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 88_sales_notfication", cn, adOpenKeyset, adLockOptimistic
                
                rs.AddNew
                rs!no_invoice_asal = G_No_RESIT_JUALAN
                rs!jenis = 3
                rs!jenis_report = 0 '0 : Jualan , 1 : Trade In
                rs!write_timestamp = LM_NOW
                rs!Status = 0
                rs!terminal = G_TERMINAL
                rs.Update
                
                rs.Close
                Set rs = Nothing
                
                Shell "cmd.exe /c " & G_SPKE_NE_PATH
                
                MsgBox "Sistem telah menghantar email berkenaan invoice ini kepada pihak pengurusan.", vbInformation, "Info"
                
            End If
        
        End If
        
    End If
End If
End Sub
Private Sub Frm85_SM_Export1_Overall_1_Click()
'on error resume next
'REPORT KESELURUHAN BELIAN BARANG KEMAS - EXCEL
'REPORT KESELURUHAN BELIAN BARANG KEMAS (IKUT No. INVOICE) - EXCEL

If GM_REPORT_MODE = 0 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_overall_beli
ElseIf GM_REPORT_MODE = 1 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_overall_beli_berat
ElseIf GM_REPORT_MODE = 8 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_overall_beli_invoice_supplier
ElseIf GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_overall_beli_no_siri
End If
End Sub
Private Sub Frm85_SM_Export1_Overall_10_Click()
'on error resume next
Call Frm85_excel_trade_in
End Sub

Private Sub Frm85_SM_Export1_Overall_11_Click()
'on error resume next
Call frm85_excel_susut_nilai
End Sub

Private Sub Frm85_SM_Export1_Overall_2_Click()
'on error resume next
'REPORT KESELURUHAN JUALAN - EXCEL
'REPORT KESELURUHAN JUALAN (No. Siri Produk) - EXCEL
'REPORT KESELURUHAN JUALAN (No. Invoice Jualan) - EXCEL

If GM_REPORT_MODE = 2 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_overall_jual
ElseIf GM_REPORT_MODE = 3 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_overall_jual_invoice
ElseIf GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_overall_jual_no_siri
End If
End Sub
Private Sub Frm85_SM_Export1_Overall_3_Click()
'On Error Resume Next
'REPORT KESELURUHAN TRADE IN BARANG KEMAS - EXCEL
'REPORT KESELURUHAN TRADE IN IKUT INVOICE - EXCEL
'REPORT KESELURUHAN TRADE IN IKUT NO SIRI - EXCEL

If GM_REPORT_MODE = 5 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_bk_trade_overall
ElseIf GM_REPORT_MODE = 4 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_bk_trade_invoice_overall
ElseIf GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_bk_trade_siri_overall
End If
End Sub
Private Sub Frm85_SM_Export1_Overall_4_Click()
'on error resume next
'REPORT STOK KESELURUHAN - EXCEL
'REPORT STOK IKUT NO SIRI - EXCEL
If GM_REPORT_MODE = 10 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_stok_overall
ElseIf GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_stok_siri
End If
End Sub
Private Sub Frm85_SM_Export1_Overall_5_Click()
'on error resume next
'REPORT KESELURUHAN BELIAN GOLD BAR - EXCEL
'REPORT KESELURUHAN BELIAN GOLD BAR IKUT NO SIRI - EXCEL
If GM_REPORT_MODE = 6 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_belian_gb_overall
ElseIf GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_belian_gb_siri
End If
End Sub
Private Sub Frm85_SM_Export1_Overall_6_Click()
'On Error Resume Next
'REPORT KESELURUHAN TRADE IN GOLD BAR - EXCEL
'REPORT KESELURUHAN TRADE IN GOLD BAR IKUT NO SIRI - EXCEL
If GM_REPORT_MODE = 7 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_buyback_gb_overall
ElseIf GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_buyback_gb_siri
End If
End Sub
Private Sub Frm85_SM_Export1_Overall_7_Click()
'On Error Resume Next
'REPORT KESELURUHAN POTONG - EXCEL
'REPORT KESELURUHAN POTONG IKUT NO SIRI - EXCEL
If GM_REPORT_MODE = 11 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_potong_overall
ElseIf GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_potong_siri
End If
End Sub
Private Sub Frm85_SM_Export1_Overall_8_Click()
'On Error Resume Next
'REPORT KESELURUHAN ANSURAN - EXCEL
'REPORT KESELURUHAN ANSURAN IKUT NO SIRI - EXCEL

If GM_REPORT_MODE = 12 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_ansuran_overall
ElseIf GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_ansuran_siri
End If
End Sub
Private Sub Frm85_SM_Export1_Overall_9_Click()
'On Error Resume Next
'REPORT KESELURUHAN TEMPAHAN - EXCEL
'REPORT KESELURUHAN TEMPAHAN IKUT NO SIRI - EXCEL

If GM_REPORT_MODE = 13 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_tempahan_overall
ElseIf GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_tempahan_siri
End If
End Sub
Private Sub Frm85_SM_Export6_Click()
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
        .Columns("B").ColumnWidth = 20 'Tarikh Belian
        .Columns("C").ColumnWidth = 20 'No. Siri Produk
        .Columns("D").ColumnWidth = 20 'Purity
        .Columns("E").ColumnWidth = 40 'Kategori Produk
        .Columns("F").ColumnWidth = 40 'Supplier
        .Columns("G").ColumnWidth = 20 'Berat (g)
        .Columns("H").ColumnWidth = 20 'Rate Penerimaan (RM/g)
        .Columns("I").ColumnWidth = 20 'Upah (RM)
        .Columns("J").ColumnWidth = 20 'Spread (%)
        .Columns("K").ColumnWidth = 20 'Harga Selepas Spread (RM)
        .Columns("L").ColumnWidth = 20 'Adjustment (RM)
        .Columns("M").ColumnWidth = 20 'Harga Belian (RM)
        .Columns("N").ColumnWidth = 20 'Dulang
        .Columns("O").ColumnWidth = 20 'Panjang
        .Columns("P").ColumnWidth = 20 'Lebar
        .Columns("Q").ColumnWidth = 20 'Saiz
    
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
        
        .Cells(7, 1) = Frm85.L71_Text 'Header Report
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Belian"
        .Cells(8, 3) = "No. Siri Produk"
        .Cells(8, 4) = "Purity"
        .Cells(8, 5) = "Kategori Produk"
        .Cells(8, 6) = "Supplier"
        .Cells(8, 7) = "Berat (g)"
        .Cells(8, 8) = "Rate Penerimaan (RM/g)"
        .Cells(8, 9) = "Upah (RM)"
        .Cells(8, 10) = "Spread (%)"
        .Cells(8, 11) = "Harga Selepas Spread (RM)"
        .Cells(8, 12) = "Adjustment (RM)"
        .Cells(8, 13) = "Harga Belian Termasuk GST (RM)"
        .Cells(8, 14) = "Dulang"
        .Cells(8, 15) = "Panjang"
        .Cells(8, 16) = "Lebar"
        .Cells(8, 17) = "Saiz"
        
        For i = 1 To 17
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
        
        Y = 0
        For x = 1 To Frm85.MSFlexGrid8.Rows - 1
            Y = Y + 1
            .Cells(8 + Y, 1) = Y 'No.
            .Cells(8 + Y, 1).HorizontalAlignment = xlCenter
            .Cells(8 + Y, 2) = "'" & Frm85.MSFlexGrid8.TextMatrix(x, 3) 'Tarikh Belian
            .Cells(8 + Y, 2).HorizontalAlignment = xlCenter
            
            .Cells(8 + Y, 3) = Frm85.MSFlexGrid8.TextMatrix(x, 4) 'No. Siri Produk
            .Cells(8 + Y, 3).HorizontalAlignment = xlCenter
            
            .Cells(8 + Y, 4) = Frm85.MSFlexGrid8.TextMatrix(x, 5) 'Purity
            .Cells(8 + Y, 4).HorizontalAlignment = xlCenter
            
            .Cells(8 + Y, 5) = Frm85.MSFlexGrid8.TextMatrix(x, 6) 'Kategori Produk
            
            .Cells(8 + Y, 6) = Frm85.MSFlexGrid8.TextMatrix(x, 7) 'Supplier
            
            .Cells(8 + Y, 7).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 7) = Frm85.MSFlexGrid8.TextMatrix(x, 8) 'Berat (g)
            .Cells(8 + Y, 7).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 8).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 8) = Frm85.MSFlexGrid8.TextMatrix(x, 9) 'Rate Penerimaan (RM/g)
            .Cells(8 + Y, 8).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 9).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 9) = Frm85.MSFlexGrid8.TextMatrix(x, 10) 'Upah (RM)
            .Cells(8 + Y, 9).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 10).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 10) = Frm85.MSFlexGrid8.TextMatrix(x, 11) 'Spread (%)
            .Cells(8 + Y, 10).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 11).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 11) = Frm85.MSFlexGrid8.TextMatrix(x, 12) 'Harga Selepas Spread (RM)
            .Cells(8 + Y, 11).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 12).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 12) = Frm85.MSFlexGrid8.TextMatrix(x, 13) 'Adjustment (RM)
            .Cells(8 + Y, 12).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 13).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 13) = Frm85.MSFlexGrid8.TextMatrix(x, 14) 'Harga Belian (RM)
            .Cells(8 + Y, 13).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 14) = Frm85.MSFlexGrid8.TextMatrix(x, 15) 'Dulang
            .Cells(8 + Y, 14).HorizontalAlignment = xlCenter
            
            .Cells(8 + Y, 15) = Frm85.MSFlexGrid8.TextMatrix(x, 16) 'Panjang
            .Cells(8 + Y, 15).HorizontalAlignment = xlCenter
            
            .Cells(8 + Y, 16) = Frm85.MSFlexGrid8.TextMatrix(x, 17) 'Lebar
            .Cells(8 + Y, 16).HorizontalAlignment = xlCenter
            
            .Cells(8 + Y, 17) = Frm85.MSFlexGrid8.TextMatrix(x, 18) 'Saiz
            .Cells(8 + Y, 17).HorizontalAlignment = xlCenter

            For Col = 1 To 17
                .Cells(8 + Y, Col).Borders.LineStyle = xlContinuous
            Next Col
        Next x
    
        Y = Y + 2
        .Cells(8 + Y, 1) = "Bilangan : " & Frm85.L65_Text 'Total Barang
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah Berat (g) : " & Frm85.L66_Text 'Total Berat
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah Modal : RM " & Frm85.L67_Text 'Total Harga Belian
        
        Y = Y + 4
        .Cells(8 + Y, 1).Font.Bold = True
        .Cells(8 + Y, 1) = "Report Generated By Sankyu System [ +6010 - 900 4788 , sankyusystem@gmail.com ]" 'Watermark Sankyu System
    End With
    
    ' This makes Excel visible
    xlObject.Visible = True
    xlObject.EnableEvents = True
End If
End Sub
Private Sub Frm85_SM_Export7_Click()
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
        .Columns("B").ColumnWidth = 20 'Tarikh Belian
        .Columns("C").ColumnWidth = 20 'No. Siri Produk
        .Columns("D").ColumnWidth = 20 'Purity
        .Columns("E").ColumnWidth = 40 'Kategori Produk
        .Columns("F").ColumnWidth = 40 'Supplier
        .Columns("G").ColumnWidth = 20 'Berat (g)
        .Columns("H").ColumnWidth = 20 'Rate Penerimaan (RM/g)
        .Columns("I").ColumnWidth = 20 'Upah (RM)
        .Columns("J").ColumnWidth = 20 'Spread (%)
        .Columns("K").ColumnWidth = 20 'Harga Selepas Spread (RM)
        .Columns("L").ColumnWidth = 20 'Adjustment (RM)
        .Columns("M").ColumnWidth = 20 'Harga Belian (RM)
        .Columns("N").ColumnWidth = 20 'Dulang
        .Columns("O").ColumnWidth = 20 'Panjang
        .Columns("P").ColumnWidth = 20 'Lebar
        .Columns("Q").ColumnWidth = 20 'Saiz
    
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
        
        .Cells(7, 1) = Frm85.L72_Text 'Header Report
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Belian"
        .Cells(8, 3) = "No. Siri Produk"
        .Cells(8, 4) = "Purity"
        .Cells(8, 5) = "Kategori Produk"
        .Cells(8, 6) = "Supplier"
        .Cells(8, 7) = "Berat (g)"
        .Cells(8, 8) = "Rate Penerimaan (RM/g)"
        .Cells(8, 9) = "Upah (RM)"
        .Cells(8, 10) = "Spread (%)"
        .Cells(8, 11) = "Harga Selepas Spread (RM)"
        .Cells(8, 12) = "Adjustment (RM)"
        .Cells(8, 13) = "Harga Belian (RM)"
        .Cells(8, 14) = "Dulang"
        .Cells(8, 15) = "Panjang"
        .Cells(8, 16) = "Lebar"
        .Cells(8, 17) = "Saiz"
        
        For i = 1 To 17
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
        
        Y = 0
        For x = 1 To Frm85.MSFlexGrid9.Rows - 1
            Y = Y + 1
            .Cells(8 + Y, 1) = Y 'No.
            .Cells(8 + Y, 1).HorizontalAlignment = xlCenter
            .Cells(8 + Y, 2) = "'" & Frm85.MSFlexGrid9.TextMatrix(x, 3) 'Tarikh Belian
            .Cells(8 + Y, 2).HorizontalAlignment = xlCenter
            
            .Cells(8 + Y, 3) = Frm85.MSFlexGrid9.TextMatrix(x, 4) 'No. Siri Produk
            .Cells(8 + Y, 3).HorizontalAlignment = xlCenter
            
            .Cells(8 + Y, 4) = Frm85.MSFlexGrid9.TextMatrix(x, 5) 'Purity
            .Cells(8 + Y, 4).HorizontalAlignment = xlCenter
            
            .Cells(8 + Y, 5) = Frm85.MSFlexGrid9.TextMatrix(x, 6) 'Kategori Produk
            
            .Cells(8 + Y, 6) = Frm85.MSFlexGrid9.TextMatrix(x, 7) 'Supplier
            
            .Cells(8 + Y, 7).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 7) = Frm85.MSFlexGrid9.TextMatrix(x, 8) 'Berat (g)
            .Cells(8 + Y, 7).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 8).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 8) = Frm85.MSFlexGrid9.TextMatrix(x, 9) 'Rate Penerimaan (RM/g)
            .Cells(8 + Y, 8).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 9).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 9) = Frm85.MSFlexGrid9.TextMatrix(x, 10) 'Upah (RM)
            .Cells(8 + Y, 9).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 10).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 10) = Frm85.MSFlexGrid9.TextMatrix(x, 11) 'Spread (%)
            .Cells(8 + Y, 10).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 11).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 11) = Frm85.MSFlexGrid9.TextMatrix(x, 12) 'Harga Selepas Spread (RM)
            .Cells(8 + Y, 11).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 12).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 12) = Frm85.MSFlexGrid9.TextMatrix(x, 13) 'Adjustment (RM)
            .Cells(8 + Y, 12).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 13).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 13) = Frm85.MSFlexGrid9.TextMatrix(x, 14) 'Harga Belian (RM)
            .Cells(8 + Y, 13).HorizontalAlignment = xlRight
            
            .Cells(8 + Y, 14) = Frm85.MSFlexGrid9.TextMatrix(x, 15) 'Dulang
            .Cells(8 + Y, 14).HorizontalAlignment = xlCenter
            
            .Cells(8 + Y, 15) = Frm85.MSFlexGrid9.TextMatrix(x, 16) 'Panjang
            .Cells(8 + Y, 15).HorizontalAlignment = xlCenter
            
            .Cells(8 + Y, 16) = Frm85.MSFlexGrid9.TextMatrix(x, 17) 'Lebar
            .Cells(8 + Y, 16).HorizontalAlignment = xlCenter
            
            .Cells(8 + Y, 17) = Frm85.MSFlexGrid9.TextMatrix(x, 18) 'Saiz
            .Cells(8 + Y, 17).HorizontalAlignment = xlCenter

            For Col = 1 To 17
                .Cells(8 + Y, Col).Borders.LineStyle = xlContinuous
            Next Col
        Next x
    
        Y = Y + 2
        .Cells(8 + Y, 1) = "Bilangan : " & Frm85.L76_Text 'Total Barang
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah Berat (g) : " & Frm85.L77_Text 'Total Berat
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah Modal : RM " & Frm85.L78_Text 'Total Harga Belian
        
        Y = Y + 4
        .Cells(8 + Y, 1).Font.Bold = True
        .Cells(8 + Y, 1) = "Report Generated By Sankyu System [ +6010 - 900 4788 , sankyusystem@gmail.com ]" 'Watermark Sankyu System
    End With
    
    ' This makes Excel visible
    xlObject.Visible = True
    xlObject.EnableEvents = True
End If
End Sub

Private Sub frm85_sm_hilang_Click()
'on error resume next
LM_FOUND = 0

frm85_LM_No_ID = vbNullString

If IsNumeric(Frm85.LV4.SelectedItem.Index) Then
    
    frm85_LM_No_ID = Frm85.LV4.ListItems(Frm85.LV4.SelectedItem.Index)
    
    If frm85_LM_No_ID <> vbNullString Then
    

        Note = "Adakah anda ingin menukar status barang ini?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Barang ini akan dimasukkan ke dalam senarai barang yang hilang , kecurian atau sebagainya." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from data_database where ID='" & frm85_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                Call frm125_initial_setting
                
                frm125.L7_Text = frm85_LM_No_ID
                If Not IsNull(rs!no_siri_Produk) Then frm125.L1_Text = rs!no_siri_Produk
                If Not IsNull(rs!kategori_Produk) Then frm125.L2_Text = rs!kategori_Produk
                If Not IsNull(rs!purity) Then frm125.L3_Text = rs!purity
                If Not IsNull(rs!beza_berat) Then frm125.L4_Text = Format(rs!beza_berat, "#,##0.00g")
                If Not IsNull(rs!harga_item) Then frm125.L5_Text = "RM " & Format(rs!harga_item, "#,##0.00")
                If Not IsNull(rs!dulang) Then frm125.L6_Text = rs!dulang

                LM_FOUND = 1
                
            Else
            
                MsgBox "Tiada maklumat berkenaan barang ini. Sila keluar dari menu ini dan cuba lagi.", vbExclamation, "Info"
                
            End If
            
            rs.Close
            Set rs = Nothing
            
            If LM_FOUND = 1 Then
                
                frm125.Show 1
                'frm125.TB1.SetFocus
                
            End If
            
        End If
        
    End If
    
End If
End Sub
Private Sub Frm85_SM_Padam_Data_Click()
'on error resume next
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim Frm85_LM_BERAT_ASAL As Double
Dim Frm85_LM_BEZA_BERAT As Double
Dim Frm85_LM_BERAT_JUALAN As Double
Dim Frm85_LM_BERAT_ASAL_COMP As Double
Dim Frm85_LM_BERAT_SELEPAS_COMP As Double
Dim Frm84_LM_MATA_ASAL As Double
Dim Frm84_LM_MATA_TEBUS As Double
Dim Frm84_LM_MATA_DAPAT As Double
Dim Frm85_LM_REFUND_ASAL As Double
Dim Frm85_LM_REFUND_GUNA As Double
Dim Frm85_SUSUT_BERAT As Double

Frm85_SUSUT_BERAT = 0
Frm85_LM_BERAT_ASAL_COMP = 0
Frm85_LM_BERAT_SELEPAS_COMP = 0

Frm84_LM_No_PELANGGAN_MATA = vbNullString

Frm84_LM_MATA_ASAL = 0
Frm84_LM_MATA_TEBUS = 0
Frm84_LM_MATA_DAPAT = 0
Frm84_LM_FLAG_MATA_ASAL = 0 '0 : Tiada perolehan mata , 1 : Ada perolehan mata (Pada awal pembelian asal)

Frm85_LM_FLAG_TI = 0 '0 : Tiada urusan trade in , 1 : Ada urusan trade in
Frm85_LM_JENIS = 0
DATA_FOUND = 0

If IsNumeric(Frm85.LV1.SelectedItem.Index) Then
    
    frm85_LM_No_ID = Frm85.LV1.ListItems(Frm85.LV1.SelectedItem.Index)
    
    If frm85_LM_No_ID <> vbNullString Then
    
        Note = "Padam Invoice Ini ?" & vbCrLf & _
                "" & vbCrLf & _
                "Semua Barang Dari No. Invoice Ini Akan Dipulangkan Ke Dalam Stok Kedai." & vbCrLf & _
                "" & vbCrLf & _
                "Teruskan ?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
        LM_NOW = Now
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 23_senarai_jualan where ID='" & frm85_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!jenis_jualan) Then
                    If rs!jenis_jualan = 0 Then
                        Frm85_LM_JENIS = 0
                    ElseIf rs!jenis_jualan = 1 Then
                        Frm85_LM_JENIS = 1
                    End If
                End If
                If Not IsNull(rs!no_resit) Then
                    Frm85_LM_NO_RESIT = rs!no_resit 'No. Resit
                    G_No_RESIT_JUALAN = rs!no_resit 'No. Invoice
                    DATA_FOUND = 1
                End If
            End If
            
            rs.Close
            Set rs = Nothing

            If DATA_FOUND = 1 Then
                If Frm85_LM_JENIS = 0 Then
                
                    If G_SPKE_ME_MAIL = "YES" Then
                        
                        LM_NOW = Now
                        
                        Set rs = New ADODB.Recordset
                        Call Main
                        rs.Open "select * from 88_sales_notfication", cn, adOpenKeyset, adLockOptimistic
                        
                        rs.AddNew
                        rs!no_invoice_asal = G_No_RESIT_JUALAN 'No. invoice rasmi
                        rs!jenis = 2
                        rs!jenis_report = 0 '0 : Jualan , 1 : Trade In
                        rs!write_timestamp = LM_NOW
                        rs!terminal = G_TERMINAL
                        rs!Status = 0
                        rs.Update
                        
                        rs.Close
                        Set rs = Nothing
        
                        Shell "cmd.exe /c " & G_SPKE_NE_PATH
                        
                    End If
                    
'### Periksa Samada Ada Pembayaran Menggunakan Simpanan Duit Di Kedai ### - Start 13/07/2015
    '### Jika Ada Perlu Pulangkan Dahulu Simpanan Tersebut ###
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from 22_jualan where no_resit='" & Frm85_LM_NO_RESIT & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                        If Not IsNull(rs!point_ari_nashi) Then
                            If rs!point_ari_nashi = 1 Then Frm84_LM_FLAG_MATA_ASAL = 1 '0 : Tiada perolehan mata , 1 : Ada perolehan mata (Pada awal pembelian asal)
                        End If
                        If Not IsNull(rs!duit_simpanan_kedai) Then
                            If rs!duit_simpanan_kedai <> "0.00" Then
                                If IsNumeric(rs!duit_simpanan_kedai) Then Frm85_LM_REFUND_ASAL = rs!duit_simpanan_kedai 'Refund : Jumlah Simpanan Asal
                                
                                If Not IsNull(rs!no_rujukan_pembeli) Then
                                    Frm85_LM_No_PELANGGAN = rs!no_rujukan_pembeli 'No. Pelanggan
                                End If
                            End If
                        End If
                        If Not IsNull(rs!flag_trade_in) Then
                            If rs!flag_trade_in = 1 Then
                            
' ### Pulangkan status barang trade in jika ada digunakan dalam urusan jualan ini ### - Start
                                Set rs1 = New ADODB.Recordset
                                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                                rs1.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & rs!no_resit_trade_in & "'", cn, adOpenKeyset, adLockOptimistic
                                
                                If Not rs1.EOF Then
                                    
                                    If Not IsNull(rs1!trade_in_status) Then
                                    
                                        If rs1!trade_in_status = 1 Then

                                            rs1!trade_in_status = 0
                                            rs1.Update
                                            
                                        End If
                                    
                                    End If
                                End If
                                
                                rs1.Close
                                Set rs1 = Nothing
' ### Pulangkan status barang trade in jika ada digunakan dalam urusan jualan ini ### - End

                            End If
                            
                            If Not IsNull(rs!jenis_trade_in) Then
                            
                                If rs!jenis_trade_in = 2 Then
                                
                                    Set rs1 = New ADODB.Recordset
                                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                                    rs1.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & rs!no_resit_trade_in & "'", cn, adOpenKeyset, adLockOptimistic
                                    
                                    If Not rs1.EOF Then
                                        
                                        rs1!Status = 0
                                        'rs1.Delete
                                        rs1.Update
                                        
                                    End If
                                    
                                    rs1.Close
                                    Set rs1 = Nothing
                                    
                                    Set rs1 = New ADODB.Recordset
                                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                                    rs1.Open "select * from data_database where bill_No_Trade_In='" & rs!no_resit_trade_in & "'", cn, adOpenKeyset, adLockOptimistic
                                    
                                    While rs1.EOF = False
                                        
                                        rs1!StatusItem = 0
                                        'rs1.Delete
                                        rs1.Update
                                        
                                        rs1.MoveNext
                                    Wend
                                    
                                    rs1.Close
                                    Set rs1 = Nothing
                                    
                                End If
                                
                            End If
                        End If
                    End If

                    rs.Close
                    Set rs = Nothing
                    
'### Pulangkan point/mata kepada ahli ### - Start
                    If Frm84_LM_FLAG_MATA_ASAL = 1 Then '0 : Tiada perolehan mata , 1 : Ada perolehan mata (Pada awal pembelian asal)
                    
                        Set rs = New ADODB.Recordset
                        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                        rs.Open "select * from 71_tebus_agih_point where no_invoice='" & Frm85_LM_NO_RESIT & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
                        
                        If Not rs.EOF Then
                        
                            If Not IsNull(rs!no_ahli) Then Frm84_LM_No_PELANGGAN_MATA = rs!no_ahli
                    
                            If Not IsNull(rs!jumlah_peroleh_point) Then 'Jumlah perolehan mata
                                If IsNumeric(rs!jumlah_peroleh_point) Then Frm84_LM_MATA_DAPAT = rs!jumlah_peroleh_point
                            End If
                            If Not IsNull(rs!jumlah_tebus_point) Then 'Jumlah mata yang ditebus
                                If IsNumeric(rs!jumlah_tebus_point) Then Frm84_LM_MATA_TEBUS = rs!jumlah_tebus_point
                            End If
                            rs!Status = 0
                            rs!write_timestamp3 = Now
                            rs.Update
                            
                        End If
                            
                        rs.Close
                        Set rs = Nothing
                        
                        If Frm84_LM_No_PELANGGAN_MATA <> vbNullString Then
                    
                            Set rs = New ADODB.Recordset
                            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                            rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm84_LM_No_PELANGGAN_MATA & "'", cn, adOpenKeyset, adLockOptimistic
                            
                            If Not rs.EOF Then
                                
                                If Not IsNull(rs!baki_point) Then 'Baki mata asal
                                    If IsNumeric(rs!baki_point) Then Frm84_LM_MATA_ASAL = rs!baki_point
                                End If
                                rs!baki_point = Frm84_LM_MATA_ASAL + Frm84_LM_MATA_TEBUS - Frm84_LM_MATA_DAPAT
                                rs.Update
                                
                            End If
                            
                            rs.Close
                            Set rs = Nothing
                        
                        End If
                        
                    End If
'### Pulangkan point/mata kepada ahli ### - End
        
                    If Frm85_LM_No_PELANGGAN <> vbNullString Then
                    
                        Set rs = New ADODB.Recordset
                        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                        rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm85_LM_No_PELANGGAN & "'", cn, adOpenKeyset, adLockOptimistic
                        
                        If Not rs.EOF Then
                            If IsNumeric(rs!baki_simpanan) Then Frm85_LM_REFUND_GUNA = rs!baki_simpanan 'Refund : Jumlah Simpan Yang Telah Digunakan Sebelum Ini
                            
                            rs!baki_simpanan = Format(Frm85_LM_REFUND_ASAL + Frm85_LM_REFUND_GUNA, "#,##0.00") 'Baki Simpanan
                            rs.Update
                        End If
                        
                        rs.Close
                        Set rs = Nothing
                
'### Padam Rekod Penggunaan Duit Pelanggan ### - Start
                        Set rs = New ADODB.Recordset
                        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                        rs.Open "select * from 24_rekod_kewangan_pelanggan where NO_RESIT='" & Frm85_LM_NO_RESIT & "'", cn, adOpenKeyset, adLockOptimistic
                        
                        If Not rs.EOF Then
                            rs.Delete
                            rs.Update
                        End If
                        
                        rs.Close
                        Set rs = Nothing
'### Padam Rekod Penggunaan Duit Pelanggan ### - End
                    End If
'### Periksa Samada Ada Pembayaran Menggunakan Simpanan Duit Di Kedai ### - End

'### Padam Rekod Belian Pembeli Yang Tidak Berdaftar Dengan Kedai ### - Start (08-07-2015)
        
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from 44_senarai_pelanggan where no_resit='" & Frm85_LM_NO_RESIT & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                        
                        rs.Delete
                        rs.Update
                    
                    End If
                    
                    rs.Close
                    Set rs = Nothing

'### Padam Rekod Belian Pembeli Yang Tidak Berdaftar Dengan Kedai ### - End (08-07-2015)
                   
'### Padam Resit ### - Start
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from 22_jualan where no_resit='" & Frm85_LM_NO_RESIT & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                        'rs.Delete
                        rs!Status = 0
                        rs.Update
                    End If
                    
                    rs.Close
                    Set rs = Nothing
'### Padam Resit ### - End

'### Padam Senarai Barang Dari No. Resit Ni Dan Pulangkan Stok ### - Start
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from 23_senarai_jualan where no_resit='" & Frm85_LM_NO_RESIT & "' AND status_rekod = 1", cn, adOpenKeyset, adLockOptimistic
    
                    While rs.EOF = False
                        Frm85_LM_BERAT_ASAL = 0
                        Frm85_LM_BEZA_BERAT = 0
                        Frm85_LM_BERAT_JUALAN = 0
                        Frm85_SUSUT_BERAT = 0
                        
                        Set rs2 = New ADODB.Recordset
                        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                        rs2.Open "select * from data_database where no_siri_Produk='" & rs!no_siri_Produk & "'", cn, adOpenKeyset, adLockOptimistic
                        
                        If Not rs2.EOF Then
                            If rs2!receiving_Status = 0 Or rs2!receiving_Status = 2 Then
                            
                                If Not IsNull(rs2!Berat) Then Frm85_LM_BERAT_ASAL = rs2!Berat
                                If Not IsNull(rs2!beza_berat) Then Frm85_LM_BEZA_BERAT = rs2!beza_berat
                                If Not IsNull(rs!berat_jualan) Then Frm85_LM_BERAT_JUALAN = rs!berat_jualan
                                If Not IsNull(rs2!susut_berat) Then Frm85_SUSUT_BERAT = Format(rs2!susut_berat, "0.00") 'Susut berat
                                
                                Frm85_LM_BERAT_ASAL_COMP = Format(Frm85_LM_BERAT_ASAL, "0.00")
                                Frm85_LM_BERAT_SELEPAS_COMP = Format(Frm85_LM_BERAT_JUALAN + Frm85_LM_BEZA_BERAT, "0.00")
                                
                                If Frm85_LM_BERAT_ASAL_COMP = Frm85_LM_BERAT_SELEPAS_COMP Then
                                    rs2!beza_berat = Format(Frm85_LM_BERAT_JUALAN + Frm85_LM_BEZA_BERAT, "0.00")
                                    rs2!StatusItem = 10
                                    rs2!tarikh_jualan1 = Null
                                Else
                                    rs2!beza_berat = Format(Frm85_LM_BERAT_JUALAN + Frm85_LM_BEZA_BERAT, "0.00")
                                    rs2!StatusItem = 12
                                    rs2!tarikh_jualan1 = DateTime.Date
                                End If
                            Else
                                rs2!StatusItem = 10
                            End If
                            rs2.Update
                        End If
                        
                        rs2.Close
                        Set rs2 = Nothing
 
                        '### Masukkan data lama ke dalam table #72_data_amendment ### - Start
                        If rs!Type = 0 Then
                        
                            Set rs3 = New ADODB.Recordset
                            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                            
                            strsql = "insert into 72_data_amendment(id_asal,no_siri_produk,kategori_produk,berat,upah,nama_supplier,jenis_barang,jenis,menu,write_timestamp)" & _
                                        "select ID,no_siri_produk,kategori_produk,Berat_Jualan,upah,harga_Semasa,0,1,1,Now() from 23_senarai_jualan WHERE no_siri_Produk='" & rs!no_siri_Produk & "'"
                            
                            Set rs3 = cn.Execute(strsql)
                            Set rs3 = Nothing
                            
                        ElseIf rs!Type = 1 Then
                        
                            Set rs3 = New ADODB.Recordset
                            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                            
                            strsql = "insert into 72_data_amendment(id_asal,no_siri_produk,kategori_produk,upah,jenis_barang,jenis,menu,write_timestamp)" & _
                                        "select ID,no_siri_produk,kategori_produk,harga_jualan,1,1,1,Now() from 23_senarai_jualan WHERE no_siri_Produk='" & rs!no_siri_Produk & "'"
                            
                            Set rs3 = cn.Execute(strsql)
                            Set rs3 = Nothing
                        
                        End If
                        '### Masukkan data lama ke dalam table #72_data_amendment ### - End
                        
                        'rs.Delete
                        rs!status_rekod = 0
                        rs.Update
                        rs.MoveNext
                    Wend
                
                    rs.Close
                    Set rs = Nothing
                    
                    '### Transfer data kepada recovery database ### - Start
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            
                    strsql = "insert into " & G_RECOVERY_DATABASE & ".85_penggunaan_ti(id_asal,tarikh,no_rujukan,purity,berat,write_timestamp,terminal,Status)" & _
                                "select ID,tarikh,no_rujukan,purity,berat,write_timestamp,terminal,Status " _
                                & "from " & G_SERVER_DATABASE & ".85_penggunaan_ti WHERE no_rujukan='" & Frm85_LM_NO_RESIT & "' AND status = 1"
                                
                    Set rs = cn.Execute(strsql)
                    Set rs = Nothing
                    '### Transfer data kepada recovery database ### - End
                    
                    '### Padam data asal (85_penggunaan_ti) ### - Start
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            
                    strsql = "DELETE FROM 85_penggunaan_ti WHERE no_rujukan='" & Frm85_LM_NO_RESIT & "' AND status = 1"
            
                    Set rs = cn.Execute(strsql)
                    Set rs = Nothing
                    '### Padam data asal (85_penggunaan_ti) ### - End

                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            
                    strsql = "UPDATE 93_trade_in_susut_niai set status = 0 , write_timestamp2='" & LM_NOW & "' WHERE no_invoice='" & Frm85_LM_NO_RESIT & "' AND status = 1"
            
                    Set rs = cn.Execute(strsql)
                    Set rs = Nothing
            
'### Update Log ### - Start
                    user = MDI_frm1.L3_Text
                    LogAct_Memory = "[" & user & "] Padam invoice jualan [" & Frm85_LM_NO_RESIT & "]"
                    LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                    Call UpdateLog_Database
'### Update Log ### - End
    
                    Note = "Data Telah Berjaya Dipadamkan." & vbCrLf & _
                            "Refresh Data Anda ? Sistem Akan Mengambil Sedikit Masa Untuk Refresh Data." & vbCrLf & _
                            "" & vbCrLf & _
                            "Teruskan ?"
    
    
                    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
                    
                    If Answer = vbNo Then
                    End If
                    If Answer = vbYes Then
                        GM_NEXT_PREV = 2
                        
                        If GM_REPORT_MODE = 2 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                            Call Frm85_Header_Report_Jualan
                            Call Frm85_Report_Jualan_page
                        ElseIf GM_REPORT_MODE = 3 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                            Call Frm85_Header_Report_Jualan
                            Call Frm85_carian_jualan_page
                        ElseIf GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
                            Call Frm85_Header_Report_Jualan
                            Call Frm85_Report_Jualan_barcode
                        End If

                    End If
'### Padam Senarai Barang Dari No. Resit Ni Dan Pulangkan Stok ### - End
                
                ElseIf Frm85_LM_JENIS = 1 Then
                
                    MsgBox "Item ini adalah jualan kepada pihak supplier/agen." & vbCrLf & _
                            "Untuk memadamkan data ini sila padam dari menu report GRN & GDN", vbInformation, "Info"
                            
                    'Call Frm85_padam_voucher
                    
                    Exit Sub
                    
                End If
                
                Call amendment_email_check
            End If
        End If
    End If
End If
End Sub
Private Sub Frm85_SM_Padam_Data2_Click()
'on error resume next
Dim rs3 As ADODB.Recordset
DATA_FOUND = 0

frm85_LM_No_ID = vbNullString

If IsNumeric(Frm85.LV2.SelectedItem.Index) Then
    
    frm85_LM_No_ID = Frm85.LV2.ListItems(Frm85.LV2.SelectedItem.Index)
    
    If frm85_LM_No_ID <> vbNullString Then
    
        Note = "Padam data stok ini?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "*** Peringatan ***" & vbCrLf & _
                "----DATA BARANG INI AKAN DIPADAMKAN DARI SISTEM DAN TIDAK DAPAT DIPULANGKAN KEMBALI DATA INI-----" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan ?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
'### Periksa Status Setiap Item Dari Batch Penerimaan Stok Ini ### - Start

            LM_NOW = Now
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from Data_Database where ID='" & frm85_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                'If Not IsNull(rs!NoRujukanSistem) Then
                    
                    If Not IsNull(rs!StatusItem) Then
                        
                        If rs!StatusItem <> "10" Then
                            
                            MsgBox "Item/Stok ini tidak dibenarkan untuk dipadamkan kerana status item ini telah berubah.", vbExclamation, "Info"
                            
                            rs.Close
                            Set rs = Nothing
                            
                            Exit Sub
                            
                        End If
                    
                    End If
                    
                    If Not IsNull(rs!no_siri_Produk) Then LM_NO_SIRI = rs!no_siri_Produk
                    
                    Frm85_LM_NO_RUJUKAN = rs!NoRujukanSistem 'No. Rujukan Belian
                    
                    rs!StatusItem = 0
                    rs.Update
                    
                    DATA_FOUND = 1
                'Else
                '    MsgBox "Tiada Data Rujukan Belian.", vbExclamation, "Error"
                'End If
            End If
            
            rs.Close
            Set rs = Nothing
    
            If DATA_FOUND = 1 Then

'### Periksa Status Setiap Item Dari Batch Penerimaan Stok Ini ### - End

                '### Masukkan data lama ke dalam table #72_data_amendment ### - Start
                Set rs3 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

                strsql = "insert into 72_data_amendment(id_asal,no_siri_produk,kategori_produk,nama_supplier,upah,jenis,menu,jenis_barang,write_timestamp)" & _
                            "select ID,no_siri_produk,kategori_produk,nama_supplier,harga_item,1,0,receiving_Status,'" & LM_NOW & "' from Data_Database WHERE NoRujukanSistem='" & Frm85_LM_NO_RUJUKAN & "' AND (receiving_Status = 1 Or receiving_Status = 3)"
                
                Set rs3 = cn.Execute(strsql)
                Set rs3 = Nothing
                
                Set rs3 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

                strsql = "insert into 72_data_amendment(id_asal,no_siri_produk,kategori_produk,nama_supplier,berat,upah,jenis,menu,jenis_barang,write_timestamp)" & _
                            "select ID,no_siri_produk,kategori_produk,nama_supplier,berat,upah,1,0,receiving_Status,'" & LM_NOW & "'  from Data_Database WHERE NoRujukanSistem='" & Frm85_LM_NO_RUJUKAN & "' AND (receiving_Status = 0 Or receiving_Status = 2 Or receiving_Status = 4 Or receiving_Status = 5)"
                
                Set rs3 = cn.Execute(strsql)
                Set rs3 = Nothing
                '### Masukkan data lama ke dalam table #72_data_amendment ### - End

'### Update Log ### - Start
                user = MDI_frm1.L3_Text
                LogAct_Memory = "[" & user & "] Padam data stok bagi [" & LM_NO_SIRI & "]"
                LogDate_Memory = LM_NOW
                Call UpdateLog_Database
'### Update Log ### - End
                
                GM_NEXT_PREV = 2
                
                If Frm101.L33_Text = 0 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                    If Frm101.CB2 = 1 Then 'Report Belian
                        Call Frm85_Header_Report_Belian
                        'Call Frm85_Report_Belian
                        Call Frm85_report_belian_page
                    End If
                    If Frm101.CB4 = 1 Then 'Report Jualan
                        Frm85_Header_Report_Buyback
                        'Call Frm85_Report_Buyback
                        Call Frm85_report_buyback_page
                    End If
                ElseIf Frm101.L33_Text = 1 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                    Call Frm85_Header_Report_Belian
                    'Call Frm85_search_berat
                    Call Frm85_search_berat_page
                ElseIf Frm101.L33_Text = 3 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                    Call Frm85_Header_Report_Buyback
                    'Call Frm85_carian_buyback
                    Call Frm85_carian_buyback_page
                ElseIf Frm101.L33_Text = 5 Then
                    Call Frm85_Header_Report_Belian
                    Call Frm85_report_belian_barcode
                End If
                
                Call amendment_email_check
                
                MsgBox "Data bagi item ini telah berjaya dipadamkan.", vbInformation, "Infro"
                
            End If
        End If
    End If
End If
End Sub
Private Sub Frm85_SM_Padam_Data3_Click()
'on error resume next
DATA_FOUND = 0
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset

Frm85_LM_JENIS_BELIAN = 0 '1 : Belian secara trade in barang , 2 : Belian secara trade in (voucher)
Frm85_LM_LOCKED = 0 '0 : Dibenarkan utk edit / padam , 1 : Tidak dibenarkan untuk edit / padam

frm85_LM_No_ID = vbNullString

If IsNumeric(Frm85.LV3.SelectedItem.Index) Then
    
    frm85_LM_No_ID = Frm85.LV3.ListItems(Frm85.LV3.SelectedItem.Index)
    
    If frm85_LM_No_ID <> vbNullString Then
    
        '### Periksa status invoice trade in ### - Start
        Set rs = New ADODB.Recordset
        'call main
        rs.Open "select * from Data_Database where ID='" & frm85_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            If rs!jenis_trade_in = 1 Then
            
                Frm85_LM_JENIS_BELIAN = 2 '1 : Belian secara trade in barang , 2 : Belian secara trade in (voucher)
                
            
            ElseIf rs!jenis_trade_in = 0 Then
                
                Frm85_LM_JENIS_BELIAN = 1 '1 : Belian secara trade in barang , 2 : Belian secara trade in (voucher)
            
            End If
            
            Set rs1 = New ADODB.Recordset
            'call main
            rs1.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & rs!bill_No_Trade_In & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs1.EOF Then
                
                If Not IsNull(rs1!trade_in_status) Then
                
                    If rs1!trade_in_status = 1 Then
                    
                        Set rs2 = New ADODB.Recordset
                        'call main
                        rs2.Open "select * from 22_jualan where no_resit_trade_in='" & rs!bill_No_Trade_In & "'", cn, adOpenKeyset, adLockOptimistic
                        
                        If Not rs2.EOF Then
                            If Not IsNull(rs2!no_resit) Then
                                Frm85_LM_No_INVOICE = rs2!no_resit 'No. invoice jualan
                                Frm85_LM_LOCKED = 1 '0 : Dibenarkan utk edit / padam , 1 : Tidak dibenarkan untuk edit / padam
                            End If
                        End If
                        
                        rs2.Close
                        Set rs2 = Nothing
                        
                    End If
                
                End If
            End If
            
            rs1.Close
            Set rs1 = Nothing
        
        End If
        
        rs.Close
        Set rs = Nothing
        
        If Frm85_LM_LOCKED = 1 Then '0 : Dibenarkan utk edit / padam , 1 : Tidak dibenarkan untuk edit / padam
            
            If Frm85_LM_JENIS_BELIAN = 1 Or Frm85_LM_JENIS_BELIAN = 2 Then '1 : Belian secara trade in barang , 2 : Belian secara trade in (voucher)
                
                MsgBox "Data bagi barang kemas ini tidak dibenarkan untuk diedit atau dipadamkan." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Sebab :" & vbCrLf & _
                        "Voucher atau data bagi data ini telah digunakan untuk jualan barang kepada pelanggan." & vbCrLf & _
                        "Untuk edit data ini perlu rujuk kepada invoice jualan [" & Frm85_LM_No_INVOICE & "] untuk edit atau padam data.", vbExclamation, "Info"
                
                Exit Sub
                
            End If
        
        End If
        '### Periksa status invoice trade in ### - End
            
        If frm85_LM_No_ID <> vbNullString Then
            Note = "Padam Data Ini ?" & vbCrLf & _
                    "" & vbCrLf & _
                    "*** Semua barang atau stok yang diterima bersama dengan item ini akan dipadamkan." & vbCrLf & _
                    "" & vbCrLf & _
                    "Teruskan ?"
                    
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbNo Then
                Exit Sub
            End If
            If Answer = vbYes Then
        '### Periksa Status Setiap Item Dari Batch Penerimaan Stok Ini ### - Start
                Set rs = New ADODB.Recordset
                'call main
                rs.Open "select * from Data_Database where ID='" & frm85_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Not IsNull(rs!NoRujukanSistem) Then
                        Frm85_LM_NO_RUJUKAN = rs!NoRujukanSistem 'No. Rujukan Belian
                        DATA_FOUND = 1
                    Else
                        MsgBox "Tiada Data Rujukan Belian.", vbExclamation, "Error"
                    End If
                End If
                
                rs.Close
                Set rs = Nothing
        
                If DATA_FOUND = 1 Then
                    Set rs = New ADODB.Recordset
                    'call main
                    rs.Open "select * from Data_Database where NoRujukanSistem='" & Frm85_LM_NO_RUJUKAN & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    While rs.EOF = False
                        If Not IsNull(rs!StatusItem) Then
                            If rs!StatusItem <> 10 And rs!StatusItem <> 0 Then
                                MsgBox "Item ini atau salah satu atau lebih dari barang/stok yang diterima bersama barang telah dijual (tiada dalam stok)." & vbCrLf & _
                                        "" & vbCrLf & _
                                        "Sila periksa setiap status bagi item barang yang diterima bersama item ini semasa proses penerimaan stok.", vbExclamation, "Info"
                                
                                rs.Close
                                Set rs = Nothing
                                
                                Exit Sub
                            End If
                        End If
                        rs.MoveNext
                    Wend
                    
                    rs.Close
                    Set rs = Nothing
        '### Periksa Status Setiap Item Dari Batch Penerimaan Stok Ini ### - End
        
                    '### Masukkan data lama ke dalam table #72_data_amendment ### - Start
                    Set rs3 = New ADODB.Recordset
                    'call main
        
                    strsql = "insert into 72_data_amendment(id_asal,no_siri_produk,kategori_produk,nama_supplier,upah,jenis,menu,jenis_barang,write_timestamp)" & _
                                "select ID,no_siri_produk,kategori_produk,nama_supplier,harga_item,1,0,receiving_Status,Now() from Data_Database WHERE NoRujukanSistem='" & Frm85_LM_NO_RUJUKAN & "' AND (receiving_Status = 1 Or receiving_Status = 3)"
                    
                    Set rs3 = cn.Execute(strsql)
                    Set rs3 = Nothing
                    
                    Set rs3 = New ADODB.Recordset
                    'call main
        
                    strsql = "insert into 72_data_amendment(id_asal,no_siri_produk,kategori_produk,nama_supplier,berat,upah,jenis,menu,jenis_barang,write_timestamp)" & _
                                "select ID,no_siri_produk,kategori_produk,nama_supplier,berat,upah,1,0,receiving_Status,Now() from Data_Database WHERE NoRujukanSistem='" & Frm85_LM_NO_RUJUKAN & "' AND (receiving_Status = 0 Or receiving_Status = 2 Or receiving_Status = 4 Or receiving_Status = 5)"
                    
                    Set rs3 = cn.Execute(strsql)
                    Set rs3 = Nothing
                    '### Masukkan data lama ke dalam table #72_data_amendment ### - End
        
        '### Padam Data - Database ### - Start
                    Set rs = New ADODB.Recordset
                    'call main
                    rs.Open "select * from Data_Database where NoRujukanSistem='" & Frm85_LM_NO_RUJUKAN & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    While rs.EOF = False
                        'rs.Delete
                        rs!StatusItem = 0
                        rs.Update
                        
                        rs.MoveNext
                    Wend
                    
                    rs.Close
                    Set rs = Nothing
        '### Padam Data - Database ### - End
        
        '### Padam Data - Akaun ### - Start
                    Set rs = New ADODB.Recordset
                    'call main
                    rs.Open "select * from 16_gold_bar_belian where no_rujukan='" & Frm85_LM_NO_RUJUKAN & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    While rs.EOF = False
                        'rs.Delete
                        rs!Status = 0
                        rs.Update
                        
                        rs.MoveNext
                    Wend
                    
                    rs.Close
                    Set rs = Nothing
        '### Padam Data - Akaun ### - End
        
        '### Update Log ### - Start
                    user = MDI_frm1.L3_Text
                    LogAct_Memory = "[" & user & "] Padam Data Stok. No. Rujukan [" & Frm85_LM_NO_RUJUKAN & "]"
                    LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                    Call UpdateLog_Database
        '### Update Log ### - End
        
                    Note = "Data Telah Berjaya Dipadamkan." & vbCrLf & _
                            "Refresh Data Anda ?"
        
                    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
                    
                    If Answer = vbNo Then
                        '
                    End If
                    If Answer = vbYes Then
                        GM_NEXT_PREV = 2
                        
                        If Frm101.L33_Text = 0 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                            If Frm101.CB2 = 1 Then 'Report Belian
                                Call Frm85_Header_Report_Belian
                                'Call Frm85_Report_Belian
                                Call Frm85_report_belian_page
                            End If
                            If Frm101.CB4 = 1 Then 'Report Jualan
                                Frm85_Header_Report_Buyback
                                'Call Frm85_Report_Buyback
                                Call Frm85_report_buyback_page
                            End If
                        ElseIf Frm101.L33_Text = 1 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                            Call Frm85_Header_Report_Belian
                            'Call Frm85_search_berat
                            Call Frm85_search_berat_page
                        ElseIf Frm101.L33_Text = 3 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                            Call Frm85_Header_Report_Buyback
                            'Call Frm85_carian_buyback
                            Call Frm85_carian_buyback_page
                        End If
                    End If
                    
                    Call amendment_email_check
                    
                End If
            End If
        End If

    End If
End If
End Sub
Private Sub Frm85_SM_Padam_Data4_Click()
'on error resume next
Dim rs3 As ADODB.Recordset
DATA_FOUND = 0

If Frm85.MSFlexGrid8 <> vbNullString Then
    Frm85_LM_ID = Frm85.MSFlexGrid8.TextMatrix(Frm85.MSFlexGrid8, 2) 'No. ID
    
    If Frm85_LM_ID <> vbNullString Then
        Note = "Padam Data Ini ?" & vbCrLf & _
                "" & vbCrLf & _
                "*** Semua barang atau stok yang diterima bersama dengan item ini akan dipadamkan." & vbCrLf & _
                "" & vbCrLf & _
                "Teruskan ?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
'### Periksa Status Setiap Item Dari Batch Penerimaan Stok Ini ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from Data_Database where ID='" & Frm85_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!NoRujukanSistem) Then
                    Frm85_LM_NO_RUJUKAN = rs!NoRujukanSistem 'No. Rujukan Belian
                    DATA_FOUND = 1
                Else
                    MsgBox "Tiada Data Rujukan Belian.", vbExclamation, "Error"
                End If
            End If
            
            rs.Close
            Set rs = Nothing
    
            If DATA_FOUND = 1 Then
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from Data_Database where NoRujukanSistem='" & Frm85_LM_NO_RUJUKAN & "'", cn, adOpenKeyset, adLockOptimistic
                
                While rs.EOF = False
                    If Not IsNull(rs!StatusItem) Then
                        If rs!StatusItem <> 10 Then
                            MsgBox "Item ini atau salah satu atau lebih dari barang/stok yang diterima bersama barang telah dijual (tiada dalam stok)." & vbCrLf & _
                                    "" & vbCrLf & _
                                    "Sila periksa setiap status bagi item barang yang diterima bersama item ini semasa proses penerimaan stok.", vbExclamation, "Info"
                            
                            rs.Close
                            Set rs = Nothing
                            
                            Exit Sub
                        End If
                    End If
                    rs.MoveNext
                Wend
                
                rs.Close
                Set rs = Nothing
'### Periksa Status Setiap Item Dari Batch Penerimaan Stok Ini ### - End

                '### Masukkan data lama ke dalam table #72_data_amendment ### - Start
                Set rs3 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

                strsql = "insert into 72_data_amendment(id_asal,no_siri_produk,kategori_produk,nama_supplier,upah,jenis,menu,jenis_barang,write_timestamp)" & _
                            "select ID,no_siri_produk,kategori_produk,nama_supplier,harga_item,1,0,receiving_Status,Now() from Data_Database WHERE NoRujukanSistem='" & Frm85_LM_NO_RUJUKAN & "' AND (receiving_Status = 1 Or receiving_Status = 3)"
                
                Set rs3 = cn.Execute(strsql)
                Set rs3 = Nothing
                
                Set rs3 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

                strsql = "insert into 72_data_amendment(id_asal,no_siri_produk,kategori_produk,nama_supplier,berat,upah,jenis,menu,jenis_barang,write_timestamp)" & _
                            "select ID,no_siri_produk,kategori_produk,nama_supplier,berat,upah,1,0,receiving_Status,Now() from Data_Database WHERE NoRujukanSistem='" & Frm85_LM_NO_RUJUKAN & "' AND (receiving_Status = 0 Or receiving_Status = 2 Or receiving_Status = 4 Or receiving_Status = 5)"
                
                Set rs3 = cn.Execute(strsql)
                Set rs3 = Nothing
                '### Masukkan data lama ke dalam table #72_data_amendment ### - End

'### Padam Data - Database ### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from Data_Database where NoRujukanSistem='" & Frm85_LM_NO_RUJUKAN & "'", cn, adOpenKeyset, adLockOptimistic
                
                While rs.EOF = False
                    'rs.Delete
                    rs!StatusItem = 0
                    rs.Update
                    
                    rs.MoveNext
                Wend
                
                rs.Close
                Set rs = Nothing
'### Padam Data - Database ### - End

'### Padam Data - Akaun ### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 16_gold_bar_belian where no_rujukan='" & Frm85_LM_NO_RUJUKAN & "'", cn, adOpenKeyset, adLockOptimistic
                
                While rs.EOF = False
                    'rs.Delete
                    rs!Status = 0
                    rs.Update
                    
                    rs.MoveNext
                Wend
                
                rs.Close
                Set rs = Nothing
'### Padam Data - Akaun ### - End

'### Update Log ### - Start
                user = MDI_frm1.L3_Text
                LogAct_Memory = "[" & user & "] Padam Data Stok. No. Rujukan [" & Frm85_LM_NO_RUJUKAN & "]"
                LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                Call UpdateLog_Database
'### Update Log ### - End

                Note = "Data Telah Berjaya Dipadamkan." & vbCrLf & _
                        "Refresh Data Anda ?"
    
                Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
                
                If Answer = vbNo Then
                    '
                End If
                If Answer = vbYes Then
                    GM_NEXT_PREV = 2
                    
                    If Frm101.L33_Text = 0 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                        If Frm101.CB2 = 1 Then 'Report Belian
                            Call Frm85_Header_Report_Belian
                            'Call Frm85_Report_Belian
                            Call Frm85_report_belian_page
                        End If
                        If Frm101.CB4 = 1 Then 'Report Jualan
                            Frm85_Header_Report_Buyback
                            'Call Frm85_Report_Buyback
                            Call Frm85_report_buyback_page
                        End If
                        If Frm101.CB11 = 1 Then 'Report Belian Gold Bar
                            Call Frm85_Header_Report_belian_gb
                            Call Frm85_report_belian_gb_page
                        End If
                        If Frm101.CB12 = 1 Then 'Report Buyback Gold Bar
                            Frm85_Header_Report_Buyback
                            Call Frm85_report_buyback_gb_page
                        End If
                    ElseIf Frm101.L33_Text = 1 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                        Call Frm85_Header_Report_Belian
                        'Call Frm85_search_berat
                        Call Frm85_search_berat_page
                    ElseIf Frm101.L33_Text = 3 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                        Call Frm85_Header_Report_Buyback
                        'Call Frm85_carian_buyback
                        Call Frm85_carian_buyback_page
                    End If

                End If
                
                Call amendment_email_check
                
            End If
        End If
    End If
End If
End Sub
Private Sub Frm85_SM_Padam_Data5_Click()
'on error resume next
DATA_FOUND = 0
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset

Frm85_LM_JENIS_BELIAN = 0 '1 : Belian secara trade in barang , 2 : Belian secara trade in (voucher)
Frm85_LM_LOCKED = 0 '0 : Dibenarkan utk edit / padam , 1 : Tidak dibenarkan untuk edit / padam

If Frm85.MSFlexGrid9 <> vbNullString Then
    Frm85_LM_ID = Frm85.MSFlexGrid9.TextMatrix(Frm85.MSFlexGrid9, 2) 'No. ID
    
    If Frm85_LM_ID <> vbNullString Then
    
'### Periksa status invoice trade in ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where ID='" & Frm85_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            If rs!jenis_trade_in = 1 Then
            
                Frm85_LM_JENIS_BELIAN = 2 '1 : Belian secara trade in barang , 2 : Belian secara trade in (voucher)
                
            
            ElseIf rs!jenis_trade_in = 0 Then
                
                Frm85_LM_JENIS_BELIAN = 1 '1 : Belian secara trade in barang , 2 : Belian secara trade in (voucher)
            
            End If
            
            Set rs1 = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs1.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & rs!bill_No_Trade_In & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs1.EOF Then
                
                If Not IsNull(rs1!trade_in_status) Then
                
                    If rs1!trade_in_status = 1 Then
                    
                        Set rs2 = New ADODB.Recordset
                        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                        rs2.Open "select * from 22_jualan where no_resit_trade_in='" & rs!bill_No_Trade_In & "'", cn, adOpenKeyset, adLockOptimistic
                        
                        If Not rs2.EOF Then
                            If Not IsNull(rs2!no_resit) Then
                                Frm85_LM_No_INVOICE = rs2!no_resit 'No. invoice jualan
                                Frm85_LM_LOCKED = 1 '0 : Dibenarkan utk edit / padam , 1 : Tidak dibenarkan untuk edit / padam
                            End If
                        End If
                        
                        rs2.Close
                        Set rs2 = Nothing
                        
                    End If
                
                End If
            End If
            
            rs1.Close
            Set rs1 = Nothing
        
        End If
        
        rs.Close
        Set rs = Nothing
        
        If Frm85_LM_LOCKED = 1 Then '0 : Dibenarkan utk edit / padam , 1 : Tidak dibenarkan untuk edit / padam
            
            If Frm85_LM_JENIS_BELIAN = 1 Or Frm85_LM_JENIS_BELIAN = 2 Then '1 : Belian secara trade in barang , 2 : Belian secara trade in (voucher)
                
                MsgBox "Data bagi barang kemas ini tidak dibenarkan untuk diedit atau dipadamkan." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Sebab :" & vbCrLf & _
                        "Voucher atau data bagi data ini telah digunakan untuk jualan barang kepada pelanggan." & vbCrLf & _
                        "Untuk edit data ini perlu rujuk kepada invoice jualan [" & Frm85_LM_No_INVOICE & "] untuk edit atau padam data.", vbExclamation, "Info"
                
                Exit Sub
                
            End If
    
        End If
'### Periksa status invoice trade in ### - End

        Note = "Padam Data Ini ?" & vbCrLf & _
                "" & vbCrLf & _
                "*** Semua barang atau stok yang diterima bersama dengan item ini akan dipadamkan." & vbCrLf & _
                "" & vbCrLf & _
                "Teruskan ?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
'### Periksa Status Setiap Item Dari Batch Penerimaan Stok Ini ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from Data_Database where ID='" & Frm85_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!NoRujukanSistem) Then
                    Frm85_LM_NO_RUJUKAN = rs!NoRujukanSistem 'No. Rujukan Belian
                    DATA_FOUND = 1
                Else
                    MsgBox "Tiada Data Rujukan Belian.", vbExclamation, "Error"
                End If
            End If
            
            rs.Close
            Set rs = Nothing
    
            If DATA_FOUND = 1 Then
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from Data_Database where NoRujukanSistem='" & Frm85_LM_NO_RUJUKAN & "'", cn, adOpenKeyset, adLockOptimistic
                
                While rs.EOF = False
                    If Not IsNull(rs!StatusItem) Then
                        If rs!StatusItem <> 10 Then
                            MsgBox "Item ini atau salah satu atau lebih dari barang/stok yang diterima bersama barang telah dijual (tiada dalam stok)." & vbCrLf & _
                                    "" & vbCrLf & _
                                    "Sila periksa setiap status bagi item barang yang diterima bersama item ini semasa proses penerimaan stok.", vbExclamation, "Info"
                            
                            rs.Close
                            Set rs = Nothing
                            
                            Exit Sub
                        End If
                    End If
                    rs.MoveNext
                Wend
                
                rs.Close
                Set rs = Nothing
'### Periksa Status Setiap Item Dari Batch Penerimaan Stok Ini ### - End

                '### Masukkan data lama ke dalam table #72_data_amendment ### - Start
                Set rs3 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

                strsql = "insert into 72_data_amendment(id_asal,no_siri_produk,kategori_produk,nama_supplier,upah,jenis,menu,jenis_barang,write_timestamp)" & _
                            "select ID,no_siri_produk,kategori_produk,nama_supplier,harga_item,1,0,receiving_Status,Now() from Data_Database WHERE NoRujukanSistem='" & Frm85_LM_NO_RUJUKAN & "' AND (receiving_Status = 1 Or receiving_Status = 3)"
                
                Set rs3 = cn.Execute(strsql)
                Set rs3 = Nothing
                
                Set rs3 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

                strsql = "insert into 72_data_amendment(id_asal,no_siri_produk,kategori_produk,nama_supplier,berat,upah,jenis,menu,jenis_barang,write_timestamp)" & _
                            "select ID,no_siri_produk,kategori_produk,nama_supplier,berat,upah,1,0,receiving_Status,Now() from Data_Database WHERE NoRujukanSistem='" & Frm85_LM_NO_RUJUKAN & "' AND (receiving_Status = 0 Or receiving_Status = 2 Or receiving_Status = 4 Or receiving_Status = 5)"
                
                Set rs3 = cn.Execute(strsql)
                Set rs3 = Nothing
                '### Masukkan data lama ke dalam table #72_data_amendment ### - End

'### Padam Data - Database ### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from Data_Database where NoRujukanSistem='" & Frm85_LM_NO_RUJUKAN & "'", cn, adOpenKeyset, adLockOptimistic
                
                While rs.EOF = False
                    rs!StatusItem = 0
                    'rs.Delete
                    rs.Update
                    
                    rs.MoveNext
                Wend
                
                rs.Close
                Set rs = Nothing
'### Padam Data - Database ### - End

'### Padam Data - Akaun ### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 16_gold_bar_belian where no_rujukan='" & Frm85_LM_NO_RUJUKAN & "'", cn, adOpenKeyset, adLockOptimistic
                
                While rs.EOF = False
                    rs!StatusItem = 0
                    'rs.Delete
                    rs.Update
                    
                    rs.MoveNext
                Wend
                
                rs.Close
                Set rs = Nothing
'### Padam Data - Akaun ### - End

'### Update Log ### - Start
                user = MDI_frm1.L3_Text
                LogAct_Memory = "[" & user & "] Padam Data Stok. No. Rujukan [" & Frm85_LM_NO_RUJUKAN & "]"
                LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                Call UpdateLog_Database
'### Update Log ### - End

                

                Note = "Data Telah Berjaya Dipadamkan." & vbCrLf & _
                        "Refresh Data Anda ?"
    
                Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
                
                If Answer = vbNo Then
                    '
                End If
                If Answer = vbYes Then
                    GM_NEXT_PREV = 2
                    
                    If Frm101.L33_Text = 0 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                        If Frm101.CB2 = 1 Then 'Report Belian
                            Call Frm85_Header_Report_Belian
                            Call Frm85_report_belian_page
                        End If
                        If Frm101.CB4 = 1 Then 'Report Jualan
                            Frm85_Header_Report_Buyback
                            Call Frm85_report_buyback_page
                        End If
                        If Frm101.CB11 = 1 Then 'Report Belian Gold Bar
                            Call Frm85_Header_Report_belian_gb
                            Call Frm85_report_belian_gb_page
                        End If
                        If Frm101.CB12 = 1 Then 'Report Buyback Gold Bar
                            Frm85_Header_Report_Buyback
                            Call Frm85_report_buyback_gb_page
                        End If
                    ElseIf Frm101.L33_Text = 1 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                        Call Frm85_Header_Report_Belian
                        'Call Frm85_search_berat
                        Call Frm85_search_berat_page
                    ElseIf Frm101.L33_Text = 3 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                        Call Frm85_Header_Report_Buyback
                        'Call Frm85_carian_buyback
                        Call Frm85_carian_buyback_page
                    End If
                End If
                
                Call amendment_email_check

            End If
        End If
    End If
End If
End Sub
Private Sub Frm85_SM_Print_Barcode_buyback_Click()
'on error resume next
DATA_FOUND = 0

frm85_LM_No_ID = vbNullString

If IsNumeric(Frm85.LV3.SelectedItem.Index) Then
    
    frm85_LM_No_ID = Frm85.LV3.ListItems(Frm85.LV3.SelectedItem.Index)
    
    If frm85_LM_No_ID <> vbNullString Then
    
        Note = "Cetak Barcode Item Ini. Teruskan ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from Data_Database where ID='" & frm85_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!NoRujukanSistem) Then
                    GM_No_RUJUKAN_BELIAN = rs!NoRujukanSistem 'No. Rujukan Belian
                    DATA_FOUND = 1
                Else
                    MsgBox "Tiada Data Rujukan Belian.", vbExclamation, "Error"
                End If
                If Not IsNull(rs!cawangan) Then G_KEDAI = rs!cawangan
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_FOUND = 1 Then
                G_FIELD = "NoRujukanSistem"
                Call Print_All_Barcode2
            End If
        End If
    End If
End If
End Sub
Private Sub Frm85_SM_Print_Barcode_buyback_gb_Click()
'on error resume next
DATA_FOUND = 0

If Frm85.MSFlexGrid9 <> vbNullString Then
    Frm85_LM_ID = Frm85.MSFlexGrid9.TextMatrix(Frm85.MSFlexGrid9, 2) 'No. ID
    
    If Frm85_LM_ID <> vbNullString Then
        Note = "Cetak Barcode Item Ini. Teruskan ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from Data_Database where ID='" & Frm85_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!NoRujukanSistem) Then
                    GM_No_RUJUKAN_BELIAN = rs!NoRujukanSistem 'No. Rujukan Belian
                    DATA_FOUND = 1
                Else
                    MsgBox "Tiada Data Rujukan Belian.", vbExclamation, "Error"
                End If
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_FOUND = 1 Then
                Call cetak_barcode_gb_all
            End If
        End If
    End If
End If
End Sub
Private Sub Frm85_SM_Print_Barcode2_buyback_Click()
'on error resume next
DATA_FOUND = 0

frm85_LM_No_ID = vbNullString

If IsNumeric(Frm85.LV3.SelectedItem.Index) Then
    
    frm85_LM_No_ID = Frm85.LV3.ListItems(Frm85.LV3.SelectedItem.Index)
    
    If frm85_LM_No_ID <> vbNullString Then
    
        Note = "Cetak Barcode Item Ini. Teruskan ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from Data_Database where ID='" & frm85_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!no_siri_Produk) Then
                    GM_No_RUJUKAN_BELIAN = rs!no_siri_Produk 'No. Siri Produk
                    DATA_FOUND = 1
                Else
                    MsgBox "Tiada Data Rujukan Belian.", vbExclamation, "Error"
                End If
                If Not IsNull(rs!cawangan) Then G_KEDAI = rs!cawangan
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_FOUND = 1 Then
                G_FIELD = "no_siri_produk"
                Call Print_All_Barcode2
            End If
        End If
    End If
End If
End Sub
Private Sub Frm85_SM_Print_Barcode2_buyback_gb_Click()
'on error resume next
DATA_FOUND = 0

If Frm85.MSFlexGrid9 <> vbNullString Then
    Frm85_LM_ID = Frm85.MSFlexGrid9.TextMatrix(Frm85.MSFlexGrid9, 2) 'No. ID
    
    If Frm85_LM_ID <> vbNullString Then
        Note = "Cetak Barcode Item Ini. Teruskan ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from Data_Database where ID='" & Frm85_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!no_siri_Produk) Then
                    GM_No_RUJUKAN_BELIAN = rs!no_siri_Produk 'No. Siri Produk
                    DATA_FOUND = 1
                Else
                    MsgBox "Tiada Data Rujukan Belian.", vbExclamation, "Error"
                End If
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_FOUND = 1 Then
                Call cetak_barcode_gb
            End If
        End If
    End If
End If
End Sub
Private Sub Frm85_SM_Print_Barcode2_Click()
'on error resume next
DATA_FOUND = 0

frm85_LM_No_ID = vbNullString

If IsNumeric(Frm85.LV2.SelectedItem.Index) Then
    
    frm85_LM_No_ID = Frm85.LV2.ListItems(Frm85.LV2.SelectedItem.Index)
    
    If frm85_LM_No_ID <> vbNullString Then
    
        Note = "Cetak Barcode Item Ini. Teruskan ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from Data_Database where ID='" & frm85_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!no_siri_Produk) Then
                    GM_No_RUJUKAN_BELIAN = rs!no_siri_Produk 'No. Siri Produk
                    DATA_FOUND = 1
                Else
                    MsgBox "Tiada Data Rujukan Belian.", vbExclamation, "Error"
                End If
                If Not IsNull(rs!cawangan) Then G_KEDAI = rs!cawangan
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_FOUND = 1 Then
                G_FIELD = "no_siri_produk"
                Call Print_All_Barcode2
            End If
        End If
    End If
End If
End Sub
Private Sub Frm85_SM_Print_Barcode3_Click()
'on error resume next
DATA_FOUND = 0

If Frm85.MSFlexGrid8 <> vbNullString Then
    Frm85_LM_ID = Frm85.MSFlexGrid8.TextMatrix(Frm85.MSFlexGrid8, 2) 'No. ID
    
    If Frm85_LM_ID <> vbNullString Then
        Note = "Cetak Barcode Item Ini. Teruskan ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from Data_Database where ID='" & Frm85_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!NoRujukanSistem) Then
                    GM_No_RUJUKAN_BELIAN = rs!NoRujukanSistem 'No. Rujukan Belian
                    DATA_FOUND = 1
                Else
                    MsgBox "Tiada Data Rujukan Belian.", vbExclamation, "Error"
                End If
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_FOUND = 1 Then
                Call cetak_barcode_gb_all
            End If
        End If
    End If
End If
End Sub
Private Sub Frm85_SM_Print_Barcode4_Click()
'on error resume next
DATA_FOUND = 0

If Frm85.MSFlexGrid8 <> vbNullString Then
    Frm85_LM_ID = Frm85.MSFlexGrid8.TextMatrix(Frm85.MSFlexGrid8, 2) 'No. ID
    
    If Frm85_LM_ID <> vbNullString Then
        Note = "Cetak Barcode Item Ini. Teruskan ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from Data_Database where ID='" & Frm85_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!no_siri_Produk) Then
                    GM_No_RUJUKAN_BELIAN = rs!no_siri_Produk 'No. Siri Produk
                    DATA_FOUND = 1
                Else
                    MsgBox "Tiada Data Rujukan Belian.", vbExclamation, "Error"
                End If
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_FOUND = 1 Then
                Call cetak_barcode_gb
            End If
        End If
    End If
End If
End Sub
Private Sub Frm85_SM_susut_berat_Click()
'on error resume next
frm85_LM_No_ID = vbNullString
DATA_FOUND = 0

If IsNumeric(Frm85.LV5.SelectedItem.Index) Then
    
    frm85_LM_No_ID = Frm85.LV5.ListItems(Frm85.LV5.SelectedItem.Index)
    
    If frm85_LM_No_ID <> vbNullString Then
    
    
        Frm85.L89_Text = vbNullString
        Frm85.L90_Text = vbNullString
        Frm85.L91_Text = vbNullString
        Frm85.L92_Text = vbNullString
        Frm85.TB1 = "0.00"
        
        If frm85_LM_No_ID <> vbNullString Then
            
            Note = "Adakah anda ingin ubah data susut nilai barang ini?" & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Teruskan?"
            
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbYes Then
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from Data_Database where ID='" & frm85_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                
                    Frm85.L4_Text = frm85_LM_No_ID
                    If Not IsNull(rs!no_siri_Produk) Then Frm85.L89_Text = rs!no_siri_Produk
                    If Not IsNull(rs!kategori_Produk) Then Frm85.L90_Text = rs!kategori_Produk
                    If Not IsNull(rs!Berat) Then Frm85.L91_Text = Format(rs!Berat, "#,##0.00")
                    If Not IsNull(rs!beza_berat) Then Frm85.L92_Text = Format(rs!beza_berat, "#,##0.00")
                    
                    If Not IsNull(rs!susut_berat) Then Frm85.TB1 = Format(rs!susut_berat, "0.00")
                    DATA_FOUND = 1
                    
                End If

                rs.Close
                Set rs = Nothing
                
                If DATA_FOUND = 1 Then
                    
                    Frm85.Pic1.Visible = True
                    
                Else
                
                    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
                
                End If
            End If
            
        End If
        
    Else
    
        MsgBox "Tiada data.", vbInformation, "Info"
        
    End If
    
Else

    MsgBox "Tiada data.", vbInformation, "Info"
    
End If
End Sub
Private Sub L3_Text_Click()
'on error resume next
Frm101.Show
Unload Frm85
End Sub
Private Sub L4_Text_Click()
'On Error Resume Next
Frm34.Show
Unload Frm85
Unload Frm101
End Sub


Private Sub LV1_DblClick()
'on error resume next
frm85_LM_No_ID = vbNullString

If IsNumeric(Frm85.LV1.SelectedItem.Index) Then
    
    frm85_LM_No_ID = Frm85.LV1.SelectedItem.Index
    
    If frm85_LM_No_ID <> vbNullString Then
    
        user_level = MDI_frm1.L4_Text
    
        If user_level = "Admin" Or user_level = "HQ" Or user_level = "Developer" Then
        
            Frm85.Frm85_SM_Edit_Data_Belian.Enabled = True
            Frm85.Frm85_SM_edit_supplier.Enabled = True
            Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = True
            Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = True
            Frm85.Frm85_SM_Padam_Data2.Enabled = True
            Frm85.Frm85_SM_Padam_Data.Enabled = True
            Frm85.Frm85_SM_Padam_Data3.Enabled = True
            Frm85.Frm85_SM_edit_supplier3.Enabled = True
                    
        ElseIf user_level = "Manager" Then
        
            Frm85.Frm85_SM_Edit_Data_Belian.Enabled = True
            Frm85.Frm85_SM_edit_supplier.Enabled = True
            Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = True
            Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = True
            Frm85.Frm85_SM_Padam_Data2.Enabled = False
            Frm85.Frm85_SM_Padam_Data.Enabled = False
            Frm85.Frm85_SM_edit_supplier3.Enabled = True
            Frm85.Frm85_SM_Padam_Data3.Enabled = False
            
        Else
        
            Frm85.Frm85_SM_Edit_Data_Belian.Enabled = False
            Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = False
            Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = False
            Frm85.Frm85_SM_Padam_Data2.Enabled = False
            Frm85.Frm85_SM_Padam_Data.Enabled = False
            Frm85.Frm85_SM_Padam_Data3.Enabled = False
            Frm85.Frm85_SM_edit_supplier.Enabled = False
            Frm85.Frm85_SM_edit_supplier3.Enabled = False
        
        End If
        
        If G_SPKE_ME_MAIL = "YES" Then
            Frm85.frm85_sm_email_jualan.Visible = True
        Else
            Frm85.frm85_sm_email_jualan.Visible = False
        End If
    
        PopupMenu Frm85_PM_Menu2
    
    Else
    
        MsgBox "Tiada Data.", vbInformation, "Info"
        
    End If
    
Else

    MsgBox "Tiada Data.", vbInformation, "Info"
    
End If
End Sub
Private Sub LV2_DblClick()
'on error resume next

frm85_LM_No_ID = vbNullString

If IsNumeric(Frm85.LV2.SelectedItem.Index) Then
    
    frm85_LM_No_ID = Frm85.LV2.SelectedItem.Index
    
    If frm85_LM_No_ID <> vbNullString Then
    
        user_level = MDI_frm1.L4_Text
        
        If user_level = "Admin" Or user_level = "HQ" Or user_level = "Developer" Then
        
            Frm85.Frm85_SM_Edit_Data_Belian.Enabled = True
            Frm85.Frm85_SM_edit_supplier.Enabled = True
            Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = True
            Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = True
            Frm85.Frm85_SM_Padam_Data2.Enabled = True
            Frm85.Frm85_SM_Padam_Data.Enabled = True
            Frm85.Frm85_SM_Padam_Data3.Enabled = True
            Frm85.Frm85_SM_edit_supplier3.Enabled = True
                    
        ElseIf user_level = "Manager" Then
        
            Frm85.Frm85_SM_Edit_Data_Belian.Enabled = True
            Frm85.Frm85_SM_edit_supplier.Enabled = True
            Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = True
            Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = True
            Frm85.Frm85_SM_Padam_Data2.Enabled = False
            Frm85.Frm85_SM_Padam_Data.Enabled = False
            Frm85.Frm85_SM_edit_supplier3.Enabled = True
            Frm85.Frm85_SM_Padam_Data3.Enabled = False
            
        Else
        
            Frm85.Frm85_SM_Edit_Data_Belian.Enabled = False
            Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = False
            Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = False
            Frm85.Frm85_SM_Padam_Data2.Enabled = False
            Frm85.Frm85_SM_Padam_Data.Enabled = False
            Frm85.Frm85_SM_Padam_Data3.Enabled = False
            Frm85.Frm85_SM_edit_supplier.Enabled = False
            Frm85.Frm85_SM_edit_supplier3.Enabled = False
        
        End If
    
        PopupMenu Frm85_PM_Menu1
    
    End If
    
Else

    MsgBox "Tiada Data.", vbInformation, "Info"
    
End If
End Sub
Private Sub LV3_DblClick()
'On Error Resume Next
frm85_LM_No_ID = vbNullString

If IsNumeric(Frm85.LV3.SelectedItem.Index) Then
    
    frm85_LM_No_ID = Frm85.LV3.ListItems(Frm85.LV3.SelectedItem.Index)
    
    If frm85_LM_No_ID <> vbNullString Then

        user_level = MDI_frm1.L4_Text
    
        If user_level = "Admin" Or user_level = "HQ" Or user_level = "Developer" Then
        
            Frm85.Frm85_SM_Edit_Data_Belian.Enabled = True
            Frm85.Frm85_SM_edit_supplier.Enabled = True
            Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = True
            Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = True
            Frm85.Frm85_SM_Padam_Data2.Enabled = True
            Frm85.Frm85_SM_Padam_Data.Enabled = True
            Frm85.Frm85_SM_Padam_Data3.Enabled = True
            Frm85.Frm85_SM_edit_supplier3.Enabled = True
                    
        ElseIf user_level = "Manager" Then
        
            Frm85.Frm85_SM_Edit_Data_Belian.Enabled = True
            Frm85.Frm85_SM_edit_supplier.Enabled = True
            Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = True
            Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = True
            Frm85.Frm85_SM_Padam_Data2.Enabled = False
            Frm85.Frm85_SM_Padam_Data.Enabled = False
            Frm85.Frm85_SM_edit_supplier3.Enabled = True
            Frm85.Frm85_SM_Padam_Data3.Enabled = False
            
        Else
        
            Frm85.Frm85_SM_Edit_Data_Belian.Enabled = False
            Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = False
            Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = False
            Frm85.Frm85_SM_Padam_Data2.Enabled = False
            Frm85.Frm85_SM_Padam_Data.Enabled = False
            Frm85.Frm85_SM_Padam_Data3.Enabled = False
            Frm85.Frm85_SM_edit_supplier.Enabled = False
            Frm85.Frm85_SM_edit_supplier3.Enabled = False
        
        End If

        PopupMenu Frm85_PM_Menu3
        
    End If
    
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub LV4_DblClick()
'On Error Resume Next
frm85_LM_No_ID = vbNullString

If IsNumeric(Frm85.LV4.SelectedItem.Index) Then
    
    frm85_LM_No_ID = Frm85.LV4.ListItems(Frm85.LV4.SelectedItem.Index)
    
    If frm85_LM_No_ID <> vbNullString Then
    
        user_level = MDI_frm1.L4_Text
    
        If user_level = "Admin" Or user_level = "HQ" Or user_level = "Developer" Then
        
            Frm85.frm85_sm_hilang.Enabled = True
                    
        ElseIf user_level = "Manager" Then
        
            Frm85.frm85_sm_hilang.Enabled = False
            
        Else
        
            Frm85.frm85_sm_hilang.Enabled = False
        
        End If
    
        PopupMenu Frm85_PM_Menu4
        
    Else
        
        MsgBox "Tiada data.", vbExclamation, "Info"
        
    End If
    
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub





Private Sub LV5_DblClick()
'on error resume next
frm85_LM_No_ID = vbNullString

If IsNumeric(Frm85.LV5.SelectedItem.Index) Then
    
    frm85_LM_No_ID = Frm85.LV5.SelectedItem.Index
    
    If frm85_LM_No_ID <> vbNullString Then
    
        user_level = MDI_frm1.L4_Text
    
        If user_level = "Admin" Or user_level = "HQ" Or user_level = "Developer" Then
        
            Frm85.Frm85_SM_Edit_Data_Belian.Enabled = True
            Frm85.Frm85_SM_edit_supplier.Enabled = True
            Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = True
            Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = True
            Frm85.Frm85_SM_Padam_Data2.Enabled = True
            Frm85.Frm85_SM_Padam_Data.Enabled = True
            Frm85.Frm85_SM_Padam_Data3.Enabled = True
            Frm85.Frm85_SM_edit_supplier3.Enabled = True
            Frm85.Frm85_SM_susut_berat.Enabled = True
                    
        ElseIf user_level = "Manager" Then
        
            Frm85.Frm85_SM_Edit_Data_Belian.Enabled = True
            Frm85.Frm85_SM_edit_supplier.Enabled = True
            Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = True
            Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = True
            Frm85.Frm85_SM_Padam_Data2.Enabled = False
            Frm85.Frm85_SM_Padam_Data.Enabled = False
            Frm85.Frm85_SM_edit_supplier3.Enabled = True
            Frm85.Frm85_SM_Padam_Data3.Enabled = False
            Frm85.Frm85_SM_susut_berat.Enabled = True
            
        Else
        
            Frm85.Frm85_SM_Edit_Data_Belian.Enabled = False
            Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = False
            Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = False
            Frm85.Frm85_SM_Padam_Data2.Enabled = False
            Frm85.Frm85_SM_Padam_Data.Enabled = False
            Frm85.Frm85_SM_Padam_Data3.Enabled = False
            Frm85.Frm85_SM_edit_supplier.Enabled = False
            Frm85.Frm85_SM_edit_supplier3.Enabled = False
            Frm85.Frm85_SM_susut_berat.Enabled = False
            
        End If
    
        PopupMenu Frm85_PM_Menu5
    
    Else
    
        MsgBox "Tiada Data.", vbInformation, "Info"
        
    End If
    
Else

    MsgBox "Tiada Data.", vbInformation, "Info"
    
End If
End Sub
Private Sub LV6_DblClick()
'On Error Resume Next
frm85_LM_No_ID = vbNullString

If IsNumeric(Frm85.LV6.SelectedItem.Index) Then
    
    frm85_LM_No_ID = Frm85.LV6.ListItems(Frm85.LV6.SelectedItem.Index)
    
    If frm85_LM_No_ID <> vbNullString Then

        user_level = MDI_frm1.L4_Text
    
        If user_level = "Admin" Or user_level = "HQ" Or user_level = "Developer" Then
        
            Frm85.Frm85_SM_Edit_Data_Belian.Enabled = True
            Frm85.Frm85_SM_edit_supplier.Enabled = True
            Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = True
            Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = True
            Frm85.Frm85_SM_Padam_Data2.Enabled = True
            Frm85.Frm85_SM_Padam_Data.Enabled = True
            Frm85.Frm85_SM_Padam_Data3.Enabled = True
            Frm85.Frm85_SM_edit_supplier3.Enabled = True
                    
        ElseIf user_level = "Manager" Then
        
            Frm85.Frm85_SM_Edit_Data_Belian.Enabled = True
            Frm85.Frm85_SM_edit_supplier.Enabled = True
            Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = True
            Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = True
            Frm85.Frm85_SM_Padam_Data2.Enabled = False
            Frm85.Frm85_SM_Padam_Data.Enabled = False
            Frm85.Frm85_SM_edit_supplier3.Enabled = True
            Frm85.Frm85_SM_Padam_Data3.Enabled = False
            
        Else
        
            Frm85.Frm85_SM_Edit_Data_Belian.Enabled = False
            Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = False
            Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = False
            Frm85.Frm85_SM_Padam_Data2.Enabled = False
            Frm85.Frm85_SM_Padam_Data.Enabled = False
            Frm85.Frm85_SM_Padam_Data3.Enabled = False
            Frm85.Frm85_SM_edit_supplier.Enabled = False
            Frm85.Frm85_SM_edit_supplier3.Enabled = False
        
        End If
    
        PopupMenu Frm85_PM_Menu9
        
    End If
    
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub MSFlexGrid10_DblClick()
'On Error Resume Next
If Frm85.MSFlexGrid10 <> vbNullString Then

user_level = MDI_frm1.L4_Text

    If user_level = "Admin" Or user_level = "HQ" Or user_level = "Developer" Then
    
        Frm85.Frm85_SM_Edit_Data_Belian.Enabled = True
        Frm85.Frm85_SM_edit_supplier.Enabled = True
        Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = True
        Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = True
        Frm85.Frm85_SM_Padam_Data2.Enabled = True
        Frm85.Frm85_SM_Padam_Data.Enabled = True
        Frm85.Frm85_SM_Padam_Data3.Enabled = True
        Frm85.Frm85_SM_edit_supplier3.Enabled = True
                
    ElseIf user_level = "Manager" Then
    
        Frm85.Frm85_SM_Edit_Data_Belian.Enabled = True
        Frm85.Frm85_SM_edit_supplier.Enabled = True
        Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = True
        Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = True
        Frm85.Frm85_SM_Padam_Data2.Enabled = False
        Frm85.Frm85_SM_Padam_Data.Enabled = False
        Frm85.Frm85_SM_edit_supplier3.Enabled = True
        Frm85.Frm85_SM_Padam_Data3.Enabled = False
        
    Else
    
        Frm85.Frm85_SM_Edit_Data_Belian.Enabled = False
        Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = False
        Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = False
        Frm85.Frm85_SM_Padam_Data2.Enabled = False
        Frm85.Frm85_SM_Padam_Data.Enabled = False
        Frm85.Frm85_SM_Padam_Data3.Enabled = False
        Frm85.Frm85_SM_edit_supplier.Enabled = False
        Frm85.Frm85_SM_edit_supplier3.Enabled = False
    
    End If

    'PopupMenu Frm85_PM_Menu10
    PopupMenu Frm85_PM_Menu11
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub



Private Sub MSFlexGrid6_DblClick()
'On Error Resume Next
If Frm85.MSFlexGrid6 <> vbNullString Then

user_level = MDI_frm1.L4_Text

    If user_level = "Admin" Or user_level = "HQ" Or user_level = "Developer" Then
    
        Frm85.Frm85_SM_Edit_Data_Belian.Enabled = True
        Frm85.Frm85_SM_edit_supplier.Enabled = True
        Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = True
        Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = True
        Frm85.Frm85_SM_Padam_Data2.Enabled = True
        Frm85.Frm85_SM_Padam_Data.Enabled = True
        Frm85.Frm85_SM_Padam_Data3.Enabled = True
        Frm85.Frm85_SM_edit_supplier3.Enabled = True
                
    ElseIf user_level = "Manager" Then
    
        Frm85.Frm85_SM_Edit_Data_Belian.Enabled = True
        Frm85.Frm85_SM_edit_supplier.Enabled = True
        Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = True
        Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = True
        Frm85.Frm85_SM_Padam_Data2.Enabled = False
        Frm85.Frm85_SM_Padam_Data.Enabled = False
        Frm85.Frm85_SM_edit_supplier3.Enabled = True
        Frm85.Frm85_SM_Padam_Data3.Enabled = False
        
    Else
    
        Frm85.Frm85_SM_Edit_Data_Belian.Enabled = False
        Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = False
        Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = False
        Frm85.Frm85_SM_Padam_Data2.Enabled = False
        Frm85.Frm85_SM_Padam_Data.Enabled = False
        Frm85.Frm85_SM_Padam_Data3.Enabled = False
        Frm85.Frm85_SM_edit_supplier.Enabled = False
        Frm85.Frm85_SM_edit_supplier3.Enabled = False
    
    End If

    PopupMenu Frm85_PM_Menu8
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub MSFlexGrid8_DblClick()
'On Error Resume Next
If Frm85.MSFlexGrid8 <> vbNullString Then

user_level = MDI_frm1.L4_Text

    If user_level = "Admin" Or user_level = "HQ" Or user_level = "Developer" Then
    
        Frm85.Frm85_SM_Edit_Data_Belian.Enabled = True
        Frm85.Frm85_SM_edit_supplier.Enabled = True
        Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = True
        Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = True
        Frm85.Frm85_SM_Padam_Data2.Enabled = True
        Frm85.Frm85_SM_Padam_Data.Enabled = True
        Frm85.Frm85_SM_Padam_Data3.Enabled = True
        Frm85.Frm85_SM_edit_supplier3.Enabled = True
                
    ElseIf user_level = "Manager" Then
    
        Frm85.Frm85_SM_Edit_Data_Belian.Enabled = True
        Frm85.Frm85_SM_edit_supplier.Enabled = True
        Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = True
        Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = True
        Frm85.Frm85_SM_Padam_Data2.Enabled = False
        Frm85.Frm85_SM_Padam_Data.Enabled = False
        Frm85.Frm85_SM_edit_supplier3.Enabled = True
        Frm85.Frm85_SM_Padam_Data3.Enabled = False
        
    Else
    
        Frm85.Frm85_SM_Edit_Data_Belian.Enabled = False
        Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = False
        Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = False
        Frm85.Frm85_SM_Padam_Data2.Enabled = False
        Frm85.Frm85_SM_Padam_Data.Enabled = False
        Frm85.Frm85_SM_Padam_Data3.Enabled = False
        Frm85.Frm85_SM_edit_supplier.Enabled = False
        Frm85.Frm85_SM_edit_supplier3.Enabled = False
    
    End If

    PopupMenu Frm85_PM_Menu6
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub MSFlexGrid9_DblClick()
'On Error Resume Next
If Frm85.MSFlexGrid9 <> vbNullString Then

user_level = MDI_frm1.L4_Text

    If user_level = "Admin" Or user_level = "HQ" Or user_level = "Developer" Then
    
        Frm85.Frm85_SM_Edit_Data_Belian.Enabled = True
        Frm85.Frm85_SM_edit_supplier.Enabled = True
        Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = True
        Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = True
        Frm85.Frm85_SM_Padam_Data2.Enabled = True
        Frm85.Frm85_SM_Padam_Data.Enabled = True
        Frm85.Frm85_SM_Padam_Data3.Enabled = True
        Frm85.Frm85_SM_edit_supplier3.Enabled = True
                
    ElseIf user_level = "Manager" Then
    
        Frm85.Frm85_SM_Edit_Data_Belian.Enabled = True
        Frm85.Frm85_SM_edit_supplier.Enabled = True
        Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = True
        Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = True
        Frm85.Frm85_SM_Padam_Data2.Enabled = False
        Frm85.Frm85_SM_Padam_Data.Enabled = False
        Frm85.Frm85_SM_edit_supplier3.Enabled = True
        Frm85.Frm85_SM_Padam_Data3.Enabled = False
        
    Else
    
        Frm85.Frm85_SM_Edit_Data_Belian.Enabled = False
        Frm85.Frm85_SM_Edit_Data_Jualan.Enabled = False
        Frm85.Frm85_SM_Edit_Data_Buyback.Enabled = False
        Frm85.Frm85_SM_Padam_Data2.Enabled = False
        Frm85.Frm85_SM_Padam_Data.Enabled = False
        Frm85.Frm85_SM_Padam_Data3.Enabled = False
        Frm85.Frm85_SM_edit_supplier.Enabled = False
        Frm85.Frm85_SM_edit_supplier3.Enabled = False
    
    End If

    PopupMenu Frm85_PM_Menu7
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub Tmr1_Timer()
'on error resume next
Frm85.L1_Text = DateTime.Date
Frm85.L2_Text = DateTime.Time$
End Sub
