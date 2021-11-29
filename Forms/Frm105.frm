VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm105 
   Caption         =   "Report Kewangan"
   ClientHeight    =   13035
   ClientLeft      =   120
   ClientTop       =   -7635
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
   Icon            =   "Frm105.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   13035
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic9 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12015
      Left            =   2280
      ScaleHeight     =   12015
      ScaleWidth      =   23535
      TabIndex        =   112
      Top             =   1800
      Visible         =   0   'False
      Width           =   23535
      Begin VB.PictureBox Pic10 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   11415
         Left            =   360
         ScaleHeight     =   11415
         ScaleWidth      =   23535
         TabIndex        =   114
         Top             =   360
         Visible         =   0   'False
         Width           =   23535
         Begin VB.CommandButton CMD16 
            BackColor       =   &H00FFFFFF&
            Height          =   840
            Left            =   11640
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm105.frx":0ECA
            MousePointer    =   99  'Custom
            Picture         =   "Frm105.frx":11D4
            Style           =   1  'Graphical
            TabIndex        =   179
            ToolTipText     =   "Paparan seterusnya"
            Top             =   9720
            Width           =   1200
         End
         Begin VB.CommandButton CMD15 
            BackColor       =   &H00FFFFFF&
            Height          =   840
            Left            =   10320
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm105.frx":1AFA
            MousePointer    =   99  'Custom
            Picture         =   "Frm105.frx":1E04
            Style           =   1  'Graphical
            TabIndex        =   178
            ToolTipText     =   "Paparan sebelumnya"
            Top             =   9720
            Width           =   1200
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid7 
            Height          =   10305
            Left            =   240
            TabIndex        =   115
            ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
            Top             =   480
            Width           =   9885
            _ExtentX        =   17436
            _ExtentY        =   18177
            _Version        =   393216
            Rows            =   1
            Cols            =   0
            FixedCols       =   0
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   8454016
            BackColorSel    =   -2147483643
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
         Begin VB.Label L113_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L113_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   12480
            TabIndex        =   204
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label L112_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L112_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   12480
            TabIndex        =   203
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   ": RM  : RM  : RM "
            Height          =   1815
            Left            =   11985
            TabIndex        =   202
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah                    Tunai                   Bank In               "
            Height          =   1575
            Left            =   10320
            TabIndex        =   201
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label L72_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L72_Text"
            Height          =   255
            Left            =   360
            TabIndex        =   123
            Top             =   120
            Width           =   15375
         End
         Begin VB.Label Label40 
            BackStyle       =   0  'Transparent
            Caption         =   "** Belian barang trade in dari pelanggan adalah secara TUNAI / BANK IN."
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
            Left            =   10320
            TabIndex        =   122
            Top             =   1560
            Width           =   5895
         End
         Begin VB.Label L73_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L73_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   12480
            TabIndex        =   121
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label L74_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "L74_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   11520
            TabIndex        =   119
            Top             =   10560
            Width           =   735
         End
         Begin VB.Label L75_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L75_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   12480
            TabIndex        =   118
            Top             =   10560
            Width           =   2295
         End
         Begin VB.Label L77_Text 
            Caption         =   "L77_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   17640
            TabIndex        =   117
            Top             =   5640
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label L76_Text 
            Caption         =   "L76_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   17640
            TabIndex        =   116
            Top             =   5160
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Paparan Muka  :          / "
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   10320
            TabIndex        =   120
            Top             =   10560
            Width           =   2295
         End
      End
      Begin VB.PictureBox Pic11 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   11415
         Left            =   5040
         ScaleHeight     =   11415
         ScaleWidth      =   23535
         TabIndex        =   184
         Top             =   2520
         Visible         =   0   'False
         Width           =   23535
         Begin VB.CommandButton CMD17 
            BackColor       =   &H00FFFFFF&
            Height          =   840
            Left            =   7080
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm105.frx":2743
            MousePointer    =   99  'Custom
            Picture         =   "Frm105.frx":2A4D
            Style           =   1  'Graphical
            TabIndex        =   196
            ToolTipText     =   "Paparan sebelumnya"
            Top             =   9960
            Width           =   1200
         End
         Begin VB.CommandButton CMD18 
            BackColor       =   &H00FFFFFF&
            Height          =   840
            Left            =   8400
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm105.frx":338C
            MousePointer    =   99  'Custom
            Picture         =   "Frm105.frx":3696
            Style           =   1  'Graphical
            TabIndex        =   195
            ToolTipText     =   "Paparan seterusnya"
            Top             =   9960
            Width           =   1200
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid8 
            Height          =   10545
            Left            =   240
            TabIndex        =   185
            ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
            Top             =   480
            Width           =   6765
            _ExtentX        =   11933
            _ExtentY        =   18600
            _Version        =   393216
            Rows            =   1
            Cols            =   0
            FixedCols       =   0
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   8454016
            BackColorSel    =   -2147483643
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
         Begin VB.Label L79_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L79_Text"
            Height          =   255
            Left            =   360
            TabIndex        =   193
            Top             =   120
            Width           =   15375
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah    : RM"
            Height          =   255
            Left            =   7200
            TabIndex        =   192
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "** Belian barang dari agen ini adalah secara TUNAI."
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
            Left            =   7200
            TabIndex        =   191
            Top             =   840
            Width           =   6615
         End
         Begin VB.Label L80_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L80_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   8520
            TabIndex        =   190
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label L81_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "L81_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   8280
            TabIndex        =   189
            Top             =   10800
            Width           =   735
         End
         Begin VB.Label L82_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L82_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   9240
            TabIndex        =   188
            Top             =   10800
            Width           =   2295
         End
         Begin VB.Label L84_Text 
            Caption         =   "L84_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   17640
            TabIndex        =   187
            Top             =   5640
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label L83_Text 
            Caption         =   "L83_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   17640
            TabIndex        =   186
            Top             =   5160
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "Paparan Muka  :          / "
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   7080
            TabIndex        =   194
            Top             =   10800
            Width           =   2295
         End
      End
      Begin VB.PictureBox Pic14 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   11415
         Left            =   12000
         ScaleHeight     =   11415
         ScaleWidth      =   23535
         TabIndex        =   148
         Top             =   2160
         Visible         =   0   'False
         Width           =   23535
         Begin VB.CommandButton CMD24 
            BackColor       =   &H00FFFFFF&
            Height          =   840
            Left            =   12000
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm105.frx":3FBC
            MousePointer    =   99  'Custom
            Picture         =   "Frm105.frx":42C6
            Style           =   1  'Graphical
            TabIndex        =   181
            ToolTipText     =   "Paparan seterusnya"
            Top             =   9840
            Width           =   1200
         End
         Begin VB.CommandButton CMD23 
            BackColor       =   &H00FFFFFF&
            Height          =   840
            Left            =   10680
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm105.frx":4BEC
            MousePointer    =   99  'Custom
            Picture         =   "Frm105.frx":4EF6
            Style           =   1  'Graphical
            TabIndex        =   180
            ToolTipText     =   "Paparan sebelumnya"
            Top             =   9840
            Width           =   1200
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid11 
            Height          =   10545
            Left            =   240
            TabIndex        =   149
            ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
            Top             =   360
            Width           =   10365
            _ExtentX        =   18283
            _ExtentY        =   18600
            _Version        =   393216
            Rows            =   1
            Cols            =   0
            FixedCols       =   0
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   8454016
            BackColorSel    =   -2147483643
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
         Begin VB.Label L103_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L103_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   12120
            TabIndex        =   161
            Top             =   840
            Width           =   2295
         End
         Begin VB.Label Label43 
            BackStyle       =   0  'Transparent
            Caption         =   ": RM  : RM  : RM "
            Height          =   1815
            Left            =   11625
            TabIndex        =   159
            Top             =   360
            Width           =   495
         End
         Begin VB.Label L101_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L101_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   12120
            TabIndex        =   158
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label L102_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L102_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   12120
            TabIndex        =   157
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label L100_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L100_Text"
            Height          =   255
            Left            =   360
            TabIndex        =   154
            Top             =   120
            Width           =   15375
         End
         Begin VB.Label L104_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "L104_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   11880
            TabIndex        =   153
            Top             =   10680
            Width           =   735
         End
         Begin VB.Label L105_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L105_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   12840
            TabIndex        =   152
            Top             =   10680
            Width           =   2295
         End
         Begin VB.Label L107_Text 
            Caption         =   "L107_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   17640
            TabIndex        =   151
            Top             =   5640
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label L106_Text 
            Caption         =   "L106_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   17640
            TabIndex        =   150
            Top             =   5160
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label Label45 
            BackStyle       =   0  'Transparent
            Caption         =   "Paparan Muka  :          / "
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   10680
            TabIndex        =   155
            Top             =   10680
            Width           =   2295
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah       Tunai         Bank In      "
            Height          =   1575
            Left            =   10680
            TabIndex        =   160
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.PictureBox Pic13 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   11415
         Left            =   6840
         ScaleHeight     =   11415
         ScaleWidth      =   23535
         TabIndex        =   136
         Top             =   720
         Visible         =   0   'False
         Width           =   23535
         Begin VB.CommandButton CMD22 
            BackColor       =   &H00FFFFFF&
            Height          =   840
            Left            =   8400
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm105.frx":5835
            MousePointer    =   99  'Custom
            Picture         =   "Frm105.frx":5B3F
            Style           =   1  'Graphical
            TabIndex        =   177
            ToolTipText     =   "Paparan seterusnya"
            Top             =   9960
            Width           =   1200
         End
         Begin VB.CommandButton CMD21 
            BackColor       =   &H00FFFFFF&
            Height          =   840
            Left            =   7080
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm105.frx":6465
            MousePointer    =   99  'Custom
            Picture         =   "Frm105.frx":676F
            Style           =   1  'Graphical
            TabIndex        =   176
            ToolTipText     =   "Paparan sebelumnya"
            Top             =   9960
            Width           =   1200
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid10 
            Height          =   10545
            Left            =   240
            TabIndex        =   137
            ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
            Top             =   480
            Width           =   6765
            _ExtentX        =   11933
            _ExtentY        =   18600
            _Version        =   393216
            Rows            =   1
            Cols            =   0
            FixedCols       =   0
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   8454016
            BackColorSel    =   -2147483643
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
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "** Semua perbelanjaan adalah secara TUNAI."
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
            Left            =   7200
            TabIndex        =   147
            Top             =   960
            Width           =   6615
         End
         Begin VB.Label L97_Text 
            Caption         =   "L97_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   17640
            TabIndex        =   144
            Top             =   5160
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label L98_Text 
            Caption         =   "L98_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   17640
            TabIndex        =   143
            Top             =   5640
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label L96_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L96_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   9240
            TabIndex        =   142
            Top             =   10800
            Width           =   2295
         End
         Begin VB.Label L95_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "L95_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   8280
            TabIndex        =   141
            Top             =   10800
            Width           =   735
         End
         Begin VB.Label L94_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L94_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   8520
            TabIndex        =   140
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah    : RM"
            Height          =   255
            Left            =   7200
            TabIndex        =   139
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label L93_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L93_Text"
            Height          =   255
            Left            =   360
            TabIndex        =   138
            Top             =   120
            Width           =   15375
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "Paparan Muka  :          / "
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   7080
            TabIndex        =   145
            Top             =   10800
            Width           =   2295
         End
      End
      Begin VB.PictureBox Pic12 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   11415
         Left            =   3720
         ScaleHeight     =   11415
         ScaleWidth      =   23535
         TabIndex        =   125
         Top             =   600
         Visible         =   0   'False
         Width           =   23535
         Begin VB.CommandButton CMD20 
            BackColor       =   &H00FFFFFF&
            Height          =   840
            Left            =   6720
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm105.frx":70AE
            MousePointer    =   99  'Custom
            Picture         =   "Frm105.frx":73B8
            Style           =   1  'Graphical
            TabIndex        =   183
            ToolTipText     =   "Paparan seterusnya"
            Top             =   9960
            Width           =   1200
         End
         Begin VB.CommandButton CMD19 
            BackColor       =   &H00FFFFFF&
            Height          =   840
            Left            =   5400
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm105.frx":7CDE
            MousePointer    =   99  'Custom
            Picture         =   "Frm105.frx":7FE8
            Style           =   1  'Graphical
            TabIndex        =   182
            ToolTipText     =   "Paparan sebelumnya"
            Top             =   9960
            Width           =   1200
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid9 
            Height          =   10545
            Left            =   240
            TabIndex        =   126
            ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
            Top             =   480
            Width           =   5085
            _ExtentX        =   8969
            _ExtentY        =   18600
            _Version        =   393216
            Rows            =   1
            Cols            =   0
            FixedCols       =   0
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   8454016
            BackColorSel    =   -2147483643
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
         Begin VB.Label L86_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L86_Text"
            Height          =   255
            Left            =   360
            TabIndex        =   133
            Top             =   120
            Width           =   15375
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah    : RM"
            Height          =   255
            Left            =   5520
            TabIndex        =   132
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label L87_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L87_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   6840
            TabIndex        =   131
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label L88_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "L88_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   6600
            TabIndex        =   130
            Top             =   10800
            Width           =   735
         End
         Begin VB.Label L89_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L89_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   7560
            TabIndex        =   129
            Top             =   10800
            Width           =   2295
         End
         Begin VB.Label L91_Text 
            Caption         =   "L91_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   17640
            TabIndex        =   128
            Top             =   5640
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label L90_Text 
            Caption         =   "L90_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   17640
            TabIndex        =   127
            Top             =   5160
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "Paparan Muka  :          / "
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   5400
            TabIndex        =   134
            Top             =   10800
            Width           =   2295
         End
      End
      Begin VB.Label L99_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Bayaran gaji"
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
         Left            =   6600
         MouseIcon       =   "Frm105.frx":8927
         MousePointer    =   99  'Custom
         TabIndex        =   156
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label L92_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Perbelanjaan kedai"
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
         Left            =   4560
         MouseIcon       =   "Frm105.frx":8C31
         MousePointer    =   99  'Custom
         TabIndex        =   146
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label L85_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ambilan tunai dari kedai"
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
         Left            =   2160
         MouseIcon       =   "Frm105.frx":8F3B
         MousePointer    =   99  'Custom
         TabIndex        =   135
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label L78_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Belian tukaran barang oleh agen"
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
         Left            =   11520
         MouseIcon       =   "Frm105.frx":9245
         MousePointer    =   99  'Custom
         TabIndex        =   124
         Top             =   120
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label L71_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Belian barang trade in"
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
         MouseIcon       =   "Frm105.frx":954F
         MousePointer    =   99  'Custom
         TabIndex        =   113
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.PictureBox Pic1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2415
      ScaleWidth      =   5865
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   5865
      Begin VB.ComboBox CBB1 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         ItemData        =   "Frm105.frx":9859
         Left            =   1500
         List            =   "Frm105.frx":985B
         Style           =   2  'Dropdown List
         TabIndex        =   197
         Top             =   1250
         Width           =   4005
      End
      Begin VB.CommandButton CMD1 
         BackColor       =   &H000080FF&
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm105.frx":985D
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   1800
         Width           =   2025
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   1500
         TabIndex        =   7
         Top             =   525
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
         Format          =   142475264
         CurrentDate     =   41561
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   360
         Left            =   1500
         TabIndex        =   8
         Top             =   885
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
         Format          =   142475264
         CurrentDate     =   41561
      End
      Begin VB.Label L111_Text 
         Caption         =   "L111_Text"
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
         Left            =   360
         TabIndex        =   200
         Top             =   2160
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label L110_Text 
         Caption         =   "L110_Text"
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
         Left            =   360
         TabIndex        =   199
         Top             =   1920
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pekerja"
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
         Left            =   315
         TabIndex        =   198
         Top             =   1280
         Width           =   1995
      End
      Begin VB.Shape Shape2 
         Height          =   1575
         Left            =   120
         Top             =   120
         Width           =   5655
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila buat pilihan tempoh report kewangan."
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
         Height          =   240
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   8610
      End
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Akhir "
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
         Left            =   315
         TabIndex        =   12
         Top             =   930
         Width           =   1995
      End
      Begin VB.Label Label62 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Mula "
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
         Left            =   315
         TabIndex        =   11
         Top             =   570
         Width           =   1995
      End
      Begin VB.Label L6_Text 
         Caption         =   "L6_Text"
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
         Left            =   4680
         TabIndex        =   10
         Top             =   1800
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label L5_Text 
         Caption         =   "L5_Text"
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
         Left            =   3960
         TabIndex        =   9
         Top             =   1800
         Visible         =   0   'False
         Width           =   660
      End
   End
   Begin VB.Timer Tmr1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Pic2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12015
      Left            =   480
      ScaleHeight     =   12015
      ScaleWidth      =   23535
      TabIndex        =   14
      Top             =   3000
      Visible         =   0   'False
      Width           =   23535
      Begin VB.PictureBox Pic4 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   11415
         Left            =   -840
         ScaleHeight     =   11415
         ScaleWidth      =   23535
         TabIndex        =   35
         Top             =   480
         Visible         =   0   'False
         Width           =   23535
         Begin VB.CommandButton CMD3 
            BackColor       =   &H00FFFFFF&
            Height          =   840
            Left            =   15720
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm105.frx":9B67
            MousePointer    =   99  'Custom
            Picture         =   "Frm105.frx":9E71
            Style           =   1  'Graphical
            TabIndex        =   165
            ToolTipText     =   "Paparan seterusnya"
            Top             =   9720
            Width           =   1200
         End
         Begin VB.CommandButton CMD2 
            BackColor       =   &H00FFFFFF&
            Height          =   840
            Left            =   14400
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm105.frx":A797
            MousePointer    =   99  'Custom
            Picture         =   "Frm105.frx":AAA1
            Style           =   1  'Graphical
            TabIndex        =   164
            ToolTipText     =   "Paparan sebelumnya"
            Top             =   9720
            Width           =   1200
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
            Height          =   10305
            Left            =   240
            TabIndex        =   36
            ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
            Top             =   480
            Width           =   13995
            _ExtentX        =   24686
            _ExtentY        =   18177
            _Version        =   393216
            Rows            =   1
            Cols            =   0
            FixedCols       =   0
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   8454016
            BackColorSel    =   -2147483643
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
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "** Jumlah kutipan dari [Kad Kredit] dan [Kad Debit] adalah tidak termasuk caj perkhidmatan yang dikenakan oleh pihak bank."
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
            Left            =   14400
            TabIndex        =   68
            Top             =   2160
            Width           =   6615
         End
         Begin VB.Label L22_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L22_Text"
            Height          =   255
            Left            =   360
            TabIndex        =   50
            Top             =   120
            Width           =   15375
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   $"Frm105.frx":B3E0
            Height          =   1575
            Left            =   14400
            TabIndex        =   49
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   ": RM  : RM  : RM  : RM  : RM  : RM"
            Height          =   1815
            Left            =   16065
            TabIndex        =   48
            Top             =   480
            Width           =   495
         End
         Begin VB.Label L23_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L23_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   16560
            TabIndex        =   47
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label L28_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L28_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   16560
            TabIndex        =   46
            Top             =   1680
            Width           =   2295
         End
         Begin VB.Label L27_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L27_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   16560
            TabIndex        =   45
            Top             =   1440
            Width           =   2295
         End
         Begin VB.Label L26_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L26_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   16560
            TabIndex        =   44
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label L24_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L24_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   16560
            TabIndex        =   43
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label L25_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L25_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   16560
            TabIndex        =   42
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label L29_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "L29_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   15600
            TabIndex        =   40
            Top             =   10560
            Width           =   735
         End
         Begin VB.Label L30_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L30_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   16560
            TabIndex        =   39
            Top             =   10560
            Width           =   2295
         End
         Begin VB.Label L32_Text 
            Caption         =   "L32_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   17640
            TabIndex        =   38
            Top             =   5640
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label L31_Text 
            Caption         =   "L31_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   17640
            TabIndex        =   37
            Top             =   5160
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Paparan Muka  :          / "
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   14400
            TabIndex        =   41
            Top             =   10560
            Width           =   2295
         End
      End
      Begin VB.PictureBox Pic3 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   11415
         Left            =   5880
         ScaleHeight     =   11415
         ScaleWidth      =   23535
         TabIndex        =   15
         Top             =   -2880
         Visible         =   0   'False
         Width           =   23535
         Begin VB.CommandButton CMD7 
            BackColor       =   &H00FFFFFF&
            Height          =   840
            Left            =   14400
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm105.frx":B46A
            MousePointer    =   99  'Custom
            Picture         =   "Frm105.frx":B774
            Style           =   1  'Graphical
            TabIndex        =   171
            ToolTipText     =   "Paparan sebelumnya"
            Top             =   9720
            Width           =   1200
         End
         Begin VB.CommandButton CMD8 
            BackColor       =   &H00FFFFFF&
            Height          =   840
            Left            =   15720
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm105.frx":C0B3
            MousePointer    =   99  'Custom
            Picture         =   "Frm105.frx":C3BD
            Style           =   1  'Graphical
            TabIndex        =   170
            ToolTipText     =   "Paparan seterusnya"
            Top             =   9720
            Width           =   1200
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   10305
            Left            =   240
            TabIndex        =   17
            ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
            Top             =   480
            Width           =   13995
            _ExtentX        =   24686
            _ExtentY        =   18177
            _Version        =   393216
            Rows            =   1
            Cols            =   0
            FixedCols       =   0
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   8454016
            BackColorSel    =   -2147483643
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
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "** Jumlah kutipan dari [Kad Kredit] dan [Kad Debit] adalah tidak termasuk caj perkhidmatan yang dikenakan oleh pihak bank."
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
            Left            =   14400
            TabIndex        =   33
            Top             =   3240
            Width           =   6015
         End
         Begin VB.Label L19_Text 
            Caption         =   "L19_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   17640
            TabIndex        =   32
            Top             =   5160
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label L20_Text 
            Caption         =   "L20_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   17640
            TabIndex        =   31
            Top             =   5640
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label L18_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L18_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   16560
            TabIndex        =   30
            Top             =   10560
            Width           =   2295
         End
         Begin VB.Label L17_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "L17_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   15600
            TabIndex        =   29
            Top             =   10560
            Width           =   735
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Paparan Muka  :          / "
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   14400
            TabIndex        =   28
            Top             =   10560
            Width           =   2295
         End
         Begin VB.Label L13_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L13_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   16680
            TabIndex        =   27
            Top             =   1080
            Width           =   2295
         End
         Begin VB.Label L12_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L12_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   16680
            TabIndex        =   26
            Top             =   840
            Width           =   2295
         End
         Begin VB.Label L14_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L14_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   16680
            TabIndex        =   25
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label L15_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L15_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   16680
            TabIndex        =   24
            Top             =   1560
            Width           =   2295
         End
         Begin VB.Label L16_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L16_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   16680
            TabIndex        =   23
            Top             =   1800
            Width           =   2295
         End
         Begin VB.Label L11_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L11_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   16680
            TabIndex        =   22
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "      RM  : RM  : RM  : RM  : RM  :"
            Height          =   1815
            Left            =   16185
            TabIndex        =   21
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   $"Frm105.frx":CCE3
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
            Left            =   14400
            TabIndex        =   20
            Top             =   2160
            Width           =   6135
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"Frm105.frx":CD92
            Height          =   1575
            Left            =   14520
            TabIndex        =   19
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label L10_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L10_Text"
            Height          =   255
            Left            =   360
            TabIndex        =   18
            Top             =   120
            Width           =   15375
         End
      End
      Begin VB.PictureBox Pic8 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   11415
         Left            =   9840
         ScaleHeight     =   11415
         ScaleWidth      =   23535
         TabIndex        =   101
         Top             =   -600
         Visible         =   0   'False
         Width           =   23535
         Begin VB.CommandButton CMD13 
            BackColor       =   &H00FFFFFF&
            Height          =   840
            Left            =   6960
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm105.frx":CE1B
            MousePointer    =   99  'Custom
            Picture         =   "Frm105.frx":D125
            Style           =   1  'Graphical
            TabIndex        =   169
            ToolTipText     =   "Paparan sebelumnya"
            Top             =   9840
            Width           =   1200
         End
         Begin VB.CommandButton CMD14 
            BackColor       =   &H00FFFFFF&
            Height          =   840
            Left            =   8280
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm105.frx":DA64
            MousePointer    =   99  'Custom
            Picture         =   "Frm105.frx":DD6E
            Style           =   1  'Graphical
            TabIndex        =   168
            ToolTipText     =   "Paparan seterusnya"
            Top             =   9840
            Width           =   1200
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid6 
            Height          =   10305
            Left            =   240
            TabIndex        =   102
            ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
            Top             =   480
            Width           =   6645
            _ExtentX        =   11721
            _ExtentY        =   18177
            _Version        =   393216
            Rows            =   1
            Cols            =   0
            FixedCols       =   0
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   8454016
            BackColorSel    =   -2147483643
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
         Begin VB.Label L65_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L65_Text"
            Height          =   255
            Left            =   360
            TabIndex        =   110
            Top             =   120
            Width           =   15375
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah        : RM"
            Height          =   255
            Left            =   6960
            TabIndex        =   109
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label L66_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L66_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   8520
            TabIndex        =   108
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label L67_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "L67_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   8160
            TabIndex        =   107
            Top             =   10680
            Width           =   735
         End
         Begin VB.Label L68_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L68_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   9120
            TabIndex        =   106
            Top             =   10680
            Width           =   2295
         End
         Begin VB.Label L70_Text 
            Caption         =   "L70_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   7560
            TabIndex        =   105
            Top             =   6840
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label L69_Text 
            Caption         =   "L69_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   7560
            TabIndex        =   104
            Top             =   6360
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "** Simpanan duit di kedai adalah secara TUNAI."
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
            Left            =   6960
            TabIndex        =   103
            Top             =   840
            Width           =   6615
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "Paparan Muka  :          / "
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   6960
            TabIndex        =   111
            Top             =   10680
            Width           =   2295
         End
      End
      Begin VB.PictureBox Pic7 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   11415
         Left            =   8520
         ScaleHeight     =   11415
         ScaleWidth      =   23535
         TabIndex        =   88
         Top             =   720
         Visible         =   0   'False
         Width           =   23535
         Begin VB.CommandButton CMD11 
            BackColor       =   &H00FFFFFF&
            Height          =   840
            Left            =   6240
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm105.frx":E694
            MousePointer    =   99  'Custom
            Picture         =   "Frm105.frx":E99E
            Style           =   1  'Graphical
            TabIndex        =   167
            ToolTipText     =   "Paparan sebelumnya"
            Top             =   9960
            Width           =   1200
         End
         Begin VB.CommandButton CMD12 
            BackColor       =   &H00FFFFFF&
            Height          =   840
            Left            =   7560
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm105.frx":F2DD
            MousePointer    =   99  'Custom
            Picture         =   "Frm105.frx":F5E7
            Style           =   1  'Graphical
            TabIndex        =   166
            ToolTipText     =   "Paparan seterusnya"
            Top             =   9960
            Width           =   1200
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid5 
            Height          =   10545
            Left            =   240
            TabIndex        =   89
            ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
            Top             =   480
            Width           =   5925
            _ExtentX        =   10451
            _ExtentY        =   18600
            _Version        =   393216
            Rows            =   1
            Cols            =   0
            FixedCols       =   0
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   8454016
            BackColorSel    =   -2147483643
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
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "** Kemasukkan duit ke kedai adalah secara TUNAI."
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
            Left            =   6360
            TabIndex        =   97
            Top             =   960
            Width           =   6615
         End
         Begin VB.Label L62_Text 
            Caption         =   "L62_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   7080
            TabIndex        =   96
            Top             =   5280
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label L63_Text 
            Caption         =   "L63_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   7080
            TabIndex        =   95
            Top             =   5760
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label L61_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L61_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   8400
            TabIndex        =   94
            Top             =   10800
            Width           =   2295
         End
         Begin VB.Label L60_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "L60_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   7440
            TabIndex        =   93
            Top             =   10800
            Width           =   735
         End
         Begin VB.Label L59_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L59_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   7920
            TabIndex        =   92
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah        : RM"
            Height          =   255
            Left            =   6360
            TabIndex        =   91
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label L58_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L58_Text"
            Height          =   255
            Left            =   360
            TabIndex        =   90
            Top             =   120
            Width           =   15375
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Paparan Muka  :          / "
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   6240
            TabIndex        =   98
            Top             =   10800
            Width           =   2295
         End
      End
      Begin VB.PictureBox Pic5 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   11415
         Left            =   600
         ScaleHeight     =   11415
         ScaleWidth      =   23535
         TabIndex        =   52
         Top             =   1440
         Visible         =   0   'False
         Width           =   23535
         Begin VB.CommandButton CMD4 
            BackColor       =   &H00FFFFFF&
            Height          =   840
            Left            =   14280
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm105.frx":FF0D
            MousePointer    =   99  'Custom
            Picture         =   "Frm105.frx":10217
            Style           =   1  'Graphical
            TabIndex        =   175
            ToolTipText     =   "Paparan sebelumnya"
            Top             =   9720
            Width           =   1200
         End
         Begin VB.CommandButton CMD5 
            BackColor       =   &H00FFFFFF&
            Height          =   840
            Left            =   15600
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm105.frx":10B56
            MousePointer    =   99  'Custom
            Picture         =   "Frm105.frx":10E60
            Style           =   1  'Graphical
            TabIndex        =   174
            ToolTipText     =   "Paparan seterusnya"
            Top             =   9720
            Width           =   1200
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
            Height          =   10305
            Left            =   240
            TabIndex        =   53
            ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
            Top             =   480
            Width           =   13995
            _ExtentX        =   24686
            _ExtentY        =   18177
            _Version        =   393216
            Rows            =   1
            Cols            =   0
            FixedCols       =   0
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   8454016
            BackColorSel    =   -2147483643
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
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "** Jumlah kutipan dari [Kad Kredit] dan [Kad Debit] adalah tidak termasuk caj perkhidmatan yang dikenakan oleh pihak bank."
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
            Left            =   14400
            TabIndex        =   69
            Top             =   2280
            Width           =   6615
         End
         Begin VB.Label L43_Text 
            Caption         =   "L43_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   17640
            TabIndex        =   66
            Top             =   5160
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label L44_Text 
            Caption         =   "L44_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   17640
            TabIndex        =   65
            Top             =   5640
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label L42_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L42_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   16440
            TabIndex        =   64
            Top             =   10560
            Width           =   2295
         End
         Begin VB.Label L41_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "L41_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   15480
            TabIndex        =   63
            Top             =   10560
            Width           =   735
         End
         Begin VB.Label L37_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L37_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   16560
            TabIndex        =   62
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label L36_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L36_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   16560
            TabIndex        =   61
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label L38_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L38_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   16560
            TabIndex        =   60
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label L39_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L39_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   16560
            TabIndex        =   59
            Top             =   1440
            Width           =   2295
         End
         Begin VB.Label L40_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L40_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   16560
            TabIndex        =   58
            Top             =   1680
            Width           =   2295
         End
         Begin VB.Label L35_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L35_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   16560
            TabIndex        =   57
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   ": RM  : RM  : RM  : RM  : RM  : RM"
            Height          =   1815
            Left            =   16065
            TabIndex        =   56
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   $"Frm105.frx":11786
            Height          =   1575
            Left            =   14400
            TabIndex        =   55
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label L34_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L34_Text"
            Height          =   255
            Left            =   360
            TabIndex        =   54
            Top             =   120
            Width           =   15375
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Paparan Muka  :          / "
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   14280
            TabIndex        =   67
            Top             =   10560
            Width           =   2295
         End
      End
      Begin VB.PictureBox Pic6 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   11415
         Left            =   2880
         ScaleHeight     =   11415
         ScaleWidth      =   23535
         TabIndex        =   70
         Top             =   720
         Visible         =   0   'False
         Width           =   23535
         Begin VB.CommandButton CMD9 
            BackColor       =   &H00FFFFFF&
            Height          =   840
            Left            =   13080
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm105.frx":11810
            MousePointer    =   99  'Custom
            Picture         =   "Frm105.frx":11B1A
            Style           =   1  'Graphical
            TabIndex        =   173
            ToolTipText     =   "Paparan sebelumnya"
            Top             =   9960
            Width           =   1200
         End
         Begin VB.CommandButton CMD10 
            BackColor       =   &H00FFFFFF&
            Height          =   840
            Left            =   14400
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm105.frx":12459
            MousePointer    =   99  'Custom
            Picture         =   "Frm105.frx":12763
            Style           =   1  'Graphical
            TabIndex        =   172
            ToolTipText     =   "Paparan seterusnya"
            Top             =   9960
            Width           =   1200
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
            Height          =   10545
            Left            =   240
            TabIndex        =   71
            ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
            Top             =   480
            Width           =   12765
            _ExtentX        =   22516
            _ExtentY        =   18600
            _Version        =   393216
            Rows            =   1
            Cols            =   0
            FixedCols       =   0
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   8454016
            BackColorSel    =   -2147483643
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
         Begin VB.Label L46_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L46_Text"
            Height          =   255
            Left            =   360
            TabIndex        =   85
            Top             =   120
            Width           =   15375
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   $"Frm105.frx":13089
            Height          =   1575
            Left            =   13200
            TabIndex        =   84
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   ": RM  : RM  : RM  : RM  : RM  : RM"
            Height          =   1815
            Left            =   14865
            TabIndex        =   83
            Top             =   600
            Width           =   495
         End
         Begin VB.Label L47_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L47_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   15360
            TabIndex        =   82
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label L52_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L52_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   15360
            TabIndex        =   81
            Top             =   1800
            Width           =   2295
         End
         Begin VB.Label L51_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L51_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   15360
            TabIndex        =   80
            Top             =   1560
            Width           =   2295
         End
         Begin VB.Label L50_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L50_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   15360
            TabIndex        =   79
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label L48_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L48_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   15360
            TabIndex        =   78
            Top             =   840
            Width           =   2295
         End
         Begin VB.Label L49_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L49_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   15360
            TabIndex        =   77
            Top             =   1080
            Width           =   2295
         End
         Begin VB.Label L53_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "L53_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   14280
            TabIndex        =   76
            Top             =   10800
            Width           =   735
         End
         Begin VB.Label L54_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L54_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   15240
            TabIndex        =   75
            Top             =   10800
            Width           =   2295
         End
         Begin VB.Label L56_Text 
            Caption         =   "L56_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   17640
            TabIndex        =   74
            Top             =   5640
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label L55_Text 
            Caption         =   "L55_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   17640
            TabIndex        =   73
            Top             =   5160
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "** Jumlah kutipan dari [Kad Kredit] dan [Kad Debit] adalah tidak termasuk caj perkhidmatan yang dikenakan oleh pihak bank."
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
            Left            =   13200
            TabIndex        =   72
            Top             =   2280
            Width           =   6615
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "Paparan Muka  :          / "
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   13080
            TabIndex        =   86
            Top             =   10800
            Width           =   2295
         End
      End
      Begin VB.Label L64_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Simpanan duit di kedai oleh pelanggan"
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
         MouseIcon       =   "Frm105.frx":13113
         MousePointer    =   99  'Custom
         TabIndex        =   100
         Top             =   120
         Width           =   3495
      End
      Begin VB.Label L57_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Kemasukkan tunai ke kedai"
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
         Left            =   5880
         MouseIcon       =   "Frm105.frx":1341D
         MousePointer    =   99  'Custom
         TabIndex        =   99
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label L45_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai bayaran tempahan"
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
         Left            =   3360
         MouseIcon       =   "Frm105.frx":13727
         MousePointer    =   99  'Custom
         TabIndex        =   87
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label L33_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai bayaran ansuran"
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
         Left            =   14880
         MouseIcon       =   "Frm105.frx":13A31
         MousePointer    =   99  'Custom
         TabIndex        =   51
         Top             =   120
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label L21_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai servis"
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
         Left            =   1800
         MouseIcon       =   "Frm105.frx":13D3B
         MousePointer    =   99  'Custom
         TabIndex        =   34
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label L9_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai jualan"
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
         MouseIcon       =   "Frm105.frx":14045
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   $"Frm105.frx":1434F
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   960
      Left            =   6000
      TabIndex        =   163
      Top             =   480
      Width           =   6570
   End
   Begin VB.Label L108_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ringkasan Report"
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
      MouseIcon       =   "Frm105.frx":143E5
      MousePointer    =   99  'Custom
      TabIndex        =   162
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label L8_Text 
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
      Left            =   3720
      MouseIcon       =   "Frm105.frx":146EF
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label L7_Text 
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
      Left            =   2040
      MouseIcon       =   "Frm105.frx":149F9
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label L1_Text 
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
      Left            =   20880
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label L2_Text 
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
      Left            =   20880
      TabIndex        =   1
      Top             =   555
      Visible         =   0   'False
      Width           =   2100
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
      Left            =   120
      MouseIcon       =   "Frm105.frx":14D03
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Menu Frm105_PM_Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm105_SM_senarai_jualan 
         Caption         =   "Cetak senarai jualan"
      End
   End
   Begin VB.Menu Frm105_PM_Menu2 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm105_SM_senarai_servis 
         Caption         =   "Cetak senarai servis"
      End
   End
   Begin VB.Menu Frm105_PM_Menu3 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm105_SM_senarai_ansuran 
         Caption         =   "Cetak senarai bayaran ansuran"
      End
   End
   Begin VB.Menu Frm105_PM_Menu4 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm105_SM_senarai_tempahan 
         Caption         =   "Cetak senarai bayaran tempahan"
      End
   End
   Begin VB.Menu Frm105_PM_Menu5 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm105_SM_senarai_trade_in 
         Caption         =   "Cetak senarai trade in"
      End
   End
   Begin VB.Menu Frm105_PM_Menu6 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm105_SM_senarai_barang_agen 
         Caption         =   "Cetak senarai belian barang dari agen"
      End
   End
   Begin VB.Menu Frm105_PM_Menu7 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm105_SM_senarai_belanja 
         Caption         =   "Cetak senarai perbelanjaan kedai"
      End
   End
   Begin VB.Menu Frm105_PM_Menu8 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm105_SM_senarai_gaji 
         Caption         =   "Cetak senarai pembayaran gaji pekerja"
      End
   End
End
Attribute VB_Name = "Frm105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD1_Click()
'on error resume next
Frm105.L5_Text = Frm105.DTPicker1 'Tarikh mula
Frm105.L6_Text = Frm105.DTPicker2 'Tarikh akhir
Frm105.L110_Text = Frm105.CBB1 'Nama pekerja
Frm105.L111_Text = Frm105.CBB1

If Frm105.CBB1 <> "Semua" Then
    If InStr(1, Frm105.CBB1, "  |  ") <> 0 Then
    
        Frm105.L110_Text = Split(Frm105.CBB1, "  |  ")(1)
        Frm105.L111_Text = Split(Frm105.CBB1, "  |  ")(0)
        
    End If
End If
        
Note = "Sistem mungkin akan mengambil sedikit masa untuk mengeluarkan report kewangan." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    
    Call Frm106_initial_setting
    Call Frm106_penyata_akaun
    
    Frm105.Pic1.Visible = False
    Frm105.L7_Text.Visible = True
    Frm105.L8_Text.Visible = True
    Frm105.L108_Text.Visible = True
    MsgBox "Sila lihat rekod kewangan dalam tempoh report dari menu [Debit] atau [Kredit] atau [Ringkasan Report].", vbInformation, "Info"
        
End If
End Sub
Private Sub CMD10_Click()
'on error resume next
Dim Frm105_LM_CURR_PAGE As Double
Dim Frm105_LM_TOTAL_PAGE As Double

Frm105_LM_CURR_PAGE = 0
Frm105_LM_TOTAL_PAGE = 0

If Frm105.L53_Text <> vbNullString And IsNumeric(Frm105.L53_Text) Then
    If Frm105.L54_Text <> vbNullString And IsNumeric(Frm105.L54_Text) Then
        Frm105_LM_CURR_PAGE = Frm105.L53_Text
        Frm105_LM_TOTAL_PAGE = Frm105.L54_Text
        
        If Frm105_LM_CURR_PAGE < Frm105_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm105_senarai_tempahan_header
            Call Frm105_senarai_tempahan
            
        End If
    End If
End If
End Sub
Private Sub CMD11_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm105_senarai_cash_in_header
Call Frm105_senarai_cash_in
End Sub
Private Sub CMD12_Click()
'on error resume next
Dim Frm105_LM_CURR_PAGE As Double
Dim Frm105_LM_TOTAL_PAGE As Double

Frm105_LM_CURR_PAGE = 0
Frm105_LM_TOTAL_PAGE = 0

If Frm105.L60_Text <> vbNullString And IsNumeric(Frm105.L60_Text) Then
    If Frm105.L61_Text <> vbNullString And IsNumeric(Frm105.L61_Text) Then
        Frm105_LM_CURR_PAGE = Frm105.L60_Text
        Frm105_LM_TOTAL_PAGE = Frm105.L61_Text
        
        If Frm105_LM_CURR_PAGE < Frm105_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm105_senarai_cash_in_header
            Call Frm105_senarai_cash_in
            
        End If
    End If
End If
End Sub
Private Sub CMD13_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm105_simpanan_duit_header
Call Frm105_simpanan_duit
End Sub
Private Sub CMD14_Click()
'on error resume next
Dim Frm105_LM_CURR_PAGE As Double
Dim Frm105_LM_TOTAL_PAGE As Double

Frm105_LM_CURR_PAGE = 0
Frm105_LM_TOTAL_PAGE = 0

If Frm105.L67_Text <> vbNullString And IsNumeric(Frm105.L67_Text) Then
    If Frm105.L68_Text <> vbNullString And IsNumeric(Frm105.L68_Text) Then
        Frm105_LM_CURR_PAGE = Frm105.L67_Text
        Frm105_LM_TOTAL_PAGE = Frm105.L68_Text
        
        If Frm105_LM_CURR_PAGE < Frm105_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm105_simpanan_duit_header
            Call Frm105_simpanan_duit
            
        End If
    End If
End If
End Sub
Private Sub CMD15_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm105_belian_trade_in_header
Call Frm105_belian_trade_in
End Sub
Private Sub CMD16_Click()
'on error resume next
Dim Frm105_LM_CURR_PAGE As Double
Dim Frm105_LM_TOTAL_PAGE As Double

Frm105_LM_CURR_PAGE = 0
Frm105_LM_TOTAL_PAGE = 0

If Frm105.L74_Text <> vbNullString And IsNumeric(Frm105.L74_Text) Then
    If Frm105.L75_Text <> vbNullString And IsNumeric(Frm105.L75_Text) Then
        Frm105_LM_CURR_PAGE = Frm105.L74_Text
        Frm105_LM_TOTAL_PAGE = Frm105.L75_Text
        
        If Frm105_LM_CURR_PAGE < Frm105_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm105_belian_trade_in_header
            Call Frm105_belian_trade_in
            
        End If
    End If
End If
End Sub
Private Sub CMD17_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm105_belian_barang_agen_header
Call Frm105_belian_barang_agen
End Sub
Private Sub CMD18_Click()
'on error resume next
Dim Frm105_LM_CURR_PAGE As Double
Dim Frm105_LM_TOTAL_PAGE As Double

Frm105_LM_CURR_PAGE = 0
Frm105_LM_TOTAL_PAGE = 0

If Frm105.L81_Text <> vbNullString And IsNumeric(Frm105.L81_Text) Then
    If Frm105.L82_Text <> vbNullString And IsNumeric(Frm105.L82_Text) Then
        Frm105_LM_CURR_PAGE = Frm105.L81_Text
        Frm105_LM_TOTAL_PAGE = Frm105.L82_Text
        
        If Frm105_LM_CURR_PAGE < Frm105_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm105_belian_barang_agen_header
            Call Frm105_belian_barang_agen
            
        End If
    End If
End If
End Sub
Private Sub CMD19_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm105_ambilan_tunai_header
Call Frm105_ambilan_tunai
End Sub
Private Sub CMD2_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm105_senarai_servis_header
Call Frm105_senarai_servis
End Sub
Private Sub CMD20_Click()
'on error resume next
Dim Frm105_LM_CURR_PAGE As Double
Dim Frm105_LM_TOTAL_PAGE As Double

Frm105_LM_CURR_PAGE = 0
Frm105_LM_TOTAL_PAGE = 0

If Frm105.L88_Text <> vbNullString And IsNumeric(Frm105.L88_Text) Then
    If Frm105.L89_Text <> vbNullString And IsNumeric(Frm105.L89_Text) Then
        Frm105_LM_CURR_PAGE = Frm105.L88_Text
        Frm105_LM_TOTAL_PAGE = Frm105.L89_Text
        
        If Frm105_LM_CURR_PAGE < Frm105_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm105_ambilan_tunai_header
            Call Frm105_ambilan_tunai
            
        End If
    End If
End If
End Sub
Private Sub CMD21_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm105_perbelanjaan_header
Call Frm105_perbelanjaan
End Sub
Private Sub CMD22_Click()
'on error resume next
Dim Frm105_LM_CURR_PAGE As Double
Dim Frm105_LM_TOTAL_PAGE As Double

Frm105_LM_CURR_PAGE = 0
Frm105_LM_TOTAL_PAGE = 0

If Frm105.L95_Text <> vbNullString And IsNumeric(Frm105.L95_Text) Then
    If Frm105.L96_Text <> vbNullString And IsNumeric(Frm105.L96_Text) Then
        Frm105_LM_CURR_PAGE = Frm105.L95_Text
        Frm105_LM_TOTAL_PAGE = Frm105.L96_Text
        
        If Frm105_LM_CURR_PAGE < Frm105_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm105_perbelanjaan_header
            Call Frm105_perbelanjaan
            
        End If
    End If
End If
End Sub
Private Sub CMD23_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm105_bayaran_gaji_header
Call Frm105_bayaran_gaji
End Sub
Private Sub CMD24_Click()
'on error resume next
Dim Frm105_LM_CURR_PAGE As Double
Dim Frm105_LM_TOTAL_PAGE As Double

Frm105_LM_CURR_PAGE = 0
Frm105_LM_TOTAL_PAGE = 0

If Frm105.L104_Text <> vbNullString And IsNumeric(Frm105.L104_Text) Then
    If Frm105.L105_Text <> vbNullString And IsNumeric(Frm105.L105_Text) Then
        Frm105_LM_CURR_PAGE = Frm105.L104_Text
        Frm105_LM_TOTAL_PAGE = Frm105.L105_Text
        
        If Frm105_LM_CURR_PAGE < Frm105_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm105_perbelanjaan_header
            Call Frm105_perbelanjaan
            
        End If
    End If
End If
End Sub
Private Sub CMD3_Click()
'on error resume next
Dim Frm105_LM_CURR_PAGE As Double
Dim Frm105_LM_TOTAL_PAGE As Double

Frm105_LM_CURR_PAGE = 0
Frm105_LM_TOTAL_PAGE = 0

If Frm105.L29_Text <> vbNullString And IsNumeric(Frm105.L29_Text) Then
    If Frm105.L30_Text <> vbNullString And IsNumeric(Frm105.L30_Text) Then
        Frm105_LM_CURR_PAGE = Frm105.L29_Text
        Frm105_LM_TOTAL_PAGE = Frm105.L30_Text
        
        If Frm105_LM_CURR_PAGE < Frm105_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm105_senarai_servis_header
            Call Frm105_senarai_servis
            
        End If
    End If
End If
End Sub
Private Sub CMD4_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm105_senarai_ansuran_header
Call Frm105_senarai_ansuran
End Sub
Private Sub CMD5_Click()
'on error resume next
Dim Frm105_LM_CURR_PAGE As Double
Dim Frm105_LM_TOTAL_PAGE As Double

Frm105_LM_CURR_PAGE = 0
Frm105_LM_TOTAL_PAGE = 0

If Frm105.L41_Text <> vbNullString And IsNumeric(Frm105.L41_Text) Then
    If Frm105.L42_Text <> vbNullString And IsNumeric(Frm105.L42_Text) Then
        Frm105_LM_CURR_PAGE = Frm105.L41_Text
        Frm105_LM_TOTAL_PAGE = Frm105.L42_Text
        
        If Frm105_LM_CURR_PAGE < Frm105_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm105_senarai_ansuran_header
            Call Frm105_senarai_ansuran
            
        End If
    End If
End If
End Sub
Private Sub CMD7_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm105_senarai_jualan_header
Call Frm105_senarai_jualan
End Sub
Private Sub CMD8_Click()
'on error resume next
Dim Frm105_LM_CURR_PAGE As Double
Dim Frm105_LM_TOTAL_PAGE As Double

Frm105_LM_CURR_PAGE = 0
Frm105_LM_TOTAL_PAGE = 0

If Frm105.L17_Text <> vbNullString And IsNumeric(Frm105.L17_Text) Then
    If Frm105.L18_Text <> vbNullString And IsNumeric(Frm105.L18_Text) Then
        Frm105_LM_CURR_PAGE = Frm105.L17_Text
        Frm105_LM_TOTAL_PAGE = Frm105.L18_Text
        
        If Frm105_LM_CURR_PAGE < Frm105_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm105_senarai_jualan_header
            Call Frm105_senarai_jualan
            
        End If
    End If
End If
End Sub
Private Sub CMD9_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm105_senarai_tempahan_header
Call Frm105_senarai_tempahan
End Sub
Private Sub Form_Load()
'on error resume next
Frm105.DTPicker1 = DateTime.Date
Frm105.DTPicker2 = DateTime.Date

Frm105.L7_Text.Visible = False
Frm105.L8_Text.Visible = False
Frm105.L108_Text.Visible = False

Frm105.CBB1.Clear
Frm105.CBB1.AddItem "Semua"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where status='" & "Aktif" & "' AND InvestorSmall = 0 AND InvestorBig = 0", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    If rs!user_level = 1 Or rs!user_level = 2 Or rs!user_level = 3 Or rs!user_level = G_LEVEL_USER Then Frm105.CBB1.AddItem rs!Samaran & "  |  " & rs!NoPekerja
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm105.CBB1 = "Semua"
End Sub
Private Sub Frm105_SM_senarai_ansuran_Click()
'on error resume next
Dim TM As Date
Dim TA As Date

TM = Frm105.L5_Text 'Tarikh Mula
TA = Frm105.L6_Text 'Tarikh Akhir

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
Report64.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report64.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report64.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report64.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report64.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report64.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report64.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report64.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report64.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report64.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Report64.Sections("Section4").Controls("L1").Caption = vbNullString 'Header
Report64.Sections("Section5").Controls("L2").Caption = "0.00" 'Jumlah
Report64.Sections("Section5").Controls("L3").Caption = "0.00" 'Tunai
Report64.Sections("Section5").Controls("L4").Caption = "0.00" 'Bank In
Report64.Sections("Section5").Controls("L5").Caption = "0.00" 'Kad Kredit
Report64.Sections("Section5").Controls("L6").Caption = "0.00" 'Kad Debit
Report64.Sections("Section5").Controls("L7").Caption = "0.00" 'Simpanan Di Kedai
Report64.Sections("Section5").Controls("L8").Caption = vbNullString 'Timestamp

Report64.Sections("Section4").Controls("L1").Caption = Frm105.L34_Text 'Header
Report64.Sections("Section5").Controls("L2").Caption = Frm105.L35_Text 'Jumlah
Report64.Sections("Section5").Controls("L3").Caption = Frm105.L36_Text 'Tunai
Report64.Sections("Section5").Controls("L4").Caption = Frm105.L37_Text 'Bank In
Report64.Sections("Section5").Controls("L5").Caption = Frm105.L38_Text 'Kad Kredit
Report64.Sections("Section5").Controls("L6").Caption = Frm105.L39_Text 'Kad Debit
Report64.Sections("Section5").Controls("L7").Caption = Frm105.L40_Text 'Simpanan Di Kedai
Report64.Sections("Section5").Controls("L8").Caption = Now 'Timestamp

'### Paparan Penyata ### - Start
Set rs = New ADODB.Recordset

If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 29_akaun_ansuran where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report64.DataSource = rs
    Report64.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing
'### Paparan Penyata ### - End
End Sub
Private Sub Frm105_SM_senarai_barang_agen_Click()
'on error resume next
Dim TM As Date
Dim TA As Date

TM = Frm105.L5_Text 'Tarikh Mula
TA = Frm105.L6_Text 'Tarikh Akhir

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
Report67.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report67.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report67.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report67.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report67.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report67.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report67.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report67.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report67.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report67.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Report67.Sections("Section4").Controls("L1").Caption = vbNullString 'Header
Report67.Sections("Section5").Controls("L2").Caption = "0.00" 'Jumlah
Report67.Sections("Section5").Controls("L8").Caption = vbNullString 'Timestamp

Report67.Sections("Section4").Controls("L1").Caption = Frm105.L79_Text 'Header
Report67.Sections("Section5").Controls("L2").Caption = Frm105.L80_Text 'Jumlah
Report67.Sections("Section5").Controls("L8").Caption = Now 'Timestamp

'### Paparan Penyata ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 51_voucher_belian_agen where tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND flag_bayaran='" & "1" & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report67.DataSource = rs
    Report67.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing
'### Paparan Penyata ### - End
End Sub
Private Sub Frm105_SM_senarai_belanja_Click()
'on error resume next
Dim TM As Date
Dim TA As Date

TM = Frm105.L5_Text 'Tarikh Mula
TA = Frm105.L6_Text 'Tarikh Akhir

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
Report68.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report68.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report68.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report68.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report68.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report68.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report68.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report68.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report68.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report68.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Report68.Sections("Section4").Controls("L1").Caption = vbNullString 'Header
Report68.Sections("Section5").Controls("L2").Caption = "0.00" 'Jumlah
Report68.Sections("Section5").Controls("L8").Caption = vbNullString 'Timestamp

Report68.Sections("Section4").Controls("L1").Caption = Frm105.L93_Text 'Header
Report68.Sections("Section5").Controls("L2").Caption = Frm105.L94_Text 'Jumlah
Report68.Sections("Section5").Controls("L8").Caption = Now 'Timestamp

'### Paparan Penyata ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 39_akaun_expense where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report68.DataSource = rs
    Report68.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing
'### Paparan Penyata ### - End
End Sub
Private Sub Frm105_SM_senarai_gaji_Click()
'on error resume next
Dim TM As Date
Dim TA As Date

TM = Frm105.L5_Text 'Tarikh Mula
TA = Frm105.L6_Text 'Tarikh Akhir

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
Report69.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report69.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report69.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report69.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report69.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report69.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report69.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report69.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report69.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report69.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Report69.Sections("Section4").Controls("L1").Caption = vbNullString 'Header
Report69.Sections("Section5").Controls("L2").Caption = "0.00" 'Jumlah
Report69.Sections("Section5").Controls("L3").Caption = "0.00" 'Tunai
Report69.Sections("Section5").Controls("L4").Caption = "0.00" 'Bank In
Report69.Sections("Section5").Controls("L8").Caption = vbNullString 'Timestamp

Report69.Sections("Section4").Controls("L1").Caption = Frm105.L100_Text 'Header
Report69.Sections("Section5").Controls("L2").Caption = Frm105.L101_Text 'Jumlah
Report69.Sections("Section5").Controls("L3").Caption = Frm105.L102_Text 'Tunai
Report69.Sections("Section5").Controls("L4").Caption = Frm105.L103_Text 'Bank In
Report69.Sections("Section5").Controls("L8").Caption = Now 'Timestamp

'### Paparan Penyata ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from payslip where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report69.DataSource = rs
    Report69.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing
'### Paparan Penyata ### - End
End Sub
Private Sub Frm105_SM_senarai_jualan_Click()
'On Error GoTo logging:
Dim TM As Date
Dim TA As Date

TM = Frm105.L5_Text 'Tarikh Mula
TA = Frm105.L6_Text 'Tarikh Akhir
If Frm105.L110_Text = "Semua" Then
    Frm105_LM_SEARCH_1 = Null
    Frm105_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm105_LM_SEARCH_1 = Frm105.L110_Text
    Frm105_LM_SEARCH_1_LOGIC = "="
End If

PRINTER_FOUND = 1 '0 : Printer Not Found , 1 : Printer Found
'### Reset maklumat kedai ### - Start
Report62.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report62.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report62.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report62.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report62.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report62.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report62.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report62.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report62.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report62.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Report62.Sections("Section4").Controls("L1").Caption = vbNullString 'Header
Report62.Sections("Section5").Controls("L2").Caption = "0.00" 'Jumlah
Report62.Sections("Section5").Controls("L3").Caption = "0.00" 'Tunai
Report62.Sections("Section5").Controls("L4").Caption = "0.00" 'Bank In
Report62.Sections("Section5").Controls("L5").Caption = "0.00" 'Kad Kredit
Report62.Sections("Section5").Controls("L6").Caption = "0.00" 'Kad Debit
Report62.Sections("Section5").Controls("L7").Caption = "0.00" 'Simpanan Di Kedai
Report62.Sections("Section5").Controls("L8").Caption = vbNullString 'Timestamp

Report62.Sections("Section4").Controls("L1").Caption = Frm105.L10_Text 'Header
Report62.Sections("Section5").Controls("L2").Caption = Frm105.L11_Text 'Jumlah
Report62.Sections("Section5").Controls("L3").Caption = Frm105.L12_Text 'Tunai
Report62.Sections("Section5").Controls("L4").Caption = Frm105.L13_Text 'Bank In
Report62.Sections("Section5").Controls("L5").Caption = Frm105.L14_Text 'Kad Kredit
Report62.Sections("Section5").Controls("L6").Caption = Frm105.L15_Text 'Kad Debit
Report62.Sections("Section5").Controls("L7").Caption = Frm105.L16_Text 'Simpanan Di Kedai
Report62.Sections("Section5").Controls("L8").Caption = Now 'Timestamp

'### Paparan Penyata ### - Start
LM_CONN = 2
re_conn_2:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where menu = 0 AND status = 1 AND no_pekerja " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report62.DataSource = rs
    Report62.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing
'### Paparan Penyata ### - End

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm105 : Frm105_SM_senarai_jualan_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If LM_ERR_NO = "3704" Or LM_ERR_NO = "-2147467259" Or LM_ERR_NO = "-2147217887" Then
    
    Call Main

    If LM_CONN = 1 Then
        Resume re_conn_1:
    ElseIf LM_CONN = 2 Then
        Resume re_conn_2:
    End If
Else
    Resume Next
End If
End Sub
Private Sub Frm105_SM_senarai_servis_Click()
'On Error GoTo logging:
Dim TM As Date
Dim TA As Date

TM = Frm105.L5_Text 'Tarikh Mula
TA = Frm105.L6_Text 'Tarikh Akhir
If Frm105.L110_Text = "Semua" Then
    Frm105_LM_SEARCH_1 = Null
    Frm105_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm105_LM_SEARCH_1 = Frm105.L110_Text
    Frm105_LM_SEARCH_1_LOGIC = "="
End If

PRINTER_FOUND = 1 '0 : Printer Not Found , 1 : Printer Found

'### Reset maklumat kedai ### - Start
Report63.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report63.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report63.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report63.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report63.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report63.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report63.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report63.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report63.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report63.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Report63.Sections("Section4").Controls("L1").Caption = vbNullString 'Header
Report63.Sections("Section5").Controls("L2").Caption = "0.00" 'Jumlah
Report63.Sections("Section5").Controls("L3").Caption = "0.00" 'Tunai
Report63.Sections("Section5").Controls("L4").Caption = "0.00" 'Bank In
Report63.Sections("Section5").Controls("L5").Caption = "0.00" 'Kad Kredit
Report63.Sections("Section5").Controls("L6").Caption = "0.00" 'Kad Debit
Report63.Sections("Section5").Controls("L7").Caption = "0.00" 'Simpanan Di Kedai
Report63.Sections("Section5").Controls("L8").Caption = vbNullString 'Timestamp

Report63.Sections("Section4").Controls("L1").Caption = Frm105.L22_Text 'Header
Report63.Sections("Section5").Controls("L2").Caption = Frm105.L23_Text 'Jumlah
Report63.Sections("Section5").Controls("L3").Caption = Frm105.L24_Text 'Tunai
Report63.Sections("Section5").Controls("L4").Caption = Frm105.L25_Text 'Bank In
Report63.Sections("Section5").Controls("L5").Caption = Frm105.L26_Text 'Kad Kredit
Report63.Sections("Section5").Controls("L6").Caption = Frm105.L27_Text 'Kad Debit
Report63.Sections("Section5").Controls("L7").Caption = Frm105.L28_Text 'Simpanan Di Kedai
Report63.Sections("Section5").Controls("L8").Caption = Now 'Timestamp

'### Paparan Penyata ### - Start
LM_CONN = 2
re_conn_2:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where menu = 1 AND no_pekerja " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND status = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report63.DataSource = rs
    Report63.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing
'### Paparan Penyata ### - End

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm105 : Frm105_SM_senarai_servis_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If LM_ERR_NO = "3704" Or LM_ERR_NO = "-2147467259" Or LM_ERR_NO = "-2147217887" Then
    
    Call Main

    If LM_CONN = 1 Then
        Resume re_conn_1:
    ElseIf LM_CONN = 2 Then
        Resume re_conn_2:
    End If
Else
    Resume Next
End If
End Sub
Private Sub Frm105_SM_senarai_tempahan_Click()
'on error resume next
Dim TM As Date
Dim TA As Date

TM = Frm105.L5_Text 'Tarikh Mula
TA = Frm105.L6_Text 'Tarikh Akhir

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
Report65.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report65.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report65.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report65.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report65.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report65.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report65.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report65.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report65.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report65.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Report65.Sections("Section4").Controls("L1").Caption = vbNullString 'Header
Report65.Sections("Section5").Controls("L2").Caption = "0.00" 'Jumlah
Report65.Sections("Section5").Controls("L3").Caption = "0.00" 'Tunai
Report65.Sections("Section5").Controls("L4").Caption = "0.00" 'Bank In
Report65.Sections("Section5").Controls("L5").Caption = "0.00" 'Kad Kredit
Report65.Sections("Section5").Controls("L6").Caption = "0.00" 'Kad Debit
Report65.Sections("Section5").Controls("L7").Caption = "0.00" 'Simpanan Di Kedai
Report65.Sections("Section5").Controls("L8").Caption = vbNullString 'Timestamp

Report65.Sections("Section4").Controls("L1").Caption = Frm105.L46_Text 'Header
Report65.Sections("Section5").Controls("L2").Caption = Frm105.L47_Text 'Jumlah
Report65.Sections("Section5").Controls("L3").Caption = Frm105.L48_Text 'Tunai
Report65.Sections("Section5").Controls("L4").Caption = Frm105.L49_Text 'Bank In
Report65.Sections("Section5").Controls("L5").Caption = Frm105.L50_Text 'Kad Kredit
Report65.Sections("Section5").Controls("L6").Caption = Frm105.L51_Text 'Kad Debit
Report65.Sections("Section5").Controls("L7").Caption = Frm105.L52_Text 'Simpanan Di Kedai
Report65.Sections("Section5").Controls("L8").Caption = Now 'Timestamp

'### Paparan Penyata ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where (menu = 2 OR menu = 3) AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND status = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report65.DataSource = rs
    Report65.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing
'### Paparan Penyata ### - End
End Sub
Private Sub Frm105_SM_senarai_trade_in_Click()
'On Error GoTo logging:
Dim TM As Date
Dim TA As Date

TM = Frm105.L5_Text 'Tarikh Mula
TA = Frm105.L6_Text 'Tarikh Akhir
If Frm105.L110_Text = "Semua" Then
    Frm105_LM_SEARCH_1 = Null
    Frm105_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm105_LM_SEARCH_1 = Frm105.L110_Text
    Frm105_LM_SEARCH_1_LOGIC = "="
End If

PRINTER_FOUND = 1 '0 : Printer Not Found , 1 : Printer Found

'### Reset maklumat kedai ### - Start
Report66.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report66.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report66.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report66.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report66.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report66.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report66.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report66.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report66.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report66.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Report66.Sections("Section4").Controls("L1").Caption = vbNullString 'Header
Report66.Sections("Section5").Controls("L2").Caption = "0.00" 'Jumlah
Report66.Sections("Section5").Controls("L8").Caption = vbNullString 'Timestamp

Report66.Sections("Section4").Controls("L1").Caption = Frm105.L72_Text 'Header
Report66.Sections("Section5").Controls("L2").Caption = Frm105.L73_Text 'Jumlah
Report66.Sections("Section5").Controls("L10").Caption = Frm105.L112_Text
Report66.Sections("Section5").Controls("L11").Caption = Frm105.L113_Text
Report66.Sections("Section5").Controls("L8").Caption = Now 'Timestamp

'### Paparan Penyata ### - Start
LM_CONN = 2
re_conn_2:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 16_gold_bar_belian where tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND no_pekerja " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND status = 1 AND flag_trade_in = 1  order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report66.DataSource = rs
    Report66.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing
'### Paparan Penyata ### - End

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Frm105 : Frm105_SM_senarai_trade_in_Click" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If LM_ERR_NO = "3704" Or LM_ERR_NO = "-2147467259" Or LM_ERR_NO = "-2147217887" Then
    
    Call Main

    If LM_CONN = 1 Then
        Resume re_conn_1:
    ElseIf LM_CONN = 2 Then
        Resume re_conn_2:
    End If
Else
    Resume Next
End If
End Sub
Private Sub L108_Text_Click()
'on error resume next
Frm106.Show
Frm106.Picture = MDI_frm1.Picture
Frm105.Hide
End Sub
Private Sub L21_Text_Click()
'on error resume next
If Frm105.Pic4.Visible = False Then
    
    Call Frm105_debit_setting
    
    GM_NEXT_PREV = 0
    Frm105.L31_Text = -1 'Titik Pencarian Data
    Frm105.L32_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm105.L29_Text = 0 'Paparan Page ke-xxx
    
    Call Frm105_senarai_servis_header
    Call Frm105_senarai_servis
   
    'Frm105.Pic3.Visible = True
Else
    Frm105.Pic4.Visible = False
End If
End Sub
Private Sub L33_Text_Click()
'on error resume next
If Frm105.Pic5.Visible = False Then
    
    Call Frm105_debit_setting
    
    GM_NEXT_PREV = 0
    Frm105.L43_Text = -1 'Titik Pencarian Data
    Frm105.L44_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm105.L41_Text = 0 'Paparan Page ke-xxx
    
    Call Frm105_senarai_ansuran_header
    Call Frm105_senarai_ansuran
   
    'Frm105.Pic3.Visible = True
Else
    Frm105.Pic5.Visible = False
End If
End Sub
Private Sub L4_Text_Click()
'on error resume next
If Frm105.Pic1.Visible = False Then
    Call Frm105_initial_setting
    Call Frm105_debit_setting
    
    Frm105.L7_Text.Visible = False
    Frm105.L8_Text.Visible = False
    Frm105.L108_Text.Visible = False
    
    Frm105.Pic1.Visible = True
Else
    Frm105.Pic1.Visible = False
End If
End Sub
Private Sub L45_Text_Click()
'on error resume next
If Frm105.Pic6.Visible = False Then
    
    Call Frm105_debit_setting
    
    GM_NEXT_PREV = 0
    Frm105.L55_Text = -1 'Titik Pencarian Data
    Frm105.L56_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm105.L53_Text = 0 'Paparan Page ke-xxx
    
    Call Frm105_senarai_tempahan_header
    Call Frm105_senarai_tempahan
   
    'Frm105.Pic3.Visible = True
Else
    Frm105.Pic6.Visible = False
End If
End Sub
Private Sub L57_Text_Click()
'on error resume next
If Frm105.Pic7.Visible = False Then
    
    Call Frm105_debit_setting
    
    GM_NEXT_PREV = 0
    Frm105.L62_Text = -1 'Titik Pencarian Data
    Frm105.L63_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm105.L60_Text = 0 'Paparan Page ke-xxx
    
    Call Frm105_senarai_cash_in_header
    Call Frm105_senarai_cash_in
   
    'Frm105.Pic3.Visible = True
Else
    Frm105.Pic7.Visible = False
End If
End Sub
Private Sub L64_Text_Click()
'on error resume next
If Frm105.Pic8.Visible = False Then
    
    Call Frm105_debit_setting
    
    GM_NEXT_PREV = 0
    Frm105.L69_Text = -1 'Titik Pencarian Data
    Frm105.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm105.L67_Text = 0 'Paparan Page ke-xxx
    
    Call Frm105_simpanan_duit_header
    Call Frm105_simpanan_duit
   
    'Frm105.Pic3.Visible = True
Else
    Frm105.Pic8.Visible = False
End If
End Sub
Private Sub L7_Text_Click()
'on error resume next
If Frm105.Pic2.Visible = False Then
    Call Frm105_initial_setting
    
    Frm105.Pic3.Visible = False
    Frm105.Pic4.Visible = False
    Frm105.Pic5.Visible = False
    Frm105.Pic6.Visible = False
    Frm105.Pic7.Visible = False
    Frm105.Pic8.Visible = False
    
    Frm105.Pic2.Visible = True
Else
    Frm105.Pic2.Visible = False
End If
End Sub
Private Sub L71_Text_Click()
'on error resume next
If Frm105.Pic10.Visible = False Then
    
    Call Frm105_kredit_setting
    
    GM_NEXT_PREV = 0
    Frm105.L76_Text = -1 'Titik Pencarian Data
    Frm105.L77_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm105.L74_Text = 0 'Paparan Page ke-xxx
    
    Call Frm105_belian_trade_in_header
    Call Frm105_belian_trade_in
   
    'Frm105.Pic3.Visible = True
Else
    Frm105.Pic10.Visible = False
End If
End Sub
Private Sub L78_Text_Click()
'on error resume next
If Frm105.Pic11.Visible = False Then
    
    Call Frm105_kredit_setting
    
    GM_NEXT_PREV = 0
    Frm105.L83_Text = -1 'Titik Pencarian Data
    Frm105.L84_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm105.L81_Text = 0 'Paparan Page ke-xxx
    
    Call Frm105_belian_barang_agen_header
    Call Frm105_belian_barang_agen
   
    'Frm105.Pic3.Visible = True
Else
    Frm105.Pic11.Visible = False
End If
End Sub
Private Sub L8_Text_Click()
'on error resume next
If Frm105.Pic9.Visible = False Then
    Call Frm105_initial_setting
    
    Frm105.Pic10.Visible = False
    Frm105.Pic11.Visible = False
    Frm105.Pic12.Visible = False
    Frm105.Pic13.Visible = False
    Frm105.Pic14.Visible = False
    
    Frm105.Pic9.Visible = True
Else
    Frm105.Pic9.Visible = False
End If
End Sub
Private Sub L85_Text_Click()
'on error resume next
If Frm105.Pic12.Visible = False Then
    
    Call Frm105_kredit_setting
    
    GM_NEXT_PREV = 0
    Frm105.L90_Text = -1 'Titik Pencarian Data
    Frm105.L91_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm105.L88_Text = 0 'Paparan Page ke-xxx
    
    Call Frm105_ambilan_tunai_header
    Call Frm105_ambilan_tunai
   
    'Frm105.Pic3.Visible = True
Else
    Frm105.Pic12.Visible = False
End If
End Sub
Private Sub L9_Text_Click()
'on error resume next
If Frm105.Pic3.Visible = False Then
    
    Call Frm105_debit_setting
    
    GM_NEXT_PREV = 0
    Frm105.L19_Text = -1 'Titik Pencarian Data
    Frm105.L20_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm105.L17_Text = 0 'Paparan Page ke-xxx
    
    Call Frm105_senarai_jualan_header
    Call Frm105_senarai_jualan
   
    'Frm105.Pic3.Visible = True
Else
    Frm105.Pic3.Visible = False
End If
End Sub
Private Sub L92_Text_Click()
'on error resume next
If Frm105.Pic13.Visible = False Then
    
    Call Frm105_kredit_setting
    
    GM_NEXT_PREV = 0
    Frm105.L97_Text = -1 'Titik Pencarian Data
    Frm105.L98_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm105.L95_Text = 0 'Paparan Page ke-xxx
    
    Call Frm105_perbelanjaan_header
    Call Frm105_perbelanjaan
   
    'Frm105.Pic3.Visible = True
Else
    Frm105.Pic13.Visible = False
End If
End Sub
Private Sub L99_Text_Click()
'on error resume next
If Frm105.Pic14.Visible = False Then
    
    Call Frm105_kredit_setting
    
    GM_NEXT_PREV = 0
    Frm105.L106_Text = -1 'Titik Pencarian Data
    Frm105.L107_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm105.L104_Text = 0 'Paparan Page ke-xxx
    
    Call Frm105_bayaran_gaji_header
    Call Frm105_bayaran_gaji
   
    'Frm105.Pic3.Visible = True
Else
    Frm105.Pic14.Visible = False
End If
End Sub
Private Sub MSFlexGrid1_DblClick()
'on error resume next
If Frm105.MSFlexGrid1 <> vbNullString Then
    PopupMenu Frm105_PM_Menu
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub MSFlexGrid10_DblClick()
'on error resume next
If Frm105.MSFlexGrid10 <> vbNullString Then
    PopupMenu Frm105_PM_Menu7
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub MSFlexGrid11_DblClick()
'on error resume next
If Frm105.MSFlexGrid11 <> vbNullString Then
    PopupMenu Frm105_PM_Menu8
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub MSFlexGrid2_DblClick()
'on error resume next
If Frm105.MSFlexGrid2 <> vbNullString Then
    PopupMenu Frm105_PM_Menu2
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub MSFlexGrid3_DblClick()
'on error resume next
If Frm105.MSFlexGrid3 <> vbNullString Then
    PopupMenu Frm105_PM_Menu3
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub MSFlexGrid4_DblClick()
'on error resume next
If Frm105.MSFlexGrid4 <> vbNullString Then
    PopupMenu Frm105_PM_Menu4
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub MSFlexGrid7_DblClick()
'on error resume next
If Frm105.MSFlexGrid7 <> vbNullString Then
    PopupMenu Frm105_PM_Menu5
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub MSFlexGrid8_DblClick()
'on error resume next
If Frm105.MSFlexGrid8 <> vbNullString Then
    PopupMenu Frm105_PM_Menu6
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub

Private Sub Tmr1_Timer()
'On Error Resume Next
Frm105.L1_Text = DateTime.Date
Frm105.L2_Text = DateTime.Time$
End Sub
