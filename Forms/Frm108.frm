VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm108 
   Caption         =   "Ambilan / hantaran stok ke cawangan /  kedai"
   ClientHeight    =   12735
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   23760
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12735
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   11400
      Left            =   16320
      ScaleHeight     =   11400
      ScaleWidth      =   21045
      TabIndex        =   2
      Top             =   11880
      Visible         =   0   'False
      Width           =   21045
   End
   Begin VB.PictureBox Pic7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   11400
      Left            =   9600
      ScaleHeight     =   11400
      ScaleWidth      =   21195
      TabIndex        =   153
      Top             =   600
      Visible         =   0   'False
      Width           =   21195
      Begin VB.CommandButton CMD26 
         BackColor       =   &H00FFFFFF&
         Height          =   740
         Left            =   18960
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm108.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Frm108.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   172
         ToolTipText     =   "Paparan sebelumnya"
         Top             =   10440
         Width           =   1000
      End
      Begin VB.CommandButton CMD27 
         BackColor       =   &H00FFFFFF&
         Height          =   740
         Left            =   20040
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm108.frx":0C49
         MousePointer    =   99  'Custom
         Picture         =   "Frm108.frx":0F53
         Style           =   1  'Graphical
         TabIndex        =   171
         ToolTipText     =   "Paparan seterusnya"
         Top             =   10440
         Width           =   1000
      End
      Begin VB.ComboBox CBB7 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   315
         ItemData        =   "Frm108.frx":1879
         Left            =   1500
         List            =   "Frm108.frx":187B
         Style           =   2  'Dropdown List
         TabIndex        =   167
         Top             =   2470
         Width           =   4725
      End
      Begin VB.CommandButton CMD28 
         BackColor       =   &H000080FF&
         Caption         =   "Report"
         Height          =   405
         Left            =   2280
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm108.frx":187D
         MousePointer    =   99  'Custom
         TabIndex        =   156
         Top             =   2880
         Width           =   2025
      End
      Begin VB.ComboBox CBB6 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   315
         ItemData        =   "Frm108.frx":1B87
         Left            =   1500
         List            =   "Frm108.frx":1B89
         Style           =   2  'Dropdown List
         TabIndex        =   155
         Top             =   2160
         Width           =   4725
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
         Left            =   285
         TabIndex        =   154
         Top             =   735
         Width           =   200
      End
      Begin MSComCtl2.DTPicker DTPicker5 
         Height          =   360
         Left            =   1500
         TabIndex        =   157
         Top             =   1275
         Width           =   4725
         _ExtentX        =   8334
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
         Format          =   415956992
         CurrentDate     =   41561
      End
      Begin MSComCtl2.DTPicker DTPicker6 
         Height          =   360
         Left            =   1500
         TabIndex        =   158
         Top             =   1635
         Width           =   4725
         _ExtentX        =   8334
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
         Format          =   415956992
         CurrentDate     =   41561
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid6 
         Height          =   9525
         Left            =   6720
         TabIndex        =   170
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   840
         Width           =   14400
         _ExtentX        =   25400
         _ExtentY        =   16801
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
      Begin VB.Label L60_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L60_Text"
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
         Height          =   195
         Left            =   600
         TabIndex        =   183
         Top             =   7560
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L61_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L61_Text"
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
         Height          =   195
         Left            =   600
         TabIndex        =   182
         Top             =   7920
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L56_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L56_Text"
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
         Left            =   18600
         TabIndex        =   178
         Top             =   10440
         Width           =   615
      End
      Begin VB.Label L55_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L55_Text"
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
         Left            =   18060
         TabIndex        =   177
         Top             =   10440
         Width           =   375
      End
      Begin VB.Label L57_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L57_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   10800
         TabIndex        =   176
         Top             =   10680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L58_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L58_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   10800
         TabIndex        =   175
         Top             =   10920
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L53_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L53_Text"
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
         Left            =   7560
         TabIndex        =   174
         Top             =   10440
         Width           =   975
      End
      Begin VB.Label L54_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L54_Text"
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
         Left            =   9975
         TabIndex        =   173
         Top             =   10440
         Width           =   2280
      End
      Begin VB.Label L52_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L52_Text"
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
         Left            =   6840
         TabIndex        =   169
         Top             =   600
         Width           =   13455
      End
      Begin VB.Label Label53 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Report"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   240
         TabIndex        =   168
         Top             =   2490
         Width           =   1695
      End
      Begin VB.Shape Shape8 
         Height          =   2895
         Left            =   120
         Top             =   600
         Width           =   6495
      End
      Begin VB.Label Label78 
         BackStyle       =   0  'Transparent
         Caption         =   "Cawangan"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   240
         TabIndex        =   166
         Top             =   2175
         Width           =   1695
      End
      Begin VB.Label Label77 
         BackStyle       =   0  'Transparent
         Caption         =   "** Jika tidak ditanda , sistem TIDAK akan mengeluarkan report mengikut tarikh."
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   600
         TabIndex        =   165
         Top             =   960
         Width           =   5850
      End
      Begin VB.Label Label76 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Mula "
         Height          =   255
         Left            =   240
         TabIndex        =   164
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label75 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Akhir "
         Height          =   255
         Left            =   240
         TabIndex        =   163
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label74 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila tandakan di sini jika ingin cari data mengikut tarikh."
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   570
         TabIndex        =   162
         Top             =   720
         Width           =   4890
      End
      Begin VB.Label L49_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L49_Text"
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
         Height          =   195
         Left            =   600
         TabIndex        =   161
         Top             =   6360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L50_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L50_Text"
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
         Height          =   195
         Left            =   600
         TabIndex        =   160
         Top             =   6720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L51_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L51_Text"
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
         Height          =   195
         Left            =   600
         TabIndex        =   159
         Top             =   7080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan :                    Jumlah Berat (g) : "
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
         Left            =   6720
         TabIndex        =   179
         Top             =   10440
         Width           =   3975
      End
      Begin VB.Label Label52 
         BackStyle       =   0  'Transparent
         Caption         =   "Paparan Muka  :       / "
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
         Left            =   16800
         TabIndex        =   180
         Top             =   10440
         Width           =   2295
      End
   End
   Begin VB.PictureBox Pic6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   11400
      Left            =   2400
      ScaleHeight     =   11400
      ScaleWidth      =   21045
      TabIndex        =   111
      Top             =   480
      Visible         =   0   'False
      Width           =   21045
      Begin VB.TextBox TB14 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   3225
         TabIndex        =   193
         Text            =   "TB14"
         Top             =   5640
         Visible         =   0   'False
         Width           =   3780
      End
      Begin VB.TextBox TB13 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   3225
         TabIndex        =   191
         Text            =   "TB13"
         Top             =   5340
         Visible         =   0   'False
         Width           =   3780
      End
      Begin VB.TextBox TB12 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   3225
         TabIndex        =   189
         Text            =   "TB12"
         Top             =   5040
         Visible         =   0   'False
         Width           =   3780
      End
      Begin VB.CheckBox CB8 
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
         Left            =   3240
         TabIndex        =   186
         Top             =   4420
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.CheckBox CB7 
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
         Left            =   1920
         TabIndex        =   185
         Top             =   4420
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.TextBox TB8 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1900
         TabIndex        =   124
         Text            =   "TB8"
         Top             =   690
         Width           =   6420
      End
      Begin VB.TextBox TB9 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1900
         TabIndex        =   123
         Text            =   "TB9"
         Top             =   990
         Width           =   6420
      End
      Begin VB.TextBox TB10 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1900
         TabIndex        =   122
         Text            =   "TB10"
         Top             =   1290
         Width           =   6420
      End
      Begin VB.ComboBox CBB4 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   315
         ItemData        =   "Frm108.frx":1B8B
         Left            =   1900
         List            =   "Frm108.frx":1B8D
         Style           =   2  'Dropdown List
         TabIndex        =   121
         Top             =   1680
         Width           =   6420
      End
      Begin VB.ComboBox CBB5 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   315
         ItemData        =   "Frm108.frx":1B8F
         Left            =   1900
         List            =   "Frm108.frx":1B91
         Style           =   2  'Dropdown List
         TabIndex        =   120
         Top             =   2000
         Width           =   6420
      End
      Begin VB.CommandButton CMD21 
         BackColor       =   &H000080FF&
         Caption         =   "Simpan Data"
         Height          =   405
         Left            =   3360
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm108.frx":1B93
         MousePointer    =   99  'Custom
         TabIndex        =   119
         Top             =   10680
         Width           =   2025
      End
      Begin VB.CommandButton CMD22 
         BackColor       =   &H000080FF&
         Caption         =   "Simpan Data"
         Height          =   405
         Left            =   2160
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm108.frx":1E9D
         MousePointer    =   99  'Custom
         TabIndex        =   118
         Top             =   10680
         Width           =   2025
      End
      Begin VB.CommandButton CMD23 
         BackColor       =   &H000080FF&
         Caption         =   "Batal"
         Height          =   405
         Left            =   4320
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm108.frx":21A7
         MousePointer    =   99  'Custom
         TabIndex        =   117
         Top             =   10680
         Width           =   2025
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
         Left            =   150
         TabIndex        =   116
         Top             =   3370
         Width           =   200
      End
      Begin VB.TextBox TB11 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1900
         TabIndex        =   115
         Text            =   "TB11"
         Top             =   3720
         Width           =   3900
      End
      Begin VB.CommandButton CMD20 
         BackColor       =   &H000080FF&
         Caption         =   "Masukkan Data"
         Height          =   405
         Left            =   5880
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm108.frx":24B1
         MousePointer    =   99  'Custom
         TabIndex        =   114
         Top             =   3660
         Width           =   2025
      End
      Begin VB.Timer Tmr2 
         Interval        =   100
         Left            =   120
         Top             =   120
      End
      Begin VB.CommandButton CMD24 
         BackColor       =   &H00FFFFFF&
         Height          =   740
         Left            =   18840
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm108.frx":27BB
         MousePointer    =   99  'Custom
         Picture         =   "Frm108.frx":2AC5
         Style           =   1  'Graphical
         TabIndex        =   113
         ToolTipText     =   "Paparan sebelumnya"
         Top             =   10365
         Width           =   1000
      End
      Begin VB.CommandButton CMD25 
         BackColor       =   &H00FFFFFF&
         Height          =   740
         Left            =   19920
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm108.frx":3404
         MousePointer    =   99  'Custom
         Picture         =   "Frm108.frx":370E
         Style           =   1  'Graphical
         TabIndex        =   112
         ToolTipText     =   "Paparan seterusnya"
         Top             =   10365
         Width           =   1000
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   360
         Left            =   1900
         TabIndex        =   125
         Top             =   2310
         Width           =   6405
         _ExtentX        =   11298
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
         Format          =   142344192
         CurrentDate     =   41561
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid5 
         Height          =   9645
         Left            =   8400
         TabIndex        =   126
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   600
         Width           =   12525
         _ExtentX        =   22093
         _ExtentY        =   17013
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
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm108.frx":4034
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   120
         TabIndex        =   195
         Top             =   6050
         Visible         =   0   'False
         Width           =   8265
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Jualan (RM) *"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         TabIndex        =   194
         Top             =   5640
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Perjanjian B *"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         TabIndex        =   192
         Top             =   5355
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Perjanjian A *"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         TabIndex        =   190
         Top             =   5055
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "(Sila masukkan data di bawah jika pilihan adalah ""Dijual"")"
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
         Height          =   285
         Left            =   3240
         TabIndex        =   188
         Top             =   4680
         Visible         =   0   'False
         Width           =   6105
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis *            Pulangan         Dijual"
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
         Height          =   285
         Left            =   840
         TabIndex        =   187
         Top             =   4395
         Visible         =   0   'False
         Width           =   7785
      End
      Begin VB.Label Label51 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Hantar *"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   148
         Top             =   2350
         Width           =   2385
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama *"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   145
         Top             =   720
         Width           =   1785
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan maklumat terperinci berkenaan cawangan / agen / pengedar yang memulangkan barang ini."
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
         Left            =   120
         TabIndex        =   144
         Top             =   240
         Width           =   8145
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Kad Pengenalan *"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   143
         Top             =   990
         Width           =   1785
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Telefon"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   142
         Top             =   1300
         Width           =   1785
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Cawangan / Kedai *"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   141
         Top             =   1710
         Width           =   1695
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pekerja *"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   140
         Top             =   2030
         Width           =   1695
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "Scanner Mode (Sila klik sini jika anda menggunakan scanner untuk scan data barang kemas)"
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   435
         TabIndex        =   139
         Top             =   3360
         Width           =   7665
      End
      Begin VB.Label L40_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L40_Text"
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
         Height          =   195
         Left            =   480
         TabIndex        =   138
         Top             =   7560
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan data yang akan dipulangkan oleh cawangan atau kedai."
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
         Height          =   420
         Left            =   120
         TabIndex        =   137
         Top             =   3000
         Width           =   7905
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Siri Produk *"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   150
         TabIndex        =   136
         Top             =   3750
         Width           =   1785
      End
      Begin VB.Shape Shape5 
         Height          =   1215
         Left            =   75
         Top             =   3240
         Width           =   8175
      End
      Begin VB.Label L41_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai barang yang akan dipulangkan / dijual."
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
         Height          =   195
         Left            =   8520
         TabIndex        =   135
         Top             =   360
         Width           =   10440
      End
      Begin VB.Label L42_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L42_Text"
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
         Height          =   240
         Left            =   9480
         TabIndex        =   134
         Top             =   10320
         Width           =   615
      End
      Begin VB.Label L45_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L45_Text"
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
         Left            =   18600
         TabIndex        =   133
         Top             =   10845
         Width           =   615
      End
      Begin VB.Label L44_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L44_Text"
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
         Left            =   18120
         TabIndex        =   132
         Top             =   10845
         Width           =   375
      End
      Begin VB.Label L46_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L46_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   8640
         TabIndex        =   131
         Top             =   10605
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L47_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L47_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   8640
         TabIndex        =   130
         Top             =   10845
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L43_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L43_Text"
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
         Height          =   240
         Left            =   11370
         TabIndex        =   129
         Top             =   10320
         Width           =   1335
      End
      Begin VB.Label L39_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L39_Text"
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
         Height          =   195
         Left            =   480
         TabIndex        =   128
         Top             =   7320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "** Sila tekan F2 untuk scan barang yang akan dipulangkan. (Jika menggunakan scanner mode)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   120
         TabIndex        =   127
         Top             =   4080
         Width           =   8265
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan :              Jumlah Berat :"
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
         Height          =   240
         Left            =   8640
         TabIndex        =   147
         Top             =   10320
         Width           =   3255
      End
      Begin VB.Label Label49 
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
         Left            =   16680
         TabIndex        =   146
         Top             =   10845
         Width           =   2295
      End
   End
   Begin VB.PictureBox Pic1 
      BorderStyle     =   0  'None
      Height          =   11400
      Left            =   2880
      ScaleHeight     =   11400
      ScaleWidth      =   7605
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   7605
      Begin VB.CommandButton CMD3 
         BackColor       =   &H000080FF&
         Caption         =   "Batal"
         Height          =   405
         Left            =   3600
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm108.frx":40C8
         MousePointer    =   99  'Custom
         TabIndex        =   33
         Top             =   840
         Width           =   2025
      End
      Begin VB.CommandButton CMD2 
         BackColor       =   &H000080FF&
         Caption         =   "Simpan Data"
         Height          =   405
         Left            =   1440
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm108.frx":43D2
         MousePointer    =   99  'Custom
         TabIndex        =   32
         Top             =   840
         Width           =   2025
      End
      Begin VB.CommandButton CMD1 
         BackColor       =   &H000080FF&
         Caption         =   "Simpan Data"
         Height          =   405
         Left            =   2520
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm108.frx":46DC
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   840
         Width           =   2025
      End
      Begin VB.CommandButton CMD4 
         BackColor       =   &H00FFFFFF&
         Height          =   740
         Left            =   5400
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm108.frx":49E6
         MousePointer    =   99  'Custom
         Picture         =   "Frm108.frx":4CF0
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Paparan sebelumnya"
         Top             =   10440
         Width           =   1000
      End
      Begin VB.CommandButton CMD5 
         BackColor       =   &H00FFFFFF&
         Height          =   740
         Left            =   6480
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm108.frx":562F
         MousePointer    =   99  'Custom
         Picture         =   "Frm108.frx":5939
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Paparan seterusnya"
         Top             =   10440
         Width           =   1000
      End
      Begin VB.TextBox TB4 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1560
         TabIndex        =   19
         Text            =   "TB4"
         Top             =   360
         Width           =   5700
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   8685
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   1680
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   15319
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
      Begin VB.Label L11_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L11_Text"
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
         Left            =   6720
         TabIndex        =   39
         Top             =   1320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label L10_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L10_Text"
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
         Height          =   240
         Left            =   1080
         TabIndex        =   35
         Top             =   10395
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan :"
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
         Height          =   240
         Left            =   240
         TabIndex        =   34
         Top             =   10395
         Width           =   975
      End
      Begin VB.Label L7_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L7_Text"
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
         Left            =   5160
         TabIndex        =   29
         Top             =   10920
         Width           =   615
      End
      Begin VB.Label L6_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L6_Text"
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
         Left            =   4635
         TabIndex        =   28
         Top             =   10920
         Width           =   375
      End
      Begin VB.Label L8_Text 
         Caption         =   "L8_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   480
         TabIndex        =   27
         Top             =   10680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L9_Text 
         Caption         =   "L9_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   480
         TabIndex        =   26
         Top             =   10920
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai cawangan / kedai / agen yang telah didaftarkan  di dalam sistem."
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Top             =   1440
         Width           =   8025
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan nama cawangan / kedai / agen bagi tujuan pendaftaran."
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   21
         Top             =   120
         Width           =   8025
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama *"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Top             =   390
         Width           =   1785
      End
      Begin VB.Label Label11 
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
         Left            =   3240
         TabIndex        =   30
         Top             =   10920
         Width           =   2295
      End
   End
   Begin VB.PictureBox Pic3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   11400
      Left            =   240
      ScaleHeight     =   11400
      ScaleWidth      =   21195
      TabIndex        =   60
      Top             =   1440
      Visible         =   0   'False
      Width           =   21195
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
         Left            =   3000
         TabIndex        =   151
         Top             =   360
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
         Left            =   1560
         TabIndex        =   150
         Top             =   360
         Width           =   200
      End
      Begin VB.CommandButton CMD13 
         BackColor       =   &H000080FF&
         Caption         =   "Carian Data"
         Height          =   360
         Left            =   3240
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm108.frx":625F
         MousePointer    =   99  'Custom
         TabIndex        =   88
         Top             =   5355
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox TB7 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1500
         TabIndex        =   87
         Text            =   "TB7"
         Top             =   5400
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.CommandButton CMD12 
         BackColor       =   &H000080FF&
         Caption         =   "Carian Data"
         Height          =   360
         Left            =   3240
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm108.frx":6569
         MousePointer    =   99  'Custom
         TabIndex        =   84
         Top             =   4035
         Width           =   1425
      End
      Begin VB.TextBox TB6 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1500
         TabIndex        =   83
         Text            =   "TB6"
         Top             =   4080
         Width           =   1620
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
         Left            =   285
         TabIndex        =   63
         Top             =   735
         Width           =   200
      End
      Begin VB.ComboBox CBB3 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   315
         ItemData        =   "Frm108.frx":6873
         Left            =   1500
         List            =   "Frm108.frx":6875
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   2145
         Width           =   4725
      End
      Begin VB.CommandButton CMD11 
         BackColor       =   &H000080FF&
         Caption         =   "Report"
         Height          =   405
         Left            =   2400
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm108.frx":6877
         MousePointer    =   99  'Custom
         TabIndex        =   61
         Top             =   2640
         Width           =   2025
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   360
         Left            =   1500
         TabIndex        =   64
         Top             =   1275
         Width           =   4725
         _ExtentX        =   8334
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
         Left            =   1500
         TabIndex        =   65
         Top             =   1635
         Width           =   4725
         _ExtentX        =   8334
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
      Begin VB.PictureBox Pic4 
         BorderStyle     =   0  'None
         Height          =   10815
         Left            =   7320
         ScaleHeight     =   10815
         ScaleWidth      =   14445
         TabIndex        =   71
         Top             =   -960
         Visible         =   0   'False
         Width           =   14445
         Begin VB.CommandButton CMD14 
            BackColor       =   &H00FFFFFF&
            Height          =   740
            Left            =   12240
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm108.frx":6B81
            MousePointer    =   99  'Custom
            Picture         =   "Frm108.frx":6E8B
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Paparan sebelumnya"
            Top             =   9960
            Width           =   1000
         End
         Begin VB.CommandButton CMD10 
            BackColor       =   &H00FFFFFF&
            Height          =   740
            Left            =   13320
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm108.frx":77CA
            MousePointer    =   99  'Custom
            Picture         =   "Frm108.frx":7AD4
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Paparan seterusnya"
            Top             =   9960
            Width           =   1000
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
            Height          =   9525
            Left            =   120
            TabIndex        =   74
            ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
            Top             =   360
            Width           =   14205
            _ExtentX        =   25056
            _ExtentY        =   16801
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
         Begin VB.Label L24_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L24_Text"
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
            Left            =   240
            TabIndex        =   80
            Top             =   120
            Width           =   13455
         End
         Begin VB.Label L27_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L27_Text"
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
            Left            =   11850
            TabIndex        =   79
            Top             =   9960
            Width           =   615
         End
         Begin VB.Label L26_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "L26_Text"
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
            Left            =   11340
            TabIndex        =   78
            Top             =   9960
            Width           =   375
         End
         Begin VB.Label L28_Text 
            BackColor       =   &H8000000C&
            Caption         =   "L28_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   8880
            TabIndex        =   77
            Top             =   9960
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label L29_Text 
            BackColor       =   &H8000000C&
            Caption         =   "L29_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   8880
            TabIndex        =   76
            Top             =   10200
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label L25_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L25_Text"
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
            Left            =   1080
            TabIndex        =   75
            Top             =   9960
            Width           =   975
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "Bilangan :   "
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
            Left            =   240
            TabIndex        =   82
            Top             =   9960
            Width           =   1815
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "Paparan Muka  :       / "
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
            Left            =   10080
            TabIndex        =   81
            Top             =   9960
            Width           =   2295
         End
      End
      Begin VB.PictureBox Pic5 
         BorderStyle     =   0  'None
         Height          =   10815
         Left            =   6720
         ScaleHeight     =   10815
         ScaleWidth      =   12765
         TabIndex        =   95
         Top             =   360
         Visible         =   0   'False
         Width           =   12765
         Begin VB.CommandButton CMD19 
            BackColor       =   &H000080FF&
            Caption         =   "Tutup Paparan Ini"
            Height          =   405
            Left            =   8400
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm108.frx":83FA
            MousePointer    =   99  'Custom
            TabIndex        =   108
            Top             =   10320
            Width           =   2025
         End
         Begin VB.CommandButton CMD16 
            BackColor       =   &H00FFFFFF&
            Height          =   740
            Left            =   11640
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm108.frx":8704
            MousePointer    =   99  'Custom
            Picture         =   "Frm108.frx":8A0E
            Style           =   1  'Graphical
            TabIndex        =   97
            ToolTipText     =   "Paparan seterusnya"
            Top             =   9960
            Width           =   1000
         End
         Begin VB.CommandButton CMD15 
            BackColor       =   &H00FFFFFF&
            Height          =   740
            Left            =   10560
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm108.frx":9334
            MousePointer    =   99  'Custom
            Picture         =   "Frm108.frx":963E
            Style           =   1  'Graphical
            TabIndex        =   96
            ToolTipText     =   "Paparan sebelumnya"
            Top             =   9960
            Width           =   1000
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
            Height          =   9525
            Left            =   120
            TabIndex        =   98
            ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
            Top             =   360
            Width           =   12525
            _ExtentX        =   22093
            _ExtentY        =   16801
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
         Begin VB.Label L62_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "Maklumat agihan"
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
            Height          =   240
            Left            =   240
            TabIndex        =   200
            Top             =   9960
            Width           =   2775
         End
         Begin VB.Label L66_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L66_Text"
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
            Height          =   240
            Left            =   4605
            TabIndex        =   199
            Top             =   10560
            Width           =   1440
         End
         Begin VB.Label L64_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L64_Text"
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
            Height          =   240
            Left            =   4605
            TabIndex        =   197
            Top             =   10170
            Width           =   975
         End
         Begin VB.Label L65_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L65_Text"
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
            Height          =   240
            Left            =   4605
            TabIndex        =   196
            Top             =   10350
            Width           =   1440
         End
         Begin VB.Label L32_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L32_Text"
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
            Height          =   240
            Left            =   1845
            TabIndex        =   107
            Top             =   10350
            Width           =   1440
         End
         Begin VB.Label L31_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L31_Text"
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
            Height          =   240
            Left            =   1850
            TabIndex        =   104
            Top             =   10170
            Width           =   975
         End
         Begin VB.Label L36_Text 
            BackColor       =   &H8000000C&
            Caption         =   "L36_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   7560
            TabIndex        =   103
            Top             =   10200
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label L35_Text 
            BackColor       =   &H8000000C&
            Caption         =   "L35_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   7560
            TabIndex        =   102
            Top             =   9960
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label L33_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "L33_Text"
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
            Left            =   9660
            TabIndex        =   101
            Top             =   9960
            Width           =   375
         End
         Begin VB.Label L34_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L34_Text"
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
            Left            =   10200
            TabIndex        =   100
            Top             =   9960
            Width           =   615
         End
         Begin VB.Label L30_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L30_Text"
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
            Left            =   240
            TabIndex        =   99
            Top             =   120
            Width           =   13455
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "Bilangan                 :                          Jumlah Berat (g) : "
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
            Height          =   900
            Left            =   240
            TabIndex        =   105
            Top             =   10170
            Width           =   2655
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Paparan Muka  :       / "
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
            Left            =   8400
            TabIndex        =   106
            Top             =   9960
            Width           =   2295
         End
         Begin VB.Label L63_Text 
            BackStyle       =   0  'Transparent
            Caption         =   $"Frm108.frx":9F7D
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
            Height          =   900
            Left            =   3000
            TabIndex        =   198
            Top             =   9960
            Width           =   2655
         End
      End
      Begin VB.Label L48_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L48_Text"
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
         Height          =   195
         Left            =   600
         TabIndex        =   152
         Top             =   8160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Report               Agihan                     Pulangan"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   149
         Top             =   340
         Width           =   4890
      End
      Begin VB.Label L37_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L37_Text"
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
         Height          =   195
         Left            =   600
         TabIndex        =   109
         Top             =   7800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L23_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L23_Text"
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
         Height          =   195
         Left            =   600
         TabIndex        =   94
         Top             =   7440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L22_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L22_Text"
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
         Height          =   195
         Left            =   600
         TabIndex        =   93
         Top             =   7080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L21_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L21_Text"
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
         Height          =   195
         Left            =   600
         TabIndex        =   92
         Top             =   6720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L20_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L20_Text"
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
         Height          =   195
         Left            =   600
         TabIndex        =   91
         Top             =   6360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Siri Produk   :"
         Height          =   255
         Left            =   240
         TabIndex        =   90
         Top             =   5400
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan No. Siri Produk bagi mencari data terperinci berkenaan agihan barang tersebut."
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   240
         TabIndex        =   89
         Top             =   4920
         Visible         =   0   'False
         Width           =   4530
      End
      Begin VB.Shape Shape4 
         Height          =   1215
         Left            =   120
         Top             =   4680
         Visible         =   0   'False
         Width           =   6495
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Rujukan      :"
         Height          =   255
         Left            =   240
         TabIndex        =   86
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan No. Rujukan bagi mencari data terperinci hantaran barang dari nombor rujukan tersebut."
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   240
         TabIndex        =   85
         Top             =   3600
         Width           =   4530
      End
      Begin VB.Shape Shape3 
         Height          =   1215
         Left            =   120
         Top             =   3360
         Width           =   6495
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila tandakan di sini jika ingin cari data mengikut tarikh."
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   570
         TabIndex        =   70
         Top             =   720
         Width           =   4890
      End
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Akhir "
         Height          =   255
         Left            =   240
         TabIndex        =   69
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label62 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Mula "
         Height          =   255
         Left            =   240
         TabIndex        =   68
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "** Jika tidak ditanda , sistem TIDAK akan mengeluarkan report mengikut tarikh."
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   600
         TabIndex        =   67
         Top             =   960
         Width           =   5850
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Cawangan"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   240
         TabIndex        =   66
         Top             =   2175
         Width           =   1695
      End
      Begin VB.Shape Shape2 
         Height          =   2655
         Left            =   120
         Top             =   600
         Width           =   6495
      End
   End
   Begin VB.PictureBox Pic2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   11400
      Left            =   960
      ScaleHeight     =   11400
      ScaleWidth      =   21045
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   21045
      Begin VB.CommandButton CMD18 
         BackColor       =   &H00FFFFFF&
         Height          =   740
         Left            =   18120
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm108.frx":A004
         MousePointer    =   99  'Custom
         Picture         =   "Frm108.frx":A30E
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Paparan seterusnya"
         Top             =   10365
         Width           =   1000
      End
      Begin VB.CommandButton CMD17 
         BackColor       =   &H00FFFFFF&
         Height          =   740
         Left            =   17040
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm108.frx":AC34
         MousePointer    =   99  'Custom
         Picture         =   "Frm108.frx":AF3E
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Paparan sebelumnya"
         Top             =   10365
         Width           =   1000
      End
      Begin VB.Timer Tmr1 
         Interval        =   100
         Left            =   120
         Top             =   120
      End
      Begin VB.CommandButton CMD9 
         BackColor       =   &H000080FF&
         Caption         =   "Carian Data"
         Height          =   405
         Left            =   5880
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm108.frx":B87D
         MousePointer    =   99  'Custom
         TabIndex        =   46
         Top             =   3660
         Width           =   2025
      End
      Begin VB.TextBox TB5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1900
         TabIndex        =   44
         Text            =   "TB5"
         Top             =   3720
         Width           =   3900
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
         Left            =   150
         TabIndex        =   40
         Top             =   3370
         Width           =   200
      End
      Begin VB.CommandButton CMD8 
         BackColor       =   &H000080FF&
         Caption         =   "Batal"
         Height          =   405
         Left            =   4440
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm108.frx":BB87
         MousePointer    =   99  'Custom
         TabIndex        =   38
         Top             =   10680
         Width           =   2025
      End
      Begin VB.CommandButton CMD7 
         BackColor       =   &H000080FF&
         Caption         =   "Simpan Data"
         Height          =   405
         Left            =   2280
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm108.frx":BE91
         MousePointer    =   99  'Custom
         TabIndex        =   37
         Top             =   10680
         Width           =   2025
      End
      Begin VB.CommandButton CMD6 
         BackColor       =   &H000080FF&
         Caption         =   "Simpan Data"
         Height          =   405
         Left            =   3360
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm108.frx":C19B
         MousePointer    =   99  'Custom
         TabIndex        =   36
         Top             =   10680
         Width           =   2025
      End
      Begin VB.ComboBox CBB2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   315
         ItemData        =   "Frm108.frx":C4A5
         Left            =   1900
         List            =   "Frm108.frx":C4A7
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2000
         Width           =   6420
      End
      Begin VB.ComboBox CBB1 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   315
         ItemData        =   "Frm108.frx":C4A9
         Left            =   1900
         List            =   "Frm108.frx":C4AB
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1680
         Width           =   6420
      End
      Begin VB.TextBox TB3 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1900
         TabIndex        =   11
         Text            =   "TB3"
         Top             =   1290
         Width           =   6420
      End
      Begin VB.TextBox TB2 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1900
         TabIndex        =   9
         Text            =   "TB2"
         Top             =   990
         Width           =   6420
      End
      Begin VB.TextBox TB1 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1900
         TabIndex        =   3
         Text            =   "TB1"
         Top             =   690
         Width           =   6420
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   1900
         TabIndex        =   15
         Top             =   2310
         Width           =   6405
         _ExtentX        =   11298
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   9645
         Left            =   8640
         TabIndex        =   48
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   600
         Width           =   10605
         _ExtentX        =   18706
         _ExtentY        =   17013
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
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "*** Sila tekan F2 untuk scan barang yang akan diagihkan. (Jika menggunakan scanner mode)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   120
         TabIndex        =   184
         Top             =   4320
         Width           =   8265
      End
      Begin VB.Label L1_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L1_Text"
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
         Height          =   195
         Left            =   960
         TabIndex        =   59
         Top             =   9000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L18_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L18_Text"
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
         Height          =   240
         Left            =   11370
         TabIndex        =   58
         Top             =   10320
         Width           =   1335
      End
      Begin VB.Label L16_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L16_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   8640
         TabIndex        =   56
         Top             =   10845
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L15_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L15_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   8640
         TabIndex        =   55
         Top             =   10605
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L13_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L13_Text"
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
         Left            =   16275
         TabIndex        =   54
         Top             =   10845
         Width           =   375
      End
      Begin VB.Label L14_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L14_Text"
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
         Left            =   16800
         TabIndex        =   53
         Top             =   10845
         Width           =   615
      End
      Begin VB.Label L17_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L17_Text"
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
         Height          =   240
         Left            =   9480
         TabIndex        =   51
         Top             =   10320
         Width           =   615
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai barang yang akan / telah dihantar kepada cawangan."
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
         Height          =   195
         Left            =   8760
         TabIndex        =   47
         Top             =   360
         Width           =   6000
      End
      Begin VB.Shape Shape1 
         Height          =   975
         Left            =   75
         Top             =   3240
         Width           =   8175
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Siri Produk *"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   150
         TabIndex        =   45
         Top             =   3750
         Width           =   1785
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan data yang akan dihantar/diagihkan kepada cawangan / kedai / agen."
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
         Height          =   420
         Left            =   120
         TabIndex        =   43
         Top             =   3000
         Width           =   7905
      End
      Begin VB.Label L12_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L12_Text"
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
         Height          =   195
         Left            =   960
         TabIndex        =   42
         Top             =   9240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Scanner Mode (Sila klik sini jika anda menggunakan scanner untuk scan data barang kemas)"
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   435
         TabIndex        =   41
         Top             =   3360
         Width           =   7665
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pekerja *"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   18
         Top             =   2030
         Width           =   1695
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Cawangan / Kedai *"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   14
         Top             =   1710
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Telefon"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   1300
         Width           =   1785
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Kad Pengenalan *"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   990
         Width           =   1785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan maklumat terperinci berkenaan cawangan / agen / pengedar yang akan mengambil barang ini."
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
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   8145
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama *"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1785
      End
      Begin VB.Label Label19 
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
         Left            =   14880
         TabIndex        =   57
         Top             =   10845
         Width           =   2295
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan :              Jumlah Berat :"
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
         Height          =   240
         Left            =   8640
         TabIndex        =   52
         Top             =   10320
         Width           =   3255
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Hantar *"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   2350
         Width           =   2385
      End
   End
   Begin VB.Label L59_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Report Inventori"
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
      Left            =   8760
      MouseIcon       =   "Frm108.frx":C4AD
      MousePointer    =   99  'Custom
      TabIndex        =   181
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label L38_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pulangan Barang"
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
      Left            =   4680
      MouseIcon       =   "Frm108.frx":C7B7
      MousePointer    =   99  'Custom
      TabIndex        =   110
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label L4_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rekod Agihan && Pulangan"
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
      MouseIcon       =   "Frm108.frx":CAC1
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label L2_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai Cawangan / Kedai / Agen"
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
      MouseIcon       =   "Frm108.frx":CDCB
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label L3_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hantaran Barang"
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
      Left            =   3120
      MouseIcon       =   "Frm108.frx":D0D5
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Menu Frm108_PM_menu1 
      Caption         =   "Menu 1"
      Visible         =   0   'False
      Begin VB.Menu Frm108_SM_edit_data 
         Caption         =   "Edit data ini"
      End
      Begin VB.Menu Frm108_SM_tukar_status 
         Caption         =   "Tukar status"
         Begin VB.Menu Frm108_SM_tidak_aktif 
            Caption         =   "Tidak aktif"
         End
      End
   End
   Begin VB.Menu Frm108_PM_menu2 
      Caption         =   "Scan Barang F2"
      Begin VB.Menu Frm108_SM_scan 
         Caption         =   "Scan barang"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu Frm108_PM_menu3 
      Caption         =   "Menu 2"
      Visible         =   0   'False
      Begin VB.Menu Frm108_SM_lihat_edit_data 
         Caption         =   "Lihat / edit data ini"
      End
      Begin VB.Menu Frm108_SM_padam_data 
         Caption         =   "Padam data"
      End
      Begin VB.Menu Frm108_SM_senarai_barang 
         Caption         =   "Senarai barang terperinci"
      End
      Begin VB.Menu Frm108_SM_cetak_penyata 
         Caption         =   "Cetak penyata bagi nombor rujukan ini"
      End
   End
   Begin VB.Menu Frm108_PM_menu4 
      Caption         =   "Menu 3"
      Visible         =   0   'False
      Begin VB.Menu Frm108_SM_remove 
         Caption         =   "Keluarkan dari senarai"
      End
   End
   Begin VB.Menu Frm108_PM_menu5 
      Caption         =   "Menu 4"
      Visible         =   0   'False
      Begin VB.Menu Frm108_SM_remove2 
         Caption         =   "Keluarkan dari senarai"
      End
   End
   Begin VB.Menu Frm108_PM_menu6 
      Caption         =   "Menu 5"
      Visible         =   0   'False
      Begin VB.Menu Frm108_SM_excel 
         Caption         =   "Export report ke dalam format EXCEL"
      End
   End
End
Attribute VB_Name = "Frm108"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB4_Click()
'on error resume next
If Frm108.CB4 = 1 Then
    Frm108.CB5 = 0
End If
End Sub
Private Sub CB5_Click()
'on error resume next
If Frm108.CB5 = 1 Then
    Frm108.CB4 = 0
End If
End Sub
Private Sub CB7_Click()
'on error resume next
If Frm108.CB7 = 1 Then
    Frm108.CB8 = 0
End If
End Sub
Private Sub CB8_Click()
'on error resume next
If Frm108.CB8 = 1 Then
    Frm108.CB7 = 0
End If
End Sub
Private Sub CMD1_Click()
'on error resume next
DATA_FOUND = 0

If Frm108.TB4 = vbNullString Then
    MsgBox "Sila masukkan [Cawangan].", vbInformation, "Info"
    Frm108.TB4.SetFocus
    Exit Sub
End If

If InStr(1, Frm108.TB4, "*") <> 0 Or InStr(1, Frm108.TB4, "/") <> 0 Or InStr(1, Frm108.TB4, "\") <> 0 Or InStr(1, Frm108.TB4, "'") <> 0 Then
    MsgBox "Cawangan mengandungi simbol. Sila buang simbol dan cuba sekali lagi.", vbExclamation, "Error"
    
    Frm108.TB4 = vbNullString
    Frm108.TB4.SetFocus
    Exit Sub
End If

Note = "Adakah anda ingin simpan data ini?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 62_senarai_cawangan where cawangan='" & UCase(Frm108.TB4) & "'", cn, adOpenKeyset, adLockOptimistic

    If rs.EOF Then
        
        rs.AddNew
        rs!cawangan = UCase(Frm108.TB4) 'Nama cawangan
        rs!Status = 1
        rs!write_timestamp = Now
        rs.Update
        
        DATA_FOUND = 1
    Else
        If rs!Status = 1 Then
        
            MsgBox "Cawangan [" & UCase(Frm108.TB4) & "] telah didaftarkan sebelum ini. Sila periksa data anda.", vbexclamtion, "Info"
            
            Frm108.TB4 = vbNullString
            Frm108.TB4.SetFocus
            
        ElseIf rs!Status = 0 Then
            rs!Status = 1
            
            rs.Update
            
            DATA_FOUND = 2
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    
    If DATA_FOUND = 1 Then
    
        '#### Update Log Aktiviti Sistem #### - Start
        user = MDI_frm1.L3_Text
        
        LogAct_Memory = "[" & user & "] Daftar cawangan. Nama cawangan [" & UCase(Frm108.TB4) & "]."
        LogDate_Memory = DateTime.Date & " " & DateTime.Time$
        Call UpdateLog_Database
        '#### Update Log Aktiviti Sistem #### - End
    
        GM_NEXT_PREV = 2
        
        Call Frm108_cawangan_initial_setting
        Call Frm108_senarai_cawangan_header
        Call Frm108_senarai_cawangan
    
        MsgBox "Data cawangan telah berjaya disimpan.", vbInformation, "Info"
        
        Frm108.TB4.SetFocus
    ElseIf DATA_FOUND = 2 Then
    
    
'### Update nama cawangan dalam table #64_agihan_barang ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

        strsql = "UPDATE 64_agihan_barang set cawangan='" & UCase(Frm108.TB4) & "'" _
        & "WHERE cawangan_id='" & Frm108.L11_Text & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update nama cawangan dalam table #64_agihan_barang ### - End
    
        '#### Update Log Aktiviti Sistem #### - Start
        user = MDI_frm1.L3_Text
        
        LogAct_Memory = "[" & user & "] Pengaktifan semula cawangan. Nama cawangan [" & UCase(Frm108.TB4) & "]."
        LogDate_Memory = DateTime.Date & " " & DateTime.Time$
        Call UpdateLog_Database
        '#### Update Log Aktiviti Sistem #### - End
    
        GM_NEXT_PREV = 2
        
        Call Frm108_cawangan_initial_setting
        Call Frm108_senarai_cawangan_header
        Call Frm108_senarai_cawangan
        
        MsgBox "Cawangan " & UCase(Frm108.TB4) & " telah berjaya diaktifkan kembali.", vbInformation, "Info"
        Frm108.TB4.SetFocus
    End If
    
End If
End Sub
Private Sub CMD10_Click()
'on error resume next
Dim Frm108_LM_CURR_PAGE As Double
Dim Frm108_LM_TOTAL_PAGE As Double

Frm108_LM_CURR_PAGE = 0
Frm108_LM_TOTAL_PAGE = 0

If Frm108.L26_Text <> vbNullString And IsNumeric(Frm108.L26_Text) Then
    If Frm108.L27_Text <> vbNullString And IsNumeric(Frm108.L27_Text) Then
        Frm108_LM_CURR_PAGE = Frm108.L26_Text
        Frm108_LM_TOTAL_PAGE = Frm108.L27_Text
        
        If Frm108_LM_CURR_PAGE < Frm108_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            If Frm108.L48_Text = 0 Then 'Memory : Jenis report , 0 : Agihan , 1 : Pulangan
                Call Frm108_senarai_agihan_barang_header
                Call Frm108_senarai_agihan_barang
            ElseIf Frm108.L48_Text = 1 Then 'Memory : Jenis report , 0 : Agihan , 1 : Pulangan
                Call Frm108_senarai_pulangan_barang_header
                Call Frm108_senarai_pulangan_barang
            End If
            
        End If
    End If
End If
End Sub
Private Sub CMD11_Click()
'on error resume next
If Frm108.CBB3 = vbNullString Then
    MsgBox "Sila pilih cawangan", vbInformation, "Info"
    
    Exit Sub
End If
If Frm108.CB4 = 0 And Frm108.CB5 = 0 Then
    MsgBox "Sila buat pilihan report", vbInformation, "Info"
    
    Exit Sub
End If

If Frm108.CB4 = 1 Then
    Note = "REPORT AGIHAN" & vbCrLf & _
            vbNullString & vbCrLf & _
            "Sistem mungkin akan mengambil masa untuk mengeluarkan report." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan ?"
End If

If Frm108.CB5 = 1 Then
    Note = "REPORT PULANGAN" & vbCrLf & _
            vbNullString & vbCrLf & _
            "Sistem mungkin akan mengambil masa untuk mengeluarkan report." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan ?"
End If

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then
    
    If Frm108.CB2 = 0 Then
        Frm108.L20_Text = 0 'Memori : Jenis report ( 0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh , 2:  No.rujukan , 3:  No.siri produk )
        
        If Frm108.CBB3 = "Semua cawangan" Then
            If Frm108.CB4 = 1 Then Frm108.L24_Text = "Senarai statement agihan barang kepada cawangan." 'Header
            If Frm108.CB5 = 1 Then Frm108.L24_Text = "Senarai statement pulangan barang oleh cawangan." 'Header
        Else
            If Frm108.CB4 = 1 Then Frm108.L24_Text = "Senarai statement agihan barang kepada cawangan [" & Frm108.CBB3 & "]." 'Header
            If Frm108.CB5 = 1 Then Frm108.L24_Text = "Senarai statement pulangan barang oleh cawangan [" & Frm108.CBB3 & "]." 'Header
        End If
        
    Else
        Frm108.L20_Text = 1 'Memori : Jenis report ( 0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh , 2:  No.rujukan , 3:  No.siri produk )
    
        If Frm108.CBB3 = "Semua cawangan" Then
            If Frm108.CB4 = 1 Then Frm108.L24_Text = "Senarai statement agihan barang kepada cawangan dari " & Frm108.DTPicker2 & " hingga " & Frm108.DTPicker3 & "." 'Header
            If Frm108.CB5 = 1 Then Frm108.L24_Text = "Senarai statement pulangan barang oleh cawangan dari " & Frm108.DTPicker2 & " hingga " & Frm108.DTPicker3 & "." 'Header
        Else
            If Frm108.CB4 = 1 Then Frm108.L24_Text = "Senarai statement agihan barang kepada cawangan [" & Frm108.CBB3 & "] dari " & Frm108.DTPicker2 & " hingga " & Frm108.DTPicker3 & "." 'Header
            If Frm108.CB5 = 1 Then Frm108.L24_Text = "Senarai statement pulangan barang oleh cawangan [" & Frm108.CBB3 & "] dari " & Frm108.DTPicker2 & " hingga " & Frm108.DTPicker3 & "." 'Header
        End If
        
    End If
    
    Call Frm108_initial_setting2

    Frm108.L21_Text = Frm108.DTPicker2 'Memori : Tarikh mula
    Frm108.L22_Text = Frm108.DTPicker3 'Memori : Tarikh akhir
    Frm108.L23_Text = Frm108.CBB3 'Memori : Supplier / No rujukan / No. siri produk
    
    GM_NEXT_PREV = 0
    
    Frm108.L28_Text = -1 'Titik Pencarian Data
    Frm108.L29_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm108.L26_Text = 0 'Paparan Page ke-xxx
    
    'Frm108.L35_Text = -1 'Titik Pencarian Data
    'Frm108.L36_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    'Frm108.L33_Text = 0 'Paparan Page ke-xxx
    
    'Call Frm108_senarai_agihan_barang_detail_header
    'Call Frm108_senarai_agihan_barang_detail
    
    If Frm108.CB4 = 1 Then
        Frm108.L48_Text = 0 'Memory : Jenis report , 0 : Agihan , 1 : Pulangan
        Call Frm108_senarai_agihan_barang_header
        Call Frm108_senarai_agihan_barang
    End If
    
    If Frm108.CB5 = 1 Then
        Frm108.L48_Text = 1 'Memory : Jenis report , 0 : Agihan , 1 : Pulangan
        Call Frm108_senarai_pulangan_barang_header
        Call Frm108_senarai_pulangan_barang
    End If
    
    Frm108.Pic4.Visible = True
    
    If Frm108.L25_Text <> vbNullString Then
        If Frm108.L25_Text = 0 Then MsgBox "Tiada rekod dijumpai.", vbInformation, "Info"
    End If
    
End If
End Sub
Private Sub CMD12_Click()
'on error resume next
If Frm108.CB4 = 0 And Frm108.CB5 = 0 Then
    MsgBox "Sila buat pilihan report", vbInformation, "Info"
    
    Exit Sub
End If

If Frm108.TB6 = vbNullString Then
    MsgBox "Sila masukkan No. Rujukan.", vbInformation, "Info"
    
    Frm108.TB6.SetFocus
    Exit Sub
End If

If InStr(1, Frm108.TB6, "*") <> 0 Or InStr(1, Frm108.TB6, "/") <> 0 Or InStr(1, Frm108.TB6, "\") <> 0 Or InStr(1, Frm108.TB6, "'") <> 0 Then
    MsgBox "Simbol tidak dibenarkan di dalam ruangan No. Rujukan.", vbExclamation, "Error"
    
    Frm108.TB6 = vbNullString
    Frm108.TB6.SetFocus
    Exit Sub
End If

If Frm108.CB4 = 1 Then
    Note = "REPORT AGIHAN" & vbCrLf & _
            vbNullString & vbCrLf & _
            "Sistem mungkin akan mengambil masa untuk mengeluarkan report." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan ?"
End If

If Frm108.CB5 = 1 Then
    Note = "REPORT PULANGAN" & vbCrLf & _
            vbNullString & vbCrLf & _
            "Sistem mungkin akan mengambil masa untuk mengeluarkan report." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan ?"
End If

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then

    Frm108.L20_Text = 2 'Memori : Jenis report ( 0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh , 2:  No.rujukan , 3:  No.siri produk )

    If Frm108.CB4 = 1 Then Frm108.L24_Text = "Senarai statement agihan barang kepada cawangan dari No. Rujukan [" & UCase(Frm108.TB6) & "]."  'Header
    If Frm108.CB5 = 1 Then Frm108.L24_Text = "Senarai statement pulangan barang oleh cawangan dari No. Rujukan [" & UCase(Frm108.TB6) & "]."  'Header
    
    Call Frm108_initial_setting2

    Frm108.L23_Text = UCase(Frm108.TB6) 'Memori : Supplier / No rujukan / No. siri produk
    
    GM_NEXT_PREV = 0
    
    Frm108.L28_Text = -1 'Titik Pencarian Data
    Frm108.L29_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm108.L26_Text = 0 'Paparan Page ke-xxx
    
    If Frm108.CB4 = 1 Then
        Frm108.L48_Text = 0 'Memory : Jenis report , 0 : Agihan , 1 : Pulangan
        Call Frm108_senarai_agihan_barang_header
        Call Frm108_senarai_agihan_barang
    End If
    
    If Frm108.CB5 = 1 Then
        Frm108.L48_Text = 1 'Memory : Jenis report , 0 : Agihan , 1 : Pulangan
        Call Frm108_senarai_pulangan_barang_header
        Call Frm108_senarai_pulangan_barang
    End If
    
    Frm108.Pic4.Visible = True
    
    If Frm108.L25_Text <> vbNullString Then
        If Frm108.L25_Text = 0 Then MsgBox "Tiada rekod dijumpai.", vbInformation, "Info"
    End If
    
End If
End Sub
Private Sub CMD14_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

If Frm108.L48_Text = 0 Then 'Memory : Jenis report , 0 : Agihan , 1 : Pulangan
    Call Frm108_senarai_agihan_barang_header
    Call Frm108_senarai_agihan_barang
ElseIf Frm108.L48_Text = 1 Then 'Memory : Jenis report , 0 : Agihan , 1 : Pulangan
    Call Frm108_senarai_pulangan_barang_header
    Call Frm108_senarai_pulangan_barang
End If
End Sub
Private Sub CMD15_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

If Frm108.L48_Text = 0 Then
    Call Frm108_senarai_agihan_barang_detail_header
    Call Frm108_senarai_agihan_barang_detail
End If

If Frm108.L48_Text = 1 Then
    Call Frm108_senarai_pulangan_barang_detail_header
    Call Frm108_senarai_pulangan_barang_detail
End If
End Sub
Private Sub CMD16_Click()
'on error resume next
Dim Frm108_LM_CURR_PAGE As Double
Dim Frm108_LM_TOTAL_PAGE As Double

Frm108_LM_CURR_PAGE = 0
Frm108_LM_TOTAL_PAGE = 0

If Frm108.L33_Text <> vbNullString And IsNumeric(Frm108.L33_Text) Then
    If Frm108.L34_Text <> vbNullString And IsNumeric(Frm108.L34_Text) Then
        Frm108_LM_CURR_PAGE = Frm108.L33_Text
        Frm108_LM_TOTAL_PAGE = Frm108.L34_Text
        
        If Frm108_LM_CURR_PAGE < Frm108_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            If Frm108.L48_Text = 0 Then
                Call Frm108_senarai_agihan_barang_detail_header
                Call Frm108_senarai_agihan_barang_detail
            End If
            
            If Frm108.L48_Text = 1 Then
                Call Frm108_senarai_pulangan_barang_detail_header
                Call Frm108_senarai_pulangan_barang_detail
            End If

        End If
    End If
End If
End Sub
Private Sub CMD17_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm108_senarai_agihan_header
Call Frm108_senarai_agihan
End Sub
Private Sub CMD18_Click()
'on error resume next
Dim Frm108_LM_CURR_PAGE As Double
Dim Frm108_LM_TOTAL_PAGE As Double

Frm108_LM_CURR_PAGE = 0
Frm108_LM_TOTAL_PAGE = 0

If Frm108.L13_Text <> vbNullString And IsNumeric(Frm108.L13_Text) Then
    If Frm108.L14_Text <> vbNullString And IsNumeric(Frm108.L14_Text) Then
        Frm108_LM_CURR_PAGE = Frm108.L13_Text
        Frm108_LM_TOTAL_PAGE = Frm108.L14_Text
        
        If Frm108_LM_CURR_PAGE < Frm108_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm108_senarai_agihan_header
            Call Frm108_senarai_agihan
            
        End If
    End If
End If
End Sub
Private Sub CMD19_Click()
'on error resume next
Frm108.Pic4.Visible = True
Frm108.Pic5.Visible = False
End Sub
Private Sub CMD2_Click()
'on error resume next
DATA_FOUND = 0
DATA_WRITE = 1 '0 : Data Write NG , Data Write OK

If Frm108.TB4 = vbNullString Then
    MsgBox "Sila masukkan [Cawangan].", vbInformation, "Info"
    Frm108.TB4.SetFocus
    Exit Sub
End If

If InStr(1, Frm108.TB4, "*") <> 0 Or InStr(1, Frm108.TB4, "/") <> 0 Or InStr(1, Frm108.TB4, "\") <> 0 Or InStr(1, Frm108.TB4, "'") <> 0 Then
    MsgBox "Cawangan mengandungi simbol. Sila buang simbol dan cuba sekali lagi.", vbExclamation, "Error"
    
    'Frm108.TB4 = vbNullString
    Frm108.TB4.SetFocus
    Exit Sub
End If

Note = "Adakah anda ingin simpan data ini?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then

'### Periksa samada nama yang telah diedit ini telah digunakan atau belum ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 62_senarai_cawangan where cawangan='" & UCase(Frm108.TB4) & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        If rs!ID <> Frm108.L11_Text Then
            MsgBox "Cawangan [" & UCase(Frm108.TB4) & "] telah didaftarkan sebelum ini. Sila periksa data anda.", vbexclamtion, "Info"
            
            Frm108.TB4.SetFocus
            DATA_WRITE = 0 '0 : Data Write NG , Data Write OK
        End If
        
    End If
    
    rs.Close
    Set rs = Nothing
'### Periksa samada nama yang telah diedit ini telah digunakan atau belum ### - End
    
    If DATA_WRITE = 1 Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 62_senarai_cawangan where ID='" & Frm108.L11_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
        If Not rs.EOF Then
                
            rs!cawangan = UCase(Frm108.TB4) 'Nama cawangan
            rs!write_timestamp2 = Now
            rs.Update
            DATA_FOUND = 1
            
        End If
        
        rs.Close
        Set rs = Nothing
    
        If DATA_FOUND = 1 Then
        
'### Update nama cawangan dalam table #63_agihan ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    
            strsql = "UPDATE 63_agihan set cawangan='" & UCase(Frm108.TB4) & "'" _
            & "WHERE cawangan_id='" & Frm108.L11_Text & "'"
            
            Set rs = cn.Execute(strsql)
            Set rs = Nothing
'### Update nama cawangan dalam table #63_agihan ### - End

'### Update nama cawangan dalam table #64_agihan_barang ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    
            strsql = "UPDATE 64_agihan_barang set cawangan='" & UCase(Frm108.TB4) & "'" _
            & "WHERE cawangan_id='" & Frm108.L11_Text & "'"
            
            Set rs = cn.Execute(strsql)
            Set rs = Nothing
'### Update nama cawangan dalam table #64_agihan_barang ### - End
        
'#### Update Log Aktiviti Sistem #### - Start
            user = MDI_frm1.L3_Text
            
            LogAct_Memory = "[" & user & "] Edit data cawangan. Nama cawangan [" & UCase(Frm108.TB4) & "]."
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End
        
            GM_NEXT_PREV = 2
            
            Call Frm108_cmd_visible_1
            Call Frm108_cawangan_initial_setting
            Call Frm108_senarai_cawangan_header
            Call Frm108_senarai_cawangan
        
            MsgBox "Data cawangan telah berjaya disimpan.", vbInformation, "Info"
            
            Frm108.TB4.SetFocus
            
        End If
    End If
End If
End Sub
Private Sub CMD20_Click()
'on error resume next
Dim Err(6)

If InStr(1, Frm108.TB11, "*") <> 0 Or InStr(1, Frm108.TB11, "/") <> 0 Or InStr(1, Frm108.TB11, "\") <> 0 Or InStr(1, Frm108.TB11, "'") <> 0 Then

    MsgBox "No. Siri Produk mengandungi simbol yang tidak sah.", vbInformation, "Info"
    
    Frm108.TB11 = vbNullString
    Exit Sub
End If

If Frm108.CB7 = 0 And Frm108.CB8 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan jenis pulangan samada [Pulangan] atau [Dijual]"
End If

If Frm108.CB8 = 1 Then

    If Frm108.TB12 = vbNullString Then
        x = x + 1
        Err(x) = "Sila masukkan [No. Perjanjian A]."
    End If
    If Frm108.TB13 = vbNullString Then
        x = x + 1
        Err(x) = "Sila masukkan [No. Perjanjian B]."
    End If
    If Frm108.TB14 = vbNullString Or (Frm108.TB14 <> vbNullString And Not IsNumeric(Frm108.TB14)) Then
        x = x + 1
        Err(x) = "Sila masukkan Harga Jualan. Hanya NOMBOR dibenarkan di dalam ruangan ini."
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

    Call Frm108_periksa_data_barang2

End If
End Sub
Private Sub CMD21_Click()
'on error resume next
Dim Err(6)
Dim Frm108_LM_No_RUJ As Integer

G_PENYATA_PULANGAN = vbNullString
            
Frm108_LM_CAW_ID = vbNullString
Frm108_LM_No_RUJ = 1

If Frm108.TB8 = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan [Nama]."
End If
If Frm108.TB9 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [No. Kad Pengenalan]."
End If
If Frm108.CBB4 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Cawangan / Kedai]."
End If
If Frm108.CBB5 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Nama Pekerja]."
End If

If Frm108.L42_Text <> vbNullString Then
    If Not IsNumeric(Frm108.L42_Text) Then
        x = x + 1
        Err(x) = "Tiada maklumat barang yang akan dipulangkan." & vbCrLf & _
                "Sila masukkan data barangan yang hendak dipulangkan atau keluar dari menu ini dan cuba sekali lagi."
    Else
        If Frm108.L42_Text = 0 Then
            x = x + 1
            Err(x) = "Tiada maklumat barang yang akan dipulangkan." & vbCrLf & _
                    "Sila masukkan data barangan yang hendak dipulangkan atau keluar dari menu ini dan cuba sekali lagi."
        End If
    End If
End If

If Frm108.L40_Text <> vbNullString Then
    If Not IsNumeric(Frm108.L40_Text) Then
        x = x + 1
        Err(x) = "Technical Error." & vbCrLf & _
                "Sila keluar dari menu ini dan cuba sekali lagi."
    End If
Else

    x = x + 1
    Err(x) = "Technical Error." & vbCrLf & _
            "Sila keluar dari menu ini dan cuba sekali lagi."

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
    
    If Answer = vbYes Then
    
'### Periksa nombor rujukan ### - Start
        Frm108_LM_No_RUJ = Frm108.L40_Text
        
'---------------------------------------No. Invoice
        LM_NOW = Now
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        rs.Open "select * from 10_rujukan_pulangan", cn2, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        rs!tarikh = Frm108.DTPicker4
        rs!terminal = G_TERMINAL
        rs!write_timestamp = LM_NOW
        rs!Status = 1
        rs!nama_staff = MDI_frm1.L3_Text
        rs.Update
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        rs.Open "select * from 10_rujukan_pulangan where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & Frm108.DTPicker4 & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!ID) Then

                Frm108_LM_No_RUJ = rs!ID 'No. Rujukan Belian
                rs!no_rujukan = "BRS" & Format(Frm108_LM_No_RUJ, "000000")
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

        GoTo a:
        
Re_Gen_No:
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 68_pulangan where no_rujukan='" & Frm108_LM_No_RUJ & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            Frm108_LM_No_RUJ = Frm108_LM_No_RUJ + 1
            Frm108.L40_Text = Frm108_LM_No_RUJ
            
            rs.Close
            Set rs = Nothing
            
            GoTo Re_Gen_No:
        End If
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 69_pulangan_barang where no_rujukan='" & Frm108_LM_No_RUJ & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            Frm108_LM_No_RUJ = Frm108_LM_No_RUJ + 1
            Frm108.L40_Text = Frm108_LM_No_RUJ
            
            rs.Close
            Set rs = Nothing
            
            GoTo Re_Gen_No:
        End If
        
        rs.Close
        Set rs = Nothing
'### Periksa nombor rujukan ### - End

a:

'### Carian No. ID cawangan ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 62_senarai_cawangan where cawangan='" & Frm108.CBB4 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!ID) Then Frm108_LM_CAW_ID = rs!ID
        End If
        
        rs.Close
        Set rs = Nothing
'### Carian No. ID cawangan ### - End

'### No Rujukan pekerja ### - Start
        If Frm108.CBB5 <> vbNullString Then
            Frm108_LM_EMP_NO = Split(Frm108.CBB5, "  |  ")(1)
        End If
'### No Rujukan pekerja ### - End

'### Masukkan data asas pulangan ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 68_pulangan", cn, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        If Frm108.L40_Text <> vbNullString Then 'No. rujukan sistem
            rs!no_rujukan = Frm108_LM_No_RUJ
        Else
            rs!no_rujukan = Null
        End If
        If Frm108.L40_Text <> vbNullString Then 'No. rujukan sistem
            rs!no_statement = "BRS" & Format(Frm108_LM_No_RUJ, "000000")
            G_PENYATA_PULANGAN = "BRS" & Format(Frm108_LM_No_RUJ, "000000")
        Else
            rs!no_statement = Null
        End If
        rs!tarikh = Frm108.DTPicker4 'Tarikh barang dipulangkan
        If Frm108.TB8 <> vbNullString Then 'Nama PIC
            rs!Nama = UCase(Frm108.TB8)
        Else
            rs!Nama = Null
        End If
        If Frm108.TB9 <> vbNullString Then 'No. IC
            rs!no_ic = UCase(Frm108.TB9)
        Else
            rs!no_ic = Null
        End If
        If Frm108.TB10 <> vbNullString Then 'No. telefon
            rs!no_tel = UCase(Frm108.TB10)
        Else
            rs!no_tel = Null
        End If
        If Frm108.CBB4 <> vbNullString Then 'Cawangan
            rs!cawangan = Frm108.CBB4
        Else
            rs!cawangan = Null
        End If
        If Frm108_LM_CAW_ID <> vbNullString Then 'No ID cawangan (dari table #62_senarai_cawangan)
            rs!cawangan_id = Frm108_LM_CAW_ID
        Else
            rs!cawangan_id = Null
        End If
        If Frm108.CBB5 <> vbNullString Then 'Nama pekerja yang daftarkan pulangan barang
            rs!nama_pekerja = Frm108_LM_EMP_NO
        Else
            rs!nama_pekerja = Null
        End If
        rs!Status = 1
        rs!write_timestamp = LM_NOW
        rs.Update
        
        rs.Close
        Set rs = Nothing
'### Masukkan data asas pulangan ### - End

'### Masukkan data di bawah ke dalam #67_pulangan_barang_temp ### - Start
'No rujukan
'Tarikh
'ID cawangan
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE " & G_PULANGAN_TEMP & " set no_rujukan='" & Frm108_LM_No_RUJ & "'," _
        & "tarikh='" & Frm108.DTPicker4 & "'," _
        & "cawangan_id='" & Frm108_LM_CAW_ID & "'" _
        & "WHERE status='" & 1 & "' OR status='" & 2 & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Masukkan data di bawah ke dalam #67_pulangan_barang_temp ### - End

'### #67_pulangan_barang_temp -> #69_pulangan_barang ### - Start (Barang yang dipulangkan)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "insert into 69_pulangan_barang(no_rujukan,no_rujukan_agihan,tarikh,cawangan_id,no_siri_produk,kategori_produk,purity,berat,status,write_timestamp)" & _
                    "select no_rujukan,no_rujukan_agihan,tarikh,cawangan_id,no_siri_produk,kategori_produk,purity,berat,1,'" & LM_NOW & "' from " & G_PULANGAN_TEMP & " WHERE status='" & 1 & "' order by no_siri_Produk ASC"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### #67_pulangan_barang_temp -> #69_pulangan_barang ### - End (Barang yang dipulangkan)

'### #67_pulangan_barang_temp -> #69_pulangan_barang ### - Start (Barang yang dijual)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "insert into 69_pulangan_barang(no_rujukan,no_rujukan_agihan,tarikh,cawangan_id,no_siri_produk,kategori_produk,purity,berat,no_perjanjian_a,no_perjanjian_b,harga_jualan,status,write_timestamp)" & _
                    "select no_rujukan,no_rujukan_agihan,tarikh,cawangan_id,no_siri_produk,kategori_produk,purity,berat,no_perjanjian_a,no_perjanjian_b,harga_jualan,2,'" & LM_NOW & "' from " & G_PULANGAN_TEMP & " WHERE status='" & 2 & "' order by no_siri_Produk ASC"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### #67_pulangan_barang_temp -> #69_pulangan_barang ### - End (Barang yang dijual)

'### Update status dan tarikh pulangan dalam table #64_agihan_barang ### - Start (Barang yang dipulangkan)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        'strsql = "UPDATE 64_agihan_barang set status='" & 3 & "'," _
        & "tarikh_jual='" & Frm108.DTPicker4 & "'," _
        & "write_timestamp3='" & Now & "'" _
        & "WHERE status='" & 1 & "'"
        
        strsql = "UPDATE 64_agihan_barang," & G_PULANGAN_TEMP & " SET 64_agihan_barang.status='" & 2 & "'," _
        & "64_agihan_barang.tarikh_jual='" & Frm108.DTPicker4 & "'," _
        & "64_agihan_barang.write_timestamp3='" & LM_NOW & "'" _
        & "WHERE 64_agihan_barang.no_siri_produk = " & G_PULANGAN_TEMP & ".no_siri_produk AND " & G_PULANGAN_TEMP & ".status='" & 1 & "' AND 64_agihan_barang.no_rujukan = " & G_PULANGAN_TEMP & ".no_rujukan_agihan"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update status dan tarikh pulangan dalam table #64_agihan_barang ### - End (Barang yang dipulangkan)

'### Update status dan tarikh pulangan dalam table #64_agihan_barang ### - Start (Barang yang dijual)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        'strsql = "UPDATE 64_agihan_barang set status='" & 3 & "'," _
        & "tarikh_jual='" & Frm108.DTPicker4 & "'," _
        & "write_timestamp3='" & Now & "'" _
        & "WHERE status='" & 1 & "'"
        
        strsql = "UPDATE 64_agihan_barang," & G_PULANGAN_TEMP & " SET 64_agihan_barang.status='" & 3 & "'," _
        & "64_agihan_barang.tarikh_jual='" & Frm108.DTPicker4 & "'," _
        & "64_agihan_barang.write_timestamp3='" & LM_NOW & "'" _
        & "WHERE 64_agihan_barang.no_siri_produk = " & G_PULANGAN_TEMP & ".no_siri_produk AND " & G_PULANGAN_TEMP & ".status='" & 2 & "' AND 64_agihan_barang.no_rujukan = " & G_PULANGAN_TEMP & ".no_rujukan_agihan"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update status dan tarikh pulangan dalam table #64_agihan_barang ### - End (Barang yang dijual)

'### Update status barang dalam table #data_database ### - Start (Barang yang dipulangkan)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE Data_Database," & G_PULANGAN_TEMP & " SET Data_Database.StatusItem='" & 10 & "'," _
        & "data_database.no_rujukan_pulang='" & Frm108_LM_No_RUJ & "'" _
        & "WHERE Data_Database.no_siri_produk = " & G_PULANGAN_TEMP & ".no_siri_produk AND " & G_PULANGAN_TEMP & ".status = 1"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update status barang dalam table #data_database ### - End (Barang yang dipulangkan)

'### Update status barang dalam table #data_database ### - Start (Barang yang dijual)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE Data_Database," & G_PULANGAN_TEMP & " SET Data_Database.StatusItem='" & 26 & "'," _
        & "data_database.no_rujukan_pulang='" & Frm108_LM_No_RUJ & "'" _
        & "WHERE Data_Database.no_siri_produk = " & G_PULANGAN_TEMP & ".no_siri_produk AND " & G_PULANGAN_TEMP & ".status = 2"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update status barang dalam table #data_database ### - End (Barang yang dijual)

'### Update status dalam #69_pulangan_barang ### - Start (Data tidak aktif)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE 69_pulangan_barang set status_caption='" & Null & "'" _
        & "WHERE status='" & 0 & "' AND no_rujukan='" & Frm108_LM_No_RUJ & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update status dalam #69_pulangan_barang ### - End (Data tidak aktif)

'### Update status dalam #69_pulangan_barang ### - Start (Data yang dipulangkan)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE 69_pulangan_barang set status_caption='" & "Pulang" & "'" _
        & "WHERE status='" & 1 & "' AND no_rujukan='" & Frm108_LM_No_RUJ & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update status dalam #69_pulangan_barang ### - End (Data yang dipulangkan)

'### Update status dalam #69_pulangan_barang ### - Start (Data yang dijual)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE 69_pulangan_barang set status_caption='" & "Jual" & "'" _
        & "WHERE status='" & 2 & "' AND no_rujukan='" & Frm108_LM_No_RUJ & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update status dalam #69_pulangan_barang ### - End (Data yang dijual)

'#### Update Log Aktiviti Sistem #### - Start
        user = MDI_frm1.L3_Text
        
        LogAct_Memory = "[" & user & "] Pulangan barang dari cawangan. No. Rujukan [" & G_PENYATA_PULANGAN & "]."
        LogDate_Memory = LM_NOW
        Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End

'### update no rujukan sistem ### - Start
        'Set rs = New ADODB.Recordset
        'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        'rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
        
        'If Not rs.EOF Then
        '    If rs!Default1 = "Default" Then
        '        rs!no_rujukan_pulangan = Frm108_LM_No_RUJ + 1 'No. rujukan sistem
        '        rs.Update
        '    End If
        'End If
        
        'rs.Close
        'Set rs = Nothing
'### update no rujukan sistem ### - End

        Call Frm108_hantaran_initial_setting
        Call Frm108_hantaran_initial_setting2

        GM_NEXT_PREV = 0
        
        Frm108.L46_Text = -1 'Titik Pencarian Data
        Frm108.L47_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
        Frm108.L44_Text = 0 'Paparan Page ke-xxx
        
        Call Frm108_senarai_pulangan_header
        Call Frm108_senarai_pulangan
        
        Note = "Data telah berjaya disimpan." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Adakah anda ingin cetak penyata pulangan barang ini?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
            If G_PENYATA_PULANGAN <> vbNullString Then
                Call Frm108_cetak_penyata_pulangan
            End If
        End If
        
        Frm108.TB8.SetFocus
    End If
End If
End Sub
Private Sub CMD22_Click()
'on error resume next
Dim Err(6)
Dim Frm108_LM_No_RUJ As Integer

G_PENYATA_PULANGAN = vbNullString
            
Frm108_LM_CAW_ID = vbNullString
Frm108_LM_No_RUJ = 1

If Frm108.TB8 = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan [Nama]."
End If
If Frm108.TB9 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [No. Kad Pengenalan]."
End If
If Frm108.CBB4 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Cawangan / Kedai]."
End If
If Frm108.CBB5 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Nama Pekerja]."
End If

If Frm108.L42_Text <> vbNullString Then
    If Not IsNumeric(Frm108.L42_Text) Then
        x = x + 1
        Err(x) = "Tiada maklumat barang yang akan dipulangkan." & vbCrLf & _
                "Sila masukkan data barangan yang hendak dipulangkan atau keluar dari menu ini dan cuba sekali lagi."
    Else
        If Frm108.L42_Text = 0 Then
            x = x + 1
            Err(x) = "Tiada maklumat barang yang akan dipulangkan." & vbCrLf & _
                    "Sila masukkan data barangan yang hendak dipulangkan atau keluar dari menu ini dan cuba sekali lagi."
        End If
    End If
End If

If Frm108.L40_Text <> vbNullString Then
    If Not IsNumeric(Frm108.L40_Text) Then
        x = x + 1
        Err(x) = "Technical Error." & vbCrLf & _
                "Sila keluar dari menu ini dan cuba sekali lagi."
    End If
Else

    x = x + 1
    Err(x) = "Technical Error." & vbCrLf & _
            "Sila keluar dari menu ini dan cuba sekali lagi."

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
    
    If Answer = vbYes Then
    
        Frm108_LM_No_RUJ = Frm108.L40_Text 'No. Rujukan sistem

'### Carian No. ID cawangan ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 62_senarai_cawangan where cawangan='" & Frm108.CBB4 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!ID) Then Frm108_LM_CAW_ID = rs!ID
        End If
        
        rs.Close
        Set rs = Nothing
'### Carian No. ID cawangan ### - End

'### No Rujukan pekerja ### - Start
        If Frm108.CBB5 <> vbNullString Then
            Frm108_LM_EMP_NO = Split(Frm108.CBB5, "  |  ")(1)
        End If
'### No Rujukan pekerja ### - End

'### Masukkan data asas pulangan ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 68_pulangan where no_rujukan='" & Frm108.L40_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!no_statement) Then G_PENYATA_PULANGAN = rs!no_statement
            rs!tarikh = Frm108.DTPicker4 'Tarikh barang dipulangkan
            If Frm108.TB8 <> vbNullString Then 'Nama PIC
                rs!Nama = UCase(Frm108.TB8)
            Else
                rs!Nama = Null
            End If
            If Frm108.TB9 <> vbNullString Then 'No. IC
                rs!no_ic = UCase(Frm108.TB9)
            Else
                rs!no_ic = Null
            End If
            If Frm108.TB10 <> vbNullString Then 'No. telefon
                rs!no_tel = UCase(Frm108.TB10)
            Else
                rs!no_tel = Null
            End If
            If Frm108.CBB4 <> vbNullString Then 'Cawangan
                rs!cawangan = Frm108.CBB4
            Else
                rs!cawangan = Null
            End If
            If Frm108_LM_CAW_ID <> vbNullString Then 'No ID cawangan (dari table #62_senarai_cawangan)
                rs!cawangan_id = Frm108_LM_CAW_ID
            Else
                rs!cawangan_id = Null
            End If
            If Frm108.CBB5 <> vbNullString Then 'Nama pekerja yang daftarkan pulangan barang
                rs!nama_pekerja = Frm108_LM_EMP_NO
            Else
                rs!nama_pekerja = Null
            End If
            rs!Status = 1
            rs!write_timestamp2 = Now
            rs.Update
        
        End If
        
        rs.Close
        Set rs = Nothing
'### Masukkan data asas pulangan ### - End

'### Masukkan data di bawah ke dalam #67_pulangan_barang_temp ### - Start
'No rujukan
'Tarikh
'ID cawangan
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE " & G_PULANGAN_TEMP & " set no_rujukan='" & Frm108_LM_No_RUJ & "'," _
        & "tarikh='" & Frm108.DTPicker4 & "'," _
        & "cawangan_id='" & Frm108_LM_CAW_ID & "'" _
        & "WHERE status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "' OR status='" & 4 & "' OR status='" & 5 & "' OR status='" & 6 & "' OR status='" & 7 & "' OR status='" & 8 & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Masukkan data di bawah ke dalam #67_pulangan_barang_temp ### - End

'### #67_pulangan_barang_temp -> #69_pulangan_barang ### - Start (data baru yang dipulangkan)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "insert into 69_pulangan_barang(no_rujukan,no_rujukan_agihan,tarikh,cawangan_id,no_siri_produk,kategori_produk,purity,berat,no_perjanjian_a,no_perjanjian_b,harga_jualan,status,write_timestamp)" & _
                    "select no_rujukan,no_rujukan_agihan,tarikh,cawangan_id,no_siri_produk,kategori_produk,purity,berat,no_perjanjian_a,no_perjanjian_b,harga_jualan,1,Now() from " & G_PULANGAN_TEMP & " WHERE status='" & 5 & "' order by no_siri_Produk ASC"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### #67_pulangan_barang_temp -> #69_pulangan_barang ### - End (data baru yang dipulangkan)

'### #67_pulangan_barang_temp -> #69_pulangan_barang ### - Start (data baru yang dijual)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "insert into 69_pulangan_barang(no_rujukan,no_rujukan_agihan,tarikh,cawangan_id,no_siri_produk,kategori_produk,purity,berat,no_perjanjian_a,no_perjanjian_b,harga_jualan,status,write_timestamp)" & _
                    "select no_rujukan,no_rujukan_agihan,tarikh,cawangan_id,no_siri_produk,kategori_produk,purity,berat,no_perjanjian_a,no_perjanjian_b,harga_jualan,2,Now() from " & G_PULANGAN_TEMP & " WHERE status='" & 6 & "' order by no_siri_Produk ASC"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### #67_pulangan_barang_temp -> #69_pulangan_barang ### - End (data baru yang dijual)

'### #67_pulangan_barang_temp -> #69_pulangan_barang ### - Start (Data yang diedit - barang yang dipulangkan)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

        strsql = "UPDATE 69_pulangan_barang," & G_PULANGAN_TEMP & " SET 69_pulangan_barang.no_rujukan = " & G_PULANGAN_TEMP & ".no_rujukan ," _
        & "69_pulangan_barang.no_rujukan_agihan = " & G_PULANGAN_TEMP & ".no_rujukan_agihan ," _
        & "69_pulangan_barang.tarikh = " & G_PULANGAN_TEMP & ".tarikh ," _
        & "69_pulangan_barang.cawangan_id = " & G_PULANGAN_TEMP & ".cawangan_id ," _
        & "69_pulangan_barang.no_siri_produk = " & G_PULANGAN_TEMP & ".no_siri_produk ," _
        & "69_pulangan_barang.kategori_produk = " & G_PULANGAN_TEMP & ".kategori_produk ," _
        & "69_pulangan_barang.purity = " & G_PULANGAN_TEMP & ".purity ," _
        & "69_pulangan_barang.berat = " & G_PULANGAN_TEMP & ".berat ," _
        & "69_pulangan_barang.no_perjanjian_a = " & G_PULANGAN_TEMP & ".no_perjanjian_a ," _
        & "69_pulangan_barang.no_perjanjian_b = " & G_PULANGAN_TEMP & ".no_perjanjian_b ," _
        & "69_pulangan_barang.harga_jualan = " & G_PULANGAN_TEMP & ".harga_jualan ," _
        & "69_pulangan_barang.status = 1 ," _
        & "69_pulangan_barang.write_timestamp2 = NOW() WHERE 69_pulangan_barang.no_siri_produk = " & G_PULANGAN_TEMP & ".no_siri_produk AND 69_pulangan_barang.no_rujukan = " & G_PULANGAN_TEMP & ".no_rujukan AND " & G_PULANGAN_TEMP & ".status = 3"

        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### #67_pulangan_barang_temp -> #69_pulangan_barang ### - End (Data yang diedit - barang yang dipulangkan)

'### #67_pulangan_barang_temp -> #69_pulangan_barang ### - Start (Data yang diedit - barang yang dijual)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

        strsql = "UPDATE 69_pulangan_barang," & G_PULANGAN_TEMP & " SET 69_pulangan_barang.no_rujukan = " & G_PULANGAN_TEMP & ".no_rujukan ," _
        & "69_pulangan_barang.no_rujukan_agihan = " & G_PULANGAN_TEMP & ".no_rujukan_agihan ," _
        & "69_pulangan_barang.tarikh = " & G_PULANGAN_TEMP & ".tarikh ," _
        & "69_pulangan_barang.cawangan_id = " & G_PULANGAN_TEMP & ".cawangan_id ," _
        & "69_pulangan_barang.no_siri_produk = " & G_PULANGAN_TEMP & ".no_siri_produk ," _
        & "69_pulangan_barang.kategori_produk = " & G_PULANGAN_TEMP & ".kategori_produk ," _
        & "69_pulangan_barang.purity = " & G_PULANGAN_TEMP & ".purity ," _
        & "69_pulangan_barang.berat = " & G_PULANGAN_TEMP & ".berat ," _
        & "69_pulangan_barang.no_perjanjian_a = " & G_PULANGAN_TEMP & ".no_perjanjian_a ," _
        & "69_pulangan_barang.no_perjanjian_b = " & G_PULANGAN_TEMP & ".no_perjanjian_b ," _
        & "69_pulangan_barang.harga_jualan = " & G_PULANGAN_TEMP & ".harga_jualan ," _
        & "69_pulangan_barang.status = 2 ," _
        & "69_pulangan_barang.write_timestamp2 = NOW() WHERE 69_pulangan_barang.no_siri_produk = " & G_PULANGAN_TEMP & ".no_siri_produk AND 69_pulangan_barang.no_rujukan = " & G_PULANGAN_TEMP & ".no_rujukan AND " & G_PULANGAN_TEMP & ".status = 4"

        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### #67_pulangan_barang_temp -> #69_pulangan_barang ### - End (Data yang diedit - barang yang dijual)

'### #67_pulangan_barang_temp -> #69_pulangan_barang ### - Start (Data yang dipadamkan - dari barang yang dipulangkan)
'Ubah status barang yang dipadamkan di dalam table #69_pulangan_barang

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

        strsql = "UPDATE 69_pulangan_barang," & G_PULANGAN_TEMP & " SET 69_pulangan_barang.status = 0 ," _
        & "69_pulangan_barang.write_timestamp3 = Now()" _
        & "WHERE 69_pulangan_barang.no_siri_produk = " & G_PULANGAN_TEMP & ".no_siri_produk AND 69_pulangan_barang.no_rujukan = " & G_PULANGAN_TEMP & ".no_rujukan AND (" & G_PULANGAN_TEMP & ".status = 9 OR " & G_PULANGAN_TEMP & ".status = 10)"
      
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### #67_pulangan_barang_temp -> #69_pulangan_barang ### - End (Data yang dipadamkan)

'### Update status barang dalam table #data_database ### - Start (barang yang dipulangkan)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE Data_Database," & G_PULANGAN_TEMP & " SET Data_Database.StatusItem='" & 10 & "'" _
        & "WHERE Data_Database.no_siri_produk = " & G_PULANGAN_TEMP & ".no_siri_produk AND (" & G_PULANGAN_TEMP & ".status = 3 OR " & G_PULANGAN_TEMP & ".status = 5) AND data_database.no_rujukan_pulang='" & Frm108_LM_No_RUJ & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update status barang dalam table #data_database ### - End (barang yang dipulangkan)

'### Update status barang dalam table #data_database ### - Start (barang yang dijual)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE Data_Database," & G_PULANGAN_TEMP & " SET Data_Database.StatusItem='" & 26 & "'" _
        & "WHERE Data_Database.no_siri_produk = " & G_PULANGAN_TEMP & ".no_siri_produk AND (" & G_PULANGAN_TEMP & ".status = 4 OR " & G_PULANGAN_TEMP & ".status = 6) AND data_database.no_rujukan_pulang='" & Frm108_LM_No_RUJ & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update status barang dalam table #data_database ### - End (barang yang dijual)

'### Update status barang dalam table #data_database ### - Start (barang yang dijual)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE Data_Database," & G_PULANGAN_TEMP & " SET Data_Database.StatusItem='" & 25 & "'" _
        & "WHERE Data_Database.no_siri_produk = " & G_PULANGAN_TEMP & ".no_siri_produk AND (" & G_PULANGAN_TEMP & ".status = 9 OR " & G_PULANGAN_TEMP & ".status = 10) AND data_database.no_rujukan_pulang='" & Frm108_LM_No_RUJ & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update status barang dalam table #data_database ### - End (barang yang dijual)

'### Update status dan tarikh pulangan dalam table #64_agihan_barang ### - Start (Data bagi barang yang dipulangkan)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE 64_agihan_barang," & G_PULANGAN_TEMP & " SET 64_agihan_barang.status='" & 2 & "'," _
        & "64_agihan_barang.tarikh_jual='" & Frm108.DTPicker4 & "'," _
        & "64_agihan_barang.write_timestamp3='" & Now & "'" _
        & "WHERE 64_agihan_barang.no_siri_produk = " & G_PULANGAN_TEMP & ".no_siri_produk AND (" & G_PULANGAN_TEMP & ".status = 3 OR " & G_PULANGAN_TEMP & ".status = 5) AND 64_agihan_barang.no_rujukan = " & G_PULANGAN_TEMP & ".no_rujukan_agihan"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update status dan tarikh pulangan dalam table #64_agihan_barang ### - End (Data bagi barang yang dipulangkan)

'### Update status dan tarikh pulangan dalam table #64_agihan_barang ### - Start (Data bagi barang yang dijual)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE 64_agihan_barang," & G_PULANGAN_TEMP & " SET 64_agihan_barang.status='" & 3 & "'," _
        & "64_agihan_barang.tarikh_jual='" & Frm108.DTPicker4 & "'," _
        & "64_agihan_barang.write_timestamp3='" & Now & "'" _
        & "WHERE 64_agihan_barang.no_siri_produk = " & G_PULANGAN_TEMP & ".no_siri_produk AND (" & G_PULANGAN_TEMP & ".status = 4 OR " & G_PULANGAN_TEMP & ".status = 6) AND 64_agihan_barang.no_rujukan = " & G_PULANGAN_TEMP & ".no_rujukan_agihan"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update status dan tarikh pulangan dalam table #64_agihan_barang ### - End (Data bagi barang yang dijual)

'### Update tarikh pulangan dalam table #64_agihan_barang ### - Start (semua data)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE 64_agihan_barang," & G_PULANGAN_TEMP & " SET 64_agihan_barang.tarikh_jual='" & Frm108.DTPicker4 & "'" _
        & "WHERE 64_agihan_barang.no_siri_produk = " & G_PULANGAN_TEMP & ".no_siri_produk AND 64_agihan_barang.no_rujukan = " & G_PULANGAN_TEMP & ".no_rujukan_agihan AND (" & G_PULANGAN_TEMP & ".status = 3 OR " & G_PULANGAN_TEMP & ".status = 4 OR " & G_PULANGAN_TEMP & ".status = 5 OR " & G_PULANGAN_TEMP & ".status = 6 OR " & G_PULANGAN_TEMP & ".status = 9)"
        
        'Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update tarikh pulangan dalam table #64_agihan_barang ### - End (semua data)

'### Update status dan tarikh pulangan dalam table #64_agihan_barang ### - Start (Padam data - barang yang dipulangkan)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE 64_agihan_barang," & G_PULANGAN_TEMP & " SET 64_agihan_barang.status='" & 1 & "'," _
        & "64_agihan_barang.tarikh_jual = NULL ," _
        & "64_agihan_barang.write_timestamp3='" & Now & "'" _
        & "WHERE 64_agihan_barang.no_siri_produk = " & G_PULANGAN_TEMP & ".no_siri_produk AND 64_agihan_barang.no_rujukan = " & G_PULANGAN_TEMP & ".no_rujukan_agihan AND " & G_PULANGAN_TEMP & ".status = 9"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update status dan tarikh pulangan dalam table #64_agihan_barang ### - End (Padam data - barang yang dipulangkan)

'### Update status dan tarikh pulangan dalam table #64_agihan_barang ### - Start (Padam data - barang yang dijual)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE 64_agihan_barang," & G_PULANGAN_TEMP & " SET 64_agihan_barang.status='" & 1 & "'," _
        & "64_agihan_barang.tarikh_jual = NULL ," _
        & "64_agihan_barang.write_timestamp3='" & Now & "'" _
        & "WHERE 64_agihan_barang.no_siri_produk = " & G_PULANGAN_TEMP & ".no_siri_produk AND 64_agihan_barang.no_rujukan = " & G_PULANGAN_TEMP & ".no_rujukan_agihan AND " & G_PULANGAN_TEMP & ".status = 10"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update status dan tarikh pulangan dalam table #64_agihan_barang ### - End (Padam data - barang yang dijual)

'### Update status dalam #69_pulangan_barang ### - Start (Data tidak aktif)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE 69_pulangan_barang set status_caption='" & Null & "'" _
        & "WHERE status='" & 0 & "' AND no_rujukan='" & Frm108_LM_No_RUJ & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update status dalam #69_pulangan_barang ### - End (Data tidak aktif)

'### Update status dalam #69_pulangan_barang ### - Start (Data yang dipulangkan)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE 69_pulangan_barang set status_caption='" & "Pulang" & "'" _
        & "WHERE status='" & 1 & "' AND no_rujukan='" & Frm108_LM_No_RUJ & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update status dalam #69_pulangan_barang ### - End (Data yang dipulangkan)

'### Update status dalam #69_pulangan_barang ### - Start (Data yang dijual)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE 69_pulangan_barang set status_caption='" & "Jual" & "'" _
        & "WHERE status='" & 2 & "' AND no_rujukan='" & Frm108_LM_No_RUJ & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update status dalam #69_pulangan_barang ### - End (Data yang dijual)

'### Update data dalam table #64_agihan_barang ### - Start
'Tarikh
'ID cawangan

'        Set rs = New ADODB.Recordset
'        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
'        strsql = "UPDATE 69_pulangan_barang set tarikh='" & Frm108.DTPicker4 & "'," _
'        & "cawangan_id='" & Frm108_LM_CAW_ID & "'" _
'        & "WHERE (status='" & 0 & "' OR status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "') AND no_rujukan='" & Frm108.L40_Text & "'"
        
'        Set rs = cn.Execute(strsql)
'        Set rs = Nothing
'### Update data dalam table #64_agihan_barang ### - End

'#### Update Log Aktiviti Sistem #### - Start
        user = MDI_frm1.L3_Text
        
        LogAct_Memory = "[" & user & "] Edit data pulangan barang oleh cawangan. No. Rujukan [" & G_PENYATA_PULANGAN & "]."
        LogDate_Memory = DateTime.Date & " " & DateTime.Time$
        Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End

        GM_NEXT_PREV = 2
    
        Call Frm108_senarai_pulangan_header
        Call Frm108_senarai_pulangan
        
        Frm108.Pic3.Visible = True
        Frm108.Pic6.Visible = False
        
        Note = "Data telah berjaya disimpan." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Adakah anda ingin cetak penyata pulangan barang ini?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
            If G_PENYATA_PULANGAN <> vbNullString Then
                Call Frm108_cetak_penyata_pulangan
            End If
        End If

    End If
End If
End Sub
Private Sub CMD23_Click()
'on error resume next
Note = "Adakah anda ingin batalkan edit data ini?" & vbCrLf & _
        "Sistem tidak akan menyimpan data jika terdapat data yang diubah." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then
    
    Frm108.Pic3.Visible = True
    Frm108.Pic6.Visible = False

End If
End Sub
Private Sub CMD24_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm108_senarai_pulangan_header
Call Frm108_senarai_pulangan
End Sub
Private Sub CMD25_Click()
'on error resume next
Dim Frm108_LM_CURR_PAGE As Double
Dim Frm108_LM_TOTAL_PAGE As Double

Frm108_LM_CURR_PAGE = 0
Frm108_LM_TOTAL_PAGE = 0

If Frm108.L44_Text <> vbNullString And IsNumeric(Frm108.L44_Text) Then
    If Frm108.L45_Text <> vbNullString And IsNumeric(Frm108.L45_Text) Then
        Frm108_LM_CURR_PAGE = Frm108.L44_Text
        Frm108_LM_TOTAL_PAGE = Frm108.L45_Text
        
        If Frm108_LM_CURR_PAGE < Frm108_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm108_senarai_pulangan_header
            Call Frm108_senarai_pulangan
            
        End If
    End If
End If
End Sub
Private Sub CMD26_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm108_report_inventory_header
Call Frm108_report_inventory
End Sub
Private Sub CMD27_Click()
'on error resume next
Dim Frm108_LM_CURR_PAGE As Double
Dim Frm108_LM_TOTAL_PAGE As Double

Frm108_LM_CURR_PAGE = 0
Frm108_LM_TOTAL_PAGE = 0

If Frm108.L55_Text <> vbNullString And IsNumeric(Frm108.L55_Text) Then
    If Frm108.L56_Text <> vbNullString And IsNumeric(Frm108.L56_Text) Then
        Frm108_LM_CURR_PAGE = Frm108.L55_Text
        Frm108_LM_TOTAL_PAGE = Frm108.L56_Text
        
        If Frm108_LM_CURR_PAGE < Frm108_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm108_report_inventory_header
            Call Frm108_report_inventory
            
        End If
    End If
End If
End Sub
Private Sub CMD28_Click()
'on error resume next
If Frm108.CBB6 = vbNullString Then
    MsgBox "Sila buat pilih cawangan", vbInformation, "Info"
    
    Exit Sub
End If
If Frm108.CBB7 = vbNullString Then
    MsgBox "Sila buat pilih jenis report", vbInformation, "Info"
    
    Exit Sub
End If

Note = "Sistem mungkin akan mengambil masa untuk mengeluarkan report." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan ?"


Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then
    
    If Frm108.CB6 = 0 Then
        Frm108.L49_Text = 0 'Memory : Jenis report , 0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh
    Else
        Frm108.L49_Text = 1 'Memory : Jenis report , 0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh
    End If
    
    Frm108.L60_Text = Frm108.DTPicker5 'Memori : Tarikh mula
    Frm108.L61_Text = Frm108.DTPicker6 'Memori : Tarikh akhir
    Frm108.L50_Text = Frm108.CBB6 'Memory : Cawangan
    Frm108.L51_Text = Frm108.CBB7 'Memory : Jenis Report
    
    If Frm108.L50_Text = vbNullString And Frm108.L51_Text = vbNullString And Frm108.L60_Text = vbNullString And Frm108.L61_Text = vbNullString Then
        
        MsgBox "Technical Error." & vbCrLf & _
                "Sila keluar dari menu ini dan cuba sekali lagi.", vbExclamation, "Error"
                
        Exit Sub
        
    Else
    
        GM_NEXT_PREV = 0
        
        Frm108.L57_Text = -1 'Titik Pencarian Data
        Frm108.L58_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
        Frm108.L55_Text = 0 'Paparan Page ke-xxx
        
        Call Frm108_report_inventory_header
        Call Frm108_report_inventory
        
        If Frm108.L53_Text <> vbNullString Then
            If Frm108.L53_Text = 0 Then MsgBox "Tiada rekod dijumpai.", vbInformation, "Info"
        End If
    
    
    End If
End If
End Sub
Private Sub CMD3_Click()
'on error resume next
Note = "Adakah anda ingin batalkan edit data ini?" & vbCrLf & _
        "Sistem tidak akan menyimpan data jika terdapat data yang diubah." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then

    Call Frm108_cmd_visible_1
    Frm108.L11_Text = 0 'Memory : No. ID cawangan
    Frm108.TB4 = vbNullString
    Frm108.TB4.SetFocus
    
End If
End Sub
Private Sub CMD4_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm108_senarai_cawangan_header
Call Frm108_senarai_cawangan
End Sub
Private Sub CMD5_Click()
'on error resume next
Dim Frm108_LM_CURR_PAGE As Double
Dim Frm108_LM_TOTAL_PAGE As Double

Frm108_LM_CURR_PAGE = 0
Frm108_LM_TOTAL_PAGE = 0

If Frm108.L6_Text <> vbNullString And IsNumeric(Frm108.L6_Text) Then
    If Frm108.L7_Text <> vbNullString And IsNumeric(Frm108.L7_Text) Then
        Frm108_LM_CURR_PAGE = Frm108.L6_Text
        Frm108_LM_TOTAL_PAGE = Frm108.L7_Text
        
        If Frm108_LM_CURR_PAGE < Frm108_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm108_senarai_cawangan_header
            Call Frm108_senarai_cawangan
            
        End If
    End If
End If
End Sub
Private Sub CMD6_Click()
'on error resume next
Dim Err(6)
Dim Frm108_LM_No_RUJ As Integer

G_PENYATA_AMBILAN = vbNullString
            
Frm108_LM_CAW_ID = vbNullString
Frm108_LM_No_RUJ = 1

If Frm108.TB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan [Nama]."
End If
If Frm108.TB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [No. Kad Pengenalan]."
End If
If Frm108.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Cawangan / Kedai]."
End If
If Frm108.CBB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Nama Pekerja]."
End If

If Frm108.L17_Text <> vbNullString Then
    If Not IsNumeric(Frm108.L17_Text) Then
        x = x + 1
        Err(x) = "Tiada maklumat barang yang akan diagihkan." & vbCrLf & _
                "Sila masukkan data barangan yang hendak diagihkan atau keluar dari menu ini dan cuba sekali lagi."
    Else
        If Frm108.L17_Text = 0 Then
            x = x + 1
            Err(x) = "Tiada maklumat barang yang akan diagihkan." & vbCrLf & _
                    "Sila masukkan data barangan yang hendak diagihkan atau keluar dari menu ini dan cuba sekali lagi."
        End If
    End If
End If

If Frm108.L12_Text <> vbNullString Then
    If Not IsNumeric(Frm108.L12_Text) Then
        x = x + 1
        Err(x) = "Technical Error." & vbCrLf & _
                "Sila keluar dari menu ini dan cuba sekali lagi."
    End If
Else

    x = x + 1
    Err(x) = "Technical Error." & vbCrLf & _
            "Sila keluar dari menu ini dan cuba sekali lagi."

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
    
    If Answer = vbYes Then
    
'### Periksa nombor rujukan ### - Start
        Frm108_LM_No_RUJ = Frm108.L12_Text

'---------------------------------------No. Invoice
        LM_NOW = Now
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        rs.Open "select * from 9_rujukan_agihan", cn2, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        rs!tarikh = Frm108.DTPicker1
        rs!terminal = G_TERMINAL
        rs!write_timestamp = LM_NOW
        rs!Status = 1
        rs!nama_staff = MDI_frm1.L3_Text
        rs.Update
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        rs.Open "select * from 9_rujukan_agihan where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & Frm108.DTPicker1 & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!ID) Then

                Frm108_LM_No_RUJ = rs!ID 'No. Rujukan Belian
                rs!no_rujukan = "BKS" & Format(Frm108_LM_No_RUJ, "000000")
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
        
        GoTo a:
Re_Gen_No:
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 63_agihan where no_rujukan='" & Frm108_LM_No_RUJ & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            Frm108_LM_No_RUJ = Frm108_LM_No_RUJ + 1
            Frm108.L12_Text = Frm108_LM_No_RUJ
            
            rs.Close
            Set rs = Nothing
            
            GoTo Re_Gen_No:
        End If
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 64_agihan_barang where no_rujukan='" & Frm108_LM_No_RUJ & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            Frm108_LM_No_RUJ = Frm108_LM_No_RUJ + 1
            Frm108.L12_Text = Frm108_LM_No_RUJ
            
            rs.Close
            Set rs = Nothing
            
            GoTo Re_Gen_No:
        End If
        
        rs.Close
        Set rs = Nothing
'### Periksa nombor rujukan ### - End

a:

'### Carian No. ID cawangan ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 62_senarai_cawangan where cawangan='" & Frm108.CBB1 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!ID) Then Frm108_LM_CAW_ID = rs!ID
        End If
        
        rs.Close
        Set rs = Nothing
'### Carian No. ID cawangan ### - End

'### No Rujukan pekerja ### - Start
        If Frm108.CBB2 <> vbNullString Then
            Frm108_LM_EMP_NO = Split(Frm108.CBB2, "  |  ")(1)
        End If
'### No Rujukan pekerja ### - End

'### Masukkan data asas agihan ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 63_agihan", cn, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        If Frm108.L12_Text <> vbNullString Then 'No. rujukan sistem
            rs!no_rujukan = Frm108_LM_No_RUJ
        Else
            rs!no_rujukan = Null
        End If
        If Frm108.L12_Text <> vbNullString Then 'No. rujukan sistem
            rs!no_statement = "BKS" & Format(Frm108_LM_No_RUJ, "000000")
            G_PENYATA_AMBILAN = "BKS" & Format(Frm108_LM_No_RUJ, "000000")
        Else
            rs!no_statement = Null
        End If
        rs!tarikh = Frm108.DTPicker1 'Tarikh barang diambil
        If Frm108.TB1 <> vbNullString Then 'Nama PIC
            rs!Nama = UCase(Frm108.TB1)
        Else
            rs!Nama = Null
        End If
        If Frm108.TB2 <> vbNullString Then 'No. IC
            rs!no_ic = UCase(Frm108.TB2)
        Else
            rs!no_ic = Null
        End If
        If Frm108.TB3 <> vbNullString Then 'No. telefon
            rs!no_tel = UCase(Frm108.TB3)
        Else
            rs!no_tel = Null
        End If
        If Frm108.CBB1 <> vbNullString Then 'Cawangan
            rs!cawangan = Frm108.CBB1
        Else
            rs!cawangan = Null
        End If
        If Frm108_LM_CAW_ID <> vbNullString Then 'No ID cawangan (dari table #62_senarai_cawangan)
            rs!cawangan_id = Frm108_LM_CAW_ID
        Else
            rs!cawangan_id = Null
        End If
        If Frm108.CBB2 <> vbNullString Then 'Nama pekerja yang daftarkan agihan barang
            rs!nama_pekerja = Frm108_LM_EMP_NO
        Else
            rs!nama_pekerja = Null
        End If
        rs!Status = 1
        rs!write_timestamp = LM_NOW
        rs.Update
        
        rs.Close
        Set rs = Nothing
'### Masukkan data asas agihan ### - End

'### Masukkan data di bawah ke dalam #65_agihan_barang_temp ### - Start
'No rujukan
'Tarikh
'ID cawangan

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE " & G_AGIHAN_TEMP & " set no_rujukan='" & Frm108_LM_No_RUJ & "'," _
        & "tarikh='" & Frm108.DTPicker1 & "'," _
        & "cawangan='" & Frm108.CBB1 & "'," _
        & "cawangan_id='" & Frm108_LM_CAW_ID & "'" _
        & "WHERE status='" & 1 & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Masukkan data di bawah ke dalam #65_agihan_barang_temp ### - End

'### #65_agihan_barang_temp -> #64_agihan_barang ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "insert into 64_agihan_barang(no_rujukan,tarikh,cawangan,cawangan_id,no_siri_produk,kategori_produk,purity,berat,status,write_timestamp)" & _
                    "select no_rujukan,tarikh,cawangan,cawangan_id,no_siri_produk,kategori_produk,purity,berat,1,'" & LM_NOW & "' from " & G_AGIHAN_TEMP & " WHERE status='" & 1 & "' order by no_siri_Produk ASC"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### #65_agihan_barang_temp -> #64_agihan_barang ### - End

'### Update status barang dalam table #data_database ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE Data_Database," & G_AGIHAN_TEMP & " SET Data_Database.StatusItem='" & 25 & "'," _
        & "Data_Database.cawangan_id='" & Frm108_LM_CAW_ID & "'" _
        & "WHERE Data_Database.no_siri_produk = " & G_AGIHAN_TEMP & ".no_siri_produk AND " & G_AGIHAN_TEMP & ".status = 1"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update status barang dalam table #data_database ### - End

'#### Update Log Aktiviti Sistem #### - Start
        user = MDI_frm1.L3_Text
        
        LogAct_Memory = "[" & user & "] Agihan barang kepada cawangan. No. Rujukan [" & G_PENYATA_AMBILAN & "]."
        LogDate_Memory = LM_NOW
        Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End

'### update no rujukan sistem ### - Start
        'Set rs = New ADODB.Recordset
        'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        'rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
        
        'If Not rs.EOF Then
        '    If rs!Default1 = "Default" Then
        '        rs!no_rujukan_agihan = Frm108_LM_No_RUJ + 1 'No. rujukan sistem
        '        rs.Update
        '    End If
        'End If
        
        'rs.Close
        'Set rs = Nothing
'### update no rujukan sistem ### - End

        Call Frm108_hantaran_initial_setting
        Call Frm108_hantaran_initial_setting2
        
        GM_NEXT_PREV = 0
        
        Frm108.L15_Text = -1 'Titik Pencarian Data
        Frm108.L16_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
        Frm108.L13_Text = 0 'Paparan Page ke-xxx
        
        Call Frm108_senarai_agihan_header
        Call Frm108_senarai_agihan
        
        Note = "Data telah berjaya disimpan." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Adakah anda ingin cetak penyata ambilan barang ini?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
            If G_PENYATA_AMBILAN <> vbNullString Then
                Call Frm108_cetak_penyata_ambilan
            End If
        End If
        
        Frm108.TB1.SetFocus
    End If
End If
End Sub
Private Sub CMD7_Click()
'on error resume next
Dim Err(6)
Dim Frm108_LM_No_RUJ As Integer

G_PENYATA_AMBILAN = vbNullString
            
Frm108_LM_CAW_ID = vbNullString
Frm108_LM_No_RUJ = 1

If Frm108.TB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan [Nama]."
End If
If Frm108.TB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [No. Kad Pengenalan]."
End If
If Frm108.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Cawangan / Kedai]."
End If
If Frm108.CBB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Nama Pekerja]."
End If

If Frm108.L17_Text <> vbNullString Then
    If Not IsNumeric(Frm108.L17_Text) Then
        x = x + 1
        Err(x) = "Tiada maklumat barang yang akan diagihkan." & vbCrLf & _
                "Sila masukkan data barangan yang hendak diagihkan atau keluar dari menu ini dan cuba sekali lagi."
    Else
        If Frm108.L17_Text = 0 Then
            x = x + 1
            Err(x) = "Tiada maklumat barang yang akan diagihkan." & vbCrLf & _
                    "Sila masukkan data barangan yang hendak diagihkan atau keluar dari menu ini dan cuba sekali lagi."
        End If
    End If
End If

If Frm108.L12_Text <> vbNullString Then
    
    If Frm108.L12_Text = 0 Then
        x = x + 1
        Err(x) = "Technical Error." & vbCrLf & _
                "Sila keluar dari menu ini dan cuba sekali lagi."
    End If
    
Else

    x = x + 1
    Err(x) = "Technical Error." & vbCrLf & _
            "Sila keluar dari menu ini dan cuba sekali lagi."

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
    
    If Answer = vbYes Then

'### Carian No. ID cawangan ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 62_senarai_cawangan where cawangan='" & Frm108.CBB1 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!ID) Then Frm108_LM_CAW_ID = rs!ID
        End If
        
        rs.Close
        Set rs = Nothing
'### Carian No. ID cawangan ### - End

'### No Rujukan pekerja ### - Start
        If Frm108.CBB2 <> vbNullString Then
            Frm108_LM_EMP_NO = Split(Frm108.CBB2, "  |  ")(1)
        End If
'### No Rujukan pekerja ### - End

'### Masukkan data asas agihan ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 63_agihan where no_rujukan='" & Frm108.L12_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!no_statement) Then G_PENYATA_AMBILAN = rs!no_statement
            rs!tarikh = Frm108.DTPicker1 'Tarikh barang diambil
            If Frm108.TB1 <> vbNullString Then 'Nama PIC
                rs!Nama = UCase(Frm108.TB1)
            Else
                rs!Nama = Null
            End If
            If Frm108.TB2 <> vbNullString Then 'No. IC
                rs!no_ic = UCase(Frm108.TB2)
            Else
                rs!no_ic = Null
            End If
            If Frm108.TB3 <> vbNullString Then 'No. telefon
                rs!no_tel = UCase(Frm108.TB3)
            Else
                rs!no_tel = Null
            End If
            If Frm108.CBB1 <> vbNullString Then 'Cawangan
                rs!cawangan = Frm108.CBB1
            Else
                rs!cawangan = Null
            End If
            If Frm108_LM_CAW_ID <> vbNullString Then 'No ID cawangan (dari table #62_senarai_cawangan)
                rs!cawangan_id = Frm108_LM_CAW_ID
            Else
                rs!cawangan_id = Null
            End If
            If Frm108.CBB2 <> vbNullString Then 'Nama pekerja yang daftarkan agihan barang
                rs!nama_pekerja = Frm108_LM_EMP_NO
            Else
                rs!nama_pekerja = Null
            End If
            rs!Status = 1
            rs!write_timestamp2 = Now
            rs.Update
        
        End If
        
        rs.Close
        Set rs = Nothing
'### Masukkan data asas agihan ### - End

'### Masukkan data di bawah ke dalam #65_agihan_barang_temp ### - Start
'No rujukan
'Tarikh
'ID cawangan

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE " & G_AGIHAN_TEMP & " set no_rujukan='" & Frm108.L12_Text & "'," _
        & "tarikh='" & Frm108.DTPicker1 & "'," _
        & "cawangan='" & Frm108.CBB1 & "'," _
        & "cawangan_id='" & Frm108_LM_CAW_ID & "'" _
        & "WHERE status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "' OR status='" & 4 & "' OR status='" & 5 & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Masukkan data di bawah ke dalam #65_agihan_barang_temp ### - End

'### #65_agihan_barang_temp -> #64_agihan_barang ### - Start (Data baru)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "insert into 64_agihan_barang(no_rujukan,tarikh,cawangan_id,no_siri_produk,kategori_produk,purity,berat,status,write_timestamp)" & _
                    "select no_rujukan,tarikh,cawangan_id,no_siri_produk,kategori_produk,purity,berat,1,Now() from " & G_AGIHAN_TEMP & " WHERE status='" & 1 & "' OR status='" & 3 & "' order by no_siri_Produk ASC"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### #65_agihan_barang_temp -> #64_agihan_barang ### - End (Data baru)

'### #65_agihan_barang_temp -> #64_agihan_barang ### - Start (Data yang diedit)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

        strsql = "UPDATE 64_agihan_barang," & G_AGIHAN_TEMP & " SET 64_agihan_barang.no_rujukan = " & G_AGIHAN_TEMP & ".no_rujukan ," _
        & "64_agihan_barang.tarikh = " & G_AGIHAN_TEMP & ".tarikh ," _
        & "64_agihan_barang.cawangan = " & G_AGIHAN_TEMP & ".cawangan ," _
        & "64_agihan_barang.cawangan_id = " & G_AGIHAN_TEMP & ".cawangan_id ," _
        & "64_agihan_barang.no_siri_produk = " & G_AGIHAN_TEMP & ".no_siri_produk ," _
        & "64_agihan_barang.kategori_produk = " & G_AGIHAN_TEMP & ".kategori_produk ," _
        & "64_agihan_barang.purity = " & G_AGIHAN_TEMP & ".purity ," _
        & "64_agihan_barang.berat = " & G_AGIHAN_TEMP & ".berat ," _
        & "64_agihan_barang.status = 1 ," _
        & "64_agihan_barang.write_timestamp2 = NOW() WHERE " & G_AGIHAN_TEMP & ".status = 5 AND 64_agihan_barang.no_siri_produk = " & G_AGIHAN_TEMP & ".no_siri_produk AND 64_agihan_barang.no_rujukan = " & G_AGIHAN_TEMP & ".no_rujukan"

        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### #65_agihan_barang_temp -> #64_agihan_barang ### - End (Data yang diedit)

'### #65_agihan_barang_temp -> #64_agihan_barang ### - Start (Data yang dipadamkan)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

        strsql = "UPDATE 64_agihan_barang," & G_AGIHAN_TEMP & " SET 64_agihan_barang.status = 0 ," _
        & "64_agihan_barang.write_timestamp3 = Now()" _
        & "WHERE " & G_AGIHAN_TEMP & ".status = 4 AND 64_agihan_barang.no_siri_produk = " & G_AGIHAN_TEMP & ".no_siri_produk AND 64_agihan_barang.no_rujukan = " & G_AGIHAN_TEMP & ".no_rujukan"
      
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### #65_agihan_barang_temp -> #64_agihan_barang ### - End (Data yang dipadamkan)

'### Update status barang dalam table #data_database ### - Start (Data baru)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE Data_Database," & G_AGIHAN_TEMP & " SET Data_Database.StatusItem='" & 25 & "'," _
        & "Data_Database.cawangan_id='" & Frm108_LM_CAW_ID & "'" _
        & "WHERE Data_Database.no_siri_produk = " & G_AGIHAN_TEMP & ".no_siri_produk AND " & G_AGIHAN_TEMP & ".status = 3"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update status barang dalam table #data_database ### - End (Data baru)

'### Update status barang dalam table #data_database ### - Start (Data padam)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE Data_Database," & G_AGIHAN_TEMP & " SET Data_Database.StatusItem='" & 10 & "'," _
        & "Data_Database.cawangan_id = NULL " _
        & "WHERE Data_Database.no_siri_produk = " & G_AGIHAN_TEMP & ".no_siri_produk AND " & G_AGIHAN_TEMP & ".status = 4"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update status barang dalam table #data_database ### - End (Data padam)

'### Update ID cawangan dalam table #data_database ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE Data_Database,64_agihan_barang SET Data_Database.cawangan_id='" & Frm108_LM_CAW_ID & "'" _
        & "WHERE Data_Database.no_siri_produk = 64_agihan_barang.no_siri_produk AND 64_agihan_barang.status = 1"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update ID cawangan dalam table #data_database ### - End

'### Update data dalam table #64_agihan_barang ### - Start
'Tarikh
'ID cawangan
'cawangan
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE 64_agihan_barang set tarikh='" & Frm108.DTPicker1 & "'," _
        & "cawangan='" & Frm108.CBB1 & "'," _
        & "cawangan_id='" & Frm108_LM_CAW_ID & "'" _
        & "WHERE (status='" & 0 & "' OR status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "') AND no_rujukan='" & Frm108.L12_Text & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update data dalam table #64_agihan_barang ### - End

'#### Update Log Aktiviti Sistem #### - Start
        user = MDI_frm1.L3_Text
        
        LogAct_Memory = "[" & user & "] Edit data agihan barang kepada cawangan. No. Rujukan [" & G_PENYATA_AMBILAN & "]."
        LogDate_Memory = DateTime.Date & " " & DateTime.Time$
        Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End

        GM_NEXT_PREV = 2
    
        Call Frm108_senarai_agihan_barang_header
        Call Frm108_senarai_agihan_barang
        
        Frm108.Pic3.Visible = True
        Frm108.Pic2.Visible = False
        
        Note = "Data telah berjaya disimpan." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Adakah anda ingin cetak penyata ambilan barang ini?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
            If G_PENYATA_AMBILAN <> vbNullString Then
                Call Frm108_cetak_penyata_ambilan
            End If
        End If

    End If
End If
End Sub
Private Sub CMD8_Click()
'on error resume next
Note = "Adakah anda ingin batalkan edit data ini?" & vbCrLf & _
        "Sistem tidak akan menyimpan data jika terdapat data yang diubah." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then
    
    Frm108.Pic3.Visible = True
    Frm108.Pic2.Visible = False

End If
End Sub
Private Sub CMD9_Click()
'on error resume next
If InStr(1, Frm108.TB5, "*") <> 0 Or InStr(1, Frm108.TB5, "/") <> 0 Or InStr(1, Frm108.TB5, "\") <> 0 Or InStr(1, Frm108.TB5, "'") <> 0 Then

    MsgBox "No. Siri Produk Mengandungi Simbol Yang Tidak Sah , ['].", vbInformation, "Info"
    
    Frm108.TB5 = vbNullString
    Exit Sub
End If

Call Frm108_periksa_data_barang
End Sub

Private Sub Form_Load()
'on error resume next
Call Frm108_one_time_reset
End Sub
Private Sub Frm108_SM_cetak_penyata_Click()
'On Error Resume Next
Frm108_LM_NO_STATEMENT = vbNullString

If Frm108.MSFlexGrid3 <> vbNullString Then
    'If Button = vbRightButton Then
    
        If IsNumeric(Frm108.MSFlexGrid3) Then

            Frm108_LM_NO_STATEMENT = Frm108.MSFlexGrid3.TextMatrix(Frm108.MSFlexGrid3, 4) 'No. Statement
            
            If Frm108_LM_NO_STATEMENT <> vbNullString Then
                
                If Frm108.L48_Text = 0 Then 'Memory : Jenis report , 0 : Agihan , 1 : Pulangan
                    G_PENYATA_AMBILAN = Frm108_LM_NO_STATEMENT
                    Call Frm108_cetak_penyata_ambilan
                ElseIf Frm108.L48_Text = 1 Then 'Memory : Jenis report , 0 : Agihan , 1 : Pulangan
                    G_PENYATA_PULANGAN = Frm108_LM_NO_STATEMENT
                    Call Frm108_cetak_penyata_pulangan
                End If
                

            End If
            
        End If
        
    'End If
End If
End Sub
Private Sub Frm108_SM_edit_data_Click()
'On Error Resume Next
Frm108_LM_ID = vbNullString
DATA_FOUND = 0

If Frm108.MSFlexGrid1 <> vbNullString Then
    'If Button = vbRightButton Then
    
        If IsNumeric(Frm108.MSFlexGrid1) Then
            Frm108_LM_ID = Frm108.MSFlexGrid1.TextMatrix(Frm108.MSFlexGrid1, 2) 'No. ID
            
            If Frm108_LM_ID <> vbNullString Then
                
                Call Frm108_cawangan_initial_setting

                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 62_senarai_cawangan where ID='" & Frm108_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Not IsNull(rs!ID) Then Frm108.L11_Text = rs!ID
                    If Not IsNull(rs!cawangan) Then Frm108.TB4 = rs!cawangan 'Nama Cawangan
                    DATA_FOUND = 1
                End If
                
                rs.Close
                Set rs = Nothing
                
                If DATA_FOUND = 1 Then
                
                    Call Frm108_cmd_invisible_1
                
                End If
                
            End If
            
        End If
        
    'End If
End If
End Sub
Private Sub Frm108_SM_excel_Click()
'on error resume next
'REPORT TRADE IN AGEN - EXCEL
Dim TA As Date
Dim TM As Date

Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem mungkin akan mengambil masa untuk mengeluarkan report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila tunggu sehingga sistem siap keluarkan report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then

    x = 0
    
    If Frm108.L49_Text = 1 Then 'Memory : Jenis report , 0 : Tiada pilihan tarikh , 1 : Ada pilihan
        TM = Frm108.L60_Text 'Tarikh mula
        TA = Frm108.L61_Text 'Tarikh akhir
    End If
    If Frm108.L50_Text = "Semua Cawangan" Then
        Frm108_LM_SEARCH_1 = Null
        Frm108_LM_SEARCH_1_LOGIC = "<>"
    Else
        Frm108_LM_SEARCH_1 = Frm108.L50_Text
        Frm108_LM_SEARCH_1_LOGIC = "="
    End If
    If Frm108.L51_Text = "Semua Jenis Report" Then
    
        Frm108_LM_SEARCH_2 = 1
        Frm108_LM_SEARCH_2_LOGIC = "="
    
        Frm108_LM_SEARCH_3 = 2
        Frm108_LM_SEARCH_3_LOGIC = "="
        
        Frm108_LM_SEARCH_4 = 3
        Frm108_LM_SEARCH_4_LOGIC = "="
        
    ElseIf Frm108.L51_Text = "Agihan" Then
    
        Frm108_LM_SEARCH_2 = 1
        Frm108_LM_SEARCH_2_LOGIC = "="
        
        Frm108_LM_SEARCH_3 = 2
        Frm108_LM_SEARCH_3_LOGIC = "="
        
        Frm108_LM_SEARCH_4 = 3
        Frm108_LM_SEARCH_4_LOGIC = "="
        
    ElseIf Frm108.L51_Text = "Pulangan" Then
    
        Frm108_LM_SEARCH_2 = 2
        Frm108_LM_SEARCH_2_LOGIC = "="
        
        Frm108_LM_SEARCH_3 = 2
        Frm108_LM_SEARCH_3_LOGIC = "="
        
        Frm108_LM_SEARCH_4 = 2
        Frm108_LM_SEARCH_4_LOGIC = "="
        
    ElseIf Frm108.L51_Text = "Dijual" Then
    
        Frm108_LM_SEARCH_2 = 3
        Frm108_LM_SEARCH_2_LOGIC = "="
        
        Frm108_LM_SEARCH_3 = 3
        Frm108_LM_SEARCH_3_LOGIC = "="
        
        Frm108_LM_SEARCH_4 = 3
        Frm108_LM_SEARCH_4_LOGIC = "="
        
    ElseIf Frm108.L51_Text = "Belum Dipulangkan" Then
    
        Frm108_LM_SEARCH_2 = 1
        Frm108_LM_SEARCH_2_LOGIC = "="
        
        Frm108_LM_SEARCH_3 = 1
        Frm108_LM_SEARCH_3_LOGIC = "="
        
        Frm108_LM_SEARCH_4 = 1
        Frm108_LM_SEARCH_4_LOGIC = "="
        
    End If
    
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 15 'Tarikh Agihan
        .Columns("C").ColumnWidth = 40 'Cawangan
        .Columns("D").ColumnWidth = 15 'No. Siri Produk
        .Columns("E").ColumnWidth = 40 'Nama Produk
        .Columns("F").ColumnWidth = 15 'Berat (g)
        .Columns("G").ColumnWidth = 15 'Status
        .Columns("H").ColumnWidth = 15 'Tarikh Pulangan
    
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
                .Cells(1, 4) = rs!nama_kedai
                .Cells(1, 4).Font.Name = "Times New Roman"
            End If
            If Not IsNull(rs!no_pendaftaran) Then .Cells(2, 4) = rs!no_pendaftaran
            If Not IsNull(rs!alamat) Then .Cells(3, 4) = rs!alamat
            If Not IsNull(rs!no_tel) Then .Cells(4, 4) = rs!no_tel
            If Not IsNull(rs!no_id_gst) Then .Cells(5, 4) = rs!no_id_gst
        End If
        
        rs.Close
        Set rs = Nothing
        '### Maklumat kedai ### - End
        
        .Cells(1, 4).Font.Bold = True
        .Cells(1, 4).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 4).HorizontalAlignment = xlCenter
        Next Row
        
        If Frm108.L49_Text = 0 Then .Cells(7, 1) = "Report inventori bagi [" & Frm108.L51_Text & "] kepada/oleh cawangan [" & Frm108.L50_Text & "]."
        If Frm108.L49_Text = 1 Then .Cells(7, 1) = "Report inventori bagi [" & Frm108.L51_Text & "] kepada/oleh cawangan [" & Frm108.L50_Text & "] dari " & TM & " hingga " & TA & "."
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Agihan"
        .Cells(8, 3) = "Cawangan"
        .Cells(8, 4) = "No. Siri Produk"
        .Cells(8, 5) = "Nama Produk"
        .Cells(8, 6) = "Berat (g)"
        .Cells(8, 7) = "Status"
        .Cells(8, 8) = "Tarikh Pulangan / Jual"
        
        For i = 1 To 8
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm108.L49_Text = 0 Then rs.Open "select * from 64_agihan_barang where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND (status='" & Frm108_LM_SEARCH_2 & "' OR status='" & Frm108_LM_SEARCH_3 & "' OR status='" & Frm108_LM_SEARCH_4 & "') order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm108.L49_Text = 1 Then rs.Open "select * from 64_agihan_barang where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND (status='" & Frm108_LM_SEARCH_2 & "' OR status='" & Frm108_LM_SEARCH_3 & "' OR status='" & Frm108_LM_SEARCH_4 & "') AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh) Then .Cells(8 + x, 2) = "'" & rs!tarikh 'Tarikh
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!cawangan) Then .Cells(8 + x, 3) = rs!cawangan 'Cawangan
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 4) = rs!no_siri_Produk 'No. Siri Produk
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 5) = rs!kategori_Produk 'Nama Produk
            
            .Cells(8 + x, 6).HorizontalAlignment = xlCenter
            If Not IsNull(rs!Berat) Then
                .Cells(8 + x, 6) = Format(rs!Berat, "#,##0.00 g") 'Berat (g)
            Else
                .Cells(8 + x, 6) = "0.00 g"
            End If
            
            If Not IsNull(rs!Status) Then
                If rs!Status = 1 Then
                    .Cells(8 + x, 7) = "Agihan"
                ElseIf rs!Status = 2 Then
                    .Cells(8 + x, 7) = "Pulang"
                ElseIf rs!Status = 3 Then
                    .Cells(8 + x, 7) = "Jual"
                End If
            Else
                .Cells(8 + x, 7) = "'-" 'Tarikh pulangan
            End If
            .Cells(8 + x, 7).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh_jual) Then
                .Cells(8 + x, 8) = "'" & rs!tarikh_jual 'Tarikh pulangan
            Else
                .Cells(8 + x, 8) = "'-" 'Tarikh pulangan
            End If
            .Cells(8 + x, 8).HorizontalAlignment = xlCenter

            For Col = 1 To 8
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 2
        
        '### Jumlah bilangan barang keseluruhan ### - Start
        .Cells(8 + Y, 1) = "Bilangan : " & 0
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm108.L49_Text = 0 Then rs.Open "select COUNT(ID) from 64_agihan_barang where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND (status='" & Frm108_LM_SEARCH_2 & "' OR status='" & Frm108_LM_SEARCH_3 & "' OR status='" & Frm108_LM_SEARCH_4 & "') order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm108.L49_Text = 1 Then rs.Open "select COUNT(ID) from 64_agihan_barang where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND (status='" & Frm108_LM_SEARCH_2 & "' OR status='" & Frm108_LM_SEARCH_3 & "' OR status='" & Frm108_LM_SEARCH_4 & "') AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic

        If Not IsNull(rs(0)) Then
            .Cells(8 + Y, 1) = "Bilangan : " & rs(0)
        Else
            .Cells(8 + Y, 1) = "Bilangan : " & 0
        End If
        
        rs.Close
        Set rs = Nothing
        '### Jumlah bilangan barang keseluruhan ### - End
        
        Y = Y + 1
        
        '### Jumlah berat barang keseluruhan ### - Start
        .Cells(8 + Y, 1) = "Jumlah berat : 0.00 g"
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm108.L49_Text = 0 Then rs.Open "select SUM(berat) from 64_agihan_barang where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND (status='" & Frm108_LM_SEARCH_2 & "' OR status='" & Frm108_LM_SEARCH_3 & "' OR status='" & Frm108_LM_SEARCH_4 & "') order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm108.L49_Text = 1 Then rs.Open "select SUM(berat) from 64_agihan_barang where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND (status='" & Frm108_LM_SEARCH_2 & "' OR status='" & Frm108_LM_SEARCH_3 & "' OR status='" & Frm108_LM_SEARCH_4 & "') AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic

        If Not IsNull(rs(0)) Then
            .Cells(8 + Y, 1) = "Jumlah berat : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Jumlah berat : 0.00 g"
        End If
        
        rs.Close
        Set rs = Nothing
        '### Jumlah berat barang keseluruhan ### - End
        
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
End Sub
Private Sub Frm108_SM_lihat_edit_data_Click()
'On Error Resume Next
Frm108_LM_ID = vbNullString
Frm108_LM_No_PEKERJA = vbNullString
DATA_FOUND = 0

If Frm108.MSFlexGrid3 <> vbNullString Then
    
    If IsNumeric(Frm108.MSFlexGrid3) Then
        Frm108_LM_ID = Frm108.MSFlexGrid3.TextMatrix(Frm108.MSFlexGrid3, 2) 'No. ID
        
        If Frm108_LM_ID <> vbNullString Then
        
            If Frm108.L48_Text = 0 Then
                Call Frm108_recall_data_agihan
            ElseIf Frm108.L48_Text = 1 Then
                Call Frm108_recall_data_pulangan
            End If
            
        End If
        
    End If
    
End If
End Sub
Private Sub Frm108_SM_padam_data_Click()
'On Error Resume Next
Frm108_LM_ID = vbNullString
DATA_FOUND = 0

If Frm108.MSFlexGrid3 <> vbNullString Then
    
    If IsNumeric(Frm108.MSFlexGrid3) Then
        Frm108_LM_ID = Frm108.MSFlexGrid3.TextMatrix(Frm108.MSFlexGrid3, 2) 'No. ID
        
        If Frm108_LM_ID <> vbNullString Then
        
            If Frm108.L48_Text = 0 Then
                Call Frm108_padam_data_agihan
            ElseIf Frm108.L48_Text = 1 Then
                Call Frm108_padam_data_pulangan
            End If
        
        End If
        
    End If
    
End If
End Sub
Private Sub Frm108_SM_remove_Click()
'On Error Resume Next
Frm108_LM_ID = vbNullString
DATA_REMOVE = 0

If Frm108.MSFlexGrid2 <> vbNullString Then
    
    If IsNumeric(Frm108.MSFlexGrid2) Then
        Frm108_LM_ID = Frm108.MSFlexGrid2.TextMatrix(Frm108.MSFlexGrid2, 2) 'No. ID
        Frm108_LM_NO_SIRI = Frm108.MSFlexGrid2.TextMatrix(Frm108.MSFlexGrid2, 3) 'No. Siri Produk
        
        If Frm108_LM_ID <> vbNullString Then

            Note = "Adakah anda yakin untuk mengeluarkan barang ini dari senarai?" & vbCrLf & _
                    "No. siri produk [" & Frm108_LM_NO_SIRI & "]." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Teruskan?"
            
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbYes Then
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from " & G_AGIHAN_TEMP & " where ID='" & Frm108_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                
                    If rs!Status = 6 Then
                        
                        MsgBox "Anda tidak dibenarkan untuk mengeluarkan barang ini dari senarai" & vbCrLf & _
                                "kerana barang ini telah dipulangkan.", vbInformation, "Info"
                                
                        rs.Close
                        Set rs = Nothing
                        
                        Exit Sub
                        
                    ElseIf rs!Status = 7 Then
                    
                        MsgBox "Anda tidak dibenarkan untuk mengeluarkan barang ini dari senarai" & vbCrLf & _
                                "kerana barang ini telah terjual.", vbInformation, "Info"
                                
                        rs.Close
                        Set rs = Nothing
                        
                        Exit Sub
                        
                    Else
                    
                        If Frm108.L1_Text = 0 Then
                        
                            rs!Status = 0
                            rs.Update
                            DATA_REMOVE = 1
                            
                        ElseIf Frm108.L1_Text = 1 Then
                        
                            If rs!Status = 3 Then
                            
                                rs!Status = 0
                                rs.Update
                                DATA_REMOVE = 1
                                
                            ElseIf rs!Status = 2 Or rs!Status = 5 Then
                            
                                rs!Status = 4
                                rs.Update
                                DATA_REMOVE = 1
                                
                            End If
                            
                        End If
                        
                    End If
                
                End If
                
                rs.Close
                Set rs = Nothing
            
                If DATA_REMOVE = 1 Then
                
                    GM_NEXT_PREV = 2
                    
                    Call Frm108_senarai_agihan_header
                    Call Frm108_senarai_agihan
                    
                    MsgBox "Barang ini telah berjaya dikeluarkan dari senarai.", vbInformation, "Info"
                
                End If
                
            End If
        
        End If
    End If
End If
End Sub
Private Sub Frm108_SM_remove2_Click()
'On Error Resume Next
Frm108_LM_ID = vbNullString
DATA_REMOVE = 0

If Frm108.MSFlexGrid5 <> vbNullString Then
    
    If IsNumeric(Frm108.MSFlexGrid5) Then
        Frm108_LM_ID = Frm108.MSFlexGrid5.TextMatrix(Frm108.MSFlexGrid5, 2) 'No. ID
        Frm108_LM_NO_SIRI = Frm108.MSFlexGrid5.TextMatrix(Frm108.MSFlexGrid5, 3) 'No. Siri Produk
        
        If Frm108_LM_ID <> vbNullString Then

            Note = "Adakah anda yakin untuk mengeluarkan barang ini dari senarai?" & vbCrLf & _
                    "No. siri produk [" & Frm108_LM_NO_SIRI & "]." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Teruskan?"
            
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbYes Then
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from " & G_PULANGAN_TEMP & " where ID='" & Frm108_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                
                    If Frm108.L39_Text = 0 Then 'Pulangan Flag , 0 : Data baru , 1 : Edit
                    
                        rs!Status = 0
                        rs.Update
                        DATA_REMOVE = 1
                        
                    ElseIf Frm108.L39_Text = 1 Then 'Pulangan Flag , 0 : Data baru , 1 : Edit
                    
                        If rs!Status = 3 Or rs!Status = 7 Then
                        
                            rs!Status = 9
                            rs.Update
                            DATA_REMOVE = 1
                            
                        ElseIf rs!Status = 4 Or rs!Status = 8 Then
                            
                            rs!Status = 10
                            rs.Update
                            DATA_REMOVE = 1
                            
                        ElseIf rs!Status = 5 Or rs!Status = 6 Then
                        
                            rs!Status = 0
                            rs.Update
                            DATA_REMOVE = 1
                            
                        End If
                        
                    End If
                
                End If
                
                rs.Close
                Set rs = Nothing
            
                If DATA_REMOVE = 1 Then
                
                    GM_NEXT_PREV = 2
                    
                    Call Frm108_senarai_pulangan_header
                    Call Frm108_senarai_pulangan
                    
                    MsgBox "Barang ini telah berjaya dikeluarkan dari senarai.", vbInformation, "Info"
                
                End If
                
            End If
        
        End If
    End If
End If
End Sub
Private Sub Frm108_SM_scan_Click()
'On Error Resume Next
If Frm108.Pic2.Visible = True Then
    Frm108.TB5 = vbNullString
    Frm108.TB5.SetFocus
End If

If Frm108.Pic6.Visible = True Then
    Frm108.TB11 = vbNullString
    Frm108.TB11.SetFocus
End If
End Sub
Private Sub Frm108_SM_senarai_barang_Click()
'On Error Resume Next
Frm108_LM_ID = vbNullString
DATA_FOUND = 0

If Frm108.MSFlexGrid3 <> vbNullString Then
    'If Button = vbRightButton Then
    
        If IsNumeric(Frm108.MSFlexGrid3) Then
            Frm108_LM_ID = Frm108.MSFlexGrid3.TextMatrix(Frm108.MSFlexGrid3, 2) 'No. ID
            
            If Frm108_LM_ID <> vbNullString Then

                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                If Frm108.L48_Text = 0 Then rs.Open "select * from 63_agihan where ID='" & Frm108_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
                If Frm108.L48_Text = 1 Then rs.Open "select * from 68_pulangan where ID='" & Frm108_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Not IsNull(rs!no_rujukan) Then
                        Frm108_LM_No_RUJ = rs!no_rujukan
                        DATA_FOUND = 1
                    End If
                End If
                
                rs.Close
                Set rs = Nothing
                
                If DATA_FOUND = 1 Then
                    Frm108.L37_Text = Frm108_LM_No_RUJ
                    
                    GM_NEXT_PREV = 0
                    
                    Frm108.L35_Text = -1 'Titik Pencarian Data
                    Frm108.L36_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
                    Frm108.L33_Text = 0 'Paparan Page ke-xxx
                    
                    If Frm108.L48_Text = 0 Then
                        Frm108.L63_Text.Visible = False
                        Frm108.L64_Text.Visible = False
                        Frm108.L65_Text.Visible = False
                        Frm108.L66_Text.Visible = False
                        Frm108.L62_Text = "Maklumat Agihan" 'Caption : Maklumat agihan / Maklumat pulangan
                        
                        Call Frm108_senarai_agihan_barang_detail_header
                        Call Frm108_senarai_agihan_barang_detail
                    End If
                    
                    If Frm108.L48_Text = 1 Then
                        Frm108.L63_Text.Visible = True
                        Frm108.L64_Text.Visible = True
                        Frm108.L65_Text.Visible = True
                        Frm108.L66_Text.Visible = True
                        Frm108.L62_Text = "Maklumat pulangan" 'Caption : Maklumat agihan / Maklumat pulangan
                        
                        Call Frm108_senarai_pulangan_barang_detail_header
                        Call Frm108_senarai_pulangan_barang_detail
                    End If
                    
                    Frm108.Pic4.Visible = False
                    Frm108.Pic5.Visible = True
                    
                End If
                
            End If
            
        End If
        
    'End If
End If
End Sub
Private Sub Frm108_SM_tidak_aktif_Click()
'On Error Resume Next
Frm108_LM_ID = vbNullString
DATA_FOUND = 0

If Frm108.MSFlexGrid1 <> vbNullString Then
    'If Button = vbRightButton Then
    
        If IsNumeric(Frm108.MSFlexGrid1) Then
            Frm108_LM_ID = Frm108.MSFlexGrid1.TextMatrix(Frm108.MSFlexGrid1, 2) 'No. ID
            
            If Frm108_LM_ID <> vbNullString Then
            
                Note = "Adakah anda ingin tukar status cawangan ini kepada TIDAK AKTIF?" & vbCrLf & _
                        "Cawangan ini tidak boleh digunakan lagi setelah diubah status ini." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Teruskan?"
                
                Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
                
                If Answer = vbYes Then

                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from 62_senarai_cawangan where ID='" & Frm108_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                        If Not IsNull(rs!cawangan) Then Frm108_LM_NAMA_CAWANGAN = rs!cawangan
                        rs!Status = 0
                        
                        rs.Update
                        DATA_FOUND = 1
                    End If
                    
                    rs.Close
                    Set rs = Nothing
                    
                    If DATA_FOUND = 1 Then
                        
                        '#### Update Log Aktiviti Sistem #### - Start
                        user = MDI_frm1.L3_Text
                        
                        LogAct_Memory = "[" & user & "] Tukar status cawangan kepada tidak aktif. Nama cawangan [" & Frm108_LM_NAMA_CAWANGAN & "]."
                        LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                        Call UpdateLog_Database
                        '#### Update Log Aktiviti Sistem #### - End
                    
                        GM_NEXT_PREV = 2

                        Call Frm108_senarai_cawangan_header
                        Call Frm108_senarai_cawangan
                        
                        MsgBox "Data telah berjaya ditukar status.", vbInformation, "Info"
                    
                    End If
                    
                End If
                
            End If
            
        End If
        
    'End If
End If
End Sub
Private Sub L2_Text_Click()
'on error resume next
If Frm108.Pic1.Visible = False Then
    Call Frm108_cmd_visible_1
    Call Frm108_initial_setting
    Call Frm108_cawangan_initial_setting
    
    GM_NEXT_PREV = 0
    
    Frm108.L8_Text = -1 'Titik Pencarian Data
    Frm108.L9_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm108.L6_Text = 0 'Paparan Page ke-xxx
    
    Call Frm108_senarai_cawangan_header
    Call Frm108_senarai_cawangan
    
    Frm108.Pic1.Visible = True
    
    Frm108.TB4.SetFocus
Else
    Frm108.Pic1.Visible = False
End If
End Sub
Private Sub L3_Text_Click()
'on error resume next
If Frm108.Pic2.Visible = False Then
    Call Frm108_cmd_visible_2
    Call Frm108_initial_setting
    Call Frm108_hantaran_initial_setting
    Call Frm108_hantaran_initial_setting2
    
    GM_NEXT_PREV = 0
    
    Frm108.L15_Text = -1 'Titik Pencarian Data
    Frm108.L16_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm108.L13_Text = 0 'Paparan Page ke-xxx
    
    Call Frm108_senarai_agihan_header
    Call Frm108_senarai_agihan
    
    Frm108.Pic2.Visible = True
    
    Frm108.TB1.SetFocus
Else
    Frm108.Pic2.Visible = False
End If
End Sub
Private Sub L38_Text_Click()
'on error resume next
If Frm108.Pic6.Visible = False Then
    Call Frm108_cmd_visible_3
    Call Frm108_initial_setting
    Call Frm108_hantaran_initial_setting
    Call Frm108_hantaran_initial_setting2
    
    GM_NEXT_PREV = 0
    
    Frm108.L46_Text = -1 'Titik Pencarian Data
    Frm108.L47_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm108.L44_Text = 0 'Paparan Page ke-xxx
    
    Call Frm108_senarai_pulangan_header
    Call Frm108_senarai_pulangan
    
    Frm108.Pic6.Visible = True
    
    Frm108.TB8.SetFocus
Else
    Frm108.Pic6.Visible = False
End If
End Sub
Private Sub L4_Text_Click()
'on error resume next
If Frm108.Pic3.Visible = False Then

    Call Frm108_initial_setting
    Call Frm108_initial_setting2
    Call Frm108_hantaran_initial_setting3
    
    'GM_NEXT_PREV = 0
    
    'Frm108.L15_Text = -1 'Titik Pencarian Data
    'Frm108.L16_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    'Frm108.L13_Text = 0 'Paparan Page ke-xxx
    
    'Call Frm108_senarai_agihan_header
    'Call Frm108_senarai_agihan
    
    Frm108.Pic3.Visible = True

Else
    Frm108.Pic3.Visible = False
End If
End Sub
Private Sub L59_Text_Click()
'on error resume next
If Frm108.Pic7.Visible = False Then

    Call Frm108_initial_setting
    Call Frm108_report_initial_setting
    
    Frm108.Pic7.Visible = True

Else
    Frm108.Pic7.Visible = False
End If
End Sub
Private Sub MSFlexGrid1_DblClick()
'On Error Resume Next
Frm108_LM_ID = vbNullString

If Frm108.MSFlexGrid1 <> vbNullString Then
    'If Button = vbRightButton Then
    
        If IsNumeric(Frm108.MSFlexGrid1) Then
            Frm108_LM_ID = Frm108.MSFlexGrid1.TextMatrix(Frm108.MSFlexGrid1, 2) 'No. ID
            
            If Frm108_LM_ID <> vbNullString Then
            
            
                user_level = MDI_frm1.L4_Text
                
                If user_level = "Admin" Or user_level = "HQ" Or user_level = "Developer" Then
                
                    Frm108.Frm108_SM_tukar_status.Enabled = True
                    Frm108.Frm108_SM_edit_data.Enabled = True
                    
                ElseIf user_level = "Manager" Then
                
                    Frm108.Frm108_SM_tukar_status.Enabled = False
                    Frm108.Frm108_SM_edit_data.Enabled = True
                
                Else
                
                    Frm108.Frm108_SM_tukar_status.Enabled = False
                    Frm108.Frm108_SM_edit_data.Enabled = False
                
                End If
 
                PopupMenu Frm108_PM_menu1, vbPopupMenuRightButton

            End If
            
        End If
        
    'End If
End If
End Sub
Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'On Error Resume Next
Frm108_LM_ID = vbNullString

If Frm108.MSFlexGrid1 <> vbNullString Then
    If Button = vbRightButton Then
    
        If IsNumeric(Frm108.MSFlexGrid1) Then
            Frm108_LM_ID = Frm108.MSFlexGrid1.TextMatrix(Frm108.MSFlexGrid1, 2) 'No. ID
            
            If Frm108_LM_ID <> vbNullString Then
            
            
                user_level = MDI_frm1.L4_Text
                
                If user_level = "Admin" Or user_level = "HQ" Or user_level = "Developer" Then
                
                    Frm108.Frm108_SM_tukar_status.Enabled = True
                    Frm108.Frm108_SM_edit_data.Enabled = True
                    
                ElseIf user_level = "Manager" Then
                
                    Frm108.Frm108_SM_tukar_status.Enabled = False
                    Frm108.Frm108_SM_edit_data.Enabled = True
                
                Else
                
                    Frm108.Frm108_SM_tukar_status.Enabled = False
                    Frm108.Frm108_SM_edit_data.Enabled = False
                
                End If
 
                PopupMenu Frm108_PM_menu1, vbPopupMenuRightButton

            End If
            
        End If
        
    End If
End If
End Sub
Private Sub MSFlexGrid2_DblClick()
'On Error Resume Next
Frm108_LM_ID = vbNullString

If Frm108.MSFlexGrid2 <> vbNullString Then
    'If Button = vbRightButton Then
    
        If IsNumeric(Frm108.MSFlexGrid2) Then
            Frm108_LM_ID = Frm108.MSFlexGrid2.TextMatrix(Frm108.MSFlexGrid2, 2) 'No. ID
            Frm108_LM_NO_SIRI = Frm108.MSFlexGrid2.TextMatrix(Frm108.MSFlexGrid2, 3) 'No. Siri Produk
            
            If Frm108_LM_ID <> vbNullString Then
                Frm108.Frm108_SM_remove.Caption = "Keluarkan dari senarai. No. siri produk [" & Frm108_LM_NO_SIRI & "]"
                PopupMenu Frm108_PM_menu4, vbPopupMenuRightButton

            End If
            
        End If
        
    'End If
End If
End Sub
Private Sub MSFlexGrid3_DblClick()
'On Error Resume Next
Frm108_LM_ID = vbNullString

If Frm108.MSFlexGrid3 <> vbNullString Then
    'If Button = vbRightButton Then
    
        If IsNumeric(Frm108.MSFlexGrid3) Then
            Frm108_LM_ID = Frm108.MSFlexGrid3.TextMatrix(Frm108.MSFlexGrid3, 2) 'No. ID
            Frm108_LM_NO_STATEMENT = Frm108.MSFlexGrid3.TextMatrix(Frm108.MSFlexGrid3, 4) 'No. Statement
            
            If Frm108_LM_ID <> vbNullString Then
                'If Frm108.CB4 = 1 Then
                
                    user_level = MDI_frm1.L4_Text
                    
                    If user_level = "Admin" Or user_level = "HQ" Or user_level = "Developer" Then
                    
                        Frm108.Frm108_SM_lihat_edit_data.Enabled = True
                        Frm108.Frm108_SM_padam_data.Enabled = True
                        
                    ElseIf user_level = "Manager" Then
                    
                        Frm108.Frm108_SM_lihat_edit_data.Enabled = True
                        Frm108.Frm108_SM_padam_data.Enabled = False
                    
                    Else
                    
                        Frm108.Frm108_SM_lihat_edit_data.Enabled = False
                        Frm108.Frm108_SM_padam_data.Enabled = False
                    
                    End If
                
                    Frm108.Frm108_SM_cetak_penyata.Caption = "Cetak penyata bagi nombor rujukan ini [" & Frm108_LM_NO_STATEMENT & "]"
                    PopupMenu Frm108_PM_menu3, vbPopupMenuRightButton
                    
                'End If
                
                'If Frm108.CB5 = 1 Then

                '    Frm108.Frm108_SM_cetak_penyata2.Caption = "Cetak penyata bagi nombor rujukan ini [" & Frm108_LM_NO_STATEMENT & "]"
                '    PopupMenu Frm108_PM_menu6, vbPopupMenuRightButton
                
                'End If

            End If
            
        End If
        
    'End If
End If
End Sub
Private Sub MSFlexGrid5_DblClick()
'On Error Resume Next
Frm108_LM_ID = vbNullString

If Frm108.MSFlexGrid5 <> vbNullString Then
    'If Button = vbRightButton Then
    
        If IsNumeric(Frm108.MSFlexGrid5) Then
            Frm108_LM_ID = Frm108.MSFlexGrid5.TextMatrix(Frm108.MSFlexGrid5, 2) 'No. ID
            Frm108_LM_NO_SIRI = Frm108.MSFlexGrid5.TextMatrix(Frm108.MSFlexGrid5, 3) 'No. Siri
            
            If Frm108_LM_ID <> vbNullString Then
                Frm108.Frm108_SM_remove2.Caption = "Keluarkan dari senarai. No. siri [" & Frm108_LM_NO_SIRI & "]"
                PopupMenu Frm108_PM_menu5, vbPopupMenuRightButton

            End If
            
        End If
        
    'End If
End If
End Sub
Private Sub MSFlexGrid6_DblClick()
'On Error Resume Next
Frm108_LM_ID = vbNullString

If Frm108.MSFlexGrid6 <> vbNullString Then
    'If Button = vbRightButton Then
    
        If IsNumeric(Frm108.MSFlexGrid6) Then
 
            PopupMenu Frm108_PM_menu6, vbPopupMenuRightButton

        End If
        
    'End If
End If
End Sub
Private Sub TB11_Change()
'on error resume next
If Frm108.CB1 = 1 And Frm108.TB11 <> vbNullString Then
    Frm108.Tmr2.Enabled = False
    Frm108.Tmr2.Enabled = True
    Frm108.Tmr2.Interval = 100
End If
End Sub
Private Sub TB5_Change()
'on error resume next
If Frm108.CB1 = 1 And Frm108.TB5 <> vbNullString Then
    Frm108.Tmr1.Enabled = False
    Frm108.Tmr1.Enabled = True
    Frm108.Tmr1.Interval = 100
End If
End Sub
Private Sub Tmr1_Timer()
'On Error Resume Next
If Frm108.CB1 = 1 And Frm108.TB5 <> vbNullString And Frm108.Tmr1.Enabled = True Then

    If Frm108.Tmr1.Interval = 100 Then
        If InStr(1, Frm108.TB5, "*") <> 0 Or InStr(1, Frm108.TB5, "/") <> 0 Or InStr(1, Frm108.TB5, "\") <> 0 Or InStr(1, Frm108.TB5, "'") <> 0 Then
        
            MsgBox "No. Siri Produk Mengandungi Simbol Yang Tidak Sah , ['].", vbInformation, "Info"
            
            Frm108.TB5 = vbNullString
            Exit Sub
        End If
        
        Call Frm108_periksa_data_barang
        
    End If
End If
End Sub
Private Sub Tmr2_Timer()
'On Error Resume Next
Dim Err(6)

If Frm108.CB3 = 1 And Frm108.TB11 <> vbNullString And Frm108.Tmr2.Enabled = True Then
        
    If InStr(1, Frm108.TB11, "*") <> 0 Or InStr(1, Frm108.TB11, "/") <> 0 Or InStr(1, Frm108.TB11, "\") <> 0 Or InStr(1, Frm108.TB11, "'") <> 0 Then
    
        MsgBox "No. Siri Produk mengandungi simbol yang tidak sah.", vbInformation, "Info"
        
        Frm108.TB11 = vbNullString
        Exit Sub
    End If
    
    If Frm108.CB7 = 0 And Frm108.CB8 = 0 Then
        x = x + 1
        Err(x) = "Sila buat pilihan jenis pulangan samada [Pulangan] atau [Dijual]"
    End If
    
    If Frm108.CB8 = 1 Then
    
        If Frm108.TB12 = vbNullString Then
            x = x + 1
            Err(x) = "Sila masukkan [No. Perjanjian A]."
        End If
        If Frm108.TB13 = vbNullString Then
            x = x + 1
            Err(x) = "Sila masukkan [No. Perjanjian B]."
        End If
        If Frm108.TB14 = vbNullString Or (Frm108.TB14 <> vbNullString And Not IsNumeric(Frm108.TB14)) Then
            x = x + 1
            Err(x) = "Sila masukkan Harga Jualan. Hanya NOMBOR dibenarkan di dalam ruangan ini."
        End If
    
    End If
    
    If x <> 0 Then
        Frm108.TB11 = vbNullString
        
        Frm6.Show
        Frm6.Pic1.Cls
        For Y = 1 To x
            Frm6.Pic1.Print Y & " - " & Err(Y)
        Next Y
        Exit Sub
    Else
    
        Call Frm108_periksa_data_barang2
    
    End If
        
End If
End Sub
