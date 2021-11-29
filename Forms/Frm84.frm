VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm84 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Jualan"
   ClientHeight    =   12735
   ClientLeft      =   225
   ClientTop       =   -795
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
   Icon            =   "Frm84.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12735
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   11295
      Left            =   17520
      ScaleHeight     =   11295
      ScaleWidth      =   19335
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   19335
      Begin VB.CommandButton CMD18 
         Caption         =   "Next"
         Height          =   810
         Left            =   18000
         MouseIcon       =   "Frm84.frx":0ECA
         MousePointer    =   99  'Custom
         Picture         =   "Frm84.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   208
         ToolTipText     =   "Tutup senarai ini."
         Top             =   10320
         Width           =   1095
      End
      Begin VB.CommandButton CMD17 
         Caption         =   "Back"
         Height          =   810
         Left            =   16800
         MouseIcon       =   "Frm84.frx":229E
         MousePointer    =   99  'Custom
         Picture         =   "Frm84.frx":25A8
         Style           =   1  'Graphical
         TabIndex        =   207
         ToolTipText     =   "Tutup senarai ini."
         Top             =   10320
         Width           =   1095
      End
      Begin VB.CommandButton CMD4 
         Caption         =   "Tutup Senarai Ini"
         Height          =   930
         Left            =   7320
         MouseIcon       =   "Frm84.frx":3672
         MousePointer    =   99  'Custom
         Picture         =   "Frm84.frx":397C
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   10320
         Width           =   3375
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   9900
         Left            =   120
         TabIndex        =   192
         Top             =   360
         Width           =   19035
         _ExtentX        =   33576
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
      Begin VB.Label L90_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L90_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1560
         TabIndex        =   213
         Top             =   10440
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L89_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L89_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1560
         TabIndex        =   212
         Top             =   10800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L87_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L87_Text"
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
         Left            =   15600
         TabIndex        =   211
         Top             =   10320
         Width           =   375
      End
      Begin VB.Label L88_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L88_Text"
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
         Left            =   16200
         TabIndex        =   210
         Top             =   10320
         Width           =   615
      End
      Begin VB.Label Label23 
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
         Left            =   14280
         TabIndex        =   209
         Top             =   10320
         Width           =   2295
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai barang yang telah dimasukkan ke dalam senarai jualan."
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   11535
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Scan Item Jualan"
      Height          =   4935
      Left            =   10080
      TabIndex        =   89
      Top             =   2400
      Visible         =   0   'False
      Width           =   11775
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Komisyen Agen Dropship"
         Height          =   1455
         Left            =   8520
         TabIndex        =   143
         Top             =   3120
         Width           =   3975
         Begin VB.TextBox TB16 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1920
            TabIndex        =   144
            Text            =   "0.00"
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Komisen         RM"
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
            Height          =   285
            Left            =   120
            TabIndex        =   145
            Top             =   390
            Width           =   2265
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Komisyen Agen Dropship"
         Height          =   1455
         Left            =   7560
         TabIndex        =   135
         Top             =   3240
         Width           =   3975
         Begin VB.TextBox TB12 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2400
            TabIndex        =   139
            Text            =   "0.00"
            Top             =   360
            Width           =   1275
         End
         Begin VB.TextBox TB13 
            Alignment       =   2  'Center
            BackColor       =   &H8000000A&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   138
            Text            =   "0.00"
            Top             =   945
            Width           =   1275
         End
         Begin VB.TextBox TB43 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   137
            Text            =   "0.00"
            Top             =   645
            Width           =   435
         End
         Begin VB.TextBox TB44 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2400
            TabIndex        =   136
            Text            =   "0.00"
            Top             =   645
            Width           =   1275
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Komisen Per Gram             RM/g"
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
            Left            =   120
            TabIndex        =   142
            Top             =   375
            Width           =   2505
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Komisen                    RM"
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
            Left            =   120
            TabIndex        =   141
            Top             =   975
            Width           =   2265
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Komisen Upah               %    RM"
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
            Left            =   120
            TabIndex        =   140
            Top             =   675
            Width           =   2385
         End
      End
      Begin VB.TextBox TB11 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9840
         Locked          =   -1  'True
         TabIndex        =   128
         Text            =   "0.00"
         Top             =   2325
         Width           =   1515
      End
      Begin VB.CheckBox CB3 
         BackColor       =   &H00FFFFFF&
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
         Left            =   7680
         TabIndex        =   127
         Top             =   1770
         Width           =   200
      End
      Begin VB.CheckBox CB2 
         BackColor       =   &H00FFFFFF&
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
         Left            =   7680
         TabIndex        =   126
         Top             =   1560
         Width           =   200
      End
      Begin VB.TextBox TB14 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9840
         Locked          =   -1  'True
         TabIndex        =   125
         Text            =   "0.00"
         Top             =   2640
         Width           =   1515
      End
      Begin VB.CheckBox CB18 
         BackColor       =   &H00FFFFFF&
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
         Left            =   7680
         TabIndex        =   124
         Top             =   1965
         Width           =   200
      End
      Begin VB.CheckBox CB12 
         BackColor       =   &H00FFFFFF&
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
         Left            =   7680
         TabIndex        =   123
         Top             =   960
         Width           =   200
      End
      Begin VB.CommandButton CMD31 
         BackColor       =   &H00FFFFC0&
         Caption         =   "SET HARGA JUALAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   10125
         MouseIcon       =   "Frm84.frx":5F46
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton CMD30 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Tambahan Aksesori"
         Height          =   350
         Left            =   0
         MouseIcon       =   "Frm84.frx":6250
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   121
         Top             =   4080
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton CMD13 
         Caption         =   "Masukkan Data Jualan"
         Height          =   350
         Left            =   1320
         MouseIcon       =   "Frm84.frx":655A
         MousePointer    =   99  'Custom
         TabIndex        =   119
         ToolTipText     =   "Masukkan data maklumat barang kemas ini ke dalam senarai jualan"
         Top             =   4200
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton CMD14 
         Caption         =   "Batal Edit Data"
         Height          =   350
         Left            =   3720
         MouseIcon       =   "Frm84.frx":6864
         MousePointer    =   99  'Custom
         TabIndex        =   118
         ToolTipText     =   "Batal urusan edit data jualan"
         Top             =   4200
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox TB9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5640
         MaxLength       =   10
         TabIndex        =   108
         Text            =   "0.00"
         Top             =   2760
         Width           =   1500
      End
      Begin VB.TextBox TB8 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   107
         Text            =   "0.00"
         Top             =   2460
         Width           =   1500
      End
      Begin VB.TextBox TB7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5640
         MaxLength       =   10
         TabIndex        =   106
         Text            =   "0.00"
         Top             =   2160
         Width           =   1500
      End
      Begin VB.TextBox TB6 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2250
         Locked          =   -1  'True
         TabIndex        =   105
         Text            =   "0.00"
         Top             =   3780
         Width           =   1500
      End
      Begin VB.TextBox TB5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2250
         MaxLength       =   10
         TabIndex        =   104
         Text            =   "0.00"
         Top             =   3060
         Width           =   1500
      End
      Begin VB.TextBox TB4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2250
         MaxLength       =   10
         TabIndex        =   103
         Text            =   "0.00"
         Top             =   2760
         Width           =   1500
      End
      Begin VB.TextBox TB3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2250
         Locked          =   -1  'True
         TabIndex        =   102
         Text            =   "0.00"
         Top             =   2460
         Width           =   1500
      End
      Begin VB.TextBox TB2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2250
         Locked          =   -1  'True
         TabIndex        =   101
         Top             =   2160
         Width           =   1500
      End
      Begin VB.TextBox TB10 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   100
         Text            =   "0.00"
         Top             =   3060
         Width           =   1500
      End
      Begin VB.TextBox TB15 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2250
         MaxLength       =   10
         TabIndex        =   99
         Text            =   "0.00"
         Top             =   3405
         Width           =   1500
      End
      Begin VB.TextBox TB22 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5640
         MaxLength       =   10
         TabIndex        =   98
         Text            =   "0.00"
         Top             =   3405
         Width           =   1500
      End
      Begin VB.TextBox TB1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   95
         Top             =   1100
         Width           =   2820
      End
      Begin VB.CommandButton CMD1 
         Caption         =   "Carian Data"
         Height          =   375
         Left            =   4320
         MouseIcon       =   "Frm84.frx":6B6E
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   94
         ToolTipText     =   "Carian maklumat tentang barang kemas / stok"
         Top             =   1040
         Width           =   2055
      End
      Begin VB.CheckBox CB1 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   92
         Top             =   735
         Width           =   200
      End
      Begin VB.CommandButton CMD3 
         Caption         =   "Masukkan Data Jualan"
         Height          =   350
         Left            =   2760
         MouseIcon       =   "Frm84.frx":6E78
         MousePointer    =   99  'Custom
         TabIndex        =   120
         ToolTipText     =   "Masukkan data maklumat barang kemas ini ke dalam senarai jualan"
         Top             =   4200
         Width           =   2295
      End
      Begin VB.Shape Shape6 
         Height          =   2595
         Left            =   120
         Top             =   1530
         Width           =   7275
      End
      Begin VB.Label L8_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   8520
         TabIndex        =   132
         Top             =   2355
         Width           =   840
      End
      Begin VB.Shape Shape13 
         Height          =   2475
         Left            =   7560
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label57 
         BackStyle       =   0  'Transparent
         Caption         =   "Zero Rated ZR(L)                           Standard Rated (SR)                      Standard Rated (SR) Inclusive  "
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
         Height          =   855
         Left            =   7920
         TabIndex        =   131
         Top             =   1560
         Width           =   3105
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Harga jualan dengan GST RM"
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
         Height          =   285
         Left            =   7680
         TabIndex        =   130
         Top             =   2670
         Width           =   2265
      End
      Begin VB.Shape Shape12 
         BackColor       =   &H00000000&
         FillStyle       =   7  'Diagonal Cross
         Height          =   135
         Left            =   7680
         Top             =   1395
         Width           =   3705
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "Tidak Bertanda : GST pada harga barang Bertanda : GST pada UPAH"
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
         Height          =   405
         Left            =   8040
         TabIndex        =   129
         Top             =   960
         Width           =   3105
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Berat Jualan                  g                                     Adjustment                RM"
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
         TabIndex        =   116
         Top             =   2760
         Width           =   6825
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Berat Asal                      g                                    Harga Lepas Diskaun  RM"
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
         TabIndex        =   115
         Top             =   2475
         Width           =   6705
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Siri Produk                                                     Diskaun                      %"
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
         TabIndex        =   114
         Top             =   2175
         Width           =   6705
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Asal                 RM"
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
         TabIndex        =   113
         Top             =   3795
         Width           =   2265
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Semasa         RM/g                                     Harga Jualan             RM"
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
         TabIndex        =   112
         Top             =   3075
         Width           =   6705
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Upah                          RM                                      Upah per gram       RM/g"
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
         TabIndex        =   111
         Top             =   3435
         Width           =   5505
      End
      Begin VB.Label L67_Text 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   6720
         TabIndex        =   110
         Top             =   3720
         Width           =   1245
      End
      Begin VB.Label L69_Text 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   4920
         TabIndex        =   109
         Top             =   3720
         Width           =   1005
      End
      Begin VB.Shape Shape2 
         Height          =   375
         Left            =   240
         Top             =   3360
         Width           =   7095
      End
      Begin VB.Label L40_Text 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm84.frx":7182
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   525
         Left            =   240
         TabIndex        =   97
         Top             =   1560
         Visible         =   0   'False
         Width           =   7035
      End
      Begin VB.Shape Shape1 
         Height          =   855
         Left            =   120
         Top             =   600
         Width           =   7275
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Siri Produk :"
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
         Height          =   225
         Left            =   0
         TabIndex        =   96
         Top             =   1120
         Width           =   1425
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Scanner Mode   (Sila scan setiap barang yang akan dijual di sini)"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   525
         TabIndex        =   93
         Top             =   720
         Width           =   6930
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Caption         =   "Maklumat Produk                                                                                              Maklumat GST"
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
         Height          =   315
         Left            =   240
         TabIndex        =   91
         Top             =   360
         Width           =   11175
      End
      Begin VB.Label Label89 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila klik F2 untuk scan barang yang hendak dijual."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2280
         TabIndex        =   90
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label Label59 
         BackStyle       =   0  'Transparent
         Caption         =   "@       %"
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
         Left            =   8640
         TabIndex        =   134
         Top             =   2355
         Width           =   840
      End
      Begin VB.Label Label58 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah GST                       RM"
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
         Left            =   7680
         TabIndex        =   133
         Top             =   2355
         Width           =   2640
      End
      Begin VB.Label L68_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Modal (RM)   :                      Jual (RM) :"
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
         Left            =   3840
         TabIndex        =   117
         Top             =   3720
         Width           =   3585
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Maklumat Trade In (0%)"
      Height          =   2295
      Left            =   3120
      TabIndex        =   214
      Top             =   4920
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton CMD20 
         Caption         =   "Batal Trade In"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   840
         MouseIcon       =   "Frm84.frx":7221
         MousePointer    =   99  'Custom
         TabIndex        =   220
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox TB49 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3480
         TabIndex        =   215
         Text            =   "TB49"
         Top             =   290
         Width           =   1155
      End
      Begin VB.TextBox TB52 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3480
         TabIndex        =   218
         Text            =   "TB52"
         Top             =   1200
         Width           =   1155
      End
      Begin VB.TextBox TB51 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3480
         TabIndex        =   217
         Text            =   "TB51"
         Top             =   915
         Width           =   1155
      End
      Begin VB.CommandButton CMD9 
         Caption         =   "KemasKini Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2760
         MouseIcon       =   "Frm84.frx":752B
         MousePointer    =   99  'Custom
         TabIndex        =   219
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox TB50 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3480
         TabIndex        =   216
         Text            =   "TB50"
         Top             =   600
         Width           =   1155
      End
      Begin VB.Shape Shape4 
         Height          =   615
         Left            =   0
         Top             =   2880
         Width           =   5055
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "g       RM / g  RM / g RM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1125
         Left            =   2760
         TabIndex        =   222
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "      Berat Trade In :      Harga Semasa Trade In :      Harga Semasa Buyback :      Caj Pertukaran :      "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1605
         Left            =   120
         TabIndex        =   221
         Top             =   360
         Width           =   2985
      End
   End
   Begin VB.PictureBox Pic8 
      BorderStyle     =   0  'None
      Height          =   2720
      Left            =   18360
      ScaleHeight     =   2715
      ScaleWidth      =   7170
      TabIndex        =   77
      Top             =   3480
      Visible         =   0   'False
      Width           =   7170
      Begin VB.CommandButton CMD29 
         Caption         =   "Tutup Paparan Ini"
         Height          =   350
         Left            =   2040
         MouseIcon       =   "Frm84.frx":7835
         MousePointer    =   99  'Custom
         TabIndex        =   80
         Top             =   2160
         Width           =   2895
      End
      Begin VB.ComboBox CBB3 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Frm84.frx":7B3F
         Left            =   1800
         List            =   "Frm84.frx":7B41
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   480
         Width           =   4965
      End
      Begin VB.ComboBox CBB4 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Frm84.frx":7B43
         Left            =   1800
         List            =   "Frm84.frx":7B45
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   840
         Width           =   4965
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kategori Produk :"
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   120
         TabIndex        =   84
         Top             =   480
         Width           =   1560
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan maklumat berkenaan produk yang akan dijual."
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
         Height          =   420
         Left            =   120
         TabIndex        =   83
         Top             =   120
         Width           =   6840
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Purity :"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   82
         Top             =   840
         Width           =   1560
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Setelah pilihan di atas telah dibuat , sila tutup paparan ini dan isikan maklumat jualan barang ini."
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
         Height          =   645
         Left            =   0
         TabIndex        =   81
         Top             =   1320
         Width           =   6930
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Invoice"
      Height          =   9615
      Left            =   11520
      TabIndex        =   146
      Top             =   0
      Visible         =   0   'False
      Width           =   7815
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Height          =   2655
         Left            =   360
         TabIndex        =   194
         Top             =   4800
         Width           =   7335
         Begin VB.TextBox TB37 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4800
            TabIndex        =   197
            Text            =   "0.00"
            Top             =   1875
            Width           =   1275
         End
         Begin VB.TextBox TB36 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4800
            TabIndex        =   196
            Text            =   "0"
            Top             =   1560
            Width           =   1275
         End
         Begin VB.TextBox TB35 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4800
            TabIndex        =   195
            Text            =   "0"
            Top             =   525
            Width           =   1275
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   $"Frm84.frx":7B47
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   2685
            Left            =   240
            TabIndex        =   203
            Top             =   240
            Width           =   4905
         End
         Begin VB.Label L78_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5220
            TabIndex        =   202
            Top             =   2145
            Width           =   2505
         End
         Begin VB.Label L77_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4800
            TabIndex        =   201
            Top             =   1320
            Width           =   2505
         End
         Begin VB.Label L76_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4800
            TabIndex        =   200
            Top             =   825
            Width           =   2505
         End
         Begin VB.Label L75_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5220
            TabIndex        =   199
            Top             =   240
            Width           =   2505
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   ": RM  :        :                  :        :        :        : RM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   2325
            Left            =   4680
            TabIndex        =   198
            Top             =   240
            Width           =   585
         End
      End
      Begin VB.TextBox TB19 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5160
         TabIndex        =   152
         Text            =   "0.00"
         Top             =   1200
         Width           =   2115
      End
      Begin VB.TextBox TB42 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5160
         TabIndex        =   151
         Text            =   "0.00"
         Top             =   1830
         Width           =   2115
      End
      Begin VB.TextBox TB20 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5160
         TabIndex        =   150
         Text            =   "0.00"
         Top             =   2475
         Width           =   2115
      End
      Begin VB.TextBox TB45 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5160
         TabIndex        =   149
         Text            =   "0.00"
         Top             =   2160
         Width           =   2115
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
         Left            =   120
         TabIndex        =   148
         Top             =   2835
         Width           =   200
      End
      Begin VB.TextBox TB34 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   147
         Text            =   "0.00"
         Top             =   2790
         Width           =   2115
      End
      Begin VB.Label L38_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L38_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         TabIndex        =   168
         Top             =   7800
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label L79_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L79_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         TabIndex        =   167
         Top             =   7800
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label L23_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   5160
         TabIndex        =   166
         Top             =   8130
         Width           =   2505
      End
      Begin VB.Label L24_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Pelanggan Perlu Bayar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   165
         Top             =   8130
         Width           =   5505
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm84.frx":7CF1
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4485
         Left            =   120
         TabIndex        =   164
         Top             =   240
         Width           =   5025
      End
      Begin VB.Label L17_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5160
         TabIndex        =   163
         Top             =   240
         Width           =   2145
      End
      Begin VB.Label L22_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5160
         TabIndex        =   162
         Top             =   4035
         Width           =   2505
      End
      Begin VB.Label L18_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5160
         TabIndex        =   161
         Top             =   555
         Width           =   2505
      End
      Begin VB.Label L19_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5160
         TabIndex        =   160
         Top             =   870
         Width           =   2505
      End
      Begin VB.Label L20_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5160
         TabIndex        =   159
         Top             =   1530
         Width           =   2505
      End
      Begin VB.Label L37_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5160
         TabIndex        =   158
         Top             =   4335
         Width           =   3705
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm84.frx":7F66
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   9285
         Left            =   4560
         TabIndex        =   157
         Top             =   240
         Width           =   705
      End
      Begin VB.Label L21_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5160
         TabIndex        =   156
         Top             =   3390
         Width           =   2505
      End
      Begin VB.Label L73_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5160
         TabIndex        =   155
         Top             =   3120
         Width           =   2505
      End
      Begin VB.Label L74_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5160
         TabIndex        =   154
         Top             =   3720
         Width           =   2505
      End
      Begin VB.Label L80_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L80_Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   2520
         TabIndex        =   153
         Top             =   2760
         Width           =   1905
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Maklumat Trade In"
      Height          =   4095
      Left            =   1680
      TabIndex        =   181
      Top             =   360
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000003&
         Caption         =   "Batal Belian Dengan Trade In"
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
         Left            =   120
         MouseIcon       =   "Frm84.frx":8035
         MousePointer    =   99  'Custom
         Picture         =   "Frm84.frx":833F
         Style           =   1  'Graphical
         TabIndex        =   193
         Top             =   2760
         Width           =   5175
      End
      Begin VB.CommandButton CMD8 
         Caption         =   "Reset Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1680
         MouseIcon       =   "Frm84.frx":A909
         MousePointer    =   99  'Custom
         TabIndex        =   191
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox TB17 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   189
         Text            =   "0.00"
         Top             =   1680
         Width           =   1725
      End
      Begin VB.TextBox TB18 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         TabIndex        =   183
         Top             =   600
         Width           =   1635
      End
      Begin VB.CommandButton CMD7 
         Caption         =   "Carian Maklumat Voucher"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MouseIcon       =   "Frm84.frx":AC13
         MousePointer    =   99  'Custom
         TabIndex        =   182
         Top             =   600
         Width           =   2175
      End
      Begin VB.Shape Shape15 
         Height          =   1935
         Left            =   120
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Voucher  : RM"
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
         Left            =   240
         TabIndex        =   190
         Top             =   1710
         Width           =   1635
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Voucher :"
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
         Left            =   240
         TabIndex        =   188
         Top             =   1395
         Width           =   1635
      End
      Begin VB.Label L16_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   1920
         TabIndex        =   187
         Top             =   1395
         Width           =   2265
      End
      Begin VB.Label Label34 
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
         Height          =   255
         Left            =   240
         TabIndex        =   186
         Top             =   1080
         Width           =   4275
      End
      Begin VB.Shape Shape14 
         Height          =   615
         Left            =   195
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Voucher:"
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
         Height          =   285
         Left            =   240
         TabIndex        =   185
         Top             =   630
         Width           =   945
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Carian Voucher Buyback / Trade In"
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
         Height          =   285
         Left            =   240
         TabIndex        =   184
         Top             =   360
         Width           =   4275
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Maklumat Cukai GST"
      Height          =   3135
      Left            =   4440
      TabIndex        =   169
      Top             =   720
      Visible         =   0   'False
      Width           =   5250
      Begin VB.Shape Shape9 
         Height          =   1140
         Left            =   120
         Top             =   240
         Width           =   4935
      End
      Begin VB.Shape Shape8 
         Height          =   1380
         Left            =   120
         Top             =   1560
         Width           =   4935
      End
      Begin VB.Label Label122 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Zero Rated ZR(L)    (RM) : Standard Rated SR (RM):"
         ForeColor       =   &H00000000&
         Height          =   660
         Left            =   120
         TabIndex        =   180
         Top             =   2340
         Width           =   2520
      End
      Begin VB.Label Label60 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Harga   Cukai GST"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2520
         TabIndex        =   179
         Top             =   2040
         Width           =   2640
      End
      Begin VB.Label L7_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2520
         TabIndex        =   178
         Top             =   2340
         Width           =   1200
      End
      Begin VB.Label L9_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3840
         TabIndex        =   177
         Top             =   2340
         Width           =   1005
      End
      Begin VB.Label L10_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2520
         TabIndex        =   176
         Top             =   2595
         Width           =   1200
      End
      Begin VB.Label L11_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3840
         TabIndex        =   175
         Top             =   2595
         Width           =   1005
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Tanpa GST   : RM"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   360
         TabIndex        =   173
         Top             =   720
         Width           =   2520
      End
      Begin VB.Label L5_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2640
         TabIndex        =   172
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label L6_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2640
         TabIndex        =   171
         Top             =   975
         Width           =   1140
      End
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Dengan GST : RM"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   360
         TabIndex        =   170
         Top             =   975
         Width           =   2505
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm84.frx":AF1D
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1725
         Left            =   240
         TabIndex        =   174
         Top             =   360
         Width           =   3240
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8640
      Top             =   9000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm84.frx":B025
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm84.frx":D5FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm84.frx":FBD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm84.frx":121B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm84.frx":1478D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm84.frx":16D67
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm84.frx":19341
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm84.frx":1B91B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   11175
      Left            =   120
      TabIndex        =   88
      Top             =   240
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
   Begin VB.CommandButton CMD11 
      Caption         =   "Maklumat Agen Dropship"
      Enabled         =   0   'False
      Height          =   1170
      Left            =   13560
      MouseIcon       =   "Frm84.frx":1DEF5
      MousePointer    =   99  'Custom
      Picture         =   "Frm84.frx":1E1FF
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CommandButton CMD12 
      Caption         =   "Info Pembeli - (Berdaftar)"
      Height          =   1170
      Left            =   13560
      MouseIcon       =   "Frm84.frx":207C9
      MousePointer    =   99  'Custom
      Picture         =   "Frm84.frx":20AD3
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CommandButton CMD21 
      Caption         =   "Info Pembeli - (Tidak berdaftar)"
      Height          =   1170
      Left            =   13560
      MouseIcon       =   "Frm84.frx":2309D
      MousePointer    =   99  'Custom
      Picture         =   "Frm84.frx":233A7
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Timer Tmr3 
      Interval        =   70
      Left            =   13800
      Top             =   0
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
      Left            =   13680
      TabIndex        =   60
      Top             =   8760
      Width           =   200
   End
   Begin VB.PictureBox Pic6 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   3165
      Left            =   12600
      ScaleHeight     =   3165
      ScaleWidth      =   3465
      TabIndex        =   19
      Top             =   3480
      Visible         =   0   'False
      Width           =   3465
      Begin VB.CommandButton CMD26 
         Caption         =   "Batal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         MouseIcon       =   "Frm84.frx":25971
         MousePointer    =   99  'Custom
         TabIndex        =   70
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton CMD25 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   480
         MouseIcon       =   "Frm84.frx":25C7B
         MousePointer    =   99  'Custom
         TabIndex        =   69
         Top             =   2640
         Width           =   1215
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
         Left            =   240
         TabIndex        =   52
         Top             =   2320
         Width           =   200
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
         Left            =   240
         TabIndex        =   49
         Top             =   1080
         Width           =   200
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
         Left            =   240
         TabIndex        =   48
         Top             =   840
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
         Left            =   240
         TabIndex        =   47
         Top             =   600
         Width           =   200
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
         Left            =   240
         TabIndex        =   46
         Top             =   1320
         Width           =   200
      End
      Begin VB.CheckBox CB10 
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
         TabIndex        =   45
         Top             =   1560
         Width           =   200
      End
      Begin VB.Label L65_Text 
         Caption         =   "L65_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2640
         TabIndex        =   59
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L64_Text 
         Caption         =   "L64_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2640
         TabIndex        =   58
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm84.frx":25F85
         ForeColor       =   &H00000000&
         Height          =   2085
         Left            =   480
         TabIndex        =   50
         Top             =   600
         Width           =   2370
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm84.frx":2606F
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2565
         Left            =   480
         TabIndex        =   51
         Top             =   240
         Width           =   2370
      End
   End
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
      Left            =   13680
      TabIndex        =   53
      Top             =   8400
      Width           =   200
   End
   Begin VB.TextBox TB41 
      BackColor       =   &H8000000A&
      Height          =   360
      Left            =   17640
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2100
      Width           =   3060
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
      Left            =   13560
      TabIndex        =   35
      Top             =   2160
      Width           =   200
   End
   Begin VB.TextBox TB33 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   13560
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "0.00"
      Top             =   840
      Width           =   7125
   End
   Begin VB.ComboBox CBB1 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   360
      Left            =   15180
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   7875
      Width           =   5565
   End
   Begin VB.Timer Tmr2 
      Interval        =   100
      Left            =   13080
      Top             =   0
   End
   Begin VB.Timer Tmr1 
      Interval        =   10
      Left            =   12240
      Top             =   0
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   360
      Left            =   15180
      TabIndex        =   13
      Top             =   7515
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
      Format          =   167510016
      CurrentDate     =   41561
   End
   Begin VB.CommandButton CMD16 
      BackColor       =   &H80000003&
      Caption         =   "Keluar"
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
      Left            =   10560
      MouseIcon       =   "Frm84.frx":26163
      MousePointer    =   99  'Custom
      Picture         =   "Frm84.frx":2646D
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   10320
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton CMD2 
      Caption         =   "Keluar"
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
      Left            =   10560
      MouseIcon       =   "Frm84.frx":28A37
      MousePointer    =   99  'Custom
      Picture         =   "Frm84.frx":28D41
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   10320
      Width           =   2775
   End
   Begin VB.TextBox TB46 
      BackColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   2520
      MultiLine       =   -1  'True
      TabIndex        =   204
      Top             =   8850
      Width           =   5700
   End
   Begin VB.CommandButton CMD15 
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
      Left            =   7680
      MouseIcon       =   "Frm84.frx":2B30B
      MousePointer    =   99  'Custom
      Picture         =   "Frm84.frx":2B615
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   10320
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton CMD5 
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
      Left            =   7680
      MouseIcon       =   "Frm84.frx":2DBDF
      MousePointer    =   99  'Custom
      Picture         =   "Frm84.frx":2DEE9
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   10320
      Width           =   2775
   End
   Begin VB.Label L91_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "Trade In : 888.88 g X RM 193.00 = RM 171,553.88"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2895
      Left            =   1680
      TabIndex        =   223
      Top             =   5040
      Visible         =   0   'False
      Width           =   11895
   End
   Begin VB.Label L86_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PENGIRAAN UPAH MENGIKUT UPAH PER ITEM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   1680
      TabIndex        =   206
      Top             =   7800
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks :"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1080
      TabIndex        =   205
      Top             =   8820
      Width           =   1425
   End
   Begin VB.Label L85_Text 
      BackColor       =   &H00C0FFFF&
      Caption         =   "L85_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1560
      TabIndex        =   87
      Top             =   12720
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label L84_Text 
      BackColor       =   &H00C0FFFF&
      Caption         =   "L84_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1560
      TabIndex        =   86
      Top             =   11880
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label L83_Text 
      BackColor       =   &H00C0FFFF&
      Caption         =   "L83_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   85
      Top             =   12720
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label L42_Text 
      BackColor       =   &H00C0FFFF&
      Caption         =   "L42_Text"
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
      Left            =   11760
      TabIndex        =   76
      Top             =   10080
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label L72_Text 
      BackColor       =   &H00C0FFFF&
      Caption         =   "L72_Text"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   65
      Top             =   11520
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label L71_Text 
      BackColor       =   &H00C0FFFF&
      Caption         =   "L71_Text"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5520
      TabIndex        =   64
      Top             =   12480
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label L70_Text 
      BackColor       =   &H00C0FFFF&
      Caption         =   "L70_Text"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6600
      TabIndex        =   63
      Top             =   12120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label L66_Text 
      BackColor       =   &H00C0FFFF&
      Caption         =   "L66_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   62
      Top             =   12240
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   $"Frm84.frx":304B3
      ForeColor       =   &H000000FF&
      Height          =   780
      Left            =   13920
      TabIndex        =   61
      Top             =   8730
      Width           =   6330
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   ":   :"
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   15000
      TabIndex        =   57
      Top             =   2445
      Width           =   240
   End
   Begin VB.Label L63_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "Kategori pelanggan           : "
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   13560
      TabIndex        =   56
      Top             =   3360
      Width           =   5145
   End
   Begin VB.Label L62_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "Jualan oleh agen dropship : TIDAK"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   13560
      TabIndex        =   55
      Top             =   3120
      Width           =   3345
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Jualan secara online (Sila tanda di sini jika jualan dibuat secara online)"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   13920
      TabIndex        =   54
      Top             =   8355
      Width           =   7050
   End
   Begin VB.Label L60_Text 
      Caption         =   "L60_Text"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6600
      TabIndex        =   44
      Top             =   11040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label L61_Text 
      Caption         =   "L61_Text"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6600
      TabIndex        =   43
      Top             =   11400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label L59_Text 
      Caption         =   "L59_Text"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5520
      TabIndex        =   42
      Top             =   12120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label L56_Text 
      Caption         =   "L56_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5520
      TabIndex        =   41
      Top             =   11040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label L57_Text 
      Caption         =   "L57_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5520
      TabIndex        =   40
      Top             =   11400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label L58_Text 
      Caption         =   "L58_Text"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5520
      TabIndex        =   39
      Top             =   11760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label L55_Text 
      Caption         =   "L55_Text"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3840
      TabIndex        =   38
      Top             =   12720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label L54_Text 
      Caption         =   "L54_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3840
      TabIndex        =   37
      Top             =   12360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label91 
      BackStyle       =   0  'Transparent
      Caption         =   "Easy Payment Plan (EPP)    Approval Code :"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   13800
      TabIndex        =   36
      Top             =   2115
      Width           =   3930
   End
   Begin VB.Label L53_Text 
      Caption         =   "L53_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3840
      TabIndex        =   34
      Top             =   12000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label L52_Text 
      Caption         =   "L52_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3840
      TabIndex        =   33
      Top             =   11520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label L51_Text 
      Caption         =   "L51_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2760
      TabIndex        =   32
      Top             =   12720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label L50_Text 
      Caption         =   "L50_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2760
      TabIndex        =   31
      Top             =   12360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label L49_Text 
      Caption         =   "L49_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2760
      TabIndex        =   30
      Top             =   12000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label L48_Text 
      Caption         =   "L48_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2760
      TabIndex        =   29
      Top             =   11520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label L46_Text 
      Caption         =   "L46_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6720
      TabIndex        =   28
      Top             =   11760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label L45_Text 
      Caption         =   "L45_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7800
      TabIndex        =   27
      Top             =   11760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label L29_Text 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   16800
      TabIndex        =   26
      Top             =   6120
      Width           =   6225
   End
   Begin VB.Label L28_Text 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   16800
      TabIndex        =   23
      Top             =   4920
      Width           =   4065
   End
   Begin VB.Label L44_Text 
      Caption         =   "L44_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   12720
      TabIndex        =   22
      Top             =   9720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label L41_Text 
      BackColor       =   &H00C0FFFF&
      Caption         =   "L41_Text"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6600
      TabIndex        =   21
      Top             =   12480
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label L39_Text 
      BackColor       =   &H00C0FFFF&
      Caption         =   "L39_Text"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7800
      TabIndex        =   20
      Top             =   12240
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label L34_Text 
      BackColor       =   &H00C0FFFF&
      Caption         =   "L34_Text"
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
      Left            =   11760
      TabIndex        =   18
      Top             =   9720
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label L25_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Bayaran"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   13560
      TabIndex        =   17
      Top             =   360
      Width           =   7995
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Tarikh Jualan  * :"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   13635
      TabIndex        =   15
      Top             =   7560
      Width           =   2385
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Pekerja * :"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   13635
      TabIndex        =   14
      Top             =   7920
      Width           =   2295
   End
   Begin VB.Label L14_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   15120
      TabIndex        =   9
      Top             =   2445
      Width           =   1320
   End
   Begin VB.Label L15_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   15120
      TabIndex        =   8
      Top             =   2700
      Width           =   1320
   End
   Begin VB.Label L13_Text 
      BackColor       =   &H00C0FFFF&
      Caption         =   "L13_Text"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   12360
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label L12_Text 
      BackColor       =   &H00C0FFFF&
      Caption         =   "L12_Text"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   11520
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label L4_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1140
      Left            =   9840
      TabIndex        =   5
      Top             =   9195
      Width           =   2295
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai barang yang telah dimasukkan ke dalam senarai :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   350
      Left            =   1560
      TabIndex        =   4
      Top             =   9840
      Width           =   8775
   End
   Begin VB.Label L3_Text 
      BackColor       =   &H00C0FFFF&
      Caption         =   "L3_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   11880
      Visible         =   0   'False
      Width           =   945
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
      Left            =   21465
      TabIndex        =   0
      Top             =   315
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Barang      Jumlah Berat (g)"
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   13560
      TabIndex        =   10
      Top             =   2445
      Width           =   1560
   End
   Begin VB.Label L27_Text 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   16800
      TabIndex        =   24
      Top             =   3720
      Width           =   4305
   End
   Begin VB.Label Label80 
      BackStyle       =   0  'Transparent
      Caption         =   $"Frm84.frx":30559
      ForeColor       =   &H00000000&
      Height          =   2865
      Left            =   16080
      TabIndex        =   25
      Top             =   3720
      Width           =   825
   End
   Begin VB.Menu Frm84_PM_Menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm84_SM_Edit 
         Caption         =   "Edit Data Ini"
      End
      Begin VB.Menu Frm84_SM_Padam 
         Caption         =   "Keluarkan Item Ini Dari Senarai Jualan"
      End
   End
   Begin VB.Menu Frm84_PM_Menu2 
      Caption         =   "Scan Mode (F2)"
      Begin VB.Menu Frm84_scan_mode 
         Caption         =   "Scan Mode"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu Frm84_PM_Menu3 
      Caption         =   "Reset / Batal (F3)"
      Begin VB.Menu Frm84_SM_reset 
         Caption         =   "Reset / Batal Jualan"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu Frm84_PM_Menu4 
      Caption         =   "Tukar Kategori Pembeli / Jualan Oleh agen (F4)"
      Begin VB.Menu Frm84_SM_tukar_kategori 
         Caption         =   "Tukar Kategori Pembeli / Jualan Oleh agen"
         Shortcut        =   {F4}
      End
   End
End
Attribute VB_Name = "Frm84"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CB10_Click()
'on error resume next
If Frm84.CB10 = 1 Then
    Frm84.CB5 = 0
    Frm84.CB6 = 0
    Frm84.CB9 = 0
    Frm84.CB4 = 0
    
    Frm84.L45_Text = 5
    Frm84.L63_Text = "Kategori pelanggan           : Platinum"
End If
End Sub
Private Sub CB12_Click()
'On Error Resume Next
Dim frm84_LM_KADAR_GST As Double
Dim Frm84_LM_HARGA As Double
Dim Frm84_LM_GST As Double

Call frm84_kiraan_gst
Call Frm84_modal_dan_jual

Exit Sub

Frm84_LM_HARGA = 0

If Frm84.CB18 = 0 Then
    If Frm84.CB3 = 1 And (Frm84.TB10 <> vbNullString And IsNumeric(Frm84.TB10)) And (Frm84.L8_Text <> vbNullString And IsNumeric(Frm84.L8_Text)) Then
        frm84_LM_KADAR_GST = Frm84.L8_Text 'Jumlah Kadar GST (%)

        If Frm84.CB12 = 0 Then
            Frm84_LM_HARGA = Frm84.TB10 'Harga (RM)
        Else
            If Frm84.TB15 <> vbNullString And IsNumeric(Frm84.TB15) Then
                Frm84_LM_HARGA = Frm84.TB15 'Upah (RM)
            End If
        End If

        
        Frm84.TB11 = Format((frm84_LM_KADAR_GST / 100) * Frm84_LM_HARGA, "#,##0.00") 'Jumlah Cukai GST (RM)
        Frm84.L44_Text = Format(Frm84_LM_HARGA, "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
    Else
        Frm84.TB11 = Format(0, "#,##0.00") 'Jumlah Cukai GST (RM)
    End If
ElseIf Frm84.CB18 = 1 Then
    If Frm84.CB3 = 1 And (Frm84.TB10 <> vbNullString And IsNumeric(Frm84.TB10)) And (Frm84.L8_Text <> vbNullString And IsNumeric(Frm84.L8_Text)) Then
        frm84_LM_KADAR_GST = Frm84.L8_Text 'Jumlah Kadar GST (%)

        If Frm84.CB12 = 0 Then
            Frm84_LM_HARGA = Frm84.TB10 'Harga (RM)
        Else
            If Frm84.TB15 <> vbNullString And IsNumeric(Frm84.TB15) Then
                Frm84_LM_HARGA = Frm84.TB15 'Upah (RM)
            End If
        End If
        
        Frm84.L44_Text = Format(Frm84_LM_HARGA / (1 + (frm84_LM_KADAR_GST / 100)), "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
        Frm84.TB11 = Format(Frm84_LM_HARGA - (Frm84_LM_HARGA / (1 + (frm84_LM_KADAR_GST / 100))), "#,##0.00") 'Jumlah Cukai GST (RM)
    Else
        Frm84.TB11 = "0.00" 'Jumlah Cukai GST (RM)
    End If
End If

Call Frm84_modal_dan_jual
End Sub
Private Sub CB14_Click()
'On Error Resume Next
Call Frm84_kiraan_potongan_kupon
End Sub
Private Sub CB18_Click()
'On Error Resume Next
Dim frm84_LM_KADAR_GST As Double
Dim Frm84_LM_HARGA As Double

If Frm84.CB18 = 1 Then
    Frm84.CB2 = 0
    Frm84.CB3 = 0
End If

Call frm84_kiraan_gst
Call Frm84_modal_dan_jual

Exit Sub

Frm84_LM_HARGA = 0

If Frm84.CB18 = 0 Then
    If Frm84.CB3 = 1 And (Frm84.TB10 <> vbNullString And IsNumeric(Frm84.TB10)) And (Frm84.L8_Text <> vbNullString And IsNumeric(Frm84.L8_Text)) Then
        frm84_LM_KADAR_GST = Frm84.L8_Text 'Jumlah Kadar GST (%)

        If Frm84.CB12 = 0 Then
            Frm84_LM_HARGA = Frm84.TB10 'Harga (RM)
        Else
            If Frm84.TB15 <> vbNullString And IsNumeric(Frm84.TB15) Then
                Frm84_LM_HARGA = Frm84.TB15 'Upah (RM)
            End If
        End If
        
        Frm84.TB11 = Format((frm84_LM_KADAR_GST / 100) * Frm84_LM_HARGA, "#,##0.00") 'Jumlah Cukai GST (RM)
        Frm84.L44_Text = Format(Frm84_LM_HARGA, "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
    Else
        Frm84.TB11 = Format(0, "#,##0.00") 'Jumlah Cukai GST (RM)
    End If
ElseIf Frm84.CB18 = 1 Then
    If Frm84.CB3 = 1 And (Frm84.TB10 <> vbNullString And IsNumeric(Frm84.TB10)) And (Frm84.L8_Text <> vbNullString And IsNumeric(Frm84.L8_Text)) Then
        frm84_LM_KADAR_GST = Frm84.L8_Text 'Jumlah Kadar GST (%)

        If Frm84.CB12 = 0 Then
            Frm84_LM_HARGA = Frm84.TB10 'Harga (RM)
        Else
            If Frm84.TB15 <> vbNullString And IsNumeric(Frm84.TB15) Then
                Frm84_LM_HARGA = Frm84.TB15 'Upah (RM)
            End If
        End If
        
        Frm84.L44_Text = Format(Frm84_LM_HARGA / (1 + (frm84_LM_KADAR_GST / 100)), "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
        Frm84.TB11 = Format(Frm84_LM_HARGA - (Frm84_LM_HARGA / (1 + (frm84_LM_KADAR_GST / 100))), "#,##0.00") 'Jumlah Cukai GST (RM)
    Else
        Frm84.TB11 = "0.00" 'Jumlah Cukai GST (RM)
    End If
End If

Call Frm84_modal_dan_jual
End Sub
Private Sub CB19_Click()
'On Error Resume Next
If Frm84.CB19 = 1 Then
    Frm84.TB41.Locked = False
    Frm84.TB41.BackColor = &HFFFFFF
    'Frm84.TB41 = vbNullString
    
    If Frm84.Visible = True Then Frm84.TB41.SetFocus
Else
    If Frm84.TB41 <> vbNullString Then
        Note = "Adakah anda ingin reset Approval Code ini ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
        
            Frm84.TB41.Locked = True
            Frm84.TB41.BackColor = &H8000000A
            Frm84.TB41 = vbNullString
            
        ElseIf Answer = vbNo Then
            
            Frm84.CB19 = 1
            
        End If
    Else
        Frm84.TB41.Locked = True
        Frm84.TB41.BackColor = &H8000000A
    End If
End If
End Sub
Private Sub CB2_Click()
'On Error Resume Next
If Frm84.CB2 = 1 Then
    Frm84.CB3 = 0
    Frm84.CB18 = 0
End If

Call frm84_kiraan_gst
Call Frm84_modal_dan_jual

Exit Sub

If Frm84.CB2 = 1 Then
    Frm84.CB3 = 0
    Frm84.CB18 = 0
    Frm84.TB11 = "0.00" 'Jumlah Cukai GST (RM)

    If Frm84.CB12 = 0 Then
        If IsNumeric(Frm84.TB10) Then
            Frm84.L44_Text = Format(Frm84.TB10, "#,##0.00")
        Else
            Frm84.L44_Text = Format(0, "#,##0.00")
        End If
    Else
        If IsNumeric(Frm84.TB15) Then
            Frm84.L44_Text = Format(Frm84.TB15, "#,##0.00")
        Else
            Frm84.L44_Text = Format(0, "#,##0.00")
        End If
    End If
End If

Call Frm84_modal_dan_jual
End Sub
Private Sub CB3_Click()
'On Error Resume Next
Dim frm84_LM_KADAR_GST As Double
Dim Frm84_LM_HARGA As Double

If Frm84.CB3 = 1 Then
    Frm84.CB2 = 0
    Frm84.CB18 = 0
End If

Call frm84_kiraan_gst
Call Frm84_modal_dan_jual

Exit Sub

If Frm84.CB3 = 1 Then
    Frm84.CB2 = 0
End If
If Frm84.CB3 = 0 Then
    Frm84.CB18 = 0
End If

Frm84_LM_HARGA = 0

If Frm84.CB3 = 1 And (Frm84.TB10 <> vbNullString And IsNumeric(Frm84.TB10)) And (Frm84.L8_Text <> vbNullString And IsNumeric(Frm84.L8_Text)) Then
    If Frm84.CB18 = 0 Then
        frm84_LM_KADAR_GST = Frm84.L8_Text 'Jumlah Kadar GST (%)
        If Frm84.CB12 = 0 Then
            Frm84_LM_HARGA = Frm84.TB10 'Harga (RM)
        Else
            If Frm84.TB15 <> vbNullString And IsNumeric(Frm84.TB15) Then
                Frm84_LM_HARGA = Frm84.TB15 'Upah (RM)
            End If
        End If
        
        Frm84.TB11 = Format((frm84_LM_KADAR_GST / 100) * Frm84_LM_HARGA, "#,##0.00") 'Jumlah Cukai GST (RM)
        Frm84.L44_Text = Format(Frm84_LM_HARGA, "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
    ElseIf Frm84.CB18 = 1 Then
        frm84_LM_KADAR_GST = Frm84.L8_Text 'Jumlah Kadar GST (%)
        If Frm84.CB12 = 0 Then
            Frm84_LM_HARGA = Frm84.TB10 'Harga (RM)
        Else
            If Frm84.TB15 <> vbNullString And IsNumeric(Frm84.TB15) Then
                Frm84_LM_HARGA = Frm84.TB15 'Upah (RM)
            End If
        End If

        Frm84.L44_Text = Format(Frm84_LM_HARGA / (1 + (frm84_LM_KADAR_GST / 100)), "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
        Frm84.TB11 = Format(Frm84_LM_HARGA - (Frm84_LM_HARGA / (1 + (frm84_LM_KADAR_GST / 100))), "#,##0.00") 'Jumlah Cukai GST (RM)
    End If
Else
    Frm84.TB11 = "0.00" 'Jumlah Cukai GST (RM)

    If Frm84.CB12 = 0 Then
        If IsNumeric(Frm84.TB10) Then
            Frm84.L44_Text = Format(Frm84.TB10, "#,##0.00")
        Else
            Frm84.L44_Text = Format(0, "#,##0.00")
        End If
    Else
        If IsNumeric(Frm84.TB15) Then
            Frm84.L44_Text = Format(Frm84.TB15, "#,##0.00")
        Else
            Frm84.L44_Text = Format(0, "#,##0.00")
        End If
    End If
End If

Call Frm84_modal_dan_jual
End Sub
Private Sub CB4_Click()
'on error resume next
If Frm84.CB4 = 1 Then
    Frm84.CB5 = 0
    Frm84.CB6 = 0
    Frm84.CB9 = 0
    Frm84.CB10 = 0
    
    Frm84.L45_Text = 1
    Frm84.L63_Text = "Kategori pelanggan           : PELANGGAN / PEMBELI BIASA"
End If
End Sub
Private Sub CB5_Click()
'on error resume next
If Frm84.CB5 = 1 Then
    Frm84.CB4 = 0
    Frm84.CB6 = 0
    Frm84.CB9 = 0
    Frm84.CB10 = 0
    
    Frm84.L45_Text = 2
    Frm84.L63_Text = "Kategori pelanggan           : AHLI BIASA"
End If
End Sub
Private Sub CB6_Click()
'on error resume next
If Frm84.CB6 = 1 Then
    Frm84.CB5 = 0
    Frm84.CB4 = 0
    Frm84.CB9 = 0
    Frm84.CB10 = 0
    
    Frm84.L45_Text = 4
    Frm84.L63_Text = "Kategori pelanggan           : SILVER"
End If
End Sub
Private Sub CB7_Click()
'on error resume next
If Frm84.CB7 = 1 Then
    Frm84.CMD11.Enabled = True
    Frm84.L62_Text = "Jualan oleh agen dropship : YA"
Else
    Frm84.CMD11.Enabled = False
    Frm84.L62_Text = "Jualan oleh agen dropship : TIDAK"
End If
End Sub
Private Sub CB9_Click()
'on error resume next
If Frm84.CB9 = 1 Then
    Frm84.CB5 = 0
    Frm84.CB6 = 0
    Frm84.CB4 = 0
    Frm84.CB10 = 0
    
    Frm84.L45_Text = 3
    Frm84.L63_Text = "Kategori pelanggan           : GOLD"
End If
End Sub
Private Sub CBB3_Click()
'on error resume next
Frm84.L12_Text = Frm84.CBB3
End Sub
Private Sub CBB4_Click()
'on error resume next
DATA_FOUND = 0
Frm84.TB3 = "0.00"

If GLOBAL_DISABLE = 0 Then
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from setting_database where Metal_Purity='" & Frm84.CBB4 & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        If Not IsNull(rs!Kod_Metal_Purity) Then
            Frm84.L13_Text = rs!Kod_Metal_Purity
            DATA_FOUND = 1
        End If

    End If
    
    rs.Close
    Set rs = Nothing
    
    If DATA_FOUND = 1 Then

        Dim LM_BERAT_ASAL As Double
        Dim LM_BERAT_GUNA As Double
        Dim LM_BERAT_TEMP As Double
        Dim LM_BERAT_TEMP_ASAL As Double
        
        LM_BERAT_ASAL = 0
        LM_BERAT_GUNA = 0
        LM_BERAT_TEMP = 0
        LM_BERAT_TEMP_ASAL = 0
        
        Frm84.TB3 = Format(0, "#,##0.00")
        
        Set rs3 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs3.Open "select SUM(beza_berat) from data_database where Purity='" & Frm84.CBB4 & "' AND (((statusitem = 10 OR statusitem = 12 OR statusitem = 20 OR statusitem = 22) AND receiving_Status = 2) OR ((statusitem = 12 OR statusitem = 20 OR statusitem = 22) AND receiving_Status = 0))", cn, adOpenKeyset, adLockOptimistic
        
        If Not IsNull(rs3(0)) Then LM_BERAT_ASAL = Format(rs3(0), "#,##0.00")
            
        rs3.Close
        Set rs3 = Nothing
        
        Set rs3 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs3.Open "select SUM(berat) from 85_penggunaan_ti where purity='" & Frm84.CBB4 & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
            
        If Not IsNull(rs3(0)) Then LM_BERAT_GUNA = Format(rs3(0), "#,##0.00")
            
        rs3.Close
        Set rs3 = Nothing
        
        Set rs3 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs3.Open "select SUM(Berat_Jualan) from " & G_JUALAN_TEMP & " where purity='" & Frm84.L13_Text & "' AND flag_barang = 1 AND (status = 1 OR Status = 2 OR Status = 3 OR Status = 4)", cn, adOpenKeyset, adLockOptimistic
            
        If Not IsNull(rs3(0)) Then LM_BERAT_TEMP = Format(rs3(0), "#,##0.00")
            
        rs3.Close
        Set rs3 = Nothing
        
        Frm84.TB3 = Format(LM_BERAT_ASAL - LM_BERAT_GUNA - LM_BERAT_TEMP, "#,##0.00")
        
        If Frm84.L3_Text <> vbNullString Then Call frm84_berat_guna_dr_invoice_ini
        
        Frm84.TB5 = Format(0, "0.00")
        
        If Frm84.L13_Text <> vbNullString Then
        
            Set rs3 = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs3.Open "select * from hargaemas where Purity='" & Frm84.L13_Text & "' AND cawangan='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs3.EOF Then
                If Not IsNull(rs3!HargaDariSupplier) Then
                    If IsNumeric(rs3!HargaDariSupplier) Then
                        Frm84.L54_Text = rs3!HargaDariSupplier
                    Else
                        Frm84.L54_Text = 0
                    End If
                Else
                    Frm84.L54_Text = 0
                End If
                If Not IsNull(rs3!harga_staff) Then
                    If IsNumeric(rs3!harga_staff) Then Frm84.L49_Text = Format(rs3!harga_staff, "0.00") 'Harga jualan kepada staff
                End If
                If Frm84.CB4 = 1 Then
                    If IsNumeric(rs3!Harga_Pelanggan) Then Frm84.TB5 = Format(rs3!Harga_Pelanggan, "0.00") 'Harga Emas Semasa Pelanggan
                ElseIf Frm84.CB5 = 1 Then
                    If IsNumeric(rs3!Harga_Member) Then Frm84.TB5 = Format(rs3!Harga_Member, "0.00") 'Harga Emas Semasa Member
                ElseIf Frm84.CB6 = 1 Then
                    If IsNumeric(rs3!Harga_Pengedar) Then Frm84.TB5 = Format(rs3!Harga_Pengedar, "0.00") 'Harga Emas Semasa Pengedar
                ElseIf Frm84.CB9 = 1 Then
                    If IsNumeric(rs3!Harga_RAF) Then Frm84.TB5 = Format(rs3!Harga_RAF, "0.00") 'Harga Emas Semasa RAF
                ElseIf Frm84.CB10 = 1 Then
                    If IsNumeric(rs3!harga_nd) Then Frm84.TB5 = Format(rs3!harga_nd, "0.00") 'Harga Emas Semasa Normal Dealer
                'ElseIf Frm84.CB11 = 1 Then
                '    If IsNumeric(rs!harga_md) Then Frm84.TB5 = Format(rs!harga_md, "0.00") 'Harga Emas Semasa Master Dealer
                End If
            End If
            
            rs3.Close
            Set rs3 = Nothing
            
        End If
        
    End If
    
End If
End Sub

Private Sub CMD1_Click()
'on error resume next
Dim Frm84_LM_LIMIT As Integer
Dim Frm84_LM_BIL As Integer

If Frm84.TB1 = vbNullString Then
    MsgBox "Sila Masukkan No. Siri Produk.", vbInformation, "Info"
    
    Frm84.TB1.SetFocus
    Exit Sub
End If

If InStr(1, Frm84.TB1, "'") <> 0 Then
    MsgBox "No. Siri Produk Mengandungi Simbol Yang Tidak Sah , ['].", vbInformation, "Info"
    
    Frm84.TB1 = vbNullString
    Frm84.TB1.SetFocus
    Exit Sub
End If

If Frm84.CB4 = 0 And Frm84.CB5 = 0 And Frm84.CB6 = 0 And Frm84.CB9 = 0 And Frm84.CB10 = 0 Then
    MsgBox "Sila pilih kategori pembeli.", vbExclamation, "info"
    Exit Sub
End If

If IsNumeric(Frm84.L46_Text) Then Frm84_LM_LIMIT = Frm84.L46_Text 'Limit Invoice
If IsNumeric(Frm84.L4_Text) Then Frm84_LM_BIL = Frm84.L4_Text 'Kuantiti Terkini

If Frm84_LM_LIMIT <> 0 Then
    If Frm84_LM_BIL >= Frm84_LM_LIMIT Then
        MsgBox "Hanya " & Frm84_LM_LIMIT & " item sahaja dibenarkan untuk dijual dalam satu invoice.", vbInformation, "Info"
    Else
        Call Frm84_Call_Product_Detail
    End If
Else
    Call Frm84_Call_Product_Detail
End If
End Sub

Private Sub CMD11_Click()
'On Error Resume Next
If Frm84.L29_Text = vbNullString Then Call Frm27_initial
Frm27.Show 1
End Sub
Private Sub CMD12_Click()
'On Error Resume Next
If Frm84.L28_Text = vbNullString Then
    
    If Frm84.L27_Text <> vbNullString Then
    
        Note = "Anda telah mengisi data pembeli barangan ini di dalam ruangan pelanggan yang TIDAK berdaftar dengan kedai." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Jika ada meneruskan menu ini , semua data pembeli di dalam ruangan pelanggan TIDAK berdaftar akan dipadamkan." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
            
            'Unload Frm26
            'Unload Frm27
            Call Frm28_initial
            
            Frm84.L27_Text = vbNullString 'Nama pembeli : Tidak berdaftar
            
            Frm28.Show 1
        End If
        
    Else
        
        'Unload Frm26
        'Unload Frm27
        Call Frm28_initial
        
        Frm28.Show 1
                
    End If
    
Else

    Frm28.Show 1
    
End If
End Sub
Private Sub CMD13_Click()
'On Error Resume Next
Dim Err(30)
Dim Frm84_LM_BERAT_ASAL As Double
Dim Frm84_LM_BERAT_JUAL As Double
Dim Frm84_LM_HARGA_MODAL As Double
Dim Frm84_LM_HARGA_JUAL As Double
Dim Frm84_LM_HARGA_SEMASA_MODAL As Double
Dim Frm84_LM_TETAPANHARGA As Double
Dim Frm84_LM_LIMIT As Double
Dim Frm84_LM_HARGA_STAFF As Double 'Tetapan harga jualan kepada staff
Dim Frm84_LM_HARGA_PELANGGAN As Double 'Tetapan harga jualan kepada pelanggan
Dim Frm84_LM_HARGA_SEMASA As Double 'Harga semasa (jualan)
Dim Frm84_LM_HARGA_SUPPLIER As Double 'Harga per gram (harga semasa) dari supplier (modal)
Dim Frm84_UPAH_MODAL As Double 'Upah modal
Dim Frm84_UPAH_JUAL As Double 'Upah jualan
Frm84_LM_HARGA_SEMASA = 0 'Harga semasa (jualan)
Frm84_LM_HARGA_SUPPLIER = 0 'Harga per gram (harga semasa) dari supplier (modal)
Dim Frm84_LM_HARGA_JUALAN_CALC As Double 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
Dim Frm84_LM_GST_CALC As Double 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
Dim Frm84_KOMISYEN_UPAH As Double 'Komisyen dari upah kepada agen dropship
Dim Frm84_LM_BERAT_OVERALL As Double
Dim Frm84_LM_SUSUT_BERAT As Double
Dim LM_HARGA_JUALAN_DGN_GST As Double
Dim LM_GST_JUAL As Double
Dim LM_MODAL_DGN_GST As Double
Dim LM_MODAL_TANPA_GST As Double
Dim LM_MODAL_TANPA_GST_GRAM As Double

LM_HARGA_JUALAN_DGN_GST = 0
LM_GST_JUAL = 0
LM_MODAL_DGN_GST = 0
LM_MODAL_TANPA_GST = 0
LM_MODAL_TANPA_GST_GRAM = 0
Frm84_LM_SUSUT_BERAT = 0
Frm84_LM_BERAT_OVERALL = 0
x = 0
Frm84_LM_BERAT_ASAL = 0
Frm84_LM_BERAT_JUAL = 0
Frm84_LM_DATA_SAVE = 0
Frm84_LM_HARGA_MODAL = 0
Frm84_LM_HARGA_JUAL = 0
Frm84_LM_HARGA_SEMASA_MODAL = 0
Frm84_LM_PRICE_CHECK = 0 '0 : Tidak Perlu Periksa Harga Semasa Jualan , 1 : Perlu Periksa Harga Semasa Jualan
Frm84_LM_TETAPANHARGA = 0
Frm84_LM_LIMIT = 0
Frm84_LM_HARGA_STAFF = 0
Frm84_LM_HARGA_PELANGGAN = 0
Frm84_UPAH_MODAL = 0 'Upah modal
Frm84_UPAH_JUAL = 0 'Upah jualan
Frm84_LM_HARGA_JUALAN_CALC = 0 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
Frm84_LM_GST_CALC = 0 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
Frm84_KOMISYEN_UPAH = 0 'Komisyen dari upah kepada agen dropship

If Frm84.L83_Text = "0" Then '0 : Stok kedai , 1 : Barang trade in/potong
    If Frm84.TB2 = vbNullString Then
        x = x + 1
        Err(x) = "Sila Masukkan [No. Siri Produk]."
    End If
End If
If Frm84.TB2 <> vbNullString And Frm84.TB3 = vbNullString And Frm84.CB12 = 1 Then
    MsgBox "Tetapan GST ke atas UPAH hanya dibenarkan untuk barang kemas SAHAJA. Sila periksa tetapan GST anda.", vbExclamation, "Info"
    Exit Sub
End If
'If (Frm84.TB14 <> vbNullString And IsNumeric(Frm84.TB14)) And (Frm84.L51_Text <> vbNullString And IsNumeric(Frm84.L51_Text)) Then
'    Frm84_LM_HARGA_STAFF = Frm84.L51_Text
'    Frm84_LM_HARGA_PELANGGAN = Frm84.TB14
    
'    If Frm84_LM_HARGA_PELANGGAN < Frm84_LM_HARGA_STAFF Then
'        X = X + 1
'        Err(X) = "Harga Jualan Minimum Yang Dibenarkan Adalah RM " & Format(Frm84_LM_HARGA_STAFF, "#,##0.00")
'    End If
'End If
    
'### Error Bagi Item BK ### - Start
If Frm84.TB3 <> vbNullString Then
    If Frm84.TB3 = vbNullString Or (Frm84.TB3 <> vbNullString And Not IsNumeric(Frm84.TB3)) Then
        x = x + 1
        Err(x) = "Sila Maklumat [Berat Asal]. Sila Scan Item Sekali Lagi."
    End If
    If Frm84.TB4 = vbNullString Or (Frm84.TB4 <> vbNullString And Not IsNumeric(Frm84.TB4)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Berat Jualan]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
    End If
    If Frm84.TB5 = vbNullString Or (Frm84.TB5 <> vbNullString And Not IsNumeric(Frm84.TB5)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Harga Semasa]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
    End If
    If Frm84.TB15 = vbNullString Or (Frm84.TB15 <> vbNullString And Not IsNumeric(Frm84.TB15)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Upah]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
    End If
    If Frm84.CB7 = 1 Then
        If Frm84.TB12 = vbNullString Or (Frm84.TB12 <> vbNullString And Not IsNumeric(Frm84.TB12)) Then
            x = x + 1
            Err(x) = "Sila Masukkan [Komisen Per Gram]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
        End If
        If Frm84.TB43 = vbNullString Or (Frm84.TB43 <> vbNullString And Not IsNumeric(Frm84.TB43)) Then
            x = x + 1
            Err(x) = "Sila Masukkan [Kadar Komisyen Upah (%)]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
        End If
        If Frm84.TB44 = vbNullString Or (Frm84.TB44 <> vbNullString And Not IsNumeric(Frm84.TB44)) Then
            x = x + 1
            Err(x) = "Sila Masukkan [Jumlah Komisyen Bagi Upah]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
        End If
        If (Frm84.TB15 <> vbNullString And IsNumeric(Frm84.TB15)) And (Frm84.TB44 <> vbNullString And IsNumeric(Frm84.TB44)) Then
            Frm84_UPAH_JUAL = Frm84.TB15 'Upah jualan
            Frm84_KOMISYEN_UPAH = Frm84.TB44 'Komisyen Upah
            
            If Frm84_KOMISYEN_UPAH > Frm84_UPAH_JUAL Then
                x = x + 1
                Err(x) = "Komisyen upah bagi agen dropship adalah melebihi dari upah asal."
            End If
        End If
    End If
End If
'### Error Bagi Item BK ### - End

'### Error Bagi Item Permata ### - Start
If Frm84.TB3 = vbNullString Then
    If Frm84.CB7 = 1 Then
        If Frm84.TB16 = vbNullString Or (Frm84.TB16 <> vbNullString And Not IsNumeric(Frm84.TB16)) Then
            x = x + 1
            Err(x) = "Sila Masukkan [Jumlah Komisen]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
        End If
    End If
End If
'### Error Bagi Item Permata ### - End

If Frm84.TB7 = vbNullString Or (Frm84.TB7 <> vbNullString And Not IsNumeric(Frm84.TB7)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Diskaun]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
End If
If Frm84.TB9 = vbNullString Or (Frm84.TB9 <> vbNullString And Not IsNumeric(Frm84.TB9)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Adjustment]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
End If
If Frm84.TB10 = vbNullString Or (Frm84.TB10 <> vbNullString And Not IsNumeric(Frm84.TB10)) Then
    x = x + 1
    Err(x) = "Tiada maklumat [Harga Jualan]. Sila keluarkan item ini dari senarai dan scan sekali lagi."
End If
If Frm84.TB11 = vbNullString Or (Frm84.TB11 <> vbNullString And Not IsNumeric(Frm84.TB11)) Then
    x = x + 1
    Err(x) = "Tiada maklumat [Jumlah GST]. Sila keluarkan item ini dari senarai dan scan sekali lagi."
End If
If Frm84.CB4 = 0 And Frm84.CB5 = 0 And Frm84.CB6 = 0 And Frm84.CB9 = 0 And Frm84.CB10 = 0 Then
    x = x + 1
    Err(x) = "Sila Buat Pilihan Kategori Pembeli."
End If
If Frm84.CB2 = 0 And Frm84.CB3 = 0 And Frm84.CB18 = 0 Then
    x = x + 1
    Err(x) = "Sila Buat Pilihan Jenis GST."
End If
If (Frm84.TB3 <> vbNullString And IsNumeric(Frm84.TB3)) And (Frm84.TB4 <> vbNullString And IsNumeric(Frm84.TB4)) Then
    Frm84_LM_BERAT_ASAL = Frm84.TB3 'Berat Asal
    Frm84_LM_BERAT_JUAL = Frm84.TB4 'Berat Jualan
    
    If Format(Frm84_LM_BERAT_JUAL, "0.00") = "0.00" Then
        x = x + 1
        Err(x) = "Berat jualan yang tidak sah Nilai 0 tidak dibenarkan di dalam ruangan ini."
    End If
    If Frm84_LM_BERAT_JUAL > Frm84_LM_BERAT_ASAL Then
        x = x + 1
        Err(x) = "Berat Jualan Melebihi Berat Asal."
    End If
End If
If Frm84.TB3 <> vbNullString And Frm84.L54_Text = vbNullString Then
    x = x + 1
    Err(x) = "Tiada maklumat harga dari supplier bagi purity ini." & vbCrLf & _
                "Sila pastikan harga dari supplier bagi purity ini telah ditetapkan dalam TETAPAN HARIAN SISTEM."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah Anda Ingin Masukkan Data Ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
'### Periksa Data Dulang ### - Start
        If Frm84.L83_Text = "0" Then '0 : Stok kedai , 1 : Barang trade in/potong
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from Data_Database where no_siri_Produk='" & Frm84.TB2 & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!dulang) Then Frm84_LM_DULANG = rs!dulang 'Dulang
                If Not IsNull(rs!susut_berat) Then Frm84_LM_SUSUT_BERAT = rs!susut_berat 'Susut berat
            End If
            
            rs.Close
            Set rs = Nothing
            
        End If
'### Periksa Data Dulang ### - End

        If Frm84.TB3 <> vbNullString And Frm84.TB4 <> vbNullString Then
            If IsNumeric(Frm84.TB3) Then Frm84_LM_BERAT_ASAL = Frm84.TB3 'Berat Asal
            If IsNumeric(Frm84.TB4) Then Frm84_LM_BERAT_JUAL = Frm84.TB4 'Berat Jualan
            
            Frm84_LM_BERAT_OVERALL = Frm84_LM_SUSUT_BERAT + Frm84_LM_BERAT_JUAL
            
            If Frm84_LM_BERAT_ASAL < Frm84_LM_BERAT_OVERALL Then
            
                MsgBox "Berat jualan melebihi berat jualan yang dibenarkan." & vbCrLf & _
                        "Berat asal : " & Format(Frm84_LM_BERAT_ASAL, "#,##0.00 g") & vbCrLf & _
                        "Susut berat : " & Format(Frm84_LM_SUSUT_BERAT, "#,##0.00 g") & vbCrLf & _
                        "Berat jualan maksimum yang dibenarkan adalah " & Format(Frm84_LM_BERAT_ASAL - Frm84_LM_SUSUT_BERAT, "#,##0.00 g"), vbInformation, "Info"
                        
                Exit Sub
                        
            End If
        End If
    
'### Periksa Kadar Penurunan Harga ### - Start
'GoTo skip_periksa_harga:
        user = MDI_frm1.L3_Text
        
        If MDI_frm1.L4_Text <> vbNullString Then
            If MDI_frm1.L4_Text = "Staff" Then
                Frm84_LM_PRICE_CHECK = 1 '0 : Tidak Perlu Periksa Harga Semasa Jualan , 1 : Perlu Periksa Harga Semasa Jualan
            End If
        End If
'skip_periksa_harga:

        If Frm84.CB13 = 0 And Frm84.CB2 = 1 And Frm84.L84_Text = "1" And Frm84.L85_Text = "0" Then
            
            Note = "Anda cuba menjual barang ini tanpa cukai GST." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sistem akan menukarkan jenis invoice jualan ini kepada TIDAK RASMI." & vbCrLf & _
                    "*** Invoice TIDAK RASMI adalah invoice yang tidak akan dikira sebagai jualan rasmi kedai." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila pilih [Yes] untuk meneruskan jualan ini dengan invoice tidak rasmi dan pilih [No] jika ingin meneruskan jualan dengan invoice rasmi."
                    
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbNo Then
                Exit Sub
            End If
            If Answer = vbYes Then

                Frm84.CB13 = 1
                
            End If
            
        End If
        
        If Frm84_LM_PRICE_CHECK = 1 Then '0 : Tidak Perlu Periksa Harga Semasa Jualan , 1 : Perlu Periksa Harga Semasa Jualan
            Frm84_LM_LIMIT_TYPE = 0 '1 : BK , 2 : Barang Permata
            
'### Periksa Purity Dan Tetapan Harga Jualan ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from Data_Database where no_siri_Produk='" & Frm84.TB2 & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!receiving_Status) Then
                    If rs!receiving_Status = 0 Or rs!receiving_Status = 2 Then '0 : BK , 1 : Barang Permata , 2 : Buyback BK , 3 : Buyback Barang Permata
                        If Not IsNull(rs!kod_Purity) Then
                            Frm84_LM_PURITY = rs!kod_Purity 'Purity
                        End If
                        Frm84_LM_LIMIT_TYPE = 1 '1 : BK , 2 : Barang Permata
                    End If
                    If rs!receiving_Status = 1 Or rs!receiving_Status = 3 Then '0 : BK , 1 : Barang Permata , 2 : Buyback BK , 3 : Buyback Barang Permata
                        If Frm84.CB4 = 1 Then
                            If IsNumeric(rs!code_Supplier) Then Frm84_LM_TETAPANHARGA = Format(rs!code_Supplier, "0.00")  'Harga Pelanggan
                        ElseIf Frm84.CB5 = 1 Then
                            If IsNumeric(rs!HargaJualan_Member) Then Frm84_LM_TETAPANHARGA = Format(rs!HargaJualan_Member, "0.00") 'Harga Member
                        ElseIf Frm84.CB9 = 1 Then
                            If IsNumeric(rs!HargaJualan_RAF) Then Frm84_LM_TETAPANHARGA = Format(rs!HargaJualan_RAF, "0.00") 'Harga RAF
                        ElseIf Frm84.CB6 = 1 Then
                            If IsNumeric(rs!HargaJualan_Pengedar) Then Frm84_LM_TETAPANHARGA = Format(rs!HargaJualan_Pengedar, "0.00") 'Harga Pengedar
                        ElseIf Frm84.CB10 = 1 Then
                            If IsNumeric(rs!hargajualan_normal_dealer) Then Frm84_LM_TETAPANHARGA = Format(rs!hargajualan_normal_dealer, "0.00") 'Harga Normal Dealer
                        'ElseIf Frm84.CB11 = 1 Then
                        '    If IsNumeric(rs!hargajualan_master_dealer) Then Frm84_LM_TETAPANHARGA = Format(rs!hargajualan_master_dealer, "0.00") 'Harga Master Dealer
                        End If
                        Frm84_LM_LIMIT_TYPE = 2 '1 : BK , 2 : Barang Permata
                    End If
                End If
            End If
            
            rs.Close
            Set rs = Nothing
            
'### Carian Harga Semasa Emas ### - Start
            If Frm84_LM_LIMIT_TYPE = 1 Then '1 : BK , 2 : Barang Permata
            
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from default_setting where Default1='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    'If rs!Default1 = "Default" Then
                        If IsNumeric(rs!limit_per_gram) Then Frm84_LM_LIMIT = rs!limit_per_gram
                    'End If
                End If
                
                rs.Close
                Set rs = Nothing
            
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from hargaemas where Purity='" & Frm84_LM_PURITY & "' AND cawangan='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Frm84.CB4 = 1 Then
                        If IsNumeric(rs!Harga_Pelanggan) Then Frm84_LM_TETAPANHARGA = Format(rs!Harga_Pelanggan, "0.00") 'Harga Pelanggan
                    ElseIf Frm84.CB5 = 1 Then
                        If IsNumeric(rs!Harga_Member) Then Frm84_LM_TETAPANHARGA = Format(rs!Harga_Member, "0.00") 'Harga Member
                    ElseIf Frm84.CB9 = 1 Then
                        If IsNumeric(rs!Harga_RAF) Then Frm84_LM_TETAPANHARGA = Format(rs!Harga_RAF, "0.00") 'Harga RAF
                    ElseIf Frm84.CB6 = 1 Then
                        If IsNumeric(rs!Harga_Pengedar) Then Frm84_LM_TETAPANHARGA = Format(rs!Harga_Pengedar, "0.00") 'Harga Pengedar
                    ElseIf Frm84.CB10 = 1 Then
                        If IsNumeric(rs!harga_normal_dealer) Then Frm84_LM_TETAPANHARGA = Format(rs!harga_normal_dealer, "0.00") 'Harga Normal Dealer
                    'ElseIf Frm84.CB11 = 1 Then
                    '    If IsNumeric(rs!harga_master_dealer) Then Frm84_LM_TETAPANHARGA = Format(rs!harga_master_dealer, "0.00") 'Harga Master Dealer
                    End If
                End If
                
                rs.Close
                Set rs = Nothing
                
                If IsNumeric(Frm84.TB5) Then
                    Frm84_LM_HARGA_JUALAN = Frm84.TB5 'Harga Semasa Jualan (RM/g)
                End If
                
                If Frm84_LM_TETAPANHARGA - Frm84_LM_HARGA_JUALAN > Frm84_LM_LIMIT Then
                    MsgBox "Harga jualan tidak mengikut pengurangan harga minimum yang ditetapkan oleh kedai!." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Tetapan Asal Harga Jualan : RM " & Format(Frm84_LM_TETAPANHARGA, "0.00") & vbCrLf & _
                    "Limit Diskaun Pengurangan Harga : RM " & Format(Frm84_LM_LIMIT, "0.00") & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila periksa data anda.", vbExclamation, "Info"
                    Exit Sub
                End If
            End If
            
            If Frm84_LM_LIMIT_TYPE = 2 Then '1 : BK , 2 : Barang Permata
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from default_setting where Default1='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    'If rs!Default1 = "Default" Then
                        If IsNumeric(rs!limit_per_item) Then Frm84_LM_LIMIT = rs!limit_per_item
                    'End If
                End If
                
                rs.Close
                Set rs = Nothing
                
                If IsNumeric(Frm84.TB10) Then
                    Frm84_LM_HARGA_JUALAN = Frm84.TB10 'Harga Jualan (RM)
                End If
                
                If Frm84_LM_TETAPANHARGA - Frm84_LM_HARGA_JUALAN > Frm84_LM_LIMIT Then
                    MsgBox "Harga jualan tidak mengikut pengurangan harga minimum yang ditetapkan oleh kedai!." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Tetapan Asal Harga Jualan : RM " & Format(Frm84_LM_TETAPANHARGA, "0.00") & vbCrLf & _
                    "Limit Diskaun Pengurangan Harga : RM " & Format(Frm84_LM_LIMIT, "0.00") & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila periksa data anda.", vbExclamation, "Info"
                    Exit Sub
                End If
            End If
'### Carian Harga Semasa Emas ### - End
        
'### Periksa Purity Dan Tetapan Harga Jualan ### - End
        End If
'### Periksa Kadar Penurunan Harga ### - End
    
'### Masukkan Data Ke Dalam Temp Table ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_JUALAN_TEMP & " where ID='" & Frm84.L39_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Frm84.L83_Text = "0" Then '0 : Stok kedai , 1 : Barang trade in/potong
                If Frm84.TB2 <> vbNullString Then
                    rs!no_siri_Produk = Frm84.TB2 'No. Siri Produk
                Else
                    rs!no_siri_Produk = Null 'No. Siri Produk
                End If
                rs!nama_purity = Null
                rs!dulang = Frm84_LM_DULANG 'Dulang
                
                If Frm84.TB3 <> vbNullString Then
                    rs!Berat_Asal = Format(Frm84.TB3, "0.00") 'Berat Asal (g)
                Else
                    rs!Berat_Asal = Null 'Berat Asal (g)
                End If
                
            Else
                rs!no_siri_Produk = "-" 'No. Siri Produk
                If Frm84.CBB4 <> vbNullString Then
                    rs!nama_purity = Frm84.CBB4
                Else
                    rs!nama_purity = Null
                End If
                rs!dulang = "-" 'Dulang
                
                If Frm84.TB4 <> vbNullString Then
                    rs!Berat_Asal = Format(Frm84.TB4, "0.00") 'Berat Asal (g)
                Else
                    rs!Berat_Asal = Null 'Berat Asal (g)
                End If
                
            End If
            If Frm84.L83_Text = "0" Then '0 : Stok kedai , 1 : Barang trade in/potong
                rs!flag_barang = 0
            ElseIf Frm84.L83_Text = "1" Then '0 : Stok kedai , 1 : Barang trade in/potong
                rs!flag_barang = 1
            End If
            If Frm84.L12_Text <> vbNullString Then
                rs!kategori_Produk = Frm84.L12_Text 'Kategori Produk
            Else
                rs!kategori_Produk = Null 'Kategori Produk
            End If
            If Frm84.L13_Text <> vbNullString Then
                rs!purity = Frm84.L13_Text 'Purity
            Else
                rs!purity = Null 'Purity
            End If
            'If Frm84.TB3 <> vbNullString Then
            '    rs!Berat_Asal = Format(Frm84.TB3, "0.00") 'Berat Asal (g)
            'Else
            '    rs!Berat_Asal = Null 'Berat Asal (g)
            'End If
            If Frm84.TB4 <> vbNullString Then
                rs!berat_jualan = Format(Frm84.TB4, "0.00") 'Berat Jualan (g)
            Else
                rs!berat_jualan = Null 'Berat Jualan (g)
            End If
            If Frm84.TB5 <> vbNullString Then
                rs!harga_Semasa = Format(Frm84.TB5, "0.00") 'Harga Semasa (RM/g)
            Else
                rs!harga_Semasa = Null 'Harga Semasa (RM/g)
            End If
            If Frm84.TB15 <> vbNullString Then
                rs!UPAH = Format(Frm84.TB15, "0.00") 'Upah (RM)
            Else
                rs!UPAH = Null 'Upah (RM)
            End If
            If Frm84.TB6 <> vbNullString Then
                rs!harga_asal = Format(Frm84.TB6, "0.00") 'Harga Asal Item (RM)
            Else
                rs!harga_asal = Null 'Harga Asal Item (RM)
            End If
            If Frm84.TB7 <> vbNullString Then
                rs!diskaun = Format(Frm84.TB7, "0.00") 'Diskaun (%)
            Else
                rs!diskaun = Null 'Diskaun (%)
            End If
            If Frm84.TB8 <> vbNullString Then
                rs!harga_lepas_diskaun = Format(Frm84.TB8, "0.00") 'Harga Selepas Diskaun (RM)
            Else
                rs!harga_lepas_diskaun = Null 'Harga Selepas Diskaun (RM)
            End If
            If Frm84.TB9 <> vbNullString Then
                rs!adjustment = Format(Frm84.TB9, "0.00") 'Adjustment (RM)
            Else
                rs!adjustment = Null 'Adjustment (RM)
            End If
            If Frm84.TB10 <> vbNullString Then
                rs!harga_jualan = Format(Frm84.TB10, "0.00") 'Harga Jualan (RM)
            Else
                rs!harga_jualan = Null 'Harga Jualan (RM)
            End If
            If Frm84.CB2 = 1 Then
            
                rs!gst_ari_nashi = "ZR (L)" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                rs!kadar_gst = Null 'Kadar Cukai GST (%)
                If Frm84.TB11 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm84.TB11, "0.00") 'Jumlah Cukai GST (RM)
                    LM_GST_JUAL = Format(Frm84.TB11, "0.00")
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah Cukai GST (RM)
                End If
                
                rs!gst_include = Null '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                'If Frm84.L85_Text = "0" Then
                '    If Frm84.L84_Text = "1" Then Frm84.CB13 = 1
                'End If
                
            ElseIf Frm84.CB3 = 1 Then
            
                rs!gst_ari_nashi = "SR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                If Frm84.L8_Text <> vbNullString Then
                    rs!kadar_gst = Frm84.L8_Text 'Kadar Cukai GST (%)
                Else
                    rs!kadar_gst = "0" 'Jumlah Cukai GST (RM)
                End If
                If Frm84.TB11 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm84.TB11, "0.00") 'Jumlah Cukai GST (RM)
                    LM_GST_JUAL = Format(Frm84.TB11, "0.00")
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah Cukai GST (RM)
                End If

                rs!gst_include = 0 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang

            ElseIf Frm84.CB18 = 1 Then
            
                rs!gst_ari_nashi = "SR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                If Frm84.L8_Text <> vbNullString Then
                    rs!kadar_gst = Frm84.L8_Text 'Kadar Cukai GST (%)
                Else
                    rs!kadar_gst = "0" 'Jumlah Cukai GST (RM)
                End If
                If Frm84.TB11 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm84.TB11, "0.00") 'Jumlah Cukai GST (RM)
                    LM_GST_JUAL = Format(Frm84.TB11, "0.00")
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah Cukai GST (RM)
                End If
                
                rs!gst_include = "1" '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang

            End If
            If Frm84.L44_Text <> vbNullString Then
                rs!harga_tanpa_gst = Format(Frm84.L44_Text, "0.00") 'Harga Jualan Tanpa GST (RM)
            Else
                rs!harga_tanpa_gst = Null 'Harga Jualan Tanpa GST (RM)
            End If
            If Frm84.TB14 <> vbNullString Then
                rs!harga_dengan_gst = Format(Frm84.TB14, "0.00") 'Harga Jualan Termasuk GST (RM)
            Else
                rs!harga_dengan_gst = Null 'Harga Jualan Termasuk GST (RM)
            End If
            If Frm84.CB7 = 1 Then
                rs!dropship = 1 '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
                If Frm84.Frame2.Visible = True Then 'Komisen Agen Dropship : BK
                    If Frm84.TB12 <> vbNullString Then
                        rs!komisyen_per_gram = Format(Frm84.TB12, "0.00") 'Komisyen Per Gram Dropship (RM/g) : BK
                    Else
                        rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g) : BK
                    End If
                    If Frm84.TB13 <> vbNullString Then
                        rs!jumlah_komisyen = Format(Frm84.TB13, "0.00") 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
                    Else
                        rs!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
                    End If
                    If Frm84.TB43 <> vbNullString Then
                        rs!kadar_komisyen_upah = Frm84.TB43 'Kadar komisyen bagi upah kepada agen dropship
                    Else
                        rs!kadar_komisyen_upah = Null 'Kadar komisyen bagi upah kepada agen dropship
                    End If
                    If Frm84.TB44 <> vbNullString Then
                        rs!komisyen_upah = Format(Frm84.TB44, "0.00") 'Jumlah komisyen bagi upah kepada agen dropship
                    Else
                        rs!komisyen_upah = Null 'Jumlah komisyen bagi upah kepada agen dropship
                    End If
                End If
                If Frm84.Frame3.Visible = True Then 'Komisen Agen Dropship : Permata
                    rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g) : BK
                    rs!kadar_komisyen_upah = Null 'Kadar komisyen bagi upah kepada agen dropship
                    rs!komisyen_upah = Null 'Jumlah komisyen bagi upah kepada agen dropship
                    If Frm84.TB16 <> vbNullString Then
                        rs!jumlah_komisyen = Format(Frm84.TB16, "0.00") 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : Permata
                    Else
                        rs!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : Permata
                    End If
                End If
            End If
                
            If Frm84.CB7 = 0 Then
                rs!dropship = 0 '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
                rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g)
                rs!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini
                rs!kadar_komisyen_upah = Null 'Kadar komisyen bagi upah kepada agen dropship
                rs!komisyen_upah = Null 'Jumlah komisyen bagi upah kepada agen dropship
            End If
            
'Status
'0 : Keluarkan Dari Senarai
'1 : Data Baru (Fresh)
'2 : Data Baru Diedit (Fresh)
'3 : Data Baru Dari Menu Edit
'4 : Data Baru Dari Menu Edit Yang Telah Diedit

            If Frm84.L41_Text = "0" Then '0 : Menu Data Baru , 1 : Menu Edit Data
                rs!Status = 1
            ElseIf Frm84.L41_Text = "1" Then
                If rs!Status = "1" Then
                    rs!Status = 4
                End If
                If rs!Status = "3" Then
                    rs!Status = 3
                End If
            End If
            
            If Frm84.TB3 = vbNullString Then
            
                rs!Type = 1 '0 : BK , 1 : Barang Permata
                rs!harga_per_gram_modal = Null 'Harga Per Gram Bagi Modal (RM/g)
                If Frm84.L34_Text <> vbNullString Then
                    rs!modal = Format(Frm84.L34_Text, "0.00") 'Harga Modal (RM)
                    LM_MODAL_DGN_GST = Format(Frm84.L34_Text, "0.00")
                Else
                    rs!modal = Null 'Harga Modal (RM)
                End If
                If Frm84.L42_Text <> vbNullString Then
                    rs!modal_tanpa_gst = Format(Frm84.L42_Text, "0.00") 'Harga Modal Tanpa GST (RM)
                    LM_MODAL_TANPA_GST = Frm84.L42_Text
                Else
                    rs!modal_tanpa_gst = Null 'Harga Modal (RM)
                End If
                If Frm84.L44_Text <> vbNullString Then
                    If IsNumeric(Frm84.L44_Text) Then Frm84_LM_HARGA_JUALAN_TANPA_GST = Frm84.L44_Text
                End If
                
                If Frm84.CB3 = 1 And Frm84.CB18 = 0 Then
                    Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
                    Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
                    
                    Frm84_LM_HARGA_JUALAN_DENGAN_GST = Format(Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC, "0.00")
                    'Field ini adalah lebih kurang kepada @harga_dengan_gst
                    'Tetapi bagi gst yang dikenakan hanya pada UPAH akan membuatkan field ini hanya memaparkan harga dengan gst hanya pada upah.
                    'Oleh itu field ini dibuat khas bagi mencampurkan gst pada upah dan harga barang.
                Else
                    Frm84_LM_HARGA_JUALAN_DENGAN_GST = Format(Frm84.TB10, "0.00") 'Harga Jualan (RM)
                End If
                
                If IsNumeric(Frm84.L34_Text) Then Frm84_LM_MODAL_DENGAN_GST = Frm84.L34_Text 'Harga Modal
                
                rs!jualan_per_gram_dengan_gst = Null
                rs!untung = Format(Frm84_LM_HARGA_JUALAN_DENGAN_GST - Frm84_LM_MODAL_DENGAN_GST, "0.00") 'Jumlah Keuntungan (pengiraan termasuk GST)
                rs!untung2 = Format(Frm84_LM_HARGA_JUALAN_TANPA_GST - Frm84_LM_MODAL_TANPA_GST, "0.00") 'Jumlah Keuntungan (pengiraan TIDAK termasuk GST)
                
                rs!harga_per_gram_supplier = Null 'Harga per gram (harga semasa) dari supplier (modal)
                rs!upah_modal = Null 'Upah modal
                rs!harga_per_gram_tanpa_gst = Null 'Harga modal per gram tanpa GST (RM)
                
            Else
            
                rs!Type = 0 '0 : BK , 1 : Barang Permata
                
                If Frm84.L34_Text <> vbNullString Then
                    rs!harga_per_gram_modal = Format(Frm84.L34_Text, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
                    If IsNumeric(Frm84.L34_Text) Then
                        Frm84_LM_HARGA_SEMASA_MODAL = Frm84.L34_Text 'Harga Per Gram Bagi Modal (RM/g)
                        
                        rs!modal = Format(Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_SEMASA_MODAL, "0.00") 'Harga Modal (RM)
                        LM_MODAL_DGN_GST = Format(Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_SEMASA_MODAL, "0.00")
                    End If
                Else
                    rs!modal = Null 'Harga Modal (RM)
                    rs!harga_per_gram_modal = Null 'Harga Per Gram Bagi Modal (RM/g)
                End If
                If Frm84.L42_Text <> vbNullString Then
                    rs!harga_per_gram_tanpa_gst = Format(Frm84.L42_Text, "0.00") 'Harga modal per gram tanpa GST (RM)
                    LM_MODAL_TANPA_GST_GRAM = Frm84.L42_Text
                    LM_MODAL_TANPA_GST = Format(Frm84_LM_BERAT_JUAL * LM_MODAL_TANPA_GST_GRAM, "0.00")
                Else
                    rs!harga_per_gram_tanpa_gst = Null 'Harga modal per gram tanpa GST (RM)
                End If
                
                If IsNumeric(Frm84.TB14) And IsNumeric(Frm84.L34_Text) And IsNumeric(Frm84.TB10) And IsNumeric(Frm84.TB11) Then
                    
                    Frm84_LM_HARGA_MODAL = Frm84.L34_Text 'Harga Modal
                    Frm84_LM_HARGA_JUAL = Frm84.TB14 'Harga Jualan
                    Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
                    Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
                    
                    If Frm84.L34_Text <> vbNullString Then
                        If IsNumeric(Frm84.L34_Text) Then Frm84_LM_MODAL_DENGAN_GST = Frm84.L34_Text 'Harga Modal
                    End If
                    
                    If Frm84.L42_Text <> vbNullString Then
                        If IsNumeric(Frm84.L42_Text) Then Frm84_LM_MODAL_TANPA_GST = Frm84.L42_Text
                    End If
                        
                    If Frm84.CB12 = 0 Then
                        
                        If Frm84.TB14 <> vbNullString Then
                            If IsNumeric(Frm84.TB14) Then Frm84_LM_HARGA_JUALAN_DENGAN_GST = Frm84.TB14 'Harga Jualan
                        End If
                        
                        If Frm84.L44_Text <> vbNullString Then
                            If IsNumeric(Frm84.L44_Text) Then Frm84_LM_HARGA_JUALAN_TANPA_GST = Frm84.L44_Text
                        End If
                        
                        rs!jualan_per_gram_dengan_gst = Format(Frm84_LM_HARGA_JUALAN_DENGAN_GST / Frm84_LM_BERAT_JUAL, "0.00")
                        rs!untung = Format((Frm84_LM_HARGA_JUALAN_DENGAN_GST) - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_DENGAN_GST), "0.00") 'Jumlah Keuntungan (pengiraan termasuk GST)
                        rs!untung2 = Format(Frm84_LM_HARGA_JUALAN_TANPA_GST - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_TANPA_GST), "0.00") 'Jumlah Keuntungan (pengiraan TIDAK termasuk GST)

                    ElseIf Frm84.CB12 = 1 Then
                        
                        If Frm84.CB2 = 1 Then
                        
                            If Frm84.TB10 <> vbNullString Then
                                If IsNumeric(Frm84.TB10) Then Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
                            End If
                             
                            If Frm84.TB11 <> vbNullString Then
                                Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
                            End If
                            
                            rs!jualan_per_gram_dengan_gst = Format((Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC) / Frm84_LM_BERAT_JUAL, "0.00")
                            rs!untung = Format((Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC) - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_DENGAN_GST), "0.00") 'Jumlah Keuntungan (pengiraan termasuk GST)
                            rs!untung2 = Format(Frm84_LM_HARGA_JUALAN_CALC - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_TANPA_GST), "0.00") 'Jumlah Keuntungan (pengiraan TIDAK termasuk GST)
                            
                        ElseIf Frm84.CB3 = 1 Then
                            
                            If Frm84.TB10 <> vbNullString Then
                                If IsNumeric(Frm84.TB10) Then Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
                            End If
                             
                            If Frm84.TB11 <> vbNullString Then
                                Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
                            End If
                            
                            rs!jualan_per_gram_dengan_gst = Format((Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC) / Frm84_LM_BERAT_JUAL, "0.00")
                            rs!untung = Format((Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC) - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_DENGAN_GST), "0.00") 'Jumlah Keuntungan (pengiraan termasuk GST)
                            rs!untung2 = Format(Frm84_LM_HARGA_JUALAN_CALC - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_TANPA_GST), "0.00") 'Jumlah Keuntungan (pengiraan TIDAK termasuk GST)

                        ElseIf Frm84.CB18 = 1 Then
                        
                            If Frm84.TB10 <> vbNullString Then
                                If IsNumeric(Frm84.TB10) Then Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
                            End If
                             
                            If Frm84.TB11 <> vbNullString Then
                                Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
                            End If
                            
                            rs!jualan_per_gram_dengan_gst = Format((Frm84_LM_HARGA_JUALAN_CALC) / Frm84_LM_BERAT_JUAL, "0.00")
                            rs!untung = Format((Frm84_LM_HARGA_JUALAN_CALC) - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_DENGAN_GST), "0.00")  'Jumlah Keuntungan (pengiraan termasuk GST)
                            rs!untung2 = Format((Frm84_LM_HARGA_JUALAN_CALC - Frm84_LM_GST_CALC) - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_TANPA_GST), "0.00") 'Jumlah Keuntungan (pengiraan TIDAK termasuk GST)

                        End If
                        
                    End If
                    
                End If
                
            End If
            
            If Format(Frm84.TB3, "0.00") = Format(Frm84.TB4, "0.00") Then
                rs!potong_flag = 0 '0 : Tiada Potong , 1 : Ada Potong
                rs!status_jualan = 0 '0 : Tiada Potong , 1 : Ada Potong
            Else
                rs!potong_flag = 1 '0 : Tiada Potong , 1 : Ada Potong
                rs!status_jualan = 1 '0 : Tiada Potong , 1 : Ada Potong
            End If
            'rs!dulang = Frm84_LM_DULANG 'Dulang
            
            '### Maklumat tetapan harga jualan kepada staff ### - Start
            If Frm84.L48_Text <> vbNullString Then
                rs!kadar_penurunan_upah = Format(Frm84.L48_Text, "0.00") 'Kadar peratusan penurunan harga upah kepada staff (%)
            Else
                rs!kadar_penurunan_upah = Null
            End If
            If Frm84.L49_Text <> vbNullString Then
                rs!harga_semasa_staff = Format(Frm84.L49_Text, "0.00") 'Harga emas semasa yang dijual kepada staff
            Else
                rs!harga_semasa_staff = Null
            End If
            If Frm84.L50_Text <> vbNullString Then
                rs!kadar_penurunan_bp = Format(Frm84.L50_Text, "0.00") 'Kadar peratusan penurunan harga barang permata kepada staff (%)
            Else
                rs!kadar_penurunan_bp = Null
            End If
            If Frm84.L51_Text <> vbNullString Then
                rs!harga_staff = Format(Frm84.L51_Text, "0.00") 'Harga yang dijual kepada staff (RM)
            Else
                rs!harga_staff = Null
            End If
            If Frm84.L52_Text <> vbNullString Then
                rs!harga_bp_asal = Format(Frm84.L52_Text, "0.00") 'Tetapan harga barang permata yang asal (RM)
            Else
                rs!harga_bp_asal = Null
            End If
            If Frm84.L53_Text <> vbNullString Then
                rs!upah_asal = Format(Frm84.L53_Text, "0.00") 'Tetapan upah asal (RM)
            Else
                rs!upah_asal = Null
            End If
            rs!komisyen_staff = Format(Frm84_LM_HARGA_PELANGGAN - Frm84_LM_HARGA_STAFF, "0.00") 'Jumlah Komisyen Staff (RM)
            '### Maklumat tetapan harga jualan kepada staff ### - End
            
            If Frm84.CB12 = 0 Then '0 : GST pada harga jualan , 1 : GST pada upah
                rs!gst_barang_atau_upah = 0
            Else
                rs!gst_barang_atau_upah = 1
            End If
            If Frm84.CB3 = 1 And Frm84.CB18 = 0 Then
                Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
                Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
                
                rs!harga_jualan_dengan_gst = Format(Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC, "0.00")
                LM_HARGA_JUALAN_DGN_GST = Format(Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC, "0.00")
                'Field ini adalah lebih kurang kepada @harga_dengan_gst
                'Tetapi bagi gst yang dikenakan hanya pada UPAH akan membuatkan field ini hanya memaparkan harga dengan gst hanya pada upah.
                'Oleh itu field ini dibuat khas bagi mencampurkan gst pada upah dam harga barang.
            Else
                rs!harga_jualan_dengan_gst = Format(Frm84.TB10, "0.00") 'Harga Jualan (RM)
                LM_HARGA_JUALAN_DGN_GST = Format(Frm84.TB10, "0.00")
            End If
            If Frm84.L67_Text <> vbNullString Then 'Purata harga jualan per gram (RM/g) bagi barang kemas , Bagi barang permata adalah merujuk kepada harga jualan
                rs!jualan_per_gram = Format(Frm84.L67_Text, "0.00")
            Else
                rs!jualan_per_gram = Null
            End If
            If Frm84.L69_Text <> vbNullString Then 'Paparan modal per gram (tanpa GST)
                rs!modal_per_gram = Format(Frm84.L69_Text, "0.00")
            Else
                rs!modal_per_gram = Null
            End If
            rs!harga_jual_excl_gst = Format(LM_HARGA_JUALAN_DGN_GST - LM_GST_JUAL, "0.00")
            rs!harga_modal_gst = Format(LM_MODAL_DGN_GST - LM_MODAL_TANPA_GST, "0.00")
            rs!harga_modal_incl_gst = Format(LM_MODAL_DGN_GST, "0.00")
            rs!harga_modal_excl_gst = Format(LM_MODAL_TANPA_GST, "0.00")
            
            rs!untung = Format(LM_HARGA_JUALAN_DGN_GST - LM_GST_JUAL - LM_MODAL_TANPA_GST, "0.00")
            rs!untung2 = Format(LM_HARGA_JUALAN_DGN_GST - LM_MODAL_DGN_GST, "0.00")
            
            rs.Update
            Frm84_LM_DATA_SAVE = 1
        End If
        
        rs.Close
        Set rs = Nothing
'### Masukkan Data Ke Dalam Temp Table ### - End
        
        If Frm84_LM_DATA_SAVE = 1 Then
            'Call Frm84_Reset
            Call Frm84_Reset_Edit

            GM_NEXT_PREV = 2 '0 : Next , 1 : Previous
                    
            Call Frm84_Senarai_Jualan_Header
            Call Frm84_Senarai_Jualan
            
            Frm84.CMD3.Visible = True
            Frm84.CMD13.Visible = False
            Frm84.CMD14.Visible = False
            
            Frm84.TB1 = vbNullString
            Frm84.TB1.Locked = False
            Frm84.TB1.BackColor = &HFFFFFF
            
            MsgBox "Data Telah Berjaya Diedit.", vbInformation, "Info"
            Frm84.TB1.SetFocus
        End If
    End If
End If
End Sub
Private Sub CMD14_Click()
'on error resume next
Call Frm84_Reset_Edit

Frm84.CMD3.Visible = True
Frm84.CMD13.Visible = False
Frm84.CMD14.Visible = False

Frm84.TB1 = vbNullString
Frm84.TB1.Locked = False
Frm84.TB1.BackColor = &HFFFFFF
Frm84.TB1.SetFocus
End Sub
Private Sub CMD15_Click()
'On Error Resume Next
Call tesuto3
'Call Frm84_save_edit_data
End Sub
Private Sub CMD16_Click()
'on error resume next
If Frm84.L4_Text <> 0 Then

    Note = "Adakah mempunyai data yang belum disimpan." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Adakah anda ingin keluar dari menu ini?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then
    
        Frm85.Show
        Unload Frm84
        Unload Frm26
        Unload Frm27
        Unload Frm28
        Unload Frm83
        MDI_frm1.L5_Text = 12
        
    End If
    
Else

    Unload Frm84
    Unload Frm26
    Unload Frm27
    Unload Frm28
    Unload Frm83
    MDI_frm1.L5_Text = 12

End If
End Sub

Private Sub CMD19_Click()
'on error resume next
Frm55.Show
Frm84.Hide
Frm55.L16_Text = 2 '0 : NIL , 1 : Menu Ansuran , 2 : Menu Jualan
End Sub

Private Sub CMD17_Click()
'on error resume next
Dim frm84_LM_CURR_PAGE As Double
Dim frm84_LM_TOTAL_PAGE As Double

frm84_LM_CURR_PAGE = 0
frm84_LM_TOTAL_PAGE = 0

If Frm84.L87_Text <> vbNullString And IsNumeric(Frm84.L87_Text) Then
    If Frm84.L88_Text <> vbNullString And IsNumeric(Frm84.L88_Text) Then
        frm84_LM_CURR_PAGE = Frm84.L87_Text
        frm84_LM_TOTAL_PAGE = Frm84.L88_Text
        
        If frm84_LM_CURR_PAGE <> 1 And frm84_LM_CURR_PAGE <> 0 Then
        
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
                    
            Call Frm84_Senarai_Jualan_Header
            Call Frm84_Senarai_Jualan
                    
        End If

    End If
End If
End Sub

Private Sub CMD18_Click()
'on error resume next
Dim frm84_LM_CURR_PAGE As Double
Dim frm84_LM_TOTAL_PAGE As Double

frm84_LM_CURR_PAGE = 0
frm84_LM_TOTAL_PAGE = 0

If Frm84.L87_Text <> vbNullString And IsNumeric(Frm84.L87_Text) Then
    If Frm84.L88_Text <> vbNullString And IsNumeric(Frm84.L88_Text) Then
        frm84_LM_CURR_PAGE = Frm84.L87_Text
        frm84_LM_TOTAL_PAGE = Frm84.L88_Text
        
        If frm84_LM_CURR_PAGE < frm84_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm84_Senarai_Jualan_Header
            Call Frm84_Senarai_Jualan
            
        End If
    End If
End If
End Sub



Private Sub CMD2_Click()
'on error resume next
If Frm84.L4_Text <> 0 Then

    Note = "Adakah mempunyai data yang belum disimpan." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Adakah anda ingin keluar dari menu ini?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then
    
        Unload Frm84
        Unload Frm26
        Unload Frm27
        Unload Frm28
        Unload Frm83
        MDI_frm1.L5_Text = 0
        
    End If
    
Else

    Unload Frm84
    Unload Frm26
    Unload Frm27
    Unload Frm28
    Unload Frm83
    MDI_frm1.L5_Text = 0

End If
End Sub

Private Sub CMD20_Click()
'on error resume next
Note = "Adakah anda yakin untuk batalkan urusan ini ?" & vbCrLf & _
        "Semua data trade in yang telah dimasukkan akan dipadamkan jika anda teruskan." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    G_TI_MODE = 0
    Call frm_kiraan_harga_selepas_ti
End If
End Sub

Private Sub CMD21_Click()
'On Error Resume Next
If Frm84.L27_Text = vbNullString Then
    
    If Frm84.L28_Text <> vbNullString Then
    
        Note = "Anda telah mengisi data pembeli barangan ini di dalam ruangan pelanggan yang berdaftar dengan kedai." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Jika ada meneruskan menu ini , semua data pembeli di dalam ruangan pelanggan berdaftar akan dipadamkan." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
        
            'Unload Frm27
            'Unload Frm28
            Call Frm26_initial
            
            Frm84.L28_Text = vbNullString 'Nama pembeli : Berdaftar
            Frm84.L79_Text = 0 '0 : Tiada kad , 1 : Ada kad
            
            Frm26.Show 1
        End If
        
    Else
    
        'Unload Frm27
        'Unload Frm28
        Call Frm26_initial
        Frm84.L79_Text = 0 '0 : Tiada kad , 1 : Ada kad
        
        Frm26.Show 1
                
    End If
    
Else
    
    Frm84.L79_Text = 0 '0 : Tiada kad , 1 : Ada kad
    Frm26.Show 1
    
End If

End Sub


Private Sub CMD25_Click()
'On Error Resume Next
If Frm84.CB4 = 0 And Frm84.CB5 = 0 And Frm84.CB6 = 0 And Frm84.CB9 = 0 And Frm84.CB10 = 0 Then
    MsgBox "Sila pilih kategori pembeli.", vbExclamation, "info"
    Exit Sub
End If

Note = "Adakah anda ingin menukar kategori pembeli atau jualan oleh agen ?" & vbCrLf & _
        vbNullString & vbCrLf & _
        "Semua data jualan yang telah discan , maklumat pembeli dan data berkaitan dengan jualan ini akan dipadamkan." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    'Exit Sub
End If
If Answer = vbYes Then
    Call Frm84_Load_Form
    Call Frm84_Reset_Edit
    Unload Frm26
    Unload Frm27
    Unload Frm28
    Unload Frm83
    
    Frm84.Pic6.Visible = False
    
    If Frm84.CB7 = 1 Then
        MDI_frm1.L5_Text = 5
    Else
        MDI_frm1.L5_Text = 4
    End If
    
    Frm84.Frame1.Visible = True
    Frm84.TB1 = vbNullString
    Frm84.TB1.Locked = False
    Frm84.TB1.BackColor = &HFFFFFF
    Frm84.TB1.SetFocus
    
End If
End Sub
Private Sub CMD26_Click()
'On Error Resume Next
If Frm84.L64_Text = "1" Then

    Frm84.CB4 = 1
    
ElseIf Frm84.L64_Text = "2" Then

    Frm84.CB5 = 1
    
ElseIf Frm84.L64_Text = "3" Then
    
    Frm84.CB9 = 1
    
ElseIf Frm84.L64_Text = "4" Then
    
    Frm84.CB6 = 1
    
ElseIf Frm84.L64_Text = "5" Then
    
    Frm84.CB10 = 1

'ElseIf Frm84.L64_Text = "6" Then
    
'    Frm84.CB11 = 1
    
End If

If Frm84.L65_Text = 0 Then

    Frm84.CMD11.Enabled = False
    Frm84.CB7 = 0
    
ElseIf Frm84.L65_Text = 1 Then

    Frm84.CMD11.Enabled = True
    Frm84.CB7 = 1
    
End If

Frm84.Pic6.Visible = False
End Sub
Private Sub CMD29_Click()
'On Error Resume Next
If Frm84.CBB3 <> vbNullString And Frm84.CBB4 <> vbNullString Then
    Frm84.L83_Text = "1" '0 : Stok kedai , 1 : Barang trade in/potong
End If

Frm84.Pic8.Visible = False
End Sub

Private Sub CMD3_Click()
'On Error Resume Next
Dim Err(30)
Dim Frm84_LM_BERAT_ASAL As Double
Dim Frm84_LM_BERAT_JUAL As Double
Dim Frm84_LM_HARGA_MODAL As Double
Dim Frm84_LM_HARGA_JUAL As Double
Dim Frm84_LM_HARGA_SEMASA_MODAL As Double
Dim Frm84_LM_TETAPANHARGA As Double
Dim Frm84_LM_LIMIT As Double
Dim Frm84_LM_HARGA_STAFF As Double 'Tetapan harga jualan kepada staff
Dim Frm84_LM_HARGA_PELANGGAN As Double 'Tetapan harga jualan kepada pelanggan
Dim Frm84_LM_HARGA_SEMASA As Double 'Harga semasa (jualan)
Dim Frm84_LM_HARGA_SUPPLIER As Double 'Harga per gram (harga semasa) dari supplier (modal)
Dim Frm84_UPAH_MODAL As Double 'Upah modal
Dim Frm84_UPAH_JUAL As Double 'Upah jualan
Dim Frm84_LM_HARGA_JUALAN_CALC As Double 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
Dim Frm84_LM_GST_CALC As Double 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
Dim Frm84_KOMISYEN_UPAH As Double 'Komisyen dari upah kepada agen dropship
Dim Frm84_LM_SUSUT_BERAT As Double
Dim Frm84_LM_BERAT_OVERALL As Double

Frm84.L89_Text = -1 'Titik Pencarian Data
Frm84.L90_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Frm84.L87_Text = 0 'Paparan Page ke-xxx
Frm84.L88_Text = 0

GM_NEXT_PREV = 0
            
Call tesutochu

Exit Sub

Frm84_LM_BERAT_OVERALL = 0
Frm84_LM_SUSUT_BERAT = 0
Frm84_LM_HARGA_SEMASA = 0 'Harga semasa (jualan)
Frm84_LM_HARGA_SUPPLIER = 0 'Harga per gram (harga semasa) dari supplier (modal)
x = 0
Frm84_LM_BERAT_ASAL = 0
Frm84_LM_BERAT_JUAL = 0
Frm84_LM_DATA_SAVE = 0
Frm84_LM_HARGA_MODAL = 0
Frm84_LM_HARGA_JUAL = 0
Frm84_LM_HARGA_SEMASA_MODAL = 0
Frm84_LM_PRICE_CHECK = 0 '0 : Tidak Perlu Periksa Harga Semasa Jualan , 1 : Perlu Periksa Harga Semasa Jualan
Frm84_LM_TETAPANHARGA = 0
Frm84_LM_LIMIT = 0
Frm84_LM_HARGA_STAFF = 0
Frm84_LM_HARGA_PELANGGAN = 0
Frm84_UPAH_MODAL = 0 'Upah modal
Frm84_UPAH_JUAL = 0 'Upah jualan
Frm84_LM_HARGA_JUALAN_CALC = 0 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
Frm84_LM_GST_CALC = 0 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
Frm84_KOMISYEN_UPAH = 0 'Komisyen dari upah kepada agen dropship

If Frm84.TB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan [No. Siri Produk]."
End If
If Frm84.TB2 <> vbNullString And Frm84.TB3 = vbNullString And Frm84.CB12 = 1 Then
    MsgBox "Tetapan GST ke atas UPAH hanya dibenarkan untuk barang kemas SAHAJA. Sila periksa tetapan GST anda.", vbExclamation, "Info"
    Exit Sub
End If
'If (Frm84.TB14 <> vbNullString And IsNumeric(Frm84.TB14)) And (Frm84.L51_Text <> vbNullString And IsNumeric(Frm84.L51_Text)) Then
'    Frm84_LM_HARGA_STAFF = Frm84.L51_Text
'    Frm84_LM_HARGA_PELANGGAN = Frm84.TB14
    
'    If Frm84_LM_HARGA_PELANGGAN < Frm84_LM_HARGA_STAFF Then
'        X = X + 1
'        Err(X) = "Harga Jualan Minimum Yang Dibenarkan Adalah RM " & Format(Frm84_LM_HARGA_STAFF, "#,##0.00")
'    End If
'End If

'### Error Bagi Item BK ### - Start
If Frm84.TB3 <> vbNullString Then
    If Frm84.TB3 = vbNullString Or (Frm84.TB3 <> vbNullString And Not IsNumeric(Frm84.TB3)) Then
        x = x + 1
        Err(x) = "Sila Maklumat [Berat Asal]. Sila Scan Item Sekali Lagi."
    End If
    If Frm84.TB4 = vbNullString Or (Frm84.TB4 <> vbNullString And Not IsNumeric(Frm84.TB4)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Berat Jualan]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
    End If
    If Frm84.TB5 = vbNullString Or (Frm84.TB5 <> vbNullString And Not IsNumeric(Frm84.TB5)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Harga Semasa]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
    End If
    If Frm84.TB15 = vbNullString Or (Frm84.TB15 <> vbNullString And Not IsNumeric(Frm84.TB15)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Upah]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
    End If
    If Frm84.CB7 = 1 Then
        If Frm84.TB12 = vbNullString Or (Frm84.TB12 <> vbNullString And Not IsNumeric(Frm84.TB12)) Then
            x = x + 1
            Err(x) = "Sila Masukkan [Komisen Per Gram]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
        End If
        If Frm84.TB43 = vbNullString Or (Frm84.TB43 <> vbNullString And Not IsNumeric(Frm84.TB43)) Then
            x = x + 1
            Err(x) = "Sila Masukkan [Kadar Komisyen Upah (%)]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
        End If
        If Frm84.TB44 = vbNullString Or (Frm84.TB44 <> vbNullString And Not IsNumeric(Frm84.TB44)) Then
            x = x + 1
            Err(x) = "Sila Masukkan [Jumlah Komisyen Bagi Upah]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
        End If
        If (Frm84.TB15 <> vbNullString And IsNumeric(Frm84.TB15)) And (Frm84.TB44 <> vbNullString And IsNumeric(Frm84.TB44)) Then
            Frm84_UPAH_JUAL = Frm84.TB15 'Upah jualan
            Frm84_KOMISYEN_UPAH = Frm84.TB44 'Komisyen Upah
            
            If Frm84_KOMISYEN_UPAH > Frm84_UPAH_JUAL Then
                x = x + 1
                Err(x) = "Komisyen upah bagi agen dropship adalah melebihi dari upah asal."
            End If
        End If
    End If
End If
'### Error Bagi Item BK ### - End

'### Error Bagi Item Permata ### - Start
If Frm84.TB3 = vbNullString Then
    If Frm84.CB7 = 1 Then
        If Frm84.TB16 = vbNullString Or (Frm84.TB16 <> vbNullString And Not IsNumeric(Frm84.TB16)) Then
            x = x + 1
            Err(x) = "Sila Masukkan [Jumlah Komisen]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
        End If
    End If
End If
'### Error Bagi Item Permata ### - End

If Frm84.TB7 = vbNullString Or (Frm84.TB7 <> vbNullString And Not IsNumeric(Frm84.TB7)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Diskaun]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
End If
If Frm84.TB9 = vbNullString Or (Frm84.TB9 <> vbNullString And Not IsNumeric(Frm84.TB9)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Adjustment]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
End If
If Frm84.CB4 = 0 And Frm84.CB5 = 0 And Frm84.CB6 = 0 And Frm84.CB9 = 0 And Frm84.CB10 = 0 Then
    x = x + 1
    Err(x) = "Sila Buat Pilihan Kategori Pembeli."
End If
If Frm84.CB2 = 0 And Frm84.CB3 = 0 Then
    x = x + 1
    Err(x) = "Sila Buat Pilihan Jenis GST."
End If
If Frm84.TB10 = vbNullString Or (Frm84.TB10 <> vbNullString And Not IsNumeric(Frm84.TB10)) Then
    x = x + 1
    Err(x) = "Tiada maklumat [Harga Jualan]. Sila keluarkan item ini dari senarai dan scan sekali lagi."
End If
If Frm84.TB11 = vbNullString Or (Frm84.TB11 <> vbNullString And Not IsNumeric(Frm84.TB11)) Then
    x = x + 1
    Err(x) = "Tiada maklumat [Jumlah GST]. Sila keluarkan item ini dari senarai dan scan sekali lagi."
End If
If (Frm84.TB3 <> vbNullString And IsNumeric(Frm84.TB3)) And (Frm84.TB4 <> vbNullString And IsNumeric(Frm84.TB4)) Then
    Frm84_LM_BERAT_ASAL = Frm84.TB3 'Berat Asal
    Frm84_LM_BERAT_JUAL = Frm84.TB4 'Berat Jualan
    
    If Frm84_LM_BERAT_JUAL > Frm84_LM_BERAT_ASAL Then
        x = x + 1
        Err(x) = "Berat Jualan Melebihi Berat Asal."
    End If
End If
If Frm84.TB3 <> vbNullString And Frm84.L54_Text = vbNullString Then
    x = x + 1
    Err(x) = "Tiada maklumat harga dari supplier bagi purity ini." & vbCrLf & _
                "Sila pastikan harga dari supplier bagi purity ini telah ditetapkan dalam TETAPAN HARIAN SISTEM."
End If
If Frm84.L70_Text = 0 Then
    If Frm84.TB22 = vbNullString Or (Frm84.TB22 <> vbNullString And Not IsNumeric(Frm84.TB22)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Upah per gram]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
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
    Note = "Adakah Anda Ingin Masukkan Data Ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
'### Periksa Data Dulang ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where no_siri_Produk='" & Frm84.TB2 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!dulang) Then Frm84_LM_DULANG = rs!dulang 'Dulang
            If Not IsNull(rs!susut_berat) Then Frm84_LM_SUSUT_BERAT = rs!susut_berat 'Susut berat
        End If
        
        rs.Close
        Set rs = Nothing
'### Periksa Data Dulang ### - End

        If Frm84.TB3 <> vbNullString And Frm84.TB4 <> vbNullString Then
            If IsNumeric(Frm84.TB3) Then Frm84_LM_BERAT_ASAL = Frm84.TB3 'Berat Asal
            If IsNumeric(Frm84.TB4) Then Frm84_LM_BERAT_JUAL = Frm84.TB4 'Berat Jualan
            
            Frm84_LM_BERAT_OVERALL = Frm84_LM_SUSUT_BERAT + Frm84_LM_BERAT_JUAL
            
            If Frm84_LM_BERAT_ASAL < Frm84_LM_BERAT_OVERALL Then
            
                MsgBox "Berat jualan melebihi berat jualan yang dibenarkan." & vbCrLf & _
                        "Berat asal : " & Format(Frm84_LM_BERAT_ASAL, "#,##0.00 g") & vbCrLf & _
                        "Susut berat : " & Format(Frm84_LM_SUSUT_BERAT, "#,##0.00 g") & vbCrLf & _
                        "Berat jualan maksimum yang dibenarkan adalah " & Format(Frm84_LM_BERAT_ASAL - Frm84_LM_SUSUT_BERAT, "#,##0.00 g"), vbInformation, "Info"
                        
                Exit Sub
                        
            End If
        End If
    
'### Periksa Kadar Penurunan Harga ### - Start
'GoTo skip_periksa_harga:
        user = MDI_frm1.L3_Text
        
        If MDI_frm1.L4_Text <> vbNullString Then
            If MDI_frm1.L4_Text = "Staff" Then
                Frm84_LM_PRICE_CHECK = 1 '0 : Tidak Perlu Periksa Harga Semasa Jualan , 1 : Perlu Periksa Harga Semasa Jualan
            End If
        End If
'skip_periksa_harga:
        
        If Frm84_LM_PRICE_CHECK = 1 Then '0 : Tidak Perlu Periksa Harga Semasa Jualan , 1 : Perlu Periksa Harga Semasa Jualan
            Frm84_LM_LIMIT_TYPE = 0 '1 : BK , 2 : Barang Permata
            
'### Periksa Purity Dan Tetapan Harga Jualan ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from Data_Database where no_siri_Produk='" & Frm84.TB2 & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!dulang) Then Frm84_LM_DULANG = rs!dulang 'Dulang
                
                If Not IsNull(rs!receiving_Status) Then
                    If rs!receiving_Status = 0 Or rs!receiving_Status = 2 Then '0 : BK , 1 : Barang Permata , 2 : Buyback BK , 3 : Buyback Barang Permata
                        If Not IsNull(rs!kod_Purity) Then
                            Frm84_LM_PURITY = rs!kod_Purity 'Purity
                        End If
                        Frm84_LM_LIMIT_TYPE = 1 '1 : BK , 2 : Barang Permata
                    End If
                    If rs!receiving_Status = 1 Or rs!receiving_Status = 3 Then '0 : BK , 1 : Barang Permata , 2 : Buyback BK , 3 : Buyback Barang Permata
                        If Frm84.CB4 = 1 Then
                            If IsNumeric(rs!code_Supplier) Then Frm84_LM_TETAPANHARGA = Format(rs!code_Supplier, "0.00")  'Harga Pelanggan
                        ElseIf Frm84.CB5 = 1 Then
                            If IsNumeric(rs!HargaJualan_Member) Then Frm84_LM_TETAPANHARGA = Format(rs!HargaJualan_Member, "0.00") 'Harga Member
                        ElseIf Frm84.CB9 = 1 Then
                            If IsNumeric(rs!HargaJualan_RAF) Then Frm84_LM_TETAPANHARGA = Format(rs!HargaJualan_RAF, "0.00") 'Harga RAF
                        ElseIf Frm84.CB6 = 1 Then
                            If IsNumeric(rs!HargaJualan_Pengedar) Then Frm84_LM_TETAPANHARGA = Format(rs!HargaJualan_Pengedar, "0.00") 'Harga Pengedar
                        ElseIf Frm84.CB10 = 1 Then
                            If IsNumeric(rs!hargajualan_normal_dealer) Then Frm84_LM_TETAPANHARGA = Format(rs!hargajualan_normal_dealer, "0.00") 'Harga Normal Dealer
                        'ElseIf Frm84.CB11 = 1 Then
                        '    If IsNumeric(rs!hargajualan_master_dealer) Then Frm84_LM_TETAPANHARGA = Format(rs!hargajualan_master_dealer, "0.00") 'Harga Master Dealer
                        End If
                        Frm84_LM_LIMIT_TYPE = 2 '1 : BK , 2 : Barang Permata
                    End If
                End If
            End If
            
            rs.Close
            Set rs = Nothing
            
'### Carian Harga Semasa Emas ### - Start
            If Frm84_LM_LIMIT_TYPE = 1 Then '1 : BK , 2 : Barang Permata
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from default_setting where Default1='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    'If rs!Default1 = "Default" Then
                        If IsNumeric(rs!limit_per_gram) Then Frm84_LM_LIMIT = rs!limit_per_gram
                    'End If
                End If
                
                rs.Close
                Set rs = Nothing
            
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from hargaemas where Purity='" & Frm84_LM_PURITY & "' AND cawangan='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Frm84.CB4 = 1 Then
                        If IsNumeric(rs!Harga_Pelanggan) Then Frm84_LM_TETAPANHARGA = Format(rs!Harga_Pelanggan, "0.00") 'Harga Pelanggan
                    ElseIf Frm84.CB5 = 1 Then
                        If IsNumeric(rs!Harga_Member) Then Frm84_LM_TETAPANHARGA = Format(rs!Harga_Member, "0.00") 'Harga Member
                    ElseIf Frm84.CB9 = 1 Then
                        If IsNumeric(rs!Harga_RAF) Then Frm84_LM_TETAPANHARGA = Format(rs!Harga_RAF, "0.00") 'Harga RAF
                    ElseIf Frm84.CB6 = 1 Then
                        If IsNumeric(rs!Harga_Pengedar) Then Frm84_LM_TETAPANHARGA = Format(rs!Harga_Pengedar, "0.00") 'Harga Pengedar
                    ElseIf Frm84.CB10 = 1 Then
                        If IsNumeric(rs!harga_normal_dealer) Then Frm84_LM_TETAPANHARGA = Format(rs!harga_normal_dealer, "0.00") 'Harga Normal Dealer
                    'ElseIf Frm84.CB11 = 1 Then
                    '    If IsNumeric(rs!harga_master_dealer) Then Frm84_LM_TETAPANHARGA = Format(rs!harga_master_dealer, "0.00") 'Harga Master Dealer
                    End If
                End If
                
                rs.Close
                Set rs = Nothing
                
                If IsNumeric(Frm84.TB5) Then
                    Frm84_LM_HARGA_JUALAN = Frm84.TB5 'Harga Semasa Jualan (RM/g)
                End If
                
                If Frm84_LM_TETAPANHARGA - Frm84_LM_HARGA_JUALAN > Frm84_LM_LIMIT Then
                    MsgBox "Harga jualan tidak mengikut pengurangan harga minimum yang ditetapkan oleh kedai!." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Tetapan Asal Harga Jualan : RM " & Format(Frm84_LM_TETAPANHARGA, "0.00") & vbCrLf & _
                    "Limit Diskaun Pengurangan Harga : RM " & Format(Frm84_LM_LIMIT, "0.00") & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila periksa data anda.", vbExclamation, "Info"
                    Exit Sub
                End If
            End If
            
            If Frm84_LM_LIMIT_TYPE = 2 Then '1 : BK , 2 : Barang Permata
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If rs!Default1 = "Default" Then
                        If IsNumeric(rs!limit_per_item) Then Frm84_LM_LIMIT = rs!limit_per_item
                    End If
                End If
                
                rs.Close
                Set rs = Nothing
                
                If IsNumeric(Frm84.TB10) Then
                    Frm84_LM_HARGA_JUALAN = Frm84.TB10 'Harga Jualan (RM)
                End If
                
                If Frm84_LM_TETAPANHARGA - Frm84_LM_HARGA_JUALAN > Frm84_LM_LIMIT Then
                    MsgBox "Harga jualan tidak mengikut pengurangan harga minimum yang ditetapkan oleh kedai!." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Tetapan Asal Harga Jualan : RM " & Format(Frm84_LM_TETAPANHARGA, "0.00") & vbCrLf & _
                    "Limit Diskaun Pengurangan Harga : RM " & Format(Frm84_LM_LIMIT, "0.00") & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila periksa data anda.", vbExclamation, "Info"
                    Exit Sub
                End If
            End If
'### Carian Harga Semasa Emas ### - End
        
'### Periksa Purity Dan Tetapan Harga Jualan ### - End
        End If
'### Periksa Kadar Penurunan Harga ### - End

'### Masukkan Data Ke Dalam Temp Table ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_JUALAN_TEMP & " where no_siri_Produk='" & Frm84.TB2 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
            rs.AddNew
            If Frm84.TB2 <> vbNullString Then
                rs!no_siri_Produk = Frm84.TB2 'No. Siri Produk
            Else
                rs!no_siri_Produk = Null 'No. Siri Produk
            End If
            If Frm84.L12_Text <> vbNullString Then
                rs!kategori_Produk = Frm84.L12_Text 'Kategori Produk
            Else
                rs!kategori_Produk = Null 'Kategori Produk
            End If
            If Frm84.L13_Text <> vbNullString Then
                rs!purity = Frm84.L13_Text 'Purity
            Else
                rs!purity = Null 'Purity
            End If
            If Frm84.TB3 <> vbNullString Then
                rs!Berat_Asal = Format(Frm84.TB3, "0.00") 'Berat Asal (g)
            Else
                rs!Berat_Asal = Null 'Berat Asal (g)
            End If
            If Frm84.TB4 <> vbNullString Then
                rs!berat_jualan = Format(Frm84.TB4, "0.00") 'Berat Jualan (g)
            Else
                rs!berat_jualan = Null 'Berat Jualan (g)
            End If
            If Frm84.TB5 <> vbNullString Then
                rs!harga_Semasa = Format(Frm84.TB5, "0.00") 'Harga Semasa (RM/g)
            Else
                rs!harga_Semasa = Null 'Harga Semasa (RM/g)
            End If
            If Frm84.TB15 <> vbNullString Then
                rs!UPAH = Format(Frm84.TB15, "0.00") 'Upah (RM)
            Else
                rs!UPAH = Null 'Upah (RM)
            End If
            If Frm84.TB6 <> vbNullString Then
                rs!harga_asal = Format(Frm84.TB6, "0.00") 'Harga Asal Item (RM)
            Else
                rs!harga_asal = Null 'Harga Asal Item (RM)
            End If
            If Frm84.TB7 <> vbNullString Then
                rs!diskaun = Format(Frm84.TB7, "0.00") 'Diskaun (%)
            Else
                rs!diskaun = Null 'Diskaun (%)
            End If
            If Frm84.TB8 <> vbNullString Then
                rs!harga_lepas_diskaun = Format(Frm84.TB8, "0.00") 'Harga Selepas Diskaun (RM)
            Else
                rs!harga_lepas_diskaun = Null 'Harga Selepas Diskaun (RM)
            End If
            If Frm84.TB9 <> vbNullString Then
                rs!adjustment = Format(Frm84.TB9, "0.00") 'Adjustment (RM)
            Else
                rs!adjustment = Null 'Adjustment (RM)
            End If
            If Frm84.TB10 <> vbNullString Then
                rs!harga_jualan = Format(Frm84.TB10, "0.00") 'Harga Jualan (RM)
            Else
                rs!harga_jualan = Null 'Harga Jualan (RM)
            End If
            If Frm84.CB2 = 1 Then
                rs!gst_ari_nashi = "ZR (L)" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                rs!kadar_gst = Null 'Kadar Cukai GST (%)
                If Frm84.TB11 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm84.TB11, "0.00") 'Jumlah Cukai GST (RM)
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah Cukai GST (RM)
                End If
            ElseIf Frm84.CB3 = 1 Then
                rs!gst_ari_nashi = "SR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                If Frm84.L8_Text <> vbNullString Then
                    rs!kadar_gst = Frm84.L8_Text 'Kadar Cukai GST (%)
                Else
                    rs!kadar_gst = "0" 'Jumlah Cukai GST (RM)
                End If
                If Frm84.TB11 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm84.TB11, "0.00") 'Jumlah Cukai GST (RM)
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah Cukai GST (RM)
                End If
                If Frm84.CB18 = 1 Then 'Jenis Cukai GST SR
                    rs!gst_include = 1 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                Else
                    rs!gst_include = 0 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                End If
            End If
            If Frm84.L44_Text <> vbNullString Then
                rs!harga_tanpa_gst = Format(Frm84.L44_Text, "0.00") 'Harga Jualan Tanpa GST (RM)
            Else
                rs!harga_tanpa_gst = Null 'Harga Jualan Tanpa GST (RM)
            End If
            If Frm84.TB14 <> vbNullString Then
                rs!harga_dengan_gst = Format(Frm84.TB14, "0.00") 'Harga Jualan Termasuk GST (RM)
            Else
                rs!harga_dengan_gst = Null 'Harga Jualan Termasuk GST (RM)
            End If
            If Frm84.CB7 = 1 Then
                rs!dropship = 1 '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
                If Frm84.Frame2.Visible = True Then 'Komisen Agen Dropship : BK
                    If Frm84.TB12 <> vbNullString Then
                        rs!komisyen_per_gram = Format(Frm84.TB12, "0.00") 'Komisyen Per Gram Dropship (RM/g) : BK
                    Else
                        rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g) : BK
                    End If
                    If Frm84.TB13 <> vbNullString Then
                        rs!jumlah_komisyen = Format(Frm84.TB13, "0.00") 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
                    Else
                        rs!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
                    End If
                    If Frm84.TB43 <> vbNullString Then
                        rs!kadar_komisyen_upah = Frm84.TB43 'Kadar komisyen bagi upah kepada agen dropship
                    Else
                        rs!kadar_komisyen_upah = Null 'Kadar komisyen bagi upah kepada agen dropship
                    End If
                    If Frm84.TB44 <> vbNullString Then
                        rs!komisyen_upah = Format(Frm84.TB44, "0.00") 'Jumlah komisyen bagi upah kepada agen dropship
                    Else
                        rs!komisyen_upah = Null 'Jumlah komisyen bagi upah kepada agen dropship
                    End If
                End If
                If Frm84.Frame3.Visible = True Then 'Komisen Agen Dropship : Permata
                    rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g) : BK
                    rs!kadar_komisyen_upah = Null 'Kadar komisyen bagi upah kepada agen dropship
                    rs!komisyen_upah = Null 'Jumlah komisyen bagi upah kepada agen dropship
                    If Frm84.TB16 <> vbNullString Then
                        rs!jumlah_komisyen = Format(Frm84.TB16, "0.00") 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : Permata
                    Else
                        rs!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : Permata
                    End If
                End If
            End If
                
            If Frm84.CB7 = 0 Then
                rs!dropship = 0 '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
                rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g)
                rs!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini
                rs!kadar_komisyen_upah = Null 'Kadar komisyen bagi upah kepada agen dropship
                rs!komisyen_upah = Null 'Jumlah komisyen bagi upah kepada agen dropship
            End If
            
            If Frm84.L41_Text = "0" Then '0 : Menu Data Baru , 1 : Menu Edit Data
                rs!Status = 1
            ElseIf Frm84.L41_Text = "1" Then
                rs!Status = 3
            End If
            
            If Frm84.TB3 = vbNullString Then
            
                rs!Type = 1 '0 : BK , 1 : Barang Permata
                rs!harga_per_gram_modal = Null 'Harga Per Gram Bagi Modal (RM/g)
                If Frm84.L34_Text <> vbNullString Then
                    rs!modal = Format(Frm84.L34_Text, "0.00") 'Harga Modal (RM)
                Else
                    rs!modal = Null 'Harga Modal (RM)
                End If
                If IsNumeric(Frm84.TB14) And IsNumeric(Frm84.L34_Text) Then
                    Frm84_LM_HARGA_MODAL = Frm84.L34_Text 'Harga Modal
                    Frm84_LM_HARGA_JUAL = Frm84.TB14 'Harga Jualan

                    rs!untung = Format(Frm84_LM_HARGA_JUAL - Frm84_LM_HARGA_MODAL, "0.00") 'Jumlah Keuntungan
                End If
                rs!harga_per_gram_supplier = Null 'Harga per gram (harga semasa) dari supplier (modal)
                rs!untung2 = Null 'Untung jika restok pada harga supplier ini
                rs!upah_modal = Null 'Upah modal
                
            Else
            
                rs!Type = 0 '0 : BK , 1 : Barang Permata
                
                If Frm84.L34_Text <> vbNullString Then
                    rs!harga_per_gram_modal = Format(Frm84.L34_Text, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
                    If IsNumeric(Frm84.L34_Text) Then
                        Frm84_LM_HARGA_SEMASA_MODAL = Frm84.L34_Text 'Harga Per Gram Bagi Modal (RM/g)
                        
                        rs!modal = Format(Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_SEMASA_MODAL, "0.00") 'Harga Modal (RM)
                    End If
                Else
                    rs!modal = Null 'Harga Modal (RM)
                    rs!harga_per_gram_modal = Null 'Harga Per Gram Bagi Modal (RM/g)
                End If
                
                If IsNumeric(Frm84.TB14) And IsNumeric(Frm84.L34_Text) And IsNumeric(Frm84.TB10) And IsNumeric(Frm84.TB11) Then
                    
                    Frm84_LM_HARGA_MODAL = Frm84.L34_Text 'Harga Modal
                    Frm84_LM_HARGA_JUAL = Frm84.TB14 'Harga Jualan
                    Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
                    Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
                    
                    If Frm84.CB12 = 0 Then
                    
                        rs!untung = Format(Frm84_LM_HARGA_JUAL - (Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_MODAL), "0.00") 'Jumlah Keuntungan
                        
                    ElseIf Frm84.CB12 = 1 Then
                        
                        If Frm84.CB3 = 1 And Frm84.CB18 = 0 Then
                            
                            rs!untung = Format((Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC) - (Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_MODAL), "0.00") 'Jumlah Keuntungan
                            
                        Else
                            
                            rs!untung = Format(Frm84_LM_HARGA_JUALAN_CALC - (Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_MODAL), "0.00") 'Jumlah Keuntungan
                            
                        End If
                        
                    End If
                    
                End If
                
                If IsNumeric(Frm84.TB4) And IsNumeric(Frm84.TB5) And IsNumeric(Frm84.L54_Text) And IsNumeric(Frm84.L55_Text) And IsNumeric(Frm84.TB15) And IsNumeric(Frm84.TB3) Then
                    
                    Frm84_LM_BERAT_JUAL = Frm84.TB4 'Berat Jualan
                    Frm84_LM_HARGA_SEMASA = Frm84.TB5 'Harga semasa (jualan)
                    Frm84_LM_HARGA_SUPPLIER = Frm84.L54_Text 'Harga per gram (harga semasa) dari supplier (modal)
                    Frm84_UPAH_MODAL = Frm84.L55_Text 'Upah modal
                    Frm84_UPAH_JUAL = Frm84.TB15 'Upah jualan
                    Frm84_LM_BERAT_ASAL = Frm84.TB3 'Berat Asal
                    
                    rs!upah_modal = Frm84.L55_Text 'Upah modal
                    rs!harga_per_gram_supplier = Frm84.L54_Text 'Harga per gram (harga semasa) dari supplier (modal)
                    rs!untung2 = Format(((Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_SEMASA) + Frm84_UPAH_JUAL) - ((Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_SUPPLIER) + (Frm84_LM_BERAT_JUAL * Frm84_UPAH_MODAL / Frm84_LM_BERAT_ASAL)), "0.00") 'Untung jika restok pada harga supplier ini

                Else
                    
                    rs!harga_per_gram_supplier = "0.00" 'Harga per gram (harga semasa) dari supplier (modal)
                    rs!untung2 = "0.00" 'Untung jika restok pada harga supplier ini
                    rs!upah_modal = "0.00" 'Upah modal
                    
                End If
                
            End If
            If Format(Frm84.TB3, "0.00") = Format(Frm84.TB4, "0.00") Then
                rs!potong_flag = 0 '0 : Tiada Potong , 1 : Ada Potong
            Else
                rs!potong_flag = 1 '0 : Tiada Potong , 1 : Ada Potong
            End If
            rs!dulang = Frm84_LM_DULANG 'Dulang
            
            '### Maklumat tetapan harga jualan kepada staff ### - Start
            If Frm84.L48_Text <> vbNullString Then
                rs!kadar_penurunan_upah = Format(Frm84.L48_Text, "0.00") 'Kadar peratusan penurunan harga upah kepada staff (%)
            Else
                rs!kadar_penurunan_upah = Null
            End If
            If Frm84.L49_Text <> vbNullString Then
                rs!harga_semasa_staff = Format(Frm84.L49_Text, "0.00") 'Harga emas semasa yang dijual kepada staff
            Else
                rs!harga_semasa_staff = Null
            End If
            If Frm84.L50_Text <> vbNullString Then
                rs!kadar_penurunan_bp = Format(Frm84.L50_Text, "0.00") 'Kadar peratusan penurunan harga barang permata kepada staff (%)
            Else
                rs!kadar_penurunan_bp = Null
            End If
            If Frm84.L51_Text <> vbNullString Then
                rs!harga_staff = Format(Frm84.L51_Text, "0.00") 'Harga yang dijual kepada staff (RM)
            Else
                rs!harga_staff = Null
            End If
            If Frm84.L52_Text <> vbNullString Then
                rs!harga_bp_asal = Format(Frm84.L52_Text, "0.00") 'Tetapan harga barang permata yang asal (RM)
            Else
                rs!harga_bp_asal = Null
            End If
            If Frm84.L53_Text <> vbNullString Then
                rs!upah_asal = Format(Frm84.L53_Text, "0.00") 'Tetapan upah asal (RM)
            Else
                rs!upah_asal = Null
            End If
            rs!komisyen_staff = Format(Frm84_LM_HARGA_PELANGGAN - Frm84_LM_HARGA_STAFF, "0.00") 'Jumlah Komisyen Staff (RM)
            '### Maklumat tetapan harga jualan kepada staff ### - End
            
            If Frm84.CB12 = 0 Then '0 : GST pada harga jualan , 1 : GST pada upah
                rs!gst_barang_atau_upah = 0
            Else
                rs!gst_barang_atau_upah = 1
            End If
            If Frm84.CB3 = 1 And Frm84.CB18 = 0 Then
                Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
                Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
                
                rs!harga_jualan_dengan_gst = Format(Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC, "0.00")
                'Field ini adalah lebih kurang kepada @harga_dengan_gst
                'Tetapi bagi gst yang dikenakan hanya pada UPAH akan membuatkan field ini hanya memaparkan harga dengan gst hanya pada upah.
                'Oleh itu field ini dibuat khas bagi mencampurkan gst pada upah dam harga barang.
            Else
                rs!harga_jualan_dengan_gst = Format(Frm84.TB10, "0.00") 'Harga Jualan (RM)
            End If
            If Frm84.L67_Text <> vbNullString Then 'Purata harga jualan per gram (RM/g) bagi barang kemas , Bagi barang permata adalah merujuk kepada harga jualan
                rs!jualan_per_gram = Format(Frm84.L67_Text, "0.00")
            Else
                rs!jualan_per_gram = Null
            End If
            If Frm84.L69_Text <> vbNullString Then 'Paparan modal per gram (tanpa GST)
                rs!modal_per_gram = Format(Frm84.L69_Text, "0.00")
            Else
                rs!modal_per_gram = Null
            End If
            If Frm84.L70_Text = 0 Then
                
                rs!flag_upah = 0
                
                If Frm84.TB22 <> vbNullString Then
                
                    rs!upah_per_gram = Format(Frm84.TB22, "0.00")
                
                Else
                
                    rs!upah_per_gram = "0.00"
                    
                End If
            
            ElseIf Frm84.L70_Text = 1 Then
                
                rs!flag_upah = 1
                rs!upah_per_gram = Null
            
            End If
            
            rs.Update
            Frm84_LM_DATA_SAVE = 1
        Else
            If Frm84.TB2 <> vbNullString Then
                rs!no_siri_Produk = Frm84.TB2 'No. Siri Produk
            Else
                rs!no_siri_Produk = Null 'No. Siri Produk
            End If
            If Frm84.L12_Text <> vbNullString Then
                rs!kategori_Produk = Frm84.L12_Text 'Kategori Produk
            Else
                rs!kategori_Produk = Null 'Kategori Produk
            End If
            If Frm84.L13_Text <> vbNullString Then
                rs!purity = Frm84.L13_Text 'Purity
            Else
                rs!purity = Null 'Purity
            End If
            If Frm84.TB3 <> vbNullString Then
                rs!Berat_Asal = Format(Frm84.TB3, "0.00") 'Berat Asal (g)
            Else
                rs!Berat_Asal = Null 'Berat Asal (g)
            End If
            If Frm84.TB4 <> vbNullString Then
                rs!berat_jualan = Format(Frm84.TB4, "0.00") 'Berat Jualan (g)
            Else
                rs!berat_jualan = Null 'Berat Jualan (g)
            End If
            If Frm84.TB5 <> vbNullString Then
                rs!harga_Semasa = Format(Frm84.TB5, "0.00") 'Harga Semasa (RM/g)
            Else
                rs!harga_Semasa = Null 'Harga Semasa (RM/g)
            End If
            If Frm84.TB15 <> vbNullString Then
                rs!UPAH = Format(Frm84.TB15, "0.00") 'Upah (RM)
            Else
                rs!UPAH = Null 'Upah (RM)
            End If
            If Frm84.TB6 <> vbNullString Then
                rs!harga_asal = Format(Frm84.TB6, "0.00") 'Harga Asal Item (RM)
            Else
                rs!harga_asal = Null 'Harga Asal Item (RM)
            End If
            If Frm84.TB7 <> vbNullString Then
                rs!diskaun = Format(Frm84.TB7, "0.00") 'Diskaun (%)
            Else
                rs!diskaun = Null 'Diskaun (%)
            End If
            If Frm84.TB8 <> vbNullString Then
                rs!harga_lepas_diskaun = Format(Frm84.TB8, "0.00") 'Harga Selepas Diskaun (RM)
            Else
                rs!harga_lepas_diskaun = Null 'Harga Selepas Diskaun (RM)
            End If
            If Frm84.TB9 <> vbNullString Then
                rs!adjustment = Format(Frm84.TB9, "0.00") 'Adjustment (RM)
            Else
                rs!adjustment = Null 'Adjustment (RM)
            End If
            If Frm84.TB10 <> vbNullString Then
                rs!harga_jualan = Format(Frm84.TB10, "0.00") 'Harga Jualan (RM)
            Else
                rs!harga_jualan = Null 'Harga Jualan (RM)
            End If
            If Frm84.CB2 = 1 Then
                rs!gst_ari_nashi = "ZR (L)" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                rs!kadar_gst = Null 'Kadar Cukai GST (%)
                
                If Frm84.TB11 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm84.TB11, "0.00") 'Jumlah Cukai GST (RM)
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah Cukai GST (RM)
                End If
            ElseIf Frm84.CB3 = 1 Then
                rs!gst_ari_nashi = "SR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                If Frm84.L8_Text <> vbNullString Then
                    rs!kadar_gst = Frm84.L8_Text 'Kadar Cukai GST (%)
                Else
                    rs!kadar_gst = "0" 'Jumlah Cukai GST (RM)
                End If
                If Frm84.TB11 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm84.TB11, "0.00") 'Jumlah Cukai GST (RM)
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah Cukai GST (RM)
                End If
                If Frm84.CB18 = 1 Then 'Jenis Cukai GST SR
                    rs!gst_include = 1 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                Else
                    rs!gst_include = 0 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                End If
            End If
            If Frm84.L44_Text <> vbNullString Then
                rs!harga_tanpa_gst = Format(Frm84.L44_Text, "0.00") 'Harga Jualan Tanpa GST (RM)
            Else
                rs!harga_tanpa_gst = Null 'Harga Jualan Tanpa GST (RM)
            End If
            If Frm84.TB14 <> vbNullString Then
                rs!harga_dengan_gst = Format(Frm84.TB14, "0.00") 'Harga Jualan Termasuk GST (RM)
            Else
                rs!harga_dengan_gst = Null 'Harga Jualan Termasuk GST (RM)
            End If
            If Frm84.CB7 = 1 Then
                rs!dropship = 1 '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
                If Frm84.Frame2.Visible = True Then 'Komisen Agen Dropship : BK
                    If Frm84.TB12 <> vbNullString Then
                        rs!komisyen_per_gram = Format(Frm84.TB12, "0.00") 'Komisyen Per Gram Dropship (RM/g) : BK
                    Else
                        rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g) : BK
                    End If
                    If Frm84.TB13 <> vbNullString Then
                        rs!jumlah_komisyen = Format(Frm84.TB13, "0.00") 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
                    Else
                        rs!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
                    End If
                End If
                If Frm84.Frame3.Visible = True Then 'Komisen Agen Dropship : Permata
                    rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g) : BK
                    If Frm84.TB16 <> vbNullString Then
                        rs!jumlah_komisyen = Format(Frm84.TB16, "0.00") 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : Permata
                    Else
                        rs!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : Permata
                    End If
                End If
            End If
                
            If Frm84.CB7 = 0 Then
                rs!dropship = 0 '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
                rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g)
                rs!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini
            End If
            
            If Frm84.L41_Text = "0" Then '0 : Menu Data Baru , 1 : Menu Edit Data
                rs!Status = 1
            ElseIf Frm84.L41_Text = "1" Then
                rs!Status = 3
            End If
            
            If Frm84.TB3 = vbNullString Then
                rs!Type = 1 '0 : BK , 1 : Barang Permata
                rs!harga_per_gram_modal = Null 'Harga Per Gram Bagi Modal (RM/g)
                If Frm84.L34_Text <> vbNullString Then
                    rs!modal = Format(Frm84.L34_Text, "0.00") 'Harga Modal (RM)
                Else
                    rs!modal = Null 'Harga Modal (RM)
                End If
                If IsNumeric(Frm84.TB14) And IsNumeric(Frm84.L34_Text) Then
                    Frm84_LM_HARGA_MODAL = Frm84.L34_Text 'Harga Modal
                    Frm84_LM_HARGA_JUAL = Frm84.TB14 'Harga Jualan
                    
                    rs!untung = Format(Frm84_LM_HARGA_JUAL - Frm84_LM_HARGA_MODAL, "0.00") 'Jumlah Keuntungan
                End If

                rs!harga_per_gram_supplier = Null 'Harga per gram (harga semasa) dari supplier (modal)
                rs!untung2 = Null 'Untung jika restok pada harga supplier ini
                rs!upah_modal = Null 'Upah modal
                
            Else
                rs!Type = 0 '0 : BK , 1 : Barang Permata
                
                If Frm84.L34_Text <> vbNullString Then
                    rs!harga_per_gram_modal = Format(Frm84.L34_Text, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
                    If IsNumeric(Frm84.L34_Text) Then
                        Frm84_LM_HARGA_SEMASA_MODAL = Frm84.L34_Text 'Harga Per Gram Bagi Modal (RM/g)
                        
                        rs!modal = Format(Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_SEMASA_MODAL, "0.00") 'Harga Modal (RM)
                    End If
                Else
                    rs!modal = Null 'Harga Modal (RM)
                    rs!harga_per_gram_modal = Null 'Harga Per Gram Bagi Modal (RM/g)
                End If
                
                If IsNumeric(Frm84.TB14) And IsNumeric(Frm84.L34_Text) And IsNumeric(Frm84.TB10) And IsNumeric(Frm84.TB11) Then
                    
                    Frm84_LM_HARGA_MODAL = Frm84.L34_Text 'Harga Modal
                    Frm84_LM_HARGA_JUAL = Frm84.TB14 'Harga Jualan
                    Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
                    Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
                    
                    If Frm84.CB12 = 0 Then
                    
                        rs!untung = Format(Frm84_LM_HARGA_JUAL - (Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_MODAL), "0.00") 'Jumlah Keuntungan
                        
                    ElseIf Frm84.CB12 = 1 Then
                        
                        If Frm84.CB3 = 1 And Frm84.CB18 = 0 Then
                            
                            rs!untung = Format((Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC) - (Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_MODAL), "0.00") 'Jumlah Keuntungan
                            
                        Else
                            
                            rs!untung = Format(Frm84_LM_HARGA_JUALAN_CALC - (Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_MODAL), "0.00") 'Jumlah Keuntungan
                            
                        End If
                        
                    End If
                End If

                If IsNumeric(Frm84.TB4) And IsNumeric(Frm84.TB5) And IsNumeric(Frm84.L54_Text) And IsNumeric(Frm84.L55_Text) And IsNumeric(Frm84.TB15) And IsNumeric(Frm84.TB3) Then
                    
                    Frm84_LM_BERAT_JUAL = Frm84.TB4 'Berat Jualan
                    Frm84_LM_HARGA_SEMASA = Frm84.TB5 'Harga semasa (jualan)
                    Frm84_LM_HARGA_SUPPLIER = Frm84.L54_Text 'Harga per gram (harga semasa) dari supplier (modal)
                    Frm84_UPAH_MODAL = Frm84.L55_Text 'Upah modal
                    Frm84_UPAH_JUAL = Frm84.TB15 'Upah jualan
                    Frm84_LM_BERAT_ASAL = Frm84.TB3 'Berat Asal
                    
                    rs!upah_modal = Frm84.L55_Text 'Upah modal
                    rs!harga_per_gram_supplier = Frm84.L54_Text 'Harga per gram (harga semasa) dari supplier (modal)
                    rs!untung2 = Format(((Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_SEMASA) + Frm84_UPAH_JUAL) - ((Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_SUPPLIER) + (Frm84_LM_BERAT_JUAL * Frm84_UPAH_MODAL / Frm84_LM_BERAT_ASAL)), "0.00") 'Untung jika restok pada harga supplier ini

                Else
                    
                    rs!harga_per_gram_supplier = "0.00" 'Harga per gram (harga semasa) dari supplier (modal)
                    rs!untung2 = "0.00" 'Untung jika restok pada harga supplier ini
                    rs!upah_modal = "0.00" 'Upah modal
                    
                End If
                
            End If
            If Format(Frm84.TB3, "0.00") = Format(Frm84.TB4, "0.00") Then
                rs!potong_flag = 0 '0 : Tiada Potong , 1 : Ada Potong
            Else
                rs!potong_flag = 1 '0 : Tiada Potong , 1 : Ada Potong
            End If
            
            rs!dulang = Frm84_LM_DULANG 'Dulang
            
            '### Maklumat tetapan harga jualan kepada staff ### - Start
            If Frm84.L48_Text <> vbNullString Then
                rs!kadar_penurunan_upah = Format(Frm84.L48_Text, "0.00") 'Kadar peratusan penurunan harga upah kepada staff (%)
            Else
                rs!kadar_penurunan_upah = Null
            End If
            If Frm84.L49_Text <> vbNullString Then
                rs!harga_semasa_staff = Format(Frm84.L49_Text, "0.00") 'Harga emas semasa yang dijual kepada staff
            Else
                rs!harga_semasa_staff = Null
            End If
            If Frm84.L50_Text <> vbNullString Then
                rs!kadar_penurunan_bp = Format(Frm84.L50_Text, "0.00") 'Kadar peratusan penurunan harga barang permata kepada staff (%)
            Else
                rs!kadar_penurunan_bp = Null
            End If
            If Frm84.L51_Text <> vbNullString Then
                rs!harga_staff = Format(Frm84.L51_Text, "0.00") 'Harga yang dijual kepada staff (RM)
            Else
                rs!harga_staff = Null
            End If
            If Frm84.L52_Text <> vbNullString Then
                rs!harga_bp_asal = Format(Frm84.L52_Text, "0.00") 'Tetapan harga barang permata yang asal (RM)
            Else
                rs!harga_bp_asal = Null
            End If
            If Frm84.L53_Text <> vbNullString Then
                rs!upah_asal = Format(Frm84.L53_Text, "0.00") 'Tetapan upah asal (RM)
            Else
                rs!upah_asal = Null
            End If
            rs!komisyen_staff = Format(Frm84_LM_HARGA_PELANGGAN - Frm84_LM_HARGA_STAFF, "0.00") 'Jumlah Komisyen Staff (RM)
            '### Maklumat tetapan harga jualan kepada staff ### - End
            
            If Frm84.CB12 = 0 Then '0 : GST pada harga jualan , 1 : GST pada upah
                rs!gst_barang_atau_upah = 0
            Else
                rs!gst_barang_atau_upah = 1
            End If
            If Frm84.CB3 = 1 And Frm84.CB18 = 0 Then
                Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
                Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
                
                rs!harga_jualan_dengan_gst = Format(Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC, "0.00")
                'Field ini adalah lebih kurang kepada @harga_dengan_gst
                'Tetapi bagi gst yang dikenakan hanya pada UPAH akan membuatkan field ini hanya memaparkan harga dengan gst hanya pada upah.
                'Oleh itu field ini dibuat khas bagi mencampurkan gst pada upah dam harga barang.
            Else
                rs!harga_jualan_dengan_gst = Format(Frm84.TB10, "0.00") 'Harga Jualan (RM)
            End If
            If Frm84.L67_Text <> vbNullString Then 'Purata harga jualan per gram (RM/g) bagi barang kemas , Bagi barang permata adalah merujuk kepada harga jualan
                rs!jualan_per_gram = Format(Frm84.L67_Text, "0.00")
            Else
                rs!jualan_per_gram = Null
            End If
            If Frm84.L69_Text <> vbNullString Then 'Paparan modal per gram (tanpa GST)
                rs!modal_per_gram = Format(Frm84.L69_Text, "0.00")
            Else
                rs!modal_per_gram = Null
            End If
            If Frm84.L70_Text = 0 Then
                
                rs!flag_upah = 0
                
                If Frm84.TB22 <> vbNullString Then
                
                    rs!upah_per_gram = Format(Frm84.TB22, "0.00")
                
                Else
                
                    rs!upah_per_gram = "0.00"
                    
                End If
            
            ElseIf Frm84.L71_Text = 1 Then
                
                rs!flag_upah = 1
                rs!upah_per_gram = Null
            
            End If

            rs.Update
            Frm84_LM_DATA_SAVE = 1
        End If
        
        rs.Close
        Set rs = Nothing
'### Masukkan Data Ke Dalam Temp Table ### - End
        
        If Frm84_LM_DATA_SAVE = 1 Then
            'Call Frm84_Reset
            Call Frm84_Reset_Edit
            Call Frm84_Senarai_Jualan_Header
            Call Frm84_Senarai_Jualan
            
            MsgBox "Data Telah Berjaya Dimasukkan Ke Dalam Senarai Jualan.", vbInformation, "Info"
            
            Frm84.TB1.SetFocus
        End If
    End If
End If
End Sub

Private Sub CMD30_Click()
'on error resume next
'Call frm84_senarai_barang_purity
If Frm84.TB2 <> vbNullString Then
    
    If Frm84.TB2 <> "-" Then
    
        Note = "Terdapat barang stok kedai yang telah discan dan cuba dijual." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Jika anda meneruskan menu ini , semua data jualan ini akan dipadamkan." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan?"
            
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
        
            Exit Sub
        
        End If
        
    End If
            
End If

If Frm84.L83_Text = "0" Then
    Call frm84_senarai_barang_purity  '0 : Stok kedai , 1 : Barang trade in/potong
    
    Call Frm84_Reset_Edit
    
    Frm84.CMD3.Visible = True
    Frm84.CMD13.Visible = False
    Frm84.CMD14.Visible = False
    
    Frm84.TB1 = vbNullString
    Frm84.TB1.Locked = False
    Frm84.TB1.BackColor = &HFFFFFF

End If

Frm84.Pic8.Visible = True
End Sub

Private Sub CMD31_Click()
'On Error Resume Next
Dim Err(15)

If Frm84.CB12 = 1 Then
    x = x + 1
    Err(x) = "Penetapan cukai GST bagi UPAH sahaja tidak dibenarkan bagi menu ini."
End If
If Frm84.L83_Text = "0" Then '0 : Stok kedai , 1 : Barang trade in/potong
    If Frm84.TB2 = vbNullString Then
        x = x + 1
        Err(x) = "Tiada detail barang yang ingin dijual (Tiada maklumat No. Siri Produk)."
    End If
End If
If Frm84.TB2 <> vbNullString And Frm84.TB3 = vbNullString And Frm84.CB12 = 1 Then
    MsgBox "Tetapan GST ke atas UPAH hanya dibenarkan untuk barang kemas SAHAJA. Sila periksa tetapan GST anda.", vbExclamation, "Info"
    Exit Sub
End If

'### Error Bagi Item BK ### - Start
If Frm84.TB3 <> vbNullString Then

    If Frm84.TB3 = vbNullString Or (Frm84.TB3 <> vbNullString And Not IsNumeric(Frm84.TB3)) Then
        x = x + 1
        Err(x) = "Tiada maklumat berat asal.Sila scan atau masukkan no. siri produk dan cari maklumat terperinci item ini."
    End If
    If Frm84.TB4 = vbNullString Or (Frm84.TB4 <> vbNullString And Not IsNumeric(Frm84.TB4)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Berat Jualan]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm84.TB5 = vbNullString Or (Frm84.TB5 <> vbNullString And Not IsNumeric(Frm84.TB5)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Harga Semasa]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm84.TB15 = vbNullString Or (Frm84.TB15 <> vbNullString And Not IsNumeric(Frm84.TB15)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Upah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If

End If
'### Error Bagi Item BK ### - End

If Frm84.TB6 = vbNullString Or (Frm84.TB6 <> vbNullString And Not IsNumeric(Frm84.TB6)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Harga Asal]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm84.TB7 = vbNullString Or (Frm84.TB7 <> vbNullString And Not IsNumeric(Frm84.TB7)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Diskaun]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm84.TB9 = vbNullString Or (Frm84.TB9 <> vbNullString And Not IsNumeric(Frm84.TB9)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Adjustment]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm84.CB2 = 0 And Frm84.CB3 = 0 And Frm84.CB18 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan jenis cukai GST."
End If
If Frm84.TB10 = vbNullString Or (Frm84.TB10 <> vbNullString And Not IsNumeric(Frm84.TB10)) Then
    x = x + 1
    Err(x) = "Tiada maklumat [Harga Jualan]. Sila keluarkan item ini dari senarai dan scan sekali lagi."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
 
Dim LM_HARGA_JUALAN As Double
Dim LM_HARGA_LEPAS_DISKAUN As Double
Dim LM_ADJUSTMENT As Double
Dim frm84_LM_KADAR_GST As Double
Dim LM_HARGA_BEFORE_GST As Double
Dim LM_BERAT_CHECK As Double
Dim LM_UPAH As Double

LM_HARGA_JUALAN = 0
LM_HARGA_LEPAS_DISKAUN = 0
LM_ADJUSTMENT = 0
frm84_LM_KADAR_GST = 0
LM_HARGA_BEFORE_GST = 0
LM_HARGA_JUALAN = 0
LM_UPAH = 0

LM_BERAT = 1
LM_BERAT_CHECK = 0

    Note = "Sila masukkan harga jualan barang ini." & vbCrLf & _
            "Hanya NOMBOR dibenarkan di dalam ruangan ini." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Sistem akan mengirakan jumlah cukai GST dan akan masukkan jumlah ADJUSTMENT."

    LM_INPUT_HARGA = InputBox(Note, "Harga Jualan", "0.00")
    
    If StrPtr(LM_INPUT_HARGA) = 0 Then
        Exit Sub
    End If
    
    If IsNumeric(LM_INPUT_HARGA) Then
    
        LM_HARGA_JUALAN = LM_INPUT_HARGA
        
        G_CALC_AUTO = 1
        
        'If Frm84.TB3 <> vbNullString Then
            
            If Frm84.TB4 <> vbNullString And IsNumeric(Frm84.TB4) Then
                
                LM_BERAT_CHECK = Frm84.TB4
                
                If LM_BERAT_CHECK <> 0 Then
                
                    LM_BERAT = LM_BERAT_CHECK
                    
                    If Frm84.TB15 <> vbNullString And IsNumeric(Frm84.TB15) Then LM_UPAH = Frm84.TB15
                    
                    Frm84.TB5 = Format((LM_HARGA_JUALAN - LM_UPAH) / LM_BERAT, "#,##0.00")
                    
                End If
                
            End If
        
        'Else
        
            If Frm84.TB8 <> vbNullString And IsNumeric(Frm84.TB8) Then LM_HARGA_LEPAS_DISKAUN = Frm84.TB8
            If Frm84.L8_Text <> vbNullString And IsNumeric(Frm84.L8_Text) Then frm84_LM_KADAR_GST = Frm84.L8_Text
            
            If Frm84.CB2 = 1 Or Frm84.CB18 = 1 Then
                Frm84.TB9 = Format(LM_HARGA_LEPAS_DISKAUN - LM_HARGA_JUALAN, "#,##0.00") 'Adjustment
            End If
            
            If Frm84.CB3 = 1 Then
            
                'LM_HARGA_BEFORE_GST = Format(LM_HARGA_JUALAN - (LM_HARGA_JUALAN / (1 + (frm84_LM_KADAR_GST / 100))), "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
                LM_HARGA_BEFORE_GST = Format((LM_HARGA_JUALAN / (1 + (frm84_LM_KADAR_GST / 100))), "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
                Frm84.TB9 = Format(LM_HARGA_LEPAS_DISKAUN - LM_HARGA_BEFORE_GST, "#,##0.00") 'Adjustment
            End If
            
        'End If
        
        Call Frm84_pengiraan_komisyen_dropship
        Call Frm84_modal_dan_jual
        
        G_CALC_AUTO = 0
    
    Else
        MsgBox "Hanya NOMBOR dibenarkan di dalam ruangan ini.", vbExclamation, "Info"
    End If
    
End If
End Sub

Private Sub CMD4_Click()
'on error resume next
Frm84.Pic3.Visible = False
Frm84.Frame1.Visible = True
Frm84.TB1.SetFocus
End Sub
Private Sub CMD5_Click()
'On Error Resume Next
Call tesuto2
End Sub
Private Sub CMD7_Click()
'on error resume next
DATA_FOUND = 0
Frm84_LM_No_RUJUKAN_PEMBELI = vbNullString
Frm_LM_DATA_PENJUAL_BUYBACK = 0 '0 : Penjual Tidak Berdaftar , 1 : Penjual Berdaftar , 2 : Ahli
Frm84_LM_NAMA = "-"
Frm84_LM_HP = "-"
Frm84_LM_IC = "-"

If Frm84.TB18 = vbNullString Then
    MsgBox "Sila Masukkan No. Resit Buyback/Trade In.", vbInformation, "Info"
    Exit Sub
End If

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & UCase(Frm84.TB18) & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!trade_in_status) Then
        If rs!trade_in_status = 0 Then
            If Not IsNull(rs!jumlah_tanpa_gst) Then
                Frm84.TB17 = Format(rs!jumlah_tanpa_gst, "0.00") 'Jumlah Nilaian Resit Trade In
                'Frm84.L22_Text = Format(rs!jumlah_tanpa_gst, "0.00") 'Jumlah Nilaian Resit Trade In
                Frm84.L58_Text = Format(rs!jumlah_tanpa_gst, "0.00") 'Jumlah Nilaian Resit Trade In
            End If
            If Not IsNull(rs!no_rujukan_pelanggan_buyback) Then Frm84_LM_No_RUJUKAN_PEMBELI = rs!no_rujukan_pelanggan_buyback 'No. Rujukan Penjual Barang Trade In
            Frm84.L16_Text = UCase(Frm84.TB18) 'No. Resit Trade In
            Frm84.L57_Text = UCase(Frm84.TB18) 'No. Resit Trade In
            Frm84.TB18 = vbNullString

            If Not IsNull(rs!kategori_penjual) Then
                If rs!kategori_penjual = 0 Then
                
                    If Not IsNull(rs!no_rujukan_pelanggan_buyback) Then
                        Frm85_LM_No_PENJUAL = rs!no_rujukan_pelanggan_buyback 'No. Rujukan Penjual (Penjual Buyback)
                        Frm_LM_DATA_PENJUAL_BUYBACK = 1 '0 : Penjual Tidak Berdaftar , 1 : Penjual Berdaftar , 2 : Ahli
                    Else
                        Frm_LM_DATA_PENJUAL_BUYBACK = 0 '0 : Penjual Tidak Berdaftar , 1 : Penjual Berdaftar , 2 : Ahli
                    End If
                
                ElseIf rs!kategori_penjual = 1 Then
                
                    If Not IsNull(rs!no_rujukan_pelanggan_buyback) Then
                        Frm85_LM_No_PENJUAL = rs!no_rujukan_pelanggan_buyback 'No. Rujukan Penjual (Penjual Buyback)
                        Frm_LM_DATA_PENJUAL_BUYBACK = 2 '0 : Penjual Tidak Berdaftar , 1 : Penjual Berdaftar , 2 : Ahli
                    End If
                    
                End If
            End If
            
            DATA_FOUND = 1

        ElseIf rs!trade_in_status = 1 Then
            MsgBox "No. Resit Trade In Ini Telah Digunakan Untuk Urusan Belian Sebelum Ini.", vbInformation, "Info"
            
            Frm84.TB18 = vbNullString
            Frm84.TB18.SetFocus
        End If
    End If
Else
    MsgBox "No. Resit Tidak Dijumpai.", vbInformation, "Info"
    
    Frm84.TB18 = vbNullString
    Frm84.TB18.SetFocus
End If

rs.Close
Set rs = Nothing

Exit Sub

'### Carian Maklumat Penjual Bagi Buyback ### - Start
If Frm_LM_DATA_PENJUAL_BUYBACK = 0 Then '0 : Penjual Tidak Berdaftar , 1 : Penjual Berdaftar , 2 : Ahli

    If Frm84.CB4 = 1 Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 44_senarai_pelanggan where no_resit='" & Frm83.L12_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!Nama) Then Frm84_LM_NAMA = rs!Nama 'Nama
            If Not IsNull(rs!no_tel) Then Frm84_LM_HP = rs!no_tel 'No. Telefon
            
            Note = "Adakah anda ingin menggunakan maklumat pembeli urusan pembelian barang ini mengikut maklumat pembeli yang tercatat dalam payment voucher trade in ini?" & vbCrLf & _
                    vbNullString & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Nama : " & Frm84_LM_NAMA & vbCrLf & _
                    "No. Telefon : " & Frm84_LM_HP & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Jika ya sila klik [Yes] dan [No] jika tidak."
                    
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbNo Then
                
            ElseIf Answer = vbYes Then
            
                Call Frm26_initial
                Call Frm27_initial
                Call Frm28_initial
            
                If Not IsNull(rs!Nama) Then
                    Frm26.TB1 = rs!Nama 'Nama
                    Frm83.L36_Text = rs!Nama 'Nama
                End If
                If Not IsNull(rs!no_tel) Then Frm26.TB2 = rs!no_tel 'No. Telefon
                
            End If
    
        End If
        
        rs.Close
        Set rs = Nothing
        
    End If
    
ElseIf Frm_LM_DATA_PENJUAL_BUYBACK = 1 Then '0 : Penjual Tidak Berdaftar , 1 : Penjual Berdaftar , 2 : Ahli
    
    If Frm84.CB4 = 1 Then

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm85_LM_No_PENJUAL & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!Nama) Then Frm84_LM_NAMA = rs!Nama 'Nama
            If Not IsNull(rs!no_ic) Then Frm84_LM_IC = rs!no_ic 'No. Kad Pengenalan
            If Not IsNull(rs!no_tel) Then Frm84_LM_HP = rs!no_tel 'No. Telefon
        
            Note = "Adakah anda ingin menggunakan maklumat pembeli urusan pembelian barang ini mengikut maklumat pembeli yang tercatat dalam payment voucher trade in ini?" & vbCrLf & _
                    vbNullString & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Nama : " & Frm84_LM_NAMA & vbCrLf & _
                    "No. Kad Pengenalan : " & Frm84_LM_IC & vbCrLf & _
                    "No. Telefon : " & Frm84_LM_HP & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Jika ya sila klik [Yes] dan [No] jika tidak."
                    
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbNo Then
                
            ElseIf Answer = vbYes Then
            
                Call Frm26_initial
                Call Frm27_initial
                Call Frm28_initial
        
                If Not IsNull(rs!Nama) Then
                    Frm27.L1_Text = rs!Nama 'Nama
                    Frm83.L37_Text = rs!Nama 'Nama
                End If
                If Not IsNull(rs!no_ic) Then Frm27.L2_Text = rs!no_ic 'No. Kad Pengenalan
                If Not IsNull(rs!no_tel) Then Frm27.L3_Text = rs!no_tel 'No. Telefon
                If Not IsNull(rs!Email) Then Frm27.L4_Text = rs!Email 'E-mail
                If Not IsNull(rs!no_pelanggan) Then Frm27.L5_Text = rs!no_pelanggan 'No. Pelanggan
                
            End If
    
        End If
        
        rs.Close
        Set rs = Nothing
        
    End If
    
ElseIf Frm_LM_DATA_PENJUAL_BUYBACK = 2 Then '0 : Penjual Tidak Berdaftar , 1 : Penjual Berdaftar , 2 : Ahli

    If Frm84.CB5 = 1 Or Frm84.CB6 = 1 Or Frm84.CB9 = 1 Or Frm84.CB10 = 1 Then

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm85_LM_No_PENJUAL & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            If Not IsNull(rs!Nama) Then Frm84_LM_NAMA = rs!Nama 'Nama
            If Not IsNull(rs!no_ic) Then Frm84_LM_IC = rs!no_ic 'No. Kad Pengenalan
            If Not IsNull(rs!no_tel) Then Frm84_LM_HP = rs!no_tel 'No. Telefon
        
            Note = "Adakah anda ingin menggunakan maklumat pembeli urusan pembelian barang ini mengikut maklumat pembeli yang tercatat dalam payment voucher trade in ini?" & vbCrLf & _
                    vbNullString & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Nama : " & Frm84_LM_NAMA & vbCrLf & _
                    "No. Kad Pengenalan : " & Frm84_LM_IC & vbCrLf & _
                    "No. Telefon : " & Frm84_LM_HP & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Jika ya sila klik [Yes] dan [No] jika tidak."
                    
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbNo Then
                
            ElseIf Answer = vbYes Then
            
                Call Frm26_initial
                Call Frm27_initial
                Call Frm28_initial
        
                If Not IsNull(rs!Nama) Then
                    Frm28.L1_Text = rs!Nama 'Nama
                    Frm83.L37_Text = rs!Nama 'Nama
                End If
                If Not IsNull(rs!no_ic) Then Frm28.L2_Text = rs!no_ic 'No. Kad Pengenalan
                If Not IsNull(rs!no_tel) Then Frm28.L3_Text = rs!no_tel 'No. Telefon
                If Not IsNull(rs!Email) Then Frm28.L4_Text = rs!Email 'E-mail
                If Not IsNull(rs!no_pelanggan) Then Frm28.L5_Text = rs!no_pelanggan 'No. Pelanggan
                
            End If
    
        End If
        
        rs.Close
        Set rs = Nothing
        
    End If
    
End If
'### Carian Maklumat Penjual Bagi Buyback ### - End
End Sub
Private Sub CMD8_Click()
'on error resume next
If Frm84.L16_Text <> vbNullString Then
    Note = "Adakah Anda Ingin Batalkan No. Resit Ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
        Call Frm26_initial
        Call Frm27_initial
        Call Frm28_initial

        Frm84.L16_Text = vbNullString
        Frm84.TB17 = "0.00"
        Frm84.L57_Text = vbNullString 'No. Voucher
        Frm84.L58_Text = vbNullString 'Jumlah trade in
        
        'Frm84.L27_Text = vbNullString
        'Frm84.L28_Text = vbNullString
    End If
Else
    MsgBox "Tiada Maklumat Tentang Resit Trade In.", vbInformation, "Info"
End If
End Sub



Private Sub CMD9_Click()
'on error resume next
If Frm84.TB49 = vbNullString Or (Frm84.TB49 <> vbNullString And Not IsNumeric(Frm84.TB49)) Then
    MsgBox "[Berat] tidak sah.", vbExclamation, "Info"
    Exit Sub
End If
If Frm84.TB50 = vbNullString Or (Frm84.TB50 <> vbNullString And Not IsNumeric(Frm84.TB50)) Then
    MsgBox "[Harga Semasa Trade In] tidak sah.", vbExclamation, "Info"
    Exit Sub
End If
If Frm84.TB51 = vbNullString Or (Frm84.TB51 <> vbNullString And Not IsNumeric(Frm84.TB51)) Then
    MsgBox "[Harga Semasa Buyback] tidak sah.", vbExclamation, "Info"
    Exit Sub
End If
If Frm84.TB52 = vbNullString Or (Frm84.TB52 <> vbNullString And Not IsNumeric(Frm84.TB52)) Then
    MsgBox "[Caj Pertukaran] tidak sah.", vbExclamation, "Info"
    Exit Sub
End If

G_TI_BERAT = Frm84.TB49
G_TI_TRADE_IN = Frm84.TB50
G_TI_BUYBACK = Frm84.TB51
G_TI_CAJ = Frm84.TB52

G_TI_MODE = 3

Call frm_kiraan_harga_selepas_ti
End Sub

Private Sub Command1_Click()
'on error resume next
If Frm84.L56_Text <> 0 Then

    Note = "Adakah anda yakin untuk batalkan urusan ini ?" & vbCrLf & _
            "Semua data trade in yang telah dimasukkan akan dipadamkan jika anda teruskan." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
        'Unload Frm83
    
        Frm84.L56_Text = 0 '0 : Tiada , 1 : Belian dengan trade in (sudah ada voucher) , 2 : Trade in
        Frm84.L57_Text = vbNullString 'No. Voucher
        Frm84.L58_Text = vbNullString 'Jumlah trade in
        
        Frm84.L16_Text = vbNullString 'No. Voucher
        Frm84.TB17 = "0.00" 'Jumlah trade in
    
        Frm84.Frame6.Visible = False
        
    End If
    
End If
End Sub
Private Sub Form_Load()
'on error resume next
With Frm84.ListView1
    Set .SmallIcons = Frm84.ImageList1
    Set .Icons = Frm84.ImageList1
    'For Sales
    .ListItems.Add , "Scan", "Scan Item", 1
    .ListItems.Add , "Senarai", "Senarai", 8
    '.ListItems.Add , "Pembeli_1", "Pembeli (Berdaftar)", 2
    '.ListItems.Add , "Pembeli_2", "Pembeli (Tidak Berdaftar)", 6
    '.ListItems.Add , "dropship", "Pembeli (Tidak Berdaftar)", 7
    .ListItems.Add , "Bayaran", "Bayaran", 3
    .ListItems.Add , "Invoice", "Invoice", 4
    .ListItems.Add , "Trade In", "Trade In", 5
    .ListItems.Add , "Trade In (0%)", "Trade In (0%)", 5
End With

GLOBAL_DISABLE = 0
Frm84.L41_Text = 0 '0 : Data Baru , 1:  Data Diedit
frm130.L41_Text = 0

Frm84.CMD2.Visible = True
Frm84.CMD5.Visible = True
Frm84.L3_Text.BackStyle = 0
Frm84.L4_Text.BackStyle = 0
Frm84.L8_Text.BackStyle = 0
'Frm84.L31_Text.BackStyle = 0
'Frm84.L32_Text.BackStyle = 0
'Frm84.L81_Text.BackStyle = 0
'Frm84.L82_Text.BackStyle = 0

Frm84.DTPicker1 = DateTime.Date

Frm84.CB4 = 0
Frm84.CB5 = 0
Frm84.CB6 = 0
Frm84.CB9 = 0
Frm84.CB10 = 0
Frm84.CB7 = 0
End Sub

Private Sub Frm84_scan_mode_Click()
'on error resume next
If Frm84.Pic3.Visible = True Or Frm84.Frame4.Visible = True Or Frm84.Frame6.Visible = True Or Frm26.Visible = True Or Frm27.Visible = True Or Frm28.Visible = True Then
    If Frm84.Pic3.Visible = True Then
        msg = "Sila tutup dahulu paparan [Maklumat Terperinci Bayaran] sebelum scan barang yang hendak dijual."
    End If
    If Frm84.Frame4.Visible = True Then
        msg = "Sila tutup dahulu paparan [Senarai Jualan] sebelum scan barang yang hendak dijual."
    End If
    If Frm84.Frame6.Visible = True Then
        msg = "Sila tutup dahulu paparan [Trade In] sebelum scan barang yang hendak dijual."
    End If
    If Frm27.Visible = True Then
        msg = "Sila tutup dahulu paparan [Maklumat Agen Dropship] sebelum scan barang yang hendak dijual."
    End If
    If Frm26.Visible = True Or Frm28.Visible = True Then
        msg = "Sila tutup dahulu paparan [Maklumat Pembeli] sebelum scan barang yang hendak dijual."
    End If
    
    MsgBox msg, vbInformation, "Info"
Else
    Frm84.TB1 = vbNullString
    Frm84.TB1.SetFocus
End If
End Sub
Private Sub Frm84_SM_Edit_Click()
'on error resume next
DATA_FOUND = 0
Frm84_LM_BARANG_PERMATA = 0
Frm84_LM_ID = vbNullString
    
If IsNumeric(Frm84.ListView2.SelectedItem.Index) Then
    
    Frm84_LM_ID = Frm84.ListView2.ListItems(Frm84.ListView2.SelectedItem.Index)
    
    If Frm84_LM_ID <> vbNullString Then
    
        Call Frm84_Reset_Edit '!! Hati-hati dengan tempat letakkan command ini!!
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_JUALAN_TEMP & " where ID='" & Frm84_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!ID) Then 'No. ID
                Frm84.L39_Text = rs!ID
            Else
                Frm84.L39_Text = vbNullString
            End If
            If Not IsNull(rs!flag_barang) Then '0 : Stok kedai , 1 : Barang trade in/potong
                If rs!flag_barang = 0 Then
                    Frm84.L83_Text = "0"
                ElseIf rs!flag_barang = 1 Then
                    Frm84.L83_Text = "1"
                    LM_FLAG_BARANG = 1
                End If
            End If
            If Not IsNull(rs!nama_purity) Then LM_NAMA_PURITY = rs!nama_purity
            
            If Not IsNull(rs!no_siri_Produk) Then 'No. Siri Produk
                Frm84.TB2 = rs!no_siri_Produk
            Else
                Frm84.TB2 = vbNullString
            End If
            If Not IsNull(rs!kategori_Produk) Then 'Kategori Produk
                Frm84.L12_Text = rs!kategori_Produk
            Else
                Frm84.L12_Text = vbNullString
            End If
            If Not IsNull(rs!purity) Then 'Purity
                Frm84.L13_Text = rs!purity
            Else
                Frm84.L13_Text = vbNullString
            End If
            If Not IsNull(rs!Berat_Asal) Then 'Berat Asal (g)
                Frm84.TB3 = rs!Berat_Asal
            Else
                Frm84.TB3 = vbNullString
            End If
            If Not IsNull(rs!berat_jualan) Then 'Berat Jualan (g)
                Frm84.TB4 = rs!berat_jualan
            Else
                Frm84.TB4 = vbNullString
            End If
            If Not IsNull(rs!harga_Semasa) Then 'Harga Semasa (RM/g)
                Frm84.TB5 = rs!harga_Semasa
            Else
                Frm84.TB5 = vbNullString
            End If
            If Not IsNull(rs!UPAH) Then 'Upah (RM)
                Frm84.TB15 = rs!UPAH
            Else
                Frm84.TB15 = vbNullString
            End If
            If Not IsNull(rs!harga_asal) Then 'Harga Asal Item (RM)
                Frm84.TB6 = rs!harga_asal
            Else
                Frm84.TB6 = vbNullString
            End If
            If Not IsNull(rs!diskaun) Then 'Diskaun (%)
                Frm84.TB7 = rs!diskaun
            Else
                Frm84.TB7 = vbNullString
            End If
            If Not IsNull(rs!harga_lepas_diskaun) Then 'Harga Selepas Diskaun (RM)
                Frm84.TB8 = rs!harga_lepas_diskaun
            Else
                Frm84.TB8 = vbNullString
            End If
            If Not IsNull(rs!adjustment) Then 'Adjustment (RM)
                Frm84.TB9 = rs!adjustment
            Else
                Frm84.TB9 = vbNullString
            End If
            If Not IsNull(rs!harga_jualan) Then 'Harga Jualan (RM)
                Frm84.TB10 = rs!harga_jualan
            Else
                Frm84.TB10 = vbNullString
            End If
            If Not IsNull(rs!gst_ari_nashi) Then 'Harga Jualan (RM)
                If rs!gst_ari_nashi = "ZR (L)" Then
                    Frm84.CB2 = 1 '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                    If Not IsNull(rs!jumlah_gst) Then
                        Frm84.TB11 = rs!jumlah_gst 'Jumlah Cukai GST (RM)
                    Else
                        Frm84.TB11 = vbNullString
                    End If
                ElseIf rs!gst_ari_nashi = "SR" Then
                    Frm84.CB3 = 1 '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                    If Not IsNull(rs!kadar_gst) Then
                        Frm84.L8_Text = rs!kadar_gst 'Kadar Cukai GST (%)
                        'frm130.L8_Text = rs!kadar_gst 'Kadar Cukai GST (%)
                    End If
                    If Not IsNull(rs!jumlah_gst) Then
                        Frm84.TB11 = rs!jumlah_gst 'Jumlah Cukai GST (RM)
                    Else
                        Frm84.TB11 = vbNullString
                    End If
                    If Not IsNull(rs!gst_include) Then
                        If rs!gst_include = "1" Then
                            Frm84.CB18 = 1
                            Frm84.CB3 = 0
                        End If
                    Else
                        Frm84.CB18 = 0
                    End If
                End If
            End If
            If Not IsNull(rs!harga_dengan_gst) Then 'Harga Jualan Termasuk GST (RM)
                Frm84.TB14 = rs!harga_dengan_gst
            Else
                Frm84.TB14 = vbNullString
            End If
            If Not IsNull(rs!dropship) Then '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
                If rs!dropship = 0 Then
                    Frm84.CB7 = 0
                    Frm84.Frame2.Visible = False
                    Frm84.Frame3.Visible = False
                ElseIf rs!dropship = 1 Then
                    Frm84.CB7 = 1
                    If Not IsNull(rs!Type) Then
                        If rs!Type = 0 Then '0 : BK , 1 : Barang Permata
                            If Not IsNull(rs!komisyen_per_gram) Then 'Komisyen Per Gram Dropship (RM/g) : BK
                                Frm84.TB12 = rs!komisyen_per_gram
                            Else
                                Frm84.TB12 = vbNullString
                            End If
                            If Not IsNull(rs!jumlah_komisyen) Then 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
                                Frm84.TB13 = rs!jumlah_komisyen
                            Else
                                Frm84.TB13 = vbNullString
                            End If
                            If Not IsNull(rs!kadar_komisyen_upah) Then 'Kadar komisyen bagi upah kepada agen dropship
                                Frm84.TB43 = rs!kadar_komisyen_upah
                            Else
                                Frm84.TB43 = vbNullString
                            End If
                            If Not IsNull(rs!komisyen_upah) Then 'Jumlah komisyen bagi upah kepada agen dropship
                                Frm84.TB44 = rs!komisyen_upah
                            Else
                                Frm84.TB44 = vbNullString
                            End If
                            
                            Frm84.Frame2.Visible = True
                        ElseIf rs!Type = 1 Then
                            If Not IsNull(rs!jumlah_komisyen) Then 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : Permata
                                Frm84.TB16 = rs!jumlah_komisyen
                            Else
                                Frm84.TB16 = vbNullString
                            End If
                            
                            Frm84.Frame3.Visible = True
                        End If
                    End If
                End If
            End If
            If Not IsNull(rs!Type) Then
                If rs!Type = 0 Then '0 : BK , 1 : Barang Permata
                
                    Frm84.L68_Text = "Modal (RM/g) :                   Jual (RM/g) :"
                    
                    If Not IsNull(rs!harga_per_gram_modal) Then 'Harga Per Gram Bagi Modal (RM/g)
                        Frm84.L34_Text = rs!harga_per_gram_modal
                    Else
                        Frm84.L34_Text = vbNullString
                    End If
                    If Not IsNull(rs!harga_per_gram_tanpa_gst) Then 'Harga modal per gram tanpa GST (RM)
                        Frm84.L42_Text = rs!harga_per_gram_tanpa_gst
                    Else
                        Frm84.L42_Text = vbNullString
                    End If
                    If Not IsNull(rs!harga_per_gram_supplier) Then 'Harga per gram (harga semasa) dari supplier (modal)
                        Frm84.L54_Text = rs!harga_per_gram_supplier
                    Else
                        Frm84.L54_Text = "0.00"
                    End If
                    If Not IsNull(rs!upah_modal) Then 'Upah modal
                        Frm84.L55_Text = rs!upah_modal
                    Else
                        Frm84.L55_Text = "0.00"
                    End If
                    
                    Frm84.TB5.Locked = False
                    Frm84.TB4.Locked = False
                    Frm84.TB15.Locked = False
                    Frm84.TB22.Locked = False
                    Frm84.TB6.Locked = True
                    'insan
                    Frm84.TB5.BackColor = &HFFFFFF
                    Frm84.TB4.BackColor = &HFFFFFF
                    Frm84.TB15.BackColor = &HFFFFFF
                    Frm84.TB22.BackColor = &HFFFFFF
                    Frm84.TB6.BackColor = &H8000000A
                    Frm84_LM_BARANG_PERMATA = 1
                    
                ElseIf rs!Type = 1 Then '0 : BK , 1 : Barang Permata
                    
                    Frm84.L68_Text = "Modal (RM)   :                      Jual (RM) :"
                    
                    If Not IsNull(rs!modal) Then 'Harga Modal (RM)
                        Frm84.L34_Text = rs!modal
                    Else
                        Frm84.L34_Text = vbNullString
                    End If
                    If Not IsNull(rs!modal_tanpa_gst) Then 'Harga Modal Tanpa GST (RM)
                        Frm84.L42_Text = rs!modal_tanpa_gst
                    Else
                        Frm84.L42_Text = vbNullString
                    End If

                    Frm84.L54_Text = vbNullString
                    
                    Frm84.TB5.Locked = True
                    Frm84.TB4.Locked = True
                    Frm84.TB15.Locked = True
                    Frm84.TB22.Locked = True
                    Frm84.TB6.Locked = False
                    
                    Frm84.TB5.BackColor = &H8000000A
                    Frm84.TB4.BackColor = &H8000000A
                    Frm84.TB15.BackColor = &H8000000A
                    Frm84.TB22.BackColor = &H8000000A
                    Frm84.TB6.BackColor = &HFFFFFF
                    
                End If
            End If
            
'### Maklumat tetapan harga jualan kepada staff ### - Start
            If Not IsNull(rs!kadar_penurunan_upah) Then Frm84.L48_Text = rs!kadar_penurunan_upah 'Kadar peratusan penurunan harga upah kepada staff (%)
            If Not IsNull(rs!harga_semasa_staff) Then 'Harga emas semasa yang dijual kepada staff
                Frm84.L49_Text = rs!harga_semasa_staff
            Else
                Frm84.L49_Text = vbNullString
            End If
            If Not IsNull(rs!kadar_penurunan_bp) Then Frm84.L50_Text = rs!kadar_penurunan_bp 'Kadar peratusan penurunan harga barang permata kepada staff (%)
            If Not IsNull(rs!harga_staff) Then 'Harga yang dijual kepada staff (RM)
                Frm84.L51_Text = rs!harga_staff
            Else
                Frm84.L51_Text = vbNullString
            End If
            If Not IsNull(rs!harga_bp_asal) Then 'Tetapan harga barang permata yang asal (RM)
                Frm84.L52_Text = rs!harga_bp_asal
            Else
                Frm84.L52_Text = vbNullString
            End If
            If Not IsNull(rs!upah_asal) Then 'Tetapan upah asal (RM)
                Frm84.L53_Text = rs!upah_asal
            Else
                Frm84.L53_Text = vbNullString
            End If
            If Not IsNull(rs!upah_asal) Then 'Tetapan upah asal (RM)
                Frm84.L53_Text = rs!upah_asal
            Else
                Frm84.L53_Text = vbNullString
            End If
'### Maklumat tetapan harga jualan kepada staff ### - End
            
            If Not IsNull(rs!gst_barang_atau_upah) Then '0 : GST pada harga jualan , 1 : GST pada upah
                If rs!gst_barang_atau_upah = 0 Then
                    Frm84.CB12 = 0
                ElseIf rs!gst_barang_atau_upah = 1 Then
                    Frm84.CB12 = 1
                End If
            Else
                Frm84.CB12 = 0
            End If
            
            If Not IsNull(rs!Type) Then '0 : BK , 1 : Barang Permata
                
                If rs!Type = 0 Then
                
                    Frm84.L68_Text = "Modal (RM/g) :                   Jual (RM/g) :"
                    
                ElseIf rs!Type = 1 Then
                    
                    Frm84.L68_Text = "Modal (RM)   :                      Jual (RM) :"
                
                End If
                
            Else
            
                Frm84.L68_Text = "Modal (RM/g) :                   Jual (RM/g) :"
            
            End If
            
            If Not IsNull(rs!jualan_per_gram) Then 'Purata harga jualan per gram (RM/g) bagi barang kemas , Bagi barang permata adalah merujuk kepada harga jualan
                
                Frm84.L67_Text = rs!jualan_per_gram

            Else
            
                Frm84.L67_Text = "0.00"
            
            End If
            
            If Not IsNull(rs!modal_per_gram) Then 'Paparan modal per gram (tanpa GST)
                
                Frm84.L69_Text = rs!modal_per_gram
                
            Else
            
                Frm84.L69_Text = "0.00"
            
            End If
            
            If Not IsNull(rs!flag_upah) Then
                
                If rs!flag_upah = 0 Then
                    
                    Frm84.L70_Text = 0
                    
                    If Frm84_LM_BARANG_PERMATA = 1 Then
                        Frm84.TB22.Locked = False
                        Frm84.TB22.BackColor = &HFFFFFF
                        Frm84.TB15.Locked = True
                        Frm84.TB15.BackColor = &H8000000A
                    End If
                    
                ElseIf rs!flag_upah = 1 Then
                    
                    Frm84.L70_Text = 1
                    
                    If Frm84_LM_BARANG_PERMATA = 1 Then
                        Frm84.TB15.Locked = False
                        Frm84.TB15.BackColor = &HFFFFFF
                        Frm84.TB22.Locked = True
                        Frm84.TB22.BackColor = &H8000000A
                    End If
                    
                End If
                
            End If
            
            If Not IsNull(rs!upah_per_gram) Then
                
                Frm84.TB22 = Format(rs!upah_per_gram, "0.00")
                
            Else
            
                Frm84.TB22 = vbNullString
            
            End If

            DATA_FOUND = 1
        End If
        
        rs.Close
        Set rs = Nothing
       
        If DATA_FOUND = 1 Then
        
            GLOBAL_DISABLE = 1
            
            If LM_FLAG_BARANG = 1 Then
            
                If LM_NAMA_PURITY <> vbNullString Then

                    On Error GoTo Err_A:
                    Frm84.CBB4 = LM_NAMA_PURITY
Restore_A:

                End If
                
                If Frm84.L12_Text <> vbNullString Then
                    
                    On Error GoTo Err_B:
                    Frm84.CBB3 = Frm84.L12_Text

Restore_B:
                End If
                
                If Frm84.CBB4 <> vbNullString Then
                    Call frm84_call_edit_berat
                    Call frm84_berat_guna_dr_invoice_ini
                End If
                
            End If
            
            GLOBAL_DISABLE = 0
            
            Frm84.CMD3.Visible = False
            Frm84.CMD13.Visible = True
            Frm84.CMD14.Visible = True
            
            Frm84.TB1 = vbNullString
            Frm84.TB1.Locked = True
            Frm84.TB1.BackColor = &H8000000A
            
            Frm84.Pic3.Visible = False
            Frm84.Frame1.Visible = True
            
        End If
    End If
End If

Exit Sub
Err_A:
Frm84.CBB4.AddItem LM_NAMA_PURITY
Frm84.CBB4 = LM_NAMA_PURITY
Resume Restore_A:

Exit Sub
Err_B:
Frm84.CBB3.AddItem Frm84.L12_Text
Frm84.CBB3 = Frm84.L12_Text
Resume Restore_B:
End Sub
Private Sub Frm84_SM_Padam_Click()
'on error resume next
DATA_FOUND = 0

If IsNumeric(Frm84.ListView2.SelectedItem.Index) Then
    
    Frm84_LM_ID = Frm84.ListView2.ListItems(Frm84.ListView2.SelectedItem.Index)
    
    If Frm84_LM_ID <> vbNullString Then
        
        Note = "Adakah anda ingin keluarkan item ini dari senarai ini ?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        If Answer = vbNo Then
            'Exit Sub
        End If
        If Answer = vbYes Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from " & G_JUALAN_TEMP & " where ID='" & Frm84_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!Status) Then
'### Kod Bagi Status ###
'==========================================
'0 : Tidak Aktif (Tidak Perlu Buat Apa-Apa)
'1 : Aktif - Kemasukkan Data Baru
'2 : Barang Sudah Terjual ****Tidak Dibenarkan Untuk Diedit Atau Dipadamkan
'3 : Kemasukkan Data Baru
'4 : Data Diedit
'5 : Keluarkan Data Dari Database (Bagi Item Asal Yang Dipadamkan)
'6 : Ignore Kemasukkan Data Ke Dalam Database

                    If rs!Status = "1" Then
                        If Frm84.L41_Text = 0 Then
                            Frm84_LM_STATUS = "0"
                            DATA_FOUND = 1
                        ElseIf Frm84.L41_Text = 1 Then
                            Frm84_LM_STATUS = "5"
                            DATA_FOUND = 1
                        End If
                    ElseIf rs!Status = "4" Then
                        Frm84_LM_STATUS = "5"
                        DATA_FOUND = 1
                    ElseIf rs!Status = "3" Then
                        Frm84_LM_STATUS = "6"
                        DATA_FOUND = 1
                    End If
                    
                    If DATA_FOUND = 1 Then
                        rs!Status = Frm84_LM_STATUS
                        rs.Update
                    End If
                End If
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_FOUND = 1 Then
    
                GM_NEXT_PREV = 2 '0 : Next , 1 : Previous
                        
                Call Frm84_Senarai_Jualan_Header
                Call Frm84_Senarai_Jualan
                
                MsgBox "Item Telah Dikeluarkan Dari Senarai.", vbInformation, "Info"
            End If
        End If
        
    End If
End If
End Sub
Private Sub Frm84_SM_reset_Click()
'on error resume next
Note = "Adakah anda ingin reset semua data jualan ini ?" & vbCrLf & _
        vbNullString & vbCrLf & _
        "Semua data jualan yang telah discan , maklumat pembeli dan data berkaitan dengan jualan ini akan dipadamkan." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    'Exit Sub
End If
If Answer = vbYes Then
    Call Frm84_Load_Form
    Call Frm84_Reset_Edit

    Unload Frm26
    Unload Frm27
    Unload Frm28
    Unload Frm83

    'Call Frm26_initial
    'Call Frm27_initial
    'Call Frm28_initial
    
    'Frm84.CMD3.Visible = True
    'Frm84.CMD13.Visible = False
    'Frm84.CMD14.Visible = False
    
    Frm84.Frame1.Visible = True
    Frm84.TB1 = vbNullString
    Frm84.TB1.Locked = False
    Frm84.TB1.BackColor = &HFFFFFF
    
    Frm84.TB1.SetFocus
End If
End Sub
Private Sub Frm84_SM_tukar_kategori_Click()
'on error resume next
Frm84.L64_Text = Frm84.L45_Text
If Frm84.L62_Text = "Jualan oleh agen dropship : YA" Then
'If Frm84.CMD11.Enabled = True Then
    Frm84.L65_Text = 1
Else
    Frm84.L65_Text = 0
End If

Call frm84_disable_frame

Frm84.Pic6.Visible = True
Frm84.Pic3.Visible = False
End Sub

Private Sub L1_Text_Click()

End Sub
Private Sub L15_Text_Change()
'on error resume next
Call Frm84_kiraan_potongan_kupon
End Sub
Private Sub L17_Text_Change()
'on error resume next
Call Frm84_kira_harga_layak_mata
End Sub



Private Sub L19_Text_Change()
'on error resume next
Call frm84_harga_selepas_diskaun
End Sub
Private Sub L2_Text_Change()
'On Error Resume Next
If Frm84.CMD13.Visible = True Then
    If Frm84.L40_Text.Visible = True Then
        Frm84.L40_Text.Visible = False
    Else
        Frm84.L40_Text.Visible = True
    End If
Else
    Frm84.L40_Text.Visible = False
End If
End Sub
Private Sub L20_Text_Change()
'on error resume next
Dim Frm84_HARGA_LEPAS_DISKAUN As Double
Dim Frm84_ADJUSTMENT As Double

Call Frm84_pengiraan_harga_jualan
End Sub
Private Sub L21_Text_Change()
'on error resume next
Call frm_kiraan_harga_selepas_ti
End Sub
Private Sub L22_Text_Change()
'on error resume next
Call frm_kiraan_harga_selepas_ti

Exit Sub

Dim Frm84_LM_HARGA As Double
Dim Frm84_LM_BUYBACK As Double
Dim Frm84_LM_HARGA_LAST As Double
Dim Frm84_LM_DEDUCT_RESIT As Double

If GLOBAL_DISABLE = 0 Then

    Frm84_LM_HARGA = 0
    Frm84_LM_BUYBACK = 0
    Frm84_LM_HARGA_LAST = 0
    Frm84_LM_TOLAKAN_RESIT = 0
    Frm84_LM_DEDUCT_RESIT = 0
    
    If ((Frm84.L21_Text <> vbNullString And IsNumeric(Frm84.L21_Text)) And (Frm84.L22_Text <> vbNullString And IsNumeric(Frm84.L22_Text))) Then
        Frm84_LM_HARGA = Frm84.L21_Text 'Harga Barang
        Frm84_LM_BUYBACK = Frm84.L22_Text 'Buyback
        
        Frm84_LM_HARGA_LAST = Frm84_LM_HARGA - Frm84_LM_BUYBACK
        
        If Frm84_LM_HARGA_LAST >= 0 Then
            Frm84.L23_Text = Format(Frm84_LM_HARGA_LAST, "#,##0.00") 'Harga Perlu Bayar
            Frm84.TB33 = Format(Frm84_LM_HARGA_LAST, "#,##0.00") 'Harga Perlu Bayar
            Frm84.L37_Text = "0.00"
            Frm84.L24_Text = "Jumlah Bayaran"
            Frm84.L25_Text = "Jumlah Bayaran"
        Else
            Frm84.L23_Text = -Format(Frm84_LM_HARGA_LAST, "#,##0.00")  'Harga Perlu Bayar
            Frm84.TB33 = -Format(Frm84_LM_HARGA_LAST, "#,##0.00")  'Harga Perlu Bayar
            Frm84.L24_Text = "Harga Kedai Perlu Bayar Pelanggan"
            Frm84.L25_Text = "Harga Kedai Perlu Bayar Pelanggan"
            Frm84_LM_TOLAKAN_RESIT = 1
        End If
        
    Else
        Frm84.L23_Text = "0.00" 'Harga Perlu Bayar
        Frm84.TB33 = "0.00" 'Harga Perlu Bayar
        Frm84.L37_Text = "0.00"
        Frm84.L24_Text = "Jumlah Bayaran"
        Frm84.L25_Text = "Jumlah Bayaran"
    End If
    
    If Frm84_LM_TOLAKAN_RESIT = 1 Then
        If IsNumeric(Frm84.L38_Text) Then
            Frm84_LM_DEDUCT_RESIT = Frm84.L38_Text 'Potongan Harga Resit Trade in (%)
            
            Frm84_LM_JUMLAH_POTONG = (Frm84_LM_DEDUCT_RESIT / 100) * (-Frm84_LM_HARGA_LAST)
            
            Frm84.L37_Text = Format(Frm84_LM_JUMLAH_POTONG, "0.00")
            Frm84.L23_Text = Format((-Frm84_LM_HARGA_LAST) - Frm84_LM_JUMLAH_POTONG, "#,##0.00")
            Frm84.TB33 = Format((-Frm84_LM_HARGA_LAST) - Frm84_LM_JUMLAH_POTONG, "#,##0.00")
            x = 0
        End If
    End If
    
End If
End Sub





Private Sub L44_Text_Change()
'On Error Resume Next
Dim frm84_LM_KADAR_GST As Double
Dim Frm84_LM_HARGA As Double
Dim Frm84_LM_GST As Double

If (Frm84.L44_Text <> vbNullString And IsNumeric(Frm84.L44_Text)) And (Frm84.TB11 <> vbNullString And IsNumeric(Frm84.TB11)) Then
    Frm84_LM_GST = Frm84.TB11 'Jumlah GST (RM)
    Frm84_LM_HARGA = Frm84.L44_Text 'Harga (RM)
    
    Frm84.TB14 = Format(Frm84_LM_HARGA + Frm84_LM_GST, "#,##0.00") 'Jumlah Harga Jualan Dengan GST (RM)
Else
    Frm84.TB14 = "#,##0.00" 'Jumlah Harga Jualan Dengan GST (RM)
End If
End Sub

Private Sub L48_Text_Change()
'On Error Resume Next
If Frm84.TB2 <> vbNullString Then
    If Frm84.TB3 <> vbNullString And IsNumeric(Frm84.TB3) Then
        Call Frm84_pengiraan_harga_staff
    End If
End If
End Sub
Private Sub L49_Text_Change()
'On Error Resume Next
If Frm84.TB2 <> vbNullString Then
    If Frm84.TB3 <> vbNullString And IsNumeric(Frm84.TB3) Then
        Call Frm84_pengiraan_harga_staff
    End If
End If
End Sub
Private Sub L53_Text_Change()
'On Error Resume Next
If Frm84.TB2 <> vbNullString Then
    If Frm84.TB3 <> vbNullString And IsNumeric(Frm84.TB3) Then
        Call Frm84_pengiraan_harga_staff
    End If
End If
End Sub
Private Sub L58_Text_Change()
'On Error Resume Next
If Frm84.L58_Text <> vbNullString Then

    If IsNumeric(Frm84.L58_Text) Then
        Frm84.L22_Text = Format(Frm84.L58_Text, "#,##0.00") 'Jumlah Nilaian Resit Trade In
    End If
    
Else

    Frm84.L22_Text = Format(0, "0.00") 'Jumlah Nilaian Resit Trade In
    
End If
End Sub

Private Sub L73_Text_Change()
'On Error Resume Next
Call Frm84_pengiraan_harga_jualan
End Sub
Private Sub L75_Text_Change()
'On Error Resume Next
Call Frm84_kira_mata_ganjaran
End Sub
Private Sub L76_Text_Change()
'on error resume next
Frm84.L74_Text = Frm84.L76_Text
End Sub
Private Sub L78_Text_Change()
'On Error Resume Next
Call Frm84_kira_harga_layak_mata
End Sub
Private Sub L8_Text_Change()
'On Error Resume Next
Call frm84_kiraan_gst
End Sub
Private Sub L80_Text_Change()
'On Error Resume Next
Call Frm84_kiraan_potongan_kupon
End Sub

Private Sub ListView1_Click()
'on error resume next
LM_KEY = Frm84.ListView1.SelectedItem.Key

If LM_KEY = "Scan" Then

    Call frm84_disable_frame
    Frm84.Frame1.Visible = True
    
    Frm84.TB1.SetFocus
    
ElseIf LM_KEY = "Senarai" Then
    
    Call frm84_disable_frame
    Frm84.Frame1.Visible = True
    
    If Frm84.L4_Text <> "" Then
        If Frm84.L4_Text <> 0 Then
        
            Frm84.L89_Text = -1 'Titik Pencarian Data
            Frm84.L90_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
            Frm84.L87_Text = 0 'Paparan Page ke-xxx
            Frm84.L88_Text = 0
            
            GM_NEXT_PREV = 0
            
            Call Frm84_Senarai_Jualan_Header
            Call Frm84_Senarai_Jualan

            Frm84.Pic3.Visible = True
        Else
            MsgBox "Tiada Barang Dalam Senarai Jualan Anda.", vbInformation, "Info"
        End If
    Else
        MsgBox "Tiada Barang Dalam Senarai Jualan Anda.", vbInformation, "Info"
    End If

ElseIf LM_KEY = "Bayaran" Then
    
    If Frm84.L25_Text = "Jumlah Bayaran" Then
        
        frm130.TB33 = Format(Frm84.TB33, "#,##0.00")
        frm130.Show vbModal
        
    Else
        
        MsgBox "Tiada bayaran yang perlu dijelaskan oleh pembeli bagi urusan pembelian ini.", vbInformation, "Info"
    
    End If
    
ElseIf LM_KEY = "Invoice" Then

    Call frm84_disable_frame
    Frm84.Frame4.Visible = True
    
ElseIf LM_KEY = "Trade In" Then
    'G_TI_MODE , 0 : Tiada TI , 1 : Voucher , 2 : TI , 3 : 0%
    If G_TI_MODE = 3 Then
        G_TI_MODE = 0
        Call frm_kiraan_harga_selepas_ti
    End If
    
    If Frm84.L56_Text = 2 Then '0 : Tiada , 1 : Belian dengan trade in (sudah ada voucher) , 2 : Trade in
    
        Note = "Adakah anda yakin untuk batalkan urusan ini ?" & vbCrLf & _
                "Semua data trade in yang telah dimasukkan akan dipadamkan jika anda teruskan." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan ?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        End If
        If Answer = vbYes Then
        
            Frm84.L56_Text = 1 '0 : Tiada , 1 : Belian dengan trade in (sudah ada voucher) , 2 : Trade in
            Frm84.L57_Text = vbNullString 'No. Voucher
            Frm84.L58_Text = vbNullString 'Jumlah trade in
            
            Frm84.L16_Text = vbNullString 'No. Voucher
            Frm84.TB17 = "0.00" 'Jumlah trade in
            G_TI_MODE = 1
            'Frm84.Pic5.Visible = True
            
        End If
        
    ElseIf Frm84.L56_Text = 0 Then  '0 : Tiada , 1 : Belian dengan trade in (sudah ada voucher) , 2 : Trade in
    
        Frm84.L56_Text = 1 '0 : Tiada , 1 : Belian dengan trade in (sudah ada voucher) , 2 : Trade in
        G_TI_MODE = 2
        Frm84.L16_Text = vbNullString 'No. Voucher
        Frm84.TB17 = "0.00" 'Jumlah trade in
        
        Frm84.Frame6.Visible = True
        
    End If
    
    Call frm84_disable_frame
    Frm84.Frame6.Visible = True

ElseIf LM_KEY = "Trade In (0%)" Then

    If G_TI_MODE = 1 Or G_TI_MODE = 2 Then
        Note = "Oleh kerana terdapat maklumat bagi trade in (menggunakan voucher)," & vbCrLf & _
                vbNullString & vbCrLf & _
                "data tersebut akan direset. Teruskan?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbNo Then
            Exit Sub
        ElseIf Answer = vbYes Then
            Frm84.L56_Text = 0 '0 : Tiada , 1 : Belian dengan trade in (sudah ada voucher) , 2 : Trade in
            Frm84.L57_Text = vbNullString 'No. Voucher
            Frm84.L58_Text = vbNullString 'Jumlah trade in
            
            Frm84.L16_Text = vbNullString 'No. Voucher
            Frm84.TB17 = "0.00" 'Jumlah trade in
        End If
    End If

    Call frm84_disable_frame
    Frm84.Frame8.Visible = True
        
End If
End Sub

Private Sub ListView2_DblClick()
'On Error Resume Next
frm84_LM_No_ID = vbNullString

If IsNumeric(Frm84.ListView2.SelectedItem.Index) Then
    
    frm84_LM_No_ID = Frm84.ListView2.SelectedItem.Index
    
    If frm84_LM_No_ID <> vbNullString Then

        Call Frm84_Reset_Edit
        
        Frm84.CMD3.Visible = True
        Frm84.CMD13.Visible = False
        Frm84.CMD14.Visible = False
        
        Frm84.TB1 = vbNullString
        Frm84.TB1.Locked = False
        Frm84.TB1.BackColor = &HFFFFFF
        
        PopupMenu Frm84_PM_Menu1
        
    Else
        
        MsgBox "Tiada data.", vbExclamation, "Info"
        
    End If
    
Else

    MsgBox "Tiada data.", vbExclamation, "Info"
    
End If
End Sub




Private Sub TB1_Change()
'on error resume next
If Frm84.CB1 = 1 And Frm84.TB1 <> vbNullString Then
    Frm84.Tmr2.Enabled = False
    Frm84.Tmr2.Enabled = True
    Frm84.Tmr2.Interval = 100
End If
End Sub
Private Sub TB1_KeyPress(KeyAscii As Integer)
'on error resume next
If KeyAscii = 13 Then

    Dim Frm84_LM_LIMIT As Integer
    Dim Frm84_LM_BIL As Integer
    
    If Frm84.TB1 = vbNullString Then
        MsgBox "Sila Masukkan No. Siri Produk.", vbInformation, "Info"
        
        Frm84.TB1.SetFocus
        Exit Sub
    End If
    
    If InStr(1, Frm84.TB1, "'") <> 0 Then
        MsgBox "No. Siri Produk Mengandungi Simbol Yang Tidak Sah , ['].", vbInformation, "Info"
        
        Frm84.TB1 = vbNullString
        Frm84.TB1.SetFocus
        Exit Sub
    End If
    
    If Frm84.CB4 = 0 And Frm84.CB5 = 0 And Frm84.CB6 = 0 And Frm84.CB9 = 0 And Frm84.CB10 = 0 Then
        MsgBox "Sila pilih kategori pembeli.", vbExclamation, "info"
        Exit Sub
    End If
    
    If IsNumeric(Frm84.L46_Text) Then Frm84_LM_LIMIT = Frm84.L46_Text 'Limit Invoice
    If IsNumeric(Frm84.L4_Text) Then Frm84_LM_BIL = Frm84.L4_Text 'Kuantiti Terkini
    
    If Frm84_LM_LIMIT <> 0 Then
        If Frm84_LM_BIL >= Frm84_LM_LIMIT Then
            MsgBox "Hanya " & Frm84_LM_LIMIT & " item sahaja dibenarkan untuk dijual dalam satu invoice.", vbInformation, "Info"
        Else
            Call Frm84_Call_Product_Detail
        End If
    Else
        Call Frm84_Call_Product_Detail
    End If

End If
End Sub
Private Sub TB10_Change()
'On Error Resume Next
Dim frm84_LM_KADAR_GST As Double
Dim Frm84_LM_HARGA As Double
Dim Frm84_LM_GST As Double

Call frm84_kiraan_gst
Call Frm84_modal_dan_jual

Exit Sub

Frm84_LM_HARGA = 0

If Frm84.CB3 = 1 And (Frm84.TB10 <> vbNullString And IsNumeric(Frm84.TB10)) And (Frm84.L8_Text <> vbNullString And IsNumeric(Frm84.L8_Text)) Then
    frm84_LM_KADAR_GST = Frm84.L8_Text 'Jumlah Kadar GST (%)

    If Frm84.CB12 = 0 Then
        Frm84_LM_HARGA = Frm84.TB10 'Harga (RM)
    Else
        If Frm84.TB15 <> vbNullString And IsNumeric(Frm84.TB15) Then
            Frm84_LM_HARGA = Frm84.TB15 'Upah (RM)
        End If
    End If
    
    Frm84.TB11 = Format((frm84_LM_KADAR_GST / 100) * Frm84_LM_HARGA, "0.00") 'Jumlah Cukai GST (RM)
Else
    Frm84.TB11 = "0.00" 'Jumlah Cukai GST (RM)
End If

If Frm84.CB18 = 0 Then
    If Frm84.CB3 = 1 And (Frm84.TB10 <> vbNullString And IsNumeric(Frm84.TB10)) And (Frm84.L8_Text <> vbNullString And IsNumeric(Frm84.L8_Text)) Then
        frm84_LM_KADAR_GST = Frm84.L8_Text 'Jumlah Kadar GST (%)

        If Frm84.CB12 = 0 Then
            Frm84_LM_HARGA = Frm84.TB10 'Harga (RM)
        Else
            If Frm84.TB15 <> vbNullString And IsNumeric(Frm84.TB15) Then
                Frm84_LM_HARGA = Frm84.TB15 'Upah (RM)
            End If
        End If

        
        Frm84.TB11 = Format((frm84_LM_KADAR_GST / 100) * Frm84_LM_HARGA, "#,##0.00") 'Jumlah Cukai GST (RM)
        Frm84.L44_Text = Format(Frm84_LM_HARGA, "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
    Else
        Frm84.TB11 = Format(0, "#,##0.00") 'Jumlah Cukai GST (RM)
    End If
ElseIf Frm84.CB18 = 1 Then
    If Frm84.CB3 = 1 And (Frm84.TB10 <> vbNullString And IsNumeric(Frm84.TB10)) And (Frm84.L8_Text <> vbNullString And IsNumeric(Frm84.L8_Text)) Then
        frm84_LM_KADAR_GST = Frm84.L8_Text 'Jumlah Kadar GST (%)

        If Frm84.CB12 = 0 Then
            Frm84_LM_HARGA = Frm84.TB10 'Harga (RM)
        Else
            If Frm84.TB15 <> vbNullString And IsNumeric(Frm84.TB15) Then
                Frm84_LM_HARGA = Frm84.TB15 'Upah (RM)
            End If
        End If
        
        Frm84.L44_Text = Format(Frm84_LM_HARGA / (1 + (frm84_LM_KADAR_GST / 100)), "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
        Frm84.TB11 = Format(Frm84_LM_HARGA - (Frm84_LM_HARGA / (1 + (frm84_LM_KADAR_GST / 100))), "#,##0.00") 'Jumlah Cukai GST (RM)
    Else
        Frm84.TB11 = "0.00" 'Jumlah Cukai GST (RM)
    End If
End If

If Frm84.CB2 = 1 Then

    If Frm84.CB12 = 0 Then
        If IsNumeric(Frm84.TB10) Then
            Frm84.L44_Text = Format(Frm84.TB10, "#,##0.00")
        Else
            Frm84.L44_Text = Format(0, "#,##0.00")
        End If
    Else
        If IsNumeric(Frm84.TB15) Then
            Frm84.L44_Text = Format(Frm84.TB15, "#,##0.00")
        Else
            Frm84.L44_Text = Format(0, "#,##0.00")
        End If
    End If
End If

Call Frm84_modal_dan_jual
End Sub
Private Sub TB11_Change()
'on error resume next
'On Error Resume Next
Dim frm84_LM_KADAR_GST As Double
Dim Frm84_LM_HARGA As Double
Dim Frm84_LM_GST As Double

Frm84_LM_HARGA = 0

If (Frm84.L44_Text <> vbNullString And IsNumeric(Frm84.L44_Text)) And (Frm84.TB11 <> vbNullString And IsNumeric(Frm84.TB11)) Then
    Frm84_LM_GST = Frm84.TB11 'Jumlah GST (RM)
    Frm84_LM_HARGA = Frm84.L44_Text 'Harga (RM)
    
    Frm84.TB14 = Format(Frm84_LM_HARGA + Frm84_LM_GST, "#,##0.00") 'Jumlah Harga Jualan Dengan GST (RM)
Else
    Frm84.TB14 = "#,##0.00" 'Jumlah Harga Jualan Dengan GST (RM)
End If

Exit Sub

'Dim Frm84_LM_GST As Double

If (Frm84.TB10 <> vbNullString And IsNumeric(Frm84.TB10)) And (Frm84.TB11 <> vbNullString And IsNumeric(Frm84.TB11)) Then
    Frm84_LM_GST = Frm84.TB11 'Jumlah GST (RM)

    If Frm84.CB12 = 0 Then
        Frm84_LM_HARGA = Frm84.TB10 'Harga (RM)
    Else
        If Frm84.TB15 <> vbNullString And IsNumeric(Frm84.TB15) Then
            Frm84_LM_HARGA = Frm84.TB15 'Upah (RM)
        End If
    End If
    
    Frm84.TB14 = Format(Frm84_LM_HARGA + Frm84_LM_GST, "0.00") 'Jumlah Harga Jualan Dengan GST (RM)
Else
    Frm84.TB14 = "0.00" 'Jumlah Harga Jualan Dengan GST (RM)
End If
End Sub
Private Sub TB12_Change()
'on error resume next
Dim Frm84_BERAT As Double
Dim Frm84_KOMISEN_PER_GRAM As Double

Call Frm84_pengiraan_komisyen_dropship

Exit Sub

If Frm84.CB7 = 1 Then
    If ((Frm84.TB4 <> vbNullString And IsNumeric(Frm84.TB4)) And (Frm84.TB12 <> vbNullString And IsNumeric(Frm84.TB12))) Then
        Frm84_BERAT = Frm84.TB4 'Berat
        Frm84_KOMISEN_PER_GRAM = Frm84.TB12 'Komisyen Per Gram
        
        Frm84.TB13 = Format((Frm84_BERAT * Frm84_KOMISEN_PER_GRAM), "0.00") 'Jumlah Komisyen
    Else
        Frm84.TB13 = "0.00" 'Jumlah Komisyen
    End If
Else
    Frm84.TB12 = "0.00" 'Komisyen Per Gram (Dropship)
    'Frm84.TB43 = 0 '% komisyen upah
    Frm84.TB44 = "0.00" 'Komisyen upah
    Frm84.TB13 = "0.00" 'Jumlah Komisyen (Dropship)
End If
End Sub

Private Sub TB15_Change()
'on error resume next
Dim Frm84_BERAT As Double
Dim Frm84_HARGA_PER_GRAM As Double
Dim Frm84_UPAH As Double

Call frm84_kiraan_harga_asal
Call frm84_kiraan_gst
Call Frm84_pengiraan_komisyen_upah

Exit Sub

If GLOBAL_DISABLE = 0 Then
    
    Call Frm84_pengiraan_komisyen_upah

    Frm84_BERAT = 0
    Frm84_HARGA_PER_GRAM = 0
    Frm84_UPAH = 0
    
    If ((Frm84.TB4 <> vbNullString And IsNumeric(Frm84.TB4)) And (Frm84.TB5 <> vbNullString And IsNumeric(Frm84.TB5)) And (Frm84.TB15 <> vbNullString And IsNumeric(Frm84.TB15))) Then
        Frm84_BERAT = Frm84.TB4 'Berat
        Frm84_HARGA_PER_GRAM = Frm84.TB5 'Harga Per Gram
        Frm84_UPAH = Frm84.TB15 'Upah
        
        Frm84.TB6 = Format((Frm84_BERAT * Frm84_HARGA_PER_GRAM) + Frm84_UPAH, "0.00") 'Harga Asal
    Else
        Frm84.TB6 = "0.00" 'Harga Asal
    End If
    
End If
End Sub
Private Sub TB19_Change()
'on error resume next
Call frm84_harga_selepas_diskaun
End Sub
Private Sub TB20_Change()
'on error resume next
Dim Frm84_HARGA_LEPAS_DISKAUN As Double
Dim Frm84_ADJUSTMENT As Double

Call Frm84_pengiraan_harga_jualan
End Sub
Private Sub TB22_Change()
'On Error Resume Next
Call Frm84_kira_upah
End Sub




Private Sub TB33_Change()
'On Error Resume Next
If Frm84.L25_Text = "Jumlah Bayaran" Then
    frm130.TB33 = Format(Frm84.TB33, "#,##0.00")
    'Call frm130_kiraan_cara_bayaran
Else
    frm130.TB33 = "0.00"
End If
End Sub

Private Sub TB34_Change()
'On Error Resume Next
Call Frm84_kira_harga_layak_mata
Call Frm84_pengiraan_harga_jualan
End Sub
Private Sub TB35_Change()
'On Error Resume Next
Call Frm84_kira_mata_ganjaran
End Sub
Private Sub TB36_Change()
'On Error Resume Next
Call Frm84_nilai_mata_tebus
End Sub

Private Sub TB37_Change()
'On Error Resume Next
Call Frm84_nilai_mata_tebus
End Sub
Private Sub TB4_Change()
'on error resume next
Dim Frm84_BERAT As Double
Dim Frm84_HARGA_PER_GRAM As Double
Dim Frm84_UPAH As Double
Dim Frm84_KOMISEN_PER_GRAM As Double

Call Frm84_pengiraan_komisyen_dropship
Call frm84_kiraan_harga_asal
Call Frm84_kira_upah

Exit Sub

If GLOBAL_DISABLE = 0 Then

    Frm84_BERAT = 0
    Frm84_HARGA_PER_GRAM = 0
    Frm84_KOMISEN_PER_GRAM = 0
    Frm84_UPAH = 0
    
    If ((Frm84.TB4 <> vbNullString And IsNumeric(Frm84.TB4)) And (Frm84.TB5 <> vbNullString And IsNumeric(Frm84.TB5)) And (Frm84.TB15 <> vbNullString And IsNumeric(Frm84.TB15))) Then
        Frm84_BERAT = Frm84.TB4 'Berat
        Frm84_HARGA_PER_GRAM = Frm84.TB5 'Harga Per Gram
        Frm84_UPAH = Frm84.TB15 'Upah
        
        Frm84.TB6 = Format((Frm84_BERAT * Frm84_HARGA_PER_GRAM) + Frm84_UPAH, "0.00") 'Harga Asal
    Else
        Frm84.TB6 = "0.00" 'Harga Asal
    End If
    
    Call Frm84_pengiraan_komisyen_dropship
    Call Frm84_modal_dan_jual
    'If Frm84.CB7 = 1 Then
    '    If ((Frm84.TB4 <> vbNullString And IsNumeric(Frm84.TB4)) And (Frm84.TB12 <> vbNullString And IsNumeric(Frm84.TB12))) Then
    '        Frm84_BERAT = Frm84.TB4 'Berat
    '        Frm84_KOMISEN_PER_GRAM = Frm84.TB12 'Komisyen Per Gram
            
    '        Frm84.TB13 = Format((Frm84_BERAT * Frm84_KOMISEN_PER_GRAM), "0.00") 'Jumlah Komisyen
    '    Else
    '        Frm84.TB13 = "0.00" 'Jumlah Komisyen
    '    End If
    'Else
    '    Frm84.TB12 = "0.00" 'Komisyen Per Gram (Dropship)
    '    Frm84.TB43 = 0 '% komisyen upah
    '    Frm84.TB44 = "0.00" 'Komisyen upah
    '    Frm84.TB13 = "0.00" 'Jumlah Komisyen (Dropship)
    'End If
    
    'If Frm84.TB2 <> vbNullString Then
    '    If Frm84.TB3 <> vbNullString And IsNumeric(Frm84.TB3) Then
    '        If IsNumeric(Frm84.TB4) And Frm84.TB2 <> vbNullString Then
    '            If Format(Frm84.TB4, "0.00") <> "0.00" Then Call Frm84_pengiraan_harga_staff
    '        End If
    '    End If
    'End If
End If
End Sub

Private Sub TB42_Change()
'on error resume next
Call Frm84_pengiraan_harga_jualan
End Sub
Private Sub TB43_Change()
'on error resume next
Call Frm84_pengiraan_komisyen_upah
End Sub
Private Sub TB44_Change()
'on error resume next
Call Frm84_pengiraan_komisyen_dropship
End Sub



Private Sub TB5_Change()
'on error resume next
Dim Frm84_BERAT As Double
Dim Frm84_HARGA_PER_GRAM As Double
Dim Frm84_UPAH As Double

Call frm84_kiraan_harga_asal

Exit Sub

If GLOBAL_DISABLE = 0 Then

    Frm84_BERAT = 0
    Frm84_HARGA_PER_GRAM = 0
    Frm84_UPAH = 0
    
    If ((Frm84.TB4 <> vbNullString And IsNumeric(Frm84.TB4)) And (Frm84.TB5 <> vbNullString And IsNumeric(Frm84.TB5)) And (Frm84.TB15 <> vbNullString And IsNumeric(Frm84.TB15))) Then
        Frm84_BERAT = Frm84.TB4 'Berat
        Frm84_HARGA_PER_GRAM = Frm84.TB5 'Harga Per Gram
        Frm84_UPAH = Frm84.TB15 'Upah
        
        Frm84.TB6 = Format((Frm84_BERAT * Frm84_HARGA_PER_GRAM) + Frm84_UPAH, "0.00") 'Harga Asal
    Else
        Frm84.TB6 = "0.00" 'Harga Asal
    End If
    
End If
End Sub
Private Sub TB6_Change()
'on error resume next
Dim Frm84_HARGA_ASAL As Double
Dim Frm84_DISKAUN As Double

Call frm84_selepas_diskaun

Exit Sub

If GLOBAL_DISABLE = 0 Then

    Frm84_HARGA_ASAL = 0
    Frm84_DISKAUN = 0
    
    If ((Frm84.TB6 <> vbNullString And IsNumeric(Frm84.TB6)) And (Frm84.TB7 <> vbNullString And IsNumeric(Frm84.TB7))) Then
        Frm84_HARGA_ASAL = Frm84.TB6 'Harga Asal
        Frm84_DISKAUN = Frm84.TB7 'Diskaun
        
        Frm84.TB8 = Format(Frm84_HARGA_ASAL - ((Frm84_DISKAUN / 100) * Frm84_HARGA_ASAL), "0.00") 'Harga Selepas Diskaun
    Else
        Frm84.TB8 = "0.00" 'Harga Selepas Diskaun
    End If
    
End If
End Sub
Private Sub TB7_Change()
'on error resume next
Dim Frm84_HARGA_ASAL As Double
Dim Frm84_DISKAUN As Double

Call frm84_selepas_diskaun

Exit Sub

If GLOBAL_DISABLE = 0 Then

    Frm84_HARGA_ASAL = 0
    Frm84_DISKAUN = 0
    
    If ((Frm84.TB6 <> vbNullString And IsNumeric(Frm84.TB6)) And (Frm84.TB7 <> vbNullString And IsNumeric(Frm84.TB7))) Then
        Frm84_HARGA_ASAL = Frm84.TB6 'Harga Asal
        Frm84_DISKAUN = Frm84.TB7 'Diskaun
        
        Frm84.TB8 = Format(Frm84_HARGA_ASAL - ((Frm84_DISKAUN / 100) * Frm84_HARGA_ASAL), "0.00") 'Harga Selepas Diskaun
    Else
        Frm84.TB8 = "0.00" 'Harga Selepas Diskaun
    End If
    
End If
End Sub
Private Sub TB8_Change()
'on error resume next
Dim Frm84_HARGA_LEPAS_DISKAUN As Double
Dim Frm84_ADJUSTMENT As Double

Call frm84_harga_jualan

Exit Sub

If GLOBAL_DISABLE = 0 Then

    Frm84_HARGA_LEPAS_DISKAUN = 0
    Frm84_ADJUSTMENT = 0
    
    If ((Frm84.TB8 <> vbNullString And IsNumeric(Frm84.TB8)) And (Frm84.TB9 <> vbNullString And IsNumeric(Frm84.TB9))) Then
        Frm84_HARGA_LEPAS_DISKAUN = Frm84.TB8 'Harga Lepas Diskaun
        Frm84_ADJUSTMENT = Frm84.TB9 'Adjustment
        
        Frm84.TB10 = Format(Frm84_HARGA_LEPAS_DISKAUN - Frm84_ADJUSTMENT, "0.00") 'Harga Jualan
    Else
        Frm84.TB10 = "0.00" 'Harga Jualan
    End If
    
    If Frm84.TB2 <> vbNullString And Frm84.TB3 = vbNullString Then
        Call Frm84_pengiraan_harga_bp_staff
    End If
    If Frm84.TB2 <> vbNullString Then
        If Frm84.TB3 <> vbNullString And IsNumeric(Frm84.TB3) Then
            If Format(Frm84.TB4, "0.00") <> "0.00" Then Call Frm84_pengiraan_harga_staff
        End If
    End If
    
End If
End Sub
Private Sub TB9_Change()
'on error resume next
Dim Frm84_HARGA_LEPAS_DISKAUN As Double
Dim Frm84_ADJUSTMENT As Double

Call frm84_harga_jualan

Exit Sub

If GLOBAL_DISABLE = 0 Then

    Frm84_HARGA_LEPAS_DISKAUN = 0
    Frm84_ADJUSTMENT = 0
    
    If ((Frm84.TB8 <> vbNullString And IsNumeric(Frm84.TB8)) And (Frm84.TB9 <> vbNullString And IsNumeric(Frm84.TB9))) Then
        Frm84_HARGA_LEPAS_DISKAUN = Frm84.TB8 'Harga Lepas Diskaun
        Frm84_ADJUSTMENT = Frm84.TB9 'Adjustment
        
        Frm84.TB10 = Format(Frm84_HARGA_LEPAS_DISKAUN - Frm84_ADJUSTMENT, "0.00") 'Harga Jualan
    Else
        Frm84.TB10 = "0.00" 'Harga Jualan
    End If
    
End If
End Sub
Private Sub Tmr1_Timer()
'on error resume next
Frm84.L2_Text = DateTime.Time$
End Sub
Private Sub Tmr2_Timer()
'On Error Resume Next
Dim Frm84_LM_LIMIT As Integer
Dim Frm84_LM_BIL As Integer

If Frm84.CB1 = 1 And Frm84.TB1 <> vbNullString And Frm84.Tmr2.Enabled = True Then
    If Frm84.Tmr2.Interval = 100 Then
        If InStr(1, Frm84.TB1, "'") <> 0 Then
            MsgBox "No. Siri Produk Mengandungi Simbol Yang Tidak Sah , ['].", vbInformation, "Info"
            
            Frm84.TB1 = vbNullString
            Exit Sub
        End If
        If Frm84.CB4 = 0 And Frm84.CB5 = 0 And Frm84.CB6 = 0 And Frm84.CB9 = 0 And Frm84.CB10 = 0 Then
            MsgBox "Sila pilih kategori pembeli.", vbExclamation, "info"
            Exit Sub
        End If
        
        
        If IsNumeric(Frm84.L46_Text) Then Frm84_LM_LIMIT = Frm84.L46_Text 'Limit Invoice
        If IsNumeric(Frm84.L4_Text) Then Frm84_LM_BIL = Frm84.L4_Text 'Kuantiti Terkini
        
        If Frm84_LM_LIMIT <> 0 Then
            If Frm84_LM_BIL >= Frm84_LM_LIMIT Then
                Frm84.TB1 = vbNullString
                MsgBox "Hanya " & Frm84_LM_LIMIT & " item sahaja dibenarkan untuk dijual dalam satu invoice.", vbInformation, "Info"
            Else
                Call Frm84_Call_Product_Detail
            End If
        Else
            Call Frm84_Call_Product_Detail
        End If
        
    End If
End If
End Sub
Private Sub Tmr3_Timer()
'On Error Resume Next
Dim Frm84_LM_BELI As Double
Dim Frm84_LM_JUAL As Double

Frm84_LM_BELI = 0
Frm84_LM_JUAL = 0

If (Frm84.L67_Text <> vbNullString And IsNumeric(Frm84.L67_Text)) And (Frm84.L69_Text <> vbNullString And IsNumeric(Frm84.L69_Text)) Then
 
    Frm84_LM_BELI = Frm84.L69_Text
    Frm84_LM_JUAL = Frm84.L67_Text
    
    If Frm84_LM_JUAL < Frm84_LM_BELI Then
        
        Frm84.L67_Text.FontBold = True
        
        If Frm84.L67_Text.Visible = True Then
            Frm84.L67_Text.Visible = False
        Else
            Frm84.L67_Text.Visible = True
        End If
        
    Else
        
        Frm84.L67_Text.Visible = True
        Frm84.L67_Text.FontBold = False
        
    End If

Else

    Frm84.L67_Text.Visible = True
    Frm84.L67_Text.FontBold = False
        
End If

End Sub
