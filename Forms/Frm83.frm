VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm83 
   Caption         =   "Penerimaan Stok Barang Kemas (Barang Kemas / Barang Permata / Gold Bar)"
   ClientHeight    =   13125
   ClientLeft      =   120
   ClientTop       =   -150
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
   Icon            =   "Frm83.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13125
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD20 
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
      Left            =   8760
      MouseIcon       =   "Frm83.frx":0ECA
      MousePointer    =   99  'Custom
      Picture         =   "Frm83.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   10320
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton CMD21 
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
      Left            =   10920
      MouseIcon       =   "Frm83.frx":379E
      MousePointer    =   99  'Custom
      Picture         =   "Frm83.frx":3AA8
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   10320
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton CMD23 
      BackColor       =   &H8000000A&
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
      Left            =   10920
      MouseIcon       =   "Frm83.frx":6072
      MousePointer    =   99  'Custom
      Picture         =   "Frm83.frx":637C
      Style           =   1  'Graphical
      TabIndex        =   165
      Top             =   10320
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.TextBox TB42 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Height          =   360
      Left            =   10350
      Locked          =   -1  'True
      TabIndex        =   182
      Text            =   "0.00"
      Top             =   10920
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox TB41 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   10350
      MaxLength       =   10
      TabIndex        =   181
      Text            =   "0.00"
      ToolTipText     =   "Ruangan ini hanya boleh diubah/diisi untuk urusan BUYBACK sahaja."
      Top             =   10560
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox TB40 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   10350
      TabIndex        =   180
      Text            =   "0.00"
      Top             =   10200
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton CMD25 
      Caption         =   "Info Pembeli - (Berdaftar)"
      Height          =   930
      Left            =   120
      MouseIcon       =   "Frm83.frx":8946
      MousePointer    =   99  'Custom
      Picture         =   "Frm83.frx":8C50
      Style           =   1  'Graphical
      TabIndex        =   172
      Top             =   10440
      Width           =   2415
   End
   Begin VB.CommandButton CMD24 
      Caption         =   "Info Pembeli - (Tidak berdaftar)"
      Height          =   1050
      Left            =   120
      MouseIcon       =   "Frm83.frx":B21A
      MousePointer    =   99  'Custom
      Picture         =   "Frm83.frx":B524
      Style           =   1  'Graphical
      TabIndex        =   171
      Top             =   9360
      Width           =   2415
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Senarai Barang Trade In"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9255
      Left            =   15240
      TabIndex        =   160
      Top             =   360
      Visible         =   0   'False
      Width           =   19095
      Begin VB.CommandButton CMD26 
         Caption         =   "Back"
         Height          =   810
         Left            =   16680
         MouseIcon       =   "Frm83.frx":C5EE
         MousePointer    =   99  'Custom
         Picture         =   "Frm83.frx":C8F8
         Style           =   1  'Graphical
         TabIndex        =   174
         ToolTipText     =   "Tutup senarai ini."
         Top             =   8040
         Width           =   1095
      End
      Begin VB.CommandButton CMD27 
         Caption         =   "Next"
         Height          =   810
         Left            =   17880
         MouseIcon       =   "Frm83.frx":D9C2
         MousePointer    =   99  'Custom
         Picture         =   "Frm83.frx":DCCC
         Style           =   1  'Graphical
         TabIndex        =   173
         ToolTipText     =   "Tutup senarai ini."
         Top             =   8040
         Width           =   1095
      End
      Begin VB.CommandButton CMD4 
         Caption         =   "Tutup Senarai Ini"
         Height          =   930
         Left            =   7800
         MouseIcon       =   "Frm83.frx":ED96
         MousePointer    =   99  'Custom
         Picture         =   "Frm83.frx":F0A0
         Style           =   1  'Graphical
         TabIndex        =   162
         ToolTipText     =   "Tutup senarai ini."
         Top             =   8160
         Width           =   3375
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   7635
         Left            =   120
         TabIndex        =   161
         Top             =   360
         Width           =   18840
         _ExtentX        =   33232
         _ExtentY        =   13467
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
      Begin VB.Label L26_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   1320
         TabIndex        =   188
         Top             =   8040
         Width           =   2400
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah : RM"
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
         TabIndex        =   186
         Top             =   8040
         Width           =   1095
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
         Left            =   16080
         TabIndex        =   178
         Top             =   8040
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
         Left            =   15480
         TabIndex        =   177
         Top             =   8040
         Width           =   375
      End
      Begin VB.Label L69_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L69_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   12480
         TabIndex        =   176
         Top             =   8520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L70_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L70_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   12360
         TabIndex        =   175
         Top             =   8160
         Visible         =   0   'False
         Width           =   855
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
         Left            =   14160
         TabIndex        =   179
         Top             =   8040
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Maklumat Stok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9255
      Left            =   21840
      TabIndex        =   92
      Top             =   0
      Visible         =   0   'False
      Width           =   18495
      Begin VB.CommandButton CMD1 
         Caption         =   "Masukkan Dalam Senarai"
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
         MouseIcon       =   "Frm83.frx":1166A
         MousePointer    =   99  'Custom
         Picture         =   "Frm83.frx":11974
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   8040
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CommandButton CMD13 
         BackColor       =   &H8000000A&
         Caption         =   "Masukkan Dalam Senarai"
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
         Left            =   6240
         MouseIcon       =   "Frm83.frx":13F3E
         MousePointer    =   99  'Custom
         Picture         =   "Frm83.frx":14248
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   8040
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CommandButton CMD14 
         BackColor       =   &H8000000A&
         Caption         =   "Batal Edit Data"
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
         Left            =   9240
         MouseIcon       =   "Frm83.frx":16812
         MousePointer    =   99  'Custom
         Picture         =   "Frm83.frx":16B1C
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   8040
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CommandButton CMD12 
         Caption         =   "Masukkan Dalam Senarai"
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
         MouseIcon       =   "Frm83.frx":190E6
         MousePointer    =   99  'Custom
         Picture         =   "Frm83.frx":193F0
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   8040
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CommandButton CMD6 
         BackColor       =   &H8000000A&
         Caption         =   "Masukkan Dalam Senarai"
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
         Left            =   6240
         MouseIcon       =   "Frm83.frx":1B9BA
         MousePointer    =   99  'Custom
         Picture         =   "Frm83.frx":1BCC4
         Style           =   1  'Graphical
         TabIndex        =   157
         Top             =   8040
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CommandButton CMD7 
         BackColor       =   &H8000000A&
         Caption         =   "Batal Edit Data"
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
         MouseIcon       =   "Frm83.frx":1E28E
         MousePointer    =   99  'Custom
         Picture         =   "Frm83.frx":1E598
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   8040
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Maklumat Tambahan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5655
         Left            =   9720
         TabIndex        =   139
         Top             =   2280
         Width           =   8535
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No. Siri Produk"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   120
            TabIndex        =   153
            Top             =   4560
            Visible         =   0   'False
            Width           =   4695
            Begin VB.TextBox TB6 
               Alignment       =   2  'Center
               BackColor       =   &H8000000A&
               Height          =   360
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   155
               Top             =   360
               Width           =   960
            End
            Begin VB.TextBox TB7 
               Alignment       =   2  'Center
               BackColor       =   &H8000000A&
               Height          =   360
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   154
               Top             =   360
               Width           =   1740
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "No. Siri Produk  * :"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   120
               TabIndex        =   156
               Top             =   360
               Width           =   1785
            End
         End
         Begin VB.TextBox TB24 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   3120
            MaxLength       =   10
            TabIndex        =   25
            Text            =   "0.00"
            Top             =   480
            Width           =   1000
         End
         Begin VB.TextBox TB25 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   3120
            TabIndex        =   26
            Text            =   "0.00"
            Top             =   840
            Width           =   1000
         End
         Begin VB.TextBox TB26 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   3120
            TabIndex        =   27
            Text            =   "0.00"
            Top             =   1200
            Width           =   1000
         End
         Begin VB.TextBox TB33 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   7200
            TabIndex        =   30
            Text            =   "0.00"
            Top             =   1200
            Visible         =   0   'False
            Width           =   1000
         End
         Begin VB.TextBox TB32 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   7200
            TabIndex        =   29
            Text            =   "0.00"
            Top             =   840
            Width           =   1000
         End
         Begin VB.TextBox TB31 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   7200
            TabIndex        =   28
            Text            =   "0.00"
            Top             =   480
            Width           =   1000
         End
         Begin VB.TextBox TB16 
            BackColor       =   &H00FFFFFF&
            Height          =   945
            Left            =   3120
            MultiLine       =   -1  'True
            TabIndex        =   31
            Top             =   1680
            Width           =   5085
         End
         Begin VB.TextBox TB15 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   3120
            TabIndex        =   33
            Top             =   3030
            Width           =   1485
         End
         Begin VB.TextBox TB36 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   3120
            TabIndex        =   34
            Top             =   3480
            Width           =   1485
         End
         Begin VB.TextBox TB37 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   3120
            TabIndex        =   35
            Top             =   3840
            Width           =   1485
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   360
            Left            =   3120
            TabIndex        =   32
            Top             =   2640
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
            Format          =   167313408
            CurrentDate     =   41561
         End
         Begin VB.Label L27_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Upah Jualan Pelanggan    RM"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   120
            TabIndex        =   152
            Top             =   510
            Width           =   2925
         End
         Begin VB.Label L28_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Upah Jualan Member       RM"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   120
            TabIndex        =   151
            Top             =   855
            Width           =   2925
         End
         Begin VB.Label L29_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Upah Jualan Pengedar     RM"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   120
            TabIndex        =   150
            Top             =   1230
            Width           =   2925
         End
         Begin VB.Label L35_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Upah Jualan M.Dealer      RM"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   4200
            TabIndex        =   149
            Top             =   1230
            Visible         =   0   'False
            Width           =   2925
         End
         Begin VB.Label L34_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Upah Jualan N.Dealer      RM"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   4200
            TabIndex        =   148
            Top             =   855
            Width           =   2925
         End
         Begin VB.Label L33_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Upah Jualan RAF             RM"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   4200
            TabIndex        =   147
            Top             =   510
            Width           =   2925
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   120
            TabIndex        =   146
            Top             =   1680
            Width           =   2925
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tarikh Belian * :"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   240
            TabIndex        =   145
            Top             =   2670
            Width           =   2865
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "**Sila masukkan No. invoice daripada supplier (Jika Ada)"
            ForeColor       =   &H00000000&
            Height          =   525
            Left            =   4680
            TabIndex        =   144
            Top             =   3000
            Width           =   3465
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "No. Invoice :"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   240
            TabIndex        =   143
            Top             =   3050
            Width           =   2865
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Code 1 :"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   240
            TabIndex        =   142
            Top             =   3510
            Width           =   2865
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Code 2 :"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   240
            TabIndex        =   141
            Top             =   3870
            Width           =   2865
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Sila masukkan Code 1 dan Code 2 seringkas yang mungkin kerana mungkin code ini akan dicetak pada tag barcode."
            ForeColor       =   &H00000000&
            Height          =   1005
            Left            =   4680
            TabIndex        =   140
            Top             =   3600
            Width           =   3465
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dimension / Ukuran"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   11520
         TabIndex        =   132
         Top             =   360
         Width           =   6735
         Begin VB.ComboBox CBB5 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   3840
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   840
            Width           =   1380
         End
         Begin VB.TextBox TB29 
            Alignment       =   2  'Center
            BackColor       =   &H8000000A&
            Height          =   360
            Left            =   3840
            Locked          =   -1  'True
            TabIndex        =   136
            Text            =   "0.00"
            Top             =   480
            Width           =   1365
         End
         Begin VB.TextBox TB14 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   1440
            TabIndex        =   23
            Top             =   1200
            Width           =   1000
         End
         Begin VB.TextBox TB13 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   1440
            TabIndex        =   22
            Top             =   840
            Width           =   1000
         End
         Begin VB.TextBox TB12 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   1440
            TabIndex        =   21
            Top             =   480
            Width           =   1000
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dulang * :"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2640
            TabIndex        =   138
            Top             =   870
            Width           =   1185
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Berat Riyal :"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2640
            TabIndex        =   137
            Top             =   495
            Width           =   1185
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Saiz :"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   120
            TabIndex        =   135
            Top             =   1230
            Width           =   1305
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Lebar (cm) :"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   120
            TabIndex        =   134
            Top             =   855
            Width           =   1305
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Panjang (cm) :"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   120
            TabIndex        =   133
            Top             =   495
            Width           =   1305
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Spesifikasi Stok"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5655
         Left            =   120
         TabIndex        =   104
         Top             =   2280
         Width           =   9495
         Begin VB.TextBox TB4 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   2760
            TabIndex        =   11
            Text            =   "0.00"
            Top             =   2880
            Width           =   1500
         End
         Begin VB.TextBox TB10 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   2760
            TabIndex        =   12
            Text            =   "0.00"
            Top             =   3240
            Width           =   1500
         End
         Begin VB.TextBox TB9 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   2760
            TabIndex        =   7
            Text            =   "0.00"
            Top             =   1200
            Width           =   1500
         End
         Begin VB.TextBox TB8 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   6
            Text            =   "0.00"
            Top             =   840
            Width           =   1500
         End
         Begin VB.TextBox TB19 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   13
            Text            =   "0.00"
            ToolTipText     =   "Ruangan ini hanya boleh diubah/diisi untuk urusan BUYBACK sahaja."
            Top             =   3600
            Width           =   1500
         End
         Begin VB.TextBox TB20 
            Alignment       =   2  'Center
            BackColor       =   &H8000000A&
            Height          =   360
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   16
            Text            =   "0.00"
            Top             =   4680
            Width           =   1500
         End
         Begin VB.TextBox TB21 
            Alignment       =   2  'Center
            BackColor       =   &H8000000A&
            Height          =   360
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   14
            Text            =   "0.00"
            Top             =   3960
            Width           =   1500
         End
         Begin VB.TextBox TB22 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   15
            Text            =   "0.00"
            Top             =   4320
            Width           =   1500
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
            Left            =   240
            TabIndex        =   9
            Top             =   2280
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
            Left            =   240
            TabIndex        =   8
            Top             =   1970
            Width           =   200
         End
         Begin VB.CheckBox CB5 
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
            Left            =   2400
            TabIndex        =   5
            Top             =   405
            Width           =   200
         End
         Begin VB.CheckBox CB4 
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
            TabIndex        =   4
            Top             =   405
            Width           =   200
         End
         Begin VB.TextBox TB35 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   10
            Text            =   "0.00"
            Top             =   2220
            Width           =   1500
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Maklumat Cukai GST"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   4560
            TabIndex        =   105
            Top             =   240
            Width           =   4695
            Begin VB.CheckBox CB12 
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
               TabIndex        =   17
               Top             =   360
               Width           =   200
            End
            Begin VB.CheckBox CB11 
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
               TabIndex        =   20
               Top             =   1440
               Width           =   200
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
               TabIndex        =   18
               Top             =   960
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
               TabIndex        =   19
               Top             =   1215
               Width           =   200
            End
            Begin VB.TextBox TB27 
               BackColor       =   &H8000000A&
               Height          =   360
               Left            =   3240
               Locked          =   -1  'True
               TabIndex        =   107
               Text            =   "0.00"
               Top             =   1750
               Width           =   1365
            End
            Begin VB.TextBox TB28 
               BackColor       =   &H8000000A&
               Height          =   360
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   106
               Top             =   2160
               Width           =   3180
            End
            Begin VB.Label Label79 
               BackStyle       =   0  'Transparent
               Caption         =   "Tidak Bertanda : GST pada harga barang Bertanda : GST pada UPAH"
               ForeColor       =   &H00000000&
               Height          =   525
               Left            =   480
               TabIndex        =   115
               Top             =   330
               Width           =   3585
            End
            Begin VB.Label Label78 
               BackStyle       =   0  'Transparent
               Caption         =   "Standard Rated Inclusive (SR)"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   480
               TabIndex        =   114
               Top             =   1390
               Width           =   3705
            End
            Begin VB.Label Label77 
               BackStyle       =   0  'Transparent
               Caption         =   "Standard Rated (SR)"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   480
               TabIndex        =   113
               Top             =   1160
               Width           =   3705
            End
            Begin VB.Label Label76 
               BackStyle       =   0  'Transparent
               Caption         =   "Zero Rated ZR(L)"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   480
               TabIndex        =   112
               Top             =   920
               Width           =   3705
            End
            Begin VB.Label Label50 
               BackStyle       =   0  'Transparent
               Caption         =   "Jumlah Cukai GST                 : RM"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   240
               TabIndex        =   111
               Top             =   1800
               Width           =   3585
            End
            Begin VB.Label Label59 
               BackStyle       =   0  'Transparent
               Caption         =   "@        %"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1920
               TabIndex        =   110
               Top             =   1800
               Width           =   840
            End
            Begin VB.Label L8_Text 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1920
               TabIndex        =   109
               Top             =   1800
               Width           =   840
            End
            Begin VB.Label Label62 
               BackStyle       =   0  'Transparent
               Caption         =   "No. ID GST :"
               ForeColor       =   &H00000000&
               Height          =   300
               Left            =   240
               TabIndex        =   108
               Top             =   2190
               Width           =   1665
            End
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Berat * (g) :"
            Height          =   255
            Left            =   330
            TabIndex        =   131
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Harga Per Gram * (RM/g) :"
            Height          =   255
            Left            =   330
            TabIndex        =   130
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Harga Asal-Spread (RM) :"
            Height          =   255
            Left            =   330
            TabIndex        =   129
            Top             =   3960
            Width           =   2415
         End
         Begin VB.Label Label56 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Upah * (RM) :"
            Height          =   255
            Left            =   330
            TabIndex        =   128
            Top             =   2880
            Width           =   2415
         End
         Begin VB.Label Label61 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Harga Asal * (RM) :"
            Height          =   255
            Left            =   330
            TabIndex        =   127
            Top             =   3240
            Width           =   2415
         End
         Begin VB.Label Label64 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Spread * (RM) :"
            Height          =   255
            Left            =   330
            TabIndex        =   126
            Top             =   3600
            Width           =   2415
         End
         Begin VB.Label Label66 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Adjustment * (RM) :"
            Height          =   255
            Left            =   330
            TabIndex        =   125
            Top             =   4320
            Width           =   2415
         End
         Begin VB.Label Label67 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Harga Belian * (RM) :"
            Height          =   255
            Left            =   330
            TabIndex        =   124
            Top             =   4680
            Width           =   2415
         End
         Begin VB.Label Label70 
            BackStyle       =   0  'Transparent
            Caption         =   "Cara tetapan upah dari supplier"
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
            Left            =   240
            TabIndex        =   123
            Top             =   1680
            Width           =   4305
         End
         Begin VB.Label Label72 
            BackStyle       =   0  'Transparent
            Caption         =   "Upah Tetap Per Item"
            Height          =   255
            Left            =   480
            TabIndex        =   122
            Top             =   1920
            Width           =   2415
         End
         Begin VB.Label Label74 
            BackStyle       =   0  'Transparent
            Caption         =   "Upah Per Gram (Upah/g) :"
            Height          =   255
            Left            =   480
            TabIndex        =   121
            Top             =   2250
            Width           =   2415
         End
         Begin VB.Label Label65 
            BackStyle       =   0  'Transparent
            Caption         =   "Barang Kemas               Barang Permata"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   480
            TabIndex        =   120
            Top             =   360
            Width           =   3705
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   2
            Height          =   495
            Left            =   120
            Top             =   240
            Width           =   4095
         End
         Begin VB.Shape Shape12 
            BorderWidth     =   2
            Height          =   1095
            Left            =   120
            Top             =   1650
            Width           =   4215
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "** Susut nilai bagi belian barang terpakai"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   4320
            TabIndex        =   119
            Top             =   3600
            Width           =   4065
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "** Harga setelah susut nilai"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   4320
            TabIndex        =   118
            Top             =   3960
            Width           =   4065
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "** Diskaun dalam RM"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   4320
            TabIndex        =   117
            Top             =   4320
            Width           =   4065
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "*** Semua maklumat ini adalah kos belian stok ini dari supplier. ***"
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
            Height          =   525
            Left            =   240
            TabIndex        =   116
            Top             =   5040
            Width           =   9105
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Maklumat Asas Produk"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   94
         Top             =   360
         Width           =   11295
         Begin VB.TextBox TB3 
            Alignment       =   2  'Center
            BackColor       =   &H8000000A&
            Height          =   360
            Left            =   9600
            Locked          =   -1  'True
            TabIndex        =   97
            Top             =   1200
            Width           =   1500
         End
         Begin VB.TextBox TB2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000A&
            Height          =   360
            Left            =   9600
            Locked          =   -1  'True
            TabIndex        =   96
            Top             =   840
            Width           =   1500
         End
         Begin VB.ComboBox CBB3 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1200
            Width           =   5500
         End
         Begin VB.ComboBox CBB2 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   840
            Width           =   5500
         End
         Begin VB.TextBox TB1 
            Alignment       =   2  'Center
            BackColor       =   &H8000000A&
            DataField       =   "Kod_Supplier"
            DataSource      =   "Adodc1"
            Height          =   360
            Left            =   9600
            Locked          =   -1  'True
            TabIndex        =   95
            Top             =   480
            Width           =   1500
         End
         Begin VB.ComboBox CBB1 
            BackColor       =   &H00FFFFFF&
            DataField       =   "Supplier"
            Height          =   360
            ItemData        =   "Frm83.frx":20B62
            Left            =   1920
            List            =   "Frm83.frx":20B64
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   480
            Width           =   5500
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier * :"
            Height          =   255
            Left            =   -400
            TabIndex        =   103
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Kod Supplier * :"
            Height          =   255
            Left            =   7200
            TabIndex        =   102
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label48 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Kategori Produk * :"
            Height          =   255
            Left            =   -400
            TabIndex        =   101
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Purity * :"
            Height          =   255
            Left            =   -400
            TabIndex        =   100
            Top             =   840
            Width           =   2295
         End
         Begin VB.Label Label52 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Kod Purity * :"
            Height          =   255
            Left            =   7200
            TabIndex        =   99
            Top             =   840
            Width           =   2295
         End
         Begin VB.Label Label55 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Kod Kategori Produk * :"
            Height          =   255
            Left            =   7200
            TabIndex        =   98
            Top             =   1200
            Width           =   2295
         End
      End
   End
   Begin VB.PictureBox Pic2 
      BorderStyle     =   0  'None
      Height          =   1485
      Left            =   6000
      ScaleHeight     =   1485
      ScaleWidth      =   8595
      TabIndex        =   80
      Top             =   0
      Visible         =   0   'False
      Width           =   8595
      Begin VB.Label Label122 
         BackStyle       =   0  'Transparent
         Caption         =   "Zero Rated ZR(L)"
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
         Left            =   3960
         TabIndex        =   91
         Top             =   840
         Width           =   1440
      End
      Begin VB.Shape Shape9 
         Height          =   1095
         Left            =   0
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Tanpa GST   : RM"
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
         Left            =   120
         TabIndex        =   90
         Top             =   720
         Width           =   2040
      End
      Begin VB.Label L11_Text 
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
         Height          =   300
         Left            =   1935
         TabIndex        =   89
         Top             =   720
         Width           =   2400
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "Maklumat Harga Belian Keseluruhan         Maklumat GST"
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
         TabIndex        =   88
         Top             =   360
         Width           =   11040
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Standard Rated SR"
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
         Left            =   3960
         TabIndex        =   87
         Top             =   1080
         Width           =   1680
      End
      Begin VB.Shape Shape8 
         Height          =   1095
         Left            =   3840
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label60 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Harga (RM)   Cukai GST  (RM)"
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
         Left            =   5160
         TabIndex        =   86
         Top             =   600
         Width           =   3480
      End
      Begin VB.Label L22_Text 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   5475
         TabIndex        =   85
         Top             =   840
         Width           =   1500
      End
      Begin VB.Label L23_Text 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   6840
         TabIndex        =   84
         Top             =   840
         Width           =   1500
      End
      Begin VB.Label L24_Text 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   5475
         TabIndex        =   83
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Label L25_Text 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   6840
         TabIndex        =   82
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Dengan GST : RM"
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
         Left            =   120
         TabIndex        =   81
         Top             =   960
         Width           =   2040
      End
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
      Left            =   4200
      TabIndex        =   77
      Top             =   10680
      Width           =   200
   End
   Begin VB.Timer Tmr1 
      Interval        =   10
      Left            =   20760
      Top             =   120
   End
   Begin VB.CheckBox CB9 
      Enabled         =   0   'False
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
      Left            =   13800
      TabIndex        =   70
      Top             =   480
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.CheckBox CB10 
      Enabled         =   0   'False
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
      Left            =   13800
      TabIndex        =   69
      Top             =   720
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.TextBox TB34 
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
      Left            =   4200
      TabIndex        =   49
      Top             =   2235
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.CommandButton CMD16 
      BackColor       =   &H000080FF&
      Caption         =   "Upload Image"
      Height          =   300
      Left            =   16920
      MaskColor       =   &H00400000&
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   5070
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.ComboBox CBB6 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   360
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Top             =   11010
      Width           =   4365
   End
   Begin VB.CheckBox CB7 
      Caption         =   "Penerimaan Stok Baru"
      Enabled         =   0   'False
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
      Left            =   10560
      TabIndex        =   48
      Top             =   840
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox CB8 
      Enabled         =   0   'False
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
      Left            =   10560
      TabIndex        =   0
      Top             =   1065
      Visible         =   0   'False
      Width           =   200
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   9255
      Left            =   120
      TabIndex        =   93
      Top             =   120
      Visible         =   0   'False
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   16325
      Arrange         =   2
      LabelEdit       =   1
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   11520
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
            Picture         =   "Frm83.frx":20B66
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm83.frx":23140
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CMD11 
      BackColor       =   &H8000000A&
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
      Left            =   18120
      MouseIcon       =   "Frm83.frx":2571A
      MousePointer    =   99  'Custom
      Picture         =   "Frm83.frx":25A24
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   10440
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton CMD2 
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
      Left            =   18120
      MouseIcon       =   "Frm83.frx":27FEE
      MousePointer    =   99  'Custom
      Picture         =   "Frm83.frx":282F8
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   10440
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton CMD22 
      BackColor       =   &H8000000A&
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
      Left            =   8760
      MouseIcon       =   "Frm83.frx":2A8C2
      MousePointer    =   99  'Custom
      Picture         =   "Frm83.frx":2ABCC
      Style           =   1  'Graphical
      TabIndex        =   164
      Top             =   10320
      Visible         =   0   'False
      Width           =   2000
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
      Left            =   15960
      MouseIcon       =   "Frm83.frx":2D196
      MousePointer    =   99  'Custom
      Picture         =   "Frm83.frx":2D4A0
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   10440
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton CMD10 
      BackColor       =   &H8000000A&
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
      Left            =   15960
      MouseIcon       =   "Frm83.frx":2FA6A
      MousePointer    =   99  'Custom
      Picture         =   "Frm83.frx":2FD74
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   10440
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.Label L12_Text 
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
      Height          =   285
      Left            =   11280
      TabIndex        =   187
      Top             =   12120
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label L101_Text 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bank In * (RM) :"
      Height          =   255
      Left            =   8760
      TabIndex        =   185
      Top             =   10560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label L100_Text 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tunai * (RM) :"
      Height          =   255
      Left            =   8760
      TabIndex        =   184
      Top             =   10200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label L102_Text 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah (RM) :"
      Height          =   255
      Left            =   8760
      TabIndex        =   183
      Top             =   10920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label L36_Text 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3360
      TabIndex        =   170
      Top             =   9360
      Width           =   5385
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nama :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2520
      TabIndex        =   169
      Top             =   9360
      Width           =   705
   End
   Begin VB.Label Label75 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nama :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2520
      TabIndex        =   168
      Top             =   10440
      Width           =   705
   End
   Begin VB.Label L37_Text 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3360
      TabIndex        =   167
      Top             =   10440
      Width           =   4665
   End
   Begin VB.Label L39_Text 
      Caption         =   "L39_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   20640
      TabIndex        =   166
      Top             =   8040
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label L14_Text 
      Caption         =   "Anda berada di dalam menu EDIT DATA."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   1920
      TabIndex        =   163
      Top             =   7560
      Visible         =   0   'False
      Width           =   11715
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai barang yang dimasukkan ke dalam senarai :"
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
      Height          =   405
      Left            =   11880
      TabIndex        =   159
      Top             =   9720
      Visible         =   0   'False
      Width           =   8775
   End
   Begin VB.Label L10_Text 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   17760
      TabIndex        =   158
      Top             =   9240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Sila tandakan ruangan di bawah jika mahukan sistem cetak barcode setelah stok diterima."
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
      Height          =   660
      Left            =   13800
      TabIndex        =   79
      Top             =   960
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Cetak barcode"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   4440
      TabIndex        =   78
      Top             =   10635
      Width           =   2385
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
      Left            =   21480
      TabIndex        =   76
      Top             =   435
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
      Left            =   21495
      TabIndex        =   75
      Top             =   120
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Penerimaan Stok Baru                  Buyback / Trade in"
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
      Height          =   525
      Left            =   10800
      TabIndex        =   74
      Top             =   840
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Label L41_Text 
      Caption         =   "L41_Text"
      Height          =   375
      Left            =   13680
      TabIndex        =   73
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label L40_Text 
      Caption         =   "L40_Text"
      Height          =   375
      Left            =   13680
      TabIndex        =   72
      Top             =   5760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label51 
      BackStyle       =   0  'Transparent
      Caption         =   "Barang Kemas / Permata       Gold Bar"
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
      Height          =   525
      Left            =   14040
      TabIndex        =   71
      Top             =   480
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.Label Label47 
      BackStyle       =   0  'Transparent
      Caption         =   "** Hanya untuk GOLD BAR."
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
      Left            =   5640
      TabIndex        =   68
      Top             =   2280
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.Label Label43 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Certificate"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2040
      TabIndex        =   67
      Top             =   2280
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.Label L38_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "Sila pastikan anda hanya menerima barang dari SATU supplier sahaja dalam satu masa / kemasukkan data ke dalam sistem."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C25418&
      Height          =   525
      Left            =   960
      TabIndex        =   66
      Top             =   6960
      Visible         =   0   'False
      Width           =   11715
   End
   Begin VB.Label L32_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L32_Text"
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
      Left            =   15240
      TabIndex        =   64
      Top             =   5280
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label L31_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L31_Text"
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
      Left            =   15960
      TabIndex        =   63
      Top             =   5040
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label69 
      BackStyle       =   0  'Transparent
      Caption         =   "Image : "
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
      Left            =   15240
      TabIndex        =   62
      Top             =   5040
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label L30_Text 
      Caption         =   "L30_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   15360
      TabIndex        =   61
      Top             =   12600
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Pekerja * :"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   2640
      TabIndex        =   60
      Top             =   11040
      Width           =   2655
   End
   Begin VB.Label L9_Text 
      Caption         =   "L9_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1200
      TabIndex        =   59
      Top             =   12600
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label L7_Text 
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5040
      TabIndex        =   58
      Top             =   12600
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label L6_Text 
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4080
      TabIndex        =   57
      Top             =   12600
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label L5_Text 
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3120
      TabIndex        =   56
      Top             =   12600
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label L4_Text 
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2160
      TabIndex        =   55
      Top             =   12600
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label L3_Text 
      Caption         =   "L3_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   54
      Top             =   12600
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Shape Shape7 
      Height          =   1575
      Left            =   10440
      Top             =   360
      Visible         =   0   'False
      Width           =   3240
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "**Anda hanya dibenarkan untuk membuat pilihan stok baru atau buyback dalam satu masa"
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
      Height          =   645
      Left            =   10560
      TabIndex        =   53
      Top             =   1320
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.Label L13_Text 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7080
      TabIndex        =   52
      Top             =   12600
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label L20_Text 
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8040
      TabIndex        =   51
      Top             =   12600
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label L21_Text 
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9000
      TabIndex        =   50
      Top             =   12600
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Menu Frm83_PM_Menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm83_SM_Edit 
         Caption         =   "Edit Data Ini"
      End
      Begin VB.Menu Frm83_SM_Padam 
         Caption         =   "Padam Data / Keluarkan Dari Senarai"
      End
   End
End
Attribute VB_Name = "Frm83"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CB10_Click()
'On Error Resume Next
If Frm83.CB10 = 1 Then
    Frm83.TB34.Locked = False 'No. Certificate
    Frm83.TB34.BackColor = &HFFFFFF 'No. Certificate
    Frm83.TB4.Locked = True 'Upah
    Frm83.TB4.BackColor = &H8000000A 'Upah
    
    Frm83.CB4.Enabled = False
    Frm83.CB5.Enabled = False
    Frm83.CB4 = 0
    Frm83.CB5 = 0
    Frm83.CB12 = 0
    Frm83.CB12.Enabled = False
End If
End Sub
Private Sub CB11_Click()
'On Error Resume Next
If Frm83.CB11 = 1 Then

    Frm83.CB3 = 0
    Frm83.CB2 = 0
    
End If

Call kiraan_gst_belian

Exit Sub

Dim frm83_LM_KADAR_GST As Double
Dim Frm83_LM_HARGA As Double

Frm83_LM_HARGA = 0

If Frm83.CB11 = 0 Then

    If Frm83.CB3 = 1 And (Frm83.TB20 <> vbNullString And IsNumeric(Frm83.TB20)) And (Frm83.L8_Text <> vbNullString And IsNumeric(Frm83.L8_Text)) Then
        frm83_LM_KADAR_GST = Frm83.L8_Text 'Jumlah Kadar GST (%)
        
        If Frm83.CB12 = 0 Then
            Frm83_LM_HARGA = Frm83.TB20 'Harga (RM)
        ElseIf Frm83.CB12 = 1 Then
            If Frm83.TB4 <> vbNullString And IsNumeric(Frm83.TB4) Then
                Frm83_LM_HARGA = Frm83.TB4 'Upah (RM)
            End If
        End If
        
        Frm83.TB27 = Format((frm83_LM_KADAR_GST / 100) * Frm83_LM_HARGA, "#,##0.00") 'Jumlah Cukai GST (RM)
        Frm83.L40_Text = Format(Frm83_LM_HARGA, "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
    Else
        Frm83.TB27 = Format(0, "#,##0.00") 'Jumlah Cukai GST (RM)
    End If
    
ElseIf Frm83.CB11 = 1 Then

    If Frm83.CB3 = 1 And (Frm83.TB20 <> vbNullString And IsNumeric(Frm83.TB20)) And (Frm83.L8_Text <> vbNullString And IsNumeric(Frm83.L8_Text)) Then
        frm83_LM_KADAR_GST = Frm83.L8_Text 'Jumlah Kadar GST (%)

        If Frm83.CB12 = 0 Then
            Frm83_LM_HARGA = Frm83.TB20 'Harga (RM)
        ElseIf Frm83.CB12 = 1 Then
            If Frm83.TB4 <> vbNullString And IsNumeric(Frm83.TB4) Then
                Frm83_LM_HARGA = Frm83.TB4 'Upah (RM)
            End If
        End If
        
        Frm83.L40_Text = Format(Frm83_LM_HARGA / (1 + (frm83_LM_KADAR_GST / 100)), "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
        Frm83.TB27 = Format(Frm83_LM_HARGA - (Frm83_LM_HARGA / (1 + (frm83_LM_KADAR_GST / 100))), "#,##0.00") 'Jumlah Cukai GST (RM)
    Else
        Frm83.TB27 = "0.00" 'Jumlah Cukai GST (RM)
    End If
    
End If
End Sub
Private Sub CB12_Click()
'On Error Resume Next

Call kiraan_gst_belian

Exit Sub

Dim frm83_LM_KADAR_GST As Double
Dim Frm83_LM_HARGA As Double

Frm83_LM_HARGA = 0

If Frm83.CB11 = 0 Then

    If Frm83.CB3 = 1 And (Frm83.TB20 <> vbNullString And IsNumeric(Frm83.TB20)) And (Frm83.L8_Text <> vbNullString And IsNumeric(Frm83.L8_Text)) Then
        frm83_LM_KADAR_GST = Frm83.L8_Text 'Jumlah Kadar GST (%)

        If Frm83.CB12 = 0 Then
            Frm83_LM_HARGA = Frm83.TB20 'Harga (RM)
        ElseIf Frm83.CB12 = 1 Then
            If Frm83.TB4 <> vbNullString And IsNumeric(Frm83.TB4) Then
                Frm83_LM_HARGA = Frm83.TB4 'Upah (RM)
            End If
        End If
        
        Frm83.TB27 = Format((frm83_LM_KADAR_GST / 100) * Frm83_LM_HARGA, "#,##0.00") 'Jumlah Cukai GST (RM)
        Frm83.L40_Text = Format(Frm83_LM_HARGA, "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
    Else
        Frm83.TB27 = Format(0, "#,##0.00") 'Jumlah Cukai GST (RM)
    End If
    
    If Frm83.CB3 = 0 And (Frm83.TB20 <> vbNullString And IsNumeric(Frm83.TB20)) Then
    
        If Frm83.CB12 = 0 Then
            Frm83_LM_HARGA = Frm83.TB20 'Harga (RM)
        ElseIf Frm83.CB12 = 1 Then
            If Frm83.TB4 <> vbNullString And IsNumeric(Frm83.TB4) Then
                Frm83_LM_HARGA = Frm83.TB4 'Upah (RM)
            End If
        End If
        
        Frm83.L40_Text = Format(Frm83_LM_HARGA, "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
    End If
    
ElseIf Frm83.CB11 = 1 Then

    If Frm83.CB3 = 1 And (Frm83.TB20 <> vbNullString And IsNumeric(Frm83.TB20)) And (Frm83.L8_Text <> vbNullString And IsNumeric(Frm83.L8_Text)) Then
        frm83_LM_KADAR_GST = Frm83.L8_Text 'Jumlah Kadar GST (%)

        If Frm83.CB12 = 0 Then
            Frm83_LM_HARGA = Frm83.TB20 'Harga (RM)
        ElseIf Frm83.CB12 = 1 Then
            If Frm83.TB4 <> vbNullString And IsNumeric(Frm83.TB4) Then
                Frm83_LM_HARGA = Frm83.TB4 'Upah (RM)
            End If
        End If
        
        Frm83.L40_Text = Format(Frm83_LM_HARGA / (1 + (frm83_LM_KADAR_GST / 100)), "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
        Frm83.TB27 = Format(Frm83_LM_HARGA - (Frm83_LM_HARGA / (1 + (frm83_LM_KADAR_GST / 100))), "#,##0.00") 'Jumlah Cukai GST (RM)
    Else
        Frm83.TB27 = "0.00" 'Jumlah Cukai GST (RM)
    End If
    
End If
End Sub
Private Sub CB14_Click()
'On Error Resume Next
If Frm83.CB14 = 1 Then
    Frm83.CB15 = 0
    Frm83.TB35 = "0.00"
    Frm83.TB4 = "0.00"
    
    Frm83.TB35.BackColor = &HFFFFFF
    Frm83.TB4.BackColor = &H8000000A
    Frm83.TB35.Locked = False
    Frm83.TB4.Locked = True
    
    Call Frm83_kira_upah
End If
End Sub
Private Sub CB15_Click()
'On Error Resume Next
If Frm83.CB15 = 1 Then
    Frm83.CB14 = 0
    Frm83.TB35 = "0.00"
    Frm83.TB4 = "0.00"
    
    Frm83.TB4.BackColor = &HFFFFFF
    Frm83.TB35.BackColor = &H8000000A
    Frm83.TB4.Locked = False
    Frm83.TB35.Locked = True
End If
End Sub
Private Sub CB2_Click()
'On Error Resume Next
If Frm83.CB2 = 1 Then

    Frm83.CB3 = 0
    Frm83.CB11 = 0
    
End If

Call kiraan_gst_belian

Exit Sub

If Frm83.CB2 = 1 Then
    Frm83.CB3 = 0
    Frm83.CB11 = 0
    Frm83.TB27 = "0.00" 'Jumlah Cukai GST (RM)
    If Frm83.CB12 = 0 Then
        If IsNumeric(Frm83.TB20) Then
            Frm83.L40_Text = Format(Frm83.TB20, "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
        Else
            Frm83.L40_Text = Format(0, "#,##0.00")
        End If
    ElseIf Frm83.CB12 = 1 Then
        If IsNumeric(Frm83.TB4) Then
            Frm83.L40_Text = Format(Frm83.TB4, "#,##0.00")
        Else
            Frm83.L40_Text = "0.00"
        End If
    End If
End If
End Sub
Private Sub CB3_Click()
'On Error Resume Next
If Frm83.CB3 = 1 Then

    Frm83.CB2 = 0
    Frm83.CB11 = 0
    
End If

Call kiraan_gst_belian

Exit Sub

Dim frm83_LM_KADAR_GST As Double
Dim Frm83_LM_HARGA As Double

If Frm83.CB3 = 1 Then
    Frm83.CB2 = 0
End If
If Frm83.CB3 = 0 Then
    Frm83.CB11 = 0
End If

Frm83_LM_HARGA = 0

If Frm83.CB3 = 1 And (Frm83.TB10 <> vbNullString And IsNumeric(Frm83.TB10)) And (Frm83.L8_Text <> vbNullString And IsNumeric(Frm83.L8_Text)) Then
    If Frm83.CB11 = 0 Then
        frm83_LM_KADAR_GST = Frm83.L8_Text 'Jumlah Kadar GST (%)
        
        If Frm83.CB12 = 0 Then
            Frm83_LM_HARGA = Frm83.TB20 'Harga (RM)
        ElseIf Frm83.CB12 = 1 Then
            If Frm83.TB4 <> vbNullString And IsNumeric(Frm83.TB4) Then
                Frm83_LM_HARGA = Frm83.TB4 'Upah (RM)
            End If
        End If
        
        Frm83.TB27 = Format((frm83_LM_KADAR_GST / 100) * Frm83_LM_HARGA, "#,##0.00") 'Jumlah Cukai GST (RM)
        Frm83.L40_Text = Format(Frm83_LM_HARGA, "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
    ElseIf Frm83.CB11 = 1 Then
        frm83_LM_KADAR_GST = Frm83.L8_Text 'Jumlah Kadar GST (%)

        If Frm83.CB12 = 0 Then
            Frm83_LM_HARGA = Frm83.TB20 'Harga (RM)
        ElseIf Frm83.CB12 = 1 Then
            If Frm83.TB4 <> vbNullString And IsNumeric(Frm83.TB4) Then
                Frm83_LM_HARGA = Frm83.TB4 'Upah (RM)
            End If
        End If
        
        Frm83.L40_Text = Format(Frm83_LM_HARGA / (1 + (frm83_LM_KADAR_GST / 100)), "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
        Frm83.TB27 = Format(Frm83_LM_HARGA - (Frm83_LM_HARGA / (1 + (frm83_LM_KADAR_GST / 100))), "#,##0.00") 'Jumlah Cukai GST (RM)
    End If
Else
    Frm83.TB27 = "0.00" 'Jumlah Cukai GST (RM)
    If Frm83.CB12 = 0 Then
        If IsNumeric(Frm83.TB20) Then
            Frm83.L40_Text = Format(Frm83.TB20, "#,##0.00")
        Else
            Frm83.L40_Text = Format(0, "#,##0.00")
        End If
    Else
        If IsNumeric(Frm83.TB4) Then
            Frm83.L40_Text = Format(Frm83.TB4, "#,##0.00")
        Else
            Frm83.L40_Text = "0.00"
        End If
    End If
End If
End Sub
Private Sub CB4_Click()
'on error resume next
If Frm83.CB4 = 1 Then
    Frm83.CB5 = 0
    
    Frm83.TB8.Locked = False
    Frm83.TB9.Locked = False
    Frm83.TB4.Locked = False
    
    Frm83.TB8.BackColor = &HFFFFFF
    Frm83.TB9.BackColor = &HFFFFFF
    Frm83.TB4.BackColor = &HFFFFFF
    
    Frm83.L27_Text = "Upah Jualan Pelanggan    RM"
    Frm83.L28_Text = "Upah Jualan Ahli               RM"
    Frm83.L29_Text = "Upah Jualan Silver            RM"
    Frm83.L33_Text = "Upah Jualan Gold             RM"
    Frm83.L34_Text = "Upah Jualan Platinum       RM"
    Frm83.L35_Text = "Upah Jualan M.Dealer      RM"
    
    Frm83.TB8 = "0.00"
    Frm83.TB9 = "0.00"
    Frm83.TB4 = 0
    Frm83.TB10 = "0.00"
    
    Frm83.CB14 = 0
    Frm83.CB15 = 0

    Frm83.CB14.Enabled = True
    Frm83.CB15.Enabled = True
    
    Frm83.CB12.Enabled = True
End If
End Sub
Private Sub CB5_Click()
'on error resume next
If Frm83.CB5 = 1 Then
    Frm83.CB4 = 0
    
    Frm83.TB8.Locked = True
    Frm83.TB9.Locked = True
    Frm83.TB4.Locked = True
    
    Frm83.TB8.BackColor = &H8000000A
    Frm83.TB9.BackColor = &H8000000A
    Frm83.TB4.BackColor = &H8000000A
    
    Frm83.L27_Text = "Harga Jualan Pelanggan   RM"
    Frm83.L28_Text = "Harga Jualan Ahli              RM"
    Frm83.L29_Text = "Harga Jualan Silver           RM"
    Frm83.L33_Text = "Harga Jualan Gold            RM"
    Frm83.L34_Text = "Harga Jualan Platinum      RM"
    
    Frm83.TB8 = vbNullString
    Frm83.TB9 = vbNullString
    Frm83.TB4 = vbNullString
    Frm83.TB10 = "0.00"
    
    Frm83.CB14 = 0
    Frm83.CB15 = 0
    Frm83.TB35 = vbNullString
    
    Frm83.CB14.Enabled = False
    Frm83.CB15.Enabled = False
    
    Frm83.TB35.BackColor = &H8000000A
    Frm83.TB35.Locked = True
    Frm83.TB4.Locked = True
    
    
    Frm83.CB12.Enabled = False
    Frm83.CB12 = 0
End If
End Sub
Private Sub CB7_Click()
'on error resume next
If Frm83.CB7 = 1 Then
    Frm83.CB8 = 0
    
    Frm83.TB19.BackColor = &H8000000A
    Frm83.TB19.Locked = True
    
End If
End Sub
Private Sub CB8_Click()
'on error resume next
If Frm83.CB8 = 1 Then
    Frm83.CB7 = 0
    
    Frm83.TB19.BackColor = &HFFFFFF
    Frm83.TB19.Locked = False
    
End If
End Sub
Private Sub CB9_Click()
'On Error Resume Next
If Frm83.CB9 = 1 Then
    Frm83.TB4.Locked = False 'Upah
    Frm83.TB4.BackColor = &HFFFFFF 'Upah
    Frm83.TB34.Locked = True 'No. Certificate
    Frm83.TB34.BackColor = &H8000000A 'No. Certificate
    
    Frm83.CB4.Enabled = True
    Frm83.CB5.Enabled = True
    Frm83.CB4 = 1
End If
End Sub
Private Sub CBB1_Change()
'On Error Resume Next
If GLOBAL_DISABLE <> 1 Then
    Frm83.TB1 = vbNullString
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from setting_database where Supplier='" & Frm83.CBB1 & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!Kod_Supplier) Then Frm83.TB1 = rs!Kod_Supplier
        If Not IsNull(rs!ID) Then
            Frm83.L4_Text = rs!ID 'ID
        End If
        
        If Not IsNull(rs!no_id_gst) Then
            Frm83.TB28 = rs!no_id_gst 'No. ID GST
            Frm83.CB2.Enabled = True
            Frm83.CB3.Enabled = True
            Frm83.CB11.Enabled = True
            If Frm83.CB9 = 1 Then
                Frm83.CB12.Enabled = True
            End If
        Else
            Frm83.TB28 = vbNullString
            Frm83.CB2 = 1
            Frm83.CB2.Enabled = False
            Frm83.CB3.Enabled = False
            Frm83.CB11.Enabled = False
            Frm83.CB12.Enabled = False
        End If
    
    End If
    
    rs.Close
    Set rs = Nothing

End If
End Sub
Private Sub CBB1_Click()
'On Error Resume Next
If GLOBAL_DISABLE <> 1 Then
    Frm83.TB1 = vbNullString
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from setting_database where Supplier='" & Frm83.CBB1 & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        If Not IsNull(rs!Kod_Supplier) Then Frm83.TB1 = rs!Kod_Supplier
        If Not IsNull(rs!ID) Then
            Frm83.L4_Text = rs!ID 'ID
        End If
        
        If Not IsNull(rs!no_id_gst) Then
            Frm83.TB28 = rs!no_id_gst 'No. ID GST
            Frm83.CB2.Enabled = True
            Frm83.CB3.Enabled = True
            Frm83.CB11.Enabled = True
            If Frm83.CB9 = 1 Then
                Frm83.CB12.Enabled = True
            End If
        Else
            Frm83.TB28 = vbNullString
            Frm83.CB2 = 1
            Frm83.CB2.Enabled = False
            Frm83.CB3.Enabled = False
            Frm83.CB11.Enabled = False
            Frm83.CB12.Enabled = False
        End If
        
    End If
    
    rs.Close
    Set rs = Nothing
End If
End Sub
Private Sub CBB2_Change()
'On Error Resume Next
If GLOBAL_DISABLE <> 1 Then
    DATA_FOUND = 0
    Frm83.TB2 = vbNullString
    Frm83.TB9 = vbNullString
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from setting_database where Metal_Purity='" & Frm83.CBB2 & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!Kod_Metal_Purity) Then
            Frm83.TB2 = rs!Kod_Metal_Purity
            DATA_FOUND = 1
        End If
        If Not IsNull(rs!ID) Then
            Frm83.L5_Text = rs!ID 'ID
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    
    If DATA_FOUND = 1 And (Frm83.CB4 = 1 Or Frm83.CB10 = 1) Then
        Set rs1 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs1.Open "select * from hargaemas where Purity='" & Frm83.TB2 & "' AND cawangan='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs1.EOF Then
            If Frm83.CB7 = 1 Then
                If Not IsNull(rs1!HargaDariSupplier) Then Frm83.TB9 = Format(rs1!HargaDariSupplier, "0.00") 'Harga Dari Supplier
            ElseIf Frm83.CB8 = 1 Then
                If Not IsNull(rs1!Harga_Pelanggan) Then Frm83.TB9 = Format(rs1!Harga_Pelanggan, "0.00") 'Harga Semasa
            End If
        End If
        
        rs1.Close
        Set rs1 = Nothing
    End If
End If
End Sub
Private Sub CBB2_Click()
'On Error Resume Next
If GLOBAL_DISABLE <> 1 Then
    DATA_FOUND = 0
    Frm83.TB2 = vbNullString
    Frm83.TB9 = vbNullString
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from setting_database where Metal_Purity='" & Frm83.CBB2 & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!Kod_Metal_Purity) Then
            Frm83.TB2 = rs!Kod_Metal_Purity
            DATA_FOUND = 1
        End If
        If Not IsNull(rs!ID) Then
            Frm83.L5_Text = rs!ID 'ID
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    
    If DATA_FOUND = 1 And (Frm83.CB4 = 1 Or Frm83.CB10 = 1) Then
        Set rs1 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs1.Open "select * from hargaemas where Purity='" & Frm83.TB2 & "' AND cawangan='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs1.EOF Then
            If Frm83.CB7 = 1 Then
                If Not IsNull(rs1!HargaDariSupplier) Then Frm83.TB9 = Format(rs1!HargaDariSupplier, "0.00") 'Harga Dari Supplier
            ElseIf Frm83.CB8 = 1 Then
                If Not IsNull(rs1!Harga_Pelanggan) Then Frm83.TB9 = Format(rs1!Harga_Pelanggan, "0.00") 'Harga Semasa
            End If
        End If
        
        rs1.Close
        Set rs1 = Nothing
    End If
End If
End Sub
Private Sub CBB3_Change()
'On Error Resume Next
If GLOBAL_DISABLE <> 1 Then
    Frm83.TB3 = vbNullString
    Frm83.TB6 = vbNullString
        
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from setting_database where Kategori_Produk='" & Frm83.CBB3 & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!Kod_Kategori_Produk) Then
            Frm83.TB3 = rs!Kod_Kategori_Produk
            Frm83.TB6 = rs!Kod_Kategori_Produk
        End If
        If Not IsNull(rs!ID) Then
            Frm83.L6_Text = rs!ID 'ID
        End If
    End If
    
    rs.Close
    Set rs = Nothing
End If
End Sub
Private Sub CBB3_Click()
'On Error Resume Next
If GLOBAL_DISABLE <> 1 Then
    Frm83.TB3 = vbNullString
    Frm83.TB6 = vbNullString
        
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from setting_database where Kategori_Produk='" & Frm83.CBB3 & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!Kod_Kategori_Produk) Then
            Frm83.TB3 = rs!Kod_Kategori_Produk
            Frm83.TB6 = rs!Kod_Kategori_Produk
        End If
        If Not IsNull(rs!ID) Then
            Frm83.L6_Text = rs!ID 'ID
        End If
    End If
    
    rs.Close
    Set rs = Nothing
End If
End Sub

Private Sub CMD1_Click()
'On Error Resume Next
Dim rs2 As ADODB.Recordset

Dim Err(35)
Dim Frm83_LM_BERAT As Double
Dim Frm83_LM_HARGA_TOTAL As Double
Dim Frm83_LM_CUKAI_GST As Double
Dim Frm83_LM_HARGA_SEMASA As Double 'Harga Semasa
Dim Frm83_LM_ADJUSTMENT As Double 'Adjustment
Dim Frm83_LM_UPAH As Double 'Upah

Frm83_LM_HARGA_SEMASA = 0 'Harga Semasa
Frm83_LM_ADJUSTMENT = 0 'Adjustment
Frm83_LM_CUKAI_GST = 0
Frm83_LM_HARGA_TOTAL = 0
Frm83_LM_BERAT = 1
Frm83_LM_UPAH = 0 'Upah

x = 0
DATA_SAVE = 0

If Frm83.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Supplier]."
End If
If Frm83.CBB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Purity]."
End If
If Frm83.CBB3 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Kategori Produk]."
End If
If Frm83.TB1 = vbNullString Then
    x = x + 1
    Err(x) = "Tiada Maklumat [Kod Supplier]."
End If
If Frm83.TB2 = vbNullString Then
    x = x + 1
    Err(x) = "Tiada Maklumat [Kod Purity]."
End If
If Frm83.TB3 = vbNullString Then
    x = x + 1
    Err(x) = "Tiada Maklumat [Kod Kategori Produk]."
End If
'If Frm83.TB6 = vbNullString Or Frm83.TB7 = vbNullString Then
'    x = x + 1
'    Err(x) = "Maklumat [No. Siri Produk] Yang Tidak Lengkap."
'End If
If Frm83.CB9 = 1 And Frm83.CB4 = 0 And Frm83.CB5 = 0 Then
    x = x + 1
    Err(x) = "Sila Buat Pilihan Penerimaan [Barang Kemas] Atau [Barang Permata]."
End If
If Frm83.TB10 = vbNullString Or (Frm83.TB10 <> vbNullString And Not IsNumeric(Frm83.TB10)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Harga Asal]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB19 = vbNullString Or (Frm83.TB19 <> vbNullString And Not IsNumeric(Frm83.TB19)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Spread (%)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB20 = vbNullString Or (Frm83.TB20 <> vbNullString And Not IsNumeric(Frm83.TB20)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Harga Belian]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB21 = vbNullString Or (Frm83.TB21 <> vbNullString And Not IsNumeric(Frm83.TB21)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Harga Asal-Spread]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB22 = vbNullString Or (Frm83.TB22 <> vbNullString And Not IsNumeric(Frm83.TB22)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Adjusment]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.CB2 = 0 And Frm83.CB3 = 0 And Frm83.CB11 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan jenis cukai GST"
End If
If Frm83.TB36 <> vbNullString Then

    If InStr(1, Frm83.TB36, "*") <> 0 Or InStr(1, Frm83.TB36, "/") <> 0 Or InStr(1, Frm83.TB36, "\") <> 0 Or InStr(1, Frm83.TB36, "'") <> 0 Then

        x = x + 1
        Err(x) = "[Code 1] mengandungi simbol yang tidak sah."
        
    End If
    
End If
If Frm83.TB37 <> vbNullString Then

    If InStr(1, Frm83.TB37, "*") <> 0 Or InStr(1, Frm83.TB37, "/") <> 0 Or InStr(1, Frm83.TB37, "\") <> 0 Or InStr(1, Frm83.TB37, "'") <> 0 Then

        x = x + 1
        Err(x) = "[Code 2] mengandungi simbol yang tidak sah."
        
    End If
    
End If
If Frm83.CB4 = 1 Then
    If Frm83.TB8 = vbNullString Or (Frm83.TB8 <> vbNullString And Not IsNumeric(Frm83.TB8)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Berat]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB8 <> vbNullString Or (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) Then
        
        Frm83_LM_BERAT = Frm83.TB8
        
        If Frm83_LM_BERAT = 0 Then
            x = x + 1
            Err(x) = "Tiada data bagi berat barang. Berat 0 tidak dibenarkan."
        End If
    End If
    If Frm83.TB9 = vbNullString Or (Frm83.TB9 <> vbNullString And Not IsNumeric(Frm83.TB9)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Harga Semasa]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB4 = vbNullString Or (Frm83.TB4 <> vbNullString And Not IsNumeric(Frm83.TB4)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Upah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB24 = vbNullString Or (Frm83.TB24 <> vbNullString And Not IsNumeric(Frm83.TB24)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Pelanggan]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB25 = vbNullString Or (Frm83.TB25 <> vbNullString And Not IsNumeric(Frm83.TB25)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Ahli]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB26 = vbNullString Or (Frm83.TB26 <> vbNullString And Not IsNumeric(Frm83.TB26)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Silver]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB31 = vbNullString Or (Frm83.TB31 <> vbNullString And Not IsNumeric(Frm83.TB31)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Gold]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB32 = vbNullString Or (Frm83.TB32 <> vbNullString And Not IsNumeric(Frm83.TB32)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Platinum]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB33 = vbNullString Or (Frm83.TB33 <> vbNullString And Not IsNumeric(Frm83.TB33)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada M.Dealer]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
ElseIf Frm83.CB5 = 1 Then
    '+++++++++++ Special Request ++++++++++ Start
    'If Frm83.TB8 = vbNullString Or (Frm83.TB8 <> vbNullString And Not IsNumeric(Frm83.TB8)) Then
    '    x = x + 1
    '    Err(x) = "Sila Masukkan [Berat]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    'End If
    'If Frm83.TB8 <> vbNullString Or (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) Then
        
    '    Frm83_LM_BERAT = Frm83.TB8
        
    '    If Frm83_LM_BERAT = 0 Then
    '        x = x + 1
    '        Err(x) = "Tiada data bagi berat barang. Berat 0 tidak dibenarkan."
    '    End If
    'End If
    '+++++++++++ Special Request ++++++++++ End
    If Frm83.TB24 = vbNullString Or (Frm83.TB24 <> vbNullString And Not IsNumeric(Frm83.TB24)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Pelanggan]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB25 = vbNullString Or (Frm83.TB25 <> vbNullString And Not IsNumeric(Frm83.TB25)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Ahli]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB26 = vbNullString Or (Frm83.TB26 <> vbNullString And Not IsNumeric(Frm83.TB26)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Silver]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB31 = vbNullString Or (Frm83.TB31 <> vbNullString And Not IsNumeric(Frm83.TB31)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Gold]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB32 = vbNullString Or (Frm83.TB32 <> vbNullString And Not IsNumeric(Frm83.TB32)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Platinum]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB33 = vbNullString Or (Frm83.TB33 <> vbNullString And Not IsNumeric(Frm83.TB33)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada M.Dealer]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If

If Frm83.CB8 = 1 Then
    If Frm83.TB19 = vbNullString Or (Frm83.TB19 <> vbNullString And Not IsNumeric(Frm83.TB19)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Spread]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If
If Frm83.CBB5 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Dulang]."
End If
'If Frm83.CB3 = 1 Then
    If Frm83.TB27 = vbNullString Or (Frm83.TB27 <> vbNullString And Not IsNumeric(Frm83.TB27)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat GST"
    End If
'End If
If Frm83.CB4 = 1 Then
    If Frm83.CB14 = 0 And Frm83.CB15 = 1 Then
        If Frm83.TB35 = vbNullString Or (Frm83.TB35 <> vbNullString And Not IsNumeric(Frm83.TB35)) Then
            x = x + 1
            Err(x) = "Sila buat tetapan pengiraan upah dari supplier"
        End If
    End If
End If
If Frm83.CB14 = 1 Then
    If Frm83.TB35 = vbNullString Or (Frm83.TB35 <> vbNullString And Not IsNumeric(Frm83.TB35)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Upah per gram]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If
If Frm83.TB15 <> vbNullString Then

    If InStr(1, Frm83.TB15, "*") <> 0 Or InStr(1, Frm83.TB15, "/") <> 0 Or InStr(1, Frm83.TB15, "\") <> 0 Or InStr(1, Frm83.TB15, "'") <> 0 Then

        x = x + 1
        Err(x) = "[No. Invoice] mengandungi simbol yang tidak sah."
        
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
    Note = "Adakah anda ingin masukkan data barang ini ke dalam senarai belian?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then

        Frm83_LM_No_SIRI = Frm83.L3_Text 'No. Turutan No. Siri
        Frm83_LM_No_RUJUKAN_BELIAN = Frm83.L9_Text 'No. Rujukan Belian
        
'Re_Gen_Code:
        
'        Set rs = New ADODB.Recordset
'        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'        If Frm83.CB9 = 1 Then rs.Open "select * from Data_Database where barcode='" & Format(Frm83_LM_No_SIRI, "000000") & "'", cn, adOpenKeyset, adLockOptimistic
'        If Frm83.CB10 = 1 Then rs.Open "select * from Data_Database where barcode='" & Format(Frm83_LM_No_SIRI, "000000") & "W" & "'", cn, adOpenKeyset, adLockOptimistic
        
'        If Not rs.EOF Then
'            Frm83_LM_No_SIRI = Frm83_LM_No_SIRI + 1
            
'            rs.Close
'            Set rs = Nothing
'            GoTo Re_Gen_Code:
'        End If
        
'        rs.Close
'        Set rs = Nothing
        
'###Masukkan Data Belian Ke Dalam Database### - Start
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_BELIAN_TEMP & "", cn, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        If Frm83.L4_Text <> vbNullString Then
            rs!supplier_ID = Frm83.L4_Text 'No. ID Bagi Supplier
        Else
            rs!supplier_ID = Null 'No. ID Bagi Supplier
        End If
        If Frm83.CBB1 <> vbNullString Then
            rs!nama_Supplier = Frm83.CBB1 'Nama Supplier
        Else
            rs!nama_Supplier = Null 'Nama Supplier
        End If
        If Frm83.TB1 <> vbNullString Then
            rs!Kod_Supplier = Frm83.TB1 'Kod Supplier
        Else
            rs!Kod_Supplier = Null 'Kod Supplier
        End If
        If Frm83.L5_Text <> vbNullString Then
            rs!purity_ID = Frm83.L5_Text 'No. ID Bagi Purity
        Else
            rs!purity_ID = Null 'No. ID Bagi Purity
        End If
        If Frm83.CBB2 <> vbNullString Then
            rs!purity = Frm83.CBB2 'Purity
        Else
            rs!purity = Null 'Purity
        End If
        If Frm83.TB2 <> vbNullString Then
            rs!kod_Purity = Frm83.TB2 'Kod Purity
        Else
            rs!kod_Purity = Null 'Kod Purity
        End If
        If Frm83.L6_Text <> vbNullString Then
            rs!kategori_produk_ID = Frm83.L6_Text 'No. ID Bagi Kategori Produk
        Else
            rs!kategori_produk_ID = Null 'No. ID Bagi Kategori Produk
        End If
        If Frm83.CBB3 <> vbNullString Then
            rs!kategori_Produk = Frm83.CBB3 'Kategori Produk
        Else
            rs!kategori_Produk = Null 'Kategori Produk
        End If
        If Frm83.TB3 <> vbNullString Then
            rs!Kod_Kategori_Produk = Frm83.TB3 'Kod Kategori Produk
        Else
            rs!Kod_Kategori_Produk = Null 'Kod Kategori Produk
        End If
        'If Frm83.TB7 <> vbNullString Then
        '    If Frm83.CB9 = 1 Then
        '        rs!Barcode = Format(Frm83_LM_No_SIRI, "000000") 'No. Barcode (6 Digit Terakhir)
        '    ElseIf Frm83.CB10 = 1 Then
        '        rs!Barcode = Format(Frm83_LM_No_SIRI, "000000") & "W" 'No. Barcode (6 Digit Terakhir)
        '    End If
        'Else
        '    rs!Barcode = Null 'No. Barcode (6 Digit Terakhir)
        'End If
        'If Frm83.CB9 = 1 Then
        '    rs!no_siri_Produk = Frm83.TB6 & Format(Frm83_LM_No_SIRI, "000000")  'No. Siri Produk
        'ElseIf Frm83.CB10 = 1 Then
        '    rs!no_siri_Produk = Frm83.TB6 & Format(Frm83_LM_No_SIRI, "000000") & "W"  'No. Siri Produk
        'End If
        If Frm83.CB12 = 0 Then '0 : Pengiraan GST pada harga belian barang , 1 : Pengiraan GST pada upah sahaja
            rs!gst_barang_atau_upah = 0
        ElseIf Frm83.CB12 = 1 Then
            rs!gst_barang_atau_upah = 1
        End If
        If Frm83.CB4 = 1 Or Frm83.CB10 = 1 Then
            If Frm83.TB8 <> vbNullString Then
                rs!Berat = Format(Frm83.TB8, "0.00") 'Berat
                rs!beza_berat = Format(Frm83.TB8, "0.00") 'Baki Berat
            Else
                rs!Berat = Null 'Berat
                rs!beza_berat = Null 'Baki Berat
            End If
            If Frm83.TB29 <> vbNullString Then
                rs!riyal = Format(Frm83.TB29, "0.00") 'Berat Riyal
            Else
                rs!riyal = Null 'Berat Riyal
            End If
            If Frm83.TB9 <> vbNullString Then
                rs!kos_Belian_Gram = Format(Frm83.TB9, "0.00") 'Harga Per Gram (Belian)
            Else
                rs!kos_Belian_Gram = Null 'Harga Per Gram (Belian)
            End If
            If Frm83.TB4 <> vbNullString Then
                rs!UPAH = Frm83.TB4 'Upah (RM)
            Else
                rs!UPAH = Null 'Upah (RM)
            End If
            If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) And (Frm83.TB22 <> vbNullString And IsNumeric(Frm83.TB22)) Then
                Frm83_LM_BERAT = Frm83.TB8 'Berat
                Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment

                If Frm83.CB12 = 0 Then
                    rs!harga_Per_Gram_Item = Format((Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST) / Frm83_LM_BERAT, "0.00") 'Harga Modal Per Gram [Average] (Harga Emas + Upah) Tanpa Tax
                    rs!harga_per_gram_tanpa_gst = Format(Frm83_LM_HARGA_TOTAL / Frm83_LM_BERAT, "0.00")
                ElseIf Frm83.CB12 = 1 Then
                    rs!harga_Per_Gram_Item = Format((((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST) / Frm83_LM_BERAT, "0.00")
                    rs!harga_per_gram_tanpa_gst = Format((((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL) / Frm83_LM_BERAT, "0.00")
                End If
            Else
                rs!harga_Per_Gram_Item = Null
            End If
            If Frm83.TB24 <> vbNullString Then
                rs!Upah_Jualan = Format(Frm83.TB24, "0.00") 'Upah Jualan Kepada Pelanggan
            Else
                rs!Upah_Jualan = Null 'Upah Jualan Kepada Pelanggan
            End If
            If Frm83.TB25 <> vbNullString Then
                rs!Upah_Member = Format(Frm83.TB25, "0.00") 'Upah Jualan Kepada Ahli / Member
            Else
                rs!Upah_Member = Null 'Upah Jualan Kepada Ahli / Member
            End If
            If Frm83.TB26 <> vbNullString Then
                rs!Upah_Pengedar = Format(Frm83.TB26, "0.00") 'Upah Jualan Kepada Pengedar
            Else
                rs!Upah_Pengedar = Null 'Upah Jualan Kepada Pengedar
            End If
            If Frm83.TB31 <> vbNullString Then
                rs!Upah_RAF = Format(Frm83.TB31, "0.00") 'Upah Jualan Kepada RAF
            Else
                rs!Upah_RAF = Null 'Upah Jualan Kepada RAF
            End If
            If Frm83.TB32 <> vbNullString Then
                rs!upah_normal_dealer = Format(Frm83.TB32, "0.00") 'Upah Jualan Kepada N.Dealer
            Else
                rs!upah_normal_dealer = Null 'Upah Jualan Kepada N.Dealer
            End If
            If Frm83.TB33 <> vbNullString Then
                rs!upah_master_dealer = Format(Frm83.TB33, "0.00") 'Upah Jualan Kepada M.Dealer
            Else
                rs!upah_master_dealer = Null 'Upah Jualan Kepada M.Dealer
            End If
            rs!code_Supplier = Null 'Harga Jualan Kepada Pelanggan
            rs!HargaJualan_Member = Null 'Harga Jualan Kepada Ahli / Member
            rs!HargaJualan_Pengedar = Null 'Harga Jualan Kepada Pengedar
            rs!HargaJualan_RAF = Null 'Harga Jualan Kepada RAF
            rs!hargajualan_normal_dealer = Null 'Harga Jualan Kepada N.Dealer
            rs!hargajualan_master_dealer = Null 'Harga Jualan Kepada M.Dealer
        Else
            '+++++++++++ Special Request ++++++++++ Start
            'rs!Berat = Null 'Berat
            'rs!beza_berat = Null 'Baki Berat
            'If Frm83.TB8 <> vbNullString Then
            '    rs!Berat = Format(Frm83.TB8, "0.00") 'Berat
            '    rs!beza_berat = Format(Frm83.TB8, "0.00") 'Baki Berat
            'Else
                rs!Berat = Null 'Berat
                rs!beza_berat = Null 'Baki Berat
            'End If
            '+++++++++++ Special Request ++++++++++ End
            rs!kos_Belian_Gram = Null 'Harga Per Gram (Belian)
            rs!UPAH = Null 'Upah (RM)
            rs!harga_Per_Gram_Item = Null 'Harga Modal Per Gram [Average] (Harga Emas + Upah) Tanpa Tax
            rs!Upah_Jualan = Null 'Upah Jualan Kepada Pelanggan
            rs!Upah_Member = Null 'Upah Jualan Kepada Ahli / Member
            rs!Upah_Pengedar = Null 'Upah Jualan Kepada Pengedar
            rs!Upah_RAF = Null 'Upah Jualan Kepada RAF
            rs!upah_normal_dealer = Null 'Upah Jualan Kepada N.Dealer
            rs!upah_master_dealer = Null 'Upah Jualan Kepada M.Dealer
        End If
        If Frm83.CB5 = 1 Then
            If Frm83.TB24 <> vbNullString Then
                rs!code_Supplier = Format(Frm83.TB24, "0.00") 'Harga Jualan Kepada Pelanggan
            Else
                rs!code_Supplier = Null 'Harga Jualan Kepada Pelanggan
            End If
            If Frm83.TB25 <> vbNullString Then
                rs!HargaJualan_Member = Format(Frm83.TB25, "0.00") 'Harga Jualan Kepada Ahli / Member
            Else
                rs!HargaJualan_Member = Null 'Harga Jualan Kepada Ahli / Member
            End If
            If Frm83.TB26 <> vbNullString Then
                rs!HargaJualan_Pengedar = Format(Frm83.TB26, "0.00") 'Harga Jualan Kepada Pengedar
            Else
                rs!HargaJualan_Pengedar = Null 'Harga Jualan Kepada Pengedar
            End If
            If Frm83.TB31 <> vbNullString Then
                rs!HargaJualan_RAF = Format(Frm83.TB31, "0.00") 'Harga Jualan Kepada RAF
            Else
                rs!HargaJualan_RAF = Null 'Harga Jualan Kepada RAF
            End If
            If Frm83.TB32 <> vbNullString Then
                rs!hargajualan_normal_dealer = Format(Frm83.TB32, "0.00") 'Harga Jualan Kepada N.Dealer
            Else
                rs!hargajualan_normal_dealer = Null 'Harga Jualan Kepada N.Dealer
            End If
            If Frm83.TB33 <> vbNullString Then
                rs!hargajualan_master_dealer = Format(Frm83.TB33, "0.00") 'Harga Jualan Kepada M.Dealer
            Else
                rs!hargajualan_master_dealer = Null 'Harga Jualan Kepada M.Dealer
            End If
            
            '+++++++++++ Special Request ++++++++++ Start
            'rs!Berat = Null 'Berat
            'rs!beza_berat = Null 'Baki Berat
            'If Frm83.TB8 <> vbNullString Then
            '    rs!Berat = Format(Frm83.TB8, "0.00") 'Berat
            '    rs!beza_berat = Format(Frm83.TB8, "0.00") 'Baki Berat
            'Else
                rs!Berat = Null 'Berat
                rs!beza_berat = Null 'Baki Berat
            'End If
            '+++++++++++ Special Request ++++++++++ End
            rs!kos_Belian_Gram = Null 'Harga Per Gram (Belian)
            rs!UPAH = Null 'Upah (RM)
            rs!harga_Per_Gram_Item = Null 'Harga Modal Per Gram [Average] (Harga Emas + Upah) Tanpa Tax
            rs!Upah_Jualan = Null 'Upah Jualan Kepada Pelanggan
            rs!Upah_Member = Null 'Upah Jualan Kepada Ahli / Member
            rs!Upah_Pengedar = Null 'Upah Jualan Kepada Pengedar
            rs!Upah_RAF = Null 'Upah Jualan Kepada RAF
            rs!upah_normal_dealer = Null 'Upah Jualan Kepada N.Dealer
            rs!upah_master_dealer = Null 'Upah Jualan Kepada M.Dealer
        Else
            rs!code_Supplier = Null 'Harga Jualan Kepada Pelanggan
            rs!HargaJualan_Member = Null 'Harga Jualan Kepada Ahli / Member
            rs!HargaJualan_Pengedar = Null 'Harga Jualan Kepada Pengedar
            rs!HargaJualan_RAF = Null 'Harga Jualan Kepada RAF
            rs!hargajualan_normal_dealer = Null 'Harga Jualan Kepada N.Dealer
            rs!hargajualan_master_dealer = Null 'Harga Jualan Kepada M.Dealer
        End If
        If Frm83.CB12 = 0 Then 'GST pada harga barang
        
            If Frm83.TB10 <> vbNullString Then
                rs!kos_Belian_Item = Format(Frm83.TB10, "0.00") 'Harga Asal (RM)
            Else
                rs!kos_Belian_Item = Null 'Harga Asal (RM)
            End If
            
        End If
        If Frm83.CB12 = 1 Then 'GST pada upah
        
            If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) Then
                
                Frm83_LM_BERAT = Frm83.TB8 'Berat
                Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment
                
                rs!kos_Belian_Item = Format((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) + Frm83_LM_HARGA_TOTAL, "0.00") 'Harga Asal (RM)
                
            End If
        
        End If
        If Frm83.CB8 = 1 Then
            If Frm83.TB19 <> vbNullString Then
                rs!Spread = Format(Frm83.TB19, "0.00") 'Spread (%)
            Else
                rs!Spread = Null 'Spread (%)
            End If
        ElseIf Frm83.CB7 = 1 Then
            rs!Spread = Null 'Spread (%)
        End If
        If Frm83.TB21 <> vbNullString Then
            rs!harga_lepas_spread = Format(Frm83.TB21, "0.00") 'Harga asal ditolak spread (RM)
        Else
            rs!harga_lepas_spread = Null 'Harga asal ditolak spread (RM)
        End If
        If Frm83.TB22 <> vbNullString Then
            rs!adjustment = Format(Frm83.TB22, "0.00") 'Adjustment (RM)
        Else
            rs!adjustment = Null 'Adjustment (RM)
        End If
        If Frm83.CB12 = 0 Then 'GST pada harga barang
            If Frm83.TB20 <> vbNullString Then
                rs!kos_item_tanpa_tax = Format(Frm83.TB20, "0.00") 'Harga Barang + Upah Tanpa Tax
            Else
                rs!kos_item_tanpa_tax = Null 'Harga Barang + Upah Tanpa Tax
            End If
        End If
        If Frm83.CB12 = 1 Then 'GST pada upah
        
            If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) And (Frm83.TB22 <> vbNullString And IsNumeric(Frm83.TB22)) Then
                Frm83_LM_BERAT = Frm83.TB8 'Berat
                Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment
                
                rs!kos_item_tanpa_tax = Format(((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL, "0.00")  'Harga Barang + Upah Tanpa Tax
                
            End If
            
        End If
        If Frm83.TB12 <> vbNullString Then
            rs!dimension_Panjang = Frm83.TB12 'Panjang
        Else
            rs!dimension_Panjang = Null 'Panjang
        End If
        If Frm83.TB13 <> vbNullString Then
            rs!dimension_Lebar = Frm83.TB13 'Lebar
        Else
            rs!dimension_Lebar = Null 'Lebar
        End If
        If Frm83.TB14 <> vbNullString Then
            rs!dimension_Saiz = Frm83.TB14 'Saiz
        Else
            rs!dimension_Saiz = Null 'Saiz
        End If
        If Frm83.TB36 <> vbNullString Then 'Code 1
            rs!code1 = UCase(Frm83.TB36)
        Else
            rs!code1 = Null
        End If
        If Frm83.TB37 <> vbNullString Then 'Code 2
            rs!code2 = UCase(Frm83.TB37)
        Else
            rs!code2 = Null
        End If
        If Frm83.CBB5 <> vbNullString Then
            rs!dulang = Frm83.CBB5 'Dulang
        Else
            rs!dulang = Null 'Dulang
        End If
        If Frm83.TB16 <> vbNullString Then
            rs!remarks = UCase(Frm83.TB16) 'Remarks
        Else
            rs!remarks = Null 'Remarks
        End If
        If Frm83.TB34 <> vbNullString Then
            rs!no_cert = UCase(Frm83.TB34) 'No. Cert
        Else
            rs!no_cert = Null 'No. Cert
        End If
        
        If Frm83.CB2 = 1 Then
        
            rs!gst_ari_nashi = 0 'Status Cukai GST : 0 : ZR(L) , 1 : SR
            rs!kadar_gst = Null 'Kadar GST (%)
            rs!jumlah_gst = Null 'Jumlah Cukai GST (RM)
            rs!gst_included = Null '0 : Harga tidak termasuk GST , 1 : Harga termasuk GST
            
        ElseIf Frm83.CB3 = 1 Then
        
            rs!gst_included = 0 '0 : Harga tidak termasuk GST , 1 : Harga termasuk GST
            rs!gst_ari_nashi = 1 'Status Cukai GST : 0 : ZR(L) , 1 : SR
            rs!kadar_gst = Format(Frm83.L8_Text, "0.00") 'Kadar GST (%)
            rs!jumlah_gst = Format(Frm83.TB27, "0.00") 'Jumlah Cukai GST (RM)
          
        ElseIf Frm83.CB11 = 1 Then

            rs!gst_included = 1 '0 : Harga tidak termasuk GST , 1 : Harga termasuk GST
            rs!gst_ari_nashi = 1 'Status Cukai GST : 0 : ZR(L) , 1 : SR
            rs!kadar_gst = Format(Frm83.L8_Text, "0.00") 'Kadar GST (%)
            rs!jumlah_gst = Format(Frm83.TB27, "0.00") 'Jumlah Cukai GST (RM)
            
        End If
        
        If Frm83.L40_Text <> vbNullString Then
            rs!harga_tanpa_gst = Format(Frm83.L40_Text, "0.00") 'Harga Barang Tanpa Tax (kalau gst included)
        Else
            rs!harga_tanpa_gst = Null 'Harga Barang Tanpa Tax (kalau gst included)
        End If
        If Frm83.CB5 = 1 Then 'Barang Permata
        
            If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) Then
                Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                
                rs!harga_item = Format(Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST, "0.00") 'Jumlah Harga Modal Keseluruhan @ Harga Barang + Upah + GST (RM)
            End If
            
        End If

        If Frm83.CB4 = 1 Or Frm83.CB10 = 1 Then 'Barang Kemas / Gold Bar
        
            If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) And (Frm83.TB22 <> vbNullString And IsNumeric(Frm83.TB22)) Then
                Frm83_LM_BERAT = Frm83.TB8 'Berat
                Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment
                
                If Frm83.CB12 = 0 Then
                    rs!harga_item = Format(Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST, "0.00") 'Jumlah Harga Modal Keseluruhan @ Harga Barang + Upah + GST (RM)
                ElseIf Frm83.CB12 = 1 Then
                    rs!harga_item = Format(((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST, "0.00") 'Jumlah Harga Modal Keseluruhan @ Harga Barang + Upah + GST (RM)
                End If
                
            End If
            
        End If
        
'### Kod Bagi Status ###
'==========================================
'0 : Tidak Aktif (Tidak Perlu Buat Apa-Apa)
'10 : Aktif - Kemasukkan Data Baru
'2 : Barang Sudah Terjual ****Tidak Dibenarkan Untuk Diedit Atau Dipadamkan
'3 : Kemasukkan Data Baru
'4 : Data Diedit
'5 : Keluarkan Data Dari Database (Bagi Item Asal Yang Dipadamkan)
'6 : Ignore Kemasukkan Data Ke Dalam Database

        rs!StatusItem = 10
        
'### Jenis ###
'0 : BK
'1 : Barang permata
'2 : Emas terpakai BK
'3 : Emas terpakai permata
'4 : gold Bar
'5 : Emas terpakai gold bar
'6 : Trade In BK
'7 : Trade In Barang Permata
'8 : Trade In Gold Bar

'=========================================================
'Frm83.L41_Text
'0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
'=========================================================

        'If Frm83.L41_Text = 1 Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
        
            'If Frm83.CB8 = 1 Then 'Buyback / Trade in
        '        If Frm83.CB4 = 1 Then 'Barang kemas
        '            rs!jenis = 6
        '        ElseIf Frm83.CB5 = 1 Then 'Barang permata
        '            rs!jenis = 7
        '        End If
        '        If Frm83.CB10 = 1 Then 'Gold bar
        '            rs!jenis = 8
        '        End If
            'End If
        
        'ElseIf Frm83.L41_Text = 0 Or Frm83.L41_Text = 2 Then
        
            If Frm83.CB7 = 1 Then 'Penerimaan stok baru
                If Frm83.CB4 = 1 Then 'Barang kemas
                    rs!jenis = 0
                ElseIf Frm83.CB5 = 1 Then 'Barang permata
                    rs!jenis = 1
                End If
                If Frm83.CB10 = 1 Then 'Gold bar
                    rs!jenis = 4
                End If
            ElseIf Frm83.CB8 = 1 Then 'Buyback / Trade in
                If Frm83.CB4 = 1 Then 'Barang kemas
                    rs!jenis = 2
                ElseIf Frm83.CB5 = 1 Then 'Barang permata
                    rs!jenis = 3
                End If
                If Frm83.CB10 = 1 Then 'Gold bar
                    rs!jenis = 5
                End If
            End If
        
        If Frm83.L41_Text = 0 Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
            
            rs!jenis_trade_in = 0 '0 : Belian emas dari pelanggan , 1 : Belian emas secara trade in
            
        ElseIf Frm83.L41_Text = 1 Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
        
            rs!jenis_trade_in = 1 '0 : Belian emas dari pelanggan , 1 : Belian emas secara trade in
            
        End If
        
        If Frm83.TB35 <> vbNullString Then
            rs!upah_per_gram = Format(Frm83.TB35, "0.00")
        Else
            rs!upah_per_gram = "0.00"
        End If
        If Frm83.CB14 = 1 Then
            rs!flag_upah = 0
        ElseIf Frm83.CB15 = 1 Then
            rs!flag_upah = 1
            
            If IsNumeric(Frm83.TB8) And Frm83.TB8 <> 0 Then
                Frm83_LM_BERAT = Frm83.TB8 'Berat
            End If
            
            Frm83_LM_UPAH = Frm83.TB4 'Upah
            
            rs!upah_per_gram = Format(Frm83_LM_UPAH / Frm83_LM_BERAT, "0.00")
        End If
        
        If Frm83.TB28 <> vbNullString Then
            rs!no_id_gst = UCase(Frm83.TB28)
        Else
            rs!no_id_gst = Null
        End If
        If Frm83.TB15 <> vbNullString Then
            rs!bill_No_Belian = UCase(Frm83.TB15)
        Else
            rs!bill_No_Belian = Null
        End If
        rs!tarikh_belian = Frm83.DTPicker1
        
        rs!write_timestamp = Now 'Tarikh & Masa Data Dimasukkan
        rs!flag_image = 0
        rs!Image = Null
        rs!terminal = G_TERMINAL
        'If Frm83.L32_Text = 1 Then
        '    Set rs2 = New ADODB.Recordset
        '    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        '    rs2.Open "select * from 1_image_item_temp where initial_flag='" & "1" & "'", cn, adOpenKeyset, adLockOptimistic
            
        '    If Not rs2.EOF Then
        '        rs!flag_image = 1
        '        rs!Image = rs2!Image
        '    End If
            
        '    rs2.Close
        '    Set rs2 = Nothing
        'End If
        
        rs.Update
        DATA_SAVE = 1
        
        rs.Close
        Set rs = Nothing
        
        If DATA_SAVE = 1 Then
            'Frm83.L3_Text = Frm83_LM_No_SIRI + 1 'No. Siri Barcode

            'If Frm83.CB9 = 1 Then
            '    Frm83.TB7 = Format(Frm83.L3_Text, "000000") 'No. Siri Barcode
            'ElseIf Frm83.CB10 = 1 Then
            '    Frm83.TB7 = Format(Frm83.L3_Text, "000000") & "W" 'No. Siri Barcode
            'End If
            
            Call Frm83_Reset_Form
            Call Frm83_Senarai_Belian_Header
            Call Frm83_Senarai_Belian
            
            MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
            'Frm83.TB11.SetFocus
        End If
    End If
End If
End Sub
Private Sub CMD10_Click()
'On Error Resume Next
Dim Data_Err(10)
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim strsql As String
Dim LM_ID As Double
Dim LM_JUM_BAYARAN As Double
Dim LM_TOTAL As Double

x = 0
Y = 0 '0 : Tiada Perubahan Pada Data , 1 : Ada Perubahan Pada Data
DATA_SAVE = 0

If Frm83.CBB6 <> vbNullString Then
    Frm83_LM_EMP_NAME = Split(Frm83.CBB6, "  |  ")(0)
End If
If Frm83.L10_Text = 0 Then
    x = x + 1
    Data_Err(x) = "Tiada senarai belian/stok."
End If
If Frm83.CBB6 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Sila Pilih [Nama Pekerja]."
End If
If Frm83.L9_Text = vbNullString Then
    'If Not IsNumeric(Frm83.L9_Text) Then
        x = x + 1
        Data_Err(x) = "Technical error. Sila keluar dari menu ini dan cuba sekali lagi."
    'End If
End If
If Frm83.TB40 = vbNullString Or (Frm83.TB40 <> vbNullString And Not IsNumeric(Frm83.TB40)) Then
    x = x + 1
    Data_Err(x) = "Sila Masukkan [Tunai (RM)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB41 = vbNullString Or (Frm83.TB41 <> vbNullString And Not IsNumeric(Frm83.TB41)) Then
    x = x + 1
    Data_Err(x) = "Sila Masukkan [Bank In (RM)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.CB8 = 1 Then
    If Frm83.L12_Text = vbNullString Then
        'If Not IsNumeric(Frm83.L12_Text) Then
            x = x + 1
            Data_Err(x) = "Technical error. Sila keluar dari menu ini dan cuba sekali lagi."
        'End If
    End If
End If
If Frm83.CB8 = 1 Then
'### Periksa Samada Maklumat Penjual Diisi Dalam Kedua-dua Ruangan Berdaftar Dan Tidak Berdaftar ### - Start
    If x = 0 Then
        If Frm83.L36_Text <> vbNullString And Frm83.L37_Text <> vbNullString Then
        
            MsgBox "Data bagi penjual telah diisi bagi kedua-dua ruangan pelanggan berdaftar dan tidak berdaftar." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila padam salah satu yang tidak berkenaan.", vbExclamation, "Info"
                        
            Exit Sub
              
        End If
    End If
'### Periksa Samada Maklumat Penjual Diisi Dalam Kedua-dua Ruangan Berdaftar Dan Tidak Berdaftar ### - End
End If


If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Data_Err(Y)
    Next Y
    Exit Sub
Else

    LM_JUM_BAYARAN = 0
    LM_TOTAL = 0
    
    LM_JUM_BAYARAN = Frm83.TB42
    LM_TOTAL = Frm83.L26_Text
    
    If LM_TOTAL <> LM_JUM_BAYARAN Then
        MsgBox "Jumlah voucher TIDAK SAMA dengan jumlah bayaran." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Jumlah Voucher : RM " & Format(LM_TOTAL, "#,##0.00") & vbCrLf & _
                "Jumlah Cara Bayaran : RM " & Format(LM_JUM_BAYARAN, "#,##0.00"), vbexclamtion, "Info"
        Exit Sub
    End If
    
    Note = "Adakah anda yakin untuk teruskan urusan belian ini ?" & vbCrLf & _
            vbNullString & vbCrLf & _
            "Data belian akan disimpan ke dalam sistem."
                        
    If Frm83.CB8 = 1 Then

        If Frm83.L37_Text <> vbNullString And Frm83.L36_Text = vbNullString Then
            Note = "Adakah anda yakin untuk teruskan urusan belian ini ?" & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Data belian akan disimpan ke dalam sistem."
        End If
        
        If Frm83.L37_Text = vbNullString And Frm83.L36_Text <> vbNullString Then
            Note = "Adakah anda yakin untuk teruskan urusan belian ini ?" & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Data belian akan disimpan ke dalam sistem." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "*** Oleh ini pembeli ini tidak berdaftar dengan sistem, oleh itu data dan rekod penjual ini tidak akan disimpan di dalam sistem ***"
        End If

        If Frm83.L37_Text = vbNullString And Frm83.L36_Text = vbNullString Then
        
            Note = "TIADA maklumat bagi penjual telah diisi." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Maklumat penjual tidak akan dicetak di dalam payment voucher." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Adakah anda yakin untuk teruskan urusan belian ini ?"
            
        End If

    End If
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
        G_JENIS_URUSAN = 3
        
'$$$$ Periksa status terkini setiap item yang hendak diedit $$$$ - Start
        LM_TRANS_VOID = 0

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select data_database.no_siri_produk from " & G_BELIAN_TEMP & ",data_database where " & G_BELIAN_TEMP & ".no_siri_produk = data_database.no_siri_produk AND data_database.statusitem <> 10 AND " & G_BELIAN_TEMP & ".statusitem = 4", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
            If LM_TRANS_VOID = 1 Then
                LM_SOLD = LM_SOLD & " , " & rs!no_siri_Produk ' & vbCrLf
            End If
            If LM_TRANS_VOID = 0 Then
                LM_SOLD = LM_SOLD & rs!no_siri_Produk  ' & vbCrLf
                
                LM_TRANS_VOID = 1
            End If
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        If LM_TRANS_VOID = 1 Then
            
            MsgBox " Barang-barang berikut tidak dibenarkan untuk diedit kerana status barang tersebut telah berubah." & vbCrLf & _
                    "Senarai barang tersebut adalah seperti di bawah : " & vbCrLf & _
                    LM_SOLD & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila periksa data anda.", vbExclamation, "Periksa Data"
                    
            Exit Sub
            
        End If
'$$$$ Periksa status terkini setiap item yang hendak diedit $$$$ - End

        '$$$ No. staff $$$ - Start
        If InStr(1, Frm83.CBB6, "  |  ") <> 0 Then
            Frm83_LM_EMP_NO = Split(Frm83.CBB6, "  |  ")(1)
            Frm83_LM_EMP_NAMA = Split(Frm83.CBB6, "  |  ")(0)
        Else
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!NoPekerja) Then Frm83_LM_EMP_NO = rs!NoPekerja
    
            End If
            
            rs.Close
            Set rs = Nothing
        End If
    
        Frm83_LM_No_RUJUKAN_BELIAN = Frm83.L9_Text 'No. Rujukan Belian
        LM_NOW = Now
        
'###Masukkan Data Belian Ke Dalam Database Akaun Belian### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 16_gold_bar_belian where no_rujukan='" & Frm83.L9_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!trade_in_status) Then
                If rs!trade_in_status = 1 Then
                
                    MsgBox "Anda tidak dibenarkan untuk edit voucher trade ini kerana status voucher telah berubah. Mungkin telah digunakan untuk urusan pembelian lain.", vbExclamation, "Info"
                    
                    rs.Close
                    Set rs = Nothing
                    
                    Exit Sub
                End If
            End If
            
            G_ID = rs!ID
            Call recovery_16_gold_bar_belian

            rs!tarikh = Frm83.DTPicker1 'Tarikh Belian
            rs!cara_bayaran = 0 'Cara Belian // 0 : Cash @ Bank in @ Kad Kredit @ Kad Debit , 1 : Cheque
            If Frm83.TB40 <> vbNullString Then
                rs!tunai = Format(Frm83.TB40, "0.00") 'Cara Bayaran : Tunai (Jumlah Keseluruhan Dengan Cukai GST)
            Else
                rs!tunai = Null 'Cara Bayaran : Tunai (Jumlah Keseluruhan Dengan Cukai GST)
            End If
            If Frm83.TB41 <> vbNullString Then
                rs!bank_in = Format(Frm83.TB41, "0.00") 'Cara Bayaran : Tunai (Jumlah Keseluruhan Dengan Cukai GST)
            Else
                rs!bank_in = Null 'Cara Bayaran : Tunai (Jumlah Keseluruhan Dengan Cukai GST)
            End If
            If Frm83.L11_Text <> vbNullString Then
                rs!jumlah_asal = Format(Frm83.L11_Text, "0.00") 'Jumlah Bayaran Asal (Jumlah Tanpa Cukai GST)
            Else
                rs!jumlah_asal = Null 'Jumlah Bayaran Asal (Jumlah Tanpa Cukai GST)
            End If
            If Frm83.L8_Text <> vbNullString Then
                rs!gst_value = Frm83.L8_Text 'Jumlah Cukai GST (%)
            Else
                rs!gst_value = "0" 'Jumlah Cukai GST (%)
            End If
            If Frm83.L22_Text <> vbNullString Then
                rs!gst_zr_harga = Format(Frm83.L22_Text, "0.00") 'Jumlah Bayaran Yang Dikenakan Cukai GST ZR (RM)
            Else
                rs!gst_zr_harga = "0.00" 'Jumlah Bayaran Yang Dikenakan Cukai GST ZR (RM)
            End If
            If Frm83.L23_Text <> vbNullString Then
                rs!gst_zr_cukai = Format(Frm83.L23_Text, "0.00") 'Jumlah Bayaran Cukai GST ZR (RM)
            Else
                rs!gst_zr_cukai = "0.00" 'Jumlah Bayaran Cukai GST ZR (RM)
            End If
            If Frm83.L24_Text <> vbNullString Then
                rs!gst_sr_harga = Format(Frm83.L24_Text, "0.00") 'Jumlah Bayaran Yang Dikenakan Cukai GST SR (RM)
            Else
                rs!gst_sr_harga = "0.00" 'Jumlah Bayaran Yang Dikenakan Cukai GST SR (RM)
            End If
            If Frm83.L25_Text <> vbNullString Then
                rs!gst_sr_cukai = Format(Frm83.L25_Text, "0.00") 'Jumlah Bayaran Cukai GST SR (RM)
            Else
                rs!gst_sr_cukai = "0.00" 'Jumlah Bayaran Cukai GST SR (RM)
            End If
            If Frm83.TB28 <> vbNullString Then
                rs!no_id_gst_supplier = Frm83.TB28 'No. ID GST Supplier
            Else
                rs!no_id_gst_supplier = Null 'No. ID GST Supplier
            End If
            If Frm83.TB15 <> vbNullString Then
                rs!no_resit_supplier = UCase(Frm83.TB15) 'No. Resit Dari Supplier (Jika Ada)
            Else
                rs!no_resit_supplier = Null 'No. Resit Dari Supplier (Jika Ada)
            End If
            If Frm83.TB1 <> vbNullString Then
                rs!Kod_Supplier = UCase(Frm83.TB1) 'Kod Supplier
            Else
                rs!Kod_Supplier = Null 'Kod Supplier
            End If
            If Frm83.L11_Text <> vbNullString Then
                rs!jumlah_tanpa_gst = Format(Frm83.L11_Text, "0.00") 'Jumlah Bayaran Asal (Jumlah Tanpa Cukai GST)
            Else
                rs!jumlah_tanpa_gst = Null 'Jumlah Bayaran Asal (Jumlah Tanpa Cukai GST)
            End If
            If Frm83.L26_Text <> vbNullString Then
                rs!jumlah_dengan_gst = Format(Frm83.L26_Text, "0.00") 'Jumlah Bayaran Keseluruhan (Jumlah Dengan Cukai GST)
            Else
                rs!jumlah_dengan_gst = Null 'Jumlah Bayaran Keseluruhan (Jumlah Dengan Cukai GST)
            End If
            If Frm83.CB7 = 1 Then
                rs!flag_trade_in = 0 'Flag Trade In // 0 : Tiada , 1 : Ada
            ElseIf Frm83.CB8 = 1 Then
                rs!flag_trade_in = 1 'Flag Trade In // 0 : Tiada , 1 : Ada
                rs!trade_in_status = 0 'Flag Samada Trade In Sudah Digunakan Atau Tidak , 0 : Tiada , 1 : Ada
                rs!no_resit_trade_in = Frm83.L12_Text 'No. Resit Trade In
            End If
            If Frm83.CB8 = 1 Then
                If Frm83.L37_Text <> vbNullString Then
                    If Frm28.L5_Text <> vbNullString Then
                        rs!no_rujukan_pelanggan_buyback = Frm28.L5_Text 'No. Rujukan Pembeli
                    Else
                        rs!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
                    End If
                Else
                    rs!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
                End If
            Else
                rs!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
            End If
            rs!no_pekerja = Frm83_LM_EMP_NO 'No. Pekerja
            rs!nama_pekerja = Frm83_LM_EMP_NAMA
            rs!jenis_urusan = G_JENIS_URUSAN
            rs!no_staff = G_LOGIN_USER 'No. rujukan pekerja yang membuat perubahan ini
            rs!write_timestamp2 = LM_NOW
            rs!remarks = "Edit data stok"
            rs!terminal = G_TERMINAL
            If Not IsNull(rs!cawangan) Then LM_CAWANGAN = rs!cawangan
            
            DATA_SAVE = 1
            rs.Update
            
        End If
        
        rs.Close
        Set rs = Nothing

        'carian kod kedai
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select kod_cawangan from 56_maklumat_kedai where cawangan ='" & LM_CAWANGAN & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!kod_cawangan) Then LM_KOD_KEDAI = rs!kod_cawangan
        End If
        
        rs.Close
        Set rs = Nothing
        
'###Padam Data Penjual Yang Tidak Berdaftar Ke Dalam Database### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 44_senarai_pelanggan where no_resit='" & Frm83.L12_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            G_ID = rs!ID
            Call recovery_44_senarai_pelanggan
            
            rs.Delete
            rs.Update
            
        End If
        
        rs.Close
        Set rs = Nothing
'###Padam Data Penjual Yang Tidak Berdaftar Ke Dalam Database### - End
        
'###Masukkan Data Penjual Yang Tidak Berdaftar Ke Dalam Database### - Start
        If Frm83.L36_Text <> vbNullString And Frm26.TB1 <> vbNullString Then

            If Frm26.TB1 <> vbNullString Then 'Nama
                LM_NAMA = UCase(Frm26.TB1)
            Else
                LM_NAMA = Null
            End If
            If Frm26.TB2 <> vbNullString Then 'No. Telefon
                LM_NO_TEL = UCase(Frm26.TB2)
            Else
                LM_NO_TEL = Null
            End If
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            strsql = "insert into 44_senarai_pelanggan(tarikh,no_resit,Nama,no_tel,write_timestamp,no_staff,terminal,jenis_urusan,cawangan)" & _
                    "select '" & Frm83.DTPicker1 & "','" & Frm83.L12_Text & "','" & LM_NAMA & "','" & LM_NO_TEL & "','" & LM_NOW & "','" & Frm83_LM_EMP_NO & "','" & G_TERMINAL & "','" & G_JENIS_URUSAN & "','" & LM_CAWANGAN & "'"
                                            
            Set rs = cn.Execute(strsql)
            Set rs = Nothing
            
        End If
'###Masukkan Data Penjual Yang Tidak Berdaftar Ke Dalam Database### - End

'###Masukkan Data Belian Ke Dalam Database### - Start
'Masukkan Data Ke Dalam Database
'### Kod Bagi Status ###
'==========================================
'0 : Tidak Aktif (Tidak Perlu Buat Apa-Apa)
'1 : Aktif - Kemasukkan Data Baru
'2 : Barang Sudah Terjual ****Tidak Dibenarkan Untuk Diedit Atau Dipadamkan
'3 : Kemasukkan Data Baru
'4 : Data Diedit
'5 : Keluarkan Data Dari Database (Bagi Item Asal Yang Dipadamkan)
'6 : Ignore Kemasukkan Data Ke Dalam Database
'--- Yang Terlibat Dalam Urusan Ini Adalah HANYA 0 , 3 Dan 4

'### Masukkan maklumat data barang ke dalam table #data_database ### - Start
'Barang / item baru

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        strsql = "insert into Data_Database(NoRujukanSistem,cawangan,nama_pekerja,tarikh_belian,bill_No_Belian,terminal,supplier_ID,nama_Supplier,Kod_Supplier,purity_ID,purity,kod_Purity,kategori_produk_ID,kategori_Produk,Kod_Kategori_Produk,Berat,beza_berat,UPAH,Upah30,riyal,kos_Belian_Gram,kos_Belian_Item,SpreadValue,harga_lepas_spread,adjustment,kos_item_tanpa_tax,cara_Belian,dimension_Panjang,dimension_Lebar,dimension_Saiz,code1,code2,harga_Per_Gram_Item,dulang,no_cert,gst_barang_atau_upah,StatusItem,Upah_Jualan,Upah_Member,Upah_RAF,Upah_Pengedar,code_Supplier,HargaJualan_Member,HargaJualan_Pengedar,upah_normal_dealer,upah_master_dealer,HargaJualan_RAF,hargajualan_normal_dealer,hargajualan_master_dealer,remarks,gst_ari_nashi,kadar_gst,jumlah_gst,harga_item,receiving_Status,harga_tanpa_gst,gst_included,jenis_trade_in,flag_upah,upah_per_gram,flag_image,write_timestamp,no_id_gst,susut_berat,no_pekerja)" & _
                    "select '" & Frm83_LM_No_RUJUKAN_BELIAN & "','" & LM_CAWANGAN & "','" & Frm83_LM_EMP_NAMA & "',tarikh_belian,bill_No_Belian,terminal,supplier_ID,nama_Supplier,Kod_Supplier,purity_ID,purity,kod_Purity,kategori_produk_ID,kategori_Produk,Kod_Kategori_Produk,Berat,beza_berat,UPAH,Upah30,riyal,kos_Belian_Gram,kos_Belian_Item,Spread,harga_lepas_spread,adjustment,kos_item_tanpa_tax,cara_Belian,dimension_Panjang,dimension_Lebar,dimension_Saiz,code1,code2,harga_Per_Gram_Item,dulang,no_cert,gst_barang_atau_upah,10,Upah_Jualan,Upah_Member,Upah_RAF,Upah_Pengedar,code_Supplier,HargaJualan_Member,HargaJualan_Pengedar,upah_normal_dealer,upah_master_dealer,HargaJualan_RAF,hargajualan_normal_dealer,hargajualan_master_dealer,remarks,gst_ari_nashi,kadar_gst,jumlah_gst,harga_item,jenis,harga_tanpa_gst,gst_included,jenis_trade_in,flag_upah,upah_per_gram,flag_image,'" & LM_NOW & "',no_id_gst,0.00,'" & Frm83_LM_EMP_NO & "' from " & G_BELIAN_TEMP & " WHERE StatusItem='" & 3 & "'"

        Set rs = cn.Execute(strsql)
        Set rs = Nothing
        
'### Masukkan maklumat data barang ke dalam table #data_database ### - End

        '### Masukkan data lama ke dalam table #72_data_amendment ### - Start
        Set rs3 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "insert into 72_data_amendment(id_asal,jenis_barang,menu)" & _
                    "select id_database,jenis,0 from " & G_BELIAN_TEMP & " WHERE " & G_BELIAN_TEMP & ".statusitem = 4"
        
        Set rs3 = cn.Execute(strsql)
        Set rs3 = Nothing
        
        Set rs3 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        'Masukkan maklumat bagi barang kemas , trade in barang kemas , gold bar , trade in gold bar shj
        strsql = "UPDATE data_database,72_data_amendment set 72_data_amendment.no_siri_produk = data_database.no_siri_produk ," _
        & "72_data_amendment.kategori_produk = data_database.kategori_produk ," _
        & "72_data_amendment.nama_Supplier = data_database.nama_Supplier ," _
        & "72_data_amendment.berat = data_database.berat ," _
        & "72_data_amendment.upah = data_database.upah ," _
        & "72_data_amendment.purity = data_database.kod_purity ," _
        & "72_data_amendment.jenis='" & 0 & "'" _
        & "WHERE 72_data_amendment.id_asal = data_database.ID AND (72_data_amendment.jenis_barang = 0 OR 72_data_amendment.jenis_barang = 2 OR 72_data_amendment.jenis_barang = 4 OR 72_data_amendment.jenis_barang = 5)"
        
        Set rs3 = cn.Execute(strsql)
        Set rs3 = Nothing
        
        'Masukkan maklumat bagi barang permata , trade in barang permata shj
        strsql = "UPDATE data_database,72_data_amendment set 72_data_amendment.no_siri_produk = data_database.no_siri_produk ," _
        & "72_data_amendment.kategori_produk = data_database.kategori_produk ," _
        & "72_data_amendment.nama_Supplier = data_database.nama_Supplier ," _
        & "72_data_amendment.upah = data_database.harga_item ," _
        & "72_data_amendment.purity = data_database.kod_purity ," _
        & "72_data_amendment.jenis='" & 0 & "'" _
        & "WHERE 72_data_amendment.id_asal = data_database.ID AND (72_data_amendment.jenis_barang = 1 OR 72_data_amendment.jenis_barang = 3)"
        
        Set rs3 = cn.Execute(strsql)
        Set rs3 = Nothing
        
'### Kod Bagi Status ###
'==========================================
'0 : Tidak Aktif (Tidak Perlu Buat Apa-Apa)
'1 : Aktif - Kemasukkan Data Baru
'2 : Barang Sudah Terjual ****Tidak Dibenarkan Untuk Diedit Atau Dipadamkan
'3 : Kemasukkan Data Baru
'4 : Data Diedit
'5 : Keluarkan Data Dari Database (Bagi Item Asal Yang Dipadamkan)
'6 : Ignore Kemasukkan Data Ke Dalam Database
        '$$$$ Recovery start $$$$
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_BELIAN_TEMP & " where statusitem = 4 OR statusitem = 5", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
        
            G_ID = rs!id_database
            Call recovery_data_database

            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        '$$$$ Recovery end $$$$
        
'### Update data barang ke dalam table #data_database ### - Start
'Barang sedia ada
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

        strsql = "UPDATE data_database," & G_BELIAN_TEMP & " SET Data_Database.NoRujukanSistem='" & Frm83_LM_No_RUJUKAN_BELIAN & "'," _
        & "Data_Database.tarikh_belian = " & G_BELIAN_TEMP & ".tarikh_belian ," _
        & "Data_Database.bill_no_belian = " & G_BELIAN_TEMP & ".bill_no_belian , Data_Database.hargajualan_pengedar = " & G_BELIAN_TEMP & ".hargajualan_pengedar , Data_Database.upah_normal_dealer = " & G_BELIAN_TEMP & ".upah_normal_dealer , Data_Database.upah_master_dealer = " & G_BELIAN_TEMP & ".upah_master_dealer , Data_Database.hargajualan_raf = " & G_BELIAN_TEMP & ".hargajualan_raf ," _
        & "Data_Database.supplier_ID = " & G_BELIAN_TEMP & ".supplier_ID , Data_Database.hargajualan_normal_dealer = " & G_BELIAN_TEMP & ".hargajualan_normal_dealer , Data_Database.hargajualan_master_dealer = " & G_BELIAN_TEMP & ".hargajualan_master_dealer , Data_Database.remarks = " & G_BELIAN_TEMP & ".remarks ," _
        & "Data_Database.nama_Supplier = " & G_BELIAN_TEMP & ".nama_Supplier , Data_Database.gst_ari_nashi = " & G_BELIAN_TEMP & ".gst_ari_nashi , Data_Database.kadar_gst = " & G_BELIAN_TEMP & ".kadar_gst , Data_Database.jumlah_gst = " & G_BELIAN_TEMP & ".jumlah_gst , Data_Database.harga_item = " & G_BELIAN_TEMP & ".harga_item , Data_Database.receiving_status = " & G_BELIAN_TEMP & ".jenis , Data_Database.harga_tanpa_gst = " & G_BELIAN_TEMP & ".harga_tanpa_gst ," _
        & "Data_Database.Kod_Supplier = " & G_BELIAN_TEMP & ".Kod_Supplier , Data_Database.gst_included = " & G_BELIAN_TEMP & ".gst_included , Data_Database.jenis_trade_in = " & G_BELIAN_TEMP & ".jenis_trade_in , Data_Database.flag_upah = " & G_BELIAN_TEMP & ".flag_upah , Data_Database.upah_per_gram = " & G_BELIAN_TEMP & ".upah_per_gram , Data_Database.flag_image = " & G_BELIAN_TEMP & ".flag_image ," _
        & "Data_Database.purity_ID = " & G_BELIAN_TEMP & ".purity_ID ," _
        & "Data_Database.purity = " & G_BELIAN_TEMP & ".purity ," _
        & "Data_Database.kod_Purity = " & G_BELIAN_TEMP & ".kod_Purity ," _
        & "Data_Database.kategori_produk_ID = " & G_BELIAN_TEMP & ".kategori_produk_ID , Data_Database.code1 = " & G_BELIAN_TEMP & ".code1 , Data_Database.code2 = " & G_BELIAN_TEMP & ".code2 ," _
        & "Data_Database.kategori_Produk = " & G_BELIAN_TEMP & ".kategori_Produk ," _
        & "Data_Database.Kod_Kategori_Produk = " & G_BELIAN_TEMP & ".Kod_Kategori_Produk , Data_Database.terminal = " & G_BELIAN_TEMP & ".terminal ," _
        & "Data_Database.Berat = " & G_BELIAN_TEMP & ".Berat ," _
        & "Data_Database.beza_berat = " & G_BELIAN_TEMP & ".beza_berat ," _
        & "Data_Database.upah = " & G_BELIAN_TEMP & ".upah ," _
        & "Data_Database.upah30 = " & G_BELIAN_TEMP & ".upah30 ," _
        & "Data_Database.no_pekerja='" & Frm83_LM_EMP_NO & "'," _
        & "Data_Database.nama_pekerja='" & Frm83_LM_EMP_NAMA & "'," _
        & "Data_Database.menu='" & G_JENIS_URUSAN & "'," _
        & "Data_Database.riyal = " & G_BELIAN_TEMP & ".riyal , Data_Database.no_id_gst = " & G_BELIAN_TEMP & ".no_id_gst ," _
        & "Data_Database.kos_belian_gram = " & G_BELIAN_TEMP & ".kos_belian_gram , Data_Database.kos_belian_item = " & G_BELIAN_TEMP & ".kos_belian_item , Data_Database.spreadvalue = " & G_BELIAN_TEMP & ".spread , Data_Database.harga_lepas_spread = " & G_BELIAN_TEMP & ".harga_lepas_spread , Data_Database.adjustment = " & G_BELIAN_TEMP & ".adjustment , Data_Database.kos_item_tanpa_tax = " & G_BELIAN_TEMP & ".kos_item_tanpa_tax , Data_Database.cara_belian = " & G_BELIAN_TEMP & ".cara_belian , Data_Database.dimension_panjang = " & G_BELIAN_TEMP & ".dimension_panjang , Data_Database.dimension_lebar = " & G_BELIAN_TEMP & ".dimension_lebar , Data_Database.dimension_saiz = " & G_BELIAN_TEMP & ".dimension_saiz ," _
        & "Data_Database.harga_per_gram_item = " & G_BELIAN_TEMP & ".harga_per_gram_item , Data_Database.dulang = " & G_BELIAN_TEMP & ".dulang , Data_Database.no_cert = " & G_BELIAN_TEMP & ".no_cert , Data_Database.gst_barang_atau_upah = " & G_BELIAN_TEMP & ".gst_barang_atau_upah , Data_Database.statusitem = 10 , Data_Database.upah_jualan = " & G_BELIAN_TEMP & ".upah_jualan , Data_Database.upah_member = " & G_BELIAN_TEMP & ".upah_member , Data_Database.upah_raf = " & G_BELIAN_TEMP & ".upah_raf , Data_Database.upah_pengedar = " & G_BELIAN_TEMP & ".upah_pengedar , Data_Database.code_supplier = " & G_BELIAN_TEMP & ".code_supplier , Data_Database.hargajualan_member = " & G_BELIAN_TEMP & ".hargajualan_member , " _
        & "Data_Database.write_timestamp2='" & LM_NOW & "' WHERE " & G_BELIAN_TEMP & ".statusitem = 4 AND Data_Database.id = " & G_BELIAN_TEMP & ".id_database"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update data barang ke dalam table #data_database ### - End

'### Update data barang ke dalam table #data_database ### - Start
'Barang yang dipadamkan
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE Data_Database," & G_BELIAN_TEMP & " SET Data_Database.statusitem='" & 0 & "', Data_Database.terminal = " & G_BELIAN_TEMP & ".terminal ," _
                & "Data_Database.no_pekerja='" & Frm83_LM_EMP_NO & "'," _
                & "Data_Database.menu='" & G_JENIS_URUSAN & "'," _
                & "Data_Database.write_timestamp2='" & LM_NOW & "' WHERE " & G_BELIAN_TEMP & ".statusitem = 5 AND Data_Database.id = " & G_BELIAN_TEMP & ".id_database"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update data barang ke dalam table #data_database ### - End


'### Update maklumat di bawah ke dalam maklumat barang ### - Start
'@no_siri_produk
'@Barcode
'@bill_no_trade_in
'@no_rujukan_pelanggan_buyback

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from data_database where NoRujukanSistem='" & Frm83_LM_No_RUJUKAN_BELIAN & "'", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
            
            LM_ID = rs!ID

            If Not IsNull(rs!ID) And Not IsNull(rs!Kod_Kategori_Produk) Then
                rs!no_siri_Produk = LM_KOD_KEDAI & "-" & rs!Kod_Kategori_Produk & Format(rs!ID, "000000")
            Else
                rs!no_siri_Produk = LM_KOD_KEDAI & "-" & Format(rs!ID, "000000")
            End If
            If Not IsNull(rs!ID) Then
                rs!Barcode = Format(rs!ID, "000000")
            Else
                rs!Barcode = Format(rs!ID, "000000")
            End If
            
            If Frm83.CB8 = 1 Then
                rs!bill_No_Trade_In = Frm83.L12_Text '"TI" & Format(Frm83_LM_NO_TI, "000000") 'No. Resit Trade In

                If Frm83.L37_Text <> vbNullString Then
                    If Frm28.L5_Text <> vbNullString Then
                        rs!no_rujukan_pelanggan_buyback = Frm28.L5_Text 'No. Rujukan Pembeli
                    Else
                        rs!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
                    End If
                Else
                    rs!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
                End If
            Else
                rs!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
            End If
            rs!nama_pekerja = Frm83_LM_EMP_NAMA
            rs.Update
        
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
'### Update maklumat di bawah ke dalam maklumat barang ### - End


        'Update data selepas perubahan bagi barang kemas , trade in barang kemas , gold bar , trade in gold bar
        Set rs3 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

        strsql = "UPDATE data_database,72_data_amendment set 72_data_amendment.no_siri_produk_new = data_database.no_siri_produk ," _
        & "72_data_amendment.kategori_produk_new = data_database.kategori_produk ," _
        & "72_data_amendment.nama_supplier_new = data_database.nama_Supplier ," _
        & "72_data_amendment.berat_new = data_database.berat ," _
        & "72_data_amendment.upah_new = data_database.upah ," _
        & "72_data_amendment.purity_new = data_database.kod_purity ," _
        & "72_data_amendment.jenis='" & 0 & "'," _
        & "72_data_amendment.terminal='" & G_TERMINAL & "'," _
        & "72_data_amendment.nama_pic='" & Frm83_LM_EMP_NAME & "'," _
        & "72_data_amendment.write_timestamp='" & LM_NOW & "'" _
        & "WHERE 72_data_amendment.id_asal = data_database.ID AND (72_data_amendment.jenis_barang = 0 OR 72_data_amendment.jenis_barang = 2 OR 72_data_amendment.jenis_barang = 4 OR 72_data_amendment.jenis_barang = 5)"
        
        Set rs3 = cn.Execute(strsql)
        Set rs3 = Nothing
        
        'Update data selepas perubahan bagi barang permata , trade in barang permata
        Set rs3 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

        strsql = "UPDATE data_database,72_data_amendment set 72_data_amendment.no_siri_produk_new = data_database.no_siri_produk ," _
        & "72_data_amendment.kategori_produk_new = data_database.kategori_produk ," _
        & "72_data_amendment.nama_supplier_new = data_database.nama_Supplier ," _
        & "72_data_amendment.upah_new = data_database.harga_item ," _
        & "72_data_amendment.purity_new = data_database.kod_purity ," _
        & "72_data_amendment.jenis='" & 0 & "'," _
        & "72_data_amendment.terminal='" & G_TERMINAL & "'," _
        & "72_data_amendment.nama_pic='" & Frm83_LM_EMP_NAME & "'," _
        & "72_data_amendment.write_timestamp='" & LM_NOW & "'" _
        & "WHERE 72_data_amendment.id_asal = data_database.ID AND (72_data_amendment.jenis_barang = 1 OR 72_data_amendment.jenis_barang = 3)"
        
        Set rs3 = cn.Execute(strsql)
        Set rs3 = Nothing
        
'### Masukkan maklumat data barang ke dalam table #data_database ### - End
DATA_SAVE = 1
        
        If DATA_SAVE = 1 Then
            If Frm83.TB15 <> vbNullString Then
                Frm83_LM_No_INVOICE_SUPPLIER = UCase(Frm83.TB15)
            Else
                Frm83_LM_No_INVOICE_SUPPLIER = Null
            End If
            
            '#### Update nombor rujukan penjual #### - Start
            If Frm83.CB8 = 1 Then
                If Frm83.L37_Text <> vbNullString Then
                
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
                    strsql = "UPDATE Data_Database set no_rujukan_pelanggan_buyback='" & Frm28.L5_Text & "'" _
                    & "WHERE NoRujukanSistem='" & Frm83.L9_Text & "'"
                    
                    Set rs = cn.Execute(strsql)
                    Set rs = Nothing
                    
                Else
                
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
                    strsql = "UPDATE Data_Database set no_rujukan_pelanggan_buyback='" & Null & "'" _
                    & "WHERE NoRujukanSistem='" & Frm83.L9_Text & "'"
                    
                    Set rs = cn.Execute(strsql)
                    Set rs = Nothing
                    
                End If
                
            End If
            '#### Update nombor rujukan penjual #### - End
        
            user = MDI_frm1.L3_Text
            If Frm83.CB7 = 1 Then LogAct_Memory = "[" & G_LOGIN_USER & "] Edit data stok [" & Frm83.L9_Text & "]."
            If Frm83.CB8 = 1 Then LogAct_Memory = "[" & G_LOGIN_USER & "] Edit data trade in.No. Voucher [" & Frm83.L12_Text & "]."
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
            
            Call amendment_email_check
                
            Note = "Data telah berjaya disimpan." & vbCrLf & _
                    "Sistem akan refresh data."

            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbNo Or Answer = vbYes Then
            
                GM_NEXT_PREV = 2
                
                If Frm101.L33_Text = 0 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                    If Frm101.CB2 = 1 Then 'Report Belian
                        Call Frm85_Header_Report_Belian
                        Call Frm85_report_belian_page
                    End If
                    If Frm101.CB4 = 1 Then 'Report Buyback
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
                    Call Frm85_Header_Report_buyback_gb
                    Call Frm85_carian_buyback_page
                ElseIf Frm101.L33_Text = 4 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                    Call Frm85_Header_Report_Belian
                    Call Frm85_search_invoice_supplier_page
                ElseIf Frm101.L33_Text = 5 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                    Call Frm85_Header_Report_Belian
                    Call Frm85_report_belian_barcode
                ElseIf Frm101.L33_Text = 6 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                    Call Frm85_Header_Report_Belian
                    Call Frm85_report_buyback_barcode
                ElseIf Frm101.L33_Text = 7 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                    Call Frm85_Header_Report_belian_gb
                    Call Frm85_report_belian_gb_barcode
                ElseIf Frm101.L33_Text = 8 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                    Call Frm85_Header_Report_belian_gb
                    Call Frm85_report_buyback_gb_barcode
                End If
                
                Frm85.Show
                Unload Frm83

                Unload Frm26
                Unload Frm27
                Unload Frm28
                
            End If
        End If
    End If
End If
End Sub
Private Sub CMD11_Click()
'On Error Resume Next
If Frm83.L10_Text <> 0 Then
    
    If MDI_frm1.L5_Text <> 4 Then
    
        Note = "Adakah mempunyai data yang belum disimpan." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Adakah anda ingin keluar dari menu ini?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
        
            If Frm83.L41_Text = "1" Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
                
                Frm84.Show
                Frm83.Hide
                
            ElseIf Frm83.L41_Text = "0" Or Frm83.L41_Text = "2" Then
            
                Frm85.Show
                Unload Frm26
                Unload Frm27
                Unload Frm28
                Unload Frm83
                
            End If
        
        End If
        
    Else
    
        Frm84.Show
        Frm83.Hide
    
    End If

Else

    If Frm83.L41_Text = "1" Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
        
        Frm84.Show
        Frm83.Hide
        
    ElseIf Frm83.L41_Text = "0" Or Frm83.L41_Text = "2" Then
    
        Frm85.Show
        Unload Frm26
        Unload Frm27
        Unload Frm28
        Unload Frm83
        
    End If

End If
End Sub
Private Sub CMD12_Click()
'On Error Resume Next
Dim Err(30)
Dim Frm83_LM_BERAT As Double
Dim Frm83_LM_UPAH As Double 'Upah
Dim Frm83_LM_HARGA_TOTAL As Double
Dim Frm83_LM_CUKAI_GST As Double
Dim Frm83_LM_HARGA_SEMASA As Double 'Harga Semasa
Dim Frm83_LM_ADJUSTMENT As Double 'Adjustment

Frm83_LM_BERAT = 1
Frm83_LM_UPAH = 0 'Upah
Frm83_LM_HARGA_TOTAL = 0
Frm83_LM_CUKAI_GST = 0
Frm83_LM_HARGA_SEMASA = 0 'Harga Semasa
Frm83_LM_ADJUSTMENT = 0 'Adjustment

x = 0
Y = 0 '0 : Tiada Perubahan Pada Data , 1 : Ada Perubahan Pada Data
DATA_SAVE = 0

If Frm83.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Supplier]."
End If
If Frm83.CBB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Purity]."
End If
If Frm83.CBB3 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Kategori Produk]."
End If
If Frm83.TB1 = vbNullString Then
    x = x + 1
    Err(x) = "Tiada Maklumat [Kod Supplier]."
End If
If Frm83.TB2 = vbNullString Then
    x = x + 1
    Err(x) = "Tiada Maklumat [Kod Purity]."
End If
If Frm83.TB3 = vbNullString Then
    x = x + 1
    Err(x) = "Tiada Maklumat [Kod Kategori Produk]."
End If
'If Frm83.TB6 = vbNullString Or Frm83.TB7 = vbNullString Then
'    x = x + 1
'    Err(x) = "Maklumat [No. Siri Produk] Yang Tidak Lengkap."
'End If
If Frm83.CB9 = 1 And Frm83.CB4 = 0 And Frm83.CB5 = 0 Then
    x = x + 1
    Err(x) = "Sila Buat Pilihan Penerimaan [Barang Kemas] Atau [Barang Permata]."
End If
If Frm83.TB36 <> vbNullString Then

    If InStr(1, Frm83.TB36, "*") <> 0 Or InStr(1, Frm83.TB36, "/") <> 0 Or InStr(1, Frm83.TB36, "\") <> 0 Or InStr(1, Frm83.TB36, "'") <> 0 Then

        x = x + 1
        Err(x) = "[Code 1] mengandungi simbol yang tidak sah."
        
    End If
    
End If
If Frm83.TB37 <> vbNullString Then

    If InStr(1, Frm83.TB37, "*") <> 0 Or InStr(1, Frm83.TB37, "/") <> 0 Or InStr(1, Frm83.TB37, "\") <> 0 Or InStr(1, Frm83.TB37, "'") <> 0 Then

        x = x + 1
        Err(x) = "[Code 2] mengandungi simbol yang tidak sah."
        
    End If
    
End If
If Frm83.CB2 = 0 And Frm83.CB3 = 0 And Frm83.CB11 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan jenis cukai GST"
End If
If Frm83.CB4 = 1 Then
    If Frm83.TB8 = vbNullString Or (Frm83.TB8 <> vbNullString And Not IsNumeric(Frm83.TB8)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Berat]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB8 <> vbNullString Or (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) Then
        
        Frm83_LM_BERAT = Frm83.TB8
        
        If Frm83_LM_BERAT = 0 Then
            x = x + 1
            Err(x) = "Tiada data bagi berat barang. Berat 0 tidak dibenarkan."
        End If
    End If
    If Frm83.TB9 = vbNullString Or (Frm83.TB9 <> vbNullString And Not IsNumeric(Frm83.TB9)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Harga Semasa]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB4 = vbNullString Or (Frm83.TB4 <> vbNullString And Not IsNumeric(Frm83.TB4)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Upah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB24 = vbNullString Or (Frm83.TB24 <> vbNullString And Not IsNumeric(Frm83.TB24)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Pelanggan]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB25 = vbNullString Or (Frm83.TB25 <> vbNullString And Not IsNumeric(Frm83.TB25)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Ahli]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB26 = vbNullString Or (Frm83.TB26 <> vbNullString And Not IsNumeric(Frm83.TB26)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Silver]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB31 = vbNullString Or (Frm83.TB31 <> vbNullString And Not IsNumeric(Frm83.TB31)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Gold]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB32 = vbNullString Or (Frm83.TB32 <> vbNullString And Not IsNumeric(Frm83.TB32)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Platinum]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB33 = vbNullString Or (Frm83.TB33 <> vbNullString And Not IsNumeric(Frm83.TB33)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada M.Dealer]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
ElseIf Frm83.CB5 = 1 Then
    '+++++++++++ Special Request ++++++++++ Start
    'If Frm83.TB8 = vbNullString Or (Frm83.TB8 <> vbNullString And Not IsNumeric(Frm83.TB8)) Then
    '    x = x + 1
    '    Err(x) = "Sila Masukkan [Berat]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    'End If
    'If Frm83.TB8 <> vbNullString Or (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) Then
        
    '    Frm83_LM_BERAT = Frm83.TB8
        
    '    If Frm83_LM_BERAT = 0 Then
    '        x = x + 1
    '        Err(x) = "Tiada data bagi berat barang. Berat 0 tidak dibenarkan."
    '    End If
    'End If
    '+++++++++++ Special Request ++++++++++ End
    If Frm83.TB24 = vbNullString Or (Frm83.TB24 <> vbNullString And Not IsNumeric(Frm83.TB24)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Pelanggan]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB25 = vbNullString Or (Frm83.TB25 <> vbNullString And Not IsNumeric(Frm83.TB25)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Ahli]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB26 = vbNullString Or (Frm83.TB26 <> vbNullString And Not IsNumeric(Frm83.TB26)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Silver]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB31 = vbNullString Or (Frm83.TB31 <> vbNullString And Not IsNumeric(Frm83.TB31)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Gold]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB32 = vbNullString Or (Frm83.TB32 <> vbNullString And Not IsNumeric(Frm83.TB32)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Platinum]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB33 = vbNullString Or (Frm83.TB33 <> vbNullString And Not IsNumeric(Frm83.TB33)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada M.Dealer]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If
If Frm83.TB10 = vbNullString Or (Frm83.TB10 <> vbNullString And Not IsNumeric(Frm83.TB10)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Harga Asal]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB19 = vbNullString Or (Frm83.TB19 <> vbNullString And Not IsNumeric(Frm83.TB19)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Spread (%)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB20 = vbNullString Or (Frm83.TB20 <> vbNullString And Not IsNumeric(Frm83.TB20)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Harga Belian]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB21 = vbNullString Or (Frm83.TB21 <> vbNullString And Not IsNumeric(Frm83.TB21)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Harga Asal-Spread]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB22 = vbNullString Or (Frm83.TB22 <> vbNullString And Not IsNumeric(Frm83.TB22)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Adjusment]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.CB8 = 1 Then
    If Frm83.TB19 = vbNullString Or (Frm83.TB19 <> vbNullString And Not IsNumeric(Frm83.TB19)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Spread]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If
If Frm83.CBB5 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Dulang]."
End If
'If Frm83.CB3 = 1 Then
    If Frm83.TB27 = vbNullString Or (Frm83.TB27 <> vbNullString And Not IsNumeric(Frm83.TB27)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat GST"
    End If
'End If
If Frm83.CB4 = 1 Then
    If Frm83.CB14 = 0 And Frm83.CB15 = 1 Then
        If Frm83.TB35 = vbNullString Or (Frm83.TB35 <> vbNullString And Not IsNumeric(Frm83.TB35)) Then
            x = x + 1
            Err(x) = "Sila buat tetapan pengiraan upah dari supplier"
        End If
    End If
End If
If Frm83.CB14 = 1 Then
    If Frm83.TB35 = vbNullString Or (Frm83.TB35 <> vbNullString And Not IsNumeric(Frm83.TB35)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Upah per gram]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If
If Frm83.TB15 <> vbNullString Then

    If InStr(1, Frm83.TB15, "*") <> 0 Or InStr(1, Frm83.TB15, "/") <> 0 Or InStr(1, Frm83.TB15, "\") <> 0 Or InStr(1, Frm83.TB15, "'") <> 0 Then

        x = x + 1
        Err(x) = "[No. Invoice] mengandungi simbol yang tidak sah."
        
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
    Note = "Adakah anda ingin masukkan data barang ini ke dalam senarai belian?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then

        Frm83_LM_No_SIRI = Frm83.L3_Text 'No. Turutan No. Siri
        Frm83_LM_No_RUJUKAN_BELIAN = Frm83.L9_Text 'No. Rujukan Belian
        
'Re_Gen_Code:
        
        'Set rs = New ADODB.Recordset
        'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        'If Frm83.CB9 = 1 Then rs.Open "select * from Data_Database where barcode='" & Format(Frm83_LM_No_SIRI, "000000") & "'", cn, adOpenKeyset, adLockOptimistic
        'If Frm83.CB10 = 1 Then rs.Open "select * from Data_Database where barcode='" & Format(Frm83_LM_No_SIRI, "000000") & "W" & "'", cn, adOpenKeyset, adLockOptimistic

        'If Not rs.EOF Then
        '    Frm83_LM_No_SIRI = Frm83_LM_No_SIRI + 1
            
        '    rs.Close
        '    Set rs = Nothing
        '    GoTo Re_Gen_Code:
        'End If
        
        'rs.Close
        'Set rs = Nothing
        
        
'###Masukkan Data Belian Ke Dalam Database### - Start
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_BELIAN_TEMP & "", cn, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        If Frm83.L4_Text <> vbNullString Then
            rs!supplier_ID = Frm83.L4_Text 'No. ID Bagi Supplier
        Else
            rs!supplier_ID = Null 'No. ID Bagi Supplier
        End If
        If Frm83.CBB1 <> vbNullString Then
            rs!nama_Supplier = Frm83.CBB1 'Nama Supplier
        Else
            rs!nama_Supplier = Null 'Nama Supplier
        End If
        If Frm83.TB1 <> vbNullString Then
            rs!Kod_Supplier = Frm83.TB1 'Kod Supplier
        Else
            rs!Kod_Supplier = Null 'Kod Supplier
        End If
        If Frm83.L5_Text <> vbNullString Then
            rs!purity_ID = Frm83.L5_Text 'No. ID Bagi Purity
        Else
            rs!purity_ID = Null 'No. ID Bagi Purity
        End If
        If Frm83.CBB2 <> vbNullString Then
            rs!purity = Frm83.CBB2 'Purity
        Else
            rs!purity = Null 'Purity
        End If
        If Frm83.TB2 <> vbNullString Then
            rs!kod_Purity = Frm83.TB2 'Kod Purity
        Else
            rs!kod_Purity = Null 'Kod Purity
        End If
        If Frm83.L6_Text <> vbNullString Then
            rs!kategori_produk_ID = Frm83.L6_Text 'No. ID Bagi Kategori Produk
        Else
            rs!kategori_produk_ID = Null 'No. ID Bagi Kategori Produk
        End If
        If Frm83.CBB3 <> vbNullString Then
            rs!kategori_Produk = Frm83.CBB3 'Kategori Produk
        Else
            rs!kategori_Produk = Null 'Kategori Produk
        End If
        If Frm83.TB3 <> vbNullString Then
            rs!Kod_Kategori_Produk = Frm83.TB3 'Kod Kategori Produk
        Else
            rs!Kod_Kategori_Produk = Null 'Kod Kategori Produk
        End If
        'If Frm83.TB7 <> vbNullString Then
        '    If Frm83.CB9 = 1 Then
        '        rs!Barcode = Format(Frm83_LM_No_SIRI, "000000") 'No. Barcode (6 Digit Terakhir)
        '    ElseIf Frm83.CB10 = 1 Then
        '        rs!Barcode = Format(Frm83_LM_No_SIRI, "000000") & "W" 'No. Barcode (6 Digit Terakhir)
        '    End If
        'Else
        '    rs!Barcode = Null 'No. Barcode (6 Digit Terakhir)
        'End If
        'If Frm83.CB9 = 1 Then
        '    rs!no_siri_Produk = Frm83.TB6 & Format(Frm83_LM_No_SIRI, "000000")  'No. Siri Produk
        'ElseIf Frm83.CB10 = 1 Then
        '    rs!no_siri_Produk = Frm83.TB6 & Format(Frm83_LM_No_SIRI, "000000") & "W"  'No. Siri Produk
        'End If
        If Frm83.CB12 = 0 Then '0 : Pengiraan GST pada harga belian barang , 1 : Pengiraan GST pada upah sahaja
            rs!gst_barang_atau_upah = 0
        ElseIf Frm83.CB12 = 1 Then
            rs!gst_barang_atau_upah = 1
        End If
        If Frm83.CB4 = 1 Or Frm83.CB10 = 1 Then
            If Frm83.TB8 <> vbNullString Then
                rs!Berat = Format(Frm83.TB8, "0.00") 'Berat
                rs!beza_berat = Format(Frm83.TB8, "0.00") 'Baki Berat
            Else
                rs!Berat = Null 'Berat
                rs!beza_berat = Null 'Baki Berat
            End If
            If Frm83.TB29 <> vbNullString Then
                rs!riyal = Format(Frm83.TB29, "0.00") 'Berat Riyal
            Else
                rs!riyal = Null 'Berat Riyal
            End If
            If Frm83.TB9 <> vbNullString Then
                rs!kos_Belian_Gram = Format(Frm83.TB9, "0.00") 'Harga Per Gram (Belian)
            Else
                rs!kos_Belian_Gram = Null 'Harga Per Gram (Belian)
            End If
            If Frm83.TB4 <> vbNullString Then
                rs!UPAH = Frm83.TB4 'Upah (RM)
            Else
                rs!UPAH = Null 'Upah (RM)
            End If
            If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) And (Frm83.TB22 <> vbNullString And IsNumeric(Frm83.TB22)) Then
                Frm83_LM_BERAT = Frm83.TB8 'Berat
                Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment

                If Frm83.CB12 = 0 Then
                    rs!harga_Per_Gram_Item = Format((Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST) / Frm83_LM_BERAT, "0.00") 'Harga Modal Per Gram [Average] (Harga Emas + Upah) Tanpa Tax
                    rs!harga_per_gram_tanpa_gst = Format(Frm83_LM_HARGA_TOTAL / Frm83_LM_BERAT, "0.00")
                ElseIf Frm83.CB12 = 1 Then
                    rs!harga_Per_Gram_Item = Format((((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST) / Frm83_LM_BERAT, "0.00")
                    rs!harga_per_gram_tanpa_gst = Format((((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL) / Frm83_LM_BERAT, "0.00")
                End If
            Else
                rs!harga_Per_Gram_Item = Null
            End If

            If Frm83.TB24 <> vbNullString Then
                rs!Upah_Jualan = Format(Frm83.TB24, "0.00") 'Upah Jualan Kepada Pelanggan
            Else
                rs!Upah_Jualan = Null 'Upah Jualan Kepada Pelanggan
            End If
            If Frm83.TB25 <> vbNullString Then
                rs!Upah_Member = Format(Frm83.TB25, "0.00") 'Upah Jualan Kepada Ahli / Member
            Else
                rs!Upah_Member = Null 'Upah Jualan Kepada Ahli / Member
            End If
            If Frm83.TB26 <> vbNullString Then
                rs!Upah_Pengedar = Format(Frm83.TB26, "0.00") 'Upah Jualan Kepada Pengedar
            Else
                rs!Upah_Pengedar = Null 'Upah Jualan Kepada Pengedar
            End If
            If Frm83.TB31 <> vbNullString Then
                rs!Upah_RAF = Format(Frm83.TB31, "0.00") 'Upah Jualan Kepada RAF
            Else
                rs!Upah_RAF = Null 'Upah Jualan Kepada RAF
            End If
            If Frm83.TB32 <> vbNullString Then
                rs!upah_normal_dealer = Format(Frm83.TB32, "0.00") 'Upah Jualan Kepada N.Dealer
            Else
                rs!upah_normal_dealer = Null 'Upah Jualan Kepada N.Dealer
            End If
            If Frm83.TB33 <> vbNullString Then
                rs!upah_master_dealer = Format(Frm83.TB33, "0.00") 'Upah Jualan Kepada M.Dealer
            Else
                rs!upah_master_dealer = Null 'Upah Jualan Kepada M.Dealer
            End If
            rs!code_Supplier = Null 'Harga Jualan Kepada Pelanggan
            rs!HargaJualan_Member = Null 'Harga Jualan Kepada Ahli / Member
            rs!HargaJualan_Pengedar = Null 'Harga Jualan Kepada Pengedar
            rs!HargaJualan_RAF = Null 'Harga Jualan Kepada RAF
            rs!hargajualan_normal_dealer = Null 'Harga Jualan Kepada N.Dealer
            rs!hargajualan_master_dealer = Null 'Harga Jualan Kepada M.Dealer
        Else
            '+++++++++++ Special Request ++++++++++ Start
            'rs!Berat = Null 'Berat
            'rs!beza_berat = Null 'Baki Berat
            'If Frm83.TB8 <> vbNullString Then
            '    rs!Berat = Format(Frm83.TB8, "0.00") 'Berat
            '    rs!beza_berat = Format(Frm83.TB8, "0.00") 'Baki Berat
            'Else
                rs!Berat = Null 'Berat
                rs!beza_berat = Null 'Baki Berat
            'End If
            '+++++++++++ Special Request ++++++++++ End
            rs!kos_Belian_Gram = Null 'Harga Per Gram (Belian)
            rs!UPAH = Null 'Upah (RM)
            rs!harga_Per_Gram_Item = Null 'Harga Modal Per Gram [Average] (Harga Emas + Upah) Tanpa Tax
            rs!Upah_Jualan = Null 'Upah Jualan Kepada Pelanggan
            rs!Upah_Member = Null 'Upah Jualan Kepada Ahli / Member
            rs!Upah_Pengedar = Null 'Upah Jualan Kepada Pengedar
            rs!Upah_RAF = Null 'Upah Jualan Kepada RAF
            rs!upah_normal_dealer = Null 'Upah Jualan Kepada N.Dealer
            rs!upah_master_dealer = Null 'Upah Jualan Kepada M.Dealer
        End If
        If Frm83.CB5 = 1 Then
            If Frm83.TB24 <> vbNullString Then
                rs!code_Supplier = Format(Frm83.TB24, "0.00") 'Harga Jualan Kepada Pelanggan
            Else
                rs!code_Supplier = Null 'Harga Jualan Kepada Pelanggan
            End If
            If Frm83.TB25 <> vbNullString Then
                rs!HargaJualan_Member = Format(Frm83.TB25, "0.00") 'Harga Jualan Kepada Ahli / Member
            Else
                rs!HargaJualan_Member = Null 'Harga Jualan Kepada Ahli / Member
            End If
            If Frm83.TB26 <> vbNullString Then
                rs!HargaJualan_Pengedar = Format(Frm83.TB26, "0.00") 'Harga Jualan Kepada Pengedar
            Else
                rs!HargaJualan_Pengedar = Null 'Harga Jualan Kepada Pengedar
            End If
            If Frm83.TB31 <> vbNullString Then
                rs!HargaJualan_RAF = Format(Frm83.TB31, "0.00") 'Harga Jualan Kepada RAF
            Else
                rs!HargaJualan_RAF = Null 'Harga Jualan Kepada RAF
            End If
            If Frm83.TB32 <> vbNullString Then
                rs!hargajualan_normal_dealer = Format(Frm83.TB32, "0.00") 'Harga Jualan Kepada N.Dealer
            Else
                rs!hargajualan_normal_dealer = Null 'Harga Jualan Kepada N.Dealer
            End If
            If Frm83.TB33 <> vbNullString Then
                rs!hargajualan_master_dealer = Format(Frm83.TB33, "0.00") 'Harga Jualan Kepada M.Dealer
            Else
                rs!hargajualan_master_dealer = Null 'Harga Jualan Kepada M.Dealer
            End If
            
            '+++++++++++ Special Request ++++++++++ Start
            'rs!Berat = Null 'Berat
            'rs!beza_berat = Null 'Baki Berat
            'If Frm83.TB8 <> vbNullString Then
            '    rs!Berat = Format(Frm83.TB8, "0.00") 'Berat
            '    rs!beza_berat = Format(Frm83.TB8, "0.00") 'Baki Berat
            'Else
                rs!Berat = Null 'Berat
                rs!beza_berat = Null 'Baki Berat
            'End If
            '+++++++++++ Special Request ++++++++++ End
            rs!kos_Belian_Gram = Null 'Harga Per Gram (Belian)
            rs!UPAH = Null 'Upah (RM)
            rs!harga_Per_Gram_Item = Null 'Harga Modal Per Gram [Average] (Harga Emas + Upah) Tanpa Tax
            rs!Upah_Jualan = Null 'Upah Jualan Kepada Pelanggan
            rs!Upah_Member = Null 'Upah Jualan Kepada Ahli / Member
            rs!Upah_Pengedar = Null 'Upah Jualan Kepada Pengedar
            rs!Upah_RAF = Null 'Upah Jualan Kepada RAF
            rs!upah_normal_dealer = Null 'Upah Jualan Kepada N.Dealer
            rs!upah_master_dealer = Null 'Upah Jualan Kepada M.Dealer
        Else
            rs!code_Supplier = Null 'Harga Jualan Kepada Pelanggan
            rs!HargaJualan_Member = Null 'Harga Jualan Kepada Ahli / Member
            rs!HargaJualan_Pengedar = Null 'Harga Jualan Kepada Pengedar
            rs!HargaJualan_RAF = Null 'Harga Jualan Kepada RAF
            rs!hargajualan_normal_dealer = Null 'Harga Jualan Kepada N.Dealer
            rs!hargajualan_master_dealer = Null 'Harga Jualan Kepada M.Dealer
        End If
        If Frm83.CB12 = 0 Then 'GST pada harga barang
        
            If Frm83.TB10 <> vbNullString Then
                rs!kos_Belian_Item = Format(Frm83.TB10, "0.00") 'Harga Asal (RM)
            Else
                rs!kos_Belian_Item = Null 'Harga Asal (RM)
            End If
            
        End If
        If Frm83.CB12 = 1 Then 'GST pada upah
        
            If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) Then
                
                Frm83_LM_BERAT = Frm83.TB8 'Berat
                Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment
                
                rs!kos_Belian_Item = Format((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) + Frm83_LM_HARGA_TOTAL, "0.00") 'Harga Asal (RM)
                
            End If
        
        End If
        If Frm83.CB8 = 1 Then
            If Frm83.TB19 <> vbNullString Then
                rs!Spread = Format(Frm83.TB19, "0.00") 'Spread (%)
            Else
                rs!Spread = Null 'Spread (%)
            End If
        ElseIf Frm83.CB7 = 1 Then
            rs!Spread = Null 'Spread (%)
        End If
        If Frm83.TB21 <> vbNullString Then
            rs!harga_lepas_spread = Format(Frm83.TB21, "0.00") 'Harga asal ditolak spread (RM)
        Else
            rs!harga_lepas_spread = Null 'Harga asal ditolak spread (RM)
        End If
        If Frm83.TB22 <> vbNullString Then
            rs!adjustment = Format(Frm83.TB22, "0.00") 'Adjustment (RM)
        Else
            rs!adjustment = Null 'Adjustment (RM)
        End If
        If Frm83.CB12 = 0 Then 'GST pada harga barang
            If Frm83.TB20 <> vbNullString Then
                rs!kos_item_tanpa_tax = Format(Frm83.TB20, "0.00") 'Harga Barang + Upah Tanpa Tax
            Else
                rs!kos_item_tanpa_tax = Null 'Harga Barang + Upah Tanpa Tax
            End If
        End If
        If Frm83.CB12 = 1 Then 'GST pada upah
        
            If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) And (Frm83.TB22 <> vbNullString And IsNumeric(Frm83.TB22)) Then
                Frm83_LM_BERAT = Frm83.TB8 'Berat
                Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment
                
                rs!kos_item_tanpa_tax = Format(((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL, "0.00")  'Harga Barang + Upah Tanpa Tax
                
            End If
            
        End If
        'If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) Then
        '    Frm83_LM_BERAT = Frm83.TB8 'Berat
        '    Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
        '    Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax

        '    rs!harga_Per_Gram_Item = Format((Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST) / Frm83_LM_BERAT, "0.00") 'Harga Modal Per Gram [Average] (Harga Emas + Upah) Tanpa Tax
        'Else
        '    rs!harga_Per_Gram_Item = Null
        'End If
        If Frm83.TB12 <> vbNullString Then
            rs!dimension_Panjang = Frm83.TB12 'Panjang
        Else
            rs!dimension_Panjang = Null 'Panjang
        End If
        If Frm83.TB13 <> vbNullString Then
            rs!dimension_Lebar = Frm83.TB13 'Lebar
        Else
            rs!dimension_Lebar = Null 'Lebar
        End If
        If Frm83.TB14 <> vbNullString Then
            rs!dimension_Saiz = Frm83.TB14 'Saiz
        Else
            rs!dimension_Saiz = Null 'Saiz
        End If
        If Frm83.TB36 <> vbNullString Then 'Code 1
            rs!code1 = UCase(Frm83.TB36)
        Else
            rs!code1 = Null
        End If
        If Frm83.TB37 <> vbNullString Then 'Code 2
            rs!code2 = UCase(Frm83.TB37)
        Else
            rs!code2 = Null
        End If
        If Frm83.CBB5 <> vbNullString Then
            rs!dulang = Frm83.CBB5 'Dulang
        Else
            rs!dulang = Null 'Dulang
        End If
        If Frm83.TB34 <> vbNullString Then
            rs!no_cert = UCase(Frm83.TB34) 'No. Cert
        Else
            rs!no_cert = Null 'No. Cert
        End If
        If Frm83.TB16 <> vbNullString Then
            rs!remarks = UCase(Frm83.TB16) 'Remarks
        Else
            rs!remarks = Null 'Remarks
        End If

        If Frm83.CB2 = 1 Then
        
            rs!gst_ari_nashi = 0 'Status Cukai GST : 0 : ZR(L) , 1 : SR
            rs!kadar_gst = Null 'Kadar GST (%)
            rs!jumlah_gst = Null 'Jumlah Cukai GST (RM)
            rs!gst_included = Null '0 : Harga tidak termasuk GST , 1 : Harga termasuk GST
            
        ElseIf Frm83.CB3 = 1 Then
        
            rs!gst_included = 0 '0 : Harga tidak termasuk GST , 1 : Harga termasuk GST
            rs!gst_ari_nashi = 1 'Status Cukai GST : 0 : ZR(L) , 1 : SR
            rs!kadar_gst = Format(Frm83.L8_Text, "0.00") 'Kadar GST (%)
            rs!jumlah_gst = Format(Frm83.TB27, "0.00") 'Jumlah Cukai GST (RM)
          
        ElseIf Frm83.CB11 = 1 Then

            rs!gst_included = 1 '0 : Harga tidak termasuk GST , 1 : Harga termasuk GST
            rs!gst_ari_nashi = 1 'Status Cukai GST : 0 : ZR(L) , 1 : SR
            rs!kadar_gst = Format(Frm83.L8_Text, "0.00") 'Kadar GST (%)
            rs!jumlah_gst = Format(Frm83.TB27, "0.00") 'Jumlah Cukai GST (RM)
            
        End If
        
        If Frm83.L40_Text <> vbNullString Then
            rs!harga_tanpa_gst = Format(Frm83.L40_Text, "0.00") 'Harga Barang Tanpa Tax (kalau gst included)
        Else
            rs!harga_tanpa_gst = Null 'Harga Barang Tanpa Tax (kalau gst included)
        End If
        If Frm83.CB5 = 1 Then 'Barang Permata
        
            If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) Then
                Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                
                rs!harga_item = Format(Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST, "0.00") 'Jumlah Harga Modal Keseluruhan @ Harga Barang + Upah + GST (RM)
            End If
            
        End If
        
        If Frm83.CB4 = 1 Or Frm83.CB10 = 1 Then 'Barang Kemas / Gold Bar
        
            If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) And (Frm83.TB22 <> vbNullString And IsNumeric(Frm83.TB22)) Then
                Frm83_LM_BERAT = Frm83.TB8 'Berat
                Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment
                
                If Frm83.CB12 = 0 Then
                    rs!harga_item = Format(Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST, "0.00") 'Jumlah Harga Modal Keseluruhan @ Harga Barang + Upah + GST (RM)
                ElseIf Frm83.CB12 = 1 Then
                    rs!harga_item = Format(((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST, "0.00") 'Jumlah Harga Modal Keseluruhan @ Harga Barang + Upah + GST (RM)
                End If
                
            End If
            
        End If
        
'### Kod Bagi Status ###
'==========================================
'0 : Tidak Aktif (Tidak Perlu Buat Apa-Apa)
'10 : Aktif - Kemasukkan Data Baru
'2 : Barang Sudah Terjual ****Tidak Dibenarkan Untuk Diedit Atau Dipadamkan
'3 : Kemasukkan Data Baru
'4 : Data Diedit
'5 : Keluarkan Data Dari Database (Bagi Item Asal Yang Dipadamkan)
'6 : Ignore Kemasukkan Data Ke Dalam Database

'### Jenis ###
'0 : BK
'1 : Barang permata
'2 : Emas terpakai BK
'3 : Emas terpakai permata
'4 : gold Bar
'5 : Emas terpakai gold bar
'6 : Trade In BK
'7 : Trade In Barang Permata
'8 : Trade In Gold Bar

'=========================================================
'Frm83.L41_Text
'0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
'=========================================================
        rs!StatusItem = 3
        'If Frm83.L41_Text = 1 Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
        
            'If Frm83.CB8 = 1 Then 'Buyback / Trade in
        '        If Frm83.CB4 = 1 Then 'Barang kemas
        '            rs!jenis = 6
        '        ElseIf Frm83.CB5 = 1 Then 'Barang permata
        '            rs!jenis = 7
        '        End If
        '        If Frm83.CB10 = 1 Then 'Gold bar
        '            rs!jenis = 8
        '        End If
            'End If
        
        'ElseIf Frm83.L41_Text = 0 Or Frm83.L41_Text = 2 Then
        
            If Frm83.CB7 = 1 Then 'Penerimaan stok baru
                If Frm83.CB4 = 1 Then 'Barang kemas
                    rs!jenis = 0
                ElseIf Frm83.CB5 = 1 Then 'Barang permata
                    rs!jenis = 1
                End If
                If Frm83.CB10 = 1 Then 'Gold bar
                    rs!jenis = 4
                End If
            ElseIf Frm83.CB8 = 1 Then 'Buyback / Trade in
                If Frm83.CB4 = 1 Then 'Barang kemas
                    rs!jenis = 2
                ElseIf Frm83.CB5 = 1 Then 'Barang permata
                    rs!jenis = 3
                End If
                If Frm83.CB10 = 1 Then 'Gold bar
                    rs!jenis = 5
                End If
            End If
        
        'End If

        If Frm83.L41_Text = 0 Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
            
            rs!jenis_trade_in = 0 '0 : Belian emas dari pelanggan , 1 : Belian emas secara trade in
            
        ElseIf Frm83.L41_Text = 1 Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
        
            rs!jenis_trade_in = 1 '0 : Belian emas dari pelanggan , 1 : Belian emas secara trade in
            
        End If
        
        If Frm83.TB35 <> vbNullString Then
            rs!upah_per_gram = Format(Frm83.TB35, "0.00")
        Else
            rs!upah_per_gram = "0.00"
        End If
        If Frm83.CB14 = 1 Then
            rs!flag_upah = 0
        ElseIf Frm83.CB15 = 1 Then
            rs!flag_upah = 1
            
            If IsNumeric(Frm83.TB8) And Frm83.TB8 <> 0 Then
                Frm83_LM_BERAT = Frm83.TB8 'Berat
            End If
            
            Frm83_LM_UPAH = Frm83.TB4 'Upah
            
            rs!upah_per_gram = Format(Frm83_LM_UPAH / Frm83_LM_BERAT, "0.00")
        End If
        If Frm83.TB28 <> vbNullString Then
            rs!no_id_gst = UCase(Frm83.TB28)
        Else
            rs!no_id_gst = Null
        End If
        If Frm83.TB15 <> vbNullString Then
            rs!bill_No_Belian = UCase(Frm83.TB15)
        Else
            rs!bill_No_Belian = Null
        End If
        rs!tarikh_belian = Frm83.DTPicker1
        rs!terminal = G_TERMINAL
        rs!write_timestamp = Now 'Tarikh & Masa Data Dimasukkan
        rs!flag_image = 0
        rs!Image = Null
        
        If Frm83.L32_Text = 1 Then
            'Set rs2 = New ADODB.Recordset
            'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            'rs2.Open "select * from 1_image_item_temp where initial_flag='" & "1" & "'", cn, adOpenKeyset, adLockOptimistic
            
            'If Not rs2.EOF Then
            '    rs!flag_image = 1
            '    rs!Image = rs2!Image
            'End If
            
            'rs2.Close
            'Set rs2 = Nothing
        End If
        
        rs.Update
        DATA_SAVE = 1
        
        rs.Close
        Set rs = Nothing
        
        If DATA_SAVE = 1 Then
            'Frm83.L3_Text = Frm83_LM_No_SIRI + 1 'No. Siri Barcode
            
            'If Frm83.CB9 = 1 Then
            '    Frm83.TB7 = Format(Frm83.L3_Text, "000000") 'No. Siri Barcode
            'ElseIf Frm83.CB10 = 1 Then
            '    Frm83.TB7 = Format(Frm83.L3_Text, "000000") & "W" 'No. Siri Barcode
            'End If
            
            Call Frm83_Reset_Form
            Call Frm83_Senarai_Belian_Header
            Call Frm83_Senarai_Belian
            
            MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
            'Frm83.TB11.SetFocus
        End If
    End If
End If
End Sub
Private Sub CMD13_Click()
'On Error Resume Next
Dim Err(30)
Dim Frm83_LM_BERAT As Double
Dim Frm83_LM_UPAH As Double 'Upah
Dim Frm83_LM_HARGA_TOTAL As Double
Dim Frm83_LM_CUKAI_GST As Double
Dim Frm83_LM_HARGA_SEMASA As Double 'Harga Semasa
Dim Frm83_LM_ADJUSTMENT As Double 'Adjustment

Frm83_LM_BERAT = 1
Frm83_LM_UPAH = 0 'Upah
Frm83_LM_HARGA_TOTAL = 0
Frm83_LM_CUKAI_GST = 0
Frm83_LM_HARGA_SEMASA = 0 'Harga Semasa
Frm83_LM_ADJUSTMENT = 0 'Adjustment

x = 0
DATA_SAVE = 0

If Frm83.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Supplier]."
End If
If Frm83.CBB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Purity]."
End If
If Frm83.CBB3 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Kategori Produk]."
End If
If Frm83.TB1 = vbNullString Then
    x = x + 1
    Err(x) = "Tiada Maklumat [Kod Supplier]."
End If
If Frm83.TB2 = vbNullString Then
    x = x + 1
    Err(x) = "Tiada Maklumat [Kod Purity]."
End If
If Frm83.TB3 = vbNullString Then
    x = x + 1
    Err(x) = "Tiada Maklumat [Kod Kategori Produk]."
End If
'If Frm83.TB6 = vbNullString Or Frm83.TB7 = vbNullString Then
'    x = x + 1
'    Err(x) = "Maklumat [No. Siri Produk] Yang Tidak Lengkap."
'End If
If Frm83.TB36 <> vbNullString Then

    If InStr(1, Frm83.TB36, "*") <> 0 Or InStr(1, Frm83.TB36, "/") <> 0 Or InStr(1, Frm83.TB36, "\") <> 0 Or InStr(1, Frm83.TB36, "'") <> 0 Then

        x = x + 1
        Err(x) = "[Code 1] mengandungi simbol yang tidak sah."
        
    End If
    
End If
If Frm83.TB37 <> vbNullString Then

    If InStr(1, Frm83.TB37, "*") <> 0 Or InStr(1, Frm83.TB37, "/") <> 0 Or InStr(1, Frm83.TB37, "\") <> 0 Or InStr(1, Frm83.TB37, "'") <> 0 Then

        x = x + 1
        Err(x) = "[Code 2] mengandungi simbol yang tidak sah."
        
    End If
    
End If
If Frm83.CB9 = 1 And Frm83.CB4 = 0 And Frm83.CB5 = 0 Then
    x = x + 1
    Err(x) = "Sila Buat Pilihan Penerimaan [Barang Kemas] Atau [Barang Permata]."
End If
If Frm83.CB2 = 0 And Frm83.CB3 = 0 And Frm83.CB11 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan jenis cukai GST"
End If
If Frm83.CB4 = 1 Then
    If Frm83.TB8 = vbNullString Or (Frm83.TB8 <> vbNullString And Not IsNumeric(Frm83.TB8)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Berat]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB8 <> vbNullString Or (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) Then
        
        Frm83_LM_BERAT = Frm83.TB8
        
        If Frm83_LM_BERAT = 0 Then
            x = x + 1
            Err(x) = "Tiada data bagi berat barang. Berat 0 tidak dibenarkan."
        End If
    End If
    If Frm83.TB9 = vbNullString Or (Frm83.TB9 <> vbNullString And Not IsNumeric(Frm83.TB9)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Harga Semasa]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB4 = vbNullString Or (Frm83.TB4 <> vbNullString And Not IsNumeric(Frm83.TB4)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Upah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB24 = vbNullString Or (Frm83.TB24 <> vbNullString And Not IsNumeric(Frm83.TB24)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Pelanggan]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB25 = vbNullString Or (Frm83.TB25 <> vbNullString And Not IsNumeric(Frm83.TB25)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Ahli]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB26 = vbNullString Or (Frm83.TB26 <> vbNullString And Not IsNumeric(Frm83.TB26)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Silver]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB31 = vbNullString Or (Frm83.TB31 <> vbNullString And Not IsNumeric(Frm83.TB31)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Gold]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB32 = vbNullString Or (Frm83.TB32 <> vbNullString And Not IsNumeric(Frm83.TB32)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Platinum]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB33 = vbNullString Or (Frm83.TB33 <> vbNullString And Not IsNumeric(Frm83.TB33)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada M.Dealer]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
ElseIf Frm83.CB5 = 1 Then
    '+++++++++++ Special Request ++++++++++ Start
    'If Frm83.TB8 = vbNullString Or (Frm83.TB8 <> vbNullString And Not IsNumeric(Frm83.TB8)) Then
    '    x = x + 1
    '    Err(x) = "Sila Masukkan [Berat]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    'End If
    'If Frm83.TB8 <> vbNullString Or (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) Then
        
    '    Frm83_LM_BERAT = Frm83.TB8
        
    '    If Frm83_LM_BERAT = 0 Then
    '        x = x + 1
    '        Err(x) = "Tiada data bagi berat barang. Berat 0 tidak dibenarkan."
    '    End If
    'End If
    '+++++++++++ Special Request ++++++++++ End
    If Frm83.TB24 = vbNullString Or (Frm83.TB24 <> vbNullString And Not IsNumeric(Frm83.TB24)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Pelanggan]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB25 = vbNullString Or (Frm83.TB25 <> vbNullString And Not IsNumeric(Frm83.TB25)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Ahli]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB26 = vbNullString Or (Frm83.TB26 <> vbNullString And Not IsNumeric(Frm83.TB26)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Silver]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB31 = vbNullString Or (Frm83.TB31 <> vbNullString And Not IsNumeric(Frm83.TB31)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Gold]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB32 = vbNullString Or (Frm83.TB32 <> vbNullString And Not IsNumeric(Frm83.TB32)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Platinum]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB33 = vbNullString Or (Frm83.TB33 <> vbNullString And Not IsNumeric(Frm83.TB33)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada M.Dealer]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If
If Frm83.TB10 = vbNullString Or (Frm83.TB10 <> vbNullString And Not IsNumeric(Frm83.TB10)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Harga Asal]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB19 = vbNullString Or (Frm83.TB19 <> vbNullString And Not IsNumeric(Frm83.TB19)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Spread (%)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB20 = vbNullString Or (Frm83.TB20 <> vbNullString And Not IsNumeric(Frm83.TB20)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Harga Belian]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB21 = vbNullString Or (Frm83.TB21 <> vbNullString And Not IsNumeric(Frm83.TB21)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Harga Asal-Spread]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB22 = vbNullString Or (Frm83.TB22 <> vbNullString And Not IsNumeric(Frm83.TB22)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Adjusment]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.CB8 = 1 Then
    If Frm83.TB19 = vbNullString Or (Frm83.TB19 <> vbNullString And Not IsNumeric(Frm83.TB19)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Spread]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If
If Frm83.CBB5 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Dulang]."
End If
If Frm83.CB3 = 1 Then
    If Frm83.TB27 = vbNullString Or (Frm83.TB27 <> vbNullString And Not IsNumeric(Frm83.TB27)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat GST"
    End If
End If
If Frm83.CB4 = 1 Then
    If Frm83.CB14 = 0 And Frm83.CB15 = 1 Then
        If Frm83.TB35 = vbNullString Or (Frm83.TB35 <> vbNullString And Not IsNumeric(Frm83.TB35)) Then
            x = x + 1
            Err(x) = "Sila buat tetapan pengiraan upah dari supplier"
        End If
    End If
End If
If Frm83.CB14 = 1 Then
    If Frm83.TB35 = vbNullString Or (Frm83.TB35 <> vbNullString And Not IsNumeric(Frm83.TB35)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Upah per gram]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If
If Frm83.TB15 <> vbNullString Then

    If InStr(1, Frm83.TB15, "*") <> 0 Or InStr(1, Frm83.TB15, "/") <> 0 Or InStr(1, Frm83.TB15, "\") <> 0 Or InStr(1, Frm83.TB15, "'") <> 0 Then

        x = x + 1
        Err(x) = "[No. Invoice] mengandungi simbol yang tidak sah."
        
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
    Note = "Adakah anda ingin masukkan data barang ini ke dalam senarai belian?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
        Frm83_LM_STATUS = 1
    
        'Frm83_LM_No_SIRI = Frm83.L3_Text 'Frm83.TB7 'No. Turutan No. Siri
        Frm83_LM_No_RUJUKAN_BELIAN = Frm83.L9_Text 'No. Rujukan Belian
        
'Re_Gen_Code:
        
        'Set rs = New ADODB.Recordset
        'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        'rs.Open "select * from Data_Database where barcode='" & Format(Frm83_LM_No_SIRI, "000000") & "'", cn, adOpenKeyset, adLockOptimistic
        
        'If Not rs.EOF Then
        '    Frm83_LM_No_SIRI = Frm83_LM_No_SIRI + 1
            
        '    rs.Close
        '    Set rs = Nothing
        '    GoTo Re_Gen_Code:
        'End If
        
        'rs.Close
        'Set rs = Nothing
          
'###Masukkan Data Belian Ke Dalam Database### - Start

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_BELIAN_TEMP & " where ID='" & Frm83.L13_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Frm83.L4_Text <> vbNullString Then
                rs!supplier_ID = Frm83.L4_Text 'No. ID Bagi Supplier
            Else
                rs!supplier_ID = Null 'No. ID Bagi Supplier
            End If
            If Frm83.CBB1 <> vbNullString Then
                rs!nama_Supplier = Frm83.CBB1 'Nama Supplier
            Else
                rs!nama_Supplier = Null 'Nama Supplier
            End If
            If Frm83.TB1 <> vbNullString Then
                rs!Kod_Supplier = Frm83.TB1 'Kod Supplier
            Else
                rs!Kod_Supplier = Null 'Kod Supplier
            End If
            If Frm83.L5_Text <> vbNullString Then
                rs!purity_ID = Frm83.L5_Text 'No. ID Bagi Purity
            Else
                rs!purity_ID = Null 'No. ID Bagi Purity
            End If
            If Frm83.CBB2 <> vbNullString Then
                rs!purity = Frm83.CBB2 'Purity
            Else
                rs!purity = Null 'Purity
            End If
            If Frm83.TB2 <> vbNullString Then
                rs!kod_Purity = Frm83.TB2 'Kod Purity
            Else
                rs!kod_Purity = Null 'Kod Purity
            End If
            If Frm83.L6_Text <> vbNullString Then
                rs!kategori_produk_ID = Frm83.L6_Text 'No. ID Bagi Kategori Produk
            Else
                rs!kategori_produk_ID = Null 'No. ID Bagi Kategori Produk
            End If
            If Frm83.CBB3 <> vbNullString Then
                rs!kategori_Produk = Frm83.CBB3 'Kategori Produk
            Else
                rs!kategori_Produk = Null 'Kategori Produk
            End If
            If Frm83.TB3 <> vbNullString Then
                rs!Kod_Kategori_Produk = Frm83.TB3 'Kod Kategori Produk
            Else
                rs!Kod_Kategori_Produk = Null 'Kod Kategori Produk
            End If
            'If Frm83.TB7 <> vbNullString Then
                'If Frm83.CB9 = 1 Then
            '        rs!Barcode = Frm83.TB7 'Format(Frm83_LM_No_SIRI, "000000") 'No. Barcode (6 Digit Terakhir)
                'ElseIf Frm83.CB10 = 1 Then
                '    rs!Barcode = Format(Frm83_LM_No_SIRI, "000000") & "W" 'No. Barcode (6 Digit Terakhir)
                'End If
            'Else
            '    rs!Barcode = Null 'No. Barcode (6 Digit Terakhir)
            'End If
            'rs!no_siri_Produk = Frm83.TB6 & Frm83.TB7 'No. Siri Produk
            'If Frm83.CB9 = 1 Then
            '    rs!no_siri_Produk = Frm83.TB6 & Format(Frm83_LM_No_SIRI, "000000")  'No. Siri Produk
            'ElseIf Frm83.CB10 = 1 Then
            '    rs!no_siri_Produk = Frm83.TB6 & Format(Frm83_LM_No_SIRI, "000000") & "W"  'No. Siri Produk
            'End If
            If Frm83.CB12 = 0 Then '0 : Pengiraan GST pada harga belian barang , 1 : Pengiraan GST pada upah sahaja
                rs!gst_barang_atau_upah = 0
            ElseIf Frm83.CB12 = 1 Then
                rs!gst_barang_atau_upah = 1
            End If
            If Frm83.CB4 = 1 Or Frm83.CB10 = 1 Then
                If Frm83.TB8 <> vbNullString Then
                    rs!Berat = Format(Frm83.TB8, "0.00") 'Berat
                    rs!beza_berat = Format(Frm83.TB8, "0.00") 'Baki Berat
                Else
                    rs!Berat = Null 'Berat
                    rs!beza_berat = Null 'Baki Berat
                End If
                If Frm83.TB29 <> vbNullString Then
                    rs!riyal = Format(Frm83.TB29, "0.00") 'Berat Riyal
                Else
                    rs!riyal = Null 'Berat Riyal
                End If
                If Frm83.TB9 <> vbNullString Then
                    rs!kos_Belian_Gram = Format(Frm83.TB9, "0.00") 'Harga Per Gram (Belian)
                Else
                    rs!kos_Belian_Gram = Null 'Harga Per Gram (Belian)
                End If
                If Frm83.TB4 <> vbNullString Then
                    rs!UPAH = Frm83.TB4 'Upah (RM)
                Else
                    rs!UPAH = Null 'Upah (RM)
                End If
                If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) And (Frm83.TB22 <> vbNullString And IsNumeric(Frm83.TB22)) Then
                    Frm83_LM_BERAT = Frm83.TB8 'Berat
                    Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                    Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                    Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                    Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment
    
                    If Frm83.CB12 = 0 Then
                        rs!harga_Per_Gram_Item = Format((Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST) / Frm83_LM_BERAT, "0.00") 'Harga Modal Per Gram [Average] (Harga Emas + Upah) Tanpa Tax
                        rs!harga_per_gram_tanpa_gst = Format(Frm83_LM_HARGA_TOTAL / Frm83_LM_BERAT, "0.00")
                    ElseIf Frm83.CB12 = 1 Then
                        rs!harga_Per_Gram_Item = Format((((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST) / Frm83_LM_BERAT, "0.00")
                        rs!harga_per_gram_tanpa_gst = Format((((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL) / Frm83_LM_BERAT, "0.00")
                    End If
                Else
                    rs!harga_Per_Gram_Item = Null
                End If

                If Frm83.TB24 <> vbNullString Then
                    rs!Upah_Jualan = Format(Frm83.TB24, "0.00") 'Upah Jualan Kepada Pelanggan
                Else
                    rs!Upah_Jualan = Null 'Upah Jualan Kepada Pelanggan
                End If
                If Frm83.TB25 <> vbNullString Then
                    rs!Upah_Member = Format(Frm83.TB25, "0.00") 'Upah Jualan Kepada Ahli / Member
                Else
                    rs!Upah_Member = Null 'Upah Jualan Kepada Ahli / Member
                End If
                If Frm83.TB26 <> vbNullString Then
                    rs!Upah_Pengedar = Format(Frm83.TB26, "0.00") 'Upah Jualan Kepada Pengedar
                Else
                    rs!Upah_Pengedar = Null 'Upah Jualan Kepada Pengedar
                End If
                If Frm83.TB31 <> vbNullString Then
                    rs!Upah_RAF = Format(Frm83.TB31, "0.00") 'Upah Jualan Kepada RAF
                Else
                    rs!Upah_RAF = Null 'Upah Jualan Kepada RAF
                End If
                If Frm83.TB32 <> vbNullString Then
                    rs!upah_normal_dealer = Format(Frm83.TB32, "0.00") 'Upah Jualan Kepada N.Dealer
                Else
                    rs!upah_normal_dealer = Null 'Upah Jualan Kepada N.Dealer
                End If
                If Frm83.TB33 <> vbNullString Then
                    rs!upah_master_dealer = Format(Frm83.TB33, "0.00") 'Upah Jualan Kepada M.Dealer
                Else
                    rs!upah_master_dealer = Null 'Upah Jualan Kepada M.Dealer
                End If
                rs!code_Supplier = Null 'Harga Jualan Kepada Pelanggan
                rs!HargaJualan_Member = Null 'Harga Jualan Kepada Ahli / Member
                rs!HargaJualan_Pengedar = Null 'Harga Jualan Kepada Pengedar
                rs!HargaJualan_RAF = Null 'Harga Jualan Kepada RAF
                rs!hargajualan_normal_dealer = Null 'Harga Jualan Kepada N.Dealer
                rs!hargajualan_master_dealer = Null 'Harga Jualan Kepada M.Dealer
            Else
                rs!Berat = Null 'Berat
                rs!beza_berat = Null 'Baki Berat
                rs!kos_Belian_Gram = Null 'Harga Per Gram (Belian)
                rs!UPAH = Null 'Upah (RM)
                rs!harga_Per_Gram_Item = Null 'Harga Modal Per Gram [Average] (Harga Emas + Upah) Tanpa Tax
                rs!Upah_Jualan = Null 'Upah Jualan Kepada Pelanggan
                rs!Upah_Member = Null 'Upah Jualan Kepada Ahli / Member
                rs!Upah_Pengedar = Null 'Upah Jualan Kepada Pengedar
                rs!Upah_RAF = Null 'Upah Jualan Kepada RAF
                rs!upah_normal_dealer = Null 'Upah Jualan Kepada N.Dealer
                rs!upah_master_dealer = Null 'Upah Jualan Kepada M.Dealer
            End If
            If Frm83.CB5 = 1 Then
                If Frm83.TB24 <> vbNullString Then
                    rs!code_Supplier = Format(Frm83.TB24, "0.00") 'Harga Jualan Kepada Pelanggan
                Else
                    rs!code_Supplier = Null 'Harga Jualan Kepada Pelanggan
                End If
                If Frm83.TB25 <> vbNullString Then
                    rs!HargaJualan_Member = Format(Frm83.TB25, "0.00") 'Harga Jualan Kepada Ahli / Member
                Else
                    rs!HargaJualan_Member = Null 'Harga Jualan Kepada Ahli / Member
                End If
                If Frm83.TB26 <> vbNullString Then
                    rs!HargaJualan_Pengedar = Format(Frm83.TB26, "0.00") 'Harga Jualan Kepada Pengedar
                Else
                    rs!HargaJualan_Pengedar = Null 'Harga Jualan Kepada Pengedar
                End If
                If Frm83.TB31 <> vbNullString Then
                    rs!HargaJualan_RAF = Format(Frm83.TB31, "0.00") 'Harga Jualan Kepada RAF
                Else
                    rs!HargaJualan_RAF = Null 'Harga Jualan Kepada RAF
                End If
                If Frm83.TB32 <> vbNullString Then
                    rs!hargajualan_normal_dealer = Format(Frm83.TB32, "0.00") 'Harga Jualan Kepada N.Dealer
                Else
                    rs!hargajualan_normal_dealer = Null 'Harga Jualan Kepada N.Dealer
                End If
                If Frm83.TB33 <> vbNullString Then
                    rs!hargajualan_master_dealer = Format(Frm83.TB33, "0.00") 'Harga Jualan Kepada M.Dealer
                Else
                    rs!hargajualan_master_dealer = Null 'Harga Jualan Kepada M.Dealer
                End If
                
                '+++++++++++ Special Request ++++++++++ Start
                'rs!Berat = Null 'Berat
                'rs!beza_berat = Null 'Baki Berat
                'If Frm83.TB8 <> vbNullString Then
                '    rs!Berat = Format(Frm83.TB8, "0.00") 'Berat
                '    rs!beza_berat = Format(Frm83.TB8, "0.00") 'Baki Berat
                'Else
                    rs!Berat = Null 'Berat
                    rs!beza_berat = Null 'Baki Berat
                'End If
                '+++++++++++ Special Request ++++++++++ End
                rs!kos_Belian_Gram = Null 'Harga Per Gram (Belian)
                rs!UPAH = Null 'Upah (RM)
                rs!harga_Per_Gram_Item = Null 'Harga Modal Per Gram [Average] (Harga Emas + Upah) Tanpa Tax
                rs!Upah_Jualan = Null 'Upah Jualan Kepada Pelanggan
                rs!Upah_Member = Null 'Upah Jualan Kepada Ahli / Member
                rs!Upah_Pengedar = Null 'Upah Jualan Kepada Pengedar
                rs!Upah_RAF = Null 'Upah Jualan Kepada RAF
                rs!upah_normal_dealer = Null 'Upah Jualan Kepada N.Dealer
                rs!upah_master_dealer = Null 'Upah Jualan Kepada M.Dealer
            Else
                rs!code_Supplier = Null 'Harga Jualan Kepada Pelanggan
                rs!HargaJualan_Member = Null 'Harga Jualan Kepada Ahli / Member
                rs!HargaJualan_Pengedar = Null 'Harga Jualan Kepada Pengedar
                rs!HargaJualan_RAF = Null 'Harga Jualan Kepada RAF
                rs!hargajualan_normal_dealer = Null 'Harga Jualan Kepada N.Dealer
                rs!hargajualan_master_dealer = Null 'Harga Jualan Kepada M.Dealer
            End If
            
            If Frm83.CB12 = 0 Then 'GST pada harga barang
            
                If Frm83.TB10 <> vbNullString Then
                    rs!kos_Belian_Item = Format(Frm83.TB10, "0.00") 'Harga Asal (RM)
                Else
                    rs!kos_Belian_Item = Null 'Harga Asal (RM)
                End If
                
            End If
            If Frm83.CB12 = 1 Then 'GST pada upah
            
                If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) Then
                    
                    Frm83_LM_BERAT = Frm83.TB8 'Berat
                    Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                    Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                    Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                    Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment
                    
                    rs!kos_Belian_Item = Format((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) + Frm83_LM_HARGA_TOTAL, "0.00") 'Harga Asal (RM)
                    
                End If
            
            End If
            If Frm83.CB8 = 1 Then
                If Frm83.TB19 <> vbNullString Then
                    rs!Spread = Format(Frm83.TB19, "0.00") 'Spread (%)
                Else
                    rs!Spread = Null 'Spread (%)
                End If
            ElseIf Frm83.CB7 = 1 Then
                rs!Spread = Null 'Spread (%)
            End If
            If Frm83.TB21 <> vbNullString Then
                rs!harga_lepas_spread = Format(Frm83.TB21, "0.00") 'Harga asal ditolak spread (RM)
            Else
                rs!harga_lepas_spread = Null 'Harga asal ditolak spread (RM)
            End If
            If Frm83.TB22 <> vbNullString Then
                rs!adjustment = Format(Frm83.TB22, "0.00") 'Adjustment (RM)
            Else
                rs!adjustment = Null 'Adjustment (RM)
            End If
            If Frm83.CB12 = 0 Then 'GST pada harga barang
                If Frm83.TB20 <> vbNullString Then
                    rs!kos_item_tanpa_tax = Format(Frm83.TB20, "0.00") 'Harga Barang + Upah Tanpa Tax
                Else
                    rs!kos_item_tanpa_tax = Null 'Harga Barang + Upah Tanpa Tax
                End If
            End If
            If Frm83.CB12 = 1 Then 'GST pada upah
            
                If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) And (Frm83.TB22 <> vbNullString And IsNumeric(Frm83.TB22)) Then
                    Frm83_LM_BERAT = Frm83.TB8 'Berat
                    Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                    Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                    Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                    Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment
                    
                    rs!kos_item_tanpa_tax = Format(((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL, "0.00")  'Harga Barang + Upah Tanpa Tax
                    
                End If
                
            End If
            'If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) Then
            '    Frm83_LM_BERAT = Frm83.TB8 'Berat
            '    Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
            '    Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax

            '    rs!harga_Per_Gram_Item = Format((Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST) / Frm83_LM_BERAT, "0.00") 'Harga Modal Per Gram [Average] (Harga Emas + Upah) Tanpa Tax
            'Else
            '    rs!harga_Per_Gram_Item = Null
            'End If
            If Frm83.TB12 <> vbNullString Then
                rs!dimension_Panjang = Frm83.TB12 'Panjang
            Else
                rs!dimension_Panjang = Null 'Panjang
            End If
            If Frm83.TB13 <> vbNullString Then
                rs!dimension_Lebar = Frm83.TB13 'Lebar
            Else
                rs!dimension_Lebar = Null 'Lebar
            End If
            If Frm83.TB14 <> vbNullString Then
                rs!dimension_Saiz = Frm83.TB14 'Saiz
            Else
                rs!dimension_Saiz = Null 'Saiz
            End If
            If Frm83.TB36 <> vbNullString Then 'Code 1
                rs!code1 = UCase(Frm83.TB36)
            Else
                rs!code1 = Null
            End If
            If Frm83.TB37 <> vbNullString Then 'Code 2
                rs!code2 = UCase(Frm83.TB37)
            Else
                rs!code2 = Null
            End If
            If Frm83.CBB5 <> vbNullString Then
                rs!dulang = Frm83.CBB5 'Dulang
            Else
                rs!dulang = Null 'Dulang
            End If
            If Frm83.TB34 <> vbNullString Then
                rs!no_cert = UCase(Frm83.TB34) 'No. Cert
            Else
                rs!no_cert = Null 'No. Cert
            End If
            If Frm83.TB16 <> vbNullString Then
                rs!remarks = UCase(Frm83.TB16) 'Remarks
            Else
                rs!remarks = Null 'Remarks
            End If

            If Frm83.CB2 = 1 Then
            
                rs!gst_ari_nashi = 0 'Status Cukai GST : 0 : ZR(L) , 1 : SR
                rs!kadar_gst = Null 'Kadar GST (%)
                rs!jumlah_gst = Null 'Jumlah Cukai GST (RM)
                rs!gst_included = Null '0 : Harga tidak termasuk GST , 1 : Harga termasuk GST
                
            ElseIf Frm83.CB3 = 1 Then
            
                rs!gst_included = 0 '0 : Harga tidak termasuk GST , 1 : Harga termasuk GST
                rs!gst_ari_nashi = 1 'Status Cukai GST : 0 : ZR(L) , 1 : SR
                rs!kadar_gst = Format(Frm83.L8_Text, "0.00") 'Kadar GST (%)
                rs!jumlah_gst = Format(Frm83.TB27, "0.00") 'Jumlah Cukai GST (RM)
              
            ElseIf Frm83.CB11 = 1 Then
    
                rs!gst_included = 1 '0 : Harga tidak termasuk GST , 1 : Harga termasuk GST
                rs!gst_ari_nashi = 1 'Status Cukai GST : 0 : ZR(L) , 1 : SR
                rs!kadar_gst = Format(Frm83.L8_Text, "0.00") 'Kadar GST (%)
                rs!jumlah_gst = Format(Frm83.TB27, "0.00") 'Jumlah Cukai GST (RM)
                
            End If
        
            If Frm83.L40_Text <> vbNullString Then
                rs!harga_tanpa_gst = Format(Frm83.L40_Text, "0.00") 'Harga Barang Tanpa Tax (kalau gst included)
            Else
                rs!harga_tanpa_gst = Null 'Harga Barang Tanpa Tax (kalau gst included)
            End If
            If Frm83.CB5 = 1 Then 'Barang Permata
            
                If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) Then
                    Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                    Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                    
                    rs!harga_item = Format(Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST, "0.00") 'Jumlah Harga Modal Keseluruhan @ Harga Barang + Upah + GST (RM)
                End If
                
            End If
            
            If Frm83.CB4 = 1 Or Frm83.CB10 = 1 Then 'Barang Kemas / Gold Bar
            
                If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) And (Frm83.TB22 <> vbNullString And IsNumeric(Frm83.TB22)) Then
                    Frm83_LM_BERAT = Frm83.TB8 'Berat
                    Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                    Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                    Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                    Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment
                    
                    If Frm83.CB12 = 0 Then
                        rs!harga_item = Format(Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST, "0.00") 'Jumlah Harga Modal Keseluruhan @ Harga Barang + Upah + GST (RM)
                    ElseIf Frm83.CB12 = 1 Then
                        rs!harga_item = Format(((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST, "0.00") 'Jumlah Harga Modal Keseluruhan @ Harga Barang + Upah + GST (RM)
                    End If
                    
                End If
                
            End If

'### Kod Bagi Status ###
'==========================================
'0 : Tidak Aktif (Tidak Perlu Buat Apa-Apa)
'1 : Aktif - Kemasukkan Data Baru
'2 : Barang Sudah Terjual ****Tidak Dibenarkan Untuk Diedit Atau Dipadamkan
'3 : Kemasukkan Data Baru
'4 : Data Diedit
'5 : Keluarkan Data Dari Database (Bagi Item Asal Yang Dipadamkan)
'6 : Ignore Kemasukkan Data Ke Dalam Database
            If rs!StatusItem = "10" Then
                Frm83_LM_STATUS = "4"
            ElseIf rs!StatusItem = "3" Then
                Frm83_LM_STATUS = "3"
            ElseIf rs!StatusItem = "4" Then
                Frm83_LM_STATUS = "4"
            End If
            rs!StatusItem = Frm83_LM_STATUS
            
'### Jenis ###
'0 : BK
'1 : Barang permata
'2 : Emas terpakai BK
'3 : Emas terpakai permata
'4 : gold Bar
'5 : Emas terpakai gold bar
'6 : Trade In BK
'7 : Trade In Barang Permata
'8 : Trade In Gold Bar

'=========================================================
'Frm83.L41_Text
'0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
'=========================================================

            'If Frm83.L41_Text = 1 Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
            
                'If Frm83.CB8 = 1 Then 'Buyback / Trade in
            '        If Frm83.CB4 = 1 Then 'Barang kemas
            '            rs!jenis = 6
            '        ElseIf Frm83.CB5 = 1 Then 'Barang permata
            '            rs!jenis = 7
            '        End If
            '        If Frm83.CB10 = 1 Then 'Gold bar
            '            rs!jenis = 8
            '        End If
            '    'End If
            
            'ElseIf Frm83.L41_Text = 0 Or Frm83.L41_Text = 2 Then
            
                If Frm83.CB7 = 1 Then 'Penerimaan stok baru
                    If Frm83.CB4 = 1 Then 'Barang kemas
                        rs!jenis = 0
                    ElseIf Frm83.CB5 = 1 Then 'Barang permata
                        rs!jenis = 1
                    End If
                    If Frm83.CB10 = 1 Then 'Gold bar
                        rs!jenis = 4
                    End If
                ElseIf Frm83.CB8 = 1 Then 'Buyback / Trade in
                    If Frm83.CB4 = 1 Then 'Barang kemas
                        rs!jenis = 2
                    ElseIf Frm83.CB5 = 1 Then 'Barang permata
                        rs!jenis = 3
                    End If
                    If Frm83.CB10 = 1 Then 'Gold bar
                        rs!jenis = 5
                    End If
                End If
            
            'End If
            
            If Frm83.L41_Text = 0 Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
                
                rs!jenis_trade_in = 0 '0 : Belian emas dari pelanggan , 1 : Belian emas secara trade in
                
            ElseIf Frm83.L41_Text = 1 Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
            
                rs!jenis_trade_in = 1 '0 : Belian emas dari pelanggan , 1 : Belian emas secara trade in
                
            End If
            rs!flag_image = 0
            rs!Image = Null

            If Frm83.TB35 <> vbNullString Then
                rs!upah_per_gram = Format(Frm83.TB35, "0.00")
            Else
                rs!upah_per_gram = "0.00"
            End If
            If Frm83.CB14 = 1 Then
                rs!flag_upah = 0
            ElseIf Frm83.CB15 = 1 Then
                rs!flag_upah = 1
                
                If IsNumeric(Frm83.TB8) And Frm83.TB8 <> 0 Then
                    Frm83_LM_BERAT = Frm83.TB8 'Berat
                End If
                
                Frm83_LM_UPAH = Frm83.TB4 'Upah
                
                rs!upah_per_gram = Format(Frm83_LM_UPAH / Frm83_LM_BERAT, "0.00")
            End If
            If Frm83.TB28 <> vbNullString Then
                rs!no_id_gst = UCase(Frm83.TB28)
            Else
                rs!no_id_gst = Null
            End If
            If Frm83.TB15 <> vbNullString Then
                rs!bill_No_Belian = UCase(Frm83.TB15)
            Else
                rs!bill_No_Belian = Null
            End If
            rs!tarikh_belian = Frm83.DTPicker1
            rs!terminal = G_TERMINAL
            If Frm83.L32_Text = 1 Then
                'Set rs2 = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                'rs2.Open "select * from 1_image_item_temp where initial_flag='" & "1" & "'", cn, adOpenKeyset, adLockOptimistic
                
                'If Not rs2.EOF Then
                '    rs!flag_image = 1
                '    rs!Image = rs2!Image
                'End If
                
                'rs2.Close
                'Set rs2 = Nothing
            End If
            
            rs.Update
            DATA_SAVE = 1
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_SAVE = 1 Then
            Frm83.CMD12.Visible = True
            Frm83.CMD13.Visible = False
            Frm83.CMD14.Visible = False
            
            Call Frm83_Cancel_Edit
            
            Call Frm83_Reset_Form
            
            GM_NEXT_PREV = 2
            
            Call Frm83_Senarai_Belian_Header
            Call Frm83_Senarai_Belian
            
            MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
            'Frm83.TB11.SetFocus
        End If
    End If
End If
End Sub
Private Sub CMD14_Click()
'on error resume next
Call Frm83_Cancel_Edit

If Frm83.CB9 = 1 Then

    Frm83.CB5 = 0
    
    Frm83.TB8.Locked = False
    Frm83.TB9.Locked = False
    Frm83.TB4.Locked = False
    
    Frm83.TB8.BackColor = &HFFFFFF
    Frm83.TB9.BackColor = &HFFFFFF
    Frm83.TB4.BackColor = &HFFFFFF
    
    Frm83.L27_Text = "Upah Jualan Pelanggan    RM"
    Frm83.L28_Text = "Upah Jualan Ahli               RM"
    Frm83.L29_Text = "Upah Jualan Silver            RM"

    Frm83.CB4.Enabled = True
    Frm83.CB5.Enabled = True

End If

Frm83.TB8 = "0.00"
Frm83.TB9 = "0.00"
Frm83.TB4 = 0
Frm83.TB10 = "0.00"
Frm83.TB35 = "0.00"
End Sub
Private Sub CMD16_Click()
'On Error Resume Next
Frm90.Show
Frm90.L3_Text = 0
Frm83.Hide

If Frm83.L32_Text = 1 Then
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 1_image_item_temp where initial_flag='" & "1" & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        Set picstrm = New ADODB.Stream
        picstrm.Type = adTypeBinary
        picstrm.Open
        If IsNull(rs!Image) = False Then
            picstrm.Write rs!Image
            picstrm.SaveToFile "" & App.Path & "\picture\a.jpg", adSaveCreateOverWrite
            Frm90.Image1.Picture = LoadPicture("" & App.Path & "\picture\a.jpg")
            picstrm.Close
            Set picstrm = Nothing
        End If
        
    End If
    
    rs.Close
    Set rs = Nothing
End If
End Sub
Private Sub CMD17_Click()

End Sub
Private Sub CMD18_Click()

End Sub

Private Sub CMD2_Click()
'On Error Resume Next
If Frm83.L10_Text <> 0 Then
    
    If MDI_frm1.L5_Text <> 4 Then
    
        Note = "Adakah mempunyai data yang belum disimpan." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Adakah anda ingin keluar dari menu ini?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
        
            If Frm83.L41_Text = "1" Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
                
                Frm84.Show
                Frm83.Hide
                
            ElseIf Frm83.L41_Text = "0" Or Frm83.L41_Text = "2" Then
            
                'Frm16.Show
                Unload Frm26
                Unload Frm27
                Unload Frm28
                Unload Frm83
                MDI_frm1.L5_Text = 0
                
            End If
            
        End If
        
    Else
    
        Frm84.Show
        Frm83.Hide
                
    End If
    
Else

    If Frm83.L41_Text = "1" Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
        
        Frm84.Show
        Frm83.Hide
        
    ElseIf Frm83.L41_Text = "0" Or Frm83.L41_Text = "2" Then
    
        'Frm16.Show
        Unload Frm26
        Unload Frm27
        Unload Frm28
        Unload Frm83
        MDI_frm1.L5_Text = 0
        
    End If

End If
End Sub

Private Sub CMD20_Click()
'On Error Resume Next
Dim rs2 As ADODB.Recordset

Dim Data_Err(35)
Dim Frm83_LM_BERAT As Double
Dim Frm83_LM_HARGA_TOTAL As Double
Dim Frm83_LM_CUKAI_GST As Double
Dim Frm83_LM_HARGA_SEMASA As Double 'Harga Semasa
Dim Frm83_LM_ADJUSTMENT As Double 'Adjustment
Dim Frm83_LM_UPAH As Double 'Upah

Frm83_LM_HARGA_SEMASA = 0 'Harga Semasa
Frm83_LM_ADJUSTMENT = 0 'Adjustment
Frm83_LM_CUKAI_GST = 0
Frm83_LM_HARGA_TOTAL = 0
Frm83_LM_BERAT = 1
Frm83_LM_UPAH = 0 'Upah

x = 0
DATA_SAVE = 0

If Frm83.CBB6 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Sila Pilih [Nama Pekerja]."
End If
If Frm83.CBB1 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Sila Pilih [Supplier]."
End If
If Frm83.CBB2 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Sila Pilih [Purity]."
End If
If Frm83.CBB3 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Sila Pilih [Kategori Produk]."
End If
If Frm83.TB1 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Tiada Maklumat [Kod Supplier]."
End If
If Frm83.TB2 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Tiada Maklumat [Kod Purity]."
End If
If Frm83.TB3 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Tiada Maklumat [Kod Kategori Produk]."
End If
'If Frm83.TB6 = vbNullString Or Frm83.TB7 = vbNullString Then
'    x = x + 1
'    Data_Err(x) = "Maklumat [No. Siri Produk] Yang Tidak Lengkap."
'End If
If Frm83.CB9 = 1 And Frm83.CB4 = 0 And Frm83.CB5 = 0 Then
    x = x + 1
    Data_Err(x) = "Sila Buat Pilihan Penerimaan [Barang Kemas] Atau [Barang Permata]."
End If
If Frm83.TB10 = vbNullString Or (Frm83.TB10 <> vbNullString And Not IsNumeric(Frm83.TB10)) Then
    x = x + 1
    Data_Err(x) = "Sila Masukkan [Harga Asal]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB19 = vbNullString Or (Frm83.TB19 <> vbNullString And Not IsNumeric(Frm83.TB19)) Then
    x = x + 1
    Data_Err(x) = "Sila Masukkan [Spread (%)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB20 = vbNullString Or (Frm83.TB20 <> vbNullString And Not IsNumeric(Frm83.TB20)) Then
    x = x + 1
    Data_Err(x) = "Sila Masukkan [Harga Belian]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB21 = vbNullString Or (Frm83.TB21 <> vbNullString And Not IsNumeric(Frm83.TB21)) Then
    x = x + 1
    Data_Err(x) = "Sila Masukkan [Harga Asal-Spread]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB22 = vbNullString Or (Frm83.TB22 <> vbNullString And Not IsNumeric(Frm83.TB22)) Then
    x = x + 1
    Data_Err(x) = "Sila Masukkan [Adjusment]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.CB2 = 0 And Frm83.CB3 = 0 And Frm83.CB11 = 0 Then
    x = x + 1
    Data_Err(x) = "Sila buat pilihan jenis cukai GST"
End If
If Frm83.TB36 <> vbNullString Then

    If InStr(1, Frm83.TB36, "*") <> 0 Or InStr(1, Frm83.TB36, "/") <> 0 Or InStr(1, Frm83.TB36, "\") <> 0 Or InStr(1, Frm83.TB36, "'") <> 0 Then

        x = x + 1
        Data_Err(x) = "[Code 1] mengandungi simbol yang tidak sah."
        
    End If
    
End If
If Frm83.TB37 <> vbNullString Then

    If InStr(1, Frm83.TB37, "*") <> 0 Or InStr(1, Frm83.TB37, "/") <> 0 Or InStr(1, Frm83.TB37, "\") <> 0 Or InStr(1, Frm83.TB37, "'") <> 0 Then

        x = x + 1
        Data_Err(x) = "[Code 2] mengandungi simbol yang tidak sah."
        
    End If
    
End If
If Frm83.TB8 <> vbNullString Or (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) Then
    
    Frm83_LM_BERAT = Frm83.TB8
    
    If Frm83_LM_BERAT = 0 Then
        x = x + 1
        Data_Err(x) = "Tiada data bagi berat barang. Berat 0 tidak dibenarkan."
    End If
End If
    
If Frm83.CB4 = 1 Then
    If Frm83.TB8 = vbNullString Or (Frm83.TB8 <> vbNullString And Not IsNumeric(Frm83.TB8)) Then
        x = x + 1
        Data_Err(x) = "Sila Masukkan [Berat]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB8 <> vbNullString Or (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) Then
        
        Frm83_LM_BERAT = Frm83.TB8
        
        If Frm83_LM_BERAT = 0 Then
            x = x + 1
            Data_Err(x) = "Tiada data bagi berat barang. Berat 0 tidak dibenarkan."
        End If
    End If
    If Frm83.TB9 = vbNullString Or (Frm83.TB9 <> vbNullString And Not IsNumeric(Frm83.TB9)) Then
        x = x + 1
        Data_Err(x) = "Sila Masukkan [Harga Semasa]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB4 = vbNullString Or (Frm83.TB4 <> vbNullString And Not IsNumeric(Frm83.TB4)) Then
        x = x + 1
        Data_Err(x) = "Sila Masukkan [Upah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB24 = vbNullString Or (Frm83.TB24 <> vbNullString And Not IsNumeric(Frm83.TB24)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Pelanggan]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB25 = vbNullString Or (Frm83.TB25 <> vbNullString And Not IsNumeric(Frm83.TB25)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Ahli]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB26 = vbNullString Or (Frm83.TB26 <> vbNullString And Not IsNumeric(Frm83.TB26)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Silver]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB31 = vbNullString Or (Frm83.TB31 <> vbNullString And Not IsNumeric(Frm83.TB31)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Gold]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB32 = vbNullString Or (Frm83.TB32 <> vbNullString And Not IsNumeric(Frm83.TB32)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Platinum]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB33 = vbNullString Or (Frm83.TB33 <> vbNullString And Not IsNumeric(Frm83.TB33)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Upah Kepada M.Dealer]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
ElseIf Frm83.CB5 = 1 Then
    '+++++++++++ Special Request ++++++++++ Start
    'If Frm83.TB8 = vbNullString Or (Frm83.TB8 <> vbNullString And Not IsNumeric(Frm83.TB8)) Then
    '    x = x + 1
    '    Data_Err(x) = "Sila Masukkan [Berat]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    'End If
    'If Frm83.TB8 <> vbNullString Or (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) Then
        
    '    Frm83_LM_BERAT = Frm83.TB8
        
    '    If Frm83_LM_BERAT = 0 Then
    '        x = x + 1
    '        Data_Err(x) = "Tiada data bagi berat barang. Berat 0 tidak dibenarkan."
    '    End If
    'End If
    '+++++++++++ Special Request ++++++++++ End
    If Frm83.TB24 = vbNullString Or (Frm83.TB24 <> vbNullString And Not IsNumeric(Frm83.TB24)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Pelanggan]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB25 = vbNullString Or (Frm83.TB25 <> vbNullString And Not IsNumeric(Frm83.TB25)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Ahli]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB26 = vbNullString Or (Frm83.TB26 <> vbNullString And Not IsNumeric(Frm83.TB26)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Silver]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB31 = vbNullString Or (Frm83.TB31 <> vbNullString And Not IsNumeric(Frm83.TB31)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Gold]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB32 = vbNullString Or (Frm83.TB32 <> vbNullString And Not IsNumeric(Frm83.TB32)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Platinum]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB33 = vbNullString Or (Frm83.TB33 <> vbNullString And Not IsNumeric(Frm83.TB33)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada M.Dealer]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If

If Frm83.CB8 = 1 Then
    If Frm83.TB19 = vbNullString Or (Frm83.TB19 <> vbNullString And Not IsNumeric(Frm83.TB19)) Then
        x = x + 1
        Data_Err(x) = "Sila Masukkan [Spread]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If
If Frm83.CBB5 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Sila Pilih [Dulang]."
End If
'If Frm83.CB3 = 1 Then
    If Frm83.TB27 = vbNullString Or (Frm83.TB27 <> vbNullString And Not IsNumeric(Frm83.TB27)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat GST"
    End If
'End If
If Frm83.CB4 = 1 Then
    If Frm83.CB14 = 0 And Frm83.CB15 = 1 Then
        If Frm83.TB35 = vbNullString Or (Frm83.TB35 <> vbNullString And Not IsNumeric(Frm83.TB35)) Then
            x = x + 1
            Data_Err(x) = "Sila buat tetapan pengiraan upah dari supplier"
        End If
    End If
End If
If Frm83.CB14 = 1 Then
    If Frm83.TB35 = vbNullString Or (Frm83.TB35 <> vbNullString And Not IsNumeric(Frm83.TB35)) Then
        x = x + 1
        Data_Err(x) = "Sila masukkan [Upah per gram]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If
If Frm83.TB15 <> vbNullString Then

    If InStr(1, Frm83.TB15, "*") <> 0 Or InStr(1, Frm83.TB15, "/") <> 0 Or InStr(1, Frm83.TB15, "\") <> 0 Or InStr(1, Frm83.TB15, "'") <> 0 Then

        x = x + 1
        Data_Err(x) = "[No. Invoice] mengandungi simbol yang tidak sah."
        
    End If
    
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Data_Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah anda ingin masukkan data barang ini ke dalam stok kedai?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
        
        
        
        G_JENIS_URUSAN = 2
        
        '$$$ No. staff $$$ - Start
        If InStr(1, Frm83.CBB6, "  |  ") <> 0 Then
            Frm83_LM_EMP_NO = Split(Frm83.CBB6, "  |  ")(1)
            Frm83_LM_EMP_NAME = Split(Frm83.CBB6, "  |  ")(0)
        Else
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!NoPekerja) Then Frm83_LM_EMP_NO = rs!NoPekerja
    
            End If
            
            rs.Close
            Set rs = Nothing
        End If
        '$$$ No. staff $$$ - End
        
        LM_NOW = Now
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from data_database where id is null", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
            rs.AddNew
            rs!terminal = G_TERMINAL
            rs!write_timestamp = LM_NOW
            If Frm83.L4_Text <> vbNullString Then
                rs!supplier_ID = Frm83.L4_Text 'No. ID Bagi Supplier
            Else
                rs!supplier_ID = Null 'No. ID Bagi Supplier
            End If
            If Frm83.CBB1 <> vbNullString Then
                rs!nama_Supplier = Frm83.CBB1 'Nama Supplier
            Else
                rs!nama_Supplier = Null 'Nama Supplier
            End If
            If Frm83.TB1 <> vbNullString Then
                rs!Kod_Supplier = Frm83.TB1 'Kod Supplier
            Else
                rs!Kod_Supplier = Null 'Kod Supplier
            End If
            If Frm83.L5_Text <> vbNullString Then
                rs!purity_ID = Frm83.L5_Text 'No. ID Bagi Purity
            Else
                rs!purity_ID = Null 'No. ID Bagi Purity
            End If
            If Frm83.CBB2 <> vbNullString Then
                rs!purity = Frm83.CBB2 'Purity
            Else
                rs!purity = Null 'Purity
            End If
            If Frm83.TB2 <> vbNullString Then
                rs!kod_Purity = Frm83.TB2 'Kod Purity
            Else
                rs!kod_Purity = Null 'Kod Purity
            End If
            If Frm83.L6_Text <> vbNullString Then
                rs!kategori_produk_ID = Frm83.L6_Text 'No. ID Bagi Kategori Produk
            Else
                rs!kategori_produk_ID = Null 'No. ID Bagi Kategori Produk
            End If
            If Frm83.CBB3 <> vbNullString Then
                rs!kategori_Produk = Frm83.CBB3 'Kategori Produk
            Else
                rs!kategori_Produk = Null 'Kategori Produk
            End If
            If Frm83.TB3 <> vbNullString Then
                rs!Kod_Kategori_Produk = Frm83.TB3 'Kod Kategori Produk
            Else
                rs!Kod_Kategori_Produk = Null 'Kod Kategori Produk
            End If
            If Frm83.CB12 = 0 Then '0 : Pengiraan GST pada harga belian barang , 1 : Pengiraan GST pada upah sahaja
                rs!gst_barang_atau_upah = 0
            ElseIf Frm83.CB12 = 1 Then
                rs!gst_barang_atau_upah = 1
            End If
            If Frm83.CB4 = 1 Or Frm83.CB10 = 1 Then
                If Frm83.TB8 <> vbNullString Then
                    rs!Berat = Format(Frm83.TB8, "0.00") 'Berat
                    rs!beza_berat = Format(Frm83.TB8, "0.00") 'Baki Berat
                    'rs!berat_asal = Format(Frm83.TB8, "0.00")
                Else
                    rs!Berat = Null 'Berat
                    rs!beza_berat = Null 'Baki Berat
                    'rs!beza_berat = Null
                End If
                If Frm83.TB29 <> vbNullString Then
                    rs!riyal = Format(Frm83.TB29, "0.00") 'Berat Riyal
                Else
                    rs!riyal = Null 'Berat Riyal
                End If
                If Frm83.TB9 <> vbNullString Then
                    rs!kos_Belian_Gram = Format(Frm83.TB9, "0.00") 'Harga Per Gram (Belian)
                Else
                    rs!kos_Belian_Gram = Null 'Harga Per Gram (Belian)
                End If
                If Frm83.TB4 <> vbNullString Then
                    rs!UPAH = Frm83.TB4 'Upah (RM)
                Else
                    rs!UPAH = Null 'Upah (RM)
                End If
                If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) And (Frm83.TB22 <> vbNullString And IsNumeric(Frm83.TB22)) Then
                    Frm83_LM_BERAT = Frm83.TB8 'Berat
                    Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                    Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                    Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                    Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment
    
                    If Frm83.CB12 = 0 Then
                        rs!harga_Per_Gram_Item = Format((Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST) / Frm83_LM_BERAT, "0.00") 'Harga Modal Per Gram [Average] (Harga Emas + Upah) Tanpa Tax
                        rs!harga_per_gram_tanpa_gst = Format(Frm83_LM_HARGA_TOTAL / Frm83_LM_BERAT, "0.00")
                    ElseIf Frm83.CB12 = 1 Then
                        rs!harga_Per_Gram_Item = Format((((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST) / Frm83_LM_BERAT, "0.00")
                        rs!harga_per_gram_tanpa_gst = Format((((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL) / Frm83_LM_BERAT, "0.00")
                    End If
                Else
                    rs!harga_Per_Gram_Item = Null
                End If
                If Frm83.TB24 <> vbNullString Then
                    rs!Upah_Jualan = Format(Frm83.TB24, "0.00") 'Upah Jualan Kepada Pelanggan
                Else
                    rs!Upah_Jualan = Null 'Upah Jualan Kepada Pelanggan
                End If
                If Frm83.TB25 <> vbNullString Then
                    rs!Upah_Member = Format(Frm83.TB25, "0.00") 'Upah Jualan Kepada Ahli / Member
                Else
                    rs!Upah_Member = Null 'Upah Jualan Kepada Ahli / Member
                End If
                If Frm83.TB26 <> vbNullString Then
                    rs!Upah_Pengedar = Format(Frm83.TB26, "0.00") 'Upah Jualan Kepada Pengedar
                Else
                    rs!Upah_Pengedar = Null 'Upah Jualan Kepada Pengedar
                End If
                If Frm83.TB31 <> vbNullString Then
                    rs!Upah_RAF = Format(Frm83.TB31, "0.00") 'Upah Jualan Kepada RAF
                Else
                    rs!Upah_RAF = Null 'Upah Jualan Kepada RAF
                End If
                If Frm83.TB32 <> vbNullString Then
                    rs!upah_normal_dealer = Format(Frm83.TB32, "0.00") 'Upah Jualan Kepada N.Dealer
                Else
                    rs!upah_normal_dealer = Null 'Upah Jualan Kepada N.Dealer
                End If
                If Frm83.TB33 <> vbNullString Then
                    rs!upah_master_dealer = Format(Frm83.TB33, "0.00") 'Upah Jualan Kepada M.Dealer
                Else
                    rs!upah_master_dealer = Null 'Upah Jualan Kepada M.Dealer
                End If
                rs!code_Supplier = Null 'Harga Jualan Kepada Pelanggan
                rs!HargaJualan_Member = Null 'Harga Jualan Kepada Ahli / Member
                rs!HargaJualan_Pengedar = Null 'Harga Jualan Kepada Pengedar
                rs!HargaJualan_RAF = Null 'Harga Jualan Kepada RAF
                rs!hargajualan_normal_dealer = Null 'Harga Jualan Kepada N.Dealer
                rs!hargajualan_master_dealer = Null 'Harga Jualan Kepada M.Dealer
            Else
                '+++++++++++ Special Request ++++++++++ Start
                'rs!Berat = Null 'Berat
                'rs!beza_berat = Null 'Baki Berat
                'If Frm83.TB8 <> vbNullString Then
                '    rs!Berat = Format(Frm83.TB8, "0.00") 'Berat
                '    rs!beza_berat = Format(Frm83.TB8, "0.00") 'Baki Berat
                'Else
                    rs!Berat = Null 'Berat
                    rs!beza_berat = Null 'Baki Berat
                'End If
                '+++++++++++ Special Request ++++++++++ End
                rs!kos_Belian_Gram = Null 'Harga Per Gram (Belian)
                rs!UPAH = Null 'Upah (RM)
                rs!harga_Per_Gram_Item = Null 'Harga Modal Per Gram [Average] (Harga Emas + Upah) Tanpa Tax
                rs!Upah_Jualan = Null 'Upah Jualan Kepada Pelanggan
                rs!Upah_Member = Null 'Upah Jualan Kepada Ahli / Member
                rs!Upah_Pengedar = Null 'Upah Jualan Kepada Pengedar
                rs!Upah_RAF = Null 'Upah Jualan Kepada RAF
                rs!upah_normal_dealer = Null 'Upah Jualan Kepada N.Dealer
                rs!upah_master_dealer = Null 'Upah Jualan Kepada M.Dealer
            End If
            If Frm83.CB5 = 1 Then
                If Frm83.TB24 <> vbNullString Then
                    rs!code_Supplier = Format(Frm83.TB24, "0.00") 'Harga Jualan Kepada Pelanggan
                Else
                    rs!code_Supplier = Null 'Harga Jualan Kepada Pelanggan
                End If
                If Frm83.TB25 <> vbNullString Then
                    rs!HargaJualan_Member = Format(Frm83.TB25, "0.00") 'Harga Jualan Kepada Ahli / Member
                Else
                    rs!HargaJualan_Member = Null 'Harga Jualan Kepada Ahli / Member
                End If
                If Frm83.TB26 <> vbNullString Then
                    rs!HargaJualan_Pengedar = Format(Frm83.TB26, "0.00") 'Harga Jualan Kepada Pengedar
                Else
                    rs!HargaJualan_Pengedar = Null 'Harga Jualan Kepada Pengedar
                End If
                If Frm83.TB31 <> vbNullString Then
                    rs!HargaJualan_RAF = Format(Frm83.TB31, "0.00") 'Harga Jualan Kepada RAF
                Else
                    rs!HargaJualan_RAF = Null 'Harga Jualan Kepada RAF
                End If
                If Frm83.TB32 <> vbNullString Then
                    rs!hargajualan_normal_dealer = Format(Frm83.TB32, "0.00") 'Harga Jualan Kepada N.Dealer
                Else
                    rs!hargajualan_normal_dealer = Null 'Harga Jualan Kepada N.Dealer
                End If
                If Frm83.TB33 <> vbNullString Then
                    rs!hargajualan_master_dealer = Format(Frm83.TB33, "0.00") 'Harga Jualan Kepada M.Dealer
                Else
                    rs!hargajualan_master_dealer = Null 'Harga Jualan Kepada M.Dealer
                End If
                
                '+++++++++++ Special Request ++++++++++ Start
                'rs!Berat = Null 'Berat
                'rs!beza_berat = Null 'Baki Berat
                'If Frm83.TB8 <> vbNullString Then
                '    rs!Berat = Format(Frm83.TB8, "0.00") 'Berat
                '    rs!beza_berat = Format(Frm83.TB8, "0.00") 'Baki Berat
                'Else
                    rs!Berat = Null 'Berat
                    rs!beza_berat = Null 'Baki Berat
                'End If
                '+++++++++++ Special Request ++++++++++ End
                rs!kos_Belian_Gram = Null 'Harga Per Gram (Belian)
                rs!UPAH = Null 'Upah (RM)
                rs!harga_Per_Gram_Item = Null 'Harga Modal Per Gram [Average] (Harga Emas + Upah) Tanpa Tax
                rs!Upah_Jualan = Null 'Upah Jualan Kepada Pelanggan
                rs!Upah_Member = Null 'Upah Jualan Kepada Ahli / Member
                rs!Upah_Pengedar = Null 'Upah Jualan Kepada Pengedar
                rs!Upah_RAF = Null 'Upah Jualan Kepada RAF
                rs!upah_normal_dealer = Null 'Upah Jualan Kepada N.Dealer
                rs!upah_master_dealer = Null 'Upah Jualan Kepada M.Dealer
            Else
                rs!code_Supplier = Null 'Harga Jualan Kepada Pelanggan
                rs!HargaJualan_Member = Null 'Harga Jualan Kepada Ahli / Member
                rs!HargaJualan_Pengedar = Null 'Harga Jualan Kepada Pengedar
                rs!HargaJualan_RAF = Null 'Harga Jualan Kepada RAF
                rs!hargajualan_normal_dealer = Null 'Harga Jualan Kepada N.Dealer
                rs!hargajualan_master_dealer = Null 'Harga Jualan Kepada M.Dealer
            End If
            If Frm83.CB12 = 0 Then 'GST pada harga barang
            
                If Frm83.TB10 <> vbNullString Then
                    rs!kos_Belian_Item = Format(Frm83.TB10, "0.00") 'Harga Asal (RM)
                Else
                    rs!kos_Belian_Item = Null 'Harga Asal (RM)
                End If
                
            End If
            If Frm83.CB12 = 1 Then 'GST pada upah
            
                If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) Then
                    
                    Frm83_LM_BERAT = Frm83.TB8 'Berat
                    Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                    Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                    Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                    Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment
                    
                    rs!kos_Belian_Item = Format((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) + Frm83_LM_HARGA_TOTAL, "0.00") 'Harga Asal (RM)
                    
                End If
            
            End If
            If Frm83.CB8 = 1 Then
                If Frm83.TB19 <> vbNullString Then
                    rs!SpreadValue = Format(Frm83.TB19, "0.00") 'Spread (%)
                Else
                    rs!SpreadValue = Null 'Spread (%)
                End If
            ElseIf Frm83.CB7 = 1 Then
                rs!SpreadValue = Null 'Spread (%)
            End If
            If Frm83.TB21 <> vbNullString Then
                rs!harga_lepas_spread = Format(Frm83.TB21, "0.00") 'Harga asal ditolak spread (RM)
            Else
                rs!harga_lepas_spread = Null 'Harga asal ditolak spread (RM)
            End If
            If Frm83.TB22 <> vbNullString Then
                rs!adjustment = Format(Frm83.TB22, "0.00") 'Adjustment (RM)
            Else
                rs!adjustment = Null 'Adjustment (RM)
            End If
            If Frm83.CB12 = 0 Then 'GST pada harga barang
                If Frm83.TB20 <> vbNullString Then
                    rs!kos_item_tanpa_tax = Format(Frm83.TB20, "0.00") 'Harga Barang + Upah Tanpa Tax
                Else
                    rs!kos_item_tanpa_tax = Null 'Harga Barang + Upah Tanpa Tax
                End If
            End If
            If Frm83.CB12 = 1 Then 'GST pada upah
            
                If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) And (Frm83.TB22 <> vbNullString And IsNumeric(Frm83.TB22)) Then
                    Frm83_LM_BERAT = Frm83.TB8 'Berat
                    Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                    Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                    Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                    Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment
                    
                    rs!kos_item_tanpa_tax = Format(((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL, "0.00")  'Harga Barang + Upah Tanpa Tax
                    
                End If
                
            End If
            If Frm83.TB12 <> vbNullString Then
                rs!dimension_Panjang = Frm83.TB12 'Panjang
            Else
                rs!dimension_Panjang = Null 'Panjang
            End If
            If Frm83.TB13 <> vbNullString Then
                rs!dimension_Lebar = Frm83.TB13 'Lebar
            Else
                rs!dimension_Lebar = Null 'Lebar
            End If
            If Frm83.TB14 <> vbNullString Then
                rs!dimension_Saiz = Frm83.TB14 'Saiz
            Else
                rs!dimension_Saiz = Null 'Saiz
            End If
            If Frm83.TB36 <> vbNullString Then 'Code 1
                rs!code1 = UCase(Frm83.TB36)
            Else
                rs!code1 = Null
            End If
            If Frm83.TB37 <> vbNullString Then 'Code 2
                rs!code2 = UCase(Frm83.TB37)
            Else
                rs!code2 = Null
            End If
            If Frm83.CBB5 <> vbNullString Then
                rs!dulang = Frm83.CBB5 'Dulang
            Else
                rs!dulang = Null 'Dulang
            End If
            If Frm83.TB16 <> vbNullString Then
                rs!remarks = UCase(Frm83.TB16) 'Remarks
            Else
                rs!remarks = Null 'Remarks
            End If
            If Frm83.TB34 <> vbNullString Then
                rs!no_cert = UCase(Frm83.TB34) 'No. Cert
            Else
                rs!no_cert = Null 'No. Cert
            End If
            
            If Frm83.CB2 = 1 Then
            
                rs!gst_ari_nashi = 0 'Status Cukai GST : 0 : ZR(L) , 1 : SR
                rs!kadar_gst = Null 'Kadar GST (%)
                rs!jumlah_gst = Null 'Jumlah Cukai GST (RM)
                rs!gst_included = Null '0 : Harga tidak termasuk GST , 1 : Harga termasuk GST
                
            ElseIf Frm83.CB3 = 1 Then
            
                rs!gst_included = 0 '0 : Harga tidak termasuk GST , 1 : Harga termasuk GST
                rs!gst_ari_nashi = 1 'Status Cukai GST : 0 : ZR(L) , 1 : SR
                rs!kadar_gst = Format(Frm83.L8_Text, "0.00") 'Kadar GST (%)
                rs!jumlah_gst = Format(Frm83.TB27, "0.00") 'Jumlah Cukai GST (RM)
              
            ElseIf Frm83.CB11 = 1 Then
    
                rs!gst_included = 1 '0 : Harga tidak termasuk GST , 1 : Harga termasuk GST
                rs!gst_ari_nashi = 1 'Status Cukai GST : 0 : ZR(L) , 1 : SR
                rs!kadar_gst = Format(Frm83.L8_Text, "0.00") 'Kadar GST (%)
                rs!jumlah_gst = Format(Frm83.TB27, "0.00") 'Jumlah Cukai GST (RM)
                
            End If
            
            If Frm83.L40_Text <> vbNullString Then
                rs!harga_tanpa_gst = Format(Frm83.L40_Text, "0.00") 'Harga Barang Tanpa Tax (kalau gst included)
            Else
                rs!harga_tanpa_gst = Null 'Harga Barang Tanpa Tax (kalau gst included)
            End If
            If Frm83.CB5 = 1 Then 'Barang Permata
            
                If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) Then
                    Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                    Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                    
                    rs!harga_item = Format(Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST, "0.00") 'Jumlah Harga Modal Keseluruhan @ Harga Barang + Upah + GST (RM)
                End If
                
            End If

            If Frm83.CB4 = 1 Or Frm83.CB10 = 1 Then 'Barang Kemas / Gold Bar
            
                If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) And (Frm83.TB22 <> vbNullString And IsNumeric(Frm83.TB22)) Then
                    Frm83_LM_BERAT = Frm83.TB8 'Berat
                    Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                    Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                    Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                    Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment
                    
                    If Frm83.CB12 = 0 Then
                        rs!harga_item = Format(Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST, "0.00") 'Jumlah Harga Modal Keseluruhan @ Harga Barang + Upah + GST (RM)
                    ElseIf Frm83.CB12 = 1 Then
                        rs!harga_item = Format(((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST, "0.00") 'Jumlah Harga Modal Keseluruhan @ Harga Barang + Upah + GST (RM)
                    End If
                    
                End If
                
            End If
        
'### Kod Bagi Status ###
'==========================================
'0 : Tidak Aktif (Tidak Perlu Buat Apa-Apa)
'10 : Aktif - Kemasukkan Data Baru
'2 : Barang Sudah Terjual ****Tidak Dibenarkan Untuk Diedit Atau Dipadamkan
'3 : Kemasukkan Data Baru
'4 : Data Diedit
'5 : Keluarkan Data Dari Database (Bagi Item Asal Yang Dipadamkan)
'6 : Ignore Kemasukkan Data Ke Dalam Database

            rs!StatusItem = 10
        
'### Jenis ###
'0 : BK
'1 : Barang permata
'2 : Emas terpakai BK
'3 : Emas terpakai permata
'4 : gold Bar
'5 : Emas terpakai gold bar
'6 : Trade In BK
'7 : Trade In Barang Permata
'8 : Trade In Gold Bar

'=========================================================
'Frm83.L41_Text
'0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
'=========================================================

        'If Frm83.L41_Text = 1 Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
        
            'If Frm83.CB8 = 1 Then 'Buyback / Trade in
        '        If Frm83.CB4 = 1 Then 'Barang kemas
        '            rs!jenis = 6
        '        ElseIf Frm83.CB5 = 1 Then 'Barang permata
        '            rs!jenis = 7
        '        End If
        '        If Frm83.CB10 = 1 Then 'Gold bar
        '            rs!jenis = 8
        '        End If
            'End If
        
        'ElseIf Frm83.L41_Text = 0 Or Frm83.L41_Text = 2 Then
        
            If Frm83.CB7 = 1 Then 'Penerimaan stok baru
                If Frm83.CB4 = 1 Then 'Barang kemas
                    rs!receiving_Status = 0
                ElseIf Frm83.CB5 = 1 Then 'Barang permata
                    rs!receiving_Status = 1
                End If
                If Frm83.CB10 = 1 Then 'Gold bar
                    rs!receiving_Status = 4
                End If
            ElseIf Frm83.CB8 = 1 Then 'Buyback / Trade in
                If Frm83.CB4 = 1 Then 'Barang kemas
                    rs!receiving_Status = 2
                ElseIf Frm83.CB5 = 1 Then 'Barang permata
                    rs!receiving_Status = 3
                End If
                If Frm83.CB10 = 1 Then 'Gold bar
                    rs!receiving_Status = 5
                End If
            End If
        
            If Frm83.L41_Text = 0 Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
                
                rs!jenis_trade_in = 0 '0 : Belian emas dari pelanggan , 1 : Belian emas secara trade in
                
            ElseIf Frm83.L41_Text = 1 Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
            
                rs!jenis_trade_in = 1 '0 : Belian emas dari pelanggan , 1 : Belian emas secara trade in
                
            End If
            
            If Frm83.TB35 <> vbNullString Then
                rs!upah_per_gram = Format(Frm83.TB35, "0.00")
            Else
                rs!upah_per_gram = "0.00"
            End If
            If Frm83.CB14 = 1 Then
                rs!flag_upah = 0
            ElseIf Frm83.CB15 = 1 Then
                rs!flag_upah = 1
                
                If IsNumeric(Frm83.TB8) And Frm83.TB8 <> 0 Then
                    Frm83_LM_BERAT = Frm83.TB8 'Berat
                End If
                
                Frm83_LM_UPAH = Frm83.TB4 'Upah
                
                rs!upah_per_gram = Format(Frm83_LM_UPAH / Frm83_LM_BERAT, "0.00")
            End If
        
            If Frm83.TB28 <> vbNullString Then
                rs!no_id_gst = UCase(Frm83.TB28)
            Else
                rs!no_id_gst = Null
            End If
            If Frm83.TB15 <> vbNullString Then
                rs!bill_No_Belian = UCase(Frm83.TB15)
            Else
                rs!bill_No_Belian = Null
            End If
            rs!tarikh_belian = Frm83.DTPicker1
            rs!flag_image = 0
            rs!cawangan = G_CAWANGAN
            rs!no_pekerja = Frm83_LM_EMP_NO
            rs!nama_pekerja = Frm83_LM_EMP_NAMA
            rs!susut_berat = "0.00"
        
        End If
        
        rs.Update
        DATA_SAVE = 1
        
        rs.Close
        Set rs = Nothing
        
        If DATA_SAVE = 1 Then
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from data_database where terminal='" & G_TERMINAL & "' AND write_timestamp='" & LM_NOW & "' AND tarikh_belian='" & Frm83.DTPicker1 & "' order by ID DESC", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
            
                If Not IsNull(rs!ID) And Not IsNull(rs!Kod_Kategori_Produk) Then
                    rs!no_siri_Produk = G_KOD_KEDAI & "-" & rs!Kod_Kategori_Produk & Format(rs!ID, "000000")
                    LM_NO_SIRI = G_KOD_KEDAI & "-" & rs!Kod_Kategori_Produk & Format(rs!ID, "000000")
                Else
                    rs!no_siri_Produk = G_KOD_KEDAI & "-" & Format(rs!ID, "000000")
                    LM_NO_SIRI = G_KOD_KEDAI & "-" & Format(rs!ID, "000000")
                End If
                If Not IsNull(rs!ID) Then
                    rs!Barcode = Format(rs!ID, "000000")
                Else
                    rs!Barcode = Format(rs!ID, "000000")
                End If
                rs!nama_pekerja = Frm83_LM_EMP_NAMA
                'rs!cawangan = "HQ"
                rs.Update
                
            End If
            
            rs.Close
            Set rs = Nothing
            
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & Frm83_LM_EMP_NAME & "] Penerimaan stok baru [" & LM_NO_SIRI & "]."
            LogDate_Memory = LM_NOW
            Call UpdateLog_Database
            
            If Frm83.CB4 = 1 Then Frm83.TB8 = "0.00"
            
'### Print Barcode ### - Start
            If Frm83.CB13 = 1 Then
                GM_No_RUJUKAN_BELIAN = LM_NO_SIRI
                G_FIELD = "no_siri_produk"
                Call Print_All_Barcode2
            End If
'### Print Barcode ### - End
            
            MsgBox "Data stok telah berjaya disimpan." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "No. siri produk adalah " & LM_NO_SIRI, vbInformation, "Info"

        End If
        
    End If
End If
End Sub

Private Sub CMD21_Click()
'On Error Resume Next
If Frm83.L10_Text <> 0 Then
    
    If MDI_frm1.L5_Text <> 4 Then
    
        Note = "Adakah mempunyai data yang belum disimpan." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Adakah anda ingin keluar dari menu ini?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
        
            If Frm83.L41_Text = "1" Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
                
                Frm84.Show
                Frm83.Hide
                
            ElseIf Frm83.L41_Text = "0" Or Frm83.L41_Text = "2" Then
            
                'Frm16.Show
                Unload Frm26
                Unload Frm27
                Unload Frm28
                Unload Frm83
                MDI_frm1.L5_Text = 0
                
            End If
            
        End If
        
    Else
    
        Frm84.Show
        Frm83.Hide
                
    End If
    
Else

    If Frm83.L41_Text = "1" Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
        
        Frm84.Show
        Frm83.Hide
        
    ElseIf Frm83.L41_Text = "0" Or Frm83.L41_Text = "2" Then
    
        'Frm16.Show
        Unload Frm26
        Unload Frm27
        Unload Frm28
        Unload Frm83
        MDI_frm1.L5_Text = 0
        
    End If

End If
End Sub

Private Sub CMD22_Click()
'On Error Resume Next
Dim rs2 As ADODB.Recordset

Dim Data_Err(35)
Dim Frm83_LM_BERAT As Double
Dim Frm83_LM_HARGA_TOTAL As Double
Dim Frm83_LM_CUKAI_GST As Double
Dim Frm83_LM_HARGA_SEMASA As Double 'Harga Semasa
Dim Frm83_LM_ADJUSTMENT As Double 'Adjustment
Dim Frm83_LM_UPAH As Double 'Upah

Frm83_LM_HARGA_SEMASA = 0 'Harga Semasa
Frm83_LM_ADJUSTMENT = 0 'Adjustment
Frm83_LM_CUKAI_GST = 0
Frm83_LM_HARGA_TOTAL = 0
Frm83_LM_BERAT = 1
Frm83_LM_UPAH = 0 'Upah

x = 0
DATA_SAVE = 0

If Frm83.CBB6 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Sila Pilih [Nama Pekerja]."
End If
If Frm83.CBB1 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Sila Pilih [Supplier]."
End If
If Frm83.CBB2 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Sila Pilih [Purity]."
End If
If Frm83.CBB3 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Sila Pilih [Kategori Produk]."
End If
If Frm83.TB1 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Tiada Maklumat [Kod Supplier]."
End If
If Frm83.TB2 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Tiada Maklumat [Kod Purity]."
End If
If Frm83.TB3 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Tiada Maklumat [Kod Kategori Produk]."
End If
'If Frm83.TB6 = vbNullString Or Frm83.TB7 = vbNullString Then
'    x = x + 1
'    Data_Err(x) = "Maklumat [No. Siri Produk] Yang Tidak Lengkap."
'End If
If Frm83.L13_Text = vbNullString Then
    x = x + 1
    Data_Err(x) = "Telah berlaku ralat pada data item ini. Sila keluar dari menu ini dan cuba lagi."
End If
If Frm83.CB9 = 1 And Frm83.CB4 = 0 And Frm83.CB5 = 0 Then
    x = x + 1
    Data_Err(x) = "Sila Buat Pilihan Penerimaan [Barang Kemas] Atau [Barang Permata]."
End If
If Frm83.TB10 = vbNullString Or (Frm83.TB10 <> vbNullString And Not IsNumeric(Frm83.TB10)) Then
    x = x + 1
    Data_Err(x) = "Sila Masukkan [Harga Asal]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB19 = vbNullString Or (Frm83.TB19 <> vbNullString And Not IsNumeric(Frm83.TB19)) Then
    x = x + 1
    Data_Err(x) = "Sila Masukkan [Spread (%)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB20 = vbNullString Or (Frm83.TB20 <> vbNullString And Not IsNumeric(Frm83.TB20)) Then
    x = x + 1
    Data_Err(x) = "Sila Masukkan [Harga Belian]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB21 = vbNullString Or (Frm83.TB21 <> vbNullString And Not IsNumeric(Frm83.TB21)) Then
    x = x + 1
    Data_Err(x) = "Sila Masukkan [Harga Asal-Spread]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB22 = vbNullString Or (Frm83.TB22 <> vbNullString And Not IsNumeric(Frm83.TB22)) Then
    x = x + 1
    Data_Err(x) = "Sila Masukkan [Adjusment]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.CB2 = 0 And Frm83.CB3 = 0 And Frm83.CB11 = 0 Then
    x = x + 1
    Data_Err(x) = "Sila buat pilihan jenis cukai GST"
End If
If Frm83.TB36 <> vbNullString Then

    If InStr(1, Frm83.TB36, "*") <> 0 Or InStr(1, Frm83.TB36, "/") <> 0 Or InStr(1, Frm83.TB36, "\") <> 0 Or InStr(1, Frm83.TB36, "'") <> 0 Then

        x = x + 1
        Data_Err(x) = "[Code 1] mengandungi simbol yang tidak sah."
        
    End If
    
End If
If Frm83.TB37 <> vbNullString Then

    If InStr(1, Frm83.TB37, "*") <> 0 Or InStr(1, Frm83.TB37, "/") <> 0 Or InStr(1, Frm83.TB37, "\") <> 0 Or InStr(1, Frm83.TB37, "'") <> 0 Then

        x = x + 1
        Data_Err(x) = "[Code 2] mengandungi simbol yang tidak sah."
        
    End If
    
End If
If Frm83.CB4 = 1 Then
    If Frm83.TB8 = vbNullString Or (Frm83.TB8 <> vbNullString And Not IsNumeric(Frm83.TB8)) Then
        x = x + 1
        Data_Err(x) = "Sila Masukkan [Berat]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB8 <> vbNullString Or (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) Then
        
        Frm83_LM_BERAT = Frm83.TB8
        
        If Frm83_LM_BERAT = 0 Then
            x = x + 1
            Data_Err(x) = "Tiada data bagi berat barang. Berat 0 tidak dibenarkan."
        End If
    End If
    If Frm83.TB9 = vbNullString Or (Frm83.TB9 <> vbNullString And Not IsNumeric(Frm83.TB9)) Then
        x = x + 1
        Data_Err(x) = "Sila Masukkan [Harga Semasa]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB4 = vbNullString Or (Frm83.TB4 <> vbNullString And Not IsNumeric(Frm83.TB4)) Then
        x = x + 1
        Data_Err(x) = "Sila Masukkan [Upah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB24 = vbNullString Or (Frm83.TB24 <> vbNullString And Not IsNumeric(Frm83.TB24)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Pelanggan]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB25 = vbNullString Or (Frm83.TB25 <> vbNullString And Not IsNumeric(Frm83.TB25)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Ahli]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB26 = vbNullString Or (Frm83.TB26 <> vbNullString And Not IsNumeric(Frm83.TB26)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Silver]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB31 = vbNullString Or (Frm83.TB31 <> vbNullString And Not IsNumeric(Frm83.TB31)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Gold]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB32 = vbNullString Or (Frm83.TB32 <> vbNullString And Not IsNumeric(Frm83.TB32)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Platinum]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB33 = vbNullString Or (Frm83.TB33 <> vbNullString And Not IsNumeric(Frm83.TB33)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Upah Kepada M.Dealer]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
ElseIf Frm83.CB5 = 1 Then
    '+++++++++++ Special Request ++++++++++ Start
    'If Frm83.TB8 = vbNullString Or (Frm83.TB8 <> vbNullString And Not IsNumeric(Frm83.TB8)) Then
    '    x = x + 1
    '    Data_Err(x) = "Sila Masukkan [Berat]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    'End If
    'If Frm83.TB8 <> vbNullString Or (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) Then
        
    '    Frm83_LM_BERAT = Frm83.TB8
        
    '    If Frm83_LM_BERAT = 0 Then
    '        x = x + 1
    '        Data_Err(x) = "Tiada data bagi berat barang. Berat 0 tidak dibenarkan."
    '    End If
    'End If
    '+++++++++++ Special Request ++++++++++ End
    If Frm83.TB24 = vbNullString Or (Frm83.TB24 <> vbNullString And Not IsNumeric(Frm83.TB24)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Pelanggan]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB25 = vbNullString Or (Frm83.TB25 <> vbNullString And Not IsNumeric(Frm83.TB25)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Ahli]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB26 = vbNullString Or (Frm83.TB26 <> vbNullString And Not IsNumeric(Frm83.TB26)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Silver]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB31 = vbNullString Or (Frm83.TB31 <> vbNullString And Not IsNumeric(Frm83.TB31)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Gold]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB32 = vbNullString Or (Frm83.TB32 <> vbNullString And Not IsNumeric(Frm83.TB32)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Platinum]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB33 = vbNullString Or (Frm83.TB33 <> vbNullString And Not IsNumeric(Frm83.TB33)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada M.Dealer]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If

If Frm83.CB8 = 1 Then
    If Frm83.TB19 = vbNullString Or (Frm83.TB19 <> vbNullString And Not IsNumeric(Frm83.TB19)) Then
        x = x + 1
        Data_Err(x) = "Sila Masukkan [Spread]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If
If Frm83.CBB5 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Sila Pilih [Dulang]."
End If
'If Frm83.CB3 = 1 Then
    If Frm83.TB27 = vbNullString Or (Frm83.TB27 <> vbNullString And Not IsNumeric(Frm83.TB27)) Then
        x = x + 1
        Data_Err(x) = "Tiada Maklumat GST"
    End If
'End If
If Frm83.CB4 = 1 Then
    If Frm83.CB14 = 0 And Frm83.CB15 = 1 Then
        If Frm83.TB35 = vbNullString Or (Frm83.TB35 <> vbNullString And Not IsNumeric(Frm83.TB35)) Then
            x = x + 1
            Data_Err(x) = "Sila buat tetapan pengiraan upah dari supplier"
        End If
    End If
End If
If Frm83.CB14 = 1 Then
    If Frm83.TB35 = vbNullString Or (Frm83.TB35 <> vbNullString And Not IsNumeric(Frm83.TB35)) Then
        x = x + 1
        Data_Err(x) = "Sila masukkan [Upah per gram]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If
If Frm83.TB15 <> vbNullString Then

    If InStr(1, Frm83.TB15, "*") <> 0 Or InStr(1, Frm83.TB15, "/") <> 0 Or InStr(1, Frm83.TB15, "\") <> 0 Or InStr(1, Frm83.TB15, "'") <> 0 Then

        x = x + 1
        Data_Err(x) = "[No. Invoice] mengandungi simbol yang tidak sah."
        
    End If
    
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Data_Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah anda ingin simpan data yang telah diedit ini?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
        
        G_JENIS_URUSAN = 2
        
        '$$$ No. staff $$$ - Start
        If InStr(1, Frm83.CBB6, "  |  ") <> 0 Then
        
            Frm83_LM_EMP_NO = Split(Frm83.CBB6, "  |  ")(1)
            Frm83_LM_EMP_NAME = Split(Frm83.CBB6, "  |  ")(0)
            
        Else
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!NoPekerja) Then Frm83_LM_EMP_NO = rs!NoPekerja
    
            End If
            
            rs.Close
            Set rs = Nothing
        
        End If
        '$$$ No. staff $$$ - End
        
        LM_NOW = Now
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from data_database where ID='" & Frm83.L13_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!StatusItem) Then
                
                If rs!StatusItem = "10" Then
                
                    G_ID = rs!ID
                    Call recovery_data_database
                    
                    rs!terminal = G_TERMINAL
                    rs!write_timestamp2 = LM_NOW
                    If Frm83.L4_Text <> vbNullString Then
                        rs!supplier_ID = Frm83.L4_Text 'No. ID Bagi Supplier
                    Else
                        rs!supplier_ID = Null 'No. ID Bagi Supplier
                    End If
                    If Frm83.CBB1 <> vbNullString Then
                        rs!nama_Supplier = Frm83.CBB1 'Nama Supplier
                    Else
                        rs!nama_Supplier = Null 'Nama Supplier
                    End If
                    If Frm83.TB1 <> vbNullString Then
                        rs!Kod_Supplier = Frm83.TB1 'Kod Supplier
                    Else
                        rs!Kod_Supplier = Null 'Kod Supplier
                    End If
                    If Frm83.L5_Text <> vbNullString Then
                        rs!purity_ID = Frm83.L5_Text 'No. ID Bagi Purity
                    Else
                        rs!purity_ID = Null 'No. ID Bagi Purity
                    End If
                    If Frm83.CBB2 <> vbNullString Then
                        rs!purity = Frm83.CBB2 'Purity
                    Else
                        rs!purity = Null 'Purity
                    End If
                    If Frm83.TB2 <> vbNullString Then
                        rs!kod_Purity = Frm83.TB2 'Kod Purity
                    Else
                        rs!kod_Purity = Null 'Kod Purity
                    End If
                    If Frm83.L6_Text <> vbNullString Then
                        rs!kategori_produk_ID = Frm83.L6_Text 'No. ID Bagi Kategori Produk
                    Else
                        rs!kategori_produk_ID = Null 'No. ID Bagi Kategori Produk
                    End If
                    If Frm83.CBB3 <> vbNullString Then
                        rs!kategori_Produk = Frm83.CBB3 'Kategori Produk
                    Else
                        rs!kategori_Produk = Null 'Kategori Produk
                    End If
                    If Frm83.TB3 <> vbNullString Then
                        rs!Kod_Kategori_Produk = Frm83.TB3 'Kod Kategori Produk
                    Else
                        rs!Kod_Kategori_Produk = Null 'Kod Kategori Produk
                    End If
                    If Frm83.CB12 = 0 Then '0 : Pengiraan GST pada harga belian barang , 1 : Pengiraan GST pada upah sahaja
                        rs!gst_barang_atau_upah = 0
                    ElseIf Frm83.CB12 = 1 Then
                        rs!gst_barang_atau_upah = 1
                    End If
                    If Frm83.CB4 = 1 Or Frm83.CB10 = 1 Then
                        If Frm83.TB8 <> vbNullString Then
                            rs!Berat = Format(Frm83.TB8, "0.00") 'Berat
                            rs!beza_berat = Format(Frm83.TB8, "0.00") 'Baki Berat
                        Else
                            rs!Berat = Null 'Berat
                            rs!beza_berat = Null 'Baki Berat
                        End If
                        If Frm83.TB29 <> vbNullString Then
                            rs!riyal = Format(Frm83.TB29, "0.00") 'Berat Riyal
                        Else
                            rs!riyal = Null 'Berat Riyal
                        End If
                        If Frm83.TB9 <> vbNullString Then
                            rs!kos_Belian_Gram = Format(Frm83.TB9, "0.00") 'Harga Per Gram (Belian)
                        Else
                            rs!kos_Belian_Gram = Null 'Harga Per Gram (Belian)
                        End If
                        If Frm83.TB4 <> vbNullString Then
                            rs!UPAH = Frm83.TB4 'Upah (RM)
                        Else
                            rs!UPAH = Null 'Upah (RM)
                        End If
                        If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) And (Frm83.TB22 <> vbNullString And IsNumeric(Frm83.TB22)) Then
                            Frm83_LM_BERAT = Frm83.TB8 'Berat
                            Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                            Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                            Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                            Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment
                    
                            If Frm83.CB12 = 0 Then
                                rs!harga_Per_Gram_Item = Format((Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST) / Frm83_LM_BERAT, "0.00") 'Harga Modal Per Gram [Average] (Harga Emas + Upah) Tanpa Tax
                                rs!harga_per_gram_tanpa_gst = Format(Frm83_LM_HARGA_TOTAL / Frm83_LM_BERAT, "0.00")
                            ElseIf Frm83.CB12 = 1 Then
                                rs!harga_Per_Gram_Item = Format((((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST) / Frm83_LM_BERAT, "0.00")
                                rs!harga_per_gram_tanpa_gst = Format((((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL) / Frm83_LM_BERAT, "0.00")
                            End If
                        Else
                            rs!harga_Per_Gram_Item = Null
                        End If
                        If Frm83.TB24 <> vbNullString Then
                            rs!Upah_Jualan = Format(Frm83.TB24, "0.00") 'Upah Jualan Kepada Pelanggan
                        Else
                            rs!Upah_Jualan = Null 'Upah Jualan Kepada Pelanggan
                        End If
                        If Frm83.TB25 <> vbNullString Then
                            rs!Upah_Member = Format(Frm83.TB25, "0.00") 'Upah Jualan Kepada Ahli / Member
                        Else
                            rs!Upah_Member = Null 'Upah Jualan Kepada Ahli / Member
                        End If
                        If Frm83.TB26 <> vbNullString Then
                            rs!Upah_Pengedar = Format(Frm83.TB26, "0.00") 'Upah Jualan Kepada Pengedar
                        Else
                            rs!Upah_Pengedar = Null 'Upah Jualan Kepada Pengedar
                        End If
                        If Frm83.TB31 <> vbNullString Then
                            rs!Upah_RAF = Format(Frm83.TB31, "0.00") 'Upah Jualan Kepada RAF
                        Else
                            rs!Upah_RAF = Null 'Upah Jualan Kepada RAF
                        End If
                        If Frm83.TB32 <> vbNullString Then
                            rs!upah_normal_dealer = Format(Frm83.TB32, "0.00") 'Upah Jualan Kepada N.Dealer
                        Else
                            rs!upah_normal_dealer = Null 'Upah Jualan Kepada N.Dealer
                        End If
                        If Frm83.TB33 <> vbNullString Then
                            rs!upah_master_dealer = Format(Frm83.TB33, "0.00") 'Upah Jualan Kepada M.Dealer
                        Else
                            rs!upah_master_dealer = Null 'Upah Jualan Kepada M.Dealer
                        End If
                        rs!code_Supplier = Null 'Harga Jualan Kepada Pelanggan
                        rs!HargaJualan_Member = Null 'Harga Jualan Kepada Ahli / Member
                        rs!HargaJualan_Pengedar = Null 'Harga Jualan Kepada Pengedar
                        rs!HargaJualan_RAF = Null 'Harga Jualan Kepada RAF
                        rs!hargajualan_normal_dealer = Null 'Harga Jualan Kepada N.Dealer
                        rs!hargajualan_master_dealer = Null 'Harga Jualan Kepada M.Dealer
                    Else
                        '+++++++++++ Special Request ++++++++++ Start
                        'rs!Berat = Null 'Berat
                        'rs!beza_berat = Null 'Baki Berat
                        'If Frm83.TB8 <> vbNullString Then
                        '    rs!Berat = Format(Frm83.TB8, "0.00") 'Berat
                        '    rs!beza_berat = Format(Frm83.TB8, "0.00") 'Baki Berat
                        'Else
                            rs!Berat = Null 'Berat
                            rs!beza_berat = Null 'Baki Berat
                        'End If
                        '+++++++++++ Special Request ++++++++++ End
                        rs!kos_Belian_Gram = Null 'Harga Per Gram (Belian)
                        rs!UPAH = Null 'Upah (RM)
                        rs!harga_Per_Gram_Item = Null 'Harga Modal Per Gram [Average] (Harga Emas + Upah) Tanpa Tax
                        rs!Upah_Jualan = Null 'Upah Jualan Kepada Pelanggan
                        rs!Upah_Member = Null 'Upah Jualan Kepada Ahli / Member
                        rs!Upah_Pengedar = Null 'Upah Jualan Kepada Pengedar
                        rs!Upah_RAF = Null 'Upah Jualan Kepada RAF
                        rs!upah_normal_dealer = Null 'Upah Jualan Kepada N.Dealer
                        rs!upah_master_dealer = Null 'Upah Jualan Kepada M.Dealer
                    End If
                    If Frm83.CB5 = 1 Then
                        If Frm83.TB24 <> vbNullString Then
                            rs!code_Supplier = Format(Frm83.TB24, "0.00") 'Harga Jualan Kepada Pelanggan
                        Else
                            rs!code_Supplier = Null 'Harga Jualan Kepada Pelanggan
                        End If
                        If Frm83.TB25 <> vbNullString Then
                            rs!HargaJualan_Member = Format(Frm83.TB25, "0.00") 'Harga Jualan Kepada Ahli / Member
                        Else
                            rs!HargaJualan_Member = Null 'Harga Jualan Kepada Ahli / Member
                        End If
                        If Frm83.TB26 <> vbNullString Then
                            rs!HargaJualan_Pengedar = Format(Frm83.TB26, "0.00") 'Harga Jualan Kepada Pengedar
                        Else
                            rs!HargaJualan_Pengedar = Null 'Harga Jualan Kepada Pengedar
                        End If
                        If Frm83.TB31 <> vbNullString Then
                            rs!HargaJualan_RAF = Format(Frm83.TB31, "0.00") 'Harga Jualan Kepada RAF
                        Else
                            rs!HargaJualan_RAF = Null 'Harga Jualan Kepada RAF
                        End If
                        If Frm83.TB32 <> vbNullString Then
                            rs!hargajualan_normal_dealer = Format(Frm83.TB32, "0.00") 'Harga Jualan Kepada N.Dealer
                        Else
                            rs!hargajualan_normal_dealer = Null 'Harga Jualan Kepada N.Dealer
                        End If
                        If Frm83.TB33 <> vbNullString Then
                            rs!hargajualan_master_dealer = Format(Frm83.TB33, "0.00") 'Harga Jualan Kepada M.Dealer
                        Else
                            rs!hargajualan_master_dealer = Null 'Harga Jualan Kepada M.Dealer
                        End If
                        
                        '+++++++++++ Special Request ++++++++++ Start
                        'rs!Berat = Null 'Berat
                        'rs!beza_berat = Null 'Baki Berat
                        'If Frm83.TB8 <> vbNullString Then
                        '    rs!Berat = Format(Frm83.TB8, "0.00") 'Berat
                        '    rs!beza_berat = Format(Frm83.TB8, "0.00") 'Baki Berat
                        'Else
                            rs!Berat = Null 'Berat
                            rs!beza_berat = Null 'Baki Berat
                        'End If
                        '+++++++++++ Special Request ++++++++++ End
                        rs!kos_Belian_Gram = Null 'Harga Per Gram (Belian)
                        rs!UPAH = Null 'Upah (RM)
                        rs!harga_Per_Gram_Item = Null 'Harga Modal Per Gram [Average] (Harga Emas + Upah) Tanpa Tax
                        rs!Upah_Jualan = Null 'Upah Jualan Kepada Pelanggan
                        rs!Upah_Member = Null 'Upah Jualan Kepada Ahli / Member
                        rs!Upah_Pengedar = Null 'Upah Jualan Kepada Pengedar
                        rs!Upah_RAF = Null 'Upah Jualan Kepada RAF
                        rs!upah_normal_dealer = Null 'Upah Jualan Kepada N.Dealer
                        rs!upah_master_dealer = Null 'Upah Jualan Kepada M.Dealer
                    Else
                        rs!code_Supplier = Null 'Harga Jualan Kepada Pelanggan
                        rs!HargaJualan_Member = Null 'Harga Jualan Kepada Ahli / Member
                        rs!HargaJualan_Pengedar = Null 'Harga Jualan Kepada Pengedar
                        rs!HargaJualan_RAF = Null 'Harga Jualan Kepada RAF
                        rs!hargajualan_normal_dealer = Null 'Harga Jualan Kepada N.Dealer
                        rs!hargajualan_master_dealer = Null 'Harga Jualan Kepada M.Dealer
                    End If
                    If Frm83.CB12 = 0 Then 'GST pada harga barang
                    
                        If Frm83.TB10 <> vbNullString Then
                            rs!kos_Belian_Item = Format(Frm83.TB10, "0.00") 'Harga Asal (RM)
                        Else
                            rs!kos_Belian_Item = Null 'Harga Asal (RM)
                        End If
                        
                    End If
                    If Frm83.CB12 = 1 Then 'GST pada upah
                    
                        If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) Then
                            
                            Frm83_LM_BERAT = Frm83.TB8 'Berat
                            Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                            Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                            Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                            Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment
                            
                            rs!kos_Belian_Item = Format((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) + Frm83_LM_HARGA_TOTAL, "0.00") 'Harga Asal (RM)
                            
                        End If
                    
                    End If
                    If Frm83.CB8 = 1 Then
                        If Frm83.TB19 <> vbNullString Then
                            rs!SpreadValue = Format(Frm83.TB19, "0.00") 'Spread (%)
                        Else
                            rs!SpreadValue = Null 'Spread (%)
                        End If
                    ElseIf Frm83.CB7 = 1 Then
                        rs!SpreadValue = Null 'Spread (%)
                    End If
                    If Frm83.TB21 <> vbNullString Then
                        rs!harga_lepas_spread = Format(Frm83.TB21, "0.00") 'Harga asal ditolak spread (RM)
                    Else
                        rs!harga_lepas_spread = Null 'Harga asal ditolak spread (RM)
                    End If
                    If Frm83.TB22 <> vbNullString Then
                        rs!adjustment = Format(Frm83.TB22, "0.00") 'Adjustment (RM)
                    Else
                        rs!adjustment = Null 'Adjustment (RM)
                    End If
                    If Frm83.CB12 = 0 Then 'GST pada harga barang
                        If Frm83.TB20 <> vbNullString Then
                            rs!kos_item_tanpa_tax = Format(Frm83.TB20, "0.00") 'Harga Barang + Upah Tanpa Tax
                        Else
                            rs!kos_item_tanpa_tax = Null 'Harga Barang + Upah Tanpa Tax
                        End If
                    End If
                    If Frm83.CB12 = 1 Then 'GST pada upah
                    
                        If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) And (Frm83.TB22 <> vbNullString And IsNumeric(Frm83.TB22)) Then
                            Frm83_LM_BERAT = Frm83.TB8 'Berat
                            Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                            Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                            Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                            Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment
                            
                            rs!kos_item_tanpa_tax = Format(((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL, "0.00")  'Harga Barang + Upah Tanpa Tax
                            
                        End If
                        
                    End If
                    If Frm83.TB12 <> vbNullString Then
                        rs!dimension_Panjang = Frm83.TB12 'Panjang
                    Else
                        rs!dimension_Panjang = Null 'Panjang
                    End If
                    If Frm83.TB13 <> vbNullString Then
                        rs!dimension_Lebar = Frm83.TB13 'Lebar
                    Else
                        rs!dimension_Lebar = Null 'Lebar
                    End If
                    If Frm83.TB14 <> vbNullString Then
                        rs!dimension_Saiz = Frm83.TB14 'Saiz
                    Else
                        rs!dimension_Saiz = Null 'Saiz
                    End If
                    If Frm83.TB36 <> vbNullString Then 'Code 1
                        rs!code1 = UCase(Frm83.TB36)
                    Else
                        rs!code1 = Null
                    End If
                    If Frm83.TB37 <> vbNullString Then 'Code 2
                        rs!code2 = UCase(Frm83.TB37)
                    Else
                        rs!code2 = Null
                    End If
                    If Frm83.CBB5 <> vbNullString Then
                        rs!dulang = Frm83.CBB5 'Dulang
                    Else
                        rs!dulang = Null 'Dulang
                    End If
                    If Frm83.TB16 <> vbNullString Then
                        rs!remarks = UCase(Frm83.TB16) 'Remarks
                    Else
                        rs!remarks = Null 'Remarks
                    End If
                    If Frm83.TB34 <> vbNullString Then
                        rs!no_cert = UCase(Frm83.TB34) 'No. Cert
                    Else
                        rs!no_cert = Null 'No. Cert
                    End If
                    
                    If Frm83.CB2 = 1 Then
                    
                        rs!gst_ari_nashi = 0 'Status Cukai GST : 0 : ZR(L) , 1 : SR
                        rs!kadar_gst = Null 'Kadar GST (%)
                        rs!jumlah_gst = Null 'Jumlah Cukai GST (RM)
                        rs!gst_included = Null '0 : Harga tidak termasuk GST , 1 : Harga termasuk GST
                        
                    ElseIf Frm83.CB3 = 1 Then
                    
                        rs!gst_included = 0 '0 : Harga tidak termasuk GST , 1 : Harga termasuk GST
                        rs!gst_ari_nashi = 1 'Status Cukai GST : 0 : ZR(L) , 1 : SR
                        rs!kadar_gst = Format(Frm83.L8_Text, "0.00") 'Kadar GST (%)
                        rs!jumlah_gst = Format(Frm83.TB27, "0.00") 'Jumlah Cukai GST (RM)
                      
                    ElseIf Frm83.CB11 = 1 Then
                    
                        rs!gst_included = 1 '0 : Harga tidak termasuk GST , 1 : Harga termasuk GST
                        rs!gst_ari_nashi = 1 'Status Cukai GST : 0 : ZR(L) , 1 : SR
                        rs!kadar_gst = Format(Frm83.L8_Text, "0.00") 'Kadar GST (%)
                        rs!jumlah_gst = Format(Frm83.TB27, "0.00") 'Jumlah Cukai GST (RM)
                        
                    End If
                    
                    If Frm83.L40_Text <> vbNullString Then
                        rs!harga_tanpa_gst = Format(Frm83.L40_Text, "0.00") 'Harga Barang Tanpa Tax (kalau gst included)
                    Else
                        rs!harga_tanpa_gst = Null 'Harga Barang Tanpa Tax (kalau gst included)
                    End If
                    If Frm83.CB5 = 1 Then 'Barang Permata
                    
                        If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) Then
                            Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                            Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                            
                            rs!harga_item = Format(Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST, "0.00") 'Jumlah Harga Modal Keseluruhan @ Harga Barang + Upah + GST (RM)
                        End If
                        
                    End If
                    
                    If Frm83.CB4 = 1 Or Frm83.CB10 = 1 Then 'Barang Kemas / Gold Bar
                    
                        If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) And (Frm83.TB22 <> vbNullString And IsNumeric(Frm83.TB22)) Then
                            Frm83_LM_BERAT = Frm83.TB8 'Berat
                            Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                            Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                            Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                            Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment
                            
                            If Frm83.CB12 = 0 Then
                                rs!harga_item = Format(Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST, "0.00") 'Jumlah Harga Modal Keseluruhan @ Harga Barang + Upah + GST (RM)
                            ElseIf Frm83.CB12 = 1 Then
                                rs!harga_item = Format(((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST, "0.00") 'Jumlah Harga Modal Keseluruhan @ Harga Barang + Upah + GST (RM)
                            End If
                            
                        End If
                        
                    End If
                    
                    '### Kod Bagi Status ###
                    '==========================================
                    '0 : Tidak Aktif (Tidak Perlu Buat Apa-Apa)
                    '10 : Aktif - Kemasukkan Data Baru
                    '2 : Barang Sudah Terjual ****Tidak Dibenarkan Untuk Diedit Atau Dipadamkan
                    '3 : Kemasukkan Data Baru
                    '4 : Data Diedit
                    '5 : Keluarkan Data Dari Database (Bagi Item Asal Yang Dipadamkan)
                    '6 : Ignore Kemasukkan Data Ke Dalam Database
                    
                    rs!StatusItem = 10
                    
                    '### Jenis ###
                    '0 : BK
                    '1 : Barang permata
                    '2 : Emas terpakai BK
                    '3 : Emas terpakai permata
                    '4 : gold Bar
                    '5 : Emas terpakai gold bar
                    '6 : Trade In BK
                    '7 : Trade In Barang Permata
                    '8 : Trade In Gold Bar
                    
                    '=========================================================
                    'Frm83.L41_Text
                    '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
                    '=========================================================
                    
                    'If Frm83.L41_Text = 1 Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
                    
                        'If Frm83.CB8 = 1 Then 'Buyback / Trade in
                    '        If Frm83.CB4 = 1 Then 'Barang kemas
                    '            rs!jenis = 6
                    '        ElseIf Frm83.CB5 = 1 Then 'Barang permata
                    '            rs!jenis = 7
                    '        End If
                    '        If Frm83.CB10 = 1 Then 'Gold bar
                    '            rs!jenis = 8
                    '        End If
                        'End If
                    
                    'ElseIf Frm83.L41_Text = 0 Or Frm83.L41_Text = 2 Then
                    
                        If Frm83.CB7 = 1 Then 'Penerimaan stok baru
                            If Frm83.CB4 = 1 Then 'Barang kemas
                                rs!receiving_Status = 0
                            ElseIf Frm83.CB5 = 1 Then 'Barang permata
                                rs!receiving_Status = 1
                            End If
                            If Frm83.CB10 = 1 Then 'Gold bar
                                rs!receiving_Status = 4
                            End If
                        ElseIf Frm83.CB8 = 1 Then 'Buyback / Trade in
                            If Frm83.CB4 = 1 Then 'Barang kemas
                                rs!receiving_Status = 2
                            ElseIf Frm83.CB5 = 1 Then 'Barang permata
                                rs!receiving_Status = 3
                            End If
                            If Frm83.CB10 = 1 Then 'Gold bar
                                rs!receiving_Status = 5
                            End If
                        End If
                    
                    If Frm83.L41_Text = 0 Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
                        
                        rs!jenis_trade_in = 0 '0 : Belian emas dari pelanggan , 1 : Belian emas secara trade in
                        
                    ElseIf Frm83.L41_Text = 1 Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
                    
                        rs!jenis_trade_in = 1 '0 : Belian emas dari pelanggan , 1 : Belian emas secara trade in
                        
                    End If
                    
                    If Frm83.TB35 <> vbNullString Then
                        rs!upah_per_gram = Format(Frm83.TB35, "0.00")
                    Else
                        rs!upah_per_gram = "0.00"
                    End If
                    If Frm83.CB14 = 1 Then
                        rs!flag_upah = 0
                    ElseIf Frm83.CB15 = 1 Then
                        rs!flag_upah = 1
                        
                        If IsNumeric(Frm83.TB8) And Frm83.TB8 <> 0 Then
                            Frm83_LM_BERAT = Frm83.TB8 'Berat
                        End If
                        
                        Frm83_LM_UPAH = Frm83.TB4 'Upah
                        
                        rs!upah_per_gram = Format(Frm83_LM_UPAH / Frm83_LM_BERAT, "0.00")
                    End If
                    
                    If Frm83.TB28 <> vbNullString Then
                        rs!no_id_gst = UCase(Frm83.TB28)
                    Else
                        rs!no_id_gst = Null
                    End If
                    If Frm83.TB15 <> vbNullString Then
                        rs!bill_No_Belian = UCase(Frm83.TB15)
                    Else
                        rs!bill_No_Belian = Null
                    End If
                    rs!tarikh_belian = Frm83.DTPicker1
                    rs!flag_image = 0
                    'rs!Image = Null
                    rs!no_pekerja = Frm83_LM_EMP_NO
                    rs!nama_pekerja = Frm83_LM_EMP_NAME
                    rs!susut_berat = "0.00"
                    rs.Update
                    DATA_SAVE = 1
                
                Else
                
                    MsgBox "Status barang ini telah berubah dan anda tidak dibenarkan untuk edit data barang ini." & vbCrLf & _
                            vbNullString & vbCrLf & _
                            "Sila batalkan urusan edit data ini dan periksa status terbaru barang ini.", vbExclamation, "Info"
                End If
            
            Else
                
                MsgBox "Status barang ini telah berubah dan anda tidak dibenarkan untuk edit data barang ini." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Sila batalkan urusan edit data ini dan periksa status terbaru barang ini.", vbExclamation, "Info"
                
            End If
            
        Else
            
            MsgBox "Tiada rekod item ini dijumpai." & vbCrLf & _
                    "Kemungkinan besar data item ini telah dipadamkan." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila batalkan urusan edit data ini dan periksa status terbaru barang ini.", vbExclamation, "Info"
                            
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_SAVE = 1 Then
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from data_database where ID='" & Frm83.L13_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
            
                If Not IsNull(rs!ID) And Not IsNull(rs!Kod_Kategori_Produk) Then
                    rs!no_siri_Produk = G_KOD_KEDAI & "-" & rs!Kod_Kategori_Produk & Format(rs!ID, "000000")
                    LM_NO_SIRI = G_KOD_KEDAI & "-" & rs!Kod_Kategori_Produk & Format(rs!ID, "000000")
                Else
                    rs!no_siri_Produk = G_KOD_KEDAI & "-" & Format(rs!ID, "000000")
                    LM_NO_SIRI = G_KOD_KEDAI & "-" & Format(rs!ID, "000000")
                End If
                If Not IsNull(rs!ID) Then
                    rs!Barcode = Format(rs!ID, "000000")
                Else
                    rs!Barcode = Format(rs!ID, "000000")
                End If
                rs!nama_pekerja = Frm83_LM_EMP_NAME
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
            
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & Frm83_LM_EMP_NAME & "] Edit data stok [" & LM_NO_SIRI & "]."
            LogDate_Memory = LM_NOW
            Call UpdateLog_Database

            GM_NEXT_PREV = 2
            
            If Frm101.L33_Text = 0 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                If Frm101.CB2 = 1 Then 'Report Belian
                    Call Frm85_Header_Report_Belian
                    Call Frm85_report_belian_page
                End If
                If Frm101.CB4 = 1 Then 'Report Buyback
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
                Call Frm85_Header_Report_buyback_gb
                Call Frm85_carian_buyback_page
            ElseIf Frm101.L33_Text = 4 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                Call Frm85_Header_Report_Belian
                Call Frm85_search_invoice_supplier_page
            ElseIf Frm101.L33_Text = 5 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                Call Frm85_Header_Report_Belian
                Call Frm85_report_belian_barcode
            ElseIf Frm101.L33_Text = 6 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                Call Frm85_Header_Report_Belian
                Call Frm85_report_buyback_barcode
            ElseIf Frm101.L33_Text = 7 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                Call Frm85_Header_Report_belian_gb
                Call Frm85_report_belian_gb_barcode
            ElseIf Frm101.L33_Text = 8 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                Call Frm85_Header_Report_belian_gb
                Call Frm85_report_buyback_gb_barcode
            End If
            
            Frm85.Show
            Unload Frm83

            Unload Frm26
            Unload Frm27
            Unload Frm28
            
            Note = "Data stok yang telah diedit berjaya disimpan." & vbCrLf & _
                    "Mungkin telah berlaku perubahan terhadap maklumat barang ini dan memerlukan untuk cetak kembali barcode barang ini." & vbCrLf & _
                    "Adakah anda ingin cetak barcode barang ini?" & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "No. siri produk adalah " & LM_NO_SIRI & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Teruskan?"
            
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbNo Then
                Exit Sub
            End If
            If Answer = vbYes Then
                
                GM_No_RUJUKAN_BELIAN = LM_NO_SIRI
                G_FIELD = "no_siri_produk"
                Call Print_All_Barcode2
                
            End If

        End If
    End If
End If
End Sub

Private Sub CMD23_Click()
'On Error Resume Next
If Frm83.L10_Text <> 0 Then
    
    If MDI_frm1.L5_Text <> 4 Then
    
        Note = "Adakah mempunyai data yang belum disimpan." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Adakah anda ingin keluar dari menu ini?"
        
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
        
            If Frm83.L41_Text = "1" Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
                
                Frm84.Show
                Frm83.Hide
                
            ElseIf Frm83.L41_Text = "0" Or Frm83.L41_Text = "2" Then
            
                Frm85.Show
                Unload Frm26
                Unload Frm27
                Unload Frm28
                Unload Frm83
                
            End If
        
        End If
        
    Else
    
        Frm84.Show
        Frm83.Hide
    
    End If

Else

    If Frm83.L41_Text = "1" Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
        
        Frm84.Show
        Frm83.Hide
        
    ElseIf Frm83.L41_Text = "0" Or Frm83.L41_Text = "2" Then
    
        Frm85.Show
        Unload Frm26
        Unload Frm27
        Unload Frm28
        Unload Frm83
        
    End If

End If
End Sub


Private Sub CMD24_Click()
'On Error Resume Next
If Frm83.L36_Text = vbNullString Then
    
    If Frm83.L37_Text <> vbNullString Then
    
        Note = "Anda telah mengisi data penjual barang kemas terpakai ini di dalam ruangan pelanggan yang berdaftar dengan kedai." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Jika ada meneruskan menu ini , semua data penjual di dalam ruangan pelanggan berdaftar akan dipadamkan." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
        
            Unload Frm27
            Unload Frm28
            Call Frm26_initial
            
            Frm83.L37_Text = vbNullString 'Nama pembeli : Berdaftar
            
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

Private Sub CMD25_Click()
'On Error Resume Next
If Frm83.L37_Text = vbNullString Then
    
    If Frm83.L36_Text <> vbNullString Then
    
        Note = "Anda telah mengisi data penjual barang kemas terpakai ini di dalam ruangan pelanggan yang TIDAK berdaftar dengan kedai." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Jika ada meneruskan menu ini , semua data penjual di dalam ruangan pelanggan TIDAK berdaftar akan dipadamkan." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
            
            'Unload Frm26
            'Unload Frm27
            Call Frm28_initial
            
            Frm83.L36_Text = vbNullString 'Nama pembeli : Tidak berdaftar
            
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

Private Sub CMD26_Click()
'on error resume next
Dim frm83_LM_CURR_PAGE As Double
Dim frm83_LM_TOTAL_PAGE As Double

frm83_LM_CURR_PAGE = 0
frm83_LM_TOTAL_PAGE = 0

If Frm83.L67_Text <> vbNullString And IsNumeric(Frm83.L67_Text) Then
    If Frm83.L68_Text <> vbNullString And IsNumeric(Frm83.L68_Text) Then
        frm83_LM_CURR_PAGE = Frm83.L67_Text
        frm83_LM_TOTAL_PAGE = Frm83.L68_Text
        
        If frm83_LM_CURR_PAGE <> 1 And frm83_LM_CURR_PAGE <> 0 Then
        
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
                    
            Call Frm83_Senarai_Belian_Header
            Call Frm83_Senarai_Belian
                    
        End If

    End If
End If
End Sub

Private Sub CMD27_Click()
'on error resume next
Dim frm83_LM_CURR_PAGE As Double
Dim frm83_LM_TOTAL_PAGE As Double

frm83_LM_CURR_PAGE = 0
frm83_LM_TOTAL_PAGE = 0

If Frm83.L67_Text <> vbNullString And IsNumeric(Frm83.L67_Text) Then
    If Frm83.L68_Text <> vbNullString And IsNumeric(Frm83.L68_Text) Then
        frm83_LM_CURR_PAGE = Frm83.L67_Text
        frm83_LM_TOTAL_PAGE = Frm83.L68_Text
        
        If frm83_LM_CURR_PAGE < frm83_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm83_Senarai_Belian_Header
            Call Frm83_Senarai_Belian
            
        End If
    End If
End If
End Sub

Private Sub CMD4_Click()
'On Error Resume Next
Frm83.Frame9.Visible = False
End Sub
Private Sub CMD5_Click()
'On Error Resume Next
Dim Data_Err(10)
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim LM_JUM_BAYARAN As Double
Dim LM_TOTAL As Double
x = 0
Y = 0
DATA_SAVE = 0

If Frm83.L10_Text = "0" Then
    x = x + 1
    Data_Err(x) = "Tiada Senarai Belian."
End If
If Frm83.CBB6 = vbNullString Then
    x = x + 1
    Data_Err(x) = "Sila Pilih [Nama Pekerja]."
End If
If Frm83.TB40 = vbNullString Or (Frm83.TB40 <> vbNullString And Not IsNumeric(Frm83.TB40)) Then
    x = x + 1
    Data_Err(x) = "Sila Masukkan [Tunai (RM)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB41 = vbNullString Or (Frm83.TB41 <> vbNullString And Not IsNumeric(Frm83.TB41)) Then
    x = x + 1
    Data_Err(x) = "Sila Masukkan [Bank In (RM)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If

If Frm83.CB8 = 1 Then
'### Periksa Samada Maklumat Penjual Diisi Dalam Kedua-dua Ruangan Berdaftar Dan Tidak Berdaftar ### - Start
    If x = 0 Then
        If Frm83.L36_Text <> vbNullString And Frm83.L37_Text <> vbNullString Then
        
            MsgBox "Data bagi penjual telah diisi bagi kedua-dua ruangan pelanggan berdaftar dan tidak berdaftar." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila padam salah satu yang tidak berkenaan.", vbExclamation, "Info"
                        
            Exit Sub
              
        End If
    End If
'### Periksa Samada Maklumat Penjual Diisi Dalam Kedua-dua Ruangan Berdaftar Dan Tidak Berdaftar ### - End
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Data_Err(Y)
    Next Y
    Exit Sub
Else

    LM_JUM_BAYARAN = 0
    LM_TOTAL = 0
    
    LM_JUM_BAYARAN = Frm83.TB42
    LM_TOTAL = Frm83.L26_Text
    
    If LM_TOTAL <> LM_JUM_BAYARAN Then
        MsgBox "Jumlah voucher TIDAK SAMA dengan jumlah bayaran." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Jumlah Voucher : RM " & Format(LM_TOTAL, "#,##0.00") & vbCrLf & _
                "Jumlah Cara Bayaran : RM " & Format(LM_JUM_BAYARAN, "#,##0.00"), vbexclamtion, "Info"
        Exit Sub
    End If
    
    Note = "Adakah anda yakin untuk teruskan urusan belian ini ?" & vbCrLf & _
            vbNullString & vbCrLf & _
            "Data belian akan disimpan ke dalam sistem."
                        
    If Frm83.CB8 = 1 Then

        If Frm83.L37_Text <> vbNullString And Frm83.L36_Text = vbNullString Then
            Note = "Adakah anda yakin untuk teruskan urusan belian ini ?" & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Data belian akan disimpan ke dalam sistem."
        End If
        
        If Frm83.L37_Text = vbNullString And Frm83.L36_Text <> vbNullString Then
            Note = "Adakah anda yakin untuk teruskan urusan belian ini ?" & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Data belian akan disimpan ke dalam sistem." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "*** Oleh ini pembeli ini tidak berdaftar dengan sistem, oleh itu data dan rekod penjual ini tidak akan disimpan di dalam sistem ***"
        End If

        If Frm83.L37_Text = vbNullString And Frm83.L36_Text = vbNullString Then
        
            Note = "TIADA maklumat bagi penjual telah diisi." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Maklumat penjual tidak akan dicetak di dalam payment voucher." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Adakah anda yakin untuk teruskan urusan belian ini ?"
            
        End If

    End If
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
        
        G_JENIS_URUSAN = 2
        
        '$$$ No. staff $$$ - Start
        If InStr(1, Frm83.CBB6, "  |  ") <> 0 Then
            Frm83_LM_EMP_NO = Split(Frm83.CBB6, "  |  ")(1)
            Frm83_LM_EMP_NAMA = Split(Frm83.CBB6, "  |  ")(0)
        Else
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!NoPekerja) Then Frm83_LM_EMP_NO = rs!NoPekerja
    
            End If
            
            rs.Close
            Set rs = Nothing
        End If
        '$$$ No. staff $$$ - End
        
        LM_NOW = Now
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        strsql = "insert into 14_senarai_voucher(tarikh,terminal,write_timestamp,Status,nama_staff)" & _
                "select '" & Frm83.DTPicker1 & "','" & G_TERMINAL & "','" & LM_NOW & "',1,'" & MDI_frm1.L3_Text & "'"
                                        
        Set rs = cn2.Execute(strsql)
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        rs.Open "select * from 14_senarai_voucher where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & Frm83.DTPicker1 & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!ID) Then
            
                Frm83_LM_No_RUJUKAN_BELIAN = rs!ID 'No. Rujukan Belian
                GM_No_RUJUKAN_BELIAN = Format(rs!ID, "000000") 'No. Rujukan Belian
                Frm83_LM_NO_TI = rs!ID
                rs!no_voucher = "PV" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm83_LM_NO_TI, "000000")
                
            End If
            
            rs.Update
        Else
        
            MsgBox "Berlaku ralat semasa data cuba disimpan. Sila keluar dari menu ini dan cuba lagi.", vbCritical, "Error"
            
            rs.Close
            Set rs = Nothing
            
            Exit Sub
            
        End If
        
        rs.Close
        Set rs = Nothing

Re_Gen_No_Rujukan:
'###Masukkan Data Belian Ke Dalam Database Akaun Belian### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 16_gold_bar_belian where no_rujukan='" & "PV" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm83_LM_NO_TI, "000000") & "'", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
            rs.AddNew
            rs!no_rujukan = "PV" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm83_LM_NO_TI, "000000") 'No. Rujukan Belian
            GM_No_RUJUKAN_BELIAN = "PV" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm83_LM_NO_TI, "000000") 'No. Rujukan Belian
            rs!tarikh = Frm83.DTPicker1 'Tarikh Belian
            rs!cara_bayaran = 0 'Cara Belian // 0 : Cash @ Bank in @ Kad Kredit @ Kad Debit , 1 : Cheque
            If Frm83.TB40 <> vbNullString Then
                rs!tunai = Format(Frm83.TB40, "0.00") 'Cara Bayaran : Tunai (Jumlah Keseluruhan Dengan Cukai GST)
            Else
                rs!tunai = Null 'Cara Bayaran : Tunai (Jumlah Keseluruhan Dengan Cukai GST)
            End If
            If Frm83.TB41 <> vbNullString Then
                rs!bank_in = Format(Frm83.TB41, "0.00") 'Cara Bayaran : Tunai (Jumlah Keseluruhan Dengan Cukai GST)
            Else
                rs!bank_in = Null 'Cara Bayaran : Tunai (Jumlah Keseluruhan Dengan Cukai GST)
            End If
            If Frm83.L11_Text <> vbNullString Then
                rs!jumlah_asal = Format(Frm83.L11_Text, "0.00") 'Jumlah Bayaran Asal (Jumlah Tanpa Cukai GST)
            Else
                rs!jumlah_asal = Null 'Jumlah Bayaran Asal (Jumlah Tanpa Cukai GST)
            End If

            If Frm83.L8_Text <> vbNullString Then
                rs!gst_value = Frm83.L8_Text 'Jumlah Cukai GST (%)
            Else
                rs!gst_value = "0" 'Jumlah Cukai GST (%)
            End If
            If Frm83.L22_Text <> vbNullString Then
                rs!gst_zr_harga = Format(Frm83.L22_Text, "0.00") 'Jumlah Bayaran Yang Dikenakan Cukai GST ZR (RM)
            Else
                rs!gst_zr_harga = "0.00" 'Jumlah Bayaran Yang Dikenakan Cukai GST ZR (RM)
            End If
            If Frm83.L23_Text <> vbNullString Then
                rs!gst_zr_cukai = Format(Frm83.L23_Text, "0.00") 'Jumlah Bayaran Cukai GST ZR (RM)
            Else
                rs!gst_zr_cukai = "0.00" 'Jumlah Bayaran Cukai GST ZR (RM)
            End If
            If Frm83.L24_Text <> vbNullString Then
                rs!gst_sr_harga = Format(Frm83.L24_Text, "0.00") 'Jumlah Bayaran Yang Dikenakan Cukai GST SR (RM)
            Else
                rs!gst_sr_harga = "0.00" 'Jumlah Bayaran Yang Dikenakan Cukai GST SR (RM)
            End If
            If Frm83.L25_Text <> vbNullString Then
                rs!gst_sr_cukai = Format(Frm83.L25_Text, "0.00") 'Jumlah Bayaran Cukai GST SR (RM)
            Else
                rs!gst_sr_cukai = "0.00" 'Jumlah Bayaran Cukai GST SR (RM)
            End If
            If Frm83.TB28 <> vbNullString Then
                rs!no_id_gst_supplier = Frm83.TB28 'No. ID GST Supplier
            Else
                rs!no_id_gst_supplier = Null 'No. ID GST Supplier
            End If
            If Frm83.TB15 <> vbNullString Then
                rs!no_resit_supplier = UCase(Frm83.TB15) 'No. Resit Dari Supplier (Jika Ada)
            Else
                rs!no_resit_supplier = Null 'No. Resit Dari Supplier (Jika Ada)
            End If
            If Frm83.TB1 <> vbNullString Then
                rs!Kod_Supplier = UCase(Frm83.TB1) 'Kod Supplier
            Else
                rs!Kod_Supplier = Null 'Kod Supplier
            End If
            If Frm83.L11_Text <> vbNullString Then
                rs!jumlah_tanpa_gst = Format(Frm83.L11_Text, "0.00") 'Jumlah Bayaran Asal (Jumlah Tanpa Cukai GST)
            Else
                rs!jumlah_tanpa_gst = Null 'Jumlah Bayaran Asal (Jumlah Tanpa Cukai GST)
            End If
            If Frm83.L26_Text <> vbNullString Then
                rs!jumlah_dengan_gst = Format(Frm83.L26_Text, "0.00") 'Jumlah Bayaran Keseluruhan (Jumlah Dengan Cukai GST)
            Else
                rs!jumlah_dengan_gst = Null 'Jumlah Bayaran Keseluruhan (Jumlah Dengan Cukai GST)
            End If
            If Frm83.CB7 = 1 Then
                rs!flag_trade_in = 0 'Flag Trade In // 0 : Tiada , 1 : Ada
            ElseIf Frm83.CB8 = 1 Then
                rs!flag_trade_in = 1 'Flag Trade In // 0 : Tiada , 1 : Ada
                rs!trade_in_status = 0 'Flag Samada Trade In Sudah Digunakan Atau Tidak , 0 : Tiada , 1 : Ada
                rs!no_resit_trade_in = "PV" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm83_LM_NO_TI, "000000") 'No. Resit Trade In
                G_No_RESIT_JUALAN = "PV" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm83_LM_NO_TI, "000000") 'No. Resit Trade In
            End If
            If Frm83.CB8 = 1 Then
                If Frm83.L37_Text <> vbNullString Then
                    If Frm28.L5_Text <> vbNullString Then
                        rs!no_rujukan_pelanggan_buyback = Frm28.L5_Text 'No. Rujukan Pembeli
                    Else
                        rs!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
                    End If
                Else
                    rs!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
                End If
            Else
                rs!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
            End If
            rs!terminal = G_TERMINAL
            rs!no_pekerja = Frm83_LM_EMP_NO 'No. Pekerja
            rs!nama_pekerja = Frm83_LM_EMP_NAMA
            rs!write_timestamp = LM_NOW
            rs!remarks = "Penerimaan stok baru"
            rs!cawangan = G_CAWANGAN
            rs!Status = 1
            rs.Update
        Else
            Frm83_LM_NO_TI = Frm83_LM_NO_TI + 1
            Frm83.L9_Text = Frm83_LM_NO_TI 'No. Rujukan Belian
            
            rs.Close
            Set rs = Nothing
            GoTo Re_Gen_No_Rujukan:
        End If
        
        rs.Close
        Set rs = Nothing
'###Masukkan Data Belian Ke Dalam Database Akaun Belian### - End

'###Masukkan Data Penjual Yang Tidak Berdaftar Ke Dalam Database### - Start
        If Frm83.L36_Text <> vbNullString And Frm83.L37_Text = vbNullString Then

            If Frm26.TB1 <> vbNullString Then 'Nama
                LM_NAMA = UCase(Frm26.TB1)
            Else
                LM_NAMA = Null
            End If
            If Frm26.TB2 <> vbNullString Then 'No. Telefon
                LM_NO_TEL = UCase(Frm26.TB2)
            Else
                LM_NO_TEL = Null
            End If
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            strsql = "insert into 44_senarai_pelanggan(tarikh,no_resit,Nama,no_tel,write_timestamp,no_staff,terminal,jenis_urusan,cawangan)" & _
                    "select '" & Frm83.DTPicker1 & "','" & "PV" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm83_LM_NO_TI, "000000") & "','" & LM_NAMA & "','" & LM_NO_TEL & "','" & LM_NOW & "','" & Frm83_LM_EMP_NO & "','" & G_TERMINAL & "','" & G_JENIS_URUSAN & "','" & G_CAWANGAN & "'"
                                            
            Set rs = cn.Execute(strsql)
            Set rs = Nothing
        End If
'###Masukkan Data Penjual Yang Tidak Berdaftar Ke Dalam Database### - End
            
'### Masukkan maklumat data barang ke dalam table #data_database ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        strsql = "insert into data_database(NoRujukanSistem,no_pekerja,cawangan,nama_pekerja,tarikh_belian,bill_No_Belian,terminal,supplier_ID,nama_Supplier,Kod_Supplier,purity_ID,purity,kod_Purity,kategori_produk_ID,kategori_Produk,Kod_Kategori_Produk,Berat,beza_berat,UPAH,Upah30,riyal,kos_Belian_Gram,kos_Belian_Item,SpreadValue,harga_lepas_spread,adjustment,kos_item_tanpa_tax,cara_Belian,dimension_Panjang,dimension_Lebar,dimension_Saiz,code1,code2,harga_Per_Gram_Item,dulang,no_cert,gst_barang_atau_upah,StatusItem,Upah_Jualan,Upah_Member,Upah_RAF,Upah_Pengedar,code_Supplier,HargaJualan_Member,HargaJualan_Pengedar,upah_normal_dealer,upah_master_dealer,HargaJualan_RAF,hargajualan_normal_dealer,hargajualan_master_dealer,remarks,gst_ari_nashi,kadar_gst,jumlah_gst,harga_item,receiving_Status,harga_tanpa_gst,gst_included,jenis_trade_in,flag_upah,upah_per_gram,flag_image,write_timestamp,no_id_gst,harga_per_gram_tanpa_gst)" & _
                    "select '" & "PV" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm83_LM_NO_TI, "000000") & "','" & Frm83_LM_EMP_NO & "','" & G_CAWANGAN & "','" & Frm83_LM_EMP_NAMA & "',tarikh_belian,bill_No_Belian,terminal,supplier_ID,nama_Supplier,Kod_Supplier,purity_ID,purity,kod_Purity,kategori_produk_ID,kategori_Produk,Kod_Kategori_Produk,Berat,beza_berat,UPAH,Upah30,riyal," _
                    & "kos_Belian_Gram,kos_Belian_Item,Spread,harga_lepas_spread,adjustment,kos_item_tanpa_tax,cara_Belian,dimension_Panjang,dimension_Lebar,dimension_Saiz,code1,code2,harga_Per_Gram_Item,dulang,no_cert,gst_barang_atau_upah,StatusItem,Upah_Jualan,Upah_Member,Upah_RAF,Upah_Pengedar,code_Supplier,HargaJualan_Member,HargaJualan_Pengedar,upah_normal_dealer,upah_master_dealer,HargaJualan_RAF,hargajualan_normal_dealer,hargajualan_master_dealer,remarks,gst_ari_nashi,kadar_gst,jumlah_gst,harga_item,jenis,harga_tanpa_gst,gst_included,jenis_trade_in,flag_upah,upah_per_gram,flag_image,'" & LM_NOW & "',no_id_gst,harga_per_gram_tanpa_gst from " & G_BELIAN_TEMP & " WHERE StatusItem='" & 10 & "'"

        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Masukkan maklumat data barang ke dalam table #data_database ### - End

'### Update maklumat di bawah ke dalam maklumat barang ### - Start
'@no_siri_produk
'@Barcode
'@bill_no_trade_in
'@no_rujukan_pelanggan_buyback

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from data_database where NoRujukanSistem='" & "PV" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm83_LM_NO_TI, "000000") & "'", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
            Y = Y + 1
            If Not IsNull(rs!ID) And Not IsNull(rs!Kod_Kategori_Produk) Then
                rs!no_siri_Produk = G_KOD_KEDAI & "-" & rs!Kod_Kategori_Produk & Format(rs!ID, "000000")
            Else
                rs!no_siri_Produk = G_KOD_KEDAI & "-" & Format(rs!ID, "000000")
            End If
            If Not IsNull(rs!ID) Then
                rs!Barcode = Format(rs!ID, "000000")
            Else
                rs!Barcode = Format(rs!ID, "000000")
            End If
            If Frm83.CB8 = 1 Then
                rs!bill_No_Trade_In = "PV" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm83_LM_NO_TI, "000000") 'No. Resit Trade In

                If Frm83.L37_Text <> vbNullString Then
                    If Frm28.L5_Text <> vbNullString Then
                        rs!no_rujukan_pelanggan_buyback = Frm28.L5_Text 'No. Rujukan Pembeli
                    Else
                        rs!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
                    End If
                Else
                    rs!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
                End If
            Else
                rs!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
            End If
            rs!susut_berat = "0.00"
            'rs!cawangan = "HQ"
            rs!nama_pekerja = Frm83_LM_EMP_NAMA
            rs.Update
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
'### Update maklumat di bawah ke dalam maklumat barang ### - End
    
        user = MDI_frm1.L3_Text
        If Frm83.CB7 = 1 Then LogAct_Memory = "[" & G_LOGIN_USER & "] Penerimaan stok baru [" & Format(Frm83_LM_No_RUJUKAN_BELIAN, "000000") & "] , Bil item [" & Y & "]."
        If Frm83.CB8 = 1 Then LogAct_Memory = "[" & G_LOGIN_USER & "] Belian trade in [" & "PV" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm83_LM_NO_TI, "000000") & "] , Bil item [" & Y & "]."
        LogDate_Memory = LM_NOW
        Call UpdateLog_Database
        
        Call Frm83_reset_list
        Call Frm83_Senarai_Belian
        
        Frm83.L69_Text = -1 'Titik Pencarian Data
        Frm83.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
        Frm83.L67_Text = 0 'Paparan Page ke-xxx
        Frm83.L68_Text = 0
            
'### Print Barcode ### - Start
        If Frm83.CB13 = 1 Then
        
            If Frm83.CB9 = 1 Then
                G_FIELD = "NoRujukanSistem"
                Call Print_All_Barcode2
            ElseIf Frm83.CB10 = 1 Then
                Call cetak_barcode_gb_all
            End If
        
        End If
'### Print Barcode ### - End
        If Frm83.CMD24.Visible = True Then
            G_KEDAI = G_CAWANGAN
            
            Call Frm84_Resit_Buyback
            Call Frm26_initial
            Call Frm27_initial
            Call Frm28_initial
        End If
        
        MsgBox "Data penerimaan stok baru telah berjaya disimpan.", vbInformation, "Info"

    End If
End If
End Sub
Private Sub CMD6_Click()
'On Error Resume Next
Dim Err(35)
Dim Frm83_LM_HARGA_TOTAL As Double
Dim Frm83_LM_CUKAI_GST As Double
Dim Frm83_LM_HARGA_SEMASA As Double 'Harga Semasa
Dim Frm83_LM_ADJUSTMENT As Double 'Adjustment
Dim Frm83_LM_BERAT As Double
Dim Frm83_LM_UPAH As Double 'Upah

Frm83_LM_HARGA_TOTAL = 0
Frm83_LM_CUKAI_GST = 0
Frm83_LM_HARGA_SEMASA = 0 'Harga Semasa
Frm83_LM_ADJUSTMENT = 0 'Adjustment
Frm83_LM_BERAT = 1
Frm83_LM_UPAH = 0 'Upah
x = 0
DATA_SAVE = 0

If Frm83.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Supplier]."
End If
If Frm83.CBB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Purity]."
End If
If Frm83.CBB3 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Kategori Produk]."
End If
If Frm83.TB1 = vbNullString Then
    x = x + 1
    Err(x) = "Tiada Maklumat [Kod Supplier]."
End If
If Frm83.TB2 = vbNullString Then
    x = x + 1
    Err(x) = "Tiada Maklumat [Kod Purity]."
End If
If Frm83.TB3 = vbNullString Then
    x = x + 1
    Err(x) = "Tiada Maklumat [Kod Kategori Produk]."
End If
'If Frm83.TB6 = vbNullString Or Frm83.TB7 = vbNullString Then
'    x = x + 1
'    Err(x) = "Maklumat [No. Siri Produk] Yang Tidak Lengkap."
'End If
If Frm83.TB36 <> vbNullString Then

    If InStr(1, Frm83.TB36, "*") <> 0 Or InStr(1, Frm83.TB36, "/") <> 0 Or InStr(1, Frm83.TB36, "\") <> 0 Or InStr(1, Frm83.TB36, "'") <> 0 Then

        x = x + 1
        Err(x) = "[Code 1] mengandungi simbol yang tidak sah."
        
    End If
    
End If
If Frm83.TB37 <> vbNullString Then

    If InStr(1, Frm83.TB37, "*") <> 0 Or InStr(1, Frm83.TB37, "/") <> 0 Or InStr(1, Frm83.TB37, "\") <> 0 Or InStr(1, Frm83.TB37, "'") <> 0 Then

        x = x + 1
        Err(x) = "[Code 2] mengandungi simbol yang tidak sah."
        
    End If
    
End If
If Frm83.CB9 = 1 And Frm83.CB4 = 0 And Frm83.CB5 = 0 Then
    x = x + 1
    Err(x) = "Sila Buat Pilihan Penerimaan [Barang Kemas] Atau [Barang Permata]."
End If
If Frm83.TB10 = vbNullString Or (Frm83.TB10 <> vbNullString And Not IsNumeric(Frm83.TB10)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Harga Asal]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB19 = vbNullString Or (Frm83.TB19 <> vbNullString And Not IsNumeric(Frm83.TB19)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Spread (%)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB20 = vbNullString Or (Frm83.TB20 <> vbNullString And Not IsNumeric(Frm83.TB20)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Harga Belian]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB21 = vbNullString Or (Frm83.TB21 <> vbNullString And Not IsNumeric(Frm83.TB21)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Harga Asal-Spread]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.TB22 = vbNullString Or (Frm83.TB22 <> vbNullString And Not IsNumeric(Frm83.TB22)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Adjusment]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm83.CB2 = 0 And Frm83.CB3 = 0 And Frm83.CB11 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan jenis cukai GST"
End If
If Frm83.CB4 = 1 Then
    If Frm83.TB8 = vbNullString Or (Frm83.TB8 <> vbNullString And Not IsNumeric(Frm83.TB8)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Berat]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB8 <> vbNullString Or (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) Then
        
        Frm83_LM_BERAT = Frm83.TB8
        
        If Frm83_LM_BERAT = 0 Then
            x = x + 1
            Err(x) = "Tiada data bagi berat barang. Berat 0 tidak dibenarkan."
        End If
    End If
    If Frm83.TB9 = vbNullString Or (Frm83.TB9 <> vbNullString And Not IsNumeric(Frm83.TB9)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Harga Semasa]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB4 = vbNullString Or (Frm83.TB4 <> vbNullString And Not IsNumeric(Frm83.TB4)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Upah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB24 = vbNullString Or (Frm83.TB24 <> vbNullString And Not IsNumeric(Frm83.TB24)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Pelanggan]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB25 = vbNullString Or (Frm83.TB25 <> vbNullString And Not IsNumeric(Frm83.TB25)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Ahli]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB26 = vbNullString Or (Frm83.TB26 <> vbNullString And Not IsNumeric(Frm83.TB26)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Silver]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB31 = vbNullString Or (Frm83.TB31 <> vbNullString And Not IsNumeric(Frm83.TB31)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Gold]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB32 = vbNullString Or (Frm83.TB32 <> vbNullString And Not IsNumeric(Frm83.TB32)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada Platinum]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB33 = vbNullString Or (Frm83.TB33 <> vbNullString And Not IsNumeric(Frm83.TB33)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Upah Kepada M.Dealer]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
ElseIf Frm83.CB5 = 1 Then
    '+++++++++++ Special Request ++++++++++ Start
    'If Frm83.TB8 = vbNullString Or (Frm83.TB8 <> vbNullString And Not IsNumeric(Frm83.TB8)) Then
    '    x = x + 1
    '    Err(x) = "Sila Masukkan [Berat]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    'End If
    'If Frm83.TB8 <> vbNullString Or (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) Then
        
    '    Frm83_LM_BERAT = Frm83.TB8
        
    '    If Frm83_LM_BERAT = 0 Then
    '        x = x + 1
    '        Err(x) = "Tiada data bagi berat barang. Berat 0 tidak dibenarkan."
    '    End If
    'End If
    '+++++++++++ Special Request ++++++++++ End
    If Frm83.TB24 = vbNullString Or (Frm83.TB24 <> vbNullString And Not IsNumeric(Frm83.TB24)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Pelanggan]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB25 = vbNullString Or (Frm83.TB25 <> vbNullString And Not IsNumeric(Frm83.TB25)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Ahli]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB26 = vbNullString Or (Frm83.TB26 <> vbNullString And Not IsNumeric(Frm83.TB26)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Silver]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB31 = vbNullString Or (Frm83.TB31 <> vbNullString And Not IsNumeric(Frm83.TB31)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Gold]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB32 = vbNullString Or (Frm83.TB32 <> vbNullString And Not IsNumeric(Frm83.TB32)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada Platinum]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm83.TB33 = vbNullString Or (Frm83.TB33 <> vbNullString And Not IsNumeric(Frm83.TB33)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat [Tetapan Harga Jualan Kepada M.Dealer]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If
If Frm83.CB8 = 1 Then
    If Frm83.TB19 = vbNullString Or (Frm83.TB19 <> vbNullString And Not IsNumeric(Frm83.TB19)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Spread]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If
If Frm83.CBB5 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Pilih [Dulang]."
End If
'If Frm83.CB3 = 1 Then
    If Frm83.TB27 = vbNullString Or (Frm83.TB27 <> vbNullString And Not IsNumeric(Frm83.TB27)) Then
        x = x + 1
        Err(x) = "Tiada Maklumat GST"
    End If
'End If
If Frm83.CB4 = 1 Then
    If Frm83.CB14 = 0 And Frm83.CB15 = 1 Then
        If Frm83.TB35 = vbNullString Or (Frm83.TB35 <> vbNullString And Not IsNumeric(Frm83.TB35)) Then
            x = x + 1
            Err(x) = "Sila buat tetapan pengiraan upah dari supplier"
        End If
    End If
End If
If Frm83.CB14 = 1 Then
    If Frm83.TB35 = vbNullString Or (Frm83.TB35 <> vbNullString And Not IsNumeric(Frm83.TB35)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Upah per gram]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If
If Frm83.TB15 <> vbNullString Then

    If InStr(1, Frm83.TB15, "*") <> 0 Or InStr(1, Frm83.TB15, "/") <> 0 Or InStr(1, Frm83.TB15, "\") <> 0 Or InStr(1, Frm83.TB15, "'") <> 0 Then

        x = x + 1
        Err(x) = "[No. Invoice] mengandungi simbol yang tidak sah."
        
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
    Note = "Adakah anda ingin simpan data ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
        
        'Frm83_LM_No_SIRI = Frm83.L3_Text 'Frm83.TB7 'No. Turutan No. Siri
        Frm83_LM_No_RUJUKAN_BELIAN = Frm83.L9_Text 'No. Rujukan Belian
        
'Re_Gen_Code:
        
'        Set rs = New ADODB.Recordset
'        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'        If Frm83.CB9 = 1 Then rs.Open "select * from Data_Database where barcode='" & Format(Frm83_LM_No_SIRI, "000000") & "'", cn, adOpenKeyset, adLockOptimistic
'        If Frm83.CB10 = 1 Then rs.Open "select * from Data_Database where barcode='" & Format(Frm83_LM_No_SIRI, "000000") & "W" & "'", cn, adOpenKeyset, adLockOptimistic
        
'        If Not rs.EOF Then
'            Frm83_LM_No_SIRI = Frm83_LM_No_SIRI + 1
            
'            rs.Close
'            Set rs = Nothing
'            GoTo Re_Gen_Code:
'        End If
        
'        rs.Close
'        Set rs = Nothing
          
'###Masukkan Data Belian Ke Dalam Database### - Start
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_BELIAN_TEMP & " where ID='" & Frm83.L13_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Frm83.L4_Text <> vbNullString Then
                rs!supplier_ID = Frm83.L4_Text 'No. ID Bagi Supplier
            Else
                rs!supplier_ID = Null 'No. ID Bagi Supplier
            End If
            If Frm83.CBB1 <> vbNullString Then
                rs!nama_Supplier = Frm83.CBB1 'Nama Supplier
            Else
                rs!nama_Supplier = Null 'Nama Supplier
            End If
            If Frm83.TB1 <> vbNullString Then
                rs!Kod_Supplier = Frm83.TB1 'Kod Supplier
            Else
                rs!Kod_Supplier = Null 'Kod Supplier
            End If
            If Frm83.L5_Text <> vbNullString Then
                rs!purity_ID = Frm83.L5_Text 'No. ID Bagi Purity
            Else
                rs!purity_ID = Null 'No. ID Bagi Purity
            End If
            If Frm83.CBB2 <> vbNullString Then
                rs!purity = Frm83.CBB2 'Purity
            Else
                rs!purity = Null 'Purity
            End If
            If Frm83.TB2 <> vbNullString Then
                rs!kod_Purity = Frm83.TB2 'Kod Purity
            Else
                rs!kod_Purity = Null 'Kod Purity
            End If
            If Frm83.L6_Text <> vbNullString Then
                rs!kategori_produk_ID = Frm83.L6_Text 'No. ID Bagi Kategori Produk
            Else
                rs!kategori_produk_ID = Null 'No. ID Bagi Kategori Produk
            End If
            If Frm83.CBB3 <> vbNullString Then
                rs!kategori_Produk = Frm83.CBB3 'Kategori Produk
            Else
                rs!kategori_Produk = Null 'Kategori Produk
            End If
            If Frm83.TB3 <> vbNullString Then
                rs!Kod_Kategori_Produk = Frm83.TB3 'Kod Kategori Produk
            Else
                rs!Kod_Kategori_Produk = Null 'Kod Kategori Produk
            End If
            'If Frm83.TB7 <> vbNullString Then
                'If Frm83.CB9 = 1 Then
                '    rs!Barcode = Frm83.TB7 'Format(Frm83_LM_No_SIRI, "000000") 'No. Barcode (6 Digit Terakhir)
                'ElseIf Frm83.CB10 = 1 Then
                '    rs!Barcode = Format(Frm83_LM_No_SIRI, "000000") & "W" 'No. Barcode (6 Digit Terakhir)
                'End If
            'Else
            '    rs!Barcode = Null 'No. Barcode (6 Digit Terakhir)
            'End If
            'rs!no_siri_Produk = Frm83.TB6 & Frm83.TB7 'No. Siri Produk
            'If Frm83.CB9 = 1 Then
            '    rs!no_siri_Produk = Frm83.TB6 & Format(Frm83_LM_No_SIRI, "000000")  'No. Siri Produk
            'ElseIf Frm83.CB10 = 1 Then
            '    rs!no_siri_Produk = Frm83.TB6 & Format(Frm83_LM_No_SIRI, "000000") & "W"  'No. Siri Produk
            'End If
            If Frm83.CB12 = 0 Then '0 : Pengiraan GST pada harga belian barang , 1 : Pengiraan GST pada upah sahaja
                rs!gst_barang_atau_upah = 0
            ElseIf Frm83.CB12 = 1 Then
                rs!gst_barang_atau_upah = 1
            End If
            If Frm83.CB4 = 1 Then
                If Frm83.TB8 <> vbNullString Then
                    rs!Berat = Format(Frm83.TB8, "0.00") 'Berat
                    rs!beza_berat = Format(Frm83.TB8, "0.00") 'Baki Berat
                Else
                    rs!Berat = Null 'Berat
                    rs!beza_berat = Null 'Baki Berat
                End If
                If Frm83.TB29 <> vbNullString Then
                    rs!riyal = Format(Frm83.TB29, "0.00") 'Berat Riyal
                Else
                    rs!riyal = Null 'Berat Riyal
                End If
                If Frm83.TB9 <> vbNullString Then
                    rs!kos_Belian_Gram = Format(Frm83.TB9, "0.00") 'Harga Per Gram (Belian)
                Else
                    rs!kos_Belian_Gram = Null 'Harga Per Gram (Belian)
                End If
                If Frm83.TB4 <> vbNullString Then
                    rs!UPAH = Frm83.TB4 'Upah (RM)
                Else
                    rs!UPAH = Null 'Upah (RM)
                End If
                
                If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) And (Frm83.TB22 <> vbNullString And IsNumeric(Frm83.TB22)) Then
                    Frm83_LM_BERAT = Frm83.TB8 'Berat
                    Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                    Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                    Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                    Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment
    
                    If Frm83.CB12 = 0 Then
                        rs!harga_Per_Gram_Item = Format((Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST) / Frm83_LM_BERAT, "0.00") 'Harga Modal Per Gram [Average] (Harga Emas + Upah) Tanpa Tax
                        rs!harga_per_gram_tanpa_gst = Format(Frm83_LM_HARGA_TOTAL / Frm83_LM_BERAT, "0.00")
                    ElseIf Frm83.CB12 = 1 Then
                        rs!harga_Per_Gram_Item = Format((((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST) / Frm83_LM_BERAT, "0.00")
                        rs!harga_per_gram_tanpa_gst = Format((((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL) / Frm83_LM_BERAT, "0.00")
                    End If
                Else
                    rs!harga_Per_Gram_Item = Null
                End If

                If Frm83.TB24 <> vbNullString Then
                    rs!Upah_Jualan = Format(Frm83.TB24, "0.00") 'Upah Jualan Kepada Pelanggan
                Else
                    rs!Upah_Jualan = Null 'Upah Jualan Kepada Pelanggan
                End If
                If Frm83.TB25 <> vbNullString Then
                    rs!Upah_Member = Format(Frm83.TB25, "0.00") 'Upah Jualan Kepada Ahli / Member
                Else
                    rs!Upah_Member = Null 'Upah Jualan Kepada Ahli / Member
                End If
                If Frm83.TB26 <> vbNullString Then
                    rs!Upah_Pengedar = Format(Frm83.TB26, "0.00") 'Upah Jualan Kepada Pengedar
                Else
                    rs!Upah_Pengedar = Null 'Upah Jualan Kepada Pengedar
                End If
                If Frm83.TB31 <> vbNullString Then
                    rs!Upah_RAF = Format(Frm83.TB31, "0.00") 'Upah Jualan Kepada RAF
                Else
                    rs!Upah_RAF = Null 'Upah Jualan Kepada RAF
                End If
                If Frm83.TB32 <> vbNullString Then
                    rs!upah_normal_dealer = Format(Frm83.TB32, "0.00") 'Upah Jualan Kepada N.Dealer
                Else
                    rs!upah_normal_dealer = Null 'Upah Jualan Kepada N.Dealer
                End If
                If Frm83.TB33 <> vbNullString Then
                    rs!upah_master_dealer = Format(Frm83.TB33, "0.00") 'Upah Jualan Kepada M.Dealer
                Else
                    rs!upah_master_dealer = Null 'Upah Jualan Kepada M.Dealer
                End If
                rs!code_Supplier = Null 'Harga Jualan Kepada Pelanggan
                rs!HargaJualan_Member = Null 'Harga Jualan Kepada Ahli / Member
                rs!HargaJualan_Pengedar = Null 'Harga Jualan Kepada Pengedar
                rs!HargaJualan_RAF = Null 'Harga Jualan Kepada RAF
                rs!hargajualan_normal_dealer = Null 'Harga Jualan Kepada N.Dealer
                rs!hargajualan_master_dealer = Null 'Harga Jualan Kepada M.Dealer
            Else
                '+++++++++++ Special Request ++++++++++ Start
                'rs!Berat = Null 'Berat
                'rs!beza_berat = Null 'Baki Berat
                'If Frm83.TB8 <> vbNullString Then
                '    rs!Berat = Format(Frm83.TB8, "0.00") 'Berat
                '    rs!beza_berat = Format(Frm83.TB8, "0.00") 'Baki Berat
                'Else
                    rs!Berat = Null 'Berat
                    rs!beza_berat = Null 'Baki Berat
                'End If
                '+++++++++++ Special Request ++++++++++ End
                rs!kos_Belian_Gram = Null 'Harga Per Gram (Belian)
                rs!UPAH = Null 'Upah (RM)
                rs!harga_Per_Gram_Item = Null 'Harga Modal Per Gram [Average] (Harga Emas + Upah) Tanpa Tax
                rs!Upah_Jualan = Null 'Upah Jualan Kepada Pelanggan
                rs!Upah_Member = Null 'Upah Jualan Kepada Ahli / Member
                rs!Upah_Pengedar = Null 'Upah Jualan Kepada Pengedar
                rs!Upah_RAF = Null 'Upah Jualan Kepada RAF
                rs!upah_normal_dealer = Null 'Upah Jualan Kepada N.Dealer
                rs!upah_master_dealer = Null 'Upah Jualan Kepada M.Dealer
            End If
            If Frm83.CB5 = 1 Then
                If Frm83.TB24 <> vbNullString Then
                    rs!code_Supplier = Format(Frm83.TB24, "0.00") 'Harga Jualan Kepada Pelanggan
                Else
                    rs!code_Supplier = Null 'Harga Jualan Kepada Pelanggan
                End If
                If Frm83.TB25 <> vbNullString Then
                    rs!HargaJualan_Member = Format(Frm83.TB25, "0.00") 'Harga Jualan Kepada Ahli / Member
                Else
                    rs!HargaJualan_Member = Null 'Harga Jualan Kepada Ahli / Member
                End If
                If Frm83.TB26 <> vbNullString Then
                    rs!HargaJualan_Pengedar = Format(Frm83.TB26, "0.00") 'Harga Jualan Kepada Pengedar
                Else
                    rs!HargaJualan_Pengedar = Null 'Harga Jualan Kepada Pengedar
                End If
                If Frm83.TB31 <> vbNullString Then
                    rs!HargaJualan_RAF = Format(Frm83.TB31, "0.00") 'Harga Jualan Kepada RAF
                Else
                    rs!HargaJualan_RAF = Null 'Harga Jualan Kepada RAF
                End If
                If Frm83.TB32 <> vbNullString Then
                    rs!hargajualan_normal_dealer = Format(Frm83.TB32, "0.00") 'Harga Jualan Kepada N.Dealer
                Else
                    rs!hargajualan_normal_dealer = Null 'Harga Jualan Kepada N.Dealer
                End If
                If Frm83.TB33 <> vbNullString Then
                    rs!hargajualan_master_dealer = Format(Frm83.TB33, "0.00") 'Harga Jualan Kepada M.Dealer
                Else
                    rs!hargajualan_master_dealer = Null 'Harga Jualan Kepada M.Dealer
                End If
                
                '+++++++++++ Special Request ++++++++++ Start
                'rs!Berat = Null 'Berat
                'rs!beza_berat = Null 'Baki Berat
                'If Frm83.TB8 <> vbNullString Then
                '    rs!Berat = Format(Frm83.TB8, "0.00") 'Berat
                '    rs!beza_berat = Format(Frm83.TB8, "0.00") 'Baki Berat
                'Else
                    rs!Berat = Null 'Berat
                    rs!beza_berat = Null 'Baki Berat
                'End If
                '+++++++++++ Special Request ++++++++++ End
                rs!kos_Belian_Gram = Null 'Harga Per Gram (Belian)
                rs!UPAH = Null 'Upah (RM)
                rs!harga_Per_Gram_Item = Null 'Harga Modal Per Gram [Average] (Harga Emas + Upah) Tanpa Tax
                rs!Upah_Jualan = Null 'Upah Jualan Kepada Pelanggan
                rs!Upah_Member = Null 'Upah Jualan Kepada Ahli / Member
                rs!Upah_Pengedar = Null 'Upah Jualan Kepada Pengedar
                rs!Upah_RAF = Null 'Upah Jualan Kepada RAF
                rs!upah_normal_dealer = Null 'Upah Jualan Kepada N.Dealer
                rs!upah_master_dealer = Null 'Upah Jualan Kepada M.Dealer
            Else
                rs!code_Supplier = Null 'Harga Jualan Kepada Pelanggan
                rs!HargaJualan_Member = Null 'Harga Jualan Kepada Ahli / Member
                rs!HargaJualan_Pengedar = Null 'Harga Jualan Kepada Pengedar
                rs!HargaJualan_RAF = Null 'Harga Jualan Kepada RAF
                rs!hargajualan_normal_dealer = Null 'Harga Jualan Kepada N.Dealer
                rs!hargajualan_master_dealer = Null 'Harga Jualan Kepada M.Dealer
            End If
            If Frm83.CB12 = 0 Then 'GST pada harga barang
            
                If Frm83.TB10 <> vbNullString Then
                    rs!kos_Belian_Item = Format(Frm83.TB10, "0.00") 'Harga Asal (RM)
                Else
                    rs!kos_Belian_Item = Null 'Harga Asal (RM)
                End If
                
            End If
            If Frm83.CB12 = 1 Then 'GST pada upah
            
                If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) Then
                    
                    Frm83_LM_BERAT = Frm83.TB8 'Berat
                    Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                    Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                    Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                    Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment
                    
                    rs!kos_Belian_Item = Format((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) + Frm83_LM_HARGA_TOTAL, "0.00") 'Harga Asal (RM)
                    
                End If
            
            End If

            If Frm83.CB8 = 1 Then
                If Frm83.TB19 <> vbNullString Then
                    rs!Spread = Format(Frm83.TB19, "0.00") 'Spread (%)
                Else
                    rs!Spread = Null 'Spread (%)
                End If
            ElseIf Frm83.CB7 = 1 Then
                rs!Spread = Null 'Spread (%)
            End If
            If Frm83.TB21 <> vbNullString Then
                rs!harga_lepas_spread = Format(Frm83.TB21, "0.00") 'Harga asal ditolak spread (RM)
            Else
                rs!harga_lepas_spread = Null 'Harga asal ditolak spread (RM)
            End If
            If Frm83.TB22 <> vbNullString Then
                rs!adjustment = Format(Frm83.TB22, "0.00") 'Adjustment (RM)
            Else
                rs!adjustment = Null 'Adjustment (RM)
            End If
            If Frm83.CB12 = 0 Then 'GST pada harga barang
                If Frm83.TB20 <> vbNullString Then
                    rs!kos_item_tanpa_tax = Format(Frm83.TB20, "0.00") 'Harga Barang + Upah Tanpa Tax
                Else
                    rs!kos_item_tanpa_tax = Null 'Harga Barang + Upah Tanpa Tax
                End If
            End If
            If Frm83.CB12 = 1 Then 'GST pada upah
            
                If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) And (Frm83.TB22 <> vbNullString And IsNumeric(Frm83.TB22)) Then
                    Frm83_LM_BERAT = Frm83.TB8 'Berat
                    Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                    Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                    Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                    Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment
                    
                    rs!kos_item_tanpa_tax = Format(((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL, "0.00")  'Harga Barang + Upah Tanpa Tax
                    
                End If
                
            End If
            'If (Frm83.TB20 <> vbNullString And IsNumeric(Frm83.TB20)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) Then
            '    Frm83_LM_BERAT = Frm83.TB8 'Berat
            '    Frm83_LM_HARGA_TOTAL = Frm83.TB20 'Harga Modal Tanpa Tax
                
            '    rs!harga_Per_Gram_Item = Format((Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST), "0.00") 'Harga Modal Per Gram [Average] (Harga Emas + Upah) Tanpa Tax
            'End If
            If Frm83.TB12 <> vbNullString Then
                rs!dimension_Panjang = Frm83.TB12 'Panjang
            Else
                rs!dimension_Panjang = Null 'Panjang
            End If
            If Frm83.TB13 <> vbNullString Then
                rs!dimension_Lebar = Frm83.TB13 'Lebar
            Else
                rs!dimension_Lebar = Null 'Lebar
            End If
            If Frm83.TB14 <> vbNullString Then
                rs!dimension_Saiz = Frm83.TB14 'Saiz
            Else
                rs!dimension_Saiz = Null 'Saiz
            End If
            If Frm83.TB36 <> vbNullString Then 'Code 1
                rs!code1 = UCase(Frm83.TB36)
            Else
                rs!code1 = Null
            End If
            If Frm83.TB37 <> vbNullString Then 'Code 2
                rs!code2 = UCase(Frm83.TB37)
            Else
                rs!code2 = Null
            End If
            If Frm83.CBB5 <> vbNullString Then
                rs!dulang = Frm83.CBB5 'Dulang
            Else
                rs!dulang = Null 'Dulang
            End If
            If Frm83.TB16 <> vbNullString Then
                rs!remarks = UCase(Frm83.TB16) 'Remarks
            Else
                rs!remarks = Null 'Remarks
            End If
            If Frm83.TB34 <> vbNullString Then
                rs!no_cert = UCase(Frm83.TB34) 'No. Cert
            Else
                rs!no_cert = Null 'No. Cert
            End If

            If Frm83.CB2 = 1 Then
            
                rs!gst_ari_nashi = 0 'Status Cukai GST : 0 : ZR(L) , 1 : SR
                rs!kadar_gst = Null 'Kadar GST (%)
                rs!jumlah_gst = Null 'Jumlah Cukai GST (RM)
                rs!gst_included = Null '0 : Harga tidak termasuk GST , 1 : Harga termasuk GST
                
            ElseIf Frm83.CB3 = 1 Then
            
                rs!gst_included = 0 '0 : Harga tidak termasuk GST , 1 : Harga termasuk GST
                rs!gst_ari_nashi = 1 'Status Cukai GST : 0 : ZR(L) , 1 : SR
                rs!kadar_gst = Format(Frm83.L8_Text, "0.00") 'Kadar GST (%)
                rs!jumlah_gst = Format(Frm83.TB27, "0.00") 'Jumlah Cukai GST (RM)
              
            ElseIf Frm83.CB11 = 1 Then
    
                rs!gst_included = 1 '0 : Harga tidak termasuk GST , 1 : Harga termasuk GST
                rs!gst_ari_nashi = 1 'Status Cukai GST : 0 : ZR(L) , 1 : SR
                rs!kadar_gst = Format(Frm83.L8_Text, "0.00") 'Kadar GST (%)
                rs!jumlah_gst = Format(Frm83.TB27, "0.00") 'Jumlah Cukai GST (RM)
                
            End If
        
            If Frm83.L40_Text <> vbNullString Then
                rs!harga_tanpa_gst = Format(Frm83.L40_Text, "0.00") 'Harga Barang Tanpa Tax (kalau gst included)
            Else
                rs!harga_tanpa_gst = Null 'Harga Barang Tanpa Tax (kalau gst included)
            End If

            If Frm83.CB5 = 1 Then 'Barang Permata
            
                If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) Then
                    Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                    Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                    
                    rs!harga_item = Format(Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST, "0.00") 'Jumlah Harga Modal Keseluruhan @ Harga Barang + Upah + GST (RM)
                End If
                
            End If
            
            If Frm83.CB4 = 1 Or Frm83.CB10 = 1 Then 'Barang Kemas / Gold Bar
            
                If (Frm83.L40_Text <> vbNullString And IsNumeric(Frm83.L40_Text)) And (Frm83.TB27 <> vbNullString And IsNumeric(Frm83.TB27)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) And (Frm83.TB22 <> vbNullString And IsNumeric(Frm83.TB22)) Then
                    Frm83_LM_BERAT = Frm83.TB8 'Berat
                    Frm83_LM_CUKAI_GST = Frm83.TB27  'Jumlah Cukai GST
                    Frm83_LM_HARGA_TOTAL = Frm83.L40_Text 'Harga Modal Tanpa Tax
                    Frm83_LM_HARGA_SEMASA = Frm83.TB9 'Harga Semasa
                    Frm83_LM_ADJUSTMENT = Frm83.TB22 'Adjustment
                    
                    If Frm83.CB12 = 0 Then
                        rs!harga_item = Format(Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST, "0.00") 'Jumlah Harga Modal Keseluruhan @ Harga Barang + Upah + GST (RM)
                    ElseIf Frm83.CB12 = 1 Then
                        rs!harga_item = Format(((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) - Frm83_LM_ADJUSTMENT) + Frm83_LM_HARGA_TOTAL + Frm83_LM_CUKAI_GST, "0.00") 'Jumlah Harga Modal Keseluruhan @ Harga Barang + Upah + GST (RM)
                    End If
                    
                End If
                
            End If
            
'### Kod Bagi Status ###
'==========================================
'0 : Tidak Aktif (Tidak Perlu Buat Apa-Apa)
'1 : Aktif - Kemasukkan Data Baru
'2 : Barang Sudah Terjual ****Tidak Dibenarkan Untuk Diedit Atau Dipadamkan
'3 : Kemasukkan Data Baru
'4 : Data Diedit
'5 : Keluarkan Data Dari Database (Bagi Item Asal Yang Dipadamkan)
'6 : Ignore Kemasukkan Data Ke Dalam Database

            rs!StatusItem = 10 '0 : Padam Data , 1 : Aktif (Tiada Perubahan) , 2 : Sudah Terjual , 3 : Data Baru , 4 : Data Diedit

'### Jenis ###
'0 : BK
'1 : Barang permata
'2 : Emas terpakai BK
'3 : Emas terpakai permata
'4 : gold Bar
'5 : Emas terpakai gold bar
'6 : Trade In BK
'7 : Trade In Barang Permata
'8 : Trade In Gold Bar

'=========================================================
'Frm83.L41_Text
'0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
'=========================================================

            'If Frm83.L41_Text = 1 Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
            
                'If Frm83.CB8 = 1 Then 'Buyback / Trade in
            '        If Frm83.CB4 = 1 Then 'Barang kemas
            '            rs!jenis = 6
            '        ElseIf Frm83.CB5 = 1 Then 'Barang permata
            '            rs!jenis = 7
            '        End If
            '        If Frm83.CB10 = 1 Then 'Gold bar
            '            rs!jenis = 8
            '        End If
                'End If
            
            'ElseIf Frm83.L41_Text = 0 Or Frm83.L41_Text = 2 Then
            
                If Frm83.CB7 = 1 Then 'Penerimaan stok baru
                    If Frm83.CB4 = 1 Then 'Barang kemas
                        rs!jenis = 0
                    ElseIf Frm83.CB5 = 1 Then 'Barang permata
                        rs!jenis = 1
                    End If
                    If Frm83.CB10 = 1 Then 'Gold bar
                        rs!jenis = 4
                    End If
                ElseIf Frm83.CB8 = 1 Then 'Buyback / Trade in
                    If Frm83.CB4 = 1 Then 'Barang kemas
                        rs!jenis = 2
                    ElseIf Frm83.CB5 = 1 Then 'Barang permata
                        rs!jenis = 3
                    End If
                    If Frm83.CB10 = 1 Then 'Gold bar
                        rs!jenis = 5
                    End If
                End If
            
            'End If
            
            If Frm83.L41_Text = 0 Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
                
                rs!jenis_trade_in = 0 '0 : Belian emas dari pelanggan , 1 : Belian emas secara trade in
                
            ElseIf Frm83.L41_Text = 1 Then '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
            
                rs!jenis_trade_in = 1 '0 : Belian emas dari pelanggan , 1 : Belian emas secara trade in
                
            End If
            rs!flag_image = 0
            rs!Image = Null

            If Frm83.TB35 <> vbNullString Then
                rs!upah_per_gram = Format(Frm83.TB35, "0.00")
            Else
                rs!upah_per_gram = "0.00"
            End If
            If Frm83.CB14 = 1 Then
                rs!flag_upah = 0
            ElseIf Frm83.CB15 = 1 Then
                rs!flag_upah = 1
                
                If IsNumeric(Frm83.TB8) And Frm83.TB8 <> 0 Then
                    Frm83_LM_BERAT = Frm83.TB8 'Berat
                End If
                
                Frm83_LM_UPAH = Frm83.TB4 'Upah
                
                rs!upah_per_gram = Format(Frm83_LM_UPAH / Frm83_LM_BERAT, "0.00")
            End If
            
            If Frm83.TB28 <> vbNullString Then
                rs!no_id_gst = UCase(Frm83.TB28)
            Else
                rs!no_id_gst = Null
            End If
            If Frm83.TB15 <> vbNullString Then
                rs!bill_No_Belian = UCase(Frm83.TB15)
            Else
                rs!bill_No_Belian = Null
            End If
            rs!tarikh_belian = Frm83.DTPicker1
            rs!terminal = G_TERMINAL
            'If Frm83.L32_Text = 1 Then
            '    Set rs2 = New ADODB.Recordset
            '    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            '    rs2.Open "select * from 1_image_item_temp where initial_flag='" & "1" & "'", cn, adOpenKeyset, adLockOptimistic
                
            '    If Not rs2.EOF Then
            '        rs!flag_image = 1
            '        rs!Image = rs2!Image
            '    End If
                
            '    rs2.Close
            '    Set rs2 = Nothing
            'End If
            
            rs.Update
            DATA_SAVE = 1
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_SAVE = 1 Then
            Call Frm83_Cancel_Edit
            
            Call Frm83_Reset_Form
            Call Frm83_Senarai_Belian_Header
            Call Frm83_Senarai_Belian
            
            MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
        End If
    End If
End If
End Sub
Private Sub CMD7_Click()
'on error resume next
Call Frm83_Cancel_Edit

If Frm83.CB9 = 1 Then

    Frm83.CB5 = 0
    
    Frm83.TB8.Locked = False
    Frm83.TB9.Locked = False
    Frm83.TB4.Locked = False
    
    Frm83.TB8.BackColor = &HFFFFFF
    Frm83.TB9.BackColor = &HFFFFFF
    Frm83.TB4.BackColor = &HFFFFFF
    
    Frm83.L27_Text = "Upah Jualan Pelanggan    RM"
    Frm83.L28_Text = "Upah Jualan Member       RM"
    Frm83.L29_Text = "Upah Jualan Pengedar     RM"
    
    Frm83.L27_Text = "Upah Jualan Pelanggan    RM"
    Frm83.L28_Text = "Upah Jualan Ahli               RM"
    Frm83.L29_Text = "Upah Jualan Silver            RM"

    Frm83.CB4.Enabled = True
    Frm83.CB5.Enabled = True

End If

Frm83.TB8 = "0.00"
Frm83.TB9 = "0.00"
Frm83.TB4 = 0
Frm83.TB10 = "0.00"
Frm83.TB35 = "0.00"
End Sub


Private Sub Form_Load()
'on error resume next
GLOBAL_DISABLE = 0
Frm83.L21_Text = 0 '0 : Data Baru , 1:  Data Diedit
Frm83.CMD10.Visible = False
Frm83.CMD11.Visible = False
Frm83.CMD2.Visible = True
Frm83.CMD5.Visible = True
Frm83.TB4 = 0 'Upah
Frm83.L8_Text = vbNullString 'Kadar Cukai GST
Call Frm83_background_color
'Frm83.CB4 = 1
End Sub
Private Sub Frm83_SM_Edit_Click()
'on error resume next
DATA_FOUND = 0
Frm83_LM_ID = vbNullString

If IsNumeric(Frm83.ListView2.SelectedItem.Index) Then
    
    Frm83_LM_ID = Frm83.ListView2.SelectedItem.Index
    
    If Frm83_LM_ID <> vbNullString Then
    
        If Frm83_LM_ID <> vbNullString Then
        
            Note = "Adakah anda ingin edit data ini ?"
                    
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

            If Answer = vbYes Then
            
                Call Frm83_Edit_Data
                
            End If
            
        End If
        
    End If
    
End If
End Sub
Private Sub Frm83_SM_Padam_Click()
'on error resume next
DATA_FOUND = 0

DATA_FOUND = 0
Frm83_LM_ID = vbNullString

If IsNumeric(Frm83.ListView2.SelectedItem.Index) Then
    
    Frm83_LM_ID = Frm83.ListView2.ListItems(Frm83.ListView2.SelectedItem.Index)
    
    If Frm83_LM_ID <> vbNullString Then
        
        Note = "Adakah anda ingin padam data ini ?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        If Answer = vbNo Then
            'Exit Sub
        End If
        If Answer = vbYes Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from " & G_BELIAN_TEMP & " where ID='" & Frm83_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!StatusItem) Then
'### Kod Bagi Status ###
'==========================================
'0 : Tidak Aktif (Tidak Perlu Buat Apa-Apa)
'1 : Aktif - Kemasukkan Data Baru
'2 : Barang Sudah Terjual ****Tidak Dibenarkan Untuk Diedit Atau Dipadamkan
'3 : Kemasukkan Data Baru
'4 : Data Diedit
'5 : Keluarkan Data Dari Database (Bagi Item Asal Yang Dipadamkan)
'6 : Ignore Kemasukkan Data Ke Dalam Database

                    If rs!StatusItem = "10" Then
                        If Frm83.L21_Text = 0 Then
                            Frm83_LM_STATUS = "0"
                            DATA_FOUND = 1
                        ElseIf Frm83.L21_Text = 1 Then
                            Frm83_LM_STATUS = "5"
                            DATA_FOUND = 1
                        End If
                    ElseIf rs!StatusItem = "4" Then
                        Frm83_LM_STATUS = "5"
                        DATA_FOUND = 1
                    ElseIf rs!StatusItem = "3" Then
                        Frm83_LM_STATUS = "0"
                        DATA_FOUND = 1
                    End If
                    rs!terminal = G_TERMINAL
                    
                    If rs!StatusItem = "11" Then
                        MsgBox "Item Ini Tidak Dibenarkan Untuk Dipadamkan Kerana Telah Dijual.", vbExclamation, "Info"
                    ElseIf rs!StatusItem = "12" Then
                        MsgBox "Item Ini Tidak Dibenarkan Untuk Dipadamkan Kerana Telah Dijual Secara Potong.", vbExclamation, "Info"
                    ElseIf rs!StatusItem = "13" Then
                        MsgBox "Item Ini Tidak Dibenarkan Untuk Dipadamkan Kerana Telah Dijual Secara Potong.", vbExclamation, "Info"
                    ElseIf rs!StatusItem = "14" Or rs!StatusItem = "21" Or rs!StatusItem = "22" Then
                        MsgBox "Item Ini Tidak Dibenarkan Untuk Dipadamkan Kerana Ditempah Oleh Pelanggan.", vbExclamation, "Info"
                    ElseIf rs!StatusItem = "15" Or rs!StatusItem = "19" Or rs!StatusItem = "20" Then
                        MsgBox "Item Ini Tidak Dibenarkan Untuk Dipadamkan Kerana Dibeli Secara Ansuran.", vbExclamation, "Info"
                    ElseIf rs!StatusItem = "16" Then
                        MsgBox "Item Ini Tidak Dibenarkan Untuk Dipadamkan Kerana Telah Dihantar Ke Ar-Rahnu.", vbExclamation, "Info"
                    ElseIf rs!StatusItem = "17" Then
                        MsgBox "Item Ini Tidak Dibenarkan Untuk Dipadamkan Kerana Dijual Secara ETA.", vbExclamation, "Info"
                    ElseIf rs!StatusItem = "23" Then
                        MsgBox "Item Ini Tidak Dibenarkan Untuk Dipadamkan Kerana Dihantar Ke Supplier/Kilang.", vbExclamation, "Info"
                    ElseIf rs!StatusItem = "24" Then
                        MsgBox "Item Ini Tidak Dibenarkan Untuk Dipadamkan Kerana Telah Dihantar Ke Supplier/Kilang.", vbExclamation, "Info"
                    ElseIf rs!StatusItem = "25" Then
                        MsgBox "Item Ini Tidak Dibenarkan Untuk Dipadamkan Kerana Telah Diagihkan Ke Cawangan.", vbExclamation, "Info"
                    ElseIf rs!StatusItem = "26" Then
                        MsgBox "Item Ini Tidak Dibenarkan Untuk Dipadamkan Kerana Telah Dijual Oleh Cawangan.", vbExclamation, "Info"
                    ElseIf rs!StatusItem = "0" Then
                        MsgBox "Item Ini Tidak Dibenarkan Untuk Dipadamkan Kerana Tiada Rekod Di Dalam Sistem.", vbExclamation, "Info"
                    ElseIf rs!StatusItem = "27" Or rs!StatusItem = "28" Then
                        MsgBox "Item Ini Telah Dijual Dari Menu GDN.", vbExclamation, "Info"
                    ElseIf rs!StatusItem = "29" Then
                        MsgBox "Item Ini Telah Diubah Status Kepada Hilang , Dicuri Dan Sebagainya.", vbExclamation, "Info"
                    End If
                    
                    If DATA_FOUND = 1 Then
                        rs!StatusItem = Frm83_LM_STATUS
                        rs.Update
                    End If
                End If
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_FOUND = 1 Then
                
                GM_NEXT_PREV = 2
                
                Call Frm83_Senarai_Belian_Header
                Call Frm83_Senarai_Belian
                
                MsgBox "Item Telah Dikeluarkan Dari Senarai.", vbInformation, "Info"
            End If
        End If
        
    End If
End If
End Sub

Private Sub L2_Text_Change()
'On Error Resume Next
If Frm83.CMD6.Visible = True Or Frm83.CMD13.Visible = True Or Frm83.CMD22.Visible = True Then
    If Frm83.L14_Text.Visible = True Then
        Frm83.L14_Text.Visible = False
    Else
        Frm83.L14_Text.Visible = True
    End If
End If

'If Frm83.L38_Text.Visible = True Then
'    Frm83.L38_Text.Visible = False
'Else
'    Frm83.L38_Text.Visible = True
'End If
End Sub



Private Sub ListView1_Click()
'on error resume next
LM_KEY = Frm83.ListView1.SelectedItem.Key

If LM_KEY = "Data Item" Then

    Frm83.Frame9.Visible = False
    Frm83.Frame1.Visible = True
            
ElseIf LM_KEY = "Senarai Item" Then
    
    If Frm83.L10_Text <> "" Then
        If Frm83.L10_Text <> "0" Then
        
            Frm83.L69_Text = -1 'Titik Pencarian Data
            Frm83.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
            Frm83.L67_Text = 0 'Paparan Page ke-xxx
            Frm83.L68_Text = 0
            
            GM_NEXT_PREV = 0
    
            Call Frm83_Senarai_Belian_Header
            Call Frm83_Senarai_Belian
            
            Frm83.Frame9.Visible = True
            Frm83.Frame1.Visible = False
            
        Else
        
            MsgBox "Tiada data di dalam senarai.", vbInformation, "Info"
            
        End If
    Else
    
        MsgBox "Tiada data di dalam senarai.", vbInformation, "Info"
        
    End If

End If
End Sub

Private Sub ListView2_DblClick()
'On Error Resume Next
Frm83_LM_No_ID = vbNullString

If IsNumeric(Frm83.ListView2.SelectedItem.Index) Then
    
    Frm83_LM_No_ID = Frm83.ListView2.SelectedItem.Index
    
    If Frm83_LM_No_ID <> vbNullString Then
        
        PopupMenu Frm83_PM_Menu1
        
    Else
        
        MsgBox "Tiada data.", vbExclamation, "Info"
        
    End If
    
Else

    MsgBox "Tiada data.", vbExclamation, "Info"
    
End If
End Sub



Private Sub Pic3_Click()

End Sub

Private Sub TB10_Change()
'On Error Resume Next
Call frm83_kiraan_harga_selepas_spread

Exit Sub

Dim Frm83_LM_HARGA_ASAL As Double
Dim Frm83_LM_SPREAD As Double

If (Frm83.TB10 <> vbNullString And IsNumeric(Frm83.TB10)) Then
    Frm83.TB20 = Format(Frm83.TB10, "0.00") 'Jumlah Harga Belian Setelah Ditolak Spread
Else
    Frm83.TB20 = "0.00" 'Jumlah Harga Belian Setelah Ditolak Spread
End If

If (Frm83.TB10 <> vbNullString And IsNumeric(Frm83.TB10)) Then
    If Frm83.CB8 = 1 Then
        If Frm83.TB19 <> vbNullString And IsNumeric(Frm83.TB19) Then
            Frm83_LM_HARGA_ASAL = Frm83.TB10 'Harga Asal
            Frm83_LM_SPREAD = Frm83.TB19 'Spread
            
            Frm83.TB21 = Format(Frm83_LM_HARGA_ASAL - ((Frm83_LM_SPREAD / 100) * Frm83_LM_HARGA_ASAL), "0.00") 'Jumlah Harga Belian Setelah Ditolak Spread
        Else
            Frm83.TB21 = "0.00" 'Jumlah Harga Belian Setelah Ditolak Spread
        End If
    Else
        Frm83.TB21 = Format(Frm83.TB10, "0.00") 'Jumlah Harga Belian Setelah Ditolak Spread
    End If
Else
    Frm83.TB21 = "0.00" 'Jumlah Harga Belian Setelah Ditolak Spread
End If
End Sub

Private Sub TB19_Change()
'On Error Resume Next
Call frm83_kiraan_harga_selepas_spread

Exit Sub

Dim Frm83_LM_HARGA_ASAL As Double
Dim Frm83_LM_SPREAD As Double

If (Frm83.TB10 <> vbNullString And IsNumeric(Frm83.TB10)) Then
    Frm83.TB20 = Format(Frm83.TB10, "0.00") 'Jumlah Harga Belian Setelah Ditolak Spread
Else
    Frm83.TB20 = "0.00" 'Jumlah Harga Belian Setelah Ditolak Spread
End If

If (Frm83.TB19 <> vbNullString And IsNumeric(Frm83.TB19)) Then
    If Frm83.CB8 = 1 Then
        If Frm83.TB10 <> vbNullString And IsNumeric(Frm83.TB10) Then
            Frm83_LM_HARGA_ASAL = Frm83.TB10 'Harga Asal
            Frm83_LM_SPREAD = Frm83.TB19 'Spread
            
            Frm83.TB21 = Format(Frm83_LM_HARGA_ASAL - ((Frm83_LM_SPREAD / 100) * Frm83_LM_HARGA_ASAL), "0.00") 'Jumlah Harga Belian Setelah Ditolak Spread
        Else
            Frm83.TB21 = "0.00" 'Jumlah Harga Belian Setelah Ditolak Spread
        End If
    Else
        Frm83.TB21 = Format(Frm83.TB10, "0.00") 'Jumlah Harga Belian Setelah Ditolak Spread
    End If
Else
    Frm83.TB21 = "0.00" 'Jumlah Harga Belian Setelah Ditolak Spread
End If
End Sub
Private Sub TB20_Change()
'On Error Resume Next
Call kiraan_gst_belian

Exit Sub

Dim frm83_LM_KADAR_GST As Double
Dim Frm83_LM_HARGA As Double

Frm83_LM_HARGA = 0

If Frm83.CB11 = 0 Then

    If Frm83.CB3 = 1 And (Frm83.TB20 <> vbNullString And IsNumeric(Frm83.TB20)) And (Frm83.L8_Text <> vbNullString And IsNumeric(Frm83.L8_Text)) Then
        frm83_LM_KADAR_GST = Frm83.L8_Text 'Jumlah Kadar GST (%)

        If Frm83.CB12 = 0 Then
            Frm83_LM_HARGA = Frm83.TB20 'Harga (RM)
        ElseIf Frm83.CB12 = 1 Then
            If Frm83.TB4 <> vbNullString And IsNumeric(Frm83.TB4) Then
                Frm83_LM_HARGA = Frm83.TB4 'Upah (RM)
            End If
        End If
        
        Frm83.TB27 = Format((frm83_LM_KADAR_GST / 100) * Frm83_LM_HARGA, "#,##0.00") 'Jumlah Cukai GST (RM)
        Frm83.L40_Text = Format(Frm83_LM_HARGA, "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
    Else
        Frm83.TB27 = Format(0, "#,##0.00") 'Jumlah Cukai GST (RM)
    End If
    
    If Frm83.CB3 = 0 And (Frm83.TB20 <> vbNullString And IsNumeric(Frm83.TB20)) Then
    
        If Frm83.CB12 = 0 Then
            Frm83_LM_HARGA = Frm83.TB20 'Harga (RM)
        ElseIf Frm83.CB12 = 1 Then
            If Frm83.TB4 <> vbNullString And IsNumeric(Frm83.TB4) Then
                Frm83_LM_HARGA = Frm83.TB4 'Upah (RM)
            End If
        End If
        
        Frm83.L40_Text = Format(Frm83_LM_HARGA, "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
    End If
    
ElseIf Frm83.CB11 = 1 Then

    If Frm83.CB3 = 1 And (Frm83.TB20 <> vbNullString And IsNumeric(Frm83.TB20)) And (Frm83.L8_Text <> vbNullString And IsNumeric(Frm83.L8_Text)) Then
        frm83_LM_KADAR_GST = Frm83.L8_Text 'Jumlah Kadar GST (%)

        If Frm83.CB12 = 0 Then
            Frm83_LM_HARGA = Frm83.TB20 'Harga (RM)
        ElseIf Frm83.CB12 = 1 Then
            If Frm83.TB4 <> vbNullString And IsNumeric(Frm83.TB4) Then
                Frm83_LM_HARGA = Frm83.TB4 'Upah (RM)
            End If
        End If
        
        Frm83.L40_Text = Format(Frm83_LM_HARGA / (1 + (frm83_LM_KADAR_GST / 100)), "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
        Frm83.TB27 = Format(Frm83_LM_HARGA - (Frm83_LM_HARGA / (1 + (frm83_LM_KADAR_GST / 100))), "#,##0.00") 'Jumlah Cukai GST (RM)
    Else
        Frm83.TB27 = "0.00" 'Jumlah Cukai GST (RM)
    End If
    
End If
End Sub
Private Sub TB21_Change()
'On Error Resume Next

Call frm83_harga_belian_lepas_adjust

Exit Sub

If (Frm83.TB21 <> vbNullString And IsNumeric(Frm83.TB21)) And (Frm83.TB22 <> vbNullString And IsNumeric(Frm83.TB22)) Then
    Frm83.TB20 = Format(Frm83.TB21 - Frm83.TB22, "0.00") 'Harga
Else
    If Frm83.TB21 <> vbNullString And IsNumeric(Frm83.TB21) Then
        Frm83.TB20 = Format(Frm83.TB21, "0.00") 'Harga Belian - Setelah DiTolak Semua
    Else
        Frm83.TB20 = "0.00"
    End If
End If
End Sub
Private Sub TB22_Change()
'On Error Resume Next

Call frm83_harga_belian_lepas_adjust

Exit Sub


If (Frm83.TB21 <> vbNullString And IsNumeric(Frm83.TB21)) And (Frm83.TB22 <> vbNullString And IsNumeric(Frm83.TB22)) Then
    Frm83.TB20 = Format(Frm83.TB21 - Frm83.TB22, "0.00") 'Harga
Else
    If Frm83.TB21 <> vbNullString And IsNumeric(Frm83.TB21) Then
        Frm83.TB20 = Format(Frm83.TB21, "0.00") 'Harga Belian - Setelah DiTolak Semua
    Else
        Frm83.TB20 = "0.00"
    End If
End If
End Sub

Private Sub TB35_Change()
'On Error Resume Next
Call Frm83_kira_upah
End Sub

Private Sub TB4_Change()
'On Error Resume Next
Call frm83_kiraan_harga_asal
Call kiraan_gst_belian

Exit Sub

Dim Frm83_LM_BERAT As Double
Dim Frm83_LM_UPAH As Double
Dim Frm83_LM_HargaPerGram As Double

If (Frm83.TB4 <> vbNullString And IsNumeric(Frm83.TB4)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) Then
    Frm83_LM_BERAT = Frm83.TB8 'Berat
    Frm83_LM_UPAH = Frm83.TB4 'Upah
    Frm83_LM_HargaPerGram = Frm83.TB9 'Harga Per Gram
    
    Frm83.TB10 = Format((Frm83_LM_BERAT * Frm83_LM_HargaPerGram) + Frm83_LM_UPAH, "0.00") 'Harga
Else
    Frm83.TB10 = "0.00"
End If
End Sub

Private Sub TB40_Change()
'On Error Resume Next
Call frm83_kiraan_cara_bayaran
End Sub
Private Sub TB41_Change()
'On Error Resume Next
Call frm83_kiraan_cara_bayaran
End Sub
Private Sub TB8_Change()
'On Error Resume Next
Call frm83_kiraan_harga_asal

If (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.L30_Text <> vbNullString And IsNumeric(Frm83.L30_Text)) Then
    Frm83.TB29 = Format(Frm83.TB8 / Frm83.L30_Text, "0.00") 'Berat Riyal
Else
    Frm83.TB29 = "0.00"
End If

Exit Sub

Dim Frm83_LM_BERAT As Double
Dim Frm83_LM_UPAH As Double
Dim Frm83_LM_HargaPerGram As Double

If (Frm83.TB4 <> vbNullString And IsNumeric(Frm83.TB4)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) Then
    Frm83_LM_BERAT = Frm83.TB8 'Berat
    Frm83_LM_UPAH = Frm83.TB4 'Upah
    Frm83_LM_HargaPerGram = Frm83.TB9 'Harga Per Gram
    
    Frm83.TB10 = Format((Frm83_LM_BERAT * Frm83_LM_HargaPerGram) + Frm83_LM_UPAH, "0.00") 'Harga
Else
    Frm83.TB10 = "0.00"
End If



Call Frm83_kira_upah
End Sub
Private Sub TB9_Change()
'On Error Resume Next
Call frm83_kiraan_harga_asal

Exit Sub

Dim Frm83_LM_BERAT As Double
Dim Frm83_LM_UPAH As Double
Dim Frm83_LM_HargaPerGram As Double

If (Frm83.TB4 <> vbNullString And IsNumeric(Frm83.TB4)) And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) Then
    Frm83_LM_BERAT = Frm83.TB8 'Berat
    Frm83_LM_UPAH = Frm83.TB4 'Upah
    Frm83_LM_HargaPerGram = Frm83.TB9 'Harga Per Gram
    
    Frm83.TB10 = Format((Frm83_LM_BERAT * Frm83_LM_HargaPerGram) + Frm83_LM_UPAH, "0.00") 'Harga
Else
    Frm83.TB10 = "0.00"
End If
End Sub
Private Sub Tmr1_Timer()
'on error resume next
Frm83.L1_Text = DateTime.Date
Frm83.L2_Text = DateTime.Time$
End Sub
