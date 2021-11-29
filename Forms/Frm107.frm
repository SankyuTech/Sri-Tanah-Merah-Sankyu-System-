VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm107 
   Caption         =   "Hantaran barang kepada supplier / kilang"
   ClientHeight    =   13035
   ClientLeft      =   120
   ClientTop       =   -9135
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
   Icon            =   "Frm107.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13035
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   9495
      Left            =   20760
      ScaleHeight     =   9495
      ScaleWidth      =   14325
      TabIndex        =   80
      Top             =   3960
      Visible         =   0   'False
      Width           =   14325
   End
   Begin VB.PictureBox Pic1 
      BorderStyle     =   0  'None
      Height          =   11895
      Left            =   480
      ScaleHeight     =   11895
      ScaleWidth      =   21255
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   21255
      Begin VB.CommandButton CMD21 
         BackColor       =   &H000080FF&
         Caption         =   "Batal"
         Height          =   405
         Left            =   13080
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm107.frx":0ECA
         MousePointer    =   99  'Custom
         TabIndex        =   144
         Top             =   10200
         Width           =   2025
      End
      Begin VB.CommandButton CMD18 
         BackColor       =   &H000080FF&
         Caption         =   "Simpan Data"
         Height          =   405
         Left            =   1560
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm107.frx":11D4
         MousePointer    =   99  'Custom
         TabIndex        =   142
         Top             =   3960
         Width           =   2025
      End
      Begin VB.CommandButton CMD19 
         BackColor       =   &H000080FF&
         Caption         =   "Batal"
         Height          =   405
         Left            =   3720
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm107.frx":14DE
         MousePointer    =   99  'Custom
         TabIndex        =   141
         Top             =   3960
         Width           =   2025
      End
      Begin VB.ComboBox CBB5 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   315
         Left            =   8700
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   1080
         Width           =   4000
      End
      Begin VB.TextBox TB3 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   720
         TabIndex        =   68
         Text            =   "TB3"
         Top             =   10440
         Width           =   1260
      End
      Begin VB.CommandButton CMD4 
         BackColor       =   &H000080FF&
         Caption         =   "Simpan Data"
         Height          =   405
         Left            =   2640
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm107.frx":17E8
         MousePointer    =   99  'Custom
         TabIndex        =   45
         Top             =   3960
         Width           =   2025
      End
      Begin VB.TextBox TB2 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1400
         TabIndex        =   42
         Text            =   "TB2"
         Top             =   3405
         Width           =   1260
      End
      Begin VB.ComboBox CBB4 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   315
         ItemData        =   "Frm107.frx":1AF2
         Left            =   1400
         List            =   "Frm107.frx":1AF4
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   3075
         Width           =   5685
      End
      Begin VB.TextBox TB1 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   1400
         MultiLine       =   -1  'True
         TabIndex        =   38
         Text            =   "Frm107.frx":1AF6
         Top             =   2520
         Width           =   5685
      End
      Begin VB.ComboBox CBB3 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   315
         ItemData        =   "Frm107.frx":1AFA
         Left            =   1400
         List            =   "Frm107.frx":1AFC
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   720
         Width           =   5685
      End
      Begin VB.CommandButton CMD5 
         BackColor       =   &H00FFFFFF&
         Height          =   740
         Left            =   8640
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm107.frx":1AFE
         MousePointer    =   99  'Custom
         Picture         =   "Frm107.frx":1E08
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Paparan sebelumnya"
         Top             =   9840
         Width           =   1000
      End
      Begin VB.CommandButton CDM6 
         BackColor       =   &H00FFFFFF&
         Height          =   740
         Left            =   9720
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm107.frx":2747
         MousePointer    =   99  'Custom
         Picture         =   "Frm107.frx":2A51
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Paparan seterusnya"
         Top             =   9840
         Width           =   1000
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   4965
         Left            =   120
         TabIndex        =   26
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   4800
         Width           =   10605
         _ExtentX        =   18706
         _ExtentY        =   8758
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
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   8700
         TabIndex        =   74
         Top             =   720
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
         Format          =   415891456
         CurrentDate     =   41561
      End
      Begin VB.CommandButton CMD7 
         BackColor       =   &H000080FF&
         Caption         =   "Simpan (Hantar Barang)"
         Height          =   405
         Left            =   10920
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm107.frx":3377
         MousePointer    =   99  'Custom
         TabIndex        =   77
         Top             =   10200
         Width           =   2025
      End
      Begin VB.CommandButton CMD20 
         BackColor       =   &H000080FF&
         Caption         =   "Simpan (Hantar Barang)"
         Height          =   405
         Left            =   10920
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm107.frx":3681
         MousePointer    =   99  'Custom
         TabIndex        =   143
         Top             =   10200
         Width           =   2025
      End
      Begin VB.Label L59_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L59_Text"
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
         Left            =   1890
         TabIndex        =   148
         Top             =   10080
         Width           =   2175
      End
      Begin VB.Label L57_Text 
         BackColor       =   &H008080FF&
         Caption         =   "L57_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   14040
         TabIndex        =   145
         Top             =   8760
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan 1.00 jika tiada tetapan yang diberikan oleh pihak kilang / supplier (Penerima)"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2700
         TabIndex        =   79
         Top             =   3600
         Width           =   6945
      End
      Begin VB.Label L31_Text 
         BackColor       =   &H008080FF&
         Caption         =   "L31_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   14040
         TabIndex        =   78
         Top             =   8400
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pekerja * :"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7440
         TabIndex        =   76
         Top             =   1125
         Width           =   2295
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Hantar * :"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7440
         TabIndex        =   75
         Top             =   765
         Width           =   2385
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
         Left            =   4080
         TabIndex        =   72
         Top             =   10920
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan 1.00 jika tiada tetapan yang diberikan oleh pihak kilang / supplier (Penerima)"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   71
         Top             =   10680
         Width           =   6945
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Berat bersih setelah ditukar kepada mutu 999.9 ialah "
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   70
         Top             =   10920
         Width           =   6945
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Mutu :                              ** Mutu yang ditentukan oleh pihak kilang / supplier"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   69
         Top             =   10470
         Width           =   6945
      End
      Begin VB.Label L28_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L28_Text"
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
         TabIndex        =   66
         Top             =   9840
         Width           =   615
      End
      Begin VB.Label L29_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L29_Text"
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
         Left            =   3600
         TabIndex        =   65
         Top             =   9840
         Width           =   1455
      End
      Begin VB.Label L10_Text 
         BackColor       =   &H008080FF&
         Caption         =   "L10_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   14040
         TabIndex        =   46
         Top             =   8040
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Butiran barang yang akan dihantar kepada supplier / kilang."
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
         TabIndex        =   44
         Top             =   2280
         Width           =   6615
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Mutu                                                 * Peratusan pertukaran mutu kepada mutu 999.9"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   43
         Top             =   3435
         Width           =   6945
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Purity"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   41
         Top             =   3105
         Width           =   1695
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Description "
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   39
         Top             =   2550
         Width           =   1665
      End
      Begin VB.Label L11_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L11_Text"
         ForeColor       =   &H00000000&
         Height          =   1140
         Left            =   1400
         TabIndex        =   37
         Top             =   1080
         Width           =   5655
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat "
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan maklumat supplier / kilang dan barang yang akan dihantar."
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
         TabIndex        =   35
         Top             =   480
         Width           =   6615
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier / Kilang"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label L14_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L14_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   8205
         TabIndex        =   31
         Top             =   9840
         Width           =   615
      End
      Begin VB.Label L13_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L13_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7680
         TabIndex        =   30
         Top             =   9840
         Width           =   375
      End
      Begin VB.Label L15_Text 
         Caption         =   "L15_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   13800
         TabIndex        =   29
         Top             =   6960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L16_Text 
         Caption         =   "L16_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   13800
         TabIndex        =   28
         Top             =   7200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai maklumat yang akan dihantar ke supplier / kilang."
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   240
         TabIndex        =   27
         Top             =   4560
         Width           =   6855
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Paparan Muka  :          / "
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   6480
         TabIndex        =   32
         Top             =   9840
         Width           =   2295
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan :                     Jumlah berat (g) :"
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
         TabIndex        =   67
         Top             =   9840
         Width           =   3615
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Modal : RM"
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
         TabIndex        =   149
         Top             =   10080
         Width           =   3615
      End
   End
   Begin VB.PictureBox Pic4 
      BorderStyle     =   0  'None
      Height          =   11055
      Left            =   360
      ScaleHeight     =   11055
      ScaleWidth      =   23205
      TabIndex        =   81
      Top             =   0
      Visible         =   0   'False
      Width           =   23205
      Begin VB.PictureBox Pic5 
         BorderStyle     =   0  'None
         Height          =   10815
         Left            =   6720
         ScaleHeight     =   10815
         ScaleWidth      =   14445
         TabIndex        =   92
         Top             =   0
         Visible         =   0   'False
         Width           =   14445
         Begin VB.CommandButton CMD10 
            BackColor       =   &H00FFFFFF&
            Height          =   740
            Left            =   13320
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm107.frx":398B
            MousePointer    =   99  'Custom
            Picture         =   "Frm107.frx":3C95
            Style           =   1  'Graphical
            TabIndex        =   96
            ToolTipText     =   "Paparan seterusnya"
            Top             =   9960
            Width           =   1000
         End
         Begin VB.CommandButton CMD9 
            BackColor       =   &H00FFFFFF&
            Height          =   740
            Left            =   12240
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm107.frx":45BB
            MousePointer    =   99  'Custom
            Picture         =   "Frm107.frx":48C5
            Style           =   1  'Graphical
            TabIndex        =   95
            ToolTipText     =   "Paparan sebelumnya"
            Top             =   9960
            Width           =   1000
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
            Height          =   9525
            Left            =   120
            TabIndex        =   93
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
         Begin VB.Label L60_Text 
            BackStyle       =   0  'Transparent
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
            Height          =   300
            Left            =   6030
            TabIndex        =   150
            Top             =   9960
            Width           =   2415
         End
         Begin VB.Label L37_Text 
            BackStyle       =   0  'Transparent
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
            Height          =   300
            Left            =   3240
            TabIndex        =   103
            Top             =   9960
            Width           =   1335
         End
         Begin VB.Label L36_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L36_Text"
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
            TabIndex        =   102
            Top             =   9960
            Width           =   615
         End
         Begin VB.Label L35_Text 
            BackColor       =   &H8000000C&
            Caption         =   "L35_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   8880
            TabIndex        =   101
            Top             =   10200
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label L34_Text 
            BackColor       =   &H8000000C&
            Caption         =   "L34_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   8880
            TabIndex        =   100
            Top             =   9960
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label L32_Text 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   11340
            TabIndex        =   98
            Top             =   9960
            Width           =   375
         End
         Begin VB.Label L33_Text 
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
            Left            =   11850
            TabIndex        =   97
            Top             =   9960
            Width           =   615
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
            Height          =   300
            Left            =   240
            TabIndex        =   94
            Top             =   120
            Width           =   13455
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
            TabIndex        =   99
            Top             =   9960
            Width           =   2295
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "Bilangan :                     Jumlah berat :                           Jumlah Modal : RM  "
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
            TabIndex        =   104
            Top             =   9960
            Width           =   6615
         End
      End
      Begin VB.PictureBox Pic6 
         BorderStyle     =   0  'None
         Height          =   10815
         Left            =   12840
         ScaleHeight     =   10815
         ScaleWidth      =   14445
         TabIndex        =   113
         Top             =   1200
         Visible         =   0   'False
         Width           =   14445
         Begin VB.CommandButton CMD13 
            BackColor       =   &H000080FF&
            Caption         =   "Paparan Sebelum"
            Height          =   360
            Left            =   10080
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm107.frx":5204
            MousePointer    =   99  'Custom
            TabIndex        =   126
            Top             =   10320
            Width           =   1905
         End
         Begin VB.CommandButton CMD11 
            BackColor       =   &H00FFFFFF&
            Height          =   740
            Left            =   12240
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm107.frx":550E
            MousePointer    =   99  'Custom
            Picture         =   "Frm107.frx":5818
            Style           =   1  'Graphical
            TabIndex        =   115
            ToolTipText     =   "Paparan sebelumnya"
            Top             =   9960
            Width           =   1000
         End
         Begin VB.CommandButton CMD12 
            BackColor       =   &H00FFFFFF&
            Height          =   740
            Left            =   13320
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm107.frx":6157
            MousePointer    =   99  'Custom
            Picture         =   "Frm107.frx":6461
            Style           =   1  'Graphical
            TabIndex        =   114
            ToolTipText     =   "Paparan seterusnya"
            Top             =   9960
            Width           =   1000
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid5 
            Height          =   9525
            Left            =   120
            TabIndex        =   116
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
            Height          =   300
            Left            =   240
            TabIndex        =   123
            Top             =   120
            Width           =   13455
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
            Left            =   11850
            TabIndex        =   122
            Top             =   9960
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
            Left            =   11340
            TabIndex        =   121
            Top             =   9960
            Width           =   375
         End
         Begin VB.Label L46_Text 
            BackColor       =   &H8000000C&
            Caption         =   "L46_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   8880
            TabIndex        =   120
            Top             =   9960
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label L47_Text 
            BackColor       =   &H8000000C&
            Caption         =   "L47_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   8880
            TabIndex        =   119
            Top             =   10200
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label L48_Text 
            BackStyle       =   0  'Transparent
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
            Height          =   300
            Left            =   1080
            TabIndex        =   118
            Top             =   9960
            Width           =   615
         End
         Begin VB.Label L49_Text 
            BackStyle       =   0  'Transparent
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
            Height          =   300
            Left            =   3240
            TabIndex        =   117
            Top             =   9960
            Width           =   1335
         End
         Begin VB.Label Label39 
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
            TabIndex        =   124
            Top             =   9960
            Width           =   2295
         End
         Begin VB.Label Label40 
            BackStyle       =   0  'Transparent
            Caption         =   "Bilangan :                     Jumlah berat :"
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
            TabIndex        =   125
            Top             =   9960
            Width           =   3615
         End
      End
      Begin VB.PictureBox Pic7 
         BorderStyle     =   0  'None
         Height          =   10815
         Left            =   11040
         ScaleHeight     =   10815
         ScaleWidth      =   14445
         TabIndex        =   127
         Top             =   2160
         Visible         =   0   'False
         Width           =   14445
         Begin VB.CommandButton CMD15 
            BackColor       =   &H00FFFFFF&
            Height          =   740
            Left            =   13320
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm107.frx":6D87
            MousePointer    =   99  'Custom
            Picture         =   "Frm107.frx":7091
            Style           =   1  'Graphical
            TabIndex        =   130
            ToolTipText     =   "Paparan seterusnya"
            Top             =   9960
            Width           =   1000
         End
         Begin VB.CommandButton CMD14 
            BackColor       =   &H00FFFFFF&
            Height          =   740
            Left            =   12240
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm107.frx":79B7
            MousePointer    =   99  'Custom
            Picture         =   "Frm107.frx":7CC1
            Style           =   1  'Graphical
            TabIndex        =   129
            ToolTipText     =   "Paparan sebelumnya"
            Top             =   9960
            Width           =   1000
         End
         Begin VB.CommandButton CMD16 
            BackColor       =   &H000080FF&
            Caption         =   "Paparan Sebelum"
            Height          =   360
            Left            =   10080
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm107.frx":8600
            MousePointer    =   99  'Custom
            TabIndex        =   128
            Top             =   10320
            Width           =   1905
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid6 
            Height          =   9525
            Left            =   120
            TabIndex        =   131
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
            Left            =   3240
            TabIndex        =   138
            Top             =   9960
            Width           =   1335
         End
         Begin VB.Label L55_Text 
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
            Left            =   1080
            TabIndex        =   137
            Top             =   9960
            Width           =   615
         End
         Begin VB.Label L54_Text 
            BackColor       =   &H8000000C&
            Caption         =   "L54_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   8880
            TabIndex        =   136
            Top             =   10200
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label L53_Text 
            BackColor       =   &H8000000C&
            Caption         =   "L53_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   8880
            TabIndex        =   135
            Top             =   9960
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label L51_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Height          =   300
            Left            =   11340
            TabIndex        =   134
            Top             =   9960
            Width           =   375
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
            Left            =   11850
            TabIndex        =   133
            Top             =   9960
            Width           =   615
         End
         Begin VB.Label L50_Text 
            BackStyle       =   0  'Transparent
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
            Height          =   300
            Left            =   240
            TabIndex        =   132
            Top             =   120
            Width           =   13455
         End
         Begin VB.Label Label41 
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
            TabIndex        =   139
            Top             =   9960
            Width           =   2295
         End
         Begin VB.Label Label42 
            BackStyle       =   0  'Transparent
            Caption         =   "Bilangan :                     Jumlah berat :"
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
            TabIndex        =   140
            Top             =   9960
            Width           =   3615
         End
      End
      Begin VB.TextBox TB4 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1500
         TabIndex        =   106
         Text            =   "TB4"
         Top             =   3480
         Width           =   1620
      End
      Begin VB.CommandButton CMD17 
         BackColor       =   &H000080FF&
         Caption         =   "Carian Data"
         Height          =   360
         Left            =   3240
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm107.frx":890A
         MousePointer    =   99  'Custom
         TabIndex        =   105
         Top             =   3440
         Width           =   1425
      End
      Begin VB.CommandButton CMD8 
         BackColor       =   &H000080FF&
         Caption         =   "Report"
         Height          =   405
         Left            =   2400
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm107.frx":8C14
         MousePointer    =   99  'Custom
         TabIndex        =   91
         Top             =   2040
         Width           =   2025
      End
      Begin VB.ComboBox CBB6 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   315
         ItemData        =   "Frm107.frx":8F1E
         Left            =   1500
         List            =   "Frm107.frx":8F20
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   1540
         Width           =   4725
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
         Left            =   285
         TabIndex        =   82
         Top             =   135
         Width           =   200
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   360
         Left            =   1500
         TabIndex        =   83
         Top             =   675
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
         Format          =   416415744
         CurrentDate     =   41561
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   360
         Left            =   1500
         TabIndex        =   84
         Top             =   1035
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
         Format          =   416415744
         CurrentDate     =   41561
      End
      Begin VB.Shape Shape2 
         Height          =   1215
         Left            =   120
         Top             =   2760
         Width           =   6495
      End
      Begin VB.Shape Shape1 
         Height          =   2655
         Left            =   120
         Top             =   0
         Width           =   6495
      End
      Begin VB.Label L41_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L41_Text"
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
         Left            =   960
         TabIndex        =   112
         Top             =   7320
         Visible         =   0   'False
         Width           =   615
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
         Height          =   300
         Left            =   960
         TabIndex        =   111
         Top             =   6960
         Visible         =   0   'False
         Width           =   615
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
         Height          =   300
         Left            =   960
         TabIndex        =   110
         Top             =   6600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label L38_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L38_Text"
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
         Left            =   960
         TabIndex        =   109
         Top             =   5760
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan No. Rujukan bagi mencari data terperinci hantaran barang dari nombor rujukan tersebut."
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   240
         TabIndex        =   108
         Top             =   3000
         Width           =   4530
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Rujukan      :"
         Height          =   255
         Left            =   240
         TabIndex        =   107
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier / Kilang"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   240
         TabIndex        =   90
         Top             =   1580
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "** Jika tidak ditanda , sistem TIDAK akan mengeluarkan report mengikut tarikh."
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   600
         TabIndex        =   88
         Top             =   360
         Width           =   5850
      End
      Begin VB.Label Label62 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Mula "
         Height          =   255
         Left            =   240
         TabIndex        =   87
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Akhir "
         Height          =   255
         Left            =   240
         TabIndex        =   86
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila tandakan di sini jika ingin cari data mengikut tarikh."
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   570
         TabIndex        =   85
         Top             =   120
         Width           =   4890
      End
   End
   Begin VB.PictureBox Pic2 
      BorderStyle     =   0  'None
      Height          =   11895
      Left            =   360
      ScaleHeight     =   11895
      ScaleWidth      =   22695
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   22695
      Begin VB.CommandButton CMD6 
         BackColor       =   &H00FFFFFF&
         Height          =   740
         Left            =   14040
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm107.frx":8F22
         MousePointer    =   99  'Custom
         Picture         =   "Frm107.frx":922C
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Paparan sebelumnya"
         Top             =   10560
         Width           =   1000
      End
      Begin VB.CommandButton CDM7 
         BackColor       =   &H00FFFFFF&
         Height          =   740
         Left            =   15120
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm107.frx":9B6B
         MousePointer    =   99  'Custom
         Picture         =   "Frm107.frx":9E75
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Paparan seterusnya"
         Top             =   10560
         Width           =   1000
      End
      Begin VB.PictureBox Pic3 
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   2160
         ScaleHeight     =   1575
         ScaleWidth      =   7125
         TabIndex        =   14
         Top             =   1680
         Visible         =   0   'False
         Width           =   7118
         Begin VB.ComboBox CBB1 
            BackColor       =   &H00FFFFFF&
            DataField       =   "Supplier"
            Height          =   315
            ItemData        =   "Frm107.frx":A79B
            Left            =   1200
            List            =   "Frm107.frx":A79D
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   120
            Width           =   4365
         End
         Begin VB.ComboBox CBB2 
            BackColor       =   &H00FFFFFF&
            DataField       =   "Supplier"
            Height          =   315
            ItemData        =   "Frm107.frx":A79F
            Left            =   1200
            List            =   "Frm107.frx":A7A1
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   430
            Width           =   4365
         End
         Begin VB.CommandButton CMD1 
            BackColor       =   &H000080FF&
            Caption         =   "Paparan Senarai"
            Height          =   405
            Left            =   2040
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm107.frx":A7A3
            MousePointer    =   99  'Custom
            TabIndex        =   15
            Top             =   840
            Width           =   2025
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Status barang"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   0
            TabIndex        =   21
            Top             =   150
            Width           =   2295
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Purity"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   0
            TabIndex        =   20
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label L6_Text 
            BackColor       =   &H8000000A&
            Caption         =   "L6_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   120
            TabIndex        =   19
            Top             =   840
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label L7_Text 
            BackColor       =   &H8000000A&
            Caption         =   "L7_Text"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   120
            TabIndex        =   18
            Top             =   1200
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   735
         Left            =   21480
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CMD3 
         BackColor       =   &H00FFFFFF&
         Height          =   740
         Left            =   7680
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm107.frx":AAAD
         MousePointer    =   99  'Custom
         Picture         =   "Frm107.frx":ADB7
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Paparan seterusnya"
         Top             =   10560
         Width           =   1000
      End
      Begin VB.CommandButton CMD2 
         BackColor       =   &H00FFFFFF&
         Height          =   740
         Left            =   6600
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm107.frx":B6DD
         MousePointer    =   99  'Custom
         Picture         =   "Frm107.frx":B9E7
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Paparan sebelumnya"
         Top             =   10560
         Width           =   1000
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   9765
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   720
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   17224
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   9765
         Left            =   8760
         TabIndex        =   53
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   720
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   17224
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
      Begin VB.Label L58_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L58_Text"
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
         Left            =   9840
         TabIndex        =   146
         Top             =   10800
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Description :"
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
         Left            =   16440
         TabIndex        =   64
         Top             =   840
         Width           =   6855
      End
      Begin VB.Label L26_Text 
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
         Left            =   9720
         TabIndex        =   62
         Top             =   10560
         Width           =   615
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
         Left            =   11700
         TabIndex        =   61
         Top             =   10560
         Width           =   1335
      End
      Begin VB.Label L23_Text 
         BackStyle       =   0  'Transparent
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
         Height          =   300
         Left            =   13560
         TabIndex        =   59
         Top             =   11040
         Width           =   615
      End
      Begin VB.Label L22_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   300
         Left            =   13035
         TabIndex        =   58
         Top             =   11040
         Width           =   375
      End
      Begin VB.Label L24_Text 
         Caption         =   "L24_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9000
         TabIndex        =   57
         Top             =   11040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L25_Text 
         Caption         =   "L25_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9000
         TabIndex        =   56
         Top             =   11280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai barang yang telah dimasukkan ke dalam senarai."
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
         Left            =   8880
         TabIndex        =   52
         Top             =   480
         Width           =   6855
      End
      Begin VB.Label L21_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Kembali"
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
         Left            =   2520
         MouseIcon       =   "Frm107.frx":C326
         MousePointer    =   99  'Custom
         TabIndex        =   51
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label L20_Text 
         BackStyle       =   0  'Transparent
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
         Height          =   1260
         Left            =   16440
         TabIndex        =   50
         Top             =   1080
         Width           =   6240
      End
      Begin VB.Label L19_Text 
         BackColor       =   &H008080FF&
         Caption         =   "L19_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   6600
         TabIndex        =   49
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L18_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tetapan Paparan Barang Kemas"
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
         MouseIcon       =   "Frm107.frx":C630
         MousePointer    =   99  'Custom
         TabIndex        =   48
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label L9_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L9_Text"
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
         TabIndex        =   12
         Top             =   10560
         Width           =   1455
      End
      Begin VB.Label L8_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L8_Text"
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
         TabIndex        =   11
         Top             =   10560
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan :                     Jumlah berat :"
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
         TabIndex        =   10
         Top             =   10560
         Width           =   3615
      End
      Begin VB.Label L5_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L5_Text"
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
         TabIndex        =   9
         Top             =   480
         Width           =   6855
      End
      Begin VB.Label L4_Text 
         Caption         =   "L4_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   240
         TabIndex        =   8
         Top             =   11040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L3_Text 
         Caption         =   "L3_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   240
         TabIndex        =   7
         Top             =   10800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L1_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   300
         Left            =   5835
         TabIndex        =   5
         Top             =   11040
         Width           =   375
      End
      Begin VB.Label L2_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L2_Text"
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
         Left            =   6360
         TabIndex        =   4
         Top             =   11040
         Width           =   615
      End
      Begin VB.Label Label3 
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
         Left            =   4440
         TabIndex        =   6
         Top             =   11040
         Width           =   2295
      End
      Begin VB.Label Label9 
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
         Left            =   11640
         TabIndex        =   60
         Top             =   11040
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan :                Jumlah berat :"
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
         Left            =   8880
         TabIndex        =   63
         Top             =   10560
         Width           =   5895
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Modal : RM"
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
         Left            =   8880
         TabIndex        =   147
         Top             =   10800
         Width           =   5895
      End
   End
   Begin VB.Label L17_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai Hantaran Barang"
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
      MouseIcon       =   "Frm107.frx":C93A
      MousePointer    =   99  'Custom
      TabIndex        =   47
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label L12_Text 
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
      Left            =   0
      MouseIcon       =   "Frm107.frx":CC44
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   0
      Width           =   1815
   End
   Begin VB.Menu Frm107_PM_menu1 
      Caption         =   "Menu 1"
      Visible         =   0   'False
      Begin VB.Menu Frm107_SM_pilih 
         Caption         =   "Pilih barang ini"
      End
   End
   Begin VB.Menu Frm107_PM_menu2 
      Caption         =   "Menu 2"
      Visible         =   0   'False
      Begin VB.Menu Frm107_SM_edit_data_desc 
         Caption         =   "Edit data description"
      End
      Begin VB.Menu Frm107_SM_padam_desc 
         Caption         =   "Padam / keluarkan dari senarai"
      End
      Begin VB.Menu Frm107_SM_pilihan_barang 
         Caption         =   "Pilihan barang yang akan dihantar"
      End
   End
   Begin VB.Menu Frm107_PM_menu3 
      Caption         =   "Menu 3"
      Visible         =   0   'False
      Begin VB.Menu Frm107_SM_remove 
         Caption         =   "Keluarkan barang ini dari senarai"
      End
   End
   Begin VB.Menu Frm107_PM_menu4 
      Caption         =   "Menu 4"
      Visible         =   0   'False
      Begin VB.Menu Frm107_SM_edit_data 
         Caption         =   "Edit penyata ini"
      End
      Begin VB.Menu Frm107_SM_padam_data 
         Caption         =   "Padam penyata ini"
      End
      Begin VB.Menu Frm107_SM_cetak_statement 
         Caption         =   "Cetak statement ini"
      End
      Begin VB.Menu Frm107_SM_senarai_description 
         Caption         =   "Senarai maklumat terperinci dari statement ini."
      End
      Begin VB.Menu Frm107_SM_senarai_barang_excel 
         Caption         =   "Senarai barang yang dihantar (Excel)"
      End
   End
   Begin VB.Menu Frm107_PM_menu5 
      Caption         =   "Menu 5"
      Visible         =   0   'False
      Begin VB.Menu Frm107_SM_senarai_barang 
         Caption         =   "Senarai barang yang dihantar"
      End
   End
End
Attribute VB_Name = "Frm107"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CBB3_Change()
'on error resume next
If GLOBAL_DISABLE <> 1 Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from setting_database where Supplier='" & Frm107.CBB3 & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        Frm107.L11_Text = vbNullString
        
        If Not IsNull(rs!alamat) Then Frm107.L11_Text = rs!alamat 'Alamat

    End If
    
    rs.Close
    Set rs = Nothing

End If
End Sub
Private Sub CBB3_Click()
'on error resume next
If GLOBAL_DISABLE <> 1 Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from setting_database where Supplier='" & Frm107.CBB3 & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        Frm107.L11_Text = vbNullString
        
        If Not IsNull(rs!alamat) Then Frm107.L11_Text = rs!alamat 'Alamat

    End If
    
    rs.Close
    Set rs = Nothing

End If
End Sub
Private Sub CDM6_Click()
'on error resume next
Dim Frm107_LM_CURR_PAGE As Double
Dim Frm107_LM_TOTAL_PAGE As Double

Frm107_LM_CURR_PAGE = 0
Frm107_LM_TOTAL_PAGE = 0

If Frm107.L13_Text <> vbNullString And IsNumeric(Frm107.L13_Text) Then
    If Frm107.L14_Text <> vbNullString And IsNumeric(Frm107.L14_Text) Then
        Frm107_LM_CURR_PAGE = Frm107.L13_Text
        Frm107_LM_TOTAL_PAGE = Frm107.L14_Text
        
        If Frm107_LM_CURR_PAGE < Frm107_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm107_senarai_description_header
            Call Frm107_senarai_description
            
        End If
    End If
End If
End Sub
Private Sub CDM7_Click()
'on error resume next
Dim Frm107_LM_CURR_PAGE As Double
Dim Frm107_LM_TOTAL_PAGE As Double

Frm107_LM_CURR_PAGE = 0
Frm107_LM_TOTAL_PAGE = 0

If Frm107.L22_Text <> vbNullString And IsNumeric(Frm107.L22_Text) Then
    If Frm107.L23_Text <> vbNullString And IsNumeric(Frm107.L23_Text) Then
        Frm107_LM_CURR_PAGE = Frm107.L22_Text
        Frm107_LM_TOTAL_PAGE = Frm107.L23_Text
        
        If Frm107_LM_CURR_PAGE < Frm107_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm107_senarai_hantar_header
            Call Frm107_senarai_hantar
            
        End If
    End If
End If
End Sub
Private Sub CMD1_Click()
'on error resume next
If Frm107.CBB1 = vbNullString Then
    MsgBox "Sila pilih [Status barang].", vbInformation, "Info"

    Exit Sub
End If
If Frm107.CBB2 = vbNullString Then
    MsgBox "Sila pilih [Purity].", vbInformation, "Info"

    Exit Sub
End If

Frm107.L6_Text = Frm107.CBB1
Frm107.L7_Text = Frm107.CBB2

GM_NEXT_PREV = 0
Frm107.L3_Text = -1 'Titik Pencarian Data
Frm107.L4_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Frm107.L1_Text = 0 'Paparan Page ke-xxx

Call Frm107_senarai_barang_header
Call Frm107_senarai_barang

Frm107.Pic3.Visible = False
End Sub
Private Sub CMD10_Click()
'on error resume next
Dim Frm107_LM_CURR_PAGE As Double
Dim Frm107_LM_TOTAL_PAGE As Double

Frm107_LM_CURR_PAGE = 0
Frm107_LM_TOTAL_PAGE = 0

If Frm107.L32_Text <> vbNullString And IsNumeric(Frm107.L32_Text) Then
    If Frm107.L33_Text <> vbNullString And IsNumeric(Frm107.L33_Text) Then
        Frm107_LM_CURR_PAGE = Frm107.L32_Text
        Frm107_LM_TOTAL_PAGE = Frm107.L33_Text
        
        If Frm107_LM_CURR_PAGE < Frm107_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm107_report_statement_header
            Call Frm107_report_statement
            
        End If
    End If
End If
End Sub
Private Sub CMD11_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm107_report_description_header
Call Frm107_report_description
End Sub
Private Sub CMD12_Click()
'on error resume next
Dim Frm107_LM_CURR_PAGE As Double
Dim Frm107_LM_TOTAL_PAGE As Double

Frm107_LM_CURR_PAGE = 0
Frm107_LM_TOTAL_PAGE = 0

If Frm107.L44_Text <> vbNullString And IsNumeric(Frm107.L44_Text) Then
    If Frm107.L45_Text <> vbNullString And IsNumeric(Frm107.L45_Text) Then
        Frm107_LM_CURR_PAGE = Frm107.L44_Text
        Frm107_LM_TOTAL_PAGE = Frm107.L45_Text
        
        If Frm107_LM_CURR_PAGE < Frm107_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm107_report_description_header
            Call Frm107_report_description
            
        End If
    End If
End If
End Sub
Private Sub CMD13_Click()
'On Error Resume Next
Note = "Adakah anda ingin tutup paparan ini?" & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then
    Frm107.Pic6.Visible = False
    Frm107.Pic5.Visible = True
End If
End Sub
Private Sub CMD14_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm107_report_senarai_hantar_header
Call Frm107_report_senarai_hantar
End Sub
Private Sub CMD15_Click()
'on error resume next
Dim Frm107_LM_CURR_PAGE As Double
Dim Frm107_LM_TOTAL_PAGE As Double

Frm107_LM_CURR_PAGE = 0
Frm107_LM_TOTAL_PAGE = 0

If Frm107.L51_Text <> vbNullString And IsNumeric(Frm107.L51_Text) Then
    If Frm107.L52_Text <> vbNullString And IsNumeric(Frm107.L52_Text) Then
        Frm107_LM_CURR_PAGE = Frm107.L51_Text
        Frm107_LM_TOTAL_PAGE = Frm107.L52_Text
        
        If Frm107_LM_CURR_PAGE < Frm107_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm107_report_senarai_hantar_header
            Call Frm107_report_senarai_hantar
            
        End If
    End If
End If
End Sub
Private Sub CMD16_Click()
'On Error Resume Next
Note = "Adakah anda ingin tutup paparan ini?" & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then
    Frm107.Pic7.Visible = False
    Frm107.Pic6.Visible = True
End If
End Sub
Private Sub CMD17_Click()
'on error resume next
If Frm107.TB4 = vbNullString Then
    MsgBox "Sila masukkan No. Rujukan.", vbInformation, "Info"
    
    Frm107.TB4.SetFocus
    Exit Sub
End If

If InStr(1, Frm107.TB4, "*") <> 0 Or InStr(1, Frm107.TB4, "/") <> 0 Or InStr(1, Frm107.TB4, "\") <> 0 Or InStr(1, Frm107.TB4, "'") <> 0 Then
    MsgBox "Simbol tidak dibenarkan di dalam ruangan No. Rujukan.", vbExclamation, "Error"
    
    Frm107.TB4 = vbNullString
    Frm107.TB4.SetFocus
    Exit Sub
End If

Note = "Sistem akan mengambil masa untuk mengeluarkan report." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then

    Call Frm107_initial_location3
    
    Frm107.L38_Text = 2 '0 : Carian ikut supplier sahaja , 1 : Carian ikut tarikh , 2 : Carian ikut No. Rujukan
    
    Frm107.L41_Text = UCase(Frm107.TB4) 'No. Rujukan
    
    GM_NEXT_PREV = 0
    Frm107.L34_Text = -1 'Titik Pencarian Data
    Frm107.L35_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm107.L32_Text = 0 'Paparan Page ke-xxx
    
    Call Frm107_report_statement_header
    Call Frm107_report_statement
    
    Frm107.Pic5.Visible = True
    
    If Frm107.L36_Text <> vbNullString Then
        If Frm107.L36_Text = 0 Then MsgBox "Tiada rekod dijumpai.", vbInformation, "Info"
    End If
    
End If
End Sub
Private Sub CMD18_Click()
'on error resume next
Dim Err(10)
Dim Frm107_LM_ID_DESC As Integer
Dim Frm107_LM_ID_FORM As Integer
Dim Frm107_LM_BERAT_BEFORE As Double
Dim Frm107_LM_MUTU As Double
Dim Frm107_LM_BERAT_AFTER As Double
Dim LM_MUTU As Double

Frm107_LM_ID_DESC = 1
Frm107_LM_ID_FORM = 1
LM_MUTU = 0
Frm107_LM_BERAT_BEFORE = 0
Frm107_LM_MUTU = 0
Frm107_LM_BERAT_AFTER = 0

If Frm107.L31_Text = vbNullString Or ((Frm107.L31_Text <> vbNullString) And Not IsNumeric(Frm107.L31_Text)) Then
    x = x + 1
    Err(x) = "Tiada maklumat ID bagi description ini. Sila keluar dari edit edit data ini dah cuba sekali lagi."
End If
If Frm107.TB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan [Description]."
End If
If Frm107.CBB4 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Purity]."
End If
If Frm107.TB2 = vbNullString Or (Frm107.TB2 <> vbNullString And Not IsNumeric(Frm107.TB2)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Mutu]. Hanya NOMBOR dibenarkan di dalam ruangan ini."
End If
If (Frm107.TB2 <> vbNullString And IsNumeric(Frm107.TB2)) Then
    
    LM_MUTU = Frm107.TB2
    
    If LM_MUTU > 1 Then
    
        x = x + 1
        Err(x) = "[Mutu] tidak boleh lebih dari 1."
        
    End If
    
End If
If Frm107.TB2 <> vbNullString And IsNumeric(Frm107.TB2) Then
    If Len(Frm107.TB2) > 20 Then
        x = x + 1
        Err(x) = "Hanya 20 digit dibenarkan dalam ruangan [Mutu]."
    End If
End If
If Frm107.L31_Text = vbNullString Or (Frm107.L31_Text <> vbNullString And Not IsNumeric(Frm107.L31_Text)) Then
    x = x + 1
    Err(x) = "[Technical Error : ID Description] Sila keluar dari menu ini dan cuba sekali lagi."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah anda ingin masukkan description yang telah diedit ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then
    
        '### Periksa nombor rujukan desc ### - Start
        Frm107_LM_ID_DESC = Frm107.L31_Text
        Frm107_LM_ID_FORM = Frm107.L10_Text
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_FORM_OUT_DESC & " where ID='" & Frm107_LM_ID_DESC & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Frm107.L10_Text <> vbNullString Then 'No rujukan sistem
                rs!no_rujukan = Frm107.L10_Text
            Else
                rs!no_rujukan = 0
            End If
            rs!id_desc = Frm107.L31_Text 'No. rujukan ID description (Running numnber) (Auto generated number)
            If Frm107.TB1 <> vbNullString Then 'Description
                rs!Description = Frm107.TB1
            Else
                rs!Description = Null
            End If
            If Frm107.CBB4 <> vbNullString Then 'Purity
                rs!purity = Frm107.CBB4
            Else
                rs!purity = Null
            End If
            If Not IsNull(rs!berat_before) Then
                If IsNumeric(rs!berat_before) Then Frm107_LM_BERAT_BEFORE = rs!berat_before 'Jumlah berat keseluruhan sebelum ditukar purity
            End If
            If Frm107.TB2 <> vbNullString Then 'Mutu (Kadar tukaran kepada mutu 999.9)
                rs!Conversion = Frm107.TB2
                Frm107_LM_MUTU = Frm107.TB2
            Else
                rs!Conversion = Null
            End If
            rs!berat_after = Format(Frm107_LM_BERAT_BEFORE * Frm107_LM_MUTU, "0.00") 'Jumlah berat selepas ditukar purity
            
            If Frm107.L57_Text = 0 Then '0 : Data baru , 1:  Data Edit
                rs!Status = 1 '0 : Batal , 1 : Aktif , 2 : Edit , 3 : Data baru (menu edit) , 4 : Padam (menu edit)
            ElseIf Frm107.L57_Text = 1 Then
                If rs!Status = 1 Then
                    rs!Status = 2 '0 : Batal , 1 : Aktif , 2 : Edit , 3 : Data baru (menu edit) , 4 : Padam (menu edit)
                End If
            End If
        
            rs.Update
        
        End If
        
        rs.Close
        Set rs = Nothing
        
        Frm107.TB1 = vbNullString
        Frm107.TB2 = vbNullString
        
        Call Frm107_visible_component_1
        
        GM_NEXT_PREV = 0
        Frm107.L15_Text = -1 'Titik Pencarian Data
        Frm107.L16_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
        Frm107.L13_Text = 0 'Paparan Page ke-xxx

        Call Frm107_senarai_description_header
        Call Frm107_senarai_description
        
        MsgBox "Data telah berjaya dimasukkan ke dalam senarai.", vbInformation, "Info"
        
        Frm107.TB1.SetFocus
    End If
    
End If
End Sub
Private Sub CMD19_Click()
'on error resume next
Note = "Adakah anda ingin batalkan edit data ini?" & vbCrLf & _
        "Sistem tidak akan menyimpan data jika terdapat data yang diubah." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then

    Call Frm107_visible_component_1
    Frm107.TB1 = vbNullString
    Frm107.TB2 = vbNullString
    Frm107.TB1.SetFocus

End If
End Sub
Private Sub CMD2_Click()
'on error resume next
If Frm107.L5_Text <> vbNullString Then

    GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
    
    Call Frm107_senarai_barang_header
    Call Frm107_senarai_barang
    
End If
End Sub
Private Sub CMD20_Click()
'on error resume next
Dim Err(10)
Dim Frm107_LM_MUTU As Double
Dim Frm107_LM_BERAT As Double
Dim Frm107_LM_BERAT_AFTER As Double
Dim Frm107_LM_NO_RUJ_SISTEM As Double
Dim LM_MUTU As Double

Frm107_LM_ID_SUPPLIER = 0
Frm107_LM_MUTU = 0
Frm107_LM_BERAT = 0
Frm107_LM_BERAT_AFTER = 0
Frm107_LM_NO_RUJ_SISTEM = 0
LM_MUTU = 0
Frm107_LM_NO_RUJ_SISTEM = Frm107.L10_Text

If Frm107.L28_Text = 0 Then
    x = x + 1
    Err(x) = "Tiada senarai barang yang akan dihantar."
End If
If Frm107.TB3 = vbNullString Or (Frm107.TB3 <> vbNullString And Not IsNumeric(Frm107.TB3)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Mutu]. Hanya NOMBOR dibenarkan di dalam ruangan ini."
End If
If (Frm107.TB3 <> vbNullString And IsNumeric(Frm107.TB3)) Then
    
    LM_MUTU = Frm107.TB3
    
    If LM_MUTU > 1 Then
    
        x = x + 1
        Err(x) = "[Mutu] tidak boleh lebih dari 1."
        
    End If
    
End If
If Frm107.L29_Text = vbNullString Or (Frm107.L29_Text <> vbNullString And Not IsNumeric(Frm107.L29_Text)) Then
    x = x + 1
    Err(x) = "Tiada maklumat Jumlah Berat keseluruhan. Sila keluar dari menu ini dan cuba sekali lagi."
End If
If Frm107.CBB3 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Supplier / Kilang]."
End If
If Frm107.CBB5 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Nama Pekerja]."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah anda ingin simpan data hantaran barang ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then
        
        Frm107_LM_RUJUKAN = Frm107.L10_Text
        
        G_No_STATMENT_FORM = vbNullString
        G_No_STATMENT_FORM = Format(Frm107.L10_Text, "000000")
        
        '### ID database bagi supplier / kilang ini ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where Supplier='" & Frm107.CBB3 & "'", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            If Not IsNull(rs!ID) Then Frm107_LM_ID_SUPPLIER = rs!ID
        End If
        
        rs.Close
        Set rs = Nothing
        '### ID database bagi supplier / kilang ini ### - End
        
        '### Pengiraan berat selepas ditukar mutu ### - Start
        If (Frm107.L29_Text <> vbNullString And IsNumeric(Frm107.L29_Text)) And (Frm107.TB3 <> vbNullString And IsNumeric(Frm107.TB3)) Then
            Frm107_LM_MUTU = Frm107.TB3
            Frm107_LM_BERAT = Frm107.L29_Text
            
            Frm107_LM_BERAT_AFTER = Frm107_LM_MUTU * Frm107_LM_BERAT
        End If
        '### Pengiraan berat selepas ditukar mutu ### - End
        
        '### No Rujukan pekerja ### - Start
        If Frm107.CBB5 <> vbNullString Then
            Frm107_LM_EMP_NO = Split(Frm107.CBB5, "  |  ")(1)
        End If
        '### No Rujukan pekerja ### - End
        
        '### Masukkan data ke dalam table 57_form_out ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 57_form_out where no_rujukan='" & Frm107.L10_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            rs!tarikh = Frm107.DTPicker1
            If Frm107.CBB3 <> vbNullString Then 'Nama syarikat / kilang yang dihantar barang-barang ini.
                rs!nama_kedai = Frm107.CBB3
            Else
                rs!nama_kedai = Null
            End If
            rs!id_kedai = Frm107_LM_ID_SUPPLIER 'Ambil dari ID dari table senarai supplier
            If Frm107.L29_Text <> vbNullString Then 'Jumlah berat keseluruhan sebelum ditukar purity
                rs!berat_before = Format(Frm107.L29_Text, "0.00")
            Else
                rs!berat_before = Null
            End If
            If Frm107.TB3 <> vbNullString Then 'Jumlah conversion yang digunakan (%) - Sama macam purity
                rs!Conversion = Frm107.TB3
            Else
                rs!Conversion = Null
            End If
            rs!berat_after = Format(Frm107_LM_BERAT_AFTER, "0.00") 'Jumlah berat selepas ditukar purity
            If Frm107.L59_Text <> vbNullString Then 'Jumlah modal
                rs!modal = Format(Frm107.L59_Text, "0.00")
            Else
                rs!modal = Null
            End If
            rs!Status = 1
            rs!nama_pekerja = Frm107_LM_EMP_NO 'Nama pekerja yang masukkan data
            rs!write_timestamp2 = Now
            rs.Update
        
        End If
        
        rs.Close
        Set rs = Nothing
        '### Masukkan data ke dalam table 57_form_out ### - End
        
        
'@status bagi #60_form_out_list_temp
'1 : tiada perubahan (tidak perlu buat apa-apa dalam database)
'2 : data telah diedit (PERLU update dalam database)
'3 : tambahan data(description) baru (PERLU tambah dalam database)
'4 : dipadam (Keluarkan/padam dari database)
        
        '### Update status barang yang dihantar forming out di dalam table #58_form_out_list ### - Start (DATA EDIT)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE 58_form_out_list," & G_FORM_OUT_DESC & " SET 58_form_out_list.no_rujukan = " & G_FORM_OUT_DESC & ".no_rujukan ," _
        & "58_form_out_list.id_desc = " & G_FORM_OUT_DESC & ".id_desc ," _
        & "58_form_out_list.description = " & G_FORM_OUT_DESC & ".description ," _
        & "58_form_out_list.berat_before = " & G_FORM_OUT_DESC & ".berat_before ," _
        & "58_form_out_list.purity = " & G_FORM_OUT_DESC & ".purity ," _
        & "58_form_out_list.conversion = " & G_FORM_OUT_DESC & ".conversion ," _
        & "58_form_out_list.berat_after = " & G_FORM_OUT_DESC & ".berat_after ," _
        & "58_form_out_list.modal = " & G_FORM_OUT_DESC & ".modal ," _
        & "58_form_out_list.status = 1 ," _
        & "58_form_out_list.write_timestamp2 = NOW() WHERE " & G_FORM_OUT_DESC & ".status = 2 AND 58_form_out_list.id_desc = " & G_FORM_OUT_DESC & ".id_desc AND 58_form_out_list.no_rujukan = " & G_FORM_OUT_DESC & ".no_rujukan"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
        '### Update status barang yang dihantar forming out di dalam table #58_form_out_list ### - End (DATA EDIT)
               
        '### Remove/padam status barang yang dihantar forming out di dalam table #58_form_out_list ### - Start (DATA PADAM)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE 58_form_out_list," & G_FORM_OUT_DESC & " SET 58_form_out_list.status = 0 ," _
        & "58_form_out_list.write_timestamp2 = NOW() WHERE " & G_FORM_OUT_DESC & ".status = 4 AND 58_form_out_list.id_desc = " & G_FORM_OUT_DESC & ".id_desc AND 58_form_out_list.no_rujukan = " & G_FORM_OUT_DESC & ".no_rujukan"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
        '### Update status barang yang dihantar forming out di dalam table #58_form_out_list ### - End (DATA PADAM)
        
        '### Pindah data dari table 60_form_out_list_temp -> 58_form_out_list ### - Start (DATA BARU)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "insert into 58_form_out_list(no_rujukan,id_desc,description,berat_before,purity,conversion,berat_after,status,write_timestamp,modal)" & _
                    "select no_rujukan,id_desc,description,berat_before,purity,conversion,berat_after,1,NOW(),modal from " & G_FORM_OUT_DESC & " WHERE status='" & 3 & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
        '### Pindah data dari table 60_form_out_list_temp -> 58_form_out_list ### - End (DATA BARU)
        
        'GoTo aaa:
        
        '### Pindah data dari table 61_form_out_item_list_temp -> 59_form_out_item_list ### - Start (DATA BARU)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "insert into 59_form_out_item_list(no_rujukan,id_rujukan,no_siri_produk,purity,berat,Status,status_asal,write_timestamp,modal,jenis_barang)" & _
                    "select no_rujukan,id_rujukan,no_siri_produk,purity,berat,1,status_asal,NOW(),modal,jenis_barang from " & G_FORM_LIST & " WHERE status='" & 3 & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
        '### Pindah data dari table 61_form_out_item_list_temp -> 59_form_out_item_list ### - End (DATA BARU)
        
        '### Pindah data dari table 61_form_out_item_list_temp -> 59_form_out_item_list ### - Start (DATA EDIT)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        'strsql = "insert into 59_form_out_item_list(no_rujukan,id_rujukan,no_siri_produk,purity,berat,Status,status_asal,write_timestamp)" & _
                    "select no_rujukan,id_rujukan,no_siri_produk,purity,berat,1,status_asal,NOW() from 61_form_out_item_list_temp WHERE status='" & 3 & "'"
        
        strsql = "UPDATE 59_form_out_item_list," & G_FORM_LIST & " SET 59_form_out_item_list.no_rujukan = " & G_FORM_LIST & ".no_rujukan ," _
        & "59_form_out_item_list.id_rujukan = " & G_FORM_LIST & ".id_rujukan ," _
        & "59_form_out_item_list.no_siri_produk = " & G_FORM_LIST & ".no_siri_produk ," _
        & "59_form_out_item_list.purity = " & G_FORM_LIST & ".purity ," _
        & "59_form_out_item_list.berat = " & G_FORM_LIST & ".berat ," _
        & "59_form_out_item_list.modal = " & G_FORM_LIST & ".modal ," _
        & "59_form_out_item_list.jenis_barang = " & G_FORM_LIST & ".jenis_barang ," _
        & "59_form_out_item_list.Status = 1 ," _
        & "59_form_out_item_list.status_asal = " & G_FORM_LIST & ".status_asal ," _
        & "59_form_out_item_list.write_timestamp2 = NOW() WHERE " & G_FORM_LIST & ".status = 2 AND 59_form_out_item_list.no_rujukan = " & G_FORM_LIST & ".no_rujukan AND 59_form_out_item_list.no_siri_produk = " & G_FORM_LIST & ".no_siri_produk"
      
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
        '### Pindah data dari table 61_form_out_item_list_temp -> 59_form_out_item_list ### - End  (DATA EDIT)
        
        '### Deactivekan status barang yang dipadamkan ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE 59_form_out_item_list," & G_FORM_LIST & " SET 59_form_out_item_list.status = 0 ," _
        & "59_form_out_item_list.write_timestamp2 = NOW() WHERE " & G_FORM_LIST & ".status = 4 AND 59_form_out_item_list.no_rujukan = " & G_FORM_LIST & ".no_rujukan AND 59_form_out_item_list.no_siri_produk = " & G_FORM_LIST & ".no_siri_produk"

        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
        '### Deactivekan status barang yang dipadamkan ### - End
        
        '### Update status barang yang dihantar forming out di dalam table #data_database ### - Start (JIKA STATUS ASAL ADALAH BARANG TRADE IN)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE Data_Database," & G_FORM_LIST & " SET Data_Database.StatusItem = 23 " _
        & "WHERE Data_Database.no_siri_produk = " & G_FORM_LIST & ".no_siri_produk AND " & G_FORM_LIST & ".status = 3 AND " & G_FORM_LIST & ".status_asal = 10"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
        '### Update status barang yang dihantar forming out di dalam table #data_database ### - End (JIKA STATUS ASAL ADALAH BARANG TRADE IN)
        
        '### Update status barang yang dihantar forming out di dalam table #data_database ### - Start (JIKA STATUS ASAL ADALAH BARANG POTONG)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE Data_Database," & G_FORM_LIST & " SET Data_Database.StatusItem = 24 " _
        & "WHERE Data_Database.no_siri_produk = " & G_FORM_LIST & ".no_siri_produk AND " & G_FORM_LIST & ".status = 3 AND (" & G_FORM_LIST & ".status_asal = 12 OR " & G_FORM_LIST & ".status_asal = 20 OR " & G_FORM_LIST & ".status_asal = 22)"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
        '### Update status barang yang dihantar forming out di dalam table #data_database ### - End (JIKA STATUS ASAL ADALAH BARANG POTONG)
        
        '### Pulangkan status barang dalam table #data_database jika data dipadamkan (keluar dari senarai) ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE Data_Database," & G_FORM_LIST & " SET Data_Database.StatusItem = " & G_FORM_LIST & ".status_asal " _
        & "WHERE Data_Database.no_siri_produk = " & G_FORM_LIST & ".no_siri_produk AND " & G_FORM_LIST & ".status = 4"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
        '### Pulangkan status barang dalam table #data_database jika data dipadamkan (keluar dari senarai) ### - End
        
        '#### Update Log Aktiviti Sistem #### - Start
        user = MDI_frm1.L3_Text
        
        LogAct_Memory = "[" & user & "] Edit data forming out. No. Rujukan [" & Format(Frm107.L10_Text, "000000") & "]."
        LogDate_Memory = DateTime.Date & " " & DateTime.Time$
        Call UpdateLog_Database
        '#### Update Log Aktiviti Sistem #### - End
        
        Call Frm107_initial_setting1
        Call Frm107_clear_status
        
        GM_NEXT_PREV = 2
        
        Call Frm107_report_statement_header
        Call Frm107_report_statement
        
        Frm107.Pic4.Visible = True
        Frm107.Pic1.Visible = False
        
        Note = "Data telah berjaya disimpan." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Adakah anda ingin cetak penyata hantaran barang ini?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
            If G_No_STATMENT_FORM <> vbNullString Then
                Call Frm107_cetak_penyata_forming
            End If
        End If

    End If
End If
End Sub
Private Sub CMD21_Click()
'on error resume next
Note = "Adakah anda ingin batalkan edit data ini?" & vbCrLf & _
        "Sistem tidak akan menyimpan data jika terdapat data yang diubah." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then

    Call Frm107_visible_component_2
    
    Frm107.Pic4.Visible = True
    Frm107.Pic1.Visible = False

End If
End Sub
Private Sub CMD3_Click()
'on error resume next
Dim Frm107_LM_CURR_PAGE As Double
Dim Frm107_LM_TOTAL_PAGE As Double

Frm107_LM_CURR_PAGE = 0
Frm107_LM_TOTAL_PAGE = 0

If Frm107.L1_Text <> vbNullString And IsNumeric(Frm107.L1_Text) Then
    If Frm107.L2_Text <> vbNullString And IsNumeric(Frm107.L2_Text) Then
        Frm107_LM_CURR_PAGE = Frm107.L1_Text
        Frm107_LM_TOTAL_PAGE = Frm107.L2_Text
        
        If Frm107_LM_CURR_PAGE < Frm107_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm107_senarai_barang_header
            Call Frm107_senarai_barang
            
        End If
    End If
End If
End Sub
Private Sub CMD4_Click()
'on error resume next
Dim Err(10)
Dim Frm107_LM_ID_DESC As Integer
Dim Frm107_LM_ID_FORM As Integer
Dim LM_MUTU As Double

Frm107_LM_ID_DESC = 1
Frm107_LM_ID_FORM = 1
LM_MUTU = 0

If Frm107.TB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan [Description]."
End If
If Frm107.CBB4 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Purity]."
End If
If Frm107.TB2 = vbNullString Or (Frm107.TB2 <> vbNullString And Not IsNumeric(Frm107.TB2)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Mutu]. Hanya NOMBOR dibenarkan di dalam ruangan ini."
End If

If (Frm107.TB2 <> vbNullString And IsNumeric(Frm107.TB2)) Then
    
    LM_MUTU = Frm107.TB2
    
    If LM_MUTU > 1 Then
    
        x = x + 1
        Err(x) = "[Mutu] tidak boleh lebih dari 1."
        
    End If
    
End If

If Frm107.TB2 <> vbNullString And IsNumeric(Frm107.TB2) Then
    If Len(Frm107.TB2) > 20 Then
        x = x + 1
        Err(x) = "Hanya 20 digit dibenarkan dalam ruangan [Mutu]."
    End If
End If
If Frm107.L31_Text = vbNullString Or (Frm107.L31_Text <> vbNullString And Not IsNumeric(Frm107.L31_Text)) Then
    x = x + 1
    Err(x) = "[Technical Error : ID Description] Sila keluar dari menu ini dan cuba sekali lagi."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah anda ingin masukkan description ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then
    
        '### Periksa nombor rujukan desc ### - Start
        Frm107_LM_ID_DESC = Frm107.L31_Text
        Frm107_LM_ID_FORM = Frm107.L10_Text
        
Re_Gen_No:
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_FORM_OUT_DESC & " where id_desc='" & Frm107_LM_ID_DESC & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            Frm107_LM_ID_DESC = Frm107_LM_ID_DESC + 1
            Frm107.L31_Text = Frm107_LM_ID_DESC
            
            rs.Close
            Set rs = Nothing
            
            GoTo Re_Gen_No:
        End If
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 59_form_out_item_list where id_rujukan='" & Frm107_LM_ID_DESC & "' AND no_rujukan='" & Frm107_LM_ID_FORM & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            Frm107_LM_ID_DESC = Frm107_LM_ID_DESC + 1
            Frm107.L31_Text = Frm107_LM_ID_DESC
            
            rs.Close
            Set rs = Nothing
            
            GoTo Re_Gen_No:
        End If
        
        rs.Close
        Set rs = Nothing
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_FORM_OUT_DESC & "", cn, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        If Frm107.L10_Text <> vbNullString Then 'No rujukan sistem
            rs!no_rujukan = Frm107.L10_Text
        Else
            rs!no_rujukan = 0
        End If
        rs!id_desc = Frm107.L31_Text 'No. rujukan ID description (Running numnber) (Auto generated number)
        If Frm107.TB1 <> vbNullString Then 'Description
            rs!Description = Frm107.TB1
        Else
            rs!Description = Null
        End If
        If Frm107.CBB4 <> vbNullString Then 'Purity
            rs!purity = Frm107.CBB4
        Else
            rs!purity = Null
        End If
        rs!berat_before = "0.00" 'Jumlah berat keseluruhan sebelum ditukar purity
        If Frm107.TB2 <> vbNullString Then 'Mutu (Kadar tukaran kepada mutu 999.9)
            rs!Conversion = Frm107.TB2
        Else
            rs!Conversion = Null
        End If
        rs!berat_after = "0.00" 'Jumlah berat selepas ditukar purity
        If Frm107.L57_Text = 0 Then '0 : Data baru , 1:  Data Edit
            rs!Status = 1 '0 : Batal , 1 : Aktif , 2 : Edit , 3 : Data baru (menu edit) , 4 : Padam (menu edit)
        ElseIf Frm107.L57_Text = 1 Then
            rs!Status = 3 '0 : Batal , 1 : Aktif , 2 : Edit , 3 : Data baru (menu edit) , 4 : Padam (menu edit)
        End If
        rs!modal = "0.00"
        rs.Update
        
        rs.Close
        Set rs = Nothing
        
        Frm107.TB1 = vbNullString
        Frm107.TB2 = vbNullString
        
        Call Frm107_visible_component_1
        
        GM_NEXT_PREV = 0
        Frm107.L15_Text = -1 'Titik Pencarian Data
        Frm107.L16_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
        Frm107.L13_Text = 0 'Paparan Page ke-xxx

        Call Frm107_senarai_description_header
        Call Frm107_senarai_description
        
        MsgBox "Data telah berjaya dimasukkan ke dalam senarai.", vbInformation, "Info"
        
        Frm107.TB1.SetFocus
    End If
    
End If
End Sub
Private Sub CMD5_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm107_senarai_description_header
Call Frm107_senarai_description
End Sub
Private Sub CMD6_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm107_senarai_hantar_header
Call Frm107_senarai_hantar
End Sub
Private Sub CMD7_Click()
'on error resume next
Dim Err(6)
Dim Frm107_LM_MUTU As Double
Dim Frm107_LM_BERAT As Double
Dim Frm107_LM_BERAT_AFTER As Double
Dim Frm107_LM_NO_RUJ_SISTEM As Double
Dim LM_MUTU As Double

Frm107_LM_ID_SUPPLIER = 0
Frm107_LM_MUTU = 0
Frm107_LM_BERAT = 0
Frm107_LM_BERAT_AFTER = 0
Frm107_LM_NO_RUJ_SISTEM = 0
LM_MUTU = 0

Frm107_LM_NO_RUJ_SISTEM = Frm107.L10_Text

If Frm107.L28_Text = 0 Then
    x = x + 1
    Err(x) = "Tiada senarai barang yang akan dihantar."
End If
If Frm107.TB3 = vbNullString Or (Frm107.TB3 <> vbNullString And Not IsNumeric(Frm107.TB3)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Mutu]. Hanya NOMBOR dibenarkan di dalam ruangan ini."
End If
If (Frm107.TB3 <> vbNullString And IsNumeric(Frm107.TB3)) Then
    
    LM_MUTU = Frm107.TB3
    
    If LM_MUTU > 1 Then
    
        x = x + 1
        Err(x) = "[Mutu] tidak boleh lebih dari 1."
        
    End If
    
End If
If Frm107.L29_Text = vbNullString Or (Frm107.L29_Text <> vbNullString And Not IsNumeric(Frm107.L29_Text)) Then
    x = x + 1
    Err(x) = "Tiada maklumat Jumlah Berat keseluruhan. Sila keluar dari menu ini dan cuba sekali lagi."
End If
If Frm107.CBB3 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Supplier / Kilang]."
End If
If Frm107.CBB5 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Nama Pekerja]."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah anda ingin simpan data hantaran barang ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then
        
        Frm107_LM_RUJUKAN = Frm107.L10_Text
        
        G_No_STATMENT_FORM = vbNullString
        G_No_STATMENT_FORM = Format(Frm107.L10_Text, "000000")
        
        '### ID database bagi supplier / kilang ini ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where Supplier='" & Frm107.CBB3 & "'", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            If Not IsNull(rs!ID) Then Frm107_LM_ID_SUPPLIER = rs!ID
        End If
        
        rs.Close
        Set rs = Nothing
        '### ID database bagi supplier / kilang ini ### - End
        
        '### Pengiraan berat selepas ditukar mutu ### - Start
        If (Frm107.L29_Text <> vbNullString And IsNumeric(Frm107.L29_Text)) And (Frm107.TB3 <> vbNullString And IsNumeric(Frm107.TB3)) Then
            Frm107_LM_MUTU = Frm107.TB3
            Frm107_LM_BERAT = Frm107.L29_Text
            
            Frm107_LM_BERAT_AFTER = Frm107_LM_MUTU * Frm107_LM_BERAT
        End If
        '### Pengiraan berat selepas ditukar mutu ### - End
        
        '### No Rujukan pekerja ### - Start
        If Frm107.CBB5 <> vbNullString Then
            Frm107_LM_EMP_NO = Split(Frm107.CBB5, "  |  ")(1)
        End If
        '### No Rujukan pekerja ### - End
        
        '### Masukkan data ke dalam table 57_form_out ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 57_form_out", cn, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        rs!tarikh = Frm107.DTPicker1
        rs!no_statement = Format(Frm107.L10_Text, "000000") 'No. rujukan sistem (No. penyata) - XXXXXX
        If Frm107.L10_Text <> vbNullString Then 'No rujukan yang akan dicatatkan pada "Bill of materials"
            rs!no_rujukan = Frm107.L10_Text
        Else
            rs!no_rujukan = Null
        End If
        If Frm107.CBB3 <> vbNullString Then 'Nama syarikat / kilang yang dihantar barang-barang ini.
            rs!nama_kedai = Frm107.CBB3
        Else
            rs!nama_kedai = Null
        End If
        rs!id_kedai = Frm107_LM_ID_SUPPLIER 'Ambil dari ID dari table senarai supplier
        If Frm107.L29_Text <> vbNullString Then 'Jumlah berat keseluruhan sebelum ditukar purity
            rs!berat_before = Format(Frm107.L29_Text, "0.00")
        Else
            rs!berat_before = Null
        End If
        If Frm107.TB3 <> vbNullString Then 'Jumlah conversion yang digunakan (%) - Sama macam purity
            rs!Conversion = Frm107.TB3
        Else
            rs!Conversion = Null
        End If
        If Frm107.L59_Text <> vbNullString Then 'Jumlah modal
            rs!modal = Format(Frm107.L59_Text, "0.00")
        Else
            rs!modal = Null
        End If
        rs!berat_after = Format(Frm107_LM_BERAT_AFTER, "0.00") 'Jumlah berat selepas ditukar purity
        rs!Status = 1
        rs!nama_pekerja = Frm107_LM_EMP_NO 'Nama pekerja yang masukkan data
        rs!write_timestamp = Now
        rs.Update
        
        rs.Close
        Set rs = Nothing
        '### Masukkan data ke dalam table 57_form_out ### - End
        
        '### Pindah data dari table 60_form_out_list_temp -> 58_form_out_list ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "insert into 58_form_out_list(no_rujukan,id_desc,description,berat_before,purity,conversion,berat_after,status,write_timestamp,modal)" & _
                    "select no_rujukan,id_desc,description,berat_before,purity,conversion,berat_after,1,NOW(),modal from " & G_FORM_OUT_DESC & " WHERE status='" & 1 & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
        '### Pindah data dari table 60_form_out_list_temp -> 58_form_out_list ### - End
        
        '### Pindah data dari table 61_form_out_item_list_temp -> 59_form_out_item_list ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "insert into 59_form_out_item_list(no_rujukan,id_rujukan,no_siri_produk,purity,berat,Status,status_asal,write_timestamp,modal,jenis_barang)" & _
                    "select no_rujukan,id_rujukan,no_siri_produk,purity,berat,1,status_asal,NOW(),modal,jenis_barang from " & G_FORM_LIST & " WHERE status='" & 1 & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
        '### Pindah data dari table 61_form_out_item_list_temp -> 59_form_out_item_list ### - End
        
        '### Update status barang yang dihantar forming out di dalam table #data_database ### - Start (JIKA STATUS ASAL ADALAH BARANG TRADE IN)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE Data_Database," & G_FORM_LIST & " SET Data_Database.StatusItem = 23 " _
        & "WHERE Data_Database.no_siri_produk = " & G_FORM_LIST & ".no_siri_produk AND " & G_FORM_LIST & ".status = 1 AND " & G_FORM_LIST & ".status_asal = 10"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
        '### Update status barang yang dihantar forming out di dalam table #data_database ### - End (JIKA STATUS ASAL ADALAH BARANG TRADE IN)
        
        '### Update status barang yang dihantar forming out di dalam table #data_database ### - Start (JIKA STATUS ASAL ADALAH BARANG POTONG)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE Data_Database," & G_FORM_LIST & " SET Data_Database.StatusItem = 24 " _
        & "WHERE Data_Database.no_siri_produk = " & G_FORM_LIST & ".no_siri_produk AND " & G_FORM_LIST & ".status = 1 AND (" & G_FORM_LIST & ".status_asal = 12 OR " & G_FORM_LIST & ".status_asal = 20 OR " & G_FORM_LIST & ".status_asal = 22)"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
        '### Update status barang yang dihantar forming out di dalam table #data_database ### - End (JIKA STATUS ASAL ADALAH BARANG POTONG)
        
        '#### Update Log Aktiviti Sistem #### - Start
        user = MDI_frm1.L3_Text
        
        LogAct_Memory = "[" & user & "] Forming out. No. Rujukan [" & Format(Frm107.L10_Text, "000000") & "]."
        LogDate_Memory = DateTime.Date & " " & DateTime.Time$
        Call UpdateLog_Database
        '#### Update Log Aktiviti Sistem #### - End
        
        '###Update No. rujukan sistem### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If rs!Default1 = "Default" Then
                rs!no_rujukan_form = Frm107_LM_NO_RUJ_SISTEM + 1 'No. rujukan sistem
                rs.Update
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        '###Update No. rujukan sistem### - End
        
        Call Frm107_initial_setting1
        Call Frm107_clear_status
        
        Note = "Data telah berjaya disimpan." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Adakah anda ingin cetak penyata hantaran barang ini?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
            If G_No_STATMENT_FORM <> vbNullString Then
                Call Frm107_cetak_penyata_forming
            End If
        End If

    End If
End If
End Sub

Private Sub CMD8_Click()
'on error resume next
If Frm107.CBB6 = vbNullString Then
    MsgBox "Sila pilih Supplier/Kilang", vbInformation, "Info"
    
    Exit Sub
End If

Note = "Sistem mungkin akan mengambil masa untuk mengeluarkan report." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then

    Call Frm107_initial_location3
    
    If Frm107.CB1 = 0 Then
        Frm107.L38_Text = 0 '0 : Carian ikut supplier sahaja , 1 : Carian ikut tarikh , 2 : Carian ikut No. Rujukan
    Else
        Frm107.L38_Text = 1 '0 : Carian ikut supplier sahaja , 1 : Carian ikut tarikh , 2 : Carian ikut No. Rujukan
    End If
    
    Frm107.L39_Text = Frm107.DTPicker2 'Tarikh mula
    Frm107.L40_Text = Frm107.DTPicker3 'Tarikh akhir
    Frm107.L41_Text = Frm107.CBB6 'Supplier
    
    GM_NEXT_PREV = 0
    Frm107.L34_Text = -1 'Titik Pencarian Data
    Frm107.L35_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm107.L32_Text = 0 'Paparan Page ke-xxx
    
    Call Frm107_report_statement_header
    Call Frm107_report_statement
    
    Frm107.Pic5.Visible = True
    
    If Frm107.L36_Text <> vbNullString Then
        If Frm107.L36_Text = 0 Then MsgBox "Tiada rekod dijumpai.", vbInformation, "Info"
    End If
    
End If
End Sub
Private Sub CMD9_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm107_report_statement_header
Call Frm107_report_statement
End Sub

Private Sub Form_Load()
'on error resume next
Call Frm107_initial_setting
Call Frm107_initial_location
Call Frm107_initial_location2

GLOBAL_DISABLE = 0
Frm107.L19_Text = 0
Frm107.DTPicker1 = DateTime.Date
Frm107.DTPicker2 = DateTime.Date
Frm107.DTPicker3 = DateTime.Date

Frm107.L13_Text = 0 'Senarai description : Paparan page
Frm107.L14_Text = 0 'Senarai description : Jumlah page
Frm107.L32_Text = 0 'Senarai statement (report) : Paparan page
Frm107.L33_Text = 0 'Senarai statement (report) : Jumlah page
Frm107.L44_Text = 0 'Senarai description (report) : Paparan page
Frm107.L45_Text = 0 'Senarai description (report) : Jumlah page
Frm107.L51_Text = 0 'Senarai barang yang dihantar (report) : Paparan page
Frm107.L52_Text = 0 'Senarai barang yang dihantar (report) : Jumlah page

Frm107.L38_Text = 0 '0 : Carian ikut supplier sahaja , 1 : Carian ikut tarikh , 2 : Carian ikut No. Rujukan
Frm107.L42_Text = "Senarai rekod hantaran tukaran barang dengan supplier / kilang." 'Header
Frm107.L43_Text = "Senarai maklumat terperinci hantaran barang kepada supplier / kilang." 'Header
Frm107.L50_Text = "Senarai maklumat barang yang dihantar." 'Header
End Sub
Private Sub Frm107_SM_cetak_statement_Click()
'on error resume next
Call Frm107_cetak_penyata_forming
End Sub
Private Sub Frm107_SM_edit_data_Click()
'on error resume next
Dim rs1 As ADODB.Recordset
Dim DATA_PEKERJA_FOUND As Integer

DATA_FOUND = 0
Frm107_LM_EMP_ID = vbNullString
Frm107_LM_ID_KEDAI = vbNullString

If IsNumeric(Frm107.MSFlexGrid4) Then
    Frm107_LM_ID = Frm107.MSFlexGrid4.TextMatrix(Frm107.MSFlexGrid4, 2) 'No. ID
    
    If Frm107_LM_ID <> vbNullString Then
        If IsNumeric(Frm107_LM_ID) Then
        
            Note = "Adakah anda ingin edit data ini?" & vbCrLf & _
                    "Sistem mungkin akan mengambil sedikit masa untuk mencari semua data bagi hantaran emas kepada supplier/kilang ini." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Teruskan?"
                    
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbYes Then

                Call Frm107_initial_location
                Call Frm107_initial_setting1
                Call Frm107_initial_setting2
                Call Frm107_initial_setting4
                Call Frm107_visible_component_1
                Call Frm107_visible_component_2
                Call Frm107_clear_status
                
                Frm107.TB3 = "1.00"
    
                GLOBAL_DISABLE = 1
                
                '### Carian no. rujukan sistem ### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 57_form_out where ID='" & Frm107_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Not IsNull(rs!no_rujukan) Then
                        Frm107_LM_NO_RUJUKAN = rs!no_rujukan
                        DATA_FOUND = 1
                    End If
                    If Not IsNull(rs!tarikh) Then Frm107.DTPicker1 = rs!tarikh
                    If Not IsNull(rs!no_rujukan) Then 'No rujukan yang akan dicatatkan pada "Bill of materials"
                        Frm107.L10_Text = rs!no_rujukan
                    Else
                        Frm107.L10_Text = 1
                    End If
                    If Not IsNull(rs!id_kedai) Then Frm107_LM_ID_KEDAI = rs!id_kedai 'Ambil dari ID dari table senarai supplier
                    
                    On Error GoTo Err_A:
                    If Not IsNull(rs!nama_kedai) Then 'Nama supplier / kilang
                        Frm107_LM_NAMA_KEDAI = rs!nama_kedai
                        Frm107.CBB3 = Frm107_LM_NAMA_KEDAI
                    End If
            
Restore_A:
                    'on error resume next

                    If Not IsNull(rs!berat_before) Then 'Jumlah berat keseluruhan sebelum ditukar purity
                        Frm107.L29_Text = Format(rs!berat_before, "0.00")
                    Else
                        Frm107.L29_Text = Format(0, "0.00")
                    End If
                    If Not IsNull(rs!Conversion) Then 'Jumlah conversion yang digunakan (%) - Sama macam purity
                        Frm107.TB3 = Format(rs!Conversion, "0.00")
                    Else
                        Frm107.TB3 = Format(1, "0.00")
                    End If
                    If Not IsNull(rs!modal) Then 'Modal
                        Frm107.L59_Text = Format(rs!modal, "0.00")
                    Else
                        Frm107.L59_Text = Format(1, "0.00")
                    End If
                    
                    If Not IsNull(rs!nama_pekerja) Then 'Nama pekerja yang masukkan data
                        Frm107_LM_EMP_ID = rs!nama_pekerja
                    Else
                        Frm107_LM_EMP_ID = vbNullString
                    End If
                   
                End If
                
                rs.Close
                Set rs = Nothing
                
                '### Carian alamat supplier / kilang ### - Start
                If Frm107_LM_ID_KEDAI <> vbNullString Then

                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from setting_database where ID='" & Frm107_LM_ID_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                        Frm107.L11_Text = vbNullString
                        
                        If Not IsNull(rs!alamat) Then Frm107.L11_Text = rs!alamat 'Alamat
                
                    End If
                    
                    rs.Close
                    Set rs = Nothing
    
                End If
                '### Carian alamat supplier / kilang ### - End
                
                If Frm107_LM_EMP_ID <> vbNullString Then
                
                    '### Carian Maklumat Penjual (Data Pekerja) ### - Start
                    DATA_PEKERJA_FOUND = 0
                    
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from employee where NoPekerja='" & Frm107_LM_EMP_ID & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                        Frm107_LM_MAKLUMAT_PEKERJA = rs!Samaran & "  |  " & rs!NoPekerja
                        DATA_PEKERJA_FOUND = 1
                    End If
                    
                    rs.Close
                    Set rs = Nothing
                    
                    If DATA_PEKERJA_FOUND = 1 Then
                        On Error GoTo Err_B:
                        Frm107.CBB5 = Frm107_LM_MAKLUMAT_PEKERJA
Restore_B:
                    End If
                    '### Carian Maklumat Penjual (Data Pekerja) ### - End
                    
                    'on error resume next
                    
                End If
                
                If DATA_FOUND = 1 Then
                
                    '### Pindah data dari table 58_form_out_list -> 60_form_out_list_temp ### - Start
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    
                    strsql = "insert into " & G_FORM_OUT_DESC & "(no_rujukan,id_desc,description,berat_before,purity,conversion,berat_after,status,modal)" & _
                                "select no_rujukan,id_desc,description,berat_before,purity,conversion,berat_after,1,modal from 58_form_out_list WHERE no_rujukan='" & Frm107_LM_NO_RUJUKAN & "' AND status='" & 1 & "'"
                    
                    Set rs = cn.Execute(strsql)
                    Set rs = Nothing
                    '### Pindah data dari table 58_form_out_list -> 60_form_out_list_temp ### - End
                    
                    '### Pindah data dari table 59_form_out_item_list -> 61_form_out_item_list_temp ### - Start
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    
                    strsql = "insert into " & G_FORM_LIST & "(no_rujukan,id_rujukan,no_siri_produk,purity,berat,Status,status_asal,modal,jenis_barang)" & _
                                "select no_rujukan,id_rujukan,no_siri_produk,purity,berat,1,status_asal,modal,jenis_barang from 59_form_out_item_list WHERE no_rujukan='" & Frm107_LM_NO_RUJUKAN & "' AND status='" & 1 & "'"
                    
                    Set rs = cn.Execute(strsql)
                    Set rs = Nothing
                    '### Pindah data dari table 59_form_out_item_list -> 61_form_out_item_list_temp ### - End
                    
                    '### Update semua status semua barang yang telah dipilih dalam table #data_database ### - Start
                    
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    
                    strsql = "UPDATE Data_Database,59_form_out_item_list SET Data_Database.form_out_status = 1 " _
                    & "WHERE Data_Database.no_siri_produk = 59_form_out_item_list.no_siri_produk AND 59_form_out_item_list.status='" & 1 & "' AND no_rujukan='" & Frm107_LM_NO_RUJUKAN & "'"
                    
                    Set rs = cn.Execute(strsql)
                    Set rs = Nothing
                    
                    '### Update semua status semua barang yang telah dipilih dalam table #data_database ### - End
                    
                    Frm107.L57_Text = 1 '0 : Data baru , 1:  Data Edit
                    
                    Frm107.CBB5.Enabled = True
                    Frm107.CBB5.BackColor = &HFFFFFF
                    
                    GLOBAL_DISABLE = 0

                    GM_NEXT_PREV = 0
                    Frm107.L15_Text = -1 'Titik Pencarian Data
                    Frm107.L16_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
                    Frm107.L13_Text = 0 'Paparan Page ke-xxx
                    
                    Call Frm107_senarai_description_header
                    Call Frm107_senarai_description
                    
                    Frm107.L24_Text = -1 'Titik Pencarian Data
                    Frm107.L25_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
                    Frm107.L22_Text = 0 'Paparan Page ke-xxx
                    Frm107.L23_Text = 0
                    
                    Call Frm107_senarai_hantar_header
                    Call Frm107_senarai_hantar
                    
                    Frm107.CMD7.Visible = False
                    Frm107.CMD20.Visible = True
                    Frm107.CMD21.Visible = True
                
                    Frm107.Pic1.Visible = True
                    Frm107.Pic4.Visible = False
                End If

            End If
        
        End If
    End If
End If

Exit Sub

Err_A:
Frm107.CBB3.AddItem Frm107_LM_NAMA_KEDAI
Frm107.CBB3 = Frm107_LM_NAMA_KEDAI
Resume Restore_A:

Exit Sub
Err_B:
Frm107.CBB5.AddItem Frm107_LM_MAKLUMAT_PEKERJA
Frm107.CBB5 = Frm107_LM_MAKLUMAT_PEKERJA
Resume Restore_B:
End Sub
Private Sub Frm107_SM_edit_data_desc_Click()
'on error resume next
DATA_FOUND = 0
Frm107_LM_NO_SIRI = vbNullString

If IsNumeric(Frm107.MSFlexGrid2) Then
    Frm107_LM_ID = Frm107.MSFlexGrid2.TextMatrix(Frm107.MSFlexGrid2, 2) 'No. ID
    
    If Frm107_LM_ID <> vbNullString Then
    
        Call Frm107_initial_setting2

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_FORM_OUT_DESC & " where ID='" & Frm107_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            GLOBAL_DISABLE = 1
            
            If Not IsNull(rs!no_rujukan) Then 'No rujukan sistem
                Frm107.L10_Text = rs!no_rujukan
            Else
                Frm107.L10_Text = rs!no_rujukan
            End If
            If Not IsNull(rs!Description) Then Frm107.TB1 = rs!Description 'Description
            If Not IsNull(rs!id_desc) Then 'No. rujukan ID description (Running numnber) (Auto generated number)
                Frm107.L31_Text = rs!id_desc
            Else
                Frm107.L31_Text = 1
            End If
            
            On Error GoTo Err_A:
            If Not IsNull(rs!purity) Then 'Purity
                Frm107_LM_PURITY = rs!purity
                Frm107.CBB4 = Frm107_LM_PURITY
            End If
    
Restore_A:
            'on error resume next
        
            If Not IsNull(rs!Conversion) Then 'Mutu (Kadar tukaran kepada mutu 999.9)
                Frm107.TB2 = rs!Conversion
            Else
                Frm107.TB2 = rs!Conversion
            End If
            
            DATA_FOUND = 1
            
        End If
        
        rs.Close
        Set rs = Nothing
        
        GLOBAL_DISABLE = 0
        
        If DATA_FOUND = 1 Then
    
            Frm107.CMD4.Visible = False
            Frm107.CMD18.Visible = True
            Frm107.CMD19.Visible = True
    
        End If
        
    End If
End If

Exit Sub

Err_A:
Frm107.CBB4.AddItem Frm107_LM_PURITY
Frm107.CBB4 = Frm107_LM_PURITY
Resume Restore_A:
End Sub
Private Sub Frm107_SM_padam_data_Click()
'on error resume next
Frm107_LM_NO_RUJUKAN = vbNullString
DATA_FOUND = 0

If IsNumeric(Frm107.MSFlexGrid4) Then
    Frm107_LM_ID = Frm107.MSFlexGrid4.TextMatrix(Frm107.MSFlexGrid4, 2) 'No. ID
    
    If Frm107_LM_ID <> vbNullString Then
        If IsNumeric(Frm107_LM_ID) Then
        
            Note = "Adakah anda ingin PADAM data ini?" & vbCrLf & _
                    "Sistem mungkin akan mengambil sedikit masa untuk MEMADAMKAN semua data berkenaan statement ini." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Teruskan?"
                    
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbYes Then
            
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 57_form_out where ID='" & Frm107_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Not IsNull(rs!no_statement) Then Frm107_LM_No_STATEMENT = rs!no_statement 'No. Statenebt
                    If Not IsNull(rs!no_rujukan) Then
                        Frm107_LM_NO_RUJUKAN = rs!no_rujukan
                        DATA_FOUND = 1
                        
                        rs!Status = 0
                        rs.Update
                    End If
                End If
                
                rs.Close
                Set rs = Nothing
                
                If DATA_FOUND = 1 Then
                
                
                    '### Tukar status semua data berkenaan dengan statement ini dalam table #58_form_out_list ### - Start
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    
                    strsql = "UPDATE 58_form_out_list SET status = 0 WHERE no_rujukan='" & Frm107_LM_NO_RUJUKAN & "'"
                    
                    Set rs = cn.Execute(strsql)
                    Set rs = Nothing
                    '### Tukar status semua data berkenaan dengan statement ini dalam table #58_form_out_list ### - End
                    
                    '### Tukar status semua data berkenaan dengan statement ini dalam table #data_database ### - Start (JIKA STATUS ASAL ADALAH BARANG TRADE IN)
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    
                    strsql = "UPDATE Data_Database,59_form_out_item_list SET Data_Database.StatusItem = 10 " _
                    & "WHERE Data_Database.no_siri_produk = 59_form_out_item_list.no_siri_produk AND 59_form_out_item_list.status = 1 AND 59_form_out_item_list.status_asal = 10 AND 59_form_out_item_list.no_rujukan='" & Frm107_LM_NO_RUJUKAN & "'"
                    
                    Set rs = cn.Execute(strsql)
                    Set rs = Nothing
                    '### Tukar status semua data berkenaan dengan statement ini dalam table #data_database ### - End (JIKA STATUS ASAL ADALAH BARANG TRADE IN)
                    
                    '### Tukar status semua data berkenaan dengan statement ini dalam table #data_database ### - Start (JIKA STATUS ASAL ADALAH BARANG POTONG)
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    
                    strsql = "UPDATE Data_Database,59_form_out_item_list SET Data_Database.StatusItem = 12 " _
                    & "WHERE Data_Database.no_siri_produk = 59_form_out_item_list.no_siri_produk AND 59_form_out_item_list.status = 1 AND (59_form_out_item_list.status_asal = 12 OR 59_form_out_item_list.status_asal = 20 OR 59_form_out_item_list.status_asal = 22) AND 59_form_out_item_list.no_rujukan='" & Frm107_LM_NO_RUJUKAN & "'"
                    
                    Set rs = cn.Execute(strsql)
                    Set rs = Nothing
                    '### Tukar status semua data berkenaan dengan statement ini dalam table #data_database ### - End (JIKA STATUS ASAL ADALAH BARANG POTONG)

                    
                    '### Tukar status semua data berkenaan dengan statement ini dalam table #58_form_out_list ### - Start
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    
                    strsql = "UPDATE 59_form_out_item_list SET status = 0 WHERE no_rujukan='" & Frm107_LM_NO_RUJUKAN & "'"
                    
                    Set rs = cn.Execute(strsql)
                    Set rs = Nothing
                    '### Tukar status semua data berkenaan dengan statement ini dalam table #58_form_out_list ### - End
                    
                    '#### Update Log Aktiviti Sistem #### - Start
                    user = MDI_frm1.L3_Text
                    
                    LogAct_Memory = "[" & user & "] Padam data hantaran barang ke supplier. No. Statement [" & Frm107_LM_No_STATEMENT & "]."
                    LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                    Call UpdateLog_Database
                    '#### Update Log Aktiviti Sistem #### - End
            
                    GM_NEXT_PREV = 2
                    
                    Call Frm107_report_statement_header
                    Call Frm107_report_statement
                    
                    If Frm107.L36_Text <> vbNullString Then
                        If Frm107.L36_Text = 0 Then MsgBox "Tiada rekod dijumpai.", vbInformation, "Info"
                    End If
                    
                    MsgBox "Data telah berjaya dipadamkan.", vbInformation, "Info"
                    
                End If
            
            End If
        
        End If
    
    End If

End If
End Sub
Private Sub Frm107_SM_padam_desc_Click()
'on error resume next
DATA_FOUND = 0
Frm107_LM_ID_DESC = vbNullString
Frm107_LM_STATUS = 0

If IsNumeric(Frm107.MSFlexGrid2) Then
    Frm107_LM_ID = Frm107.MSFlexGrid2.TextMatrix(Frm107.MSFlexGrid2, 2) 'No. ID
    
    If Frm107.L57_Text = vbNullString Or ((Frm107.L57_Text <> vbNullString) And Not IsNumeric(Frm107.L57_Text)) Then
        MsgBox "Sistem menghadapi masalah teknikal semasa cuba memadam data ini." & vbCrLf & _
                "Sila keluar dari menu ini dan cuba sekali lagi.", vbExclamation, "Info"
        
        Exit Sub
    End If
    
    If Frm107_LM_ID <> vbNullString Then
        
        Note = "Adakah anda ingin keluarkan/padam description ini?" & vbCrLf & _
                "Semua data / barang yang dipilih dari description ini akan dipulangkan ke dalam stok kedai semula." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from " & G_FORM_OUT_DESC & " where ID='" & Frm107_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!Status) Then
                    Frm107_LM_STATUS = rs!Status
                End If
            End If
            
            rs.Close
            Set rs = Nothing
            
            '### Padam / keluarkan senarai jika adalah data baru ### - Start
            If Frm107.L57_Text = 0 Or Frm107_LM_STATUS = 3 Then  '0 : Data baru , 1:  Data Edit
            
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from " & G_FORM_OUT_DESC & " where ID='" & Frm107_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Not IsNull(rs!id_desc) Then Frm107_LM_ID_DESC = rs!id_desc 'No. rujukan ID description (Running numnber) (Auto generated number)
                    
                    rs.Delete
                    
                    rs.Update
                    
                End If
                
                rs.Close
                Set rs = Nothing
                
                If Frm107_LM_ID_DESC <> vbNullString Then
                
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    
                    strsql = "UPDATE Data_Database," & G_FORM_LIST & " SET Data_Database.form_out_status = 0 " _
                    & "WHERE Data_Database.no_siri_produk = " & G_FORM_LIST & ".no_siri_produk AND " & G_FORM_LIST & ".id_rujukan='" & Frm107_LM_ID_DESC & "'"
                    
                    Set rs = cn.Execute(strsql)
                    Set rs = Nothing

                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    
                    strsql = "DELETE from " & G_FORM_LIST & " WHERE id_rujukan='" & Frm107_LM_ID_DESC & "'"
                    
                    Set rs = cn.Execute(strsql)
                    Set rs = Nothing
                
                End If
                
            End If
            '### Padam / keluarkan senarai jika adalah data baru ### - End
        
            If Frm107.L57_Text = 1 Then '0 : Data baru , 1:  Data Edit
            
                If Frm107_LM_STATUS = 1 Or Frm107_LM_STATUS = 2 Then
            
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from " & G_FORM_OUT_DESC & " where ID='" & Frm107_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                        If Not IsNull(rs!id_desc) Then Frm107_LM_ID_DESC = rs!id_desc 'No. rujukan ID description (Running numnber) (Auto generated number)
                        rs!Status = 4 '0 : Batal , 1 : Aktif , 2 : Edit , 3 : Data baru (menu edit) , 4 : Padam (menu edit)
                        rs.Update
                    
                    End If
                    
                    rs.Close
                    Set rs = Nothing
                
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    
                    strsql = "UPDATE Data_Database," & G_FORM_LIST & " SET Data_Database.form_out_status = 4 " _
                    & "WHERE Data_Database.no_siri_produk = " & G_FORM_LIST & ".no_siri_produk AND " & G_FORM_LIST & ".id_rujukan='" & Frm107_LM_ID_DESC & "'"
                    
                    Set rs = cn.Execute(strsql)
                    Set rs = Nothing
                
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    
                    strsql = "UPDATE " & G_FORM_LIST & " SET status = 4 " _
                    & "WHERE id_rujukan='" & Frm107_LM_ID_DESC & "'"
                    
                    Set rs = cn.Execute(strsql)
                    Set rs = Nothing
                    
                End If
            
            End If
            
            
            GM_NEXT_PREV = 2
    
            Call Frm107_senarai_description_header
            Call Frm107_senarai_description
        
        End If
    
    End If
End If
End Sub
Private Sub Frm107_SM_pilih_Click()
'on error resume next
Dim rs1 As ADODB.Recordset

DATA_FOUND = 0
Frm107_LM_NO_SIRI = vbNullString

If IsNumeric(Frm107.MSFlexGrid1) Then
    Frm107_LM_ID = Frm107.MSFlexGrid1.TextMatrix(Frm107.MSFlexGrid1, 2) 'No. ID
    
    If Frm107_LM_ID <> vbNullString Then
    
        Note = "Adakah anda ingin MASUKKAN barang ini ke dalam senarai?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from " & G_FORM_OUT_DESC & " where ID='" & Frm107.L31_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Frm107.L57_Text = 0 Then '0 : Data baru , 1:  Data Edit
                    rs!Status = 1 '0 : Batal , 1 : Aktif , 2 : Edit , 3 : Data baru (menu edit) , 4 : Padam (menu edit)
                ElseIf Frm107.L57_Text = 1 Then
                    If rs!Status = 1 Then
                        rs!Status = 2 '0 : Batal , 1 : Aktif , 2 : Edit , 3 : Data baru (menu edit) , 4 : Padam (menu edit)
                    End If
                End If
            
                rs.Update
            
            End If
            
            rs.Close
            Set rs = Nothing
            
            '### Tukar status barang yang dipilih kepada 1 ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from Data_Database where ID='" & Frm107_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!no_siri_Produk) Then Frm107_LM_NO_SIRI = rs!no_siri_Produk 'No. Siri Produk
                rs!form_out_status = 1
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
            '### Tukar status barang yang dipilih kepada 1 ### - End
            
            If Frm107_LM_NO_SIRI <> vbNullString Then
                '### Masukkan barang yang dipilih ke dalam temp table ### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from " & G_FORM_LIST & " where no_siri_produk='" & Frm107_LM_NO_SIRI & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                
                    If Frm107.L10_Text <> vbNullString Then
                        rs!no_rujukan = Frm107.L10_Text
                    Else
                        rs!no_rujukan = Null
                    End If
                    If Frm107.L19_Text <> vbNullString Then
                        rs!id_rujukan = Frm107.L19_Text
                    Else
                        rs!id_rujukan = Null
                    End If
                    If Frm107.L57_Text = 0 Then
                        rs!Status = 1
                    ElseIf Frm107.L57_Text = 1 Then
                        If rs!Status = 1 Or rs!Status = 4 Then
                            rs!Status = 2
                        ElseIf rs!Status = 0 Or rs!Status = 3 Then
                            rs!Status = 3
                        End If
                    End If
    
                    rs.Update
                    
                Else

                    Set rs1 = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs1.Open "select * from Data_Database where no_siri_produk='" & Frm107_LM_NO_SIRI & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs1.EOF Then
                    
                        rs.AddNew
                        
                        If Frm107.L10_Text <> vbNullString Then
                            rs!no_rujukan = Frm107.L10_Text
                        Else
                            rs!no_rujukan = Null
                        End If
                        If Frm107.L19_Text <> vbNullString Then
                            rs!id_rujukan = Frm107.L19_Text
                        Else
                            rs!id_rujukan = Null
                        End If

                        If Not IsNull(rs1!no_siri_Produk) Then rs!no_siri_Produk = rs1!no_siri_Produk 'No. siri produk
                        If Not IsNull(rs1!kod_Purity) Then rs!purity = rs1!kod_Purity 'Purity
                        If Not IsNull(rs1!beza_berat) Then rs!Berat = rs1!beza_berat 'Berat
                        If Not IsNull(rs1!StatusItem) Then rs!status_asal = rs1!StatusItem 'Status asal barang ini dari table data_database (sebelum dihantar forming out)
                        If Frm107.L57_Text = 0 Then
                            rs!Status = 1
                        ElseIf Frm107.L57_Text = 1 Then
                            rs!Status = 3
                        End If
                        
                        If Not IsNull(rs1!receiving_Status) Then
                            
                            If rs1!receiving_Status = "0" Or rs1!receiving_Status = "2" Or rs1!receiving_Status = "4" Or rs1!receiving_Status = "5" Or rs1!receiving_Status = "6" Or rs1!receiving_Status = "8" Then
                            
                                If Not IsNull(rs1!beza_berat) And Not IsNull(rs1!harga_Per_Gram_Item) Then
                                    rs!modal = Format(rs1!beza_berat * rs1!harga_Per_Gram_Item, "0.00") 'Modal
                                End If
                                
                            ElseIf rs1!receiving_Status = "1" Or rs1!receiving_Status = "3" Or rs1!receiving_Status = "7" Then
                    
                                If Not IsNull(rs1!harga_item) Then
                                    rs!modal = Format(rs1!harga_item, "0.00") 'Modal
                                End If
                                
                            End If
                            
                            If rs1!receiving_Status = "0" Or rs1!receiving_Status = "1" Or rs1!receiving_Status = "4" Then
                            
                                rs!jenis_barang = "Baru"
                                
                                If rs1!StatusItem = "12" Or rs1!StatusItem = "20" Or rs1!StatusItem = "22" Then
                                    
                                    rs!jenis_barang = "Baru - potong"
                                
                                End If
                                
                            End If
                            
                            If rs1!receiving_Status = "2" Or rs1!receiving_Status = "3" Or rs1!receiving_Status = "5" Or rs1!receiving_Status = "6" Or rs1!receiving_Status = "7" Or rs1!receiving_Status = "8" Then
                            
                                rs!jenis_barang = "Trade In"
                                
                                If rs1!StatusItem = "12" Or rs1!StatusItem = "20" Or rs1!StatusItem = "22" Then
                                    
                                    rs!jenis_barang = "Trade In - potong"
                                
                                End If
                                
                            End If
                            
                        Else
                        
                            rs!modal = Format(0, "0.00") 'Modal
                            
                        End If
                        
                        
                        
'0:  BK
'4:  gold Bar
'1:  Barang permata

'2 : Trade In BK
'3 : Trade In Barang Permata
'5 : Trade In Gold Bar
'6 : Emas terpakai BK
'7 : Emas terpakai permata
'8 : Emas terpakai gold bar

'12 : In Stock - Potong
'20 : Terjual Secara Ansuran - Potong (Jelas)
'22 : Terjual Secara Tempahan - Potong (Siap)

                        
                        rs.Update
                        
                    End If
                    
                    rs1.Close
                    Set rs1 = Nothing

                End If
                
                rs.Close
                Set rs = Nothing
                
                '### Masukkan barang yang dipilih ke dalam temp table ### - End
            End If

            
            GM_NEXT_PREV = 2
            
            Call Frm107_senarai_barang_header
            Call Frm107_senarai_barang '1
            
            Call Frm107_senarai_hantar_header
            Call Frm107_senarai_hantar
            
        End If
        
    End If
    
End If
End Sub
Private Sub Frm107_SM_pilihan_barang_Click()
'on error resume next
If IsNumeric(Frm107.MSFlexGrid2) Then
    Frm107_LM_ID = Frm107.MSFlexGrid2.TextMatrix(Frm107.MSFlexGrid2, 2) 'No. ID
    
    If Frm107_LM_ID <> vbNullString Then
        Frm107.L5_Text = vbNullString
        Call Frm107_senarai_barang_header
        Frm107.Pic1.Visible = False
        Frm107.Pic2.Visible = True

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_FORM_OUT_DESC & " where ID='" & Frm107_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then

            If Not IsNull(rs!id_desc) Then 'No. rujukan ID description (Running numnber) (Auto generated number)
                Frm107.L31_Text = rs!id_desc
            End If
            
        End If

        GM_NEXT_PREV = 0
        Frm107.L3_Text = -1 'Titik Pencarian Data
        Frm107.L4_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
        Frm107.L1_Text = 0 'Paparan Page ke-xxx
        
        Frm107.L24_Text = -1 'Titik Pencarian Data
        Frm107.L25_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
        Frm107.L22_Text = 0 'Paparan Page ke-xxx
        Frm107.L23_Text = 0
        
        Frm107.L12_Text.Visible = False
        Frm107.L17_Text.Visible = False
        
        Frm107.L2_Text = 0
        Frm107.L8_Text = 0
        Frm107.L9_Text = "0.00 g"
        
        Call Frm107_senarai_hantar_header
        Call Frm107_senarai_hantar
        
    End If
End If
End Sub
Private Sub Frm107_SM_remove_Click()
'on error resume next
Dim rs1 As ADODB.Recordset

DATA_FOUND = 0
Frm107_LM_NO_SIRI = vbNullString

If IsNumeric(Frm107.MSFlexGrid3) Then
    Frm107_LM_ID = Frm107.MSFlexGrid3.TextMatrix(Frm107.MSFlexGrid3, 2) 'No. ID
    
    If Frm107_LM_ID <> vbNullString Then
    
        Note = "Adakah anda ingin KELUARKAN barang ini ke dalam senarai?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from " & G_FORM_OUT_DESC & " where ID='" & Frm107.L31_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Frm107.L57_Text = 0 Then '0 : Data baru , 1:  Data Edit
                    rs!Status = 1 '0 : Batal , 1 : Aktif , 2 : Edit , 3 : Data baru (menu edit) , 4 : Padam (menu edit)
                ElseIf Frm107.L57_Text = 1 Then
                    If rs!Status = 1 Then
                        rs!Status = 2 '0 : Batal , 1 : Aktif , 2 : Edit , 3 : Data baru (menu edit) , 4 : Padam (menu edit)
                    End If
                End If
            
                rs.Update
            
            End If
            
            rs.Close
            Set rs = Nothing
            
            '### Tukar status barang yang dipilih kepada 0 ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from " & G_FORM_LIST & " where ID='" & Frm107_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!no_siri_Produk) Then Frm107_LM_NO_SIRI = rs!no_siri_Produk 'No. Siri Produk
                If Frm107.L57_Text = 0 Then
                    rs!Status = 0
                ElseIf Frm107.L57_Text = 1 Then
                    If rs!Status = 1 Or rs!Status = 2 Then
                        rs!Status = 4
                    ElseIf rs!Status = 3 Then
                        rs!Status = 0
                    End If
                End If
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
            '### Tukar status barang yang dipilih kepada 0 ### - End
                
            If Frm107_LM_NO_SIRI <> vbNullString Then
            
                '### Tukar status barang dari table data_database kepada 0 ### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from Data_Database where no_siri_produk='" & Frm107_LM_NO_SIRI & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Frm107.L57_Text = 0 Then
                        rs!form_out_status = 0
                    ElseIf Frm107.L57_Text = 1 Then
                        rs!form_out_status = 4
                    End If
                    rs.Update
                End If
                
                rs.Close
                Set rs = Nothing
                
                '### Masukkan barang yang dipilih ke dalam temp table ### - End
                
            End If
            
            GM_NEXT_PREV = 2
            
            If Frm107.L3_Text <> -1 Then
              
                Call Frm107_senarai_barang_header
                Call Frm107_senarai_barang '2
                
            End If
            
            Call Frm107_senarai_hantar_header
            Call Frm107_senarai_hantar
            
        End If
        
    End If
    
End If
End Sub
Private Sub Frm107_SM_senarai_barang_Click()
'On Error Resume Next
If Frm107.MSFlexGrid5 <> vbNullString Then
    Frm107_LM_ID = Frm107.MSFlexGrid5.TextMatrix(Frm107.MSFlexGrid5, 2) 'No. ID
    
    If Frm107_LM_ID <> vbNullString Then
        If IsNumeric(Frm107_LM_ID) Then
        
            Note = "Sistem mungkin mengambil sedikit masa untuk mengeluarkan senarai ini." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Teruskan?"
            
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbYes Then
            
                GM_NEXT_PREV = 0
                Frm107.L53_Text = -1 'Titik Pencarian Data
                Frm107.L54_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
                Frm107.L51_Text = 0 'Paparan Page ke-xxx
                
                Call Frm107_report_senarai_hantar_header
                Call Frm107_report_senarai_hantar
                
                If Frm107.L55_Text = 0 Then
                    MsgBox "Tiada senarai dijumpai.", vbInformation, "Info"
                Else
                    Frm107.Pic7.Visible = True
                    Frm107.Pic6.Visible = False
                End If
                
            End If

        End If
    End If
End If
    
    
    

End Sub
Private Sub Frm107_SM_senarai_barang_excel_Click()
'On Error Resume Next
If G_No_STATMENT_FORM <> vbNullString Then
    Call Frm107_senarai_barang_hantar_excel
End If
End Sub
Private Sub Frm107_SM_senarai_description_Click()
'On Error Resume Next
Note = "Sistem mungkin mengambil sedikit masa untuk mengeluarkan senarai ini." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then

    GM_NEXT_PREV = 0
    Frm107.L46_Text = -1 'Titik Pencarian Data
    Frm107.L47_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm107.L44_Text = 0 'Paparan Page ke-xxx
    
    Call Frm107_report_description_header
    Call Frm107_report_description
    
    If Frm107.L48_Text = 0 Then
        MsgBox "Tiada senarai dijumpai.", vbInformation, "Info"
    Else
        Frm107.Pic6.Visible = True
        Frm107.Pic5.Visible = False
    End If
    
End If
End Sub
Private Sub L12_Text_Click()
'on error resume next
If Frm107.Pic1.Visible = False Then
    Call Frm107_initial_location
    Call Frm107_initial_location2
    Call Frm107_initial_location3
    Call Frm107_initial_setting1
    Call Frm107_initial_setting2
    Call Frm107_initial_setting4
    Call Frm107_visible_component_1
    Call Frm107_visible_component_2
    Call Frm107_clear_status
    
    Frm107.TB3 = "1.00"
    
    Frm107.Pic1.Visible = True
Else
    Frm107.Pic1.Visible = False
End If
End Sub
Private Sub L17_Text_Click()
'on error resume next
If Frm107.Pic4.Visible = False Then
    Call Frm107_initial_location
    Call Frm107_initial_location2
    Call Frm107_initial_location3
    Call Frm107_initial_setting3
  
    Frm107.Pic4.Visible = True
Else
    Frm107.Pic4.Visible = False
End If
End Sub
Private Sub L18_Text_Click()
'on error resume next
If Frm107.Pic3.Visible = False Then
    Call Frm107_initial_location2
    Call Frm107_initial_setting
    
    Frm107.Pic3.Visible = True
Else
    Frm107.Pic3.Visible = False
End If
End Sub
Private Sub L21_Text_Click()
'On Error Resume Next
Note = "Adakah anda ingin kembali ke menu sebelumnya?" & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then

    GM_NEXT_PREV = 2

    Call Frm107_senarai_description_header
    Call Frm107_senarai_description
    
    Frm107.Pic1.Visible = True
    Frm107.Pic2.Visible = False
    
    Frm107.L12_Text.Visible = True
    Frm107.L17_Text.Visible = True

End If
End Sub
Private Sub L29_Text_Change()
'on error resume next
Call Frm107_tukaran_mutu
End Sub



Private Sub MSFlexGrid1_DblClick()
'On Error Resume Next
Frm107_LM_STATUS = vbNullString

If Frm107.MSFlexGrid1 <> vbNullString Then
    
    If IsNumeric(Frm107.MSFlexGrid1) Then
        Frm107_LM_STATUS = Frm107.MSFlexGrid1.TextMatrix(Frm107.MSFlexGrid1, 8) 'Status
        
        If Frm107_LM_STATUS <> vbNullString Then
            If Frm107_LM_STATUS = "Sudah dipilih" Then
                MsgBox "Barang ini telah dimasukkan ke dalam senarai barang yang akan dihantar kepada kilang/supplier.", vbExclamation, "Info"
            ElseIf Frm107_LM_STATUS = "Belum dipilih" Then
                Frm107.Frm107_SM_pilih.Caption = "Pilih barang ini? No. siri produk : " & Frm107.MSFlexGrid1.TextMatrix(Frm107.MSFlexGrid1, 3)
                PopupMenu Frm107_PM_menu1, vbPopupMenuRightButton
            End If
        End If
        
    End If
        
End If
End Sub
Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'On Error Resume Next
Frm107_LM_STATUS = vbNullString

If Frm107.MSFlexGrid1 <> vbNullString Then
    If Button = vbRightButton Then
    
        If IsNumeric(Frm107.MSFlexGrid1) Then
            Frm107_LM_STATUS = Frm107.MSFlexGrid1.TextMatrix(Frm107.MSFlexGrid1, 7) 'Status
            
            If Frm107_LM_STATUS <> vbNullString Then
                If Frm107_LM_STATUS = "Sudah dipilih" Then
                    MsgBox "Barang ini telah dimasukkan ke dalam senarai barang yang akan dihantar kepada kilang/supplier.", vbExclamation, "Info"
                ElseIf Frm107_LM_STATUS = "Belum dipilih" Then
                    Frm107.Frm107_SM_pilih.Caption = "Pilih barang ini? No. siri produk : " & Frm107.MSFlexGrid1.TextMatrix(Frm107.MSFlexGrid1, 3)
                    PopupMenu Frm107_PM_menu1, vbPopupMenuRightButton
                End If
            End If
            
        End If
        
    End If
End If
End Sub
Private Sub MSFlexGrid2_DblClick()
'On Error Resume Next
Frm107_LM_ID = vbNullString

If Frm107.MSFlexGrid2 <> vbNullString Then
    
    If IsNumeric(Frm107.MSFlexGrid2) Then
        Frm107_LM_ID = Frm107.MSFlexGrid2.TextMatrix(Frm107.MSFlexGrid2, 2) 'ID
        
        If Frm107_LM_ID <> vbNullString Then
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from " & G_FORM_OUT_DESC & " where ID='" & Frm107_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!id_desc) Then
                    Frm107.L19_Text = rs!id_desc 'No. rujukan ID description (Running numnber) (Auto generated number)
                End If
            End If
            
            rs.Close
            Set rs = Nothing
            
            Frm107.L20_Text = Frm107.MSFlexGrid2.TextMatrix(Frm107.MSFlexGrid2, 3) 'Description
            
            PopupMenu Frm107_PM_menu2, vbPopupMenuRightButton

        End If
    End If
        
End If
End Sub
Private Sub MSFlexGrid2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'On Error Resume Next
Frm107_LM_ID = vbNullString

If Frm107.MSFlexGrid2 <> vbNullString Then
    If Button = vbRightButton Then
    
        If IsNumeric(Frm107.MSFlexGrid2) Then
            Frm107_LM_ID = Frm107.MSFlexGrid2.TextMatrix(Frm107.MSFlexGrid2, 2) 'ID
            
            If Frm107_LM_ID <> vbNullString Then
            
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from " & G_FORM_OUT_DESC & " where ID='" & Frm107_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Not IsNull(rs!id_desc) Then
                        Frm107.L19_Text = rs!id_desc 'No. rujukan ID description (Running numnber) (Auto generated number)
                    End If
                End If
                
                rs.Close
                Set rs = Nothing
            
                Frm107.L20_Text = Frm107.MSFlexGrid2.TextMatrix(Frm107.MSFlexGrid2, 3) 'Description
                
                PopupMenu Frm107_PM_menu2, vbPopupMenuRightButton

            End If
            
        End If
        
    End If
End If
End Sub
Private Sub MSFlexGrid3_DblClick()
'On Error Resume Next
Frm107_LM_ID = vbNullString

If Frm107.MSFlexGrid3 <> vbNullString Then
    'If Button = vbRightButton Then
    
        If IsNumeric(Frm107.MSFlexGrid3) Then
            Frm107_LM_ID = Frm107.MSFlexGrid3.TextMatrix(Frm107.MSFlexGrid3, 2) 'ID
            
            If Frm107_LM_ID <> vbNullString Then
                Frm107.Frm107_SM_remove.Caption = "Keluarkan barang ini dari senarai? No. siri produk : " & Frm107.MSFlexGrid3.TextMatrix(Frm107.MSFlexGrid3, 3)
                PopupMenu Frm107_PM_menu3, vbPopupMenuRightButton
            End If
            
        End If
        
    'End If
End If
End Sub
Private Sub MSFlexGrid3_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'On Error Resume Next
Frm107_LM_ID = vbNullString

If Frm107.MSFlexGrid3 <> vbNullString Then
    If Button = vbRightButton Then
    
        If IsNumeric(Frm107.MSFlexGrid3) Then
            Frm107_LM_ID = Frm107.MSFlexGrid3.TextMatrix(Frm107.MSFlexGrid3, 2) 'ID
            
            If Frm107_LM_ID <> vbNullString Then
                Frm107.Frm107_SM_remove.Caption = "Keluarkan barang ini dari senarai? No. siri produk : " & Frm107.MSFlexGrid3.TextMatrix(Frm107.MSFlexGrid3, 3)
                PopupMenu Frm107_PM_menu3, vbPopupMenuRightButton
            End If
            
        End If
        
    End If
End If
End Sub
Private Sub MSFlexGrid4_DblClick()
'On Error Resume Next
Frm107_LM_ID = vbNullString

If Frm107.MSFlexGrid4 <> vbNullString Then
    'If Button = vbRightButton Then
    
        If IsNumeric(Frm107.MSFlexGrid4) Then
            Frm107_LM_ID = Frm107.MSFlexGrid4.TextMatrix(Frm107.MSFlexGrid4, 2) 'ID
            
            If Frm107_LM_ID <> vbNullString Then

                user_level = MDI_frm1.L4_Text
                
                If user_level = "Admin" Or user_level = "HQ" Or user_level = "Developer" Then
                
                    Frm107.Frm107_SM_edit_data.Enabled = True
                    Frm107.Frm107_SM_padam_data.Enabled = True
                            
                ElseIf user_level = "Manager" Then
                
                    Frm107.Frm107_SM_edit_data.Enabled = True
                    Frm107.Frm107_SM_padam_data.Enabled = False
                    
                Else
                
                    Frm107.Frm107_SM_edit_data.Enabled = False
                    Frm107.Frm107_SM_padam_data.Enabled = False
                
                End If
                
                Frm107.Frm107_SM_cetak_statement.Caption = "Cetak statement ini? No. Rujukan : " & Frm107.MSFlexGrid4.TextMatrix(Frm107.MSFlexGrid4, 4)
                Frm107.Frm107_SM_senarai_description.Caption = "Senarai maklumat terperinci dari statement ini. No. Rujukan : " & Frm107.MSFlexGrid4.TextMatrix(Frm107.MSFlexGrid4, 4)
                G_No_STATMENT_FORM = vbNullString
                G_No_STATMENT_FORM = Frm107.MSFlexGrid4.TextMatrix(Frm107.MSFlexGrid4, 4)
                PopupMenu Frm107_PM_menu4, vbPopupMenuRightButton
            End If
            
        End If
        
    'End If
End If
End Sub
Private Sub MSFlexGrid4_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'On Error Resume Next
Frm107_LM_ID = vbNullString

If Frm107.MSFlexGrid4 <> vbNullString Then
    If Button = vbRightButton Then
    
        If IsNumeric(Frm107.MSFlexGrid4) Then
            Frm107_LM_ID = Frm107.MSFlexGrid4.TextMatrix(Frm107.MSFlexGrid4, 2) 'ID
            
            If Frm107_LM_ID <> vbNullString Then
                Frm107.Frm107_SM_cetak_statement.Caption = "Cetak statement ini? No. Rujukan : " & Frm107.MSFlexGrid4.TextMatrix(Frm107.MSFlexGrid4, 4)
                G_No_STATMENT_FORM = vbNullString
                G_No_STATMENT_FORM = Frm107.MSFlexGrid4.TextMatrix(Frm107.MSFlexGrid4, 4)
                PopupMenu Frm107_PM_menu4, vbPopupMenuRightButton
            End If
            
        End If
        
    End If
End If
End Sub
Private Sub MSFlexGrid5_DblClick()
'On Error Resume Next
Frm107_LM_ID = vbNullString

If Frm107.MSFlexGrid5 <> vbNullString Then
    'If Button = vbRightButton Then
    
        If IsNumeric(Frm107.MSFlexGrid5) Then
            Frm107_LM_ID = Frm107.MSFlexGrid5.TextMatrix(Frm107.MSFlexGrid5, 2) 'ID
            
            If Frm107_LM_ID <> vbNullString Then
            
                G_No_RUJUKAN_FORM = vbNullString
                'G_No_STATMENT_FORM = vbNullString
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 58_form_out_list where ID='" & Frm107_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    'If Not IsNull(rs!no_rujukan) Then G_No_STATMENT_FORM = rs!no_rujukan
                    If Not IsNull(rs!id_desc) Then G_No_RUJUKAN_FORM = rs!id_desc
                End If
                
                rs.Close
                Set rs = Nothing
                
                PopupMenu Frm107_PM_menu5, vbPopupMenuRightButton
            End If
            
        End If
        
    'End If
End If
End Sub
Private Sub MSFlexGrid5_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'On Error Resume Next
Frm107_LM_ID = vbNullString

If Frm107.MSFlexGrid5 <> vbNullString Then
    If Button = vbRightButton Then
    
        If IsNumeric(Frm107.MSFlexGrid5) Then
            Frm107_LM_ID = Frm107.MSFlexGrid5.TextMatrix(Frm107.MSFlexGrid5, 2) 'ID
            
            If Frm107_LM_ID <> vbNullString Then
            
                G_No_RUJUKAN_FORM = vbNullString
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 58_form_out_list where ID='" & Frm107_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Not IsNull(rs!id_desc) Then G_No_RUJUKAN_FORM = rs!id_desc
                End If
                
                rs.Close
                Set rs = Nothing
                
                PopupMenu Frm107_PM_menu5, vbPopupMenuRightButton
            End If
            
        End If
        
    End If
End If
End Sub

Private Sub TB3_Change()
'on error resume next
Call Frm107_tukaran_mutu
End Sub
