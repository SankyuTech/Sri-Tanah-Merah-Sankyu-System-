VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm54 
   Caption         =   "Tetapan Harga Belian / Jualan"
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
   Icon            =   "Frm54.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tetapan Upah"
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
      Left            =   7200
      TabIndex        =   70
      Top             =   1320
      Visible         =   0   'False
      Width           =   21135
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
         Left            =   2235
         MouseIcon       =   "Frm54.frx":0ECA
         MousePointer    =   99  'Custom
         Picture         =   "Frm54.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox TB20 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2235
         MaxLength       =   10
         TabIndex        =   76
         Text            =   "TB20"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox TB21 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2235
         MaxLength       =   10
         TabIndex        =   75
         Text            =   "TB21"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox TB23 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2235
         MaxLength       =   10
         TabIndex        =   74
         Text            =   "TB23"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox TB22 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2235
         MaxLength       =   10
         TabIndex        =   73
         Text            =   "TB22"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox TB24 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2235
         MaxLength       =   10
         TabIndex        =   72
         Text            =   "TB24"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox TB25 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2235
         MaxLength       =   10
         TabIndex        =   71
         Text            =   "TB25"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm54.frx":379E
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2925
         Left            =   360
         TabIndex        =   84
         Top             =   4200
         Width           =   7920
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tetapan kenaikan upah mengikut kategori pelanggan."
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   83
         Top             =   360
         Width           =   5640
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "a ) Pelanggan biasa"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   82
         Top             =   765
         Width           =   3000
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "b ) Ahli biasa"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   81
         Top             =   1110
         Width           =   3000
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "d ) Gold"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   80
         Top             =   1800
         Width           =   3000
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "c ) Silver"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   79
         Top             =   1440
         Width           =   3000
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "e ) Platinum"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   78
         Top             =   2160
         Width           =   3000
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "f ) Master Dealer"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   77
         Top             =   2520
         Visible         =   0   'False
         Width           =   3000
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
      Height          =   9135
      Left            =   21360
      ScaleHeight     =   9135
      ScaleWidth      =   23535
      TabIndex        =   24
      Top             =   2280
      Visible         =   0   'False
      Width           =   23535
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tetapan Harga"
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
      Left            =   600
      TabIndex        =   27
      Top             =   2640
      Visible         =   0   'False
      Width           =   21135
      Begin VB.CommandButton CMD2 
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
         MouseIcon       =   "Frm54.frx":3BB0
         MousePointer    =   99  'Custom
         Picture         =   "Frm54.frx":3EBA
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   4680
         Width           =   2775
      End
      Begin VB.TextBox TB6 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   11760
         TabIndex        =   57
         Text            =   "TB6"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox TB31 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3480
         TabIndex        =   40
         Text            =   "TB31"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox TB30 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3480
         TabIndex        =   39
         Text            =   "TB30"
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox TB5 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3480
         TabIndex        =   38
         Text            =   "TB5"
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox TB4 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3480
         TabIndex        =   37
         Text            =   "TB4"
         Top             =   3360
         Width           =   1935
      End
      Begin VB.TextBox TB3 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3480
         TabIndex        =   36
         Text            =   "TB3"
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox TB2 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3480
         TabIndex        =   35
         Text            =   "TB2"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox TB12 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   9960
         TabIndex        =   31
         Text            =   "TB12"
         Top             =   1005
         Width           =   3135
      End
      Begin VB.TextBox TB1 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3250
         TabIndex        =   30
         Text            =   "TB1"
         Top             =   1005
         Width           =   3135
      End
      Begin VB.ComboBox CBB1 
         Height          =   360
         Left            =   3250
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   11640
         TabIndex        =   60
         Top             =   1965
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Harga semasa jualan"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8760
         TabIndex        =   59
         Top             =   1920
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1 ) Tetapan harga jualan kepada staff kedai."
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8760
         TabIndex        =   58
         Top             =   1560
         Visible         =   0   'False
         Width           =   8145
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "f ) Master Dealer"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   480
         TabIndex        =   56
         Top             =   4080
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.Label L14_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L14_Text"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5640
         TabIndex        =   55
         Top             =   4080
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "e ) Platinum"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   480
         TabIndex        =   54
         Top             =   3720
         Width           =   3000
      End
      Begin VB.Label L13_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L13_Text"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5640
         TabIndex        =   53
         Top             =   3720
         Width           =   1995
      End
      Begin VB.Label L6_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L6_Text"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5640
         TabIndex        =   52
         Top             =   3000
         Width           =   1995
      End
      Begin VB.Label L5_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L5_Text"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5640
         TabIndex        =   51
         Top             =   3360
         Width           =   1995
      End
      Begin VB.Label L4_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L4_Text"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5640
         TabIndex        =   50
         Top             =   2670
         Width           =   1995
      End
      Begin VB.Label L3_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L3_Text"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5640
         TabIndex        =   49
         Top             =   2325
         Width           =   1995
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "c ) Silver"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   480
         TabIndex        =   48
         Top             =   3000
         Width           =   3000
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "d ) Gold"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   480
         TabIndex        =   47
         Top             =   3360
         Width           =   3000
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "b ) Ahli biasa"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   480
         TabIndex        =   46
         Top             =   2670
         Width           =   3000
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "a ) Pelanggan biasa"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   480
         TabIndex        =   45
         Top             =   2325
         Width           =   3000
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Jualan Per Gram RM/g"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5640
         TabIndex        =   44
         Top             =   1920
         Width           =   2595
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Diskaun (RM/g)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3600
         TabIndex        =   43
         Top             =   1920
         Width           =   1875
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sasaran Pelanggan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   42
         Top             =   1920
         Width           =   3195
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tetapan harga jualan kepada setiap sasaran PELANGGAN."
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   480
         TabIndex        =   41
         Top             =   1560
         Width           =   8145
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Semasa Dari Supplier (RM/g) * :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   6300
         TabIndex        =   34
         Top             =   1035
         Width           =   3585
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Purity * :"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   360
         TabIndex        =   33
         Top             =   660
         Width           =   2865
      End
      Begin VB.Label Label58 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Semasa Jualan (RM/g) * :"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   360
         TabIndex        =   32
         Top             =   1035
         Width           =   2865
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Penetapan harga belian dan jualan emas semasa."
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   14475
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
      Height          =   8295
      Left            =   120
      ScaleHeight     =   8295
      ScaleWidth      =   16935
      TabIndex        =   0
      Top             =   11760
      Visible         =   0   'False
      Width           =   16935
      Begin VB.CommandButton CMD9 
         BackColor       =   &H000080FF&
         Caption         =   "Simpan Data"
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
         Left            =   11640
         MaskColor       =   &H00400000&
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Keluar Dari Menu Ini"
         Top             =   7440
         Width           =   3000
      End
      Begin VB.ComboBox CBB2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   12600
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   6480
         Width           =   3375
      End
      Begin VB.TextBox TB19 
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
         Left            =   12600
         TabIndex        =   13
         Top             =   6960
         Width           =   3375
      End
      Begin VB.CommandButton CMD8 
         BackColor       =   &H000080FF&
         Caption         =   "Daftar"
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
         Left            =   2520
         MaskColor       =   &H00400000&
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Keluar Dari Menu Ini"
         Top             =   7440
         Width           =   3000
      End
      Begin VB.TextBox TB18 
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
         Left            =   3480
         TabIndex        =   6
         Top             =   6840
         Width           =   3375
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4155
         Left            =   480
         TabIndex        =   2
         Top             =   1800
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   7329
         _Version        =   393216
         Rows            =   1
         BackColor       =   16777088
         BackColorFixed  =   -2147483645
         BackColorBkg    =   16777215
         GridColor       =   0
         WordWrap        =   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   3
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
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   4155
         Left            =   9360
         TabIndex        =   12
         Top             =   1800
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   7329
         _Version        =   393216
         Rows            =   1
         BackColor       =   16777088
         BackColorFixed  =   -2147483645
         BackColorBkg    =   16777215
         GridColor       =   0
         WordWrap        =   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   3
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
      Begin VB.Shape Shape7 
         Height          =   7455
         Left            =   9120
         Top             =   720
         Width           =   7335
      End
      Begin VB.Shape Shape6 
         Height          =   1935
         Left            =   9360
         Top             =   6120
         Width           =   6855
      End
      Begin VB.Label Label66 
         BackStyle       =   0  'Transparent
         Caption         =   "Tetapan upah bagi setiap kategori."
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
         Left            =   9600
         TabIndex        =   19
         Top             =   6120
         Width           =   6465
      End
      Begin VB.Label Label62 
         BackStyle       =   0  'Transparent
         Caption         =   "Tetapan Upah *   (RM)"
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
         Left            =   9840
         TabIndex        =   18
         Top             =   7005
         Width           =   2625
      End
      Begin VB.Label Label65 
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
         Left            =   12480
         TabIndex        =   17
         Top             =   6525
         Width           =   135
      End
      Begin VB.Label Label64 
         BackStyle       =   0  'Transparent
         Caption         =   "Kategori Upah *"
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
         Left            =   9840
         TabIndex        =   16
         Top             =   6525
         Width           =   2265
      End
      Begin VB.Label Label63 
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
         Left            =   12480
         TabIndex        =   14
         Top             =   7005
         Width           =   135
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "Tetapan upah bagi setiap kategori yang telah ditetapkan. "
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
         Left            =   9360
         TabIndex        =   11
         Top             =   1320
         Width           =   6705
      End
      Begin VB.Label Label60 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tetapan Upah"
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
         Left            =   9360
         TabIndex        =   10
         Top             =   840
         Width           =   6855
      End
      Begin VB.Shape Shape5 
         Height          =   7455
         Left            =   240
         Top             =   720
         Width           =   7335
      End
      Begin VB.Shape Shape3 
         Height          =   1935
         Left            =   480
         Top             =   6120
         Width           =   6855
      End
      Begin VB.Label Label57 
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
         Left            =   3360
         TabIndex        =   8
         Top             =   6885
         Width           =   135
      End
      Begin VB.Label Label56 
         BackStyle       =   0  'Transparent
         Caption         =   "Kategori Upah *"
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
         Left            =   720
         TabIndex        =   7
         Top             =   6885
         Width           =   2265
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Caption         =   "Anda boleh daftarkan kategori upah dengan mengisi butiran   di bawah."
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
         Height          =   525
         Left            =   600
         TabIndex        =   5
         Top             =   6240
         Width           =   6705
      End
      Begin VB.Label Label54 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Senarai Kategori Bagi Upah"
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
         Left            =   480
         TabIndex        =   4
         Top             =   840
         Width           =   6855
      End
      Begin VB.Label Label53 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai di bawah adalah kategori upah yang telah didaftarkan di dalam sistem."
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
         Height          =   525
         Left            =   600
         TabIndex        =   3
         Top             =   1200
         Width           =   6705
      End
      Begin VB.Label Label52 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm54.frx":6484
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
         Height          =   645
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   16200
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Senarai Harga"
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
      Left            =   840
      TabIndex        =   25
      Top             =   840
      Visible         =   0   'False
      Width           =   21135
      Begin VB.CommandButton CMD6 
         Caption         =   "Back"
         Height          =   810
         Left            =   18600
         MouseIcon       =   "Frm54.frx":6528
         MousePointer    =   99  'Custom
         Picture         =   "Frm54.frx":6832
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Paparan Sebelum"
         Top             =   10560
         Width           =   1095
      End
      Begin VB.CommandButton CMD14 
         Caption         =   "Next"
         Height          =   810
         Left            =   19800
         MouseIcon       =   "Frm54.frx":78FC
         MousePointer    =   99  'Custom
         Picture         =   "Frm54.frx":7C06
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Paparan Seterusnya"
         Top             =   10560
         Width           =   1095
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   9930
         Left            =   240
         TabIndex        =   26
         Top             =   600
         Width           =   20655
         _ExtentX        =   36433
         _ExtentY        =   17515
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
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label67 
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
         Left            =   16320
         TabIndex        =   69
         Top             =   10560
         Width           =   2295
      End
      Begin VB.Label L62_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L62_Text"
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
         Left            =   18240
         TabIndex        =   68
         Top             =   10560
         Width           =   615
      End
      Begin VB.Label L61_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   300
         Left            =   17640
         TabIndex        =   67
         Top             =   10560
         Width           =   375
      End
      Begin VB.Label L63_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L63_Text"
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
         Height          =   300
         Left            =   16800
         TabIndex        =   66
         Top             =   10920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L64_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L64_Text"
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
         Height          =   300
         Left            =   15720
         TabIndex        =   65
         Top             =   10920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Tetapan harga jualan bagi setiap kategori sasaran yang telah dibuat."
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   360
         TabIndex        =   62
         Top             =   280
         Width           =   13320
      End
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Tetapan Upah"
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
      MouseIcon       =   "Frm54.frx":8CD0
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Tetapan Harga"
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
      Left            =   3960
      MouseIcon       =   "Frm54.frx":8FDA
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Senarai Harga Jualan Mengikut Purity"
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
      MouseIcon       =   "Frm54.frx":92E4
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   120
      Width           =   3615
   End
   Begin VB.Menu Frm54_Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm54_PadamData 
         Caption         =   "Padam Data"
      End
   End
End
Attribute VB_Name = "Frm54"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CBB1_Change()
'On Error Resume Next
Call Frm54_ClearAllField
SearchDisable = 1

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from hargaemas where Purity='" & Frm54.CBB1 & "' AND cawangan='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm54.TB12 = Format(rs!HargaDariSupplier, "0.00") 'Harga Dari Supplier
    Frm54.L3_Text = Format(rs!Harga_Pelanggan, "0.00") 'Harga Pelanggan
    Frm54.L4_Text = Format(rs!Harga_Member, "0.00") 'Harga Member
    Frm54.L5_Text = Format(rs!Harga_RAF, "0.00") 'Harga RAF
    Frm54.L6_Text = Format(rs!Harga_Pengedar, "0.00") 'Harga Pengedar
    'Frm54.TB1 = Format(rs!HargaMKS, "0.00") 'Harga MKS
    Frm54.TB2 = Format(rs!Pemalar_Pelanggan, "0.00") 'Pemalar Pelanggan
    Frm54.TB3 = Format(rs!Pemalar_Member, "0.00") 'Pemalar Member
    Frm54.TB4 = Format(rs!Pemalar_RAF, "0.00") 'Pemalar RAF
    Frm54.TB5 = Format(rs!Pemalar_Pengedar, "0.00") 'Pemalar Pengedar
    Frm54.TB1 = Format(rs!HargaMKS, "0.00") 'Harga MKS
    Frm54.TB30 = Format(rs!pemalar_nd, "0.00") 'Pemalar Bagi Normal Dealer
    Frm54.TB31 = Format(rs!pemalar_md, "0.00") 'Pemalar Bagi Master Dealer
    'Call TetapanHargaJualan
End If

rs.Close
Set rs = Nothing

SearchDisable = 0
End Sub
Private Sub CBB1_Click()
'On Error Resume Next
Call Frm54_ClearAllField
SearchDisable = 1

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from hargaemas where Purity='" & Frm54.CBB1 & "' AND cawangan='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!harga_staff) Then Frm54.TB6 = Format(rs!harga_staff, "0.00") 'Harga jualan kepada staff
    If Not IsNull(rs!HargaDariSupplier) Then Frm54.TB12 = Format(rs!HargaDariSupplier, "0.00") 'Harga Dari Supplier
    If Not IsNull(rs!Harga_Pelanggan) Then Frm54.L3_Text = Format(rs!Harga_Pelanggan, "0.00") 'Harga Pelanggan
    If Not IsNull(rs!Harga_Member) Then Frm54.L4_Text = Format(rs!Harga_Member, "0.00") 'Harga Member
    If Not IsNull(rs!Harga_RAF) Then Frm54.L5_Text = Format(rs!Harga_RAF, "0.00") 'Harga RAF
    If Not IsNull(rs!Harga_Pengedar) Then Frm54.L6_Text = Format(rs!Harga_Pengedar, "0.00") 'Harga Pengedar
    If Not IsNull(rs!HargaMKS) Then Frm54.TB1 = Format(rs!HargaMKS, "0.00") 'Harga MKS
    If Not IsNull(rs!Pemalar_Pelanggan) Then Frm54.TB2 = Format(rs!Pemalar_Pelanggan, "0.00") 'Pemalar Pelanggan
    If Not IsNull(rs!Pemalar_Member) Then Frm54.TB3 = Format(rs!Pemalar_Member, "0.00") 'Pemalar Member
    If Not IsNull(rs!Pemalar_RAF) Then Frm54.TB4 = Format(rs!Pemalar_RAF, "0.00") 'Pemalar RAF
    If Not IsNull(rs!Pemalar_Pengedar) Then Frm54.TB5 = Format(rs!Pemalar_Pengedar, "0.00") 'Pemalar Pengedar
    If Not IsNull(rs!pemalar_nd) Then Frm54.TB30 = Format(rs!pemalar_nd, "0.00") 'Pemalar Bagi Normal Dealer
    If Not IsNull(rs!pemalar_md) Then Frm54.TB31 = Format(rs!pemalar_md, "0.00") 'Pemalar Bagi Master Dealer
    'Call TetapanHargaJualan
End If

rs.Close
Set rs = Nothing

SearchDisable = 0
End Sub
Private Sub CMD1_Click()
'On Error Resume Next
Dim Err(20)
DATA_SAVE = 0

If Frm54.TB20 = vbNullString Or (Frm54.TB20 <> vbNullString And Not IsNumeric(Frm54.TB20)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kenaikan upah bagi pelanggan biasa]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm54.TB21 = vbNullString Or (Frm54.TB21 <> vbNullString And Not IsNumeric(Frm54.TB21)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kenaikan upah bagi ahli biasa]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm54.TB22 = vbNullString Or (Frm54.TB22 <> vbNullString And Not IsNumeric(Frm54.TB22)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kenaikan upah bagi silver]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm54.TB23 = vbNullString Or (Frm54.TB23 <> vbNullString And Not IsNumeric(Frm54.TB23)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kenaikan upah bagi gold]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm54.TB24 = vbNullString Or (Frm54.TB24 <> vbNullString And Not IsNumeric(Frm54.TB24)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kenaikan upah bagi platinum]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
'If Frm54.TB25 = vbNullString Or (Frm54.TB25 <> vbNullString And Not IsNumeric(Frm54.TB25)) Then
'    x = x + 1
'    Err(x) = "Sila masukkan [Kenaikan upah bagi master dealer]. Hanya NOMBOR dibenarkan dalam ruangan ini."
'End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah anda ingin simpan tetapan ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 73_tetapan_upah where default_setting='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            If Frm54.TB20 <> vbNullString Then 'Kenaikan upah bagi pelanggan
                rs!pelanggan = Format(Frm54.TB20, "0.00")
            Else
                rs!pelanggan = Format(0, "0.00")
            End If
            If Frm54.TB21 <> vbNullString Then 'Kenaikan upah bagi Member
                rs!Member = Format(Frm54.TB21, "0.00")
            Else
                rs!Member = Format(0, "0.00")
            End If
            If Frm54.TB22 <> vbNullString Then 'Kenaikan upah bagi Pengedar
                rs!Pengedar = Format(Frm54.TB22, "0.00")
            Else
                rs!Pengedar = Format(0, "0.00")
            End If
            If Frm54.TB23 <> vbNullString Then 'Kenaikan upah bagi RAF
                rs!raf = Format(Frm54.TB23, "0.00")
            Else
                rs!raf = Format(0, "0.00")
            End If
            If Frm54.TB24 <> vbNullString Then 'Kenaikan upah bagi Normal Dealer
                rs!normal_dealer = Format(Frm54.TB24, "0.00")
            Else
                rs!normal_dealer = Format(0, "0.00")
            End If
            If Frm54.TB25 <> vbNullString Then 'Kenaikan upah bagi Master Dealer
                rs!master_dealer = Format(Frm54.TB25, "0.00")
            Else
                rs!master_dealer = Format(0, "0.00")
            End If
            rs.Update
            
            DATA_SAVE = 1
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_SAVE = 1 Then
        
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & user & "] Tetapan kenaikan upah mengikut kategori."
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
            
            MsgBox "Tetapan telah berjaya disimpan.", vbInformation, "Info"
        End If
    
    End If
End If
End Sub

Private Sub CMD14_Click()
'on error resume next
Dim Frm54_LM_CURR_PAGE As Double
Dim Frm54_LM_TOTAL_PAGE As Double

Frm54_LM_CURR_PAGE = 0
Frm54_LM_TOTAL_PAGE = 0

If Frm54.L61_Text <> vbNullString And IsNumeric(Frm54.L61_Text) Then
    If Frm54.L62_Text <> vbNullString And IsNumeric(Frm54.L62_Text) Then
        Frm54_LM_CURR_PAGE = Frm54.L61_Text
        Frm54_LM_TOTAL_PAGE = Frm54.L62_Text
        
        If Frm54_LM_CURR_PAGE < Frm54_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call frm54_senarai_harga_header
            Call frm54_senarai_harga
            
        End If
    End If
End If
End Sub

Private Sub CMD2_Click()
'On Error Resume Next
Dim Err(20)
Dim Frm54_LM_HARGA_PELANGGAN As Double
Dim Frm54_LM_HARGA_MEMBER As Double
Dim Frm54_LM_HARGA_PENGEDAR As Double
Dim Frm54_LM_HARGA_RAF As Double
Dim Frm54_LM_HARGA_NORMAL As Double
Dim Frm54_LM_HARGA_MASTER As Double
Dim Frm54_LM_HARGA_STAFF As Double
x = 0

Frm54_LM_HARGA_PELANGGAN = 0
Frm54_LM_HARGA_MEMBER = 0
Frm54_LM_HARGA_PENGEDAR = 0
Frm54_LM_HARGA_RAF = 0
Frm54_LM_HARGA_NORMAL = 0
Frm54_LM_HARGA_MASTER = 0
Frm54_LM_HARGA_STAFF = 0

If Frm54.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Purity]."
End If
If Frm54.TB1 = vbNullString Or (Frm54.TB1 <> vbNullString And Not IsNumeric(Frm54.TB1)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Harga Semasa (MKS)]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm54.TB12 = vbNullString Or (Frm54.TB12 <> vbNullString And Not IsNumeric(Frm54.TB12)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Harga Dari Supplier]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm54.TB2 = vbNullString Or (Frm54.TB2 <> vbNullString And Not IsNumeric(Frm54.TB2)) Then
    x = x + 1
    Err(x) = "Sila Masukkan Diskaun Bagi [Pelanggan biasa] Dalam Tetapan Harga Jualan. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm54.TB3 = vbNullString Or (Frm54.TB3 <> vbNullString And Not IsNumeric(Frm54.TB3)) Then
    x = x + 1
    Err(x) = "Sila Masukkan Diskaun Bagi [Ahli biasa] Dalam Tetapan Harga Jualan. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm54.TB4 = vbNullString Or (Frm54.TB4 <> vbNullString And Not IsNumeric(Frm54.TB4)) Then
    x = x + 1
    Err(x) = "Sila Masukkan Diskaun Bagi [Gold] Dalam Tetapan Harga Jualan. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm54.TB5 = vbNullString Or (Frm54.TB5 <> vbNullString And Not IsNumeric(Frm54.TB5)) Then
    x = x + 1
    Err(x) = "Sila Masukkan Diskaun Bagi [Silver] Dalam Tetapan Harga Jualan. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm54.TB30 = vbNullString Or (Frm54.TB30 <> vbNullString And Not IsNumeric(Frm54.TB30)) Then
    x = x + 1
    Err(x) = "Sila Masukkan Diskaun Bagi [Platinum] Dalam Tetapan Harga Jualan. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
'If Frm54.TB31 = vbNullString Or (Frm54.TB31 <> vbNullString And Not IsNumeric(Frm54.TB31)) Then
'    x = x + 1
'    Err(x) = "Sila Masukkan Diskaun Bagi [Master Dealer] Dalam Tetapan Harga Jualan Secara Ansuran. Hanya NOMBOR dibenarkan dalam ruangan ini."
'End If

GoTo Skip_harga_staff:
If Frm54.TB6 = vbNullString Or (Frm54.TB6 <> vbNullString And Not IsNumeric(Frm54.TB6)) Then
    x = x + 1
    Err(x) = "Sila Masukkan Harga Jualan Kepada Staff Kedai. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If (Frm54.TB6 <> vbNullString And IsNumeric(Frm54.TB6)) Then
    If Format(Frm54.TB6, "0.00") = "0.00" Then
        x = x + 1
        Err(x) = "[0.00] Tidak Dibenarkan Bagi Harga Jualan Kepada Staff Kedai."
    End If
End If

'### Periksa tetapan harga jualan kepada staff ### - Start
If Frm54.TB6 <> vbNullString And IsNumeric(Frm54.TB6) Then
    Frm54_LM_HARGA_STAFF = Frm54.TB6 'Harga jualan kepada staff
End If
If Frm54.L3_Text <> vbNullString And IsNumeric(Frm54.L3_Text) Then
    Frm54_LM_HARGA_PELANGGAN = Frm54.L3_Text 'Harga jualan kepada pelanggan
End If
If Frm54.L4_Text <> vbNullString And IsNumeric(Frm54.L4_Text) Then
    Frm54_LM_HARGA_MEMBER = Frm54.L4_Text 'Harga jualan kepada member
End If
If Frm54.L6_Text <> vbNullString And IsNumeric(Frm54.L6_Text) Then
    Frm54_LM_HARGA_PENGEDAR = Frm54.L6_Text 'Harga jualan kepada pengedar
End If
If Frm54.L5_Text <> vbNullString And IsNumeric(Frm54.L5_Text) Then
    Frm54_LM_HARGA_RAF = Frm54.L5_Text 'Harga jualan kepada RAF
End If
If Frm54.L13_Text <> vbNullString And IsNumeric(Frm54.L13_Text) Then
    Frm54_LM_HARGA_NORMAL = Frm54.L13_Text 'Harga jualan kepada Normal Dealer
End If
If Frm54.L14_Text <> vbNullString And IsNumeric(Frm54.L14_Text) Then
    Frm54_LM_HARGA_MASTER = Frm54.L14_Text 'Harga jualan kepada Master Dealer
End If

If Frm54_LM_HARGA_STAFF > Frm54_LM_HARGA_PELANGGAN Then
    x = x + 1
    Err(x) = "Harga Jualan Kepada Pelanggan Tidak Boleh Melebihi Dari Harga Jualan Kepada Staff Kedai."
End If
If Frm54_LM_HARGA_STAFF > Frm54_LM_HARGA_MEMBER Then
    x = x + 1
    Err(x) = "Harga Jualan Kepada Member Tidak Boleh Melebihi Dari Harga Jualan Kepada Staff Kedai."
End If
If Frm54_LM_HARGA_STAFF > Frm54_LM_HARGA_PENGEDAR Then
    x = x + 1
    Err(x) = "Harga Jualan Kepada Pengedar Tidak Boleh Melebihi Dari Harga Jualan Kepada Staff Kedai."
End If
If Frm54_LM_HARGA_STAFF > Frm54_LM_HARGA_RAF Then
    x = x + 1
    Err(x) = "Harga Jualan Kepada RAF Tidak Boleh Melebihi Dari Harga Jualan Kepada Staff Kedai."
End If
If Frm54_LM_HARGA_STAFF > Frm54_LM_HARGA_NORMAL Then
    x = x + 1
    Err(x) = "Harga Jualan Kepada Normal Dealer Tidak Boleh Melebihi Dari Harga Jualan Kepada Staff Kedai."
End If
If Frm54_LM_HARGA_STAFF > Frm54_LM_HARGA_MASTER Then
    x = x + 1
    Err(x) = "Harga Jualan Kepada Master Dealer Tidak Boleh Melebihi Dari Harga Jualan Kepada Staff Kedai."
End If
'### Periksa tetapan harga jualan kepada staff ### - End
Skip_harga_staff:

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
        
        LM_NOW = Now
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from hargaemas where Purity='" & Frm54.CBB1 & "' AND cawangan='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            rs!HargaMKS = Format(Frm54.TB1, "0.00") 'Harga MKS
            rs!HargaDariSupplier = Format(Frm54.TB12, "0.00") 'Harga Dari Supplier
            'rs!harga_staff = Format(Frm54.TB6, "0.00") 'Harga Jualan Kepada Staff
            rs!Harga_Pelanggan = Format(Frm54.L3_Text, "0.00") 'Harga Pelanggan
            rs!Harga_Member = Format(Frm54.L4_Text, "0.00") 'Harga Member
            rs!Harga_RAF = Format(Frm54.L5_Text, "0.00") 'Harga RAF
            rs!Harga_Pengedar = Format(Frm54.L6_Text, "0.00") 'Harga Pengedar
            rs!harga_nd = Format(Frm54.L13_Text, "0.00") 'Harga ND
            rs!harga_md = Format(Frm54.L14_Text, "0.00") 'Harga MD
            rs!Pemalar_Pelanggan = Format(Frm54.TB2, "0.00") 'Pemalar Pelanggan
            rs!Pemalar_Member = Format(Frm54.TB3, "0.00") 'Pemalar Member
            rs!Pemalar_RAF = Format(Frm54.TB4, "0.00") 'Pemalar RAF
            rs!Pemalar_Pengedar = Format(Frm54.TB5, "0.00") 'Pemalar Pengedar
            rs!pemalar_nd = Format(Frm54.TB30, "0.00") 'Pemalar Bagi Normal Dealer
            rs!pemalar_md = Format(Frm54.TB31, "0.00") 'Pemalar Bagi Master Dealer
            rs!cawangan = MDI_frm1.L20_Text
            rs!write_timestamp = LM_NOW
            rs.Update
        Else
            rs.AddNew
            rs!purity = Frm54.CBB1 'Purity
            rs!HargaMKS = Format(Frm54.TB1, "0.00") 'Harga MKS
            rs!HargaDariSupplier = Format(Frm54.TB12, "0.00") 'Harga Dari Supplier
            'rs!harga_staff = Format(Frm54.TB6, "0.00") 'Harga Jualan Kepada Staff
            rs!Harga_Pelanggan = Format(Frm54.L3_Text, "0.00") 'Harga Pelanggan
            rs!Harga_Member = Format(Frm54.L4_Text, "0.00") 'Harga Member
            rs!Harga_RAF = Format(Frm54.L5_Text, "0.00") 'Harga RAF
            rs!Harga_Pengedar = Format(Frm54.L6_Text, "0.00") 'Harga Pengedar
            rs!harga_nd = Format(Frm54.L13_Text, "0.00") 'Harga ND
            rs!harga_md = Format(Frm54.L14_Text, "0.00") 'Harga MD
            rs!Pemalar_Pelanggan = Format(Frm54.TB2, "0.00") 'Pemalar Pelanggan
            rs!Pemalar_Member = Format(Frm54.TB3, "0.00") 'Pemalar Member
            rs!Pemalar_RAF = Format(Frm54.TB4, "0.00") 'Pemalar RAF
            rs!Pemalar_Pengedar = Format(Frm54.TB5, "0.00") 'Pemalar Pengedar
            rs!pemalar_nd = Format(Frm54.TB30, "0.00") 'Pemalar Bagi Normal Dealer
            rs!pemalar_md = Format(Frm54.TB31, "0.00") 'Pemalar Bagi Master Dealer
            rs!cawangan = MDI_frm1.L20_Text
            rs!write_timestamp = LM_NOW
            rs.Update
        End If
        
        rs.Close
        Set rs = Nothing
        
        Call Frm54_rekod_tetapan_harga
        
        user = MDI_frm1.L3_Text
        LogAct_Memory = "[" & user & "] Update Harga Jualan [" & Frm54.CBB1 & "]"
        LogDate_Memory = LM_NOW
        Call UpdateLog_Database
        
        MsgBox "Data telah BERJAYA disimpan.", vbInformation, "Info"
    End If
End If
End Sub
Private Sub CMD7_Click()
'On Error Resume Next
Call Frm54_ClearAllField
Call Frm54_KategoriUpah
Frm54.Frame1.Visible = False
Frm54.Frame2.Visible = False
Frm54.Frame3.Visible = False
Frm54.Pic5.Visible = True
End Sub

Private Sub CMD6_Click()
'on error resume next
Dim Frm54_LM_CURR_PAGE As Double
Dim Frm54_LM_TOTAL_PAGE As Double

Frm54_LM_CURR_PAGE = 0
Frm54_LM_TOTAL_PAGE = 0

If Frm54.L61_Text <> vbNullString And IsNumeric(Frm54.L61_Text) Then
    If Frm54.L62_Text <> vbNullString And IsNumeric(Frm54.L62_Text) Then
        Frm54_LM_CURR_PAGE = Frm54.L61_Text
        Frm54_LM_TOTAL_PAGE = Frm54.L62_Text
        
        If Frm54_LM_CURR_PAGE <> 1 And Frm54_LM_CURR_PAGE <> 0 Then
        
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
                    
            Call frm54_senarai_harga_header
            Call frm54_senarai_harga
                    
        End If

    End If
End If
End Sub

Private Sub CMD8_Click()
'On Error Resume Next
If Frm54.TB18 = vbNullString Then
    MsgBox "Sila Masukkan Kategori Upah.", vbExclamation, "Error"
    Exit Sub
End If

Note = "Adakah Anda Ingin Simpan Data Ini ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from tetapanupah where KategoriUpah='" & Frm54.TB18 & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        MsgBox "Kategori Upah Bagi " & UCase(Frm54.TB18) & " Telah Didaftarkan Sebelum Ini.", vbExclamation, "Info"
    Else
        rs.AddNew
        rs!KategoriUpah = UCase(Frm54.TB18) 'Kategori Upah
        rs.Update
    End If
    
    rs.Close
    Set rs = Nothing
    
    user = MDI_frm1.L3_Text
    LogAct_Memory = "[" & user & "] Pendaftaran Kategori Upah [" & UCase(Frm54.TB18) & "]."
    LogDate_Memory = DateTime.Date & " " & DateTime.Time$
    UpdateLog_Database
    
    Call Frm54_KategoriUpah
    Frm54.TB18 = vbNullString
    MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
End If
End Sub
Private Sub CMD9_Click()
'On Error Resume Next
If Frm54.CBB2 = vbNullString Then
    MsgBox "Sila Pilih Kategori Upah.", vbExclamation, "Error"
    Exit Sub
End If
If Frm54.TB19 = vbNullString Or (Frm54.TB19 <> vbNullString And Not IsNumeric(Frm54.TB19)) Then
    MsgBox "Sila Masukkan Tetapan Upah. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini.", vbExclamation, "Error"
    Exit Sub
End If

Note = "Adakah Anda Ingin Simpan Data Ini ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from tetapanupah where KategoriUpah='" & Frm54.CBB2 & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        rs!tetapanupah = Frm54.TB19 'Format(Frm54.TB19, "0.00") 'Tetapan Upah
        rs.Update
    End If
    
    rs.Close
    Set rs = Nothing
    
    user = MDI_frm1.L3_Text
    LogAct_Memory = "[" & user & "] Tetapan Upah [" & UCase(Frm54.CBB2) & "]."
    LogDate_Memory = DateTime.Date & " " & DateTime.Time$
    UpdateLog_Database
    
    Call Frm54_KategoriUpah
    Frm54.TB19 = vbNullString
    MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
End If
End Sub

Private Sub Frm54_PadamData_Click()
'On Error Resume Next
Note = "Adakah Anda Ingin Padam Data Ini ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    OBJEK = Frm54.MSFlexGrid1.TextMatrix(Frm54.MSFlexGrid1, 2) 'Kategori Upah
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from tetapanupah where KategoriUpah='" & OBJEK & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        rs.Delete
        rs.Update
    End If
    
    rs.Close
    Set rs = Nothing
    
    user = MDI_frm1.L3_Text
    LogAct_Memory = "[" & user & "] Padam Kategori Upah [" & OBJEK & "]."
    LogDate_Memory = DateTime.Date & " " & DateTime.Time$
    UpdateLog_Database
    
    Call Frm54_KategoriUpah
    MsgBox "Data Telah Berjaya Dipadam.", vbInformation, "Info"
End If
End Sub
Private Sub Label24_Click()
'On Error Resume Next
If Frm54.Frame1.Visible = False Then
    
    Call frm54_initial_location
    
    GM_NEXT_PREV = 0
    
    Frm54.L63_Text = -1 'Titik Pencarian Data
    Frm54.L64_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm54.L61_Text = 0 'Paparan Page ke-xxx
    
    Call frm54_senarai_harga_header
    Call frm54_senarai_harga
    
    Frm54.Frame1.Visible = True

Else

    Frm54.Frame1.Visible = False
    
End If
End Sub
Private Sub Label25_Click()
'On Error Resume Next
If MDI_frm1.L20_Text = "Semua cawangan" Then

    Frm96.CMD2.Visible = True
    Frm96.CMD1.Visible = False

    Call Frm96_initial
    
    Frm96.Show vbModal
    
End If
    
If Frm54.Frame2.Visible = False Then
    
    Call frm54_initial_location
    Call Frm54_ClearAllField
    
    Frm54.Frame2.Visible = True

Else

    Frm54.Frame2.Visible = False
    
End If
End Sub
Private Sub Label26_Click()
'On Error Resume Next
If MDI_frm1.L20_Text = "Semua cawangan" Then

    Frm96.CMD2.Visible = True
    Frm96.CMD1.Visible = False

    Call Frm96_initial
    
    Frm96.Show vbModal
    
End If
    
If Frm54.Frame3.Visible = False Then
    
    Call frm54_initial_location
    Call Frm54_ClearAllField
    Call Frm54_call_setting_upah
    
    Frm54.Frame3.Visible = True

Else

    Frm54.Frame3.Visible = False
    
End If
End Sub
Private Sub MSFlexGrid1_DblClick()
'On Error Resume Next
PopupMenu Frm54_Menu
End Sub
Private Sub TB1_Change()
'On Error Resume Next
Dim a As Double 'Harga Semasa
Dim b As Double 'Diskaun Pelanggan
Dim c As Double 'Diskaun Member
Dim d As Double 'Diskaun Pengedar
Dim e As Double 'Diskaun RAF
Dim f As Double 'Diskaun Normal Dealer
Dim g As Double 'Diskaun Master Dealer

a = 0 'Harga Semasa
b = 0 'Diskaun Pelanggan
c = 0 'Diskaun Member
d = 0 'Diskaun Pengedar
e = 0 'Diskaun RAF
f = 0 'Diskaun Normal Dealer
g = 0 'Diskaun Master Dealer

If SearchDisable = 0 Then
    'Call TetapanHargaJualan
    
    If IsNumeric(Frm54.TB1) Then a = Frm54.TB1 'Harga Semasa
    If IsNumeric(Frm54.TB2) Then b = Frm54.TB2 'Diskaun Pelanggan
    If IsNumeric(Frm54.TB3) Then c = Frm54.TB3 'Diskaun Member
    If IsNumeric(Frm54.TB5) Then d = Frm54.TB5 'Diskaun Pengedar
    If IsNumeric(Frm54.TB4) Then e = Frm54.TB4 'Diskaun RAF
    If IsNumeric(Frm54.TB30) Then f = Frm54.TB30 'Diskaun Normal Dealer
    If IsNumeric(Frm54.TB31) Then g = Frm54.TB31 'Diskaun Master Dealer
    
    If a <> 0 Then
        Frm54.L3_Text = Format(a - b, "0.00") 'Harga Semasa Bagi Pelanggan
        Frm54.L4_Text = Format(a - c, "0.00") 'Harga Semasa Bagi Member
        Frm54.L6_Text = Format(a - d, "0.00") 'Harga Semasa Bagi Pengedar
        Frm54.L5_Text = Format(a - e, "0.00") 'Harga Semasa Bagi RAF
        Frm54.L13_Text = Format(a - f, "0.00") 'Harga Semasa Bagi Normal Dealer
        Frm54.L14_Text = Format(a - g, "0.00") 'Harga Semasa Bagi Master Dealer
    End If
End If
End Sub
Private Sub TB2_Change()
'On Error Resume Next
Dim a As Double
Dim b As Double

If IsNumeric(Frm54.TB1) And IsNumeric(Frm54.TB2) Then
    a = Frm54.TB1
    b = Frm54.TB2
    
    Frm54.L3_Text = Format((a - b), "0.00") 'Harga Jualan Bagi Pelanggan
Else
    Frm54.L3_Text = "XXX.XX"
End If
End Sub
Private Sub TB3_Change()
'On Error Resume Next
Dim a As Double
Dim b As Double

If IsNumeric(Frm54.TB1) And IsNumeric(Frm54.TB3) Then
    a = Frm54.TB1
    b = Frm54.TB3
    
    Frm54.L4_Text = Format((a - b), "0.00") 'Harga Jualan Bagi Member
Else
    Frm54.L4_Text = "XXX.XX"
End If
End Sub
Private Sub TB30_Change()
'On Error Resume Next
Dim a As Double
Dim b As Double

If IsNumeric(Frm54.TB1) And IsNumeric(Frm54.TB30) Then
    a = Frm54.TB1
    b = Frm54.TB30
    
    Frm54.L13_Text = Format((a - b), "0.00") 'Harga Jualan Bagi Normal Dealer
Else
    Frm54.L13_Text = "XXX.XX"
End If
End Sub
Private Sub TB31_Change()
'On Error Resume Next
Dim a As Double
Dim b As Double

If IsNumeric(Frm54.TB1) And IsNumeric(Frm54.TB31) Then
    a = Frm54.TB1
    b = Frm54.TB31
    
    Frm54.L14_Text = Format((a - b), "0.00") 'Harga Jualan Bagi Master Dealer
Else
    Frm54.L14_Text = "XXX.XX"
End If
End Sub
Private Sub TB4_Change()
'On Error Resume Next
Dim a As Double
Dim b As Double

If IsNumeric(Frm54.TB1) And IsNumeric(Frm54.TB4) Then
    a = Frm54.TB1
    b = Frm54.TB4
    
    Frm54.L5_Text = Format((a - b), "0.00") 'Harga Jualan Bagi RAF
Else
    Frm54.L5_Text = "XXX.XX"
End If
End Sub
Private Sub TB5_Change()
'On Error Resume Next
Dim a As Double
Dim b As Double

If IsNumeric(Frm54.TB1) And IsNumeric(Frm54.TB5) Then
    a = Frm54.TB1
    b = Frm54.TB5
    
    Frm54.L6_Text = Format((a - b), "0.00") 'Harga Jualan Bagi Pengedar
Else
    Frm54.L6_Text = "XXX.XX"
End If
End Sub

