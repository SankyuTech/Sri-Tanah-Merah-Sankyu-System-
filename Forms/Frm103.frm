VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm103 
   Caption         =   "Penyata Untung Rugi (Runcit)"
   ClientHeight    =   13035
   ClientLeft      =   120
   ClientTop       =   -11295
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
   Icon            =   "Frm103.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13035
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tetapan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   7695
      Begin VB.ComboBox CBB2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Frm103.frx":0ECA
         Left            =   1900
         List            =   "Frm103.frx":0ECC
         Style           =   2  'Dropdown List
         TabIndex        =   92
         Top             =   1440
         Width           =   5565
      End
      Begin VB.CommandButton CMD1 
         Caption         =   "Report"
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
         Left            =   2520
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm103.frx":0ECE
         MousePointer    =   99  'Custom
         Picture         =   "Frm103.frx":11D8
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2400
         Width           =   2865
      End
      Begin VB.ComboBox CBB1 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Frm103.frx":37A2
         Left            =   1900
         List            =   "Frm103.frx":37A4
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1800
         Width           =   5565
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   1905
         TabIndex        =   3
         Top             =   720
         Width           =   5550
         _ExtentX        =   9790
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
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   360
         Left            =   1905
         TabIndex        =   4
         Top             =   1080
         Width           =   5550
         _ExtentX        =   9790
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
      Begin VB.Label L35_Text 
         Caption         =   "L35_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   6480
         TabIndex        =   94
         Top             =   3120
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dulang * :"
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
         Height          =   300
         Left            =   360
         TabIndex        =   93
         Top             =   1470
         Width           =   1500
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm103.frx":37A6
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
         Height          =   1680
         Left            =   240
         TabIndex        =   10
         Top             =   3480
         Width           =   7290
      End
      Begin VB.Label L31_Text 
         Caption         =   "L31_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   6480
         TabIndex        =   14
         Top             =   2760
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label L6_Text 
         Caption         =   "L6_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   5640
         TabIndex        =   13
         Top             =   3120
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label L5_Text 
         Caption         =   "L5_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   5640
         TabIndex        =   12
         Top             =   2640
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Perhatian :"
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
         Height          =   240
         Left            =   240
         TabIndex        =   11
         Top             =   3120
         Width           =   8610
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cawangan * :"
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
         Height          =   300
         Left            =   360
         TabIndex        =   8
         Top             =   1830
         Width           =   1500
      End
      Begin VB.Label Label63 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Akhir * :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1125
         Width           =   1500
      End
      Begin VB.Label Label62 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Mula * :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   765
         Width           =   1500
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila buat pilihan tarikh bagi penyata untung rugi."
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
         Height          =   240
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   5010
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11415
      Left            =   4440
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   22455
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   10815
         Left            =   5760
         TabIndex        =   21
         Top             =   960
         Visible         =   0   'False
         Width           =   21015
         Begin VB.CommandButton CMD9 
            BackColor       =   &H000080FF&
            Caption         =   "Cetak senarai invoice jualan"
            Height          =   405
            Left            =   360
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm103.frx":3947
            MousePointer    =   99  'Custom
            TabIndex        =   78
            Top             =   8280
            Visible         =   0   'False
            Width           =   3225
         End
         Begin VB.CommandButton CMD10 
            BackColor       =   &H000080FF&
            Caption         =   "Cetak senarai modal dan harga jualan"
            Height          =   405
            Left            =   360
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm103.frx":3C51
            MousePointer    =   99  'Custom
            TabIndex        =   77
            Top             =   7080
            Width           =   3225
         End
         Begin VB.CommandButton CMD11 
            BackColor       =   &H000080FF&
            Caption         =   "Cetak ringkasan untung rugi"
            Height          =   405
            Left            =   360
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm103.frx":3F5B
            MousePointer    =   99  'Custom
            TabIndex        =   76
            Top             =   7680
            Width           =   3225
         End
         Begin VB.Label Label51 
            BackStyle       =   0  'Transparent
            Caption         =   "** Jumlah keuntungan bersih hasil dari jual beli emas dengan GST"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   6000
            TabIndex        =   89
            Top             =   6360
            Width           =   8055
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "RM"
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
            Height          =   375
            Left            =   3000
            TabIndex        =   88
            Top             =   6360
            Width           =   615
         End
         Begin VB.Label L34_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L34_Text"
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
            Height          =   375
            Left            =   3720
            TabIndex        =   87
            Top             =   6360
            Width           =   4095
         End
         Begin VB.Label Label40 
            BackStyle       =   0  'Transparent
            Caption         =   "Untung Bersih 2"
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
            Height          =   375
            Left            =   360
            TabIndex        =   86
            Top             =   6360
            Width           =   4095
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "RM"
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
            Height          =   375
            Left            =   3000
            TabIndex        =   75
            Top             =   5040
            Width           =   615
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Kos Bersih (Modal)"
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
            Height          =   375
            Left            =   360
            TabIndex        =   74
            Top             =   5070
            Width           =   4095
         End
         Begin VB.Label L22_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L22_Text"
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
            Height          =   375
            Left            =   3720
            TabIndex        =   71
            Top             =   5070
            Width           =   4095
         End
         Begin VB.Shape Shape3 
            BorderWidth     =   3
            Height          =   15
            Left            =   3720
            Top             =   5040
            Width           =   2100
         End
         Begin VB.Label L23_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L23_Text"
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
            Height          =   375
            Left            =   3720
            TabIndex        =   70
            Top             =   5550
            Visible         =   0   'False
            Width           =   4095
         End
         Begin VB.Label L24_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L24_Text"
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
            Height          =   375
            Left            =   3720
            TabIndex        =   69
            Top             =   6030
            Width           =   4095
         End
         Begin VB.Shape Shape4 
            BorderWidth     =   3
            Height          =   15
            Left            =   3720
            Top             =   5955
            Width           =   2100
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "- RM"
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
            Height          =   375
            Left            =   3000
            TabIndex        =   68
            Top             =   5550
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "RM"
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
            Height          =   375
            Left            =   3000
            TabIndex        =   67
            Top             =   6030
            Width           =   615
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "** Jumlah harga modal belian stok emas. Harga TIDAK termasuk GST."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   6000
            TabIndex        =   66
            Top             =   5070
            Width           =   8055
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "** Jumlah komisyen yang diberikan kepada agen dropship."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   6000
            TabIndex        =   65
            Top             =   5550
            Visible         =   0   'False
            Width           =   8055
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "** Jumlah keuntungan bersih hasil dari jual beli emas tanpa GST"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   6000
            TabIndex        =   64
            Top             =   6030
            Width           =   8055
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "RM"
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
            Height          =   375
            Left            =   3000
            TabIndex        =   63
            Top             =   4155
            Width           =   615
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "RM"
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
            Height          =   375
            Left            =   3000
            TabIndex        =   62
            Top             =   3600
            Width           =   615
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Jualan Bersih (Jualan)"
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
            Height          =   375
            Left            =   360
            TabIndex        =   61
            Top             =   3675
            Width           =   4095
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "GST Modal"
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
            Height          =   375
            Left            =   360
            TabIndex        =   60
            Top             =   4515
            Width           =   4095
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Kos Modal (Temasuk GST)"
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
            Height          =   375
            Left            =   360
            TabIndex        =   59
            Top             =   4155
            Width           =   4095
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   3
            Height          =   15
            Left            =   3720
            Top             =   3600
            Width           =   2100
         End
         Begin VB.Label L20_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L20_Text"
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
            Height          =   375
            Left            =   3720
            TabIndex        =   58
            Top             =   4155
            Width           =   4095
         End
         Begin VB.Label L21_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L21_Text"
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
            Height          =   375
            Left            =   3720
            TabIndex        =   57
            Top             =   4515
            Width           =   4095
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "- RM"
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
            Height          =   375
            Left            =   3000
            TabIndex        =   56
            Top             =   4515
            Width           =   615
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "** Jumlah harga belian stok emas. Harga termasuk GST."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   6000
            TabIndex        =   55
            Top             =   4155
            Width           =   8055
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "** Jumlah GST dari belian stok emas."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   6000
            TabIndex        =   54
            Top             =   4515
            Width           =   8055
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "** Jumlah harga jualan bersih terkumpul. (Setelah ditolak GST dan adjustment)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   6000
            TabIndex        =   53
            Top             =   3675
            Width           =   8055
         End
         Begin VB.Label L27_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L27_Text"
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
            Height          =   375
            Left            =   3720
            TabIndex        =   52
            Top             =   3675
            Width           =   4095
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "** Jumlah ADJUSTMENT(diskaun) terkumpul yang diberikan di dalam invoice jualan."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   6000
            TabIndex        =   51
            Top             =   2040
            Width           =   9615
         End
         Begin VB.Label L26_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L26_Text"
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
            Height          =   375
            Left            =   3720
            TabIndex        =   50
            Top             =   2040
            Width           =   4095
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "GST Jualan"
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
            Height          =   375
            Left            =   360
            TabIndex        =   49
            Top             =   1320
            Width           =   4095
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "- RM"
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
            Height          =   375
            Left            =   3000
            TabIndex        =   48
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label L18_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L18_Text"
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
            Height          =   375
            Left            =   3720
            TabIndex        =   47
            Top             =   1320
            Width           =   4095
         End
         Begin VB.Label L19_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L19_Text"
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
            Height          =   375
            Left            =   3720
            TabIndex        =   46
            Top             =   1680
            Width           =   4095
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "RM"
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
            Height          =   375
            Left            =   3000
            TabIndex        =   45
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "** Jumlah jualan emas kepada pelanggan. Harga adalah termasuk GST."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   6000
            TabIndex        =   44
            Top             =   960
            Width           =   8055
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "** Jumlah GST dari jualan emas kepada pelanggan."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   6000
            TabIndex        =   43
            Top             =   1320
            Width           =   8055
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "** Jumlah jualan emas kepada pelanggan. Harga TIDAK termasuk GST."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   6000
            TabIndex        =   42
            Top             =   1680
            Width           =   8055
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "- RM"
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
            Height          =   375
            Left            =   3000
            TabIndex        =   41
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "Adjustment"
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
            Height          =   375
            Left            =   360
            TabIndex        =   40
            Top             =   2040
            Width           =   2535
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "Jualan (Tanpa GST)"
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
            Height          =   375
            Left            =   360
            TabIndex        =   39
            Top             =   1680
            Width           =   2895
         End
         Begin VB.Label Label41 
            BackStyle       =   0  'Transparent
            Caption         =   "Diskaun"
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
            Height          =   375
            Left            =   360
            TabIndex        =   38
            Top             =   2400
            Width           =   2535
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "- RM"
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
            Height          =   375
            Left            =   3000
            TabIndex        =   37
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label L28_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L28_Text"
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
            Height          =   375
            Left            =   3720
            TabIndex        =   36
            Top             =   2400
            Width           =   4095
         End
         Begin VB.Label Label44 
            BackStyle       =   0  'Transparent
            Caption         =   "** Jumlah diskaun terkumpul yang diberikan di dalam semua invioce jualan."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   6000
            TabIndex        =   35
            Top             =   2400
            Width           =   9615
         End
         Begin VB.Label Label45 
            BackStyle       =   0  'Transparent
            Caption         =   "Kupon diskaun"
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
            Height          =   375
            Left            =   360
            TabIndex        =   34
            Top             =   2760
            Width           =   2535
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "- RM"
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
            Height          =   375
            Left            =   3000
            TabIndex        =   33
            Top             =   2760
            Width           =   615
         End
         Begin VB.Label L29_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L29_Text"
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
            Height          =   375
            Left            =   3720
            TabIndex        =   32
            Top             =   2760
            Width           =   4095
         End
         Begin VB.Label Label48 
            BackStyle       =   0  'Transparent
            Caption         =   "** Jumlah kupon diskaun yang digunakan oleh pelanggan."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   6000
            TabIndex        =   31
            Top             =   2760
            Width           =   9615
         End
         Begin VB.Label Label49 
            BackStyle       =   0  'Transparent
            Caption         =   "Tebus mata ganjaran"
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
            Height          =   375
            Left            =   360
            TabIndex        =   30
            Top             =   3120
            Width           =   2535
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "- RM"
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
            Height          =   375
            Left            =   3000
            TabIndex        =   29
            Top             =   3120
            Width           =   615
         End
         Begin VB.Label L30_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L30_Text"
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
            Height          =   375
            Left            =   3720
            TabIndex        =   28
            Top             =   3120
            Width           =   4095
         End
         Begin VB.Label Label52 
            BackStyle       =   0  'Transparent
            Caption         =   "** Jumlah nilaian mata ganjaran yang ditebus oleh pelanggan untuk membuat bayaran."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   6000
            TabIndex        =   27
            Top             =   3120
            Width           =   9615
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "RM"
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
            Height          =   375
            Left            =   3000
            TabIndex        =   26
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Jualan (Temasuk GST)"
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
            Height          =   375
            Left            =   360
            TabIndex        =   25
            Top             =   960
            Width           =   4095
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Ringkasan penyata."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   240
            TabIndex        =   24
            Top             =   240
            Width           =   4095
         End
         Begin VB.Label L17_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L17_Text"
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
            Height          =   375
            Left            =   3720
            TabIndex        =   23
            Top             =   960
            Width           =   4095
         End
         Begin VB.Label L25_Text 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   360
            TabIndex        =   22
            Top             =   600
            Width           =   12255
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Komisyen Agen Dropship"
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
            Height          =   375
            Left            =   360
            TabIndex        =   73
            Top             =   5550
            Visible         =   0   'False
            Width           =   4095
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Untung Bersih 1"
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
            Height          =   375
            Left            =   360
            TabIndex        =   72
            Top             =   6030
            Width           =   4095
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   10815
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Visible         =   0   'False
         Width           =   21015
         Begin VB.CommandButton CMD8 
            Caption         =   "Next"
            Height          =   810
            Left            =   16440
            MouseIcon       =   "Frm103.frx":4265
            MousePointer    =   99  'Custom
            Picture         =   "Frm103.frx":456F
            Style           =   1  'Graphical
            TabIndex        =   91
            ToolTipText     =   "Paparan Seterusnya"
            Top             =   9480
            Width           =   1095
         End
         Begin VB.CommandButton CMD7 
            Caption         =   "Back"
            Height          =   810
            Left            =   15240
            MouseIcon       =   "Frm103.frx":5639
            MousePointer    =   99  'Custom
            Picture         =   "Frm103.frx":5943
            Style           =   1  'Graphical
            TabIndex        =   90
            ToolTipText     =   "Paparan Sebelum"
            Top             =   9480
            Width           =   1095
         End
         Begin MSComctlLib.ListView LV1 
            Height          =   10170
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Width           =   15015
            _ExtentX        =   26485
            _ExtentY        =   17939
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
         Begin VB.Label L16_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L16_Text"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   16950
            TabIndex        =   84
            Top             =   10400
            Width           =   375
         End
         Begin VB.Label L15_Text 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "L15_Text"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   16440
            TabIndex        =   83
            Top             =   10400
            Width           =   375
         End
         Begin VB.Label L14_Text 
            Caption         =   "L14_Text"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   18720
            TabIndex        =   82
            Top             =   5040
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label L13_Text 
            Caption         =   "L13_Text"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   17640
            TabIndex        =   81
            Top             =   5040
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   $"Frm103.frx":6A0D
            ForeColor       =   &H00000000&
            Height          =   1455
            Left            =   15360
            TabIndex        =   80
            Top             =   840
            Width           =   5715
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "Perhatian :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   15360
            TabIndex        =   79
            Top             =   480
            Width           =   4995
         End
         Begin VB.Label L8_Text 
            BackStyle       =   0  'Transparent
            Caption         =   "L8_Text"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   240
            Width           =   12255
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Paparan muka :           /"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   15240
            TabIndex        =   85
            Top             =   10400
            Width           =   2415
         End
      End
      Begin VB.Label L33_Text 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Ringkasan"
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
         Left            =   2280
         MouseIcon       =   "Frm103.frx":6C1D
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label L32_Text 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Senarai Report"
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
         Left            =   240
         MouseIcon       =   "Frm103.frx":6F27
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Label L4_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tetapan Penyata"
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
      MouseIcon       =   "Frm103.frx":7231
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
   Begin VB.Menu Frm103_PM_Menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm102_SM_excel1 
         Caption         =   "Export ke excel"
      End
   End
   Begin VB.Menu Frm103_PM_Menu2 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm102_SM_excel2 
         Caption         =   "Export ke excel"
      End
   End
End
Attribute VB_Name = "Frm103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD1_Click()
'on error resume next
Frm103.L5_Text = Frm103.DTPicker1 'Tarikh mula
Frm103.L6_Text = Frm103.DTPicker2 'Tarikh akhir
Frm103.L31_Text = Frm103.CBB1 'Cawangan
Frm103.L35_Text = Frm103.CBB2 'Dulang

If Frm103.CBB1 = vbNullString Then

    MsgBox "Sila buat pilihan cawangan.", vbExclamation, "Info"
    
    Exit Sub

End If

If Frm103.CBB2 = vbNullString Then

    MsgBox "Sila buat pilihan dulang.", vbExclamation, "Info"
    
    Exit Sub

End If

Note = "Sistem akan mengambil sedikit masa untuk mengeluarkan penyata untung rugi." & vbCrLf & _
        "" & vbCrLf & _
        "Sila tunggu sehingga sistem selasai melakukan pengiraan." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    
    Frm103.L13_Text = -1
    Frm103.L15_Text = 0
    Frm103.L14_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    GM_NEXT_PREV = 0
    
    Call Frm103_senarai_modal_jualan_header
    Call Frm103_senarai_modal_jualan
    Call Frm103_kira_untung_rugi
    
    If Frm103.L15_Text <> "0" Then
    
        Frm103.Frame2.Visible = True
        Frm103.Frame3.Visible = True
        Frm103.Frame1.Visible = False
    
    Else
        
        MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    
    End If
    'Call Frm103_reset_penyata
    'Call Frm103_kiraan_untung_rugi
    'Call Frm103_initial_setting2
    
End If
End Sub
Private Sub CMD10_Click()
'on error resume next
Call Frm103_cetak_penyata_modal_jual
End Sub
Private Sub CMD11_Click()
'on error resume next
Call Frm103_cetak_penyata_untung_rugi
End Sub
Private Sub CMD7_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm103_senarai_modal_jualan_header
Call Frm103_senarai_modal_jualan
End Sub
Private Sub CMD8_Click()
'on error resume next
Dim LM_CURRENT_PAGE As Single
Dim LM_PAGE_QTY As Single

LM_CURRENT_PAGE = 0
LM_PAGE_QTY = 0

If Frm103.L15_Text <> vbNullString And IsNumeric(Frm103.L15_Text) Then
    If Frm103.L16_Text <> vbNullString And IsNumeric(Frm103.L16_Text) Then
        LM_PAGE_QTY = Frm103.L16_Text
        LM_CURRENT_PAGE = Frm103.L15_Text
        
        If LM_CURRENT_PAGE < LM_PAGE_QTY Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm103_senarai_modal_jualan_header
            Call Frm103_senarai_modal_jualan
            
        End If
    End If
End If
End Sub
Private Sub CMD9_Click()
'on error resume next
Call Frm103_cetak_penyata_invoice
End Sub
Private Sub Form_Load()
'on error resume next
Frm103.Picture = MDI_frm1.Picture
Frm103.DTPicker1 = DateTime.Date
Frm103.DTPicker2 = DateTime.Date
End Sub
Private Sub Frm103_SM_excel2_Click()
'on error resume next
Call Frm103_senarai_modal_jual_excel
End Sub
Private Sub Frm102_SM_excel2_Click()
'on error resume next
Call Frm103_senarai_modal_jual_excel
End Sub

Private Sub L17_Text_Change()
'on error resume next
Call Frm103_untung_rugi_summary
End Sub

Private Sub L18_Text_Change()
'on error resume next
Call Frm103_untung_rugi_summary
End Sub

Private Sub L19_Text_Change()
'on error resume next
Call Frm103_untung_rugi_summary
End Sub

Private Sub L20_Text_Change()
'on error resume next
Call Frm103_untung_rugi_summary
End Sub



Private Sub L21_Text_Change()
'on error resume next
Call Frm103_untung_rugi_summary
End Sub



Private Sub L22_Text_Change()
'on error resume next
Call Frm103_untung_rugi_summary
End Sub
Private Sub L23_Text_Change()
'on error resume next
Call Frm103_untung_rugi_summary
End Sub

Private Sub L26_Text_Change()
'on error resume next
Call Frm103_untung_rugi_summary
End Sub
Private Sub L3_Text_Click()
'on error resume next
Frm34.Show
Unload Frm103
End Sub
Private Sub L27_Text_Change()
'on error resume next
Call Frm103_untung_rugi_summary
End Sub

Private Sub L28_Text_Change()
'on error resume next
Call Frm103_untung_rugi_summary
End Sub


Private Sub L29_Text_Change()
'on error resume next
Call Frm103_untung_rugi_summary
End Sub


Private Sub L30_Text_Change()
'on error resume next
Call Frm103_untung_rugi_summary
End Sub

Private Sub L32_Text_Click()
'on error resume next
If Frm103.Frame3.Visible = False Then

    Call Frm103_initial_setting2
    
    'Frm103.L13_Text = -1
    'Frm103.L15_Text = 0
    'Frm103.L14_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    'GM_NEXT_PREV = 0
    
    'Call Frm103_senarai_modal_jualan_header
    'Call Frm103_senarai_modal_jualan
    
    Frm103.Frame3.Visible = True
    
Else

    Frm103.Frame3.Visible = False
    
End If
End Sub

Private Sub L33_Text_Click()
'on error resume next
If Frm103.Frame4.Visible = False Then

    Call Frm103_initial_setting2
    
    Frm103.Frame4.Visible = True
    
Else

    Frm103.Frame4.Visible = False
    
End If
End Sub

Private Sub L4_Text_Click()
'on error resume next
If Frm103.Frame1.Visible = False Then

    Call Frm103_initial_setting
    
    Frm103.Frame1.Visible = True
    
Else

    Frm103.Frame1.Visible = False
    
End If
End Sub



Private Sub LV1_DblClick()
'on error resume next
frm103_LM_No_ID = vbNullString

If IsNumeric(Frm103.LV1.SelectedItem.Index) Then
    
    frm103_LM_No_ID = Frm103.LV1.SelectedItem.Index
    
    If frm103_LM_No_ID <> vbNullString Then
    
    
        PopupMenu Frm103_PM_Menu2
    
    Else
    
        MsgBox "Tiada Data.", vbInformation, "Info"
        
    End If
    
Else

    MsgBox "Tiada Data.", vbInformation, "Info"
    
End If
End Sub

