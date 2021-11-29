VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDI_frm1 
   BackColor       =   &H8000000C&
   Caption         =   "SPKE 106.1.20"
   ClientHeight    =   13035
   ClientLeft      =   225
   ClientTop       =   270
   ClientWidth     =   23880
   Icon            =   "MDI_frm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDI_frm1.frx":0ECA
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Tmr3 
      Interval        =   50
      Left            =   4680
      Top             =   1320
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   1140
      Left            =   0
      TabIndex        =   72
      Top             =   0
      Width           =   23880
      _ExtentX        =   42122
      _ExtentY        =   2011
      ButtonWidth     =   2328
      ButtonHeight    =   1852
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "TERIMA STOK"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "URUSAN KEDAI"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "REPORT"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ADMIN"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "LOG OUT"
            ImageIndex      =   5
         EndProperty
      EndProperty
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   9000
         ScaleHeight     =   1095
         ScaleWidth      =   14895
         TabIndex        =   76
         Top             =   0
         Width           =   14895
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   375
            Left            =   5040
            TabIndex        =   84
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            Height          =   615
            Left            =   2880
            TabIndex        =   83
            Top             =   240
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   615
            Left            =   1800
            TabIndex        =   82
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label L17_Text 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "ONLINE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   0
            TabIndex        =   77
            Top             =   720
            Width           =   1815
         End
         Begin VB.Image Image3 
            Height          =   720
            Left            =   480
            Picture         =   "MDI_frm1.frx":7D070
            Stretch         =   -1  'True
            Top             =   0
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Image Image1 
            Height          =   720
            Left            =   600
            Picture         =   "MDI_frm1.frx":7F63A
            Stretch         =   -1  'True
            Top             =   0
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label L9_Text 
            Alignment       =   2  'Center
            Caption         =   "SANKYU SYSTEM"
            BeginProperty Font 
               Name            =   "Script MT Bold"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1215
            Left            =   120
            TabIndex        =   78
            Top             =   -120
            Width           =   14745
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1605
         Left            =   6600
         Picture         =   "MDI_frm1.frx":8FE84
         ScaleHeight     =   1605
         ScaleWidth      =   2415
         TabIndex        =   73
         Top             =   0
         Width           =   2415
         Begin VB.Label amination 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "======================= Sankyu System  sankyusystem@gmail.com +6010 - 900 4788 #SankyuSystem"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1695
            Index           =   1
            Left            =   120
            TabIndex        =   75
            Top             =   480
            Width           =   2175
            WordWrap        =   -1  'True
         End
         Begin VB.Label amination 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "======================= Sankyu System  sankyusystem@gmail.com +6010 - 900 4788 #SankyuSystem"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1695
            Index           =   0
            Left            =   120
            TabIndex        =   74
            Top             =   -480
            Width           =   2175
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   0
         Picture         =   "MDI_frm1.frx":928CF
         Stretch         =   -1  'True
         Top             =   0
         Width           =   840
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   720
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":94E99
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":97473
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":99A4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":9C027
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":9E601
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_frm1.frx":A0BDB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Tmr2 
      Interval        =   50
      Left            =   4200
      Top             =   1320
   End
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   3720
      Top             =   1320
   End
   Begin VB.Timer Tmr1 
      Interval        =   10
      Left            =   3240
      Top             =   1320
   End
   Begin VB.PictureBox Pic1 
      Align           =   4  'Align Right
      BackColor       =   &H00C0FFFF&
      Height          =   11895
      Left            =   21465
      ScaleHeight     =   11835
      ScaleWidth      =   2355
      TabIndex        =   0
      Top             =   1140
      Width           =   2415
      Begin VB.CommandButton CMD44 
         Caption         =   "Tukar Cawangan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         Picture         =   "MDI_frm1.frx":A31B5
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   1680
         Width           =   2100
      End
      Begin VB.PictureBox Pic4 
         BackColor       =   &H00FFFFFF&
         Height          =   7000
         Left            =   2280
         ScaleHeight     =   6945
         ScaleWidth      =   1995
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton CMD43 
            BackColor       =   &H80000016&
            Caption         =   "REPORT TRADE IN / POTONG"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":AC67F
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   1000
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD42 
            BackColor       =   &H80000016&
            Caption         =   "LOG"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":AC989
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   5830
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD41 
            BackColor       =   &H80000016&
            Caption         =   "BARANG HILANG"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":ACC93
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   5280
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD40 
            BackColor       =   &H80000016&
            Caption         =   "REPORT TRADE IN"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":ACF9D
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   6360
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton CMD34 
            BackColor       =   &H80000016&
            Caption         =   "GRN / GDN / INV / VOU"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":AD2A7
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   1530
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD29 
            BackColor       =   &H80000016&
            Caption         =   "SENARAI INVOICE"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":AD5B1
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Senarai pengeluaran invoice"
            Top             =   2580
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD19 
            BackColor       =   &H80000016&
            Caption         =   "INVENTORI"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":AD8BB
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   4180
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD17 
            BackColor       =   &H80000016&
            Caption         =   "PENYATA UNTUNG RUGI (Runcit)"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":ADBC5
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   3650
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD16 
            BackColor       =   &H80000016&
            Caption         =   "REPORT GST"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":ADECF
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   4720
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD15 
            BackColor       =   &H80000016&
            Caption         =   "PENYATA UNTUNG RUGI (Restock)"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":AE1D9
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   3120
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD14 
            BackColor       =   &H80000016&
            Caption         =   "REPORT KEWANGAN"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":AE4E3
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   2060
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD7 
            BackColor       =   &H80000016&
            Caption         =   "REPORT KESELURUHAN"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":AE7ED
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   480
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "REPORT"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   615
            Left            =   0
            TabIndex        =   13
            Top             =   120
            Width           =   1935
         End
      End
      Begin VB.PictureBox Pic5 
         BackColor       =   &H00FFFFFF&
         Height          =   7000
         Left            =   2280
         ScaleHeight     =   6945
         ScaleWidth      =   1995
         TabIndex        =   21
         Top             =   840
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton CMD31 
            BackColor       =   &H80000016&
            Caption         =   "SETTING PRINTER"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   550
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":AEAF7
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   6120
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton CMD30 
            BackColor       =   &H80000016&
            Caption         =   "TETAPAN SISTEM"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   550
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":AEE01
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   1920
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD28 
            BackColor       =   &H80000016&
            Caption         =   "ANALISA HARGA EMAS"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   550
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":AF10B
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   4320
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD25 
            BackColor       =   &H80000016&
            Caption         =   "BACKUP DATABASE"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   550
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":AF415
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   4920
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD23 
            BackColor       =   &H80000016&
            Caption         =   "PAYROLL"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   550
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":AF71F
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   3720
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD18 
            BackColor       =   &H80000016&
            Caption         =   "DATA PEKERJA"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   550
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":AFA29
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   2520
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD11 
            BackColor       =   &H80000016&
            Caption         =   "TETAPAN BARCODE"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   550
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":AFD33
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   3120
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD9 
            BackColor       =   &H80000016&
            Caption         =   "TETAPAN HARGA JUALAN EMAS"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   550
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":B003D
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   1320
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD8 
            BackColor       =   &H80000016&
            Caption         =   "TETAPAN ASAS SISTEM"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   550
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":B0347
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   720
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TETAPAN SISTEM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   615
            Left            =   0
            TabIndex        =   23
            Top             =   120
            Width           =   1935
         End
      End
      Begin VB.PictureBox Pic3 
         BackColor       =   &H00FFFFFF&
         Height          =   7000
         Left            =   2040
         ScaleHeight     =   6945
         ScaleWidth      =   1995
         TabIndex        =   10
         Top             =   1440
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton CMD39 
            BackColor       =   &H80000016&
            Caption         =   "GOODS DESPATCH NOTE (BULK)"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":B0651
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   5880
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton CMD36 
            BackColor       =   &H80000016&
            Caption         =   "TRADE IN"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":B095B
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   1080
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD35 
            BackColor       =   &H80000016&
            Caption         =   "INVOICE / VOUCHER"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":B0C65
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   5520
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD33 
            BackColor       =   &H80000016&
            Caption         =   "GOODS RECEIVED NOTE"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":B0F6F
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   4965
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD32 
            BackColor       =   &H80000016&
            Caption         =   "GOODS DESPATCH NOTE (PER ITEM)"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":B1279
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   4425
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD27 
            BackColor       =   &H80000016&
            Caption         =   "AGIHAN STOK KE CAWANGAN / AGEN"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":B1583
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   3840
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD26 
            BackColor       =   &H80000016&
            Caption         =   "FORMING OUT"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   0
            MouseIcon       =   "MDI_frm1.frx":B188D
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   6480
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton CMD24 
            BackColor       =   &H80000016&
            Caption         =   "E-MAIL PROMOSI"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":B1B97
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   3480
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD22 
            BackColor       =   &H80000016&
            Caption         =   "PENGURUSAN BUKU CEK"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":B1EA1
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   2910
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD21 
            BackColor       =   &H80000016&
            Caption         =   "PENGELUARAN / KEMASUKAN TUNAI"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":B21AB
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   2350
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD20 
            BackColor       =   &H80000016&
            Caption         =   "TEMPAHAN"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":B24B5
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   720
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD13 
            BackColor       =   &H80000016&
            Caption         =   "ANSURAN"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   1200
            MouseIcon       =   "MDI_frm1.frx":B27BF
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   6600
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton CMD12 
            BackColor       =   &H80000016&
            Caption         =   "SERVIS && BELANJA"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":B2AC9
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   1440
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD10 
            BackColor       =   &H80000016&
            Caption         =   "MAKLUMAT PELANGGAN"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":B2DD3
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   1800
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD6 
            BackColor       =   &H80000016&
            Caption         =   "JUALAN KEPADA AGEN"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   -240
            MouseIcon       =   "MDI_frm1.frx":B30DD
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   6600
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton CMD5 
            BackColor       =   &H80000016&
            Caption         =   "JUALAN"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":B33E7
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "URUSAN KEDAI"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   615
            Left            =   360
            TabIndex        =   11
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.PictureBox Pic2 
         BackColor       =   &H00FFFFFF&
         Height          =   7000
         Left            =   2040
         ScaleHeight     =   6945
         ScaleWidth      =   1995
         TabIndex        =   7
         Top             =   1440
         Visible         =   0   'False
         Width           =   2055
         Begin VB.CommandButton CMD3 
            BackColor       =   &H80000016&
            Caption         =   "BELIAN EMAS TERPAKAI (BARANG KEMAS)"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":B36F1
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   1680
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton CMD2 
            BackColor       =   &H80000016&
            Caption         =   "PENERIMAAN STOK GOLD BAR (BARU)"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":B39FB
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   2640
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton CMD1 
            BackColor       =   &H80000016&
            Caption         =   "PENERIMAAN STOK BARANG KEMAS (BARU)"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            MouseIcon       =   "MDI_frm1.frx":B3D05
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   840
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "PENERIMAAN STOK"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   615
            Left            =   0
            TabIndex        =   8
            Top             =   120
            Width           =   1935
         End
      End
      Begin VB.Label L22_Text 
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
         Height          =   255
         Left            =   240
         TabIndex        =   80
         Top             =   3600
         Visible         =   0   'False
         Width           =   495
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
         Height          =   495
         Left            =   840
         TabIndex        =   79
         Top             =   990
         Width           =   1365
      End
      Begin VB.Label L19_Text 
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
         Height          =   255
         Left            =   960
         TabIndex        =   71
         Top             =   3600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label L18_Text 
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
         Height          =   255
         Left            =   1560
         TabIndex        =   70
         Top             =   3600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label L16_Text 
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
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   960
         TabIndex        =   59
         Top             =   8040
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label L15_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Kalkulator trade in"
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         MouseIcon       =   "MDI_frm1.frx":B400F
         MousePointer    =   99  'Custom
         TabIndex        =   58
         ToolTipText     =   "Kalkulator"
         Top             =   9885
         Width           =   2175
      End
      Begin VB.Label L14_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Note pad"
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         MouseIcon       =   "MDI_frm1.frx":B4319
         MousePointer    =   99  'Custom
         TabIndex        =   50
         ToolTipText     =   "Note pad"
         Top             =   10125
         Width           =   2295
      End
      Begin VB.Label L13_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Kalkulator"
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         MouseIcon       =   "MDI_frm1.frx":B4623
         MousePointer    =   99  'Custom
         TabIndex        =   49
         ToolTipText     =   "Kalkulator"
         Top             =   9615
         Width           =   2175
      End
      Begin VB.Label L12_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Tukar user"
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   0
         MouseIcon       =   "MDI_frm1.frx":B492D
         MousePointer    =   99  'Custom
         TabIndex        =   48
         ToolTipText     =   "Tukar password"
         Top             =   7680
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label L11_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Maklumat sistem"
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         MouseIcon       =   "MDI_frm1.frx":B4C37
         MousePointer    =   99  'Custom
         TabIndex        =   47
         ToolTipText     =   "Maklumat sistem"
         Top             =   10365
         Width           =   2295
      End
      Begin VB.Label L10_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Tukar password"
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         MouseIcon       =   "MDI_frm1.frx":B4F41
         MousePointer    =   99  'Custom
         TabIndex        =   46
         ToolTipText     =   "Tukar password"
         Top             =   9350
         Width           =   2300
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   3
         Height          =   735
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   10680
         Width           =   2295
      End
      Begin VB.Label Label36 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Sankyu System"
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   11040
         Width           =   2055
      End
      Begin VB.Label Label35 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Powered By"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   10800
         Width           =   1335
      End
      Begin VB.Label L8_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Tetapan harga emas"
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         MouseIcon       =   "MDI_frm1.frx":B524B
         MousePointer    =   99  'Custom
         TabIndex        =   42
         ToolTipText     =   "Tetapan harga emas"
         Top             =   9100
         Width           =   2300
      End
      Begin VB.Label L7_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Jualan"
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         MouseIcon       =   "MDI_frm1.frx":B5555
         MousePointer    =   99  'Custom
         TabIndex        =   41
         ToolTipText     =   "Menu jualan"
         Top             =   8850
         Width           =   2300
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Quick Launcher"
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         MouseIcon       =   "MDI_frm1.frx":B585F
         MousePointer    =   99  'Custom
         TabIndex        =   40
         Top             =   8280
         Width           =   1935
      End
      Begin VB.Label L6_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Maklumat Pelanggan"
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         MouseIcon       =   "MDI_frm1.frx":B5B69
         MousePointer    =   99  'Custom
         TabIndex        =   14
         ToolTipText     =   "Maklumat pelanggan"
         Top             =   8600
         Width           =   2300
      End
      Begin VB.Label L5_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L5_Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   8040
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         Height          =   1035
         Left            =   30
         Shape           =   4  'Rounded Rectangle
         Top             =   540
         Width           =   2895
      End
      Begin VB.Label L4_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L4_Text"
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
         Left            =   840
         TabIndex        =   6
         Top             =   810
         Width           =   2025
      End
      Begin VB.Label L2_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L2_Text"
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
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label L1_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L1_Text"
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
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   2775
      End
      Begin VB.Label L3_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L3_Text"
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
         Left            =   840
         TabIndex        =   3
         Top             =   610
         Width           =   2025
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   ":    :     :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   720
         TabIndex        =   2
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User        Level  Branch"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label37 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   30
         TabIndex        =   45
         Top             =   10680
         Width           =   2295
      End
   End
   Begin ComctlLib.ImageList ImageList4 
      Left            =   1200
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI_frm1.frx":B5E73
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1800
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI_frm1.frx":B668D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MDI_PM_menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu MDI_SM_tukar_background 
         Caption         =   "&Tukar warna background sistem"
         Begin VB.Menu MD1_SM_SM_1 
            Caption         =   "1"
         End
         Begin VB.Menu MD1_SM_SM_2 
            Caption         =   "2"
         End
         Begin VB.Menu MD1_SM_SM_3 
            Caption         =   "3"
         End
         Begin VB.Menu MD1_SM_SM_4 
            Caption         =   "4"
         End
         Begin VB.Menu MD1_SM_SM_5 
            Caption         =   "5"
         End
         Begin VB.Menu MD1_SM_SM_6 
            Caption         =   "6"
         End
         Begin VB.Menu MD1_SM_SM_7 
            Caption         =   "7"
         End
      End
   End
End
Attribute VB_Name = "MDI_frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xLogOff As Boolean
Private Sub CMD1_Click()
'on error resume next
If MDI_frm1.L20_Text = "Semua cawangan" Then

    Frm96.CMD2.Visible = True
    Frm96.CMD1.Visible = False

    Call Frm96_initial
        
    Frm96.Show vbModal
    
End If

Call MDI_frm1_unload_all_menu

'Call Frm96_background_color
'Call Frm96_initial
MDI_frm1.L5_Text = 1

Frm83.CB7 = 1 'Barang baru
Frm83.CB8 = 0 'Used gold

'ID :
    '1 : Stock In -> Penerimaan stok baru (Barang kemas & permata)
    '2 : Stock In -> Penerimaan stok baru (Gold bar)
    
If MDI_frm1.L5_Text = 1 Then
    Frm83.CB9 = 1 'Barang kemas / permata
    Frm83.CB10 = 0 'Gold bar
ElseIf MDI_frm1.L5_Text = 2 Then
    Frm83.CB9 = 0 'Barang kemas / permata
    Frm83.CB10 = 1 'Gold bar
    
    Frm83.CB14 = 0
    Frm83.CB15 = 0
    
    Frm83.CB14.Enabled = False
    Frm83.CB15.Enabled = False
End If

Call Frm83_background_color
Call Frm83_form_load
Call frm83_flag_barang_baru

If MDI_frm1.L5_Text = 2 Then

    Call Frm83_mode_gold_bar
    
End If

Frm83.Show
End Sub
Private Sub CMD10_Click()
'on error resume next
Call Frm68_background_color
Call MDI_frm1_unload_all_menu
Call Frm68_background_color
MDI_frm1.L5_Text = 11

Frm68.Show
Frm68.L36_Text = 0 '0 : Terus dari menu data pelanggan , 1 : Data pembeli , 2 : Data agen dropship
End Sub
Private Sub CMD11_Click()
'On Error Resume Next
Call MDI_frm1_unload_all_menu
Frm56x.Picture = MDI_frm1.Picture

Frm56.Show
MDI_frm1.L5_Text = 20
End Sub
Private Sub CMD12_Click()
'on error resume next
Call MDI_frm1_unload_all_menu
Call Frm92_background_color
MDI_frm1.L5_Text = 10
Frm92.Show
End Sub
Private Sub CMD13_Click()
'on error resume next
Call MDI_frm1_unload_all_menu
Call Frm87_background_color
MDI_frm1.L5_Text = 7
Frm87.Show
End Sub

Private Sub CMD14_Click()
'On Error Resume Next
Call MDI_frm1_unload_all_menu
Call Frm105_background_color
Frm106.Picture = MDI_frm1.Picture

Frm105.Show
MDI_frm1.L5_Text = 13
End Sub
Private Sub CMD15_Click()
'On Error Resume Next
Call MDI_frm1_unload_all_menu
Call Frm104_background_color

Frm104.Show
MDI_frm1.L5_Text = 14
End Sub
Private Sub CMD16_Click()
'On Error Resume Next
Call MDI_frm1_unload_all_menu
Call Frm75_background_color

Frm75.L17_Text = "0.00"
Frm75.L18_Text = "0.00"
Frm75.L19_Text = vbNullString

Frm75.L69_Text = -1 'Titik Pencarian Data
Frm75.L75_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Frm75.L67_Text = 0 'Paparan Page ke-xxx
Frm75.L68_Text = 0

Frm75.L62_Text = -1 'Start Point
Frm75.L60_Text = 0 'Current Page
Frm75.L61_Text = 0 'Current Page
Frm75.L63_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    
Frm75.Show
MDI_frm1.L5_Text = 15
End Sub
Private Sub CMD17_Click()
'On Error Resume Next
Call MDI_frm1_unload_all_menu
Call Frm103_background_color

Frm103.Show
MDI_frm1.L5_Text = 15
End Sub
Private Sub CMD18_Click()
'On Error Resume Next
Call MDI_frm1_unload_all_menu
Call Frm49_background_color
Call frm49_Default

Frm49.Show
MDI_frm1.L5_Text = 17
End Sub
Private Sub CMD19_Click()
'on error resume next
Call MDI_frm1_unload_all_menu
Call Frm57_background_color

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If rs!SenaraiDulang <> vbNullString And rs!Status = 1 Then
        Frm57.CBB1.AddItem rs!SenaraiDulang 'Senarai Dulang
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Call Frm57_M_Clear

Frm57.Show
Frm57.Pic1.Visible = True

MDI_frm1.L5_Text = 21
End Sub
Private Sub CMD2_Click()
'on error resume next
Call MDI_frm1_unload_all_menu
'Call Frm96_background_color
'Call Frm96_initial

MDI_frm1.L5_Text = 2
'Frm96.Show

Frm83.CB7 = 1 'Barang baru
Frm83.CB8 = 0 'Used gold

If MDI_frm1.L5_Text = 1 Then
    Frm83.CB9 = 1 'Barang kemas / permata
    Frm83.CB10 = 0 'Gold bar
ElseIf MDI_frm1.L5_Text = 2 Then
    Frm83.CB9 = 0 'Barang kemas / permata
    Frm83.CB10 = 1 'Gold bar
    
    Frm83.CB14 = 0
    Frm83.CB15 = 0
    
    Frm83.CB14.Enabled = False
    Frm83.CB15.Enabled = False
End If

Call Frm83_background_color
Call Frm83_form_load

If MDI_frm1.L5_Text = 2 Then

    Call Frm83_mode_gold_bar
    
End If
End Sub
Private Sub CMD20_Click()
'on error resume next
Call MDI_frm1_unload_all_menu
Call Frm93_background_color
MDI_frm1.L5_Text = 8
Frm93.Show
End Sub
Private Sub CMD21_Click()
'On Error Resume Next
Call MDI_frm1_unload_all_menu
Call Frm100_background_color

Frm100.Show
MDI_frm1.L5_Text = 22
End Sub
Private Sub CMD22_Click()
'On Error Resume Next
Call MDI_frm1_unload_all_menu
Call Frm86_background_color

Call Frm86_Initial_Setting
Frm86.Show
MDI_frm1.L5_Text = 24
End Sub
Private Sub CMD23_Click()
'on error resume next
Call MDI_frm1_unload_all_menu
Call Frm48_background_color

Call Frm48_Default
Frm48.Show
MDI_frm1.L5_Text = 23
End Sub
Private Sub CMD24_Click()
'On Error Resume Next
Call MDI_frm1_unload_all_menu
Call Frm97_background_color

Frm97.Show
MDI_frm1.L5_Text = 25
End Sub
Private Sub CMD25_Click()
'on error resume next
Note = "Adakah anda ingin backup database ini?" & vbCrLf & _
        "Sistem tidak dapat beroperasi semasa database dibackup sehingga selesai." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then
        
    Call backup_database
    
End If
End Sub
Private Sub CMD26_Click()
'on error resume next
Call MDI_frm1_unload_all_menu
Call Frm107_background_color

Frm107.Show

MDI_frm1.L5_Text = 27
End Sub
Private Sub CMD27_Click()
'on error resume next
Call MDI_frm1_unload_all_menu
Call Frm108_background_color

Frm108.Show
MDI_frm1.L5_Text = 28
End Sub
Private Sub CMD28_Click()
'on error resume next
Call MDI_frm1_unload_all_menu
Frm109.Show
MDI_frm1.L5_Text = 29
End Sub
Private Sub CMD29_Click()
'On Error Resume Next
Call MDI_frm1_unload_all_menu
Call Frm110_background_color

Frm110.Show
MDI_frm1.L5_Text = 30
End Sub

Private Sub CMD30_Click()
'On Error Resume Next
Call MDI_frm1_unload_all_menu
Call Frm111_background_color

Call Frm111_initial_setting
Call Frm111_setting

Frm111.Show
MDI_frm1.L5_Text = 30
End Sub
Private Sub CMD31_Click()
'on error resume next
Call MDI_frm1_unload_all_menu
Frm112.Picture = MDI_frm1.Picture

Frm112.Show
MDI_frm1.L5_Text = 31
End Sub

Private Sub CMD32_Click()
'on error resume next
Note = "Menu ini adalah dikhususkan bagi hantaran barang/tukaran barang dengan pihak supplier atau agen." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Setiap barang yang ditukarkan perlu dipilih SATU PER SATU dari senarai atau SCAN barang tersebut." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If

If Answer = vbYes Then

    Call MDI_frm1_unload_all_menu
    Call Frm115_background_color
    
    GLOBAL_DISABLE = 0
    Frm115.TB1 = vbNullString
    
    Call Frm115_reset_1
    Call Frm115_reset_2
    Call Frm115_reset_3
    Call Frm115_reset_main
    Call Frm28_initial
    Call Frm115_reset_main2
    
    Frm115.DTPicker1 = DateTime.Date$
    MDI_frm1.L5_Text = 16
    
    Frm115.L69_Text = -1 'Titik Pencarian Data
    Frm115.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm115.L67_Text = 0 'Paparan Page ke-xxx
    Frm115.L68_Text = 0
    
    Call frm115_initial_setting_stok
    Call Frm115_Senarai_Jualan_Header
    Call frm115_reset_gdn_list
    
    Frm115.CMD8.Visible = True
    Frm115.CMD9.Visible = True
    Frm115.CMD10.Visible = False
    Frm115.CMD11.Visible = False
    
    Frm115.Picture = MDI_frm1.Picture
    Frm115.Show
    
    Frm115.L32_Text = 0 '0 : Data Baru , 1 : Edit Data
    
    Frm115.TB1.SetFocus
    
End If
End Sub

Private Sub CMD33_Click()
'on error resume next
Call MDI_frm1_unload_all_menu
Call Frm116_background_color

Call frm116_one_time_reset
Call frm116_reset_1
Call Frm116_reset_3

frm116.L69_Text = -1 'Titik Pencarian Data
frm116.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
frm116.L67_Text = 0 'Paparan Page ke-xxx
frm116.L68_Text = 0

Call Frm116_Senarai_Belian_Header

frm116.Show
frm116.L32_Text = 0 '0 : Data Baru , 1 : Edit Data

GLOBAL_DISABLE = 0

MDI_frm1.L5_Text = 17
End Sub

Private Sub CMD34_Click()
'on error resume next
Call MDI_frm1_unload_all_menu
Call Frm117_background_color
Call frm117_pic_ena_disable
Call frm117_initial_setting

MDI_frm1.L5_Text = 32

frm117.L69_Text = -1 'Titik Pencarian Data
frm117.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
frm117.L67_Text = 0 'Paparan Page ke-xxx
frm117.L68_Text = 0
End Sub

Private Sub CMD35_Click()
'on error resume next
Call MDI_frm1_unload_all_menu
Call Frm118_background_color

Call frm118_initial_setting
MDI_frm1.L5_Text = 33
End Sub

Private Sub CMD36_Click()
'on error resume next
Call MDI_frm1_unload_all_menu
Call Frm83_background_color

Frm83.CBB1.Enabled = False
Frm83.CBB1.BackColor = &H8000000A

Call Frm83_form_load_trade_in
Call Frm83_form_load
Call Frm83_form_load_trade_in
Call frm83_flag_barang_trade_in
MDI_frm1.L5_Text = 3
End Sub

Private Sub CMD39_Click()
'on error resume next
Note = "Menu ini adalah dikhususkan bagi hantaran barang/tukaran barang dengan pihak supplier atau agen." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Barangan ini akan dihantar secara BULK." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If

If Answer = vbYes Then

    Call MDI_frm1_unload_all_menu
    Call frm123_background_color
    
    Call Frm123_one_time_reset
    Call Frm123_reset_1
    Call Frm123_reset_3
    
    frm123.L69_Text = -1 'Titik Pencarian Data
    frm123.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    frm123.L67_Text = 0 'Paparan Page ke-xxx
    frm123.L68_Text = 0
    
    Call Frm123_Senarai_Belian_Header
    
    frm123.Show
    frm123.L32_Text = 0 '0 : Data Baru , 1 : Edit Data
    
    GLOBAL_DISABLE = 0
    
    MDI_frm1.L5_Text = 34
    
End If
End Sub

Private Sub CMD40_Click()
'on error resume next
Call MDI_frm1_unload_all_menu
Call Frm124_background_color

Call frm124_initial_setting

frm124.Show
End Sub

Private Sub CMD41_Click()
'on error resume next
Call MDI_frm1_unload_all_menu
Call Frm126_background_color

frm126.L69_Text = -1 'Titik Pencarian Data
frm126.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
frm126.L67_Text = 0 'Paparan Page ke-xxx
frm126.L68_Text = 0

frm126.L10_Text = 0
frm126.L11_Text = "0.00 g"
frm126.L12_Text = "RM 0.00"

frm126.Show
End Sub

Private Sub CMD42_Click()
'on error resume next
Call MDI_frm1_unload_all_menu
Call Frm127_background_color

frm127.L69_Text = -1 'Titik Pencarian Data
frm127.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
frm127.L67_Text = 0 'Paparan Page ke-xxx
frm127.L68_Text = 0

GM_NEXT_PREV = 0

Call frm127_log_header
Call frm127_log

frm127.Show
End Sub

Private Sub CMD43_Click()
'on error resume next
Call MDI_frm1_unload_all_menu
Call Frm129_background_color

Call frm129_initial_setting

frm129.Show
End Sub

Private Sub CMD44_Click()
'on error resume next
'If MDI_frm1.L20_Text = "Semua cawangan" Then

    Frm96.CMD2.Visible = True
    Frm96.CMD1.Visible = False

    Call Frm96_initial
        
    Frm96.Show vbModal
    
'End If
End Sub

Private Sub CMD5_Click()
'on error resume next
If MDI_frm1.L20_Text = "Semua cawangan" Then

    Frm96.CMD2.Visible = True
    Frm96.CMD1.Visible = False

    Call Frm96_initial
        
    Frm96.Show vbModal
    
End If

Call MDI_frm1_unload_all_menu
Call Frm84_background_color

Frm84.CB4 = 1
Call Frm84_form_load
'Frm84.CB4 = 1
Frm84.L62_Text = "Jualan oleh agen dropship : TIDAK"
MDI_frm1.L5_Text = 4
End Sub
Private Sub CMD7_Click()
'On Error Resume Next
Call MDI_frm1_unload_all_menu
Call Frm101_background_color

Call Frm101_initial_setting
Frm101.Show

Frm101.Pic1.Visible = True
Frm101.CB2 = 1
MDI_frm1.L5_Text = 12
End Sub
Private Sub CMD6_Click()
'on error resume next
Call MDI_frm1_unload_all_menu
Call Frm102_background_color

GLOBAL_DISABLE = 0
Frm102.TB1 = vbNullString

Call frm102_reset_1
Call frm102_reset_2
Call frm102_reset_3
Call frm102_reset_main
Call Frm28_initial

Frm102.L26_Text.BackStyle = 0
Frm102.L27_Text.BackStyle = 0

Frm102.DTPicker1 = DateTime.Date$
MDI_frm1.L5_Text = 6

Call Frm102_Senarai_Jualan_Header
Call Frm102_senarai_belian_header

Frm102.CMD8.Visible = True
Frm102.CMD9.Visible = True
Frm102.CMD10.Visible = False
Frm102.CMD11.Visible = False

Frm102.Picture = MDI_frm1.Picture
Frm102.Show

Frm102.L32_Text = 0 '0 : Data Baru , 1 : Edit Data

Frm102.TB1.SetFocus
End Sub

Private Sub jcbutton7_Click()
'on error resume next
Call MDI_frm1_unload_all_menu
Frm103.Picture = MDI_frm1.Picture

Frm103.Show
End Sub
Private Sub CMD8_Click()
'On Error Resume Next
Call MDI_frm1_unload_all_menu
Call Frm95_background_color

MDI_frm1.L5_Text = 18
End Sub
Private Sub CMD9_Click()
'On Error Resume Next
Call MDI_frm1_unload_all_menu
Call Frm54_background_color

Frm54.CBB1.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database order by Kod_Metal_Purity DESC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If rs!Metal_Purity <> vbNullString Then
        Frm54.CBB1.AddItem rs!Kod_Metal_Purity
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm54.Show

MDI_frm1.L5_Text = 19
End Sub

Private Sub Command2_Click()
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim LM_JUALAN As Double
Dim LM_JUALAN_23 As Double

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from delete_item", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    LM_INVOICE = vbNullString
    LM_INVOICE = rs!no_invoice
    LM_JUALAN = 0
    LM_JUALAN_23 = 0
    
    Set rs1 = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs1.Open "select * from 22_jualan where no_resit='" & rs!no_invoice & "' and status = 0", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs1.EOF Then
        If Not IsNull(rs1!harga_jualan) Then LM_JUALAN = rs1!harga_jualan
    End If
    
    rs1.Close
    Set rs1 = Nothing
    
    Set rs2 = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs2.Open "select SUM(harga_jualan) from 23_senarai_jualan where no_resit='" & rs!no_invoice & "' and status_rekod = 0", cn, adOpenKeyset, adLockOptimistic
    
    If Not IsNull(rs2(0)) Then LM_JUALAN_23 = rs2(0)
    
    rs2.Close
    Set rs2 = Nothing
    
    If LM_JUALAN <> LM_JUALAN_23 Then MsgBox rs!no_invoice
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

MsgBox "done"

End Sub

Private Sub Command3_Click()
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim LM_JUALAN As Double
Dim LM_JUALAN_23 As Double

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from delete_item", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    LM_INVOICE = vbNullString
    LM_INVOICE = rs!no_invoice
    LM_JUALAN = 0
    LM_JUALAN_23 = 0
    
    Set rs1 = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

    strsql = "UPDATE 22_jualan set status = 1 where no_resit='" & rs!no_invoice & "' and status = 0"
    
    Set rs1 = cn.Execute(strsql)
    Set rs1 = Nothing
    
Dim Frm84_LM_BERAT_ASAL As Double
Dim Frm84_LM_BERAT_JUALAN As Double

    Set rs1 = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs1.Open "select * from 23_senarai_jualan where status_rekod = 0 and no_resit='" & rs!no_invoice & "'", cn, adOpenKeyset, adLockOptimistic
    
    While rs1.EOF = False
    
        Set rs2 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs2.Open "select * from Data_Database where no_siri_produk='" & rs1!no_siri_Produk & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs2.EOF Then
            Frm84_LM_BERAT_ASAL = 0
            Frm84_LM_BERAT_JUALAN = 0
            
            If rs1!Type = 0 Then
                Frm84_LM_BERAT_ASAL = rs2!beza_berat 'Berat Asal (g)
                Frm84_LM_BERAT_JUALAN = rs1!berat_jualan 'Berat Jualan (g)
                
                If Frm84_LM_BERAT_JUALAN = Frm84_LM_BERAT_ASAL Then
                    rs2!beza_berat = "0.00" 'Baki Berat
                    rs2!susut_berat = "0.00" 'Susut berat
                    rs2!StatusItem = 11
                Else
                    rs2!beza_berat = Format(Frm84_LM_BERAT_ASAL - Frm84_LM_BERAT_JUALAN, "0.00") 'Baki Berat
                    rs2!susut_berat = "0.00" 'Susut berat
                    rs2!StatusItem = 12
                End If
            Else
                rs2!StatusItem = 11
            End If
            rs2!remarks = "restore data"
            rs2.Update
        End If
        
        rs2.Close
        Set rs2 = Nothing
        
        rs1!status_rekod = 1
        rs1.Update
        rs1.MoveNext
    Wend
    
    rs1.Close
    Set rs1 = Nothing

    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

MsgBox "done"



End Sub

Private Sub L10_Text_Click()
'on error resume next
Call MDI_frm1_unload_all_menu
Frm7.Picture = MDI_frm1.Picture

Frm7.Show
MDI_frm1.L5_Text = 26
Frm7.TB1 = MDI_frm1.L3_Text
Frm7.TB2.SetFocus
End Sub
Private Sub L11_Text_Click()
'on error resume next
Frm4.Picture = MDI_frm1.Picture

Frm4.L1_Text = vbNullString
Frm4.L2_Text = vbNullString
Frm4.L3_Text = vbNullString
Frm4.L4_Text = vbNullString
Frm4.L5_Text = vbNullString
Frm4.L6_Text = vbNullString

If MDI_frm1.L4_Text = "HQ" Then
    
    G_KEDAI = "HQ"
    
Else

    G_KEDAI = MDI_frm1.L20_Text
    
End If

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai_2) Then Frm4.L1_Text = rs!nama_kedai_2 'Nama kedai
    If Not IsNull(rs!version_sistem) Then Frm4.L2_Text = rs!version_sistem 'Version Sistem
    If Not IsNull(rs!version_database) Then Frm4.L3_Text = rs!version_database 'Version database
    If Not IsNull(rs!version_database_image) Then Frm4.L4_Text = rs!version_database_image 'Versin database image
    If Not IsNull(rs!version_database_ae) Then Frm4.L5_Text = rs!version_database_ae 'Version automation email
    If Not IsNull(rs!version_database_ab) Then Frm4.L6_Text = rs!version_database_ab 'Version auto backup
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End


Frm4.Show 1
End Sub
Private Sub L12_Text_Click()
'on error resume next
Frm3.cmdlogin.Visible = False
Frm3.CMD1.Visible = False

Frm3.CMD2.Visible = True
Frm3.CMD3.Visible = True

Frm3.Show 1
End Sub
Private Sub L13_Text_Click()
'on error resume next
Shell "calc.exe", vbNormalFocus
End Sub
Private Sub L14_Text_Click()
'on error resume next
Shell "notepad.exe", vbNormalFocus
End Sub

Private Sub L15_Text_Click()
'on error resume next
Frm114.Show
End Sub

Private Sub L17_Text_Change()
'on error resume next
If G_SYSTEM_TYPE = "ONLINE" And MDI_frm1.L17_Text = "OFFLINE" Then
    Call MDI_frm1_unload_all_menu
    
    'MsgBox "Tiada sambungan sistem dan internet. OFFLINE." & vbCrLf & _
            vbNullString & vbrlf & _
            "Semua menu telah ditutup.", vbCritical, "Connection Error"
    
End If
End Sub

Private Sub L17_Text_Click()
'on error resume next
If MDI_frm1.L17_Text = "ONLINE" Then
    G_MAIN_CONN = 1
ElseIf MDI_frm1.L17_Text = "OFFLINE" Then
    G_MAIN_CONN = 0
End If
End Sub

Private Sub L2_Text_Change()
'on error resume next
Call check_internet_connection_main
End Sub

Private Sub L6_Text_Click()
'on error resume next
Call Frm68_background_color

If MDI_frm1.L5_Text = "1" Then

ElseIf MDI_frm1.L5_Text = "2" Then

ElseIf MDI_frm1.L5_Text = "3" Then

    Frm83.Hide
    Frm68.Show
    
ElseIf MDI_frm1.L5_Text = "4" Then

    Frm84.Hide
    Frm68.Show
    
ElseIf MDI_frm1.L5_Text = "5" Then

ElseIf MDI_frm1.L5_Text = "6" Then

ElseIf MDI_frm1.L5_Text = "7" Then

    Frm87.Hide
    Frm68.Show
    
ElseIf MDI_frm1.L5_Text = "8" Then

    Frm93.Hide
    Frm68.Show
    
ElseIf MDI_frm1.L5_Text = "10" Then

    Frm92.Hide
    Frm68.Show

Else

    Call MDI_frm1_unload_all_menu
    
    Frm68.Show
    Frm68.L36_Text = 0 '0 : Terus dari menu data pelanggan , 1 : Data pembeli , 2 : Data agen dropship
    
End If

Call Frm68_background_color
End Sub
Private Sub L7_Text_Click()
'on error resume next
If MDI_frm1.L20_Text = "Semua cawangan" Then

    Frm96.CMD2.Visible = True
    Frm96.CMD1.Visible = False

    Call Frm96_initial
        
    Frm96.Show vbModal
    
End If

Call MDI_frm1_unload_all_menu
Call Frm84_background_color

G_CALC_AUTO = 0

Frm84.CB4 = 1
Call Frm84_form_load

'Note = "Sila buat pilihan jenis pengiraan upah." & vbCrLf & _
        vbNullString & vbCrLf & _
        "YES : Upah mengikut tetapan per item" & vbCrLf & _
        "NO  : Upah mengikut berat"

'Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

'If Answer = vbNo Then
'    G_KIRAAN_UPAH = 0
'    Frm84.L86_Text = "PENGIRAAN UPAH MENGIKUT BERAT"
'Else
'    G_KIRAAN_UPAH = 1
'    Frm84.L86_Text = "PENGIRAAN UPAH MENGIKUT UPAH PER ITEM"
'End If
    
'Frm84.CB4 = 1
Frm84.L62_Text = "Jualan oleh agen dropship : TIDAK"
MDI_frm1.L5_Text = 4
End Sub
Private Sub L8_Text_Click()
'On Error Resume Next
If MDI_frm1.L4_Text = "Admin" Or MDI_frm1.L4_Text = "Manager" Or MDI_frm1.L4_Text = "HQ" Or user_level = "Developer" Then

    Call MDI_frm1_unload_all_menu
    Call Frm54_background_color
    
    Frm54.CBB1.Clear
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from setting_database order by Kod_Metal_Purity DESC", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        If rs!Metal_Purity <> vbNullString Then
            Frm54.CBB1.AddItem rs!Kod_Metal_Purity
        End If
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
    Frm54.Show
    
    MDI_frm1.L5_Text = 19
    
Else

    MsgBox "Anda tidak dibenarkan untuk memasuki menu ini.", vbExclamation, "Info"
    
End If
End Sub

Private Sub Label8_Click()

End Sub

Private Sub MD1_SM_SM_1_Click()
'on error resume next
MDI_frm1.Picture = LoadPicture(App.Path & "\Backgrounds\3.jpg")
End Sub
Private Sub MD1_SM_SM_2_Click()
'on error resume next
MDI_frm1.Picture = LoadPicture(App.Path & "\Backgrounds\2.jpg")
End Sub
Private Sub MD1_SM_SM_3_Click()
'on error resume next
MDI_frm1.Picture = LoadPicture(App.Path & "\Backgrounds\3.jpg")
End Sub
Private Sub MD1_SM_SM_4_Click()
'on error resume next
MDI_frm1.Picture = LoadPicture(App.Path & "\Backgrounds\4.jpg")
End Sub
Private Sub MD1_SM_SM_5_Click()
'on error resume next
MDI_frm1.Picture = LoadPicture(App.Path & "\Backgrounds\5.jpg")
End Sub
Private Sub MD1_SM_SM_6_Click()
'on error resume next
MDI_frm1.Picture = LoadPicture(App.Path & "\Backgrounds\6.jpg")
End Sub
Private Sub MD1_SM_SM_7_Click()
'on error resume next
MDI_frm1.Picture = LoadPicture(App.Path & "\Backgrounds\7.jpg")
End Sub
Private Sub MDIForm_Load()
'on error resume next
xLogOff = False

MDI_frm1.L3_Text = "-----------------" 'User
MDI_frm1.L4_Text = "-----------------" 'Level

'MDI_frm1.Picture = LoadPicture(App.Path & "\Backgrounds\1.jpg")

MDI_frm1.L5_Text = 0

MDI_frm1.L18_Text = "0"
MDI_frm1.L19_Text = "0"
MDI_frm1.L22_Text = "0"
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
'on error resume next

'### Tutup form dengan klik [X] ### - Start
If xLogOff = False Then

    MDI_frm1.L5_Text = 0

    Note = "Adakah anda ingin menutup sistem ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Cancel = 1
        Exit Sub
    End If
    If Answer = vbYes Then
        Call amendment_email_check
        
        End
    End If
    
End If

'### Tutup form dengan klik [X] ### - End

End Sub

Private Sub Timer2_Timer()
    If amination(0).Top = -480 Then
        amination(1).Top = 720
    End If
    
    If amination(1).Top = -480 Then
        amination(0).Top = 720
    End If
    
    amination(0).Top = amination(0).Top - 5
    amination(1).Top = amination(1).Top - 5
End Sub
Private Sub Tmr1_Timer()
'on error resume next
MDI_frm1.L1_Text = DateTime.Date
MDI_frm1.L2_Text = DateTime.Time
End Sub

Private Sub Tmr2_Timer()
'on error resume next
If MDI_frm1.L17_Text = "OFFLINE" Then
    
    If MDI_frm1.Image3.Visible = True Then
        MDI_frm1.Image3.Visible = False
    Else
        MDI_frm1.Image3.Visible = True
    End If
    
End If
End Sub

Private Sub Tmr3_Timer()
'on error resume next
Call check_internet_interval
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
'on error resume next
MDI_LM_INDEX = Button.Index

If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
    
    MDI_frm1.CMD44.Enabled = True
    
Else
    
    MDI_frm1.CMD44.Enabled = False

End If

Call MDI_frm1_unload_sub_menu
Call MDI_frm1_unload_all_menu

Select Case Button.Index

Case 1 'Penerimaan Stok

    Call MDI_frm1_unload_sub_menu
    
    frm131.LV1.ListItems.Clear
    
    With frm131.LV1
        Set .SmallIcons = frm131.ImageList1
        Set .Icons = frm131.ImageList1
    
        .ListItems.Add , "Penerimaan Stok Baru", "Penerimaan Stok Baru", 6
        .ListItems.Add , "Belian Emas Terpakai", "Belian Emas Terpakai", 6
    
    End With
    
Case 2 'Urusan Kedai
    
    Call MDI_frm1_unload_sub_menu
    
    frm131.LV1.ListItems.Clear
    
    With frm131.LV1
        Set .SmallIcons = frm131.ImageList1
        Set .Icons = frm131.ImageList1
    
        .ListItems.Add , "Jualan", "Jualan", 1
        .ListItems.Add , "Tempahan", "Tempahan", 8
        .ListItems.Add , "Trade In", "Trade In", 6
        .ListItems.Add , "Servis & Belanja", "Servis & Belanja", 7
        .ListItems.Add , "Maklumat Pelanggan", "Maklumat Pelanggan", 3
        .ListItems.Add , "Pengeluaran & Kemasukkan Tunai", "Pengeluaran & Kemasukkan Tunai", 10
        .ListItems.Add , "E-mail Promosi", "E-mail Promosi", 2
        .ListItems.Add , "Update Dulang", "Update Dulang", 21
        If G_GDN_SUBSCRIBE = "YES" Then .ListItems.Add , "Goods Despatch Note (Per Item)", "Goods Despatch Note (Per Item)", 14
        If G_GDN_SUBSCRIBE = "YES" Then .ListItems.Add , "Goods Despatch Note (Bulk)", "Goods Despatch Note (Bulk)", 14
        If G_GDN_SUBSCRIBE = "YES" Then .ListItems.Add , "Goods Received Note", "Goods Received Note", 14
        If G_GDN_SUBSCRIBE = "YES" Then .ListItems.Add , "Invoice / Voucher", "Invoice / Voucher", 14
        If G_SYSTEM_TYPE = "OFFLINE" Then .ListItems.Add , "Backup Database", "Backup Database", 22
        '.ListItems.Add , "Pengurusan Buku Cek", "Pengurusan Buku Cek", 11
        '.ListItems.Add , "Agihan Stok", "Agihan Stok", 13
        
    End With

Case 3 'Report
    
    If MDI_frm1.L4_Text = "Admin" Or MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
    
        Call MDI_frm1_unload_sub_menu
        
        frm131.LV1.ListItems.Clear
        
        With frm131.LV1
            Set .SmallIcons = frm131.ImageList1
            Set .Icons = frm131.ImageList1
        
            .ListItems.Add , "Report Keseluruhan", "Report Keseluruhan", 16
            .ListItems.Add , "Report Trade In", "Report Trade In", 16
            If G_GDN_SUBSCRIBE = "YES" Then .ListItems.Add , "GDN/GRN", "GDN/GRN", 16
            .ListItems.Add , "Report Kewangan", "Report Kewangan", 16
            .ListItems.Add , "Senarai Invoice", "Senarai Invoice", 17
            .ListItems.Add , "Penyata Untung Rugi (Restock)", "Penyata Untung Rugi (Restock)", 18
            .ListItems.Add , "Penyata Untung Rugi (Runcit)", "Penyata Untung Rugi (Runcit)", 18
            .ListItems.Add , "Inventori Dulang", "Inventori Dulang", 4
            .ListItems.Add , "Report Stok Dulang", "Report Stok Dulang", 4
            .ListItems.Add , "Report GST", "Report GST", 5
            .ListItems.Add , "Barang Hilang", "Barang Hilang", 19
            .ListItems.Add , "Log", "Log", 9
            
        End With

    Else
    
        MsgBox "Anda tidak dibenarkan untuk memasuki menu ini.", vbExclamation, "Info"
        
    End If
Case 4

    If MDI_frm1.L4_Text = "Admin" Or MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
    
        Call MDI_frm1_unload_sub_menu
        Call MDI_frm1_unload_all_menu
    
        frm131.LV1.ListItems.Clear
        
        With frm131.LV1
            Set .SmallIcons = frm131.ImageList1
            Set .Icons = frm131.ImageList1
        
            .ListItems.Add , "Tetapan Asas Sistem", "Tetapan Asas Sistem", 20
            .ListItems.Add , "Tetapan Harga Jualan Emas", "Tetapan Harga Jualan Emas", 24
            .ListItems.Add , "Tetapan Sistem", "Tetapan Sistem", 20
            .ListItems.Add , "Data Pekerja", "Data Pekerja", 3
            .ListItems.Add , "Tetapan Barcode", "Tetapan Barcode", 21
            .ListItems.Add , "Payroll", "Payroll", 23
            .ListItems.Add , "Analisa Harga Emas", "Analisa Harga Emas", 18
            If G_SYSTEM_TYPE = "OFFLINE" Then .ListItems.Add , "Backup Database", "Backup Database", 22
            'If MDI_frm1.L4_Text = "Developer" Then .ListItems.Add , "Setting Invoice", "Setting Invoice", 20
            If MDI_frm1.L4_Text = "Developer" Then .ListItems.Add , "Developer", "Developer", 25
        
        End With
        
    Else
    
        MsgBox "Anda tidak dibenarkan untuk memasuki menu ini.", vbExclamation, "Info"
        
    End If
    
Case 5

    Note = "Adakah anda ingin keluar dari sistem ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then
        
        xLogOff = True
        Frm3.Show
        Unload MDI_frm1
        
    Else
    
        MDI_frm1.CMD44.Enabled = False
        
    End If
    
End Select

Exit Sub

Select Case Button.Index

Case 1
    Call MDI_frm1_unload_sub_menu
    Call MDI_frm1_unload_all_menu
    
    MDI_frm1.Pic2.Visible = True
Case 2
    Call MDI_frm1_unload_sub_menu
    Call MDI_frm1_unload_all_menu
    
    MDI_frm1.Pic3.Visible = True
Case 3
    Call MDI_frm1_unload_sub_menu
    Call MDI_frm1_unload_all_menu
    
    MDI_frm1.Pic4.Visible = True
Case 4

    If MDI_frm1.L4_Text = "Admin" Or MDI_frm1.L4_Text = "HQ" Or user_level = "Developer" Then
    
        Call MDI_frm1_unload_sub_menu
        Call MDI_frm1_unload_all_menu
    
        MDI_frm1.Pic5.Visible = True
        
    Else
    
        MsgBox "Anda tidak dibenarkan untuk memasuki menu ini.", vbExclamation, "Info"
        
    End If
    
Case 5

    Note = "Adakah anda ingin keluar dari sistem ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then
        
        xLogOff = True
        Frm3.Show
        Unload MDI_frm1
        
    End If
    
End Select
End Sub
