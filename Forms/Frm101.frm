VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm101 
   Caption         =   "Report Keseluruhan"
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
   Icon            =   "Frm101.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13035
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic1 
      BorderStyle     =   0  'None
      Height          =   10215
      Left            =   480
      ScaleHeight     =   10215
      ScaleWidth      =   18465
      TabIndex        =   2
      Top             =   720
      Width           =   18465
      Begin VB.ComboBox CBB8 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2145
         Style           =   2  'Dropdown List
         TabIndex        =   85
         Top             =   4360
         Width           =   7000
      End
      Begin VB.ComboBox CBB7 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   360
         ItemData        =   "Frm101.frx":0ECA
         Left            =   2145
         List            =   "Frm101.frx":0ECC
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   2560
         Width           =   7000
      End
      Begin VB.CommandButton CMD6 
         BackColor       =   &H000080FF&
         Caption         =   "Carian Data"
         Height          =   405
         Left            =   4320
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm101.frx":0ECE
         MousePointer    =   99  'Custom
         TabIndex        =   78
         ToolTipText     =   "Carian data bagi setiap report dengan menggunakan No. Siri Produk"
         Top             =   9480
         Width           =   2145
      End
      Begin VB.CommandButton CMD4 
         BackColor       =   &H000080FF&
         Caption         =   "Ringkasan Report"
         Height          =   405
         Left            =   4680
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm101.frx":11D8
         MousePointer    =   99  'Custom
         TabIndex        =   77
         ToolTipText     =   $"Frm101.frx":14E2
         Top             =   7920
         Width           =   2385
      End
      Begin VB.CommandButton CMD1 
         BackColor       =   &H000080FF&
         Caption         =   "Report Terperinci"
         Height          =   405
         Left            =   1920
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm101.frx":1580
         MousePointer    =   99  'Custom
         TabIndex        =   76
         ToolTipText     =   "Report ini akan memaparkan maklumat terperinci bagi setiap jenis report. Data dari report ini  boleh dieksport ke EXCEL."
         Top             =   7920
         Width           =   2385
      End
      Begin VB.ComboBox CBB6 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   4020
         Width           =   3030
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
         Left            =   360
         TabIndex        =   71
         Top             =   6960
         Width           =   200
      End
      Begin VB.ComboBox CBB5 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   13680
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   5280
         Visible         =   0   'False
         Width           =   3030
      End
      Begin VB.TextBox TB4 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2160
         TabIndex        =   64
         Text            =   "TB4"
         Top             =   9540
         Width           =   1980
      End
      Begin VB.ComboBox CBB4 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   360
         ItemData        =   "Frm101.frx":188A
         Left            =   2145
         List            =   "Frm101.frx":188C
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   2925
         Width           =   7000
      End
      Begin VB.PictureBox Pic9 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   13320
         ScaleHeight     =   1815
         ScaleWidth      =   5505
         TabIndex        =   55
         Top             =   4440
         Visible         =   0   'False
         Width           =   5505
         Begin VB.CommandButton CMD5 
            BackColor       =   &H000080FF&
            Caption         =   "Carian Data"
            Height          =   405
            Left            =   1320
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm101.frx":188E
            MousePointer    =   99  'Custom
            TabIndex        =   79
            ToolTipText     =   "Carian maklumat / report belian melalui No. Invoice dari supplier."
            Top             =   1320
            Width           =   2145
         End
         Begin VB.TextBox TB3 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   2160
            TabIndex        =   56
            Text            =   "TB3"
            Top             =   840
            Width           =   1980
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Invoice        :"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   600
            TabIndex        =   58
            Top             =   870
            Width           =   2505
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Sila masukkan [No. Invoice Belian Dari Supplier] bagi mencari data terperinci bagi belian dari supplier."
            ForeColor       =   &H00000000&
            Height          =   525
            Left            =   240
            TabIndex        =   57
            Top             =   240
            Width           =   5000
         End
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
         Left            =   360
         TabIndex        =   27
         Top             =   6720
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
         Left            =   360
         TabIndex        =   26
         Top             =   6480
         Width           =   200
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
         Left            =   360
         TabIndex        =   25
         Top             =   6000
         Width           =   200
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
         Left            =   360
         TabIndex        =   24
         Top             =   5775
         Width           =   200
      End
      Begin VB.CheckBox CB1 
         Caption         =   "Print Barcode      :"
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
         TabIndex        =   23
         Top             =   700
         Width           =   200
      End
      Begin VB.ComboBox CBB2 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2145
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3645
         Width           =   7000
      End
      Begin VB.ComboBox CBB1 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   360
         ItemData        =   "Frm101.frx":1B98
         Left            =   2160
         List            =   "Frm101.frx":1B9A
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   3285
         Width           =   7000
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
         Left            =   4200
         TabIndex        =   20
         Top             =   5760
         Width           =   200
      End
      Begin VB.PictureBox Pic7 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   14640
         ScaleHeight     =   1815
         ScaleWidth      =   5505
         TabIndex        =   16
         Top             =   720
         Visible         =   0   'False
         Width           =   5505
         Begin VB.CommandButton CMD2 
            BackColor       =   &H000080FF&
            Caption         =   "Carian Data"
            Height          =   405
            Left            =   1560
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm101.frx":1B9C
            MousePointer    =   99  'Custom
            TabIndex        =   80
            ToolTipText     =   "Carian maklumat barang melalui berat barang tersebut."
            Top             =   1320
            Width           =   2145
         End
         Begin VB.TextBox TB1 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   2160
            TabIndex        =   17
            Text            =   "TB1"
            Top             =   840
            Width           =   1980
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "Sila masukkan [Berat Stok] bagi mencari semua stok yang mempunyai berat yang hendak dicari."
            ForeColor       =   &H00000000&
            Height          =   525
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   5370
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "Berat    (g):"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   1080
            TabIndex        =   18
            Top             =   870
            Width           =   1905
         End
      End
      Begin VB.PictureBox Pic8 
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   9960
         ScaleHeight     =   3135
         ScaleWidth      =   5505
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   5505
         Begin VB.CommandButton CMD3 
            BackColor       =   &H000080FF&
            Caption         =   "Carian Data"
            Height          =   405
            Left            =   1560
            MaskColor       =   &H00400000&
            MouseIcon       =   "Frm101.frx":1EA6
            MousePointer    =   99  'Custom
            TabIndex        =   81
            ToolTipText     =   "Carian maklumat / report jualan melalui No. Invoice"
            Top             =   2640
            Width           =   2145
         End
         Begin VB.TextBox TB2 
            BackColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   2640
            TabIndex        =   11
            Text            =   "TB2"
            Top             =   840
            Width           =   1740
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
            Left            =   1200
            TabIndex        =   10
            Top             =   1935
            Width           =   200
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
            Left            =   1200
            TabIndex        =   9
            Top             =   2190
            Width           =   200
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Invoice / Voucher :"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   600
            TabIndex        =   15
            Top             =   870
            Width           =   2385
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Sila masukkan [No. Invoice / Voucher] bagi mencari rekod dan maklumat bagi No. Invoice / Voucher ini."
            ForeColor       =   &H00000000&
            Height          =   525
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   5085
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Jualan                       Voucher Buyback / Trade In"
            ForeColor       =   &H00000000&
            Height          =   630
            Left            =   1440
            TabIndex        =   13
            Top             =   1905
            Width           =   3165
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Sila pilih samada [Invoice Jualan] atau [Voucher Buyback / Trade In]"
            ForeColor       =   &H00000000&
            Height          =   525
            Left            =   840
            TabIndex        =   12
            Top             =   1440
            Width           =   3765
         End
         Begin VB.Shape Shape3 
            Height          =   1215
            Left            =   600
            Top             =   1320
            Width           =   4095
         End
      End
      Begin VB.ComboBox CBB3 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2145
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   4005
         Width           =   2550
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
         Left            =   360
         TabIndex        =   6
         Top             =   6240
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
         Left            =   8160
         TabIndex        =   5
         Top             =   5760
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.CheckBox CB11 
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
         TabIndex        =   4
         Top             =   6000
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.CheckBox CB12 
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
         TabIndex        =   3
         Top             =   6240
         Visible         =   0   'False
         Width           =   200
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   2160
         TabIndex        =   28
         Top             =   1245
         Width           =   7005
         _ExtentX        =   12356
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
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   360
         Left            =   2160
         TabIndex        =   29
         Top             =   1605
         Width           =   7005
         _ExtentX        =   12356
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
      Begin VB.Label L47_Text 
         Caption         =   "L47_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   11280
         TabIndex        =   87
         Top             =   8520
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Staff *"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   240
         TabIndex        =   86
         Top             =   4430
         Width           =   2295
      End
      Begin VB.Label L46_Text 
         Caption         =   "L46_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   11280
         TabIndex        =   84
         Top             =   8160
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Cawangan *"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   240
         TabIndex        =   83
         Top             =   2575
         Width           =   2295
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "** Report stok ini tidak termasuk barang trade in."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1800
         TabIndex        =   75
         Top             =   6720
         Width           =   7515
      End
      Begin VB.Label L32_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Carian data mengikut No. Invoice / Voucher"
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
         Height          =   285
         Left            =   9720
         MouseIcon       =   "Frm101.frx":21B0
         MousePointer    =   99  'Custom
         TabIndex        =   36
         Top             =   1200
         Width           =   4530
      End
      Begin VB.Label L35_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Carian data mengikut No. Invoice dari supplier"
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
         Height          =   285
         Left            =   9720
         MouseIcon       =   "Frm101.frx":24BA
         MousePointer    =   99  'Custom
         TabIndex        =   54
         Top             =   1440
         Width           =   4530
      End
      Begin VB.Label L45_Text 
         Caption         =   "L45_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   11280
         TabIndex        =   74
         Top             =   7800
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Cara Jualan *"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   4800
         TabIndex        =   73
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label L44_Text 
         Caption         =   "L44_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   11280
         TabIndex        =   70
         Top             =   7440
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm101.frx":27C4
         ForeColor       =   &H00000000&
         Height          =   885
         Left            =   13080
         TabIndex        =   69
         Top             =   7560
         Visible         =   0   'False
         Width           =   4995
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "** Jualan Biasa : Merujuk kepada jualan melalui MENU JUALAN kepada pelanggan."
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   13080
         TabIndex        =   68
         Top             =   6840
         Visible         =   0   'False
         Width           =   4995
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Tetapan bagi ""Jenis Jualan"" ini hanya digunakan untuk memaparkan REPORT JUALAN sahaja."
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   13200
         TabIndex        =   67
         Top             =   6240
         Visible         =   0   'False
         Width           =   4995
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   12480
         X2              =   13080
         Y1              =   6360
         Y2              =   6360
      End
      Begin VB.Shape Shape6 
         Height          =   1395
         Left            =   120
         Top             =   8640
         Width           =   9255
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Siri Produk   :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   600
         TabIndex        =   65
         Top             =   9540
         Width           =   2505
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm101.frx":284F
         ForeColor       =   &H00000000&
         Height          =   540
         Left            =   360
         TabIndex        =   63
         Top             =   9000
         Width           =   8655
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Carian data No. Siri Produk"
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
         Height          =   405
         Left            =   240
         TabIndex        =   62
         Top             =   8760
         Width           =   5535
      End
      Begin VB.Label L37_Text 
         Caption         =   "L37_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   11280
         TabIndex        =   61
         Top             =   6720
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier *"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   240
         TabIndex        =   60
         Top             =   2940
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Purity *"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   240
         TabIndex        =   52
         Top             =   3285
         Width           =   2295
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm101.frx":28DB
         ForeColor       =   &H00000000&
         Height          =   1725
         Left            =   645
         TabIndex        =   51
         Top             =   5760
         Width           =   4215
      End
      Begin VB.Shape Shape4 
         Height          =   3255
         Left            =   120
         Top             =   5280
         Width           =   9255
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Pilihan Report"
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
         Height          =   405
         Left            =   360
         TabIndex        =   50
         Top             =   5400
         Width           =   5535
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Tetapan Report"
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
         Height          =   405
         Left            =   360
         TabIndex        =   49
         Top             =   2250
         Width           =   5535
      End
      Begin VB.Shape Shape2 
         Height          =   2895
         Left            =   120
         Top             =   2205
         Width           =   9255
      End
      Begin VB.Shape Shape1 
         Height          =   1455
         Left            =   120
         Top             =   645
         Width           =   9255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm101.frx":2A1A
         ForeColor       =   &H00000000&
         Height          =   600
         Left            =   525
         TabIndex        =   48
         Top             =   690
         Width           =   8610
      End
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Akhir "
         Height          =   255
         Left            =   195
         TabIndex        =   47
         Top             =   1650
         Width           =   2895
      End
      Begin VB.Label Label62 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Mula "
         Height          =   255
         Left            =   195
         TabIndex        =   46
         Top             =   1290
         Width           =   2535
      End
      Begin VB.Label L9_Text 
         Caption         =   "L9_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9480
         TabIndex        =   45
         Top             =   7800
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label L8_Text 
         Caption         =   "L8_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9480
         TabIndex        =   44
         Top             =   7440
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label L7_Text 
         Caption         =   "L7_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9480
         TabIndex        =   43
         Top             =   7080
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label L6_Text 
         Caption         =   "L6_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9480
         TabIndex        =   42
         Top             =   6720
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label L5_Text 
         Caption         =   "L5_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9480
         TabIndex        =   41
         Top             =   6360
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Kategori Produk *"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   240
         TabIndex        =   40
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila buat pilihan krateria report seperti di bawah."
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
         Height          =   360
         Left            =   360
         TabIndex        =   39
         Top             =   240
         Width           =   8610
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Report Stok Potong "
         ForeColor       =   &H00000000&
         Height          =   1125
         Left            =   4440
         TabIndex        =   38
         Top             =   5760
         Width           =   4695
      End
      Begin VB.Label L31_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Carian data mengikut berat stok"
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
         Height          =   285
         Left            =   9720
         MouseIcon       =   "Frm101.frx":2AA9
         MousePointer    =   99  'Custom
         TabIndex        =   37
         Top             =   960
         Width           =   4530
      End
      Begin VB.Label L33_Text 
         Caption         =   "L33_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9480
         TabIndex        =   35
         Top             =   8160
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm101.frx":2DB3
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
         Height          =   480
         Left            =   360
         TabIndex        =   34
         Top             =   7200
         Width           =   8610
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Dulang *                                        Jenis Jualan *"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   240
         TabIndex        =   33
         Top             =   4080
         Width           =   6375
      End
      Begin VB.Label L34_Text 
         Caption         =   "L34_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   11280
         TabIndex        =   32
         Top             =   6360
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label L43_Text 
         Caption         =   "L43_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   11280
         TabIndex        =   31
         Top             =   7080
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Menu (Carian Data)"
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
         Left            =   9600
         TabIndex        =   30
         Top             =   600
         Width           =   5970
      End
   End
   Begin VB.Timer Tmr1 
      Interval        =   10
      Left            =   0
      Top             =   0
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
      Left            =   240
      MouseIcon       =   "Frm101.frx":2E3A
      MousePointer    =   99  'Custom
      TabIndex        =   53
      Top             =   120
      Width           =   2175
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
      Left            =   21600
      TabIndex        =   1
      Top             =   435
      Visible         =   0   'False
      Width           =   2100
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
      Left            =   21600
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2100
   End
End
Attribute VB_Name = "Frm101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CB1_Click()
'on error resume next
If Frm101.CB1 = 1 Then
    Frm101.L9_Text = 1
Else
    Frm101.L9_Text = 0
End If
End Sub

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
    Frm101.CB13 = 0
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
    Frm101.CB13 = 0
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
    Frm101.CB13 = 0
End If
End Sub
Private Sub CB13_Click()
'on error resume next
If Frm101.CB13 = 1 Then
    Frm101.CB3 = 0
    Frm101.CB4 = 0
    Frm101.CB2 = 0
    Frm101.CB6 = 0
    Frm101.CB9 = 0
    Frm101.CB10 = 0
    Frm101.CB11 = 0
    Frm101.CB12 = 0
    Frm101.CB5 = 0
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
    Frm101.CB13 = 0
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
    Frm101.CB13 = 0
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
    Frm101.CB13 = 0
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
    Frm101.CB13 = 0
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
    Frm101.CB13 = 0
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
    Frm101.CB13 = 0
End If
End Sub
Private Sub CMD1_Click()
'on error resume next
If Frm101.CB2 = 0 And Frm101.CB3 = 0 And Frm101.CB4 = 0 And Frm101.CB5 = 0 And Frm101.CB6 = 0 And Frm101.CB9 = 0 And Frm101.CB10 = 0 And Frm101.CB11 = 0 And Frm101.CB12 = 0 And Frm101.CB13 = 0 Then
    MsgBox "Sila Buat Pilihan Tetapan Report.", vbInformation, "Info"
    Exit Sub
End If

Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report. Teruskan ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then

    If Frm101.CB3 = 1 Then 'Report Jualan
    
        LM_REPORT_JUALAN = vbNullString ' 1 : Barang Baru Sahaja , 2 : Barang Trade In Sahaja , 3 : Semua
        
        LM_REPORT_JUALAN = InputBox("Sila pilih jenis barang jualan." & _
                vbCrLf & _
                vbCrLf & vbTab & "1 - Barang Baru Sahaja" & _
                vbCrLf & vbTab & "2 - Barang Trade In Sahaja" & _
                vbCrLf & vbTab & "3 - Semua", "Jenis Jualan")
                 
        Select Case LM_REPORT_JUALAN
            Case "1"
        
                G_JENIS_JUALAN = "Barang Baru Sahaja"
                
            Case "2"
                
                G_JENIS_JUALAN = "Barang Trade In Sahaja"
                
            Case "3"

                G_JENIS_JUALAN = "Barang Baru Dan Barang Trade In"
                
            Case Else
            
                MsgBox "Tiada pilihan dibuat atau pilihan yang tidak sah.", vbInformation, "Info"
                
                Exit Sub
                
        End Select

    End If
    
    Call Frm85_Initial_Setting
    Call Frm85_background_color
    
    Frm101.L43_Text = -1
    Frm85.L79_Text = 0
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    GM_NEXT_PREV = 0
    
    Frm101.L5_Text = Frm101.CBB1 'Purity
    Frm101.L6_Text = Frm101.CBB2 'Kategori Produk
    Frm101.L34_Text = Frm101.CBB3 'Dulang
    Frm101.L37_Text = Frm101.CBB4 'Supplier
    Frm101.L45_Text = Frm101.CBB6 'Cara jualan
    Frm101.L46_Text = Frm101.CBB7 'Cawangan
    Frm101.L47_Text = Frm101.CBB8 'Pekerja
    
    If Frm101.CBB5 = "Semua Jenis" Then 'Jenis Jualan
        Frm101.L44_Text = 2 '0 : Jualan biasa , 1 : Jualan kepada agen , 2 : Semua jenis jualan
    ElseIf Frm101.CBB5 = "Jualan Biasa" Then
        Frm101.L44_Text = 0 '0 : Jualan biasa , 1 : Jualan kepada agen , 2 : Semua jenis jualan
    ElseIf Frm101.CBB5 = "Jualan Kepada Agen" Then
        Frm101.L44_Text = 1 '0 : Jualan biasa , 1 : Jualan kepada agen , 2 : Semua jenis jualan
    End If
    
    Frm101.L7_Text = Frm101.DTPicker1 'Tarikh Mula
    Frm101.L8_Text = Frm101.DTPicker2 'Tarikh Akhir
    
    Frm101.L33_Text = 0 '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
    
    If Frm101.CB2 = 1 Then 'Report Belian
        Call Frm85_Header_Report_Belian
        GM_REPORT_MODE = 0 '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
        Call Frm85_report_belian_page
    End If
    If Frm101.CB3 = 1 Then 'Report Jualan
        Call Frm85_Header_Report_Jualan
        GM_REPORT_MODE = 2 '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
        Call Frm85_Report_Jualan_page
    End If
    If Frm101.CB4 = 1 Then 'Report Buyback
        GM_REPORT_MODE = 5 '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
        Call Frm85_Header_Report_Buyback
        Call Frm85_report_buyback_page
    End If
    If Frm101.CB5 = 1 Then 'Report Stok
        GM_REPORT_MODE = 10 '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
        Call Frm85_Header_Report_Stok
        Call Frm85_report_stok_page
    End If
    If Frm101.CB6 = 1 Then 'Report Potong
        GM_REPORT_MODE = 11 '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
        Call Frm85_Header_Report_Potong
        Call Frm85_report_potong_page
    End If
    If Frm101.CB9 = 1 Then 'Report Jualan Secara Ansuran
        GM_REPORT_MODE = 12 '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
        Call Frm85_Header_Report_Ansuran
        'Call Frm85_Report_Ansuran
        Call Frm85_report_ansuran_page
    End If
    If Frm101.CB10 = 1 Then 'Report Jualan Secara Tempahan
        GM_REPORT_MODE = 13 '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
        Call Frm85_Header_Report_Tempahan
        'Call Frm85_Report_Tempahan
        Call Frm85_report_tempahan_page
    End If
    If Frm101.CB11 = 1 Then 'Report Buyback
        GM_REPORT_MODE = 6 '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
        Call Frm85_Header_Report_belian_gb
        Call Frm85_report_belian_gb_page
    End If
    If Frm101.CB12 = 1 Then 'Report Buyback
        GM_REPORT_MODE = 7 '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
        Call Frm85_Header_Report_buyback_gb
        Call Frm85_report_buyback_gb_page
    End If
    If Frm101.CB13 = 1 Then 'Report trade in agen
        'Call Frm85_header_report_trade_in_agen
        'Call Frm85_report_trade_in_agen
        Call Frm85_header_report_trade_in_susut_nilai
        Call Frm85_report_trade_in_susut_nilai
    End If
End If
End Sub
Private Sub CMD2_Click()
'on error resume next
If Frm101.TB1 = vbNullString Or (Frm101.TB1 <> vbNullString And Not IsNumeric(Frm101.TB1)) Then
    MsgBox "Sila Masukkan [Berat]. Hanya NOMBOR dibenarkan dalam ruangan ini.", vbExclamation, "Info"
    Exit Sub
End If

If Frm101.TB1 <> vbNullString Then

    If InStr(1, Frm101.TB1, "*") <> 0 Or InStr(1, Frm101.TB1, "/") <> 0 Or InStr(1, Frm101.TB1, "\") <> 0 Or InStr(1, Frm101.TB1, "'") <> 0 Then
        MsgBox "Berat mengandungi simbol yang tidak sah.", vbExclamation, "Info"
        Exit Sub
    End If

End If

Call Frm85_Initial_Setting
Frm101.L46_Text = Frm101.CBB7 'Cawangan
Frm101.L33_Text = 1 '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
Frm101.L5_Text = Format(Frm101.TB1, "0.00") 'Berat (g)
Frm101.L43_Text = -1
GM_NEXT_PREV = 0

GM_REPORT_MODE = 1 '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
'Call Frm85_search_berat
Call Frm85_Header_Report_Belian
Call Frm85_search_berat_page
End Sub
Private Sub CMD3_Click()
'on error resume next
If Frm101.CB7 = 0 And Frm101.CB8 = 0 Then

    MsgBox "Sila Pilih Samada [No. Invoice Jualan Atau Voucher Buyback/Trade In].", vbExclamation, "Info"
    Exit Sub
    
End If

If Frm101.TB2 = vbNullString Then
    MsgBox "Sila Masukkan [No. Voucher].", vbExclamation, "Info"
    Exit Sub
End If
If Frm101.TB2 <> vbNullString Then

    If InStr(1, Frm101.TB2, "*") <> 0 Or InStr(1, Frm101.TB2, "/") <> 0 Or InStr(1, Frm101.TB2, "\") <> 0 Or InStr(1, Frm101.TB2, "'") <> 0 Then
        
        If Frm101.CB7 = 1 Then
        
            MsgBox "Invoice jualan mengandungi simbol yang tidak sah.", vbExclamation, "Info"
            
        ElseIf Frm101.CB8 = 1 Then
            
            MsgBox "Voucher trade in mengandungi simbol yang tidak sah.", vbExclamation, "Info"
            
        End If
        
        Exit Sub
    End If

End If

Frm101.L46_Text = Frm101.CBB7 'Cawangan
Frm101.L5_Text = Frm101.TB2 'No. Resit
Call Frm85_Initial_Setting

If Frm101.CB7 = 1 Then
    Frm101.L33_Text = 2 '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
    Frm101.L43_Text = -1
    GM_NEXT_PREV = 0
    
    GM_REPORT_MODE = 3 '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_Header_Report_Jualan
    Call Frm85_carian_jualan_page
ElseIf Frm101.CB8 = 1 Then
    Frm101.L33_Text = 3 '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
    Frm101.L43_Text = -1
    GM_NEXT_PREV = 0
    
    GM_REPORT_MODE = 4 '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
    Call Frm85_Header_Report_Buyback
    Call Frm85_carian_buyback_page
End If
End Sub
Private Sub CMD4_Click()
'on error resume next
If Frm101.CB2 = 0 And Frm101.CB3 = 0 And Frm101.CB4 = 0 And Frm101.CB5 = 0 And Frm101.CB6 = 0 And Frm101.CB9 = 0 And Frm101.CB10 = 0 And Frm101.CB11 = 0 And Frm101.CB12 = 0 And Frm101.CB13 = 0 Then
    MsgBox "Sila Buat Pilihan Tetapan Report.", vbInformation, "Info"
    Exit Sub
End If

Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report. Teruskan ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then

    If Frm101.CB3 = 1 Then 'Report Jualan
    
        LM_REPORT_JUALAN = vbNullString ' 1 : Barang Baru Sahaja , 2 : Barang Trade In Sahaja , 3 : Semua
        
        LM_REPORT_JUALAN = InputBox("Sila pilih jenis barang jualan." & _
                vbCrLf & _
                vbCrLf & vbTab & "1 - Barang Baru Sahaja" & _
                vbCrLf & vbTab & "2 - Barang Trade In Sahaja" & _
                vbCrLf & vbTab & "3 - Semua", "Jenis Jualan")
                 
        Select Case LM_REPORT_JUALAN
            Case "1"
        
                G_JENIS_JUALAN = "Barang Baru Sahaja"
                
            Case "2"
                
                G_JENIS_JUALAN = "Barang Trade In Sahaja"
                
            Case "3"

                G_JENIS_JUALAN = "Barang Baru Dan Barang Trade In"
                
            Case Else
            
                MsgBox "Tiada pilihan dibuat atau pilihan yang tidak sah.", vbInformation, "Info"
                
                Exit Sub
                
        End Select

    End If
    
    Frm101.L5_Text = Frm101.CBB1 'Purity
    Frm101.L6_Text = Frm101.CBB2 'Kategori Produk
    Frm101.L34_Text = Frm101.CBB3 'Dulang
    Frm101.L37_Text = Frm101.CBB4 'Supplier
    Frm101.L45_Text = Frm101.CBB6 'Cara jualan
    Frm101.L46_Text = Frm101.CBB7 'Cawangan
    Frm101.L47_Text = Frm101.CBB8 'Pekerja
    
    If Frm101.CBB5 = "Semua Jenis" Then 'Jenis Jualan
        Frm101.L44_Text = 2 '0 : Jualan biasa , 1 : Jualan kepada agen , 2 : Semua jenis jualan
    ElseIf Frm101.CBB5 = "Jualan Biasa" Then
        Frm101.L44_Text = 0 '0 : Jualan biasa , 1 : Jualan kepada agen , 2 : Semua jenis jualan
    ElseIf Frm101.CBB5 = "Jualan Kepada Agen" Then
        Frm101.L44_Text = 1 '0 : Jualan biasa , 1 : Jualan kepada agen , 2 : Semua jenis jualan
    End If
    
    Frm101.L7_Text = Frm101.DTPicker1 'Tarikh Mula
    Frm101.L8_Text = Frm101.DTPicker2 'Tarikh Akhir
    
    Frm101.L33_Text = 0 '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
    
    If Frm101.CB2 = 1 Then 'Report Belian
        Call Frm85_summary_report_belian
    End If
    If Frm101.CB3 = 1 Then 'Report Jualan
        Call Frm85_summary_report_jualan
    End If
    If Frm101.CB4 = 1 Then 'Report Buyback
        Call Frm85_summary_report_buyback
    End If
    If Frm101.CB5 = 1 Then 'Report Stok
        Call Frm85_summary_report_stok
    End If
    If Frm101.CB6 = 1 Then 'Report Potong
        Call Frm85_summary_report_potong
    End If
    If Frm101.CB9 = 1 Then 'Report Jualan Secara Ansuran
        Call Frm85_summary_report_ansuran
    End If
    If Frm101.CB10 = 1 Then 'Report Jualan Secara Tempahan
        Call Frm85_summary_report_tempahan
    End If
    If Frm101.CB11 = 1 Then 'Report Belian Gold Bar
        Call Frm85_summary_report_belian_gb
    End If
    If Frm101.CB12 = 1 Then 'Report Buyback Gold Bar
        Call Frm85_summary_report_buyback_gb
    End If
    If Frm101.CB13 = 1 Then 'Report trade in agen
        'Call Frm85_summary_report_trade_in
    End If
End If
End Sub
Private Sub CMD5_Click()
'on error resume next
If Frm101.TB3 = vbNullString Then
    MsgBox "Sila Masukkan [No. Invoice Dari Supplier].", vbExclamation, "Info"
    Exit Sub
End If

If Frm101.TB3 <> vbNullString Then

    If InStr(1, Frm101.TB3, "*") <> 0 Or InStr(1, Frm101.TB3, "/") <> 0 Or InStr(1, Frm101.TB3, "\") <> 0 Or InStr(1, Frm101.TB3, "'") <> 0 Then
        MsgBox "No. invoice mengandungi simbol yang tidak sah.", vbExclamation, "Info"
        Exit Sub
    End If

End If

Frm101.L46_Text = Frm101.CBB7 'Cawangan

Call Frm85_Initial_Setting

Frm101.L33_Text = 4 '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)  , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk
Frm101.L5_Text = UCase(Frm101.TB3) 'No. Invoice Suplier
Frm101.L43_Text = -1
GM_NEXT_PREV = 0

GM_REPORT_MODE = 8 '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan  , 6 : Report mengikut no invoice supplier

Call Frm85_Header_Report_Belian
Call Frm85_search_invoice_supplier_page
End Sub
Private Sub CMD6_Click()
'on error resume next
If Frm101.CB2 = 0 And Frm101.CB3 = 0 And Frm101.CB4 = 0 And Frm101.CB5 = 0 And Frm101.CB6 = 0 And Frm101.CB9 = 0 And Frm101.CB10 = 0 And Frm101.CB11 = 0 And Frm101.CB12 = 0 Then
    MsgBox "Sila Buat Pilihan Tetapan Report.", vbInformation, "Info"
    Exit Sub
End If

If Frm101.TB4 = vbNullString Then
    MsgBox "Sila Masukkan [No. Siri Produk].", vbExclamation, "Info"
    Exit Sub
End If

If Frm101.TB4 <> vbNullString Then

    If InStr(1, Frm101.TB4, "*") <> 0 Or InStr(1, Frm101.TB4, "/") <> 0 Or InStr(1, Frm101.TB4, "\") <> 0 Or InStr(1, Frm101.TB4, "'") <> 0 Then
        MsgBox "No. siri produk mengandungi simbol yang tidak sah.", vbExclamation, "Info"
        Exit Sub
    End If

End If

Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report. Teruskan ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    Call Frm85_Initial_Setting
    
    Frm101.L43_Text = -1
    Frm85.L79_Text = 0
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm101.L5_Text = UCase(Frm101.TB4) 'No. Siri Produk
    Frm101.L46_Text = Frm101.CBB7 'Cawangan
    GM_NEXT_PREV = 0
    
    Frm101.L33_Text = 5 '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
    
    If Frm101.CB2 = 1 Then 'Report Belian
        GM_REPORT_MODE = 9 '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
        
        Call Frm85_Header_Report_Belian
        Call Frm85_report_belian_barcode
    End If
    If Frm101.CB3 = 1 Then 'Report Jualan
        GM_REPORT_MODE = 9 '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
        
        Call Frm85_Header_Report_Jualan
        Call Frm85_Report_Jualan_barcode
    End If
    If Frm101.CB4 = 1 Then 'Report Buyback
        GM_REPORT_MODE = 9 '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
        Frm101.L33_Text = 6 '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
        Call Frm85_Header_Report_Buyback
        Call Frm85_report_buyback_barcode
    End If
    If Frm101.CB5 = 1 Then 'Report Stok
        GM_REPORT_MODE = 9 '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
        
        Call Frm85_Header_Report_Stok
        Call Frm85_report_stok_barcode
    End If
    If Frm101.CB6 = 1 Then 'Report Potong
        GM_REPORT_MODE = 9 '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
        
        Call Frm85_Header_Report_Potong
        Call Frm85_report_potong_barcode
    End If
    If Frm101.CB9 = 1 Then 'Report Jualan Secara Ansuran
        GM_REPORT_MODE = 9 '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
        
        Call Frm85_Header_Report_Ansuran
        Call Frm85_report_ansuran_barcode
    End If
    If Frm101.CB10 = 1 Then 'Report Jualan Secara Tempahan
        GM_REPORT_MODE = 9 '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
        
        Call Frm85_Header_Report_Tempahan
        Call Frm85_report_tempahan_barcode
    End If
    If Frm101.CB11 = 1 Then 'Report Buyback
        GM_REPORT_MODE = 9 '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
        Frm101.L33_Text = 7 '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
        Call Frm85_Header_Report_belian_gb
        Call Frm85_report_belian_gb_barcode
    End If
    If Frm101.CB12 = 1 Then 'Report Buyback
        GM_REPORT_MODE = 9 '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
        Frm101.L33_Text = 8 '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
        Call Frm85_Header_Report_buyback_gb
        Call Frm85_report_buyback_gb_barcode
    End If
End If

End Sub
Private Sub Form_Load()
'on error resume next
Frm101.DTPicker1 = DateTime.Date
Frm101.DTPicker2 = DateTime.Date

Frm101.L5_Text = vbNullString
Frm101.L6_Text = vbNullString
Frm101.L7_Text = vbNullString
Frm101.L8_Text = vbNullString
Frm101.L9_Text = 0
Frm85.L10_Text = vbNullString
Frm85.L79_Text = 0
Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir

Frm101.CB2 = 1
Frm101.CB7 = 1
Frm101.CB8 = 0

Frm101.CBB1.Clear
Frm101.CBB2.Clear
Frm101.CBB3.Clear
Frm101.CBB4.Clear
Frm101.CBB5.Clear
Frm101.CBB6.Clear

Frm101.CBB1.AddItem "Semua Purity"
Frm101.CBB2.AddItem "Semua Kategori Produk"
Frm101.CBB3.AddItem "Semua Dulang"
Frm101.CBB4.AddItem "Semua Supplier"
Frm101.CBB5.AddItem "Semua Jenis"
Frm101.CBB6.AddItem "Kedai & Online"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database order by supplier ASC , Kod_Metal_Purity ASC , kategori_Produk ASC , SenaraiDulang ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!Kod_Metal_Purity) Then Frm101.CBB1.AddItem rs!Kod_Metal_Purity
    If Not IsNull(rs!kategori_Produk) Then Frm101.CBB2.AddItem rs!kategori_Produk
    If Not IsNull(rs!SenaraiDulang) Then Frm101.CBB3.AddItem rs!SenaraiDulang
    If Not IsNull(rs!supplier) Then Frm101.CBB4.AddItem rs!supplier
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm101.CBB7.Clear

Frm101.CBB7.AddItem "Semua cawangan"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where status='" & 1 & "' order by cawangan ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!cawangan) Then Frm101.CBB7.AddItem rs!cawangan
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm101.CBB8.Clear
Frm101.CBB8.AddItem "Semua"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where status='" & "Aktif" & "' AND InvestorSmall = 0 AND InvestorBig = 0", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    If rs!user_level = 1 Or rs!user_level = 2 Or rs!user_level = 3 Or rs!user_level = G_LEVEL_USER Then Frm101.CBB8.AddItem rs!Samaran & "  |  " & rs!NoPekerja
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm101.CBB8 = "Semua"

Frm101.CBB5.AddItem "Jualan Biasa"
Frm101.CBB5.AddItem "Jualan Kepada Agen"

Frm101.CBB6.AddItem "Kedai Sahaja"
Frm101.CBB6.AddItem "Online Sahaja"

Frm101.CBB1 = "Semua Purity"
Frm101.CBB2 = "Semua Kategori Produk"
Frm101.CBB3 = "Semua Dulang"
Frm101.CBB4 = "Semua Supplier"
Frm101.CBB5 = "Semua Jenis"
Frm101.CBB6 = "Kedai & Online"
Frm101.CBB7 = "Semua cawangan"

If MDI_frm1.L4_Text <> "HQ" And MDI_frm1.L4_Text <> "Developer" Then
    Frm101.CBB7 = MDI_frm1.L20_Text
    Frm101.CBB7.Enabled = False
Else
    Frm101.CBB7.Enabled = True
End If

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

Frm101.TB1 = vbNullString
Frm101.TB2 = vbNullString
Frm101.TB3 = vbNullString
Frm101.TB4 = vbNullString
End Sub
Private Sub L3_Text_Click()
'on error resume next
If Frm101.Pic1.Visible = False Then
    Call Frm101_initial_setting
    
    Frm101.Pic1.Visible = True
Else
    Frm101.Pic1.Visible = False
End If
End Sub
Private Sub L31_Text_Click()
'On Error Resume Next
If Frm101.Pic7.Visible = False Then
    Frm101.L32_Text.Top = 3000
    Frm101.L35_Text.Top = 3240
    
    Frm101.Pic7.Visible = True
    Frm101.Pic8.Visible = False
    Frm101.Pic9.Visible = False
    
    Frm101.TB1.SetFocus
Else
    Frm101.Pic7.Visible = False
    Frm101.L32_Text.Top = 1200
    Frm101.L35_Text.Top = 1440
End If
End Sub
Private Sub L32_Text_Click()
'On Error Resume Next
If Frm101.Pic8.Visible = False Then
    Frm101.L32_Text.Top = 1200
    Frm101.L35_Text.Top = 4560
    
    Frm101.Pic8.Top = 1440
    Frm101.Pic8.Visible = True
    Frm101.Pic9.Visible = False
    Frm101.Pic7.Visible = False
    
    Frm101.TB2.SetFocus
Else
    Frm101.Pic8.Visible = False
    Frm101.L32_Text.Top = 1200
    Frm101.L35_Text.Top = 1440
End If
End Sub
Private Sub L35_Text_Click()
'On Error Resume Next
If Frm101.Pic9.Visible = False Then
    Frm101.L32_Text.Top = 1200
    Frm101.L35_Text.Top = 1440
    
    Frm101.Pic9.Visible = True
    Frm101.Pic7.Visible = False
    Frm101.Pic8.Visible = False
    
    Frm101.TB3.SetFocus
Else
    Frm101.Pic9.Visible = False
    Frm101.L32_Text.Top = 1200
    Frm101.L35_Text.Top = 1440
End If
End Sub

Private Sub L4_Text_Click()

End Sub

Private Sub Tmr1_Timer()
'On Error Resume Next
Frm101.L1_Text = DateTime.Date
Frm101.L2_Text = DateTime.Time$
End Sub
