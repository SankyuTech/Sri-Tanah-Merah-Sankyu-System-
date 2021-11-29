VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm111 
   Caption         =   "Tetapan sistem"
   ClientHeight    =   13035
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   22170
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13035
   ScaleWidth      =   22170
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tetapan Sistem"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11415
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   16815
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tetapan 7"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   8040
         TabIndex        =   48
         Top             =   720
         Width           =   8295
         Begin VB.TextBox TB56 
            Height          =   360
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   129
            Text            =   "TB56"
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox TB55 
            Height          =   360
            Left            =   5760
            MaxLength       =   10
            TabIndex        =   127
            Text            =   "TB55"
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox TB54 
            Height          =   360
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   126
            Text            =   "TB54"
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox TB33 
            Height          =   360
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   51
            Text            =   "TB33"
            Top             =   2400
            Width           =   1575
         End
         Begin VB.TextBox TB51 
            Height          =   360
            Left            =   5400
            MaxLength       =   10
            TabIndex        =   50
            Text            =   "TB51"
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Caj Pertukaran : RM/g"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   130
            Top             =   885
            Width           =   5535
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Kadar Trade In : RM/g                          Kadar Buyback : RM/g"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   128
            Top             =   525
            Width           =   5535
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "Harga GDN  (RM/g) :                           Harga GRN (RM/g) :"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   52
            Top             =   2445
            Width           =   5535
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   $"Frm111.frx":0000
            Height          =   615
            Left            =   240
            TabIndex        =   49
            Top             =   1800
            Width           =   7935
         End
      End
      Begin VB.CheckBox CB15 
         BackColor       =   &H008080FF&
         Caption         =   "Generate Barcode"
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
         Left            =   3960
         TabIndex        =   124
         Top             =   390
         Width           =   200
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tetapan 13"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8040
         TabIndex        =   121
         Top             =   9240
         Width           =   8295
         Begin VB.CheckBox CB14 
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
            TabIndex        =   122
            Top             =   405
            Width           =   200
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Tandakan di sini jika ""Invoice Tidak Rasmi"" sebagai default invoice."
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   600
            TabIndex        =   123
            Top             =   360
            Width           =   7095
         End
      End
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
         Left            =   10680
         MouseIcon       =   "Frm111.frx":0088
         MousePointer    =   99  'Custom
         Picture         =   "Frm111.frx":0392
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   10200
         Width           =   2775
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tetapan 11"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   8040
         TabIndex        =   69
         Top             =   6600
         Width           =   8295
         Begin VB.TextBox TB40 
            Height          =   360
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   72
            Text            =   "TB40"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox TB41 
            Height          =   360
            Left            =   2640
            MaxLength       =   10
            TabIndex        =   71
            Text            =   "TB41"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox TB42 
            Height          =   360
            Left            =   4080
            MaxLength       =   10
            TabIndex        =   70
            Text            =   "TB42"
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Tael  :                Public  :               SA  : "
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   480
            TabIndex        =   74
            Top             =   765
            Width           =   7695
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Tetapan untuk pengiraan harga belian barang kemas terpakai."
            Height          =   255
            Left            =   240
            TabIndex        =   73
            Top             =   360
            Width           =   7935
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tetapan 10"
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
         Left            =   8040
         TabIndex        =   65
         Top             =   5520
         Width           =   8295
         Begin VB.TextBox TB35 
            Height          =   360
            Left            =   3840
            MaxLength       =   10
            TabIndex        =   66
            Text            =   "TB35"
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Diskaun / gram                         RM  :"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   480
            TabIndex        =   68
            Top             =   525
            Width           =   3495
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Penetapan kadar diskaun per gram bagi penggunaan kupon."
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   240
            Width           =   7935
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tetapan 9"
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
         Left            =   8040
         TabIndex        =   57
         Top             =   3600
         Width           =   8295
         Begin VB.TextBox TB53 
            Height          =   360
            Left            =   3840
            MaxLength       =   10
            TabIndex        =   62
            Text            =   "TB53"
            Top             =   1320
            Width           =   2895
         End
         Begin VB.CheckBox CB10 
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
            TabIndex        =   60
            Top             =   720
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
            Left            =   3765
            TabIndex        =   59
            Top             =   720
            Width           =   200
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Tetapan top margin hanya digunakan bagi pre-printed invoice."
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   1050
            Width           =   7935
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Top Margin                                       :"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   360
            TabIndex        =   63
            Top             =   1365
            Width           =   3495
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Pre-printed                                           Header dari sistem"
            Height          =   255
            Left            =   480
            TabIndex        =   61
            Top             =   680
            Width           =   7935
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Sila buat pilihan header invoice kedai."
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   360
            Width           =   7935
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tetapan 8"
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
         Left            =   8040
         TabIndex        =   53
         Top             =   2280
         Width           =   8295
         Begin VB.TextBox TB34 
            Height          =   360
            Left            =   3720
            MaxLength       =   10
            TabIndex        =   54
            Text            =   "TB34"
            Top             =   720
            Width           =   2895
         End
         Begin VB.Label Label41 
            BackStyle       =   0  'Transparent
            Caption         =   "Penetapan kadar bagi pengiraan komisyen upah kepada agen dropship."
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   360
            Width           =   7935
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "Kadar Komisyen Upah               (%)  :"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   360
            TabIndex        =   55
            Top             =   765
            Width           =   3495
         End
      End
      Begin VB.TextBox Text2 
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
         Left            =   12240
         TabIndex        =   47
         Top             =   360
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tetapan 6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   240
         TabIndex        =   41
         Top             =   9480
         Width           =   7695
         Begin VB.TextBox TB32 
            Height          =   360
            Left            =   3600
            MaxLength       =   10
            TabIndex        =   45
            Text            =   "TB32"
            Top             =   1200
            Width           =   2895
         End
         Begin VB.CheckBox CB6 
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
            TabIndex        =   43
            Top             =   885
            Width           =   200
         End
         Begin VB.Label Label80 
            BackStyle       =   0  'Transparent
            Caption         =   "Kadar Diskaun                         (%)  :"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   240
            TabIndex        =   46
            Top             =   1245
            Width           =   3495
         End
         Begin VB.Label Label78 
            BackStyle       =   0  'Transparent
            Caption         =   "Diskaun Bagi Jualan Barang Kemas / Permata"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   600
            TabIndex        =   44
            Top             =   840
            Width           =   4575
         End
         Begin VB.Label Label69 
            BackStyle       =   0  'Transparent
            Caption         =   "Sila tandakan jika di bawah jika pihak kedai ada memberi diskaun dalam urusan jualan barang kemas / permata."
            Height          =   615
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   7935
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tetapan 5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   37
         Top             =   8100
         Width           =   7695
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
            TabIndex        =   39
            Top             =   1005
            Width           =   200
         End
         Begin VB.Label Label90 
            BackStyle       =   0  'Transparent
            Caption         =   "Upah Bagi Jualan Emas"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   600
            TabIndex        =   40
            Top             =   960
            Width           =   3855
         End
         Begin VB.Label Label79 
            BackStyle       =   0  'Transparent
            Caption         =   $"Frm111.frx":295C
            Height          =   855
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   7335
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tetapan 4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   31
         Top             =   6720
         Width           =   7695
         Begin VB.TextBox TB20 
            Height          =   360
            Left            =   3075
            MaxLength       =   10
            TabIndex        =   34
            Text            =   "TB20"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox TB29 
            Height          =   360
            Left            =   6840
            MaxLength       =   10
            TabIndex        =   33
            Text            =   "TB29"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox TB52 
            Height          =   360
            Left            =   3765
            MaxLength       =   10
            TabIndex        =   32
            Text            =   "TB52"
            Top             =   765
            Width           =   735
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Kadar Komisen Barang Kemas % :           Kadar Komisen Barang Permata% :"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   120
            TabIndex        =   36
            Top             =   405
            Width           =   8055
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Kadar Komisen Pekerja Per Gram: (RM/g)"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   810
            Width           =   8055
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tetapan 12"
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
         Left            =   8040
         TabIndex        =   22
         Top             =   7920
         Width           =   8295
         Begin VB.CheckBox CB7 
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
            Left            =   2520
            TabIndex        =   78
            Top             =   645
            Width           =   200
         End
         Begin VB.CheckBox CB5 
            Caption         =   "Scanner Mode"
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
            TabIndex        =   77
            Top             =   840
            Width           =   200
         End
         Begin VB.TextBox TB24 
            Height          =   360
            Left            =   7200
            MaxLength       =   10
            TabIndex        =   76
            Top             =   600
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.CheckBox CB3 
            Caption         =   "Scanner Mode"
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
            TabIndex        =   75
            Top             =   645
            Width           =   200
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Penetapan GST Belian"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   480
            TabIndex        =   81
            Top             =   840
            Width           =   3375
         End
         Begin VB.Label Label91 
            BackStyle       =   0  'Transparent
            Caption         =   "Penetapan GST Jualan       GST termasuk dalam harga jualan               "
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   480
            TabIndex        =   80
            Top             =   600
            Width           =   6615
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Tetapan untuk cukai GST pada menu belian stok dan jualan."
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   360
            Width           =   7935
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tetapan 3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   240
         TabIndex        =   15
         Top             =   5040
         Width           =   7695
         Begin VB.CheckBox CB13 
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
            Left            =   3000
            TabIndex        =   29
            Top             =   1350
            Width           =   200
         End
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
            TabIndex        =   28
            Top             =   1350
            Width           =   200
         End
         Begin VB.CheckBox CB8 
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
            TabIndex        =   24
            Top             =   650
            Width           =   200
         End
         Begin VB.CheckBox CB9 
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
            Left            =   3000
            TabIndex        =   23
            Top             =   650
            Width           =   200
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Mengikut berat barang              Mengikut upah yang telah ditetapkan"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   480
            TabIndex        =   30
            Top             =   1320
            Width           =   6735
         End
         Begin VB.Label Label43 
            BackStyle       =   0  'Transparent
            Caption         =   "Tetapan cara pengiraan upah bagi penerimaan stok dari supplier."
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Width           =   6975
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Mengikut berat barang              Mengikut upah yang telah ditetapkan"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   480
            TabIndex        =   26
            Top             =   600
            Width           =   6735
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Tetapan cara pengiraan upah jualan bagi barang kemas."
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   6975
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tetapan 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   240
         TabIndex        =   14
         Top             =   2550
         Width           =   7695
         Begin VB.TextBox TB28 
            Height          =   360
            Left            =   4560
            MaxLength       =   10
            TabIndex        =   18
            Top             =   1920
            Width           =   1695
         End
         Begin VB.TextBox TB16 
            Height          =   360
            Left            =   4560
            MaxLength       =   10
            TabIndex        =   17
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "Potongan jika harga belian barangan kedai oleh pelanggan melebihi dari nilaian barang trade in."
            Height          =   615
            Left            =   480
            TabIndex        =   21
            Top             =   1515
            Width           =   7215
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "Kadar Potongan Barang Trade In            % :"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   600
            TabIndex        =   20
            Top             =   1965
            Width           =   4695
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Spread Trade In Cash                           % :"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   600
            TabIndex        =   19
            Top             =   1125
            Width           =   4695
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   $"Frm111.frx":29F0
            Height          =   615
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   7215
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tetapan 1"
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
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   7695
         Begin VB.TextBox TB26 
            Height          =   360
            Left            =   4560
            MaxLength       =   10
            TabIndex        =   11
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox TB25 
            Height          =   360
            Left            =   4560
            MaxLength       =   10
            TabIndex        =   10
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label89 
            BackStyle       =   0  'Transparent
            Caption         =   "Had Penurunan Harga Barang Per Item      RM :"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   360
            TabIndex        =   13
            Top             =   1365
            Width           =   4575
         End
         Begin VB.Label Label86 
            BackStyle       =   0  'Transparent
            Caption         =   "Had Penurunan Harga Barang Per Gram  RM/g :"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   360
            TabIndex        =   12
            Top             =   1005
            Width           =   4575
         End
         Begin VB.Label Label87 
            BackStyle       =   0  'Transparent
            Caption         =   "Ini adalah tetapan bagi had maksimum penurunan harga yang boleh dibuat oleh pekerja semasa urusan jualan barang kedai."
            Height          =   615
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   7215
         End
      End
      Begin VB.CheckBox CB2 
         BackColor       =   &H008080FF&
         Caption         =   "Generate Barcode"
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
         Left            =   2160
         TabIndex        =   5
         Top             =   390
         Width           =   200
      End
      Begin VB.CheckBox CB1 
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
         Top             =   390
         Width           =   200
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Cetak Barcode Trade In"
         Height          =   255
         Left            =   4200
         TabIndex        =   125
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "Cetak Barcode"
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "Scanner Mode"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Penetapan Cas Kad Kredit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11415
      Left            =   1560
      TabIndex        =   100
      Top             =   600
      Visible         =   0   'False
      Width           =   16815
      Begin VB.CommandButton CMD6 
         Caption         =   "Next"
         Height          =   810
         Left            =   9120
         MouseIcon       =   "Frm111.frx":2A88
         MousePointer    =   99  'Custom
         Picture         =   "Frm111.frx":2D92
         Style           =   1  'Graphical
         TabIndex        =   119
         ToolTipText     =   "Tutup senarai ini."
         Top             =   10440
         Width           =   1095
      End
      Begin VB.CommandButton CMD5 
         Caption         =   "Back"
         Height          =   810
         Left            =   7920
         MouseIcon       =   "Frm111.frx":3E5C
         MousePointer    =   99  'Custom
         Picture         =   "Frm111.frx":4166
         Style           =   1  'Graphical
         TabIndex        =   118
         ToolTipText     =   "Tutup senarai ini."
         Top             =   10440
         Width           =   1095
      End
      Begin VB.TextBox TB36 
         Height          =   360
         Left            =   2400
         MaxLength       =   100
         TabIndex        =   105
         Text            =   "TB36"
         Top             =   645
         Width           =   6375
      End
      Begin VB.TextBox TB37 
         Height          =   360
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   104
         Text            =   "TB37"
         Top             =   1005
         Width           =   1695
      End
      Begin VB.CommandButton CMD2 
         BackColor       =   &H000080FF&
         Caption         =   "Simpan Tetapan"
         Height          =   405
         Left            =   3240
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm111.frx":5230
         MousePointer    =   99  'Custom
         TabIndex        =   103
         ToolTipText     =   "Simpan tetapan sistem yang telah dibuat."
         Top             =   1560
         Width           =   2385
      End
      Begin VB.CommandButton CMD3 
         BackColor       =   &H000080FF&
         Caption         =   "Simpan Tetapan"
         Height          =   405
         Left            =   2040
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm111.frx":553A
         MousePointer    =   99  'Custom
         TabIndex        =   102
         ToolTipText     =   "Simpan tetapan sistem yang telah dibuat."
         Top             =   1560
         Width           =   2385
      End
      Begin VB.CommandButton CMD4 
         BackColor       =   &H000080FF&
         Caption         =   "Batal"
         Height          =   405
         Left            =   4560
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm111.frx":5844
         MousePointer    =   99  'Custom
         TabIndex        =   101
         ToolTipText     =   "Simpan tetapan sistem yang telah dibuat."
         Top             =   1560
         Width           =   2385
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   7665
         Left            =   240
         TabIndex        =   106
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   2640
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   13520
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
      Begin VB.Label L12_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L12_Text"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7440
         TabIndex        =   116
         Top             =   10440
         Width           =   735
      End
      Begin VB.Label L11_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L11_Text"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6390
         TabIndex        =   115
         Top             =   10440
         Width           =   855
      End
      Begin VB.Label L10_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L10_Text"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10680
         TabIndex        =   114
         Top             =   6840
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label L9_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L9_Text"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10680
         TabIndex        =   113
         Top             =   6480
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label L7_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai kad kredit/debit dan caj perkhidmatan."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   112
         Top             =   2400
         Width           =   8655
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis kad kredit / debit  :"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   111
         Top             =   690
         Width           =   2415
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Caj perkhidmatan  (%) :"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   110
         Top             =   1050
         Width           =   2415
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan jenis dan caj perkhidmatan bagi kad tersebut."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   109
         Top             =   360
         Width           =   8655
      End
      Begin VB.Label L13_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Anda berada di dalam menu edit data."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   360
         TabIndex        =   108
         Top             =   2040
         Visible         =   0   'False
         Width           =   8415
      End
      Begin VB.Label L14_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L14_Text"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10800
         TabIndex        =   107
         Top             =   5040
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Paparan muka :           /"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5280
         TabIndex        =   117
         Top             =   10440
         Width           =   2415
      End
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Penetapan Mata Ganjaran Keahlian"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11415
      Left            =   14040
      TabIndex        =   83
      Top             =   720
      Visible         =   0   'False
      Width           =   16815
      Begin VB.CommandButton CMD7 
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
         Left            =   1800
         MouseIcon       =   "Frm111.frx":5B4E
         MousePointer    =   99  'Custom
         Picture         =   "Frm111.frx":5E58
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox TB44 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   91
         Text            =   "TB44"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox TB43 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   90
         Text            =   "TB43"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox TB45 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   89
         Text            =   "TB45"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox TB46 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   88
         Text            =   "TB46"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox TB47 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   87
         Text            =   "TB47"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox TB48 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   86
         Text            =   "TB48"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox TB49 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   85
         Text            =   "TB49"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox TB50 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   84
         Text            =   "TB50"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila masukkan kadar pembahagian dan tebusan mata ganjaran kepada setiap kategori pelanggan."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   98
         Top             =   480
         Width           =   8655
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Silver"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   97
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "Ahli biasa"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   96
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Gold"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   95
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Platinum"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   94
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pembahagian mata"
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   1800
         TabIndex        =   93
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tebusan mata"
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   3075
         TabIndex        =   92
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Timer Tmr1 
      Interval        =   80
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "User perlu log out dari sistem (Semua station) bagi memboleh tetapan ini diupdate selepas tetapan ini disimpan."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   240
      TabIndex        =   120
      Top             =   11040
      Width           =   8175
   End
   Begin VB.Label L4_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tetapan Mata Ganjaran"
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
      Left            =   5280
      MouseIcon       =   "Frm111.frx":8422
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label L2_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tetapan Kadar Caj Kad Kredit/debit"
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
      MouseIcon       =   "Frm111.frx":872C
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label L3_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tetapan Sistem"
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
      MouseIcon       =   "Frm111.frx":8A36
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Menu Frm111_PM_menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm111_SM_edit_data 
         Caption         =   "Edit data"
      End
      Begin VB.Menu Frm111_SM_padam_data 
         Caption         =   "Padam data"
      End
   End
End
Attribute VB_Name = "Frm111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB10_Click()
'On Error Resume Next
If Frm111.CB10 = 1 Then
    Frm111.CB11 = 0
End If
End Sub
Private Sub CB11_Click()
'On Error Resume Next
If Frm111.CB11 = 1 Then
    Frm111.CB10 = 0
End If
End Sub
Private Sub CB3_Click()
'On Error Resume Next
If Frm111.CB3 = 1 Then
    Frm111.CB7 = 0
End If
End Sub

Private Sub CB7_Click()
'On Error Resume Next
If Frm111.CB7 = 1 Then
    Frm111.CB3 = 0
End If
End Sub

Private Sub CB8_Click()
'On Error Resume Next
If Frm111.CB8 = 1 Then
    Frm111.CB9 = 0
End If
End Sub
Private Sub CB9_Click()
'On Error Resume Next
If Frm111.CB9 = 1 Then
    Frm111.CB8 = 0
End If
End Sub
Private Sub CB12_Click()
'On Error Resume Next
If Frm111.CB12 = 1 Then
    Frm111.CB13 = 0
End If
End Sub
Private Sub CB13_Click()
'On Error Resume Next
If Frm111.CB13 = 1 Then
    Frm111.CB12 = 0
End If
End Sub
Private Sub CMD1_Click()
'On Error Resume Next
Dim Err(30)

x = 0
If Frm111.TB16 = vbNullString Or (Frm111.TB16 <> vbNullString And Not IsNumeric(Frm111.TB16)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Spread Trade In], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
'If Frm111.TB17 = vbNullString Or (Frm111.TB17 <> vbNullString And Not IsNumeric(Frm111.TB17)) Then
'    x = x + 1
'    Err(x) = "Sila masukkan [Cas Kad Kredit], Hanya NOMBOR dibenarkan dalam ruangan ini."
'End If
'If Frm111.TB27 = vbNullString Or (Frm111.TB27 <> vbNullString And Not IsNumeric(Frm111.TB27)) Then
'    x = x + 1
'    Err(x) = "Sila masukkan [Cas Kad Debit], Hanya NOMBOR dibenarkan dalam ruangan ini."
'End If
If Frm111.TB20 = vbNullString Or (Frm111.TB20 <> vbNullString And Not IsNumeric(Frm111.TB20)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar Komisen Barang Kemas (%)], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm111.TB29 = vbNullString Or (Frm111.TB29 <> vbNullString And Not IsNumeric(Frm111.TB29)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar Komisen Barang Permata (%)], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm111.TB52 = vbNullString Or (Frm111.TB52 <> vbNullString And Not IsNumeric(Frm111.TB52)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar Komisen Pekerja Per Gram], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm111.TB25 = vbNullString Or (Frm111.TB25 <> vbNullString And Not IsNumeric(Frm111.TB25)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Had Penurunan Harga Barang Per Gram], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm111.TB26 = vbNullString Or (Frm111.TB26 <> vbNullString And Not IsNumeric(Frm111.TB26)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Had Penurunan Harga Barang Per Item], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm111.TB28 = vbNullString Or (Frm111.TB28 <> vbNullString And Not IsNumeric(Frm111.TB28)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar Potongan Barang Trade In], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm111.TB33 = vbNullString Or (Frm111.TB33 <> vbNullString And Not IsNumeric(Frm111.TB33)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Harga emas per gram 999.9 bagi GDN], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm111.TB51 = vbNullString Or (Frm111.TB51 <> vbNullString And Not IsNumeric(Frm111.TB51)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Harga emas per gram 999.9 bagi GRN], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm111.TB34 = vbNullString Or (Frm111.TB34 <> vbNullString And Not IsNumeric(Frm111.TB34)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar Komisyen Upah (%)], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If

If Frm111.CB3 = 1 Then
    If Frm111.TB24 = vbNullString Or (Frm111.TB24 <> vbNullString And Not IsNumeric(Frm111.TB24)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Kadar GST], Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If
If Frm111.CB6 = 1 Then
    If Frm111.TB32 = vbNullString Or (Frm111.TB32 <> vbNullString And Not IsNumeric(Frm111.TB32)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Kadar Diskaun], Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
End If
If Frm111.TB24 <> vbNullString And Not IsNumeric(Frm111.TB24) Then
    x = x + 1
    Err(x) = "[Kadar GST] yang tidak sah, Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm111.CB8 = 0 And Frm111.CB9 = 0 Then
    x = x + 1
    Err(x) = "Sila pilih cara pengiraan upah jualan bagi barang kemas."
End If
If Frm111.CB10 = 0 And Frm111.CB11 = 0 Then
    x = x + 1
    Err(x) = "Sila pilih jenis header invoice."
End If
If Frm111.TB35 = vbNullString Or (Frm111.TB35 <> vbNullString And Not IsNumeric(Frm111.TB35)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Diskaun / gram], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm111.TB40 = vbNullString Or (Frm111.TB40 <> vbNullString And Not IsNumeric(Frm111.TB40)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Tael], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm111.TB41 = vbNullString Or (Frm111.TB41 <> vbNullString And Not IsNumeric(Frm111.TB41)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Public], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm111.TB42 = vbNullString Or (Frm111.TB42 <> vbNullString And Not IsNumeric(Frm111.TB42)) Then
    x = x + 1
    Err(x) = "Sila masukkan [SA], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm111.CB12 = 0 And Frm111.CB13 = 0 Then
    x = x + 1
    Err(x) = "Sila pilih cara pengiraan upah bagi penerimaan stok dari supplier."
End If
If Frm111.TB53 = vbNullString Or (Frm111.TB53 <> vbNullString And Not IsNumeric(Frm111.TB53)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Top Margin], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm111.TB54 = vbNullString Or (Frm111.TB54 <> vbNullString And Not IsNumeric(Frm111.TB54)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar Trade In], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm111.TB55 = vbNullString Or (Frm111.TB55 <> vbNullString And Not IsNumeric(Frm111.TB55)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar Buyback], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm111.TB56 = vbNullString Or (Frm111.TB56 <> vbNullString And Not IsNumeric(Frm111.TB56)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Caj Tukaran], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah anda ingin simpan tetapan ini ?" & vbCrLf & _
            "Sistem mungkin mengambil sedikit masa untuk menyimpan tetapan ini." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
    
        If MDI_frm1.L20_Text = "Semua cawangan" Then
            
            LM_CAWANGAN = "HQ"
            
        Else
            
            LM_CAWANGAN = MDI_frm1.L20_Text
            
        End If
        
        LM_NOW = Now
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from default_setting where Default1='" & LM_CAWANGAN & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            'If rs!Default1 = "Default" Then
                If Frm111.TB16 <> vbNullString Then 'Spread Trade In
                    rs!spread_Cash_Trade_In = Frm111.TB16
                Else
                    rs!spread_Cash_Trade_In = "0.00"
                End If
                'If Frm111.TB17 <> vbNullString Then 'Cas Kad Kredit
                '    rs!cas_Kad_Kredit = Format(Frm111.TB17, "0.00")
                'Else
                '    rs!cas_Kad_Kredit = Format(0, "0.00")
                'End If
                'If Frm111.TB27 <> vbNullString Then 'Cas Kad Debit
                '    rs!cas_debit_kad = Format(Frm111.TB27, "0.00")
                'Else
                '    rs!cas_debit_kad = Format(0, "0.00")
                'End If
                rs!cara_Belian = "Tunai" 'Cara Belian Frm111.CBB1
                If Frm111.TB20 <> vbNullString Then 'Kadar Komisen Barang Kemas (%)
                    rs!komisen = Format(Frm111.TB20, "0.00")
                Else
                    rs!komisen = "0.00"
                End If
                If Frm111.TB52 <> vbNullString Then 'Kadar Komisen Pekerja Per Gram
                    rs!komisen_per_gram = Format(Frm111.TB52, "0.00")
                Else
                    rs!komisen_per_gram = "0.00"
                End If
                If Frm111.TB33 <> vbNullString Then 'Harga emas per gram 999.9 (GDN)
                    rs!harga_999 = Format(Frm111.TB33, "0.00")
                Else
                    rs!harga_999 = "0.00"
                End If
                If Frm111.TB51 <> vbNullString Then 'Harga emas per gram 999.9 (GRN)
                    rs!harga_beli_999 = Format(Frm111.TB51, "0.00")
                Else
                    rs!harga_beli_999 = "0.00"
                End If
                If Frm111.TB34 <> vbNullString Then 'Kadar komisyen upah kepada agen dropship (%)
                    rs!kadar_komisyen_upah = Frm111.TB34
                Else
                    rs!kadar_komisyen_upah = 0
                End If
                If Frm111.TB29 <> vbNullString Then 'Kadar Komisen Barang Permata (%)
                    rs!komisen_permata = Format(Frm111.TB29, "0.00")
                Else
                    rs!komisen_permata = "0.00"
                End If
                If Frm111.TB24 <> vbNullString Then
                    rs!gst_value = Frm111.TB24 'Kadar GST
                Else
                    rs!gst_value = Null
                End If
                If Frm111.CB3 = 1 Then
                    rs!gst_arinashi = 1
                    rs!gst_value = Frm111.TB24 'Kadar GST
                ElseIf Frm111.CB7 = 1 Then
                    rs!gst_value = Frm111.TB24 'Kadar GST
                    rs!gst_arinashi = 2
                Else
                    rs!gst_arinashi = 0
                End If
                If Frm111.CB5 = 1 Then
                    rs!gst_arinashi_belian = 1
                    rs!gst_value = Frm111.TB24 'Kadar GST
                Else
                    rs!gst_arinashi_belian = 0
                End If
                If Frm111.CB1 = 1 Then
                    rs!ScannerMode = 1
                Else
                    rs!ScannerMode = 0
                End If
                If Frm111.CB2 = 1 Then
                    rs!BarcodeYesNo = 1
                Else
                    rs!BarcodeYesNo = 0
                End If
                If Frm111.CB15 = 1 Then
                    rs!printer_mode_ti = 1
                Else
                    rs!printer_mode_ti = 0
                End If
                If Frm111.CB4 = 1 Then
                    rs!flag_upah = 1
                Else
                    rs!flag_upah = 0
                End If
                If Frm111.CB6 = 1 Then
                    rs!diskaun_ari_nashi = 1
                    rs!diskaun = Frm111.TB32 'Kadar Diskaun
                Else
                    rs!diskaun_ari_nashi = 0
                    rs!diskaun = 0 'Kadar Diskaun
                End If
                If Frm111.TB25 <> vbNullString Then 'Had Kadar Penurunan Harga Jualan Per Gram
                    rs!limit_per_gram = Format(Frm111.TB25, "0.00")
                Else
                    rs!limit_per_gram = Format(0, "0.00")
                End If
                If Frm111.TB26 <> vbNullString Then 'Had Kadar Penurunan Harga Jualan Per Gram
                    rs!limit_per_item = Format(Frm111.TB26, "0.00")
                Else
                    rs!limit_per_item = Format(0, "0.00")
                End If
                If Frm111.TB28 <> vbNullString Then
                    rs!potongan_trade_in = Frm111.TB28 'Kadar Potongan Trade In - Jika Kedai Perlu Bayar (%)
                Else
                    rs!potongan_trade_in = Null 'Kadar Potongan Trade In - Jika Kedai Perlu Bayar (%)
                End If
                If Frm111.CB7 = 1 Then
                    rs!gst_jualan_included = 1
                Else
                    rs!gst_jualan_included = 0
                End If
                If Frm111.CB8 = 1 Then '0 : Pengiraan mengikut berat barang , 1 : Pengiraan mengikut tetapan asal
                    rs!kiraan_upah = 0
                ElseIf Frm111.CB9 = 1 Then
                    rs!kiraan_upah = 1
                End If
                If Frm111.CB10 = 1 Then 'Jenis header bagi invoice , 0 : Pre-printed , 1 : Header dari sistem
                    rs!jenis_header = 0
                ElseIf Frm111.CB11 = 1 Then
                    rs!jenis_header = 1
                End If
                If Frm111.TB35 <> vbNullString Then 'Kadar diskaun per gram bagi penggunaan kupon
                    rs!kupon_diskaun = Format(Frm111.TB35, "0.00")
                Else
                    rs!kupon_diskaun = Format(0, "0.00")
                End If
                If Frm111.TB40 <> vbNullString Then 'Tael
                    rs!tael = Frm111.TB40
                Else
                    rs!tael = 0
                End If
                If Frm111.TB41 <> vbNullString Then 'Public
                    rs!public = Frm111.TB41
                Else
                    rs!public = 0
                End If
                If Frm111.TB42 <> vbNullString Then 'SA
                    rs!sa = Frm111.TB42
                Else
                    rs!sa = 0
                End If
                If Frm111.CB12 = 1 Then 'jenis penetapan upah dari supplier , 0 : Upah ikut berat , 1 : Upah ikut harga tetap
                    rs!upah_supplier = 0
                ElseIf Frm111.CB13 = 1 Then
                    rs!upah_supplier = 1
                End If
                If Frm111.TB53 <> vbNullString Then 'Top Margin
                    rs!top_margin = Frm111.TB53
                Else
                    rs!top_margin = 0
                End If
                If Frm111.CB14 = 1 Then
                    rs!invoice_tak_rasmi = 0 '0 : Tidak Rasmi , 1 : Rasmi
                Else
                    rs!invoice_tak_rasmi = 1
                End If
                If Frm111.TB54 <> vbNullString Then
                    rs!rate_trade_in = Frm111.TB54
                Else
                    rs!rate_trade_in = 0
                End If
                If Frm111.TB55 <> vbNullString Then
                    rs!rate_buyback = Frm111.TB55
                Else
                    rs!rate_buyback = 0
                End If
                If Frm111.TB56 <> vbNullString Then
                    rs!rate_caj_pertukaran = Frm111.TB56
                Else
                    rs!rate_caj_pertukaran = 0
                End If
                rs!write_timestamp = LM_NOW

                rs.Update
            'End If
        End If
        
        rs.Close
        Set rs = Nothing
        
        Call main_setting
        
        user = MDI_frm1.L3_Text
        LogAct_Memory = "[" & user & "] Tetapan sistem."
        LogDate_Memory = LM_NOW
        Call UpdateLog_Database
        
        Call Frm111_setting
        
        MsgBox "Tetapan telah BERJAYA disimpan.", vbInformation, "Info"
    End If
End If
End Sub
Private Sub CMD2_Click()
'on error resume next
Dim Err(5)
DATA_SAVE = 0

If Frm111.TB36 = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan [Jenis kad]."
End If
If Frm111.TB37 = vbNullString Or (Frm111.TB37 <> vbNullString And Not IsNumeric(Frm111.TB37)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Caj Perkhidmatan (%)]. Hanya NOMBOR dibenarkan di dalam ruangan ini."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah anda simpan data jenis kad ini?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 74_cas_kad_kredit where jenis_kad='" & UCase(Frm111.TB36) & "' AND status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
            rs.AddNew
            If Frm111.TB36 <> vbNullString Then 'Jenis kad
                rs!jenis_kad = UCase(Frm111.TB36)
            Else
                rs!jenis_kad = Null
            End If
            If Frm111.TB37 <> vbNullString Then 'Caj perkhidmatan
                rs!cas_kad = Format(Frm111.TB37, "0.00")
            Else
                rs!cas_kad = "0.00"
            End If
            rs!Status = 1
            rs!write_timestamp = Now
            rs.Update
            
            DATA_SAVE = 1
        Else
            
            MsgBox "Jenis kad [" & UCase(Frm111.TB36) & "] telah didaftarkan sebelum ini. Sila periksa data anda.", vbInformation, "Info"
            
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_SAVE = 1 Then
        
'#### Update Log Aktiviti Sistem #### - Start
            user = MDI_frm1.L3_Text
            
            LogAct_Memory = "[" & user & "] Pendaftaran jenis kad [" & UCase(Frm111.TB36) & "]."
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End
            
            Call Frm111_initial_setting2
            
            Frm111.L9_Text = -1 'Senarai jenis kad : Paparan page
            Frm111.L11_Text = 0 'Senarai jenis kad : Titik carian data (default = -1)
            Frm111.L10_Text = 0 'Senarai jenis kad : Jumlah page
            GM_NEXT_PREV = 0
            
            Call Frm111_senarai_jenis_kad_header
            Call Frm111_senarai_jenis_kad
            
            MsgBox "Data telah berjaya disimpan.", vbInformation, "Info"
            
            Frm111.TB36.SetFocus
        
        End If
            
    End If
    
End If
End Sub
Private Sub CMD3_Click()
'on error resume next
Dim Err(5)
DATA_SAVE = 0

If Frm111.TB36 = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan [Jenis kad]."
End If
If Frm111.TB37 = vbNullString Or (Frm111.TB37 <> vbNullString And Not IsNumeric(Frm111.TB37)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Caj Perkhidmatan (%)]. Hanya NOMBOR dibenarkan di dalam ruangan ini."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah anda simpan data jenis kad ini?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 74_cas_kad_kredit where jenis_kad='" & UCase(Frm111.TB36) & "' AND status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If rs!ID <> Frm111.L14_Text Then
            
                MsgBox "Jenis kad [" & UCase(Frm111.TB36) & "] telah didaftarkan sebelum ini. Sila periksa data anda.", vbInformation, "Info"
                
                rs.Close
                Set rs = Nothing
                
                Exit Sub
                
            End If
        
        End If
        
        rs.Close
        Set rs = Nothing
    
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 74_cas_kad_kredit where ID='" & Frm111.L14_Text & "' AND status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            If Not IsNull(rs!jenis_kad) Then Frm111_LM_JENIS = rs!jenis_kad
            If Frm111.TB36 <> vbNullString Then 'Jenis kad
                rs!jenis_kad = UCase(Frm111.TB36)
            Else
                rs!jenis_kad = Null
            End If
            If Frm111.TB37 <> vbNullString Then 'Caj perkhidmatan
                rs!cas_kad = Format(Frm111.TB37, "0.00")
            Else
                rs!cas_kad = "0.00"
            End If
            rs!Status = 1
            rs!write_timestamp2 = Now
            rs.Update
            
            DATA_SAVE = 1

        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_SAVE = 1 Then
        
'#### Update Log Aktiviti Sistem #### - Start
            user = MDI_frm1.L3_Text
            
            LogAct_Memory = "[" & user & "] Edit jenis kad [" & Frm111_LM_JENIS & " -> " & UCase(Frm111.TB36) & "]."
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End
            
            Call Frm111_initial_setting2
            
            Frm111.L9_Text = -1 'Senarai jenis kad : Paparan page
            Frm111.L11_Text = 0 'Senarai jenis kad : Titik carian data (default = -1)
            Frm111.L10_Text = 0 'Senarai jenis kad : Jumlah page
            GM_NEXT_PREV = 0
            
            Call Frm111_senarai_jenis_kad_header
            Call Frm111_senarai_jenis_kad
            
            MsgBox "Data telah berjaya disimpan.", vbInformation, "Info"
            
            Frm111.TB36.SetFocus
        
        End If
            
    End If
    
End If
End Sub
Private Sub CMD4_Click()
'on error resume next
Note = "Batal edit data ini?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then

    Frm111.TB36 = vbNullString
    Frm111.TB37 = "0.00"
    
    
    Frm111.TB36.SetFocus

    Frm111.CMD2.Visible = True
    Frm111.CMD3.Visible = False
    Frm111.CMD4.Visible = False
    
    Frm111.L13_Text.Visible = False
    
End If
End Sub

Private Sub CMD5_Click()
'on error resume next
Dim Frm111_LM_CURR_PAGE As Double
Dim Frm111_LM_TOTAL_PAGE As Double

Frm111_LM_CURR_PAGE = 0
Frm111_LM_TOTAL_PAGE = 0

If Frm111.L11_Text <> vbNullString And IsNumeric(Frm111.L11_Text) Then
    If Frm111.L12_Text <> vbNullString And IsNumeric(Frm111.L12_Text) Then
        Frm111_LM_CURR_PAGE = Frm111.L11_Text
        Frm111_LM_TOTAL_PAGE = Frm111.L12_Text
        
        If Frm111_LM_CURR_PAGE <> 1 And Frm111_LM_CURR_PAGE <> 0 Then
        
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            Call Frm111_senarai_jenis_kad_header
            Call Frm111_senarai_jenis_kad
            
        End If
    End If
End If
End Sub

Private Sub CMD6_Click()
'on error resume next
Dim Frm111_LM_CURR_PAGE As Double
Dim Frm111_LM_TOTAL_PAGE As Double

Frm111_LM_CURR_PAGE = 0
Frm111_LM_TOTAL_PAGE = 0

If Frm111.L11_Text <> vbNullString And IsNumeric(Frm111.L11_Text) Then
    If Frm111.L12_Text <> vbNullString And IsNumeric(Frm111.L12_Text) Then
        Frm111_LM_CURR_PAGE = Frm111.L11_Text
        Frm111_LM_TOTAL_PAGE = Frm111.L12_Text
        
        If Frm111_LM_CURR_PAGE < Frm111_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm111_senarai_jenis_kad_header
            Call Frm111_senarai_jenis_kad
            
        End If
    End If
End If
End Sub
Private Sub CMD7_Click()
'On Error Resume Next
Dim Err(10)

x = 0
If Frm111.TB43 = vbNullString Or (Frm111.TB43 <> vbNullString And Not IsNumeric(Frm111.TB43)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar PEMBAHAGIAN mata ganjaran bagi ahli biasa], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm111.TB44 = vbNullString Or (Frm111.TB44 <> vbNullString And Not IsNumeric(Frm111.TB44)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar TEBUSAN mata ganjaran bagi ahli biasa], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm111.TB45 = vbNullString Or (Frm111.TB45 <> vbNullString And Not IsNumeric(Frm111.TB45)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar PEMBAHAGIAN mata ganjaran bagi silver], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm111.TB46 = vbNullString Or (Frm111.TB46 <> vbNullString And Not IsNumeric(Frm111.TB46)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar TEBUSAN mata ganjaran bagi silver], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm111.TB47 = vbNullString Or (Frm111.TB47 <> vbNullString And Not IsNumeric(Frm111.TB47)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar PEMBAHAGIAN mata ganjaran bagi gold], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm111.TB48 = vbNullString Or (Frm111.TB48 <> vbNullString And Not IsNumeric(Frm111.TB48)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar TEBUSAN mata ganjaran bagi gold], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm111.TB49 = vbNullString Or (Frm111.TB49 <> vbNullString And Not IsNumeric(Frm111.TB49)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar PEMBAHAGIAN mata ganjaran bagi platinum], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm111.TB50 = vbNullString Or (Frm111.TB50 <> vbNullString And Not IsNumeric(Frm111.TB50)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar TEBUSAN mata ganjaran bagi platinum], Hanya NOMBOR dibenarkan dalam ruangan ini."
End If

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
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
        
        LM_NOW = Now
        
        If MDI_frm1.L20_Text = "Semua cawangan" Then
            
            LM_CAWANGAN = "HQ"
            
        Else
            
            LM_CAWANGAN = MDI_frm1.L20_Text
            
        End If

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from default_setting where Default1='" & LM_CAWANGAN & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then

            If Frm111.TB43 <> vbNullString Then 'Kadar perolehan mata ganjaran (ahli biasa)
                rs!pemalar_bonus_biasa = Frm111.TB43
                G_PEMALAR_BONUS_BIASA = Frm111.TB43
            Else
                rs!pemalar_bonus_biasa = 0
            End If
            If Frm111.TB44 <> vbNullString Then 'Kadar tebusan mata ganjaran (ahli biasa)
                rs!pemalar_tebus_bonus_biasa = Frm111.TB44
                G_PEMALAR_TEBUS_BIASA = Frm111.TB44
            Else
                rs!pemalar_tebus_bonus_biasa = 0
            End If
            If Frm111.TB45 <> vbNullString Then 'Kadar perolehan mata ganjaran (silver)
                rs!pemalar_bonus_silver = Frm111.TB45
                G_PEMALAR_BONUS_SILVER = Frm111.TB45
            Else
                rs!pemalar_bonus_silver = 0
            End If
            If Frm111.TB46 <> vbNullString Then 'Kadar tebusan mata ganjaran (silver)
                rs!pemalar_tebus_bonus_silver = Frm111.TB46
                G_PEMALAR_TEBUS_SILVER = Frm111.TB46
            Else
                rs!pemalar_tebus_bonus_silver = 0
            End If
            If Frm111.TB47 <> vbNullString Then 'Kadar perolehan mata ganjaran (gold)
                rs!pemalar_bonus_gold = Frm111.TB47
                G_PEMALAR_BONUS_GOLD = Frm111.TB47
            Else
                rs!pemalar_bonus_gold = 0
            End If
            If Frm111.TB48 <> vbNullString Then 'Kadar tebusan mata ganjaran (gold)
                rs!pemalar_tebus_bonus_gold = Frm111.TB48
                G_PEMALAR_TEBUS_GOLD = Frm111.TB48
            Else
                rs!pemalar_tebus_bonus_gold = 0
            End If
            
            If Frm111.TB49 <> vbNullString Then 'Kadar perolehan mata ganjaran (platinum)
                rs!pemalar_bonus_platinum = Frm111.TB49
                G_PEMALAR_BONUS_PLATINUM = Frm111.TB49
            Else
                rs!pemalar_bonus_platinum = 0
            End If
            If Frm111.TB50 <> vbNullString Then 'Kadar tebusan mata ganjaran (platinum)
                rs!pemalar_tebus_bonus_platinum = Frm111.TB50
                G_PEMALAR_TEBUS_PLATINUM = Frm111.TB50
            Else
                rs!pemalar_tebus_bonus_platinum = 0
            End If

            rs.Update

        End If
        
        rs.Close
        Set rs = Nothing
        
        user = MDI_frm1.L3_Text
        LogAct_Memory = "[" & user & "] Tetapan sistem (Tetapan kadar perolehan / tebus mata ganjaran)."
        LogDate_Memory = LM_NOW
        Call UpdateLog_Database
        
        Call Frm111_setting
        
        MsgBox "Tetapan telah BERJAYA disimpan.", vbInformation, "Info"
    End If
End If
End Sub
Private Sub Form_Load()
'on error resume next
Frm111.L9_Text = 0 'Senarai jenis kad : Paparan page
Frm111.L10_Text = 0 'Senarai jenis kad : Jumlah page
Frm111.L11_Text = 0 'Senarai jenis kad : Titik carian data (default = -1)
Frm111.L12_Text = 0 'Senarai jenis kad : Flag page terakhir
End Sub

Private Sub Frm111_SM_edit_data_Click()
'On Error Resume Next
Frm111_LM_ID = vbNullString
DATA_FOUND = 0

If Frm111.MSFlexGrid1 <> vbNullString Then
    'If Button = vbRightButton Then
    
        If IsNumeric(Frm111.MSFlexGrid1) Then
            Frm111_LM_ID = Frm111.MSFlexGrid1.TextMatrix(Frm111.MSFlexGrid1, 2) 'No. ID
            
            If Frm111_LM_ID <> vbNullString Then
                
                Call Frm111_initial_setting2
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 74_cas_kad_kredit where ID='" & Frm111_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                
                    Frm111.L14_Text = Frm111_LM_ID
                    If Not IsNull(rs!jenis_kad) Then Frm111.TB36 = rs!jenis_kad
                    If Not IsNull(rs!cas_kad) Then Frm111.TB37 = Format(rs!cas_kad, "0.00")
                    DATA_FOUND = 1
                    
                End If
                
                rs.Close
                Set rs = Nothing
                
                If DATA_FOUND = 1 Then

                    Frm111.CMD2.Visible = False
                    Frm111.CMD3.Visible = True
                    Frm111.CMD4.Visible = True
                
                End If

            End If
            
        End If
        
    'End If
End If
End Sub

Private Sub Frm111_SM_padam_data_Click()
'On Error Resume Next
Frm111_LM_ID = vbNullString
DATA_FOUND = 0

If Frm111.MSFlexGrid1 <> vbNullString Then
    'If Button = vbRightButton Then
    
        If IsNumeric(Frm111.MSFlexGrid1) Then
            Frm111_LM_ID = Frm111.MSFlexGrid1.TextMatrix(Frm111.MSFlexGrid1, 2) 'No. ID
            
            If Frm111_LM_ID <> vbNullString Then
                
                Note = "Adakah anda ingin padam data ini?"
                
                Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
                
                If Answer = vbYes Then
                
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from 74_cas_kad_kredit where ID='" & Frm111_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                        If Not IsNull(rs!jenis_kad) Then Frm111_LM_JENIS = rs!jenis_kad
                        rs!Status = 0
                        rs!write_timestamp3 = Now
                        DATA_FOUND = 1
                        rs.Update
                    End If
                    
                    rs.Close
                    Set rs = Nothing
                    
                    If DATA_FOUND = 1 Then
    
'#### Update Log Aktiviti Sistem #### - Start
                        user = MDI_frm1.L3_Text
                        
                        LogAct_Memory = "[" & user & "] Padam data kad [" & Frm111_LM_JENIS & "]."
                        LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                        Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End
                        
                        Frm111.L9_Text = -1 'Senarai jenis kad : Paparan page
                        Frm111.L11_Text = 0 'Senarai jenis kad : Titik carian data (default = -1)
                        Frm111.L10_Text = 0 'Senarai jenis kad : Jumlah page
                        GM_NEXT_PREV = 0
                        
                        Call Frm111_senarai_jenis_kad_header
                        Call Frm111_senarai_jenis_kad
                        
                        MsgBox "Data telah berjaya dipadamkan.", vbInformation, "Info"
                        
                        Frm111.TB36.SetFocus
            
                    End If
                    
                End If

            End If
            
        End If
        
    'End If
End If
End Sub

Private Sub L2_Text_Click()
'on error resume next
If Frm111.Frame15.Visible = False Then

    Call Frm111_initial_setting
    Call Frm111_initial_setting2
    
    Frm111.L9_Text = -1 'Senarai jenis kad : Paparan page
    Frm111.L11_Text = 0 'Senarai jenis kad : Titik carian data (default = -1)
    Frm111.L10_Text = 0 'Senarai jenis kad : Jumlah page
    GM_NEXT_PREV = 0
    
    Call Frm111_senarai_jenis_kad_header
    Call Frm111_senarai_jenis_kad
    
    Frm111.Frame15.Visible = True
    
    Frm111.TB36.SetFocus
    
Else

    Frm111.Frame15.Visible = False
    
End If
End Sub
Private Sub L3_Text_Click()
'on error resume next
If Frm111.Frame1.Visible = False Then

    If MDI_frm1.L20_Text = "Semua cawangan" Then
    
        Frm96.CMD2.Visible = True
        Frm96.CMD1.Visible = False
    
        Call Frm96_initial
        
        Frm96.Show vbModal
        
    End If
    
    If G_GST_SYSTEM = "YES" Then
        Frm111.Frame16.Visible = True
    Else
        Frm111.Frame16.Visible = False
    End If
    
    Call Frm111_initial_setting
    Call Frm111_setting
    
    Frm111.Frame1.Visible = True
Else
    Frm111.Frame1.Visible = False
End If
End Sub
Private Sub L4_Text_Click()
'on error resume next
If Frm111.Frame14.Visible = False Then

    If MDI_frm1.L20_Text = "Semua cawangan" Then
    
        Frm96.CMD2.Visible = True
        Frm96.CMD1.Visible = False
    
        Call Frm96_initial
        
        Frm96.Show vbModal
        
    End If
    
    Call Frm111_initial_setting
    Call Frm111_setting2
    
    Frm111.Frame14.Visible = True
Else
    Frm111.Frame14.Visible = False
End If
End Sub
Private Sub MSFlexGrid1_DblClick()
'On Error Resume Next
Frm111_LM_ID = vbNullString

If Frm111.MSFlexGrid1 <> vbNullString Then
    'If Button = vbRightButton Then
    
        If IsNumeric(Frm111.MSFlexGrid1) Then
            Frm111_LM_ID = Frm111.MSFlexGrid1.TextMatrix(Frm111.MSFlexGrid1, 2) 'No. ID
            Frm111_LM_JENIS = Frm111.MSFlexGrid1.TextMatrix(Frm111.MSFlexGrid1, 3) 'Jenis Kad
            
            If Frm111_LM_ID <> vbNullString Then
                
                Frm111.Frm111_SM_edit_data.Caption = "Edit data. (" & Frm111_LM_JENIS & ")"
                Frm111.Frm111_SM_padam_data.Caption = "Padam data. (" & Frm111_LM_JENIS & ")"
                
                PopupMenu Frm111_PM_menu1, vbPopupMenuRightButton

            End If
            
        End If
        
    'End If
End If
End Sub

Private Sub Tmr1_Timer()
'On Error Resume Next
If Frm111.CMD3.Visible = True Then
    If Frm111.L13_Text.Visible = True Then
        Frm111.L13_Text.Visible = False
    Else
        Frm111.L13_Text.Visible = True
    End If
End If
End Sub
