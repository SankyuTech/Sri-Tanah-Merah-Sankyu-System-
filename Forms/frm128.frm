VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm128 
   Caption         =   "Maklumat deposit dan perbelanjaan pelanggan."
   ClientHeight    =   13035
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   23880
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
   ScaleWidth      =   23880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rekod"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9375
      Left            =   960
      TabIndex        =   41
      Top             =   2040
      Visible         =   0   'False
      Width           =   20535
      Begin VB.CommandButton CMD22 
         Caption         =   "Next"
         Height          =   810
         Left            =   19200
         MouseIcon       =   "frm128.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frm128.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Tutup senarai ini."
         Top             =   8400
         Width           =   1095
      End
      Begin VB.CommandButton CMD21 
         Caption         =   "Back"
         Height          =   810
         Left            =   18000
         MouseIcon       =   "frm128.frx":13D4
         MousePointer    =   99  'Custom
         Picture         =   "frm128.frx":16DE
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Tutup senarai ini."
         Top             =   8400
         Width           =   1095
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   7860
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   20235
         _ExtentX        =   35692
         _ExtentY        =   13864
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
      Begin VB.Label L14_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai rekod simpanan , pulangan dan penggunaan duit pelanggan ini.                      "
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
         TabIndex        =   55
         Top             =   240
         Width           =   15855
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
         Left            =   17520
         TabIndex        =   53
         Top             =   8400
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
         Left            =   16920
         TabIndex        =   52
         Top             =   8400
         Width           =   375
      End
      Begin VB.Label L69_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L69_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   14640
         TabIndex        =   51
         Top             =   8640
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L70_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L70_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   13680
         TabIndex        =   50
         Top             =   8640
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   $"frm128.frx":27A8
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
         TabIndex        =   49
         Top             =   8400
         Width           =   9615
      End
      Begin VB.Label L27_Text 
         BackColor       =   &H8000000C&
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
         Left            =   4200
         TabIndex        =   48
         Top             =   8400
         Width           =   1695
      End
      Begin VB.Label L26_Text 
         BackColor       =   &H8000000C&
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
         Left            =   1560
         TabIndex        =   47
         Top             =   8400
         Width           =   1335
      End
      Begin VB.Label L28_Text 
         BackColor       =   &H8000000C&
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
         Left            =   7350
         TabIndex        =   46
         Top             =   8400
         Width           =   1695
      End
      Begin VB.Label L29_Text 
         BackColor       =   &H8000000C&
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
         Left            =   9600
         TabIndex        =   45
         Top             =   8400
         Width           =   1695
      End
      Begin VB.Label Label12 
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
         Left            =   15600
         TabIndex        =   54
         Top             =   8400
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Simpanan Duit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   5400
      TabIndex        =   10
      Top             =   5880
      Visible         =   0   'False
      Width           =   7575
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
         Left            =   2400
         MouseIcon       =   "frm128.frx":2845
         MousePointer    =   99  'Custom
         Picture         =   "frm128.frx":2B4F
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3960
         Width           =   2775
      End
      Begin VB.CheckBox CB1 
         BackColor       =   &H8000000C&
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
         Left            =   2880
         TabIndex        =   21
         Top             =   3435
         Width           =   200
      End
      Begin VB.CheckBox CB2 
         BackColor       =   &H8000000C&
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
         Left            =   2880
         TabIndex        =   20
         Top             =   3675
         Width           =   200
      End
      Begin VB.ComboBox CBB1 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1935
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1560
         Width           =   5205
      End
      Begin VB.TextBox TB1 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1935
         TabIndex        =   12
         Text            =   "TB1"
         Top             =   840
         Width           =   5205
      End
      Begin VB.TextBox TB3 
         BackColor       =   &H00FFFFFF&
         Height          =   1320
         Left            =   1935
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "frm128.frx":5119
         Top             =   1920
         Width           =   5205
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   1935
         TabIndex        =   14
         Top             =   1200
         Width           =   5205
         _ExtentX        =   9181
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
         Format          =   414842880
         CurrentDate     =   41561
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bayaran diterima secara * :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   -120
         TabIndex        =   23
         Top             =   3360
         Width           =   2835
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Tunai                    Bank In"
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   3120
         TabIndex        =   22
         Top             =   3390
         Width           =   1995
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pekerja * :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   -120
         TabIndex        =   19
         Top             =   1560
         Width           =   1995
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh * :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   -120
         TabIndex        =   18
         Top             =   1200
         Width           =   1995
      End
      Begin VB.Label Label70 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah (RM) :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   -120
         TabIndex        =   17
         Top             =   840
         Width           =   1995
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Simpanan duit di kedai."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   480
         TabIndex        =   16
         Top             =   360
         Width           =   7305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tujuan * :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   -120
         TabIndex        =   15
         Top             =   1920
         Width           =   1995
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pulangan Duit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   1680
      TabIndex        =   25
      Top             =   2520
      Visible         =   0   'False
      Width           =   7575
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
         Left            =   2280
         MouseIcon       =   "frm128.frx":511D
         MousePointer    =   99  'Custom
         Picture         =   "frm128.frx":5427
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   4200
         Width           =   2775
      End
      Begin VB.CheckBox CB3 
         BackColor       =   &H8000000C&
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
         Left            =   2760
         TabIndex        =   37
         Top             =   3435
         Width           =   200
      End
      Begin VB.CheckBox CB4 
         BackColor       =   &H8000000C&
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
         Left            =   2760
         TabIndex        =   36
         Top             =   3675
         Width           =   200
      End
      Begin VB.CheckBox CB5 
         BackColor       =   &H8000000C&
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
         Left            =   2760
         TabIndex        =   35
         Top             =   3915
         Width           =   200
      End
      Begin VB.TextBox TB2 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1815
         TabIndex        =   28
         Text            =   "TB2"
         Top             =   840
         Width           =   5205
      End
      Begin VB.ComboBox CBB2 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1815
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   1560
         Width           =   5205
      End
      Begin VB.TextBox TB4 
         BackColor       =   &H00FFFFFF&
         Height          =   1320
         Left            =   1815
         MultiLine       =   -1  'True
         TabIndex        =   26
         Text            =   "frm128.frx":79F1
         Top             =   1920
         Width           =   5205
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   360
         Left            =   1815
         TabIndex        =   29
         Top             =   1200
         Width           =   5205
         _ExtentX        =   9181
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
         Format          =   414842880
         CurrentDate     =   41561
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bayaran dibuat secara * :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   -240
         TabIndex        =   39
         Top             =   3360
         Width           =   2835
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Tunai                    Bank In                   Cek"
         ForeColor       =   &H00000000&
         Height          =   885
         Left            =   3000
         TabIndex        =   38
         Top             =   3390
         Width           =   1995
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Pulangan Duit."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         TabIndex        =   34
         Top             =   360
         Width           =   7305
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah (RM) :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   -240
         TabIndex        =   33
         Top             =   870
         Width           =   1995
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh * :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   -240
         TabIndex        =   32
         Top             =   1200
         Width           =   1995
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pekerja * :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   -240
         TabIndex        =   31
         Top             =   1560
         Width           =   1995
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sebab * :"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   -240
         TabIndex        =   30
         Top             =   1920
         Width           =   1995
      End
   End
   Begin VB.Label L8_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kembali"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      MouseIcon       =   "frm128.frx":79F5
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label L7_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rekod Simpanan / Pulangan / Penggunaan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      MouseIcon       =   "frm128.frx":7CFF
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   1560
      Width           =   5415
   End
   Begin VB.Label L6_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pulangan Duit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      MouseIcon       =   "frm128.frx":8009
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label L5_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   -120
      MouseIcon       =   "frm128.frx":8313
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "Maklumat pelanggan."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6945
   End
   Begin VB.Label L1_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L1_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2760
      TabIndex        =   4
      Top             =   360
      Width           =   8625
   End
   Begin VB.Label Label47 
      BackStyle       =   0  'Transparent
      Caption         =   $"frm128.frx":861D
      ForeColor       =   &H00000000&
      Height          =   1005
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   2625
   End
   Begin VB.Label L2_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L2_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Top             =   600
      Width           =   8625
   End
   Begin VB.Label L3_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L3_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   840
      Width           =   8625
   End
   Begin VB.Label L4_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L4_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2760
      TabIndex        =   0
      Top             =   1080
      Width           =   8625
   End
   Begin VB.Menu frm128_pm_menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu frm128_sm_cetak 
         Caption         =   "Cetak Resit / Payment Voucher"
      End
   End
End
Attribute VB_Name = "frm128"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB1_Click()
'On Error Resume Next
If frm128.CB1 = 1 Then
    frm128.CB2 = 0
End If
End Sub
Private Sub CB2_Click()
'On Error Resume Next
If frm128.CB2 = 1 Then
    frm128.CB1 = 0
End If
End Sub
Private Sub CB3_Click()
'On Error Resume Next
If frm128.CB3 = 1 Then
    frm128.CB4 = 0
    frm128.CB5 = 0
End If
End Sub
Private Sub CB4_Click()
'On Error Resume Next
If frm128.CB4 = 1 Then
    frm128.CB3 = 0
    frm128.CB5 = 0
End If
End Sub
Private Sub CB5_Click()
'On Error Resume Next
If frm128.CB5 = 1 Then
    frm128.CB3 = 0
    frm128.CB4 = 0
End If
End Sub
Private Sub CMD1_Click()
'On Error Resume Next
Dim Err(6)
Dim Frm128_LM_SIMPANAN_ASAL As Double
Dim Frm128_LM_SIMPANAN_BARU As Double

DATA_SAVE = 0
Frm128_LM_SIMPANAN_ASAL = 0
Frm128_LM_SIMPANAN_BARU = 0

If frm128.L4_Text = vbNullString Then
    x = x + 1
    Err(x) = "Ralat telah berlaku. Sila keluar dari menu ini dan cuba lagi."
End If
If frm128.TB1 = vbNullString Or (frm128.TB1 <> vbNullString And Not IsNumeric(frm128.TB1)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Jumlah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If frm128.TB1 <> vbNullString Then
    If IsNumeric(frm128.TB1) Then
        
        If Format(frm128.TB1, "0.00") = "0.00" Then
            x = x + 1
            Err(x) = "Nilai 0 bagi jumlah tidak dibenarkan. Hanya NOMBOR dibenarkan dalam ruangan ini."
        End If
        
    End If
End If
If frm128.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Nama Pekerja]."
End If
If frm128.TB3 = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan tujuan bayaran."
End If
If frm128.CB1 = 0 And frm128.CB2 = 0 Then
    x = x + 1
    Err(x) = "Sila pilih cara bayaran diterima."
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
        
        LM_NOW = Now
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        rs.Open "select * from 1_senarai_invoice_deposit", cn2, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        rs!tarikh = frm128.DTPicker1
        rs!terminal = G_TERMINAL
        rs!write_timestamp = LM_NOW
        rs!Status = 1
        rs!nama_staff = MDI_frm1.L3_Text
        rs.Update
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        rs.Open "select * from 1_senarai_invoice_deposit where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & frm128.DTPicker1 & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!ID) Then
            
                rs!no_invoice = "REC" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(rs!ID, "000000")
                Frm128_LM_NO_REC = rs!ID
                
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
        
        '"REC" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm128_LM_NO_REC, "000000")
re_gen_no_rec:
        
'###Masukkan Data Belian Ke Dalam Database Akaun Belian### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 24_rekod_kewangan_pelanggan where no_resit='" & "REC" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm128_LM_NO_REC, "000000") & "' AND jenis='" & "0" & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
        
            rs.AddNew
            rs!tarikh = frm128.DTPicker1 'Tarikh
            rs!jenis = 0 '0 : Simpanan , 1 : Penggunaan Duit , 2 : Pulangan wang pelanggan
            rs!Status = 1
            If frm128.L4_Text <> vbNullString Then 'No. Rujukan Pelanggan
                rs!no_rujukan_pelanggan = frm128.L4_Text
            Else
                rs!no_rujukan_pelanggan = Null
            End If
            rs!no_resit = "REC" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm128_LM_NO_REC, "000000") 'No. Rujukan
            If frm128.TB1 <> vbNullString Then 'Jumlah Simpanan / Penggunaan (RM)
                rs!jumlah = Format(frm128.TB1, "0.00")
                Frm128_LM_SIMPANAN_BARU = frm128.TB1 'Simpanan Yang Baru (RM)
                
                If frm128.CB1 = 1 Then
                    rs!tunai = Format(frm128.TB1, "0.00")
                ElseIf frm128.CB2 = 1 Then
                    rs!bank_in = Format(frm128.TB1, "0.00")
                End If
            Else
                rs!jumlah = Null 'Jumlah Simpanan / Penggunaan (RM)
            End If
            If frm128.CBB1 <> vbNullString Then
                Frm128_LM_EMP_NO = Split(frm128.CBB1, "  |  ")(1)
                Frm128_LM_EMP_NAMA = Split(frm128.CBB1, "  |  ")(0)
                rs!no_rujukan_pekerja = Frm128_LM_EMP_NO 'No. Pekerja
            End If
            If frm128.TB3 <> vbNullString Then
                rs!jenis_penggunaan = frm128.TB3
            Else
                rs!jenis_penggunaan = Null
            End If
            rs!cawangan = G_CAWANGAN
            DATA_SAVE = 1
            rs.Update
            
        Else
        
            Frm128_LM_NO_REC = Frm128_LM_NO_REC + 1
            
            rs.Close
            Set rs = Nothing
            
            GoTo re_gen_no_rec:
            
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_SAVE = 1 Then

            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from senarai_pelanggan where no_pelanggan='" & frm128.L4_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
            If Not rs.EOF Then
            
                If Not IsNull(rs!baki_simpanan) Then
                    If IsNumeric(rs!baki_simpanan) Then Frm128_LM_SIMPANAN_ASAL = rs!baki_simpanan
                End If
                rs!baki_simpanan = Format(Frm128_LM_SIMPANAN_ASAL + Frm128_LM_SIMPANAN_BARU, "0.00") 'Baki Simpanan Terbaru (RM)
                rs.Update
                
            End If
            
            rs.Close
            Set rs = Nothing
            
            G_PREVIEW = 1
            G_No_RESIT_JUALAN = "REC" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm128_LM_NO_REC, "000000")
            G_KEDAI = G_CAWANGAN
        
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & Frm128_LM_EMP_NAMA & "] Simpanan duit di kedai. No rujukan [" & "REC" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm128_LM_NO_REC, "000000") & "]."
            LogDate_Memory = LM_NOW
            Call UpdateLog_Database
            
            Call frm128_reset_simpanan
            Call frm128_cetak_receipt
            
            MsgBox "Data Telah Berjaya Disimpan", vbInformation, "Info"
            
        End If
    End If
End If
End Sub
Private Sub CMD2_Click()
'On Error Resume Next
Dim Err(6)
Dim Frm128_LM_SIMPANAN_ASAL As Double
Dim Frm128_LM_REFUND As Double

DATA_SAVE = 0
Frm128_LM_SIMPANAN_ASAL = 0
Frm128_LM_REFUND = 0

If frm128.L4_Text = vbNullString Then
    x = x + 1
    Err(x) = "Ralat telah berlaku. Sila keluar dari menu ini dan cuba lagi."
End If
If frm128.TB2 = vbNullString Or (frm128.TB2 <> vbNullString And Not IsNumeric(frm128.TB2)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Jumlah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If frm128.TB2 <> vbNullString Then
    If IsNumeric(frm128.TB2) Then
        
        If Format(frm128.TB2, "0.00") = "0.00" Then
            x = x + 1
            Err(x) = "Nilai 0 bagi jumlah tidak dibenarkan. Hanya NOMBOR dibenarkan dalam ruangan ini."
        End If
        
    End If
End If
If frm128.CBB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Nama Pekerja]."
End If
If frm128.TB4 = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan sebab bayaran."
End If
If frm128.CB3 = 0 And frm128.CB4 = 0 And frm128.CB5 = 0 Then
    x = x + 1
    Err(x) = "Sila pilih cara bayaran dibuat."
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
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where no_pelanggan='" & frm128.L4_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
        If Not rs.EOF Then
        
            If Not IsNull(rs!baki_simpanan) Then
                If IsNumeric(rs!baki_simpanan) Then Frm128_LM_SIMPANAN_ASAL = rs!baki_simpanan
            End If
            
        End If
        
        rs.Close
        Set rs = Nothing
        
        Frm128_LM_REFUND = frm128.TB2 'Simpanan Yang Baru (RM)
        
        If Frm128_LM_REFUND > Frm128_LM_SIMPANAN_ASAL Then
            
            MsgBox "Jumlah pulangan adalah melebihi jumlah duit terkumpul." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Jumlah pulangan : RM " & Format(Frm128_LM_REFUND, "#,##0.00") & vbCrLf & _
                    "Jumlah duit terkumpul : RM " & Format(Frm128_LM_SIMPANAN_ASAL, "#,##0.00") & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila periksa data anda.", vbExclamation, "Info"
            
            Exit Sub
        End If
        
        LM_NOW = Now

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        rs.Open "select * from 14_senarai_voucher", cn2, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        rs!tarikh = frm128.DTPicker2
        rs!terminal = G_TERMINAL
        rs!write_timestamp = LM_NOW
        rs!Status = 1
        rs!nama_staff = MDI_frm1.L3_Text
        rs.Update
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        rs.Open "select * from 14_senarai_voucher where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & frm128.DTPicker2 & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!ID) Then
            
                rs!no_voucher = "PV" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(rs!ID, "000000")
                Frm128_LM_NO_VOUCHER = rs!ID
                
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
        
        '"PV" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm128_LM_NO_VOUCHER, "000000")
        
search_next_no:
        
'###Masukkan Data Belian Ke Dalam Database Akaun Belian### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 24_rekod_kewangan_pelanggan where no_resit='" & "PV" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm128_LM_NO_VOUCHER, "000000") & "' AND jenis='" & "2" & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
        
            rs.AddNew
            rs!tarikh = frm128.DTPicker2 'Tarikh
            rs!jenis = 2 '0 : Simpanan , 1 : Penggunaan Duit , 2 : Pulangan wang pelanggan
            rs!Status = 1
            If frm128.L4_Text <> vbNullString Then 'No. Rujukan Pelanggan
                rs!no_rujukan_pelanggan = frm128.L4_Text
            Else
                rs!no_rujukan_pelanggan = Null
            End If
            rs!no_resit = "PV" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm128_LM_NO_VOUCHER, "000000") 'No. Rujukan
            If frm128.TB2 <> vbNullString Then 'Jumlah Simpanan / Penggunaan (RM)
                rs!jumlah = Format(frm128.TB2, "0.00")
                Frm128_LM_REFUND = frm128.TB2 'Simpanan Yang Baru (RM)

                If frm128.CB3 = 1 Then
                    rs!tunai = Format(frm128.TB2, "0.00")
                ElseIf frm128.CB4 = 1 Then
                    rs!bank_in = Format(frm128.TB2, "0.00")
                ElseIf frm128.CB5 = 1 Then
                    rs!cek = Format(frm128.TB2, "0.00")
                End If
                
            Else
                rs!jumlah = Null 'Jumlah Simpanan / Penggunaan (RM)
            End If
            If frm128.CBB2 <> vbNullString Then
                Frm128_LM_EMP_NO = Split(frm128.CBB2, "  |  ")(1)
                Frm128_LM_EMP_NAMA = Split(frm128.CBB2, "  |  ")(0)
                rs!no_rujukan_pekerja = Frm128_LM_EMP_NO 'No. Pekerja
            End If
            If frm128.TB4 <> vbNullString Then
                rs!jenis_penggunaan = frm128.TB4
            Else
                rs!jenis_penggunaan = Null
            End If
            rs!cawangan = G_CAWANGAN
            DATA_SAVE = 1
            rs.Update
            
        Else
        
            Frm128_LM_NO_VOUCHER = Frm128_LM_NO_VOUCHER + 1
            
            rs.Close
            Set rs = Nothing
            
            GoTo search_next_no:
            
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_SAVE = 1 Then

            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from senarai_pelanggan where no_pelanggan='" & frm128.L4_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
            If Not rs.EOF Then
            
                If Not IsNull(rs!baki_simpanan) Then
                    If IsNumeric(rs!baki_simpanan) Then Frm128_LM_SIMPANAN_ASAL = rs!baki_simpanan
                End If
                rs!baki_simpanan = Format(Frm128_LM_SIMPANAN_ASAL - Frm128_LM_REFUND, "0.00") 'Baki Simpanan Terbaru (RM)
                rs.Update
                
            End If
            
            rs.Close
            Set rs = Nothing
            
            G_PREVIEW = 1
            G_No_RESIT_JUALAN = "PV" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm128_LM_NO_VOUCHER, "000000")
            G_KEDAI = G_CAWANGAN
            
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & Frm128_LM_EMP_NAMA & "] Pulangan duit pelanggan. No rujukan [" & "PV" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm128_LM_NO_VOUCHER, "000000") & "]."
            LogDate_Memory = LM_NOW
            Call UpdateLog_Database
            
            Call frm128_reset_pulangan
            Call frm128_cetak_pv
            
            MsgBox "Data Telah Berjaya Disimpan", vbInformation, "Info"
            
        End If
    End If
End If
End Sub

Private Sub CMD21_Click()
'on error resume next
Dim frm128_LM_CURR_PAGE As Double
Dim frm128_LM_TOTAL_PAGE As Double

frm128_LM_CURR_PAGE = 0
frm128_LM_TOTAL_PAGE = 0

If frm128.L67_Text <> vbNullString And IsNumeric(frm128.L67_Text) Then
    If frm128.L68_Text <> vbNullString And IsNumeric(frm128.L68_Text) Then
        frm128_LM_CURR_PAGE = frm128.L67_Text
        frm128_LM_TOTAL_PAGE = frm128.L68_Text
        
        If frm128_LM_CURR_PAGE <> 1 And frm128_LM_CURR_PAGE <> 0 Then
        
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
                    
            Call frm128_report_simpanan_header
            Call frm128_report_simpanan
                    
        End If

    End If
End If
End Sub
Private Sub CMD22_Click()
'on error resume next
Dim frm128_LM_CURR_PAGE As Double
Dim frm128_LM_TOTAL_PAGE As Double

frm128_LM_CURR_PAGE = 0
frm128_LM_TOTAL_PAGE = 0

If frm128.L67_Text <> vbNullString And IsNumeric(frm128.L67_Text) Then
    If frm128.L68_Text <> vbNullString And IsNumeric(frm128.L68_Text) Then
        frm128_LM_CURR_PAGE = frm128.L67_Text
        frm128_LM_TOTAL_PAGE = frm128.L68_Text
        
        If frm128_LM_CURR_PAGE < frm128_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call frm128_report_simpanan_header
            Call frm128_report_simpanan
            
        End If
    End If
End If
End Sub
Private Sub Form_Load()
'on error resume next
Call frm128_background_color
End Sub

Private Sub frm128_sm_cetak_Click()
'on error resume next
DATA_FOUND = 0
Frm128_LM_INVOICE_TYPE = 0 'Unlimited , 1 : Limited
LM_JENIS = 3

If IsNumeric(frm128.LV1.SelectedItem.Index) Then
    
    Frm128_LM_No_ID = frm128.LV1.ListItems(frm128.LV1.SelectedItem.Index)
    
    If Frm128_LM_No_ID <> vbNullString Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 24_rekod_kewangan_pelanggan where ID='" & Frm128_LM_No_ID & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            If Not IsNull(rs!no_resit) Then G_No_RESIT_JUALAN = rs!no_resit
            
            If Not IsNull(rs!jenis) Then '0 : Simpanan , 1 : Penggunaan Duit , 2 : Pulangan wang pelanggan
                
                If rs!jenis = 0 Then
                    LM_JENIS = 0
                ElseIf rs!jenis = 1 Then
                    LM_JENIS = 1
                ElseIf rs!jenis = 2 Then
                    LM_JENIS = 2
                End If
            
            End If
            
        End If
        
        rs.Close
        Set rs = Nothing
        
        G_PREVIEW = 1
        
        If LM_JENIS = 3 Then
            
            MsgBox "Tiada maklumat dijumpai.", vbInformation, "Info"
            
        ElseIf LM_JENIS = 1 Then
            
            MsgBox "Tiada resit atau payment voucher bagi data ini.", vbInformation, "Info"
        
        ElseIf LM_JENIS = 0 Then
        
            Call frm128_cetak_receipt
            
        ElseIf LM_JENIS = 2 Then
            
            Call frm128_cetak_pv
        
        End If
        
        
    Else
    
        MsgBox "Tiada Data.", vbInformation, "Info"
        
    End If
    
Else

    MsgBox "Tiada Data.", vbInformation, "Info"
    
End If
End Sub

Private Sub L5_Text_Click()
'on error resume next
If frm128.Frame1.Visible = False Then

    If MDI_frm1.L20_Text = "Semua cawangan" Then
    
        Frm96.CMD2.Visible = True
        Frm96.CMD1.Visible = False
    
        Call Frm96_initial
        
        Frm96.Show vbModal
        
    End If
    
    Call frm128_pic_ena_disable
    Call frm128_default_setting
    Call frm128_reset_simpanan
    Call frm128_jurujual
    
    frm128.Frame1.Visible = True
    
    frm128.TB1.SetFocus
Else

    frm128.Frame1.Visible = False
    
End If
End Sub

Private Sub L6_Text_Click()
'on error resume next
If frm128.Frame2.Visible = False Then

    If MDI_frm1.L20_Text = "Semua cawangan" Then
    
        Frm96.CMD2.Visible = True
        Frm96.CMD1.Visible = False
    
        Call Frm96_initial
        
        Frm96.Show vbModal
        
    End If
    
    Call frm128_pic_ena_disable
    Call frm128_default_setting
    Call frm128_reset_pulangan
    Call frm128_jurujual
    
    frm128.Frame2.Visible = True
    
    frm128.TB2.SetFocus
    
Else

    frm128.Frame2.Visible = False
    
End If
End Sub

Private Sub L7_Text_Click()
'on error resume next
If frm128.Frame3.Visible = False Then

    Call frm128_pic_ena_disable
    'Call frm128_default_setting
    
    frm128.L69_Text = -1 'Titik Pencarian Data
    frm128.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    frm128.L67_Text = 0 'Paparan Page ke-xxx
    frm128.L68_Text = 0
    
    GM_NEXT_PREV = 0
    
    Call frm128_report_simpanan_header
    Call frm128_report_simpanan
    
    frm128.Frame3.Visible = True
    
Else

    frm128.Frame3.Visible = False
    
End If
End Sub

Private Sub L8_Text_Click()
'on error resume next

GM_NEXT_PREV = 2 '0 : Next , 1 : Previous
            
Call frm68_senarai_pelanggan_header
Call frm68_senarai_pelanggan

Frm68.Show
Unload frm128
End Sub
Private Sub LV1_DblClick()
'on error resume next
Frm128_LM_No_ID = vbNullString

If IsNumeric(frm128.LV1.SelectedItem.Index) Then
    
    Frm128_LM_No_ID = frm128.LV1.SelectedItem.Index
    
    If Frm128_LM_No_ID <> vbNullString Then

        PopupMenu frm128_pm_menu
    
    Else
    
        MsgBox "Tiada Data.", vbInformation, "Info"
        
    End If
    
Else

    MsgBox "Tiada Data.", vbInformation, "Info"
    
End If
End Sub

