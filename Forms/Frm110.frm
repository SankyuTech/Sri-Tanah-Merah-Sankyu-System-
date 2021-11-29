VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm110 
   Caption         =   "Senarai Pengeluaran Invoice"
   ClientHeight    =   13035
   ClientLeft      =   120
   ClientTop       =   435
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
   ScaleHeight     =   13035
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Senarai Invoice"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11280
      Left            =   2040
      TabIndex        =   21
      Top             =   3240
      Visible         =   0   'False
      Width           =   22455
      Begin VB.TextBox TB1 
         BackColor       =   &H00FFFFFF&
         Height          =   1380
         Left            =   17625
         TabIndex        =   51
         Text            =   "TB1"
         Top             =   5040
         Width           =   3630
      End
      Begin VB.CommandButton CMD4 
         BackColor       =   &H80000004&
         Caption         =   "Carian"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   18360
         MaskColor       =   &H00400000&
         Picture         =   "Frm110.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Carian Maklumat Pembeli"
         Top             =   7320
         Width           =   2145
      End
      Begin VB.CommandButton CMD2 
         Caption         =   "Next"
         Height          =   810
         Left            =   18000
         MouseIcon       =   "Frm110.frx":09AA
         MousePointer    =   99  'Custom
         Picture         =   "Frm110.frx":0CB4
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Paparan Seterusnya"
         Top             =   9840
         Width           =   1095
      End
      Begin VB.CommandButton CMD3 
         Caption         =   "Back"
         Height          =   810
         Left            =   16800
         MouseIcon       =   "Frm110.frx":1D7E
         MousePointer    =   99  'Custom
         Picture         =   "Frm110.frx":2088
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Paparan Sebelum"
         Top             =   9840
         Width           =   1095
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   10650
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   16335
         _ExtentX        =   28813
         _ExtentY        =   18785
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Keyword :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   16080
         TabIndex        =   53
         Top             =   5040
         Width           =   1500
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "*** Sistem akan mencari keyword ini di dalam field ""No. Invoice"" , ""No. Tracking"" atau ""Remarks""."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   16800
         TabIndex        =   52
         Top             =   6480
         Width           =   4575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   ": RM   : RM   : RM   : RM  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   18960
         TabIndex        =   47
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Tunai                               Online Transfer                 Kad Kredit                   Simpanan Di Kedai"
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
         Left            =   16680
         TabIndex        =   46
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label L23_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L23_Text"
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
         Left            =   19440
         TabIndex        =   45
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Label L22_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L22_Text"
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
         Left            =   19440
         TabIndex        =   44
         Top             =   4200
         Width           =   2295
      End
      Begin VB.Label L20_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L20_Text"
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
         Left            =   19440
         TabIndex        =   43
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label L21_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L21_Text"
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
         Left            =   19440
         TabIndex        =   42
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label L14_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L14_Text"
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
         Left            =   19440
         TabIndex        =   40
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label L17_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L17_Text"
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
         Left            =   19440
         TabIndex        =   39
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label L13_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L13_Text"
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
         Left            =   19440
         TabIndex        =   38
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label L12_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L12_Text"
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
         Left            =   19440
         TabIndex        =   37
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label L15_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L15_Text"
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
         Left            =   19440
         TabIndex        =   36
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label L10_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L10_Text"
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
         Left            =   19080
         TabIndex        =   35
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label L11_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L11_Text"
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
         Left            =   19440
         TabIndex        =   34
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label L16_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L16_Text"
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
         Left            =   19440
         TabIndex        =   33
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label L18_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L18_Text"
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
         Left            =   19440
         TabIndex        =   32
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Ringkasan."
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
         Height          =   300
         Left            =   16560
         TabIndex        =   31
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   ":         : RM   : RM   : RM   : RM   : RM   : RM   : RM    : RM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   18960
         TabIndex        =   30
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm110.frx":3152
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   16680
         TabIndex        =   29
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label L6_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L6_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   18525
         TabIndex        =   25
         Top             =   10680
         Width           =   2295
      End
      Begin VB.Label L5_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L5_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   17640
         TabIndex        =   24
         Top             =   10680
         Width           =   735
      End
      Begin VB.Label L9_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L9_Text"
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
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   15375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Paparan Muka  :          / "
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   16800
         TabIndex        =   26
         Top             =   10680
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
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
      Height          =   3615
      Left            =   16080
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   6015
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
         Left            =   2280
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm110.frx":325C
         MousePointer    =   99  'Custom
         Picture         =   "Frm110.frx":3566
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2400
         Width           =   2865
      End
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
         ItemData        =   "Frm110.frx":5B30
         Left            =   1750
         List            =   "Frm110.frx":5B32
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   360
         Width           =   4005
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
         ItemData        =   "Frm110.frx":5B34
         Left            =   1750
         List            =   "Frm110.frx":5B36
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   720
         Width           =   4005
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   1755
         TabIndex        =   13
         Top             =   1545
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
         Format          =   142409728
         CurrentDate     =   41561
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   360
         Left            =   1755
         TabIndex        =   14
         Top             =   1905
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
         Format          =   142409728
         CurrentDate     =   41561
      End
      Begin VB.Label L24_Text 
         BackColor       =   &H0080FFFF&
         Caption         =   "L24_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1080
         TabIndex        =   49
         Top             =   3000
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label L25_Text 
         BackColor       =   &H0080FFFF&
         Caption         =   "L25_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1080
         TabIndex        =   48
         Top             =   3360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label L19_Text 
         BackColor       =   &H0080FFFF&
         Caption         =   "L19_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1080
         TabIndex        =   41
         Top             =   2640
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label L3_Text 
         BackColor       =   &H0080FFFF&
         Caption         =   "L3_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   240
         TabIndex        =   20
         Top             =   3000
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label L2_Text 
         BackColor       =   &H0080FFFF&
         Caption         =   "L2_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   240
         TabIndex        =   19
         Top             =   2640
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label L4_Text 
         BackColor       =   &H0080FFFF&
         Caption         =   "L4_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   240
         TabIndex        =   18
         Top             =   3360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label7 
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
         Left            =   120
         TabIndex        =   16
         Top             =   390
         Width           =   1575
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
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label62 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Mula  * :"
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
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila buat pilihan tempoh report."
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
         Left            =   360
         TabIndex        =   10
         Top             =   1200
         Width           =   6690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Invoice * :"
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
         Left            =   120
         TabIndex        =   9
         Top             =   750
         Width           =   1575
      End
   End
   Begin VB.PictureBox Pic2 
      BorderStyle     =   0  'None
      Height          =   11415
      Left            =   720
      ScaleHeight     =   11415
      ScaleWidth      =   23535
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   23535
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   10545
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   360
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   18600
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
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "** Senarai yang dipaparkan hanya bagi invoice jualan , invoice tempahan emas dan invoice servis SAHAJA."
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
         Height          =   780
         Left            =   12360
         TabIndex        =   6
         Top             =   2520
         Width           =   7215
      End
      Begin VB.Label L8_Text 
         Caption         =   "L8_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   12360
         TabIndex        =   5
         Top             =   9240
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label L7_Text 
         Caption         =   "L7_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   12360
         TabIndex        =   4
         Top             =   9000
         Visible         =   0   'False
         Width           =   2295
      End
   End
   Begin VB.PictureBox Pic1 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2415
      ScaleWidth      =   5865
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   5865
   End
   Begin VB.Label L1_Text 
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
      Left            =   120
      MouseIcon       =   "Frm110.frx":5B38
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Menu Frm110_PM_menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm110_SM_excel 
         Caption         =   "Export excel report"
      End
      Begin VB.Menu Frm110_SM_ringkasan 
         Caption         =   "Ringkasan report"
         Visible         =   0   'False
      End
      Begin VB.Menu frm110_sm_cetak_invoice_ini 
         Caption         =   "Cetak invoice ini SAHAJA"
      End
      Begin VB.Menu frm110_sm_cetak_semua_invoice 
         Caption         =   "Cetak semua invoice"
      End
   End
End
Attribute VB_Name = "Frm110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD1_Click()
'on error resume next
If Frm110.CBB1 = vbNullString Then

    MsgBox "Sila pilih [Jenis invoice].", vbInformation, "Info"

    Exit Sub
    
End If
If Frm110.CBB2 = vbNullString Then

    MsgBox "Sila pilih [Cawangan].", vbInformation, "Info"

    Exit Sub
    
End If

Frm110.L2_Text = Frm110.CBB1 'Jenis invoice
Frm110.L3_Text = Frm110.DTPicker1 'Tarikh mula
Frm110.L4_Text = Frm110.DTPicker2 'Tarikh akhir

If Frm110.L2_Text <> vbNullString And Frm110.L3_Text <> vbNullString And Frm110.L4_Text <> vbNullString And Frm110.L19_Text <> vbNullString Then
    
    Frm110.L25_Text = "0" '0 : Filter , 1 : Keyword
    
    GM_NEXT_PREV = 0
    Frm110.L7_Text = -1 'Titik Pencarian Data
    Frm110.L8_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm110.L5_Text = 0 'Paparan Page ke-xxx
    Frm110.L19_Text = Frm110.CBB2 'Cawangan
    
    Frm110.TB1 = vbNullString
    
    Call Frm110_senarai_jualan_header
    Call Frm110_senarai_jualan
    
    If Frm110.L10_Text = 0 Then
        MsgBox "Tiada invoice dijumpai.", vbInformation, "Info"
    End If

End If
End Sub
Private Sub CMD2_Click()
'on error resume next
Dim Frm110_LM_CURR_PAGE As Double
Dim Frm110_LM_TOTAL_PAGE As Double

Frm110_LM_CURR_PAGE = 0
Frm110_LM_TOTAL_PAGE = 0

If Frm110.L5_Text <> vbNullString And IsNumeric(Frm110.L5_Text) Then
    If Frm110.L6_Text <> vbNullString And IsNumeric(Frm110.L6_Text) Then
        Frm110_LM_CURR_PAGE = Frm110.L5_Text
        Frm110_LM_TOTAL_PAGE = Frm110.L6_Text
        
        If Frm110_LM_CURR_PAGE < Frm110_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm110_senarai_jualan_header
            Call Frm110_senarai_jualan
            
        End If
    End If
End If
End Sub
Private Sub CMD3_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm110_senarai_jualan_header
Call Frm110_senarai_jualan
End Sub

Private Sub CMD4_Click()
'on error resume next
Frm110.L24_Text = vbNullString 'Keyword

If Frm110.TB1 = vbNullString Then
    MsgBox "Sila masukkan [Keyword].", vbExclamation, "Info"
    
    Frm110.TB1.SetFocus
    Exit Sub
End If

If Frm110.TB1 <> vbNullString Then

    If InStr(1, Frm110.TB1, "*") <> 0 Or InStr(1, Frm110.TB1, "/") <> 0 Or InStr(1, Frm110.TB1, "\") <> 0 Or InStr(1, Frm110.TB1, "'") <> 0 Then
        MsgBox "Keyword mengandungi simbol yang tidak sah.", vbExclamation, "Info"
        
        Frm110.TB1.SetFocus
        Exit Sub
    End If

End If

Frm110.L24_Text = UCase(Frm110.TB1) 'Keyword

If Frm110.L2_Text <> vbNullString And Frm110.L24_Text <> vbNullString Then
    
    Frm110.L25_Text = "1" '0 : Filter , 1 : Keyword
    
    GM_NEXT_PREV = 0
    Frm110.L2_Text = Frm110.CBB1 'Jenis invoice
    Frm110.L7_Text = -1 'Titik Pencarian Data
    Frm110.L8_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm110.L5_Text = 0 'Paparan Page ke-xxx
    Frm110.L19_Text = Frm110.CBB2 'Cawangan
    
    Call Frm110_senarai_jualan_header
    Call Frm110_senarai_jualan2
    
    If Frm110.L10_Text = 0 Then
        MsgBox "Tiada maklumat dijumpai.", vbInformation, "Info"
    End If

End If
End Sub

Private Sub Form_Load()
'on error resume next
Frm110.CBB1.Clear

user_level = MDI_frm1.L4_Text

Frm110.CBB1.AddItem "Semua Invoice"

If user_level <> "Guest/User" And user_level <> "Administration" Then

    Frm110.CBB1.AddItem "Invoice Rasmi"
    Frm110.CBB1.AddItem "Invoice Tidak Rasmi"

End If

Frm110.CBB1 = "Semua Invoice"

Frm110.DTPicker1 = DateTime.Date
Frm110.DTPicker2 = DateTime.Date

Frm110.L2_Text = vbNullString
Frm110.L3_Text = vbNullString
Frm110.L4_Text = vbNullString
End Sub

Private Sub frm110_sm_cetak_invoice_ini_Click()
'on error resume next
Dim rs20 As ADODB.Recordset
LM_FOUND = 0
Frm110_LM_No_ID = vbNullString

If IsNumeric(Frm110.LV1.SelectedItem.Index) Then
    
    Frm110_LM_No_ID = Frm110.LV1.ListItems(Frm110.LV1.SelectedItem.Index)
    
    If Frm110_LM_No_ID <> vbNullString Then

        Note = "Adakah anda ingin cetak ini?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then

            Set rs20 = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs20.Open "select * from 22_jualan where ID='" & Frm110_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs20.EOF Then
                
                If Not IsNull(rs20!Menu) Then
                
                    If Not IsNull(rs20!no_resit) Then
                        G_No_RESIT_JUALAN = rs20!no_resit
                        
                        Note = "Adakah anda ingin preview invoice ini sebelum cetak?." & vbCrLf & _
                                vbNullString & vbCrLf & _
                                "No. Invoice : " & G_No_RESIT_JUALAN & vbCrLf & _
                                vbNullString & vbCrLf & _
                                "Jika YA : Invoice akan dipaparkan dahulu." & vbCrLf & _
                                "Jika TIDAK : Invoice tidak akan dipaparkan dan sistem akan cetak invoice tersebut."
                                
                        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
                        
                        If Answer = vbNo Then
                            G_PREVIEW = 0
                        End If
                        If Answer = vbYes Then
                            G_PREVIEW = 1
                        End If
                        
                        G_No_RESIT_SERVIS = G_No_RESIT_JUALAN
                        G_No_INV_BOOK = G_No_RESIT_JUALAN
                        
                        If Not IsNull(rs20!cawangan) Then
                        
                            G_KEDAI = rs20!cawangan
                        
                            If rs20!Menu = "0" Then Call Frm84_Resit_Jualan
                            If rs20!Menu = "1" Then Call Frm92_Resit_Servis
                            If rs20!Menu = "2" Then Call Frm94_invoice_deposit_tempahan
                            If rs20!Menu = "3" Then Call Frm94_invoice_siap_tempahan
                            If rs20!Menu = "4" Then Call frm118_cetak_inv_vou
                        
                            If G_PREVIEW = 0 Then
                            
                                MsgBox "Invoice telah berjaya dicetak.", vbInformation, "Info"
                                        
                            End If
                            
                        End If
                        
                    End If
                        
                Else
                    
                    MsgBox "Berlaku ralat semasa cuba cetak invoice ini. Sila hubungi developer dan nyatakan no. invoice yang berlaku ralat ini.", vbCritical, "Error"
                    
                End If
                
            End If
            
            rs20.Close
            Set rs20 = Nothing

        End If
        
    End If
    
End If
End Sub

Private Sub frm110_sm_cetak_semua_invoice_Click()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim rs20 As ADODB.Recordset
                
LM_FOUND = 0
Frm110_LM_No_ID = vbNullString

If IsNumeric(Frm110.LV1.SelectedItem.Index) Then
    
    Frm110_LM_No_ID = Frm110.LV1.ListItems(Frm110.LV1.SelectedItem.Index)
    
    If Frm110_LM_No_ID <> vbNullString Then

        Note = "Adakah anda ingin cetak semua invoice ini?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Sistem akan mencetak semua invoice , sila tunggu sistem selesai cetak semua invoice ini." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Teruskan?"
                
        Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
        
        If Answer = vbYes Then
            
            x = 0
            
            TM = Frm110.L3_Text 'Tarikh Mula
            TA = Frm110.L4_Text 'Tarikh Akhir
            
            If Frm110.L2_Text = "Semua Invoice" Then
            
                Frm110_LM_SEARCH_1 = 0
                Frm110_LM_SEARCH_2 = 1
                
            ElseIf Frm110.L2_Text = "Invoice Rasmi" Then
            
                Frm110_LM_SEARCH_1 = 1
                Frm110_LM_SEARCH_2 = 1
                
            ElseIf Frm110.L2_Text = "Invoice Tidak Rasmi" Then
            
                Frm110_LM_SEARCH_1 = 0
                Frm110_LM_SEARCH_2 = 0
            
            End If
            
            user_level = MDI_frm1.L4_Text
            
            LM_INVOICE_RASMI = 0
                
            If user_level = "Guest/User" Then
                Frm85_LM_SEARCH_6 = 1
                Frm85_LM_SEARCH_6_LOGIC = "="
                LM_INVOICE_RASMI = 1
                Frm85_LM_SEARCH_7 = 1
                Frm85_LM_SEARCH_7_LOGIC = "="
            Else
                Frm85_LM_SEARCH_6 = 0
                Frm85_LM_SEARCH_6_LOGIC = "="
                
                Frm85_LM_SEARCH_7 = 1
                Frm85_LM_SEARCH_7_LOGIC = "="
            End If
            
            If user_level = "Administration" Then
            
                Frm110_LM_SEARCH_1 = 1
                Frm110_LM_SEARCH_2 = 1
                
            End If
            
            If Frm110.L19_Text = "Semua cawangan" Then
            
                Frm85_SEARCH_8 = Null
                Frm85_SEARCH_8_LOGIC = "<>"
                Frm85_SEARCH_9 = Null
                Frm85_SEARCH_9_LOGIC = "<>"
                
            Else
            
                Frm85_SEARCH_8 = Frm110.L19_Text
                Frm85_SEARCH_8_LOGIC = "="
                Frm85_SEARCH_9 = "HQ"
                Frm85_SEARCH_9_LOGIC = "="
                
            End If
            
            G_PREVIEW = 0
            
            Set rs20 = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs20.Open "select * from 22_jualan where status = 1 AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (bil_rasmi='" & Frm110_LM_SEARCH_1 & "' OR bil_rasmi='" & Frm110_LM_SEARCH_2 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC , no_resit ASC", cn, adOpenKeyset, adLockOptimistic

            While rs20.EOF = False
                    
                If Not IsNull(rs20!Menu) Then
                
                    If Not IsNull(rs20!no_resit) Then
                    
                        x = x + 1
                        
                        G_No_RESIT_JUALAN = rs20!no_resit
                        G_No_RESIT_SERVIS = rs20!no_resit
                        G_No_INV_BOOK = rs20!no_resit

                        If rs20!Menu = "0" Then Call Frm84_Resit_Jualan
                        If rs20!Menu = "1" Then Call Frm92_Resit_Servis
                        If rs20!Menu = "2" Then Call Frm94_invoice_deposit_tempahan
                        If rs20!Menu = "3" Then Call Frm94_invoice_siap_tempahan
                        If rs20!Menu = "4" Then Call frm118_cetak_inv_vou
                        
                    End If
                    
                Else
                    
                    MsgBox "Berlaku ralat semasa cuba cetak invoice ini. Sila hubungi developer dan nyatakan no. invoice yang berlaku ralat ini.", vbCritical, "Error"
                    
                End If
                
                rs20.MoveNext
            Wend
            
            rs20.Close
            Set rs20 = Nothing
    
        End If

    End If
    
End If
End Sub

Private Sub Frm110_SM_excel_Click()
'on error resume next
If Frm110.L25_Text = "0" Then '0 : Filter , 1 : Keyword
    Call frm110_excel_filter
Else
    Call frm110_excel_keyword
End If
End Sub
Private Sub Frm110_SM_ringkasan_Click()
'on error resume next
Dim TM As Date
Dim TA As Date

'Set rs = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from default_setting where default1='" & "default" & "'", cn, adOpenKeyset, adLockOptimistic

'If Not rs.EOF Then
    
'    If Not IsNull(rs!default_printer) Then LM_PRINTER = rs!default_printer
    
'End If

'rs.Close
'Set rs = Nothing

'For Each oPrn In Printers
'    If oPrn.DeviceName = LM_PRINTER Then
'        Set Printer = oPrn
'        PRINTER_FOUND = 1 '0 : Printer Not Found , 1 : Printer Found
'        Exit For
'    End If
'Next

TM = Frm110.L3_Text 'Tarikh Mula
TA = Frm110.L4_Text 'Tarikh Akhir

If Frm110.L2_Text = "Semua Invoice" Then

    Frm110_LM_SEARCH_1 = 0
    Frm110_LM_SEARCH_2 = 1
    
ElseIf Frm110.L2_Text = "Invoice Rasmi" Then

    Frm110_LM_SEARCH_1 = 1
    Frm110_LM_SEARCH_2 = 1
    
ElseIf Frm110.L2_Text = "Invoice Tidak Rasmi" Then

    Frm110_LM_SEARCH_1 = 0
    Frm110_LM_SEARCH_2 = 0

End If

Report74.Caption = "Senarai Invoice"

'### Reset maklumat kedai ### - Start
Report74.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report74.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report74.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report74.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report74.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report74.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report74.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report74.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report74.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report74.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Report74.Sections("Section4").Controls("L1").Caption = Frm110.L9_Text 'Header Report
Report74.Sections("Section5").Controls("L2").Caption = "Report dikeluarkan pada " & Now

Report74.Sections("Section5").Controls("L3").Caption = Frm110.L10_Text 'Bilangan invoice
Report74.Sections("Section5").Controls("L4").Caption = Frm110.L11_Text 'Jumlah harga barang
Report74.Sections("Section5").Controls("L5").Caption = Frm110.L12_Text 'Jumlah trade in
Report74.Sections("Section5").Controls("L6").Caption = Frm110.L13_Text 'Jumlah adjustment
Report74.Sections("Section5").Controls("L7").Caption = Frm110.L14_Text 'Jumlah pos laju
Report74.Sections("Section5").Controls("L8").Caption = Frm110.L15_Text 'Jumlah bayaran bersih

Report74.Sections("Section5").Controls("L9").Caption = Frm110.L18_Text 'Jumlah diskaun
Report74.Sections("Section5").Controls("L11").Caption = Frm110.L16_Text 'Jumlah kupon
Report74.Sections("Section5").Controls("L12").Caption = Frm110.L17_Text 'Jumlah mata ganjaran

'### Paparan Penyata ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where status = 1 AND (bil_rasmi='" & Frm110_LM_SEARCH_1 & "' OR bil_rasmi='" & Frm110_LM_SEARCH_2 & "') AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report74.DataSource = rs
    Report74.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing
'### Paparan Penyata ### - End
End Sub
Private Sub L1_Text_Click()
'on error resume next
If Frm110.Frame1.Visible = False Then

    Call Frm110_initial_setting
    
    Frm110.L25_Text = "0" '0 : Filter , 1 : Keyword
    
    Frm110.Frame1.Visible = True
    
Else

    Frm110.Frame1.Visible = False
    
End If
End Sub


Private Sub LV1_DblClick()
'on error resume next
Frm110_LM_No_ID = vbNullString

If IsNumeric(Frm110.LV1.SelectedItem.Index) Then
    
    Frm110_LM_No_ID = Frm110.LV1.SelectedItem.Index
    
    If Frm110_LM_No_ID <> vbNullString Then
    
    
        PopupMenu Frm110_PM_menu1
    
    Else
    
        MsgBox "Tiada Data.", vbInformation, "Info"
        
    End If
    
Else

    MsgBox "Tiada Data.", vbInformation, "Info"
    
End If
End Sub

Private Sub MSFlexGrid1_DblClick()
'On Error Resume Next
Frm110_LM_ID = vbNullString

If Frm110.MSFlexGrid1 <> vbNullString Then
    
    If IsNumeric(Frm110.MSFlexGrid1) Then
        Frm110_LM_ID = Frm110.MSFlexGrid1.TextMatrix(Frm110.MSFlexGrid1, 2) 'Status
        
        PopupMenu Frm110_PM_menu1, vbPopupMenuRightButton
        
    Else
        MsgBox "Tiada data.", vbeclamation, "Info"
    End If
        
End If
End Sub


