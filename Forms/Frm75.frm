VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm75 
   Caption         =   "Report GST"
   ClientHeight    =   12915
   ClientLeft      =   60
   ClientTop       =   -13440
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
   Icon            =   "Frm75.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic3 
      BorderStyle     =   0  'None
      Height          =   11415
      Left            =   1080
      ScaleHeight     =   11415
      ScaleWidth      =   23535
      TabIndex        =   12
      Top             =   2880
      Visible         =   0   'False
      Width           =   23535
      Begin VB.CommandButton CMD23 
         BackColor       =   &H00FFFFFF&
         Height          =   840
         Left            =   12960
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm75.frx":0ECA
         MousePointer    =   99  'Custom
         Picture         =   "Frm75.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Paparan seterusnya"
         Top             =   10440
         Width           =   1100
      End
      Begin VB.CommandButton CMD21 
         BackColor       =   &H00FFFFFF&
         Height          =   840
         Left            =   11760
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm75.frx":1AFA
         MousePointer    =   99  'Custom
         Picture         =   "Frm75.frx":1E04
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Paparan sebelumnya"
         Top             =   10440
         Width           =   1100
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   10035
         Left            =   120
         TabIndex        =   36
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   360
         Width           =   13965
         _ExtentX        =   24633
         _ExtentY        =   17701
         _Version        =   393216
         Rows            =   1
         Cols            =   0
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   12648384
         BackColorSel    =   16777215
         ForeColorSel    =   16711680
         BackColorBkg    =   16777215
         GridColor       =   0
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
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
      Begin VB.Label L75_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L75_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7320
         TabIndex        =   46
         Top             =   10560
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L69_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L69_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   6360
         TabIndex        =   45
         Top             =   10560
         Visible         =   0   'False
         Width           =   855
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
         Left            =   10890
         TabIndex        =   44
         Top             =   11100
         Width           =   375
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
         Left            =   11400
         TabIndex        =   43
         Top             =   11100
         Width           =   615
      End
      Begin VB.Label Label22 
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
         Left            =   9480
         TabIndex        =   42
         Top             =   11100
         Width           =   2295
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Harga SR : RM"
         Height          =   255
         Left            =   14160
         TabIndex        =   35
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label L14_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L14_Text"
         Height          =   255
         Left            =   16140
         TabIndex        =   34
         Top             =   360
         Width           =   1995
      End
      Begin VB.Label L15_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L15_Text"
         Height          =   255
         Left            =   16155
         TabIndex        =   33
         Top             =   600
         Width           =   1995
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Cukai SR  : RM"
         Height          =   255
         Left            =   14160
         TabIndex        =   32
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Cukai ZR  : RM"
         Height          =   255
         Left            =   14160
         TabIndex        =   31
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label L23_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L23_Text"
         Height          =   255
         Left            =   16155
         TabIndex        =   30
         Top             =   1080
         Width           =   1995
      End
      Begin VB.Label L22_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L22_Text"
         Height          =   255
         Left            =   16125
         TabIndex        =   29
         Top             =   840
         Width           =   1995
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Harga ZR : RM"
         Height          =   255
         Left            =   14160
         TabIndex        =   28
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label L10_Text 
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   120
         Width           =   14295
      End
   End
   Begin VB.PictureBox Pic1 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   13800
      ScaleHeight     =   1935
      ScaleWidth      =   6135
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton CMD2 
         BackColor       =   &H000080FF&
         Caption         =   "Batal"
         Height          =   405
         Left            =   3240
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm75.frx":2743
         MousePointer    =   99  'Custom
         TabIndex        =   39
         Top             =   1320
         Width           =   1545
      End
      Begin VB.CommandButton CMD1 
         BackColor       =   &H000080FF&
         Caption         =   "Carian"
         Default         =   -1  'True
         Height          =   405
         Left            =   1560
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm75.frx":2A4D
         MousePointer    =   99  'Custom
         TabIndex        =   38
         Top             =   1320
         Width           =   1545
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   1800
         TabIndex        =   5
         Top             =   480
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
         Format          =   415301632
         CurrentDate     =   41561
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   360
         Left            =   1800
         TabIndex        =   6
         Top             =   840
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
         Format          =   415301632
         CurrentDate     =   41561
      End
      Begin VB.Label Label62 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Mula *    :"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   525
         Width           =   1695
      End
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Akhir *    :"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   885
         Width           =   1695
      End
      Begin VB.Label Label64 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila buat pilihan tarikh bagi tempoh report GST."
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   5775
      End
      Begin VB.Label L6_Text 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label L7_Text 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.PictureBox Pic2 
      BorderStyle     =   0  'None
      Height          =   11415
      Left            =   11760
      ScaleHeight     =   11415
      ScaleWidth      =   23535
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   23535
      Begin VB.CommandButton CMD26 
         BackColor       =   &H00FFFFFF&
         Height          =   840
         Left            =   18120
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm75.frx":2D57
         MousePointer    =   99  'Custom
         Picture         =   "Frm75.frx":3061
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Paparan seterusnya"
         Top             =   10450
         Width           =   1300
      End
      Begin VB.CommandButton CMD25 
         BackColor       =   &H00FFFFFF&
         Height          =   840
         Left            =   16680
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm75.frx":3987
         MousePointer    =   99  'Custom
         Picture         =   "Frm75.frx":3C91
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Paparan sebelumnya"
         Top             =   10450
         Width           =   1300
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   10035
         Left            =   120
         TabIndex        =   37
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   360
         Width           =   19365
         _ExtentX        =   34158
         _ExtentY        =   17701
         _Version        =   393216
         Rows            =   1
         Cols            =   0
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   12648384
         BackColorSel    =   16777215
         ForeColorSel    =   16711680
         BackColorBkg    =   16777215
         GridColor       =   0
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
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
      Begin VB.Label L63_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L63_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   19680
         TabIndex        =   52
         Top             =   7920
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label L62_Text 
         BackColor       =   &H8000000A&
         Caption         =   "L62_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   19680
         TabIndex        =   51
         Top             =   7560
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label L61_Text 
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
         Height          =   285
         Left            =   16140
         TabIndex        =   50
         Top             =   11040
         Width           =   705
      End
      Begin VB.Label L60_Text 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   15555
         TabIndex        =   49
         Top             =   11040
         Width           =   465
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Harga ZR : RM"
         Height          =   255
         Left            =   7440
         TabIndex        =   27
         Top             =   10440
         Width           =   2415
      End
      Begin VB.Label L20_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L20_Text"
         Height          =   255
         Left            =   9420
         TabIndex        =   26
         Top             =   10440
         Width           =   1995
      End
      Begin VB.Label L21_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L21_Text"
         Height          =   255
         Left            =   13140
         TabIndex        =   25
         Top             =   10440
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Cukai ZR  : RM"
         Height          =   255
         Left            =   11160
         TabIndex        =   24
         Top             =   10440
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Cukai SR  : RM"
         Height          =   255
         Left            =   3720
         TabIndex        =   17
         Top             =   10440
         Width           =   2415
      End
      Begin VB.Label L13_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L13_Text"
         Height          =   255
         Left            =   5715
         TabIndex        =   16
         Top             =   10440
         Width           =   1995
      End
      Begin VB.Label L8_Text 
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Width           =   14775
      End
      Begin VB.Label L12_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L12_Text"
         Height          =   255
         Left            =   2205
         TabIndex        =   14
         Top             =   10440
         Width           =   1995
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Harga SR : RM"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   10440
         Width           =   2415
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Caption         =   "Paparan Muka :          /"
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
         Left            =   14280
         TabIndex        =   53
         Top             =   11040
         Width           =   2505
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Kutipan GST  : RM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   54
      Top             =   1320
      Width           =   4935
   End
   Begin VB.Label L19_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   240
      TabIndex        =   23
      Top             =   2400
      Width           =   15255
   End
   Begin VB.Label L18_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   22
      Top             =   1800
      Width           =   8295
   End
   Begin VB.Label L17_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   21
      Top             =   1320
      Width           =   8295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Bayaran GST : RM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   20
      Top             =   1800
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Kesimpulan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label L5_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kutipan GST"
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
      Left            =   3360
      MouseIcon       =   "Frm75.frx":45D0
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label L4_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bayaran GST"
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
      MouseIcon       =   "Frm75.frx":48DA
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   120
      Width           =   1335
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
      Left            =   120
      MouseIcon       =   "Frm75.frx":4BE4
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Menu Frm75_Menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm75_Export_Excel_1 
         Caption         =   "Export Excel Report"
      End
   End
   Begin VB.Menu Frm75_Menu3 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm75_Export_Excel_3 
         Caption         =   "Export Excel Report"
      End
   End
End
Attribute VB_Name = "Frm75"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD1_Click()
'on error resume next
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
    vbNullString & vbCrLf & _
    "Teruskan ?"
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    Frm75.L6_Text = Frm75.DTPicker1
    Frm75.L7_Text = Frm75.DTPicker2
    
    Frm75.L69_Text = -1 'Titik Pencarian Data
    Frm75.L75_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm75.L67_Text = 0 'Paparan Page ke-xxx
    Frm75.L68_Text = 0

    Frm75.L62_Text = -1 'Start Point
    Frm75.L60_Text = 0 'Current Page
    Frm75.L63_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    
    GM_NEXT_PREV = 0
    
    Call Frm75_report_gst_bayar_header
    Call Frm75_report_gst_bayar
    
    Call Frm75_report_gst_kutip_header
    Call Frm75_report_gst_kutip
    
    'Call Frm75_Report_GST_Header
    'Call Frm75_Report_GST_BAYARAN
    'Call Frm75_Report_GST_KUTIPAN
    
    Frm75.Pic1.Visible = False
    Frm75.Pic2.Visible = True
    
    If Frm75.L19_Text <> vbNullString Then MsgBox Frm75.L19_Text
End If
End Sub

Private Sub CMD21_Click()
'on error resume next
Dim frm75_LM_CURR_PAGE As Double
Dim frm75_LM_TOTAL_PAGE As Double

frm75_LM_CURR_PAGE = 0
frm75_LM_TOTAL_PAGE = 0

If Frm75.L67_Text <> vbNullString And IsNumeric(Frm75.L67_Text) Then
    If Frm75.L68_Text <> vbNullString And IsNumeric(Frm75.L68_Text) Then
        frm75_LM_CURR_PAGE = Frm75.L67_Text
        frm75_LM_TOTAL_PAGE = Frm75.L68_Text
        
        If frm75_LM_CURR_PAGE <> 1 And frm75_LM_CURR_PAGE <> 0 Then
        
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
                    
            Call Frm75_report_gst_kutip_header
            Call Frm75_report_gst_kutip
            
        End If

    End If
End If
End Sub
Private Sub CMD23_Click()
'on error resume next
Dim frm75_LM_CURR_PAGE As Double
Dim frm75_LM_TOTAL_PAGE As Double

frm75_LM_CURR_PAGE = 0
frm75_LM_TOTAL_PAGE = 0

If Frm75.L67_Text <> vbNullString And IsNumeric(Frm75.L67_Text) Then
    If Frm75.L68_Text <> vbNullString And IsNumeric(Frm75.L68_Text) Then
        frm75_LM_CURR_PAGE = Frm75.L67_Text
        frm75_LM_TOTAL_PAGE = Frm75.L68_Text
        
        If frm75_LM_CURR_PAGE < frm75_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm75_report_gst_kutip_header
            Call Frm75_report_gst_kutip
            
        End If
    End If
End If
End Sub

Private Sub CMD25_Click()
'on error resume next
Dim frm75_LM_CURR_PAGE As Double
Dim frm75_LM_TOTAL_PAGE As Double

frm75_LM_CURR_PAGE = 0
frm75_LM_TOTAL_PAGE = 0

If Frm75.L60_Text <> vbNullString And IsNumeric(Frm75.L60_Text) Then
    If Frm75.L61_Text <> vbNullString And IsNumeric(Frm75.L61_Text) Then
        frm75_LM_CURR_PAGE = Frm75.L60_Text
        frm75_LM_TOTAL_PAGE = Frm75.L61_Text
        
        If frm75_LM_CURR_PAGE <> 1 And frm75_LM_CURR_PAGE <> 0 Then
        
        GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
        
        Call Frm75_report_gst_bayar_header
        Call Frm75_report_gst_bayar
            
        End If
    End If
End If
End Sub
Private Sub CMD26_Click()
'on error resume next
Dim frm75_LM_CURR_PAGE As Double
Dim frm75_LM_TOTAL_PAGE As Double

frm75_LM_CURR_PAGE = 0
frm75_LM_TOTAL_PAGE = 0

If Frm75.L60_Text <> vbNullString And IsNumeric(Frm75.L60_Text) Then
    If Frm75.L61_Text <> vbNullString And IsNumeric(Frm75.L61_Text) Then
        frm75_LM_CURR_PAGE = Frm75.L60_Text
        frm75_LM_TOTAL_PAGE = Frm75.L61_Text
        
        If frm75_LM_CURR_PAGE < frm75_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm75_report_gst_bayar_header
            Call Frm75_report_gst_bayar
            
        End If
    End If
End If
End Sub


Private Sub Form_Load()
'on error resume next
Frm75.DTPicker1 = DateTime.Date
Frm75.DTPicker2 = DateTime.Date
Call Frm75_Initial_Setting

'### Bayaran
Frm75.L12_Text = "0.00" 'jumlah SR
Frm75.L13_Text = "0.00" 'Cukai SR
Frm75.L20_Text = "0.00" 'Harga ZR
Frm75.L21_Text = "0.00" 'Cukai ZR

'### Kutipan
Frm75.L14_Text = "0.00" 'jumlah SR
Frm75.L15_Text = "0.00" 'Cukai SR
Frm75.L22_Text = "0.00" 'Harga ZR
Frm75.L23_Text = "0.00" 'Cukai ZR
End Sub
Private Sub Frm75_Export_Excel_1_Click()
'On Error Resume Next
Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report. Teruskan ?"
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 20
        .Columns("C").ColumnWidth = 20
        .Columns("D").ColumnWidth = 20
        .Columns("E").ColumnWidth = 40
        .Columns("F").ColumnWidth = 20
        .Columns("G").ColumnWidth = 20
        .Columns("H").ColumnWidth = 20
        .Columns("I").ColumnWidth = 20
        .Columns("J").ColumnWidth = 20
        .Columns("K").ColumnWidth = 10
        .Columns("L").ColumnWidth = 10
        .Columns("M").ColumnWidth = 10
        .Columns("N").ColumnWidth = 10
        .Columns("O").ColumnWidth = 10
        .Columns("P").ColumnWidth = 10
        .Columns("Q").ColumnWidth = 20
        
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
            LM_NAMA_HEADER = "HQ"
            
        Else
            
            LM_NAMA_HEADER = MDI_frm1.L20_Text
            
        End If
                
        '### Maklumat kedai ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 56_maklumat_kedai where cawangan='" & LM_NAMA_HEADER & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!nama_kedai) Then
                .Cells(1, 5) = rs!nama_kedai
                .Cells(1, 5).Font.Name = "Times New Roman"
            End If
            If Not IsNull(rs!no_pendaftaran) Then .Cells(2, 5) = rs!no_pendaftaran
            If Not IsNull(rs!alamat) Then .Cells(3, 5) = rs!alamat
            If Not IsNull(rs!no_tel) Then .Cells(4, 5) = rs!no_tel
            If Not IsNull(rs!no_id_gst) Then .Cells(5, 5) = rs!no_id_gst
        End If
        
        rs.Close
        Set rs = Nothing
        '### Maklumat kedai ### - End
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        .Cells(7, 1) = Frm75.L8_Text 'Header Report
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh"
        .Cells(8, 3) = "No. Resit"
        .Cells(8, 4) = "No. ID GST (Supplier)"
        .Cells(8, 5) = "Perkara"
        .Cells(8, 6) = "Jumlah Harga SR (RM)"
        .Cells(8, 7) = "Jumlah Cukai SR (RM)"
        .Cells(8, 8) = "Jumlah Harga ZR (RM)"
        .Cells(8, 9) = "Jumlah Cukai ZR (RM)"
    
        For i = 1 To 9
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
        
        Y = 0
        For x = 1 To Frm75.MSFlexGrid1.Rows - 1
            Y = Y + 1
            .Cells(8 + Y, 1) = Y
            .Cells(8 + Y, 2) = "'" & Frm75.MSFlexGrid1.TextMatrix(x, 3) 'Tarikh
            .Cells(8 + Y, 2).HorizontalAlignment = xlCenter
            .Cells(8 + Y, 3) = "'" & Frm75.MSFlexGrid1.TextMatrix(x, 4) 'No. Resit
            .Cells(8 + Y, 4) = "'" & Frm75.MSFlexGrid1.TextMatrix(x, 5) 'No. ID GST Supplier
            .Cells(8 + Y, 5) = Frm75.MSFlexGrid1.TextMatrix(x, 6) 'Perkara
            .Cells(8 + Y, 6).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 6) = Frm75.MSFlexGrid1.TextMatrix(x, 7) 'Jumlah Harga SR
            .Cells(8 + Y, 6).HorizontalAlignment = xlCenter
            .Cells(8 + Y, 7).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 7) = Frm75.MSFlexGrid1.TextMatrix(x, 8) 'Jumlah Cukai SR
            .Cells(8 + Y, 7).HorizontalAlignment = xlCenter
            .Cells(8 + Y, 8).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 8) = Frm75.MSFlexGrid1.TextMatrix(x, 9) 'Jumlah Harga ZR
            .Cells(8 + Y, 8).HorizontalAlignment = xlCenter
            .Cells(8 + Y, 9).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 9) = Frm75.MSFlexGrid1.TextMatrix(x, 10) 'Jumlah Cukai ZR
            .Cells(8 + Y, 9).HorizontalAlignment = xlCenter
            For Col = 1 To 9
                .Cells(8 + Y, Col).Borders.LineStyle = xlContinuous
            Next Col
        Next x
    
        Y = Y + 2
        .Cells(8 + Y, 1) = "Jumlah Harga Barang Yang Dikenakan Cukai Standard Rated (SR) : RM " & Frm75.L12_Text
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah Bayaran Cukai Standard Rated (SR) : RM " & Frm75.L13_Text
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah Harga Barang Yang Dikenakan Cukai Zero Rated (ZR) : RM " & Frm75.L20_Text
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah Bayaran Cukai Zero Rated (ZR) : RM " & Frm75.L21_Text
        
    End With
    
    ' This makes Excel visible
    xlObject.Visible = True
    xlObject.EnableEvents = True
End If
End Sub
Private Sub Frm75_Export_Excel_3_Click()
'On Error Resume Next
Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report. Teruskan ?"
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 20
        .Columns("C").ColumnWidth = 20
        .Columns("D").ColumnWidth = 40
        .Columns("E").ColumnWidth = 20
        .Columns("F").ColumnWidth = 20
        .Columns("G").ColumnWidth = 20
        .Columns("H").ColumnWidth = 20
        .Columns("I").ColumnWidth = 20
        .Columns("J").ColumnWidth = 10
        .Columns("K").ColumnWidth = 10
        .Columns("L").ColumnWidth = 10
        .Columns("M").ColumnWidth = 10
        .Columns("N").ColumnWidth = 10
        .Columns("O").ColumnWidth = 10
        .Columns("P").ColumnWidth = 10
        .Columns("Q").ColumnWidth = 20
        
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
            LM_NAMA_HEADER = "HQ"
            
        Else
            
            LM_NAMA_HEADER = MDI_frm1.L20_Text
            
        End If
                
        '### Maklumat kedai ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 56_maklumat_kedai where cawangan='" & LM_NAMA_HEADER & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!nama_kedai) Then
                .Cells(1, 4) = rs!nama_kedai
                .Cells(1, 4).Font.Name = "Times New Roman"
            End If
            If Not IsNull(rs!no_pendaftaran) Then .Cells(2, 4) = rs!no_pendaftaran
            If Not IsNull(rs!alamat) Then .Cells(3, 4) = rs!alamat
            If Not IsNull(rs!no_tel) Then .Cells(4, 4) = rs!no_tel
            If Not IsNull(rs!no_id_gst) Then .Cells(5, 4) = rs!no_id_gst
        End If
        
        rs.Close
        Set rs = Nothing
        '### Maklumat kedai ### - End
        
        .Cells(1, 4).Font.Bold = True
        .Cells(1, 4).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 4).HorizontalAlignment = xlCenter
        Next Row
        
        .Cells(7, 1) = Frm75.L10_Text 'Header Report
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh"
        .Cells(8, 3) = "No. Resit"
        .Cells(8, 4) = "Perkara"
        .Cells(8, 5) = "Jumlah Harga SR (RM)"
        .Cells(8, 6) = "Jumlah Cukai SR (RM)"
        .Cells(8, 7) = "Jumlah Harga ZR (RM)"
        .Cells(8, 8) = "Jumlah Cukai ZR (RM)"
    
        For i = 1 To 8
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
        
        Y = 0
        For x = 1 To Frm75.MSFlexGrid3.Rows - 1
            Y = Y + 1
            .Cells(8 + Y, 1) = Y
            .Cells(8 + Y, 2) = "'" & Frm75.MSFlexGrid3.TextMatrix(x, 3) 'Tarikh
            .Cells(8 + Y, 2).HorizontalAlignment = xlCenter
            .Cells(8 + Y, 3) = "'" & Frm75.MSFlexGrid3.TextMatrix(x, 4) 'No. Resit
            .Cells(8 + Y, 3).HorizontalAlignment = xlCenter
            .Cells(8 + Y, 4) = "'" & Frm75.MSFlexGrid3.TextMatrix(x, 5) 'Perkara
            .Cells(8 + Y, 5).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 5) = Frm75.MSFlexGrid3.TextMatrix(x, 6) 'Jumlah Harga SR
            .Cells(8 + Y, 5).HorizontalAlignment = xlCenter
            .Cells(8 + Y, 6).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 6) = Frm75.MSFlexGrid3.TextMatrix(x, 7) 'Jumlah Cukai SR
            .Cells(8 + Y, 6).HorizontalAlignment = xlCenter
            .Cells(8 + Y, 7).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 7) = Frm75.MSFlexGrid3.TextMatrix(x, 8) 'Jumlah Harga ZR
            .Cells(8 + Y, 7).HorizontalAlignment = xlCenter
            .Cells(8 + Y, 8).NumberFormat = "#,##0.00"
            .Cells(8 + Y, 8) = Frm75.MSFlexGrid3.TextMatrix(x, 9) 'Jumlah Cukai ZR
            .Cells(8 + Y, 8).HorizontalAlignment = xlCenter
            
            For Col = 1 To 8
                .Cells(8 + Y, Col).Borders.LineStyle = xlContinuous
            Next Col
        Next x

        Y = Y + 2
        .Cells(8 + Y, 1) = "Jumlah Harga Barang Yang Dikenakan Cukai Standard Rated (SR) : RM " & Frm75.L14_Text
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah Kutipan Cukai Standard Rated (SR) : RM " & Frm75.L15_Text
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah Harga Barang Yang Dikenakan Cukai Zero Rated (ZR) : RM " & Frm75.L22_Text
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah Kutipan Cukai Zero Rated (ZR) : RM " & Frm75.L23_Text
        
    End With
    
    ' This makes Excel visible
    xlObject.Visible = True
    xlObject.EnableEvents = True
End If
End Sub
Private Sub Frm75_Export_Excel_4_Click()
'On Error Resume Next
'Dim xlObject As Excel.Application
'Dim xlWB As Excel.Workbook
       
'Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report. Teruskan ?"
'Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
'If Answer = vbNo Then
'    Exit Sub
'End If
'If Answer = vbYes Then
'    Set xlObject = New Excel.Application
'    Set xlWB = xlObject.Workbooks.Add
'
'    'xlObject.Visible = True
'    With xlObject.ActiveWorkbook.ActiveSheet
'        .Cells.VerticalAlignment = xlCenter
'        .Columns("A").ColumnWidth = 5
'        .Columns("B").ColumnWidth = 20
'        .Columns("C").ColumnWidth = 20
 '       .Columns("D").ColumnWidth = 60
 '       .Columns("E").ColumnWidth = 20
  '      .Columns("F").ColumnWidth = 20
  '      .Columns("G").ColumnWidth = 20
  '      .Columns("H").ColumnWidth = 10
 '       .Columns("I").ColumnWidth = 10
 '       .Columns("J").ColumnWidth = 10
'        .Columns("K").ColumnWidth = 10
'        .Columns("L").ColumnWidth = 10
'        .Columns("M").ColumnWidth = 10
'        .Columns("N").ColumnWidth = 10
'        .Columns("O").ColumnWidth = 10
'        .Columns("P").ColumnWidth = 10
'        .Columns("Q").ColumnWidth = 20
'
'        .Cells(1, 4) = "Kedai Emas Sri Harmoni"
'        .Cells(1, 4).Font.Name = "Times New Roman"
'        .Cells(2, 4) = "No Pendaftaran :  KT0302567-K"
'        .Cells(3, 4) = "355B Jalan Temenggong , 15000 Kota Bharu , Kelantan"
'        .Cells(4, 4) = "No. Telefon : +609 - 746 1093"
'        .Cells(5, 4) = vbNullString
'
'        .Cells(1, 4).Font.Bold = True
'        .Cells(1, 4).Font.Size = 30
'
'        For Row = 1 To 5
'            .Cells(Row, 4).HorizontalAlignment = xlCenter
'        Next Row
'
'        .Cells(7, 1) = Frm75.L11_Text 'Header Report
'
'        .Cells(8, 1) = "No."
'        .Cells(8, 2) = "Tarikh"
'        .Cells(8, 3) = "No. Resit"
'        .Cells(8, 4) = "Detail"
'        .Cells(8, 5) = "Jumlah (RM)"
'        .Cells(8, 6) = "GST (RM)"
'        .Cells(8, 7) = "Jumlah Keseluruhan (RM)"
'
'        For i = 1 To 7
'            .Cells(8, i).HorizontalAlignment = xlCenter
'            .Cells(8, i).Interior.ColorIndex = 15
'            .Cells(8, i).WrapText = True
'            .Cells(8, i).Borders.LineStyle = xlContinuous
'        Next i
'
'        Y = 0
'        For X = 1 To Frm75.MSFlexGrid4.Rows - 1
'            Y = Y + 1
'            .Cells(8 + Y, 1) = Y
'            .Cells(8 + Y, 2) = "'" & Frm75.MSFlexGrid4.TextMatrix(X, 2) 'Tarikh
'            .Cells(8 + Y, 2).HorizontalAlignment = xlCenter
'            .Cells(8 + Y, 3) = "'" & Frm75.MSFlexGrid4.TextMatrix(X, 3) 'No. Resit
'            .Cells(8 + Y, 3).HorizontalAlignment = xlCenter
'            .Cells(8 + Y, 4) = "'" & Frm75.MSFlexGrid4.TextMatrix(X, 4) 'Detail
'            .Cells(8 + Y, 5) = "'" & Frm75.MSFlexGrid4.TextMatrix(X, 5) 'Jumlah (RM)
'            .Cells(8 + Y, 5).HorizontalAlignment = xlCenter
'            .Cells(8 + Y, 6) = "'" & Frm75.MSFlexGrid4.TextMatrix(X, 6) 'GST (RM)
'            .Cells(8 + Y, 6).HorizontalAlignment = xlCenter
'            .Cells(8 + Y, 7) = "'" & Frm75.MSFlexGrid4.TextMatrix(X, 7) 'Jumlah Keseluruhan (RM)
'            .Cells(8 + Y, 7).HorizontalAlignment = xlCenter
'            For Col = 1 To 7
'                .Cells(8 + Y, Col).Borders.LineStyle = xlContinuous
'            Next Col
'        Next X
'
'        Y = Y + 2
'        .Cells(8 + Y, 1) = "Jumlah Bayaran GST : RM " & Frm75.L15_Text
'    End With
'
'    ' This makes Excel visible
'    xlObject.Visible = True
'    xlObject.EnableEvents = True
'End If
End Sub

Private Sub L13_Text_Change()
'on error resume next
If IsNumeric(Frm75.L13_Text) Then Frm75.L18_Text = Format(Frm75.L13_Text, "#,##0.00")
End Sub

Private Sub L15_Text_Change()
'on error resume next
If IsNumeric(Frm75.L15_Text) Then Frm75.L17_Text = Format(Frm75.L15_Text, "#,##0.00")
End Sub
Private Sub L17_Text_Change()
'on error resume next
Call frm75_kiraan_summary_gst
End Sub

Private Sub L18_Text_Change()
'on error resume next
Call frm75_kiraan_summary_gst
End Sub

Private Sub L3_Text_Click()
'on error resume next
If Frm75.Pic1.Visible = False Then
    Call Frm75_Initial_Setting
    
    Frm75.Pic1.Visible = True
Else
    Frm75.Pic1.Visible = False
End If
End Sub
Private Sub L4_Text_Click()
'on error resume next
If Frm75.Pic2.Visible = False Then
    Call Frm75_Initial_Setting

    Frm75.Pic2.Visible = True
Else
    Frm75.Pic2.Visible = False
End If
End Sub
Private Sub L5_Text_Click()
'on error resume next
If Frm75.Pic3.Visible = False Then
    Call Frm75_Initial_Setting
    
    Frm75.Pic3.Visible = True
Else
    Frm75.Pic3.Visible = False
End If
End Sub

Private Sub MSFlexGrid1_DblClick()
'On Error Resume Next
'If Frm75.MSFlexGrid1 <> vbNullString Then
'    PopupMenu Frm75_Menu1
'Else
'    MsgBox "Tiada Data.", vbExclamation, "Info"
'End If
End Sub
Private Sub MSFlexGrid3_DblClick()
'On Error Resume Next
'If Frm75.MSFlexGrid3 <> vbNullString Then
'    PopupMenu Frm75_Menu3
'Else
'    MsgBox "Tiada Data.", vbExclamation, "Info"
'End If
End Sub
Private Sub MSFlexGrid4_DblClick()
'On Error Resume Next
'If Frm75.MSFlexGrid4 <> vbNullString Then
'    PopupMenu Frm75_Menu4
'Else
'    MsgBox "Tiada Data.", vbExclamation, "Info"
'End If
End Sub

