VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm57 
   Caption         =   "Inventori Mengikut Dulang"
   ClientHeight    =   12615
   ClientLeft      =   120
   ClientTop       =   -10635
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
   Icon            =   "Frm57.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12615
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic3 
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
      Height          =   3735
      Left            =   8640
      ScaleHeight     =   3735
      ScaleWidth      =   6255
      TabIndex        =   12
      Top             =   -480
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton CMD4 
         Caption         =   "Inventori Selesai"
         Height          =   375
         Left            =   2280
         MouseIcon       =   "Frm57.frx":0ECA
         MousePointer    =   99  'Custom
         TabIndex        =   69
         Top             =   3120
         Width           =   2055
      End
      Begin VB.CommandButton CMD3 
         Caption         =   "Carian"
         Height          =   375
         Left            =   2280
         MouseIcon       =   "Frm57.frx":11D4
         MousePointer    =   99  'Custom
         TabIndex        =   68
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Report"
         Height          =   375
         Left            =   -3480
         MouseIcon       =   "Frm57.frx":14DE
         MousePointer    =   99  'Custom
         TabIndex        =   67
         ToolTipText     =   "Senarai barang yang telah dimasukkan ke dalam senarai belian"
         Top             =   -3840
         Width           =   2535
      End
      Begin VB.CheckBox CB1 
         Caption         =   "Scanner Mode"
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
         Left            =   300
         TabIndex        =   16
         Top             =   180
         Width           =   200
      End
      Begin VB.TextBox TB1 
         Alignment       =   2  'Center
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
         IMEMode         =   3  'DISABLE
         Left            =   2160
         TabIndex        =   15
         Top             =   840
         Width           =   3405
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   3
         Height          =   1215
         Left            =   480
         Top             =   1800
         Width           =   5415
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Status carian barang."
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   720
         TabIndex        =   59
         Top             =   1920
         Width           =   2280
      End
      Begin VB.Label L6_Text 
         BackStyle       =   0  'Transparent
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
         Left            =   5520
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label L5_Text 
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
         Left            =   720
         TabIndex        =   19
         Top             =   2520
         Width           =   4995
      End
      Begin VB.Label L4_Text 
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
         Left            =   720
         TabIndex        =   18
         Top             =   2200
         Width           =   4995
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Scanner Mode"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   17
         Top             =   150
         Width           =   1680
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Siri Produk  :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   600
         TabIndex        =   14
         Top             =   900
         Width           =   2280
      End
      Begin VB.Label L3_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila scan setiap item dari dulang."
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         TabIndex        =   13
         Top             =   480
         Width           =   9240
      End
   End
   Begin VB.PictureBox Pic7 
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
      Height          =   11655
      Left            =   10920
      ScaleHeight     =   11655
      ScaleWidth      =   21375
      TabIndex        =   41
      Top             =   1200
      Visible         =   0   'False
      Width           =   21375
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid5 
         Height          =   10755
         Left            =   120
         TabIndex        =   63
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   360
         Width           =   21135
         _ExtentX        =   37280
         _ExtentY        =   18971
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
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Berat   :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3120
         TabIndex        =   48
         Top             =   11100
         Width           =   1800
      End
      Begin VB.Label L23_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4560
         TabIndex        =   47
         Top             =   11100
         Width           =   2280
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Modal  :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7320
         TabIndex        =   46
         Top             =   11100
         Width           =   1800
      End
      Begin VB.Label L22_Text 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   44
         Top             =   40
         Width           =   12720
      End
      Begin VB.Label L25_Text 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         TabIndex        =   43
         Top             =   11100
         Width           =   2280
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Kuantiti   :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   42
         Top             =   11100
         Width           =   1200
      End
      Begin VB.Label L24_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8760
         TabIndex        =   45
         Top             =   11100
         Width           =   2280
      End
   End
   Begin VB.PictureBox Pic2 
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
      Height          =   11655
      Left            =   9600
      ScaleHeight     =   11655
      ScaleWidth      =   21375
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   21375
      Begin VB.CommandButton CMD18 
         BackColor       =   &H00FFFFFF&
         Height          =   700
         Left            =   19800
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm57.frx":17E8
         MousePointer    =   99  'Custom
         Picture         =   "Frm57.frx":1AF2
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Paparan seterusnya"
         Top             =   10440
         Width           =   1200
      End
      Begin VB.CommandButton CMD17 
         BackColor       =   &H00FFFFFF&
         Height          =   700
         Left            =   18480
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm57.frx":2418
         MousePointer    =   99  'Custom
         Picture         =   "Frm57.frx":2722
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Paparan sebelumnya"
         Top             =   10440
         Width           =   1200
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   10035
         Left            =   120
         TabIndex        =   61
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   360
         Width           =   20895
         _ExtentX        =   36856
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
      Begin VB.Label L40_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L40_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   17040
         TabIndex        =   75
         Top             =   10920
         Width           =   735
      End
      Begin VB.Label L42_Text 
         Caption         =   "L42_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   12720
         TabIndex        =   74
         Top             =   10560
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L43_Text 
         Caption         =   "L43_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   13800
         TabIndex        =   73
         Top             =   10560
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L41_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L41_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   18000
         TabIndex        =   72
         Top             =   10920
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Berat   :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2880
         TabIndex        =   11
         Top             =   10500
         Width           =   1800
      End
      Begin VB.Label L10_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4440
         TabIndex        =   10
         Top             =   10500
         Width           =   1680
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Modal  :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6240
         TabIndex        =   9
         Top             =   10500
         Width           =   1800
      End
      Begin VB.Label L11_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7800
         TabIndex        =   8
         Top             =   10500
         Width           =   1680
      End
      Begin VB.Label L9_Text 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   40
         Width           =   14880
      End
      Begin VB.Label L32_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   10500
         Width           =   1320
      End
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   "Kuantiti   :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   10500
         Width           =   1080
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Paparan Muka  :          / "
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   15840
         TabIndex        =   76
         Top             =   10920
         Width           =   2295
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Tmr2 
      Interval        =   100
      Left            =   0
      Top             =   960
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
      Height          =   1575
      Left            =   2400
      ScaleHeight     =   1575
      ScaleWidth      =   7215
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CommandButton CMD2 
         Caption         =   "Report"
         Height          =   375
         Left            =   2160
         MouseIcon       =   "Frm57.frx":3061
         MousePointer    =   99  'Custom
         TabIndex        =   66
         Top             =   960
         Width           =   2535
      End
      Begin VB.ComboBox CBB1 
         Height          =   360
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   450
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila buat pilihan DULANG bagi mengetahui status inventori bagi setiap dulang."
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   9240
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Pilihan Dulang  :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   600
         TabIndex        =   2
         Top             =   480
         Width           =   2040
      End
   End
   Begin VB.Timer Tmr1 
      Interval        =   10
      Left            =   0
      Top             =   240
   End
   Begin VB.PictureBox Pic6 
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
      Height          =   11655
      Left            =   9240
      ScaleHeight     =   11655
      ScaleWidth      =   21375
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   21375
      Begin VB.CommandButton CMD9 
         BackColor       =   &H00FFFFFF&
         Height          =   700
         Left            =   10320
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm57.frx":336B
         MousePointer    =   99  'Custom
         Picture         =   "Frm57.frx":3675
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   "Paparan sebelumnya"
         Top             =   10440
         Width           =   1200
      End
      Begin VB.CommandButton CMD10 
         BackColor       =   &H00FFFFFF&
         Height          =   700
         Left            =   11640
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm57.frx":3FB4
         MousePointer    =   99  'Custom
         Picture         =   "Frm57.frx":42BE
         Style           =   1  'Graphical
         TabIndex        =   91
         ToolTipText     =   "Paparan seterusnya"
         Top             =   10440
         Width           =   1200
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
         Height          =   10035
         Left            =   120
         TabIndex        =   64
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   360
         Width           =   12735
         _ExtentX        =   22463
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
      Begin VB.Label L53_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L53_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9960
         TabIndex        =   96
         Top             =   10920
         Width           =   2295
      End
      Begin VB.Label L55_Text 
         Caption         =   "L55_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   6960
         TabIndex        =   95
         Top             =   10560
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L54_Text 
         Caption         =   "L54_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   5880
         TabIndex        =   94
         Top             =   10560
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L52_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L52_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9000
         TabIndex        =   93
         Top             =   10920
         Width           =   735
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Kuantiti   :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   40
         Top             =   10500
         Width           =   1200
      End
      Begin VB.Label L26_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         TabIndex        =   39
         Top             =   10500
         Width           =   2280
      End
      Begin VB.Label L20_Text 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   38
         Top             =   40
         Width           =   12720
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Paparan Muka  :          / "
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7800
         TabIndex        =   97
         Top             =   10920
         Width           =   2295
      End
   End
   Begin VB.PictureBox Pic4 
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
      Height          =   11655
      Left            =   8880
      ScaleHeight     =   11655
      ScaleWidth      =   23535
      TabIndex        =   21
      Top             =   -480
      Visible         =   0   'False
      Width           =   23535
      Begin VB.CommandButton CMD5 
         BackColor       =   &H00FFFFFF&
         Height          =   700
         Left            =   18480
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm57.frx":4BE4
         MousePointer    =   99  'Custom
         Picture         =   "Frm57.frx":4EEE
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Paparan sebelumnya"
         Top             =   10440
         Width           =   1200
      End
      Begin VB.CommandButton CMD6 
         BackColor       =   &H00FFFFFF&
         Height          =   700
         Left            =   19800
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm57.frx":582D
         MousePointer    =   99  'Custom
         Picture         =   "Frm57.frx":5B37
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Paparan seterusnya"
         Top             =   10440
         Width           =   1200
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   10035
         Left            =   120
         TabIndex        =   65
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   360
         Width           =   20895
         _ExtentX        =   36856
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
      Begin VB.Label L45_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L45_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   18000
         TabIndex        =   82
         Top             =   10920
         Width           =   2295
      End
      Begin VB.Label L47_Text 
         Caption         =   "L47_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   15480
         TabIndex        =   81
         Top             =   10680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L46_Text 
         Caption         =   "L46_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   14400
         TabIndex        =   80
         Top             =   10680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L44_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L44_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   17040
         TabIndex        =   79
         Top             =   10920
         Width           =   735
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Kuantiti   :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         TabIndex        =   28
         Top             =   10500
         Width           =   1200
      End
      Begin VB.Label L15_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1320
         TabIndex        =   27
         Top             =   10500
         Width           =   2280
      End
      Begin VB.Label L12_Text 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   26
         Top             =   40
         Width           =   12720
      End
      Begin VB.Label L14_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7800
         TabIndex        =   25
         Top             =   10500
         Width           =   2280
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Modal  :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6360
         TabIndex        =   24
         Top             =   10500
         Width           =   1800
      End
      Begin VB.Label L13_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4560
         TabIndex        =   23
         Top             =   10500
         Width           =   2280
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Berat   :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3120
         TabIndex        =   22
         Top             =   10500
         Width           =   1800
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Paparan Muka  :          / "
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   15840
         TabIndex        =   83
         Top             =   10920
         Width           =   2295
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
      Height          =   11655
      Left            =   9720
      ScaleHeight     =   11655
      ScaleWidth      =   21375
      TabIndex        =   29
      Top             =   -480
      Visible         =   0   'False
      Width           =   21375
      Begin VB.CommandButton CMD8 
         BackColor       =   &H00FFFFFF&
         Height          =   700
         Left            =   19800
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm57.frx":645D
         MousePointer    =   99  'Custom
         Picture         =   "Frm57.frx":6767
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Paparan seterusnya"
         Top             =   10440
         Width           =   1200
      End
      Begin VB.CommandButton CDM7 
         BackColor       =   &H00FFFFFF&
         Height          =   700
         Left            =   18480
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm57.frx":708D
         MousePointer    =   99  'Custom
         Picture         =   "Frm57.frx":7397
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Paparan sebelumnya"
         Top             =   10440
         Width           =   1200
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   10035
         Left            =   120
         TabIndex        =   62
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   360
         Width           =   20895
         _ExtentX        =   36856
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
      Begin VB.Label L48_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L48_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   16920
         TabIndex        =   89
         Top             =   10920
         Width           =   735
      End
      Begin VB.Label L50_Text 
         Caption         =   "L50_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   14280
         TabIndex        =   88
         Top             =   10800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L51_Text 
         Caption         =   "L51_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   15360
         TabIndex        =   87
         Top             =   10800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label L49_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L49_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   17880
         TabIndex        =   86
         Top             =   10920
         Width           =   2295
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Berat   :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2640
         TabIndex        =   36
         Top             =   10500
         Width           =   1800
      End
      Begin VB.Label L17_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4080
         TabIndex        =   35
         Top             =   10500
         Width           =   2280
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Modal  :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6240
         TabIndex        =   34
         Top             =   10500
         Width           =   1800
      End
      Begin VB.Label L18_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7680
         TabIndex        =   33
         Top             =   10500
         Width           =   2280
      End
      Begin VB.Label L16_Text 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   32
         Top             =   40
         Width           =   12720
      End
      Begin VB.Label L19_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         TabIndex        =   31
         Top             =   10500
         Width           =   2280
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Kuantiti   :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   30
         Top             =   10500
         Width           =   1200
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Paparan Muka  :          / "
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   15720
         TabIndex        =   90
         Top             =   10920
         Width           =   2295
      End
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   $"Frm57.frx":7CD6
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
      Height          =   645
      Left            =   240
      TabIndex        =   98
      Top             =   5760
      Width           =   9240
   End
   Begin VB.Label L39_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "Sila tekan F2 untuk scan barang dari dulang ."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   60
      Top             =   6480
      Width           =   9975
   End
   Begin VB.Label L37_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lain-lain Status"
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
      Left            =   10440
      MouseIcon       =   "Frm57.frx":7D85
      MousePointer    =   99  'Custom
      TabIndex        =   58
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label L36_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Barang Diluar Kawalan"
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
      Left            =   8160
      MouseIcon       =   "Frm57.frx":808F
      MousePointer    =   99  'Custom
      TabIndex        =   57
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label L35_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Barang Terkawal"
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
      Left            =   6240
      MouseIcon       =   "Frm57.frx":8399
      MousePointer    =   99  'Custom
      TabIndex        =   56
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label L34_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Barang Yang Sepatut Berada Di Dalam Dulang"
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
      MouseIcon       =   "Frm57.frx":86A3
      MousePointer    =   99  'Custom
      TabIndex        =   55
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label L33_Text 
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
      MouseIcon       =   "Frm57.frx":89AD
      MousePointer    =   99  'Custom
      TabIndex        =   54
      Top             =   0
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   3615
      Left            =   120
      Top             =   2040
      Width           =   9375
   End
   Begin VB.Line Line12 
      X1              =   4800
      X2              =   4560
      Y1              =   5160
      Y2              =   4800
   End
   Begin VB.Line Line11 
      X1              =   4800
      X2              =   5040
      Y1              =   5160
      Y2              =   4800
   End
   Begin VB.Line Line10 
      X1              =   4800
      X2              =   4800
      Y1              =   4680
      Y2              =   5160
   End
   Begin VB.Line Line6 
      X1              =   4800
      X2              =   4560
      Y1              =   4320
      Y2              =   3960
   End
   Begin VB.Line Line5 
      X1              =   4800
      X2              =   5040
      Y1              =   4320
      Y2              =   3960
   End
   Begin VB.Line Line4 
      X1              =   4800
      X2              =   4800
      Y1              =   3840
      Y2              =   4320
   End
   Begin VB.Line Line3 
      X1              =   4800
      X2              =   4560
      Y1              =   3600
      Y2              =   3240
   End
   Begin VB.Line Line2 
      X1              =   4800
      X2              =   5040
      Y1              =   3600
      Y2              =   3240
   End
   Begin VB.Line Line1 
      X1              =   4800
      X2              =   4800
      Y1              =   3120
      Y2              =   3600
   End
   Begin VB.Shape Shape1 
      Height          =   420
      Left            =   360
      Top             =   2280
      Width           =   8895
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sistem akan keluarkan report inventori bagi dulang yang dipilih."
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
      TabIndex        =   53
      Top             =   5280
      Width           =   9240
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Setelah semua selesai klik ""Inventori Selesai"""
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
      TabIndex        =   52
      Top             =   4440
      Width           =   9240
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Scan setiap item di dalam dulang"
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
      TabIndex        =   51
      Top             =   3600
      Width           =   9240
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Buat Pilihan Dulang"
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
      TabIndex        =   50
      Top             =   2880
      Width           =   9240
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Carta aliran sistem inventori dulang"
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
      Left            =   0
      TabIndex        =   49
      Top             =   2280
      Width           =   9240
   End
   Begin VB.Menu Frm57_PM_Menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm57_PM_Excel1 
         Caption         =   "Export Excel Report"
      End
   End
   Begin VB.Menu Frm57_PM_Menu2 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm57_PM_Excel2 
         Caption         =   "Export Excel Report"
      End
   End
   Begin VB.Menu Frm57_PM_Menu3 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm57_PM_Excel3 
         Caption         =   "Export Excel Report"
      End
   End
   Begin VB.Menu Frm57_PM_Menu4 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm57_PM_Excel4 
         Caption         =   "Export Excel Report"
      End
   End
   Begin VB.Menu Frm57_PM_Scan 
      Caption         =   "Scan Mode F2"
      Begin VB.Menu Frm57_SM_scan_mode 
         Caption         =   "Scan Mode"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "Frm57"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB1_Click()
'On Error Resume Next
'If Frm57.CB1 = 1 Then
'    Frm57.TB1.SetFocus
'End If
End Sub
Private Sub CMD1_Click()
'On Error Resume Next
Call Frm57_M_Clear
Frm57.Pic1.Visible = True
End Sub
Private Sub CDM7_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm57_M_Inventori_Luar_Kawalan_header
Call Frm57_M_Inventori_Luar_Kawalan
End Sub
Private Sub CMD10_Click()
'on error resume next
Dim Frm57_LM_CURR_PAGE As Double
Dim Frm57_LM_TOTAL_PAGE As Double

Frm57_LM_CURR_PAGE = 0
Frm57_LM_TOTAL_PAGE = 0

If Frm57.L52_Text <> vbNullString And IsNumeric(Frm57.L52_Text) Then
    If Frm57.L53_Text <> vbNullString And IsNumeric(Frm57.L53_Text) Then
        Frm57_LM_CURR_PAGE = Frm57.L52_Text
        Frm57_LM_TOTAL_PAGE = Frm57.L53_Text
        
        If Frm57_LM_CURR_PAGE < Frm57_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm57_M_Inventori_Lain_header
            Call Frm57_M_Inventori_Lain
            
        End If
    End If
End If
End Sub
Private Sub CMD17_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm57_M_RekodInventori_Stok_header
Call Frm57_M_RekodInventori_Stok
End Sub
Private Sub CMD18_Click()
'on error resume next
Dim Frm57_LM_CURR_PAGE As Double
Dim Frm57_LM_TOTAL_PAGE As Double

Frm57_LM_CURR_PAGE = 0
Frm57_LM_TOTAL_PAGE = 0

If Frm57.L40_Text <> vbNullString And IsNumeric(Frm57.L40_Text) Then
    If Frm57.L41_Text <> vbNullString And IsNumeric(Frm57.L41_Text) Then
        Frm57_LM_CURR_PAGE = Frm57.L40_Text
        Frm57_LM_TOTAL_PAGE = Frm57.L41_Text
        
        If Frm57_LM_CURR_PAGE < Frm57_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm57_M_RekodInventori_Stok_header
            Call Frm57_M_RekodInventori_Stok
            
        End If
    End If
End If
End Sub
Private Sub CMD2_Click()
'On Error Resume Next
If Frm57.CBB1 = vbNullString Then
    MsgBox "Sila pilih DULANG.", vbInformation, "Info"
Else
    Note = "Sistem akan mengambil masa untuk mengeluarkan report." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan ?"
            
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
        Frm57.L4_Text = vbNullString
        Frm57.L5_Text = vbNullString
        Frm57.Pic3.Visible = False
        Call Frm57_M_Inventory
        
        
        GM_NEXT_PREV = 0
        Frm57.L42_Text = -1 'Titik Pencarian Data
        Frm57.L43_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
        Frm57.L40_Text = 0 'Paparan Page ke-xxx
        
        Frm57.L46_Text = -1 'Titik Pencarian Data
        Frm57.L47_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
        Frm57.L44_Text = 0 'Paparan Page ke-xxx
        
        Frm57.L50_Text = -1 'Titik Pencarian Data
        Frm57.L51_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
        Frm57.L48_Text = 0 'Paparan Page ke-xxx
        
        Frm57.L54_Text = -1 'Titik Pencarian Data
        Frm57.L55_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
        Frm57.L52_Text = 0 'Paparan Page ke-xxx
        
        Frm57.L34_Text.Visible = False
        Frm57.L35_Text.Visible = False
        Frm57.L36_Text.Visible = False
        Frm57.L37_Text.Visible = False
    
    End If
End If
End Sub
Private Sub CMD3_Click()
'On Error Resume Next
If Frm57.TB1 = vbNullString Then
    MsgBox "Sila masukkan No. Siri Produk.", vbInformation, "Info"
Else
    Call Frm57_M_Carian
End If
End Sub
Private Sub CMD4_Click()
'On Error Resume Next
Note = "Sistem Akan Mengambil Masa Untuk Menganalisa Data Inventori Yang Telah Dimasukkan." & vbCrLf & _
    "Teruskan ?"
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then

    Frm57.Pic1.Visible = False
    Frm57.Pic3.Visible = False

    Call Frm57_M_RekodInventori_Stok_header
    Call Frm57_M_RekodInventori_Stok
    Call Frm57_M_Inventori_Kawalan_header
    Call Frm57_M_Inventori_Kawalan
    Call Frm57_M_Inventori_Luar_Kawalan_header
    Call Frm57_M_Inventori_Luar_Kawalan
    Call Frm57_M_Inventori_Lain_header
    Call Frm57_M_Inventori_Lain
    Call Frm57_initial_setting
    
    Frm57.Pic2.Visible = True
    
    Frm57.L34_Text.Visible = True
    Frm57.L35_Text.Visible = True
    Frm57.L36_Text.Visible = True
    Frm57.L37_Text.Visible = True

End If
End Sub
Private Sub CMD5_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm57_M_Inventori_Kawalan_header
Call Frm57_M_Inventori_Kawalan
End Sub
Private Sub CMD6_Click()
'on error resume next
Dim Frm57_LM_CURR_PAGE As Double
Dim Frm57_LM_TOTAL_PAGE As Double

Frm57_LM_CURR_PAGE = 0
Frm57_LM_TOTAL_PAGE = 0

If Frm57.L44_Text <> vbNullString And IsNumeric(Frm57.L44_Text) Then
    If Frm57.L45_Text <> vbNullString And IsNumeric(Frm57.L45_Text) Then
        Frm57_LM_CURR_PAGE = Frm57.L44_Text
        Frm57_LM_TOTAL_PAGE = Frm57.L45_Text
        
        If Frm57_LM_CURR_PAGE < Frm57_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm57_M_Inventori_Kawalan_header
            Call Frm57_M_Inventori_Kawalan
            
        End If
    End If
End If
End Sub
Private Sub CMD8_Click()
'on error resume next
Dim Frm57_LM_CURR_PAGE As Double
Dim Frm57_LM_TOTAL_PAGE As Double

Frm57_LM_CURR_PAGE = 0
Frm57_LM_TOTAL_PAGE = 0

If Frm57.L48_Text <> vbNullString And IsNumeric(Frm57.L48_Text) Then
    If Frm57.L49_Text <> vbNullString And IsNumeric(Frm57.L49_Text) Then
        Frm57_LM_CURR_PAGE = Frm57.L48_Text
        Frm57_LM_TOTAL_PAGE = Frm57.L49_Text
        
        If Frm57_LM_CURR_PAGE < Frm57_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm57_M_Inventori_Luar_Kawalan_header
            Call Frm57_M_Inventori_Luar_Kawalan
            
        End If
    End If
End If
End Sub
Private Sub CMD9_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm57_M_Inventori_Lain_header
Call Frm57_M_Inventori_Lain
End Sub
Private Sub Form_Load()
'On Error Resume Next
Frm57.L4_Text.BackStyle = 0
Frm57.L5_Text.BackStyle = 0
End Sub
Private Sub Frm57_PM_Excel1_Click()
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
        .Columns("B").ColumnWidth = 15
        .Columns("C").ColumnWidth = 15
        .Columns("D").ColumnWidth = 35
        .Columns("E").ColumnWidth = 40
        .Columns("F").ColumnWidth = 15
        .Columns("G").ColumnWidth = 15
        .Columns("H").ColumnWidth = 15
        .Columns("I").ColumnWidth = 15
        .Columns("J").ColumnWidth = 15
        .Columns("K").ColumnWidth = 10
        .Columns("L").ColumnWidth = 10
        .Columns("M").ColumnWidth = 10
        .Columns("N").ColumnWidth = 10
        .Columns("O").ColumnWidth = 10
        .Columns("P").ColumnWidth = 10
        .Columns("Q").ColumnWidth = 10
        .Columns("R").ColumnWidth = 10
        .Columns("S").ColumnWidth = 10
        .Columns("T").ColumnWidth = 10
        .Columns("U").ColumnWidth = 10
        
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
                .Cells(1, 6) = rs!nama_kedai
                .Cells(1, 6).Font.Name = "Times New Roman"
            End If
            If Not IsNull(rs!no_pendaftaran) Then .Cells(2, 6) = rs!no_pendaftaran
            If Not IsNull(rs!alamat) Then .Cells(3, 6) = rs!alamat
            If Not IsNull(rs!no_tel) Then .Cells(4, 6) = rs!no_tel
            If Not IsNull(rs!no_id_gst) Then .Cells(5, 6) = rs!no_id_gst
        End If
        
        rs.Close
        Set rs = Nothing
        '### Maklumat kedai ### - End
        
        .Cells(1, 6).Font.Bold = True
        .Cells(1, 6).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 6).HorizontalAlignment = xlCenter
        Next Row
        
        .Cells(7, 1) = Frm57.L9_Text 'Header Report
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Belian"
        .Cells(8, 3) = "No. Siri Produk"
        .Cells(8, 4) = "Nama Produk"
        .Cells(8, 5) = "Supplier"
        .Cells(8, 6) = "Purity"
        .Cells(8, 7) = "Berat (g)"
        .Cells(8, 8) = "Kos Per Gram (RM/g)"
        .Cells(8, 9) = "Upah (RM)"
        .Cells(8, 10) = "Harga Belian (RM)"
        .Cells(8, 11) = "Dulang"
        .Cells(8, 12) = "Panjang"
        .Cells(8, 13) = "Lebar"
        .Cells(8, 14) = "Dia"
        .Cells(8, 15) = "Saiz"
    
        For i = 1 To 15
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from inventory", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
            x = x + 1
            .Cells(8 + x, 1) = x
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            If Not IsNull(rs!tarikh_belian) Then .Cells(8 + x, 2) = "'" & rs!tarikh_belian 'Tarikh Belian
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            If Not IsNull(rs!no_siri) Then .Cells(8 + x, 3) = rs!no_siri 'No. Siri Produk
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            If Not IsNull(rs!nama_produk) Then .Cells(8 + x, 4) = rs!nama_produk 'Nama Produk
            If Not IsNull(rs!supplier) Then .Cells(8 + x, 5) = rs!supplier 'Nama Supplier
            If Not IsNull(rs!purity) Then .Cells(8 + x, 6) = rs!purity 'Purity
            .Cells(8 + x, 6).HorizontalAlignment = xlCenter
            If Not IsNull(rs!Berat) Then .Cells(8 + x, 7) = Format(rs!Berat, "#,##0.00") 'Berat Jualan
            .Cells(8 + x, 7).HorizontalAlignment = xlCenter
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            If Not IsNull(rs!KOSPERGRAM) Then .Cells(8 + x, 8) = Format(rs!KOSPERGRAM, "#,##0.00") 'Harga Belian Per Gram
            .Cells(8 + x, 8).HorizontalAlignment = xlCenter
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            If Not IsNull(rs!UPAH) Then .Cells(8 + x, 9) = Format(rs!UPAH, "#,##0.00") 'Upah
            .Cells(8 + x, 9).HorizontalAlignment = xlCenter
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            If Not IsNull(rs!harga_belian) Then .Cells(8 + x, 10) = Format(rs!harga_belian, "#,##0.00") 'Kos Belian Item
            .Cells(8 + x, 10).HorizontalAlignment = xlCenter
            .Cells(8 + x, 10).NumberFormat = "#,##0.00"
            If Not IsNull(rs!dulang) Then .Cells(8 + x, 11) = rs!dulang 'Dulang
            .Cells(8 + x, 11).HorizontalAlignment = xlCenter
            If Not IsNull(rs!panjang) Then .Cells(8 + x, 12) = rs!panjang 'Panjang
            .Cells(8 + x, 12).HorizontalAlignment = xlCenter
            If Not IsNull(rs!lebar) Then .Cells(8 + x, 13) = rs!lebar 'Lebar
            .Cells(8 + x, 13).HorizontalAlignment = xlCenter
            If Not IsNull(rs!dia) Then .Cells(8 + x, 14) = rs!dia 'Dia
            .Cells(8 + x, 14).HorizontalAlignment = xlCenter
            If Not IsNull(rs!Size) Then .Cells(8 + x, 15) = rs!Size 'Saiz
            .Cells(8 + x, 15).HorizontalAlignment = xlCenter
            For Col = 1 To 15
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
    
        x = x + 2
        .Cells(8 + x, 1) = "Kuantiti : " & Frm57.L32_Text
        .Cells(8 + x, 4) = "Jumlah Berat : " & Frm57.L10_Text
        .Cells(8 + x, 6) = "Jumlah Modal : " & Frm57.L11_Text
    End With
    
    ' This makes Excel visible
    xlObject.Visible = True
    xlObject.EnableEvents = True
End If
End Sub
Private Sub Frm57_PM_Excel2_Click()
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
        .Columns("B").ColumnWidth = 15
        .Columns("C").ColumnWidth = 15
        .Columns("D").ColumnWidth = 35
        .Columns("E").ColumnWidth = 40
        .Columns("F").ColumnWidth = 15
        .Columns("G").ColumnWidth = 15
        .Columns("H").ColumnWidth = 15
        .Columns("I").ColumnWidth = 15
        .Columns("J").ColumnWidth = 15
        .Columns("K").ColumnWidth = 10
        .Columns("L").ColumnWidth = 10
        .Columns("M").ColumnWidth = 10
        .Columns("N").ColumnWidth = 10
        .Columns("O").ColumnWidth = 10
        .Columns("P").ColumnWidth = 10
        .Columns("Q").ColumnWidth = 10
        .Columns("R").ColumnWidth = 10
        .Columns("S").ColumnWidth = 10
        .Columns("T").ColumnWidth = 10
        .Columns("U").ColumnWidth = 10
        
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
                .Cells(1, 6) = rs!nama_kedai
                .Cells(1, 6).Font.Name = "Times New Roman"
            End If
            If Not IsNull(rs!no_pendaftaran) Then .Cells(2, 6) = rs!no_pendaftaran
            If Not IsNull(rs!alamat) Then .Cells(3, 6) = rs!alamat
            If Not IsNull(rs!no_tel) Then .Cells(4, 6) = rs!no_tel
            If Not IsNull(rs!no_id_gst) Then .Cells(5, 6) = rs!no_id_gst
        End If
        
        rs.Close
        Set rs = Nothing
        '### Maklumat kedai ### - End
        
        .Cells(1, 6).Font.Bold = True
        .Cells(1, 6).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 6).HorizontalAlignment = xlCenter
        Next Row
        
        .Cells(7, 1) = Frm57.L12_Text 'Header Report
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Belian"
        .Cells(8, 3) = "No. Siri Produk"
        .Cells(8, 4) = "Nama Produk"
        .Cells(8, 5) = "Supplier"
        .Cells(8, 6) = "Purity"
        .Cells(8, 7) = "Berat (g)"
        .Cells(8, 8) = "Kos Per Gram (RM/g)"
        .Cells(8, 9) = "Upah (RM)"
        .Cells(8, 10) = "Harga Belian (RM)"
        .Cells(8, 11) = "Dulang"
        .Cells(8, 12) = "Panjang"
        .Cells(8, 13) = "Lebar"
        .Cells(8, 14) = "Dia"
        .Cells(8, 15) = "Saiz"
    
        For i = 1 To 15
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from inventory where status='" & "1" & "'", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
            x = x + 1
            .Cells(8 + x, 1) = x
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            If Not IsNull(rs!tarikh_belian) Then .Cells(8 + x, 2) = "'" & rs!tarikh_belian 'Tarikh Belian
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            If Not IsNull(rs!no_siri) Then .Cells(8 + x, 3) = rs!no_siri 'No. Siri Produk
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            If Not IsNull(rs!nama_produk) Then .Cells(8 + x, 4) = rs!nama_produk 'Nama Produk
            If Not IsNull(rs!supplier) Then .Cells(8 + x, 5) = rs!supplier 'Nama Supplier
            If Not IsNull(rs!purity) Then .Cells(8 + x, 6) = rs!purity 'Purity
            .Cells(8 + x, 6).HorizontalAlignment = xlCenter
            If Not IsNull(rs!Berat) Then .Cells(8 + x, 7) = Format(rs!Berat, "#,##0.00") 'Berat Jualan
            .Cells(8 + x, 7).HorizontalAlignment = xlCenter
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            If Not IsNull(rs!KOSPERGRAM) Then .Cells(8 + x, 8) = Format(rs!KOSPERGRAM, "#,##0.00") 'Harga Belian Per Gram
            .Cells(8 + x, 8).HorizontalAlignment = xlCenter
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            If Not IsNull(rs!UPAH) Then .Cells(8 + x, 9) = Format(rs!UPAH, "#,##0.00") 'Upah
            .Cells(8 + x, 9).HorizontalAlignment = xlCenter
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            If Not IsNull(rs!harga_belian) Then .Cells(8 + x, 10) = Format(rs!harga_belian, "#,##0.00") 'Kos Belian Item
            .Cells(8 + x, 10).HorizontalAlignment = xlCenter
            .Cells(8 + x, 10).NumberFormat = "#,##0.00"
            If Not IsNull(rs!dulang) Then .Cells(8 + x, 11) = rs!dulang 'Dulang
            .Cells(8 + x, 11).HorizontalAlignment = xlCenter
            If Not IsNull(rs!panjang) Then .Cells(8 + x, 12) = rs!panjang 'Panjang
            .Cells(8 + x, 12).HorizontalAlignment = xlCenter
            If Not IsNull(rs!lebar) Then .Cells(8 + x, 13) = rs!lebar 'Lebar
            .Cells(8 + x, 13).HorizontalAlignment = xlCenter
            If Not IsNull(rs!dia) Then .Cells(8 + x, 14) = rs!dia 'Dia
            .Cells(8 + x, 14).HorizontalAlignment = xlCenter
            If Not IsNull(rs!Size) Then .Cells(8 + x, 15) = rs!Size 'Saiz
            .Cells(8 + x, 15).HorizontalAlignment = xlCenter
            For Col = 1 To 15
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
    
        x = x + 2
        .Cells(8 + x, 1) = "Kuantiti : " & Frm57.L15_Text
        .Cells(8 + x, 4) = "Jumlah Berat : " & Frm57.L13_Text
        .Cells(8 + x, 6) = "Jumlah Modal : " & Frm57.L14_Text
    End With
    
    ' This makes Excel visible
    xlObject.Visible = True
    xlObject.EnableEvents = True
End If
End Sub
Private Sub Frm57_PM_Excel3_Click()
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
        .Columns("B").ColumnWidth = 15
        .Columns("C").ColumnWidth = 15
        .Columns("D").ColumnWidth = 35
        .Columns("E").ColumnWidth = 40
        .Columns("F").ColumnWidth = 15
        .Columns("G").ColumnWidth = 15
        .Columns("H").ColumnWidth = 15
        .Columns("I").ColumnWidth = 15
        .Columns("J").ColumnWidth = 15
        .Columns("K").ColumnWidth = 10
        .Columns("L").ColumnWidth = 10
        .Columns("M").ColumnWidth = 10
        .Columns("N").ColumnWidth = 10
        .Columns("O").ColumnWidth = 10
        .Columns("P").ColumnWidth = 10
        .Columns("Q").ColumnWidth = 10
        .Columns("R").ColumnWidth = 10
        .Columns("S").ColumnWidth = 10
        .Columns("T").ColumnWidth = 10
        .Columns("U").ColumnWidth = 10
        
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
                .Cells(1, 6) = rs!nama_kedai
                .Cells(1, 6).Font.Name = "Times New Roman"
            End If
            If Not IsNull(rs!no_pendaftaran) Then .Cells(2, 6) = rs!no_pendaftaran
            If Not IsNull(rs!alamat) Then .Cells(3, 6) = rs!alamat
            If Not IsNull(rs!no_tel) Then .Cells(4, 6) = rs!no_tel
            If Not IsNull(rs!no_id_gst) Then .Cells(5, 6) = rs!no_id_gst
        End If
        
        rs.Close
        Set rs = Nothing
        '### Maklumat kedai ### - End
        
        .Cells(1, 6).Font.Bold = True
        .Cells(1, 6).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 6).HorizontalAlignment = xlCenter
        Next Row
        
        .Cells(7, 1) = Frm57.L16_Text 'Header Report
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Belian"
        .Cells(8, 3) = "No. Siri Produk"
        .Cells(8, 4) = "Nama Produk"
        .Cells(8, 5) = "Supplier"
        .Cells(8, 6) = "Purity"
        .Cells(8, 7) = "Berat (g)"
        .Cells(8, 8) = "Kos Per Gram (RM/g)"
        .Cells(8, 9) = "Upah (RM)"
        .Cells(8, 10) = "Harga Belian (RM)"
        .Cells(8, 11) = "Dulang"
        .Cells(8, 12) = "Panjang"
        .Cells(8, 13) = "Lebar"
        .Cells(8, 14) = "Dia"
        .Cells(8, 15) = "Saiz"
    
        For i = 1 To 15
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from inventory where status='" & "0" & "'", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
            x = x + 1
            .Cells(8 + x, 1) = x
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            If Not IsNull(rs!tarikh_belian) Then .Cells(8 + x, 2) = "'" & rs!tarikh_belian 'Tarikh Belian
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            If Not IsNull(rs!no_siri) Then .Cells(8 + x, 3) = rs!no_siri 'No. Siri Produk
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            If Not IsNull(rs!nama_produk) Then .Cells(8 + x, 4) = rs!nama_produk 'Nama Produk
            If Not IsNull(rs!supplier) Then .Cells(8 + x, 5) = rs!supplier 'Nama Supplier
            If Not IsNull(rs!purity) Then .Cells(8 + x, 6) = rs!purity 'Purity
            .Cells(8 + x, 6).HorizontalAlignment = xlCenter
            If Not IsNull(rs!Berat) Then .Cells(8 + x, 7) = Format(rs!Berat, "#,##0.00") 'Berat Jualan
            .Cells(8 + x, 7).HorizontalAlignment = xlCenter
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            If Not IsNull(rs!KOSPERGRAM) Then .Cells(8 + x, 8) = Format(rs!KOSPERGRAM, "#,##0.00") 'Harga Belian Per Gram
            .Cells(8 + x, 8).HorizontalAlignment = xlCenter
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            If Not IsNull(rs!UPAH) Then .Cells(8 + x, 9) = Format(rs!UPAH, "#,##0.00") 'Upah
            .Cells(8 + x, 9).HorizontalAlignment = xlCenter
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            If Not IsNull(rs!harga_belian) Then .Cells(8 + x, 10) = Format(rs!harga_belian, "#,##0.00") 'Kos Belian Item
            .Cells(8 + x, 10).HorizontalAlignment = xlCenter
            .Cells(8 + x, 10).NumberFormat = "#,##0.00"
            If Not IsNull(rs!dulang) Then .Cells(8 + x, 11) = rs!dulang 'Dulang
            .Cells(8 + x, 11).HorizontalAlignment = xlCenter
            If Not IsNull(rs!panjang) Then .Cells(8 + x, 12) = rs!panjang 'Panjang
            .Cells(8 + x, 12).HorizontalAlignment = xlCenter
            If Not IsNull(rs!lebar) Then .Cells(8 + x, 13) = rs!lebar 'Lebar
            .Cells(8 + x, 13).HorizontalAlignment = xlCenter
            If Not IsNull(rs!dia) Then .Cells(8 + x, 14) = rs!dia 'Dia
            .Cells(8 + x, 14).HorizontalAlignment = xlCenter
            If Not IsNull(rs!Size) Then .Cells(8 + x, 15) = rs!Size 'Saiz
            .Cells(8 + x, 15).HorizontalAlignment = xlCenter
            For Col = 1 To 15
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
    
        x = x + 2
        .Cells(8 + x, 1) = "Kuantiti : " & Frm57.L19_Text
        .Cells(8 + x, 4) = "Jumlah Berat : " & Frm57.L17_Text
        .Cells(8 + x, 6) = "Jumlah Modal : " & Frm57.L18_Text
    End With
    
    ' This makes Excel visible
    xlObject.Visible = True
    xlObject.EnableEvents = True
End If
End Sub
Private Sub Frm57_PM_Excel4_Click()
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
        .Columns("C").ColumnWidth = 100

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
                .Cells(1, 3) = rs!nama_kedai
                .Cells(1, 3).Font.Name = "Times New Roman"
            End If
            If Not IsNull(rs!no_pendaftaran) Then .Cells(2, 3) = rs!no_pendaftaran
            If Not IsNull(rs!alamat) Then .Cells(3, 3) = rs!alamat
            If Not IsNull(rs!no_tel) Then .Cells(4, 3) = rs!no_tel
            If Not IsNull(rs!no_id_gst) Then .Cells(5, 3) = rs!no_id_gst
        End If
        
        rs.Close
        Set rs = Nothing
        '### Maklumat kedai ### - End
        
        .Cells(1, 3).Font.Bold = True
        .Cells(1, 3).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 3).HorizontalAlignment = xlCenter
        Next Row
        
        .Cells(7, 1) = Frm57.L20_Text 'Header Report
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "No. Siri Produk"
        .Cells(8, 3) = "Detail"
        
        For i = 1 To 3
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from inventory2", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
            x = x + 1
            .Cells(8 + x, 1) = x
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            If Not IsNull(rs!no_siri) Then .Cells(8 + x, 2) = rs!no_siri 'No. Siri Produk
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            If Not IsNull(rs!Detail) Then .Cells(8 + x, 3) = rs!Detail 'Detail

            For Col = 1 To 3
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
    
        x = x + 2

        .Cells(8 + x, 1) = "Kuantiti : " & Frm57.L26_Text
    End With
    
    ' This makes Excel visible
    xlObject.Visible = True
    xlObject.EnableEvents = True
End If
End Sub
Private Sub Frm57_SM_scan_mode_Click()
'On Error Resume Next
If Frm57.Pic3.Visible = True Then
    Frm57.TB1.SetFocus
Else
    MsgBox "Sila buat tetapan report sebelum scan barang.", vbInformation, "Info"
End If
End Sub



Private Sub L33_Text_Click()
'on error resume next
If Frm57.Pic1.Visible = False Then
    Call Frm57_M_Clear
    Frm57.L34_Text.Visible = False
    Frm57.L35_Text.Visible = False
    Frm57.L36_Text.Visible = False
    Frm57.L37_Text.Visible = False
        
    Call Frm57_initial_setting
    
    Frm57.Pic1.Visible = True

Else
    Frm57.Pic1.Visible = False
End If
End Sub
Private Sub L34_Text_Click()
'On Error Resume Next
If Frm57.Pic2.Visible = False Then

    Call Frm57_initial_setting
    
    Frm57.Pic2.Visible = True

Else
    Frm57.Pic2.Visible = False
End If
End Sub
Private Sub L35_Text_Click()
'On Error Resume Next
If Frm57.Pic4.Visible = False Then

    Call Frm57_initial_setting
    
    Frm57.Pic4.Visible = True

Else
    Frm57.Pic4.Visible = False
End If
End Sub
Private Sub L36_Text_Click()
'On Error Resume Next
If Frm57.Pic5.Visible = False Then

    Call Frm57_initial_setting
    
    Frm57.Pic5.Visible = True

Else
    Frm57.Pic5.Visible = False
End If
End Sub
Private Sub L37_Text_Click()
'On Error Resume Next
If Frm57.Pic6.Visible = False Then

    Call Frm57_initial_setting
    
    Frm57.Pic6.Visible = True

Else
    Frm57.Pic6.Visible = False
End If
End Sub
Private Sub MSFlexGrid1_DblClick()
'On Error Resume Next
If Frm57.MSFlexGrid1 <> vbNullString Then
    PopupMenu Frm57_PM_Menu1
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub MSFlexGrid2_DblClick()
'On Error Resume Next
If Frm57.MSFlexGrid2 <> vbNullString Then
    PopupMenu Frm57_PM_Menu2
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub MSFlexGrid3_DblClick()
'On Error Resume Next
If Frm57.MSFlexGrid3 <> vbNullString Then
    PopupMenu Frm57_PM_Menu3
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub MSFlexGrid4_DblClick()
'On Error Resume Next
If Frm57.MSFlexGrid4 <> vbNullString Then
    PopupMenu Frm57_PM_Menu4
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub
Private Sub MSFlexGrid5_DblClick()
'On Error Resume Next
If Frm57.MSFlexGrid4 <> vbNullString Then
    PopupMenu Frm57_PM_Menu4
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub

Private Sub TB1_Change()
'On Error Resume Next
If Frm57.CB1 = 1 And Frm57.TB1 <> vbNullString Then
    Frm57.Tmr2.Enabled = False
    Frm57.Tmr2.Enabled = True
    Frm57.Tmr2.Interval = 100
End If
End Sub
Private Sub Tmr2_Timer()
'On Error Resume Next
If Frm57.CB1 = 1 And Frm57.TB1 <> vbNullString And Frm57.Tmr2.Enabled = True Then
    If Frm57.Tmr2.Interval = 100 Then
        Call Frm57_M_Carian
        'Frm14.TB1.SetFocus
    End If
End If
End Sub
