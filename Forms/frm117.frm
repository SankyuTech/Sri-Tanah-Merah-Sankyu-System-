VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm117 
   Caption         =   "Report Goods Received Note & Goods Delivery Note"
   ClientHeight    =   13035
   ClientLeft      =   120
   ClientTop       =   465
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
   Icon            =   "frm117.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13035
   ScaleWidth      =   23880
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic2 
      BorderStyle     =   0  'None
      Height          =   11565
      Left            =   5160
      ScaleHeight     =   11565
      ScaleWidth      =   22245
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   22245
      Begin VB.CommandButton CMD21 
         BackColor       =   &H00FFFFFF&
         Height          =   650
         Left            =   14520
         MaskColor       =   &H00400000&
         MouseIcon       =   "frm117.frx":0ECA
         MousePointer    =   99  'Custom
         Picture         =   "frm117.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Paparan sebelumnya"
         Top             =   10320
         Width           =   1100
      End
      Begin VB.CommandButton CMD22 
         BackColor       =   &H00FFFFFF&
         Height          =   650
         Left            =   15720
         MaskColor       =   &H00400000&
         MouseIcon       =   "frm117.frx":1B13
         MousePointer    =   99  'Custom
         Picture         =   "frm117.frx":1E1D
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Paparan seterusnya"
         Top             =   10320
         Width           =   1100
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   9645
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   240
         Width           =   16725
         _ExtentX        =   29501
         _ExtentY        =   17013
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
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "** Jika nilai di atas adalah negatif bermakna pihak kedai berhutang dengan supplier/agen."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   360
         TabIndex        =   43
         Top             =   11040
         Width           =   8775
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
         Left            =   960
         TabIndex        =   42
         Top             =   10800
         Width           =   2055
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
         Left            =   960
         TabIndex        =   41
         Top             =   10560
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tunai : "
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
         Left            =   360
         TabIndex        =   40
         Top             =   10800
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Emas : "
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
         Left            =   360
         TabIndex        =   39
         Top             =   10560
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Summary"
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
         TabIndex        =   38
         Top             =   10320
         Width           =   4215
      End
      Begin VB.Label L20_Text 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000C&
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
         Height          =   300
         Left            =   8760
         TabIndex        =   37
         Top             =   9960
         Width           =   1095
      End
      Begin VB.Label L21_Text 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "L21_Text"
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
         Left            =   10800
         TabIndex        =   36
         Top             =   9960
         Width           =   1095
      End
      Begin VB.Label L22_Text 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000C&
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
         Left            =   12000
         TabIndex        =   35
         Top             =   9960
         Width           =   1095
      End
      Begin VB.Label L23_Text 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000C&
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
         Left            =   13200
         TabIndex        =   34
         Top             =   9960
         Width           =   1095
      End
      Begin VB.Label L24_Text 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "L24_Text"
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
         Left            =   14400
         TabIndex        =   33
         Top             =   9960
         Width           =   1095
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
         Left            =   14040
         TabIndex        =   32
         Top             =   10320
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
         Left            =   13440
         TabIndex        =   31
         Top             =   10320
         Width           =   375
      End
      Begin VB.Label L69_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L69_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   11040
         TabIndex        =   30
         Top             =   11160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L70_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L70_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   11160
         TabIndex        =   29
         Top             =   10800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan :"
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
         TabIndex        =   25
         Top             =   9960
         Width           =   975
      End
      Begin VB.Label L25_Text 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "L25_Text"
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
         TabIndex        =   24
         Top             =   9960
         Width           =   1095
      End
      Begin VB.Label L14_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai report                       "
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
         TabIndex        =   23
         Top             =   0
         Width           =   15855
      End
      Begin VB.Label L13_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L13_Text"
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
         TabIndex        =   22
         Top             =   9975
         Width           =   1335
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
         Left            =   12120
         TabIndex        =   28
         Top             =   10320
         Width           =   2295
      End
   End
   Begin VB.PictureBox Pic1 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   120
      ScaleHeight     =   3735
      ScaleWidth      =   11385
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   11385
      Begin VB.ComboBox CBB1 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   360
         ItemData        =   "frm117.frx":2743
         Left            =   2150
         List            =   "frm117.frx":2745
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1680
         Width           =   7005
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
         Left            =   240
         TabIndex        =   3
         Top             =   320
         Width           =   200
      End
      Begin VB.CommandButton CMD3 
         BackColor       =   &H000080FF&
         Caption         =   "Report"
         Height          =   405
         Left            =   3480
         MaskColor       =   &H00400000&
         MouseIcon       =   "frm117.frx":2747
         MousePointer    =   99  'Custom
         TabIndex        =   2
         ToolTipText     =   "Report"
         Top             =   2520
         Width           =   2385
      End
      Begin VB.ComboBox CBB2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   360
         ItemData        =   "frm117.frx":2A51
         Left            =   2150
         List            =   "frm117.frx":2A53
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   2040
         Width           =   7005
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   2145
         TabIndex        =   5
         Top             =   645
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
         Left            =   2145
         TabIndex        =   6
         Top             =   1005
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
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis * "
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   435
         TabIndex        =   18
         Top             =   2040
         Width           =   1860
      End
      Begin VB.Label L8_Text 
         BackColor       =   &H00C0C0FF&
         Caption         =   "L8_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9600
         TabIndex        =   17
         Top             =   1560
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier / Agen * "
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   435
         TabIndex        =   16
         Top             =   1680
         Width           =   1860
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila klik jika ingin melihat senarai GDN && GRN di dalam tempoh tarikh di bawah."
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   480
         TabIndex        =   15
         Top             =   270
         Width           =   8370
      End
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Akhir "
         Height          =   300
         Left            =   435
         TabIndex        =   14
         Top             =   1050
         Width           =   2895
      End
      Begin VB.Label Label62 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Mula "
         Height          =   300
         Left            =   435
         TabIndex        =   13
         Top             =   690
         Width           =   2535
      End
      Begin VB.Label L7_Text 
         BackColor       =   &H00C0C0FF&
         Caption         =   "L7_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9600
         TabIndex        =   12
         Top             =   1200
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label L6_Text 
         BackColor       =   &H00C0C0FF&
         Caption         =   "L6_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9600
         TabIndex        =   11
         Top             =   840
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label L5_Text 
         BackColor       =   &H00C0C0FF&
         Caption         =   "L5_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9600
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label L9_Text 
         BackColor       =   &H00C0C0FF&
         Caption         =   "L9_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9600
         TabIndex        =   9
         Top             =   2040
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label L17_Text 
         BackColor       =   &H00C0C0FF&
         Caption         =   "L17_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9600
         TabIndex        =   8
         Top             =   2400
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label L18_Text 
         BackColor       =   &H00C0C0FF&
         Caption         =   "L18_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9600
         TabIndex        =   7
         Top             =   2760
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Shape Shape1 
         Height          =   3375
         Left            =   120
         Top             =   120
         Width           =   9255
      End
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
      Left            =   0
      MouseIcon       =   "frm117.frx":2A55
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   0
      Width           =   2055
   End
   Begin VB.Menu frm117_pm_menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu frm117_sm_cetak 
         Caption         =   "Cetak"
      End
      Begin VB.Menu frm117_sm_cetak_penyata 
         Caption         =   "Cetak penyata"
         Visible         =   0   'False
      End
      Begin VB.Menu frm117_sm_excel 
         Caption         =   "Export excel"
      End
      Begin VB.Menu frm117_sm_bar1 
         Caption         =   "-"
      End
      Begin VB.Menu frm117_sm_edit 
         Caption         =   "Edit data"
      End
      Begin VB.Menu frm117_sm_bar2 
         Caption         =   "-"
      End
      Begin VB.Menu frm117_sm_padam 
         Caption         =   "Padam data"
      End
   End
End
Attribute VB_Name = "frm117"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD21_Click()
'on error resume next
Dim frm117_LM_CURR_PAGE As Double
Dim frm117_LM_TOTAL_PAGE As Double

frm117_LM_CURR_PAGE = 0
frm117_LM_TOTAL_PAGE = 0

If frm117.L67_Text <> vbNullString And IsNumeric(frm117.L67_Text) Then
    If frm117.L68_Text <> vbNullString And IsNumeric(frm117.L68_Text) Then
        frm117_LM_CURR_PAGE = frm117.L67_Text
        frm117_LM_TOTAL_PAGE = frm117.L68_Text
        
        If frm117_LM_CURR_PAGE <> 1 And frm117_LM_CURR_PAGE <> 0 Then
        
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
                    
            Call frm117_report_gdn_grn_header
            Call frm117_report_gdn_grn
            
        End If

    End If
End If
End Sub
Private Sub CMD22_Click()
'on error resume next
Dim frm117_LM_CURR_PAGE As Double
Dim frm117_LM_TOTAL_PAGE As Double

frm117_LM_CURR_PAGE = 0
frm117_LM_TOTAL_PAGE = 0

If frm117.L67_Text <> vbNullString And IsNumeric(frm117.L67_Text) Then
    If frm117.L68_Text <> vbNullString And IsNumeric(frm117.L68_Text) Then
        frm117_LM_CURR_PAGE = frm117.L67_Text
        frm117_LM_TOTAL_PAGE = frm117.L68_Text
        
        If frm117_LM_CURR_PAGE < frm117_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call frm117_report_gdn_grn_header
            Call frm117_report_gdn_grn
            
        End If
    End If
End If
End Sub

Private Sub CMD3_Click()
'On Error Resume Next
If frm117.CB1 = 1 Then 'Pilihan tarikh
    frm117.L5_Text = 1
Else
    frm117.L5_Text = 0
End If

If frm117.CBB1 = vbNullString Then

    MsgBox "Sila buat pilihan [Supplier/agen].", vbInformation, "Info"
    
    Exit Sub
    
End If
If frm117.CBB2 = vbNullString Then

    MsgBox "Sila buat pilihan [Jenis].", vbInformation, "Info"
    
    Exit Sub
    
End If

frm117.L6_Text = frm117.DTPicker1 'Tarikh mula
frm117.L7_Text = frm117.DTPicker2 'Tarikh akhir

frm117.L8_Text = frm117.CBB1 'Nama Supplier
frm117.L9_Text = frm117.CBB2 'Jenis

frm117.L69_Text = -1 'Titik Pencarian Data
frm117.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
frm117.L67_Text = 0 'Paparan Page ke-xxx
frm117.L68_Text = 0

GM_NEXT_PREV = 0

Call frm117_report_gdn_grn_header
Call frm117_report_gdn_grn

If frm117.L13_Text <> "0" Then
    frm117.Pic1.Visible = False
    frm117.Pic2.Visible = True
Else
    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If
End Sub

Private Sub frm117_sm_cetak_Click()
'on error resume next
LM_YES = 0
DATA_FOUND = 0
frm117_LM_No_ID = vbNullString

If frm117.MSFlexGrid1 <> vbNullString Then

    If IsNumeric(frm117.MSFlexGrid1) Then
    
        frm117_LM_No_ID = frm117.MSFlexGrid1.TextMatrix(frm117.MSFlexGrid1, 2) 'No. ID
        
        If frm117_LM_No_ID <> vbNullString Then
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 77_gdn_grn where ID='" & frm117_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!cawangan) Then G_KEDAI = rs!cawangan
                If Not IsNull(rs!no_rujukan) Then
                
                    DATA_FOUND = 1
                    G_No_RESIT_JUALAN = rs!no_rujukan
                
                End If
                
                If Not IsNull(rs!jenis) Then
                
                    If rs!jenis = "GRN" Then
                        
                        LM_JENIS = "Goods Received Note"
                        
                    ElseIf rs!jenis = "GDN" Then
                    
                    
                        LM_JENIS = "Goods Despatch Note"
                        
                    ElseIf rs!jenis = "INV" Then
                    
                    
                        LM_JENIS = "INV"
                        
                    ElseIf rs!jenis = "VOU" Then

                        LM_JENIS = "VOU"
                        
                    Else
                    
                        LM_JENIS = "Goods Despatch Note"
                        
                    End If
                    
                End If
                
                If Not IsNull(rs!jenis_urusan) Then LM_JENIS_GDN = rs!jenis_urusan
                
                Note = "Anda telah memilih data bagi " & LM_JENIS & "." & vbCrLf & _
                        vbNullString & bcrlf & _
                        "Adakah anda ingin cetak penyata ini."
                        
                Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
                If Answer = vbYes Then
                
                    LM_YES = 1
                    
                End If
                
            End If
            
            rs.Close
            Set rs = Nothing
            
            If DATA_FOUND = 1 Then
            
                If LM_YES = 1 Then
                    G_PREVIEW = 1
                    
                    If LM_JENIS = "Goods Received Note" Then Call Frm116_cetak_grn
                    If LM_JENIS = "Goods Despatch Note" Then
                        
                        If LM_JENIS_GDN = 4 Then
                            
                            Call frm123_cetak_gdn
                        
                        Else
                        
                            Call Frm115_cetak_gdn
                            
                        End If
                        
                    End If
                    If LM_JENIS = "INV" Or LM_JENIS = "VOU" Then Call frm118_cetak_inv_vou
                End If
                
            Else
            
                MsgBox "Tiada data dijumpai. Sila keluar dari menu ini dan cuba sekali lagi.", vbExclamation, "Info"
            
            End If
            
        End If
        
    End If
    
End If
End Sub

Private Sub frm117_sm_cetak_penyata_Click()
'on error resume next
DATA_FOUND = 0
frm117_LM_No_ID = vbNullString

If frm117.MSFlexGrid1 <> vbNullString Then

    If IsNumeric(frm117.MSFlexGrid1) Then
    
        frm117_LM_No_ID = frm117.MSFlexGrid1.TextMatrix(frm117.MSFlexGrid1, 2) 'No. ID
        
        If frm117_LM_No_ID <> vbNullString Then
        
            Note = "Adakah anda ingin cetak penyata ini?"
                    
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

            If Answer = vbYes Then
                
                Call frm117_cetak_statement
            
            End If
            
        End If
        
    End If
    
End If
End Sub

Private Sub frm117_sm_edit_Click()
'on error resume next
DATA_FOUND = 0
frm117_LM_No_ID = vbNullString
LM_JENIS = vbNullString

If frm117.MSFlexGrid1 <> vbNullString Then

    If IsNumeric(frm117.MSFlexGrid1) Then
    
        frm117_LM_No_ID = frm117.MSFlexGrid1.TextMatrix(frm117.MSFlexGrid1, 2) 'No. ID
        
        If frm117_LM_No_ID <> vbNullString Then
        
            Note = "Adakah anda ingin edit data ini?"
                    
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

            If Answer = vbYes Then
            
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 77_gdn_grn where ID='" & frm117_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    
                    If Not IsNull(rs!no_rujukan) Then
                    
                        DATA_FOUND = 1
                        G_No_RESIT_JUALAN = rs!no_rujukan
                    
                    End If
                    
                    If Not IsNull(rs!jenis) Then
                    
                        If rs!jenis = "GRN" Then
                            
                            LM_JENIS = "GRN"
                        
                        ElseIf rs!jenis = "GDN" Then
                            
                            LM_JENIS = "GDN"
                        
                        ElseIf rs!jenis = "INV" Then
                        
                            LM_JENIS = "INV"
                            
                        ElseIf rs!jenis = "VOU" Then
                        
                            LM_JENIS = "VOU"
                            
                        End If
                        
                    End If
                    
                    If Not IsNull(rs!jenis_urusan) Then LM_JENIS_GDN = rs!jenis_urusan
                    
                End If
                
                rs.Close
                Set rs = Nothing
                
                If DATA_FOUND = 1 Then
                    
                    If LM_JENIS <> vbNullString Then
                    
                        If LM_JENIS = "GRN" Then Call frm117_edit_data_grn
                        
                        If (LM_JENIS = "INV" Or LM_JENIS = "VOU") Then Call frm117_edit_data_inv_vou
                        
                        If LM_JENIS = "GDN" Then
                            
                            If LM_JENIS_GDN = 4 Then
                                
                                Call frm123_edit_data_gdn_bulk
                            
                            Else
                                
                                GLOBAL_DISABLE = 0
                                Frm115.TB1 = vbNullString
                                
                                Call Frm115_reset_1
                                Call Frm115_reset_2
                                Call Frm115_reset_3
                                Call Frm115_reset_main
                                Call Frm115_reset_main2
                                
                                Frm115.DTPicker1 = DateTime.Date$
                                
                                Frm115.L32_Text = 1 '0 : Data Baru , 1 : Edit Data
                                Frm115.L54_Text = LM_ID
                                
                                Frm115.CMD8.Visible = False
                                Frm115.CMD9.Visible = False
                                Frm115.CMD10.Visible = True
                                Frm115.CMD11.Visible = True
                                
                                Frm115.L23_Text = G_No_RESIT_JUALAN
                                
                                Call frm115_initial_setting_stok
                                Call Frm115_Senarai_Jualan_Header
                                Call frm115_reset_gdn_list
                                
                                MDI_frm1.L5_Text = 16
                                Call Frm115_recall_edit_jualan
                                Call Frm115_background_color
                                
                                Frm115.L71_Text = "1"
                                
                                Frm115.TB1.SetFocus
                                
                            End If

                
                        End If
                        
                    Else
                        
                        MsgBox "Tiada data dijumpai. Sila keluar dari menu ini dan cuba sekali lagi.", vbExclamation, "Info"
                    
                    End If
                    
                End If
                
            End If
            
        End If
        
    End If
    
End If
End Sub

Private Sub frm117_sm_excel_Click()
'on error resume next
Dim TA As Date
Dim TM As Date

Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook

If frm117.MSFlexGrid1 <> vbNullString Then

    If IsNumeric(frm117.MSFlexGrid1) Then
    
        frm117_LM_No_ID = frm117.MSFlexGrid1.TextMatrix(frm117.MSFlexGrid1, 2) 'No. ID
        
        If frm117_LM_No_ID <> vbNullString Then
        
            Note = "Adakah anda ingin export semua data ini ke excel?" & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sistem mungkin mengambil masa untuk export semua data ini." & vbCrLf & _
                    "Sila tunggu sehingga sistem selesai export data ini." & vbCrLf & _
                    "Teruskan?"
                    
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

            If Answer = vbYes Then


                Set xlObject = New Excel.Application
                Set xlWB = xlObject.Workbooks.Add
                           
                'xlObject.Visible = True
                With xlObject.ActiveWorkbook.ActiveSheet
                
                    .Cells.VerticalAlignment = xlCenter
                    .Columns("A").ColumnWidth = 5 'No.
                    .Columns("B").ColumnWidth = 15 'Tarikh
                    .Columns("C").ColumnWidth = 15 'Jenis
                    .Columns("D").ColumnWidth = 20 'No. Rujukan
                    .Columns("E").ColumnWidth = 40 'Nama Supplier/Agen
                    .Columns("F").ColumnWidth = 15 'Berat (g)
                    .Columns("G").ColumnWidth = 15 'Mutu
                    .Columns("H").ColumnWidth = 15 'Hutang (g)
                    .Columns("I").ColumnWidth = 15 'Bayar (g)
                    .Columns("J").ColumnWidth = 15 'Hutang (RM)
                    .Columns("K").ColumnWidth = 15 'Bayar (RM)
                    .Columns("L").ColumnWidth = 15 'GST (RM)
                
                    '### Maklumat kedai ### - Start
                    If MDI_frm1.L4_Text = "HQ" Then
                        
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
                    
                    x = 0
                    
                    LM_TARIKH = "tiada pilihan tempoh"
                    LM_SUPPLIER = "semua supplier dan agen"
                    LM_JENIS = "GDN/GRN/INV/VOU"
                    
                    If frm117.L5_Text = "1" Then '0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh
                    
                        TM = frm117.L6_Text 'Tarikh mula
                        TA = frm117.L7_Text 'Tarikh akhir
                        LM_TARIKH = " dari " & frm117.L6_Text & " hingga " & frm117.L7_Text
                        
                    End If
                    
                    If frm117.L8_Text = "Semua supplier dan agen" Then
                        
                        frm117_LM_SEARCH_1 = Null
                        frm117_LM_SEARCH_1_LOGIC = "<>"
                        
                    Else
                        
                        frm117_LM_SEARCH_1 = frm117.L8_Text
                        frm117_LM_SEARCH_1_LOGIC = "="
                        LM_SUPPLIER = frm117.L8_Text
                        
                    End If
                    
                    If frm117.L9_Text = "Semua GDN/GRN/INV/VOU" Then
                        
                        frm117_LM_SEARCH_2 = Null
                        frm117_LM_SEARCH_2_LOGIC = "<>"
                    
                    Else
                        
                        frm117_LM_SEARCH_2 = frm117.L9_Text
                        frm117_LM_SEARCH_2_LOGIC = "="
                        LM_JENIS = frm117.L9_Text
                        
                    End If
                
                    .Cells(1, 5).Font.Bold = True
                    .Cells(1, 5).Font.Size = 30
                    
                    For Row = 1 To 5
                        .Cells(Row, 5).HorizontalAlignment = xlCenter
                    Next Row
                    
                    .Cells(7, 1) = "Senarai " & LM_JENIS & " bagi " & LM_SUPPLIER & " dan " & LM_TARIKH & "."
                    
                    .Cells(8, 1) = "No."
                    .Cells(8, 2) = "Tarikh"
                    .Cells(8, 3) = "Jenis"
                    .Cells(8, 4) = "No. Rujukan"
                    .Cells(8, 5) = "Nama Supplier/Agen"
                    .Cells(8, 6) = "Berat (g)"
                    .Cells(8, 7) = "Mutu"
                    .Cells(8, 8) = "Hutang (g)"
                    .Cells(8, 9) = "Bayar (g)"
                    .Cells(8, 10) = "Hutang (RM)"
                    .Cells(8, 11) = "Bayar (RM)"
                    .Cells(8, 12) = "GST (RM)"
                    
                    For i = 1 To 12
                        .Cells(8, i).HorizontalAlignment = xlCenter
                        .Cells(8, i).Interior.ColorIndex = 15
                        .Cells(8, i).WrapText = True
                        .Cells(8, i).Borders.LineStyle = xlContinuous
                    Next i
                    
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    If frm117.L5_Text = 0 Then rs.Open "select * from 77_gdn_grn where supplier_agen " & frm117_LM_SEARCH_1_LOGIC & "'" & frm117_LM_SEARCH_1 & "' AND status = 1 AND jenis " & frm117_LM_SEARCH_2_LOGIC & "'" & frm117_LM_SEARCH_2 & "' order by write_timestamp ASC", cn, adOpenKeyset, adLockOptimistic
                    If frm117.L5_Text = 1 Then rs.Open "select * from 77_gdn_grn where supplier_agen " & frm117_LM_SEARCH_1_LOGIC & "'" & frm117_LM_SEARCH_1 & "' AND status = 1 AND jenis " & frm117_LM_SEARCH_2_LOGIC & "'" & frm117_LM_SEARCH_2 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by write_timestamp ASC", cn, adOpenKeyset, adLockOptimistic
                    
                    While rs.EOF = False
                    
                        x = x + 1
                        .Cells(8 + x, 1) = x 'No.
                        .Cells(8 + x, 1).HorizontalAlignment = xlCenter
                        
                        If Not IsNull(rs!tarikh) Then .Cells(8 + x, 2) = "'" & rs!tarikh 'Tarikh
                        .Cells(8 + x, 2).HorizontalAlignment = xlCenter
                    
                        If Not IsNull(rs!jenis) Then .Cells(8 + x, 3) = rs!jenis 'Jenis
                        .Cells(8 + x, 3).HorizontalAlignment = xlCenter
                        
                        If Not IsNull(rs!no_rujukan) Then .Cells(8 + x, 4) = rs!no_rujukan 'No. Rujukan
                
                        If Not IsNull(rs!supplier_agen) Then .Cells(8 + x, 5) = rs!supplier_agen 'Supplier/Agen
                
                        .Cells(8 + x, 6).HorizontalAlignment = xlRight
                        If Not IsNull(rs!Berat_Asal) Then
                            .Cells(8 + x, 6) = Format(rs!Berat_Asal, "#,##0.00") 'Berat (g)
                            .Cells(8 + x, 6).NumberFormat = "#,##0.00"
                        End If
                        
                        If Not IsNull(rs!kadar_tukaran) Then .Cells(8 + x, 7) = rs!kadar_tukaran 'Mutu
                        .Cells(8 + x, 7).NumberFormat = "#,##0.000"
                        .Cells(8 + x, 7).HorizontalAlignment = xlCenter
                        
                        .Cells(8 + x, 8).HorizontalAlignment = xlRight
                        If Not IsNull(rs!berat_tukaran_grn) Then
                            .Cells(8 + x, 8) = Format(rs!berat_tukaran_grn, "#,##0.00") 'Hutang (Emas)
                            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
                        End If
                        
                        .Cells(8 + x, 9).HorizontalAlignment = xlRight
                        If Not IsNull(rs!berat_tukaran) Then
                            .Cells(8 + x, 9) = Format(rs!berat_tukaran, "#,##0.00") 'Bayar (Emas)
                            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
                        End If
                        
                        If Not IsNull(rs!jenis_urusan) Then
                            
                            If rs!jenis_urusan <> "3" Then
                                
                                .Cells(8 + x, 10).HorizontalAlignment = xlRight
                                If Not IsNull(rs!harga_dengan_gst_grn) Then
                                    .Cells(8 + x, 10) = Format(rs!harga_dengan_gst_grn, "#,##0.00") 'Hutang (RM)
                                    .Cells(8 + x, 10).NumberFormat = "#,##0.00"
                                End If
                                
                                .Cells(8 + x, 11).HorizontalAlignment = xlRight
                                If Not IsNull(rs!harga_dengan_gst) Then
                                    .Cells(8 + x, 11) = Format(rs!harga_dengan_gst, "#,##0.00") 'Bayar (RM)
                                    .Cells(8 + x, 11).NumberFormat = "#,##0.00"
                                End If
                        
                                .Cells(8 + x, 12).HorizontalAlignment = xlRight
                                If Not IsNull(rs!jumlah_gst) Then
                                    .Cells(8 + x, 12) = Format(rs!jumlah_gst, "#,##0.00") 'GST (RM)
                                    .Cells(8 + x, 12).NumberFormat = "#,##0.00"
                                End If
                            
                            Else
                                
                                If rs!umum_berat = "0" Then
                                    
                                    .Cells(8 + x, 10).HorizontalAlignment = xlRight
                                    If Not IsNull(rs!harga_dengan_gst_grn) Then
                                        .Cells(8 + x, 10) = Format(rs!harga_dengan_gst_grn, "#,##0.00") 'Hutang (RM)
                                        .Cells(8 + x, 10).NumberFormat = "#,##0.00"
                                    End If
                                    
                                    .Cells(8 + x, 11).HorizontalAlignment = xlRight
                                    If Not IsNull(rs!harga_dengan_gst) Then
                                        .Cells(8 + x, 11) = Format(rs!harga_dengan_gst, "#,##0.00") 'Bayar (RM)
                                        .Cells(8 + x, 11).NumberFormat = "#,##0.00"
                                    End If
                            
                                    .Cells(8 + x, 12).HorizontalAlignment = xlRight
                                    If Not IsNull(rs!jumlah_gst) Then
                                        .Cells(8 + x, 12) = Format(rs!jumlah_gst, "#,##0.00") 'GST (RM)
                                        .Cells(8 + x, 12).NumberFormat = "#,##0.00"
                                    End If
                    
                                ElseIf rs!umum_berat = "1" Then
                                
                                
                                End If
                                
                            End If
                        Else
                            
                            .Cells(8 + x, 10).HorizontalAlignment = xlRight
                            If Not IsNull(rs!harga_dengan_gst_grn) Then
                                .Cells(8 + x, 10) = Format(rs!harga_dengan_gst_grn, "#,##0.00") 'Hutang (RM)
                                .Cells(8 + x, 10).NumberFormat = "#,##0.00"
                            End If
                            
                            .Cells(8 + x, 11).HorizontalAlignment = xlRight
                            If Not IsNull(rs!harga_dengan_gst) Then
                                .Cells(8 + x, 11) = Format(rs!harga_dengan_gst, "#,##0.00") 'Bayar (RM)
                                .Cells(8 + x, 11).NumberFormat = "#,##0.00"
                            End If
                    
                            .Cells(8 + x, 12).HorizontalAlignment = xlRight
                            If Not IsNull(rs!jumlah_gst) Then
                                .Cells(8 + x, 12) = Format(rs!jumlah_gst, "#,##0.00") 'GST (RM)
                                .Cells(8 + x, 12).NumberFormat = "#,##0.00"
                            End If
                            
                        End If
                                            
                        For Col = 1 To 12
                            .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
                        Next Col
                            
                        rs.MoveNext
                        
                    Wend
                    
                    rs.Close
                    Set rs = Nothing
                    
                    Y = 1
                    Y = x + 1
                    
                    .Cells(8 + Y, 1).HorizontalAlignment = xlCenter 'Bilangan Data
                    .Cells(8 + Y, 1) = "Bil : " & frm117.L13_Text
                    .Cells(8 + Y, 1).Font.Bold = True
                    
                    .Cells(8 + Y, 6).HorizontalAlignment = xlRight 'Jumlah berat
                    .Cells(8 + Y, 6) = Format(frm117.L20_Text, "#,##0.00 g")
                    .Cells(8 + Y, 6).NumberFormat = "#,##0.00"
                    .Cells(8 + Y, 6).Font.Bold = True
                    
                    .Cells(8 + Y, 8).HorizontalAlignment = xlRight 'Jumlah berat (Hutang)
                    .Cells(8 + Y, 8) = Format(frm117.L21_Text, "#,##0.00 g")
                    .Cells(8 + Y, 8).NumberFormat = "#,##0.00"
                    .Cells(8 + Y, 8).Font.Bold = True
                    
                    .Cells(8 + Y, 9).HorizontalAlignment = xlRight 'Jumlah berat (Bayar)
                    .Cells(8 + Y, 9) = Format(frm117.L22_Text, "#,##0.00 g")
                    .Cells(8 + Y, 9).NumberFormat = "#,##0.00"
                    .Cells(8 + Y, 9).Font.Bold = True
                    
                    .Cells(8 + Y, 10).HorizontalAlignment = xlRight 'Jumlah Tunai (Hutang)
                    .Cells(8 + Y, 10) = "RM " & Format(frm117.L23_Text, "#,##0.00")
                    .Cells(8 + Y, 10).NumberFormat = "#,##0.00"
                    .Cells(8 + Y, 10).Font.Bold = True
                    
                    .Cells(8 + Y, 11).HorizontalAlignment = xlRight 'Jumlah Tunai (Bayar)
                    .Cells(8 + Y, 11) = "RM " & Format(frm117.L24_Text, "#,##0.00")
                    .Cells(8 + Y, 11).NumberFormat = "#,##0.00"
                    .Cells(8 + Y, 11).Font.Bold = True
                    
                    .Cells(8 + Y, 12).HorizontalAlignment = xlRight 'Jumlah GST
                    .Cells(8 + Y, 12) = "RM " & Format(frm117.L25_Text, "#,##0.00")
                    .Cells(8 + Y, 12).NumberFormat = "#,##0.00"
                    .Cells(8 + Y, 12).Font.Bold = True
                    
                    Y = Y + 2
                    .Cells(8 + Y, 1).Font.Bold = True
                    .Cells(8 + Y, 1) = "Ringkasan"
                    Y = Y + 1
                    .Cells(8 + Y, 1) = "Emas : " & Format(frm117.L26_Text, "#,##0.00 g")
                    Y = Y + 1
                    .Cells(8 + Y, 1) = "Tunai : RM " & Format(frm117.L27_Text, "#,##0.00")
                    Y = Y + 1
                    .Cells(8 + Y, 1) = "** Jika nilai di atas adalah negatif bermakna pihak kedai berhutang dengan supplier/agen."
                    
                    Y = Y + 4
                    .Cells(8 + Y, 1).Font.Bold = True
                    .Cells(8 + Y, 1) = "Report Generated By Sankyu System" 'Watermark Sankyu System
                    Y = Y + 1
                    .Cells(8 + Y, 1).Font.Bold = True
                    .Cells(8 + Y, 1) = "Sankyu System , +6010 - 900 4788 , sankyusystem@gmail.com" 'Watermark Sankyu System
                End With
                    
                ' This makes Excel visible
                xlObject.Visible = True
                xlObject.EnableEvents = True
            
            End If
            
        End If
        
    End If
    
End If
End Sub

Private Sub frm117_sm_padam_Click()
'on error resume next
DATA_FOUND = 0
frm117_LM_No_ID = vbNullString
LM_JENIS = vbNullString

If frm117.MSFlexGrid1 <> vbNullString Then

    If IsNumeric(frm117.MSFlexGrid1) Then
    
        frm117_LM_No_ID = frm117.MSFlexGrid1.TextMatrix(frm117.MSFlexGrid1, 2) 'No. ID
        
        If frm117_LM_No_ID <> vbNullString Then
        
            Note = "Adakah anda ingin padam data ini?"
                    
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

            If Answer = vbYes Then
            
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 77_gdn_grn where ID='" & frm117_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    
                    If Not IsNull(rs!no_rujukan) Then
                    
                        DATA_FOUND = 1
                        If Not IsNull(rs!no_rujukan) Then G_No_RESIT_JUALAN = rs!no_rujukan
                        If Not IsNull(rs!ID) Then G_ID = rs!ID
                        If Not IsNull(rs!jenis) Then
                        
                            If rs!jenis = "GRN" Then
                                
                                LM_JENIS = "GRN"
                            
                            ElseIf rs!jenis = "GDN" Then
                                
                                LM_JENIS = "GDN"
                                
                            ElseIf (rs!jenis = "INV" Or rs!jenis = "VOU") Then
                                
                                LM_JENIS = "INV"
                                
                            End If
                            
                        End If
                        If Not IsNull(rs!jenis_urusan) Then LM_JENIS_GDN = rs!jenis_urusan
                    End If
                    
                End If
                
                rs.Close
                Set rs = Nothing
                
                If DATA_FOUND = 1 Then
                
                    Note = "Adakah anda ingin padamkan data ini?" & vbCrLf & _
                            "Data ini akan dipadamkan dari database." & vbCrLf & _
                            vbNullString & vbCrLf & _
                            "No. rujukan GRN/GDN/INV/VOU yang akan dipadamkan adalah " & G_No_RESIT_JUALAN & vbCrLf & _
                            vbNullString & vbCrLf & _
                            "Sila masukkan password bagi memadamkan data ini."
                
                    LM_PASSWORD = InputBox(Note, "Padam Invoice", "Sila masukkan password anda")
                    
                    If StrPtr(LM_PASSWORD) = 0 Then
                        Exit Sub
                    End If
                    
                    If StrPtr(LM_PASSWORD) <> 0 Then
                
                        If InStr(1, LM_PASSWORD, "*") <> 0 Or InStr(1, LM_PASSWORD, "&") <> 0 Or InStr(1, LM_PASSWORD, "/") <> 0 Or InStr(1, LM_PASSWORD, "\") <> 0 Or InStr(1, LM_PASSWORD, "'") <> 0 Then
                            MsgBox "Password mengandungi simbol yang tidak sah.", vbExclamation, "Error"
                            
                            Exit Sub
                        End If
                        
                        If MDI_frm1.L3_Text <> vbNullString Then
                        
                            Set rs = New ADODB.Recordset
                            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                            rs.Open "select * from employee where samaran='" & MDI_frm1.L3_Text & "' and password='" & LM_PASSWORD & "'", cn, adOpenKeyset, adLockOptimistic
                            
                            If Not rs.EOF Then
                            
                                LM_USER_FOUND = 1
                                
                            Else
                            
                                MsgBox "Password yang dimasukkan tidak betul/sah." & vbCrLf & _
                                        vbNullString & vbCrLf & _
                                        "Sila cuba sekali lagi.", vbExclamation, "Info"
                                        
                            End If
                            
                            rs.Close
                            Set rs = Nothing
                            
                            If LM_USER_FOUND = 1 Then
                                
                                If LM_JENIS <> vbNullString Then
                                
                                    If LM_JENIS = "GRN" Then Call frm117_padam_grn
                                    If LM_JENIS = "GDN" Then
                                        If LM_JENIS_GDN = 4 Then
                                            Call frm123_padam_gdn_bulk
                                        Else
                                            Call frm117_padam_gdn
                                        End If
                                    End If
                                    If LM_JENIS = "INV" Then Call frm117_padam_inv_vou
                                    
                                Else
                                    
                                    MsgBox "Tiada data dijumpai. Sila keluar dari menu ini dan cuba sekali lagi.", vbExclamation, "Info"
                                    
                                End If
                                
                            End If
                            
                        End If
                    
                    End If
                
                End If
                
            End If
            
        End If
        
    End If
    
End If
End Sub

Private Sub L1_Text_Click()
'On Error Resume Next
Call frm117_pic_ena_disable
frm117.Pic1.Visible = True
End Sub
Private Sub MSFlexGrid1_DblClick()
'On Error Resume Next
frm117_LM_No_ID = vbNullString

If frm117.MSFlexGrid1 <> vbNullString Then

    If IsNumeric(frm117.MSFlexGrid1) Then
    
        frm117_LM_No_ID = frm117.MSFlexGrid1.TextMatrix(frm117.MSFlexGrid1, 2) 'No. ID
        
        If frm117_LM_No_ID <> vbNullString Then
        
        
            user_level = MDI_frm1.L4_Text
        
            If user_level = "Admin" Or user_level = "HQ" Or user_level = "Developer" Then
            
                frm117.frm117_sm_edit.Enabled = True
                frm117.frm117_sm_padam.Enabled = True
                        
            ElseIf user_level = "Manager" Then
            
                frm117.frm117_sm_edit.Enabled = True
                frm117.frm117_sm_padam.Enabled = False
                
            Else
            
                frm117.frm117_sm_edit.Enabled = False
                frm117.frm117_sm_padam.Enabled = False
                
            End If
        
            PopupMenu frm117_pm_menu
            
        Else
            
            MsgBox "Tiada data.", vbExclamation, "Info"
            
        End If
        
    Else
    
        MsgBox "Tiada data.", vbExclamation, "Info"
        
    End If
    
End If
End Sub


