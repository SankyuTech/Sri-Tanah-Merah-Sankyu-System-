VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm113 
   Caption         =   "Mata ganjaran"
   ClientHeight    =   13035
   ClientLeft      =   120
   ClientTop       =   465
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
   Begin VB.PictureBox Pic1 
      BorderStyle     =   0  'None
      Height          =   11565
      Left            =   7200
      ScaleHeight     =   11565
      ScaleWidth      =   22245
      TabIndex        =   1
      Top             =   -960
      Width           =   22245
      Begin VB.CommandButton CMD1 
         BackColor       =   &H00FFFFFF&
         Height          =   740
         Left            =   14760
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm113.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Frm113.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Paparan sebelumnya"
         Top             =   10080
         Width           =   1000
      End
      Begin VB.CommandButton CMD2 
         BackColor       =   &H00FFFFFF&
         Height          =   740
         Left            =   15840
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm113.frx":0C49
         MousePointer    =   99  'Custom
         Picture         =   "Frm113.frx":0F53
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Paparan seterusnya"
         Top             =   10080
         Width           =   1000
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   9645
         Left            =   120
         TabIndex        =   4
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
         Left            =   3120
         TabIndex        =   16
         Top             =   9975
         Width           =   1335
      End
      Begin VB.Label L12_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L12_Text"
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
         Left            =   3120
         TabIndex        =   15
         Top             =   10560
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   ":    :    :    :"
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
         Height          =   780
         Left            =   3000
         TabIndex        =   14
         Top             =   9960
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Senarai mata ganjaran ahli.                                "
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
         TabIndex        =   13
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label L6_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L6_Text"
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
         Left            =   14370
         TabIndex        =   10
         Top             =   10080
         Width           =   615
      End
      Begin VB.Label L5_Text 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L5_Text"
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
         Left            =   13860
         TabIndex        =   9
         Top             =   10080
         Width           =   375
      End
      Begin VB.Label L7_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L7_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   11400
         TabIndex        =   8
         Top             =   10440
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L8_Text 
         BackColor       =   &H8000000C&
         Caption         =   "L8_Text"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   11400
         TabIndex        =   7
         Top             =   10680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label L10_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L10_Text"
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
         Left            =   3120
         TabIndex        =   6
         Top             =   10155
         Width           =   1575
      End
      Begin VB.Label L11_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L11_Text"
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
         Left            =   3120
         TabIndex        =   5
         Top             =   10365
         Width           =   1335
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "Paparan Muka  :       / "
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
         Left            =   12600
         TabIndex        =   11
         Top             =   10080
         Width           =   2295
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm113.frx":1879
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
         Height          =   780
         Left            =   120
         TabIndex        =   12
         Top             =   9960
         Width           =   3015
      End
   End
   Begin VB.PictureBox Pic2 
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   120
      ScaleHeight     =   5415
      ScaleWidth      =   6345
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   6345
      Begin VB.CommandButton CMD5 
         BackColor       =   &H000080FF&
         Caption         =   "Batal"
         Height          =   405
         Left            =   3120
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm113.frx":1906
         MousePointer    =   99  'Custom
         TabIndex        =   39
         Top             =   4200
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.CommandButton CMD4 
         BackColor       =   &H000080FF&
         Caption         =   "Simpan"
         Height          =   405
         Left            =   960
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm113.frx":1C10
         MousePointer    =   99  'Custom
         TabIndex        =   38
         Top             =   4200
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.TextBox TB1 
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
         Height          =   360
         Left            =   1770
         TabIndex        =   30
         Text            =   "TB1"
         Top             =   1560
         Width           =   4245
      End
      Begin VB.ComboBox CBB1 
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
         Height          =   360
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   2280
         Width           =   4245
      End
      Begin VB.CheckBox CB1 
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
         Left            =   120
         TabIndex        =   26
         Top             =   615
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
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   200
      End
      Begin VB.CommandButton CMD3 
         BackColor       =   &H000080FF&
         Caption         =   "Simpan"
         Height          =   405
         Left            =   2160
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm113.frx":1F1A
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   4200
         Width           =   2025
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   1770
         TabIndex        =   31
         Top             =   1920
         Width           =   4245
         _ExtentX        =   7488
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
         Format          =   414777344
         CurrentDate     =   41561
      End
      Begin VB.TextBox TB2 
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
         Height          =   1200
         Left            =   1770
         MaxLength       =   75
         TabIndex        =   28
         Text            =   "TB2"
         Top             =   2640
         Width           =   4245
      End
      Begin VB.Label L18_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "Anda tidak dibenarkan untuk menukar jenis pilihan ini bagi menu edit data.Sila padam data jika anda tersilap membuat pilihan ini."
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
         Height          =   525
         Left            =   120
         TabIndex        =   41
         Top             =   1080
         Visible         =   0   'False
         Width           =   5895
      End
      Begin VB.Label L17_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L17_Text"
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
         Left            =   480
         TabIndex        =   40
         Top             =   4440
         Visible         =   0   'False
         Width           =   975
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
         Left            =   4320
         TabIndex        =   37
         Top             =   3840
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Max. tulisan 75. Baki adalah "
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
         Left            =   1800
         TabIndex        =   36
         Top             =   3840
         Width           =   2655
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah mata *"
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
         TabIndex        =   35
         Top             =   1590
         Width           =   1545
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh *"
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
         TabIndex        =   34
         Top             =   1920
         Width           =   2385
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pekerja *"
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
         TabIndex        =   33
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         TabIndex        =   32
         Top             =   2670
         Width           =   1545
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Tambahan mata ganjaran Potongan mata ganjaran"
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
         Height          =   525
         Left            =   405
         TabIndex        =   27
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila pilih samada tambahan atau potongan mata bagi ahli ini."
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
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   8610
      End
   End
   Begin VB.Label L15_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tambahan / potongan mata"
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
      Left            =   2640
      MouseIcon       =   "Frm113.frx":2224
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label L14_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai mata ganjaran"
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
      MouseIcon       =   "Frm113.frx":252E
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   720
      Width           =   2295
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
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   1320
      TabIndex        =   22
      Top             =   330
      Width           =   7095
   End
   Begin VB.Label L9_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L9_Text"
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
      Left            =   1320
      TabIndex        =   21
      Top             =   510
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Maklumat mata ganjaran bagi   Nama              :                                  No. Keahlian  :"
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
      Height          =   660
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label L3_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Keluar"
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
      MouseIcon       =   "Frm113.frx":2838
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Menu Frm113_PM_menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm113_SM_edit 
         Caption         =   "Edit / lihat data ini"
      End
      Begin VB.Menu Frm113_SM_cetak_penyata_mata 
         Caption         =   "Cetak penyata mata ganjaran"
      End
   End
End
Attribute VB_Name = "Frm113"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB1_Click()
'on error resume next
If Frm113.CB1 = 1 Then
    Frm113.CB2 = 0
End If
End Sub
Private Sub CB2_Click()
'on error resume next
If Frm113.CB2 = 1 Then
    Frm113.CB1 = 0
End If
End Sub
Private Sub CMD1_Click()
'on error resume next
GM_NEXT_PREV = 1 '0 : Next , 1 : Previous

Call Frm113_senarai_mata_ganjaran_header
Call Frm113_senarai_mata_ganjaran
End Sub
Private Sub CMD2_Click()
'on error resume next
Dim Frm113_LM_CURR_PAGE As Double
Dim Frm113_LM_TOTAL_PAGE As Double

Frm113_LM_CURR_PAGE = 0
Frm113_LM_TOTAL_PAGE = 0

If Frm113.L5_Text <> vbNullString And IsNumeric(Frm113.L5_Text) Then
    If Frm113.L6_Text <> vbNullString And IsNumeric(Frm113.L6_Text) Then
        Frm113_LM_CURR_PAGE = Frm113.L5_Text
        Frm113_LM_TOTAL_PAGE = Frm113.L6_Text
        
        If Frm113_LM_CURR_PAGE < Frm113_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call Frm113_senarai_mata_ganjaran_header
            Call Frm113_senarai_mata_ganjaran
            
        End If
    End If
End If
End Sub
Private Sub CMD3_Click()
'On Error Resume Next
Dim Err(5)
Dim Frm113_LM_MATA_TERKINI As Double
Dim Frm113_LM_MATA_DEDUCT As Double
Dim Frm113_LM_MATA_TERBARU As Double

DATA_UPDATE_1 = 1
DATA_UPDATE_2 = 1

Frm113_LM_EMP_NO = vbNullString
Frm113_LM_MATA_TERKINI = 0
Frm113_LM_MATA_DEDUCT = 0
Frm113_LM_MATA_TERBARU = 0

If Frm113.CB1 = 0 And Frm113.CB2 = 0 Then
    x = x + 1
    Err(x) = "Sila pilih samada [Tambahan mata ganjaran] atau [Potongan mata ganjaran]."
End If
If Frm113.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Nama Pekerja]."
End If
If Frm113.TB1 = vbNullString Or (Frm113.TB1 <> vbNullString And Not IsNumeric(Frm113.TB1)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Jumlah mata]. Hanya NOMBOR dibenarkan dalam ruangan ini."
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
    
    If Answer = vbYes Then
    
        If Frm113.CB1 = 1 Then
            
            Note = "Sebanyak " & Frm113.TB1 & " mata ganjaran akan diberikan kepada ahli ini." & vbCrLf & _
                    "Teruskan?"
            
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbNo Then
                
                Exit Sub
                
            End If
        
        End If
        If Frm113.CB2 = 1 Then
        
            Note = "Sebanyak " & Frm113.TB1 & " mata ganjaran akan ditolak dari  kepada ahli ini." & vbCrLf & _
                    "Teruskan?"
            
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbNo Then
                
                Exit Sub
                
            End If
            
            '### Periksa mata terkini ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm113.L9_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!baki_point) Then
                    If IsNumeric(rs!baki_point) Then Frm113_LM_MATA_TERKINI = rs!baki_point
                End If
            Else
                
                MsgBox "Tiada maklumat berkenaan dengan keahlian ini." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Kemungkinan ahli ini telah dipadamkan data dari sistem.", vbInformation, "Info"
                
                Exit Sub
                
            End If
            
            rs.Close
            Set rs = Nothing
            '### Periksa mata terkini ### - End
            
            Frm113_LM_MATA_DEDUCT = Frm113.TB1 'Mata yang dipotong / ditambah
            
            Frm113_LM_MATA_TERBARU = Frm113_LM_MATA_TERKINI - Frm113_LM_MATA_DEDUCT
            
            If Frm113_LM_MATA_TERBARU < 0 Then
                
                MsgBox "Mata yang ingin dipotong adalah lebih dari mata terkumpul yang dimiliki oleh ahli ini." & vbCrLf & _
                        "Maklumat baki terkini adalah seperti di bawah :" & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Baki mata ganjaran terkumpul : " & Frm113_LM_MATA_TERKINI & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Anda tidak dibenarkan untuk memotong mata ganjaran kurang dari mata terkumpul", vbeclamation, "Info"
                        
                        Exit Sub
            
            End If
            
        End If
        
        If Frm113.CBB1 <> vbNullString Then
            Frm113_LM_EMP_NO = Split(Frm113.CBB1, "  |  ")(1)
        End If
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 71_tebus_agih_point", cn, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        If Frm113.DTPicker1 <> vbNullString Then
            rs!tarikh = Frm113.DTPicker1 'Tarikh
        Else
            rs!tarikh = Null 'Tarikh
        End If
        If Frm113.L9_Text <> vbNullString Then 'No. Rujukan Pembeli
            rs!no_ahli = Frm113.L9_Text
        Else
            rs!no_ahli = Null
        End If
        If Frm113.CB1 = 1 Then
        
            rs!jumlah_peroleh_point = Frm113.TB1 'Jumlah mata yang ditambah
            rs!jumlah_tebus_point = Null 'Jumlah mata yang dipotong
            rs!Type = 2
            
        End If
        
        If Frm113.CB2 = 1 Then
        
            rs!jumlah_peroleh_point = Null 'Jumlah mata yang ditambah
            rs!jumlah_tebus_point = Frm113.TB1 'Jumlah mata yang dipotong
            rs!Type = 3
            
        End If
        rs!remarks = Frm113.TB2 'Remarks
        rs!write_timestamp = Now
        rs!Status = 1 '0 : Tidak aktif , 1 : Aktif
        rs!no_pekerja = Frm113_LM_EMP_NO
        DATA_UPDATE_1 = 1
        
        rs.Update
        
        rs.Close
        Set rs = Nothing
        
        '### Update mata ganjaran terkini ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm113.L9_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then

            If Not IsNull(rs!baki_point) Then
                If IsNumeric(rs!baki_point) Then Frm113_LM_MATA_TERKINI = rs!baki_point
            End If
                
            Frm113_LM_MATA_DEDUCT = Frm113.TB1 'Mata yang dipotong / ditambah
            
            If Frm113.CB1 = 1 Then
                rs!baki_point = Frm113_LM_MATA_TERKINI + Frm113_LM_MATA_DEDUCT
            ElseIf Frm113.CB2 = 1 Then
                rs!baki_point = Frm113_LM_MATA_TERKINI - Frm113_LM_MATA_DEDUCT
            End If
            DATA_UPDATE_2 = 1
            rs.Update
            
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_UPDATE_1 = 1 And DATA_UPDATE_2 = 1 Then
        
            user = MDI_frm1.L3_Text
            
            If Frm113.CB1 Then
            
                LogAct_Memory = "[" & user & "] Penambahan mata ganjaran (" & Frm113_LM_MATA_DEDUCT & ") kepada ahli.[" & Frm113.L9_Text & "]"
                MsgBox "Sebanyak " & Frm113.TB1 & " telah berjaya ditambah kepada ahli ini.", vbInformation, "Info"
            
            End If
            
            If Frm113.CB2 Then
            
                LogAct_Memory = "[" & user & "] Penolakan mata ganjaran (" & Frm113_LM_MATA_DEDUCT & ") kepada ahli.[" & Frm113.L9_Text & "]"
                MsgBox "Sebanyak " & Frm113.TB1 & " telah berjaya dipotong dari mata ganjaran terkumpul ahli ini.", vbInformation, "Info"
            
            End If

            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
            
            Frm113.TB1 = vbNullString
            Frm113.TB2 = vbNullString
            Frm113.L16_Text = 75
            
            Frm113.TB1.SetFocus
        End If
        
    End If
    
End If
End Sub
Private Sub CMD4_Click()
'On Error Resume Next
Dim Err(5)
Dim Frm113_LM_MATA_TERKINI As Double
Dim Frm113_LM_MATA_DEDUCT As Double
Dim Frm113_LM_MATA_TERBARU As Double
Dim Frm113_LM_MATA_ASAL As Double

DATA_UPDATE_1 = 1
DATA_UPDATE_2 = 1

Frm113_LM_EMP_NO = vbNullString
Frm113_LM_MATA_TERKINI = 0
Frm113_LM_MATA_DEDUCT = 0
Frm113_LM_MATA_TERBARU = 0
Frm113_LM_MATA_ASAL = 0

If Frm113.CB1 = 0 And Frm113.CB2 = 0 Then
    x = x + 1
    Err(x) = "Sila pilih samada [Tambahan mata ganjaran] atau [Potongan mata ganjaran]."
End If
If Frm113.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih [Nama Pekerja]."
End If
If Frm113.TB1 = vbNullString Or (Frm113.TB1 <> vbNullString And Not IsNumeric(Frm113.TB1)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Jumlah mata]. Hanya NOMBOR dibenarkan dalam ruangan ini."
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
    
    If Answer = vbYes Then
    
        '### Mata asal dari urusan ini ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 71_tebus_agih_point where ID='" & Frm113.L17_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            If Not IsNull(rs!Type) Then
            
                If rs!Type = 2 Then
                
                    If Not IsNull(rs!jumlah_peroleh_point) Then
                        If IsNumeric(rs!jumlah_peroleh_point) Then Frm113_LM_MATA_ASAL = rs!jumlah_peroleh_point
                    End If
                    
                ElseIf rs!Type = 3 Then
                
                    If Not IsNull(rs!jumlah_tebus_point) Then
                        If IsNumeric(rs!jumlah_tebus_point) Then Frm113_LM_MATA_ASAL = rs!jumlah_tebus_point
                    End If
                
                End If
                
            End If
                
            
        End If
        
        rs.Close
        Set rs = Nothing
        '### Mata asal dari urusan ini ### - End
          
        '### Periksa mata terkini ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm113.L9_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!baki_point) Then
                If IsNumeric(rs!baki_point) Then Frm113_LM_MATA_TERKINI = rs!baki_point
            End If
        Else
            
            MsgBox "Tiada maklumat berkenaan dengan keahlian ini." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Kemungkinan ahli ini telah dipadamkan data dari sistem.", vbInformation, "Info"
            
            Exit Sub
            
        End If
        
        rs.Close
        Set rs = Nothing
        '### Periksa mata terkini ### - End
        
        Frm113_LM_MATA_DEDUCT = Frm113.TB1 'Mata yang dipotong / ditambah
        
        If Frm113.CB1 = 1 Then
        
            Frm113_LM_MATA_TERBARU = Frm113_LM_MATA_TERKINI - Frm113_LM_MATA_ASAL + Frm113_LM_MATA_DEDUCT
            
        ElseIf Frm113.CB2 = 1 Then
        
            Frm113_LM_MATA_TERBARU = Frm113_LM_MATA_TERKINI + Frm113_LM_MATA_ASAL - Frm113_LM_MATA_DEDUCT
            
        End If
        
        If Frm113_LM_MATA_TERBARU < 0 Then
            
            MsgBox "Mata yang ingin ditambah / dipotong adalah menyebabkan mata terkumpul ahli ini di bawah 0." & vbCrLf & _
                    "Maklumat baki terkini adalah seperti di bawah :" & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Baki mata ganjaran terkumpul : " & Frm113_LM_MATA_TERKINI & ".", vbeclamation, "Info"
                    
                    Exit Sub
        
        End If

        
        If Frm113.CBB1 <> vbNullString Then
            Frm113_LM_EMP_NO = Split(Frm113.CBB1, "  |  ")(1)
        End If
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 71_tebus_agih_point where ID='" & Frm113.L17_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Frm113.DTPicker1 <> vbNullString Then
                rs!tarikh = Frm113.DTPicker1 'Tarikh
            Else
                rs!tarikh = Null 'Tarikh
            End If
            If Frm113.L9_Text <> vbNullString Then 'No. Rujukan Pembeli
                rs!no_ahli = Frm113.L9_Text
            Else
                rs!no_ahli = Null
            End If
            If Frm113.CB1 = 1 Then
            
                rs!jumlah_peroleh_point = Frm113.TB1 'Jumlah mata yang ditambah
                rs!jumlah_tebus_point = Null 'Jumlah mata yang dipotong
                rs!Type = 2
                
            End If
            
            If Frm113.CB2 = 1 Then
            
                rs!jumlah_peroleh_point = Null 'Jumlah mata yang ditambah
                rs!jumlah_tebus_point = Frm113.TB1 'Jumlah mata yang dipotong
                rs!Type = 3
                
            End If
            rs!remarks = Frm113.TB2 'Remarks
            rs!write_timestamp2 = Now
            rs!Status = 1 '0 : Tidak aktif , 1 : Aktif
            rs!no_pekerja = Frm113_LM_EMP_NO
            DATA_UPDATE_1 = 1
            
            rs.Update
        End If
        
        rs.Close
        Set rs = Nothing
        
        '### Update mata ganjaran terkini ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm113.L9_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then

            If Not IsNull(rs!baki_point) Then
                If IsNumeric(rs!baki_point) Then Frm113_LM_MATA_TERKINI = rs!baki_point
            End If
                
            Frm113_LM_MATA_DEDUCT = Frm113.TB1 'Mata yang dipotong / ditambah
            
            If Frm113.CB1 = 1 Then
                rs!baki_point = Frm113_LM_MATA_TERKINI - Frm113_LM_MATA_ASAL + Frm113_LM_MATA_DEDUCT
            ElseIf Frm113.CB2 = 1 Then
                rs!baki_point = Frm113_LM_MATA_TERKINI + Frm113_LM_MATA_ASAL - Frm113_LM_MATA_DEDUCT
            End If
            DATA_UPDATE_2 = 1
            rs.Update
        End If
        
        rs.Close
        Set rs = Nothing
        
        If DATA_UPDATE_1 = 1 And DATA_UPDATE_2 = 1 Then
            
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & user & "] Edit data mata ganjaran (" & Frm113_LM_MATA_DEDUCT & ") kepada ahli.[" & Frm113.L9_Text & "]"
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database

            MsgBox "Data telah berjaya disimpan.", vbInformation, "Info"
            
            GM_NEXT_PREV = 2
            Call Frm113_senarai_mata_ganjaran_header
            Call Frm113_senarai_mata_ganjaran
            
            Frm113.Pic2.Visible = False
            Frm113.Pic1.Visible = True
        End If
        
    End If
    
End If
End Sub
Private Sub CMD5_Click()
'On Error Resume Next
Frm113.Pic2.Visible = False
Frm113.Pic1.Visible = True
End Sub
Private Sub Form_Load()
'On Error Resume Next
Frm113.Picture = MDI_frm1.Picture
Frm113.Pic1 = MDI_frm1.Picture
Frm113.Pic2 = MDI_frm1.Picture
Frm113.DTPicker1 = DateTime.Date
Call Frm113_initial
Call Frm113_initial2
End Sub
Private Sub Frm113_SM_cetak_penyata_mata_Click()
'On Error Resume Next
If Frm113.L9_Text <> vbNullString Then

    Report76.Sections("Section2").Controls("L1").Caption = vbNullString 'Nama
    Report76.Sections("Section2").Controls("L2").Caption = vbNullString 'No. kad pengenalan
    Report76.Sections("Section2").Controls("L3").Caption = vbNullString 'No. keahlian
    Report76.Sections("Section2").Controls("L4").Caption = vbNullString 'No. telefon
    Report76.Sections("Section2").Controls("L5").Caption = vbNullString 'Keahlian sejak
    
    Report76.Sections("Section5").Controls("L6").Caption = vbNullString 'Jumlah perolehan mata
    Report76.Sections("Section5").Controls("L7").Caption = vbNullString 'Jumlah tebusan mata
    Report76.Sections("Section5").Controls("L8").Caption = vbNullString 'Jumlah mata terkini
    Report76.Sections("Section5").Controls("L9").Caption = vbNullString 'Report
    
    Report76.Sections("Section5").Controls("L6").Caption = Frm113.L10_Text 'Jumlah perolehan mata
    Report76.Sections("Section5").Controls("L7").Caption = Frm113.L11_Text 'Jumlah tebusan mata
    Report76.Sections("Section5").Controls("L8").Caption = Frm113.L12_Text 'Jumlah mata terkini
    Report76.Sections("Section5").Controls("L9").Caption = "Penyata dikeluarkan pada " & Now 'Report
    
    '### Reset maklumat kedai ### - Start
    Report76.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
    Report76.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
    Report76.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
    Report76.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
    Report76.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
    '### Reset maklumat kedai ### - End
    
    '### Maklumat kedai ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!nama_kedai) Then Report76.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
        If Not IsNull(rs!no_pendaftaran) Then Report76.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
        If Not IsNull(rs!alamat) Then Report76.Sections("Section4").Controls("L202").Caption = rs!alamat
        If Not IsNull(rs!no_tel) Then Report76.Sections("Section4").Controls("L203").Caption = rs!no_tel
        If Not IsNull(rs!no_id_gst) Then Report76.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
    End If
    
    rs.Close
    Set rs = Nothing
    '### Maklumat kedai ### - End
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm113.L9_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        
        If Not IsNull(rs!Nama) Then Report76.Sections("Section2").Controls("L1").Caption = rs!Nama 'Nama
        If Not IsNull(rs!no_ic) Then Report76.Sections("Section2").Controls("L2").Caption = rs!no_ic 'No. kad pengenalan
        Report76.Sections("Section2").Controls("L3").Caption = Frm113.L9_Text 'No. keahlian
        If Not IsNull(rs!no_tel) Then Report76.Sections("Section2").Controls("L4").Caption = rs!no_tel 'No. telefon
        If Not IsNull(rs!tarikh) Then Report76.Sections("Section2").Controls("L5").Caption = rs!tarikh 'Keahlian sejak
    
    End If
    
    rs.Close
    Set rs = Nothing
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 71_tebus_agih_point where no_ahli='" & Frm113.L9_Text & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        Set Report76.DataSource = rs
        Report76.Show
        rs.MoveNext
    Wend
    
    'rs.Close
    Set rs = Nothing
            
End If
End Sub
Private Sub Frm113_SM_edit_Click()
'on error resume next
Frm113_LM_No_ID = vbNullString
Frm113_LM_STAFF_ID = vbNullString
DATA_PEKERJA_FOUND = 0
DATA_FOUND = 0

Call Frm113_initial

If Frm113.MSFlexGrid1 <> vbNullString Then

    If IsNumeric(Frm113.MSFlexGrid1) Then
    
        Frm113_LM_No_ID = Frm113.MSFlexGrid1.TextMatrix(Frm113.MSFlexGrid1, 2) 'No. ID
        
        If Frm113_LM_No_ID <> vbNullString Then
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 71_tebus_agih_point where ID='" & Frm113_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!Type) Then
                    
                    If rs!Type = 2 Or rs!Type = 3 Then
                        
                        Frm113.L17_Text = Frm113_LM_No_ID
                        
                        If Not IsNull(rs!tarikh) Then Frm113.DTPicker1 = rs!tarikh 'Tarikh
                        If Not IsNull(rs!Type) Then
                            If rs!Type = 2 Then
                            
                                Frm113.CB1 = 1
                                Frm113.CB2 = 0
                                If Not IsNull(rs!jumlah_peroleh_point) Then Frm113.TB1 = rs!jumlah_peroleh_point 'Jumlah mata yang ditambah
                                
                            ElseIf rs!Type = 3 Then
                            
                                Frm113.CB1 = 0
                                Frm113.CB2 = 1
                                If Not IsNull(rs!jumlah_tebus_point) Then Frm113.TB1 = rs!jumlah_tebus_point 'Jumlah mata yang dipotong
                                
                            End If
                        End If
                        
                        If Not IsNull(rs!remarks) Then Frm113.TB2 = rs!remarks 'Remarks
                        If Not IsNull(rs!no_pekerja) Then Frm113_LM_STAFF_ID = rs!no_pekerja
                        
                        DATA_FOUND = 1
                        
                    ElseIf rs!Type = 1 Then
                        
                        MsgBox "Anda tidak dibenarkan untuk edit data ini kerana pemberian / potongan mata ganjaran ini adalah dari urusan belian yang telah dibuat oleh ahli ini.", vbExclamation, "Info"
                        
                        Exit Sub
                        
                    End If
                
                Else
                    
                    MsgBox "Anda tidak dibenarkan untuk edit data ini kerana pemberian / potongan mata ganjaran ini adalah dari urusan belian yang telah dibuat oleh ahli ini.", vbExclamation, "Info"
                    
                    Exit Sub
                    
                End If
            
            End If
            
            rs.Close
            Set rs = Nothing
            
            If Frm113_LM_STAFF_ID <> vbNullString Then
            
                '### Carian Maklumat Penjual (Data Pekerja) ### - Start
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from employee where NoPekerja='" & Frm113_LM_STAFF_ID & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    Frm113_LM_MAKLUMAT_PEKERJA = rs!Samaran & "  |  " & rs!NoPekerja
                    DATA_PEKERJA_FOUND = 1
                End If
                
                rs.Close
                Set rs = Nothing
                
                If DATA_PEKERJA_FOUND = 1 Then
                    On Error GoTo Err_A:
                    Frm113.CBB1 = Frm113_LM_MAKLUMAT_PEKERJA
Restore_A:
                End If
                '### Carian Maklumat Penjual (Data Pekerja) ### - End

            End If
            
            If DATA_FOUND = 1 Then
            
                Frm113.CMD3.Visible = False
                Frm113.CMD4.Visible = True
                Frm113.CMD5.Visible = True
                Frm113.CB1.Enabled = False
                Frm113.CB2.Enabled = False
                Frm113.L18_Text.Visible = True
                
                Frm113.Pic2.Visible = True
                Frm113.Pic1.Visible = False
                
            End If
        
        End If
        
    End If
    
End If

Exit Sub
Err_A:
Frm113.CBB1.AddItem Frm113_LM_MAKLUMAT_PEKERJA
Frm113.CBB1 = Frm113_LM_MAKLUMAT_PEKERJA
Resume Restore_A:
End Sub

Private Sub L14_Text_Click()
'on error resume next
If Frm113.Pic1.Visible = False Then
    Call Frm113_initial
    Call Frm113_initial2
    
    GM_NEXT_PREV = 0
    Frm113.L7_Text = -1 'Titik Pencarian Data
    Frm113.L8_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm113.L5_Text = 0 'Paparan Page ke-xxx
    
    Call Frm113_senarai_mata_ganjaran_header
    Call Frm113_senarai_mata_ganjaran
    
    If Frm113.L13_Text <> 0 Then
        Frm113.Pic1.Visible = True
    Else
        MsgBox "Tiada rekod dijumpai.", vbInformation, "Info"
    End If
    
Else
    Frm113.Pic1.Visible = False
End If
End Sub
Private Sub L15_Text_Click()
'on error resume next
If Frm113.Pic2.Visible = False Then
    Call Frm113_initial
    Call Frm113_initial2
    
    Frm113.CB1.Enabled = True
    Frm113.CB2.Enabled = True
    Frm113.L18_Text.Visible = False
    
    Frm113.CMD3.Visible = True
    Frm113.CMD4.Visible = False
    Frm113.CMD5.Visible = False
    
    Frm113.Pic2.Visible = True
    
    Frm113.TB1.SetFocus
Else
    Frm113.Pic2.Visible = False
End If
End Sub
Private Sub L3_Text_Click()
'On Error Resume Next
GM_NEXT_PREV = 2 '0 : Next , 1 : Previous

Call frm68_senarai_pelanggan_header
Call frm68_senarai_pelanggan
            
Frm68.Show
Unload Frm113
End Sub
Private Sub MSFlexGrid1_DblClick()
'on error resume next
Frm113_LM_No_ID = vbNullString

If Frm113.MSFlexGrid1 <> vbNullString Then

    If IsNumeric(Frm113.MSFlexGrid1) Then
    
        Frm113_LM_No_ID = Frm113.MSFlexGrid1.TextMatrix(Frm113.MSFlexGrid1, 2) 'No. ID
        
        If Frm113_LM_No_ID <> vbNullString Then
            
            PopupMenu Frm113_PM_menu
            
        End If
        
    Else
    
        MsgBox "Tiada data.", vbInformation, "Info"
        
    End If
    
Else

    MsgBox "Tiada data.", vbInformation, "Info"
    
End If
End Sub
Private Sub TB2_Change()
'On Error Resume Next
Dim Frm113_LM_REMARKS As Integer

Frm113_LM_REMARKS = 0

Frm113_LM_REMARKS = Len(Frm113.TB2)

Frm113.L16_Text = 75 - Frm113_LM_REMARKS
End Sub
