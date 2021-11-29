VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm99 
   Caption         =   "Komisyen Pekerja"
   ClientHeight    =   12930
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
   Icon            =   "Frm99.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Frm99.frx":0ECA
   ScaleHeight     =   12930
   ScaleWidth      =   23760
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pic2 
      BorderStyle     =   0  'None
      Height          =   10215
      Left            =   120
      Picture         =   "Frm99.frx":29EDF
      ScaleHeight     =   10215
      ScaleWidth      =   23475
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   23475
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   9405
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
         Top             =   360
         Width           =   13785
         _ExtentX        =   24315
         _ExtentY        =   16589
         _Version        =   393216
         Rows            =   1
         Cols            =   0
         FixedCols       =   0
         BackColor       =   16777088
         ForeColor       =   0
         BackColorFixed  =   8454016
         BackColorBkg    =   12640511
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
      Begin VB.Label L7_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L7_Text"
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
         Height          =   270
         Left            =   5640
         TabIndex        =   17
         Top             =   9840
         Width           =   1785
      End
      Begin VB.Label L6_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L6_Text"
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
         Height          =   270
         Left            =   1200
         TabIndex        =   16
         Top             =   9840
         Width           =   1785
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
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   18345
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilangan :                                      Jumlah Komisyen : RM"
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
         Height          =   270
         Left            =   240
         TabIndex        =   18
         Top             =   9840
         Width           =   7425
      End
   End
   Begin VB.PictureBox Pic1 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   120
      Picture         =   "Frm99.frx":52EF4
      ScaleHeight     =   2295
      ScaleWidth      =   7155
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   7155
      Begin VB.CommandButton CMD1 
         BackColor       =   &H000080FF&
         Caption         =   "Pengiraan Komisyen"
         Height          =   345
         Left            =   2520
         MaskColor       =   &H00400000&
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1800
         Width           =   2025
      End
      Begin VB.ComboBox CBB1 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Supplier"
         Height          =   360
         ItemData        =   "Frm99.frx":7BF09
         Left            =   1800
         List            =   "Frm99.frx":7BF0B
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   600
         Width           =   5055
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   1800
         TabIndex        =   9
         Top             =   960
         Width           =   5055
         _ExtentX        =   8916
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
         Format          =   136773632
         CurrentDate     =   41561
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   360
         Left            =   1800
         TabIndex        =   10
         Top             =   1320
         Width           =   5055
         _ExtentX        =   8916
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
         Format          =   136773632
         CurrentDate     =   41561
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pekerja *"
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   240
         TabIndex        =   14
         Top             =   615
         Width           =   2295
      End
      Begin VB.Label Label79 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Akhir *"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1365
         Width           =   1695
      End
      Begin VB.Label Label80 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Mula *  "
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1005
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila buat tetapan pengiraan komisyen."
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
         TabIndex        =   8
         Top             =   240
         Width           =   5385
      End
   End
   Begin VB.Timer Tmr1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kedai Emas Sri Harmoni"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   4680
      TabIndex        =   20
      Top             =   480
      Width           =   14295
   End
   Begin VB.Label Label67 
      BackColor       =   &H00000000&
      Caption         =   "Label7"
      ForeColor       =   &H00FFFFFF&
      Height          =   45
      Left            =   0
      TabIndex        =   19
      Top             =   2040
      Width           =   24000
   End
   Begin VB.Label L4_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tetapan Pengiraan Komisyen"
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
      MouseIcon       =   "Frm99.frx":7BF0D
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2160
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
      Left            =   21480
      MouseIcon       =   "Frm99.frx":7C217
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label L1_Text 
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
      Left            =   21720
      TabIndex        =   1
      Top             =   1320
      Width           =   2100
   End
   Begin VB.Label L2_Text 
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
      Left            =   21705
      TabIndex        =   0
      Top             =   1635
      Width           =   2100
   End
   Begin VB.Label Label85 
      BackColor       =   &H00FFFFFF&
      Height          =   2070
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   24000
   End
   Begin VB.Menu Frm99_PM_Menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Frm99_SM_cetak_penyata 
         Caption         =   "Cetak penyata komisyen"
      End
   End
End
Attribute VB_Name = "Frm99"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD1_Click()
'on error resume next
Dim TM As Date
Dim TA As Date

If Frm99.CBB1 = vbNullString Then
    MsgBox "Sila pilih nama pekerja.", vbExclamation, "Info"
    Exit Sub
End If

TM = Frm99.DTPicker1 'Tarikh Mula
TA = Frm99.DTPicker2 'Tarikh Akhir
Frm99_LM_NAMA = Split(Frm99.CBB1, " -> ")(1)


Note = "Adakah anda ingin mengira jumlah komisyen bagi " & Frm99_LM_NAMA & " dari " & TM & " hingga " & TA & "?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    Call Frm99_senarai_komisyen_header
    Call Frm99_senarai_komisyen
End If
End Sub
Private Sub Frm99_SM_cetak_penyata_Click()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim Frm99_LM_KOMISYEN As Double

Frm99_LM_NO_STAFF = vbNullString
Frm99_LM_KOMISYEN = 0

TM = Frm99.DTPicker1 'Tarikh Mula
TA = Frm99.DTPicker2 'Tarikh Akhir
Frm99_LM_NO_STAFF = Split(Frm99.CBB1, " -> ")(0)
Frm99_LM_NAMA_STAFF = Split(Frm99.CBB1, " -> ")(1)

If Frm99_LM_NO_STAFF <> vbNullString Then

    Report55.Sections("Section4").Controls("L1").Caption = vbNullString
    Report55.Sections("Section4").Controls("L2").Caption = vbNullString
    Report55.Sections("Section5").Controls("L3").Caption = Frm99.L6_Text
    Report55.Sections("Section5").Controls("L4").Caption = Frm99.L7_Text
    
    Set rs = New ADODB.Recordset
    Call Main
    rs.Open "select * from employee where NoPekerja='" & Frm99_LM_NO_STAFF & "'", cn, adOpenKeyset, adLockOptimistic

    If Not rs.EOF Then
        If Not IsNull(rs!Nama) Then Report55.Sections("Section4").Controls("L1").Caption = rs!Nama
    End If
    
    rs.Close
    Set rs = Nothing

    Report55.Sections("Section4").Controls("L2").Caption = TM & " hingga " & TA
    
    '### Paparan Penyata ### - Start
    Set rs = New ADODB.Recordset
    Call Main
    rs.Open "select * from 23_senarai_jualan where no_pekerja='" & Frm99_LM_NO_STAFF & "' AND tarikh BETWEEN '" & TM & "'  AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        Set Report55.DataSource = rs
        Report55.Show
        rs.MoveNext
    Wend
    
    'rs.Close
    Set rs = Nothing
    '### Paparan Penyata ### - End
    
End If
End Sub
Private Sub L3_Text_Click()
'on error resume next
Frm30.Show
Unload Frm99
End Sub
Private Sub L4_Text_Click()
'on error resume next
If Frm99.Pic1.Visible = False Then
    Call Frm99_initial_setting
    
    Frm99.Pic1.Visible = True
Else
    Frm99.Pic1.Visible = False
End If
End Sub
Private Sub MSFlexGrid1_DblClick()
'On Error Resume Next
If Frm99.MSFlexGrid1 <> vbNullString Then
    PopupMenu Frm99_PM_Menu1
Else
    MsgBox "Tiada Data.", vbExclamation, "Info"
End If
End Sub

Private Sub Tmr1_Timer()
'on error resume next
Frm99.L1_Text = DateTime.Date
Frm99.L2_Text = DateTime.Time$
End Sub
