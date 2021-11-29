VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm134 
   Caption         =   "Report Stok Setiap Dulang"
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
   Begin VB.CommandButton CMD21 
      Caption         =   "Back"
      Height          =   810
      Left            =   10560
      MouseIcon       =   "frm134.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frm134.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Tutup senarai ini."
      Top             =   10800
      Width           =   1095
   End
   Begin VB.CommandButton CMD22 
      Caption         =   "Next"
      Height          =   810
      Left            =   11760
      MouseIcon       =   "frm134.frx":13D4
      MousePointer    =   99  'Custom
      Picture         =   "frm134.frx":16DE
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Tutup senarai ini."
      Top             =   10800
      Width           =   1095
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   10380
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   18309
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
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label L1 
      BackStyle       =   0  'Transparent
      Caption         =   "L1"
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
      TabIndex        =   9
      Top             =   10800
      Width           =   12615
   End
   Begin VB.Label L70_Text 
      BackColor       =   &H8000000C&
      Caption         =   "L70_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   6240
      TabIndex        =   8
      Top             =   11040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label L69_Text 
      BackColor       =   &H8000000C&
      Caption         =   "L69_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   7200
      TabIndex        =   7
      Top             =   11040
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
      Left            =   9480
      TabIndex        =   6
      Top             =   10800
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
      Left            =   10080
      TabIndex        =   5
      Top             =   10800
      Width           =   615
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
      Left            =   8160
      TabIndex        =   4
      Top             =   10800
      Width           =   2295
   End
   Begin VB.Label L14_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "Report stok mengikut dulang."
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
      TabIndex        =   3
      Top             =   120
      Width           =   12615
   End
   Begin VB.Menu frm134_pm_menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu frm134_sm_cetak_penyata 
         Caption         =   "Cetak Penyata"
      End
   End
End
Attribute VB_Name = "frm134"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD21_Click()
'on error resume next
Dim frm134_LM_CURR_PAGE As Double
Dim frm134_LM_TOTAL_PAGE As Double

frm134_LM_CURR_PAGE = 0
frm134_LM_TOTAL_PAGE = 0

If frm134.L67_Text <> vbNullString And IsNumeric(frm134.L67_Text) Then
    If frm134.L68_Text <> vbNullString And IsNumeric(frm134.L68_Text) Then
        frm134_LM_CURR_PAGE = frm134.L67_Text
        frm134_LM_TOTAL_PAGE = frm134.L68_Text
        
        If frm134_LM_CURR_PAGE <> 1 And frm134_LM_CURR_PAGE <> 0 Then
        
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
                    
            Call frm134_report_stok_header
            Call frm134_report_stok
                    
        End If

    End If
End If
End Sub
Private Sub CMD22_Click()
'on error resume next
Dim frm134_LM_CURR_PAGE As Double
Dim frm134_LM_TOTAL_PAGE As Double

frm134_LM_CURR_PAGE = 0
frm134_LM_TOTAL_PAGE = 0

If frm134.L67_Text <> vbNullString And IsNumeric(frm134.L67_Text) Then
    If frm134.L68_Text <> vbNullString And IsNumeric(frm134.L68_Text) Then
        frm134_LM_CURR_PAGE = frm134.L67_Text
        frm134_LM_TOTAL_PAGE = frm134.L68_Text
        
        If frm134_LM_CURR_PAGE < frm134_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call frm134_report_stok_header
            Call frm134_report_stok
            
        End If
    End If
End If
End Sub

Private Sub frm134_sm_cetak_penyata_Click()
'on error resume next
DATA_FOUND = 0

If IsNumeric(frm134.LV1.SelectedItem.Index) Then
    
    frm134_LM_No_ID = frm134.LV1.ListItems(frm134.LV1.SelectedItem.Index)
    
    If frm134_LM_No_ID <> vbNullString Then
        
        G_PREVIEW = 1
        Call frm134_cetak_penyata_dulang
        
    Else
    
        MsgBox "Tiada Data.", vbInformation, "Info"
        
    End If
    
Else

    MsgBox "Tiada Data.", vbInformation, "Info"
    
End If
End Sub

Private Sub LV1_DblClick()
'on error resume next
frm134_LM_No_ID = vbNullString

If IsNumeric(frm134.LV1.SelectedItem.Index) Then
    
    frm134_LM_No_ID = frm134.LV1.SelectedItem.Index
    
    If frm134_LM_No_ID <> vbNullString Then

        PopupMenu frm134_pm_menu
    
    Else
    
        MsgBox "Tiada Data.", vbInformation, "Info"
        
    End If
    
Else

    MsgBox "Tiada Data.", vbInformation, "Info"
    
End If
End Sub
