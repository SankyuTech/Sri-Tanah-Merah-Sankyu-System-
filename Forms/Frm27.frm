VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm27 
   BackColor       =   &H80000003&
   Caption         =   "Maklumat Agen Dropship"
   ClientHeight    =   12015
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12915
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
   Icon            =   "Frm27.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12015
   ScaleWidth      =   12915
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CMD2 
      BackColor       =   &H80000004&
      Caption         =   "Padam Maklumat Pembeli"
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
      Left            =   3840
      MaskColor       =   &H00400000&
      Picture         =   "Frm27.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Padam maklumat pembeli"
      Top             =   10200
      Width           =   4305
   End
   Begin VB.CommandButton CMD21 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   10440
      MouseIcon       =   "Frm27.frx":1874
      MousePointer    =   99  'Custom
      Picture         =   "Frm27.frx":1B7E
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Tutup senarai ini."
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton CMD22 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   11640
      MouseIcon       =   "Frm27.frx":2C48
      MousePointer    =   99  'Custom
      Picture         =   "Frm27.frx":2F52
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Tutup senarai ini."
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton CMD1 
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
      Left            =   7920
      MaskColor       =   &H00400000&
      Picture         =   "Frm27.frx":401C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Carian Maklumat Pembeli"
      Top             =   1440
      Width           =   2145
   End
   Begin VB.TextBox TB1 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2145
      TabIndex        =   3
      Text            =   "TB1"
      Top             =   1530
      Width           =   5700
   End
   Begin VB.CommandButton CMD4 
      BackColor       =   &H80000004&
      Caption         =   "Carian / Pendaftaran"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   13320
      MaskColor       =   &H00400000&
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Carian maklumat pembeli / Pendaftaran maklumat pembeli."
      Top             =   8040
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.CommandButton CMD3 
      BackColor       =   &H80000004&
      Caption         =   "Kembali Ke Menu Sebelum"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4920
      MaskColor       =   &H00400000&
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Kembali Ke Menu Sebelum"
      Top             =   11160
      Width           =   2145
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   4860
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   8573
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Shape Shape3 
      Height          =   2535
      Left            =   120
      Top             =   8520
      Width           =   12615
   End
   Begin VB.Label L4_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L4_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2160
      TabIndex        =   31
      Top             =   9600
      Width           =   8835
   End
   Begin VB.Label L3_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L3_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2160
      TabIndex        =   30
      Top             =   9360
      Width           =   8835
   End
   Begin VB.Label L2_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L2_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2160
      TabIndex        =   29
      Top             =   9120
      Width           =   8835
   End
   Begin VB.Label L1_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L1_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2160
      TabIndex        =   28
      Top             =   8880
      Width           =   8835
   End
   Begin VB.Label L5_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L5_Text"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2160
      TabIndex        =   27
      Top             =   9840
      Width           =   8835
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Kad Pengenalan :"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   26
      Top             =   9120
      Width           =   2010
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nama :"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   25
      Top             =   8880
      Width           =   2010
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Telefon :"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   24
      Top             =   9360
      Width           =   2010
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail :"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   23
      Top             =   9600
      Width           =   2010
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Pelanggan :"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   22
      Top             =   9840
      Width           =   2010
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Maklumat pelanggan yang telah dipilih."
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
      Height          =   315
      Left            =   240
      TabIndex        =   21
      Top             =   8640
      Width           =   4650
   End
   Begin VB.Label Label19 
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
      Left            =   8040
      TabIndex        =   20
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Senarai Agen Dropship."
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
      Height          =   315
      Left            =   240
      TabIndex        =   19
      Top             =   2400
      Width           =   4650
   End
   Begin VB.Label L70_Text 
      BackColor       =   &H8000000C&
      Caption         =   "L70_Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   6120
      TabIndex        =   18
      Top             =   7800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label L69_Text 
      BackColor       =   &H8000000C&
      Caption         =   "L69_Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   7080
      TabIndex        =   17
      Top             =   7800
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
      Left            =   9360
      TabIndex        =   16
      Top             =   7560
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
      Left            =   9960
      TabIndex        =   15
      Top             =   7560
      Width           =   615
   End
   Begin VB.Label Label21 
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
      TabIndex        =   14
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label L71_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L71_Text"
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
      TabIndex        =   13
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label L72_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L72_Text"
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
      TabIndex        =   12
      Top             =   7800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   2175
      Left            =   120
      Top             =   120
      Width           =   12735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "**Hanya maklumat ahli yang sudah berdaftar dengan sistem sahaja akan dipaparkan."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   7395
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"Frm27.frx":49C6
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
      Height          =   525
      Left            =   360
      TabIndex        =   7
      Top             =   600
      Width           =   10515
   End
   Begin VB.Label Label45 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Keyword Carian :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Top             =   1540
      Width           =   1785
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Carian data agen dropship"
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
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   6555
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "** Anda perlu menutup menu ini dahulu sebelum boleh ke menu seterusnya."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   11640
      Width           =   7095
   End
   Begin VB.Menu frm27_pm_menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu frm27_sm_pilih 
         Caption         =   "Pilih Maklumat Agen Ini"
      End
   End
End
Attribute VB_Name = "Frm27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD1_Click()
'on error resume next
Call frm27_periksa_carian
End Sub
Private Sub CMD2_Click()
'on error resume next
Note = "Padamkan semua maklumat pembeli ini ?" & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
ElseIf Answer = vbYes Then
    If MDI_frm1.L5_Text = 4 Then Frm84.L29_Text = UCase(Frm27.L1_Text) 'Nama Agen Dropship
    'Frm84.L29_Text = UCase(Frm27.L1_Text) 'Nama Pembeli
    Frm27.TB1 = vbNullString
    Frm27.L1_Text = vbNullString
    Frm27.L2_Text = vbNullString
    Frm27.L3_Text = vbNullString
    Frm27.L4_Text = vbNullString
    Frm27.L5_Text = vbNullString
    
    'If Frm84.Visible = True Then Frm84.L29_Text = vbNullString 'Nama
End If
End Sub

Private Sub CMD21_Click()
'on error resume next
Dim frm27_LM_CURR_PAGE As Double
Dim frm27_LM_TOTAL_PAGE As Double

frm27_LM_CURR_PAGE = 0
frm27_LM_TOTAL_PAGE = 0

If Frm27.L67_Text <> vbNullString And IsNumeric(Frm27.L67_Text) Then
    If Frm27.L68_Text <> vbNullString And IsNumeric(Frm27.L68_Text) Then
        frm27_LM_CURR_PAGE = Frm27.L67_Text
        frm27_LM_TOTAL_PAGE = Frm27.L68_Text
        
        If frm27_LM_CURR_PAGE <> 1 And frm27_LM_CURR_PAGE <> 0 Then
        
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
                    
            Call frm27_senarai_dropship_header
            Call frm27_senarai_dropship
                    
        End If

    End If
End If
End Sub
Private Sub CMD22_Click()
'on error resume next
Dim frm27_LM_CURR_PAGE As Double
Dim frm27_LM_TOTAL_PAGE As Double

frm27_LM_CURR_PAGE = 0
frm27_LM_TOTAL_PAGE = 0

If Frm27.L67_Text <> vbNullString And IsNumeric(Frm27.L67_Text) Then
    If Frm27.L68_Text <> vbNullString And IsNumeric(Frm27.L68_Text) Then
        frm27_LM_CURR_PAGE = Frm27.L67_Text
        frm27_LM_TOTAL_PAGE = Frm27.L68_Text
        
        If frm27_LM_CURR_PAGE < frm27_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call frm27_senarai_dropship_header
            Call frm27_senarai_dropship
            
        End If
    End If
End If
End Sub

Private Sub CMD3_Click()
'on error resume next
Frm27.Hide
End Sub
Private Sub CMD4_Click()
'on error resume next

'Data Pelanggan
'---------------------
'0:  Pendaftaran Biasa
'1 : Jualan Gold Bar
'2 : Buyback Gold Bar
'3:  Jualan BK
'4:  Buyback BK
'5:  Ansuran
'6:  Servis
'7:  Tempahan

'Data Agen Drophip
'--------------------
'20 : Jualan

If Frm84.Visible = True Then Frm68.L15_Text = 20

Frm68.L36_Text = 2 '0 : Terus dari menu data pelanggan , 1 : Data pembeli , 2 : Data agen dropship

Frm68.Show
Frm84.Hide
Frm27.Hide
End Sub

Private Sub frm27_sm_pilih_Click()
'on error resume next
DATA_FOUND = 0

If IsNumeric(Frm27.LV1.SelectedItem.Index) Then
    
    frm27_LM_No_ID = Frm27.LV1.ListItems(Frm27.LV1.SelectedItem.Index)
    
    If frm27_LM_No_ID <> vbNullString Then
        
        Frm27.L1_Text = vbNullString
        Frm27.L2_Text = vbNullString
        Frm27.L3_Text = vbNullString
        Frm27.L4_Text = vbNullString
        Frm27.L5_Text = vbNullString
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where ID='" & frm27_LM_No_ID & "' AND dropship = 1", cn, adOpenKeyset, adLockOptimistic
    
        If Not rs.EOF Then

            If Not IsNull(rs!Nama) Then Frm27.L1_Text = rs!Nama 'Nama
            If Not IsNull(rs!no_ic) Then Frm27.L2_Text = rs!no_ic 'No. IC
            If Not IsNull(rs!no_tel) Then Frm27.L3_Text = rs!no_tel 'No. Telefon
            If Not IsNull(rs!Email) Then Frm27.L4_Text = rs!Email 'E-mail
            If Not IsNull(rs!no_pelanggan) Then Frm27.L5_Text = rs!no_pelanggan 'No. Customer
        
        End If
        
        rs.Close
        Set rs = Nothing
        
    Else
    
        MsgBox "Tiada Data.", vbInformation, "Info"
        
    End If
    
Else

    MsgBox "Tiada Data.", vbInformation, "Info"
    
End If
End Sub

Private Sub L1_Text_Change()
'on error resume next
If MDI_frm1.L5_Text = "4" Or MDI_frm1.L5_Text = "5" Then Frm84.L29_Text = UCase(Frm27.L1_Text) 'Nama Agen Dropship
'If Frm84.Visible = True Then Frm84.L29_Text = UCase(Frm27.L1_Text) 'Nama Pembeli
'If Frm83.Visible = True Then Frm83.L37_Text = UCase(Frm27.L1_Text) 'Nama Pembeli
'If Frm87.Visible = True Then Frm87.L6_Text = UCase(Frm27.L1_Text) 'Nama Pembeli
'If Frm93.Visible = True Then Frm93.L36_Text = UCase(Frm27.L1_Text) 'Nama Pembeli
'If Frm92.Visible = True Then Frm92.L52_Text = UCase(Frm27.L1_Text) 'Nama Pembeli
End Sub

Private Sub LV1_DblClick()
'on error resume next
frm27_LM_No_ID = vbNullString

If IsNumeric(Frm27.LV1.SelectedItem.Index) Then
    
    frm27_LM_No_ID = Frm27.LV1.SelectedItem.Index
    
    If frm27_LM_No_ID <> vbNullString Then

        PopupMenu frm27_pm_menu
    
    Else
    
        MsgBox "Tiada Data.", vbInformation, "Info"
        
    End If
    
Else

    MsgBox "Tiada Data.", vbInformation, "Info"
    
End If
End Sub

Private Sub TB1_KeyPress(KeyAscii As Integer)
'on error resume next
If KeyAscii = 13 Then
    
    Call frm27_periksa_carian

End If
End Sub
