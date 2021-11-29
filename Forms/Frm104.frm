VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm104 
   Caption         =   "Penyata Untung Rugi (Restock)"
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
   Icon            =   "Frm104.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13035
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic4 
      BorderStyle     =   0  'None
      Height          =   10215
      Left            =   4800
      ScaleHeight     =   10215
      ScaleWidth      =   23415
      TabIndex        =   14
      Top             =   2880
      Visible         =   0   'False
      Width           =   23415
      Begin VB.CommandButton CMD2 
         BackColor       =   &H000080FF&
         Caption         =   "Cetak Penyata"
         Height          =   405
         Left            =   120
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm104.frx":0ECA
         MousePointer    =   99  'Custom
         TabIndex        =   40
         Top             =   7320
         Width           =   2025
      End
      Begin VB.TextBox TB3 
         Height          =   285
         Left            =   3000
         TabIndex        =   25
         Text            =   "TB3"
         Top             =   3960
         Width           =   1300
      End
      Begin VB.TextBox TB2 
         Height          =   285
         Left            =   3000
         TabIndex        =   24
         Text            =   "TB2"
         Top             =   3660
         Width           =   1300
      End
      Begin VB.TextBox TB1 
         Height          =   285
         Left            =   3000
         TabIndex        =   23
         Text            =   "TB1"
         Top             =   3360
         Width           =   1300
      End
      Begin VB.Label L21_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L21_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2700
         TabIndex        =   50
         Top             =   2060
         Width           =   1305
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm104.frx":11D4
         ForeColor       =   &H00000000&
         Height          =   1245
         Left            =   240
         TabIndex        =   49
         Top             =   1680
         Width           =   2010
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "** Upah sudah termasuk dalam harga jualan"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3840
         TabIndex        =   48
         Top             =   2505
         Width           =   5250
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   ":        :         : RM  : RM  : RM   : RM    "
         ForeColor       =   &H00000000&
         Height          =   1155
         Left            =   2280
         TabIndex        =   47
         Top             =   1680
         Width           =   450
      End
      Begin VB.Label L12_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L12_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   46
         Top             =   1680
         Width           =   2730
      End
      Begin VB.Label L13_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L13_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   45
         Top             =   1870
         Width           =   2730
      End
      Begin VB.Label L14_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L14_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2700
         TabIndex        =   44
         Top             =   2250
         Width           =   1305
      End
      Begin VB.Label L15_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L15_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2700
         TabIndex        =   43
         Top             =   2450
         Width           =   1305
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "** Upah modal bagi barangan yang telah terjual"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3840
         TabIndex        =   42
         Top             =   2700
         Width           =   5250
      End
      Begin VB.Label L16_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L16_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2700
         TabIndex        =   41
         Top             =   2660
         Width           =   1305
      End
      Begin VB.Label L20_text 
         BackStyle       =   0  'Transparent
         Caption         =   "Report generated on 2015/01/01 00:00:00"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   39
         Top             =   7920
         Width           =   8490
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Perhatian."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   38
         Top             =   6120
         Width           =   6210
      End
      Begin VB.Label L19_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L19_Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   37
         Top             =   5640
         Width           =   10170
      End
      Begin VB.Label L18_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L18_Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   36
         Top             =   5280
         Width           =   10170
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm104.frx":1270
         ForeColor       =   &H00000000&
         Height          =   1365
         Left            =   120
         TabIndex        =   35
         Top             =   6360
         Width           =   6930
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Analisa untung rugi restock barang kemas."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   34
         Top             =   4920
         Width           =   6210
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "** Sila masukkan upah bagi barang kemas ini yang ditetapkan oleh pihak pembekal."
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4440
         TabIndex        =   33
         Top             =   4040
         Width           =   6330
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "** Sila masukkan harga semasa (modal) yang ditetapkan oleh pihak pembekal."
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4440
         TabIndex        =   32
         Top             =   3720
         Width           =   6330
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "** Sila masukkan berat barang kemas yang ingin di restock."
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4440
         TabIndex        =   31
         Top             =   3360
         Width           =   5250
      End
      Begin VB.Shape Shape1 
         Height          =   375
         Left            =   180
         Top             =   4380
         Width           =   4695
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
         Height          =   285
         Left            =   3000
         TabIndex        =   30
         Top             =   4440
         Width           =   1665
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Harga Restock                 RM :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   29
         Top             =   4440
         Width           =   2730
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Upah (Tanpa GST)                       RM :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   28
         Top             =   3960
         Width           =   2730
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Semasa (Dari Pembekal)RM/g :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   27
         Top             =   3675
         Width           =   2730
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Berat Restock                                 g :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   26
         Top             =   3375
         Width           =   2730
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Belian stok baru (Restock)."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   22
         Top             =   3000
         Width           =   6210
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Barangan yang terjual."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   6210
      End
      Begin VB.Label L11_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L11_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1320
         TabIndex        =   20
         Top             =   1000
         Width           =   2730
      End
      Begin VB.Label L10_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L10_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1320
         TabIndex        =   19
         Top             =   800
         Width           =   2730
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   ":    :    :       "
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   1200
         TabIndex        =   18
         Top             =   600
         Width           =   210
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Mula    Hingga         Purity           "
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Tetapan analisa."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   6210
      End
      Begin VB.Label L9_Text 
         BackStyle       =   0  'Transparent
         Caption         =   "L9_Text"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   600
         Width           =   2730
      End
   End
   Begin VB.PictureBox Pic1 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   120
      ScaleHeight     =   2655
      ScaleWidth      =   5775
      TabIndex        =   1
      Top             =   240
      Width           =   5775
      Begin VB.CommandButton CMD1 
         BackColor       =   &H000080FF&
         Caption         =   "Paparan analisa"
         Height          =   405
         Left            =   1800
         MaskColor       =   &H00400000&
         MouseIcon       =   "Frm104.frx":1414
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   2040
         Width           =   2025
      End
      Begin VB.ComboBox CBB1 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "Frm104.frx":171E
         Left            =   1545
         List            =   "Frm104.frx":1720
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   510
         Width           =   4005
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1545
         TabIndex        =   4
         Top             =   840
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   16744576
         Format          =   62390272
         CurrentDate     =   41561
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   1545
         TabIndex        =   5
         Top             =   1170
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   16744576
         Format          =   110362624
         CurrentDate     =   41561
      End
      Begin VB.Label L7_Text 
         Caption         =   "L7_Text"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3960
         TabIndex        =   13
         Top             =   2280
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label L6_Text 
         Caption         =   "L6_Text"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4680
         TabIndex        =   12
         Top             =   2040
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label L5_Text 
         Caption         =   "L5_Text"
         ForeColor       =   &H00000000&
         Height          =   200
         Left            =   3960
         TabIndex        =   11
         Top             =   2040
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Purity"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   550
         Width           =   1995
      End
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   "Hingga"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1200
         Width           =   1995
      End
      Begin VB.Label Label62 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarikh Mula "
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   885
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "* Sistem akan mencari senarai barang yang telah dijual mengikut tetapan yang telah dibuat."
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   5415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sila buat pilihan jenis mutu (purity) dan tarikh jualan."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.Label L4_Text 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tetapan Analisa"
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
      MouseIcon       =   "Frm104.frx":1722
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "Frm104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD1_Click()
'On Error Resume Next
Dim Err(3)

If Frm104.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila buat pilihan purity."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Sistem mungkin akan mengambil sedikit masa untuk menganalisa senarai jualan dari pilihan purity dan tempoh report." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Teruskan ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
        
        Call Frm104_transfer_list_jualan
            
    End If
End If
End Sub
Private Sub CMD2_Click()
'On Error Resume Next
Dim Err(3)

If Frm104.TB1 = vbNullString Or (Frm104.TB1 <> vbNullString And Not IsNumeric(Frm104.TB1)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Berat Restock]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm104.TB2 = vbNullString Or (Frm104.TB2 <> vbNullString And Not IsNumeric(Frm104.TB2)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Harga Semasa]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm104.TB3 = vbNullString Or (Frm104.TB3 <> vbNullString And Not IsNumeric(Frm104.TB3)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Upah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah anda ingin cetak penyata analisa untung rugi restock ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then

        Call Frm104_penyata_restock
        
    End If
End If
End Sub
Private Sub Form_Load()
'on error resume next
Frm104_LM_FOUND = 0

Frm104.CBB1.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database order by kadar_tukaran_9999 DESC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!Kod_Metal_Purity) Then
        Frm104.CBB1.AddItem rs!Kod_Metal_Purity
        If Frm104_LM_FOUND = 0 Then
            Frm104_LM_PURITY = rs!Kod_Metal_Purity
            Frm104_LM_FOUND = 1
        End If
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

If Frm104_LM_FOUND = 1 Then
    Frm104.CBB1 = Frm104_LM_PURITY
End If

Call Frm104_initial_setting
Frm104.Pic1.Visible = True

Frm104.DTPicker1 = DateTime.Date$
Frm104.DTPicker2 = DateTime.Date$
End Sub
Private Sub L13_Text_Change()
'on error resume next
If IsNumeric(Frm104.L13_Text) Then
    Frm104.TB1 = Frm104.L13_Text
    
    Call Frm104_analisa_untung_rugi_restock
Else
    Frm104.TB1 = "0.00"
End If
End Sub
Private Sub L14_Text_Change()
'On Error Resume Next
Call Frm104_analisa_untung_rugi_restock
End Sub
Private Sub L16_Text_Change()
'on error resume next
If IsNumeric(Frm104.L16_Text) Then
    Frm104.TB3 = Frm104.L16_Text
Else
    Frm104.TB3 = "0.00"
End If
End Sub
Private Sub L17_Text_Change()
'On Error Resume Next
Call Frm104_analisa_untung_rugi_restock
End Sub
Private Sub L4_Text_Click()
'on error resume next
If Frm104.Pic1.Visible = False Then
    Call Frm104_initial_setting
    
    Frm104.Pic1.Visible = True
Else
    'Frm104.Pic1.Visible = False
End If
End Sub
Private Sub TB1_Change()
'On Error Resume Next
Call Frm104_harga_restock
Call Frm104_analisa_untung_rugi_restock
End Sub
Private Sub TB2_Change()
'On Error Resume Next
Call Frm104_harga_restock
End Sub
Private Sub TB3_Change()
'On Error Resume Next
Call Frm104_harga_restock
End Sub
