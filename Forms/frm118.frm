VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm118 
   Caption         =   "Invoice & Voucher (Urusan dengan supplier / agen)"
   ClientHeight    =   13035
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13035
   ScaleWidth      =   23760
   WindowState     =   2  'Maximized
   Begin VB.TextBox TB8 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   50
      Text            =   "TB8"
      Top             =   7320
      Width           =   1500
   End
   Begin VB.CheckBox CB9 
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
      Left            =   1920
      TabIndex        =   44
      Top             =   3765
      Width           =   200
   End
   Begin VB.CheckBox CB8 
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
      Left            =   1920
      TabIndex        =   43
      Top             =   3540
      Width           =   200
   End
   Begin VB.TextBox TB7 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2400
      TabIndex        =   42
      Text            =   "TB7"
      Top             =   4080
      Width           =   1500
   End
   Begin VB.ComboBox CBB2 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frm118.frx":0000
      Left            =   1560
      List            =   "frm118.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   2280
      Width           =   4815
   End
   Begin VB.TextBox TB6 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "TB6"
      Top             =   6960
      Width           =   1500
   End
   Begin VB.TextBox TB5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7200
      TabIndex        =   33
      Text            =   "TB5"
      Top             =   5040
      Width           =   1500
   End
   Begin VB.CheckBox CB7 
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
      Left            =   2880
      TabIndex        =   30
      Top             =   5685
      Width           =   200
   End
   Begin VB.CheckBox CB6 
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
      Left            =   2880
      TabIndex        =   29
      Top             =   5460
      Width           =   200
   End
   Begin VB.TextBox TB4 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "TB4"
      Top             =   6300
      Width           =   1500
   End
   Begin VB.TextBox TB3 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "TB3"
      Top             =   6000
      Width           =   1500
   End
   Begin VB.CommandButton CMD3 
      BackColor       =   &H8000000C&
      Caption         =   "Batal"
      Height          =   360
      Left            =   4440
      MouseIcon       =   "frm118.frx":0004
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Sila pastikan semua data telah dimasukkan dengan betul sebelum masukkan data ini ke dalam senarai jualan."
      Top             =   8280
      Width           =   3015
   End
   Begin VB.CommandButton CMD2 
      BackColor       =   &H8000000C&
      Caption         =   "Simpan Data"
      Height          =   360
      Left            =   1320
      MouseIcon       =   "frm118.frx":030E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8280
      Width           =   3015
   End
   Begin VB.CommandButton CMD1 
      Caption         =   "Simpan Data"
      Height          =   360
      Left            =   2760
      MouseIcon       =   "frm118.frx":0618
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   8280
      Width           =   3015
   End
   Begin VB.CheckBox CB5 
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
      Left            =   3240
      TabIndex        =   19
      Top             =   7965
      Width           =   200
   End
   Begin VB.CheckBox CB3 
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
      Left            =   3240
      TabIndex        =   16
      Top             =   7500
      Width           =   200
   End
   Begin VB.CheckBox CB4 
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
      Left            =   3240
      TabIndex        =   15
      Top             =   7725
      Width           =   200
   End
   Begin VB.TextBox TB2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3120
      TabIndex        =   13
      Text            =   "TB2"
      Top             =   5040
      Width           =   1500
   End
   Begin VB.ComboBox CBB1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frm118.frx":0922
      Left            =   1560
      List            =   "frm118.frx":0924
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1920
      Width           =   4815
   End
   Begin VB.TextBox TB1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1560
      TabIndex        =   6
      Text            =   "TB1"
      Top             =   1560
      Width           =   1500
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
      Left            =   1200
      TabIndex        =   1
      Top             =   630
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
      Left            =   2520
      TabIndex        =   0
      Top             =   630
      Width           =   200
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1560
      TabIndex        =   10
      Top             =   2640
      Width           =   4815
      _ExtentX        =   8493
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
      Format          =   415825920
      CurrentDate     =   41561
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Berat Bayaran (g) :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4080
      TabIndex        =   51
      Top             =   7320
      Width           =   2985
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Tujuan bayaran :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   49
      Top             =   3480
      Width           =   3465
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Umum                         Bayaran belian stok emas"
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   2160
      TabIndex        =   48
      Top             =   3495
      Width           =   2625
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Contoh : Bayaran bagi UPAH"
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2880
      TabIndex        =   47
      Top             =   3480
      Width           =   5145
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Semasa : RM /g"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   46
      Top             =   4080
      Width           =   2985
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Nilaian harga semasa yang diberikan oleh pihak pembekal."
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   4080
      TabIndex        =   45
      Top             =   4080
      Width           =   5145
   End
   Begin VB.Label Label109 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Pekerja :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   -480
      TabIndex        =   41
      Top             =   2280
      Width           =   2025
   End
   Begin VB.Label L3_Text 
      BackColor       =   &H8000000C&
      Caption         =   "L3_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9960
      TabIndex        =   39
      Top             =   6360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label L2_Text 
      BackColor       =   &H8000000C&
      Caption         =   "L2_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9960
      TabIndex        =   38
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label L1_Text 
      BackColor       =   &H8000000C&
      Caption         =   "L1_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9960
      TabIndex        =   37
      Top             =   5640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Keseluruhan (RM) :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4080
      TabIndex        =   36
      Top             =   6975
      Width           =   2985
   End
   Begin VB.Shape Shape2 
      Height          =   2295
      Left            =   240
      Top             =   4635
      Width           =   8655
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah (RM) :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4080
      TabIndex        =   34
      Top             =   5040
      Width           =   2985
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Zero Rated (ZR)"
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
      Left            =   5880
      TabIndex        =   32
      Top             =   4725
      Width           =   3465
   End
   Begin VB.Shape Shape1 
      Height          =   2175
      Left            =   360
      Top             =   4680
      Width           =   5295
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Rated         Standard Rated (Inclusive)"
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   3120
      TabIndex        =   31
      Top             =   5400
      Width           =   2625
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Termasuk GST (RM) :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   0
      TabIndex        =   28
      Top             =   6315
      Width           =   2985
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah GST (RM) :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   0
      TabIndex        =   26
      Top             =   6000
      Width           =   2985
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Rated (SR)"
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
      Left            =   480
      TabIndex        =   24
      Top             =   4725
      Width           =   3465
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Maklumat bayaran dan cukai GST"
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
      TabIndex        =   23
      Top             =   3120
      Width           =   3465
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Tunai      Bank In      Cek"
      ForeColor       =   &H00000000&
      Height          =   1005
      Left            =   3480
      TabIndex        =   18
      Top             =   7440
      Width           =   1185
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Cara bayaran diterima / dibuat :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   17
      Top             =   7440
      Width           =   3465
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah (RM) :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   0
      TabIndex        =   14
      Top             =   5040
      Width           =   2985
   End
   Begin VB.Label Label43 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier/Agen * :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   0
      TabIndex        =   12
      Top             =   1920
      Width           =   1545
   End
   Begin VB.Label Label110 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tarikh :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   -480
      TabIndex        =   11
      Top             =   2640
      Width           =   2025
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sila masukkan No. Invoice yang diterima dari supplier/agen (jika ada)."
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3120
      TabIndex        =   8
      Top             =   1560
      Width           =   6465
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Rujukan :"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   1575
      Width           =   1425
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "** Voucher : Bayaran yang dibuat kepada pihak supplier / agen."
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   1080
      Width           =   6105
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "** Invoice : Bayaran yang diterima dari supplier / agen."
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   840
      Width           =   5145
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Menu ini digunakan bagi merekod antara urusan kedai dan supplier/agen sahaja."
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7785
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis *            Invoice            Voucher"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   585
      Width           =   3465
   End
End
Attribute VB_Name = "frm118"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB1_Click()
'on error resume next
If frm118.CB1 = 1 Then
    frm118.CB2 = 0
    
    frm118.CB8 = 0
    frm118.CB9 = 0
    
    frm118.CB8.Enabled = False
    frm118.CB9.Enabled = False
    
    frm118.TB1 = vbNullString
    frm118.TB1.BackColor = &H8000000A
    frm118.TB1.Locked = True
End If
End Sub
Private Sub CB2_Click()
'on error resume next
If frm118.CB2 = 1 Then
    frm118.CB1 = 0
    
    frm118.CB8 = 1
    frm118.CB9 = 0
    
    frm118.CB8.Enabled = True
    frm118.CB9.Enabled = True
    
    frm118.TB1.BackColor = &HFFFFFF
    frm118.TB1.Locked = False
End If
End Sub
Private Sub CB3_Click()
'on error resume next
If frm118.CB3 = 1 Then
    frm118.CB4 = 0
    frm118.CB5 = 0
End If
End Sub
Private Sub CB4_Click()
'on error resume next
If frm118.CB4 = 1 Then
    frm118.CB3 = 0
    frm118.CB5 = 0
End If
End Sub
Private Sub CB5_Click()
'on error resume next
If frm118.CB5 = 1 Then
    frm118.CB4 = 0
    frm118.CB3 = 0
End If
End Sub
Private Sub CB6_Click()
'on error resume next
If frm118.CB6 = 1 Then
    frm118.CB7 = 0
End If

Call frm118_calc_gst_1
End Sub
Private Sub CB7_Click()
'on error resume next
If frm118.CB7 = 1 Then
    frm118.CB6 = 0
End If

Call frm118_calc_gst_1
End Sub

Private Sub CB8_Click()
'On Error Resume Next
If frm118.CB8 = 1 Then
    frm118.CB9 = 0
    frm118.TB7 = vbNullString
    frm118.TB8 = vbNullString
    frm118.TB7.BackColor = &H8000000A
    frm118.TB7.Locked = True
End If
End Sub
Private Sub CB9_Click()
'On Error Resume Next
If frm118.CB9 = 1 Then
    frm118.CB8 = 0
    frm118.TB7 = "0.00"
    frm118.TB8 = "0.00"
    frm118.TB7.BackColor = &HFFFFFF
    frm118.TB7.Locked = False
End If
End Sub
Private Sub CMD1_Click()
'On Error Resume Next
Dim Err(15)
Dim LM_HARGA_SEMASA As Double
Dim LM_HARGA As Double

LM_HARGA_SEMASA = 0
LM_HARGA = 0

If frm118.CB1 = 0 And frm118.CB2 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan jenis urusan."
End If
If frm118.TB1 <> vbNullString Then

    If InStr(1, frm118.TB1, "*") <> 0 Or InStr(1, frm118.TB1, "/") <> 0 Or InStr(1, frm118.TB1, "\") <> 0 Or InStr(1, frm118.TB1, "'") <> 0 Then

        x = x + 1
        Err(x) = "No. rujukan mengandungi simbol yang tidak sah."
        
    End If
    
End If
If frm118.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih supplier/agen."
End If
If frm118.L2_Text = vbNullString Or (frm118.L2_Text <> vbNullString And Not IsNumeric(frm118.L2_Text)) Then
    x = x + 1
    Err(x) = "Tiada maklumat tentang kadar cukai GST. Sila keluar dari menu ini dan cuba lagi."
End If
If frm118.L3_Text = vbNullString Or (frm118.L3_Text <> vbNullString And Not IsNumeric(frm118.L3_Text)) Then
    x = x + 1
    Err(x) = "Tiada maklumat tentang jumlah termasuk GST. Sila keluar dari menu ini dan cuba lagi."
End If
If frm118.TB2 = vbNullString Or (frm118.TB2 <> vbNullString And Not IsNumeric(frm118.TB2)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Jumlah bagi bayaran dengan cukai GST SR]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If frm118.TB3 = vbNullString Or (frm118.TB3 <> vbNullString And Not IsNumeric(frm118.TB3)) Then
    x = x + 1
    Err(x) = "Tiada maklumat tentang jumlah cukai GST. Sila keluar dari menu ini dan cuba lagi."
End If
If frm118.TB4 = vbNullString Or (frm118.TB4 <> vbNullString And Not IsNumeric(frm118.TB4)) Then
    x = x + 1
    Err(x) = "Tiada maklumat tentang jumlah termasuk GST. Sila keluar dari menu ini dan cuba lagi."
End If
If frm118.TB5 = vbNullString Or (frm118.TB5 <> vbNullString And Not IsNumeric(frm118.TB5)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Jumlah bagi bayaran dengan cukai GST ZR]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If frm118.CB6 = 0 And frm118.CB7 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan jenis GST SR."
End If
If frm118.CBB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih nama pekerja."
End If
If frm118.CB9 = 1 Then
    If frm118.TB7 = vbNullString Or (frm118.TB7 <> vbNullString And Not IsNumeric(frm118.TB7)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Harga Semasa]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If (frm118.TB7 <> vbNullString And IsNumeric(frm118.TB7)) Then
        
        LM_HARGA_SEMASA = frm118.TB7
        
        If LM_HARGA_SEMASA = 0 Then
            x = x + 1
            Err(x) = "Nilai 0 tidak dibenarkan di dalam ruangan [Harga Semasa]."
        End If
    End If
End If
If (frm118.TB6 <> vbNullString And IsNumeric(frm118.TB6)) Then
    
    LM_HARGA = frm118.TB6
    
    If LM_HARGA = 0 Then
        x = x + 1
        Err(x) = "Nilai bayaran adalah 0."
    End If
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
    
'### Periksa NO INVOICE sebelum simpan data ke dalam database ### - Start
        LM_NO_RUJUKAN = 1
        
'---------------------------------------No. Invoice
        LM_NOW = Now
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        If frm118.CB1 = 1 Then rs.Open "select * from 4_senarai_invoice_rasmi", cn2, adOpenKeyset, adLockOptimistic
        If frm118.CB2 = 1 Then rs.Open "select * from 8_gn_grn_vouher", cn2, adOpenKeyset, adLockOptimistic
        
        rs.AddNew
        rs!tarikh = frm118.DTPicker1
        rs!terminal = G_TERMINAL
        rs!write_timestamp = LM_NOW
        rs!Status = 1
        rs!nama_staff = MDI_frm1.L3_Text
        rs.Update
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        If frm118.CB1 = 1 Then rs.Open "select * from 4_senarai_invoice_rasmi where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & frm118.DTPicker1 & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic
        If frm118.CB2 = 1 Then rs.Open "select * from 8_gn_grn_vouher where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & frm118.DTPicker1 & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!ID) Then
                
                If frm118.CB1 = 1 Then rs!no_invoice = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(rs!ID, "000000")
                If frm118.CB2 = 1 Then rs!no_voucher = "PV" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(rs!ID, "000000")
                LM_NO_RUJUKAN = rs!ID 'No. Rujukan Belian
                rs.Update
                
            End If
            
        Else
        
            MsgBox "Berlaku ralat semasa data cuba disimpan. Sila keluar dari menu ini dan cuba lagi.", vbCritical, "Error"
            
            rs.Close
            Set rs = Nothing
            
            Exit Sub
            
        End If
        
        rs.Close
        Set rs = Nothing
'---------------------------------------No. Invoice
        
        GoTo a:
        
        If frm118.CB1 = 1 Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from default_setting where Default1='" & "Default" & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!ResitNo) Then
                    If IsNumeric(rs!ResitNo) Then LM_NO_RUJUKAN = rs!ResitNo 'No. invoice rasmi
                End If
            End If
            
            rs.Close
            Set rs = Nothing
        
Re_gen_no_resit:
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 22_jualan where no_resit='" & "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_NO_RUJUKAN, "000000") & "' AND bil_rasmi = 1", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
            
                LM_NO_RUJUKAN = LM_NO_RUJUKAN + 1
                
                rs.Close
                Set rs = Nothing
                
                GoTo Re_gen_no_resit:
            End If
            
            rs.Close
            Set rs = Nothing
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from default_setting where Default1='" & "Default" & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
            
                rs!ResitNo = LM_NO_RUJUKAN + 1 'No. invoice rasmi
                rs.Update
                
            End If
            
            rs.Close
            Set rs = Nothing
            
        End If
        

        If frm118.CB2 = 1 Then
' ### Periksa No. voucher ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from default_setting where Default1='" & "Default" & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
            
                If Not IsNull(rs!no_voc_grn) Then
                    If IsNumeric(rs!no_voc_grn) Then LM_NO_RUJUKAN = rs!no_voc_grn 'No. voucher
                End If
                
            End If
            
            rs.Close
            Set rs = Nothing
        
Re_gen_no_resit2:
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 77_gdn_grn where no_rujukan='" & "VOU" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_NO_RUJUKAN, "000000") & "' AND jenis_urusan = 1", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
            
                LM_NO_RUJUKAN = LM_NO_RUJUKAN + 1
                
                rs.Close
                Set rs = Nothing
                
                GoTo Re_gen_no_resit2:
                
            End If
            
            rs.Close
            Set rs = Nothing
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from default_setting where Default1='" & "Default" & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                rs!no_voc_grn = LM_NO_RUJUKAN + 1
                rs.Update
            
            End If
            
            rs.Close
            Set rs = Nothing
' ### Periksa No. Voucher ### - End

        End If
        
a:

        If frm118.CBB2 <> vbNullString Then
        
            frm118_LM_EMP_NAMA = Split(frm118.CBB2, "  |  ")(0)
            frm118_LM_EMP_NO = Split(frm118.CBB2, "  |  ")(1)
            
        End If
            
'### Masukkan maklumat Good Delivery Note (GRN) ### - Start
        LM_NOW = Now
        LM_TARIKH = DateTime.Date$
        LM_MASA = DateTime.Time$
        
        LM_GRN_RE_GEN = 0
        
Re_gen_no_resit3:

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If frm118.CB1 = 1 Then rs.Open "select * from 77_gdn_grn where no_rujukan='" & "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_NO_RUJUKAN, "000000") & "' AND jenis_urusan = 2", cn, adOpenKeyset, adLockOptimistic
        If frm118.CB2 = 1 Then rs.Open "select * from 77_gdn_grn where no_rujukan='" & "PV" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_NO_RUJUKAN, "000000") & "' AND jenis_urusan = 3", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
        
            rs.AddNew
            rs!tarikh = frm118.DTPicker1
            rs!masa = LM_MASA
            rs!write_timestamp = LM_NOW
            
            If frm118.CB1 = 1 Then
                rs!no_rujukan = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_NO_RUJUKAN, "000000")
                G_No_RESIT_JUALAN = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_NO_RUJUKAN, "000000")
                
                If frm118.TB6 <> vbNullString Then
                    rs!harga_dengan_gst_grn = Format(frm118.TB6, "0.00")
                Else
                    rs!harga_dengan_gst_grn = Null
                End If
                rs!jenis_urusan = 2
                rs!jenis = "INV"
            End If
            If frm118.CB2 = 1 Then
                rs!no_rujukan = "PV" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_NO_RUJUKAN, "000000")
                G_No_RESIT_JUALAN = "PV" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_NO_RUJUKAN, "000000")
                
                If frm118.TB6 <> vbNullString Then
                    rs!harga_dengan_gst = Format(frm118.TB6, "0.00")
                Else
                    rs!harga_dengan_gst = Null
                End If
                rs!jenis_urusan = 3
                rs!jenis = "VOU"
            End If

            If frm118.L3_Text <> vbNullString Then
                rs!harga_tanpa_gst = Format(frm118.L3_Text, "0.00")
            Else
                rs!harga_tanpa_gst = "0.00"
            End If
            If frm118.TB3 <> vbNullString Then
                rs!jumlah_gst = Format(frm118.TB3, "0.00")
            Else
                rs!jumlah_gst = "0.00"
            End If
            If frm118.L2_Text <> vbNullString Then
                rs!kadar_gst = Format(frm118.L2_Text, "0.00")
            Else
                rs!kadar_gst = "0.00"
            End If
            If frm118.TB5 <> vbNullString Then
                rs!gst_zr_harga = Format(frm118.TB5, "0.00")
            Else
                rs!gst_zr_harga = "0.00"
            End If
            If frm118.L3_Text <> vbNullString Then
                rs!gst_sr_harga = Format(frm118.L3_Text, "0.00")
            Else
                rs!gst_sr_harga = "0.00"
            End If
            rs!gst_zr_cukai = "0.00"
            If frm118.TB3 <> vbNullString Then
                rs!gst_sr_cukai = Format(frm118.TB3, "0.00")
            Else
                rs!gst_sr_cukai = "0.00"
            End If
            If frm118.TB1 <> vbNullString Then
                rs!no_rujukan_supplier = UCase(frm118.TB1)
            Else
                rs!no_rujukan_supplier = Null
            End If
            If frm118.TB2 <> vbNullString Then
                rs!jumlah = Format(frm118.TB2, "0.00")
            Else
                rs!jumlah = "0.00"
            End If
            If frm118.CB6 = 1 Then
                rs!jenis_gst = 0
            ElseIf frm118.CB7 = 1 Then
                rs!jenis_gst = 1
            End If
            If frm118.CB3 = 1 Then
                rs!cara_bayaran = 0
            ElseIf frm118.CB4 = 1 Then
                rs!cara_bayaran = 1
            ElseIf frm118.CB5 = 1 Then
                rs!cara_bayaran = 2
            End If
            If frm118.CB9 = 1 Then
            
                If frm118.TB6 <> vbNullString Then
                    rs!nilaian_harga_emas = Format(frm118.TB6, "0.00")
                Else
                    rs!nilaian_harga_emas = "0.00"
                End If
                If frm118.TB7 <> vbNullString Then
                    rs!harga_999 = Format(frm118.TB7, "0.00")
                Else
                    rs!harga_999 = "0.00"
                End If
                If frm118.TB8 <> vbNullString Then
                    rs!berat_tukaran = Format(frm118.TB8, "0.00")
                Else
                    rs!berat_tukaran = "0.00"
                End If
                
            Else
                
                rs!nilaian_harga_emas = Null
                rs!harga_999 = Null
                rs!berat_tukaran = Null
                
            End If
            If frm118.CB8 = 1 Then
            
                rs!umum_berat = 0 '0 : Umum , 1 :Berat
            
            ElseIf frm118.CB9 = 1 Then
            
                rs!umum_berat = 1 '0 : Umum , 1 :Berat
                
            End If
            
            rs!Status = 1
            rs!terminal = G_TERMINAL
            If frm118.CBB1 <> vbNullString Then
                rs!supplier_agen = frm118.CBB1
            Else
                rs!supplier_agen = Null
            End If
            rs!user = frm118_LM_EMP_NAMA 'Nama Pekerja
            rs!cawangan = G_KEDAI
            rs.Update
            DATA_SAVE = 1
            
        Else
        
            LM_NO_RUJUKAN = LM_NO_RUJUKAN + 1
            LM_GRN_RE_GEN = 1
            
            rs.Close
            Set rs = Nothing
            
            GoTo Re_gen_no_resit3:
            
        End If
        
        rs.Close
        Set rs = Nothing
'### Masukkan data voucher / invoice bagi belian agen ini ### - End

        'If LM_GRN_RE_GEN = 1 Then

        '    If frm118.CB1 = 1 Then

        '        Set rs = New ADODB.Recordset
        '        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        '        rs.Open "select * from default_setting where Default1='" & "Default" & "'", cn, adOpenKeyset, adLockOptimistic
                
        '        If Not rs.EOF Then
                
        '            rs!ResitNo = LM_NO_RUJUKAN + 1 'No. invoice rasmi
        '            rs.Update
                    
        '        End If
                
        '        rs.Close
        '        Set rs = Nothing
            
        '    End If
            
        '    If frm118.CB2 = 1 Then

        '        Set rs = New ADODB.Recordset
        '        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        '        rs.Open "select * from default_setting where Default1='" & "Default" & "'", cn, adOpenKeyset, adLockOptimistic
                
        '        If Not rs.EOF Then
                    
        '            rs!no_voc_grn = LM_NO_RUJUKAN + 1
        '            rs.Update
                
        '        End If
                
        '        rs.Close
        '        Set rs = Nothing
                
        '    End If
        
        'End If

'### Masukkan Data Jualan Ke Dalam Table Jualan ### - Start

        If DATA_SAVE = 1 Then

'### Simpan data di dalam table 22_jualan ### - Start
            If frm118.CB1 = 1 Then
            
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 22_jualan where no_resit='" & G_No_RESIT_JUALAN & "' AND bil_rasmi = 1 AND status = 1", cn, adOpenKeyset, adLockOptimistic
                
                If rs.EOF Then
                
                    rs.AddNew
                    rs!no_resit = G_No_RESIT_JUALAN 'No. invoice rasmi
                    rs!bil_rasmi = 1 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
                    rs!tarikh = frm118.DTPicker1 'Tarikh Jualan
                    rs!status_r = 0
                    rs!tunai = Format(0, "0.00")
                    rs!bank_in = Format(0, "0.00") 'Cara Bayaran : Bank In
                    rs!kad_kredit = Format(0, "0.00") 'Cara Bayaran : Kad Kredit
                    rs!duit_simpanan_kedai = Format(0, "0.00") 'Cara Bayaran : Simpanan Duit Di Kedai
                    rs!cek = Format(0, "0.00") 'Cara Bayaran : Cek
            
                    If frm118.CB3 = 1 Then
                    
                        If frm118.TB6 <> vbNullString Then
                            rs!tunai = Format(frm118.TB6, "0.00") 'Cara Bayaran : Tunai
                        Else
                            rs!tunai = Null 'Cara Bayaran : Tunai
                        End If
                        
                    ElseIf frm118.CB4 = 1 Then
                    
                        If frm118.TB6 <> vbNullString Then
                            rs!bank_in = Format(frm118.TB6, "0.00") 'Cara Bayaran : Bank In
                        Else
                            rs!bank_in = Null 'Cara Bayaran : Bank In
                        End If
                        
                    ElseIf frm118.CB5 = 1 Then
                
                        If frm118.TB6 <> vbNullString Then
                            rs!cek = Format(frm118.TB6, "0.00") 'Cara Bayaran : Cek
                        Else
                            rs!cek = Null 'Cara Bayaran : Cek
                        End If
                        
                    Else
                        
                        'rs!Tunai = Null 'Cara Bayaran : Tunai
                        'rs!bank_in = Null 'Cara Bayaran : Bank In
                        'rs!cek = Null 'Cara Bayaran : Cek
                        
                    End If
                    
                    If frm118.TB6 <> vbNullString Then
                        rs!jumlah_bayaran = Format(frm118.TB6, "0.00") 'Cara Bayaran : Jumlah Bayaran
                    Else
                        rs!jumlah_bayaran = Null 'Cara Bayaran : Jumlah Bayaran
                    End If
                    If frm118.TB6 <> vbNullString Then
                        rs!harga_barang = Format(frm118.TB6, "0.00") 'Jumlah Harga Barang Tanpa GST (RM)
                    Else
                        rs!harga_barang = Null 'Jumlah Harga Barang Tanpa GST (RM)
                    End If
                    If frm118.TB3 <> vbNullString Then
                        rs!jumlah_cukai_gst = Format(frm118.TB3, "0.00") 'Jumlah Cukai GST (ZR + SR)
                    Else
                        rs!jumlah_cukai_gst = Null 'Jumlah Cukai GST (ZR + SR)
                    End If
                    If frm118.TB6 <> vbNullString Then
                        rs!harga_barang_dengan_gst = Format(frm118.TB6, "0.00") 'Jumlah Harga Barang Dengan GST (RM)
                    Else
                        rs!harga_barang_dengan_gst = Null 'Jumlah Harga Barang Dengan GST (RM)
                    End If
                    If frm118.TB6 <> vbNullString Then
                        rs!harga_jualan = Format(frm118.TB6, "0.00") 'Jumlah Harga Jualan (RM)
                    Else
                        rs!harga_jualan = Null 'Jumlah Harga Jualan (RM)
                    End If
                    rs!flag_bayaran = 0 '0 : Pembeli Bayar , 1 : Kedai Bayar
                    If frm118.TB6 <> vbNullString Then
                        rs!jumlah_perlu_bayar = Format(frm118.TB6, "0.00") 'Jumlah Bayaran Yang Perlu Dibuat (RM)
                    Else
                        rs!jumlah_perlu_bayar = Null 'Jumlah Bayaran Yang Perlu Dibuat (RM)
                    End If
                    If frm118.TB5 <> vbNullString Then
                        rs!gst_zr_harga = Format(frm118.TB5, "0.00") 'Harga Keseluruhan Bagi Barang ZR
                    Else
                        rs!gst_zr_harga = Null 'Harga Keseluruhan Bagi Barang ZR
                    End If
                    rs!gst_zr_cukai = Format(0, "0.00") 'Jumlah Cukai Bagi ZR
                    If frm118.L3_Text <> vbNullString Then
                        rs!gst_sr_harga = Format(frm118.L3_Text, "0.00") 'Harga Keseluruhan Bagi Barang SR
                    Else
                        rs!gst_sr_harga = Null 'Harga Keseluruhan Bagi Barang SR
                    End If
                    If frm118.TB3 <> vbNullString Then
                        rs!gst_sr_cukai = Format(frm118.TB3, "0.00") 'Jumlah Cukai Bagi SR
                    Else
                        rs!gst_sr_cukai = Null 'Jumlah Cukai Bagi SR
                    End If
                    
                    rs!no_pekerja = frm118_LM_EMP_NO 'No. Pekerja
                    rs!jualan_online = 0
                    rs!Status = 1
                    rs!terminal = G_TERMINAL
                    rs!write_timestamp = LM_NOW
                    rs!cawangan = G_KEDAI
                    rs!Menu = 4

                    DATA_SAVE = 1
                    rs.Update
                    
                End If
                
                rs.Close
                Set rs = Nothing
                
            ElseIf frm118.CB2 = 1 Then
            
                Call frm118_save_data_expenses
        
            End If
'### Simpan data di dalam table 22_jualan ### - End

    '#### Update Log Aktiviti Sistem #### - Start
            'User = MDI_frm1.L3_Text
            If frm118.CB1 = 1 Then LogAct_Memory = "[" & frm118_LM_EMP_NAMA & "] Pengeluaran INV kepada agen/supplier. No. Rujukan [" & G_No_RESIT_JUALAN & "]."
            If frm118.CB2 = 1 Then LogAct_Memory = "[" & frm118_LM_EMP_NAMA & "] Pengeluaran VOU kepada agen/supplier. No. Rujukan [" & G_No_RESIT_JUALAN & "]."
            LogDate_Memory = LM_NOW
            Call UpdateLog_Database
    '#### Update Log Aktiviti Sistem #### - End
            
            Call frm118_initial_setting
            
            G_PREVIEW = 1
            Call frm118_cetak_inv_vou
            MsgBox "Data telah berjaya disimpan.", vbInformation, "Info"
            
        End If
    
    End If
    
End If
End Sub

Private Sub CMD2_Click()
'On Error Resume Next
Dim Err(15)
Dim LM_HARGA_SEMASA As Double
Dim LM_HARGA As Double

LM_HARGA_SEMASA = 0
LM_HARGA = 0

If frm118.L1_Text = vbNullString Then
    x = x + 1
    Err(x) = "Tiada maklumat asas bagi INV atau VOU ini. Sila keluar dari menu ini dan cuba lagi."
End If
If frm118.CB1 = 0 And frm118.CB2 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan jenis urusan."
End If
If frm118.TB1 <> vbNullString Then

    If InStr(1, frm118.TB1, "*") <> 0 Or InStr(1, frm118.TB1, "/") <> 0 Or InStr(1, frm118.TB1, "\") <> 0 Or InStr(1, frm118.TB1, "'") <> 0 Then

        x = x + 1
        Err(x) = "No. rujukan mengandungi simbol yang tidak sah."
        
    End If
    
End If
If frm118.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih supplier/agen."
End If
If frm118.L2_Text = vbNullString Or (frm118.L2_Text <> vbNullString And Not IsNumeric(frm118.L2_Text)) Then
    x = x + 1
    Err(x) = "Tiada maklumat tentang kadar cukai GST. Sila keluar dari menu ini dan cuba lagi."
End If
If frm118.L3_Text = vbNullString Or (frm118.L3_Text <> vbNullString And Not IsNumeric(frm118.L3_Text)) Then
    x = x + 1
    Err(x) = "Tiada maklumat tentang jumlah termasuk GST. Sila keluar dari menu ini dan cuba lagi."
End If
If frm118.TB2 = vbNullString Or (frm118.TB2 <> vbNullString And Not IsNumeric(frm118.TB2)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Jumlah bagi bayaran dengan cukai GST SR]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If frm118.TB3 = vbNullString Or (frm118.TB3 <> vbNullString And Not IsNumeric(frm118.TB3)) Then
    x = x + 1
    Err(x) = "Tiada maklumat tentang jumlah cukai GST. Sila keluar dari menu ini dan cuba lagi."
End If
If frm118.TB4 = vbNullString Or (frm118.TB4 <> vbNullString And Not IsNumeric(frm118.TB4)) Then
    x = x + 1
    Err(x) = "Tiada maklumat tentang jumlah termasuk GST. Sila keluar dari menu ini dan cuba lagi."
End If
If frm118.TB5 = vbNullString Or (frm118.TB5 <> vbNullString And Not IsNumeric(frm118.TB5)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Jumlah bagi bayaran dengan cukai GST ZR]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If frm118.CB6 = 0 And frm118.CB7 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan jenis GST SR."
End If
If frm118.CBB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih nama pekerja."
End If
If frm118.CB9 = 1 Then
    If frm118.TB7 = vbNullString Or (frm118.TB7 <> vbNullString And Not IsNumeric(frm118.TB7)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Harga Semasa]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If (frm118.TB7 <> vbNullString And IsNumeric(frm118.TB7)) Then
        
        LM_HARGA_SEMASA = frm118.TB7
        
        If LM_HARGA_SEMASA = 0 Then
            x = x + 1
            Err(x) = "Nilai 0 tidak dibenarkan di dalam ruangan [Harga Semasa]."
        End If
    End If
End If
If (frm118.TB6 <> vbNullString And IsNumeric(frm118.TB6)) Then
    
    LM_HARGA = frm118.TB6
    
    If LM_HARGA = 0 Then
        x = x + 1
        Err(x) = "Nilai bayaran adalah 0."
    End If
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

        If frm118.CBB2 <> vbNullString Then
        
            frm118_LM_EMP_NAMA = Split(frm118.CBB2, "  |  ")(0)
            frm118_LM_EMP_NO = Split(frm118.CBB2, "  |  ")(1)
            
        End If
            
'### Masukkan maklumat Good Delivery Note (GRN) ### - Start
        LM_NOW = Now
        LM_TARIKH = DateTime.Date$
        LM_MASA = DateTime.Time$

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 77_gdn_grn where no_rujukan='" & G_No_RESIT_JUALAN & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            G_ID = rs!ID
            Call recovery_77_gdn_grn
            
            rs!tarikh = frm118.DTPicker1
            
            If frm118.CB1 = 1 Then

                If frm118.TB6 <> vbNullString Then
                    rs!harga_dengan_gst_grn = Format(frm118.TB6, "0.00")
                Else
                    rs!harga_dengan_gst_grn = Null
                End If
                rs!jenis_urusan = 2
                rs!jenis = "INV"
                
            End If
            If frm118.CB2 = 1 Then

                If frm118.TB6 <> vbNullString Then
                    rs!harga_dengan_gst = Format(frm118.TB6, "0.00")
                Else
                    rs!harga_dengan_gst = Null
                End If
                rs!jenis_urusan = 3
                rs!jenis = "VOU"
                
            End If

            If frm118.L3_Text <> vbNullString Then
                rs!harga_tanpa_gst = Format(frm118.L3_Text, "0.00")
            Else
                rs!harga_tanpa_gst = "0.00"
            End If
            If frm118.TB3 <> vbNullString Then
                rs!jumlah_gst = Format(frm118.TB3, "0.00")
            Else
                rs!jumlah_gst = "0.00"
            End If
            If frm118.L2_Text <> vbNullString Then
                rs!kadar_gst = Format(frm118.L2_Text, "0.00")
            Else
                rs!kadar_gst = "0.00"
            End If
            If frm118.TB5 <> vbNullString Then
                rs!gst_zr_harga = Format(frm118.TB5, "0.00")
            Else
                rs!gst_zr_harga = "0.00"
            End If
            If frm118.L3_Text <> vbNullString Then
                rs!gst_sr_harga = Format(frm118.L3_Text, "0.00")
            Else
                rs!gst_sr_harga = "0.00"
            End If
            rs!gst_zr_cukai = "0.00"
            If frm118.TB3 <> vbNullString Then
                rs!gst_sr_cukai = Format(frm118.TB3, "0.00")
            Else
                rs!gst_sr_cukai = "0.00"
            End If
            If frm118.TB1 <> vbNullString Then
                rs!no_rujukan_supplier = UCase(frm118.TB1)
            Else
                rs!no_rujukan_supplier = Null
            End If
            If frm118.TB2 <> vbNullString Then
                rs!jumlah = Format(frm118.TB2, "0.00")
            Else
                rs!jumlah = "0.00"
            End If
            If frm118.CB6 = 1 Then
                rs!jenis_gst = 0
            ElseIf frm118.CB7 = 1 Then
                rs!jenis_gst = 1
            End If
            If frm118.CB3 = 1 Then
                rs!cara_bayaran = 0
            ElseIf frm118.CB4 = 1 Then
                rs!cara_bayaran = 1
            ElseIf frm118.CB5 = 1 Then
                rs!cara_bayaran = 2
            End If
            If frm118.CB9 = 1 Then
            
                If frm118.TB6 <> vbNullString Then
                    rs!nilaian_harga_emas = Format(frm118.TB6, "0.00")
                Else
                    rs!nilaian_harga_emas = "0.00"
                End If
                If frm118.TB7 <> vbNullString Then
                    rs!harga_999 = Format(frm118.TB7, "0.00")
                Else
                    rs!harga_999 = "0.00"
                End If
                If frm118.TB8 <> vbNullString Then
                    rs!berat_tukaran = Format(frm118.TB8, "0.00")
                Else
                    rs!berat_tukaran = "0.00"
                End If
                
            Else
                
                rs!nilaian_harga_emas = Null
                rs!harga_999 = Null
                rs!berat_tukaran = Null
                
            End If
            If frm118.CB8 = 1 Then
            
                rs!umum_berat = 0 '0 : Umum , 1 :Berat
            
            ElseIf frm118.CB9 = 1 Then
            
                rs!umum_berat = 1 '0 : Umum , 1 :Berat
                
            End If
            
            rs!Status = 1
            rs!terminal = G_TERMINAL
            If frm118.CBB1 <> vbNullString Then
                rs!supplier_agen = frm118.CBB1
            Else
                rs!supplier_agen = Null
            End If
            rs!user = frm118_LM_EMP_NAMA 'Nama Pekerja
            If Not IsNull(rs!cawangan) Then LM_CAWANGAN = rs!cawangan
            rs.Update
            DATA_SAVE = 1
            
        End If
        
        rs.Close
        Set rs = Nothing
'### Masukkan data voucher / invoice bagi belian agen ini ### - End

'### Masukkan Data Jualan Ke Dalam Table Jualan ### - Start

        If DATA_SAVE = 1 Then

'### Simpan data di dalam table 22_jualan ### - Start
            If frm118.CB1 = 1 Then
            
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 22_jualan where no_resit='" & G_No_RESIT_JUALAN & "' AND bil_rasmi = 1 AND status = 1", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
    
                    rs!no_resit = G_No_RESIT_JUALAN 'No. invoice rasmi
                    rs!bil_rasmi = 1 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
                    rs!tarikh = frm118.DTPicker1 'Tarikh Jualan
                    
                    rs!tunai = Format(0, "0.00")
                    rs!bank_in = Format(0, "0.00") 'Cara Bayaran : Bank In
                    rs!kad_kredit = Format(0, "0.00") 'Cara Bayaran : Kad Kredit
                    rs!duit_simpanan_kedai = Format(0, "0.00") 'Cara Bayaran : Simpanan Duit Di Kedai
                    rs!cek = Format(0, "0.00") 'Cara Bayaran : Cek
            
                    If frm118.CB3 = 1 Then
                    
                        If frm118.TB6 <> vbNullString Then
                            rs!tunai = Format(frm118.TB6, "0.00") 'Cara Bayaran : Tunai
                        Else
                            rs!tunai = Null 'Cara Bayaran : Tunai
                        End If
                        'rs!bank_in = "0.00" 'Cara Bayaran : Bank In
                        'rs!cek = Null 'Cara Bayaran : Cek
                        
                    ElseIf frm118.CB4 = 1 Then
                    
                        If frm118.TB6 <> vbNullString Then
                            rs!bank_in = Format(frm118.TB6, "0.00") 'Cara Bayaran : Bank In
                        Else
                            rs!bank_in = Null 'Cara Bayaran : Bank In
                        End If
                        'rs!Tunai = Null 'Cara Bayaran : Tunai
                        'rs!cek = Null 'Cara Bayaran : Cek
                        
                    ElseIf frm118.CB5 = 1 Then
                
                        If frm118.TB6 <> vbNullString Then
                            rs!cek = Format(frm118.TB6, "0.00") 'Cara Bayaran : Cek
                        Else
                            rs!cek = Null 'Cara Bayaran : Cek
                        End If
                        'rs!Tunai = Null 'Cara Bayaran : Tunai
                        'rs!bank_in = Null 'Cara Bayaran : Bank In
                        
                    Else
                        
                        'rs!Tunai = Null 'Cara Bayaran : Tunai
                        'rs!bank_in = Null 'Cara Bayaran : Bank In
                        'rs!cek = Null 'Cara Bayaran : Cek
                        
                    End If
                    
                    If frm118.TB6 <> vbNullString Then
                        rs!jumlah_bayaran = Format(frm118.TB6, "0.00") 'Cara Bayaran : Jumlah Bayaran
                    Else
                        rs!jumlah_bayaran = Null 'Cara Bayaran : Jumlah Bayaran
                    End If
                    If frm118.TB6 <> vbNullString Then
                        rs!harga_barang = Format(frm118.TB6, "0.00") 'Jumlah Harga Barang Tanpa GST (RM)
                    Else
                        rs!harga_barang = Null 'Jumlah Harga Barang Tanpa GST (RM)
                    End If
                    If frm118.TB3 <> vbNullString Then
                        rs!jumlah_cukai_gst = Format(frm118.TB3, "0.00") 'Jumlah Cukai GST (ZR + SR)
                    Else
                        rs!jumlah_cukai_gst = Null 'Jumlah Cukai GST (ZR + SR)
                    End If
                    If frm118.TB6 <> vbNullString Then
                        rs!harga_barang_dengan_gst = Format(frm118.TB6, "0.00") 'Jumlah Harga Barang Dengan GST (RM)
                    Else
                        rs!harga_barang_dengan_gst = Null 'Jumlah Harga Barang Dengan GST (RM)
                    End If
                    If frm118.TB6 <> vbNullString Then
                        rs!harga_jualan = Format(frm118.TB6, "0.00") 'Jumlah Harga Jualan (RM)
                    Else
                        rs!harga_jualan = Null 'Jumlah Harga Jualan (RM)
                    End If
                    rs!flag_bayaran = 0 '0 : Pembeli Bayar , 1 : Kedai Bayar
                    If frm118.TB6 <> vbNullString Then
                        rs!jumlah_perlu_bayar = Format(frm118.TB6, "0.00") 'Jumlah Bayaran Yang Perlu Dibuat (RM)
                    Else
                        rs!jumlah_perlu_bayar = Null 'Jumlah Bayaran Yang Perlu Dibuat (RM)
                    End If
                    If frm118.TB5 <> vbNullString Then
                        rs!gst_zr_harga = Format(frm118.TB5, "0.00") 'Harga Keseluruhan Bagi Barang ZR
                    Else
                        rs!gst_zr_harga = Null 'Harga Keseluruhan Bagi Barang ZR
                    End If
                    rs!gst_zr_cukai = Format(0, "0.00") 'Jumlah Cukai Bagi ZR
                    If frm118.L3_Text <> vbNullString Then
                        rs!gst_sr_harga = Format(frm118.L3_Text, "0.00") 'Harga Keseluruhan Bagi Barang SR
                    Else
                        rs!gst_sr_harga = Null 'Harga Keseluruhan Bagi Barang SR
                    End If
                    If frm118.TB3 <> vbNullString Then
                        rs!gst_sr_cukai = Format(frm118.TB3, "0.00") 'Jumlah Cukai Bagi SR
                    Else
                        rs!gst_sr_cukai = Null 'Jumlah Cukai Bagi SR
                    End If
                    
                    rs!no_pekerja = frm118_LM_EMP_NO 'No. Pekerja
                    rs!jualan_online = 0
                    rs!Status = 1
                    rs!terminal = G_TERMINAL
                    rs!write_timestamp = LM_NOW
                    rs!Menu = 4
                    
                    DATA_SAVE = 1
                    rs.Update
                    
                End If
                
                rs.Close
                Set rs = Nothing
                
            ElseIf frm118.CB2 = 1 Then
            
                Call frm118_save_data_expenses_edit
        
            End If
'### Simpan data di dalam table 22_jualan ### - End

    '#### Update Log Aktiviti Sistem #### - Start
            'User = MDI_frm1.L3_Text
            If frm118.CB1 = 1 Then LogAct_Memory = "[" & frm118_LM_EMP_NAMA & "] Edit data INV kepada agen/supplier. No. Rujukan [" & G_No_RESIT_JUALAN & "]."
            If frm118.CB2 = 1 Then LogAct_Memory = "[" & frm118_LM_EMP_NAMA & "] Edit data VOU kepada agen/supplier. No. Rujukan [" & G_No_RESIT_JUALAN & "]."
            LogDate_Memory = LM_NOW
            Call UpdateLog_Database
    '#### Update Log Aktiviti Sistem #### - End
            
            frm117.Show
            Unload frm118
            
            GM_NEXT_PREV = 2
            
            Call frm117_report_gdn_grn_header
            Call frm117_report_gdn_grn
            
            MsgBox "Data telah berjaya disimpan.", vbInformation, "Info"
            
        End If
    
    End If
    
End If
End Sub

Private Sub CMD3_Click()
'on error resume next
Note = "Adakah anda ingin batalkan edit data ini?" & vbCrLf & _
        "Sistem tidak akan menyimpan data jika terdapat data yang diubah." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then
    
    frm117.Show
    Unload frm118

End If
End Sub



Private Sub L2_Text_Change()
'on error resume next
Call frm118_calc_gst_1
End Sub
Private Sub L3_Text_Change()
'on error resume next
Call frm118_calc_gst_2
End Sub
Private Sub TB2_Change()
'on error resume next
Call frm118_calc_gst_1
End Sub
Private Sub TB3_Change()
'on error resume next
Call frm118_calc_gst_2
End Sub
Private Sub TB4_Change()
'on error resume next
Call frm118_calc_gst_3
End Sub
Private Sub TB5_Change()
'on error resume next
Call frm118_calc_gst_3
End Sub

Private Sub TB6_Change()
'On Error Resume Next
Call frm118_kiraan_berat_bayaran
End Sub

Private Sub TB7_Change()
'On Error Resume Next
Call frm118_kiraan_berat_bayaran
End Sub


