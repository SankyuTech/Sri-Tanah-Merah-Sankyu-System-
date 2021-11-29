VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frm2 
   BackColor       =   &H0080C0FF&
   Caption         =   "Menu Utama [Sistem Pengurusan Kedai Emas (Sankyu System)] Version 51.0.1"
   ClientHeight    =   12615
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   23760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm2.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Frm2.frx":0ECA
   ScaleHeight     =   12615
   ScaleWidth      =   23760
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton CMD1 
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   18240
      MaskColor       =   &H8000000B&
      Picture         =   "Frm2.frx":4D1C0C
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Tutup Sistem"
      Top             =   11280
      Width           =   855
   End
   Begin VB.CommandButton CMD_logout 
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   17280
      MaskColor       =   &H8000000B&
      Picture         =   "Frm2.frx":4D4570
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "User Logout"
      Top             =   11280
      Width           =   855
   End
   Begin VB.CommandButton CMD24 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   17880
      Picture         =   "Frm2.frx":4D67C8
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Menu tetapan sistem kedai."
      Top             =   2280
      Width           =   5700
   End
   Begin VB.CommandButton MAIN_BUT 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      Picture         =   "Frm2.frx":4D8A42
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ruangan untuk penerimaan stok baru samada secara mengikut berat ataupun mengikut item."
      Top             =   2280
      Width           =   5700
   End
   Begin VB.CommandButton MAIN_BUT2 
      BackColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   6120
      Picture         =   "Frm2.frx":4E3D44
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ruangan menu urusniaga kedai dengan pelanggan."
      Top             =   2280
      Width           =   5700
   End
   Begin VB.Timer Tmr1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   12240
      Width           =   23760
      _ExtentX        =   41910
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   13600
            MinWidth        =   1411
            Text            =   $"Frm2.frx":4E7BFA
            TextSave        =   $"Frm2.frx":4E7C97
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   13864
            Text            =   "Menu Utama                                                                                            "
            TextSave        =   "Menu Utama                                                                                            "
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   13864
            Text            =   "User :"
            TextSave        =   "User :"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6255
      Left            =   195
      TabIndex        =   6
      ToolTipText     =   "Sila klik di sini untuk update aktiviti sistem terbaru."
      Top             =   5520
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   11033
      _Version        =   393216
      Cols            =   0
      FixedCols       =   0
      BackColor       =   16777088
      ForeColor       =   0
      BackColorFixed  =   8454016
      BackColorBkg    =   12640511
      GridColor       =   0
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   3855
      Left            =   11280
      TabIndex        =   21
      ToolTipText     =   "Sila klik di sini untuk update analisis harga terbaru."
      Top             =   5520
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   0
      FixedCols       =   0
      BackColor       =   16777088
      ForeColor       =   0
      BackColorFixed  =   8454016
      BackColorBkg    =   12640511
      GridColor       =   0
      WordWrap        =   -1  'True
      TextStyle       =   4
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton CMD39 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   12000
      Picture         =   "Frm2.frx":4E7D34
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Report inventori , belian , jualan dan kewangan kedai."
      Top             =   2280
      Width           =   5700
   End
   Begin VB.Label Label67 
      BackColor       =   &H00000000&
      Caption         =   "Label7"
      ForeColor       =   &H00FFFFFF&
      Height          =   45
      Left            =   0
      TabIndex        =   26
      Top             =   2040
      Width           =   24000
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
      TabIndex        =   25
      Top             =   480
      Width           =   14295
   End
   Begin VB.Label L9_Text 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Update Terkini : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   24
      Top             =   4800
      Width           =   4215
   End
   Begin VB.Label L8_Text 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Update Terkini : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14520
      TabIndex        =   23
      Top             =   4920
      Width           =   4215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "****Sila klik senarai di atas untuk update analisis harga terbaru.****"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      TabIndex        =   22
      Top             =   9360
      Width           =   7455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Analisis Harga Emas Kedai"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   11280
      TabIndex        =   19
      Top             =   5160
      Width           =   7455
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   3
      Height          =   735
      Left            =   21360
      Shape           =   4  'Rounded Rectangle
      Top             =   11040
      Width           =   2295
   End
   Begin VB.Label Label86 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Powered By"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   21480
      TabIndex        =   17
      Top             =   11160
      Width           =   1335
   End
   Begin VB.Label Label87 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Sankyu System"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   21480
      TabIndex        =   16
      Top             =   11400
      Width           =   2055
   End
   Begin VB.Label Label44 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "sankyusystem@gmail.com / 010 - 900 4788"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   17760
      TabIndex        =   15
      Top             =   11880
      Width           =   5865
   End
   Begin VB.Label L2_Text 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "High Tower Text"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   405
      Left            =   17880
      TabIndex        =   12
      Top             =   4320
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Label L4_Text 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
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
      Left            =   21700
      TabIndex        =   11
      Top             =   1635
      Width           =   2100
   End
   Begin VB.Label L3_Text 
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
      Left            =   21700
      TabIndex        =   10
      Top             =   1320
      Width           =   2100
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Log Aktiviti Sistem"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   195
      TabIndex        =   9
      Top             =   5160
      Width           =   10815
   End
   Begin VB.Label Label_User 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   1560
      TabIndex        =   8
      Top             =   4830
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User   :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   360
      TabIndex        =   7
      Top             =   4830
      Width           =   1215
   End
   Begin VB.Label L_1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   11640
      Visible         =   0   'False
      Width           =   10695
   End
   Begin VB.Label Label88 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   21360
      TabIndex        =   18
      Top             =   11040
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "****Sila klik senarai di atas untuk update aktiviti sistem terbaru.****"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   20
      Top             =   11880
      Width           =   10455
   End
   Begin VB.Label Label85 
      BackColor       =   &H00FFFFFF&
      Height          =   2070
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   24000
   End
   Begin VB.Menu Frm2_MenuUtama 
      Caption         =   "Menu Utama"
      Visible         =   0   'False
      Begin VB.Menu Frm2_ReceivingPerGram 
         Caption         =   "Receving (Per Gram)"
      End
      Begin VB.Menu Frm2_ReceivingPerItem 
         Caption         =   "Receiving (Per Item)"
      End
      Begin VB.Menu Frm_Jualan_Fresh 
         Caption         =   "Jualan"
      End
      Begin VB.Menu Frm2_BelianTradeInCash 
         Caption         =   "Belian Trade In - Cash"
      End
      Begin VB.Menu Frm2_BelianTradeInItem 
         Caption         =   "Trade In Dengan Barang"
      End
      Begin VB.Menu Frm2_ServisPadaPelanggan 
         Caption         =   "Servis Pada Pelanggan"
      End
      Begin VB.Menu Frm2_ServisLain 
         Caption         =   "Pembelian Kedai Selain Dari Emas"
      End
      Begin VB.Menu Frm2_Insentif 
         Caption         =   "Insentif Pada Marketer"
      End
      Begin VB.Menu Frm2_ReportInsentif 
         Caption         =   "Report Insentif Pada Marketer"
      End
      Begin VB.Menu Frm2_Nego 
         Caption         =   "Runding Harga"
      End
      Begin VB.Menu Frm2_Ansuran 
         Caption         =   "Jualan Secara Ansuran"
      End
      Begin VB.Menu Frm2_DataAnsuran 
         Caption         =   "Lihat Data Ansuran"
      End
      Begin VB.Menu Frm2_Promosi 
         Caption         =   "Promosi"
      End
      Begin VB.Menu Frm2_TetapanSistem 
         Caption         =   "Tetapan Sistem"
      End
      Begin VB.Menu Frm2_DetailData 
         Caption         =   "Lihat Detail Data Produk"
      End
      Begin VB.Menu Frm2_PengeluaranCek 
         Caption         =   "Pengeluaran Cek"
      End
      Begin VB.Menu Frm2_Tempahan 
         Caption         =   "Tempahan"
      End
      Begin VB.Menu Frm2_Delay 
         Caption         =   "Delay"
      End
      Begin VB.Menu Frm2_Rahnu 
         Caption         =   "ArRahnu"
      End
   End
   Begin VB.Menu Frm2_Report 
      Caption         =   "Report"
      Visible         =   0   'False
      Begin VB.Menu Frm2_Report1 
         Caption         =   "Report Belian Dan Jualan"
      End
      Begin VB.Menu Frm2_Report2 
         Caption         =   "Report Belian Dari Supplier"
      End
      Begin VB.Menu Frm2_Report3 
         Caption         =   "Report Akaun (Harian)"
      End
      Begin VB.Menu Frm2_ReportJualanKategori 
         Caption         =   "Report Jualan Mengikut Kategori"
      End
      Begin VB.Menu Frm2_AkaunTerperinci 
         Caption         =   "Report Akaun Terperinci"
      End
      Begin VB.Menu Frm2_ReportCek 
         Caption         =   "Report Cek"
      End
      Begin VB.Menu Frm2_ReportTempahan 
         Caption         =   "Report Tempahan"
      End
   End
   Begin VB.Menu Frm2_User 
      Caption         =   "User"
      Visible         =   0   'False
      Begin VB.Menu Frm2_TukarPassword 
         Caption         =   "Tukar Password"
      End
   End
   Begin VB.Menu Frm2_Keluar 
      Caption         =   "Keluar"
      Begin VB.Menu Frm2_Logout 
         Caption         =   "Logout"
      End
      Begin VB.Menu Frm2_Exit 
         Caption         =   "Tutup Sistem"
      End
   End
   Begin VB.Menu Frm2_TentangSistem1 
      Caption         =   "Tentang Sistem"
      Begin VB.Menu Frm2_TentangSistem 
         Caption         =   "Sistem Version"
      End
      Begin VB.Menu Frm2_HubungiDeveloper 
         Caption         =   "Hubungi Developer"
      End
   End
End
Attribute VB_Name = "Frm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public UnloadSystem_OK As Integer
Private Sub CMD_logout_Click()
UnloadSystem_OK = 0

Note = "Adakah anda ingin keluar dari sistem ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    UnloadSystem_OK = 1
    Unload Frm2
    MsgBox "Anda telah berjaya keluar dari sistem.", vbInformation, "Logout Berjaya"
    Frm3.TxtUsername = vbNullString
    Frm3.Show
    Frm3.TxtUsername.SetFocus
    UnloadSystem_OK = 0
End If
End Sub
Private Sub CMD1_Click()
UnloadSystem_OK = 0
Note = "Adakah anda ingin tutup sistem ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    UnloadSystem_OK = 1
    Unload Frm2
    UnloadSystem_OK = 0
    MsgBox "Sistem telah berjaya ditutup.", vbInformation, "Tutup Sistem"
    End
End If
End Sub
Private Sub CMD3_1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'Frm2.CMD3_1.Visible = False
    'Frm2.CMD3.Visible = True
    Frm2.L_1.Caption = "Daftar User Baru"
End Sub
Private Sub CMD24_Click()
'On Error Resume Next
user = MDI_frm1.L3_Text

Set rs2 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs2.Open "select * from tblelogin where username='" & user & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs2.EOF Then
    If rs2!usertype = "Developer" Or rs2!usertype = "Admin" Then
        Frm30.Show
        Frm2.Hide
    Else
        MsgBox "Hanya ADMIN Sahaja Yang Dibenarkan Untuk Memasuki Menu Ini.", vbExclamation, "Info"
    End If
End If
rs2.Close
Set rs2 = Nothing
End Sub
Private Sub CMD3_Click()
If InStr(1, Frm2.StatusBar1.Panels(2), "Menu Utama") <> 0 Then
    Frm8.Show
    Frm2.Hide
    Frm8.TB1.SetFocus
    Frm2.StatusBar1.Panels(2) = "Daftar User Baru"
Else
    MsgBox "Anda perlu keluar dari Menu [" & Frm2.StatusBar1.Panels(2) & "] dahulu.", vbInformation, "Tutup Menu"
End If
End Sub
Private Sub CMD4_1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'Frm2.CMD4_1.Visible = False
    'Frm2.CMD4.Visible = True
    Frm2.L_1.Caption = "Padam Data User"
End Sub
Private Sub CMD3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Frm2.L2_Text.Visible = True
Frm2.L2_Text = "Daftar User Baru"
End Sub
Private Sub CMD39_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Frm2.L2_Text.Left = 12000
Frm2.L2_Text.Top = 4320
Frm2.L2_Text.Visible = True
Frm2.L2_Text = "Report"
End Sub
Private Sub CMD5_1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'Frm2.CMD5_1.Visible = False
    'Frm2.CMD5.Visible = True
    Frm2.L_1.Caption = "Admin Setting"
End Sub

Private Sub CMD4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Frm2.L2_Text.Visible = True
Frm2.L2_Text = "Padam Data User"
End Sub
Private Sub CMD24_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Frm2.L2_Text.Left = 17880
Frm2.L2_Text.Top = 4320
Frm2.L2_Text.Visible = True
Frm2.L2_Text = "Admin"
End Sub
Private Sub CMD29_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Frm2.L2_Text.Visible = True
Frm2.L2_Text = "Cetak Barcode"
End Sub

Private Sub CMD39_Click()
'On Error Resume Next
Frm34.Show
Frm2.Hide
End Sub
Private Sub CMD6_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Frm2.L2_Text.Visible = True
Frm2.L2_Text = "Tetapan Asas Sistem"
End Sub
Private Sub Form_Activate()
'On Error Resume Next
Frm2.Tmr1.Enabled = True
Frm2.Tmr1.Interval = 1
'Call Call_ReportPurityItem
'Call Call_BelianJualanToday
'Call AnalystSpot
'Call UpdateLog_Main
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Frm2.Label_User <> vbNullString Then
    If InStr(1, Split(Frm2.Label_User, "  [")(1), "Guest") = 0 Then
        'Frm2.CMD3_1.Visible = True
        'Frm2.CMD3.Visible = False
        'Frm2.CMD4_1.Visible = True
        'Frm2.CMD4.Visible = False
        'Frm2.CMD5_1.Visible = True
        'Frm2.CMD5.Visible = False
        Frm2.L_1.Visible = True
        Frm2.L_1.Caption = vbNullString
        'Frm2.CMD17.BackColor = &H8000000C
        'Frm2.CMD18.BackColor = &H8000000C
        'Frm2.CMD19.BackColor = &H8000000C
    End If
End If
Frm2.L2_Text.Visible = False
Frm2.L2_Text = vbNullString
End Sub
Private Sub Form_Unload(Cancel As Integer)
If UnloadSystem_OK = 0 Then Cancel = True
End Sub
Private Sub Frm2_Exit_Click()
UnloadSystem_OK = 0
Note = "Adakah anda ingin tutup sistem ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    UnloadSystem_OK = 1
    Unload Frm2
    UnloadSystem_OK = 0
    MsgBox "Sistem telah berjaya ditutup.", vbInformation, "Tutup Sistem"
    End
End If
End Sub

Private Sub Frm2_HubungiDeveloper_Click()
'On Error Resume Next
MsgBox "Developer : INSAN" & vbCrLf & _
       "No. Telefon : +6010 - 900 4788" & vbCrLf & _
       "Email : sankyusystem@gmail.com " & vbCrLf & _
       "Facebook : https://www.facebook.com/ExcelVisualBasicApplicationVba " & vbCrLf & _
       "Website : http://sankyutech-visualbasic.weebly.com " & vbCrLf & _
       "" & vbCrLf & _
       "Terima Kasih Menggunakan Sistem Ini.", vbInformation, "Info"
End Sub
Private Sub Frm2_Logout_Click()
'On Error Resume Next
UnloadSystem_OK = 0
Note = "Adakah anda ingin keluar dari sistem ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    UnloadSystem_OK = 1
    Unload Frm2
    MsgBox "Anda telah berjaya keluar dari sistem.", vbInformation, "Logout Berjaya"
    Frm3.TxtUsername = vbNullString
    Frm3.Show
    Frm3.TxtUsername.SetFocus
    UnloadSystem_OK = 0
    End If
End Sub
Private Sub Frm2_TentangSistem_Click()
'On Error Resume Next
MsgBox "Sankyu System" & vbCrLf & _
       "Sistem Pengurusan Kedai Emas (SPKE 51.0.1)" & vbCrLf & _
       "Version Sistem : SPKE 51.0.1" & vbCrLf & _
       "Version Database : spke5100" & vbCrLf & _
       "Version Database Image : spke5100_image" & vbCrLf & _
       "Version AE : ae300 / AE3.0.0", vbInformation, "Info"
End Sub
Private Sub CMD29_Click()
'On Error Resume Next
Set rs2 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs2.Open "select * from Data_Database order by ID DESC", cn, adOpenKeyset, adLockOptimistic

While rs2.EOF = False
    With Frm33.List1
        If rs2!no_siri_Produk <> vbNullString Then .AddItem rs2!no_siri_Produk
    End With
    rs2.MoveNext
Wend
rs2.Close
Set rs2 = Nothing

Frm2.Hide
Frm33.Show
End Sub

Private Sub MAIN_BUT_Click()
'On Error Resume Next
Frm16.CB9 = 1
Frm16.Show
Frm2.Hide
End Sub
Private Sub MAIN_BUT_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'On Error Resume Next
Frm2.L2_Text.Left = 240
Frm2.L2_Text.Top = 4320
Frm2.L2_Text.Visible = True
Frm2.L2_Text = "Stock In"
End Sub
Private Sub MAIN_BUT2_Click()
'On Error Resume Next
'Frm2.Pic1.Enabled = False
'Frm53.L1_Text = 1
'Frm53.L2_Text = 2
'Frm53.Show
Frm15.Show
Frm2.Hide
End Sub
Private Sub MAIN_BUT2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'On Error Resume Next
Frm2.L2_Text.Left = 6120
Frm2.L2_Text.Top = 4320
Frm2.L2_Text.Visible = True
Frm2.L2_Text = "Transaksi Kedai"
End Sub
Private Sub MSFlexGrid1_DblClick()
'On Error Resume Next
Call UpdateLog_Main
End Sub
Private Sub MSFlexGrid2_DblClick()
'On Error Resume Next
Call Frm2_AnalystSpot
End Sub
Private Sub Pic1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'On Error Resume Next
Frm2.L2_Text.Visible = False
Frm2.L2_Text = vbNullString
End Sub
Private Sub Tmr1_Timer()
'On Error Resume Next
Frm2.L3_Text = DateTime.Date
Frm2.L4_Text = DateTime.Time$
End Sub
