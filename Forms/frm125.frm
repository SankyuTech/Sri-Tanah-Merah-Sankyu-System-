VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm125 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Barang hilang , kecurian dan sebagainya."
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8160
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
   Icon            =   "frm125.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CMD2 
      BackColor       =   &H000080FF&
      Caption         =   "Batal"
      Height          =   405
      Left            =   4920
      MaskColor       =   &H00400000&
      MouseIcon       =   "frm125.frx":0ECA
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4440
      Width           =   2865
   End
   Begin VB.CommandButton CMD1 
      BackColor       =   &H000080FF&
      Caption         =   "Simpan Data"
      Height          =   405
      Left            =   1920
      MaskColor       =   &H00400000&
      MouseIcon       =   "frm125.frx":11D4
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   4440
      Width           =   2865
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
      Height          =   1740
      Left            =   1860
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frm125.frx":14DE
      Top             =   2640
      Width           =   6015
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   360
      Left            =   1860
      TabIndex        =   2
      Top             =   2280
      Width           =   6015
      _ExtentX        =   10610
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
      Format          =   415498240
      CurrentDate     =   41561
   End
   Begin VB.Label L7_Text 
      BackColor       =   &H8000000A&
      Caption         =   "L7_Text"
      Height          =   300
      Left            =   360
      TabIndex        =   19
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sebab * :"
      Height          =   300
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tarikh * :"
      Height          =   300
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label L6_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L6_Text"
      Height          =   300
      Left            =   1845
      TabIndex        =   16
      Top             =   1920
      Width           =   5055
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dulang :"
      Height          =   300
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label L5_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L5_Text"
      Height          =   300
      Left            =   1845
      TabIndex        =   14
      Top             =   1680
      Width           =   5055
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Harga modal :"
      Height          =   300
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label L4_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L4_Text"
      Height          =   300
      Left            =   1845
      TabIndex        =   12
      Top             =   1440
      Width           =   5055
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Berat :"
      Height          =   300
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label L3_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L3_Text"
      Height          =   300
      Left            =   1845
      TabIndex        =   10
      Top             =   1200
      Width           =   5055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Purity :"
      Height          =   300
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label L2_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L2_Text"
      Height          =   300
      Left            =   1845
      TabIndex        =   8
      Top             =   960
      Width           =   5055
   End
   Begin VB.Label L1_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "L1_Text"
      Height          =   300
      Left            =   1850
      TabIndex        =   7
      Top             =   720
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Berikut adalah maklumat barang yang hilang atau kecurian. Sila masukkan sebab barang ini dikeluarkan dari senarai stok kedai."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   8535
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Kategori Produk :"
      Height          =   300
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label62 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Siri Produk :"
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "frm125"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD1_Click()
'on error resume next
If frm125.L7_Text = vbNullString Then

    MsgBox "Berlaku ralat. Sila keluar dari menu ini dan cuba lagi.", vbExclamation, "Info"
    
    Exit Sub
    
End If
If frm125.TB1 = vbNullString Then

    MsgBox "Sila masukkan sebab.", vbExclamation, "Info"
    
    Exit Sub
    
End If

Note = "Adakah anda ingin menukar status barang ini?" & vbCrLf & _
        vbNullString & vbCrLf & _
        "Barang ini akan dikeluarkan dari stok kedai dan barang ini akan dimasukkan ke dalam senarai barang yang hilang , kecurian atau sebagainya." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Teruskan?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then
    
    GoTo skip_carian_user:
    
    If MDI_frm1.L3_Text <> vbNullString Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!NoPekerja) Then G_LOGIN_USER = rs!NoPekerja
    
        End If
        
        rs.Close
        Set rs = Nothing
        
    End If
    
skip_carian_user:
            
    LM_NOW = Now

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from Data_Database where ID='" & frm125.L7_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        
        If Not IsNull(rs!StatusItem) Then
            
            If rs!StatusItem = "10" Then
            
                G_ID = rs!ID
                Call recovery_data_database
                
                If Not IsNull(rs!no_siri_Produk) Then LM_NO_SIRI = rs!no_siri_Produk
                rs!StatusItem = 29
                rs!write_timestamp2 = LM_NOW
                rs!no_pekerja = G_LOGIN_USER
                rs!terminal = G_TERMINAL
                
                rs.Update
                
            Else
                
                MsgBox "Status barang ini telah berubah. Oleh itu anda tidak dibenarkan untuk tukar status ini. Sila periksa status terbaru barang ini.", vbExclamation, "Info"
                
                rs.Close
                Set rs = Nothing
                
                Exit Sub
            
            End If
            
        End If
        
    End If
    
    rs.Close
    Set rs = Nothing
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    
    strsql = "insert into 86_barang_hilang(id_data,Purity,kategori_Produk,no_siri_Produk,Beza_Berat,Dulang,tarikh,sebab,nama_pekerja,write_timestamp,harga_item,Status)" & _
                "select ID,Purity,kategori_Produk,no_siri_Produk,Beza_Berat,Dulang,'" & frm125.DTPicker1 & "','" & UCase(frm125.TB1) & "','" & MDI_frm1.L3_Text & "','" & LM_NOW & "',harga_item,1 from Data_Database WHERE ID='" & frm125.L7_Text & "'"
    
    Set rs = cn.Execute(strsql)
    Set rs = Nothing
    
'#### Update Log Aktiviti Sistem #### - Start
    user = MDI_frm1.L3_Text
    LogAct_Memory = "[" & user & "] Tukar status barang [" & LM_NO_SIRI & "] kepada hilang , dicuri dan sebagainya."
    LogDate_Memory = LM_NOW
    Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End
    
    Unload frm125
    
    GM_NEXT_PREV = 2 '0 : Next , 1 : Previous
    
    If GM_REPORT_MODE = 9 Then '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
        Call Frm85_Header_Report_Stok
        Call Frm85_report_stok_barcode
    Else
        Call Frm85_Header_Report_Stok
        Call Frm85_report_stok_page
    End If
    
    MsgBox "Status barang ini telah berjaya diubah.", vbInformation, "Info"
    
End If
End Sub
Private Sub CMD2_Click()
'on error resume next
Unload frm125

MsgBox "Urusan telah dibatalkan.", vbInformation, "Info"
End Sub
