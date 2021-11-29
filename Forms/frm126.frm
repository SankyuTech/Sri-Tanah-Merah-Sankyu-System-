VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm126 
   Caption         =   "Report Barang Hilang , Dicuri Dan Sebagainya."
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
   Icon            =   "frm126.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13035
   ScaleWidth      =   23880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD3 
      BackColor       =   &H000080FF&
      Caption         =   "Report"
      Height          =   405
      Left            =   3360
      MaskColor       =   &H00400000&
      MouseIcon       =   "frm126.frx":0ECA
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Report"
      Top             =   1440
      Width           =   2385
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
      Left            =   120
      TabIndex        =   2
      Top             =   160
      Width           =   200
   End
   Begin VB.CommandButton CMD22 
      BackColor       =   &H00FFFFFF&
      Height          =   650
      Left            =   18600
      MaskColor       =   &H00400000&
      MouseIcon       =   "frm126.frx":11D4
      MousePointer    =   99  'Custom
      Picture         =   "frm126.frx":14DE
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Paparan seterusnya"
      Top             =   10995
      Width           =   1100
   End
   Begin VB.CommandButton CMD21 
      BackColor       =   &H00FFFFFF&
      Height          =   650
      Left            =   17400
      MaskColor       =   &H00400000&
      MouseIcon       =   "frm126.frx":1E04
      MousePointer    =   99  'Custom
      Picture         =   "frm126.frx":210E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Paparan sebelumnya"
      Top             =   10995
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   360
      Left            =   2025
      TabIndex        =   4
      Top             =   495
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
      Format          =   415432704
      CurrentDate     =   41561
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   360
      Left            =   2025
      TabIndex        =   5
      Top             =   855
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
      Format          =   415432704
      CurrentDate     =   41561
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   8805
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
      Top             =   2115
      Width           =   19485
      _ExtentX        =   34369
      _ExtentY        =   15531
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
   Begin VB.Label Label62 
      BackStyle       =   0  'Transparent
      Caption         =   "Tarikh Mula "
      Height          =   300
      Left            =   315
      TabIndex        =   25
      Top             =   540
      Width           =   2535
   End
   Begin VB.Label Label63 
      BackStyle       =   0  'Transparent
      Caption         =   "Tarikh Akhir "
      Height          =   300
      Left            =   315
      TabIndex        =   24
      Top             =   900
      Width           =   2895
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Sila klik jika ingin melihat senarai rekod di dalam tempoh tarikh di bawah."
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   360
      TabIndex        =   23
      Top             =   120
      Width           =   8370
   End
   Begin VB.Label L5_Text 
      BackColor       =   &H00C0C0FF&
      Caption         =   "L5_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9120
      TabIndex        =   22
      Top             =   555
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label L6_Text 
      BackColor       =   &H00C0C0FF&
      Caption         =   "L6_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9120
      TabIndex        =   21
      Top             =   915
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label L7_Text 
      BackColor       =   &H00C0C0FF&
      Caption         =   "L7_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9120
      TabIndex        =   20
      Top             =   1275
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label L8_Text 
      BackColor       =   &H00C0C0FF&
      Caption         =   "L8_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   10800
      TabIndex        =   19
      Top             =   960
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label L9_Text 
      BackColor       =   &H00C0C0FF&
      Caption         =   "L9_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   10800
      TabIndex        =   18
      Top             =   600
      Visible         =   0   'False
      Width           =   1500
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
      Left            =   360
      TabIndex        =   17
      Top             =   1850
      Width           =   15495
   End
   Begin VB.Label L70_Text 
      BackColor       =   &H8000000C&
      Caption         =   "L70_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3720
      TabIndex        =   16
      Top             =   11115
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label L69_Text 
      BackColor       =   &H8000000C&
      Caption         =   "L69_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3720
      TabIndex        =   15
      Top             =   11475
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
      Left            =   16320
      TabIndex        =   14
      Top             =   10995
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
      Left            =   16920
      TabIndex        =   13
      Top             =   10995
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bil. Barang :"
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
      Top             =   10995
      Width           =   1455
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
      Left            =   1920
      TabIndex        =   11
      Top             =   10995
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Berat :"
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
      TabIndex        =   10
      Top             =   11235
      Width           =   1455
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
      Left            =   1920
      TabIndex        =   9
      Top             =   11235
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Modal :"
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
      TabIndex        =   8
      Top             =   11475
      Width           =   1455
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
      Left            =   1920
      TabIndex        =   7
      Top             =   11475
      Width           =   1695
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
      Left            =   15000
      TabIndex        =   26
      Top             =   10995
      Width           =   2295
   End
   Begin VB.Menu frm126_pm_menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu frm126_sm_excel 
         Caption         =   "Report Excel"
      End
      Begin VB.Menu frm126_sm_pulang_stok 
         Caption         =   "Pulangkan Ke Stok Kedai"
      End
   End
End
Attribute VB_Name = "frm126"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD21_Click()
'on error resume next
Dim frm126_LM_CURR_PAGE As Double
Dim frm126_LM_TOTAL_PAGE As Double

frm126_LM_CURR_PAGE = 0
frm126_LM_TOTAL_PAGE = 0

If frm126.L67_Text <> vbNullString And IsNumeric(frm126.L67_Text) Then
    If frm126.L68_Text <> vbNullString And IsNumeric(frm126.L68_Text) Then
        frm126_LM_CURR_PAGE = frm126.L67_Text
        frm126_LM_TOTAL_PAGE = frm126.L68_Text
        
        If frm126_LM_CURR_PAGE <> 1 And frm126_LM_CURR_PAGE <> 0 Then
        
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
                    
            Call frm126_barang_hilang_header
            Call frm126_barang_hilang
            
        End If

    End If
End If
End Sub
Private Sub CMD22_Click()
'on error resume next
Dim frm126_LM_CURR_PAGE As Double
Dim frm126_LM_TOTAL_PAGE As Double

frm126_LM_CURR_PAGE = 0
frm126_LM_TOTAL_PAGE = 0

If frm126.L67_Text <> vbNullString And IsNumeric(frm126.L67_Text) Then
    If frm126.L68_Text <> vbNullString And IsNumeric(frm126.L68_Text) Then
        frm126_LM_CURR_PAGE = frm126.L67_Text
        frm126_LM_TOTAL_PAGE = frm126.L68_Text
        
        If frm126_LM_CURR_PAGE < frm126_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call frm126_barang_hilang_header
            Call frm126_barang_hilang
                        
        End If
    End If
End If
End Sub

Private Sub CMD3_Click()
'On Error Resume Next
If frm126.CB1 = 1 Then 'Pilihan tarikh
    frm126.L5_Text = 1
Else
    frm126.L5_Text = 0
End If

frm126.L6_Text = frm126.DTPicker1 'Tarikh mula
frm126.L7_Text = frm126.DTPicker2 'Tarikh akhir

frm126.L69_Text = -1 'Titik Pencarian Data
frm126.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
frm126.L67_Text = 0 'Paparan Page ke-xxx
frm126.L68_Text = 0

GM_NEXT_PREV = 0

Call frm126_barang_hilang_header
Call frm126_barang_hilang
End Sub
Private Sub Form_Load()
'On Error Resume Next
frm126.CB1 = 0
frm126.DTPicker1 = DateTime.Date
frm126.DTPicker2 = DateTime.Date

frm124.L5_Text = 0
frm124.L6_Text = frm124.DTPicker1 'Tarikh mula
frm124.L7_Text = frm124.DTPicker2 'Tarikh akhir
End Sub

Private Sub frm126_sm_excel_Click()
'on error resume next
Dim TM As Date
Dim TA As Date

LM_FOUND = 0
frm126_LM_No_ID = vbNullString

If frm126.MSFlexGrid1 <> vbNullString Then

    If IsNumeric(frm126.MSFlexGrid1) Then
    
        frm126_LM_No_ID = frm126.MSFlexGrid1.TextMatrix(frm126.MSFlexGrid1, 2) 'No. ID
        
        If frm126_LM_No_ID <> vbNullString Then

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
                    .Columns("C").ColumnWidth = 15 'No. Siri Produk
                    .Columns("D").ColumnWidth = 40 'Kategori Produk
                    .Columns("E").ColumnWidth = 15 'Purity
                    .Columns("F").ColumnWidth = 15 'Berat (g)
                    .Columns("G").ColumnWidth = 15 'Modal (RM)
                    .Columns("H").ColumnWidth = 15 'Dulang
                    .Columns("I").ColumnWidth = 70 'Sebab
                
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
                
                    .Cells(1, 5).Font.Bold = True
                    .Cells(1, 5).Font.Size = 30
                    
                    For Row = 1 To 5
                        .Cells(Row, 5).HorizontalAlignment = xlCenter
                    Next Row
                    
                    .Cells(7, 1) = frm126.L14_Text
                    
                    .Cells(8, 1) = "No."
                    .Cells(8, 2) = "Tarikh"
                    .Cells(8, 3) = "No. Siri Produk"
                    .Cells(8, 4) = "Kategori Produk"
                    .Cells(8, 5) = "Purity"
                    .Cells(8, 6) = "Berat (g)"
                    .Cells(8, 7) = "Modal (RM)"
                    .Cells(8, 8) = "Dulang"
                    .Cells(8, 9) = "Sebab"

                    For i = 1 To 9
                        .Cells(8, i).HorizontalAlignment = xlCenter
                        .Cells(8, i).Interior.ColorIndex = 15
                        .Cells(8, i).WrapText = True
                        .Cells(8, i).Borders.LineStyle = xlContinuous
                    Next i
            
                    If frm126.L5_Text = "1" Then '0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh
                    
                        TM = frm126.L6_Text 'Tarikh mula
                        TA = frm126.L7_Text 'Tarikh akhir
                    
                    End If
                    
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    If frm126.L5_Text = 0 Then rs.Open "select * from 86_barang_hilang where status = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
                    If frm126.L5_Text = 1 Then rs.Open "select * from 86_barang_hilang where status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
                    
                    While rs.EOF = False
                    
                        x = x + 1
                        .Cells(8 + x, 1) = x 'No.
                        .Cells(8 + x, 1).HorizontalAlignment = xlCenter
                        
                        If Not IsNull(rs!tarikh) Then .Cells(8 + x, 2) = "'" & rs!tarikh 'Tarikh
                        .Cells(8 + x, 2).HorizontalAlignment = xlCenter
                                            
                        If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 3) = rs!no_siri_Produk 'No. Siri Produk
                        If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 4) = rs!kategori_Produk 'Kategori Produk
                        
                        If Not IsNull(rs!purity) Then .Cells(8 + x, 5) = rs!purity 'Purity
                        .Cells(8 + x, 5).HorizontalAlignment = xlCenter
                        
                        .Cells(8 + x, 6).HorizontalAlignment = xlRight
                        If Not IsNull(rs!beza_berat) Then .Cells(8 + x, 6) = Format(rs!beza_berat, "#,##0.00") 'Berat (g)
                        .Cells(8 + x, 6).NumberFormat = "#,##0.00"
                        
                        .Cells(8 + x, 7).HorizontalAlignment = xlRight
                        If Not IsNull(rs!harga_item) Then .Cells(8 + x, 7) = Format(rs!harga_item, "#,##0.00") 'Modal (RM)
                        .Cells(8 + x, 7).NumberFormat = "#,##0.00"
                        
                        .Cells(8 + x, 8).HorizontalAlignment = xlCenter
                        If Not IsNull(rs!dulang) Then .Cells(8 + x, 8) = rs!dulang 'Dulang
                        
                        If Not IsNull(rs!sebab) Then .Cells(8 + x, 9) = rs!sebab  'Sebab
                        
                        For Col = 1 To 9
                            .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
                        Next Col
                                            
                        rs.MoveNext
                    Wend
                    
                    rs.Close
                    Set rs = Nothing
                    
                    Y = 0
                    Y = x + 2
                    
                    .Cells(8 + Y, 1) = "Bilangan : " & frm126.L10_Text
                    .Cells(8 + Y, 1).Font.Bold = True
                    
                    Y = Y + 1
                    .Cells(8 + Y, 1) = "Jumlah Berat : " & frm126.L11_Text
                    .Cells(8 + Y, 1).Font.Bold = True
                        
                    Y = Y + 1
                    .Cells(8 + Y, 1) = "Jumlah Modal : " & frm126.L12_Text
                    .Cells(8 + Y, 1).Font.Bold = True
                    
                    Y = Y + 2
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

Private Sub frm126_sm_pulang_stok_Click()
'on error resume next
LM_FOUND = 0
frm126_LM_No_ID = vbNullString

If frm126.MSFlexGrid1 <> vbNullString Then

    If IsNumeric(frm126.MSFlexGrid1) Then
    
        frm126_LM_No_ID = frm126.MSFlexGrid1.TextMatrix(frm126.MSFlexGrid1, 2) 'No. ID
        
        If frm126_LM_No_ID <> vbNullString Then

            Note = "Adakah anda pulangkan barang ini ke dalam stok kedai?" & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Teruskan?"
                    
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbYes Then
            
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 86_barang_hilang where status = 1 AND ID='" & frm126_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    
                    If Not IsNull(rs!no_siri_Produk) Then LM_NO_SIRI = rs!no_siri_Produk
                    If Not IsNull(rs!id_data) Then
                        LM_ID = rs!id_data
                        
                        rs!Status = 0
                        
                        rs.Update
                        
                        LM_FOUND = 1
                    End If
                    
                End If
                
                rs.Close
                Set rs = Nothing
                
                If LM_FOUND = 1 Then
                    
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
                    rs.Open "select * from Data_Database where ID='" & LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                        
                        G_ID = rs!ID
                        Call recovery_data_database
                        
                        rs!StatusItem = 10
                        rs!write_timestamp2 = LM_NOW
                        rs!no_pekerja = G_LOGIN_USER
                        rs!terminal = G_TERMINAL
                        
                        rs.Update
                                
                    End If
                    
                    rs.Close
                    Set rs = Nothing
                    
'#### Update Log Aktiviti Sistem #### - Start
                    user = MDI_frm1.L3_Text
                    LogAct_Memory = "[" & user & "] Pulangkan status barang ke dalam stok kedai.[" & LM_NO_SIRI & "]."
                    LogDate_Memory = LM_NOW
                    Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End
                    
                    GM_NEXT_PREV = 2
                    
                    Call frm126_barang_hilang_header
                    Call frm126_barang_hilang
                    
                    MsgBox "Status item ini telah berjaya pulangkan ke dalam stok kedai.", vbInformation, "Info"
                    
                End If
                
            End If
            
        End If
        
    End If
    
End If
End Sub

Private Sub MSFlexGrid1_DblClick()
'On Error Resume Next
frm126_LM_No_ID = vbNullString

If frm126.MSFlexGrid1 <> vbNullString Then

    If IsNumeric(frm126.MSFlexGrid1) Then
    
        frm126_LM_No_ID = frm126.MSFlexGrid1.TextMatrix(frm126.MSFlexGrid1, 2) 'No. ID
        
        If frm126_LM_No_ID <> vbNullString Then
    
            PopupMenu frm126_pm_menu
            
        Else
            
            MsgBox "Tiada data.", vbExclamation, "Info"
            
        End If
        
    Else
    
        MsgBox "Tiada data.", vbExclamation, "Info"
        
    End If
    
End If
End Sub
