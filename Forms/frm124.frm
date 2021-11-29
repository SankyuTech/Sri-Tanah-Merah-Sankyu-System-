VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm124 
   Caption         =   "Rekod penggunaan stok barangan trade in dan stok barangan potong"
   ClientHeight    =   13545
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
   Icon            =   "frm124.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13545
   ScaleWidth      =   23880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD21 
      BackColor       =   &H00FFFFFF&
      Height          =   650
      Left            =   6840
      MaskColor       =   &H00400000&
      MouseIcon       =   "frm124.frx":0ECA
      MousePointer    =   99  'Custom
      Picture         =   "frm124.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Paparan sebelumnya"
      Top             =   11040
      Width           =   1100
   End
   Begin VB.CommandButton CMD22 
      BackColor       =   &H00FFFFFF&
      Height          =   650
      Left            =   8040
      MaskColor       =   &H00400000&
      MouseIcon       =   "frm124.frx":1B13
      MousePointer    =   99  'Custom
      Picture         =   "frm124.frx":1E1D
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Paparan seterusnya"
      Top             =   11040
      Width           =   1100
   End
   Begin VB.ComboBox CBB2 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Supplier"
      Height          =   360
      ItemData        =   "frm124.frx":2743
      Left            =   2040
      List            =   "frm124.frx":2745
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1680
      Width           =   7005
   End
   Begin VB.ComboBox CBB1 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Supplier"
      Height          =   360
      ItemData        =   "frm124.frx":2747
      Left            =   2040
      List            =   "frm124.frx":2749
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1320
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
      Left            =   120
      TabIndex        =   1
      Top             =   200
      Width           =   200
   End
   Begin VB.CommandButton CMD3 
      BackColor       =   &H000080FF&
      Caption         =   "Report"
      Height          =   405
      Left            =   3720
      MaskColor       =   &H00400000&
      MouseIcon       =   "frm124.frx":274B
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Report"
      Top             =   2080
      Width           =   2385
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   360
      Left            =   2025
      TabIndex        =   3
      Top             =   540
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
      Format          =   413859840
      CurrentDate     =   41561
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   360
      Left            =   2025
      TabIndex        =   4
      Top             =   900
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
      Format          =   413859840
      CurrentDate     =   41561
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   8085
      Left            =   240
      TabIndex        =   18
      ToolTipText     =   "Sila ""Double Click"" untuk menu seterusnya."
      Top             =   2880
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   14261
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
      TabIndex        =   30
      Top             =   11520
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Baki :"
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
      TabIndex        =   29
      Top             =   11520
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
      TabIndex        =   28
      Top             =   11280
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jual / Guna :"
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
      TabIndex        =   27
      Top             =   11280
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
      TabIndex        =   26
      Top             =   11040
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Berat Asal :"
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
      TabIndex        =   25
      Top             =   11040
      Width           =   1455
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
      Left            =   6360
      TabIndex        =   23
      Top             =   11040
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
      Left            =   5760
      TabIndex        =   22
      Top             =   11040
      Width           =   375
   End
   Begin VB.Label L69_Text 
      BackColor       =   &H8000000C&
      Caption         =   "L69_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3720
      TabIndex        =   21
      Top             =   11520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label L70_Text 
      BackColor       =   &H8000000C&
      Caption         =   "L70_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3720
      TabIndex        =   20
      Top             =   11160
      Visible         =   0   'False
      Width           =   855
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
      TabIndex        =   19
      Top             =   2640
      Width           =   15495
   End
   Begin VB.Label L9_Text 
      BackColor       =   &H00C0C0FF&
      Caption         =   "L9_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9120
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Purity * "
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   315
      TabIndex        =   14
      Top             =   1680
      Width           =   1860
   End
   Begin VB.Label L8_Text 
      BackColor       =   &H00C0C0FF&
      Caption         =   "L8_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9120
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label L7_Text 
      BackColor       =   &H00C0C0FF&
      Caption         =   "L7_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9120
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label L6_Text 
      BackColor       =   &H00C0C0FF&
      Caption         =   "L6_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9120
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label L5_Text 
      BackColor       =   &H00C0C0FF&
      Caption         =   "L5_Text"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   9120
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Urusan * "
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   315
      TabIndex        =   8
      Top             =   1320
      Width           =   1860
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Sila klik jika ingin melihat senarai rekod di dalam tempoh tarikh di bawah."
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   360
      TabIndex        =   7
      Top             =   165
      Width           =   8370
   End
   Begin VB.Label Label63 
      BackStyle       =   0  'Transparent
      Caption         =   "Tarikh Akhir "
      Height          =   300
      Left            =   315
      TabIndex        =   6
      Top             =   945
      Width           =   2895
   End
   Begin VB.Label Label62 
      BackStyle       =   0  'Transparent
      Caption         =   "Tarikh Mula "
      Height          =   300
      Left            =   315
      TabIndex        =   5
      Top             =   585
      Width           =   2535
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
      Left            =   4440
      TabIndex        =   24
      Top             =   11040
      Width           =   2295
   End
   Begin VB.Menu frm124_pm_menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu frm124_sm_excel 
         Caption         =   "Report excel"
      End
   End
End
Attribute VB_Name = "frm124"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD21_Click()
'on error resume next
Dim frm124_LM_CURR_PAGE As Double
Dim frm124_LM_TOTAL_PAGE As Double

frm124_LM_CURR_PAGE = 0
frm124_LM_TOTAL_PAGE = 0

If frm124.L67_Text <> vbNullString And IsNumeric(frm124.L67_Text) Then
    If frm124.L68_Text <> vbNullString And IsNumeric(frm124.L68_Text) Then
        frm124_LM_CURR_PAGE = frm124.L67_Text
        frm124_LM_TOTAL_PAGE = frm124.L68_Text
        
        If frm124_LM_CURR_PAGE <> 1 And frm124_LM_CURR_PAGE <> 0 Then
        
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
                    
            Call frm124_report_trade_in_header
            Call frm124_report_trade_in
            
        End If

    End If
End If
End Sub
Private Sub CMD22_Click()
'on error resume next
Dim frm124_LM_CURR_PAGE As Double
Dim frm124_LM_TOTAL_PAGE As Double

frm124_LM_CURR_PAGE = 0
frm124_LM_TOTAL_PAGE = 0

If frm124.L67_Text <> vbNullString And IsNumeric(frm124.L67_Text) Then
    If frm124.L68_Text <> vbNullString And IsNumeric(frm124.L68_Text) Then
        frm124_LM_CURR_PAGE = frm124.L67_Text
        frm124_LM_TOTAL_PAGE = frm124.L68_Text
        
        If frm124_LM_CURR_PAGE < frm124_LM_TOTAL_PAGE Then
        
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
            
            Call frm124_report_trade_in_header
            Call frm124_report_trade_in
            
        End If
    End If
End If
End Sub
Private Sub CMD3_Click()
'On Error Resume Next
If frm124.CB1 = 1 Then 'Pilihan tarikh
    frm124.L5_Text = 1
Else
    frm124.L5_Text = 0
End If

If frm124.CBB1 = vbNullString Then

    MsgBox "Sila buat pilihan jenis urusan.", vbInformation, "Info"
    
    Exit Sub
    
End If

If frm124.CBB2 = vbNullString Then

    MsgBox "Sila buat pilihan purity.", vbInformation, "Info"
    
    Exit Sub
    
End If

frm124.L6_Text = frm124.DTPicker1 'Tarikh mula
frm124.L7_Text = frm124.DTPicker2 'Tarikh akhir

frm124.L8_Text = frm124.CBB1 'Jenis urusan
frm124.L9_Text = frm124.CBB2 'Purity

frm124.L69_Text = -1 'Titik Pencarian Data
frm124.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
frm124.L67_Text = 0 'Paparan Page ke-xxx
frm124.L68_Text = 0

GM_NEXT_PREV = 0

Call frm124_report_trade_in_header
Call frm124_report_trade_in
End Sub
Private Sub frm124_sm_excel_Click()
'on error resume next
Dim TM As Date
Dim TA As Date

LM_FOUND = 0
frm124_LM_No_ID = vbNullString

If frm124.MSFlexGrid1 <> vbNullString Then

    If IsNumeric(frm124.MSFlexGrid1) Then
    
        frm124_LM_No_ID = frm124.MSFlexGrid1.TextMatrix(frm124.MSFlexGrid1, 2) 'No. ID
        
        If frm124_LM_No_ID <> vbNullString Then

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
                    .Columns("B").ColumnWidth = 15 'Jenis
                    .Columns("C").ColumnWidth = 15 'Tarikh
                    .Columns("D").ColumnWidth = 20 'No. Rujukan
                    .Columns("E").ColumnWidth = 15 'Purity
                    .Columns("F").ColumnWidth = 15 'Berat (g)
                
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
                    
                    .Cells(7, 1) = frm124.L14_Text
                    
                    .Cells(8, 1) = "No."
                    .Cells(8, 2) = "Jenis"
                    .Cells(8, 3) = "Tarikh"
                    .Cells(8, 4) = "No. Rujukan"
                    .Cells(8, 5) = "Purity"
                    .Cells(8, 6) = "Berat (g)"
                    
                    For i = 1 To 6
                        .Cells(8, i).HorizontalAlignment = xlCenter
                        .Cells(8, i).Interior.ColorIndex = 15
                        .Cells(8, i).WrapText = True
                        .Cells(8, i).Borders.LineStyle = xlContinuous
                    Next i
            
                    If frm124.L5_Text = "1" Then '0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh
                    
                        TM = frm124.L6_Text 'Tarikh mula
                        TA = frm124.L7_Text 'Tarikh akhir
                    
                    End If
                    
                    If frm124.L8_Text = "semua urusan" Then
                        
                        frm124_LM_SEARCH_1 = 0
                        frm124_LM_SEARCH_2 = 1
                        
                    ElseIf frm124.L8_Text = "GDN" Then
                        
                        frm124_LM_SEARCH_1 = 0
                        frm124_LM_SEARCH_2 = 0
                    
                    ElseIf frm124.L8_Text = "Jualan" Then
                        
                        frm124_LM_SEARCH_1 = 1
                        frm124_LM_SEARCH_2 = 1
                        
                    End If
                    
                    If frm124.L9_Text = "semua purity" Then
                        
                        frm124_LM_SEARCH_3 = Null
                        frm124_LM_SEARCH_3_LOGIC = "<>"
                        
                    Else
                        
                        frm124_LM_SEARCH_3 = frm124.L9_Text
                        frm124_LM_SEARCH_3_LOGIC = "="
                    
                    End If
                    
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    If frm124.L5_Text = 0 Then rs.Open "select * from 85_penggunaan_ti where (menu='" & frm124_LM_SEARCH_1 & "' OR menu='" & frm124_LM_SEARCH_2 & "') AND purity " & frm124_LM_SEARCH_3_LOGIC & "'" & frm124_LM_SEARCH_3 & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
                    If frm124.L5_Text = 1 Then rs.Open "select * from 85_penggunaan_ti where (menu='" & frm124_LM_SEARCH_1 & "' OR menu='" & frm124_LM_SEARCH_2 & "') AND purity " & frm124_LM_SEARCH_3_LOGIC & "'" & frm124_LM_SEARCH_3 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic


                    While rs.EOF = False
                    
                        x = x + 1
                        .Cells(8 + x, 1) = x 'No.
                        .Cells(8 + x, 1).HorizontalAlignment = xlCenter
                        
                        If Not IsNull(rs!Menu) Then 'Jenis
                            
                            If rs!Menu = 0 Then
                                
                                .Cells(8 + x, 2) = "GDN"
                                
                            ElseIf rs!Menu = 1 Then
                                
                                .Cells(8 + x, 2) = "Jualan"
                                
                            End If
                            
                        End If
                        
                        If Not IsNull(rs!tarikh) Then .Cells(8 + x, 3) = "'" & rs!tarikh 'Tarikh
                        .Cells(8 + x, 3).HorizontalAlignment = xlCenter
                        
                        If Not IsNull(rs!no_rujukan) Then .Cells(8 + x, 4) = rs!no_rujukan 'No. Rujukan
                        
                        If Not IsNull(rs!purity) Then .Cells(8 + x, 5) = rs!purity 'Purity
                        .Cells(8 + x, 5).HorizontalAlignment = xlCenter
                        
                        .Cells(8 + x, 6).HorizontalAlignment = xlRight
                        If Not IsNull(rs!Berat) Then
                            .Cells(8 + x, 6) = Format(rs!Berat, "#,##0.00") 'Berat (g)
                            .Cells(8 + x, 6).NumberFormat = "#,##0.00"
                        End If
                                            
                        For Col = 1 To 6
                            .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
                        Next Col
                            
                        rs.MoveNext
                        
                    Wend
                    
                    rs.Close
                    Set rs = Nothing
                    
                    Y = 0
                    Y = x + 2
                    
                    .Cells(8 + Y, 1) = "Berat Asal : " & frm124.L10_Text
                    .Cells(8 + Y, 1).Font.Bold = True
                    
                    Y = Y + 1
                    .Cells(8 + Y, 1) = "Jual / Guna : " & frm124.L11_Text
                    .Cells(8 + Y, 1).Font.Bold = True
                        
                    Y = Y + 1
                    .Cells(8 + Y, 1) = "Baki : " & frm124.L12_Text
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

Private Sub MSFlexGrid1_DblClick()
'On Error Resume Next
frm124_LM_No_ID = vbNullString

If frm124.MSFlexGrid1 <> vbNullString Then

    If IsNumeric(frm124.MSFlexGrid1) Then
    
        frm124_LM_No_ID = frm124.MSFlexGrid1.TextMatrix(frm124.MSFlexGrid1, 2) 'No. ID
        
        If frm124_LM_No_ID <> vbNullString Then
    
            PopupMenu frm124_pm_menu
            
        Else
            
            MsgBox "Tiada data.", vbExclamation, "Info"
            
        End If
        
    Else
    
        MsgBox "Tiada data.", vbExclamation, "Info"
        
    End If
    
End If
End Sub
