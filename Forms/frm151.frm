VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm151 
   Caption         =   "Invoice"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11280
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
   ScaleHeight     =   2745
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.TextBox TB4 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   3900
      TabIndex        =   10
      Text            =   "TB4"
      Top             =   1200
      Width           =   1620
   End
   Begin VB.CommandButton CMD2 
      Caption         =   "Simpan Data Trade In"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8760
      MouseIcon       =   "frm151.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frm151.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Simpan data"
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check Data"
      Height          =   855
      Left            =   7440
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton CMD1 
      Caption         =   "Simpan Data Jualan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5640
      MouseIcon       =   "frm151.frx":0CB4
      MousePointer    =   99  'Custom
      Picture         =   "frm151.frx":0FBE
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Simpan data"
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox TB3 
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1500
      TabIndex        =   0
      Text            =   "TB3"
      Top             =   1200
      Width           =   1500
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   360
      Left            =   1500
      TabIndex        =   1
      Top             =   240
      Width           =   4005
      _ExtentX        =   7064
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
      Format          =   141819904
      CurrentDate     =   41561
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   360
      Left            =   1500
      TabIndex        =   2
      Top             =   600
      Width           =   4005
      _ExtentX        =   7064
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
      Format          =   141819904
      CurrentDate     =   41561
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tahun :"
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   2520
      TabIndex        =   11
      Top             =   1230
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "** Sila masukkan no turutan terakhir yang telah digunakan."
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   7575
   End
   Begin VB.Label Label63 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tarikh Akhir :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   645
      Width           =   1275
   End
   Begin VB.Label Label62 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tarikh Mula :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   285
      Width           =   1275
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Mula :"
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   120
      TabIndex        =   3
      Top             =   1230
      Width           =   1275
   End
End
Attribute VB_Name = "frm151"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD1_Click()
'On Error GoTo aaa:
Dim rs1 As ADODB.Recordset
Dim TA As Date
Dim TM As Date
Dim LM_SEQ As Integer
Dim LM_JENIS_JUALAN As Double

Dim Err(5)
DATA_SAVE = 0

If frm151.TB3 = vbNullString Or (frm151.TB3 <> vbNullString And Not IsNumeric(frm151.TB3)) Then
    x = x + 1
    Err(x) = "Sila masukkan [No. Mula]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If frm151.TB4 = vbNullString Or (frm151.TB4 <> vbNullString And Not IsNumeric(frm151.TB4)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Tahun]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Teruskan?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then

        TM = frm151.DTPicker1
        TA = frm151.DTPicker2
        LM_SEQ = frm151.TB3
        LM_TAHUN = frm151.TB4
        LM_NOW = Now
        x = 0

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    
        strsql = "UPDATE 23_senarai_jualan set write_timestamp3 = NULL WHERE status_rekod = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    
        strsql = "UPDATE 22_jualan set write_timestamp3 = NULL WHERE status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    
        strsql = "UPDATE 35_senarai_servis set write_timestamp2 = NULL WHERE tarikh BETWEEN '" & TM & "' AND '" & TA & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    
        strsql = "UPDATE 40_tempahan_deposit set write_timestamp3 = NULL WHERE status_invoice = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    
        strsql = "UPDATE 42_tempahan_siap set write_timestamp3 = NULL WHERE status_invoice = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 22_jualan where status = 1 AND (menu = 0 OR menu = 1 OR menu = 3 OR menu = 4) AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC , ID ASC", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
        
            LM_NO_RESIT = vbNullString
            LM_JENIS_JUALAN = 0
            x = x + 1
    
            LM_SEQ = LM_SEQ + 1
            If Not IsNull(rs!no_resit) Then LM_NO_RESIT = rs!no_resit
            rs!no_resit = "BK-" & LM_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_SEQ, "000000")
            rs!write_timestamp3 = LM_NOW


            If Not IsNull(rs!Menu) Then
                If rs!Menu = 0 Then
                    Set rs1 = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                    strsql = "UPDATE 23_senarai_jualan set write_timestamp3='" & LM_NOW & "' , no_resit='" & "BK-" & LM_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_SEQ, "000000") & "' WHERE write_timestamp3 is NULL AND status_rekod = 1 AND no_resit='" & LM_NO_RESIT & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'"
                    
                    Set rs1 = cn.Execute(strsql)
                    Set rs1 = Nothing
                ElseIf rs!Menu = 1 Then
                    Set rs1 = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                    strsql = "UPDATE 35_senarai_servis set write_timestamp2='" & LM_NOW & "' , no_resit_servis='" & "BK-" & LM_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_SEQ, "000000") & "' WHERE write_timestamp2 is NULL AND no_resit_servis='" & LM_NO_RESIT & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'"
                    
                    Set rs1 = cn.Execute(strsql)
                    Set rs1 = Nothing
                ElseIf rs!Menu = 2 Then
                    Set rs1 = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                    strsql = "UPDATE 40_tempahan_deposit set write_timestamp3='" & LM_NOW & "' , no_resit_tempahan='" & "BK-" & LM_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_SEQ, "000000") & "' WHERE write_timestamp3 is NULL AND status_invoice = 1 AND no_resit_tempahan='" & LM_NO_RESIT & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'"
                    
                    Set rs1 = cn.Execute(strsql)
                    Set rs1 = Nothing
                ElseIf rs!Menu = 3 Then
                    Set rs1 = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                    strsql = "UPDATE 42_tempahan_siap set write_timestamp3='" & LM_NOW & "' , no_resit_tempahan='" & "BK-" & LM_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_SEQ, "000000") & "' WHERE write_timestamp3 is NULL AND status_invoice = 1 AND no_resit_tempahan='" & LM_NO_RESIT & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'"
                    
                    Set rs1 = cn.Execute(strsql)
                    Set rs1 = Nothing
                End If
            
            End If

            rs.Update
            
            rs.MoveNext
        Wend
               
        rs.Close
        Set rs = Nothing
        
        LogAct_Memory = "Maintenance (Jualan)." & TM & " hingga " & TA & " <" & frm151.TB3 & "><" & x & ">"
        LogDate_Memory = LM_NOW
        Call UpdateLog_Database
    
        MsgBox "Selesai. " & x
        
    End If

End If

Exit Sub
aaa:

MsgBox "Error " & LM_SEQ
End Sub

Private Sub CMD2_Click()
'On Error GoTo aaa:
Dim rs1 As ADODB.Recordset
Dim TA As Date
Dim TM As Date
Dim LM_SEQ As Integer
Dim LM_JENIS_JUALAN As Double

Dim Err(5)
DATA_SAVE = 0

If frm151.TB3 = vbNullString Or (frm151.TB3 <> vbNullString And Not IsNumeric(frm151.TB3)) Then
    x = x + 1
    Err(x) = "Sila masukkan [No. Mula]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If frm151.TB4 = vbNullString Or (frm151.TB4 <> vbNullString And Not IsNumeric(frm151.TB4)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Tahun]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Teruskan?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then

        TM = frm151.DTPicker1
        TA = frm151.DTPicker2
        LM_SEQ = frm151.TB3
        LM_TAHUN = frm151.TB4
        LM_NOW = Now
        x = 0

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    
        strsql = "UPDATE 16_gold_bar_belian set write_timestamp3 = NULL WHERE status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    
        strsql = "UPDATE data_database set write_timestamp3 = NULL WHERE bill_No_Trade_In is not null AND tarikh_Belian BETWEEN '" & TM & "' AND '" & TA & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
    
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 16_gold_bar_belian where status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC , no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
        
            LM_NO_RESIT = vbNullString
            LM_JENIS_JUALAN = 0
            x = x + 1
    
            LM_SEQ = LM_SEQ + 1
            If Not IsNull(rs!no_rujukan) Then LM_NO_RESIT = rs!no_rujukan
            rs!no_rujukan = "PV-" & LM_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_SEQ, "000000")
            rs!write_timestamp3 = LM_NOW

            Set rs1 = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            strsql = "UPDATE data_database set write_timestamp3='" & LM_NOW & "' , bill_No_Trade_In='" & "PV-" & LM_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_SEQ, "000000") & "' WHERE write_timestamp3 is NULL AND bill_No_Trade_In='" & LM_NO_RESIT & "' AND tarikh_Belian BETWEEN '" & TM & "' AND '" & TA & "'"
            
            Set rs1 = cn.Execute(strsql)
            Set rs1 = Nothing

            Set rs1 = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            strsql = "UPDATE 22_jualan set no_resit_trade_in='" & "PV-" & LM_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(LM_SEQ, "000000") & "' WHERE no_resit_trade_in='" & LM_NO_RESIT & "'"
            
            Set rs1 = cn.Execute(strsql)
            Set rs1 = Nothing
            
            rs.Update
            
            rs.MoveNext
        Wend
               
        rs.Close
        Set rs = Nothing
        
        LogAct_Memory = "Maintenance (Trade In)." & TM & " hingga " & TA & " <" & frm151.TB3 & "><" & x & ">"
        LogDate_Memory = LM_NOW
        Call UpdateLog_Database
    
        MsgBox "Selesai. " & x
        
    End If

End If

Exit Sub
aaa:

MsgBox "Error " & LM_SEQ
End Sub

Private Sub Command2_Click()
'On Error GoTo aaa:
Dim rs1 As ADODB.Recordset
Dim TA As Date
Dim TM As Date
Dim LM_BIL_22 As Double
Dim LM_BERAT_22 As Double
Dim LM_BIL_23 As Double
Dim LM_BERAT_23 As Double
Dim LM_TARIKH_22 As Date
Dim LM_TARIKH_23 As Date

Note = "Teruskan?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then

    TM = frm151.DTPicker1
    TA = frm151.DTPicker2
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 22_jualan where status = 1 AND menu = 0 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by no_resit ASC", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
    
        LM_BIL_22 = 0
        LM_BERAT_22 = 0
        LM_BIL_23 = 0
        LM_BERAT_23 = 0
        LM_TARIKH_22 = "2010-01-01"
        LM_TARIKH_23 = "2010-01-02"
        
        If Not IsNull(rs!kuantiti_barang) Then LM_BIL_22 = rs!kuantiti_barang
        If Not IsNull(rs!JUMLAH_BERAT) Then LM_BERAT_22 = rs!JUMLAH_BERAT
        If Not IsNull(rs!tarikh) Then LM_TARIKH_22 = rs!tarikh
        
            
        Set rs1 = New ADODB.Recordset
        'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        'rs1.Open "select berat_jualan , tarikh from 23_senarai_jualan where status_rekod = 1 AND no_resit='" & rs!no_resit & "' AND (write_timestamp='" & rs!write_timestamp & "' or write_timestamp='" & rs!write_timestamp2 & "')", cn, adOpenKeyset, adLockOptimistic
        rs1.Open "select berat_jualan , tarikh from 23_senarai_jualan where status_rekod = 1 AND no_resit='" & rs!no_resit & "'", cn, adOpenKeyset, adLockOptimistic
        
        While rs1.EOF = False
            LM_BIL_23 = LM_BIL_23 + 1
            If Not IsNull(rs1!berat_jualan) Then LM_BERAT_23 = LM_BERAT_23 + rs1!berat_jualan
            If Not IsNull(rs1!tarikh) Then LM_TARIKH_23 = rs1!tarikh
            rs1.MoveNext
        Wend
        
        rs1.Close
        Set rs1 = Nothing
        
        If LM_BIL_22 <> LM_BIL_23 Or Format(LM_BERAT_22, "0.00") <> Format(LM_BERAT_23, "0.00") Or LM_TARIKH_22 <> LM_TARIKH_23 Then
        
            sFilename = App.Path & "\error_resit.txt"

            ' Open the file to write
            filenumber = FreeFile
            Open sFilename For Append As #filenumber
            
            Print #filenumber, Now & " " & rs!Menu & " / " & rs!no_resit & " / " & LM_BIL_22 & " " & LM_BIL_23 & " / " & LM_BERAT_22 & " " & LM_BERAT_23 & " / " & LM_TARIKH_22 & " " & LM_TARIKH_23
            
            Close #filenumber
        
        End If
        
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
    MsgBox "Selesai"
End If

Exit Sub
aaa:

MsgBox "Error " & LM_SEQ

End Sub

Private Sub Command1_Click()
'On Error GoTo aaa:
Dim rs1 As ADODB.Recordset
Dim TA As Date
Dim TM As Date
Dim LM_BIL_22 As Double
Dim LM_BERAT_22 As Double
Dim LM_BIL_23 As Double
Dim LM_BERAT_23 As Double
Dim LM_TARIKH_22 As Date
Dim LM_TARIKH_23 As Date

Note = "Teruskan?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then

    TM = frm151.DTPicker1
    TA = frm151.DTPicker2
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    'rs.Open "select * from 22_jualan where status = 1 AND (menu = 0 OR menu = 1 OR menu = 3 OR menu = 4) AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC , ID ASC", cn, adOpenKeyset, adLockOptimistic
    rs.Open "select * from 22_jualan where status = 1 AND menu = 0 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by no_resit ASC", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
    
        LM_BIL_22 = 0
        LM_BERAT_22 = 0
        LM_BIL_23 = 0
        LM_BERAT_23 = 0
        LM_TARIKH_22 = "2010-01-01"
        LM_TARIKH_23 = "2010-01-02"
        
        If Not IsNull(rs!kuantiti_barang) Then LM_BIL_22 = rs!kuantiti_barang
        If Not IsNull(rs!JUMLAH_BERAT) Then LM_BERAT_22 = rs!JUMLAH_BERAT
        If Not IsNull(rs!tarikh) Then LM_TARIKH_22 = rs!tarikh
        
            
        Set rs1 = New ADODB.Recordset
        'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        'rs1.Open "select berat_jualan , tarikh from 23_senarai_jualan where status_rekod = 1 AND no_resit='" & rs!no_resit & "' AND (write_timestamp='" & rs!write_timestamp & "' or write_timestamp='" & rs!write_timestamp2 & "')", cn, adOpenKeyset, adLockOptimistic
        rs1.Open "select berat_jualan , tarikh from 23_senarai_jualan where status_rekod = 1 AND no_resit='" & rs!no_resit & "'", cn, adOpenKeyset, adLockOptimistic
        
        While rs1.EOF = False
            LM_BIL_23 = LM_BIL_23 + 1
            If Not IsNull(rs1!berat_jualan) Then LM_BERAT_23 = LM_BERAT_23 + rs1!berat_jualan
            If Not IsNull(rs1!tarikh) Then LM_TARIKH_23 = rs1!tarikh
            rs1.MoveNext
        Wend
        
        rs1.Close
        Set rs1 = Nothing
        
        If LM_BIL_22 <> LM_BIL_23 Or Format(LM_BERAT_22, "0.00") <> Format(LM_BERAT_23, "0.00") Or LM_TARIKH_22 <> LM_TARIKH_23 Then
        
            sFilename = App.Path & "\error_resit.txt"

            ' Open the file to write
            filenumber = FreeFile
            Open sFilename For Append As #filenumber
            
            Print #filenumber, Now & " " & rs!Menu & " / " & rs!no_resit & " / " & LM_BIL_22 & " " & LM_BIL_23 & " / " & LM_BERAT_22 & " " & LM_BERAT_23 & " / " & LM_TARIKH_22 & " " & LM_TARIKH_23
            
            Close #filenumber
        
        End If
        
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
    MsgBox "Selesai"
End If

Exit Sub
aaa:

MsgBox "Error " & LM_SEQ
End Sub

Private Sub Form_Load()
'on error resume next
frm151.Picture = MDI_frm1.Picture

frm151.DTPicker1 = DateTime.Date
frm151.DTPicker2 = DateTime.Date
frm151.TB3 = 1
frm151.TB4 = 1
End Sub
