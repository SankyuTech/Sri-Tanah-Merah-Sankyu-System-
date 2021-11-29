Attribute VB_Name = "Module14"
Sub Frm54_KategoriUpah()
'On Error Resume Next
Frm54.MSFlexGrid1.Clear
Frm54.MSFlexGrid1.RowHeight(0) = 600
Frm54.MSFlexGrid1.FormatString = "No.|<No.|<Kategori Upah"

Frm54.MSFlexGrid1.Rows = 1
Frm54.MSFlexGrid1.ColWidth(0) = 600
Frm54.MSFlexGrid1.ColWidth(1) = 0
Frm54.MSFlexGrid1.ColWidth(2) = 6100

Frm54.MSFlexGrid2.Clear
Frm54.MSFlexGrid2.RowHeight(0) = 600
Frm54.MSFlexGrid2.FormatString = "No.|<No.|<Kategori Upah|<Tetapan Upah (RM)"

Frm54.MSFlexGrid2.Rows = 1
Frm54.MSFlexGrid2.ColWidth(0) = 600
Frm54.MSFlexGrid2.ColWidth(1) = 0
Frm54.MSFlexGrid2.ColWidth(2) = 3100
Frm54.MSFlexGrid2.ColWidth(3) = 3000

Frm54.CBB2.Clear
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from tetapanupah", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If rs!KategoriUpah <> vbNullString Then
        Frm54.CBB2.AddItem rs!KategoriUpah
        x = x + 1
        Frm54.MSFlexGrid1.Rows = x + 1
        Frm54.MSFlexGrid1.TextMatrix(x, 0) = x
        Frm54.MSFlexGrid1.TextMatrix(x, 1) = x
        If Not IsNull(rs!KategoriUpah) Then Frm54.MSFlexGrid1.TextMatrix(x, 2) = rs!KategoriUpah 'Kategori Upah
        Frm54.MSFlexGrid2.Rows = x + 1
        Frm54.MSFlexGrid2.TextMatrix(x, 0) = x
        Frm54.MSFlexGrid2.TextMatrix(x, 1) = x
        If Not IsNull(rs!KategoriUpah) Then Frm54.MSFlexGrid2.TextMatrix(x, 2) = rs!KategoriUpah 'Kategori Upah
        If Not IsNull(rs!tetapanupah) Then Frm54.MSFlexGrid2.TextMatrix(x, 3) = rs!tetapanupah 'Tetapan Upah
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
End Sub
Sub frm49_Default()
'On Error Resume Next
Frm49.TB1 = vbNullString
Frm49.TB2 = vbNullString
Frm49.TB3 = vbNullString
Frm49.TB4 = vbNullString
Frm49.TB5 = vbNullString
Frm49.TB6 = vbNullString
Frm49.TB7 = vbNullString
Frm49.TB8 = vbNullString
Frm49.TB9 = vbNullString
Frm49.TB10 = vbNullString
Frm49.TB12 = vbNullString
Frm49.TB13 = vbNullString
Frm49.TB14 = "0.00"
Frm49.TB15 = "0.00"
Frm49.TB16 = vbNullString
Frm49.TB19 = vbNullString
Frm49.CB1 = 0
Frm49.CB2 = 0
Frm49.CB3 = 0
Frm49.CB4 = 0
Frm49.L4_Text = 0 'Memory : ID

Frm49.TB2.Locked = False
Frm49.TB13.Locked = False
Frm49.CMD4.Visible = True
Frm49.CMD6.Visible = False
Frm49.CMD7.Visible = False

Frm49.L1_Label.Visible = False
Frm49.L2_Label.Visible = False
Frm49.DTPicker4.Visible = False

Frm49.TB13.Locked = False
Frm49.TB13.BackColor = &HFFFFFF
            
Frm49.CBB1.Clear
With Frm49.CBB1
    .AddItem "Aktif"
    .AddItem "Berhenti"
End With

If MDI_frm1.L4_Text = "HQ" Then
    Frm49.CB1.Enabled = True
Else
    Frm49.CB1.Enabled = False
End If

Frm49.DTPicker1 = DateTime.Date
Frm49.DTPicker4 = DateTime.Date

'Set rs = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

'If Not rs.EOF Then
'    If rs!Default1 = "Default" Then
'        Frm49.TB4 = rs!EmpNo 'No. Pekerja
'    End If
'End If

'rs.Close
'Set rs = Nothing
End Sub
Sub frm49_disable_form()
'On Error Resume Next
Frm49.Frame1.Top = 120
Frm49.Frame1.Left = 1680

Frm49.Frame2.Top = 120
Frm49.Frame2.Left = 1680

Frm49.Frame3.Top = 120
Frm49.Frame3.Left = 1680

Frm49.Frame1.Visible = False
Frm49.Frame2.Visible = False
Frm49.Frame3.Visible = False
End Sub
Sub frm49_cawangan()
'On Error Resume Next
Frm49.CBB2.Clear
Frm49.CBB3.Clear

Frm49.CBB3.AddItem "Semua cawangan"
'Frm49.CBB3.AddItem "HQ"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where status='" & 1 & "' order by cawangan ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!cawangan) Then Frm49.CBB2.AddItem rs!cawangan
    If Not IsNull(rs!cawangan) Then Frm49.CBB3.AddItem rs!cawangan
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm49.CBB3 = "Semua cawangan"

If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
    
    Frm49.CBB2.Enabled = True
    
Else
    
    Frm49.CBB2.Enabled = False
    Frm49.CBB2 = MDI_frm1.L20_Text
    
    Frm49.CBB3.Enabled = False
    Frm49.CBB3 = MDI_frm1.L20_Text
    
End If
End Sub
Sub frm49_senarai_staff_header()
'On Error Resume Next
With Frm49.LV2
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear
    
    Frm49.LV2.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "Nama", 3500
    .ColumnHeaders.Add 5, , "No. Kad Pengenalan", 1700
    .ColumnHeaders.Add 6, , "No. Pekerja", 1500
    .ColumnHeaders.Add 7, , "No. Telefon", 2000
    .ColumnHeaders.Add 8, , "No. EPF", 1500
    .ColumnHeaders.Add 9, , "No. Income Tax", 1800
    .ColumnHeaders.Add 10, , "Jawatan", 2000
    .ColumnHeaders.Add 11, , "Status", 1500
    .ColumnHeaders.Add 12, , "Tarikh Masuk", 1800
    .ColumnHeaders.Add 13, , "Tarikh Berhenti", 1800
    .ColumnHeaders.Add 14, , "User Level", 1700
    .ColumnHeaders.Add 15, , "Username", 1800
    .ColumnHeaders.Add 16, , "Password", 1800
    .ColumnHeaders.Add 17, , "Cawangan", 2000
    .ColumnHeaders.Add 18, , "E-mail", 3500
    .ColumnHeaders.Add 19, , "Komisen", 1800
    .ColumnHeaders.Add 20, , "Gaji (RM)", 1800, 1
    .ColumnHeaders.Add 21, , "Elaun (RM)", 1800, 1
    .ColumnHeaders.Add 22, , "Nama Bank", 3500
    .ColumnHeaders.Add 23, , "No. Akaun", 3500
    
End With
End Sub
Sub frm49_senarai_staff()
'On Error Resume Next
Dim frm49_LM_TOTAL_PAGE As Double
Dim frm49_field_1 As String

frm49_LM_TOTAL_PAGE = 0
frm49_PAGE_SIZE = 34
x = 0

re_gen_report:

frm49_LM_SEARCH_1 = Frm49.L6_Text
frm49_LM_SEARCH_1_LOGIC = "="

If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then

    frm49_LM_SEARCH_2 = Null
    frm49_LM_SEARCH_2_LOGIC = "<>"
    
Else

    frm49_LM_SEARCH_2 = MDI_frm1.L20_Text
    frm49_LM_SEARCH_2_LOGIC = "="
    
End If

If Frm49.L7_Text = "0" Then '0 : Carian mengikut cawangan , 1 : Carian mengikut maklumat pekerja

    frm49_field_1 = "cawangan"
    
    If Frm49.L6_Text = "Semua cawangan" Then
        frm49_LM_SEARCH_1 = Null
        frm49_LM_SEARCH_1_LOGIC = "<>"
    Else
        frm49_LM_SEARCH_1 = Frm49.L6_Text
        frm49_LM_SEARCH_1_LOGIC = "="
    End If
    
ElseIf Frm49.L7_Text = "1" Then

    If Frm49.Option1 = True Then frm49_field_1 = "NoIC"
    If Frm49.Option2 = True Then frm49_field_1 = "Nama"
    If Frm49.Option3 = True Then frm49_field_1 = "NoPekerja"
    
    frm49_LM_SEARCH_1 = Frm49.L6_Text
    frm49_LM_SEARCH_1_LOGIC = "="
        
End If

LM_START_ROW = Frm49.L69_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + frm49_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm49.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - frm49_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        Frm49.L67_Text = 1
    End If
End If

frm49_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where cawangan " & frm49_LM_SEARCH_2_LOGIC & "'" & frm49_LM_SEARCH_2 & "' AND " & frm49_field_1 & " " & frm49_LM_SEARCH_1_LOGIC & "'" & frm49_LM_SEARCH_1 & "' AND (user_level <> 4 AND user_level <> 5 AND user_level <> 6 AND user_level <> 7) order by nama ASC LIMIT " & LM_START_ROW & "," & frm49_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1
    If frm49_LM_PAGE_FOUND = 0 Then
        If Frm49.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm49.L67_Text = Frm49.L67_Text + 1 'Paparan Page ke-xxx
                frm49_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm49.L67_Text) Then
                    If Frm49.L67_Text <> 1 Then
                        Frm49.L67_Text = Frm49.L67_Text - 1 'Paparan Page ke-xxx
                        frm49_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If

    Y = ((Frm49.L67_Text - 1) * frm49_PAGE_SIZE) + x

    With Frm49.LV2.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!Nama) Then 'Nama
            .ListSubItems.Add , , rs!Nama
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!NoIC) Then 'No. Kad Pengenalan
            .ListSubItems.Add , , rs!NoIC
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!NoPekerja) Then 'No. Pekerja
            .ListSubItems.Add , , rs!NoPekerja
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!NoTel) Then 'No. Tel
            .ListSubItems.Add , , rs!NoTel
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!NoKWSP) Then 'No. EPF
            .ListSubItems.Add , , rs!NoKWSP
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!NoSocso) Then 'No. Income Tax
            .ListSubItems.Add , , rs!NoSocso
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!Jawatan) Then 'Jawatan
            .ListSubItems.Add , , rs!Jawatan
        Else
            .ListSubItems.Add , , ""
        End If
    
        If Not IsNull(rs!Status) Then 'Status
            .ListSubItems.Add , , rs!Status
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!TarikhMula) Then 'Tarikh masuk
            .ListSubItems.Add , , rs!TarikhMula
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!TarikhBerhenti) Then 'Tarikh berenti
            .ListSubItems.Add , , rs!TarikhBerhenti
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!user_level) Then 'Tarikh berenti

            If rs!user_level = 1 Then
                .ListSubItems.Add , , "Admin" 'User Level
            ElseIf rs!user_level = 2 Then
                .ListSubItems.Add , , "Manager" 'User Level
            ElseIf rs!user_level = 3 Then
                .ListSubItems.Add , , "Staff" 'User Level
            End If
    
        Else
            .ListSubItems.Add , , "Staff" 'User Level
        End If
        
        If Not IsNull(rs!Samaran) Then 'username
            .ListSubItems.Add , , rs!Samaran
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Password) Then 'Password
            .ListSubItems.Add , , rs!Password
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!cawangan) Then 'Cawangan
            .ListSubItems.Add , , rs!cawangan
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!mail) Then 'E-mail
            .ListSubItems.Add , , rs!mail
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!komisen) Then
            
            If rs!komisen = 0 Then
                .ListSubItems.Add , , "Tiada"
            ElseIf rs!komisen = 1 Then
                .ListSubItems.Add , , "Ada"
            End If
            
        Else
        
            .ListSubItems.Add , , "Tiada"
        
        End If
        
        If Not IsNull(rs!Gaji) Then 'Gaji (RM)
            .ListSubItems.Add , , Format(rs!Gaji, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If

        If Not IsNull(rs!Elaun) Then 'Elaun (RM)
            .ListSubItems.Add , , Format(rs!Elaun, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        
        If Not IsNull(rs!alamat2) Then 'Nama Bank
            .ListSubItems.Add , , rs!alamat2
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!alamat3) Then 'No. Akaun
            .ListSubItems.Add , , rs!alamat3
        Else
            .ListSubItems.Add , , ""
        End If
    
    End With

    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from employee where cawangan " & frm49_LM_SEARCH_2_LOGIC & "'" & frm49_LM_SEARCH_2 & "' AND " & frm49_field_1 & " " & frm49_LM_SEARCH_1_LOGIC & "'" & frm49_LM_SEARCH_1 & "' AND (user_level <> 4 AND user_level <> 5 AND user_level <> 6 AND user_level <> 7)", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    frm49_LM_TOTAL_PAGE = Format(rs(0) / frm49_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, frm49_LM_TOTAL_PAGE, ".") <> 0 Then
    
        frm49_LM_PAGE = Split(frm49_LM_TOTAL_PAGE, ".")(0)
        frm49_LM_PAGE_LEBIHAN = Split(frm49_LM_TOTAL_PAGE, ".")(1)
        
        If frm49_LM_PAGE_LEBIHAN <> "00" Then
            Frm49.L68_Text = frm49_LM_PAGE + 1
        Else
            Frm49.L68_Text = frm49_LM_PAGE
        End If
        
    Else
    
        Frm49.L68_Text = frm49_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm49.L68_Text = 0
    End If
Else
    Frm49.L68_Text = 0
End If

If Not IsNull(rs(0)) Then Frm49.L71_Text = rs(0)

rs.Close
Set rs = Nothing

If x <> 0 Then
    Frm49.L69_Text = LM_START_ROW
End If

If Frm49.L67_Text <> vbNullString And IsNumeric(Frm49.L67_Text) Then
    If Frm49.L68_Text <> vbNullString And IsNumeric(Frm49.L68_Text) Then
        frm49_LM_CURR_PAGE = Frm49.L67_Text
        frm49_LM_TOTAL_PAGE = Frm49.L68_Text
        
        If frm49_LM_CURR_PAGE > frm49_LM_TOTAL_PAGE Then
            
            Frm49.L67_Text = Frm49.L67_Text - 1
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            GoTo re_gen_report:
            
        End If
    End If
End If

Frm49.Frame2.Visible = True
Frm49.Frame3.Visible = False
End Sub
Sub Frm49_EmpList()
'On Error Resume Next
Frm49.MSFlexGrid1.Clear
Frm49.MSFlexGrid1.RowHeight(0) = 800
'Frm49.MSFlexGrid1.FormatString = "No.|<No.|<Nama|<No. Kad Pengenalan|<No. Pekerja|<No. KWSP|<No. Socso|<No. Tel|<Jawatan|<Nama Samaran|<Gaji (RM)|<Elaun (RM)|<Elaun Profit|<Elaun Investor (Small)|<Elaun Investor (Big)"
Frm49.MSFlexGrid1.FormatString = "No.|<No.|<Nama|<No. Kad Pengenalan|<No. Pekerja|<No. EPF|<No. Income Tax|<No. Tel|<Jawatan|<Nama Samaran|<Password|<User Level|<Gaji (RM)|<Elaun (RM)|<Komisen Jualan|<Status"

Frm49.MSFlexGrid1.Rows = 1
Frm49.MSFlexGrid1.ColWidth(0) = 600
Frm49.MSFlexGrid1.ColWidth(1) = 0
Frm49.MSFlexGrid1.ColWidth(2) = 4800
Frm49.MSFlexGrid1.ColWidth(3) = 1700
Frm49.MSFlexGrid1.ColWidth(4) = 1200
Frm49.MSFlexGrid1.ColWidth(5) = 1200
Frm49.MSFlexGrid1.ColWidth(6) = 1200
Frm49.MSFlexGrid1.ColWidth(7) = 1200
Frm49.MSFlexGrid1.ColWidth(8) = 1200
Frm49.MSFlexGrid1.ColWidth(9) = 1200
Frm49.MSFlexGrid1.ColWidth(10) = 1200
Frm49.MSFlexGrid1.ColWidth(11) = 1200
Frm49.MSFlexGrid1.ColWidth(12) = 1000
Frm49.MSFlexGrid1.ColWidth(13) = 1000
Frm49.MSFlexGrid1.ColWidth(14) = 1000
Frm49.MSFlexGrid1.ColWidth(15) = 1000
'Frm49.MSFlexGrid1.ColWidth(14) = 1200

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where (user_level <> 4 AND user_level <> 5)", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm49.MSFlexGrid1.Rows = x + 1
    Frm49.MSFlexGrid1.TextMatrix(x, 0) = x
    Frm49.MSFlexGrid1.TextMatrix(x, 1) = x
    If Not IsNull(rs!Nama) Then Frm49.MSFlexGrid1.TextMatrix(x, 2) = rs!Nama 'Nama Pekerja
    If Not IsNull(rs!NoIC) Then Frm49.MSFlexGrid1.TextMatrix(x, 3) = rs!NoIC 'No. Kad Pengenalan
    If Not IsNull(rs!NoPekerja) Then Frm49.MSFlexGrid1.TextMatrix(x, 4) = rs!NoPekerja 'No Pekerja
    If Not IsNull(rs!NoKWSP) Then Frm49.MSFlexGrid1.TextMatrix(x, 5) = rs!NoKWSP 'No. KWSP
    If Not IsNull(rs!NoSocso) Then Frm49.MSFlexGrid1.TextMatrix(x, 6) = rs!NoSocso 'No. Socso
    If Not IsNull(rs!NoTel) Then Frm49.MSFlexGrid1.TextMatrix(x, 7) = rs!NoTel 'No. Tel
    If Not IsNull(rs!Jawatan) Then Frm49.MSFlexGrid1.TextMatrix(x, 8) = rs!Jawatan 'Jawatan
    If Not IsNull(rs!Samaran) Then Frm49.MSFlexGrid1.TextMatrix(x, 9) = rs!Samaran 'Nama Samaran
    If Not IsNull(rs!Password) Then Frm49.MSFlexGrid1.TextMatrix(x, 10) = rs!Password 'Password
    If Not IsNull(rs!user_level) Then
        If rs!user_level = 1 Then
            Frm49.MSFlexGrid1.TextMatrix(x, 11) = "Admin" 'User Level
        ElseIf rs!user_level = 2 Then
            Frm49.MSFlexGrid1.TextMatrix(x, 11) = "Manager" 'User Level
        ElseIf rs!user_level = 3 Then
            Frm49.MSFlexGrid1.TextMatrix(x, 11) = "Staff" 'User Level
        End If
    Else
        Frm49.MSFlexGrid1.TextMatrix(x, 11) = "Staff" 'User Level
    End If
    If Not IsNull(rs!Gaji) Then Frm49.MSFlexGrid1.TextMatrix(x, 12) = rs!Gaji 'Gaji
    If Not IsNull(rs!Elaun) Then Frm49.MSFlexGrid1.TextMatrix(x, 13) = rs!Elaun 'Elaun
    If Not IsNull(rs!komisen) Then
        If rs!komisen = 0 Then
            Frm49.MSFlexGrid1.TextMatrix(x, 14) = "Tidak"
        Else
            Frm49.MSFlexGrid1.TextMatrix(x, 14) = "Ya"
        End If
    End If
    If Not IsNull(rs!Status) Then
        Frm49.MSFlexGrid1.TextMatrix(x, 15) = rs!Status 'Status
    Else
        Frm49.MSFlexGrid1.TextMatrix(x, 15) = "Aktif" 'Status
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
End Sub


