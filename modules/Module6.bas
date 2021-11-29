Attribute VB_Name = "Module6"
Sub Frm113_initial_setting()
'on error resume next
Frm113.L4_Text = vbNullString
Frm113.L9_Text = vbNullString
Frm113.L5_Text = 0
Frm113.L6_Text = 0

Frm113.L10_Text = 0
Frm113.L11_Text = 0
Frm113.L12_Text = 0
Frm113.L13_Text = 0
End Sub
Sub Frm113_initial()
'on error resume next
Frm113.TB1 = vbNullString
Frm113.TB2 = vbNullString
Frm113.L17_Text = vbNullString
Frm113.L18_Text.Visible = False

Frm113.L16_Text = 75

Frm113.CB1 = 1
Frm113.CB2 = 0

Frm113.CBB1.Clear

'###Senarai Nama Pekerja###
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where status='" & "Aktif" & "' AND InvestorSmall = 0 AND InvestorBig = 0", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    If rs!user_level = 1 Or rs!user_level = 2 Or rs!user_level = 3 Or rs!user_level = G_LEVEL_USER Then Frm113.CBB1.AddItem rs!Samaran & "  |  " & rs!NoPekerja
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
End Sub
Sub Frm113_initial2()
'on error resume next
Frm113.Pic1.Left = 120
Frm113.Pic1.Top = 960
Frm113.Pic2.Left = 120
Frm113.Pic2.Top = 960

Frm113.Pic1.Visible = False
Frm113.Pic2.Visible = False
End Sub
Sub Frm113_senarai_mata_ganjaran_header()
'on error resume next
'#### Header Report Servis #### - Start
Frm113.MSFlexGrid1.Clear
Frm113.MSFlexGrid1.RowHeight(0) = 900
Frm113.MSFlexGrid1.FormatString = "No.|<No.|<ID|<Jenis|<Tarikh|<No. Invoice|<Harga Layak Mata (RM)|<Kadar Perolehan Mata|<Jumlah Perolehan Mata|<Jumlah Tebus Mata|<Kadar Tebus Mata|<Nilaian Tebusan Mata (RM)|<Remarks"

Frm113.MSFlexGrid1.Rows = 1
Frm113.MSFlexGrid1.ColWidth(0) = 0 'No.
Frm113.MSFlexGrid1.ColWidth(1) = 600 'No.
Frm113.MSFlexGrid1.ColWidth(2) = 0 'ID
Frm113.MSFlexGrid1.ColWidth(3) = 1300 'Jenis
Frm113.MSFlexGrid1.ColWidth(4) = 1000 'Tarikh
Frm113.MSFlexGrid1.ColWidth(5) = 1200 'No. Invoice
Frm113.MSFlexGrid1.ColWidth(6) = 1000 'Harga Layak Mata (RM)
Frm113.MSFlexGrid1.ColWidth(7) = 830 'Kadar Perolehan Mata
Frm113.MSFlexGrid1.ColWidth(8) = 1000 'Jumlah Perolehan Mata
Frm113.MSFlexGrid1.ColWidth(9) = 1000 'Jumlah Tebus Mata
Frm113.MSFlexGrid1.ColWidth(10) = 700 'Kadar Tebus Mata
Frm113.MSFlexGrid1.ColWidth(11) = 1000 'Nilaian Tebusan Mata
Frm113.MSFlexGrid1.ColWidth(12) = 7000 'Remarks
'#### Header Report Servis #### - End
End Sub
Sub Frm113_senarai_mata_ganjaran()
'on error resume next
Dim Frm113_LM_TOTAL_PAGE As Double
Dim Frm113_LM_PEROLEH As Double
Dim Frm113_LM_TEBUS As Double

Frm113_PAGE_SIZE = 36
Frm113_LM_TOTAL_PAGE = 0
Frm113_LM_PEROLEH = 0
Frm113_LM_TEBUS = 0
x = 0

LM_START_ROW = Frm113.L7_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm113_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm113.L8_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm113_PAGE_SIZE
        End If
    End If
End If

Frm113_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 71_tebus_agih_point where no_ahli='" & Frm113.L9_Text & "' AND status = 1 order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm113_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm113_LM_PAGE_FOUND = 0 Then
        If Frm113.L8_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm113.L5_Text = Frm113.L5_Text + 1 'Paparan Page ke-xxx
                Frm113_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm113.L5_Text) Then
                    If Frm113.L5_Text <> 1 Then
                        Frm113.L5_Text = Frm113.L5_Text - 1 'Paparan Page ke-xxx
                        Frm113_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm113.L5_Text - 1) * Frm113_PAGE_SIZE) + x
    Frm113.MSFlexGrid1.Rows = x + 1
    Frm113.MSFlexGrid1.TextMatrix(x, 0) = x 'No.
    Frm113.MSFlexGrid1.TextMatrix(x, 1) = Y 'No.
    Frm113.MSFlexGrid1.ColAlignment(1) = 4
    
    Frm113.MSFlexGrid1.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!Type) Then
        If rs!Type = 1 Then
            Frm113.MSFlexGrid1.TextMatrix(x, 3) = "Belian"
        ElseIf rs!Type = 2 Then
            Frm113.MSFlexGrid1.TextMatrix(x, 3) = "Pemberian Mata"
        ElseIf rs!Type = 3 Then
            Frm113.MSFlexGrid1.TextMatrix(x, 3) = "Potongan Mata"
        End If
    End If
    
    If Not IsNull(rs!tarikh) Then Frm113.MSFlexGrid1.TextMatrix(x, 4) = rs!tarikh 'Tarikh
    Frm113.MSFlexGrid1.ColAlignment(4) = 4
    
    If Not IsNull(rs!no_invoice) Then Frm113.MSFlexGrid1.TextMatrix(x, 5) = rs!no_invoice
    Frm113.MSFlexGrid1.ColAlignment(5) = 4
    
    If Not IsNull(rs!harga_layak_bonus) Then 'Harga Layak Mata (RM)
        Frm113.MSFlexGrid1.TextMatrix(x, 6) = Format(rs!harga_layak_bonus, "#,##0.00")
    Else
        Frm113.MSFlexGrid1.TextMatrix(x, 6) = "0.00"
    End If
    Frm113.MSFlexGrid1.ColAlignment(6) = 4
    
    If Not IsNull(rs!kadar_peroleh_point) Then Frm113.MSFlexGrid1.TextMatrix(x, 7) = rs!kadar_peroleh_point 'Kadar Perolehan Mata
    Frm113.MSFlexGrid1.ColAlignment(7) = 4
    
    If Not IsNull(rs!jumlah_peroleh_point) Then Frm113.MSFlexGrid1.TextMatrix(x, 8) = rs!jumlah_peroleh_point 'Jumlah Perolehan Mata
    Frm113.MSFlexGrid1.ColAlignment(8) = 4
    
    If Not IsNull(rs!jumlah_tebus_point) Then Frm113.MSFlexGrid1.TextMatrix(x, 9) = rs!jumlah_tebus_point 'Jumlah Tebus Mata
    Frm113.MSFlexGrid1.ColAlignment(9) = 4

    If Not IsNull(rs!kadar_tebus_point) Then Frm113.MSFlexGrid1.TextMatrix(x, 10) = rs!kadar_tebus_point 'Kadar Tebus Mata
    Frm113.MSFlexGrid1.ColAlignment(10) = 4

    If Not IsNull(rs!nilaian_tebus_point) Then 'Nilaian Tebusan Mata
        Frm113.MSFlexGrid1.TextMatrix(x, 11) = Format(rs!nilaian_tebus_point, "#,##0.00")
    Else
        Frm113.MSFlexGrid1.TextMatrix(x, 11) = "0.00"
    End If
    Frm113.MSFlexGrid1.ColAlignment(11) = 4

    If Not IsNull(rs!remarks) Then Frm113.MSFlexGrid1.TextMatrix(x, 12) = rs!remarks 'Remarks

    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 71_tebus_agih_point where no_ahli='" & Frm113.L9_Text & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm113_LM_TOTAL_PAGE = Format(rs(0) / Frm113_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm113_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm113_LM_PAGE = Split(Frm113_LM_TOTAL_PAGE, ".")(0)
        Frm113_LM_PAGE_LEBIHAN = Split(Frm113_LM_TOTAL_PAGE, ".")(1)
        
        If Frm113_LM_PAGE_LEBIHAN <> "00" Then
            Frm113.L6_Text = Frm113_LM_PAGE + 1
        Else
            Frm113.L6_Text = Frm113_LM_PAGE
        End If
        
    Else
    
        Frm113.L6_Text = Frm113_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm113.L6_Text = 0
    End If
Else
    Frm113.L6_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm113.L6_Text = vbNullString Then
    Frm113.L6_Text = 0
End If
'### Jumlah Data ### - End

'### Bilangan Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 71_tebus_agih_point where no_ahli='" & Frm113.L9_Text & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm113.L13_Text = rs(0) 'Bilangan data

rs.Close
Set rs = Nothing
'### Bilangan Data ### - End

'### Jumlah Perolehan Mata ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_peroleh_point) from 71_tebus_agih_point where no_ahli='" & Frm113.L9_Text & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm113_LM_PEROLEH = rs(0) 'Jumlah Perolehan Mata

rs.Close
Set rs = Nothing
'### Jumlah Perolehan Mata ### - End

'### Jumlah Tebusan Mata ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_tebus_point) from 71_tebus_agih_point where no_ahli='" & Frm113.L9_Text & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm113_LM_TEBUS = rs(0) 'Jumlah Tebusan Mata

rs.Close
Set rs = Nothing
'### Jumlah Tebusan Mata ### - End

Frm113.L10_Text = Frm113_LM_PEROLEH
Frm113.L11_Text = Frm113_LM_TEBUS
Frm113.L12_Text = Frm113_LM_PEROLEH - Frm113_LM_TEBUS

If x <> 0 Then
    Frm113.L8_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm113.L7_Text = LM_START_ROW 'Titik Pencarian Data
    
    'Frm113.Pic1.Visible = False
    'Frm113.Pic2.Visible = True
Else
    Frm113.L8_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If

End Sub
Sub frm134_report_stok_dulang()
'on error resume next
Dim rs1 As ADODB.Recordset

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "UPDATE setting_database set stok_qty = 0 , stok_qty_ti = 0 , total_stok = 0 where status = 1"

Set rs = cn.Execute(strsql)
Set rs = Nothing
                    
If MDI_frm1.L20_Text = "Semua cawangan" Then

    Frm85_LM_SEARCH_5 = Null
    Frm85_LM_SEARCH_5_LOGIC = "<>"
    Frm85_LM_SEARCH_6 = Null
    Frm85_LM_SEARCH_6_LOGIC = "<>"
    
Else

    Frm85_LM_SEARCH_5 = MDI_frm1.L20_Text
    Frm85_LM_SEARCH_5_LOGIC = "="
    Frm85_LM_SEARCH_6 = "HQ"
    Frm85_LM_SEARCH_6_LOGIC = "="
    
End If

user_level = MDI_frm1.L4_Text

If user_level = "Administration" Or user_level = "Guest/User" Then
    Frm85_LM_SEARCH_10 = 1
    Frm85_LM_SEARCH_10_LOGIC = "="
    
    Frm85_LM_SEARCH_11 = 1
    Frm85_LM_SEARCH_11_LOGIC = "="
Else
    Frm85_LM_SEARCH_10 = 1
    Frm85_LM_SEARCH_10_LOGIC = "="
    
    Frm85_LM_SEARCH_11 = 0
    Frm85_LM_SEARCH_11_LOGIC = "="
End If

Dim LM_STOK_BARU As Double

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where status='" & 1 & "' AND SenaraiDulang is not null", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    LM_STOK_BARU = 0
    
    Set rs1 = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs1.Open "select COUNT(ID) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "') AND dulang='" & rs!SenaraiDulang & "' AND (receiving_Status='" & 0 & "' OR receiving_Status='" & 1 & "' OR receiving_Status='" & 4 & "')", cn, adOpenKeyset, adLockOptimistic
    
    If Not IsNull(rs1(0)) Then
    
        rs!stok_qty = rs1(0)
        LM_STOK_BARU = rs1(0)
        rs.Update
        
    End If
    
    rs1.Close
    Set rs1 = Nothing
    
    Set rs1 = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs1.Open "select COUNT(ID) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "') AND dulang='" & rs!SenaraiDulang & "' AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "')", cn, adOpenKeyset, adLockOptimistic
    
    If Not IsNull(rs1(0)) Then
    
        rs!stok_qty_ti = rs1(0)
        rs!total_stok = rs1(0) + LM_STOK_BARU
        
        rs.Update
        
    End If
    
    rs1.Close
    Set rs1 = Nothing

    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

frm134.L1 = "Report dikeluarkan pada " & Now
End Sub
Sub frm134_report_stok_header()
'on error resume next
With frm134.LV1
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear

    frm134.LV1.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 1000, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "Dulang", 2200, 2
    .ColumnHeaders.Add 5, , "Stok (Baru)", 3000, 2
    .ColumnHeaders.Add 6, , "Stok (Trade In)", 3500, 2
    .ColumnHeaders.Add 7, , "Jumlah", 3000, 2
    
End With
End Sub
Sub frm134_report_stok()
'On Error Resume Next
Dim frm134_LM_TOTAL_PAGE As Double

frm134_PAGE_SIZE = 19
frm134_LM_TOTAL_PAGE = 0
x = 0

re_gen_report:

LM_START_ROW = frm134.L69_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + frm134_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If frm134.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - frm134_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        frm134.L67_Text = 1
    End If
End If

frm134_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where status = 1 AND SenaraiDulang is not NULL order by SenaraiDulang ASC LIMIT " & LM_START_ROW & "," & frm134_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1
    If frm134_LM_PAGE_FOUND = 0 Then
        If frm134.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                frm134.L67_Text = frm134.L67_Text + 1 'Paparan Page ke-xxx
                frm134_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(frm134.L67_Text) Then
                    If frm134.L67_Text <> 1 Then
                        frm134.L67_Text = frm134.L67_Text - 1 'Paparan Page ke-xxx
                        frm134_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If

    Y = ((frm134.L67_Text - 1) * frm134_PAGE_SIZE) + x

    With frm134.LV1.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!SenaraiDulang) Then
            .ListSubItems.Add , , rs!SenaraiDulang
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!stok_qty) Then
            .ListSubItems.Add , , Format(rs!stok_qty, "#,##0")
        Else
            .ListSubItems.Add , , "0"
        End If
        
        If Not IsNull(rs!stok_qty_ti) Then
            .ListSubItems.Add , , Format(rs!stok_qty_ti, "#,##0")
        Else
            .ListSubItems.Add , , "0"
        End If
        
        If Not IsNull(rs!total_stok) Then
            .ListSubItems.Add , , Format(rs!total_stok, "#,##0")
        Else
            .ListSubItems.Add , , "0"
        End If

    End With
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from setting_database where status = 1 AND SenaraiDulang is not NULL", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    frm134_LM_TOTAL_PAGE = Format(rs(0) / frm134_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, frm134_LM_TOTAL_PAGE, ".") <> 0 Then
    
        frm134_LM_PAGE = Split(frm134_LM_TOTAL_PAGE, ".")(0)
        frm134_LM_PAGE_LEBIHAN = Split(frm134_LM_TOTAL_PAGE, ".")(1)
        
        If frm134_LM_PAGE_LEBIHAN <> "00" Then
            frm134.L68_Text = frm134_LM_PAGE + 1
        Else
            frm134.L68_Text = frm134_LM_PAGE
        End If
        
    Else
    
        frm134.L68_Text = frm134_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        frm134.L68_Text = 0
    End If
Else
    frm134.L68_Text = 0
End If

rs.Close
Set rs = Nothing

If x <> 0 Then
    frm134.L69_Text = LM_START_ROW
End If

If frm134.L67_Text <> vbNullString And IsNumeric(frm134.L67_Text) Then
    If frm134.L68_Text <> vbNullString And IsNumeric(frm134.L68_Text) Then
        frm134_LM_CURR_PAGE = frm134.L67_Text
        frm134_LM_TOTAL_PAGE = frm134.L68_Text
        
        If frm134_LM_CURR_PAGE > frm134_LM_TOTAL_PAGE Then
            
            frm134.L67_Text = frm134.L67_Text - 1
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            GoTo re_gen_report:
            
        End If
    End If
End If
End Sub
Sub frm134_cetak_penyata_dulang()
'on error resume next
'### Reset maklumat kedai ### - Start
Report82.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report82.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report82.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report82.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report82.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

Report82.Sections("Section5").Controls("L1").Caption = frm134.L1

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_CAWANGAN & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report82.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report82.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report82.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report82.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report82.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End
   
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where status = 1 AND SenaraiDulang is not NULL order by SenaraiDulang ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report82.DataSource = rs
    If G_PREVIEW = 1 Then Report82.Show
    
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing

If G_PREVIEW = 0 Then Report82.PrintReport
End Sub
