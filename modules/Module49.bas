Attribute VB_Name = "Module49"
Sub frm129_initial_setting()
'on error resume next
frm129.CB1 = 0

frm129.CB2 = 1
frm129.CB3 = 0
frm129.CB4 = 0

frm129.DTPicker1 = DateTime.Date
frm129.DTPicker2 = DateTime.Date

frm129.CBB1.Clear
frm129.CBB2.Clear

frm129.CBB1.AddItem "Semua kategori"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where Kategori_Produk<>'" & Null & "' AND status = 1 order by Metal_Purity ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!kategori_Produk) Then frm129.CBB1.AddItem rs!kategori_Produk
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

frm129.CBB1 = "Semua kategori"

frm129.CBB2.AddItem "semua purity"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where Kod_Metal_Purity<>'" & Null & "' AND status = 1 order by Metal_Purity ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!Kod_Metal_Purity) Then frm129.CBB2.AddItem rs!Kod_Metal_Purity
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

frm129.CBB2 = "semua purity"

frm129.L69_Text = -1 'Titik Pencarian Data
frm129.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
frm129.L67_Text = 0 'Paparan Page ke-xxx
frm129.L68_Text = 0

frm129.L5_Text = vbNullString
frm129.L6_Text = vbNullString
frm129.L7_Text = vbNullString
frm129.L8_Text = vbNullString
frm129.L9_Text = vbNullString
frm129.L15_Text = vbNullString

frm129.L10_Text = 0
frm129.L11_Text = Format(0, "#,##0.00 g")
End Sub
Sub frm129_report_trade_in_header()
'on error resume next
frm129.MSFlexGrid1.Clear
frm129.MSFlexGrid1.RowHeight(0) = 700
frm129.MSFlexGrid1.FormatString = "No.|<No.|<ID|<No. Siri Produk|<Purity|<Nama Produk|<Berat (g)"

'No.
'No. Siri Produk
'Purity
'Nama Produk
'Berat Asal (g)
'Beza Berat (g)

frm129.MSFlexGrid1.Rows = 1
frm129.MSFlexGrid1.ColWidth(0) = 0 'No.
frm129.MSFlexGrid1.ColAlignment(0) = 4

frm129.MSFlexGrid1.ColWidth(1) = 800 'No.
frm129.MSFlexGrid1.ColAlignment(1) = 4

frm129.MSFlexGrid1.ColWidth(2) = 0 'No. ID

frm129.MSFlexGrid1.ColWidth(3) = 1500 'No. Siri Produk

frm129.MSFlexGrid1.ColWidth(4) = 1700 'Purity

frm129.MSFlexGrid1.ColWidth(5) = 9500 'Nama Produk

frm129.MSFlexGrid1.ColWidth(6) = 1500 'Berat Asal (g)
frm129.MSFlexGrid1.ColAlignment(6) = 7
End Sub
Sub frm129_report_trade_in_belian()
'on error resume next
Dim frm129_LM_TOTAL_PAGE As Double

frm129_PAGE_SIZE = 32
frm129_LM_TOTAL_PAGE = 0
x = 0

re_gen_report:

frm129.L10_Text = "0"
frm129.L11_Text = "0.00 g"

If frm129.L5_Text = "1" Then '0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh

    TM = frm129.L6_Text 'Tarikh mula
    TA = frm129.L7_Text 'Tarikh akhir

End If

If frm129.L8_Text = "Semua kategori" Then
    
    frm129_LM_SEARCH_1 = Null
    frm129_LM_SEARCH_1_LOGIC = "<>"
    
Else

    frm129_LM_SEARCH_1 = frm129.L8_Text
    frm129_LM_SEARCH_1_LOGIC = "="
    
End If

If frm129.L9_Text = "semua purity" Then
    
    frm129_LM_SEARCH_2 = Null
    frm129_LM_SEARCH_2_LOGIC = "<>"
    
Else
    
    frm129_LM_SEARCH_2 = frm129.L9_Text
    frm129_LM_SEARCH_2_LOGIC = "="

End If

If frm129.L5_Text = 0 Then frm129.L14_Text = "Rekod belian trade in bagi [" & frm129.L8_Text & "] dan purity [" & frm129.L9_Text & "]."
If frm129.L5_Text = 1 Then frm129.L14_Text = "Rekod belian trade in bagi [" & frm129.L8_Text & "] dan purity [" & frm129.L9_Text & "] dari " & TM & " hingga " & TA & "."

LM_START_ROW = frm129.L69_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + frm129_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If frm129.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - frm129_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        frm129.L67_Text = 1
    End If
End If

frm129_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm129.L5_Text = 0 Then rs.Open "select * from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND StatusItem <> 0 AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "') order by tarikh_belian ASC , ID ASC LIMIT " & LM_START_ROW & "," & frm129_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If frm129.L5_Text = 1 Then rs.Open "select * from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "') order by tarikh_belian ASC , ID ASC LIMIT " & LM_START_ROW & "," & frm129_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1
    If frm129_LM_PAGE_FOUND = 0 Then
        If frm129.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                frm129.L67_Text = frm129.L67_Text + 1 'Paparan Page ke-xxx
                frm129_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(frm129.L67_Text) Then
                    If frm129.L67_Text <> 1 Then
                        frm129.L67_Text = frm129.L67_Text - 1 'Paparan Page ke-xxx
                        frm129_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If

    Y = ((frm129.L67_Text - 1) * frm129_PAGE_SIZE) + x
    frm129.MSFlexGrid1.Rows = x + 1
    frm129.MSFlexGrid1.TextMatrix(x, 0) = x 'No.
    frm129.MSFlexGrid1.TextMatrix(x, 1) = Y 'No.
    frm129.MSFlexGrid1.ColAlignment(1) = 4
    frm129.MSFlexGrid1.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!no_siri_Produk) Then frm129.MSFlexGrid1.TextMatrix(x, 3) = rs!no_siri_Produk 'No. Siri Produk
    If Not IsNull(rs!kod_Purity) Then frm129.MSFlexGrid1.TextMatrix(x, 4) = rs!kod_Purity 'Purity
    If Not IsNull(rs!kategori_Produk) Then frm129.MSFlexGrid1.TextMatrix(x, 5) = rs!kategori_Produk 'Nama Produk
    If Not IsNull(rs!Berat) Then frm129.MSFlexGrid1.TextMatrix(x, 6) = Format(rs!Berat, "#,##0.00") 'Berat (g)
    'If Not IsNull(rs!Beza_Berat) Then frm129.MSFlexGrid1.TextMatrix(x, 7) = Format(rs!Beza_Berat, "#,##0.00") 'Beza Berat (g)

    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm129.L5_Text = 0 Then rs.Open "select COUNT(ID) from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND StatusItem <> 0 AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "')", cn, adOpenKeyset, adLockOptimistic
If frm129.L5_Text = 1 Then rs.Open "select COUNT(ID) from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "')", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    frm129_LM_TOTAL_PAGE = Format(rs(0) / frm129_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, frm129_LM_TOTAL_PAGE, ".") <> 0 Then
    
        frm129_LM_PAGE = Split(frm129_LM_TOTAL_PAGE, ".")(0)
        frm129_LM_PAGE_LEBIHAN = Split(frm129_LM_TOTAL_PAGE, ".")(1)
        
        If frm129_LM_PAGE_LEBIHAN <> "00" Then
            frm129.L68_Text = frm129_LM_PAGE + 1
        Else
            frm129.L68_Text = frm129_LM_PAGE
        End If
        
    Else
    
        frm129.L68_Text = frm129_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        frm129.L68_Text = 0
    End If
Else
    frm129.L68_Text = 0
End If

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm129.L5_Text = 0 Then rs.Open "select COUNT(ID) , SUM(berat) from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND StatusItem <> 0 AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "')", cn, adOpenKeyset, adLockOptimistic
If frm129.L5_Text = 1 Then rs.Open "select COUNT(ID) , SUM(berat) from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "')", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then frm129.L10_Text = rs(0)
If Not IsNull(rs(1)) Then frm129.L11_Text = Format(rs(1), "#,##0.00 g")

rs.Close
Set rs = Nothing

If x <> 0 Then
    frm129.L69_Text = LM_START_ROW
End If

If frm129.L67_Text <> vbNullString And IsNumeric(frm129.L67_Text) Then
    If frm129.L68_Text <> vbNullString And IsNumeric(frm129.L68_Text) Then
        frm129_LM_CURR_PAGE = frm129.L67_Text
        frm129_LM_TOTAL_PAGE = frm129.L68_Text
        
        If frm129_LM_CURR_PAGE > frm129_LM_TOTAL_PAGE Then
            
            frm129.L67_Text = frm129.L67_Text - 1
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            GoTo re_gen_report:
            
        End If
    End If
End If
End Sub
Sub frm129_report_trade_in_jualan()
'on error resume next
Dim frm129_LM_TOTAL_PAGE As Double

frm129_PAGE_SIZE = 32
frm129_LM_TOTAL_PAGE = 0
x = 0

re_gen_report:

frm129.L10_Text = "0"
frm129.L11_Text = "0.00 g"

If frm129.L5_Text = "1" Then '0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh

    TM = frm129.L6_Text 'Tarikh mula
    TA = frm129.L7_Text 'Tarikh akhir

End If

If frm129.L8_Text = "Semua kategori" Then
    
    frm129_LM_SEARCH_1 = Null
    frm129_LM_SEARCH_1_LOGIC = "<>"
    
Else

    frm129_LM_SEARCH_1 = frm129.L8_Text
    frm129_LM_SEARCH_1_LOGIC = "="
    
End If

If frm129.L9_Text = "semua purity" Then
    
    frm129_LM_SEARCH_2 = Null
    frm129_LM_SEARCH_2_LOGIC = "<>"
    
Else
    
    frm129_LM_SEARCH_2 = frm129.L9_Text
    frm129_LM_SEARCH_2_LOGIC = "="

End If

If frm129.L5_Text = 0 Then frm129.L14_Text = "Rekod jualan trade in bagi [" & frm129.L8_Text & "] dan purity [" & frm129.L9_Text & "]."
If frm129.L5_Text = 1 Then frm129.L14_Text = "Rekod jualan trade in bagi [" & frm129.L8_Text & "] dan purity [" & frm129.L9_Text & "] dari " & TM & " hingga " & TA & "."

LM_START_ROW = frm129.L69_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + frm129_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If frm129.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - frm129_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        frm129.L67_Text = 1
    End If
End If

frm129_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm129.L5_Text = 0 Then rs.Open "select * from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND (StatusItem = 11 OR StatusItem = 12 OR StatusItem = 21 OR StatusItem = 22 OR StatusItem = 27 OR StatusItem = 28) AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "') order by tarikh_belian ASC , ID ASC LIMIT " & LM_START_ROW & "," & frm129_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If frm129.L5_Text = 1 Then rs.Open "select * from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND (StatusItem = 11 OR StatusItem = 12 OR StatusItem = 21 OR StatusItem = 22 OR StatusItem = 27 OR StatusItem = 28) AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "') order by tarikh_belian ASC , ID ASC LIMIT " & LM_START_ROW & "," & frm129_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1
    If frm129_LM_PAGE_FOUND = 0 Then
        If frm129.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                frm129.L67_Text = frm129.L67_Text + 1 'Paparan Page ke-xxx
                frm129_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(frm129.L67_Text) Then
                    If frm129.L67_Text <> 1 Then
                        frm129.L67_Text = frm129.L67_Text - 1 'Paparan Page ke-xxx
                        frm129_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If

    Y = ((frm129.L67_Text - 1) * frm129_PAGE_SIZE) + x
    frm129.MSFlexGrid1.Rows = x + 1
    frm129.MSFlexGrid1.TextMatrix(x, 0) = x 'No.
    frm129.MSFlexGrid1.TextMatrix(x, 1) = Y 'No.
    frm129.MSFlexGrid1.ColAlignment(1) = 4
    frm129.MSFlexGrid1.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!no_siri_Produk) Then frm129.MSFlexGrid1.TextMatrix(x, 3) = rs!no_siri_Produk 'No. Siri Produk
    If Not IsNull(rs!kod_Purity) Then frm129.MSFlexGrid1.TextMatrix(x, 4) = rs!kod_Purity 'Purity
    If Not IsNull(rs!kategori_Produk) Then frm129.MSFlexGrid1.TextMatrix(x, 5) = rs!kategori_Produk 'Nama Produk
    If Not IsNull(rs!Berat) Then frm129.MSFlexGrid1.TextMatrix(x, 6) = Format(rs!Berat - rs!beza_berat, "#,##0.00") 'Berat (g)
    'If Not IsNull(rs!Beza_Berat) Then frm129.MSFlexGrid1.TextMatrix(x, 7) = Format(rs!Beza_Berat, "#,##0.00") 'Beza Berat (g)

    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm129.L5_Text = 0 Then rs.Open "select COUNT(ID) from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND (StatusItem = 11 OR StatusItem = 12 OR StatusItem = 21 OR StatusItem = 22 OR StatusItem = 27 OR StatusItem = 28) AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "')", cn, adOpenKeyset, adLockOptimistic
If frm129.L5_Text = 1 Then rs.Open "select COUNT(ID) from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND (StatusItem = 11 OR StatusItem = 12 OR StatusItem = 21 OR StatusItem = 22 OR StatusItem = 27 OR StatusItem = 28) AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "')", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    frm129_LM_TOTAL_PAGE = Format(rs(0) / frm129_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, frm129_LM_TOTAL_PAGE, ".") <> 0 Then
    
        frm129_LM_PAGE = Split(frm129_LM_TOTAL_PAGE, ".")(0)
        frm129_LM_PAGE_LEBIHAN = Split(frm129_LM_TOTAL_PAGE, ".")(1)
        
        If frm129_LM_PAGE_LEBIHAN <> "00" Then
            frm129.L68_Text = frm129_LM_PAGE + 1
        Else
            frm129.L68_Text = frm129_LM_PAGE
        End If
        
    Else
    
        frm129.L68_Text = frm129_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        frm129.L68_Text = 0
    End If
Else
    frm129.L68_Text = 0
End If

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm129.L5_Text = 0 Then rs.Open "select COUNT(ID) , SUM(Berat - Beza_Berat) from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND (StatusItem = 11 OR StatusItem = 12 OR StatusItem = 21 OR StatusItem = 22 OR StatusItem = 27 OR StatusItem = 28) AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "')", cn, adOpenKeyset, adLockOptimistic
If frm129.L5_Text = 1 Then rs.Open "select COUNT(ID) , SUM(Berat - Beza_Berat) from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND (StatusItem = 11 OR StatusItem = 12 OR StatusItem = 21 OR StatusItem = 22 OR StatusItem = 27 OR StatusItem = 28) AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "')", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then frm129.L10_Text = rs(0)
If Not IsNull(rs(1)) Then frm129.L11_Text = Format(rs(1), "#,##0.00 g")

rs.Close
Set rs = Nothing

If x <> 0 Then
    frm129.L69_Text = LM_START_ROW
End If

If frm129.L67_Text <> vbNullString And IsNumeric(frm129.L67_Text) Then
    If frm129.L68_Text <> vbNullString And IsNumeric(frm129.L68_Text) Then
        frm129_LM_CURR_PAGE = frm129.L67_Text
        frm129_LM_TOTAL_PAGE = frm129.L68_Text
        
        If frm129_LM_CURR_PAGE > frm129_LM_TOTAL_PAGE Then
            
            frm129.L67_Text = frm129.L67_Text - 1
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            GoTo re_gen_report:
            
        End If
    End If
End If
End Sub
Sub frm129_report_trade_in_stok()
'on error resume next
Dim frm129_LM_TOTAL_PAGE As Double

frm129_PAGE_SIZE = 32
frm129_LM_TOTAL_PAGE = 0
x = 0

re_gen_report:

frm129.L10_Text = "0"
frm129.L11_Text = "0.00 g"

If frm129.L5_Text = "1" Then '0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh

    TM = frm129.L6_Text 'Tarikh mula
    TA = frm129.L7_Text 'Tarikh akhir

End If

If frm129.L8_Text = "Semua kategori" Then
    
    frm129_LM_SEARCH_1 = Null
    frm129_LM_SEARCH_1_LOGIC = "<>"
    
Else

    frm129_LM_SEARCH_1 = frm129.L8_Text
    frm129_LM_SEARCH_1_LOGIC = "="
    
End If

If frm129.L9_Text = "semua purity" Then
    
    frm129_LM_SEARCH_2 = Null
    frm129_LM_SEARCH_2_LOGIC = "<>"
    
Else
    
    frm129_LM_SEARCH_2 = frm129.L9_Text
    frm129_LM_SEARCH_2_LOGIC = "="

End If

If frm129.L5_Text = 0 Then frm129.L14_Text = "Rekod stok trade in bagi [" & frm129.L8_Text & "] dan purity [" & frm129.L9_Text & "]."
If frm129.L5_Text = 1 Then frm129.L14_Text = "Rekod stok trade in bagi [" & frm129.L8_Text & "] dan purity [" & frm129.L9_Text & "] dari " & TM & " hingga " & TA & "."

LM_START_ROW = frm129.L69_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + frm129_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If frm129.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - frm129_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        frm129.L67_Text = 1
    End If
End If

frm129_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm129.L5_Text = 0 Then rs.Open "select * from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND (StatusItem = 10 OR StatusItem = 12 OR StatusItem = 22 OR StatusItem = 28) AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "') order by tarikh_belian ASC , ID ASC LIMIT " & LM_START_ROW & "," & frm129_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If frm129.L5_Text = 1 Then rs.Open "select * from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND (StatusItem = 10 OR StatusItem = 12 OR StatusItem = 22 OR StatusItem = 28) AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "') order by tarikh_belian ASC , ID ASC LIMIT " & LM_START_ROW & "," & frm129_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1
    If frm129_LM_PAGE_FOUND = 0 Then
        If frm129.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                frm129.L67_Text = frm129.L67_Text + 1 'Paparan Page ke-xxx
                frm129_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(frm129.L67_Text) Then
                    If frm129.L67_Text <> 1 Then
                        frm129.L67_Text = frm129.L67_Text - 1 'Paparan Page ke-xxx
                        frm129_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If

    Y = ((frm129.L67_Text - 1) * frm129_PAGE_SIZE) + x
    frm129.MSFlexGrid1.Rows = x + 1
    frm129.MSFlexGrid1.TextMatrix(x, 0) = x 'No.
    frm129.MSFlexGrid1.TextMatrix(x, 1) = Y 'No.
    frm129.MSFlexGrid1.ColAlignment(1) = 4
    frm129.MSFlexGrid1.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!no_siri_Produk) Then frm129.MSFlexGrid1.TextMatrix(x, 3) = rs!no_siri_Produk 'No. Siri Produk
    If Not IsNull(rs!kod_Purity) Then frm129.MSFlexGrid1.TextMatrix(x, 4) = rs!kod_Purity 'Purity
    If Not IsNull(rs!kategori_Produk) Then frm129.MSFlexGrid1.TextMatrix(x, 5) = rs!kategori_Produk 'Nama Produk
    
    If rs!StatusItem = "10" Then
        If Not IsNull(rs!Berat) Then frm129.MSFlexGrid1.TextMatrix(x, 6) = Format(rs!Berat, "#,##0.00") 'Berat (g)
    Else
        If Not IsNull(rs!beza_berat) Then frm129.MSFlexGrid1.TextMatrix(x, 6) = Format(rs!beza_berat, "#,##0.00") 'Berat (g)
    End If
    'If Not IsNull(rs!Beza_Berat) Then frm129.MSFlexGrid1.TextMatrix(x, 7) = Format(rs!Beza_Berat, "#,##0.00") 'Beza Berat (g)

    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm129.L5_Text = 0 Then rs.Open "select COUNT(ID) from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND (StatusItem = 10 OR StatusItem = 12 OR StatusItem = 22 OR StatusItem = 28) AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "')", cn, adOpenKeyset, adLockOptimistic
If frm129.L5_Text = 1 Then rs.Open "select COUNT(ID) from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND (StatusItem = 10 OR StatusItem = 12 OR StatusItem = 22 OR StatusItem = 28) AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "')", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then frm129.L10_Text = rs(0)

If Not rs.EOF Then

    frm129_LM_TOTAL_PAGE = Format(rs(0) / frm129_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, frm129_LM_TOTAL_PAGE, ".") <> 0 Then
    
        frm129_LM_PAGE = Split(frm129_LM_TOTAL_PAGE, ".")(0)
        frm129_LM_PAGE_LEBIHAN = Split(frm129_LM_TOTAL_PAGE, ".")(1)
        
        If frm129_LM_PAGE_LEBIHAN <> "00" Then
            frm129.L68_Text = frm129_LM_PAGE + 1
        Else
            frm129.L68_Text = frm129_LM_PAGE
        End If
        
    Else
    
        frm129.L68_Text = frm129_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        frm129.L68_Text = 0
    End If
Else
    frm129.L68_Text = 0
End If

rs.Close
Set rs = Nothing

Dim LM_BERAT_1 As Double
Dim LM_BERAT_2 As Double

LM_BERAT_1 = 0
LM_BERAT_2 = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm129.L5_Text = 0 Then rs.Open "select SUM(Berat) from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND StatusItem = 10 AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "')", cn, adOpenKeyset, adLockOptimistic
If frm129.L5_Text = 1 Then rs.Open "select SUM(Berat) from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND StatusItem = 10 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "')", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then LM_BERAT_1 = Format(rs(0), "#,##0.00")

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm129.L5_Text = 0 Then rs.Open "select SUM(Beza_Berat) from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND (StatusItem = 12 OR StatusItem = 22 OR StatusItem = 28) AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "')", cn, adOpenKeyset, adLockOptimistic
If frm129.L5_Text = 1 Then rs.Open "select SUM(Beza_Berat) from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND (StatusItem = 12 OR StatusItem = 22 OR StatusItem = 28) AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "')", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then LM_BERAT_2 = rs(0)

rs.Close
Set rs = Nothing

frm129.L11_Text = Format(LM_BERAT_1 + LM_BERAT_2, "#,##0.00 g")

If x <> 0 Then
    frm129.L69_Text = LM_START_ROW
End If

If frm129.L67_Text <> vbNullString And IsNumeric(frm129.L67_Text) Then
    If frm129.L68_Text <> vbNullString And IsNumeric(frm129.L68_Text) Then
        frm129_LM_CURR_PAGE = frm129.L67_Text
        frm129_LM_TOTAL_PAGE = frm129.L68_Text
        
        If frm129_LM_CURR_PAGE > frm129_LM_TOTAL_PAGE Then
            
            frm129.L67_Text = frm129.L67_Text - 1
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            GoTo re_gen_report:
            
        End If
    End If
End If
End Sub
Sub frm129_excel_trade_in_belian()
'on error resume next
LM_FOUND = 0

Set xlObject = New Excel.Application
Set xlWB = xlObject.Workbooks.Add
           
'xlObject.Visible = True
With xlObject.ActiveWorkbook.ActiveSheet

    .Cells.VerticalAlignment = xlCenter
    .Columns("A").ColumnWidth = 5 'No.
    .Columns("B").ColumnWidth = 15 'No. Siri Produk
    .Columns("C").ColumnWidth = 15 'Purity
    .Columns("D").ColumnWidth = 40 'Nama Produk
    .Columns("E").ColumnWidth = 15 'Berat (g)
    .Columns("F").ColumnWidth = 15 '

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
            .Cells(1, 4) = rs!nama_kedai
            .Cells(1, 4).Font.Name = "Times New Roman"
        End If
        If Not IsNull(rs!no_pendaftaran) Then .Cells(2, 4) = rs!no_pendaftaran
        If Not IsNull(rs!alamat) Then .Cells(3, 4) = rs!alamat
        If Not IsNull(rs!no_tel) Then .Cells(4, 4) = rs!no_tel
        If Not IsNull(rs!no_id_gst) Then .Cells(5, 4) = rs!no_id_gst
    End If
    
    rs.Close
    Set rs = Nothing
    '### Maklumat kedai ### - End
    
    x = 0

    .Cells(1, 4).Font.Bold = True
    .Cells(1, 4).Font.Size = 30
    
    For Row = 1 To 5
        .Cells(Row, 4).HorizontalAlignment = xlCenter
    Next Row
    
    .Cells(7, 1) = frm129.L14_Text
    
    .Cells(8, 1) = "No."
    .Cells(8, 2) = "No. Siri Produk"
    .Cells(8, 3) = "Purity"
    .Cells(8, 4) = "Nama Produk"
    .Cells(8, 5) = "Berat (g)"
    '.Cells(8, 6) = "Berat (g)"
    
    For i = 1 To 5
        .Cells(8, i).HorizontalAlignment = xlCenter
        .Cells(8, i).Interior.ColorIndex = 15
        .Cells(8, i).WrapText = True
        .Cells(8, i).Borders.LineStyle = xlContinuous
    Next i

    If frm129.L5_Text = "1" Then '0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh
    
        TM = frm129.L6_Text 'Tarikh mula
        TA = frm129.L7_Text 'Tarikh akhir
    
    End If
    
    If frm129.L8_Text = "Semua kategori" Then
        
        frm129_LM_SEARCH_1 = Null
        frm129_LM_SEARCH_1_LOGIC = "<>"
        
    Else
    
        frm129_LM_SEARCH_1 = frm129.L8_Text
        frm129_LM_SEARCH_1_LOGIC = "="
        
    End If
    
    If frm129.L9_Text = "semua purity" Then
        
        frm129_LM_SEARCH_2 = Null
        frm129_LM_SEARCH_2_LOGIC = "<>"
        
    Else
        
        frm129_LM_SEARCH_2 = frm129.L9_Text
        frm129_LM_SEARCH_2_LOGIC = "="
    
    End If
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    If frm129.L5_Text = 0 Then rs.Open "select * from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND StatusItem <> 0 AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "') order by tarikh_belian ASC , ID ASC", cn, adOpenKeyset, adLockOptimistic
    If frm129.L5_Text = 1 Then rs.Open "select * from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

    While rs.EOF = False
    
        x = x + 1
        .Cells(8 + x, 1) = x 'No.
        .Cells(8 + x, 1).HorizontalAlignment = xlCenter
        
        If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 2) = rs!no_siri_Produk 'No. Siri Produk

        If Not IsNull(rs!kod_Purity) Then .Cells(8 + x, 3) = rs!kod_Purity 'Purity
        
        If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 4) = rs!kategori_Produk 'Nama Produk
        
        .Cells(8 + x, 5).HorizontalAlignment = xlRight
        If Not IsNull(rs!Berat) Then
            .Cells(8 + x, 5) = Format(rs!Berat, "#,##0.00") 'Berat (g)
            .Cells(8 + x, 5).NumberFormat = "#,##0.00"
        End If
                            
        For Col = 1 To 5
            .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
        Next Col
            
        rs.MoveNext
        
    Wend
    
    rs.Close
    Set rs = Nothing
    
    Y = 0
    Y = x + 2
    
    .Cells(8 + Y, 1) = "Bil : " & frm129.L10_Text
    .Cells(8 + Y, 1).Font.Bold = True
    
    Y = Y + 1
    .Cells(8 + Y, 1) = "Berat : " & frm129.L11_Text
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
End Sub
Sub frm129_excel_trade_in_jualan()
'on error resume next
LM_FOUND = 0

Set xlObject = New Excel.Application
Set xlWB = xlObject.Workbooks.Add
           
'xlObject.Visible = True
With xlObject.ActiveWorkbook.ActiveSheet

    .Cells.VerticalAlignment = xlCenter
    .Columns("A").ColumnWidth = 5 'No.
    .Columns("B").ColumnWidth = 15 'No. Siri Produk
    .Columns("C").ColumnWidth = 15 'Purity
    .Columns("D").ColumnWidth = 40 'Nama Produk
    .Columns("E").ColumnWidth = 15 'Berat (g)
    .Columns("F").ColumnWidth = 15 '

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
            .Cells(1, 4) = rs!nama_kedai
            .Cells(1, 4).Font.Name = "Times New Roman"
        End If
        If Not IsNull(rs!no_pendaftaran) Then .Cells(2, 4) = rs!no_pendaftaran
        If Not IsNull(rs!alamat) Then .Cells(3, 4) = rs!alamat
        If Not IsNull(rs!no_tel) Then .Cells(4, 4) = rs!no_tel
        If Not IsNull(rs!no_id_gst) Then .Cells(5, 4) = rs!no_id_gst
    End If
    
    rs.Close
    Set rs = Nothing
    '### Maklumat kedai ### - End
    
    x = 0

    .Cells(1, 4).Font.Bold = True
    .Cells(1, 4).Font.Size = 30
    
    For Row = 1 To 5
        .Cells(Row, 4).HorizontalAlignment = xlCenter
    Next Row
    
    .Cells(7, 1) = frm129.L14_Text
    
    .Cells(8, 1) = "No."
    .Cells(8, 2) = "No. Siri Produk"
    .Cells(8, 3) = "Purity"
    .Cells(8, 4) = "Nama Produk"
    .Cells(8, 5) = "Berat (g)"
    '.Cells(8, 6) = "Berat (g)"
    
    For i = 1 To 5
        .Cells(8, i).HorizontalAlignment = xlCenter
        .Cells(8, i).Interior.ColorIndex = 15
        .Cells(8, i).WrapText = True
        .Cells(8, i).Borders.LineStyle = xlContinuous
    Next i

    If frm129.L5_Text = "1" Then '0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh
    
        TM = frm129.L6_Text 'Tarikh mula
        TA = frm129.L7_Text 'Tarikh akhir
    
    End If
    
    If frm129.L8_Text = "Semua kategori" Then
        
        frm129_LM_SEARCH_1 = Null
        frm129_LM_SEARCH_1_LOGIC = "<>"
        
    Else
    
        frm129_LM_SEARCH_1 = frm129.L8_Text
        frm129_LM_SEARCH_1_LOGIC = "="
        
    End If
    
    If frm129.L9_Text = "semua purity" Then
        
        frm129_LM_SEARCH_2 = Null
        frm129_LM_SEARCH_2_LOGIC = "<>"
        
    Else
        
        frm129_LM_SEARCH_2 = frm129.L9_Text
        frm129_LM_SEARCH_2_LOGIC = "="
    
    End If
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    If frm129.L5_Text = 0 Then rs.Open "select * from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND (StatusItem = 11 OR StatusItem = 12 OR StatusItem = 21 OR StatusItem = 22 OR StatusItem = 27 OR StatusItem = 28) AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "') order by tarikh_belian ASC , ID ASC", cn, adOpenKeyset, adLockOptimistic
    If frm129.L5_Text = 1 Then rs.Open "select * from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND (StatusItem = 11 OR StatusItem = 12 OR StatusItem = 21 OR StatusItem = 22 OR StatusItem = 27 OR StatusItem = 28) AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "') order by tarikh_belian ASC , ID ASC", cn, adOpenKeyset, adLockOptimistic

    While rs.EOF = False
    
        x = x + 1
        .Cells(8 + x, 1) = x 'No.
        .Cells(8 + x, 1).HorizontalAlignment = xlCenter
        
        If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 2) = rs!no_siri_Produk 'No. Siri Produk

        If Not IsNull(rs!kod_Purity) Then .Cells(8 + x, 3) = rs!kod_Purity 'Purity
        
        If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 4) = rs!kategori_Produk 'Nama Produk
        
        .Cells(8 + x, 5).HorizontalAlignment = xlRight
        If Not IsNull(rs!Berat) Then
            .Cells(8 + x, 5) = Format(rs!Berat - rs!beza_berat, "#,##0.00") 'Berat (g)
            .Cells(8 + x, 5).NumberFormat = "#,##0.00"
        End If
                            
        For Col = 1 To 5
            .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
        Next Col
            
        rs.MoveNext
        
    Wend
    
    rs.Close
    Set rs = Nothing
    
    Y = 0
    Y = x + 2
    
    .Cells(8 + Y, 1) = "Bil : " & frm129.L10_Text
    .Cells(8 + Y, 1).Font.Bold = True
    
    Y = Y + 1
    .Cells(8 + Y, 1) = "Berat : " & frm129.L11_Text
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
End Sub
Sub frm129_excel_trade_in_stok()
'on error resume next
LM_FOUND = 0

Set xlObject = New Excel.Application
Set xlWB = xlObject.Workbooks.Add
           
'xlObject.Visible = True
With xlObject.ActiveWorkbook.ActiveSheet

    .Cells.VerticalAlignment = xlCenter
    .Columns("A").ColumnWidth = 5 'No.
    .Columns("B").ColumnWidth = 15 'No. Siri Produk
    .Columns("C").ColumnWidth = 15 'Purity
    .Columns("D").ColumnWidth = 40 'Nama Produk
    .Columns("E").ColumnWidth = 15 'Berat (g)
    .Columns("F").ColumnWidth = 15 '

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
            .Cells(1, 4) = rs!nama_kedai
            .Cells(1, 4).Font.Name = "Times New Roman"
        End If
        If Not IsNull(rs!no_pendaftaran) Then .Cells(2, 4) = rs!no_pendaftaran
        If Not IsNull(rs!alamat) Then .Cells(3, 4) = rs!alamat
        If Not IsNull(rs!no_tel) Then .Cells(4, 4) = rs!no_tel
        If Not IsNull(rs!no_id_gst) Then .Cells(5, 4) = rs!no_id_gst
    End If
    
    rs.Close
    Set rs = Nothing
    '### Maklumat kedai ### - End
    
    x = 0

    .Cells(1, 4).Font.Bold = True
    .Cells(1, 4).Font.Size = 30
    
    For Row = 1 To 5
        .Cells(Row, 4).HorizontalAlignment = xlCenter
    Next Row
    
    .Cells(7, 1) = frm129.L14_Text
    
    .Cells(8, 1) = "No."
    .Cells(8, 2) = "No. Siri Produk"
    .Cells(8, 3) = "Purity"
    .Cells(8, 4) = "Nama Produk"
    .Cells(8, 5) = "Berat (g)"
    '.Cells(8, 6) = "Berat (g)"
    
    For i = 1 To 5
        .Cells(8, i).HorizontalAlignment = xlCenter
        .Cells(8, i).Interior.ColorIndex = 15
        .Cells(8, i).WrapText = True
        .Cells(8, i).Borders.LineStyle = xlContinuous
    Next i

    If frm129.L5_Text = "1" Then '0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh
    
        TM = frm129.L6_Text 'Tarikh mula
        TA = frm129.L7_Text 'Tarikh akhir
    
    End If
    
    If frm129.L8_Text = "Semua kategori" Then
        
        frm129_LM_SEARCH_1 = Null
        frm129_LM_SEARCH_1_LOGIC = "<>"
        
    Else
    
        frm129_LM_SEARCH_1 = frm129.L8_Text
        frm129_LM_SEARCH_1_LOGIC = "="
        
    End If
    
    If frm129.L9_Text = "semua purity" Then
        
        frm129_LM_SEARCH_2 = Null
        frm129_LM_SEARCH_2_LOGIC = "<>"
        
    Else
        
        frm129_LM_SEARCH_2 = frm129.L9_Text
        frm129_LM_SEARCH_2_LOGIC = "="
    
    End If
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    If frm129.L5_Text = 0 Then rs.Open "select * from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND (StatusItem = 10 OR StatusItem = 12 OR StatusItem = 22 OR StatusItem = 28) AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "') order by tarikh_belian ASC , ID ASC", cn, adOpenKeyset, adLockOptimistic
    If frm129.L5_Text = 1 Then rs.Open "select * from Data_Database where kod_Purity " & frm129_LM_SEARCH_2_LOGIC & "'" & frm129_LM_SEARCH_2 & "' AND kategori_Produk " & frm129_LM_SEARCH_1_LOGIC & "'" & frm129_LM_SEARCH_1 & "' AND (StatusItem = 10 OR StatusItem = 12 OR StatusItem = 22 OR StatusItem = 28) AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "') order by tarikh_belian ASC , ID ASC", cn, adOpenKeyset, adLockOptimistic

    While rs.EOF = False
    
        x = x + 1
        .Cells(8 + x, 1) = x 'No.
        .Cells(8 + x, 1).HorizontalAlignment = xlCenter
        
        If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 2) = rs!no_siri_Produk 'No. Siri Produk

        If Not IsNull(rs!kod_Purity) Then .Cells(8 + x, 3) = rs!kod_Purity 'Purity
        
        If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 4) = rs!kategori_Produk 'Nama Produk
        
        .Cells(8 + x, 5).HorizontalAlignment = xlRight
        If rs!StatusItem = "10" Then
            If Not IsNull(rs!Berat) Then .Cells(8 + x, 5) = Format(rs!Berat, "#,##0.00") 'Berat (g)
        Else
            If Not IsNull(rs!beza_berat) Then .Cells(8 + x, 5) = Format(rs!beza_berat, "#,##0.00") 'Berat (g)
        End If
        
        .Cells(8 + x, 5).NumberFormat = "#,##0.00"
                            
        For Col = 1 To 5
            .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
        Next Col
            
        rs.MoveNext
        
    Wend
    
    rs.Close
    Set rs = Nothing
    
    Y = 0
    Y = x + 2
    
    .Cells(8 + Y, 1) = "Bil : " & frm129.L10_Text
    .Cells(8 + Y, 1).Font.Bold = True
    
    Y = Y + 1
    .Cells(8 + Y, 1) = "Berat : " & frm129.L11_Text
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
End Sub

