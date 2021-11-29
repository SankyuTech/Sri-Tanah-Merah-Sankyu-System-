Attribute VB_Name = "Module47"
Sub frm126_barang_hilang_header()
'on error resume next
frm126.MSFlexGrid1.Clear
frm126.MSFlexGrid1.RowHeight(0) = 700
frm126.MSFlexGrid1.FormatString = "No.|<No.|<ID|<Tarikh|<No. Siri Produk|<Kategori Produk|<Purity|<Berat (g)|<Modal (RM)|<Dulang|<Sebab"

'No.
'Tarikh
'No. Siri Produk
'Kategori Produk
'Purity
'Berat (g)
'Modal (RM)
'Dulang
'Sebab

frm126.MSFlexGrid1.Rows = 1
frm126.MSFlexGrid1.ColWidth(0) = 0 'No.
frm126.MSFlexGrid1.ColAlignment(0) = 4

frm126.MSFlexGrid1.ColWidth(1) = 700 'No.
frm126.MSFlexGrid1.ColAlignment(1) = 4

frm126.MSFlexGrid1.ColWidth(2) = 0 'No. ID

frm126.MSFlexGrid1.ColWidth(3) = 1500 'Tarikh
frm126.MSFlexGrid1.ColAlignment(3) = 4

frm126.MSFlexGrid1.ColWidth(4) = 1700 'No. Siri Produk
frm126.MSFlexGrid1.ColAlignment(4) = 4

frm126.MSFlexGrid1.ColWidth(5) = 3700 'Kategori Produk

frm126.MSFlexGrid1.ColWidth(6) = 1000 'Purity
frm126.MSFlexGrid1.ColAlignment(6) = 4

frm126.MSFlexGrid1.ColWidth(7) = 1000 'Berat (g)
frm126.MSFlexGrid1.ColAlignment(7) = 7

frm126.MSFlexGrid1.ColWidth(8) = 1000 'Modal (RM)
frm126.MSFlexGrid1.ColAlignment(8) = 7

frm126.MSFlexGrid1.ColWidth(9) = 900 'Dulang
frm126.MSFlexGrid1.ColAlignment(9) = 4

frm126.MSFlexGrid1.ColWidth(10) = 7800 'Sebab
End Sub
Sub frm126_barang_hilang()
'on error resume next
Dim frm126_LM_TOTAL_PAGE As Double

frm126_PAGE_SIZE = 33
frm126_LM_TOTAL_PAGE = 0
x = 0

re_gen_report:

frm126.L10_Text = 0
frm126.L11_Text = "0.00 g"
frm126.L12_Text = "RM 0.00"

If frm126.L5_Text = "1" Then '0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh

    TM = frm126.L6_Text 'Tarikh mula
    TA = frm126.L7_Text 'Tarikh akhir

End If

If frm126.L5_Text = 0 Then frm126.L14_Text = "Senarai barang yang hilang atau dicuri."
If frm126.L5_Text = 1 Then frm126.L14_Text = "Senarai barang yang hilang atau dicuri dari " & TM & " hingga " & TA & "."

LM_START_ROW = frm126.L69_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + frm126_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If frm126.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - frm126_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        frm126.L67_Text = 1
    End If
End If

frm126_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm126.L5_Text = 0 Then rs.Open "select * from 86_barang_hilang where status = 1 order by tarikh ASC LIMIT " & LM_START_ROW & "," & frm126_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If frm126.L5_Text = 1 Then rs.Open "select * from 86_barang_hilang where status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & frm126_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1
    If frm126_LM_PAGE_FOUND = 0 Then
        If frm126.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                frm126.L67_Text = frm126.L67_Text + 1 'Paparan Page ke-xxx
                frm126_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(frm126.L67_Text) Then
                    If frm126.L67_Text <> 1 Then
                        frm126.L67_Text = frm126.L67_Text - 1 'Paparan Page ke-xxx
                        frm126_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If

    Y = ((frm126.L67_Text - 1) * frm126_PAGE_SIZE) + x
    frm126.MSFlexGrid1.Rows = x + 1
    frm126.MSFlexGrid1.TextMatrix(x, 0) = x 'No.
    frm126.MSFlexGrid1.TextMatrix(x, 1) = Y 'No.
    frm126.MSFlexGrid1.ColAlignment(1) = 4
    frm126.MSFlexGrid1.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!tarikh) Then frm126.MSFlexGrid1.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!no_siri_Produk) Then frm126.MSFlexGrid1.TextMatrix(x, 4) = rs!no_siri_Produk 'No. Siri Produk
    If Not IsNull(rs!kategori_Produk) Then frm126.MSFlexGrid1.TextMatrix(x, 5) = rs!kategori_Produk 'Kategori Produk
    If Not IsNull(rs!purity) Then frm126.MSFlexGrid1.TextMatrix(x, 6) = rs!purity 'Purity
    If Not IsNull(rs!beza_berat) Then frm126.MSFlexGrid1.TextMatrix(x, 7) = Format(rs!beza_berat, "#,##0.00") 'Berat (g)
    If Not IsNull(rs!harga_item) Then frm126.MSFlexGrid1.TextMatrix(x, 8) = Format(rs!harga_item, "#,##0.00") 'Modal (RM)
    If Not IsNull(rs!dulang) Then frm126.MSFlexGrid1.TextMatrix(x, 9) = rs!dulang 'Dulang
    If Not IsNull(rs!sebab) Then frm126.MSFlexGrid1.TextMatrix(x, 10) = rs!sebab 'Sebab
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm126.L5_Text = 0 Then rs.Open "select COUNT(ID) from 86_barang_hilang where status = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
If frm126.L5_Text = 1 Then rs.Open "select COUNT(ID) from 86_barang_hilang where status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    frm126_LM_TOTAL_PAGE = Format(rs(0) / frm126_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, frm126_LM_TOTAL_PAGE, ".") <> 0 Then
    
        frm126_LM_PAGE = Split(frm126_LM_TOTAL_PAGE, ".")(0)
        frm126_LM_PAGE_LEBIHAN = Split(frm126_LM_TOTAL_PAGE, ".")(1)
        
        If frm126_LM_PAGE_LEBIHAN <> "00" Then
            frm126.L68_Text = frm126_LM_PAGE + 1
        Else
            frm126.L68_Text = frm126_LM_PAGE
        End If
        
    Else
    
        frm126.L68_Text = frm126_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        frm126.L68_Text = 0
    End If
Else
    frm126.L68_Text = 0
End If

rs.Close
Set rs = Nothing

Dim LM_BERAT_ASAL As Double
Dim LM_BERAT_GUNA As Double

LM_BERAT_ASAL = 0
LM_BERAT_GUNA = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm126.L5_Text = 0 Then rs.Open "select COUNT(ID) , SUM(beza_berat) , SUM(harga_item) from 86_barang_hilang where status = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
If frm126.L5_Text = 1 Then rs.Open "select COUNT(ID) , SUM(beza_berat) , SUM(harga_item) from 86_barang_hilang where status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then frm126.L10_Text = rs(0)
If Not IsNull(rs(1)) Then frm126.L11_Text = Format(rs(1), "#,##0.00 g")
If Not IsNull(rs(2)) Then frm126.L12_Text = "RM " & Format(rs(2), "#,##0.00")

rs.Close
Set rs = Nothing

If x <> 0 Then
    frm126.L69_Text = LM_START_ROW
End If

If frm126.L67_Text <> vbNullString And IsNumeric(frm126.L67_Text) Then
    If frm126.L68_Text <> vbNullString And IsNumeric(frm126.L68_Text) Then
        frm126_LM_CURR_PAGE = frm126.L67_Text
        frm126_LM_TOTAL_PAGE = frm126.L68_Text
        
        If frm126_LM_CURR_PAGE > frm126_LM_TOTAL_PAGE Then
            
            frm126.L67_Text = frm126.L67_Text - 1
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            GoTo re_gen_report:
            
        End If
    End If
End If

End Sub
