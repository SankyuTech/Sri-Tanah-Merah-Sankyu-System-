Attribute VB_Name = "Module45"
Sub frm124_initial_setting()
'on error resume next
frm124.CB1 = 0
frm124.DTPicker1 = DateTime.Date
frm124.DTPicker2 = DateTime.Date

frm124.CBB1.Clear
frm124.CBB2.Clear

frm124.CBB1.AddItem "semua urusan"
frm124.CBB1.AddItem "Jualan"
frm124.CBB1.AddItem "GDN"

frm124.CBB1 = "semua urusan"

frm124.CBB2.AddItem "semua purity"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where Metal_Purity<>'" & Null & "' AND status = 1 order by Metal_Purity ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!Metal_Purity) Then frm124.CBB2.AddItem rs!Metal_Purity
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

frm124.CBB2 = "semua purity"

frm124.L5_Text = vbNullString
frm124.L6_Text = vbNullString
frm124.L7_Text = vbNullString
frm124.L8_Text = vbNullString
frm124.L9_Text = vbNullString

frm124.L10_Text = Format(0, "#,##0.00 g")
frm124.L11_Text = Format(0, "#,##0.00 g")
frm124.L12_Text = Format(0, "#,##0.00 g")
End Sub
Sub frm124_report_trade_in_header()
'on error resume next
frm124.MSFlexGrid1.Clear
frm124.MSFlexGrid1.RowHeight(0) = 700
frm124.MSFlexGrid1.FormatString = "No.|<No.|<ID|<Jenis|<Tarikh|<No. Rujukan|<Purity|<Berat (g)"

'No.
'Jenis
'Tarikh
'No. Rujukan
'Purity
'Berat (g)

frm124.MSFlexGrid1.Rows = 1
frm124.MSFlexGrid1.ColWidth(0) = 0 'No.
frm124.MSFlexGrid1.ColAlignment(0) = 4

frm124.MSFlexGrid1.ColWidth(1) = 800 'No.
frm124.MSFlexGrid1.ColAlignment(1) = 4

frm124.MSFlexGrid1.ColWidth(2) = 0 'No. ID

frm124.MSFlexGrid1.ColWidth(3) = 1500 'Jenis
frm124.MSFlexGrid1.ColAlignment(3) = 4

frm124.MSFlexGrid1.ColWidth(4) = 1500 'Tarikh
frm124.MSFlexGrid1.ColAlignment(4) = 4

frm124.MSFlexGrid1.ColWidth(5) = 2000 'No. Rujukan

frm124.MSFlexGrid1.ColWidth(6) = 1500 'Purity
frm124.MSFlexGrid1.ColAlignment(6) = 4

frm124.MSFlexGrid1.ColWidth(7) = 1500 'Berat (g)
frm124.MSFlexGrid1.ColAlignment(7) = 7
End Sub
Sub frm124_report_trade_in()
'on error resume next
Dim frm124_LM_TOTAL_PAGE As Double

frm124_PAGE_SIZE = 30
frm124_LM_TOTAL_PAGE = 0
x = 0

re_gen_report:

frm124.L10_Text = "0.00 g"
frm124.L11_Text = "0.00 g"
frm124.L12_Text = "0.00 g"

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

If frm124.L5_Text = 0 Then frm124.L14_Text = "Rekod penggunaan/jualan barang trade in atau barang potong bagi " & frm124.L8_Text & " dan " & frm124.L9_Text & "."
If frm124.L5_Text = 1 Then frm124.L14_Text = "Rekod penggunaan/jualan barang trade in atau barang potong bagi " & frm124.L8_Text & " dan " & frm124.L9_Text & " dari " & TM & " hingga " & TA & "."

LM_START_ROW = frm124.L69_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + frm124_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If frm124.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - frm124_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        frm124.L67_Text = 1
    End If
End If

frm124_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm124.L5_Text = 0 Then rs.Open "select * from 85_penggunaan_ti where (menu='" & frm124_LM_SEARCH_1 & "' OR menu='" & frm124_LM_SEARCH_2 & "') AND purity " & frm124_LM_SEARCH_3_LOGIC & "'" & frm124_LM_SEARCH_3 & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & frm124_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If frm124.L5_Text = 1 Then rs.Open "select * from 85_penggunaan_ti where (menu='" & frm124_LM_SEARCH_1 & "' OR menu='" & frm124_LM_SEARCH_2 & "') AND purity " & frm124_LM_SEARCH_3_LOGIC & "'" & frm124_LM_SEARCH_3 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & frm124_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1
    If frm124_LM_PAGE_FOUND = 0 Then
        If frm124.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                frm124.L67_Text = frm124.L67_Text + 1 'Paparan Page ke-xxx
                frm124_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(frm124.L67_Text) Then
                    If frm124.L67_Text <> 1 Then
                        frm124.L67_Text = frm124.L67_Text - 1 'Paparan Page ke-xxx
                        frm124_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If

    Y = ((frm124.L67_Text - 1) * frm124_PAGE_SIZE) + x
    frm124.MSFlexGrid1.Rows = x + 1
    frm124.MSFlexGrid1.TextMatrix(x, 0) = x 'No.
    frm124.MSFlexGrid1.TextMatrix(x, 1) = Y 'No.
    frm124.MSFlexGrid1.ColAlignment(1) = 4
    frm124.MSFlexGrid1.TextMatrix(x, 2) = rs!ID 'No. ID
    
    If Not IsNull(rs!Menu) Then 'Jenis
        
        If rs!Menu = 0 Then
            
            frm124.MSFlexGrid1.TextMatrix(x, 3) = "GDN"
            
        ElseIf rs!Menu = 1 Then
            
            frm124.MSFlexGrid1.TextMatrix(x, 3) = "Jualan"
            
        End If
        
    End If
    
    If Not IsNull(rs!tarikh) Then frm124.MSFlexGrid1.TextMatrix(x, 4) = rs!tarikh 'Tarikh
    
    If Not IsNull(rs!no_rujukan) Then frm124.MSFlexGrid1.TextMatrix(x, 5) = rs!no_rujukan 'No. Rujukan
    
    If Not IsNull(rs!purity) Then frm124.MSFlexGrid1.TextMatrix(x, 6) = rs!purity 'Purity
    
    If Not IsNull(rs!Berat) Then frm124.MSFlexGrid1.TextMatrix(x, 7) = Format(rs!Berat, "#,##0.00") 'Berat (g)

    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm124.L5_Text = 0 Then rs.Open "select COUNT(ID) from 85_penggunaan_ti where (menu='" & frm124_LM_SEARCH_1 & "' OR menu='" & frm124_LM_SEARCH_2 & "') AND purity " & frm124_LM_SEARCH_3_LOGIC & "'" & frm124_LM_SEARCH_3 & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
If frm124.L5_Text = 1 Then rs.Open "select COUNT(ID) from 85_penggunaan_ti where (menu='" & frm124_LM_SEARCH_1 & "' OR menu='" & frm124_LM_SEARCH_2 & "') AND purity " & frm124_LM_SEARCH_3_LOGIC & "'" & frm124_LM_SEARCH_3 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    frm124_LM_TOTAL_PAGE = Format(rs(0) / frm124_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, frm124_LM_TOTAL_PAGE, ".") <> 0 Then
    
        frm124_LM_PAGE = Split(frm124_LM_TOTAL_PAGE, ".")(0)
        frm124_LM_PAGE_LEBIHAN = Split(frm124_LM_TOTAL_PAGE, ".")(1)
        
        If frm124_LM_PAGE_LEBIHAN <> "00" Then
            frm124.L68_Text = frm124_LM_PAGE + 1
        Else
            frm124.L68_Text = frm124_LM_PAGE
        End If
        
    Else
    
        frm124.L68_Text = frm124_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        frm124.L68_Text = 0
    End If
Else
    frm124.L68_Text = 0
End If

rs.Close
Set rs = Nothing

Dim LM_BERAT_ASAL As Double
Dim LM_BERAT_GUNA As Double

LM_BERAT_ASAL = 0
LM_BERAT_GUNA = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(beza_berat) from data_database where Purity " & frm124_LM_SEARCH_3_LOGIC & "'" & frm124_LM_SEARCH_3 & "' AND (((statusitem = 10 OR statusitem = 12 OR statusitem = 20 OR statusitem = 22) AND receiving_Status = 2) OR ((statusitem = 12 OR statusitem = 20 OR statusitem = 22) AND receiving_Status = 0))", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then
    frm124.L10_Text = Format(rs(0), "#,##0.00 g")
    LM_BERAT_ASAL = rs(0)
End If

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm124.L5_Text = 0 Then rs.Open "select SUM(berat) from 85_penggunaan_ti where (menu='" & frm124_LM_SEARCH_1 & "' OR menu='" & frm124_LM_SEARCH_2 & "') AND purity " & frm124_LM_SEARCH_3_LOGIC & "'" & frm124_LM_SEARCH_3 & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
If frm124.L5_Text = 1 Then rs.Open "select SUM(berat) from 85_penggunaan_ti where (menu='" & frm124_LM_SEARCH_1 & "' OR menu='" & frm124_LM_SEARCH_2 & "') AND purity " & frm124_LM_SEARCH_3_LOGIC & "'" & frm124_LM_SEARCH_3 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then
    frm124.L11_Text = Format(rs(0), "#,##0.00 g")
    LM_BERAT_GUNA = rs(0)
End If

rs.Close
Set rs = Nothing

frm124.L12_Text = Format(LM_BERAT_ASAL - LM_BERAT_GUNA, "#,##0.00 g")

If x <> 0 Then
    frm124.L69_Text = LM_START_ROW
End If

If frm124.L67_Text <> vbNullString And IsNumeric(frm124.L67_Text) Then
    If frm124.L68_Text <> vbNullString And IsNumeric(frm124.L68_Text) Then
        frm124_LM_CURR_PAGE = frm124.L67_Text
        frm124_LM_TOTAL_PAGE = frm124.L68_Text
        
        If frm124_LM_CURR_PAGE > frm124_LM_TOTAL_PAGE Then
            
            frm124.L67_Text = frm124.L67_Text - 1
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            GoTo re_gen_report:
            
        End If
    End If
End If

End Sub
