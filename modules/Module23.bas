Attribute VB_Name = "Module16"
Sub frm117_pic_ena_disable()
'On Error Resume Next
frm117.Pic1.Left = 120
frm117.Pic1.Top = 240
frm117.Pic2.Left = 120
frm117.Pic2.Top = 240

frm117.Pic1.Visible = False
frm117.Pic2.Visible = False
End Sub
Sub frm117_initial_setting()
'On Error Resume Next
frm117.CB1 = 0

frm117.DTPicker1 = DateTime.Date
frm117.DTPicker2 = DateTime.Date

frm117.CBB1.Clear

frm117.CBB1.AddItem "Semua supplier dan agen"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where status='" & 1 & "' order by supplier ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!supplier) Then frm117.CBB1.AddItem rs!supplier
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

frm117.CBB1 = "Semua supplier dan agen"

frm117.CBB2.Clear

frm117.CBB2.AddItem "Semua GDN/GRN/INV/VOU"
frm117.CBB2.AddItem "GDN"
frm117.CBB2.AddItem "GRN"
frm117.CBB2.AddItem "INV"
frm117.CBB2.AddItem "VOU"

frm117.CBB2 = "Semua GDN/GRN/INV/VOU"
End Sub
Sub frm117_report_gdn_grn_header()
'on error resume next
frm117.MSFlexGrid1.Clear
frm117.MSFlexGrid1.RowHeight(0) = 700
frm117.MSFlexGrid1.FormatString = "No.|<No.|<ID|<Tarikh|<Jenis|<No. Rujukan|<Supplier/Agen|<Berat (g)|<Mutu|<Hutang (Emas)|<Bayar (Emas)|<Hutang (RM)|<Bayar (RM)|<GST (RM)"

'No.
'Tarikh
'Jenis
'No. Rujukan
'Supplier/Agen
'Berat (g)
'Mutu
'Hutang (Emas)
'Bayar (Emas)
'Hutang (RM)
'Bayar (RM)
'GST (RM)

frm117.MSFlexGrid1.Rows = 1
frm117.MSFlexGrid1.ColWidth(0) = 600 'No.
frm117.MSFlexGrid1.ColAlignment(0) = 4

frm117.MSFlexGrid1.ColWidth(1) = 0 'No.
frm117.MSFlexGrid1.ColWidth(2) = 0 'No. ID

frm117.MSFlexGrid1.ColWidth(3) = 1200 'Tarikh

frm117.MSFlexGrid1.ColWidth(4) = 800 'Jenis
frm117.MSFlexGrid1.ColAlignment(4) = 4

frm117.MSFlexGrid1.ColWidth(5) = 1500 'No. Rujukan
frm117.MSFlexGrid1.ColAlignment(5) = 4

frm117.MSFlexGrid1.ColWidth(6) = 4500 'Supplier/Agen

frm117.MSFlexGrid1.ColWidth(7) = 1200 'Berat (g)
frm117.MSFlexGrid1.ColAlignment(7) = 7

frm117.MSFlexGrid1.ColWidth(8) = 800 'Mutu
frm117.MSFlexGrid1.ColAlignment(8) = 7

frm117.MSFlexGrid1.ColWidth(9) = 1200 'Hutang (Emas)
frm117.MSFlexGrid1.ColAlignment(9) = 7

frm117.MSFlexGrid1.ColWidth(10) = 1200 'Bayar (Emas)
frm117.MSFlexGrid1.ColAlignment(10) = 7

frm117.MSFlexGrid1.ColWidth(11) = 1200 'Hutang (RM)
frm117.MSFlexGrid1.ColAlignment(11) = 7

frm117.MSFlexGrid1.ColWidth(12) = 1200 'Bayar (RM)
frm117.MSFlexGrid1.ColAlignment(12) = 7

frm117.MSFlexGrid1.ColWidth(13) = 1200 'GST (RM)
frm117.MSFlexGrid1.ColAlignment(13) = 7
End Sub
Sub frm117_report_gdn_grn()
'On Error Resume Next
Dim frm117_LM_TOTAL_PAGE As Double
Dim TM As Date
Dim TA As Date
Dim LM_BERAT_BAYAR As Double
Dim LM_BERAT_HUTANG As Double
Dim LM_TUNAI_BAYAR As Double
Dim LM_TUNAI_HUTANG As Double
Dim LM_BAYAR_EMAS As Double
Dim LM_TUNAI_HUTANG2 As Double
Dim LM_GST_1 As Double
Dim LM_GST_2 As Double

frm117_PAGE_SIZE = 37
frm117_LM_TOTAL_PAGE = 0
x = 0

re_gen_report:

frm117.L13_Text = 0
frm117.L20_Text = "0.00"
frm117.L21_Text = "0.00"
frm117.L22_Text = "0.00"
frm117.L23_Text = "0.00"
frm117.L24_Text = "0.00"
frm117.L25_Text = "0.00"
frm117.L26_Text = "0.00"
frm117.L27_Text = "0.00"

If frm117.L5_Text = "1" Then '0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh

    TM = frm117.L6_Text 'Tarikh mula
    TA = frm117.L7_Text 'Tarikh akhir

End If

If frm117.L8_Text = "Semua supplier dan agen" Then
    
    frm117_LM_SEARCH_1 = Null
    frm117_LM_SEARCH_1_LOGIC = "<>"
    
Else
    
    frm117_LM_SEARCH_1 = frm117.L8_Text
    frm117_LM_SEARCH_1_LOGIC = "="

End If

If frm117.L9_Text = "Semua GDN/GRN/INV/VOU" Then
    
    frm117_LM_SEARCH_2 = Null
    frm117_LM_SEARCH_2_LOGIC = "<>"

Else
    
    frm117_LM_SEARCH_2 = frm117.L9_Text
    frm117_LM_SEARCH_2_LOGIC = "="
    
End If

LM_START_ROW = frm117.L69_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + frm117_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If frm117.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - frm117_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        frm117.L67_Text = 1
    End If
End If

frm117_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm117.L5_Text = 0 Then rs.Open "select * from 77_gdn_grn where supplier_agen " & frm117_LM_SEARCH_1_LOGIC & "'" & frm117_LM_SEARCH_1 & "' AND status = 1 AND jenis " & frm117_LM_SEARCH_2_LOGIC & "'" & frm117_LM_SEARCH_2 & "' order by write_timestamp ASC LIMIT " & LM_START_ROW & "," & frm117_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If frm117.L5_Text = 1 Then rs.Open "select * from 77_gdn_grn where supplier_agen " & frm117_LM_SEARCH_1_LOGIC & "'" & frm117_LM_SEARCH_1 & "' AND status = 1 AND jenis " & frm117_LM_SEARCH_2_LOGIC & "'" & frm117_LM_SEARCH_2 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by write_timestamp ASC LIMIT " & LM_START_ROW & "," & frm117_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1
    If frm117_LM_PAGE_FOUND = 0 Then
        If frm117.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                frm117.L67_Text = frm117.L67_Text + 1 'Paparan Page ke-xxx
                frm117_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(frm117.L67_Text) Then
                    If frm117.L67_Text <> 1 Then
                        frm117.L67_Text = frm117.L67_Text - 1 'Paparan Page ke-xxx
                        frm117_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If

    Y = ((frm117.L67_Text - 1) * frm117_PAGE_SIZE) + x
    frm117.MSFlexGrid1.Rows = x + 1
    frm117.MSFlexGrid1.TextMatrix(x, 0) = x 'No.
    frm117.MSFlexGrid1.TextMatrix(x, 1) = Y 'No.
    frm117.MSFlexGrid1.ColAlignment(1) = 4
    frm117.MSFlexGrid1.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!tarikh) Then frm117.MSFlexGrid1.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!jenis) Then frm117.MSFlexGrid1.TextMatrix(x, 4) = rs!jenis 'Jenis
    If Not IsNull(rs!no_rujukan) Then frm117.MSFlexGrid1.TextMatrix(x, 5) = rs!no_rujukan 'No. Rujukan
    If Not IsNull(rs!supplier_agen) Then frm117.MSFlexGrid1.TextMatrix(x, 6) = rs!supplier_agen 'Supplier/Agen
    If Not IsNull(rs!Berat_Asal) Then frm117.MSFlexGrid1.TextMatrix(x, 7) = Format(rs!Berat_Asal, "#,##0.00") 'Berat (g)
    If Not IsNull(rs!kadar_tukaran) Then frm117.MSFlexGrid1.TextMatrix(x, 8) = Format(rs!kadar_tukaran, "0.000") 'Mutu
    If Not IsNull(rs!berat_tukaran_grn) Then frm117.MSFlexGrid1.TextMatrix(x, 9) = Format(rs!berat_tukaran_grn, "#,##0.00") 'Hutang (Emas)
    If Not IsNull(rs!berat_tukaran) Then frm117.MSFlexGrid1.TextMatrix(x, 10) = Format(rs!berat_tukaran, "#,##0.00") 'Bayar (Emas)
    
    If Not IsNull(rs!jenis_urusan) Then
        
        If rs!jenis_urusan <> "3" Then
        
            If Not IsNull(rs!harga_dengan_gst_grn) Then frm117.MSFlexGrid1.TextMatrix(x, 11) = Format(rs!harga_dengan_gst_grn, "#,##0.00") 'Hutang (RM)
            If Not IsNull(rs!harga_dengan_gst) Then frm117.MSFlexGrid1.TextMatrix(x, 12) = Format(rs!harga_dengan_gst, "#,##0.00") 'Bayar (RM)
            If Not IsNull(rs!jumlah_gst) Then frm117.MSFlexGrid1.TextMatrix(x, 13) = Format(rs!jumlah_gst, "#,##0.00") 'GST (RM)
        
        Else
            
            If rs!umum_berat = "0" Then
        
                If Not IsNull(rs!harga_dengan_gst_grn) Then frm117.MSFlexGrid1.TextMatrix(x, 11) = Format(rs!harga_dengan_gst_grn, "#,##0.00") 'Hutang (RM)
                If Not IsNull(rs!harga_dengan_gst) Then frm117.MSFlexGrid1.TextMatrix(x, 12) = Format(rs!harga_dengan_gst, "#,##0.00") 'Bayar (RM)
                If Not IsNull(rs!jumlah_gst) Then frm117.MSFlexGrid1.TextMatrix(x, 13) = Format(rs!jumlah_gst, "#,##0.00") 'GST (RM)
                
            ElseIf rs!umum_berat = "1" Then
            
            
            End If
            
        End If
    Else
        
        If Not IsNull(rs!harga_dengan_gst_grn) Then frm117.MSFlexGrid1.TextMatrix(x, 11) = Format(rs!harga_dengan_gst_grn, "#,##0.00") 'Hutang (RM)
        If Not IsNull(rs!harga_dengan_gst) Then frm117.MSFlexGrid1.TextMatrix(x, 12) = Format(rs!harga_dengan_gst, "#,##0.00") 'Bayar (RM)
        If Not IsNull(rs!jumlah_gst) Then frm117.MSFlexGrid1.TextMatrix(x, 13) = Format(rs!jumlah_gst, "#,##0.00") 'GST (RM)
        
    End If
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm117.L5_Text = 0 Then rs.Open "select COUNT(ID) from 77_gdn_grn where supplier_agen " & frm117_LM_SEARCH_1_LOGIC & "'" & frm117_LM_SEARCH_1 & "' AND status = 1 AND jenis " & frm117_LM_SEARCH_2_LOGIC & "'" & frm117_LM_SEARCH_2 & "'", cn, adOpenKeyset, adLockOptimistic
If frm117.L5_Text = 1 Then rs.Open "select COUNT(ID) from 77_gdn_grn where supplier_agen " & frm117_LM_SEARCH_1_LOGIC & "'" & frm117_LM_SEARCH_1 & "' AND status = 1 AND jenis " & frm117_LM_SEARCH_2_LOGIC & "'" & frm117_LM_SEARCH_2 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    frm117_LM_TOTAL_PAGE = Format(rs(0) / frm117_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, frm117_LM_TOTAL_PAGE, ".") <> 0 Then
    
        frm117_LM_PAGE = Split(frm117_LM_TOTAL_PAGE, ".")(0)
        frm117_LM_PAGE_LEBIHAN = Split(frm117_LM_TOTAL_PAGE, ".")(1)
        
        If frm117_LM_PAGE_LEBIHAN <> "00" Then
            frm117.L68_Text = frm117_LM_PAGE + 1
        Else
            frm117.L68_Text = frm117_LM_PAGE
        End If
        
    Else
    
        frm117.L68_Text = frm117_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        frm117.L68_Text = 0
    End If
Else
    frm117.L68_Text = 0
End If

rs.Close
Set rs = Nothing

LM_BERAT_BAYAR = 0
LM_BERAT_HUTANG = 0
LM_TUNAI_BAYAR = 0
LM_TUNAI_HUTANG = 0
LM_BAYAR_EMAS = 0
LM_TUNAI_HUTANG2 = 0
LM_GST_1 = 0
LM_GST_2 = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm117.L5_Text = 0 Then rs.Open "select COUNT(ID) , SUM(berat_asal) , SUM(berat_tukaran_grn) , SUM(berat_tukaran) , SUM(harga_dengan_gst_grn) , SUM(harga_dengan_gst) , SUM(jumlah_gst) from 77_gdn_grn where supplier_agen " & frm117_LM_SEARCH_1_LOGIC & "'" & frm117_LM_SEARCH_1 & "' AND status = 1 AND jenis " & frm117_LM_SEARCH_2_LOGIC & "'" & frm117_LM_SEARCH_2 & "'", cn, adOpenKeyset, adLockOptimistic
If frm117.L5_Text = 1 Then rs.Open "select COUNT(ID) , SUM(berat_asal) , SUM(berat_tukaran_grn) , SUM(berat_tukaran) , SUM(harga_dengan_gst_grn) , SUM(harga_dengan_gst) , SUM(jumlah_gst) from 77_gdn_grn where supplier_agen " & frm117_LM_SEARCH_1_LOGIC & "'" & frm117_LM_SEARCH_1 & "' AND status = 1 AND jenis " & frm117_LM_SEARCH_2_LOGIC & "'" & frm117_LM_SEARCH_2 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then frm117.L13_Text = rs(0) 'Jumlah bilangan barang jualan
If Not IsNull(rs(1)) Then frm117.L20_Text = Format(rs(1), "#,##0.00")
If Not IsNull(rs(2)) Then
    frm117.L21_Text = Format(rs(2), "#,##0.00")
    LM_BERAT_BAYAR = rs(2)
End If
If Not IsNull(rs(3)) Then
    frm117.L22_Text = Format(rs(3), "#,##0.00")
    LM_BERAT_HUTANG = rs(3)
End If
If Not IsNull(rs(4)) Then
    frm117.L23_Text = Format(rs(4), "#,##0.00")
    LM_TUNAI_BAYAR = rs(4)
End If
If Not IsNull(rs(5)) Then
    'frm117.L24_Text = Format(rs(5), "#,##0.00")
    LM_TUNAI_HUTANG = rs(5)
End If
'If Not IsNull(rs(6)) Then frm117.L25_Text = Format(rs(6), "#,##0.00")
If Not IsNull(rs(6)) Then LM_GST_1 = Format(rs(6), "#,##0.00")

'SUM(berat_asal) , SUM(berat_tukaran_grn) , SUM(berat_tukaran) , SUM(harga_dengan_gst_grn) , SUM(harga_dengan_gst) , SUM(jumlah_gst)
rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm117.L5_Text = 0 Then rs.Open "select SUM(harga_dengan_gst) , SUM(jumlah_gst) from 77_gdn_grn where jenis_urusan = 3 AND umum_berat = 1 AND supplier_agen " & frm117_LM_SEARCH_1_LOGIC & "'" & frm117_LM_SEARCH_1 & "' AND status = 1 AND jenis " & frm117_LM_SEARCH_2_LOGIC & "'" & frm117_LM_SEARCH_2 & "'", cn, adOpenKeyset, adLockOptimistic
If frm117.L5_Text = 1 Then rs.Open "select SUM(harga_dengan_gst) , SUM(jumlah_gst) from 77_gdn_grn where jenis_urusan = 3 AND umum_berat = 1 AND supplier_agen " & frm117_LM_SEARCH_1_LOGIC & "'" & frm117_LM_SEARCH_1 & "' AND status = 1 AND jenis " & frm117_LM_SEARCH_2_LOGIC & "'" & frm117_LM_SEARCH_2 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then
    'frm117.L24_Text = Format(rs(5), "#,##0.00")
    LM_TUNAI_HUTANG2 = rs(0)
End If
If Not IsNull(rs(1)) Then LM_GST_2 = Format(rs(1), "#,##0.00")

rs.Close
Set rs = Nothing

frm117.L24_Text = Format(LM_TUNAI_HUTANG - LM_TUNAI_HUTANG2, "#,##0.00")
frm117.L25_Text = Format(LM_GST_1 - LM_GST_2, "#,##0.00")

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm117.L5_Text = 0 Then rs.Open "select SUM(harga_dengan_gst) from 77_gdn_grn where jenis_urusan = 3 AND umum_berat = 1 AND supplier_agen " & frm117_LM_SEARCH_1_LOGIC & "'" & frm117_LM_SEARCH_1 & "' AND status = 1 AND jenis " & frm117_LM_SEARCH_2_LOGIC & "'" & frm117_LM_SEARCH_2 & "'", cn, adOpenKeyset, adLockOptimistic
If frm117.L5_Text = 1 Then rs.Open "select SUM(harga_dengan_gst) from 77_gdn_grn where jenis_urusan = 3 AND umum_berat = 1 AND supplier_agen " & frm117_LM_SEARCH_1_LOGIC & "'" & frm117_LM_SEARCH_1 & "' AND status = 1 AND jenis " & frm117_LM_SEARCH_2_LOGIC & "'" & frm117_LM_SEARCH_2 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then LM_BAYAR_EMAS = rs(0) 'Bayaran emas secara RM

'SUM(berat_asal) , SUM(berat_tukaran_grn) , SUM(berat_tukaran) , SUM(harga_dengan_gst_grn) , SUM(harga_dengan_gst) , SUM(jumlah_gst)
rs.Close
Set rs = Nothing

frm117.L26_Text = Format(LM_BERAT_HUTANG - LM_BERAT_BAYAR, "#,##0.00 g")
frm117.L27_Text = "RM " & Format(LM_TUNAI_HUTANG - LM_BAYAR_EMAS - LM_TUNAI_BAYAR, "#,##0.00")

If x <> 0 Then
    frm117.L69_Text = LM_START_ROW
End If

If frm117.L67_Text <> vbNullString And IsNumeric(frm117.L67_Text) Then
    If frm117.L68_Text <> vbNullString And IsNumeric(frm117.L68_Text) Then
        frm117_LM_CURR_PAGE = frm117.L67_Text
        frm117_LM_TOTAL_PAGE = frm117.L68_Text
        
        If frm117_LM_CURR_PAGE > frm117_LM_TOTAL_PAGE Then
            
            frm117.L67_Text = frm117.L67_Text - 1
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            GoTo re_gen_report:
            
        End If
    End If
End If
End Sub
Sub frm117_edit_data_grn()
'on error resume next
LM_FOUND = 0

If G_No_RESIT_JUALAN <> vbNullString Then
    
    Frm116_LM_USER = vbNullString
    
    frm116.L69_Text = -1 'Titik Pencarian Data
    frm116.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    frm116.L67_Text = 0 'Paparan Page ke-xxx
    
    Call frm116_one_time_reset
    Call frm116_reset_1
    Call Frm116_reset_3

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    
    strsql = "insert into " & G_GRN_TEMP & "(id_database,purity,berat_asal," _
                & "berat_tukaran_grn,upah,harga_tanpa_gst_grn,gst_ari_nashi," _
                & "gst_include,jumlah_gst,kadar_gst,harga_dengan_gst_grn,nilaian_harga_emas," _
                & "kadar_tukaran,Status)" & _
                "select ID,purity,berat_asal," _
                & "berat_tukaran_grn,upah,harga_tanpa_gst_grn,gst_ari_nashi," _
                & "gst_include,jumlah_gst,kadar_gst,harga_dengan_gst_grn,nilaian_harga_emas," _
                & "kadar_tukaran,2 " _
                & "from 79_grn WHERE no_rujukan='" & G_No_RESIT_JUALAN & "' AND status = 1"

    Set rs = cn.Execute(strsql)
    Set rs = Nothing
    
    Call Frm116_Senarai_Belian_Header
    Call Frm116_Senarai_Belian

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 77_gdn_grn where no_rujukan='" & G_No_RESIT_JUALAN & "' AND jenis_urusan = 1 AND status = 1", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        
        If Not IsNull(rs!ID) Then frm116.L71_Text = rs!ID
        If Not IsNull(rs!tarikh) Then frm116.DTPicker1 = rs!tarikh
        If Not IsNull(rs!Berat_Asal) Then frm116.L48_Text = Format(rs!Berat_Asal, "#,##0.00") 'Berat asal sebelum tukaran mutu
        If Not IsNull(rs!kadar_tukaran) Then frm116.TB8 = rs!kadar_tukaran
        If Not IsNull(rs!berat_tukaran_grn) Then frm116.L9_Text = Format(rs!berat_tukaran_grn, "#,##0.00")
        If Not IsNull(rs!harga_tanpa_gst) Then frm116.L51_Text = Format(rs!harga_tanpa_gst, "#,##0.00")
        If Not IsNull(rs!jumlah_gst) Then frm116.L52_Text = Format(rs!jumlah_gst, "#,##0.00")
        If Not IsNull(rs!kadar_gst) Then frm116.L22_Text = Format(rs!kadar_gst, "#,##0.00")
        If Not IsNull(rs!harga_dengan_gst_grn) Then frm116.L53_Text = Format(rs!harga_dengan_gst_grn, "#,##0.00")
        If Not IsNull(rs!harga_999) Then frm116.TB6 = Format(rs!harga_999, "#,##0.00")
        If Not IsNull(rs!nilaian_harga_emas) Then frm116.L12_Text = Format(rs!nilaian_harga_emas, "#,##0.00")
        If Not IsNull(rs!gst_zr_harga) Then frm116.L17_Text = Format(rs!gst_zr_harga, "#,##0.00")
        If Not IsNull(rs!gst_sr_harga) Then frm116.L18_Text = Format(rs!gst_sr_harga, "#,##0.00")
        If Not IsNull(rs!gst_zr_cukai) Then frm116.L19_Text = Format(rs!gst_zr_cukai, "#,##0.00")
        If Not IsNull(rs!gst_sr_cukai) Then frm116.L20_Text = Format(rs!gst_sr_cukai, "#,##0.00")
        If Not IsNull(rs!bil_barang) Then frm116.L43_Text = rs!bil_barang
        If Not IsNull(rs!no_rujukan_supplier) Then frm116.TB9 = rs!no_rujukan_supplier
        
        If Not IsNull(rs!supplier_agen) Then
            'on error goto Err_A:
            Frm116_LM_SUPPLIER = rs!supplier_agen
            frm116.CBB2 = Frm116_LM_SUPPLIER
        
Restore_A:
        End If
        
        If Not IsNull(rs!user) Then

            Frm116_LM_USER = rs!user

        End If
        'on error resume next
        LM_FOUND = 1
        
    End If
    
    rs.Close
    Set rs = Nothing
'### Masukkan data voucher / invoice bagi belian agen ini ### - End

'### Carian Maklumat Penjual (Data Pekerja) ### - Start
    If Frm116_LM_USER <> vbNullString Then
    
        DATA_PEKERJA_FOUND = 0
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where Samaran='" & Frm116_LM_USER & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            Frm116_LM_MAKLUMAT_PEKERJA = rs!Samaran & "  |  " & rs!NoPekerja
            DATA_PEKERJA_FOUND = 1
            
        End If
        
        rs.Close
        Set rs = Nothing
    
        If DATA_PEKERJA_FOUND = 1 Then
            'On Error GoTo Err_B:
            frm116.CBB4 = Frm116_LM_MAKLUMAT_PEKERJA
            
Restore_B:
        End If
        
        'on error resume next
    End If
'### Carian Maklumat Penjual (Data Pekerja) ### - End

    If LM_FOUND = 1 Then
    
        frm116.CBB4.Enabled = True
        frm116.CBB4.BackColor = &HFFFFFF
    
        frm116.CMD8.Visible = False
        frm116.CMD9.Visible = False
        frm116.CMD10.Visible = True
        frm116.CMD11.Visible = True
        
        frm116.Show
        frm117.Hide
        
    End If

End If
     
Exit Sub

Err_A:

frm116.CBB2.AddItem Frm116_LM_SUPPLIER
frm116.CBB2 = Frm116_LM_SUPPLIER
            
Resume Restore_A:

Exit Sub
Err_B:
frm116.CBB4.AddItem Frm116_LM_MAKLUMAT_PEKERJA
frm116.CBB4 = Frm116_LM_MAKLUMAT_PEKERJA
Resume Restore_B:
End Sub
Sub frm117_padam_grn()
'On Error Resume Next
'### Masukkan maklumat Good Delivery Note (GRN) ### - Start
DATA_SAVE = 0

LM_NOW = Now
LM_TARIKH = DateTime.Date$
LM_MASA = DateTime.Time$
LM_NO_RUJUKAN = vbNullString

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 77_gdn_grn where ID='" & G_ID & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    G_ID = rs!ID
    Call recovery_77_gdn_grn
    
    rs!Status = 0
    rs!jenis_urusan = 1
    rs!terminal = G_TERMINAL
    rs!user = MDI_frm1.L3_Text 'Nama Pekerja
    rs.Update
    DATA_SAVE = 1
    
End If

rs.Close
Set rs = Nothing
'### Masukkan data voucher / invoice bagi belian agen ini ### - End

'### Masukkan Data Jualan Ke Dalam Table Jualan ### - Start

If DATA_SAVE = 1 Then
    
    '### Transfer data kepada recovery database ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

    strsql = "insert into " & G_RECOVERY_DATABASE & ".79_grn(id_asal,tarikh,masa,write_timestamp,no_rujukan,no_rujukan_rev,purity,berat_asal," _
                & "berat_tukaran_grn,upah,harga_tanpa_gst_grn,gst_ari_nashi," _
                & "gst_include,jumlah_gst,kadar_gst,harga_dengan_gst_grn,nilaian_harga_emas," _
                & "kadar_tukaran,Status,terminal,user)" & _
                "select ID,tarikh,masa,write_timestamp,no_rujukan,no_rujukan_rev,purity,berat_asal," _
                & "berat_tukaran_grn,upah,harga_tanpa_gst_grn,gst_ari_nashi," _
                & "gst_include,jumlah_gst,kadar_gst,harga_dengan_gst_grn,nilaian_harga_emas," _
                & "kadar_tukaran,Status,terminal,user " _
                & "from " & G_SERVER_DATABASE & ".79_grn WHERE no_rujukan='" & G_No_RESIT_JUALAN & "' AND status = 1"
                
    Set rs = cn.Execute(strsql)
    Set rs = Nothing
    '### Transfer data kepada recovery database ### - End
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

    strsql = "UPDATE 79_grn set status='" & 0 & "'," _
    & "user='" & MDI_frm1.L3_Text & "'," _
    & "terminal='" & G_TERMINAL & "'" _
    & "WHERE no_rujukan='" & G_No_RESIT_JUALAN & "' AND status = 1"
    
    Set rs = cn.Execute(strsql)
    Set rs = Nothing

'#### Update Log Aktiviti Sistem #### - Start
    'User = MDI_frm1.L3_Text
    LogAct_Memory = "[" & MDI_frm1.L3_Text & "] Padam data GRN kepada agen/supplier. No. Rujukan [" & G_No_RESIT_JUALAN & "]."
    LogDate_Memory = LM_NOW
    Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End
    
    GM_NEXT_PREV = 2
    
    Call frm117_report_gdn_grn_header
    Call frm117_report_gdn_grn

    MsgBox "Data GRN telah berjaya dipadamkan.", vbInformation, "Info"

End If
End Sub
Sub frm117_cetak_statement()
'on error resume next
Dim TM As Date
Dim TA As Date

If frm117.L5_Text = 1 Then '0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh

    TM = frm117.L6_Text 'Tarikh mula
    TA = frm117.L7_Text 'Tarikh akhir

End If

If frm117.L8_Text = "Semua supplier dan agen" Then
    
    frm117_LM_SEARCH_1 = Null
    frm117_LM_SEARCH_1_LOGIC = "<>"
    
Else
    
    frm117_LM_SEARCH_1 = frm117.L8_Text
    frm117_LM_SEARCH_1_LOGIC = "="
    
End If

If frm117.L9_Text = "Semua GDN/GRN/INV/VOU" Then
    
    frm117_LM_SEARCH_2 = Null
    frm117_LM_SEARCH_2_LOGIC = "<>"
    
Else
    
    frm117_LM_SEARCH_2 = frm117.L9_Text
    frm117_LM_SEARCH_2_LOGIC = "="
    
End If

'Set rs = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from default_setting where default1='" & "default" & "'", cn, adOpenKeyset, adLockOptimistic

'If Not rs.EOF Then
    
'    If Not IsNull(rs!default_printer) Then LM_PRINTER = rs!default_printer
    
'End If

'rs.Close
'Set rs = Nothing

'For Each oPrn In Printers
'    If oPrn.DeviceName = LM_PRINTER Then
'        Set Printer = oPrn
'        PRINTER_FOUND = 1 '0 : Printer Not Found , 1 : Printer Found
'        Exit For
'    End If
'Next

Report79.Sections("Section5").Controls("L2").Caption = Now
Report79.Sections("Section5").Controls("L3").Caption = Format(0, "#,##0.00")
Report79.Sections("Section5").Controls("L4").Caption = Format(0, "#,##0.00")
Report79.Sections("Section5").Controls("L5").Caption = Format(0, "#,##0.00")
Report79.Sections("Section5").Controls("L6").Caption = Format(0, "#,##0.00")
Report79.Sections("Section5").Controls("L7").Caption = Format(0, "#,##0.00")
Report79.Sections("Section5").Controls("L8").Caption = Format(0, "#,##0.00")
Report79.Sections("Section5").Controls("L9").Caption = "RM " & Format(0, "#,##0.00")

Report79.Sections("Section4").Controls("L10").Caption = "Semua GDN/GRN/INV/VOU"
Report79.Sections("Section4").Controls("L11").Caption = "Semua supplier dan agen"
Report79.Sections("Section4").Controls("L12").Caption = "-"

If frm117.L5_Text = "1" Then '0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh

    Report79.Sections("Section4").Controls("L12").Caption = frm117.L6_Text & " hingga " & frm117.L7_Text
    
End If

If frm117.L8_Text <> "Semua supplier dan agen" Then
    
    Report79.Sections("Section4").Controls("L11").Caption = frm117.L8_Text

End If

If frm117.L9_Text <> "Semua GDN/GRN/INV/VOU" Then
    
    Report79.Sections("Section4").Controls("L10").Caption = frm117.L9_Text
    
End If

'### Reset maklumat kedai ### - Start
Report79.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report79.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report79.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report79.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report79.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

Report79.Sections("Section4").Controls("L205").Caption = "Goods Received Note"

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report79.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report79.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report79.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report79.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report79.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Report79.Sections("Section5").Controls("L2").Caption = Now
Report79.Sections("Section5").Controls("L3").Caption = Format(frm117.L20_Text, "#,##0.00")
Report79.Sections("Section5").Controls("L4").Caption = Format(frm117.L21_Text, "#,##0.00")
Report79.Sections("Section5").Controls("L5").Caption = Format(frm117.L22_Text, "#,##0.00")
Report79.Sections("Section5").Controls("L6").Caption = Format(frm117.L23_Text, "#,##0.00")
Report79.Sections("Section5").Controls("L7").Caption = Format(frm117.L24_Text, "#,##0.00")
Report79.Sections("Section5").Controls("L8").Caption = Format(frm117.L26_Text, "#,##0.00")
Report79.Sections("Section5").Controls("L9").Caption = "RM " & Format(frm117.L27_Text, "#,##0.00")

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm117.L5_Text = 0 Then rs.Open "select * from 77_gdn_grn where supplier_agen " & frm117_LM_SEARCH_1_LOGIC & "'" & frm117_LM_SEARCH_1 & "' AND status = 1 AND jenis " & frm117_LM_SEARCH_2_LOGIC & "'" & frm117_LM_SEARCH_2 & "' order by write_timestamp ASC", cn, adOpenKeyset, adLockOptimistic
If frm117.L5_Text = 1 Then rs.Open "select * from 77_gdn_grn where supplier_agen " & frm117_LM_SEARCH_1_LOGIC & "'" & frm117_LM_SEARCH_1 & "' AND status = 1 AND jenis " & frm117_LM_SEARCH_2_LOGIC & "'" & frm117_LM_SEARCH_2 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by write_timestamp ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report79.DataSource = rs
    Report79.Show
    
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing
End Sub
Sub frm117_padam_gdn()
'on error resume next
Dim rs2 As ADODB.Recordset
Dim frm117_LM_BERAT_ASAL As Double
Dim frm117_LM_BEZA_BERAT As Double
Dim frm117_LM_BERAT_JUALAN As Double
Dim frm117_LM_BERAT_ASAL_COMP As Double
Dim frm117_LM_BERAT_SELEPAS_COMP As Double
Dim frm117_SUSUT_BERAT As Double

LM_NOW = Now
LM_TARIKH = DateTime.Date$
LM_MASA = DateTime.Time$

'### Padam data voucher / invoice bagi belian agen ini ### - Start

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 77_gdn_grn where no_rujukan='" & G_No_RESIT_JUALAN & "' AND jenis_urusan = 0 AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    G_ID = rs!ID
    Call recovery_77_gdn_grn

    rs!Status = 0
    rs!terminal = G_TERMINAL
    rs!user = MDI_frm1.L3_Text 'Nama Pekerja
    rs.Update
    DATA_SAVE = 1

End If

rs.Close
Set rs = Nothing
'### Padam data voucher / invoice bagi belian agen ini ### - End

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 23_senarai_jualan where no_resit='" & G_No_RESIT_JUALAN & "' AND status_rekod = 1", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    frm117_LM_BERAT_ASAL = 0
    frm117_LM_BEZA_BERAT = 0
    frm117_LM_BERAT_JUALAN = 0
    frm117_SUSUT_BERAT = 0
    frm117_LM_BERAT_ASAL_COMP = 0
    frm117_LM_BERAT_SELEPAS_COMP = 0
    
    Set rs2 = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs2.Open "select * from data_database where no_siri_Produk='" & rs!no_siri_Produk & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs2.EOF Then
    
        If rs2!receiving_Status = 0 Or rs2!receiving_Status = 2 Then
            If Not IsNull(rs2!Berat) Then frm117_LM_BERAT_ASAL = rs2!Berat
            If Not IsNull(rs2!beza_berat) Then frm117_LM_BEZA_BERAT = rs2!beza_berat
            If Not IsNull(rs!berat_jualan) Then frm117_LM_BERAT_JUALAN = rs!berat_jualan
            If Not IsNull(rs2!susut_berat) Then frm117_SUSUT_BERAT = Format(rs2!susut_berat, "0.00") 'Susut berat
            
            frm117_LM_BERAT_ASAL_COMP = Format(frm117_LM_BERAT_ASAL, "0.00")
            frm117_LM_BERAT_SELEPAS_COMP = Format(frm117_LM_BERAT_JUALAN + frm117_LM_BEZA_BERAT - frm117_SUSUT_BERAT, "0.00")
            
            If Format(frm117_LM_BERAT_ASAL, "0.00") = Format(frm117_LM_BERAT_JUALAN + frm117_LM_BEZA_BERAT, "0.00") Then
                rs2!beza_berat = Format(frm117_LM_BERAT_JUALAN + frm117_LM_BEZA_BERAT, "0.00")
                rs2!StatusItem = 10
            Else
                rs2!beza_berat = Format(frm117_LM_BERAT_JUALAN + frm117_LM_BEZA_BERAT, "0.00")
                rs2!StatusItem = 12
                rs2!tarikh_jualan1 = DateTime.Date
            End If
        Else
            rs2!StatusItem = 10
        End If
        rs2.Update
    End If
    
    rs2.Close
    Set rs2 = Nothing

    rs!status_rekod = 0
    rs.Update
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'#### Update Log Aktiviti Sistem #### - Start
user = MDI_frm1.L3_Text
LogAct_Memory = "[" & user & "] Padam data GDN. No. Rujukan [" & G_No_RESIT_JUALAN & "]."
LogDate_Memory = LM_NOW
Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End

MsgBox "Data GDN telah berjaya dipadamkan.", vbInformation, "Info"

GM_NEXT_PREV = 2

Call frm117_report_gdn_grn_header
Call frm117_report_gdn_grn
End Sub
Sub frm117_edit_data_inv_vou()
'On Error Resume Next
frm118_LM_USER = vbNullString
LM_FOUND = 0

Call Frm118_background_color
Call frm118_initial_setting

frm118.CB1.Enabled = False
frm118.CB2.Enabled = False

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 77_gdn_grn where no_rujukan='" & G_No_RESIT_JUALAN & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    frm118.L1_Text = G_No_RESIT_JUALAN
    If Not IsNull(rs!tarikh) Then frm118.DTPicker1 = rs!tarikh
    If Not IsNull(rs!jenis_gst) Then
        If rs!jenis_gst = 0 Then
            frm118.CB6 = 1
            frm118.CB7 = 0
        ElseIf rs!jenis_gst = 1 Then
            frm118.CB6 = 0
            frm118.CB7 = 1
        End If
    End If
    If Not IsNull(rs!jenis_urusan) Then
    
        If rs!jenis_urusan = 2 Then
            frm118.CB1 = 1
            If Not IsNull(rs!harga_dengan_gst_grn) Then frm118.TB6 = Format(rs!harga_dengan_gst_grn, "#,##0.00")
        End If
        If rs!jenis_urusan = 3 Then
            frm118.CB2 = 1
            If Not IsNull(rs!harga_dengan_gst) Then frm118.TB6 = Format(rs!harga_dengan_gst, "#,##0.00")
        End If
        
    End If
    
    If Not IsNull(rs!kadar_gst) Then frm118.L2_Text = rs!kadar_gst
    If Not IsNull(rs!jumlah) Then frm118.TB2 = Format(rs!jumlah, "#,##0.00")
    If Not IsNull(rs!gst_sr_cukai) Then frm118.TB3 = Format(rs!gst_sr_cukai, "#,##0.00")
    If Not IsNull(rs!gst_sr_harga) Then frm118.L3_Text = Format(rs!gst_sr_harga, "#,##0.00")
    If Not IsNull(rs!gst_zr_harga) Then frm118.TB5 = Format(rs!gst_zr_harga, "#,##0.00")
    If Not IsNull(rs!no_rujukan_supplier) Then frm118.TB1 = rs!no_rujukan_supplier
    If Not IsNull(rs!cara_bayaran) Then
        If rs!cara_bayaran = 0 Then
            frm118.CB3 = 1
            frm118.CB4 = 0
            frm118.CB5 = 0
        ElseIf rs!cara_bayaran = 1 Then
            frm118.CB3 = 0
            frm118.CB4 = 1
            frm118.CB5 = 0
        ElseIf rs!cara_bayaran = 2 Then
            frm118.CB3 = 0
            frm118.CB4 = 0
            frm118.CB5 = 1
        End If
    End If
    If Not IsNull(rs!umum_berat) Then
        If rs!umum_berat = 0 Then
        
            frm118.CB8 = 1
            frm118.CB9 = 0
        
        ElseIf rs!umum_berat = 1 Then
        
            frm118.CB8 = 0
            frm118.CB9 = 1
            
            If Not IsNull(rs!nilaian_harga_emas) Then frm118.TB6 = Format(rs!nilaian_harga_emas, "#,##0.00")
            If Not IsNull(rs!harga_999) Then frm118.TB7 = Format(rs!harga_999, "#,##0.00")
            If Not IsNull(rs!berat_tukaran) Then frm118.TB8 = Format(rs!berat_tukaran, "#,##0.00")
            
        End If
    
    End If

    If Not IsNull(rs!supplier_agen) Then
        'on error goto Err_A:
        Frm118_LM_SUPPLIER = rs!supplier_agen
        frm118.CBB1 = Frm118_LM_SUPPLIER
    
Restore_A:

    End If
    
    If Not IsNull(rs!user) Then

        frm118_LM_USER = rs!user

    End If
    'on error resume next
    LM_FOUND = 1
    
End If

rs.Close
Set rs = Nothing
'### Masukkan data voucher / invoice bagi belian agen ini ### - End

'### Carian Maklumat Penjual (Data Pekerja) ### - Start
If frm118_LM_USER <> vbNullString Then

    DATA_PEKERJA_FOUND = 0
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where Samaran='" & frm118_LM_USER & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        Frm118_LM_MAKLUMAT_PEKERJA = rs!Samaran & "  |  " & rs!NoPekerja
        DATA_PEKERJA_FOUND = 1
        
    End If
    
    rs.Close
    Set rs = Nothing

    If DATA_PEKERJA_FOUND = 1 Then
        'On Error GoTo Err_B:
        frm118.CBB2 = Frm118_LM_MAKLUMAT_PEKERJA
        
Restore_B:
    End If
    
    'on error resume next
End If
'### Carian Maklumat Penjual (Data Pekerja) ### - End

If LM_FOUND = 1 Then

    frm118.CBB2.Enabled = True
    frm118.CBB2.BackColor = &HFFFFFF
        
    frm118.CMD1.Visible = False
    frm118.CMD2.Visible = True
    frm118.CMD3.Visible = True
    
    frm118.Show
    frm117.Hide
End If

Exit Sub

Err_A:

frm118.CBB1.AddItem Frm118_LM_SUPPLIER
frm118.CBB1 = Frm118_LM_SUPPLIER
            
Resume Restore_A:

Exit Sub
Err_B:
frm118.CBB2.AddItem Frm118_LM_MAKLUMAT_PEKERJA
frm118.CBB2 = Frm118_LM_MAKLUMAT_PEKERJA
Resume Restore_B:
End Sub
Sub frm118_cetak_inv_vou()
'on error resume next
Frm115_LM_CUST = vbNullString

'Set rs = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from default_setting where default1='" & "default" & "'", cn, adOpenKeyset, adLockOptimistic

'If Not rs.EOF Then
    
'    If Not IsNull(rs!default_printer) Then LM_PRINTER = rs!default_printer
    
'End If

'rs.Close
'Set rs = Nothing

'For Each oPrn In Printers
'    If oPrn.DeviceName = LM_PRINTER Then
'        Set Printer = oPrn
'        PRINTER_FOUND = 1 '0 : Printer Not Found , 1 : Printer Found
'        Exit For
'    End If
'Next

Report80.Sections("Section4").Controls("L1").Caption = vbNullString 'No. Rujukan
Report80.Sections("Section4").Controls("L2").Caption = vbNullString 'Tarikh
Report80.Sections("Section4").Controls("L3").Caption = vbNullString 'Nama Pembeli
Report80.Sections("Section4").Controls("L4").Caption = vbNullString 'No. Telefon
Report80.Sections("Section4").Controls("L17").Caption = vbNullString 'Jurujual
Report80.Sections("Section4").Controls("L18").Caption = "-" 'No. ID GST
Report80.Sections("Section4").Controls("L21").Caption = vbNullString 'No. Rujukan Dari Supplier

Report80.Sections("Section2").Controls("L5").Caption = "0.00" 'Jumlah Cukai GST
Report80.Sections("Section2").Controls("L6").Caption = "0.00" 'Jumlah Bayaran
Report80.Sections("Section2").Controls("L8").Caption = "0.00" 'Jumlah Bayaran

'### Reset maklumat kedai ### - Start
Report80.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report80.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report80.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report80.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report80.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
Report80.Sections("Section2").Controls("L20").Caption = vbNullString 'Tujuan
'### Reset maklumat kedai ### - End

Report80.Sections("Section4").Controls("L205").Caption = "Invoice / Voucher"

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report80.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report80.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report80.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report80.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report80.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Report80.Sections("Section4").Controls("L1").Caption = G_No_RESIT_JUALAN 'No. Invoice

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 77_gdn_grn where no_rujukan='" & G_No_RESIT_JUALAN & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Not IsNull(rs!tarikh) Then Report80.Sections("Section4").Controls("L2").Caption = rs!tarikh 'Tarikh
    If Not IsNull(rs!user) Then Report80.Sections("Section4").Controls("L17").Caption = rs!user 'Jurujual
    If Not IsNull(rs!bil_barang) Then Report80.Sections("Section5").Controls("L15").Caption = rs!bil_barang 'Bilangan barang
    If Not IsNull(rs!Berat_Asal) Then Report80.Sections("Section5").Controls("L16").Caption = Format(rs!Berat_Asal, "#,##0.00 g") 'Berat Asal (g)
    If Not IsNull(rs!kadar_tukaran) Then Report80.Sections("Section5").Controls("L19").Caption = rs!kadar_tukaran 'Mutu
    If Not IsNull(rs!berat_tukaran_grn) Then Report80.Sections("Section5").Controls("L20").Caption = Format(rs!berat_tukaran_grn, "#,##0.00 g") 'Berat 999.9 (g)
    If Not IsNull(rs!supplier_agen) Then Frm115_LM_CUST = rs!supplier_agen
    If Not IsNull(rs!no_rujukan_supplier) Then Report80.Sections("Section4").Controls("L21").Caption = "No. Invoice Supplier          : " & rs!no_rujukan_supplier 'No. Rujukan Dari Supplier
    
    If rs!jenis = "INV" Then
        Report80.Sections("Section4").Controls("L205").Caption = "Invoice"
        If Not IsNull(rs!harga_dengan_gst_grn) Then Report80.Sections("Section2").Controls("L6").Caption = Format(rs!harga_dengan_gst_grn, "#,##0.00")
        If Not IsNull(rs!harga_dengan_gst_grn) Then Report80.Sections("Section2").Controls("L8").Caption = "Jumlah : RM " & Format(rs!harga_dengan_gst_grn, "#,##0.00")
    ElseIf rs!jenis = "VOU" Then
        Report80.Sections("Section4").Controls("L205").Caption = "Voucher"
        If Not IsNull(rs!harga_dengan_gst) Then Report80.Sections("Section2").Controls("L6").Caption = Format(rs!harga_dengan_gst, "#,##0.00")
        If Not IsNull(rs!harga_dengan_gst) Then Report80.Sections("Section2").Controls("L8").Caption = "Jumlah : RM " & Format(rs!harga_dengan_gst, "#,##0.00")
    End If
    
    If Not IsNull(rs!jumlah_gst) Then Report80.Sections("Section2").Controls("L5").Caption = Format(rs!jumlah_gst, "#,##0.00") 'Jumlah GST
    If rs!umum_berat = 0 Then
        Report80.Sections("Section2").Controls("L20").Caption = "Tujuan : Bayaran belian stok/perkhidmatan lain-lain." 'Tujuan
    ElseIf rs!umum_berat = 1 Then
        
        If Not IsNull(rs!berat_tukaran) Then LM_BERAT = rs!berat_tukaran
        If Not IsNull(rs!harga_999) Then LM_HARGA_999 = rs!harga_999
        
        Report80.Sections("Section2").Controls("L20").Caption = "Tujuan : Bayaran belian stok/perkhidmatan lain-lain. (Berat : " & Format(LM_BERAT, "#,##0.00 g") & " @ RM " & Format(LM_HARGA_999, "#,##0.00") & " /g)" 'Tujuan
    End If
    
End If

rs.Close
Set rs = Nothing

If Frm115_LM_CUST <> vbNullString Then
 
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from setting_database where Supplier='" & Frm115_LM_CUST & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        
        If Not IsNull(rs!supplier) Then Report80.Sections("Section4").Controls("L3").Caption = rs!supplier 'Nama Pembeli
        If Not IsNull(rs!no_tel_hp) Then Report80.Sections("Section4").Controls("L4").Caption = rs!no_tel_hp 'No. Telefon
        If Not IsNull(rs!no_id_gst) Then Report80.Sections("Section4").Controls("L18").Caption = rs!no_id_gst 'No. ID GST

    End If
    
    rs.Close
    Set rs = Nothing
    
End If
   
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 77_gdn_grn where no_rujukan='" & G_No_RESIT_JUALAN & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report80.DataSource = rs
    If G_PREVIEW = 1 Then Report80.Show
    
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing

If G_PREVIEW = 0 Then Report80.PrintReport
End Sub
Sub frm117_padam_inv_vou()
'On Error Resume Next
DATA_SAVE = 0

LM_NOW = Now
LM_TARIKH = DateTime.Date$
LM_MASA = DateTime.Time$
LM_NO_RUJUKAN = vbNullString

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 77_gdn_grn where ID='" & G_ID & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    G_ID = rs!ID
    Call recovery_77_gdn_grn
    
    If Not IsNull(rs!jenis) Then LM_JENIS = rs!jenis
    rs!Status = 0
    'rs!jenis_urusan = 1
    rs!terminal = G_TERMINAL
    rs!user = MDI_frm1.L3_Text 'Nama Pekerja
    rs.Update
    DATA_SAVE = 1
    
End If

rs.Close
Set rs = Nothing
'### Masukkan data voucher / invoice bagi belian agen ini ### - End

'### Masukkan Data Jualan Ke Dalam Table Jualan ### - Start

If DATA_SAVE = 1 Then
    
    If LM_JENIS = "INV" Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 22_jualan where no_resit='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            G_ID = rs!ID
            Call recovery_22_jualan
            
            rs!Status = 0
            rs.Update
            
        End If
        
        rs.Close
        Set rs = Nothing
                    
    End If
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 39_akaun_expense where no_rujukan_expense='" & G_No_RESIT_JUALAN & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        
        rs!Status = 0
        rs.Update
        
    End If
    
    rs.Close
    Set rs = Nothing

'#### Update Log Aktiviti Sistem #### - Start
    'User = MDI_frm1.L3_Text
    LogAct_Memory = "[" & MDI_frm1.L3_Text & "] Padam data " & LM_JENIS & " dari/kepada agen/supplier. No. Rujukan [" & G_No_RESIT_JUALAN & "]."
    LogDate_Memory = LM_NOW
    Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End
    
    GM_NEXT_PREV = 2
    
    Call frm117_report_gdn_grn_header
    Call frm117_report_gdn_grn

    MsgBox "Data " & LM_JENIS & " telah berjaya dipadamkan.", vbInformation, "Info"

End If
End Sub

