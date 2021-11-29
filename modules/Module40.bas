Attribute VB_Name = "Module40"
Sub Frm83_Initial_Setting()
'on error resume next
Frm83.Pic2.Left = 120
Frm83.Pic2.Top = 9840

Frm83.L3_Text.BackStyle = 0
Frm83.L8_Text.BackStyle = 0
Frm83.L10_Text.BackStyle = 0
Frm83.L14_Text.BackStyle = 0

Frm83.TB2 = vbNullString
Frm83.TB3 = vbNullString
Frm83.TB4 = 0
Frm83.TB6 = vbNullString
Frm83.TB7 = vbNullString
Frm83.TB8 = vbNullString
Frm83.TB9 = vbNullString
Frm83.TB10 = vbNullString
Frm83.TB12 = vbNullString
Frm83.TB13 = vbNullString
Frm83.TB14 = vbNullString
Frm83.TB15 = vbNullString
Frm83.TB16 = vbNullString
Frm83.TB19 = "0.00"
Frm83.TB20 = vbNullString
Frm83.TB21 = vbNullString
Frm83.TB22 = vbNullString
Frm83.TB24 = "0.00"
Frm83.TB25 = "0.00"
Frm83.TB26 = "0.00"
Frm83.TB31 = "0.00"
Frm83.TB32 = "0.00"
Frm83.TB33 = "0.00"
Frm83.TB27 = "0.00"
Frm83.TB29 = "0.00"
Frm83.L10_Text = 0
Frm83.TB21 = "0.00"
Frm83.TB22 = "0.00"
Frm83.TB34 = vbNullString
Frm83.CB14 = 1
Frm83.CB15 = 0
Frm83.TB35 = "0.00"
Frm83.TB36 = vbNullString
Frm83.TB37 = vbNullString

Frm83.L11_Text = "0.00"
Frm83.L22_Text = "0.00"
Frm83.L23_Text = "0.00"
Frm83.L24_Text = "0.00"
Frm83.L25_Text = "0.00"
Frm83.L26_Text = "0.00"
Frm83.L30_Text = 0
Frm83.L31_Text = "Tiada"
Frm83.L32_Text = 0
Frm83.L36_Text = vbNullString
Frm83.L37_Text = vbNullString
Frm83.L39_Text = 1
Frm83.DTPicker1 = DateTime.Date

Frm83.TB40 = "0.00"
Frm83.TB41 = "0.00"
Frm83.TB42 = "0.00"

Frm83.CBB1.Clear
Frm83.CBB2.Clear
Frm83.CBB3.Clear
Frm83.CBB5.Clear
Frm83.CBB6.Clear

'GoTo aaaa:

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where status='" & 1 & "' order by supplier ASC , Metal_Purity ASC , kategori_Produk ASC , SenaraiDulang ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!supplier) Then
        
        If rs!jenis_supplier = "Supplier" Then
            Frm83.CBB1.AddItem rs!supplier
        End If
        
    End If
    If Not IsNull(rs!Metal_Purity) Then Frm83.CBB2.AddItem rs!Metal_Purity
    If Not IsNull(rs!kategori_Produk) Then Frm83.CBB3.AddItem rs!kategori_Produk
    If Not IsNull(rs!SenaraiDulang) Then Frm83.CBB5.AddItem rs!SenaraiDulang
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm83.L8_Text = G_RATE_GST
Frm83.L30_Text = G_RIYAL

If G_UPAH_SUPPLIER = 0 Then
    Frm83.CB14 = 1
    Frm83.CB15 = 0
ElseIf G_UPAH_SUPPLIER = 1 Then
    Frm83.CB14 = 0
    Frm83.CB15 = 1
End If
If Frm83.CB8 = 1 Then
    Frm83.TB19 = Format(G_SPREAD, "#,##0.00") 'Spread Trade In %
ElseIf Frm83.CB7 = 1 Then
    Frm83.TB19 = "0.00" 'Spread Trade In %
End If
If G_PRINT_BARCODE = 0 Then
    Frm83.CB13 = 0
Else
    Frm83.CB13 = 1
End If

'###Senarai Nama Pekerja###
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where status='" & "Aktif" & "' AND InvestorSmall = 0 AND InvestorBig = 0", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    If rs!user_level = 1 Or rs!user_level = 2 Or rs!user_level = 3 Or rs!user_level = G_LEVEL_USER Then
        Frm83.CBB6.AddItem rs!Samaran & "  |  " & rs!NoPekerja
        xx = 1
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'###Padam Temp Table### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "TRUNCATE TABLE " & G_BELIAN_TEMP & ""

Set rs = cn.Execute(strsql)
Set rs = Nothing
'###Padam Temp Table### - End

'aaaa:
End Sub
Sub Frm83_initial_setting2()
'on error resume next

'==================================================================
'Module ini digunakan untuk reset maklumat supplier dan gst
'juga reset maklumat penerimaan stok bagi belian buyback / trade in
'==================================================================

Frm83.TB1 = vbNullString
Frm83.TB28 = vbNullString

'Call frm83_flag_barang_baru

'Frm83.CBB1.Clear

'Set rs = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from setting_database where jenis_supplier='" & "Supplier" & "'", cn, adOpenKeyset, adLockOptimistic

'While rs.EOF = False
'    If Not IsNull(rs!supplier) Then Frm83.CBB1.AddItem rs!supplier
'    rs.MoveNext
'Wend

'rs.Close
'Set rs = Nothing

If G_GST_INCOMING = 1 Then
    Frm83.CB2 = 0
    Frm83.CB3 = 1
    Frm83.CB11 = 1
Else
    Frm83.CB2 = 1
    Frm83.CB3 = 0
    Frm83.CB11 = 0
End If

'Set rs = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

'If Not rs.EOF Then
'    If rs!Default1 = "Default" Then
'        If Not IsNull(rs!gst_arinashi_belian) Then
'            If rs!gst_arinashi_belian = 1 Then
'                Frm83.CB2 = 0
'                Frm83.CB3 = 1
'                Frm83.CB11 = 1
'            Else
'                Frm83.CB2 = 1
'                Frm83.CB3 = 0
'                Frm83.CB11 = 0
'            End If
'        End If
'    End If
'End If

'rs.Close
'Set rs = Nothing
End Sub
Sub Frm83_Reset_Form()
'on error resume next
'Frm83.TB11 = vbNullString
'Frm83.TB22 = "0.00"
Frm83.L31_Text = "Tiada"
Frm83.L32_Text = 0
If Frm83.CB4 = 1 Then Frm83.TB8 = "0.00"
Frm83_LM_No_RESIT = Frm83.L12_Text 'No. Resit

Frm83.L69_Text = -1 'Titik Pencarian Data
Frm83.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Frm83.L67_Text = 0 'Paparan Page ke-xxx
Frm83.L68_Text = 0

Exit Sub

If Frm83.L21_Text = 0 Then
Re_gen_no_resit:
    
    '###Carian No. Resit###
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & Format(Frm83_LM_No_RESIT, "000000") & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        Frm83_LM_No_RESIT = Frm83_LM_No_RESIT + 1
        
        rs.Close
        Set rs = Nothing
        
        GoTo Re_gen_no_resit:
    End If
    
    rs.Close
    Set rs = Nothing
    
    Frm83.L12_Text = Frm83_LM_No_RESIT 'No. Resit
End If
Exit Sub

'Set rs = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from 14_gold_bar_tetapan", cn, adOpenKeyset, adLockOptimistic

'If Not rs.EOF Then
'    If rs!default_setting = "Default" Then
'        If Not IsNull(rs!no_rujukan_belian) Then Frm83.L9_Text = rs!no_rujukan_belian 'No. Rujukan Belian
'        If Not IsNull(rs!siri_barcode) Then
'            Frm83.TB7 = Format(rs!siri_barcode, "00000") 'No. Siri Barcode
'            Frm83.L3_Text = rs!siri_barcode 'No. Siri Barcode
'        End If
'    End If
'End If

'rs.Close
'Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Default1 = "Default" Then
        If Not IsNull(rs!NoRujukanSistem) Then Frm83.L9_Text = rs!NoRujukanSistem 'No. Rujukan Sistem
        If Not IsNull(rs!tarikh) Then Frm83.DTPicker1 = rs!tarikh
        If Not IsNull(rs!gst_value) Then Frm83.L8_Text = rs!gst_value
        If Not IsNull(rs!NoRujukanStock) Then
            If Frm83.CB9 = 1 Then
                'Frm83.TB7 = Format(rs!NoRujukanStock, "000000") 'No. Siri Barcode
                'Frm83.L3_Text = rs!NoRujukanStock 'No. Siri Barcode
            ElseIf Frm83.CB10 = 1 Then
                'Frm83.TB7 = Format(rs!no_siri_gb, "000000") & "W" 'No. Siri Barcode
                'Frm83.L3_Text = rs!no_siri_gb 'No. Siri Barcode
            End If
        End If
        'If Frm83.CB8 = 1 Then
            If Not IsNull(rs!ResitNo) Then Frm83.L12_Text = rs!ResitNo 'No. Resit
        'End If
        If Frm83.CB8 = 1 Then
            If Not IsNull(rs!spread_Cash_Trade_In) Then Frm83.TB19 = Format(rs!spread_Cash_Trade_In, "0.00") 'Spread Trade In %
        ElseIf Frm83.CB7 = 1 Then
            Frm83.TB19 = "0.00" 'Spread Trade In %
        End If
        If Not IsNull(rs!gst_arinashi_belian) Then
            If rs!gst_arinashi_belian = 1 Then
                Frm83.CB2 = 0
                Frm83.CB3 = 1
                Frm83.CB11 = 1
            Else
                Frm83.CB2 = 1
                Frm83.CB3 = 0
                Frm83.CB11 = 0
            End If
        End If
    End If
End If

rs.Close
Set rs = Nothing
End Sub
Sub Frm83_Senarai_Belian_Header()
'on error resume next

With Frm83.ListView2
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear

    Frm83.ListView2.ListItems.Clear
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "Tarikh Belian", 2000
    .ColumnHeaders.Add 5, , "No. Invoice", 1500
    .ColumnHeaders.Add 6, , "Supplier", 2500
    .ColumnHeaders.Add 7, , "Purity", 1500, 2
    .ColumnHeaders.Add 8, , "Kategori Produk", 3500
    .ColumnHeaders.Add 9, , "No. Siri Produk", 2000
    .ColumnHeaders.Add 10, , "Berat (g)", 1200, 1
    .ColumnHeaders.Add 11, , "Harga Per Gram (RM/g)", 2500, 1
    .ColumnHeaders.Add 12, , "Upah (RM)", 1200, 1
    .ColumnHeaders.Add 13, , "Spread (%)", 1400, 1
    .ColumnHeaders.Add 14, , "Harga Lepas Spread (RM)", 2700, 1
    .ColumnHeaders.Add 15, , "Adjustment (RM)", 2000, 1
    .ColumnHeaders.Add 16, , "Harga Belian Tanpa GST (RM)", 3100, 1
    .ColumnHeaders.Add 17, , "Dulang", 1500, 2
    .ColumnHeaders.Add 18, , "Panjang", 1500
    .ColumnHeaders.Add 19, , "Lebar", 1500
    .ColumnHeaders.Add 20, , "Saiz", 1500
    .ColumnHeaders.Add 21, , "Jenis Cukai GST", 2200, 1
    .ColumnHeaders.Add 22, , "Jumlah GST (RM)", 2200, 1
    .ColumnHeaders.Add 23, , "Harga Belian Dengan GST (RM)", 3200, 1
    .ColumnHeaders.Add 24, , "Code 1", 1000
    .ColumnHeaders.Add 25, , "Code 2", 1000
    
End With

'No.
'No.
'ID
'Tarikh Belian
'No. Invoice
'Supplier
'Purity
'Kategori Produk
'No. Siri Produk
'Berat (g)
'Harga Per Gram (RM/g)
'Upah (RM)
'Spread (%)
'Harga Lepas Spread (RM)
'Adjustment (RM)
'Harga Belian Tanpa GST (RM)
'Dulang
'Panjang
'Lebar
'Saiz
'Jenis Cukai GST
'Jumlah GST (RM)
'Harga Belian Dengan GST (RM)
End Sub
Sub Frm83_Senarai_Belian()
'on error resume next
Dim Frm83_LM_HARGA_TANPA_GST As Double 'Harga Belian Tanpa Cukai GST
Dim Frm83_LM_HARGA_DENGAN_GST As Double 'Harga Belian Tanpa Cukai GST
Dim Frm83_LM_GST_SR As Double 'Kutipan GST : SR
Dim Frm83_LM_GST_ZR As Double 'Kutipan GST : ZR
Dim Frm83_LM_JUMLAH_HARGA_SR As Double 'Total Harga Yang Dikenakan GST SR
Dim Frm83_LM_JUMLAH_HARGA_ZR As Double 'Total Harga Yang Dikenakan GST ZR

Dim frm83_LM_TOTAL_PAGE As Double

frm83_PAGE_SIZE = 27
frm83_LM_TOTAL_PAGE = 0
x = 0

re_gen_report:

LM_START_ROW = Frm83.L69_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + frm83_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm83.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - frm83_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        Frm83.L67_Text = 1
    End If
End If

frm83_LM_PAGE_FOUND = 0

Frm83_LM_HARGA_TANPA_GST = 0
Frm83_LM_HARGA_DENGAN_GST = 0
Frm83_LM_GST_SR = 0
Frm83_LM_GST_ZR = 0
Frm83_LM_JUMLAH_HARGA_SR = 0
Frm83_LM_JUMLAH_HARGA_ZR = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from " & G_BELIAN_TEMP & " where StatusItem<>'" & "0" & "' LIMIT " & LM_START_ROW & "," & frm83_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    'If rs!StatusItem = "10" Or rs!StatusItem = "3" Or rs!StatusItem = "4" Or rs!StatusItem = "1" Then
    If rs!StatusItem <> "5" Then

        x = x + 1
        If frm83_LM_PAGE_FOUND = 0 Then
            If Frm83.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
                If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                    Frm83.L67_Text = Frm83.L67_Text + 1 'Paparan Page ke-xxx
                    frm83_LM_PAGE_FOUND = 1
                ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                    If IsNumeric(Frm83.L67_Text) Then
                        If Frm83.L67_Text <> 1 Then
                            Frm83.L67_Text = Frm83.L67_Text - 1 'Paparan Page ke-xxx
                            frm83_LM_PAGE_FOUND = 1
                        End If
                    End If
                End If
            End If
        End If
    
        Y = ((Frm83.L67_Text - 1) * frm83_PAGE_SIZE) + x

        With Frm83.ListView2.ListItems.Add(, , rs!ID)
        
            .ListSubItems.Add , , Y
            
            If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
            
            If Not IsNull(rs!tarikh_belian) Then 'Tarikh Belian
                .ListSubItems.Add , , rs!tarikh_belian
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!bill_No_Belian) Then 'No. Invoice
                .ListSubItems.Add , , rs!bill_No_Belian
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!nama_Supplier) Then 'Supplier
                .ListSubItems.Add , , rs!nama_Supplier
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!purity) Then 'Purity
                .ListSubItems.Add , , rs!purity
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!kategori_Produk) Then 'Kategori Produk
                .ListSubItems.Add , , rs!kategori_Produk
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!no_siri_Produk) Then 'No. Siri Produk
                .ListSubItems.Add , , rs!no_siri_Produk
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!Berat) Then 'Berat (g)
                .ListSubItems.Add , , Format(rs!Berat, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!kos_Belian_Gram) Then 'Harga Per Gram (RM/g)
                .ListSubItems.Add , , Format(rs!kos_Belian_Gram, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!UPAH) Then 'Upah (RM)
                .ListSubItems.Add , , Format(rs!UPAH, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!Spread) Then 'Spread (%)
                .ListSubItems.Add , , Format(rs!Spread, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!harga_lepas_spread) Then 'Harga Lepas Spread (RM)
                .ListSubItems.Add , , Format(rs!harga_lepas_spread, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!adjustment) Then 'Adjustment (RM)
                .ListSubItems.Add , , Format(rs!adjustment, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If

            If Not IsNull(rs!harga_tanpa_gst) Then
                
                If Not IsNull(rs!harga_tanpa_gst) Then 'Harga Belian Tanpa GST (RM)
                    .ListSubItems.Add , , Format(rs!harga_tanpa_gst, "#,##0.00")
                Else
                    .ListSubItems.Add , , ""
                End If
            
                If IsNumeric(rs!harga_tanpa_gst) Then
                    Frm83_LM_HARGA_TANPA_GST = Frm83_LM_HARGA_TANPA_GST + rs!harga_tanpa_gst 'Harga Belian Tanpa GST (RM)
                    
                    If Not IsNull(rs!gst_ari_nashi) Then
                        If rs!gst_ari_nashi = 0 Then
                            Frm83_LM_JUMLAH_HARGA_ZR = Frm83_LM_JUMLAH_HARGA_ZR + rs!harga_tanpa_gst 'Total Harga Yang Dikenakan GST ZR
                        ElseIf rs!gst_ari_nashi = 1 Then
                            Frm83_LM_JUMLAH_HARGA_SR = Frm83_LM_JUMLAH_HARGA_SR + rs!harga_tanpa_gst 'Total Harga Yang Dikenakan GST SR
                        End If
                    End If
                    
                End If
                
            End If
            
            If Not IsNull(rs!dulang) Then 'Dulang
                .ListSubItems.Add , , rs!dulang
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!dimension_Panjang) Then 'Panjang
                .ListSubItems.Add , , rs!dimension_Panjang
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!dimension_Lebar) Then 'Lebar
                .ListSubItems.Add , , rs!dimension_Lebar
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!dimension_Saiz) Then 'Saiz
                .ListSubItems.Add , , rs!dimension_Saiz
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!gst_ari_nashi) Then 'Jenis Cukai GST
            
                If rs!gst_ari_nashi = 0 Then

                    .ListSubItems.Add , , "ZR(L)"
                    
                    If IsNumeric(rs!jumlah_gst) Then Frm83_LM_GST_ZR = Frm83_LM_GST_ZR + rs!jumlah_gst
                    
                ElseIf rs!gst_ari_nashi = 1 Then

                    .ListSubItems.Add , , "SR"
                    
                    If IsNumeric(rs!jumlah_gst) Then Frm83_LM_GST_SR = Frm83_LM_GST_SR + rs!jumlah_gst
                    
                End If
            End If
            
            If Not IsNull(rs!jumlah_gst) Then 'Jumlah GST (RM)
                .ListSubItems.Add , , Format(rs!jumlah_gst, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!harga_item) Then 'Harga Belian Dengan GST (RM)
                .ListSubItems.Add , , Format(rs!harga_item, "#,##0.00")
                
                If IsNumeric(rs!harga_item) Then Frm83_LM_HARGA_DENGAN_GST = Frm83_LM_HARGA_DENGAN_GST + rs!harga_item 'Harga Belian Dengan GST (RM)
                
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!code1) Then 'Code 1
                .ListSubItems.Add , , rs!code1
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!code2) Then 'Code 2
                .ListSubItems.Add , , rs!code2
            Else
                .ListSubItems.Add , , ""
            End If
            
        End With
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

LM_X = 0
Frm83.L11_Text = Format(0, "#,##0.00")
Frm83.L26_Text = Format(0, "0.00") 'Harga Belian Dengan GST (RM)
If Frm83.L41_Text = 1 Then
    Frm84.L58_Text = Format(0, "0.00") 'Harga Belian Dengan GST (RM)
End If
Frm83.TB40 = Format(0, "0.00")

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(id) , SUM(harga_tanpa_gst) , SUM(harga_item) from " & G_BELIAN_TEMP & " where (StatusItem <> 0 and StatusItem <> 5)", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    frm83_LM_TOTAL_PAGE = Format(rs(0) / frm83_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, frm83_LM_TOTAL_PAGE, ".") <> 0 Then
    
        frm83_LM_PAGE = Split(frm83_LM_TOTAL_PAGE, ".")(0)
        frm83_LM_PAGE_LEBIHAN = Split(frm83_LM_TOTAL_PAGE, ".")(1)
        
        If frm83_LM_PAGE_LEBIHAN <> "00" Then
            Frm83.L68_Text = frm83_LM_PAGE + 1
        Else
            Frm83.L68_Text = frm83_LM_PAGE
        End If
        
    Else
    
        Frm83.L68_Text = frm83_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm83.L68_Text = 0
    End If
Else
    Frm83.L68_Text = 0
End If

If Not IsNull(rs(0)) Then LM_X = rs(0)
If Not IsNull(rs(1)) Then Frm83.L11_Text = Format(rs(1), "#,##0.00")
If Not IsNull(rs(2)) Then
    Frm83.L26_Text = Format(rs(2), "0.00") 'Harga Belian Dengan GST (RM)
    If Frm83.L41_Text = 1 Then
        Frm84.L58_Text = Format(rs(2), "0.00") 'Harga Belian Dengan GST (RM)
    End If
    Frm83.TB40 = Format(rs(2), "0.00")
End If

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_gst) , SUM(harga_tanpa_gst) from " & G_BELIAN_TEMP & " where (StatusItem <> 0 and StatusItem <> 5) AND gst_ari_nashi = 0", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm83.L22_Text = Format(rs(0), "#,##0.00")  'Total Harga Yang Dikenakan GST ZR
If Not IsNull(rs(1)) Then Frm83.L23_Text = Format(rs(1), "#,##0.00")  'Jumlah Kutipan GST ZR(L)

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_gst) , SUM(harga_tanpa_gst) from " & G_BELIAN_TEMP & " where (StatusItem <> 0 and StatusItem <> 5) AND gst_ari_nashi = 1", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm83.L24_Text = Format(rs(0), "#,##0.00")  'Total Harga Yang Dikenakan GST SR
If Not IsNull(rs(1)) Then Frm83.L25_Text = Format(rs(1), "#,##0.00")  'Jumlah Kutipan GST SR

rs.Close
Set rs = Nothing
            
            
Frm83.L10_Text = LM_X

If x <> 0 Then
    Frm83.L69_Text = LM_START_ROW
End If

If Frm83.L67_Text <> vbNullString And IsNumeric(Frm83.L67_Text) Then
    If Frm83.L68_Text <> vbNullString And IsNumeric(Frm83.L68_Text) Then
        frm83_LM_CURR_PAGE = Frm83.L67_Text
        frm83_LM_TOTAL_PAGE = Frm83.L68_Text
        
        If frm83_LM_CURR_PAGE > frm83_LM_TOTAL_PAGE Then
            
            Frm83.L67_Text = Frm83.L67_Text - 1
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            GoTo re_gen_report:
            
        End If
    End If
End If
End Sub
Sub Frm84_Senarai_Jualan2()
'on error resume next
Dim Frm84_LM_HARGA_TANPA_GST As Double 'Harga Jualan Tanpa Cukai GST
Dim Frm84_LM_HARGA_DENGAN_GST As Double 'Harga Jualan Dengan Cukai GST
Dim Frm84_LM_GST_SR As Double 'Kutipan GST : SR
Dim Frm84_LM_GST_ZR As Double 'Kutipan GST : ZR
Dim Frm84_LM_JUMLAH_HARGA_SR As Double 'Total Harga Yang Dikenakan GST SR
Dim Frm84_LM_JUMLAH_HARGA_ZR As Double 'Total Harga Yang Dikenakan GST ZR
Dim Frm84_LM_BERAT As Double 'Berat Jualan
Dim Frm84_LM_JUALAN_GST As Double
Dim Frm84_LM_JUALAN_DENGAN_GST As Double
Dim Frm84_LM_JUALAN_TANPA_GST As Double
Dim Frm84_LM_HARGA_JUALAN_TANPA_GST As Double 'Harga jualan barang kemas tanpa GST

x = 0
Frm84_LM_HARGA_TANPA_GST = 0
Frm84_LM_HARGA_DENGAN_GST = 0
Frm84_LM_GST_SR = 0
Frm84_LM_GST_ZR = 0
Frm84_LM_JUMLAH_HARGA_SR = 0
Frm84_LM_JUMLAH_HARGA_ZR = 0
Frm84_LM_HARGA_JUALAN_TANPA_GST = 0 'Harga jualan barang kemas tanpa GST

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from " & G_JUALAN_TEMP & "", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    Frm84_LM_JUALAN_GST = 0
    Frm84_LM_JUALAN_DENGAN_GST = 0
    Frm84_LM_JUALAN_TANPA_GST = 0
    
    If rs!Status = 1 Or rs!Status = 3 Or rs!Status = 4 Then
        x = x + 1
        
        With Frm84.ListView2.ListItems.Add(, , rs!ID)
            .ListSubItems.Add , , x
            If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
            
            If Not IsNull(rs!no_siri_Produk) Then 'No. Siri Produk
                .ListSubItems.Add , , rs!no_siri_Produk
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!kategori_Produk) Then 'Kategori Produk
                .ListSubItems.Add , , rs!kategori_Produk
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!purity) Then 'Purity
                .ListSubItems.Add , , rs!purity
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!Berat_Asal) Then 'Berat Asal (g)
                .ListSubItems.Add , , Format(rs!Berat_Asal, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!berat_jualan) Then 'Berat Jualan (g)
                .ListSubItems.Add , , Format(rs!berat_jualan, "#,##0.00")
                If IsNumeric(rs!berat_jualan) Then Frm84_LM_BERAT = Frm84_LM_BERAT + rs!berat_jualan 'Jumlah Berat Jualan
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!harga_Semasa) Then 'Harga Semasa Bagi Purity Ini (RM/g)
                .ListSubItems.Add , , Format(rs!harga_Semasa, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!UPAH) Then 'Upah (RM)
                .ListSubItems.Add , , Format(rs!UPAH, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!harga_asal) Then 'Harga Asal Item (RM)
                .ListSubItems.Add , , Format(rs!harga_asal, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!diskaun) Then 'Diskaun (%)
                .ListSubItems.Add , , Format(rs!diskaun, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!harga_lepas_diskaun) Then 'Harga Lepas Diskaun (RM)
                .ListSubItems.Add , , Format(rs!harga_lepas_diskaun, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!adjustment) Then 'Adjustment (RM)
                .ListSubItems.Add , , Format(rs!adjustment, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!harga_jualan) Then 'Harga Jualan (RM)
                .ListSubItems.Add , , Format(rs!harga_jualan, "#,##0.00")
                If IsNumeric(rs!harga_tanpa_gst) Then Frm84_LM_HARGA_TANPA_GST = Frm84_LM_HARGA_TANPA_GST + rs!harga_tanpa_gst 'Harga Jualan Tanpa GST (RM)
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!gst_ari_nashi) Then 'Jenis GST
            
                If rs!gst_ari_nashi = "ZR (L)" Then
                    .ListSubItems.Add , , "ZR(L)"  'Jenis GST : Zero Rated
                    If IsNumeric(rs!jumlah_gst) Then Frm84_LM_GST_ZR = Frm84_LM_GST_ZR + rs!jumlah_gst 'Jumlah Kutipan GST ZR(L)
                    If IsNumeric(rs!harga_dengan_gst) Then Frm84_LM_JUMLAH_HARGA_ZR = Frm84_LM_JUMLAH_HARGA_ZR + rs!harga_dengan_gst 'Total Harga Yang Dikenakan GST ZR
                ElseIf rs!gst_ari_nashi = "SR" Then
                    .ListSubItems.Add , , "SR"  'Jenis GST : Standard Rated
                    If IsNumeric(rs!jumlah_gst) Then Frm84_LM_GST_SR = Frm84_LM_GST_SR + rs!jumlah_gst 'Jumlah Kutipan GST SR
                    If IsNumeric(rs!harga_tanpa_gst) Then Frm84_LM_JUMLAH_HARGA_SR = Frm84_LM_JUMLAH_HARGA_SR + rs!harga_tanpa_gst 'Total Harga Yang Dikenakan GST SR
                End If
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!gst_include) Then
                If rs!gst_include = 0 Then
                    .ListSubItems.Add , , "Tidak"  'Harga Termasuk GST
                Else
                    .ListSubItems.Add , , "Ya" 'Harga Termasuk GST
                End If
            Else
                .ListSubItems.Add , , "Tidak"  'Harga Termasuk GST
            End If
            
            If Not IsNull(rs!jumlah_gst) Then 'Jumlah GST (RM)
                .ListSubItems.Add , , Format(rs!jumlah_gst, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!harga_jualan_dengan_gst) Then 'Harga Dengan GST (RM)
                .ListSubItems.Add , , Format(rs!harga_jualan_dengan_gst, "#,##0.00")
                
                If IsNumeric(rs!harga_jualan_dengan_gst) And IsNumeric(rs!jumlah_gst) Then
                    Frm84_LM_HARGA_JUALAN_TANPA_GST = Frm84_LM_HARGA_JUALAN_TANPA_GST + (rs!harga_jualan_dengan_gst - rs!jumlah_gst) 'Harga jualan barang kemas tanpa GST
                End If
                
                If Not IsNull(rs!harga_jualan_dengan_gst) Then
                    If IsNumeric(rs!harga_dengan_gst) Then Frm84_LM_HARGA_DENGAN_GST = Frm84_LM_HARGA_DENGAN_GST + rs!harga_dengan_gst 'Harga Jualan Dengan GST (RM)
                End If
        
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!komisyen_per_gram) Then 'Komisen Per Gram (RM/g)
                .ListSubItems.Add , , Format(rs!komisyen_per_gram, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!komisyen_upah) Then 'Jumlah Komisyen Bagi Upah (RM)
                .ListSubItems.Add , , Format(rs!komisyen_upah, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!jumlah_komisyen) Then 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini (RM)
                .ListSubItems.Add , , Format(rs!jumlah_komisyen, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If

        End With

    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm84.L4_Text = x
Frm84.L14_Text = x
Frm84.L5_Text = Format(Frm84_LM_HARGA_TANPA_GST, "#,##0.00") 'Harga Jualan Tanpa GST (RM)
Frm84.L17_Text = Format(Frm84_LM_HARGA_JUALAN_TANPA_GST, "#,##0.00") 'Harga Jualan Tanpa GST (RM)
Frm84.L6_Text = Format(Frm84_LM_HARGA_DENGAN_GST, "#,##0.00") 'Harga Jualan Dengan GST (RM)
Frm84.L18_Text = Format(Frm84_LM_GST_SR + Frm84_LM_GST_ZR, "#,##0.00") 'Jumlah Cukai GST (RM)
Frm84.L19_Text = Format(Frm84_LM_HARGA_JUALAN_TANPA_GST + Frm84_LM_GST_SR + Frm84_LM_GST_ZR, "#,##0.00") 'Harga Jualan Dengan GST (RM)
Frm84.L7_Text = Format(Frm84_LM_JUMLAH_HARGA_ZR, "#,##0.00")  'Total Harga Yang Dikenakan GST ZR
Frm84.L9_Text = Format(Frm84_LM_GST_ZR, "#,##0.00")  'Jumlah Kutipan GST ZR(L)
Frm84.L10_Text = Format(Frm84_LM_JUMLAH_HARGA_SR, "#,##0.00")  'Total Harga Yang Dikenakan GST SR
Frm84.L11_Text = Format(Frm84_LM_GST_SR, "#,##0.00")  'Jumlah Kutipan GST SR
Frm84.L15_Text = Format(Frm84_LM_BERAT, "#,##0.00") 'Jumlah Berat Jualan
End Sub
Sub Frm83_Reset_After_Save()
'on error resume next
If Frm83.CB8 = 1 Then
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If rs!Default1 = "Default" Then
            If Not IsNull(rs!NoRujukanSistem) Then Frm83.L9_Text = rs!NoRujukanSistem 'No. Rujukan Sistem
            If Not IsNull(rs!NoRujukanStock) Then
                If Frm83.CB9 = 1 Then
                '    Frm83.TB7 = Format(rs!NoRujukanStock, "000000") 'No. Siri Barcode
                '    Frm83.L3_Text = rs!NoRujukanStock 'No. Siri Barcode
                ElseIf Frm83.CB10 = 1 Then
                '    Frm83.TB7 = Format(rs!no_siri_gb, "000000") & "W" 'No. Siri Barcode
                '    Frm83.L3_Text = rs!no_siri_gb 'No. Siri Barcode
                End If
            End If
            'If Frm83.CB8 = 1 Then
                If Not IsNull(rs!no_resit_trade_in) Then Frm83.L12_Text = rs!no_resit_trade_in 'No. Resit
            'End If
        End If
    End If
    
    rs.Close
    Set rs = Nothing
End If

'###Padam Temp Table###
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 17_gold_bar_buy_temp", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    rs.Delete
    rs.Update
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm83.TB15 = vbNullString
'Frm83.TB28 = vbNullString
End Sub
Sub Frm83_Edit_Data()
'on error resume next
Dim rs2 As ADODB.Recordset

DATA_FOUND = 0

Frm83_LM_ID = vbNullString

If IsNumeric(Frm83.ListView2.SelectedItem.Index) Then
    
    Frm83_LM_ID = Frm83.ListView2.ListItems(Frm83.ListView2.SelectedItem.Index)
    
Else
    
    MsgBox "Tiada Data Dijumpai.", vbInformation, "Info"
    
    Exit Sub
    
End If

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from " & G_BELIAN_TEMP & " where ID='" & Frm83_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!StatusItem) Then
        If rs!StatusItem = "10" Or rs!StatusItem = "3" Or rs!StatusItem = "4" Then
            GLOBAL_DISABLE = 1
            If Not IsNull(rs!ID) Then Frm83.L13_Text = rs!ID 'No. ID
            If Not IsNull(rs!id_database) Then Frm83.L20_Text = rs!id_database 'ID Dari Database Asal
            If Not IsNull(rs!supplier_ID) Then Frm83.L4_Text = rs!supplier_ID 'No. ID Bagi Supplier
            If Not IsNull(rs!Kod_Supplier) Then Frm83.TB1 = rs!Kod_Supplier 'Kod Supplier
            If Not IsNull(rs!purity_ID) Then Frm83.L5_Text = rs!purity_ID 'No. ID Bagi Purity
            If Not IsNull(rs!kod_Purity) Then Frm83.TB2 = rs!kod_Purity 'Kod Purity
            If Not IsNull(rs!no_cert) Then Frm83.TB34 = rs!no_cert 'No. Cert
            If Not IsNull(rs!upah_per_gram) Then Frm83.TB35 = rs!upah_per_gram
            
            If Not IsNull(rs!kategori_produk_ID) Then Frm83.L6_Text = rs!kategori_produk_ID 'No. ID Bagi Kategori Produk
            If Not IsNull(rs!Kod_Kategori_Produk) Then
                Frm83.TB3 = rs!Kod_Kategori_Produk 'Kod Kategori Produk
                Frm83.TB6 = rs!Kod_Kategori_Produk 'Kod Kategori Produk
            End If
            If Not IsNull(rs!Barcode) Then Frm83.TB7 = rs!Barcode 'No. Barcode (6 Digit Terakhir)
            If Not IsNull(rs!gst_barang_atau_upah) Then '0 : Pengiraan GST pada harga belian barang , 1 : Pengiraan GST pada upah sahaja
                If rs!gst_barang_atau_upah = 0 Then
                    Frm83.CB12 = 0
                ElseIf rs!gst_barang_atau_upah = 1 Then
                    Frm83.CB12 = 1
                End If
            Else
                Frm83.CB12 = 0
            End If
            
'### Jenis ###
'0 : BK
'1 : Barang permata
'2 : Emas terpakai BK
'3 : Emas terpakai permata
'4 : gold Bar
'5 : Emas terpakai gold bar
'6 : Trade In BK
'7 : Trade In Barang Permata
'8 : Trade In Gold Bar

            If Not IsNull(rs!jenis) Then
                If rs!jenis = 0 Then '0 : BK , 1 : Barang Permata , 2 : Buyback BK , 3 : Buyback Barang Permata
                    Frm83.CB7 = 1 'Penerimaan stok baru
                    Frm83.CB4 = 1 'Barang kemas
                    Frm83.CB5 = 0 'Barang permata
                    If Not IsNull(rs!Upah_Jualan) Then Frm83.TB24 = rs!Upah_Jualan 'Upah Jualan Kepada Pelanggan
                    If Not IsNull(rs!Upah_Member) Then Frm83.TB25 = rs!Upah_Member 'Upah Jualan Kepada Ahli / Member
                    If Not IsNull(rs!Upah_Pengedar) Then Frm83.TB26 = rs!Upah_Pengedar 'Upah Jualan Kepada Pengedar
                    If Not IsNull(rs!Upah_RAF) Then Frm83.TB31 = rs!Upah_RAF 'Upah Jualan Kepada RAF
                    If Not IsNull(rs!upah_normal_dealer) Then Frm83.TB32 = rs!upah_normal_dealer 'Upah Jualan Kepada Normal Dealer
                    If Not IsNull(rs!upah_master_dealer) Then Frm83.TB33 = rs!upah_master_dealer 'Upah Jualan Kepada Master Dealer
                ElseIf rs!jenis = 1 Then '0 : BK , 1 : Barang Permata , 2 : Buyback BK , 3 : Buyback Barang Permata
                    Frm83.CB7 = 1 'Penerimaan stok baru
                    Frm83.CB5 = 1 'Barang permata
                    Frm83.CB4 = 0 'Barang kemas
                    If Not IsNull(rs!code_Supplier) Then Frm83.TB24 = rs!code_Supplier 'Harga Jualan Kepada Pelanggan
                    If Not IsNull(rs!HargaJualan_Member) Then Frm83.TB25 = rs!HargaJualan_Member 'Harga Jualan Kepada Ahli / Member
                    If Not IsNull(rs!HargaJualan_Pengedar) Then Frm83.TB26 = rs!HargaJualan_Pengedar 'Harga Jualan Kepada Pengedar
                    If Not IsNull(rs!HargaJualan_RAF) Then Frm83.TB31 = rs!HargaJualan_RAF 'Harga Jualan Kepada RAF
                    If Not IsNull(rs!hargajualan_normal_dealer) Then Frm83.TB32 = rs!hargajualan_normal_dealer 'Harga Jualan Kepada Normal Dealer
                    If Not IsNull(rs!hargajualan_master_dealer) Then Frm83.TB33 = rs!hargajualan_master_dealer 'Harga Jualan Kepada Master Dealer
                ElseIf rs!jenis = 2 Or rs!jenis = 6 Then '0 : BK , 1 : Barang Permata , 2 : Buyback BK , 3 : Buyback Barang Permata
                    Frm83.CB8 = 1 'Buyback / Trade in
                    Frm83.CB4 = 1 'Barang kemas
                    Frm83.CB5 = 0 'Barang permata
                    If Not IsNull(rs!Upah_Jualan) Then Frm83.TB24 = rs!Upah_Jualan 'Upah Jualan Kepada Pelanggan
                    If Not IsNull(rs!Upah_Member) Then Frm83.TB25 = rs!Upah_Member 'Upah Jualan Kepada Ahli / Member
                    If Not IsNull(rs!Upah_Pengedar) Then Frm83.TB26 = rs!Upah_Pengedar 'Upah Jualan Kepada Pengedar
                    If Not IsNull(rs!Upah_RAF) Then Frm83.TB31 = rs!Upah_RAF 'Upah Jualan Kepada RAF
                    If Not IsNull(rs!upah_normal_dealer) Then Frm83.TB32 = rs!upah_normal_dealer 'Upah Jualan Kepada Normal Dealer
                    If Not IsNull(rs!upah_master_dealer) Then Frm83.TB33 = rs!upah_master_dealer 'Upah Jualan Kepada Master Dealer
                ElseIf rs!jenis = 3 Or rs!jenis = 7 Then '0 : BK , 1 : Barang Permata , 2 : Buyback BK , 3 : Buyback Barang Permata
                    Frm83.CB8 = 1 'Buyback / Trade in
                    Frm83.CB5 = 1 'Barang permata
                    Frm83.CB4 = 0 'Barang kemas
                    If Not IsNull(rs!code_Supplier) Then Frm83.TB24 = rs!code_Supplier 'Harga Jualan Kepada Pelanggan
                    If Not IsNull(rs!HargaJualan_Member) Then Frm83.TB25 = rs!HargaJualan_Member 'Harga Jualan Kepada Ahli / Member
                    If Not IsNull(rs!HargaJualan_Pengedar) Then Frm83.TB26 = rs!HargaJualan_Pengedar 'Harga Jualan Kepada Pengedar
                    If Not IsNull(rs!HargaJualan_RAF) Then Frm83.TB31 = rs!HargaJualan_RAF 'Harga Jualan Kepada RAF
                    If Not IsNull(rs!hargajualan_normal_dealer) Then Frm83.TB32 = rs!hargajualan_normal_dealer 'Harga Jualan Kepada Normal Dealer
                    If Not IsNull(rs!hargajualan_master_dealer) Then Frm83.TB33 = rs!hargajualan_master_dealer 'Harga Jualan Kepada Master Dealer
                End If
            End If
            
            If Not IsNull(rs!Berat) Then Frm83.TB8 = rs!Berat 'Berat
            If Not IsNull(rs!kos_Belian_Gram) Then Frm83.TB9 = rs!kos_Belian_Gram 'Harga Per Gram (Belian)
            If Not IsNull(rs!UPAH) Then Frm83.TB4 = rs!UPAH 'Upah (RM)
            If Not IsNull(rs!kos_Belian_Item) Then Frm83.TB10 = rs!kos_Belian_Item 'Harga Asal (RM)
            If Not IsNull(rs!Spread) Then
                Frm83.TB19 = rs!Spread 'Spread (%)
            Else
                Frm83.TB19 = "0.00" 'Spread (%)
            End If
            'If Not IsNull(rs!jenis) Then
            
            'Else
            
            'End If
            '    If rs!jenis = 0 Then
            '        Frm83.TB19 = "0.00" 'Spread (%)
            '    ElseIf rs!jenis = 1 Then
            '        If Not IsNull(rs!Spread) Then Frm83.TB19 = rs!Spread 'Spread (%)
            '    End If
            'End If
            If Not IsNull(rs!harga_lepas_spread) Then Frm83.TB21 = rs!harga_lepas_spread 'Harga asal ditolak spread (RM)
            If Not IsNull(rs!adjustment) Then Frm83.TB22 = rs!adjustment 'Adjustment (RM)
            If Not IsNull(rs!kos_item_tanpa_tax) Then Frm83.TB20 = rs!kos_item_tanpa_tax 'Harga Barang + Upah Tanpa Tax
            If Not IsNull(rs!dimension_Panjang) Then
                Frm83.TB12 = rs!dimension_Panjang 'Panjang
            Else
                Frm83.TB12 = vbNullString 'Panjang
            End If
            If Not IsNull(rs!dimension_Lebar) Then
                Frm83.TB13 = rs!dimension_Lebar 'Lebar
            Else
                Frm83.TB13 = vbNullString 'Lebar
            End If
            If Not IsNull(rs!dimension_Saiz) Then
                Frm83.TB14 = rs!dimension_Saiz 'Saiz
            Else
                Frm83.TB14 = vbNullString 'Saiz
            End If
            
            If Not IsNull(rs!code1) Then 'Code 1
                Frm83.TB36 = rs!code1
            Else
                Frm83.TB36 = vbNullString
            End If
            If Not IsNull(rs!code2) Then 'Code 2
                Frm83.TB37 = rs!code2
            Else
                Frm83.TB37 = vbNullString
            End If
            
            If Not IsNull(rs!remarks) Then
                Frm83.TB16 = rs!remarks 'Remarks
            Else
                Frm83.TB16 = vbNullString 'Remarks
            End If
            If Not IsNull(rs!flag_upah) Then
                
                If rs!flag_upah = 0 Then
                    Frm83.CB14 = 1
                    Frm83.CB15 = 0
                ElseIf rs!flag_upah = 1 Then
                    Frm83.CB14 = 0
                    Frm83.CB15 = 1
                End If
                
            End If
            If Not IsNull(rs!UPAH) Then Frm83.TB4 = rs!UPAH 'Upah (RM)
            If Not IsNull(rs!upah_per_gram) Then Frm83.TB35 = Format(rs!upah_per_gram, "0.00")
            
            If rs!flag_image = 1 Then
                'Set rs2 = New ADODB.Recordset
                'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                'rs2.Open "select * from 1_image_item_temp where initial_flag='" & "1" & "'", cn, adOpenKeyset, adLockOptimistic
                
                'If rs2.EOF Then
                '    rs2.AddNew
                '    rs2!initial_flag = 1
                '    rs2!Image = rs!Image
                '    rs2.Update
                'Else
                '    rs2!Image = rs!Image
                '    rs2.Update
                'End If
                
                'rs2.Close
                'Set rs2 = Nothing
                
                Frm83.L31_Text = "Ada"
                Frm83.L32_Text = 1
            Else
                Frm83.L31_Text = "Tiada"
                Frm83.L32_Text = 0
            End If
        
            If Not IsNull(rs!gst_ari_nashi) Then
                
                If rs!gst_ari_nashi = 0 Then
                    
                    Frm83.CB2 = 1
                    Frm83.CB3 = 0
                    Frm83.CB11 = 0
                    
                ElseIf rs!gst_ari_nashi = 1 Then
                    
                    If rs!gst_included = 0 Then
                    
                        Frm83.CB2 = 0
                        Frm83.CB3 = 1
                        Frm83.CB11 = 0
                        
                    ElseIf rs!gst_included = 1 Then
                
                        Frm83.CB2 = 0
                        Frm83.CB3 = 0
                        Frm83.CB11 = 1
                    
                    End If
                    
                    If Not IsNull(rs!kadar_gst) Then
                        Frm83.L8_Text = rs!kadar_gst 'Kadar GST (%)
                    Else
                        Frm83.L8_Text = "0.00" 'Kadar GST (%)
                    End If
                    If Not IsNull(rs!jumlah_gst) Then
                        Frm83.TB27 = rs!jumlah_gst 'Jumlah Cukai GST (RM)
                    Else
                        Frm83.TB27 = "0.00" 'Jumlah Cukai GST (RM)
                    End If
                
                End If
                
            End If
        
'If Not IsNull(rs!gst_ari_nashi) Then
'    If rs!gst_ari_nashi = 0 Then 'Status Cukai GST : 0 : ZR(L) , 1 : SR
'        Frm83.CB3 = 0
'        Frm83.CB11 = 0
'        Frm83.CB2 = 1
'        Frm83.CB3 = 0
        'If Not IsNull(rs!kadar_gst) Then
        '    Frm83.L8_Text = rs!kadar_gst 'Kadar GST (%)
        'Else
        '    Frm83.L8_Text = vbNullString 'Kadar GST (%)
        'End If
'        If Not IsNull(rs!jumlah_gst) Then
'            Frm83.TB27 = rs!jumlah_gst 'Jumlah Cukai GST (RM)
'        Else
'            Frm83.TB27 = "0.00" 'Jumlah Cukai GST (RM)
'        End If
'    ElseIf rs!gst_ari_nashi = 1 Then
'        Frm83.CB2 = 0
'        Frm83.CB3 = 1
'        Frm83.CB11 = 1
        
'        If Not IsNull(rs!gst_included) Then
'            If rs!gst_included = 0 Then
'                Frm83.CB11 = 0
'            ElseIf rs!gst_included = 1 Then
'                Frm83.CB11 = 1
'            End If
'        Else
'            Frm83.CB11 = 0
'        End If
        
'        If Not IsNull(rs!kadar_gst) Then
'            Frm83.L8_Text = rs!kadar_gst 'Kadar GST (%)
'        Else
'            Frm83.L8_Text = "0.00" 'Kadar GST (%)
'        End If
'        If Not IsNull(rs!jumlah_gst) Then
'            Frm83.TB27 = rs!jumlah_gst 'Jumlah Cukai GST (RM)
'        Else
'            Frm83.TB27 = "0.00" 'Jumlah Cukai GST (RM)
'        End If
'    End If
'End If
            If Not IsNull(rs!no_id_gst) Then
                Frm83.TB28 = rs!no_id_gst
            Else
                Frm83.TB28 = vbNullString
            End If
            If Not IsNull(rs!bill_No_Belian) Then
                Frm83.TB15 = rs!bill_No_Belian
            Else
                Frm83.TB15 = vbNullString
            End If
            If Not IsNull(rs!tarikh_belian) Then Frm83.DTPicker1 = rs!tarikh_belian
            
            On Error GoTo Err_A:
            If Not IsNull(rs!nama_Supplier) Then
                Frm83_LM_Supplier = rs!nama_Supplier 'Nama Supplier
                Frm83.CBB1 = Frm83_LM_Supplier 'Nama Supplier
            End If
            
Restore_A:
        
            On Error GoTo Err_B:
            If Not IsNull(rs!purity) Then
                Frm83_LM_BRAND = rs!purity 'Purity
                Frm83.CBB2 = Frm83_LM_BRAND 'Purity
            End If
            
Restore_B:
        
            On Error GoTo Err_C:
            If Not IsNull(rs!kategori_Produk) Then
                Frm83_LM_PRODUK = rs!kategori_Produk 'Kategori Produk
                Frm83.CBB3 = Frm83_LM_PRODUK 'Kategori Produk
            End If
            
Restore_C:
        
            On Error GoTo Err_E:
            If Not IsNull(rs!dulang) Then
                Frm83_LM_DULANG = rs!dulang 'Dulang
                Frm83.CBB5 = Frm83_LM_DULANG 'Dulang
            End If
            
Restore_E:
            
            'rs!Status = 1
            DATA_FOUND = 1
            GLOBAL_DISABLE = 0
        ElseIf rs!StatusItem = "11" Then
            MsgBox "Item Ini Tidak Dibenarkan Untuk Diedit Kerana Telah Dijual.", vbExclamation, "Info"
        ElseIf rs!StatusItem = "12" Then
            MsgBox "Item Ini Tidak Dibenarkan Untuk Diedit Kerana Telah Dijual Secara Potong.", vbExclamation, "Info"
        ElseIf rs!StatusItem = "13" Then
            MsgBox "Item Ini Tidak Dibenarkan Untuk Diedit Kerana Telah Dijual Secara Potong.", vbExclamation, "Info"
        ElseIf rs!StatusItem = "14" Or rs!StatusItem = "21" Or rs!StatusItem = "22" Then
            MsgBox "Item Ini Tidak Dibenarkan Untuk Diedit Kerana Ditempah Oleh Pelanggan.", vbExclamation, "Info"
        ElseIf rs!StatusItem = "15" Or rs!StatusItem = "19" Or rs!StatusItem = "20" Then
            MsgBox "Item Ini Tidak Dibenarkan Untuk Diedit Kerana Dibeli Secara Ansuran.", vbExclamation, "Info"
        ElseIf rs!StatusItem = "16" Then
            MsgBox "Item Ini Tidak Dibenarkan Untuk Diedit Kerana Telah Dihantar Ke Ar-Rahnu.", vbExclamation, "Info"
        ElseIf rs!StatusItem = "17" Then
            MsgBox "Item Ini Tidak Dibenarkan Untuk Diedit Kerana Dijual Secara ETA.", vbExclamation, "Info"
        ElseIf rs!StatusItem = "23" Then
            MsgBox "Item Ini Tidak Dibenarkan Untuk Diedit Kerana Telah Dihantar Ke Supplier/Kilang.", vbExclamation, "Info"
        ElseIf rs!StatusItem = "24" Then
            MsgBox "Item Ini Tidak Dibenarkan Untuk Diedit Kerana Telah Dihantar Ke Supplier/Kilang.", vbExclamation, "Info"
        ElseIf rs!StatusItem = "25" Then
            MsgBox "Item Ini Tidak Dibenarkan Untuk Diedit Kerana Telah Diagihkan Ke Cawangan.", vbExclamation, "Info"
        ElseIf rs!StatusItem = "26" Then
            MsgBox "Item Ini Tidak Dibenarkan Untuk Diedit Kerana Telah Dijual Oleh Cawangan.", vbExclamation, "Info"
        ElseIf rs!StatusItem = "0" Then
            MsgBox "Item Ini Tidak Dibenarkan Untuk Diedit Kerana Telah Dipadamkan Dari Sistem.", vbExclamation, "Info"
        ElseIf rs!StatusItem = "27" Or rs!StatusItem = "28" Then
            MsgBox "Item Ini Telah Dijual Dari Menu GDN.", vbExclamation, "Info"
        ElseIf rs!StatusItem = "29" Then
            MsgBox "Item Ini Telah Diubah Status Kepada Hilang , Dicuri Dan Sebagainya.", vbExclamation, "Info"
        End If
        
    End If
End If

rs.Close
Set rs = Nothing

If DATA_FOUND = 1 Then
    'Frm83.Pic1.Visible = False
    Frm83.L14_Text.Visible = True
    
    Frm83.CB4.Enabled = False
    Frm83.CB5.Enabled = False

    If Frm83.L21_Text = 0 Then
        Frm83.CMD1.Visible = False
        Frm83.CMD6.Visible = True
        Frm83.CMD7.Visible = True
    ElseIf Frm83.L21_Text = 1 Then
        Frm83.CMD12.Visible = False
        Frm83.CMD13.Visible = True
        Frm83.CMD14.Visible = True
    End If
    
    Frm83.Frame1.Visible = True
    Frm83.Frame9.Visible = False
End If

Exit Sub
Err_A:
Frm83.CBB1.AddItem Frm83_LM_Supplier
Frm83.CBB1 = Frm83_LM_Supplier
Resume Restore_A:

Exit Sub
Err_B:
Frm83.CBB2.AddItem Frm83_LM_BRAND
Frm83.CBB2 = Frm83_LM_BRAND
Resume Restore_B:

Exit Sub
Err_C:
Frm83.CBB3.AddItem Frm83_LM_PRODUK
Frm83.CBB3 = Frm83_LM_PRODUK
Resume Restore_C:

Exit Sub
Err_E:
Frm83.CBB5.AddItem Frm83_LM_DULANG
Frm83.CBB5 = Frm83_LM_DULANG
Resume Restore_E:
End Sub
Sub Frm83_Cancel_Edit()
'on error resume next
Frm83.L14_Text.Visible = False

If Frm83.L21_Text = 0 Then
    Frm83.CMD1.Visible = True
    Frm83.CMD6.Visible = False
    Frm83.CMD7.Visible = False
End If
If Frm83.L21_Text = 1 Then
    Frm83.CMD12.Visible = True
    Frm83.CMD13.Visible = False
    Frm83.CMD14.Visible = False
End If

'Frm83.TB1 = vbNullString
Frm83.TB2 = vbNullString
Frm83.TB3 = vbNullString
Frm83.TB4 = 0
Frm83.TB6 = vbNullString
Frm83.TB7 = vbNullString
Frm83.TB8 = "0.00"
Frm83.TB9 = "0.00"
Frm83.TB10 = "0.00"
Frm83.TB12 = vbNullString
Frm83.TB13 = vbNullString
Frm83.TB14 = vbNullString
Frm83.TB15 = vbNullString
Frm83.TB16 = vbNullString
Frm83.TB19 = vbNullString
Frm83.TB20 = vbNullString
Frm83.TB21 = vbNullString
'Frm83.TB28 = vbNullString
Frm83.TB22 = "0.00"
Frm83.TB24 = "0.00"
Frm83.TB25 = "0.00"
Frm83.TB26 = "0.00"
Frm83.TB31 = "0.00"
Frm83.TB32 = "0.00"
Frm83.TB33 = "0.00"
Frm83.TB34 = vbNullString
Frm83.TB35 = "0.00"
Frm83.TB36 = vbNullString
Frm83.TB37 = vbNullString

If Frm83.CB9 = 1 Then
    Frm83.CB4 = 1
    Frm83.CB5 = 0
    Frm83.CB4.Enabled = True
    Frm83.CB5.Enabled = True
    
    Frm83.TB4.Locked = False
    Frm83.TB4.BackColor = &HFFFFFF
End If

Frm83.TB8.Locked = False
Frm83.TB9.Locked = False

Frm83.TB8.BackColor = &HFFFFFF
Frm83.TB9.BackColor = &HFFFFFF
Frm83.L27_Text = "Upah Jualan Pelanggan    RM"
Frm83.L28_Text = "Upah Jualan Ahli               RM"
Frm83.L29_Text = "Upah Jualan Silver            RM"

Frm83.L31_Text = "Tiada"
Frm83.L32_Text = 0

'Frm83.CBB1.Clear
Frm83.CBB2.Clear
Frm83.CBB3.Clear
Frm83.CBB5.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database order by Metal_Purity ASC , kategori_Produk ASC , SenaraiDulang ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    'If Not IsNull(rs!Supplier) Then Frm83.CBB1.AddItem rs!Supplier
    If Not IsNull(rs!Metal_Purity) Then Frm83.CBB2.AddItem rs!Metal_Purity
    If Not IsNull(rs!kategori_Produk) Then Frm83.CBB3.AddItem rs!kategori_Produk
    If Not IsNull(rs!SenaraiDulang) Then Frm83.CBB5.AddItem rs!SenaraiDulang
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 14_gold_bar_tetapan", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!default_setting = "Default" Then
        If Not IsNull(rs!Spread) Then Frm83.TB19 = Format(rs!Spread, "0.00") 'Spread %
    End If
End If

rs.Close
Set rs = Nothing

'If Frm83.CB9 = 1 Then
'    Frm83.TB7 = Format(Frm83.L3_Text, "000000") 'No. Turutan Barcode
'ElseIf Frm83.CB10 = 1 Then
'    Frm83.TB7 = Format(Frm83.L3_Text, "000000") & "W" 'No. Turutan Barcode
'End If
End Sub
Sub Frm83_Resit_Buyback_GB()
'on error resume next
Dim Frm83_LM_TOTAL_BERAT As Double

DATA_FOUND = 0
'G_No_RESIT_BUYBACK_GB = "000018" 'No. Resit Buyback

x = 0
Frm83_LM_TOTAL_BERAT = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 12_gold_bar_database where no_resit_trade_in='" & G_No_RESIT_BUYBACK_GB & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!no_rujukan_pelanggan_buyback) Then Frm83_LM_No_PELANGGAN = rs!no_rujukan_pelanggan_buyback 'No. Pelanggan (Buyback)
    If Not IsNull(rs!no_pekerja_belian) Then Frm83_LM_No_PEKERJA = rs!no_pekerja_belian 'No. Pelanggan
    If Not IsNull(rs!tarikh_belian) Then Report12.Sections("Section4").Controls("L4").Caption = rs!tarikh_belian 'Tarikh Belian
    If Not IsNull(rs!no_resit_trade_in) Then Report12.Sections("Section4").Controls("L3").Caption = rs!no_resit_trade_in 'No. Resit Jualan
    DATA_FOUND = 1
End If

rs.Close
Set rs = Nothing

If DATA_FOUND = 1 Then
'### Carian Maklumat Pelanggan ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm83_LM_No_PELANGGAN & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!Nama) Then Report12.Sections("Section4").Controls("L1").Caption = rs!Nama 'Nama Pelanggan
        If Not IsNull(rs!no_tel) Then Report12.Sections("Section4").Controls("L2").Caption = rs!no_tel 'No. Telefon
    End If
    
    rs.Close
    Set rs = Nothing
'### Carian Maklumat Pelanggan ### - End

'### Carian Maklumat Pekerja ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where NoPekerja='" & Frm83_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!Samaran) Then Report12.Sections("Section5").Controls("L7").Caption = rs!Samaran 'Nama Samaran Pekerja
    End If
    
    rs.Close
    Set rs = Nothing
'### Carian Maklumat Pekerja ### - End

'### Carian Maklumat Belian ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & G_No_RESIT_BUYBACK_GB & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!jumlah_asal) Then Report12.Sections("Section5").Controls("L8").Caption = Format(rs!jumlah_asal, "0.00") 'Jumlah Asal (Sub Total)
        If Not IsNull(rs!jumlah_gst) Then Report12.Sections("Section5").Controls("L9").Caption = Format(rs!jumlah_gst, "0.00") 'Jumlah GST
        If Not IsNull(rs!jumlah_sebenar) Then Report12.Sections("Section5").Controls("L10").Caption = Format(rs!jumlah_sebenar, "0.00") 'Jumlah Bayaran (Belian)
    End If
    
    rs.Close
    Set rs = Nothing
'### Carian Maklumat Belian ### - End

'### Paparan Resit ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 12_gold_bar_database where no_resit_trade_in='" & G_No_RESIT_BUYBACK_GB & "'", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        x = x + 1
        If Not IsNull(rs!Berat) Then
            If IsNumeric(rs!Berat) Then
                Frm83_LM_TOTAL_BERAT = Frm83_LM_TOTAL_BERAT + rs!Berat 'Jumlah Berat
            End If
        End If
        Report12.Sections("Section5").Controls("L5").Caption = x 'Jumlah Bilangan Barang
        Report12.Sections("Section5").Controls("L6").Caption = Format(Frm83_LM_TOTAL_BERAT, "0.00") 'Jumlah Brat
        Set Report12.DataSource = rs
        Report12.Show
        rs.MoveNext
    Wend
    
    'rs.Close
    Set rs = Nothing
End If
'### Paparan Resit ### - End

G_No_RESIT_BUYBACK_GB = vbNullString
End Sub
Sub Frm83_upload_image()
'on error resume next
DATA_SAVE = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 1_image_item_temp where initial_flag='" & "1" & "'", cn, adOpenKeyset, adLockOptimistic

If rs.EOF Then
    rs.AddNew
    'rs!barcode = Frm10.TB21 & Frm10.TB7 'Barcode
    
    rs!initial_flag = 1
    Set picstrm = New ADODB.Stream
    picstrm.Type = adTypeBinary
    picstrm.Open
    picstrm.LoadFromFile strpic
    rs!Image = picstrm.Read
    picstrm.Close
    Set picstrm = Nothing
    
    rs!write_timestamp = Now
    
    'Frm58.Image1 = Nothing
    strpic = vbNullString
    DATA_SAVE = 1
    rs.Update
Else
    'rs!barcode = Frm10.TB21 & Frm10.TB7 'Barcode

    Set picstrm = New ADODB.Stream
    picstrm.Type = adTypeBinary
    picstrm.Open
    picstrm.LoadFromFile strpic
    rs!Image = picstrm.Read
    picstrm.Close
    Set picstrm = Nothing
    
    rs!write_timestamp = Now
    
    'Frm58.Image1 = Nothing
    strpic = vbNullString
    DATA_SAVE = 1
    rs.Update
End If

rs.Close
Set rs = Nothing

If DATA_SAVE = 1 Then
    Frm10.L8_Text = 1
    MsgBox "Gambar telah berjaya disimpan.", vbInformation, "Info"
End If
End Sub
Sub Frm83_reset_list()
'on error resume next
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Default1 = "Default" Then
        If Not IsNull(rs!NoRujukanSistem) Then Frm83.L9_Text = rs!NoRujukanSistem 'No. Rujukan Sistem
        'If Not IsNull(rs!tarikh) Then Frm83.DTPicker1 = rs!tarikh
        If Not IsNull(rs!gst_value) Then Frm83.L8_Text = rs!gst_value
        If Not IsNull(rs!riyal) Then Frm83.L30_Text = rs!riyal
        If Not IsNull(rs!NoRujukanStock) Then
            If Frm83.CB9 = 1 Then
            '    Frm83.TB7 = Format(rs!NoRujukanStock, "000000") 'No. Siri Barcode
            '    Frm83.L3_Text = rs!NoRujukanStock 'No. Siri Barcode
            ElseIf Frm83.CB10 = 1 Then
            '    Frm83.TB7 = Format(rs!no_siri_gb, "000000") & "W" 'No. Siri Barcode
            '    Frm83.L3_Text = rs!no_siri_gb 'No. Siri Barcode
            End If
        End If
        'If Frm83.CB8 = 1 Then
            If Not IsNull(rs!no_resit_trade_in) Then Frm83.L12_Text = rs!no_resit_trade_in 'No. Resit
        'End If
        If Frm83.CB8 = 1 Then
            If Not IsNull(rs!spread_Cash_Trade_In) Then Frm83.TB19 = Format(rs!spread_Cash_Trade_In, "0.00") 'Spread Trade In %
        ElseIf Frm83.CB7 = 1 Then
            Frm83.TB19 = "0.00" 'Spread Trade In %
        End If
        'If Not IsNull(rs!gst_arinashi_belian) Then
        '    If rs!gst_arinashi_belian = 1 Then
        '        Frm83.CB2 = 0
        '        Frm83.CB3 = 1
        '    Else
        '        Frm83.CB2 = 1
        '        Frm83.CB3 = 0
        '    End If
        'End If
    End If
End If

rs.Close
Set rs = Nothing

Frm83.TB40 = "0.00"
Frm83.TB41 = "0.00"
Frm83.TB42 = "0.00"

'###Padam Temp Table### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "TRUNCATE TABLE " & G_BELIAN_TEMP & ""

Set rs = cn.Execute(strsql)
Set rs = Nothing
'###Padam Temp Table### - End
End Sub
Sub Frm83_kira_upah()
'on error resume next
Dim Frm83_LM_BERAT
Dim Frm83_LM_UPAH_GRAM

Frm83_LM_BERAT = 0
Frm83_LM_UPAH_GRAM = 0

If Frm83.CB14 = 1 And (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) And (Frm83.TB35 <> vbNullString And IsNumeric(Frm83.TB35)) Then

    Frm83_LM_BERAT = Frm83.TB8
    Frm83_LM_UPAH_GRAM = Frm83.TB35
    
    Frm83.TB4 = Format(Frm83_LM_BERAT * Frm83_LM_UPAH_GRAM, "0.00")
End If
End Sub
Sub Frm83_mode_gold_bar()
'on error resume next
Frm83.CB14 = 0
Frm83.CB15 = 0
Frm83.TB35 = vbNullString
Frm83.TB35.BackColor = &H8000000A
Frm83.TB35.Locked = True
Frm83.CB14.Enabled = False
Frm83.CB15.Enabled = False
End Sub
Sub Frm83_jurujual()
'On Error Resume Next
If MDI_frm1.L3_Text <> vbNullString Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        Frm83.CBB6 = rs!Samaran & "  |  " & rs!NoPekerja
        
    Else
        
        Frm83.CBB6.AddItem "" & "  |  " & rs!Samaran
        Frm83.CBB6 = "" & "  |  " & rs!Samaran
        
    End If
    
    rs.Close
    Set rs = Nothing

    If G_LOCK_JURUJUAL = "YES" Then
    
        Frm83.CBB6.Enabled = False
        Frm83.CBB6.BackColor = &H8000000A

    Else
    
        Frm83.CBB6.Enabled = True
        Frm83.CBB6.BackColor = &HFFFFFF

    End If
End If
End Sub
Sub kiraan_gst_belian()
'On Error Resume Next
Dim Frm83_LM_HARGA As Double
Dim frm83_LM_KADAR_GST As Double
Dim frm83_TOTAL_GST As Double
Dim frm83_HARGA_TANPA_GST As Double

Frm83_LM_HARGA = 0
frm83_LM_KADAR_GST = 0
frm83_TOTAL_GST = 0
frm83_HARGA_TANPA_GST = 0

If Frm83.CB12 = 0 Then

    If Frm83.TB20 <> vbNullString And IsNumeric(Frm83.TB20) Then
        Frm83_LM_HARGA = Frm83.TB20
    End If
    
Else

    If Frm83.TB4 <> vbNullString And IsNumeric(Frm83.TB4) Then
        Frm83_LM_HARGA = Frm83.TB4
    End If
    
End If

If Frm83.L8_Text <> vbNullString And IsNumeric(Frm83.L8_Text) Then
    frm83_LM_KADAR_GST = Frm83.L8_Text
End If

If Frm83.CB2 = 1 Then
    
    Frm83.TB27 = Format(frm83_TOTAL_GST, "#,##0.00") 'Jumlah Cukai GST (RM)
    Frm83.L40_Text = Format(Frm83_LM_HARGA, "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)

ElseIf Frm83.CB3 = 1 Then

    Frm83.TB27 = Format((frm83_LM_KADAR_GST / 100) * Frm83_LM_HARGA, "#,##0.00") 'Jumlah Cukai GST (RM)
    Frm83.L40_Text = Format(Frm83_LM_HARGA, "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
    
ElseIf Frm83.CB11 = 1 Then

    Frm83.L40_Text = Format(Frm83_LM_HARGA / (1 + (frm83_LM_KADAR_GST / 100)), "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
    Frm83.TB27 = Format(Frm83_LM_HARGA - (Frm83_LM_HARGA / (1 + (frm83_LM_KADAR_GST / 100))), "#,##0.00") 'Jumlah Cukai GST (RM)
        
End If
End Sub
Sub frm83_kiraan_harga_asal()
'On Error Resume Next
Dim Frm83_LM_BERAT As Double
Dim Frm83_LM_UPAH As Double
Dim Frm83_LM_HARGA_SEMASA As Double

Frm83_LM_BERAT = 0
Frm83_LM_UPAH = 0
Frm83_LM_HARGA_SEMASA = 0

If (Frm83.TB8 <> vbNullString And IsNumeric(Frm83.TB8)) Then Frm83_LM_BERAT = Frm83.TB8
If (Frm83.TB4 <> vbNullString And IsNumeric(Frm83.TB4)) Then Frm83_LM_UPAH = Frm83.TB4
If (Frm83.TB9 <> vbNullString And IsNumeric(Frm83.TB9)) Then Frm83_LM_HARGA_SEMASA = Frm83.TB9

Frm83.TB10 = Format((Frm83_LM_BERAT * Frm83_LM_HARGA_SEMASA) + Frm83_LM_UPAH, "#,##0.00") 'Harga

Call Frm83_kira_upah
End Sub
Sub frm83_kiraan_harga_selepas_spread()
'On Error Resume Next
Dim Frm83_LM_HARGA_ASAL As Double
Dim Frm83_LM_SPREAD As Double
Dim Frm83_LM_HARGA_LEPAS_SPREAD As Double

Frm83_LM_HARGA_ASAL = 0
Frm83_LM_SPREAD = 0
Frm83_LM_HARGA_LEPAS_SPREAD = 0

If (Frm83.TB10 <> vbNullString And IsNumeric(Frm83.TB10)) Then Frm83_LM_HARGA_ASAL = Frm83.TB10
If (Frm83.TB19 <> vbNullString And IsNumeric(Frm83.TB19)) Then Frm83_LM_SPREAD = Frm83.TB19

Frm83_LM_HARGA_LEPAS_SPREAD = Frm83_LM_HARGA_ASAL - ((Frm83_LM_SPREAD / 100) * Frm83_LM_HARGA_ASAL)

If Frm83.CB7 = 1 Then

    Frm83.TB21 = Format(Frm83_LM_HARGA_ASAL, "#,##0.00") 'Jumlah Harga Belian Setelah Ditolak Spread

ElseIf Frm83.CB8 = 1 Then
    
    Frm83.TB21 = Format(Frm83_LM_HARGA_LEPAS_SPREAD, "#,##0.00") 'Jumlah Harga Belian Setelah Ditolak Spread
    
Else
    
    Frm83.TB21 = Format(Frm83_LM_HARGA_ASAL, "#,##0.00") 'Jumlah Harga Belian Setelah Ditolak Spread
    
End If
End Sub
Sub frm83_harga_belian_lepas_adjust()
'On Error Resume Next
Dim Frm83_LM_HARGA_ASAL As Double
Dim Frm83_LM_ADJUST As Double

Frm83_LM_HARGA_ASAL = 0
Frm83_LM_ADJUST = 0

If (Frm83.TB21 <> vbNullString And IsNumeric(Frm83.TB21)) Then Frm83_LM_HARGA_ASAL = Frm83.TB21
If (Frm83.TB22 <> vbNullString And IsNumeric(Frm83.TB22)) Then Frm83_LM_ADJUST = Frm83.TB22

Frm83.TB20 = Format(Frm83_LM_HARGA_ASAL - Frm83_LM_ADJUST, "#,##0.00") 'Harga belian
End Sub
Sub frm83_flag_barang_baru()
'on error resume next
Frm83.Pic2.Visible = False
Frm83.Frame8.Visible = False

Frm83.CMD1.Visible = False
Frm83.CMD6.Visible = False
Frm83.CMD7.Visible = False
Frm83.CMD12.Visible = False
Frm83.CMD13.Visible = False
Frm83.CMD14.Visible = False

Frm83.CMD20.Visible = True
Frm83.CMD21.Visible = True
Frm83.CMD22.Visible = False
Frm83.CMD23.Visible = False
End Sub
Sub frm83_flag_barang_trade_in()
'on error resume next
Frm83.CMD1.Visible = True
Frm83.CMD6.Visible = False
Frm83.CMD7.Visible = False
Frm83.CMD12.Visible = False
Frm83.CMD13.Visible = False
Frm83.CMD14.Visible = False

Frm83.CMD20.Visible = False
Frm83.CMD21.Visible = False
Frm83.CMD22.Visible = False
Frm83.CMD23.Visible = False
End Sub
Sub frm83_kiraan_cara_bayaran()
'on error resume next
Dim Frm83_LM_TUNAI As Double
Dim Frm83_LM_BANK_IN As Double

Frm83_LM_TUNAI = 0
Frm83_LM_BANK_IN = 0

If (Frm83.TB40 <> vbNullString And IsNumeric(Frm83.TB40)) Then Frm83_LM_TUNAI = Frm83.TB40
If (Frm83.TB41 <> vbNullString And IsNumeric(Frm83.TB41)) Then Frm83_LM_BANK_IN = Frm83.TB41

Frm83.TB42 = Format(Frm83_LM_TUNAI + Frm83_LM_BANK_IN, "#,##0.00") 'Jumlah
End Sub
