Attribute VB_Name = "Module41"
Sub Frm85_Initial_Setting()
'on error resume next
Frm85.Pic2.Left = 120
Frm85.Pic2.Top = 240
Frm85.Pic3.Left = 120
Frm85.Pic3.Top = 240
Frm85.Pic4.Left = 120
Frm85.Pic4.Top = 240
Frm85.Pic5.Left = 120
Frm85.Pic5.Top = 240
Frm85.Pic6.Left = 120
Frm85.Pic6.Top = 240
Frm85.Pic9.Left = 120
Frm85.Pic9.Top = 240
Frm85.Pic10.Left = 120
Frm85.Pic10.Top = 240
Frm85.Pic11.Left = 120
Frm85.Pic11.Top = 240
Frm85.Pic12.Left = 120
Frm85.Pic12.Top = 240
Frm85.Pic13.Left = 120
Frm85.Pic13.Top = 240

Frm85.Pic1.Visible = False
Frm85.Pic2.Visible = False
Frm85.Pic3.Visible = False
Frm85.Pic4.Visible = False
Frm85.Pic5.Visible = False
Frm85.Pic6.Visible = False
Frm85.Pic9.Visible = False
Frm85.Pic10.Visible = False
Frm85.Pic11.Visible = False
Frm85.Pic12.Visible = False
Frm85.Pic13.Visible = False

'Frm85.L85_Text.BackStyle = 0
'Frm85.L86_Text.BackStyle = 0

Frm85.L44_Text = 0
Frm85.L45_Text = "0.00 g"
Frm85.L46_Text = "RM 0.00"
Frm85.L47_Text = 0
Frm85.L48_Text = "0.00 g"
Frm85.L49_Text = "RM 0.00"
Frm85.L50_Text = "RM 0.00"
Frm85.L51_Text = 0
Frm85.L52_Text = "0.00 g"
Frm85.L53_Text = "RM 0.00"
Frm85.L54_Text = 0
Frm85.L55_Text = "0.00 g"
Frm85.L56_Text = "RM 0.00"
Frm85.L57_Text = 0
Frm85.L58_Text = "0.00 g"
Frm85.L59_Text = 0
Frm85.L60_Text = "0.00 g"
Frm85.L61_Text = "RM 0.00"
Frm85.L62_Text = 0
Frm85.L63_Text = "0.00 g"
Frm85.L64_Text = "RM 0.00"
Frm85.L65_Text = 0
Frm85.L66_Text = "0.00 g"
Frm85.L67_Text = "RM 0.00"
Frm85.L68_Text = 0
Frm85.L69_Text = "0.00 g"
Frm85.L70_Text = "RM 0.00"
Frm85.L73_Text = 0
Frm85.L74_Text = "0.00 g"
Frm85.L75_Text = "RM 0.00"
Frm85.L76_Text = 0
Frm85.L77_Text = "0.00 g"
Frm85.L78_Text = "RM 0.00"
Frm85.L83_Text = "0.00 g"
Frm85.L84_Text = "0.00 g"
'Frm85.L85_Text = vbNullString
'Frm85.L86_Text = vbNullString

Frm85.L4_Text = vbNullString
Frm85.L89_Text = vbNullString
Frm85.L90_Text = vbNullString
Frm85.L91_Text = "0.00"
Frm85.L92_Text = "0.00"
Frm85.TB1 = "0.00"

Frm101.L33_Text = 0 '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
End Sub
Sub Frm85_Header_Report_Belian()
'on error resume next
With Frm85.LV2
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear
    
    Frm85.LV2.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "Tarikh Belian", 1500, 2
    .ColumnHeaders.Add 5, , "No. Siri Produk", 1700
    .ColumnHeaders.Add 6, , "Purity", 1500
    .ColumnHeaders.Add 7, , "Kategori Produk", 4000
    .ColumnHeaders.Add 8, , "Supplier", 4000
    .ColumnHeaders.Add 9, , "Berat (g)", 1400, 1
    .ColumnHeaders.Add 10, , "Rate Penerimaan (RM/g)", 2400, 1
    .ColumnHeaders.Add 11, , "Upah (RM)", 1500, 1
    .ColumnHeaders.Add 12, , "Spread (%)", 1200, 1
    .ColumnHeaders.Add 13, , "Harga Selepas Spread (RM)", 2800, 1
    .ColumnHeaders.Add 14, , "Adjustment (RM)", 1700, 1
    .ColumnHeaders.Add 15, , "Harga Belian Termasuk GST (RM)", 3300, 1
    .ColumnHeaders.Add 16, , "Upah Jualan (RM) : Pelanggan", 3200, 1
    .ColumnHeaders.Add 17, , "Upah Jualan (RM) : Ahli Biasa", 3200, 1
    .ColumnHeaders.Add 18, , "Upah Jualan (RM) : Silver", 3200, 1
    .ColumnHeaders.Add 19, , "Upah Jualan (RM) : Gold", 3200, 1
    .ColumnHeaders.Add 20, , "Upah Jualan (RM) : Platinum", 3200, 1
    .ColumnHeaders.Add 21, , "Upah Jualan (RM) : Master Dealer", 0, 1
    .ColumnHeaders.Add 22, , "Tetapan Harga Jualan (RM) : Pelanggan", 3700, 1
    .ColumnHeaders.Add 23, , "Tetapan Harga Jualan (RM) : Ahli", 3500, 1
    .ColumnHeaders.Add 24, , "Tetapan Harga Jualan (RM) : Silver", 3500, 1
    .ColumnHeaders.Add 25, , "Tetapan Harga Jualan (RM) : Gold", 3500, 1
    .ColumnHeaders.Add 26, , "Tetapan Harga Jualan (RM) : Platinum", 3500, 1
    .ColumnHeaders.Add 27, , "Tetapan Harga Jualan (RM) : Master Dealer", 0, 1
    .ColumnHeaders.Add 28, , "Dulang", 1700, 2
    .ColumnHeaders.Add 29, , "Panjang", 1700, 2
    .ColumnHeaders.Add 30, , "Lebar", 1700, 2
    .ColumnHeaders.Add 31, , "Saiz", 1700, 2
    .ColumnHeaders.Add 32, , "No. Invoice", 2000, 0
    .ColumnHeaders.Add 33, , "Code 1", 1700, 0
    .ColumnHeaders.Add 34, , "Code 2", 1700, 0
    .ColumnHeaders.Add 35, , "Cawangan", 2500, 0
    .ColumnHeaders.Add 36, , "Nama Pekerja", 2500, 0
    
End With
End Sub
Sub Frm85_Header_Report_belian_gb()
'on error resume next

'#### Header Report Belian #### - Start
Frm85.MSFlexGrid8.Clear
Frm85.MSFlexGrid8.Rows = 1
Frm85.MSFlexGrid8.RowHeight(0) = 1500
Frm85.MSFlexGrid8.FormatString = "<No.|<No.|<No. ID|<Tarikh Belian|<No. Siri Produk|<Purity|<Kategori Produk|<Supplier|<Berat (g)|<Rate Penerimaan (RM/g)|<Upah (RM)|<Spread (%)|<Harga Selepas Spread (RM)|<Adjustment (RM)|<Harga Belian Termasuk GST (RM)|<Dulang|<Panjang|<Lebar|<Saiz"

Frm85.MSFlexGrid8.ColWidth(0) = 0 'No.
Frm85.MSFlexGrid8.ColWidth(1) = 600 'No.
Frm85.MSFlexGrid8.ColAlignment(1) = 7

Frm85.MSFlexGrid8.ColWidth(2) = 0 'No. ID
Frm85.MSFlexGrid8.ColWidth(3) = 1200 'Tarikh Belian
Frm85.MSFlexGrid8.ColAlignment(3) = 4

Frm85.MSFlexGrid8.ColWidth(4) = 1500 'No. Siri Produk
Frm85.MSFlexGrid8.ColAlignment(4) = 4

Frm85.MSFlexGrid8.ColWidth(5) = 1500 'Purity
Frm85.MSFlexGrid8.ColAlignment(5) = 4

Frm85.MSFlexGrid8.ColWidth(6) = 4000 'Kategori Produk

Frm85.MSFlexGrid8.ColWidth(7) = 4000 'Supplier

Frm85.MSFlexGrid8.ColWidth(8) = 1000 'Berat (g)
Frm85.MSFlexGrid8.ColAlignment(8) = 7

Frm85.MSFlexGrid8.ColWidth(9) = 1100 'Rate Penerimaan (RM/g)
Frm85.MSFlexGrid8.ColAlignment(9) = 7

Frm85.MSFlexGrid8.ColWidth(10) = 1000 'Upah (RM)
Frm85.MSFlexGrid8.ColAlignment(10) = 7

Frm85.MSFlexGrid8.ColWidth(11) = 1000 'Spread (%)
Frm85.MSFlexGrid8.ColAlignment(11) = 7

Frm85.MSFlexGrid8.ColWidth(12) = 1000 'Harga Selepas Spread (RM)
Frm85.MSFlexGrid8.ColAlignment(12) = 7

Frm85.MSFlexGrid8.ColWidth(13) = 900 'Adjustment (RM)
Frm85.MSFlexGrid8.ColAlignment(13) = 7

Frm85.MSFlexGrid8.ColWidth(14) = 1000 'Harga Belian (RM)
Frm85.MSFlexGrid8.ColAlignment(14) = 7

Frm85.MSFlexGrid8.ColWidth(15) = 800 'Dulang
Frm85.MSFlexGrid8.ColAlignment(15) = 4

Frm85.MSFlexGrid8.ColWidth(16) = 800 'Panjang
Frm85.MSFlexGrid8.ColAlignment(16) = 4

Frm85.MSFlexGrid8.ColWidth(17) = 800 'Lebar
Frm85.MSFlexGrid8.ColAlignment(17) = 4

Frm85.MSFlexGrid8.ColWidth(18) = 800 'Saiz
Frm85.MSFlexGrid8.ColAlignment(18) = 4
End Sub
Sub Frm85_Header_Report_buyback_gb()
'on error resume next

'#### Header Report Belian #### - Start
Frm85.MSFlexGrid9.Clear
Frm85.MSFlexGrid9.Rows = 1
Frm85.MSFlexGrid9.RowHeight(0) = 1500
Frm85.MSFlexGrid9.FormatString = "<No.|<No.|<No. ID|<Tarikh Belian|<No. Siri Produk|<Purity|<Kategori Produk|<Supplier|<Berat (g)|<Rate Penerimaan (RM/g)|<Upah (RM)|<Spread (%)|<Harga Selepas Spread (RM)|<Adjustment (RM)|<Harga Belian Tanpa GST (RM)|<Dulang|<Panjang|<Lebar|<Saiz"

Frm85.MSFlexGrid9.ColWidth(0) = 0 'No.
Frm85.MSFlexGrid9.ColWidth(1) = 600 'No.
Frm85.MSFlexGrid9.ColAlignment(1) = 4

Frm85.MSFlexGrid9.ColWidth(2) = 0 'No. ID
Frm85.MSFlexGrid9.ColWidth(3) = 1200 'Tarikh Belian
Frm85.MSFlexGrid9.ColAlignment(3) = 4

Frm85.MSFlexGrid9.ColWidth(4) = 1500 'No. Siri Produk
Frm85.MSFlexGrid9.ColAlignment(4) = 4

Frm85.MSFlexGrid9.ColWidth(5) = 1500 'Purity
Frm85.MSFlexGrid9.ColAlignment(5) = 4

Frm85.MSFlexGrid9.ColWidth(6) = 4000 'Kategori Produk

Frm85.MSFlexGrid9.ColWidth(7) = 4000 'Supplier

Frm85.MSFlexGrid9.ColWidth(8) = 1000 'Berat (g)
Frm85.MSFlexGrid9.ColAlignment(8) = 7

Frm85.MSFlexGrid9.ColWidth(9) = 1100 'Rate Penerimaan (RM/g)
Frm85.MSFlexGrid9.ColAlignment(9) = 7

Frm85.MSFlexGrid9.ColWidth(10) = 1000 'Upah (RM)
Frm85.MSFlexGrid9.ColAlignment(10) = 7

Frm85.MSFlexGrid9.ColWidth(11) = 1000 'Spread (%)
Frm85.MSFlexGrid9.ColAlignment(11) = 7

Frm85.MSFlexGrid9.ColWidth(12) = 1000 'Harga Selepas Spread (RM)
Frm85.MSFlexGrid9.ColAlignment(12) = 7

Frm85.MSFlexGrid9.ColWidth(13) = 900 'Adjustment (RM)
Frm85.MSFlexGrid9.ColAlignment(13) = 7

Frm85.MSFlexGrid9.ColWidth(14) = 1000 'Harga Belian (RM)
Frm85.MSFlexGrid9.ColAlignment(14) = 7

Frm85.MSFlexGrid9.ColWidth(15) = 800 'Dulang
Frm85.MSFlexGrid9.ColAlignment(15) = 5

Frm85.MSFlexGrid9.ColWidth(16) = 800 'Panjang
Frm85.MSFlexGrid9.ColAlignment(16) = 5

Frm85.MSFlexGrid9.ColWidth(17) = 800 'Lebar
Frm85.MSFlexGrid9.ColAlignment(17) = 5

Frm85.MSFlexGrid9.ColWidth(18) = 800 'Saiz
Frm85.MSFlexGrid9.ColAlignment(18) = 5
End Sub
Sub Frm85_report_belian_page()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date
Dim Frm85_LM_TOTAL_PAGE As Double

x = 0
Y = 0
Frm85_PAGE_SIZE = 34
Frm85_LM_TOTAL_PAGE = 0

Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

Frm85.L44_Text = 0
Frm85.L45_Text = Format(0, "#,##0.00 g")
Frm85.L46_Text = "RM " & Format(0, "#,##0.00")

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If
If Frm101.L6_Text = "Semua Kategori Produk" Then
    Frm85_LM_SEARCH_2 = Null
    Frm85_LM_SEARCH_2_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_2 = Frm101.L6_Text
    Frm85_LM_SEARCH_2_LOGIC = "="
End If
If Frm101.L34_Text = "Semua Dulang" Then
    Frm85_LM_SEARCH_3 = Null
    Frm85_LM_SEARCH_3_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_3 = Frm101.L34_Text
    Frm85_LM_SEARCH_3_LOGIC = "="
End If
If Frm101.L37_Text = "Semua Supplier" Then
    Frm85_LM_SEARCH_4 = Null
    Frm85_LM_SEARCH_4_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_4 = Frm101.L37_Text
    Frm85_LM_SEARCH_4_LOGIC = "="
End If
If Frm101.L46_Text = "Semua cawangan" Then

    Frm85_LM_SEARCH_5 = Null
    Frm85_LM_SEARCH_5_LOGIC = "<>"
    Frm85_LM_SEARCH_6 = Null
    Frm85_LM_SEARCH_6_LOGIC = "<>"
    
Else

    Frm85_LM_SEARCH_5 = Frm101.L46_Text
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

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
        End If
    End If
End If

Frm85_LM_PAGE_FOUND = 0

'#### Header Report ###
If Frm101.L9_Text = 0 Then Frm85.L10_Text = "Report Belian Bagi Supplier [" & Frm101.L37_Text & "] , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "]." 'Report Header"
If Frm101.L9_Text = 1 Then Frm85.L10_Text = "Report Belian Bagi Supplier [" & Frm101.L37_Text & "] , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "] Dari " & TM & " Hingga " & TA & "." 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select * from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND (receiving_Status='" & 1 & "' OR receiving_Status='" & 0 & "') order by tarikh_belian ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select * from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND StatusItem <> 0 AND (receiving_Status='" & 1 & "' OR receiving_Status='" & 0 & "') order by tarikh_belian ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    'If rs!receiving_Status = "0" Or rs!receiving_Status = "1" Then
        x = x + 1
        If Frm85_LM_PAGE_FOUND = 0 Then
            If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
                If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                    Frm85.L79_Text = Frm85.L79_Text + 1
                    Frm85_LM_PAGE_FOUND = 1
                ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                    If IsNumeric(Frm85.L79_Text) Then
                        If Frm85.L79_Text <> 1 Then
                            Frm85.L79_Text = Frm85.L79_Text - 1
                            Frm85_LM_PAGE_FOUND = 1
                        End If
                    End If
                End If
            End If
        End If
        Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x

        With Frm85.LV2.ListItems.Add(, , rs!ID)
        
            .ListSubItems.Add , , Y
            
            If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
            
            If Not IsNull(rs!tarikh_belian) Then 'Tarikh Belian
                .ListSubItems.Add , , rs!tarikh_belian
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!no_siri_Produk) Then 'No. Siri Produk
                .ListSubItems.Add , , rs!no_siri_Produk
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!kod_Purity) Then 'Purity
                .ListSubItems.Add , , rs!kod_Purity
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!kategori_Produk) Then 'Kategori Produk
                .ListSubItems.Add , , rs!kategori_Produk
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!nama_Supplier) Then 'Nama Supplier
                .ListSubItems.Add , , rs!nama_Supplier
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!Berat) Then 'Berat (g)
                .ListSubItems.Add , , Format(rs!Berat, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!kos_Belian_Gram) Then 'Rate Penerimaan (RM/g)
                .ListSubItems.Add , , Format(rs!kos_Belian_Gram, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!UPAH) Then 'Upah (RM)
                .ListSubItems.Add , , Format(rs!UPAH, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!SpreadValue) Then 'Spread (%)
                .ListSubItems.Add , , Format(rs!SpreadValue, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!harga_lepas_spread) Then 'Harga Selepas Spread (RM)
                .ListSubItems.Add , , Format(rs!harga_lepas_spread, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If

            If Not IsNull(rs!adjustment) Then 'Adjustment (RM)
                .ListSubItems.Add , , Format(rs!adjustment, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!harga_item) Then 'Harga Belian (RM) : Tidak Campur Cukai GST
                .ListSubItems.Add , , Format(rs!harga_item, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!Upah_Jualan) Then 'Upah Jualan (RM) : Pelanggan
                .ListSubItems.Add , , Format(rs!Upah_Jualan, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If

            If Not IsNull(rs!Upah_Member) Then 'Upah Jualan (RM) : Ahli Biasa
                .ListSubItems.Add , , Format(rs!Upah_Member, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If

            If Not IsNull(rs!Upah_Pengedar) Then 'Upah Jualan (RM) : Silver
                .ListSubItems.Add , , Format(rs!Upah_Pengedar, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!Upah_RAF) Then 'Upah Jualan (RM) : Gold
                .ListSubItems.Add , , Format(rs!Upah_RAF, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!upah_normal_dealer) Then 'Upah Jualan (RM) : Platinum
                .ListSubItems.Add , , Format(rs!upah_normal_dealer, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!upah_master_dealer) Then 'Upah Jualan (RM) : Master Dealer
                .ListSubItems.Add , , Format(rs!upah_master_dealer, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!code_Supplier) Then 'Tetapan Harga Jualan (RM) : Pelanggan
                .ListSubItems.Add , , Format(rs!code_Supplier, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!HargaJualan_Member) Then 'Tetapan Harga Jualan (RM) : Pelanggan
                .ListSubItems.Add , , Format(rs!HargaJualan_Member, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!HargaJualan_Pengedar) Then 'Tetapan Harga Jualan (RM) : Pengedar
                .ListSubItems.Add , , Format(rs!HargaJualan_Pengedar, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!HargaJualan_RAF) Then 'Tetapan Harga Jualan (RM) : RAF
                .ListSubItems.Add , , Format(rs!HargaJualan_RAF, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!hargajualan_normal_dealer) Then 'Tetapan Harga Jualan (RM) : Normal Dealer
                .ListSubItems.Add , , Format(rs!hargajualan_normal_dealer, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If

            If Not IsNull(rs!hargajualan_master_dealer) Then 'Tetapan Harga Jualan (RM) : Master Dealer
                .ListSubItems.Add , , Format(rs!hargajualan_master_dealer, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
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
            
            If Not IsNull(rs!dimension_Saiz) Then 'Tebal
                .ListSubItems.Add , , rs!dimension_Saiz
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!bill_No_Belian) Then 'No. Invoice Supplier
                .ListSubItems.Add , , rs!bill_No_Belian
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
            
            If Not IsNull(rs!cawangan) Then 'Cawangan
                .ListSubItems.Add , , rs!cawangan
            Else
                .ListSubItems.Add , , ""
            End If

            If Not IsNull(rs!nama_pekerja) Then 'Pekerja
                .ListSubItems.Add , , rs!nama_pekerja
            Else
                .ListSubItems.Add , , ""
            End If
            
        End With
    'End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select COUNT(ID),SUM(Berat),SUM(harga_item) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND (receiving_Status='" & 1 & "' OR receiving_Status='" & 0 & "')", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select COUNT(ID),SUM(Berat),SUM(harga_item) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 1 & "' OR receiving_Status='" & 0 & "')", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

If Not IsNull(rs(0)) Then Frm85.L44_Text = Format(rs(0), "#,##0")
If Not IsNull(rs(1)) Then Frm85.L45_Text = Format(rs(1), "#,##0.00 g")
If Not IsNull(rs(2)) Then Frm85.L46_Text = "RM " & Format(rs(2), "#,##0.00")

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    Frm85.Pic2.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    MsgBox "Tiada Rekod Belian Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_report_belian_gb_page()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date
Dim Frm85_LM_TOTAL_PAGE As Double

x = 0
Y = 0
Frm85_PAGE_SIZE = 34
Frm85_LM_TOTAL_PAGE = 0

Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If
If Frm101.L6_Text = "Semua Kategori Produk" Then
    Frm85_LM_SEARCH_2 = Null
    Frm85_LM_SEARCH_2_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_2 = Frm101.L6_Text
    Frm85_LM_SEARCH_2_LOGIC = "="
End If
If Frm101.L34_Text = "Semua Dulang" Then
    Frm85_LM_SEARCH_3 = Null
    Frm85_LM_SEARCH_3_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_3 = Frm101.L34_Text
    Frm85_LM_SEARCH_3_LOGIC = "="
End If
If Frm101.L37_Text = "Semua Supplier" Then
    Frm85_LM_SEARCH_4 = Null
    Frm85_LM_SEARCH_4_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_4 = Frm101.L37_Text
    Frm85_LM_SEARCH_4_LOGIC = "="
End If

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
        End If
    End If
End If

Frm85_LM_PAGE_FOUND = 0

'#### Header Report ###
If Frm101.L9_Text = 0 Then Frm85.L71_Text = "Report Belian Gold Bar Bagi Supplier [" & Frm101.L37_Text & "] , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "]" 'Report Header"
If Frm101.L9_Text = 1 Then Frm85.L71_Text = "Report Belian Gold Bar Bagi Supplier [" & Frm101.L37_Text & "] , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] Dari " & TM & " Hingga " & TA 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select * from Data_Database where nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND receiving_Status='" & 4 & "' order by tarikh_belian ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select * from Data_Database where nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND receiving_Status='" & 4 & "' order by tarikh_belian ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm85_LM_PAGE_FOUND = 0 Then
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm85.L79_Text = Frm85.L79_Text + 1
                Frm85_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm85.L79_Text) Then
                    If Frm85.L79_Text <> 1 Then
                        Frm85.L79_Text = Frm85.L79_Text - 1
                        Frm85_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x
    Frm85.MSFlexGrid8.Rows = x + 1
    Frm85.MSFlexGrid8.TextMatrix(x, 0) = x 'No.
    Frm85.MSFlexGrid8.TextMatrix(x, 1) = Y 'No.
    Frm85.MSFlexGrid8.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh_belian) Then Frm85.MSFlexGrid8.TextMatrix(x, 3) = rs!tarikh_belian 'Tarikh Belian
    If Not IsNull(rs!no_siri_Produk) Then Frm85.MSFlexGrid8.TextMatrix(x, 4) = rs!no_siri_Produk 'No. Siri Produk
    If Not IsNull(rs!kod_Purity) Then Frm85.MSFlexGrid8.TextMatrix(x, 5) = rs!kod_Purity 'Purity
    If Not IsNull(rs!kategori_Produk) Then Frm85.MSFlexGrid8.TextMatrix(x, 6) = rs!kategori_Produk 'Kategori Produk
    If Not IsNull(rs!nama_Supplier) Then Frm85.MSFlexGrid8.TextMatrix(x, 7) = rs!nama_Supplier 'Nama Supplier
    If Not IsNull(rs!Berat) Then
        Frm85.MSFlexGrid8.TextMatrix(x, 8) = Format(rs!Berat, "#,##0.00") 'Berat (g)
        If IsNumeric(rs!Berat) Then Frm85_LM_BERAT = Frm85_LM_BERAT + rs!Berat 'Total Berat (g)
    End If
    If Not IsNull(rs!kos_Belian_Gram) Then Frm85.MSFlexGrid8.TextMatrix(x, 9) = Format(rs!kos_Belian_Gram, "#,##0.00") 'Rate Penerimaan (RM/g)
    If Not IsNull(rs!UPAH) Then Frm85.MSFlexGrid8.TextMatrix(x, 10) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
    If Not IsNull(rs!SpreadValue) Then Frm85.MSFlexGrid8.TextMatrix(x, 11) = rs!SpreadValue 'Spread (%)
    If Not IsNull(rs!harga_lepas_spread) Then Frm85.MSFlexGrid8.TextMatrix(x, 12) = Format(rs!harga_lepas_spread, "#,##0.00") 'Harga Selepas Spread (RM)
    If Not IsNull(rs!adjustment) Then Frm85.MSFlexGrid8.TextMatrix(x, 13) = Format(rs!adjustment, "#,##0.00") 'Adjustment (RM)
    If Not IsNull(rs!harga_item) Then
        Frm85.MSFlexGrid8.TextMatrix(x, 14) = Format(rs!harga_item, "#,##0.00") 'Harga Belian (RM) : Tidak Campur Cukai GST
        If IsNumeric(rs!harga_item) Then Frm85_LM_HARGA = Frm85_LM_HARGA + rs!harga_item 'Total Harga Belian (RM) : Tidak Campur Cukai GST
    End If
    If Not IsNull(rs!dulang) Then Frm85.MSFlexGrid8.TextMatrix(x, 15) = rs!dulang 'Dulang
    If Not IsNull(rs!dimension_Panjang) Then Frm85.MSFlexGrid8.TextMatrix(x, 16) = rs!dimension_Panjang 'Panjang
    If Not IsNull(rs!dimension_Lebar) Then Frm85.MSFlexGrid8.TextMatrix(x, 17) = rs!dimension_Lebar 'Lebar
    If Not IsNull(rs!dimension_Saiz) Then Frm85.MSFlexGrid8.TextMatrix(x, 18) = rs!dimension_Saiz 'Tebal
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm85.L65_Text = x 'Total Barang
Frm85.L66_Text = Format(Frm85_LM_BERAT, "#,##0.00 g") 'Total Berat
Frm85.L67_Text = "RM " & Format(Frm85_LM_HARGA, "#,##0.00") 'Total Harga Belian

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select COUNT(ID) from Data_Database where nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND receiving_Status='" & 4 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select COUNT(ID) from Data_Database where nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND receiving_Status='" & 4 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

'#### Jumlah Bilangan Barang Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select COUNT(no_siri_Produk) from Data_Database where nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND receiving_Status='" & 4 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select COUNT(no_siri_Produk) from Data_Database where nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND receiving_Status='" & 4 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L68_Text = rs(0)
Else
    Frm85.L68_Text = 0
End If

rs.Close
Set rs = Nothing
'#### Jumlah Bilangan Barang Keseluruhan #### - End

'#### Jumlah Berat Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select SUM(Berat) from Data_Database where nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND receiving_Status='" & 4 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select SUM(Berat) from Data_Database where nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND receiving_Status='" & 4 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L69_Text = Format(rs(0), "#,##0.00 g")
Else
    Frm85.L69_Text = "0.00 g"
End If

rs.Close
Set rs = Nothing
'#### Jumlah Berat Keseluruhan #### - End

'#### Jumlah Modal Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select SUM(harga_item) from Data_Database where nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND receiving_Status='" & 4 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select SUM(harga_item) from Data_Database where nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND receiving_Status='" & 4 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L70_Text = "RM " & Format(rs(0), "#,##0.00")
Else
    Frm85.L70_Text = "RM 0.00"
End If

rs.Close
Set rs = Nothing
'#### Jumlah Modal Keseluruhan #### - End

If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    
    Frm85.Pic11.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    
    MsgBox "Tiada Rekod Belian Gold Bar Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_report_buyback_gb_page()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date
Dim Frm85_LM_TOTAL_PAGE As Double

x = 0
Y = 0
Frm85_PAGE_SIZE = 34
Frm85_LM_TOTAL_PAGE = 0

Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If
If Frm101.L6_Text = "Semua Kategori Produk" Then
    Frm85_LM_SEARCH_2 = Null
    Frm85_LM_SEARCH_2_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_2 = Frm101.L6_Text
    Frm85_LM_SEARCH_2_LOGIC = "="
End If
If Frm101.L34_Text = "Semua Dulang" Then
    Frm85_LM_SEARCH_3 = Null
    Frm85_LM_SEARCH_3_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_3 = Frm101.L34_Text
    Frm85_LM_SEARCH_3_LOGIC = "="
End If

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
        End If
    End If
End If

Frm85_LM_PAGE_FOUND = 0

'#### Header Report ###
If Frm101.L9_Text = 0 Then Frm85.L72_Text = "Report Buyback / Trade In Gold Bar Bagi Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "]" 'Report Header"
If Frm101.L9_Text = 1 Then Frm85.L72_Text = "Report Buyback / Trade In Gold Bar Bagi Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] Dari " & TM & " Hingga " & TA 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select * from Data_Database where kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND receiving_Status='" & 5 & "' order by tarikh_belian ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select * from Data_Database where kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND receiving_Status='" & 5 & "' order by tarikh_belian ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm85_LM_PAGE_FOUND = 0 Then
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm85.L79_Text = Frm85.L79_Text + 1
                Frm85_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm85.L79_Text) Then
                    If Frm85.L79_Text <> 1 Then
                        Frm85.L79_Text = Frm85.L79_Text - 1
                        Frm85_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x
    Frm85.MSFlexGrid9.Rows = x + 1
    Frm85.MSFlexGrid9.TextMatrix(x, 0) = x 'No.
    Frm85.MSFlexGrid9.TextMatrix(x, 1) = Y 'No.
    Frm85.MSFlexGrid9.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh_belian) Then Frm85.MSFlexGrid9.TextMatrix(x, 3) = rs!tarikh_belian 'Tarikh Belian
    If Not IsNull(rs!no_siri_Produk) Then Frm85.MSFlexGrid9.TextMatrix(x, 4) = rs!no_siri_Produk 'No. Siri Produk
    If Not IsNull(rs!kod_Purity) Then Frm85.MSFlexGrid9.TextMatrix(x, 5) = rs!kod_Purity 'Purity
    If Not IsNull(rs!kategori_Produk) Then Frm85.MSFlexGrid9.TextMatrix(x, 6) = rs!kategori_Produk 'Kategori Produk
    If Not IsNull(rs!nama_Supplier) Then Frm85.MSFlexGrid9.TextMatrix(x, 7) = rs!nama_Supplier 'Nama Supplier
    If Not IsNull(rs!Berat) Then
        Frm85.MSFlexGrid9.TextMatrix(x, 8) = Format(rs!Berat, "#,##0.00") 'Berat (g)
        If IsNumeric(rs!Berat) Then Frm85_LM_BERAT = Frm85_LM_BERAT + rs!Berat 'Total Berat (g)
    End If
    If Not IsNull(rs!kos_Belian_Gram) Then Frm85.MSFlexGrid9.TextMatrix(x, 9) = Format(rs!kos_Belian_Gram, "#,##0.00") 'Rate Penerimaan (RM/g)
    If Not IsNull(rs!UPAH) Then Frm85.MSFlexGrid9.TextMatrix(x, 10) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
    If Not IsNull(rs!SpreadValue) Then Frm85.MSFlexGrid9.TextMatrix(x, 11) = rs!SpreadValue 'Spread (%)
    If Not IsNull(rs!harga_lepas_spread) Then Frm85.MSFlexGrid9.TextMatrix(x, 12) = Format(rs!harga_lepas_spread, "#,##0.00") 'Harga Selepas Spread (RM)
    If Not IsNull(rs!adjustment) Then Frm85.MSFlexGrid9.TextMatrix(x, 13) = Format(rs!adjustment, "#,##0.00") 'Adjustment (RM)
    If Not IsNull(rs!harga_item) Then
        Frm85.MSFlexGrid9.TextMatrix(x, 14) = Format(rs!harga_item, "#,##0.00") 'Harga Belian (RM) : Tidak Campur Cukai GST
        If IsNumeric(rs!harga_item) Then Frm85_LM_HARGA = Frm85_LM_HARGA + rs!harga_item 'Total Harga Belian (RM) : Tidak Campur Cukai GST
    End If
    If Not IsNull(rs!dulang) Then Frm85.MSFlexGrid9.TextMatrix(x, 15) = rs!dulang 'Dulang
    If Not IsNull(rs!dimension_Panjang) Then Frm85.MSFlexGrid9.TextMatrix(x, 16) = rs!dimension_Panjang 'Panjang
    If Not IsNull(rs!dimension_Lebar) Then Frm85.MSFlexGrid9.TextMatrix(x, 17) = rs!dimension_Lebar 'Lebar
    If Not IsNull(rs!dimension_Saiz) Then Frm85.MSFlexGrid9.TextMatrix(x, 18) = rs!dimension_Saiz 'Tebal
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm85.L73_Text = x 'Total Barang
Frm85.L74_Text = Format(Frm85_LM_BERAT, "#,##0.00 g") 'Total Berat
Frm85.L75_Text = "RM " & Format(Frm85_LM_HARGA, "#,##0.00") 'Total Harga Belian

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select COUNT(ID) from Data_Database where kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select COUNT(ID) from Data_Database where kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

'#### Jumlah Bilangan Barang Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select COUNT(no_siri_Produk) from Data_Database where kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select COUNT(no_siri_Produk) from Data_Database where kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L76_Text = rs(0)
Else
    Frm85.L76_Text = 0
End If

rs.Close
Set rs = Nothing
'#### Jumlah Bilangan Barang Keseluruhan #### - End

'#### Jumlah Berat Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select SUM(Berat) from Data_Database where kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select SUM(Berat) from Data_Database where kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L77_Text = Format(rs(0), "#,##0.00 g")
Else
    Frm85.L77_Text = "0.00 g"
End If

rs.Close
Set rs = Nothing
'#### Jumlah Berat Keseluruhan #### - End

'#### Jumlah Modal Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select SUM(harga_item) from Data_Database where kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select SUM(harga_item) from Data_Database where kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L78_Text = "RM " & Format(rs(0), "#,##0.00")
Else
    Frm85.L78_Text = "RM 0.00"
End If

rs.Close
Set rs = Nothing
'#### Jumlah Modal Keseluruhan #### - End

If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    
    Frm85.Pic12.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    
    MsgBox "Tiada Rekod Buyback / Trade In Gold Bar Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_Report_Jualan_page()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim Frm85_LM_UNTUNG As Double
Dim Frm85_LM_UNTUNG2 As Double
Dim TA As Date
Dim TM As Date
Dim Frm85_LM_TOTAL_PAGE As Double

Frm85_PAGE_SIZE = 35
Frm85_LM_TOTAL_PAGE = 0
x = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0
Frm85_LM_UNTUNG = 0
Frm85_LM_UNTUNG2 = 0

Frm85.L47_Text = Format(0, "#,##0")
Frm85.L48_Text = Format(0, "#,##0.00 g")
Frm85.L49_Text = "RM " & Format(0, "#,##0.00")
Frm85.L50_Text = "RM " & Format(0, "#,##0.00")
Frm85.L88_Text = "RM " & Format(0, "#,##0.00")

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If
If Frm101.L6_Text = "Semua Kategori Produk" Then
    Frm85_LM_SEARCH_2 = Null
    Frm85_LM_SEARCH_2_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_2 = Frm101.L6_Text
    Frm85_LM_SEARCH_2_LOGIC = "="
End If
If Frm101.L34_Text = "Semua Dulang" Then
    Frm85_LM_SEARCH_3 = Null
    Frm85_LM_SEARCH_3_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_3 = Frm101.L34_Text
    Frm85_LM_SEARCH_3_LOGIC = "="
End If
If Frm101.L44_Text = 2 Then '0 : Jualan biasa , 1 : Jualan kepada agen , 2 : Semua jenis jualan
    Frm85_LM_SEARCH_4 = Null
    Frm85_LM_SEARCH_4_LOGIC = "<>"
ElseIf Frm101.L44_Text = 0 Then '0 : Jualan biasa , 1 : Jualan kepada agen , 2 : Semua jenis jualan
    Frm85_LM_SEARCH_4 = 0
    Frm85_LM_SEARCH_4_LOGIC = "="
ElseIf Frm101.L44_Text = 1 Then '0 : Jualan biasa , 1 : Jualan kepada agen , 2 : Semua jenis jualan
    Frm85_LM_SEARCH_4 = 1
    Frm85_LM_SEARCH_4_LOGIC = "="
End If
If Frm101.L45_Text = "Kedai & Online" Then
    Frm85_LM_SEARCH_5 = Null
    Frm85_LM_SEARCH_5_LOGIC = "<>"
ElseIf Frm101.L45_Text = "Kedai Sahaja" Then
    Frm85_LM_SEARCH_5 = 0
    Frm85_LM_SEARCH_5_LOGIC = "="
ElseIf Frm101.L45_Text = "Online Sahaja" Then
    Frm85_LM_SEARCH_5 = 1
    Frm85_LM_SEARCH_5_LOGIC = "="
End If
If Frm101.L46_Text = "Semua cawangan" Then
    Frm85_SEARCH_8 = Null
    Frm85_SEARCH_8_LOGIC = "<>"
    Frm85_SEARCH_9 = Null
    Frm85_SEARCH_9_LOGIC = "<>"
Else
    Frm85_SEARCH_8 = Frm101.L46_Text
    Frm85_SEARCH_8_LOGIC = "="
    Frm85_SEARCH_9 = "HQ"
    Frm85_SEARCH_9_LOGIC = "="
End If
If Frm101.L47_Text = "Semua" Then
    Frm85_LM_SEARCH_20 = Null
    Frm85_LM_SEARCH_20_LOGIC = "<>"
Else
    If InStr(1, Frm101.L47_Text, "  |  ") <> 0 Then
        Frm84_LM_EMP_NO = Split(Frm101.L47_Text, "  |  ")(1)
        Frm84_LM_EMP_NAMA = Split(Frm101.L47_Text, "  |  ")(0)
    End If
    Frm85_LM_SEARCH_20 = Frm84_LM_EMP_NAMA
    Frm85_LM_SEARCH_20_LOGIC = "="
End If

user_level = MDI_frm1.L4_Text

LM_INVOICE_RASMI = 0

If user_level = "Guest/User" Then
    Frm85_LM_SEARCH_6 = 1
    Frm85_LM_SEARCH_6_LOGIC = "="
    
    Frm85_LM_SEARCH_7 = 1
    Frm85_LM_SEARCH_7_LOGIC = "="
    
    LM_INVOICE_RASMI = 1
Else
    Frm85_LM_SEARCH_6 = 0
    Frm85_LM_SEARCH_6_LOGIC = "="
    
    Frm85_LM_SEARCH_7 = 1
    Frm85_LM_SEARCH_7_LOGIC = "="
End If

If user_level = "Administration" Then
    Frm85_LM_SEARCH_10 = 1
    Frm85_LM_SEARCH_10_LOGIC = "="
    
    Frm85_LM_SEARCH_11 = 1
    Frm85_LM_SEARCH_11_LOGIC = "="
Else
    Frm85_LM_SEARCH_10 = 0
    Frm85_LM_SEARCH_10_LOGIC = "="
    
    Frm85_LM_SEARCH_11 = 1
    Frm85_LM_SEARCH_11_LOGIC = "="
End If

If G_JENIS_JUALAN = "Barang Baru Sahaja" Then

    Frm85_LM_SEARCH_12 = 0
    Frm85_LM_SEARCH_12_LOGIC = "="
    
    Frm85_LM_SEARCH_13 = 0
    Frm85_LM_SEARCH_13_LOGIC = "="
    
ElseIf G_JENIS_JUALAN = "Barang Trade In Sahaja" Then

    Frm85_LM_SEARCH_12 = 1
    Frm85_LM_SEARCH_12_LOGIC = "="
    
    Frm85_LM_SEARCH_13 = 1
    Frm85_LM_SEARCH_13_LOGIC = "="
    
ElseIf G_JENIS_JUALAN = "Barang Baru Dan Barang Trade In" Then

    Frm85_LM_SEARCH_12 = 0
    Frm85_LM_SEARCH_12_LOGIC = "="
    
    Frm85_LM_SEARCH_13 = 1
    Frm85_LM_SEARCH_13_LOGIC = "="
    
End If

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
        End If
    End If
End If

Frm85_LM_PAGE_FOUND = 0

'#### Header Report ###
If Frm101.L9_Text = 0 Then Frm85.L14_Text = "Report Jualan Bagi " & G_JENIS_JUALAN & " , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Jenis Jualan [" & Frm101.CBB5 & "] , Jualan Secara [" & Frm101.L45_Text & "] , Cawangan [" & Frm101.L46_Text & "]." 'Report Header
If Frm101.L9_Text = 1 Then Frm85.L14_Text = "Report Jualan Bagi " & G_JENIS_JUALAN & " , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Jenis Jualan [" & Frm101.CBB5 & "] , Jualan Secara [" & Frm101.L45_Text & "] , Cawangan [" & Frm101.L46_Text & "] Dari " & TM & " Hingga " & TA & "." 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select * from 23_senarai_jualan where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (baru_or_ti " & Frm85_LM_SEARCH_12_LOGIC & "'" & Frm85_LM_SEARCH_12 & "' OR baru_or_ti " & Frm85_LM_SEARCH_13_LOGIC & "'" & Frm85_LM_SEARCH_13 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND " _
& "dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND jenis_jualan " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND jualan_online " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by no_resit ASC , tarikh ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select * from 23_senarai_jualan where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (baru_or_ti " & Frm85_LM_SEARCH_12_LOGIC & "'" & Frm85_LM_SEARCH_12 & "' OR baru_or_ti " & Frm85_LM_SEARCH_13_LOGIC & "'" & Frm85_LM_SEARCH_13 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND " _
& "dulang " & Frm85_LM_SEARCH_3_LOGIC & " '" & Frm85_LM_SEARCH_3 & "' AND jenis_jualan " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND jualan_online " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by no_resit ASC , tarikh ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm85_LM_PAGE_FOUND = 0 Then
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm85.L79_Text = Frm85.L79_Text + 1
                Frm85_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm85.L79_Text) Then
                    If Frm85.L79_Text <> 1 Then
                        Frm85.L79_Text = Frm85.L79_Text - 1
                        Frm85_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x
    
    With Frm85.LV1.ListItems.Add(, , rs!ID)
     
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!tarikh) Then 'Tarikh Jualan
            .ListSubItems.Add , , rs!tarikh
        Else
            .ListSubItems.Add , , ""
        End If
        
        If LM_INVOICE_RASMI = 0 Then
        
            If Not IsNull(rs!no_resit) Then .ListSubItems.Add , , rs!no_resit
            If Not IsNull(rs!no_invoice_r) Then .ListSubItems.Add , , rs!no_invoice_r
            
        Else
            .ListSubItems.Add , , ""
        End If
        
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
        
        If Not IsNull(rs!berat_jualan) Then  'Berat Jualan (g)
            .ListSubItems.Add , , Format(rs!berat_jualan, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_Semasa) Then 'Harga Semasa (RM/g)
            .ListSubItems.Add , , Format(rs!harga_Semasa, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!UPAH) Then 'Upah (RM)
            .ListSubItems.Add , , Format(rs!UPAH, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_asal) Then 'Harga Asal (RM)
            .ListSubItems.Add , , Format(rs!harga_asal, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!diskaun) Then 'Diskaun (%)
            .ListSubItems.Add , , Format(rs!diskaun, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!harga_lepas_diskaun) Then 'Harga Selepas Diskaun (RM)
            .ListSubItems.Add , , Format(rs!harga_lepas_diskaun, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!adjustment) Then 'Adjustment (RM)
            .ListSubItems.Add , , Format(rs!adjustment, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_jualan_dengan_gst) Then 'Harga Jualan (RM)
            .ListSubItems.Add , , Format(rs!harga_jualan_dengan_gst, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!gst_ari_nashi) Then
        
            If rs!gst_ari_nashi = "ZR (L)" Then '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                .ListSubItems.Add , , "ZR (L)"
            ElseIf rs!gst_ari_nashi = "SR" Then
                .ListSubItems.Add , , "SR"
            End If
            
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!jumlah_gst) Then 'Jumlah GST (RM)
            .ListSubItems.Add , , Format(rs!jumlah_gst, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!cawangan) Then
            .ListSubItems.Add , , rs!cawangan
        Else
            .ListSubItems.Add , , ""
        End If
        If Not IsNull(rs!nama_pekerja) Then
            .ListSubItems.Add , , rs!nama_pekerja
        Else
            .ListSubItems.Add , , ""
        End If
        If Not IsNull(rs!dulang) Then
            .ListSubItems.Add , , rs!dulang
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
If Frm101.L9_Text = 0 Then rs.Open "select COUNT(ID),SUM(Berat_Jualan),SUM(harga_jualan_dengan_gst),SUM(untung),SUM(untung2) from 23_senarai_jualan where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (baru_or_ti " & Frm85_LM_SEARCH_12_LOGIC & "'" & Frm85_LM_SEARCH_12 & "' OR baru_or_ti " & Frm85_LM_SEARCH_13_LOGIC & "'" & Frm85_LM_SEARCH_13 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND " _
& "kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND jenis_jualan " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND jualan_online " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select COUNT(ID),SUM(Berat_Jualan),SUM(harga_jualan_dengan_gst),SUM(untung),SUM(untung2) from 23_senarai_jualan where (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (baru_or_ti " & Frm85_LM_SEARCH_12_LOGIC & "'" & Frm85_LM_SEARCH_12 & "' OR baru_or_ti " & Frm85_LM_SEARCH_13_LOGIC & "'" & Frm85_LM_SEARCH_13 & "') AND " _
& "kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND jenis_jualan " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND jualan_online " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

If Not IsNull(rs(0)) Then Frm85.L47_Text = Format(rs(0), "#,##0")
If Not IsNull(rs(1)) Then Frm85.L48_Text = Format(rs(1), "#,##0.00 g")
If Not IsNull(rs(2)) Then Frm85.L49_Text = "RM " & Format(rs(2), "#,##0.00")
If Not IsNull(rs(3)) Then Frm85.L50_Text = "RM " & Format(rs(3), "#,##0.00")
If Not IsNull(rs(4)) Then Frm85.L88_Text = "RM " & Format(rs(4), "#,##0.00")

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    
    Frm85.Pic3.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else

    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    
    MsgBox "Tiada Rekod Jualan Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_Header_Report_Jualan()
'on error resume next
With Frm85.LV1
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear

    Frm85.LV1.ListItems.Clear
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "Tarikh Jualan", 1500, 2
    .ColumnHeaders.Add 5, , "No. Invoice", 2200
    .ColumnHeaders.Add 6, , "No. Siri Produk", 2000
    .ColumnHeaders.Add 7, , "Kategori Produk", 4000
    .ColumnHeaders.Add 8, , "Purity", 1200
    .ColumnHeaders.Add 9, , "Berat Asal (g)", 1400, 1
    .ColumnHeaders.Add 10, , "Berat Jualan (g)", 1600, 1
    .ColumnHeaders.Add 11, , "Harga Semasa (RM/g)", 2100, 1
    .ColumnHeaders.Add 12, , "Upah (RM)", 1200, 1
    .ColumnHeaders.Add 13, , "Harga Asal (RM)", 1700, 1
    .ColumnHeaders.Add 14, , "Diskaun (%)", 1300, 1
    .ColumnHeaders.Add 15, , "Harga Selepas Diskaun (RM)", 3000, 1
    .ColumnHeaders.Add 16, , "Adjustment (RM)", 1800, 1
    .ColumnHeaders.Add 17, , "Harga Jualan Termasuk GST(RM)", 3100, 1
    .ColumnHeaders.Add 18, , "Jenis GST", 1300, 2
    .ColumnHeaders.Add 19, , "Jumlah GST(RM)", 1700, 1
    .ColumnHeaders.Add 20, , "Cawangan", 2500
    .ColumnHeaders.Add 21, , "Nama Pekerja", 2000
    .ColumnHeaders.Add 22, , "Dulang", 1500

End With
End Sub
Sub Frm85_Header_Report_Buyback()
'on error resume next
With Frm85.LV3
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear
    
    Frm85.LV3.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "Tarikh Belian", 1500, 2
    .ColumnHeaders.Add 5, , "No. Siri Produk", 1700
    .ColumnHeaders.Add 6, , "Purity", 1500
    .ColumnHeaders.Add 7, , "Kategori Produk", 4000
    .ColumnHeaders.Add 8, , "Supplier", 4000
    .ColumnHeaders.Add 9, , "Berat (g)", 1400, 1
    .ColumnHeaders.Add 10, , "Rate Penerimaan (RM/g)", 2400, 1
    .ColumnHeaders.Add 11, , "Upah (RM)", 1500, 1
    .ColumnHeaders.Add 12, , "Spread (%)", 1200, 1
    .ColumnHeaders.Add 13, , "Harga Selepas Spread (RM)", 2800, 1
    .ColumnHeaders.Add 14, , "Adjustment (RM)", 1700, 1
    .ColumnHeaders.Add 15, , "Harga Belian Termasuk GST (RM)", 3300, 1
    .ColumnHeaders.Add 16, , "Upah Jualan (RM) : Pelanggan", 3200, 1
    .ColumnHeaders.Add 17, , "Upah Jualan (RM) : Ahli Biasa", 3200, 1
    .ColumnHeaders.Add 18, , "Upah Jualan (RM) : Silver", 3200, 1
    .ColumnHeaders.Add 19, , "Upah Jualan (RM) : Gold", 3200, 1
    .ColumnHeaders.Add 20, , "Upah Jualan (RM) : Platinum", 3200, 1
    .ColumnHeaders.Add 21, , "Upah Jualan (RM) : Master Dealer", 0, 1
    .ColumnHeaders.Add 22, , "Tetapan Harga Jualan (RM) : Pelanggan", 3700, 1
    .ColumnHeaders.Add 23, , "Tetapan Harga Jualan (RM) : Ahli", 3500, 1
    .ColumnHeaders.Add 24, , "Tetapan Harga Jualan (RM) : Silver", 3500, 1
    .ColumnHeaders.Add 25, , "Tetapan Harga Jualan (RM) : Gold", 3500, 1
    .ColumnHeaders.Add 26, , "Tetapan Harga Jualan (RM) : Platinum", 3500, 1
    .ColumnHeaders.Add 27, , "Tetapan Harga Jualan (RM) : Master Dealer", 0, 1
    .ColumnHeaders.Add 28, , "Dulang", 1700, 2
    .ColumnHeaders.Add 29, , "Panjang", 1700, 2
    .ColumnHeaders.Add 30, , "Lebar", 1700, 2
    .ColumnHeaders.Add 31, , "Saiz", 1700, 2
    .ColumnHeaders.Add 32, , "No. Invoice", 2000, 0
    .ColumnHeaders.Add 33, , "Code 1", 1700, 0
    .ColumnHeaders.Add 34, , "Code 2", 1700, 0
    .ColumnHeaders.Add 35, , "Cawangan", 2500, 0
    .ColumnHeaders.Add 36, , "No. Voucher", 2500, 0
    
End With
End Sub
Sub Frm85_report_buyback_page()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date
Dim Frm85_LM_TOTAL_PAGE As Double

Frm85_LM_TOTAL_PAGE = 0
Frm85_PAGE_SIZE = 34
x = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

Frm85.L51_Text = Format(0, "#,##0")
Frm85.L52_Text = Format(0, "#,##0.00 g")
Frm85.L53_Text = "RM " & Format(0, "#,##0.00")

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If
If Frm101.L6_Text = "Semua Kategori Produk" Then
    Frm85_LM_SEARCH_2 = Null
    Frm85_LM_SEARCH_2_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_2 = Frm101.L6_Text
    Frm85_LM_SEARCH_2_LOGIC = "="
End If
If Frm101.L34_Text = "Semua Dulang" Then
    Frm85_LM_SEARCH_3 = Null
    Frm85_LM_SEARCH_3_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_3 = Frm101.L34_Text
    Frm85_LM_SEARCH_3_LOGIC = "="
End If
If Frm101.L46_Text = "Semua cawangan" Then

    Frm85_LM_SEARCH_5 = Null
    Frm85_LM_SEARCH_5_LOGIC = "<>"
    Frm85_LM_SEARCH_6 = Null
    Frm85_LM_SEARCH_6_LOGIC = "<>"
    
Else

    Frm85_LM_SEARCH_5 = Frm101.L46_Text
    Frm85_LM_SEARCH_5_LOGIC = "="
    Frm85_LM_SEARCH_6 = "HQ"
    Frm85_LM_SEARCH_6_LOGIC = "="
    
End If
If Frm101.L47_Text = "Semua" Then
    Frm85_LM_SEARCH_20 = Null
    Frm85_LM_SEARCH_20_LOGIC = "<>"
Else
    If InStr(1, Frm101.L47_Text, "  |  ") <> 0 Then
        Frm84_LM_EMP_NO = Split(Frm101.L47_Text, "  |  ")(1)
        Frm84_LM_EMP_NAMA = Split(Frm101.L47_Text, "  |  ")(0)
    End If
    Frm85_LM_SEARCH_20 = Frm84_LM_EMP_NO
    Frm85_LM_SEARCH_20_LOGIC = "="
End If

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
        End If
    End If
End If

Frm85_LM_PAGE_FOUND = 0

'#### Header Report ###
If Frm101.L9_Text = 0 Then Frm85.L19_Text = "Report Belian Buyback / Trade In Bagi Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "]." 'Report Header"
If Frm101.L9_Text = 1 Then Frm85.L19_Text = "Report Belian Buyback / Trade In Bagi Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "] Dari " & TM & " Hingga " & TA 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select * from Data_Database where no_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "') order by tarikh_belian ASC , bill_No_Trade_In ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select * from Data_Database where no_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 6 & "'OR receiving_Status='" & 7 & "') order by tarikh_belian ASC , bill_No_Trade_In ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    'If rs!receiving_Status = "2" Or rs!receiving_Status = "3" Then
        x = x + 1
        If Frm85_LM_PAGE_FOUND = 0 Then
            If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
                If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                    Frm85.L79_Text = Frm85.L79_Text + 1
                    Frm85_LM_PAGE_FOUND = 1
                ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                    If IsNumeric(Frm85.L79_Text) Then
                        If Frm85.L79_Text <> 1 Then
                            Frm85.L79_Text = Frm85.L79_Text - 1
                            Frm85_LM_PAGE_FOUND = 1
                        End If
                    End If
                End If
            End If
        End If
        Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x
        
        With Frm85.LV3.ListItems.Add(, , rs!ID)
        
            .ListSubItems.Add , , Y
            
            If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
            
            If Not IsNull(rs!tarikh_belian) Then 'Tarikh Belian
                .ListSubItems.Add , , rs!tarikh_belian
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!no_siri_Produk) Then 'No. Siri Produk
                .ListSubItems.Add , , rs!no_siri_Produk
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!kod_Purity) Then 'Purity
                .ListSubItems.Add , , rs!kod_Purity
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!kategori_Produk) Then 'Kategori Produk
                .ListSubItems.Add , , rs!kategori_Produk
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!nama_Supplier) Then 'Nama Supplier
                .ListSubItems.Add , , rs!nama_Supplier
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!Berat) Then 'Berat (g)
                .ListSubItems.Add , , Format(rs!Berat, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!kos_Belian_Gram) Then 'Rate Penerimaan (RM/g)
                .ListSubItems.Add , , Format(rs!kos_Belian_Gram, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!UPAH) Then 'Upah (RM)
                .ListSubItems.Add , , Format(rs!UPAH, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!SpreadValue) Then 'Spread (%)
                .ListSubItems.Add , , Format(rs!SpreadValue, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!harga_lepas_spread) Then 'Harga Selepas Spread (RM)
                .ListSubItems.Add , , Format(rs!harga_lepas_spread, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If

            If Not IsNull(rs!adjustment) Then 'Adjustment (RM)
                .ListSubItems.Add , , Format(rs!adjustment, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!harga_item) Then 'Harga Belian (RM) : Tidak Campur Cukai GST
                .ListSubItems.Add , , Format(rs!harga_item, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!Upah_Jualan) Then 'Upah Jualan (RM) : Pelanggan
                .ListSubItems.Add , , Format(rs!Upah_Jualan, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If

            If Not IsNull(rs!Upah_Member) Then 'Upah Jualan (RM) : Ahli Biasa
                .ListSubItems.Add , , Format(rs!Upah_Member, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If

            If Not IsNull(rs!Upah_Pengedar) Then 'Upah Jualan (RM) : Silver
                .ListSubItems.Add , , Format(rs!Upah_Pengedar, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!Upah_RAF) Then 'Upah Jualan (RM) : Gold
                .ListSubItems.Add , , Format(rs!Upah_RAF, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!upah_normal_dealer) Then 'Upah Jualan (RM) : Platinum
                .ListSubItems.Add , , Format(rs!upah_normal_dealer, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!upah_master_dealer) Then 'Upah Jualan (RM) : Master Dealer
                .ListSubItems.Add , , Format(rs!upah_master_dealer, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!code_Supplier) Then 'Tetapan Harga Jualan (RM) : Pelanggan
                .ListSubItems.Add , , Format(rs!code_Supplier, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!HargaJualan_Member) Then 'Tetapan Harga Jualan (RM) : Pelanggan
                .ListSubItems.Add , , Format(rs!HargaJualan_Member, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!HargaJualan_Pengedar) Then 'Tetapan Harga Jualan (RM) : Pengedar
                .ListSubItems.Add , , Format(rs!HargaJualan_Pengedar, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!HargaJualan_RAF) Then 'Tetapan Harga Jualan (RM) : RAF
                .ListSubItems.Add , , Format(rs!HargaJualan_RAF, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!hargajualan_normal_dealer) Then 'Tetapan Harga Jualan (RM) : Normal Dealer
                .ListSubItems.Add , , Format(rs!hargajualan_normal_dealer, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If

            If Not IsNull(rs!hargajualan_master_dealer) Then 'Tetapan Harga Jualan (RM) : Master Dealer
                .ListSubItems.Add , , Format(rs!hargajualan_master_dealer, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
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
            
            If Not IsNull(rs!dimension_Saiz) Then 'Tebal
                .ListSubItems.Add , , rs!dimension_Saiz
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!bill_No_Belian) Then 'No invoice supplier
                .ListSubItems.Add , , rs!bill_No_Belian
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
            
            If Not IsNull(rs!cawangan) Then 'Cawangan
                .ListSubItems.Add , , rs!cawangan
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!bill_No_Trade_In) Then
                .ListSubItems.Add , , rs!bill_No_Trade_In
            Else
                .ListSubItems.Add , , ""
            End If
        End With

    'End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select COUNT(ID),SUM(Berat),SUM(harga_item) from Data_Database where no_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select COUNT(ID),SUM(Berat),SUM(harga_item) from Data_Database where no_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

If Not IsNull(rs(0)) Then Frm85.L51_Text = Format(rs(0), "#,##0")
If Not IsNull(rs(1)) Then Frm85.L52_Text = Format(rs(1), "#,##0.00 g")
If Not IsNull(rs(2)) Then Frm85.L53_Text = "RM " & Format(rs(2), "#,##0.00")

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    
    Frm85.Pic4.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    
    MsgBox "Tiada Rekod Belian Buyback / Trade In Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_Header_Report_Stok()
'on error resume next
With Frm85.LV4
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear
    
    Frm85.LV4.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "Tarikh Belian", 1500, 2
    .ColumnHeaders.Add 5, , "No. Siri Produk", 1700
    .ColumnHeaders.Add 6, , "Purity", 1500
    .ColumnHeaders.Add 7, , "Kategori Produk", 4000
    .ColumnHeaders.Add 8, , "Supplier", 4000
    .ColumnHeaders.Add 9, , "Berat (g)", 1400, 1
    .ColumnHeaders.Add 10, , "Rate Penerimaan (RM/g)", 2400, 1
    .ColumnHeaders.Add 11, , "Upah (RM)", 1500, 1
    .ColumnHeaders.Add 12, , "Spread (%)", 1200, 1
    .ColumnHeaders.Add 13, , "Harga Selepas Spread (RM)", 2800, 1
    .ColumnHeaders.Add 14, , "Adjustment (RM)", 1700, 1
    .ColumnHeaders.Add 15, , "Harga Belian Termasuk GST (RM)", 3300, 1
    .ColumnHeaders.Add 16, , "Upah Jualan (RM) : Pelanggan", 3200, 1
    .ColumnHeaders.Add 17, , "Upah Jualan (RM) : Ahli Biasa", 3200, 1
    .ColumnHeaders.Add 18, , "Upah Jualan (RM) : Silver", 3200, 1
    .ColumnHeaders.Add 19, , "Upah Jualan (RM) : Gold", 3200, 1
    .ColumnHeaders.Add 20, , "Upah Jualan (RM) : Platinum", 3200, 1
    .ColumnHeaders.Add 21, , "Upah Jualan (RM) : Master Dealer", 0, 1
    .ColumnHeaders.Add 22, , "Tetapan Harga Jualan (RM) : Pelanggan", 3700, 1
    .ColumnHeaders.Add 23, , "Tetapan Harga Jualan (RM) : Ahli", 3500, 1
    .ColumnHeaders.Add 24, , "Tetapan Harga Jualan (RM) : Silver", 3500, 1
    .ColumnHeaders.Add 25, , "Tetapan Harga Jualan (RM) : Gold", 3500, 1
    .ColumnHeaders.Add 26, , "Tetapan Harga Jualan (RM) : Platinum", 3500, 1
    .ColumnHeaders.Add 27, , "Tetapan Harga Jualan (RM) : Master Dealer", 0, 1
    .ColumnHeaders.Add 28, , "Dulang", 1700, 2
    .ColumnHeaders.Add 29, , "Panjang", 1700, 2
    .ColumnHeaders.Add 30, , "Lebar", 1700, 2
    .ColumnHeaders.Add 31, , "Saiz", 1700, 2
    .ColumnHeaders.Add 32, , "No. Invoice", 2000, 0
    .ColumnHeaders.Add 33, , "Code 1", 1700, 0
    .ColumnHeaders.Add 34, , "Code 2", 1700, 0
    .ColumnHeaders.Add 35, , "Cawangan", 2500, 0
    
End With
End Sub
Sub Frm85_report_stok_page()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date
Dim Frm85_LM_TOTAL_PAGE As Double

Frm85_LM_TOTAL_PAGE = 0
Frm85_PAGE_SIZE = 34
x = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

Frm85.L54_Text = Format(0, "#,##0")
Frm85.L55_Text = Format(0, "#,##0.00 g")
Frm85.L56_Text = "RM " & Format(0, "#,##0.00")

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If
If Frm101.L6_Text = "Semua Kategori Produk" Then
    Frm85_LM_SEARCH_2 = Null
    Frm85_LM_SEARCH_2_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_2 = Frm101.L6_Text
    Frm85_LM_SEARCH_2_LOGIC = "="
End If
If Frm101.L34_Text = "Semua Dulang" Then
    Frm85_LM_SEARCH_3 = Null
    Frm85_LM_SEARCH_3_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_3 = Frm101.L34_Text
    Frm85_LM_SEARCH_3_LOGIC = "="
End If
If Frm101.L37_Text = "Semua Supplier" Then
    Frm85_LM_SEARCH_4 = Null
    Frm85_LM_SEARCH_4_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_4 = Frm101.L37_Text
    Frm85_LM_SEARCH_4_LOGIC = "="
End If
If Frm101.L46_Text = "Semua cawangan" Then

    Frm85_LM_SEARCH_5 = Null
    Frm85_LM_SEARCH_5_LOGIC = "<>"
    Frm85_LM_SEARCH_6 = Null
    Frm85_LM_SEARCH_6_LOGIC = "<>"
    
Else

    Frm85_LM_SEARCH_5 = Frm101.L46_Text
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

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
        End If
    End If
End If

Frm85_LM_PAGE_FOUND = 0

'#### Header Report ###
If Frm101.L9_Text = 0 Then Frm85.L23_Text = "Report Stok Bagi Supplier [" & Frm101.L37_Text & "] , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "]." 'Report Header"
If Frm101.L9_Text = 1 Then Frm85.L23_Text = "Report Stok Bagi Supplier [" & Frm101.L37_Text & "] , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "] Dari " & TM & " Hingga " & TA 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select * from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "') AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND (receiving_Status='" & 0 & "' OR receiving_Status='" & 1 & "' OR receiving_Status='" & 4 & "') order by tarikh_belian ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select * from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "') AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 0 & "' OR receiving_Status='" & 1 & "' OR receiving_Status='" & 4 & "') order by tarikh_belian ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm85_LM_PAGE_FOUND = 0 Then
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm85.L79_Text = Frm85.L79_Text + 1
                Frm85_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm85.L79_Text) Then
                    If Frm85.L79_Text <> 1 Then
                        Frm85.L79_Text = Frm85.L79_Text - 1
                        Frm85_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x

    With Frm85.LV4.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!tarikh_belian) Then 'Tarikh Belian
            .ListSubItems.Add , , rs!tarikh_belian
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_siri_Produk) Then 'No. Siri Produk
            .ListSubItems.Add , , rs!no_siri_Produk
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kod_Purity) Then 'Purity
            .ListSubItems.Add , , rs!kod_Purity
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kategori_Produk) Then 'Kategori Produk
            .ListSubItems.Add , , rs!kategori_Produk
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!nama_Supplier) Then 'Nama Supplier
            .ListSubItems.Add , , rs!nama_Supplier
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Berat) Then 'Berat (g)
            .ListSubItems.Add , , Format(rs!Berat, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kos_Belian_Gram) Then 'Rate Penerimaan (RM/g)
            .ListSubItems.Add , , Format(rs!kos_Belian_Gram, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!UPAH) Then 'Upah (RM)
            .ListSubItems.Add , , Format(rs!UPAH, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!SpreadValue) Then 'Spread (%)
            .ListSubItems.Add , , Format(rs!SpreadValue, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_lepas_spread) Then 'Harga Selepas Spread (RM)
            .ListSubItems.Add , , Format(rs!harga_lepas_spread, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!adjustment) Then 'Adjustment (RM)
            .ListSubItems.Add , , Format(rs!adjustment, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_item) Then 'Harga Belian (RM) : Tidak Campur Cukai GST
            .ListSubItems.Add , , Format(rs!harga_item, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Upah_Jualan) Then 'Upah Jualan (RM) : Pelanggan
            .ListSubItems.Add , , Format(rs!Upah_Jualan, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!Upah_Member) Then 'Upah Jualan (RM) : Ahli Biasa
            .ListSubItems.Add , , Format(rs!Upah_Member, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!Upah_Pengedar) Then 'Upah Jualan (RM) : Silver
            .ListSubItems.Add , , Format(rs!Upah_Pengedar, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Upah_RAF) Then 'Upah Jualan (RM) : Gold
            .ListSubItems.Add , , Format(rs!Upah_RAF, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!upah_normal_dealer) Then 'Upah Jualan (RM) : Platinum
            .ListSubItems.Add , , Format(rs!upah_normal_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!upah_master_dealer) Then 'Upah Jualan (RM) : Master Dealer
            .ListSubItems.Add , , Format(rs!upah_master_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!code_Supplier) Then 'Tetapan Harga Jualan (RM) : Pelanggan
            .ListSubItems.Add , , Format(rs!code_Supplier, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!HargaJualan_Member) Then 'Tetapan Harga Jualan (RM) : Pelanggan
            .ListSubItems.Add , , Format(rs!HargaJualan_Member, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!HargaJualan_Pengedar) Then 'Tetapan Harga Jualan (RM) : Pengedar
            .ListSubItems.Add , , Format(rs!HargaJualan_Pengedar, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!HargaJualan_RAF) Then 'Tetapan Harga Jualan (RM) : RAF
            .ListSubItems.Add , , Format(rs!HargaJualan_RAF, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!hargajualan_normal_dealer) Then 'Tetapan Harga Jualan (RM) : Normal Dealer
            .ListSubItems.Add , , Format(rs!hargajualan_normal_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!hargajualan_master_dealer) Then 'Tetapan Harga Jualan (RM) : Master Dealer
            .ListSubItems.Add , , Format(rs!hargajualan_master_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
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
        
        If Not IsNull(rs!dimension_Saiz) Then 'Tebal
            .ListSubItems.Add , , rs!dimension_Saiz
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!bill_No_Belian) Then 'No. Invoice Supplier
            .ListSubItems.Add , , rs!bill_No_Belian
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
        
        If Not IsNull(rs!cawangan) Then 'Cawangan
            .ListSubItems.Add , , rs!cawangan
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
If Frm101.L9_Text = 0 Then rs.Open "select COUNT(ID),SUM(Berat),SUM(harga_item) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "') AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND (receiving_Status='" & 0 & "' OR receiving_Status='" & 1 & "' OR receiving_Status='" & 4 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select COUNT(ID),SUM(Berat),SUM(harga_item) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "') AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 0 & "' OR receiving_Status='" & 1 & "' OR receiving_Status='" & 4 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

If Not IsNull(rs(0)) Then Frm85.L54_Text = Format(rs(0), "#,##0")
If Not IsNull(rs(1)) Then Frm85.L55_Text = Format(rs(1), "#,##0.00 g")
If Not IsNull(rs(2)) Then Frm85.L56_Text = "RM " & Format(rs(2), "#,##0.00")

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    
    Frm85.Pic5.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    
    MsgBox "Tiada Data Stok Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_Recall_Data_Belian()
'on error resume next
Dim rs2 As ADODB.Recordset

DATA_FOUND = 0
Frm_LM_DATA_PENJUAL_BUYBACK = 0
Frm85_LM_No_PENJUAL = vbNullString

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from Data_Database where NoRujukanSistem='" & Frm83.L9_Text & "' order by ID ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Not IsNull(rs!receiving_Status) Then
        
        If rs!receiving_Status = 0 Or rs!receiving_Status = 1 Or rs!receiving_Status = 2 Or rs!receiving_Status = 3 Or rs!receiving_Status = 6 Or rs!receiving_Status = 7 Then
            Frm83.CB9 = 1
            Frm83.CB10 = 0
        End If
        If rs!receiving_Status = 4 Or rs!receiving_Status = 5 Or rs!receiving_Status = 8 Then
            Frm83.CB9 = 0
            Frm83.CB10 = 1
        End If
        
    End If
        
End If

rs.Close
Set rs = Nothing

'### Masukkan maklumat data barang ke dalam table #data_database ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "insert into " & G_BELIAN_TEMP & "(id_database,terminal,tarikh_belian,no_siri_produk,bill_no_belian,supplier_ID,nama_Supplier,Kod_Supplier,purity_ID,purity,kod_Purity,kategori_produk_ID,kategori_Produk,Kod_Kategori_Produk,Berat,beza_berat,UPAH,Upah30,riyal,kos_Belian_Gram,kos_Belian_Item,Spread,harga_lepas_spread,adjustment,kos_item_tanpa_tax,cara_Belian,dimension_Panjang,dimension_Lebar,dimension_Saiz,code1,code2,harga_Per_Gram_Item,dulang,no_cert,gst_barang_atau_upah,StatusItem,Upah_Jualan,Upah_Member,Upah_RAF,Upah_Pengedar,code_Supplier,HargaJualan_Member,HargaJualan_Pengedar,upah_normal_dealer,upah_master_dealer,HargaJualan_RAF,hargajualan_normal_dealer,hargajualan_master_dealer,remarks,gst_ari_nashi,kadar_gst,jumlah_gst,harga_item,jenis,harga_tanpa_gst,gst_included,jenis_trade_in,flag_upah,upah_per_gram,flag_image,no_id_gst)" & _
            "select ID,terminal,tarikh_belian,no_siri_produk,bill_no_belian,supplier_ID,nama_Supplier,Kod_Supplier,purity_ID,purity,kod_Purity,kategori_produk_ID,kategori_Produk,Kod_Kategori_Produk,Berat,beza_berat,UPAH,Upah30,riyal,kos_Belian_Gram,kos_Belian_Item,SpreadValue,harga_lepas_spread,adjustment,kos_item_tanpa_tax,cara_Belian,dimension_Panjang,dimension_Lebar,dimension_Saiz,code1,code2,harga_Per_Gram_Item,dulang,no_cert,gst_barang_atau_upah,StatusItem,Upah_Jualan,Upah_Member,Upah_RAF,Upah_Pengedar,code_Supplier,HargaJualan_Member,HargaJualan_Pengedar,upah_normal_dealer,upah_master_dealer,HargaJualan_RAF,hargajualan_normal_dealer,hargajualan_master_dealer,remarks,gst_ari_nashi,kadar_gst,jumlah_gst,harga_item,receiving_Status,harga_tanpa_gst,gst_included,jenis_trade_in,flag_upah,upah_per_gram,flag_image,no_id_gst from Data_Database WHERE norujukansistem='" & Frm83.L9_Text & "'"

Set rs = cn.Execute(strsql)
Set rs = Nothing
'### Masukkan maklumat data barang ke dalam table #data_database ### - End

GM_NEXT_PREV = 0

Call Frm83_Senarai_Belian_Header
Call Frm83_Senarai_Belian

'### Maklumat Belian / Buyback (Akaun) ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 16_gold_bar_belian where no_rujukan='" & Frm83.L9_Text & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!tarikh) Then Frm83.DTPicker1 = rs!tarikh 'Tarikh Belian
    If Not IsNull(rs!tunai) Then Frm83.L26_Text = rs!tunai 'Cara Bayaran : Tunai (Jumlah Keseluruhan Dengan Cukai GST)
    If Not IsNull(rs!jumlah_asal) Then Frm83.L11_Text = rs!jumlah_asal 'Jumlah Bayaran Asal (Jumlah Tanpa Cukai GST)
    If Not IsNull(rs!gst_value) Then Frm83.L8_Text = rs!gst_value ''Jumlah Cukai GST (%)'Jumlah Cukai GST (%)
    If Not IsNull(rs!gst_zr_harga) Then Frm83.L22_Text = rs!gst_zr_harga 'Jumlah Bayaran Yang Dikenakan Cukai GST ZR (RM)
    If Not IsNull(rs!gst_zr_cukai) Then Frm83.L23_Text = rs!gst_zr_cukai 'Jumlah Bayaran Cukai GST ZR (RM)
    If Not IsNull(rs!gst_sr_harga) Then Frm83.L24_Text = rs!gst_sr_harga 'Jumlah Bayaran Yang Dikenakan Cukai GST SR (RM)
    If Not IsNull(rs!gst_sr_cukai) Then Frm83.L25_Text = rs!gst_sr_cukai 'Jumlah Bayaran Cukai GST SR (RM)
    If Not IsNull(rs!tunai) Then Frm83.TB40 = Format(rs!tunai, "#,##0.00")
    If Not IsNull(rs!bank_in) Then Frm83.TB41 = Format(rs!bank_in, "#,##0.00")
    If Not IsNull(rs!no_id_gst_supplier) Then Frm83.TB28 = rs!no_id_gst_supplier 'No. ID GST Supplier
    'If Not IsNull(rs!no_resit_supplier) Then Frm83.TB15 = rs!no_resit_supplier 'No. Resit Dari Supplier (Jika Ada)
    If Not IsNull(rs!jumlah_dengan_gst) Then Frm83.L26_Text = rs!jumlah_dengan_gst 'Jumlah Bayaran Keseluruhan (Jumlah Dengan Cukai GST)
    If Not IsNull(rs!flag_trade_in) Then
        If rs!flag_trade_in = 0 Then
            Frm83.CB7 = 1 'Flag Trade In // 0 : Tiada , 1 : Ada
        ElseIf rs!flag_trade_in = 1 Then
            Frm_LM_DATA_PENJUAL_BUYBACK = 1
            Frm83.CB8 = 1 'Flag Trade In // 0 : Tiada , 1 : Ada
            'Frm83.Pic3.Visible = True
            If Not IsNull(rs!kategori_penjual) Then Frm83.L39_Text = rs!kategori_penjual
            If Not IsNull(rs!no_rujukan_pelanggan_buyback) Then Frm85_LM_No_PENJUAL = rs!no_rujukan_pelanggan_buyback 'No. Rujukan Penjual (Penjual Buyback)
            If Not IsNull(rs!no_resit_trade_in) Then Frm83.L12_Text = rs!no_resit_trade_in 'No. Resit Trade In
            MDI_frm1.L5_Text = 3
        End If
    End If
    If Not IsNull(rs!no_pekerja) Then
        Frm83_LM_No_PEKERJA = rs!no_pekerja 'No. Pekerja
    End If
End If

rs.Close
Set rs = Nothing
'### Maklumat Belian / Buyback (Akaun) ### - End

'### Carian Maklumat Penjual Bagi Buyback ### - Start
If Frm85_LM_No_PENJUAL = vbNullString Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 44_senarai_pelanggan where no_resit='" & Frm83.L12_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        If Not IsNull(rs!Nama) Then
            Frm26.TB1 = rs!Nama 'Nama
            Frm83.L36_Text = rs!Nama 'Nama
        End If
        If Not IsNull(rs!no_tel) Then Frm26.TB2 = rs!no_tel 'No. Telefon

    End If
    
    rs.Close
    Set rs = Nothing
    
ElseIf Frm85_LM_No_PENJUAL <> vbNullString Then '0 : Penjual Tidak Berdaftar , 1 : Penjual Berdaftar , 2 : Ahli

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm85_LM_No_PENJUAL & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        If Not IsNull(rs!Nama) Then
            Frm28.L1_Text = rs!Nama 'Nama
            Frm83.L37_Text = rs!Nama 'Nama
        End If
        If Not IsNull(rs!no_ic) Then Frm28.L2_Text = rs!no_ic 'No. Kad Pengenalan
        If Not IsNull(rs!no_tel) Then Frm28.L3_Text = rs!no_tel 'No. Telefon
        If Not IsNull(rs!Email) Then Frm28.L4_Text = rs!Email 'E-mail
        If Not IsNull(rs!no_pelanggan) Then Frm28.L5_Text = rs!no_pelanggan 'No. Pelanggan

    End If
    
    rs.Close
    Set rs = Nothing
    
End If
'### Carian Maklumat Penjual Bagi Buyback ### - End

Frm83.CBB6.Enabled = True
Frm83.CBB6.BackColor = &HFFFFFF

'### Carian Maklumat Penjual (Data Pekerja) ### - Start
DATA_PEKERJA_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where NoPekerja='" & Frm83_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm83_LM_MAKLUMAT_PEKERJA = rs!Samaran & "  |  " & rs!NoPekerja
    DATA_PEKERJA_FOUND = 1
End If

rs.Close
Set rs = Nothing

If DATA_PEKERJA_FOUND = 1 Then
    On Error GoTo Err_A:
    Frm83.CBB6 = Frm83_LM_MAKLUMAT_PEKERJA
Restore_A:
End If
'### Carian Maklumat Penjual (Data Pekerja) ### - End

'GM_NEXT_PREV = 0

'Call Frm83_Senarai_Belian_Header
'Call Frm83_Senarai_Belian

'Frm83.Pic1.Visible = True
        
Frm83.Show
Frm85.Hide

Frm83.L21_Text = 1
Frm83.CMD1.Visible = False
Frm83.CMD12.Visible = True
Frm83.CMD13.Visible = False
Frm83.CMD14.Visible = False

Frm83.CMD2.Visible = False
Frm83.CMD5.Visible = False
Frm83.CMD10.Visible = True
Frm83.CMD11.Visible = True

Exit Sub
Err_A:
Frm83.CBB6.AddItem Frm83_LM_MAKLUMAT_PEKERJA
Frm83.CBB6 = Frm83_LM_MAKLUMAT_PEKERJA
Resume Restore_A:
End Sub
Sub Frm85_Recall_Data_Jualan()
'on error resume next
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim Frm84_LM_SIMPANAN_ASAL As Double 'Jumlah Simpanan Asal Yang Ada (RM)
Dim Frm84_LM_SIMPANAN_DIGUNAKAN As Double 'Jumlah Simpanan Yang Digunakan (RM)
Dim Frm84_LM_MATA_ASAL As Double
Dim Frm84_LM_MATA_TEBUS As Double
Dim Frm84_LM_MATA_DAPAT As Double

DATA_FOUND = 0
Frm85_LM_DISKAUN = 0 '0 Tiada Diskaun , 1 : Ada Diskaun
Frm84_LM_No_PEKERJA = vbNullString
Frm84_LM_Flag_SIMPANAN = 0 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
Frm84_LM_SIMPANAN_ASAL = 0 'Jumlah Simpanan Asal Yang Ada (RM)
Frm84_LM_SIMPANAN_DIGUNAKAN = 0 'Jumlah Simpanan Yang Digunakan (RM)
Frm84_LM_KATEGORI_PEMBELI = 0
Frm84_LM_No_PEMBELI = vbNullString
Frm85_LM_No_AGEN = vbNullString
Frm84_LM_JUALAN_TRADE = 0 '0 : Jualan tanpa trade in , 1 : Jualan dengan trade in
Frm84_LM_MATA_ASAL = 0
Frm84_LM_MATA_TEBUS = 0
Frm84_LM_MATA_DAPAT = 0


Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "insert into " & G_JUALAN_TEMP & "(id_database,no_resit,status,baru_or_ti,no_siri_produk,flag_barang,nama_purity,kategori_produk,purity,berat_asal,berat_jualan,harga_semasa," _
            & "upah,harga_asal,diskaun,harga_lepas_diskaun,adjustment,harga_jualan,gst_ari_nashi,kadar_gst,jumlah_gst,harga_dengan_gst,harga_jual_excl_gst,harga_modal_gst,harga_modal_incl_gst,harga_modal_excl_gst,dropship,komisyen_per_gram," _
            & "jumlah_komisyen,status_jualan,type,potong_flag,modal_tanpa_gst,harga_per_gram_tanpa_gst,jualan_per_gram_dengan_gst,harga_per_gram_modal," _
            & "modal,untung,harga_per_gram_supplier,untung2,dulang,gst_include,harga_tanpa_gst,harga_koperasi,kadar_penurunan_upah,kadar_penurunan_bp," _
            & "harga_semasa_staff,harga_bp_asal,upah_asal,harga_staff,komisyen_staff,pemalar_tukaran_999,berat_999,upah_modal,gst_barang_atau_upah," _
            & "harga_jualan_dengan_gst,kadar_komisyen_upah,komisyen_upah,jualan_per_gram,modal_per_gram,flag_upah,upah_per_gram)" & _
            "select ID,'" & Frm84.L3_Text & "',1,baru_or_ti,no_siri_produk,flag_barang,nama_purity,kategori_produk,purity,berat_asal,berat_jualan,harga_semasa," _
            & "upah,harga_asal,diskaun,harga_lepas_diskaun,adjustment,harga_jualan,gst_ari_nashi,kadar_gst,jumlah_gst,harga_dengan_gst,harga_jual_excl_gst,harga_modal_gst,harga_modal_incl_gst,harga_modal_excl_gst,dropship,komisyen_per_gram," _
            & "jumlah_komisyen,1,type,potong_flag,modal_tanpa_gst,harga_per_gram_tanpa_gst,jualan_per_gram_dengan_gst,harga_per_gram_modal," _
            & "modal,untung,harga_per_gram_supplier,untung2,dulang,gst_include,harga_tanpa_gst,harga_koperasi,kadar_penurunan_upah,kadar_penurunan_bp," _
            & "harga_semasa_staff,harga_bp_asal,upah_asal,harga_staff,komisyen_staff,pemalar_tukaran_999,berat_999,upah_modal,gst_barang_atau_upah," _
            & "harga_jualan_dengan_gst,kadar_komisyen_upah,komisyen_upah,jualan_per_gram,modal_per_gram,flag_upah,upah_per_gram from 23_senarai_jualan WHERE status_rekod = 1 AND no_resit='" & Frm84.L3_Text & "'"
    
Set rs = cn.Execute(strsql)
Set rs = Nothing
                    
Frm84.L41_Text = 1
frm130.L41_Text = 1
DATA_FOUND = 1

GoTo aaaa:

'### Masukkan Data Jualan Ke Dalam Table Jualan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 23_senarai_jualan where no_resit='" & Frm84.L3_Text & "' AND status_rekod = 1", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    Set rs1 = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs1.Open "select * from " & G_JUALAN_TEMP & "", cn, adOpenKeyset, adLockOptimistic

    rs1.AddNew
    If Not IsNull(rs!ID) Then
        rs1!id_database = rs!ID 'No. ID Database Asal
    Else
        rs1!id_database = Null 'No. ID Database Asal
    End If
    rs1!no_resit = Frm84.L3_Text 'No. Resit
    If Not IsNull(rs!no_siri_Produk) Then
        rs1!no_siri_Produk = rs!no_siri_Produk 'No. Siri Produk
    Else
        rs1!no_siri_Produk = Null 'No. Siri Produk
    End If
    If Not IsNull(rs!kategori_Produk) Then
        rs1!kategori_Produk = rs!kategori_Produk 'Kategori Produk
    Else
        rs1!no_siri_Produk = Null 'Kategori Produk
    End If
    If Not IsNull(rs!flag_barang) Then
        rs1!flag_barang = rs!flag_barang
    Else
        rs1!flag_barang = Null 'Purity
    End If
    If Not IsNull(rs!nama_purity) Then
        rs1!nama_purity = rs!nama_purity
    Else
        rs1!nama_purity = Null
    End If
    If Not IsNull(rs!purity) Then
        rs1!purity = rs!purity 'Purity
    Else
        rs1!purity = Null 'Purity
    End If
    If Not IsNull(rs!Berat_Asal) Then
        rs1!Berat_Asal = rs!Berat_Asal 'Berat Asal (g)
    Else
        rs1!Berat_Asal = Null 'Berat Asal (g)
    End If
    If Not IsNull(rs!berat_jualan) Then
        rs1!berat_jualan = rs!berat_jualan 'Berat Jualan (g)
    Else
        rs1!berat_jualan = Null 'Berat Jualan (g)
    End If
    If Not IsNull(rs!harga_Semasa) Then
        rs1!harga_Semasa = Format(rs!harga_Semasa, "0.00") 'Harga Semasa (RM/g)
    Else
        rs1!harga_Semasa = Null 'Harga Semasa (RM/g)
    End If
    If Not IsNull(rs!UPAH) Then
        rs1!UPAH = Format(rs!UPAH, "0.00") 'Upah (RM)
    Else
        rs1!UPAH = Null 'Upah (RM)
    End If
    If Not IsNull(rs!harga_asal) Then
        rs1!harga_asal = Format(rs!harga_asal, "0.00") 'Harga Asal Item (RM)
    Else
        rs1!harga_asal = Null 'Harga Asal Item (RM)
    End If
    If Not IsNull(rs!diskaun) Then
        rs1!diskaun = Format(rs!diskaun, "0.00") 'Diskaun (%)
        If rs1!diskaun <> "0.00" Then
            Frm85_LM_DISKAUN = 1 '0 Tiada Diskaun , 1 : Ada Diskaun
        End If
    Else
        rs1!diskaun = Null 'Diskaun (%)
    End If
    If Not IsNull(rs!harga_lepas_diskaun) Then
        rs1!harga_lepas_diskaun = Format(rs!harga_lepas_diskaun, "0.00") 'Harga Selepas Diskaun (RM)
    Else
        rs1!harga_lepas_diskaun = Null 'Harga Selepas Diskaun (RM)
    End If
    If Not IsNull(rs!adjustment) Then
        rs1!adjustment = Format(rs!adjustment, "0.00") 'Harga Selepas Diskaun (RM)
    Else
        rs1!adjustment = Null 'Harga Selepas Diskaun (RM)
    End If
    If Not IsNull(rs!harga_jualan) Then
        rs1!harga_jualan = Format(rs!harga_jualan, "0.00") 'Harga Jualan (RM)
    Else
        rs1!harga_jualan = Null 'Harga Jualan (RM)
    End If
    If Not IsNull(rs!gst_ari_nashi) Then
        rs1!gst_ari_nashi = rs!gst_ari_nashi '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
    Else
        rs1!gst_ari_nashi = Null '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
    End If
    If Not IsNull(rs!kadar_gst) Then
        rs1!kadar_gst = rs!kadar_gst 'Kadar Cukai GST (%)
    Else
        rs1!kadar_gst = Null 'Kadar Cukai GST (%)
    End If
    If Not IsNull(rs!jumlah_gst) Then
        rs1!jumlah_gst = rs!jumlah_gst 'Jumlah Cukai GST (RM)
    Else
        rs1!jumlah_gst = Null 'Jumlah Cukai GST (RM)
    End If
    If Not IsNull(rs!harga_dengan_gst) Then
        rs1!harga_dengan_gst = Format(rs!harga_dengan_gst, "0.00") 'Harga Jualan Termasuk GST (RM)
    Else
        rs1!harga_dengan_gst = Null 'Harga Jualan Termasuk GST (RM)
    End If
    If Not IsNull(rs!dropship) Then
        rs1!dropship = rs!dropship '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
        If rs!dropship = 1 Then
            Frm84.CB7 = 1
        Else
            Frm84.CB7 = 0
        End If
    Else
        rs1!dropship = Null '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
    End If
    If Not IsNull(rs!komisyen_per_gram) Then
        rs1!komisyen_per_gram = Format(rs!komisyen_per_gram, "0.00") 'Komisyen Per Gram Dropship (RM/g) : BK
    Else
        rs1!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g) : BK
    End If
    
    If Not IsNull(rs!kadar_komisyen_upah) Then 'Kadar komisyen bagi upah kepada agen dropship
        rs1!kadar_komisyen_upah = rs!kadar_komisyen_upah
    Else
        rs1!kadar_komisyen_upah = Null
    End If
    If Not IsNull(rs!komisyen_upah) Then 'Jumlah komisyen bagi upah kepada agen dropship
        rs1!komisyen_upah = Format(rs!komisyen_upah, "0.00")
    Else
        rs1!komisyen_upah = Null
    End If
            
    If Not IsNull(rs!jumlah_komisyen) Then
        rs1!jumlah_komisyen = Format(rs!jumlah_komisyen, "0.00") 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
    Else
        rs1!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
    End If
    If Not IsNull(rs!harga_per_gram_modal) Then
        rs1!harga_per_gram_modal = Format(rs!harga_per_gram_modal, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
    Else
        rs1!harga_per_gram_modal = Null 'Harga Per Gram Bagi Modal (RM/g)
    End If
    If Not IsNull(rs!modal) Then
        rs1!modal = Format(rs!modal, "0.00") 'Harga Modal (RM)
    Else
        rs1!modal = Null 'Harga Modal (RM)
    End If
    If Not IsNull(rs!untung) Then
        rs1!untung = Format(rs!untung, "0.00") 'Jumlah Keuntungan
    Else
        rs1!untung = Null 'Jumlah Keuntungan
    End If
    If Not IsNull(rs!harga_per_gram_supplier) Then
        rs1!harga_per_gram_supplier = Format(rs!harga_per_gram_supplier, "0.00") 'Harga per gram (harga semasa) dari supplier (modal)
    Else
        rs1!harga_per_gram_supplier = Null 'Harga per gram (harga semasa) dari supplier (modal)
    End If
    If Not IsNull(rs!upah_modal) Then
        rs1!upah_modal = Format(rs!upah_modal, "0.00") 'Upah modal
    Else
        rs1!upah_modal = Null 'Upah modal
    End If
    If Not IsNull(rs!untung2) Then
        rs1!untung2 = Format(rs!untung2, "0.00") 'Untung jika restok pada harga supplier ini
    Else
        rs1!untung2 = Null 'Untung jika restok pada harga supplier ini
    End If
    If Not IsNull(rs!dulang) Then
        rs1!dulang = rs!dulang 'Dulang
    Else
        rs1!dulang = Null 'Dulang
    End If
    If Not IsNull(rs!potong_flag) Then
        rs1!potong_flag = rs!potong_flag '0 : Tiada Potong , 1 : Ada Potong
    Else
        rs1!potong_flag = Null '0 : Tiada Potong , 1 : Ada Potong
    End If
    If Not IsNull(rs!Type) Then
        rs1!Type = rs!Type '0 : BK , 1 : Barang Permata
    Else
        rs1!Type = Null '0 : BK , 1 : Barang Permata
    End If
    'If Not IsNull(rs!gst_include) Then 'Pilihan Cukai GST (SR) Samada Pelanggan Bayar Atau Kedai Bayar
    '    If rs!gst_include = "**Harga Termasuk GST" Then
    '        rs1!gst_include = 1
    '    End If
    'Else
    '    rs1!gst_include = 0
    'End If
    
    If Not IsNull(rs!gst_include) Then 'Pilihan Cukai GST (SR) Samada Pelanggan Bayar Atau Kedai Bayar
        If rs!gst_include = 1 Then
            rs1!gst_include = 1
        End If
    'Else
    '    rs1!gst_include = 0
    End If
    
    If Not IsNull(rs!harga_tanpa_gst) Then
        rs1!harga_tanpa_gst = Format(rs!harga_tanpa_gst, "#,##0.00") 'Harga Semasa (RM/g)
    Else
        rs1!harga_tanpa_gst = Null 'Harga Semasa (RM/g)
    End If
    rs1!Status = 1
    
'### Maklumat tetapan harga jualan kepada staff ### - Start
    If Not IsNull(rs!kadar_penurunan_upah) Then 'Kadar peratusan penurunan harga upah kepada staff (%)
        rs1!kadar_penurunan_upah = Format(rs!kadar_penurunan_upah, "0.00")
    Else
        rs1!kadar_penurunan_upah = Null
    End If
    If Not IsNull(rs!harga_semasa_staff) Then 'Harga emas semasa yang dijual kepada staff
        rs1!harga_semasa_staff = Format(rs!harga_semasa_staff, "0.00")
    Else
        rs1!harga_semasa_staff = Null
    End If
    If Not IsNull(rs!kadar_penurunan_bp) Then 'Kadar peratusan penurunan harga barang permata kepada staff (%)
        rs1!kadar_penurunan_bp = Format(rs!kadar_penurunan_bp, "0.00")
    Else
        rs1!kadar_penurunan_bp = Null
    End If
    If Not IsNull(rs!harga_staff) Then 'Harga yang dijual kepada staff (RM)
        rs1!harga_staff = Format(rs!harga_staff, "0.00")
    Else
        rs1!harga_staff = Null
    End If
    If Not IsNull(rs!harga_bp_asal) Then 'Tetapan harga barang permata yang asal (RM)
        rs1!harga_bp_asal = Format(rs!harga_bp_asal, "0.00")
    Else
        rs1!harga_bp_asal = Null
    End If
    If Not IsNull(rs!upah_asal) Then 'Tetapan upah asal (RM)
        rs1!upah_asal = Format(rs!upah_asal, "0.00")
    Else
        rs1!upah_asal = Null
    End If
    If Not IsNull(rs!komisyen_staff) Then
        rs1!komisyen_staff = Format(rs!komisyen_staff, "0.00")
    Else
        rs1!komisyen_staff = Null
    End If
'### Maklumat tetapan harga jualan kepada staff ### - End

    If Not IsNull(rs!gst_barang_atau_upah) Then '0 : GST pada harga jualan , 1 : GST pada upah
        rs1!gst_barang_atau_upah = rs!gst_barang_atau_upah
    Else
        rs1!gst_barang_atau_upah = 0
    End If
    If Not IsNull(rs!harga_jualan_dengan_gst) Then
        rs1!harga_jualan_dengan_gst = rs!harga_jualan_dengan_gst
    Else
        rs1!harga_jualan_dengan_gst = 0
    End If
    If Not IsNull(rs!harga_per_gram_tanpa_gst) Then
        rs1!harga_per_gram_tanpa_gst = rs!harga_per_gram_tanpa_gst
    Else
        rs1!harga_per_gram_tanpa_gst = 0
    End If
    
    If Not IsNull(rs!jualan_per_gram) Then
        rs1!jualan_per_gram = rs!jualan_per_gram
    Else
        rs1!jualan_per_gram = 0
    End If
    If Not IsNull(rs!modal_per_gram) Then
        rs1!modal_per_gram = rs!modal_per_gram
    Else
        rs1!modal_per_gram = 0
    End If
    If Not IsNull(rs!flag_upah) Then
        rs1!flag_upah = rs!flag_upah
    Else
        rs1!flag_upah = 1
    End If
    If Not IsNull(rs!upah_per_gram) Then
        rs1!upah_per_gram = rs!upah_per_gram
    Else
        rs1!upah_per_gram = Null
    End If

    DATA_FOUND = 1
    rs1.Update
    
    rs1.Close
    Set rs1 = Nothing

    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

aaaa:


'### Enable / Disable ruangan diskaun ### - Start
If Frm85_LM_DISKAUN = 1 Then '0 Tiada Diskaun , 1 : Ada Diskaun
    Frm84.TB7.Locked = False
    Frm84.TB7.BackColor = &HFFFFFF
Else
    Frm84.TB7.Locked = True
    Frm84.TB7.BackColor = &H8000000A
End If
'### Enable / Disable ruangan diskaun ### - Start

Call Frm84_Senarai_Jualan_Header
Call Frm84_Senarai_Jualan

'###Masukkan Jualan Ke Dalam Table Akaun Jualan### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where no_resit='" & Frm84.L3_Text & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    'GLOBAL_DISABLE = 1
    If Not IsNull(rs!kadar_gst) Then
        frm130.L8_Text = rs!kadar_gst 'Kadar Cukai GST (%)
        Frm84.L8_Text = rs!kadar_gst 'Kadar Cukai GST (%)
    Else
        frm130.L8_Text = G_RATE_GST 'Kadar Cukai GST (%)
    End If
    If Not IsNull(rs!tarikh) Then 'Tarikh Jualan
        Frm84.DTPicker1 = rs!tarikh
    Else
        'Frm84.DTPicker1 = vbNullString
    End If
    If Not IsNull(rs!epp) Then '0 : Bayaran selain dari EPP , 1 : Bayaran secara EPP
        If rs!epp = 1 Then
            Frm84.CB19 = 1
        Else
            Frm84.CB19 = 0
        End If
    Else
        Frm84.CB19 = 0
    End If
    If Not IsNull(rs!approval_code_epp) Then 'Approval Code (EPP)
        Frm84.TB41 = rs!approval_code_epp
    Else
        Frm84.TB41 = vbNullString
    End If
    If Not IsNull(rs!harga_barang) Then 'Jumlah Harga Barang Tanpa GST (RM)
        Frm84.L17_Text = rs!harga_barang
    Else
        Frm84.L17_Text = "0.00"
    End If
    If Not IsNull(rs!jumlah_cukai_gst) Then 'Jumlah Cukai GST (ZR + SR)
        Frm84.L18_Text = rs!jumlah_cukai_gst
    Else
        Frm84.L18_Text = "0.00"
    End If
    If Not IsNull(rs!harga_barang_dengan_gst) Then 'Jumlah Harga Barang Dengan GST (RM)
        Frm84.L19_Text = rs!harga_barang_dengan_gst
    Else
        Frm84.L19_Text = "0.00"
    End If
    If Not IsNull(rs!diskaun) Then 'Jumlah Diskaun (%)
        Frm84.TB19 = rs!diskaun
    Else
        Frm84.TB19 = "0.00"
    End If
    If Not IsNull(rs!harga_lepas_diskaun) Then 'Harga Selepas Diskaun (RM)
        Frm84.L20_Text = rs!harga_lepas_diskaun
    Else
        Frm84.L20_Text = "0.00"
    End If
    If Not IsNull(rs!adjustment) Then 'Adjustment (RM)
        Frm84.TB20 = rs!adjustment
    Else
        Frm84.TB20 = "0.00"
    End If
    If Not IsNull(rs!harga_jualan) Then 'Jumlah Harga Jualan (RM)
        Frm84.L21_Text = rs!harga_jualan
    Else
        Frm84.L21_Text = "0.00"
    End If
    If Not IsNull(rs!loss_trade_in) Then 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (%)
        Frm84.L38_Text = rs!loss_trade_in
    Else
        Frm84.L38_Text = "0.00"
    End If
    If Not IsNull(rs!loss_trade_in_rm) Then 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (RM)
        Frm84.L37_Text = rs!loss_trade_in_rm
    Else
        Frm84.L37_Text = "0.00"
    End If
    If Not IsNull(rs!flag_bayaran) Then '0 : Pembeli Bayar , 1 : Kedai Bayar
        If rs!flag_bayaran = 0 Then
            Frm84.L24_Text = "Jumlah Bayaran"
        ElseIf rs!flag_bayaran = 1 Then
            Frm84.L24_Text = "Harga Kedai Perlu Bayar Pelanggan"
        End If
    End If
    If Not IsNull(rs!jumlah_perlu_bayar) Then 'Jumlah Bayaran Yang Perlu Dibuat (RM)
        Frm84.L23_Text = rs!jumlah_perlu_bayar
    Else
        Frm84.L23_Text = "0.00"
    End If
    If Not IsNull(rs!kuantiti_barang) Then 'Kuantiti Barang Yang Dijual
        Frm84.L14_Text = rs!kuantiti_barang
    Else
        Frm84.L14_Text = "0.00"
    End If
    If Not IsNull(rs!JUMLAH_BERAT) Then 'Jumlah Berat Barang Yang Dijual
        Frm84.L15_Text = rs!JUMLAH_BERAT
    Else
        Frm84.L15_Text = "0.00"
    End If
    If Not IsNull(rs!gst_zr_harga) Then 'Harga Keseluruhan Bagi Barang ZR
        Frm84.L7_Text = rs!gst_zr_harga
    Else
        Frm84.L7_Text = "0.00"
    End If
    If Not IsNull(rs!gst_zr_cukai) Then 'Jumlah Cukai Bagi ZR
        Frm84.L9_Text = rs!gst_zr_cukai
    Else
        Frm84.L9_Text = "0.00"
    End If
    If Not IsNull(rs!gst_sr_harga) Then 'Harga Keseluruhan Bagi Barang SR
        Frm84.L10_Text = rs!gst_sr_harga
    Else
        Frm84.L10_Text = "0.00"
    End If
    If Not IsNull(rs!gst_sr_cukai) Then 'Jumlah Cukai Bagi SR
        Frm84.L11_Text = rs!gst_sr_cukai
    Else
        Frm84.L11_Text = "0.00"
    End If
    If Not IsNull(rs!no_rujukan_agen_dropship) Then 'No. Rujukan Agen Dropship
        Frm85_LM_No_AGEN = rs!no_rujukan_agen_dropship
    End If
    If Not IsNull(rs!flag_trade_in) Then
        
        If rs!jenis_trade_in = 3 Then
            G_TI_MODE = 3
        ElseIf rs!flag_trade_in = 1 Then '0 : Tiada Urusan Trade in , 1 : Ada Urusan Trade In
            
            If rs!jenis_trade_in = 3 Then
                G_TI_MODE = 3
            Else
                If Not IsNull(rs!no_resit_trade_in) Then 'No. Resit Trade In
                    Frm84.L57_Text = rs!no_resit_trade_in
                    Frm84.L60_Text = rs!no_resit_trade_in
                Else
                    Frm84.L57_Text = vbNullString
                    Frm84.L60_Text = vbNullString
                End If
                
                If Not IsNull(rs!jumlah_trade_in) Then 'Jumlah trade in (RM)
                    Frm84.L58_Text = rs!jumlah_trade_in
                Else
                    Frm84.L58_Text = vbNullString
                End If
            
                If Not IsNull(rs!jenis_trade_in) Then
                    
                    Frm84.L56_Text = rs!jenis_trade_in '0 : Tiada , 1 : Belian dengan trade in (sudah ada voucher) , 2 : Trade in
                    
                    If rs!jenis_trade_in = 1 Then
                    
                        If Not IsNull(rs!no_resit_trade_in) Then 'No. Resit Trade In
                            Frm84.L16_Text = rs!no_resit_trade_in
                        Else
                            Frm84.L16_Text = vbNullString
                        End If
                        
                        If Not IsNull(rs!jumlah_trade_in) Then 'Jumlah trade in (RM)
                            Frm84.TB17 = rs!jumlah_trade_in
                        Else
                            Frm84.TB17 = vbNullString
                        End If
                        
                        'Frm84.Pic5.Visible = True
                
                    ElseIf rs!jenis_trade_in = 2 Then
                    
                        Frm84_LM_JUALAN_TRADE = 1 '0 : Jualan tanpa trade in , 1 : Jualan dengan trade in
        
                    End If
                    
                Else
                
                    Frm84.L56_Text = 0 '0 : Tiada , 1 : Belian dengan trade in (sudah ada voucher) , 2 : Trade in
                    
                End If
            End If
        Else
            
            Frm84.L57_Text = vbNullString 'No. Resit Trade In
            Frm84.L58_Text = vbNullString 'Jumlah trade in (RM)
            
            Frm84.TB17 = "0.00"
            Frm84.L16_Text = vbNullString
            
        End If
        
    End If
    If Not IsNull(rs!invoice_type) Then '0 : Unlimited , Selain 0 (Limited : Mengikut nombor yang dimasukkan)
        Frm84.L46_Text = rs!invoice_type
    Else
        Frm84.L46_Text = 0
    End If
    If Not IsNull(rs!no_pekerja) Then Frm84_LM_No_PEKERJA = rs!no_pekerja
    
    If Not IsNull(rs!kategori_pembeli) Then
        If rs!kategori_pembeli = 1 Then
            Frm84.CB4 = 1
        ElseIf rs!kategori_pembeli = 2 Then
            Frm84.CB5 = 1
        ElseIf rs!kategori_pembeli = 3 Then
            Frm84.CB6 = 1
        ElseIf rs!kategori_pembeli = 4 Then
            Frm84.CB9 = 1
        ElseIf rs!kategori_pembeli = 5 Then
            Frm84.CB10 = 1
        'ElseIf rs!kategori_pembeli = 6 Then
        '    Frm84.CB11 = 1
        End If
    End If
    
    If Not IsNull(rs!caj_pos) Then 'Jumlah caj pos laju (postage)
        Frm84.TB42 = Format(rs!caj_pos, "0.00")
    Else
        Frm84.TB42 = "0.00"
    End If
    If Not IsNull(rs!no_tracking) Then 'No. Tracking pos laju
        Frm84.TB45 = rs!no_tracking
    Else
        Frm84.TB45 = vbNullString
    End If
    If Not IsNull(rs!jualan_online) Then
        If rs!jualan_online = 0 Then
            Frm84.CB27 = 0
        ElseIf rs!jualan_online = 1 Then
            Frm84.CB27 = 1
        End If
    Else
        Frm84.CB27 = 0
    End If
    
    If Not IsNull(rs!no_rujukan_pembeli) Then Frm84_LM_No_PEMBELI = rs!no_rujukan_pembeli 'No. Rujukan Pembeli
    
    If Not IsNull(rs!point_ari_nashi) Then
        If rs!point_ari_nashi = 1 Then
            Frm84_LM_POINT_ARI_NASHI = 1 '0 : Tiada mata , 1 : Ada mata
            Frm84.L79_Text = 1
        End If
    End If
    If Not IsNull(rs!kadar_peroleh_point) Then Frm84.TB35 = rs!kadar_peroleh_point
    If Not IsNull(rs!kadar_tebus_point) Then Frm84.TB37 = rs!kadar_tebus_point
    If Not IsNull(rs!kadar_diskaun) Then
        If rs!kadar_diskaun <> "0.00" Then
            Frm84.L80_Text = "RM " & Format(rs!kadar_diskaun, "0.00") & " /g"
        End If
    End If
    If Not IsNull(rs!kupon_diskaun) Then
        If rs!kupon_diskaun <> "0.00" Then
            Frm84.CB14 = 1
        Else
            Frm84.CB14 = 0
        End If
    Else
        Frm84.CB14 = 0
    End If
    If Not IsNull(rs!tunai) Then 'Cara Bayaran : Tunai
        frm130.TB27 = rs!tunai
    Else
        frm130.TB27 = "0.00"
    End If
    If Not IsNull(rs!remarks) Then
        Frm84.TB46 = rs!remarks
    Else
        Frm84.TB46 = vbNullString
    End If
    If Not IsNull(rs!kad_kredit) Then 'Cara Bayaran : Kad Kredit
        frm130.TB29 = rs!kad_kredit
    Else
        frm130.TB29 = "0.00"
    End If

    On Error GoTo Err_B:
    If Not IsNull(rs!jenis_kad) Then
        Frm84_LM_JENIS_KAD = rs!jenis_kad
        frm130.CBB2 = Frm84_LM_JENIS_KAD
        
Restore_B:
    End If
    'on error resume next

    If Not IsNull(rs!cas_Kad_Kredit) Then 'Cara Bayaran : Cas Kad Kredit (%)
        frm130.L31_Text = rs!cas_Kad_Kredit
    Else
        frm130.L31_Text = "0.00"
    End If
    If Not IsNull(rs!jumlah_cas_kad_kredit) Then 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
        frm130.L32_Text = rs!jumlah_cas_kad_kredit
    Else
        frm130.L32_Text = "0.00"
    End If
    If Not IsNull(rs!gst_kad_kredit) Then 'Cara Bayaran : Jumlah GST kad kredit (RM)
        frm130.L81_Text = rs!gst_kad_kredit
    Else
        frm130.L81_Text = "0.00"
    End If
    If Not IsNull(rs!jumlah_potongan_kad_kredit) Then 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
        frm130.L82_Text = rs!jumlah_potongan_kad_kredit
    Else
        frm130.L82_Text = "0.00"
    End If
    
    If Not IsNull(rs!duit_simpanan_kedai) Then 'Cara Bayaran : Simpanan Duit Di Kedai
        frm130.TB21 = rs!duit_simpanan_kedai
        If rs!duit_simpanan_kedai <> "0.00" Then
            Frm84_LM_Flag_SIMPANAN = 1 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
            Frm84_LM_SIMPANAN_DIGUNAKAN = rs!duit_simpanan_kedai 'Jumlah Simpanan Yang Digunakan (RM)
        End If
    Else
        frm130.TB21 = "0.00"
    End If
    If Not IsNull(rs!jumlah_bayaran) Then 'Cara Bayaran : Jumlah Bayaran
        frm130.TB32 = Format(rs!jumlah_bayaran, "#,##0.00")
        frm130.TB33 = Format(rs!jumlah_bayaran, "#,##0.00")
    Else
        frm130.TB32 = "0.00"
        frm130.TB33 = "0.00"
    End If
    If Not IsNull(rs!tunai) Then 'Cara Bayaran : Tunai
        frm130.TB27 = rs!tunai
    Else
        frm130.TB27 = "0.00"
    End If
    If Not IsNull(rs!bank_in) Then 'Cara Bayaran : Bank In
        frm130.TB28 = rs!bank_in
    Else
        frm130.TB28 = "0.00"
    End If
    
    If Not IsNull(rs!jenis_trade_in) Then
        G_TI_MODE = rs!jenis_trade_in
        If rs!jenis_trade_in = 3 Then
            If Not IsNull(rs!berat_trade_in) Then
                Frm84.TB49 = Format(rs!berat_trade_in, "#,##0.00")
                G_TI_BERAT = rs!berat_trade_in
            End If
            If Not IsNull(rs!harga_semasa_trade_in) Then
                Frm84.TB50 = Format(rs!harga_semasa_trade_in, "#,##0.00")
                G_TI_TRADE_IN = Format(rs!harga_semasa_trade_in, "#,##0.00")
            End If
            If Not IsNull(rs!harga_semasa_buyback) Then
                Frm84.TB51 = Format(rs!harga_semasa_buyback, "#,##0.00")
                G_TI_BUYBACK = Format(rs!harga_semasa_buyback, "#,##0.00")
            End If
            If Not IsNull(rs!caj_pertukaran) Then
                Frm84.TB52 = Format(rs!caj_pertukaran, "#,##0.00")
                G_TI_CAJ = Format(rs!caj_pertukaran, "#,##0.00")
            End If
        End If
    End If
    'GLOBAL_DISABLE = 0
End If

rs.Close
Set rs = Nothing
'###Masukkan Jualan Ke Dalam Table Akaun Jualan### - End

'### Maklumat Pembeli Yang Tidak Berdaftar Dengan Kedai ### - Start (08-07-2015)
If Frm84_LM_No_PEMBELI = vbNullString Then
    
    'Frm84.CMD11.Visible = True
    Call Frm26_initial
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 44_senarai_pelanggan where no_resit='" & Frm84.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!Nama) Then 'Nama
            Frm26.TB1 = rs!Nama
            'Frm84.L27_Text = rs!Nama
        Else
            Frm26.TB1 = vbNullString
        End If
        If Not IsNull(rs!no_tel) Then 'No. Telefon
            Frm26.TB2 = rs!no_tel
        Else
            Frm26.TB2 = vbNullString
        End If
    End If
    
    rs.Close
    Set rs = Nothing
End If
'### Maklumat Pembeli Yang Tidak Berdaftar Dengan Kedai ### - End

'###Update Data Pelanggan & Simpanan Duit Pelanggan### - Start
If Frm84_LM_No_PEMBELI <> vbNullString Then
    
    Call Frm28_initial
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm84_LM_No_PEMBELI & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!Nama) Then
            Frm28.L1_Text = rs!Nama 'Nama
            Frm84.L28_Text = rs!Nama
        End If
        If Not IsNull(rs!no_ic) Then Frm28.L2_Text = rs!no_ic 'No. Kad Pengenalan
        If Not IsNull(rs!no_tel) Then Frm28.L3_Text = rs!no_tel 'No. Telefon
        If Not IsNull(rs!Email) Then Frm28.L4_Text = rs!Email 'E-mail
        If Not IsNull(rs!no_pelanggan) Then Frm28.L5_Text = rs!no_pelanggan 'No. Pelanggan
        If Not IsNull(rs!baki_simpanan) Then
            frm130.L26_Text = Format(rs!baki_simpanan, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
            If IsNumeric(rs!baki_simpanan) Then
                Frm84_LM_SIMPANAN_ASAL = rs!baki_simpanan 'Jumlah Simpanan Asal Yang Ada (RM)
                
                frm130.L26_Text = Format(Frm84_LM_SIMPANAN_ASAL + Frm84_LM_SIMPANAN_DIGUNAKAN, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
            End If
        End If
        If Not IsNull(rs!baki_point) Then Frm84_LM_MATA_ASAL = rs!baki_point
    End If
    
    rs.Close
    Set rs = Nothing
    
    If Frm84_LM_POINT_ARI_NASHI = 1 Then '0 : Tiada mata , 1 : Ada mata
'### Maklumat agihan point ### - Start

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 71_tebus_agih_point where no_invoice='" & Frm84.L3_Text & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!harga_layak_bonus) Then Frm84.L75_Text = rs!harga_layak_bonus 'Harga yang membolehkan untuk mendaparkan point
            If Not IsNull(rs!kadar_peroleh_point) Then Frm84.TB35 = rs!kadar_peroleh_point 'Kadar perolehan point (eg. 0.5)
            If Not IsNull(rs!jumlah_peroleh_point) Then
                Frm84.L76_Text = rs!jumlah_peroleh_point 'Jumlah perolehan mata
                If IsNumeric(rs!jumlah_peroleh_point) Then Frm84_LM_MATA_DAPAT = rs!jumlah_peroleh_point
            End If
            If Not IsNull(rs!jumlah_tebus_point) Then
                Frm84.TB36 = rs!jumlah_tebus_point 'Jumlah mata yang ditebus
                If IsNumeric(rs!jumlah_tebus_point) Then Frm84_LM_MATA_TEBUS = rs!jumlah_tebus_point
            End If
            If Not IsNull(rs!kadar_tebus_point) Then Frm84.TB37 = rs!kadar_tebus_point 'Kadar tebusan mata
            If Not IsNull(rs!nilaian_tebus_point) Then Frm84.L78_Text = rs!nilaian_tebus_point 'Jumlah nilaian mata yang ditebus

        End If
        
        rs.Close
        Set rs = Nothing
        
        Frm84.L77_Text = Frm84_LM_MATA_TEBUS + Frm84_LM_MATA_ASAL - Frm84_LM_MATA_DAPAT

'### Maklumat agihan point ### - End
    End If
    
End If
'###Update Data Pelanggan & Simpanan Duit Pelanggan### - End

'### Update data agen drophip ### - Start
If Frm85_LM_No_AGEN <> vbNullString Then
    
    Call Frm27_initial
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm85_LM_No_AGEN & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!Nama) Then
            Frm27.L1_Text = rs!Nama 'Nama
            Frm84.L29_Text = rs!Nama
        End If
        If Not IsNull(rs!no_ic) Then Frm27.L2_Text = rs!no_ic 'No. Kad Pengenalan
        If Not IsNull(rs!no_tel) Then Frm27.L3_Text = rs!no_tel 'No. Telefon
        If Not IsNull(rs!Email) Then Frm27.L4_Text = rs!Email 'E-mail
        If Not IsNull(rs!no_pelanggan) Then Frm27.L5_Text = rs!no_pelanggan 'No. Pelanggan
    End If
    
    rs.Close
    Set rs = Nothing
    
    Frm84.CB7 = 1
    
End If
'### Update data agen drophip ### - End

If Frm84_LM_No_PEKERJA <> vbNullString Then
    '### Carian Maklumat Penjual (Data Pekerja) ### - Start
    DATA_PEKERJA_FOUND = 0
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where NoPekerja='" & Frm84_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        Frm84_LM_MAKLUMAT_PEKERJA = rs!Samaran & "  |  " & rs!NoPekerja
        DATA_PEKERJA_FOUND = 1
    End If
    
    rs.Close
    Set rs = Nothing
    
    If DATA_PEKERJA_FOUND = 1 Then
        On Error GoTo Err_A:
        Frm84.CBB1 = Frm84_LM_MAKLUMAT_PEKERJA
Restore_A:
    End If
    '### Carian Maklumat Penjual (Data Pekerja) ### - End
End If

Frm84.CBB1.Enabled = True
Frm84.CBB1.BackColor = &HFFFFFF

If DATA_FOUND = 1 Then
    
    Frm84.L85_Text = "1" '0 : Barang baru , 1 : Edit
    
    If Frm84_LM_JUALAN_TRADE = 1 Then '0 : Jualan tanpa trade in , 1 : Jualan dengan trade in
        
        Frm83.CB9 = 1
        Frm83.CB10 = 0
        
        Call Frm83_Initial_Setting
        Call Frm83_initial_setting2
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where bill_No_Trade_In='" & Frm84.L57_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!NoRujukanSistem) Then
                Frm83.L9_Text = rs!NoRujukanSistem 'No. Rujukan Belian
                Frm84.L61_Text = rs!NoRujukanSistem 'No. Rujukan Belian
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        
        Frm83.L41_Text = 1 '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru
        Frm84.L59_Text = 1 '0 : Barang baru , 1 : Edit
        
        Call Frm84_recall_trade_in_data
    
        If Frm83.TB28 = vbNullString Then
            Frm83.CB2 = 1
            Frm83.CB2.Enabled = False
            Frm83.CB3.Enabled = False
            Frm83.CB11.Enabled = False
            Frm83.CB12.Enabled = False
        Else
            Frm83.CB2.Enabled = True
            Frm83.CB3.Enabled = True
            Frm83.CB11.Enabled = True
            Frm83.CB12.Enabled = True
        End If
        
    End If
    
    'Note = "Sila buat pilihan jenis pengiraan upah." & vbCrLf & _
            vbNullString & vbCrLf & _
            "YES : Upah mengikut tetapan per item" & vbCrLf & _
            "NO  : Upah mengikut berat"
    
    'Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    'If Answer = vbNo Then
    '    G_KIRAAN_UPAH = 0
    '    Frm84.L86_Text = "PENGIRAAN UPAH MENGIKUT BERAT"
    'Else
    '    G_KIRAAN_UPAH = 1
    '    Frm84.L86_Text = "PENGIRAAN UPAH MENGIKUT UPAH PER ITEM"
    'End If
    
    Frm84.Show
    Frm83.Hide
    Frm85.Hide
    
    Frm84.Pic3.Visible = True
    Frm84.L41_Text = 1
    frm130.L41_Text = 1
    Frm84.CMD2.Visible = False
    Frm84.CMD5.Visible = False
    Frm84.CMD15.Visible = True
    Frm84.CMD16.Visible = True
End If

Exit Sub
Err_A:
Frm84.CBB1.AddItem Frm84_LM_MAKLUMAT_PEKERJA
Frm84.CBB1 = Frm84_LM_MAKLUMAT_PEKERJA
Resume Restore_A:

Exit Sub
Err_B:
frm130.CBB2.AddItem Frm84_LM_JENIS_KAD
frm130.CBB2 = Frm84_LM_JENIS_KAD
Resume Restore_B:
End Sub
Sub Frm85_search_berat_page()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim Frm85_LM_TOTAL_PAGE As Double

x = 0
Frm85_PAGE_SIZE = 34
Frm85_LM_TOTAL_PAGE = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
    End If
End If

If Frm101.L46_Text = "Semua cawangan" Then

    Frm85_LM_SEARCH_5 = Null
    Frm85_LM_SEARCH_5_LOGIC = "<>"
    Frm85_LM_SEARCH_6 = Null
    Frm85_LM_SEARCH_6_LOGIC = "<>"
    
Else

    Frm85_LM_SEARCH_5 = Frm101.L46_Text
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

Frm85_LM_PAGE_FOUND = 0

Frm85.L10_Text = "Report Semua Item Dengan Berat [" & Format(Frm101.L5_Text, "0.00 g") & "]" 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND Berat='" & Frm101.L5_Text & "' AND StatusItem <> 0 order by tarikh_belian ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1

    If Frm85_LM_PAGE_FOUND = 0 Then
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm85.L79_Text = Frm85.L79_Text + 1
                Frm85_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm85.L79_Text) Then
                    If Frm85.L79_Text <> 1 Then
                        Frm85.L79_Text = Frm85.L79_Text - 1
                        Frm85_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x
        
    With Frm85.LV2.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!tarikh_belian) Then 'Tarikh Belian
            .ListSubItems.Add , , rs!tarikh_belian
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_siri_Produk) Then 'No. Siri Produk
            .ListSubItems.Add , , rs!no_siri_Produk
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kod_Purity) Then 'Purity
            .ListSubItems.Add , , rs!kod_Purity
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kategori_Produk) Then 'Kategori Produk
            .ListSubItems.Add , , rs!kategori_Produk
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!nama_Supplier) Then 'Nama Supplier
            .ListSubItems.Add , , rs!nama_Supplier
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Berat) Then 'Berat (g)
            .ListSubItems.Add , , Format(rs!Berat, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kos_Belian_Gram) Then 'Rate Penerimaan (RM/g)
            .ListSubItems.Add , , Format(rs!kos_Belian_Gram, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!UPAH) Then 'Upah (RM)
            .ListSubItems.Add , , Format(rs!UPAH, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!SpreadValue) Then 'Spread (%)
            .ListSubItems.Add , , Format(rs!SpreadValue, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_lepas_spread) Then 'Harga Selepas Spread (RM)
            .ListSubItems.Add , , Format(rs!harga_lepas_spread, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!adjustment) Then 'Adjustment (RM)
            .ListSubItems.Add , , Format(rs!adjustment, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_item) Then 'Harga Belian (RM) : Tidak Campur Cukai GST
            .ListSubItems.Add , , Format(rs!harga_item, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Upah_Jualan) Then 'Upah Jualan (RM) : Pelanggan
            .ListSubItems.Add , , Format(rs!Upah_Jualan, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!Upah_Member) Then 'Upah Jualan (RM) : Ahli Biasa
            .ListSubItems.Add , , Format(rs!Upah_Member, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!Upah_Pengedar) Then 'Upah Jualan (RM) : Silver
            .ListSubItems.Add , , Format(rs!Upah_Pengedar, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Upah_RAF) Then 'Upah Jualan (RM) : Gold
            .ListSubItems.Add , , Format(rs!Upah_RAF, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!upah_normal_dealer) Then 'Upah Jualan (RM) : Platinum
            .ListSubItems.Add , , Format(rs!upah_normal_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!upah_master_dealer) Then 'Upah Jualan (RM) : Master Dealer
            .ListSubItems.Add , , Format(rs!upah_master_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!code_Supplier) Then 'Tetapan Harga Jualan (RM) : Pelanggan
            .ListSubItems.Add , , Format(rs!code_Supplier, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!HargaJualan_Member) Then 'Tetapan Harga Jualan (RM) : Pelanggan
            .ListSubItems.Add , , Format(rs!HargaJualan_Member, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!HargaJualan_Pengedar) Then 'Tetapan Harga Jualan (RM) : Pengedar
            .ListSubItems.Add , , Format(rs!HargaJualan_Pengedar, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!HargaJualan_RAF) Then 'Tetapan Harga Jualan (RM) : RAF
            .ListSubItems.Add , , Format(rs!HargaJualan_RAF, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!hargajualan_normal_dealer) Then 'Tetapan Harga Jualan (RM) : Normal Dealer
            .ListSubItems.Add , , Format(rs!hargajualan_normal_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!hargajualan_master_dealer) Then 'Tetapan Harga Jualan (RM) : Master Dealer
            .ListSubItems.Add , , Format(rs!hargajualan_master_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
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
        
        If Not IsNull(rs!dimension_Saiz) Then 'Tebal
            .ListSubItems.Add , , rs!dimension_Saiz
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!bill_No_Belian) Then 'No. Invoice Supplier
            .ListSubItems.Add , , rs!bill_No_Belian
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
        
        If Not IsNull(rs!cawangan) Then 'Cawangan
            .ListSubItems.Add , , rs!cawangan
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
rs.Open "select COUNT(ID) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND Berat='" & Frm101.L5_Text & "' AND StatusItem <> 0 order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

'#### Jumlah Bilangan Barang Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(no_siri_Produk) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND Berat='" & Frm101.L5_Text & "' AND StatusItem <> 0 order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L44_Text = rs(0)
Else
    Frm85.L44_Text = 0
End If

rs.Close
Set rs = Nothing
'#### Jumlah Bilangan Barang Keseluruhan #### - End

'#### Jumlah Berat Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(Berat) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND Berat='" & Frm101.L5_Text & "' AND StatusItem <> 0 order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L45_Text = Format(rs(0), "#,##0.00 g")
Else
    Frm85.L45_Text = "0.00 g"
End If

rs.Close
Set rs = Nothing
'#### Jumlah Berat Keseluruhan #### - End

'#### Jumlah Modal Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_item) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND Berat='" & Frm101.L5_Text & "' AND StatusItem <> 0 order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L46_Text = "RM " & Format(rs(0), "#,##0.00")
Else
    Frm85.L46_Text = "RM 0.00"
End If

rs.Close
Set rs = Nothing
'#### Jumlah Modal Keseluruhan #### - End

If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    Frm85.Pic2.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    MsgBox "Tiada Rekod Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_search_invoice_supplier_page()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim Frm85_LM_TOTAL_PAGE As Double

x = 0
Frm85_PAGE_SIZE = 34
Frm85_LM_TOTAL_PAGE = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
    End If
End If

If Frm101.L46_Text = "Semua cawangan" Then

    Frm85_LM_SEARCH_5 = Null
    Frm85_LM_SEARCH_5_LOGIC = "<>"
    Frm85_LM_SEARCH_6 = Null
    Frm85_LM_SEARCH_6_LOGIC = "<>"
    
Else

    Frm85_LM_SEARCH_5 = Frm101.L46_Text
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

Frm85_LM_PAGE_FOUND = 0

Frm85.L10_Text = "Report belian dari No. Invoice [" & Frm101.L5_Text & "]" 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND bill_No_Belian='" & Frm101.L5_Text & "' AND StatusItem <> 0 order by tarikh_belian ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1

    If Frm85_LM_PAGE_FOUND = 0 Then
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm85.L79_Text = Frm85.L79_Text + 1
                Frm85_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm85.L79_Text) Then
                    If Frm85.L79_Text <> 1 Then
                        Frm85.L79_Text = Frm85.L79_Text - 1
                        Frm85_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x
    
    With Frm85.LV2.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!tarikh_belian) Then 'Tarikh Belian
            .ListSubItems.Add , , rs!tarikh_belian
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_siri_Produk) Then 'No. Siri Produk
            .ListSubItems.Add , , rs!no_siri_Produk
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kod_Purity) Then 'Purity
            .ListSubItems.Add , , rs!kod_Purity
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kategori_Produk) Then 'Kategori Produk
            .ListSubItems.Add , , rs!kategori_Produk
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!nama_Supplier) Then 'Nama Supplier
            .ListSubItems.Add , , rs!nama_Supplier
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Berat) Then 'Berat (g)
            .ListSubItems.Add , , Format(rs!Berat, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kos_Belian_Gram) Then 'Rate Penerimaan (RM/g)
            .ListSubItems.Add , , Format(rs!kos_Belian_Gram, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!UPAH) Then 'Upah (RM)
            .ListSubItems.Add , , Format(rs!UPAH, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!SpreadValue) Then 'Spread (%)
            .ListSubItems.Add , , Format(rs!SpreadValue, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_lepas_spread) Then 'Harga Selepas Spread (RM)
            .ListSubItems.Add , , Format(rs!harga_lepas_spread, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!adjustment) Then 'Adjustment (RM)
            .ListSubItems.Add , , Format(rs!adjustment, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_item) Then 'Harga Belian (RM) : Tidak Campur Cukai GST
            .ListSubItems.Add , , Format(rs!harga_item, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Upah_Jualan) Then 'Upah Jualan (RM) : Pelanggan
            .ListSubItems.Add , , Format(rs!Upah_Jualan, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!Upah_Member) Then 'Upah Jualan (RM) : Ahli Biasa
            .ListSubItems.Add , , Format(rs!Upah_Member, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!Upah_Pengedar) Then 'Upah Jualan (RM) : Silver
            .ListSubItems.Add , , Format(rs!Upah_Pengedar, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Upah_RAF) Then 'Upah Jualan (RM) : Gold
            .ListSubItems.Add , , Format(rs!Upah_RAF, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!upah_normal_dealer) Then 'Upah Jualan (RM) : Platinum
            .ListSubItems.Add , , Format(rs!upah_normal_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!upah_master_dealer) Then 'Upah Jualan (RM) : Master Dealer
            .ListSubItems.Add , , Format(rs!upah_master_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!code_Supplier) Then 'Tetapan Harga Jualan (RM) : Pelanggan
            .ListSubItems.Add , , Format(rs!code_Supplier, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!HargaJualan_Member) Then 'Tetapan Harga Jualan (RM) : Pelanggan
            .ListSubItems.Add , , Format(rs!HargaJualan_Member, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!HargaJualan_Pengedar) Then 'Tetapan Harga Jualan (RM) : Pengedar
            .ListSubItems.Add , , Format(rs!HargaJualan_Pengedar, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!HargaJualan_RAF) Then 'Tetapan Harga Jualan (RM) : RAF
            .ListSubItems.Add , , Format(rs!HargaJualan_RAF, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!hargajualan_normal_dealer) Then 'Tetapan Harga Jualan (RM) : Normal Dealer
            .ListSubItems.Add , , Format(rs!hargajualan_normal_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!hargajualan_master_dealer) Then 'Tetapan Harga Jualan (RM) : Master Dealer
            .ListSubItems.Add , , Format(rs!hargajualan_master_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
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
        
        If Not IsNull(rs!dimension_Saiz) Then 'Tebal
            .ListSubItems.Add , , rs!dimension_Saiz
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!bill_No_Belian) Then 'No. Invoice Supplier
            .ListSubItems.Add , , rs!bill_No_Belian
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
        
        If Not IsNull(rs!cawangan) Then 'Cawangan
            .ListSubItems.Add , , rs!cawangan
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
rs.Open "select COUNT(ID) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND bill_No_Belian='" & Frm101.L5_Text & "' AND StatusItem <> 0 order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

'#### Jumlah Bilangan Barang Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(no_siri_Produk) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND bill_No_Belian='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L44_Text = rs(0)
Else
    Frm85.L44_Text = 0
End If
    
rs.Close
Set rs = Nothing
'#### Jumlah Bilangan Barang Keseluruhan #### - End

'#### Jumlah Berat Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(Berat) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND bill_No_Belian='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L45_Text = Format(rs(0), "#,##0.00 g")
Else
    Frm85.L45_Text = "0.00 g"
End If

rs.Close
Set rs = Nothing
'#### Jumlah Berat Keseluruhan #### - End

'#### Jumlah Modal Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_item) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND bill_No_Belian='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L46_Text = "RM " & Format(rs(0), "#,##0.00")
Else
    Frm85.L46_Text = "RM 0.00"
End If

rs.Close
Set rs = Nothing
'#### Jumlah Modal Keseluruhan #### - End

If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    Frm85.Pic2.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    MsgBox "Tiada Rekod Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_carian_jualan_page()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim Frm85_LM_UNTUNG As Double
Dim Frm85_LM_TOTAL_PAGE As Double
Dim Frm85_LM_UNTUNG2 As Double

Frm85_PAGE_SIZE = 34
Frm85_LM_TOTAL_PAGE = 0
x = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0
Frm85_LM_UNTUNG = 0
Frm85_LM_UNTUNG2 = 0

user_level = MDI_frm1.L4_Text

Dim LM_FIELD As String
LM_INVOICE_RASMI = 0

If Frm101.L46_Text = "Semua cawangan" Then

    Frm85_SEARCH_8 = Null
    Frm85_SEARCH_8_LOGIC = "<>"
    Frm85_SEARCH_9 = Null
    Frm85_SEARCH_9_LOGIC = "<>"
    
Else

    Frm85_SEARCH_8 = Frm101.L46_Text
    Frm85_SEARCH_8_LOGIC = "="
    Frm85_SEARCH_9 = "HQ"
    Frm85_SEARCH_9_LOGIC = "="
    
End If

If user_level = "Guest/User" Then
    Frm85_LM_SEARCH_6 = 1
    Frm85_LM_SEARCH_6_LOGIC = "="
    LM_INVOICE_RASMI = 1
    Frm85_LM_SEARCH_7 = 1
    Frm85_LM_SEARCH_7_LOGIC = "="
    
    LM_FIELD = "no_invoice_r"
Else
    Frm85_LM_SEARCH_6 = 0
    Frm85_LM_SEARCH_6_LOGIC = "="
    
    Frm85_LM_SEARCH_7 = 1
    Frm85_LM_SEARCH_7_LOGIC = "="
    
    LM_FIELD = "no_resit"
End If

If user_level = "Administration" Then
    Frm85_LM_SEARCH_10 = 1
    Frm85_LM_SEARCH_10_LOGIC = "="
    
    Frm85_LM_SEARCH_11 = 1
    Frm85_LM_SEARCH_11_LOGIC = "="
Else
    Frm85_LM_SEARCH_10 = 0
    Frm85_LM_SEARCH_10_LOGIC = "="
    
    Frm85_LM_SEARCH_11 = 1
    Frm85_LM_SEARCH_11_LOGIC = "="
End If

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
    End If
End If

Frm85_LM_PAGE_FOUND = 0

'#### Header Report ###
Frm85.L14_Text = "Report Jualan Bagi No. Invoice [" & UCase(Frm101.L5_Text) & "]" 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 23_senarai_jualan where " & LM_FIELD & "='" & UCase(Frm101.L5_Text) & "' AND (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm85_LM_PAGE_FOUND = 0 Then
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm85.L79_Text = Frm85.L79_Text + 1
                Frm85_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm85.L79_Text) Then
                    If Frm85.L79_Text <> 1 Then
                        Frm85.L79_Text = Frm85.L79_Text - 1
                        Frm85_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x
    
    With Frm85.LV1.ListItems.Add(, , rs!ID)
     
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!tarikh) Then 'Tarikh Jualan
            .ListSubItems.Add , , rs!tarikh
        Else
            .ListSubItems.Add , , ""
        End If
        
        If LM_INVOICE_RASMI = 0 Then
        
            If Not IsNull(rs!no_resit) Then .ListSubItems.Add , , rs!no_resit
            If Not IsNull(rs!no_invoice_r) Then .ListSubItems.Add , , rs!no_invoice_r
            
        Else
            .ListSubItems.Add , , ""
        End If
        
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
        
        If Not IsNull(rs!berat_jualan) Then  'Berat Jualan (g)
            .ListSubItems.Add , , Format(rs!berat_jualan, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_Semasa) Then 'Harga Semasa (RM/g)
            .ListSubItems.Add , , Format(rs!harga_Semasa, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!UPAH) Then 'Upah (RM)
            .ListSubItems.Add , , Format(rs!UPAH, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_asal) Then 'Harga Asal (RM)
            .ListSubItems.Add , , Format(rs!harga_asal, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!diskaun) Then 'Diskaun (%)
            .ListSubItems.Add , , Format(rs!diskaun, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!harga_lepas_diskaun) Then 'Harga Selepas Diskaun (RM)
            .ListSubItems.Add , , Format(rs!harga_lepas_diskaun, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!adjustment) Then 'Adjustment (RM)
            .ListSubItems.Add , , Format(rs!adjustment, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_jualan_dengan_gst) Then 'Harga Jualan (RM)
            .ListSubItems.Add , , Format(rs!harga_jualan_dengan_gst, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!gst_ari_nashi) Then
        
            If rs!gst_ari_nashi = "ZR (L)" Then '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                .ListSubItems.Add , , "ZR (L)"
            ElseIf rs!gst_ari_nashi = "SR" Then
                .ListSubItems.Add , , "SR"
            End If
            
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!jumlah_gst) Then 'Jumlah GST (RM)
            .ListSubItems.Add , , Format(rs!jumlah_gst, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!cawangan) Then
            .ListSubItems.Add , , rs!cawangan
        Else
            .ListSubItems.Add , , ""
        End If
        If Not IsNull(rs!nama_pekerja) Then
            .ListSubItems.Add , , rs!nama_pekerja
        Else
            .ListSubItems.Add , , ""
        End If
        If Not IsNull(rs!dulang) Then
            .ListSubItems.Add , , rs!dulang
        Else
            .ListSubItems.Add , , ""
        End If

    End With
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'Frm85.L15_Text = x 'Total Barang
'Frm85.L16_Text = Format(Frm85_LM_BERAT, "#,##0.00 g")  'Total Berat
'Frm85.L17_Text = "RM " & Format(Frm85_LM_HARGA, "#,##0.00")  'Total Harga Jualan
'Frm85.L18_Text = "RM " & Format(Frm85_LM_UNTUNG, "#,##0.00")  'Total Keuntungan Jualan
'Frm85.L87_Text = "RM " & Format(Frm85_LM_UNTUNG2, "#,##0.00")  'Total Keuntungan Jualan 2

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 23_senarai_jualan where " & LM_FIELD & "='" & UCase(Frm101.L5_Text) & "' AND (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

'#### Jumlah Bilangan Barang Terjual Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(no_siri_Produk) from 23_senarai_jualan where " & LM_FIELD & "='" & UCase(Frm101.L5_Text) & "' AND (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L47_Text = rs(0)
Else
    Frm85.L47_Text = 0
End If

rs.Close
Set rs = Nothing
'#### Jumlah Bilangan Barang Terjual Keseluruhan #### - End

'#### Jumlah Berat Jualan Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(Berat_Jualan) from 23_senarai_jualan where " & LM_FIELD & "='" & UCase(Frm101.L5_Text) & "' AND (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L48_Text = Format(rs(0), "#,##0.00 g")
Else
    Frm85.L48_Text = "0.00 g"
End If

rs.Close
Set rs = Nothing
'#### Jumlah Berat Jualan Keseluruhan #### - End

'#### Jumlah Harga Jualan Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_dengan_gst) from 23_senarai_jualan where " & LM_FIELD & "='" & UCase(Frm101.L5_Text) & "' AND (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L49_Text = "RM " & Format(rs.Fields(0), "#,##0.00")
Else
    Frm85.L49_Text = "RM 0.00"
End If

rs.Close
Set rs = Nothing
'#### Jumlah Harga Jualan Keseluruhan #### - End

'#### Jumlah Keuntungan Keseluruhan 1 #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(untung) from 23_senarai_jualan where " & LM_FIELD & "='" & UCase(Frm101.L5_Text) & "' AND (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then
    Frm85.L50_Text = "RM " & Format(rs(0), "#,##0.00")
Else
    Frm85.L50_Text = "RM 0.00"
End If

rs.Close
Set rs = Nothing
'#### Jumlah Keuntungan Keseluruhan 1 #### - End

'#### Jumlah Keuntungan Keseluruhan 2 #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(untung2) from 23_senarai_jualan where " & LM_FIELD & "='" & UCase(Frm101.L5_Text) & "' AND (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then
    Frm85.L88_Text = "RM " & Format(rs(0), "#,##0.00")
Else
    Frm85.L88_Text = "RM 0.00"
End If

rs.Close
Set rs = Nothing
'#### Jumlah Keuntungan Keseluruhan 2 #### - End

If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    
    Frm85.Pic3.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else

    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    
    MsgBox "Tiada Rekod No. Invoice Ini Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_carian_buyback_page()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim Frm85_LM_TOTAL_PAGE As Double

Frm85_LM_TOTAL_PAGE = 0
Frm85_PAGE_SIZE = 34
x = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
    End If
End If

If Frm101.L46_Text = "Semua cawangan" Then

    Frm85_LM_SEARCH_5 = Null
    Frm85_LM_SEARCH_5_LOGIC = "<>"
    Frm85_LM_SEARCH_6 = Null
    Frm85_LM_SEARCH_6_LOGIC = "<>"
    
Else

    Frm85_LM_SEARCH_5 = Frm101.L46_Text
    Frm85_LM_SEARCH_5_LOGIC = "="
    Frm85_LM_SEARCH_6 = "HQ"
    Frm85_LM_SEARCH_6_LOGIC = "="
    
End If

Frm85_LM_PAGE_FOUND = 0

'#### Header Report ###
Frm85.L19_Text = "Report Dari No. Voucher Buyback / Trade In [" & UCase(Frm101.L5_Text) & "]" 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from Data_Database where bill_No_Trade_In='" & UCase(Frm101.L5_Text) & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 order by tarikh_belian ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    
    If Frm85_LM_PAGE_FOUND = 0 Then
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm85.L79_Text = Frm85.L79_Text + 1
                Frm85_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm85.L79_Text) Then
                    If Frm85.L79_Text <> 1 Then
                        Frm85.L79_Text = Frm85.L79_Text - 1
                        Frm85_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x
        
    With Frm85.LV3.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!tarikh_belian) Then 'Tarikh Belian
            .ListSubItems.Add , , rs!tarikh_belian
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_siri_Produk) Then 'No. Siri Produk
            .ListSubItems.Add , , rs!no_siri_Produk
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kod_Purity) Then 'Purity
            .ListSubItems.Add , , rs!kod_Purity
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kategori_Produk) Then 'Kategori Produk
            .ListSubItems.Add , , rs!kategori_Produk
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!nama_Supplier) Then 'Nama Supplier
            .ListSubItems.Add , , rs!nama_Supplier
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Berat) Then 'Berat (g)
            .ListSubItems.Add , , Format(rs!Berat, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kos_Belian_Gram) Then 'Rate Penerimaan (RM/g)
            .ListSubItems.Add , , Format(rs!kos_Belian_Gram, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!UPAH) Then 'Upah (RM)
            .ListSubItems.Add , , Format(rs!UPAH, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!SpreadValue) Then 'Spread (%)
            .ListSubItems.Add , , Format(rs!SpreadValue, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_lepas_spread) Then 'Harga Selepas Spread (RM)
            .ListSubItems.Add , , Format(rs!harga_lepas_spread, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!adjustment) Then 'Adjustment (RM)
            .ListSubItems.Add , , Format(rs!adjustment, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_item) Then 'Harga Belian (RM) : Tidak Campur Cukai GST
            .ListSubItems.Add , , Format(rs!harga_item, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Upah_Jualan) Then 'Upah Jualan (RM) : Pelanggan
            .ListSubItems.Add , , Format(rs!Upah_Jualan, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!Upah_Member) Then 'Upah Jualan (RM) : Ahli Biasa
            .ListSubItems.Add , , Format(rs!Upah_Member, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!Upah_Pengedar) Then 'Upah Jualan (RM) : Silver
            .ListSubItems.Add , , Format(rs!Upah_Pengedar, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Upah_RAF) Then 'Upah Jualan (RM) : Gold
            .ListSubItems.Add , , Format(rs!Upah_RAF, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!upah_normal_dealer) Then 'Upah Jualan (RM) : Platinum
            .ListSubItems.Add , , Format(rs!upah_normal_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!upah_master_dealer) Then 'Upah Jualan (RM) : Master Dealer
            .ListSubItems.Add , , Format(rs!upah_master_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!code_Supplier) Then 'Tetapan Harga Jualan (RM) : Pelanggan
            .ListSubItems.Add , , Format(rs!code_Supplier, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!HargaJualan_Member) Then 'Tetapan Harga Jualan (RM) : Pelanggan
            .ListSubItems.Add , , Format(rs!HargaJualan_Member, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!HargaJualan_Pengedar) Then 'Tetapan Harga Jualan (RM) : Pengedar
            .ListSubItems.Add , , Format(rs!HargaJualan_Pengedar, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!HargaJualan_RAF) Then 'Tetapan Harga Jualan (RM) : RAF
            .ListSubItems.Add , , Format(rs!HargaJualan_RAF, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!hargajualan_normal_dealer) Then 'Tetapan Harga Jualan (RM) : Normal Dealer
            .ListSubItems.Add , , Format(rs!hargajualan_normal_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!hargajualan_master_dealer) Then 'Tetapan Harga Jualan (RM) : Master Dealer
            .ListSubItems.Add , , Format(rs!hargajualan_master_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
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
        
        If Not IsNull(rs!dimension_Saiz) Then 'Tebal
            .ListSubItems.Add , , rs!dimension_Saiz
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!bill_No_Belian) Then 'No invoice supplier
            .ListSubItems.Add , , rs!bill_No_Belian
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
        
        If Not IsNull(rs!cawangan) Then 'Cawangan
            .ListSubItems.Add , , rs!cawangan
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
rs.Open "select COUNT(ID) from Data_Database where bill_No_Trade_In='" & UCase(Frm101.L5_Text) & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

'#### Jumlah Bilangan Barang Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(no_siri_Produk) from Data_Database where bill_No_Trade_In='" & UCase(Frm101.L5_Text) & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L51_Text = rs(0)
Else
    Frm85.L51_Text = 0
End If
'#### Jumlah Bilangan Barang Keseluruhan #### - End

'#### Jumlah Berat Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(Berat) from Data_Database where bill_No_Trade_In='" & UCase(Frm101.L5_Text) & "' AND StatusItem <> 0 order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L52_Text = Format(rs(0), "#,##0.00 g")
Else
    Frm85.L52_Text = "0.00 g"
End If
'#### Jumlah Berat Keseluruhan #### - End

'#### Jumlah Modal Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_item) from Data_Database where bill_No_Trade_In='" & UCase(Frm101.L5_Text) & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L53_Text = "RM " & Format(rs(0), "#,##0.00")
Else
    Frm85.L53_Text = "RM 0.00"
End If
'#### Jumlah Modal Keseluruhan #### - End

If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    
    Frm85.Pic4.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    
    MsgBox "Tiada Rekod No. Resit Buyback / Trade In Ini.", vbInformation, "Info"
End If
End Sub
Sub Frm85_Header_Report_Potong()
'on error resume next
With Frm85.LV5
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear

    Frm85.LV5.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "Tarikh Potong", 1800
    .ColumnHeaders.Add 5, , "No. Siri Produk", 2200
    .ColumnHeaders.Add 6, , "Purity", 1500
    .ColumnHeaders.Add 7, , "Kategori Produk", 3500
    .ColumnHeaders.Add 8, , "Supplier", 4400
    .ColumnHeaders.Add 9, , "Berat Asal (g)", 1500, 1
    .ColumnHeaders.Add 10, , "Susut Berat (g)", 1500, 1
    .ColumnHeaders.Add 11, , "Baki Berat (g)", 1500, 1
    .ColumnHeaders.Add 12, , "Dulang", 1000, 2
    .ColumnHeaders.Add 13, , "Cawangan", 3500
    .ColumnHeaders.Add 14, , "Nama Pekerja", 2500
    
End With
End Sub
Sub Frm85_report_potong_page()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date
Dim Frm85_LM_TOTAL_PAGE As Double

Frm85_LM_TOTAL_PAGE = 0
Frm85_PAGE_SIZE = 34
x = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If
If Frm101.L6_Text = "Semua Kategori Produk" Then
    Frm85_LM_SEARCH_2 = Null
    Frm85_LM_SEARCH_2_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_2 = Frm101.L6_Text
    Frm85_LM_SEARCH_2_LOGIC = "="
End If
If Frm101.L34_Text = "Semua Dulang" Then
    Frm85_LM_SEARCH_3 = Null
    Frm85_LM_SEARCH_3_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_3 = Frm101.L34_Text
    Frm85_LM_SEARCH_3_LOGIC = "="
End If
If Frm101.L37_Text = "Semua Supplier" Then
    Frm85_LM_SEARCH_4 = Null
    Frm85_LM_SEARCH_4_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_4 = Frm101.L37_Text
    Frm85_LM_SEARCH_4_LOGIC = "="
End If
If Frm101.L46_Text = "Semua cawangan" Then

    Frm85_LM_SEARCH_5 = Null
    Frm85_LM_SEARCH_5_LOGIC = "<>"
    Frm85_LM_SEARCH_6 = Null
    Frm85_LM_SEARCH_6_LOGIC = "<>"
    
Else

    Frm85_LM_SEARCH_5 = Frm101.L46_Text
    Frm85_LM_SEARCH_5_LOGIC = "="
    Frm85_LM_SEARCH_6 = "HQ"
    Frm85_LM_SEARCH_6_LOGIC = "="
    
End If
If Frm101.L47_Text = "Semua" Then
    Frm85_LM_SEARCH_20 = Null
    Frm85_LM_SEARCH_20_LOGIC = "<>"
Else
    If InStr(1, Frm101.L47_Text, "  |  ") <> 0 Then
        Frm84_LM_EMP_NO = Split(Frm101.L47_Text, "  |  ")(1)
        Frm84_LM_EMP_NAMA = Split(Frm101.L47_Text, "  |  ")(0)
    End If
    Frm85_LM_SEARCH_20 = Frm84_LM_EMP_NAMA
    Frm85_LM_SEARCH_20_LOGIC = "="
End If

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
        End If
    End If
End If

Frm85_LM_PAGE_FOUND = 0

'#### Header Report ###
If Frm101.L9_Text = 0 Then Frm85.L27_Text = "Report Potong Bagi Supplier [" & Frm101.L37_Text & "] , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "]." 'Report Header"
If Frm101.L9_Text = 1 Then Frm85.L27_Text = "Report Potong Bagi Supplier [" & Frm101.L37_Text & "] , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "] Dari " & TM & " Hingga " & TA 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select * from Data_Database where nama_pekerja_potong " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND (StatusItem='" & 12 & "' OR StatusItem='" & 20 & "' OR StatusItem='" & 28 & "' OR StatusItem='" & 22 & "') order by tarikh_Jualan1 ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select * from Data_Database where nama_pekerja_potong " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh_Jualan1 BETWEEN '" & TM & "' AND '" & TA & "' AND (StatusItem='" & 12 & "' OR StatusItem='" & 20 & "' OR StatusItem='" & 28 & "' OR StatusItem='" & 22 & "') order by tarikh_Jualan1 ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm85_LM_PAGE_FOUND = 0 Then
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm85.L79_Text = Frm85.L79_Text + 1
                Frm85_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm85.L79_Text) Then
                    If Frm85.L79_Text <> 1 Then
                        Frm85.L79_Text = Frm85.L79_Text - 1
                        Frm85_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x
    
    With Frm85.LV5.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!tarikh_jualan1) Then 'Tarikh potong
            .ListSubItems.Add , , rs!tarikh_jualan1
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_siri_Produk) Then 'No. Siri Produk
            .ListSubItems.Add , , rs!no_siri_Produk
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kod_Purity) Then 'Purity
            .ListSubItems.Add , , rs!kod_Purity
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kategori_Produk) Then 'Kategori Produk
            .ListSubItems.Add , , rs!kategori_Produk
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!nama_Supplier) Then 'Nama Supplier
            .ListSubItems.Add , , rs!nama_Supplier
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!Berat) Then 'Berat (g)
            .ListSubItems.Add , , Format(rs!Berat, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!susut_berat) Then 'Susut berat (g)
            .ListSubItems.Add , , Format(rs!susut_berat, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!beza_berat) Then 'Susut berat (g)
            .ListSubItems.Add , , Format(rs!beza_berat, "#,##0.00")
            If IsNumeric(rs!beza_berat) Then Frm85_LM_BERAT = Frm85_LM_BERAT + rs!beza_berat 'Total Berat (g)
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!dulang) Then 'Dulang
            .ListSubItems.Add , , rs!dulang
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!cawangan) Then 'Cawangan
            .ListSubItems.Add , , rs!cawangan
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!nama_pekerja_potong) Then
            .ListSubItems.Add , , rs!nama_pekerja_potong
        Else
            .ListSubItems.Add , , ""
        End If
        
    End With

    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'Frm85.L28_Text = x 'Total Barang
'Frm85.L29_Text = Format(Frm85_LM_BERAT, "#,##0.00 g") 'Total Berat
'Frm85.L13_Text = Format(Frm85_LM_HARGA, "#,##0.00") 'Total Harga Belian

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select COUNT(ID) , SUM(beza_berat) from Data_Database where nama_pekerja_potong " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND (StatusItem='" & 12 & "' OR StatusItem='" & 20 & "' OR StatusItem='" & 28 & "' OR StatusItem='" & 22 & "')", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select COUNT(ID) , SUM(beza_berat) from Data_Database where nama_pekerja_potong " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh_Jualan1 BETWEEN '" & TM & "' AND '" & TA & "' AND (StatusItem='" & 12 & "' OR StatusItem='" & 20 & "' OR StatusItem='" & 28 & "' OR StatusItem='" & 22 & "')", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

If Not IsNull(rs(0)) Then
    Frm85.L57_Text = rs(0)
Else
    Frm85.L57_Text = 0
End If
If Not IsNull(rs(1)) Then
    Frm85.L58_Text = Format(rs(1), "#,##0.00 g")
Else
    Frm85.L58_Text = "0.00 g"
End If

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    
    Frm85.Pic6.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    
    MsgBox "Tiada Rekod Stok Potong Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_summary_report_belian()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

Frm101.L5_Text = Frm101.CBB1 'Purity
Frm101.L6_Text = Frm101.CBB2 'Kategori Produk
Frm101.L34_Text = Frm101.CBB3 'Dulang
Frm101.L37_Text = Frm101.CBB4 'Supplier

Frm101.L7_Text = Frm101.DTPicker1 'Tarikh Mula
Frm101.L8_Text = Frm101.DTPicker2 'Tarikh Akhir

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


'### Reset maklumat kedai ### - Start
Report43.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report43.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report43.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report43.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report43.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
    
    LM_NAMA_HEADER = "HQ"
    
Else
    
    LM_NAMA_HEADER = MDI_frm1.L20_Text
    
End If
        
'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & LM_NAMA_HEADER & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report43.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report43.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report43.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report43.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report43.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

x = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If
If Frm101.L6_Text = "Semua Kategori Produk" Then
    Frm85_LM_SEARCH_2 = Null
    Frm85_LM_SEARCH_2_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_2 = Frm101.L6_Text
    Frm85_LM_SEARCH_2_LOGIC = "="
End If
If Frm101.L34_Text = "Semua Dulang" Then
    Frm85_LM_SEARCH_3 = Null
    Frm85_LM_SEARCH_3_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_3 = Frm101.L34_Text
    Frm85_LM_SEARCH_3_LOGIC = "="
End If
If Frm101.L37_Text = "Semua Supplier" Then
    Frm85_LM_SEARCH_4 = Null
    Frm85_LM_SEARCH_4_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_4 = Frm101.L37_Text
    Frm85_LM_SEARCH_4_LOGIC = "="
End If
If Frm101.L46_Text = "Semua cawangan" Then

    Frm85_LM_SEARCH_5 = Null
    Frm85_LM_SEARCH_5_LOGIC = "<>"
    Frm85_LM_SEARCH_6 = Null
    Frm85_LM_SEARCH_6_LOGIC = "<>"
    
Else

    Frm85_LM_SEARCH_5 = Frm101.L46_Text
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

'#### Header Report ###
If Frm101.L9_Text = 0 Then Report43.Sections("Section4").Controls("L1").Caption = "Report Belian Bagi Supplier [" & Frm101.L37_Text & "] , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "]." 'Report Header"
If Frm101.L9_Text = 1 Then Report43.Sections("Section4").Controls("L1").Caption = "Report Belian Bagi Supplier [" & Frm101.L37_Text & "] , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "] Dari " & TM & " Hingga " & TA & "." 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select * from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND (receiving_Status='" & 1 & "' OR receiving_Status='" & 0 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select * from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND (receiving_Status='" & 1 & "' OR receiving_Status='" & 0 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Not IsNull(rs!Berat) Then
        If IsNumeric(rs!Berat) Then Frm85_LM_BERAT = Frm85_LM_BERAT + rs!Berat
    End If
    If Not IsNull(rs!harga_item) Then
        If IsNumeric(rs!harga_item) Then Frm85_LM_HARGA = Frm85_LM_HARGA + rs!harga_item
    End If
    Set Report43.DataSource = rs
    Report43.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing

Report43.Sections("Section5").Controls("L3").Caption = x 'Qty Barang
Report43.Sections("Section5").Controls("L4").Caption = Format(Frm85_LM_BERAT, "#,##0.00 g") 'Jumlah Berat (g)
Report43.Sections("Section5").Controls("L5").Caption = Format(Frm85_LM_HARGA, "#,##0.00") 'Jumlah Harga RM

If x = 0 Then
    MsgBox "Tiada Rekod Belian Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_summary_report_belian_gb()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

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

Frm101.L5_Text = Frm101.CBB1 'Purity
Frm101.L6_Text = Frm101.CBB2 'Kategori Produk
Frm101.L34_Text = Frm101.CBB3 'Dulang

Frm101.L7_Text = Frm101.DTPicker1 'Tarikh Mula
Frm101.L8_Text = Frm101.DTPicker2 'Tarikh Akhir

'### Reset maklumat kedai ### - Start
Report43.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report43.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report43.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report43.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report43.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report43.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report43.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report43.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report43.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report43.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End
    
x = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If
If Frm101.L6_Text = "Semua Kategori Produk" Then
    Frm85_LM_SEARCH_2 = Null
    Frm85_LM_SEARCH_2_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_2 = Frm101.L6_Text
    Frm85_LM_SEARCH_2_LOGIC = "="
End If
If Frm101.L34_Text = "Semua Dulang" Then
    Frm85_LM_SEARCH_3 = Null
    Frm85_LM_SEARCH_3_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_3 = Frm101.L34_Text
    Frm85_LM_SEARCH_3_LOGIC = "="
End If
If Frm101.L37_Text = "Semua Supplier" Then
    Frm85_LM_SEARCH_4 = Null
    Frm85_LM_SEARCH_4_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_4 = Frm101.L37_Text
    Frm85_LM_SEARCH_4_LOGIC = "="
End If


'#### Header Report ###
If Frm101.L9_Text = 0 Then Report43.Sections("Section4").Controls("L1").Caption = "Report Belian Gold Bar Bagi Supplier [" & Frm101.L37_Text & "] , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "]" 'Report Header"
If Frm101.L9_Text = 1 Then Report43.Sections("Section4").Controls("L1").Caption = "Report Belian Gold Bar Bagi Supplier [" & Frm101.L37_Text & "] , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] Dari " & TM & " Hingga " & TA 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select * from Data_Database where nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND receiving_Status='" & 4 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select * from Data_Database where nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND receiving_Status='" & 4 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Not IsNull(rs!Berat) Then
        If IsNumeric(rs!Berat) Then Frm85_LM_BERAT = Frm85_LM_BERAT + rs!Berat
    End If
    If Not IsNull(rs!harga_item) Then
        If IsNumeric(rs!harga_item) Then Frm85_LM_HARGA = Frm85_LM_HARGA + rs!harga_item
    End If
    Set Report43.DataSource = rs
    Report43.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing

Report43.Sections("Section5").Controls("L3").Caption = x 'Qty Barang
Report43.Sections("Section5").Controls("L4").Caption = Format(Frm85_LM_BERAT, "#,##0.00 g") 'Jumlah Berat (g)
Report43.Sections("Section5").Controls("L5").Caption = Format(Frm85_LM_HARGA, "#,##0.00") 'Jumlah Harga RM

If x = 0 Then
    MsgBox "Tiada Rekod Belian Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_summary_report_stok()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

x = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

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

'### Reset maklumat kedai ### - Start
Report44.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report44.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report44.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report44.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report44.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
    
    LM_NAMA_HEADER = "HQ"
    
Else
    
    LM_NAMA_HEADER = MDI_frm1.L20_Text
    
End If

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & LM_NAMA_HEADER & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report44.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report44.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report44.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report44.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report44.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If
If Frm101.L6_Text = "Semua Kategori Produk" Then
    Frm85_LM_SEARCH_2 = Null
    Frm85_LM_SEARCH_2_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_2 = Frm101.L6_Text
    Frm85_LM_SEARCH_2_LOGIC = "="
End If
If Frm101.L34_Text = "Semua Dulang" Then
    Frm85_LM_SEARCH_3 = Null
    Frm85_LM_SEARCH_3_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_3 = Frm101.L34_Text
    Frm85_LM_SEARCH_3_LOGIC = "="
End If
If Frm101.L37_Text = "Semua Supplier" Then
    Frm85_LM_SEARCH_4 = Null
    Frm85_LM_SEARCH_4_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_4 = Frm101.L37_Text
    Frm85_LM_SEARCH_4_LOGIC = "="
End If
If Frm101.L46_Text = "Semua cawangan" Then

    Frm85_LM_SEARCH_5 = Null
    Frm85_LM_SEARCH_5_LOGIC = "<>"
    Frm85_LM_SEARCH_6 = Null
    Frm85_LM_SEARCH_6_LOGIC = "<>"
    
Else

    Frm85_LM_SEARCH_5 = Frm101.L46_Text
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

'#### Header Report ###
If Frm101.L9_Text = 0 Then Report44.Sections("Section4").Controls("L1").Caption = "Report Stok Bagi Supplier [" & Frm101.L37_Text & "] , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "]." 'Report Header"
If Frm101.L9_Text = 1 Then Report44.Sections("Section4").Controls("L1").Caption = "Report Stok Bagi Supplier [" & Frm101.L37_Text & "] , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "] Dari " & TM & " Hingga " & TA 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select * from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "') AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND (receiving_Status='" & 0 & "' OR receiving_Status='" & 1 & "' OR receiving_Status='" & 4 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select * from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "') AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 0 & "' OR receiving_Status='" & 1 & "' OR receiving_Status='" & 4 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Not IsNull(rs!Berat) Then
        If IsNumeric(rs!Berat) Then Frm85_LM_BERAT = Frm85_LM_BERAT + rs!Berat
    End If
    If Not IsNull(rs!harga_item) Then
        If IsNumeric(rs!harga_item) Then Frm85_LM_HARGA = Frm85_LM_HARGA + rs!harga_item
    End If
    Set Report44.DataSource = rs
    Report44.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing

Report44.Sections("Section5").Controls("L3").Caption = x 'Qty Barang
Report44.Sections("Section5").Controls("L4").Caption = Format(Frm85_LM_BERAT, "#,##0.00 g") 'Jumlah Berat (g)
Report44.Sections("Section5").Controls("L5").Caption = Format(Frm85_LM_HARGA, "#,##0.00") 'Jumlah Harga RM

If x = 0 Then
    MsgBox "Tiada Data Stok Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_summary_report_potong()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

x = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

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

If Frm101.L46_Text = "Semua cawangan" Then

    Frm85_LM_SEARCH_5 = Null
    Frm85_LM_SEARCH_5_LOGIC = "<>"
    Frm85_LM_SEARCH_6 = Null
    Frm85_LM_SEARCH_6_LOGIC = "<>"
    
Else

    Frm85_LM_SEARCH_5 = Frm101.L46_Text
    Frm85_LM_SEARCH_5_LOGIC = "="
    Frm85_LM_SEARCH_6 = "HQ"
    Frm85_LM_SEARCH_6_LOGIC = "="
    
End If
If Frm101.L47_Text = "Semua" Then
    Frm85_LM_SEARCH_20 = Null
    Frm85_LM_SEARCH_20_LOGIC = "<>"
Else
    If InStr(1, Frm101.L47_Text, "  |  ") <> 0 Then
        Frm84_LM_EMP_NO = Split(Frm101.L47_Text, "  |  ")(1)
        Frm84_LM_EMP_NAMA = Split(Frm101.L47_Text, "  |  ")(0)
    End If
    Frm85_LM_SEARCH_20 = Frm84_LM_EMP_NAMA
    Frm85_LM_SEARCH_20_LOGIC = "="
End If

'### Reset maklumat kedai ### - Start
Report47.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report47.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report47.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report47.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report47.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
    LM_NAMA_HEADER = "HQ"
Else
    LM_NAMA_HEADER = MDI_frm1.L20_Text
End If

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & LM_NAMA_HEADER & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report47.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report47.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report47.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report47.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report47.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If
If Frm101.L6_Text = "Semua Kategori Produk" Then
    Frm85_LM_SEARCH_2 = Null
    Frm85_LM_SEARCH_2_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_2 = Frm101.L6_Text
    Frm85_LM_SEARCH_2_LOGIC = "="
End If
If Frm101.L34_Text = "Semua Dulang" Then
    Frm85_LM_SEARCH_3 = Null
    Frm85_LM_SEARCH_3_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_3 = Frm101.L34_Text
    Frm85_LM_SEARCH_3_LOGIC = "="
End If
If Frm101.L37_Text = "Semua Supplier" Then
    Frm85_LM_SEARCH_4 = Null
    Frm85_LM_SEARCH_4_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_4 = Frm101.L37_Text
    Frm85_LM_SEARCH_4_LOGIC = "="
End If

'#### Header Report ###
If Frm101.L9_Text = 0 Then Report47.Sections("Section4").Controls("L1").Caption = "Report Potong Bagi Supplier [" & Frm101.L37_Text & "] , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "]" 'Report Header"
If Frm101.L9_Text = 1 Then Report47.Sections("Section4").Controls("L1").Caption = "Report Potong Bagi Supplier [" & Frm101.L37_Text & "] , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] Dari " & TM & " Hingga " & TA 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select * from Data_Database where nama_pekerja_potong " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND (StatusItem='" & 12 & "' OR StatusItem='" & 20 & "' OR StatusItem='" & 28 & "' OR StatusItem='" & 22 & "') order by tarikh_Jualan1 ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select * from Data_Database where nama_pekerja_potong " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh_Jualan1 BETWEEN '" & TM & "' AND '" & TA & "' AND (StatusItem='" & 12 & "' OR StatusItem='" & 20 & "' OR StatusItem='" & 28 & "' OR StatusItem='" & 22 & "') order by tarikh_Jualan1 ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Not IsNull(rs!beza_berat) Then
        If IsNumeric(rs!beza_berat) Then Frm85_LM_BERAT = Frm85_LM_BERAT + rs!beza_berat 'Total Berat (g)
    End If
    
    Set Report47.DataSource = rs
    Report47.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing

Report47.Sections("Section5").Controls("L3").Caption = x 'Total Barang
Report47.Sections("Section5").Controls("L4").Caption = Format(Frm85_LM_BERAT, "#,##0.00 g") 'Total Berat

If x = 0 Then
    MsgBox "Tiada Rekod Stok Potong Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_summary_report_buyback()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

x = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

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

'### Reset maklumat kedai ### - Start
Report45.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report45.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report45.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report45.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report45.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
    
    LM_NAMA_HEADER = "HQ"
    
Else
    
    LM_NAMA_HEADER = MDI_frm1.L20_Text
    
End If

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & LM_NAMA_HEADER & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report45.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report45.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report45.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report45.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report45.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If
If Frm101.L6_Text = "Semua Kategori Produk" Then
    Frm85_LM_SEARCH_2 = Null
    Frm85_LM_SEARCH_2_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_2 = Frm101.L6_Text
    Frm85_LM_SEARCH_2_LOGIC = "="
End If
If Frm101.L34_Text = "Semua Dulang" Then
    Frm85_LM_SEARCH_3 = Null
    Frm85_LM_SEARCH_3_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_3 = Frm101.L34_Text
    Frm85_LM_SEARCH_3_LOGIC = "="
End If
If Frm101.L46_Text = "Semua cawangan" Then

    Frm85_LM_SEARCH_5 = Null
    Frm85_LM_SEARCH_5_LOGIC = "<>"
    Frm85_LM_SEARCH_6 = Null
    Frm85_LM_SEARCH_6_LOGIC = "<>"
    
Else

    Frm85_LM_SEARCH_5 = Frm101.L46_Text
    Frm85_LM_SEARCH_5_LOGIC = "="
    Frm85_LM_SEARCH_6 = "HQ"
    Frm85_LM_SEARCH_6_LOGIC = "="
    
End If

'#### Header Report ###
If Frm101.L9_Text = 0 Then Report45.Sections("Section4").Controls("L1").Caption = "Report Belian Buyback / Trade In Bagi Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "]." 'Report Header"
If Frm101.L9_Text = 1 Then Report45.Sections("Section4").Controls("L1").Caption = "Report Belian Buyback / Trade In Bagi Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "] Dari " & TM & " Hingga " & TA 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select * from Data_Database where kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "') order by tarikh_belian ASC , bill_No_Trade_In ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select * from Data_Database where kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "') order by tarikh_belian ASC , bill_No_Trade_In ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Not IsNull(rs!Berat) Then
        If IsNumeric(rs!Berat) Then Frm85_LM_BERAT = Frm85_LM_BERAT + rs!Berat
    End If
    If Not IsNull(rs!harga_item) Then
        If IsNumeric(rs!harga_item) Then Frm85_LM_HARGA = Frm85_LM_HARGA + rs!harga_item
    End If
    Set Report45.DataSource = rs
    Report45.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing

Report45.Sections("Section5").Controls("L3").Caption = x 'Qty Barang
Report45.Sections("Section5").Controls("L4").Caption = Format(Frm85_LM_BERAT, "#,##0.00 g") 'Jumlah Berat (g)
Report45.Sections("Section5").Controls("L5").Caption = Format(Frm85_LM_HARGA, "#,##0.00") 'Jumlah Harga RM

If x = 0 Then
    MsgBox "Tiada Rekod Belian Buyback / Trade In Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_summary_report_buyback_gb()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

x = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

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

'### Reset maklumat kedai ### - Start
Report45.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report45.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report45.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report45.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report45.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report45.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report45.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report45.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report45.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report45.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If
If Frm101.L6_Text = "Semua Kategori Produk" Then
    Frm85_LM_SEARCH_2 = Null
    Frm85_LM_SEARCH_2_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_2 = Frm101.L6_Text
    Frm85_LM_SEARCH_2_LOGIC = "="
End If
If Frm101.L34_Text = "Semua Dulang" Then
    Frm85_LM_SEARCH_3 = Null
    Frm85_LM_SEARCH_3_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_3 = Frm101.L34_Text
    Frm85_LM_SEARCH_3_LOGIC = "="
End If
If Frm101.L37_Text = "Semua Supplier" Then
    Frm85_LM_SEARCH_4 = Null
    Frm85_LM_SEARCH_4_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_4 = Frm101.L37_Text
    Frm85_LM_SEARCH_4_LOGIC = "="
End If

'#### Header Report ###
If Frm101.L9_Text = 0 Then Report45.Sections("Section4").Controls("L1").Caption = "Report Buyback / Trade In Gold Bar Bagi Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "]" 'Report Header"
If Frm101.L9_Text = 1 Then Report45.Sections("Section4").Controls("L1").Caption = "Report Buyback / Trade In Gold Bar Bagi Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] Dari " & TM & " Hingga " & TA 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select * from Data_Database where kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select * from Data_Database where kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Not IsNull(rs!Berat) Then
        If IsNumeric(rs!Berat) Then Frm85_LM_BERAT = Frm85_LM_BERAT + rs!Berat
    End If
    If Not IsNull(rs!harga_item) Then
        If IsNumeric(rs!harga_item) Then Frm85_LM_HARGA = Frm85_LM_HARGA + rs!harga_item
    End If
    Set Report45.DataSource = rs
    Report45.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing

Report45.Sections("Section5").Controls("L3").Caption = x 'Qty Barang
Report45.Sections("Section5").Controls("L4").Caption = Format(Frm85_LM_BERAT, "#,##0.00 g") 'Jumlah Berat (g)
Report45.Sections("Section5").Controls("L5").Caption = Format(Frm85_LM_HARGA, "#,##0.00") 'Jumlah Harga RM

If x = 0 Then
    MsgBox "Tiada Rekod Belian Buyback / Trade In Gold Bar Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_summary_report_trade_in()
'on error resume next
Dim rs1 As ADODB.Recordset
Dim TA As Date
Dim TM As Date

x = 0

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
''        Set Printer = oPrn
'        PRINTER_FOUND = 1 '0 : Printer Not Found , 1 : Printer Found
'        Exit For
'    End If
'Next

'### Reset maklumat kedai ### - Start
Report58.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report58.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report58.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report58.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report58.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
    
    LM_NAMA_HEADER = "HQ"
    
Else
    
    LM_NAMA_HEADER = MDI_frm1.L20_Text
    
End If

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & LM_NAMA_HEADER & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report58.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report58.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report58.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report58.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report58.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Report58.Sections("Section5").Controls("L1").Caption = vbNullString
Report58.Sections("Section4").Controls("L2").Caption = vbNullString

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If

'#### Header Report ###
If Frm101.L9_Text = 0 Then Report58.Sections("Section4").Controls("L2").Caption = "Report trade in dari agen bagi purity [" & Frm101.L5_Text & "]." 'Report Header"
If Frm101.L9_Text = 1 Then Report58.Sections("Section4").Controls("L2").Caption = "Report trade in dari agen bagi purity [" & Frm101.L5_Text & "] dari " & TM & " hingga " & TA & "."  'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select * from 50_belian_emas_agen where status='" & 1 & "' AND kod_purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select * from 50_belian_emas_agen where status='" & 1 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = 1
    Set Report58.DataSource = rs
    Report58.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing

'#### Jumlah Berat Keseluruhan #### - Start
Set rs1 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs1.Open "select SUM(Berat_Asal) from 50_belian_emas_agen where status='" & 1 & "' AND kod_purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs1.Open "select SUM(Berat_Asal) from 50_belian_emas_agen where status='" & 1 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs1(0)) Then
    Report58.Sections("Section5").Controls("L1").Caption = Format(rs1(0), "#,##0.00 g")
Else
    Report58.Sections("Section5").Controls("L1").Caption = Format(0, "#,##0.00 g")
End If
'#### Jumlah Berat Keseluruhan #### - End

rs1.Close
Set rs1 = Nothing

If x = 0 Then
    MsgBox "Tiada rekod trade in oleh agen dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_summary_report_jualan()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim Frm85_LM_UNTUNG As Double
Dim TA As Date
Dim TM As Date

x = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0
Frm85_LM_UNTUNG = 0

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

'### Reset maklumat kedai ### - Start
Report46.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report46.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report46.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report46.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report46.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
    
    LM_NAMA_HEADER = "HQ"
    
Else
    
    LM_NAMA_HEADER = MDI_frm1.L20_Text
    
End If

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & LM_NAMA_HEADER & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report46.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report46.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report46.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report46.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report46.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If
If Frm101.L6_Text = "Semua Kategori Produk" Then
    Frm85_LM_SEARCH_2 = Null
    Frm85_LM_SEARCH_2_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_2 = Frm101.L6_Text
    Frm85_LM_SEARCH_2_LOGIC = "="
End If
If Frm101.L34_Text = "Semua Dulang" Then
    Frm85_LM_SEARCH_3 = Null
    Frm85_LM_SEARCH_3_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_3 = Frm101.L34_Text
    Frm85_LM_SEARCH_3_LOGIC = "="
End If
If Frm101.L44_Text = 2 Then '0 : Jualan biasa , 1 : Jualan kepada agen , 2 : Semua jenis jualan
    Frm85_LM_SEARCH_4 = Null
    Frm85_LM_SEARCH_4_LOGIC = "<>"
ElseIf Frm101.L44_Text = 0 Then '0 : Jualan biasa , 1 : Jualan kepada agen , 2 : Semua jenis jualan
    Frm85_LM_SEARCH_4 = 0
    Frm85_LM_SEARCH_4_LOGIC = "="
ElseIf Frm101.L44_Text = 1 Then '0 : Jualan biasa , 1 : Jualan kepada agen , 2 : Semua jenis jualan
    Frm85_LM_SEARCH_4 = 1
    Frm85_LM_SEARCH_4_LOGIC = "="
End If
If Frm101.L45_Text = "Kedai & Online" Then
    Frm85_LM_SEARCH_5 = Null
    Frm85_LM_SEARCH_5_LOGIC = "<>"
ElseIf Frm101.L45_Text = "Kedai Sahaja" Then
    Frm85_LM_SEARCH_5 = 0
    Frm85_LM_SEARCH_5_LOGIC = "="
ElseIf Frm101.L45_Text = "Online Sahaja" Then
    Frm85_LM_SEARCH_5 = 1
    Frm85_LM_SEARCH_5_LOGIC = "="
End If

user_level = MDI_frm1.L4_Text

LM_INVOICE_RASMI = 0

If user_level = "Guest/User" Then
    Frm85_LM_SEARCH_6 = 1
    Frm85_LM_SEARCH_6_LOGIC = "="
    
    Report46.Sections("Section1").Controls("Text9").DataField = "no_invoice_r"
    
    Frm85_LM_SEARCH_7 = 1
    Frm85_LM_SEARCH_7_LOGIC = "="
    
    LM_INVOICE_RASMI = 1
Else
    Frm85_LM_SEARCH_6 = 0
    Frm85_LM_SEARCH_6_LOGIC = "="
    
    Frm85_LM_SEARCH_7 = 1
    Frm85_LM_SEARCH_7_LOGIC = "="
End If

If user_level = "Administration" Then
    Frm85_LM_SEARCH_10 = 1
    Frm85_LM_SEARCH_10_LOGIC = "="
    
    Frm85_LM_SEARCH_11 = 1
    Frm85_LM_SEARCH_11_LOGIC = "="
Else
    Frm85_LM_SEARCH_10 = 0
    Frm85_LM_SEARCH_10_LOGIC = "="
    
    Frm85_LM_SEARCH_11 = 1
    Frm85_LM_SEARCH_11_LOGIC = "="
End If

If Frm101.L46_Text = "Semua cawangan" Then

    Frm85_SEARCH_8 = Null
    Frm85_SEARCH_8_LOGIC = "<>"
    Frm85_SEARCH_9 = Null
    Frm85_SEARCH_9_LOGIC = "<>"
    
Else

    Frm85_SEARCH_8 = Frm101.L46_Text
    Frm85_SEARCH_8_LOGIC = "="
    Frm85_SEARCH_9 = "HQ"
    Frm85_SEARCH_9_LOGIC = "="
    
End If

If G_JENIS_JUALAN = "Barang Baru Sahaja" Then

    Frm85_LM_SEARCH_12 = 0
    Frm85_LM_SEARCH_12_LOGIC = "="
    
    Frm85_LM_SEARCH_13 = 0
    Frm85_LM_SEARCH_13_LOGIC = "="
    
ElseIf G_JENIS_JUALAN = "Barang Trade In Sahaja" Then

    Frm85_LM_SEARCH_12 = 1
    Frm85_LM_SEARCH_12_LOGIC = "="
    
    Frm85_LM_SEARCH_13 = 1
    Frm85_LM_SEARCH_13_LOGIC = "="
    
ElseIf G_JENIS_JUALAN = "Barang Baru Dan Barang Trade In" Then

    Frm85_LM_SEARCH_12 = 0
    Frm85_LM_SEARCH_12_LOGIC = "="
    
    Frm85_LM_SEARCH_13 = 1
    Frm85_LM_SEARCH_13_LOGIC = "="
    
End If
If Frm101.L47_Text = "Semua" Then
    Frm85_LM_SEARCH_20 = Null
    Frm85_LM_SEARCH_20_LOGIC = "<>"
Else
    If InStr(1, Frm101.L47_Text, "  |  ") <> 0 Then
        Frm84_LM_EMP_NO = Split(Frm101.L47_Text, "  |  ")(1)
        Frm84_LM_EMP_NAMA = Split(Frm101.L47_Text, "  |  ")(0)
    End If
    Frm85_LM_SEARCH_20 = Frm84_LM_EMP_NAMA
    Frm85_LM_SEARCH_20_LOGIC = "="
End If

'#### Header Report ###
If Frm101.L9_Text = 0 Then Report46.Sections("Section4").Controls("L1").Caption = "Report Jualan Bagi " & G_JENIS_JUALAN & " , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Jenis Jualan [" & Frm101.CBB5 & "] , Jualan Secara [" & Frm101.L45_Text & "] , Cawangan [" & Frm101.L46_Text & "]." 'Report Header"
If Frm101.L9_Text = 1 Then Report46.Sections("Section4").Controls("L1").Caption = "Report Jualan Bagi " & G_JENIS_JUALAN & " , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Jenis Jualan [" & Frm101.CBB5 & "] , Jualan Secara [" & Frm101.L45_Text & "] , Cawangan [" & Frm101.L46_Text & "] Dari " & TM & " Hingga " & TA 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select * from 23_senarai_jualan where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND(bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (baru_or_ti " & Frm85_LM_SEARCH_12_LOGIC & "'" & Frm85_LM_SEARCH_12 & "' OR baru_or_ti " & Frm85_LM_SEARCH_13_LOGIC & "'" & Frm85_LM_SEARCH_13 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND " _
& "kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND jenis_jualan " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND jualan_online " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' AND status_rekod = 1 order by no_resit ASC , tarikh ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select * from 23_senarai_jualan where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND(bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (baru_or_ti " & Frm85_LM_SEARCH_12_LOGIC & "'" & Frm85_LM_SEARCH_12 & "' OR baru_or_ti " & Frm85_LM_SEARCH_13_LOGIC & "'" & Frm85_LM_SEARCH_13 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND " _
& "kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND jenis_jualan " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND jualan_online " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' AND status_rekod = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by no_resit ASC , tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Not IsNull(rs!berat_jualan) Then 'Berat Jualan (g)
        If IsNumeric(rs!berat_jualan) Then Frm85_LM_BERAT = Frm85_LM_BERAT + rs!berat_jualan 'Total Berat Jualan (g)
    End If
    If Not IsNull(rs!harga_jualan_dengan_gst) Then
        If IsNumeric(rs!harga_jualan_dengan_gst) Then Frm85_LM_HARGA = Frm85_LM_HARGA + rs!harga_jualan_dengan_gst 'Total Harga Jualan (RM)
    End If
    
    Set Report46.DataSource = rs
    Report46.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing

Report46.Sections("Section5").Controls("L3").Caption = x 'Qty Barang
Report46.Sections("Section5").Controls("L4").Caption = Format(Frm85_LM_BERAT, "#,##0.00 g") 'Jumlah Berat (g)
Report46.Sections("Section5").Controls("L5").Caption = Format(Frm85_LM_HARGA, "#,##0.00") 'Jumlah Harga RM

If x = 0 Then
    MsgBox "Tiada Rekod Jualan Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_Header_Report_Ansuran()
'on error resume next
'#### Header Report Jualan Secara Ansuran #### - Start
Frm85.MSFlexGrid6.Clear
Frm85.MSFlexGrid6.Rows = 1
Frm85.MSFlexGrid6.RowHeight(0) = 1500
Frm85.MSFlexGrid6.FormatString = "<No.|<No.|<No. ID|<Tarikh|<Jenis Ansuran|<No. Siri Produk|<Kategori Produk|<Purity|<Berat Asal (g)|<Berat Jualan (g)|<Upah (RM)|<Jumlah Harga (RM)|<Dulang"

Frm85.MSFlexGrid6.ColWidth(0) = 0 'No.
Frm85.MSFlexGrid6.ColWidth(1) = 700 'No.
Frm85.MSFlexGrid6.ColAlignment(1) = 4

Frm85.MSFlexGrid6.ColWidth(2) = 0 'No. ID
Frm85.MSFlexGrid6.ColWidth(3) = 1500 'Tarikh
Frm85.MSFlexGrid6.ColAlignment(3) = 4

Frm85.MSFlexGrid6.ColWidth(4) = 1500 'Jenis Ansuran
Frm85.MSFlexGrid6.ColAlignment(4) = 4

Frm85.MSFlexGrid6.ColWidth(5) = 1500 'No. Siri Produk
Frm85.MSFlexGrid6.ColAlignment(5) = 4

Frm85.MSFlexGrid6.ColWidth(6) = 4000 'Kategori Produk

Frm85.MSFlexGrid6.ColWidth(7) = 1200 'Purity
Frm85.MSFlexGrid6.ColAlignment(7) = 4

Frm85.MSFlexGrid6.ColWidth(8) = 1200 'Berat Asal (g)
Frm85.MSFlexGrid6.ColAlignment(8) = 7

Frm85.MSFlexGrid6.ColWidth(9) = 1200 'Berat Jualan (g)
Frm85.MSFlexGrid6.ColAlignment(9) = 7

Frm85.MSFlexGrid6.ColWidth(10) = 1200 'Upah (RM)
Frm85.MSFlexGrid6.ColAlignment(10) = 7

Frm85.MSFlexGrid6.ColWidth(11) = 1500 'Jumlah Harga (RM)
Frm85.MSFlexGrid6.ColAlignment(11) = 7

Frm85.MSFlexGrid6.ColWidth(12) = 1000 'Dulang
Frm85.MSFlexGrid6.ColAlignment(12) = 4
End Sub
Sub Frm85_report_ansuran_page()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date
Dim Frm85_LM_TOTAL_PAGE As Double

Frm85_LM_TOTAL_PAGE = 0
Frm85_PAGE_SIZE = 34
x = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If
If Frm101.L6_Text = "Semua Kategori Produk" Then
    Frm85_LM_SEARCH_2 = Null
    Frm85_LM_SEARCH_2_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_2 = Frm101.L6_Text
    Frm85_LM_SEARCH_2_LOGIC = "="
End If
If Frm101.L34_Text = "Semua Dulang" Then
    Frm85_LM_SEARCH_3 = Null
    Frm85_LM_SEARCH_3_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_3 = Frm101.L34_Text
    Frm85_LM_SEARCH_3_LOGIC = "="
End If

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
        End If
    End If
End If

Frm85_LM_PAGE_FOUND = 0

'#### Header Report ###
If Frm101.L9_Text = 0 Then Frm85.L35_Text = "Report Jualan Secara Ansuran Bagi Purity [" & Frm101.L5_Text & "] & Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "]" 'Report Header"
If Frm101.L9_Text = 1 Then Frm85.L35_Text = "Report Jualan Secara Ansuran Bagi Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] Dari " & TM & " Hingga " & TA 'Report Header"

'#### Report Jualan Dari Ansuran #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select * from 27_senarai_ansuran where status='" & "Jelas" & "' AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' order by tarikh_jelas ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select * from 27_senarai_ansuran where status='" & "Jelas" & "' AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh_jelas BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh_jelas ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm85_LM_PAGE_FOUND = 0 Then
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm85.L79_Text = Frm85.L79_Text + 1
                Frm85_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm85.L79_Text) Then
                    If Frm85.L79_Text <> 1 Then
                        Frm85.L79_Text = Frm85.L79_Text - 1
                        Frm85_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x
    Frm85.MSFlexGrid6.Rows = x + 1
    Frm85.MSFlexGrid6.TextMatrix(x, 0) = x 'No.
    Frm85.MSFlexGrid6.TextMatrix(x, 1) = Y 'No.
    If Not IsNull(rs!ID) Then Frm85.MSFlexGrid6.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh_jelas) Then Frm85.MSFlexGrid6.TextMatrix(x, 3) = rs!tarikh_jelas 'Tarikh
    If Not IsNull(rs!jenis_ansuran) Then 'Jenis Ansuran
        If rs!jenis_ansuran = 0 Then
            Frm85.MSFlexGrid6.TextMatrix(x, 4) = "Harga Semasa"
        ElseIf rs!jenis_ansuran = 1 Then
            Frm85.MSFlexGrid6.TextMatrix(x, 4) = "Harga Tetap"
        End If
    End If
    If Not IsNull(rs!no_siri_Produk) Then Frm85.MSFlexGrid6.TextMatrix(x, 5) = rs!no_siri_Produk 'No. Siri Produk
    If Not IsNull(rs!kategori_Produk) Then Frm85.MSFlexGrid6.TextMatrix(x, 6) = rs!kategori_Produk 'Kategori Produk
    If Not IsNull(rs!purity) Then Frm85.MSFlexGrid6.TextMatrix(x, 7) = rs!purity 'Purity
    If Not IsNull(rs!Berat_Asal) Then Frm85.MSFlexGrid6.TextMatrix(x, 8) = Format(rs!Berat_Asal, "#,##0.00") 'Berat Asal (g)
    If Not IsNull(rs!berat_jualan) Then
        Frm85.MSFlexGrid6.TextMatrix(x, 9) = Format(rs!berat_jualan, "#,##0.00") 'Berat Jualan (g)
        If IsNumeric(rs!berat_jualan) Then Frm85_LM_BERAT = Frm85_LM_BERAT + rs!berat_jualan
    End If
    If Not IsNull(rs!UPAH) Then Frm85.MSFlexGrid6.TextMatrix(x, 10) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
    If Not IsNull(rs!jumlah_bayaran) Then
        Frm85.MSFlexGrid6.TextMatrix(x, 11) = Format(rs!jumlah_bayaran, "#,##0.00") 'Harga Asal Jualan (RM)
        If IsNumeric(rs!jumlah_bayaran) Then Frm85_LM_HARGA = Frm85_LM_HARGA + rs!jumlah_bayaran
    End If
    If Not IsNull(rs!dulang) Then Frm85.MSFlexGrid6.TextMatrix(x, 12) = rs!dulang 'Dulang
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'#### Report Jualan Dari Ansuran #### - End

Frm85.L36_Text = x 'Total Barang
Frm85.L37_Text = Format(Frm85_LM_BERAT, "#,##0.00 g")  'Total Berat
Frm85.L38_Text = "RM " & Format(Frm85_LM_HARGA, "#,##0.00")  'Total Harga Jualan

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select COUNT(ID) from 27_senarai_ansuran where status='" & "Jelas" & "' AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select COUNT(ID) from 27_senarai_ansuran where status='" & "Jelas" & "' AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh_jelas BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

'#### Jumlah Bilangan Barang Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select COUNT(no_siri_Produk) from 27_senarai_ansuran where status='" & "Jelas" & "' AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select COUNT(no_siri_Produk) from 27_senarai_ansuran where status='" & "Jelas" & "' AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh_jelas BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L59_Text = rs(0)
Else
    Frm85.L59_Text = 0
End If
'#### Jumlah Bilangan Barang Keseluruhan #### - End

'#### Jumlah Berat Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select SUM(Berat_Jualan) from 27_senarai_ansuran where status='" & "Jelas" & "' AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select SUM(Berat_Jualan) from 27_senarai_ansuran where status='" & "Jelas" & "' AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh_jelas BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L60_Text = Format(rs(0), "#,##0.00 g")
Else
    Frm85.L60_Text = "0.00 g"
End If
'#### Jumlah Berat Keseluruhan #### - End

'#### Jumlah Harga Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select SUM(jumlah_bayaran) from 27_senarai_ansuran where status='" & "Jelas" & "' AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select SUM(jumlah_bayaran) from 27_senarai_ansuran where status='" & "Jelas" & "' AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh_jelas BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L61_Text = "RM " & Format(rs(0), "#,##0.00")
Else
    Frm85.L61_Text = "RM " & "0.00"
End If
'#### Jumlah Harga Keseluruhan #### - End

If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    
    Frm85.Pic9.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    
    MsgBox "Tiada Rekod Jualan Secara Ansuran Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_Header_Report_Tempahan()
'on error resume next
With Frm85.LV6
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear
    
    Frm85.LV6.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "Tarikh", 1500, 2
    .ColumnHeaders.Add 5, , "No. Siri Produk", 1700
    .ColumnHeaders.Add 6, , "Kategori Produk", 3200
    .ColumnHeaders.Add 7, , "Purity", 1500
    .ColumnHeaders.Add 8, , "Berat Asal (g)", 1500, 1
    .ColumnHeaders.Add 9, , "Berat Jualan (g)", 1700, 1
    .ColumnHeaders.Add 10, , "Upah (RM)", 1300, 1
    .ColumnHeaders.Add 11, , "Jumlah Harga (RM)", 1900, 1
    .ColumnHeaders.Add 12, , "Dulang", 1000, 2
    .ColumnHeaders.Add 13, , "Cawangan", 4000

End With
End Sub
Sub Frm85_report_tempahan_page()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date
Dim Frm85_LM_TOTAL_PAGE As Double

Frm85_LM_TOTAL_PAGE = 0
Frm85_PAGE_SIZE = 34
x = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If
If Frm101.L6_Text = "Semua Kategori Produk" Then
    Frm85_LM_SEARCH_2 = Null
    Frm85_LM_SEARCH_2_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_2 = Frm101.L6_Text
    Frm85_LM_SEARCH_2_LOGIC = "="
End If
If Frm101.L34_Text = "Semua Dulang" Then
    Frm85_LM_SEARCH_3 = Null
    Frm85_LM_SEARCH_3_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_3 = Frm101.L34_Text
    Frm85_LM_SEARCH_3_LOGIC = "="
End If
If Frm101.L46_Text = "Semua cawangan" Then

    Frm85_LM_SEARCH_5 = Null
    Frm85_LM_SEARCH_5_LOGIC = "<>"
    Frm85_LM_SEARCH_6 = Null
    Frm85_LM_SEARCH_6_LOGIC = "<>"
    
Else

    Frm85_LM_SEARCH_5 = Frm101.L46_Text
    Frm85_LM_SEARCH_5_LOGIC = "="
    Frm85_LM_SEARCH_6 = "HQ"
    Frm85_LM_SEARCH_6_LOGIC = "="
    
End If

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
        End If
    End If
End If

Frm85_LM_PAGE_FOUND = 0

'#### Header Report ###
If Frm101.L9_Text = 0 Then Frm85.L39_Text = "Report Jualan Secara Tempahan Bagi Purity [" & Frm101.L5_Text & "] & Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "]." 'Report Header"
If Frm101.L9_Text = 1 Then Frm85.L39_Text = "Report Jualan Secara Tempahan Bagi Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "] Dari " & TM & " Hingga " & TA & "." 'Report Header"

'#### Report Jualan Dari Tempahan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select * from 42_tempahan_siap where status_invoice = 1 AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select * from 42_tempahan_siap where status_invoice = 1 AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm85_LM_PAGE_FOUND = 0 Then
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm85.L79_Text = Frm85.L79_Text + 1
                Frm85_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm85.L79_Text) Then
                    If Frm85.L79_Text <> 1 Then
                        Frm85.L79_Text = Frm85.L79_Text - 1
                        Frm85_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x

    With Frm85.LV6.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!tarikh) Then 'Tarikh
            .ListSubItems.Add , , rs!tarikh
        Else
            .ListSubItems.Add , , ""
        End If
        
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
            If IsNumeric(rs!berat_jualan) Then Frm85_LM_BERAT = Frm85_LM_BERAT + rs!berat_jualan
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!UPAH) Then  'Upah (RM)
            .ListSubItems.Add , , Format(rs!UPAH, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_dengan_gst) Then 'Harga Asal Jualan (RM)
            .ListSubItems.Add , , Format(rs!harga_dengan_gst, "#,##0.00")
            If IsNumeric(rs!harga_dengan_gst) Then Frm85_LM_HARGA = Frm85_LM_HARGA + rs!harga_dengan_gst
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!dulang) Then 'Dulang
            .ListSubItems.Add , , rs!dulang
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!cawangan) Then 'Cawangan
            .ListSubItems.Add , , rs!cawangan
        Else
            .ListSubItems.Add , , ""
        End If
        
    End With
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'#### Report Jualan Dari Tempahan #### - End

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select COUNT(ID) from 42_tempahan_siap where status_invoice = 1 AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select COUNT(ID) from 42_tempahan_siap where status_invoice = 1 AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

'#### Jumlah Bilangan Barang Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select COUNT(no_siri_Produk) from 42_tempahan_siap where status_invoice = 1 AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select COUNT(no_siri_Produk) from 42_tempahan_siap where status_invoice = 1 AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L62_Text = rs(0)
Else
    Frm85.L62_Text = 0
End If
'#### Jumlah Bilangan Barang Keseluruhan #### - End

'#### Jumlah Berat Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select SUM(Berat_Jualan) from 42_tempahan_siap where status_invoice = 1 AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select SUM(Berat_Jualan) from 42_tempahan_siap where status_invoice = 1 AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L63_Text = Format(rs(0), "#,##0.00 g")
Else
    Frm85.L63_Text = "0.00 g"
End If
'#### Jumlah Berat Keseluruhan #### - End

'#### Jumlah Harga Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select SUM(harga_dengan_gst) from 42_tempahan_siap where status_invoice = 1 AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select SUM(harga_dengan_gst) from 42_tempahan_siap where status_invoice = 1 AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L64_Text = "RM " & Format(rs(0), "#,##0.00")
Else
    Frm85.L64_Text = "RM " & "0.00"
End If
'#### Jumlah Harga Keseluruhan #### - End

If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    
    Frm85.Pic10.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    
    MsgBox "Tiada Rekod Jualan Secara Tempahan Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_summary_report_tempahan()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

x = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

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

'### Reset maklumat kedai ### - Start
Report51.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report51.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report51.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report51.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report51.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
    
    LM_NAMA_HEADER = "HQ"
    
Else
    
    LM_NAMA_HEADER = MDI_frm1.L20_Text
    
End If

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & LM_NAMA_HEADER & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report51.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report51.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report51.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report51.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report51.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If
If Frm101.L6_Text = "Semua Kategori Produk" Then
    Frm85_LM_SEARCH_2 = Null
    Frm85_LM_SEARCH_2_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_2 = Frm101.L6_Text
    Frm85_LM_SEARCH_2_LOGIC = "="
End If
If Frm101.L34_Text = "Semua Dulang" Then
    Frm85_LM_SEARCH_3 = Null
    Frm85_LM_SEARCH_3_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_3 = Frm101.L34_Text
    Frm85_LM_SEARCH_3_LOGIC = "="
End If
If Frm101.L46_Text = "Semua cawangan" Then

    Frm85_LM_SEARCH_5 = Null
    Frm85_LM_SEARCH_5_LOGIC = "<>"
    Frm85_LM_SEARCH_6 = Null
    Frm85_LM_SEARCH_6_LOGIC = "<>"
    
Else

    Frm85_LM_SEARCH_5 = Frm101.L46_Text
    Frm85_LM_SEARCH_5_LOGIC = "="
    Frm85_LM_SEARCH_6 = "HQ"
    Frm85_LM_SEARCH_6_LOGIC = "="
    
End If

' #### No ID GST #### - Start
'Set rs = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

'If Not rs.EOF Then
'    If rs!Default1 = "Default" Then
'        If Not IsNull(rs!id_gst) Then
'            Report51.Sections("Section2").Controls("L100").Caption = "GST ID : " & rs!id_gst 'No. ID GST
'        Else
'            Report51.Sections("Section2").Controls("L100").Caption = vbNullString 'No. ID GST
'        End If
'    End If
'End If

'rs.Close
'Set rs = Nothing
' #### No ID GST #### - End

'#### Header Report ###
If Frm101.L9_Text = 0 Then Report51.Sections("Section4").Controls("L1").Caption = "Report Jualan Tempahan Bagi Purity [" & Frm101.L5_Text & "] & Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "]." 'Report Header"
If Frm101.L9_Text = 1 Then Report51.Sections("Section4").Controls("L1").Caption = "Report Jualan Tempahan Bagi Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "] Dari " & TM & " Hingga " & TA & "." 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select * from 42_tempahan_siap where status_invoice = 1 AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select * from 42_tempahan_siap where status_invoice = 1 AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Not IsNull(rs!berat_jualan) Then 'Berat Jualan (g)
        If IsNumeric(rs!berat_jualan) Then Frm85_LM_BERAT = Frm85_LM_BERAT + rs!berat_jualan 'Total Berat Jualan (g)
    End If
    If Not IsNull(rs!harga_dengan_gst) Then
        If IsNumeric(rs!harga_dengan_gst) Then Frm85_LM_HARGA = Frm85_LM_HARGA + rs!harga_dengan_gst 'Total Harga Jualan (RM)
    End If
    
    Set Report51.DataSource = rs
    Report51.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing

Report51.Sections("Section5").Controls("L3").Caption = x 'Qty Barang
Report51.Sections("Section5").Controls("L4").Caption = Format(Frm85_LM_BERAT, "#,##0.00 g") 'Jumlah Berat (g)
Report51.Sections("Section5").Controls("L5").Caption = "RM " & Format(Frm85_LM_HARGA, "#,##0.00") 'Jumlah Harga RM

If x = 0 Then
    MsgBox "Tiada Rekod Jualan Secara Tempahan Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_summary_report_ansuran()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

x = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

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

'### Reset maklumat kedai ### - Start
Report52.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report52.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report52.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report52.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report52.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report52.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report52.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report52.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report52.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report52.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If
If Frm101.L6_Text = "Semua Kategori Produk" Then
    Frm85_LM_SEARCH_2 = Null
    Frm85_LM_SEARCH_2_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_2 = Frm101.L6_Text
    Frm85_LM_SEARCH_2_LOGIC = "="
End If
If Frm101.L34_Text = "Semua Dulang" Then
    Frm85_LM_SEARCH_3 = Null
    Frm85_LM_SEARCH_3_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_3 = Frm101.L34_Text
    Frm85_LM_SEARCH_3_LOGIC = "="
End If

'#### Header Report ###
If Frm101.L9_Text = 0 Then Report52.Sections("Section4").Controls("L1").Caption = "Report Jualan Ansuran Bagi Purity [" & Frm101.L5_Text & "] & Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "]" 'Report Header"
If Frm101.L9_Text = 1 Then Report52.Sections("Section4").Controls("L1").Caption = "Report Jualan Ansuran Bagi Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] Dari " & TM & " Hingga " & TA 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select * from 27_senarai_ansuran where status='" & "Jelas" & "' AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select * from 27_senarai_ansuran where status='" & "Jelas" & "' AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh_jelas BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Not IsNull(rs!berat_jualan) Then 'Berat Jualan (g)
        If IsNumeric(rs!berat_jualan) Then Frm85_LM_BERAT = Frm85_LM_BERAT + rs!berat_jualan 'Total Berat Jualan (g)
    End If
    If Not IsNull(rs!jumlah_bayaran) Then
        If IsNumeric(rs!jumlah_bayaran) Then Frm85_LM_HARGA = Frm85_LM_HARGA + rs!jumlah_bayaran 'Total Harga Jualan (RM)
    End If
    
    Set Report52.DataSource = rs
    Report52.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing

Report52.Sections("Section5").Controls("L3").Caption = x 'Qty Barang
Report52.Sections("Section5").Controls("L4").Caption = Format(Frm85_LM_BERAT, "#,##0.00 g") 'Jumlah Berat (g)
Report52.Sections("Section5").Controls("L5").Caption = "RM " & Format(Frm85_LM_HARGA, "#,##0.00") 'Jumlah Harga RM

If x = 0 Then
    MsgBox "Tiada Rekod Jualan Secara Ansuran Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm101_initial_setting()
'on error resume next
Frm101.Pic1.Left = 120
Frm101.Pic1.Top = 480
Frm101.Pic7.Left = 9720
Frm101.Pic7.Top = 1200
Frm101.Pic8.Left = 9720
Frm101.Pic8.Top = 1440
Frm101.Pic9.Left = 9720
Frm101.Pic9.Top = 1680

Frm101.L31_Text.Left = 9720
Frm101.L31_Text.Top = 960
Frm101.L32_Text.Left = 9720
Frm101.L32_Text.Top = 1200

Frm101.Pic1.Visible = False
Frm101.Pic7.Visible = False
Frm101.Pic8.Visible = False

Frm101.L44_Text = 0
End Sub
Sub Frm85_report_belian_barcode()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim Frm85_LM_TOTAL_PAGE As Double

x = 0
Frm85_PAGE_SIZE = 34
Frm85_LM_TOTAL_PAGE = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
    End If
End If

If Frm101.L46_Text = "Semua cawangan" Then

    Frm85_LM_SEARCH_5 = Null
    Frm85_LM_SEARCH_5_LOGIC = "<>"
    Frm85_LM_SEARCH_6 = Null
    Frm85_LM_SEARCH_6_LOGIC = "<>"
    
Else

    Frm85_LM_SEARCH_5 = Frm101.L46_Text
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

Frm85_LM_PAGE_FOUND = 0

Frm85.L10_Text = "Report belian bagi item dengan no. siri produk [" & Frm101.L5_Text & "]" 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' AND StatusItem <> 0 AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND (receiving_Status='" & 1 & "' OR receiving_Status='" & 0 & "') order by tarikh_belian ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1

    If Frm85_LM_PAGE_FOUND = 0 Then
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm85.L79_Text = Frm85.L79_Text + 1
                Frm85_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm85.L79_Text) Then
                    If Frm85.L79_Text <> 1 Then
                        Frm85.L79_Text = Frm85.L79_Text - 1
                        Frm85_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x
        
    With Frm85.LV2.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!tarikh_belian) Then 'Tarikh Belian
            .ListSubItems.Add , , rs!tarikh_belian
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_siri_Produk) Then 'No. Siri Produk
            .ListSubItems.Add , , rs!no_siri_Produk
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kod_Purity) Then 'Purity
            .ListSubItems.Add , , rs!kod_Purity
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kategori_Produk) Then 'Kategori Produk
            .ListSubItems.Add , , rs!kategori_Produk
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!nama_Supplier) Then 'Nama Supplier
            .ListSubItems.Add , , rs!nama_Supplier
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Berat) Then 'Berat (g)
            .ListSubItems.Add , , Format(rs!Berat, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kos_Belian_Gram) Then 'Rate Penerimaan (RM/g)
            .ListSubItems.Add , , Format(rs!kos_Belian_Gram, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!UPAH) Then 'Upah (RM)
            .ListSubItems.Add , , Format(rs!UPAH, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!SpreadValue) Then 'Spread (%)
            .ListSubItems.Add , , Format(rs!SpreadValue, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_lepas_spread) Then 'Harga Selepas Spread (RM)
            .ListSubItems.Add , , Format(rs!harga_lepas_spread, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!adjustment) Then 'Adjustment (RM)
            .ListSubItems.Add , , Format(rs!adjustment, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_item) Then 'Harga Belian (RM) : Tidak Campur Cukai GST
            .ListSubItems.Add , , Format(rs!harga_item, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Upah_Jualan) Then 'Upah Jualan (RM) : Pelanggan
            .ListSubItems.Add , , Format(rs!Upah_Jualan, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!Upah_Member) Then 'Upah Jualan (RM) : Ahli Biasa
            .ListSubItems.Add , , Format(rs!Upah_Member, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!Upah_Pengedar) Then 'Upah Jualan (RM) : Silver
            .ListSubItems.Add , , Format(rs!Upah_Pengedar, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Upah_RAF) Then 'Upah Jualan (RM) : Gold
            .ListSubItems.Add , , Format(rs!Upah_RAF, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!upah_normal_dealer) Then 'Upah Jualan (RM) : Platinum
            .ListSubItems.Add , , Format(rs!upah_normal_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!upah_master_dealer) Then 'Upah Jualan (RM) : Master Dealer
            .ListSubItems.Add , , Format(rs!upah_master_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!code_Supplier) Then 'Tetapan Harga Jualan (RM) : Pelanggan
            .ListSubItems.Add , , Format(rs!code_Supplier, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!HargaJualan_Member) Then 'Tetapan Harga Jualan (RM) : Pelanggan
            .ListSubItems.Add , , Format(rs!HargaJualan_Member, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!HargaJualan_Pengedar) Then 'Tetapan Harga Jualan (RM) : Pengedar
            .ListSubItems.Add , , Format(rs!HargaJualan_Pengedar, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!HargaJualan_RAF) Then 'Tetapan Harga Jualan (RM) : RAF
            .ListSubItems.Add , , Format(rs!HargaJualan_RAF, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!hargajualan_normal_dealer) Then 'Tetapan Harga Jualan (RM) : Normal Dealer
            .ListSubItems.Add , , Format(rs!hargajualan_normal_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!hargajualan_master_dealer) Then 'Tetapan Harga Jualan (RM) : Master Dealer
            .ListSubItems.Add , , Format(rs!hargajualan_master_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
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
        
        If Not IsNull(rs!dimension_Saiz) Then 'Tebal
            .ListSubItems.Add , , rs!dimension_Saiz
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!bill_No_Belian) Then 'No. Invoice Supplier
            .ListSubItems.Add , , rs!bill_No_Belian
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
        
        If Not IsNull(rs!cawangan) Then 'Cawangan
            .ListSubItems.Add , , rs!cawangan
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
rs.Open "select COUNT(ID) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' AND StatusItem <> 0 AND (receiving_Status='" & 1 & "' OR receiving_Status='" & 0 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

'#### Jumlah Bilangan Barang Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' AND StatusItem <> 0 AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND (receiving_Status='" & 1 & "' OR receiving_Status='" & 0 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L44_Text = rs(0)
Else
    Frm85.L44_Text = 0
End If

rs.Close
Set rs = Nothing
'#### Jumlah Bilangan Barang Keseluruhan #### - End

'#### Jumlah Berat Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(Berat) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' AND StatusItem <> 0 AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND (receiving_Status='" & 1 & "' OR receiving_Status='" & 0 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L45_Text = Format(rs(0), "#,##0.00 g")
Else
    Frm85.L45_Text = "0.00 g"
End If

rs.Close
Set rs = Nothing
'#### Jumlah Berat Keseluruhan #### - End

'#### Jumlah Modal Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_item) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND (receiving_Status='" & 1 & "' OR receiving_Status='" & 0 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L46_Text = "RM " & Format(rs(0), "#,##0.00")
Else
    Frm85.L46_Text = "RM 0.00"
End If

rs.Close
Set rs = Nothing
'#### Jumlah Modal Keseluruhan #### - End

If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    Frm85.Pic2.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    MsgBox "Tiada Rekod Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_Report_Jualan_barcode()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim Frm85_LM_UNTUNG As Double
Dim TA As Date
Dim TM As Date
Dim Frm85_LM_TOTAL_PAGE As Double
Dim Frm85_LM_UNTUNG2 As Double

Frm85_LM_UNTUNG2 = 0
Frm85_PAGE_SIZE = 34
Frm85_LM_TOTAL_PAGE = 0
x = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0
Frm85_LM_UNTUNG = 0

user_level = MDI_frm1.L4_Text

LM_INVOICE_RASMI = 0

If user_level = "Guest/User" Then
    Frm85_LM_SEARCH_6 = 1
    Frm85_LM_SEARCH_6_LOGIC = "="
    LM_INVOICE_RASMI = 1
    Frm85_LM_SEARCH_7 = 1
    Frm85_LM_SEARCH_7_LOGIC = "="
Else
    Frm85_LM_SEARCH_6 = 0
    Frm85_LM_SEARCH_6_LOGIC = "="
    
    Frm85_LM_SEARCH_7 = 1
    Frm85_LM_SEARCH_7_LOGIC = "="
End If

If user_level = "Administration" Then
    Frm85_LM_SEARCH_10 = 1
    Frm85_LM_SEARCH_10_LOGIC = "="
    
    Frm85_LM_SEARCH_11 = 1
    Frm85_LM_SEARCH_11_LOGIC = "="
Else
    Frm85_LM_SEARCH_10 = 0
    Frm85_LM_SEARCH_10_LOGIC = "="
    
    Frm85_LM_SEARCH_11 = 1
    Frm85_LM_SEARCH_11_LOGIC = "="
End If

If Frm101.L46_Text = "Semua cawangan" Then

    Frm85_SEARCH_8 = Null
    Frm85_SEARCH_8_LOGIC = "<>"
    Frm85_SEARCH_9 = Null
    Frm85_SEARCH_9_LOGIC = "<>"
    
Else

    Frm85_SEARCH_8 = Frm101.L46_Text
    Frm85_SEARCH_8_LOGIC = "="
    Frm85_SEARCH_9 = "HQ"
    Frm85_SEARCH_9_LOGIC = "="
    
End If

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
        End If
    End If
End If

Frm85_LM_PAGE_FOUND = 0

'#### Header Report ###
Frm85.L14_Text = "Report jualan bagi item dengan no. siri produk [" & Frm101.L5_Text & "]." 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 23_senarai_jualan where (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND no_siri_produk='" & Frm101.L5_Text & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm85_LM_PAGE_FOUND = 0 Then
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm85.L79_Text = Frm85.L79_Text + 1
                Frm85_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm85.L79_Text) Then
                    If Frm85.L79_Text <> 1 Then
                        Frm85.L79_Text = Frm85.L79_Text - 1
                        Frm85_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x

    With Frm85.LV1.ListItems.Add(, , rs!ID)
     
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!tarikh) Then 'Tarikh Jualan
            .ListSubItems.Add , , rs!tarikh
        Else
            .ListSubItems.Add , , ""
        End If
        
        If LM_INVOICE_RASMI = 0 Then
        
            If Not IsNull(rs!no_resit) Then .ListSubItems.Add , , rs!no_resit
            If Not IsNull(rs!no_invoice_r) Then .ListSubItems.Add , , rs!no_invoice_r
            
        Else
            .ListSubItems.Add , , ""
        End If
        
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
        
        If Not IsNull(rs!berat_jualan) Then  'Berat Jualan (g)
            .ListSubItems.Add , , Format(rs!berat_jualan, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_Semasa) Then 'Harga Semasa (RM/g)
            .ListSubItems.Add , , Format(rs!harga_Semasa, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!UPAH) Then 'Upah (RM)
            .ListSubItems.Add , , Format(rs!UPAH, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_asal) Then 'Harga Asal (RM)
            .ListSubItems.Add , , Format(rs!harga_asal, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!diskaun) Then 'Diskaun (%)
            .ListSubItems.Add , , Format(rs!diskaun, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!harga_lepas_diskaun) Then 'Harga Selepas Diskaun (RM)
            .ListSubItems.Add , , Format(rs!harga_lepas_diskaun, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!adjustment) Then 'Adjustment (RM)
            .ListSubItems.Add , , Format(rs!adjustment, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_jualan_dengan_gst) Then 'Harga Jualan (RM)
            .ListSubItems.Add , , Format(rs!harga_jualan_dengan_gst, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!gst_ari_nashi) Then
        
            If rs!gst_ari_nashi = "ZR (L)" Then '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                .ListSubItems.Add , , "ZR (L)"
            ElseIf rs!gst_ari_nashi = "SR" Then
                .ListSubItems.Add , , "SR"
            End If
            
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!jumlah_gst) Then 'Jumlah GST (RM)
            .ListSubItems.Add , , Format(rs!jumlah_gst, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!cawangan) Then
            .ListSubItems.Add , , rs!cawangan
        Else
            .ListSubItems.Add , , ""
        End If
        If Not IsNull(rs!nama_pekerja) Then
            .ListSubItems.Add , , rs!nama_pekerja
        Else
            .ListSubItems.Add , , ""
        End If
        If Not IsNull(rs!dulang) Then
            .ListSubItems.Add , , rs!dulang
        Else
            .ListSubItems.Add , , ""
        End If
    
    End With
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'Frm85.L15_Text = x 'Total Barang
'Frm85.L16_Text = Format(Frm85_LM_BERAT, "#,##0.00 g")  'Total Berat
'Frm85.L17_Text = "RM " & Format(Frm85_LM_HARGA, "#,##0.00")  'Total Harga Jualan
'Frm85.L18_Text = "RM " & Format(Frm85_LM_UNTUNG, "#,##0.00")  'Total Keuntungan Jualan
'Frm85.L87_Text = "RM " & Format(Frm85_LM_UNTUNG2, "#,##0.00")  'Total Keuntungan Jualan 2

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 23_senarai_jualan where (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND no_siri_produk='" & Frm101.L5_Text & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

'#### Jumlah Bilangan Barang Terjual Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 23_senarai_jualan where (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L47_Text = rs(0)
Else
    Frm85.L47_Text = 0
End If

rs.Close
Set rs = Nothing
'#### Jumlah Bilangan Barang Terjual Keseluruhan #### - End

'#### Jumlah Berat Jualan Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(Berat_Jualan) from 23_senarai_jualan where (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L48_Text = Format(rs(0), "#,##0.00 g")
Else
    Frm85.L48_Text = "0.00 g"
End If

rs.Close
Set rs = Nothing
'#### Jumlah Berat Jualan Keseluruhan #### - End

'#### Jumlah Harga Jualan Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_dengan_gst) from 23_senarai_jualan where (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L49_Text = "RM " & Format(rs(0), "#,##0.00")
Else
    Frm85.L49_Text = "RM 0.00"
End If

rs.Close
Set rs = Nothing
'#### Jumlah Harga Jualan Keseluruhan #### - End

'#### Jumlah Keuntungan Keseluruhan 1 #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(untung) from 23_senarai_jualan where (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then
    Frm85.L50_Text = "RM " & Format(rs(0), "#,##0.00")
Else
    Frm85.L50_Text = "RM 0.00"
End If

rs.Close
Set rs = Nothing
'#### Jumlah Keuntungan Keseluruhan 1 #### - End

'#### Jumlah Keuntungan Keseluruhan 2 #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(untung2) from 23_senarai_jualan where (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then
    Frm85.L88_Text = "RM " & Format(rs(0), "#,##0.00")
Else
    Frm85.L88_Text = "RM 0.00"
End If

rs.Close
Set rs = Nothing
'#### Jumlah Keuntungan Keseluruhan 2 #### - End

If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    
    Frm85.Pic3.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else

    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    
    MsgBox "Tiada Rekod Jualan Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_report_buyback_barcode()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date
Dim Frm85_LM_TOTAL_PAGE As Double

Frm85_LM_TOTAL_PAGE = 0
Frm85_PAGE_SIZE = 34
x = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
        End If
    End If
End If

If Frm101.L46_Text = "Semua cawangan" Then

    Frm85_LM_SEARCH_5 = Null
    Frm85_LM_SEARCH_5_LOGIC = "<>"
    Frm85_LM_SEARCH_6 = Null
    Frm85_LM_SEARCH_6_LOGIC = "<>"
    
Else

    Frm85_LM_SEARCH_5 = Frm101.L46_Text
    Frm85_LM_SEARCH_5_LOGIC = "="
    Frm85_LM_SEARCH_6 = "HQ"
    Frm85_LM_SEARCH_6_LOGIC = "="
    
End If

Frm85_LM_PAGE_FOUND = 0

'#### Header Report ###
Frm85.L19_Text = "Report belian buyback / trade in bagi item dengan no. siri produk [" & Frm101.L5_Text & "]." 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "') order by tarikh_belian ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    'If rs!receiving_Status = "2" Or rs!receiving_Status = "3" Then
        x = x + 1
        If Frm85_LM_PAGE_FOUND = 0 Then
            If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
                If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                    Frm85.L79_Text = Frm85.L79_Text + 1
                    Frm85_LM_PAGE_FOUND = 1
                ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                    If IsNumeric(Frm85.L79_Text) Then
                        If Frm85.L79_Text <> 1 Then
                            Frm85.L79_Text = Frm85.L79_Text - 1
                            Frm85_LM_PAGE_FOUND = 1
                        End If
                    End If
                End If
            End If
        End If
        Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x

        With Frm85.LV3.ListItems.Add(, , rs!ID)
        
            .ListSubItems.Add , , Y
            
            If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
            
            If Not IsNull(rs!tarikh_belian) Then 'Tarikh Belian
                .ListSubItems.Add , , rs!tarikh_belian
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!no_siri_Produk) Then 'No. Siri Produk
                .ListSubItems.Add , , rs!no_siri_Produk
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!kod_Purity) Then 'Purity
                .ListSubItems.Add , , rs!kod_Purity
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!kategori_Produk) Then 'Kategori Produk
                .ListSubItems.Add , , rs!kategori_Produk
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!nama_Supplier) Then 'Nama Supplier
                .ListSubItems.Add , , rs!nama_Supplier
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!Berat) Then 'Berat (g)
                .ListSubItems.Add , , Format(rs!Berat, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!kos_Belian_Gram) Then 'Rate Penerimaan (RM/g)
                .ListSubItems.Add , , Format(rs!kos_Belian_Gram, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!UPAH) Then 'Upah (RM)
                .ListSubItems.Add , , Format(rs!UPAH, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!SpreadValue) Then 'Spread (%)
                .ListSubItems.Add , , Format(rs!SpreadValue, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!harga_lepas_spread) Then 'Harga Selepas Spread (RM)
                .ListSubItems.Add , , Format(rs!harga_lepas_spread, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If

            If Not IsNull(rs!adjustment) Then 'Adjustment (RM)
                .ListSubItems.Add , , Format(rs!adjustment, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!harga_item) Then 'Harga Belian (RM) : Tidak Campur Cukai GST
                .ListSubItems.Add , , Format(rs!harga_item, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!Upah_Jualan) Then 'Upah Jualan (RM) : Pelanggan
                .ListSubItems.Add , , Format(rs!Upah_Jualan, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If

            If Not IsNull(rs!Upah_Member) Then 'Upah Jualan (RM) : Ahli Biasa
                .ListSubItems.Add , , Format(rs!Upah_Member, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If

            If Not IsNull(rs!Upah_Pengedar) Then 'Upah Jualan (RM) : Silver
                .ListSubItems.Add , , Format(rs!Upah_Pengedar, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!Upah_RAF) Then 'Upah Jualan (RM) : Gold
                .ListSubItems.Add , , Format(rs!Upah_RAF, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!upah_normal_dealer) Then 'Upah Jualan (RM) : Platinum
                .ListSubItems.Add , , Format(rs!upah_normal_dealer, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!upah_master_dealer) Then 'Upah Jualan (RM) : Master Dealer
                .ListSubItems.Add , , Format(rs!upah_master_dealer, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!code_Supplier) Then 'Tetapan Harga Jualan (RM) : Pelanggan
                .ListSubItems.Add , , Format(rs!code_Supplier, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!HargaJualan_Member) Then 'Tetapan Harga Jualan (RM) : Pelanggan
                .ListSubItems.Add , , Format(rs!HargaJualan_Member, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!HargaJualan_Pengedar) Then 'Tetapan Harga Jualan (RM) : Pengedar
                .ListSubItems.Add , , Format(rs!HargaJualan_Pengedar, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!HargaJualan_RAF) Then 'Tetapan Harga Jualan (RM) : RAF
                .ListSubItems.Add , , Format(rs!HargaJualan_RAF, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!hargajualan_normal_dealer) Then 'Tetapan Harga Jualan (RM) : Normal Dealer
                .ListSubItems.Add , , Format(rs!hargajualan_normal_dealer, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If

            If Not IsNull(rs!hargajualan_master_dealer) Then 'Tetapan Harga Jualan (RM) : Master Dealer
                .ListSubItems.Add , , Format(rs!hargajualan_master_dealer, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
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
            
            If Not IsNull(rs!dimension_Saiz) Then 'Tebal
                .ListSubItems.Add , , rs!dimension_Saiz
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!bill_No_Belian) Then 'No invoice supplier
                .ListSubItems.Add , , rs!bill_No_Belian
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
            
            If Not IsNull(rs!cawangan) Then 'Cawangan
                .ListSubItems.Add , , rs!cawangan
            Else
                .ListSubItems.Add , , ""
            End If

            If Not IsNull(rs!bill_No_Trade_In) Then
                .ListSubItems.Add , , rs!bill_No_Trade_In
            Else
                .ListSubItems.Add , , ""
            End If
            
        End With

    'End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

'#### Jumlah Bilangan Barang Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L51_Text = rs(0)
Else
    Frm85.L51_Text = 0
End If
'#### Jumlah Bilangan Barang Keseluruhan #### - End

'#### Jumlah Berat Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(Berat) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L52_Text = Format(rs(0), "#,##0.00 g")
Else
    Frm85.L52_Text = "0.00 g"
End If
'#### Jumlah Berat Keseluruhan #### - End

'#### Jumlah Modal Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_item) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L53_Text = "RM " & Format(rs(0), "#,##0.00")
Else
    Frm85.L53_Text = "RM 0.00"
End If
'#### Jumlah Modal Keseluruhan #### - End

If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    
    Frm85.Pic4.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    
    MsgBox "Tiada Rekod Belian Buyback / Trade In Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_report_stok_barcode()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date
Dim Frm85_LM_TOTAL_PAGE As Double

Frm85_LM_TOTAL_PAGE = 0
Frm85_PAGE_SIZE = 34
x = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
        End If
    End If
End If

If Frm101.L46_Text = "Semua cawangan" Then

    Frm85_LM_SEARCH_5 = Null
    Frm85_LM_SEARCH_5_LOGIC = "<>"
    Frm85_LM_SEARCH_6 = Null
    Frm85_LM_SEARCH_6_LOGIC = "<>"
    
Else

    Frm85_LM_SEARCH_5 = Frm101.L46_Text
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

Frm85_LM_PAGE_FOUND = 0

'#### Header Report ###
Frm85.L23_Text = "Report stok bagi item dengan no. siri produk [" & Frm101.L5_Text & "]." 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from Data_Database where (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "') order by tarikh_belian ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm85_LM_PAGE_FOUND = 0 Then
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm85.L79_Text = Frm85.L79_Text + 1
                Frm85_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm85.L79_Text) Then
                    If Frm85.L79_Text <> 1 Then
                        Frm85.L79_Text = Frm85.L79_Text - 1
                        Frm85_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x

    With Frm85.LV4.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!tarikh_belian) Then 'Tarikh Belian
            .ListSubItems.Add , , rs!tarikh_belian
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_siri_Produk) Then 'No. Siri Produk
            .ListSubItems.Add , , rs!no_siri_Produk
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kod_Purity) Then 'Purity
            .ListSubItems.Add , , rs!kod_Purity
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kategori_Produk) Then 'Kategori Produk
            .ListSubItems.Add , , rs!kategori_Produk
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!nama_Supplier) Then 'Nama Supplier
            .ListSubItems.Add , , rs!nama_Supplier
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Berat) Then 'Berat (g)
            .ListSubItems.Add , , Format(rs!Berat, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kos_Belian_Gram) Then 'Rate Penerimaan (RM/g)
            .ListSubItems.Add , , Format(rs!kos_Belian_Gram, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!UPAH) Then 'Upah (RM)
            .ListSubItems.Add , , Format(rs!UPAH, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!SpreadValue) Then 'Spread (%)
            .ListSubItems.Add , , Format(rs!SpreadValue, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_lepas_spread) Then 'Harga Selepas Spread (RM)
            .ListSubItems.Add , , Format(rs!harga_lepas_spread, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!adjustment) Then 'Adjustment (RM)
            .ListSubItems.Add , , Format(rs!adjustment, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_item) Then 'Harga Belian (RM) : Tidak Campur Cukai GST
            .ListSubItems.Add , , Format(rs!harga_item, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Upah_Jualan) Then 'Upah Jualan (RM) : Pelanggan
            .ListSubItems.Add , , Format(rs!Upah_Jualan, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!Upah_Member) Then 'Upah Jualan (RM) : Ahli Biasa
            .ListSubItems.Add , , Format(rs!Upah_Member, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!Upah_Pengedar) Then 'Upah Jualan (RM) : Silver
            .ListSubItems.Add , , Format(rs!Upah_Pengedar, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Upah_RAF) Then 'Upah Jualan (RM) : Gold
            .ListSubItems.Add , , Format(rs!Upah_RAF, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!upah_normal_dealer) Then 'Upah Jualan (RM) : Platinum
            .ListSubItems.Add , , Format(rs!upah_normal_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!upah_master_dealer) Then 'Upah Jualan (RM) : Master Dealer
            .ListSubItems.Add , , Format(rs!upah_master_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!code_Supplier) Then 'Tetapan Harga Jualan (RM) : Pelanggan
            .ListSubItems.Add , , Format(rs!code_Supplier, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!HargaJualan_Member) Then 'Tetapan Harga Jualan (RM) : Pelanggan
            .ListSubItems.Add , , Format(rs!HargaJualan_Member, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!HargaJualan_Pengedar) Then 'Tetapan Harga Jualan (RM) : Pengedar
            .ListSubItems.Add , , Format(rs!HargaJualan_Pengedar, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!HargaJualan_RAF) Then 'Tetapan Harga Jualan (RM) : RAF
            .ListSubItems.Add , , Format(rs!HargaJualan_RAF, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!hargajualan_normal_dealer) Then 'Tetapan Harga Jualan (RM) : Normal Dealer
            .ListSubItems.Add , , Format(rs!hargajualan_normal_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!hargajualan_master_dealer) Then 'Tetapan Harga Jualan (RM) : Master Dealer
            .ListSubItems.Add , , Format(rs!hargajualan_master_dealer, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
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
        
        If Not IsNull(rs!dimension_Saiz) Then 'Tebal
            .ListSubItems.Add , , rs!dimension_Saiz
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!bill_No_Belian) Then 'No. Invoice Supplier
            .ListSubItems.Add , , rs!bill_No_Belian
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
        
        If Not IsNull(rs!cawangan) Then 'Cawangan
            .ListSubItems.Add , , rs!cawangan
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
rs.Open "select COUNT(ID) from Data_Database where (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

'#### Jumlah Bilangan Barang Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L54_Text = rs(0)
Else
    Frm85.L54_Text = 0
End If
'#### Jumlah Bilangan Barang Keseluruhan #### - End

'#### Jumlah Berat Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(Berat) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L55_Text = Format(rs(0), "#,##0.00 g")
Else
    Frm85.L55_Text = "0.00 g"
End If
'#### Jumlah Berat Keseluruhan #### - End

'#### Jumlah Modal Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_item) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L56_Text = "RM " & Format(rs(0), "#,##0.00")
Else
    Frm85.L56_Text = "RM 0.00"
End If
'#### Jumlah Modal Keseluruhan #### - End

If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    
    Frm85.Pic5.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    
    MsgBox "Tiada Data Stok Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_report_potong_barcode()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date
Dim Frm85_LM_TOTAL_PAGE As Double

Frm85_LM_TOTAL_PAGE = 0
Frm85_PAGE_SIZE = 34
x = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
        End If
    End If
End If

If Frm101.L46_Text = "Semua cawangan" Then

    Frm85_LM_SEARCH_5 = Null
    Frm85_LM_SEARCH_5_LOGIC = "<>"
    Frm85_LM_SEARCH_6 = Null
    Frm85_LM_SEARCH_6_LOGIC = "<>"
    
Else

    Frm85_LM_SEARCH_5 = Frm101.L46_Text
    Frm85_LM_SEARCH_5_LOGIC = "="
    Frm85_LM_SEARCH_6 = "HQ"
    Frm85_LM_SEARCH_6_LOGIC = "="
    
End If

Frm85_LM_PAGE_FOUND = 0

'#### Header Report ###
Frm85.L27_Text = "Report stok potong bagi item dengan no. siri produk [" & Frm101.L5_Text & "]." 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND (StatusItem='" & 12 & "' OR StatusItem='" & 20 & "' OR StatusItem='" & 28 & "' OR StatusItem='" & 22 & "') order by tarikh_belian ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm85_LM_PAGE_FOUND = 0 Then
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm85.L79_Text = Frm85.L79_Text + 1
                Frm85_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm85.L79_Text) Then
                    If Frm85.L79_Text <> 1 Then
                        Frm85.L79_Text = Frm85.L79_Text - 1
                        Frm85_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x

    With Frm85.LV5.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!tarikh_jualan1) Then 'Tarikh potong
            .ListSubItems.Add , , rs!tarikh_jualan1
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_siri_Produk) Then 'No. Siri Produk
            .ListSubItems.Add , , rs!no_siri_Produk
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kod_Purity) Then 'Purity
            .ListSubItems.Add , , rs!kod_Purity
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kategori_Produk) Then 'Kategori Produk
            .ListSubItems.Add , , rs!kategori_Produk
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!nama_Supplier) Then 'Nama Supplier
            .ListSubItems.Add , , rs!nama_Supplier
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!Berat) Then 'Berat (g)
            .ListSubItems.Add , , Format(rs!Berat, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!susut_berat) Then 'Susut berat (g)
            .ListSubItems.Add , , Format(rs!susut_berat, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!beza_berat) Then 'Susut berat (g)
            .ListSubItems.Add , , Format(rs!beza_berat, "#,##0.00")
            If IsNumeric(rs!beza_berat) Then Frm85_LM_BERAT = Frm85_LM_BERAT + rs!beza_berat 'Total Berat (g)
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!dulang) Then 'Dulang
            .ListSubItems.Add , , rs!dulang
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!cawangan) Then 'Cawangan
            .ListSubItems.Add , , rs!cawangan
        Else
            .ListSubItems.Add , , ""
        End If
        
    End With
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'Frm85.L28_Text = x 'Total Barang
'Frm85.L29_Text = Format(Frm85_LM_BERAT, "#,##0.00 g") 'Total Berat
'Frm85.L13_Text = Format(Frm85_LM_HARGA, "#,##0.00") 'Total Harga Belian

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND (StatusItem='" & 12 & "' OR StatusItem='" & 20 & "' OR StatusItem='" & 28 & "' OR StatusItem='" & 22 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

'#### Jumlah Bilangan Barang Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND (StatusItem='" & 12 & "' OR StatusItem='" & 20 & "' OR StatusItem='" & 28 & "' OR StatusItem='" & 22 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L57_Text = rs(0)
Else
    Frm85.L57_Text = 0
End If
'#### Jumlah Bilangan Barang Keseluruhan #### - End

'#### Jumlah Berat Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(beza_berat) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND (StatusItem='" & 12 & "' OR StatusItem='" & 20 & "' OR StatusItem='" & 28 & "' OR StatusItem='" & 22 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L58_Text = Format(rs(0), "#,##0.00 g")
Else
    Frm85.L58_Text = "0.00 g"
End If
'#### Jumlah Berat Keseluruhan #### - End

If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    
    Frm85.Pic6.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    
    MsgBox "Tiada Rekod Stok Potong Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_report_ansuran_barcode()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date
Dim Frm85_LM_TOTAL_PAGE As Double

Frm85_LM_TOTAL_PAGE = 0
Frm85_PAGE_SIZE = 34
x = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
        End If
    End If
End If

Frm85_LM_PAGE_FOUND = 0

'#### Header Report ###
Frm85.L35_Text = "Report jualan secara ansuran bagi item dengan no siri produk [" & Frm101.L5_Text & "]." 'Report Header"

'#### Report Jualan Dari Ansuran #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 27_senarai_ansuran where status='" & "Jelas" & "' AND no_siri_Produk='" & Frm101.L5_Text & "' order by tarikh_jelas ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm85_LM_PAGE_FOUND = 0 Then
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm85.L79_Text = Frm85.L79_Text + 1
                Frm85_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm85.L79_Text) Then
                    If Frm85.L79_Text <> 1 Then
                        Frm85.L79_Text = Frm85.L79_Text - 1
                        Frm85_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x
    Frm85.MSFlexGrid6.Rows = x + 1
    Frm85.MSFlexGrid6.TextMatrix(x, 0) = x 'No.
    Frm85.MSFlexGrid6.TextMatrix(x, 1) = Y 'No.
    If Not IsNull(rs!ID) Then Frm85.MSFlexGrid6.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh_jelas) Then Frm85.MSFlexGrid6.TextMatrix(x, 3) = rs!tarikh_jelas 'Tarikh
    If Not IsNull(rs!jenis_ansuran) Then 'Jenis Ansuran
        If rs!jenis_ansuran = 0 Then
            Frm85.MSFlexGrid6.TextMatrix(x, 4) = "Harga Semasa"
        ElseIf rs!jenis_ansuran = 1 Then
            Frm85.MSFlexGrid6.TextMatrix(x, 4) = "Harga Tetap"
        End If
    End If
    If Not IsNull(rs!no_siri_Produk) Then Frm85.MSFlexGrid6.TextMatrix(x, 5) = rs!no_siri_Produk 'No. Siri Produk
    If Not IsNull(rs!kategori_Produk) Then Frm85.MSFlexGrid6.TextMatrix(x, 6) = rs!kategori_Produk 'Kategori Produk
    If Not IsNull(rs!purity) Then Frm85.MSFlexGrid6.TextMatrix(x, 7) = rs!purity 'Purity
    If Not IsNull(rs!Berat_Asal) Then Frm85.MSFlexGrid6.TextMatrix(x, 8) = Format(rs!Berat_Asal, "#,##0.00") 'Berat Asal (g)
    If Not IsNull(rs!berat_jualan) Then
        Frm85.MSFlexGrid6.TextMatrix(x, 9) = Format(rs!berat_jualan, "#,##0.00") 'Berat Jualan (g)
        If IsNumeric(rs!berat_jualan) Then Frm85_LM_BERAT = Frm85_LM_BERAT + rs!berat_jualan
    End If
    If Not IsNull(rs!UPAH) Then Frm85.MSFlexGrid6.TextMatrix(x, 10) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
    If Not IsNull(rs!jumlah_bayaran) Then
        Frm85.MSFlexGrid6.TextMatrix(x, 11) = Format(rs!jumlah_bayaran, "#,##0.00") 'Harga Asal Jualan (RM)
        If IsNumeric(rs!jumlah_bayaran) Then Frm85_LM_HARGA = Frm85_LM_HARGA + rs!jumlah_bayaran
    End If
    If Not IsNull(rs!dulang) Then Frm85.MSFlexGrid6.TextMatrix(x, 12) = rs!dulang 'Dulang
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'#### Report Jualan Dari Ansuran #### - End

Frm85.L36_Text = x 'Total Barang
Frm85.L37_Text = Format(Frm85_LM_BERAT, "#,##0.00 g")  'Total Berat
Frm85.L38_Text = "RM " & Format(Frm85_LM_HARGA, "#,##0.00")  'Total Harga Jualan

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 27_senarai_ansuran where status='" & "Jelas" & "' AND no_siri_Produk='" & Frm101.L5_Text & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

'#### Jumlah Bilangan Barang Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 27_senarai_ansuran where status='" & "Jelas" & "' AND no_siri_Produk='" & Frm101.L5_Text & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L59_Text = rs(0)
Else
    Frm85.L59_Text = 0
End If
'#### Jumlah Bilangan Barang Keseluruhan #### - End

'#### Jumlah Berat Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(Berat_Jualan) from 27_senarai_ansuran where status='" & "Jelas" & "' AND no_siri_Produk='" & Frm101.L5_Text & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L60_Text = Format(rs(0), "#,##0.00 g")
Else
    Frm85.L60_Text = "0.00 g"
End If
'#### Jumlah Berat Keseluruhan #### - End

'#### Jumlah Harga Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_bayaran) from 27_senarai_ansuran where status='" & "Jelas" & "' AND no_siri_Produk='" & Frm101.L5_Text & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic
If Not rs.EOF Then
    Frm85.L61_Text = "RM " & Format(rs(0), "#,##0.00")
Else
    Frm85.L61_Text = "RM " & "0.00"
End If
'#### Jumlah Harga Keseluruhan #### - End

If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    
    Frm85.Pic9.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    
    MsgBox "Tiada Rekod Jualan Secara Ansuran Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_report_tempahan_barcode()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date
Dim Frm85_LM_TOTAL_PAGE As Double

Frm85_LM_TOTAL_PAGE = 0
Frm85_PAGE_SIZE = 34
x = 0
Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
        End If
    End If
End If

Frm85_LM_PAGE_FOUND = 0

If Frm101.L46_Text = "Semua cawangan" Then

    Frm85_LM_SEARCH_5 = Null
    Frm85_LM_SEARCH_5_LOGIC = "<>"
    Frm85_LM_SEARCH_6 = Null
    Frm85_LM_SEARCH_6_LOGIC = "<>"
    
Else

    Frm85_LM_SEARCH_5 = Frm101.L46_Text
    Frm85_LM_SEARCH_5_LOGIC = "="
    Frm85_LM_SEARCH_6 = "HQ"
    Frm85_LM_SEARCH_6_LOGIC = "="
    
End If

'#### Header Report ###
Frm85.L39_Text = "Report jualan secara tempahan bagi item dengan no. siri produk [" & Frm101.L5_Text & "]." 'Report Header"

'#### Report Jualan Dari Tempahan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 42_tempahan_siap where status_invoice = 1 AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm85_LM_PAGE_FOUND = 0 Then
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm85.L79_Text = Frm85.L79_Text + 1
                Frm85_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm85.L79_Text) Then
                    If Frm85.L79_Text <> 1 Then
                        Frm85.L79_Text = Frm85.L79_Text - 1
                        Frm85_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x

    With Frm85.LV6.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!tarikh) Then 'Tarikh
            .ListSubItems.Add , , rs!tarikh
        Else
            .ListSubItems.Add , , ""
        End If
        
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
            If IsNumeric(rs!berat_jualan) Then Frm85_LM_BERAT = Frm85_LM_BERAT + rs!berat_jualan
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!UPAH) Then  'Upah (RM)
            .ListSubItems.Add , , Format(rs!UPAH, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_dengan_gst) Then 'Harga Asal Jualan (RM)
            .ListSubItems.Add , , Format(rs!harga_dengan_gst, "#,##0.00")
            If IsNumeric(rs!harga_dengan_gst) Then Frm85_LM_HARGA = Frm85_LM_HARGA + rs!harga_dengan_gst
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!dulang) Then 'Dulang
            .ListSubItems.Add , , rs!dulang
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!cawangan) Then 'Cawangan
            .ListSubItems.Add , , rs!cawangan
        Else
            .ListSubItems.Add , , ""
        End If
        
    End With
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'#### Report Jualan Dari Tempahan #### - End

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 42_tempahan_siap where status_invoice = 1 AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

'#### Jumlah Bilangan Barang Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 42_tempahan_siap where status_invoice = 1 AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L62_Text = rs(0)
Else
    Frm85.L62_Text = 0
End If
'#### Jumlah Bilangan Barang Keseluruhan #### - End

'#### Jumlah Berat Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(Berat_Jualan) from 42_tempahan_siap where status_invoice = 1 AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L63_Text = Format(rs(0), "#,##0.00 g")
Else
    Frm85.L63_Text = "0.00 g"
End If
'#### Jumlah Berat Keseluruhan #### - End

'#### Jumlah Harga Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_dengan_gst) from 42_tempahan_siap where status_invoice = 1 AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L64_Text = "RM " & Format(rs(0), "#,##0.00")
Else
    Frm85.L64_Text = "RM " & "0.00"
End If
'#### Jumlah Harga Keseluruhan #### - End

If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    
    Frm85.Pic10.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    
    MsgBox "Tiada Rekod Jualan Secara Tempahan Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_report_belian_gb_barcode()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date
Dim Frm85_LM_TOTAL_PAGE As Double

x = 0
Y = 0
Frm85_PAGE_SIZE = 34
Frm85_LM_TOTAL_PAGE = 0

Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
        End If
    End If
End If

Frm85_LM_PAGE_FOUND = 0

'#### Header Report ###
Frm85.L71_Text = "Report belian gold bar bagi item dengan no. siri produk [" & Frm101.L5_Text & "]." 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND StatusItem <> 0 AND receiving_Status='" & 4 & "' order by tarikh_belian ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm85_LM_PAGE_FOUND = 0 Then
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm85.L79_Text = Frm85.L79_Text + 1
                Frm85_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm85.L79_Text) Then
                    If Frm85.L79_Text <> 1 Then
                        Frm85.L79_Text = Frm85.L79_Text - 1
                        Frm85_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x
    Frm85.MSFlexGrid8.Rows = x + 1
    Frm85.MSFlexGrid8.TextMatrix(x, 0) = x 'No.
    Frm85.MSFlexGrid8.TextMatrix(x, 1) = Y 'No.
    Frm85.MSFlexGrid8.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh_belian) Then Frm85.MSFlexGrid8.TextMatrix(x, 3) = rs!tarikh_belian 'Tarikh Belian
    If Not IsNull(rs!no_siri_Produk) Then Frm85.MSFlexGrid8.TextMatrix(x, 4) = rs!no_siri_Produk 'No. Siri Produk
    If Not IsNull(rs!kod_Purity) Then Frm85.MSFlexGrid8.TextMatrix(x, 5) = rs!kod_Purity 'Purity
    If Not IsNull(rs!kategori_Produk) Then Frm85.MSFlexGrid8.TextMatrix(x, 6) = rs!kategori_Produk 'Kategori Produk
    If Not IsNull(rs!nama_Supplier) Then Frm85.MSFlexGrid8.TextMatrix(x, 7) = rs!nama_Supplier 'Nama Supplier
    If Not IsNull(rs!Berat) Then
        Frm85.MSFlexGrid8.TextMatrix(x, 8) = Format(rs!Berat, "#,##0.00") 'Berat (g)
        If IsNumeric(rs!Berat) Then Frm85_LM_BERAT = Frm85_LM_BERAT + rs!Berat 'Total Berat (g)
    End If
    If Not IsNull(rs!kos_Belian_Gram) Then Frm85.MSFlexGrid8.TextMatrix(x, 9) = Format(rs!kos_Belian_Gram, "#,##0.00") 'Rate Penerimaan (RM/g)
    If Not IsNull(rs!UPAH) Then Frm85.MSFlexGrid8.TextMatrix(x, 10) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
    If Not IsNull(rs!SpreadValue) Then Frm85.MSFlexGrid8.TextMatrix(x, 11) = rs!SpreadValue 'Spread (%)
    If Not IsNull(rs!harga_lepas_spread) Then Frm85.MSFlexGrid8.TextMatrix(x, 12) = Format(rs!harga_lepas_spread, "#,##0.00") 'Harga Selepas Spread (RM)
    If Not IsNull(rs!adjustment) Then Frm85.MSFlexGrid8.TextMatrix(x, 13) = Format(rs!adjustment, "#,##0.00") 'Adjustment (RM)
    If Not IsNull(rs!harga_item) Then
        Frm85.MSFlexGrid8.TextMatrix(x, 14) = Format(rs!harga_item, "#,##0.00") 'Harga Belian (RM) : Tidak Campur Cukai GST
        If IsNumeric(rs!harga_item) Then Frm85_LM_HARGA = Frm85_LM_HARGA + rs!harga_item 'Total Harga Belian (RM) : Tidak Campur Cukai GST
    End If
    If Not IsNull(rs!dulang) Then Frm85.MSFlexGrid8.TextMatrix(x, 15) = rs!dulang 'Dulang
    If Not IsNull(rs!dimension_Panjang) Then Frm85.MSFlexGrid8.TextMatrix(x, 16) = rs!dimension_Panjang 'Panjang
    If Not IsNull(rs!dimension_Lebar) Then Frm85.MSFlexGrid8.TextMatrix(x, 17) = rs!dimension_Lebar 'Lebar
    If Not IsNull(rs!dimension_Saiz) Then Frm85.MSFlexGrid8.TextMatrix(x, 18) = rs!dimension_Saiz 'Tebal
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm85.L65_Text = x 'Total Barang
Frm85.L66_Text = Format(Frm85_LM_BERAT, "#,##0.00 g") 'Total Berat
Frm85.L67_Text = "RM " & Format(Frm85_LM_HARGA, "#,##0.00") 'Total Harga Belian

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND StatusItem <> 0 AND receiving_Status='" & 4 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

'#### Jumlah Bilangan Barang Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND StatusItem <> 0 AND receiving_Status='" & 4 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L68_Text = rs(0)
Else
    Frm85.L68_Text = 0
End If

rs.Close
Set rs = Nothing
'#### Jumlah Bilangan Barang Keseluruhan #### - End

'#### Jumlah Berat Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(Berat) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND StatusItem <> 0 AND receiving_Status='" & 4 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L69_Text = Format(rs(0), "#,##0.00 g")
Else
    Frm85.L69_Text = "0.00 g"
End If

rs.Close
Set rs = Nothing
'#### Jumlah Berat Keseluruhan #### - End

'#### Jumlah Modal Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_item) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND StatusItem <> 0 AND receiving_Status='" & 4 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L70_Text = "RM " & Format(rs(0), "#,##0.00")
Else
    Frm85.L70_Text = "RM 0.00"
End If

rs.Close
Set rs = Nothing
'#### Jumlah Modal Keseluruhan #### - End

If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    
    Frm85.Pic11.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    
    MsgBox "Tiada Rekod Belian Gold Bar Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_report_buyback_gb_barcode()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date
Dim Frm85_LM_TOTAL_PAGE As Double

x = 0
Y = 0
Frm85_PAGE_SIZE = 34
Frm85_LM_TOTAL_PAGE = 0

Frm85_LM_BERAT = 0
Frm85_LM_HARGA = 0

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
        End If
    End If
End If

Frm85_LM_PAGE_FOUND = 0

'#### Header Report ###
Frm85.L72_Text = "Report buyback / trade in gold bar dengan no. siri produk [" & Frm101.L5_Text & "]." 'Report Header"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND StatusItem <> 0 AND receiving_Status='" & 5 & "' order by tarikh_belian ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm85_LM_PAGE_FOUND = 0 Then
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm85.L79_Text = Frm85.L79_Text + 1
                Frm85_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm85.L79_Text) Then
                    If Frm85.L79_Text <> 1 Then
                        Frm85.L79_Text = Frm85.L79_Text - 1
                        Frm85_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x
    Frm85.MSFlexGrid9.Rows = x + 1
    Frm85.MSFlexGrid9.TextMatrix(x, 0) = x 'No.
    Frm85.MSFlexGrid9.TextMatrix(x, 1) = Y 'No.
    Frm85.MSFlexGrid9.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh_belian) Then Frm85.MSFlexGrid9.TextMatrix(x, 3) = rs!tarikh_belian 'Tarikh Belian
    If Not IsNull(rs!no_siri_Produk) Then Frm85.MSFlexGrid9.TextMatrix(x, 4) = rs!no_siri_Produk 'No. Siri Produk
    If Not IsNull(rs!kod_Purity) Then Frm85.MSFlexGrid9.TextMatrix(x, 5) = rs!kod_Purity 'Purity
    If Not IsNull(rs!kategori_Produk) Then Frm85.MSFlexGrid9.TextMatrix(x, 6) = rs!kategori_Produk 'Kategori Produk
    If Not IsNull(rs!nama_Supplier) Then Frm85.MSFlexGrid9.TextMatrix(x, 7) = rs!nama_Supplier 'Nama Supplier
    If Not IsNull(rs!Berat) Then
        Frm85.MSFlexGrid9.TextMatrix(x, 8) = Format(rs!Berat, "#,##0.00") 'Berat (g)
        If IsNumeric(rs!Berat) Then Frm85_LM_BERAT = Frm85_LM_BERAT + rs!Berat 'Total Berat (g)
    End If
    If Not IsNull(rs!kos_Belian_Gram) Then Frm85.MSFlexGrid9.TextMatrix(x, 9) = Format(rs!kos_Belian_Gram, "#,##0.00") 'Rate Penerimaan (RM/g)
    If Not IsNull(rs!UPAH) Then Frm85.MSFlexGrid9.TextMatrix(x, 10) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
    If Not IsNull(rs!SpreadValue) Then Frm85.MSFlexGrid9.TextMatrix(x, 11) = rs!SpreadValue 'Spread (%)
    If Not IsNull(rs!harga_lepas_spread) Then Frm85.MSFlexGrid9.TextMatrix(x, 12) = Format(rs!harga_lepas_spread, "#,##0.00") 'Harga Selepas Spread (RM)
    If Not IsNull(rs!adjustment) Then Frm85.MSFlexGrid9.TextMatrix(x, 13) = Format(rs!adjustment, "#,##0.00") 'Adjustment (RM)
    If Not IsNull(rs!harga_item) Then
        Frm85.MSFlexGrid9.TextMatrix(x, 14) = Format(rs!harga_item, "#,##0.00") 'Harga Belian (RM) : Tidak Campur Cukai GST
        If IsNumeric(rs!harga_item) Then Frm85_LM_HARGA = Frm85_LM_HARGA + rs!harga_item 'Total Harga Belian (RM) : Tidak Campur Cukai GST
    End If
    If Not IsNull(rs!dulang) Then Frm85.MSFlexGrid9.TextMatrix(x, 15) = rs!dulang 'Dulang
    If Not IsNull(rs!dimension_Panjang) Then Frm85.MSFlexGrid9.TextMatrix(x, 16) = rs!dimension_Panjang 'Panjang
    If Not IsNull(rs!dimension_Lebar) Then Frm85.MSFlexGrid9.TextMatrix(x, 17) = rs!dimension_Lebar 'Lebar
    If Not IsNull(rs!dimension_Saiz) Then Frm85.MSFlexGrid9.TextMatrix(x, 18) = rs!dimension_Saiz 'Tebal
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm85.L73_Text = x 'Total Barang
Frm85.L74_Text = Format(Frm85_LM_BERAT, "#,##0.00 g") 'Total Berat
Frm85.L75_Text = "RM " & Format(Frm85_LM_HARGA, "#,##0.00") 'Total Harga Belian

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND StatusItem <> 0 AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

'#### Jumlah Bilangan Barang Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND StatusItem <> 0 AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L76_Text = rs(0)
Else
    Frm85.L76_Text = 0
End If

rs.Close
Set rs = Nothing
'#### Jumlah Bilangan Barang Keseluruhan #### - End

'#### Jumlah Berat Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(Berat) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND StatusItem <> 0 AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L77_Text = Format(rs(0), "#,##0.00 g")
Else
    Frm85.L77_Text = "0.00 g"
End If

rs.Close
Set rs = Nothing
'#### Jumlah Berat Keseluruhan #### - End

'#### Jumlah Modal Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_item) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND StatusItem <> 0 AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L78_Text = "RM " & Format(rs(0), "#,##0.00")
Else
    Frm85.L78_Text = "RM 0.00"
End If

rs.Close
Set rs = Nothing
'#### Jumlah Modal Keseluruhan #### - End

If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    
    Frm85.Pic12.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    
    MsgBox "Tiada Rekod Buyback / Trade In Gold Bar Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_header_report_trade_in_agen()
'### REPORT TRADE IN AGEN
'on error resume next
'#### Header Report #### - Start
Frm85.MSFlexGrid10.Clear
Frm85.MSFlexGrid10.Rows = 1
Frm85.MSFlexGrid10.RowHeight(0) = 1500
Frm85.MSFlexGrid10.FormatString = "<No.|<No.|<No. ID|<Tarikh|<No. Invoice|<Purity|<Berat Asal (g)|<Kadar Tukaran|<Berat Selepas Tukaran (g)"

Frm85.MSFlexGrid10.ColWidth(0) = 0 'No.
Frm85.MSFlexGrid10.ColWidth(1) = 600 'No.
Frm85.MSFlexGrid10.ColAlignment(1) = 4

Frm85.MSFlexGrid10.ColWidth(2) = 0 'No. ID
Frm85.MSFlexGrid10.ColWidth(3) = 1200 'Tarikh
Frm85.MSFlexGrid10.ColAlignment(3) = 4

Frm85.MSFlexGrid10.ColWidth(4) = 1200 'No. Invoice
Frm85.MSFlexGrid10.ColAlignment(4) = 4

Frm85.MSFlexGrid10.ColWidth(5) = 1200 'Purity
Frm85.MSFlexGrid10.ColAlignment(5) = 4

Frm85.MSFlexGrid10.ColWidth(6) = 1200 'Berat Asal (g)
Frm85.MSFlexGrid10.ColAlignment(6) = 7

Frm85.MSFlexGrid10.ColWidth(7) = 1200 'Kadar Tukaran
Frm85.MSFlexGrid10.ColAlignment(7) = 7

Frm85.MSFlexGrid10.ColWidth(8) = 1200 'Berat Selepas Tukaran (g)
Frm85.MSFlexGrid10.ColAlignment(8) = 7
'#### Header Report #### - End
End Sub
Sub Frm85_header_report_trade_in_susut_nilai()
'### REPORT TRADE IN AGEN
'on error resume next
'#### Header Report #### - Start
Frm85.MSFlexGrid10.Clear
Frm85.MSFlexGrid10.Rows = 1
Frm85.MSFlexGrid10.RowHeight(0) = 1500
Frm85.MSFlexGrid10.FormatString = "<No.|<No.|<No. ID|<Tarikh|<No. Invoice|<Jenis|<Berat (g)|<Harga Semasa (RM/g)|<Harga (RM)|<Nama Pekerja"

Frm85.MSFlexGrid10.ColWidth(0) = 0 'No.
Frm85.MSFlexGrid10.ColWidth(1) = 600 'No.
Frm85.MSFlexGrid10.ColAlignment(1) = 4

Frm85.MSFlexGrid10.ColWidth(2) = 0 'No. ID
Frm85.MSFlexGrid10.ColWidth(3) = 1200 'Tarikh
Frm85.MSFlexGrid10.ColAlignment(3) = 4

Frm85.MSFlexGrid10.ColWidth(4) = 1500 'No. Invoice

Frm85.MSFlexGrid10.ColWidth(5) = 1500 'Jenis

Frm85.MSFlexGrid10.ColWidth(6) = 1300 'Berat (g)
Frm85.MSFlexGrid10.ColAlignment(6) = 7

Frm85.MSFlexGrid10.ColWidth(7) = 1300 'Harga Semasa
Frm85.MSFlexGrid10.ColAlignment(7) = 7

Frm85.MSFlexGrid10.ColWidth(8) = 1300 'Harga
Frm85.MSFlexGrid10.ColAlignment(8) = 7

Frm85.MSFlexGrid10.ColWidth(9) = 2000
'#### Header Report #### - End
End Sub
Sub frm85_excel_susut_nilai()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then

    x = 0
    
    If Frm101.L9_Text = 1 Then
        TM = Frm101.L7_Text 'Tarikh Mula
        TA = Frm101.L8_Text 'Tarikh Akhir
    End If
    If Frm101.L47_Text = "Semua" Then
        Frm85_LM_SEARCH_20 = Null
        Frm85_LM_SEARCH_20_LOGIC = "<>"
    Else
        If InStr(1, Frm101.L47_Text, "  |  ") <> 0 Then
            Frm84_LM_EMP_NO = Split(Frm101.L47_Text, "  |  ")(1)
            Frm84_LM_EMP_NAMA = Split(Frm101.L47_Text, "  |  ")(0)
        End If
        Frm85_LM_SEARCH_20 = Frm84_LM_EMP_NAMA
        Frm85_LM_SEARCH_20_LOGIC = "="
    End If

    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh
        .Columns("C").ColumnWidth = 20 'No. Invoice
        .Columns("D").ColumnWidth = 20 'Jenis
        .Columns("E").ColumnWidth = 20 'Berat (g)
        .Columns("F").ColumnWidth = 20 'Harga Semasa (RM/g)
        .Columns("G").ColumnWidth = 20 'Harga (RM)
        .Columns("H").ColumnWidth = 20 'Nama Pekerja
        
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row

        .Cells(7, 1) = Frm85.L82_Text 'Header Report
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh"
        .Cells(8, 3) = "No. Invoice"
        .Cells(8, 4) = "Jenis"
        .Cells(8, 5) = "Berat (g)"
        .Cells(8, 6) = "Harga Semasa (RM/g)"
        .Cells(8, 7) = "Harga (RM)"
        .Cells(8, 8) = "Nama Pekerja"
        
        For i = 1 To 8
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
        
        '#### Report Trade In Dari Agen #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select * from 93_trade_in_susut_niai where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND status = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select * from 93_trade_in_susut_niai where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh) Then .Cells(8 + x, 2) = "'" & rs!tarikh 'Tarikh Jualan
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter

            If Not IsNull(rs!no_invoice) Then .Cells(8 + x, 3) = rs!no_invoice 'No. Resit

            If Not IsNull(rs!jenis) Then 'Jenis
                If rs!jenis = 0 Or rs!jenis = 1 Then
                    If rs!jenis = 0 Then .Cells(8 + x, 4) = "Trade In"
                    If rs!jenis = 1 Then .Cells(8 + x, 4) = "Buyback"
        
                    .Cells(8 + x, 5).NumberFormat = "#,##0.00"
                    .Cells(8 + x, 5).HorizontalAlignment = xlRight
                    If Not IsNull(rs!Berat) Then .Cells(8 + x, 5) = Format(rs!Berat, "#,##0.00")
                    
                    .Cells(8 + x, 6).NumberFormat = "#,##0.00"
                    .Cells(8 + x, 6).HorizontalAlignment = xlRight
                    If Not IsNull(rs!harga_Semasa) Then .Cells(8 + x, 6) = Format(rs!harga_Semasa, "#,##0.00")
        
                    .Cells(8 + x, 7).NumberFormat = "#,##0.00"
                    .Cells(8 + x, 7).HorizontalAlignment = xlRight
                    If Not IsNull(rs!harga) Then .Cells(8 + x, 7) = Format(rs!harga, "#,##0.00")
                End If
                If rs!jenis = 2 Then
                    .Cells(8 + x, 4) = "Caj Pertukaran"
                    If Not IsNull(rs!harga) Then .Cells(8 + x, 7) = Format(rs!harga, "#,##0.00")
                End If
                If Not IsNull(rs!nama_pekerja) Then .Cells(8 + x, 8) = rs!nama_pekerja
            End If

            For Col = 1 To 8
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 1
        
        '### Jumlah Data ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select SUM(berat) , SUM(harga) from 93_trade_in_susut_niai where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select SUM(berat) , SUM(harga) from 93_trade_in_susut_niai where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not IsNull(rs(0)) Then
            .Cells(8 + Y, 1) = "Berat : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat : " & "0.00 g"
        End If
        Y = Y + 1
        If Not IsNull(rs(1)) Then
            .Cells(8 + Y, 1) = "Harga : RM " & Format(rs(1), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Harga : RM " & "0.00"
        End If
        
        rs.Close
        Set rs = Nothing

        Y = Y + 4
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

End Sub
Sub Frm85_report_trade_in_susut_nilai()
'on error resume next

'### REPORT TRADE IN AGEN
'Report ini hanya boleh difilter melalui 2 krateria sahaja iaitu TARIKH dan PURITY

Dim Frm85_LM_BERAT As Double
Dim TA As Date
Dim TM As Date
Dim Frm85_LM_TOTAL_PAGE As Double

Frm85_LM_TOTAL_PAGE = 0
Frm85_PAGE_SIZE = 34
x = 0
Y = 0
Frm85_LM_BERAT = 0

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L47_Text = "Semua" Then
    Frm85_LM_SEARCH_20 = Null
    Frm85_LM_SEARCH_20_LOGIC = "<>"
Else
    If InStr(1, Frm101.L47_Text, "  |  ") <> 0 Then
        Frm84_LM_EMP_NO = Split(Frm101.L47_Text, "  |  ")(1)
        Frm84_LM_EMP_NAMA = Split(Frm101.L47_Text, "  |  ")(0)
    End If
    Frm85_LM_SEARCH_20 = Frm84_LM_EMP_NAMA
    Frm85_LM_SEARCH_20_LOGIC = "="
End If

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
        End If
    End If
End If

Frm85_LM_PAGE_FOUND = 0

'#### Header Report ###
If Frm101.L9_Text = 0 Then Frm85.L82_Text = "Report trade in 0% susut nilai." 'Report Header"
If Frm101.L9_Text = 1 Then Frm85.L82_Text = "Report trade in 0% susut nilai dari " & TM & " hingga " & TA & "."  'Report Header"

'#### Report Trade In Dari Agen #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select * from 93_trade_in_susut_niai where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND status = 1 order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select * from 93_trade_in_susut_niai where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm85_LM_PAGE_FOUND = 0 Then
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm85.L79_Text = Frm85.L79_Text + 1
                Frm85_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm85.L79_Text) Then
                    If Frm85.L79_Text <> 1 Then
                        Frm85.L79_Text = Frm85.L79_Text - 1
                        Frm85_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x
    
    Frm85.MSFlexGrid10.Rows = x + 1
    Frm85.MSFlexGrid10.TextMatrix(x, 0) = x 'No.
    Frm85.MSFlexGrid10.TextMatrix(x, 1) = Y 'No.
    If Not IsNull(rs!ID) Then Frm85.MSFlexGrid10.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm85.MSFlexGrid10.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!no_invoice) Then Frm85.MSFlexGrid10.TextMatrix(x, 4) = rs!no_invoice 'No. Voucher
    If Not IsNull(rs!jenis) Then 'Jenis
        If rs!jenis = 0 Or rs!jenis = 1 Then
            If rs!jenis = 0 Then Frm85.MSFlexGrid10.TextMatrix(x, 5) = "Trade In"
            If rs!jenis = 1 Then Frm85.MSFlexGrid10.TextMatrix(x, 5) = "Buyback"
            
            If Not IsNull(rs!Berat) Then Frm85.MSFlexGrid10.TextMatrix(x, 6) = Format(rs!Berat, "#,##0.00")
            If Not IsNull(rs!harga_Semasa) Then Frm85.MSFlexGrid10.TextMatrix(x, 7) = Format(rs!harga_Semasa, "#,##0.00")
            If Not IsNull(rs!harga) Then Frm85.MSFlexGrid10.TextMatrix(x, 8) = Format(rs!harga, "#,##0.00")
        End If
        If rs!jenis = 2 Then
            Frm85.MSFlexGrid10.TextMatrix(x, 5) = "Caj Pertukaran"
            If Not IsNull(rs!harga) Then Frm85.MSFlexGrid10.TextMatrix(x, 8) = Format(rs!harga, "#,##0.00")
        End If
        If Not IsNull(rs!nama_pekerja) Then Frm85.MSFlexGrid10.TextMatrix(x, 9) = rs!nama_pekerja
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'#### Report Trade In Dari Agen #### - End

Frm85.L83_Text = Format(0, "#,##0.00 g") 'Total Berat (Paparan ini)
Frm85.L84_Text = Format(0, "#,##0.00") 'Total Berat (Paparan ini)

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select COUNT(ID) , SUM(berat) , SUM(harga) from 93_trade_in_susut_niai where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND status = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select COUNT(ID) , SUM(berat) , SUM(harga) from 93_trade_in_susut_niai where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

If Not IsNull(rs(1)) Then Frm85.L83_Text = Format(rs(1), "#,##0.00 g") 'Total Berat (Paparan ini)
If Not IsNull(rs(2)) Then Frm85.L84_Text = Format(rs(2), "#,##0.00") 'Total Berat (Paparan ini)

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    
    Call Frm85_berat_purity
    Frm85.Pic13.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    
    MsgBox "Tiada rekod trade in oleh dijumpai.", vbInformation, "Info"
End If

End Sub
Sub Frm85_report_trade_in_agen()
'on error resume next

'### REPORT TRADE IN AGEN
'Report ini hanya boleh difilter melalui 2 krateria sahaja iaitu TARIKH dan PURITY

Dim Frm85_LM_BERAT As Double
Dim TA As Date
Dim TM As Date
Dim Frm85_LM_TOTAL_PAGE As Double

Frm85_LM_TOTAL_PAGE = 0
Frm85_PAGE_SIZE = 34
x = 0
Y = 0
Frm85_LM_BERAT = 0

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If

LM_START_ROW = Frm101.L43_Text

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm85_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm85_PAGE_SIZE
        End If
    End If
End If

Frm85_LM_PAGE_FOUND = 0

'#### Header Report ###
If Frm101.L9_Text = 0 Then Frm85.L82_Text = "Report trade in dari agen bagi purity [" & Frm101.L5_Text & "]." 'Report Header"
If Frm101.L9_Text = 1 Then Frm85.L82_Text = "Report trade in dari agen bagi purity [" & Frm101.L5_Text & "] dari " & TM & " hingga " & TA & "."  'Report Header"

'#### Report Trade In Dari Agen #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select * from 50_belian_emas_agen where status='" & 1 & "' AND kod_purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select * from 50_belian_emas_agen where status='" & 1 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm85_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm85_LM_PAGE_FOUND = 0 Then
        If Frm85.L80_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm85.L79_Text = Frm85.L79_Text + 1
                Frm85_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm85.L79_Text) Then
                    If Frm85.L79_Text <> 1 Then
                        Frm85.L79_Text = Frm85.L79_Text - 1
                        Frm85_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm85.L79_Text - 1) * Frm85_PAGE_SIZE) + x
    
    Frm85.MSFlexGrid10.Rows = x + 1
    Frm85.MSFlexGrid10.TextMatrix(x, 0) = x 'No.
    Frm85.MSFlexGrid10.TextMatrix(x, 1) = Y 'No.
    If Not IsNull(rs!ID) Then Frm85.MSFlexGrid10.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm85.MSFlexGrid10.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!no_invoice) Then Frm85.MSFlexGrid10.TextMatrix(x, 4) = rs!no_invoice 'No. Voucher
    If Not IsNull(rs!kod_Purity) Then Frm85.MSFlexGrid10.TextMatrix(x, 5) = rs!kod_Purity 'Purity (Kod purity)
    If Not IsNull(rs!Berat_Asal) Then
        Frm85.MSFlexGrid10.TextMatrix(x, 6) = Format(rs!Berat_Asal, "#,##0.00") 'Berat Asal (g)
        If IsNumeric(rs!Berat_Asal) Then Frm85_LM_BERAT = Frm85_LM_BERAT + rs!Berat_Asal
    End If
    If Not IsNull(rs!kadar_tukaran) Then Frm85.MSFlexGrid10.TextMatrix(x, 7) = rs!kadar_tukaran 'Kadar Tukaran
    If Not IsNull(rs!berat_tukaran) Then Frm85.MSFlexGrid10.TextMatrix(x, 8) = Format(rs!berat_tukaran, "#,##0.00") 'Berat Selepas Tukaran (g)
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'#### Report Trade In Dari Agen #### - End

Frm85.L83_Text = Format(Frm85_LM_BERAT, "#,##0.00 g") 'Total Berat (Paparan ini)

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select COUNT(ID) from 50_belian_emas_agen where status='" & 1 & "' AND kod_purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select COUNT(ID) from 50_belian_emas_agen where status='" & 1 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85_LM_TOTAL_PAGE = Format(rs(0) / Frm85_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm85_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm85_LM_PAGE = Split(Frm85_LM_TOTAL_PAGE, ".")(0)
        Frm85_LM_PAGE_LEBIHAN = Split(Frm85_LM_TOTAL_PAGE, ".")(1)
        
        If Frm85_LM_PAGE_LEBIHAN <> "00" Then
            Frm85.L81_Text = Frm85_LM_PAGE + 1
        Else
            Frm85.L81_Text = Frm85_LM_PAGE
        End If
        
    Else
    
        Frm85.L81_Text = Frm85_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm85.L81_Text = 0
    End If
Else
    Frm85.L81_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm85.L81_Text = vbNullString Then
    Frm85.L81_Text = 0
End If
'### Jumlah Data ### - End

'#### Jumlah Berat Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm101.L9_Text = 0 Then rs.Open "select SUM(Berat_Asal) from 50_belian_emas_agen where status='" & 1 & "' AND kod_purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
If Frm101.L9_Text = 1 Then rs.Open "select SUM(Berat_Asal) from 50_belian_emas_agen where status='" & 1 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm85.L84_Text = Format(rs(0), "#,##0.00 g")
Else
    Frm85.L84_Text = "0.00 g"
End If
'#### Jumlah Berat Keseluruhan #### - End

rs.Close
Set rs = Nothing

If Frm85.L84_Text = vbNullString Then
    Frm85.L84_Text = "0.00 g"
End If


If x <> 0 Then
    Frm101.L43_Text = LM_START_ROW
End If

If x <> 0 Then
    Frm85.Show
    Frm101.Hide
    
    Call Frm85_berat_purity
    Frm85.Pic13.Visible = True
    
    Frm85.L80_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm85.L80_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    
    MsgBox "Tiada rekod trade in oleh agen dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm85_berat_purity()
'on error resume next
Dim rs1 As ADODB.Recordset
Dim TA As Date
Dim TM As Date

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If

A1 = vbNullString
B1 = vbNullString

'Frm85.L85_Text = A1
'Frm85.L86_Text = B1

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where Kod_Metal_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!Kod_Metal_Purity) Then
        '#### Jumlah Berat Keseluruhan #### - Start
        Frm85_LM_BERAT = 0
        
        Set rs1 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs1.Open "select SUM(berat_asal) from 50_belian_emas_agen where status='" & 1 & "' AND kod_purity='" & rs!Kod_Metal_Purity & "'", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs1.Open "select SUM(berat_asal) from 50_belian_emas_agen where status='" & 1 & "' AND kod_purity='" & rs!Kod_Metal_Purity & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not IsNull(rs1(0)) Then
            Frm85_LM_BERAT = Format(rs1(0), "#,##0.00 g")
        Else
            Frm85_LM_BERAT = Format(0, "#,##0.00 g")
        End If
        '#### Jumlah Berat Keseluruhan #### - End
        
        rs1.Close
        Set rs1 = Nothing
        
        A1 = A1 & rs!Metal_Purity & vbCrLf
        B1 = B1 & Frm85_LM_BERAT & vbCrLf
        
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'Frm85.L85_Text = A1
'Frm85.L86_Text = B1
End Sub
Sub Frm85_excel_trade_in()
'on error resume next
'REPORT TRADE IN AGEN - EXCEL
Dim TA As Date
Dim TM As Date

Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem akan mengambil masa untuk mengeluarkan report." & vbCrLf & _
        "" & vbCrLf & _
        "sila tunggu sehingga sistem siap keluarkan report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then

    x = 0
    
    If Frm101.L9_Text = 1 Then
        TM = Frm101.L7_Text 'Tarikh Mula
        TA = Frm101.L8_Text 'Tarikh Akhir
    End If
    If Frm101.L5_Text = "Semua Purity" Then
        Frm85_LM_SEARCH_1 = Null
        Frm85_LM_SEARCH_1_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_1 = Frm101.L5_Text
        Frm85_LM_SEARCH_1_LOGIC = "="
    End If
    
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh
        .Columns("C").ColumnWidth = 20 'No. Invoice
        .Columns("D").ColumnWidth = 20 'Purity
        .Columns("E").ColumnWidth = 20 'Berat Asal (g)
        .Columns("F").ColumnWidth = 20 'Kadar Tukaran
        .Columns("G").ColumnWidth = 20 'Berat Selepas Tukaran (g)
    
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        .Cells(7, 1) = Frm85.L14_Text 'Header Report
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh"
        .Cells(8, 3) = "No. Invoice"
        .Cells(8, 4) = "Purity"
        .Cells(8, 5) = "Berat Asal (g)"
        .Cells(8, 6) = "Kadar Tukaran"
        .Cells(8, 7) = "Berat Selepas Tukaran (g)"
        
        For i = 1 To 7
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
        
        '#### Header Report ###
        .Cells(7, 1) = Frm85.L82_Text 'Report Header"
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select * from 50_belian_emas_agen where status='" & 1 & "' AND kod_purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select * from 50_belian_emas_agen where status='" & 1 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh) Then .Cells(8 + x, 2) = "'" & rs!tarikh 'Tarikh
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_invoice) Then .Cells(8 + x, 3) = rs!no_invoice 'No. Voucher
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kod_Purity) Then .Cells(8 + x, 4) = rs!kod_Purity 'Purity (Kod purity)
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!Berat_Asal) Then 'Berat Asal (g)
                .Cells(8 + x, 5) = Format(rs!Berat_Asal, "#,##0.00")
            Else
                .Cells(8 + x, 5) = Format(0, "#,##0.00")
            End If
            .Cells(8 + x, 5).HorizontalAlignment = xlRight
            .Cells(8 + x, 5).NumberFormat = "#,##0.00"
            
            If Not IsNull(rs!kadar_tukaran) Then 'Kadar Tukaran
                .Cells(8 + x, 6) = rs!kadar_tukaran
            Else
                .Cells(8 + x, 6) = "0.00"
            End If
            .Cells(8 + x, 6).HorizontalAlignment = xlRight

            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!berat_tukaran) Then 'Berat Selepas Tukaran (g)
                .Cells(8 + x, 7) = Format(rs!berat_tukaran, "#,##0.00")
            Else
                .Cells(8 + x, 7) = "0.00"
            End If
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"

            For Col = 1 To 7
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 2
        
        '#### Jumlah Berat Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select SUM(Berat_Asal) from 50_belian_emas_agen where status='" & 1 & "' AND kod_purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select SUM(Berat_Asal) from 50_belian_emas_agen where status='" & 1 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Berat keseluruhan : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat keseluruhan : " & "0.00 g"
        End If
        '#### Jumlah Berat Keseluruhan #### - End
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = vbNullString Then
            .Cells(8 + Y, 1) = "Berat keseluruhan : " & "0.00 g"
        End If
        
        Y = Y + 1
        .Cells(8 + Y, 1) = "*** Berat keseluruhan adalah merujuk kepada BERAT ASAL."
        
        Y = Y + 3
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
End Sub
Sub Frm85_overall_beli()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    
    x = 0
    
    Frm85_LM_BERAT = 0
    Frm85_LM_HARGA = 0
    
    If Frm101.L9_Text = 1 Then
        TM = Frm101.L7_Text 'Tarikh Mula
        TA = Frm101.L8_Text 'Tarikh Akhir
    End If
    If Frm101.L5_Text = "Semua Purity" Then
        Frm85_LM_SEARCH_1 = Null
        Frm85_LM_SEARCH_1_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_1 = Frm101.L5_Text
        Frm85_LM_SEARCH_1_LOGIC = "="
    End If
    If Frm101.L6_Text = "Semua Kategori Produk" Then
        Frm85_LM_SEARCH_2 = Null
        Frm85_LM_SEARCH_2_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_2 = Frm101.L6_Text
        Frm85_LM_SEARCH_2_LOGIC = "="
    End If
    If Frm101.L34_Text = "Semua Dulang" Then
        Frm85_LM_SEARCH_3 = Null
        Frm85_LM_SEARCH_3_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_3 = Frm101.L34_Text
        Frm85_LM_SEARCH_3_LOGIC = "="
    End If
    If Frm101.L37_Text = "Semua Supplier" Then
        Frm85_LM_SEARCH_4 = Null
        Frm85_LM_SEARCH_4_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_4 = Frm101.L37_Text
        Frm85_LM_SEARCH_4_LOGIC = "="
    End If
    If Frm101.L46_Text = "Semua cawangan" Then
    
        Frm85_LM_SEARCH_5 = Null
        Frm85_LM_SEARCH_5_LOGIC = "<>"
        Frm85_LM_SEARCH_6 = Null
        Frm85_LM_SEARCH_6_LOGIC = "<>"
        
    Else
    
        Frm85_LM_SEARCH_5 = Frm101.L46_Text
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
    
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh Belian
        .Columns("C").ColumnWidth = 20 'No. Siri Produk
        .Columns("D").ColumnWidth = 20 'Purity
        .Columns("E").ColumnWidth = 40 'Kategori Produk
        .Columns("F").ColumnWidth = 40 'Supplier
        .Columns("G").ColumnWidth = 20 'Berat (g)
        .Columns("H").ColumnWidth = 20 'Rate Penerimaan (RM/g)
        .Columns("I").ColumnWidth = 20 'Upah (RM)
        .Columns("J").ColumnWidth = 20 'Spread (%)
        .Columns("K").ColumnWidth = 20 'Harga Selepas Spread (RM)
        .Columns("L").ColumnWidth = 20 'Adjustment (RM)
        .Columns("M").ColumnWidth = 20 'Harga Belian (RM)
        .Columns("N").ColumnWidth = 20 'Upah Jualan (RM) : Pelanggan
        .Columns("O").ColumnWidth = 20 'Upah Jualan (RM) : Ahli Biasa
        .Columns("P").ColumnWidth = 20 'Upah Jualan (RM) : Silver
        .Columns("Q").ColumnWidth = 20 'Upah Jualan (RM) : Gold
        .Columns("R").ColumnWidth = 20 'Upah Jualan (RM) : Platinum
        .Columns("S").ColumnWidth = 0 'Upah Jualan (RM) : Master Dealer
        .Columns("T").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Pelanggan
        .Columns("U").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Member
        .Columns("V").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Silver
        .Columns("W").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Gold
        .Columns("X").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Platinum
        .Columns("Y").ColumnWidth = 0 'Tetapan Harga Jualan (RM) : Master Dealer
        .Columns("Z").ColumnWidth = 20 'Dulang
        .Columns("AA").ColumnWidth = 20 'Panjang
        .Columns("AB").ColumnWidth = 20 'Lebar
        .Columns("AC").ColumnWidth = 20 'Saiz
        .Columns("AD").ColumnWidth = 20 'No. Invoice Dari Supplier
        .Columns("AE").ColumnWidth = 20 'Code 1
        .Columns("AF").ColumnWidth = 20 'Code 2
        .Columns("AG").ColumnWidth = 40 'Cawangan
        
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        '#### Header Report ###
        If Frm101.L9_Text = 0 Then .Cells(7, 1) = "Report Belian Bagi Supplier [" & Frm101.L37_Text & "] , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "]." 'Report Header"
        If Frm101.L9_Text = 1 Then .Cells(7, 1) = "Report Belian Bagi Supplier [" & Frm101.L37_Text & "] , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "] Dari " & TM & " Hingga " & TA 'Report Header"

        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Belian"
        .Cells(8, 3) = "No. Siri Produk"
        .Cells(8, 4) = "Purity"
        .Cells(8, 5) = "Kategori Produk"
        .Cells(8, 6) = "Supplier"
        .Cells(8, 7) = "Berat (g)"
        .Cells(8, 8) = "Rate Penerimaan (RM/g)"
        .Cells(8, 9) = "Upah (RM)"
        .Cells(8, 10) = "Spread (%)"
        .Cells(8, 11) = "Harga Selepas Spread (RM)"
        .Cells(8, 12) = "Adjustment (RM)"
        .Cells(8, 13) = "Harga Belian Termasuk GST (RM)"
        .Cells(8, 14) = "Upah Jualan (RM) : Pelanggan"
        .Cells(8, 15) = "Upah Jualan (RM) : Ahli"
        .Cells(8, 16) = "Upah Jualan (RM) : Silver"
        .Cells(8, 17) = "Upah Jualan (RM) : Gold"
        .Cells(8, 18) = "Upah Jualan (RM) : Platinum"
        .Cells(8, 19) = "Upah Jualan (RM) : Master Dealer"
        .Cells(8, 20) = "Tetapan Harga Jualan (RM) : Pelanggan"
        .Cells(8, 21) = "Tetapan Harga Jualan (RM) : Ahli"
        .Cells(8, 22) = "Tetapan Harga Jualan (RM) : Silver"
        .Cells(8, 23) = "Tetapan Harga Jualan (RM) : Gold"
        .Cells(8, 24) = "Tetapan Harga Jualan (RM) : Platinum"
        .Cells(8, 25) = "Tetapan Harga Jualan (RM) : Master Dealer"
        .Cells(8, 26) = "Dulang"
        .Cells(8, 27) = "Panjang"
        .Cells(8, 28) = "Lebar"
        .Cells(8, 29) = "Saiz"
        .Cells(8, 30) = "No. Invoice"
        .Cells(8, 31) = "Code 1"
        .Cells(8, 32) = "Code 2"
        .Cells(8, 33) = "Cawangan"
        
        For i = 1 To 33
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select * from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND (receiving_Status='" & 1 & "' OR receiving_Status='" & 0 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select * from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 1 & "' OR receiving_Status='" & 0 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh_belian) Then .Cells(8 + x, 2) = "'" & rs!tarikh_belian 'Tarikh Belian
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 3) = rs!no_siri_Produk 'No. Siri Produk
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!purity) Then .Cells(8 + x, 4) = rs!purity 'Purity
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 5) = rs!kategori_Produk 'Kategori Produk
            
            If Not IsNull(rs!nama_Supplier) Then .Cells(8 + x, 6) = rs!nama_Supplier 'Nama Supplier
            
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!Berat) Then
                .Cells(8 + x, 7) = Format(rs!Berat, "#,##0.00") 'Berat (g)
            Else
                .Cells(8 + x, 7) = "0.00" 'Berat (g)
            End If
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!kos_Belian_Gram) Then
                .Cells(8 + x, 8) = Format(rs!kos_Belian_Gram, "#,##0.00") 'Rate Penerimaan (RM/g)
            Else
                .Cells(8 + x, 8) = "0.00" 'Rate Penerimaan (RM/g)
            End If
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!UPAH) Then
                .Cells(8 + x, 9) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
            Else
                .Cells(8 + x, 9) = "0.00" 'Upah (RM)
            End If
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 10).HorizontalAlignment = xlRight
            If Not IsNull(rs!SpreadValue) Then
                .Cells(8 + x, 10) = rs!SpreadValue 'Spread (%)
            Else
                .Cells(8 + x, 10) = "0.00" 'Spread (%)
            End If
            .Cells(8 + x, 10).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 11).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_lepas_spread) Then
                .Cells(8 + x, 11) = Format(rs!harga_lepas_spread, "#,##0.00") 'Harga Selepas Spread (RM)
            Else
                .Cells(8 + x, 11) = "0.00" 'Harga Selepas Spread (RM)
            End If
            .Cells(8 + x, 11).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 12).HorizontalAlignment = xlRight
            If Not IsNull(rs!adjustment) Then
                .Cells(8 + x, 12) = Format(rs!adjustment, "#,##0.00") 'Adjustment (RM)
            Else
                .Cells(8 + x, 12) = "0.00" 'Adjustment (RM)
            End If
            .Cells(8 + x, 12).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 13).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_item) Then
                .Cells(8 + x, 13) = Format(rs!harga_item, "#,##0.00") 'Harga Belian (RM) : Tidak Campur Cukai GST
            Else
                .Cells(8 + x, 13) = "0.00" 'Harga Belian (RM) : Tidak Campur Cukai GST
            End If
            .Cells(8 + x, 13).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 14).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Jualan) Then
                .Cells(8 + x, 14) = Format(rs!Upah_Jualan, "#,##0.00") 'Upah Jualan (RM) : Pelanggan
            Else
                .Cells(8 + x, 14) = vbNullString 'Upah Jualan (RM) : Pelanggan
            End If
            .Cells(8 + x, 14).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 15).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Member) Then
                .Cells(8 + x, 15) = Format(rs!Upah_Member, "#,##0.00") 'Upah Jualan (RM) : Ahli Biasa
            Else
                .Cells(8 + x, 15) = vbNullString 'Upah Jualan (RM) : Ahli Biasa
            End If
            .Cells(8 + x, 15).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 16).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Pengedar) Then
                .Cells(8 + x, 16) = Format(rs!Upah_Pengedar, "#,##0.00") 'Upah Jualan (RM) : Silver
            Else
                .Cells(8 + x, 16) = vbNullString 'Upah Jualan (RM) : Silver
            End If
            .Cells(8 + x, 16).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 17).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_RAF) Then
                .Cells(8 + x, 17) = Format(rs!Upah_RAF, "#,##0.00") 'Upah Jualan (RM) : Gold
            Else
                .Cells(8 + x, 17) = vbNullString 'Upah Jualan (RM) : Gold
            End If
            .Cells(8 + x, 17).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 18).HorizontalAlignment = xlRight
            If Not IsNull(rs!upah_normal_dealer) Then
                .Cells(8 + x, 18) = Format(rs!upah_normal_dealer, "#,##0.00") 'Upah Jualan (RM) : Platinum
            Else
                .Cells(8 + x, 18) = vbNullString 'Upah Jualan (RM) : Platinum
            End If
            .Cells(8 + x, 18).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 19).HorizontalAlignment = xlRight
            If Not IsNull(rs!upah_master_dealer) Then
                .Cells(8 + x, 19) = Format(rs!upah_master_dealer, "#,##0.00") 'Upah Jualan (RM) : Master Dealer
            Else
                .Cells(8 + x, 19) = vbNullString 'Upah Jualan (RM) : Master Dealer
            End If
            .Cells(8 + x, 19).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 20).HorizontalAlignment = xlRight
            If Not IsNull(rs!code_Supplier) Then
                .Cells(8 + x, 20) = Format(rs!code_Supplier, "#,##0.00") 'Tetapan Harga Jualan (RM) : Pelanggan
            Else
                .Cells(8 + x, 20) = vbNullString 'Tetapan Harga Jualan (RM) : Pelanggan
            End If
            .Cells(8 + x, 20).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 21).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_Member) Then
                .Cells(8 + x, 21) = Format(rs!HargaJualan_Member, "#,##0.00") 'Tetapan Harga Jualan (RM) : Member
            Else
                .Cells(8 + x, 21) = vbNullString 'Tetapan Harga Jualan (RM) : Member
            End If
            .Cells(8 + x, 21).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 22).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_Pengedar) Then
                .Cells(8 + x, 22) = Format(rs!HargaJualan_Pengedar, "#,##0.00") 'Tetapan Harga Jualan (RM) : Pengedar
            Else
                .Cells(8 + x, 22) = vbNullString 'Tetapan Harga Jualan (RM) : Pengedar
            End If
            .Cells(8 + x, 22).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 23).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_RAF) Then
                .Cells(8 + x, 23) = Format(rs!HargaJualan_RAF, "#,##0.00") 'Tetapan Harga Jualan (RM) : RAF
            Else
                .Cells(8 + x, 23) = vbNullString 'Tetapan Harga Jualan (RM) : RAF
            End If
            .Cells(8 + x, 23).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 24).HorizontalAlignment = xlRight
            If Not IsNull(rs!hargajualan_normal_dealer) Then
                .Cells(8 + x, 24) = Format(rs!hargajualan_normal_dealer, "#,##0.00") 'Tetapan Harga Jualan (RM) : Normal Dealer
            Else
                .Cells(8 + x, 24) = vbNullString 'Tetapan Harga Jualan (RM) : Normal Dealer
            End If
            .Cells(8 + x, 24).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 25).HorizontalAlignment = xlRight
            If Not IsNull(rs!hargajualan_master_dealer) Then
                .Cells(8 + x, 25) = Format(rs!hargajualan_master_dealer, "#,##0.00") 'Tetapan Harga Jualan (RM) : Master Dealer
            Else
                .Cells(8 + x, 25) = vbNullString 'Tetapan Harga Jualan (RM) : Master Dealer
            End If
            .Cells(8 + x, 25).NumberFormat = "#,##0.00"
            
            If Not IsNull(rs!dulang) Then
                .Cells(8 + x, 26) = rs!dulang 'Dulang
                .Cells(8 + x, 26).HorizontalAlignment = xlCenter
            End If
                
            If Not IsNull(rs!dimension_Panjang) Then
                .Cells(8 + x, 27) = rs!dimension_Panjang 'Panjang
                .Cells(8 + x, 27).HorizontalAlignment = xlCenter
            End If
    
            If Not IsNull(rs!dimension_Lebar) Then
                .Cells(8 + x, 28) = rs!dimension_Lebar 'Lebar
                .Cells(8 + x, 28).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!dimension_Saiz) Then
                .Cells(8 + x, 29) = rs!dimension_Saiz 'Tebal
                .Cells(8 + x, 29).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!bill_No_Belian) Then .Cells(8 + x, 30) = "'" & rs!bill_No_Belian 'No Invoice Dari Supplier
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!code1) Then .Cells(8 + x, 31) = "'" & rs!code1 'No Invoice Dari Supplier
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!code2) Then .Cells(8 + x, 32) = "'" & rs!code2 'No Invoice Dari Supplier
            
            If Not IsNull(rs!cawangan) Then .Cells(8 + x, 33) = rs!cawangan 'Cawangan
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            For Col = 1 To 33
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 2
        .Cells(8 + Y, 1) = "Bilangan : " & x 'Total Barang
        Y = Y + 1
        
        '#### Jumlah Berat Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select SUM(Berat) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND (receiving_Status='" & 1 & "' OR receiving_Status='" & 0 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select SUM(Berat) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 1 & "' OR receiving_Status='" & 0 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        '#### Jumlah Berat Keseluruhan #### - End
        
        Y = Y + 1
        
        '#### Jumlah Modal Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select SUM(harga_item) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND (receiving_Status='" & 1 & "' OR receiving_Status='" & 0 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select SUM(harga_item) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 1 & "' OR receiving_Status='" & 0 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "0.00"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "RM 0.00"
        End If
        '#### Jumlah Modal Keseluruhan #### - End
        
        Y = Y + 4
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
End Sub
Sub Frm85_overall_beli_berat()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    
    x = 0
    
    If Frm101.L46_Text = "Semua cawangan" Then
    
        Frm85_LM_SEARCH_5 = Null
        Frm85_LM_SEARCH_5_LOGIC = "<>"
        Frm85_LM_SEARCH_6 = Null
        Frm85_LM_SEARCH_6_LOGIC = "<>"
        
    Else
    
        Frm85_LM_SEARCH_5 = Frm101.L46_Text
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

    Frm85_LM_BERAT = 0
    Frm85_LM_HARGA = 0
    
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh Belian
        .Columns("C").ColumnWidth = 20 'No. Siri Produk
        .Columns("D").ColumnWidth = 20 'Purity
        .Columns("E").ColumnWidth = 40 'Kategori Produk
        .Columns("F").ColumnWidth = 40 'Supplier
        .Columns("G").ColumnWidth = 20 'Berat (g)
        .Columns("H").ColumnWidth = 20 'Rate Penerimaan (RM/g)
        .Columns("I").ColumnWidth = 20 'Upah (RM)
        .Columns("J").ColumnWidth = 20 'Spread (%)
        .Columns("K").ColumnWidth = 20 'Harga Selepas Spread (RM)
        .Columns("L").ColumnWidth = 20 'Adjustment (RM)
        .Columns("M").ColumnWidth = 20 'Harga Belian (RM)
        .Columns("N").ColumnWidth = 20 'Upah Jualan (RM) : Pelanggan
        .Columns("O").ColumnWidth = 20 'Upah Jualan (RM) : Ahli Biasa
        .Columns("P").ColumnWidth = 20 'Upah Jualan (RM) : Silver
        .Columns("Q").ColumnWidth = 20 'Upah Jualan (RM) : Gold
        .Columns("R").ColumnWidth = 20 'Upah Jualan (RM) : Platinum
        .Columns("S").ColumnWidth = 0 'Upah Jualan (RM) : Master Dealer
        .Columns("T").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Pelanggan
        .Columns("U").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Member
        .Columns("V").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Silver
        .Columns("W").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Gold
        .Columns("X").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Platinum
        .Columns("Y").ColumnWidth = 0 'Tetapan Harga Jualan (RM) : Master Dealer
        .Columns("Z").ColumnWidth = 20 'Dulang
        .Columns("AA").ColumnWidth = 20 'Panjang
        .Columns("AB").ColumnWidth = 20 'Lebar
        .Columns("AC").ColumnWidth = 20 'Saiz
        .Columns("AD").ColumnWidth = 20 'No. Invoice Dari Supplier
        .Columns("AE").ColumnWidth = 20 'Code 1
        .Columns("AF").ColumnWidth = 20 'Code 2
        .Columns("AG").ColumnWidth = 40 'Cawangan
        
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        '#### Header Report ###
        .Cells(7, 1) = Frm85.L10_Text 'Report Header"
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Belian"
        .Cells(8, 3) = "No. Siri Produk"
        .Cells(8, 4) = "Purity"
        .Cells(8, 5) = "Kategori Produk"
        .Cells(8, 6) = "Supplier"
        .Cells(8, 7) = "Berat (g)"
        .Cells(8, 8) = "Rate Penerimaan (RM/g)"
        .Cells(8, 9) = "Upah (RM)"
        .Cells(8, 10) = "Spread (%)"
        .Cells(8, 11) = "Harga Selepas Spread (RM)"
        .Cells(8, 12) = "Adjustment (RM)"
        .Cells(8, 13) = "Harga Belian Termasuk GST (RM)"
        .Cells(8, 14) = "Upah Jualan (RM) : Pelanggan"
        .Cells(8, 15) = "Upah Jualan (RM) : Ahli"
        .Cells(8, 16) = "Upah Jualan (RM) : Silver"
        .Cells(8, 17) = "Upah Jualan (RM) : Gold"
        .Cells(8, 18) = "Upah Jualan (RM) : Platinum"
        .Cells(8, 19) = "Upah Jualan (RM) : Master Dealer"
        .Cells(8, 20) = "Tetapan Harga Jualan (RM) : Pelanggan"
        .Cells(8, 21) = "Tetapan Harga Jualan (RM) : Ahli"
        .Cells(8, 22) = "Tetapan Harga Jualan (RM) : Silver"
        .Cells(8, 23) = "Tetapan Harga Jualan (RM) : Gold"
        .Cells(8, 24) = "Tetapan Harga Jualan (RM) : Platinum"
        .Cells(8, 25) = "Tetapan Harga Jualan (RM) : Master Dealer"
        .Cells(8, 26) = "Dulang"
        .Cells(8, 27) = "Panjang"
        .Cells(8, 28) = "Lebar"
        .Cells(8, 29) = "Saiz"
        .Cells(8, 30) = "No. Invoice"
        .Cells(8, 31) = "Code 1"
        .Cells(8, 32) = "Code 2"
        .Cells(8, 33) = "Cawangan"
        
        For i = 1 To 33
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND Berat='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
    
        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh_belian) Then .Cells(8 + x, 2) = "'" & rs!tarikh_belian 'Tarikh Belian
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 3) = rs!no_siri_Produk 'No. Siri Produk
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!purity) Then .Cells(8 + x, 4) = rs!purity 'Purity
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 5) = rs!kategori_Produk 'Kategori Produk
            
            If Not IsNull(rs!nama_Supplier) Then .Cells(8 + x, 6) = rs!nama_Supplier 'Nama Supplier
            
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!Berat) Then
                .Cells(8 + x, 7) = Format(rs!Berat, "#,##0.00") 'Berat (g)
            Else
                .Cells(8 + x, 7) = "0.00" 'Berat (g)
            End If
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!kos_Belian_Gram) Then
                .Cells(8 + x, 8) = Format(rs!kos_Belian_Gram, "#,##0.00") 'Rate Penerimaan (RM/g)
            Else
                .Cells(8 + x, 8) = "0.00" 'Rate Penerimaan (RM/g)
            End If
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!UPAH) Then
                .Cells(8 + x, 9) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
            Else
                .Cells(8 + x, 9) = "0.00" 'Upah (RM)
            End If
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 10).HorizontalAlignment = xlRight
            If Not IsNull(rs!SpreadValue) Then
                .Cells(8 + x, 10) = rs!SpreadValue 'Spread (%)
            Else
                .Cells(8 + x, 10) = "0.00" 'Spread (%)
            End If
            .Cells(8 + x, 10).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 11).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_lepas_spread) Then
                .Cells(8 + x, 11) = Format(rs!harga_lepas_spread, "#,##0.00") 'Harga Selepas Spread (RM)
            Else
                .Cells(8 + x, 11) = "0.00" 'Harga Selepas Spread (RM)
            End If
            .Cells(8 + x, 11).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 12).HorizontalAlignment = xlRight
            If Not IsNull(rs!adjustment) Then
                .Cells(8 + x, 12) = Format(rs!adjustment, "#,##0.00") 'Adjustment (RM)
            Else
                .Cells(8 + x, 12) = "0.00" 'Adjustment (RM)
            End If
            .Cells(8 + x, 12).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 13).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_item) Then
                .Cells(8 + x, 13) = Format(rs!harga_item, "#,##0.00") 'Harga Belian (RM) : Tidak Campur Cukai GST
            Else
                .Cells(8 + x, 13) = "0.00" 'Harga Belian (RM) : Tidak Campur Cukai GST
            End If
            .Cells(8 + x, 13).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 14).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Jualan) Then
                .Cells(8 + x, 14) = Format(rs!Upah_Jualan, "#,##0.00") 'Upah Jualan (RM) : Pelanggan
            Else
                .Cells(8 + x, 14) = vbNullString 'Upah Jualan (RM) : Pelanggan
            End If
            .Cells(8 + x, 14).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 15).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Member) Then
                .Cells(8 + x, 15) = Format(rs!Upah_Member, "#,##0.00") 'Upah Jualan (RM) : Ahli Biasa
            Else
                .Cells(8 + x, 15) = vbNullString 'Upah Jualan (RM) : Ahli Biasa
            End If
            .Cells(8 + x, 15).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 16).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Pengedar) Then
                .Cells(8 + x, 16) = Format(rs!Upah_Pengedar, "#,##0.00") 'Upah Jualan (RM) : Silver
            Else
                .Cells(8 + x, 16) = vbNullString 'Upah Jualan (RM) : Silver
            End If
            .Cells(8 + x, 16).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 17).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_RAF) Then
                .Cells(8 + x, 17) = Format(rs!Upah_RAF, "#,##0.00") 'Upah Jualan (RM) : Gold
            Else
                .Cells(8 + x, 17) = vbNullString 'Upah Jualan (RM) : Gold
            End If
            .Cells(8 + x, 17).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 18).HorizontalAlignment = xlRight
            If Not IsNull(rs!upah_normal_dealer) Then
                .Cells(8 + x, 18) = Format(rs!upah_normal_dealer, "#,##0.00") 'Upah Jualan (RM) : Platinum
            Else
                .Cells(8 + x, 18) = vbNullString 'Upah Jualan (RM) : Platinum
            End If
            .Cells(8 + x, 18).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 19).HorizontalAlignment = xlRight
            If Not IsNull(rs!upah_master_dealer) Then
                .Cells(8 + x, 19) = Format(rs!upah_master_dealer, "#,##0.00") 'Upah Jualan (RM) : Master Dealer
            Else
                .Cells(8 + x, 19) = vbNullString 'Upah Jualan (RM) : Master Dealer
            End If
            .Cells(8 + x, 19).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 20).HorizontalAlignment = xlRight
            If Not IsNull(rs!code_Supplier) Then
                .Cells(8 + x, 20) = Format(rs!code_Supplier, "#,##0.00") 'Tetapan Harga Jualan (RM) : Pelanggan
            Else
                .Cells(8 + x, 20) = vbNullString 'Tetapan Harga Jualan (RM) : Pelanggan
            End If
            .Cells(8 + x, 20).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 21).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_Member) Then
                .Cells(8 + x, 21) = Format(rs!HargaJualan_Member, "#,##0.00") 'Tetapan Harga Jualan (RM) : Member
            Else
                .Cells(8 + x, 21) = vbNullString 'Tetapan Harga Jualan (RM) : Member
            End If
            .Cells(8 + x, 21).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 22).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_Pengedar) Then
                .Cells(8 + x, 22) = Format(rs!HargaJualan_Pengedar, "#,##0.00") 'Tetapan Harga Jualan (RM) : Pengedar
            Else
                .Cells(8 + x, 22) = vbNullString 'Tetapan Harga Jualan (RM) : Pengedar
            End If
            .Cells(8 + x, 22).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 23).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_RAF) Then
                .Cells(8 + x, 23) = Format(rs!HargaJualan_RAF, "#,##0.00") 'Tetapan Harga Jualan (RM) : RAF
            Else
                .Cells(8 + x, 23) = vbNullString 'Tetapan Harga Jualan (RM) : RAF
            End If
            .Cells(8 + x, 23).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 24).HorizontalAlignment = xlRight
            If Not IsNull(rs!hargajualan_normal_dealer) Then
                .Cells(8 + x, 24) = Format(rs!hargajualan_normal_dealer, "#,##0.00") 'Tetapan Harga Jualan (RM) : Normal Dealer
            Else
                .Cells(8 + x, 24) = vbNullString 'Tetapan Harga Jualan (RM) : Normal Dealer
            End If
            .Cells(8 + x, 24).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 25).HorizontalAlignment = xlRight
            If Not IsNull(rs!hargajualan_master_dealer) Then
                .Cells(8 + x, 25) = Format(rs!hargajualan_master_dealer, "#,##0.00") 'Tetapan Harga Jualan (RM) : Master Dealer
            Else
                .Cells(8 + x, 25) = vbNullString 'Tetapan Harga Jualan (RM) : Master Dealer
            End If
            .Cells(8 + x, 25).NumberFormat = "#,##0.00"
            
            If Not IsNull(rs!dulang) Then
                .Cells(8 + x, 26) = rs!dulang 'Dulang
                .Cells(8 + x, 26).HorizontalAlignment = xlCenter
            End If
                
            If Not IsNull(rs!dimension_Panjang) Then
                .Cells(8 + x, 27) = rs!dimension_Panjang 'Panjang
                .Cells(8 + x, 27).HorizontalAlignment = xlCenter
            End If
    
            If Not IsNull(rs!dimension_Lebar) Then
                .Cells(8 + x, 28) = rs!dimension_Lebar 'Lebar
                .Cells(8 + x, 28).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!dimension_Saiz) Then
                .Cells(8 + x, 29) = rs!dimension_Saiz 'Tebal
                .Cells(8 + x, 29).HorizontalAlignment = xlCenter
            End If

            If Not IsNull(rs!bill_No_Belian) Then .Cells(8 + x, 30) = "'" & rs!bill_No_Belian 'No Invoice Dari Supplier
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!code1) Then .Cells(8 + x, 31) = "'" & rs!code1 'No Invoice Dari Supplier
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!code2) Then .Cells(8 + x, 32) = "'" & rs!code2 'No Invoice Dari Supplier
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!cawangan) Then .Cells(8 + x, 33) = rs!cawangan 'Cawangan
            
            For Col = 1 To 33
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 2
        .Cells(8 + Y, 1) = "Bilangan : " & x 'Total Barang
        Y = Y + 1
        
        '#### Jumlah Berat Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(Berat) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND Berat='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        '#### Jumlah Berat Keseluruhan #### - End
        
        Y = Y + 1
        
        '#### Jumlah Modal Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(harga_item) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND Berat='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "0.00"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "RM 0.00"
        End If
        '#### Jumlah Modal Keseluruhan #### - End
        
        Y = Y + 4
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
End Sub
Sub Frm85_overall_beli_no_siri()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    
    x = 0
    
    Frm85_LM_BERAT = 0
    Frm85_LM_HARGA = 0
    
    If Frm101.L46_Text = "Semua cawangan" Then
    
        Frm85_LM_SEARCH_5 = Null
        Frm85_LM_SEARCH_5_LOGIC = "<>"
        Frm85_LM_SEARCH_6 = Null
        Frm85_LM_SEARCH_6_LOGIC = "<>"
        
    Else
    
        Frm85_LM_SEARCH_5 = Frm101.L46_Text
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
    
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh Belian
        .Columns("C").ColumnWidth = 20 'No. Siri Produk
        .Columns("D").ColumnWidth = 20 'Purity
        .Columns("E").ColumnWidth = 40 'Kategori Produk
        .Columns("F").ColumnWidth = 40 'Supplier
        .Columns("G").ColumnWidth = 20 'Berat (g)
        .Columns("H").ColumnWidth = 20 'Rate Penerimaan (RM/g)
        .Columns("I").ColumnWidth = 20 'Upah (RM)
        .Columns("J").ColumnWidth = 20 'Spread (%)
        .Columns("K").ColumnWidth = 20 'Harga Selepas Spread (RM)
        .Columns("L").ColumnWidth = 20 'Adjustment (RM)
        .Columns("M").ColumnWidth = 20 'Harga Belian (RM)
        .Columns("N").ColumnWidth = 20 'Upah Jualan (RM) : Pelanggan
        .Columns("O").ColumnWidth = 20 'Upah Jualan (RM) : Ahli Biasa
        .Columns("P").ColumnWidth = 20 'Upah Jualan (RM) : Silver
        .Columns("Q").ColumnWidth = 20 'Upah Jualan (RM) : Gold
        .Columns("R").ColumnWidth = 20 'Upah Jualan (RM) : Platinum
        .Columns("S").ColumnWidth = 0 'Upah Jualan (RM) : Master Dealer
        .Columns("T").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Pelanggan
        .Columns("U").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Member
        .Columns("V").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Silver
        .Columns("W").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Gold
        .Columns("X").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Platinum
        .Columns("Y").ColumnWidth = 0 'Tetapan Harga Jualan (RM) : Master Dealer
        .Columns("Z").ColumnWidth = 20 'Dulang
        .Columns("AA").ColumnWidth = 20 'Panjang
        .Columns("AB").ColumnWidth = 20 'Lebar
        .Columns("AC").ColumnWidth = 20 'Saiz
        .Columns("AD").ColumnWidth = 20 'No. Invoice Dari Supplier
        .Columns("AE").ColumnWidth = 20 'Code 1
        .Columns("AF").ColumnWidth = 20 'Code 2
        .Columns("AG").ColumnWidth = 40 'Cawangan
        
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        '#### Header Report ###
        .Cells(7, 1) = Frm85.L10_Text 'Report Header"
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Belian"
        .Cells(8, 3) = "No. Siri Produk"
        .Cells(8, 4) = "Purity"
        .Cells(8, 5) = "Kategori Produk"
        .Cells(8, 6) = "Supplier"
        .Cells(8, 7) = "Berat (g)"
        .Cells(8, 8) = "Rate Penerimaan (RM/g)"
        .Cells(8, 9) = "Upah (RM)"
        .Cells(8, 10) = "Spread (%)"
        .Cells(8, 11) = "Harga Selepas Spread (RM)"
        .Cells(8, 12) = "Adjustment (RM)"
        .Cells(8, 13) = "Harga Belian Termasuk GST (RM)"
        .Cells(8, 14) = "Upah Jualan (RM) : Pelanggan"
        .Cells(8, 15) = "Upah Jualan (RM) : Ahli"
        .Cells(8, 16) = "Upah Jualan (RM) : Silver"
        .Cells(8, 17) = "Upah Jualan (RM) : Gold"
        .Cells(8, 18) = "Upah Jualan (RM) : Platinum"
        .Cells(8, 19) = "Upah Jualan (RM) : Master Dealer"
        .Cells(8, 20) = "Tetapan Harga Jualan (RM) : Pelanggan"
        .Cells(8, 21) = "Tetapan Harga Jualan (RM) : Ahli"
        .Cells(8, 22) = "Tetapan Harga Jualan (RM) : Silver"
        .Cells(8, 23) = "Tetapan Harga Jualan (RM) : Gold"
        .Cells(8, 24) = "Tetapan Harga Jualan (RM) : Platinum"
        .Cells(8, 25) = "Tetapan Harga Jualan (RM) : Master Dealer"
        .Cells(8, 26) = "Dulang"
        .Cells(8, 27) = "Panjang"
        .Cells(8, 28) = "Lebar"
        .Cells(8, 29) = "Saiz"
        .Cells(8, 30) = "No. Invoice"
        .Cells(8, 31) = "Code 1"
        .Cells(8, 32) = "Code 2"
        .Cells(8, 33) = "Cawangan"
        
        For i = 1 To 33
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 AND (receiving_Status='" & 1 & "' OR receiving_Status='" & 0 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh_belian) Then .Cells(8 + x, 2) = "'" & rs!tarikh_belian 'Tarikh Belian
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 3) = rs!no_siri_Produk 'No. Siri Produk
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!purity) Then .Cells(8 + x, 4) = rs!purity 'Purity
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 5) = rs!kategori_Produk 'Kategori Produk
            
            If Not IsNull(rs!nama_Supplier) Then .Cells(8 + x, 6) = rs!nama_Supplier 'Nama Supplier
            
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!Berat) Then
                .Cells(8 + x, 7) = Format(rs!Berat, "#,##0.00") 'Berat (g)
            Else
                .Cells(8 + x, 7) = "0.00" 'Berat (g)
            End If
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!kos_Belian_Gram) Then
                .Cells(8 + x, 8) = Format(rs!kos_Belian_Gram, "#,##0.00") 'Rate Penerimaan (RM/g)
            Else
                .Cells(8 + x, 8) = "0.00" 'Rate Penerimaan (RM/g)
            End If
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!UPAH) Then
                .Cells(8 + x, 9) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
            Else
                .Cells(8 + x, 9) = "0.00" 'Upah (RM)
            End If
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 10).HorizontalAlignment = xlRight
            If Not IsNull(rs!SpreadValue) Then
                .Cells(8 + x, 10) = rs!SpreadValue 'Spread (%)
            Else
                .Cells(8 + x, 10) = "0.00" 'Spread (%)
            End If
            .Cells(8 + x, 10).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 11).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_lepas_spread) Then
                .Cells(8 + x, 11) = Format(rs!harga_lepas_spread, "#,##0.00") 'Harga Selepas Spread (RM)
            Else
                .Cells(8 + x, 11) = "0.00" 'Harga Selepas Spread (RM)
            End If
            .Cells(8 + x, 11).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 12).HorizontalAlignment = xlRight
            If Not IsNull(rs!adjustment) Then
                .Cells(8 + x, 12) = Format(rs!adjustment, "#,##0.00") 'Adjustment (RM)
            Else
                .Cells(8 + x, 12) = "0.00" 'Adjustment (RM)
            End If
            .Cells(8 + x, 12).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 13).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_item) Then
                .Cells(8 + x, 13) = Format(rs!harga_item, "#,##0.00") 'Harga Belian (RM) : Tidak Campur Cukai GST
            Else
                .Cells(8 + x, 13) = "0.00" 'Harga Belian (RM) : Tidak Campur Cukai GST
            End If
            .Cells(8 + x, 13).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 14).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Jualan) Then
                .Cells(8 + x, 14) = Format(rs!Upah_Jualan, "#,##0.00") 'Upah Jualan (RM) : Pelanggan
            Else
                .Cells(8 + x, 14) = vbNullString 'Upah Jualan (RM) : Pelanggan
            End If
            .Cells(8 + x, 14).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 15).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Member) Then
                .Cells(8 + x, 15) = Format(rs!Upah_Member, "#,##0.00") 'Upah Jualan (RM) : Ahli Biasa
            Else
                .Cells(8 + x, 15) = vbNullString 'Upah Jualan (RM) : Ahli Biasa
            End If
            .Cells(8 + x, 15).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 16).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Pengedar) Then
                .Cells(8 + x, 16) = Format(rs!Upah_Pengedar, "#,##0.00") 'Upah Jualan (RM) : Silver
            Else
                .Cells(8 + x, 16) = vbNullString 'Upah Jualan (RM) : Silver
            End If
            .Cells(8 + x, 16).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 17).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_RAF) Then
                .Cells(8 + x, 17) = Format(rs!Upah_RAF, "#,##0.00") 'Upah Jualan (RM) : Gold
            Else
                .Cells(8 + x, 17) = vbNullString 'Upah Jualan (RM) : Gold
            End If
            .Cells(8 + x, 17).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 18).HorizontalAlignment = xlRight
            If Not IsNull(rs!upah_normal_dealer) Then
                .Cells(8 + x, 18) = Format(rs!upah_normal_dealer, "#,##0.00") 'Upah Jualan (RM) : Platinum
            Else
                .Cells(8 + x, 18) = vbNullString 'Upah Jualan (RM) : Platinum
            End If
            .Cells(8 + x, 18).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 19).HorizontalAlignment = xlRight
            If Not IsNull(rs!upah_master_dealer) Then
                .Cells(8 + x, 19) = Format(rs!upah_master_dealer, "#,##0.00") 'Upah Jualan (RM) : Master Dealer
            Else
                .Cells(8 + x, 19) = vbNullString 'Upah Jualan (RM) : Master Dealer
            End If
            .Cells(8 + x, 19).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 20).HorizontalAlignment = xlRight
            If Not IsNull(rs!code_Supplier) Then
                .Cells(8 + x, 20) = Format(rs!code_Supplier, "#,##0.00") 'Tetapan Harga Jualan (RM) : Pelanggan
            Else
                .Cells(8 + x, 20) = vbNullString 'Tetapan Harga Jualan (RM) : Pelanggan
            End If
            .Cells(8 + x, 20).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 21).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_Member) Then
                .Cells(8 + x, 21) = Format(rs!HargaJualan_Member, "#,##0.00") 'Tetapan Harga Jualan (RM) : Member
            Else
                .Cells(8 + x, 21) = vbNullString 'Tetapan Harga Jualan (RM) : Member
            End If
            .Cells(8 + x, 21).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 22).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_Pengedar) Then
                .Cells(8 + x, 22) = Format(rs!HargaJualan_Pengedar, "#,##0.00") 'Tetapan Harga Jualan (RM) : Pengedar
            Else
                .Cells(8 + x, 22) = vbNullString 'Tetapan Harga Jualan (RM) : Pengedar
            End If
            .Cells(8 + x, 22).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 23).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_RAF) Then
                .Cells(8 + x, 23) = Format(rs!HargaJualan_RAF, "#,##0.00") 'Tetapan Harga Jualan (RM) : RAF
            Else
                .Cells(8 + x, 23) = vbNullString 'Tetapan Harga Jualan (RM) : RAF
            End If
            .Cells(8 + x, 23).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 24).HorizontalAlignment = xlRight
            If Not IsNull(rs!hargajualan_normal_dealer) Then
                .Cells(8 + x, 24) = Format(rs!hargajualan_normal_dealer, "#,##0.00") 'Tetapan Harga Jualan (RM) : Normal Dealer
            Else
                .Cells(8 + x, 24) = vbNullString 'Tetapan Harga Jualan (RM) : Normal Dealer
            End If
            .Cells(8 + x, 24).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 25).HorizontalAlignment = xlRight
            If Not IsNull(rs!hargajualan_master_dealer) Then
                .Cells(8 + x, 25) = Format(rs!hargajualan_master_dealer, "#,##0.00") 'Tetapan Harga Jualan (RM) : Master Dealer
            Else
                .Cells(8 + x, 25) = vbNullString 'Tetapan Harga Jualan (RM) : Master Dealer
            End If
            .Cells(8 + x, 25).NumberFormat = "#,##0.00"
            
            If Not IsNull(rs!dulang) Then
                .Cells(8 + x, 26) = rs!dulang 'Dulang
                .Cells(8 + x, 26).HorizontalAlignment = xlCenter
            End If
                
            If Not IsNull(rs!dimension_Panjang) Then
                .Cells(8 + x, 27) = rs!dimension_Panjang 'Panjang
                .Cells(8 + x, 27).HorizontalAlignment = xlCenter
            End If
    
            If Not IsNull(rs!dimension_Lebar) Then
                .Cells(8 + x, 28) = rs!dimension_Lebar 'Lebar
                .Cells(8 + x, 28).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!dimension_Saiz) Then
                .Cells(8 + x, 29) = rs!dimension_Saiz 'Tebal
                .Cells(8 + x, 29).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!bill_No_Belian) Then .Cells(8 + x, 30) = "'" & rs!bill_No_Belian 'No Invoice Dari Supplier
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!code1) Then .Cells(8 + x, 31) = "'" & rs!code1 'No Invoice Dari Supplier
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!code2) Then .Cells(8 + x, 32) = "'" & rs!code2 'No Invoice Dari Supplier
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!cawangan) Then .Cells(8 + x, 33) = rs!cawangan 'Cawangan
            
            For Col = 1 To 33
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 2
        .Cells(8 + Y, 1) = "Bilangan : " & x 'Total Barang
        Y = Y + 1
        
        '#### Jumlah Berat Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(Berat) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 AND (receiving_Status='" & 1 & "' OR receiving_Status='" & 0 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        '#### Jumlah Berat Keseluruhan #### - End
        
        Y = Y + 1
        
        '#### Jumlah Modal Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(harga_item) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' AND StatusItem <> 0 AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND (receiving_Status='" & 1 & "' OR receiving_Status='" & 0 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "0.00"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "RM 0.00"
        End If
        '#### Jumlah Modal Keseluruhan #### - End
        
        Y = Y + 4
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
End Sub
Sub Frm85_overall_beli_invoice_supplier()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    
    x = 0
    
    Frm85_LM_BERAT = 0
    Frm85_LM_HARGA = 0
    
    If Frm101.L46_Text = "Semua cawangan" Then
    
        Frm85_LM_SEARCH_5 = Null
        Frm85_LM_SEARCH_5_LOGIC = "<>"
        Frm85_LM_SEARCH_6 = Null
        Frm85_LM_SEARCH_6_LOGIC = "<>"
        
    Else
    
        Frm85_LM_SEARCH_5 = Frm101.L46_Text
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

    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh Belian
        .Columns("C").ColumnWidth = 20 'No. Siri Produk
        .Columns("D").ColumnWidth = 20 'Purity
        .Columns("E").ColumnWidth = 40 'Kategori Produk
        .Columns("F").ColumnWidth = 40 'Supplier
        .Columns("G").ColumnWidth = 20 'Berat (g)
        .Columns("H").ColumnWidth = 20 'Rate Penerimaan (RM/g)
        .Columns("I").ColumnWidth = 20 'Upah (RM)
        .Columns("J").ColumnWidth = 20 'Spread (%)
        .Columns("K").ColumnWidth = 20 'Harga Selepas Spread (RM)
        .Columns("L").ColumnWidth = 20 'Adjustment (RM)
        .Columns("M").ColumnWidth = 20 'Harga Belian (RM)
        .Columns("N").ColumnWidth = 20 'Upah Jualan (RM) : Pelanggan
        .Columns("O").ColumnWidth = 20 'Upah Jualan (RM) : Ahli Biasa
        .Columns("P").ColumnWidth = 20 'Upah Jualan (RM) : Silver
        .Columns("Q").ColumnWidth = 20 'Upah Jualan (RM) : Gold
        .Columns("R").ColumnWidth = 20 'Upah Jualan (RM) : Platinum
        .Columns("S").ColumnWidth = 0 'Upah Jualan (RM) : Master Dealer
        .Columns("T").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Pelanggan
        .Columns("U").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Member
        .Columns("V").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Silver
        .Columns("W").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Gold
        .Columns("X").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Platinum
        .Columns("Y").ColumnWidth = 0 'Tetapan Harga Jualan (RM) : Master Dealer
        .Columns("Z").ColumnWidth = 20 'Dulang
        .Columns("AA").ColumnWidth = 20 'Panjang
        .Columns("AB").ColumnWidth = 20 'Lebar
        .Columns("AC").ColumnWidth = 20 'Saiz
        .Columns("AD").ColumnWidth = 20 'No. Invoice Dari Supplier
        .Columns("AE").ColumnWidth = 20 'Code 1
        .Columns("AF").ColumnWidth = 20 'Code 2
        .Columns("AG").ColumnWidth = 40 'Cawangan
        
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        '#### Header Report ###
        .Cells(7, 1) = Frm85.L10_Text 'Report Header"
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Belian"
        .Cells(8, 3) = "No. Siri Produk"
        .Cells(8, 4) = "Purity"
        .Cells(8, 5) = "Kategori Produk"
        .Cells(8, 6) = "Supplier"
        .Cells(8, 7) = "Berat (g)"
        .Cells(8, 8) = "Rate Penerimaan (RM/g)"
        .Cells(8, 9) = "Upah (RM)"
        .Cells(8, 10) = "Spread (%)"
        .Cells(8, 11) = "Harga Selepas Spread (RM)"
        .Cells(8, 12) = "Adjustment (RM)"
        .Cells(8, 13) = "Harga Belian Termasuk GST (RM)"
        .Cells(8, 14) = "Upah Jualan (RM) : Pelanggan"
        .Cells(8, 15) = "Upah Jualan (RM) : Ahli"
        .Cells(8, 16) = "Upah Jualan (RM) : Silver"
        .Cells(8, 17) = "Upah Jualan (RM) : Gold"
        .Cells(8, 18) = "Upah Jualan (RM) : Platinum"
        .Cells(8, 19) = "Upah Jualan (RM) : Master Dealer"
        .Cells(8, 20) = "Tetapan Harga Jualan (RM) : Pelanggan"
        .Cells(8, 21) = "Tetapan Harga Jualan (RM) : Ahli"
        .Cells(8, 22) = "Tetapan Harga Jualan (RM) : Silver"
        .Cells(8, 23) = "Tetapan Harga Jualan (RM) : Gold"
        .Cells(8, 24) = "Tetapan Harga Jualan (RM) : Platinum"
        .Cells(8, 25) = "Tetapan Harga Jualan (RM) : Master Dealer"
        .Cells(8, 26) = "Dulang"
        .Cells(8, 27) = "Panjang"
        .Cells(8, 28) = "Lebar"
        .Cells(8, 29) = "Saiz"
        .Cells(8, 30) = "No. Invoice"
        .Cells(8, 31) = "Code 1"
        .Cells(8, 32) = "Code 2"
        .Cells(8, 33) = "Cawangan"
        
        For i = 1 To 33
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND bill_No_Belian='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
    
        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh_belian) Then .Cells(8 + x, 2) = "'" & rs!tarikh_belian 'Tarikh Belian
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 3) = rs!no_siri_Produk 'No. Siri Produk
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!purity) Then .Cells(8 + x, 4) = rs!purity 'Purity
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 5) = rs!kategori_Produk 'Kategori Produk
            
            If Not IsNull(rs!nama_Supplier) Then .Cells(8 + x, 6) = rs!nama_Supplier 'Nama Supplier
            
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!Berat) Then
                .Cells(8 + x, 7) = Format(rs!Berat, "#,##0.00") 'Berat (g)
            Else
                .Cells(8 + x, 7) = "0.00" 'Berat (g)
            End If
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!kos_Belian_Gram) Then
                .Cells(8 + x, 8) = Format(rs!kos_Belian_Gram, "#,##0.00") 'Rate Penerimaan (RM/g)
            Else
                .Cells(8 + x, 8) = "0.00" 'Rate Penerimaan (RM/g)
            End If
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!UPAH) Then
                .Cells(8 + x, 9) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
            Else
                .Cells(8 + x, 9) = "0.00" 'Upah (RM)
            End If
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 10).HorizontalAlignment = xlRight
            If Not IsNull(rs!SpreadValue) Then
                .Cells(8 + x, 10) = rs!SpreadValue 'Spread (%)
            Else
                .Cells(8 + x, 10) = "0.00" 'Spread (%)
            End If
            .Cells(8 + x, 10).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 11).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_lepas_spread) Then
                .Cells(8 + x, 11) = Format(rs!harga_lepas_spread, "#,##0.00") 'Harga Selepas Spread (RM)
            Else
                .Cells(8 + x, 11) = "0.00" 'Harga Selepas Spread (RM)
            End If
            .Cells(8 + x, 11).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 12).HorizontalAlignment = xlRight
            If Not IsNull(rs!adjustment) Then
                .Cells(8 + x, 12) = Format(rs!adjustment, "#,##0.00") 'Adjustment (RM)
            Else
                .Cells(8 + x, 12) = "0.00" 'Adjustment (RM)
            End If
            .Cells(8 + x, 12).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 13).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_item) Then
                .Cells(8 + x, 13) = Format(rs!harga_item, "#,##0.00") 'Harga Belian (RM) : Tidak Campur Cukai GST
            Else
                .Cells(8 + x, 13) = "0.00" 'Harga Belian (RM) : Tidak Campur Cukai GST
            End If
            .Cells(8 + x, 13).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 14).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Jualan) Then
                .Cells(8 + x, 14) = Format(rs!Upah_Jualan, "#,##0.00") 'Upah Jualan (RM) : Pelanggan
            Else
                .Cells(8 + x, 14) = vbNullString 'Upah Jualan (RM) : Pelanggan
            End If
            .Cells(8 + x, 14).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 15).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Member) Then
                .Cells(8 + x, 15) = Format(rs!Upah_Member, "#,##0.00") 'Upah Jualan (RM) : Ahli Biasa
            Else
                .Cells(8 + x, 15) = vbNullString 'Upah Jualan (RM) : Ahli Biasa
            End If
            .Cells(8 + x, 15).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 16).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Pengedar) Then
                .Cells(8 + x, 16) = Format(rs!Upah_Pengedar, "#,##0.00") 'Upah Jualan (RM) : Silver
            Else
                .Cells(8 + x, 16) = vbNullString 'Upah Jualan (RM) : Silver
            End If
            .Cells(8 + x, 16).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 17).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_RAF) Then
                .Cells(8 + x, 17) = Format(rs!Upah_RAF, "#,##0.00") 'Upah Jualan (RM) : Gold
            Else
                .Cells(8 + x, 17) = vbNullString 'Upah Jualan (RM) : Gold
            End If
            .Cells(8 + x, 17).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 18).HorizontalAlignment = xlRight
            If Not IsNull(rs!upah_normal_dealer) Then
                .Cells(8 + x, 18) = Format(rs!upah_normal_dealer, "#,##0.00") 'Upah Jualan (RM) : Platinum
            Else
                .Cells(8 + x, 18) = vbNullString 'Upah Jualan (RM) : Platinum
            End If
            .Cells(8 + x, 18).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 19).HorizontalAlignment = xlRight
            If Not IsNull(rs!upah_master_dealer) Then
                .Cells(8 + x, 19) = Format(rs!upah_master_dealer, "#,##0.00") 'Upah Jualan (RM) : Master Dealer
            Else
                .Cells(8 + x, 19) = vbNullString 'Upah Jualan (RM) : Master Dealer
            End If
            .Cells(8 + x, 19).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 20).HorizontalAlignment = xlRight
            If Not IsNull(rs!code_Supplier) Then
                .Cells(8 + x, 20) = Format(rs!code_Supplier, "#,##0.00") 'Tetapan Harga Jualan (RM) : Pelanggan
            Else
                .Cells(8 + x, 20) = vbNullString 'Tetapan Harga Jualan (RM) : Pelanggan
            End If
            .Cells(8 + x, 20).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 21).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_Member) Then
                .Cells(8 + x, 21) = Format(rs!HargaJualan_Member, "#,##0.00") 'Tetapan Harga Jualan (RM) : Member
            Else
                .Cells(8 + x, 21) = vbNullString 'Tetapan Harga Jualan (RM) : Member
            End If
            .Cells(8 + x, 21).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 22).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_Pengedar) Then
                .Cells(8 + x, 22) = Format(rs!HargaJualan_Pengedar, "#,##0.00") 'Tetapan Harga Jualan (RM) : Pengedar
            Else
                .Cells(8 + x, 22) = vbNullString 'Tetapan Harga Jualan (RM) : Pengedar
            End If
            .Cells(8 + x, 22).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 23).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_RAF) Then
                .Cells(8 + x, 23) = Format(rs!HargaJualan_RAF, "#,##0.00") 'Tetapan Harga Jualan (RM) : RAF
            Else
                .Cells(8 + x, 23) = vbNullString 'Tetapan Harga Jualan (RM) : RAF
            End If
            .Cells(8 + x, 23).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 24).HorizontalAlignment = xlRight
            If Not IsNull(rs!hargajualan_normal_dealer) Then
                .Cells(8 + x, 24) = Format(rs!hargajualan_normal_dealer, "#,##0.00") 'Tetapan Harga Jualan (RM) : Normal Dealer
            Else
                .Cells(8 + x, 24) = vbNullString 'Tetapan Harga Jualan (RM) : Normal Dealer
            End If
            .Cells(8 + x, 24).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 25).HorizontalAlignment = xlRight
            If Not IsNull(rs!hargajualan_master_dealer) Then
                .Cells(8 + x, 25) = Format(rs!hargajualan_master_dealer, "#,##0.00") 'Tetapan Harga Jualan (RM) : Master Dealer
            Else
                .Cells(8 + x, 25) = vbNullString 'Tetapan Harga Jualan (RM) : Master Dealer
            End If
            .Cells(8 + x, 25).NumberFormat = "#,##0.00"
            
            If Not IsNull(rs!dulang) Then
                .Cells(8 + x, 26) = rs!dulang 'Dulang
                .Cells(8 + x, 26).HorizontalAlignment = xlCenter
            End If
                
            If Not IsNull(rs!dimension_Panjang) Then
                .Cells(8 + x, 27) = rs!dimension_Panjang 'Panjang
                .Cells(8 + x, 27).HorizontalAlignment = xlCenter
            End If
    
            If Not IsNull(rs!dimension_Lebar) Then
                .Cells(8 + x, 28) = rs!dimension_Lebar 'Lebar
                .Cells(8 + x, 28).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!dimension_Saiz) Then
                .Cells(8 + x, 29) = rs!dimension_Saiz 'Tebal
                .Cells(8 + x, 29).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!bill_No_Belian) Then .Cells(8 + x, 30) = "'" & rs!bill_No_Belian 'No Invoice Dari Supplier
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!code1) Then .Cells(8 + x, 31) = "'" & rs!code1 'No Invoice Dari Supplier
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!code2) Then .Cells(8 + x, 32) = "'" & rs!code2 'No Invoice Dari Supplier
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!cawangan) Then .Cells(8 + x, 33) = rs!cawangan 'Cawangan
            
            For Col = 1 To 33
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 2
        .Cells(8 + Y, 1) = "Bilangan : " & x 'Total Barang
        Y = Y + 1
        
        '#### Jumlah Berat Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(Berat) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND bill_No_Belian='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        '#### Jumlah Berat Keseluruhan #### - End
        
        Y = Y + 1
        
        '#### Jumlah Modal Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(harga_item) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND bill_No_Belian='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "0.00"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "RM 0.00"
        End If
        '#### Jumlah Modal Keseluruhan #### - End
        
        Y = Y + 4
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
End Sub
Sub Frm85_overall_jual()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then

    x = 0
    Frm85_LM_BERAT = 0
    Frm85_LM_HARGA = 0
    Frm85_LM_UNTUNG = 0
    
    If Frm101.L9_Text = 1 Then
        TM = Frm101.L7_Text 'Tarikh Mula
        TA = Frm101.L8_Text 'Tarikh Akhir
    End If
    If Frm101.L5_Text = "Semua Purity" Then
        Frm85_LM_SEARCH_1 = Null
        Frm85_LM_SEARCH_1_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_1 = Frm101.L5_Text
        Frm85_LM_SEARCH_1_LOGIC = "="
    End If
    If Frm101.L6_Text = "Semua Kategori Produk" Then
        Frm85_LM_SEARCH_2 = Null
        Frm85_LM_SEARCH_2_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_2 = Frm101.L6_Text
        Frm85_LM_SEARCH_2_LOGIC = "="
    End If
    If Frm101.L34_Text = "Semua Dulang" Then
        Frm85_LM_SEARCH_3 = Null
        Frm85_LM_SEARCH_3_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_3 = Frm101.L34_Text
        Frm85_LM_SEARCH_3_LOGIC = "="
    End If
    If Frm101.L44_Text = 2 Then '0 : Jualan biasa , 1 : Jualan kepada agen , 2 : Semua jenis jualan
        Frm85_LM_SEARCH_4 = Null
        Frm85_LM_SEARCH_4_LOGIC = "<>"
    ElseIf Frm101.L44_Text = 0 Then '0 : Jualan biasa , 1 : Jualan kepada agen , 2 : Semua jenis jualan
        Frm85_LM_SEARCH_4 = 0
        Frm85_LM_SEARCH_4_LOGIC = "="
    ElseIf Frm101.L44_Text = 1 Then '0 : Jualan biasa , 1 : Jualan kepada agen , 2 : Semua jenis jualan
        Frm85_LM_SEARCH_4 = 1
        Frm85_LM_SEARCH_4_LOGIC = "="
    End If
    If Frm101.L45_Text = "Kedai & Online" Then
        Frm85_LM_SEARCH_5 = Null
        Frm85_LM_SEARCH_5_LOGIC = "<>"
    ElseIf Frm101.L45_Text = "Kedai Sahaja" Then
        Frm85_LM_SEARCH_5 = 0
        Frm85_LM_SEARCH_5_LOGIC = "="
    ElseIf Frm101.L45_Text = "Online Sahaja" Then
        Frm85_LM_SEARCH_5 = 1
        Frm85_LM_SEARCH_5_LOGIC = "="
    End If
    
    If Frm101.L46_Text = "Semua cawangan" Then
    
        Frm85_SEARCH_8 = Null
        Frm85_SEARCH_8_LOGIC = "<>"
        Frm85_SEARCH_9 = Null
        Frm85_SEARCH_9_LOGIC = "<>"
        
    Else
    
        Frm85_SEARCH_8 = Frm101.L46_Text
        Frm85_SEARCH_8_LOGIC = "="
        Frm85_SEARCH_9 = "HQ"
        Frm85_SEARCH_9_LOGIC = "="
        
    End If
    
    user_level = MDI_frm1.L4_Text
    LM_INVOICE_RASMI = 0
    
    If user_level = "Guest/User" Then
        Frm85_LM_SEARCH_6 = 1
        Frm85_LM_SEARCH_6_LOGIC = "="
        
        Frm85_LM_SEARCH_7 = 1
        Frm85_LM_SEARCH_7_LOGIC = "="
        
        LM_INVOICE_RASMI = 1
    Else
        Frm85_LM_SEARCH_6 = 0
        Frm85_LM_SEARCH_6_LOGIC = "="
        
        Frm85_LM_SEARCH_7 = 1
        Frm85_LM_SEARCH_7_LOGIC = "="
    End If
    
    If user_level = "Administration" Then
        Frm85_LM_SEARCH_10 = 1
        Frm85_LM_SEARCH_10_LOGIC = "="
        
        Frm85_LM_SEARCH_11 = 1
        Frm85_LM_SEARCH_11_LOGIC = "="
    Else
        Frm85_LM_SEARCH_10 = 0
        Frm85_LM_SEARCH_10_LOGIC = "="
        
        Frm85_LM_SEARCH_11 = 1
        Frm85_LM_SEARCH_11_LOGIC = "="
    End If
    
    If G_JENIS_JUALAN = "Barang Baru Sahaja" Then
    
        Frm85_LM_SEARCH_12 = 0
        Frm85_LM_SEARCH_12_LOGIC = "="
        
        Frm85_LM_SEARCH_13 = 0
        Frm85_LM_SEARCH_13_LOGIC = "="
        
    ElseIf G_JENIS_JUALAN = "Barang Trade In Sahaja" Then
    
        Frm85_LM_SEARCH_12 = 1
        Frm85_LM_SEARCH_12_LOGIC = "="
        
        Frm85_LM_SEARCH_13 = 1
        Frm85_LM_SEARCH_13_LOGIC = "="
        
    ElseIf G_JENIS_JUALAN = "Barang Baru Dan Barang Trade In" Then
    
        Frm85_LM_SEARCH_12 = 0
        Frm85_LM_SEARCH_12_LOGIC = "="
        
        Frm85_LM_SEARCH_13 = 1
        Frm85_LM_SEARCH_13_LOGIC = "="
        
    End If
    If Frm101.L47_Text = "Semua" Then
        Frm85_LM_SEARCH_20 = Null
        Frm85_LM_SEARCH_20_LOGIC = "<>"
    Else
        If InStr(1, Frm101.L47_Text, "  |  ") <> 0 Then
            Frm84_LM_EMP_NO = Split(Frm101.L47_Text, "  |  ")(1)
            Frm84_LM_EMP_NAMA = Split(Frm101.L47_Text, "  |  ")(0)
        End If
        Frm85_LM_SEARCH_20 = Frm84_LM_EMP_NAMA
        Frm85_LM_SEARCH_20_LOGIC = "="
    End If

    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh Jualan
        .Columns("C").ColumnWidth = 20 'No. Resit
        .Columns("D").ColumnWidth = 20 'No. Siri Produk
        .Columns("E").ColumnWidth = 40 'Kategori Produk
        .Columns("F").ColumnWidth = 20 'Purity
        .Columns("G").ColumnWidth = 20 'Berat Asal (g)
        .Columns("H").ColumnWidth = 20 'Berat Jualan (g)
        .Columns("I").ColumnWidth = 20 'Harga Semasa (RM/g)
        .Columns("J").ColumnWidth = 20 'Upah (RM)
        .Columns("K").ColumnWidth = 20 'Diskaun (%)
        .Columns("L").ColumnWidth = 20 'Harga Asal (RM)
        .Columns("M").ColumnWidth = 20 'Harga Selepas Diskaun (RM)
        .Columns("N").ColumnWidth = 20 'Adjustment (RM)
        .Columns("O").ColumnWidth = 20 'Harga Jualan (RM)
        .Columns("P").ColumnWidth = 20 'Jenis GST
        .Columns("Q").ColumnWidth = 20 'Jumlah GST (RM)
        .Columns("R").ColumnWidth = 0 'Untung(RM)
        .Columns("S").ColumnWidth = 0 'Harga Semasa Jika Restok (RM/g)
        .Columns("T").ColumnWidth = 0 'Untung 2 (RM)
        .Columns("U").ColumnWidth = 40 'Cawangan
        .Columns("V").ColumnWidth = 30 'Nama Pekerja
        .Columns("W").ColumnWidth = 20 'Dulang
    
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        .Cells(7, 1) = Frm85.L14_Text 'Header Report
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Jualan"
        .Cells(8, 3) = "No. Invoice"
        .Cells(8, 4) = "No. Siri Produk"
        .Cells(8, 5) = "Kategori Produk"
        .Cells(8, 6) = "Purity"
        .Cells(8, 7) = "Berat Asal (g)"
        .Cells(8, 8) = "Berat Jualan (g)"
        .Cells(8, 9) = "Harga Semasa (RM/g)"
        .Cells(8, 10) = "Upah (RM)"
        .Cells(8, 11) = "Harga Asal (RM)"
        .Cells(8, 12) = "Diskaun (%)"
        .Cells(8, 13) = "Harga Selepas Diskaun (RM)"
        .Cells(8, 14) = "Adjustment (RM)"
        .Cells(8, 15) = "Harga Jualan Termasuk GST(RM)"
        .Cells(8, 16) = "Jenis GST"
        .Cells(8, 17) = "Jumlah GST (RM)"
        .Cells(8, 18) = "Untung 1 (RM)"
        .Cells(8, 19) = "Harga Semasa Jika Restok (RM/g)"
        .Cells(8, 20) = "Untung 2 (RM)"
        .Cells(8, 21) = "Cawangan"
        .Cells(8, 22) = "Nama Pekerja"
        .Cells(8, 23) = "Dulang"
        
        For i = 1 To 23
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
        
        '#### Header Report ###
        'If Frm101.L9_Text = 0 Then .Cells(7, 1) = "Report Jualan Bagi Purity [" & Frm101.L5_Text & "] & Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "]" 'Report Header"
        'If Frm101.L9_Text = 1 Then .Cells(7, 1) = "Report Jualan Bagi Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] Dari " & TM & " Hingga " & TA 'Report Header"
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select * from 23_senarai_jualan where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND(bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (baru_or_ti " & Frm85_LM_SEARCH_12_LOGIC & "'" & Frm85_LM_SEARCH_12 & "' OR baru_or_ti " & Frm85_LM_SEARCH_13_LOGIC & "'" & Frm85_LM_SEARCH_13 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND " _
        & "kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND jenis_jualan " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND jualan_online " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select * from 23_senarai_jualan where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND(bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (baru_or_ti " & Frm85_LM_SEARCH_12_LOGIC & "'" & Frm85_LM_SEARCH_12 & "' OR baru_or_ti " & Frm85_LM_SEARCH_13_LOGIC & "'" & Frm85_LM_SEARCH_13 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND " _
        & "kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND jenis_jualan " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND jualan_online " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh) Then .Cells(8 + x, 2) = "'" & rs!tarikh 'Tarikh Jualan
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If LM_INVOICE_RASMI = 0 Then
                If Not IsNull(rs!no_resit) Then .Cells(8 + x, 3) = rs!no_resit 'No. Resit
            Else
                If Not IsNull(rs!no_invoice_r) Then .Cells(8 + x, 3) = rs!no_invoice_r 'No. Resit
            End If
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 4) = rs!no_siri_Produk 'No. Siri Produk
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 5) = rs!kategori_Produk 'Kategori Produk
            
            If Not IsNull(rs!purity) Then .Cells(8 + x, 6) = rs!purity 'Purity
            .Cells(8 + x, 6).HorizontalAlignment = xlCenter

            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!Berat_Asal) Then
                .Cells(8 + x, 7) = Format(rs!Berat_Asal, "#,##0.00") 'Berat Asal (g)
            Else
                .Cells(8 + x, 7) = "0.00" 'Berat Asal (g)
            End If
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"

            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!berat_jualan) Then
                .Cells(8 + x, 8) = Format(rs!berat_jualan, "#,##0.00") 'Berat Jualan (g)
            Else
                .Cells(8 + x, 8) = "0.00" 'Berat Jualan (g)
            End If
            
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_Semasa) Then
                .Cells(8 + x, 9) = Format(rs!harga_Semasa, "#,##0.00") 'Harga Semasa (RM/g)
            Else
                .Cells(8 + x, 9) = "0.00" 'Harga Semasa (RM/g)
            End If
            
            .Cells(8 + x, 10).NumberFormat = "#,##0.00"
            .Cells(8 + x, 10).HorizontalAlignment = xlRight
            If Not IsNull(rs!UPAH) Then
                .Cells(8 + x, 10) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
            Else
                .Cells(8 + x, 10) = "0.00" 'Upah (RM)
            End If
            
            .Cells(8 + x, 11).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_asal) Then
                .Cells(8 + x, 11) = Format(rs!harga_asal, "#,##0.00") 'Harga Asal (RM)
            Else
                .Cells(8 + x, 11) = "0.00" 'Harga Asal (RM)
            End If
            .Cells(8 + x, 11).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 12).NumberFormat = "#,##0.00"
            .Cells(8 + x, 12).HorizontalAlignment = xlRight
            If Not IsNull(rs!diskaun) Then
                .Cells(8 + x, 12) = rs!diskaun 'Diskaun (%)
            Else
                .Cells(8 + x, 12) = "0.00" 'Diskaun (%)
            End If
            
            .Cells(8 + x, 13).NumberFormat = "#,##0.00"
            .Cells(8 + x, 13).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_lepas_diskaun) Then
                .Cells(8 + x, 13) = Format(rs!harga_lepas_diskaun, "#,##0.00") 'Harga Selepas Diskaun (RM)
            Else
                .Cells(8 + x, 13) = "0.00" 'Harga Selepas Diskaun (RM)
            End If

            .Cells(8 + x, 14).NumberFormat = "#,##0.00"
            .Cells(8 + x, 14).HorizontalAlignment = xlRight
            If Not IsNull(rs!adjustment) Then
                .Cells(8 + x, 14) = Format(rs!adjustment, "#,##0.00") 'Adjustment (RM)
            Else
                .Cells(8 + x, 14) = "0.00" 'Adjustment (RM)
            End If

            .Cells(8 + x, 15).NumberFormat = "#,##0.00"
            .Cells(8 + x, 15).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_jualan_dengan_gst) Then
                .Cells(8 + x, 15) = Format(rs!harga_jualan_dengan_gst, "#,##0.00")
            Else
                .Cells(8 + x, 15) = "0.00"
            End If
            
            .Cells(8 + x, 16).HorizontalAlignment = xlCenter
            If Not IsNull(rs!gst_ari_nashi) Then
                If rs!gst_ari_nashi = "ZR (L)" Then '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                    .Cells(8 + x, 16) = "ZR (L)"
                ElseIf rs!gst_ari_nashi = "SR" Then
                    .Cells(8 + x, 16) = "SR"
                End If
            End If

            .Cells(8 + x, 17).NumberFormat = "#,##0.00"
            .Cells(8 + x, 17).HorizontalAlignment = xlRight
            If Not IsNull(rs!jumlah_gst) Then
                .Cells(8 + x, 17) = Format(rs!jumlah_gst, "#,##0.00") 'Jumlah GST (RM)
            Else
                .Cells(8 + x, 17) = "0.00" 'Jumlah GST (RM)
            End If

            .Cells(8 + x, 18).NumberFormat = "#,##0.00"
            .Cells(8 + x, 18).HorizontalAlignment = xlRight
            If Not IsNull(rs!untung) Then
                .Cells(8 + x, 18) = Format(rs!untung, "#,##0.00") 'Untung(RM)
            Else
                .Cells(8 + x, 18) = "0.00" 'Untung(RM)
            End If
            
            .Cells(8 + x, 19).NumberFormat = "#,##0.00"
            .Cells(8 + x, 19).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_per_gram_supplier) Then
                .Cells(8 + x, 19) = Format(rs!harga_per_gram_supplier, "#,##0.00") 'Harga Semasa Jika Restok (RM/g)
            Else
                .Cells(8 + x, 19) = "0.00" 'Harga Semasa Jika Restok (RM/g)
            End If
            
            .Cells(8 + x, 20).NumberFormat = "#,##0.00"
            .Cells(8 + x, 20).HorizontalAlignment = xlRight
            If Not IsNull(rs!untung2) Then
                .Cells(8 + x, 20) = Format(rs!untung2, "#,##0.00") 'Untung(RM)
            Else
                .Cells(8 + x, 20) = "0.00" 'Untung(RM)
            End If
            If Not IsNull(rs!cawangan) Then
                .Cells(8 + x, 21) = rs!cawangan
            Else
                .Cells(8 + x, 21) = vbNullString
            End If
            If Not IsNull(rs!nama_pekerja) Then
                .Cells(8 + x, 22) = rs!nama_pekerja
            Else
                .Cells(8 + x, 22) = vbNullString
            End If
            If Not IsNull(rs!dulang) Then
                .Cells(8 + x, 23) = rs!dulang
            Else
                .Cells(8 + x, 23) = vbNullString
            End If

            For Col = 1 To 23
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 2
        
        '#### Jumlah Bilangan Barang Terjual Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select COUNT(no_siri_Produk) from 23_senarai_jualan where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND(bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (baru_or_ti " & Frm85_LM_SEARCH_12_LOGIC & "'" & Frm85_LM_SEARCH_12 & "' OR baru_or_ti " & Frm85_LM_SEARCH_13_LOGIC & "'" & Frm85_LM_SEARCH_13 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND " _
        & "kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND jenis_jualan " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND jualan_online " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select COUNT(no_siri_Produk) from 23_senarai_jualan where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND(bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (baru_or_ti " & Frm85_LM_SEARCH_12_LOGIC & "'" & Frm85_LM_SEARCH_12 & "' OR baru_or_ti " & Frm85_LM_SEARCH_13_LOGIC & "'" & Frm85_LM_SEARCH_13 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND " _
        & "kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND jenis_jualan " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND jualan_online " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Bilangan Barang : " & rs(0)
        Else
            .Cells(8 + Y, 1) = "Bilangan Barang : " & 0
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Bilangan Barang : " & 0
        End If
        '#### Jumlah Bilangan Barang Terjual Keseluruhan #### - End
        
        Y = Y + 1

        '#### Jumlah Berat Jualan Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select SUM(Berat_Jualan) from 23_senarai_jualan where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND(bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (baru_or_ti " & Frm85_LM_SEARCH_12_LOGIC & "'" & Frm85_LM_SEARCH_12 & "' OR baru_or_ti " & Frm85_LM_SEARCH_13_LOGIC & "'" & Frm85_LM_SEARCH_13 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND " _
        & "kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND jenis_jualan " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND jualan_online " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select SUM(Berat_Jualan) from 23_senarai_jualan where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND(bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (baru_or_ti " & Frm85_LM_SEARCH_12_LOGIC & "'" & Frm85_LM_SEARCH_12 & "' OR baru_or_ti " & Frm85_LM_SEARCH_13_LOGIC & "'" & Frm85_LM_SEARCH_13 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND " _
        & "kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND jenis_jualan " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND jualan_online " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Berat Jualan : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat Jualan : " & Format(0, "#,##0.00 g")
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Berat Jualan : " & Format(0, "#,##0.00 g")
        End If
        '#### Jumlah Berat Jualan Keseluruhan #### - End
        
        Y = Y + 1

        '#### Jumlah Harga Jualan Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select SUM(harga_jualan_dengan_gst) from 23_senarai_jualan where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND(bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (baru_or_ti " & Frm85_LM_SEARCH_12_LOGIC & "'" & Frm85_LM_SEARCH_12 & "' OR baru_or_ti " & Frm85_LM_SEARCH_13_LOGIC & "'" & Frm85_LM_SEARCH_13 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND " _
        & "kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND jenis_jualan " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND jualan_online " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select SUM(harga_jualan_dengan_gst) from 23_senarai_jualan where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND(bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (baru_or_ti " & Frm85_LM_SEARCH_12_LOGIC & "'" & Frm85_LM_SEARCH_12 & "' OR baru_or_ti " & Frm85_LM_SEARCH_13_LOGIC & "'" & Frm85_LM_SEARCH_13 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND " _
        & "kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND jenis_jualan " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND jualan_online " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Jualan : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Jualan : RM " & "0.00"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Jualan : RM " & "0.00"
        End If
        '#### Jumlah Harga Jualan Keseluruhan #### - End

GoTo a:

        Y = Y + 1

        '#### Jumlah Keuntungan Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select SUM(untung) from 23_senarai_jualan where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND(bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND jenis_jualan " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND jualan_online " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select SUM(untung) from 23_senarai_jualan where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND(bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND jenis_jualan " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND jualan_online " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Keuntungan 1 : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Keuntungan 1 : RM " & "0.00"
        End If
        
        rs.Close
        Set rs = Nothing

        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Keuntungan 1 : RM " & "0.00"
        End If
        '#### Jumlah Keuntungan Keseluruhan #### - End
        
        Y = Y + 1

        '#### Jumlah Keuntungan Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select SUM(untung2) from 23_senarai_jualan where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND(bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND jenis_jualan " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND jualan_online " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select SUM(untung2) from 23_senarai_jualan where nama_pekerja " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND(bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND jenis_jualan " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND jualan_online " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Keuntungan 2 : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Keuntungan 2 : RM " & "0.00"
        End If
        
        rs.Close
        Set rs = Nothing

        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Keuntungan 2 : RM " & "0.00"
        End If
        '#### Jumlah Keuntungan Keseluruhan #### - End
        
a:
        
        Y = Y + 4
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
End Sub
Sub Frm85_overall_jual_invoice()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then

    x = 0
    Frm85_LM_BERAT = 0
    Frm85_LM_HARGA = 0
    Frm85_LM_UNTUNG = 0
    
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh Jualan
        .Columns("C").ColumnWidth = 20 'No. Resit
        .Columns("D").ColumnWidth = 20 'No. Siri Produk
        .Columns("E").ColumnWidth = 40 'Kategori Produk
        .Columns("F").ColumnWidth = 20 'Purity
        .Columns("G").ColumnWidth = 20 'Berat Asal (g)
        .Columns("H").ColumnWidth = 20 'Berat Jualan (g)
        .Columns("I").ColumnWidth = 20 'Harga Semasa (RM/g)
        .Columns("J").ColumnWidth = 20 'Upah (RM)
        .Columns("K").ColumnWidth = 20 'Diskaun (%)
        .Columns("L").ColumnWidth = 20 'Harga Asal (RM)
        .Columns("M").ColumnWidth = 20 'Harga Selepas Diskaun (RM)
        .Columns("N").ColumnWidth = 20 'Adjustment (RM)
        .Columns("O").ColumnWidth = 20 'Harga Jualan (RM)
        .Columns("P").ColumnWidth = 20 'Jenis GST
        .Columns("Q").ColumnWidth = 20 'Jumlah GST (RM)
        .Columns("R").ColumnWidth = 0 'Untung(RM)
        .Columns("S").ColumnWidth = 0 'Harga Semasa Jika Restok (RM/g)
        .Columns("T").ColumnWidth = 0 'Untung 2 (RM)
        .Columns("U").ColumnWidth = 40 'Cawangan
        .Columns("V").ColumnWidth = 30 'Nama Pekerja
        .Columns("W").ColumnWidth = 20 'Dulang
        
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
        
        user_level = MDI_frm1.L4_Text
        
        Dim LM_FIELD As String
        LM_INVOICE_RASMI = 0
        
        If user_level = "Guest/User" Then
            Frm85_LM_SEARCH_6 = 1
            Frm85_LM_SEARCH_6_LOGIC = "="
            
            Frm85_LM_SEARCH_7 = 1
            Frm85_LM_SEARCH_7_LOGIC = "="
            LM_INVOICE_RASMI = 1
            LM_FIELD = "no_invoice_r"
        Else
            Frm85_LM_SEARCH_6 = 0
            Frm85_LM_SEARCH_6_LOGIC = "="
            
            Frm85_LM_SEARCH_7 = 1
            Frm85_LM_SEARCH_7_LOGIC = "="
            
            LM_FIELD = "no_resit"
        End If
        
        If user_level = "Administration" Then
            Frm85_LM_SEARCH_10 = 1
            Frm85_LM_SEARCH_10_LOGIC = "="
            
            Frm85_LM_SEARCH_11 = 1
            Frm85_LM_SEARCH_11_LOGIC = "="
        Else
            Frm85_LM_SEARCH_10 = 0
            Frm85_LM_SEARCH_10_LOGIC = "="
            
            Frm85_LM_SEARCH_11 = 1
            Frm85_LM_SEARCH_11_LOGIC = "="
        End If
        
        If Frm101.L46_Text = "Semua cawangan" Then
        
            Frm85_SEARCH_8 = Null
            Frm85_SEARCH_8_LOGIC = "<>"
            Frm85_SEARCH_9 = Null
            Frm85_SEARCH_9_LOGIC = "<>"
            
        Else
        
            Frm85_SEARCH_8 = Frm101.L46_Text
            Frm85_SEARCH_8_LOGIC = "="
            Frm85_SEARCH_9 = "HQ"
            Frm85_SEARCH_9_LOGIC = "="
            
        End If
         
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        .Cells(7, 1) = Frm85.L14_Text 'Header Report
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Jualan"
        .Cells(8, 3) = "No. Invoice"
        .Cells(8, 4) = "No. Siri Produk"
        .Cells(8, 5) = "Kategori Produk"
        .Cells(8, 6) = "Purity"
        .Cells(8, 7) = "Berat Asal (g)"
        .Cells(8, 8) = "Berat Jualan (g)"
        .Cells(8, 9) = "Harga Semasa (RM/g)"
        .Cells(8, 10) = "Upah (RM)"
        .Cells(8, 11) = "Harga Asal (RM)"
        .Cells(8, 12) = "Diskaun (%)"
        .Cells(8, 13) = "Harga Selepas Diskaun (RM)"
        .Cells(8, 14) = "Adjustment (RM)"
        .Cells(8, 15) = "Harga Jualan Termasuk GST (RM)"
        .Cells(8, 16) = "Jenis GST"
        .Cells(8, 17) = "Jumlah GST (RM)"
        .Cells(8, 18) = "Untung 1 (RM)"
        .Cells(8, 19) = "Harga Semasa Jika Restok (RM/g)"
        .Cells(8, 20) = "Untung 2 (RM)"
        .Cells(8, 21) = "Cawangan"
        .Cells(8, 22) = "Nama Pekerja"
        .Cells(8, 23) = "Dulang"
        
        For i = 1 To 23
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 23_senarai_jualan where " & LM_FIELD & "='" & UCase(Frm101.L5_Text) & "' AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by no_resit ASC , tarikh ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh) Then .Cells(8 + x, 2) = "'" & rs!tarikh 'Tarikh Jualan
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_resit) Then .Cells(8 + x, 3) = rs!no_resit 'No. Resit
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If LM_INVOICE_RASMI = 0 Then
                If Not IsNull(rs!no_resit) Then .Cells(8 + x, 3) = rs!no_resit 'No. Resit
            Else
                If Not IsNull(rs!no_invoice_r) Then .Cells(8 + x, 3) = rs!no_invoice_r 'No. Resit
            End If
    
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 4) = rs!no_siri_Produk 'No. Siri Produk
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 5) = rs!kategori_Produk 'Kategori Produk
            
            If Not IsNull(rs!purity) Then .Cells(8 + x, 6) = rs!purity 'Purity
            .Cells(8 + x, 6).HorizontalAlignment = xlCenter

            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!Berat_Asal) Then
                .Cells(8 + x, 7) = Format(rs!Berat_Asal, "#,##0.00") 'Berat Asal (g)
            Else
                .Cells(8 + x, 7) = "0.00" 'Berat Asal (g)
            End If

            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!berat_jualan) Then
                .Cells(8 + x, 8) = Format(rs!berat_jualan, "#,##0.00") 'Berat Jualan (g)
            Else
                .Cells(8 + x, 8) = "0.00" 'Berat Jualan (g)
            End If
            
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_Semasa) Then
                .Cells(8 + x, 9) = Format(rs!harga_Semasa, "#,##0.00") 'Harga Semasa (RM/g)
            Else
                .Cells(8 + x, 9) = "0.00" 'Harga Semasa (RM/g)
            End If
            
            .Cells(8 + x, 10).NumberFormat = "#,##0.00"
            .Cells(8 + x, 10).HorizontalAlignment = xlRight
            If Not IsNull(rs!UPAH) Then
                .Cells(8 + x, 10) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
            Else
                .Cells(8 + x, 10) = "0.00" 'Upah (RM)
            End If
            
            .Cells(8 + x, 11).NumberFormat = "#,##0.00"
            .Cells(8 + x, 11).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_asal) Then
                .Cells(8 + x, 11) = Format(rs!harga_asal, "#,##0.00") 'Harga Asal (RM)
            Else
                .Cells(8 + x, 11) = "0.00" 'Harga Asal (RM)
            End If
            
            .Cells(8 + x, 12).NumberFormat = "#,##0.00"
            .Cells(8 + x, 12).HorizontalAlignment = xlRight
            If Not IsNull(rs!diskaun) Then
                .Cells(8 + x, 12) = rs!diskaun 'Diskaun (%)
            Else
                .Cells(8 + x, 12) = "0.00" 'Diskaun (%)
            End If
            
            .Cells(8 + x, 13).NumberFormat = "#,##0.00"
            .Cells(8 + x, 13).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_lepas_diskaun) Then
                .Cells(8 + x, 13) = Format(rs!harga_lepas_diskaun, "#,##0.00") 'Harga Selepas Diskaun (RM)
            Else
                .Cells(8 + x, 13) = "0.00" 'Harga Selepas Diskaun (RM)
            End If

            .Cells(8 + x, 14).NumberFormat = "#,##0.00"
            .Cells(8 + x, 14).HorizontalAlignment = xlRight
            If Not IsNull(rs!adjustment) Then
                .Cells(8 + x, 14) = Format(rs!adjustment, "#,##0.00") 'Adjustment (RM)
            Else
                .Cells(8 + x, 14) = "0.00" 'Adjustment (RM)
            End If

            .Cells(8 + x, 15).NumberFormat = "#,##0.00"
            .Cells(8 + x, 15).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_jualan_dengan_gst) Then
                .Cells(8 + x, 15) = Format(rs!harga_jualan_dengan_gst, "#,##0.00")
            Else
                .Cells(8 + x, 15) = "0.00"
            End If
            
            .Cells(8 + x, 16).HorizontalAlignment = xlCenter
            If Not IsNull(rs!gst_ari_nashi) Then
                If rs!gst_ari_nashi = "ZR (L)" Then '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                    .Cells(8 + x, 16) = "ZR (L)"
                ElseIf rs!gst_ari_nashi = "SR" Then
                    .Cells(8 + x, 16) = "SR"
                End If
            End If

            .Cells(8 + x, 17).NumberFormat = "#,##0.00"
            .Cells(8 + x, 17).HorizontalAlignment = xlRight
            If Not IsNull(rs!jumlah_gst) Then
                .Cells(8 + x, 17) = Format(rs!jumlah_gst, "#,##0.00") 'Jumlah GST (RM)
            Else
                .Cells(8 + x, 17) = "0.00" 'Jumlah GST (RM)
            End If

            .Cells(8 + x, 18).NumberFormat = "#,##0.00"
            .Cells(8 + x, 18).HorizontalAlignment = xlRight
            If Not IsNull(rs!untung) Then
                .Cells(8 + x, 18) = Format(rs!untung, "#,##0.00") 'Untung(RM)
            Else
                .Cells(8 + x, 18) = "0.00" 'Untung(RM)
            End If
            
            .Cells(8 + x, 19).NumberFormat = "#,##0.00"
            .Cells(8 + x, 19).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_per_gram_supplier) Then
                .Cells(8 + x, 19) = Format(rs!harga_per_gram_supplier, "#,##0.00") 'Harga Semasa Jika Restok (RM/g)
            Else
                .Cells(8 + x, 19) = "0.00" 'Harga Semasa Jika Restok (RM/g)
            End If
            
            .Cells(8 + x, 20).NumberFormat = "#,##0.00"
            .Cells(8 + x, 20).HorizontalAlignment = xlRight
            If Not IsNull(rs!untung2) Then
                .Cells(8 + x, 20) = Format(rs!untung2, "#,##0.00") 'Untung(RM)
            Else
                .Cells(8 + x, 20) = "0.00" 'Untung(RM)
            End If
            
            If Not IsNull(rs!cawangan) Then
                .Cells(8 + x, 21) = rs!cawangan
            Else
                .Cells(8 + x, 21) = vbNullString
            End If
            If Not IsNull(rs!nama_pekerja) Then
                .Cells(8 + x, 22) = rs!nama_pekerja
            Else
                .Cells(8 + x, 22) = vbNullString
            End If
            If Not IsNull(rs!dulang) Then
                .Cells(8 + x, 23) = rs!dulang
            Else
                .Cells(8 + x, 23) = vbNullString
            End If

            For Col = 1 To 23
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 2
        
        '#### Jumlah Bilangan Barang Terjual Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select COUNT(no_siri_Produk) from 23_senarai_jualan where " & LM_FIELD & "='" & UCase(Frm101.L5_Text) & "' AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Bilangan Barang : " & rs(0)
        Else
            .Cells(8 + Y, 1) = "Bilangan Barang : " & 0
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Bilangan Barang : " & 0
        End If
        '#### Jumlah Bilangan Barang Terjual Keseluruhan #### - End
        
        Y = Y + 1

        '#### Jumlah Berat Jualan Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(Berat_Jualan) from 23_senarai_jualan where " & LM_FIELD & "='" & UCase(Frm101.L5_Text) & "' AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Berat Jualan : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat Jualan : " & Format(0, "#,##0.00 g")
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Berat Jualan : " & Format(0, "#,##0.00 g")
        End If
        '#### Jumlah Berat Jualan Keseluruhan #### - End
        
        Y = Y + 1

        '#### Jumlah Harga Jualan Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(harga_jualan_dengan_gst) from 23_senarai_jualan where " & LM_FIELD & "='" & UCase(Frm101.L5_Text) & "' AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Jualan : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Jualan : RM " & "0.00"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Jualan : RM " & "0.00"
        End If
        '#### Jumlah Harga Jualan Keseluruhan #### - End

GoTo a:

        Y = Y + 1

        '#### Jumlah Keuntungan Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(untung) from 23_senarai_jualan where " & LM_FIELD & "='" & UCase(Frm101.L5_Text) & "' AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Keuntungan 1 : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Keuntungan 1 : RM " & "0.00"
        End If
        
        rs.Close
        Set rs = Nothing

        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Keuntungan 1 : RM " & "0.00"
        End If
        '#### Jumlah Keuntungan Keseluruhan #### - End
        
        Y = Y + 1

        '#### Jumlah Keuntungan Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(untung2) from 23_senarai_jualan where " & LM_FIELD & "='" & UCase(Frm101.L5_Text) & "' AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Keuntungan 2 : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Keuntungan 2 : RM " & "0.00"
        End If
        
        rs.Close
        Set rs = Nothing

        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Keuntungan 2 : RM " & "0.00"
        End If
        '#### Jumlah Keuntungan Keseluruhan #### - End
        
a:
        
        Y = Y + 4
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
End Sub
Sub Frm85_overall_jual_no_siri()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then

    x = 0
    Frm85_LM_BERAT = 0
    Frm85_LM_HARGA = 0
    Frm85_LM_UNTUNG = 0
    
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh Jualan
        .Columns("C").ColumnWidth = 20 'No. Resit
        .Columns("D").ColumnWidth = 20 'No. Siri Produk
        .Columns("E").ColumnWidth = 40 'Kategori Produk
        .Columns("F").ColumnWidth = 20 'Purity
        .Columns("G").ColumnWidth = 20 'Berat Asal (g)
        .Columns("H").ColumnWidth = 20 'Berat Jualan (g)
        .Columns("I").ColumnWidth = 20 'Harga Semasa (RM/g)
        .Columns("J").ColumnWidth = 20 'Upah (RM)
        .Columns("K").ColumnWidth = 20 'Diskaun (%)
        .Columns("L").ColumnWidth = 20 'Harga Asal (RM)
        .Columns("M").ColumnWidth = 20 'Harga Selepas Diskaun (RM)
        .Columns("N").ColumnWidth = 20 'Adjustment (RM)
        .Columns("O").ColumnWidth = 20 'Harga Jualan (RM)
        .Columns("P").ColumnWidth = 20 'Jenis GST
        .Columns("Q").ColumnWidth = 20 'Jumlah GST (RM)
        .Columns("R").ColumnWidth = 0 'Untung(RM)
        .Columns("S").ColumnWidth = 0 'Harga Semasa Jika Restok (RM/g)
        .Columns("T").ColumnWidth = 0 'Untung 2 (RM)
        .Columns("U").ColumnWidth = 40 'Cawangan
        .Columns("V").ColumnWidth = 30 'Nama Pekerja
        .Columns("W").ColumnWidth = 20 'Dulang
        
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        .Cells(7, 1) = Frm85.L14_Text 'Header Report
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Jualan"
        .Cells(8, 3) = "No. Invoice"
        .Cells(8, 4) = "No. Siri Produk"
        .Cells(8, 5) = "Kategori Produk"
        .Cells(8, 6) = "Purity"
        .Cells(8, 7) = "Berat Asal (g)"
        .Cells(8, 8) = "Berat Jualan (g)"
        .Cells(8, 9) = "Harga Semasa (RM/g)"
        .Cells(8, 10) = "Upah (RM)"
        .Cells(8, 11) = "Harga Asal (RM)"
        .Cells(8, 12) = "Diskaun (%)"
        .Cells(8, 13) = "Harga Selepas Diskaun (RM)"
        .Cells(8, 14) = "Adjustment (RM)"
        .Cells(8, 15) = "Harga Jualan Termasuk GST (RM)"
        .Cells(8, 16) = "Jenis GST"
        .Cells(8, 17) = "Jumlah GST (RM)"
        .Cells(8, 18) = "Untung 1 (RM)"
        .Cells(8, 19) = "Harga Semasa Jika Restok (RM/g)"
        .Cells(8, 20) = "Untung 2 (RM)"
        .Cells(8, 21) = "Cawangan"
        .Cells(8, 22) = "Nama Pekerja"
        .Cells(8, 23) = "Dulang"
        
        For i = 1 To 23
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
        
        user_level = MDI_frm1.L4_Text
        LM_INVOICE_RASMI = 0
        
        If user_level = "Guest/User" Then
            Frm85_LM_SEARCH_6 = 1
            Frm85_LM_SEARCH_6_LOGIC = "="
            LM_INVOICE_RASMI = 1
            Frm85_LM_SEARCH_7 = 1
            Frm85_LM_SEARCH_7_LOGIC = "="
        Else
            Frm85_LM_SEARCH_6 = 0
            Frm85_LM_SEARCH_6_LOGIC = "="
            
            Frm85_LM_SEARCH_7 = 1
            Frm85_LM_SEARCH_7_LOGIC = "="
        End If
        
        If user_level = "Administration" Then
            Frm85_LM_SEARCH_10 = 1
            Frm85_LM_SEARCH_10_LOGIC = "="
            
            Frm85_LM_SEARCH_11 = 1
            Frm85_LM_SEARCH_11_LOGIC = "="
        Else
            Frm85_LM_SEARCH_10 = 0
            Frm85_LM_SEARCH_10_LOGIC = "="
            
            Frm85_LM_SEARCH_11 = 1
            Frm85_LM_SEARCH_11_LOGIC = "="
        End If
        If Frm101.L46_Text = "Semua cawangan" Then
        
            Frm85_SEARCH_8 = Null
            Frm85_SEARCH_8_LOGIC = "<>"
            Frm85_SEARCH_9 = Null
            Frm85_SEARCH_9_LOGIC = "<>"
            
        Else
        
            Frm85_SEARCH_8 = Frm101.L46_Text
            Frm85_SEARCH_8_LOGIC = "="
            Frm85_SEARCH_9 = "HQ"
            Frm85_SEARCH_9_LOGIC = "="
            
        End If
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 23_senarai_jualan where (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh) Then .Cells(8 + x, 2) = "'" & rs!tarikh 'Tarikh Jualan
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If LM_INVOICE_RASMI = 0 Then
                If Not IsNull(rs!no_resit) Then .Cells(8 + x, 3) = rs!no_resit 'No. Resit
            Else
                If Not IsNull(rs!no_invoice_r) Then .Cells(8 + x, 3) = rs!no_invoice_r 'No. Resit
            End If
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 4) = rs!no_siri_Produk 'No. Siri Produk
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 5) = rs!kategori_Produk 'Kategori Produk
            
            If Not IsNull(rs!purity) Then .Cells(8 + x, 6) = rs!purity 'Purity
            .Cells(8 + x, 6).HorizontalAlignment = xlCenter

            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!Berat_Asal) Then
                .Cells(8 + x, 7) = Format(rs!Berat_Asal, "#,##0.00") 'Berat Asal (g)
            Else
                .Cells(8 + x, 7) = "0.00" 'Berat Asal (g)
            End If

            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!berat_jualan) Then
                .Cells(8 + x, 8) = Format(rs!berat_jualan, "#,##0.00") 'Berat Jualan (g)
            Else
                .Cells(8 + x, 8) = "0.00" 'Berat Jualan (g)
            End If
            
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_Semasa) Then
                .Cells(8 + x, 9) = Format(rs!harga_Semasa, "#,##0.00") 'Harga Semasa (RM/g)
            Else
                .Cells(8 + x, 9) = "0.00" 'Harga Semasa (RM/g)
            End If
            
            .Cells(8 + x, 10).NumberFormat = "#,##0.00"
            .Cells(8 + x, 10).HorizontalAlignment = xlRight
            If Not IsNull(rs!UPAH) Then
                .Cells(8 + x, 10) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
            Else
                .Cells(8 + x, 10) = "0.00" 'Upah (RM)
            End If
            
            .Cells(8 + x, 11).NumberFormat = "#,##0.00"
            .Cells(8 + x, 11).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_asal) Then
                .Cells(8 + x, 11) = Format(rs!harga_asal, "#,##0.00") 'Harga Asal (RM)
            Else
                .Cells(8 + x, 11) = "0.00" 'Harga Asal (RM)
            End If
            
            .Cells(8 + x, 12).NumberFormat = "#,##0.00"
            .Cells(8 + x, 12).HorizontalAlignment = xlRight
            If Not IsNull(rs!diskaun) Then
                .Cells(8 + x, 12) = rs!diskaun 'Diskaun (%)
            Else
                .Cells(8 + x, 12) = "0.00" 'Diskaun (%)
            End If
            
            .Cells(8 + x, 13).NumberFormat = "#,##0.00"
            .Cells(8 + x, 13).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_lepas_diskaun) Then
                .Cells(8 + x, 13) = Format(rs!harga_lepas_diskaun, "#,##0.00") 'Harga Selepas Diskaun (RM)
            Else
                .Cells(8 + x, 13) = "0.00" 'Harga Selepas Diskaun (RM)
            End If

            .Cells(8 + x, 14).NumberFormat = "#,##0.00"
            .Cells(8 + x, 14).HorizontalAlignment = xlRight
            If Not IsNull(rs!adjustment) Then
                .Cells(8 + x, 14) = Format(rs!adjustment, "#,##0.00") 'Adjustment (RM)
            Else
                .Cells(8 + x, 14) = "0.00" 'Adjustment (RM)
            End If

            .Cells(8 + x, 15).NumberFormat = "#,##0.00"
            .Cells(8 + x, 15).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_jualan_dengan_gst) Then
                .Cells(8 + x, 15) = Format(rs!harga_jualan_dengan_gst, "#,##0.00")
            Else
                .Cells(8 + x, 15) = "0.00"
            End If
            
            .Cells(8 + x, 16).HorizontalAlignment = xlCenter
            If Not IsNull(rs!gst_ari_nashi) Then
                If rs!gst_ari_nashi = "ZR (L)" Then '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                    .Cells(8 + x, 16) = "ZR (L)"
                ElseIf rs!gst_ari_nashi = "SR" Then
                    .Cells(8 + x, 16) = "SR"
                End If
            End If

            .Cells(8 + x, 17).NumberFormat = "#,##0.00"
            .Cells(8 + x, 17).HorizontalAlignment = xlRight
            If Not IsNull(rs!jumlah_gst) Then
                .Cells(8 + x, 17) = Format(rs!jumlah_gst, "#,##0.00") 'Jumlah GST (RM)
            Else
                .Cells(8 + x, 17) = "0.00" 'Jumlah GST (RM)
            End If

            .Cells(8 + x, 18).NumberFormat = "#,##0.00"
            .Cells(8 + x, 18).HorizontalAlignment = xlRight
            If Not IsNull(rs!untung) Then
                .Cells(8 + x, 18) = Format(rs!untung, "#,##0.00") 'Untung(RM)
            Else
                .Cells(8 + x, 18) = "0.00" 'Untung(RM)
            End If
            
            .Cells(8 + x, 19).NumberFormat = "#,##0.00"
            .Cells(8 + x, 19).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_per_gram_supplier) Then
                .Cells(8 + x, 19) = Format(rs!harga_per_gram_supplier, "#,##0.00") 'Harga Semasa Jika Restok (RM/g)
            Else
                .Cells(8 + x, 19) = "0.00" 'Harga Semasa Jika Restok (RM/g)
            End If
            
            .Cells(8 + x, 20).NumberFormat = "#,##0.00"
            .Cells(8 + x, 20).HorizontalAlignment = xlRight
            If Not IsNull(rs!untung2) Then
                .Cells(8 + x, 20) = Format(rs!untung2, "#,##0.00") 'Untung(RM)
            Else
                .Cells(8 + x, 20) = "0.00" 'Untung(RM)
            End If
            
            If Not IsNull(rs!cawangan) Then
                .Cells(8 + x, 21) = rs!cawangan
            Else
                .Cells(8 + x, 21) = vbNullString
            End If

            If Not IsNull(rs!nama_pekerja) Then
                .Cells(8 + x, 22) = rs!nama_pekerja
            Else
                .Cells(8 + x, 22) = vbNullString
            End If
            If Not IsNull(rs!dulang) Then
                .Cells(8 + x, 23) = rs!dulang
            Else
                .Cells(8 + x, 23) = vbNullString
            End If

            For Col = 1 To 23
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 2
        
        '#### Jumlah Bilangan Barang Terjual Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select COUNT(no_siri_Produk) from 23_senarai_jualan where (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_produk='" & Frm101.L5_Text & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Bilangan Barang : " & rs(0)
        Else
            .Cells(8 + Y, 1) = "Bilangan Barang : " & 0
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Bilangan Barang : " & 0
        End If
        '#### Jumlah Bilangan Barang Terjual Keseluruhan #### - End
        
        Y = Y + 1

        '#### Jumlah Berat Jualan Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(Berat_Jualan) from 23_senarai_jualan where (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_produk='" & Frm101.L5_Text & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Berat Jualan : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat Jualan : " & Format(0, "#,##0.00 g")
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Berat Jualan : " & Format(0, "#,##0.00 g")
        End If
        '#### Jumlah Berat Jualan Keseluruhan #### - End
        
        Y = Y + 1

        '#### Jumlah Harga Jualan Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(harga_dengan_gst) from 23_senarai_jualan where (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_produk='" & Frm101.L5_Text & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Jualan : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Jualan : RM " & "0.00"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Jualan : RM " & "0.00"
        End If
        '#### Jumlah Harga Jualan Keseluruhan #### - End

GoTo a:

        Y = Y + 1

        '#### Jumlah Keuntungan Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(untung) from 23_senarai_jualan where (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_produk='" & Frm101.L5_Text & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Keuntungan 1 : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Keuntungan 1 : RM " & "0.00"
        End If
        
        rs.Close
        Set rs = Nothing

        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Keuntungan 1 : RM " & "0.00"
        End If
        '#### Jumlah Keuntungan Keseluruhan #### - End
        
        Y = Y + 1

        '#### Jumlah Keuntungan Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(untung2) from 23_senarai_jualan where (bil_rasmi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR bil_rasmi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_produk='" & Frm101.L5_Text & "' AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND status_rekod = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Keuntungan 2 : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Keuntungan 2 : RM " & "0.00"
        End If
        
        rs.Close
        Set rs = Nothing

        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Keuntungan 2 : RM " & "0.00"
        End If
        '#### Jumlah Keuntungan Keseluruhan #### - End
        
a:
        
        Y = Y + 4
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
End Sub
Sub Frm85_bk_trade_overall()
'on error resume next
Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If
If Frm101.L6_Text = "Semua Kategori Produk" Then
    Frm85_LM_SEARCH_2 = Null
    Frm85_LM_SEARCH_2_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_2 = Frm101.L6_Text
    Frm85_LM_SEARCH_2_LOGIC = "="
End If
If Frm101.L34_Text = "Semua Dulang" Then
    Frm85_LM_SEARCH_3 = Null
    Frm85_LM_SEARCH_3_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_3 = Frm101.L34_Text
    Frm85_LM_SEARCH_3_LOGIC = "="
End If
If Frm101.L46_Text = "Semua cawangan" Then

    Frm85_LM_SEARCH_5 = Null
    Frm85_LM_SEARCH_5_LOGIC = "<>"
    Frm85_LM_SEARCH_6 = Null
    Frm85_LM_SEARCH_6_LOGIC = "<>"
    
Else

    Frm85_LM_SEARCH_5 = Frm101.L46_Text
    Frm85_LM_SEARCH_5_LOGIC = "="
    Frm85_LM_SEARCH_6 = "HQ"
    Frm85_LM_SEARCH_6_LOGIC = "="
    
End If

Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh Belian
        .Columns("C").ColumnWidth = 20 'No. Siri Produk
        .Columns("D").ColumnWidth = 20 'Purity
        .Columns("E").ColumnWidth = 40 'Kategori Produk
        .Columns("F").ColumnWidth = 40 'Supplier
        .Columns("G").ColumnWidth = 20 'Berat (g)
        .Columns("H").ColumnWidth = 20 'Rate Penerimaan (RM/g)
        .Columns("I").ColumnWidth = 20 'Upah (RM)
        .Columns("J").ColumnWidth = 20 'Spread (%)
        .Columns("K").ColumnWidth = 20 'Harga Selepas Spread (RM)
        .Columns("L").ColumnWidth = 20 'Adjustment (RM)
        .Columns("M").ColumnWidth = 20 'Harga Belian (RM)
        .Columns("N").ColumnWidth = 20 'Upah Jualan (RM) : Pelanggan
        .Columns("O").ColumnWidth = 20 'Upah Jualan (RM) : Ahli Biasa
        .Columns("P").ColumnWidth = 20 'Upah Jualan (RM) : Silver
        .Columns("Q").ColumnWidth = 20 'Upah Jualan (RM) : Gold
        .Columns("R").ColumnWidth = 20 'Upah Jualan (RM) : Platinum
        .Columns("S").ColumnWidth = 0 'Upah Jualan (RM) : Master Dealer
        .Columns("T").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Pelanggan
        .Columns("U").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Member
        .Columns("V").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Pengedar
        .Columns("W").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : RAF
        .Columns("X").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Normal Dealer
        .Columns("Y").ColumnWidth = 0 'Tetapan Harga Jualan (RM) : Master Dealer
        .Columns("Z").ColumnWidth = 20 'Dulang
        .Columns("AA").ColumnWidth = 20 'Panjang
        .Columns("AB").ColumnWidth = 20 'Lebar
        .Columns("AC").ColumnWidth = 20 'Saiz
        .Columns("AD").ColumnWidth = 20 'No. Invoice Dari Supplier
        .Columns("AE").ColumnWidth = 20 'Code 1
        .Columns("AF").ColumnWidth = 20 'Code 2
        .Columns("AG").ColumnWidth = 40 'Cawangan
        .Columns("AH").ColumnWidth = 40 'No. Voucher
        
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        '#### Header Report ###
        If Frm101.L9_Text = 0 Then .Cells(7, 1) = "Report Belian Buyback / Trade In Bagi Purity [" & Frm101.L5_Text & "] & Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "]." 'Report Header"
        If Frm101.L9_Text = 1 Then .Cells(7, 1) = "Report Belian Buyback / Trade In Bagi Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "] Dari " & TM & " Hingga " & TA 'Report Header"

        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Belian"
        .Cells(8, 3) = "No. Siri Produk"
        .Cells(8, 4) = "Purity"
        .Cells(8, 5) = "Kategori Produk"
        .Cells(8, 6) = "Supplier"
        .Cells(8, 7) = "Berat (g)"
        .Cells(8, 8) = "Rate Penerimaan (RM/g)"
        .Cells(8, 9) = "Upah (RM)"
        .Cells(8, 10) = "Spread (%)"
        .Cells(8, 11) = "Harga Selepas Spread (RM)"
        .Cells(8, 12) = "Adjustment (RM)"
        .Cells(8, 13) = "Harga Belian (RM)"
        .Cells(8, 14) = "Upah Jualan (RM) : Pelanggan"
        .Cells(8, 15) = "Upah Jualan (RM) : Ahli Biasa"
        .Cells(8, 16) = "Upah Jualan (RM) : Silver"
        .Cells(8, 17) = "Upah Jualan (RM) : Gold"
        .Cells(8, 18) = "Upah Jualan (RM) : Platinum"
        .Cells(8, 19) = "Upah Jualan (RM) : Master Dealer"
        .Cells(8, 20) = "Tetapan Harga Jualan (RM) : Pelanggan"
        .Cells(8, 21) = "Tetapan Harga Jualan (RM) : Ahli Biasa"
        .Cells(8, 22) = "Tetapan Harga Jualan (RM) : Silver"
        .Cells(8, 23) = "Tetapan Harga Jualan (RM) : Gold"
        .Cells(8, 24) = "Tetapan Harga Jualan (RM) : Platinum"
        .Cells(8, 25) = "Tetapan Harga Jualan (RM) : Master Dealer"
        .Cells(8, 26) = "Dulang"
        .Cells(8, 27) = "Panjang"
        .Cells(8, 28) = "Lebar"
        .Cells(8, 29) = "Saiz"
        .Cells(8, 30) = "No. Invoice"
        .Cells(8, 31) = "Code 1"
        .Cells(8, 32) = "Code 2"
        .Cells(8, 33) = "Cawangan"
        .Cells(8, 34) = "No. Voucher"
        
        For i = 1 To 34
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

        If Frm101.L9_Text = 0 Then rs.Open "select * from Data_Database where kod_purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select * from Data_Database where kod_purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh_belian) Then .Cells(8 + x, 2) = "'" & rs!tarikh_belian 'Tarikh Belian
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 3) = rs!no_siri_Produk 'No. Siri Produk
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!purity) Then .Cells(8 + x, 4) = rs!purity 'Purity
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 5) = rs!kategori_Produk 'Kategori Produk
            
            If Not IsNull(rs!nama_Supplier) Then .Cells(8 + x, 6) = rs!nama_Supplier 'Nama Supplier
            
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!Berat) Then
                .Cells(8 + x, 7) = rs!Berat 'Berat (g)
            Else
                .Cells(8 + x, 7) = "0.00" 'Berat (g)
            End If
            
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!kos_Belian_Gram) Then
                .Cells(8 + x, 8) = Format(rs!kos_Belian_Gram, "#,##0.00") 'Rate Penerimaan (RM/g)
            Else
                .Cells(8 + x, 8) = "0.00" 'Rate Penerimaan (RM/g)
            End If
            
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!UPAH) Then
                .Cells(8 + x, 9) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
            Else
                .Cells(8 + x, 9) = "0.00" 'Upah (RM)
            End If
            
            .Cells(8 + x, 10).NumberFormat = "#,##0.00"
            .Cells(8 + x, 10).HorizontalAlignment = xlRight
            If Not IsNull(rs!SpreadValue) Then
                .Cells(8 + x, 10) = rs!SpreadValue 'Spread (%)
            Else
                .Cells(8 + x, 10) = "0.00" 'Spread (%)
            End If
            
            .Cells(8 + x, 11).NumberFormat = "#,##0.00"
            .Cells(8 + x, 11).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_lepas_spread) Then
                .Cells(8 + x, 11) = Format(rs!harga_lepas_spread, "#,##0.00") 'Harga Selepas Spread (RM)
            Else
                .Cells(8 + x, 11) = "0.00" 'Harga Selepas Spread (RM)
            End If
            
            .Cells(8 + x, 12).NumberFormat = "#,##0.00"
            .Cells(8 + x, 12).HorizontalAlignment = xlRight
            If Not IsNull(rs!adjustment) Then
                .Cells(8 + x, 12) = Format(rs!adjustment, "#,##0.00") 'Adjustment (RM)
            Else
                .Cells(8 + x, 12) = "0.00" 'Adjustment (RM)
            End If
            
            .Cells(8 + x, 13).NumberFormat = "#,##0.00"
            .Cells(8 + x, 13).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_item) Then
                .Cells(8 + x, 13) = Format(rs!harga_item, "#,##0.00") 'Harga Belian (RM) : Tidak Campur Cukai GST
            Else
                .Cells(8 + x, 13) = "0.00" 'Harga Belian (RM) : Tidak Campur Cukai GST
            End If
            
            .Cells(8 + x, 14).NumberFormat = "#,##0.00"
            .Cells(8 + x, 14).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Jualan) Then
                .Cells(8 + x, 14) = Format(rs!Upah_Jualan, "#,##0.00") 'Upah Jualan (RM) : Pelanggan
            Else
                .Cells(8 + x, 14) = "0.00" 'Upah Jualan (RM) : Pelanggan
            End If
            
            .Cells(8 + x, 15).NumberFormat = "#,##0.00"
            .Cells(8 + x, 15).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Member) Then
                .Cells(8 + x, 15) = Format(rs!Upah_Member, "#,##0.00") 'Upah Jualan (RM) : Ahli Biasa
            Else
                .Cells(8 + x, 15) = "0.00" 'Upah Jualan (RM) : Ahli Biasa
            End If
            
            .Cells(8 + x, 16).NumberFormat = "#,##0.00"
            .Cells(8 + x, 16).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Pengedar) Then
                .Cells(8 + x, 16) = Format(rs!Upah_Pengedar, "#,##0.00") 'Upah Jualan (RM) : Silver
            Else
                .Cells(8 + x, 16) = "0.00" 'Upah Jualan (RM) : Silver
            End If
            
            .Cells(8 + x, 17).NumberFormat = "#,##0.00"
            .Cells(8 + x, 17).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_RAF) Then
                .Cells(8 + x, 17) = Format(rs!Upah_RAF, "#,##0.00") 'Upah Jualan (RM) : Gold
            Else
                .Cells(8 + x, 17) = "0.00" 'Upah Jualan (RM) : Gold
            End If
            
            .Cells(8 + x, 18).NumberFormat = "#,##0.00"
            .Cells(8 + x, 18).HorizontalAlignment = xlRight
            If Not IsNull(rs!upah_normal_dealer) Then
                .Cells(8 + x, 18) = Format(rs!upah_normal_dealer, "#,##0.00") 'Upah Jualan (RM) : Platinum
            Else
                .Cells(8 + x, 18) = "0.00" 'Upah Jualan (RM) : Platinum
            End If
            
            .Cells(8 + x, 19).NumberFormat = "#,##0.00"
            .Cells(8 + x, 19).HorizontalAlignment = xlRight
            If Not IsNull(rs!upah_master_dealer) Then
                .Cells(8 + x, 19) = Format(rs!upah_master_dealer, "#,##0.00") 'Upah Jualan (RM) : Master Dealer
            Else
                .Cells(8 + x, 19) = "0.00" 'Upah Jualan (RM) : Master Dealer
            End If
            
            .Cells(8 + x, 20).NumberFormat = "#,##0.00"
            .Cells(8 + x, 20).HorizontalAlignment = xlRight
            If Not IsNull(rs!code_Supplier) Then
                .Cells(8 + x, 20) = Format(rs!code_Supplier, "#,##0.00") 'Tetapan Harga Jualan (RM) : Pelanggan
            Else
                .Cells(8 + x, 20) = "0.00" 'Tetapan Harga Jualan (RM) : Pelanggan
            End If
            
            .Cells(8 + x, 21).NumberFormat = "#,##0.00"
            .Cells(8 + x, 21).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_Member) Then
                .Cells(8 + x, 21) = Format(rs!HargaJualan_Member, "#,##0.00") 'Tetapan Harga Jualan (RM) : Member
            Else
                .Cells(8 + x, 21) = "0.00" 'Tetapan Harga Jualan (RM) : Member
            End If
            
            .Cells(8 + x, 22).NumberFormat = "#,##0.00"
            .Cells(8 + x, 22).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_Pengedar) Then
                .Cells(8 + x, 22) = Format(rs!HargaJualan_Pengedar, "#,##0.00") 'Tetapan Harga Jualan (RM) : Pengedar
            Else
                .Cells(8 + x, 22) = "0.00" 'Tetapan Harga Jualan (RM) : Pengedar
            End If
            
            .Cells(8 + x, 23).NumberFormat = "#,##0.00"
            .Cells(8 + x, 23).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_RAF) Then
                .Cells(8 + x, 23) = Format(rs!HargaJualan_RAF, "#,##0.00") 'Tetapan Harga Jualan (RM) : RAF
            Else
                .Cells(8 + x, 23) = "0.00" 'Tetapan Harga Jualan (RM) : RAF
            End If
            
            .Cells(8 + x, 24).NumberFormat = "#,##0.00"
            .Cells(8 + x, 24).HorizontalAlignment = xlRight
            If Not IsNull(rs!hargajualan_normal_dealer) Then
                .Cells(8 + x, 24) = Format(rs!hargajualan_normal_dealer, "#,##0.00") 'Tetapan Harga Jualan (RM) : Normal Dealer
            Else
                .Cells(8 + x, 24) = "0.00" 'Tetapan Harga Jualan (RM) : Normal Dealer
            End If
            
            .Cells(8 + x, 25).NumberFormat = "#,##0.00"
            .Cells(8 + x, 25).HorizontalAlignment = xlRight
            If Not IsNull(rs!hargajualan_master_dealer) Then
                .Cells(8 + x, 25) = Format(rs!hargajualan_master_dealer, "#,##0.00") 'Tetapan Harga Jualan (RM) : Master Dealer
            Else
                .Cells(8 + x, 25) = "0.00" 'Tetapan Harga Jualan (RM) : Master Dealer
            End If
            
            If Not IsNull(rs!dulang) Then
                .Cells(8 + x, 26) = rs!dulang 'Dulang
                .Cells(8 + x, 26).HorizontalAlignment = xlCenter
            End If
                
            If Not IsNull(rs!dimension_Panjang) Then
                .Cells(8 + x, 27) = rs!dimension_Panjang 'Panjang
                .Cells(8 + x, 27).HorizontalAlignment = xlCenter
            End If
    
            If Not IsNull(rs!dimension_Lebar) Then
                .Cells(8 + x, 28) = rs!dimension_Lebar 'Lebar
                .Cells(8 + x, 28).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!dimension_Saiz) Then
                .Cells(8 + x, 29) = rs!dimension_Saiz 'Tebal
                .Cells(8 + x, 29).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!bill_No_Belian) Then .Cells(8 + x, 30) = "'" & rs!bill_No_Belian 'No Invoice Dari Supplier
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!code1) Then .Cells(8 + x, 31) = "'" & rs!code1 'No Invoice Dari Supplier
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!code2) Then .Cells(8 + x, 32) = "'" & rs!code2 'No Invoice Dari Supplier
            
            If Not IsNull(rs!cawangan) Then .Cells(8 + x, 33) = rs!cawangan 'Cawangan
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!bill_No_Trade_In) Then .Cells(8 + x, 34) = rs!bill_No_Trade_In
            
            For Col = 1 To 34
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
    
        Y = x + 2

        '#### Jumlah Bilangan Barang Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select COUNT(no_siri_Produk) from Data_Database where kod_purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select COUNT(no_siri_Produk) from Data_Database where kod_purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
                
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Bilangan : " & rs(0)
        Else
            .Cells(8 + Y, 1) = "Bilangan : " & 0
        End If
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Bilangan : " & 0
        End If
        '#### Jumlah Bilangan Barang Keseluruhan #### - End
        
        Y = Y + 1
        
        '#### Jumlah Berat Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select SUM(Berat) from Data_Database where kod_purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select SUM(Berat) from Data_Database where kod_purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
                
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        '#### Jumlah Berat Keseluruhan #### - End
        
        Y = Y + 1
        
        '#### Jumlah Modal Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select SUM(harga_item) from Data_Database where kod_purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select SUM(harga_item) from Data_Database where kod_purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "0.00"
        End If
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "0.00"
        End If
        '#### Jumlah Modal Keseluruhan #### - End
        
        Y = Y + 4
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
End Sub
Sub Frm85_bk_trade_invoice_overall()
'on error resume next
Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh Belian
        .Columns("C").ColumnWidth = 20 'No. Siri Produk
        .Columns("D").ColumnWidth = 20 'Purity
        .Columns("E").ColumnWidth = 40 'Kategori Produk
        .Columns("F").ColumnWidth = 40 'Supplier
        .Columns("G").ColumnWidth = 20 'Berat (g)
        .Columns("H").ColumnWidth = 20 'Rate Penerimaan (RM/g)
        .Columns("I").ColumnWidth = 20 'Upah (RM)
        .Columns("J").ColumnWidth = 20 'Spread (%)
        .Columns("K").ColumnWidth = 20 'Harga Selepas Spread (RM)
        .Columns("L").ColumnWidth = 20 'Adjustment (RM)
        .Columns("M").ColumnWidth = 20 'Harga Belian (RM)
        .Columns("N").ColumnWidth = 20 'Upah Jualan (RM) : Pelanggan
        .Columns("O").ColumnWidth = 20 'Upah Jualan (RM) : Ahli Biasa
        .Columns("P").ColumnWidth = 20 'Upah Jualan (RM) : Silver
        .Columns("Q").ColumnWidth = 20 'Upah Jualan (RM) : Gold
        .Columns("R").ColumnWidth = 20 'Upah Jualan (RM) : Platinum
        .Columns("S").ColumnWidth = 0 'Upah Jualan (RM) : Master Dealer
        .Columns("T").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Pelanggan
        .Columns("U").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Member
        .Columns("V").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Pengedar
        .Columns("W").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : RAF
        .Columns("X").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Normal Dealer
        .Columns("Y").ColumnWidth = 0 'Tetapan Harga Jualan (RM) : Master Dealer
        .Columns("Z").ColumnWidth = 20 'Dulang
        .Columns("AA").ColumnWidth = 20 'Panjang
        .Columns("AB").ColumnWidth = 20 'Lebar
        .Columns("AC").ColumnWidth = 20 'Saiz
        .Columns("AD").ColumnWidth = 20 'No. Invoice Dari Supplier
        .Columns("AE").ColumnWidth = 20 'Code 1
        .Columns("AF").ColumnWidth = 20 'Code 2
        .Columns("AG").ColumnWidth = 40 'Cawangan
        
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        '#### Header Report ###
        .Cells(7, 1) = Frm85.L19_Text 'Report Header"
       
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Belian"
        .Cells(8, 3) = "No. Siri Produk"
        .Cells(8, 4) = "Purity"
        .Cells(8, 5) = "Kategori Produk"
        .Cells(8, 6) = "Supplier"
        .Cells(8, 7) = "Berat (g)"
        .Cells(8, 8) = "Rate Penerimaan (RM/g)"
        .Cells(8, 9) = "Upah (RM)"
        .Cells(8, 10) = "Spread (%)"
        .Cells(8, 11) = "Harga Selepas Spread (RM)"
        .Cells(8, 12) = "Adjustment (RM)"
        .Cells(8, 13) = "Harga Belian (RM)"
        .Cells(8, 14) = "Upah Jualan (RM) : Pelanggan"
        .Cells(8, 15) = "Upah Jualan (RM) : Ahli Biasa"
        .Cells(8, 16) = "Upah Jualan (RM) : Silver"
        .Cells(8, 17) = "Upah Jualan (RM) : Gold"
        .Cells(8, 18) = "Upah Jualan (RM) : Platinum"
        .Cells(8, 19) = "Upah Jualan (RM) : Master Dealer"
        .Cells(8, 20) = "Tetapan Harga Jualan (RM) : Pelanggan"
        .Cells(8, 21) = "Tetapan Harga Jualan (RM) : Ahli Biasa"
        .Cells(8, 22) = "Tetapan Harga Jualan (RM) : Silver"
        .Cells(8, 23) = "Tetapan Harga Jualan (RM) : Gold"
        .Cells(8, 24) = "Tetapan Harga Jualan (RM) : Platinum"
        .Cells(8, 25) = "Tetapan Harga Jualan (RM) : Master Dealer"
        .Cells(8, 26) = "Dulang"
        .Cells(8, 27) = "Panjang"
        .Cells(8, 28) = "Lebar"
        .Cells(8, 29) = "Saiz"
        .Cells(8, 30) = "No. Invoice"
        .Cells(8, 31) = "Code 1"
        .Cells(8, 32) = "Code 2"
        .Cells(8, 33) = "Cawangan"
        
        For i = 1 To 33
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
        
        If Frm101.L46_Text = "Semua cawangan" Then
        
            Frm85_LM_SEARCH_5 = Null
            Frm85_LM_SEARCH_5_LOGIC = "<>"
            Frm85_LM_SEARCH_6 = Null
            Frm85_LM_SEARCH_6_LOGIC = "<>"
            
        Else
        
            Frm85_LM_SEARCH_5 = Frm101.L46_Text
            Frm85_LM_SEARCH_5_LOGIC = "="
            Frm85_LM_SEARCH_6 = "HQ"
            Frm85_LM_SEARCH_6_LOGIC = "="
            
        End If
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where bill_No_Trade_In='" & UCase(Frm101.L5_Text) & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh_belian) Then .Cells(8 + x, 2) = "'" & rs!tarikh_belian 'Tarikh Belian
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 3) = rs!no_siri_Produk 'No. Siri Produk
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!purity) Then .Cells(8 + x, 4) = rs!purity 'Purity
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 5) = rs!kategori_Produk 'Kategori Produk
            
            If Not IsNull(rs!nama_Supplier) Then .Cells(8 + x, 6) = rs!nama_Supplier 'Nama Supplier
            
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!Berat) Then
                .Cells(8 + x, 7) = Format(rs!Berat, "#,##0.00") 'Berat (g)
            Else
                .Cells(8 + x, 7) = "0.00" 'Berat (g)
            End If
            
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!kos_Belian_Gram) Then
                .Cells(8 + x, 8) = Format(rs!kos_Belian_Gram, "#,##0.00") 'Rate Penerimaan (RM/g)
            Else
                .Cells(8 + x, 8) = "0.00" 'Rate Penerimaan (RM/g)
            End If
            
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!UPAH) Then
                .Cells(8 + x, 9) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
            Else
                .Cells(8 + x, 9) = "0.00" 'Upah (RM)
            End If
            
            .Cells(8 + x, 10).NumberFormat = "#,##0.00"
            .Cells(8 + x, 10).HorizontalAlignment = xlRight
            If Not IsNull(rs!SpreadValue) Then
                .Cells(8 + x, 10) = rs!SpreadValue 'Spread (%)
            Else
                .Cells(8 + x, 10) = "0.00" 'Spread (%)
            End If
            
            .Cells(8 + x, 11).NumberFormat = "#,##0.00"
            .Cells(8 + x, 11).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_lepas_spread) Then
                .Cells(8 + x, 11) = Format(rs!harga_lepas_spread, "#,##0.00") 'Harga Selepas Spread (RM)
            Else
                .Cells(8 + x, 11) = "0.00" 'Harga Selepas Spread (RM)
            End If
            
            .Cells(8 + x, 12).NumberFormat = "#,##0.00"
            .Cells(8 + x, 12).HorizontalAlignment = xlRight
            If Not IsNull(rs!adjustment) Then
                .Cells(8 + x, 12) = Format(rs!adjustment, "#,##0.00") 'Adjustment (RM)
            Else
                .Cells(8 + x, 12) = "0.00" 'Adjustment (RM)
            End If
            
            .Cells(8 + x, 13).NumberFormat = "#,##0.00"
            .Cells(8 + x, 13).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_item) Then
                .Cells(8 + x, 13) = Format(rs!harga_item, "#,##0.00") 'Harga Belian (RM) : Tidak Campur Cukai GST
            Else
                .Cells(8 + x, 13) = "0.00" 'Harga Belian (RM) : Tidak Campur Cukai GST
            End If
            
            .Cells(8 + x, 14).NumberFormat = "#,##0.00"
            .Cells(8 + x, 14).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Jualan) Then
                .Cells(8 + x, 14) = Format(rs!Upah_Jualan, "#,##0.00") 'Upah Jualan (RM) : Pelanggan
            Else
                .Cells(8 + x, 14) = "0.00" 'Upah Jualan (RM) : Pelanggan
            End If
            
            .Cells(8 + x, 15).NumberFormat = "#,##0.00"
            .Cells(8 + x, 15).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Member) Then
                .Cells(8 + x, 15) = Format(rs!Upah_Member, "#,##0.00") 'Upah Jualan (RM) : Ahli Biasa
            Else
                .Cells(8 + x, 15) = "0.00" 'Upah Jualan (RM) : Ahli Biasa
            End If
            
            .Cells(8 + x, 16).NumberFormat = "#,##0.00"
            .Cells(8 + x, 16).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Pengedar) Then
                .Cells(8 + x, 16) = Format(rs!Upah_Pengedar, "#,##0.00") 'Upah Jualan (RM) : Silver
            Else
                .Cells(8 + x, 16) = "0.00" 'Upah Jualan (RM) : Silver
            End If
            
            .Cells(8 + x, 17).NumberFormat = "#,##0.00"
            .Cells(8 + x, 17).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_RAF) Then
                .Cells(8 + x, 17) = Format(rs!Upah_RAF, "#,##0.00") 'Upah Jualan (RM) : Gold
            Else
                .Cells(8 + x, 17) = "0.00" 'Upah Jualan (RM) : Gold
            End If
            
            .Cells(8 + x, 18).NumberFormat = "#,##0.00"
            .Cells(8 + x, 18).HorizontalAlignment = xlRight
            If Not IsNull(rs!upah_normal_dealer) Then
                .Cells(8 + x, 18) = Format(rs!upah_normal_dealer, "#,##0.00") 'Upah Jualan (RM) : Platinum
            Else
                .Cells(8 + x, 18) = "0.00" 'Upah Jualan (RM) : Platinum
            End If
            
            .Cells(8 + x, 19).NumberFormat = "#,##0.00"
            .Cells(8 + x, 19).HorizontalAlignment = xlRight
            If Not IsNull(rs!upah_master_dealer) Then
                .Cells(8 + x, 19) = Format(rs!upah_master_dealer, "#,##0.00") 'Upah Jualan (RM) : Master Dealer
            Else
                .Cells(8 + x, 19) = "0.00" 'Upah Jualan (RM) : Master Dealer
            End If
            
            .Cells(8 + x, 20).NumberFormat = "#,##0.00"
            .Cells(8 + x, 20).HorizontalAlignment = xlRight
            If Not IsNull(rs!code_Supplier) Then
                .Cells(8 + x, 20) = Format(rs!code_Supplier, "#,##0.00") 'Tetapan Harga Jualan (RM) : Pelanggan
            Else
                .Cells(8 + x, 20) = "0.00" 'Tetapan Harga Jualan (RM) : Pelanggan
            End If
            
            .Cells(8 + x, 21).NumberFormat = "#,##0.00"
            .Cells(8 + x, 21).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_Member) Then
                .Cells(8 + x, 21) = Format(rs!HargaJualan_Member, "#,##0.00") 'Tetapan Harga Jualan (RM) : Member
            Else
                .Cells(8 + x, 21) = "0.00" 'Tetapan Harga Jualan (RM) : Member
            End If
            
            .Cells(8 + x, 22).NumberFormat = "#,##0.00"
            .Cells(8 + x, 22).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_Pengedar) Then
                .Cells(8 + x, 22) = Format(rs!HargaJualan_Pengedar, "#,##0.00") 'Tetapan Harga Jualan (RM) : Pengedar
            Else
                .Cells(8 + x, 22) = "0.00" 'Tetapan Harga Jualan (RM) : Pengedar
            End If
            
            .Cells(8 + x, 23).NumberFormat = "#,##0.00"
            .Cells(8 + x, 23).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_RAF) Then
                .Cells(8 + x, 23) = Format(rs!HargaJualan_RAF, "#,##0.00") 'Tetapan Harga Jualan (RM) : RAF
            Else
                .Cells(8 + x, 23) = "0.00" 'Tetapan Harga Jualan (RM) : RAF
            End If
            
            .Cells(8 + x, 24).NumberFormat = "#,##0.00"
            .Cells(8 + x, 24).HorizontalAlignment = xlRight
            If Not IsNull(rs!hargajualan_normal_dealer) Then
                .Cells(8 + x, 24) = Format(rs!hargajualan_normal_dealer, "#,##0.00") 'Tetapan Harga Jualan (RM) : Normal Dealer
            Else
                .Cells(8 + x, 24) = "0.00" 'Tetapan Harga Jualan (RM) : Normal Dealer
            End If
            
            .Cells(8 + x, 25).NumberFormat = "#,##0.00"
            .Cells(8 + x, 25).HorizontalAlignment = xlRight
            If Not IsNull(rs!hargajualan_master_dealer) Then
                .Cells(8 + x, 25) = Format(rs!hargajualan_master_dealer, "#,##0.00") 'Tetapan Harga Jualan (RM) : Master Dealer
            Else
                .Cells(8 + x, 25) = "0.00" 'Tetapan Harga Jualan (RM) : Master Dealer
            End If
            
            If Not IsNull(rs!dulang) Then
                .Cells(8 + x, 26) = rs!dulang 'Dulang
                .Cells(8 + x, 26).HorizontalAlignment = xlCenter
            End If
                
            If Not IsNull(rs!dimension_Panjang) Then
                .Cells(8 + x, 27) = rs!dimension_Panjang 'Panjang
                .Cells(8 + x, 27).HorizontalAlignment = xlCenter
            End If
    
            If Not IsNull(rs!dimension_Lebar) Then
                .Cells(8 + x, 28) = rs!dimension_Lebar 'Lebar
                .Cells(8 + x, 28).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!dimension_Saiz) Then
                .Cells(8 + x, 29) = rs!dimension_Saiz 'Tebal
                .Cells(8 + x, 29).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!bill_No_Belian) Then .Cells(8 + x, 30) = "'" & rs!bill_No_Belian 'No Invoice Dari Supplier
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!code1) Then .Cells(8 + x, 31) = "'" & rs!code1 'No Invoice Dari Supplier
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!code2) Then .Cells(8 + x, 32) = "'" & rs!code2 'Code 1
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!cawangan) Then .Cells(8 + x, 33) = rs!cawangan 'Cawangan
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            For Col = 1 To 33
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
    
        Y = x + 2

        '#### Jumlah Bilangan Barang Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select COUNT(no_siri_Produk) from Data_Database where bill_No_Trade_In='" & UCase(Frm101.L5_Text) & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Bilangan : " & rs(0)
        Else
            .Cells(8 + Y, 1) = "Bilangan : " & 0
        End If
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Bilangan : " & 0
        End If
        '#### Jumlah Bilangan Barang Keseluruhan #### - End
        
        Y = Y + 1
        
        '#### Jumlah Berat Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(Berat) from Data_Database where bill_No_Trade_In='" & UCase(Frm101.L5_Text) & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        '#### Jumlah Berat Keseluruhan #### - End
        
        Y = Y + 1
        
        '#### Jumlah Modal Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(harga_item) from Data_Database where bill_No_Trade_In='" & UCase(Frm101.L5_Text) & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "0.00"
        End If
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "0.00"
        End If
        '#### Jumlah Modal Keseluruhan #### - End
        
        Y = Y + 4
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
End Sub
Sub Frm85_bk_trade_siri_overall()
'on error resume next
Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If
If Frm101.L6_Text = "Semua Kategori Produk" Then
    Frm85_LM_SEARCH_2 = Null
    Frm85_LM_SEARCH_2_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_2 = Frm101.L6_Text
    Frm85_LM_SEARCH_2_LOGIC = "="
End If
If Frm101.L34_Text = "Semua Dulang" Then
    Frm85_LM_SEARCH_3 = Null
    Frm85_LM_SEARCH_3_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_3 = Frm101.L34_Text
    Frm85_LM_SEARCH_3_LOGIC = "="
End If
If Frm101.L46_Text = "Semua cawangan" Then

    Frm85_LM_SEARCH_5 = Null
    Frm85_LM_SEARCH_5_LOGIC = "<>"
    Frm85_LM_SEARCH_6 = Null
    Frm85_LM_SEARCH_6_LOGIC = "<>"
    
Else

    Frm85_LM_SEARCH_5 = Frm101.L46_Text
    Frm85_LM_SEARCH_5_LOGIC = "="
    Frm85_LM_SEARCH_6 = "HQ"
    Frm85_LM_SEARCH_6_LOGIC = "="
    
End If
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh Belian
        .Columns("C").ColumnWidth = 20 'No. Siri Produk
        .Columns("D").ColumnWidth = 20 'Purity
        .Columns("E").ColumnWidth = 40 'Kategori Produk
        .Columns("F").ColumnWidth = 40 'Supplier
        .Columns("G").ColumnWidth = 20 'Berat (g)
        .Columns("H").ColumnWidth = 20 'Rate Penerimaan (RM/g)
        .Columns("I").ColumnWidth = 20 'Upah (RM)
        .Columns("J").ColumnWidth = 20 'Spread (%)
        .Columns("K").ColumnWidth = 20 'Harga Selepas Spread (RM)
        .Columns("L").ColumnWidth = 20 'Adjustment (RM)
        .Columns("M").ColumnWidth = 20 'Harga Belian (RM)
        .Columns("N").ColumnWidth = 20 'Upah Jualan (RM) : Pelanggan
        .Columns("O").ColumnWidth = 20 'Upah Jualan (RM) : Ahli Biasa
        .Columns("P").ColumnWidth = 20 'Upah Jualan (RM) : Silver
        .Columns("Q").ColumnWidth = 20 'Upah Jualan (RM) : Gold
        .Columns("R").ColumnWidth = 20 'Upah Jualan (RM) : Platinum
        .Columns("S").ColumnWidth = 0 'Upah Jualan (RM) : Master Dealer
        .Columns("T").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Pelanggan
        .Columns("U").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Member
        .Columns("V").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Pengedar
        .Columns("W").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : RAF
        .Columns("X").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Normal Dealer
        .Columns("Y").ColumnWidth = 0 'Tetapan Harga Jualan (RM) : Master Dealer
        .Columns("Z").ColumnWidth = 20 'Dulang
        .Columns("AA").ColumnWidth = 20 'Panjang
        .Columns("AB").ColumnWidth = 20 'Lebar
        .Columns("AC").ColumnWidth = 20 'Saiz
        .Columns("AD").ColumnWidth = 20 'No. Invoice Dari Supplier
        .Columns("AE").ColumnWidth = 20 'Code 1
        .Columns("AF").ColumnWidth = 20 'Code 2
        .Columns("AG").ColumnWidth = 40 'Cawangan
        
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        '#### Header Report ###
        .Cells(7, 1) = Frm85.L19_Text 'Report Header"

        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Belian"
        .Cells(8, 3) = "No. Siri Produk"
        .Cells(8, 4) = "Purity"
        .Cells(8, 5) = "Kategori Produk"
        .Cells(8, 6) = "Supplier"
        .Cells(8, 7) = "Berat (g)"
        .Cells(8, 8) = "Rate Penerimaan (RM/g)"
        .Cells(8, 9) = "Upah (RM)"
        .Cells(8, 10) = "Spread (%)"
        .Cells(8, 11) = "Harga Selepas Spread (RM)"
        .Cells(8, 12) = "Adjustment (RM)"
        .Cells(8, 13) = "Harga Belian (RM)"
        .Cells(8, 14) = "Upah Jualan (RM) : Pelanggan"
        .Cells(8, 15) = "Upah Jualan (RM) : Ahli Biasa"
        .Cells(8, 16) = "Upah Jualan (RM) : Silver"
        .Cells(8, 17) = "Upah Jualan (RM) : Gold"
        .Cells(8, 18) = "Upah Jualan (RM) : Platinum"
        .Cells(8, 19) = "Upah Jualan (RM) : Master Dealer"
        .Cells(8, 20) = "Tetapan Harga Jualan (RM) : Pelanggan"
        .Cells(8, 21) = "Tetapan Harga Jualan (RM) : Ahli Biasa"
        .Cells(8, 22) = "Tetapan Harga Jualan (RM) : Silver"
        .Cells(8, 23) = "Tetapan Harga Jualan (RM) : Gold"
        .Cells(8, 24) = "Tetapan Harga Jualan (RM) : Platinum"
        .Cells(8, 25) = "Tetapan Harga Jualan (RM) : Master Dealer"
        .Cells(8, 26) = "Dulang"
        .Cells(8, 27) = "Panjang"
        .Cells(8, 28) = "Lebar"
        .Cells(8, 29) = "Saiz"
        .Cells(8, 30) = "No. Invoice"
        .Cells(8, 31) = "Code 1"
        .Cells(8, 32) = "Code 2"
        .Cells(8, 33) = "Cawangan"
        
        For i = 1 To 33
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh_belian) Then .Cells(8 + x, 2) = "'" & rs!tarikh_belian 'Tarikh Belian
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 3) = rs!no_siri_Produk 'No. Siri Produk
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!purity) Then .Cells(8 + x, 4) = rs!purity 'Purity
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 5) = rs!kategori_Produk 'Kategori Produk
            
            If Not IsNull(rs!nama_Supplier) Then .Cells(8 + x, 6) = rs!nama_Supplier 'Nama Supplier
            
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!Berat) Then
                .Cells(8 + x, 7) = Format(rs!Berat, "#,##0.00") 'Berat (g)
            Else
                .Cells(8 + x, 7) = "0.00" 'Berat (g)
            End If
            
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!kos_Belian_Gram) Then
                .Cells(8 + x, 8) = Format(rs!kos_Belian_Gram, "#,##0.00") 'Rate Penerimaan (RM/g)
            Else
                .Cells(8 + x, 8) = "0.00" 'Rate Penerimaan (RM/g)
            End If
            
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!UPAH) Then
                .Cells(8 + x, 9) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
            Else
                .Cells(8 + x, 9) = "0.00" 'Upah (RM)
            End If
            
            .Cells(8 + x, 10).NumberFormat = "#,##0.00"
            .Cells(8 + x, 10).HorizontalAlignment = xlRight
            If Not IsNull(rs!SpreadValue) Then
                .Cells(8 + x, 10) = rs!SpreadValue 'Spread (%)
            Else
                .Cells(8 + x, 10) = "0.00" 'Spread (%)
            End If
            
            .Cells(8 + x, 11).NumberFormat = "#,##0.00"
            .Cells(8 + x, 11).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_lepas_spread) Then
                .Cells(8 + x, 11) = Format(rs!harga_lepas_spread, "#,##0.00") 'Harga Selepas Spread (RM)
            Else
                .Cells(8 + x, 11) = "0.00" 'Harga Selepas Spread (RM)
            End If
            
            .Cells(8 + x, 12).NumberFormat = "#,##0.00"
            .Cells(8 + x, 12).HorizontalAlignment = xlRight
            If Not IsNull(rs!adjustment) Then
                .Cells(8 + x, 12) = Format(rs!adjustment, "#,##0.00") 'Adjustment (RM)
            Else
                .Cells(8 + x, 12) = "0.00" 'Adjustment (RM)
            End If
            
            .Cells(8 + x, 13).NumberFormat = "#,##0.00"
            .Cells(8 + x, 13).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_item) Then
                .Cells(8 + x, 13) = Format(rs!harga_item, "#,##0.00") 'Harga Belian (RM) : Tidak Campur Cukai GST
            Else
                .Cells(8 + x, 13) = "0.00" 'Harga Belian (RM) : Tidak Campur Cukai GST
            End If
            
            .Cells(8 + x, 14).NumberFormat = "#,##0.00"
            .Cells(8 + x, 14).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Jualan) Then
                .Cells(8 + x, 14) = Format(rs!Upah_Jualan, "#,##0.00") 'Upah Jualan (RM) : Pelanggan
            Else
                .Cells(8 + x, 14) = "0.00" 'Upah Jualan (RM) : Pelanggan
            End If
            
            .Cells(8 + x, 15).NumberFormat = "#,##0.00"
            .Cells(8 + x, 15).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Member) Then
                .Cells(8 + x, 15) = Format(rs!Upah_Member, "#,##0.00") 'Upah Jualan (RM) : Ahli Biasa
            Else
                .Cells(8 + x, 15) = "0.00" 'Upah Jualan (RM) : Ahli Biasa
            End If
            
            .Cells(8 + x, 16).NumberFormat = "#,##0.00"
            .Cells(8 + x, 16).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Pengedar) Then
                .Cells(8 + x, 16) = Format(rs!Upah_Pengedar, "#,##0.00") 'Upah Jualan (RM) : Silver
            Else
                .Cells(8 + x, 16) = "0.00" 'Upah Jualan (RM) : Silver
            End If
            
            .Cells(8 + x, 17).NumberFormat = "#,##0.00"
            .Cells(8 + x, 17).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_RAF) Then
                .Cells(8 + x, 17) = Format(rs!Upah_RAF, "#,##0.00") 'Upah Jualan (RM) : Gold
            Else
                .Cells(8 + x, 17) = "0.00" 'Upah Jualan (RM) : Gold
            End If
            
            .Cells(8 + x, 18).NumberFormat = "#,##0.00"
            .Cells(8 + x, 18).HorizontalAlignment = xlRight
            If Not IsNull(rs!upah_normal_dealer) Then
                .Cells(8 + x, 18) = Format(rs!upah_normal_dealer, "#,##0.00") 'Upah Jualan (RM) : Platinum
            Else
                .Cells(8 + x, 18) = "0.00" 'Upah Jualan (RM) : Platinum
            End If
            
            .Cells(8 + x, 19).NumberFormat = "#,##0.00"
            .Cells(8 + x, 19).HorizontalAlignment = xlRight
            If Not IsNull(rs!upah_master_dealer) Then
                .Cells(8 + x, 19) = Format(rs!upah_master_dealer, "#,##0.00") 'Upah Jualan (RM) : Master Dealer
            Else
                .Cells(8 + x, 19) = "0.00" 'Upah Jualan (RM) : Master Dealer
            End If
            
            .Cells(8 + x, 20).NumberFormat = "#,##0.00"
            .Cells(8 + x, 20).HorizontalAlignment = xlRight
            If Not IsNull(rs!code_Supplier) Then
                .Cells(8 + x, 20) = Format(rs!code_Supplier, "#,##0.00") 'Tetapan Harga Jualan (RM) : Pelanggan
            Else
                .Cells(8 + x, 20) = "0.00" 'Tetapan Harga Jualan (RM) : Pelanggan
            End If
            
            .Cells(8 + x, 21).NumberFormat = "#,##0.00"
            .Cells(8 + x, 21).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_Member) Then
                .Cells(8 + x, 21) = Format(rs!HargaJualan_Member, "#,##0.00") 'Tetapan Harga Jualan (RM) : Member
            Else
                .Cells(8 + x, 21) = "0.00" 'Tetapan Harga Jualan (RM) : Member
            End If
            
            .Cells(8 + x, 22).NumberFormat = "#,##0.00"
            .Cells(8 + x, 22).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_Pengedar) Then
                .Cells(8 + x, 22) = Format(rs!HargaJualan_Pengedar, "#,##0.00") 'Tetapan Harga Jualan (RM) : Pengedar
            Else
                .Cells(8 + x, 22) = "0.00" 'Tetapan Harga Jualan (RM) : Pengedar
            End If
            
            .Cells(8 + x, 23).NumberFormat = "#,##0.00"
            .Cells(8 + x, 23).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_RAF) Then
                .Cells(8 + x, 23) = Format(rs!HargaJualan_RAF, "#,##0.00") 'Tetapan Harga Jualan (RM) : RAF
            Else
                .Cells(8 + x, 23) = "0.00" 'Tetapan Harga Jualan (RM) : RAF
            End If
            
            .Cells(8 + x, 24).NumberFormat = "#,##0.00"
            .Cells(8 + x, 24).HorizontalAlignment = xlRight
            If Not IsNull(rs!hargajualan_normal_dealer) Then
                .Cells(8 + x, 24) = Format(rs!hargajualan_normal_dealer, "#,##0.00") 'Tetapan Harga Jualan (RM) : Normal Dealer
            Else
                .Cells(8 + x, 24) = "0.00" 'Tetapan Harga Jualan (RM) : Normal Dealer
            End If
            
            .Cells(8 + x, 25).NumberFormat = "#,##0.00"
            .Cells(8 + x, 25).HorizontalAlignment = xlRight
            If Not IsNull(rs!hargajualan_master_dealer) Then
                .Cells(8 + x, 25) = Format(rs!hargajualan_master_dealer, "#,##0.00") 'Tetapan Harga Jualan (RM) : Master Dealer
            Else
                .Cells(8 + x, 25) = "0.00" 'Tetapan Harga Jualan (RM) : Master Dealer
            End If
            
            If Not IsNull(rs!dulang) Then
                .Cells(8 + x, 26) = rs!dulang 'Dulang
                .Cells(8 + x, 26).HorizontalAlignment = xlCenter
            End If
                
            If Not IsNull(rs!dimension_Panjang) Then
                .Cells(8 + x, 27) = rs!dimension_Panjang 'Panjang
                .Cells(8 + x, 27).HorizontalAlignment = xlCenter
            End If
    
            If Not IsNull(rs!dimension_Lebar) Then
                .Cells(8 + x, 28) = rs!dimension_Lebar 'Lebar
                .Cells(8 + x, 28).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!dimension_Saiz) Then
                .Cells(8 + x, 29) = rs!dimension_Saiz 'Tebal
                .Cells(8 + x, 29).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!bill_No_Belian) Then .Cells(8 + x, 30) = "'" & rs!bill_No_Belian 'No Invoice Dari Supplier
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!code1) Then .Cells(8 + x, 31) = "'" & rs!code1 'No Invoice Dari Supplier
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!code2) Then .Cells(8 + x, 32) = "'" & rs!code2 'Code 1
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!cawangan) Then .Cells(8 + x, 33) = rs!cawangan 'Cawangan
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            For Col = 1 To 33
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
    
        Y = x + 2

        '#### Jumlah Bilangan Barang Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select COUNT(no_siri_Produk) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
                
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Bilangan : " & rs(0)
        Else
            .Cells(8 + Y, 1) = "Bilangan : " & 0
        End If
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Bilangan : " & 0
        End If
        '#### Jumlah Bilangan Barang Keseluruhan #### - End
        
        Y = Y + 1
        
        '#### Jumlah Berat Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(Berat) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
                
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        '#### Jumlah Berat Keseluruhan #### - End
        
        Y = Y + 1
        
        '#### Jumlah Modal Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(harga_item) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND StatusItem <> 0 AND (receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "0.00"
        End If
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "0.00"
        End If
        '#### Jumlah Modal Keseluruhan #### - End
        
        Y = Y + 4
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
End Sub
Sub Frm85_stok_overall()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    
    x = 0
    
    Frm85_LM_BERAT = 0
    Frm85_LM_HARGA = 0
    
    If Frm101.L9_Text = 1 Then
        TM = Frm101.L7_Text 'Tarikh Mula
        TA = Frm101.L8_Text 'Tarikh Akhir
    End If
    If Frm101.L5_Text = "Semua Purity" Then
        Frm85_LM_SEARCH_1 = Null
        Frm85_LM_SEARCH_1_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_1 = Frm101.L5_Text
        Frm85_LM_SEARCH_1_LOGIC = "="
    End If
    If Frm101.L6_Text = "Semua Kategori Produk" Then
        Frm85_LM_SEARCH_2 = Null
        Frm85_LM_SEARCH_2_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_2 = Frm101.L6_Text
        Frm85_LM_SEARCH_2_LOGIC = "="
    End If
    If Frm101.L34_Text = "Semua Dulang" Then
        Frm85_LM_SEARCH_3 = Null
        Frm85_LM_SEARCH_3_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_3 = Frm101.L34_Text
        Frm85_LM_SEARCH_3_LOGIC = "="
    End If
    If Frm101.L37_Text = "Semua Supplier" Then
        Frm85_LM_SEARCH_4 = Null
        Frm85_LM_SEARCH_4_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_4 = Frm101.L37_Text
        Frm85_LM_SEARCH_4_LOGIC = "="
    End If
    If Frm101.L46_Text = "Semua cawangan" Then
    
        Frm85_LM_SEARCH_5 = Null
        Frm85_LM_SEARCH_5_LOGIC = "<>"
        Frm85_LM_SEARCH_6 = Null
        Frm85_LM_SEARCH_6_LOGIC = "<>"
        
    Else
    
        Frm85_LM_SEARCH_5 = Frm101.L46_Text
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

    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh Belian
        .Columns("C").ColumnWidth = 20 'No. Siri Produk
        .Columns("D").ColumnWidth = 20 'Purity
        .Columns("E").ColumnWidth = 40 'Kategori Produk
        .Columns("F").ColumnWidth = 40 'Supplier
        .Columns("G").ColumnWidth = 20 'Berat (g)
        .Columns("H").ColumnWidth = 20 'Rate Penerimaan (RM/g)
        .Columns("I").ColumnWidth = 20 'Upah (RM)
        .Columns("J").ColumnWidth = 20 'Spread (%)
        .Columns("K").ColumnWidth = 20 'Harga Selepas Spread (RM)
        .Columns("L").ColumnWidth = 20 'Adjustment (RM)
        .Columns("M").ColumnWidth = 20 'Harga Belian (RM)
        .Columns("N").ColumnWidth = 20 'Upah Jualan (RM) : Pelanggan
        .Columns("O").ColumnWidth = 20 'Upah Jualan (RM) : Ahli
        .Columns("P").ColumnWidth = 20 'Upah Jualan (RM) : Silver
        .Columns("Q").ColumnWidth = 20 'Upah Jualan (RM) : Gold
        .Columns("R").ColumnWidth = 20 'Upah Jualan (RM) : Platinum
        .Columns("S").ColumnWidth = 0 'Upah Jualan (RM) : Master Dealer
        .Columns("T").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Pelanggan
        .Columns("U").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Ahli
        .Columns("V").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Silver
        .Columns("W").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Gold
        .Columns("X").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Platinum
        .Columns("Y").ColumnWidth = 0 'Tetapan Harga Jualan (RM) : Master Dealer
        .Columns("Z").ColumnWidth = 20 'Dulang
        .Columns("AA").ColumnWidth = 20 'Panjang
        .Columns("AB").ColumnWidth = 20 'Lebar
        .Columns("AC").ColumnWidth = 20 'Saiz
        .Columns("AD").ColumnWidth = 20 'Invoice Supplier
        .Columns("AE").ColumnWidth = 20 'Code 1
        .Columns("AF").ColumnWidth = 20 'Code 2
        .Columns("AG").ColumnWidth = 40 'Cawangan
        
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
            LM_NAMA_HEADER = "HQ"
            
        Else
            
            LM_NAMA_HEADER = MDI_frm1.L20_Text
            
        End If
                
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        '#### Header Report ###
        If Frm101.L9_Text = 0 Then .Cells(7, 1) = "Report Stok Bagi Supplier [" & Frm101.L37_Text & "] , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "]." 'Report Header"
        If Frm101.L9_Text = 1 Then .Cells(7, 1) = "Report Stok Bagi Supplier [" & Frm101.L37_Text & "] , Purity [" & Frm101.L5_Text & "] , Kategori Produk [" & Frm101.L6_Text & "] , Dulang [" & Frm101.L34_Text & "] , Cawangan [" & Frm101.L46_Text & "] Dari " & TM & " Hingga " & TA 'Report Header"
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Belian"
        .Cells(8, 3) = "No. Siri Produk"
        .Cells(8, 4) = "Purity"
        .Cells(8, 5) = "Kategori Produk"
        .Cells(8, 6) = "Supplier"
        .Cells(8, 7) = "Berat (g)"
        .Cells(8, 8) = "Rate Penerimaan (RM/g)"
        .Cells(8, 9) = "Upah (RM)"
        .Cells(8, 10) = "Spread (%)"
        .Cells(8, 11) = "Harga Selepas Spread (RM)"
        .Cells(8, 12) = "Adjustment (RM)"
        .Cells(8, 13) = "Harga Belian (RM)"
        .Cells(8, 14) = "Upah Jualan (RM) : Pelanggan"
        .Cells(8, 15) = "Upah Jualan (RM) : Ahli"
        .Cells(8, 16) = "Upah Jualan (RM) : Silver"
        .Cells(8, 17) = "Upah Jualan (RM) : Gold"
        .Cells(8, 18) = "Upah Jualan (RM) : Platinum"
        .Cells(8, 19) = "Upah Jualan (RM) : Master Dealer"
        .Cells(8, 20) = "Tetapan Harga Jualan (RM) : Pelanggan"
        .Cells(8, 21) = "Tetapan Harga Jualan (RM) : Ahli"
        .Cells(8, 22) = "Tetapan Harga Jualan (RM) : Silver"
        .Cells(8, 23) = "Tetapan Harga Jualan (RM) : Gold"
        .Cells(8, 24) = "Tetapan Harga Jualan (RM) : Platinum"
        .Cells(8, 25) = "Tetapan Harga Jualan (RM) : Master Dealer"
        .Cells(8, 26) = "Dulang"
        .Cells(8, 27) = "Panjang"
        .Cells(8, 28) = "Lebar"
        .Cells(8, 29) = "Saiz"
        .Cells(8, 30) = "Invoice Supplier"
        .Cells(8, 31) = "Code 1"
        .Cells(8, 32) = "Code 2"
        .Cells(8, 33) = "Cawangan"
        
        For i = 1 To 33
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select * from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "') AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND (receiving_Status='" & 0 & "' OR receiving_Status='" & 1 & "' OR receiving_Status='" & 4 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select * from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "') AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 0 & "' OR receiving_Status='" & 1 & "' OR receiving_Status='" & 4 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh_belian) Then .Cells(8 + x, 2) = "'" & rs!tarikh_belian 'Tarikh Belian
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 3) = rs!no_siri_Produk 'No. Siri Produk
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!purity) Then .Cells(8 + x, 4) = rs!purity 'Purity
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 5) = rs!kategori_Produk 'Kategori Produk
            
            If Not IsNull(rs!nama_Supplier) Then .Cells(8 + x, 6) = rs!nama_Supplier 'Nama Supplier
            
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!Berat) Then
                .Cells(8 + x, 7) = Format(rs!Berat, "#,##0.00") 'Berat (g)
            Else
                .Cells(8 + x, 7) = "0.00" 'Berat (g)
            End If
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!kos_Belian_Gram) Then
                .Cells(8 + x, 8) = Format(rs!kos_Belian_Gram, "#,##0.00") 'Rate Penerimaan (RM/g)
            Else
                .Cells(8 + x, 8) = "0.00" 'Rate Penerimaan (RM/g)
            End If
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!UPAH) Then
                .Cells(8 + x, 9) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
            Else
                .Cells(8 + x, 9) = "0.00" 'Upah (RM)
            End If
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 10).HorizontalAlignment = xlRight
            If Not IsNull(rs!SpreadValue) Then
                .Cells(8 + x, 10) = rs!SpreadValue 'Spread (%)
            Else
                .Cells(8 + x, 10) = "0.00" 'Spread (%)
            End If
            .Cells(8 + x, 10).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 11).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_lepas_spread) Then
                .Cells(8 + x, 11) = Format(rs!harga_lepas_spread, "#,##0.00") 'Harga Selepas Spread (RM)
            Else
                .Cells(8 + x, 11) = "0.00" 'Harga Selepas Spread (RM)
            End If
            .Cells(8 + x, 11).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 12).HorizontalAlignment = xlRight
            If Not IsNull(rs!adjustment) Then
                .Cells(8 + x, 12) = Format(rs!adjustment, "#,##0.00") 'Adjustment (RM)
            Else
                .Cells(8 + x, 12) = "0.00" 'Adjustment (RM)
            End If
            .Cells(8 + x, 12).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 13).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_item) Then
                .Cells(8 + x, 13) = Format(rs!harga_item, "#,##0.00") 'Harga Belian (RM) : Tidak Campur Cukai GST
            Else
                .Cells(8 + x, 13) = "0.00" 'Harga Belian (RM) : Tidak Campur Cukai GST
            End If
            .Cells(8 + x, 13).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 14).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Jualan) Then
                .Cells(8 + x, 14) = Format(rs!Upah_Jualan, "#,##0.00") 'Upah Jualan (RM) : Pelanggan
            Else
                .Cells(8 + x, 14) = "0.00" 'Upah Jualan (RM) : Pelanggan
            End If
            .Cells(8 + x, 14).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 15).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Member) Then
                .Cells(8 + x, 15) = Format(rs!Upah_Member, "#,##0.00") 'Upah Jualan (RM) : Ahli Biasa
            Else
                .Cells(8 + x, 15) = "0.00" 'Upah Jualan (RM) : Ahli Biasa
            End If
            .Cells(8 + x, 15).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 16).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Pengedar) Then
                .Cells(8 + x, 16) = Format(rs!Upah_Pengedar, "#,##0.00") 'Upah Jualan (RM) : Silver
            Else
                .Cells(8 + x, 16) = "0.00" 'Upah Jualan (RM) : Silver
            End If
            .Cells(8 + x, 16).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 17).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_RAF) Then
                .Cells(8 + x, 17) = Format(rs!Upah_RAF, "#,##0.00") 'Upah Jualan (RM) : Gold
            Else
                .Cells(8 + x, 17) = "0.00" 'Upah Jualan (RM) : Gold
            End If
            .Cells(8 + x, 17).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 18).HorizontalAlignment = xlRight
            If Not IsNull(rs!upah_normal_dealer) Then
                .Cells(8 + x, 18) = Format(rs!upah_normal_dealer, "#,##0.00") 'Upah Jualan (RM) : Platinum
            Else
                .Cells(8 + x, 18) = "0.00" 'Upah Jualan (RM) : Platinum
            End If
            .Cells(8 + x, 18).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 19).HorizontalAlignment = xlRight
            If Not IsNull(rs!upah_master_dealer) Then
                .Cells(8 + x, 19) = Format(rs!upah_master_dealer, "#,##0.00") 'Upah Jualan (RM) : Master Dealer
            Else
                .Cells(8 + x, 19) = "0.00" 'Upah Jualan (RM) : Master Dealer
            End If
            .Cells(8 + x, 19).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 20).HorizontalAlignment = xlRight
            If Not IsNull(rs!code_Supplier) Then
                .Cells(8 + x, 20) = Format(rs!code_Supplier, "#,##0.00") 'Tetapan Harga Jualan (RM) : Pelanggan
            Else
                .Cells(8 + x, 20) = "0.00" 'Tetapan Harga Jualan (RM) : Pelanggan
            End If
            .Cells(8 + x, 20).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 21).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_Member) Then
                .Cells(8 + x, 21) = Format(rs!HargaJualan_Member, "#,##0.00") 'Tetapan Harga Jualan (RM) : Member
            Else
                .Cells(8 + x, 21) = "0.00" 'Tetapan Harga Jualan (RM) : Member
            End If
            .Cells(8 + x, 21).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 22).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_Pengedar) Then
                .Cells(8 + x, 22) = Format(rs!HargaJualan_Pengedar, "#,##0.00") 'Tetapan Harga Jualan (RM) : Pengedar
            Else
                .Cells(8 + x, 22) = "0.00" 'Tetapan Harga Jualan (RM) : Pengedar
            End If
            .Cells(8 + x, 22).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 23).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_RAF) Then
                .Cells(8 + x, 23) = Format(rs!HargaJualan_RAF, "#,##0.00") 'Tetapan Harga Jualan (RM) : RAF
            Else
                .Cells(8 + x, 23) = "0.00" 'Tetapan Harga Jualan (RM) : RAF
            End If
            .Cells(8 + x, 23).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 24).HorizontalAlignment = xlRight
            If Not IsNull(rs!hargajualan_normal_dealer) Then
                .Cells(8 + x, 24) = Format(rs!hargajualan_normal_dealer, "#,##0.00") 'Tetapan Harga Jualan (RM) : Normal Dealer
            Else
                .Cells(8 + x, 24) = "0.00" 'Tetapan Harga Jualan (RM) : Normal Dealer
            End If
            .Cells(8 + x, 24).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 25).HorizontalAlignment = xlRight
            If Not IsNull(rs!hargajualan_master_dealer) Then
                .Cells(8 + x, 25) = Format(rs!hargajualan_master_dealer, "#,##0.00") 'Tetapan Harga Jualan (RM) : Master Dealer
            Else
                .Cells(8 + x, 25) = "0.00" 'Tetapan Harga Jualan (RM) : Master Dealer
            End If
            .Cells(8 + x, 25).NumberFormat = "#,##0.00"
            
            If Not IsNull(rs!dulang) Then
                .Cells(8 + x, 26) = rs!dulang 'Dulang
                .Cells(8 + x, 26).HorizontalAlignment = xlCenter
            End If
                
            If Not IsNull(rs!dimension_Panjang) Then
                .Cells(8 + x, 27) = rs!dimension_Panjang 'Panjang
                .Cells(8 + x, 27).HorizontalAlignment = xlCenter
            End If
    
            If Not IsNull(rs!dimension_Lebar) Then
                .Cells(8 + x, 28) = rs!dimension_Lebar 'Lebar
                .Cells(8 + x, 28).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!dimension_Saiz) Then
                .Cells(8 + x, 29) = rs!dimension_Saiz 'Tebal
                .Cells(8 + x, 29).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!bill_No_Belian) Then .Cells(8 + x, 30) = "'" & rs!bill_No_Belian 'No Invoice Dari Supplier
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!code1) Then .Cells(8 + x, 31) = "'" & rs!code1 'No Invoice Dari Supplier
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!code2) Then .Cells(8 + x, 32) = "'" & rs!code2 'No Invoice Dari Supplier
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!cawangan) Then .Cells(8 + x, 33) = rs!cawangan 'Cawangan
            
            For Col = 1 To 33
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 2
        .Cells(8 + Y, 1) = "Bilangan : " & x 'Total Barang
        Y = Y + 1
        
        '#### Jumlah Berat Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select SUM(Berat) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "') AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND (receiving_Status='" & 0 & "' OR receiving_Status='" & 1 & "' OR receiving_Status='" & 4 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select SUM(Berat) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "') AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 0 & "' OR receiving_Status='" & 1 & "' OR receiving_Status='" & 4 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        '#### Jumlah Berat Keseluruhan #### - End
        
        Y = Y + 1
        
        '#### Jumlah Modal Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select SUM(harga_item) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "') AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND (receiving_Status='" & 0 & "' OR receiving_Status='" & 1 & "' OR receiving_Status='" & 4 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select SUM(harga_item) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "') AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND (receiving_Status='" & 0 & "' OR receiving_Status='" & 1 & "' OR receiving_Status='" & 4 & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "0.00"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "RM 0.00"
        End If
        '#### Jumlah Modal Keseluruhan #### - End
        
        Y = Y + 4
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
End Sub
Sub Frm85_stok_siri()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    
    x = 0
    
    Frm85_LM_BERAT = 0
    Frm85_LM_HARGA = 0
    
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

    If Frm101.L9_Text = 1 Then
        TM = Frm101.L7_Text 'Tarikh Mula
        TA = Frm101.L8_Text 'Tarikh Akhir
    End If
    If Frm101.L5_Text = "Semua Purity" Then
        Frm85_LM_SEARCH_1 = Null
        Frm85_LM_SEARCH_1_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_1 = Frm101.L5_Text
        Frm85_LM_SEARCH_1_LOGIC = "="
    End If
    If Frm101.L6_Text = "Semua Kategori Produk" Then
        Frm85_LM_SEARCH_2 = Null
        Frm85_LM_SEARCH_2_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_2 = Frm101.L6_Text
        Frm85_LM_SEARCH_2_LOGIC = "="
    End If
    If Frm101.L34_Text = "Semua Dulang" Then
        Frm85_LM_SEARCH_3 = Null
        Frm85_LM_SEARCH_3_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_3 = Frm101.L34_Text
        Frm85_LM_SEARCH_3_LOGIC = "="
    End If
    If Frm101.L46_Text = "Semua cawangan" Then
    
        Frm85_LM_SEARCH_5 = Null
        Frm85_LM_SEARCH_5_LOGIC = "<>"
        Frm85_LM_SEARCH_6 = Null
        Frm85_LM_SEARCH_6_LOGIC = "<>"
        
    Else
    
        Frm85_LM_SEARCH_5 = Frm101.L46_Text
        Frm85_LM_SEARCH_5_LOGIC = "="
        Frm85_LM_SEARCH_6 = "HQ"
        Frm85_LM_SEARCH_6_LOGIC = "="
        
    End If
    
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh Belian
        .Columns("C").ColumnWidth = 20 'No. Siri Produk
        .Columns("D").ColumnWidth = 20 'Purity
        .Columns("E").ColumnWidth = 40 'Kategori Produk
        .Columns("F").ColumnWidth = 40 'Supplier
        .Columns("G").ColumnWidth = 20 'Berat (g)
        .Columns("H").ColumnWidth = 20 'Rate Penerimaan (RM/g)
        .Columns("I").ColumnWidth = 20 'Upah (RM)
        .Columns("J").ColumnWidth = 20 'Spread (%)
        .Columns("K").ColumnWidth = 20 'Harga Selepas Spread (RM)
        .Columns("L").ColumnWidth = 20 'Adjustment (RM)
        .Columns("M").ColumnWidth = 20 'Harga Belian (RM)
        .Columns("N").ColumnWidth = 20 'Upah Jualan (RM) : Pelanggan
        .Columns("O").ColumnWidth = 20 'Upah Jualan (RM) : Ahli
        .Columns("P").ColumnWidth = 20 'Upah Jualan (RM) : Silver
        .Columns("Q").ColumnWidth = 20 'Upah Jualan (RM) : Gold
        .Columns("R").ColumnWidth = 20 'Upah Jualan (RM) : Platinum
        .Columns("S").ColumnWidth = 0 'Upah Jualan (RM) : Master Dealer
        .Columns("T").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Pelanggan
        .Columns("U").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Ahli
        .Columns("V").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Silver
        .Columns("W").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Gold
        .Columns("X").ColumnWidth = 20 'Tetapan Harga Jualan (RM) : Platinum
        .Columns("Y").ColumnWidth = 0 'Tetapan Harga Jualan (RM) : Master Dealer
        .Columns("Z").ColumnWidth = 20 'Dulang
        .Columns("AA").ColumnWidth = 20 'Panjang
        .Columns("AB").ColumnWidth = 20 'Lebar
        .Columns("AC").ColumnWidth = 20 'Saiz
        .Columns("AD").ColumnWidth = 20 'Invoice Supplier
        .Columns("AE").ColumnWidth = 20 'Code 1
        .Columns("AF").ColumnWidth = 20 'Code 2
        .Columns("AG").ColumnWidth = 40 'Cawangan
        
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        '#### Header Report ###
        .Cells(7, 1) = Frm85.L23_Text 'Report Header"
               
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Belian"
        .Cells(8, 3) = "No. Siri Produk"
        .Cells(8, 4) = "Purity"
        .Cells(8, 5) = "Kategori Produk"
        .Cells(8, 6) = "Supplier"
        .Cells(8, 7) = "Berat (g)"
        .Cells(8, 8) = "Rate Penerimaan (RM/g)"
        .Cells(8, 9) = "Upah (RM)"
        .Cells(8, 10) = "Spread (%)"
        .Cells(8, 11) = "Harga Selepas Spread (RM)"
        .Cells(8, 12) = "Adjustment (RM)"
        .Cells(8, 13) = "Harga Belian (RM)"
        .Cells(8, 14) = "Upah Jualan (RM) : Pelanggan"
        .Cells(8, 15) = "Upah Jualan (RM) : Ahli"
        .Cells(8, 16) = "Upah Jualan (RM) : Silver"
        .Cells(8, 17) = "Upah Jualan (RM) : Gold"
        .Cells(8, 18) = "Upah Jualan (RM) : Platinum"
        .Cells(8, 19) = "Upah Jualan (RM) : Master Dealer"
        .Cells(8, 20) = "Tetapan Harga Jualan (RM) : Pelanggan"
        .Cells(8, 21) = "Tetapan Harga Jualan (RM) : Ahli"
        .Cells(8, 22) = "Tetapan Harga Jualan (RM) : Silver"
        .Cells(8, 23) = "Tetapan Harga Jualan (RM) : Gold"
        .Cells(8, 24) = "Tetapan Harga Jualan (RM) : Platinum"
        .Cells(8, 25) = "Tetapan Harga Jualan (RM) : Master Dealer"
        .Cells(8, 26) = "Dulang"
        .Cells(8, 27) = "Panjang"
        .Cells(8, 28) = "Lebar"
        .Cells(8, 29) = "Saiz"
        .Cells(8, 30) = "Invoice Supplier"
        .Cells(8, 31) = "Code 1"
        .Cells(8, 32) = "Code 2"
        .Cells(8, 33) = "Cawangan"
        
        For i = 1 To 33
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh_belian) Then .Cells(8 + x, 2) = "'" & rs!tarikh_belian 'Tarikh Belian
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 3) = rs!no_siri_Produk 'No. Siri Produk
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!purity) Then .Cells(8 + x, 4) = rs!purity 'Purity
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 5) = rs!kategori_Produk 'Kategori Produk
            
            If Not IsNull(rs!nama_Supplier) Then .Cells(8 + x, 6) = rs!nama_Supplier 'Nama Supplier
            
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!Berat) Then
                .Cells(8 + x, 7) = Format(rs!Berat, "#,##0.00") 'Berat (g)
            Else
                .Cells(8 + x, 7) = "0.00" 'Berat (g)
            End If
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!kos_Belian_Gram) Then
                .Cells(8 + x, 8) = Format(rs!kos_Belian_Gram, "#,##0.00") 'Rate Penerimaan (RM/g)
            Else
                .Cells(8 + x, 8) = "0.00" 'Rate Penerimaan (RM/g)
            End If
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!UPAH) Then
                .Cells(8 + x, 9) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
            Else
                .Cells(8 + x, 9) = "0.00" 'Upah (RM)
            End If
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 10).HorizontalAlignment = xlRight
            If Not IsNull(rs!SpreadValue) Then
                .Cells(8 + x, 10) = rs!SpreadValue 'Spread (%)
            Else
                .Cells(8 + x, 10) = "0.00" 'Spread (%)
            End If
            .Cells(8 + x, 10).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 11).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_lepas_spread) Then
                .Cells(8 + x, 11) = Format(rs!harga_lepas_spread, "#,##0.00") 'Harga Selepas Spread (RM)
            Else
                .Cells(8 + x, 11) = "0.00" 'Harga Selepas Spread (RM)
            End If
            .Cells(8 + x, 11).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 12).HorizontalAlignment = xlRight
            If Not IsNull(rs!adjustment) Then
                .Cells(8 + x, 12) = Format(rs!adjustment, "#,##0.00") 'Adjustment (RM)
            Else
                .Cells(8 + x, 12) = "0.00" 'Adjustment (RM)
            End If
            .Cells(8 + x, 12).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 13).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_item) Then
                .Cells(8 + x, 13) = Format(rs!harga_item, "#,##0.00") 'Harga Belian (RM) : Tidak Campur Cukai GST
            Else
                .Cells(8 + x, 13) = "0.00" 'Harga Belian (RM) : Tidak Campur Cukai GST
            End If
            .Cells(8 + x, 13).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 14).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Jualan) Then
                .Cells(8 + x, 14) = Format(rs!Upah_Jualan, "#,##0.00") 'Upah Jualan (RM) : Pelanggan
            Else
                .Cells(8 + x, 14) = "0.00" 'Upah Jualan (RM) : Pelanggan
            End If
            .Cells(8 + x, 14).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 15).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Member) Then
                .Cells(8 + x, 15) = Format(rs!Upah_Member, "#,##0.00") 'Upah Jualan (RM) : Ahli Biasa
            Else
                .Cells(8 + x, 15) = "0.00" 'Upah Jualan (RM) : Ahli Biasa
            End If
            .Cells(8 + x, 15).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 16).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_Pengedar) Then
                .Cells(8 + x, 16) = Format(rs!Upah_Pengedar, "#,##0.00") 'Upah Jualan (RM) : Silver
            Else
                .Cells(8 + x, 16) = "0.00" 'Upah Jualan (RM) : Silver
            End If
            .Cells(8 + x, 16).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 17).HorizontalAlignment = xlRight
            If Not IsNull(rs!Upah_RAF) Then
                .Cells(8 + x, 17) = Format(rs!Upah_RAF, "#,##0.00") 'Upah Jualan (RM) : Gold
            Else
                .Cells(8 + x, 17) = "0.00" 'Upah Jualan (RM) : Gold
            End If
            .Cells(8 + x, 17).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 18).HorizontalAlignment = xlRight
            If Not IsNull(rs!upah_normal_dealer) Then
                .Cells(8 + x, 18) = Format(rs!upah_normal_dealer, "#,##0.00") 'Upah Jualan (RM) : Platinum
            Else
                .Cells(8 + x, 18) = "0.00" 'Upah Jualan (RM) : Platinum
            End If
            .Cells(8 + x, 18).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 19).HorizontalAlignment = xlRight
            If Not IsNull(rs!upah_master_dealer) Then
                .Cells(8 + x, 19) = Format(rs!upah_master_dealer, "#,##0.00") 'Upah Jualan (RM) : Master Dealer
            Else
                .Cells(8 + x, 19) = "0.00" 'Upah Jualan (RM) : Master Dealer
            End If
            .Cells(8 + x, 19).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 20).HorizontalAlignment = xlRight
            If Not IsNull(rs!code_Supplier) Then
                .Cells(8 + x, 20) = Format(rs!code_Supplier, "#,##0.00") 'Tetapan Harga Jualan (RM) : Pelanggan
            Else
                .Cells(8 + x, 20) = "0.00" 'Tetapan Harga Jualan (RM) : Pelanggan
            End If
            .Cells(8 + x, 20).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 21).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_Member) Then
                .Cells(8 + x, 21) = Format(rs!HargaJualan_Member, "#,##0.00") 'Tetapan Harga Jualan (RM) : Member
            Else
                .Cells(8 + x, 21) = "0.00" 'Tetapan Harga Jualan (RM) : Member
            End If
            .Cells(8 + x, 21).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 22).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_Pengedar) Then
                .Cells(8 + x, 22) = Format(rs!HargaJualan_Pengedar, "#,##0.00") 'Tetapan Harga Jualan (RM) : Pengedar
            Else
                .Cells(8 + x, 22) = "0.00" 'Tetapan Harga Jualan (RM) : Pengedar
            End If
            .Cells(8 + x, 22).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 23).HorizontalAlignment = xlRight
            If Not IsNull(rs!HargaJualan_RAF) Then
                .Cells(8 + x, 23) = Format(rs!HargaJualan_RAF, "#,##0.00") 'Tetapan Harga Jualan (RM) : RAF
            Else
                .Cells(8 + x, 23) = "0.00" 'Tetapan Harga Jualan (RM) : RAF
            End If
            .Cells(8 + x, 23).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 24).HorizontalAlignment = xlRight
            If Not IsNull(rs!hargajualan_normal_dealer) Then
                .Cells(8 + x, 24) = Format(rs!hargajualan_normal_dealer, "#,##0.00") 'Tetapan Harga Jualan (RM) : Normal Dealer
            Else
                .Cells(8 + x, 24) = "0.00" 'Tetapan Harga Jualan (RM) : Normal Dealer
            End If
            .Cells(8 + x, 24).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 25).HorizontalAlignment = xlRight
            If Not IsNull(rs!hargajualan_master_dealer) Then
                .Cells(8 + x, 25) = Format(rs!hargajualan_master_dealer, "#,##0.00") 'Tetapan Harga Jualan (RM) : Master Dealer
            Else
                .Cells(8 + x, 25) = "0.00" 'Tetapan Harga Jualan (RM) : Master Dealer
            End If
            .Cells(8 + x, 25).NumberFormat = "#,##0.00"
            
            If Not IsNull(rs!dulang) Then
                .Cells(8 + x, 26) = rs!dulang 'Dulang
                .Cells(8 + x, 26).HorizontalAlignment = xlCenter
            End If
                
            If Not IsNull(rs!dimension_Panjang) Then
                .Cells(8 + x, 27) = rs!dimension_Panjang 'Panjang
                .Cells(8 + x, 27).HorizontalAlignment = xlCenter
            End If
    
            If Not IsNull(rs!dimension_Lebar) Then
                .Cells(8 + x, 28) = rs!dimension_Lebar 'Lebar
                .Cells(8 + x, 28).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!dimension_Saiz) Then
                .Cells(8 + x, 29) = rs!dimension_Saiz 'Tebal
                .Cells(8 + x, 29).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!bill_No_Belian) Then .Cells(8 + x, 30) = "'" & rs!bill_No_Belian 'No Invoice Dari Supplier
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!code1) Then .Cells(8 + x, 31) = "'" & rs!code1 'No Invoice Dari Supplier
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!code2) Then .Cells(8 + x, 32) = "'" & rs!code2 'No Invoice Dari Supplier
            '.Cells(8 + x, 30).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!cawangan) Then .Cells(8 + x, 33) = rs!cawangan 'Cawangan
            
            For Col = 1 To 33
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 2
        .Cells(8 + Y, 1) = "Bilangan : " & x 'Total Barang
        Y = Y + 1
        
        '#### Jumlah Berat Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(Berat) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
                        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        '#### Jumlah Berat Keseluruhan #### - End
        
        Y = Y + 1
        
        '#### Jumlah Modal Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(harga_item) from Data_Database where (gst_ari_nashi " & Frm85_LM_SEARCH_10_LOGIC & "'" & Frm85_LM_SEARCH_10 & "' OR gst_ari_nashi " & Frm85_LM_SEARCH_11_LOGIC & "'" & Frm85_LM_SEARCH_11 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "') order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "0.00"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "RM 0.00"
        End If
        '#### Jumlah Modal Keseluruhan #### - End
        
        Y = Y + 4
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
End Sub
Sub Frm85_potong_overall()
'On Error Resume Next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    
    x = 0
    
    Frm85_LM_BERAT = 0
    Frm85_LM_HARGA = 0
    
    If Frm101.L9_Text = 1 Then
        TM = Frm101.L7_Text 'Tarikh Mula
        TA = Frm101.L8_Text 'Tarikh Akhir
    End If
    If Frm101.L5_Text = "Semua Purity" Then
        Frm85_LM_SEARCH_1 = Null
        Frm85_LM_SEARCH_1_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_1 = Frm101.L5_Text
        Frm85_LM_SEARCH_1_LOGIC = "="
    End If
    If Frm101.L6_Text = "Semua Kategori Produk" Then
        Frm85_LM_SEARCH_2 = Null
        Frm85_LM_SEARCH_2_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_2 = Frm101.L6_Text
        Frm85_LM_SEARCH_2_LOGIC = "="
    End If
    If Frm101.L34_Text = "Semua Dulang" Then
        Frm85_LM_SEARCH_3 = Null
        Frm85_LM_SEARCH_3_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_3 = Frm101.L34_Text
        Frm85_LM_SEARCH_3_LOGIC = "="
    End If
    If Frm101.L46_Text = "Semua cawangan" Then
    
        Frm85_LM_SEARCH_5 = Null
        Frm85_LM_SEARCH_5_LOGIC = "<>"
        Frm85_LM_SEARCH_6 = Null
        Frm85_LM_SEARCH_6_LOGIC = "<>"
        
    Else
    
        Frm85_LM_SEARCH_5 = Frm101.L46_Text
        Frm85_LM_SEARCH_5_LOGIC = "="
        Frm85_LM_SEARCH_6 = "HQ"
        Frm85_LM_SEARCH_6_LOGIC = "="
        
    End If

    If Frm101.L47_Text = "Semua" Then
        Frm85_LM_SEARCH_20 = Null
        Frm85_LM_SEARCH_20_LOGIC = "<>"
    Else
        If InStr(1, Frm101.L47_Text, "  |  ") <> 0 Then
            Frm84_LM_EMP_NO = Split(Frm101.L47_Text, "  |  ")(1)
            Frm84_LM_EMP_NAMA = Split(Frm101.L47_Text, "  |  ")(0)
        End If
        Frm85_LM_SEARCH_20 = Frm84_LM_EMP_NAMA
        Frm85_LM_SEARCH_20_LOGIC = "="
    End If

    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh Belian
        .Columns("C").ColumnWidth = 20 'No. Siri Produk
        .Columns("D").ColumnWidth = 20 'Purity
        .Columns("E").ColumnWidth = 40 'Kategori Produk
        .Columns("F").ColumnWidth = 40 'Supplier
        .Columns("G").ColumnWidth = 20 'Berat Asal (g)
        .Columns("H").ColumnWidth = 20 'Susut Berat (g)
        .Columns("I").ColumnWidth = 20 'Baki Berat (g)
        .Columns("J").ColumnWidth = 20 'Dulang
        .Columns("K").ColumnWidth = 50 'Cawangan
        .Columns("L").ColumnWidth = 20 'Nama Pekerja
        
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        '#### Header Report ###
        .Cells(7, 1) = Frm85.L27_Text 'Report Header"
       
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Potong"
        .Cells(8, 3) = "No. Siri Produk"
        .Cells(8, 4) = "Purity"
        .Cells(8, 5) = "Kategori Produk"
        .Cells(8, 6) = "Supplier"
        .Cells(8, 7) = "Berat Asal (g)"
        .Cells(8, 8) = "Susut Berat (g)"
        .Cells(8, 9) = "Baki Berat (g)"
        .Cells(8, 10) = "Dulang"
        .Cells(8, 11) = "Cawangan"
        .Cells(8, 12) = "Nama Pekerja"
        
        For i = 1 To 12
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select * from Data_Database where nama_pekerja_potong " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND kod_purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND (StatusItem='" & 12 & "' OR StatusItem='" & 20 & "' OR StatusItem='" & 28 & "' OR StatusItem='" & 22 & "') order by tarikh_Jualan1 ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select * from Data_Database where nama_pekerja_potong " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND kod_purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh_Jualan1 BETWEEN '" & TM & "' AND '" & TA & "' AND (StatusItem='" & 12 & "' OR StatusItem='" & 20 & "' OR StatusItem='" & 28 & "' OR StatusItem='" & 22 & "') order by tarikh_Jualan1 ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh_jualan1) Then .Cells(8 + x, 2) = "'" & rs!tarikh_jualan1 'Tarikh Belian
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 3) = rs!no_siri_Produk 'No. Siri Produk
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!purity) Then .Cells(8 + x, 4) = rs!purity 'Purity
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 5) = rs!kategori_Produk 'Kategori Produk
            
            If Not IsNull(rs!nama_Supplier) Then .Cells(8 + x, 6) = rs!nama_Supplier 'Nama Supplier
            
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!Berat) Then
                .Cells(8 + x, 7) = Format(rs!Berat, "#,##0.00") 'Berat (g)
            Else
                .Cells(8 + x, 7) = "0.00" 'Berat (g)
            End If
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!susut_berat) Then
                .Cells(8 + x, 8) = Format(rs!susut_berat, "#,##0.00") 'Berat (g)
            Else
                .Cells(8 + x, 8) = "0.00" 'Berat (g)
            End If
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!beza_berat) Then
                .Cells(8 + x, 9) = Format(rs!beza_berat, "#,##0.00") 'Baki Berat (g)
            Else
                .Cells(8 + x, 9) = "0.00" 'Baki Berat (g)
            End If
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            
            If Not IsNull(rs!dulang) Then
                .Cells(8 + x, 10) = rs!dulang 'Dulang
                .Cells(8 + x, 10).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!cawangan) Then
                .Cells(8 + x, 11) = rs!cawangan 'Dulang
                '.Cells(8 + x, 11).HorizontalAlignment = xlCenter
            End If
            If Not IsNull(rs!nama_pekerja_potong) Then .Cells(8 + x, 12) = rs!nama_pekerja_potong
            
            For Col = 1 To 12
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 2
        .Cells(8 + Y, 1) = "Bilangan : " & x 'Total Barang
        Y = Y + 1
        
        '#### Jumlah Berat Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select SUM(beza_berat) from Data_Database where nama_pekerja_potong " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND kod_purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND (StatusItem='" & 12 & "' OR StatusItem='" & 20 & "' OR StatusItem='" & 28 & "' OR StatusItem='" & 22 & "') order by tarikh_Jualan1 ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select SUM(beza_berat) from Data_Database where nama_pekerja_potong " & Frm85_LM_SEARCH_20_LOGIC & "'" & Frm85_LM_SEARCH_20 & "' AND kod_purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh_Jualan1 BETWEEN '" & TM & "' AND '" & TA & "' AND (StatusItem='" & 12 & "' OR StatusItem='" & 20 & "' OR StatusItem='" & 28 & "' OR StatusItem='" & 22 & "') order by tarikh_Jualan1 ASC", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        '#### Jumlah Berat Keseluruhan #### - End
        
        Y = Y + 4
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
End Sub
Sub Frm85_potong_siri()
'On Error Resume Next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    
    x = 0
    
    Frm85_LM_BERAT = 0
    Frm85_LM_HARGA = 0
    
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh Belian
        .Columns("C").ColumnWidth = 20 'No. Siri Produk
        .Columns("D").ColumnWidth = 20 'Purity
        .Columns("E").ColumnWidth = 40 'Kategori Produk
        .Columns("F").ColumnWidth = 40 'Supplier
        .Columns("G").ColumnWidth = 20 'Berat Asal (g)
        .Columns("H").ColumnWidth = 20 'Susut Berat (g)
        .Columns("I").ColumnWidth = 20 'Baki Berat (g)
        .Columns("J").ColumnWidth = 20 'Dulang
        .Columns("K").ColumnWidth = 20 'Cawangan
        
        If Frm101.L46_Text = "Semua cawangan" Then
        
            Frm85_LM_SEARCH_5 = Null
            Frm85_LM_SEARCH_5_LOGIC = "<>"
            Frm85_LM_SEARCH_6 = Null
            Frm85_LM_SEARCH_6_LOGIC = "<>"
            
        Else
        
            Frm85_LM_SEARCH_5 = Frm101.L46_Text
            Frm85_LM_SEARCH_5_LOGIC = "="
            Frm85_LM_SEARCH_6 = "HQ"
            Frm85_LM_SEARCH_6_LOGIC = "="
            
        End If

        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        '#### Header Report ###
        .Cells(7, 1) = Frm85.L27_Text 'Report Header"
       
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Potong"
        .Cells(8, 3) = "No. Siri Produk"
        .Cells(8, 4) = "Purity"
        .Cells(8, 5) = "Kategori Produk"
        .Cells(8, 6) = "Supplier"
        .Cells(8, 7) = "Berat Asal (g)"
        .Cells(8, 8) = "Susut Berat (g)"
        .Cells(8, 9) = "Baki Berat (g)"
        .Cells(8, 10) = "Dulang"
        .Cells(8, 11) = "Cawangan"
        
        For i = 1 To 11
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND (StatusItem='" & 12 & "' OR StatusItem='" & 20 & "' OR StatusItem='" & 28 & "' OR StatusItem='" & 22 & "') order by tarikh_Jualan1 ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh_jualan1) Then .Cells(8 + x, 2) = "'" & rs!tarikh_jualan1 'Tarikh Belian
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 3) = rs!no_siri_Produk 'No. Siri Produk
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!purity) Then .Cells(8 + x, 4) = rs!purity 'Purity
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 5) = rs!kategori_Produk 'Kategori Produk
            
            If Not IsNull(rs!nama_Supplier) Then .Cells(8 + x, 6) = rs!nama_Supplier 'Nama Supplier
            
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!Berat) Then
                .Cells(8 + x, 7) = Format(rs!Berat, "#,##0.00") 'Berat (g)
            Else
                .Cells(8 + x, 7) = "0.00" 'Berat (g)
            End If
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!susut_berat) Then
                .Cells(8 + x, 8) = Format(rs!susut_berat, "#,##0.00") 'Berat (g)
            Else
                .Cells(8 + x, 8) = "0.00" 'Berat (g)
            End If
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!beza_berat) Then
                .Cells(8 + x, 9) = Format(rs!beza_berat, "#,##0.00") 'Baki Berat (g)
            Else
                .Cells(8 + x, 9) = "0.00" 'Baki Berat (g)
            End If
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            
            If Not IsNull(rs!dulang) Then
                .Cells(8 + x, 10) = rs!dulang 'Dulang
                .Cells(8 + x, 10).HorizontalAlignment = xlCenter
            End If
            If Not IsNull(rs!cawangan) Then
                .Cells(8 + x, 11) = rs!cawangan 'Dulang
                '.Cells(8 + x, 11).HorizontalAlignment = xlCenter
            End If

            For Col = 1 To 11
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 2
        .Cells(8 + Y, 1) = "Bilangan : " & x 'Total Barang
        Y = Y + 1
        
        '#### Jumlah Berat Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(beza_berat) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND (StatusItem='" & 12 & "' OR StatusItem='" & 20 & "' OR StatusItem='" & 28 & "' OR StatusItem='" & 22 & "') order by tarikh_Jualan1 ASC", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        '#### Jumlah Berat Keseluruhan #### - End
        
        Y = Y + 4
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
End Sub
Sub Frm85_belian_gb_overall()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    
    x = 0
    
    Frm85_LM_BERAT = 0
    Frm85_LM_HARGA = 0
    
    If Frm101.L9_Text = 1 Then
        TM = Frm101.L7_Text 'Tarikh Mula
        TA = Frm101.L8_Text 'Tarikh Akhir
    End If
    If Frm101.L5_Text = "Semua Purity" Then
        Frm85_LM_SEARCH_1 = Null
        Frm85_LM_SEARCH_1_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_1 = Frm101.L5_Text
        Frm85_LM_SEARCH_1_LOGIC = "="
    End If
    If Frm101.L6_Text = "Semua Kategori Produk" Then
        Frm85_LM_SEARCH_2 = Null
        Frm85_LM_SEARCH_2_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_2 = Frm101.L6_Text
        Frm85_LM_SEARCH_2_LOGIC = "="
    End If
    If Frm101.L34_Text = "Semua Dulang" Then
        Frm85_LM_SEARCH_3 = Null
        Frm85_LM_SEARCH_3_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_3 = Frm101.L34_Text
        Frm85_LM_SEARCH_3_LOGIC = "="
    End If
    If Frm101.L37_Text = "Semua Supplier" Then
        Frm85_LM_SEARCH_4 = Null
        Frm85_LM_SEARCH_4_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_4 = Frm101.L37_Text
        Frm85_LM_SEARCH_4_LOGIC = "="
    End If
    
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
    
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh Belian
        .Columns("C").ColumnWidth = 20 'No. Siri Produk
        .Columns("D").ColumnWidth = 20 'Purity
        .Columns("E").ColumnWidth = 40 'Kategori Produk
        .Columns("F").ColumnWidth = 40 'Supplier
        .Columns("G").ColumnWidth = 20 'Berat (g)
        .Columns("H").ColumnWidth = 20 'Rate Penerimaan (RM/g)
        .Columns("I").ColumnWidth = 20 'Upah (RM)
        .Columns("J").ColumnWidth = 20 'Spread (%)
        .Columns("K").ColumnWidth = 20 'Harga Selepas Spread (RM)
        .Columns("L").ColumnWidth = 20 'Adjustment (RM)
        .Columns("M").ColumnWidth = 20 'Harga Belian (RM)
        .Columns("N").ColumnWidth = 20 'Dulang
        .Columns("O").ColumnWidth = 20 'Panjang
        .Columns("P").ColumnWidth = 20 'Lebar
        .Columns("Q").ColumnWidth = 20 'Saiz
    
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        .Cells(7, 1) = Frm85.L71_Text 'Header Report
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Belian"
        .Cells(8, 3) = "No. Siri Produk"
        .Cells(8, 4) = "Purity"
        .Cells(8, 5) = "Kategori Produk"
        .Cells(8, 6) = "Supplier"
        .Cells(8, 7) = "Berat (g)"
        .Cells(8, 8) = "Rate Penerimaan (RM/g)"
        .Cells(8, 9) = "Upah (RM)"
        .Cells(8, 10) = "Spread (%)"
        .Cells(8, 11) = "Harga Selepas Spread (RM)"
        .Cells(8, 12) = "Adjustment (RM)"
        .Cells(8, 13) = "Harga Belian Termasuk GST (RM)"
        .Cells(8, 14) = "Dulang"
        .Cells(8, 15) = "Panjang"
        .Cells(8, 16) = "Lebar"
        .Cells(8, 17) = "Saiz"
        
        For i = 1 To 17
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select * from Data_Database where nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND receiving_Status='" & 4 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select * from Data_Database where nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND receiving_Status='" & 4 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh_belian) Then .Cells(8 + x, 2) = "'" & rs!tarikh_belian 'Tarikh Belian
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 3) = rs!no_siri_Produk 'No. Siri Produk
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kod_Purity) Then .Cells(8 + x, 4) = rs!kod_Purity 'Purity
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 5) = rs!kategori_Produk 'Kategori Produk
            
            If Not IsNull(rs!nama_Supplier) Then .Cells(8 + x, 6) = rs!nama_Supplier 'Nama Supplier
            
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!Berat) Then
                .Cells(8 + x, 7) = Format(rs!Berat, "#,##0.00") 'Berat (g)
            Else
                .Cells(8 + x, 7) = "0.00" 'Berat (g)
            End If
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!kos_Belian_Gram) Then
                .Cells(8 + x, 8) = Format(rs!kos_Belian_Gram, "#,##0.00") 'Rate Penerimaan (RM/g)
            Else
                .Cells(8 + x, 8) = "0.00" 'Rate Penerimaan (RM/g)
            End If
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!UPAH) Then
                .Cells(8 + x, 9) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
            Else
                .Cells(8 + x, 9) = "0.00" 'Upah (RM)
            End If
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 10).HorizontalAlignment = xlRight
            If Not IsNull(rs!SpreadValue) Then
                .Cells(8 + x, 10) = rs!SpreadValue 'Spread (%)
            Else
                .Cells(8 + x, 10) = "0.00" 'Spread (%)
            End If
            .Cells(8 + x, 10).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 11).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_lepas_spread) Then
                .Cells(8 + x, 11) = Format(rs!harga_lepas_spread, "#,##0.00") 'Harga Selepas Spread (RM)
            Else
                .Cells(8 + x, 11) = "0.00" 'Harga Selepas Spread (RM)
            End If
            .Cells(8 + x, 11).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 12).HorizontalAlignment = xlRight
            If Not IsNull(rs!adjustment) Then
                .Cells(8 + x, 12) = Format(rs!adjustment, "#,##0.00") 'Adjustment (RM)
            Else
                .Cells(8 + x, 12) = "0.00" 'Adjustment (RM)
            End If
            .Cells(8 + x, 12).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 13).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_item) Then
                .Cells(8 + x, 13) = Format(rs!harga_item, "#,##0.00") 'Harga Belian (RM) : Tidak Campur Cukai GST
            Else
                .Cells(8 + x, 13) = "0.00" 'Harga Belian (RM) : Tidak Campur Cukai GST
            End If
            .Cells(8 + x, 13).NumberFormat = "#,##0.00"
            
            If Not IsNull(rs!dulang) Then
                .Cells(8 + x, 14) = rs!dulang 'Dulang
                .Cells(8 + x, 14).HorizontalAlignment = xlCenter
            End If
                
            If Not IsNull(rs!dimension_Panjang) Then
                .Cells(8 + x, 15) = rs!dimension_Panjang 'Panjang
                .Cells(8 + x, 15).HorizontalAlignment = xlCenter
            End If
    
            If Not IsNull(rs!dimension_Lebar) Then
                .Cells(8 + x, 16) = rs!dimension_Lebar 'Lebar
                .Cells(8 + x, 16).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!dimension_Saiz) Then
                .Cells(8 + x, 17) = rs!dimension_Saiz 'Tebal
                .Cells(8 + x, 17).HorizontalAlignment = xlCenter
            End If
            
            For Col = 1 To 17
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 2
        .Cells(8 + Y, 1) = "Bilangan : " & x 'Total Barang
        Y = Y + 1
        
        '#### Jumlah Berat Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select SUM(Berat) from Data_Database where nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND receiving_Status='" & 4 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select SUM(Berat) from Data_Database where nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND receiving_Status='" & 4 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        '#### Jumlah Berat Keseluruhan #### - End
        
        Y = Y + 1
        
        '#### Jumlah Modal Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select SUM(harga_item) from Data_Database where nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND receiving_Status='" & 4 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select SUM(harga_item) from Data_Database where nama_Supplier " & Frm85_LM_SEARCH_4_LOGIC & "'" & Frm85_LM_SEARCH_4 & "' AND kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND receiving_Status='" & 4 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "0.00"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "RM 0.00"
        End If
        '#### Jumlah Modal Keseluruhan #### - End
        
        Y = Y + 4
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
End Sub
Sub Frm85_belian_gb_siri()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    
    x = 0
    
    Frm85_LM_BERAT = 0
    Frm85_LM_HARGA = 0
    
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
    
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh Belian
        .Columns("C").ColumnWidth = 20 'No. Siri Produk
        .Columns("D").ColumnWidth = 20 'Purity
        .Columns("E").ColumnWidth = 40 'Kategori Produk
        .Columns("F").ColumnWidth = 40 'Supplier
        .Columns("G").ColumnWidth = 20 'Berat (g)
        .Columns("H").ColumnWidth = 20 'Rate Penerimaan (RM/g)
        .Columns("I").ColumnWidth = 20 'Upah (RM)
        .Columns("J").ColumnWidth = 20 'Spread (%)
        .Columns("K").ColumnWidth = 20 'Harga Selepas Spread (RM)
        .Columns("L").ColumnWidth = 20 'Adjustment (RM)
        .Columns("M").ColumnWidth = 20 'Harga Belian (RM)
        .Columns("N").ColumnWidth = 20 'Dulang
        .Columns("O").ColumnWidth = 20 'Panjang
        .Columns("P").ColumnWidth = 20 'Lebar
        .Columns("Q").ColumnWidth = 20 'Saiz
    
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        .Cells(7, 1) = Frm85.L71_Text 'Header Report
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Belian"
        .Cells(8, 3) = "No. Siri Produk"
        .Cells(8, 4) = "Purity"
        .Cells(8, 5) = "Kategori Produk"
        .Cells(8, 6) = "Supplier"
        .Cells(8, 7) = "Berat (g)"
        .Cells(8, 8) = "Rate Penerimaan (RM/g)"
        .Cells(8, 9) = "Upah (RM)"
        .Cells(8, 10) = "Spread (%)"
        .Cells(8, 11) = "Harga Selepas Spread (RM)"
        .Cells(8, 12) = "Adjustment (RM)"
        .Cells(8, 13) = "Harga Belian Termasuk GST (RM)"
        .Cells(8, 14) = "Dulang"
        .Cells(8, 15) = "Panjang"
        .Cells(8, 16) = "Lebar"
        .Cells(8, 17) = "Saiz"
        
        For i = 1 To 17
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND receiving_Status='" & 4 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh_belian) Then .Cells(8 + x, 2) = "'" & rs!tarikh_belian 'Tarikh Belian
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 3) = rs!no_siri_Produk 'No. Siri Produk
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kod_Purity) Then .Cells(8 + x, 4) = rs!kod_Purity 'Purity
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 5) = rs!kategori_Produk 'Kategori Produk
            
            If Not IsNull(rs!nama_Supplier) Then .Cells(8 + x, 6) = rs!nama_Supplier 'Nama Supplier
            
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!Berat) Then
                .Cells(8 + x, 7) = Format(rs!Berat, "#,##0.00") 'Berat (g)
            Else
                .Cells(8 + x, 7) = "0.00" 'Berat (g)
            End If
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!kos_Belian_Gram) Then
                .Cells(8 + x, 8) = Format(rs!kos_Belian_Gram, "#,##0.00") 'Rate Penerimaan (RM/g)
            Else
                .Cells(8 + x, 8) = "0.00" 'Rate Penerimaan (RM/g)
            End If
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!UPAH) Then
                .Cells(8 + x, 9) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
            Else
                .Cells(8 + x, 9) = "0.00" 'Upah (RM)
            End If
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 10).HorizontalAlignment = xlRight
            If Not IsNull(rs!SpreadValue) Then
                .Cells(8 + x, 10) = rs!SpreadValue 'Spread (%)
            Else
                .Cells(8 + x, 10) = "0.00" 'Spread (%)
            End If
            .Cells(8 + x, 10).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 11).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_lepas_spread) Then
                .Cells(8 + x, 11) = Format(rs!harga_lepas_spread, "#,##0.00") 'Harga Selepas Spread (RM)
            Else
                .Cells(8 + x, 11) = "0.00" 'Harga Selepas Spread (RM)
            End If
            .Cells(8 + x, 11).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 12).HorizontalAlignment = xlRight
            If Not IsNull(rs!adjustment) Then
                .Cells(8 + x, 12) = Format(rs!adjustment, "#,##0.00") 'Adjustment (RM)
            Else
                .Cells(8 + x, 12) = "0.00" 'Adjustment (RM)
            End If
            .Cells(8 + x, 12).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 13).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_item) Then
                .Cells(8 + x, 13) = Format(rs!harga_item, "#,##0.00") 'Harga Belian (RM) : Tidak Campur Cukai GST
            Else
                .Cells(8 + x, 13) = "0.00" 'Harga Belian (RM) : Tidak Campur Cukai GST
            End If
            .Cells(8 + x, 13).NumberFormat = "#,##0.00"
            
            If Not IsNull(rs!dulang) Then
                .Cells(8 + x, 14) = rs!dulang 'Dulang
                .Cells(8 + x, 14).HorizontalAlignment = xlCenter
            End If
                
            If Not IsNull(rs!dimension_Panjang) Then
                .Cells(8 + x, 15) = rs!dimension_Panjang 'Panjang
                .Cells(8 + x, 15).HorizontalAlignment = xlCenter
            End If
    
            If Not IsNull(rs!dimension_Lebar) Then
                .Cells(8 + x, 16) = rs!dimension_Lebar 'Lebar
                .Cells(8 + x, 16).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!dimension_Saiz) Then
                .Cells(8 + x, 17) = rs!dimension_Saiz 'Tebal
                .Cells(8 + x, 17).HorizontalAlignment = xlCenter
            End If
            
            For Col = 1 To 17
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 2
        .Cells(8 + Y, 1) = "Bilangan : " & x 'Total Barang
        Y = Y + 1
        
        '#### Jumlah Berat Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(Berat) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND receiving_Status='" & 4 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        '#### Jumlah Berat Keseluruhan #### - End
        
        Y = Y + 1
        
        '#### Jumlah Modal Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(harga_item) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND receiving_Status='" & 4 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "0.00"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "RM 0.00"
        End If
        '#### Jumlah Modal Keseluruhan #### - End
        
        Y = Y + 4
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
End Sub
Sub Frm85_buyback_gb_overall()
'On Error Resume Next
Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook

If Frm101.L9_Text = 1 Then
    TM = Frm101.L7_Text 'Tarikh Mula
    TA = Frm101.L8_Text 'Tarikh Akhir
End If
If Frm101.L5_Text = "Semua Purity" Then
    Frm85_LM_SEARCH_1 = Null
    Frm85_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_1 = Frm101.L5_Text
    Frm85_LM_SEARCH_1_LOGIC = "="
End If
If Frm101.L6_Text = "Semua Kategori Produk" Then
    Frm85_LM_SEARCH_2 = Null
    Frm85_LM_SEARCH_2_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_2 = Frm101.L6_Text
    Frm85_LM_SEARCH_2_LOGIC = "="
End If
If Frm101.L34_Text = "Semua Dulang" Then
    Frm85_LM_SEARCH_3 = Null
    Frm85_LM_SEARCH_3_LOGIC = "<>"
Else
    Frm85_LM_SEARCH_3 = Frm101.L34_Text
    Frm85_LM_SEARCH_3_LOGIC = "="
End If
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh Belian
        .Columns("C").ColumnWidth = 20 'No. Siri Produk
        .Columns("D").ColumnWidth = 20 'Purity
        .Columns("E").ColumnWidth = 40 'Kategori Produk
        .Columns("F").ColumnWidth = 40 'Supplier
        .Columns("G").ColumnWidth = 20 'Berat (g)
        .Columns("H").ColumnWidth = 20 'Rate Penerimaan (RM/g)
        .Columns("I").ColumnWidth = 20 'Upah (RM)
        .Columns("J").ColumnWidth = 20 'Spread (%)
        .Columns("K").ColumnWidth = 20 'Harga Selepas Spread (RM)
        .Columns("L").ColumnWidth = 20 'Adjustment (RM)
        .Columns("M").ColumnWidth = 20 'Harga Belian (RM)
        .Columns("N").ColumnWidth = 20 'Dulang
        .Columns("O").ColumnWidth = 20 'Panjang
        .Columns("P").ColumnWidth = 20 'Lebar
        .Columns("Q").ColumnWidth = 20 'Saiz
    
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        .Cells(7, 1) = Frm85.L72_Text 'Header Report
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Belian"
        .Cells(8, 3) = "No. Siri Produk"
        .Cells(8, 4) = "Purity"
        .Cells(8, 5) = "Kategori Produk"
        .Cells(8, 6) = "Supplier"
        .Cells(8, 7) = "Berat (g)"
        .Cells(8, 8) = "Rate Penerimaan (RM/g)"
        .Cells(8, 9) = "Upah (RM)"
        .Cells(8, 10) = "Spread (%)"
        .Cells(8, 11) = "Harga Selepas Spread (RM)"
        .Cells(8, 12) = "Adjustment (RM)"
        .Cells(8, 13) = "Harga Belian Termasuk GST (RM)"
        .Cells(8, 14) = "Dulang"
        .Cells(8, 15) = "Panjang"
        .Cells(8, 16) = "Lebar"
        .Cells(8, 17) = "Saiz"
        
        For i = 1 To 17
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select * from Data_Database where kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select * from Data_Database where kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh_belian) Then .Cells(8 + x, 2) = "'" & rs!tarikh_belian 'Tarikh Belian
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 3) = rs!no_siri_Produk 'No. Siri Produk
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kod_Purity) Then .Cells(8 + x, 4) = rs!kod_Purity 'Purity
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 5) = rs!kategori_Produk 'Kategori Produk
            
            If Not IsNull(rs!nama_Supplier) Then .Cells(8 + x, 6) = rs!nama_Supplier 'Nama Supplier
            
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!Berat) Then
                .Cells(8 + x, 7) = Format(rs!Berat, "#,##0.00") 'Berat (g)
            Else
                .Cells(8 + x, 7) = "0.00" 'Berat (g)
            End If
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!kos_Belian_Gram) Then
                .Cells(8 + x, 8) = Format(rs!kos_Belian_Gram, "#,##0.00") 'Rate Penerimaan (RM/g)
            Else
                .Cells(8 + x, 8) = "0.00" 'Rate Penerimaan (RM/g)
            End If
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!UPAH) Then
                .Cells(8 + x, 9) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
            Else
                .Cells(8 + x, 9) = "0.00" 'Upah (RM)
            End If
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 10).HorizontalAlignment = xlRight
            If Not IsNull(rs!SpreadValue) Then
                .Cells(8 + x, 10) = rs!SpreadValue 'Spread (%)
            Else
                .Cells(8 + x, 10) = "0.00" 'Spread (%)
            End If
            .Cells(8 + x, 10).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 11).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_lepas_spread) Then
                .Cells(8 + x, 11) = Format(rs!harga_lepas_spread, "#,##0.00") 'Harga Selepas Spread (RM)
            Else
                .Cells(8 + x, 11) = "0.00" 'Harga Selepas Spread (RM)
            End If
            .Cells(8 + x, 11).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 12).HorizontalAlignment = xlRight
            If Not IsNull(rs!adjustment) Then
                .Cells(8 + x, 12) = Format(rs!adjustment, "#,##0.00") 'Adjustment (RM)
            Else
                .Cells(8 + x, 12) = "0.00" 'Adjustment (RM)
            End If
            .Cells(8 + x, 12).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 13).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_item) Then
                .Cells(8 + x, 13) = Format(rs!harga_item, "#,##0.00") 'Harga Belian (RM) : Tidak Campur Cukai GST
            Else
                .Cells(8 + x, 13) = "0.00" 'Harga Belian (RM) : Tidak Campur Cukai GST
            End If
            .Cells(8 + x, 13).NumberFormat = "#,##0.00"
            
            If Not IsNull(rs!dulang) Then
                .Cells(8 + x, 14) = rs!dulang 'Dulang
                .Cells(8 + x, 14).HorizontalAlignment = xlCenter
            End If
                
            If Not IsNull(rs!dimension_Panjang) Then
                .Cells(8 + x, 15) = rs!dimension_Panjang 'Panjang
                .Cells(8 + x, 15).HorizontalAlignment = xlCenter
            End If
    
            If Not IsNull(rs!dimension_Lebar) Then
                .Cells(8 + x, 16) = rs!dimension_Lebar 'Lebar
                .Cells(8 + x, 16).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!dimension_Saiz) Then
                .Cells(8 + x, 17) = rs!dimension_Saiz 'Tebal
                .Cells(8 + x, 17).HorizontalAlignment = xlCenter
            End If
            
            For Col = 1 To 17
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 2

        '#### Jumlah Bilangan Barang Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select COUNT(no_siri_Produk) from Data_Database where kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select COUNT(no_siri_Produk) from Data_Database where kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Bilangan : " & rs(0)
        Else
            .Cells(8 + Y, 1) = "Bilangan : " & 0
        End If
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Bilangan : " & 0
        End If
        '#### Jumlah Bilangan Barang Keseluruhan #### - End
        
        Y = Y + 1
        
        '#### Jumlah Berat Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select SUM(Berat) from Data_Database where kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select SUM(Berat) from Data_Database where kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        '#### Jumlah Berat Keseluruhan #### - End
        
        Y = Y + 1
        
        '#### Jumlah Modal Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select SUM(harga_item) from Data_Database where kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select SUM(harga_item) from Data_Database where kod_Purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_Produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "0.00"
        End If
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "0.00"
        End If
        '#### Jumlah Modal Keseluruhan #### - End
        
        Y = Y + 4
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
End Sub
Sub Frm85_buyback_gb_siri()
'On Error Resume Next
Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh Belian
        .Columns("C").ColumnWidth = 20 'No. Siri Produk
        .Columns("D").ColumnWidth = 20 'Purity
        .Columns("E").ColumnWidth = 40 'Kategori Produk
        .Columns("F").ColumnWidth = 40 'Supplier
        .Columns("G").ColumnWidth = 20 'Berat (g)
        .Columns("H").ColumnWidth = 20 'Rate Penerimaan (RM/g)
        .Columns("I").ColumnWidth = 20 'Upah (RM)
        .Columns("J").ColumnWidth = 20 'Spread (%)
        .Columns("K").ColumnWidth = 20 'Harga Selepas Spread (RM)
        .Columns("L").ColumnWidth = 20 'Adjustment (RM)
        .Columns("M").ColumnWidth = 20 'Harga Belian (RM)
        .Columns("N").ColumnWidth = 20 'Dulang
        .Columns("O").ColumnWidth = 20 'Panjang
        .Columns("P").ColumnWidth = 20 'Lebar
        .Columns("Q").ColumnWidth = 20 'Saiz
    
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        .Cells(7, 1) = Frm85.L72_Text 'Header Report
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh Belian"
        .Cells(8, 3) = "No. Siri Produk"
        .Cells(8, 4) = "Purity"
        .Cells(8, 5) = "Kategori Produk"
        .Cells(8, 6) = "Supplier"
        .Cells(8, 7) = "Berat (g)"
        .Cells(8, 8) = "Rate Penerimaan (RM/g)"
        .Cells(8, 9) = "Upah (RM)"
        .Cells(8, 10) = "Spread (%)"
        .Cells(8, 11) = "Harga Selepas Spread (RM)"
        .Cells(8, 12) = "Adjustment (RM)"
        .Cells(8, 13) = "Harga Belian Termasuk GST (RM)"
        .Cells(8, 14) = "Dulang"
        .Cells(8, 15) = "Panjang"
        .Cells(8, 16) = "Lebar"
        .Cells(8, 17) = "Saiz"
        
        For i = 1 To 17
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND StatusItem <> 0 AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh_belian) Then .Cells(8 + x, 2) = "'" & rs!tarikh_belian 'Tarikh Belian
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 3) = rs!no_siri_Produk 'No. Siri Produk
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kod_Purity) Then .Cells(8 + x, 4) = rs!kod_Purity 'Purity
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 5) = rs!kategori_Produk 'Kategori Produk
            
            If Not IsNull(rs!nama_Supplier) Then .Cells(8 + x, 6) = rs!nama_Supplier 'Nama Supplier
            
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!Berat) Then
                .Cells(8 + x, 7) = Format(rs!Berat, "#,##0.00") 'Berat (g)
            Else
                .Cells(8 + x, 7) = "0.00" 'Berat (g)
            End If
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!kos_Belian_Gram) Then
                .Cells(8 + x, 8) = Format(rs!kos_Belian_Gram, "#,##0.00") 'Rate Penerimaan (RM/g)
            Else
                .Cells(8 + x, 8) = "0.00" 'Rate Penerimaan (RM/g)
            End If
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!UPAH) Then
                .Cells(8 + x, 9) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
            Else
                .Cells(8 + x, 9) = "0.00" 'Upah (RM)
            End If
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 10).HorizontalAlignment = xlRight
            If Not IsNull(rs!SpreadValue) Then
                .Cells(8 + x, 10) = rs!SpreadValue 'Spread (%)
            Else
                .Cells(8 + x, 10) = "0.00" 'Spread (%)
            End If
            .Cells(8 + x, 10).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 11).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_lepas_spread) Then
                .Cells(8 + x, 11) = Format(rs!harga_lepas_spread, "#,##0.00") 'Harga Selepas Spread (RM)
            Else
                .Cells(8 + x, 11) = "0.00" 'Harga Selepas Spread (RM)
            End If
            .Cells(8 + x, 11).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 12).HorizontalAlignment = xlRight
            If Not IsNull(rs!adjustment) Then
                .Cells(8 + x, 12) = Format(rs!adjustment, "#,##0.00") 'Adjustment (RM)
            Else
                .Cells(8 + x, 12) = "0.00" 'Adjustment (RM)
            End If
            .Cells(8 + x, 12).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 13).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_item) Then
                .Cells(8 + x, 13) = Format(rs!harga_item, "#,##0.00") 'Harga Belian (RM) : Tidak Campur Cukai GST
            Else
                .Cells(8 + x, 13) = "0.00" 'Harga Belian (RM) : Tidak Campur Cukai GST
            End If
            .Cells(8 + x, 13).NumberFormat = "#,##0.00"
            
            If Not IsNull(rs!dulang) Then
                .Cells(8 + x, 14) = rs!dulang 'Dulang
                .Cells(8 + x, 14).HorizontalAlignment = xlCenter
            End If
                
            If Not IsNull(rs!dimension_Panjang) Then
                .Cells(8 + x, 15) = rs!dimension_Panjang 'Panjang
                .Cells(8 + x, 15).HorizontalAlignment = xlCenter
            End If
    
            If Not IsNull(rs!dimension_Lebar) Then
                .Cells(8 + x, 16) = rs!dimension_Lebar 'Lebar
                .Cells(8 + x, 16).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!dimension_Saiz) Then
                .Cells(8 + x, 17) = rs!dimension_Saiz 'Tebal
                .Cells(8 + x, 17).HorizontalAlignment = xlCenter
            End If
            
            For Col = 1 To 17
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        Y = x + 2

        '#### Jumlah Bilangan Barang Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select COUNT(no_siri_Produk) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND StatusItem <> 0 AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Bilangan : " & rs(0)
        Else
            .Cells(8 + Y, 1) = "Bilangan : " & 0
        End If
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Bilangan : " & 0
        End If
        '#### Jumlah Bilangan Barang Keseluruhan #### - End
        
        Y = Y + 1
        
        '#### Jumlah Berat Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(Berat) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND StatusItem <> 0 AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        '#### Jumlah Berat Keseluruhan #### - End
        
        Y = Y + 1
        
        '#### Jumlah Modal Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(harga_item) from Data_Database where no_siri_Produk='" & Frm101.L5_Text & "' AND StatusItem <> 0 AND receiving_Status='" & 5 & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "0.00"
        End If
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Modal : RM " & "0.00"
        End If
        '#### Jumlah Modal Keseluruhan #### - End
        
        Y = Y + 4
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
End Sub
Sub Frm85_ansuran_overall()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    
    x = 0
    
    Frm85_LM_BERAT = 0
    Frm85_LM_HARGA = 0
    
    If Frm101.L9_Text = 1 Then
        TM = Frm101.L7_Text 'Tarikh Mula
        TA = Frm101.L8_Text 'Tarikh Akhir
    End If
    If Frm101.L5_Text = "Semua Purity" Then
        Frm85_LM_SEARCH_1 = Null
        Frm85_LM_SEARCH_1_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_1 = Frm101.L5_Text
        Frm85_LM_SEARCH_1_LOGIC = "="
    End If
    If Frm101.L6_Text = "Semua Kategori Produk" Then
        Frm85_LM_SEARCH_2 = Null
        Frm85_LM_SEARCH_2_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_2 = Frm101.L6_Text
        Frm85_LM_SEARCH_2_LOGIC = "="
    End If
    If Frm101.L34_Text = "Semua Dulang" Then
        Frm85_LM_SEARCH_3 = Null
        Frm85_LM_SEARCH_3_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_3 = Frm101.L34_Text
        Frm85_LM_SEARCH_3_LOGIC = "="
    End If
    
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh
        .Columns("C").ColumnWidth = 20 'Jenis Ansuran
        .Columns("D").ColumnWidth = 20 'No. Siri Produk
        .Columns("E").ColumnWidth = 40 'Kategori Produk
        .Columns("F").ColumnWidth = 20 'Purity
        .Columns("G").ColumnWidth = 20 'Berat Asal (g)
        .Columns("H").ColumnWidth = 20 'Berat Jualan (g)
        .Columns("I").ColumnWidth = 20 'Upah (RM)
        .Columns("J").ColumnWidth = 20 'Jumlah Harga (RM)
        .Columns("K").ColumnWidth = 20 'Dulang
    
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        '#### Header Report ###
        .Cells(7, 1) = Frm85.L35_Text 'Report Header"
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh"
        .Cells(8, 3) = "Jenis Ansuran"
        .Cells(8, 4) = "No. Siri Produk"
        .Cells(8, 5) = "Kategori Produk"
        .Cells(8, 6) = "Purity"
        .Cells(8, 7) = "Berat Asal (g)"
        .Cells(8, 8) = "Berat Jualan (g)"
        .Cells(8, 9) = "Upah (RM)"
        .Cells(8, 10) = "Jumlah Harga (RM)"
        .Cells(8, 11) = "Dulang"
        
        For i = 1 To 11
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select * from 27_senarai_ansuran where status='" & "Jelas" & "' AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select * from 27_senarai_ansuran where status='" & "Jelas" & "' AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh_jelas BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh_jelas) Then .Cells(8 + x, 2) = "'" & rs!tarikh_jelas 'Tarikh
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!jenis_ansuran) Then 'Jenis Ansuran
                If rs!jenis_ansuran = 0 Then
                    .Cells(8 + x, 3) = "Harga Semasa"
                ElseIf rs!jenis_ansuran = 1 Then
                    .Cells(8 + x, 3) = "Harga Tetap"
                End If
            End If
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 4) = rs!no_siri_Produk 'No. Siri Produk
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 5) = rs!kategori_Produk 'Kategori Produk
            
            If Not IsNull(rs!purity) Then .Cells(8 + x, 6) = rs!purity 'Purity
            .Cells(8 + x, 6).HorizontalAlignment = xlCenter
            
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!Berat_Asal) Then
                .Cells(8 + x, 7) = Format(rs!Berat_Asal, "#,##0.00") 'Berat Asal (g)
            Else
                .Cells(8 + x, 7) = "0.00" 'Berat (g)
            End If
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!berat_jualan) Then
                .Cells(8 + x, 8) = Format(rs!berat_jualan, "#,##0.00") 'Berat Jualan (g)
            Else
                .Cells(8 + x, 8) = "0.00" 'Berat Jualan (g)
            End If
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!UPAH) Then
                .Cells(8 + x, 9) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
            Else
                .Cells(8 + x, 9) = "0.00" 'Upah (RM)
            End If
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 10).HorizontalAlignment = xlRight
            If Not IsNull(rs!jumlah_bayaran) Then
                .Cells(8 + x, 10) = Format(rs!jumlah_bayaran, "#,##0.00") 'Harga Asal Jualan (RM)
            Else
                .Cells(8 + x, 10) = "0.00" 'Harga Asal Jualan (RM)
            End If
            .Cells(8 + x, 10).NumberFormat = "#,##0.00"
            
            If Not IsNull(rs!dulang) Then
                .Cells(8 + x, 11) = rs!dulang 'Dulang
                .Cells(8 + x, 11).HorizontalAlignment = xlCenter
            End If
            
            For Col = 1 To 11
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 2
        
        '#### Jumlah Bilangan Barang Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select COUNT(no_siri_Produk) from 27_senarai_ansuran where status='" & "Jelas" & "' AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select COUNT(no_siri_Produk) from 27_senarai_ansuran where status='" & "Jelas" & "' AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh_jelas BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Bilangan : " & rs(0)
        Else
            .Cells(8 + Y, 1) = "Bilangan : " & 0
        End If
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Bilangan : " & 0
        End If
        '#### Jumlah Bilangan Barang Keseluruhan #### - End
        
        '#### Jumlah Berat Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select SUM(Berat_Jualan) from 27_senarai_ansuran where status='" & "Jelas" & "' AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select SUM(Berat_Jualan) from 27_senarai_ansuran where status='" & "Jelas" & "' AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh_jelas BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        '#### Jumlah Berat Keseluruhan #### - End
        
        Y = Y + 1
        
        '#### Jumlah Modal Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select SUM(jumlah_bayaran) from 27_senarai_ansuran where status='" & "Jelas" & "' AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select SUM(jumlah_bayaran) from 27_senarai_ansuran where status='" & "Jelas" & "' AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh_jelas BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Harga : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Harga : RM " & "0.00"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Harga : RM " & "RM 0.00"
        End If
        '#### Jumlah Modal Keseluruhan #### - End
        
        Y = Y + 4
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
End Sub
Sub Frm85_ansuran_siri()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    
    x = 0
    
    Frm85_LM_BERAT = 0
    Frm85_LM_HARGA = 0
    
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh
        .Columns("C").ColumnWidth = 20 'Jenis Ansuran
        .Columns("D").ColumnWidth = 20 'No. Siri Produk
        .Columns("E").ColumnWidth = 40 'Kategori Produk
        .Columns("F").ColumnWidth = 20 'Purity
        .Columns("G").ColumnWidth = 20 'Berat Asal (g)
        .Columns("H").ColumnWidth = 20 'Berat Jualan (g)
        .Columns("I").ColumnWidth = 20 'Upah (RM)
        .Columns("J").ColumnWidth = 20 'Jumlah Harga (RM)
        .Columns("K").ColumnWidth = 20 'Dulang
    
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        '#### Header Report ###
        .Cells(7, 1) = Frm85.L35_Text 'Report Header"
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh"
        .Cells(8, 3) = "Jenis Ansuran"
        .Cells(8, 4) = "No. Siri Produk"
        .Cells(8, 5) = "Kategori Produk"
        .Cells(8, 6) = "Purity"
        .Cells(8, 7) = "Berat Asal (g)"
        .Cells(8, 8) = "Berat Jualan (g)"
        .Cells(8, 9) = "Upah (RM)"
        .Cells(8, 10) = "Jumlah Harga (RM)"
        .Cells(8, 11) = "Dulang"
        
        For i = 1 To 11
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 27_senarai_ansuran where status='" & "Jelas" & "' AND no_siri_Produk='" & Frm101.L5_Text & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh_jelas) Then .Cells(8 + x, 2) = "'" & rs!tarikh_jelas 'Tarikh
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!jenis_ansuran) Then 'Jenis Ansuran
                If rs!jenis_ansuran = 0 Then
                    .Cells(8 + x, 3) = "Harga Semasa"
                ElseIf rs!jenis_ansuran = 1 Then
                    .Cells(8 + x, 3) = "Harga Tetap"
                End If
            End If
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 4) = rs!no_siri_Produk 'No. Siri Produk
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 5) = rs!kategori_Produk 'Kategori Produk
            
            If Not IsNull(rs!purity) Then .Cells(8 + x, 6) = rs!purity 'Purity
            .Cells(8 + x, 6).HorizontalAlignment = xlCenter
            
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!Berat_Asal) Then
                .Cells(8 + x, 7) = Format(rs!Berat_Asal, "#,##0.00") 'Berat Asal (g)
            Else
                .Cells(8 + x, 7) = "0.00" 'Berat (g)
            End If
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!berat_jualan) Then
                .Cells(8 + x, 8) = Format(rs!berat_jualan, "#,##0.00") 'Berat Jualan (g)
            Else
                .Cells(8 + x, 8) = "0.00" 'Berat Jualan (g)
            End If
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!UPAH) Then
                .Cells(8 + x, 9) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
            Else
                .Cells(8 + x, 9) = "0.00" 'Upah (RM)
            End If
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 10).HorizontalAlignment = xlRight
            If Not IsNull(rs!jumlah_bayaran) Then
                .Cells(8 + x, 10) = Format(rs!jumlah_bayaran, "#,##0.00") 'Harga Asal Jualan (RM)
            Else
                .Cells(8 + x, 10) = "0.00" 'Harga Asal Jualan (RM)
            End If
            .Cells(8 + x, 10).NumberFormat = "#,##0.00"
            
            If Not IsNull(rs!dulang) Then
                .Cells(8 + x, 11) = rs!dulang 'Dulang
                .Cells(8 + x, 11).HorizontalAlignment = xlCenter
            End If
            
            For Col = 1 To 11
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 2
        
        '#### Jumlah Bilangan Barang Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select COUNT(no_siri_Produk) from 27_senarai_ansuran where status='" & "Jelas" & "' AND no_siri_Produk='" & Frm101.L5_Text & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Bilangan : " & rs(0)
        Else
            .Cells(8 + Y, 1) = "Bilangan : " & 0
        End If
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Bilangan : " & 0
        End If
        '#### Jumlah Bilangan Barang Keseluruhan #### - End
        
        '#### Jumlah Berat Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(Berat_Jualan) from 27_senarai_ansuran where status='" & "Jelas" & "' AND no_siri_Produk='" & Frm101.L5_Text & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        '#### Jumlah Berat Keseluruhan #### - End
        
        Y = Y + 1
        
        '#### Jumlah Modal Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(jumlah_bayaran) from 27_senarai_ansuran where status='" & "Jelas" & "' AND no_siri_Produk='" & Frm101.L5_Text & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Harga : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Harga : RM " & "0.00"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Harga : RM " & "RM 0.00"
        End If
        '#### Jumlah Modal Keseluruhan #### - End
        
        Y = Y + 4
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
End Sub
Sub Frm85_tempahan_overall()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    
    x = 0
    
    Frm85_LM_BERAT = 0
    Frm85_LM_HARGA = 0
    
    If Frm101.L9_Text = 1 Then
        TM = Frm101.L7_Text 'Tarikh Mula
        TA = Frm101.L8_Text 'Tarikh Akhir
    End If
    If Frm101.L5_Text = "Semua Purity" Then
        Frm85_LM_SEARCH_1 = Null
        Frm85_LM_SEARCH_1_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_1 = Frm101.L5_Text
        Frm85_LM_SEARCH_1_LOGIC = "="
    End If
    If Frm101.L6_Text = "Semua Kategori Produk" Then
        Frm85_LM_SEARCH_2 = Null
        Frm85_LM_SEARCH_2_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_2 = Frm101.L6_Text
        Frm85_LM_SEARCH_2_LOGIC = "="
    End If
    If Frm101.L34_Text = "Semua Dulang" Then
        Frm85_LM_SEARCH_3 = Null
        Frm85_LM_SEARCH_3_LOGIC = "<>"
    Else
        Frm85_LM_SEARCH_3 = Frm101.L34_Text
        Frm85_LM_SEARCH_3_LOGIC = "="
    End If
    If Frm101.L46_Text = "Semua cawangan" Then
    
        Frm85_LM_SEARCH_5 = Null
        Frm85_LM_SEARCH_5_LOGIC = "<>"
        Frm85_LM_SEARCH_6 = Null
        Frm85_LM_SEARCH_6_LOGIC = "<>"
        
    Else
    
        Frm85_LM_SEARCH_5 = Frm101.L46_Text
        Frm85_LM_SEARCH_5_LOGIC = "="
        Frm85_LM_SEARCH_6 = "HQ"
        Frm85_LM_SEARCH_6_LOGIC = "="
        
    End If

    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh
        .Columns("C").ColumnWidth = 20 'No. Siri Produk
        .Columns("D").ColumnWidth = 40 'Kategori Produk
        .Columns("E").ColumnWidth = 20 'Purity
        .Columns("F").ColumnWidth = 20 'Berat Asal (g)
        .Columns("G").ColumnWidth = 20 'Berat Jualan (g)
        .Columns("H").ColumnWidth = 20 'Upah (RM)
        .Columns("I").ColumnWidth = 20 'Jumlah Harga (RM)
        .Columns("J").ColumnWidth = 20 'Dulang
        .Columns("K").ColumnWidth = 50 'Cawangan
        
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        '#### Header Report ###
        .Cells(7, 1) = Frm85.L39_Text 'Report Header"
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh"
        .Cells(8, 3) = "No. Siri Produk"
        .Cells(8, 4) = "Kategori Produk"
        .Cells(8, 5) = "Purity"
        .Cells(8, 6) = "Berat Asal (g)"
        .Cells(8, 7) = "Berat Jualan (g)"
        .Cells(8, 8) = "Upah (RM)"
        .Cells(8, 9) = "Jumlah Harga (RM)"
        .Cells(8, 10) = "Dulang"
        .Cells(8, 11) = "Cawangan"
        
        For i = 1 To 11
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select * from 42_tempahan_siap where status_invoice = 1 AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select * from 42_tempahan_siap where status_invoice = 1 AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh) Then .Cells(8 + x, 2) = "'" & rs!tarikh 'Tarikh
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 3) = rs!no_siri_Produk 'No. Siri Produk
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 4) = rs!kategori_Produk 'Kategori Produk
            
            If Not IsNull(rs!purity) Then .Cells(8 + x, 5) = rs!purity 'Purity
            .Cells(8 + x, 5).HorizontalAlignment = xlCenter
            
            .Cells(8 + x, 6).HorizontalAlignment = xlRight
            If Not IsNull(rs!Berat_Asal) Then
                .Cells(8 + x, 6) = Format(rs!Berat_Asal, "#,##0.00") 'Berat Asal (g)
            Else
                .Cells(8 + x, 6) = "0.00" 'Berat (g)
            End If
            .Cells(8 + x, 6).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!berat_jualan) Then
                .Cells(8 + x, 7) = Format(rs!berat_jualan, "#,##0.00") 'Berat Jualan (g)
            Else
                .Cells(8 + x, 7) = "0.00" 'Berat Jualan (g)
            End If
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!UPAH) Then
                .Cells(8 + x, 8) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
            Else
                .Cells(8 + x, 8) = "0.00" 'Upah (RM)
            End If
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_dengan_gst) Then
                .Cells(8 + x, 9) = Format(rs!harga_dengan_gst, "#,##0.00") 'Harga Jualan (RM)
            Else
                .Cells(8 + x, 9) = "0.00" 'Harga Jualan (RM)
            End If
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            
            If Not IsNull(rs!dulang) Then
                .Cells(8 + x, 10) = rs!dulang 'Dulang
                .Cells(8 + x, 10).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!cawangan) Then
                .Cells(8 + x, 11) = rs!cawangan 'cawangan
                '.Cells(8 + x, 11).HorizontalAlignment = xlCenter
            End If
            
            For Col = 1 To 11
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 2
        
        '#### Jumlah Bilangan Barang Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select COUNT(no_siri_Produk) from 42_tempahan_siap where status_invoice = 1 AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select COUNT(no_siri_Produk) from 42_tempahan_siap where status_invoice = 1 AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Bilangan : " & rs(0)
        Else
            .Cells(8 + Y, 1) = "Bilangan : " & 0
        End If
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Bilangan : " & 0
        End If
        '#### Jumlah Bilangan Barang Keseluruhan #### - End
        
        '#### Jumlah Berat Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select SUM(Berat_Jualan) from 42_tempahan_siap where status_invoice = 1 AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select SUM(Berat_Jualan) from 42_tempahan_siap where status_invoice = 1 AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        '#### Jumlah Berat Keseluruhan #### - End
        
        Y = Y + 1
        
        '#### Jumlah Modal Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm101.L9_Text = 0 Then rs.Open "select SUM(harga_dengan_gst) from 42_tempahan_siap where status_invoice = 1 AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        If Frm101.L9_Text = 1 Then rs.Open "select SUM(harga_dengan_gst) from 42_tempahan_siap where status_invoice = 1 AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND purity " & Frm85_LM_SEARCH_1_LOGIC & "'" & Frm85_LM_SEARCH_1 & "' AND kategori_produk " & Frm85_LM_SEARCH_2_LOGIC & "'" & Frm85_LM_SEARCH_2 & "' AND dulang " & Frm85_LM_SEARCH_3_LOGIC & "'" & Frm85_LM_SEARCH_3 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Harga : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Harga : RM " & "0.00"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Harga : RM " & "RM 0.00"
        End If
        '#### Jumlah Modal Keseluruhan #### - End
        
        Y = Y + 4
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
End Sub
Sub Frm85_tempahan_siri()
'on error resume next
Dim Frm85_LM_BERAT As Double
Dim Frm85_LM_HARGA As Double
Dim TA As Date
Dim TM As Date

Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    
    x = 0
    
    Frm85_LM_BERAT = 0
    Frm85_LM_HARGA = 0
    
    If Frm101.L46_Text = "Semua cawangan" Then
    
        Frm85_LM_SEARCH_5 = Null
        Frm85_LM_SEARCH_5_LOGIC = "<>"
        Frm85_LM_SEARCH_6 = Null
        Frm85_LM_SEARCH_6_LOGIC = "<>"
        
    Else
    
        Frm85_LM_SEARCH_5 = Frm101.L46_Text
        Frm85_LM_SEARCH_5_LOGIC = "="
        Frm85_LM_SEARCH_6 = "HQ"
        Frm85_LM_SEARCH_6_LOGIC = "="
        
    End If
    
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh
        .Columns("C").ColumnWidth = 20 'No. Siri Produk
        .Columns("D").ColumnWidth = 40 'Kategori Produk
        .Columns("E").ColumnWidth = 20 'Purity
        .Columns("F").ColumnWidth = 20 'Berat Asal (g)
        .Columns("G").ColumnWidth = 20 'Berat Jualan (g)
        .Columns("H").ColumnWidth = 20 'Upah (RM)
        .Columns("I").ColumnWidth = 20 'Jumlah Harga (RM)
        .Columns("J").ColumnWidth = 20 'Dulang
        .Columns("K").ColumnWidth = 50 'Cawangan
    
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
        
        .Cells(1, 5).Font.Bold = True
        .Cells(1, 5).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 5).HorizontalAlignment = xlCenter
        Next Row
        
        '#### Header Report ###
        .Cells(7, 1) = Frm85.L39_Text 'Report Header"
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh"
        .Cells(8, 3) = "No. Siri Produk"
        .Cells(8, 4) = "Kategori Produk"
        .Cells(8, 5) = "Purity"
        .Cells(8, 6) = "Berat Asal (g)"
        .Cells(8, 7) = "Berat Jualan (g)"
        .Cells(8, 8) = "Upah (RM)"
        .Cells(8, 9) = "Jumlah Harga (RM)"
        .Cells(8, 10) = "Dulang"
        .Cells(8, 11) = "Cawangan"
        
        For i = 1 To 11
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 42_tempahan_siap where status_invoice = 1 AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh) Then .Cells(8 + x, 2) = "'" & rs!tarikh 'Tarikh
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 3) = rs!no_siri_Produk 'No. Siri Produk
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!kategori_Produk) Then .Cells(8 + x, 4) = rs!kategori_Produk 'Kategori Produk
            
            If Not IsNull(rs!purity) Then .Cells(8 + x, 5) = rs!purity 'Purity
            .Cells(8 + x, 5).HorizontalAlignment = xlCenter
            
            .Cells(8 + x, 6).HorizontalAlignment = xlRight
            If Not IsNull(rs!Berat_Asal) Then
                .Cells(8 + x, 6) = Format(rs!Berat_Asal, "#,##0.00") 'Berat Asal (g)
            Else
                .Cells(8 + x, 6) = "0.00" 'Berat (g)
            End If
            .Cells(8 + x, 6).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!berat_jualan) Then
                .Cells(8 + x, 7) = Format(rs!berat_jualan, "#,##0.00") 'Berat Jualan (g)
            Else
                .Cells(8 + x, 7) = "0.00" 'Berat Jualan (g)
            End If
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!UPAH) Then
                .Cells(8 + x, 8) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
            Else
                .Cells(8 + x, 8) = "0.00" 'Upah (RM)
            End If
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_dengan_gst) Then
                .Cells(8 + x, 9) = Format(rs!harga_dengan_gst, "#,##0.00") 'Harga Jualan (RM)
            Else
                .Cells(8 + x, 9) = "0.00" 'Harga Jualan (RM)
            End If
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            
            If Not IsNull(rs!dulang) Then
                .Cells(8 + x, 10) = rs!dulang 'Dulang
                .Cells(8 + x, 10).HorizontalAlignment = xlCenter
            End If
            
            If Not IsNull(rs!cawangan) Then
                .Cells(8 + x, 11) = rs!cawangan 'cawangan
                '.Cells(8 + x, 11).HorizontalAlignment = xlCenter
            End If
            
            For Col = 1 To 11
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 2
        
        '#### Jumlah Bilangan Barang Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select COUNT(no_siri_Produk) from 42_tempahan_siap where status_invoice = 1 AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Bilangan : " & rs(0)
        Else
            .Cells(8 + Y, 1) = "Bilangan : " & 0
        End If
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Bilangan : " & 0
        End If
        '#### Jumlah Bilangan Barang Keseluruhan #### - End
        
        '#### Jumlah Berat Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(Berat_Jualan) from 42_tempahan_siap where status_invoice = 1 AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & Format(rs(0), "#,##0.00 g")
        Else
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
        End If
        '#### Jumlah Berat Keseluruhan #### - End
        
        Y = Y + 1
        
        '#### Jumlah Modal Keseluruhan #### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(harga_dengan_gst) from 42_tempahan_siap where status_invoice = 1 AND (cawangan " & Frm85_LM_SEARCH_5_LOGIC & "'" & Frm85_LM_SEARCH_5 & "' OR cawangan " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "') AND no_siri_Produk='" & Frm101.L5_Text & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            .Cells(8 + Y, 1) = "Jumlah Harga : RM " & Format(rs(0), "#,##0.00")
        Else
            .Cells(8 + Y, 1) = "Jumlah Harga : RM " & "0.00"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If .Cells(8 + Y, 1) = "" Then
            .Cells(8 + Y, 1) = "Jumlah Harga : RM " & "RM 0.00"
        End If
        '#### Jumlah Modal Keseluruhan #### - End
        
        Y = Y + 4
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
End Sub
Sub frm83_recall_data_penerimaan_stok()
'on error resume next
Dim rs2 As ADODB.Recordset

DATA_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from data_database where ID='" & G_ID & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Not IsNull(rs!StatusItem) Then
        
        If rs!StatusItem = "10" Then
            
            GLOBAL_DISABLE = 1
            If Not IsNull(rs!no_pekerja) Then
                Frm83_LM_No_PEKERJA = rs!no_pekerja 'No. Pekerja
            End If
            If Not IsNull(rs!ID) Then Frm83.L13_Text = rs!ID 'No. ID
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

            If Not IsNull(rs!receiving_Status) Then
                If rs!receiving_Status = 0 Then '0 : BK , 1 : Barang Permata , 2 : Buyback BK , 3 : Buyback Barang Permata
                    Frm83.CB7 = 1 'Penerimaan stok baru
                    Frm83.CB4 = 1 'Barang kemas
                    Frm83.CB5 = 0 'Barang permata
                    If Not IsNull(rs!Upah_Jualan) Then Frm83.TB24 = rs!Upah_Jualan 'Upah Jualan Kepada Pelanggan
                    If Not IsNull(rs!Upah_Member) Then Frm83.TB25 = rs!Upah_Member 'Upah Jualan Kepada Ahli / Member
                    If Not IsNull(rs!Upah_Pengedar) Then Frm83.TB26 = rs!Upah_Pengedar 'Upah Jualan Kepada Pengedar
                    If Not IsNull(rs!Upah_RAF) Then Frm83.TB31 = rs!Upah_RAF 'Upah Jualan Kepada RAF
                    If Not IsNull(rs!upah_normal_dealer) Then Frm83.TB32 = rs!upah_normal_dealer 'Upah Jualan Kepada Normal Dealer
                    If Not IsNull(rs!upah_master_dealer) Then Frm83.TB33 = rs!upah_master_dealer 'Upah Jualan Kepada Master Dealer
                ElseIf rs!receiving_Status = 1 Then '0 : BK , 1 : Barang Permata , 2 : Buyback BK , 3 : Buyback Barang Permata
                    Frm83.CB7 = 1 'Penerimaan stok baru
                    Frm83.CB5 = 1 'Barang permata
                    Frm83.CB4 = 0 'Barang kemas
                    If Not IsNull(rs!code_Supplier) Then Frm83.TB24 = rs!code_Supplier 'Harga Jualan Kepada Pelanggan
                    If Not IsNull(rs!HargaJualan_Member) Then Frm83.TB25 = rs!HargaJualan_Member 'Harga Jualan Kepada Ahli / Member
                    If Not IsNull(rs!HargaJualan_Pengedar) Then Frm83.TB26 = rs!HargaJualan_Pengedar 'Harga Jualan Kepada Pengedar
                    If Not IsNull(rs!HargaJualan_RAF) Then Frm83.TB31 = rs!HargaJualan_RAF 'Harga Jualan Kepada RAF
                    If Not IsNull(rs!hargajualan_normal_dealer) Then Frm83.TB32 = rs!hargajualan_normal_dealer 'Harga Jualan Kepada Normal Dealer
                    If Not IsNull(rs!hargajualan_master_dealer) Then Frm83.TB33 = rs!hargajualan_master_dealer 'Harga Jualan Kepada Master Dealer
                ElseIf rs!receiving_Status = 2 Or rs!receiving_Status = 6 Then '0 : BK , 1 : Barang Permata , 2 : Buyback BK , 3 : Buyback Barang Permata
                    Frm83.CB8 = 1 'Buyback / Trade in
                    Frm83.CB4 = 1 'Barang kemas
                    Frm83.CB5 = 0 'Barang permata
                    If Not IsNull(rs!Upah_Jualan) Then Frm83.TB24 = rs!Upah_Jualan 'Upah Jualan Kepada Pelanggan
                    If Not IsNull(rs!Upah_Member) Then Frm83.TB25 = rs!Upah_Member 'Upah Jualan Kepada Ahli / Member
                    If Not IsNull(rs!Upah_Pengedar) Then Frm83.TB26 = rs!Upah_Pengedar 'Upah Jualan Kepada Pengedar
                    If Not IsNull(rs!Upah_RAF) Then Frm83.TB31 = rs!Upah_RAF 'Upah Jualan Kepada RAF
                    If Not IsNull(rs!upah_normal_dealer) Then Frm83.TB32 = rs!upah_normal_dealer 'Upah Jualan Kepada Normal Dealer
                    If Not IsNull(rs!upah_master_dealer) Then Frm83.TB33 = rs!upah_master_dealer 'Upah Jualan Kepada Master Dealer
                ElseIf rs!receiving_Status = 3 Or rs!receiving_Status = 7 Then '0 : BK , 1 : Barang Permata , 2 : Buyback BK , 3 : Buyback Barang Permata
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
            If Not IsNull(rs!SpreadValue) Then
                Frm83.TB19 = rs!SpreadValue 'Spread (%)
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
            
        Else
        
            MsgBox "Status barang ini telah berubah dan anda tidak dibenarkan untuk edit data barang ini." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila batalkan urusan edit data ini dan periksa status terbaru barang ini.", vbExclamation, "Info"
                
        End If
    
    Else
        
        MsgBox "Tiada maklumat status bagi barang ini. Sila keluar dari menu ini dan cuba lagi.", vbExclamation, "Info"
    
    End If
    
Else
    
    MsgBox "Tiada maklumat bagi item ini.", vbExclamation, "Info"
    
End If

rs.Close
Set rs = Nothing

If DATA_FOUND = 1 Then

    '### Carian Maklumat Penjual (Data Pekerja) ### - Start
    DATA_PEKERJA_FOUND = 0
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where NoPekerja='" & Frm83_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        Frm83_LM_MAKLUMAT_PEKERJA = rs!Samaran & "  |  " & rs!NoPekerja
        DATA_PEKERJA_FOUND = 1
    End If
    
    rs.Close
    Set rs = Nothing
    
    If DATA_PEKERJA_FOUND = 1 Then
        On Error GoTo Err_D:
        Frm83.CBB6 = Frm83_LM_MAKLUMAT_PEKERJA
Restore_D:
    End If
    '### Carian Maklumat Penjual (Data Pekerja) ### - End
    
    Frm83.CBB6.Enabled = True
    Frm83.CBB6.BackColor = &HFFFFFF

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
    
    Call frm83_flag_barang_baru
    
    Frm83.CMD20.Visible = False
    Frm83.CMD21.Visible = False
    Frm83.CMD22.Visible = True
    Frm83.CMD23.Visible = True
    
    Frm83.Show
    Frm85.Hide
    
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

Exit Sub
Err_D:
Frm83.CBB6.AddItem Frm83_LM_MAKLUMAT_PEKERJA
Frm83.CBB6 = Frm83_LM_MAKLUMAT_PEKERJA
Resume Restore_D:

End Sub


