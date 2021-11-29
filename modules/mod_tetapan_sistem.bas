Attribute VB_Name = "Mod_tetapan_sistem"
Sub Frm111_setting()
'On Error Resume Next
If MDI_frm1.L20_Text = "Semua cawangan" Then
    LM_CAWANGAN = "HQ"
Else
    LM_CAWANGAN = MDI_frm1.L20_Text
End If

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting where Default1='" & LM_CAWANGAN & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Not IsNull(rs!spread_Cash_Trade_In) Then Frm111.TB16 = rs!spread_Cash_Trade_In 'Spread Trade In
    'If Not IsNull(rs!cas_Kad_Kredit) Then Frm111.TB17 = rs!cas_Kad_Kredit 'Cas Kad Kredit
    'If Not IsNull(rs!cas_debit_kad) Then Frm111.TB27 = rs!cas_debit_kad 'Cas Kad Debit
    'If Not IsNull(rs!komisen) Then Frm111.TB20 = rs!komisen 'Komisen Per Gram (RM)

    If Not IsNull(rs!komisen) Then 'Kadar Komisen Barang Kemas (%)
        Frm111.TB20 = rs!komisen
    Else
        Frm111.TB20 = "0.00"
    End If
    If Not IsNull(rs!komisen_permata) Then 'Kadar Komisen Barang Permata (%)
        Frm111.TB29 = rs!komisen_permata
    Else
        Frm111.TB29 = "0.00"
    End If
    If Not IsNull(rs!harga_999) Then 'Harga emas per gram 999.9
        Frm111.TB33 = rs!harga_999
    Else
        Frm111.TB33 = "0.00"
    End If
    If Not IsNull(rs!komisen_per_gram) Then 'Kadar Komisen Staff Per Gram
        Frm111.TB52 = rs!komisen_per_gram
    Else
        Frm111.TB52 = "0.00"
    End If
    If Not IsNull(rs!harga_beli_999) Then 'Harga emas per gram 999.9
        Frm111.TB51 = rs!harga_beli_999
    Else
        Frm111.TB51 = "0.00"
    End If
    If Not IsNull(rs!kadar_komisyen_upah) Then 'Kadar komisyen upah kepada agen dropship (%)
        Frm111.TB34 = rs!kadar_komisyen_upah
    Else
        Frm111.TB34 = 0
    End If
    
    If Not IsNull(rs!limit_per_gram) Then Frm111.TB25 = Format(rs!limit_per_gram, "0.00") 'Had Kadar Penurunan Harga Jualan Per Gram
    If Not IsNull(rs!limit_per_item) Then Frm111.TB26 = Format(rs!limit_per_item, "0.00") 'Had Kadar Penurunan Harga Jualan Per Item
    If Not IsNull(rs!potongan_trade_in) Then Frm111.TB28 = rs!potongan_trade_in 'Kadar Potongan Trade In - Jika Kedai Perlu Bayar (%)
    If Not IsNull(rs!gst_jualan_included) Then
        If rs!gst_jualan_included = 1 Then
            Frm111.CB7 = 1
        ElseIf rs!gst_jualan_included = 0 Then
            Frm111.CB7 = 0
        End If
    Else
        Frm111.CB7 = 0
    End If
    If Not IsNull(rs!gst_arinashi) Then
        If rs!gst_arinashi = 1 Then
            Frm111.CB3 = 1
        Else
            Frm111.CB3 = 0
            Frm111.TB24 = vbNullString
        End If
        If Not IsNull(rs!gst_value) Then Frm111.TB24 = rs!gst_value 'Kadar GST
    End If
    If Not IsNull(rs!gst_arinashi_belian) Then
        If rs!gst_arinashi_belian = 1 Then
            Frm111.CB5 = 1
        Else
            Frm111.CB5 = 0
            Frm111.TB24 = vbNullString
        End If
        If Not IsNull(rs!gst_value) Then Frm111.TB24 = rs!gst_value 'Kadar GST
    End If
    If Not IsNull(rs!ScannerMode) Then
        If rs!ScannerMode = 1 Then
            Frm111.CB1 = 1
        Else
            Frm111.CB1 = 0
        End If
    End If
    If Not IsNull(rs!BarcodeYesNo) Then
        If rs!BarcodeYesNo = 1 Then
            Frm111.CB2 = 1
        Else
            Frm111.CB2 = 0
        End If
    End If
    If Not IsNull(rs!printer_mode_ti) Then
        If rs!printer_mode_ti = 1 Then
            Frm111.CB15 = 1
        Else
            Frm111.CB15 = 0
        End If
    End If
    If Not IsNull(rs!flag_upah) Then
        If rs!flag_upah = 1 Then
            Frm111.CB4 = 1
        Else
            Frm111.CB4 = 0
        End If
    End If
    If Not IsNull(rs!diskaun_ari_nashi) Then
        If rs!diskaun_ari_nashi = 1 Then
            Frm111.CB6 = 1
        Else
            Frm111.CB6 = 0
        End If
        If Not IsNull(rs!diskaun) Then Frm111.TB32 = rs!diskaun 'Kadar Diskaun (%)
    Else
        Frm111.CB6 = 0
        Frm111.TB32 = 0
    End If
    If Not IsNull(rs!kiraan_upah) Then
        
        If rs!kiraan_upah = 0 Then
            Frm111.CB8 = 1
            Frm111.CB9 = 0
        ElseIf rs!kiraan_upah = 1 Then
            Frm111.CB9 = 1
            Frm111.CB8 = 0
        End If
    
    End If
    
    If Not IsNull(rs!upah_supplier) Then
        
        If rs!upah_supplier = 0 Then
            Frm111.CB12 = 1
            Frm111.CB13 = 0
        ElseIf rs!upah_supplier = 1 Then
            Frm111.CB13 = 1
            Frm111.CB12 = 0
        End If
    
    End If
    
    If Not IsNull(rs!jenis_header) Then
        
        If rs!jenis_header = 0 Then
            
            Frm111.CB10 = 1
            Frm111.CB11 = 0
            
        ElseIf rs!jenis_header = 1 Then
            
            Frm111.CB10 = 0
            Frm111.CB11 = 1
            
        End If
    
    Else
    
        Frm111.CB10 = 0
        Frm111.CB11 = 1
        
    End If

    If Not IsNull(rs!kupon_diskaun) Then
        If IsNumeric(rs!kupon_diskaun) Then
            Frm111.TB35 = Format(rs!kupon_diskaun, "#,##0.00")
        Else
            Frm111.TB35 = Format(0, "#,##0.00")
        End If
    Else
        Frm111.TB35 = Format(0, "#,##0.00")
    End If
    
    If Not IsNull(rs!tael) Then Frm111.TB40 = rs!tael 'Tael
    If Not IsNull(rs!public) Then Frm111.TB41 = rs!public 'Public
    If Not IsNull(rs!sa) Then Frm111.TB42 = rs!sa 'SA
    If Not IsNull(rs!rate_trade_in) Then Frm111.TB54 = Format(rs!rate_trade_in, "#,##0.00")
    If Not IsNull(rs!rate_buyback) Then Frm111.TB55 = Format(rs!rate_buyback, "#,##0.00")
    If Not IsNull(rs!rate_caj_pertukaran) Then Frm111.TB56 = Format(rs!rate_caj_pertukaran, "#,##0.00")

    If Not IsNull(rs!top_margin) Then 'Top Margin
        Frm111.TB53 = rs!top_margin
    Else
        Frm111.TB53 = 0
    End If
    If Not IsNull(rs!invoice_tak_rasmi) Then
    
        If rs!invoice_tak_rasmi = 0 Then
            Frm111.CB14 = 1
        ElseIf rs!invoice_tak_rasmi = 1 Then
            Frm111.CB14 = 0
        End If
        
    End If
    
End If

rs.Close
Set rs = Nothing
End Sub
Sub Frm111_setting2()
'On Error Resume Next
Frm111.TB43 = G_PEMALAR_BONUS_BIASA 'Kadar perolehan mata ganjaran (ahli biasa)
Frm111.TB44 = G_PEMALAR_TEBUS_BIASA 'Kadar tebusan mata ganjaran (ahli biasa)
Frm111.TB45 = G_PEMALAR_BONUS_SILVER 'Kadar perolehan mata ganjaran (silver)
Frm111.TB46 = G_PEMALAR_TEBUS_SILVER 'Kadar tebusan mata ganjaran (silver)
Frm111.TB47 = G_PEMALAR_BONUS_GOLD 'Kadar perolehan mata ganjaran (gold)
Frm111.TB48 = G_PEMALAR_TEBUS_GOLD 'Kadar tebusan mata ganjaran (gold)
Frm111.TB49 = G_PEMALAR_BONUS_PLATINUM 'Kadar perolehan mata ganjaran (platinum)
Frm111.TB50 = G_PEMALAR_TEBUS_PLATINUM 'Kadar tebusan mata ganjaran (platinum)

Exit Sub

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting where Default1='" & "Default" & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    
    If Not IsNull(rs!pemalar_bonus_biasa) Then Frm111.TB43 = rs!pemalar_bonus_biasa 'Kadar perolehan mata ganjaran (ahli biasa)
    If Not IsNull(rs!pemalar_tebus_bonus_biasa) Then Frm111.TB44 = rs!pemalar_tebus_bonus_biasa 'Kadar tebusan mata ganjaran (ahli biasa)
                
    If Not IsNull(rs!pemalar_bonus_silver) Then Frm111.TB45 = rs!pemalar_bonus_silver 'Kadar perolehan mata ganjaran (silver)
    If Not IsNull(rs!pemalar_tebus_bonus_silver) Then Frm111.TB46 = rs!pemalar_tebus_bonus_silver 'Kadar tebusan mata ganjaran (silver)
    
    If Not IsNull(rs!pemalar_bonus_gold) Then Frm111.TB47 = rs!pemalar_bonus_gold 'Kadar perolehan mata ganjaran (gold)
    If Not IsNull(rs!pemalar_tebus_bonus_gold) Then Frm111.TB48 = rs!pemalar_tebus_bonus_gold 'Kadar tebusan mata ganjaran (gold)
    
    If Not IsNull(rs!pemalar_bonus_platinum) Then Frm111.TB49 = rs!pemalar_bonus_platinum 'Kadar perolehan mata ganjaran (platinum)
    If Not IsNull(rs!pemalar_tebus_bonus_platinum) Then Frm111.TB50 = rs!pemalar_tebus_bonus_platinum 'Kadar tebusan mata ganjaran (platinum)
End If

rs.Close
Set rs = Nothing
End Sub
Sub Frm111_initial_setting()
'On Error Resume Next
Frm111.Frame14.Top = 360
Frm111.Frame14.Left = 120
Frm111.Frame15.Top = 360
Frm111.Frame15.Left = 120
Frm111.Frame1.Top = 360
Frm111.Frame1.Left = 120

Frm111.Frame14.Visible = False
Frm111.Frame15.Visible = False
Frm111.Frame1.Visible = False

Frm111.TB32 = 0 'Kadar Diskaun (%)
Frm111.TB34 = 0 'Kadar komisyen upah kepada agen dropship (%)
Frm111.TB35 = "0.00" 'Kadar diskaun per gram bagi penggunaan kupon diskaun (RM/g)

Frm111.TB40 = 0
Frm111.TB41 = 0
Frm111.TB42 = 0

Frm111.TB43 = 0
Frm111.TB44 = 0
Frm111.TB45 = 0
Frm111.TB46 = 0
Frm111.TB47 = 0
Frm111.TB48 = 0
Frm111.TB49 = 0
Frm111.TB50 = 0
Frm111.TB53 = 0

Frm111.TB54 = 0
Frm111.TB55 = 0
Frm111.TB56 = 0

Frm111.CB3 = 0
Frm111.CB7 = 0
Frm111.CB10 = 0
Frm111.CB11 = 1
Frm111.CB12 = 0
Frm111.CB13 = 0
Frm111.TB24 = vbNullString 'Kadar GST
Frm111.TB25 = vbNullString 'Kadar Limit Penurunan Harga Barang Per Gram
Frm111.TB26 = vbNullString 'Kadar Limit Penurunan Harga Barang Per Item
'Frm111.TB27 = vbNullString 'Cas Kad Debit
Frm111.TB51 = vbNullString 'Harga GDN
Frm111.TB28 = vbNullString 'Kadar Potongan Trade In - Jika Kedai Perlu Bayar (%)
End Sub
Sub Frm111_initial_setting2()
'On Error Resume Next
Frm111.TB36 = vbNullString 'Jenis kad kredit / debit
Frm111.TB37 = "0.00" 'Cas kad kredit / debit

Frm111.L14_Text = vbNullString
Frm111.L13_Text.Visible = False

Frm111.CMD2.Visible = True
Frm111.CMD3.Visible = False
Frm111.CMD4.Visible = False
End Sub
Sub Frm111_senarai_jenis_kad_header()
'on error resume next
'#### Header #### - Start
Frm111.MSFlexGrid1.Clear
Frm111.MSFlexGrid1.Rows = 1
Frm111.MSFlexGrid1.RowHeight(0) = 800
Frm111.MSFlexGrid1.FormatString = "<No.|<No.|<No. ID|<Jenis Kad|<Caj Perkhidmatan (%)"

Frm111.MSFlexGrid1.ColWidth(0) = 0 'No.
Frm111.MSFlexGrid1.ColWidth(1) = 600 'No.
Frm111.MSFlexGrid1.ColAlignment(1) = 4
Frm111.MSFlexGrid1.ColWidth(2) = 0 'No. ID
Frm111.MSFlexGrid1.ColWidth(3) = 7800 'Jenis Kad
Frm111.MSFlexGrid1.ColWidth(4) = 1500 'Caj Perkhidmatan (%)
Frm111.MSFlexGrid1.ColAlignment(4) = 4
'#### Header #### - End
End Sub
Sub Frm111_senarai_jenis_kad()
'on error resume next
Dim TA As Date
Dim TM As Date
Dim Frm111_LM_TOTAL_PAGE As Double

x = 0
Y = 0

Frm111_PAGE_SIZE = 28
Frm111_LM_TOTAL_PAGE = 0

LM_START_ROW = Frm111.L9_Text 'Start row

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm111_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm111.L10_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm111_PAGE_SIZE
        End If
    End If
End If

Frm111_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 74_cas_kad_kredit where status='" & 1 & "' order by jenis_kad ASC LIMIT " & LM_START_ROW & "," & Frm111_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm111_LM_PAGE_FOUND = 0 Then
        If Frm111.L10_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm111.L11_Text = Frm111.L11_Text + 1
                Frm111_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm111.L11_Text) Then
                    If Frm111.L11_Text <> 1 Then
                        Frm111.L11_Text = Frm111.L11_Text - 1
                        Frm111_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm111.L11_Text - 1) * Frm111_PAGE_SIZE) + x
    
    Frm111.MSFlexGrid1.Rows = x + 1
    Frm111.MSFlexGrid1.TextMatrix(x, 0) = x 'No.
    Frm111.MSFlexGrid1.TextMatrix(x, 1) = Y 'No.
    Frm111.MSFlexGrid1.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!jenis_kad) Then Frm111.MSFlexGrid1.TextMatrix(x, 3) = rs!jenis_kad 'Jenis kad
    If Not IsNull(rs!cas_kad) Then Frm111.MSFlexGrid1.TextMatrix(x, 4) = Format(rs!cas_kad, "#,##0.00") 'Caj perkhidmatan kad kredit
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 74_cas_kad_kredit where status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm111_LM_TOTAL_PAGE = Format(rs(0) / Frm111_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm111_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm111_LM_PAGE = Split(Frm111_LM_TOTAL_PAGE, ".")(0)
        Frm111_LM_PAGE_LEBIHAN = Split(Frm111_LM_TOTAL_PAGE, ".")(1)
        
        If Frm111_LM_PAGE_LEBIHAN <> "00" Then
            Frm111.L12_Text = Frm111_LM_PAGE + 1
        Else
            Frm111.L12_Text = Frm111_LM_PAGE
        End If
        
    Else
    
        Frm111.L12_Text = Frm111_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm111.L12_Text = 0
    End If
Else
    Frm111.L12_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm111.L12_Text = vbNullString Then
    Frm111.L12_Text = 0
End If
'### Jumlah Data ### - End

If x <> 0 Then
    Frm111.L9_Text = LM_START_ROW
    Frm111.L10_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm111.L10_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
