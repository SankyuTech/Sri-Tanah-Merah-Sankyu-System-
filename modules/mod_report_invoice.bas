Attribute VB_Name = "mod_report_invoice"
Sub Frm110_initial_setting()
'on error resume next
Frm110.Frame1.Left = 120
Frm110.Frame1.Top = 360
Frm110.Frame2.Left = 120
Frm110.Frame2.Top = 360

Frm110.L10_Text = 0 'Bilangan
Frm110.L11_Text = "0.00" 'Jumlah harga barang (RM)
Frm110.L12_Text = "0.00" 'Jumlah trade in (RM)
Frm110.L13_Text = "0.00" 'Jumlah adjustment (RM)
Frm110.L14_Text = "0.00" 'Jumlah pos laju (RM)
Frm110.L15_Text = "0.00" 'Jumlah bayaran bersih (RM)
Frm110.L16_Text = "0.00" 'Jumlah Kupon Diskaun (RM)
Frm110.L17_Text = "0.00" 'Jumlah Tebusan Mata Ganjaran (RM)
Frm110.L18_Text = "0.00" 'Jumlah Diskaun (RM)

Frm110.Frame1.Visible = False
Frm110.Frame2.Visible = False

Frm110.CBB2.Clear

Frm110.CBB2.AddItem "Semua cawangan"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where status='" & 1 & "' order by cawangan ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!cawangan) Then Frm110.CBB2.AddItem rs!cawangan
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm110.CBB2 = "Semua cawangan"

If MDI_frm1.L4_Text <> "HQ" And MDI_frm1.L4_Text <> "Developer" Then

    Frm110.CBB2 = MDI_frm1.L20_Text
    Frm110.CBB2.Enabled = False
    
Else
    
    Frm110.CBB2.Enabled = True
    
End If
End Sub
Sub Frm110_senarai_jualan_header()
'on error resume next
'#### Header Report Senarai Jualan #### - Start
Frm110.MSFlexGrid1.Clear
Frm110.MSFlexGrid1.Rows = 1
Frm110.MSFlexGrid1.RowHeight(0) = 600
'Frm110.MSFlexGrid1.FormatString = "No.|<No.|<No. ID|<Tarikh|<No. Invoice|<Jumlah (RM)|<Jumlah Harga SR (RM)|<Jumlah Cukai SR (RM)|<Jumlah Harga ZR(L) (RM)|<Jumlah Cukai ZR (RM)"
Frm110.MSFlexGrid1.FormatString = "No.|<No.|<No. ID|<Tarikh|<No. Invoice|<Harga Barang (RM)|<Trade in (RM)|<Adjustment (RM)|<Kupon Diskaun (RM)|<Tebusan Mata Ganjaran (RM)|<Pos Laju (RM)|<Jumlah Bayaran (RM)"

Frm110.MSFlexGrid1.ColWidth(0) = 0 'No.
Frm110.MSFlexGrid1.ColWidth(1) = 600 'No.
Frm110.MSFlexGrid1.ColAlignment(1) = 4

Frm110.MSFlexGrid1.ColWidth(2) = 0 'No. ID

Frm110.MSFlexGrid1.ColWidth(3) = 1100 'Tarikh
Frm110.MSFlexGrid1.ColAlignment(3) = 4

Frm110.MSFlexGrid1.ColWidth(4) = 1800 'No. Invoice
Frm110.MSFlexGrid1.ColAlignment(4) = 4

Frm110.MSFlexGrid1.ColWidth(5) = 1300 'Harga Barang (RM)
Frm110.MSFlexGrid1.ColAlignment(5) = 7

Frm110.MSFlexGrid1.ColWidth(6) = 1200 'Trade in (RM)
Frm110.MSFlexGrid1.ColAlignment(6) = 7

Frm110.MSFlexGrid1.ColWidth(7) = 1100 'Adjustment (RM)
Frm110.MSFlexGrid1.ColAlignment(7) = 7

Frm110.MSFlexGrid1.ColWidth(8) = 1100 'Kupon Diskaun (RM)
Frm110.MSFlexGrid1.ColAlignment(8) = 7

Frm110.MSFlexGrid1.ColWidth(9) = 1200 'Tebusan Mata Ganjaran (RM)
Frm110.MSFlexGrid1.ColAlignment(9) = 7

Frm110.MSFlexGrid1.ColWidth(10) = 1100 'Pos Laju (RM)
Frm110.MSFlexGrid1.ColAlignment(10) = 7

Frm110.MSFlexGrid1.ColWidth(11) = 1300 'Jumlah Bayaran (RM)
Frm110.MSFlexGrid1.ColAlignment(11) = 7

With Frm110.LV1
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear

    Frm110.LV1.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "Tarikh", 1500, 2
    .ColumnHeaders.Add 5, , "No. Invoice", 2200
    .ColumnHeaders.Add 6, , "Harga Barang (RM)", 2000, 1
    .ColumnHeaders.Add 7, , "Trade in (RM)", 1800, 1
    .ColumnHeaders.Add 8, , "Adjustment (RM)", 1800, 1
    .ColumnHeaders.Add 9, , "Kupon Diskaun (RM)", 2000, 1
    .ColumnHeaders.Add 10, , "Tebusan Mata Ganjaran (RM)", 2800, 1
    .ColumnHeaders.Add 11, , "Pos Laju (RM)", 1800, 1
    .ColumnHeaders.Add 12, , "No. Tracking", 1800
    .ColumnHeaders.Add 13, , "Jumlah Bayaran (RM)", 2200, 1
    .ColumnHeaders.Add 14, , "Cawangan", 3000
    .ColumnHeaders.Add 15, , "Tunai (RM)", 1500, 1
    .ColumnHeaders.Add 16, , "Online Transfer (RM)", 2000, 1
    .ColumnHeaders.Add 17, , "Kad Kredit (RM)", 1800, 1
    .ColumnHeaders.Add 18, , "Simpanan Di Kedai (RM)", 2300, 1
    .ColumnHeaders.Add 19, , "Remarks", 5000
    
End With
End Sub
Sub Frm110_senarai_jualan()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim Frm110_LM_TOTAL_PAGE As Double
Dim rs1 As ADODB.Recordset

Frm110_PAGE_SIZE = 39
Frm110_LM_TOTAL_PAGE = 0
x = 0

TM = Frm110.L3_Text 'Tarikh Mula
TA = Frm110.L4_Text 'Tarikh Akhir

If Frm110.L19_Text = "Semua cawangan" Then

    LM_CAWANGAN = Frm110.L19_Text
    
Else

    Frm85_SEARCH_8 = Frm110.L19_Text

End If

If Frm110.L2_Text = "Semua Invoice" Then

    Frm110_LM_SEARCH_1 = 0
    Frm110_LM_SEARCH_2 = 1
    
    Frm110.L9_Text = "Senarai semua jenis invoice , cawangan [" & Frm110.L19_Text & "] dari " & TM & " hingga " & TA & "." 'Report Header
    
ElseIf Frm110.L2_Text = "Invoice Rasmi" Then

    Frm110_LM_SEARCH_1 = 1
    Frm110_LM_SEARCH_2 = 1
    
    Frm110.L9_Text = "Senarai invoice rasmi , cawangan [" & Frm110.L19_Text & "] dari " & TM & " hingga " & TA & "." 'Report Header
    
ElseIf Frm110.L2_Text = "Invoice Tidak Rasmi" Then

    Frm110_LM_SEARCH_1 = 0
    Frm110_LM_SEARCH_2 = 0
    
    Frm110.L9_Text = "Senarai invoice tidak rasmi , cawangan [" & Frm110.L19_Text & "] dari " & TM & " hingga " & TA & "." 'Report Header

End If

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

    Frm110_LM_SEARCH_1 = 1
    Frm110_LM_SEARCH_2 = 1
    
End If

If Frm110.L19_Text = "Semua cawangan" Then

    Frm85_SEARCH_8 = Null
    Frm85_SEARCH_8_LOGIC = "<>"
    Frm85_SEARCH_9 = Null
    Frm85_SEARCH_9_LOGIC = "<>"
    
Else

    Frm85_SEARCH_8 = Frm110.L19_Text
    Frm85_SEARCH_8_LOGIC = "="
    Frm85_SEARCH_9 = "HQ"
    Frm85_SEARCH_9_LOGIC = "="
    
End If


LM_START_ROW = Frm110.L7_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm110_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm110.L8_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm110_PAGE_SIZE
        End If
    End If
End If

Frm110_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from 22_jualan where flag_bayaran='" & "0" & "' AND status = 1 AND (bil_rasmi='" & Frm110_LM_SEARCH_1 & "' OR bil_rasmi='" & Frm110_LM_SEARCH_2 & "') AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by bil_rasmi ASC LIMIT " & LM_START_ROW & "," & Frm110_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
rs.Open "select * from 22_jualan where status = 1 AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (bil_rasmi='" & Frm110_LM_SEARCH_1 & "' OR bil_rasmi='" & Frm110_LM_SEARCH_2 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC , no_resit ASC LIMIT " & LM_START_ROW & "," & Frm110_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm110_LM_PAGE_FOUND = 0 Then
        If Frm110.L8_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm110.L5_Text = Frm110.L5_Text + 1 'Paparan Page ke-xxx
                Frm110_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm110.L5_Text) Then
                    If Frm110.L5_Text <> 1 Then
                        Frm110.L5_Text = Frm110.L5_Text - 1 'Paparan Page ke-xxx
                        Frm110_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm110.L5_Text - 1) * Frm110_PAGE_SIZE) + x

    With Frm110.LV1.ListItems.Add(, , rs!ID)
     
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!tarikh) Then 'Tarikh Jualan
            .ListSubItems.Add , , rs!tarikh
        Else
            .ListSubItems.Add , , ""
        End If
        
        If LM_INVOICE_RASMI = 0 Then
            If Not IsNull(rs!no_resit) Then .ListSubItems.Add , , rs!no_resit 'No. Invoice
        Else
            If Not IsNull(rs!no_invoice_r) Then .ListSubItems.Add , , rs!no_invoice_r  'No. Invoice
        End If

        If Not IsNull(rs!harga_lepas_diskaun) Then 'Harga Barang (RM)
            .ListSubItems.Add , , Format(rs!harga_lepas_diskaun, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        
        If Not IsNull(rs!jumlah_trade_in) Then 'Trade in (RM)
            .ListSubItems.Add , , Format(rs!jumlah_trade_in, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        
        If Not IsNull(rs!adjustment) Then 'Adjustment(RM)
            .ListSubItems.Add , , Format(rs!adjustment, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        
        If Not IsNull(rs!kupon_diskaun) Then 'Kupon Diskaun(RM)
            .ListSubItems.Add , , Format(rs!kupon_diskaun, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        
        If Not IsNull(rs!redeem_point) Then 'Redeem Mata Ganjaran(RM)
            .ListSubItems.Add , , Format(rs!redeem_point, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        
        If Not IsNull(rs!caj_pos) Then 'Pos Laju(RM)
            .ListSubItems.Add , , Format(rs!caj_pos, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        
        If Not IsNull(rs!no_tracking) Then 'No. Tracking Pos Laju
            .ListSubItems.Add , , rs!no_tracking
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!jumlah_perlu_bayar) Then 'Jumlah Bayaran(RM)
            .ListSubItems.Add , , Format(rs!jumlah_perlu_bayar, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        
        If Not IsNull(rs!cawangan) Then 'Cawangan
            .ListSubItems.Add , , rs!cawangan
        Else
            .ListSubItems.Add , , ""
        End If
        If Not IsNull(rs!tunai) Then 'Tunai(RM)
            .ListSubItems.Add , , Format(rs!tunai, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        If Not IsNull(rs!bank_in) Then 'Bank In(RM)
            .ListSubItems.Add , , Format(rs!bank_in, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        If Not IsNull(rs!kad_kredit) Then 'Kad Kredit(RM)
            .ListSubItems.Add , , Format(rs!kad_kredit, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        If Not IsNull(rs!duit_simpanan_kedai) Then 'Simpanan Di Kedai(RM)
            .ListSubItems.Add , , Format(rs!duit_simpanan_kedai, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        
        If Not IsNull(rs!remarks) Then 'Remarks
            .ListSubItems.Add , , rs!remarks
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
'rs.Open "select COUNT(ID) from 22_jualan where flag_bayaran='" & "0" & "' AND status = 1 AND (bil_rasmi='" & Frm110_LM_SEARCH_1 & "' OR bil_rasmi='" & Frm110_LM_SEARCH_2 & "') AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
rs.Open "select COUNT(ID) , SUM(harga_barang_dengan_gst-harga_lepas_diskaun) , SUM(harga_barang_dengan_gst) , SUM(jumlah_trade_in) , SUM(adjustment) , SUM(kupon_diskaun) , SUM(caj_pos) , SUM(redeem_point) , SUM(tunai) , SUM(bank_in) , SUM(kad_kredit) , SUM(duit_simpanan_kedai) from 22_jualan where status = 1 AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (bil_rasmi='" & Frm110_LM_SEARCH_1 & "' OR bil_rasmi='" & Frm110_LM_SEARCH_2 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm110_LM_TOTAL_PAGE = Format(rs(0) / Frm110_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm110_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm110_LM_PAGE = Split(Frm110_LM_TOTAL_PAGE, ".")(0)
        Frm110_LM_PAGE_LEBIHAN = Split(Frm110_LM_TOTAL_PAGE, ".")(1)
        
        If Frm110_LM_PAGE_LEBIHAN <> "00" Then
            Frm110.L6_Text = Frm110_LM_PAGE + 1
        Else
            Frm110.L6_Text = Frm110_LM_PAGE
        End If
        
    Else
    
        Frm110.L6_Text = Frm110_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm110.L6_Text = 0
    End If
Else
    Frm110.L6_Text = 0
End If

If Not IsNull(rs(0)) Then Frm110.L10_Text = rs(0) 'Bilangan
If Not IsNull(rs(1)) Then Frm110.L18_Text = Format(rs(1), "#,##0.00") 'Jumlah diskaun (RM)
If Not IsNull(rs(2)) Then Frm110.L11_Text = Format(rs(2), "#,##0.00") 'Jumlah harga barang (RM)
If Not IsNull(rs(3)) Then Frm110.L12_Text = Format(rs(3), "#,##0.00") 'Jumlah trade in (RM)
If Not IsNull(rs(4)) Then Frm110.L13_Text = Format(rs(4), "#,##0.00") 'Jumlah adjustment (RM)
If Not IsNull(rs(5)) Then Frm110.L16_Text = Format(rs(5), "#,##0.00") 'Jumlah kupon diskaun (RM)
If Not IsNull(rs(6)) Then Frm110.L14_Text = Format(rs(6), "#,##0.00") 'Jumlah pos laju (RM)
If Not IsNull(rs(7)) Then Frm110.L17_Text = Format(rs(7), "#,##0.00") 'Jumlah tebusan mata ganjaran (RM)
If Not IsNull(rs(8)) Then Frm110.L20_Text = Format(rs(8), "#,##0.00")
If Not IsNull(rs(9)) Then Frm110.L21_Text = Format(rs(9), "#,##0.00")
If Not IsNull(rs(10)) Then Frm110.L22_Text = Format(rs(10), "#,##0.00")
If Not IsNull(rs(11)) Then Frm110.L23_Text = Format(rs(11), "#,##0.00")
rs.Close
Set rs = Nothing

If Frm110.L6_Text = vbNullString Then
    Frm110.L6_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah tebusan mata ganjaran ### - Start
'Set rs = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select SUM(nilaian_tebus_point) from 71_tebus_agih_point where status = 1 AND (bil_rasmi='" & Frm110_LM_SEARCH_1 & "' OR bil_rasmi='" & Frm110_LM_SEARCH_2 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
'rs.Open "select SUM(nilaian_tebus_point) from 71_tebus_agih_point where status = 1 AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (bil_rasmi='" & Frm110_LM_SEARCH_1 & "' OR bil_rasmi='" & Frm110_LM_SEARCH_2 & "') AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

'If Not IsNull(rs(0)) Then Frm110.L17_Text = Format(rs(0), "#,##0.00") 'Jumlah tebusan mata ganjaran (RM)

'rs.Close
'Set rs = Nothing
'### Jumlah tebusan mata ganjaran ### - End

'### Jumlah bayaran bersih ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_perlu_bayar) from 22_jualan where status = 1 AND flag_bayaran = 0 AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (bil_rasmi='" & Frm110_LM_SEARCH_1 & "' OR bil_rasmi='" & Frm110_LM_SEARCH_2 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm110.L15_Text = Format(rs(0), "#,##0.00") 'Jumlah bayaran bersih (RM)

rs.Close
Set rs = Nothing
'### Jumlah bayaran bersih ### - End

If x <> 0 Then
    Frm110.L8_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm110.L7_Text = LM_START_ROW 'Titik Pencarian Data
    
    Frm110.Frame1.Visible = False
    Frm110.Frame2.Visible = True
Else
    Frm110.L8_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If

End Sub
Sub Frm110_senarai_jualan2()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim Frm110_LM_TOTAL_PAGE As Double
Dim rs1 As ADODB.Recordset

Frm110_PAGE_SIZE = 39
Frm110_LM_TOTAL_PAGE = 0
x = 0

TM = Frm110.L3_Text 'Tarikh Mula
TA = Frm110.L4_Text 'Tarikh Akhir

If Frm110.L19_Text = "Semua cawangan" Then

    LM_CAWANGAN = Frm110.L19_Text
    
Else

    Frm85_SEARCH_8 = Frm110.L19_Text

End If

If Frm110.L2_Text = "Semua Invoice" Then

    Frm110_LM_SEARCH_1 = 0
    Frm110_LM_SEARCH_2 = 1
    
ElseIf Frm110.L2_Text = "Invoice Rasmi" Then

    Frm110_LM_SEARCH_1 = 1
    Frm110_LM_SEARCH_2 = 1
    
ElseIf Frm110.L2_Text = "Invoice Tidak Rasmi" Then

    Frm110_LM_SEARCH_1 = 0
    Frm110_LM_SEARCH_2 = 0

End If

Frm110.L9_Text = "Carian mengikut keyword [" & Frm110.L24_Text & "]." 'Report Header

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

    Frm110_LM_SEARCH_1 = 1
    Frm110_LM_SEARCH_2 = 1
    
End If

If Frm110.L19_Text = "Semua cawangan" Then

    Frm85_SEARCH_8 = Null
    Frm85_SEARCH_8_LOGIC = "<>"
    Frm85_SEARCH_9 = Null
    Frm85_SEARCH_9_LOGIC = "<>"
    
Else

    Frm85_SEARCH_8 = Frm110.L19_Text
    Frm85_SEARCH_8_LOGIC = "="
    Frm85_SEARCH_9 = "HQ"
    Frm85_SEARCH_9_LOGIC = "="
    
End If


LM_START_ROW = Frm110.L7_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm110_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm110.L8_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm110_PAGE_SIZE
        End If
    End If
End If

Frm110_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where status = 1 AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (bil_rasmi='" & Frm110_LM_SEARCH_1 & "' OR bil_rasmi='" & Frm110_LM_SEARCH_2 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND (no_resit LIKE '%" & Frm110.L24_Text & "%' OR no_tracking LIKE '%" & Frm110.L24_Text & "%' OR remarks LIKE '%" & Frm110.L24_Text & "%') order by tarikh ASC , ID ASC LIMIT " & LM_START_ROW & "," & Frm110_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm110_LM_PAGE_FOUND = 0 Then
        If Frm110.L8_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm110.L5_Text = Frm110.L5_Text + 1 'Paparan Page ke-xxx
                Frm110_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm110.L5_Text) Then
                    If Frm110.L5_Text <> 1 Then
                        Frm110.L5_Text = Frm110.L5_Text - 1 'Paparan Page ke-xxx
                        Frm110_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm110.L5_Text - 1) * Frm110_PAGE_SIZE) + x

    With Frm110.LV1.ListItems.Add(, , rs!ID)
     
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!tarikh) Then 'Tarikh Jualan
            .ListSubItems.Add , , rs!tarikh
        Else
            .ListSubItems.Add , , ""
        End If
        
        If LM_INVOICE_RASMI = 0 Then
            If Not IsNull(rs!no_resit) Then .ListSubItems.Add , , rs!no_resit 'No. Invoice
        Else
            If Not IsNull(rs!no_invoice_r) Then .ListSubItems.Add , , rs!no_invoice_r  'No. Invoice
        End If

        If Not IsNull(rs!harga_lepas_diskaun) Then 'Harga Barang (RM)
            .ListSubItems.Add , , Format(rs!harga_lepas_diskaun, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        
        If Not IsNull(rs!jumlah_trade_in) Then 'Trade in (RM)
            .ListSubItems.Add , , Format(rs!jumlah_trade_in, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        
        If Not IsNull(rs!adjustment) Then 'Adjustment(RM)
            .ListSubItems.Add , , Format(rs!adjustment, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        
        If Not IsNull(rs!kupon_diskaun) Then 'Kupon Diskaun(RM)
            .ListSubItems.Add , , Format(rs!kupon_diskaun, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        
        If Not IsNull(rs!redeem_point) Then 'Redeem Mata Ganjaran(RM)
            .ListSubItems.Add , , Format(rs!redeem_point, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        
        If Not IsNull(rs!caj_pos) Then 'Pos Laju(RM)
            .ListSubItems.Add , , Format(rs!caj_pos, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        
        If Not IsNull(rs!no_tracking) Then 'No. Tracking Pos Laju
            .ListSubItems.Add , , rs!no_tracking
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!jumlah_perlu_bayar) Then 'Jumlah Bayaran(RM)
            .ListSubItems.Add , , Format(rs!jumlah_perlu_bayar, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        
        If Not IsNull(rs!cawangan) Then 'Cawangan
            .ListSubItems.Add , , rs!cawangan
        Else
            .ListSubItems.Add , , ""
        End If
        If Not IsNull(rs!tunai) Then 'Tunai(RM)
            .ListSubItems.Add , , Format(rs!tunai, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        If Not IsNull(rs!bank_in) Then 'Bank In(RM)
            .ListSubItems.Add , , Format(rs!bank_in, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        If Not IsNull(rs!kad_kredit) Then 'Kad Kredit(RM)
            .ListSubItems.Add , , Format(rs!kad_kredit, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        If Not IsNull(rs!duit_simpanan_kedai) Then 'Simpanan Di Kedai(RM)
            .ListSubItems.Add , , Format(rs!duit_simpanan_kedai, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        
        If Not IsNull(rs!remarks) Then 'Remarks
            .ListSubItems.Add , , rs!remarks
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
'rs.Open "select COUNT(ID) from 22_jualan where flag_bayaran='" & "0" & "' AND status = 1 AND (bil_rasmi='" & Frm110_LM_SEARCH_1 & "' OR bil_rasmi='" & Frm110_LM_SEARCH_2 & "') AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
rs.Open "select COUNT(ID) , SUM(harga_barang_dengan_gst-harga_lepas_diskaun) , SUM(harga_barang_dengan_gst) , SUM(jumlah_trade_in) , SUM(adjustment) , SUM(kupon_diskaun) , SUM(caj_pos) , SUM(redeem_point) , SUM(tunai) , SUM(bank_in) , SUM(kad_kredit) , SUM(duit_simpanan_kedai) from 22_jualan where status = 1 AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (bil_rasmi='" & Frm110_LM_SEARCH_1 & "' OR bil_rasmi='" & Frm110_LM_SEARCH_2 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND (no_resit LIKE '%" & Frm110.L24_Text & "%' OR no_tracking LIKE '%" & Frm110.L24_Text & "%' OR remarks LIKE '%" & Frm110.L24_Text & "%')", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm110_LM_TOTAL_PAGE = Format(rs(0) / Frm110_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm110_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm110_LM_PAGE = Split(Frm110_LM_TOTAL_PAGE, ".")(0)
        Frm110_LM_PAGE_LEBIHAN = Split(Frm110_LM_TOTAL_PAGE, ".")(1)
        
        If Frm110_LM_PAGE_LEBIHAN <> "00" Then
            Frm110.L6_Text = Frm110_LM_PAGE + 1
        Else
            Frm110.L6_Text = Frm110_LM_PAGE
        End If
        
    Else
    
        Frm110.L6_Text = Frm110_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm110.L6_Text = 0
    End If
Else
    Frm110.L6_Text = 0
End If

If Not IsNull(rs(0)) Then Frm110.L10_Text = rs(0) 'Bilangan
If Not IsNull(rs(1)) Then Frm110.L18_Text = Format(rs(1), "#,##0.00") 'Jumlah diskaun (RM)
If Not IsNull(rs(2)) Then Frm110.L11_Text = Format(rs(2), "#,##0.00") 'Jumlah harga barang (RM)
If Not IsNull(rs(3)) Then Frm110.L12_Text = Format(rs(3), "#,##0.00") 'Jumlah trade in (RM)
If Not IsNull(rs(4)) Then Frm110.L13_Text = Format(rs(4), "#,##0.00") 'Jumlah adjustment (RM)
If Not IsNull(rs(5)) Then Frm110.L16_Text = Format(rs(5), "#,##0.00") 'Jumlah kupon diskaun (RM)
If Not IsNull(rs(6)) Then Frm110.L14_Text = Format(rs(6), "#,##0.00") 'Jumlah pos laju (RM)
If Not IsNull(rs(7)) Then Frm110.L17_Text = Format(rs(7), "#,##0.00") 'Jumlah tebusan mata ganjaran (RM)
If Not IsNull(rs(8)) Then Frm110.L20_Text = Format(rs(8), "#,##0.00")
If Not IsNull(rs(9)) Then Frm110.L21_Text = Format(rs(9), "#,##0.00")
If Not IsNull(rs(10)) Then Frm110.L22_Text = Format(rs(10), "#,##0.00")
If Not IsNull(rs(11)) Then Frm110.L23_Text = Format(rs(11), "#,##0.00")
rs.Close
Set rs = Nothing

If Frm110.L6_Text = vbNullString Then
    Frm110.L6_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah tebusan mata ganjaran ### - Start
'Set rs = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select SUM(nilaian_tebus_point) from 71_tebus_agih_point where status = 1 AND (bil_rasmi='" & Frm110_LM_SEARCH_1 & "' OR bil_rasmi='" & Frm110_LM_SEARCH_2 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
'rs.Open "select SUM(nilaian_tebus_point) from 71_tebus_agih_point where status = 1 AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (bil_rasmi='" & Frm110_LM_SEARCH_1 & "' OR bil_rasmi='" & Frm110_LM_SEARCH_2 & "') AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

'If Not IsNull(rs(0)) Then Frm110.L17_Text = Format(rs(0), "#,##0.00") 'Jumlah tebusan mata ganjaran (RM)

'rs.Close
'Set rs = Nothing
'### Jumlah tebusan mata ganjaran ### - End

'### Jumlah bayaran bersih ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_perlu_bayar) from 22_jualan where status = 1 AND flag_bayaran = 0 AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (bil_rasmi='" & Frm110_LM_SEARCH_1 & "' OR bil_rasmi='" & Frm110_LM_SEARCH_2 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND (no_resit LIKE '%" & Frm110.L24_Text & "%' OR no_tracking LIKE '%" & Frm110.L24_Text & "%' OR remarks LIKE '%" & Frm110.L24_Text & "%')", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm110.L15_Text = Format(rs(0), "#,##0.00") 'Jumlah bayaran bersih (RM)

rs.Close
Set rs = Nothing
'### Jumlah bayaran bersih ### - End

If x <> 0 Then
    Frm110.L8_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm110.L7_Text = LM_START_ROW 'Titik Pencarian Data
    
    Frm110.Frame1.Visible = False
    Frm110.Frame2.Visible = True
Else
    Frm110.L8_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If

End Sub
Sub frm110_excel_filter()
'on error resume next
Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
Dim TM As Date
Dim TA As Date

Note = "Sistem akan mengambil masa untuk mengeluarkan report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila tunggu sehingga sistem siap keluarkan report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    
    x = 0
    
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 15 'Tarikh
        .Columns("C").ColumnWidth = 20 'No. invoice
        .Columns("D").ColumnWidth = 15 'Harga Barang (RM)
        .Columns("E").ColumnWidth = 15 'Trade in (RM)
        .Columns("F").ColumnWidth = 15 'Adjustment (RM)
        .Columns("G").ColumnWidth = 15 'Diskaun Kupon (RM)
        .Columns("H").ColumnWidth = 15 'Tebusan Mata Ganjaran (RM)
        .Columns("I").ColumnWidth = 15 'Pos Laju (RM)
        .Columns("J").ColumnWidth = 25 'No Tracking (RM)
        .Columns("K").ColumnWidth = 15 'Jumlah Bayaran (RM)
        .Columns("L").ColumnWidth = 25 'Cawangan
        .Columns("M").ColumnWidth = 20 'Tunai
        .Columns("N").ColumnWidth = 20 'Online Transfer
        .Columns("O").ColumnWidth = 20 'Kad Kredit
        .Columns("P").ColumnWidth = 20 'Simpanan Di Kedai
        .Columns("Q").ColumnWidth = 50 'Remarks
        
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
        .Cells(7, 1) = Frm110.L9_Text 'Report Header"

        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh"
        .Cells(8, 3) = "No. invoice"
        .Cells(8, 4) = "Harga Barang (RM)"
        .Cells(8, 5) = "Trade in (RM)"
        .Cells(8, 6) = "Adjustment (RM)"
        .Cells(8, 7) = "Kupon Diskaun (RM)"
        .Cells(8, 8) = "Tebusan Mata Ganjaran (RM)"
        .Cells(8, 9) = "Pos Laju (RM)"
        .Cells(8, 10) = "No. Tracking"
        .Cells(8, 11) = "Jumlah Bayaran (RM)"
        .Cells(8, 12) = "Cawangan"
        .Cells(8, 13) = "Tunai"
        .Cells(8, 14) = "Online Transfer"
        .Cells(8, 15) = "Kad Kredit"
        .Cells(8, 16) = "Simpanan Di Kedai"
        .Cells(8, 17) = "Remarks"
        
        For i = 1 To 17
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
        
        TM = Frm110.L3_Text 'Tarikh Mula
        TA = Frm110.L4_Text 'Tarikh Akhir
        
        If Frm110.L2_Text = "Semua Invoice" Then
        
            Frm110_LM_SEARCH_1 = 0
            Frm110_LM_SEARCH_2 = 1
        
        ElseIf Frm110.L2_Text = "Invoice Rasmi" Then
        
            Frm110_LM_SEARCH_1 = 1
            Frm110_LM_SEARCH_2 = 1
            
        ElseIf Frm110.L2_Text = "Invoice Tidak Rasmi" Then
        
            Frm110_LM_SEARCH_1 = 0
            Frm110_LM_SEARCH_2 = 0
        
        End If
        
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
        
            Frm110_LM_SEARCH_1 = 1
            Frm110_LM_SEARCH_2 = 1
            
        End If
        
        If Frm110.L19_Text = "Semua cawangan" Then
        
            Frm85_SEARCH_8 = Null
            Frm85_SEARCH_8_LOGIC = "<>"
            Frm85_SEARCH_9 = Null
            Frm85_SEARCH_9_LOGIC = "<>"
            
        Else
        
            Frm85_SEARCH_8 = Frm110.L19_Text
            Frm85_SEARCH_8_LOGIC = "="
            Frm85_SEARCH_9 = "HQ"
            Frm85_SEARCH_9_LOGIC = "="
            
        End If
            
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 22_jualan where status = 1 AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (bil_rasmi='" & Frm110_LM_SEARCH_1 & "' OR bil_rasmi='" & Frm110_LM_SEARCH_2 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC , no_resit ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh) Then .Cells(8 + x, 2) = "'" & rs!tarikh 'Tarikh
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If LM_INVOICE_RASMI = 0 Then
                If Not IsNull(rs!no_resit) Then .Cells(8 + x, 3) = rs!no_resit 'No. invoice
            Else
                If Not IsNull(rs!no_invoice_r) Then .Cells(8 + x, 3) = rs!no_invoice_r 'No. invoice
            End If
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            .Cells(8 + x, 4).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_lepas_diskaun) Then 'Harga Barang (RM)
                .Cells(8 + x, 4) = Format(rs!harga_lepas_diskaun, "#,##0.00")
            Else
                .Cells(8 + x, 4) = "0.00"
            End If
            .Cells(8 + x, 4).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 5).HorizontalAlignment = xlRight
            If Not IsNull(rs!jumlah_trade_in) Then 'Trade in (RM)
                .Cells(8 + x, 5) = Format(rs!jumlah_trade_in, "#,##0.00")
            Else
                .Cells(8 + x, 5) = "0.00"
            End If
            .Cells(8 + x, 5).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 6).HorizontalAlignment = xlRight
            If Not IsNull(rs!adjustment) Then 'Adjustment(RM)
                .Cells(8 + x, 6) = Format(rs!adjustment, "#,##0.00")
            Else
                .Cells(8 + x, 6) = "0.00"
            End If
            .Cells(8 + x, 6).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!kupon_diskaun) Then 'Kupon Diskaun(RM)
                .Cells(8 + x, 7) = Format(rs!kupon_diskaun, "#,##0.00")
            Else
                .Cells(8 + x, 7) = "0.00"
            End If
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!redeem_point) Then 'Redeem mata ganjaran(RM)
                .Cells(8 + x, 8) = Format(rs!redeem_point, "#,##0.00")
            Else
                .Cells(8 + x, 8) = "0.00"
            End If
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            
            'Set rs1 = New ADODB.Recordset
            'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            'rs1.Open "select * from 71_tebus_agih_point where no_invoice='" & rs!no_resit & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
            
            'If Not rs1.EOF Then
            
            '    If Not IsNull(rs1!nilaian_tebus_point) Then 'Tebusan mata(RM)
            '        .Cells(8 + x, 8) = Format(rs1!nilaian_tebus_point, "#,##0.00")
            '    Else
            '        .Cells(8 + x, 8) = "0.00"
            '    End If
            '
            'Else
            '
            '    .Cells(8 + x, 8) = "0.00"
            '
            'End If
            
            'rs1.Close
            'Set rs1 = Nothing
            '.Cells(8 + x, 8).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!caj_pos) Then 'Pos Laju(RM)
                .Cells(8 + x, 9) = Format(rs!caj_pos, "#,##0.00")
            Else
                .Cells(8 + x, 9) = "0.00"
            End If
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            
            If Not IsNull(rs!no_tracking) Then 'No. Tracking
                .Cells(8 + x, 10) = rs!no_tracking
            Else
                .Cells(8 + x, 10) = ""
            End If
            
            .Cells(8 + x, 11).HorizontalAlignment = xlRight
            If Not IsNull(rs!jumlah_perlu_bayar) Then 'Jumlah Bayaran(RM)
                If rs!flag_bayaran = 0 Then
                    .Cells(8 + x, 11) = Format(rs!jumlah_perlu_bayar, "#,##0.00")
                Else
                    .Cells(8 + x, 11) = "0.00"
                End If
            Else
                .Cells(8 + x, 11) = "0.00"
            End If
            .Cells(8 + x, 11).NumberFormat = "#,##0.00"
            
            If Not IsNull(rs!cawangan) Then .Cells(8 + x, 12) = rs!cawangan 'Cawangan
            
            .Cells(8 + x, 13).HorizontalAlignment = xlRight
            If Not IsNull(rs!tunai) Then 'Tunai(RM)
                .Cells(8 + x, 13) = Format(rs!tunai, "#,##0.00")
            Else
                .Cells(8 + x, 13) = "0.00"
            End If
            .Cells(8 + x, 13).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 14).HorizontalAlignment = xlRight
            If Not IsNull(rs!bank_in) Then 'Bank In(RM)
                .Cells(8 + x, 14) = Format(rs!bank_in, "#,##0.00")
            Else
                .Cells(8 + x, 14) = "0.00"
            End If
            .Cells(8 + x, 14).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 15).HorizontalAlignment = xlRight
            If Not IsNull(rs!kad_kredit) Then 'Kad Kredit(RM)
                .Cells(8 + x, 15) = Format(rs!kad_kredit, "#,##0.00")
            Else
                .Cells(8 + x, 15) = "0.00"
            End If
            .Cells(8 + x, 15).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 16).HorizontalAlignment = xlRight
            If Not IsNull(rs!duit_simpanan_kedai) Then 'Simpanan Di Kedai(RM)
                .Cells(8 + x, 16) = Format(rs!duit_simpanan_kedai, "#,##0.00")
            Else
                .Cells(8 + x, 16) = "0.00"
            End If
            .Cells(8 + x, 16).NumberFormat = "#,##0.00"
            
            If Not IsNull(rs!remarks) Then
                .Cells(8 + x, 17) = rs!remarks
            Else
                .Cells(8 + x, 17) = ""
            End If
            
            For Col = 1 To 17
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 2
        .Cells(8 + Y, 1) = "Bilangan invoice : " & Frm110.L10_Text
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah harga barang : RM " & Frm110.L11_Text
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah Diskaun : RM " & Frm110.L18_Text
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah trade in : RM " & Frm110.L12_Text
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah adjustment : RM " & Frm110.L13_Text
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah kupon Diskaun : RM " & Frm110.L16_Text
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah tebus mata ganjaran : RM " & Frm110.L17_Text
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah pos laju : RM " & Frm110.L14_Text
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah harga bersih : RM " & Frm110.L15_Text
        Y = Y + 2
        .Cells(8 + Y, 1) = Frm110.Label6
        
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
End If
End Sub
Sub frm110_excel_keyword()
'on error resume next
Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
Dim TM As Date
Dim TA As Date

Note = "Sistem akan mengambil masa untuk mengeluarkan report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila tunggu sehingga sistem siap keluarkan report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    
    x = 0
    
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 15 'Tarikh
        .Columns("C").ColumnWidth = 20 'No. invoice
        .Columns("D").ColumnWidth = 15 'Harga Barang (RM)
        .Columns("E").ColumnWidth = 15 'Trade in (RM)
        .Columns("F").ColumnWidth = 15 'Adjustment (RM)
        .Columns("G").ColumnWidth = 15 'Diskaun Kupon (RM)
        .Columns("H").ColumnWidth = 15 'Tebusan Mata Ganjaran (RM)
        .Columns("I").ColumnWidth = 15 'Pos Laju (RM)
        .Columns("J").ColumnWidth = 25 'No Tracking (RM)
        .Columns("K").ColumnWidth = 15 'Jumlah Bayaran (RM)
        .Columns("L").ColumnWidth = 25 'Cawangan
        .Columns("M").ColumnWidth = 20 'Tunai
        .Columns("N").ColumnWidth = 20 'Online Transfer
        .Columns("O").ColumnWidth = 20 'Kad Kredit
        .Columns("P").ColumnWidth = 20 'Simpanan Di Kedai
        .Columns("Q").ColumnWidth = 50 'Remarks
        
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
        .Cells(7, 1) = Frm110.L9_Text 'Report Header"

        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh"
        .Cells(8, 3) = "No. invoice"
        .Cells(8, 4) = "Harga Barang (RM)"
        .Cells(8, 5) = "Trade in (RM)"
        .Cells(8, 6) = "Adjustment (RM)"
        .Cells(8, 7) = "Kupon Diskaun (RM)"
        .Cells(8, 8) = "Tebusan Mata Ganjaran (RM)"
        .Cells(8, 9) = "Pos Laju (RM)"
        .Cells(8, 10) = "No. Tracking"
        .Cells(8, 11) = "Jumlah Bayaran (RM)"
        .Cells(8, 12) = "Cawangan"
        .Cells(8, 13) = "Tunai"
        .Cells(8, 14) = "Online Transfer"
        .Cells(8, 15) = "Kad Kredit"
        .Cells(8, 16) = "Simpanan Di Kedai"
        .Cells(8, 17) = "Remarks"
        
        For i = 1 To 17
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
        
        TM = Frm110.L3_Text 'Tarikh Mula
        TA = Frm110.L4_Text 'Tarikh Akhir
        
        If Frm110.L2_Text = "Semua Invoice" Then
        
            Frm110_LM_SEARCH_1 = 0
            Frm110_LM_SEARCH_2 = 1
        
        ElseIf Frm110.L2_Text = "Invoice Rasmi" Then
        
            Frm110_LM_SEARCH_1 = 1
            Frm110_LM_SEARCH_2 = 1
            
        ElseIf Frm110.L2_Text = "Invoice Tidak Rasmi" Then
        
            Frm110_LM_SEARCH_1 = 0
            Frm110_LM_SEARCH_2 = 0
        
        End If
        
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
        
            Frm110_LM_SEARCH_1 = 1
            Frm110_LM_SEARCH_2 = 1
            
        End If
        
        If Frm110.L19_Text = "Semua cawangan" Then
        
            Frm85_SEARCH_8 = Null
            Frm85_SEARCH_8_LOGIC = "<>"
            Frm85_SEARCH_9 = Null
            Frm85_SEARCH_9_LOGIC = "<>"
            
        Else
        
            Frm85_SEARCH_8 = Frm110.L19_Text
            Frm85_SEARCH_8_LOGIC = "="
            Frm85_SEARCH_9 = "HQ"
            Frm85_SEARCH_9_LOGIC = "="
            
        End If
            
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        'rs.Open "select * from 22_jualan where status = 1 AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (bil_rasmi='" & Frm110_LM_SEARCH_1 & "' OR bil_rasmi='" & Frm110_LM_SEARCH_2 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
        rs.Open "select * from 22_jualan where status = 1 AND (cawangan " & Frm85_SEARCH_8_LOGIC & "'" & Frm85_SEARCH_8 & "' OR cawangan " & Frm85_SEARCH_9_LOGIC & "'" & Frm85_SEARCH_9 & "') AND (bil_rasmi='" & Frm110_LM_SEARCH_1 & "' OR bil_rasmi='" & Frm110_LM_SEARCH_2 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') AND (no_resit LIKE '%" & Frm110.L24_Text & "%' OR no_tracking LIKE '%" & Frm110.L24_Text & "%' OR remarks LIKE '%" & Frm110.L24_Text & "%') order by tarikh ASC , ID ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh) Then .Cells(8 + x, 2) = "'" & rs!tarikh 'Tarikh
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If LM_INVOICE_RASMI = 0 Then
                If Not IsNull(rs!no_resit) Then .Cells(8 + x, 3) = rs!no_resit 'No. invoice
            Else
                If Not IsNull(rs!no_invoice_r) Then .Cells(8 + x, 3) = rs!no_invoice_r 'No. invoice
            End If
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            .Cells(8 + x, 4).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_lepas_diskaun) Then 'Harga Barang (RM)
                .Cells(8 + x, 4) = Format(rs!harga_lepas_diskaun, "#,##0.00")
            Else
                .Cells(8 + x, 4) = "0.00"
            End If
            .Cells(8 + x, 4).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 5).HorizontalAlignment = xlRight
            If Not IsNull(rs!jumlah_trade_in) Then 'Trade in (RM)
                .Cells(8 + x, 5) = Format(rs!jumlah_trade_in, "#,##0.00")
            Else
                .Cells(8 + x, 5) = "0.00"
            End If
            .Cells(8 + x, 5).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 6).HorizontalAlignment = xlRight
            If Not IsNull(rs!adjustment) Then 'Adjustment(RM)
                .Cells(8 + x, 6) = Format(rs!adjustment, "#,##0.00")
            Else
                .Cells(8 + x, 6) = "0.00"
            End If
            .Cells(8 + x, 6).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!kupon_diskaun) Then 'Kupon Diskaun(RM)
                .Cells(8 + x, 7) = Format(rs!kupon_diskaun, "#,##0.00")
            Else
                .Cells(8 + x, 7) = "0.00"
            End If
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!redeem_point) Then 'Redeem mata ganjaran(RM)
                .Cells(8 + x, 8) = Format(rs!redeem_point, "#,##0.00")
            Else
                .Cells(8 + x, 8) = "0.00"
            End If
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            
            'Set rs1 = New ADODB.Recordset
            'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            'rs1.Open "select * from 71_tebus_agih_point where no_invoice='" & rs!no_resit & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
            
            'If Not rs1.EOF Then
            
            '    If Not IsNull(rs1!nilaian_tebus_point) Then 'Tebusan mata(RM)
            '        .Cells(8 + x, 8) = Format(rs1!nilaian_tebus_point, "#,##0.00")
            '    Else
            '        .Cells(8 + x, 8) = "0.00"
            '    End If
            '
            'Else
            '
            '    .Cells(8 + x, 8) = "0.00"
            '
            'End If
            
            'rs1.Close
            'Set rs1 = Nothing
            '.Cells(8 + x, 8).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!caj_pos) Then 'Pos Laju(RM)
                .Cells(8 + x, 9) = Format(rs!caj_pos, "#,##0.00")
            Else
                .Cells(8 + x, 9) = "0.00"
            End If
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            
            If Not IsNull(rs!no_tracking) Then 'No. Tracking
                .Cells(8 + x, 10) = rs!no_tracking
            Else
                .Cells(8 + x, 10) = ""
            End If
            
            .Cells(8 + x, 11).HorizontalAlignment = xlRight
            If Not IsNull(rs!jumlah_perlu_bayar) Then 'Jumlah Bayaran(RM)
                If rs!flag_bayaran = 0 Then
                    .Cells(8 + x, 11) = Format(rs!jumlah_perlu_bayar, "#,##0.00")
                Else
                    .Cells(8 + x, 11) = "0.00"
                End If
            Else
                .Cells(8 + x, 11) = "0.00"
            End If
            .Cells(8 + x, 11).NumberFormat = "#,##0.00"
            
            If Not IsNull(rs!cawangan) Then .Cells(8 + x, 12) = rs!cawangan 'Cawangan
            
            .Cells(8 + x, 13).HorizontalAlignment = xlRight
            If Not IsNull(rs!tunai) Then 'Tunai(RM)
                .Cells(8 + x, 13) = Format(rs!tunai, "#,##0.00")
            Else
                .Cells(8 + x, 13) = "0.00"
            End If
            .Cells(8 + x, 13).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 14).HorizontalAlignment = xlRight
            If Not IsNull(rs!bank_in) Then 'Bank In(RM)
                .Cells(8 + x, 14) = Format(rs!bank_in, "#,##0.00")
            Else
                .Cells(8 + x, 14) = "0.00"
            End If
            .Cells(8 + x, 14).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 15).HorizontalAlignment = xlRight
            If Not IsNull(rs!kad_kredit) Then 'Kad Kredit(RM)
                .Cells(8 + x, 15) = Format(rs!kad_kredit, "#,##0.00")
            Else
                .Cells(8 + x, 15) = "0.00"
            End If
            .Cells(8 + x, 15).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 16).HorizontalAlignment = xlRight
            If Not IsNull(rs!duit_simpanan_kedai) Then 'Simpanan Di Kedai(RM)
                .Cells(8 + x, 16) = Format(rs!duit_simpanan_kedai, "#,##0.00")
            Else
                .Cells(8 + x, 16) = "0.00"
            End If
            .Cells(8 + x, 16).NumberFormat = "#,##0.00"
            
            If Not IsNull(rs!remarks) Then
                .Cells(8 + x, 17) = rs!remarks
            Else
                .Cells(8 + x, 17) = ""
            End If
            
            For Col = 1 To 17
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 2
        .Cells(8 + Y, 1) = "Bilangan invoice : " & Frm110.L10_Text
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah harga barang : RM " & Frm110.L11_Text
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah Diskaun : RM " & Frm110.L18_Text
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah trade in : RM " & Frm110.L12_Text
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah adjustment : RM " & Frm110.L13_Text
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah kupon Diskaun : RM " & Frm110.L16_Text
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah tebus mata ganjaran : RM " & Frm110.L17_Text
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah pos laju : RM " & Frm110.L14_Text
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jumlah harga bersih : RM " & Frm110.L15_Text
        Y = Y + 2
        .Cells(8 + Y, 1) = Frm110.Label6
        
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
End If
End Sub




