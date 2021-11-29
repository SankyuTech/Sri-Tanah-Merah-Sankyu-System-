Attribute VB_Name = "mod_report_kewangan"
Sub Frm105_initial_setting()
'on error resume next
Frm105.Pic1.Left = 120
Frm105.Pic1.Top = 360
Frm105.Pic2.Left = 120
Frm105.Pic2.Top = 360
Frm105.Pic9.Left = 120
Frm105.Pic9.Top = 360

Frm105.Pic1.Visible = False
Frm105.Pic2.Visible = False
Frm105.Pic9.Visible = False

Frm105.L10_Text = vbNullString 'Header : Senarai jualan
Frm105.L22_Text = vbNullString 'Header : Senarai servis
Frm105.L34_Text = vbNullString 'Header : Senarai bayaran ansuran
Frm105.L46_Text = vbNullString 'Header : Senarai bayaran tempahan
Frm105.L58_Text = vbNullString 'Header : Senarai kemasukkan duit ke kedai
Frm105.L65_Text = vbNullString 'Header : Senarai simpanan duit di kedai oleh pelanggan
Frm105.L72_Text = vbNullString 'Header : Senarai belian trade in
Frm105.L79_Text = vbNullString 'Header : Senarai belian tukaran barang oleh agen
Frm105.L86_Text = vbNullString 'Header : Ambilan tunai dari kedai
Frm105.L93_Text = vbNullString 'Header : Perbelanjaan kedai
Frm105.L100_Text = vbNullString 'Header : Bayaran gaji

Frm105.L17_Text = 0 'Rekod Jualan : Paparan page
Frm105.L18_Text = 0 'Rekod Jualan : Jumlah page
Frm105.L29_Text = 0 'Rekod servis : Paparan page
Frm105.L30_Text = 0 'Rekod servis : Jumlah page
Frm105.L41_Text = 0 'Rekod ansuran : Paparan page
Frm105.L42_Text = 0 'Rekod ansuran : Jumlah page
Frm105.L53_Text = 0 'Rekod tempahan : Paparan page
Frm105.L54_Text = 0 'Rekod tempahan : Jumlah page
Frm105.L60_Text = 0 'Rekod kemasukkan duit ke kedai : Paparan page
Frm105.L61_Text = 0 'Rekod kemasukkan duit ke kedai : Jumlah page
Frm105.L67_Text = 0 'Senarai simpanan duit di kedai oleh pelanggan : Paparan page
Frm105.L68_Text = 0 'Senarai simpanan duit di kedai oleh pelanggan : Jumlah page
Frm105.L74_Text = 0 'Senarai belian trade in : Paparan page
Frm105.L75_Text = 0 'Senarai belian trade in : Jumlah page
Frm105.L81_Text = 0 'Senarai belian tukaran barang oleh agen : Paparan page
Frm105.L82_Text = 0 'Senarai belian tukaran barang oleh agen : Jumlah page
Frm105.L88_Text = 0 'Ambilan tunai dari kedai : Paparan page
Frm105.L89_Text = 0 'Ambilan tunai dari kedai : Jumlah page
Frm105.L95_Text = 0 'Perbelanjaan kedai : Paparan page
Frm105.L96_Text = 0 'Perbelanjaan kedai : Jumlah page
Frm105.L104_Text = 0 'Bayaran gaji : Paparan page
Frm105.L105_Text = 0 'Bayaran gaji : Jumlah page

Frm105.L11_Text = "0.00" 'Rekod Jualan : Jumlah
Frm105.L12_Text = "0.00" 'Rekod Jualan : Tunai
Frm105.L13_Text = "0.00" 'Rekod Jualan : Bank In
Frm105.L14_Text = "0.00" 'Rekod Jualan : Kad Kredit
Frm105.L15_Text = "0.00" 'Rekod Jualan : Kad Debit
Frm105.L16_Text = "0.00" 'Rekod Jualan : Simpanan Di Kedai
Frm105.L23_Text = "0.00" 'Rekod servis : Jumlah
Frm105.L24_Text = "0.00" 'Rekod servis : Tunai
Frm105.L25_Text = "0.00" 'Rekod servis : Bank In
Frm105.L26_Text = "0.00" 'Rekod servis : Kad Kredit
Frm105.L27_Text = "0.00" 'Rekod servis : Kad Debit
Frm105.L28_Text = "0.00" 'Rekod servis : Simpanan Di Kedai
Frm105.L35_Text = "0.00" 'Rekod ansuran : Jumlah
Frm105.L36_Text = "0.00" 'Rekod ansuran : Tunai
Frm105.L37_Text = "0.00" 'Rekod ansuran : Bank In
Frm105.L38_Text = "0.00" 'Rekod ansuran : Kad Kredit
Frm105.L39_Text = "0.00" 'Rekod ansuran : Kad Debit
Frm105.L40_Text = "0.00" 'Rekod ansuran : Simpanan Di Kedai
Frm105.L47_Text = "0.00" 'Rekod tempahan : Jumlah
Frm105.L48_Text = "0.00" 'Rekod tempahan : Tunai
Frm105.L49_Text = "0.00" 'Rekod tempahan : Bank In
Frm105.L50_Text = "0.00" 'Rekod tempahan : Kad Kredit
Frm105.L51_Text = "0.00" 'Rekod tempahan : Kad Debit
Frm105.L52_Text = "0.00" 'Rekod tempahan : Simpanan Di Kedai
Frm105.L59_Text = "0.00" 'Rekod kemasukkan duit ke kedai : Jumlah
Frm105.L66_Text = "0.00" 'Senarai simpanan duit di kedai oleh pelanggan : Jumlah
Frm105.L73_Text = "0.00" 'Senarai belian trade in : Jumlah
Frm105.L112_Text = "0.00" 'Trade in : Tunai
Frm105.L113_Text = "0.00" 'Trade in : Bank in
Frm105.L80_Text = "0.00" 'Senarai belian tukaran barang oleh agen : Jumlah
Frm105.L87_Text = "0.00" 'Ambilan tunai dari kedai : Jumlah
Frm105.L94_Text = "0.00" 'Perbelanjaan kedai : Jumlah
Frm105.L101_Text = "0.00" 'Bayaran gaji : Jumlah
Frm105.L102_Text = "0.00" 'Bayaran gaji : Tunai
Frm105.L103_Text = "0.00" 'Bayaran gaji : Bank In
End Sub
Sub Frm105_debit_setting()
'on error resume next
Frm105.Pic3.Left = 120
Frm105.Pic3.Top = 480
Frm105.Pic4.Left = 120
Frm105.Pic4.Top = 480
Frm105.Pic5.Left = 120
Frm105.Pic5.Top = 480
Frm105.Pic6.Left = 120
Frm105.Pic6.Top = 480
Frm105.Pic7.Left = 120
Frm105.Pic7.Top = 480
Frm105.Pic8.Left = 120
Frm105.Pic8.Top = 480

Frm105.Pic3.Visible = False
Frm105.Pic4.Visible = False
Frm105.Pic5.Visible = False
Frm105.Pic6.Visible = False
Frm105.Pic7.Visible = False
Frm105.Pic8.Visible = False
End Sub
Sub Frm105_kredit_setting()
'on error resume next
Frm105.Pic10.Left = 120
Frm105.Pic10.Top = 480
Frm105.Pic11.Left = 120
Frm105.Pic11.Top = 480
Frm105.Pic12.Left = 120
Frm105.Pic12.Top = 480
Frm105.Pic13.Left = 120
Frm105.Pic13.Top = 480
Frm105.Pic14.Left = 120
Frm105.Pic14.Top = 480

Frm105.Pic10.Visible = False
Frm105.Pic11.Visible = False
Frm105.Pic12.Visible = False
Frm105.Pic13.Visible = False
Frm105.Pic14.Visible = False
End Sub
Sub Frm105_senarai_jualan_header()
'on error resume next
'#### Header Report Senarai Jualan #### - Start
Frm105.MSFlexGrid1.Clear
Frm105.MSFlexGrid1.Rows = 1
Frm105.MSFlexGrid1.RowHeight(0) = 600
Frm105.MSFlexGrid1.FormatString = "No.|<No.|<No. ID|<Tarikh|<No. Invoice|<Jumlah (RM)|<Tunai (RM)|<Bank In (RM)|<Kad Kredit (RM)|<Kad Debit (RM)|<Simpanan Di Kedai (RM)|<Nama Pekerja"

Frm105.MSFlexGrid1.ColWidth(0) = 0 'No.
Frm105.MSFlexGrid1.ColWidth(1) = 600 'No.
Frm105.MSFlexGrid1.ColWidth(2) = 0 'No. ID
Frm105.MSFlexGrid1.ColWidth(3) = 1300 'Tarikh
Frm105.MSFlexGrid1.ColWidth(4) = 1500 'No. Invoice
Frm105.MSFlexGrid1.ColWidth(5) = 1500 'Jumlah (RM)
Frm105.MSFlexGrid1.ColWidth(6) = 1500 'Tunai (RM)
Frm105.MSFlexGrid1.ColWidth(7) = 1500 'Bank In (RM)
Frm105.MSFlexGrid1.ColWidth(8) = 1500 'Kad Kredit (RM)
Frm105.MSFlexGrid1.ColWidth(9) = 1500 'Kad Debit (RM)
Frm105.MSFlexGrid1.ColWidth(10) = 1500 'Simpanan Di Kedai (RM)
Frm105.MSFlexGrid1.ColWidth(11) = 1500 'Nama Pekerja
End Sub
Sub Frm105_senarai_jualan()
'On Error GoTo logging:
Dim TM As Date
Dim TA As Date
Dim Frm105_LM_TOTAL_PAGE As Double
Dim Frm105_LM_JUMLAH As Double
Dim Frm105_LM_TUNAI As Double
Dim Frm105_LM_BANK_IN As Double
Dim Frm105_LM_KAD_KREDIT As Double
Dim Frm105_LM_KAD_DEBIT As Double
Dim Frm105_LM_SIMPANAN As Double

Frm105_PAGE_SIZE = 40
Frm105_LM_TOTAL_PAGE = 0
x = 0
Frm105_LM_JUMLAH = 0
Frm105_LM_TUNAI = 0
Frm105_LM_BANK_IN = 0
Frm105_LM_KAD_KREDIT = 0
Frm105_LM_KAD_DEBIT = 0
Frm105_LM_SIMPANAN = 0

TM = Frm105.L5_Text 'Tarikh Mula
TA = Frm105.L6_Text 'Tarikh Akhir
If Frm105.L110_Text = "Semua" Then
    Frm105_LM_SEARCH_1 = Null
    Frm105_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm105_LM_SEARCH_1 = Frm105.L110_Text
    Frm105_LM_SEARCH_1_LOGIC = "="
End If

Frm105.L10_Text = "Senarai jualan dari " & TM & " hingga " & TA & " bagi nama pekerja [" & Frm105.L111_Text & "]." 'Report Header

LM_START_ROW = Frm105.L19_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm105_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm105.L20_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm105_PAGE_SIZE
        End If
    End If
End If

Frm105_LM_PAGE_FOUND = 0

LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where menu = 0 AND status = 1 AND no_pekerja " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm105_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm105_LM_PAGE_FOUND = 0 Then
        If Frm105.L20_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm105.L17_Text = Frm105.L17_Text + 1 'Paparan Page ke-xxx
                Frm105_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm105.L17_Text) Then
                    If Frm105.L17_Text <> 1 Then
                        Frm105.L17_Text = Frm105.L17_Text - 1 'Paparan Page ke-xxx
                        Frm105_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm105.L17_Text - 1) * Frm105_PAGE_SIZE) + x
    Frm105.MSFlexGrid1.Rows = x + 1
    Frm105.MSFlexGrid1.TextMatrix(x, 0) = x 'No.
    Frm105.MSFlexGrid1.TextMatrix(x, 1) = Y 'No.
    Frm105.MSFlexGrid1.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!tarikh) Then Frm105.MSFlexGrid1.TextMatrix(x, 3) = rs!tarikh 'Tarikh Jualan
    If Not IsNull(rs!no_resit) Then Frm105.MSFlexGrid1.TextMatrix(x, 4) = rs!no_resit 'No. Invoice
    'If Not IsNull(rs!harga_jualan) Then 'Jumlah Bayaran / Jumlah Invoice (RM)
    '    Frm105.MSFlexGrid1.TextMatrix(x, 5) = Format(rs!harga_jualan, "#,##0.00") 'Jumlah Bayaran / Jumlah Invoice (RM)
    'Else
    '    Frm105.MSFlexGrid1.TextMatrix(x, 5) = "0.00" 'Jumlah Bayaran / Jumlah Invoice (RM)
    'End If
    If rs!flag_bayaran = 0 Then
        If Not IsNull(rs!jumlah_perlu_bayar) Then 'Jumlah Bayaran / Jumlah Invoice (RM)'harga_lepas_diskaun
            Frm105.MSFlexGrid1.TextMatrix(x, 5) = Format(rs!jumlah_perlu_bayar, "#,##0.00") 'Jumlah Bayaran / Jumlah Invoice (RM)
        Else
            Frm105.MSFlexGrid1.TextMatrix(x, 5) = "0.00" 'Jumlah Bayaran / Jumlah Invoice (RM)
        End If
    Else
        If Not IsNull(rs!harga_lepas_diskaun) Then 'Jumlah Bayaran / Jumlah Invoice (RM)'harga_lepas_diskaun
            Frm105.MSFlexGrid1.TextMatrix(x, 5) = Format(rs!harga_lepas_diskaun, "#,##0.00") 'Jumlah Bayaran / Jumlah Invoice (RM)
        Else
            Frm105.MSFlexGrid1.TextMatrix(x, 5) = "0.00" 'Jumlah Bayaran / Jumlah Invoice (RM)
        End If
    End If
    
    If Not IsNull(rs!tunai) Then 'Jumlah Kutipan TUNAI (RM)
        Frm105.MSFlexGrid1.TextMatrix(x, 6) = Format(rs!tunai, "#,##0.00") 'Jumlah Kutipan TUNAI (RM)
    Else
        Frm105.MSFlexGrid1.TextMatrix(x, 6) = "0.00" 'Jumlah Kutipan TUNAI (RM)
    End If
    If Not IsNull(rs!bank_in) Then 'Jumlah Kutipan BANK IN (RM)
        Frm105.MSFlexGrid1.TextMatrix(x, 7) = Format(rs!bank_in, "#,##0.00") 'Jumlah Kutipan BANK IN (RM)
    Else
        Frm105.MSFlexGrid1.TextMatrix(x, 7) = "0.00" 'Jumlah Kutipan BANK IN (RM)
    End If
    If Not IsNull(rs!kad_kredit) Then 'Jumlah Kutipan KAD KREDIT (RM)
        Frm105.MSFlexGrid1.TextMatrix(x, 8) = Format(rs!kad_kredit, "#,##0.00") 'Jumlah Kutipan KAD KREDIT (RM)
    Else
        Frm105.MSFlexGrid1.TextMatrix(x, 8) = "0.00" 'Jumlah Kutipan KAD KREDIT (RM)
    End If
    If Not IsNull(rs!kad_debit) Then 'Jumlah Kutipan KAD DEBIT (RM)
        Frm105.MSFlexGrid1.TextMatrix(x, 9) = Format(rs!kad_debit, "#,##0.00") 'Jumlah Kutipan KAD DEBIT (RM)
    Else
        Frm105.MSFlexGrid1.TextMatrix(x, 9) = "0.00" 'Jumlah Kutipan KAD DEBIT (RM)
    End If
    If Not IsNull(rs!duit_simpanan_kedai) Then 'Jumlah Kutipan Simpanan Di Kedai (RM)
        Frm105.MSFlexGrid1.TextMatrix(x, 10) = Format(rs!duit_simpanan_kedai, "#,##0.00") 'Jumlah Kutipan Simpanan Di Kedai (RM)
    Else
        Frm105.MSFlexGrid1.TextMatrix(x, 10) = "0.00" 'Jumlah Kutipan Simpanan Di Kedai (RM)
    End If
    If Not IsNull(rs!nama_pekerja) Then Frm105.MSFlexGrid1.TextMatrix(x, 11) = rs!nama_pekerja
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
LM_CONN = 2
re_conn_2:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 22_jualan where menu = 0 AND no_pekerja " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm105_LM_TOTAL_PAGE = Format(rs(0) / Frm105_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm105_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm105_LM_PAGE = Split(Frm105_LM_TOTAL_PAGE, ".")(0)
        Frm105_LM_PAGE_LEBIHAN = Split(Frm105_LM_TOTAL_PAGE, ".")(1)
        
        If Frm105_LM_PAGE_LEBIHAN <> "00" Then
            Frm105.L18_Text = Frm105_LM_PAGE + 1
        Else
            Frm105.L18_Text = Frm105_LM_PAGE
        End If
        
    Else
    
        Frm105.L18_Text = Frm105_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm105.L18_Text = 0
    End If
Else
    Frm105.L18_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm105.L18_Text = vbNullString Then
    Frm105.L18_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah ### - Start
LM_CONN = 3
re_conn_3:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_jualan),SUM(Tunai),SUM(bank_in),SUM(kad_kredit),SUM(kad_debit),SUM(duit_simpanan_kedai) from 22_jualan where menu = 0 AND no_pekerja " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND flag_bayaran = 0 AND status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm105_LM_JUMLAH = rs(0) 'Jumlah (RM)
If Not IsNull(rs(1)) Then Frm105_LM_TUNAI = rs(1) 'Jumlah Tunai (RM)
If Not IsNull(rs(2)) Then Frm105_LM_BANK_IN = rs(2) 'Jumlah Bank In (RM)
If Not IsNull(rs(3)) Then Frm105_LM_KAD_KREDIT = rs(3) 'Jumlah Kad Kredit (RM)
If Not IsNull(rs(4)) Then Frm105_LM_KAD_DEBIT = rs(4) 'Jumlah Kad Debit (RM)
If Not IsNull(rs(5)) Then Frm105_LM_SIMPANAN = rs(5) 'Jumlah Simpanan Di Kedai (RM)

rs.Close
Set rs = Nothing
'### Jumlah ### - End

Frm105.L11_Text = Format(Frm105_LM_JUMLAH, "#,##0.00") 'Jumlah (RM)
Frm105.L12_Text = Format(Frm105_LM_TUNAI, "#,##0.00") 'Jumlah Tunai (RM)
Frm105.L13_Text = Format(Frm105_LM_BANK_IN, "#,##0.00") 'Jumlah Bank In (RM)
Frm105.L14_Text = Format(Frm105_LM_KAD_KREDIT, "#,##0.00") 'Jumlah Kad Kredit (RM)
Frm105.L15_Text = Format(Frm105_LM_KAD_DEBIT, "#,##0.00") 'Jumlah Kad Debit (RM)
Frm105.L16_Text = Format(Frm105_LM_SIMPANAN, "#,##0.00") 'Jumlah Simpanan Di Kedai (RM)

If x <> 0 Then
    Frm105.L19_Text = LM_START_ROW 'Titik Pencarian Data
    Frm105.Pic3.Visible = True
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    If Frm105_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm105.L20_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm105.L20_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_report_kewangan : Frm105_senarai_jualan" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If LM_ERR_NO = "3704" Or LM_ERR_NO = "-2147467259" Or LM_ERR_NO = "-2147217887" Then
    
    Call Main

    If LM_CONN = 1 Then
        Resume re_conn_1:
    ElseIf LM_CONN = 2 Then
        Resume re_conn_2:
    ElseIf LM_CONN = 3 Then
        Resume re_conn_3:
    End If
Else
    Resume Next
End If
End Sub
Sub Frm105_senarai_servis_header()
'On Error Resume Next
'#### Header Report Servis #### - Start
Frm105.MSFlexGrid2.Clear
Frm105.MSFlexGrid2.Rows = 1
Frm105.MSFlexGrid2.RowHeight(0) = 600
Frm105.MSFlexGrid2.FormatString = "No.|<No.|<No. ID|<Tarikh|<No. Invoice|<Jumlah (RM)|<Tunai (RM)|<Bank In (RM)|<Kad Kredit (RM)|<Kad Debit (RM)|<Simpanan Di Kedai (RM)|<Nama Pekerja"

Frm105.MSFlexGrid2.ColWidth(0) = 0 'No.
Frm105.MSFlexGrid2.ColWidth(1) = 600 'No.
Frm105.MSFlexGrid2.ColWidth(2) = 0 'No. ID
Frm105.MSFlexGrid2.ColWidth(3) = 1300 'Tarikh
Frm105.MSFlexGrid2.ColWidth(4) = 1500 'No. Invoice
Frm105.MSFlexGrid2.ColWidth(5) = 1500 'Jumlah (RM)
Frm105.MSFlexGrid2.ColWidth(6) = 1500 'Tunai (RM)
Frm105.MSFlexGrid2.ColWidth(7) = 1500 'Bank In (RM)
Frm105.MSFlexGrid2.ColWidth(8) = 1500 'Kad Kredit (RM)
Frm105.MSFlexGrid2.ColWidth(9) = 1500 'Kad Debit (RM)
Frm105.MSFlexGrid2.ColWidth(10) = 1500 'Simpanan Di Kedai (RM)
Frm105.MSFlexGrid2.ColWidth(11) = 1500 'Nama Pekerja
'#### Header Report Servis #### - End
End Sub
Sub Frm105_senarai_servis()
'On Error GoTo logging:
Dim TM As Date
Dim TA As Date
Dim Frm105_LM_TOTAL_PAGE As Double
Dim Frm105_LM_JUMLAH As Double
Dim Frm105_LM_TUNAI As Double
Dim Frm105_LM_BANK_IN As Double
Dim Frm105_LM_KAD_KREDIT As Double
Dim Frm105_LM_KAD_DEBIT As Double
Dim Frm105_LM_SIMPANAN As Double

Frm105_PAGE_SIZE = 40
Frm105_LM_TOTAL_PAGE = 0
x = 0
Frm105_LM_JUMLAH = 0
Frm105_LM_TUNAI = 0
Frm105_LM_BANK_IN = 0
Frm105_LM_KAD_KREDIT = 0
Frm105_LM_KAD_DEBIT = 0
Frm105_LM_SIMPANAN = 0

TM = Frm105.L5_Text 'Tarikh Mula
TA = Frm105.L6_Text 'Tarikh Akhir
If Frm105.L110_Text = "Semua" Then
    Frm105_LM_SEARCH_1 = Null
    Frm105_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm105_LM_SEARCH_1 = Frm105.L110_Text
    Frm105_LM_SEARCH_1_LOGIC = "="
End If

Frm105.L22_Text = "Senarai servis dari " & TM & " hingga " & TA & " bagi nama pekerja [" & Frm105.L111_Text & "]." 'Report Header

LM_START_ROW = Frm105.L31_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm105_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm105.L32_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm105_PAGE_SIZE
        End If
    End If
End If

Frm105_LM_PAGE_FOUND = 0

LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where menu = 1 AND no_pekerja " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm105_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm105_LM_PAGE_FOUND = 0 Then
        If Frm105.L32_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm105.L29_Text = Frm105.L29_Text + 1 'Paparan Page ke-xxx
                Frm105_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm105.L29_Text) Then
                    If Frm105.L29_Text <> 1 Then
                        Frm105.L29_Text = Frm105.L29_Text - 1 'Paparan Page ke-xxx
                        Frm105_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm105.L29_Text - 1) * Frm105_PAGE_SIZE) + x
    Frm105.MSFlexGrid2.Rows = x + 1
    Frm105.MSFlexGrid2.TextMatrix(x, 0) = x 'No.
    Frm105.MSFlexGrid2.TextMatrix(x, 1) = Y 'No.
    Frm105.MSFlexGrid2.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!tarikh) Then Frm105.MSFlexGrid2.TextMatrix(x, 3) = rs!tarikh 'Tarikh Jualan
    If Not IsNull(rs!no_resit) Then Frm105.MSFlexGrid2.TextMatrix(x, 4) = rs!no_resit 'No. Invoice
    'If Not IsNull(rs!harga_jualan) Then 'Jumlah Bayaran / Jumlah Invoice (RM)
    '    Frm105.MSFlexGrid2.TextMatrix(x, 5) = Format(rs!harga_jualan, "#,##0.00") 'Jumlah Bayaran / Jumlah Invoice (RM)
    'Else
    '    Frm105.MSFlexGrid2.TextMatrix(x, 5) = "0.00" 'Jumlah Bayaran / Jumlah Invoice (RM)
    'End If
    If rs!flag_bayaran = 0 Then
        If Not IsNull(rs!jumlah_perlu_bayar) Then 'Jumlah Bayaran / Jumlah Invoice (RM)'harga_lepas_diskaun
            Frm105.MSFlexGrid2.TextMatrix(x, 5) = Format(rs!jumlah_perlu_bayar, "#,##0.00") 'Jumlah Bayaran / Jumlah Invoice (RM)
        Else
            Frm105.MSFlexGrid2.TextMatrix(x, 5) = "0.00" 'Jumlah Bayaran / Jumlah Invoice (RM)
        End If
    Else
        If Not IsNull(rs!harga_lepas_diskaun) Then 'Jumlah Bayaran / Jumlah Invoice (RM)'harga_lepas_diskaun
            Frm105.MSFlexGrid2.TextMatrix(x, 5) = Format(rs!harga_lepas_diskaun, "#,##0.00") 'Jumlah Bayaran / Jumlah Invoice (RM)
        Else
            Frm105.MSFlexGrid2.TextMatrix(x, 5) = "0.00" 'Jumlah Bayaran / Jumlah Invoice (RM)
        End If
    End If
    
    If Not IsNull(rs!tunai) Then 'Jumlah Kutipan TUNAI (RM)
        Frm105.MSFlexGrid2.TextMatrix(x, 6) = Format(rs!tunai, "#,##0.00") 'Jumlah Kutipan TUNAI (RM)
    Else
        Frm105.MSFlexGrid2.TextMatrix(x, 6) = "0.00" 'Jumlah Kutipan TUNAI (RM)
    End If
    If Not IsNull(rs!bank_in) Then 'Jumlah Kutipan BANK IN (RM)
        Frm105.MSFlexGrid2.TextMatrix(x, 7) = Format(rs!bank_in, "#,##0.00") 'Jumlah Kutipan BANK IN (RM)
    Else
        Frm105.MSFlexGrid2.TextMatrix(x, 7) = "0.00" 'Jumlah Kutipan BANK IN (RM)
    End If
    If Not IsNull(rs!kad_kredit) Then 'Jumlah Kutipan KAD KREDIT (RM)
        Frm105.MSFlexGrid2.TextMatrix(x, 8) = Format(rs!kad_kredit, "#,##0.00") 'Jumlah Kutipan KAD KREDIT (RM)
    Else
        Frm105.MSFlexGrid2.TextMatrix(x, 8) = "0.00" 'Jumlah Kutipan KAD KREDIT (RM)
    End If
    If Not IsNull(rs!kad_debit) Then 'Jumlah Kutipan KAD DEBIT (RM)
        Frm105.MSFlexGrid2.TextMatrix(x, 9) = Format(rs!kad_debit, "#,##0.00") 'Jumlah Kutipan KAD DEBIT (RM)
    Else
        Frm105.MSFlexGrid2.TextMatrix(x, 9) = "0.00" 'Jumlah Kutipan KAD DEBIT (RM)
    End If
    If Not IsNull(rs!duit_simpanan_kedai) Then 'Jumlah Kutipan Simpanan Di Kedai (RM)
        Frm105.MSFlexGrid2.TextMatrix(x, 10) = Format(rs!duit_simpanan_kedai, "#,##0.00") 'Jumlah Kutipan Simpanan Di Kedai (RM)
    Else
        Frm105.MSFlexGrid2.TextMatrix(x, 10) = "0.00" 'Jumlah Kutipan Simpanan Di Kedai (RM)
    End If
    If Not IsNull(rs!nama_pekerja) Then Frm105.MSFlexGrid2.TextMatrix(x, 11) = rs!nama_pekerja
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
LM_CONN = 2
re_conn_2:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID),SUM(harga_jualan),SUM(Tunai),SUM(bank_in),SUM(kad_kredit),SUM(kad_debit),SUM(duit_simpanan_kedai) from 22_jualan where menu = 1 AND no_pekerja " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm105_LM_TOTAL_PAGE = Format(rs(0) / Frm105_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm105_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm105_LM_PAGE = Split(Frm105_LM_TOTAL_PAGE, ".")(0)
        Frm105_LM_PAGE_LEBIHAN = Split(Frm105_LM_TOTAL_PAGE, ".")(1)
        
        If Frm105_LM_PAGE_LEBIHAN <> "00" Then
            Frm105.L30_Text = Frm105_LM_PAGE + 1
        Else
            Frm105.L30_Text = Frm105_LM_PAGE
        End If
        
    Else
    
        Frm105.L30_Text = Frm105_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm105.L30_Text = 0
    End If
Else
    Frm105.L30_Text = 0
End If

If Not IsNull(rs(1)) Then Frm105_LM_JUMLAH = rs(1) 'Jumlah (RM)
If Not IsNull(rs(2)) Then Frm105_LM_TUNAI = rs(2) 'Jumlah Tunai (RM)
If Not IsNull(rs(3)) Then Frm105_LM_BANK_IN = rs(3) 'Jumlah Bank In (RM)
If Not IsNull(rs(4)) Then Frm105_LM_KAD_KREDIT = rs(4) 'Jumlah Kad Kredit (RM)
If Not IsNull(rs(5)) Then Frm105_LM_KAD_DEBIT = rs(5) 'Jumlah Kad Debit (RM)
If Not IsNull(rs(6)) Then Frm105_LM_SIMPANAN = rs(6) 'Jumlah Simpanan Di Kedai (RM)

rs.Close
Set rs = Nothing

If Frm105.L30_Text = vbNullString Then
    Frm105.L30_Text = 0
End If
'### Jumlah Data ### - End

Frm105.L23_Text = Format(Frm105_LM_JUMLAH, "#,##0.00") 'Jumlah (RM)
Frm105.L24_Text = Format(Frm105_LM_TUNAI, "#,##0.00") 'Jumlah Tunai (RM)
Frm105.L25_Text = Format(Frm105_LM_BANK_IN, "#,##0.00") 'Jumlah Bank In (RM)
Frm105.L26_Text = Format(Frm105_LM_KAD_KREDIT, "#,##0.00") 'Jumlah Kad Kredit (RM)
Frm105.L27_Text = Format(Frm105_LM_KAD_DEBIT, "#,##0.00") 'Jumlah Kad Debit (RM)
Frm105.L28_Text = Format(Frm105_LM_SIMPANAN, "#,##0.00") 'Jumlah Simpanan Di Kedai (RM)

If x <> 0 Then
    Frm105.L31_Text = LM_START_ROW 'Titik Pencarian Data
    Frm105.Pic4.Visible = True
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    If Frm105_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm105.L32_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm105.L32_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_report_kewangan : Frm105_senarai_servis" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If LM_ERR_NO = "3704" Or LM_ERR_NO = "-2147467259" Or LM_ERR_NO = "-2147217887" Then
    
    Call Main

    If LM_CONN = 1 Then
        Resume re_conn_1:
    ElseIf LM_CONN = 2 Then
        Resume re_conn_2:
    End If
Else
    Resume Next
End If
End Sub
Sub Frm105_senarai_ansuran_header()
'On Error Resume Next
'#### Header Report Servis #### - Start
Frm105.MSFlexGrid3.Clear
Frm105.MSFlexGrid3.Rows = 1
Frm105.MSFlexGrid3.RowHeight(0) = 600
Frm105.MSFlexGrid3.FormatString = "No.|<No.|<No. ID|<Tarikh|<No. Invoice|<Jumlah (RM)|<Tunai (RM)|<Bank In (RM)|<Kad Kredit (RM)|<Kad Debit (RM)|<Simpanan Di Kedai (RM)"

Frm105.MSFlexGrid3.ColWidth(0) = 0 'No.
Frm105.MSFlexGrid3.ColWidth(1) = 600 'No.
Frm105.MSFlexGrid3.ColWidth(2) = 0 'No. ID
Frm105.MSFlexGrid3.ColWidth(3) = 1500 'Tarikh
Frm105.MSFlexGrid3.ColWidth(4) = 1500 'No. Invoice
Frm105.MSFlexGrid3.ColWidth(5) = 1500 'Jumlah (RM)
Frm105.MSFlexGrid3.ColWidth(6) = 1500 'Tunai (RM)
Frm105.MSFlexGrid3.ColWidth(7) = 1500 'Bank In (RM)
Frm105.MSFlexGrid3.ColWidth(8) = 1500 'Kad Kredit (RM)
Frm105.MSFlexGrid3.ColWidth(9) = 1500 'Kad Debit (RM)
Frm105.MSFlexGrid3.ColWidth(10) = 1500 'Simpanan Di Kedai (RM)
'#### Header Report Servis #### - End
End Sub
Sub Frm105_senarai_ansuran()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim Frm105_LM_TOTAL_PAGE As Double
Dim Frm105_LM_JUMLAH As Double
Dim Frm105_LM_TUNAI As Double
Dim Frm105_LM_BANK_IN As Double
Dim Frm105_LM_KAD_KREDIT As Double
Dim Frm105_LM_KAD_DEBIT As Double
Dim Frm105_LM_SIMPANAN As Double

Frm105_PAGE_SIZE = 34
Frm105_LM_TOTAL_PAGE = 0
x = 0
Frm105_LM_JUMLAH = 0
Frm105_LM_TUNAI = 0
Frm105_LM_BANK_IN = 0
Frm105_LM_KAD_KREDIT = 0
Frm105_LM_KAD_DEBIT = 0
Frm105_LM_SIMPANAN = 0

TM = Frm105.L5_Text 'Tarikh Mula
TA = Frm105.L6_Text 'Tarikh Akhir

Frm105.L34_Text = "Senarai bayaran ansuran dari " & TM & " hingga " & TA & "." 'Report Header

LM_START_ROW = Frm105.L43_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm105_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm105.L44_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm105_PAGE_SIZE
        End If
    End If
End If

Frm105_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 29_akaun_ansuran where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm105_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm105_LM_PAGE_FOUND = 0 Then
        If Frm105.L44_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm105.L41_Text = Frm105.L41_Text + 1 'Paparan Page ke-xxx
                Frm105_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm105.L41_Text) Then
                    If Frm105.L41_Text <> 1 Then
                        Frm105.L41_Text = Frm105.L41_Text - 1 'Paparan Page ke-xxx
                        Frm105_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm105.L41_Text - 1) * Frm105_PAGE_SIZE) + x
    Frm105.MSFlexGrid3.Rows = x + 1
    Frm105.MSFlexGrid3.TextMatrix(x, 0) = x 'No.
    Frm105.MSFlexGrid3.TextMatrix(x, 1) = Y 'No.
    Frm105.MSFlexGrid3.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!tarikh) Then Frm105.MSFlexGrid3.TextMatrix(x, 3) = rs!tarikh 'Tarikh Jualan
    If Not IsNull(rs!no_resit) Then Frm105.MSFlexGrid3.TextMatrix(x, 4) = rs!no_resit 'No. Invoice
    If Not IsNull(rs!jumlah) Then 'Jumlah Bayaran / Jumlah Invoice (RM)
        Frm105.MSFlexGrid3.TextMatrix(x, 5) = Format(rs!jumlah, "#,##0.00") 'Jumlah Bayaran / Jumlah Invoice (RM)
    Else
        Frm105.MSFlexGrid3.TextMatrix(x, 5) = "0.00" 'Jumlah Bayaran / Jumlah Invoice (RM)
    End If
    If Not IsNull(rs!tunai) Then 'Jumlah Kutipan TUNAI (RM)
        Frm105.MSFlexGrid3.TextMatrix(x, 6) = Format(rs!tunai, "#,##0.00") 'Jumlah Kutipan TUNAI (RM)
    Else
        Frm105.MSFlexGrid3.TextMatrix(x, 6) = "0.00" 'Jumlah Kutipan TUNAI (RM)
    End If
    If Not IsNull(rs!bank_in) Then 'Jumlah Kutipan BANK IN (RM)
        Frm105.MSFlexGrid3.TextMatrix(x, 7) = Format(rs!bank_in, "#,##0.00") 'Jumlah Kutipan BANK IN (RM)
    Else
        Frm105.MSFlexGrid3.TextMatrix(x, 7) = "0.00" 'Jumlah Kutipan BANK IN (RM)
    End If
    If Not IsNull(rs!kad_kredit) Then 'Jumlah Kutipan KAD KREDIT (RM)
        Frm105.MSFlexGrid3.TextMatrix(x, 8) = Format(rs!kad_kredit, "#,##0.00") 'Jumlah Kutipan KAD KREDIT (RM)
    Else
        Frm105.MSFlexGrid3.TextMatrix(x, 8) = "0.00" 'Jumlah Kutipan KAD KREDIT (RM)
    End If
    If Not IsNull(rs!kad_debit) Then 'Jumlah Kutipan KAD DEBIT (RM)
        Frm105.MSFlexGrid3.TextMatrix(x, 9) = Format(rs!kad_debit, "#,##0.00") 'Jumlah Kutipan KAD DEBIT (RM)
    Else
        Frm105.MSFlexGrid3.TextMatrix(x, 9) = "0.00" 'Jumlah Kutipan KAD DEBIT (RM)
    End If
    If Not IsNull(rs!duit_simpanan_kedai) Then 'Jumlah Kutipan Simpanan Di Kedai (RM)
        Frm105.MSFlexGrid3.TextMatrix(x, 10) = Format(rs!duit_simpanan_kedai, "#,##0.00") 'Jumlah Kutipan Simpanan Di Kedai (RM)
    Else
        Frm105.MSFlexGrid3.TextMatrix(x, 10) = "0.00" 'Jumlah Kutipan Simpanan Di Kedai (RM)
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 29_akaun_ansuran where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm105_LM_TOTAL_PAGE = Format(rs(0) / Frm105_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm105_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm105_LM_PAGE = Split(Frm105_LM_TOTAL_PAGE, ".")(0)
        Frm105_LM_PAGE_LEBIHAN = Split(Frm105_LM_TOTAL_PAGE, ".")(1)
        
        If Frm105_LM_PAGE_LEBIHAN <> "00" Then
            Frm105.L42_Text = Frm105_LM_PAGE + 1
        Else
            Frm105.L42_Text = Frm105_LM_PAGE
        End If
        
    Else
    
        Frm105.L42_Text = Frm105_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm105.L42_Text = 0
    End If
Else
    Frm105.L42_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm105.L42_Text = vbNullString Then
    Frm105.L42_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah) from 29_akaun_ansuran where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm105_LM_JUMLAH = rs(0) 'Jumlah (RM)

rs.Close
Set rs = Nothing
'### Jumlah ### - End

'### Tunai Terkumpul ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(Tunai) from 29_akaun_ansuran where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm105_LM_TUNAI = rs(0) 'Jumlah Tunai (RM)

rs.Close
Set rs = Nothing
'### Tunai Terkumpul ### - End

'### Bank In Terkumpul ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(bank_in) from 29_akaun_ansuran where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm105_LM_BANK_IN = rs(0) 'Jumlah Bank In (RM)

rs.Close
Set rs = Nothing
'### Bank In Terkumpul ### - End

'### Kad Kredit Terkumpul ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(kad_kredit) from 29_akaun_ansuran where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm105_LM_KAD_KREDIT = rs(0) 'Jumlah Kad Kredit (RM)

rs.Close
Set rs = Nothing
'### Kad Kredit Terkumpul ### - End

'### Kad Debit Terkumpul ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(kad_debit) from 29_akaun_ansuran where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm105_LM_KAD_DEBIT = rs(0) 'Jumlah Kad Debit (RM)

rs.Close
Set rs = Nothing
'### Kad Debit Terkumpul ### - End

'### Simpanan Di Kedai Terkumpul ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(duit_simpanan_kedai) from 29_akaun_ansuran where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm105_LM_SIMPANAN = rs(0) 'Jumlah Simpanan Di Kedai (RM)

rs.Close
Set rs = Nothing
'### Simpanan Di Kedai Terkumpul ### - End

Frm105.L35_Text = Format(Frm105_LM_JUMLAH, "#,##0.00") 'Jumlah (RM)
Frm105.L36_Text = Format(Frm105_LM_TUNAI, "#,##0.00") 'Jumlah Tunai (RM)
Frm105.L37_Text = Format(Frm105_LM_BANK_IN, "#,##0.00") 'Jumlah Bank In (RM)
Frm105.L38_Text = Format(Frm105_LM_KAD_KREDIT, "#,##0.00") 'Jumlah Kad Kredit (RM)
Frm105.L39_Text = Format(Frm105_LM_KAD_DEBIT, "#,##0.00") 'Jumlah Kad Debit (RM)
Frm105.L40_Text = Format(Frm105_LM_SIMPANAN, "#,##0.00") 'Jumlah Simpanan Di Kedai (RM)

If x <> 0 Then
    Frm105.L43_Text = LM_START_ROW 'Titik Pencarian Data
    Frm105.Pic5.Visible = True
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    If Frm105_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm105.L44_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm105.L44_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub Frm105_senarai_tempahan_header()
'On Error Resume Next
'#### Header Report Servis #### - Start
Frm105.MSFlexGrid4.Clear
Frm105.MSFlexGrid4.Rows = 1
Frm105.MSFlexGrid4.RowHeight(0) = 600
Frm105.MSFlexGrid4.FormatString = "No.|<No.|<No. ID|<Tarikh|<No. Invoice|<Jumlah (RM)|<Tunai (RM)|<Bank In (RM)|<Kad Kredit (RM)|<Kad Debit (RM)|<Simpanan Di Kedai (RM)"

Frm105.MSFlexGrid4.ColWidth(0) = 0 'No.
Frm105.MSFlexGrid4.ColWidth(1) = 600 'No.
Frm105.MSFlexGrid4.ColWidth(2) = 0 'No. ID
Frm105.MSFlexGrid4.ColWidth(3) = 1500 'Tarikh
Frm105.MSFlexGrid4.ColWidth(4) = 1500 'No. Invoice
Frm105.MSFlexGrid4.ColWidth(5) = 1500 'Jumlah (RM)
Frm105.MSFlexGrid4.ColWidth(6) = 1500 'Tunai (RM)
Frm105.MSFlexGrid4.ColWidth(7) = 1500 'Bank In (RM)
Frm105.MSFlexGrid4.ColWidth(8) = 1500 'Kad Kredit (RM)
Frm105.MSFlexGrid4.ColWidth(9) = 1500 'Kad Debit (RM)
Frm105.MSFlexGrid4.ColWidth(10) = 1500 'Simpanan Di Kedai (RM)
'#### Header Report Servis #### - End
End Sub
Sub Frm105_senarai_tempahan()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim Frm105_LM_TOTAL_PAGE As Double
Dim Frm105_LM_JUMLAH As Double
Dim Frm105_LM_TUNAI As Double
Dim Frm105_LM_BANK_IN As Double
Dim Frm105_LM_KAD_KREDIT As Double
Dim Frm105_LM_KAD_DEBIT As Double
Dim Frm105_LM_SIMPANAN As Double

Frm105_PAGE_SIZE = 34
Frm105_LM_TOTAL_PAGE = 0
x = 0
Frm105_LM_JUMLAH = 0
Frm105_LM_TUNAI = 0
Frm105_LM_BANK_IN = 0
Frm105_LM_KAD_KREDIT = 0
Frm105_LM_KAD_DEBIT = 0
Frm105_LM_SIMPANAN = 0

TM = Frm105.L5_Text 'Tarikh Mula
TA = Frm105.L6_Text 'Tarikh Akhir

Frm105.L46_Text = "Senarai bayaran tempahan dari " & TM & " hingga " & TA & "." 'Report Header

LM_START_ROW = Frm105.L55_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm105_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm105.L56_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm105_PAGE_SIZE
        End If
    End If
End If

Frm105_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from 22_jualan where flag_bayaran='" & "0" & "' AND status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm105_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
rs.Open "select * from 22_jualan where (menu = 2 OR menu = 3) AND status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm105_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm105_LM_PAGE_FOUND = 0 Then
        If Frm105.L56_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm105.L53_Text = Frm105.L53_Text + 1 'Paparan Page ke-xxx
                Frm105_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm105.L53_Text) Then
                    If Frm105.L53_Text <> 1 Then
                        Frm105.L53_Text = Frm105.L53_Text - 1 'Paparan Page ke-xxx
                        Frm105_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm105.L17_Text - 1) * Frm105_PAGE_SIZE) + x
    Frm105.MSFlexGrid4.Rows = x + 1
    Frm105.MSFlexGrid4.TextMatrix(x, 0) = x 'No.
    Frm105.MSFlexGrid4.TextMatrix(x, 1) = Y 'No.
    Frm105.MSFlexGrid4.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!tarikh) Then Frm105.MSFlexGrid4.TextMatrix(x, 3) = rs!tarikh 'Tarikh Jualan
    If Not IsNull(rs!no_resit) Then Frm105.MSFlexGrid4.TextMatrix(x, 4) = rs!no_resit 'No. Invoice
    'If Not IsNull(rs!harga_jualan) Then 'Jumlah Bayaran / Jumlah Invoice (RM)
    '    Frm105.MSFlexGrid4.TextMatrix(x, 5) = Format(rs!harga_jualan, "#,##0.00") 'Jumlah Bayaran / Jumlah Invoice (RM)
    'Else
    '    Frm105.MSFlexGrid4.TextMatrix(x, 5) = "0.00" 'Jumlah Bayaran / Jumlah Invoice (RM)
    'End If
    If rs!flag_bayaran = 0 Then
        If Not IsNull(rs!jumlah_perlu_bayar) Then 'Jumlah Bayaran / Jumlah Invoice (RM)'harga_lepas_diskaun
            Frm105.MSFlexGrid4.TextMatrix(x, 5) = Format(rs!jumlah_perlu_bayar, "#,##0.00") 'Jumlah Bayaran / Jumlah Invoice (RM)
        Else
            Frm105.MSFlexGrid4.TextMatrix(x, 5) = "0.00" 'Jumlah Bayaran / Jumlah Invoice (RM)
        End If
    Else
        If Not IsNull(rs!harga_lepas_diskaun) Then 'Jumlah Bayaran / Jumlah Invoice (RM)'harga_lepas_diskaun
            Frm105.MSFlexGrid4.TextMatrix(x, 5) = Format(rs!harga_lepas_diskaun, "#,##0.00") 'Jumlah Bayaran / Jumlah Invoice (RM)
        Else
            Frm105.MSFlexGrid4.TextMatrix(x, 5) = "0.00" 'Jumlah Bayaran / Jumlah Invoice (RM)
        End If
    End If
    
    If Not IsNull(rs!tunai) Then 'Jumlah Kutipan TUNAI (RM)
        Frm105.MSFlexGrid4.TextMatrix(x, 6) = Format(rs!tunai, "#,##0.00") 'Jumlah Kutipan TUNAI (RM)
    Else
        Frm105.MSFlexGrid4.TextMatrix(x, 6) = "0.00" 'Jumlah Kutipan TUNAI (RM)
    End If
    If Not IsNull(rs!bank_in) Then 'Jumlah Kutipan BANK IN (RM)
        Frm105.MSFlexGrid4.TextMatrix(x, 7) = Format(rs!bank_in, "#,##0.00") 'Jumlah Kutipan BANK IN (RM)
    Else
        Frm105.MSFlexGrid4.TextMatrix(x, 7) = "0.00" 'Jumlah Kutipan BANK IN (RM)
    End If
    If Not IsNull(rs!kad_kredit) Then 'Jumlah Kutipan KAD KREDIT (RM)
        Frm105.MSFlexGrid4.TextMatrix(x, 8) = Format(rs!kad_kredit, "#,##0.00") 'Jumlah Kutipan KAD KREDIT (RM)
    Else
        Frm105.MSFlexGrid4.TextMatrix(x, 8) = "0.00" 'Jumlah Kutipan KAD KREDIT (RM)
    End If
    If Not IsNull(rs!kad_debit) Then 'Jumlah Kutipan KAD DEBIT (RM)
        Frm105.MSFlexGrid4.TextMatrix(x, 9) = Format(rs!kad_debit, "#,##0.00") 'Jumlah Kutipan KAD DEBIT (RM)
    Else
        Frm105.MSFlexGrid4.TextMatrix(x, 9) = "0.00" 'Jumlah Kutipan KAD DEBIT (RM)
    End If
    If Not IsNull(rs!duit_simpanan_kedai) Then 'Jumlah Kutipan Simpanan Di Kedai (RM)
        Frm105.MSFlexGrid4.TextMatrix(x, 10) = Format(rs!duit_simpanan_kedai, "#,##0.00") 'Jumlah Kutipan Simpanan Di Kedai (RM)
    Else
        Frm105.MSFlexGrid4.TextMatrix(x, 10) = "0.00" 'Jumlah Kutipan Simpanan Di Kedai (RM)
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 22_jualan where (menu = 2 OR menu = 3) AND status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm105_LM_TOTAL_PAGE = Format(rs(0) / Frm105_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm105_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm105_LM_PAGE = Split(Frm105_LM_TOTAL_PAGE, ".")(0)
        Frm105_LM_PAGE_LEBIHAN = Split(Frm105_LM_TOTAL_PAGE, ".")(1)
        
        If Frm105_LM_PAGE_LEBIHAN <> "00" Then
            Frm105.L54_Text = Frm105_LM_PAGE + 1
        Else
            Frm105.L54_Text = Frm105_LM_PAGE
        End If
        
    Else
    
        Frm105.L54_Text = Frm105_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm105.L54_Text = 0
    End If
Else
    Frm105.L54_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm105.L54_Text = vbNullString Then
    Frm105.L54_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_bayaran) from 22_jualan where (menu = 2 OR menu = 3) AND status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm105_LM_JUMLAH = rs(0) 'Jumlah (RM)

rs.Close
Set rs = Nothing
'### Jumlah ### - End

'### Tunai Terkumpul ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(Tunai) from 22_jualan where (menu = 2 OR menu = 3) AND status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm105_LM_TUNAI = rs(0) 'Jumlah Tunai (RM)

rs.Close
Set rs = Nothing
'### Tunai Terkumpul ### - End

'### Bank In Terkumpul ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(bank_in) from 22_jualan where (menu = 2 OR menu = 3) AND status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm105_LM_BANK_IN = rs(0) 'Jumlah Bank In (RM)

rs.Close
Set rs = Nothing
'### Bank In Terkumpul ### - End

'### Kad Kredit Terkumpul ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(kad_kredit) from 22_jualan where (menu = 2 OR menu = 3) AND status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm105_LM_KAD_KREDIT = rs(0) 'Jumlah Kad Kredit (RM)

rs.Close
Set rs = Nothing
'### Kad Kredit Terkumpul ### - End

'### Kad Debit Terkumpul ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(kad_debit) from 22_jualan where (menu = 2 OR menu = 3) AND status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm105_LM_KAD_DEBIT = rs(0) 'Jumlah Kad Debit (RM)

rs.Close
Set rs = Nothing
'### Kad Debit Terkumpul ### - End

'### Simpanan Di Kedai Terkumpul ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(duit_simpanan_kedai) from 22_jualan where (menu = 2 OR menu = 3) AND status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm105_LM_SIMPANAN = rs(0) 'Jumlah Simpanan Di Kedai (RM)

rs.Close
Set rs = Nothing
'### Simpanan Di Kedai Terkumpul ### - End

Frm105.L47_Text = Format(Frm105_LM_JUMLAH, "#,##0.00") 'Jumlah (RM)
Frm105.L48_Text = Format(Frm105_LM_TUNAI, "#,##0.00") 'Jumlah Tunai (RM)
Frm105.L49_Text = Format(Frm105_LM_BANK_IN, "#,##0.00") 'Jumlah Bank In (RM)
Frm105.L50_Text = Format(Frm105_LM_KAD_KREDIT, "#,##0.00") 'Jumlah Kad Kredit (RM)
Frm105.L51_Text = Format(Frm105_LM_KAD_DEBIT, "#,##0.00") 'Jumlah Kad Debit (RM)
Frm105.L52_Text = Format(Frm105_LM_SIMPANAN, "#,##0.00") 'Jumlah Simpanan Di Kedai (RM)

If x <> 0 Then
    Frm105.L55_Text = LM_START_ROW 'Titik Pencarian Data
    Frm105.Pic6.Visible = True
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    If Frm105_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm105.L56_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm105.L56_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub Frm105_senarai_cash_in_header()
'on error resume next
'#### Header Report #### - Start
Frm105.MSFlexGrid5.Clear
Frm105.MSFlexGrid5.Rows = 1
Frm105.MSFlexGrid5.RowHeight(0) = 700
Frm105.MSFlexGrid5.FormatString = "No.|<No.|<No. ID|<Tarikh|<Jumlah (RM)"

Frm105.MSFlexGrid5.ColWidth(0) = 0 'No.
Frm105.MSFlexGrid5.ColWidth(1) = 600 'No.
Frm105.MSFlexGrid5.ColWidth(2) = 0 'No. ID
Frm105.MSFlexGrid5.ColWidth(3) = 2500 'Tarikh
Frm105.MSFlexGrid5.ColWidth(4) = 2400 'Jumlah (RM)
'#### Header Report #### - End
End Sub
Sub Frm105_senarai_cash_in()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim Frm105_LM_TOTAL_PAGE As Double
Dim Frm105_LM_JUMLAH As Double

Frm105_PAGE_SIZE = 34
Frm105_LM_TOTAL_PAGE = 0
x = 0
Frm105_LM_JUMLAH = 0

TM = Frm105.L5_Text 'Tarikh Mula
TA = Frm105.L6_Text 'Tarikh Akhir

Frm105.L58_Text = "Senarai kemasukkan tunai ke kedai dari " & TM & " hingga " & TA & "." 'Report Header

LM_START_ROW = Frm105.L62_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm105_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm105.L63_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm105_PAGE_SIZE
        End If
    End If
End If

Frm105_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 47_account_close where status='" & 1 & "' AND jenis='" & 0 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm105_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm105_LM_PAGE_FOUND = 0 Then
        If Frm105.L63_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm105.L60_Text = Frm105.L60_Text + 1 'Paparan Page ke-xxx
                Frm105_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm105.L60_Text) Then
                    If Frm105.L60_Text <> 1 Then
                        Frm105.L60_Text = Frm105.L60_Text - 1 'Paparan Page ke-xxx
                        Frm105_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm105.L60_Text - 1) * Frm105_PAGE_SIZE) + x

    Frm105.MSFlexGrid5.Rows = x + 1
    Frm105.MSFlexGrid5.TextMatrix(x, 0) = x 'No.
    Frm105.MSFlexGrid5.TextMatrix(x, 1) = Y 'No.
    Frm105.MSFlexGrid5.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm105.MSFlexGrid5.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!jumlah) Then 'Jumlah (RM)
        Frm105.MSFlexGrid5.TextMatrix(x, 4) = Format(rs!jumlah, "#,##0.00")
    Else
        Frm105.MSFlexGrid5.TextMatrix(x, 4) = "0.00"
    End If
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 47_account_close where status='" & 1 & "' AND jenis='" & 0 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm105_LM_TOTAL_PAGE = Format(rs(0) / Frm105_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm105_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm105_LM_PAGE = Split(Frm105_LM_TOTAL_PAGE, ".")(0)
        Frm105_LM_PAGE_LEBIHAN = Split(Frm105_LM_TOTAL_PAGE, ".")(1)
        
        If Frm105_LM_PAGE_LEBIHAN <> "00" Then
            Frm105.L61_Text = Frm105_LM_PAGE + 1
        Else
            Frm105.L61_Text = Frm105_LM_PAGE
        End If
        
    Else
    
        Frm105.L61_Text = Frm105_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm105.L61_Text = 0
    End If
Else
    Frm105.L61_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm105.L61_Text = vbNullString Then
    Frm105.L61_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah Kemasukkan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah) from 47_account_close where status='" & 1 & "' AND jenis='" & 0 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm105_LM_JUMLAH = rs(0) 'Jumlah (RM)

rs.Close
Set rs = Nothing
'### Jumlah Kemasukkan ### - End

Frm105.L59_Text = Format(Frm105_LM_JUMLAH, "#,##0.00") 'Jumlah (RM)

If x <> 0 Then
    Frm105.L62_Text = LM_START_ROW 'Titik Pencarian Data
    Frm105.Pic7.Visible = True
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    If Frm105_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm105.L63_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm105.L63_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub Frm105_simpanan_duit_header()
'on error resume next
'#### Header Report #### - Start
Frm105.MSFlexGrid6.Clear
Frm105.MSFlexGrid6.Rows = 1
Frm105.MSFlexGrid6.RowHeight(0) = 700
Frm105.MSFlexGrid6.FormatString = "No.|<No.|<No. ID|<Tarikh|<No. Invoice|<Jumlah (RM)"

Frm105.MSFlexGrid6.ColWidth(0) = 0 'No.
Frm105.MSFlexGrid6.ColWidth(1) = 600 'No.
Frm105.MSFlexGrid6.ColWidth(2) = 0 'No. ID
Frm105.MSFlexGrid6.ColWidth(3) = 2000 'Tarikh
Frm105.MSFlexGrid6.ColWidth(4) = 2000 'No. Invoice
Frm105.MSFlexGrid6.ColWidth(5) = 2000 'Jumlah (RM)
'#### Header Report #### - End
End Sub
Sub Frm105_simpanan_duit()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim Frm105_LM_TOTAL_PAGE As Double
Dim Frm105_LM_JUMLAH As Double

Frm105_PAGE_SIZE = 34
Frm105_LM_TOTAL_PAGE = 0
x = 0
Frm105_LM_JUMLAH = 0

TM = Frm105.L5_Text 'Tarikh Mula
TA = Frm105.L6_Text 'Tarikh Akhir

Frm105.L65_Text = "Senarai simpanan duit di kedai oleh pelanggan dari " & TM & " hingga " & TA & "." 'Report Header

LM_START_ROW = Frm105.L69_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm105_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm105.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm105_PAGE_SIZE
        End If
    End If
End If

Frm105_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 24_rekod_kewangan_pelanggan where tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND jenis='" & "0" & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm105_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm105_LM_PAGE_FOUND = 0 Then
        If Frm105.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm105.L67_Text = Frm105.L67_Text + 1 'Paparan Page ke-xxx
                Frm105_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm105.L67_Text) Then
                    If Frm105.L67_Text <> 1 Then
                        Frm105.L67_Text = Frm105.L67_Text - 1 'Paparan Page ke-xxx
                        Frm105_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm105.L67_Text - 1) * Frm105_PAGE_SIZE) + x

    Frm105.MSFlexGrid6.Rows = x + 1
    Frm105.MSFlexGrid6.TextMatrix(x, 0) = x 'No.
    Frm105.MSFlexGrid6.TextMatrix(x, 1) = Y 'No.
    Frm105.MSFlexGrid6.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm105.MSFlexGrid6.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!no_resit) Then Frm105.MSFlexGrid6.TextMatrix(x, 4) = rs!no_resit 'No. Invoice
    If Not IsNull(rs!jumlah) Then 'Jumlah (RM)
        Frm105.MSFlexGrid6.TextMatrix(x, 5) = Format(rs!jumlah, "#,##0.00")
    Else
        Frm105.MSFlexGrid6.TextMatrix(x, 5) = "0.00"
    End If
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 24_rekod_kewangan_pelanggan where tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND jenis='" & "0" & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm105_LM_TOTAL_PAGE = Format(rs(0) / Frm105_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm105_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm105_LM_PAGE = Split(Frm105_LM_TOTAL_PAGE, ".")(0)
        Frm105_LM_PAGE_LEBIHAN = Split(Frm105_LM_TOTAL_PAGE, ".")(1)
        
        If Frm105_LM_PAGE_LEBIHAN <> "00" Then
            Frm105.L68_Text = Frm105_LM_PAGE + 1
        Else
            Frm105.L68_Text = Frm105_LM_PAGE
        End If
        
    Else
    
        Frm105.L68_Text = Frm105_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm105.L68_Text = 0
    End If
Else
    Frm105.L68_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm105.L68_Text = vbNullString Then
    Frm105.L68_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah Kemasukkan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah) from 24_rekod_kewangan_pelanggan where tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND jenis='" & "0" & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm105_LM_JUMLAH = rs(0) 'Jumlah (RM)

rs.Close
Set rs = Nothing
'### Jumlah Kemasukkan ### - End

Frm105.L66_Text = Format(Frm105_LM_JUMLAH, "#,##0.00") 'Jumlah (RM)

If x <> 0 Then
    Frm105.L69_Text = LM_START_ROW 'Titik Pencarian Data
    Frm105.Pic8.Visible = True
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    If Frm105_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm105.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm105.L70_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub Frm105_belian_trade_in_header()
'on error resume next
'#### Header Report #### - Start
Frm105.MSFlexGrid7.Clear
Frm105.MSFlexGrid7.Rows = 1
Frm105.MSFlexGrid7.RowHeight(0) = 700
Frm105.MSFlexGrid7.FormatString = "No.|<No.|<No. ID|<Tarikh|<No. Voucher|<Tunai (RM)|<Bank In (RM)|<Jumlah (RM)|<Nama Pekerja"

Frm105.MSFlexGrid7.ColWidth(0) = 0 'No.
Frm105.MSFlexGrid7.ColWidth(1) = 600 'No.
Frm105.MSFlexGrid7.ColWidth(2) = 0 'No. ID
Frm105.MSFlexGrid7.ColWidth(3) = 1200 'Tarikh
Frm105.MSFlexGrid7.ColWidth(4) = 1500 'No. Voucher
Frm105.MSFlexGrid7.ColWidth(5) = 1500 'Tunai (RM)
Frm105.MSFlexGrid7.ColWidth(6) = 1500 'Bank In
Frm105.MSFlexGrid7.ColWidth(7) = 1500 'Jumlah
Frm105.MSFlexGrid7.ColWidth(8) = 2000 'Nama Pekerja
'#### Header Report #### - End
End Sub
Sub Frm105_belian_trade_in()
'On Error GoTo logging:
Dim TM As Date
Dim TA As Date
Dim Frm105_LM_TOTAL_PAGE As Double
Dim Frm105_LM_JUMLAH As Double

Frm105_PAGE_SIZE = 40
Frm105_LM_TOTAL_PAGE = 0
x = 0
Frm105_LM_JUMLAH = 0

TM = Frm105.L5_Text 'Tarikh Mula
TA = Frm105.L6_Text 'Tarikh Akhir
If Frm105.L110_Text = "Semua" Then
    Frm105_LM_SEARCH_1 = Null
    Frm105_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm105_LM_SEARCH_1 = Frm105.L110_Text
    Frm105_LM_SEARCH_1_LOGIC = "="
End If

Frm105.L72_Text = "Senarai belian trade in dari " & TM & " hingga " & TA & " bagi nama pekerja [" & Frm105.L111_Text & "]." 'Report Header

LM_START_ROW = Frm105.L76_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm105_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm105.L77_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm105_PAGE_SIZE
        End If
    End If
End If

Frm105_LM_PAGE_FOUND = 0
LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 16_gold_bar_belian where tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND no_pekerja " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND status = 1 AND flag_trade_in = 1 order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm105_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm105_LM_PAGE_FOUND = 0 Then
        If Frm105.L77_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm105.L74_Text = Frm105.L74_Text + 1 'Paparan Page ke-xxx
                Frm105_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm105.L74_Text) Then
                    If Frm105.L74_Text <> 1 Then
                        Frm105.L74_Text = Frm105.L74_Text - 1 'Paparan Page ke-xxx
                        Frm105_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm105.L74_Text - 1) * Frm105_PAGE_SIZE) + x

    Frm105.MSFlexGrid7.Rows = x + 1
    Frm105.MSFlexGrid7.TextMatrix(x, 0) = x 'No.
    Frm105.MSFlexGrid7.TextMatrix(x, 1) = Y 'No.
    Frm105.MSFlexGrid7.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm105.MSFlexGrid7.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!no_resit_trade_in) Then Frm105.MSFlexGrid7.TextMatrix(x, 4) = rs!no_resit_trade_in 'No. Voucher

    If Not IsNull(rs!tunai) Then 'Tunai (RM)
        Frm105.MSFlexGrid7.TextMatrix(x, 5) = Format(rs!tunai, "#,##0.00")
    Else
        Frm105.MSFlexGrid7.TextMatrix(x, 5) = "0.00"
    End If
    If Not IsNull(rs!bank_in) Then 'Bank In (RM)
        Frm105.MSFlexGrid7.TextMatrix(x, 6) = Format(rs!bank_in, "#,##0.00")
    Else
        Frm105.MSFlexGrid7.TextMatrix(x, 6) = "0.00"
    End If
    
    If Not IsNull(rs!jumlah_dengan_gst) Then 'Jumlah (RM)
        Frm105.MSFlexGrid7.TextMatrix(x, 7) = Format(rs!jumlah_dengan_gst, "#,##0.00")
    Else
        Frm105.MSFlexGrid7.TextMatrix(x, 7) = "0.00"
    End If
    If Not IsNull(rs!nama_pekerja) Then Frm105.MSFlexGrid7.TextMatrix(x, 8) = rs!nama_pekerja
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
LM_CONN = 2
re_conn_2:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID),SUM(jumlah_dengan_gst),SUM(tunai),SUM(bank_in) from 16_gold_bar_belian where tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND no_pekerja " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND status = 1 AND flag_trade_in = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm105_LM_TOTAL_PAGE = Format(rs(0) / Frm105_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm105_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm105_LM_PAGE = Split(Frm105_LM_TOTAL_PAGE, ".")(0)
        Frm105_LM_PAGE_LEBIHAN = Split(Frm105_LM_TOTAL_PAGE, ".")(1)
        
        If Frm105_LM_PAGE_LEBIHAN <> "00" Then
            Frm105.L75_Text = Frm105_LM_PAGE + 1
        Else
            Frm105.L75_Text = Frm105_LM_PAGE
        End If
        
    Else
    
        Frm105.L75_Text = Frm105_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm105.L75_Text = 0
    End If
Else
    Frm105.L75_Text = 0
End If

If Not IsNull(rs(1)) Then Frm105_LM_JUMLAH = rs(1) 'Jumlah (RM)
If Not IsNull(rs(2)) Then Frm105_LM_CASH = rs(2)
If Not IsNull(rs(3)) Then Frm105_LM_BANK_IN = rs(3)

rs.Close
Set rs = Nothing

If Frm105.L75_Text = vbNullString Then
    Frm105.L75_Text = 0
End If
'### Jumlah Data ### - End

Frm105.L73_Text = Format(Frm105_LM_JUMLAH, "#,##0.00") 'Jumlah (RM)
Frm105.L112_Text = Format(Frm105_LM_CASH, "#,##0.00") 'Jumlah (RM)
Frm105.L113_Text = Format(Frm105_LM_BANK_IN, "#,##0.00") 'Jumlah (RM)

If x <> 0 Then
    Frm105.L76_Text = LM_START_ROW 'Titik Pencarian Data
    Frm105.Pic10.Visible = True
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    If Frm105_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm105.L77_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm105.L77_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_report_kewangan : Frm105_belian_trade_in" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If LM_ERR_NO = "3704" Or LM_ERR_NO = "-2147467259" Or LM_ERR_NO = "-2147217887" Then
    
    Call Main

    If LM_CONN = 1 Then
        Resume re_conn_1:
    ElseIf LM_CONN = 2 Then
        Resume re_conn_2:
    End If
Else
    Resume Next
End If
End Sub
Sub Frm105_belian_barang_agen_header()
'on error resume next
'#### Header Report #### - Start
Frm105.MSFlexGrid8.Clear
Frm105.MSFlexGrid8.Rows = 1
Frm105.MSFlexGrid8.RowHeight(0) = 700
Frm105.MSFlexGrid8.FormatString = "No.|<No.|<No. ID|<Tarikh|<No. Voucher|<Jumlah (RM)"

Frm105.MSFlexGrid8.ColWidth(0) = 0 'No.
Frm105.MSFlexGrid8.ColWidth(1) = 600 'No.
Frm105.MSFlexGrid8.ColWidth(2) = 0 'No. ID
Frm105.MSFlexGrid8.ColWidth(3) = 2000 'Tarikh
Frm105.MSFlexGrid8.ColWidth(4) = 2000 'No. Voucher
Frm105.MSFlexGrid8.ColWidth(5) = 2000 'Jumlah (RM)
'#### Header Report #### - End
End Sub
Sub Frm105_belian_barang_agen()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim Frm105_LM_TOTAL_PAGE As Double
Dim Frm105_LM_JUMLAH As Double

Frm105_PAGE_SIZE = 34
Frm105_LM_TOTAL_PAGE = 0
x = 0
Frm105_LM_JUMLAH = 0

TM = Frm105.L5_Text 'Tarikh Mula
TA = Frm105.L6_Text 'Tarikh Akhir

Frm105.L79_Text = "Senarai belian barang dari agen dari " & TM & " hingga " & TA & "." 'Report Header

LM_START_ROW = Frm105.L83_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm105_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm105.L84_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm105_PAGE_SIZE
        End If
    End If
End If

Frm105_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 51_voucher_belian_agen where tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND flag_bayaran='" & "1" & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm105_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm105_LM_PAGE_FOUND = 0 Then
        If Frm105.L84_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm105.L81_Text = Frm105.L81_Text + 1 'Paparan Page ke-xxx
                Frm105_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm105.L81_Text) Then
                    If Frm105.L81_Text <> 1 Then
                        Frm105.L81_Text = Frm105.L81_Text - 1 'Paparan Page ke-xxx
                        Frm105_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm105.L81_Text - 1) * Frm105_PAGE_SIZE) + x

    Frm105.MSFlexGrid8.Rows = x + 1
    Frm105.MSFlexGrid8.TextMatrix(x, 0) = x 'No.
    Frm105.MSFlexGrid8.TextMatrix(x, 1) = Y 'No.
    Frm105.MSFlexGrid8.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm105.MSFlexGrid8.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!no_voucher) Then Frm105.MSFlexGrid8.TextMatrix(x, 4) = rs!no_voucher 'No. Voucher
    If Not IsNull(rs!harga_emas) Then 'Jumlah (RM)
        Frm105.MSFlexGrid8.TextMatrix(x, 5) = Format(rs!harga_emas, "#,##0.00")
    Else
        Frm105.MSFlexGrid8.TextMatrix(x, 5) = "0.00"
    End If
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 51_voucher_belian_agen where tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND flag_bayaran='" & "1" & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm105_LM_TOTAL_PAGE = Format(rs(0) / Frm105_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm105_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm105_LM_PAGE = Split(Frm105_LM_TOTAL_PAGE, ".")(0)
        Frm105_LM_PAGE_LEBIHAN = Split(Frm105_LM_TOTAL_PAGE, ".")(1)
        
        If Frm105_LM_PAGE_LEBIHAN <> "00" Then
            Frm105.L82_Text = Frm105_LM_PAGE + 1
        Else
            Frm105.L82_Text = Frm105_LM_PAGE
        End If
        
    Else
    
        Frm105.L82_Text = Frm105_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm105.L82_Text = 0
    End If
Else
    Frm105.L82_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm105.L82_Text = vbNullString Then
    Frm105.L82_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah belian item agen ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_emas) from 51_voucher_belian_agen where tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND flag_bayaran='" & "1" & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm105_LM_JUMLAH = rs(0) 'Jumlah (RM)

rs.Close
Set rs = Nothing
'### Jumlah belian item agen ### - End

Frm105.L80_Text = Format(Frm105_LM_JUMLAH, "#,##0.00") 'Jumlah (RM)

If x <> 0 Then
    Frm105.L83_Text = LM_START_ROW 'Titik Pencarian Data
    Frm105.Pic11.Visible = True
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    If Frm105_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm105.L84_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm105.L84_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub Frm105_ambilan_tunai_header()
'on error resume next
'#### Header Report #### - Start
Frm105.MSFlexGrid9.Clear
Frm105.MSFlexGrid9.Rows = 1
Frm105.MSFlexGrid9.RowHeight(0) = 700
Frm105.MSFlexGrid9.FormatString = "No.|<No.|<No. ID|<Tarikh|<Jumlah (RM)"

Frm105.MSFlexGrid9.ColWidth(0) = 0 'No.
Frm105.MSFlexGrid9.ColWidth(1) = 600 'No.
Frm105.MSFlexGrid9.ColWidth(2) = 0 'No. ID
Frm105.MSFlexGrid9.ColWidth(3) = 2000 'Tarikh
Frm105.MSFlexGrid9.ColWidth(4) = 2000 'Jumlah (RM)
'#### Header Report #### - End
End Sub
Sub Frm105_ambilan_tunai()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim Frm105_LM_TOTAL_PAGE As Double
Dim Frm105_LM_JUMLAH As Double

Frm105_PAGE_SIZE = 34
Frm105_LM_TOTAL_PAGE = 0
x = 0
Frm105_LM_JUMLAH = 0

TM = Frm105.L5_Text 'Tarikh Mula
TA = Frm105.L6_Text 'Tarikh Akhir

Frm105.L86_Text = "Senarai ambilan tunai kedai dari " & TM & " hingga " & TA & "." 'Report Header

LM_START_ROW = Frm105.L90_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm105_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm105.L91_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm105_PAGE_SIZE
        End If
    End If
End If

Frm105_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 47_account_close where status='" & 1 & "' AND jenis='" & 1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm105_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm105_LM_PAGE_FOUND = 0 Then
        If Frm105.L91_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm105.L88_Text = Frm105.L88_Text + 1 'Paparan Page ke-xxx
                Frm105_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm105.L88_Text) Then
                    If Frm105.L88_Text <> 1 Then
                        Frm105.L88_Text = Frm105.L88_Text - 1 'Paparan Page ke-xxx
                        Frm105_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm105.L88_Text - 1) * Frm105_PAGE_SIZE) + x

    Frm105.MSFlexGrid9.Rows = x + 1
    Frm105.MSFlexGrid9.TextMatrix(x, 0) = x 'No.
    Frm105.MSFlexGrid9.TextMatrix(x, 1) = Y 'No.
    Frm105.MSFlexGrid9.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm105.MSFlexGrid9.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!jumlah) Then 'Jumlah (RM)
        Frm105.MSFlexGrid9.TextMatrix(x, 4) = Format(rs!jumlah, "#,##0.00")
    Else
        Frm105.MSFlexGrid9.TextMatrix(x, 4) = "0.00"
    End If
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 47_account_close where status='" & 1 & "' AND jenis='" & 1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm105_LM_TOTAL_PAGE = Format(rs(0) / Frm105_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm105_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm105_LM_PAGE = Split(Frm105_LM_TOTAL_PAGE, ".")(0)
        Frm105_LM_PAGE_LEBIHAN = Split(Frm105_LM_TOTAL_PAGE, ".")(1)
        
        If Frm105_LM_PAGE_LEBIHAN <> "00" Then
            Frm105.L89_Text = Frm105_LM_PAGE + 1
        Else
            Frm105.L89_Text = Frm105_LM_PAGE
        End If
        
    Else
    
        Frm105.L89_Text = Frm105_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm105.L89_Text = 0
    End If
Else
    Frm105.L89_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm105.L89_Text = vbNullString Then
    Frm105.L89_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah ambilan duit ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah) from 47_account_close where status='" & 1 & "' AND jenis='" & 1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm105_LM_JUMLAH = rs(0) 'Jumlah (RM)

rs.Close
Set rs = Nothing
'### Jumlah ambilan duit ### - End

Frm105.L87_Text = Format(Frm105_LM_JUMLAH, "#,##0.00") 'Jumlah (RM)

If x <> 0 Then
    Frm105.L90_Text = LM_START_ROW 'Titik Pencarian Data
    Frm105.Pic12.Visible = True
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    If Frm105_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm105.L91_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm105.L91_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub Frm105_perbelanjaan_header()
'on error resume next
'#### Header Report #### - Start
Frm105.MSFlexGrid10.Clear
Frm105.MSFlexGrid10.Rows = 1
Frm105.MSFlexGrid10.RowHeight(0) = 700
Frm105.MSFlexGrid10.FormatString = "No.|<No.|<No. ID|<Tarikh|<No. Invoice / Resit|<Jumlah (RM)"

Frm105.MSFlexGrid10.ColWidth(0) = 0 'No.
Frm105.MSFlexGrid10.ColWidth(1) = 600 'No.
Frm105.MSFlexGrid10.ColWidth(2) = 0 'No. ID
Frm105.MSFlexGrid10.ColWidth(3) = 2000 'Tarikh
Frm105.MSFlexGrid10.ColWidth(4) = 2000 'No. Invoice / Resit
Frm105.MSFlexGrid10.ColWidth(5) = 2000 'Jumlah (RM)
'#### Header Report #### - End
End Sub
Sub Frm105_perbelanjaan()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim Frm105_LM_TOTAL_PAGE As Double
Dim Frm105_LM_JUMLAH As Double

Frm105_PAGE_SIZE = 34
Frm105_LM_TOTAL_PAGE = 0
x = 0
Frm105_LM_JUMLAH = 0

TM = Frm105.L5_Text 'Tarikh Mula
TA = Frm105.L6_Text 'Tarikh Akhir

Frm105.L93_Text = "Senarai perbelanjaan kedai dari " & TM & " hingga " & TA & "." 'Report Header

LM_START_ROW = Frm105.L97_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm105_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm105.L98_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm105_PAGE_SIZE
        End If
    End If
End If

Frm105_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 39_akaun_expense where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm105_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm105_LM_PAGE_FOUND = 0 Then
        If Frm105.L98_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm105.L95_Text = Frm105.L95_Text + 1 'Paparan Page ke-xxx
                Frm105_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm105.L95_Text) Then
                    If Frm105.L95_Text <> 1 Then
                        Frm105.L95_Text = Frm105.L95_Text - 1 'Paparan Page ke-xxx
                        Frm105_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm105.L95_Text - 1) * Frm105_PAGE_SIZE) + x

    Frm105.MSFlexGrid10.Rows = x + 1
    Frm105.MSFlexGrid10.TextMatrix(x, 0) = x 'No.
    Frm105.MSFlexGrid10.TextMatrix(x, 1) = Y 'No.
    Frm105.MSFlexGrid10.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm105.MSFlexGrid10.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!no_resit) Then Frm105.MSFlexGrid10.TextMatrix(x, 4) = rs!no_resit 'No. Resit
    If Not IsNull(rs!harga_dengan_gst) Then 'Jumlah (RM)
        Frm105.MSFlexGrid10.TextMatrix(x, 5) = Format(rs!harga_dengan_gst, "#,##0.00")
    Else
        Frm105.MSFlexGrid10.TextMatrix(x, 5) = "0.00"
    End If
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 39_akaun_expense where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm105_LM_TOTAL_PAGE = Format(rs(0) / Frm105_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm105_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm105_LM_PAGE = Split(Frm105_LM_TOTAL_PAGE, ".")(0)
        Frm105_LM_PAGE_LEBIHAN = Split(Frm105_LM_TOTAL_PAGE, ".")(1)
        
        If Frm105_LM_PAGE_LEBIHAN <> "00" Then
            Frm105.L96_Text = Frm105_LM_PAGE + 1
        Else
            Frm105.L96_Text = Frm105_LM_PAGE
        End If
        
    Else
    
        Frm105.L96_Text = Frm105_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm105.L96_Text = 0
    End If
Else
    Frm105.L96_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm105.L96_Text = vbNullString Then
    Frm105.L96_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah perbelanjaan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_dengan_gst) from 39_akaun_expense where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm105_LM_JUMLAH = rs(0) 'Jumlah (RM)

rs.Close
Set rs = Nothing
'### Jumlah perbelanjaan ### - End

Frm105.L94_Text = Format(Frm105_LM_JUMLAH, "#,##0.00") 'Jumlah (RM)

If x <> 0 Then
    Frm105.L97_Text = LM_START_ROW 'Titik Pencarian Data
    Frm105.Pic13.Visible = True
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    If Frm105_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm105.L98_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm105.L98_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub Frm105_bayaran_gaji_header()
'on error resume next
'#### Header Report #### - Start
Frm105.MSFlexGrid11.Clear
Frm105.MSFlexGrid11.Rows = 1
Frm105.MSFlexGrid11.RowHeight(0) = 700
Frm105.MSFlexGrid11.FormatString = "No.|<No.|<No. ID|<Tarikh|<Nama Pekerja|<Jumlah (RM)|<Tunai (RM)|<Bank In (RM)"

Frm105.MSFlexGrid11.ColWidth(0) = 0 'No.
Frm105.MSFlexGrid11.ColWidth(1) = 600 'No.
Frm105.MSFlexGrid11.ColWidth(2) = 0 'No. ID
Frm105.MSFlexGrid11.ColWidth(3) = 2000 'Tarikh
Frm105.MSFlexGrid11.ColWidth(4) = 2000 'Nama Pekerja
Frm105.MSFlexGrid11.ColWidth(5) = 2000 'Jumlah (RM)
Frm105.MSFlexGrid11.ColWidth(6) = 1800 'Tunai (RM)
Frm105.MSFlexGrid11.ColWidth(7) = 1800 'Bank In (RM)
'#### Header Report #### - End
End Sub
Sub Frm105_bayaran_gaji()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim Frm105_LM_TOTAL_PAGE As Double
Dim Frm105_LM_JUMLAH As Double
Dim Frm105_LM_TUNAI As Double
Dim Frm105_LM_BANK_IN As Double

Frm105_PAGE_SIZE = 34
Frm105_LM_TOTAL_PAGE = 0
x = 0
Frm105_LM_JUMLAH = 0
Frm105_LM_TUNAI = 0
Frm105_LM_BANK_IN = 0

TM = Frm105.L5_Text 'Tarikh Mula
TA = Frm105.L6_Text 'Tarikh Akhir

Frm105.L100_Text = "Senarai pembayaran gaji pekerja dari " & TM & " hingga " & TA & "." 'Report Header

LM_START_ROW = Frm105.L106_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm105_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm105.L107_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm105_PAGE_SIZE
        End If
    End If
End If

Frm105_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from payslip where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm105_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm105_LM_PAGE_FOUND = 0 Then
        If Frm105.L107_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm105.L104_Text = Frm105.L104_Text + 1 'Paparan Page ke-xxx
                Frm105_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm105.L104_Text) Then
                    If Frm105.L104_Text <> 1 Then
                        Frm105.L104_Text = Frm105.L104_Text - 1 'Paparan Page ke-xxx
                        Frm105_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm105.L104_Text - 1) * Frm105_PAGE_SIZE) + x

    Frm105.MSFlexGrid11.Rows = x + 1
    Frm105.MSFlexGrid11.TextMatrix(x, 0) = x 'No.
    Frm105.MSFlexGrid11.TextMatrix(x, 1) = Y 'No.
    Frm105.MSFlexGrid11.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm105.MSFlexGrid11.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!payroll_namapenuh) Then Frm105.MSFlexGrid11.TextMatrix(x, 4) = rs!payroll_namapenuh 'Nama Pekerja
    If Not IsNull(rs!payroll_bersih) Then 'Jumlah (RM)
        Frm105.MSFlexGrid11.TextMatrix(x, 5) = Format(rs!payroll_bersih, "#,##0.00")
    Else
        Frm105.MSFlexGrid11.TextMatrix(x, 5) = "0.00"
    End If
    If Not IsNull(rs!tunai) Then 'Jumlah (RM)
        Frm105.MSFlexGrid11.TextMatrix(x, 6) = Format(rs!tunai, "#,##0.00")
    Else
        Frm105.MSFlexGrid11.TextMatrix(x, 6) = "0.00"
    End If
    If Not IsNull(rs!bank_in) Then 'Jumlah (RM)
        Frm105.MSFlexGrid11.TextMatrix(x, 7) = Format(rs!bank_in, "#,##0.00")
    Else
        Frm105.MSFlexGrid11.TextMatrix(x, 7) = "0.00"
    End If
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from payslip where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm105_LM_TOTAL_PAGE = Format(rs(0) / Frm105_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm105_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm105_LM_PAGE = Split(Frm105_LM_TOTAL_PAGE, ".")(0)
        Frm105_LM_PAGE_LEBIHAN = Split(Frm105_LM_TOTAL_PAGE, ".")(1)
        
        If Frm105_LM_PAGE_LEBIHAN <> "00" Then
            Frm105.L105_Text = Frm105_LM_PAGE + 1
        Else
            Frm105.L105_Text = Frm105_LM_PAGE
        End If
        
    Else
    
        Frm105.L105_Text = Frm105_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm105.L105_Text = 0
    End If
Else
    Frm105.L105_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm105.L105_Text = vbNullString Then
    Frm105.L105_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah bayaran gaji ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(payroll_bersih) from payslip where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm105_LM_JUMLAH = rs(0) 'Jumlah (RM)

rs.Close
Set rs = Nothing
'### Jumlah bayaran gaji ### - End

'### Jumlah bayaran gaji (tunai) ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(tunai) from payslip where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm105_LM_TUNAI = rs(0) 'Jumlah Tunai (RM)

rs.Close
Set rs = Nothing
'### Jumlah bayaran gaji (tunai) ###  - End

'### Jumlah bayaran gaji (bank in) ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(bank_in) from payslip where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm105_LM_BANK_IN = rs(0) 'Jumlah Tunai (RM)

rs.Close
Set rs = Nothing
'### Jumlah bayaran gaji (bank in) ###  - End

Frm105.L101_Text = Format(Frm105_LM_JUMLAH, "#,##0.00") 'Jumlah (RM)
Frm105.L102_Text = Format(Frm105_LM_TUNAI, "#,##0.00") 'Jumlah (RM)
Frm105.L103_Text = Format(Frm105_LM_BANK_IN, "#,##0.00") 'Jumlah (RM)

If x <> 0 Then
    Frm105.L106_Text = LM_START_ROW 'Titik Pencarian Data
    Frm105.Pic14.Visible = True
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    If Frm105_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm105.L107_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm105.L107_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub Frm106_initial_setting()
'on error resume next
Frm106.L8_Text = "0.00" 'Senarai jualan : Jumlah
Frm106.L9_Text = "0.00" 'Senarai jualan : Tunai
Frm106.L10_Text = "0.00" 'Senarai jualan : Bank in
Frm106.L11_Text = "0.00" 'Senarai jualan : Kad kredit
Frm106.L12_Text = "0.00" 'Senarai jualan : Kad debit
Frm106.L13_Text = "0.00" 'Senarai jualan : Simpanan di kedai
Frm106.L14_Text = "0.00" 'Senarai servis : Jumlah
Frm106.L15_Text = "0.00" 'Senarai servis : Tunai
Frm106.L16_Text = "0.00" 'Senarai servis : Bank in
Frm106.L17_Text = "0.00" 'Senarai servis : Kad kredit
Frm106.L18_Text = "0.00" 'Senarai servis : Kad debit
Frm106.L19_Text = "0.00" 'Senarai servis : Simpanan di kedai
Frm106.L20_Text = "0.00" 'Senarai bayaran ansuran : Jumlah
Frm106.L21_Text = "0.00" 'Senarai bayaran ansuran : Tunai
Frm106.L22_Text = "0.00" 'Senarai bayaran ansuran : Bank in
Frm106.L23_Text = "0.00" 'Senarai bayaran ansuran : Kad kredit
Frm106.L24_Text = "0.00" 'Senarai bayaran ansuran : Kad debit
Frm106.L25_Text = "0.00" 'Senarai bayaran ansuran : Simpanan di kedai
Frm106.L26_Text = "0.00" 'Senarai bayaran tempahan : Jumlah
Frm106.L27_Text = "0.00" 'Senarai bayaran tempahan : Tunai
Frm106.L28_Text = "0.00" 'Senarai bayaran tempahan : Bank in
Frm106.L29_Text = "0.00" 'Senarai bayaran tempahan : Kad kredit
Frm106.L30_Text = "0.00" 'Senarai bayaran tempahan : Kad debit
Frm106.L31_Text = "0.00" 'Senarai bayaran tempahan : Simpanan di kedai
Frm106.L32_Text = "0.00" 'Senarai kemasukkan tunai ke kedai : Jumlah
Frm106.L33_Text = "0.00" 'Senarai kemasukkan tunai ke kedai : Tunai
Frm106.L34_Text = "0.00" 'Senarai simpanan duit di kedai oleh pelanggan : Jumlah
Frm106.L35_Text = "0.00" 'Senarai simpanan duit di kedai oleh pelanggan : Tunai
Frm106.L36_Text = "0.00" 'Debit : Jumlah
Frm106.L37_Text = "0.00" 'Debit : Tunai
Frm106.L38_Text = "0.00" 'Debit : Bank in
Frm106.L39_Text = "0.00" 'Debit : Kad kredit
Frm106.L40_Text = "0.00" 'Debit : Kad debit
Frm106.L41_Text = "0.00" 'Debit : Simpanan di kedai
Frm106.L42_Text = "0.00" 'Belian barang trade in : Jumlah
Frm106.L43_Text = "0.00" 'Belian barang trade in : Tunai
Frm106.L86_Text = "0.00"
Frm106.L44_Text = "0.00" 'Belian tukaran barang oleh agen : Jumlah
Frm106.L45_Text = "0.00" 'Belian tukaran barang oleh agen : Tunai
Frm106.L46_Text = "0.00" 'Ambilan tunai dari kedai : Jumlah
Frm106.L47_Text = "0.00" 'Ambilan tunai dari kedai : Tunai
Frm106.L48_Text = "0.00" 'Perbelanjaan kedai : Jumlah
Frm106.L49_Text = "0.00" 'Perbelanjaan kedai : Tunai
Frm106.L50_Text = "0.00" 'Bayaran gaji : Jumlah
Frm106.L51_Text = "0.00" 'Bayaran gaji : Tunai
Frm106.L52_Text = "0.00" 'Bayaran gaji : Bank in
Frm106.L53_Text = "0.00" 'Kredit : Jumlah
Frm106.L54_Text = "0.00" 'Kredit : Tunai
Frm106.L55_Text = "0.00" 'Kredit : Bank in
Frm106.L56_Text = "0.00" 'Kesimpulan : Tunai
Frm106.L57_Text = "0.00" 'Kesimpulan : Bank in
Frm106.L58_Text = "0.00" 'Kesimpulan : Kad kredit
Frm106.L59_Text = "0.00" 'Kesimpulan : Kad debit
Frm106.L60_Text = "0.00" 'Kesimpulan : Simpanan di kedai
Frm106.L61_Text = "0.00" 'Bayaran belian barang kemas terpakai yang dibayar secara tunai
Frm106.L62_Text = "0.00"
Frm106.L63_Text = "0.00" 'Yuran keahlian : Jumlah
Frm106.L64_Text = "0.00" 'Yuran keahlian : Tunai

Frm106.L70_Text = "0.00" 'Invoice GDN/GRN : Jumlah
Frm106.L71_Text = "0.00" 'Invoice GDN/GRN : Tunai
Frm106.L72_Text = "0.00" 'Invoice GDN/GRN : Bank In
Frm106.L73_Text = "0.00" 'Invoice GDN/GRN : Kad Kredit
Frm106.L74_Text = "0.00" 'Invoice GDN/GRN : Simpanan Di Kedai
Frm106.L79_Text = "0.00" 'Invoice GDN/GRN : Cek
Frm106.L80_Text = "0.00" 'Jumlah :Cek

Frm106.L75_Text = "0.00" 'Voucher GDN/GRN : Jumlah
Frm106.L76_Text = "0.00" 'Voucher GDN/GRN : Tunai
Frm106.L77_Text = "0.00" 'Voucher GDN/GRN : Bank In
Frm106.L78_Text = "0.00" 'Voucher GDN/GRN : Cek

Frm106.L81_Text = "0.00" 'Simpanan duit di kedai : Bank In

Frm106.L82_Text = "0.00" 'Pulangan duit pelanggan : Jumlah
Frm106.L83_Text = "0.00" 'Pulangan duit pelanggan : Tunai
Frm106.L84_Text = "0.00" 'Pulangan duit pelanggan : Bank In
Frm106.L85_Text = "0.00" 'Pulangan duit pelanggan : Cek
End Sub
Sub Frm106_penyata_akaun()
'On Error GoTo logging:
Dim TM As Date
Dim TA As Date
Dim Frm106_LM_JUALAN_JUMLAH As Double 'Senarai jualan : Jumlah
Dim Frm106_LM_JUALAN_TUNAI As Double 'Senarai jualan : Tunai
Dim Frm106_LM_JUALAN_BANK_IN As Double 'Senarai jualan : Bank in
Dim Frm106_LM_JUALAN_KREDIT As Double 'Senarai jualan : Kad kredit
Dim Frm106_LM_JUALAN_DEBIT As Double 'Senarai jualan : Kad debit
Dim Frm106_LM_JUALAN_SIMPANAN As Double 'Senarai jualan : Simpanan di kedai
Dim Frm106_LM_SERVIS_JUMLAH As Double 'Senarai servis : Jumlah
Dim Frm106_LM_SERVIS_TUNAI As Double 'Senarai servis : Tunai
Dim Frm106_LM_SERVIS_BANK_IN As Double 'Senarai servis : Bank in
Dim Frm106_LM_SERVIS_KREDIT As Double 'Senarai servis : Kad kredit
Dim Frm106_LM_SERVIS_DEBIT As Double 'Senarai servis : Kad debit
Dim Frm106_LM_SERVIS_SIMPANAN As Double 'Senarai servis : Simpanan di kedai
Dim Frm106_LM_ANSURAN_JUMLAH As Double 'Senarai bayaran ansuran : Jumlah
Dim Frm106_LM_ANSURAN_TUNAI As Double 'Senarai bayaran ansuran : Tunai
Dim Frm106_LM_ANSURAN_BANK_IN As Double 'Senarai bayaran ansuran : Bank in
Dim Frm106_LM_ANSURAN_KREDIT As Double 'Senarai bayaran ansuran : Kad kredit
Dim Frm106_LM_ANSURAN_DEBIT As Double 'Senarai bayaran ansuran : Kad debit
Dim Frm106_LM_ANSURAN_SIMPANAN As Double 'Senarai bayaran ansuran : Simpanan di kedai
Dim Frm106_LM_TEMPAHAN_JUMLAH As Double 'Senarai bayaran tempahan : Jumlah
Dim Frm106_LM_TEMPAHAN_TUNAI As Double 'Senarai bayaran tempahan : Tunai
Dim Frm106_LM_TEMPAHAN_BANK_IN As Double 'Senarai bayaran tempahan : Bank in
Dim Frm106_LM_TEMPAHAN_KREDIT As Double 'Senarai bayaran tempahan : Kad kredit
Dim Frm106_LM_TEMPAHAN_DEBIT As Double 'Senarai bayaran tempahan : Kad debit
Dim Frm106_LM_TEMPAHAN_SIMPANAN As Double 'Senarai bayaran tempahan : Simpanan di kedai
Dim Frm106_LM_CASH_IN_JUMLAH As Double 'Senarai kemasukkan tunai ke kedai : Jumlah
Dim Frm106_LM_SAVING_JUMLAH As Double 'Senarai simpanan duit di kedai oleh pelanggan : Jumlah
Dim Frm106_LM_TRADE_IN_JUMLAH As Double 'Belian barang trade in : Jumlah (RM)
Dim Frm106_LM_AGEN_JUMLAH As Double 'Belian tukaran barang oleh agen : Jumlah
Dim Frm106_LM_CASH_OUT_JUMLAH As Double 'Ambilan tunai dari kedai : Jumlah
Dim Frm106_LM_EXPENSES_JUMLAH As Double 'Perbelanjaan kedai : Jumlah
Dim Frm106_LM_PAYSLIP_JUMLAH As Double 'Bayaran gaji : Jumlah
Dim Frm106_LM_PAYSLIP_TUNAI As Double 'Bayaran gaji : Tunai
Dim Frm106_LM_PAYSLIP_BANK_IN As Double 'Bayaran gaji : Bank In
Dim Frm106_LM_AHLI_JUMLAH As Double 'Yuran keahlian : Jumlah
Dim Frm106_LM_AHLI_TUNAI As Double 'Yuran keahlian : Tunai

Dim Frm106_LM_JUALAN_TI As Double 'Jumlah bayaran yang dibayar secara trade in (Pelanggan -> kedai) : Tunai
Dim Frm106_LM_JUALAN_TI_LEBIH As Double 'Jumlah lebihan bayaran trade in (Kedai -> Pelanggan) : Tunai
Dim Frm106_LM_JUALAN_TI2 As Double 'Jumlah bayaran yang dibayar secara trade in (Pelanggan -> kedai) : Tunai : Tempahan
Dim Frm106_LM_JUALAN_TI_LEBIH2 As Double 'Jumlah lebihan bayaran trade in (Kedai -> Pelanggan) : Tunai : Tempahan
Dim Frm106_LM_TRADE_IN_CASH_JUMLAH As Double 'Jumlah trade in barang yang ambil cash : Tunai
Dim Frm106_LM_EXPENSES_TUNAI As Double
Dim Frm106_LM_EXPENSES_BANK As Double
Dim Frm106_LM_EXPENSES_CEK As Double
Dim Frm106_LM_TRADE_IN_BANK_JUMLAH As Double

Dim Frm106_LM_INVOICE_JUMLAH As Double
Dim Frm106_LM_INVOICE_TUNAI As Double
Dim Frm106_LM_INVOICE_BANK_IN As Double
Dim Frm106_LM_INVOICE_SIMPANAN As Double
Dim Frm106_LM_INVOICE_CEK As Double
Dim Frm106_LM_VOUCHER_JUMLAH As Double
Dim Frm106_LM_VOUCHER_TUNAI As Double
Dim Frm106_LM_VOUCHER_BANK_IN As Double
Dim Frm106_LM_VOUCHER_CEK As Double

Dim Frm106_LM_SIMPANAN_TUNAI As Double
Dim Frm106_LM_SIMPANAN_BANK_IN As Double
Dim Frm106_LM_PULANGAN_JUMLAH As Double
Dim Frm106_LM_PULANGAN_TUNAI As Double
Dim Frm106_LM_PULANGAN_BANK_IN As Double
Dim Frm106_LM_PULANGAN_CEK As Double

Frm106_LM_SIMPANAN_TUNAI = 0
Frm106_LM_SIMPANAN_BANK_IN = 0
Frm106_LM_PULANGAN_JUMLAH = 0
Frm106_LM_PULANGAN_TUNAI = 0
Frm106_LM_PULANGAN_BANK_IN = 0
Frm106_LM_PULANGAN_CEK = 0

Frm106_LM_EXPENSES_TUNAI = 0
Frm106_LM_EXPENSES_BANK = 0
Frm106_LM_EXPENSES_CEK = 0

Call Frm106_initial_setting

TM = Frm105.L5_Text 'Tarikh Mula
TA = Frm105.L6_Text 'Tarikh Akhir

Frm106_LM_INVOICE_JUMLAH = 0 'Invoice GDN/GRN : Jumlah
Frm106_LM_INVOICE_TUNAI = 0 'Invoice GDN/GRN : Tunai
Frm106_LM_INVOICE_BANK_IN = 0 'Invoice GDN/GRN : Bank In
Frm106_LM_INVOICE_SIMPANAN = 0 'Invoice GDN/GRN : Simpanan Di Kedai
Frm106_LM_INVOICE_CEK = 0 'Invoice GDN/GRN : Cek

Frm106_LM_VOUCHER_JUMLAH = 0 'Voucher GDN/GRN : Jumlah
Frm106_LM_VOUCHER_TUNAI = 0 'Voucher GDN/GRN : Tunai
Frm106_LM_VOUCHER_BANK_IN = 0 'Voucher GDN/GRN : Bank In
Frm106_LM_VOUCHER_CEK = 0 'Voucher GDN/GRN : Cek

Frm106_LM_JUALAN_JUMLAH = 0 'Senarai jualan : Jumlah
Frm106_LM_JUALAN_TUNAI = 0 'Senarai jualan : Tunai
Frm106_LM_JUALAN_BANK_IN = 0 'Senarai jualan : Bank in
Frm106_LM_JUALAN_KREDIT = 0 'Senarai jualan : Kad kredit
Frm106_LM_JUALAN_DEBIT = 0 'Senarai jualan : Kad debit
Frm106_LM_JUALAN_SIMPANAN = 0 'Senarai jualan : Simpanan di kedai
Frm106_LM_SERVIS_JUMLAH = 0 'Senarai servis : Jumlah
Frm106_LM_SERVIS_TUNAI = 0 'Senarai servis : Tunai
Frm106_LM_SERVIS_BANK_IN = 0 'Senarai servis : Bank in
Frm106_LM_SERVIS_KREDIT = 0 'Senarai servis : Kad kredit
Frm106_LM_SERVIS_DEBIT = 0 'Senarai servis : Kad debit
Frm106_LM_SERVIS_SIMPANAN = 0 'Senarai servis : Simpanan di kedai
Frm106_LM_ANSURAN_JUMLAH = 0 'Senarai bayaran ansuran : Jumlah
Frm106_LM_ANSURAN_TUNAI = 0 'Senarai bayaran ansuran : Tunai
Frm106_LM_ANSURAN_BANK_IN = 0 'Senarai bayaran ansuran : Bank in
Frm106_LM_ANSURAN_KREDIT = 0 'Senarai bayaran ansuran : Kad kredit
Frm106_LM_ANSURAN_DEBIT = 0 'Senarai bayaran ansuran : Kad debit
Frm106_LM_ANSURAN_SIMPANAN = 0 'Senarai bayaran ansuran : Simpanan di kedai
Frm106_LM_TEMPAHAN_JUMLAH = 0 'Senarai bayaran tempahan : Jumlah
Frm106_LM_TEMPAHAN_TUNAI = 0 'Senarai bayaran tempahan : Tunai
Frm106_LM_TEMPAHAN_BANK_IN = 0 'Senarai bayaran tempahan : Bank in
Frm106_LM_TEMPAHAN_KREDIT = 0 'Senarai bayaran tempahan : Kad kredit
Frm106_LM_TEMPAHAN_DEBIT = 0 'Senarai bayaran tempahan : Kad debit
Frm106_LM_TEMPAHAN_SIMPANAN = 0 'Senarai bayaran tempahan : Simpanan di kedai
Frm106_LM_CASH_IN_JUMLAH = 0 'Senarai kemasukkan tunai ke kedai : Jumlah
Frm106_LM_SAVING_JUMLAH = 0 'Senarai simpanan duit di kedai oleh pelanggan : Jumlah
Frm106_LM_TRADE_IN_JUMLAH = 0 'Belian barang trade in : Jumlah (RM)
Frm106_LM_AGEN_JUMLAH = 0 'Belian tukaran barang oleh agen : Jumlah
Frm106_LM_CASH_OUT_JUMLAH = 0 'Ambilan tunai dari kedai : Jumlah
Frm106_LM_EXPENSES_JUMLAH = 0 'Perbelanjaan kedai : Jumlah
Frm106_LM_PAYSLIP_JUMLAH = 0 'Bayaran gaji : Jumlah
Frm106_LM_PAYSLIP_TUNAI = 0 'Bayaran gaji : Tunai
Frm106_LM_PAYSLIP_BANK_IN = 0 'Bayaran gaji : Bank In
Frm106_LM_JUALAN_TI = 0 'Jumlah bayaran yang dibayar secara trade in (Pelanggan -> kedai) : Tunai
Frm106_LM_JUALAN_TI_LEBIH = 0 'Jumlah lebihan bayaran trade in (Kedai -> Pelanggan) : Tunai
Frm106_LM_TRADE_IN_CASH_JUMLAH = 0 'Jumlah trade in barang yang ambil cash : Tunai
Frm106_LM_TRADE_IN_BANK_JUMLAH = 0
Frm106_LM_AHLI_JUMLAH = 0 'Yuran keahlian : Jumlah
Frm106_LM_AHLI_TUNAI = 0 'Yuran keahlian : Tunai
Frm106_LM_JUALAN_TI2 = 0 'Jumlah bayaran yang dibayar secara trade in (Pelanggan -> kedai) : Tunai : Tempahan
Frm106_LM_JUALAN_TI_LEBIH2 = 0 'Jumlah lebihan bayaran trade in (Kedai -> Pelanggan) : Tunai : Tempahan

If Frm105.L110_Text = "Semua" Then
    Frm105_LM_SEARCH_1 = Null
    Frm105_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm105_LM_SEARCH_1 = Frm105.L110_Text
    Frm105_LM_SEARCH_1_LOGIC = "="
End If

'====================================================================== Jualan - Start
'### Jumlah ### - Start
LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_perlu_bayar) from 22_jualan where status = 1 AND no_pekerja " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND flag_bayaran = 0 AND menu = 0 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_JUALAN_JUMLAH = rs(0) 'Jumlah (RM)

rs.Close
Set rs = Nothing
'### Jumlah ### - End

'### Jumlah bayaran oleh pelanggan menggunakan barang trade in ### - Start
LM_CONN = 2
re_conn_2:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_trade_in) from 22_jualan where flag_bayaran = 0 AND no_pekerja " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND status = 1 AND flag_trade_in = 1 AND menu = 0 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_JUALAN_TI = rs(0) 'Jumlah bayaran yang dibayar secara trade in (Pelanggan -> kedai) : Tunai (RM)

rs.Close
Set rs = Nothing
'### Jumlah bayaran oleh pelanggan menggunakan barang trade in ###  - End

'### Jumlah lebihan bayaran trade in yang kedai bayar kepada pelanggan ### - Start
LM_CONN = 3
re_conn_3:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_perlu_bayar) from 22_jualan where flag_bayaran = 1 AND no_pekerja " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND status = 1 AND flag_trade_in = 1 AND menu = 0 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_JUALAN_TI_LEBIH = rs(0) 'Jumlah lebihan bayaran trade in (Kedai -> Pelanggan) : Tunai (RM)
'jumlah_perlu_bayar
rs.Close
Set rs = Nothing
'### Jumlah lebihan bayaran trade in yang kedai bayar kepada pelanggan ###   - End

'### Tunai Terkumpul ### - Start
LM_CONN = 4
re_conn_4:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(Tunai),SUM(bank_in),SUM(kad_kredit),SUM(kad_debit),SUM(duit_simpanan_kedai) from 22_jualan where flag_bayaran = 0 AND no_pekerja " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND status = 1 AND menu = 0 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_JUALAN_TUNAI = rs(0) 'Jumlah Tunai (RM)
If Not IsNull(rs(1)) Then Frm106_LM_JUALAN_BANK_IN = rs(1) 'Jumlah Bank In (RM)
If Not IsNull(rs(2)) Then Frm106_LM_JUALAN_KREDIT = rs(2) 'Jumlah Kad Kredit (RM)
If Not IsNull(rs(3)) Then Frm106_LM_JUALAN_DEBIT = rs(3) 'Jumlah Kad Debit (RM)
If Not IsNull(rs(4)) Then Frm106_LM_JUALAN_SIMPANAN = rs(4) 'Jumlah Simpanan Di Kedai (RM)

rs.Close
Set rs = Nothing
'### Tunai Terkumpul ### - End
'====================================================================== Jualan - End
'====================================================================== Servis - Start
'### Jumlah ### - Start
LM_CONN = 5
re_conn_5:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_perlu_bayar) from 22_jualan where tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND no_pekerja " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND status = 1 AND flag_bayaran = 0 AND menu = 1", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_SERVIS_JUMLAH = rs(0) 'Jumlah (RM)

rs.Close
Set rs = Nothing
'### Jumlah ### - End

'### Tunai Terkumpul ### - Start
LM_CONN = 6
re_conn_6:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(Tunai),SUM(bank_in),SUM(kad_kredit),SUM(duit_simpanan_kedai) from 22_jualan where flag_bayaran = 0 AND no_pekerja " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND status = 1 AND menu = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_SERVIS_TUNAI = rs(0) 'Jumlah Tunai (RM)
If Not IsNull(rs(1)) Then Frm106_LM_SERVIS_BANK_IN = rs(1) 'Jumlah Bank In (RM)
If Not IsNull(rs(2)) Then Frm106_LM_SERVIS_KREDIT = rs(2) 'Jumlah Kad Kredit (RM)
If Not IsNull(rs(3)) Then Frm106_LM_SERVIS_SIMPANAN = rs(3) 'Jumlah Simpanan Di Kedai (RM)

rs.Close
Set rs = Nothing
'### Tunai Terkumpul ### - End
'====================================================================== Servis - Start

'### Jumlah yuran keahlian ### - Start
LM_CONN = 7
re_conn_7:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_yuran) from senarai_pelanggan where yuran_flag='" & "1" & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then
    Frm106_LM_AHLI_JUMLAH = rs(0) 'Yuran keahlian : Jumlah
    Frm106_LM_AHLI_TUNAI = rs(0) 'Yuran keahlian : Tunai
End If

rs.Close
Set rs = Nothing
'### Jumlah yuran keahlian ### - End

GoTo skip_ansuran:
'====================================================================== Senarai bayaran ansuran - Start
'### Jumlah ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah) from 29_akaun_ansuran where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_ANSURAN_JUMLAH = rs(0) 'Jumlah (RM)

rs.Close
Set rs = Nothing
'### Jumlah ### - End

'### Tunai Terkumpul ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(Tunai) from 29_akaun_ansuran where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_ANSURAN_TUNAI = rs(0) 'Jumlah Tunai (RM)

rs.Close
Set rs = Nothing
'### Tunai Terkumpul ### - End

'### Bank In Terkumpul ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(bank_in) from 29_akaun_ansuran where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_ANSURAN_BANK_IN = rs(0) 'Jumlah Bank In (RM)

rs.Close
Set rs = Nothing
'### Bank In Terkumpul ### - End

'### Kad Kredit Terkumpul ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(kad_kredit) from 29_akaun_ansuran where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_ANSURAN_KREDIT = rs(0) 'Jumlah Kad Kredit (RM)

rs.Close
Set rs = Nothing
'### Kad Kredit Terkumpul ### - End

'### Kad Debit Terkumpul ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(kad_debit) from 29_akaun_ansuran where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_ANSURAN_DEBIT = rs(0) 'Jumlah Kad Debit (RM)

rs.Close
Set rs = Nothing
'### Kad Debit Terkumpul ### - End

'### Simpanan Di Kedai Terkumpul ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(duit_simpanan_kedai) from 29_akaun_ansuran where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_ANSURAN_SIMPANAN = rs(0) 'Jumlah Simpanan Di Kedai (RM)

rs.Close
Set rs = Nothing
'### Simpanan Di Kedai Terkumpul ### - End
'====================================================================== Senarai bayaran ansuran - End
skip_ansuran:

'====================================================================== Tempahan - Start
'### Jumlah ### - Start
LM_CONN = 8
re_conn_8:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_perlu_bayar) from 22_jualan where tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND no_pekerja " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND status = 1 AND flag_bayaran = 0 AND (menu = 2 OR menu = 3)", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_TEMPAHAN_JUMLAH = rs(0) 'Jumlah (RM)

rs.Close
Set rs = Nothing
'### Jumlah ### - End

'### Jumlah bayaran oleh pelanggan menggunakan barang trade in ### - Start
LM_CONN = 9
re_conn_9:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_trade_in) from 22_jualan where flag_bayaran = 0 AND no_pekerja " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND status = 1 AND flag_trade_in = 1 AND (menu = 2 OR menu = 3) AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_JUALAN_TI2 = rs(0) 'Jumlah bayaran yang dibayar secara trade in (Pelanggan -> kedai) : Tunai (RM)

rs.Close
Set rs = Nothing
'### Jumlah bayaran oleh pelanggan menggunakan barang trade in ###  - End

'### Jumlah lebihan bayaran trade in yang kedai bayar kepada pelanggan ### - Start
LM_CONN = 10
re_conn_10:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_perlu_bayar) from 22_jualan where flag_bayaran = 1 AND no_pekerja " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND status = 1 AND flag_trade_in = 1 AND (menu = 2 OR menu = 3) AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_JUALAN_TI_LEBIH2 = rs(0) 'Jumlah lebihan bayaran trade in (Kedai -> Pelanggan) : Tunai (RM)
'jumlah_perlu_bayar
rs.Close
Set rs = Nothing
'### Jumlah lebihan bayaran trade in yang kedai bayar kepada pelanggan ###   - End

'### Tunai Terkumpul ### - Start
LM_CONN = 11
re_conn_11:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(Tunai),SUM(bank_in),SUM(kad_kredit),SUM(duit_simpanan_kedai) from 22_jualan where flag_bayaran = 0 AND no_pekerja " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND status = 1 AND (menu = 2 OR menu = 3) AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_TEMPAHAN_TUNAI = rs(0) 'Jumlah Tunai (RM)
If Not IsNull(rs(1)) Then Frm106_LM_TEMPAHAN_BANK_IN = rs(1) 'Jumlah Bank In (RM)
If Not IsNull(rs(2)) Then Frm106_LM_TEMPAHAN_KREDIT = rs(2) 'Jumlah Kad Kredit (RM)
If Not IsNull(rs(3)) Then Frm106_LM_TEMPAHAN_SIMPANAN = rs(3) 'Jumlah Simpanan Di Kedai (RM)

rs.Close
Set rs = Nothing
'### Tunai Terkumpul ### - End
'====================================================================== Tempahan - End

'====================================================================== Senarai kemasukkan tunai ke kedai - Start
'### Jumlah Kemasukkan ### - Start
LM_CONN = 12
re_conn_12:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah) from 47_account_close where status='" & 1 & "' AND staff_id " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND jenis='" & 0 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_CASH_IN_JUMLAH = rs(0) 'Jumlah (RM)

rs.Close
Set rs = Nothing
'### Jumlah Kemasukkan ### - End
'====================================================================== Senarai kemasukkan tunai ke kedai - End

'====================================================================== Senarai simpanan duit di kedai oleh pelanggan - Start
'### Jumlah Kemasukkan ### - Start
LM_CONN = 13
re_conn_13:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah) , SUM(tunai) , SUM(bank_in) from 24_rekod_kewangan_pelanggan where tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND jenis='" & "0" & "' AND status = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_SAVING_JUMLAH = rs(0) 'Jumlah (RM)
If Not IsNull(rs(1)) Then Frm106_LM_SIMPANAN_TUNAI = rs(1) 'Tunai (RM)
If Not IsNull(rs(2)) Then Frm106_LM_SIMPANAN_BANK_IN = rs(2) 'Bank In (RM)

rs.Close
Set rs = Nothing
'### Jumlah Kemasukkan ### - End
'====================================================================== Senarai simpanan duit di kedai oleh pelanggan - End

'====================================================================== Pulangan duit pelanggan - Start
'### Jumlah Kemasukkan ### - Start
LM_CONN = 14
re_conn_14:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah) , SUM(tunai) , SUM(bank_in) , sum(cek) from 24_rekod_kewangan_pelanggan where tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND jenis='" & "2" & "' AND status = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_PULANGAN_JUMLAH = rs(0) 'Jumlah (RM)
If Not IsNull(rs(1)) Then Frm106_LM_PULANGAN_TUNAI = rs(1) 'Tunai (RM)
If Not IsNull(rs(2)) Then Frm106_LM_PULANGAN_BANK_IN = rs(2) 'Bank In (RM)
If Not IsNull(rs(3)) Then Frm106_LM_PULANGAN_CEK = rs(3) 'Cek (RM)

rs.Close
Set rs = Nothing
'### Jumlah Kemasukkan ### - End
'====================================================================== Pulangan duit pelanggan - End

'====================================================================== Belian barang trade in - Start
'### Jumlah trade in ### - Start
LM_CONN = 15
re_conn_15:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_dengan_gst) from 16_gold_bar_belian where tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND no_pekerja " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND status = 1 AND flag_trade_in = 1", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_TRADE_IN_JUMLAH = rs(0) 'Jumlah (RM)

rs.Close
Set rs = Nothing
'### Jumlah trade in ### - End
'====================================================================== Belian barang trade in - End

'====================================================================== Belian barang trade in - Start
'### Jumlah trade in (Yang ambil duit) ### - Start
LM_CONN = 16
re_conn_16:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(tunai),SUM(bank_in) from 16_gold_bar_belian where trade_in_status = 0 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND no_pekerja " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND status = 1 AND flag_trade_in = 1", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_TRADE_IN_CASH_JUMLAH = rs(0) 'Tunai (RM)
If Not IsNull(rs(1)) Then Frm106_LM_TRADE_IN_BANK_JUMLAH = rs(1) 'Bank In (RM)

rs.Close
Set rs = Nothing
'### Jumlah trade in (Yang ambil duit) ### - End
'====================================================================== Belian barang trade in - End

'====================================================================== Belian tukaran barang oleh agen - Start
'### Jumlah belian item agen ### - Start
LM_CONN = 17
re_conn_17:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_emas) from 51_voucher_belian_agen where tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND flag_bayaran='" & "1" & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_AGEN_JUMLAH = rs(0) 'Jumlah (RM)

rs.Close
Set rs = Nothing
'### Jumlah belian item agen ### - End
'====================================================================== Belian tukaran barang oleh agen - End

'====================================================================== Ambilan tunai dari kedai - Start
'### Jumlah ambilan duit ### - Start
LM_CONN = 18
re_conn_18:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah) from 47_account_close where status='" & 1 & "' AND staff_id " & Frm105_LM_SEARCH_1_LOGIC & "'" & Frm105_LM_SEARCH_1 & "' AND jenis='" & 1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_CASH_OUT_JUMLAH = rs(0) 'Jumlah (RM)

rs.Close
Set rs = Nothing
'### Jumlah ambilan duit ### - End
'====================================================================== Ambilan tunai dari kedai - End

'====================================================================== Perbelanjaan kedai - Start
'### Jumlah perbelanjaan ### - Start
LM_CONN = 19
re_conn_19:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_dengan_gst) from 39_akaun_expense where status = 1 AND menu = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_EXPENSES_JUMLAH = Format(rs(0), "#,##0.00") 'Jumlah (RM)

rs.Close
Set rs = Nothing
'### Jumlah perbelanjaan ### - End

'### Jumlah perbelanjaan (Tunai) ### - Start
LM_CONN = 20
re_conn_20:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_dengan_gst) from 39_akaun_expense where status = 1 AND menu = 1 AND cara_bayaran = 0 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_EXPENSES_TUNAI = Format(rs(0), "#,##0.00") 'Jumlah Tunai (RM)

rs.Close
Set rs = Nothing
'### Jumlah perbelanjaan (Tunai) ### - End

'### Jumlah perbelanjaan (Bank In) ### - Start
LM_CONN = 21
re_conn_21:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_dengan_gst) from 39_akaun_expense where status = 1 AND menu = 1 AND cara_bayaran = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_EXPENSES_BANK = Format(rs(0), "#,##0.00") 'Jumlah Bank In (RM)

rs.Close
Set rs = Nothing
'### Jumlah perbelanjaan (Bank In) ### - End

'### Jumlah perbelanjaan (Cek) ### - Start
LM_CONN = 22
re_conn_22:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_dengan_gst) from 39_akaun_expense where status = 1 AND menu = 1 AND cara_bayaran = 2 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_EXPENSES_CEK = Format(rs(0), "#,##0.00") 'Jumlah Cek (RM)

rs.Close
Set rs = Nothing
'### Jumlah perbelanjaan (Cek) ### - End
'====================================================================== Perbelanjaan kedai - End

'====================================================================== Bayaran gaji - Start
'### Jumlah bayaran gaji ### - Start
LM_CONN = 23
re_conn_23:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(payroll_bersih),SUM(tunai),SUM(bank_in) from payslip where tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm106_LM_PAYSLIP_JUMLAH = rs(0) 'Jumlah (RM)
If Not IsNull(rs(1)) Then Frm106_LM_PAYSLIP_TUNAI = rs(1) 'Jumlah Tunai (RM)
If Not IsNull(rs(2)) Then Frm106_LM_PAYSLIP_BANK_IN = rs(2) 'Jumlah Tunai (RM)

rs.Close
Set rs = Nothing
'### Jumlah bayaran gaji ### - End
'====================================================================== Bayaran gaji - End

'Debit=================== Start
Frm106.L8_Text = Format(Frm106_LM_JUALAN_JUMLAH, "#,##0.00") 'Senarai jualan : Jumlah
Frm106.L9_Text = Format(Frm106_LM_JUALAN_TUNAI, "#,##0.00") 'Senarai jualan : Tunai
Frm106.L10_Text = Format(Frm106_LM_JUALAN_BANK_IN, "#,##0.00") 'Senarai jualan : Bank in
Frm106.L11_Text = Format(Frm106_LM_JUALAN_KREDIT, "#,##0.00") 'Senarai jualan : Kad kredit
Frm106.L12_Text = Format(Frm106_LM_JUALAN_DEBIT, "#,##0.00") 'Senarai jualan : Kad debit
Frm106.L13_Text = Format(Frm106_LM_JUALAN_SIMPANAN, "#,##0.00") 'Senarai jualan : Simpanan di kedai
Frm106.L14_Text = Format(Frm106_LM_SERVIS_JUMLAH, "#,##0.00") 'Senarai servis : Jumlah (RM)
Frm106.L15_Text = Format(Frm106_LM_SERVIS_TUNAI, "#,##0.00") 'Senarai servis : Tunai (RM)
Frm106.L16_Text = Format(Frm106_LM_SERVIS_BANK_IN, "#,##0.00") 'Senarai servis : Bank In (RM)
Frm106.L17_Text = Format(Frm106_LM_SERVIS_KREDIT, "#,##0.00") 'Senarai servis : Kad Kredit (RM)
Frm106.L18_Text = Format(Frm106_LM_SERVIS_DEBIT, "#,##0.00") 'Senarai servis : Kad Debit (RM)
Frm106.L19_Text = Format(Frm106_LM_SERVIS_SIMPANAN, "#,##0.00") 'Senarai servis : Simpanan Di Kedai (RM)
Frm106.L20_Text = Format(Frm106_LM_ANSURAN_JUMLAH, "#,##0.00") 'Senarai bayaran ansuran : Jumlah (RM)
Frm106.L21_Text = Format(Frm106_LM_ANSURAN_TUNAI, "#,##0.00") 'Senarai bayaran ansuran : Tunai (RM)
Frm106.L22_Text = Format(Frm106_LM_ANSURAN_BANK_IN, "#,##0.00") 'Senarai bayaran ansuran : Bank In (RM)
Frm106.L23_Text = Format(Frm106_LM_ANSURAN_KREDIT, "#,##0.00") 'Senarai bayaran ansuran : Kad Kredit (RM)
Frm106.L24_Text = Format(Frm106_LM_ANSURAN_DEBIT, "#,##0.00") 'Senarai bayaran ansuran : Kad Debit (RM)
Frm106.L25_Text = Format(Frm106_LM_ANSURAN_SIMPANAN, "#,##0.00") 'Senarai bayaran ansuran : Simpanan Di Kedai (RM)
Frm106.L26_Text = Format(Frm106_LM_TEMPAHAN_JUMLAH, "#,##0.00") 'Senarai bayaran tempahan : Jumlah (RM)
Frm106.L27_Text = Format(Frm106_LM_TEMPAHAN_TUNAI, "#,##0.00") 'Senarai bayaran tempahan : Tunai (RM)
Frm106.L28_Text = Format(Frm106_LM_TEMPAHAN_BANK_IN, "#,##0.00") 'Senarai bayaran tempahan : Bank In (RM)
Frm106.L29_Text = Format(Frm106_LM_TEMPAHAN_KREDIT, "#,##0.00") 'Senarai bayaran tempahan : Kad Kredit (RM)
Frm106.L30_Text = Format(Frm106_LM_TEMPAHAN_DEBIT, "#,##0.00") 'Senarai bayaran tempahan : Kad Debit (RM)
Frm106.L31_Text = Format(Frm106_LM_TEMPAHAN_SIMPANAN, "#,##0.00") 'Senarai bayaran tempahan : Simpanan Di Kedai (RM)
Frm106.L32_Text = Format(Frm106_LM_CASH_IN_JUMLAH, "#,##0.00") 'Senarai kemasukkan tunai ke kedai : Jumlah (RM)
Frm106.L33_Text = Format(Frm106_LM_CASH_IN_JUMLAH, "#,##0.00") 'Senarai kemasukkan tunai ke kedai : Tunai (RM)
Frm106.L34_Text = Format(Frm106_LM_SAVING_JUMLAH, "#,##0.00") 'Senarai simpanan duit di kedai oleh pelanggan : Jumlah (RM)
Frm106.L35_Text = Format(Frm106_LM_SIMPANAN_TUNAI, "#,##0.00") 'Senarai simpanan duit di kedai oleh pelanggan : Tunai (RM)
Frm106.L81_Text = Format(Frm106_LM_SIMPANAN_BANK_IN, "#,##0.00") 'Senarai simpanan duit di kedai oleh pelanggan : Tunai (RM)

Frm106.L63_Text = Format(Frm106_LM_AHLI_JUMLAH, "#,##0.00") 'Yuran keahlian : Jumlah (RM)
Frm106.L64_Text = Format(Frm106_LM_AHLI_TUNAI, "#,##0.00") 'Yuran keahlian : Tunai (RM)
'Debit=================== End

'Kredit=================== Start
Frm106.L42_Text = Format(Frm106_LM_TRADE_IN_JUMLAH, "#,##0.00") 'Belian barang trade in : Jumlah (RM)
Frm106.L43_Text = Format(Frm106_LM_TRADE_IN_CASH_JUMLAH, "#,##0.00") 'Belian barang trade in : Tunai (RM)
Frm106.L86_Text = Format(Frm106_LM_TRADE_IN_BANK_JUMLAH, "#,##0.00") 'Belian barang trade in : Bank In (RM)
Frm106.L44_Text = Format(Frm106_LM_AGEN_JUMLAH, "#,##0.00") 'Belian tukaran barang oleh agen : Jumlah (RM)
Frm106.L45_Text = Format(Frm106_LM_AGEN_JUMLAH, "#,##0.00") 'Belian tukaran barang oleh agen : Tunai (RM)
Frm106.L46_Text = Format(Frm106_LM_CASH_OUT_JUMLAH, "#,##0.00") 'Ambilan tunai dari kedai : Jumlah (RM)
Frm106.L47_Text = Format(Frm106_LM_CASH_OUT_JUMLAH, "#,##0.00") 'Ambilan tunai dari kedai : Tunai (RM)
Frm106.L48_Text = Format(Frm106_LM_EXPENSES_JUMLAH, "#,##0.00") 'Perbelanjaan kedai : Jumlah (RM)
Frm106.L49_Text = Format(Frm106_LM_EXPENSES_TUNAI, "#,##0.00") 'Perbelanjaan kedai : Tunai (RM)

Frm106.L50_Text = Format(Frm106_LM_PAYSLIP_JUMLAH, "#,##0.00") 'Bayaran gaji : Jumlah (RM)
Frm106.L51_Text = Format(Frm106_LM_PAYSLIP_TUNAI, "#,##0.00") 'Bayaran gaji : Tunai (RM)
Frm106.L52_Text = Format(Frm106_LM_PAYSLIP_BANK_IN, "#,##0.00") 'Bayaran gaji : Bank in (RM)

Frm106.L82_Text = Format(Frm106_LM_PULANGAN_JUMLAH, "#,##0.00") 'Pulangan duit pelanggan : Jumlah (RM)
Frm106.L83_Text = Format(Frm106_LM_PULANGAN_TUNAI, "#,##0.00") 'Pulangan duit pelanggan : Tunai (RM)
Frm106.L84_Text = Format(Frm106_LM_PULANGAN_BANK_IN, "#,##0.00") 'Pulangan duit pelanggan : Bank In (RM)
Frm106.L85_Text = Format(Frm106_LM_PULANGAN_CEK, "#,##0.00") 'Pulangan duit pelanggan : Cek (RM)
'Kredit=================== End

Frm106.L65_Text = Format(Frm106_LM_EXPENSES_BANK, "#,##0.00") 'Perbelanjaan kedai : Bank in (RM)
Frm106.L66_Text = Format(Frm106_LM_EXPENSES_CEK, "#,##0.00") 'Perbelanjaan kedai : Cek (RM)

Frm106.L70_Text = Format(Frm106_LM_INVOICE_JUMLAH, "#,##0.00") 'Invoice GDN/GRN : Jumlah
Frm106.L71_Text = Format(Frm106_LM_INVOICE_TUNAI, "#,##0.00") 'Invoice GDN/GRN : Tunai
Frm106.L72_Text = Format(Frm106_LM_INVOICE_BANK_IN, "#,##0.00") 'Invoice GDN/GRN : Bank In
Frm106.L73_Text = Format(0, "#,##0.00") 'Invoice GDN/GRN : Kad Kredit
Frm106.L74_Text = Format(Frm106_LM_INVOICE_SIMPANAN, "#,##0.00") 'Invoice GDN/GRN : Simpanan Di Kedai
Frm106.L79_Text = Format(Frm106_LM_INVOICE_CEK, "#,##0.00") 'Invoice GDN/GRN : Cek
Frm106.L80_Text = Format(Frm106_LM_INVOICE_CEK, "#,##0.00") 'Jumlah :Cek

Frm106.L75_Text = Format(Frm106_LM_VOUCHER_JUMLAH, "#,##0.00") 'Voucher GDN/GRN : Jumlah
Frm106.L76_Text = Format(Frm106_LM_VOUCHER_TUNAI, "#,##0.00") 'Voucher GDN/GRN : Tunai
Frm106.L77_Text = Format(Frm106_LM_VOUCHER_BANK_IN, "#,##0.00") 'Voucher GDN/GRN : Bank In
Frm106.L78_Text = Format(Frm106_LM_VOUCHER_CEK, "#,##0.00") 'Voucher GDN/GRN : Cek



'#### Summary
Frm106.L36_Text = Format(Frm106_LM_INVOICE_JUMLAH + Frm106_LM_JUALAN_JUMLAH + Frm106_LM_SERVIS_JUMLAH + Frm106_LM_ANSURAN_JUMLAH + Frm106_LM_TEMPAHAN_JUMLAH + Frm106_LM_CASH_IN_JUMLAH + Frm106_LM_SAVING_JUMLAH + Frm106_LM_AHLI_JUMLAH, "#,##0.00") 'Debit : Jumlah
Frm106.L37_Text = Format(Frm106_LM_INVOICE_TUNAI + Frm106_LM_JUALAN_TUNAI + Frm106_LM_SERVIS_TUNAI + Frm106_LM_ANSURAN_TUNAI + Frm106_LM_TEMPAHAN_TUNAI + Frm106_LM_CASH_IN_JUMLAH + Frm106_LM_SIMPANAN_TUNAI + Frm106_LM_AHLI_TUNAI, "#,##0.00") 'Debit : Tunai
Frm106.L38_Text = Format(Frm106_LM_INVOICE_BANK_IN + Frm106_LM_JUALAN_BANK_IN + Frm106_LM_SERVIS_BANK_IN + Frm106_LM_ANSURAN_BANK_IN + Frm106_LM_TEMPAHAN_BANK_IN + Frm106_LM_SIMPANAN_BANK_IN, "#,##0.00") 'Debit : Bank in
Frm106.L39_Text = Format(Frm106_LM_JUALAN_KREDIT + Frm106_LM_SERVIS_KREDIT + Frm106_LM_ANSURAN_KREDIT + Frm106_LM_TEMPAHAN_KREDIT, "#,##0.00") 'Debit : Kad kredit
Frm106.L40_Text = Format(Frm106_LM_JUALAN_DEBIT + Frm106_LM_SERVIS_DEBIT + Frm106_LM_ANSURAN_DEBIT + Frm106_LM_TEMPAHAN_DEBIT, "#,##0.00") 'Debit : Kad debit
Frm106.L41_Text = Format(Frm106_LM_JUALAN_SIMPANAN + Frm106_LM_SERVIS_SIMPANAN + Frm106_LM_ANSURAN_SIMPANAN + Frm106_LM_TEMPAHAN_SIMPANAN, "#,##0.00") 'Debit : Simpanan di kedai

Frm106.L53_Text = Format(Frm106_LM_PULANGAN_JUMLAH + Frm106_LM_VOUCHER_JUMLAH + Frm106_LM_TRADE_IN_JUMLAH + Frm106_LM_AGEN_JUMLAH + Frm106_LM_CASH_OUT_JUMLAH + Frm106_LM_EXPENSES_JUMLAH + Frm106_LM_PAYSLIP_JUMLAH, "#,##0.00") 'Kredit : Jumlah
'Frm106.L54_Text = Format(Frm106_LM_VOUCHER_TUNAI + Frm106_LM_TRADE_IN_JUMLAH + Frm106_LM_AGEN_JUMLAH + Frm106_LM_CASH_OUT_JUMLAH + Frm106_LM_EXPENSES_JUMLAH + Frm106_LM_PAYSLIP_TUNAI + Frm106_LM_EXPENSES_TUNAI, "#,##0.00") 'Kredit : Tunai
Frm106.L54_Text = Format(Frm106_LM_PULANGAN_TUNAI + Frm106_LM_VOUCHER_TUNAI + Frm106_LM_TRADE_IN_JUMLAH + Frm106_LM_AGEN_JUMLAH + Frm106_LM_CASH_OUT_JUMLAH + Frm106_LM_PAYSLIP_TUNAI + Frm106_LM_EXPENSES_TUNAI, "#,##0.00") 'Kredit : Tunai
Frm106.L55_Text = Format(Frm106_LM_PULANGAN_BANK_IN + Frm106_LM_VOUCHER_BANK_IN + Frm106_LM_PAYSLIP_BANK_IN + Frm106_LM_EXPENSES_BANK, "#,##0.00") 'Kredit : Bank in

'Frm106.L56_Text = Format(Frm106_LM_JUALAN_TUNAI + Frm106_LM_SERVIS_TUNAI + Frm106_LM_ANSURAN_TUNAI + Frm106_LM_TEMPAHAN_TUNAI + Frm106_LM_CASH_IN_JUMLAH + Frm106_LM_SAVING_JUMLAH - Frm106_LM_TRADE_IN_JUMLAH - Frm106_LM_AGEN_JUMLAH - Frm106_LM_CASH_OUT_JUMLAH - Frm106_LM_EXPENSES_JUMLAH - Frm106_LM_PAYSLIP_TUNAI, "#,##0.00") 'Kesimpulan : Tunai
Frm106.L56_Text = Format(Frm106_LM_SIMPANAN_TUNAI + Frm106_LM_INVOICE_TUNAI + Frm106_LM_JUALAN_TUNAI + Frm106_LM_SERVIS_TUNAI + Frm106_LM_ANSURAN_TUNAI + Frm106_LM_TEMPAHAN_TUNAI + Frm106_LM_CASH_IN_JUMLAH + Frm106_LM_AHLI_TUNAI - Frm106_LM_VOUCHER_TUNAI - Frm106_LM_TRADE_IN_CASH_JUMLAH - Frm106_LM_JUALAN_TI_LEBIH - Frm106_LM_JUALAN_TI_LEBIH2 - Frm106_LM_AGEN_JUMLAH - Frm106_LM_CASH_OUT_JUMLAH - Frm106_LM_EXPENSES_TUNAI - Frm106_LM_PAYSLIP_TUNAI - Frm106_LM_PULANGAN_TUNAI, "#,##0.00") 'Kesimpulan : Tunai
Frm106.L57_Text = Format(Frm106_LM_SIMPANAN_BANK_IN + Frm106_LM_INVOICE_BANK_IN + Frm106_LM_JUALAN_BANK_IN + Frm106_LM_SERVIS_BANK_IN + Frm106_LM_ANSURAN_BANK_IN + Frm106_LM_TEMPAHAN_BANK_IN - Frm106_LM_TRADE_IN_BANK_JUMLAH - Frm106_LM_VOUCHER_BANK_IN - Frm106_LM_PAYSLIP_BANK_IN - Frm106_LM_EXPENSES_BANK - Frm106_LM_PULANGAN_BANK_IN, "#,##0.00") 'Kesimpulan : Bank in
Frm106.L58_Text = Format(Frm106_LM_JUALAN_KREDIT + Frm106_LM_SERVIS_KREDIT + Frm106_LM_ANSURAN_KREDIT + Frm106_LM_TEMPAHAN_KREDIT, "#,##0.00") 'Kesimpulan : Kad kredit
Frm106.L59_Text = Format(Frm106_LM_JUALAN_DEBIT + Frm106_LM_SERVIS_DEBIT + Frm106_LM_ANSURAN_DEBIT + Frm106_LM_TEMPAHAN_DEBIT, "#,##0.00") 'Kesimpulan : Kad debit
Frm106.L60_Text = Format(Frm106_LM_INVOICE_SIMPANAN + Frm106_LM_JUALAN_SIMPANAN + Frm106_LM_SERVIS_SIMPANAN + Frm106_LM_ANSURAN_SIMPANAN + Frm106_LM_TEMPAHAN_SIMPANAN, "#,##0.00") 'Kesimpulan : Simpanan di kedai

Frm106.L61_Text = Format(Frm106_LM_TRADE_IN_CASH_JUMLAH + Frm106_LM_TRADE_IN_BANK_JUMLAH + Frm106_LM_JUALAN_TI_LEBIH + Frm106_LM_JUALAN_TI_LEBIH2, "#,##0.00")
Frm106.L62_Text = Format(Frm106_LM_JUALAN_TI + Frm106_LM_JUALAN_TI, "#,##0.00")

Frm106.L67_Text = Format(Frm106_LM_PULANGAN_CEK + Frm106_LM_EXPENSES_CEK + Frm106_LM_VOUCHER_CEK, "#,##0.00")

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_report_kewangan : Frm106_penyata_akaun" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If LM_ERR_NO = "3704" Or LM_ERR_NO = "-2147467259" Or LM_ERR_NO = "-2147217887" Then
    
    Call Main

    If LM_CONN = 1 Then
        Resume re_conn_1:
    ElseIf LM_CONN = 2 Then
        Resume re_conn_2:
    ElseIf LM_CONN = 3 Then
        Resume re_conn_3:
    ElseIf LM_CONN = 4 Then
        Resume re_conn_4:
    ElseIf LM_CONN = 5 Then
        Resume re_conn_5:
    ElseIf LM_CONN = 6 Then
        Resume re_conn_6:
    ElseIf LM_CONN = 7 Then
        Resume re_conn_7:
    ElseIf LM_CONN = 8 Then
        Resume re_conn_8:
    ElseIf LM_CONN = 9 Then
        Resume re_conn_9:
    ElseIf LM_CONN = 10 Then
        Resume re_conn_10:
    ElseIf LM_CONN = 11 Then
        Resume re_conn_11:
    ElseIf LM_CONN = 12 Then
        Resume re_conn_12:
    ElseIf LM_CONN = 13 Then
        Resume re_conn_13:
    ElseIf LM_CONN = 14 Then
        Resume re_conn_14:
    ElseIf LM_CONN = 15 Then
        Resume re_conn_15:
    ElseIf LM_CONN = 16 Then
        Resume re_conn_16:
    ElseIf LM_CONN = 17 Then
        Resume re_conn_17:
    ElseIf LM_CONN = 18 Then
        Resume re_conn_18:
    ElseIf LM_CONN = 19 Then
        Resume re_conn_19:
    ElseIf LM_CONN = 20 Then
        Resume re_conn_20:
    ElseIf LM_CONN = 21 Then
        Resume re_conn_21:
    ElseIf LM_CONN = 22 Then
        Resume re_conn_22:
    ElseIf LM_CONN = 23 Then
        Resume re_conn_23:
    End If
Else
    Resume Next
End If
End Sub

