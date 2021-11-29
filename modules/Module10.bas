Attribute VB_Name = "Module10"
Sub Frm96_initial()
'on error resume next
Frm96.CBB1.Clear

Frm96.CBB1.AddItem "Semua cawangan"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where status='" & 1 & "' order by cawangan ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!cawangan) Then Frm96.CBB1.AddItem rs!cawangan
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
End Sub
Sub Frm26_initial()
'on error resume next
Frm26.TB1 = vbNullString
Frm26.TB2 = vbNullString
End Sub
Sub Frm27_initial()
'on error resume next
Frm27.TB1 = vbNullString
Frm27.L1_Text = vbNullString
Frm27.L2_Text = vbNullString
Frm27.L3_Text = vbNullString
Frm27.L4_Text = vbNullString
Frm27.L5_Text = vbNullString

Frm27.L69_Text = -1 'Titik Pencarian Data
Frm27.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Frm27.L67_Text = 0 'Paparan Page ke-xxx
Frm27.L68_Text = 0
Frm27.L71_Text = 0
End Sub
Sub Frm28_initial()
'on error resume next
Frm28.TB1 = vbNullString
Frm28.L1_Text = vbNullString
Frm28.L2_Text = vbNullString
Frm28.L3_Text = vbNullString
Frm28.L4_Text = vbNullString
Frm28.L5_Text = vbNullString

Frm28.TB2 = vbNullString
Frm28.TB3 = vbNullString
Frm28.TB4 = vbNullString

Frm28.L69_Text = -1 'Titik Pencarian Data
Frm28.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Frm28.L67_Text = 0 'Paparan Page ke-xxx
Frm28.L68_Text = 0
Frm28.L71_Text = 0

Exit Sub

Frm28.CB2 = 0
Frm28.CB3 = 1

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Default1 = "Default" Then
        If Not IsNull(rs!ScannerMode) Then
            If rs!ScannerMode = 1 Then
                Frm28.CB1 = 1
            Else
                Frm28.CB1 = 0
            End If
        Else
            Frm28.CB1 = 0
        End If
    End If
End If

rs.Close
Set rs = Nothing
End Sub
Sub Frm28_carian_ahli()
'on error resume next
Dim rs3 As ADODB.Recordset
Dim Frm28_LM_FIELD As String

Frm28_LM_KATEGORI_ASAL = 1

If Frm28.CB2 = 1 Then
    Frm28_LM_FIELD = "no_ic"
ElseIf Frm28.CB3 = 1 Then
    Frm28_LM_FIELD = "no_pelanggan"
End If

Set rs3 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs3.Open "select * from senarai_pelanggan where " & Frm28_LM_FIELD & "='" & UCase(Frm28.TB1) & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs3.EOF Then

    Call Frm28_initial
    
    If Not IsNull(rs3!Nama) Then Frm28.L1_Text = rs3!Nama 'Nama
    If Not IsNull(rs3!no_ic) Then Frm28.L2_Text = rs3!no_ic 'No. Kad Pengenalan
    If Not IsNull(rs3!no_tel) Then Frm28.L3_Text = rs3!no_tel 'No. Telefon
    If Not IsNull(rs3!Email) Then Frm28.L4_Text = rs3!Email 'E-mail
    If Not IsNull(rs3!no_pelanggan) Then Frm28.L5_Text = rs3!no_pelanggan 'No. Ahli
    
    If Not IsNull(rs3!baki_simpanan) Then
    
        If Not IsNull(rs3!baki_simpanan) Then
            
            frm130.L26_Text = Format(rs3!baki_simpanan, "#,##0.00")
            If MDI_frm1.L5_Text = 7 Then Frm87.L27_Text = Format(rs3!baki_simpanan, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
            If MDI_frm1.L5_Text = 10 Then frm130.L26_Text = Format(rs3!baki_simpanan, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
            If MDI_frm1.L5_Text = 8 Then frm130.L26_Text = Format(rs3!baki_simpanan, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
            If MDI_frm1.L5_Text = 9 Then frm130.L26_Text = Format(rs3!baki_simpanan, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
        
        Else
            
            frm130.L26_Text = Format(0, "0.00")
            If MDI_frm1.L5_Text = 7 Then Frm87.L27_Text = Format(0, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
            If MDI_frm1.L5_Text = 10 Then frm130.L26_Text = Format(0, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
            If MDI_frm1.L5_Text = 8 Then frm130.L26_Text = Format(0, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
            If MDI_frm1.L5_Text = 9 Then frm130.L26_Text = Format(0, "0.00") 'Baki Simpanan Pelanggan Ini (RM)
        
        End If
        
    End If
    
    If MDI_frm1.L5_Text = 4 Or MDI_frm1.L5_Text = 5 Then
    
        If Not IsNull(rs3!membership_card) Then
            If rs3!membership_card = 0 Then
            
                If MDI_frm1.L5_Text = 4 Or MDI_frm1.L5_Text = 5 Then
                    Frm84.L79_Text = 0 '0 : Tiada kad , 1 : Ada kad
                End If
    
                Frm84.L77_Text = "0"
    
            ElseIf rs3!membership_card = 1 Then
            
                If MDI_frm1.L5_Text = 4 Or MDI_frm1.L5_Text = 5 Then
                    Frm84.L79_Text = 1 '0 : Tiada kad , 1 : Ada kad
                End If
                If Not IsNull(rs3!baki_point) Then
                    Frm84.L77_Text = rs3!baki_point
                Else
                    Frm84.L77_Text = "0"
                End If
            End If
        Else
        
            If MDI_frm1.L5_Text = 4 Or MDI_frm1.L5_Text = 5 Then
                Frm84.L79_Text = 0 '0 : Tiada kad , 1 : Ada kad
            End If
    
            Frm84.L77_Text = "0"
    
        End If
        
    End If
    
Else

    Frm28.TB1 = vbNullString
    MsgBox "Tiada maklumat dijumpai ATAU data pelanggan ini sudah tidak aktif.", vbInformation, "Info"
    
    Frm28.TB1.SetFocus
    
End If

rs3.Close
Set rs3 = Nothing
End Sub
Sub frm133_setting()
'on error resume next
frm133.CBB1.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where SenaraiDulang <> '" & Null & "' AND status='" & 1 & "' order by SenaraiDulang ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!SenaraiDulang) Then frm133.CBB1.AddItem rs!SenaraiDulang
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
End Sub
Sub frm132_setting()
'on error resume next
frm132.CB1 = 1
frm132.TB1 = vbNullString
frm132.L1_Text = vbNullString
frm132.L2_Text = vbNullString
End Sub
Sub frm132_tukar_dulang()
'on error resume next
DATA_UDPATE = 0

LM_NOW = Now

LM_BARCODE = UCase(frm132.TB1)

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from data_database where no_siri_produk='" & UCase(frm132.TB1) & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Not IsNull(rs!cawangan) Then
        
        If MDI_frm1.L20_Text <> rs!cawangan Then
            
            MsgBox "Stok ini adalah milik cawangan [" & rs!cawangan & "]. Anda tidak dibenarkan untuk tukar maklumat dulang bagi barang ini.", vbExclamation, "Info"
            
            frm132.TB1 = vbNullString
            frm132.TB1.SetFocus
            
            rs.Close
            Set rs = Nothing
            
            Exit Sub
            
        End If
    
    End If
 
    If Not IsNull(rs!StatusItem) Then
    
        If rs!StatusItem = "10" Then
            
            GM_No_RUJUKAN_BELIAN = rs!no_siri_Produk 'No. Siri Produk
            If Not IsNull(rs!cawangan) Then G_KEDAI = rs!cawangan
            If Not IsNull(rs!dulang) Then LM_DULANG_ASAL = rs!dulang
            rs!dulang = frm132.L1_Text
            rs.Update
            DATA_UDPATE = 1
            
        Else
        
            MsgBox "Anda tidak dibenarkan untuk ubah data dulang bagi item ini kerana status item ini telah berubah." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila periksa status terbaru item ini.", vbExclamation, "Info"
                    
            frm132.TB1 = vbNullString
            frm132.TB1.SetFocus
            
            rs.Close
            Set rs = Nothing
            
            Exit Sub
            
        End If
        
    End If
            
End If

rs.Close
Set rs = Nothing

If DATA_UDPATE = 1 Then
    
'### Update Log ### - Start
    user = MDI_frm1.L3_Text
    LogAct_Memory = "[" & user & "] Tukar data dulang bagi [" & GM_No_RUJUKAN_BELIAN & "] dari " & LM_DULANG_ASAL & " ke " & frm132.L1_Text & "."
    LogDate_Memory = LM_NOW
    Call UpdateLog_Database
'### Update Log ### - End

    'frm132.L2_Text.Print LogAct_Memory
    Call Print_All_Barcode2
    
    frm132.L2_Text = LM_BARCODE & " BERJAYA DIUPDATE"
    
Else

    frm132.L2_Text = LM_BARCODE & " TIDAK DIUPDATE"
    
End If

frm132.TB1 = vbNullString
frm132.TB1.SetFocus

End Sub
Sub frm28_senarai_pelanggan_header()
'on error resume next
With Frm28.LV1
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear

    Frm28.LV1.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "Nama", 4600
    .ColumnHeaders.Add 5, , "No. Kad Pengenalan", 2100
    .ColumnHeaders.Add 6, , "No. Telefon", 1600
    .ColumnHeaders.Add 7, , "No. Keahlian", 1700
    .ColumnHeaders.Add 8, , "Kategori", 2000
    
End With
End Sub
Sub frm28_senarai_pelanggan()
'On Error Resume Next
Dim frm28_LM_TOTAL_PAGE As Double

frm28_PAGE_SIZE = 12
frm28_LM_TOTAL_PAGE = 0
x = 0
Frm28_LM_SEARCH_1 = "%" & Frm28.L72_Text & "%"

re_gen_report:

LM_START_ROW = Frm28.L69_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + frm28_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm28.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - frm28_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        Frm28.L67_Text = 1
    End If
End If

Frm28.L71_Text = 0

frm28_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from senarai_pelanggan where (nama LIKE'" & Frm28_LM_SEARCH_1 & "' OR no_ic LIKE'" & Frm28_LM_SEARCH_1 & "' OR no_tel LIKE'" & Frm28_LM_SEARCH_1 & "' OR no_pelanggan LIKE'" & Frm28_LM_SEARCH_1 & "') AND status = 1 order by nama ASC LIMIT " & LM_START_ROW & "," & frm28_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1
    If frm28_LM_PAGE_FOUND = 0 Then
        If Frm28.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm28.L67_Text = Frm28.L67_Text + 1 'Paparan Page ke-xxx
                frm28_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm28.L67_Text) Then
                    If Frm28.L67_Text <> 1 Then
                        Frm28.L67_Text = Frm28.L67_Text - 1 'Paparan Page ke-xxx
                        frm28_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If

    Y = ((Frm28.L67_Text - 1) * frm28_PAGE_SIZE) + x

    With Frm28.LV1.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!Nama) Then
            .ListSubItems.Add , , rs!Nama
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_ic) Then
            .ListSubItems.Add , , rs!no_ic
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_tel) Then
            .ListSubItems.Add , , rs!no_tel
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_pelanggan) Then
            .ListSubItems.Add , , rs!no_pelanggan
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kategori_pelanggan) Then 'Kategori
        
            If rs!kategori_pelanggan = 1 Then .ListSubItems.Add , , "Pelanggan Biasa"
            If rs!kategori_pelanggan = 2 Then .ListSubItems.Add , , "Ahli Biasa"
            If rs!kategori_pelanggan = 3 Then .ListSubItems.Add , , "Silver"
            If rs!kategori_pelanggan = 4 Then .ListSubItems.Add , , "Gold"
            If rs!kategori_pelanggan = 5 Then .ListSubItems.Add , , "Platinum"
        
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
rs.Open "select COUNT(ID) from senarai_pelanggan where (nama LIKE'" & Frm28_LM_SEARCH_1 & "' OR no_ic LIKE'" & Frm28_LM_SEARCH_1 & "' OR no_tel LIKE'" & Frm28_LM_SEARCH_1 & "' OR no_pelanggan LIKE'" & Frm28_LM_SEARCH_1 & "') AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    frm28_LM_TOTAL_PAGE = Format(rs(0) / frm28_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, frm28_LM_TOTAL_PAGE, ".") <> 0 Then
    
        frm28_LM_PAGE = Split(frm28_LM_TOTAL_PAGE, ".")(0)
        frm28_LM_PAGE_LEBIHAN = Split(frm28_LM_TOTAL_PAGE, ".")(1)
        
        If frm28_LM_PAGE_LEBIHAN <> "00" Then
            Frm28.L68_Text = frm28_LM_PAGE + 1
        Else
            Frm28.L68_Text = frm28_LM_PAGE
        End If
        
    Else
    
        Frm28.L68_Text = frm28_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm28.L68_Text = 0
    End If
Else
    Frm28.L68_Text = 0
End If

If Not IsNull(rs(0)) Then Frm28.L71_Text = rs(0)

rs.Close
Set rs = Nothing

If x <> 0 Then
    Frm28.L69_Text = LM_START_ROW
End If

If Frm28.L67_Text <> vbNullString And IsNumeric(Frm28.L67_Text) Then
    If Frm28.L68_Text <> vbNullString And IsNumeric(Frm28.L68_Text) Then
        frm28_LM_CURR_PAGE = Frm28.L67_Text
        frm28_LM_TOTAL_PAGE = Frm28.L68_Text
        
        If frm28_LM_CURR_PAGE > frm28_LM_TOTAL_PAGE Then
            
            Frm28.L67_Text = Frm28.L67_Text - 1
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            GoTo re_gen_report:
            
        End If
    End If
End If
End Sub
Sub frm28_periksa_carian()
'on error resume next
'If Frm28.CB2 = 0 And Frm28.CB3 = 0 Then
'    MsgBox "Sila buat pilihan carian mengikut [No. Kad Pengenalan] atau [No. Keahlian].", vbInformation, "Info"
'    Exit Sub
'End If

If Frm28.TB1 = vbNullString Then
    MsgBox "Sila masukkan [Keyword].", vbInformation, "Info"
    Exit Sub
End If

If InStr(1, Frm28.TB1, "'") <> 0 Or InStr(1, Frm28.TB1, "*") <> 0 Or InStr(1, Frm28.TB1, "&") <> 0 Or InStr(1, Frm28.TB1, "-") <> 0 Then

    MsgBox "Keyword mengandungi simbol yang tidak sah.", vbInformation, "Info"
    Exit Sub
End If

Frm28.L72_Text = UCase(Frm28.TB1)

Frm28.L69_Text = -1 'Titik Pencarian Data
Frm28.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Frm28.L67_Text = 0 'Paparan Page ke-xxx
Frm28.L68_Text = 0

GM_NEXT_PREV = 0

Call frm28_senarai_pelanggan_header
Call frm28_senarai_pelanggan
End Sub
Sub frm27_periksa_carian()
'on error resume next
'If frm27.CB2 = 0 And frm27.CB3 = 0 Then
'    MsgBox "Sila buat pilihan carian mengikut [No. Kad Pengenalan] atau [No. Keahlian].", vbInformation, "Info"
'    Exit Sub
'End If

If Frm27.TB1 = vbNullString Then
    MsgBox "Sila masukkan [Keyword].", vbInformation, "Info"
    Exit Sub
End If

If InStr(1, Frm27.TB1, "'") <> 0 Or InStr(1, Frm27.TB1, "*") <> 0 Or InStr(1, Frm27.TB1, "&") <> 0 Or InStr(1, Frm27.TB1, "-") <> 0 Then

    MsgBox "Keyword mengandungi simbol yang tidak sah.", vbInformation, "Info"
    Exit Sub
End If

Frm27.L72_Text = UCase(Frm27.TB1)

Frm27.L69_Text = -1 'Titik Pencarian Data
Frm27.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Frm27.L67_Text = 0 'Paparan Page ke-xxx
Frm27.L68_Text = 0

GM_NEXT_PREV = 0

Call frm27_senarai_dropship_header
Call frm27_senarai_dropship
End Sub
Sub frm27_senarai_dropship_header()
'on error resume next
With Frm27.LV1
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear

    Frm27.LV1.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "Nama", 4600
    .ColumnHeaders.Add 5, , "No. Kad Pengenalan", 2100
    .ColumnHeaders.Add 6, , "No. Telefon", 1600
    .ColumnHeaders.Add 7, , "No. Keahlian", 1700
    .ColumnHeaders.Add 8, , "Kategori", 2000
    
End With
End Sub
Sub frm27_senarai_dropship()
'On Error Resume Next
Dim frm27_LM_TOTAL_PAGE As Double

frm27_PAGE_SIZE = 17
frm27_LM_TOTAL_PAGE = 0
x = 0
frm27_LM_SEARCH_1 = "%" & Frm27.L72_Text & "%"

re_gen_report:

LM_START_ROW = Frm27.L69_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + frm27_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm27.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - frm27_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        Frm27.L67_Text = 1
    End If
End If

Frm27.L71_Text = 0

frm27_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from senarai_pelanggan where dropship = 1 AND (nama LIKE'" & frm27_LM_SEARCH_1 & "' OR no_ic LIKE'" & frm27_LM_SEARCH_1 & "' OR no_tel LIKE'" & frm27_LM_SEARCH_1 & "' OR no_pelanggan LIKE'" & frm27_LM_SEARCH_1 & "') AND status = 1 order by nama ASC LIMIT " & LM_START_ROW & "," & frm27_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1
    If frm27_LM_PAGE_FOUND = 0 Then
        If Frm27.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm27.L67_Text = Frm27.L67_Text + 1 'Paparan Page ke-xxx
                frm27_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm27.L67_Text) Then
                    If Frm27.L67_Text <> 1 Then
                        Frm27.L67_Text = Frm27.L67_Text - 1 'Paparan Page ke-xxx
                        frm27_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If

    Y = ((Frm27.L67_Text - 1) * frm27_PAGE_SIZE) + x

    With Frm27.LV1.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!Nama) Then
            .ListSubItems.Add , , rs!Nama
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_ic) Then
            .ListSubItems.Add , , rs!no_ic
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_tel) Then
            .ListSubItems.Add , , rs!no_tel
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_pelanggan) Then
            .ListSubItems.Add , , rs!no_pelanggan
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kategori_pelanggan) Then 'Kategori
        
            If rs!kategori_pelanggan = 1 Then .ListSubItems.Add , , "Pelanggan Biasa"
            If rs!kategori_pelanggan = 2 Then .ListSubItems.Add , , "Ahli Biasa"
            If rs!kategori_pelanggan = 3 Then .ListSubItems.Add , , "Silver"
            If rs!kategori_pelanggan = 4 Then .ListSubItems.Add , , "Gold"
            If rs!kategori_pelanggan = 5 Then .ListSubItems.Add , , "Platinum"
        
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
rs.Open "select COUNT(ID) from senarai_pelanggan where dropship = 1 AND (nama LIKE'" & frm27_LM_SEARCH_1 & "' OR no_ic LIKE'" & frm27_LM_SEARCH_1 & "' OR no_tel LIKE'" & frm27_LM_SEARCH_1 & "' OR no_pelanggan LIKE'" & frm27_LM_SEARCH_1 & "') AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    frm27_LM_TOTAL_PAGE = Format(rs(0) / frm27_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, frm27_LM_TOTAL_PAGE, ".") <> 0 Then
    
        frm27_LM_PAGE = Split(frm27_LM_TOTAL_PAGE, ".")(0)
        frm27_LM_PAGE_LEBIHAN = Split(frm27_LM_TOTAL_PAGE, ".")(1)
        
        If frm27_LM_PAGE_LEBIHAN <> "00" Then
            Frm27.L68_Text = frm27_LM_PAGE + 1
        Else
            Frm27.L68_Text = frm27_LM_PAGE
        End If
        
    Else
    
        Frm27.L68_Text = frm27_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm27.L68_Text = 0
    End If
Else
    Frm27.L68_Text = 0
End If

If Not IsNull(rs(0)) Then Frm27.L71_Text = rs(0)

rs.Close
Set rs = Nothing

If x <> 0 Then
    Frm27.L69_Text = LM_START_ROW
End If

If Frm27.L67_Text <> vbNullString And IsNumeric(Frm27.L67_Text) Then
    If Frm27.L68_Text <> vbNullString And IsNumeric(Frm27.L68_Text) Then
        frm27_LM_CURR_PAGE = Frm27.L67_Text
        frm27_LM_TOTAL_PAGE = Frm27.L68_Text
        
        If frm27_LM_CURR_PAGE > frm27_LM_TOTAL_PAGE Then
            
            Frm27.L67_Text = Frm27.L67_Text - 1
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            GoTo re_gen_report:
            
        End If
    End If
End If
End Sub
