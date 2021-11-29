Attribute VB_Name = "Module9"
Sub Frm95_on_time_reset()
'on error resume next
'Frm95.L19_Text = 0 'Senarai supplier : Paparan page
'Frm95.L20_Text = 0 'Senarai supplier : Jumlah page
'Frm95.L21_Text = 0 'Senarai supplier : Titik carian data (default = -1)
'Frm95.L22_Text = 0 'Senarai supplier : Flag page terakhir
Frm95.L23_Text = 0 'Senarai supplier : Bilangan tidak aktif
Frm95.L24_Text = 0 'Senarai supplier : Bilangan aktif
Frm95.L25_Text = 0 'Senarai purity : Bilangan tidak aktif
Frm95.L26_Text = 0 'Senarai purity : Bilangan aktif
Frm95.L27_Text = 0 'Senarai senarai produk : Bilangan tidak aktif
Frm95.L28_Text = 0 'Senarai senarai produk : Bilangan aktif
Frm95.L29_Text = 0 'Senarai dulang : Bilangan tidak aktif
Frm95.L30_Text = 0 'SSenarai dulang : Bilangan aktif
End Sub
Sub Frm95_initial()
'on error resume next
Frm95.Frame1.Left = 2850
Frm95.Frame1.Top = 0
Frm95.Frame2.Left = 2850
Frm95.Frame2.Top = 0
Frm95.Frame3.Left = 2850
Frm95.Frame3.Top = 0
Frm95.Frame4.Left = 2850
Frm95.Frame4.Top = 0
Frm95.Frame5.Left = 2850
Frm95.Frame5.Top = 0
Frm95.Frame6.Left = 2850
Frm95.Frame6.Top = 0
Frm95.Frame7.Left = 2850
Frm95.Frame7.Top = 0
Frm95.Frame8.Top = 0
Frm95.Frame8.Left = 2850
Frm95.Pic9.Top = 0
Frm95.Pic9.Left = 2850
Frm95.Pic10.Top = 0
Frm95.Pic10.Left = 2850

Frm95.CB1 = 1
Frm95.CB2 = 0

Frm95.TB1 = vbNullString
Frm95.TB2 = vbNullString
Frm95.TB3 = vbNullString
Frm95.TB4 = vbNullString
Frm95.TB5 = vbNullString
Frm95.TB6 = vbNullString
Frm95.TB7 = vbNullString
Frm95.TB8 = vbNullString
Frm95.TB9 = vbNullString
Frm95.TB10 = vbNullString
Frm95.TB11 = vbNullString
Frm95.TB12 = vbNullString
Frm95.TB13 = vbNullString
Frm95.TB14 = vbNullString
Frm95.TB15 = vbNullString
Frm95.TB16 = "0.00"
Frm95.TB17 = "0.00"
Frm95.TB18 = "0.00"

Frm95.CMD1.Visible = True
Frm95.CMD2.Visible = False
Frm95.CMD3.Visible = False
Frm95.CMD4.Visible = True
Frm95.CMD5.Visible = False
Frm95.CMD6.Visible = False
Frm95.CMD7.Visible = True
Frm95.CMD8.Visible = False
Frm95.CMD9.Visible = False
Frm95.CMD10.Visible = True
Frm95.CMD11.Visible = False
Frm95.CMD12.Visible = False
Frm95.CMD13.Visible = True
Frm95.CMD14.Visible = False
Frm95.CMD15.Visible = False
End Sub
Sub Frm95_invisible()
'on error resume next
Frm95.Frame1.Visible = False
Frm95.Frame2.Visible = False
Frm95.Frame3.Visible = False
Frm95.Frame4.Visible = False
Frm95.Frame5.Visible = False
Frm95.Frame6.Visible = False
Frm95.Frame7.Visible = False
Frm95.Frame8.Visible = False
Frm95.Pic9.Visible = False
Frm95.Pic10.Visible = False
End Sub
Sub Frm95_senarai_supplier_header()
'on error resume next
With Frm95.LV1

    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear
    
    Frm95.LV1.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0
    .ColumnHeaders.Add 4, , "Status", 1300, 2
    .ColumnHeaders.Add 5, , "Jenis", 1700
    .ColumnHeaders.Add 6, , "Nama Supplier", 4500
    .ColumnHeaders.Add 7, , "Kod Supplier", 1700
    .ColumnHeaders.Add 8, , "No. ID GST", 3000
    .ColumnHeaders.Add 9, , "No. Pendaftaran", 2500
    .ColumnHeaders.Add 10, , "No. Telefon (O)", 2500
    .ColumnHeaders.Add 11, , "No. Telefon (HP)", 2500
    .ColumnHeaders.Add 12, , "Alamat", 5000
    .ColumnHeaders.Add 13, , "Nama Bank", 4500
    .ColumnHeaders.Add 14, , "No. Akaun", 2500
    
End With
End Sub
Sub Frm95_senarai_supplier()
'on error resume next
x = 0

Frm95.L23_Text = 0 'Senarai supplier : Bilangan aktif
Frm95.L24_Text = 0 'Senarai supplier : Bilangan tidak aktif

Y = 0 'Bilangan aktif
Z = 0 'Bilangan tidak aktif

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where Supplier <>'" & Null & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1

    With Frm95.LV1.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , x
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!Status) Then 'Status
            
            If rs!Status = 1 Then
                .ListSubItems.Add , , "Aktif"
                Y = Y + 1
            Else
                .ListSubItems.Add , , "Tidak aktif"
                Z = Z + 1
            End If
            
        Else
            .ListSubItems.Add , , "Tidak aktif"
        End If
        
        If Not IsNull(rs!jenis_supplier) Then 'Jenis
            .ListSubItems.Add , , rs!jenis_supplier
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!supplier) Then 'Nama Supplier
            .ListSubItems.Add , , rs!supplier
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Kod_Supplier) Then 'Kod Supplier
            .ListSubItems.Add , , rs!Kod_Supplier
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_id_gst) Then 'No ID GST
            .ListSubItems.Add , , rs!no_id_gst
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!no_pendaftaran) Then 'No. Pendaftaran
            .ListSubItems.Add , , rs!no_pendaftaran
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_tel_off) Then 'No. Telefon (O)
            .ListSubItems.Add , , rs!no_tel_off
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_tel_hp) Then 'No. Telefon (HP)
            .ListSubItems.Add , , rs!no_tel_hp
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!alamat) Then 'Alamat
            .ListSubItems.Add , , rs!alamat
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!nama_bank) Then 'Nama Bank
            .ListSubItems.Add , , rs!nama_bank
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!no_akaun) Then 'No. Akaun
            .ListSubItems.Add , , rs!no_akaun
        Else
            .ListSubItems.Add , , ""
        End If
        
    End With
            
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm95.L23_Text = Y
Frm95.L24_Text = Z

GoTo skip_supplier:

'### Jumlah supplier yang aktif ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(id) from setting_database where Supplier <>'" & Null & "' AND status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm95.L23_Text = rs(0) 'Senarai supplier : Bilangan aktif

rs.Close
Set rs = Nothing
'### Jumlah supplier yang aktif ### - End

'### Jumlah supplier yang tidak aktif ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(id) from setting_database where Supplier <>'" & Null & "' AND status='" & 0 & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm95.L24_Text = rs(0) 'Senarai supplier : Bilangan tidak aktif

rs.Close
Set rs = Nothing
'### Jumlah supplier yang tidak aktif ### - End

skip_supplier:

If x <> 0 Then
    Frm95.Frame2.Visible = True
    Frm95.Frame1.Visible = False
Else
    MsgBox "Tiada senarai supplier dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm95_senarai_purity_header()
'on error resume next
With Frm95.LV2

    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear
    
    Frm95.LV2.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0
    .ColumnHeaders.Add 4, , "Status", 1550, 2
    .ColumnHeaders.Add 5, , "Nama Purity", 3100
    .ColumnHeaders.Add 6, , "Kod Purity", 1700
    .ColumnHeaders.Add 7, , "Assay", 1700, 1
    .ColumnHeaders.Add 8, , "Kadar Trade In", 2500, 1
    
End With
End Sub
Sub Frm95_senarai_purity()
'on error resume next
x = 0
Frm95.L25_Text = 0 'Senarai purity : Bilangan tidak aktif
Frm95.L26_Text = 0 'Senarai purity : Bilangan aktif

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where Metal_Purity <>'" & Null & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1

    With Frm95.LV2.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , x
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!Status) Then 'Status
            
            If rs!Status = 1 Then
                .ListSubItems.Add , , "Aktif"
                Y = Y + 1
            Else
                .ListSubItems.Add , , "Tidak aktif"
                Z = Z + 1
            End If
            
        Else
            .ListSubItems.Add , , "Tidak aktif"
        End If
        
        If Not IsNull(rs!Metal_Purity) Then 'Nama Purity
            .ListSubItems.Add , , rs!Metal_Purity
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Kod_Metal_Purity) Then 'Kod Purity
            .ListSubItems.Add , , rs!Kod_Metal_Purity
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!assay) Then 'Assay
            .ListSubItems.Add , , rs!assay
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!trade_in) Then 'Kadar Trade In
            .ListSubItems.Add , , rs!trade_in
        Else
            .ListSubItems.Add , , ""
        End If
        
    End With
        
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah purity yang aktif ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(id) from setting_database where Metal_Purity <>'" & Null & "' AND status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm95.L25_Text = rs(0) 'Senarai purity : Bilangan aktif

rs.Close
Set rs = Nothing
'### Jumlah purity yang aktif ### - End

'### Jumlah purity yang tidak aktif ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(id) from setting_database where Metal_Purity <>'" & Null & "' AND status='" & 0 & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm95.L26_Text = rs(0) 'Senarai purity : Bilangan tidak aktif

rs.Close
Set rs = Nothing
'### Jumlah purity yang tidak aktif ### - End

If x <> 0 Then
    Frm95.Frame4.Visible = True
    Frm95.Frame3.Visible = False
Else
    MsgBox "Tiada Senarai Purity Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm95_senarai_produk_header()
'on error resume next
With Frm95.LV3

    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear
    
    Frm95.LV3.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0
    .ColumnHeaders.Add 4, , "Status", 1550, 2
    .ColumnHeaders.Add 5, , "Nama Produk", 5600
    .ColumnHeaders.Add 6, , "Kod Produk", 3400
    
End With
End Sub
Sub Frm95_senarai_produk()
'on error resume next
x = 0

Frm95.L27_Text = 0 'Senarai senarai produk : Bilangan tidak aktif
Frm95.L28_Text = 0 'Senarai senarai produk : Bilangan aktif

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where kategori_Produk <>'" & Null & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1

    With Frm95.LV3.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , x
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!Status) Then 'Status
            
            If rs!Status = 1 Then
                .ListSubItems.Add , , "Aktif"
                Y = Y + 1
            Else
                .ListSubItems.Add , , "Tidak aktif"
                Z = Z + 1
            End If
            
        Else
            .ListSubItems.Add , , "Tidak aktif"
        End If
        
        If Not IsNull(rs!kategori_Produk) Then 'Nama Produk
            .ListSubItems.Add , , rs!kategori_Produk
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Kod_Kategori_Produk) Then 'Kod Produk
            .ListSubItems.Add , , rs!Kod_Kategori_Produk
        Else
            .ListSubItems.Add , , ""
        End If
        
    End With
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah senarai produk yang aktif ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(id) from setting_database where kategori_Produk <>'" & Null & "' AND status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm95.L27_Text = rs(0) 'Senarai senarai produk : Bilangan aktif

rs.Close
Set rs = Nothing
'### Jumlah senarai produk yang aktif ### - End

'### Jumlah senarai produk yang tidak aktif ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(id) from setting_database where kategori_Produk <>'" & Null & "' AND status='" & 0 & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm95.L28_Text = rs(0) 'Senarai senarai produk : Bilangan tidak aktif

rs.Close
Set rs = Nothing
'### Jumlah senarai produk yang tidak aktif ### - End

If x <> 0 Then
    Frm95.Frame6.Visible = True
    Frm95.Frame5.Visible = False
Else
    MsgBox "Tiada Senarai Produk Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm95_senarai_dulang_header()
'on error resume next
With Frm95.LV4

    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear
    
    Frm95.LV4.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0
    .ColumnHeaders.Add 4, , "Status", 1550, 2
    .ColumnHeaders.Add 5, , "Nama Dulang", 2900
    
End With
End Sub
Sub Frm95_senarai_dulang()
'on error resume next
x = 0
Frm95.L29_Text = 0 'Senarai dulang : Bilangan tidak aktif
Frm95.L30_Text = 0 'Senarai dulang : Bilangan aktif

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where SenaraiDulang <>'" & Null & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1

    With Frm95.LV4.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , x
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!Status) Then 'Status
            
            If rs!Status = 1 Then
                .ListSubItems.Add , , "Aktif"
                Y = Y + 1
            Else
                .ListSubItems.Add , , "Tidak aktif"
                Z = Z + 1
            End If
            
        Else
            .ListSubItems.Add , , "Tidak aktif"
        End If
        
        If Not IsNull(rs!SenaraiDulang) Then 'Nama Dulang
            .ListSubItems.Add , , rs!SenaraiDulang
        Else
            .ListSubItems.Add , , ""
        End If
        
    End With

    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah dulang yang aktif ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(id) from setting_database where SenaraiDulang <>'" & Null & "' AND status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm95.L29_Text = rs(0) 'Senarai dulang : Bilangan aktif

rs.Close
Set rs = Nothing
'### Jumlah dulang yang aktif ### - End

'### Jumlah dulang yang tidak aktif ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(id) from setting_database where SenaraiDulang <>'" & Null & "' AND status='" & 0 & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm95.L30_Text = rs(0) 'Senarai dulang : Bilangan tidak aktif

rs.Close
Set rs = Nothing
'### Jumlah dulang yang tidak aktif ### - End

If x <> 0 Then
    Frm95.Frame8.Visible = True
    Frm95.Frame7.Visible = False
Else
    MsgBox "Tiada Senarai Dulang Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm95_senarai_tukang_header()
'on error resume next
Frm95.MSFlexGrid5.Clear
Frm95.MSFlexGrid5.Rows = 1
Frm95.MSFlexGrid5.RowHeight(0) = 1500
Frm95.MSFlexGrid5.FormatString = "<No.|<No.|<No. ID|<Nama Tukang Emas"

Frm95.MSFlexGrid5.ColWidth(0) = 600 'No.
Frm95.MSFlexGrid5.ColWidth(1) = 0 'No.
Frm95.MSFlexGrid5.ColWidth(2) = 0 'No. ID
Frm95.MSFlexGrid5.ColWidth(3) = 6400 'Nama Tukang Emas
End Sub
Sub Frm95_senarai_tukang()
'on error resume next
x = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!tukang_emas) Then
        x = x + 1
        Frm95.MSFlexGrid5.Rows = x + 1
        Frm95.MSFlexGrid5.TextMatrix(x, 0) = x 'No.
        Frm95.MSFlexGrid5.TextMatrix(x, 1) = x 'No.
        Frm95.MSFlexGrid5.TextMatrix(x, 2) = rs!ID 'No. ID
        If Not IsNull(rs!tukang_emas) Then Frm95.MSFlexGrid5.TextMatrix(x, 3) = rs!tukang_emas 'Nama Tukang Emas
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

If x <> 0 Then
    Frm95.Pic10.Visible = True
    Frm95.Pic9.Visible = False
Else
    MsgBox "Tiada Senarai Tukang Emas Dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm99_initial_setting()
'on error resume next
Frm99.Pic1.Left = 120
Frm99.Pic1.Top = 2400
Frm99.Pic2.Left = 120
Frm99.Pic2.Top = 2400

Frm99.Pic1.Visible = False
Frm99.Pic2.Visible = False
End Sub
Sub Frm99_senarai_komisyen_header()
'on error resume next
Frm99.MSFlexGrid1.Clear
Frm99.MSFlexGrid1.Rows = 1
Frm99.MSFlexGrid1.RowHeight(0) = 1500
Frm99.MSFlexGrid1.FormatString = "<No.|<No.|<No. ID|<Tarikh|<No. Invoice|<No. Siri Produk|<Nama Barang|<Harga Jualan (RM)|<Harga Staff (RM)|<Komisyen (RM)"

Frm99.MSFlexGrid1.ColWidth(0) = 600 'No.
Frm99.MSFlexGrid1.ColWidth(1) = 0 'No.
Frm99.MSFlexGrid1.ColWidth(2) = 0 'No. ID
Frm99.MSFlexGrid1.ColWidth(3) = 1500 'Tarikh
Frm99.MSFlexGrid1.ColWidth(4) = 1500 'No. Invoice
Frm99.MSFlexGrid1.ColWidth(5) = 1500 'No. Siri Produk
Frm99.MSFlexGrid1.ColWidth(6) = 4000 'Nama Barang
Frm99.MSFlexGrid1.ColWidth(7) = 1500 'Harga Jualan (RM)
Frm99.MSFlexGrid1.ColWidth(8) = 1500 'Harga Staff (RM)
Frm99.MSFlexGrid1.ColWidth(9) = 1500 'Komisyen (RM)
End Sub
Sub Frm99_senarai_komisyen()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim Frm99_LM_KOMISYEN As Double

Frm99_LM_NO_STAFF = vbNullString
Frm99_LM_KOMISYEN = 0

TM = Frm99.DTPicker1 'Tarikh Mula
TA = Frm99.DTPicker2 'Tarikh Akhir
Frm99_LM_NO_STAFF = Split(Frm99.CBB1, " -> ")(0)
Frm99_LM_NAMA_STAFF = Split(Frm99.CBB1, " -> ")(1)

If Frm99_LM_NO_STAFF <> vbNullString Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 23_senarai_jualan where no_pekerja='" & Frm99_LM_NO_STAFF & "' AND tarikh BETWEEN '" & TM & "'  AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        Y = Y + 1
        Frm99.MSFlexGrid1.Rows = Y + 1
        Frm99.MSFlexGrid1.TextMatrix(Y, 0) = Y
        Frm99.MSFlexGrid1.TextMatrix(Y, 1) = Y
        If Not IsNull(rs!ID) Then Frm99.MSFlexGrid1.TextMatrix(Y, 2) = rs!ID 'No ID
        If Not IsNull(rs!tarikh) Then Frm99.MSFlexGrid1.TextMatrix(Y, 3) = rs!tarikh 'Tarikh
        If Not IsNull(rs!no_resit) Then Frm99.MSFlexGrid1.TextMatrix(Y, 4) = rs!no_resit 'No. Invoice
        If Not IsNull(rs!no_siri_Produk) Then Frm99.MSFlexGrid1.TextMatrix(Y, 5) = rs!no_siri_Produk 'No. Siri Produk
        If Not IsNull(rs!kategori_Produk) Then Frm99.MSFlexGrid1.TextMatrix(Y, 6) = rs!kategori_Produk 'Nama Produk
        If Not IsNull(rs!harga_dengan_gst) Then Frm99.MSFlexGrid1.TextMatrix(Y, 7) = Format(rs!harga_dengan_gst, "0.00") 'Harga Jualan (RM)
        If Not IsNull(rs!harga_staff) Then Frm99.MSFlexGrid1.TextMatrix(Y, 8) = Format(rs!harga_staff, "0.00") 'Harga Staff (RM)
        If Not IsNull(rs!komisyen_staff) Then
            Frm99.MSFlexGrid1.TextMatrix(Y, 9) = Format(rs!komisyen_staff, "0.00") 'Jumlah Komisyen Staff Staff (RM)
            'If IsNumeric(rs!komisyen_staff) Then Frm99_LM_KOMISYEN = Frm99_LM_KOMISYEN + rs!komisyen_staff
        End If
        rs.MoveNext
    Wend
        
    rs.Close
    Set rs = Nothing
    
'#### Jumlah Komisyen Staff #### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select SUM(komisyen_staff) from 23_senarai_jualan where no_pekerja='" & Frm99_LM_NO_STAFF & "' AND tarikh BETWEEN '" & TM & "'  AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        Frm99.L7_Text = Format(rs(0), "#,##0.00")
        If rs(0) = vbNullString Then
            Frm99.L7_Text = "0.00"
        End If
    Else
        Frm99.L7_Text = "0.00"
    End If
    
    rs.Close
    Set rs = Nothing
'#### Jumlah Komisyen Staff #### - End
    
    Frm99.L5_Text = "Senarai jualan dan komisyen bagi pekerja " & Frm99_LM_NAMA_STAFF & " dari " & TM & " hingga " & TA & "."
    Frm99.L6_Text = Y
    If Frm99.L7_Text = vbNullString Then Frm99.L7_Text = "0.00"
    
    If Y <> 0 Then
        Frm99.Pic1.Visible = False
        Frm99.Pic2.Visible = True
    Else
        MsgBox "Tiada rekod jualan yang dilakukan oleh pekerja ini di dalam tempoh masa yang telah ditetapkan.", vbInformation, "Info"
    End If
    
End If
End Sub
