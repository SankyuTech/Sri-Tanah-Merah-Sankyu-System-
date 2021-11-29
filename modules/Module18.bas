Attribute VB_Name = "Module18"
Sub Frm57_M_Clear()
'On Error Resume Next
Frm57.Pic1.Left = 120
Frm57.Pic1.Top = 240
Frm57.Pic2.Left = 120
Frm57.Pic2.Top = 240
Frm57.Pic3.Left = 10440
Frm57.Pic3.Top = 240
Frm57.Pic4.Left = 120
Frm57.Pic4.Top = 240
Frm57.Pic5.Left = 120
Frm57.Pic5.Top = 240
Frm57.Pic6.Left = 120
Frm57.Pic6.Top = 240
Frm57.Pic7.Left = 120
Frm57.Pic7.Top = 240

Frm57.Pic1.Visible = False
Frm57.Pic2.Visible = False
Frm57.Pic3.Visible = False
Frm57.Pic4.Visible = False
Frm57.Pic5.Visible = False
Frm57.Pic6.Visible = False
Frm57.Pic7.Visible = False

Frm57.L32_Text = 0
Frm57.L10_Text = "0.00 g"
Frm57.L11_Text = "RM 0.00"
Frm57.L15_Text = 0
Frm57.L13_Text = "0.00 g"
Frm57.L14_Text = "RM 0.00"
Frm57.L19_Text = 0
Frm57.L17_Text = "0.00 g"
Frm57.L18_Text = "RM 0.00"
Frm57.L26_Text = 0

'Frm57.L34_Text.Visible = False
'Frm57.L35_Text.Visible = False
'Frm57.L36_Text.Visible = False
'Frm57.L37_Text.Visible = False

Frm57.L39_Text.Visible = False

Frm57.TB1 = vbNullString
End Sub
Sub Frm57_initial_setting()
'On Error Resume Next
Frm57.Pic1.Visible = False
Frm57.Pic2.Visible = False
Frm57.Pic3.Visible = False
Frm57.Pic4.Visible = False
Frm57.Pic5.Visible = False
Frm57.Pic6.Visible = False
Frm57.Pic7.Visible = False
End Sub
Sub Frm57_M_Inventory()
'On Error Resume Next
Dim TOTALBERAT As Double
Dim TOTALMODAL As Double
Dim rs1 As ADODB.Recordset

x = 0
Frm57.L9_Text = "Report Inventori Dari Dulang [" & Frm57.CBB1 & "]."

'###Padam Table Inventory### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "TRUNCATE TABLE inventory"

Set rs = cn.Execute(strsql)
Set rs = Nothing
'###Padam Table Inventory### - End

'###Padam Table Inventory### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "TRUNCATE TABLE inventory2"

Set rs = cn.Execute(strsql)
Set rs = Nothing
'###Padam Table Inventory### - End


Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

'strsql = "insert into inventory(NoRujukanSistem,tarikh_belian,no_siri,nama_produk,Supplier,purity,Berat,upah,KOSPERGRAM,harga_belian,dulang,panjang,lebar,dia,Size,status)" & _
            "select NoRujukanSistem,tarikh_belian,no_siri_Produk,kategori_Produk,nama_Supplier,kod_Purity,Berat,upah,harga_Per_Gram_Item,harga_item,dulang,dimension_Panjang,dimension_Lebar,dimension_Dia,dimension_Saiz,0 from Data_Database WHERE Dulang='" & Frm57.CBB1 & "' AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "') order by no_siri_Produk ASC"
strsql = "insert into inventory(NoRujukanSistem,tarikh_belian,no_siri,nama_produk,Supplier,purity,Berat,upah,KOSPERGRAM,harga_belian,dulang,panjang,lebar,dia,Size,status)" & _
            "select NoRujukanSistem,tarikh_belian,no_siri_Produk,kategori_Produk,nama_Supplier,kod_Purity,Berat,upah,harga_Per_Gram_Item,harga_item,dulang,dimension_Panjang,dimension_Lebar,dimension_Dia,dimension_Saiz,0 from Data_Database WHERE Dulang='" & Frm57.CBB1 & "' AND (StatusItem='" & "10" & "' OR StatusItem='" & "10" & "') order by no_siri_Produk ASC"


Set rs = cn.Execute(strsql)
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from Data_Database where Dulang='" & Frm57.CBB1 & "' AND (StatusItem='" & "10" & "' OR StatusItem='" & "25" & "')", cn, adOpenKeyset, adLockOptimistic
rs.Open "select * from Data_Database where Dulang='" & Frm57.CBB1 & "' AND (StatusItem='" & "10" & "' OR StatusItem='" & "10" & "')", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    x = x + 1
End If

rs.Close
Set rs = Nothing

If x = 0 Then
    MsgBox "Tiada barang dijumpai dari dulang ini.", vbInformation, "Info"
Else
    Frm57.L6_Text = Frm57.CBB1
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If rs!Default1 = "Default" Then
            If rs!ScannerMode = 1 Then
                Frm57.CB1 = 1
            Else
                Frm57.CB1 = 0
            End If
        End If
    End If
    
    rs.Close
    Set rs = Nothing

    Frm57.Pic3.Visible = True
    Frm57.L39_Text.Visible = True
    
    Frm57.L3_Text = "Sila scan setiap item dari dulang [" & Frm57.CBB1 & "]."
    MsgBox "Data telah dijumpai. Sila scan setiap item kedai dari dulang [" & Frm57.CBB1 & "].", vbInformation, "Info"
    Frm57.TB1.SetFocus
End If
End Sub
Sub Frm57_M_Carian()
'On Error Resume Next
DATA_FOUND = 0
DATA_FOUND2 = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from inventory where no_siri='" & Frm57.TB1 & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    rs!Status = 1 '0 : Belum Scan , 1 : Sudah Scan
    rs!write_timestamp = Now
    rs.Update
    DATA_FOUND = 1
End If

rs.Close
Set rs = Nothing

If DATA_FOUND = 0 Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from Data_Database where no_siri_Produk='" & Frm57.TB1 & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If rs!StatusItem <> "10" Then
            If rs!StatusItem <> "0" Then
                Frm57_LM_DETAIL = "Item ini tiada dalam stok kedai lagi. Kemungkinan telah terjual."
            Else
                Frm57_LM_DETAIL = "Item ini sudah dipadamkan dari sistem."
            End If
        End If
        If rs!StatusItem = "10" And rs!dulang <> Frm57.TB1 Then
            Frm57_LM_DETAIL = "Item ini dari dulang [" & rs!dulang & "]."
        End If
        DATA_FOUND2 = 1
        DATA_FOUND = 1
    End If
    
    rs.Close
    Set rs = Nothing
    
End If

If DATA_FOUND2 = 1 Then
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from inventory2 where no_siri='" & Frm57.TB1 & "'", cn, adOpenKeyset, adLockOptimistic
    
    If rs.EOF Then
        rs.AddNew
        rs!no_siri = Frm57.TB1 'No. Siri Produk
        rs!Detail = Frm57_LM_DETAIL 'Detail
        rs!write_timestamp = Now
        rs.Update
    End If
    
    rs.Close
    Set rs = Nothing
End If

If DATA_FOUND = 0 Then
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from inventory2 where no_siri='" & Frm57.TB1 & "'", cn, adOpenKeyset, adLockOptimistic
    
    If rs.EOF Then
        rs.AddNew
        rs!no_siri = Frm57.TB1 'No. Siri Produk
        rs!Detail = "Tiada data bagi item ini." 'Detail
        rs!write_timestamp = Now
        rs.Update
    End If
    
    rs.Close
    Set rs = Nothing
End If

Frm57.L4_Text = "No. Siri Telah Dicari : " & Frm57.TB1
Frm57.L5_Text = "Status Carian   : Selesai"

Frm57.TB1 = vbNullString
Frm57.TB1.SetFocus
End Sub
Sub Frm57_M_RekodInventori_Stok_header()
'////////////Report Invertori Dulang Keseluruhan////////////////
Frm57.MSFlexGrid1.Clear
Frm57.MSFlexGrid1.Rows = 1
Frm57.MSFlexGrid1.RowHeight(0) = 700
Frm57.MSFlexGrid1.FormatString = "<No.|<No.|<No. Rujukan|<Tarikh Belian|<No. Siri Produk|<Nama Produk|<Supplier|<Purity|<Berat (g)|<Kos Belian Per Gram (RM/g)|<Upah (RM)|<Harga Belian (RM)|<Dulang|<Panjang|<Lebar|<Dia|<Saiz"

Frm57.MSFlexGrid1.ColWidth(0) = 0
Frm57.MSFlexGrid1.ColWidth(1) = 600
Frm57.MSFlexGrid1.ColWidth(2) = 0
Frm57.MSFlexGrid1.ColWidth(3) = 1200
Frm57.MSFlexGrid1.ColWidth(4) = 1500
Frm57.MSFlexGrid1.ColWidth(5) = 3200
Frm57.MSFlexGrid1.ColWidth(6) = 3200
Frm57.MSFlexGrid1.ColWidth(7) = 1200
Frm57.MSFlexGrid1.ColWidth(8) = 1200
Frm57.MSFlexGrid1.ColWidth(9) = 1200
Frm57.MSFlexGrid1.ColWidth(10) = 1200
Frm57.MSFlexGrid1.ColWidth(11) = 1200
Frm57.MSFlexGrid1.ColWidth(12) = 1200
Frm57.MSFlexGrid1.ColWidth(13) = 1000
Frm57.MSFlexGrid1.ColWidth(14) = 1000
Frm57.MSFlexGrid1.ColWidth(15) = 1000
Frm57.MSFlexGrid1.ColWidth(16) = 1000
End Sub
Sub Frm57_M_RekodInventori_Stok()
'On Error Resume Next
Dim TOTALBERAT As Double
Dim TOTALMODAL As Double
Dim Frm57_LM_TOTAL_PAGE As Double
Dim Frm57_LM_JUMLAH As Double

Frm57_PAGE_SIZE = 38
Frm57_LM_TOTAL_PAGE = 0

x = 0
TOTALBERAT = 0
TOTALMODAL = 0

LM_START_ROW = Frm57.L42_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm57_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm57.L43_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm57_PAGE_SIZE
        End If
    End If
End If

Frm57_LM_PAGE_FOUND = 0

Frm57.L9_Text = "Report inventori bagi dulang [" & Frm57.L6_Text & "] (Ini adalah senarai semua barang yang patut berada di dalam dulang ini)."

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from inventory LIMIT " & LM_START_ROW & "," & Frm57_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm57_LM_PAGE_FOUND = 0 Then
        If Frm57.L43_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm57.L40_Text = Frm57.L40_Text + 1 'Paparan Page ke-xxx
                Frm57_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm57.L40_Text) Then
                    If Frm57.L40_Text <> 1 Then
                        Frm57.L40_Text = Frm57.L40_Text - 1 'Paparan Page ke-xxx
                        Frm57_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm57.L40_Text - 1) * Frm57_PAGE_SIZE) + x
    
    Frm57.MSFlexGrid1.Rows = x + 1
    Frm57.MSFlexGrid1.TextMatrix(x, 0) = x
    Frm57.MSFlexGrid1.TextMatrix(x, 1) = Y
    If Not IsNull(rs!NoRujukanSistem) Then Frm57.MSFlexGrid1.TextMatrix(x, 2) = rs!NoRujukanSistem 'No. Rujukan Sistem
    If Not IsNull(rs!tarikh_belian) Then Frm57.MSFlexGrid1.TextMatrix(x, 3) = rs!tarikh_belian 'Tarikh Belian
    If Not IsNull(rs!no_siri) Then Frm57.MSFlexGrid1.TextMatrix(x, 4) = rs!no_siri   'No. Siri Produk
    If Not IsNull(rs!nama_produk) Then Frm57.MSFlexGrid1.TextMatrix(x, 5) = rs!nama_produk   'Nama Produk
    If Not IsNull(rs!supplier) Then Frm57.MSFlexGrid1.TextMatrix(x, 6) = rs!supplier   'Nama Supplier
    If Not IsNull(rs!purity) Then Frm57.MSFlexGrid1.TextMatrix(x, 7) = rs!purity   'Purity
    If IsNumeric(rs!Berat) Then
        Frm57.MSFlexGrid1.TextMatrix(x, 8) = Format(rs!Berat, "#,##0.00")  'Berat Jualan
    Else
        Frm57.MSFlexGrid1.TextMatrix(x, 8) = "-"
    End If
    If IsNumeric(rs!KOSPERGRAM) Then
        Frm57.MSFlexGrid1.TextMatrix(x, 9) = Format(rs!KOSPERGRAM, "#,##0.00")  'Harga Belian Per Gram
    Else
        Frm57.MSFlexGrid1.TextMatrix(x, 9) = "-"
    End If
    If IsNumeric(rs!UPAH) Then
        Frm57.MSFlexGrid1.TextMatrix(x, 10) = Format(rs!UPAH, "#,##0.00") 'Upah
    Else
        Frm57.MSFlexGrid1.TextMatrix(x, 10) = "-"
    End If
    If IsNumeric(rs!harga_belian) Then
        Frm57.MSFlexGrid1.TextMatrix(x, 11) = Format(rs!harga_belian, "#,##0.00")  'Kos Belian Item
    Else
        Frm57.MSFlexGrid1.TextMatrix(x, 11) = "-"
    End If

    If Not IsNull(rs!dulang) Then Frm57.MSFlexGrid1.TextMatrix(x, 12) = rs!dulang 'Dulang
    If Not IsNull(rs!panjang) Then Frm57.MSFlexGrid1.TextMatrix(x, 13) = rs!panjang   'Panjang
    If Not IsNull(rs!lebar) Then Frm57.MSFlexGrid1.TextMatrix(x, 14) = rs!lebar   'Lebar
    If Not IsNull(rs!dia) Then Frm57.MSFlexGrid1.TextMatrix(x, 15) = rs!dia   'Dia
    If Not IsNull(rs!Size) Then Frm57.MSFlexGrid1.TextMatrix(x, 16) = rs!Size   'Saiz
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from inventory", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm57_LM_TOTAL_PAGE = Format(rs(0) / Frm57_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm57_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm57_LM_PAGE = Split(Frm57_LM_TOTAL_PAGE, ".")(0)
        Frm57_LM_PAGE_LEBIHAN = Split(Frm57_LM_TOTAL_PAGE, ".")(1)
        
        If Frm57_LM_PAGE_LEBIHAN <> "00" Then
            Frm57.L41_Text = Frm57_LM_PAGE + 1
        Else
            Frm57.L41_Text = Frm57_LM_PAGE
        End If
        
    Else
    
        Frm57.L41_Text = Frm57_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm57.L41_Text = 0
    End If
Else
    Frm57.L41_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm57.L41_Text = vbNullString Then
    Frm57.L41_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah bilangan barang ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from inventory", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm57.L32_Text = rs(0)

rs.Close
Set rs = Nothing
'### Jumlah bilangan barang ### - End

'### Jumlah berat ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(Berat) from inventory", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm57.L10_Text = Format(rs(0), "#,##0.00 g") 'Jumlah (g)

rs.Close
Set rs = Nothing
'### Jumlah berat ### - End

'### Jumlah harga modal ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_belian) from inventory", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm57.L11_Text = "RM " & Format(rs(0), "#,##0.00") 'Jumlah (RM)

rs.Close
Set rs = Nothing
'### Jumlah harga modal ### - End

If x <> 0 Then
    Frm57.L42_Text = LM_START_ROW 'Titik Pencarian Data
    Frm57.Pic2.Visible = True
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    'If Frm57_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm57.L43_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm57.L43_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If

'If X = 0 Then
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
'Else
'    Call Frm57_M_Inventori_Kawalan
'    Call Frm57_M_Inventori_Luar_Kawalan
'    Call Frm57_M_Inventori_Lain
    
    Frm57.L34_Text.Visible = True
    Frm57.L35_Text.Visible = True
    Frm57.L36_Text.Visible = True
    Frm57.L37_Text.Visible = True

'    Frm57.Pic1.Visible = False
'    Frm57.Pic2.Visible = True
'    Frm57.Pic3.Visible = False
'End If
End Sub
Sub Frm57_M_Inventori_Kawalan_header()
'On Error Resume Next
'////////////Report Inventori Yang Telah Discan////////////////
Frm57.MSFlexGrid2.Clear
Frm57.MSFlexGrid2.Rows = 1
Frm57.MSFlexGrid2.RowHeight(0) = 700
Frm57.MSFlexGrid2.FormatString = "<No.|<No.|<No. Rujukan|<Tarikh Belian|<No. Siri Produk|<Nama Produk|<Supplier|<Purity|<Berat (g)|<Kos Belian Per Gram (RM/g)|<Upah (RM)|<Harga Belian (RM)|<Dulang|<Panjang|<Lebar|<Dia|<Saiz"

Frm57.MSFlexGrid2.ColWidth(0) = 0
Frm57.MSFlexGrid2.ColWidth(1) = 600
Frm57.MSFlexGrid2.ColWidth(2) = 0
Frm57.MSFlexGrid2.ColWidth(3) = 1200
Frm57.MSFlexGrid2.ColWidth(4) = 1500
Frm57.MSFlexGrid2.ColWidth(5) = 3200
Frm57.MSFlexGrid2.ColWidth(6) = 3200
Frm57.MSFlexGrid2.ColWidth(7) = 1200
Frm57.MSFlexGrid2.ColWidth(8) = 1200
Frm57.MSFlexGrid2.ColWidth(9) = 1200
Frm57.MSFlexGrid2.ColWidth(10) = 1200
Frm57.MSFlexGrid2.ColWidth(11) = 1200
Frm57.MSFlexGrid2.ColWidth(12) = 1200
Frm57.MSFlexGrid2.ColWidth(13) = 1000
Frm57.MSFlexGrid2.ColWidth(14) = 1000
Frm57.MSFlexGrid2.ColWidth(15) = 1000
Frm57.MSFlexGrid2.ColWidth(16) = 1000
End Sub
Sub Frm57_M_Inventori_Kawalan()
'On Error Resume Next
Dim TOTALBERAT As Double
Dim TOTALMODAL As Double
Dim Frm57_LM_TOTAL_PAGE As Double
Dim Frm57_LM_JUMLAH As Double

Frm57_PAGE_SIZE = 38
Frm57_LM_TOTAL_PAGE = 0

x = 0
TOTALBERAT = 0
TOTALMODAL = 0

LM_START_ROW = Frm57.L46_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm57_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm57.L47_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm57_PAGE_SIZE
        End If
    End If
End If

Frm57_LM_PAGE_FOUND = 0

Frm57.L12_Text = "Report inventori bagi dulang [" & Frm57.L6_Text & "] (Ini adalah senarai semua barang yang telah berjaya dijumpai dari dulang ini)."

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from inventory where status='" & "1" & "' LIMIT " & LM_START_ROW & "," & Frm57_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm57_LM_PAGE_FOUND = 0 Then
        If Frm57.L47_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm57.L44_Text = Frm57.L44_Text + 1 'Paparan Page ke-xxx
                Frm57_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm57.L44_Text) Then
                    If Frm57.L44_Text <> 1 Then
                        Frm57.L44_Text = Frm57.L44_Text - 1 'Paparan Page ke-xxx
                        Frm57_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm57.L44_Text - 1) * Frm57_PAGE_SIZE) + x
    
    Frm57.MSFlexGrid2.Rows = x + 1
    Frm57.MSFlexGrid2.TextMatrix(x, 0) = x
    Frm57.MSFlexGrid2.TextMatrix(x, 1) = Y
    If Not IsNull(rs!NoRujukanSistem) Then Frm57.MSFlexGrid2.TextMatrix(x, 2) = rs!NoRujukanSistem 'No. Rujukan Sistem
    If Not IsNull(rs!tarikh_belian) Then Frm57.MSFlexGrid2.TextMatrix(x, 3) = rs!tarikh_belian 'Tarikh Belian
    If Not IsNull(rs!no_siri) Then Frm57.MSFlexGrid2.TextMatrix(x, 4) = rs!no_siri   'No. Siri Produk
    If Not IsNull(rs!nama_produk) Then Frm57.MSFlexGrid2.TextMatrix(x, 5) = rs!nama_produk   'Nama Produk
    If Not IsNull(rs!supplier) Then Frm57.MSFlexGrid2.TextMatrix(x, 6) = rs!supplier   'Nama Supplier
    If Not IsNull(rs!purity) Then Frm57.MSFlexGrid2.TextMatrix(x, 7) = rs!purity   'Purity
    If IsNumeric(rs!Berat) Then
        Frm57.MSFlexGrid2.TextMatrix(x, 8) = Format(rs!Berat, "#,##0.00")  'Berat Jualan
        TOTALBERAT = TOTALBERAT + rs!Berat
    Else
        Frm57.MSFlexGrid2.TextMatrix(x, 8) = "-"
    End If
    If IsNumeric(rs!KOSPERGRAM) Then
        Frm57.MSFlexGrid2.TextMatrix(x, 9) = Format(rs!KOSPERGRAM, "#,##0.00")  'Harga Belian Per Gram
    Else
        Frm57.MSFlexGrid2.TextMatrix(x, 9) = "-"
    End If
    
    If IsNumeric(rs!UPAH) Then
        Frm57.MSFlexGrid2.TextMatrix(x, 10) = Format(rs!UPAH, "#,##0.00") 'Upah
    Else
        Frm57.MSFlexGrid2.TextMatrix(x, 10) = "-"
    End If
    If IsNumeric(rs!harga_belian) Then
        Frm57.MSFlexGrid2.TextMatrix(x, 11) = Format(rs!harga_belian, "#,##0.00")  'Kos Belian Item
        TOTALMODAL = TOTALMODAL + rs!harga_belian
    Else
        Frm57.MSFlexGrid2.TextMatrix(x, 11) = "-"
    End If
    If Not IsNull(rs!dulang) Then Frm57.MSFlexGrid2.TextMatrix(x, 12) = rs!dulang 'Dulang
    If Not IsNull(rs!panjang) Then Frm57.MSFlexGrid2.TextMatrix(x, 13) = rs!panjang   'Panjang
    If Not IsNull(rs!lebar) Then Frm57.MSFlexGrid2.TextMatrix(x, 14) = rs!lebar   'Lebar
    If Not IsNull(rs!dia) Then Frm57.MSFlexGrid2.TextMatrix(x, 15) = rs!dia   'Dia
    If Not IsNull(rs!Size) Then Frm57.MSFlexGrid2.TextMatrix(x, 16) = rs!Size   'Saiz
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from inventory where status='" & "1" & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm57_LM_TOTAL_PAGE = Format(rs(0) / Frm57_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm57_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm57_LM_PAGE = Split(Frm57_LM_TOTAL_PAGE, ".")(0)
        Frm57_LM_PAGE_LEBIHAN = Split(Frm57_LM_TOTAL_PAGE, ".")(1)
        
        If Frm57_LM_PAGE_LEBIHAN <> "00" Then
            Frm57.L45_Text = Frm57_LM_PAGE + 1
        Else
            Frm57.L45_Text = Frm57_LM_PAGE
        End If
        
    Else
    
        Frm57.L45_Text = Frm57_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm57.L45_Text = 0
    End If
Else
    Frm57.L45_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm57.L45_Text = vbNullString Then
    Frm57.L45_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah bilangan barang ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from inventory where status='" & "1" & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm57.L15_Text = rs(0)

rs.Close
Set rs = Nothing
'### Jumlah bilangan barang ### - End

'### Jumlah berat ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(Berat) from inventory where status='" & "1" & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm57.L13_Text = Format(rs(0), "#,##0.00 g") 'Jumlah (g)

rs.Close
Set rs = Nothing
'### Jumlah berat ### - End

'### Jumlah harga modal ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_belian) from inventory where status='" & "1" & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm57.L14_Text = "RM " & Format(rs(0), "#,##0.00") 'Jumlah (RM)

rs.Close
Set rs = Nothing
'### Jumlah harga modal ### - End

If x <> 0 Then
    Frm57.L46_Text = LM_START_ROW 'Titik Pencarian Data
    Frm57.Pic4.Visible = True
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    'If Frm57_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm57.L47_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm57.L47_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub Frm57_M_Inventori_Luar_Kawalan_header()
'On Error Resume Next
'////////////Report Inventori Yang Tidak Dijumpai////////////////
Frm57.MSFlexGrid3.Clear
Frm57.MSFlexGrid3.Rows = 1
Frm57.MSFlexGrid3.RowHeight(0) = 700
Frm57.MSFlexGrid3.FormatString = "<No.|<No.|<No. Rujukan|<Tarikh Belian|<No. Siri Produk|<Nama Produk|<Supplier|<Purity|<Berat (g)|<Kos Belian Per Gram (RM/g)|<Upah (RM)|<Harga Belian (RM)|<Dulang|<Panjang|<Lebar|<Dia|<Saiz"

Frm57.MSFlexGrid3.ColWidth(0) = 0
Frm57.MSFlexGrid3.ColWidth(1) = 600
Frm57.MSFlexGrid3.ColWidth(2) = 0
Frm57.MSFlexGrid3.ColWidth(3) = 1200
Frm57.MSFlexGrid3.ColWidth(4) = 1500
Frm57.MSFlexGrid3.ColWidth(5) = 3200
Frm57.MSFlexGrid3.ColWidth(6) = 3200
Frm57.MSFlexGrid3.ColWidth(7) = 1200
Frm57.MSFlexGrid3.ColWidth(8) = 1200
Frm57.MSFlexGrid3.ColWidth(9) = 1200
Frm57.MSFlexGrid3.ColWidth(10) = 1200
Frm57.MSFlexGrid3.ColWidth(11) = 1200
Frm57.MSFlexGrid3.ColWidth(12) = 1200
Frm57.MSFlexGrid3.ColWidth(13) = 1000
Frm57.MSFlexGrid3.ColWidth(14) = 1000
Frm57.MSFlexGrid3.ColWidth(15) = 1000
Frm57.MSFlexGrid3.ColWidth(16) = 1000
End Sub
Sub Frm57_M_Inventori_Luar_Kawalan()
'On Error Resume Next
Dim TOTALBERAT As Double
Dim TOTALMODAL As Double
Dim Frm57_LM_TOTAL_PAGE As Double
Dim Frm57_LM_JUMLAH As Double

Frm57_PAGE_SIZE = 38
Frm57_LM_TOTAL_PAGE = 0

x = 0
TOTALBERAT = 0
TOTALMODAL = 0

LM_START_ROW = Frm57.L50_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm57_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm57.L51_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm57_PAGE_SIZE
        End If
    End If
End If

Frm57_LM_PAGE_FOUND = 0

Frm57.L16_Text = "Report inventori bagi dulang [" & Frm57.L6_Text & "] (Ini adalah senarai barang yang TIADA dalam dulang)."

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from inventory where status='" & "0" & "' LIMIT " & LM_START_ROW & "," & Frm57_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm57_LM_PAGE_FOUND = 0 Then
        If Frm57.L51_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm57.L48_Text = Frm57.L48_Text + 1 'Paparan Page ke-xxx
                Frm57_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm57.L48_Text) Then
                    If Frm57.L48_Text <> 1 Then
                        Frm57.L48_Text = Frm57.L48_Text - 1 'Paparan Page ke-xxx
                        Frm57_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm57.L48_Text - 1) * Frm57_PAGE_SIZE) + x
    
    Frm57.MSFlexGrid3.Rows = x + 1
    Frm57.MSFlexGrid3.TextMatrix(x, 0) = x
    Frm57.MSFlexGrid3.TextMatrix(x, 1) = Y
    If Not IsNull(rs!NoRujukanSistem) Then Frm57.MSFlexGrid3.TextMatrix(x, 2) = rs!NoRujukanSistem 'No. Rujukan Sistem
    If Not IsNull(rs!tarikh_belian) Then Frm57.MSFlexGrid3.TextMatrix(x, 3) = rs!tarikh_belian 'Tarikh Belian
    If Not IsNull(rs!no_siri) Then Frm57.MSFlexGrid3.TextMatrix(x, 4) = rs!no_siri   'No. Siri Produk
    If Not IsNull(rs!nama_produk) Then Frm57.MSFlexGrid3.TextMatrix(x, 5) = rs!nama_produk   'Nama Produk
    If Not IsNull(rs!supplier) Then Frm57.MSFlexGrid3.TextMatrix(x, 6) = rs!supplier   'Nama Supplier
    If Not IsNull(rs!purity) Then Frm57.MSFlexGrid3.TextMatrix(x, 7) = rs!purity   'Purity
    If IsNumeric(rs!Berat) Then
        Frm57.MSFlexGrid3.TextMatrix(x, 8) = Format(rs!Berat, "#,##0.00")  'Berat Jualan
        TOTALBERAT = TOTALBERAT + rs!Berat
    Else
        Frm57.MSFlexGrid3.TextMatrix(x, 8) = "-"
    End If
    If IsNumeric(rs!UPAH) Then
        Frm57.MSFlexGrid3.TextMatrix(x, 9) = Format(rs!UPAH, "#,##0.00") 'Upah
    Else
        Frm57.MSFlexGrid3.TextMatrix(x, 9) = "-"
    End If
    If IsNumeric(rs!harga_belian) Then
        Frm57.MSFlexGrid3.TextMatrix(x, 10) = Format(rs!harga_belian, "#,##0.00")  'Kos Belian Item
        TOTALMODAL = TOTALMODAL + rs!harga_belian
    Else
        Frm57.MSFlexGrid3.TextMatrix(x, 10) = "-"
    End If
    If IsNumeric(rs!KOSPERGRAM) Then
        Frm57.MSFlexGrid3.TextMatrix(x, 11) = Format(rs!KOSPERGRAM, "#,##0.00")  'Harga Belian Per Gram
    Else
        Frm57.MSFlexGrid3.TextMatrix(x, 11) = "-"
    End If
    If Not IsNull(rs!dulang) Then Frm57.MSFlexGrid3.TextMatrix(x, 12) = rs!dulang 'Dulang
    If Not IsNull(rs!panjang) Then Frm57.MSFlexGrid3.TextMatrix(x, 13) = rs!panjang   'Panjang
    If Not IsNull(rs!lebar) Then Frm57.MSFlexGrid3.TextMatrix(x, 14) = rs!lebar   'Lebar
    If Not IsNull(rs!dia) Then Frm57.MSFlexGrid3.TextMatrix(x, 15) = rs!dia   'Dia
    If Not IsNull(rs!Size) Then Frm57.MSFlexGrid3.TextMatrix(x, 16) = rs!Size   'Saiz
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'Frm57.L17_Text = Format(TOTALBERAT, "#,##0.00 g")
'Frm57.L18_Text = "RM " & Format(TOTALMODAL, "#,##0.00")
'Frm57.L19_Text = X

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from inventory where status='" & "0" & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm57_LM_TOTAL_PAGE = Format(rs(0) / Frm57_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm57_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm57_LM_PAGE = Split(Frm57_LM_TOTAL_PAGE, ".")(0)
        Frm57_LM_PAGE_LEBIHAN = Split(Frm57_LM_TOTAL_PAGE, ".")(1)
        
        If Frm57_LM_PAGE_LEBIHAN <> "00" Then
            Frm57.L49_Text = Frm57_LM_PAGE + 1
        Else
            Frm57.L49_Text = Frm57_LM_PAGE
        End If
        
    Else
    
        Frm57.L49_Text = Frm57_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm57.L49_Text = 0
    End If
Else
    Frm57.L49_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm57.L49_Text = vbNullString Then
    Frm57.L49_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah bilangan barang ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from inventory where status='" & "0" & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm57.L19_Text = rs(0)

rs.Close
Set rs = Nothing
'### Jumlah bilangan barang ### - End

'### Jumlah berat ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(Berat) from inventory where status='" & "0" & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm57.L17_Text = Format(rs(0), "#,##0.00 g") 'Jumlah (g)

rs.Close
Set rs = Nothing
'### Jumlah berat ### - End

'### Jumlah harga modal ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_belian) from inventory where status='" & "0" & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm57.L18_Text = "RM " & Format(rs(0), "#,##0.00") 'Jumlah (RM)

rs.Close
Set rs = Nothing
'### Jumlah harga modal ### - End

If x <> 0 Then
    Frm57.L50_Text = LM_START_ROW 'Titik Pencarian Data
    Frm57.Pic5.Visible = True
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    'If Frm57_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm57.L51_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm57.L51_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub Frm57_M_Inventori_Lain_header()
'On Error Resume Next
'////////////Report Inventori Yang Tiada Kaitan Dengan Dulang////////////////
Frm57.MSFlexGrid4.Clear
Frm57.MSFlexGrid4.Rows = 1
Frm57.MSFlexGrid4.RowHeight(0) = 700
Frm57.MSFlexGrid4.FormatString = "<No.|<No.|<No. Siri Produk|<Detail"

Frm57.MSFlexGrid4.ColWidth(0) = 0
Frm57.MSFlexGrid4.ColWidth(1) = 600
Frm57.MSFlexGrid4.ColWidth(2) = 2000
Frm57.MSFlexGrid4.ColWidth(3) = 10000
End Sub
Sub Frm57_M_Inventori_Lain()
'On Error Resume Next
Dim TOTALBERAT As Double
Dim TOTALMODAL As Double
Dim Frm57_LM_TOTAL_PAGE As Double
Dim Frm57_LM_JUMLAH As Double

Frm57_PAGE_SIZE = 38
Frm57_LM_TOTAL_PAGE = 0

x = 0
TOTALBERAT = 0
TOTALMODAL = 0

LM_START_ROW = Frm57.L54_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm57_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm57.L55_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm57_PAGE_SIZE
        End If
    End If
End If

Frm57_LM_PAGE_FOUND = 0

Frm57.L20_Text = "Report inventori bagi dulang [" & Frm57.L6_Text & "] (Ini adalah senarai barang yang TIADA kaitan dengan dulang ini)."

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from inventory2 LIMIT " & LM_START_ROW & "," & Frm57_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm57_LM_PAGE_FOUND = 0 Then
        If Frm57.L55_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm57.L52_Text = Frm57.L52_Text + 1 'Paparan Page ke-xxx
                Frm57_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm57.L52_Text) Then
                    If Frm57.L52_Text <> 1 Then
                        Frm57.L52_Text = Frm57.L52_Text - 1 'Paparan Page ke-xxx
                        Frm57_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm57.L52_Text - 1) * Frm57_PAGE_SIZE) + x
    
    Frm57.MSFlexGrid4.Rows = x + 1
    Frm57.MSFlexGrid4.TextMatrix(x, 0) = x
    Frm57.MSFlexGrid4.TextMatrix(x, 1) = Y
    If Not IsNull(rs!no_siri) Then Frm57.MSFlexGrid4.TextMatrix(x, 2) = rs!no_siri 'No. Siri Produk
    If Not IsNull(rs!Detail) Then Frm57.MSFlexGrid4.TextMatrix(x, 3) = rs!Detail 'Detail
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'Frm57.L26_Text = X

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from inventory2", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm57_LM_TOTAL_PAGE = Format(rs(0) / Frm57_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm57_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm57_LM_PAGE = Split(Frm57_LM_TOTAL_PAGE, ".")(0)
        Frm57_LM_PAGE_LEBIHAN = Split(Frm57_LM_TOTAL_PAGE, ".")(1)
        
        If Frm57_LM_PAGE_LEBIHAN <> "00" Then
            Frm57.L53_Text = Frm57_LM_PAGE + 1
        Else
            Frm57.L53_Text = Frm57_LM_PAGE
        End If
        
    Else
    
        Frm57.L53_Text = Frm57_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm57.L53_Text = 0
    End If
Else
    Frm57.L53_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm57.L53_Text = vbNullString Then
    Frm57.L53_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah bilangan barang ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from inventory2", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm57.L26_Text = rs(0)

rs.Close
Set rs = Nothing
'### Jumlah bilangan barang ### - End

If x <> 0 Then
    Frm57.L54_Text = LM_START_ROW 'Titik Pencarian Data
    Frm57.Pic6.Visible = True
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    'If Frm57_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm57.L55_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm57.L55_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
