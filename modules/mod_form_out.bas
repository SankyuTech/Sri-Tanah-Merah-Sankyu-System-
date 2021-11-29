Attribute VB_Name = "mod_form_out"
Sub Frm107_initial_setting()
'on error resume next
Frm107.CBB1.Clear
'Frm107.CBB1.AddItem "Semua status"
Frm107.CBB1.AddItem "Barang trade in"
Frm107.CBB1.AddItem "Barang potong"

Frm107.CBB1 = "Barang trade in"

Frm107.CBB2.Clear
Frm107.CBB2.AddItem "Semua purity"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database order by kadar_tukaran_9999 DESC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!Kod_Metal_Purity) Then
        Frm107.CBB2.AddItem rs!Kod_Metal_Purity
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm107.CBB2 = "Semua purity"
End Sub
Sub Frm107_clear_status()
'on error resume next
' ### Reset semua status ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "UPDATE Data_Database set form_out_status = Null"

Set rs = cn.Execute(strsql)
Set rs = Nothing
' ### Reset semua status ### - End

'#### Tukar status semua barang trade in dan potong kepada 0 #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "UPDATE Data_Database set form_out_status='" & 0 & "'" _
& "WHERE StatusItem='" & 10 & "' AND ( receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 5 & "')"

Set rs = cn.Execute(strsql)
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "UPDATE Data_Database set form_out_status='" & 0 & "'" _
& "WHERE StatusItem='" & 12 & "' OR StatusItem='" & 20 & "' OR StatusItem='" & 22 & "'"

Set rs = cn.Execute(strsql)
Set rs = Nothing
'#### Tukar status semua barang trade in dan potong kepada 0 #### - End
End Sub
Sub Frm107_senarai_barang_header()
'on error resume next
'#### Header senarai barang #### - Start
Frm107.MSFlexGrid1.Clear
Frm107.MSFlexGrid1.Rows = 1
Frm107.MSFlexGrid1.RowHeight(0) = 600
Frm107.MSFlexGrid1.FormatString = "No.|<No.|<No. ID|<No. Siri Produk|<Nama Produk|<Purity|<Berat (g)|<Modal (RM)|<Status"

Frm107.MSFlexGrid1.ColWidth(0) = 0 'No.
Frm107.MSFlexGrid1.ColWidth(1) = 600 'No.
Frm107.MSFlexGrid1.ColWidth(2) = 0 'No. ID
Frm107.MSFlexGrid1.ColWidth(3) = 1000 'No. Siri Produk
Frm107.MSFlexGrid1.ColWidth(4) = 2500 'Nama Produk
Frm107.MSFlexGrid1.ColWidth(5) = 800 'Purity
Frm107.MSFlexGrid1.ColAlignment(5) = 4

Frm107.MSFlexGrid1.ColWidth(6) = 800 'Berat (g)
Frm107.MSFlexGrid1.ColAlignment(6) = 7

Frm107.MSFlexGrid1.ColWidth(7) = 1350 'Modal
Frm107.MSFlexGrid1.ColAlignment(7) = 7

Frm107.MSFlexGrid1.ColWidth(8) = 1350 'Status
Frm107.MSFlexGrid1.ColAlignment(8) = 4
'#### Header senarai barang #### - End
End Sub
Sub Frm107_senarai_barang()
'on error resume next
Dim Frm107_LM_TOTAL_PAGE As Double

Frm107_PAGE_SIZE = 34
Frm107_LM_TOTAL_PAGE = 0
Frm107_LM_MODE = 0 '0 : Barang trade in , 1 : Barang potong
x = 0

Frm107.L8_Text = 0
Frm107.L9_Text = "0.00 g"

If Frm107.L6_Text = "Barang trade in" Then
    Frm107_LM_MODE = 0 '0 : Barang trade in , 1 : Barang potong
Else
    Frm107_LM_MODE = 1 '0 : Barang trade in , 1 : Barang potong
End If

If Frm107.L7_Text = "Semua purity" Then
    Frm107_LM_SEARCH_1 = Null
    Frm107_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm107_LM_SEARCH_1 = Frm107.L7_Text
    Frm107_LM_SEARCH_1_LOGIC = "="
End If

'Header senarai
If Frm107_LM_MODE = 0 Then '0 : Barang trade in , 1 : Barang potong
    Frm107.L5_Text = "Senarai barang trade in bagi " & LCase(Frm107.L7_Text) & "." 'Report Header
ElseIf Frm107_LM_MODE = 1 Then '0 : Barang trade in , 1 : Barang potong
    Frm107.L5_Text = "Senarai barang potong bagi " & LCase(Frm107.L7_Text) & "." 'Report Header
End If

LM_START_ROW = Frm107.L3_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm107_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm107.L4_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm107_PAGE_SIZE
        End If
    End If
End If

Frm107_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm107_LM_MODE = 0 Then rs.Open "select * from data_database where kod_purity " & Frm107_LM_SEARCH_1_LOGIC & "'" & Frm107_LM_SEARCH_1 & "' AND ((StatusItem='" & 10 & "' AND ( receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 5 & "')) OR (StatusItem='" & 23 & "' AND (form_out_status='" & 1 & "' OR form_out_status='" & 4 & "'))) order by no_siri_produk ASC LIMIT " & LM_START_ROW & "," & Frm107_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If Frm107_LM_MODE = 1 Then rs.Open "select * from data_database where kod_purity " & Frm107_LM_SEARCH_1_LOGIC & "'" & Frm107_LM_SEARCH_1 & "' AND (StatusItem='" & 12 & "' OR StatusItem='" & 20 & "' OR StatusItem='" & 22 & "') OR (StatusItem='" & 24 & "' AND (form_out_status='" & 1 & "' OR form_out_status='" & 4 & "')) order by no_siri_produk ASC LIMIT " & LM_START_ROW & "," & Frm107_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm107_LM_PAGE_FOUND = 0 Then
        If Frm107.L4_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm107.L1_Text = Frm107.L1_Text + 1 'Paparan Page ke-xxx
                Frm107_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm107.L1_Text) Then
                    If Frm107.L1_Text <> 1 Then
                        Frm107.L1_Text = Frm107.L1_Text - 1 'Paparan Page ke-xxx
                        Frm107_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm107.L1_Text - 1) * Frm107_PAGE_SIZE) + x
    Frm107.MSFlexGrid1.Rows = x + 1
    Frm107.MSFlexGrid1.TextMatrix(x, 0) = x 'No.
    Frm107.MSFlexGrid1.TextMatrix(x, 1) = Y 'No.
    Frm107.MSFlexGrid1.ColAlignment(1) = 4
    Frm107.MSFlexGrid1.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!no_siri_Produk) Then Frm107.MSFlexGrid1.TextMatrix(x, 3) = rs!no_siri_Produk 'No. siri produk
    If Not IsNull(rs!kategori_Produk) Then Frm107.MSFlexGrid1.TextMatrix(x, 4) = rs!kategori_Produk 'Nama produk
    If Not IsNull(rs!kod_Purity) Then Frm107.MSFlexGrid1.TextMatrix(x, 5) = rs!kod_Purity 'Purity
    If Not IsNull(rs!beza_berat) Then Frm107.MSFlexGrid1.TextMatrix(x, 6) = Format(rs!beza_berat, "#,##0.00 g") 'Berat
    If Not IsNull(rs!receiving_Status) Then
        
        If rs!receiving_Status = "0" Or rs!receiving_Status = "2" Or rs!receiving_Status = "4" Or rs!receiving_Status = "5" Or rs!receiving_Status = "6" Or rs!receiving_Status = "8" Then
        
            If Not IsNull(rs!beza_berat) And Not IsNull(rs!harga_Per_Gram_Item) Then
                Frm107.MSFlexGrid1.TextMatrix(x, 7) = Format(rs!beza_berat * rs!harga_Per_Gram_Item, "#,##0.00") 'Modal
            End If
            
        ElseIf rs!receiving_Status = "1" Or rs!receiving_Status = "3" Or rs!receiving_Status = "7" Then

            If Not IsNull(rs!harga_item) Then
                Frm107.MSFlexGrid1.TextMatrix(x, 7) = Format(rs!harga_item, "#,##0.00") 'Modal
            End If
            
        End If
        
    Else
    
        Frm107.MSFlexGrid1.TextMatrix(x, 7) = Format(0, "#,##0.00") 'Modal
        
    End If

'0:  BK
'2 : Trade In BK
'4:  gold Bar
'5 : Trade In Gold Bar
'6 : Emas terpakai BK
'8 : Emas terpakai gold bar

'1:  Barang permata
'3 : Trade In Barang Permata
'7 : Emas terpakai permata
    
    If Not IsNull(rs!form_out_status) Then
        If rs!form_out_status = 0 Or rs!form_out_status = 4 Then
            Frm107.MSFlexGrid1.TextMatrix(x, 8) = "Belum dipilih"
        ElseIf rs!form_out_status = 1 Then
            Frm107.MSFlexGrid1.TextMatrix(x, 8) = "SUDAH DIPILIH"
        End If
    Else
        Frm107.MSFlexGrid1.TextMatrix(x, 8) = "Belum dipilih"
    End If
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm107_LM_MODE = 0 Then rs.Open "select COUNT(ID) from data_database where kod_purity " & Frm107_LM_SEARCH_1_LOGIC & "'" & Frm107_LM_SEARCH_1 & "' AND ((StatusItem='" & 10 & "' AND ( receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 5 & "')) OR (StatusItem='" & 23 & "' AND (form_out_status='" & 1 & "' OR form_out_status='" & 4 & "')))", cn, adOpenKeyset, adLockOptimistic
If Frm107_LM_MODE = 1 Then rs.Open "select COUNT(ID) from data_database where kod_purity " & Frm107_LM_SEARCH_1_LOGIC & "'" & Frm107_LM_SEARCH_1 & "' AND (StatusItem='" & 12 & "' OR StatusItem='" & 20 & "' OR StatusItem='" & 22 & "') OR (StatusItem='" & 24 & "' AND (form_out_status='" & 1 & "' OR form_out_status='" & 4 & "'))", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm107_LM_TOTAL_PAGE = Format(rs(0) / Frm107_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm107_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm107_LM_PAGE = Split(Frm107_LM_TOTAL_PAGE, ".")(0)
        Frm107_LM_PAGE_LEBIHAN = Split(Frm107_LM_TOTAL_PAGE, ".")(1)
        
        If Frm107_LM_PAGE_LEBIHAN <> "00" Then
            Frm107.L2_Text = Frm107_LM_PAGE + 1
        Else
            Frm107.L2_Text = Frm107_LM_PAGE
        End If
        
    Else
    
        Frm107.L2_Text = Frm107_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm107.L2_Text = 0
    End If
Else
    Frm107.L2_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm107.L2_Text = vbNullString Then
    Frm107.L2_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah bilangan barang keseluruhan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm107_LM_MODE = 0 Then rs.Open "select COUNT(ID) from data_database where kod_purity " & Frm107_LM_SEARCH_1_LOGIC & "'" & Frm107_LM_SEARCH_1 & "' AND ((StatusItem='" & 10 & "' AND ( receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 5 & "')) OR (StatusItem='" & 23 & "' AND (form_out_status='" & 1 & "' OR form_out_status='" & 4 & "')))", cn, adOpenKeyset, adLockOptimistic
If Frm107_LM_MODE = 1 Then rs.Open "select COUNT(ID) from data_database where kod_purity " & Frm107_LM_SEARCH_1_LOGIC & "'" & Frm107_LM_SEARCH_1 & "' AND (StatusItem='" & 12 & "' OR StatusItem='" & 20 & "' OR StatusItem='" & 22 & "') OR (StatusItem='" & 24 & "' AND (form_out_status='" & 1 & "' OR form_out_status='" & 4 & "'))", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm107.L8_Text = rs(0)

rs.Close
Set rs = Nothing
'### Jumlah bilangan barang keseluruhan ### - End

'### Jumlah berat keseluruhan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm107_LM_MODE = 0 Then rs.Open "select SUM(beza_berat) from data_database where kod_purity " & Frm107_LM_SEARCH_1_LOGIC & "'" & Frm107_LM_SEARCH_1 & "' AND ((StatusItem='" & 10 & "' AND ( receiving_Status='" & 2 & "' OR receiving_Status='" & 3 & "' OR receiving_Status='" & 5 & "')) OR (StatusItem='" & 23 & "' AND (form_out_status='" & 1 & "' OR form_out_status='" & 4 & "')))", cn, adOpenKeyset, adLockOptimistic
If Frm107_LM_MODE = 1 Then rs.Open "select SUM(beza_berat) from data_database where kod_purity " & Frm107_LM_SEARCH_1_LOGIC & "'" & Frm107_LM_SEARCH_1 & "' AND (StatusItem='" & 12 & "' OR StatusItem='" & 20 & "' OR StatusItem='" & 22 & "') OR (StatusItem='" & 24 & "' AND (form_out_status='" & 1 & "' OR form_out_status='" & 4 & "'))", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm107.L9_Text = Format(rs(0), "#,##0.00 g")

rs.Close
Set rs = Nothing
'### Jumlah berat keseluruhan ### - End

If x <> 0 Then
    Frm107.L3_Text = LM_START_ROW 'Titik Pencarian Data
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    If Frm107_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm107.L4_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm107.L4_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub Frm107_initial_location()
'on error resume next
Frm107.Pic1.Left = 120
Frm107.Pic1.Top = 240
Frm107.Pic2.Left = 120
Frm107.Pic2.Top = 240
Frm107.Pic4.Left = 120
Frm107.Pic4.Top = 240

Frm107.Pic1.Visible = False
Frm107.Pic2.Visible = False
Frm107.Pic4.Visible = False
End Sub
Sub Frm107_initial_location2()
'on error resume next
Frm107.Pic3.Left = 120
Frm107.Pic3.Top = 720

Frm107.Pic3.Visible = False
End Sub
Sub Frm107_initial_location3()
'on error resume next
Frm107.Pic5.Left = 6720
Frm107.Pic5.Top = 120
Frm107.Pic6.Left = 6360
Frm107.Pic6.Top = 120
Frm107.Pic7.Left = 6360
Frm107.Pic7.Top = 120

Frm107.Pic5.Visible = False
Frm107.Pic6.Visible = False
Frm107.Pic7.Visible = False
End Sub
Sub Frm107_initial_setting1()
'on error resume next
'### Setting bagi NO RUJUKAN SISTEM
'### Setting bagi NAMA SUPPLIER / KILANG
Dim Frm107_LM_NO_RUJUKAN As Long

Frm107_LM_NO_RUJUKAN = 1

Frm107.L11_Text = vbNullString
Frm107.L31_Text = 1 'Memori : No. rujukan turutan bagi description (auto generated number)

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Default1 = "Default" Then
        If Not IsNull(rs!no_rujukan_form) Then 'No. rujukan sistem
            If IsNumeric(rs!no_rujukan_form) Then
                Frm107.L10_Text = rs!no_rujukan_form
            Else
                Frm107.L10_Text = 1
            End If
        Else
            Frm107.L10_Text = 1
        End If
    Else
        Frm107.L10_Text = 1
    End If
End If

rs.Close
Set rs = Nothing

'### Periksa samada nombor rujukan ini telah digunakan atau belum ### - Start

Frm107_LM_NO_RUJUKAN = Frm107.L10_Text

Re_Gen_No_Rujukan:

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 57_form_out where no_rujukan='" & Frm107_LM_NO_RUJUKAN & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm107_LM_NO_RUJUKAN = Frm107_LM_NO_RUJUKAN + 1
    
    Frm107.L10_Text = Frm107_LM_NO_RUJUKAN 'No. rujukan sistem
    
    rs.Close
    Set rs = Nothing
    GoTo Re_Gen_No_Rujukan:
End If

rs.Close
Set rs = Nothing
'### Periksa samada nombor rujukan ini telah digunakan atau belum ### - End

Frm107.CBB3.Clear
Frm107.CBB6.Clear

Frm107.CBB6.AddItem "Semua supplier"

'### Nama supplier / kilang ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!supplier) Then
        If rs!Status = 1 Then Frm107.CBB3.AddItem rs!supplier
    
        Frm107.CBB6.AddItem rs!supplier
    End If
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'### Nama supplier / kilang ### - End

Frm107.CBB6 = "Semua supplier"

'###Padam Table 60_form_out_list_temp ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "TRUNCATE TABLE " & G_FORM_OUT_DESC & ""

Set rs = cn.Execute(strsql)
Set rs = Nothing
'###Padam Table 60_form_out_list_temp ### - End

'###Padam Table 61_form_out_item_list_temp ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "TRUNCATE TABLE " & G_FORM_LIST & ""

Set rs = cn.Execute(strsql)
Set rs = Nothing
'###Padam Table 61_form_out_item_list_temp ### - End

GM_NEXT_PREV = 0
Frm107.L15_Text = -1 'Titik Pencarian Data
Frm107.L16_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Frm107.L13_Text = 0 'Paparan Page ke-xxx

Call Frm107_senarai_description_header
Call Frm107_senarai_description

Frm107.L12_Text.Visible = True
Frm107.L17_Text.Visible = True

Frm107.L57_Text = 0 '0 : Data baru , 1:  Data Edit
End Sub
Sub Frm107_initial_setting2()
'on error resume next
'### Setting bagi PURITY
'### Clear semua component bagi description

Frm107.TB1 = vbNullString
Frm107.TB2 = vbNullString
'Frm107.TB3 = "1.00"

Frm107.CBB4.Clear

'### Senarai purity ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!Kod_Metal_Purity) Then Frm107.CBB4.AddItem rs!Kod_Metal_Purity
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'### Senarai purity ### - End
End Sub
Sub Frm107_initial_setting3()
'on error resume next
'### Digunakan pada bahagian report

Frm107.CBB6.Clear
Frm107.CBB6.AddItem "Semua supplier"

'### Nama supplier / kilang ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!supplier) Then Frm107.CBB6.AddItem rs!supplier
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'### Nama supplier / kilang ### - End

Frm107.CBB6 = "Semua supplier"

Frm107.TB4 = vbNullString
End Sub
Sub Frm107_initial_setting4()
'on error resume next
'###Senarai Nama Pekerja###
Frm107.CBB5.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where status='" & "Aktif" & "' AND InvestorSmall = 0 AND InvestorBig = 0", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    If rs!user_level = 1 Or rs!user_level = 2 Or rs!user_level = 3 Or rs!user_level = G_LEVEL_USER Then Frm107.CBB5.AddItem rs!Samaran & "  |  " & rs!NoPekerja
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Call Frm107_jurujual
End Sub
Sub Frm107_senarai_description_header()
'on error resume next
'#### Header senarai description #### - Start

Frm107.MSFlexGrid2.Clear
Frm107.MSFlexGrid2.Rows = 1
Frm107.MSFlexGrid2.RowHeight(0) = 600
Frm107.MSFlexGrid2.FormatString = "No.|<No.|<No. ID|<Description|<Purity|<Berat Asal (g)|<Mutu|<Berat 999.9 (g)|<Modal (RM)"

Frm107.MSFlexGrid2.ColWidth(0) = 0 'No.
Frm107.MSFlexGrid2.ColWidth(1) = 600 'No.
Frm107.MSFlexGrid2.ColWidth(2) = 0 'No. ID
Frm107.MSFlexGrid2.ColWidth(3) = 4500 'Description
Frm107.MSFlexGrid2.ColWidth(4) = 800 'Purity

Frm107.MSFlexGrid2.ColWidth(5) = 1000 'Berat Asal (g)
Frm107.MSFlexGrid2.ColAlignment(5) = 7

Frm107.MSFlexGrid2.ColWidth(6) = 800 'Mutu
Frm107.MSFlexGrid2.ColAlignment(6) = 4

Frm107.MSFlexGrid2.ColWidth(7) = 1000 'Berat 999.9 (g)
Frm107.MSFlexGrid2.ColAlignment(7) = 7

Frm107.MSFlexGrid2.ColWidth(8) = 1800 'Modal (RM)
Frm107.MSFlexGrid2.ColAlignment(8) = 7
'#### Header senarai description #### - End
End Sub
Sub Frm107_senarai_description()
'on error resume next
Dim Frm107_LM_TOTAL_PAGE As Double

Frm107_PAGE_SIZE = 7
Frm107_LM_TOTAL_PAGE = 0
Frm107_LM_MODE = 0 '0 : Barang trade in , 1 : Barang potong
x = 0

Frm107.L28_Text = 0
Frm107.L29_Text = "0.00"
Frm107.L59_Text = "0.00"

LM_START_ROW = Frm107.L15_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm107_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm107.L16_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm107_PAGE_SIZE
        End If
    End If
End If

Frm107_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from " & G_FORM_OUT_DESC & " where (status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "') order by ID ASC LIMIT " & LM_START_ROW & "," & Frm107_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm107_LM_PAGE_FOUND = 0 Then
        If Frm107.L16_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm107.L13_Text = Frm107.L13_Text + 1 'Paparan Page ke-xxx
                Frm107_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm107.L13_Text) Then
                    If Frm107.L13_Text <> 1 Then
                        Frm107.L13_Text = Frm107.L13_Text - 1 'Paparan Page ke-xxx
                        Frm107_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm107.L13_Text - 1) * Frm107_PAGE_SIZE) + x
    Frm107.MSFlexGrid2.Rows = x + 1
    Frm107.MSFlexGrid2.TextMatrix(x, 0) = x 'No.
    Frm107.MSFlexGrid2.TextMatrix(x, 1) = Y 'No.
    Frm107.MSFlexGrid2.ColAlignment(1) = 4
    Frm107.MSFlexGrid2.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!Description) Then Frm107.MSFlexGrid2.TextMatrix(x, 3) = rs!Description 'Description
    If Not IsNull(rs!purity) Then Frm107.MSFlexGrid2.TextMatrix(x, 4) = rs!purity 'Purity
    
    If Not IsNull(rs!berat_before) Then 'Berat asal (g)
        Frm107.MSFlexGrid2.TextMatrix(x, 5) = Format(rs!berat_before, "#,##0.00 g")
    Else
        Frm107.MSFlexGrid2.TextMatrix(x, 5) = "0.00 g"
    End If
    
    If Not IsNull(rs!Conversion) Then 'Mutu
        Frm107.MSFlexGrid2.TextMatrix(x, 6) = rs!Conversion
    Else
        Frm107.MSFlexGrid2.TextMatrix(x, 6) = "0.00"
    End If
    
    If Not IsNull(rs!berat_after) Then 'Berat 999.9
        Frm107.MSFlexGrid2.TextMatrix(x, 7) = Format(rs!berat_after, "#,##0.00 g")
    Else
        Frm107.MSFlexGrid2.TextMatrix(x, 7) = "0.00 g"
    End If
    
    If Not IsNull(rs!modal) Then Frm107.MSFlexGrid2.TextMatrix(x, 8) = Format(rs!modal, "#,##0.00")

    Frm107.MSFlexGrid2.RowHeight(x) = 700
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from " & G_FORM_OUT_DESC & " where (status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "')", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm107_LM_TOTAL_PAGE = Format(rs(0) / Frm107_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm107_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm107_LM_PAGE = Split(Frm107_LM_TOTAL_PAGE, ".")(0)
        Frm107_LM_PAGE_LEBIHAN = Split(Frm107_LM_TOTAL_PAGE, ".")(1)
        
        If Frm107_LM_PAGE_LEBIHAN <> "00" Then
            Frm107.L14_Text = Frm107_LM_PAGE + 1
        Else
            Frm107.L14_Text = Frm107_LM_PAGE
        End If
        
    Else
    
        Frm107.L14_Text = Frm107_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm107.L14_Text = 0
    End If
Else
    Frm107.L14_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm107.L14_Text = vbNullString Then
    Frm107.L14_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah bilangan description ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from " & G_FORM_OUT_DESC & " where (status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "')", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm107.L28_Text = rs(0)

rs.Close
Set rs = Nothing
'### Jumlah bilangan description ### - End

'### Jumlah berat description ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(berat_after) from " & G_FORM_OUT_DESC & " where (status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "')", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm107.L29_Text = Format(rs(0), "#,##0.00")

rs.Close
Set rs = Nothing
'### Jumlah berat description ### - End

'### Jumlah modal ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(modal) from " & G_FORM_OUT_DESC & " where (status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "')", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm107.L59_Text = Format(rs(0), "#,##0.00")

rs.Close
Set rs = Nothing
'### Jumlah modal ### - End

If x <> 0 Then
    Frm107.L15_Text = LM_START_ROW 'Titik Pencarian Data
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    'If Frm107_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm107.L16_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm107.L16_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub Frm107_senarai_hantar_header()
'on error resume next
'#### Header senarai description #### - Start
Frm107.MSFlexGrid3.Clear
Frm107.MSFlexGrid3.Rows = 1
Frm107.MSFlexGrid3.RowHeight(0) = 600
Frm107.MSFlexGrid3.FormatString = "No.|<No.|<No. ID|<No. Siri Produk|<Purity|<Berat (g)|<Modal (RM)|<Jenis Barang"

Frm107.MSFlexGrid3.ColWidth(0) = 0 'No.
Frm107.MSFlexGrid3.ColWidth(1) = 800 'No.
Frm107.MSFlexGrid3.ColWidth(2) = 0 'No. ID
Frm107.MSFlexGrid3.ColWidth(3) = 1300 'No. Siri Produk
Frm107.MSFlexGrid3.ColWidth(4) = 1000 'Purity
Frm107.MSFlexGrid3.ColAlignment(4) = 4

Frm107.MSFlexGrid3.ColWidth(5) = 1000 'Berat (g)
Frm107.MSFlexGrid3.ColAlignment(5) = 7

Frm107.MSFlexGrid3.ColWidth(6) = 1300 'Modal (RM)
Frm107.MSFlexGrid3.ColAlignment(6) = 7

Frm107.MSFlexGrid3.ColWidth(7) = 1800 'Jenis Barang
Frm107.MSFlexGrid3.ColAlignment(7) = 4
'#### Header senarai description #### - End
End Sub
Sub Frm107_senarai_hantar()
'on error resume next
Dim Frm107_LM_TOTAL_PAGE As Double
Dim Frm107_LM_WEIGHT As Double
Dim Frm107_LM_MUTU As Double

Frm107_PAGE_SIZE = 34
Frm107_LM_TOTAL_PAGE = 0
x = 0
Frm107_LM_WEIGHT = 0
Frm107_LM_MUTU = 0

Frm107.L26_Text = 0
Frm107.L27_Text = "0.00 g"
Frm107.L58_Text = "0.00"

LM_START_ROW = Frm107.L24_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm107_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm107.L25_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm107_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        Frm107.L22_Text = 1
    End If
End If

Frm107_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from " & G_FORM_LIST & " where id_rujukan='" & Frm107.L19_Text & "' AND (status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "') order by no_siri_produk ASC LIMIT " & LM_START_ROW & "," & Frm107_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm107_LM_PAGE_FOUND = 0 Then
        If Frm107.L25_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm107.L22_Text = Frm107.L22_Text + 1 'Paparan Page ke-xxx
                Frm107_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm107.L22_Text) Then
                    If Frm107.L22_Text <> 1 Then
                        Frm107.L22_Text = Frm107.L22_Text - 1 'Paparan Page ke-xxx
                        Frm107_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm107.L22_Text - 1) * Frm107_PAGE_SIZE) + x
    Frm107.MSFlexGrid3.Rows = x + 1
    Frm107.MSFlexGrid3.TextMatrix(x, 0) = x 'No.
    Frm107.MSFlexGrid3.TextMatrix(x, 1) = Y 'No.
    Frm107.MSFlexGrid3.ColAlignment(1) = 4
    Frm107.MSFlexGrid3.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!no_siri_Produk) Then Frm107.MSFlexGrid3.TextMatrix(x, 3) = rs!no_siri_Produk 'No. siri produk
    If Not IsNull(rs!purity) Then Frm107.MSFlexGrid3.TextMatrix(x, 4) = rs!purity 'Purity

    If Not IsNull(rs!Berat) Then Frm107.MSFlexGrid3.TextMatrix(x, 5) = Format(rs!Berat, "#,##0.00 g") 'Berat
    
    If Not IsNull(rs!modal) Then Frm107.MSFlexGrid3.TextMatrix(x, 6) = Format(rs!modal, "#,##0.00") 'Modal
    
    If Not IsNull(rs!jenis_barang) Then Frm107.MSFlexGrid3.TextMatrix(x, 7) = rs!jenis_barang 'Jenis Barang

    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from " & G_FORM_LIST & " where id_rujukan='" & Frm107.L19_Text & "' AND (status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "')", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm107_LM_TOTAL_PAGE = Format(rs(0) / Frm107_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm107_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm107_LM_PAGE = Split(Frm107_LM_TOTAL_PAGE, ".")(0)
        Frm107_LM_PAGE_LEBIHAN = Split(Frm107_LM_TOTAL_PAGE, ".")(1)
        
        If Frm107_LM_PAGE_LEBIHAN <> "00" Then
            Frm107.L23_Text = Frm107_LM_PAGE + 1
        Else
            Frm107.L23_Text = Frm107_LM_PAGE
        End If
        
    Else
    
        Frm107.L23_Text = Frm107_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm107.L23_Text = 0
    End If
Else
    Frm107.L23_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm107.L23_Text = vbNullString Then
    Frm107.L23_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah bilangan barang keseluruhan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from " & G_FORM_LIST & " where id_rujukan='" & Frm107.L19_Text & "' AND (status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "')", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm107.L26_Text = rs(0)

rs.Close
Set rs = Nothing
'### Jumlah bilangan barang keseluruhan ### - End

'### Jumlah berat barang keseluruhan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(berat) from " & G_FORM_LIST & " where id_rujukan='" & Frm107.L19_Text & "' AND (status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "')", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then
    Frm107.L27_Text = Format(rs(0), "#,##0.00 g")
    Frm107_LM_WEIGHT = rs(0)
End If

rs.Close
Set rs = Nothing
'### Jumlah berat barang keseluruhan ### - End

'### Jumlah MODAL keseluruhan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(modal) from " & G_FORM_LIST & " where id_rujukan='" & Frm107.L19_Text & "' AND (status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "')", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then
    Frm107.L58_Text = Format(rs(0), "#,##0.00")
End If

rs.Close
Set rs = Nothing
'### Jumlah MODAL keseluruhan ### - End

'### Update berat terkumpul dalam table 60_form_out_list_temp ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from " & G_FORM_OUT_DESC & " where no_rujukan='" & Frm107.L10_Text & "' AND id_desc='" & Frm107.L19_Text & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!Conversion) Then
        If IsNumeric(rs!Conversion) Then
            Frm107_LM_MUTU = rs!Conversion
        End If
    End If
    rs!berat_before = Format(Frm107_LM_WEIGHT, "#,##0.00")
    rs!berat_after = Format(Frm107_LM_MUTU * Frm107_LM_WEIGHT, "#,##0.00")
    If Frm107.L58_Text <> vbNullString Then
        rs!modal = Format(Frm107.L58_Text, "0.00")
    Else
        rs!modal = Format(0, "0.00")
    End If
    rs.Update
End If

rs.Close
Set rs = Nothing
'### Update berat terkumpul dalam table 60_form_out_list_temp ### - End

If x <> 0 Then
    Frm107.L24_Text = LM_START_ROW 'Titik Pencarian Data
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    'If Frm107_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm107.L25_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm107.L25_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub Frm107_tukaran_mutu()
'on error resume next
Dim Frm107_LM_MUTU As Double
Dim Frm107_LM_BERAT As Double

Frm107_LM_MUTU = 0
Frm107_LM_BERAT = 0

If (Frm107.L29_Text <> vbNullString And IsNumeric(Frm107.L29_Text)) And (Frm107.TB3 <> vbNullString And IsNumeric(Frm107.TB3)) Then
    Frm107_LM_MUTU = Frm107.TB3
    Frm107_LM_BERAT = Frm107.L29_Text
    
    Frm107.L30_Text = Format(Frm107_LM_MUTU * Frm107_LM_BERAT, "#,##0.00 g")
    
Else
    Frm107.L30_Text = Format(0, "#,##0.00 g")
End If
End Sub
Sub Frm107_cetak_penyata_forming()
'on error resume next
Dim Frm107_LM_NO_RUJUKAN

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
        PRINTER_FOUND = 1 '0 : Printer Not Found , 1 : Printer Found
'        Exit For
'    End If
'Next

'G_No_STATMENT_FORM = "000002"
Frm107_LM_NO_RUJUKAN = vbNullString
Frm107_LM_ID_KEDAI = vbNullString

'### Reset maklumat kedai ### - Start
Report71.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report71.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report71.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report71.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report71.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report71.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report71.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report71.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report71.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report71.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

'### Reset semua butiran ### - Start
Report71.Sections("Section4").Controls("L1").Caption = vbNullString 'Kilang / Supplier
Report71.Sections("Section4").Controls("L2").Caption = vbNullString 'Alamat
Report71.Sections("Section4").Controls("L3").Caption = vbNullString 'No. Rujukan / No. Statement
Report71.Sections("Section4").Controls("L4").Caption = vbNullString 'Tarikh

Report71.Sections("Section5").Controls("L5").Caption = "0.00 g" 'Jumlah berat
Report71.Sections("Section5").Controls("L6").Caption = "0.00" '% mutu
Report71.Sections("Section5").Controls("L7").Caption = "0.00 g" 'Jumlah berat 999.9
'### Reset semua butiran ### - End

Report71.Sections("Section4").Controls("L3").Caption = G_No_STATMENT_FORM 'No. Rujukan / No. Statement

'### Carian no rujukan bagi statement ini ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 57_form_out where no_statement='" & G_No_STATMENT_FORM & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!no_rujukan) Then Frm107_LM_NO_RUJUKAN = rs!no_rujukan
    If Not IsNull(rs!nama_kedai) Then Report71.Sections("Section4").Controls("L1").Caption = rs!nama_kedai 'Kilang / Supplier
    If Not IsNull(rs!tarikh) Then Report71.Sections("Section4").Controls("L4").Caption = rs!tarikh 'Tarikh
    If Not IsNull(rs!berat_before) Then Report71.Sections("Section5").Controls("L5").Caption = Format(rs!berat_before, "#,##0.00 g") 'Jumlah berat
    If Not IsNull(rs!Conversion) Then Report71.Sections("Section5").Controls("L6").Caption = rs!Conversion '% mutu
    If Not IsNull(rs!berat_after) Then Report71.Sections("Section5").Controls("L7").Caption = Format(rs!berat_after, "#,##0.00 g") 'Jumlah berat 999.9
    
    If Not IsNull(rs!id_kedai) Then
        Frm107_LM_ID_KEDAI = rs!id_kedai
    Else
        Frm107_LM_ID_KEDAI = 1
    End If
End If

rs.Close
Set rs = Nothing
'### Carian no rujukan bagi statement ini ### - End

If Frm107_LM_NO_RUJUKAN <> vbNullString Then

    '### Carian alamat kedai ### - Start
    If Frm107_LM_ID_KEDAI <> vbNullString Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from setting_database where ID='" & Frm107_LM_ID_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            If Not IsNull(rs!alamat) Then Report71.Sections("Section4").Controls("L2").Caption = rs!alamat 'Alamat
    
        End If
        
        rs.Close
        Set rs = Nothing
        
    End If
    '### Carian alamat kedai ### - Start
    
    '### Paparan statement ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 58_form_out_list where no_rujukan='" & Frm107_LM_NO_RUJUKAN & "' AND status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        Set Report71.DataSource = rs
        Report71.Show
        rs.MoveNext
    Wend
    
    'rs.Close
    Set rs = Nothing
    '### Paparan statement ### - End


End If
End Sub
Sub Frm107_report_statement_header()
'on error resume next
'#### Header senarai description #### - Start
Frm107.MSFlexGrid4.Clear
Frm107.MSFlexGrid4.Rows = 1
Frm107.MSFlexGrid4.RowHeight(0) = 600
Frm107.MSFlexGrid4.FormatString = "No.|<No.|<No. ID|<Tarikh|<No. Rujukan|<Supplier / Kilang|<Berat Asal (g)|<Mutu (%)|<Nett (g)|<Modal (RM)"

Frm107.MSFlexGrid4.ColWidth(0) = 0 'No.
Frm107.MSFlexGrid4.ColWidth(1) = 600 'No.
Frm107.MSFlexGrid4.ColWidth(2) = 0 'No. ID
Frm107.MSFlexGrid4.ColWidth(3) = 1200 'Tarikh
Frm107.MSFlexGrid4.ColWidth(4) = 1200 'No. Rujukan
Frm107.MSFlexGrid4.ColWidth(5) = 5000 'Supplier / Kilang

Frm107.MSFlexGrid4.ColWidth(6) = 1500 'Berat Asal (g)
Frm107.MSFlexGrid4.ColAlignment(6) = 7

Frm107.MSFlexGrid4.ColWidth(7) = 1000 'Mutu (%)
Frm107.MSFlexGrid4.ColAlignment(7) = 7

Frm107.MSFlexGrid4.ColWidth(8) = 1500 'Nett (g)
Frm107.MSFlexGrid4.ColAlignment(8) = 7

Frm107.MSFlexGrid4.ColWidth(9) = 2000 'Modal (RM)
Frm107.MSFlexGrid4.ColAlignment(9) = 7
'#### Header senarai description #### - End
End Sub
Sub Frm107_report_statement()
'on error resume next
Dim Frm107_LM_TOTAL_PAGE As Double
Dim TM As Date
Dim TA As Date

Frm107_PAGE_SIZE = 30
Frm107_LM_TOTAL_PAGE = 0
x = 0

Frm107.L36_Text = 0
Frm107.L37_Text = "0.00 g"
Frm107.L60_Text = "0.00"

If Frm107.L38_Text = 0 Or Frm107.L38_Text = 1 Then

    If Frm107.L38_Text = 1 Then '0 : Carian ikut supplier sahaja , 1 : Carian ikut tarikh , 2 : Carian ikut No. Rujukan
        TM = Frm107.L39_Text 'Tarikh mula
        TA = Frm107.L40_Text 'Tarikh akhir
    End If
    If Frm107.L41_Text = "Semua supplier" Then
        Frm107_LM_SEARCH_1 = Null
        Frm107_LM_SEARCH_1_LOGIC = "<>"
    Else
        Frm107_LM_SEARCH_1 = Frm107.L41_Text
        Frm107_LM_SEARCH_1_LOGIC = "="
    End If
    
End If

LM_START_ROW = Frm107.L34_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm107_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm107.L35_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm107_PAGE_SIZE
        End If
    End If
End If

Frm107_LM_PAGE_FOUND = 0

'#### Header Report ###
If Frm107.L38_Text = 0 Then Frm107.L42_Text = "Senarai rekod hantaran tukaran barang dengan supplier / kilang. Supplier [" & Frm107.L41_Text & "]." 'Report Header
If Frm107.L38_Text = 1 Then Frm107.L42_Text = "Senarai rekod hantaran tukaran barang dengan supplier / kilang. Supplier [" & Frm107.L41_Text & "] dari " & TM & " hingga " & TA 'Report Header
If Frm107.L38_Text = 2 Then Frm107.L42_Text = "Senarai rekod hantaran tukaran barang dengan supplier / kilang. No. Rujukan [" & Frm107.L41_Text & "]." 'Report Header

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm107.L38_Text = 0 Then rs.Open "select * from 57_form_out where status='" & 1 & "' AND nama_kedai " & Frm107_LM_SEARCH_1_LOGIC & "'" & Frm107_LM_SEARCH_1 & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm107_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If Frm107.L38_Text = 1 Then rs.Open "select * from 57_form_out where status='" & 1 & "' AND nama_kedai " & Frm107_LM_SEARCH_1_LOGIC & "'" & Frm107_LM_SEARCH_1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm107_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If Frm107.L38_Text = 2 Then rs.Open "select * from 57_form_out where status='" & 1 & "' AND no_statement='" & Frm107.L41_Text & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm107_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm107_LM_PAGE_FOUND = 0 Then
        If Frm107.L35_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm107.L32_Text = Frm107.L32_Text + 1 'Paparan Page ke-xxx
                Frm107_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm107.L32_Text) Then
                    If Frm107.L32_Text <> 1 Then
                        Frm107.L32_Text = Frm107.L32_Text - 1 'Paparan Page ke-xxx
                        Frm107_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm107.L32_Text - 1) * Frm107_PAGE_SIZE) + x
    Frm107.MSFlexGrid4.Rows = x + 1
    Frm107.MSFlexGrid4.TextMatrix(x, 0) = x 'No.
    Frm107.MSFlexGrid4.TextMatrix(x, 1) = Y 'No.
    Frm107.MSFlexGrid4.ColAlignment(1) = 4
    Frm107.MSFlexGrid4.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!tarikh) Then Frm107.MSFlexGrid4.TextMatrix(x, 3) = rs!tarikh 'Tarikh

    If Not IsNull(rs!no_statement) Then Frm107.MSFlexGrid4.TextMatrix(x, 4) = rs!no_statement 'No. Rujukan

    If Not IsNull(rs!nama_kedai) Then Frm107.MSFlexGrid4.TextMatrix(x, 5) = rs!nama_kedai 'Supplier / Kilang
    
    If Not IsNull(rs!berat_before) Then 'Berat Asal (g)
        Frm107.MSFlexGrid4.TextMatrix(x, 6) = Format(rs!berat_before, "#,##0.00 g")
    Else
        Frm107.MSFlexGrid4.TextMatrix(x, 6) = "0.00 g"
    End If
    
    If Not IsNull(rs!Conversion) Then 'Mutu (%)
        Frm107.MSFlexGrid4.TextMatrix(x, 7) = rs!Conversion
    Else
        Frm107.MSFlexGrid4.TextMatrix(x, 7) = "0.00"
    End If
    
    If Not IsNull(rs!berat_after) Then 'Nett (g)
        Frm107.MSFlexGrid4.TextMatrix(x, 8) = Format(rs!berat_after, "#,##0.00 g")
    Else
        Frm107.MSFlexGrid4.TextMatrix(x, 8) = "0.00 g"
    End If
    
    If Not IsNull(rs!modal) Then 'Modal
        Frm107.MSFlexGrid4.TextMatrix(x, 9) = Format(rs!modal, "#,##0.00")
    Else
        Frm107.MSFlexGrid4.TextMatrix(x, 9) = "0.00"
    End If

    rs.MoveNext
Wend

rs.Close
Set rs = Nothing


'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm107.L38_Text = 0 Then rs.Open "select COUNT(ID) from 57_form_out where status='" & 1 & "' AND nama_kedai " & Frm107_LM_SEARCH_1_LOGIC & "'" & Frm107_LM_SEARCH_1 & "'", cn, adOpenKeyset, adLockOptimistic
If Frm107.L38_Text = 1 Then rs.Open "select COUNT(ID) from 57_form_out where status='" & 1 & "' AND nama_kedai " & Frm107_LM_SEARCH_1_LOGIC & "'" & Frm107_LM_SEARCH_1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic
If Frm107.L38_Text = 2 Then rs.Open "select COUNT(ID) from 57_form_out where status='" & 1 & "' AND no_statement='" & Frm107.L41_Text & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm107_LM_TOTAL_PAGE = Format(rs(0) / Frm107_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm107_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm107_LM_PAGE = Split(Frm107_LM_TOTAL_PAGE, ".")(0)
        Frm107_LM_PAGE_LEBIHAN = Split(Frm107_LM_TOTAL_PAGE, ".")(1)
        
        If Frm107_LM_PAGE_LEBIHAN <> "00" Then
            Frm107.L33_Text = Frm107_LM_PAGE + 1
        Else
            Frm107.L33_Text = Frm107_LM_PAGE
        End If
        
    Else
    
        Frm107.L33_Text = Frm107_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm107.L33_Text = 0
    End If
Else
    Frm107.L33_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm107.L33_Text = vbNullString Then
    Frm107.L33_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah bilangan statement ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm107.L38_Text = 0 Then rs.Open "select COUNT(ID) from 57_form_out where status='" & 1 & "' AND nama_kedai " & Frm107_LM_SEARCH_1_LOGIC & "'" & Frm107_LM_SEARCH_1 & "'", cn, adOpenKeyset, adLockOptimistic
If Frm107.L38_Text = 1 Then rs.Open "select COUNT(ID) from 57_form_out where status='" & 1 & "' AND nama_kedai " & Frm107_LM_SEARCH_1_LOGIC & "'" & Frm107_LM_SEARCH_1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic
If Frm107.L38_Text = 2 Then rs.Open "select COUNT(ID) from 57_form_out where status='" & 1 & "' AND no_statement='" & Frm107.L41_Text & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm107.L36_Text = rs(0)

rs.Close
Set rs = Nothing
'### Jumlah bilangan description ### - End

'### Jumlah berat description ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm107.L38_Text = 0 Then rs.Open "select SUM(berat_after) from 57_form_out where status='" & 1 & "' AND nama_kedai " & Frm107_LM_SEARCH_1_LOGIC & "'" & Frm107_LM_SEARCH_1 & "'", cn, adOpenKeyset, adLockOptimistic
If Frm107.L38_Text = 1 Then rs.Open "select SUM(berat_after) from 57_form_out where status='" & 1 & "' AND nama_kedai " & Frm107_LM_SEARCH_1_LOGIC & "'" & Frm107_LM_SEARCH_1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic
If Frm107.L38_Text = 2 Then rs.Open "select SUM(berat_after) from 57_form_out where status='" & 1 & "' AND no_statement='" & Frm107.L41_Text & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm107.L37_Text = Format(rs(0), "#,##0.00 g")

rs.Close
Set rs = Nothing
'### Jumlah berat description ### - End

'### Jumlah modal ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm107.L38_Text = 0 Then rs.Open "select SUM(modal) from 57_form_out where status='" & 1 & "' AND nama_kedai " & Frm107_LM_SEARCH_1_LOGIC & "'" & Frm107_LM_SEARCH_1 & "'", cn, adOpenKeyset, adLockOptimistic
If Frm107.L38_Text = 1 Then rs.Open "select SUM(modal) from 57_form_out where status='" & 1 & "' AND nama_kedai " & Frm107_LM_SEARCH_1_LOGIC & "'" & Frm107_LM_SEARCH_1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic
If Frm107.L38_Text = 2 Then rs.Open "select SUM(modal) from 57_form_out where status='" & 1 & "' AND no_statement='" & Frm107.L41_Text & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm107.L60_Text = Format(rs(0), "#,##0.00")

rs.Close
Set rs = Nothing
'### Jumlah modal ### - End

If x <> 0 Then
    Frm107.L34_Text = LM_START_ROW 'Titik Pencarian Data
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    'If Frm107_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm107.L35_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm107.L35_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub Frm107_report_description_header()
'on error resume next
'#### Header senarai description #### - Start
Frm107.MSFlexGrid5.Clear
Frm107.MSFlexGrid5.Rows = 1
Frm107.MSFlexGrid5.RowHeight(0) = 600
Frm107.MSFlexGrid5.FormatString = "No.|<No.|<No. ID|<Description|<Purity|<Berat Asal (g)|<Mutu|<Berat 999.9 (g)"

Frm107.MSFlexGrid5.ColWidth(0) = 0 'No.
Frm107.MSFlexGrid5.ColWidth(1) = 600 'No.
Frm107.MSFlexGrid5.ColWidth(2) = 0 'No. ID
Frm107.MSFlexGrid5.ColWidth(3) = 6500 'Description
Frm107.MSFlexGrid5.ColWidth(4) = 1200 'Purity
Frm107.MSFlexGrid5.ColWidth(5) = 1500 'Berat Asal (g)
Frm107.MSFlexGrid5.ColWidth(6) = 1000 'Mutu
Frm107.MSFlexGrid5.ColWidth(7) = 1500 'Berat 999.9 (g)
'#### Header senarai description #### - End
End Sub
Sub Frm107_report_description()
'on error resume next
Dim Frm107_LM_TOTAL_PAGE As Double

Frm107_LM_NO_RUJUKAN = vbNullString
Frm107_PAGE_SIZE = 7
Frm107_LM_TOTAL_PAGE = 0
Frm107_LM_MODE = 0 '0 : Barang trade in , 1 : Barang potong
x = 0

Frm107.L48_Text = 0
Frm107.L49_Text = "0.00 g"

LM_START_ROW = Frm107.L46_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm107_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm107.L47_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm107_PAGE_SIZE
        End If
    End If
End If

Frm107_LM_PAGE_FOUND = 0

'### Carian no. rujukan sistem ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 57_form_out where no_statement='" & G_No_STATMENT_FORM & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!no_rujukan) Then Frm107_LM_NO_RUJUKAN = rs!no_rujukan 'No. rujukan sistem
End If

rs.Close
Set rs = Nothing
'### Carian no. rujukan sistem ### - End

If Frm107_LM_NO_RUJUKAN <> vbNullString Then

    Frm107.L43_Text = "Senarai maklumat terperinci dari statement " & G_No_STATMENT_FORM 'Header

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 58_form_out_list where status='" & 1 & "' AND no_rujukan='" & Frm107_LM_NO_RUJUKAN & "' order by id_desc ASC LIMIT " & LM_START_ROW & "," & Frm107_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        x = x + 1
        If Frm107_LM_PAGE_FOUND = 0 Then
            If Frm107.L47_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
                If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                    Frm107.L44_Text = Frm107.L44_Text + 1 'Paparan Page ke-xxx
                    Frm107_LM_PAGE_FOUND = 1
                ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                    If IsNumeric(Frm107.L44_Text) Then
                        If Frm107.L44_Text <> 1 Then
                            Frm107.L44_Text = Frm107.L44_Text - 1 'Paparan Page ke-xxx
                            Frm107_LM_PAGE_FOUND = 1
                        End If
                    End If
                End If
            End If
        End If
        Y = ((Frm107.L44_Text - 1) * Frm107_PAGE_SIZE) + x
        Frm107.MSFlexGrid5.Rows = x + 1
        Frm107.MSFlexGrid5.TextMatrix(x, 0) = x 'No.
        Frm107.MSFlexGrid5.TextMatrix(x, 1) = Y 'No.
        Frm107.MSFlexGrid5.ColAlignment(1) = 4
        Frm107.MSFlexGrid5.TextMatrix(x, 2) = rs!ID 'No. ID

        If Not IsNull(rs!Description) Then Frm107.MSFlexGrid5.TextMatrix(x, 3) = rs!Description 'Description
        If Not IsNull(rs!purity) Then Frm107.MSFlexGrid5.TextMatrix(x, 4) = rs!purity 'Purity
        Frm107.MSFlexGrid5.ColAlignment(4) = 4
        
        If Not IsNull(rs!berat_before) Then 'Berat asal (g)
            Frm107.MSFlexGrid5.TextMatrix(x, 5) = Format(rs!berat_before, "#,##0.00 g")
        Else
            Frm107.MSFlexGrid5.TextMatrix(x, 5) = "0.00 g"
        End If
        Frm107.MSFlexGrid5.ColAlignment(5) = 4
        
        If Not IsNull(rs!Conversion) Then 'Mutu
            Frm107.MSFlexGrid5.TextMatrix(x, 6) = rs!Conversion
        Else
            Frm107.MSFlexGrid5.TextMatrix(x, 6) = "0.00"
        End If
        Frm107.MSFlexGrid5.ColAlignment(6) = 4
        
        If Not IsNull(rs!berat_after) Then 'Berat 999.9
            Frm107.MSFlexGrid5.TextMatrix(x, 7) = Format(rs!berat_after, "#,##0.00 g")
        Else
            Frm107.MSFlexGrid5.TextMatrix(x, 7) = "0.00 g"
        End If
        Frm107.MSFlexGrid5.ColAlignment(7) = 4
        Frm107.MSFlexGrid5.RowHeight(x) = 700
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing

    '### Jumlah Data ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select COUNT(ID) from 58_form_out_list where status='" & 1 & "' AND no_rujukan='" & Frm107_LM_NO_RUJUKAN & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        Frm107_LM_TOTAL_PAGE = Format(rs(0) / Frm107_PAGE_SIZE, "0.00") 'Jumlah Page
        
        'Periksa Samada ada titik perpuluhan atau tidak
        If InStr(1, Frm107_LM_TOTAL_PAGE, ".") <> 0 Then
        
            Frm107_LM_PAGE = Split(Frm107_LM_TOTAL_PAGE, ".")(0)
            Frm107_LM_PAGE_LEBIHAN = Split(Frm107_LM_TOTAL_PAGE, ".")(1)
            
            If Frm107_LM_PAGE_LEBIHAN <> "00" Then
                Frm107.L45_Text = Frm107_LM_PAGE + 1
            Else
                Frm107.L45_Text = Frm107_LM_PAGE
            End If
            
        Else
        
            Frm107.L45_Text = Frm107_LM_TOTAL_PAGE
            
        End If
    
        If rs(0) = vbNullString Then
            Frm107.L45_Text = 0
        End If
    Else
        Frm107.L45_Text = 0
    End If

    rs.Close
    Set rs = Nothing
    
    If Frm107.L45_Text = vbNullString Then
        Frm107.L45_Text = 0
    End If
    '### Jumlah Data ### - End
    
    '### Jumlah bilangan description ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select COUNT(ID) from 58_form_out_list where status='" & 1 & "' AND no_rujukan='" & Frm107_LM_NO_RUJUKAN & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not IsNull(rs(0)) Then Frm107.L48_Text = rs(0)
    
    rs.Close
    Set rs = Nothing
    '### Jumlah bilangan description ### - End
    
    '### Jumlah berat description ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select SUM(berat_after) from 58_form_out_list where status='" & 1 & "' AND no_rujukan='" & Frm107_LM_NO_RUJUKAN & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not IsNull(rs(0)) Then Frm107.L49_Text = Format(rs(0), "#,##0.00 g")
    
    rs.Close
    Set rs = Nothing
    '### Jumlah berat description ### - End
    
    If x <> 0 Then
        Frm107.L46_Text = LM_START_ROW 'Titik Pencarian Data
    Else
    '    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
        'If Frm107_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    End If
    
    If x <> 0 Then
        Frm107.L47_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Else
        Frm107.L47_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    End If

End If
End Sub
Sub Frm107_report_senarai_hantar_header()
'on error resume next
'#### Header senarai description #### - Start

Frm107.MSFlexGrid6.Clear
Frm107.MSFlexGrid6.Rows = 1
Frm107.MSFlexGrid6.RowHeight(0) = 600
Frm107.MSFlexGrid6.FormatString = "No.|<No.|<No. ID|<No. Siri Produk|<Purity|<Berat (g)"

Frm107.MSFlexGrid6.ColWidth(0) = 0 'No.
Frm107.MSFlexGrid6.ColWidth(1) = 600 'No.
Frm107.MSFlexGrid6.ColWidth(2) = 0 'No. ID
Frm107.MSFlexGrid6.ColWidth(3) = 1500 'No. Siri Produk
Frm107.MSFlexGrid6.ColWidth(4) = 1200 'Purity
Frm107.MSFlexGrid6.ColWidth(5) = 1200 'Berat (g)
'#### Header senarai description #### - End
End Sub
Sub Frm107_report_senarai_hantar()
'on error resume next
Dim Frm107_LM_TOTAL_PAGE As Double

Frm107_LM_NO_RUJUKAN = vbNullString
Frm107_PAGE_SIZE = 34
Frm107_LM_TOTAL_PAGE = 0
x = 0

Frm107.L55_Text = 0
Frm107.L56_Text = "0.00 g"

LM_START_ROW = Frm107.L53_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm107_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm107.L54_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm107_PAGE_SIZE
        End If
    End If
'ElseIf GM_NEXT_PREV = 2 Then
'    If LM_START_ROW = -1 Then
'        LM_START_ROW = 0
'        Frm107.L51_Text = 1
'    End If
End If

'### Carian no. rujukan sistem ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 57_form_out where no_statement='" & G_No_STATMENT_FORM & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!no_rujukan) Then Frm107_LM_NO_RUJUKAN = rs!no_rujukan 'No. rujukan sistem
End If

rs.Close
Set rs = Nothing
'### Carian no. rujukan sistem ### - End

If Frm107_LM_NO_RUJUKAN <> vbNullString Then

    Frm107_LM_PAGE_FOUND = 0
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 59_form_out_item_list where no_rujukan='" & Frm107_LM_NO_RUJUKAN & "' AND id_rujukan='" & G_No_RUJUKAN_FORM & "' AND status='" & 1 & "' order by no_siri_produk ASC LIMIT " & LM_START_ROW & "," & Frm107_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        x = x + 1
        If Frm107_LM_PAGE_FOUND = 0 Then
            If Frm107.L54_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
                If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                    Frm107.L51_Text = Frm107.L51_Text + 1 'Paparan Page ke-xxx
                    Frm107_LM_PAGE_FOUND = 1
                ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                    If IsNumeric(Frm107.L51_Text) Then
                        If Frm107.L51_Text <> 1 Then
                            Frm107.L51_Text = Frm107.L51_Text - 1 'Paparan Page ke-xxx
                            Frm107_LM_PAGE_FOUND = 1
                        End If
                    End If
                End If
            End If
        End If
        Y = ((Frm107.L51_Text - 1) * Frm107_PAGE_SIZE) + x
        Frm107.MSFlexGrid6.Rows = x + 1
        Frm107.MSFlexGrid6.TextMatrix(x, 0) = x 'No.
        Frm107.MSFlexGrid6.TextMatrix(x, 1) = Y 'No.
        Frm107.MSFlexGrid6.ColAlignment(1) = 4
        Frm107.MSFlexGrid6.TextMatrix(x, 2) = rs!ID 'No. ID
    
        If Not IsNull(rs!no_siri_Produk) Then Frm107.MSFlexGrid6.TextMatrix(x, 3) = rs!no_siri_Produk 'No. siri produk
        If Not IsNull(rs!purity) Then Frm107.MSFlexGrid6.TextMatrix(x, 4) = rs!purity 'Purity
        Frm107.MSFlexGrid6.ColAlignment(4) = 4
        If Not IsNull(rs!Berat) Then Frm107.MSFlexGrid6.TextMatrix(x, 5) = Format(rs!Berat, "#,##0.00 g") 'Berat
        Frm107.MSFlexGrid6.ColAlignment(5) = 4
        
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
    '### Jumlah Data ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select COUNT(ID) from 59_form_out_item_list where no_rujukan='" & Frm107_LM_NO_RUJUKAN & "' AND id_rujukan='" & G_No_RUJUKAN_FORM & "' AND status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        Frm107_LM_TOTAL_PAGE = Format(rs(0) / Frm107_PAGE_SIZE, "0.00") 'Jumlah Page
        
        'Periksa Samada ada titik perpuluhan atau tidak
        If InStr(1, Frm107_LM_TOTAL_PAGE, ".") <> 0 Then
        
            Frm107_LM_PAGE = Split(Frm107_LM_TOTAL_PAGE, ".")(0)
            Frm107_LM_PAGE_LEBIHAN = Split(Frm107_LM_TOTAL_PAGE, ".")(1)
            
            If Frm107_LM_PAGE_LEBIHAN <> "00" Then
                Frm107.L52_Text = Frm107_LM_PAGE + 1
            Else
                Frm107.L52_Text = Frm107_LM_PAGE
            End If
            
        Else
        
            Frm107.L52_Text = Frm107_LM_TOTAL_PAGE
            
        End If
    
        If rs(0) = vbNullString Then
            Frm107.L52_Text = 0
        End If
    Else
        Frm107.L52_Text = 0
    End If
    
    rs.Close
    Set rs = Nothing
    
    If Frm107.L52_Text = vbNullString Then
        Frm107.L52_Text = 0
    End If
    '### Jumlah Data ### - End
    
    '### Jumlah bilangan barang keseluruhan ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select COUNT(ID) from 59_form_out_item_list where no_rujukan='" & Frm107_LM_NO_RUJUKAN & "' AND id_rujukan='" & G_No_RUJUKAN_FORM & "' AND status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not IsNull(rs(0)) Then Frm107.L55_Text = rs(0)
    
    rs.Close
    Set rs = Nothing
    '### Jumlah bilangan barang keseluruhan ### - End
    
    '### Jumlah berat barang keseluruhan ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select SUM(berat) from 59_form_out_item_list where no_rujukan='" & Frm107_LM_NO_RUJUKAN & "' AND id_rujukan='" & G_No_RUJUKAN_FORM & "' AND status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not IsNull(rs(0)) Then
        Frm107.L56_Text = Format(rs(0), "#,##0.00 g")
    End If
    
    rs.Close
    Set rs = Nothing
    '### Jumlah berat barang keseluruhan ### - End
    
    If x <> 0 Then
        Frm107.L53_Text = LM_START_ROW 'Titik Pencarian Data
    Else
    '    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
        'If Frm107_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    End If
    
    If x <> 0 Then
        Frm107.L54_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Else
        Frm107.L54_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    End If
End If
End Sub
Sub Frm107_senarai_barang_hantar_excel()
'on error resume next
Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook

Frm107_LM_NO_RUJUKAN = vbNullString

Note = "Sistem mungkin akan mengambil masa untuk mengeluarkan report ini." & vbCrLf & _
        "" & vbCrLf & _
        "Sila tunggu sehingga sistem siap mengeluarkan report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then

    '### Carian no. rujukan sistem ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 57_form_out where no_statement='" & G_No_STATMENT_FORM & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!no_rujukan) Then Frm107_LM_NO_RUJUKAN = rs!no_rujukan 'No. rujukan sistem
    End If
    
    rs.Close
    Set rs = Nothing
    '### Carian no. rujukan sistem ### - End
    
    If Frm107_LM_NO_RUJUKAN <> vbNullString Then
    
        x = 0
        
        Set xlObject = New Excel.Application
        Set xlWB = xlObject.Workbooks.Add
                   
        'xlObject.Visible = True
        With xlObject.ActiveWorkbook.ActiveSheet
        
            .Cells.VerticalAlignment = xlCenter
            .Columns("A").ColumnWidth = 5 'No.
            .Columns("B").ColumnWidth = 30 'No. Siri Produk
            .Columns("C").ColumnWidth = 20 'Purity
            .Columns("D").ColumnWidth = 20 'Berat (g)
    
            '### Maklumat kedai ### - Start
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
            
            .Cells(7, 1) = "Senarai barang yang dihantar kepada supplier / kilang dari No. Rujukan " & G_No_STATMENT_FORM 'Header Report
            
            .Cells(8, 1) = "No."
            .Cells(8, 2) = "No. Siri Produk"
            .Cells(8, 3) = "Purity"
            .Cells(8, 4) = "Berat (g)"
            
            For i = 1 To 4
                .Cells(8, i).HorizontalAlignment = xlCenter
                .Cells(8, i).Interior.ColorIndex = 15
                .Cells(8, i).WrapText = True
                .Cells(8, i).Borders.LineStyle = xlContinuous
            Next i
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 59_form_out_item_list where no_rujukan='" & Frm107_LM_NO_RUJUKAN & "' AND status='" & 1 & "' order by id_rujukan", cn, adOpenKeyset, adLockOptimistic
                
            While rs.EOF = False
            
                x = x + 1
                .Cells(8 + x, 1) = x 'No.
                .Cells(8 + x, 1).HorizontalAlignment = xlCenter
                
                If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 2) = rs!no_siri_Produk 'No. Siri Produk
                .Cells(8 + x, 2).HorizontalAlignment = xlCenter
                
                If Not IsNull(rs!purity) Then .Cells(8 + x, 3) = rs!purity 'Purity
                .Cells(8 + x, 3).HorizontalAlignment = xlCenter
                
                .Cells(8 + x, 4).HorizontalAlignment = xlCenter
                If Not IsNull(rs!Berat) Then
                    .Cells(8 + x, 4) = rs!Berat 'Berat (g)
                Else
                    .Cells(8 + x, 4) = "0.00" 'Berat (g)
                End If
                .Cells(8 + x, 4).NumberFormat = "#,##0.00"
                
                For Col = 1 To 4
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
            rs.Open "select SUM(berat) from 59_form_out_item_list where no_rujukan='" & Frm107_LM_NO_RUJUKAN & "' AND status='" & 1 & "' order by id_rujukan", cn, adOpenKeyset, adLockOptimistic
    
            If Not IsNull(rs(0)) Then
                .Cells(8 + Y, 1) = "Berat Keseluruhan : " & Format(rs(0), "#,##0.00 g")
            Else
                .Cells(8 + Y, 1) = "Berat Keseluruhan : " & "0.00 g"
            End If
            
            rs.Close
            Set rs = Nothing
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
End If
End Sub
Sub Frm107_visible_component_1()
'on error resume next
Frm107.CMD4.Visible = True
Frm107.CMD18.Visible = False
Frm107.CMD19.Visible = False
End Sub
Sub Frm107_visible_component_2()
'on error resume next
Frm107.CMD7.Visible = True
Frm107.CMD20.Visible = False
Frm107.CMD21.Visible = False
End Sub
Sub Frm107_jurujual()
'On Error Resume Next
If MDI_frm1.L3_Text <> vbNullString Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        Frm107.CBB5 = rs!Samaran & "  |  " & rs!NoPekerja
        
    Else
        
        Frm107.CBB5.AddItem "" & "  |  " & rs!Samaran
        Frm107.CBB5 = "" & "  |  " & rs!Samaran
        
    End If
    
    rs.Close
    Set rs = Nothing
    
    If G_LOCK_JURUJUAL = "YES" Then
    
        Frm107.CBB5.Enabled = False
        Frm107.CBB5.BackColor = &H8000000A

    Else
    
        Frm107.CBB5.Enabled = True
        Frm107.CBB5.BackColor = &HFFFFFF

    End If

End If
End Sub
