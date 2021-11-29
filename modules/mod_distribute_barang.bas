Attribute VB_Name = "mod_distribute_barang"
Sub Frm108_cmd_visible_1()
'on error resume next
'Data baru pendaftaran cawangan

Frm108.CMD1.Visible = True
Frm108.CMD2.Visible = False
Frm108.CMD3.Visible = False
End Sub
Sub Frm108_cmd_invisible_1()
'on error resume next
'Edit data cawangan

Frm108.CMD1.Visible = False
Frm108.CMD2.Visible = True
Frm108.CMD3.Visible = True
End Sub
Sub Frm108_cmd_visible_2()
'on error resume next
'Data baru penghantaran barang kepada cawangan / kedai

Frm108.CMD6.Visible = True
Frm108.CMD7.Visible = False
Frm108.CMD8.Visible = False
End Sub
Sub Frm108_cmd_invisible_2()
'on error resume next
'Edit data penghantaran barang kepada cawangan / kedai

Frm108.CMD6.Visible = False
Frm108.CMD7.Visible = True
Frm108.CMD8.Visible = True
End Sub
Sub Frm108_cmd_visible_3()
'on error resume next
'Data baru pulangan barang oleh cawangan / kedai

Frm108.CMD21.Visible = True
Frm108.CMD22.Visible = False
Frm108.CMD23.Visible = False
End Sub
Sub Frm108_cmd_invisible_3()
'on error resume next
'Data baru pulangan barang oleh cawangan / kedai

Frm108.CMD21.Visible = False
Frm108.CMD22.Visible = True
Frm108.CMD23.Visible = True
End Sub
Sub Frm108_cawangan_initial_setting()
'on error resume next
'Reset semua component untuk pendaftaran cawangan

Frm108.TB4 = vbNullString
End Sub
Sub Frm108_initial_setting()
'on error resume next
'Setting position bagi setiap picture

Frm108.Pic1.Left = 120
Frm108.Pic1.Top = 240
Frm108.Pic2.Left = 120
Frm108.Pic2.Top = 240
Frm108.Pic3.Left = 120
Frm108.Pic3.Top = 240
Frm108.Pic6.Left = 120
Frm108.Pic6.Top = 240
Frm108.Pic7.Left = 120
Frm108.Pic7.Top = 240

Frm108.Pic1.Visible = False
Frm108.Pic2.Visible = False
Frm108.Pic3.Visible = False
Frm108.Pic6.Visible = False
Frm108.Pic7.Visible = False
End Sub
Sub Frm108_initial_setting2()
'on error resume next
'Setting position bagi setiap dalam report hantaran
Frm108.Pic4.Left = 6720
Frm108.Pic4.Top = 120
Frm108.Pic5.Left = 6720
Frm108.Pic5.Top = 120

Frm108.Pic4.Visible = False
Frm108.Pic5.Visible = False
End Sub
Sub Frm108_hantaran_initial_setting()
'on error resume next
'Reset semua component untuk hantaran barang ke cawangan

Frm108.TB1 = vbNullString 'Ambilan : Nama
Frm108.TB2 = vbNullString 'Ambilan : No. Kad Pengenalan
Frm108.TB3 = vbNullString 'Ambilan : No. Telefon
Frm108.TB8 = vbNullString 'Pulangan : No. Kad Pengenalan
Frm108.TB9 = vbNullString 'Pulangan : No. Telefon
Frm108.TB10 = vbNullString 'Pulangan : No. Siri Produk
Frm108.TB12 = vbNullString 'No. Perjanjian A
Frm108.TB13 = vbNullString 'No. Perjanjian B
Frm108.TB14 = vbNullString 'Harga Jualan
End Sub
Sub Frm108_hantaran_initial_setting2()
'on error resume next
'Reset semua component untuk hantaran barang ke cawangan sebelum dan selepas data disimpan

Frm108.L1_Text = 0 '0 : Data baru , 1 : Edit
Frm108.L39_Text = 0 '0 : Data baru , 1 : Edit
Frm108.L17_Text = 0
Frm108.L18_Text = "0.00 g"

'Frm108.L19_Text.Visible = False

'### update no rujukan sistem ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If rs!Default1 = "Default" Then
        If Not IsNull(rs!no_rujukan_agihan) Then 'No. rujukan sistem (Agihan)
            If IsNumeric(rs!no_rujukan_agihan) Then
                Frm108.L12_Text = rs!no_rujukan_agihan
            Else
                Frm108.L12_Text = 1
            End If
        Else
            Frm108.L12_Text = 1
        End If
        
        If Not IsNull(rs!no_rujukan_pulangan) Then 'No. rujukan sistem (Pulangan)
            If IsNumeric(rs!no_rujukan_pulangan) Then
                Frm108.L40_Text = rs!no_rujukan_pulangan
            Else
                Frm108.L40_Text = 1
            End If
        Else
            Frm108.L40_Text = 1
        End If
    Else
        Frm108.L12_Text = 1
        Frm108.L40_Text = 1
    End If

End If

rs.Close
Set rs = Nothing
'### update no rujukan sistem ### - End

'### Senarai cawangan ### - Start
Frm108.CBB1.Clear
Frm108.CBB4.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 62_senarai_cawangan where status='" & 1 & "' order by cawangan ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!cawangan) Then Frm108.CBB1.AddItem rs!cawangan
    If Not IsNull(rs!cawangan) Then Frm108.CBB4.AddItem rs!cawangan
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'### Senarai cawangan ### - End

'### Senarai pekerja ### - Start
Frm108.CBB2.Clear
Frm108.CBB5.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where status='" & "Aktif" & "' AND InvestorSmall = 0 AND InvestorBig = 0", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    If rs!user_level = 1 Or rs!user_level = 2 Or rs!user_level = 3 Or rs!user_level = G_LEVEL_USER Then
        Frm108.CBB2.AddItem rs!Samaran & "  |  " & rs!NoPekerja
        Frm108.CBB5.AddItem rs!Samaran & "  |  " & rs!NoPekerja
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'### Senarai pekerja ### - End

'###Padam Table 65_agihan_barang_temp### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "TRUNCATE TABLE " & G_AGIHAN_TEMP & ""

Set rs = cn.Execute(strsql)
Set rs = Nothing
'###Padam Table 65_agihan_barang_temp### - End

'###Padam Table 67_pulangan_barang_temp### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "TRUNCATE TABLE " & G_PULANGAN_TEMP & ""

Set rs = cn.Execute(strsql)
Set rs = Nothing
'###Padam Table 67_pulangan_barang_temp### - End

Call frm108_jurujual
End Sub
Sub Frm108_hantaran_initial_setting3()
'on error resume next
'### Senarai cawangan ### - Start
Frm108.CBB3.Clear
Frm108.CBB3.AddItem "Semua Cawangan"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 62_senarai_cawangan where status='" & 1 & "' order by cawangan ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!cawangan) Then Frm108.CBB3.AddItem rs!cawangan
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'### Senarai cawangan ### - End

Frm108.CBB3 = "Semua Cawangan"

Frm108.TB6 = vbNullString 'Report : No. Rujukan
Frm108.TB7 = vbNullString 'Report : No. Siri Produk
End Sub
Sub Frm108_report_initial_setting()
'on error resume next
'### Senarai cawangan ### - Start
Frm108.CBB6.Clear
Frm108.CBB6.AddItem "Semua Cawangan"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 62_senarai_cawangan where status='" & 1 & "' order by cawangan ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!cawangan) Then Frm108.CBB6.AddItem rs!cawangan
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'### Senarai cawangan ### - End

Frm108.CBB6 = "Semua Cawangan"

Frm108.CBB7.Clear
Frm108.CBB7.AddItem "Semua Jenis Report"
Frm108.CBB7.AddItem "Agihan"
Frm108.CBB7.AddItem "Pulangan"
'Frm108.CBB7.AddItem "Dijual"
Frm108.CBB7.AddItem "Belum Dipulangkan"
Frm108.CBB7 = "Semua Jenis Report"
End Sub
Sub Frm108_one_time_reset()
'on error resume next
'Digunakan hanya untuk reset data asas sistem
'Kemungkinan sekali sahaja digunakan semasa activate kan form

Frm108.L6_Text = 0 'Senarai cawangan : Paparan page
Frm108.L7_Text = 0 'Senarai cawangan : Jumlah page
Frm108.L8_Text = 0 'Senarai cawangan : Titik carian data (default = -1)
Frm108.L9_Text = 0 'Senarai cawangan : Flag page terakhir
Frm108.L13_Text = 0 'Senarai barang agihan : Paparan page
Frm108.L14_Text = 0 'Senarai barang agihan : Jumlah page
Frm108.L15_Text = 0 'Senarai barang agihan : Titik carian data (default = -1)
Frm108.L16_Text = 0 'Senarai barang agihan : Flag page terakhir
Frm108.L17_Text = 0 'Bilangan barang agihan
Frm108.L18_Text = "0.00 g" 'Jumlah berat barang agihan
Frm108.L10_Text = 0 'Bilangan cawangan
Frm108.L11_Text = 0 'Memory : No. ID cawangan
Frm108.L12_Text = 0 'Ambilan : No. Rujukan Sistem
Frm108.L20_Text = 0 'Memori : Jenis report ( 0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh , 2:  No.rujukan , 3:  No.siri produk )
Frm108.L21_Text = vbNullString 'Memori : Tarikh mula
Frm108.L22_Text = vbNullString 'Memori : Tarikh akhir
Frm108.L23_Text = vbNullString 'Memori : Supplier / No rujukan / No. siri produk
Frm108.L24_Text = "Maklumat statement agihan barang." 'Header : maklumat statement
Frm108.L25_Text = 0 'Report : bilangan
Frm108.L26_Text = 0 'Report statement : Paparan page
Frm108.L27_Text = 0 'Report statement : Jumlah page
Frm108.L28_Text = 0 'Report statement : Titik carian data (default = -1)
Frm108.L29_Text = 0 'Report statement : Flag page terakhir
Frm108.L30_Text = "Senarai barang yang diagihkan." 'Header : Senarai barang yang diagihkan
Frm108.L31_Text = 0 'Report senarai barang yang diagih : bilangan
Frm108.L32_Text = "0.00 g" 'Report senarai barang yang diagih : bilangan
Frm108.L33_Text = 0 'Report senarai barang yang diagih : Paparan page
Frm108.L34_Text = 0 'Report senarai barang yang diagih : Jumlah page
Frm108.L35_Text = 0 'Report senarai barang yang diagih : Titik carian data (default = -1)
Frm108.L36_Text = 0 'Report senarai barang yang diagih : Flag page terakhir
Frm108.L40_Text = 0 'Pulangan : No. Rujukan Sistem
Frm108.L41_Text = "Senarai barang yang akan dipulangkan / dijual." 'Header : Senarai barang yang akan dipulangkan
Frm108.L44_Text = 0 'Senarai barang pulang : Paparan page
Frm108.L45_Text = 0 'Senarai barang pulang : Jumlah page
Frm108.L46_Text = 0 'Senarai barang pulang : Titik carian data (default = -1)
Frm108.L47_Text = 0 'Senarai barang pulang : Flag page terakhir
Frm108.L42_Text = 0 'Report senarai barang yang dipulangkan : bilangan
Frm108.L43_Text = "0.00 g" 'Report senarai barang yang dipulangkan : Jumlah berat
Frm108.TB5 = vbNullString
Frm108.TB11 = vbNullString
Frm108.L48_Text = 0 'Memory : Jenis report , 0 : Agihan , 1 : Pulangan
Frm108.L49_Text = 0 'Memory : Jenis report inventory
Frm108.L52_Text = "Report inventori agihan / pulangan barang oleh cawangan." 'Header : header report inventory
Frm108.L53_Text = 0 'Report inventory : Bilangan (Report inventory)
Frm108.L54_Text = "0.00 g" 'Report inventory : Jumlah berat (Report inventory)
Frm108.L55_Text = 0 'Report inventory : Paparan page
Frm108.L56_Text = 0 'Report inventory : Jumlah page
Frm108.L57_Text = 0 'Report inventory : Titik carian data (default = -1)
Frm108.L58_Text = 0 'Report inventory : Flag page terakhir

Frm108.L50_Text = vbNullString 'Memory : Cawangan
Frm108.L51_Text = vbNullString 'Memory : Jenis Report
Frm108.L60_Text = vbNullString 'Memory : Tarikh mula
Frm108.L61_Text = vbNullString 'Memory : tarikh akhir
Frm108.L62_Text = "Maklumat Agihan" 'Caption : Maklumat agihan / Maklumat pulangan

Frm108.L64_Text = 0 'Report inventory : Flag page terakhir
Frm108.L65_Text = "0.00 g" 'Report inventory : Flag page terakhir
Frm108.L66_Text = "RM 0.00" 'Report inventory : Flag page terakhir

Frm108.L63_Text.Visible = False
Frm108.L64_Text.Visible = False
Frm108.L65_Text.Visible = False
Frm108.L66_Text.Visible = False

Frm108.CB4 = 1 'Jenis report : Agihan
Frm108.CB5 = 0 'Jenis report : Pulangan
Frm108.CB7 = 1 'Jenis Pulangan : Pulangan
Frm108.CB8 = 0 'Jenis Pulangan : Dijual

Frm108.DTPicker1 = DateTime.Date
Frm108.DTPicker2 = DateTime.Date
Frm108.DTPicker3 = DateTime.Date
Frm108.DTPicker4 = DateTime.Date
Frm108.DTPicker5 = DateTime.Date
Frm108.DTPicker6 = DateTime.Date

Frm108.TB12 = vbNullString 'No. Perjanjian A
Frm108.TB13 = vbNullString 'No. Perjanjian B
Frm108.TB14 = vbNullString 'Harga Jualan

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Default1 = "Default" Then
        If Not IsNull(rs!ScannerMode) Then
        
            If rs!ScannerMode = 1 Then
                Frm108.CB1 = 1
                Frm108.CB3 = 1
            Else
                Frm108.CB1 = 0
                Frm108.CB3 = 0
            End If
        Else
            Frm108.CB1 = 0
            Frm108.CB3 = 0
        End If
        
    End If
End If

rs.Close
Set rs = Nothing
End Sub
Sub Frm108_senarai_cawangan_header()
'on error resume next
'#### Header senarai cawangan #### - Start
Frm108.MSFlexGrid1.Clear
Frm108.MSFlexGrid1.Rows = 1
Frm108.MSFlexGrid1.RowHeight(0) = 600
Frm108.MSFlexGrid1.FormatString = "No.|<No.|<No. ID|<Cawangan"

Frm108.MSFlexGrid1.ColWidth(0) = 0 'No.
Frm108.MSFlexGrid1.ColWidth(1) = 600 'No.
Frm108.MSFlexGrid1.ColWidth(2) = 0 'No. ID
Frm108.MSFlexGrid1.ColWidth(3) = 6700 'Cawangan
'#### Header senarai cawangan #### - End
End Sub
Sub Frm108_senarai_cawangan()
'on error resume next
Dim Frm108_LM_TOTAL_PAGE As Double

Frm108_PAGE_SIZE = 33
Frm108_LM_TOTAL_PAGE = 0
x = 0

Frm108.L10_Text = 0

LM_START_ROW = Frm108.L8_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm108_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm108.L9_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm108_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        Frm108.L6_Text = 1
    End If
End If

Frm108_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 62_senarai_cawangan where status='" & 1 & "' order by cawangan ASC LIMIT " & LM_START_ROW & "," & Frm108_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm108_LM_PAGE_FOUND = 0 Then
        If Frm108.L9_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm108.L6_Text = Frm108.L6_Text + 1 'Paparan Page ke-xxx
                Frm108_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm108.L6_Text) Then
                    If Frm108.L6_Text <> 1 Then
                        Frm108.L6_Text = Frm108.L6_Text - 1 'Paparan Page ke-xxx
                        Frm108_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm108.L6_Text - 1) * Frm108_PAGE_SIZE) + x
    Frm108.MSFlexGrid1.Rows = x + 1
    Frm108.MSFlexGrid1.TextMatrix(x, 0) = x 'No.
    Frm108.MSFlexGrid1.TextMatrix(x, 1) = Y 'No.
    Frm108.MSFlexGrid1.ColAlignment(1) = 4
    Frm108.MSFlexGrid1.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!cawangan) Then Frm108.MSFlexGrid1.TextMatrix(x, 3) = rs!cawangan 'Nama Cawangan
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 62_senarai_cawangan where status='" & 1 & "' order by cawangan ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm108_LM_TOTAL_PAGE = Format(rs(0) / Frm108_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm108_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm108_LM_PAGE = Split(Frm108_LM_TOTAL_PAGE, ".")(0)
        Frm108_LM_PAGE_LEBIHAN = Split(Frm108_LM_TOTAL_PAGE, ".")(1)
        
        If Frm108_LM_PAGE_LEBIHAN <> "00" Then
            Frm108.L7_Text = Frm108_LM_PAGE + 1
        Else
            Frm108.L7_Text = Frm108_LM_PAGE
        End If
        
    Else
    
        Frm108.L7_Text = Frm108_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm108.L7_Text = 0
    End If
Else
    Frm108.L7_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm108.L7_Text = vbNullString Then
    Frm108.L7_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah bilangan barang keseluruhan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 62_senarai_cawangan where status='" & 1 & "' order by cawangan ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm108.L10_Text = rs(0)

rs.Close
Set rs = Nothing
'### Jumlah bilangan barang keseluruhan ### - End

If x <> 0 Then
    Frm108.L8_Text = LM_START_ROW 'Titik Pencarian Data
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    'If Frm108_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm108.L9_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm108.L9_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub Frm108_periksa_data_barang()
'on error resume next
Dim rs1 As ADODB.Recordset

Frm108_LM_NO_SIRI = UCase(Frm108.TB5)
Frm108.TB5 = vbNullString
DATA_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from " & G_AGIHAN_TEMP & " where no_siri_produk='" & Frm108_LM_NO_SIRI & "' AND (status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "' OR status='" & 4 & "')", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    
    If rs!Status = 0 Then
        
        rs.Delete
        rs.Update
        
    ElseIf rs!Status = 4 Then

        rs!Status = 5
        rs.Update
        
    ElseIf rs!Status = 6 Then
    
    Else
    
        MsgBox "Barang dengan nombor siri [" & Frm108_LM_NO_SIRI & "] telah dimasukkan ke dalam senarai sebelum ini.", vbInformation, "Info"
        Frm108.TB5.SetFocus
        
        rs.Close
        Set rs = Nothing
        
        Exit Sub
    
    End If
    
End If

rs.Close
Set rs = Nothing

'### Carian status barang dalam #data_database ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from Data_Database where no_siri_Produk='" & Frm108_LM_NO_SIRI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!StatusItem = "10" Then
    
        Set rs1 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs1.Open "select * from " & G_AGIHAN_TEMP & "", cn, adOpenKeyset, adLockOptimistic
        
        rs1.AddNew
        If Not IsNull(rs!no_siri_Produk) Then rs1!no_siri_Produk = rs!no_siri_Produk 'No. siri produk
        If Not IsNull(rs!kategori_Produk) Then rs1!kategori_Produk = rs!kategori_Produk 'Kategori Produk
        If Not IsNull(rs!kod_Purity) Then rs1!purity = rs!kod_Purity 'Kod Purity
        If Not IsNull(rs!beza_berat) Then rs1!Berat = rs!beza_berat
        If Frm108.L1_Text = 0 Then '0 : Data baru , 1 : Edit Then
            rs1!Status = 1
        ElseIf Frm108.L1_Text = 1 Then '0 : Data baru , 1 : Edit
            rs1!Status = 3
        End If
        DATA_FOUND = 1
        
        rs1.Update
        
        rs1.Close
        Set rs1 = Nothing
        
    ElseIf rs!StatusItem = "11" Then
        MsgBox "Item Ini Telah Dijual. No. Siri Produk [" & Frm108_LM_NO_SIRI & "]", vbExclamation, "Info"
        
        Frm108.TB5.SetFocus
    ElseIf rs!StatusItem = "12" Then
        MsgBox "Item Ini Telah Dijual Secara Potong. No. Siri Produk [" & Frm108_LM_NO_SIRI & "]", vbExclamation, "Info"
        
        Frm108.TB5.SetFocus
    ElseIf rs!StatusItem = "13" Then
        MsgBox "Item Ini Telah Dijual Secara Potong. No. Siri Produk [" & Frm108_LM_NO_SIRI & "]", vbExclamation, "Info"
        
        Frm108.TB5.SetFocus
    ElseIf rs!StatusItem = "14" Or rs!StatusItem = "21" Or rs!StatusItem = "22" Then
        MsgBox "Item Ini Telah Ditempah Oleh Pelanggan. No. Siri Produk [" & Frm108_LM_NO_SIRI & "]", vbExclamation, "Info"
        
        Frm108.TB5.SetFocus
    ElseIf rs!StatusItem = "15" Or rs!StatusItem = "19" Or rs!StatusItem = "20" Then
        MsgBox "Item Ini Telah Dibeli Secara Ansuran. No. Siri Produk [" & Frm108_LM_NO_SIRI & "]", vbExclamation, "Info"
        
        Frm108.TB5.SetFocus
    ElseIf rs!StatusItem = "16" Then
        MsgBox "Item Ini Telah Dihantar Ke Ar-Rahnu. No. Siri Produk [" & Frm108_LM_NO_SIRI & "]", vbExclamation, "Info"
        
        Frm108.TB5.SetFocus
    ElseIf rs!StatusItem = "17" Then
        MsgBox "Item Ini Telah Dijual Secara ETA. No. Siri Produk [" & Frm108_LM_NO_SIRI & "]", vbExclamation, "Info"
        
        Frm108.TB5.SetFocus
    ElseIf rs!StatusItem = "23" Then
        MsgBox "Item Ini Telah Dihantar Ke Supplier/Kilang. No. Siri Produk [" & Frm108_LM_NO_SIRI & "]", vbExclamation, "Info"

        Frm108.TB5.SetFocus
    ElseIf rs!StatusItem = "24" Then
        MsgBox "Item Ini Telah Dihantar Ke Supplier/Kilang. No. Siri Produk [" & Frm108_LM_NO_SIRI & "]", vbExclamation, "Info"

        Frm108.TB5.SetFocus
    ElseIf rs!StatusItem = "25" Then
        MsgBox "Item Ini Telah Diagihkan Ke Cawangan. No. Siri Produk [" & Frm108_LM_NO_SIRI & "]", vbExclamation, "Info"

        Frm108.TB5.SetFocus
    ElseIf rs!StatusItem = "26" Then
        MsgBox "Item Ini Telah Dijual Oleh Cawangan. No. Siri Produk [" & Frm108_LM_NO_SIRI & "]", vbExclamation, "Info"

        Frm108.TB5.SetFocus
    ElseIf rs!StatusItem = "0" Then
        MsgBox "Item Ini Telah Dipadamkan Dari Sistem. No. Siri Produk [" & Frm108_LM_NO_SIRI & "]", vbExclamation, "Info"
        
        Frm108.TB5.SetFocus
    ElseIf rs!StatusItem = "27" Or rs!StatusItem = "28" Then
        MsgBox "Item Ini Telah Dijual Dari Menu GDN. No. Siri Produk [" & Frm108_LM_NO_SIRI & "]", vbExclamation, "Info"

        Frm108.TB5.SetFocus
    ElseIf rs!StatusItem = "29" Then
        MsgBox "Item Ini Telah Diubah Status Kepada Hilang , Dicuri Dan Sebagainya. No. Siri Produk [" & Frm108_LM_NO_SIRI & "]", vbExclamation, "Info"

        Frm108.TB5.SetFocus
    End If
Else
    MsgBox "Produk dengan nombor siri [" & Frm108_LM_NO_SIRI & "] tidak dijumpai.", vbExclamation, "Info"
    
    Frm108.TB5.SetFocus
End If

rs.Close
Set rs = Nothing

If DATA_FOUND = 1 Then

    GM_NEXT_PREV = 0
    
    Frm108.L15_Text = -1 'Titik Pencarian Data
    Frm108.L16_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm108.L13_Text = 0 'Paparan Page ke-xxx
    
    Call Frm108_senarai_agihan_header
    Call Frm108_senarai_agihan
    
    MsgBox "Data telah berjaya dimasukkan ke dalam senarai.", vbInformation, "Info"
    
    Frm108.TB5.SetFocus
    
End If
End Sub
Sub Frm108_periksa_data_barang2()
'on error resume next
Dim rs1 As ADODB.Recordset

Frm108_LM_NO_SIRI = UCase(Frm108.TB11)
Frm108.TB11 = vbNullString
DATA_FOUND = 0
DATA_VALIDATION = 0
DATA_INSERT = 0

'### Periksa samada barang ini sedang dipinjamakan atau tidak ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 64_agihan_barang where no_siri_produk='" & Frm108_LM_NO_SIRI & "' and status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic

If rs.EOF Then
    If Frm108.L48_Text = 0 Then
    
        MsgBox "Barang dengan nombor siri [" & Frm108_LM_NO_SIRI & "] tiada dalam rekod senarai barang yang sedang dipinjam.", vbInformation, "Info"
        Frm108.TB11.SetFocus
        
        rs.Close
        Set rs = Nothing
        
        Exit Sub
        
    ElseIf Frm108.L48_Text = 1 Then
    
        DATA_VALIDATION = 1
        
    End If

End If
    

rs.Close
Set rs = Nothing
'### Periksa samada barang ini sedang dipinjamakan atau tidak ### - End

'### Periksa samada barang ini telah dimasukkan ke dalam senarai atau tidak ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from " & G_PULANGAN_TEMP & " where no_siri_produk='" & Frm108_LM_NO_SIRI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    
    If rs!Status = 0 Then
        
        If Frm108.CB8 = 1 Then
        
            If Frm108.TB12 <> vbNullString Then 'No. Perjanjian A
                rs!no_perjanjian_a = UCase(Frm108.TB12)
            Else
                rs!no_perjanjian_a = Null
            End If
            If Frm108.TB13 <> vbNullString Then 'No. Perjanjian B
                rs!no_perjanjian_b = UCase(Frm108.TB13)
            Else
                rs!no_perjanjian_b = Null
            End If
            If Frm108.TB14 <> vbNullString Then 'Harga Jualan (RM)
                rs!harga_jualan = Format(Frm108.TB14, "0.00")
            Else
                rs!harga_jualan = Null
            End If
            If Frm108.L39_Text = 0 Then
                rs!Status = 2
            ElseIf Frm108.L39_Text = 1 Then
                rs!Status = 6
            End If
            
            rs.Update
            
        Else
        
            rs!no_perjanjian_a = Null 'No. Perjanjian A
            rs!no_perjanjian_b = Null 'No. Perjanjian B
            rs!harga_jualan = Null 'Harga Jualan (RM)
            If Frm108.L39_Text = 0 Then
                rs!Status = 1
            ElseIf Frm108.L39_Text = 1 Then
                rs!Status = 5
            End If
            
            rs.Update
            
        End If
        
        DATA_FOUND = 1
        DATA_VALIDATION = 0
        DATA_INSERT = 1
        
    ElseIf rs!Status = 9 Or rs!Status = 10 Then
    
        If Frm108.CB8 = 1 Then
        
            If Frm108.TB12 <> vbNullString Then 'No. Perjanjian A
                rs!no_perjanjian_a = UCase(Frm108.TB12)
            Else
                rs!no_perjanjian_a = Null
            End If
            If Frm108.TB13 <> vbNullString Then 'No. Perjanjian B
                rs!no_perjanjian_b = UCase(Frm108.TB13)
            Else
                rs!no_perjanjian_b = Null
            End If
            If Frm108.TB14 <> vbNullString Then 'Harga Jualan (RM)
                rs!harga_jualan = Format(Frm108.TB14, "0.00")
            Else
                rs!harga_jualan = Null
            End If
            rs!Status = 4
            rs.Update
            
        Else
        
            rs!no_perjanjian_a = Null 'No. Perjanjian A
            rs!no_perjanjian_b = Null 'No. Perjanjian B
            rs!harga_jualan = Null 'Harga Jualan (RM)
            rs!Status = 3
            rs.Update
            
        End If

        rs.Update
        DATA_FOUND = 1
        DATA_VALIDATION = 0
        DATA_INSERT = 1
        
    Else

        MsgBox "Barang dengan nombor siri [" & Frm108_LM_NO_SIRI & "] telah dimasukkan ke dalam senarai sebelum ini.", vbInformation, "Info"
        Frm108.TB11.SetFocus
        
        rs.Close
        Set rs = Nothing
        
        Exit Sub
        
    End If

End If

rs.Close
Set rs = Nothing
'### Periksa samada barang ini telah dimasukkan ke dalam senarai atau tidak ### - End

If DATA_VALIDATION = 1 Then

    MsgBox "Barang dengan nombor siri [" & Frm108_LM_NO_SIRI & "] tiada dalam rekod senarai barang yang sedang dipinjam.", vbInformation, "Info"
    Frm108.TB11.SetFocus
    
    Exit Sub
        
End If

'### Masukkan data ke dalam 67_pulangan_barang_temp ### - Start
'If DATA_FOUND = 1 Then

'### #65_agihan_barang_temp -> #64_agihan_barang ### - Start
    If DATA_INSERT = 0 Then

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        If Frm108.L39_Text = 0 Then
        
            If Frm108.CB7 = 1 Then
                strsql = "insert into " & G_PULANGAN_TEMP & "(no_rujukan_agihan,tarikh,cawangan_id,no_siri_produk,kategori_produk,purity,berat,status)" & _
                        "select no_rujukan,tarikh,cawangan_id,no_siri_produk,kategori_produk,purity,berat,1 from 64_agihan_barang WHERE status='" & 1 & "' AND no_siri_produk='" & Frm108_LM_NO_SIRI & "'"
            End If
        
            If Frm108.CB8 = 1 Then
                strsql = "insert into " & G_PULANGAN_TEMP & "(no_rujukan_agihan,tarikh,cawangan_id,no_siri_produk,kategori_produk,purity,berat,status)" & _
                        "select no_rujukan,tarikh,cawangan_id,no_siri_produk,kategori_produk,purity,berat,2 from 64_agihan_barang WHERE status='" & 1 & "' AND no_siri_produk='" & Frm108_LM_NO_SIRI & "'"
            End If
            
        End If
        
        If Frm108.L39_Text = 1 Then
        
            If Frm108.CB7 = 1 Then
                strsql = "insert into " & G_PULANGAN_TEMP & "(no_rujukan_agihan,tarikh,cawangan_id,no_siri_produk,kategori_produk,purity,berat,status)" & _
                        "select no_rujukan,tarikh,cawangan_id,no_siri_produk,kategori_produk,purity,berat,5 from 64_agihan_barang WHERE status='" & 1 & "' AND no_siri_produk='" & Frm108_LM_NO_SIRI & "'"
            End If
        
            If Frm108.CB8 = 1 Then
                strsql = "insert into " & G_PULANGAN_TEMP & "(no_rujukan_agihan,tarikh,cawangan_id,no_siri_produk,kategori_produk,purity,berat,status)" & _
                        "select no_rujukan,tarikh,cawangan_id,no_siri_produk,kategori_produk,purity,berat,6 from 64_agihan_barang WHERE status='" & 1 & "' AND no_siri_produk='" & Frm108_LM_NO_SIRI & "'"
            End If
        
        End If
  
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
        
        If Frm108.CB7 = 1 Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            
            If Frm108.L39_Text = 0 Then strsql = "UPDATE " & G_PULANGAN_TEMP & " SET no_perjanjian_a='" & UCase(Frm108.TB12) & "'," _
                                                & "no_perjanjian_b='" & UCase(Frm108.TB13) & "'," _
                                                & "harga_jualan='" & Format(Frm108.TB14, "0.00") & "'" _
                                                & "WHERE no_siri_produk='" & Frm108_LM_NO_SIRI & "' AND status = 1"
            
            If Frm108.L39_Text = 1 Then strsql = "UPDATE " & G_PULANGAN_TEMP & " SET no_perjanjian_a='" & UCase(Frm108.TB12) & "'," _
                                                & "no_perjanjian_b='" & UCase(Frm108.TB13) & "'," _
                                                & "harga_jualan='" & Format(Frm108.TB14, "0.00") & "'" _
                                                & "WHERE no_siri_produk='" & Frm108_LM_NO_SIRI & "' AND status = 5"
                                                
            Set rs = cn.Execute(strsql)
            Set rs = Nothing
        
        End If
        
        
        If Frm108.CB8 = 1 Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            
            If Frm108.L39_Text = 0 Then strsql = "UPDATE " & G_PULANGAN_TEMP & " SET no_perjanjian_a='" & UCase(Frm108.TB12) & "'," _
                                                & "no_perjanjian_b='" & UCase(Frm108.TB13) & "'," _
                                                & "harga_jualan='" & Format(Frm108.TB14, "0.00") & "'" _
                                                & "WHERE no_siri_produk='" & Frm108_LM_NO_SIRI & "' AND status = 2"
            
            If Frm108.L39_Text = 1 Then strsql = "UPDATE " & G_PULANGAN_TEMP & " SET no_perjanjian_a='" & UCase(Frm108.TB12) & "'," _
                                                & "no_perjanjian_b='" & UCase(Frm108.TB13) & "'," _
                                                & "harga_jualan='" & Format(Frm108.TB14, "0.00") & "'" _
                                                & "WHERE no_siri_produk='" & Frm108_LM_NO_SIRI & "' AND status = 6"
                                                
            Set rs = cn.Execute(strsql)
            Set rs = Nothing
        
        End If
        
    End If
'### #65_agihan_barang_temp -> #64_agihan_barang ### - End

    GM_NEXT_PREV = 0
    
    Frm108.L46_Text = -1 'Titik Pencarian Data
    Frm108.L47_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm108.L44_Text = 0 'Paparan Page ke-xxx
    
    Call Frm108_senarai_pulangan_header
    Call Frm108_senarai_pulangan
    
    MsgBox "Data telah berjaya dimasukkan ke dalam senarai.", vbInformation, "Info"

'End If
'### Masukkan data ke dalam 67_pulangan_barang_temp ### - End


Exit Sub

If DATA_FOUND = 1 Then

    GM_NEXT_PREV = 0
    
    Frm108.L46_Text = -1 'Titik Pencarian Data
    Frm108.L47_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm108.L44_Text = 0 'Paparan Page ke-xxx
    
    Call Frm108_senarai_pulangan_header
    Call Frm108_senarai_pulangan
    
    MsgBox "Data telah berjaya dimasukkan ke dalam senarai.", vbInformation, "Info"
    
End If
End Sub
Sub Frm108_senarai_agihan_header()
'on error resume next
'#### Header senarai agihan #### - Start
Frm108.MSFlexGrid2.Clear
Frm108.MSFlexGrid2.Rows = 1
Frm108.MSFlexGrid2.RowHeight(0) = 600
Frm108.MSFlexGrid2.FormatString = "No.|<No.|<No. ID|<No. Siri Produk|<Nama Produk|<Purity|<Berat (g)|<Status|<Tarikh Dipulangkan"

Frm108.MSFlexGrid2.ColWidth(0) = 0 'No.
Frm108.MSFlexGrid2.ColWidth(1) = 600 'No.
Frm108.MSFlexGrid2.ColWidth(2) = 0 'No. ID
Frm108.MSFlexGrid2.ColWidth(3) = 1300 'No. Siri Produk
Frm108.MSFlexGrid2.ColWidth(4) = 3400 'Nama Produk
Frm108.MSFlexGrid2.ColWidth(5) = 1000 'Purity
Frm108.MSFlexGrid2.ColWidth(6) = 1200 'Berat (g)
Frm108.MSFlexGrid2.ColWidth(7) = 1700 'Status
Frm108.MSFlexGrid2.ColWidth(8) = 1300 'Tarikh Dipulangkan
'#### Header senarai agihan #### - End
End Sub
Sub Frm108_senarai_agihan()
'on error resume next
Dim Frm108_LM_TOTAL_PAGE As Double

Frm108_PAGE_SIZE = 37
Frm108_LM_TOTAL_PAGE = 0
x = 0

Frm108.L17_Text = 0
Frm108.L18_Text = Format(0, "#,##0.00 g")

LM_START_ROW = Frm108.L15_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm108_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm108.L16_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm108_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        Frm108.L13_Text = 1
    End If
End If

Frm108_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from " & G_AGIHAN_TEMP & " where (status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "' OR status='" & 5 & "' OR status='" & 6 & "' OR status='" & 7 & "') order by no_siri_produk ASC LIMIT " & LM_START_ROW & "," & Frm108_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm108_LM_PAGE_FOUND = 0 Then
        If Frm108.L16_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm108.L13_Text = Frm108.L13_Text + 1 'Paparan Page ke-xxx
                Frm108_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm108.L13_Text) Then
                    If Frm108.L13_Text <> 1 Then
                        Frm108.L13_Text = Frm108.L13_Text - 1 'Paparan Page ke-xxx
                        Frm108_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm108.L13_Text - 1) * Frm108_PAGE_SIZE) + x
    Frm108.MSFlexGrid2.Rows = x + 1
    Frm108.MSFlexGrid2.TextMatrix(x, 0) = x 'No.
    Frm108.MSFlexGrid2.TextMatrix(x, 1) = Y 'No.
    Frm108.MSFlexGrid2.ColAlignment(1) = 4
    Frm108.MSFlexGrid2.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!no_siri_Produk) Then Frm108.MSFlexGrid2.TextMatrix(x, 3) = rs!no_siri_Produk 'No. Siri Produk
    Frm108.MSFlexGrid2.ColAlignment(3) = 4
    If Not IsNull(rs!kategori_Produk) Then Frm108.MSFlexGrid2.TextMatrix(x, 4) = rs!kategori_Produk 'Kategori produk
    If Not IsNull(rs!purity) Then Frm108.MSFlexGrid2.TextMatrix(x, 5) = rs!purity 'Purity
    Frm108.MSFlexGrid2.ColAlignment(5) = 4
    If Not IsNull(rs!Berat) Then Frm108.MSFlexGrid2.TextMatrix(x, 6) = Format(rs!Berat, "#,##0.00 g") 'Berat
    Frm108.MSFlexGrid2.ColAlignment(6) = 4
    If Not IsNull(rs!Status) Then
        If rs!Status = 1 Then
            Frm108.MSFlexGrid2.TextMatrix(x, 7) = "Sedia dipinjamkan"
        ElseIf rs!Status = 2 Then
            Frm108.MSFlexGrid2.TextMatrix(x, 7) = "Sedang dipinjamkan"
        ElseIf rs!Status = 3 Then
            Frm108.MSFlexGrid2.TextMatrix(x, 7) = "Sedia dipinjamkan"
        ElseIf rs!Status = 6 Then
            Frm108.MSFlexGrid2.TextMatrix(x, 7) = "SUDAH DIPULANGKAN"
        ElseIf rs!Status = 7 Then
            Frm108.MSFlexGrid2.TextMatrix(x, 7) = "SUDAH TERJUAL"
        End If
    End If
    If Not IsNull(rs!tarikh_jual) Then
        Frm108.MSFlexGrid2.TextMatrix(x, 8) = rs!tarikh_jual
    Else
        Frm108.MSFlexGrid2.TextMatrix(x, 8) = "-"
    End If
    Frm108.MSFlexGrid2.ColAlignment(8) = 4
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from " & G_AGIHAN_TEMP & " where (status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "' OR status='" & 5 & "' OR status='" & 6 & "' OR status='" & 7 & "')", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm108_LM_TOTAL_PAGE = Format(rs(0) / Frm108_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm108_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm108_LM_PAGE = Split(Frm108_LM_TOTAL_PAGE, ".")(0)
        Frm108_LM_PAGE_LEBIHAN = Split(Frm108_LM_TOTAL_PAGE, ".")(1)
        
        If Frm108_LM_PAGE_LEBIHAN <> "00" Then
            Frm108.L14_Text = Frm108_LM_PAGE + 1
        Else
            Frm108.L14_Text = Frm108_LM_PAGE
        End If
        
    Else
    
        Frm108.L14_Text = Frm108_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm108.L14_Text = 0
    End If
Else
    Frm108.L14_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm108.L14_Text = vbNullString Then
    Frm108.L14_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah bilangan barang keseluruhan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from " & G_AGIHAN_TEMP & " where (status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "' OR status='" & 5 & "' OR status='" & 6 & "' OR status='" & 7 & "')", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm108.L17_Text = rs(0)

rs.Close
Set rs = Nothing
'### Jumlah bilangan barang keseluruhan ### - End

'### Jumlah berat keseluruhan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(berat) from " & G_AGIHAN_TEMP & " where (status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "' OR status='" & 5 & "' OR status='" & 6 & "' OR status='" & 7 & "')", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm108.L18_Text = Format(rs(0), "#,##0.00 g")

rs.Close
Set rs = Nothing
'### Jumlah berat keseluruhan ### - End

If x <> 0 Then
    Frm108.L15_Text = LM_START_ROW 'Titik Pencarian Data
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    'If Frm108_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm108.L16_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm108.L16_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub Frm108_cetak_penyata_ambilan()
'on error resume next
'G_PENYATA_AMBILAN = "000007"
Frm108_LM_RUJUKAN = vbNullString
Frm108_LM_No_PEKERJA = vbNullString

'Set rs = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from default_setting where default1='" & "default" & "'", cn, adOpenKeyset, adLockOptimistic

'If Not rs.EOF Then
    
'    If Not IsNull(rs!default_printer) Then LM_PRINTER = rs!default_printer
'
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

Report72.Caption = "Ambilan stok / barang kemas oleh cawangan / pengedar"

'### Reset maklumat kedai ### - Start
Report72.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report72.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report72.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report72.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report72.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

Report72.Sections("Section5").Controls("L8").Caption = 0 'Jumlah bilangan
Report72.Sections("Section5").Controls("L9").Caption = "0.00 g" 'Jumlah berat

Report72.Sections("Section4").Controls("L10").Caption = "Ambilan stok / barang kemas oleh cawangan atau pengedar."
Report72.Sections("Section5").Controls("L11").Caption = "Tandatangan Penerima"

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report72.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report72.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report72.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report72.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report72.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Report72.Sections("Section4").Controls("L5").Caption = G_PENYATA_AMBILAN 'No. Rujukan (xxxxxx)

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 63_agihan where no_statement='" & G_PENYATA_AMBILAN & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!no_rujukan) Then Frm108_LM_RUJUKAN = rs!no_rujukan 'No. rujukan
    If Not IsNull(rs!cawangan) Then Report72.Sections("Section4").Controls("L1").Caption = rs!cawangan 'Cawangan
    If Not IsNull(rs!Nama) Then Report72.Sections("Section4").Controls("L2").Caption = rs!Nama 'Nama
    If Not IsNull(rs!no_ic) Then Report72.Sections("Section4").Controls("L3").Caption = rs!no_ic 'No. IC
    If Not IsNull(rs!no_tel) Then Report72.Sections("Section4").Controls("L4").Caption = rs!no_tel 'No. telefon
    If Not IsNull(rs!tarikh) Then Report72.Sections("Section4").Controls("L6").Caption = rs!tarikh 'Tarikh barang diambil
    If Not IsNull(rs!nama_pekerja) Then Frm108_LM_No_PEKERJA = rs!nama_pekerja 'No. Pekerja
End If

rs.Close
Set rs = Nothing
        
If Frm108_LM_RUJUKAN <> vbNullString Then

'### Maklumat pekerja ### - Start
    If Frm108_LM_No_PEKERJA <> vbNullString Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where NoPekerja='" & Frm108_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!Samaran) Then Report72.Sections("Section4").Controls("L7").Caption = rs!Samaran
        End If
        
        rs.Close
        Set rs = Nothing
        
    End If
'### Maklumat pekerja ### - End

'### Jumlah bilangan barang keseluruhan ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select COUNT(ID) from 64_agihan_barang where status='" & 1 & "' AND no_rujukan='" & Frm108_LM_RUJUKAN & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not IsNull(rs(0)) Then Report72.Sections("Section5").Controls("L8").Caption = rs(0)
    
    rs.Close
    Set rs = Nothing
'### Jumlah bilangan barang keseluruhan ### - End

'### Jumlah berat keseluruhan ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select SUM(berat) from 64_agihan_barang where status='" & 1 & "' AND no_rujukan='" & Frm108_LM_RUJUKAN & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not IsNull(rs(0)) Then Report72.Sections("Section5").Controls("L9").Caption = Format(rs(0), "#,##0.00 g")
    
    rs.Close
    Set rs = Nothing
'### Jumlah berat keseluruhan ### - End

    '### Paparan statement ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    'rs.Open "select * from 64_agihan_barang where no_rujukan='" & Frm108_LM_RUJUKAN & "' AND status='" & 1 & "' order by no_siri_produk ASC", cn, adOpenKeyset, adLockOptimistic
    rs.Open "select * from 64_agihan_barang where no_rujukan='" & Frm108_LM_RUJUKAN & "' AND (status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "') order by no_siri_produk ASC", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        Set Report72.DataSource = rs
        Report72.Show
        rs.MoveNext
    Wend
    
    'rs.Close
    Set rs = Nothing
    '### Paparan statement ### - End
End If
End Sub
Sub Frm108_cetak_penyata_pulangan()
'on error resume next
'G_PENYATA_PULANGAN = "000007"
Frm108_LM_RUJUKAN = vbNullString
Frm108_LM_No_PEKERJA = vbNullString

'Set rs = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from default_setting where default1='" & "default" & "'", cn, adOpenKeyset, adLockOptimistic

'If Not rs.EOF Then
    
'    If Not IsNull(rs!default_printer) Then LM_PRINTER = rs!default_printer
'
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

Report73.Caption = "Pulang barang oleh cawangan"

'### Reset maklumat kedai ### - Start
Report73.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report73.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report73.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report73.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report73.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

Report73.Sections("Section5").Controls("L8").Caption = 0 'Jumlah bilangan (Pulang)
Report73.Sections("Section5").Controls("L9").Caption = "0.00 g" 'Jumlah berat (Pulang)

Report73.Sections("Section5").Controls("L13").Caption = 0 'Jumlah bilangan (Jual)
Report73.Sections("Section5").Controls("L14").Caption = "0.00 g" 'Jumlah berat (Jual)
Report73.Sections("Section5").Controls("L15").Caption = "RM 0.00" 'Jumlah Harga (Jual)

Report73.Sections("Section4").Controls("L10").Caption = "Pulangan barang kemas oleh cawangan."
Report73.Sections("Section5").Controls("L11").Caption = "Tandatangan Penghantar"

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report73.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report73.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report73.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report73.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report73.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Report73.Sections("Section4").Controls("L5").Caption = G_PENYATA_PULANGAN 'No. Rujukan (xxxxxx)

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 68_pulangan where no_statement='" & G_PENYATA_PULANGAN & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!no_rujukan) Then Frm108_LM_RUJUKAN = rs!no_rujukan 'No. rujukan
    If Not IsNull(rs!cawangan) Then Report73.Sections("Section4").Controls("L1").Caption = rs!cawangan 'Cawangan
    If Not IsNull(rs!Nama) Then Report73.Sections("Section4").Controls("L2").Caption = rs!Nama 'Nama
    If Not IsNull(rs!no_ic) Then Report73.Sections("Section4").Controls("L3").Caption = rs!no_ic 'No. IC
    If Not IsNull(rs!no_tel) Then Report73.Sections("Section4").Controls("L4").Caption = rs!no_tel 'No. telefon
    If Not IsNull(rs!tarikh) Then Report73.Sections("Section4").Controls("L6").Caption = rs!tarikh 'Tarikh barang diambil
    If Not IsNull(rs!nama_pekerja) Then Frm108_LM_No_PEKERJA = rs!nama_pekerja 'No. Pekerja
End If

rs.Close
Set rs = Nothing
        
If Frm108_LM_RUJUKAN <> vbNullString Then

'### Maklumat pekerja ### - Start
    If Frm108_LM_No_PEKERJA <> vbNullString Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where NoPekerja='" & Frm108_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!Samaran) Then Report73.Sections("Section4").Controls("L7").Caption = rs!Samaran
        End If
        
        rs.Close
        Set rs = Nothing
        
    End If
'### Maklumat pekerja ### - End

'### Jumlah bilangan barang keseluruhan ### - Start (Barang yang dipulangkan)
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select COUNT(ID) from 69_pulangan_barang where status='" & 1 & "' AND no_rujukan='" & Frm108_LM_RUJUKAN & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not IsNull(rs(0)) Then Report73.Sections("Section5").Controls("L8").Caption = rs(0)
    
    rs.Close
    Set rs = Nothing
'### Jumlah bilangan barang keseluruhan ### - End (Barang yang dipulangkan)

'### Jumlah berat keseluruhan ### - Start (Barang yang dipulangkan)
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select SUM(berat) from 69_pulangan_barang where status='" & 1 & "' AND no_rujukan='" & Frm108_LM_RUJUKAN & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not IsNull(rs(0)) Then Report73.Sections("Section5").Controls("L9").Caption = Format(rs(0), "#,##0.00 g")
    
    rs.Close
    Set rs = Nothing
'### Jumlah berat keseluruhan ### - End (Barang yang dipulangkan)

'### Jumlah bilangan barang keseluruhan ### - Start (Barang yang dijual)
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select COUNT(ID) from 69_pulangan_barang where status='" & 2 & "' AND no_rujukan='" & Frm108_LM_RUJUKAN & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not IsNull(rs(0)) Then Report73.Sections("Section5").Controls("L13").Caption = rs(0)
    
    rs.Close
    Set rs = Nothing
'### Jumlah bilangan barang keseluruhan ### - End (Barang yang dijual)

'### Jumlah berat keseluruhan ### - Start (Barang yang dijual)
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select SUM(berat) from 69_pulangan_barang where status='" & 2 & "' AND no_rujukan='" & Frm108_LM_RUJUKAN & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not IsNull(rs(0)) Then Report73.Sections("Section5").Controls("L14").Caption = Format(rs(0), "#,##0.00 g")
    
    rs.Close
    Set rs = Nothing
'### Jumlah berat keseluruhan ### - End (Barang yang dijual)

'### Jumlah harga keseluruhan ### - Start (Barang yang dijual)
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select SUM(harga_jualan) from 69_pulangan_barang where status='" & 2 & "' AND no_rujukan='" & Frm108_LM_RUJUKAN & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not IsNull(rs(0)) Then Report73.Sections("Section5").Controls("L15").Caption = "RM " & Format(rs(0), "#,##0.00")
    
    rs.Close
    Set rs = Nothing
'### Jumlah harga keseluruhan ### - End (Barang yang dijual)

    '### Paparan statement ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    'rs.Open "select * from 64_agihan_barang where no_rujukan='" & Frm108_LM_RUJUKAN & "' AND status='" & 1 & "' order by no_siri_produk ASC", cn, adOpenKeyset, adLockOptimistic
    rs.Open "select * from 69_pulangan_barang where no_rujukan='" & Frm108_LM_RUJUKAN & "' AND (status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "') order by no_siri_produk ASC", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        Set Report73.DataSource = rs
        Report73.Show
        rs.MoveNext
    Wend
    
    'rs.Close
    Set rs = Nothing
    '### Paparan statement ### - End
End If
End Sub
Sub Frm108_senarai_agihan_barang_header()
'on error resume next
'#### Header senarai cawangan #### - Start
Frm108.MSFlexGrid3.Clear
Frm108.MSFlexGrid3.Rows = 1
Frm108.MSFlexGrid3.RowHeight(0) = 600
Frm108.MSFlexGrid3.FormatString = "No.|<No.|<No. ID|<Tarikh|<No. Rujukan|<Cawangan|<Nama PIC"

Frm108.MSFlexGrid3.ColWidth(0) = 0 'No.
Frm108.MSFlexGrid3.ColWidth(1) = 600 'No.
Frm108.MSFlexGrid3.ColWidth(2) = 0 'No. ID
Frm108.MSFlexGrid3.ColWidth(3) = 1500 'Tarikh
Frm108.MSFlexGrid3.ColWidth(4) = 1500 'No. Rujukan
Frm108.MSFlexGrid3.ColWidth(5) = 4000 'Cawangan
Frm108.MSFlexGrid3.ColWidth(6) = 6500 'Nama PIC
'#### Header senarai cawangan #### - End
End Sub
Sub Frm108_senarai_agihan_barang()
'on error resume next
Dim Frm108_LM_TOTAL_PAGE As Double
Dim TM As Date
Dim TA As Date

Frm108_PAGE_SIZE = 36
Frm108_LM_TOTAL_PAGE = 0
x = 0

Frm108.L25_Text = 0

If Frm108.L20_Text = 1 Then '0 : Carian ikut supplier sahaja , 1 : Carian ikut tarikh , 2 : Carian ikut No. Rujukan
    TM = Frm108.L21_Text 'Tarikh mula
    TA = Frm108.L22_Text 'Tarikh akhir
End If
If Frm108.L23_Text = "Semua Cawangan" Then
    Frm108_LM_SEARCH_1 = Null
    Frm108_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm108_LM_SEARCH_1 = Frm108.L23_Text
    Frm108_LM_SEARCH_1_LOGIC = "="
End If

LM_START_ROW = Frm108.L28_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm108_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm108.L29_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm108_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        Frm108.L26_Text = 1
    End If
End If

Frm108_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm108.L20_Text = 0 Then rs.Open "select * from 63_agihan where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND status='" & 1 & "' order by no_rujukan ASC LIMIT " & LM_START_ROW & "," & Frm108_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If Frm108.L20_Text = 1 Then rs.Open "select * from 63_agihan where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND status='" & 1 & "' order by no_rujukan ASC LIMIT " & LM_START_ROW & "," & Frm108_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If Frm108.L20_Text = 2 Then rs.Open "select * from 63_agihan where no_statement='" & Frm108.L23_Text & "' AND status='" & 1 & "'order by no_rujukan ASC LIMIT " & LM_START_ROW & "," & Frm108_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm108_LM_PAGE_FOUND = 0 Then
        If Frm108.L29_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm108.L26_Text = Frm108.L26_Text + 1 'Paparan Page ke-xxx
                Frm108_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm108.L26_Text) Then
                    If Frm108.L26_Text <> 1 Then
                        Frm108.L26_Text = Frm108.L26_Text - 1 'Paparan Page ke-xxx
                        Frm108_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm108.L26_Text - 1) * Frm108_PAGE_SIZE) + x
    Frm108.MSFlexGrid3.Rows = x + 1
    Frm108.MSFlexGrid3.TextMatrix(x, 0) = x 'No.
    Frm108.MSFlexGrid3.TextMatrix(x, 1) = Y 'No.
    Frm108.MSFlexGrid3.ColAlignment(1) = 4
    Frm108.MSFlexGrid3.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!tarikh) Then Frm108.MSFlexGrid3.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    Frm108.MSFlexGrid3.ColAlignment(3) = 4
    
    If Not IsNull(rs!no_statement) Then Frm108.MSFlexGrid3.TextMatrix(x, 4) = rs!no_statement 'No. rujukan (No. statement)
    Frm108.MSFlexGrid3.ColAlignment(4) = 4
    
    If Not IsNull(rs!cawangan) Then Frm108.MSFlexGrid3.TextMatrix(x, 5) = rs!cawangan 'Cawangan
    If Not IsNull(rs!Nama) Then Frm108.MSFlexGrid3.TextMatrix(x, 6) = rs!Nama 'Nama PIC
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm108.L20_Text = 0 Then rs.Open "select COUNT(ID) from 63_agihan where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND status='" & 1 & "' order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic
If Frm108.L20_Text = 1 Then rs.Open "select COUNT(ID) from 63_agihan where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND status='" & 1 & "' order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic
If Frm108.L20_Text = 2 Then rs.Open "select COUNT(ID) from 63_agihan where no_statement='" & Frm108.L23_Text & "' AND status='" & 1 & "' order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm108_LM_TOTAL_PAGE = Format(rs(0) / Frm108_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm108_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm108_LM_PAGE = Split(Frm108_LM_TOTAL_PAGE, ".")(0)
        Frm108_LM_PAGE_LEBIHAN = Split(Frm108_LM_TOTAL_PAGE, ".")(1)
        
        If Frm108_LM_PAGE_LEBIHAN <> "00" Then
            Frm108.L27_Text = Frm108_LM_PAGE + 1
        Else
            Frm108.L27_Text = Frm108_LM_PAGE
        End If
        
    Else
    
        Frm108.L27_Text = Frm108_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm108.L27_Text = 0
    End If
Else
    Frm108.L27_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm108.L27_Text = vbNullString Then
    Frm108.L27_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah bilangan barang keseluruhan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm108.L20_Text = 0 Then rs.Open "select COUNT(ID) from 63_agihan where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND status='" & 1 & "' order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic
If Frm108.L20_Text = 1 Then rs.Open "select COUNT(ID) from 63_agihan where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND status='" & 1 & "' order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic
If Frm108.L20_Text = 2 Then rs.Open "select COUNT(ID) from 63_agihan where no_statement='" & Frm108.L23_Text & "' AND status='" & 1 & "' order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm108.L25_Text = rs(0)

rs.Close
Set rs = Nothing
'### Jumlah bilangan barang keseluruhan ### - End

If x <> 0 Then
    Frm108.L28_Text = LM_START_ROW 'Titik Pencarian Data
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    'If Frm108_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm108.L29_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm108.L29_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub Frm108_senarai_agihan_barang_detail_header()
'on error resume next
'#### Header senarai agihan #### - Start
Frm108.MSFlexGrid4.Clear
Frm108.MSFlexGrid4.Rows = 1
Frm108.MSFlexGrid4.RowHeight(0) = 600
Frm108.MSFlexGrid4.FormatString = "No.|<No.|<No. ID|<No. Siri Produk|<Nama Produk|<Purity|<Berat (g)|<Status|<Tarikh Dipulangkan"

Frm108.MSFlexGrid4.ColWidth(0) = 0 'No.
Frm108.MSFlexGrid4.ColWidth(1) = 600 'No.
Frm108.MSFlexGrid4.ColWidth(2) = 0 'No. ID
Frm108.MSFlexGrid4.ColWidth(3) = 1300 'No. Siri Produk
Frm108.MSFlexGrid4.ColWidth(4) = 3600 'Nama Produk
Frm108.MSFlexGrid4.ColWidth(5) = 1200 'Purity
Frm108.MSFlexGrid4.ColWidth(6) = 1200 'Berat (g)
Frm108.MSFlexGrid4.ColWidth(7) = 1900 'Status
Frm108.MSFlexGrid4.ColWidth(8) = 1300 'Tarikh Dipulangkan
'#### Header senarai agihan #### - End
End Sub
Sub Frm108_senarai_agihan_barang_detail()
'on error resume next
Dim Frm108_LM_TOTAL_PAGE As Double

Frm108_PAGE_SIZE = 36
Frm108_LM_TOTAL_PAGE = 0
x = 0

Frm108.L31_Text = 0
Frm108.L32_Text = Format(0, "#,##0.00 g")

LM_START_ROW = Frm108.L35_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm108_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm108.L36_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm108_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        Frm108.L33_Text = 1
    End If
End If

Frm108_LM_PAGE_FOUND = 0

Frm108.L30_Text = "Senarai barang yang diagihkan dari No. Rujukan [" & Format(Frm108.L37_Text, "000000") & "]" 'Header : Senarai barang yang diagihkan

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 64_agihan_barang where no_rujukan='" & Frm108.L37_Text & "' AND (status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "' OR status='" & 5 & "') order by no_siri_produk ASC LIMIT " & LM_START_ROW & "," & Frm108_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm108_LM_PAGE_FOUND = 0 Then
        If Frm108.L36_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm108.L33_Text = Frm108.L33_Text + 1 'Paparan Page ke-xxx
                Frm108_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm108.L33_Text) Then
                    If Frm108.L33_Text <> 1 Then
                        Frm108.L33_Text = Frm108.L33_Text - 1 'Paparan Page ke-xxx
                        Frm108_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm108.L33_Text - 1) * Frm108_PAGE_SIZE) + x
    Frm108.MSFlexGrid4.Rows = x + 1
    Frm108.MSFlexGrid4.TextMatrix(x, 0) = x 'No.
    Frm108.MSFlexGrid4.TextMatrix(x, 1) = Y 'No.
    Frm108.MSFlexGrid4.ColAlignment(1) = 4
    Frm108.MSFlexGrid4.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!no_siri_Produk) Then Frm108.MSFlexGrid4.TextMatrix(x, 3) = rs!no_siri_Produk 'No. Siri Produk
    Frm108.MSFlexGrid4.ColAlignment(3) = 4
    If Not IsNull(rs!kategori_Produk) Then Frm108.MSFlexGrid4.TextMatrix(x, 4) = rs!kategori_Produk 'Kategori produk
    If Not IsNull(rs!purity) Then Frm108.MSFlexGrid4.TextMatrix(x, 5) = rs!purity 'Purity
    Frm108.MSFlexGrid4.ColAlignment(5) = 4
    If Not IsNull(rs!Berat) Then Frm108.MSFlexGrid4.TextMatrix(x, 6) = Format(rs!Berat, "#,##0.00 g") 'Berat
    Frm108.MSFlexGrid4.ColAlignment(6) = 4
    If Not IsNull(rs!Status) Then
        If rs!Status = 1 Then
            Frm108.MSFlexGrid4.TextMatrix(x, 7) = "Sedang dipinjamkan"
        ElseIf rs!Status = 2 Then
            Frm108.MSFlexGrid4.TextMatrix(x, 7) = "Sudah Dipulangkan"
        ElseIf rs!Status = 3 Then
            Frm108.MSFlexGrid4.TextMatrix(x, 7) = "Sudah Dijual"
        End If
    End If
    If Not IsNull(rs!tarikh_jual) Then
        Frm108.MSFlexGrid4.TextMatrix(x, 8) = rs!tarikh_jual
    Else
        Frm108.MSFlexGrid4.TextMatrix(x, 8) = "-"
    End If
    Frm108.MSFlexGrid4.ColAlignment(8) = 4
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 64_agihan_barang where no_rujukan='" & Frm108.L37_Text & "' AND (status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "' OR status='" & 5 & "') order by no_siri_produk ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm108_LM_TOTAL_PAGE = Format(rs(0) / Frm108_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm108_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm108_LM_PAGE = Split(Frm108_LM_TOTAL_PAGE, ".")(0)
        Frm108_LM_PAGE_LEBIHAN = Split(Frm108_LM_TOTAL_PAGE, ".")(1)
        
        If Frm108_LM_PAGE_LEBIHAN <> "00" Then
            Frm108.L34_Text = Frm108_LM_PAGE + 1
        Else
            Frm108.L34_Text = Frm108_LM_PAGE
        End If
        
    Else
    
        Frm108.L34_Text = Frm108_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm108.L34_Text = 0
    End If
Else
    Frm108.L34_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm108.L34_Text = vbNullString Then
    Frm108.L34_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah bilangan barang keseluruhan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 64_agihan_barang where no_rujukan='" & Frm108.L37_Text & "' AND (status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "' OR status='" & 5 & "') order by no_siri_produk ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm108.L31_Text = rs(0)

rs.Close
Set rs = Nothing
'### Jumlah bilangan barang keseluruhan ### - End

'### Jumlah berat keseluruhan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(berat) from 64_agihan_barang where no_rujukan='" & Frm108.L37_Text & "' AND (status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "' OR status='" & 5 & "') order by no_siri_produk ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm108.L32_Text = Format(rs(0), "#,##0.00 g")

rs.Close
Set rs = Nothing
'### Jumlah berat keseluruhan ### - End

If x <> 0 Then
    Frm108.L35_Text = LM_START_ROW 'Titik Pencarian Data
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    'If Frm108_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm108.L36_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm108.L36_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub Frm108_senarai_pulangan_barang_detail_header()
'on error resume next
'#### Header senarai agihan #### - Start
Frm108.MSFlexGrid4.Clear
Frm108.MSFlexGrid4.Rows = 1
Frm108.MSFlexGrid4.RowHeight(0) = 600
Frm108.MSFlexGrid4.FormatString = "No.|<No.|<No. ID|<No. Siri Produk|<Nama Produk|<Purity|<Berat (g)|<Status|<Tarikh Dipulangkan"
Frm108.MSFlexGrid4.FormatString = "No.|<No.|<No. ID|<No. Siri Produk|<Nama Produk|<Berat (g)|<Status|<No. Perjanjian A|<No. Perjanjian B|<Harga Jualan (RM)"

Frm108.MSFlexGrid4.ColWidth(0) = 0 'No.
Frm108.MSFlexGrid4.ColWidth(1) = 600 'No.
Frm108.MSFlexGrid4.ColWidth(2) = 0 'No. ID
Frm108.MSFlexGrid4.ColWidth(3) = 1200 'No. Siri Produk
Frm108.MSFlexGrid4.ColWidth(4) = 3000 'Nama Produk
Frm108.MSFlexGrid4.ColWidth(5) = 1000 'Berat (g)
Frm108.MSFlexGrid4.ColWidth(6) = 2400 'Status
Frm108.MSFlexGrid4.ColWidth(7) = 1500 'No. Perjanjian A
Frm108.MSFlexGrid4.ColWidth(8) = 1500 'No. Perjanjian B
Frm108.MSFlexGrid4.ColWidth(9) = 1200 'Harga Jualan (RM)
'#### Header senarai agihan #### - End
End Sub
Sub Frm108_senarai_pulangan_barang_detail()
'on error resume next
Dim Frm108_LM_TOTAL_PAGE As Double

Frm108_PAGE_SIZE = 36
Frm108_LM_TOTAL_PAGE = 0
x = 0

Frm108.L31_Text = 0
Frm108.L32_Text = Format(0, "#,##0.00 g")

Frm108.L64_Text = 0
Frm108.L65_Text = Format(0, "#,##0.00 g")
Frm108.L66_Text = "RM " & Format(0, "#,##0.00")

LM_START_ROW = Frm108.L35_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm108_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm108.L36_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm108_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        Frm108.L33_Text = 1
    End If
End If

Frm108_LM_PAGE_FOUND = 0

Frm108.L30_Text = "Senarai barang yang dipulangkan dari No. Rujukan [" & Format(Frm108.L37_Text, "000000") & "]" 'Header : Senarai barang yang diagihkan

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 69_pulangan_barang where no_rujukan='" & Frm108.L37_Text & "' AND (status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "') order by no_siri_produk ASC LIMIT " & LM_START_ROW & "," & Frm108_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm108_LM_PAGE_FOUND = 0 Then
        If Frm108.L36_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm108.L33_Text = Frm108.L33_Text + 1 'Paparan Page ke-xxx
                Frm108_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm108.L33_Text) Then
                    If Frm108.L33_Text <> 1 Then
                        Frm108.L33_Text = Frm108.L33_Text - 1 'Paparan Page ke-xxx
                        Frm108_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm108.L33_Text - 1) * Frm108_PAGE_SIZE) + x
    Frm108.MSFlexGrid4.Rows = x + 1
    Frm108.MSFlexGrid4.TextMatrix(x, 0) = x 'No.
    Frm108.MSFlexGrid4.TextMatrix(x, 1) = Y 'No.
    Frm108.MSFlexGrid4.ColAlignment(1) = 4
    Frm108.MSFlexGrid4.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!no_siri_Produk) Then Frm108.MSFlexGrid4.TextMatrix(x, 3) = rs!no_siri_Produk 'No. Siri Produk
    Frm108.MSFlexGrid4.ColAlignment(3) = 4
    
    If Not IsNull(rs!kategori_Produk) Then Frm108.MSFlexGrid4.TextMatrix(x, 4) = rs!kategori_Produk 'Kategori produk
    
    If Not IsNull(rs!Berat) Then Frm108.MSFlexGrid4.TextMatrix(x, 5) = Format(rs!Berat, "#,##0.00 g") 'Berat
    Frm108.MSFlexGrid4.ColAlignment(5) = 4
    
    If Not IsNull(rs!Status) Then
        If rs!Status = 1 Then
            Frm108.MSFlexGrid4.TextMatrix(x, 6) = "SUDAH DIPULANGKAN"
        ElseIf rs!Status = 2 Then
            Frm108.MSFlexGrid4.TextMatrix(x, 6) = "SUDAH DIJUAL"
        End If
    End If
    
    If Not IsNull(rs!no_perjanjian_a) Then Frm108.MSFlexGrid4.TextMatrix(x, 7) = rs!no_perjanjian_a 'No. Perjanjian A
    Frm108.MSFlexGrid4.ColAlignment(7) = 4
    
    If Not IsNull(rs!no_perjanjian_b) Then Frm108.MSFlexGrid4.TextMatrix(x, 8) = rs!no_perjanjian_b 'No. Perjanjian B
    Frm108.MSFlexGrid4.ColAlignment(8) = 4
    
    If Not IsNull(rs!harga_jualan) Then Frm108.MSFlexGrid4.TextMatrix(x, 9) = Format(rs!harga_jualan, "#,##0.00") 'Harga Jualan (RM)
    Frm108.MSFlexGrid4.ColAlignment(9) = 4
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 69_pulangan_barang where no_rujukan='" & Frm108.L37_Text & "' AND (status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "') order by no_siri_produk ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm108_LM_TOTAL_PAGE = Format(rs(0) / Frm108_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm108_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm108_LM_PAGE = Split(Frm108_LM_TOTAL_PAGE, ".")(0)
        Frm108_LM_PAGE_LEBIHAN = Split(Frm108_LM_TOTAL_PAGE, ".")(1)
        
        If Frm108_LM_PAGE_LEBIHAN <> "00" Then
            Frm108.L34_Text = Frm108_LM_PAGE + 1
        Else
            Frm108.L34_Text = Frm108_LM_PAGE
        End If
        
    Else
    
        Frm108.L34_Text = Frm108_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm108.L34_Text = 0
    End If
Else
    Frm108.L34_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm108.L34_Text = vbNullString Then
    Frm108.L34_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah bilangan barang keseluruhan ### - Start (Dipulangkan)
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 69_pulangan_barang where no_rujukan='" & Frm108.L37_Text & "' AND status='" & 1 & "' order by no_siri_produk ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm108.L31_Text = rs(0)

rs.Close
Set rs = Nothing
'### Jumlah bilangan barang keseluruhan ### - End (Dipulangkan)

'### Jumlah berat keseluruhan ### - Start (Dijual)
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(berat) from 69_pulangan_barang where no_rujukan='" & Frm108.L37_Text & "' AND status='" & 1 & "' order by no_siri_produk ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm108.L32_Text = Format(rs(0), "#,##0.00 g")

rs.Close
Set rs = Nothing
'### Jumlah berat keseluruhan ### - End (Dijual)

'### Jumlah bilangan barang keseluruhan ### - Start (Barang yang dijual)
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 69_pulangan_barang where status='" & 2 & "' AND no_rujukan='" & Frm108.L37_Text & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm108.L64_Text = rs(0)

rs.Close
Set rs = Nothing
'### Jumlah bilangan barang keseluruhan ### - End (Barang yang dijual)

'### Jumlah berat keseluruhan ### - Start (Barang yang dijual)
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(berat) from 69_pulangan_barang where status='" & 2 & "' AND no_rujukan='" & Frm108.L37_Text & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm108.L65_Text = Format(rs(0), "#,##0.00 g")

rs.Close
Set rs = Nothing
'### Jumlah berat keseluruhan ### - End (Barang yang dijual)

'### Jumlah harga keseluruhan ### - Start (Barang yang dijual)
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_jualan) from 69_pulangan_barang where status='" & 2 & "' AND no_rujukan='" & Frm108.L37_Text & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm108.L66_Text = "RM " & Format(rs(0), "#,##0.00")

rs.Close
Set rs = Nothing
'### Jumlah harga keseluruhan ### - End (Barang yang dijual)

If x <> 0 Then
    Frm108.L35_Text = LM_START_ROW 'Titik Pencarian Data
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    'If Frm108_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm108.L36_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm108.L36_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub Frm108_senarai_pulangan_header()
'on error resume next
'#### Header senarai agihan #### - Start
Frm108.MSFlexGrid5.Clear
Frm108.MSFlexGrid5.Rows = 1
Frm108.MSFlexGrid5.RowHeight(0) = 600
Frm108.MSFlexGrid5.FormatString = "No.|<No.|<No. ID|<No. Siri Produk|<Nama Produk|<Berat (g)|<Status|<No. Perjanjian A|<No. Perjanjian B|<Harga Jualan (RM)"

Frm108.MSFlexGrid5.ColWidth(0) = 0 'No.
Frm108.MSFlexGrid5.ColWidth(1) = 600 'No.
Frm108.MSFlexGrid5.ColWidth(2) = 0 'No. ID
Frm108.MSFlexGrid5.ColWidth(3) = 1200 'No. Siri Produk
Frm108.MSFlexGrid5.ColWidth(4) = 3000 'Nama Produk
Frm108.MSFlexGrid5.ColWidth(5) = 1000 'Berat (g)
Frm108.MSFlexGrid5.ColWidth(6) = 2400 'Status
Frm108.MSFlexGrid5.ColWidth(7) = 1500 'No. Perjanjian A
Frm108.MSFlexGrid5.ColWidth(8) = 1500 'No. Perjanjian B
Frm108.MSFlexGrid5.ColWidth(9) = 1200 'Harga Jualan (RM)
'#### Header senarai agihan #### - End
End Sub
Sub Frm108_senarai_pulangan()
'on error resume next
Dim Frm108_LM_TOTAL_PAGE As Double

Frm108_PAGE_SIZE = 37
Frm108_LM_TOTAL_PAGE = 0
x = 0

Frm108.L42_Text = 0
Frm108.L43_Text = Format(0, "#,##0.00 g")

LM_START_ROW = Frm108.L46_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm108_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm108.L47_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm108_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        Frm108.L44_Text = 1
    End If
End If

Frm108_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from " & G_PULANGAN_TEMP & " where status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "' OR status='" & 4 & "' OR status='" & 5 & "' OR status='" & 6 & "' OR status='" & 7 & "' OR status='" & 8 & "' OR status='" & 9 & "' OR status='" & 10 & "' order by no_siri_produk ASC LIMIT " & LM_START_ROW & "," & Frm108_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm108_LM_PAGE_FOUND = 0 Then
        If Frm108.L47_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm108.L44_Text = Frm108.L44_Text + 1 'Paparan Page ke-xxx
                Frm108_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm108.L44_Text) Then
                    If Frm108.L44_Text <> 1 Then
                        Frm108.L44_Text = Frm108.L44_Text - 1 'Paparan Page ke-xxx
                        Frm108_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm108.L44_Text - 1) * Frm108_PAGE_SIZE) + x
    Frm108.MSFlexGrid5.Rows = x + 1
    Frm108.MSFlexGrid5.TextMatrix(x, 0) = x 'No.
    Frm108.MSFlexGrid5.TextMatrix(x, 1) = Y 'No.
    Frm108.MSFlexGrid5.ColAlignment(1) = 4
    Frm108.MSFlexGrid5.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!no_siri_Produk) Then Frm108.MSFlexGrid5.TextMatrix(x, 3) = rs!no_siri_Produk 'No. Siri Produk
    Frm108.MSFlexGrid5.ColAlignment(3) = 4
    
    If Not IsNull(rs!kategori_Produk) Then Frm108.MSFlexGrid5.TextMatrix(x, 4) = rs!kategori_Produk 'Kategori produk
    
    If Not IsNull(rs!Berat) Then Frm108.MSFlexGrid5.TextMatrix(x, 5) = Format(rs!Berat, "#,##0.00 g") 'Berat
    Frm108.MSFlexGrid5.ColAlignment(5) = 4
    
    If Not IsNull(rs!Status) Then
        If rs!Status = 1 Or rs!Status = 5 Then
            Frm108.MSFlexGrid5.TextMatrix(x, 6) = "Sedia dipulangkan"
        ElseIf rs!Status = 3 Or rs!Status = 7 Then
            Frm108.MSFlexGrid5.TextMatrix(x, 6) = "SUDAH DIPULANGKAN"
        ElseIf rs!Status = 2 Or rs!Status = 6 Then
            Frm108.MSFlexGrid5.TextMatrix(x, 6) = "Sedia dijual"
        ElseIf rs!Status = 4 Or rs!Status = 8 Then
            Frm108.MSFlexGrid5.TextMatrix(x, 6) = "SUDAH DIJUAL"
        ElseIf rs!Status = 9 Or rs!Status = 10 Then
            Frm108.MSFlexGrid5.TextMatrix(x, 6) = "PADAM"
        'ElseIf rs!Status = 3 Then
        '    Frm108.MSFlexGrid5.TextMatrix(X, 6) = "SUDAH DIPULANGKAN"
        End If
    End If
    
    If Not IsNull(rs!no_perjanjian_a) Then Frm108.MSFlexGrid5.TextMatrix(x, 7) = rs!no_perjanjian_a 'No. Perjanjian A
    Frm108.MSFlexGrid5.ColAlignment(7) = 4
    
    If Not IsNull(rs!no_perjanjian_b) Then Frm108.MSFlexGrid5.TextMatrix(x, 8) = rs!no_perjanjian_b 'No. Perjanjian B
    Frm108.MSFlexGrid5.ColAlignment(8) = 4
    
    If Not IsNull(rs!harga_jualan) Then Frm108.MSFlexGrid5.TextMatrix(x, 9) = Format(rs!harga_jualan, "#,##0.00") 'Harga Jualan (RM)
    Frm108.MSFlexGrid5.ColAlignment(9) = 4
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from " & G_PULANGAN_TEMP & " where status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "' OR status='" & 4 & "' OR status='" & 5 & "' OR status='" & 6 & "' OR status='" & 7 & "' OR status='" & 8 & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm108_LM_TOTAL_PAGE = Format(rs(0) / Frm108_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm108_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm108_LM_PAGE = Split(Frm108_LM_TOTAL_PAGE, ".")(0)
        Frm108_LM_PAGE_LEBIHAN = Split(Frm108_LM_TOTAL_PAGE, ".")(1)
        
        If Frm108_LM_PAGE_LEBIHAN <> "00" Then
            Frm108.L45_Text = Frm108_LM_PAGE + 1
        Else
            Frm108.L45_Text = Frm108_LM_PAGE
        End If
        
    Else
    
        Frm108.L45_Text = Frm108_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm108.L45_Text = 0
    End If
Else
    Frm108.L45_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm108.L45_Text = vbNullString Then
    Frm108.L45_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah bilangan barang keseluruhan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from " & G_PULANGAN_TEMP & " where status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "' OR status='" & 4 & "' OR status='" & 5 & "' OR status='" & 6 & "' OR status='" & 7 & "' OR status='" & 8 & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm108.L42_Text = rs(0)

rs.Close
Set rs = Nothing
'### Jumlah bilangan barang keseluruhan ### - End

'### Jumlah berat keseluruhan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(berat) from " & G_PULANGAN_TEMP & " where status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "' OR status='" & 4 & "' OR status='" & 5 & "' OR status='" & 6 & "' OR status='" & 7 & "' OR status='" & 8 & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm108.L43_Text = Format(rs(0), "#,##0.00 g")

rs.Close
Set rs = Nothing
'### Jumlah berat keseluruhan ### - End

If x <> 0 Then
    Frm108.L46_Text = LM_START_ROW 'Titik Pencarian Data
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    'If Frm108_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm108.L47_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm108.L47_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub Frm108_senarai_pulangan_barang_header()
'on error resume next
'#### Header senarai cawangan #### - Start
Frm108.MSFlexGrid3.Clear
Frm108.MSFlexGrid3.Rows = 1
Frm108.MSFlexGrid3.RowHeight(0) = 600
Frm108.MSFlexGrid3.FormatString = "No.|<No.|<No. ID|<Tarikh|<No. Rujukan|<Cawangan|<Nama PIC"

Frm108.MSFlexGrid3.ColWidth(0) = 0 'No.
Frm108.MSFlexGrid3.ColWidth(1) = 600 'No.
Frm108.MSFlexGrid3.ColWidth(2) = 0 'No. ID
Frm108.MSFlexGrid3.ColWidth(3) = 1500 'Tarikh
Frm108.MSFlexGrid3.ColWidth(4) = 1500 'No. Rujukan
Frm108.MSFlexGrid3.ColWidth(5) = 4000 'Cawangan
Frm108.MSFlexGrid3.ColWidth(6) = 6500 'Nama PIC
'#### Header senarai cawangan #### - End
End Sub
Sub Frm108_senarai_pulangan_barang()
'on error resume next
Dim Frm108_LM_TOTAL_PAGE As Double
Dim TM As Date
Dim TA As Date

Frm108_PAGE_SIZE = 36
Frm108_LM_TOTAL_PAGE = 0
x = 0

Frm108.L25_Text = 0

If Frm108.L20_Text = 1 Then '0 : Carian ikut supplier sahaja , 1 : Carian ikut tarikh , 2 : Carian ikut No. Rujukan
    TM = Frm108.L21_Text 'Tarikh mula
    TA = Frm108.L22_Text 'Tarikh akhir
End If
If Frm108.L23_Text = "Semua Cawangan" Then
    Frm108_LM_SEARCH_1 = Null
    Frm108_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm108_LM_SEARCH_1 = Frm108.L23_Text
    Frm108_LM_SEARCH_1_LOGIC = "="
End If

LM_START_ROW = Frm108.L28_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm108_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm108.L29_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm108_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        Frm108.L26_Text = 1
    End If
End If

Frm108_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm108.L20_Text = 0 Then rs.Open "select * from 68_pulangan where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND status='" & 1 & "' order by no_rujukan ASC LIMIT " & LM_START_ROW & "," & Frm108_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If Frm108.L20_Text = 1 Then rs.Open "select * from 68_pulangan where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND status='" & 1 & "' order by no_rujukan ASC LIMIT " & LM_START_ROW & "," & Frm108_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If Frm108.L20_Text = 2 Then rs.Open "select * from 68_pulangan where no_statement='" & Frm108.L23_Text & "' AND status='" & 1 & "'order by no_rujukan ASC LIMIT " & LM_START_ROW & "," & Frm108_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm108_LM_PAGE_FOUND = 0 Then
        If Frm108.L29_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm108.L26_Text = Frm108.L26_Text + 1 'Paparan Page ke-xxx
                Frm108_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm108.L26_Text) Then
                    If Frm108.L26_Text <> 1 Then
                        Frm108.L26_Text = Frm108.L26_Text - 1 'Paparan Page ke-xxx
                        Frm108_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm108.L26_Text - 1) * Frm108_PAGE_SIZE) + x
    Frm108.MSFlexGrid3.Rows = x + 1
    Frm108.MSFlexGrid3.TextMatrix(x, 0) = x 'No.
    Frm108.MSFlexGrid3.TextMatrix(x, 1) = Y 'No.
    Frm108.MSFlexGrid3.ColAlignment(1) = 4
    Frm108.MSFlexGrid3.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!tarikh) Then Frm108.MSFlexGrid3.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    Frm108.MSFlexGrid3.ColAlignment(3) = 4
    
    If Not IsNull(rs!no_statement) Then Frm108.MSFlexGrid3.TextMatrix(x, 4) = rs!no_statement 'No. rujukan (No. statement)
    Frm108.MSFlexGrid3.ColAlignment(4) = 4
    
    If Not IsNull(rs!cawangan) Then Frm108.MSFlexGrid3.TextMatrix(x, 5) = rs!cawangan 'Cawangan
    If Not IsNull(rs!Nama) Then Frm108.MSFlexGrid3.TextMatrix(x, 6) = rs!Nama 'Nama PIC
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm108.L20_Text = 0 Then rs.Open "select COUNT(ID) from 68_pulangan where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND status='" & 1 & "' order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic
If Frm108.L20_Text = 1 Then rs.Open "select COUNT(ID) from 68_pulangan where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND status='" & 1 & "' order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic
If Frm108.L20_Text = 2 Then rs.Open "select COUNT(ID) from 68_pulangan where no_statement='" & Frm108.L23_Text & "' AND status='" & 1 & "' order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm108_LM_TOTAL_PAGE = Format(rs(0) / Frm108_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm108_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm108_LM_PAGE = Split(Frm108_LM_TOTAL_PAGE, ".")(0)
        Frm108_LM_PAGE_LEBIHAN = Split(Frm108_LM_TOTAL_PAGE, ".")(1)
        
        If Frm108_LM_PAGE_LEBIHAN <> "00" Then
            Frm108.L27_Text = Frm108_LM_PAGE + 1
        Else
            Frm108.L27_Text = Frm108_LM_PAGE
        End If
        
    Else
    
        Frm108.L27_Text = Frm108_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm108.L27_Text = 0
    End If
Else
    Frm108.L27_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm108.L27_Text = vbNullString Then
    Frm108.L27_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah bilangan barang keseluruhan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm108.L20_Text = 0 Then rs.Open "select COUNT(ID) from 68_pulangan where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND status='" & 1 & "' order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic
If Frm108.L20_Text = 1 Then rs.Open "select COUNT(ID) from 68_pulangan where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND status='" & 1 & "' order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic
If Frm108.L20_Text = 2 Then rs.Open "select COUNT(ID) from 68_pulangan where no_statement='" & Frm108.L23_Text & "' AND status='" & 1 & "' order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm108.L25_Text = rs(0)

rs.Close
Set rs = Nothing
'### Jumlah bilangan barang keseluruhan ### - End

If x <> 0 Then
    Frm108.L28_Text = LM_START_ROW 'Titik Pencarian Data
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    'If Frm108_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm108.L29_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm108.L29_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub Frm108_recall_data_pulangan()
'On Error Resume Next
Frm108_LM_ID = vbNullString
Frm108_LM_No_PEKERJA = vbNullString
DATA_FOUND = 0

If Frm108.MSFlexGrid3 <> vbNullString Then
    
    If IsNumeric(Frm108.MSFlexGrid3) Then
        Frm108_LM_ID = Frm108.MSFlexGrid3.TextMatrix(Frm108.MSFlexGrid3, 2) 'No. ID
        
        If Frm108_LM_ID <> vbNullString Then
        
            Call Frm108_cmd_visible_3
            Call Frm108_hantaran_initial_setting
            Call Frm108_hantaran_initial_setting2
            
            '### Maklumat asas agihan ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 68_pulangan where ID='" & Frm108_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic

            If Not rs.EOF Then
                If Not IsNull(rs!no_rujukan) Then
                    Frm108_LM_No_RUJ = rs!no_rujukan
                    Frm108.L40_Text = rs!no_rujukan
                    DATA_FOUND = 1
                End If
                
                If Not IsNull(rs!tarikh) Then Frm108.DTPicker4 = rs!tarikh 'Tarikh barang diambil
                If Not IsNull(rs!Nama) Then Frm108.TB8 = rs!Nama 'Nama PIC
                If Not IsNull(rs!no_ic) Then Frm108.TB9 = rs!no_ic 'No. IC
                If Not IsNull(rs!no_tel) Then Frm108.TB10 = rs!no_tel 'No. telefon
                
                On Error GoTo Err_A:
                If Not IsNull(rs!cawangan) Then 'Cawangan
                    Frm108_LM_CAWANGAN = rs!cawangan
                    Frm108.CBB4 = Frm108_LM_CAWANGAN
                End If
        
Restore_A:
                'on error resume next
                If Not IsNull(rs!nama_pekerja) Then Frm108_LM_No_PEKERJA = rs!nama_pekerja

                
                
            End If
            
            rs.Close
            Set rs = Nothing
            '### Maklumat asas agihan ### - End
            
            If Frm108_LM_No_PEKERJA <> vbNullString Then
            
                '### Carian Maklumat Penjual (Data Pekerja) ### - Start
                DATA_PEKERJA_FOUND = 0
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from employee where NoPekerja='" & Frm108_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    Frm108_LM_MAKLUMAT_PEKERJA = rs!Samaran & "  |  " & rs!NoPekerja
                    DATA_PEKERJA_FOUND = 1
                End If
                
                rs.Close
                Set rs = Nothing
                
                If DATA_PEKERJA_FOUND = 1 Then
                    On Error GoTo Err_B:
                    Frm108.CBB5 = Frm108_LM_MAKLUMAT_PEKERJA
Restore_B:
                End If
                '### Carian Maklumat Penjual (Data Pekerja) ### - End
                
                'on error resume next
                
            End If
            
            If DATA_FOUND = 1 Then
            
            
'### #69_pulangan_barang -> #67_pulangan_barang_temp ### - Start (Barang dipulangkan)
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "insert into " & G_PULANGAN_TEMP & "(no_rujukan,no_rujukan_agihan,tarikh,no_siri_produk,kategori_produk,purity,berat,no_perjanjian_a,no_perjanjian_b,harga_jualan,status)" & _
                            "select no_rujukan,no_rujukan_agihan,tarikh,no_siri_produk,kategori_produk,purity,berat,no_perjanjian_a,no_perjanjian_b,harga_jualan,3 from 69_pulangan_barang WHERE no_rujukan='" & Frm108_LM_No_RUJ & "' AND status='" & 1 & "' order by no_siri_Produk ASC"

                Set rs = cn.Execute(strsql)
                Set rs = Nothing
'### #69_pulangan_barang -> #67_pulangan_barang_temp ### - End (Barang dipulangkan)

'### #69_pulangan_barang -> #67_pulangan_barang_temp ### - Start (Barang dijual)
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "insert into " & G_PULANGAN_TEMP & "(no_rujukan,no_rujukan_agihan,tarikh,no_siri_produk,kategori_produk,purity,berat,no_perjanjian_a,no_perjanjian_b,harga_jualan,status)" & _
                            "select no_rujukan,no_rujukan_agihan,tarikh,no_siri_produk,kategori_produk,purity,berat,no_perjanjian_a,no_perjanjian_b,harga_jualan,4 from 69_pulangan_barang WHERE no_rujukan='" & Frm108_LM_No_RUJ & "' AND status='" & 2 & "' order by no_siri_Produk ASC"

                Set rs = cn.Execute(strsql)
                Set rs = Nothing
'### #69_pulangan_barang -> #67_pulangan_barang_temp ### - End (Barang dijual)

                GM_NEXT_PREV = 0
                
                Frm108.L46_Text = -1 'Titik Pencarian Data
                Frm108.L47_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
                Frm108.L44_Text = 0 'Paparan Page ke-xxx
                
                Call Frm108_senarai_pulangan_header
                Call Frm108_senarai_pulangan
                Call Frm108_cmd_invisible_3
                
                Frm108.L39_Text = 1 '0 : Data baru , 1 : Edit
                Frm108.Pic6.Visible = True
                Frm108.Pic3.Visible = False
                Frm108.TB11.SetFocus
            
            End If
        
        End If
        
    End If
    
End If


Exit Sub

Err_A:
Frm108.CBB4.AddItem Frm108_LM_CAWANGAN
Frm108.CBB4 = Frm108_LM_CAWANGAN
Resume Restore_A:

Exit Sub
Err_B:
Frm108.CBB5.AddItem Frm108_LM_MAKLUMAT_PEKERJA
Frm108.CBB5 = Frm108_LM_MAKLUMAT_PEKERJA
Resume Restore_B:
End Sub
Sub Frm108_recall_data_agihan()
'On Error Resume Next
Frm108_LM_ID = vbNullString
Frm108_LM_No_PEKERJA = vbNullString
DATA_FOUND = 0

If Frm108.MSFlexGrid3 <> vbNullString Then
    
    If IsNumeric(Frm108.MSFlexGrid3) Then
        Frm108_LM_ID = Frm108.MSFlexGrid3.TextMatrix(Frm108.MSFlexGrid3, 2) 'No. ID
        
        If Frm108_LM_ID <> vbNullString Then
        
            Call Frm108_cmd_visible_2
            Call Frm108_hantaran_initial_setting
            Call Frm108_hantaran_initial_setting2
            
            '### Maklumat asas agihan ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 63_agihan where ID='" & Frm108_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!no_rujukan) Then
                    Frm108_LM_No_RUJ = rs!no_rujukan
                    Frm108.L12_Text = rs!no_rujukan
                    DATA_FOUND = 1
                End If
                
                If Not IsNull(rs!tarikh) Then Frm108.DTPicker1 = rs!tarikh 'Tarikh barang diambil
                If Not IsNull(rs!Nama) Then Frm108.TB1 = rs!Nama 'Nama PIC
                If Not IsNull(rs!no_ic) Then Frm108.TB2 = rs!no_ic 'No. IC
                If Not IsNull(rs!no_tel) Then Frm108.TB3 = rs!no_tel 'No. telefon
                
                On Error GoTo Err_A:
                If Not IsNull(rs!cawangan) Then 'Cawangan
                    Frm108_LM_CAWANGAN = rs!cawangan
                    Frm108.CBB1 = Frm108_LM_CAWANGAN
                End If
        
Restore_A:
                'on error resume next
                If Not IsNull(rs!nama_pekerja) Then Frm108_LM_No_PEKERJA = rs!nama_pekerja

                
                
            End If
            
            rs.Close
            Set rs = Nothing
            '### Maklumat asas agihan ### - End
            
            If Frm108_LM_No_PEKERJA <> vbNullString Then
            
                '### Carian Maklumat Penjual (Data Pekerja) ### - Start
                DATA_PEKERJA_FOUND = 0
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from employee where NoPekerja='" & Frm108_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    Frm108_LM_MAKLUMAT_PEKERJA = rs!Samaran & "  |  " & rs!NoPekerja
                    DATA_PEKERJA_FOUND = 1
                End If
                
                rs.Close
                Set rs = Nothing
                
                If DATA_PEKERJA_FOUND = 1 Then
                    On Error GoTo Err_B:
                    Frm108.CBB2 = Frm108_LM_MAKLUMAT_PEKERJA
Restore_B:
                End If
                '### Carian Maklumat Penjual (Data Pekerja) ### - End
                
                'on error resume next
                
            End If
            
            If DATA_FOUND = 1 Then
            
            
'### #64_agihan_barang -> #65_agihan_barang_temp ### - Start (Barang yang aktif (Belum dipulangkan))
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "insert into " & G_AGIHAN_TEMP & "(no_siri_produk,kategori_produk,purity,berat,status)" & _
                            "select no_siri_produk,kategori_produk,purity,berat,2 from 64_agihan_barang WHERE no_rujukan='" & Frm108_LM_No_RUJ & "' AND status='" & 1 & "' order by no_siri_Produk ASC"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
'### #64_agihan_barang -> #65_agihan_barang_temp ### - End (Barang yang aktif (Belum dipulangkan))

            
'### #64_agihan_barang -> #65_agihan_barang_temp ### - Start (Barang yang telah dipulangkan)
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "insert into " & G_AGIHAN_TEMP & "(no_siri_produk,kategori_produk,purity,berat,status)" & _
                            "select no_siri_produk,kategori_produk,purity,berat,6 from 64_agihan_barang WHERE no_rujukan='" & Frm108_LM_No_RUJ & "' AND status='" & 2 & "' order by no_siri_Produk ASC"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
'### #64_agihan_barang -> #65_agihan_barang_temp ### - End (Barang yang telah dipulangkan)

'### #64_agihan_barang -> #65_agihan_barang_temp ### - Start (Barang yang telah terjual)
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                
                strsql = "insert into " & G_AGIHAN_TEMP & "(no_siri_produk,kategori_produk,purity,berat,status)" & _
                            "select no_siri_produk,kategori_produk,purity,berat,7 from 64_agihan_barang WHERE no_rujukan='" & Frm108_LM_No_RUJ & "' AND status='" & 3 & "' order by no_siri_Produk ASC"
                
                Set rs = cn.Execute(strsql)
                Set rs = Nothing
'### #64_agihan_barang -> #65_agihan_barang_temp ### - End (Barang yang telah terjual)

                GM_NEXT_PREV = 0
                
                Frm108.L15_Text = -1 'Titik Pencarian Data
                Frm108.L16_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
                Frm108.L13_Text = 0 'Paparan Page ke-xxx
                
                Call Frm108_senarai_agihan_header
                Call Frm108_senarai_agihan
                Call Frm108_cmd_invisible_2
                
                Frm108.L1_Text = 1 '0 : Data baru , 1 : Edit
                Frm108.Pic2.Visible = True
                Frm108.Pic3.Visible = False
                Frm108.TB1.SetFocus
            
            End If
        
        End If
        
    End If
    
End If


Exit Sub

Err_A:
Frm108.CBB1.AddItem Frm108_LM_CAWANGAN
Frm108.CBB1 = Frm108_LM_CAWANGAN
Resume Restore_A:

Exit Sub
Err_B:
Frm108.CBB2.AddItem Frm108_LM_MAKLUMAT_PEKERJA
Frm108.CBB2 = Frm108_LM_MAKLUMAT_PEKERJA
Resume Restore_B:
End Sub
Sub Frm108_padam_data_agihan()
'On Error Resume Next
Frm108_LM_ID = vbNullString
DATA_FOUND = 0

If Frm108.MSFlexGrid3 <> vbNullString Then
    
    If IsNumeric(Frm108.MSFlexGrid3) Then
        Frm108_LM_ID = Frm108.MSFlexGrid3.TextMatrix(Frm108.MSFlexGrid3, 2) 'No. ID
        
        If Frm108_LM_ID <> vbNullString Then
        
            Note = "Adakah anda yakin untuk memadamkan data agihan ini?" & vbCrLf & _
                    "Semua data yang telah diagihkan akan dipulangkan ke dalam stok kedai." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Teruskan?"
            
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbYes Then
            
                '### Maklumat asas agihan ### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 63_agihan where ID='" & Frm108_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    
                    If Not IsNull(rs!no_rujukan) Then
                        Frm108_LM_NO_RUJUKAN = rs!no_rujukan
                        If Not IsNull(rs!no_statement) Then Frm108_LM_STATEMENT = rs!no_statement
                        rs!Status = 0
                        rs!write_timestamp3 = Now
                        rs.Update
                        
                        DATA_FOUND = 1
                    End If
                    
                End If
                
                rs.Close
                Set rs = Nothing
                
                If DATA_FOUND = 1 Then
                
            '### Update status barang dalam table #data_database ### - Start
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    
                    strsql = "UPDATE Data_Database,64_agihan_barang SET Data_Database.StatusItem='" & 10 & "'," _
                    & "Data_Database.cawangan_id = NULL " _
                    & "WHERE Data_Database.no_siri_produk = 64_agihan_barang.no_siri_produk AND 64_agihan_barang.no_rujukan='" & Frm108_LM_NO_RUJUKAN & "' AND 64_agihan_barang.status = 1"
                    
                    Set rs = cn.Execute(strsql)
                    Set rs = Nothing
            '### Update status barang dalam table #data_database ### - End
            
            '### #65_agihan_barang_temp -> #64_agihan_barang ### - Start (Data yang dipadamkan)
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            
                    strsql = "UPDATE 64_agihan_barang SET status = 0 ," _
                    & "write_timestamp3 = Now()" _
                    & "WHERE no_rujukan='" & Frm108_LM_NO_RUJUKAN & "'"
                  
                    Set rs = cn.Execute(strsql)
                    Set rs = Nothing
            '### #65_agihan_barang_temp -> #64_agihan_barang ### - End (Data yang dipadamkan)
            
            '#### Update Log Aktiviti Sistem #### - Start
                    user = MDI_frm1.L3_Text
                    
                    LogAct_Memory = "[" & user & "] Padam data agihan kepada cawangan. No. Rujukan [" & Frm108_LM_STATEMENT & "]."
                    LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                    Call UpdateLog_Database
            '#### Update Log Aktiviti Sistem #### - End
            
                    GM_NEXT_PREV = 2
            
                    Call Frm108_senarai_agihan_barang_header
                    Call Frm108_senarai_agihan_barang
                    
                    MsgBox "Data telah berjaya dipadamkan.", vbInformation, "Info"
                    
                End If
            
            End If
        
        End If
        
    End If

End If
End Sub
Sub Frm108_padam_data_pulangan()
'On Error Resume Next
Frm108_LM_ID = vbNullString
DATA_FOUND = 0

If Frm108.MSFlexGrid3 <> vbNullString Then
    
    If IsNumeric(Frm108.MSFlexGrid3) Then
        Frm108_LM_ID = Frm108.MSFlexGrid3.TextMatrix(Frm108.MSFlexGrid3, 2) 'No. ID
        
        If Frm108_LM_ID <> vbNullString Then
        
            Note = "Adakah anda yakin untuk memadamkan data pulangan ini?" & vbCrLf & _
                    "Semua data yang akan dipulangkan dari penyata ini akan ditukar status kepada agihan mengikut data sebelum barang dipulangkan." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Teruskan?"
            
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbYes Then
            
                '### Maklumat asas pulangan ### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 68_pulangan where ID='" & Frm108_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    
                    If Not IsNull(rs!no_rujukan) Then
                        Frm108_LM_NO_RUJUKAN = rs!no_rujukan
                        If Not IsNull(rs!no_statement) Then Frm108_LM_STATEMENT = rs!no_statement
                        rs!Status = 0
                        rs!write_timestamp3 = Now
                        rs.Update
                        
                        DATA_FOUND = 1
                    End If
                    
                End If
                
                rs.Close
                Set rs = Nothing
                
                If DATA_FOUND = 1 Then
                
            '### Update status barang dalam table #data_database ### - Start (Bagi barang yang dipulangkan dan dijual)
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    
                    strsql = "UPDATE Data_Database,69_pulangan_barang SET Data_Database.StatusItem='" & 25 & "'" _
                    & "WHERE Data_Database.no_siri_produk = 69_pulangan_barang.no_siri_produk AND 69_pulangan_barang.no_rujukan='" & Frm108_LM_NO_RUJUKAN & "' AND (69_pulangan_barang.status = 1 OR 69_pulangan_barang.status = 2) AND data_database.no_rujukan_pulang='" & Frm108_LM_NO_RUJUKAN & "'"
                    
                    Set rs = cn.Execute(strsql)
                    Set rs = Nothing
            '### Update status barang dalam table #data_database ### - End (Bagi barang yang dipulangkan dan dijual)

            
            '### Update status dan tarikh pulangan dalam table #64_agihan_barang ### - Start (Padam data)
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    
                    strsql = "UPDATE 64_agihan_barang,69_pulangan_barang SET 64_agihan_barang.status='" & 1 & "'," _
                    & "64_agihan_barang.tarikh_jual = NULL ," _
                    & "64_agihan_barang.write_timestamp3='" & Now & "'" _
                    & "WHERE 64_agihan_barang.no_siri_produk = 69_pulangan_barang.no_siri_produk AND 64_agihan_barang.no_rujukan = 69_pulangan_barang.no_rujukan_agihan AND (69_pulangan_barang.status = 1 OR 69_pulangan_barang.status = 2)"
                    
                    Set rs = cn.Execute(strsql)
                    Set rs = Nothing
            '### Update status dan tarikh pulangan dalam table #64_agihan_barang ### - End (Padam data)
            
            '### Update status #69_pulangan_barang ### - Start
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            
                    strsql = "UPDATE 69_pulangan_barang SET status = 0 ," _
                    & "write_timestamp3 = Now()" _
                    & "WHERE no_rujukan='" & Frm108_LM_NO_RUJUKAN & "' AND (69_pulangan_barang.status = 1 OR 69_pulangan_barang.status = 2)"
                  
                    Set rs = cn.Execute(strsql)
                    Set rs = Nothing
            '### Update status #69_pulangan_barang ### - End
            
            '#### Update Log Aktiviti Sistem #### - Start
                    user = MDI_frm1.L3_Text
                    
                    LogAct_Memory = "[" & user & "] Padam data pulangan barang kepada cawangan. No. Rujukan [" & Frm108_LM_STATEMENT & "]."
                    LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                    Call UpdateLog_Database
            '#### Update Log Aktiviti Sistem #### - End
            
                    GM_NEXT_PREV = 2
            
                    Call Frm108_senarai_pulangan_barang_header
                    Call Frm108_senarai_pulangan_barang
    
                    MsgBox "Data telah berjaya dipadamkan.", vbInformation, "Info"
                    
                End If
            
            End If
        
        End If
        
    End If

End If
End Sub
Sub Frm108_report_inventory_header()
'on error resume next
'#### Header senarai cawangan #### - Start
Frm108.MSFlexGrid6.Clear
Frm108.MSFlexGrid6.Rows = 1
Frm108.MSFlexGrid6.RowHeight(0) = 600
Frm108.MSFlexGrid6.FormatString = "No.|<No.|<No. ID|<Tarikh Agihan|<Cawangan|<No. Siri Produk|<Nama Produk|<Berat (g)|<Status|<Tarikh Pulangan / Jual"

Frm108.MSFlexGrid6.ColWidth(0) = 0 'No.
Frm108.MSFlexGrid6.ColWidth(1) = 600 'No.
Frm108.MSFlexGrid6.ColWidth(2) = 0 'No. ID
Frm108.MSFlexGrid6.ColWidth(3) = 1200 'Tarikh Agihan
Frm108.MSFlexGrid6.ColWidth(4) = 4000 'Cawangan
Frm108.MSFlexGrid6.ColWidth(5) = 1300 'No. Siri Produk
Frm108.MSFlexGrid6.ColWidth(6) = 3700 'Nama Produk
Frm108.MSFlexGrid6.ColWidth(7) = 1200 'Berat (g)
Frm108.MSFlexGrid6.ColWidth(8) = 1000 'Status
Frm108.MSFlexGrid6.ColWidth(9) = 1200 'Tarikh Pulangan / Jual
'#### Header senarai cawangan #### - End
End Sub
Sub Frm108_report_inventory()
'on error resume next
Dim Frm108_LM_TOTAL_PAGE As Double
Dim TM As Date
Dim TA As Date

Frm108_PAGE_SIZE = 37
Frm108_LM_TOTAL_PAGE = 0
x = 0

Frm108.L53_Text = 0 'Report inventory : Bilangan
Frm108.L54_Text = "0.00 g" 'Report inventory : Jumlah berat

If Frm108.L49_Text = 1 Then 'Memory : Jenis report , 0 : Tiada pilihan tarikh , 1 : Ada pilihan
    TM = Frm108.L60_Text 'Tarikh mula
    TA = Frm108.L61_Text 'Tarikh akhir
End If
If Frm108.L50_Text = "Semua Cawangan" Then
    Frm108_LM_SEARCH_1 = Null
    Frm108_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm108_LM_SEARCH_1 = Frm108.L50_Text
    Frm108_LM_SEARCH_1_LOGIC = "="
End If
If Frm108.L51_Text = "Semua Jenis Report" Then

    Frm108_LM_SEARCH_2 = 1
    Frm108_LM_SEARCH_2_LOGIC = "="

    Frm108_LM_SEARCH_3 = 2
    Frm108_LM_SEARCH_3_LOGIC = "="
    
    Frm108_LM_SEARCH_4 = 3
    Frm108_LM_SEARCH_4_LOGIC = "="
    
ElseIf Frm108.L51_Text = "Agihan" Then

    Frm108_LM_SEARCH_2 = 1
    Frm108_LM_SEARCH_2_LOGIC = "="
    
    Frm108_LM_SEARCH_3 = 2
    Frm108_LM_SEARCH_3_LOGIC = "="
    
    Frm108_LM_SEARCH_4 = 3
    Frm108_LM_SEARCH_4_LOGIC = "="
    
ElseIf Frm108.L51_Text = "Pulangan" Then

    Frm108_LM_SEARCH_2 = 2
    Frm108_LM_SEARCH_2_LOGIC = "="
    
    Frm108_LM_SEARCH_3 = 2
    Frm108_LM_SEARCH_3_LOGIC = "="
    
    Frm108_LM_SEARCH_4 = 2
    Frm108_LM_SEARCH_4_LOGIC = "="
    
ElseIf Frm108.L51_Text = "Dijual" Then

    Frm108_LM_SEARCH_2 = 3
    Frm108_LM_SEARCH_2_LOGIC = "="
    
    Frm108_LM_SEARCH_3 = 3
    Frm108_LM_SEARCH_3_LOGIC = "="
    
    Frm108_LM_SEARCH_4 = 3
    Frm108_LM_SEARCH_4_LOGIC = "="
    
ElseIf Frm108.L51_Text = "Belum Dipulangkan" Then

    Frm108_LM_SEARCH_2 = 1
    Frm108_LM_SEARCH_2_LOGIC = "="
    
    Frm108_LM_SEARCH_3 = 1
    Frm108_LM_SEARCH_3_LOGIC = "="
    
    Frm108_LM_SEARCH_4 = 1
    Frm108_LM_SEARCH_4_LOGIC = "="
    
End If

If Frm108.L49_Text = 0 Then Frm108.L52_Text = "Report inventori bagi [" & Frm108.L51_Text & "] kepada/oleh cawangan [" & Frm108.L50_Text & "]."
If Frm108.L49_Text = 1 Then Frm108.L52_Text = "Report inventori bagi [" & Frm108.L51_Text & "] kepada/oleh cawangan [" & Frm108.L50_Text & "] dari " & TM & " hingga " & TA & "."

LM_START_ROW = Frm108.L57_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm108_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm108.L58_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm108_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        Frm108.L55_Text = 1
    End If
End If

Frm108_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm108.L49_Text = 0 Then rs.Open "select * from 64_agihan_barang where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND (status='" & Frm108_LM_SEARCH_2 & "' OR status='" & Frm108_LM_SEARCH_3 & "' OR status='" & Frm108_LM_SEARCH_4 & "') order by no_rujukan ASC LIMIT " & LM_START_ROW & "," & Frm108_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If Frm108.L49_Text = 1 Then rs.Open "select * from 64_agihan_barang where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND (status='" & Frm108_LM_SEARCH_2 & "' OR status='" & Frm108_LM_SEARCH_3 & "' OR status='" & Frm108_LM_SEARCH_4 & "') AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by no_rujukan ASC LIMIT " & LM_START_ROW & "," & Frm108_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm108_LM_PAGE_FOUND = 0 Then
        If Frm108.L58_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm108.L55_Text = Frm108.L55_Text + 1 'Paparan Page ke-xxx
                Frm108_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm108.L55_Text) Then
                    If Frm108.L55_Text <> 1 Then
                        Frm108.L55_Text = Frm108.L55_Text - 1 'Paparan Page ke-xxx
                        Frm108_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm108.L55_Text - 1) * Frm108_PAGE_SIZE) + x
    Frm108.MSFlexGrid6.Rows = x + 1
    Frm108.MSFlexGrid6.TextMatrix(x, 0) = x 'No.
    Frm108.MSFlexGrid6.TextMatrix(x, 1) = Y 'No.
    Frm108.MSFlexGrid6.ColAlignment(1) = 4
    Frm108.MSFlexGrid6.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!tarikh) Then Frm108.MSFlexGrid6.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    Frm108.MSFlexGrid6.ColAlignment(3) = 4
    
    If Not IsNull(rs!cawangan) Then Frm108.MSFlexGrid6.TextMatrix(x, 4) = rs!cawangan 'Cawangan
    
    If Not IsNull(rs!no_siri_Produk) Then Frm108.MSFlexGrid6.TextMatrix(x, 5) = rs!no_siri_Produk 'No. Siri Produk
    Frm108.MSFlexGrid6.ColAlignment(5) = 4
    
    If Not IsNull(rs!kategori_Produk) Then Frm108.MSFlexGrid6.TextMatrix(x, 6) = rs!kategori_Produk 'Nama Produk
    
    If Not IsNull(rs!Berat) Then Frm108.MSFlexGrid6.TextMatrix(x, 7) = Format(rs!Berat, "#,##0.00 g") 'Berat (g)
    Frm108.MSFlexGrid6.ColAlignment(7) = 4
    
    
    If Not IsNull(rs!Status) Then
        If rs!Status = 1 Then
            Frm108.MSFlexGrid6.TextMatrix(x, 8) = "Agihan"
        ElseIf rs!Status = 2 Then
            Frm108.MSFlexGrid6.TextMatrix(x, 8) = "Pulang"
        ElseIf rs!Status = 3 Then
            Frm108.MSFlexGrid6.TextMatrix(x, 8) = "Jual"
        End If
    Else
        Frm108.MSFlexGrid6.TextMatrix(x, 8) = "-"
    End If
    Frm108.MSFlexGrid6.ColAlignment(8) = 4
    
    If Not IsNull(rs!tarikh_jual) Then
        Frm108.MSFlexGrid6.TextMatrix(x, 9) = rs!tarikh_jual 'Tarikh pulangan
    Else
        Frm108.MSFlexGrid6.TextMatrix(x, 9) = "-" 'Tarikh pulangan
    End If
    Frm108.MSFlexGrid6.ColAlignment(9) = 4
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm108.L49_Text = 0 Then rs.Open "select COUNT(ID) from 64_agihan_barang where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND (status='" & Frm108_LM_SEARCH_2 & "' OR status='" & Frm108_LM_SEARCH_3 & "' OR status='" & Frm108_LM_SEARCH_4 & "') order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic
If Frm108.L49_Text = 1 Then rs.Open "select COUNT(ID) from 64_agihan_barang where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND (status='" & Frm108_LM_SEARCH_2 & "' OR status='" & Frm108_LM_SEARCH_3 & "' OR status='" & Frm108_LM_SEARCH_4 & "') AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm108_LM_TOTAL_PAGE = Format(rs(0) / Frm108_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm108_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm108_LM_PAGE = Split(Frm108_LM_TOTAL_PAGE, ".")(0)
        Frm108_LM_PAGE_LEBIHAN = Split(Frm108_LM_TOTAL_PAGE, ".")(1)
        
        If Frm108_LM_PAGE_LEBIHAN <> "00" Then
            Frm108.L56_Text = Frm108_LM_PAGE + 1
        Else
            Frm108.L56_Text = Frm108_LM_PAGE
        End If
        
    Else
    
        Frm108.L56_Text = Frm108_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm108.L56_Text = 0
    End If
Else
    Frm108.L56_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm108.L56_Text = vbNullString Then
    Frm108.L56_Text = 0
End If
'### Jumlah Data ### - End

'### Jumlah bilangan barang keseluruhan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm108.L49_Text = 0 Then rs.Open "select COUNT(ID) from 64_agihan_barang where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND (status='" & Frm108_LM_SEARCH_2 & "' OR status='" & Frm108_LM_SEARCH_3 & "' OR status='" & Frm108_LM_SEARCH_4 & "') order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic
If Frm108.L49_Text = 1 Then rs.Open "select COUNT(ID) from 64_agihan_barang where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND (status='" & Frm108_LM_SEARCH_2 & "' OR status='" & Frm108_LM_SEARCH_3 & "' OR status='" & Frm108_LM_SEARCH_4 & "') AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm108.L53_Text = rs(0)

rs.Close
Set rs = Nothing
'### Jumlah bilangan barang keseluruhan ### - End

'### Jumlah bilangan barang keseluruhan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm108.L49_Text = 0 Then rs.Open "select SUM(berat) from 64_agihan_barang where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND (status='" & Frm108_LM_SEARCH_2 & "' OR status='" & Frm108_LM_SEARCH_3 & "' OR status='" & Frm108_LM_SEARCH_4 & "') order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic
If Frm108.L49_Text = 1 Then rs.Open "select SUM(berat) from 64_agihan_barang where cawangan " & Frm108_LM_SEARCH_1_LOGIC & "'" & Frm108_LM_SEARCH_1 & "' AND (status='" & Frm108_LM_SEARCH_2 & "' OR status='" & Frm108_LM_SEARCH_3 & "' OR status='" & Frm108_LM_SEARCH_4 & "') AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by no_rujukan ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm108.L54_Text = Format(rs(0), "#,##0.00 g")

rs.Close
Set rs = Nothing
'### Jumlah bilangan barang keseluruhan ### - End

If x <> 0 Then
    Frm108.L57_Text = LM_START_ROW 'Titik Pencarian Data
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    'If Frm108_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm108.L58_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm108.L58_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub frm108_jurujual()
'On Error Resume Next
If MDI_frm1.L3_Text <> vbNullString Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        Frm108.CBB2 = rs!Samaran & "  |  " & rs!NoPekerja
        Frm108.CBB5 = rs!Samaran & "  |  " & rs!NoPekerja
        
    Else
        
        Frm108.CBB2.AddItem "" & "  |  " & rs!Samaran
        Frm108.CBB2 = "" & "  |  " & rs!Samaran
        
        Frm108.CBB5.AddItem "" & "  |  " & rs!Samaran
        Frm108.CBB5 = "" & "  |  " & rs!Samaran
        
    End If
    
    rs.Close
    Set rs = Nothing
    
    If G_LOCK_JURUJUAL = "YES" Then
    
        Frm108.CBB2.Enabled = False
        Frm108.CBB2.BackColor = &H8000000A
        
        Frm108.CBB5.Enabled = False
        Frm108.CBB5.BackColor = &H8000000A

    Else
    
        Frm108.CBB2.Enabled = True
        Frm108.CBB2.BackColor = &HFFFFFF
        
        Frm108.CBB5.Enabled = True
        Frm108.CBB5.BackColor = &HFFFFFF

    End If

End If
End Sub
