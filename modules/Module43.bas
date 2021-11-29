Attribute VB_Name = "mod_buku_cek"
Sub Frm86_Initial_Setting()
'on error resume next
GLOBAL_DISABLE = 0
Frm86.Pic1.Left = 120
Frm86.Pic1.Top = 240
Frm86.Pic2.Left = 120
Frm86.Pic2.Top = 240
Frm86.Pic3.Left = 120
Frm86.Pic3.Top = 240
Frm86.Pic4.Left = 120
Frm86.Pic4.Top = 240

Frm86.L4_Text = 0
Frm86.L7_Text = vbNullString

Frm86.L4_Text.BackStyle = 0

Frm86.TB1 = vbNullString
Frm86.TB2 = vbNullString
Frm86.TB3 = vbNullString
Frm86.TB4 = "0.00"
Frm86.TB5 = vbNullString
Frm86.TB6 = vbNullString

Frm86.DTPicker1 = DateTime.Date

Frm86.CMD1.Visible = True
Frm86.CMD2.Visible = False
Frm86.Pic1.Visible = False
Frm86.Pic2.Visible = False
Frm86.Pic3.Visible = False
Frm86.Pic4.Visible = False

Frm86.CMD3.Visible = True
Frm86.CMD4.Visible = False
Frm86.CMD5.Visible = False

Frm86.CBB1.Clear
Frm86.CBB2.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 25_buku_cek", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Frm86.CBB1.AddItem rs!nama_bank
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where status='" & "Aktif" & "' AND InvestorSmall = 0 AND InvestorBig = 0", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    If rs!user_level = 1 Or rs!user_level = 2 Or rs!user_level = 3 Or rs!user_level = G_LEVEL_USER Then Frm86.CBB2.AddItem rs!Samaran & "  |  " & rs!NoPekerja
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
End Sub
Sub Frm86_Senarai_Buku_Cek_Header()
'on error resume next
Frm86.MSFlexGrid1.Clear
Frm86.MSFlexGrid1.RowHeight(0) = 800
Frm86.MSFlexGrid1.FormatString = "No.|<No.|<ID|<Nama Bank|<No. Akaun"

Frm86.MSFlexGrid1.Rows = 1
Frm86.MSFlexGrid1.ColWidth(0) = 600
Frm86.MSFlexGrid1.ColWidth(1) = 0
Frm86.MSFlexGrid1.ColWidth(2) = 0 'No. ID
Frm86.MSFlexGrid1.ColWidth(3) = 6000 'Nama Bank
Frm86.MSFlexGrid1.ColWidth(4) = 2800 'No. Akaun
End Sub
Sub Frm86_Senarai_Buku_Cek()
'on error resume next
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 25_buku_cek", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm86.MSFlexGrid1.Rows = x + 1
    Frm86.MSFlexGrid1.TextMatrix(x, 0) = x 'No.
    Frm86.MSFlexGrid1.TextMatrix(x, 1) = x 'No.
    If Not IsNull(rs!ID) Then Frm86.MSFlexGrid1.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!nama_bank) Then Frm86.MSFlexGrid1.TextMatrix(x, 3) = rs!nama_bank 'Nama Bank Bagi Buku Cek Ini
    If Not IsNull(rs!no_akaun) Then Frm86.MSFlexGrid1.TextMatrix(x, 4) = rs!no_akaun 'No. Akaun Bagi No. Bank Ini
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
End Sub
Sub Frm86_Senarai_Cek_Header()
'on error resume next
Frm86.MSFlexGrid2.Clear
Frm86.MSFlexGrid2.RowHeight(0) = 800
Frm86.MSFlexGrid2.FormatString = "No.|<No.|<ID|<Tarikh|<Nama Bank|<No. Akaun|<No. Cek|<Jumlah (RM)|<Dibayar Kepada|<Remarks"

Frm86.MSFlexGrid2.Rows = 1
Frm86.MSFlexGrid2.ColWidth(0) = 600
Frm86.MSFlexGrid2.ColWidth(1) = 0
Frm86.MSFlexGrid2.ColWidth(2) = 0 'No. ID
Frm86.MSFlexGrid2.ColWidth(3) = 1300 'Tarikh
Frm86.MSFlexGrid2.ColWidth(4) = 3650 'Nama Bank
Frm86.MSFlexGrid2.ColWidth(5) = 3000 'No. Akaun
Frm86.MSFlexGrid2.ColWidth(6) = 2500 'No. Cek
Frm86.MSFlexGrid2.ColWidth(7) = 1700 'Jumlah (RM)
Frm86.MSFlexGrid2.ColWidth(8) = 5000 'Dibayar Kepada
Frm86.MSFlexGrid2.ColWidth(9) = 5000 'Remarks
End Sub
Sub Frm86_Senarai_Cek()
'on error resume next
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 26_senarai_cek", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm86.MSFlexGrid2.Rows = x + 1
    Frm86.MSFlexGrid2.TextMatrix(x, 0) = x 'No.
    Frm86.MSFlexGrid2.TextMatrix(x, 1) = x 'No.
    If Not IsNull(rs!ID) Then Frm86.MSFlexGrid2.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm86.MSFlexGrid2.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!nama_bank) Then Frm86.MSFlexGrid2.TextMatrix(x, 4) = rs!nama_bank 'Nama Bank
    If Not IsNull(rs!no_akaun) Then Frm86.MSFlexGrid2.TextMatrix(x, 5) = rs!no_akaun 'No. Akaun Bank
    If Not IsNull(rs!no_cek) Then Frm86.MSFlexGrid2.TextMatrix(x, 6) = rs!no_cek 'No. Cek
    If Not IsNull(rs!jumlah) Then Frm86.MSFlexGrid2.TextMatrix(x, 7) = rs!jumlah 'Jumlah (RM)
    If Not IsNull(rs!penerima) Then Frm86.MSFlexGrid2.TextMatrix(x, 8) = rs!penerima 'Nama Penerima
    If Not IsNull(rs!remarks) Then Frm86.MSFlexGrid2.TextMatrix(x, 9) = rs!remarks 'Remarks
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
End Sub
