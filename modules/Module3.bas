Attribute VB_Name = "Module3"
Sub UpdateLog_Database()
'On Error Resume Next
If MDI_frm1.L20_Text = "Semua cawangan" Then
    LM_NAMA_CAWANGAN = "HQ"
Else
    LM_NAMA_CAWANGAN = G_CAWANGAN
End If

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main2 Else Exit Sub
strsql = "insert into log(Log_Tarikh,Log_Aktiviti,terminal,UserName,cawangan)" & _
            "select '" & LogDate_Memory & "','" & LogAct_Memory & "','" & G_TERMINAL & "','" & G_LOGIN_USER & "','" & LM_NAMA_CAWANGAN & "'"

Set rs = cn3.Execute(strsql)
Set rs = Nothing

LogDate_Memory = vbNullString
LogAct_Memory = vbNullString
End Sub
Sub Frm48_Default()
'On Error Resume Next
Frm48.Pic1.Top = 360
Frm48.Pic2.Top = 360
Frm48.Pic3.Top = 360
Frm48.Pic4.Top = 360
Frm48.Pic5.Top = 360
Frm48.Pic1.Left = 120
Frm48.Pic2.Left = 120
Frm48.Pic3.Left = 120
Frm48.Pic4.Left = 120
Frm48.Pic5.Left = 120
'Frm48.Pic6.Top = 2520
'Frm48.Pic6.Left = 8160
'Frm48.Pic7.Top = 2520
'Frm48.Pic7.Left = 8160

Frm48.Pic1.Visible = False
Frm48.Pic2.Visible = False
Frm48.Pic3.Visible = False
Frm48.Pic4.Visible = False
Frm48.Pic5.Visible = False
'Frm48.Pic6.Visible = False
'Frm48.Pic7.Visible = False

Frm48.L27_Text.BackStyle = 0

Frm48.CBB1.Clear
Frm48.CBB2.Clear
With Frm48.CBB1
    .AddItem "Januari"
    .AddItem "Februari"
    .AddItem "Mac"
    .AddItem "April"
    .AddItem "Mei"
    .AddItem "Jun"
    .AddItem "Julai"
    .AddItem "Ogos"
    .AddItem "September"
    .AddItem "Oktober"
    .AddItem "November"
    .AddItem "Disember"
End With
With Frm48.CBB2
    .AddItem 2016
    .AddItem 2017
    .AddItem 2018
    .AddItem 2019
    .AddItem 2020
End With
Frm48.DTPicker1 = DateTime.Date
Frm48.DTPicker2 = DateTime.Date
End Sub
Sub Frm48_ListPayroll()
'On Error Resume Next
Frm48.MSFlexGrid1.Clear
Frm48.MSFlexGrid1.RowHeight(0) = 600
Frm48.MSFlexGrid1.FormatString = "No.|<No.|<Bulan|<Tahun|<Tarikh Mula|<Tarikh Akhir"

Frm48.MSFlexGrid1.Rows = 1
Frm48.MSFlexGrid1.ColWidth(0) = 600
Frm48.MSFlexGrid1.ColWidth(1) = 0
Frm48.MSFlexGrid1.ColWidth(2) = 1950
Frm48.MSFlexGrid1.ColWidth(3) = 1950
Frm48.MSFlexGrid1.ColWidth(4) = 2000
Frm48.MSFlexGrid1.ColWidth(5) = 2000

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from tetapan_Payslip", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm48.MSFlexGrid1.Rows = x + 1
    Frm48.MSFlexGrid1.TextMatrix(x, 0) = x
    Frm48.MSFlexGrid1.TextMatrix(x, 1) = x
    If Not IsNull(rs!Bulan) Then Frm48.MSFlexGrid1.TextMatrix(x, 2) = rs!Bulan 'Bulan
    If Not IsNull(rs!Tahun) Then Frm48.MSFlexGrid1.TextMatrix(x, 3) = rs!Tahun 'Tahun
    If Not IsNull(rs!TarikhMula) Then Frm48.MSFlexGrid1.TextMatrix(x, 4) = rs!TarikhMula 'Tarikh Mula
    If Not IsNull(rs!TarikhAkhir) Then Frm48.MSFlexGrid1.TextMatrix(x, 5) = rs!TarikhAkhir 'Tarikh Akhir
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
End Sub
Sub Frm48_CalcDefault()
'On Error Resume Next
Frm48.CBB3.Clear
Frm48.CBB4.Clear

'%%%% TukangemaS %%%%
Frm48.L29_Text = "0.00"
Frm48.L30_Text = "0.00"
Frm48.L31_Text = "0.00"
Frm48.L32_Text = "0.00"
Frm48.L33_Text = "0.00"
Frm48.L34_Text = "0.00"
Frm48.L36_Text = "0.00"
Frm48.L35_Text.Visible = False
Frm48.CMD7.Visible = False
'%%%% TukangemaS %%%%

Frm48.TB11 = "0.00"
Frm48.TB12 = "0.00"
Frm48.TB13 = "0.00"

Frm48.CB1 = 1
Frm48.CB2 = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from tetapan_Payslip", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!Bulan) And Not IsNull(rs!Tahun) Then
        Frm48.CBB3.AddItem rs!Bulan & " " & rs!Tahun
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where status='" & "Aktif" & "' AND InvestorSmall = 0 AND InvestorBig = 0", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    If rs!user_level = 1 Or rs!user_level = 2 Or rs!user_level = 3 Or rs!user_level = G_LEVEL_USER Then Frm48.CBB4.AddItem rs!Samaran & "  |  " & rs!NoPekerja
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting where default1='" & G_CAWANGAN & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    'If rs!Default1 = "Default" Then
        'If Not IsNull(rs!komisen) Then Frm48.TB6 = Format(rs!komisen, "0.00") 'Komisen (%)
        
        '%%% TukangemaS %%%%
        If Not IsNull(rs!komisen) Then 'Kadar Komisen Barang Kemas (%)
            If IsNumeric(rs!komisen) Then
                Frm48.L30_Text = Format(rs!komisen, "0.00")
            Else
                Frm48.L30_Text = "0.00"
            End If
        Else
            Frm48.L30_Text = "0.00"
        End If
        If Not IsNull(rs!komisen_permata) Then 'Kadar Komisen Barang Permata (%)
            If IsNumeric(rs!komisen_permata) Then
                Frm48.L33_Text = Format(rs!komisen_permata, "0.00")
            Else
                Frm48.L33_Text = "0.00"
            End If
        Else
            Frm48.L33_Text = "0.00"
        End If
        If Not IsNull(rs!komisen_per_gram) Then 'Kadar Komisen Barang Permata (%)
            If IsNumeric(rs!komisen_per_gram) Then
                Frm48.L41_Text = Format(rs!komisen_per_gram, "0.00")
            Else
                Frm48.L41_Text = "0.00"
            End If
        Else
            Frm48.L41_Text = "0.00"
        End If
        '%%% TukangemaS %%%%
        
        If Not IsNull(rs!Profit) Then Frm48.L9_Text = Format(rs!Profit, "0.00") 'Profit (%)
    'End If
End If

rs.Close
Set rs = Nothing

Frm48.TB1 = vbNullString
Frm48.TB2 = vbNullString
Frm48.TB3 = vbNullString
Frm48.TB4 = vbNullString
Frm48.TB5 = vbNullString
Frm48.TB7 = vbNullString
Frm48.TB8 = vbNullString
Frm48.TB9 = "0.00"
Frm48.TB10 = "0.00"
Frm48.TB15 = "0.00"
Frm48.L6_Text = vbNullString
Frm48.L7_Text = vbNullString
Frm48.L13_Text = vbNullString
Frm48.L10_Text = 0
Frm48.L3_Text.Visible = False

Frm48.MSFlexGrid2.Clear
Frm48.MSFlexGrid2.RowHeight(0) = 600
Frm48.MSFlexGrid2.FormatString = "No.|<No.|<No. Siri|<Kategori|<Tarikh Jualan|<Berat Jualan (g)|<Harga Jualan (RM)"

Frm48.MSFlexGrid2.Rows = 1
Frm48.MSFlexGrid2.ColWidth(0) = 700
Frm48.MSFlexGrid2.ColWidth(1) = 0
Frm48.MSFlexGrid2.ColWidth(2) = 3450
Frm48.MSFlexGrid2.ColWidth(3) = 3450
Frm48.MSFlexGrid2.ColWidth(4) = 2750
Frm48.MSFlexGrid2.ColWidth(5) = 2750
Frm48.MSFlexGrid2.ColWidth(6) = 2750
End Sub
Sub Frm48_CalcGross()
Dim a As Double 'Gaji Pokok
Dim b As Double 'Elaun
Dim c As Double 'Komisen
Dim d As Double 'Profit
Dim e As Double 'KWSP
Dim f As Double 'Socso
Dim g As Double 'Komisen Investor

a = 0 'Gaji Pokok
b = 0 'Elaun
c = 0 'Komisen
d = 0 'Profit
e = 0 'KWSP
f = 0 'Socso
g = 0 'Komisen Investor

If IsNumeric(Frm48.TB3) Then
    a = Frm48.TB3 'Gaji Pokok
End If
If IsNumeric(Frm48.TB4) Then
    b = Frm48.TB4 'Elaun
End If
If IsNumeric(Frm48.TB7) Then
    c = Frm48.TB7 'Komisen
End If
If IsNumeric(Frm48.TB8) Then
    d = Frm48.TB8 'Profit
End If
If IsNumeric(Frm48.TB9) Then
    e = Frm48.TB9 'KWSP
End If
If IsNumeric(Frm48.TB10) Then
    f = Frm48.TB10 'Socso
End If
If IsNumeric(Frm48.TB14) Then
    g = Frm48.TB14 'Komisen Investor
End If

Frm48.TB11 = Format(a + b + c + d + g, "0.00") 'Gaji Kasar
Frm48.TB12 = Format(e + f, "0.00") 'Penolakan
Frm48.TB13 = Format((a + b + c + d + g) - (e + f), "0.00") 'Bersih
End Sub
Sub Frm48_RekodPayslip()
'On Error Resume Next
Frm48.MSFlexGrid3.Clear
Frm48.MSFlexGrid3.RowHeight(0) = 1000
Frm48.MSFlexGrid3.FormatString = "No.|<No.|<Bulan Payroll|<Nama|<No. Kad Pengenalan|<Gaji Pokok (RM)|<Elaun (RM)|<Overtime (RM)|<Elaun Perjalanan (RM)|<Lain-lain (RM)|<Jumlah Komisen (RM)|<KWSP (RM)|<Socso (RM)|<Lain-lain (RM)|<Zakat (RM)|<Income Tax (RM)|<Advance (RM)|<Pendapatan Kasar (RM)|<Jumlah Penolakan (RM)|<Pendapatan Bersih (RM)"

Frm48.MSFlexGrid3.Rows = 1
Frm48.MSFlexGrid3.ColWidth(0) = 600
Frm48.MSFlexGrid3.ColWidth(1) = 0
Frm48.MSFlexGrid3.ColWidth(2) = 1500
Frm48.MSFlexGrid3.ColWidth(3) = 4000
Frm48.MSFlexGrid3.ColWidth(4) = 2400
Frm48.MSFlexGrid3.ColWidth(5) = 1450
Frm48.MSFlexGrid3.ColWidth(6) = 1450
Frm48.MSFlexGrid3.ColWidth(7) = 1450
Frm48.MSFlexGrid3.ColWidth(8) = 1450
Frm48.MSFlexGrid3.ColWidth(9) = 1450
Frm48.MSFlexGrid3.ColWidth(10) = 1450
Frm48.MSFlexGrid3.ColWidth(11) = 1450
Frm48.MSFlexGrid3.ColWidth(12) = 1450
Frm48.MSFlexGrid3.ColWidth(13) = 1450
Frm48.MSFlexGrid3.ColWidth(14) = 1450
Frm48.MSFlexGrid3.ColWidth(15) = 1450
Frm48.MSFlexGrid3.ColWidth(16) = 1450
Frm48.MSFlexGrid3.ColWidth(17) = 1450
Frm48.MSFlexGrid3.ColWidth(18) = 1450
Frm48.MSFlexGrid3.ColWidth(19) = 1450
'Frm48.MSFlexGrid3.ColWidth(20) = 1450

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from payslip order by ID DESC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm48.MSFlexGrid3.Rows = x + 1
    Frm48.MSFlexGrid3.TextMatrix(x, 0) = x
    Frm48.MSFlexGrid3.TextMatrix(x, 1) = x
    If Not IsNull(rs!payroll_bulan) Then Frm48.MSFlexGrid3.TextMatrix(x, 2) = rs!payroll_bulan 'Bulan Payroll
    If Not IsNull(rs!payroll_namapenuh) Then Frm48.MSFlexGrid3.TextMatrix(x, 3) = rs!payroll_namapenuh 'Nama Penuh
    If Not IsNull(rs!payroll_ic) Then Frm48.MSFlexGrid3.TextMatrix(x, 4) = rs!payroll_ic 'No IC
    If Not IsNull(rs!payroll_gajipokok) Then Frm48.MSFlexGrid3.TextMatrix(x, 5) = Format(rs!payroll_gajipokok, "#,##0.00") 'Gaji Pokok (RM)
    If Not IsNull(rs!payroll_elaun) Then Frm48.MSFlexGrid3.TextMatrix(x, 6) = Format(rs!payroll_elaun, "#,##0.00") 'Elaun (RM)
    If Not IsNull(rs!overtime) Then Frm48.MSFlexGrid3.TextMatrix(x, 7) = Format(rs!overtime, "#,##0.00") 'Overtime (RM)
    If Not IsNull(rs!elaun_perjalanan) Then Frm48.MSFlexGrid3.TextMatrix(x, 8) = Format(rs!elaun_perjalanan, "#,##0.00") 'Elaun Perjalanan (RM)
    If Not IsNull(rs!pendapatan_lain) Then Frm48.MSFlexGrid3.TextMatrix(x, 9) = Format(rs!pendapatan_lain, "#,##0.00") 'Lain-lain (RM)
    If Not IsNull(rs!payroll_jumlah_komisen) Then Frm48.MSFlexGrid3.TextMatrix(x, 10) = Format(rs!payroll_jumlah_komisen, "#,##0.00") 'Jumlah Komisen (RM)
    If Not IsNull(rs!payroll_kwsp) Then Frm48.MSFlexGrid3.TextMatrix(x, 11) = Format(rs!payroll_kwsp, "#,##0.00") 'KWSP (RM)
    If Not IsNull(rs!payroll_socso) Then Frm48.MSFlexGrid3.TextMatrix(x, 12) = Format(rs!payroll_socso, "#,##0.00") 'Socso (RM)
    If Not IsNull(rs!payroll_lain) Then Frm48.MSFlexGrid3.TextMatrix(x, 13) = Format(rs!payroll_lain, "#,##0.00") 'Lain-lain (RM)
    If Not IsNull(rs!zakat) Then Frm48.MSFlexGrid3.TextMatrix(x, 14) = Format(rs!zakat, "#,##0.00") 'Zakat (RM)
    If Not IsNull(rs!tax) Then Frm48.MSFlexGrid3.TextMatrix(x, 15) = Format(rs!tax, "#,##0.00") 'Income Tax (RM)
    If Not IsNull(rs!advance) Then Frm48.MSFlexGrid3.TextMatrix(x, 16) = Format(rs!advance, "#,##0.00") 'Advance (RM)
    If Not IsNull(rs!payroll_kasar) Then Frm48.MSFlexGrid3.TextMatrix(x, 17) = Format(rs!payroll_kasar, "#,##0.00") 'Pendapatan Kasar (RM)
    If Not IsNull(rs!payroll_tolak) Then Frm48.MSFlexGrid3.TextMatrix(x, 18) = Format(rs!payroll_tolak, "#,##0.00") 'Jumlah Penolakan (RM)
    If Not IsNull(rs!payroll_bersih) Then Frm48.MSFlexGrid3.TextMatrix(x, 19) = Format(rs!payroll_bersih, "#,##0.00") 'Pendapatan Bersih (RM)
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
End Sub
Sub Frm48_M_cetak_payslip()
'On Error Resume Next

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

Report7_1.Sections("Section1").Controls("L1").Caption = vbNullString 'Bulan
Report7_1.Sections("Section1").Controls("L2").Caption = vbNullString 'Nama
Report7_1.Sections("Section1").Controls("L3").Caption = vbNullString 'No. Kad Pengenalan
Report7_1.Sections("Section1").Controls("L4").Caption = vbNullString 'No. EPF
Report7_1.Sections("Section1").Controls("L5").Caption = vbNullString 'No. Income Tax
Report7_1.Sections("Section1").Controls("L6").Caption = "0.00" 'Gaji Pokok
Report7_1.Sections("Section1").Controls("L7").Caption = "0.00" 'Elaun
Report7_1.Sections("Section1").Controls("L8").Caption = "0.00" 'Caruman KWSP
Report7_1.Sections("Section1").Controls("L9").Caption = "0.00" 'Socso
Report7_1.Sections("Section1").Controls("L10").Caption = "0.00" 'Jumlah Pendapatan
Report7_1.Sections("Section1").Controls("L11").Caption = "0.00" 'Jumlah Potongan
Report7_1.Sections("Section1").Controls("L12").Caption = "0.00" 'Gaji Bersih
Report7_1.Sections("Section1").Controls("L13").Caption = vbNullString 'Bank
Report7_1.Sections("Section1").Controls("L14").Caption = vbNullString 'No. Akaun
Report7_1.Sections("Section1").Controls("L15").Caption = "0.00" 'Komisen (RM)
Report7_1.Sections("Section1").Controls("L16").Caption = "0.00" 'Lain-lain (RM)
Report7_1.Sections("Section1").Controls("L17").Caption = "0.00" 'Overtime (RM)
Report7_1.Sections("Section1").Controls("L18").Caption = "0.00" 'Elaun Perjalanan (RM)
Report7_1.Sections("Section1").Controls("L19").Caption = "0.00" 'Lain-lain (RM)
Report7_1.Sections("Section1").Controls("L20").Caption = "0.00" 'Zakat (RM)
Report7_1.Sections("Section1").Controls("L21").Caption = "0.00" 'Income Tax (RM)
Report7_1.Sections("Section1").Controls("L22").Caption = "0.00" 'Advance (RM)

Report7_1.Sections("Section1").Controls("L1").Caption = G_PAYSLIP_BULAN 'Bulan

'### Reset maklumat kedai ### - Start
Report7_1.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report7_1.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report7_1.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report7_1.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report7_1.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report7_1.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report7_1.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report7_1.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report7_1.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report7_1.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where NoIC='" & G_PAYSLIP_IC & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!Nama) Then Report7_1.Sections("Section1").Controls("L2").Caption = rs!Nama 'Nama
    If Not IsNull(rs!NoIC) Then Report7_1.Sections("Section1").Controls("L3").Caption = rs!NoIC 'No. Kad Pengenalan
    If Not IsNull(rs!NoKWSP) Then Report7_1.Sections("Section1").Controls("L4").Caption = rs!NoKWSP 'No. EPF
    If Not IsNull(rs!NoSocso) Then Report7_1.Sections("Section1").Controls("L5").Caption = rs!NoSocso 'No. Income Tax
    If Not IsNull(rs!alamat2) Then Report7_1.Sections("Section1").Controls("L13").Caption = rs!alamat2 'Bank
    If Not IsNull(rs!alamat3) Then Report7_1.Sections("Section1").Controls("L14").Caption = rs!alamat3 'No. Akaun
End If

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from payslip where payroll_bulan='" & G_PAYSLIP_BULAN & "' AND payroll_ic='" & G_PAYSLIP_IC & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!payroll_gajipokok) Then Report7_1.Sections("Section1").Controls("L6").Caption = Format(rs!payroll_gajipokok, "#,##0.00") 'Gaji Pokok
    If Not IsNull(rs!payroll_elaun) Then Report7_1.Sections("Section1").Controls("L7").Caption = Format(rs!payroll_elaun, "#,##0.00") 'Elaun
    If Not IsNull(rs!payroll_kwsp) Then Report7_1.Sections("Section1").Controls("L8").Caption = Format(rs!payroll_kwsp, "#,##0.00") 'Caruman KWSP
    If Not IsNull(rs!payroll_socso) Then Report7_1.Sections("Section1").Controls("L9").Caption = Format(rs!payroll_socso, "#,##0.00") 'Socso
    If Not IsNull(rs!payroll_kasar) Then Report7_1.Sections("Section1").Controls("L10").Caption = Format(rs!payroll_kasar, "#,##0.00") 'Jumlah Pendapatan
    If Not IsNull(rs!payroll_tolak) Then Report7_1.Sections("Section1").Controls("L11").Caption = Format(rs!payroll_tolak, "#,##0.00") 'Jumlah Potongan
    If Not IsNull(rs!payroll_bersih) Then Report7_1.Sections("Section1").Controls("L12").Caption = Format(rs!payroll_bersih, "#,##0.00") 'Gaji Bersih
    If Not IsNull(rs!payroll_jumlah_komisen) Then Report7_1.Sections("Section1").Controls("L15").Caption = Format(rs!payroll_jumlah_komisen, "#,##0.00") 'Komisen
    If Not IsNull(rs!payroll_lain) Then Report7_1.Sections("Section1").Controls("L16").Caption = Format(rs!payroll_lain, "#,##0.00") 'Lain-lain
    If Not IsNull(rs!overtime) Then Report7_1.Sections("Section1").Controls("L17").Caption = Format(rs!overtime, "#,##0.00") 'Overtime (RM)
    If Not IsNull(rs!elaun_perjalanan) Then Report7_1.Sections("Section1").Controls("L18").Caption = Format(rs!elaun_perjalanan, "#,##0.00") 'Elaun Perjalanan (RM)
    If Not IsNull(rs!pendapatan_lain) Then Report7_1.Sections("Section1").Controls("L19").Caption = Format(rs!pendapatan_lain, "#,##0.00") 'Lain-lain (RM)
    If Not IsNull(rs!zakat) Then Report7_1.Sections("Section1").Controls("L20").Caption = Format(rs!zakat, "#,##0.00") 'Zakat (RM)
    If Not IsNull(rs!tax) Then Report7_1.Sections("Section1").Controls("L21").Caption = Format(rs!tax, "#,##0.00") 'Income Tax (RM)
    If Not IsNull(rs!advance) Then Report7_1.Sections("Section1").Controls("L22").Caption = Format(rs!advance, "#,##0.00") 'Advance (RM)

    Set Report7_1.DataSource = rs
    Report7_1.Show
End If

'rs.Close
Set rs = Nothing

End Sub
Sub frm48_reset_gaji()
'on error resume next
Frm48.L10_Text = 0
Frm48.TB9 = Format(0, "#,##0.00")
Frm48.TB10 = Format(0, "#,##0.00")
Frm48.TB11 = Format(0, "#,##0.00")
Frm48.TB12 = Format(0, "#,##0.00")
Frm48.TB13 = Format(0, "#,##0.00")
Frm48.TB15 = Format(0, "#,##0.00")
Frm48.TB16 = Format(0, "#,##0.00")
Frm48.TB17 = Format(0, "#,##0.00")
Frm48.TB18 = Format(0, "#,##0.00")
Frm48.TB19 = Format(0, "#,##0.00")
Frm48.TB20 = Format(0, "#,##0.00")
Frm48.TB21 = Format(0, "#,##0.00")
Frm48.L40_Text = Format(0, "#,##0.00")
Frm48.L41_Text = Format(0, "#,##0.00")
Frm48.L42_Text = Format(0, "#,##0.00")
Frm48.L29_Text = Format(0, "#,##0.00")
Frm48.L30_Text = Format(0, "#,##0.00")
Frm48.L31_Text = Format(0, "#,##0.00")
Frm48.L32_Text = Format(0, "#,##0.00")
Frm48.L33_Text = Format(0, "#,##0.00")
Frm48.L34_Text = Format(0, "#,##0.00")
Frm48.L36_Text = Format(0, "#,##0.00")
End Sub
Sub frm48_kiraan_pendapatan()
'on error resume next
Dim Frm48_LM_GAJI As Double
Dim Frm48_LM_ELAUN As Double
Dim Frm48_LM_OT As Double
Dim Frm48_LM_ELAUN_TRANSPORT As Double
Dim Frm48_LM_LAIN As Double
Dim Frm48_LM_KOMISEN As Double

Frm48_LM_GAJI = 0
Frm48_LM_ELAUN = 0
Frm48_LM_OT = 0
Frm48_LM_ELAUN_TRANSPORT = 0
Frm48_LM_LAIN = 0
Frm48_LM_KOMISEN = 0

If (Frm48.TB3 <> vbNullString And IsNumeric(Frm48.TB3)) Then Frm48_LM_GAJI = Frm48.TB3
If (Frm48.TB4 <> vbNullString And IsNumeric(Frm48.TB4)) Then Frm48_LM_ELAUN = Frm48.TB4
If (Frm48.TB16 <> vbNullString And IsNumeric(Frm48.TB16)) Then Frm48_LM_OT = Frm48.TB16
If (Frm48.TB17 <> vbNullString And IsNumeric(Frm48.TB17)) Then Frm48_LM_ELAUN_TRANSPORT = Frm48.TB17
If (Frm48.TB18 <> vbNullString And IsNumeric(Frm48.TB18)) Then Frm48_LM_LAIN = Frm48.TB18
If (Frm48.L36_Text <> vbNullString And IsNumeric(Frm48.L36_Text)) Then Frm48_LM_KOMISEN = Frm48.L36_Text

Frm48.TB11 = Format(Frm48_LM_GAJI + Frm48_LM_ELAUN + Frm48_LM_OT + Frm48_LM_ELAUN_TRANSPORT + Frm48_LM_LAIN + Frm48_LM_KOMISEN, "#,##0.00")
End Sub
Sub frm48_kiraan_tolakan()
'on error resume next
Dim Frm48_LM_KWSP As Double
Dim Frm48_LM_SOCSO As Double
Dim Frm48_LM_LAIN As Double
Dim Frm48_LM_ZAKAT As Double
Dim Frm48_LM_TAX As Double
Dim Frm48_LM_ADVANCE As Double

Frm48_LM_KWSP = 0
Frm48_LM_SOCSO = 0
Frm48_LM_LAIN = 0
Frm48_LM_ZAKAT = 0
Frm48_LM_TAX = 0
Frm48_LM_ADVANCE = 0

If (Frm48.TB9 <> vbNullString And IsNumeric(Frm48.TB9)) Then Frm48_LM_KWSP = Frm48.TB9
If (Frm48.TB10 <> vbNullString And IsNumeric(Frm48.TB10)) Then Frm48_LM_SOCSO = Frm48.TB10
If (Frm48.TB15 <> vbNullString And IsNumeric(Frm48.TB15)) Then Frm48_LM_LAIN = Frm48.TB15
If (Frm48.TB19 <> vbNullString And IsNumeric(Frm48.TB19)) Then Frm48_LM_ZAKAT = Frm48.TB19
If (Frm48.TB20 <> vbNullString And IsNumeric(Frm48.TB20)) Then Frm48_LM_TAX = Frm48.TB20
If (Frm48.TB21 <> vbNullString And IsNumeric(Frm48.TB21)) Then Frm48_LM_ADVANCE = Frm48.TB21

Frm48.TB12 = Format(Frm48_LM_KWSP + Frm48_LM_SOCSO + Frm48_LM_LAIN + Frm48_LM_ZAKAT + Frm48_LM_TAX + Frm48_LM_ADVANCE, "#,##0.00")
End Sub
Sub frm48_kiraan_bersih()
'on error resume next
Dim Frm48_LM_KASAR As Double
Dim Frm48_LM_TOLAK As Double

Frm48_LM_KASAR = 0
Frm48_LM_TOLAK = 0

If (Frm48.TB11 <> vbNullString And IsNumeric(Frm48.TB11)) Then Frm48_LM_KASAR = Frm48.TB11
If (Frm48.TB12 <> vbNullString And IsNumeric(Frm48.TB12)) Then Frm48_LM_TOLAK = Frm48.TB12

Frm48.TB13 = Format(Frm48_LM_KASAR - Frm48_LM_TOLAK, "#,##0.00")
End Sub
Sub frm48_kiraan_komisen_berat()
'on error resume next
Dim Frm48_LM_BERAT As Double
Dim Frm48_LM_KADAR As Double

Frm48_LM_BERAT = 0
Frm48_LM_KADAR = 0

If (Frm48.L40_Text <> vbNullString And IsNumeric(Frm48.L40_Text)) Then Frm48_LM_BERAT = Frm48.L40_Text
If (Frm48.L41_Text <> vbNullString And IsNumeric(Frm48.L41_Text)) Then Frm48_LM_KADAR = Frm48.L41_Text

Frm48.L42_Text = Format(Frm48_LM_BERAT * Frm48_LM_KADAR, "#,##0.00")
End Sub
Sub frm48_kiraan_komisen_bk()
'on error resume next
Dim Frm48_LM_HARGA As Double
Dim Frm48_LM_KADAR As Double

Frm48_LM_HARGA = 0
Frm48_LM_KADAR = 0

If (Frm48.L29_Text <> vbNullString And IsNumeric(Frm48.L29_Text)) Then Frm48_LM_HARGA = Frm48.L29_Text
If (Frm48.L30_Text <> vbNullString And IsNumeric(Frm48.L30_Text)) Then Frm48_LM_KADAR = Frm48.L30_Text

Frm48.L31_Text = Format(Frm48_LM_HARGA * (Frm48_LM_KADAR / 100), "#,##0.00")
End Sub
Sub frm48_kiraan_komisen_permata()
'on error resume next
Dim Frm48_LM_HARGA As Double
Dim Frm48_LM_KADAR As Double

Frm48_LM_HARGA = 0
Frm48_LM_KADAR = 0

If (Frm48.L32_Text <> vbNullString And IsNumeric(Frm48.L32_Text)) Then Frm48_LM_HARGA = Frm48.L32_Text
If (Frm48.L33_Text <> vbNullString And IsNumeric(Frm48.L33_Text)) Then Frm48_LM_KADAR = Frm48.L33_Text

Frm48.L34_Text = Format(Frm48_LM_HARGA * (Frm48_LM_KADAR / 100), "#,##0.00")
End Sub
Sub frm48_kiraan_komisen()
'on error resume next
Dim Frm48_LM_KOMISEN_BK As Double
Dim Frm48_LM_KOMISEN_PERMATA As Double
Dim Frm48_LM_KOMISEN_BERAT As Double

Frm48_LM_KOMISEN_BK = 0
Frm48_LM_KOMISEN_PERMATA = 0
Frm48_LM_KOMISEN_BERAT = 0

If (Frm48.L31_Text <> vbNullString And IsNumeric(Frm48.L31_Text)) Then Frm48_LM_KOMISEN_BK = Frm48.L31_Text
If (Frm48.L34_Text <> vbNullString And IsNumeric(Frm48.L34_Text)) Then Frm48_LM_KOMISEN_PERMATA = Frm48.L34_Text
If (Frm48.L42_Text <> vbNullString And IsNumeric(Frm48.L42_Text)) Then Frm48_LM_KOMISEN_BERAT = Frm48.L42_Text

Frm48.L36_Text = Format(Frm48_LM_KOMISEN_BK + Frm48_LM_KOMISEN_PERMATA + Frm48_LM_KOMISEN_BERAT, "#,##0.00")
End Sub
Sub frm48_pic_enable()
'on error resume next
Frm48.Pic1.Top = 360
Frm48.Pic2.Top = 360
Frm48.Pic3.Top = 360
Frm48.Pic1.Left = 120
Frm48.Pic2.Left = 120
Frm48.Pic3.Left = 120

Frm48.Pic1.Visible = False
Frm48.Pic2.Visible = False
Frm48.Pic3.Visible = False
End Sub
