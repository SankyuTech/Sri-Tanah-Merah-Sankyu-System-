Attribute VB_Name = "Module38"
Sub frm128_pic_ena_disable()
'On Error Resume Next
frm128.Frame1.Left = 150
frm128.Frame1.Top = 1900
frm128.Frame2.Left = 150
frm128.Frame2.Top = 1900
frm128.Frame3.Left = 150
frm128.Frame3.Top = 1900

frm128.Frame1.Visible = False
frm128.Frame2.Visible = False
frm128.Frame3.Visible = False
End Sub
Sub frm128_reset_data_utama()
'on error resume next
frm128.L1_Text = vbNullString
frm128.L2_Text = vbNullString
frm128.L3_Text = vbNullString
frm128.L4_Text = vbNullString
End Sub
Sub frm128_default_setting()
'On Error Resume Next
frm128.CBB1.Clear
frm128.CBB2.Clear

'###Senarai Nama Pekerja###
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where status='" & "Aktif" & "' AND InvestorSmall = 0 AND InvestorBig = 0", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    If rs!user_level = 1 Or rs!user_level = 2 Or rs!user_level = 3 Or rs!user_level = G_LEVEL_USER Then
        frm128.CBB1.AddItem rs!Samaran & "  |  " & rs!NoPekerja
        frm128.CBB2.AddItem rs!Samaran & "  |  " & rs!NoPekerja
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
End Sub
Sub frm128_reset_simpanan()
'On Error Resume Next
frm128.TB1 = "0.00"
frm128.DTPicker1 = DateTime.Date$
frm128.TB3 = vbNullString
End Sub
Sub frm128_reset_pulangan()
'On Error Resume Next
frm128.TB2 = "0.00"
frm128.DTPicker2 = DateTime.Date$
frm128.TB4 = vbNullString
End Sub
Sub frm128_jurujual()
'On Error Resume Next
If MDI_frm1.L3_Text <> vbNullString Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        frm128.CBB1 = rs!Samaran & "  |  " & rs!NoPekerja
        frm128.CBB2 = rs!Samaran & "  |  " & rs!NoPekerja
        
    Else
        
        frm128.CBB1.AddItem "" & "  |  " & rs!Samaran
        frm128.CBB1 = "" & "  |  " & rs!Samaran
        frm128.CBB2.AddItem "" & "  |  " & rs!Samaran
        frm128.CBB2 = "" & "  |  " & rs!Samaran
        
    End If
    
    rs.Close
    Set rs = Nothing

    If G_LOCK_JURUJUAL = "YES" Then
    
        frm128.CBB1.Enabled = False
        frm128.CBB1.BackColor = &H8000000A
        frm128.CBB2.Enabled = False
        frm128.CBB2.BackColor = &H8000000A
        
    Else
    
        frm128.CBB1.Enabled = True
        frm128.CBB1.BackColor = &HFFFFFF
        frm128.CBB2.Enabled = True
        frm128.CBB2.BackColor = &HFFFFFF

    End If
End If
End Sub
Sub frm128_report_simpanan_header()
'on error resume next

With frm128.LV1
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear

    frm128.LV1.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "Tarikh", 1700
    .ColumnHeaders.Add 5, , "Jenis", 2000
    .ColumnHeaders.Add 6, , "No. Rujukan", 2500
    .ColumnHeaders.Add 7, , "Jumlah (RM)", 2000, 1
    .ColumnHeaders.Add 8, , "Remarks", 8800
    .ColumnHeaders.Add 9, , "Cawangan", 2500
    
End With
End Sub
Sub frm128_report_simpanan()
'On Error Resume Next
Dim frm128_LM_TOTAL_PAGE As Double

Dim LM_SIMPAN As Double
Dim LM_REFUND As Double
Dim LM_GUNA As Double

frm128_PAGE_SIZE = 29
frm128_LM_TOTAL_PAGE = 0
x = 0

re_gen_report:

frm128.L26_Text = "0.00"
frm128.L27_Text = "0.00"
frm128.L28_Text = "0.00"
frm128.L29_Text = "0.00"

LM_START_ROW = frm128.L69_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + frm128_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If frm128.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - frm128_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        frm128.L67_Text = 1
    End If
End If

frm128_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 24_rekod_kewangan_pelanggan where no_rujukan_pelanggan='" & frm128.L4_Text & "' AND status = 1 order by tarikh ASC LIMIT " & LM_START_ROW & "," & frm128_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1
    If frm128_LM_PAGE_FOUND = 0 Then
        If frm128.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                frm128.L67_Text = frm128.L67_Text + 1 'Paparan Page ke-xxx
                frm128_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(frm128.L67_Text) Then
                    If frm128.L67_Text <> 1 Then
                        frm128.L67_Text = frm128.L67_Text - 1 'Paparan Page ke-xxx
                        frm128_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If

    Y = ((frm128.L67_Text - 1) * frm128_PAGE_SIZE) + x

    With frm128.LV1.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!tarikh) Then 'Tarikh
            .ListSubItems.Add , , rs!tarikh
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!jenis) Then
            
            If rs!jenis = 0 Then
            
                .ListSubItems.Add , , "Simpanan Duit"
                
            ElseIf rs!jenis = 1 Then
                
                .ListSubItems.Add , , "Penggunaan Duit"
                
            
            ElseIf rs!jenis = 2 Then
                
                .ListSubItems.Add , , "Pulangan Duit"
                
            End If
            
        End If
        
        If Not IsNull(rs!no_resit) Then 'No. Rujukan
            .ListSubItems.Add , , rs!no_resit
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!jumlah) Then 'Jumlah (RM)
            .ListSubItems.Add , , Format(rs!jumlah, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!jenis) Then
            
            If rs!jenis = 0 Then
            
                If Not IsNull(rs!jenis_penggunaan) Then .ListSubItems.Add , , rs!jenis_penggunaan
                
            ElseIf rs!jenis = 1 Then
                
                If Not IsNull(rs!jenis_penggunaan) Then
                    If rs!jenis_penggunaan = 0 Then
                        .ListSubItems.Add , , "Belian Barangan Kemas" 'Tujuaan Penggunaan
                    ElseIf rs!jenis_penggunaan = 1 Then
                        .ListSubItems.Add , , "Bayaran Ansuran Emas" 'Tujuaan Penggunaan
                    ElseIf rs!jenis_penggunaan = 2 Then
                        .ListSubItems.Add , , "Bayaran Deposit Tempahan Emas" 'Tujuaan Penggunaan
                    ElseIf rs!jenis_penggunaan = 3 Then
                        .ListSubItems.Add , , "Bayaran Servis" 'Tujuaan Penggunaan
                    ElseIf rs!jenis_penggunaan = 4 Then
                        .ListSubItems.Add , , "Bayaran Ambilan Tempahan Emas" 'Tujuaan Penggunaan
                    End If
                End If
            
            ElseIf rs!jenis = 2 Then
            
                .ListSubItems.Add , , rs!jenis_penggunaan
                
            End If
            
        End If
        
        If Not IsNull(rs!cawangan) Then
            .ListSubItems.Add , , rs!cawangan
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
rs.Open "select COUNT(ID) from 24_rekod_kewangan_pelanggan where no_rujukan_pelanggan='" & frm128.L4_Text & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    frm128_LM_TOTAL_PAGE = Format(rs(0) / frm128_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, frm128_LM_TOTAL_PAGE, ".") <> 0 Then
    
        frm128_LM_PAGE = Split(frm128_LM_TOTAL_PAGE, ".")(0)
        frm128_LM_PAGE_LEBIHAN = Split(frm128_LM_TOTAL_PAGE, ".")(1)
        
        If frm128_LM_PAGE_LEBIHAN <> "00" Then
            frm128.L68_Text = frm128_LM_PAGE + 1
        Else
            frm128.L68_Text = frm128_LM_PAGE
        End If
        
    Else
    
        frm128.L68_Text = frm128_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        frm128.L68_Text = 0
    End If
Else
    frm128.L68_Text = 0
End If

rs.Close
Set rs = Nothing

LM_SIMPAN = 0
LM_REFUND = 0
LM_GUNA = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah) from 24_rekod_kewangan_pelanggan where no_rujukan_pelanggan='" & frm128.L4_Text & "' AND status = 1 AND jenis = 0", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then LM_SIMPAN = Format(rs(0), "#,##0.00")

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah) from 24_rekod_kewangan_pelanggan where no_rujukan_pelanggan='" & frm128.L4_Text & "' AND status = 1 AND jenis = 2", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then LM_REFUND = Format(rs(0), "#,##0.00")

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah) from 24_rekod_kewangan_pelanggan where no_rujukan_pelanggan='" & frm128.L4_Text & "' AND status = 1 AND jenis = 1", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then LM_GUNA = Format(rs(0), "#,##0.00")

rs.Close
Set rs = Nothing

frm128.L26_Text = Format(LM_SIMPAN, "#,##0.00")
frm128.L27_Text = Format(LM_REFUND, "#,##0.00")
frm128.L28_Text = Format(LM_GUNA, "#,##0.00")
frm128.L29_Text = Format(LM_SIMPAN - LM_REFUND - LM_GUNA, "#,##0.00")

If x <> 0 Then
    frm128.L69_Text = LM_START_ROW
End If

If frm128.L67_Text <> vbNullString And IsNumeric(frm128.L67_Text) Then
    If frm128.L68_Text <> vbNullString And IsNumeric(frm128.L68_Text) Then
        frm128_LM_CURR_PAGE = frm128.L67_Text
        frm128_LM_TOTAL_PAGE = frm128.L68_Text
        
        If frm128_LM_CURR_PAGE > frm128_LM_TOTAL_PAGE Then
            
            frm128.L67_Text = frm128.L67_Text - 1
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            GoTo re_gen_report:
            
        End If
    End If
End If
End Sub
Sub frm128_cetak_pv()
'on error resume next
DATA_FOUND = 0
Frm84_DATA_PEKERJA_FOUND = 0


Report81.Sections("Section2").Controls("L5").Caption = vbNullString 'Nama kedai
Report81.Sections("Section2").Controls("L7").Caption = "No. ID GST : -" 'No ID GST
Report81.Sections("Section2").Controls("L3").Caption = vbNullString 'No. Payment Voucher
Report81.Sections("Section2").Controls("L4").Caption = vbNullString 'Date
Report81.Sections("Section2").Controls("L17").Caption = vbNullString 'Staff

Report81.Sections("Section1").Controls("L14").Caption = vbNullString
Report81.Sections("Section1").Controls("L8").Caption = "1" 'Kuantiti
Report81.Sections("Section1").Controls("L9").Caption = "0.00" 'Unit price
Report81.Sections("Section1").Controls("L10").Caption = "0.00" 'Total

Report81.Sections("Section1").Controls("L11").Caption = "0.00" 'Exclude GST
Report81.Sections("Section1").Controls("L12").Caption = "0.00" 'GST
Report81.Sections("Section1").Controls("L13").Caption = "0.00" 'Include GST

'### Reset maklumat kedai ### - Start
Report81.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report81.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report81.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report81.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report81.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report81.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report81.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report81.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report81.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report81.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Report81.Sections("Section1").Controls("L11").Visible = False
Report81.Sections("Section1").Controls("L12").Visible = False
Report81.Sections("Section1").Controls("L18").Visible = False
Report81.Sections("Section1").Controls("L19").Visible = False
Report81.Sections("Section1").Controls("L20").Caption = "Total : RM"
LM_NO_PELANGGAN = vbNullString

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 24_rekod_kewangan_pelanggan where status = 1 AND no_resit='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Not IsNull(rs!no_rujukan_pelanggan) Then LM_NO_PELANGGAN = rs!no_rujukan_pelanggan
    Report81.Sections("Section2").Controls("L3").Caption = G_No_RESIT_JUALAN
    If Not IsNull(rs!tarikh) Then Report81.Sections("Section2").Controls("L4").Caption = rs!tarikh
    If Not IsNull(rs!jenis_penggunaan) Then Report81.Sections("Section1").Controls("L14").Caption = rs!jenis_penggunaan
    If Not IsNull(rs!jumlah) Then
        Report81.Sections("Section1").Controls("L9").Caption = Format(rs!jumlah, "#,##0.00")
        Report81.Sections("Section1").Controls("L10").Caption = Format(rs!jumlah, "#,##0.00")
        Report81.Sections("Section1").Controls("L13").Caption = Format(rs!jumlah, "#,##0.00")
    End If
    'If Not IsNull(rs!jumlah_tanpa_gst) Then Report81.Sections("Section1").Controls("L11").Caption = Format(rs!harga_dengan_gst - rs!gst_sr_cukai, "#,##0.00")
    'If Not IsNull(rs!gst_sr_cukai) Then Report81.Sections("Section1").Controls("L12").Caption = Format(rs!gst_sr_cukai, "#,##0.00")
    'If Not IsNull(rs!harga_dengan_gst) Then Report81.Sections("Section1").Controls("L13").Caption = Format(rs!harga_dengan_gst, "#,##0.00")

    If Not IsNull(rs!no_rujukan_pekerja) Then
        Frm84_LM_No_PEKERJA = rs!no_rujukan_pekerja
        Frm84_DATA_PEKERJA_FOUND = 1
    End If

    DATA_FOUND = 1
End If

rs.Close
Set rs = Nothing

If DATA_FOUND = 1 Then
    
    If LM_NO_PELANGGAN <> vbNullString Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where no_pelanggan='" & LM_NO_PELANGGAN & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            If Not IsNull(rs!Nama) Then Report81.Sections("Section2").Controls("L5").Caption = rs!Nama 'Nama
            If Not IsNull(rs!no_tel) Then Report81.Sections("Section2").Controls("L7").Caption = rs!no_tel 'No. Tel
            
        End If
        
        rs.Close
        Set rs = Nothing
        
    End If

    If Frm84_DATA_PEKERJA_FOUND = 1 Then
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where NoPekerja='" & Frm84_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!Samaran) Then Report81.Sections("Section2").Controls("L17").Caption = rs!Samaran 'Nama Samaran
        End If

        rs.Close
        Set rs = Nothing
    End If

    '### Paparan Resit ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 24_rekod_kewangan_pelanggan where status = 1 AND no_resit='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        Set Report81.DataSource = rs
        If G_PREVIEW = 1 Then Report81.Show
        rs.MoveNext
    Wend
    
    'rs.Close
    Set rs = Nothing
    '### Paparan Resit ### - End
    
    If G_PREVIEW = 0 Then Report81.PrintReport
     
    G_No_RESIT_JUALAN = vbNullString
End If
End Sub
Sub frm128_cetak_receipt()
'on error resume next
DATA_FOUND = 0
Frm84_DATA_PEKERJA_FOUND = 0


Report81.Sections("Section2").Controls("L5").Caption = vbNullString 'Nama kedai
Report81.Sections("Section2").Controls("L7").Caption = "No. ID GST : -" 'No ID GST
Report81.Sections("Section2").Controls("L3").Caption = vbNullString 'No. Payment Voucher
Report81.Sections("Section2").Controls("L4").Caption = vbNullString 'Date
Report81.Sections("Section2").Controls("L17").Caption = vbNullString 'Staff

Report81.Sections("Section1").Controls("L14").Caption = vbNullString
Report81.Sections("Section1").Controls("L8").Caption = "1" 'Kuantiti
Report81.Sections("Section1").Controls("L9").Caption = "0.00" 'Unit price
Report81.Sections("Section1").Controls("L10").Caption = "0.00" 'Total

Report81.Sections("Section1").Controls("L11").Caption = "0.00" 'Exclude GST
Report81.Sections("Section1").Controls("L12").Caption = "0.00" 'GST
Report81.Sections("Section1").Controls("L13").Caption = "0.00" 'Include GST

'### Reset maklumat kedai ### - Start
Report81.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report81.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report81.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report81.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report81.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report81.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report81.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report81.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report81.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report81.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Report81.Sections("Section2").Controls("L21").Caption = "Paid By"
Report81.Sections("Section2").Controls("L205").Caption = "Receipt"
Report81.Sections("Section2").Controls("L22").Caption = "Receipt No."
Report81.Sections("Section1").Controls("L23").Caption = "Payment for the followings :"
Report81.Caption = "Receipt"
Report81.Sections("Section1").Controls("L11").Visible = False
Report81.Sections("Section1").Controls("L12").Visible = False
Report81.Sections("Section1").Controls("L18").Visible = False
Report81.Sections("Section1").Controls("L19").Visible = False
Report81.Sections("Section1").Controls("L20").Caption = "Total : RM"
Report81.Sections("Section1").Controls("L26").Caption = "Paid By :"

LM_NO_PELANGGAN = vbNullString

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 24_rekod_kewangan_pelanggan where status = 1 AND no_resit='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Not IsNull(rs!no_rujukan_pelanggan) Then LM_NO_PELANGGAN = rs!no_rujukan_pelanggan
    Report81.Sections("Section2").Controls("L3").Caption = G_No_RESIT_JUALAN
    If Not IsNull(rs!tarikh) Then Report81.Sections("Section2").Controls("L4").Caption = rs!tarikh
    If Not IsNull(rs!jenis_penggunaan) Then Report81.Sections("Section1").Controls("L14").Caption = rs!jenis_penggunaan
    If Not IsNull(rs!jumlah) Then
        Report81.Sections("Section1").Controls("L9").Caption = Format(rs!jumlah, "#,##0.00")
        Report81.Sections("Section1").Controls("L10").Caption = Format(rs!jumlah, "#,##0.00")
        Report81.Sections("Section1").Controls("L13").Caption = Format(rs!jumlah, "#,##0.00")
    End If
    'If Not IsNull(rs!jumlah_tanpa_gst) Then Report81.Sections("Section1").Controls("L11").Caption = Format(rs!harga_dengan_gst - rs!gst_sr_cukai, "#,##0.00")
    'If Not IsNull(rs!gst_sr_cukai) Then Report81.Sections("Section1").Controls("L12").Caption = Format(rs!gst_sr_cukai, "#,##0.00")
    'If Not IsNull(rs!harga_dengan_gst) Then Report81.Sections("Section1").Controls("L13").Caption = Format(rs!harga_dengan_gst, "#,##0.00")

    If Not IsNull(rs!no_rujukan_pekerja) Then
        Frm84_LM_No_PEKERJA = rs!no_rujukan_pekerja
        Frm84_DATA_PEKERJA_FOUND = 1
    End If

    DATA_FOUND = 1
End If

rs.Close
Set rs = Nothing

If DATA_FOUND = 1 Then
    
    If LM_NO_PELANGGAN <> vbNullString Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where no_pelanggan='" & LM_NO_PELANGGAN & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            If Not IsNull(rs!Nama) Then Report81.Sections("Section2").Controls("L5").Caption = rs!Nama 'Nama
            If Not IsNull(rs!no_tel) Then Report81.Sections("Section2").Controls("L7").Caption = rs!no_tel 'No. Tel
            
        End If
        
        rs.Close
        Set rs = Nothing
        
    End If

    If Frm84_DATA_PEKERJA_FOUND = 1 Then
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where NoPekerja='" & Frm84_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!Samaran) Then Report81.Sections("Section2").Controls("L17").Caption = rs!Samaran 'Nama Samaran
        End If

        rs.Close
        Set rs = Nothing
    End If

    '### Paparan Resit ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 24_rekod_kewangan_pelanggan where status = 1 AND no_resit='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        Set Report81.DataSource = rs
        If G_PREVIEW = 1 Then Report81.Show
        rs.MoveNext
    Wend
    
    'rs.Close
    Set rs = Nothing
    '### Paparan Resit ### - End
    
    If G_PREVIEW = 0 Then Report81.PrintReport
     
    G_No_RESIT_JUALAN = vbNullString
End If
End Sub
