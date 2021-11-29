Attribute VB_Name = "Module8"
Sub Frm93_initial_setting()
'on error resume next
GLOBAL_DISABLE = 0

Frm93.CB2 = 0
Frm93.CB3 = 0
Frm93.CB6 = 0
Frm93.CB9 = 0
Frm93.CB13 = 0
Frm93.CB14 = 0
Frm93.CB15 = 0
Frm93.CB16 = 0
Frm93.CB17 = 0
'Frm93.CB18 = 0
Frm93.CB19 = 1
Frm93.CB20 = 0
'Frm93.CB23 = 1
'Frm93.CB24 = 0
'Frm93.CB25 = 0

Frm93.TB1 = "0.00" 'Anggaran Berat g
Frm93.TB2 = "0.00" 'Harga Semasa RM/g
Frm93.TB3 = "0.00" 'Upah RM
Frm93.TB4 = "0.00" 'Anggaran Harga RM
Frm93.TB5 = vbNullString 'No. Siri Produk
Frm93.TB6 = vbNullString 'No. Siri Produk
Frm93.TB7 = "0.00" 'Berat Asal g
Frm93.TB8 = "0.00" 'Berat Jualan g
Frm93.TB9 = "0.00" 'Harga Semasa RM/g
Frm93.TB10 = "0.00" 'Upah RM
Frm93.TB11 = "0.00" 'Harga Asal RM
Frm93.TB12 = "0.00" 'Adjustment RM
Frm93.TB13 = "0.00" 'Harga Jualan RM
'Frm93.TB16 = "0.00" 'Jumlah Cukai GST RM
Frm93.TB17 = "0.00" 'Jumlah Nilaian Resit Trade In RM
Frm93.TB18 = vbNullString 'Carian No. Resit Trade In
'Frm93.TB19 = "0.00" 'Jumlah Deposit Dengan GST (RM)
Frm93.TB20 = "0.00" 'Deposit RM
Frm93.TB22 = "0.00" 'Deposit Trade In RM
Frm93.TB23 = "0.00" 'Jumlah Deposit RM
'Frm93.TB24 = "0.00" 'Adjustment RM
'Frm93.TB25 = vbNullString 'Carian : No. Siri Produk
'Frm93.TB26 = vbNullString 'Carian : No. Invoice

frm130.TB27 = "0.00" 'Cara Bayaran : Tunai
frm130.TB28 = "0.00" 'Cara Bayaran : Bank In
frm130.TB29 = "0.00" 'Cara Bayaran : Kad Kredit
'Frm93.TB30 = "0.00" 'Cara Bayaran : Cas Kad Kredit
'Frm93.TB31 = "0.00" 'Cara Bayaran : Jumlah Potongan Kad Kredit
frm130.TB21 = "0.00" 'Cara Bayaran : Duit Simpanan Di Kedai
'Frm93.TB38 = "0.00" 'Cara Bayaran : Kad Debit
'Frm93.TB39 = "0.00" 'Cara Bayaran : Cas Kad Debit
'Frm93.TB40 = "0.00" 'Cara Bayaran : Jumlah Potongan Kad Debit
Frm93.TB32 = "0.00" 'Cara Bayaran : Jumlah Bayaran
Frm93.TB33 = vbNullString 'Remarks
Frm93.TB41 = vbNullString 'No. keahlian

Frm93.L4_Text = vbNullString 'Kategori Produk
Frm93.L13_Text = 2

Frm93.L15_Text = vbNullString 'No. Resit Trade In
Frm93.L19_Text.BackStyle = 0
Frm93.L20_Text = "0.00" 'Ringkasan Maklumat Bayaran : Deposit Dibayar Secara Tunai
Frm93.L21_Text = "0.00" 'Ringkasan Maklumat Bayaran : Deposit Dari Barang Trade In
'Frm93.L22_Text = "0.00" 'Ringkasan Maklumat Bayaran : Jumlah GST
'Frm93.L23_Text = "0.00" 'Ringkasan Maklumat Bayaran : Jumlah Perlu Bayar (Sebelum Adjustment)
'Frm93.L24_Text = "0.00" 'Ringkasan Maklumat Bayaran : Jumlah Perlu Bayar (Selepas Adjustment)
frm130.L26_Text = "0.00" 'Simpanan Duit Di Kedai
Frm93.L27_Text = 0 '0 : Jenis Tempahan , Status , 1:  No.Siri Produk , 2:  No.Invoice
Frm93.L28_Text = vbNullString 'Memory : Krateria Carian 1
Frm93.L29_Text = vbNullString 'Memory : Krateria Carian 2
Frm93.L33_Text = 0
Frm93.L35_Text = vbNullString 'Nama Pembeli (Tidak Berdaftar)
Frm93.L36_Text = vbNullString 'Nama Pembeli (Berdaftar)

Frm93.L39_Text = "0.00 g" 'Berat belum siap
Frm93.L40_Text = "0.00 g" 'Berat sudah siap

frm130.L31_Text = "0.00"
frm130.L32_Text = "0.00"
frm130.L81_Text = "0.00"
frm130.L82_Text = "0.00"

Frm93.TB1.Locked = False
Frm93.TB2.Locked = False
Frm93.TB3.Locked = False
Frm93.TB4.Locked = True

Frm93.TB1.BackColor = &HFFFFFF
Frm93.TB2.BackColor = &HFFFFFF
Frm93.TB3.BackColor = &HFFFFFF
Frm93.TB4.BackColor = &H8000000A

Frm93.DTPicker1 = DateTime.Date$

Frm93.CB9.Enabled = True

Frm93.CMD12.Visible = True
Frm93.CMD14.Visible = False
Frm93.CMD15.Visible = False
Frm93.CMD19.Enabled = True
'Frm93.CMD21.Enabled = True
            
Frm93.CBB1.Clear
Frm93.CBB2.Clear
Frm93.CBB3.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!kategori_Produk) Then Frm93.CBB1.AddItem rs!kategori_Produk
    If Not IsNull(rs!Kod_Metal_Purity) Then Frm93.CBB2.AddItem rs!Kod_Metal_Purity
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where status='" & "Aktif" & "' AND InvestorSmall = 0 AND InvestorBig = 0", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    If rs!user_level = 1 Or rs!user_level = 2 Or rs!user_level = 3 Or rs!user_level = G_LEVEL_USER Then Frm93.CBB3.AddItem rs!Samaran & "  |  " & rs!NoPekerja
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'frm130.CBB2.Clear

'Set rs = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from 74_cas_kad_kredit where status = 1 order by jenis_kad ASC", cn, adOpenKeyset, adLockOptimistic

'While rs.EOF = False
'    If Not IsNull(rs!jenis_kad) Then frm130.CBB2.AddItem rs!jenis_kad
'    rs.MoveNext
'Wend

'rs.Close
'Set rs = Nothing

If IsNumeric(G_RATE_GST) Then
    Frm93.L19_Text = G_RATE_GST
Else
    Frm93.L19_Text = 6
End If
If G_SCANNER_MODE = 1 Then
    Frm93.CB1 = 1
Else
    Frm93.CB1 = 0
End If
        


GoTo skip_setting:

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Default1 = "Default" Then
        If Not IsNull(rs!gst_value) Then
            Frm93.L19_Text = rs!gst_value 'Jumlah Kadar GST
        Else
            Frm93.L19_Text = 0
        End If
        'If Not IsNull(rs!gst_arinashi) Then
        '    If rs!gst_arinashi = 1 Then
        '        Frm93.CB21 = 0
        '        Frm93.CB22 = 1
        '    Else
        '        Frm93.CB21 = 1
        '        Frm93.CB22 = 0
        '    End If
        'End If
        
        'If Not IsNull(rs!cas_Kad_Kredit) Then Frm93.L31_Text = Format(rs!cas_Kad_Kredit, "0.00") 'Cas Kad Kredit
        'If Not IsNull(rs!cas_debit_kad) Then Frm93.L32_Text = Format(rs!cas_debit_kad, "0.00") 'Cas Debit Kredit

        ''If Not IsNull(rs!no_rujukan_book) Then 'No rujukan tempahan
        ''    Frm93.L17_Text = rs!no_rujukan_book
        ''Else
        ''    Frm93.L17_Text = 1
        ''End If
        ''If Not IsNull(rs!no_rujukan_tak_rasmi) Then
        ''    Frm93.L21_Text = rs!no_rujukan_tak_rasmi 'No. invoice tidak rasmi
        ''Else
        ''    Frm93.L21_Text = 1 'No. invoice tidak rasmi
        ''End If
        ''If Not IsNull(rs!ResitNo) Then 'No. invoice rasmi
        ''    Frm93.L18_Text = rs!ResitNo
        ''Else
        ''    Frm93.L18_Text = 1
        ''End If
        ''If rs!ScannerMode = 1 Then
        ''    Frm93.CB1 = 1
        ''Else
        ''    Frm93.CB1 = 0
        ''End If
    End If
End If

rs.Close
Set rs = Nothing

skip_setting:
'Frm93.CB21.Enabled = True
'Frm93.CB22.Enabled = True
'Frm93.CB25.Enabled = True

Exit Sub

user_level = MDI_frm1.L4_Text

If user_level = "Administration" Or user_level = "Guest/User" Then

    Frm93.CB9.Visible = False
    Frm93.Label24.Visible = False
    Frm93.Label30.Visible = False
    
Else
    
    If G_GST_SYSTEM = "YES" Then
        Frm93.CB9.Visible = True
        Frm93.Label24.Visible = True
        Frm93.Label30.Visible = True
        
        If G_INVOICE_RASMI = 0 Then
            Frm93.CB9 = 1
        Else
            Frm93.CB9 = 0
        End If
        
    Else
        Frm93.CB9.Visible = False
        Frm93.Label24.Visible = False
        Frm93.Label30.Visible = False
    End If
End If
End Sub
Sub frm93_initial_setting2()
'on error resume next
Frm93.Frame2.Left = 120
Frm93.Frame2.Top = 360
Frm93.Frame4.Left = 240
Frm93.Frame4.Top = 1080
Frm93.Frame5.Left = 240
Frm93.Frame5.Top = 1080
Frm93.Pic4.Left = 120
Frm93.Pic4.Top = 360
'Frm93.Pic6.Left = 11280
'Frm93.Pic6.Top = 7200
Frm93.Frame1.Left = 120
Frm93.Frame1.Top = 360
Frm93.Frame8.Left = 120
Frm93.Frame8.Top = 360

Frm93.Frame2.Visible = False
Frm93.Frame4.Visible = False
Frm93.Frame5.Visible = False
Frm93.Pic4.Visible = False
'Frm93.Pic6.Visible = False
Frm93.Frame1.Visible = False
Frm93.Frame8.Visible = False
End Sub
Sub frm93_setting_report()
'on error resume next
Frm93.CBB4.Clear
Frm93.CBB5.Clear
Frm93.CBB7.Clear
Frm93.CBB8.Clear

Frm93.CBB4.AddItem "Tiada filter tarikh"
Frm93.CBB4.AddItem "Mengikut tarikh di bawah"
'Frm93.CBB4.AddItem "Deposit tempahan"
'Frm93.CBB4.AddItem "Tempahan siap"

Frm93.CBB4 = "Tiada filter tarikh"

Frm93.CBB5.AddItem "Semua jenis tempahan"
Frm93.CBB5.AddItem "Tempahan barang baru"
Frm93.CBB5.AddItem "Tempahan barang kedai"

Frm93.CBB5 = "Semua jenis tempahan"

Frm93.CBB7.AddItem "Semua status"
Frm93.CBB7.AddItem "Siap"
Frm93.CBB7.AddItem "Belum Siap"

Frm93.CBB7 = "Semua status"

Frm93.CBB8.AddItem "-"
Frm93.CBB8.AddItem "No. siri produk"
Frm93.CBB8.AddItem "No. invoice deposit"
Frm93.CBB8.AddItem "No. invoice ambilan barang"

Frm93.CBB8 = "-"

Frm93.DTPicker2 = DateTime.Date
Frm93.DTPicker3 = DateTime.Date

Frm93.CBB9.Clear

Frm93.CBB9.AddItem "Semua cawangan"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where status='" & 1 & "' order by cawangan ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!cawangan) Then Frm93.CBB9.AddItem rs!cawangan
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm93.CBB9 = "Semua cawangan"

If MDI_frm1.L4_Text <> "HQ" And MDI_frm1.L4_Text <> "Developer" Then

    Frm93.CBB9 = MDI_frm1.L20_Text
    Frm93.CBB9.Enabled = False
    
Else
    
    Frm93.CBB9.Enabled = True
    
End If
End Sub
Sub Frm93_Call_Product_Detail()
'on error resume next
Frm93_LM_KOD_PURITY = vbNullString

Frm93_LM_No_SIRI = UCase(Frm93.TB5) 'No. Siri Produk

Frm93_LM_KATEGORI = 1
'1:  Pelanggan Biasa
'2 : Member / Ahli
'3:  Pengedar
'4:  RAF
'5:  Normal Dealer(ND)
'6:  Master Dealer(MD)


' ### Periksa kategori pembeli ### - Start
If Frm93.L36_Text <> vbNullString Then
    If Frm28.L5_Text <> vbNullString Then
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
        If Not rs.EOF Then
        
            If Not IsNull(rs!kategori_pelanggan) Then Frm93_LM_KATEGORI = rs!kategori_pelanggan
            
        End If
        
        rs.Close
        Set rs = Nothing
        
    End If
End If
' ### Periksa kategori pembeli ### - End

'###Carian Data Basic Bagi Item Ini### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from Data_Database where no_siri_Produk='" & Frm93_LM_No_SIRI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Not IsNull(rs!cawangan) Then
        
        If MDI_frm1.L20_Text <> rs!cawangan Then
            
            MsgBox "Stok ini adalah milik cawangan [" & rs!cawangan & "]. Anda tidak dibenarkan untuk jual barang ini.", vbExclamation, "Info"
            
            rs.Close
            Set rs = Nothing
            
            Exit Sub
            
        End If
    
    End If
        
    If rs!StatusItem = "10" Then
        If Not IsNull(rs!receiving_Status) Then
        
            If rs!receiving_Status = 2 Or rs!receiving_Status = 3 Then
                
                Frm93.TB5 = vbNullString
                
                MsgBox "Barang/item trade in tidak dibenarkan untuk dijual.", vbExclamation, "Info"
                
                rs.Close
                Set rs = Nothing
                
                Exit Sub
                'GoTo end_jualan:
            
            End If
                
            If rs!receiving_Status = 0 Or rs!receiving_Status = 2 Then
                Frm93.L4_Text = vbNullString
                
                Frm93.TB6 = Frm93_LM_No_SIRI 'No. Siri Produk
                Frm93.TB7 = Format(rs!beza_berat, "0.00") 'Berat Asal (g)
                Frm93.TB8 = Format(rs!beza_berat, "0.00") 'Berat Jualan (g)
                
                Frm93.TB8.Locked = False
                Frm93.TB9.Locked = False
                Frm93.TB10.Locked = False
                Frm93.TB11.Locked = True
                
                Frm93.TB8.BackColor = &HFFFFFF
                Frm93.TB9.BackColor = &HFFFFFF
                Frm93.TB10.BackColor = &HFFFFFF
                Frm93.TB11.BackColor = &H8000000A
                
                If Frm93_LM_KATEGORI = 1 Then
                    If Not IsNull(rs!Upah_Jualan) Then Frm93.TB10 = Format(rs!Upah_Jualan, "0.00") 'Upah Jualan Kepada Pelanggan (RM/g)
                ElseIf Frm93_LM_KATEGORI = 2 Then
                    If Not IsNull(rs!Upah_Member) Then Frm93.TB10 = Format(rs!Upah_Member, "0.00") 'Upah Jualan Kepada Member (RM/g)
                ElseIf Frm93_LM_KATEGORI = 3 Then
                    If Not IsNull(rs!Upah_Pengedar) Then Frm93.TB10 = Format(rs!Upah_Pengedar, "0.00") 'Upah Jualan Kepada Pengedar (RM/g)
                ElseIf Frm93_LM_KATEGORI = 4 Then
                    If Not IsNull(rs!Upah_RAF) Then Frm93.TB10 = Format(rs!Upah_RAF, "0.00") 'Upah Jualan Kepada RAF (RM/g)
                ElseIf Frm93_LM_KATEGORI = 5 Then
                    If Not IsNull(rs!upah_normal_dealer) Then Frm93.TB10 = Format(rs!upah_normal_dealer, "0.00") 'Upah Jualan Kepada Normal Dealer (RM/g)
                ElseIf Frm93_LM_KATEGORI = 6 Then
                    If Not IsNull(rs!upah_master_dealer) Then Frm93.TB10 = Format(rs!upah_master_dealer, "0.00") 'Upah Jualan Kepada Master Dealer (RM/g)
                End If
                
                If Not IsNull(rs!kod_Purity) Then
                    Frm93_LM_KOD_PURITY = rs!kod_Purity 'Kod Purity
                End If
                
                If Not IsNull(rs!kategori_Produk) Then Frm93.L4_Text = rs!kategori_Produk 'Kategori Produk
                Frm93.L13_Text = 0 'Flag Kategori Produk , 0 : BK , 1 : Permata
                Frm93_LM_READY_TO_SAVE = 1 'Flag : Ready To Save
            Else
                Frm93.TB6 = Frm93_LM_No_SIRI 'No. Siri Produk
                
                Frm93.L4_Text = vbNullString
                Frm93.TB7 = vbNullString
                Frm93.TB8 = vbNullString
                Frm93.TB9 = vbNullString
                'Frm93.TB10 = vbNullString
                
                Frm93.TB7.Locked = True
                Frm93.TB8.Locked = True
                Frm93.TB9.Locked = True
                'Frm93.TB10.Locked = True
                Frm93.TB11.Locked = False
                
                Frm93.TB7.BackColor = &H8000000A
                Frm93.TB8.BackColor = &H8000000A
                Frm93.TB9.BackColor = &H8000000A
                'Frm93.TB10.BackColor = &H8000000A
                Frm93.TB11.BackColor = &HFFFFFF
                
                If Frm93_LM_KATEGORI = 1 Then
                    If Not IsNull(rs!code_Supplier) Then Frm93.TB11 = Format(rs!code_Supplier, "0.00") 'Harga Jualan Kepada Pelanggan (RM)
                ElseIf Frm93_LM_KATEGORI = 2 Then
                    If Not IsNull(rs!HargaJualan_Member) Then Frm93.TB11 = Format(rs!HargaJualan_Member, "0.00") 'Harga Jualan Kepada Member (RM)
                ElseIf Frm93_LM_KATEGORI = 4 Then
                    If Not IsNull(rs!HargaJualan_Pengedar) Then Frm93.TB11 = Format(rs!HargaJualan_Pengedar, "0.00") 'Harga Jualan Kepada Pengedar (RM)
                ElseIf Frm93_LM_KATEGORI = 3 Then
                    If Not IsNull(rs!HargaJualan_RAF) Then Frm93.TB11 = Format(rs!HargaJualan_RAF, "0.00") 'Harga Jualan Kepada RAF (RM)
                ElseIf Frm93_LM_KATEGORI = 5 Then
                    If Not IsNull(rs!hargajualan_normal_dealer) Then Frm93.TB11 = Format(rs!hargajualan_normal_dealer, "0.00") 'Harga Jualan Kepada Normal Dealer (RM)
                ElseIf Frm93_LM_KATEGORI = 6 Then
                    If Not IsNull(rs!hargajualan_master_dealer) Then Frm93.TB11 = Format(rs!hargajualan_master_dealer, "0.00") 'Harga Jualan Kepada Master Dealer (RM)
                End If
                
                Frm93.L13_Text = 1 'Flag Kategori Produk , 0 : BK , 1 : Permata
                
                Frm93_LM_PERMATA = 1
            End If
        End If
        
        Frm93.TB12 = "0.00"
        If Not IsNull(rs!kategori_Produk) Then Frm93.L4_Text = rs!kategori_Produk 'Kategori Produk
        
        If Frm93_LM_PERMATA = 1 Then
            Frm93.TB7 = Format(Frm93_LM_HARGA_JUALAN, "0.00")
        End If
    ElseIf rs!StatusItem = "11" Then
        MsgBox "Item Ini Telah Dijual. No. Siri Produk [" & Frm93_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm93.TB5 = vbNullString
        Frm93.TB5.SetFocus
    ElseIf rs!StatusItem = "12" Then
        MsgBox "Item Ini Telah Dijual Secara Potong. No. Siri Produk [" & Frm93_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm93.TB5 = vbNullString
        Frm93.TB5.SetFocus
    ElseIf rs!StatusItem = "13" Then
        MsgBox "Item Ini Telah Dijual Secara Potong. No. Siri Produk [" & Frm93_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm93.TB5 = vbNullString
        Frm93.TB5.SetFocus
    ElseIf rs!StatusItem = "14" Or rs!StatusItem = "21" Or rs!StatusItem = "22" Then
        MsgBox "Item Ini Telah Ditempah Oleh Pelanggan. No. Siri Produk [" & Frm93_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm93.TB5 = vbNullString
        Frm93.TB5.SetFocus
    ElseIf rs!StatusItem = "15" Or rs!StatusItem = "19" Or rs!StatusItem = "20" Then
        MsgBox "Item Ini Telah Dibeli Secara Ansuran. No. Siri Produk [" & Frm93_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm93.TB5 = vbNullString
        Frm93.TB5.SetFocus
    ElseIf rs!StatusItem = "16" Then
        MsgBox "Item Ini Telah Dihantar Ke Ar-Rahnu. No. Siri Produk [" & Frm93_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm93.TB5 = vbNullString
        Frm93.TB5.SetFocus
    ElseIf rs!StatusItem = "17" Then
        MsgBox "Item Ini Telah Dijual Secara ETA. No. Siri Produk [" & Frm93_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm93.TB5 = vbNullString
        Frm93.TB5.SetFocus
    ElseIf rs!StatusItem = "23" Then
        MsgBox "Item Ini Telah Dihantar Ke Supplier/Kilang. No. Siri Produk [" & Frm93_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm93.TB5 = vbNullString
        Frm93.TB5.SetFocus
    ElseIf rs!StatusItem = "24" Then
        MsgBox "Item Ini Telah Dihantar Ke Supplier/Kilang. No. Siri Produk [" & Frm93_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm93.TB5 = vbNullString
        Frm93.TB5.SetFocus
    ElseIf rs!StatusItem = "25" Then
        MsgBox "Item Ini Telah Diagihkan Ke Cawangan. No. Siri Produk [" & Frm93_LM_No_SIRI & "]", vbExclamation, "Info"

        Frm93.TB5 = vbNullString
        Frm93.TB5.SetFocus
    ElseIf rs!StatusItem = "26" Then
        MsgBox "Item Ini Telah Dijual Oleh Cawangan. No. Siri Produk [" & Frm93_LM_No_SIRI & "]", vbExclamation, "Info"

        Frm93.TB5 = vbNullString
        Frm93.TB5.SetFocus
    ElseIf rs!StatusItem = "0" Then
        MsgBox "Item Ini Telah Dipadamkan Dari Sistem. No. Siri Produk [" & Frm93_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm93.TB5 = vbNullString
        Frm93.TB5.SetFocus
    End If
Else
    Frm93.TB5 = vbNullString
    
    MsgBox "No. Siri Produk Ini [" & Frm93_LM_No_SIRI & "] Tidak Dijumpai.", vbExclamation, "Info"
    
    Frm93.TB5.SetFocus
End If

rs.Close
Set rs = Nothing

'###Carian Data Basic Bagi Item Ini### - End


'###Periksa Data Produk### - Start
If Frm93_LM_READY_TO_SAVE = 1 Then 'Flag : Ready To Save
    If Frm93_LM_KOD_PURITY <> vbNullString Then
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from hargaemas where Purity='" & Frm93_LM_KOD_PURITY & "' AND cawangan='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Frm93_LM_KATEGORI = 1 Then
                If IsNumeric(rs!Harga_Pelanggan) Then Frm93.TB9 = Format(rs!Harga_Pelanggan, "0.00") 'Harga Semasa Bagi Pelanggan (RM/g)
            ElseIf Frm93_LM_KATEGORI = 2 Then
                If IsNumeric(rs!Harga_Member) Then Frm93.TB9 = Format(rs!Harga_Member, "0.00") 'Harga Semasa Bagi Member (RM/g)
            ElseIf Frm93_LM_KATEGORI = 3 Then
                If IsNumeric(rs!Harga_Pengedar) Then Frm93.TB9 = Format(rs!Harga_Pengedar, "0.00") 'Harga Semasa Bagi Pengedar (RM/g)
            ElseIf Frm93_LM_KATEGORI = 4 Then
                If IsNumeric(rs!Harga_RAF) Then Frm93.TB9 = Format(rs!Harga_RAF, "0.00") 'Harga Semasa Bagi RAF (RM/g)
            ElseIf Frm93_LM_KATEGORI = 5 Then
                If IsNumeric(rs!harga_nd) Then Frm93.TB9 = Format(rs!harga_nd, "0.00") 'Harga Semasa Bagi Normal Dealer (RM/g)
            ElseIf Frm93_LM_KATEGORI = 6 Then
                If IsNumeric(rs!harga_md) Then Frm93.TB9 = Format(rs!harga_md, "0.00") 'Harga Semasa Bagi Master Dealer (RM/g)
            End If
        End If
        
        rs.Close
        Set rs = Nothing
    End If
End If

Frm93.TB5 = vbNullString
'###Periksa Data Produk### - End
End Sub
Sub frm93_tempahan_header()
'on error resume next
With Frm93.LV1
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear

    Frm93.LV1.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "Tarikh Tempahan", 2000, 2
    .ColumnHeaders.Add 5, , "Jenis Tempahan", 2000, 2
    .ColumnHeaders.Add 6, , "Status", 1400, 2
    .ColumnHeaders.Add 7, , "No. Siri Produk", 2000
    .ColumnHeaders.Add 8, , "Kategori Produk", 4500
    .ColumnHeaders.Add 9, , "Purity", 0
    .ColumnHeaders.Add 10, , "Berat / Anggaran Berat (g)", 2500, 1
    .ColumnHeaders.Add 11, , "Harga Semasa (RM/g)", 2300, 1
    .ColumnHeaders.Add 12, , "Upah (RM)", 1200, 1
    .ColumnHeaders.Add 13, , "Harga / Anggaran Harga (RM)", 2800, 1
    .ColumnHeaders.Add 14, , "Nama", 5000
    .ColumnHeaders.Add 15, , "No. Kad Pengenalan", 2500
    .ColumnHeaders.Add 16, , "No. Telefon", 2500
    .ColumnHeaders.Add 17, , "No. Rujukan", 0
    .ColumnHeaders.Add 18, , "Cawangan", 3500

End With
End Sub
Sub frm93_tempahan()
'on error resume next
Dim frm93_field_3 As String
Dim Frm93_LM_TOTAL_PAGE As Double
Dim TM As Date
Dim TA As Date
Dim Frm93_LM_CURR_PAGE As Double

Frm93_LM_CURR_PAGE = 0
Frm93_LM_TOTAL_PAGE = 0

x = 0
Y = 0
Frm93_PAGE_SIZE = 37

Frm93.L33_Text = 0

re_gen_report:

Frm93.L39_Text = "0.00 g" 'Berat belum siap
Frm93.L40_Text = "0.00 g" 'Berat sudah siap

If Frm93.L27_Text = 1 Then '0 : Tiada carian mengikut tarikh , 1 : Carian mengikut tarikh
    TM = Frm93.L28_Text 'Tarikh mula
    TA = Frm93.L29_Text 'Tarikh akhir
End If

If Frm93.L30_Text = "Semua jenis tempahan" Then 'Jenis tempahan

    Frm93_LM_SEARCH_1 = Null
    Frm93_LM_SEARCH_1_LOGIC = "<>"
    
Else

    If Frm93.L30_Text = "Tempahan barang baru" Then
        Frm93_LM_SEARCH_1 = 0 'Jenis Tempahan , 0 : Tempahan Barang Baru , 1 : Tempahan Barang Kedai
    ElseIf Frm93.L30_Text = "Tempahan barang kedai" Then
        Frm93_LM_SEARCH_1 = 1 'Jenis Tempahan , 0 : Tempahan Barang Baru , 1 : Tempahan Barang Kedai
    End If

    Frm93_LM_SEARCH_1_LOGIC = "="
    
End If

If Frm93.L31_Text = "Semua status" Then 'Status

    Frm93_LM_SEARCH_2 = Null
    Frm93_LM_SEARCH_2_LOGIC = "<>"
    
Else

    If Frm93.L31_Text = "Siap" Then
        Frm93_LM_SEARCH_2 = "Siap"
    ElseIf Frm93.L31_Text = "Belum Siap" Then
        Frm93_LM_SEARCH_2 = "Belum Siap"
    End If

    Frm93_LM_SEARCH_2_LOGIC = "="
    
End If

If Frm93.L32_Text = "-" Then 'Lain-lain

    Frm93_LM_SEARCH_3 = Null
    Frm93_LM_SEARCH_3_LOGIC = "<>"
    frm93_field_3 = "no_resit_tempahan"
    
Else

    If Frm93.L32_Text = "No. siri produk" Then
    
        frm93_field_3 = "no_siri_produk"
        
    ElseIf Frm93.L32_Text = "No. invoice deposit" Then
    
        frm93_field_3 = "no_resit_tempahan"
        
    ElseIf Frm93.L32_Text = "No. invoice ambilan barang" Then
    
        frm93_field_3 = "invoice_siap"
        
    End If
    
    Frm93_LM_SEARCH_3 = Frm93.L38_Text
    Frm93_LM_SEARCH_3_LOGIC = "="
    
End If

If Frm93.L45_Text = "Semua cawangan" Then 'Cawangan

    Frm93_LM_SEARCH_4 = Null
    Frm93_LM_SEARCH_4_LOGIC = "<>"
    
Else

    Frm93_LM_SEARCH_4 = Frm93.L45_Text
    Frm93_LM_SEARCH_4_LOGIC = "="
    
End If

If Frm93.L27_Text = 0 Then Frm93.L25_Text = "Senarai tempahan." 'Header
If Frm93.L27_Text = 1 Then Frm93.L25_Text = "Senarai tempahan dari " & TM & " hingga " & TA & "." 'Header

LM_START_ROW = Frm93.L62_Text

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm93_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm93.L63_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm93_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        Frm93.L60_Text = 1
    End If
End If

Frm93_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm93.L27_Text = 0 Then rs.Open "select * from 40_tempahan_deposit where status_invoice = 1 AND cawangan " & Frm93_LM_SEARCH_4_LOGIC & "'" & Frm93_LM_SEARCH_4 & "' AND jenis_tempahan " & Frm93_LM_SEARCH_1_LOGIC & "'" & Frm93_LM_SEARCH_1 & "' AND status " & Frm93_LM_SEARCH_2_LOGIC & "'" & Frm93_LM_SEARCH_2 & "' AND " & frm93_field_3 & Frm93_LM_SEARCH_3_LOGIC & "'" & Frm93_LM_SEARCH_3 & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm93_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If Frm93.L27_Text = 1 Then rs.Open "select * from 40_tempahan_deposit where status_invoice = 1 AND cawangan " & Frm93_LM_SEARCH_4_LOGIC & "'" & Frm93_LM_SEARCH_4 & "' AND jenis_tempahan " & Frm93_LM_SEARCH_1_LOGIC & "'" & Frm93_LM_SEARCH_1 & "' AND status " & Frm93_LM_SEARCH_2_LOGIC & "'" & Frm93_LM_SEARCH_2 & "' AND " & frm93_field_3 & Frm93_LM_SEARCH_3_LOGIC & "'" & Frm93_LM_SEARCH_3 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm93_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm93_LM_PAGE_FOUND = 0 Then
        If Frm93.L63_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm93.L60_Text = Frm93.L60_Text + 1
                Frm93_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm93.L60_Text) Then
                    If Frm93.L60_Text <> 1 Then
                        Frm93.L60_Text = Frm93.L60_Text - 1
                        Frm93_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm93.L60_Text - 1) * Frm93_PAGE_SIZE) + x
        
    With Frm93.LV1.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!tarikh) Then 'Tarikh
            .ListSubItems.Add , , rs!tarikh
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!jenis_tempahan) Then
            If rs!jenis_tempahan = 0 Then
                .ListSubItems.Add , , "Barang Baru" 'Jenis Tempahan
            ElseIf rs!jenis_tempahan = 1 Then
                .ListSubItems.Add , , "Barang Kedai" 'Jenis Tempahan
            End If
        End If
        
        If Not IsNull(rs!Status) Then 'Status
            .ListSubItems.Add , , rs!Status
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_siri_Produk) Then 'No. Siri Produk
            .ListSubItems.Add , , rs!no_siri_Produk
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kategori_Produk) Then 'Kategori Produk
            .ListSubItems.Add , , rs!kategori_Produk
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!purity) Then 'Purity
            .ListSubItems.Add , , rs!purity
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!jenis_tempahan) Then
        
            If rs!jenis_tempahan = 0 Then

                If Not IsNull(rs!anggaran_berat) Then 'Anggaran Berat
                    .ListSubItems.Add , , Format(rs!anggaran_berat, "#,##0.00")
                Else
                    .ListSubItems.Add , , Format(0, "#,##0.00")
                End If
        
            ElseIf rs!jenis_tempahan = 1 Then
            
                If Not IsNull(rs!berat_jualan) Then 'Anggaran Berat
                    .ListSubItems.Add , , Format(rs!berat_jualan, "#,##0.00")
                Else
                    .ListSubItems.Add , , Format(0, "#,##0.00")
                End If
                
            End If
            
        End If

        If Not IsNull(rs!harga_Semasa) Then 'Harga Semasa
            .ListSubItems.Add , , Format(rs!harga_Semasa, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        
        If Not IsNull(rs!UPAH) Then 'Upah
            .ListSubItems.Add , , Format(rs!UPAH, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        
        If Not IsNull(rs!anggaran_harga) Then 'Anggaran Harga
            .ListSubItems.Add , , Format(rs!anggaran_harga, "#,##0.00")
        Else
            .ListSubItems.Add , , Format(0, "#,##0.00")
        End If
        
        If Not IsNull(rs!Nama) Then 'Nama
            .ListSubItems.Add , , rs!Nama
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!no_ic) Then 'No. Kad Pengenalan
            .ListSubItems.Add , , rs!no_ic
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_tel) Then 'No. Telefon
            .ListSubItems.Add , , rs!no_tel
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_rujukan_tempahan) Then 'No. Rujukan
            .ListSubItems.Add , , rs!no_rujukan_tempahan
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!cawangan) Then 'Cawangan
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
If Frm93.L27_Text = 0 Then rs.Open "select COUNT(ID) from 40_tempahan_deposit where status_invoice = 1 AND cawangan " & Frm93_LM_SEARCH_4_LOGIC & "'" & Frm93_LM_SEARCH_4 & "' AND jenis_tempahan " & Frm93_LM_SEARCH_1_LOGIC & "'" & Frm93_LM_SEARCH_1 & "' AND status " & Frm93_LM_SEARCH_2_LOGIC & "'" & Frm93_LM_SEARCH_2 & "' AND " & frm93_field_3 & Frm93_LM_SEARCH_3_LOGIC & "'" & Frm93_LM_SEARCH_3 & "'", cn, adOpenKeyset, adLockOptimistic
If Frm93.L27_Text = 1 Then rs.Open "select COUNT(ID) from 40_tempahan_deposit where status_invoice = 1 AND cawangan " & Frm93_LM_SEARCH_4_LOGIC & "'" & Frm93_LM_SEARCH_4 & "' AND jenis_tempahan " & Frm93_LM_SEARCH_1_LOGIC & "'" & Frm93_LM_SEARCH_1 & "' AND status " & Frm93_LM_SEARCH_2_LOGIC & "'" & Frm93_LM_SEARCH_2 & "' AND " & frm93_field_3 & Frm93_LM_SEARCH_3_LOGIC & "'" & Frm93_LM_SEARCH_3 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs(0)) Then
        Frm93_LM_TOTAL_PAGE = Format(rs(0) / Frm93_PAGE_SIZE, "0.00") 'Jumlah Page
        
        'Periksa Samada ada titik perpuluhan atau tidak
        If InStr(1, Frm93_LM_TOTAL_PAGE, ".") <> 0 Then
        
            Frm85_LM_PAGE = Split(Frm93_LM_TOTAL_PAGE, ".")(0)
            Frm85_LM_PAGE_LEBIHAN = Split(Frm93_LM_TOTAL_PAGE, ".")(1)
            
            If Frm85_LM_PAGE_LEBIHAN <> "00" Then
                Frm93.L61_Text = Frm85_LM_PAGE + 1 'Total Page
            Else
                Frm93.L61_Text = Frm85_LM_PAGE
            End If
            
        Else
        
            Frm93.L61_Text = Frm93_LM_TOTAL_PAGE
            
        End If
    
        If rs(0) = vbNullString Then
            Frm93.L61_Text = 0
        End If
    End If
Else
    Frm93.L61_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm93.L61_Text = vbNullString Then
    Frm93.L61_Text = 0
End If
'### Jumlah Data ### - End

'#### Jumlah Data #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm93.L27_Text = 0 Then rs.Open "select COUNT(ID) from 40_tempahan_deposit where status_invoice = 1 AND cawangan " & Frm93_LM_SEARCH_4_LOGIC & "'" & Frm93_LM_SEARCH_4 & "' AND jenis_tempahan " & Frm93_LM_SEARCH_1_LOGIC & "'" & Frm93_LM_SEARCH_1 & "' AND status " & Frm93_LM_SEARCH_2_LOGIC & "'" & Frm93_LM_SEARCH_2 & "' AND " & frm93_field_3 & Frm93_LM_SEARCH_3_LOGIC & "'" & Frm93_LM_SEARCH_3 & "'", cn, adOpenKeyset, adLockOptimistic
If Frm93.L27_Text = 1 Then rs.Open "select COUNT(ID) from 40_tempahan_deposit where status_invoice = 1 AND cawangan " & Frm93_LM_SEARCH_4_LOGIC & "'" & Frm93_LM_SEARCH_4 & "' AND jenis_tempahan " & Frm93_LM_SEARCH_1_LOGIC & "'" & Frm93_LM_SEARCH_1 & "' AND status " & Frm93_LM_SEARCH_2_LOGIC & "'" & Frm93_LM_SEARCH_2 & "' AND " & frm93_field_3 & Frm93_LM_SEARCH_3_LOGIC & "'" & Frm93_LM_SEARCH_3 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm93.L33_Text = rs(0)

rs.Close
Set rs = Nothing
'#### Jumlah Data#### - End

If x <> 0 Then
    Frm93.L62_Text = LM_START_ROW
End If

Dim LM_BERAT_BELUM_SIAP1 As Double
Dim LM_BERAT_BELUM_SIAP2 As Double
Dim LM_BERAT_SUDAH_SIAP1 As Double
Dim LM_BERAT_SUDAH_SIAP2 As Double

'#### Jumlah berat belum siap #### - Start
Set rs = New ADODB.Recordset
Call Main
If Frm93.L27_Text = 0 Then rs.Open "select SUM(berat_jualan) from 40_tempahan_deposit where status_invoice = 1 AND cawangan " & Frm93_LM_SEARCH_4_LOGIC & "'" & Frm93_LM_SEARCH_4 & "' AND jenis_tempahan " & Frm93_LM_SEARCH_1_LOGIC & "'" & Frm93_LM_SEARCH_1 & "' AND status='" & "Belum Siap" & "' AND " & frm93_field_3 & Frm93_LM_SEARCH_3_LOGIC & "'" & Frm93_LM_SEARCH_3 & "'", cn, adOpenKeyset, adLockOptimistic
If Frm93.L27_Text = 1 Then rs.Open "select SUM(berat_jualan) from 40_tempahan_deposit where status_invoice = 1 AND cawangan " & Frm93_LM_SEARCH_4_LOGIC & "'" & Frm93_LM_SEARCH_4 & "' AND jenis_tempahan " & Frm93_LM_SEARCH_1_LOGIC & "'" & Frm93_LM_SEARCH_1 & "' AND status='" & "Belum Siap" & "' AND " & frm93_field_3 & Frm93_LM_SEARCH_3_LOGIC & "'" & Frm93_LM_SEARCH_3 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm93.L39_Text = Format(rs(0), "#,##0.00 g") 'Berat belum siap

rs.Close
Set rs = Nothing
'#### Jumlah berat belum siap #### - End

'#### Jumlah berat siap #### - Start
Set rs = New ADODB.Recordset
Call Main
If Frm93.L27_Text = 0 Then rs.Open "select SUM(berat_jualan) from 40_tempahan_deposit where status_invoice = 1 AND cawangan " & Frm93_LM_SEARCH_4_LOGIC & "'" & Frm93_LM_SEARCH_4 & "' AND jenis_tempahan " & Frm93_LM_SEARCH_1_LOGIC & "'" & Frm93_LM_SEARCH_1 & "' AND status='" & "Siap" & "' AND " & frm93_field_3 & Frm93_LM_SEARCH_3_LOGIC & "'" & Frm93_LM_SEARCH_3 & "'", cn, adOpenKeyset, adLockOptimistic
If Frm93.L27_Text = 1 Then rs.Open "select SUM(berat_jualan) from 40_tempahan_deposit where status_invoice = 1 AND cawangan " & Frm93_LM_SEARCH_4_LOGIC & "'" & Frm93_LM_SEARCH_4 & "' AND jenis_tempahan " & Frm93_LM_SEARCH_1_LOGIC & "'" & Frm93_LM_SEARCH_1 & "' AND status='" & "Siap" & "' AND " & frm93_field_3 & Frm93_LM_SEARCH_3_LOGIC & "'" & Frm93_LM_SEARCH_3 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm93.L40_Text = Format(rs(0), "#,##0.00 g") 'Berat belum siap

rs.Close
Set rs = Nothing
'#### Jumlah berat siap #### - End

'Frm93.L39_Text = "0.00 g" 'Berat belum siap
'Frm93.L40_Text = "0.00 g" 'Berat sudah siap

If Frm93.L60_Text <> vbNullString And IsNumeric(Frm93.L60_Text) Then
    If Frm93.L61_Text <> vbNullString And IsNumeric(Frm93.L61_Text) Then
        Frm93_LM_CURR_PAGE = Frm93.L60_Text
        Frm93_LM_TOTAL_PAGE = Frm93.L61_Text
        
        If Frm93_LM_CURR_PAGE > Frm93_LM_TOTAL_PAGE Then
            
            Frm93.L60_Text = Frm93.L60_Text - 1
            
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            GoTo re_gen_report:
            
        End If
    End If
End If

Frm93.Frame1.Visible = True
Frm93.Frame8.Visible = False
End Sub
Sub Frm93_padam_data_deposit()
'Ini digunakan bagi padam data deposit sebelum edit data tempahan
'on error resume next
Dim Frm93_LM_SIMPANAN_DIGUNAKAN As Double
Dim Frm93_LM_SIMPANAN_ASAL As Double

Frm93_LM_SIMPANAN_DIGUNAKAN = 0
Frm93_LM_SIMPANAN_ASAL = 0
Frm93_LM_FLAG_TI = 0
Frm93_LM_FLAG_BARANG_KEDAI = 0 'Flag Barang Kedai

G_JENIS_URUSAN = 9

'$$$ No. staff $$$ - Start
If InStr(1, Frm93.CBB1, "  |  ") <> 0 Then

    Frm93_LM_EMP_NO = Split(Frm93.CBB3, "  |  ")(1)
    
Else

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        
        If Not IsNull(rs!NoPekerja) Then Frm93_LM_EMP_NO = rs!NoPekerja

    End If
    
    rs.Close
    Set rs = Nothing

End If

GoTo skip_carian_user:

If MDI_frm1.L3_Text <> vbNullString Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        If Not IsNull(rs!NoPekerja) Then G_LOGIN_USER = rs!NoPekerja

    End If
    
    rs.Close
    Set rs = Nothing
    
End If
'$$$ No. staff $$$ - End

skip_carian_user:
  
LM_NOW = Now

'### Padam Data Dari Senarai Tempahan (Deposit) ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 40_tempahan_deposit where no_rujukan_tempahan='" & Frm93.L17_Text & "' AND status_invoice = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    
    G_ID = rs!ID
    Call recovery_40_tempahan_deposit
    
    If Not IsNull(rs!no_resit_tempahan) Then Frm93_LM_No_RESIT = rs!no_resit_tempahan 'No. Resit
    If Not IsNull(rs!flag_trade_in) Then
        If rs!flag_trade_in = 1 Then
            Frm93_LM_FLAG_TI = 1
            If Not IsNull(rs!no_resit_trade_in) Then Frm93_LM_No_RESIT_TI = rs!no_resit_trade_in 'No. Resit Trade In
        End If
    End If
    
    If Not IsNull(rs!jenis_tempahan) Then
        If rs!jenis_tempahan = 1 Then
            Frm93_LM_FLAG_BARANG_KEDAI = 1 'Flag Barang Kedai
            If Not IsNull(rs!no_siri_Produk) Then Frm93_LM_No_SIRI = rs!no_siri_Produk 'No. Siri Produk
        End If
    End If
    
    If Not IsNull(rs!no_rujukan_pelanggan) Then Frm93_LM_No_RUJUKAN_PEMBELI = rs!no_rujukan_pelanggan 'No. Rujukan Pembeli
    
    rs!status_invoice = 0
    
    rs!terminal = G_TERMINAL
    rs!write_timestamp2 = LM_NOW
    rs!no_staff = G_LOGIN_USER 'No. staff yang memasukkan/edit data (Login)
    rs.Update
    
End If

rs.Close
Set rs = Nothing
'### Padam Data Dari Senarai Tempahan (Deposit) ### - End

'### Pulangkan Status Barang Trade In ### - Start
If Frm93_LM_FLAG_TI = 1 Then '0 : Tiada Urusan Trade In , 1 : Ada Urusan Trade In
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & Frm93_LM_No_RESIT_TI & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        
        G_ID = rs!ID
        Call recovery_16_gold_bar_belian
        
        rs!trade_in_status = 0
        rs!no_staff = Frm93_LM_EMP_NO 'No. Pekerja
        rs!terminal = G_TERMINAL
        rs!write_timestamp2 = LM_NOW
        rs!jenis_urusan = G_JENIS_URUSAN
        rs!remarks = "Kembalikan status trade in - edit data deposit tempahan"
        rs.Update
                
    End If
    
    rs.Close
    Set rs = Nothing
End If
'### Pulangkan Status Barang Trade In ### - End

'### Pulangkan Status Item Dalam Database ### - Start
If Frm93_LM_FLAG_BARANG_KEDAI = 1 Then 'Flag Barang Kedai
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from Data_Database where no_siri_produk='" & Frm93_LM_No_SIRI & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        G_ID = rs!ID
        Call recovery_data_database
        
        rs!StatusItem = 10
        rs!write_timestamp2 = LM_NOW
        rs!no_pekerja = Frm93_LM_EMP_NO
        rs!terminal = G_TERMINAL
        rs!Menu = 3
                
        rs.Update
                
    End If
    
    rs.Close
    Set rs = Nothing
End If
'### Pulangkan Status Item Dalam Database ### - End

'###Padam Akaun Tempahan### - Start
Frm93_LM_FLAG_SAVING = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where no_resit='" & Frm93_LM_No_RESIT & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    
    G_ID = rs!ID
    Call recovery_22_jualan
    
    If Not IsNull(rs!duit_simpanan_kedai) Then
        If Format(rs!duit_simpanan_kedai, "0.00") <> "0.00" Then
            If IsNumeric(rs!duit_simpanan_kedai) Then Frm93_LM_SIMPANAN_DIGUNAKAN = rs!duit_simpanan_kedai
            Frm93_LM_FLAG_SAVING = 1
        End If
    End If

    rs!Status = 0
    rs!terminal = G_TERMINAL
    rs!no_staff = G_LOGIN_USER
    rs!write_timestamp2 = LM_NOW
    rs!Menu = 3
    rs.Update
    
End If

rs.Close
Set rs = Nothing
            
'###Padam Akaun Tempahan### - End

'### Padam Rekod Belian Pembeli Yang Tidak Berdaftar Dengan Kedai ### - Start (08-07-2015)
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 44_senarai_pelanggan where no_resit='" & Frm93_LM_No_RESIT & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    G_ID = rs!ID
    Call recovery_44_senarai_pelanggan
    
    rs.Delete
    rs.Update

End If

rs.Close
Set rs = Nothing
'### Padam Rekod Belian Pembeli Yang Tidak Berdaftar Dengan Kedai ### - End (08-07-2015)

'###Update Simpanan Duit Di Kedai### - Start
If Frm93_LM_FLAG_SAVING = 1 Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm93_LM_No_RUJUKAN_PEMBELI & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        G_ID = rs!ID
        Call recovery_senarai_pelanggan
                
        If Not IsNull(rs!baki_simpanan) Then
            If IsNumeric(rs!baki_simpanan) Then Frm93_LM_SIMPANAN_ASAL = rs!baki_simpanan 'Baki Simpanan Pelanggan Ini (RM)
        End If
        
        rs!baki_simpanan = Format(Frm93_LM_SIMPANAN_ASAL + Frm93_LM_SIMPANAN_DIGUNAKAN, "0.00") 'Baki Simpanan Terkini Pelanggan Ini (RM)
        
        rs!write_timestamp2 = LM_NOW
        rs!no_staff = Frm93_LM_EMP_NO 'No. Pekerja
        rs!terminal = G_TERMINAL
        rs!jenis_urusan = G_JENIS_URUSAN
                
        rs.Update
    End If
    
    rs.Close
    Set rs = Nothing

'###Padam Rekod Bayaran Dalam Table Simpanan### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 24_rekod_kewangan_pelanggan where NO_RESIT='" & Frm93_LM_No_RESIT & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then

        G_ID = rs!ID
        Call recovery_24_rekod_kewangan_pelanggan
                
        rs.Delete
        rs.Update
    End If
    
    rs.Close
    Set rs = Nothing
'###Padam Rekod Bayaran Dalam Table Simpanan### - End
        
End If
'###Update Simpanan Duit Di Kedai### - End
End Sub
Sub Frm94_initial_setting()
'on error resume next
Frm94.L7_Text = 0
Frm94.TB1 = vbNullString 'No. Siri Produk
Frm94.TB2 = vbNullString 'No. Siri Produk
Frm94.TB3 = "0.00" 'Berat Asal
Frm94.TB4 = "0.00" 'Berat Jualan
Frm94.TB5 = "0.00" 'Harga Semasa
Frm94.TB6 = "0.00" 'Upah
Frm94.TB7 = "0.00" 'Harga Asal
Frm94.TB8 = "0.00" 'Adjustment
Frm94.TB9 = "0.00" 'Harga Jualan
Frm94.TB10 = "0.00" 'Jumlah Keseluruhan
Frm94.TB11 = "0.00" 'Bayaran Sudah Jelas
Frm94.TB12 = "0.00" 'Baki
Frm94.TB13 = vbNullString 'Carian No. Voucher Trade In
Frm94.TB14 = "0.00" 'Nilaian Trade In
Frm94.TB17 = "0.00" 'Jumlah Cukai GST
Frm94.TB22 = "0.00" 'Bayaran Dari Barang Trade In
frm130.TB21 = "0.00" 'Cara Bayaran : Duit Simpanan Di Kedai
frm130.TB27 = "0.00" 'Cara Bayaran : Tunai
frm130.TB28 = "0.00" 'Cara Bayaran : Bank In
frm130.TB29 = "0.00" 'Cara Bayaran : Kad Kredit
frm130.TB32 = "0.00" 'Cara Bayaran : Jumlah Bayaran

Frm94.L3_Text = vbNullString
Frm94.L4_Text = vbNullString
Frm94.L6_Text = "Baki (RM)                                 :"

Frm94.L8_Text.BackStyle = 0
Frm94.L9_Text = 1
Frm94.L10_Text = 1
Frm94.L13_Text.Visible = False
Frm94.L14_Text = 1
Frm94.L15_Text = 1
frm130.L26_Text = "0.00"
Frm94.L32_Text = 0

frm130.L31_Text = "0.00"
frm130.L32_Text = "0.00"
frm130.L81_Text = "0.00"
frm130.L82_Text = "0.00"

Frm94.DTPicker1 = DateTime.Date

Frm94.TB1.Locked = False
Frm94.TB1.BackColor = &HFFFFFF
Frm94.CMD1.Enabled = True

Frm94.CBB1.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where status='" & "Aktif" & "' AND InvestorSmall = 0 AND InvestorBig = 0", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    If rs!user_level = 1 Or rs!user_level = 2 Or rs!user_level = 3 Or rs!user_level = G_LEVEL_USER Then Frm94.CBB1.AddItem rs!Samaran & "  |  " & rs!NoPekerja
    'If rs!Status = "Aktif" And rs!InvestorSmall = 0 And rs!InvestorBig = 0 Then
    '    Frm94.CBB1.AddItem rs!Samaran & "  |  " & rs!NoPekerja
    'End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'frm130.CBB2.Clear

'Set rs = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from 74_cas_kad_kredit where status = 1 order by jenis_kad ASC", cn, adOpenKeyset, adLockOptimistic

'While rs.EOF = False
'    If Not IsNull(rs!jenis_kad) Then frm130.CBB2.AddItem rs!jenis_kad
'    rs.MoveNext
'Wend

'rs.Close
'Set rs = Nothing

If IsNumeric(G_RATE_GST) Then
    Frm94.L8_Text = G_RATE_GST 'Jumlah Kadar GST
    frm130.L8_Text = G_RATE_GST
Else
    Frm94.L8_Text = 6
    frm130.L8_Text = 6
End If
If G_GST_JUAL = 1 Then
    Frm94.CB3 = 0
    Frm94.CB4 = 1
Else
    Frm94.CB3 = 1
    Frm94.CB4 = 0
End If
If G_SCANNER_MODE = 1 Then
    Frm94.CB1 = 1
Else
    Frm94.CB1 = 0
End If

GoTo skip_setting:

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Default1 = "Default" Then
        If Not IsNull(rs!gst_value) Then
            Frm94.L8_Text = rs!gst_value 'Jumlah Kadar GST
        Else
            Frm94.L8_Text = 0
        End If
        If Not IsNull(rs!gst_arinashi) Then
            If rs!gst_arinashi = 1 Then
                Frm94.CB3 = 0
                Frm94.CB4 = 1
            Else
                Frm94.CB3 = 1
                Frm94.CB4 = 0
            End If
        End If

        'If Not IsNull(rs!no_rujukan_book) Then 'No rujukan tempahan
        '    Frm94.L9_Text = rs!no_rujukan_book
        'Else
        '    Frm94.L9_Text = 1
        'End If
        If Not IsNull(rs!ResitNo) Then
            Frm94.L10_Text = rs!ResitNo 'No. invoice
        Else
            Frm94.L10_Text = 1 'No. invoice
        End If
        If Not IsNull(rs!no_rujukan_tak_rasmi) Then
            Frm94.L32_Text = rs!no_rujukan_tak_rasmi 'No. invoice tidak rasmi
        Else
            Frm94.L32_Text = 1 'No. invoice tidak rasmi
        End If

        If rs!ScannerMode = 1 Then
            Frm94.CB1 = 1
        Else
            Frm94.CB1 = 0
        End If
    End If
End If

rs.Close
Set rs = Nothing

skip_setting:

user_level = MDI_frm1.L4_Text

If user_level = "Administration" Or user_level = "Guest/User" Then

    Frm94.CB9.Visible = False
    Frm94.Label24.Visible = False
    Frm94.Label30.Visible = False
    
Else
    
    If G_GST_SYSTEM = "YES" Then
        Frm94.CB9.Visible = True
        Frm94.Label24.Visible = True
        Frm94.Label30.Visible = True
        
        If G_INVOICE_RASMI = 0 Then
            Frm94.CB9 = 1
        Else
            Frm94.CB9 = 0
        End If
        
    Else
        Frm94.CB9.Visible = False
        Frm94.Label24.Visible = False
        Frm94.Label30.Visible = False
    End If
End If
End Sub
Sub Frm94_Call_Product_Detail()
'on error resume next
Frm94_LM_KOD_PURITY = vbNullString

Frm94_LM_No_SIRI = UCase(Frm94.TB1) 'No. Siri Produk

'###Carian Data Basic Bagi Item Ini### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from Data_Database where no_siri_Produk='" & Frm94_LM_No_SIRI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Not IsNull(rs!cawangan) Then
        
        If MDI_frm1.L20_Text <> rs!cawangan Then
            
            MsgBox "Stok ini adalah milik cawangan [" & rs!cawangan & "]. Anda tidak dibenarkan untuk jual barang ini.", vbExclamation, "Info"
            
            rs.Close
            Set rs = Nothing
            
            Exit Sub
            
        End If
    
    End If
        
    If rs!StatusItem = "10" Then
        If Not IsNull(rs!receiving_Status) Then
        
            If rs!receiving_Status = 2 Or rs!receiving_Status = 3 Then
                
                Frm94.TB1 = vbNullString
                
                MsgBox "Barang/item trade in tidak dibenarkan untuk dijual.", vbExclamation, "Info"
                
                rs.Close
                Set rs = Nothing
                
                Exit Sub
                'GoTo end_jualan:
            
            End If
                
            If rs!receiving_Status = 0 Or rs!receiving_Status = 2 Then
                
                If Frm94.L15_Text = 0 Then 'Jenis Barang Tempahan , 0 : Barang Kemas , 1 : Barang Permata
                    Frm94.TB2 = Frm94_LM_No_SIRI 'No. Siri Produk
                    Frm94.TB3 = Format(rs!beza_berat, "0.00") 'Berat Asal (g)
                    Frm94.TB4 = Format(rs!beza_berat, "0.00") 'Berat Jualan (g)
                    
                    Frm94.TB4.Locked = False
                    Frm94.TB5.Locked = False
                    Frm94.TB7.Locked = True
                    
                    Frm94.TB4.BackColor = &HFFFFFF
                    Frm94.TB5.BackColor = &HFFFFFF
                    Frm94.TB7.BackColor = &H8000000A
                    
                    If Frm94.L5_Text = 1 Then
                        If Frm94.L14_Text = "1" Then
                            If Not IsNull(rs!Upah_Jualan) Then Frm94.TB6 = Format(rs!Upah_Jualan, "0.00") 'Upah Jualan Kepada Pelanggan (RM/g)
                        ElseIf Frm94.L14_Text = "2" Then
                            If Not IsNull(rs!Upah_Member) Then Frm94.TB6 = Format(rs!Upah_Member, "0.00") 'Upah Jualan Kepada Member (RM/g)
                        ElseIf Frm94.L14_Text = "3" Then
                            If Not IsNull(rs!Upah_Pengedar) Then Frm94.TB6 = Format(rs!Upah_Pengedar, "0.00") 'Upah Jualan Kepada Pengedar (RM/g)
                        ElseIf Frm94.L14_Text = "4" Then
                            If Not IsNull(rs!Upah_RAF) Then Frm94.TB6 = Format(rs!Upah_RAF, "0.00") 'Upah Jualan Kepada RAF (RM/g)
                        ElseIf Frm94.L14_Text = "5" Then
                            If Not IsNull(rs!upah_normal_dealer) Then Frm94.TB6 = Format(rs!upah_normal_dealer, "0.00") 'Upah Jualan Kepada Normal Dealer (RM/g)
                        ElseIf Frm94.L14_Text = "6" Then
                            If Not IsNull(rs!upah_master_dealer) Then Frm94.TB6 = Format(rs!upah_master_dealer, "0.00") 'Upah Jualan Kepada Master Dealer (RM/g)
                        End If
                    End If
                    
                    If Not IsNull(rs!kod_Purity) Then
                        Frm94_LM_KOD_PURITY = rs!kod_Purity 'Kod Purity
                    End If
                    
                    If Not IsNull(rs!kategori_Produk) Then Frm94.L3_Text = rs!kategori_Produk 'Kategori Produk
                    Frm94_LM_READY_TO_SAVE = 1 'Flag : Ready To Save
                Else
                    rs.Close
                    Set rs = Nothing
                    
                    Frm94.TB1 = vbNullString
                    MsgBox "Item Ini [" & Frm94_LM_No_SIRI & "] adalah BARANG KEMAS , tempahan pelanggan ini adalah BARANG PERMATA.", vbExclamationm, "Info"
                    Frm94.TB1.SetFocus
                    Exit Sub
                End If
            Else
                If Frm94.L15_Text = 1 Then 'Jenis Barang Tempahan , 0 : Barang Kemas , 1 : Barang Permata
                    Frm94.TB2 = Frm94_LM_No_SIRI 'No. Siri Produk
                    
                    Frm94.TB3 = vbNullString
                    Frm94.TB4 = vbNullString
                    Frm94.TB5 = vbNullString
                    
                    Frm94.TB4.Locked = True
                    Frm94.TB5.Locked = True
                    Frm94.TB7.Locked = False
                    
                    Frm94.TB4.BackColor = &H8000000A
                    Frm94.TB5.BackColor = &H8000000A
                    Frm94.TB7.BackColor = &HFFFFFF
                    
                    If Frm94.L5_Text = 1 Then
                        If Frm94.L14_Text = "1" Then
                            If Not IsNull(rs!code_Supplier) Then Frm94.TB7 = Format(rs!code_Supplier, "0.00") 'Harga Jualan Kepada Pelanggan (RM)
                        ElseIf Frm94.L14_Text = "2" Then
                            If Not IsNull(rs!HargaJualan_Member) Then Frm94.TB7 = Format(rs!HargaJualan_Member, "0.00") 'Harga Jualan Kepada Member (RM)
                        ElseIf Frm94.L14_Text = "3" Then
                            If Not IsNull(rs!HargaJualan_Pengedar) Then Frm94.TB7 = Format(rs!HargaJualan_Pengedar, "0.00") 'Harga Jualan Kepada Pengedar (RM)
                        ElseIf Frm94.L14_Text = "4" Then
                            If Not IsNull(rs!HargaJualan_RAF) Then Frm94.TB7 = Format(rs!HargaJualan_RAF, "0.00") 'Harga Jualan Kepada RAF (RM)
                        ElseIf Frm94.L14_Text = "5" Then
                            If Not IsNull(rs!hargajualan_normal_dealer) Then Frm94.TB7 = Format(rs!hargajualan_normal_dealer, "0.00") 'Harga Jualan Kepada Normal Dealer (RM)
                        ElseIf Frm94.L14_Text = "6" Then
                            If Not IsNull(rs!hargajualan_master_dealer) Then Frm94.TB7 = Format(rs!hargajualan_master_dealer, "0.00") 'Harga Jualan Kepada Master Dealer (RM)
                        End If
                    End If
                    
                    If Not IsNull(rs!kategori_Produk) Then Frm94.L3_Text = rs!kategori_Produk 'Kategori Produk
                    'Frm94.L13_Text = 1 'Flag Kategori Produk , 0 : BK , 1 : Permata
                    
                    Frm94_LM_PERMATA = 1
                Else
                    rs.Close
                    Set rs = Nothing
                    
                    Frm94.TB1 = vbNullString
                    
                    MsgBox "Item Ini [" & Frm94_LM_No_SIRI & "] adalah BARANG PERMATA , tempahan pelanggan ini adalah BARANG KEMAS.", vbExclamationm, "Info"
                    Frm94.TB1.SetFocus
                    Exit Sub
                End If
            End If
        End If
        
        If Not IsNull(rs!kategori_Produk) Then Frm94.L3_Text = rs!kategori_Produk 'Kategori Produk
        
        'If Frm94_LM_PERMATA = 1 Then
        '    Frm94.TB7 = Format(Frm94_LM_HARGA_JUALAN, "0.00")
        'End If
    ElseIf rs!StatusItem = "11" Then
        MsgBox "Item Ini Telah Dijual. No. Siri Produk [" & Frm94_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm94.TB1 = vbNullString
        Frm94.TB1.SetFocus
    ElseIf rs!StatusItem = "12" Then
        MsgBox "Item Ini Telah Dijual Secara Potong. No. Siri Produk [" & Frm94_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm94.TB1 = vbNullString
        Frm94.TB1.SetFocus
    ElseIf rs!StatusItem = "13" Then
        MsgBox "Item Ini Telah Dijual Secara Potong. No. Siri Produk [" & Frm94_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm94.TB1 = vbNullString
        Frm94.TB1.SetFocus
    ElseIf rs!StatusItem = "14" Or rs!StatusItem = "21" Or rs!StatusItem = "22" Then
        MsgBox "Item Ini Telah Ditempah Oleh Pelanggan. No. Siri Produk [" & Frm94_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm94.TB1 = vbNullString
        Frm94.TB1.SetFocus
    ElseIf rs!StatusItem = "15" Or rs!StatusItem = "19" Or rs!StatusItem = "20" Then
        MsgBox "Item Ini Telah Dibeli Secara Ansuran. No. Siri Produk [" & Frm94_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm94.TB1 = vbNullString
        Frm94.TB1.SetFocus
    ElseIf rs!StatusItem = "16" Then
        MsgBox "Item Ini Telah Dihantar Ke Ar-Rahnu. No. Siri Produk [" & Frm94_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm94.TB1 = vbNullString
        Frm94.TB1.SetFocus
    ElseIf rs!StatusItem = "17" Then
        MsgBox "Item Ini Telah Dijual Secara ETA. No. Siri Produk [" & Frm94_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm94.TB1 = vbNullString
        Frm94.TB1.SetFocus
    ElseIf rs!StatusItem = "23" Then
        MsgBox "Item Ini Telah Dihantar Ke Supplier/Kilang. No. Siri Produk [" & Frm94_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm94.TB1 = vbNullString
        Frm94.TB1.SetFocus
    ElseIf rs!StatusItem = "24" Then
        MsgBox "Item Ini Telah Dihantar Ke Supplier/Kilang. No. Siri Produk [" & Frm94_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm94.TB1 = vbNullString
        Frm94.TB1.SetFocus
    ElseIf rs!StatusItem = "25" Then
        MsgBox "Item Ini Telah Diagihkan Ke Cawangan. No. Siri Produk [" & Frm94_LM_No_SIRI & "]", vbExclamation, "Info"

        Frm94.TB1 = vbNullString
        Frm94.TB1.SetFocus
    ElseIf rs!StatusItem = "26" Then
        MsgBox "Item Ini Telah Dijual Oleh Cawangan. No. Siri Produk [" & Frm94_LM_No_SIRI & "]", vbExclamation, "Info"

        Frm94.TB1 = vbNullString
        Frm94.TB1.SetFocus
    ElseIf rs!StatusItem = "0" Then
        MsgBox "Item Ini Telah Dipadamkan Dari Sistem. No. Siri Produk [" & Frm94_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm94.TB1 = vbNullString
        Frm94.TB1.SetFocus
    End If
Else
    Frm94.TB1 = vbNullString
    
    MsgBox "No. Siri Produk Ini [" & Frm94_LM_No_SIRI & "] Tidak Dijumpai.", vbExclamation, "Info"
    
    Frm94.TB1.SetFocus
End If

rs.Close
Set rs = Nothing

'###Carian Data Basic Bagi Item Ini### - End

'###Periksa Data Produk### - Start
If Frm94_LM_READY_TO_SAVE = 1 And Frm94.L5_Text = 1 Then 'Flag : Ready To Save
    If Frm94_LM_KOD_PURITY <> vbNullString Then
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from hargaemas where Purity='" & Frm94_LM_KOD_PURITY & "' AND cawangan='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Frm94.L14_Text = "1" Then
                If IsNumeric(rs!Harga_Pelanggan) Then Frm94.TB5 = Format(rs!Harga_Pelanggan, "0.00") 'Harga Semasa Bagi Pelanggan (RM/g)
            ElseIf Frm94.L14_Text = "2" Then
                If IsNumeric(rs!Harga_Member) Then Frm94.TB5 = Format(rs!Harga_Member, "0.00") 'Harga Semasa Bagi Member (RM/g)
            ElseIf Frm94.L14_Text = "3" Then
                If IsNumeric(rs!Harga_Pengedar) Then Frm94.TB5 = Format(rs!Harga_Pengedar, "0.00") 'Harga Semasa Bagi Pengedar (RM/g)
            ElseIf Frm94.L14_Text = "4" Then
                If IsNumeric(rs!Harga_RAF) Then Frm94.TB5 = Format(rs!Harga_RAF, "0.00") 'Harga Semasa Bagi RAF (RM/g)
            ElseIf Frm94.L14_Text = "5" Then
                If IsNumeric(rs!harga_nd) Then Frm94.TB5 = Format(rs!harga_nd, "0.00") 'Harga Semasa Bagi Normal Dealer (RM/g)
            ElseIf Frm94.L14_Text = "6" Then
                If IsNumeric(rs!harga_md) Then Frm94.TB5 = Format(rs!harga_md, "0.00") 'Harga Semasa Bagi Master Dealer (RM/g)
            End If
        End If
        
        rs.Close
        Set rs = Nothing
    End If
End If

Frm94.TB1 = vbNullString
'###Periksa Data Produk### - End
End Sub
Sub Frm94_invoice_deposit_tempahan()
'on error resume next
DATA_FOUND = 0

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

'### Reset Maklumat Pembeli ### - Start
Report48.Sections("Section1").Controls("L3").Caption = vbNullString 'No invoice
Report48.Sections("Section1").Controls("L4").Caption = vbNullString 'Tarikh
Report48.Sections("Section1").Controls("L5").Caption = vbNullString 'Nama
Report48.Sections("Section1").Controls("L7").Caption = vbNullString 'No. Telefon
'### Reset Maklumat Pembeli ### - End

Report48.Sections("Section1").Controls("L8").Caption = vbNullString
Report48.Sections("Section1").Controls("L9").Caption = vbNullString
Report48.Sections("Section1").Controls("L10").Caption = vbNullString
Report48.Sections("Section1").Controls("L11").Caption = vbNullString
Report48.Sections("Section1").Controls("L12").Caption = vbNullString
Report48.Sections("Section1").Controls("L13").Caption = vbNullString
Report48.Sections("Section1").Controls("L14").Caption = vbNullString
Report48.Sections("Section1").Controls("L15").Caption = vbNullString
Report48.Sections("Section1").Controls("L16").Caption = vbNullString
Report48.Sections("Section1").Controls("L17").Caption = vbNullString
Report48.Sections("Section1").Controls("L18").Caption = vbNullString

Report48.Sections("Section1").Controls("L26").Caption = vbNullString

'### Reset maklumat kedai ### - Start
Report48.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report48.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report48.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report48.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report48.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report48.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report48.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report48.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report48.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report48.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
    If Not IsNull(rs!gst_ari_nashi) Then
        If rs!gst_ari_nashi = 0 Then
            Report48.Sections("Section4").Controls("L205").Caption = "INVOICE"
        ElseIf rs!gst_ari_nashi = 1 Then
            Report48.Sections("Section4").Controls("L205").Caption = "TAX INVOICE"
        End If
    Else
        Report48.Sections("Section4").Controls("L205").Caption = "INVOICE"
    End If
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

LM_NO_PEKERJA = vbNullString

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 40_tempahan_deposit where no_resit_tempahan='" & G_No_INV_BOOK & "' AND status_invoice = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!no_pekerja) Then LM_NO_PEKERJA = rs!no_pekerja
    If Not IsNull(rs!no_resit_tempahan) Then
        Report48.Sections("Section1").Controls("L3").Caption = rs!no_resit_tempahan 'No. Invoice
        Frm93_LM_No_VOUCHER = rs!no_resit_tempahan 'No. Invoice
    End If
    If Not IsNull(rs!tarikh) Then Report48.Sections("Section1").Controls("L4").Caption = rs!tarikh 'Tarikh
    If Not IsNull(rs!Nama) Then Report48.Sections("Section1").Controls("L5").Caption = rs!Nama 'Nama
    If Not IsNull(rs!no_tel) Then Report48.Sections("Section1").Controls("L7").Caption = rs!no_tel 'No. Telefon
    
    If Not IsNull(rs!no_siri_Produk) Then
        Report48.Sections("Section1").Controls("L8").Caption = rs!no_siri_Produk 'No. Siri Produk
    Else
        Report48.Sections("Section1").Controls("L8").Caption = "-" 'No. Siri Produk
    End If
    If Not IsNull(rs!kategori_Produk) Then
        Report48.Sections("Section1").Controls("L9").Caption = rs!kategori_Produk 'Kategori Produk
    Else
        Report48.Sections("Section1").Controls("L9").Caption = "-" 'Kategori Produk
    End If
    If Not IsNull(rs!purity) Then
        Report48.Sections("Section1").Controls("L10").Caption = rs!purity 'Purity
    Else
        Report48.Sections("Section1").Controls("L10").Caption = "-" 'Purity
    End If
    If Not IsNull(rs!purity) Then
        Report48.Sections("Section1").Controls("L10").Caption = rs!purity 'Purity
    Else
        Report48.Sections("Section1").Controls("L10").Caption = "-" 'Purity
    End If
    
    If Not IsNull(rs!jenis_tempahan) Then
        If rs!jenis_tempahan = 0 Then
            Report48.Sections("Section1").Controls("L21").Caption = " / " 'Jenis Tempahan (Barang Baru)
            Report48.Sections("Section1").Controls("L22").Caption = " " 'Jenis Tempahan (Barang Kedai)
            
            If Not IsNull(rs!type_barang_kemas) Then
                If rs!type_barang_kemas = 0 Then
                    If Not IsNull(rs!anggaran_berat) Then
                        Report48.Sections("Section1").Controls("L11").Caption = Format(rs!anggaran_berat, "#,##0.00 g") 'Anggaran Berat
                    Else
                        Report48.Sections("Section1").Controls("L11").Caption = "-" 'Anggaran Berat
                    End If
                    If Not IsNull(rs!harga_Semasa) Then
                        Report48.Sections("Section1").Controls("L12").Caption = Format(rs!harga_Semasa, "#,##0.00") 'Harga Emas Semasa
                    Else
                        Report48.Sections("Section1").Controls("L12").Caption = "-" 'Harga Emas Semasa
                    End If
                Else
                    Report48.Sections("Section1").Controls("L11").Caption = "-" 'Anggaran Berat
                    Report48.Sections("Section1").Controls("L12").Caption = "-" 'Harga Emas Semasa
                End If
                
                If Not IsNull(rs!UPAH) Then
                    Report48.Sections("Section1").Controls("L13").Caption = Format(rs!UPAH, "#,##0.00") 'Upah
                Else
                    Report48.Sections("Section1").Controls("L13").Caption = "-" 'Upah
                End If
                If Not IsNull(rs!anggaran_harga) Then
                    Report48.Sections("Section1").Controls("L14").Caption = Format(rs!anggaran_harga, "#,##0.00") 'Anggaran Harga
                Else
                    Report48.Sections("Section1").Controls("L14").Caption = "-" 'Anggaran Harga
                End If
                    
            End If
        ElseIf rs!jenis_tempahan = 1 Then
            Report48.Sections("Section1").Controls("L21").Caption = " " 'Jenis Tempahan (Barang Baru)
            Report48.Sections("Section1").Controls("L22").Caption = " / " 'Jenis Tempahan (Barang Kedai)
            
            If Not IsNull(rs!type_barang_kemas) Then
                If rs!type_barang_kemas = 0 Then
                    If Not IsNull(rs!berat_jualan) Then
                        Report48.Sections("Section1").Controls("L11").Caption = Format(rs!berat_jualan, "#,##0.00") 'Anggaran Berat
                    Else
                        Report48.Sections("Section1").Controls("L11").Caption = "-" 'Anggaran Berat
                    End If
                    If Not IsNull(rs!harga_Semasa) Then
                        Report48.Sections("Section1").Controls("L12").Caption = Format(rs!harga_Semasa, "#,##0.00") 'Harga Emas Semasa
                    Else
                        Report48.Sections("Section1").Controls("L12").Caption = "-" 'Harga Emas Semasa
                    End If
                Else
                    Report48.Sections("Section1").Controls("L11").Caption = "-" 'Anggaran Berat
                    Report48.Sections("Section1").Controls("L12").Caption = "-" 'Harga Emas Semasa
                End If
                
                If Not IsNull(rs!UPAH) Then
                    Report48.Sections("Section1").Controls("L13").Caption = Format(rs!UPAH, "#,##0.00") 'Upah
                Else
                    Report48.Sections("Section1").Controls("L13").Caption = "-" 'Upah
                End If
                If Not IsNull(rs!anggaran_harga) Then
                    Report48.Sections("Section1").Controls("L14").Caption = Format(rs!anggaran_harga, "#,##0.00") 'Anggaran Harga
                Else
                    Report48.Sections("Section1").Controls("L14").Caption = "-" 'Anggaran Harga
                End If
                    
            End If

        End If
    End If
    If Not IsNull(rs!jumlah_deposit_tunai) Then
        Report48.Sections("Section1").Controls("L15").Caption = Format(rs!jumlah_deposit_tunai, "#,##0.00") 'Jumlah Deposit Tunai
    Else
        Report48.Sections("Section1").Controls("L15").Caption = Format(0, "#,##0.00") 'Jumlah Deposit
    End If
    If Not IsNull(rs!flag_trade_in) Then
        If rs!flag_trade_in = 0 Then
            Report48.Sections("Section1").Controls("L16").Caption = "-" 'Nilaian Voucher
            Report48.Sections("Section1").Controls("L17").Caption = "-" 'Nilaian Voucher
            Report48.Sections("Section1").Controls("L23").Caption = " " 'Trade In
        ElseIf rs!flag_trade_in = 1 Then
            Report48.Sections("Section1").Controls("L23").Caption = " / " 'Trade In
            If Not IsNull(rs!no_resit_trade_in) Then
                Report48.Sections("Section1").Controls("L16").Caption = rs!no_resit_trade_in 'No. Voucher Trade In
            Else
                Report48.Sections("Section1").Controls("L16").Caption = "-" 'Anggaran Harga
            End If
            If Not IsNull(rs!nilaian_trade_in) Then
                Report48.Sections("Section1").Controls("L17").Caption = Format(rs!nilaian_trade_in, "#,##0.00") 'Nilaian Voucher
            Else
                Report48.Sections("Section1").Controls("L17").Caption = "-" 'Anggaran Harga
            End If
        End If
    End If
    If Not IsNull(rs!jumlah_dengan_gst) Then
        Report48.Sections("Section1").Controls("L18").Caption = Format(rs!jumlah_dengan_gst, "#,##0.00") 'Jumlah Deposit
    Else
        Report48.Sections("Section1").Controls("L18").Caption = Format(0, "#,##0.00") 'Jumlah Deposit
    End If
    If Not IsNull(rs!remarks) Then
        Report48.Sections("Section1").Controls("L26").Caption = rs!remarks 'Remarks
    Else
        Report48.Sections("Section1").Controls("L26").Caption = " " 'Remarks
    End If
End If

rs.Close
Set rs = Nothing

If LM_NO_PEKERJA <> vbNullString Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where NoPekerja='" & LM_NO_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        If Not IsNull(rs!Samaran) Then
            Report48.Sections("Section1").Controls("L19").Caption = rs!Samaran
        End If
        
    End If
    
    rs.Close
    Set rs = Nothing
            
End If

'### Paparan Resit ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 40_tempahan_deposit where no_resit_tempahan='" & G_No_INV_BOOK & "' AND status_invoice = 1", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report48.DataSource = rs
    Report48.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing
'### Paparan Resit ### - End

End Sub
Sub Frm94_invoice_siap_tempahan()
'on error resume next
Dim Frm94_LM_JUMLAH As Double

Frm94_LM_JUMLAH = 0
DATA_FOUND = 0

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

'### Reset Maklumat Pembeli ### - Start
Report49.Sections("Section1").Controls("L3").Caption = vbNullString 'No. Invoice
Report49.Sections("Section1").Controls("L4").Caption = vbNullString 'Tarikh
Report49.Sections("Section1").Controls("L5").Caption = vbNullString 'Nama
Report49.Sections("Section1").Controls("L7").Caption = vbNullString 'No. Telefon
Report49.Sections("Section1").Controls("L8").Caption = "-" 'No. siri produk
Report49.Sections("Section1").Controls("L9").Caption = "-" 'Kategori produk
Report49.Sections("Section1").Controls("L10").Caption = vbNullString 'Caption lebihan bayaran
Report49.Sections("Section1").Controls("L10").Visible = False
Report49.Sections("Section1").Controls("L11").Caption = "-" 'Berat
Report49.Sections("Section1").Controls("L12").Caption = "-" 'Harga semasa
Report49.Sections("Section1").Controls("L13").Caption = "-" 'Upah
Report49.Sections("Section1").Controls("L14").Caption = "-" 'Harga barang
Report49.Sections("Section1").Controls("L15").Caption = "0.00" 'Deposit
Report49.Sections("Section1").Controls("L16").Caption = "0.00" 'Potongan trade in
Report49.Sections("Section1").Controls("L18").Caption = "0.00" 'Harga barang dengan GST
Report49.Sections("Section1").Controls("L19").Caption = "0.00" 'Jumlah GST
Report49.Sections("Section1").Controls("L20").Caption = "0.00" 'Bayaran
Report49.Sections("Section1").Controls("L21").Caption = vbNullString 'Jenis tempahan - baru
Report49.Sections("Section1").Controls("L22").Caption = vbNullString 'Jenis tempahan - barang kedai
Report49.Sections("Section1").Controls("L25").Caption = "0.00" 'Harga barang tanpa GST
Report49.Sections("Section1").Controls("L26").Caption = vbNullString 'Remarks
Report49.Sections("Section1").Controls("L27").Caption = vbNullString 'Jurujual
Report49.Sections("Section1").Controls("L28").Caption = vbNullString 'No voucher trade in
'### Reset Maklumat Pembeli ### - End

'### Reset maklumat kedai ### - Start
Report49.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report49.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report49.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report49.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report49.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report49.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report49.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report49.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report49.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report49.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
    If Not IsNull(rs!gst_ari_nashi) Then
        If rs!gst_ari_nashi = 0 Then
            Report49.Sections("Section4").Controls("L205").Caption = "INVOICE"
        ElseIf rs!gst_ari_nashi = 1 Then
            Report49.Sections("Section4").Controls("L205").Caption = "TAX INVOICE"
        End If
    Else
        Report49.Sections("Section4").Controls("L205").Caption = "INVOICE"
    End If
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 42_tempahan_siap where no_resit_tempahan='" & G_No_INV_BOOK & "' AND status_invoice = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!no_pekerja) Then LM_NO_PEKERJA = rs!no_pekerja
    If Not IsNull(rs!no_rujukan_tempahan) Then Frm93_LM_No_RUJUKAN = rs!no_rujukan_tempahan
    If Not IsNull(rs!no_resit_tempahan) Then
        Report49.Sections("Section1").Controls("L3").Caption = rs!no_resit_tempahan 'No. Invoice
        Frm93_LM_No_VOUCHER = rs!no_resit_tempahan 'No. Invoice
    End If
    If Not IsNull(rs!write_timestamp) Then Report49.Sections("Section1").Controls("L4").Caption = rs!write_timestamp 'Tarikh
    If Not IsNull(rs!Nama) Then Report49.Sections("Section1").Controls("L5").Caption = rs!Nama 'Nama
    If Not IsNull(rs!no_tel) Then Report49.Sections("Section1").Controls("L7").Caption = rs!no_tel 'No. Telefon
    
    If Not IsNull(rs!no_siri_Produk) Then
        Report49.Sections("Section1").Controls("L8").Caption = rs!no_siri_Produk 'No. Siri Produk
    Else
        Report49.Sections("Section1").Controls("L8").Caption = "-" 'No. Siri Produk
    End If
    If Not IsNull(rs!kategori_Produk) Then
        Report49.Sections("Section1").Controls("L9").Caption = rs!kategori_Produk 'Kategori Produk
    Else
        Report49.Sections("Section1").Controls("L9").Caption = "-" 'Kategori Produk
    End If
    
    If Not IsNull(rs!jenis_tempahan) Then
        If rs!jenis_tempahan = 0 Then
            Report49.Sections("Section1").Controls("L21").Caption = " / " 'Jenis Tempahan (Barang Baru)
            Report49.Sections("Section1").Controls("L22").Caption = " " 'Jenis Tempahan (Barang Kedai)
            
            If Not IsNull(rs!type_barang_kemas) Then
                If rs!type_barang_kemas = 0 Then
                    If Not IsNull(rs!berat_jualan) Then
                        Report49.Sections("Section1").Controls("L11").Caption = Format(rs!berat_jualan, "#,##0.00") 'Berat
                    Else
                        Report49.Sections("Section1").Controls("L11").Caption = "-" 'Berat
                    End If
                    If Not IsNull(rs!harga_Semasa) Then
                        Report49.Sections("Section1").Controls("L12").Caption = Format(rs!harga_Semasa, "#,##0.00") 'Harga Emas Semasa
                    Else
                        Report49.Sections("Section1").Controls("L12").Caption = "-" 'Harga Emas Semasa
                    End If
                Else
                    Report49.Sections("Section1").Controls("L11").Caption = "-" 'Anggaran Berat
                    Report49.Sections("Section1").Controls("L12").Caption = "-" 'Harga Emas Semasa
                End If
                
                If Not IsNull(rs!UPAH) Then
                    Report49.Sections("Section1").Controls("L13").Caption = Format(rs!UPAH, "#,##0.00") 'Upah
                Else
                    Report49.Sections("Section1").Controls("L13").Caption = "-" 'Upah
                End If
                If Not IsNull(rs!harga_dengan_gst) Then
                    Report49.Sections("Section1").Controls("L14").Caption = Format(rs!harga_dengan_gst, "#,##0.00") 'Harga
                Else
                    Report49.Sections("Section1").Controls("L14").Caption = "0.00" 'Harga
                End If
                    
            End If
        ElseIf rs!jenis_tempahan = 1 Then
            Report49.Sections("Section1").Controls("L21").Caption = " " 'Jenis Tempahan (Barang Baru)
            Report49.Sections("Section1").Controls("L22").Caption = " / " 'Jenis Tempahan (Barang Kedai)
            
            If Not IsNull(rs!type_barang_kemas) Then
                If rs!type_barang_kemas = 0 Then
                    If Not IsNull(rs!berat_jualan) Then
                        Report49.Sections("Section1").Controls("L11").Caption = Format(rs!berat_jualan, "#,##0.00") 'Berat
                    Else
                        Report49.Sections("Section1").Controls("L11").Caption = "-" 'Berat
                    End If
                    If Not IsNull(rs!harga_Semasa) Then
                        Report49.Sections("Section1").Controls("L12").Caption = Format(rs!harga_Semasa, "#,##0.00") 'Harga Emas Semasa
                    Else
                        Report49.Sections("Section1").Controls("L12").Caption = "-" 'Harga Emas Semasa
                    End If
                Else
                    Report49.Sections("Section1").Controls("L11").Caption = "-" 'Berat
                    Report49.Sections("Section1").Controls("L12").Caption = "-" 'Harga Emas Semasa
                End If
                
                If Not IsNull(rs!UPAH) Then
                    Report49.Sections("Section1").Controls("L13").Caption = Format(rs!UPAH, "#,##0.00") 'Upah
                Else
                    Report49.Sections("Section1").Controls("L13").Caption = "-" 'Upah
                End If
                If Not IsNull(rs!harga_dengan_gst) Then
                    Report49.Sections("Section1").Controls("L14").Caption = Format(rs!harga_dengan_gst, "#,##0.00") 'Harga
                Else
                    Report49.Sections("Section1").Controls("L14").Caption = "0.00" 'Harga
                End If
                    
            End If

        End If
    End If
    If Not IsNull(rs!bayaran_sudah_jelas) Then
        Report49.Sections("Section1").Controls("L15").Caption = Format(rs!bayaran_sudah_jelas, "#,##0.00") 'Jumlah Deposit Tunai
    Else
        Report49.Sections("Section1").Controls("L15").Caption = "0.00" 'Jumlah Deposit
    End If
    If Not IsNull(rs!flag_trade_in) Then
        If rs!flag_trade_in = 0 Then
            'Report49.Sections("Section1").Controls("L23").Caption = " " 'Trade In
            
            Report49.Sections("Section1").Controls("L16").Caption = "-" 'Nilaian Voucher
            'Report49.Sections("Section1").Controls("L17").Caption = "-" 'Nilaian Voucher
        ElseIf rs!flag_trade_in = 1 Then
            'Report49.Sections("Section1").Controls("L23").Caption = " / " 'Trade In
            If Not IsNull(rs!no_resit_trade_in) Then
                Report49.Sections("Section1").Controls("L28").Caption = "No. voucher trade in : " & rs!no_resit_trade_in 'No. Voucher Trade In
                Report49.Sections("Section1").Controls("L28").Visible = True
            Else
                Report49.Sections("Section1").Controls("L28").Caption = vbNullString 'No. Voucher Trade In
                Report49.Sections("Section1").Controls("L28").Visible = False
            End If
            If Not IsNull(rs!nilaian_trade_in) Then
                Report49.Sections("Section1").Controls("L16").Caption = Format(rs!nilaian_trade_in, "#,##0.00") 'Nilaian Voucher
            Else
                Report49.Sections("Section1").Controls("L16").Caption = "0.00" 'Nilaian Voucher
            End If
        End If
    End If
    If Not IsNull(rs!harga_dengan_gst) Then
        Report49.Sections("Section1").Controls("L18").Caption = Format(rs!harga_dengan_gst, "#,##0.00") 'Baki Tanpa GST
        'If IsNumeric(rs!baki) Then Frm94_LM_JUMLAH = rs!baki
    Else
        Report49.Sections("Section1").Controls("L18").Caption = "0.00" 'Baki Tanpa GST
    End If
    If Not IsNull(rs!jumlah_gst) Then
        Report49.Sections("Section1").Controls("L19").Caption = Format(rs!jumlah_gst, "#,##0.00") 'Jumlah GST
    Else
        Report49.Sections("Section1").Controls("L19").Caption = "0.00" 'Jumlah GST
    End If
    If Not IsNull(rs!harga_tanpa_gst) Then
        Report49.Sections("Section1").Controls("L25").Caption = Format(rs!harga_tanpa_gst, "#,##0.00") 'Harga barang tanpa GST
    Else
        Report49.Sections("Section1").Controls("L25").Caption = "0.00" 'Jumlah Adjustment
    End If
    
End If

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where no_resit='" & G_No_INV_BOOK & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!jumlah_perlu_bayar) Then
        Report49.Sections("Section1").Controls("L20").Caption = Format(rs!jumlah_perlu_bayar, "#,##0.00") 'Jumlah Bayaran
        Frm94_LM_JUMLAH = rs!jumlah_perlu_bayar
    Else
        Report49.Sections("Section1").Controls("L20").Caption = "0.00" 'Jumlah Bayaran
    End If
    If Not IsNull(rs!flag_bayaran) Then
        If rs!flag_bayaran = 1 Then
            Report49.Sections("Section1").Controls("L10").Visible = True
            Report49.Sections("Section1").Controls("L10").Caption = "RM " & Format(Frm94_LM_JUMLAH, "#,##0.00") & " perlu dipulangkan kepada pembeli ini kerana lebihan bayaran oleh pembeli." 'Jumlah Bayaran
        End If
    End If
End If

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 40_tempahan_deposit where no_rujukan_tempahan='" & Frm93_LM_No_RUJUKAN & "' AND status_invoice = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!remarks) Then
        Report49.Sections("Section1").Controls("L26").Caption = rs!remarks 'Remarks
    Else
        Report49.Sections("Section1").Controls("L26").Caption = " " 'Remarks
    End If
End If

rs.Close
Set rs = Nothing

If LM_NO_PEKERJA <> vbNullString Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where NoPekerja='" & LM_NO_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        If Not IsNull(rs!Samaran) Then
            Report49.Sections("Section1").Controls("L27").Caption = rs!Samaran
        End If
        
    End If
    
    rs.Close
    Set rs = Nothing
            
End If

'### Paparan Resit ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 42_tempahan_siap where no_resit_tempahan='" & G_No_INV_BOOK & "' AND status_invoice = 1", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report49.DataSource = rs
    Report49.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing
'### Paparan Resit ### - End
End Sub
Sub Frm93_padam_data_tempahan()
'on error resume next
Dim Frm93_LM_SIMPANAN_DIGUNAKAN As Double
Dim Frm93_LM_SIMPANAN_ASAL As Double
Dim Frm93_LM_BERAT_JUALAN As Double
Dim Frm93_LM_BERAT_ASAL As Double
Dim Frm93_LM_BEZA_BERAT As Double
Dim Frm93_LM_BERAT_PULANGAN As Double

Frm93_LM_BERAT_PULANGAN = 0
Frm93_LM_BERAT_ASAL = 0
Frm93_LM_BERAT_JUALAN = 0
Frm93_LM_BEZA_BERAT = 0
Frm93_LM_SIMPANAN_DIGUNAKAN = 0
Frm93_LM_SIMPANAN_ASAL = 0
Frm93_LM_FLAG_TI = 0
Frm93_LM_FLAG_BARANG_KEDAI = 0 'Flag Barang Kedai
DATA_FOUND = 0
LM_FOUND = 0
Frm93_LM_STATUS = vbNullString
Frm93_LM_FLAG_BARANG_KEMAS = 0

frm93_LM_No_ID = vbNullString

If IsNumeric(Frm93.LV1.SelectedItem.Index) Then
    
    frm93_LM_No_ID = Frm93.LV1.ListItems(Frm93.LV1.SelectedItem.Index)
    
    If frm93_LM_No_ID <> vbNullString Then
    
        If frm93_LM_No_ID <> vbNullString Then
            
            LM_NOW = Now
            
            If G_TEMPAHAN = 0 Then '0 : Padam data , 1 : Tukar status kepada belum siap
            
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 40_tempahan_deposit where ID='" & frm93_LM_No_ID & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Not IsNull(rs!invoice_siap) Then Frm93_LM_No_RESIT = rs!invoice_siap 'No. Resit
                    If Not IsNull(rs!no_rujukan_tempahan) Then Frm93_LM_No_RUJUKAN = rs!no_rujukan_tempahan
                    LM_FOUND = 1
                End If
                
                rs.Close
                Set rs = Nothing
            
            ElseIf G_TEMPAHAN = 1 Then  '0 : Padam data , 1 : Tukar status kepada belum siap
            
    '### Padam Data Dari Senarai Tempahan (Deposit) ### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 40_tempahan_deposit where ID='" & frm93_LM_No_ID & "' AND status_invoice = 1", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                
                    If Not IsNull(rs!no_rujukan_tempahan) Then Frm93_LM_No_RUJUKAN = rs!no_rujukan_tempahan
                    
                    G_ID = rs!ID
                    Call recovery_40_tempahan_deposit
                    
                    If Not IsNull(rs!Status) Then
                        If rs!Status = "Siap" Then
                            LM_STATUS = 1
                        End If
                    End If
                    rs!Status = "Belum Siap"
                    If Not IsNull(rs!invoice_siap) Then Frm93_LM_No_RESIT = rs!invoice_siap 'No. Resit
                    If Not IsNull(rs!flag_trade_in) Then
                        If rs!flag_trade_in = 1 Then
                            Frm93_LM_FLAG_TI = 1
                            If Not IsNull(rs!no_resit_trade_in) Then Frm93_LM_No_RESIT_TI = rs!no_resit_trade_in 'No. Resit Trade In
                        End If
                    End If
                    
                    If Not IsNull(rs!jenis_tempahan) Then
                        If rs!jenis_tempahan = 1 Then
                            Frm93_LM_FLAG_BARANG_KEDAI = 1 'Flag Barang Kedai
                            If Not IsNull(rs!no_siri_Produk) Then Frm93_LM_No_SIRI = rs!no_siri_Produk 'No. Siri Produk
                        End If
                    End If
                    
                    If Not IsNull(rs!no_rujukan_pelanggan) Then Frm93_LM_No_ID_PEMBELI = rs!no_rujukan_pelanggan 'No. Rujukan Pembeli
                    
                    'rs!status_invoice = 0
                    
                    rs!terminal = G_TERMINAL
                    rs!write_timestamp2 = LM_NOW
                    rs!no_staff = G_LOGIN_USER 'No. staff yang memasukkan/edit data (Login)
                    LM_FOUND = 1
                    rs.Update
                    
                End If
                
                rs.Close
                Set rs = Nothing
    '### Padam Data Dari Senarai Tempahan (Deposit) ### - End
            
            End If
            
            If LM_FOUND = 1 Then
            
                G_JENIS_URUSAN = 11
                
                GoTo skip_carian_user:
                
                If MDI_frm1.L3_Text <> vbNullString Then
                
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                    
                        If Not IsNull(rs!NoPekerja) Then G_LOGIN_USER = rs!NoPekerja
                
                    End If
                    
                    rs.Close
                    Set rs = Nothing
                    
                End If
                '$$$ No. staff $$$ - End
                
skip_carian_user:

        '### Padam Data Dari Senarai Tempahan (Deposit) ### - Start
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 42_tempahan_siap where no_rujukan_tempahan='" & Frm93_LM_No_RUJUKAN & "' AND status_invoice = 1", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                
                    G_ID = rs!ID
                    Call recovery_42_tempahan_siap
                    
                    DATA_FOUND = 1
                    If Not IsNull(rs!no_siri_Produk) Then Frm93_LM_No_SIRI = rs!no_siri_Produk 'No. Siri Produk
                    If Not IsNull(rs!berat_jualan) Then Frm93_LM_BERAT_JUALAN = rs!berat_jualan 'Berat Jualan
                    If Not IsNull(rs!no_rujukan_tempahan) Then Frm93_LM_No_RUJUKAN = rs!no_rujukan_tempahan 'No. Rujukan Tempahan
                    If Not IsNull(rs!no_resit_tempahan) Then Frm93_LM_No_RESIT = rs!no_resit_tempahan 'No. Resit
                    If Not IsNull(rs!flag_trade_in) Then
                        If rs!flag_trade_in = 1 Then
                            Frm93_LM_FLAG_TI = 1
                            If Not IsNull(rs!no_resit_trade_in) Then Frm93_LM_No_RESIT_TI = rs!no_resit_trade_in 'No. Resit Trade In
                        End If
                    End If
                    If Not IsNull(rs!type_barang_kemas) Then
                        If rs!type_barang_kemas = 0 Then
                            Frm93_LM_FLAG_BARANG_KEMAS = 1
                        End If
                    End If
                    
                    If Not IsNull(rs!jenis_tempahan) Then
                        If rs!jenis_tempahan = 1 Then
                            Frm93_LM_FLAG_BARANG_KEDAI = 1 'Flag Barang Kedai
                            If Not IsNull(rs!no_siri_Produk) Then Frm93_LM_No_SIRI = rs!no_siri_Produk 'No. Siri Produk
                        End If
                    End If
                    
                    If Not IsNull(rs!no_rujukan_pelanggan) Then Frm93_LM_No_RUJUKAN_PEMBELI = rs!no_rujukan_pelanggan 'No. Rujukan Pembeli
                    
                    rs!terminal = G_TERMINAL
                    'LM_NOW = Now
                    rs!write_timestamp2 = LM_NOW
                    rs!status_invoice = 0 '0 : Tidak aktif (dibatalkan) , 1:  Aktif
                    rs!no_staff = G_LOGIN_USER 'No. staff yang memasukkan/edit data (Login)
                    'rs.Delete
                    rs.Update
                End If
                
                rs.Close
                Set rs = Nothing
        '### Padam Data Dari Senarai Tempahan (Deposit) ### - End
            
                If DATA_FOUND = 1 Then
        '### Ubah Status Tempahan ### - Start
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from 40_tempahan_deposit where no_rujukan_tempahan='" & Frm93_LM_No_RUJUKAN & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                    
                        G_ID = rs!ID
                        Call recovery_40_tempahan_deposit
                        
                        rs!invoice_siap = Null
                        rs!Status = "Belum Siap"
                        rs!terminal = G_TERMINAL
                        rs!write_timestamp2 = LM_NOW
                        rs!no_staff = G_LOGIN_USER 'No. staff yang memasukkan/edit data (Login)
                        rs.Update

                    End If
                    
                    rs.Close
                    Set rs = Nothing
        '### Ubah Status Tempahan ### - End
        
        '### Pulangkan Status Barang Trade In ### - Start
                    If Frm93_LM_FLAG_TI = 1 Then '0 : Tiada Urusan Trade In , 1 : Ada Urusan Trade In
                        Set rs = New ADODB.Recordset
                        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                        rs.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & Frm93_LM_No_RESIT_TI & "'", cn, adOpenKeyset, adLockOptimistic
                        
                        If Not rs.EOF Then
                        
                            G_ID = rs!ID
                            Call recovery_16_gold_bar_belian
                
                            rs!trade_in_status = 0
                            rs!no_staff = G_LOGIN_USER 'No. Pekerja
                            rs!terminal = G_TERMINAL
                            rs!write_timestamp2 = LM_NOW
                            rs!jenis_urusan = G_JENIS_URUSAN
                            rs!remarks = "Ubah status flag trade in bagi pembatalan data tempahan"
                            rs.Update

                        End If

                        rs.Close
                        Set rs = Nothing
                    End If
        '### Pulangkan Status Barang Trade In ### - End
    
        '### Pulangkan Status Item Dalam Database ### - Start
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from Data_Database where no_siri_produk='" & Frm93_LM_No_SIRI & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                    
                        G_ID = rs!ID
                        Call recovery_data_database

                        If Frm93_LM_FLAG_BARANG_KEMAS = 1 Then
                        
                            If IsNumeric(rs!beza_berat) Then Frm93_LM_BEZA_BERAT = rs!beza_berat
                            
                            If IsNumeric(rs!Berat) Then Frm93_LM_BERAT_ASAL = rs!Berat

                            Frm93_LM_BERAT_PULANGAN = Frm93_LM_BERAT_JUALAN + Frm93_LM_BEZA_BERAT
                            
                            rs!beza_berat = Format(Frm93_LM_BERAT_JUALAN + Frm93_LM_BEZA_BERAT, "0.00") 'Beza Berat
                            
                            If G_TEMPAHAN = 0 Then '0 : Padam data , 1 : Tukar status kepada belum siap
                            
                                If Frm93_LM_BERAT_PULANGAN = Frm93_LM_BERAT_ASAL Then
                                
                                    rs!StatusItem = 10
                                    
                                Else
                                
                                    rs!StatusItem = 12
                                    
                                End If
                                
                            Else
                                
                                If Frm93_LM_FLAG_BARANG_KEDAI = 1 Then 'Flag Barang Kedai
                                
                                    rs!StatusItem = 14
                                    
                                Else
                                
                                    rs!StatusItem = 10
                                    
                                End If
                            
                            End If
                            
                        Else
                            
                            If G_TEMPAHAN = 0 Then '0 : Padam data , 1 : Tukar status kepada belum siap
                                
                                'If Frm93_LM_FLAG_BARANG_KEDAI = 1 Then 'Flag Barang Kedai
                                
                                    rs!StatusItem = 10
                                    
                                'Else
                                    
                                    
                                
                                'End If
                                
                            Else
                                
                                If Frm93_LM_FLAG_BARANG_KEDAI = 1 Then 'Flag Barang Kedai
                                
                                    rs!StatusItem = 14
                                    
                                Else
                                    
                                    rs!StatusItem = 10
                                
                                End If
                            
                            End If
                            
                        End If
                        
                        rs!write_timestamp2 = LM_NOW
                        rs!no_pekerja = G_LOGIN_USER
                        rs!terminal = G_TERMINAL
                        rs!Menu = 4
                        rs.Update

                    End If
                    
                    rs.Close
                    Set rs = Nothing
        '### Pulangkan Status Item Dalam Database ### - End
    
        '###Padam Akaun Tempahan### - Start
                    Frm93_LM_FLAG_SAVING = 0
                    
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from 22_jualan where no_resit='" & Frm93_LM_No_RESIT & "' AND Status = 1", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                        
                        G_ID = rs!ID
                        Call recovery_22_jualan
                        
                        If Not IsNull(rs!duit_simpanan_kedai) Then
                            If Format(rs!duit_simpanan_kedai, "0.00") <> "0.00" Then
                                If IsNumeric(rs!duit_simpanan_kedai) Then Frm93_LM_SIMPANAN_DIGUNAKAN = rs!duit_simpanan_kedai
                                Frm93_LM_FLAG_SAVING = 1
                            End If
                        End If
                    
                        rs!Status = 0
                        rs!terminal = G_TERMINAL
                        rs!no_staff = G_LOGIN_USER
                        rs!write_timestamp2 = LM_NOW
                        rs!Menu = 4
                        rs.Update
                        
                    End If
                    
                    rs.Close
                    Set rs = Nothing
        '###Padam Akaun Tempahan### - End

        '###Update Simpanan Duit Di Kedai### - Start
                    If Frm93_LM_FLAG_SAVING = 1 Then
                    
                        Set rs = New ADODB.Recordset
                        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                        rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm93_LM_No_RUJUKAN_PEMBELI & "'", cn, adOpenKeyset, adLockOptimistic
                        
                        If Not rs.EOF Then
                    
                            G_ID = rs!ID
                            Call recovery_senarai_pelanggan
                    
                            If Not IsNull(rs!baki_simpanan) Then
                                If IsNumeric(rs!baki_simpanan) Then Frm93_LM_SIMPANAN_ASAL = rs!baki_simpanan 'Baki Simpanan Pelanggan Ini (RM)
                            End If
                            
                            rs!baki_simpanan = Format(Frm93_LM_SIMPANAN_ASAL + Frm93_LM_SIMPANAN_DIGUNAKAN, "0.00") 'Baki Simpanan Terkini Pelanggan Ini (RM)
                            
                            rs!write_timestamp2 = LM_NOW
                            rs!no_staff = G_LOGIN_USER 'No. Pekerja
                            rs!terminal = G_TERMINAL
                            rs!jenis_urusan = G_JENIS_URUSAN
                    
                            rs.Update
                        End If
                        
                        rs.Close
                        Set rs = Nothing
                        
        '###Padam Rekod Bayaran Dalam Table Simpanan### - Start
                        Set rs = New ADODB.Recordset
                        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                        rs.Open "select * from 24_rekod_kewangan_pelanggan where NO_RESIT='" & Frm93_LM_No_RESIT & "'", cn, adOpenKeyset, adLockOptimistic
                        
                        If Not rs.EOF Then
                        
                            G_ID = rs!ID
                            Call recovery_24_rekod_kewangan_pelanggan
                    
                            rs.Delete
                            rs.Update
                        End If
                        
                        rs.Close
                        Set rs = Nothing
        '###Padam Rekod Bayaran Dalam Table Simpanan### - End
                    
                    End If
    '###Update Simpanan Duit Di Kedai### - End
                
                    If G_TEMPAHAN = 1 Then '0 : Padam data , 1 : Tukar status kepada belum siap

            '### Update Log ### - Start
                        'User = MDI_frm1.L3_Text
                        LogAct_Memory = "[" & G_LOGIN_USER & "] Tukar status tempahan kepada belum siap. No. ID tempahan [" & frm93_LM_No_ID & "]"
                        LogDate_Memory = DateTime.Date & " " & DateTime.Time$
                        Call UpdateLog_Database
            '### Update Log ### - End
            
                        GM_NEXT_PREV = 2
                        
                        Call frm93_tempahan_header
                        Call frm93_tempahan
                            
                        MsgBox "Data tempahan telah berjaya ditukar status kepada belum siap.", vbInformation, "Info"
                    
                    End If
                
                End If
                
            Else
            
                MsgBox "Maklumat tempahan siap tidak berjaya dipadamkan. Sila hubungi pihak Sankyu System.", vbCritical, "Info"
                
            End If
        
        End If
        
    End If
    
End If
End Sub
Sub frm93_kira_harga_tempahan1()
'on error resume next
'Pengiraan bagi harga (ANGGRAN harga) bagi tempahan baru (Barang yang belum ada di kedai)
Dim Frm93_BERAT As Double
Dim Frm93_HARGA_PER_GRAM As Double
Dim Frm93_UPAH As Double

If GLOBAL_DISABLE = 0 Then

    Frm93_BERAT = 0
    Frm93_HARGA_PER_GRAM = 0
    Frm93_UPAH = 0
    
    If ((Frm93.TB1 <> vbNullString And IsNumeric(Frm93.TB1)) And (Frm93.TB2 <> vbNullString And IsNumeric(Frm93.TB2)) And (Frm93.TB3 <> vbNullString And IsNumeric(Frm93.TB3))) Then
        Frm93_BERAT = Frm93.TB1 'Berat
        Frm93_HARGA_PER_GRAM = Frm93.TB2 'Harga Per Gram
        Frm93_UPAH = Frm93.TB3 'Upah
        
        Frm93.TB4 = Format((Frm93_BERAT * Frm93_HARGA_PER_GRAM) + Frm93_UPAH, "#,##0.00") 'Harga Asal
    Else
        Frm93.TB4 = "0.00" 'Harga Asal
    End If
    
End If
End Sub
Sub frm93_kira_harga_tempahan2()
'on error resume next
'Pengiraan bagi harga bagi tempahan barangan kedai (barang sedia ada)
Dim Frm93_BERAT As Double
Dim Frm93_HARGA_PER_GRAM As Double
Dim Frm93_UPAH As Double

If GLOBAL_DISABLE = 0 Then

    Frm93_BERAT = 0
    Frm93_HARGA_PER_GRAM = 0
    Frm93_UPAH = 0
    
    If ((Frm93.TB8 <> vbNullString And IsNumeric(Frm93.TB8)) And (Frm93.TB9 <> vbNullString And IsNumeric(Frm93.TB9)) And (Frm93.TB10 <> vbNullString And IsNumeric(Frm93.TB10))) Then
        Frm93_BERAT = Frm93.TB8 'Berat
        Frm93_HARGA_PER_GRAM = Frm93.TB9 'Harga Per Gram
        Frm93_UPAH = Frm93.TB10 'Upah
        
        Frm93.TB11 = Format((Frm93_BERAT * Frm93_HARGA_PER_GRAM) + Frm93_UPAH, "#,##0.00") 'Harga Asal
    Else
        Frm93.TB11 = "0.00" 'Harga Asal
    End If
    
End If
End Sub
Sub frm93_harga_jualan()
'on error resume next
Dim Frm93_HARGA_ASAL As Double
Dim Frm93_ADJUSTMENT As Double

If GLOBAL_DISABLE = 0 Then

    Frm93_HARGA_ASAL = 0
    Frm93_ADJUSTMENT = 0
    
    If ((Frm93.TB11 <> vbNullString And IsNumeric(Frm93.TB11)) And (Frm93.TB12 <> vbNullString And IsNumeric(Frm93.TB12))) Then
        Frm93_HARGA_ASAL = Frm93.TB11 'Harga Asal
        Frm93_ADJUSTMENT = Frm93.TB12 'Adjustment
        
        Frm93.TB13 = Format(Frm93_HARGA_ASAL - Frm93_ADJUSTMENT, "#,##0.00") 'Harga Jualan
    Else
        Frm93.TB13 = "0.00" 'Harga Jualan
    End If
    
End If
End Sub
Sub Frm93_jurujual()
'On Error Resume Next
If MDI_frm1.L3_Text <> vbNullString Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        Frm93.CBB3 = rs!Samaran & "  |  " & rs!NoPekerja
        
    Else
        
        Frm93.CBB3.AddItem "" & "  |  " & rs!Samaran
        Frm93.CBB3 = "" & "  |  " & rs!Samaran
        
    End If
    
    rs.Close
    Set rs = Nothing

    If G_LOCK_JURUJUAL = "YES" Then
    
        Frm93.CBB3.Enabled = False
        Frm93.CBB3.BackColor = &H8000000A

    Else
    
        Frm93.CBB3.Enabled = True
        Frm93.CBB3.BackColor = &HFFFFFF

    End If
    
End If
End Sub
Sub Frm94_jurujual()
'On Error Resume Next
If MDI_frm1.L3_Text <> vbNullString Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        Frm94.CBB1 = rs!Samaran & "  |  " & rs!NoPekerja
        
    Else
        
        Frm94.CBB1.AddItem "" & "  |  " & rs!Samaran
        Frm94.CBB1 = "" & "  |  " & rs!Samaran
        
    End If
    
    rs.Close
    Set rs = Nothing
    
    If G_LOCK_JURUJUAL = "YES" Then
    
        Frm94.CBB1.Enabled = False
        Frm94.CBB1.BackColor = &H8000000A

    Else
    
        Frm94.CBB1.Enabled = True
        Frm94.CBB1.BackColor = &HFFFFFF

    End If

End If
End Sub
Sub Frm93_kira_jumlah_bayaran()
'on error resume next
Dim Frm93_LM_TUNAI As Double
Dim Frm93_LM_BANK As Double
Dim Frm93_LM_KREDIT As Double
Dim Frm93_LM_SIMPANAN As Double

Frm93_LM_TUNAI = 0
Frm93_LM_BANK = 0
Frm93_LM_KREDIT = 0
Frm93_LM_SIMPANAN = 0

If IsNumeric(frm130.TB27) Then
    Frm93_LM_TUNAI = frm130.TB27
End If
If IsNumeric(frm130.TB28) Then
    Frm93_LM_BANK = frm130.TB28
End If
If IsNumeric(frm130.TB29) Then
    Frm93_LM_KREDIT = frm130.TB29
End If
If IsNumeric(frm130.TB21) Then
    Frm93_LM_SIMPANAN = frm130.TB21
End If

Frm93.TB32 = Format(Frm93_LM_TUNAI + Frm93_LM_BANK + Frm93_LM_KREDIT + Frm93_LM_SIMPANAN, "#,##0.00")  'Jumlah Bayaran Keseluruhan
End Sub
Sub Frm93_kira_caj_kad_kredit()
'on error resume next
Dim LM_KAD_KREDIT As Double
Dim LM_CAJ As Double

LM_KAD_KREDIT = 0
LM_CAJ = 0

If ((frm130.TB29 <> vbNullString And IsNumeric(frm130.TB29)) And (frm130.L31_Text <> vbNullString And IsNumeric(frm130.L31_Text))) Then
    LM_KAD_KREDIT = frm130.TB29
    LM_CAJ = frm130.L31_Text
    
    frm130.L32_Text = Format(LM_KAD_KREDIT * (LM_CAJ / 100), "#,##0.00")
Else
    frm130.L32_Text = "0.00"
End If
End Sub
Sub Frm93_kira_caj_gst_kad_kredit()
'on error resume next
Dim LM_CAJ As Double
Dim LM_RATE_GST As Double

LM_CAJ = 0
LM_RATE_GST = 0

If ((frm130.L32_Text <> vbNullString And IsNumeric(frm130.L32_Text)) And (Frm93.L19_Text <> vbNullString And IsNumeric(Frm93.L19_Text))) Then
    LM_CAJ = frm130.L32_Text
    LM_RATE_GST = Frm93.L19_Text
    
    frm130.L81_Text = Format(LM_CAJ * (LM_RATE_GST / 100), "#,##0.00")
Else
    frm130.L81_Text = "0.00"
End If
End Sub
Sub Frm93_kira_potongan_kad_kredit()
'on error resume next
Dim LM_KAD_KREDIT As Double
Dim LM_CAJ As Double
Dim LM_GST As Double

LM_KAD_KREDIT = 0
LM_CAJ = 0
LM_GST = 0

If ((frm130.TB29 <> vbNullString And IsNumeric(frm130.TB29)) And (frm130.L32_Text <> vbNullString And IsNumeric(frm130.L32_Text)) And (frm130.L81_Text <> vbNullString And IsNumeric(frm130.L81_Text))) Then
    LM_KAD_KREDIT = frm130.TB29
    LM_CAJ = frm130.L32_Text
    LM_GST = frm130.L81_Text
    
    frm130.L82_Text = Format(LM_KAD_KREDIT + LM_CAJ + LM_GST, "#,##0.00")
Else
    frm130.L82_Text = "0.00"
End If
End Sub
Sub frm93_jumlah_deposit()
'On Error Resume Next
Dim Frm93_LM_DEPOSIT_TUNAI As Double
Dim Frm93_LM_DEPOSIT_TRADE_IN As Double

Frm93_LM_DEPOSIT_TUNAI = 0
Frm93_LM_DEPOSIT_TRADE_IN = 0

'If GLOBAL_DISABLE = 0 Then
    
    If ((Frm93.TB20 <> vbNullString And IsNumeric(Frm93.TB20)) And (Frm93.TB22 <> vbNullString And IsNumeric(Frm93.TB22))) Then
        Frm93_LM_DEPOSIT_TUNAI = Frm93.TB20 'Deposit Tunai
        Frm93_LM_DEPOSIT_TRADE_IN = Frm93.TB22 'Deposit Trade In
        
        Frm93.TB23 = Format(Frm93_LM_DEPOSIT_TUNAI + Frm93_LM_DEPOSIT_TRADE_IN, "#,##0.00") 'Jumlah Deposit
    Else
        Frm93.TB23 = "0.00" 'Jumlah Deposit
    End If
    
'End If
End Sub
Sub frm94_kira_harga_emas()
'On Error Resume Next
Dim Frm94_BERAT As Double
Dim Frm94_HARGA_PER_GRAM As Double
Dim Frm94_UPAH As Double

If GLOBAL_DISABLE = 0 Then

    Frm94_BERAT = 0
    Frm94_HARGA_PER_GRAM = 0
    Frm94_UPAH = 0
    
    If ((Frm94.TB4 <> vbNullString And IsNumeric(Frm94.TB4)) And (Frm94.TB5 <> vbNullString And IsNumeric(Frm94.TB5)) And (Frm94.TB6 <> vbNullString And IsNumeric(Frm94.TB6))) Then
        Frm94_BERAT = Frm94.TB4 'Berat
        Frm94_HARGA_PER_GRAM = Frm94.TB5 'Harga Per Gram
        Frm94_UPAH = Frm94.TB6 'Upah
        
        Frm94.TB7 = Format((Frm94_BERAT * Frm94_HARGA_PER_GRAM) + Frm94_UPAH, "#,##0.00") 'Harga Asal
    Else
        Frm94.TB7 = "0.00" 'Harga Asal
    End If
    
End If
End Sub
Sub frm94_kira_harga_bersih()
'on error resume next
Dim Frm94_HARGA_ASAL As Double
Dim Frm94_ADJUSTMENT As Double

If GLOBAL_DISABLE = 0 Then

    Frm94_HARGA_ASAL = 0
    Frm94_ADJUSTMENT = 0
    
    If ((Frm94.TB7 <> vbNullString And IsNumeric(Frm94.TB7)) And (Frm94.TB8 <> vbNullString And IsNumeric(Frm94.TB8))) Then
        Frm94_HARGA_ASAL = Frm94.TB7 'Harga Asal
        Frm94_ADJUSTMENT = Frm94.TB8 'Adjustment
        
        Frm94.TB9 = Format(Frm94_HARGA_ASAL - Frm94_ADJUSTMENT, "#,##0.00") 'Harga Jualan
    Else
        Frm94.TB9 = "0.00" 'Harga Jualan
    End If
    
End If
End Sub
Sub frm94_kiraan_cukai_gst()
'on error resume next
Dim Frm94_LM_KADAR_GST As Double
Dim Frm94_LM_HARGA As Double
Dim Frm94_LM_GST As Double

Frm94_LM_KADAR_GST = 0
Frm94_LM_HARGA = 0
Frm94_LM_GST = 0

If GLOBAL_DISABLE = 0 Then

    If (Frm94.TB10 <> vbNullString And IsNumeric(Frm94.TB10)) And (Frm94.L8_Text <> vbNullString And IsNumeric(Frm94.L8_Text)) Then

        If IsNumeric(Frm94.L8_Text) Then Frm94_LM_KADAR_GST = Frm94.L8_Text 'Jumlah Kadar GST (%)
        If IsNumeric(Frm94.TB10) Then Frm94_LM_HARGA = Frm94.TB10 'Jumlah Bayaran (RM)
        
        If Frm94.CB3 = 1 Then
        
            Frm94_LM_GST = 0
            
            Frm94.TB17 = Format(Frm94_LM_GST, "#,##0.00") 'Jumlah cukai GST
            Frm94.TB23 = Format(Frm94_LM_HARGA + Frm94_LM_GST, "#,##0.00") 'Jumlah harga tanpa GST
            Frm94.TB24 = Format(Frm94_LM_HARGA + Frm94_LM_GST, "#,##0.00") 'Jumlah harga dengan GST
            
        ElseIf Frm94.CB4 = 1 Then
        
            Frm94_LM_GST = Frm94_LM_HARGA * (Frm94_LM_KADAR_GST / 100)
            
            Frm94.TB17 = Format(Frm94_LM_GST, "#,##0.00") 'Jumlah cukai GST
            Frm94.TB23 = Format(Frm94_LM_HARGA, "#,##0.00")  'Jumlah harga tanpa GST
            Frm94.TB24 = Format(Frm94_LM_HARGA + Frm94_LM_GST, "#,##0.00") 'Jumlah harga dengan GST
        
        ElseIf Frm94.CB5 = 1 Then
        
            Frm94_LM_GST = Frm94_LM_HARGA - (Frm94_LM_HARGA / (1 + (Frm94_LM_KADAR_GST / 100)))
            
            Frm94.TB17 = Format(Frm94_LM_GST, "#,##0.00") 'Jumlah cukai GST
            Frm94.TB23 = Format((Frm94_LM_HARGA / (1 + (Frm94_LM_KADAR_GST / 100))), "#,##0.00")  'Jumlah harga tanpa GST
            Frm94.TB24 = Format(Frm94_LM_HARGA, "#,##0.00")  'Jumlah harga dengan GST
        
        Else
        
            Frm94.TB17 = Format(0, "#,##0.00") 'Jumlah cukai GST
            Frm94.TB23 = Format(0, "#,##0.00")  'Jumlah harga tanpa GST
            Frm94.TB24 = Format(0, "#,##0.00") 'Jumlah harga dengan GST
        
        End If
        
        
    Else
    
        Frm94.TB17 = Format(0, "#,##0.00") 'Jumlah cukai GST
        Frm94.TB23 = Format(0, "#,##0.00")  'Jumlah harga tanpa GST
        Frm94.TB24 = Format(0, "#,##0.00") 'Jumlah harga dengan GST
        
    End If
    
End If
End Sub
Sub frm94_kira_baki()
'on error resume next
Dim Frm94_HARGA As Double
Dim Frm94_TRADE_IN As Double
Dim Frm94_DEPOSIT As Double
Dim Frm94_JUMLAH_DEPO As Double

If GLOBAL_DISABLE = 0 Then

    Frm94_HARGA = 0
    Frm94_TRADE_IN = 0
    Frm94_DEPOSIT = 0
    Frm94_JUMLAH_DEPO = 0
    
    If ((Frm94.TB22 <> vbNullString And IsNumeric(Frm94.TB22)) And (Frm94.TB11 <> vbNullString And IsNumeric(Frm94.TB11)) And (Frm94.TB24 <> vbNullString And IsNumeric(Frm94.TB24))) Then
        
        If IsNumeric(Frm94.TB24) Then Frm94_HARGA = Frm94.TB24 'Harga Barang
        If IsNumeric(Frm94.TB22) Then Frm94_TRADE_IN = Frm94.TB22 'Jumlah trade in
        If IsNumeric(Frm94.TB11) Then Frm94_DEPOSIT = Frm94.TB11 'Jumlah deposit
        
        Frm94_JUMLAH_DEPO = Frm94_TRADE_IN + Frm94_DEPOSIT
        
        If Frm94_HARGA >= Frm94_JUMLAH_DEPO Then
            Frm94.L6_Text = "Baki (RM)                                 :"
            Frm94.TB12 = Format(Frm94_HARGA - Frm94_JUMLAH_DEPO, "#,##0.00") 'Baki
            Frm94.L13_Text.Visible = False
            frm130.TB33 = Format(Frm94_HARGA - Frm94_JUMLAH_DEPO, "#,##0.00") 'Baki
        Else
            Frm94.L6_Text = "Lebihan Kedai Perlu Bayar (RM)  :"
            Frm94.TB12 = Format(Frm94_JUMLAH_DEPO - Frm94_HARGA, "#,##0.00") 'Tunai
            Frm94.L13_Text.Visible = True
            
            frm130.TB33 = Format(0, "#,##0.00") 'Tunai
            frm130.TB28 = Format(0, "#,##0.00") 'Bank in
            frm130.TB29 = Format(0, "#,##0.00") 'Kad kredit
            frm130.TB21 = Format(0, "#,##0.00") 'Simpanan
        End If
        
    Else
        Frm94.L6_Text = "Baki (RM)                                 :"
        Frm94.TB12 = "0.00" 'Baki
        Frm94.L13_Text.Visible = False
    End If
    
End If
End Sub
Sub Frm94_kira_jumlah_bayaran()
'on error resume next
Dim Frm94_LM_TUNAI As Double
Dim Frm94_LM_BANK As Double
Dim Frm94_LM_KREDIT As Double
Dim Frm94_LM_SIMPANAN As Double

Frm94_LM_TUNAI = 0
Frm94_LM_BANK = 0
Frm94_LM_KREDIT = 0
Frm94_LM_SIMPANAN = 0

If IsNumeric(frm130.TB27) Then
    Frm94_LM_TUNAI = frm130.TB27
End If
If IsNumeric(frm130.TB28) Then
    Frm94_LM_BANK = frm130.TB28
End If
If IsNumeric(frm130.TB29) Then
    Frm94_LM_KREDIT = frm130.TB29
End If
If IsNumeric(frm130.TB21) Then
    Frm94_LM_SIMPANAN = frm130.TB21
End If

frm130.TB32 = Format(Frm94_LM_TUNAI + Frm94_LM_BANK + Frm94_LM_KREDIT + Frm94_LM_SIMPANAN, "#,##0.00")  'Jumlah Bayaran Keseluruhan
End Sub
Sub Frm94_kira_caj_kad_kredit()
'on error resume next
Dim LM_KAD_KREDIT As Double
Dim LM_CAJ As Double

LM_KAD_KREDIT = 0
LM_CAJ = 0

If ((frm130.TB29 <> vbNullString And IsNumeric(frm130.TB29)) And (frm130.L31_Text <> vbNullString And IsNumeric(frm130.L31_Text))) Then
    LM_KAD_KREDIT = frm130.TB29
    LM_CAJ = frm130.L31_Text
    
    frm130.L32_Text = Format(LM_KAD_KREDIT * (LM_CAJ / 100), "#,##0.00")
Else
    frm130.L32_Text = "0.00"
End If
End Sub
Sub Frm94_kira_caj_gst_kad_kredit()
'on error resume next
Dim LM_CAJ As Double
Dim LM_RATE_GST As Double

LM_CAJ = 0
LM_RATE_GST = 0

If ((frm130.L32_Text <> vbNullString And IsNumeric(frm130.L32_Text)) And (Frm94.L8_Text <> vbNullString And IsNumeric(Frm94.L8_Text))) Then
    LM_CAJ = frm130.L32_Text
    LM_RATE_GST = Frm94.L8_Text
    
    frm130.L81_Text = Format(LM_CAJ * (LM_RATE_GST / 100), "#,##0.00")
Else
    frm130.L81_Text = "0.00"
End If
End Sub
Sub Frm94_kira_potongan_kad_kredit()
'on error resume next
Dim LM_KAD_KREDIT As Double
Dim LM_CAJ As Double
Dim LM_GST As Double

LM_KAD_KREDIT = 0
LM_CAJ = 0
LM_GST = 0

If ((frm130.TB29 <> vbNullString And IsNumeric(frm130.TB29)) And (frm130.L32_Text <> vbNullString And IsNumeric(frm130.L32_Text)) And (frm130.L81_Text <> vbNullString And IsNumeric(frm130.L81_Text))) Then
    LM_KAD_KREDIT = frm130.TB29
    LM_CAJ = frm130.L32_Text
    LM_GST = frm130.L81_Text
    
    frm130.L82_Text = Format(LM_KAD_KREDIT + LM_CAJ + LM_GST, "#,##0.00")
Else
    frm130.L82_Text = "0.00"
End If
End Sub


