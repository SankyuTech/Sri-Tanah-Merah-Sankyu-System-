Attribute VB_Name = "Module44"
Sub Frm87_Initial_Setting()
'on error resume next
GLOBAL_DISABLE = 0

Frm87.Pic2.Left = 120
Frm87.Pic2.Top = 240
Frm87.Pic4.Left = 120
Frm87.Pic4.Top = 240
Frm87.Pic5.Left = 120
Frm87.Pic5.Top = 240
Frm87.Pic6.Left = 240
Frm87.Pic6.Top = 1920
Frm87.Pic7.Left = 120
Frm87.Pic7.Top = 240

Frm87.Pic2.Visible = False
Frm87.Pic4.Visible = False
Frm87.Pic5.Visible = False
Frm87.Pic7.Visible = False

Frm87.CB14 = 1
Frm87.CB15 = 0
Frm87.CB20 = 0
Frm87.CB21 = 0
Frm87.CB24 = 1
Frm87.CB25 = 0
Frm87.CB26 = 0

Frm87.TB2 = vbNullString
Frm87.TB3 = vbNullString
Frm87.TB4 = vbNullString
Frm87.TB5 = vbNullString
Frm87.TB6 = vbNullString
Frm87.TB7 = vbNullString
Frm87.TB8 = vbNullString
Frm87.TB9 = vbNullString
Frm87.TB10 = vbNullString
Frm87.TB12 = "0.00"
Frm87.TB13 = "0.00"
Frm87.TB14 = "0.00"
Frm87.TB15 = "0.00"
Frm87.TB16 = "0.00"
Frm87.TB17 = "0.00"
Frm87.TB18 = "0.00"
Frm87.TB19 = "0.00"
Frm87.TB20 = "0.00"
Frm87.TB21 = "0.00"
Frm87.TB21 = "0.00"
Frm87.TB27 = "0.00"
Frm87.TB28 = "0.00"
Frm87.TB29 = "0.00"
Frm87.TB30 = "0.00"
Frm87.TB31 = "0.00"
Frm87.TB32 = "0.00"
Frm87.TB38 = "0.00"
Frm87.TB39 = "0.00"
Frm87.TB40 = "0.00"
Frm87.TB41 = vbNullString

Frm87.L19_Text = "0.00"
Frm87.L20_Text = "0.00"
Frm87.L21_Text = "0.00"

Frm87.L5_Text = vbNullString
Frm87.L6_Text = vbNullString
Frm87.L10_Text = vbNullString
Frm87.L11_Text = vbNullString
Frm87.L12_Text = vbNullString
Frm87.L13_Text = 0 'Flag Kategori Produk , 0 : BK , 1 : Permata
Frm87.L17_Text = 0
Frm87.L18_Text = vbNullString
Frm87.L30_Text = vbNullString

Frm87.L27_Text = "0.00"
Frm87.L28_Text = "0.00"
Frm87.L31_Text = "0.00"
Frm87.L32_Text = "0.00"

Frm87.L36_Text = 0
Frm87.L37_Text = 0
Frm87.L38_Text = 0
Frm87.L40_Text = 1

Frm87.L17_Text.BackStyle = 0
Frm87.L19_Text.BackStyle = 0
Frm87.L20_Text.BackStyle = 0
Frm87.L28_Text.BackStyle = 0
Frm87.L29_Text.BackStyle = 0
Frm87.L30_Text.BackStyle = 0

Frm87.CMD3.Visible = True

Frm87.DTPicker1 = DateTime.Date
Frm87.DTPicker2 = DateTime.Date

Frm87.CMD6.Visible = True
Frm87.CMD9.Visible = True
Frm87.CMD10.Visible = True
Frm87.CMD13.Visible = False
Frm87.CMD14.Visible = False
Frm87.CMD16.Visible = False
Frm87.CMD17.Visible = False

Frm87.CMD1.Enabled = True
Frm87.CMD3.Enabled = True

Frm87.TB8.Locked = False
Frm87.TB8.BackColor = &HFFFFFF
            
Frm87.CB14.Enabled = True
Frm87.CB15.Enabled = True

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Default1 = "Default" Then
        If Not IsNull(rs!gst_value) Then
            Frm87.L17_Text = rs!gst_value 'Jumlah Kadar GST
        Else
            Frm87.L17_Text = 0
        End If
        If Not IsNull(rs!gst_arinashi) Then
            If rs!gst_arinashi = 1 Then
                Frm87.CB18 = 0
                Frm87.CB19 = 1
                Frm87.CB22 = 0
                Frm87.CB23 = 1
            Else
                Frm87.CB22 = 1
                Frm87.CB23 = 0
            End If
        End If
        
        If Not IsNull(rs!cas_Kad_Kredit) Then Frm87.L31_Text = Format(rs!cas_Kad_Kredit, "0.00") 'Cas Kad Kredit
        If Not IsNull(rs!cas_debit_kad) Then Frm87.L32_Text = Format(rs!cas_debit_kad, "0.00") 'Cas Debit Kredit
        
        If Not IsNull(rs!no_rujukan_ansuran) Then
            Frm87.L11_Text = rs!no_rujukan_ansuran 'No. Rujukan Ansuran
        Else
            Frm87.L11_Text = 1
        End If
        If Not IsNull(rs!no_resit_ansuran) Then
            Frm87.L12_Text = rs!no_resit_ansuran 'No. Resit Ansuran
        Else
            Frm87.L12_Text = 1 'No. Resit Ansuran
        End If
        If rs!ScannerMode = 1 Then
            Frm87.CB13 = 1
        Else
            Frm87.CB13 = 0
        End If
    End If
End If

rs.Close
Set rs = Nothing

Frm87.CBB1.Clear
Frm87.CBB2.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where status='" & "Aktif" & "' AND InvestorSmall = 0 AND InvestorBig = 0", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    If rs!user_level = 1 Or rs!user_level = 2 Or rs!user_level = 3 Or rs!user_level = G_LEVEL_USER Then
        Frm87.CBB1.AddItem rs!Samaran & "  |  " & rs!NoPekerja
        Frm87.CBB2.AddItem rs!Samaran & "  |  " & rs!NoPekerja
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
End Sub
Sub Frm87_Call_Product_Detail()
'on error resume next
Frm87_LM_KOD_PURITY = vbNullString
Frm87_LM_KATEGORI = 1
'1:  Pelanggan Biasa
'2 : Member / Ahli
'3:  Pengedar
'4:  RAF
'5:  Normal Dealer(ND)
'6:  Master Dealer(MD)

Frm87_LM_No_SIRI = UCase(Frm87.TB8) 'No. Siri Produk

' ### Periksa kategori pembeli ### - Start
If Frm87.L6_Text <> vbNullString Then
    If Frm28.L5_Text <> vbNullString Then
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
        If Not rs.EOF Then
        
            If Not IsNull(rs!kategori_pelanggan) Then Frm87_LM_KATEGORI = rs!kategori_pelanggan
            
        End If
        
        rs.Close
        Set rs = Nothing
        
    End If
End If
' ### Periksa kategori pembeli ### - End

'###Carian Data Basic Bagi Item Ini### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from Data_Database where no_siri_Produk='" & Frm87_LM_No_SIRI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!StatusItem = "10" Then
        If Not IsNull(rs!receiving_Status) Then
            If rs!receiving_Status = 0 Or rs!receiving_Status = 2 Then
                Frm87.L10_Text = vbNullString
                
                Frm87.TB2 = Frm87_LM_No_SIRI 'No. Siri Produk
                Frm87.TB3 = Format(rs!beza_berat, "0.00") 'Berat Asal (g)
                Frm87.TB4 = Format(rs!beza_berat, "0.00") 'Berat Jualan (g)
                
                Frm87.TB4.Locked = False
                Frm87.TB5.Locked = False
                Frm87.TB6.Locked = False
                Frm87.TB7.Locked = True
                
                Frm87.TB4.BackColor = &HFFFFFF
                Frm87.TB5.BackColor = &HFFFFFF
                Frm87.TB6.BackColor = &HFFFFFF
                Frm87.TB7.BackColor = &H8000000A
                
                If Frm87_LM_KATEGORI = 1 Then
                    If Not IsNull(rs!Upah_Jualan) Then Frm87.TB6 = Format(rs!Upah_Jualan, "0.00") 'Upah Jualan Kepada Pelanggan (RM/g)
                ElseIf Frm87_LM_KATEGORI = 2 Then
                    If Not IsNull(rs!Upah_Member) Then Frm87.TB6 = Format(rs!Upah_Member, "0.00") 'Upah Jualan Kepada Member (RM/g)
                ElseIf Frm87_LM_KATEGORI = 3 Then
                    If Not IsNull(rs!Upah_Pengedar) Then Frm87.TB6 = Format(rs!Upah_Pengedar, "0.00") 'Upah Jualan Kepada Pengedar (RM/g)
                ElseIf Frm87_LM_KATEGORI = 4 Then
                    If Not IsNull(rs!HargaJualan_RAF) Then Frm87.TB6 = Format(rs!HargaJualan_RAF, "0.00") 'Upah Jualan Kepada RAF (RM/g)
                ElseIf Frm87_LM_KATEGORI = 5 Then
                    If Not IsNull(rs!upah_normal_dealer) Then Frm87.TB6 = Format(rs!upah_normal_dealer, "0.00") 'Upah Jualan Kepada Normal Dealer (RM/g)
                ElseIf Frm87_LM_KATEGORI = 6 Then
                    If Not IsNull(rs!upah_master_dealer) Then Frm87.TB6 = Format(rs!upah_master_dealer, "0.00") 'Upah Jualan Kepada Master Dealer (RM/g)
                End If
                
                If Not IsNull(rs!kod_Purity) Then
                    Frm87_LM_KOD_PURITY = rs!kod_Purity 'Kod Purity
                End If
                
                If Not IsNull(rs!kategori_Produk) Then Frm87.L10_Text = rs!kategori_Produk 'Kategori Produk
                Frm87.L13_Text = 0 'Flag Kategori Produk , 0 : BK , 1 : Permata
                Frm87_LM_READY_TO_SAVE = 1 'Flag : Ready To Save
                'Frm87.CB14 = 0
                'Frm87.CB15 = 0
                Frm87.CB14.Enabled = True
                Frm87.CB15.Enabled = True
            Else
                Frm87.TB2 = Frm87_LM_No_SIRI 'No. Siri Produk
                
                Frm87.L10_Text = vbNullString
                Frm87.TB3 = vbNullString
                Frm87.TB4 = vbNullString
                Frm87.TB5 = vbNullString
                Frm87.TB6 = vbNullString
                
                Frm87.TB4.Locked = True
                Frm87.TB5.Locked = True
                Frm87.TB6.Locked = True
                Frm87.TB7.Locked = False
                
                Frm87.TB4.BackColor = &H8000000A
                Frm87.TB5.BackColor = &H8000000A
                Frm87.TB6.BackColor = &H8000000A
                Frm87.TB7.BackColor = &HFFFFFF
                
                If Frm87_LM_KATEGORI = 1 Then
                    If Not IsNull(rs!code_Supplier) Then Frm87.TB7 = Format(rs!code_Supplier, "0.00") 'Harga Jualan Kepada Pelanggan (RM)
                ElseIf Frm87_LM_KATEGORI = 2 Then
                    If Not IsNull(rs!HargaJualan_Member) Then Frm87.TB7 = Format(rs!HargaJualan_Member, "0.00") 'Harga Jualan Kepada Member (RM)
                ElseIf Frm87_LM_KATEGORI = 3 Then
                    If Not IsNull(rs!HargaJualan_Pengedar) Then Frm87.TB7 = Format(rs!HargaJualan_Pengedar, "0.00") 'Harga Jualan Kepada Pengedar (RM)
                ElseIf Frm87_LM_KATEGORI = 4 Then
                    If Not IsNull(rs!HargaJualan_RAF) Then Frm87.TB7 = Format(rs!HargaJualan_RAF, "0.00") 'Harga Jualan Kepada RAF (RM)
                ElseIf Frm87_LM_KATEGORI = 5 Then
                    If Not IsNull(rs!hargajualan_normal_dealer) Then Frm87.TB7 = Format(rs!hargajualan_normal_dealer, "0.00") 'Harga Jualan Kepada Normal Dealer (RM)
                ElseIf Frm87_LM_KATEGORI = 6 Then
                    If Not IsNull(rs!hargajualan_master_dealer) Then Frm87.TB7 = Format(rs!hargajualan_master_dealer, "0.00") 'Harga Jualan Kepada Master Dealer (RM)
                End If
                
                Frm87.L13_Text = 1 'Flag Kategori Produk , 0 : BK , 1 : Permata
                Frm87.CB14 = 0
                Frm87.CB15 = 1
                Frm87.CB14.Enabled = False
                Frm87.CB15.Enabled = False
                Frm87_LM_PERMATA = 1
            End If
        End If
        
        Frm87.TB9 = "0.00"
        If Not IsNull(rs!kategori_Produk) Then Frm87.L10_Text = rs!kategori_Produk 'Kategori Produk
        
        If Frm87_LM_PERMATA = 1 Then
            Frm87.TB6 = Format(Frm87_LM_HARGA_JUALAN, "0.00")
        End If
    ElseIf rs!StatusItem = "11" Then
        MsgBox "Item Ini Telah Dijual. No. Siri Produk [" & Frm87_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm87.TB8 = vbNullString
        Frm87.TB8.SetFocus
    ElseIf rs!StatusItem = "12" Then
        MsgBox "Item Ini Telah Dijual Secara Potong. No. Siri Produk [" & Frm87_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm87.TB8 = vbNullString
        Frm87.TB8.SetFocus
    ElseIf rs!StatusItem = "13" Then
        MsgBox "Item Ini Telah Dijual Secara Potong. No. Siri Produk [" & Frm87_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm87.TB8 = vbNullString
        Frm87.TB8.SetFocus
    ElseIf rs!StatusItem = "14" Or rs!StatusItem = "21" Or rs!StatusItem = "22" Then
        MsgBox "Item Ini Telah Ditempah Oleh Pelanggan. No. Siri Produk [" & Frm87_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm87.TB8 = vbNullString
        Frm87.TB8.SetFocus
    ElseIf rs!StatusItem = "15" Or rs!StatusItem = "19" Or rs!StatusItem = "20" Then
        MsgBox "Item Ini Telah Dibeli Secara Ansuran. No. Siri Produk [" & Frm87_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm87.TB8 = vbNullString
        Frm87.TB8.SetFocus
    ElseIf rs!StatusItem = "16" Then
        MsgBox "Item Ini Telah Dihantar Ke Ar-Rahnu. No. Siri Produk [" & Frm87_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm87.TB8 = vbNullString
        Frm87.TB8.SetFocus
    ElseIf rs!StatusItem = "17" Then
        MsgBox "Item Ini Telah Dijual Secara ETA. No. Siri Produk [" & Frm87_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm87.TB8 = vbNullString
        Frm87.TB8.SetFocus
    ElseIf rs!StatusItem = "23" Then
        MsgBox "Item Ini Telah Dihantar Ke Supplier/Kilang. No. Siri Produk [" & Frm87_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm87.TB8 = vbNullString
        Frm87.TB8.SetFocus
    ElseIf rs!StatusItem = "24" Then
        MsgBox "Item Ini Telah Dihantar Ke Supplier/Kilang. No. Siri Produk [" & Frm87_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm87.TB8 = vbNullString
        Frm87.TB8.SetFocus
    ElseIf rs!StatusItem = "25" Then
        MsgBox "Item Ini Telah Diagihkan Ke Cawangan. No. Siri Produk [" & Frm87_LM_No_SIRI & "]", vbExclamation, "Info"

        Frm87.TB8 = vbNullString
        Frm87.TB8.SetFocus
    ElseIf rs!StatusItem = "26" Then
        MsgBox "Item Ini Telah Dijual Oleh Cawangan. No. Siri Produk [" & Frm87_LM_No_SIRI & "]", vbExclamation, "Info"

        Frm87.TB8 = vbNullString
        Frm87.TB8.SetFocus
    ElseIf rs!StatusItem = "0" Then
        MsgBox "Item Ini Telah Dipadamkan Dari Sistem. No. Siri Produk [" & Frm87_LM_No_SIRI & "]", vbExclamation, "Info"
        
        Frm87.TB8 = vbNullString
        Frm87.TB8.SetFocus
    ElseIf rs!StatusItem = "27" Or rs!StatusItem = "28" Then
        MsgBox "Item Ini Telah Dijual Dari Menu GDN. No. Siri Produk [" & Frm87_LM_No_SIRI & "]", vbExclamation, "Info"

        Frm87.TB8 = vbNullString
        Frm87.TB8.SetFocus
    ElseIf rs!StatusItem = "29" Then
        MsgBox "Item Ini Telah Diubah Status Kepada Hilang , Dicuri Dan Sebagainya. No. Siri Produk [" & Frm87_LM_No_SIRI & "]", vbExclamation, "Info"

        Frm87.TB8 = vbNullString
        Frm87.TB8.SetFocus
    End If
Else
    MsgBox "No. Siri Produk Ini [" & Frm87_LM_No_SIRI & "] Tidak Dijumpai.", vbExclamation, "Info"
    
    'Frm87.TB8 = vbNullString
    Frm87.TB8.SetFocus
End If

rs.Close
Set rs = Nothing

'###Carian Data Basic Bagi Item Ini### - End


'###Periksa Data Produk### - Start
If Frm87_LM_READY_TO_SAVE = 1 Then 'Flag : Ready To Save
    If Frm87_LM_KOD_PURITY <> vbNullString Then
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from hargaemas where Purity='" & Frm87_LM_KOD_PURITY & "' AND cawangan='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Frm87_LM_KATEGORI = 1 Then
                If IsNumeric(rs!Harga_Pelanggan) Then Frm87.TB5 = Format(rs!Harga_Pelanggan, "0.00") 'Harga Semasa Bagi Pelanggan (RM/g)
            ElseIf Frm87_LM_KATEGORI = 2 Then
                If IsNumeric(rs!Harga_Member) Then Frm87.TB5 = Format(rs!Harga_Member, "0.00") 'Harga Semasa Bagi Member (RM/g)
            ElseIf Frm87_LM_KATEGORI = 3 Then
                If IsNumeric(rs!Harga_Pengedar) Then Frm87.TB5 = Format(rs!Harga_Pengedar, "0.00") 'Harga Semasa Bagi Pengedar (RM/g)
            ElseIf Frm87_LM_KATEGORI = 4 Then
                If IsNumeric(rs!Harga_RAF) Then Frm87.TB5 = Format(rs!Harga_RAF, "0.00") 'Harga Semasa Bagi RAF (RM/g)
            ElseIf Frm87_LM_KATEGORI = 5 Then
                If IsNumeric(rs!harga_nd) Then Frm87.TB5 = Format(rs!harga_nd, "0.00") 'Harga Semasa Bagi Normal Dealer (RM/g)
            ElseIf Frm87_LM_KATEGORI = 6 Then
                If IsNumeric(rs!harga_md) Then Frm87.TB5 = Format(rs!harga_md, "0.00") 'Harga Semasa Bagi Master Dealer (RM/g)
            End If
        End If
        
        rs.Close
        Set rs = Nothing
    End If
End If

Frm87.TB8 = vbNullString
'###Periksa Data Produk### - End
End Sub
Sub Frm87_Senarai_Ansuran_Header()
'on error resume next
Frm87.MSFlexGrid1.Clear
Frm87.MSFlexGrid1.RowHeight(0) = 650
Frm87.MSFlexGrid1.FormatString = "No.|<No.|<ID|<Tarikh|<Jenis Ansuran|<Status|<Nama|<No. Kad Pengenalan|<No. Telefon|<No. Siri|<Kategori Produk|<Berat Asal (g)|<Berat Jualan (g)|<Harga Semasa (RM/g)|<Upah (RM)|<Harga Asal (RM)|<Adjustment (RM)|<Harga Jualan (RM)|<Kategori Pembeli"

Frm87.MSFlexGrid1.Rows = 1
Frm87.MSFlexGrid1.ColWidth(0) = 600
Frm87.MSFlexGrid1.ColWidth(1) = 0
Frm87.MSFlexGrid1.ColWidth(2) = 0 'No. ID
Frm87.MSFlexGrid1.ColWidth(3) = 1000 'Tarikh
Frm87.MSFlexGrid1.ColWidth(4) = 1100 'Jenis Ansuran
Frm87.MSFlexGrid1.ColWidth(5) = 1100 'Status
Frm87.MSFlexGrid1.ColWidth(6) = 4000 'Nama
Frm87.MSFlexGrid1.ColWidth(7) = 1300 'No. Kad Pengenalan
Frm87.MSFlexGrid1.ColWidth(8) = 1300 'No. Telefon
Frm87.MSFlexGrid1.ColWidth(9) = 1500 'No. Siri
Frm87.MSFlexGrid1.ColWidth(10) = 2400 'Kategori Produk
Frm87.MSFlexGrid1.ColWidth(11) = 1000 'Berat Asal (g)
Frm87.MSFlexGrid1.ColWidth(12) = 1000 'Berat Jualan (g)
Frm87.MSFlexGrid1.ColWidth(13) = 1000 'Harga Senasa (RM/g)
Frm87.MSFlexGrid1.ColWidth(14) = 1000 'Upah (RM)
Frm87.MSFlexGrid1.ColWidth(15) = 1000 'Harga Asal (RM)
Frm87.MSFlexGrid1.ColWidth(16) = 1200 'Adjustment (RM)
Frm87.MSFlexGrid1.ColWidth(17) = 1200 'Harga Jualan (RM)
Frm87.MSFlexGrid1.ColWidth(18) = 1200 'Kategori Pembeli
End Sub
Sub Frm87_Senarai_Ansuran()
'on error resume next
Dim Frm87_LM_BERAT As Double

x = 0
Frm87_LM_BERAT = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 27_senarai_ansuran", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm87.MSFlexGrid1.Rows = x + 1
    Frm87.MSFlexGrid1.TextMatrix(x, 0) = x 'No.
    Frm87.MSFlexGrid1.TextMatrix(x, 1) = x 'No.
    If Not IsNull(rs!ID) Then Frm87.MSFlexGrid1.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm87.MSFlexGrid1.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!jenis_ansuran) Then 'Jenis Ansuran
        If rs!jenis_ansuran = 0 Then
            Frm87.MSFlexGrid1.TextMatrix(x, 4) = "Harga Semasa"
        ElseIf rs!jenis_ansuran = 1 Then
            Frm87.MSFlexGrid1.TextMatrix(x, 4) = "Harga Tetap"
        End If
    End If
    If Not IsNull(rs!Status) Then Frm87.MSFlexGrid1.TextMatrix(x, 5) = rs!Status 'Status
    If Not IsNull(rs!Nama) Then Frm87.MSFlexGrid1.TextMatrix(x, 6) = rs!Nama 'Nama
    If Not IsNull(rs!no_ic) Then Frm87.MSFlexGrid1.TextMatrix(x, 7) = rs!no_ic 'No. Kad Pengenalan
    If Not IsNull(rs!no_tel) Then Frm87.MSFlexGrid1.TextMatrix(x, 8) = rs!no_tel 'No. Telefon
    If Not IsNull(rs!no_siri_Produk) Then Frm87.MSFlexGrid1.TextMatrix(x, 9) = rs!no_siri_Produk 'No. Siri Produk
    If Not IsNull(rs!kategori_Produk) Then Frm87.MSFlexGrid1.TextMatrix(x, 10) = rs!kategori_Produk 'Kategori Produk
    If Not IsNull(rs!Berat_Asal) Then Frm87.MSFlexGrid1.TextMatrix(x, 11) = rs!Berat_Asal 'Berat Asal (g)
    If Not IsNull(rs!berat_jualan) Then
        Frm87.MSFlexGrid1.TextMatrix(x, 12) = rs!berat_jualan 'Berat Jualan (g)
        If IsNumeric(rs!berat_jualan) Then Frm87_LM_BERAT = Frm87_LM_BERAT + rs!berat_jualan
    End If
    If Not IsNull(rs!harga_Semasa) Then Frm87.MSFlexGrid1.TextMatrix(x, 13) = rs!harga_Semasa 'Harga Semasa (RM/g)
    If Not IsNull(rs!UPAH) Then Frm87.MSFlexGrid1.TextMatrix(x, 14) = rs!UPAH 'Upah (RM)
    If Not IsNull(rs!harga_asal) Then Frm87.MSFlexGrid1.TextMatrix(x, 15) = rs!harga_asal 'Harga Asal Jualan (RM)
    If Not IsNull(rs!adjustment) Then Frm87.MSFlexGrid1.TextMatrix(x, 16) = rs!adjustment 'Adjustment (RM)
    If Not IsNull(rs!harga_jualan) Then Frm87.MSFlexGrid1.TextMatrix(x, 17) = rs!harga_jualan 'Harga Jualan (RM)
    If Not IsNull(rs!kategori_pembeli) Then
        If rs!kategori_pembeli = 1 Then
            Frm87.MSFlexGrid1.TextMatrix(x, 18) = "Pelanggan"
        ElseIf rs!kategori_pembeli = 2 Then
            Frm87.MSFlexGrid1.TextMatrix(x, 18) = "Ahli"
        ElseIf rs!kategori_pembeli = 4 Then
            Frm87.MSFlexGrid1.TextMatrix(x, 18) = "Pengedar"
        ElseIf rs!kategori_pembeli = 3 Then
            Frm87.MSFlexGrid1.TextMatrix(x, 18) = "RAF"
        ElseIf rs!kategori_pembeli = 5 Then
            Frm87.MSFlexGrid1.TextMatrix(x, 18) = "Normal Dealer"
        ElseIf rs!kategori_pembeli = 6 Then
            Frm87.MSFlexGrid1.TextMatrix(x, 18) = "Master Dealer"
       End If
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm87.L34_Text = x 'Bilangan Data
Frm87.L35_Text = Format(Frm87_LM_BERAT, "0.00 g") 'Jumlah Berat (g)
End Sub
Sub Frm87_Carian_Ansuran()
'on error resume next
Dim Frm87_LM_BERAT As Double

x = 0
Frm87_LM_BERAT = 0

If Frm87.CB25 = 0 Then
    Frm87_LM_SEARCH_1 = Null
    Frm87_LM_SEARCH_1_LOGIC = "<>"
Else
    Frm87_LM_SEARCH_1 = UCase(Frm87.TB41)
    Frm87_LM_SEARCH_1_LOGIC = "="
End If

If Frm87.CB26 = 0 Then
    Frm87_LM_SEARCH_2 = Null
    Frm87_LM_SEARCH_2_LOGIC = "<>"
Else
    Frm87_LM_SEARCH_2 = UCase(Frm87.TB41)
    Frm87_LM_SEARCH_2_LOGIC = "="
End If


Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 27_senarai_ansuran where no_ic " & Frm87_LM_SEARCH_1_LOGIC & "'" & Frm87_LM_SEARCH_1 & "' AND no_siri_Produk " & Frm87_LM_SEARCH_2_LOGIC & "'" & Frm87_LM_SEARCH_2 & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm87.MSFlexGrid1.Rows = x + 1
    Frm87.MSFlexGrid1.TextMatrix(x, 0) = x 'No.
    Frm87.MSFlexGrid1.TextMatrix(x, 1) = x 'No.
    If Not IsNull(rs!ID) Then Frm87.MSFlexGrid1.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm87.MSFlexGrid1.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!jenis_ansuran) Then 'Jenis Ansuran
        If rs!jenis_ansuran = 0 Then
            Frm87.MSFlexGrid1.TextMatrix(x, 4) = "Harga Semasa"
        ElseIf rs!jenis_ansuran = 1 Then
            Frm87.MSFlexGrid1.TextMatrix(x, 4) = "Harga Tetap"
        End If
    End If
    If Not IsNull(rs!Status) Then Frm87.MSFlexGrid1.TextMatrix(x, 5) = rs!Status 'Status
    If Not IsNull(rs!Nama) Then Frm87.MSFlexGrid1.TextMatrix(x, 6) = rs!Nama 'Nama
    If Not IsNull(rs!no_ic) Then Frm87.MSFlexGrid1.TextMatrix(x, 7) = rs!no_ic 'No. Kad Pengenalan
    If Not IsNull(rs!no_tel) Then Frm87.MSFlexGrid1.TextMatrix(x, 8) = rs!no_tel 'No. Telefon
    If Not IsNull(rs!no_siri_Produk) Then Frm87.MSFlexGrid1.TextMatrix(x, 9) = rs!no_siri_Produk 'No. Siri Produk
    If Not IsNull(rs!kategori_Produk) Then Frm87.MSFlexGrid1.TextMatrix(x, 10) = rs!kategori_Produk 'Kategori Produk
    If Not IsNull(rs!Berat_Asal) Then Frm87.MSFlexGrid1.TextMatrix(x, 11) = rs!Berat_Asal 'Berat Asal (g)
    If Not IsNull(rs!berat_jualan) Then
        Frm87.MSFlexGrid1.TextMatrix(x, 12) = rs!berat_jualan 'Berat Jualan (g)
        If IsNumeric(rs!berat_jualan) Then Frm87_LM_BERAT = Frm87_LM_BERAT + rs!berat_jualan
    End If
    If Not IsNull(rs!harga_Semasa) Then Frm87.MSFlexGrid1.TextMatrix(x, 13) = rs!harga_Semasa 'Harga Semasa (RM/g)
    If Not IsNull(rs!UPAH) Then Frm87.MSFlexGrid1.TextMatrix(x, 14) = rs!UPAH 'Upah (RM)
    If Not IsNull(rs!harga_asal) Then Frm87.MSFlexGrid1.TextMatrix(x, 15) = rs!harga_asal 'Harga Asal Jualan (RM)
    If Not IsNull(rs!adjustment) Then Frm87.MSFlexGrid1.TextMatrix(x, 16) = rs!adjustment 'Adjustment (RM)
    If Not IsNull(rs!harga_jualan) Then Frm87.MSFlexGrid1.TextMatrix(x, 17) = rs!harga_jualan 'Harga Jualan (RM)
    If Not IsNull(rs!kategori_pembeli) Then
        If rs!kategori_pembeli = 1 Then
            Frm87.MSFlexGrid1.TextMatrix(x, 18) = "Pelanggan"
        ElseIf rs!kategori_pembeli = 2 Then
            Frm87.MSFlexGrid1.TextMatrix(x, 18) = "Ahli"
        ElseIf rs!kategori_pembeli = 4 Then
            Frm87.MSFlexGrid1.TextMatrix(x, 18) = "Pengedar"
        ElseIf rs!kategori_pembeli = 3 Then
            Frm87.MSFlexGrid1.TextMatrix(x, 18) = "RAF"
        ElseIf rs!kategori_pembeli = 5 Then
            Frm87.MSFlexGrid1.TextMatrix(x, 18) = "Normal Dealer"
        ElseIf rs!kategori_pembeli = 6 Then
            Frm87.MSFlexGrid1.TextMatrix(x, 18) = "Master Dealer"
       End If
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm87.L34_Text = x 'Bilangan Data
Frm87.L35_Text = Format(Frm87_LM_BERAT, "0.00 g") 'Jumlah Berat (g)
End Sub
Sub Frm87_Carian_Ansuran2()
'on error resume next
Dim Frm87_LM_BERAT As Double

x = 0
Frm87_LM_BERAT = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 27_senarai_ansuran", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If InStr(1, rs!Nama, UCase(Frm87.TB41)) <> 0 Then
        x = x + 1
        Frm87.MSFlexGrid1.Rows = x + 1
        Frm87.MSFlexGrid1.TextMatrix(x, 0) = x 'No.
        Frm87.MSFlexGrid1.TextMatrix(x, 1) = x 'No.
        If Not IsNull(rs!ID) Then Frm87.MSFlexGrid1.TextMatrix(x, 2) = rs!ID 'No. ID
        If Not IsNull(rs!tarikh) Then Frm87.MSFlexGrid1.TextMatrix(x, 3) = rs!tarikh 'Tarikh
        If Not IsNull(rs!jenis_ansuran) Then 'Jenis Ansuran
            If rs!jenis_ansuran = 0 Then
                Frm87.MSFlexGrid1.TextMatrix(x, 4) = "Harga Semasa"
            ElseIf rs!jenis_ansuran = 1 Then
                Frm87.MSFlexGrid1.TextMatrix(x, 4) = "Harga Tetap"
            End If
        End If
        If Not IsNull(rs!Status) Then Frm87.MSFlexGrid1.TextMatrix(x, 5) = rs!Status 'Status
        If Not IsNull(rs!Nama) Then Frm87.MSFlexGrid1.TextMatrix(x, 6) = rs!Nama 'Nama
        If Not IsNull(rs!no_ic) Then Frm87.MSFlexGrid1.TextMatrix(x, 7) = rs!no_ic 'No. Kad Pengenalan
        If Not IsNull(rs!no_tel) Then Frm87.MSFlexGrid1.TextMatrix(x, 8) = rs!no_tel 'No. Telefon
        If Not IsNull(rs!no_siri_Produk) Then Frm87.MSFlexGrid1.TextMatrix(x, 9) = rs!no_siri_Produk 'No. Siri Produk
        If Not IsNull(rs!kategori_Produk) Then Frm87.MSFlexGrid1.TextMatrix(x, 10) = rs!kategori_Produk 'Kategori Produk
        If Not IsNull(rs!Berat_Asal) Then Frm87.MSFlexGrid1.TextMatrix(x, 11) = rs!Berat_Asal 'Berat Asal (g)
        If Not IsNull(rs!berat_jualan) Then
            Frm87.MSFlexGrid1.TextMatrix(x, 12) = rs!berat_jualan 'Berat Jualan (g)
            If IsNumeric(rs!berat_jualan) Then Frm87_LM_BERAT = Frm87_LM_BERAT + rs!berat_jualan
        End If
        If Not IsNull(rs!harga_Semasa) Then Frm87.MSFlexGrid1.TextMatrix(x, 13) = rs!harga_Semasa 'Harga Semasa (RM/g)
        If Not IsNull(rs!UPAH) Then Frm87.MSFlexGrid1.TextMatrix(x, 14) = rs!UPAH 'Upah (RM)
        If Not IsNull(rs!harga_asal) Then Frm87.MSFlexGrid1.TextMatrix(x, 15) = rs!harga_asal 'Harga Asal Jualan (RM)
        If Not IsNull(rs!adjustment) Then Frm87.MSFlexGrid1.TextMatrix(x, 16) = rs!adjustment 'Adjustment (RM)
        If Not IsNull(rs!harga_jualan) Then Frm87.MSFlexGrid1.TextMatrix(x, 17) = rs!harga_jualan 'Harga Jualan (RM)
        If Not IsNull(rs!kategori_pembeli) Then
            If rs!kategori_pembeli = 1 Then
                Frm87.MSFlexGrid1.TextMatrix(x, 18) = "Pelanggan"
            ElseIf rs!kategori_pembeli = 2 Then
                Frm87.MSFlexGrid1.TextMatrix(x, 18) = "Ahli"
            ElseIf rs!kategori_pembeli = 4 Then
                Frm87.MSFlexGrid1.TextMatrix(x, 18) = "Pengedar"
            ElseIf rs!kategori_pembeli = 3 Then
                Frm87.MSFlexGrid1.TextMatrix(x, 18) = "RAF"
            ElseIf rs!kategori_pembeli = 5 Then
                Frm87.MSFlexGrid1.TextMatrix(x, 18) = "Normal Dealer"
            ElseIf rs!kategori_pembeli = 6 Then
                Frm87.MSFlexGrid1.TextMatrix(x, 18) = "Master Dealer"
            End If
        End If
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm87.L34_Text = x 'Bilangan Data
Frm87.L35_Text = Format(Frm87_LM_BERAT, "0.00 g") 'Jumlah Berat (g)
End Sub
Sub Frm87_Rekod_Ansuran_Header()
'on error resume next
Frm87.MSFlexGrid2.Clear
Frm87.MSFlexGrid2.RowHeight(0) = 800
Frm87.MSFlexGrid2.FormatString = "No.|<No.|<ID|<Tarikh|<No. Resit|<Jumlah Ansuran (RM)|<Berat Diperolehi (g)|<Jenis GST (Ansuran)|<Jumlah Upah (RM)|<Jenis GST (Upah)|<Jumlah Keseluruhan (RM)|<Jumlah GST (RM)|<Adjustment (RM)|<Jumlah Bayaran (RM)"

Frm87.MSFlexGrid2.Rows = 1
Frm87.MSFlexGrid2.ColWidth(0) = 600
Frm87.MSFlexGrid2.ColWidth(1) = 0
Frm87.MSFlexGrid2.ColWidth(2) = 0 'No. ID
Frm87.MSFlexGrid2.ColWidth(3) = 1400 'Tarikh
Frm87.MSFlexGrid2.ColWidth(4) = 1400 'No. Resit
Frm87.MSFlexGrid2.ColWidth(5) = 1400 'Jumlah Ansuran (RM)
Frm87.MSFlexGrid2.ColWidth(6) = 1400 'Berat Diperolehi
Frm87.MSFlexGrid2.ColWidth(7) = 800 'Jenis GST (Ansuran)
Frm87.MSFlexGrid2.ColWidth(8) = 1400 'Jumlah Upah (RM)
Frm87.MSFlexGrid2.ColWidth(9) = 800 'Jenis GST Upah
Frm87.MSFlexGrid2.ColWidth(10) = 1400 'Jumlah Keseluruhan (RM)
Frm87.MSFlexGrid2.ColWidth(11) = 1400 'Jumlah GST (RM)
Frm87.MSFlexGrid2.ColWidth(12) = 1400 'Adjustment (RM)
Frm87.MSFlexGrid2.ColWidth(13) = 1400 'Jumlah Bayaran (RM)
End Sub
Sub Frm87_Rekod_Ansuran()
'on error resume next
x = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 28_rekod_ansuran where id_database_reg='" & Frm87.L18_Text & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm87.MSFlexGrid2.Rows = x + 1
    Frm87.MSFlexGrid2.TextMatrix(x, 0) = x 'No.
    Frm87.MSFlexGrid2.TextMatrix(x, 1) = x 'No.
    If Not IsNull(rs!ID) Then Frm87.MSFlexGrid2.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm87.MSFlexGrid2.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!no_resit_ansuran) Then Frm87.MSFlexGrid2.TextMatrix(x, 4) = rs!no_resit_ansuran 'No. Resit Ansuran
    If Not IsNull(rs!jumlah_ansuran) Then Frm87.MSFlexGrid2.TextMatrix(x, 5) = rs!jumlah_ansuran 'Jumlah Ansuran (RM)
    If Not IsNull(rs!berat_diperoleh) Then Frm87.MSFlexGrid2.TextMatrix(x, 6) = rs!berat_diperoleh 'Berat Diperolehi (g)
    If Not IsNull(rs!flag_ansuran_zr) Then 'Jenis GST (Ansuran)
        If rs!flag_ansuran_zr = 1 Then
            Frm87.MSFlexGrid2.TextMatrix(x, 7) = "ZR (L)"
        End If
    End If
    If Not IsNull(rs!flag_ansuran_sr) Then 'Jenis GST (Ansuran)
        If rs!flag_ansuran_sr = 1 Then
            Frm87.MSFlexGrid2.TextMatrix(x, 7) = "SR"
        End If
    End If
    If Not IsNull(rs!JUMLAH_UPAH) Then Frm87.MSFlexGrid2.TextMatrix(x, 8) = rs!JUMLAH_UPAH 'Jumlah Upah (RM)
    If Not IsNull(rs!flag_upah_zr) Then 'Jenis GST (Upah)
        If rs!flag_upah_zr = 1 Then
            Frm87.MSFlexGrid2.TextMatrix(x, 9) = "ZR (L)"
        End If
    End If
    If Not IsNull(rs!flag_upah_sr) Then 'Jenis GST (Upah)
        If rs!flag_upah_sr = 1 Then
            Frm87.MSFlexGrid2.TextMatrix(x, 9) = "SR"
        End If
    End If
    If Not IsNull(rs!jumlah_bayaran) Then Frm87.MSFlexGrid2.TextMatrix(x, 10) = rs!jumlah_bayaran 'Jumlah Keseluruhan (RM)
    If Not IsNull(rs!jumlah_gst) Then Frm87.MSFlexGrid2.TextMatrix(x, 11) = rs!jumlah_gst 'Jumlah GST (RM)
    If Not IsNull(rs!adjustment) Then Frm87.MSFlexGrid2.TextMatrix(x, 12) = rs!adjustment 'Adjustment (RM)
    If Not IsNull(rs!jumlah_keseluruhan) Then Frm87.MSFlexGrid2.TextMatrix(x, 13) = rs!jumlah_keseluruhan 'Jumlah keseluruhan (RM)
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
End Sub
Sub Frm87_LM_Detail_GST()
'on error resume next
Dim Frm87_LM_JUMLAH As Double
Dim Frm87_LM_GST As Double

Frm87_LM_JUMLAH = 0
Frm87_LM_GST = 0

If Frm87.TB42 <> vbNullString And IsNumeric(Frm87.TB42) Then
    Frm87_LM_JUMLAH = Frm87.TB42 'Jumlah Ansuran
End If
If Frm87.TB13 <> vbNullString And IsNumeric(Frm87.TB13) Then
    Frm87_LM_GST = Frm87.TB13 'Jumlah GST
End If

If Frm87.CB22 = 1 And Frm87.CB23 = 0 Then

    Frm87.L25_Text = Format(Frm87_LM_JUMLAH, "#,##0.00") 'Jumlah ZR
    Frm87.L26_Text = Format(Frm87_LM_GST, "#,##0.00") 'Jumlah GST ZR
    Frm87.L23_Text = Format(0, "0.00") 'Jumlah SR
    Frm87.L24_Text = Format(0, "0.00") 'Jumlah GST SR
    
End If

If Frm87.CB23 = 1 And Frm87.CB22 = 0 Then

    Frm87.L23_Text = Format(Frm87_LM_JUMLAH, "#,##0.00") 'Jumlah ZR
    Frm87.L24_Text = Format(Frm87_LM_GST, "#,##0.00") 'Jumlah GST ZR
    Frm87.L25_Text = Format(0, "0.00") 'Jumlah SR
    Frm87.L26_Text = Format(0, "0.00") 'Jumlah GST SR
    
End If

If Frm87.CB23 = 0 And Frm87.CB22 = 0 Then

    Frm87.L23_Text = Format(0, "#,##0.00") 'Jumlah ZR
    Frm87.L24_Text = Format(0, "#,##0.00") 'Jumlah GST ZR
    Frm87.L25_Text = Format(0, "0.00") 'Jumlah SR
    Frm87.L26_Text = Format(0, "0.00") 'Jumlah GST SR
    
End If
End Sub
Sub Frm87_Resit_Ansuran()
'on error resume next
DATA_FOUND = 0
Frm87_LM_KATEGORI_PEMBELI = 0 '0 : Pembeli Tidak Berdaftar , 1 : Pembeli Berdaftar , 2 : Ahli
Frm87_LM_No_PEMBELI = vbNullString

'### Reset Maklumat Pembeli ### - Start
Report40.Sections("Section2").Controls("L5").Caption = vbNullString 'Maklumat Pembeli : Nama
Report40.Sections("Section2").Controls("L7").Caption = vbNullString 'Maklumat Pembeli : No. Telefon
'### Reset Maklumat Pembeli ### - End

'### Reset maklumat kedai ### - Start
Report40.Sections("Section2").Controls("L200").Caption = vbNullString 'Nama kedai
Report40.Sections("Section2").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report40.Sections("Section2").Controls("L202").Caption = vbNullString 'Alamat
Report40.Sections("Section2").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report40.Sections("Section2").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

LM_HEADER = 1 '0 : Pre Printed , 1 : Sistem

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting where default1='" & "default" & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    
    If Not IsNull(rs!jenis_header) Then
        If rs!jenis_header = 0 Then
            LM_HEADER = 0 '0 : Pre Printed , 1 : Sistem
        ElseIf rs!jenis_header = 1 Then
            LM_HEADER = 1 '0 : Pre Printed , 1 : Sistem
        End If
    Else
        LM_HEADER = 1 '0 : Pre Printed , 1 : Sistem
    End If
    'If Not IsNull(rs!default_printer) Then LM_PRINTER = rs!default_printer
    
End If

rs.Close
Set rs = Nothing

'For Each oPrn In Printers
'    If oPrn.DeviceName = LM_PRINTER Then
'        Set Printer = oPrn
        PRINTER_FOUND = 1 '0 : Printer Not Found , 1 : Printer Found
'        Exit For
'    End If
'Next

If LM_HEADER = 1 Then '0 : Pre Printed , 1 : Sistem
    '### Maklumat kedai ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!nama_kedai) Then Report40.Sections("Section2").Controls("L200").Caption = rs!nama_kedai
        If Not IsNull(rs!no_pendaftaran) Then Report40.Sections("Section2").Controls("L201").Caption = rs!no_pendaftaran
        If Not IsNull(rs!alamat) Then Report40.Sections("Section2").Controls("L202").Caption = rs!alamat
        If Not IsNull(rs!no_tel) Then Report40.Sections("Section2").Controls("L203").Caption = rs!no_tel
        If Not IsNull(rs!no_id_gst) Then Report40.Sections("Section2").Controls("L204").Caption = rs!no_id_gst
        If Not IsNull(rs!gst_ari_nashi) Then
            If rs!gst_ari_nashi = 0 Then
                Report40.Sections("Section2").Controls("L205").Caption = "INVOICE"
            ElseIf rs!gst_ari_nashi = 1 Then
                Report40.Sections("Section2").Controls("L205").Caption = "TAX INVOICE"
            End If
        Else
            Report40.Sections("Section2").Controls("L205").Caption = "INVOICE"
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    '### Maklumat kedai ### - End
    
    Report40.Sections("Section4").Visible = True
Else
    Report40.Sections("Section4").Visible = False
End If

Report40.Sections("Section2").Controls("L3").Caption = G_No_RESIT_ANSURAN 'No. Resit Ansuran

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 28_rekod_ansuran where no_resit_ansuran='" & G_No_RESIT_ANSURAN & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!id_database_reg) Then Frm87_LM_ID = rs!id_database_reg
    If Not IsNull(rs!tarikh) Then
        Report40.Sections("Section2").Controls("L4").Caption = rs!tarikh 'Tarikh
        Frm87_LM_TARIKH = rs!tarikh 'Tarikh
    End If
    If Not IsNull(rs!jumlah_ansuran) Then
        Report40.Sections("Section1").Controls("L10").Caption = rs!jumlah_ansuran 'Jumlah Ansuran (RM)
    Else
        Report40.Sections("Section1").Controls("L10").Caption = "-"
    End If
    If Not IsNull(rs!harga_Semasa) Then
        Report40.Sections("Section1").Controls("L11").Caption = rs!harga_Semasa 'Harga Semasa (RM/g)
    Else
        Report40.Sections("Section1").Controls("L11").Caption = "-"
    End If
    If Not IsNull(rs!berat_diperoleh) Then
        Report40.Sections("Section1").Controls("L12").Caption = rs!berat_diperoleh 'Berat Diperolehi (g)
    Else
        Report40.Sections("Section1").Controls("L12").Caption = "-"
    End If
    If Not IsNull(rs!flag_ansuran_zr) Then
        If rs!flag_ansuran_zr = 1 Then
            Report40.Sections("Section1").Controls("L13").Caption = "ZR(L)" 'Jenis GST (ZR)
        End If
    End If
    If Not IsNull(rs!flag_ansuran_sr) Then
        If rs!flag_ansuran_sr = 1 Then
            Report40.Sections("Section1").Controls("L13").Caption = "SR" 'Jenis GST (SR)
        End If
    End If
    If IsNull(rs!flag_ansuran_zr) And IsNull(rs!flag_ansuran_sr) Then
        Report40.Sections("Section1").Controls("L13").Caption = "-"
    End If

    If Not IsNull(rs!JUMLAH_UPAH) Then
        Report40.Sections("Section1").Controls("L14").Caption = rs!JUMLAH_UPAH 'Upah (RM)
    Else
        Report40.Sections("Section1").Controls("L14").Caption = "-"
    End If
    If Not IsNull(rs!flag_upah_zr) Then
        If rs!flag_upah_zr = 1 Then
            Report40.Sections("Section1").Controls("L15").Caption = "ZR(L)" 'Jenis GST (ZR)
        End If
    End If
    If Not IsNull(rs!flag_upah_sr) Then
        If rs!flag_upah_sr = 1 Then
            Report40.Sections("Section1").Controls("L15").Caption = "SR" 'Jenis GST (SR)
        End If
    End If
    If IsNull(rs!flag_upah_zr) And IsNull(rs!flag_upah_sr) Then
        Report40.Sections("Section1").Controls("L15").Caption = "-"
    End If
    If Not IsNull(rs!jumlah_bayaran) Then
        Report40.Sections("Section1").Controls("L21").Caption = rs!jumlah_bayaran 'Jumlah Bayaran Ansuran + Upah (Tanpa GST)
    Else
        Report40.Sections("Section1").Controls("L21").Caption = "-"
    End If
    If Not IsNull(rs!jumlah_asal) Then
        Report40.Sections("Section1").Controls("L22").Caption = rs!jumlah_asal 'Jumlah Bayaran Ansuran + Upah (Dengan GST)
    Else
        Report40.Sections("Section1").Controls("L22").Caption = "-"
    End If
    If Not IsNull(rs!adjustment) Then
        Report40.Sections("Section1").Controls("L27").Caption = rs!adjustment 'Adjustment (RM)
    Else
        Report40.Sections("Section1").Controls("L27").Caption = "-"
    End If
    If Not IsNull(rs!jumlah_keseluruhan) Then
        Report40.Sections("Section1").Controls("L25").Caption = rs!jumlah_keseluruhan 'Jumlah Keseluruhan (RM)
    Else
        Report40.Sections("Section1").Controls("L25").Caption = "-"
    End If
    If Not IsNull(rs!no_rujukan_pekerja) Then Frm87_LM_No_PEKERJA = rs!no_rujukan_pekerja 'No. Pekerja
End If

rs.Close
Set rs = Nothing

'### Maklumat Tambahan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 27_senarai_ansuran where ID='" & Frm87_LM_ID & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!no_rujukan_pelanggan) Then
        Frm87_LM_No_PEMBELI = rs!no_rujukan_pelanggan 'No. Rujukan Pembeli
    Else
        
        If Not IsNull(rs!Nama) Then Report40.Sections("Section2").Controls("L5").Caption = rs!Nama 'Maklumat Pembeli : Nama
        If Not IsNull(rs!no_tel) Then Report40.Sections("Section2").Controls("L7").Caption = rs!no_tel 'Maklumat Pembeli : No. Telefon
        
    End If
    If Not IsNull(rs!tarikh) Then
        'Report40.Sections("Section2").Controls("L4").Caption = rs!tarikh 'Tarikh
        'Frm87_LM_TARIKH = rs!tarikh 'Tarikh
    End If
    If Not IsNull(rs!Status) Then
        Frm87_LM_STATUS = rs!Status 'Status
    End If
    If Not IsNull(rs!no_siri_Produk) Then Report40.Sections("Section2").Controls("L23").Caption = rs!no_siri_Produk 'No. Siri Produk
    If Not IsNull(rs!kategori_Produk) Then Report40.Sections("Section2").Controls("L24").Caption = rs!kategori_Produk 'Kategori Produk
End If

rs.Close
Set rs = Nothing
'### Maklumat Tambahan ### - End


'### Maklumat Terperinci GST ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 29_akaun_ansuran where no_resit='" & G_No_RESIT_ANSURAN & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!gst_sr_harga) Then Report40.Sections("Section1").Controls("L17").Caption = rs!gst_sr_harga 'Jumlah Harga SR
    If Not IsNull(rs!gst_sr_cukai) Then Report40.Sections("Section1").Controls("L18").Caption = rs!gst_sr_cukai 'Jumlah Cukai SR
    If Not IsNull(rs!gst_zr_harga) Then Report40.Sections("Section1").Controls("L19").Caption = rs!gst_zr_harga 'Jumlah Harga ZR
    If Not IsNull(rs!gst_zr_cukai) Then Report40.Sections("Section1").Controls("L20").Caption = rs!gst_zr_cukai 'Jumlah Cukai ZR
End If

rs.Close
Set rs = Nothing
'### Maklumat Terperinci GST ### - End


'### Nama Pekerja ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where NoPekerja='" & Frm87_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!Samaran) Then Report40.Sections("Section2").Controls("L26").Caption = rs!Samaran 'Nama Samaran
End If

rs.Close
Set rs = Nothing
'### Nama Pekerja ### - End

Report40.Sections("Section1").Controls("L28").Caption = "*** Status belian secara ansuran ini setakat " & Frm87_LM_TARIKH & " adalah [" & Frm87_LM_STATUS & "]"

'### Maklumat Pembeli ### - Start
If Frm87_LM_No_PEMBELI <> vbNullString Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm87_LM_No_PEMBELI & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!Nama) Then Report40.Sections("Section2").Controls("L5").Caption = rs!Nama 'Maklumat Pembeli : Nama
        If Not IsNull(rs!no_tel) Then Report40.Sections("Section2").Controls("L7").Caption = rs!no_tel 'Maklumat Pembeli : No. Telefon
    End If
    
    rs.Close
    Set rs = Nothing
    
ElseIf Frm87_LM_No_PEMBELI = vbNullString Then

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 44_senarai_pelanggan where no_resit='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!Nama) Then Report40.Sections("Section2").Controls("L5").Caption = rs!Nama 'Maklumat Pembeli : Nama
            If Not IsNull(rs!no_tel) Then Report40.Sections("Section2").Controls("L7").Caption = rs!no_tel 'Maklumat Pembeli : No. Telefon
        End If
        
        rs.Close
        Set rs = Nothing
    
End If
'### Maklumat Pembeli ### - End

'### Paparan Resit Ansuran ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 28_rekod_ansuran where no_resit_ansuran='" & G_No_RESIT_ANSURAN & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report40.DataSource = rs
    Report40.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing
'### Paparan Resit Ansuran ### - End
    
G_No_RESIT_ANSURAN = vbNullString

End Sub
