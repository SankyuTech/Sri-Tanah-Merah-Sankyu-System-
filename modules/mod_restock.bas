Attribute VB_Name = "mod_restock"
Sub Frm104_initial_setting()
'on error resume next
Frm104.Pic1.Left = 120
Frm104.Pic1.Top = 240

Frm104.Pic4.Left = 120
Frm104.Pic4.Top = 240

Frm104.Pic1.Visible = False

Frm104.Pic4.Visible = False
End Sub
Sub Frm104_reset_component()
'on error resume next
Frm104.L12_Text = 0 'Summary : Bil. Berat Terjual
Frm104.L13_Text = "0.00" 'Summary : Jumlah Berat Terjual
Frm104.L14_Text = "0.00" 'Summary : Harga Jualan (Dengan GST)
Frm104.L21_Text = "0.00" 'Summary : Harga Jualan (Tanpa GST)
Frm104.L15_Text = "0.00" 'Summary : Jumlah Upah Jualan
Frm104.L16_Text = "0.00" 'Summary : Modal Upah
Frm104.L17_Text = "0.00" 'Summary : Jumlah Harga Restock

Frm104.TB1 = "0.00" 'Berat restock
Frm104.TB2 = "0.00" 'Harga semasa (dari supplier)
Frm104.TB3 = "0.00" 'Upah (Tanpa GST)
End Sub
Sub Frm104_transfer_list_jualan()
'on error resume next
Dim TM As Date
Dim TA As Date

Frm104.L5_Text = Frm104.DTPicker1 'Tarikh Mula
Frm104.L6_Text = Frm104.DTPicker2 'Tarikh Akhir
Frm104.L7_Text = Frm104.CBB1 'Purity

Frm104.L9_Text = Frm104.DTPicker1 'Tarikh Mula
Frm104.L10_Text = Frm104.DTPicker2 'Tarikh Akhir
Frm104.L11_Text = Frm104.CBB1 'Purity

TM = Frm104.L5_Text
TA = Frm104.L6_Text
Frm104_LM_PURITY = Frm104.L7_Text 'Purity

Frm104_LM_STATUS = 1 '0 : Tidak termasuk dalam pengiraan , 1 : Termasuk dalam pengiraan

'Kosongkan table #54_restock_list_jualan bagi diisi dengan data yang baru mengikut data jualan yang baru
'=======================================================================================================
'###Padam Table #54_restock_list_jualan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "TRUNCATE TABLE 54_restock_list_jualan"

Set rs = cn.Execute(strsql)
Set rs = Nothing
'###Padam Table #54_restock_list_jualan ### - End

'Masukkan senarai jualan barang kemas dari menu ANSURAN ke dalam table #54_restock_list_jualan
'Hanya senarai BARANG KEMAS yang dimasukkan ke dalam senarai ini (@jenis_produk = 0)
'================================================================
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "insert into 54_restock_list_jualan(no_id_ansuran,tarikh,no_siri_produk,kategori_produk,purity,berat_asal,berat_jualan,upah,kategori_jualan)" & _
            "select ID,tarikh_jelas,no_siri_produk,kategori_produk,purity,berat_asal,berat_jualan,upah,1 from 27_senarai_ansuran WHERE jenis_produk='" & "0" & "' AND purity='" & Frm104_LM_PURITY & "' AND tarikh_jelas BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh_jelas ASC"

Set rs = cn.Execute(strsql)
Set rs = Nothing

'Masukkan jumlah harga jualan mengikut maklumat di bawah (Untuk barangan yang dijual secara ANSURAN)
'Harga jualan dengan GST
'Harga jualan tanpa GST
'Jumlah GST
'=========================================================
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "UPDATE 54_restock_list_jualan,28_rekod_ansuran SET 54_restock_list_jualan.harga_dengan_gst = (SELECT SUM(28_rekod_ansuran.jumlah_keseluruhan) FROM 28_rekod_ansuran WHERE 28_rekod_ansuran.id_database_reg = 54_restock_list_jualan.no_id_ansuran AND 54_restock_list_jualan.kategori_jualan = 1) ," _
& "54_restock_list_jualan.jumlah_gst = (SELECT SUM(28_rekod_ansuran.jumlah_gst) FROM 28_rekod_ansuran WHERE 28_rekod_ansuran.id_database_reg = 54_restock_list_jualan.no_id_ansuran AND 54_restock_list_jualan.kategori_jualan = 1) ," _
& "54_restock_list_jualan.harga_tanpa_gst = (SELECT SUM(28_rekod_ansuran.jumlah_keseluruhan - 28_rekod_ansuran.jumlah_gst) FROM 28_rekod_ansuran WHERE 28_rekod_ansuran.id_database_reg = 54_restock_list_jualan.no_id_ansuran AND 54_restock_list_jualan.kategori_jualan = 1) WHERE 54_restock_list_jualan.kategori_jualan = 1"

Set rs = cn.Execute(strsql)
Set rs = Nothing

'Masukkan senarai jualan barang kemas dari menu TEMPAHAN ke dalam table #54_restock_list_jualan
'Hanya senarai BARANG KEMAS yang dimasukkan ke dalam senarai ini
'================================================================
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "insert into 54_restock_list_jualan(tarikh,no_siri_produk,no_resit,no_rujukan_tempahan,kategori_produk,purity,berat_asal,berat_jualan,harga_semasa,upah,harga_asal,adjustment,harga_jualan,kategori_jualan)" & _
            "select tarikh,no_siri_produk,no_resit_tempahan,no_rujukan_tempahan,kategori_produk,purity,berat_asal,berat_jualan,harga_semasa,upah,harga_asal,adjustment,harga,2 from 42_tempahan_siap WHERE type_barang_kemas='" & "0" & "' AND purity='" & Frm104_LM_PURITY & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC"

Set rs = cn.Execute(strsql)
Set rs = Nothing

'Masukkan jumlah harga jualan mengikut maklumat di bawah (Untuk barangan yang dijual secara TEMPAHAN)
'Harga jualan dengan GST
'Harga jualan tanpa GST
'Jumlah GST
'=========================================================
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "UPDATE 54_restock_list_jualan,41_akaun_tempahan SET 54_restock_list_jualan.harga_dengan_gst = (SELECT SUM(41_akaun_tempahan.jumlah_deposit_dengan_gst) FROM 41_akaun_tempahan WHERE 41_akaun_tempahan.no_rujukan_tempahan = 54_restock_list_jualan.no_rujukan_tempahan AND 54_restock_list_jualan.kategori_jualan = 2) ," _
& "54_restock_list_jualan.jumlah_gst = (SELECT SUM(41_akaun_tempahan.jumlah_gst) FROM 41_akaun_tempahan WHERE 41_akaun_tempahan.no_rujukan_tempahan = 54_restock_list_jualan.no_rujukan_tempahan AND 54_restock_list_jualan.kategori_jualan = 2) ," _
& "54_restock_list_jualan.harga_tanpa_gst = (SELECT SUM(41_akaun_tempahan.jumlah_deposit_dengan_gst - 41_akaun_tempahan.jumlah_gst) FROM 41_akaun_tempahan WHERE 41_akaun_tempahan.no_rujukan_tempahan = 54_restock_list_jualan.no_rujukan_tempahan AND 54_restock_list_jualan.kategori_jualan = 2) WHERE 54_restock_list_jualan.kategori_jualan = 2"

Set rs = cn.Execute(strsql)
Set rs = Nothing

'Masukkan senarai jualan ke dalam table #54_restock_list_jualan
'Hanya senarai BARANG KEMAS yang dimasukkan ke dalam senarai ini
'================================================================

'Masukkan senarai barang yang dijual dari jualan terus kepada pelanggan (jenis_jualan=0)
'=======================================================================================
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "insert into 54_restock_list_jualan(tarikh,no_resit,no_siri_produk,kategori_produk,purity,berat_jualan,harga_semasa,upah,harga_lepas_diskaun,adjustment,harga_jualan,harga_tanpa_gst,jumlah_gst,harga_dengan_gst,jenis_jualan,berat_asal,harga_asal,diskaun)" & _
            "select tarikh,no_resit,no_siri_produk,kategori_produk,purity,berat_jualan,harga_semasa,upah,harga_lepas_diskaun,adjustment,harga_jualan,harga_tanpa_gst,jumlah_gst,harga_dengan_gst,jenis_jualan,berat_asal,harga_asal,diskaun from 23_senarai_jualan WHERE purity='" & Frm104_LM_PURITY & "' AND status_rekod = 1 AND type='" & "0" & "' AND gst_barang_atau_upah='" & "0" & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC"

Set rs = cn.Execute(strsql)
Set rs = Nothing

'Masukkan senarai barang yang dijual dari jualan terus kepada pelanggan (jenis_jualan=1)
'Perbezaan adalah pada "jumlah harga tanpa gst"
'=======================================================================================
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "insert into 54_restock_list_jualan(tarikh,no_resit,no_siri_produk,kategori_produk,purity,berat_jualan,harga_semasa,upah,harga_lepas_diskaun,adjustment,harga_jualan,harga_tanpa_gst,jumlah_gst,harga_dengan_gst,jenis_jualan,berat_asal,harga_asal,diskaun)" & _
            "select tarikh,no_resit,no_siri_produk,kategori_produk,purity,berat_jualan,harga_semasa,upah,harga_lepas_diskaun,adjustment,harga_jualan,harga_jualan_dengan_gst-jumlah_gst,jumlah_gst,harga_dengan_gst,jenis_jualan,berat_asal,harga_asal,diskaun from 23_senarai_jualan WHERE purity='" & Frm104_LM_PURITY & "' AND status_rekod = 1 AND type='" & "0" & "' AND gst_barang_atau_upah='" & "1" & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC"

Set rs = cn.Execute(strsql)
Set rs = Nothing

'Masukkan upah modal bagi setiap barang yang terjual
'Hanya upah bagi gst yang dikenakan pada harga barang sahaja diupdate
'====================================================================
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "UPDATE 54_restock_list_jualan, data_database SET 54_restock_list_jualan.upah_asal = data_database.upah WHERE 54_restock_list_jualan.no_siri_produk = data_database.no_siri_produk AND data_database.gst_barang_atau_upah='" & "0" & "'"

Set rs = cn.Execute(strsql)
Set rs = Nothing

'Masukkan upah modal bagi setiap barang yang terjual
'Hanya upah bagi gst yang dikenakan pada upah sahaja diupdate
'============================================================
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "UPDATE 54_restock_list_jualan, data_database SET 54_restock_list_jualan.upah_asal = data_database.harga_tanpa_gst WHERE 54_restock_list_jualan.no_siri_produk = data_database.no_siri_produk AND data_database.gst_barang_atau_upah='" & "1" & "'"

Set rs = cn.Execute(strsql)
Set rs = Nothing

'Masukkan status setiap jualan
'0 : Tidak termasuk dalam pengiraan
'1 : Termasuk dalam pengiraan
'==================================
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "UPDATE 54_restock_list_jualan set status='" & Frm104_LM_STATUS & "'"

Set rs = cn.Execute(strsql)
Set rs = Nothing

Call Frm104_reset_component
Call Frm104_untung_rugi_restock

End Sub
Sub Frm104_untung_rugi_restock()
'on error resume next
DATA_ARI_NASHI = 0 '0 : Tiada barang yang terjual dalam tempoh report , 1 : Ada barang yang terjual dalam tempoh report

'#### Bilangan barang yang terjual (yang dikira dalam pengiraan untung rugi ini) #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(no_siri_produk) from 54_restock_list_jualan where status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then
    Frm104.L12_Text = rs(0) 'Summary : Bil. Berat Terjual
    If rs(0) <> 0 Then DATA_ARI_NASHI = 1 '0 : Tiada barang yang terjual dalam tempoh report , 1 : Ada barang yang terjual dalam tempoh report
End If

rs.Close
Set rs = Nothing
'#### Bilangan barang yang terjual (yang dikira dalam pengiraan untung rugi ini) #### - End

'#### Jumlah berat barang yang telah terjual #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(berat_jualan) from 54_restock_list_jualan where status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm104.L13_Text = Format(rs(0), "#,##0.00") 'Summary : Jumlah Berat Terjual

rs.Close
Set rs = Nothing
'#### Jumlah berat barang yang telah terjual #### - End

'#### Jumlah harga jualan terkumpul tanpa GST #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_tanpa_gst) from 54_restock_list_jualan where status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm104.L21_Text = Format(rs(0), "#,##0.00") 'Summary : Harga Jualan (Tanpa GST)

rs.Close
Set rs = Nothing
'#### Jumlah harga jualan terkumpul tanpa GST #### - End

'#### Jumlah harga jualan terkumpul dengan GST #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select SUM(harga_tanpa_gst) from 54_restock_list_jualan where status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic
rs.Open "select SUM(harga_dengan_gst) from 54_restock_list_jualan where status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm104.L14_Text = Format(rs(0), "#,##0.00") 'Summary : Harga Jualan (Tanpa GST)

rs.Close
Set rs = Nothing
'#### Jumlah harga jualan terkumpul dengan GST #### - End

'#### Jumlah upah jualan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(upah) from 54_restock_list_jualan where status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm104.L15_Text = Format(rs(0), "#,##0.00") 'Summary : Jumlah Upah Jualan

rs.Close
Set rs = Nothing
'#### Jumlah upah jualan #### - End

'#### Jumlah upah jualan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(upah_asal) from 54_restock_list_jualan where status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm104.L16_Text = Format(rs(0), "#,##0.00") 'Summary : Modal Upah

rs.Close
Set rs = Nothing
'#### Jumlah upah jualan #### - End

Frm104.L20_Text = "Report generated on " & Now

If DATA_ARI_NASHI = 1 Then '0 : Tiada barang yang terjual dalam tempoh report , 1 : Ada barang yang terjual dalam tempoh report
    Frm104.Pic1.Visible = False
    Frm104.Pic4.Visible = True
Else
    MsgBox "Tiada barang yang terjual dalam tempoh report.", vbInformation, "Info"
End If
End Sub
Sub Frm104_harga_restock()
'on error resume next
Dim Frm104_LM_BERAT As Double 'Berat Belian
Dim Frm104_LM_HARGA_SEMASA As Double 'Harga Semasa Dari Supplier
Dim Frm104_LM_UPAH As Double 'Upah Belian

Frm104_LM_BERAT = 0 'Berat Belian
Frm104_LM_HARGA_SEMASA = 0 'Harga Semasa Dari Supplier
Frm104_LM_UPAH = 0 'Upah Belian

If Frm104.TB1 <> vbNullString And IsNumeric(Frm104.TB1) Then Frm104_LM_BERAT = Frm104.TB1 'Berat Belian
If Frm104.TB2 <> vbNullString And IsNumeric(Frm104.TB2) Then Frm104_LM_HARGA_SEMASA = Frm104.TB2 'Harga Semasa Dari Supplier
If Frm104.TB3 <> vbNullString And IsNumeric(Frm104.TB3) Then Frm104_LM_UPAH = Frm104.TB3 'Upah Belian

Frm104.L17_Text = Format((Frm104_LM_BERAT * Frm104_LM_HARGA_SEMASA) + Frm104_LM_UPAH, "#,##0.00") 'Jumlah Harga Restock
End Sub
Sub Frm104_analisa_untung_rugi_restock()
'on error resume next
Dim Frm104_LM_BERAT_JUALAN As Double 'Berat jualan
Dim Frm104_LM_BERAT_BELIAN As Double 'Berat belian
Dim Frm104_LM_HARGA_JUALAN As Double 'Harga jualan
Dim Frm104_LM_HARGA_RESTOCK As Double 'Harga restock

Frm104_LM_BERAT_JUALAN = 0 'Berat jualan
Frm104_LM_BERAT_BELIAN = 0 'Berat belian
Frm104_LM_HARGA_JUALAN = 0 'Harga jualan
Frm104_LM_HARGA_RESTOCK = 0 'Harga restock

If Frm104.TB1 <> vbNullString And IsNumeric(Frm104.TB1) Then Frm104_LM_BERAT_BELIAN = Frm104.TB1 'Berat Belian
If Frm104.L13_Text <> vbNullString And IsNumeric(Frm104.L13_Text) Then Frm104_LM_BERAT_JUALAN = Frm104.L13_Text 'Berat jualan

If Frm104.L14_Text <> vbNullString And IsNumeric(Frm104.L14_Text) Then Frm104_LM_HARGA_JUALAN = Frm104.L14_Text 'Harga jualan
If Frm104.L17_Text <> vbNullString And IsNumeric(Frm104.L17_Text) Then Frm104_LM_HARGA_RESTOCK = Frm104.L17_Text 'Harga restock

If IsNumeric(Frm104.L13_Text) And IsNumeric(Frm104.TB1) Then
    Frm104_LM_BEZA_BERAT = Frm104_LM_BERAT_BELIAN - Frm104_LM_BERAT_JUALAN
    If Frm104_LM_BEZA_BERAT > 0 Then
        Frm104.L18_Text.ForeColor = &H0&
        Frm104.L18_Text = "Lebihan berat sebanyak " & Format(Frm104_LM_BEZA_BERAT, "#,##0.00 g")
    ElseIf Frm104_LM_BEZA_BERAT < 0 Then
        Frm104.L18_Text.ForeColor = &HFF&
        Frm104.L18_Text = "Susut berat sebanyak " & Format(-Frm104_LM_BEZA_BERAT, "#,##0.00 g")
    ElseIf Frm104_LM_BEZA_BERAT = 0 Then
        Frm104.L18_Text.ForeColor = &H0&
        Frm104.L18_Text = "Tiada lebihan atau susut berat."
    End If
Else
    Frm104.L18_Text.ForeColor = &H0&
    Frm104.L18_Text = "-------------------------------------------------"
End If

If IsNumeric(Frm104.L14_Text) And IsNumeric(Frm104.L17_Text) Then
    Frm104_LM_BEZA_HARGA = Frm104_LM_HARGA_JUALAN - Frm104_LM_HARGA_RESTOCK
    If Frm104_LM_BEZA_HARGA > 0 Then
        Frm104.L19_Text.ForeColor = &H0&
        Frm104.L19_Text = "Lebihan duit sebanyak RM " & Format(Frm104_LM_BEZA_HARGA, "#,##0.00")
    ElseIf Frm104_LM_BEZA_HARGA < 0 Then
        Frm104.L19_Text.ForeColor = &HFF&
        Frm104.L19_Text = "Susut duit sebanyak RM " & Format(-Frm104_LM_BEZA_HARGA, "#,##0.00")
    ElseIf Frm104_LM_BEZA_HARGA = 0 Then
        Frm104.L19_Text.ForeColor = &H0&
        Frm104.L19_Text = "Tiada lebihan atau susut duit."
    End If
Else
    Frm104.L19_Text.ForeColor = &H0&
    Frm104.L19_Text = "-------------------------------------------------"
End If

End Sub
Sub Frm104_penyata_restock()
'on error resume next

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

Report61.Sections("Section4").Controls("L1").Caption = vbNullString 'Header
Report61.Sections("Section4").Controls("L2").Caption = 0 'Bil. barang terjual
Report61.Sections("Section4").Controls("L3").Caption = "0.00" 'Jumlah berat terjual
Report61.Sections("Section4").Controls("L4").Caption = "0.00" 'Harga jualan
Report61.Sections("Section4").Controls("L5").Caption = "0.00" 'Jumlah upah jualan
Report61.Sections("Section4").Controls("L6").Caption = "0.00" 'Jumlah upah modal
Report61.Sections("Section4").Controls("L7").Caption = "0.00" 'Berat restock
Report61.Sections("Section4").Controls("L8").Caption = "0.00" 'Harga semasa
Report61.Sections("Section4").Controls("L9").Caption = "0.00" 'Upah
Report61.Sections("Section4").Controls("L10").Caption = "0.00" 'Harga restock
Report61.Sections("Section4").Controls("L11").Caption = vbNullString 'Analisa : berat
Report61.Sections("Section4").Controls("L12").Caption = vbNullString 'Analisa : duit
Report61.Sections("Section5").Controls("L13").Caption = vbNullString 'Timestamp

'### Reset maklumat kedai ### - Start
Report61.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report61.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report61.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report61.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report61.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report61.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report61.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report61.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report61.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report61.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Report61.Sections("Section4").Controls("L1").Caption = "Report analisa untung rugi restock bagi purity " & Frm104.L11_Text & " yang terjual dari " & Frm104.L9_Text & " hingga " & Frm104.L10_Text & "." 'Header

Report61.Sections("Section4").Controls("L2").Caption = Frm104.L12_Text 'Bil. barang terjual
Report61.Sections("Section4").Controls("L3").Caption = Format(Frm104.L13_Text, "#,##0.00 g") 'Jumlah berat terjual
Report61.Sections("Section4").Controls("L4").Caption = "RM " & Format(Frm104.L14_Text, "#,##0.00") 'Harga jualan
Report61.Sections("Section4").Controls("L5").Caption = "RM " & Format(Frm104.L15_Text, "#,##0.00") 'Jumlah upah jualan
Report61.Sections("Section4").Controls("L6").Caption = "RM " & Format(Frm104.L16_Text, "#,##0.00") 'Jumlah upah modal
Report61.Sections("Section4").Controls("L7").Caption = Format(Frm104.TB1, "#,##0.00 g") 'Berat restock
Report61.Sections("Section4").Controls("L8").Caption = "RM " & Format(Frm104.TB2, "#,##0.00") 'Harga semasa
Report61.Sections("Section4").Controls("L9").Caption = "RM " & Format(Frm104.TB3, "#,##0.00") 'Upah
Report61.Sections("Section4").Controls("L10").Caption = "RM " & Format(Frm104.L17_Text, "#,##0.00") 'Harga restock
Report61.Sections("Section4").Controls("L11").Caption = Frm104.L18_Text 'Analisa : berat
Report61.Sections("Section4").Controls("L12").Caption = Frm104.L19_Text 'Analisa : duit
Report61.Sections("Section5").Controls("L13").Caption = Now

'### Paparan Penyata ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 54_restock_list_jualan order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report61.DataSource = rs
    Report61.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing
'### Paparan Penyata ### - End

End Sub
