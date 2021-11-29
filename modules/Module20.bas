Attribute VB_Name = "Module20"
Sub frm102_reset_main2()
'on error resume next
'### Digunakan untuk reset paparan / komponen pada setiap kali penjualan barang atau pembatalan jualan
Frm102.L3_Text = vbNullString 'No. siri produk
Frm102.L4_Text = vbNullString 'Purity
Frm102.L5_Text = vbNullString 'Kategori Produk
Frm102.L6_Text = "0.00" 'Berat asal
Frm102.L7_Text = "0.00" 'Berat jualan dalam purity 999.9
Frm102.L8_Text = "0.00" 'Trade In : Berat dalam 999.9
Frm102.L9_Text = "0.00" 'Berat jualan 999.9
Frm102.L10_Text = "0.00" 'Berat belian 999.9
Frm102.L11_Text = "0.00" 'Beza berat 999.9
Frm102.L12_Text = "0.00" 'Harga emas
Frm102.L13_Text = "0.00" 'Overall : Upah + GST
Frm102.L14_Text = "0.00" 'Overall : Jumlah bayaran
Frm102.L15_Text = "0.00" 'Maklumat GST : Jumlah harga tanpa GST
Frm102.L16_Text = "0.00" 'Maklumat GST : Jumlah harga dengan GST
Frm102.L17_Text = "0.00" 'Maklumat GST : Jumlah harga ZR
Frm102.L18_Text = "0.00" 'Maklumat GST : Jumlah harga SR
Frm102.L19_Text = "0.00" 'Maklumat GST : Jumlah GST ZR
Frm102.L20_Text = "0.00" 'Maklumat GST : Jumlah GST SR
Frm102.L24_Text = 0 'No. id jualan
Frm102.L25_Text = 0 'No. id trade in
Frm102.L28_Text = "0.00" 'Jumlah simpanan duit terkumpul di kedai

Frm102.TB1 = vbNullString 'No. Siri Produk (Scan)
Frm102.TB2 = "0.00" 'Harga emas semasa 999.9 (Untuk tujuan jualan kepada pelanggan)
Frm102.TB3 = "0.00" 'Berat jualan (g)
Frm102.TB4 = "0.00" 'Upah
Frm102.TB5 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
Frm102.TB6 = "0.00" 'Jumlah Upah + GST (Bagi jualan setiap item)
Frm102.TB7 = "0.00" 'Jumlah upah tanpa GST (keseluruhan)
Frm102.TB8 = "0.00" 'Jumlah GST (Keseluruhan)
Frm102.TB9 = "0.00" 'Jumlah Upah + GST (Keseluruhan)
Frm102.TB10 = "0.00" 'Berat (barang trade in)
Frm102.TB11 = "0.00" 'Harga emas semasa 999.9 (Overall)
Frm102.TB12 = "0.00" 'Overall : Jumlah GST
Frm102.TB13 = "0.00" 'Overall : Harga emas dengan GST
Frm102.TB14 = "0.00" 'Cara bayaran : Tunai
Frm102.TB15 = "0.00" 'Cara bayaran : Bank in
Frm102.TB16 = "0.00" 'Cara bayaran : Kad kredit
Frm102.TB17 = "0.00" 'Cara bayaran : Cas kad kredit
Frm102.TB18 = "0.00" 'Cara bayaran : Potongan kad kredit
Frm102.TB19 = "0.00" 'Cara bayaran : Kad debit
Frm102.TB20 = "0.00" 'Cara bayaran : Cas kad debit
Frm102.TB21 = "0.00" 'Cara bayaran : Potongan kad debit
Frm102.TB22 = "0.00" 'Cara bayaran : Duit simpanan
Frm102.TB23 = "0.00" 'Jumlah bayaran

Frm102.CMD1.Visible = True 'Masukkan dalam senarai jualan
Frm102.CMD2.Visible = False 'Masukkan dalam senarai jualan (Edit)
Frm102.CMD3.Visible = False 'Batal edit data
Frm102.CMD4.Visible = True 'Masukkan dalam senarai trade in
Frm102.CMD5.Visible = False 'Masukkan dalam senarai trade in (Edit)
Frm102.CMD6.Visible = False 'Batal edit data
End Sub
Sub frm102_reset_1()
'on error resume next
'### Digunakan untuk reset paparan / komponen pada setiap kali penjualan barang atau pembatalan jualan

Frm102.L3_Text = vbNullString 'No. siri produk
Frm102.L4_Text = vbNullString 'Purity
Frm102.L5_Text = vbNullString 'Kategori Produk
Frm102.L6_Text = "0.00" 'Berat asal
Frm102.L7_Text = "0.00" 'Berat jualan dalam purity 999.9
Frm102.L24_Text = 0 'No. id jualan
'Frm102.L43_Text = 0 'Jumlah bilangan barang jualan
Frm102.L43_Text.BackStyle = 0 'Jumlah bilangan barang jualan
'Frm102.L48_Text = "0.00" 'Jumlah berat (g)
Frm102.L48_Text.BackStyle = 0 'Jumlah berat (g)
Frm102.L49_Text = "0.00"
Frm102.L50_Text = "0.00"

'Frm102.TB1 = vbNullString 'No. Siri Produk (Scan)
Frm102.TB3 = "0.00" 'Berat jualan (g)
Frm102.TB4 = "0.00" 'Upah
Frm102.TB5 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
Frm102.TB6 = "0.00" 'Jumlah Upah + GST (Bagi jualan setiap item)

Frm102.CMD1.Visible = True 'Masukkan dalam senarai jualan
Frm102.CMD2.Visible = False 'Masukkan dalam senarai jualan (Edit)
Frm102.CMD3.Visible = False 'Batal edit data
End Sub
Sub frm102_reset_2()
'on error resume next
'### Digunakan untuk reset paparan / komponen pada setiap kali pembelian barang trade in atau pembatalan pembelian barang trade in
Frm102.TB10 = "0.00" 'Berat (barang trade in)
Frm102.L25_Text = 0 'No. id trade in
'Frm102.L44_Text = 0 'Jumlah bilangan barang trade in
Frm102.L47_Text = vbNullString 'Kod purity barang
Frm102.L44_Text.BackStyle = 0 'Jumlah bilangan barang trade in

Frm102.L8_Text = "0.00" 'Trade In : Berat dalam 999.9

Frm102.CMD4.Visible = True 'Masukkan dalam senarai trade in
Frm102.CMD5.Visible = False 'Masukkan dalam senarai trade in (Edit)
Frm102.CMD6.Visible = False 'Batal edit data
End Sub
Sub frm102_reset_3()
'on error resume next
'### Digunakan untuk reset paparan / komponen semua komponen transaksi
Frm102.L9_Text = "0.00" 'Berat jualan 999.9
Frm102.L10_Text = "0.00" 'Berat belian 999.9
Frm102.L11_Text = "0.00" 'Beza berat 999.9
Frm102.L12_Text = "0.00" 'Harga emas
Frm102.L13_Text = "0.00" 'Overall : Upah + GST
Frm102.L14_Text = "0.00" 'Overall : Jumlah bayaran
Frm102.L15_Text = "0.00" 'Maklumat GST : Jumlah harga tanpa GST
Frm102.L16_Text = "0.00" 'Maklumat GST : Jumlah harga dengan GST
Frm102.L17_Text = "0.00" 'Maklumat GST : Jumlah harga ZR
Frm102.L18_Text = "0.00" 'Maklumat GST : Jumlah harga SR
Frm102.L19_Text = "0.00" 'Maklumat GST : Jumlah GST ZR
Frm102.L20_Text = "0.00" 'Maklumat GST : Jumlah GST SR
Frm102.L24_Text = 0 'No. id jualan
Frm102.L25_Text = 0 'No. id trade in
Frm102.L28_Text = "0.00" 'Jumlah simpanan duit terkumpul di kedai

Frm102.TB11 = "0.00" 'Harga emas semasa 999.9 (Overall)
Frm102.TB12 = "0.00" 'Overall : Jumlah GST
Frm102.TB13 = "0.00" 'Overall : Harga emas dengan GST
Frm102.TB14 = "0.00" 'Cara bayaran : Tunai
Frm102.TB15 = "0.00" 'Cara bayaran : Bank in
Frm102.TB16 = "0.00" 'Cara bayaran : Kad kredit
Frm102.TB17 = "0.00" 'Cara bayaran : Cas kad kredit
Frm102.TB18 = "0.00" 'Cara bayaran : Potongan kad kredit
Frm102.TB19 = "0.00" 'Cara bayaran : Kad debit
Frm102.TB20 = "0.00" 'Cara bayaran : Cas kad debit
Frm102.TB21 = "0.00" 'Cara bayaran : Potongan kad debit
Frm102.TB22 = "0.00" 'Cara bayaran : Duit simpanan
Frm102.TB23 = "0.00" 'Jumlah bayaran
End Sub
Sub frm102_reset_main()
'on error resume next
'### Digunakan untuk reset / update komponen dari database setelah penjualan atau pembatalan jualan
Frm102.TB24 = "0.00" 'Trade In : Kadar tukaran purity 999.9
Frm102.TB2 = "0.00" 'Harga emas semasa 999.9 (Untuk tujuan jualan kepada pelanggan)
Frm102.TB7 = "0.00" 'Jumlah upah tanpa GST (keseluruhan)
Frm102.TB8 = "0.00" 'Jumlah GST (Keseluruhan)
Frm102.TB9 = "0.00" 'Jumlah Upah + GST (Keseluruhan)
Frm102.L34_Text = "Jumlah ini adalah nilai yang perlu dibayar oleh pihak kedai kepada agen."
Frm102.L34_Text.BackStyle = 0
Frm102.L34_Text.Visible = False
Frm102.Pic1.Visible = False
Frm102.Pic1.Left = 120
Frm102.Pic1.Top = 6360
Frm102.L45_Text = 0 'Flag bagi jika ada pengeluaran voucher bagi urusan ini , 0 : Tiada voucher / Tiada history pengeluaran voucher , 1 : Ada voucher / Ada history pengeluaran voucher

Frm102.L35_Text = 0
Frm102.L36_Text = 0
Frm102.L37_Text = 0
Frm102.L38_Text = 0
Frm102.L39_Text = 0
Frm102.L40_Text = 0
Frm102.L41_Text = 0
Frm102.L42_Text = 0
Frm102.L46_Text = vbNullString

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Default1 = "Default" Then
        GLOBAL_DISABLE = 1
        If Not IsNull(rs!ScannerMode) Then 'Tetapan penggunaan scanner
            If rs!ScannerMode = 1 Then
                Frm102.CB1 = 1
            Else
                Frm102.CB1 = 0
            End If
        Else
            Frm102.CB1 = 0
        End If
        If Not IsNull(rs!gst_value) Then Frm102.L21_Text = rs!gst_value 'Jumlah Kadar GST
        If Not IsNull(rs!gst_arinashi) Then 'Tetapan GST , ZR atau SR
            If rs!gst_arinashi = 1 Then 'SR
                Frm102.CB3 = 1
                Frm102.CB6 = 1
                Frm102.CB2 = 0
                Frm102.CB5 = 0
            Else 'ZR
                Frm102.CB2 = 1
                Frm102.CB5 = 1
                Frm102.CB3 = 0
                Frm102.CB6 = 0
            End If
        End If
        If Not IsNull(rs!gst_jualan_included) Then
            If rs!gst_jualan_included = 1 Then
                Frm102.CB4 = 1
                Frm102.CB7 = 1
            ElseIf rs!gst_jualan_included = 0 Then
                Frm102.CB5 = 0
                Frm102.CB7 = 0
            End If
        Else
            Frm102.CB5 = 0
            Frm102.CB7 = 0
        End If
        
        If Not IsNull(rs!NoRujukanSistem) Then Frm102.L29_Text = rs!NoRujukanSistem 'No. Rujukan Sistem
        If Not IsNull(rs!no_trade_in_agen) Then Frm102.L22_Text = rs!no_trade_in_agen 'No. Voucher Trade In
        If Not IsNull(rs!ResitNo) Then Frm102.L23_Text = rs!ResitNo 'No. Invoice
        If Not IsNull(rs!cas_Kad_Kredit) Then Frm102.L26_Text = Format(rs!cas_Kad_Kredit, "0.00") 'Cas Kad Kredit
        If Not IsNull(rs!cas_debit_kad) Then Frm102.L27_Text = Format(rs!cas_debit_kad, "0.00") 'Cas Debit Kredit
        If Not IsNull(rs!harga_999) Then
            Frm102.TB2 = Format(rs!harga_999, "0.00")
            Frm102.TB11 = Format(rs!harga_999, "0.00")
        Else
            Frm102.TB2 = "0.00"
            Frm102.TB11 = "0.00"
        End If

        GLOBAL_DISABLE = 0
    End If
End If

rs.Close
Set rs = Nothing

'###Senarai Nama Pekerja###
Frm102.CBB4.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where status='" & "Aktif" & "' AND InvestorSmall = 0 AND InvestorBig = 0", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    If rs!user_level = 1 Or rs!user_level = 2 Or rs!user_level = 3 Or rs!user_level = G_LEVEL_USER Then Frm102.CBB4.AddItem rs!Samaran & "  |  " & rs!NoPekerja
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm102.CBB1.Clear
Frm102.CBB2.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database order by kadar_tukaran_9999 DESC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!kadar_tukaran_9999) Then Frm102.CBB1.AddItem rs!kadar_tukaran_9999
    If Not IsNull(rs!Metal_Purity) Then Frm102.CBB2.AddItem rs!Metal_Purity
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'###Padam Table Jualan Temp### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "TRUNCATE TABLE " & G_JUALAN_TEMP & ""

Set rs = cn.Execute(strsql)
Set rs = Nothing
'###Padam Table Jualan Temp### - End

'###Padam Table Belian Temp### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "TRUNCATE TABLE 49_belian_temp"

Set rs = cn.Execute(strsql)
Set rs = Nothing
'###Padam Table Belian Temp### - End

Call Frm102_jurujual
End Sub
Sub frm102_calc1()
'On Error Resume Next
Dim Frm102_LM_BERAT As Double
Dim Frm102_LM_KADAR_TUKARAN As Double

Frm102_LM_BERAT = 0 'Berat jualan (g)
Frm102_LM_KADAR_TUKARAN = 0 'Kadar tukaran kepada purity 999.9

If ((Frm102.TB3 <> vbNullString And IsNumeric(Frm102.TB3)) And (Frm102.CBB1 <> vbNullString And IsNumeric(Frm102.CBB1))) Then

    Frm102_LM_BERAT = Frm102.TB3 'Berat jualan (g)
    Frm102_LM_KADAR_TUKARAN = Frm102.CBB1 'Kadar tukaran kepada purity 999.9
    
    Frm102.L7_Text = Format(Frm102_LM_BERAT * Frm102_LM_KADAR_TUKARAN, "#,##0.00") 'Berat 999.9
    
Else

    Frm102.L7_Text = "0.00" 'Berat 999.9
    
End If
End Sub
Sub frm102_calc2()
'On Error Resume Next
Dim Frm102_LM_KADAR_GST As Double
Dim Frm102_LM_UPAH As Double

Frm102_LM_KADAR_GST = 0
Frm102_LM_UPAH = 0

If IsNumeric(Frm102.L21_Text) Then Frm102_LM_KADAR_GST = Frm102.L21_Text 'Kadar gst (%)
If IsNumeric(Frm102.TB4) Then Frm102_LM_UPAH = Frm102.TB4 'Upah (RM)

If Frm102.L21_Text <> vbNullString And IsNumeric(Frm102.L21_Text) Then

    If Frm102.TB4 <> vbNullString And IsNumeric(Frm102.TB4) Then
    
        If Frm102.CB2 = 1 Then 'Upah : GST ZR
        
            Frm102.L30_Text = Format(Frm102.TB4, "#,##0.00") 'Harga upah tanpa GST
            Frm102.TB5 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
            
        End If
        
        If Frm102.CB3 = 1 Then
        
            If Frm102.CB4 = 0 Then
                
                Frm102.L30_Text = Format(Frm102_LM_UPAH, "#,##0.00") 'Harga upah tanpa GST
                Frm102.TB5 = Format(Frm102_LM_UPAH * (Frm102_LM_KADAR_GST / 100), "#,##0.00") 'Jumlah GST (Bagi jualan setiap item)
            
            ElseIf Frm102.CB4 = 1 Then
            
                Frm102.L30_Text = Format(Frm102_LM_UPAH / (1 + (Frm102_LM_KADAR_GST / 100)), "#,##0.00") 'Harga upah tanpa GST
                Frm102.TB5 = Format(Frm102_LM_UPAH - (Frm102_LM_UPAH / (1 + (Frm102_LM_KADAR_GST / 100))), "#,##0.00") 'Jumlah GST (Bagi jualan setiap item)
            
            End If
            
        End If

    Else
    
        Frm102.L30_Text = Format(0, "#,##0.00") 'Harga upah tanpa GST
        Frm102.TB5 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
        
    End If
    
    If Frm102.CB3 = 0 And Frm102.CB4 = 0 Then
    
        If IsNumeric(Frm102.TB4) Then
        
            Frm102.L30_Text = Format(Frm102.TB4, "#,##0.00") 'Harga upah tanpa GST
            Frm102.TB5 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
            
        Else
            
            Frm102.L30_Text = Format(0, "#,##0.00") 'Harga upah tanpa GST
            Frm102.TB5 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
            
        End If
    
    End If

Else

    If IsNumeric(Frm102.TB4) Then
    
        Frm102.L30_Text = Format(Frm102.TB4, "#,##0.00") 'Harga upah tanpa GST
        Frm102.TB5 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
        
    Else
        
        Frm102.L30_Text = Format(0, "#,##0.00") 'Harga upah tanpa GST
        Frm102.TB5 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
        
    End If
    
End If
End Sub
Sub frm102_calc3()
'On Error Resume Next
Dim Frm102_LM_UPAH_TANPA_GST As Double
Dim Frm102_LM_GST As Double

Frm102_LM_UPAH_TANPA_GST = 0 'Jumlah upah tanpa GST
Frm102_LM_GST = 0 'Jumlah GST

If ((Frm102.TB5 <> vbNullString And IsNumeric(Frm102.TB5)) And (Frm102.L30_Text <> vbNullString And IsNumeric(Frm102.L30_Text))) Then

    Frm102_LM_GST = Frm102.TB5 'Jumlah GST (Bagi jualan setiap item)
    Frm102_LM_UPAH_TANPA_GST = Frm102.L30_Text 'Harga upah tanpa GST
    
    Frm102.TB6 = Format(Frm102_LM_GST + Frm102_LM_UPAH_TANPA_GST, "#,##0.00") 'Jumlah Upah + GST (Bagi jualan setiap item)
    
Else

    Frm102.TB6 = "0.00" 'Jumlah Upah + GST (Bagi jualan setiap item)
    
End If
End Sub
Sub frm102_calc4()
'On Error Resume Next
Dim Frm102_LM_BERAT_JUALAN As Double
Dim Frm102_LM_BERAT_BELIAN As Double
Dim Frm102_LM_BEZA_BERAT As Double

Frm102_LM_BERAT_JUALAN = 0 'Berat jualan (g)
Frm102_LM_BERAT_BELIAN = 0 'Kadar belian (g)
Frm102_LM_BEZA_BERAT = 0 'Beza Berat (g)

If ((Frm102.L9_Text <> vbNullString And IsNumeric(Frm102.L9_Text)) And (Frm102.L10_Text <> vbNullString And IsNumeric(Frm102.L10_Text))) Then

    Frm102_LM_BERAT_JUALAN = Frm102.L9_Text 'Berat jualan (g)
    Frm102_LM_BERAT_BELIAN = Frm102.L10_Text 'Kadar belian (g)
    
    Frm102_LM_BEZA_BERAT = Frm102_LM_BERAT_JUALAN - Frm102_LM_BERAT_BELIAN 'Beza Berat (g)
    
    If Frm102_LM_BEZA_BERAT >= 0 Then
        
        Frm102.L34_Text.Visible = False
        Frm102.L11_Text = Format(Frm102_LM_BEZA_BERAT, "#,##0.00") 'Beza Berat
        
    Else
        
        Frm102.L34_Text.Visible = True
        Frm102.L11_Text = Format((-1) * Frm102_LM_BEZA_BERAT, "#,##0.00") 'Beza Berat
    
    End If
    
Else

    Frm102.L11_Text = "0.00" 'Beza Berat
    Frm102.L34_Text.Visible = False
    
End If
End Sub
Sub frm102_calc5()
'On Error Resume Next
Dim Frm102_LM_BEZA_BERAT As Double
Dim Frm102_LM_HARGA_SEMASA As Double

Frm102_LM_BEZA_BERAT = 0 'Beza berat (g)
Frm102_LM_HARGA_SEMASA = 0 'Harga semasa (RM/g)

If ((Frm102.L11_Text <> vbNullString And IsNumeric(Frm102.L11_Text)) And (Frm102.TB11 <> vbNullString And IsNumeric(Frm102.TB11))) Then
    Frm102_LM_BEZA_BERAT = Frm102.L11_Text 'Berat jualan (g)
    Frm102_LM_HARGA_SEMASA = Frm102.TB11 'Kadar belian (g)
    
    Frm102.L12_Text = Format(Frm102_LM_BEZA_BERAT * Frm102_LM_HARGA_SEMASA, "#,##0.00") 'Harga jualan
Else
    Frm102.L12_Text = "0.00" 'Harga jualan
End If
End Sub
Sub frm102_calc6()
'On Error Resume Next
Dim Frm102_LM_BERAT As Double
Dim Frm102_LM_KADAR_TUKARAN As Double

Frm102_LM_BERAT = 0 'Berat (g)
Frm102_LM_KADAR_TUKARAN = 0 'Kadar Tukaran

If ((Frm102.TB10 <> vbNullString And IsNumeric(Frm102.TB10)) And (Frm102.TB24 <> vbNullString And IsNumeric(Frm102.TB24))) Then
    Frm102_LM_BERAT = Frm102.TB10 'Berat (g)
    Frm102_LM_KADAR_TUKARAN = Frm102.TB24 'Kadar Tukaran
    
    Frm102.L8_Text = Format(Frm102_LM_BERAT * Frm102_LM_KADAR_TUKARAN, "#,##0.00") 'Berat dalam purity 999.9
Else
    Frm102.L8_Text = "0.00" 'Berat dalam purity 999.9
End If
End Sub
Sub frm102_calc7()
'On Error Resume Next
Dim Frm102_LM_HARGA_EMAS As Double
Dim Frm102_LM_UPAH As Double

Frm102_LM_HARGA_EMAS = 0 'Jumlah harga emas (RM)
Frm102_LM_UPAH = 0 'Jumlah Upah (RM)

If ((Frm102.L13_Text <> vbNullString And IsNumeric(Frm102.L13_Text)) And (Frm102.TB13 <> vbNullString And IsNumeric(Frm102.TB13))) Then
    Frm102_LM_HARGA_EMAS = Frm102.TB13 'Jumlah harga emas (RM)
    Frm102_LM_UPAH = Frm102.L13_Text 'Jumlah Upah (RM)
    
    Frm102.L14_Text = Format(Frm102_LM_HARGA_EMAS + Frm102_LM_UPAH, "#,##0.00") 'Jumlah bayaran keseluruhan (RM)
Else
    Frm102.L14_Text = "0.00" 'Jumlah bayaran keseluruhan (RM)
End If

Call frm102_calc10
End Sub
Sub frm102_calc8()
'On Error Resume Next
Dim Frm102_LM_KADAR_GST As Double
Dim Frm102_LM_UPAH As Double

Frm102_LM_KADAR_GST = 0
Frm102_LM_UPAH = 0

If IsNumeric(Frm102.L21_Text) Then Frm102_LM_KADAR_GST = Frm102.L21_Text 'Kadar gst (%)
If IsNumeric(Frm102.L12_Text) Then Frm102_LM_UPAH = Frm102.L12_Text 'Upah (RM)

If Frm102.L34_Text.Visible = False Then
    If Frm102.L21_Text <> vbNullString And IsNumeric(Frm102.L21_Text) Then
    
        If Frm102.L12_Text <> vbNullString And IsNumeric(Frm102.L12_Text) Then
        
            If Frm102.CB5 = 1 Then 'Upah : GST ZR
            
                Frm102.L31_Text = Format(Frm102.L12_Text, "#,##0.00") 'Harga upah tanpa GST
                Frm102.TB12 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
                
            End If
            
            If Frm102.CB6 = 1 Then
            
                If Frm102.CB7 = 0 Then
                    
                    Frm102.L31_Text = Format(Frm102_LM_UPAH, "#,##0.00") 'Harga upah tanpa GST
                    Frm102.TB12 = Format(Frm102_LM_UPAH * (Frm102_LM_KADAR_GST / 100), "#,##0.00") 'Jumlah GST (Bagi jualan setiap item)
                
                ElseIf Frm102.CB7 = 1 Then
                
                    Frm102.L31_Text = Format(Frm102_LM_UPAH / (1 + (Frm102_LM_KADAR_GST / 100)), "#,##0.00") 'Harga upah tanpa GST
                    Frm102.TB12 = Format(Frm102_LM_UPAH - (Frm102_LM_UPAH / (1 + (Frm102_LM_KADAR_GST / 100))), "#,##0.00") 'Jumlah GST (Bagi jualan setiap item)
                
                End If
                
            End If
    
        Else
        
            Frm102.L31_Text = Format(0, "#,##0.00") 'Harga upah tanpa GST
            Frm102.TB12 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
            
        End If
        
        If Frm102.CB6 = 0 And Frm102.CB7 = 0 Then
        
            If IsNumeric(Frm102.L12_Text) Then
            
                Frm102.L31_Text = Format(Frm102.L12_Text, "#,##0.00") 'Harga upah tanpa GST
                Frm102.TB12 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
                
            Else
                
                Frm102.L31_Text = Format(0, "#,##0.00") 'Harga upah tanpa GST
                Frm102.TB12 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
                
            End If
        
        End If
    
    Else
    
        If IsNumeric(Frm102.L12_Text) Then
        
            Frm102.L31_Text = Format(Frm102.L12_Text, "#,##0.00") 'Harga upah tanpa GST
            Frm102.TB12 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
            
        Else
            
            Frm102.L31_Text = Format(0, "#,##0.00") 'Harga upah tanpa GST
            Frm102.TB12 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
            
        End If
        
    End If
Else
    Frm102.TB12 = "0.00"
    Frm102.TB13 = "0.00"
    Frm102.L31_Text = Format(0, "#,##0.00") 'Harga upah tanpa GST
End If

'### Pengiraan GST bagi harga emas SAHAJA ### - Start
If Frm102.CB5 = 1 Then

    Frm102.L39_Text = Frm102.L31_Text
    Frm102.L40_Text = Frm102.TB12
    Frm102.L41_Text = 0
    Frm102.L42_Text = 0
    
ElseIf Frm102.CB6 = 1 Then

    Frm102.L41_Text = Frm102.L31_Text
    Frm102.L42_Text = Frm102.TB12
    Frm102.L39_Text = 0
    Frm102.L40_Text = 0

End If

Call frm102_calc10
End Sub
Sub frm102_calc9()
'On Error Resume Next
Dim Frm102_LM_HARGA_EMAS_TANPA_GST As Double
Dim Frm102_LM_GST As Double

Frm102_LM_HARGA_EMAS_TANPA_GST = 0 'Jumlah upah tanpa GST
Frm102_LM_GST = 0 'Jumlah GST

If ((Frm102.TB12 <> vbNullString And IsNumeric(Frm102.TB12)) And (Frm102.L31_Text <> vbNullString And IsNumeric(Frm102.L31_Text))) Then

    Frm102_LM_GST = Frm102.TB12 'Jumlah GST (Bagi jualan setiap item)
    Frm102_LM_HARGA_EMAS_TANPA_GST = Frm102.L31_Text 'Harga upah tanpa GST
    
    Frm102.TB13 = Format(Frm102_LM_GST + Frm102_LM_HARGA_EMAS_TANPA_GST, "#,##0.00") 'Jumlah Upah + GST (Bagi jualan setiap item)
    
Else

    Frm102.TB13 = "0.00" 'Jumlah Upah + GST (Bagi jualan setiap item)
    
End If

Call frm102_calc10
End Sub
Sub frm102_calc10()
'On Error Resume Next
Dim Frm102_LM_HARGA_ZR_UPAH As Double
Dim Frm102_LM_HARGA_SR_UPAH As Double
Dim Frm102_LM_HARGA_ZR_EMAS As Double
Dim Frm102_LM_HARGA_SR_EMAS As Double
Dim Frm102_LM_GST_ZR_UPAH As Double
Dim Frm102_LM_GST_SR_UPAH As Double
Dim Frm102_LM_GST_ZR_EMAS As Double
Dim Frm102_LM_GST_SR_EMAS As Double

Frm102_LM_HARGA_ZR_UPAH = 0
Frm102_LM_HARGA_SR_UPAH = 0
Frm102_LM_HARGA_ZR_EMAS = 0
Frm102_LM_HARGA_SR_EMAS = 0
Frm102_LM_GST_ZR_UPAH = 0
Frm102_LM_GST_SR_UPAH = 0
Frm102_LM_GST_ZR_EMAS = 0
Frm102_LM_GST_SR_EMAS = 0

If ((Frm102.L35_Text <> vbNullString And IsNumeric(Frm102.L35_Text)) And (Frm102.L39_Text <> vbNullString And IsNumeric(Frm102.L39_Text))) Then

    Frm102_LM_HARGA_ZR_UPAH = Frm102.L35_Text 'Harga ZR (Upah)
    Frm102_LM_HARGA_ZR_EMAS = Frm102.L39_Text 'Harga ZR (Emas)
    
    Frm102.L17_Text = Format(Frm102_LM_HARGA_ZR_UPAH + Frm102_LM_HARGA_ZR_EMAS, "#,##0.00") 'Jumlah Harga ZR
    
Else

    Frm102.L17_Text = "0.00" 'Jumlah Harga ZR
    
End If

If ((Frm102.L37_Text <> vbNullString And IsNumeric(Frm102.L37_Text)) And (Frm102.L41_Text <> vbNullString And IsNumeric(Frm102.L41_Text))) Then

    Frm102_LM_HARGA_SR_UPAH = Frm102.L37_Text 'Harga SR (Upah)
    Frm102_LM_HARGA_SR_EMAS = Frm102.L41_Text 'Harga SR (Emas)
    
    Frm102.L18_Text = Format(Frm102_LM_HARGA_SR_UPAH + Frm102_LM_HARGA_SR_EMAS, "#,##0.00") 'Jumlah Harga SR
    
Else

    Frm102.L18_Text = "0.00" 'Jumlah Harga SR
    
End If

If ((Frm102.L36_Text <> vbNullString And IsNumeric(Frm102.L36_Text)) And (Frm102.L40_Text <> vbNullString And IsNumeric(Frm102.L40_Text))) Then

    Frm102_LM_GST_SR_UPAH = Frm102.L36_Text 'GST ZR (Upah)
    Frm102_LM_GST_SR_EMAS = Frm102.L40_Text 'GST ZR (Emas)
    
    Frm102.L20_Text = Format(Frm102_LM_GST_SR_UPAH + Frm102_LM_GST_SR_EMAS, "#,##0.00") 'Jumlah GST ZR
    
Else

    Frm102.L20_Text = "0.00" 'Jumlah GST ZR
    
End If

If ((Frm102.L38_Text <> vbNullString And IsNumeric(Frm102.L38_Text)) And (Frm102.L42_Text <> vbNullString And IsNumeric(Frm102.L42_Text))) Then

    Frm102_LM_GST_ZR_UPAH = Frm102.L38_Text 'GST SR (Upah)
    Frm102_LM_GST_ZR_EMAS = Frm102.L42_Text 'GST SR (Emas)
    
    Frm102.L20_Text = Format(Frm102_LM_GST_ZR_UPAH + Frm102_LM_GST_ZR_EMAS, "#,##0.00") 'Jumlah GST SR
    
Else

    Frm102.L20_Text = "0.00" 'Jumlah GST SR
    
End If

Frm102.L15_Text = Format(Frm102_LM_HARGA_ZR_UPAH + Frm102_LM_HARGA_ZR_EMAS + Frm102_LM_HARGA_SR_UPAH + Frm102_LM_HARGA_SR_EMAS, "#,##0.00") 'Jumlah harga tanpa GST
Frm102.L16_Text = Format(Frm102_LM_HARGA_ZR_UPAH + Frm102_LM_HARGA_ZR_EMAS + Frm102_LM_HARGA_SR_UPAH + Frm102_LM_HARGA_SR_EMAS + Frm102_LM_GST_SR_UPAH + Frm102_LM_GST_SR_EMAS + Frm102_LM_GST_ZR_UPAH + Frm102_LM_GST_ZR_EMAS, "#,##0.00") 'Jumlah harga dengan GST
End Sub
Sub frm102_calc11()
'On Error Resume Next
Dim Frm102_LM_TUNAI As Double
Dim Frm102_LM_BANK As Double
Dim Frm102_LM_KREDIT As Double
Dim Frm102_LM_DEBIT As Double
Dim Frm102_LM_SIMPANAN As Double

Frm102_LM_TUNAI = 0
Frm102_LM_BANK = 0
Frm102_LM_KREDIT = 0
Frm102_LM_DEBIT = 0
Frm102_LM_SIMPANAN = 0

If (Frm102.TB14 <> vbNullString And IsNumeric(Frm102.TB14)) Then
    Frm102_LM_TUNAI = Frm102.TB14 'Tunai
End If

If (Frm102.TB15 <> vbNullString And IsNumeric(Frm102.TB15)) Then
    Frm102_LM_BANK = Frm102.TB15 'Bank In
End If

If (Frm102.TB16 <> vbNullString And IsNumeric(Frm102.TB16)) Then
    Frm102_LM_KREDIT = Frm102.TB16 'Kad Kredit
End If

If (Frm102.TB19 <> vbNullString And IsNumeric(Frm102.TB19)) Then
    Frm102_LM_DEBIT = Frm102.TB19 'Kad Debit
End If

If (Frm102.TB22 <> vbNullString And IsNumeric(Frm102.TB22)) Then
    Frm102_LM_SIMPANAN = Frm102.TB22 'Simpanan di kedai
End If

Frm102.TB23 = Format(Frm102_LM_TUNAI + Frm102_LM_BANK + Frm102_LM_KREDIT + Frm102_LM_DEBIT + Frm102_LM_SIMPANAN, "#,##0.00") 'Jumlah bayaran
End Sub
Sub frm102_calc12()
'On Error Resume Next
Dim Frm102_LM_KAD As Double
Dim Frm102_LM_CAS As Double

Frm102_LM_KAD = 0 'Jumlah Kad
Frm102_LM_CAS = 0 'Kadar Cas Kad

If ((Frm102.L26_Text <> vbNullString And IsNumeric(Frm102.L26_Text)) And (Frm102.TB16 <> vbNullString And IsNumeric(Frm102.TB16))) Then

    Frm102_LM_CAS = Frm102.L26_Text 'Kadar Cas Kad
    Frm102_LM_KAD = Frm102.TB16 'Jumlah Kad
    
    Frm102.TB17 = Format(Frm102_LM_KAD * (Frm102_LM_CAS / 100), "#,##0.00") 'Jumlah cas kad
    
Else

    Frm102.TB17 = "0.00" 'Jumlah cas kad
    
End If
End Sub
Sub frm102_calc13()
'On Error Resume Next
Dim Frm102_LM_KAD As Double
Dim Frm102_LM_CAS As Double

Frm102_LM_KAD = 0 'Jumlah Kad
Frm102_LM_CAS = 0 'Jumlah Cas Kad

If ((Frm102.TB17 <> vbNullString And IsNumeric(Frm102.TB17)) And (Frm102.TB16 <> vbNullString And IsNumeric(Frm102.TB16))) Then

    Frm102_LM_CAS = Frm102.TB17 'Jumlah Cas Kad
    Frm102_LM_KAD = Frm102.TB16 'Jumlah Kad
    
    Frm102.TB18 = Format(Frm102_LM_KAD + Frm102_LM_CAS, "#,##0.00") 'Jumlah cas kad
    
Else

    Frm102.TB18 = "0.00" 'Jumlah cas kad
    
End If
End Sub
Sub frm102_calc14()
'On Error Resume Next
Dim Frm102_LM_KAD As Double
Dim Frm102_LM_CAS As Double

Frm102_LM_KAD = 0 'Jumlah Kad
Frm102_LM_CAS = 0 'Kadar Cas Kad

If ((Frm102.L27_Text <> vbNullString And IsNumeric(Frm102.L27_Text)) And (Frm102.TB19 <> vbNullString And IsNumeric(Frm102.TB19))) Then

    Frm102_LM_CAS = Frm102.L27_Text 'Kadar Cas Kad
    Frm102_LM_KAD = Frm102.TB19 'Jumlah Kad
    
    Frm102.TB20 = Format(Frm102_LM_KAD * (Frm102_LM_CAS / 100), "#,##0.00") 'Jumlah cas kad
    
Else

    Frm102.TB20 = "0.00" 'Jumlah cas kad
    
End If
End Sub
Sub frm102_calc15()
'On Error Resume Next
Dim Frm102_LM_KAD As Double
Dim Frm102_LM_CAS As Double

Frm102_LM_KAD = 0 'Jumlah Kad
Frm102_LM_CAS = 0 'Jumlah Cas Kad

If ((Frm102.TB20 <> vbNullString And IsNumeric(Frm102.TB20)) And (Frm102.TB19 <> vbNullString And IsNumeric(Frm102.TB19))) Then

    Frm102_LM_CAS = Frm102.TB20 'Jumlah Cas Kad
    Frm102_LM_KAD = Frm102.TB19 'Jumlah Kad
    
    Frm102.TB21 = Format(Frm102_LM_KAD + Frm102_LM_CAS, "#,##0.00") 'Jumlah cas kad
    
Else

    Frm102.TB21 = "0.00" 'Jumlah cas kad
    
End If
End Sub
Sub Frm102_Call_Product_Detail()
'on error resume next
Dim Frm102_LM_BERAT As Double

Frm102_LM_DATA_FOUND = 0
Frm102_LM_BERAT = 0
Frm102_LM_READY_TO_SAVE = 0 'Flag : Ready To Save
Frm102_LM_UpdateList = 0
Frm102_LM_KOD_PURITY = vbNullString
Frm102_LM_PERMATA = 0

Frm102_LM_No_SIRI = UCase(Frm102.TB1) 'No. Siri Produk
Frm102.TB1 = vbNullString

Frm102_LM_KATEGORI = 1
'1:  Pelanggan Biasa
'2 : Member / Ahli
'3:  Pengedar
'4:  RAF
'5:  Normal Dealer(ND)
'6:  Master Dealer(MD)


' ### Periksa kategori pembeli ### - Start
If Frm102.L46_Text <> vbNullString Then
    If Frm28.L5_Text <> vbNullString Then
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
        If Not rs.EOF Then
        
            If Not IsNull(rs!kategori_pelanggan) Then Frm102_LM_KATEGORI = rs!kategori_pelanggan
            
        End If
        
        rs.Close
        Set rs = Nothing
        
    End If
End If
' ### Periksa kategori pembeli ### - End

'###Periksa Samada Data Ini Telah Dimasukkan Ke Dalam Temp Table### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from " & G_JUALAN_TEMP & " where no_siri_produk='" & Frm102_LM_No_SIRI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Frm102.L32_Text = "0" Then 'Data Baru (Kemasukkan Baru)
        If rs!Status = "1" Or rs!Status = "4" Then
        
            MsgBox "Item ini telah dimasukkan ke dalam senarai jualan sebelum ini.", vbInformation, "Info"
            
            Frm102.TB1 = vbNullString
            Frm102.TB1.SetFocus
            
        ElseIf rs!Status = 0 Then
            rs!Status = 1 '0 : Padam Data , 1 : Aktif (Tiada Perubahan) , 3 : Data Baru , 4 : Data Diedit , 5 : Keluarkan Data Dari Database Asal , 6 : Ignore Kemasukkan Data Ke Dalam Database
            rs.Update
            
            Frm102_LM_UpdateList = 1
        End If
    ElseIf Frm102.L32_Text = "1" Then 'Edit Data Lama + Kemasukkan Baru
        If rs!Status = "1" Or rs!Status = "4" Or rs!Status = "3" Then
        
            MsgBox "Item ini telah dimasukkan ke dalam senarai jualan sebelum ini.", vbInformation, "Info"
            
            Frm102.TB1 = vbNullString
            Frm102.TB1.SetFocus
            
        ElseIf rs!Status = "5" Or rs!Status = "6" Then
            If rs!Status = "5" Then rs!Status = "4" '0 : Padam Data , 1 : Aktif (Tiada Perubahan) , 3 : Data Baru , 4 : Data Diedit , 5 : Keluarkan Data Dari Database Asal , 6 : Ignore Kemasukkan Data Ke Dalam Database
            If rs!Status = "6" Then rs!Status = "3" '0 : Padam Data , 1 : Aktif (Tiada Perubahan) , 3 : Data Baru , 4 : Data Diedit , 5 : Keluarkan Data Dari Database Asal , 6 : Ignore Kemasukkan Data Ke Dalam Database
            rs.Update
            
            Frm102.TB1 = vbNullString
            Frm102.TB1.SetFocus
            
            Frm102_LM_UpdateList = 1
        End If
    End If
    Frm102_LM_DATA_FOUND = 1
    If rs!Status = "0" Or rs!Status = "5" Then Frm102_LM_DATA_FOUND = 0
End If

rs.Close
Set rs = Nothing
'###Periksa Samada Data Ini Telah Dimasukkan Ke Dalam Temp Table### - End

'###Carian Data Basic Bagi Item Ini### - Start
If Frm102_LM_DATA_FOUND = 0 Then

'###Periksa Mode Upah### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from default_setting where Default1='" & "Default" & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!flag_upah) Then
            If rs!flag_upah = 1 Then
                LM_UPAH_MODE = 1
            Else
                LM_UPAH_MODE = 0
            End If
        End If
    End If
    
    rs.Close
    Set rs = Nothing
'###Periksa Mode Upah### - End

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from Data_Database where no_siri_Produk='" & Frm102_LM_No_SIRI & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If rs!StatusItem = "10" Then
            If Not IsNull(rs!receiving_Status) Then
                If rs!receiving_Status = 0 Or rs!receiving_Status = 2 Or rs!receiving_Status = 4 Or rs!receiving_Status = 5 Then
                    
                    Frm102.L3_Text = Frm102_LM_No_SIRI 'No. Siri Produk
                    Frm102.L6_Text = Format(rs!beza_berat, "0.00") 'Berat Asal (g)
                    Frm102.TB3 = Format(rs!beza_berat, "0.00") 'Berat Jualan (g)
                    Frm102.L33_Text = Format(rs!harga_Per_Gram_Item, "0.00") 'Harga Per Gram Item (RM/g)
                    Frm102.L50_Text = rs!UPAH 'Upah modal
                    Frm102.TB4.Locked = False 'Upah
                    Frm102.TB4.BackColor = &HFFFFFF 'Upah
                    
                    Frm102_LM_READY_TO_SAVE = 1 'Flag : Ready To Save
                Else
                    MsgBox "Barang yang ingin dijual [" & Frm102_LM_No_SIRI & "] adalah barang permata." & vbCrLf & _
                            vbNullString & vbCrLf & _
                            "Hanya barang kemas (yang mempunyai berat) SAHAJA dibenarkan dijual dalam menu ini.", vbInformation, "Info"
                            
                    Frm102.TB1 = vbNullString
                    Frm102.TB1.SetFocus
                            
                    Exit Sub
                End If
            End If
            
            If LM_UPAH_MODE = 1 Then
                If Frm102_LM_KATEGORI = 1 Then
                    If Not IsNull(rs!Upah_Jualan) Then
                        Frm102.TB4 = Format(rs!Upah_Jualan, "0.00") 'Upah Pelanggan
                    End If
                ElseIf Frm102_LM_KATEGORI = 2 Then
                    If Not IsNull(rs!Upah_Member) Then
                        Frm102.TB4 = Format(rs!Upah_Member, "0.00") 'Upah Member
                    End If
                ElseIf Frm102_LM_KATEGORI = 3 Then
                    If Not IsNull(rs!Upah_RAF) Then
                        Frm102.TB4 = Format(rs!Upah_RAF, "0.00") 'Upah RAF
                    End If
                ElseIf Frm102_LM_KATEGORI = 4 Then
                    If Not IsNull(rs!Upah_Pengedar) Then
                        Frm102.TB4 = Format(rs!Upah_Pengedar, "0.00") 'Upah Pengedar
                    End If
                ElseIf Frm102_LM_KATEGORI = 5 Then
                    If Not IsNull(rs!upah_normal_dealer) Then
                        Frm102.TB4 = Format(rs!upah_normal_dealer, "0.00") 'Upah Normal Dealer
                    End If
                ElseIf Frm102_LM_KATEGORI = 6 Then
                    If Not IsNull(rs!upah_master_dealer) Then
                        Frm102.TB4 = Format(rs!upah_master_dealer, "0.00") 'Upah Master Dealer
                    End If
                End If
            Else
                Frm102.TB4 = Format(0, "0.00") 'Upah
            End If
            
            If Not IsNull(rs!kategori_Produk) Then Frm102.L5_Text = rs!kategori_Produk 'Kategori Produk
            If Not IsNull(rs!kod_Purity) Then
                Frm102_LM_KOD_PURITY = rs!kod_Purity 'Kod Purity
                Frm102.L4_Text = rs!kod_Purity 'Kod Purity
            End If
        ElseIf rs!StatusItem = "11" Then
            MsgBox "Item ini telah dijual. No. Siri Produk [" & Frm102_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm102.TB1 = vbNullString
            Frm102.TB1.SetFocus
        ElseIf rs!StatusItem = "12" Then
            MsgBox "Item ini telah dijual secara potong. No. Siri Produk [" & Frm102_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm102.TB1 = vbNullString
            Frm102.TB1.SetFocus
        ElseIf rs!StatusItem = "13" Then
            MsgBox "Item ini telah dijual secara potong. No. Siri Produk [" & Frm102_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm102.TB1 = vbNullString
            Frm102.TB1.SetFocus
        ElseIf rs!StatusItem = "14" Or rs!StatusItem = "21" Or rs!StatusItem = "22" Then
            MsgBox "Item ini telah ditempah oleh pelanggan. No. Siri Produk [" & Frm102_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm102.TB1 = vbNullString
            Frm102.TB1.SetFocus
        ElseIf rs!StatusItem = "15" Or rs!StatusItem = "19" Or rs!StatusItem = "20" Then
            MsgBox "Item ini telah dibeli secara ansuran. No. Siri Produk [" & Frm102_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm102.TB1 = vbNullString
            Frm102.TB1.SetFocus
        ElseIf rs!StatusItem = "16" Then
            MsgBox "Item ini telah dihantar ke Ar-Rahnu. No. Siri Produk [" & Frm102_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm102.TB1 = vbNullString
            Frm102.TB1.SetFocus
        ElseIf rs!StatusItem = "17" Then
            MsgBox "Item ini telah dijual secara ETA. No. Siri Produk [" & Frm102_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm102.TB1 = vbNullString
            Frm102.TB1.SetFocus
        ElseIf rs!StatusItem = "23" Then
            MsgBox "Item Ini Telah Dihantar Ke Supplier/Kilang. No. Siri Produk [" & Frm102_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm102.TB1 = vbNullString
            Frm102.TB1.SetFocus
        ElseIf rs!StatusItem = "24" Then
            MsgBox "Item Ini Telah Dihantar Ke Supplier/Kilang. No. Siri Produk [" & Frm102_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm102.TB1 = vbNullString
            Frm102.TB1.SetFocus
        ElseIf rs!StatusItem = "25" Then
            MsgBox "Item Ini Telah Diagihkan Ke Cawangan. No. Siri Produk [" & Frm102_LM_No_SIRI & "]", vbExclamation, "Info"
    
            Frm102.TB1 = vbNullString
            Frm102.TB1.SetFocus
        ElseIf rs!StatusItem = "26" Then
            MsgBox "Item Ini Telah Dijual Oleh Cawangan. No. Siri Produk [" & Frm102_LM_No_SIRI & "]", vbExclamation, "Info"
    
            Frm102.TB1 = vbNullString
            Frm102.TB1.SetFocus
        ElseIf rs!StatusItem = "0" Then
            MsgBox "Item Ini Telah Dipadamkan Dari Sistem. No. Siri Produk [" & Frm102_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm102.TB1 = vbNullString
            Frm102.TB1.SetFocus
        ElseIf rs!StatusItem = "27" Or rs!StatusItem = "28" Then
            MsgBox "Item Ini Telah Dijual Dari Menu GDN. No. Siri Produk [" & Frm102_LM_No_SIRI & "]", vbExclamation, "Info"
    
            Frm102.TB1 = vbNullString
            Frm102.TB1.SetFocus
        ElseIf rs!StatusItem = "29" Then
            MsgBox "Item Ini Telah Diubah Status Kepada Hilang , Dicuri Dan Sebagainya. No. Siri Produk [" & Frm102_LM_No_SIRI & "]", vbExclamation, "Info"
    
            Frm102.TB1 = vbNullString
            Frm102.TB1.SetFocus
        End If
        
    Else
        MsgBox "No. Siri Produk Ini [" & Frm102_LM_No_SIRI & "] Tidak Dijumpai.", vbExclamation, "Info"
        
        Frm102.TB1 = vbNullString
        Frm102.TB1.SetFocus
    End If
    
    rs.Close
    Set rs = Nothing
    
    If Frm102_LM_KOD_PURITY <> vbNullString Then

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from hargaemas where Purity='" & Frm102_LM_KOD_PURITY & "' AND cawangan='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            If Not IsNull(rs!HargaDariSupplier) Then
                If IsNumeric(rs!HargaDariSupplier) Then
                    Frm102.L49_Text = rs!HargaDariSupplier
                Else
                    Frm102.L49_Text = 0
                End If
            Else
                Frm102.L49_Text = 0
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        
    End If
End If
'###Carian Data Basic Bagi Item Ini### - End

'###Periksa Data Produk### - Start

'Frm102.TB1 = vbNullString
'If Frm102.CB1 = 1 Then Call Frm102_auto_insert_data

'If Frm102_LM_UpdateList = 1 Then
    'Call Frm102_Senarai_Jualan_Header
    'Call Frm102_Senarai_Jualan
    Frm102.TB1.SetFocus
'End If
'###Periksa Data Produk### - End
End Sub
Sub Frm102_Senarai_Jualan_Header()
'on error resume next
Frm102.MSFlexGrid1.Clear
Frm102.MSFlexGrid1.RowHeight(0) = 500
Frm102.MSFlexGrid1.FormatString = "No.|<No.|<ID|<No. Siri Produk|<Kategori Produk|<Berat Asal (g)|<Berat Jualan (g)|<Kadar Tukaran 999.9|<Berat 999.9 (g)|<Upah (RM)|<Jenis GST|<Jumlah GST (RM)|<Upah + GST (RM)"

Frm102.MSFlexGrid1.Rows = 1
Frm102.MSFlexGrid1.ColWidth(0) = 600 'No.
Frm102.MSFlexGrid1.ColWidth(1) = 0 'No.
Frm102.MSFlexGrid1.ColWidth(2) = 0 'No. ID
Frm102.MSFlexGrid1.ColWidth(3) = 1600 'No. Siri Produk
Frm102.MSFlexGrid1.ColWidth(4) = 4300 'Kategori Produk
Frm102.MSFlexGrid1.ColWidth(5) = 1000 'Berat Asal (g)
Frm102.MSFlexGrid1.ColWidth(6) = 1000 'Berat Jualan (g)
Frm102.MSFlexGrid1.ColWidth(7) = 1000 'Kadar Tukaran 999.9
Frm102.MSFlexGrid1.ColWidth(8) = 1000 'Berat 999.9 (g)
Frm102.MSFlexGrid1.ColWidth(9) = 1000 'Upah (RM)
Frm102.MSFlexGrid1.ColWidth(10) = 1000 'Jenis GST
Frm102.MSFlexGrid1.ColWidth(11) = 1000 'Jumlah GST (RM)
Frm102.MSFlexGrid1.ColWidth(12) = 1000 'Upah + GST (RM)
End Sub
Sub Frm102_Senarai_Jualan()
'on error resume next
Dim Frm102_LM_UPAH_TANPA_GST As Double 'Harga Jualan Tanpa Cukai GST
Dim Frm102_LM_UPAH_DENGAN_GST As Double 'Harga Jualan Dengan Cukai GST
Dim Frm102_LM_GST_SR As Double 'Kutipan GST : SR
Dim Frm102_LM_GST_ZR As Double 'Kutipan GST : ZR
Dim Frm102_LM_JUMLAH_UPAH_SR As Double 'Total Harga Yang Dikenakan GST SR
Dim Frm102_LM_JUMLAH_UPAH_ZR As Double 'Total Harga Yang Dikenakan GST ZR
Dim Frm102_LM_BERAT As Double 'Berat Jualan
Dim Frm102_LM_BERAT_ASAL As Double 'Berat Asal (Sebelum tukar kepada purity 999.9)

x = 0
Frm102_LM_UPAH_TANPA_GST = 0
Frm102_LM_UPAH_DENGAN_GST = 0
Frm102_LM_GST_SR = 0
Frm102_LM_GST_ZR = 0
Frm102_LM_JUMLAH_UPAH_SR = 0
Frm102_LM_JUMLAH_UPAH_ZR = 0
Frm102_LM_BERAT = 0
Frm102_LM_BERAT_ASAL = 0 'Berat Asal (Sebelum tukar kepada purity 999.9)

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from " & G_JUALAN_TEMP & " where Status='" & 1 & "' Or Status='" & 2 & "' Or Status='" & 3 & "' Or Status='" & 4 & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    'If rs!Status = 1 Or rs!Status = 3 Or rs!Status = 4 Then
        x = x + 1
        Frm102.MSFlexGrid1.Rows = x + 1
        Frm102.MSFlexGrid1.TextMatrix(x, 0) = x 'No.
        Frm102.MSFlexGrid1.TextMatrix(x, 1) = x 'No.
        If Not IsNull(rs!ID) Then Frm102.MSFlexGrid1.TextMatrix(x, 2) = rs!ID 'No. ID
        If Not IsNull(rs!no_siri_Produk) Then Frm102.MSFlexGrid1.TextMatrix(x, 3) = rs!no_siri_Produk 'No. Siri Produk
        If Not IsNull(rs!kategori_Produk) Then Frm102.MSFlexGrid1.TextMatrix(x, 4) = rs!kategori_Produk 'Kategori Produk
        If Not IsNull(rs!Berat_Asal) Then
            Frm102.MSFlexGrid1.TextMatrix(x, 5) = Format(rs!Berat_Asal, "#,##0.00") 'Berat Asal (g)
            If IsNumeric(rs!Berat_Asal) Then Frm102_LM_BERAT_ASAL = Frm102_LM_BERAT_ASAL + rs!Berat_Asal
        End If
        If Not IsNull(rs!berat_jualan) Then Frm102.MSFlexGrid1.TextMatrix(x, 6) = Format(rs!berat_jualan, "#,##0.00") 'Berat Jualan (g)
        If Not IsNull(rs!pemalar_tukaran_999) Then Frm102.MSFlexGrid1.TextMatrix(x, 7) = rs!pemalar_tukaran_999 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
        If Not IsNull(rs!berat_999) Then 'Berat barang kemas selepas ditukar kepada purity 999.9
            Frm102.MSFlexGrid1.TextMatrix(x, 8) = Format(rs!berat_999, "#,##0.00") 'Berat Jualan (g)
            If IsNumeric(rs!berat_999) Then Frm102_LM_BERAT = Frm102_LM_BERAT + rs!berat_999
        End If
        If Not IsNull(rs!UPAH) Then
            Frm102.MSFlexGrid1.TextMatrix(x, 9) = Format(rs!harga_tanpa_gst, "#,##0.00") 'Upah (RM)
            If IsNumeric(rs!harga_tanpa_gst) Then Frm102_LM_UPAH_TANPA_GST = Frm102_LM_UPAH_TANPA_GST + rs!harga_tanpa_gst
        End If
        If Not IsNull(rs!gst_ari_nashi) Then
            If rs!gst_ari_nashi = "ZR (L)" Then
                Frm102.MSFlexGrid1.TextMatrix(x, 10) = "ZR(L)" 'Jenis GST : Zero Rated
                If IsNumeric(rs!jumlah_gst) Then Frm102_LM_GST_ZR = Frm102_LM_GST_ZR + rs!jumlah_gst 'Jumlah Kutipan GST ZR(L)
                If IsNumeric(rs!harga_tanpa_gst) Then Frm102_LM_JUMLAH_UPAH_ZR = Frm102_LM_JUMLAH_UPAH_ZR + rs!harga_tanpa_gst 'Total Harga Yang Dikenakan GST ZR
            ElseIf rs!gst_ari_nashi = "SR" Then
                Frm102.MSFlexGrid1.TextMatrix(x, 10) = "SR" 'Jenis GST : Standard Rated
                If IsNumeric(rs!jumlah_gst) Then Frm102_LM_GST_SR = Frm102_LM_GST_SR + rs!jumlah_gst 'Jumlah Kutipan GST SR
                If IsNumeric(rs!harga_tanpa_gst) Then Frm102_LM_JUMLAH_UPAH_SR = Frm102_LM_JUMLAH_UPAH_SR + rs!harga_tanpa_gst 'Total Harga Yang Dikenakan GST SR
            End If
        End If
        If Not IsNull(rs!jumlah_gst) Then Frm102.MSFlexGrid1.TextMatrix(x, 11) = Format(rs!jumlah_gst, "#,##0.00") 'Jumlah GST (RM)
        If Not IsNull(rs!harga_dengan_gst) Then
            Frm102.MSFlexGrid1.TextMatrix(x, 12) = Format(rs!harga_dengan_gst, "#,##0.00") 'Harga Dengan GST (RM)
            If IsNumeric(rs!harga_dengan_gst) Then Frm102_LM_UPAH_DENGAN_GST = Frm102_LM_UPAH_DENGAN_GST + rs!harga_dengan_gst 'Harga Jualan Dengan GST (RM)
        End If
    'End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm102.L43_Text = x 'Jumlah bilangan barang jualan
Frm102.L48_Text = Format(Frm102_LM_BERAT_ASAL, "0.00") 'Jumlah berat jualan
Frm102.L35_Text = Format(Frm102_LM_JUMLAH_UPAH_ZR, "#,##0.00") 'Maklumat GST : Jumlah harga ZR
Frm102.L37_Text = Format(Frm102_LM_JUMLAH_UPAH_SR, "#,##0.00") 'Maklumat GST : Jumlah harga SR
Frm102.L36_Text = Format(Frm102_LM_GST_ZR, "#,##0.00") 'Maklumat GST : Jumlah GST ZR
Frm102.L38_Text = Format(Frm102_LM_GST_SR, "#,##0.00")  'Maklumat GST : Jumlah GST SR
Frm102.L9_Text = Format(Frm102_LM_BERAT, "#,##0.00") 'Berat jualan 999.9
Frm102.TB7 = Format(Frm102_LM_UPAH_TANPA_GST, "#,##0.00") 'Jumlah upah tanpa GST (keseluruhan)
Frm102.TB8 = Format(Frm102_LM_GST_SR, "#,##0.00") 'Jumlah GST (Keseluruhan)
Frm102.TB9 = Format(Frm102_LM_UPAH_DENGAN_GST, "#,##0.00") 'Jumlah Upah + GST (Keseluruhan)
End Sub
Sub Frm102_senarai_belian_header()
'on error resume next
Frm102.MSFlexGrid2.Clear
Frm102.MSFlexGrid2.RowHeight(0) = 400
Frm102.MSFlexGrid2.FormatString = "No.|<No.|<ID|<Purity|<Berat (g)|<Kadar Tukaran|<Berat 999.9 (g)"

Frm102.MSFlexGrid2.Rows = 1
Frm102.MSFlexGrid2.ColWidth(0) = 600 'No.
Frm102.MSFlexGrid2.ColWidth(1) = 0 'No.
Frm102.MSFlexGrid2.ColWidth(2) = 0 'No. ID
Frm102.MSFlexGrid2.ColWidth(3) = 1800 'Purity
Frm102.MSFlexGrid2.ColWidth(4) = 1200 'Berat (g)
Frm102.MSFlexGrid2.ColWidth(5) = 1200 'Kadar Tukaran
Frm102.MSFlexGrid2.ColWidth(6) = 1200 'Berat 999.9 (g)
End Sub
Sub Frm102_senarai_belian()
'on error resume next
Dim Frm102_LM_BERAT_999 As Double

Frm102_LM_BERAT_999 = 0 'Berat dalam purity 999.9

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 49_belian_temp where status='" & 1 & "' OR status='" & 2 & "' OR status='" & 3 & "' OR status='" & 4 & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm102.MSFlexGrid2.Rows = x + 1
    Frm102.MSFlexGrid2.TextMatrix(x, 0) = x 'No.
    Frm102.MSFlexGrid2.TextMatrix(x, 1) = x 'No.
    If Not IsNull(rs!ID) Then Frm102.MSFlexGrid2.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!purity) Then Frm102.MSFlexGrid2.TextMatrix(x, 3) = rs!purity 'Purity
    If Not IsNull(rs!Berat_Asal) Then Frm102.MSFlexGrid2.TextMatrix(x, 4) = Format(rs!Berat_Asal, "#,##0.00") 'Berat (g)
    If Not IsNull(rs!kadar_tukaran) Then Frm102.MSFlexGrid2.TextMatrix(x, 5) = rs!kadar_tukaran 'Kadar Tukaran
    If Not IsNull(rs!berat_tukaran) Then
        Frm102.MSFlexGrid2.TextMatrix(x, 6) = Format(rs!berat_tukaran, "#,##0.00") 'Berat 999.9 (g)
        If IsNumeric(rs!berat_tukaran) Then Frm102_LM_BERAT_999 = Frm102_LM_BERAT_999 + rs!berat_tukaran 'Berat dalam purity 999.9
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm102.L44_Text = x 'Jumlah bilangan barang trade in
Frm102.L10_Text = Format(Frm102_LM_BERAT_999, "#,##0.00")
End Sub
Sub Frm102_recall_edit_jualan()
'on error resume next
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset

Frm102_LM_No_PEKERJA = vbNullString
Frm102_LM_No_PEMBELI = vbNullString

GLOBAL_DISABLE = 1

'### Maklumat asas bagi invoice ### - Start
'Nama pekerja
'Tarikh
'Kadar caj kad kredit
'Kadar caj kad debit

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where no_resit='" & Frm102.L23_Text & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Not IsNull(rs!tarikh) Then Frm102.DTPicker1 = rs!tarikh 'Tarikh Jualan
    If Not IsNull(rs!cas_Kad_Kredit) Then Frm102.L26_Text = rs!cas_Kad_Kredit 'Cas Kad Kredit (%)
    If Not IsNull(rs!cas_kad_debit) Then Frm102.L27_Text = rs!cas_kad_debit 'Cas Kad Debit (%)
    If Not IsNull(rs!no_rujukan_pembeli) Then Frm102_LM_No_PEMBELI = rs!no_rujukan_pembeli 'No. Rujukan Pembeli
    If Not IsNull(rs!no_pekerja) Then 'No. Pekerja
        Frm102_LM_No_PEKERJA = rs!no_pekerja
    End If
    
End If

rs.Close
Set rs = Nothing

'Kadar caj GST
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 51_voucher_belian_agen where no_invoice='" & Frm102.L23_Text & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Not IsNull(rs!kadar_gst) Then Frm102.L21_Text = rs!kadar_gst 'Kadar Cukai GST (%)
    
End If

rs.Close
Set rs = Nothing

'### Maklumat asas bagi invoice ### - End
    
'### Masukkan Data Jualan Ke Dalam Table Jualan (Temp) ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 23_senarai_jualan where no_resit='" & Frm102.L23_Text & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set rs1 = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs1.Open "select * from " & G_JUALAN_TEMP & "", cn, adOpenKeyset, adLockOptimistic

    rs1.AddNew
    If Not IsNull(rs!ID) Then rs1!id_database = rs!ID 'No. ID
    If Not IsNull(rs!no_siri_Produk) Then
        rs1!no_siri_Produk = rs!no_siri_Produk 'No. Siri Produk
    Else
        rs1!no_siri_Produk = Null 'No. Siri Produk
    End If
    If Not IsNull(rs!kategori_Produk) Then
        rs1!kategori_Produk = rs!kategori_Produk 'Kategori Produk
    Else
        rs1!no_siri_Produk = Null 'Kategori Produk
    End If
    If Not IsNull(rs!purity) Then
        rs1!purity = rs!purity 'Purity
    Else
        rs1!purity = Null 'Purity
    End If
    If Not IsNull(rs!Berat_Asal) Then
        rs1!Berat_Asal = rs!Berat_Asal 'Berat Asal (g)
    Else
        rs1!Berat_Asal = Null 'Berat Asal (g)
    End If
    If Not IsNull(rs!berat_jualan) Then
        rs1!berat_jualan = Format(rs!berat_jualan, "0.00") 'Berat Jualan (g)
    Else
        rs1!berat_jualan = Null 'Berat Jualan (g)
    End If
    If Not IsNull(rs!harga_Semasa) Then
        rs1!harga_Semasa = Format(rs!harga_Semasa, "0.00") 'Harga Semasa (RM/g)
    Else
        rs1!harga_Semasa = Null 'Harga Semasa (RM/g)
    End If
    If Not IsNull(rs!UPAH) Then
        rs1!UPAH = Format(rs!UPAH, "0.00") 'Upah (RM)
    Else
        rs1!UPAH = Null 'Upah (RM)
    End If
    If Not IsNull(rs!harga_asal) Then
        rs1!harga_asal = Format(rs!harga_asal, "0.00") 'Harga Asal Item (RM)
    Else
        rs1!harga_asal = Null 'Harga Asal Item (RM)
    End If
    If Not IsNull(rs!diskaun) Then
        rs1!diskaun = Format(rs!diskaun, "0.00") 'Diskaun (%)
    Else
        rs1!diskaun = Null 'Diskaun (%)
    End If
    If Not IsNull(rs!harga_lepas_diskaun) Then
        rs1!harga_lepas_diskaun = Format(rs!harga_lepas_diskaun, "0.00") 'Harga Selepas Diskaun (RM)
    Else
        rs1!harga_lepas_diskaun = Null 'Harga Selepas Diskaun (RM)
    End If
    If Not IsNull(rs!adjustment) Then
        rs1!adjustment = Format(rs!adjustment, "0.00") 'Harga Selepas Diskaun (RM)
    Else
        rs1!adjustment = Null 'Harga Selepas Diskaun (RM)
    End If
    If Not IsNull(rs!harga_jualan) Then
        rs1!harga_jualan = Format(rs!harga_jualan, "0.00") 'Harga Jualan (RM)
    Else
        rs1!harga_jualan = Null 'Harga Jualan (RM)
    End If
    If Not IsNull(rs!gst_ari_nashi) Then
        rs1!gst_ari_nashi = rs!gst_ari_nashi '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
    Else
        rs1!gst_ari_nashi = Null '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
    End If
    If Not IsNull(rs!kadar_gst) Then
        rs1!kadar_gst = rs!kadar_gst 'Kadar Cukai GST (%)
    Else
        rs1!kadar_gst = Null 'Kadar Cukai GST (%)
    End If
    If Not IsNull(rs!jumlah_gst) Then
        rs1!jumlah_gst = Format(rs!jumlah_gst, "0.00") 'Jumlah Cukai GST (RM)
    Else
        rs1!jumlah_gst = Null 'Jumlah Cukai GST (RM)
    End If
    If Not IsNull(rs!harga_dengan_gst) Then
        rs1!harga_dengan_gst = Format(rs!harga_dengan_gst, "0.00") 'Harga Jualan Termasuk GST (RM)
    Else
        rs1!harga_dengan_gst = Null 'Harga Jualan Termasuk GST (RM)
    End If
    If Not IsNull(rs!dropship) Then
        rs1!dropship = rs!dropship '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
    Else
        rs1!dropship = Null '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
    End If
    If Not IsNull(rs!komisyen_per_gram) Then
        rs1!komisyen_per_gram = Format(rs!komisyen_per_gram, "0.00") 'Komisyen Per Gram Dropship (RM/g) : BK
    Else
        rs1!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g) : BK
    End If
    If Not IsNull(rs!jumlah_komisyen) Then
        rs1!jumlah_komisyen = Format(rs!jumlah_komisyen, "0.00") 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
    Else
        rs1!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
    End If
    If Not IsNull(rs!harga_per_gram_modal) Then
        rs1!harga_per_gram_modal = Format(rs!harga_per_gram_modal, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
    Else
        rs1!harga_per_gram_modal = Null 'Harga Per Gram Bagi Modal (RM/g)
    End If
    If Not IsNull(rs!modal) Then
        rs1!modal = Format(rs!modal, "0.00") 'Harga Modal (RM)
    Else
        rs1!modal = Null 'Harga Modal (RM)
    End If
    If Not IsNull(rs!untung) Then
        rs1!untung = Format(rs!untung, "0.00") 'Jumlah Keuntungan
    Else
        rs1!untung = Null 'Jumlah Keuntungan
    End If
    If Not IsNull(rs!harga_per_gram_supplier) Then
        rs1!harga_per_gram_supplier = Format(rs!harga_per_gram_supplier, "0.00") 'Harga per gram (harga semasa) dari supplier (modal)
    Else
        rs1!harga_per_gram_supplier = Null 'Harga per gram (harga semasa) dari supplier (modal)
    End If
    If Not IsNull(rs!upah_modal) Then
        rs1!upah_modal = Format(rs!upah_modal, "0.00") 'Upah modal
    Else
        rs1!upah_modal = Null 'Upah modal
    End If
    If Not IsNull(rs!untung2) Then
        rs1!untung2 = Format(rs!untung2, "0.00") 'Jumlah Keuntungan
    Else
        rs1!untung2 = Null 'Jumlah Keuntungan
    End If
    If Not IsNull(rs!dulang) Then
        rs1!dulang = rs!dulang 'Dulang
    Else
        rs1!dulang = Null 'Dulang
    End If
    If Not IsNull(rs!potong_flag) Then
        rs1!potong_flag = rs!potong_flag '0 : Tiada Potong , 1 : Ada Potong
    Else
        rs1!potong_flag = Null '0 : Tiada Potong , 1 : Ada Potong
    End If
    If Not IsNull(rs!Type) Then
        rs1!Type = rs!Type '0 : BK , 1 : Barang Permata
    Else
        rs1!Type = Null '0 : BK , 1 : Barang Permata
    End If
    'If Frm28.L5_Text <> vbNullString Then
    '    rs1!no_rujukan_pembeli = Frm28.L5_Text 'No. Rujukan Pembeli
    'Else
    '    rs1!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
    'End If
    'If Frm27.L5_Text <> vbNullString Then
    '    rs1!no_rujukan_agen_dropship = Frm27.L5_Text 'No. Rujukan Agen Dropship
    'Else
    '    rs1!no_rujukan_agen_dropship = Null 'No. Rujukan Agen Dropship
    'End If
    
'1:  Pelanggan
'2:  Member
'3:  RAF
'4:  Pengedar
'5:  Normal Dealer
'6:  Master Dealer

    'If Frm102.CB4 = 1 Then
    '    rs1!kategori_pembeli = 1
    'ElseIf Frm102.CB5 = 1 Then
    '    rs1!kategori_pembeli = 2
    'ElseIf Frm102.CB6 = 1 Then
    '    rs1!kategori_pembeli = 4
    'ElseIf Frm102.CB9 = 1 Then
    '    rs1!kategori_pembeli = 3
    'ElseIf Frm102.CB10 = 1 Then
    '    rs1!kategori_pembeli = 5
    'ElseIf Frm102.CB11 = 1 Then
    '    rs1!kategori_pembeli = 6
    'End If
    
    If Not IsNull(rs!gst_include) Then 'Pilihan Cukai GST (SR) Samada Pelanggan Bayar Atau Kedai Bayar
        If rs!gst_include = "**Harga Termasuk GST" Then
            rs1!gst_include = 1
        Else
            rs1!gst_include = 0
        End If
    Else
        rs1!gst_include = 0
    End If
    If Not IsNull(rs!harga_tanpa_gst) Then
        rs1!harga_tanpa_gst = Format(rs!harga_tanpa_gst, "0.00") 'Harga Semasa (RM/g)
    Else
        rs1!harga_tanpa_gst = Null 'Harga Semasa (RM/g)
    End If

'### Maklumat tetapan harga jualan kepada staff ### - Start
    If Not IsNull(rs!kadar_penurunan_upah) Then 'Kadar peratusan penurunan harga upah kepada staff (%)
        rs1!kadar_penurunan_upah = Format(rs!kadar_penurunan_upah, "0.00")
    Else
        rs1!kadar_penurunan_upah = Null
    End If
    If Not IsNull(rs!harga_semasa_staff) Then 'Harga emas semasa yang dijual kepada staff
        rs1!harga_semasa_staff = Format(rs!harga_semasa_staff, "0.00")
    Else
        rs1!harga_semasa_staff = Null
    End If
    If Not IsNull(rs!kadar_penurunan_bp) Then 'Kadar peratusan penurunan harga barang permata kepada staff (%)
        rs1!kadar_penurunan_bp = Format(rs!kadar_penurunan_bp, "0.00")
    Else
        rs1!kadar_penurunan_bp = Null
    End If
    If Not IsNull(rs!harga_staff) Then 'Harga yang dijual kepada staff (RM)
        rs1!harga_staff = Format(rs!harga_staff, "0.00")
    Else
        rs1!harga_staff = Null
    End If
    If Not IsNull(rs!harga_bp_asal) Then 'Tetapan harga barang permata yang asal (RM)
        rs1!harga_bp_asal = Format(rs!harga_bp_asal, "0.00")
    Else
        rs1!harga_bp_asal = Null
    End If
    If Not IsNull(rs!upah_asal) Then 'Tetapan upah asal (RM)
        rs1!upah_asal = Format(rs!upah_asal, "0.00")
    Else
        rs1!upah_asal = Null
    End If
    If Not IsNull(rs!komisyen_staff) Then 'Tetapan upah asal (RM)
        rs1!komisyen_staff = Format(rs!komisyen_staff, "0.00")
    Else
        rs1!komisyen_staff = Null
    End If
'### Maklumat tetapan harga jualan kepada staff ### - End

    rs1!Status = 2
    If Not IsNull(rs!pemalar_tukaran_999) Then 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
        rs1!pemalar_tukaran_999 = rs!pemalar_tukaran_999
    Else
        rs1!pemalar_tukaran_999 = Null
    End If
    If Not IsNull(rs!berat_999) Then 'Berat jualan dalam purity 999.9
        rs1!berat_999 = Format(rs!berat_999, "0.00")
    Else
        rs1!berat_999 = Null
    End If

    If Not IsNull(rs!gst_barang_atau_upah) Then '0 : GST pada harga jualan , 1 : GST pada upah
        rs1!gst_barang_atau_upah = rs!gst_barang_atau_upah
    Else
        rs1!gst_barang_atau_upah = 0
    End If
    If Not IsNull(rs!harga_jualan_dengan_gst) Then
        rs1!harga_jualan_dengan_gst = rs!harga_jualan_dengan_gst
    Else
        rs1!harga_jualan_dengan_gst = 0
    End If

    rs1.Update
    
    rs1.Close
    Set rs1 = Nothing

    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'### Masukkan Data Jualan Ke Dalam Table Jualan (Temp) ### - End

Call Frm102_Senarai_Jualan_Header
Call Frm102_Senarai_Jualan

'### Masukkan data belian barang dari agen ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 50_belian_emas_agen where no_invoice='" & Frm102.L23_Text & "' AND status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set rs1 = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs1.Open "select * from 49_belian_temp", cn, adOpenKeyset, adLockOptimistic

    rs1.AddNew
    If Not IsNull(rs!ID) Then rs1!id_database = rs!ID 'No. ID
    If Not IsNull(rs!Berat_Asal) Then rs1!Berat_Asal = rs!Berat_Asal 'Berat asal barang
    If Not IsNull(rs!purity) Then rs1!purity = rs!purity 'Purity barang
    If Not IsNull(rs!kod_Purity) Then rs1!kod_Purity = rs!kod_Purity 'Kod purity barang
    If Not IsNull(rs!kadar_tukaran) Then rs1!kadar_tukaran = rs!kadar_tukaran 'Kadar tukaran kepada purity 999.9
    If Not IsNull(rs!berat_tukaran) Then rs1!berat_tukaran = rs!berat_tukaran 'Berat setelah ditukar kepada purity 999.9
    rs1!Status = 2
    
    rs1.Update
    
    rs1.Close
    Set rs1 = Nothing
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'### Masukkan data belian barang dari agen ### - End

Call Frm102_senarai_belian_header
Call Frm102_senarai_belian

'### Masukkan data voucher / invoice bagi belian agen ini ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 51_voucher_belian_agen where no_invoice='" & Frm102.L23_Text & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!no_voucher) Then
        Frm102.L22_Text = rs!no_voucher 'No. Voucher
        Frm102.L45_Text = 1 'Flag bagi jika ada pengeluaran voucher bagi urusan ini , 0 : Tiada voucher / Tiada history pengeluaran voucher , 1 : Ada voucher / Ada history pengeluaran voucher
        Frm102.L34_Text.Visible = True
    Else
        Frm102.L34_Text.Visible = False
        Frm102.L45_Text = 0 'Flag bagi jika ada pengeluaran voucher bagi urusan ini , 0 : Tiada voucher / Tiada history pengeluaran voucher , 1 : Ada voucher / Ada history pengeluaran voucher
    End If
    If Not IsNull(rs!berat_jualan) Then 'Berat jualan keseluruhan barang kedai
        Frm102.L9_Text = Format(rs!berat_jualan, "#,##0.00")
    Else
        Frm102.L9_Text = "0.00"
    End If
    If Not IsNull(rs!berat_belian) Then 'Berat belian keseluruhan (Barang trade in)
        Frm102.L10_Text = Format(rs!berat_belian, "#,##0.00")
    Else
        Frm102.L10_Text = "0.00"
    End If
    If Not IsNull(rs!beza_berat) Then 'Beza antara berat jualan dan belian
        Frm102.L11_Text = Format(rs!beza_berat, "#,##0.00")
    Else
        Frm102.L11_Text = "0.00"
    End If
    If Not IsNull(rs!harga_Semasa) Then 'Harga semasa (penilaian harga emas oleh pihak kedai)
        Frm102.TB11 = Format(rs!harga_Semasa, "#,##0.00")
    Else
        Frm102.TB11 = "0.00"
    End If
    If Not IsNull(rs!harga_emas) Then 'Nilaian harga emas oleh pihak kedai terhadap beza berat tersebut (jika bayaran perlu dibuat oleh pihak kedai sahaja)
        Frm102.L12_Text = Format(rs!harga_emas, "#,##0.00")
    Else
        Frm102.L12_Text = "0.00"
    End If
    If Not IsNull(rs!harga_tanpa_gst) Then 'Harga emas tanpa GST
        Frm102.L31_Text = Format(rs!harga_tanpa_gst, "#,##0.00")
    Else
        Frm102.L31_Text = "0.00"
    End If
    If Not IsNull(rs!gst_ari_nashi) Then
        If rs!gst_ari_nashi = "SR" Then
            Frm102.CB7 = 1
        ElseIf rs!gst_ari_nashi = "ZR (L)" Then
            Frm102.CB5 = 1
            Frm102.CB6 = 0
            Frm102.CB7 = 0
        End If
    Else
        Frm102.CB5 = 1
        Frm102.CB6 = 0
        Frm102.CB7 = 0
    End If
    If Not IsNull(rs!gst_include) Then
        If rs!gst_include = 0 Then '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
            Frm102.CB7 = 0
        ElseIf rs!gst_include = 1 Then
            Frm102.CB7 = 1
        End If
    End If
    If Not IsNull(rs!kadar_gst) Then 'Kadar Cukai GST (%)
        Frm102.L21_Text = rs!kadar_gst
    Else
        Frm102.L21_Text = "0.00"
    End If
    If Not IsNull(rs!harga_tanpa_gst) Then 'Harga emas tanpa GST
        Frm102.L31_Text = Format(rs!harga_tanpa_gst, "#,##0.00")
    Else
        Frm102.L31_Text = "0.00"
    End If
    If Not IsNull(rs!jumlah_gst) Then 'Jumlah Cukai GST (RM)
        Frm102.TB12 = Format(rs!jumlah_gst, "#,##0.00")
    Else
        Frm102.TB12 = "0.00"
    End If
    If Not IsNull(rs!harga_dengan_gst) Then 'Jumlah emas + GST (RM)
        Frm102.TB13 = Format(rs!harga_dengan_gst, "#,##0.00")
    Else
        Frm102.TB13 = "0.00"
    End If

    rs.Update
End If

rs.Close
Set rs = Nothing
'### Masukkan data voucher / invoice bagi belian agen ini ### - End

'###Masukkan akaun bagi jualan ### -Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where no_resit='" & Frm102.L23_Text & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Not IsNull(rs!harga_barang) Then Frm102.L15_Text = rs!harga_barang 'Jumlah Harga Barang Tanpa GST (RM)
    If Not IsNull(rs!gst_zr_harga) Then Frm102.L17_Text = rs!gst_zr_harga 'Harga Keseluruhan Bagi Barang ZR
    If Not IsNull(rs!gst_zr_cukai) Then Frm102.L19_Text = rs!gst_zr_cukai 'Jumlah Cukai Bagi ZR
    If Not IsNull(rs!gst_sr_harga) Then Frm102.L18_Text = rs!gst_sr_harga 'Harga Keseluruhan Bagi Barang SR
    If Not IsNull(rs!gst_sr_cukai) Then Frm102.L20_Text = rs!gst_sr_cukai 'Jumlah Cukai Bagi SR
    If Not IsNull(rs!harga_barang_dengan_gst) Then Frm102.L16_Text = rs!harga_barang_dengan_gst 'Jumlah Harga Barang Dengan GST (RM)
    If Not IsNull(rs!tunai) Then Frm102.TB14 = rs!tunai 'Cara Bayaran : Tunai
    If Not IsNull(rs!bank_in) Then Frm102.TB15 = rs!bank_in 'Cara Bayaran : Bank In
    If Not IsNull(rs!kad_kredit) Then Frm102.TB16 = rs!kad_kredit 'Cara Bayaran : Kad Kredit
    If Not IsNull(rs!jumlah_cas_kad_kredit) Then Frm102.TB17 = rs!jumlah_cas_kad_kredit 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
    If Not IsNull(rs!jumlah_potongan_kad_kredit) Then Frm102.TB18 = rs!jumlah_potongan_kad_kredit 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
    If Not IsNull(rs!duit_simpanan_kedai) Then Frm102.TB22 = rs!duit_simpanan_kedai 'Cara Bayaran : Simpanan Duit Di Kedai
    If Not IsNull(rs!kad_debit) Then Frm102.TB19 = rs!kad_debit 'Cara Bayaran : Kad Debit
    If Not IsNull(rs!jumlah_cas_kad_debit) Then Frm102.TB20 = rs!jumlah_cas_kad_debit 'Cara Bayaran : Jumlah Cas Kad Debit (RM)
    If Not IsNull(rs!jumlah_potongan_kad_debit) Then Frm102.TB21 = rs!jumlah_potongan_kad_debit 'Cara Bayaran : Jumah Potongan Kad Debit (RM)
    If Not IsNull(rs!jumlah_bayaran) Then Frm102.TB23 = rs!jumlah_bayaran 'Cara Bayaran : Jumlah Bayaran

End If

rs.Close
Set rs = Nothing
'###Masukkan akaun bagi jualan ### - End

'### Maklumat Agen ### - Start
If Frm102_LM_No_PEMBELI <> vbNullString Then
    
    Call Frm28_initial
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm102_LM_No_PEMBELI & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!Nama) Then
            Frm28.L1_Text = rs!Nama 'Nama
            'Frm102.L46_Text = rs!Nama
        End If
        If Not IsNull(rs!no_ic) Then Frm28.L2_Text = rs!no_ic 'No. Kad Pengenalan
        If Not IsNull(rs!no_tel) Then Frm28.L3_Text = rs!no_tel 'No. Telefon
        If Not IsNull(rs!Email) Then Frm28.L4_Text = rs!Email 'E-mail
        If Not IsNull(rs!no_pelanggan) Then Frm28.L5_Text = rs!no_pelanggan 'No. Pelanggan
    End If
    
    rs.Close
    Set rs = Nothing

End If
'### Maklumat Agen ### - End

'### Carian Maklumat Penjual (Data Pekerja) ### - Start
If Frm102_LM_No_PEKERJA <> vbNullString Then

    DATA_PEKERJA_FOUND = 0
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where NoPekerja='" & Frm102_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        Frm102_LM_MAKLUMAT_PEKERJA = rs!Samaran & "  |  " & rs!NoPekerja
        DATA_PEKERJA_FOUND = 1
        
    End If
    
    rs.Close
    Set rs = Nothing
    
    If DATA_PEKERJA_FOUND = 1 Then
        On Error GoTo Err_A:
        Frm102.CBB4 = Frm102_LM_MAKLUMAT_PEKERJA
Restore_A:
    End If
    
    'on error resume next
End If
'### Carian Maklumat Penjual (Data Pekerja) ### - End

GLOBAL_DISABLE = 0

Frm102.CBB4.Enabled = True
Frm102.CBB4.BackColor = &HFFFFFF

Frm102.Show
Frm85.Hide

Exit Sub
Err_A:
Frm102.CBB4.AddItem Frm102_LM_MAKLUMAT_PEKERJA
Frm102.CBB4 = Frm102_LM_MAKLUMAT_PEKERJA
Resume Restore_A:
End Sub
Sub Frm102_cetak_invoice()
'on error resume next
Frm102_LM_No_CUST = vbNullString

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

Report56.Sections("Section4").Controls("L1").Caption = vbNullString 'No. Invoice
Report56.Sections("Section4").Controls("L2").Caption = vbNullString 'Tarikh
Report56.Sections("Section4").Controls("L3").Caption = vbNullString 'Nama Pembeli
Report56.Sections("Section4").Controls("L4").Caption = vbNullString 'No. Telefon
Report56.Sections("Section5").Controls("L5").Caption = vbNullString 'Lain-lain perkara
Report56.Sections("Section5").Controls("L6").Caption = vbNullString 'Harga lain-lain
Report56.Sections("Section5").Controls("L7").Caption = vbNullString 'Jenis GST lain-lain
Report56.Sections("Section5").Controls("L8").Caption = "0.00" 'Jumlah harga SR
Report56.Sections("Section5").Controls("L9").Caption = "0.00" 'Jumlah harga ZR
Report56.Sections("Section5").Controls("L10").Caption = "0.00" 'Jumlah GST SR
Report56.Sections("Section5").Controls("L11").Caption = "0.00" 'Jumlah GST ZR
Report56.Sections("Section5").Controls("L12").Caption = "0.00" 'Jumlah
Report56.Sections("Section5").Controls("L13").Caption = "0.00" 'Jumlah GST
Report56.Sections("Section5").Controls("L14").Caption = "0.00" 'Jumlah keseluruhan
Report56.Sections("Section5").Controls("L15").Caption = "0" 'Bilangan barang
Report56.Sections("Section5").Controls("L16").Caption = "0.00" 'Jumlah berat (g)

'### Reset maklumat kedai ### - Start
Report56.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report56.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report56.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report56.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report56.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

Report56.Sections("Section4").Controls("L205").Caption = "Goods Despatch Note"

If LM_HEADER = 1 Then '0 : Pre Printed , 1 : Sistem
    '### Maklumat kedai ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!nama_kedai) Then Report56.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
        If Not IsNull(rs!no_pendaftaran) Then Report56.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
        If Not IsNull(rs!alamat) Then Report56.Sections("Section4").Controls("L202").Caption = rs!alamat
        If Not IsNull(rs!no_tel) Then Report56.Sections("Section4").Controls("L203").Caption = rs!no_tel
        If Not IsNull(rs!no_id_gst) Then Report56.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
    End If
    
    rs.Close
    Set rs = Nothing
    '### Maklumat kedai ### - End
    
    Report56.Sections("Section4").Visible = True
Else
    Report56.Sections("Section4").Visible = False
End If


Report56.Sections("Section4").Controls("L1").Caption = G_No_RESIT_JUALAN 'No. Invoice

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 51_voucher_belian_agen where no_invoice='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!flag_bayaran) Then
        If rs!flag_bayaran = 0 Then
            If Not IsNull(rs!beza_berat) And Not IsNull(rs!gst_ari_nashi) And Not IsNull(rs!harga_dengan_gst) Then  'Beza antara berat jualan dan belian
                Report56.Sections("Section5").Controls("L5").Caption = "Emas purity 999.9 dengan berat " & Format(rs!beza_berat, "#,##0.00 g") 'Lain-lain perkara
                Report56.Sections("Section5").Controls("L6").Caption = Format(rs!harga_dengan_gst, "#,##0.00") 'Harga lain-lain
                Report56.Sections("Section5").Controls("L7").Caption = rs!gst_ari_nashi 'Jenis GST lain-lain
            Else
                MsgBox "Telah berlaku sedikit kekeliruan bagi invoice ini." & vbCrLf & _
                        vbNullString & vbCrLf & _
                        "Sila hubungi pihak Sankyu System bagi membetulkan kekeliruan ini ," & vbCrLf & _
                        "dan sila nyatakan No. Invoice [" & G_No_RESIT_JUALAN & "] kepada pihak mereka.", vbCritical, "Error"
                        
                Exit Sub
            End If
        End If
    End If
End If

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where no_resit='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!write_timestamp) Then
        Report56.Sections("Section4").Controls("L2").Caption = rs!write_timestamp 'Tarikh
    Else
        Report56.Sections("Section4").Controls("L2").Caption = rs!tarikh 'Tarikh
    End If
    If Not IsNull(rs!gst_sr_harga) Then Report56.Sections("Section5").Controls("L8").Caption = Format(rs!gst_sr_harga, "#,##0.00") 'Jumlah harga SR
    If Not IsNull(rs!gst_zr_harga) Then Report56.Sections("Section5").Controls("L9").Caption = Format(rs!gst_zr_harga, "#,##0.00") 'Jumlah harga ZR
    If Not IsNull(rs!gst_sr_cukai) Then Report56.Sections("Section5").Controls("L10").Caption = Format(rs!gst_sr_cukai, "#,##0.00") 'Jumlah GST SR
    If Not IsNull(rs!gst_zr_cukai) Then Report56.Sections("Section5").Controls("L11").Caption = Format(rs!gst_zr_cukai, "#,##0.00") 'Jumlah GST ZR
    If Not IsNull(rs!harga_barang) Then Report56.Sections("Section5").Controls("L12").Caption = Format(rs!harga_barang, "#,##0.00") 'Jumlah RM
    If Not IsNull(rs!jumlah_cukai_gst) Then Report56.Sections("Section5").Controls("L13").Caption = Format(rs!jumlah_cukai_gst, "#,##0.00") 'Jumlah GST
    If Not IsNull(rs!harga_barang_dengan_gst) Then Report56.Sections("Section5").Controls("L14").Caption = Format(rs!harga_barang_dengan_gst, "#,##0.00") 'Jumlah keseluruhan
    If Not IsNull(rs!no_rujukan_pembeli) Then Frm102_LM_No_CUST = rs!no_rujukan_pembeli
    If Not IsNull(rs!kuantiti_barang) Then Report56.Sections("Section5").Controls("L15").Caption = rs!kuantiti_barang 'Bilangan barang
    If Not IsNull(rs!JUMLAH_BERAT) Then Report56.Sections("Section5").Controls("L16").Caption = Format(rs!JUMLAH_BERAT, "#,##0.00") 'Jumlah berat (g)
End If

rs.Close
Set rs = Nothing

'### Data jika pembeli adalah berdaftar ### - Start
If Frm102_LM_No_CUST <> vbNullString Then
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm102_LM_No_CUST & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!Nama) Then Report56.Sections("Section4").Controls("L3").Caption = rs!Nama 'Maklumat Pembeli : Nama
        If Not IsNull(rs!no_tel) Then Report56.Sections("Section4").Controls("L4").Caption = rs!no_tel 'Maklumat Pembeli : No. Telefon
    End If
    
    rs.Close
    Set rs = Nothing
End If
'### Data jika pembeli adalah berdaftar ### - End
            

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 23_senarai_jualan where no_resit='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report56.DataSource = rs
    Report56.Show
    
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing

End Sub
Sub Frm102_cetak_voucher()
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

Frm101_LM_NO_INVOICE = G_No_RESIT_JUALAN
Frm102_LM_No_CUST = vbNullString
Report57.Sections("Section4").Controls("L1").Caption = vbNullString 'No. Invoice
Report57.Sections("Section4").Controls("L2").Caption = vbNullString 'Tarikh
Report57.Sections("Section4").Controls("L3").Caption = vbNullString 'Nama Pembeli
Report57.Sections("Section4").Controls("L4").Caption = vbNullString 'No. Telefon
Report57.Sections("Section4").Controls("L5").Caption = vbNullString 'Caption : Voucher / Statement

'Report57.Sections("Section2").Controls("L1").Caption = Frm101_LM_NO_INVOICE 'No. Invoice
Report57.Sections("Section5").Controls("L8").Caption = "*** Sila rujuk invoice " & Frm101_LM_NO_INVOICE & " bagi maklumat jualan terperinci."

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 51_voucher_belian_agen where no_invoice='" & Frm101_LM_NO_INVOICE & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!berat_belian) Then Report57.Sections("Section5").Controls("L6").Caption = Format(rs!berat_belian, "#,##0.00") 'Berat jualan emas oleh agen kepada kedai
    If Not IsNull(rs!write_timestamp) Then Report57.Sections("Section4").Controls("L1").Caption = rs!write_timestamp 'Tarikh & Masa
    
    If Not IsNull(rs!flag_bayaran) Then
        If rs!flag_bayaran = 1 Then
            If Not IsNull(rs!harga_emas) Then Report57.Sections("Section5").Controls("L7").Caption = "Jumlah yang perlu dibayar kepada agen / penjual : RM " & Format(rs!harga_emas, "#,##0.00")
            Report57.Sections("Section5").Controls("L7").Visible = True
            Report57.Sections("Section4").Controls("L5").Caption = "Voucher"
            Report57.Sections("Section4").Controls("L9").Visible = True 'Caption : No. Voucher
            Report57.Sections("Section4").Controls("L10").Visible = True 'Caption : [:]
            If Not IsNull(rs!no_voucher) Then Report57.Sections("Section4").Controls("L2").Caption = rs!no_voucher 'No. Voucher / Statement
        Else
            Report57.Sections("Section5").Controls("L7").Visible = False
            Report57.Sections("Section4").Controls("L5").Caption = "Statement"
            Report57.Sections("Section4").Controls("L9").Visible = False 'Caption : No. Voucher
            Report57.Sections("Section4").Controls("L10").Visible = False 'Caption : [:]
        End If
    Else
        Report57.Sections("Section5").Controls("L7").Visible = False
        Report57.Sections("Section2").Controls("L5").Caption = "Voucher"
    End If
    
End If

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where no_resit='" & Frm101_LM_NO_INVOICE & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!no_rujukan_pembeli) Then Frm102_LM_No_CUST = rs!no_rujukan_pembeli
End If

rs.Close
Set rs = Nothing

'### Data jika pembeli adalah berdaftar ### - Start
If Frm102_LM_No_CUST <> vbNullString Then
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm102_LM_No_CUST & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!Nama) Then Report57.Sections("Section4").Controls("L3").Caption = rs!Nama 'Maklumat Pembeli : Nama
        If Not IsNull(rs!no_tel) Then Report57.Sections("Section4").Controls("L4").Caption = rs!no_tel 'Maklumat Pembeli : No. Telefon
    End If
    
    rs.Close
    Set rs = Nothing
End If
'### Data jika pembeli adalah berdaftar ### - End
            
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 50_belian_emas_agen where no_invoice='" & Frm101_LM_NO_INVOICE & "' AND status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic
         
While rs.EOF = False
    Set Report57.DataSource = rs
    Report57.Show
    
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing

End Sub
Sub Frm85_padam_voucher()
'on error resume next
Dim rs2 As ADODB.Recordset
Dim Frm85_LM_BERAT_ASAL As Double
Dim Frm85_LM_BEZA_BERAT As Double
Dim Frm85_LM_BERAT_JUALAN As Double

'### Padam data berkenaan akaun invoice ini ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where no_resit='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    rs.Delete
    rs.Update
End If

rs.Close
Set rs = Nothing
'### Padam data berkenaan akaun invoice ini ### - End

'### Pulangkan stok barang kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 23_senarai_jualan where no_resit='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Frm85_LM_BERAT_ASAL = 0
    Frm85_LM_BEZA_BERAT = 0
    Frm85_LM_BERAT_JUALAN = 0
    
    Set rs2 = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs2.Open "select * from data_database where no_siri_Produk='" & rs!no_siri_Produk & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs2.EOF Then
        If rs2!receiving_Status = 0 Or rs2!receiving_Status = 2 Then
            If Not IsNull(rs2!Berat) Then Frm85_LM_BERAT_ASAL = rs2!Berat
            If Not IsNull(rs2!beza_berat) Then Frm85_LM_BEZA_BERAT = rs2!beza_berat
            If Not IsNull(rs!berat_jualan) Then Frm85_LM_BERAT_JUALAN = rs!berat_jualan
            
            If Frm85_LM_BERAT_ASAL = Frm85_LM_BERAT_JUALAN + Frm85_LM_BEZA_BERAT Then
                rs2!beza_berat = Format(Frm85_LM_BERAT_JUALAN + Frm85_LM_BEZA_BERAT, "0.00")
                rs2!StatusItem = 10
            Else
                rs2!beza_berat = Format(Frm85_LM_BERAT_JUALAN + Frm85_LM_BEZA_BERAT, "0.00")
                rs2!StatusItem = 12
            End If
        Else
            rs2!StatusItem = 10
        End If
        rs2.Update
    End If
    
    rs2.Close
    Set rs2 = Nothing
    
    rs.Delete
    rs.Update
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'### Pulangkan stok barang kedai ### - End

'### Tukar status barangan yang telah di trade in oleh agen ### - Start
strsql = "UPDATE 50_belian_emas_agen set status='" & 0 & "'," _
& "write_timestamp2='" & Now & "'" _
& "WHERE no_invoice='" & G_No_RESIT_JUALAN & "'"

Set rs = cn.Execute(strsql)
Set rs = Nothing
'### Tukar status barangan yang telah di trade in oleh agen ### - End

'### Padam voucher / statement berkenaan barangan yang di trade in oleh agen ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 51_voucher_belian_agen where no_invoice='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    rs.Delete
    rs.Update
End If

rs.Close
Set rs = Nothing
'### Padam voucher / statement berkenaan barangan yang di trade in oleh agen ### - End

'#### Update Log Aktiviti Sistem #### - Start
user = MDI_frm1.L3_Text
LogAct_Memory = "[" & user & "] Padam invoice jualan kepada agen. No. Invoice [" & G_No_RESIT_JUALAN & "]."
LogDate_Memory = DateTime.Date & " " & DateTime.Time$
Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End

Note = "Data Telah Berjaya Dipadamkan." & vbCrLf & _
        "Refresh Data Anda ? Sistem Akan Mengambil Sedikit Masa Untuk Refresh Data." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"


Answer = MsgBox(Note, vbOK, "Confirmation")

If Answer = vbOK Then
    GM_NEXT_PREV = 2
    
    Call Frm85_Header_Report_Jualan
    Call Frm85_Report_Jualan_page
Else
    GM_NEXT_PREV = 2
    
    Call Frm85_Header_Report_Jualan
    Call Frm85_Report_Jualan_page
End If
End Sub
Sub Frm102_jurujual()
'On Error Resume Next
If MDI_frm1.L3_Text <> vbNullString Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        Frm102.CBB4 = rs!Samaran & "  |  " & rs!NoPekerja
        
    Else
        
        Frm102.CBB4.AddItem "" & "  |  " & rs!Samaran
        Frm102.CBB4 = "" & "  |  " & rs!Samaran
        
    End If
    
    rs.Close
    Set rs = Nothing
    
    If G_LOCK_JURUJUAL = "YES" Then
    
        Frm102.CBB4.Enabled = False
        Frm102.CBB4.BackColor = &H8000000A

    Else
    
        Frm102.CBB4.Enabled = True
        Frm102.CBB4.BackColor = &HFFFFFF

    End If

End If
End Sub
