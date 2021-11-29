Attribute VB_Name = "Module15"
Sub Frm115_reset_main2()
'on error resume next
'### Digunakan untuk reset paparan / komponen pada setiap kali penjualan barang atau pembatalan jualan
Frm115.L3_Text = vbNullString 'No. siri produk
Frm115.L4_Text = vbNullString 'Purity
Frm115.L5_Text = vbNullString 'Kategori Produk
Frm115.L6_Text = "0.00" 'Berat asal
Frm115.L7_Text = "0.00" 'Berat jualan dalam purity 999.9

Frm115.L9_Text = "0.00" 'Berat jualan 999.9
Frm115.L12_Text = "0.00" 'Harga emas

Frm115.L15_Text = "0.00" 'Maklumat GST : Jumlah harga tanpa GST
Frm115.L16_Text = "0.00" 'Maklumat GST : Jumlah harga dengan GST
Frm115.L17_Text = "0.00" 'Maklumat GST : Jumlah harga ZR
Frm115.L18_Text = "0.00" 'Maklumat GST : Jumlah harga SR
Frm115.L19_Text = "0.00" 'Maklumat GST : Jumlah GST ZR
Frm115.L20_Text = "0.00" 'Maklumat GST : Jumlah GST SR
Frm115.L24_Text = 0 'No. id jualan
Frm115.L25_Text = 0 'No. id trade in

Frm115.TB1 = vbNullString 'No. Siri Produk (Scan)
'Frm115.TB2 = "0.00" 'Harga emas semasa 999.9 (Untuk tujuan jualan kepada pelanggan)
Frm115.TB3 = "0.00" 'Berat jualan (g)
Frm115.TB4 = "0.00" 'Upah
Frm115.TB5 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
Frm115.TB6 = "0.00" 'Jumlah Upah + GST (Bagi jualan setiap item)
Frm115.TB7 = "1.00" 'Kadar tukaran emas (mutu) - Per item
Frm115.TB8 = "1.00" 'Kadar tukaran emas (mutu) - Urusan keseluruhan
Frm115.L51_Text = "0.00" 'Jumlah upah tanpa GST (keseluruhan)
Frm115.L52_Text = "0.00" 'Jumlah GST (Keseluruhan)
Frm115.L53_Text = "0.00" 'Jumlah Upah + GST (Keseluruhan)

'Frm115.TB2 = "0.00" 'Harga emas semasa 999.9 (Overall)

Frm115.CMD1.Visible = True 'Masukkan dalam senarai jualan
Frm115.CMD2.Visible = False 'Masukkan dalam senarai jualan (Edit)
Frm115.CMD3.Visible = False 'Batal edit data
End Sub
Sub Frm115_reset_1()
'on error resume next
'### Digunakan untuk reset paparan / komponen pada setiap kali penjualan barang atau pembatalan jualan

Frm115.L3_Text = vbNullString 'No. siri produk
Frm115.L4_Text = vbNullString 'Purity
Frm115.L5_Text = vbNullString 'Kategori Produk
Frm115.L6_Text = "0.00" 'Berat asal
Frm115.L7_Text = "0.00" 'Berat jualan dalam purity 999.9
Frm115.L24_Text = 0 'No. id jualan
'Frm115.L43_Text = 0 'Jumlah bilangan barang jualan
Frm115.L43_Text.BackStyle = 0 'Jumlah bilangan barang jualan
'Frm115.L48_Text = "0.00" 'Jumlah berat (g)
Frm115.L48_Text.BackStyle = 0 'Jumlah berat (g)
Frm115.L49_Text = "0.00"
Frm115.L50_Text = "0.00"

'Frm115.TB1 = vbNullString 'No. Siri Produk (Scan)
Frm115.TB3 = "0.00" 'Berat jualan (g)
Frm115.TB4 = "0.00" 'Upah
Frm115.TB5 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
Frm115.TB6 = "0.00" 'Jumlah Upah + GST (Bagi jualan setiap item)

Frm115.TB1.Locked = False
Frm115.TB1.BackColor = &HFFFFFF

Frm115.TB3.Locked = False
Frm115.TB3.BackColor = &HFFFFFF

Frm115.CMD1.Visible = True 'Masukkan dalam senarai jualan
Frm115.CMD2.Visible = False 'Masukkan dalam senarai jualan (Edit)
Frm115.CMD3.Visible = False 'Batal edit data
End Sub
Sub Frm115_reset_2()
'on error resume next
'### Digunakan untuk reset paparan / komponen pada setiap kali pembelian barang trade in atau pembatalan pembelian barang trade in

Frm115.L25_Text = 0 'No. id trade in
'Frm115.L44_Text = 0 'Jumlah bilangan barang trade in
Frm115.L47_Text = vbNullString 'Kod purity barang
End Sub
Sub Frm115_reset_3()
'on error resume next
'### Digunakan untuk reset paparan / komponen semua komponen transaksi
Frm115.L9_Text = "0.00" 'Berat jualan 999.9
Frm115.L12_Text = "0.00" 'Harga emas

Frm115.L15_Text = "0.00" 'Maklumat GST : Jumlah harga tanpa GST
Frm115.L16_Text = "0.00" 'Maklumat GST : Jumlah harga dengan GST
Frm115.L17_Text = "0.00" 'Maklumat GST : Jumlah harga ZR
Frm115.L18_Text = "0.00" 'Maklumat GST : Jumlah harga SR
Frm115.L19_Text = "0.00" 'Maklumat GST : Jumlah GST ZR
Frm115.L20_Text = "0.00" 'Maklumat GST : Jumlah GST SR
Frm115.L24_Text = 0 'No. id jualan
Frm115.L25_Text = 0 'No. id trade in

'Frm115.TB2 = "0.00" 'Harga emas semasa 999.9 (Overall)
End Sub
Sub Frm115_reset_main()
'on error resume next
'### Digunakan untuk reset / update komponen dari database setelah penjualan atau pembatalan jualan
'Frm115.TB2 = "0.00" 'Harga emas semasa 999.9 (Untuk tujuan jualan kepada pelanggan)
Frm115.L51_Text = "0.00" 'Jumlah upah tanpa GST (keseluruhan)
Frm115.L52_Text = "0.00" 'Jumlah GST (Keseluruhan)
Frm115.L53_Text = "0.00" 'Jumlah Upah + GST (Keseluruhan)
Frm115.Frame4.Visible = False
Frm115.Frame4.Left = 15360
Frm115.Frame4.Top = 120
Frm115.L45_Text = 0 'Flag bagi jika ada pengeluaran voucher bagi urusan ini , 0 : Tiada voucher / Tiada history pengeluaran voucher , 1 : Ada voucher / Ada history pengeluaran voucher

Frm115.L35_Text = 0
Frm115.L36_Text = 0
Frm115.L37_Text = 0
Frm115.L38_Text = 0
Frm115.L39_Text = 0
Frm115.L40_Text = 0
Frm115.L41_Text = 0
Frm115.L42_Text = 0
Frm115.L54_Text = vbNullString

GLOBAL_DISABLE = 1
If G_SCANNER_MODE = 1 Then 'Tetapan penggunaan scanner
    Frm115.CB1 = 1
Else
    Frm115.CB1 = 0
End If
Frm115.L21_Text = G_RATE_GST 'Jumlah Kadar GST
If G_GST_JUAL = 1 Then 'SR

    Frm115.CB3 = 1
    Frm115.CB2 = 0
    
    If G_GST_JUALAN_INC = 1 Then
    
        Frm115.CB4 = 1
        
    Else
        
        Frm115.CB4 = 0
        
    End If

Else 'ZR

    Frm115.CB2 = 1
    Frm115.CB3 = 0
    Frm115.CB4 = 0
    
End If

Frm115.TB2 = Format(G_HARGA_999, "#,##0.00")
GLOBAL_DISABLE = 0

'Set rs = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

'If Not rs.EOF Then
'    If rs!Default1 = "Default" Then
'        GLOBAL_DISABLE = 1
'        If Not IsNull(rs!ScannerMode) Then 'Tetapan penggunaan scanner
'            If rs!ScannerMode = 1 Then
'                Frm115.CB1 = 1
'            Else
'                Frm115.CB1 = 0
'            End If
'        Else
'            Frm115.CB1 = 0
'        End If
'        If Not IsNull(rs!gst_value) Then Frm115.L21_Text = rs!gst_value 'Jumlah Kadar GST
'        If Not IsNull(rs!gst_arinashi) Then 'Tetapan GST , ZR atau SR
'            If rs!gst_arinashi = 1 Then 'SR
'                Frm115.CB3 = 1
'                Frm115.CB2 = 0
'                If Not IsNull(rs!gst_jualan_included) Then
'                    If rs!gst_jualan_included = 1 Then
'
'                        Frm115.CB4 = 1
'
'                    Else
'
'                        Frm115.CB4 = 0
'
'                    End If
'                End If
'            Else 'ZR
'                Frm115.CB2 = 1
'                Frm115.CB3 = 0
'                Frm115.CB4 = 0
'            End If
'        End If
'        If Not IsNull(rs!NoRujukanSistem) Then Frm115.L29_Text = rs!NoRujukanSistem 'No. Rujukan Sistem
'        If Not IsNull(rs!no_trade_in_agen) Then Frm115.L22_Text = rs!no_trade_in_agen 'No. Voucher Trade In
'        If Not IsNull(rs!no_gdn) Then Frm115.L23_Text = rs!no_gdn 'No. GDN

'        If Not IsNull(rs!harga_999) Then
'            Frm115.TB2 = Format(rs!harga_999, "0.00")
'            Frm115.TB2 = Format(rs!harga_999, "0.00")
'        Else
'            Frm115.TB2 = "0.00"
'            Frm115.TB2 = "0.00"
'        End If

'        GLOBAL_DISABLE = 0
'    End If
'End If

'rs.Close
'Set rs = Nothing

'###Senarai Nama Pekerja###
Frm115.CBB4.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where status='" & "Aktif" & "' AND InvestorSmall = 0 AND InvestorBig = 0", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    If rs!user_level = 1 Or rs!user_level = 2 Or rs!user_level = 3 Or rs!user_level = G_LEVEL_USER Then Frm115.CBB4.AddItem rs!Samaran & "  |  " & rs!NoPekerja
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm115.CBB2.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where status='" & 1 & "' order by supplier ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!supplier) Then Frm115.CBB2.AddItem rs!supplier
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "UPDATE Data_Database set gdn_temp = 0"

Set rs = cn.Execute(strsql)
Set rs = Nothing

'###Padam Table Jualan Temp### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "TRUNCATE TABLE " & G_GDN_TEMP & ""

Set rs = cn.Execute(strsql)
Set rs = Nothing
'###Padam Table Jualan Temp### - End

Call Frm115_jurujual
End Sub
Sub Frm115_calc1()
'On Error Resume Next
Dim Frm115_LM_BERAT As Double
Dim Frm115_LM_KADAR_TUKARAN As Double

Frm115_LM_BERAT = 0 'Berat jualan (g)
Frm115_LM_KADAR_TUKARAN = 0 'Kadar tukaran kepada purity 999.9

If ((Frm115.TB3 <> vbNullString And IsNumeric(Frm115.TB3)) And (Frm115.TB7 <> vbNullString And IsNumeric(Frm115.TB7))) Then

    Frm115_LM_BERAT = Frm115.TB3 'Berat jualan (g)
    Frm115_LM_KADAR_TUKARAN = Frm115.TB7 'Kadar tukaran kepada purity 999.9
    
    Frm115.L7_Text = Format(Frm115_LM_BERAT * Frm115_LM_KADAR_TUKARAN, "#,##0.00") 'Berat 999.9
    
Else

    Frm115.L7_Text = "0.00" 'Berat 999.9
    
End If
End Sub
Sub Frm115_calc2()
'On Error Resume Next
Dim Frm115_LM_KADAR_GST As Double
Dim Frm115_LM_UPAH As Double

Frm115_LM_KADAR_GST = 0
Frm115_LM_UPAH = 0

If IsNumeric(Frm115.L21_Text) Then Frm115_LM_KADAR_GST = Frm115.L21_Text 'Kadar gst (%)
If IsNumeric(Frm115.TB4) Then Frm115_LM_UPAH = Frm115.TB4 'Upah (RM)

If Frm115.L21_Text <> vbNullString And IsNumeric(Frm115.L21_Text) Then

    If Frm115.TB4 <> vbNullString And IsNumeric(Frm115.TB4) Then
    
        If Frm115.CB2 = 1 Then 'Upah : GST ZR
        
            Frm115.L30_Text = Format(Frm115.TB4, "#,##0.00") 'Harga upah tanpa GST
            Frm115.TB5 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
            
        End If
        
        If Frm115.CB3 = 1 Then
        
            Frm115.L30_Text = Format(Frm115_LM_UPAH, "#,##0.00") 'Harga upah tanpa GST
            Frm115.TB5 = Format(Frm115_LM_UPAH * (Frm115_LM_KADAR_GST / 100), "#,##0.00") 'Jumlah GST (Bagi jualan setiap item)
            
        End If
        
        
        If Frm115.CB4 = 1 Then
    
            Frm115.L30_Text = Format(Frm115_LM_UPAH / (1 + (Frm115_LM_KADAR_GST / 100)), "#,##0.00") 'Harga upah tanpa GST
            Frm115.TB5 = Format(Frm115_LM_UPAH - (Frm115_LM_UPAH / (1 + (Frm115_LM_KADAR_GST / 100))), "#,##0.00") 'Jumlah GST (Bagi jualan setiap item)
                
        End If

    Else
    
        Frm115.L30_Text = Format(0, "#,##0.00") 'Harga upah tanpa GST
        Frm115.TB5 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
        
    End If

Else

    If IsNumeric(Frm115.TB4) Then
    
        Frm115.L30_Text = Format(Frm115.TB4, "#,##0.00") 'Harga upah tanpa GST
        Frm115.TB5 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
        
    Else
        
        Frm115.L30_Text = Format(0, "#,##0.00") 'Harga upah tanpa GST
        Frm115.TB5 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
        
    End If
    
End If
End Sub
Sub Frm115_calc3()
'On Error Resume Next
Dim Frm115_LM_UPAH_TANPA_GST As Double
Dim Frm115_LM_GST As Double

Frm115_LM_UPAH_TANPA_GST = 0 'Jumlah upah tanpa GST
Frm115_LM_GST = 0 'Jumlah GST

If ((Frm115.TB5 <> vbNullString And IsNumeric(Frm115.TB5)) And (Frm115.L30_Text <> vbNullString And IsNumeric(Frm115.L30_Text))) Then

    Frm115_LM_GST = Frm115.TB5 'Jumlah GST (Bagi jualan setiap item)
    Frm115_LM_UPAH_TANPA_GST = Frm115.L30_Text 'Harga upah tanpa GST
    
    Frm115.TB6 = Format(Frm115_LM_GST + Frm115_LM_UPAH_TANPA_GST, "#,##0.00") 'Jumlah Upah + GST (Bagi jualan setiap item)
    
Else

    Frm115.TB6 = "0.00" 'Jumlah Upah + GST (Bagi jualan setiap item)
    
End If
End Sub
Sub Frm115_calc4()
'On Error Resume Next
Dim Frm115_LM_BERAT_ASAL_ASAL As Double
Dim Frm115_LM_KADAR_TUKARAN As Double

Frm115_LM_BERAT_ASAL = 0 'Berat jualan (g)
Frm115_LM_KADAR_TUKARAN = 0 'Kadar tukaran kepada purity 999.9

If ((Frm115.L48_Text <> vbNullString And IsNumeric(Frm115.L48_Text)) And (Frm115.TB8 <> vbNullString And IsNumeric(Frm115.TB8))) Then

    Frm115_LM_BERAT_ASAL = Frm115.L48_Text 'Berat jualan (g)
    Frm115_LM_KADAR_TUKARAN = Frm115.TB8 'Kadar tukaran kepada purity 999.9
    
    Frm115.L9_Text = Format(Frm115_LM_BERAT_ASAL * Frm115_LM_KADAR_TUKARAN, "#,##0.00") 'Berat 999.9
    
Else

    Frm115.L9_Text = "0.00" 'Berat 999.9
    
End If
End Sub
Sub Frm115_calc5()
'On Error Resume Next
Dim Frm115_LM_BEZA_BERAT As Double
Dim Frm115_LM_HARGA_SEMASA As Double

Frm115_LM_BEZA_BERAT = 0 'Beza berat (g)
Frm115_LM_HARGA_SEMASA = 0 'Harga semasa (RM/g)

If ((Frm115.L9_Text <> vbNullString And IsNumeric(Frm115.L9_Text)) And (Frm115.TB2 <> vbNullString And IsNumeric(Frm115.TB2))) Then
    Frm115_LM_BEZA_BERAT = Frm115.L9_Text 'Berat jualan (g)
    Frm115_LM_HARGA_SEMASA = Frm115.TB2 'Kadar belian (g)
    
    Frm115.L12_Text = Format(Frm115_LM_BEZA_BERAT * Frm115_LM_HARGA_SEMASA, "#,##0.00") 'Harga jualan
Else
    Frm115.L12_Text = "0.00" 'Harga jualan
End If
End Sub
Sub Frm115_calc10()
'On Error Resume Next
Dim Frm115_LM_HARGA_ZR_UPAH As Double
Dim Frm115_LM_HARGA_SR_UPAH As Double
Dim Frm115_LM_HARGA_ZR_EMAS As Double
Dim Frm115_LM_HARGA_SR_EMAS As Double
Dim Frm115_LM_GST_ZR_UPAH As Double
Dim Frm115_LM_GST_SR_UPAH As Double
Dim Frm115_LM_GST_ZR_EMAS As Double
Dim Frm115_LM_GST_SR_EMAS As Double

Frm115_LM_HARGA_ZR_UPAH = 0
Frm115_LM_HARGA_SR_UPAH = 0
Frm115_LM_HARGA_ZR_EMAS = 0
Frm115_LM_HARGA_SR_EMAS = 0
Frm115_LM_GST_ZR_UPAH = 0
Frm115_LM_GST_SR_UPAH = 0
Frm115_LM_GST_ZR_EMAS = 0
Frm115_LM_GST_SR_EMAS = 0

If ((Frm115.L35_Text <> vbNullString And IsNumeric(Frm115.L35_Text)) And (Frm115.L39_Text <> vbNullString And IsNumeric(Frm115.L39_Text))) Then

    Frm115_LM_HARGA_ZR_UPAH = Frm115.L35_Text 'Harga ZR (Upah)
    'Frm115_LM_HARGA_ZR_EMAS = Frm115.L39_Text 'Harga ZR (Emas)
    
    Frm115.L17_Text = Format(Frm115_LM_HARGA_ZR_UPAH + Frm115_LM_HARGA_ZR_EMAS, "#,##0.00") 'Jumlah Harga ZR
    
Else

    Frm115.L17_Text = "0.00" 'Jumlah Harga ZR
    
End If

If ((Frm115.L37_Text <> vbNullString And IsNumeric(Frm115.L37_Text)) And (Frm115.L41_Text <> vbNullString And IsNumeric(Frm115.L41_Text))) Then

    Frm115_LM_HARGA_SR_UPAH = Frm115.L37_Text 'Harga SR (Upah)
    'Frm115_LM_HARGA_SR_EMAS = Frm115.L41_Text 'Harga SR (Emas)
    
    Frm115.L18_Text = Format(Frm115_LM_HARGA_SR_UPAH + Frm115_LM_HARGA_SR_EMAS, "#,##0.00") 'Jumlah Harga SR
    
Else

    Frm115.L18_Text = "0.00" 'Jumlah Harga SR
    
End If

If ((Frm115.L36_Text <> vbNullString And IsNumeric(Frm115.L36_Text)) And (Frm115.L40_Text <> vbNullString And IsNumeric(Frm115.L40_Text))) Then

    Frm115_LM_GST_SR_UPAH = Frm115.L36_Text 'GST ZR (Upah)
    'Frm115_LM_GST_SR_EMAS = Frm115.L40_Text 'GST ZR (Emas)
    
    Frm115.L20_Text = Format(Frm115_LM_GST_SR_UPAH + Frm115_LM_GST_SR_EMAS, "#,##0.00") 'Jumlah GST ZR
    
Else

    Frm115.L20_Text = "0.00" 'Jumlah GST ZR
    
End If

If ((Frm115.L38_Text <> vbNullString And IsNumeric(Frm115.L38_Text)) And (Frm115.L42_Text <> vbNullString And IsNumeric(Frm115.L42_Text))) Then

    Frm115_LM_GST_ZR_UPAH = Frm115.L38_Text 'GST SR (Upah)
    'Frm115_LM_GST_ZR_EMAS = Frm115.L42_Text 'GST SR (Emas)
    
    Frm115.L20_Text = Format(Frm115_LM_GST_ZR_UPAH + Frm115_LM_GST_ZR_EMAS, "#,##0.00") 'Jumlah GST SR
    
Else

    Frm115.L20_Text = "0.00" 'Jumlah GST SR
    
End If

Frm115.L15_Text = Format(Frm115_LM_HARGA_ZR_UPAH + Frm115_LM_HARGA_ZR_EMAS + Frm115_LM_HARGA_SR_UPAH + Frm115_LM_HARGA_SR_EMAS, "#,##0.00") 'Jumlah harga tanpa GST
Frm115.L16_Text = Format(Frm115_LM_HARGA_ZR_UPAH + Frm115_LM_HARGA_ZR_EMAS + Frm115_LM_HARGA_SR_UPAH + Frm115_LM_HARGA_SR_EMAS + Frm115_LM_GST_SR_UPAH + Frm115_LM_GST_SR_EMAS + Frm115_LM_GST_ZR_UPAH + Frm115_LM_GST_ZR_EMAS, "#,##0.00") 'Jumlah harga dengan GST
End Sub
Sub Frm115_Call_Product_Detail()
'on error resume next
Dim Frm115_LM_BERAT As Double

Frm115_LM_DATA_FOUND = 0
Frm115_LM_BERAT = 0
Frm115_LM_READY_TO_SAVE = 0 'Flag : Ready To Save
Frm115_LM_UpdateList = 0
Frm115_LM_KOD_PURITY = vbNullString
Frm115_LM_PERMATA = 0

frm115_LM_No_SIRI = UCase(Frm115.TB1) 'No. Siri Produk
Frm115.TB1 = vbNullString

Frm115_LM_KATEGORI = 1
'1:  Pelanggan Biasa
'2 : Member / Ahli
'3:  Pengedar
'4:  RAF
'5:  Normal Dealer(ND)
'6:  Master Dealer(MD)


' ### Periksa kategori pembeli ### - Start
'If Frm115.L46_Text <> vbNullString Then
'    If Frm28.L5_Text <> vbNullString Then
        
'        Set rs = New ADODB.Recordset
'        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'        rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
'        If Not rs.EOF Then
        
'            If Not IsNull(rs!kategori_pelanggan) Then Frm115_LM_KATEGORI = rs!kategori_pelanggan
            
'        End If
        
'        rs.Close
'        Set rs = Nothing
        
'    End If
'End If
' ### Periksa kategori pembeli ### - End

'###Periksa Samada Data Ini Telah Dimasukkan Ke Dalam Temp Table### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from " & G_GDN_TEMP & " where no_siri_produk='" & frm115_LM_No_SIRI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Frm115.L32_Text = "0" Then 'Data Baru (Kemasukkan Baru)
        If rs!Status = "1" Or rs!Status = "4" Then
        
            MsgBox "Item ini telah dimasukkan ke dalam senarai jualan sebelum ini.", vbInformation, "Info"
            
            Frm115.TB1 = vbNullString
            Frm115.TB1.SetFocus
            
        ElseIf rs!Status = 0 Then
            rs!Status = 1 '0 : Padam Data , 1 : Aktif (Tiada Perubahan) , 3 : Data Baru , 4 : Data Diedit , 5 : Keluarkan Data Dari Database Asal , 6 : Ignore Kemasukkan Data Ke Dalam Database
            rs.Update
            
            Frm115_LM_UpdateList = 1
        End If
    ElseIf Frm115.L32_Text = "1" Then 'Edit Data Lama + Kemasukkan Baru
        If rs!Status = "1" Or rs!Status = "4" Or rs!Status = "3" Then
        
            MsgBox "Item ini telah dimasukkan ke dalam senarai jualan sebelum ini.", vbInformation, "Info"
            
            Frm115.TB1 = vbNullString
            Frm115.TB1.SetFocus
            
        ElseIf rs!Status = "5" Or rs!Status = "6" Then
            If rs!Status = "5" Then rs!Status = "4" '0 : Padam Data , 1 : Aktif (Tiada Perubahan) , 3 : Data Baru , 4 : Data Diedit , 5 : Keluarkan Data Dari Database Asal , 6 : Ignore Kemasukkan Data Ke Dalam Database
            If rs!Status = "6" Then rs!Status = "3" '0 : Padam Data , 1 : Aktif (Tiada Perubahan) , 3 : Data Baru , 4 : Data Diedit , 5 : Keluarkan Data Dari Database Asal , 6 : Ignore Kemasukkan Data Ke Dalam Database
            rs.Update
            
            Frm115.TB1 = vbNullString
            Frm115.TB1.SetFocus
            
            Frm115_LM_UpdateList = 1
        End If
    End If
    Frm115_LM_DATA_FOUND = 1
    If rs!Status = "0" Or rs!Status = "5" Then Frm115_LM_DATA_FOUND = 0
End If

rs.Close
Set rs = Nothing
'###Periksa Samada Data Ini Telah Dimasukkan Ke Dalam Temp Table### - End

'###Carian Data Basic Bagi Item Ini### - Start
If Frm115_LM_DATA_FOUND = 0 Then

'###Periksa Mode Upah### - Start
    'Set rs = New ADODB.Recordset
    'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    'rs.Open "select * from default_setting where Default1='" & "Default" & "'", cn, adOpenKeyset, adLockOptimistic
    
    'If Not rs.EOF Then
    '    If Not IsNull(rs!flag_upah) Then
    '        If rs!flag_upah = 1 Then
    '            LM_UPAH_MODE = 1
    '        Else
    '            LM_UPAH_MODE = 0
    '        End If
    '    End If
    'End If
    
    'rs.Close
    'Set rs = Nothing
'###Periksa Mode Upah### - End

    If G_UPAH_MODE = 1 Then
        LM_UPAH_MODE = 1
    Else
        LM_UPAH_MODE = 0
    End If

    LM_FLAG_BARANG = 0 '0 : Barang yang belum pernah jual , 1 : Potong
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from Data_Database where no_siri_Produk='" & frm115_LM_No_SIRI & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        If rs!StatusItem = "10" Then
            LM_FLAG_BARANG = 0 '0 : Barang yang belum pernah jual , 1 : Potong
        ElseIf rs!StatusItem = "12" Or rs!StatusItem = "20" Or rs!StatusItem = "22" Or rs!StatusItem = "28" Then
            LM_FLAG_BARANG = 1 '0 : Barang yang belum pernah jual , 1 : Potong
        End If
        
        If rs!StatusItem = "10" Or rs!StatusItem = "12" Or rs!StatusItem = "20" Or rs!StatusItem = "22" Or rs!StatusItem = "28" Then
        
            If Not IsNull(rs!cawangan) Then
                
                If MDI_frm1.L20_Text <> rs!cawangan Then
                    
                    MsgBox "Stok ini adalah milik cawangan [" & rs!cawangan & "]. Anda tidak dibenarkan untuk jual barang ini.", vbExclamation, "Info"
                    
                    rs.Close
                    Set rs = Nothing
                    
                    Exit Sub
                    
                End If
            
            End If
        
            If Not IsNull(rs!receiving_Status) Then
                If rs!receiving_Status = 0 Or rs!receiving_Status = 2 Or rs!receiving_Status = 4 Or rs!receiving_Status = 5 Then
                    
                    Frm115.L3_Text = frm115_LM_No_SIRI 'No. Siri Produk
                    Frm115.L6_Text = Format(rs!beza_berat, "0.00") 'Berat Asal (g)
                    Frm115.TB3 = Format(rs!beza_berat, "0.00") 'Berat Jualan (g)
                    Frm115.L33_Text = Format(rs!harga_Per_Gram_Item, "0.00") 'Harga Per Gram Item (RM/g)
                    Frm115.L50_Text = rs!UPAH 'Upah modal
                    Frm115.TB4.Locked = False 'Upah
                    Frm115.TB4.BackColor = &HFFFFFF 'Upah
                    
                    If LM_FLAG_BARANG = 0 Then '0 : Barang yang belum pernah jual , 1 : Potong
                        Frm115.TB3.Locked = False
                        Frm115.TB3.BackColor = &HFFFFFF
                    Else
                        Frm115.TB3.Locked = True
                        Frm115.TB3.BackColor = &H8000000A
                    End If
                    
                    Frm115_LM_READY_TO_SAVE = 1 'Flag : Ready To Save
                Else
                    MsgBox "Barang yang ingin dijual [" & frm115_LM_No_SIRI & "] adalah barang permata." & vbCrLf & _
                            vbNullString & vbCrLf & _
                            "Hanya barang kemas (yang mempunyai berat) SAHAJA dibenarkan dijual dalam menu ini.", vbInformation, "Info"
                            
                    Frm115.TB1 = vbNullString
                    Frm115.TB1.SetFocus
                            
                    Exit Sub
                End If
            End If
            
            If LM_UPAH_MODE = 1 And Frm115.CB5 = 1 Then
                If Frm115_LM_KATEGORI = 1 Then
                    If Not IsNull(rs!Upah_Jualan) Then
                        Frm115.TB4 = Format(rs!Upah_Jualan, "0.00") 'Upah Pelanggan
                    End If
                ElseIf Frm115_LM_KATEGORI = 2 Then
                    If Not IsNull(rs!Upah_Member) Then
                        Frm115.TB4 = Format(rs!Upah_Member, "0.00") 'Upah Member
                    End If
                ElseIf Frm115_LM_KATEGORI = 3 Then
                    If Not IsNull(rs!Upah_RAF) Then
                        Frm115.TB4 = Format(rs!Upah_RAF, "0.00") 'Upah RAF
                    End If
                ElseIf Frm115_LM_KATEGORI = 4 Then
                    If Not IsNull(rs!Upah_Pengedar) Then
                        Frm115.TB4 = Format(rs!Upah_Pengedar, "0.00") 'Upah Pengedar
                    End If
                ElseIf Frm115_LM_KATEGORI = 5 Then
                    If Not IsNull(rs!upah_normal_dealer) Then
                        Frm115.TB4 = Format(rs!upah_normal_dealer, "0.00") 'Upah Normal Dealer
                    End If
                ElseIf Frm115_LM_KATEGORI = 6 Then
                    If Not IsNull(rs!upah_master_dealer) Then
                        Frm115.TB4 = Format(rs!upah_master_dealer, "0.00") 'Upah Master Dealer
                    End If
                End If
            Else
                Frm115.TB4 = Format(0, "0.00") 'Upah
            End If
            
            If Not IsNull(rs!kategori_Produk) Then Frm115.L5_Text = rs!kategori_Produk 'Kategori Produk
            If Not IsNull(rs!kod_Purity) Then
                Frm115_LM_KOD_PURITY = rs!kod_Purity 'Kod Purity
                Frm115.L4_Text = rs!kod_Purity 'Kod Purity
            End If
        ElseIf rs!StatusItem = "11" Then
            MsgBox "Item ini telah dijual. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm115.TB1 = vbNullString
            Frm115.TB1.SetFocus
        'ElseIf rs!StatusItem = "12" Then
        '    MsgBox "Item ini telah dijual secara potong. No. Siri Produk [" & Frm115_LM_No_SIRI & "]", vbExclamation, "Info"
            
        '    Frm115.TB1 = vbNullString
        '    Frm115.TB1.SetFocus
        ElseIf rs!StatusItem = "13" Then
            MsgBox "Item ini telah dijual secara potong. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm115.TB1 = vbNullString
            Frm115.TB1.SetFocus
        ElseIf rs!StatusItem = "14" Or rs!StatusItem = "21" Then
            MsgBox "Item ini telah ditempah oleh pelanggan. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm115.TB1 = vbNullString
            Frm115.TB1.SetFocus
        ElseIf rs!StatusItem = "15" Or rs!StatusItem = "19" Or rs!StatusItem = "20" Then
            MsgBox "Item ini telah dibeli secara ansuran. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm115.TB1 = vbNullString
            Frm115.TB1.SetFocus
        ElseIf rs!StatusItem = "16" Then
            MsgBox "Item ini telah dihantar ke Ar-Rahnu. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm115.TB1 = vbNullString
            Frm115.TB1.SetFocus
        ElseIf rs!StatusItem = "17" Then
            MsgBox "Item ini telah dijual secara ETA. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm115.TB1 = vbNullString
            Frm115.TB1.SetFocus
        ElseIf rs!StatusItem = "23" Then
            MsgBox "Item Ini Telah Dihantar Ke Supplier/Kilang. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm115.TB1 = vbNullString
            Frm115.TB1.SetFocus
        ElseIf rs!StatusItem = "24" Then
            MsgBox "Item Ini Telah Dihantar Ke Supplier/Kilang. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm115.TB1 = vbNullString
            Frm115.TB1.SetFocus
        ElseIf rs!StatusItem = "25" Then
            MsgBox "Item Ini Telah Diagihkan Ke Cawangan. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
    
            Frm115.TB1 = vbNullString
            Frm115.TB1.SetFocus
        ElseIf rs!StatusItem = "26" Then
            MsgBox "Item Ini Telah Dijual Oleh Cawangan. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
    
            Frm115.TB1 = vbNullString
            Frm115.TB1.SetFocus
        ElseIf rs!StatusItem = "0" Then
            MsgBox "Item Ini Telah Dipadamkan Dari Sistem. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm115.TB1 = vbNullString
            Frm115.TB1.SetFocus
        ElseIf rs!StatusItem = "27" Then
            MsgBox "Item Ini Telah Dijual Dari Menu GDN. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
    
            Frm115.TB1 = vbNullString
            Frm115.TB1.SetFocus
        ElseIf rs!StatusItem = "29" Then
            MsgBox "Item Ini Telah Diubah Status Kepada Hilang , Dicuri Dan Sebagainya. No. Siri Produk [" & frm115_LM_No_SIRI & "]", vbExclamation, "Info"
    
            Frm115.TB1 = vbNullString
            Frm115.TB1.SetFocus
        End If
    Else
        MsgBox "No. Siri Produk Ini [" & frm115_LM_No_SIRI & "] Tidak Dijumpai.", vbExclamation, "Info"
        
        Frm115.TB1 = vbNullString
        Frm115.TB1.SetFocus
    End If
    
    rs.Close
    Set rs = Nothing
    
    If Frm115_LM_KOD_PURITY <> vbNullString Then

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from hargaemas where Purity='" & Frm115_LM_KOD_PURITY & "' AND cawangan='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic

        If Not rs.EOF Then
            If Not IsNull(rs!HargaDariSupplier) Then
                If IsNumeric(rs!HargaDariSupplier) Then
                    Frm115.L49_Text = rs!HargaDariSupplier
                Else
                    Frm115.L49_Text = 0
                End If
            Else
                Frm115.L49_Text = 0
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        
    End If
End If
'###Carian Data Basic Bagi Item Ini### - End

'###Periksa Data Produk### - Start

'Frm115.TB1 = vbNullString
'If Frm115.CB1 = 1 Then Call Frm115_auto_insert_data

'If Frm115_LM_UpdateList = 1 Then
    'Call Frm115_Senarai_Jualan_Header
    'Call Frm115_Senarai_Jualan
    Frm115.TB1.SetFocus
'End If
'###Periksa Data Produk### - End
End Sub
Sub Frm115_Senarai_Jualan_Header()
'on error resume next

With Frm115.LV2
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear

    Frm115.LV2.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "No. Siri Produk", 1800
    .ColumnHeaders.Add 5, , "Kategori Produk", 3000
    .ColumnHeaders.Add 6, , "Berat Asal (g)", 1400, 1
    .ColumnHeaders.Add 7, , "Berat Jualan (g)", 1700, 1
    .ColumnHeaders.Add 8, , "Kadar Tukaran 999.9", 2000, 1
    .ColumnHeaders.Add 9, , "Berat 999.9 (g)", 1600, 1
    .ColumnHeaders.Add 10, , "Upah (RM)", 1400, 1
    .ColumnHeaders.Add 11, , "Jenis GST", 1400, 2
    .ColumnHeaders.Add 12, , "Jumlah GST (RM)", 1700, 1
    .ColumnHeaders.Add 13, , "Upah + GST (RM)", 1700, 1

    
End With
End Sub
Sub Frm115_Senarai_Jualan()
'on error resume next
Dim frm115_LM_TOTAL_PAGE As Double
Dim frm115_LM_FIELD As String
Dim Frm115_LM_UPAH_TANPA_GST As Double 'Harga Jualan Tanpa Cukai GST
Dim Frm115_LM_UPAH_DENGAN_GST As Double 'Harga Jualan Dengan Cukai GST
Dim Frm115_LM_GST_SR As Double 'Kutipan GST : SR
Dim Frm115_LM_GST_ZR As Double 'Kutipan GST : ZR
Dim Frm115_LM_JUMLAH_UPAH_SR As Double 'Total Harga Yang Dikenakan GST SR
Dim Frm115_LM_JUMLAH_UPAH_ZR As Double 'Total Harga Yang Dikenakan GST ZR
Dim Frm115_LM_BERAT As Double 'Berat Jualan
Dim Frm115_LM_BERAT_ASAL As Double 'Berat Asal (Sebelum tukar kepada purity 999.9)

frm115_PAGE_SIZE = 26
frm115_LM_TOTAL_PAGE = 0
x = 0
Frm115_LM_UPAH_TANPA_GST = 0
Frm115_LM_UPAH_DENGAN_GST = 0
Frm115_LM_GST_SR = 0
Frm115_LM_GST_ZR = 0
Frm115_LM_JUMLAH_UPAH_SR = 0
Frm115_LM_JUMLAH_UPAH_ZR = 0
Frm115_LM_BERAT = 0
Frm115_LM_BERAT_ASAL = 0 'Berat Asal (Sebelum tukar kepada purity 999.9)

re_gen_report:

Frm115.L43_Text = x 'Jumlah bilangan barang jualan
Frm115.L48_Text = Format(0, "#,##0.00") 'Jumlah berat jualan
Frm115.L35_Text = Format(0, "#,##0.00") 'Maklumat GST : Jumlah harga ZR
Frm115.L37_Text = Format(0, "#,##0.00") 'Maklumat GST : Jumlah harga SR
Frm115.L36_Text = Format(0, "#,##0.00") 'Maklumat GST : Jumlah GST ZR
Frm115.L38_Text = Format(0, "#,##0.00")  'Maklumat GST : Jumlah GST SR
'Frm115.L9_Text = Format(0, "#,##0.00") 'Berat jualan 999.9
Frm115.L51_Text = Format(0, "#,##0.00") 'Jumlah upah tanpa GST (keseluruhan)
Frm115.L52_Text = Format(0, "#,##0.00") 'Jumlah GST (Keseluruhan)
Frm115.L53_Text = Format(0, "#,##0.00") 'Jumlah Upah + GST (Keseluruhan)

LM_START_ROW = Frm115.L69_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + frm115_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm115.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - frm115_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        Frm115.L67_Text = 1
    End If
End If

frm115_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from " & G_GDN_TEMP & " where Status='" & 1 & "' Or Status='" & 2 & "' Or Status='" & 3 & "' Or Status='" & 4 & "' order by kategori_Produk ASC LIMIT " & LM_START_ROW & "," & frm115_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1
    If frm115_LM_PAGE_FOUND = 0 Then
        If Frm115.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm115.L67_Text = Frm115.L67_Text + 1 'Paparan Page ke-xxx
                frm115_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm115.L67_Text) Then
                    If Frm115.L67_Text <> 1 Then
                        Frm115.L67_Text = Frm115.L67_Text - 1 'Paparan Page ke-xxx
                        frm115_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    
    Y = ((Frm115.L67_Text - 1) * frm115_PAGE_SIZE) + x

    With Frm115.LV2.ListItems.Add(, , rs!ID)
     
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID

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
        
        If Not IsNull(rs!Berat_Asal) Then 'Berat Asal (g)
            .ListSubItems.Add , , Format(rs!Berat_Asal, "#,##0.00")
            If IsNumeric(rs!Berat_Asal) Then Frm115_LM_BERAT_ASAL = Frm115_LM_BERAT_ASAL + rs!Berat_Asal
        Else
            .ListSubItems.Add , , "0.00"
        End If
        
        If Not IsNull(rs!berat_jualan) Then 'Berat Jualan (g)
            .ListSubItems.Add , , Format(rs!berat_jualan, "#,##0.00")
        Else
            .ListSubItems.Add , , "0.00"
        End If
        
        If Not IsNull(rs!pemalar_tukaran_999) Then 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
            .ListSubItems.Add , , Format(rs!pemalar_tukaran_999, "#,##0.00")
        Else
            .ListSubItems.Add , , "0.00"
        End If
        
        If Not IsNull(rs!berat_999) Then 'Berat barang kemas selepas ditukar kepada purity 999.9
            .ListSubItems.Add , , Format(rs!berat_999, "#,##0.00")
            If IsNumeric(rs!berat_999) Then Frm115_LM_BERAT = Frm115_LM_BERAT + rs!berat_999
        Else
            .ListSubItems.Add , , "0.00"
        End If
        
        If Not IsNull(rs!UPAH) Then 'Upah (RM)
            .ListSubItems.Add , , Format(rs!UPAH, "#,##0.00")
            If IsNumeric(rs!harga_tanpa_gst) Then Frm115_LM_UPAH_TANPA_GST = Frm115_LM_UPAH_TANPA_GST + rs!harga_tanpa_gst
        Else
            .ListSubItems.Add , , "0.00"
        End If
        
        If Not IsNull(rs!gst_ari_nashi) Then 'Jenis GST

            If rs!gst_ari_nashi = "ZR (L)" Then
                .ListSubItems.Add , , "ZR(L)"  'Jenis GST : Zero Rated
                If IsNumeric(rs!jumlah_gst) Then Frm115_LM_GST_ZR = Frm115_LM_GST_ZR + rs!jumlah_gst 'Jumlah Kutipan GST ZR(L)
                If IsNumeric(rs!harga_tanpa_gst) Then Frm115_LM_JUMLAH_UPAH_ZR = Frm115_LM_JUMLAH_UPAH_ZR + rs!harga_tanpa_gst 'Total Harga Yang Dikenakan GST ZR
            ElseIf rs!gst_ari_nashi = "SR" Then
                .ListSubItems.Add , , "SR"  'Jenis GST : Standard Rated
                If IsNumeric(rs!jumlah_gst) Then Frm115_LM_GST_SR = Frm115_LM_GST_SR + rs!jumlah_gst 'Jumlah Kutipan GST SR
                If IsNumeric(rs!harga_tanpa_gst) Then Frm115_LM_JUMLAH_UPAH_SR = Frm115_LM_JUMLAH_UPAH_SR + rs!harga_tanpa_gst 'Total Harga Yang Dikenakan GST SR
            End If
        
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!jumlah_gst) Then 'Jumlah GST (RM)
            .ListSubItems.Add , , Format(rs!jumlah_gst, "#,##0.00")
        Else
            .ListSubItems.Add , , "0.00"
        End If
        
        If Not IsNull(rs!harga_dengan_gst) Then 'Harga Dengan GST (RM)
            .ListSubItems.Add , , Format(rs!harga_dengan_gst, "#,##0.00")
            If IsNumeric(rs!harga_dengan_gst) Then Frm115_LM_UPAH_DENGAN_GST = Frm115_LM_UPAH_DENGAN_GST + rs!harga_dengan_gst 'Harga Jualan Dengan GST (RM)
        Else
            .ListSubItems.Add , , "0.00"
        End If
        
    End With

    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from " & G_GDN_TEMP & " where Status='" & 1 & "' Or Status='" & 2 & "' Or Status='" & 3 & "' Or Status='" & 4 & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    LM_BILANGAN_AHLI = rs(0)
    frm115_LM_TOTAL_PAGE = Format(rs(0) / frm115_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, frm115_LM_TOTAL_PAGE, ".") <> 0 Then
    
        frm115_LM_PAGE = Split(frm115_LM_TOTAL_PAGE, ".")(0)
        frm115_LM_PAGE_LEBIHAN = Split(frm115_LM_TOTAL_PAGE, ".")(1)
        
        If frm115_LM_PAGE_LEBIHAN <> "00" Then
            Frm115.L68_Text = frm115_LM_PAGE + 1
        Else
            Frm115.L68_Text = frm115_LM_PAGE
        End If
        
    Else
    
        Frm115.L68_Text = frm115_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm115.L68_Text = 0
    End If
Else
    Frm115.L68_Text = 0
End If

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) , SUM(berat_999) , SUM(berat_999) , SUM(harga_tanpa_gst) , SUM(jumlah_gst) , SUM(harga_dengan_gst) from " & G_GDN_TEMP & " where Status='" & 1 & "' Or Status='" & 2 & "' Or Status='" & 3 & "' Or Status='" & 4 & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm115.L43_Text = rs(0) 'Jumlah bilangan barang jualan
If Not IsNull(rs(1)) Then Frm115.L48_Text = Format(rs(1), "#,##0.00") 'Jumlah berat jualan
'If Not IsNull(rs(2)) Then Frm115.L9_Text = Format(rs(2), "#,##0.00") 'Berat jualan 999.9
If Not IsNull(rs(3)) Then Frm115.L51_Text = Format(rs(3), "#,##0.00") 'Jumlah upah tanpa GST (keseluruhan)
If Not IsNull(rs(4)) Then Frm115.L52_Text = Format(rs(4), "#,##0.00") 'Jumlah GST (Keseluruhan)
If Not IsNull(rs(5)) Then Frm115.L53_Text = Format(rs(5), "#,##0.00") 'Jumlah Upah + GST (Keseluruhan)

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_gst) , SUM(harga_tanpa_gst) from " & G_GDN_TEMP & " where (Status='" & 1 & "' Or Status='" & 2 & "' Or Status='" & 3 & "' Or Status='" & 4 & "') AND gst_ari_nashi='" & "ZR (L)" & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm115.L36_Text = Format(rs(0), "#,##0.00") 'Maklumat GST : Jumlah GST ZR
If Not IsNull(rs(1)) Then Frm115.L35_Text = Format(rs(1), "#,##0.00") 'Maklumat GST : Jumlah harga ZR

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_gst) , SUM(harga_tanpa_gst) from " & G_GDN_TEMP & " where (Status='" & 1 & "' Or Status='" & 2 & "' Or Status='" & 3 & "' Or Status='" & 4 & "') AND gst_ari_nashi='" & "SR" & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm115.L38_Text = Format(rs(0), "#,##0.00") 'Maklumat GST : Jumlah GST SR
If Not IsNull(rs(1)) Then Frm115.L37_Text = Format(rs(1), "#,##0.00") 'Maklumat GST : Jumlah harga SR

rs.Close
Set rs = Nothing

If x <> 0 Then
    Frm115.L69_Text = LM_START_ROW
End If

If Frm115.L67_Text <> vbNullString And IsNumeric(Frm115.L67_Text) Then
    If Frm115.L68_Text <> vbNullString And IsNumeric(Frm115.L68_Text) Then
        frm115_LM_CURR_PAGE = Frm115.L67_Text
        frm115_LM_TOTAL_PAGE = Frm115.L68_Text
        
        If frm115_LM_CURR_PAGE > frm115_LM_TOTAL_PAGE Then
            
            Frm115.L67_Text = Frm115.L67_Text - 1
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            GoTo re_gen_report:
            
        End If
    End If
End If

End Sub
Sub Frm115_recall_edit_jualan()
'on error resume next
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset

Frm115_LM_EMP_NAMA = vbNullString
Frm115_LM_SUPPLIER = vbNullString

GLOBAL_DISABLE = 1

'### Maklumat asas bagi invoice ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where no_resit='" & Frm115.L23_Text & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Not IsNull(rs!tarikh) Then Frm115.DTPicker1 = rs!tarikh 'Tarikh Jualan
    If Not IsNull(rs!no_rujukan_pembeli) Then Frm115_LM_No_PEMBELI = rs!no_rujukan_pembeli 'No. Rujukan Pembeli
    If Not IsNull(rs!no_pekerja) Then 'No. Pekerja
        Frm115_LM_No_PEKERJA = rs!no_pekerja
    End If
    
End If

rs.Close
Set rs = Nothing

'-----------------
'Set rs = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

'strsql = "UPDATE 23_senarai_jualan,data_database SET data_database.gdn_temp = 1 FROM 23_senarai_jualan WHERE 23_senarai_jualan.no_siri_produk = data_database.no_siri_produk AND no_resit='" & Frm115.L23_Text & "' AND status_rekod='" & 1 & "'"
'GDN2016-000003
'Set rs = cn.Execute(strsql)
'Set rs = Nothing
'-----------------

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "insert into " & G_GDN_TEMP & "(id_database,no_siri_produk,kategori_produk,purity,berat_asal,berat_jualan,harga_semasa,upah,harga_asal,diskaun,harga_lepas_diskaun,adjustment,harga_jualan,gst_ari_nashi,kadar_gst,jumlah_gst,harga_dengan_gst,dropship,komisyen_per_gram,jumlah_komisyen,type,potong_flag,harga_per_gram_modal,modal,untung,harga_per_gram_supplier,untung2,dulang,gst_include,harga_tanpa_gst,harga_koperasi,kadar_penurunan_upah,kadar_penurunan_bp,harga_semasa_staff,harga_bp_asal,upah_asal,harga_staff,komisyen_staff,pemalar_tukaran_999,berat_999,upah_modal,gst_barang_atau_upah,harga_jualan_dengan_gst,kadar_komisyen_upah,komisyen_upah,jualan_per_gram,modal_per_gram,flag_upah,upah_per_gram,status)" & _
            "select ID,no_siri_produk,kategori_produk," _
            & "purity,berat_asal,berat_jualan,harga_semasa,upah,harga_asal,diskaun,harga_lepas_diskaun,adjustment,harga_jualan," _
            & "gst_ari_nashi,kadar_gst,jumlah_gst,harga_dengan_gst,dropship,komisyen_per_gram,jumlah_komisyen,type," _
            & "potong_flag,harga_per_gram_modal,modal,untung,harga_per_gram_supplier,untung2,dulang,gst_include,harga_tanpa_gst," _
            & "harga_koperasi,kadar_penurunan_upah,kadar_penurunan_bp,harga_semasa_staff,harga_bp_asal,upah_asal,harga_staff," _
            & "komisyen_staff,pemalar_tukaran_999,berat_999,upah_modal,gst_barang_atau_upah,harga_jualan_dengan_gst," _
            & "kadar_komisyen_upah,komisyen_upah,jualan_per_gram,modal_per_gram,flag_upah,upah_per_gram,2 " _
            & "from 23_senarai_jualan WHERE no_resit='" & Frm115.L23_Text & "' AND status_rekod='" & 1 & "'"
    
Set rs = cn.Execute(strsql)
Set rs = Nothing

GoTo skip_b:
    
'### Masukkan Data Jualan Ke Dalam Table Jualan (Temp) ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 23_senarai_jualan where no_resit='" & Frm115.L23_Text & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set rs1 = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs1.Open "select * from " & G_GDN_TEMP & "", cn, adOpenKeyset, adLockOptimistic

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

    'If Frm115.CB4 = 1 Then
    '    rs1!kategori_pembeli = 1
    'ElseIf Frm115.CB5 = 1 Then
    '    rs1!kategori_pembeli = 2
    'ElseIf Frm115.CB6 = 1 Then
    '    rs1!kategori_pembeli = 4
    'ElseIf Frm115.CB9 = 1 Then
    '    rs1!kategori_pembeli = 3
    'ElseIf Frm115.CB10 = 1 Then
    '    rs1!kategori_pembeli = 5
    'ElseIf Frm115.CB11 = 1 Then
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

skip_b:

GM_NEXT_PREV = 0

Frm115.L69_Text = -1 'Titik Pencarian Data
Frm115.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Frm115.L67_Text = 0 'Paparan Page ke-xxx
Frm115.L68_Text = 0

Call Frm115_Senarai_Jualan_Header
Call Frm115_Senarai_Jualan

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 77_gdn_grn where no_rujukan='" & Frm115.L23_Text & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    
    If Not IsNull(rs!tarikh) Then Frm115.DTPicker1 = rs!tarikh 'Tarikh
    If Not IsNull(rs!Berat_Asal) Then Frm115.L48_Text = Format(rs!Berat_Asal, "#,##0.00")
    If Not IsNull(rs!kadar_tukaran) Then Frm115.TB8 = rs!kadar_tukaran
    If Not IsNull(rs!harga_999) Then Frm115.TB2 = Format(rs!harga_999, "#,##0.00")
    If Not IsNull(rs!berat_tukaran) Then Frm115.L9_Text = Format(rs!berat_tukaran, "#,##0.00")
    If Not IsNull(rs!harga_tanpa_gst) Then Frm115.L51_Text = Format(rs!harga_tanpa_gst, "#,##0.00")
    If Not IsNull(rs!kadar_gst) Then Frm115.L21_Text = rs!kadar_gst
    If Not IsNull(rs!jumlah_gst) Then Frm115.L52_Text = Format(rs!jumlah_gst, "#,##0.00")
    If Not IsNull(rs!harga_dengan_gst) Then Frm115.L53_Text = Format(rs!harga_dengan_gst, "#,##0.00")
    If Not IsNull(rs!nilaian_harga_emas) Then Frm115.L12_Text = Format(rs!nilaian_harga_emas, "#,##0.00")
    If Not IsNull(rs!gst_zr_harga) Then Frm115.L17_Text = Format(rs!gst_zr_harga, "#,##0.00")
    If Not IsNull(rs!gst_sr_harga) Then Frm115.L18_Text = Format(rs!gst_sr_harga, "#,##0.00")
    If Not IsNull(rs!gst_zr_cukai) Then Frm115.L19_Text = Format(rs!gst_zr_cukai, "#,##0.00")
    If Not IsNull(rs!gst_sr_cukai) Then Frm115.L20_Text = Format(rs!gst_sr_cukai, "#,##0.00")
    If Not IsNull(rs!user) Then Frm115_LM_EMP_NAMA = rs!user
    If Not IsNull(rs!supplier_agen) Then Frm115_LM_SUPPLIER = rs!supplier_agen
    
End If

rs.Close
Set rs = Nothing

'### Carian Maklumat Penjual (Data Pekerja) ### - Start
If Frm115_LM_EMP_NAMA <> vbNullString Then

    DATA_PEKERJA_FOUND = 0
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where Samaran='" & Frm115_LM_EMP_NAMA & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        Frm115_LM_MAKLUMAT_PEKERJA = rs!Samaran & "  |  " & rs!NoPekerja
        DATA_PEKERJA_FOUND = 1
        
    End If
    
    rs.Close
    Set rs = Nothing

    If DATA_PEKERJA_FOUND = 1 Then
        'On Error GoTo Err_A:
        Frm115.CBB4 = Frm115_LM_MAKLUMAT_PEKERJA
        
Restore_A:
    End If
    
    'on error resume next
End If
'### Carian Maklumat Penjual (Data Pekerja) ### - End

'### Maklumat Supplier ### - Start
If Frm115_LM_SUPPLIER <> vbNullString Then
     
     'On Error GoTo Err_B:
     Frm115.CBB2 = Frm115_LM_SUPPLIER

Restore_B:
        
End If
'### Maklumat Supplier ### - End

GLOBAL_DISABLE = 0

Frm115.CBB4.Enabled = True
Frm115.CBB4.BackColor = &HFFFFFF

Frm115.Show
Frm85.Hide

Exit Sub
Err_A:
Frm115.CBB4.AddItem Frm115_LM_MAKLUMAT_PEKERJA
Frm115.CBB4 = Frm115_LM_MAKLUMAT_PEKERJA
Resume Restore_A:

Exit Sub
Err_B:
Frm115.CBB2.AddItem Frm115_LM_SUPPLIER
Frm115.CBB2 = Frm115_LM_SUPPLIER
Resume Restore_B:
End Sub
Sub Frm115_cetak_gdn()
'on error resume next
Frm115_LM_CUST = vbNullString

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
'        PRINTER_FOUND = 1 '0 : Printer Not Found , 1 : Printer Found
'        Exit For
'    End If
'Next

Report56.Sections("Section4").Controls("L1").Caption = vbNullString 'No. Rujukan
Report56.Sections("Section4").Controls("L2").Caption = vbNullString 'Tarikh
Report56.Sections("Section4").Controls("L3").Caption = vbNullString 'Nama Pembeli
Report56.Sections("Section4").Controls("L4").Caption = vbNullString 'No. Telefon
Report56.Sections("Section4").Controls("L17").Caption = vbNullString 'Jurujual
Report56.Sections("Section4").Controls("L18").Caption = "-" 'No. ID GST
Report56.Sections("Section5").Controls("L15").Caption = "0" 'Bilangan barang
Report56.Sections("Section5").Controls("L16").Caption = "0.00" 'Berat Asal (g)
Report56.Sections("Section5").Controls("L19").Caption = "1.00" 'Mutu
Report56.Sections("Section5").Controls("L20").Caption = "0.00" 'Berat 999.9 (g)
Report56.Sections("Section5").Controls("L13").Caption = "0.00" 'Jumlah GST
Report56.Sections("Section5").Controls("L14").Caption = "0.00" 'Jumlah keseluruhan (Upah + GST)
Report56.Sections("Section5").Controls("L8").Caption = "0.00" 'Jumlah harga SR
Report56.Sections("Section5").Controls("L9").Caption = "0.00" 'Jumlah harga ZR
Report56.Sections("Section5").Controls("L10").Caption = "0.00" 'Jumlah GST SR
Report56.Sections("Section5").Controls("L11").Caption = "0.00" 'Jumlah GST ZR

'### Reset maklumat kedai ### - Start
Report56.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report56.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report56.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report56.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report56.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

Report56.Sections("Section4").Controls("L205").Caption = "Goods Despatch Note"

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

Report56.Sections("Section4").Controls("L1").Caption = G_No_RESIT_JUALAN 'No. Invoice

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 77_gdn_grn where no_rujukan='" & G_No_RESIT_JUALAN & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Not IsNull(rs!tarikh) Then Report56.Sections("Section4").Controls("L2").Caption = rs!tarikh 'Tarikh
    If Not IsNull(rs!user) Then Report56.Sections("Section4").Controls("L17").Caption = rs!user 'Jurujual
    If Not IsNull(rs!bil_barang) Then Report56.Sections("Section5").Controls("L15").Caption = rs!bil_barang 'Bilangan barang
    If Not IsNull(rs!Berat_Asal) Then Report56.Sections("Section5").Controls("L16").Caption = Format(rs!Berat_Asal, "#,##0.00 g") 'Berat Asal (g)
    If Not IsNull(rs!kadar_tukaran) Then Report56.Sections("Section5").Controls("L19").Caption = rs!kadar_tukaran 'Mutu
    If Not IsNull(rs!berat_tukaran) Then Report56.Sections("Section5").Controls("L20").Caption = Format(rs!berat_tukaran, "#,##0.00 g") 'Berat 999.9 (g)
    If Not IsNull(rs!jumlah_gst) Then Report56.Sections("Section5").Controls("L13").Caption = Format(rs!jumlah_gst, "#,##0.00") 'Jumlah GST
    If Not IsNull(rs!harga_dengan_gst) Then Report56.Sections("Section5").Controls("L14").Caption = Format(rs!harga_dengan_gst, "#,##0.00") 'Jumlah keseluruhan (Upah + GST)
    If Not IsNull(rs!gst_sr_harga) Then Report56.Sections("Section5").Controls("L8").Caption = Format(rs!gst_sr_harga, "#,##0.00") 'Jumlah harga SR
    If Not IsNull(rs!gst_zr_harga) Then Report56.Sections("Section5").Controls("L9").Caption = Format(rs!gst_zr_harga, "#,##0.00") 'Jumlah harga ZR
    If Not IsNull(rs!gst_sr_cukai) Then Report56.Sections("Section5").Controls("L10").Caption = Format(rs!gst_sr_cukai, "#,##0.00") 'Jumlah GST SR
    If Not IsNull(rs!gst_zr_cukai) Then Report56.Sections("Section5").Controls("L11").Caption = Format(rs!gst_zr_cukai, "#,##0.00") 'Jumlah GST ZR
    If Not IsNull(rs!supplier_agen) Then Frm115_LM_CUST = rs!supplier_agen

End If

rs.Close
Set rs = Nothing

If Frm115_LM_CUST <> vbNullString Then
 
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from setting_database where Supplier='" & Frm115_LM_CUST & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        
        If Not IsNull(rs!supplier) Then Report56.Sections("Section4").Controls("L3").Caption = rs!supplier 'Nama Pembeli
        If Not IsNull(rs!no_tel_hp) Then Report56.Sections("Section4").Controls("L4").Caption = rs!no_tel_hp 'No. Telefon
        If Not IsNull(rs!no_id_gst) Then Report56.Sections("Section4").Controls("L18").Caption = rs!no_id_gst 'No. ID GST

    End If
    
    rs.Close
    Set rs = Nothing
    
End If
   
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 23_senarai_jualan where no_resit='" & G_No_RESIT_JUALAN & "' AND status_rekod = 1", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report56.DataSource = rs
    Report56.Show
    
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing
End Sub
Sub frm123_cetak_gdn()
'on error resume next
frm123_LM_CUST = vbNullString

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

Report78.Sections("Section4").Controls("L1").Caption = vbNullString 'No. Rujukan
Report78.Sections("Section4").Controls("L2").Caption = vbNullString 'Tarikh
Report78.Sections("Section4").Controls("L3").Caption = vbNullString 'Nama Pembeli
Report78.Sections("Section4").Controls("L4").Caption = vbNullString 'No. Telefon
Report78.Sections("Section4").Controls("L17").Caption = vbNullString 'Jurujual
Report78.Sections("Section4").Controls("L18").Caption = "-" 'No. ID GST
Report78.Sections("Section4").Controls("L21").Caption = vbNullString 'No. Rujukan Dari Supplier
Report78.Sections("Section5").Controls("L15").Caption = "0" 'Bilangan barang
Report78.Sections("Section5").Controls("L16").Caption = "0.00" 'Berat Asal (g)
Report78.Sections("Section5").Controls("L19").Caption = "1.00" 'Mutu
Report78.Sections("Section5").Controls("L20").Caption = "0.00" 'Berat 999.9 (g)
Report78.Sections("Section5").Controls("L13").Caption = "0.00" 'Jumlah GST
Report78.Sections("Section5").Controls("L14").Caption = "0.00" 'Jumlah keseluruhan (Upah + GST)
Report78.Sections("Section5").Controls("L8").Caption = "0.00" 'Jumlah harga SR
Report78.Sections("Section5").Controls("L9").Caption = "0.00" 'Jumlah harga ZR
Report78.Sections("Section5").Controls("L10").Caption = "0.00" 'Jumlah GST SR
Report78.Sections("Section5").Controls("L11").Caption = "0.00" 'Jumlah GST ZR

'### Reset maklumat kedai ### - Start
Report78.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report78.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report78.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report78.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report78.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

Report78.Sections("Section4").Controls("L205").Caption = "Goods Despatch Note"

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report78.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report78.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report78.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report78.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report78.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Report78.Sections("Section4").Controls("L1").Caption = G_No_RESIT_JUALAN 'No. Invoice

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 77_gdn_grn where no_rujukan='" & G_No_RESIT_JUALAN & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Not IsNull(rs!tarikh) Then Report78.Sections("Section4").Controls("L2").Caption = rs!tarikh 'Tarikh
    If Not IsNull(rs!user) Then Report78.Sections("Section4").Controls("L17").Caption = rs!user 'Jurujual
    If Not IsNull(rs!bil_barang) Then Report78.Sections("Section5").Controls("L15").Caption = rs!bil_barang 'Bilangan barang
    If Not IsNull(rs!Berat_Asal) Then Report78.Sections("Section5").Controls("L16").Caption = Format(rs!Berat_Asal, "#,##0.00 g") 'Berat Asal (g)
    If Not IsNull(rs!kadar_tukaran) Then Report78.Sections("Section5").Controls("L19").Caption = rs!kadar_tukaran 'Mutu
    If Not IsNull(rs!berat_tukaran) Then Report78.Sections("Section5").Controls("L20").Caption = Format(rs!berat_tukaran, "#,##0.00 g") 'Berat 999.9 (g)
    If Not IsNull(rs!jumlah_gst) Then Report78.Sections("Section5").Controls("L13").Caption = Format(rs!jumlah_gst, "#,##0.00") 'Jumlah GST
    If Not IsNull(rs!harga_dengan_gst) Then Report78.Sections("Section5").Controls("L14").Caption = Format(rs!harga_dengan_gst, "#,##0.00") 'Jumlah keseluruhan (Upah + GST)
    If Not IsNull(rs!gst_sr_harga) Then Report78.Sections("Section5").Controls("L8").Caption = Format(rs!gst_sr_harga, "#,##0.00") 'Jumlah harga SR
    If Not IsNull(rs!gst_zr_harga) Then Report78.Sections("Section5").Controls("L9").Caption = Format(rs!gst_zr_harga, "#,##0.00") 'Jumlah harga ZR
    If Not IsNull(rs!gst_sr_cukai) Then Report78.Sections("Section5").Controls("L10").Caption = Format(rs!gst_sr_cukai, "#,##0.00") 'Jumlah GST SR
    If Not IsNull(rs!gst_zr_cukai) Then Report78.Sections("Section5").Controls("L11").Caption = Format(rs!gst_zr_cukai, "#,##0.00") 'Jumlah GST ZR
    If Not IsNull(rs!supplier_agen) Then frm123_LM_CUST = rs!supplier_agen
    If Not IsNull(rs!no_rujukan_supplier) Then Report78.Sections("Section4").Controls("L21").Caption = "No. Rujukan Supplier         : " & rs!no_rujukan_supplier 'No. Rujukan Dari Supplier

End If

rs.Close
Set rs = Nothing

If frm123_LM_CUST <> vbNullString Then
 
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from setting_database where Supplier='" & frm123_LM_CUST & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        
        If Not IsNull(rs!supplier) Then Report78.Sections("Section4").Controls("L3").Caption = rs!supplier 'Nama Pembeli
        If Not IsNull(rs!no_tel_hp) Then Report78.Sections("Section4").Controls("L4").Caption = rs!no_tel_hp 'No. Telefon
        If Not IsNull(rs!no_id_gst) Then Report78.Sections("Section4").Controls("L18").Caption = rs!no_id_gst 'No. ID GST

    End If
    
    rs.Close
    Set rs = Nothing
    
End If
   
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 79_grn where no_rujukan='" & G_No_RESIT_JUALAN & "' AND status = 1 order by status ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report78.DataSource = rs
    Report78.Show
    
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing
End Sub
Sub Frm115_cetak_voucher()
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
'        PRINTER_FOUND = 1 '0 : Printer Not Found , 1 : Printer Found
'        Exit For
'    End If
'Next

Frm101_LM_NO_INVOICE = G_No_RESIT_JUALAN
Frm115_LM_No_CUST = vbNullString
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
    If Not IsNull(rs!no_rujukan_pembeli) Then Frm115_LM_No_CUST = rs!no_rujukan_pembeli
End If

rs.Close
Set rs = Nothing

'### Data jika pembeli adalah berdaftar ### - Start
If Frm115_LM_No_CUST <> vbNullString Then
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm115_LM_No_CUST & "'", cn, adOpenKeyset, adLockOptimistic
    
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
Sub Frm115_padam_voucher()
'on error resume next
Dim rs2 As ADODB.Recordset
Dim Frm85_LM_BERAT_ASAL As Double
Dim Frm85_LM_BEZA_BERAT As Double
Dim Frm85_LM_BERAT_JUALAN As Double
Dim Frm85_SUSUT_BERAT As Double

LM_FOUND = 0
'### Padam data GDN ### - Start
LM_NOW = Now
LM_TARIKH = DateTime.Date$
LM_MASA = DateTime.Time$

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 77_gdn_grn where no_rujukan='" & G_No_RESIT_JUALAN & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    G_ID = rs!ID
    Call recovery_77_gdn_grn
    
    rs!tarikh = LM_TARIKH
    rs!masa = LM_MASA
    rs!write_timestamp = LM_NOW
    rs!Status = 0
    rs!terminal = G_TERMINAL
    rs!user = G_LOGIN_USER 'Nama Pekerja
    rs.Update
    LM_FOUND = 1
    
End If

rs.Close
Set rs = Nothing
'### Padam data GDN ### - End

If LM_FOUND = 1 Then
    '### Pulangkan stok barang kedai ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 23_senarai_jualan where no_resit='" & G_No_RESIT_JUALAN & "' AND status_rekod = 1", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
    
        Frm85_LM_BERAT_ASAL = 0
        Frm85_LM_BEZA_BERAT = 0
        Frm85_LM_BERAT_JUALAN = 0
        Frm85_SUSUT_BERAT = 0
        
        Set rs2 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs2.Open "select * from data_database where no_siri_Produk='" & rs!no_siri_Produk & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs2.EOF Then
        
            G_ID = rs2!ID
            Call recovery_data_database
                        
            If rs2!receiving_Status = 0 Or rs2!receiving_Status = 2 Then
                If Not IsNull(rs2!Berat) Then Frm85_LM_BERAT_ASAL = rs2!Berat
                If Not IsNull(rs2!beza_berat) Then Frm85_LM_BEZA_BERAT = rs2!beza_berat
                If Not IsNull(rs!berat_jualan) Then Frm85_LM_BERAT_JUALAN = rs!berat_jualan
                If Not IsNull(rs2!susut_berat) Then Frm85_SUSUT_BERAT = Format(rs2!susut_berat, "0.00") 'Susut berat
                
                Frm85_LM_BERAT_ASAL_COMP = Format(Frm85_LM_BERAT_ASAL, "0.00")
                Frm85_LM_BERAT_SELEPAS_COMP = Format(Frm85_LM_BERAT_JUALAN + Frm85_LM_BEZA_BERAT - Frm85_SUSUT_BERAT, "0.00")
                
                If Frm85_LM_BERAT_ASAL_COMP = Frm85_LM_BERAT_SELEPAS_COMP Then
                    rs2!beza_berat = Format(Frm85_LM_BERAT_JUALAN + Frm85_LM_BEZA_BERAT, "0.00")
                    rs2!StatusItem = 10
                Else
                    rs2!beza_berat = Format(Frm85_LM_BERAT_ASAL - Frm85_SUSUT_BERAT, "0.00")
                    rs2!StatusItem = 12
                End If
            Else
                rs2!StatusItem = 10
            End If
            
            rs2!write_timestamp2 = LM_NOW
            rs2!no_pekerja = G_LOGIN_USER
            rs2!terminal = G_TERMINAL
            rs2!Menu = 7
            
            rs2.Update
            
        End If
        
        rs2.Close
        Set rs2 = Nothing
    
        rs!write_timestamp2 = LM_NOW
        rs!no_staff = G_LOGIN_USER
        rs!status_rekod = 0
        rs.Update
        
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
    '#### Update Log Aktiviti Sistem #### - Start
    user = MDI_frm1.L3_Text
    LogAct_Memory = "[" & user & "] Padam data GDN. No. Rujukan [" & G_No_RESIT_JUALAN & "]."
    LogDate_Memory = LM_NOW
    Call UpdateLog_Database
    '#### Update Log Aktiviti Sistem #### - End
    
    Note = "Data telah berjaya dipadamkan." & vbCrLf & _
            "Refresh data anda ? Sistem akan mengambil sedikit masa untuk refresh data." & vbCrLf & _
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
    
End If
'### Pulangkan stok barang kedai ### - End
End Sub
Sub Frm115_jurujual()
'On Error Resume Next
If MDI_frm1.L3_Text <> vbNullString Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        Frm115.CBB4 = rs!Samaran & "  |  " & rs!NoPekerja
        
    Else
        
        Frm115.CBB4.AddItem "" & "  |  " & rs!Samaran
        Frm115.CBB4 = "" & "  |  " & rs!Samaran
        
    End If
    
    rs.Close
    Set rs = Nothing
    
    If G_LOCK_JURUJUAL = "YES" Then
    
        Frm115.CBB4.Enabled = False
        Frm115.CBB4.BackColor = &H8000000A

    Else
    
        Frm115.CBB4.Enabled = True
        Frm115.CBB4.BackColor = &HFFFFFF

    End If

End If
End Sub
Sub frm115_initial_setting_stok()
'on error resume next
Frm115.Frame1.Left = 9200
Frm115.Frame1.Top = 120

Frm115.CBB5.Clear
Frm115.CBB6.Clear

Frm115.CBB5.AddItem "Semua kategori produk"
Frm115.CBB6.AddItem "Semua purity"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database order by kategori_produk ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    If Not IsNull(rs!kategori_Produk) Then Frm115.CBB5.AddItem rs!kategori_Produk

    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database order by Kod_Metal_Purity ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    If Not IsNull(rs!Kod_Metal_Purity) Then Frm115.CBB6.AddItem rs!Kod_Metal_Purity

    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm115.CBB5 = "Semua kategori produk"
Frm115.CBB6 = "Semua purity"

Frm115.L55_Text = "Semua kategori produk"
Frm115.L56_Text = "Semua purity"
End Sub
Sub frm115_reset_gdn_list()
'on error resume next
GM_NEXT_PREV = 0

Frm115.L63_Text = -1 'Titik Pencarian Data
Frm115.L64_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Frm115.L61_Text = 0 'Paparan Page ke-xxx

Call frm115_gdn_list_header
Call frm115_gdn_list
End Sub
Sub frm115_gdn_list_header()
'on error resume next
With Frm115.LV1
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear

    Frm115.LV1.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "No. Siri Produk", 2000
    .ColumnHeaders.Add 5, , "Kategori Produk", 3300
    .ColumnHeaders.Add 6, , "Purity", 1500
    .ColumnHeaders.Add 7, , "Berat (g)", 1500, 1
    .ColumnHeaders.Add 8, , "Upah (RM)", 1500, 1
    .ColumnHeaders.Add 9, , "Status", 1400


End With
End Sub
Sub frm115_gdn_list()
'on error resume next
Dim frm115_LM_TOTAL_PAGE As Double

frm115_PAGE_SIZE = 31
frm115_LM_TOTAL_PAGE = 0
x = 0

re_gen_report:

Frm115.L60_Text = 0

LM_START_ROW = Frm115.L63_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + frm115_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm115.L64_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - frm115_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        Frm115.L61_Text = 1
    End If
End If

If Frm115.L55_Text = "Semua kategori produk" Then
    frm115_LM_SEARCH_1 = Null
    frm115_LM_SEARCH_1_LOGIC = "<>"
Else
    frm115_LM_SEARCH_1 = Frm115.L55_Text
    frm115_LM_SEARCH_1_LOGIC = "="
End If
If Frm115.L56_Text = "Semua purity" Then
    frm115_LM_SEARCH_2 = Null
    frm115_LM_SEARCH_2_LOGIC = "<>"
Else
    frm115_LM_SEARCH_2 = Frm115.L56_Text
    frm115_LM_SEARCH_2_LOGIC = "="
End If
If MDI_frm1.L20_Text = "Semua cawangan" Then

    Frm115_SEARCH_3 = Null
    Frm115_SEARCH_3_LOGIC = "<>"
    Frm115_SEARCH_4 = Null
    Frm115_SEARCH_4_LOGIC = "<>"
    
Else

    Frm115_SEARCH_3 = MDI_frm1.L20_Text
    Frm115_SEARCH_3_LOGIC = "="
    Frm115_SEARCH_4 = "HQ"
    Frm115_SEARCH_4_LOGIC = "="
    
End If

Frm115.L28_Text = "Paparan stok mengikut kategori produk [" & Frm115.L55_Text & "] dan purity [" & Frm115.L56_Text & "]."

frm115_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from data_database where statusitem = 10 AND (receiving_Status = 0 Or receiving_Status = 1 Or receiving_Status = 4) AND kategori_Produk " & frm115_LM_SEARCH_1_LOGIC & "'" & frm115_LM_SEARCH_1 & "' AND kod_Purity " & frm115_LM_SEARCH_2_LOGIC & "'" & frm115_LM_SEARCH_2 & "' order by no_siri_produk ASC LIMIT " & LM_START_ROW & "," & frm115_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
rs.Open "select * from data_database where (cawangan " & Frm115_SEARCH_3_LOGIC & "'" & Frm115_SEARCH_3 & "' OR cawangan " & Frm115_SEARCH_4_LOGIC & "'" & Frm115_SEARCH_4 & "') AND (statusitem = 10 OR statusitem = 12 OR statusitem = 20 OR statusitem = 22 OR statusitem = 28) AND (receiving_Status = 0 Or receiving_Status = 2) AND kategori_Produk " & frm115_LM_SEARCH_1_LOGIC & "'" & frm115_LM_SEARCH_1 & "' AND kod_Purity " & frm115_LM_SEARCH_2_LOGIC & "'" & frm115_LM_SEARCH_2 & "' order by no_siri_produk ASC , ID ASC LIMIT " & LM_START_ROW & "," & frm115_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1
    If frm115_LM_PAGE_FOUND = 0 Then
        If Frm115.L64_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm115.L61_Text = Frm115.L61_Text + 1 'Paparan Page ke-xxx
                frm115_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm115.L61_Text) Then
                    If Frm115.L61_Text <> 1 Then
                        Frm115.L61_Text = Frm115.L61_Text - 1 'Paparan Page ke-xxx
                        frm115_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If

    Y = ((Frm115.L61_Text - 1) * frm115_PAGE_SIZE) + x

    With Frm115.LV1.ListItems.Add(, , rs!ID)
     
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID

        If Not IsNull(rs!no_siri_Produk) Then  'No. Siri Produk
            .ListSubItems.Add , , rs!no_siri_Produk
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kategori_Produk) Then 'Kategori Produk
            .ListSubItems.Add , , rs!kategori_Produk
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kod_Purity) Then 'Purity
            .ListSubItems.Add , , rs!kod_Purity
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!beza_berat) Then 'Berat (g)
            .ListSubItems.Add , , Format(rs!beza_berat, "#,##0.00")
        Else
            .ListSubItems.Add , , "0.00"
        End If
        
        'If Not IsNull(rs!Upah_Jualan) Then 'Upah (RM)
        '    .ListSubItems.Add , , Format(rs!Upah_Jualan, "#,##0.00")
        'Else
        '    .ListSubItems.Add , , "0.00"
        'End If
        
        If Not IsNull(rs!UPAH) Then 'Upah (RM)
            .ListSubItems.Add , , Format(rs!UPAH, "#,##0.00")
        Else
            .ListSubItems.Add , , "0.00"
        End If
        
        If Not IsNull(rs!gdn_temp) Then
            
            If rs!gdn_temp = 0 Then
                .ListSubItems.Add , , "Belum Dipilih"
            ElseIf rs!gdn_temp = 1 Then
                .ListSubItems.Add , , "Telah Dipilih"
            End If
            
        Else
        
            .ListSubItems.Add , , "Belum Dipilih"
            
        End If
        
    End With
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select COUNT(ID) from data_database where statusitem = 10 AND kategori_Produk " & frm115_LM_SEARCH_1_LOGIC & "'" & frm115_LM_SEARCH_1 & "' AND kod_Purity " & frm115_LM_SEARCH_2_LOGIC & "'" & frm115_LM_SEARCH_2 & "'", cn, adOpenKeyset, adLockOptimistic
rs.Open "select COUNT(ID) from data_database where (cawangan " & Frm115_SEARCH_3_LOGIC & "'" & Frm115_SEARCH_3 & "' OR cawangan " & Frm115_SEARCH_4_LOGIC & "'" & Frm115_SEARCH_4 & "') AND (statusitem = 10 OR statusitem = 12 OR statusitem = 20 OR statusitem = 22 OR statusitem = 28) AND (receiving_Status = 0 Or receiving_Status = 2) AND kategori_Produk " & frm115_LM_SEARCH_1_LOGIC & "'" & frm115_LM_SEARCH_1 & "' AND kod_Purity " & frm115_LM_SEARCH_2_LOGIC & "'" & frm115_LM_SEARCH_2 & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    frm115_LM_TOTAL_PAGE = Format(rs(0) / frm115_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, frm115_LM_TOTAL_PAGE, ".") <> 0 Then
    
        frm115_LM_PAGE = Split(frm115_LM_TOTAL_PAGE, ".")(0)
        frm115_LM_PAGE_LEBIHAN = Split(frm115_LM_TOTAL_PAGE, ".")(1)
        
        If frm115_LM_PAGE_LEBIHAN <> "00" Then
            Frm115.L62_Text = frm115_LM_PAGE + 1
        Else
            Frm115.L62_Text = frm115_LM_PAGE
        End If
        
    Else
    
        Frm115.L62_Text = frm115_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm115.L62_Text = 0
    End If
Else
    Frm115.L62_Text = 0
End If

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select COUNT(ID) from data_database where statusitem = 10 AND (receiving_Status = 0 Or receiving_Status = 1 Or receiving_Status = 4) AND kategori_Produk " & frm115_LM_SEARCH_1_LOGIC & "'" & frm115_LM_SEARCH_1 & "' AND kod_Purity " & frm115_LM_SEARCH_2_LOGIC & "'" & frm115_LM_SEARCH_2 & "'", cn, adOpenKeyset, adLockOptimistic
rs.Open "select COUNT(ID) from data_database where (cawangan " & Frm115_SEARCH_3_LOGIC & "'" & Frm115_SEARCH_3 & "' OR cawangan " & Frm115_SEARCH_4_LOGIC & "'" & Frm115_SEARCH_4 & "') AND (statusitem = 10 OR statusitem = 12 OR statusitem = 20 OR statusitem = 22 OR statusitem = 28) AND (receiving_Status = 0 Or receiving_Status = 2) AND kategori_Produk " & frm115_LM_SEARCH_1_LOGIC & "'" & frm115_LM_SEARCH_1 & "' AND kod_Purity " & frm115_LM_SEARCH_2_LOGIC & "'" & frm115_LM_SEARCH_2 & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm115.L60_Text = rs(0) 'Jumlah bilangan

rs.Close
Set rs = Nothing

If x <> 0 Then
    Frm115.L63_Text = LM_START_ROW
End If

If Frm115.L61_Text <> vbNullString And IsNumeric(Frm115.L61_Text) Then
    If Frm115.L62_Text <> vbNullString And IsNumeric(Frm115.L62_Text) Then
        frm115_LM_CURR_PAGE = Frm115.L61_Text
        frm115_LM_TOTAL_PAGE = Frm115.L62_Text
        
        If frm115_LM_CURR_PAGE > frm115_LM_TOTAL_PAGE Then
            
            Frm115.L61_Text = Frm115.L61_Text - 1
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            GoTo re_gen_report:
            
        End If
    End If
End If

End Sub
Sub frm115_insert_data()
'On Error Resume Next
Dim Err(30)
Dim Frm115_LM_BERAT_ASAL As Double
Dim Frm115_LM_BERAT_JUAL As Double
Dim Frm115_LM_HARGA_MODAL As Double
Dim Frm115_LM_HARGA_JUAL As Double
Dim Frm115_LM_HARGA_SEMASA_MODAL As Double
Dim Frm115_LM_TETAPANHARGA As Double
Dim Frm115_LM_LIMIT As Double
Dim Frm115_LM_HARGA_STAFF As Double 'Tetapan harga jualan kepada staff
Dim Frm115_LM_HARGA_PELANGGAN As Double 'Tetapan harga jualan kepada pelanggan
Dim Frm115_LM_HARGA_SEMASA As Double 'Harga semasa (jualan)
Dim Frm115_LM_BERAT_JUAL_ASAL As Double 'Berat Jualan (Purity Asal)
Dim Frm115_LM_HARGA_SEMASA_999 As Double 'Harga semasa (jualan) (Purity 999.9)
Dim Frm115_LM_HARGA_SUPPLIER As Double 'Harga per gram (harga semasa) dari supplier (modal)
Dim Frm115_LM_BERAT_999 As Double 'Berat Jualan (Purity Asal)
Dim Frm115_UPAH_MODAL As Double 'Upah modal
Dim Frm115_UPAH_JUAL As Double 'Upah jualan
Dim LM_KADAR_TUKARAN As Double

LM_KADAR_TUKARAN = 0
Frm115_UPAH_MODAL = 0 'Upah modal
Frm115_UPAH_JUAL = 0 'Upah jualan
Frm115_LM_BERAT_JUAL_ASAL = 0 'Berat Jualan (Purity Asal)
Frm115_LM_HARGA_SEMASA_999 = 0 'Harga semasa (jualan) (Purity 999.9)
Frm115_LM_HARGA_SUPPLIER = 0 'Harga per gram (harga semasa) dari supplier (modal)
Frm115_LM_BERAT_999 = 0 'Berat Jualan (Purity Asal)

Frm115_LM_HARGA_SEMASA = 0 'Harga semasa (jualan)
Frm115_LM_HARGA_SUPPLIER = 0 'Harga per gram (harga semasa) dari supplier (modal)

x = 0
Frm115_LM_BERAT_ASAL = 0
Frm115_LM_BERAT_JUAL = 0
Frm115_LM_DATA_SAVE = 0
Frm115_LM_HARGA_MODAL = 0
Frm115_LM_HARGA_JUAL = 0
Frm115_LM_HARGA_SEMASA_MODAL = 0
Frm115_LM_PRICE_CHECK = 0 '0 : Tidak Perlu Periksa Harga Semasa Jualan , 1 : Perlu Periksa Harga Semasa Jualan
Frm115_LM_TETAPANHARGA = 0
Frm115_LM_LIMIT = 0
Frm115_LM_HARGA_STAFF = 0
Frm115_LM_HARGA_PELANGGAN = 0

If Frm115.L3_Text = vbNullString Then
    x = x + 1
    Err(x) = "Sila masukkan [No. Siri Produk]."
End If
If Frm115.L33_Text = vbNullString Then
    x = x + 1
    Err(x) = "Maklumat harga semasa modal belian item ini yang tidak sah. Sila keluar dari menu ini dan cuba sekali lagi."
End If
If Frm115.L50_Text = vbNullString Or (Frm115.L50_Text <> vbNullString And Not IsNumeric(Frm115.L50_Text)) Then
    x = x + 1
    Err(x) = "Maklumat upah modal yang tidak sah. Sila keluar dari menu ini dan cuba sekali lagi."
End If
If Frm115.L6_Text = vbNullString Or (Frm115.L6_Text <> vbNullString And Not IsNumeric(Frm115.L6_Text)) Then
    x = x + 1
    Err(x) = "Sila maklumat [Berat Asal]. Sila scan item sekali lagi."
End If
If Frm115.TB3 = vbNullString Or (Frm115.TB3 <> vbNullString And Not IsNumeric(Frm115.TB3)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Berat Jualan]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm115.TB2 = vbNullString Or (Frm115.TB2 <> vbNullString And Not IsNumeric(Frm115.TB2)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Harga Semasa Emas 999.9]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm115.TB2 <> vbNullString And IsNumeric(Frm115.TB2) Then

    If Format(Frm115.TB2, "0.00") = "0.00" Then
        x = x + 1
        Err(x) = "Harga emas semasa 999.9 yang tidak sah. Nilai 0.00 tidak dibenarkan."
    End If
    
End If
If Frm115.TB7 = vbNullString Or (Frm115.TB7 <> vbNullString And Not IsNumeric(Frm115.TB7)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Kadar Tukaran Purity 999.9]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If (Frm115.TB7 <> vbNullString And IsNumeric(Frm115.TB7)) Then
    
    LM_KADAR_TUKARAN = Frm115.TB7
    
    If LM_KADAR_TUKARAN > 1 Then
    
        x = x + 1
        Err(x) = "[Kadar Tukaran Purity 999.9] tidak boleh lebih dari 1.00. Hanya NOMBOR dibenarkan dalam ruangan ini."
        
    End If
    
End If
If Frm115.L7_Text = vbNullString Or (Frm115.L7_Text <> vbNullString And Not IsNumeric(Frm115.L7_Text)) Then
    x = x + 1
    Err(x) = "[Berat 999.9] yang tidak sah. Sila scan item sekali lagi."
End If
If Frm115.TB4 = vbNullString Or (Frm115.TB4 <> vbNullString And Not IsNumeric(Frm115.TB4)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Upah]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm115.CB2 = 0 And Frm115.CB3 = 0 And Frm115.CB4 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan jenis GST bagi upah."
End If
If Frm115.TB5 = vbNullString Or Frm115.TB6 = vbNullString Then
    x = x + 1
    Err(x) = "Maklumat berkenaan GST yang tidak sah. Sila keluar dari menu ini dan cuba sekali lagi."
End If

If (Frm115.L6_Text <> vbNullString And IsNumeric(Frm115.L6_Text)) And (Frm115.TB3 <> vbNullString And IsNumeric(Frm115.TB3)) Then
    Frm115_LM_BERAT_ASAL = Frm115.L6_Text 'Berat Asal
    Frm115_LM_BERAT_JUAL = Frm115.TB3 'Berat Jualan
    
    If Frm115_LM_BERAT_JUAL > Frm115_LM_BERAT_ASAL Then
        x = x + 1
        Err(x) = "Berat jualan melebihi berat asal."
    End If
End If
If Frm115.L49_Text = vbNullString Or (Frm115.L49_Text <> vbNullString And Not IsNumeric(Frm115.L49_Text)) Then
    x = x + 1
    Err(x) = "Tiada maklumat harga dari supplier bagi purity ini." & vbCrLf & _
                "Sila pastikan harga dari supplier bagi purity ini telah ditetapkan dalam TETAPAN HARIAN SISTEM."
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else

'### Periksa Data Dulang ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from Data_Database where no_siri_Produk='" & Frm115.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!dulang) Then Frm115_LM_DULANG = rs!dulang 'Dulang
    End If
    
    rs.Close
    Set rs = Nothing
'### Periksa Data Dulang ### - End
    
'### Masukkan Data Ke Dalam Temp Table ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from " & G_GDN_TEMP & " where no_siri_Produk='" & Frm115.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If rs.EOF Then
    
        rs.AddNew
        If Frm115.L3_Text <> vbNullString Then
            rs!no_siri_Produk = Frm115.L3_Text 'No. Siri Produk
        Else
            rs!no_siri_Produk = Null 'No. Siri Produk
        End If
        If Frm115.L5_Text <> vbNullString Then
            rs!kategori_Produk = Frm115.L5_Text 'Kategori Produk
        Else
            rs!kategori_Produk = Null 'Kategori Produk
        End If
        If Frm115.L4_Text <> vbNullString Then
            rs!purity = Frm115.L4_Text 'Purity
        Else
            rs!purity = Null 'Purity
        End If
        If Frm115.L6_Text <> vbNullString Then
            rs!Berat_Asal = Format(Frm115.L6_Text, "0.00") 'Berat Asal (g)
        Else
            rs!Berat_Asal = Null 'Berat Asal (g)
        End If
        If Frm115.TB3 <> vbNullString Then
            rs!berat_jualan = Format(Frm115.TB3, "0.00") 'Berat Jualan (g)
        Else
            rs!berat_jualan = Null 'Berat Jualan (g)
        End If
        If Frm115.TB2 <> vbNullString Then
            rs!harga_Semasa = Format(Frm115.TB2, "0.00") 'Harga Semasa (RM/g)
        Else
            rs!harga_Semasa = Null 'Harga Semasa (RM/g)
        End If
        If Frm115.TB4 <> vbNullString Then
            rs!UPAH = Format(Frm115.TB4, "0.00") 'Upah (RM)
        Else
            rs!UPAH = Null 'Upah (RM)
        End If
        
        Frm115_LM_HARGA_SEMASA = Frm115.TB2 'Harga emas semasa 999.9 (Untuk tujuan jualan kepada pelanggan)
        Frm115_LM_BERAT_JUALAN_9999 = Frm115.L7_Text 'Berat jualan dalam purity 999.9
        Frm115_LM_UPAH_DAN_GST = Frm115.TB6 'Jumlah Upah + GST (Bagi jualan setiap item)

        If Frm115.TB6 <> vbNullString Then
            rs!harga_asal = Format((Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
        Else
            rs!harga_asal = Null 'Harga Asal Item (RM)
        End If
        
        rs!diskaun = "0.00" 'Diskaun (%)
        rs!harga_lepas_diskaun = Format((Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
        rs!adjustment = Format(0, "0.00") 'Adjustment (RM)
        rs!harga_jualan = Format((Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
        rs!harga_jualan_dengan_gst = Format((Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
        
        If Frm115.CB2 = 1 Then
            rs!gst_ari_nashi = "ZR (L)" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
            rs!kadar_gst = Null 'Kadar Cukai GST (%)
            If Frm115.TB5 <> vbNullString Then
                rs!jumlah_gst = Format(Frm115.TB5, "0.00") 'Jumlah GST (Bagi jualan setiap item)
            Else
                rs!jumlah_gst = "0.00" 'Jumlah GST (Bagi jualan setiap item)
            End If
        ElseIf Frm115.CB3 = 1 Or Frm115.CB4 = 1 Then
            rs!gst_ari_nashi = "SR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
            If Frm115.L21_Text <> vbNullString Then
                rs!kadar_gst = Frm115.L21_Text 'Kadar Cukai GST (%)
            Else
                rs!kadar_gst = "0" 'Jumlah Cukai GST (RM)
            End If
            If Frm115.TB5 <> vbNullString Then
                rs!jumlah_gst = Format(Frm115.TB5, "0.00") 'Jumlah GST (Bagi jualan setiap item)
            Else
                rs!jumlah_gst = "0.00" 'Jumlah GST (Bagi jualan setiap item)
            End If
            If Frm115.CB4 = 1 Then 'Jenis Cukai GST SR
                rs!gst_include = 1 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
            Else
                rs!gst_include = 0 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
            End If
        End If
        If Frm115.L30_Text <> vbNullString Then
            rs!harga_tanpa_gst = Format(Frm115.L30_Text, "0.00") 'Harga Jualan Tanpa GST (RM)
        Else
            rs!harga_tanpa_gst = Null 'Harga Jualan Tanpa GST (RM)
        End If
        If Frm115.TB6 <> vbNullString Then
            rs!harga_dengan_gst = Format(Frm115.TB6, "0.00") 'Harga Jualan Termasuk GST (RM)
        Else
            rs!harga_dengan_gst = Null 'Harga Jualan Termasuk GST (RM)
        End If
        rs!dropship = 0 '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
        rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g)
        rs!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini
        If Frm115.L32_Text = "0" Then '0 : Menu Data Baru , 1 : Menu Edit Data
            rs!Status = 1
        ElseIf Frm115.L32_Text = "1" Then
            rs!Status = 4
        End If
        rs!Type = 0 '0 : BK , 1 : Barang Permata
        If Frm115.L33_Text <> vbNullString Then
            rs!harga_per_gram_modal = Format(Frm115.L33_Text, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
            Frm115_LM_HARGA_SEMASA_MODAL = Frm115.L33_Text
        Else
            rs!harga_per_gram_modal = Format(0, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
        End If
        rs!modal = Format(Frm115_LM_HARGA_SEMASA_MODAL * Frm115_LM_BERAT_JUALAN_9999, "0.00") 'Harga Modal (RM)
        If IsNumeric(Frm115.TB6) And IsNumeric(Frm115.L33_Text) And IsNumeric(Frm115.TB3) Then
            Frm115_LM_HARGA_MODAL = Frm115.L33_Text * Frm115.TB3 'Harga modal
            Frm115_LM_HARGA_JUAL = (Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST 'Harga jualan
            
            rs!untung = Format(Frm115_LM_HARGA_JUAL - Frm115_LM_HARGA_MODAL, "0.00") 'Jumlah Keuntungan
        Else
            rs!untung = Format(0, "0.00") 'Jumlah Keuntungan
        End If
        
        If Frm115.L49_Text <> vbNullString Then 'Harga per gram (harga semasa) dari supplier (modal)
            rs!harga_per_gram_supplier = Frm115.L49_Text
        Else
            rs!harga_per_gram_supplier = 0
        End If
        
        If IsNumeric(Frm115.TB3) And IsNumeric(Frm115.TB2) And IsNumeric(Frm115.L49_Text) And IsNumeric(Frm115.L7_Text) And IsNumeric(Frm115.L6_Text) And IsNumeric(Frm115.L50_Text) And IsNumeric(Frm115.TB4) Then
            Frm115_LM_BERAT_JUAL_ASAL = Frm115.TB3 'Berat Jualan (Purity Asal)
            Frm115_LM_BERAT_ASAL = Frm115.L6_Text 'Berat Asal (Purity Asal)
            Frm115_UPAH_JUAL = Frm115.TB4 'Upah jualan
            Frm115_UPAH_MODAL = Frm115.L50_Text 'Upah modal
            Frm115_LM_HARGA_SEMASA_999 = Frm115.TB2 'Harga semasa (jualan) (Purity 999.9)
            Frm115_LM_HARGA_SUPPLIER = Frm115.L49_Text 'Harga per gram (harga semasa) dari supplier (modal)
            Frm115_LM_BERAT_999 = Frm115.L7_Text 'Berat emas dalam purity 999.9
            
            rs!upah_modal = Frm115.L50_Text 'Upah modal
            rs!harga_per_gram_supplier = Frm115.L49_Text 'Harga per gram (harga semasa) dari supplier (modal)
            rs!untung2 = Format(((Frm115_LM_BERAT_999 * Frm115_LM_HARGA_SEMASA_999) + Frm115_UPAH_JUAL) - ((Frm115_LM_BERAT_JUAL_ASAL * Frm115_LM_HARGA_SUPPLIER) + (Frm115_LM_BERAT_JUAL_ASAL / Frm115_LM_BERAT_ASAL) * Frm115_UPAH_MODAL), "0.00") 'Untung jika restok pada harga supplier ini
            
        Else
        
            rs!harga_per_gram_supplier = "0.00" 'Harga per gram (harga semasa) dari supplier (modal)
            rs!untung2 = "0.00" 'Untung jika restok pada harga supplier ini
            rs!upah_modal = "0.00" 'Upah modal
            
        End If
            
        If Format(Frm115.L6_Text, "0.00") = Format(Frm115.TB3, "0.00") Then
            rs!potong_flag = 0 '0 : Tiada Potong , 1 : Ada Potong
        Else
            rs!potong_flag = 1 '0 : Tiada Potong , 1 : Ada Potong
        End If
        rs!dulang = Frm115_LM_DULANG 'Dulang
        If Frm115.TB7 <> vbNullString Then
            rs!pemalar_tukaran_999 = Frm115.TB7 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
        Else
            rs!pemalar_tukaran_999 = Format(0, "0.00") 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
        End If
        If Frm115.L7_Text <> vbNullString Then
            rs!berat_999 = Format(Frm115.L7_Text, "0.00") 'Berat jualan dalam purity 999.9
        Else
            rs!berat_999 = Null 'Berat jualan dalam purity 999.9
        End If
        rs!gst_barang_atau_upah = 1 '0 : GST pada harga jualan , 1 : GST pada upah
        
        rs.Update
        Frm115_LM_DATA_SAVE = 1
    Else
        If Frm115.L3_Text <> vbNullString Then
            rs!no_siri_Produk = Frm115.L3_Text 'No. Siri Produk
        Else
            rs!no_siri_Produk = Null 'No. Siri Produk
        End If
        If Frm115.L5_Text <> vbNullString Then
            rs!kategori_Produk = Frm115.L5_Text 'Kategori Produk
        Else
            rs!kategori_Produk = Null 'Kategori Produk
        End If
        If Frm115.L4_Text <> vbNullString Then
            rs!purity = Frm115.L4_Text 'Purity
        Else
            rs!purity = Null 'Purity
        End If
        If Frm115.L6_Text <> vbNullString Then
            rs!Berat_Asal = Format(Frm115.L6_Text, "0.00") 'Berat Asal (g)
        Else
            rs!Berat_Asal = Null 'Berat Asal (g)
        End If
        If Frm115.TB3 <> vbNullString Then
            rs!berat_jualan = Format(Frm115.TB3, "0.00") 'Berat Jualan (g)
        Else
            rs!berat_jualan = Null 'Berat Jualan (g)
        End If
        If Frm115.TB2 <> vbNullString Then
            rs!harga_Semasa = Format(Frm115.TB2, "0.00") 'Harga Semasa (RM/g)
        Else
            rs!harga_Semasa = Null 'Harga Semasa (RM/g)
        End If
        If Frm115.TB4 <> vbNullString Then
            rs!UPAH = Format(Frm115.TB4, "0.00") 'Upah (RM)
        Else
            rs!UPAH = Null 'Upah (RM)
        End If
        
        Frm115_LM_HARGA_SEMASA = Frm115.TB2 'Harga emas semasa 999.9 (Untuk tujuan jualan kepada pelanggan)
        Frm115_LM_BERAT_JUALAN_9999 = Frm115.L7_Text 'Berat jualan dalam purity 999.9
        Frm115_LM_UPAH_DAN_GST = Frm115.TB6 'Jumlah Upah + GST (Bagi jualan setiap item)

        If Frm115.TB6 <> vbNullString Then
            rs!harga_asal = Format((Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
        Else
            rs!harga_asal = Null 'Harga Asal Item (RM)
        End If
        
        rs!diskaun = "0.00" 'Diskaun (%)
        rs!harga_lepas_diskaun = Format((Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
        rs!adjustment = Format(0, "0.00") 'Adjustment (RM)
        rs!harga_jualan = Format((Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
        rs!harga_jualan_dengan_gst = Format((Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST, "0.00") 'Harga Asal Item (RM)
        
        If Frm115.CB2 = 1 Then
            rs!gst_ari_nashi = "ZR (L)" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
            rs!kadar_gst = Null 'Kadar Cukai GST (%)
            If Frm115.TB5 <> vbNullString Then
                rs!jumlah_gst = Format(Frm115.TB5, "0.00") 'Jumlah GST (Bagi jualan setiap item)
            Else
                rs!jumlah_gst = "0.00" 'Jumlah GST (Bagi jualan setiap item)
            End If
        ElseIf Frm115.CB3 = 1 Then
            rs!gst_ari_nashi = "SR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
            If Frm115.L21_Text <> vbNullString Then
                rs!kadar_gst = Frm115.L21_Text 'Kadar Cukai GST (%)
            Else
                rs!kadar_gst = "0" 'Jumlah Cukai GST (RM)
            End If
            If Frm115.TB5 <> vbNullString Then
                rs!jumlah_gst = Format(Frm115.TB5, "0.00") 'Jumlah GST (Bagi jualan setiap item)
            Else
                rs!jumlah_gst = "0.00" 'Jumlah GST (Bagi jualan setiap item)
            End If
            If Frm115.CB4 = 1 Then 'Jenis Cukai GST SR
                rs!gst_include = 1 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
            Else
                rs!gst_include = 0 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
            End If
        End If
        If Frm115.L30_Text <> vbNullString Then
            rs!harga_tanpa_gst = Format(Frm115.L30_Text, "0.00") 'Harga Jualan Tanpa GST (RM)
        Else
            rs!harga_tanpa_gst = Null 'Harga Jualan Tanpa GST (RM)
        End If
        If Frm115.TB6 <> vbNullString Then
            rs!harga_dengan_gst = Format(Frm115.TB6, "0.00") 'Harga Jualan Termasuk GST (RM)
        Else
            rs!harga_dengan_gst = Null 'Harga Jualan Termasuk GST (RM)
        End If
        rs!dropship = 0 '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
        rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g)
        rs!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini
        If Frm115.L32_Text = "0" Then '0 : Menu Data Baru , 1 : Menu Edit Data
            rs!Status = 1
        ElseIf Frm115.L32_Text = "1" Then
            rs!Status = 3
        End If
        rs!Type = 0 '0 : BK , 1 : Barang Permata
        If Frm115.L33_Text <> vbNullString Then
            rs!harga_per_gram_modal = Format(Frm115.L33_Text, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
            Frm115_LM_HARGA_SEMASA_MODAL = Frm115.L33_Text
        Else
            rs!harga_per_gram_modal = Format(0, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
        End If
        rs!modal = Format(Frm115_LM_HARGA_SEMASA_MODAL * Frm115_LM_BERAT_JUALAN_9999, "0.00") 'Harga Modal (RM)
        If IsNumeric(Frm115.TB6) And IsNumeric(Frm115.L33_Text) And IsNumeric(Frm115.TB3) Then
            Frm115_LM_HARGA_MODAL = Frm115.L33_Text * Frm115.TB3 'Harga modal
            Frm115_LM_HARGA_JUAL = (Frm115_LM_HARGA_SEMASA * Frm115_LM_BERAT_JUALAN_9999) + Frm115_LM_UPAH_DAN_GST 'Harga jualan
            
            rs!untung = Format(Frm115_LM_HARGA_JUAL - Frm115_LM_HARGA_MODAL, "0.00") 'Jumlah Keuntungan
        Else
            rs!untung = Format(0, "0.00") 'Jumlah Keuntungan
        End If
        If Frm115.L49_Text <> vbNullString Then 'Harga per gram (harga semasa) dari supplier (modal)
            rs!harga_per_gram_supplier = Frm115.L49_Text
        Else
            rs!harga_per_gram_supplier = 0
        End If
        
        If IsNumeric(Frm115.TB3) And IsNumeric(Frm115.TB2) And IsNumeric(Frm115.L49_Text) And IsNumeric(Frm115.L7_Text) And IsNumeric(Frm115.L6_Text) And IsNumeric(Frm115.L50_Text) And IsNumeric(Frm115.TB4) Then
            Frm115_LM_BERAT_JUAL_ASAL = Frm115.TB3 'Berat Jualan (Purity Asal)
            Frm115_LM_BERAT_ASAL = Frm115.L6_Text 'Berat Asal (Purity Asal)
            Frm115_UPAH_JUAL = Frm115.TB4 'Upah jualan
            Frm115_UPAH_MODAL = Frm115.L50_Text 'Upah modal
            Frm115_LM_HARGA_SEMASA_999 = Frm115.TB2 'Harga semasa (jualan) (Purity 999.9)
            Frm115_LM_HARGA_SUPPLIER = Frm115.L49_Text 'Harga per gram (harga semasa) dari supplier (modal)
            Frm115_LM_BERAT_999 = Frm115.L7_Text 'Berat emas dalam purity 999.9
            
            rs!upah_modal = Frm115.L50_Text 'Upah modal
            rs!harga_per_gram_supplier = Frm115.L49_Text 'Harga per gram (harga semasa) dari supplier (modal)
            rs!untung2 = Format(((Frm115_LM_BERAT_999 * Frm115_LM_HARGA_SEMASA_999) + Frm115_UPAH_JUAL) - ((Frm115_LM_BERAT_JUAL_ASAL * Frm115_LM_HARGA_SUPPLIER) + (Frm115_LM_BERAT_JUAL_ASAL / Frm115_LM_BERAT_ASAL) * Frm115_UPAH_MODAL), "0.00") 'Untung jika restok pada harga supplier ini
            
        Else
        
            rs!harga_per_gram_supplier = "0.00" 'Harga per gram (harga semasa) dari supplier (modal)
            rs!untung2 = "0.00" 'Untung jika restok pada harga supplier ini
            rs!upah_modal = "0.00" 'Upah modal
            
        End If
        
        If Format(Frm115.L6_Text, "0.00") = Format(Frm115.TB3, "0.00") Then
            rs!potong_flag = 0 '0 : Tiada Potong , 1 : Ada Potong
        Else
            rs!potong_flag = 1 '0 : Tiada Potong , 1 : Ada Potong
        End If
        rs!dulang = Frm115_LM_DULANG 'Dulang
        If Frm115.TB7 <> vbNullString Then
            rs!pemalar_tukaran_999 = Frm115.TB7 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
        Else
            rs!pemalar_tukaran_999 = Format(0, "0.00") 'Pemalar tukaran emas kepada 999.9 bagi urusan jualan
        End If
        If Frm115.L7_Text <> vbNullString Then
            rs!berat_999 = Format(Frm115.L7_Text, "0.00") 'Berat jualan dalam purity 999.9
        Else
            rs!berat_999 = Null 'Berat jualan dalam purity 999.9
        End If
        rs!gst_barang_atau_upah = 1 '0 : GST pada harga jualan , 1 : GST pada upah
        
        rs.Update
        Frm115_LM_DATA_SAVE = 1
    End If
    
    rs.Close
    Set rs = Nothing
'### Masukkan Data Ke Dalam Temp Table ### - End
    
    If Frm115_LM_DATA_SAVE = 1 Then
    
        GM_NEXT_PREV = 0
        
        Frm115.L69_Text = -1 'Titik Pencarian Data
        Frm115.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
        Frm115.L67_Text = 0 'Paparan Page ke-xxx

        Call Frm115_reset_1
        Call Frm115_Senarai_Jualan_Header
        Call Frm115_Senarai_Jualan
        
        'MsgBox "Data telah berjaya dimasukkan ke dalam senarai jualan.", vbInformation, "Info"
        
        Frm115.TB1.SetFocus
    End If
End If
End Sub


