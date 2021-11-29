Attribute VB_Name = "Module21"
Sub Frm103_initial_setting()
'on error resume next
Frm103.Frame1.Left = 120
Frm103.Frame1.Top = 240
Frm103.Frame2.Left = 120
Frm103.Frame2.Top = 240

Frm103.Frame1.Visible = False
Frm103.Frame2.Visible = False
End Sub
Sub Frm103_initial_setting2()
'on error resume next
Frm103.Frame4.Left = 120
Frm103.Frame4.Top = 840

Frm103.Frame3.Left = 120
Frm103.Frame3.Top = 600
Frm103.Frame4.Left = 120
Frm103.Frame4.Top = 600

Frm103.Frame3.Visible = False
Frm103.Frame4.Visible = False
End Sub
Sub Frm103_reset_penyata()
'on error resume next
Frm103.L17_Text = "0.00" 'Jualan (Termasuk GST)
Frm103.L18_Text = "0.00" 'GST Jualan
Frm103.L19_Text = "0.00" 'Jualan Bersih (Belum tolak adjustment)
Frm103.L20_Text = "0.00" 'Kos Modal (Termasuk GST)
Frm103.L21_Text = "0.00" 'GST Modal
Frm103.L22_Text = "0.00" 'Kos Bersih (Modal)
Frm103.L23_Text = "0.00" 'Komisyen Staff
Frm103.L24_Text = "0.00" 'Untung Bersih
Frm103.L26_Text = "0.00" 'Adjustment
Frm103.L27_Text = "0.00" 'Jualan bersih (Setelah tolak adjustment)
Frm103.L28_Text = "0.00" 'Jualan diskaun
Frm103.L29_Text = "0.00" 'Jualan kupon diskaun
Frm103.L30_Text = "0.00" 'Jualan tebus mata ganjaran
Frm103.L34_Text = "0.00" 'Untung Bersih
End Sub
Sub Frm103_kiraan_untung_rugi()
'on error resume next
Dim TA As Date
Dim TM As Date
Dim Frm103_LM_ADJUSTMENT As Double 'Jumlah adjustment yang telah dibuat
Dim Frm103_LM_BERAT_ASAL As Double 'Berat asal barang
Dim Frm103_LM_BERAT As Double
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset

TM = Frm103.L5_Text 'Tarikh mula
TA = Frm103.L6_Text 'Tarikh akhir
DATA_FOUND = 0
Frm103_LM_ADJUSTMENT = 0 'Jumlah adjustment yang telah dibuat

'###Padam Table### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "TRUNCATE TABLE 52_senarai_invoice_jualan"

Set rs = cn.Execute(strsql)
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "TRUNCATE TABLE 53_senarai_modal_jualan"

Set rs = cn.Execute(strsql)
Set rs = Nothing
'###Padam Table### - End
    
'#### Carian barang yang dijual di dalam tempoh report #### - Start (Jualan terus & kepada agen)
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 23_senarai_jualan where tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND status_rekod = 1 order by no_resit ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    If Not IsNull(rs!jenis_jualan) Then
    
        Frm103_LM_MODAL_TANPA_GST = 0
        Frm103_LM_MODAL_GST = 0
        Frm103_LM_MODAL_DENGAN_GST = 0
        Frm103_LM_JUALAN_TANPA_GST = 0
        Frm103_LM_JUALAN_GST = 0
        Frm103_LM_JUALAN_DENGAN_GST = 0
        Frm103_LM_KOMISYEN = 0
        Frm103_LM_ADJUSTMENT = 0
        Frm103_LM_JENIS = 0 '0 : Barang kemas , 1 : Barang permata
        Frm103_LM_BERAT = 0
        Frm103_LM_BERAT_ASAL = 0
        
        If Not IsNull(rs!berat_jualan) Then
            Frm103_LM_BERAT = rs!berat_jualan
        End If
        If rs!gst_barang_atau_upah = 0 Then

            If Not IsNull(rs!harga_tanpa_gst) Then
                If IsNumeric(rs!harga_tanpa_gst) Then Frm103_LM_JUALAN_TANPA_GST = rs!harga_tanpa_gst 'Harga jualan tanpa GST
            End If

            If Not IsNull(rs!jumlah_gst) Then
                If IsNumeric(rs!jumlah_gst) Then Frm103_LM_JUALAN_GST = rs!jumlah_gst 'GST Jualan
            End If
            
            If Not IsNull(rs!harga_dengan_gst) Then
                If IsNumeric(rs!harga_dengan_gst) Then Frm103_LM_JUALAN_DENGAN_GST = rs!harga_dengan_gst 'Harga jualan dengan GST
            End If
            
        ElseIf rs!gst_barang_atau_upah = 1 Then
        
            If Not IsNull(rs!jumlah_gst) Then
                If IsNumeric(rs!jumlah_gst) Then Frm103_LM_JUALAN_GST = rs!jumlah_gst 'GST Jualan
            End If
            
            If Not IsNull(rs!harga_jualan) Then
                If IsNumeric(rs!harga_jualan_dengan_gst) Then Frm103_LM_JUALAN_DENGAN_GST = rs!harga_jualan_dengan_gst 'Harga jualan dengan GST
            End If
        
            Frm103_LM_JUALAN_TANPA_GST = Frm103_LM_JUALAN_DENGAN_GST - Frm103_LM_JUALAN_GST 'Harga jualan tanpa GST
    
        End If
        
        If Not IsNull(rs!harga_dengan_gst) Then
            If IsNumeric(rs!jumlah_komisyen) Then Frm103_LM_KOMISYEN = rs!jumlah_komisyen 'Komisyen kepada agen dropship
        End If
    
'#### Data belian bagi barang dari no. siri produk ini ####- Start

        Set rs1 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs1.Open "select * from data_database where no_siri_Produk='" & rs!no_siri_Produk & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs1.EOF Then
            If Not IsNull(rs1!receiving_Status) Then
                If rs1!receiving_Status = 0 Or rs1!receiving_Status = 2 Or rs1!receiving_Status = 4 Or rs1!receiving_Status = 5 Then
                    Frm103_LM_JENIS = 0 '0 : Barang kemas , 1 : Barang permata
                Else
                    Frm103_LM_JENIS = 1 '0 : Barang kemas , 1 : Barang permata
                End If
            End If
            If Not IsNull(rs1!Berat) Then
                Frm103_LM_BERAT_ASAL = rs1!Berat
            End If
            If Not IsNull(rs1!gst_barang_atau_upah) Then
                If rs1!gst_barang_atau_upah = 0 Then
                    If Not IsNull(rs1!harga_tanpa_gst) Then
                        If IsNumeric(rs1!harga_tanpa_gst) Then Frm103_LM_MODAL_TANPA_GST = rs1!harga_tanpa_gst 'Harga modal tanpa GST
                    End If
                ElseIf rs1!gst_barang_atau_upah = 1 Then
                    If Not IsNull(rs1!kos_item_tanpa_tax) Then
                        If IsNumeric(rs1!kos_item_tanpa_tax) Then Frm103_LM_MODAL_TANPA_GST = rs1!kos_item_tanpa_tax 'Harga modal tanpa GST
                    End If
                End If
            Else
                If Not IsNull(rs1!harga_tanpa_gst) Then
                    If IsNumeric(rs1!harga_tanpa_gst) Then Frm103_LM_MODAL_TANPA_GST = rs1!harga_tanpa_gst 'Harga modal tanpa GST
                End If
            End If
        
            'If Not IsNull(rs1!harga_tanpa_gst) Then
            '    If IsNumeric(rs1!harga_tanpa_gst) Then Frm103_LM_MODAL_TANPA_GST = rs1!harga_tanpa_gst 'Harga modal tanpa GST
            'End If
            
            If Not IsNull(rs1!jumlah_gst) Then
                If IsNumeric(rs1!jumlah_gst) Then Frm103_LM_MODAL_GST = rs1!jumlah_gst 'GST modal
            End If
            
            If Not IsNull(rs1!harga_item) Then
                If IsNumeric(rs1!harga_item) Then Frm103_LM_MODAL_DENGAN_GST = rs1!harga_item 'Harga modal dengan GST
            End If
        
        End If
        
        rs1.Close
        Set rs1 = Nothing
'#### Data belian bagi barang dari no. siri produk ini ####- End
    
        Set rs2 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs2.Open "select * from 53_senarai_modal_jualan", cn, adOpenKeyset, adLockOptimistic
        
        rs2.AddNew
        If Not IsNull(rs!tarikh) Then rs2!tarikh = rs!tarikh 'Tarikh
        If Not IsNull(rs!no_resit) Then rs2!no_resit = rs!no_resit 'No. Invoice
        If Not IsNull(rs!no_siri_Produk) Then rs2!no_siri_Produk = rs!no_siri_Produk 'No. siri produk
        If Frm103_LM_JENIS = 0 Then
            If Frm103_LM_BERAT <> 0 Then
                If Frm103_LM_BERAT_ASAL = Frm103_LM_BERAT Then
                    rs2!modal_harga_tanpa_gst = Frm103_LM_MODAL_TANPA_GST 'Harga modal belian barang ini tidak termasuk GST
                    rs2!modal_jumlah_gst = Frm103_LM_MODAL_GST 'Jumlah GST modal barang ini
                    rs2!harga_item = Frm103_LM_MODAL_DENGAN_GST 'Jumlah modal belian barang ini dengan GST
                    rs2!untung = Format(Frm103_LM_JUALAN_TANPA_GST - Frm103_LM_MODAL_TANPA_GST, "0.00") 'Keuntungan tanpa GST (RM)
                Else
                    If Frm103_LM_BERAT_ASAL <> 0 Then
                        rs2!modal_harga_tanpa_gst = Format(Frm103_LM_BERAT * (Frm103_LM_MODAL_TANPA_GST / Frm103_LM_BERAT_ASAL), "0.00") 'Harga modal belian barang ini tidak termasuk GST
                        rs2!modal_jumlah_gst = Format(Frm103_LM_BERAT * (Frm103_LM_MODAL_GST / Frm103_LM_BERAT_ASAL), "0.00") 'Jumlah GST modal barang ini
                        rs2!harga_item = Format(Frm103_LM_BERAT * (Frm103_LM_MODAL_DENGAN_GST / Frm103_LM_BERAT_ASAL), "0.00") 'Jumlah modal belian barang ini dengan GST
                        rs2!untung = Format(Frm103_LM_JUALAN_TANPA_GST - (Frm103_LM_BERAT * (Frm103_LM_MODAL_DENGAN_GST / Frm103_LM_BERAT_ASAL)), "0.00") 'Keuntungan tanpa GST (RM)
                    Else
                        rs2!modal_harga_tanpa_gst = "0.00"
                        rs2!modal_jumlah_gst = "0.00"
                        rs2!harga_item = "0.00"
                        rs2!untung = "0.00"
                    End If
                End If
            Else
                rs2!modal_harga_tanpa_gst = Format(0, "0.00") 'Harga modal belian barang ini tidak termasuk GST
                rs2!modal_jumlah_gst = Format(0, "0.00") 'Jumlah GST modal barang ini
                rs2!harga_item = Format(0, "0.00") 'Jumlah modal belian barang ini dengan GST
                rs2!untung = Format(0, "0.00") 'Keuntungan tanpa GST (RM)
            End If
        Else
            rs2!modal_harga_tanpa_gst = Frm103_LM_MODAL_TANPA_GST 'Harga modal belian barang ini tidak termasuk GST
            rs2!modal_jumlah_gst = Frm103_LM_MODAL_GST 'Jumlah GST modal barang ini
            rs2!harga_item = Frm103_LM_MODAL_DENGAN_GST 'Jumlah modal belian barang ini dengan GST
            rs2!untung = Format(Frm103_LM_JUALAN_TANPA_GST - Frm103_LM_MODAL_TANPA_GST, "0.00") 'Keuntungan tanpa GST (RM)
        End If
        rs2!jualan_harga_tanpa_gst = Frm103_LM_JUALAN_TANPA_GST 'Harga jualan barang ini tidak termasuk GST
        rs2!jualan_jumlah_gst = Frm103_LM_JUALAN_GST 'Jumlah GST jualan barang ini
        rs2!harga_dengan_gst = Frm103_LM_JUALAN_DENGAN_GST 'Harga jualan dengan GST
        rs2!komisyen_staff = Frm103_LM_KOMISYEN 'Komisyen staff
        rs2!adjustment = Format(Frm103_LM_ADJUSTMENT, "0.00") 'Adjustment
        rs2.Update
        
        rs2.Close
        Set rs2 = Nothing
    
    End If
    rs.MoveNext
    
    DATA_FOUND = 1
Wend

rs.Close
Set rs = Nothing
'#### Carian barang yang dijual di dalam tempoh report #### - End (Jualan terus & kepada agen)

'#### Carian barang yang dijual di dalam tempoh report #### - Start (Tempahan)
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 42_tempahan_siap where tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND status_invoice = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Frm103_LM_MODAL_TANPA_GST = 0
    Frm103_LM_MODAL_GST = 0
    Frm103_LM_MODAL_DENGAN_GST = 0
    Frm103_LM_JUALAN_TANPA_GST = 0
    Frm103_LM_JUALAN_GST = 0
    Frm103_LM_JUALAN_DENGAN_GST = 0
    Frm103_LM_ADJUSTMENT = 0 'Jumlah adjustment yang telah dibuat
    Frm103_LM_KOMISYEN = 0
    
    'Set rs1 = New ADODB.Recordset
    'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    'rs1.Open "select SUM(jumlah_deposit_dengan_gst) from 41_akaun_tempahan where no_rujukan_tempahan='" & rs!no_rujukan_tempahan & "'", cn, adOpenKeyset, adLockOptimistic
    
    'If Not IsNull(rs1(0)) Then
    '    Frm103_LM_JUALAN_DENGAN_GST = rs1(0)
    'End If
    
    'rs1.Close
    'Set rs1 = Nothing
    
    'Set rs1 = New ADODB.Recordset
    'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    'rs1.Open "select SUM(jumlah_gst) from 41_akaun_tempahan where no_rujukan_tempahan='" & rs!no_rujukan_tempahan & "'", cn, adOpenKeyset, adLockOptimistic
    
    'If Not IsNull(rs1(0)) Then
    '    Frm103_LM_JUALAN_GST = rs1(0)
    'End If
    
    'rs1.Close
    'Set rs1 = Nothing
    
    'Set rs1 = New ADODB.Recordset
    'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    'rs1.Open "select SUM(adjustment_bayaran) from 41_akaun_tempahan where no_rujukan_tempahan='" & rs!no_rujukan_tempahan & "'", cn, adOpenKeyset, adLockOptimistic
    
    'If Not IsNull(rs1(0)) Then
    '    Frm103_LM_ADJUSTMENT = rs1(0)
    'End If
    
    'rs1.Close
    'Set rs1 = Nothing
    
    If Not IsNull(rs!harga_dengan_gst) Then Frm103_LM_JUALAN_DENGAN_GST = rs!harga_dengan_gst
    If Not IsNull(rs!jumlah_gst) Then Frm103_LM_JUALAN_GST = rs!jumlah_gst
    If Not IsNull(rs!baki_adjustment) Then Frm103_LM_ADJUSTMENT = rs!baki_adjustment
    
    Frm103_LM_JUALAN_TANPA_GST = Frm103_LM_JUALAN_DENGAN_GST - Frm103_LM_JUALAN_GST

    
'#### Data belian bagi barang dari no. siri produk ini ####- Start
    If Not IsNull(rs!no_siri_Produk) Then
    
        Set rs1 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs1.Open "select * from data_database where no_siri_Produk='" & rs!no_siri_Produk & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs1.EOF Then
        
            If Not IsNull(rs1!gst_barang_atau_upah) Then
                If rs1!gst_barang_atau_upah = 0 Then
                    If Not IsNull(rs1!harga_tanpa_gst) Then
                        If IsNumeric(rs1!harga_tanpa_gst) Then Frm103_LM_MODAL_TANPA_GST = rs1!harga_tanpa_gst 'Harga modal tanpa GST
                    End If
                ElseIf rs1!gst_barang_atau_upah = 1 Then
                    If Not IsNull(rs1!kos_item_tanpa_tax) Then
                        If IsNumeric(rs1!kos_item_tanpa_tax) Then Frm103_LM_MODAL_TANPA_GST = rs1!kos_item_tanpa_tax 'Harga modal tanpa GST
                    End If
                End If
            Else
                If Not IsNull(rs1!harga_tanpa_gst) Then
                    If IsNumeric(rs1!harga_tanpa_gst) Then Frm103_LM_MODAL_TANPA_GST = rs1!harga_tanpa_gst 'Harga modal tanpa GST
                End If
            End If
            
            If Not IsNull(rs1!jumlah_gst) Then
                If IsNumeric(rs1!jumlah_gst) Then Frm103_LM_MODAL_GST = rs1!jumlah_gst 'GST modal
            End If
            
            If Not IsNull(rs1!harga_item) Then
                If IsNumeric(rs1!harga_item) Then Frm103_LM_MODAL_DENGAN_GST = rs1!harga_item 'Harga modal dengan GST
            End If
        
        End If
        
        rs1.Close
        Set rs1 = Nothing
        
    End If
'#### Data belian bagi barang dari no. siri produk ini ####- End

    Set rs2 = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs2.Open "select * from 53_senarai_modal_jualan", cn, adOpenKeyset, adLockOptimistic
    
    rs2.AddNew
    If Not IsNull(rs!tarikh) Then rs2!tarikh = rs!tarikh 'Tarikh
    If Not IsNull(rs!no_resit_tempahan) Then rs2!no_resit = rs!no_resit_tempahan 'No. Invoice
    If Not IsNull(rs!no_siri_Produk) Then rs2!no_siri_Produk = rs!no_siri_Produk 'No. siri produk
    rs2!modal_harga_tanpa_gst = Frm103_LM_MODAL_TANPA_GST 'Harga modal belian barang ini tidak termasuk GST
    rs2!modal_jumlah_gst = Frm103_LM_MODAL_GST 'Jumlah GST modal barang ini
    rs2!harga_item = Frm103_LM_MODAL_DENGAN_GST 'Jumlah modal belian barang ini dengan GST
    rs2!jualan_harga_tanpa_gst = Frm103_LM_JUALAN_TANPA_GST 'Harga jualan barang ini tidak termasuk GST
    rs2!jualan_jumlah_gst = Frm103_LM_JUALAN_GST 'Jumlah GST jualan barang ini
    rs2!harga_dengan_gst = Frm103_LM_JUALAN_DENGAN_GST 'Harga jualan dengan GST
    rs2!komisyen_staff = Frm103_LM_KOMISYEN 'Komisyen staff
    rs2!untung = Format(Frm103_LM_JUALAN_TANPA_GST - Frm103_LM_MODAL_TANPA_GST - Frm103_LM_ADJUSTMENT, "0.00") 'Keuntungan tanpa GST (RM)
    rs2!adjustment = Format(Frm103_LM_ADJUSTMENT, "0.00") 'Adjustment
    rs2.Update
    
    rs2.Close
    Set rs2 = Nothing
    
    DATA_FOUND = 1
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'#### Carian barang yang dijual di dalam tempoh report #### - End (Tempahan)

'#### Carian barang yang dijual di dalam tempoh report #### - Start (Ansuran)
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 27_senarai_ansuran where tarikh_jelas BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh_jelas ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Frm103_LM_MODAL_TANPA_GST = 0
    Frm103_LM_MODAL_GST = 0
    Frm103_LM_MODAL_DENGAN_GST = 0
    Frm103_LM_JUALAN_TANPA_GST = 0
    Frm103_LM_JUALAN_GST = 0
    Frm103_LM_JUALAN_DENGAN_GST = 0
    Frm103_LM_ADJUSTMENT = 0 'Jumlah adjustment yang telah dibuat
    Frm103_LM_KOMISYEN = 0
    
    Set rs1 = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs1.Open "select SUM(jumlah_keseluruhan) from 28_rekod_ansuran where id_database_reg='" & rs!ID & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not IsNull(rs1(0)) Then
        Frm103_LM_JUALAN_DENGAN_GST = rs1(0)
    End If
    
    rs1.Close
    Set rs1 = Nothing
    
    Set rs1 = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs1.Open "select SUM(jumlah_gst) from 28_rekod_ansuran where id_database_reg='" & rs!ID & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not IsNull(rs1(0)) Then
        Frm103_LM_JUALAN_GST = rs1(0)
    End If
    
    rs1.Close
    Set rs1 = Nothing
    
    Set rs1 = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs1.Open "select SUM(adjustment) from 28_rekod_ansuran where id_database_reg='" & rs!ID & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not IsNull(rs1(0)) Then
        Frm103_LM_ADJUSTMENT = rs1(0)
    End If
    
    rs1.Close
    Set rs1 = Nothing
    
    Frm103_LM_JUALAN_TANPA_GST = Frm103_LM_JUALAN_DENGAN_GST - Frm103_LM_JUALAN_GST

    
'#### Data belian bagi barang dari no. siri produk ini ####- Start
    If Not IsNull(rs!no_siri_Produk) Then
    
        Set rs1 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs1.Open "select * from data_database where no_siri_Produk='" & rs!no_siri_Produk & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs1.EOF Then
        
            If Not IsNull(rs1!gst_barang_atau_upah) Then
                If rs1!gst_barang_atau_upah = 0 Then
                    If Not IsNull(rs1!harga_tanpa_gst) Then
                        If IsNumeric(rs1!harga_tanpa_gst) Then Frm103_LM_MODAL_TANPA_GST = rs1!harga_tanpa_gst 'Harga modal tanpa GST
                    End If
                ElseIf rs1!gst_barang_atau_upah = 1 Then
                    If Not IsNull(rs1!kos_item_tanpa_tax) Then
                        If IsNumeric(rs1!kos_item_tanpa_tax) Then Frm103_LM_MODAL_TANPA_GST = rs1!kos_item_tanpa_tax 'Harga modal tanpa GST
                    End If
                End If
            Else
                If Not IsNull(rs1!harga_tanpa_gst) Then
                    If IsNumeric(rs1!harga_tanpa_gst) Then Frm103_LM_MODAL_TANPA_GST = rs1!harga_tanpa_gst 'Harga modal tanpa GST
                End If
            End If
            
            If Not IsNull(rs1!jumlah_gst) Then
                If IsNumeric(rs1!jumlah_gst) Then Frm103_LM_MODAL_GST = rs1!jumlah_gst 'GST modal
            End If
            
            If Not IsNull(rs1!harga_item) Then
                If IsNumeric(rs1!harga_item) Then Frm103_LM_MODAL_DENGAN_GST = rs1!harga_item 'Harga modal dengan GST
            End If
        
        End If
        
        rs1.Close
        Set rs1 = Nothing
        
    End If
'#### Data belian bagi barang dari no. siri produk ini ####- End

    Set rs2 = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs2.Open "select * from 53_senarai_modal_jualan", cn, adOpenKeyset, adLockOptimistic
    
    rs2.AddNew
    If Not IsNull(rs!tarikh_jelas) Then rs2!tarikh = rs!tarikh_jelas 'Tarikh
    'If Not IsNull(rs!no_resit_tempahan) Then rs2!no_resit = rs!no_resit_tempahan 'No. Invoice
    If Not IsNull(rs!no_siri_Produk) Then rs2!no_siri_Produk = rs!no_siri_Produk 'No. siri produk
    rs2!modal_harga_tanpa_gst = Frm103_LM_MODAL_TANPA_GST 'Harga modal belian barang ini tidak termasuk GST
    rs2!modal_jumlah_gst = Frm103_LM_MODAL_GST 'Jumlah GST modal barang ini
    rs2!harga_item = Frm103_LM_MODAL_DENGAN_GST 'Jumlah modal belian barang ini dengan GST
    rs2!jualan_harga_tanpa_gst = Frm103_LM_JUALAN_TANPA_GST 'Harga jualan barang ini tidak termasuk GST
    rs2!jualan_jumlah_gst = Frm103_LM_JUALAN_GST 'Jumlah GST jualan barang ini
    rs2!harga_dengan_gst = Frm103_LM_JUALAN_DENGAN_GST 'Harga jualan dengan GST
    rs2!komisyen_staff = Frm103_LM_KOMISYEN 'Komisyen staff
    rs2!untung = Format(Frm103_LM_JUALAN_TANPA_GST - Frm103_LM_MODAL_TANPA_GST - Frm103_LM_ADJUSTMENT, "0.00") 'Keuntungan tanpa GST (RM)
    rs2!adjustment = Format(Frm103_LM_ADJUSTMENT, "0.00") 'Adjustment
    rs2.Update
    
    rs2.Close
    Set rs2 = Nothing
    
    DATA_FOUND = 1
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'#### Carian barang yang dijual di dalam tempoh report #### - End (Ansuran)

        
If DATA_FOUND = 1 Then
    Call Frm103_kira_untung_rugi
    
    Frm103.Frame2.Visible = True
    Frm103.Frame1.Visible = False
Else
    MsgBox "Tiada data jualan dijumpai dari " & TM & " hingga " & TA, vbInformation, "Info"
End If
End Sub
Sub Frm103_kiraan_untung_rugi_ori()
'on error resume next
Dim TA As Date
Dim TM As Date
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset

TM = Frm103.L5_Text 'Tarikh mula
TA = Frm103.L6_Text 'Tarikh akhir
DATA_FOUND = 0

'###Padam Table### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "TRUNCATE TABLE 52_senarai_invoice_jualan"

Set rs = cn.Execute(strsql)
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "TRUNCATE TABLE 53_senarai_modal_jualan"

Set rs = cn.Execute(strsql)
Set rs = Nothing
'###Padam Table### - End

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!no_resit) Then
    
        Set rs1 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs1.Open "select * from 52_senarai_invoice_jualan", cn, adOpenKeyset, adLockOptimistic
    
        rs1.AddNew
        If Not IsNull(rs!tarikh) Then rs1!tarikh = rs!tarikh 'Tarikh invoice
        If Not IsNull(rs!no_resit) Then rs1!no_resit = rs!no_resit 'No. invoice
        If Not IsNull(rs!harga_barang) Then
            rs1!harga_barang = rs!harga_barang 'Jumlah harga jualan barang tanpa cukai gst
        Else
            rs1!harga_barang = "0.00" 'Jumlah harga jualan barang tanpa cukai gst
        End If
        If Not IsNull(rs!jumlah_cukai_gst) Then
            rs1!jumlah_cukai_gst = rs!jumlah_cukai_gst 'Jumlah cukai GST
        Else
            rs1!jumlah_cukai_gst = "0.00" 'Jumlah cukai GST
        End If
        If Not IsNull(rs!harga_barang_dengan_gst) Then
            rs1!harga_barang_dengan_gst = rs!harga_barang_dengan_gst 'Jumlah harga jualan dengan GST
        Else
            rs1!harga_barang_dengan_gst = "0.00" 'Jumlah harga jualan dengan GST
        End If
        If Not IsNull(rs!diskaun) Then
            rs1!diskaun = rs!diskaun 'Jumlah diskaun (%)
        Else
            rs1!diskaun = "0.00" 'Jumlah diskaun (%)
        End If
        If Not IsNull(rs!harga_lepas_diskaun) Then
            rs1!harga_lepas_diskaun = rs!harga_lepas_diskaun 'Harga setelah diskaun
        Else
            rs1!harga_lepas_diskaun = "0.00" 'Harga setelah diskaun
        End If
        If Not IsNull(rs!adjustment) Then
            rs1!adjustment = rs!adjustment 'Adjustment (RM)
        Else
            rs1!adjustment = "0.00" 'Adjustment (RM)
        End If
        If Not IsNull(rs!harga_jualan) Then
            rs1!harga_jualan = rs!harga_jualan 'Jumlah harga keseluruhan invoice (RM)
        Else
            rs1!harga_jualan = "0.00" 'Jumlah harga keseluruhan invoice (RM)
        End If
        rs1.Update
        
        rs1.Close
        Set rs1 = Nothing
    
        Set rs1 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs1.Open "select * from 23_senarai_jualan where no_resit='" & rs!no_resit & "' order by no_siri_produk ASC", cn, adOpenKeyset, adLockOptimistic
        
        While rs1.EOF = False
            Frm103_LM_MODAL_TANPA_GST = 0
            Frm103_LM_MODAL_GST = 0
            Frm103_LM_MODAL_DENGAN_GST = 0
            Frm103_LM_JUALAN_TANPA_GST = 0
            Frm103_LM_JUALAN_GST = 0
            Frm103_LM_JUALAN_DENGAN_GST = 0
            Frm103_LM_KOMISYEN = 0
            
            If Not IsNull(rs1!harga_tanpa_gst) Then
                If IsNumeric(rs1!harga_tanpa_gst) Then Frm103_LM_JUALAN_TANPA_GST = rs1!harga_tanpa_gst 'Harga jualan tanpa GST
            End If
            
            If Not IsNull(rs1!jumlah_gst) Then
                If IsNumeric(rs1!jumlah_gst) Then Frm103_LM_JUALAN_GST = rs1!jumlah_gst 'GST Jualan
            End If
            
            If Not IsNull(rs1!harga_dengan_gst) Then
                If IsNumeric(rs1!harga_dengan_gst) Then Frm103_LM_JUALAN_DENGAN_GST = rs1!harga_dengan_gst 'Harga jualan dengan GST
            End If
            
            If Not IsNull(rs1!harga_dengan_gst) Then
                If IsNumeric(rs1!komisyen_staff) Then Frm103_LM_KOMISYEN = rs1!komisyen_staff 'Komisyen staff
            End If
            
            Set rs2 = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs2.Open "select * from data_database where no_siri_Produk='" & rs1!no_siri_Produk & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs2.EOF Then
            
                If Not IsNull(rs2!harga_tanpa_gst) Then
                    If IsNumeric(rs2!harga_tanpa_gst) Then Frm103_LM_MODAL_TANPA_GST = rs2!harga_tanpa_gst 'Harga modal tanpa GST
                End If
                
                If Not IsNull(rs2!jumlah_gst) Then
                    If IsNumeric(rs2!jumlah_gst) Then Frm103_LM_MODAL_GST = rs2!jumlah_gst 'GST modal
                End If
                
                If Not IsNull(rs2!harga_item) Then
                    If IsNumeric(rs2!harga_item) Then Frm103_LM_MODAL_DENGAN_GST = rs2!harga_item 'Harga modal dengan GST
                End If
            
            End If
            
            rs2.Close
            Set rs2 = Nothing
            
            Set rs3 = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs3.Open "select * from 53_senarai_modal_jualan", cn, adOpenKeyset, adLockOptimistic
            
            rs3.AddNew
            If Not IsNull(rs!tarikh) Then rs3!tarikh = rs!tarikh 'Tarikh
            If Not IsNull(rs!no_resit) Then rs3!no_resit = rs!no_resit 'No. Invoice
            If Not IsNull(rs1!no_siri_Produk) Then rs3!no_siri_Produk = rs1!no_siri_Produk 'No. siri produk
            rs3!modal_harga_tanpa_gst = Frm103_LM_MODAL_TANPA_GST 'Harga modal belian barang ini tidak termasuk GST
            rs3!modal_jumlah_gst = Frm103_LM_MODAL_GST 'Jumlah GST modal barang ini
            rs3!harga_item = Frm103_LM_MODAL_DENGAN_GST 'Jumlah modal belian barang ini dengan GST
            rs3!jualan_harga_tanpa_gst = Frm103_LM_JUALAN_TANPA_GST 'Harga jualan barang ini tidak termasuk GST
            rs3!jualan_jumlah_gst = Frm103_LM_JUALAN_GST 'Jumlah GST jualan barang ini
            rs3!harga_dengan_gst = Frm103_LM_JUALAN_DENGAN_GST 'Harga jualan dengan GST
            rs3!komisyen_staff = Frm103_LM_KOMISYEN 'Komisyen staff
            rs3.Update
            
            rs3.Close
            Set rs3 = Nothing
            
            rs1.MoveNext
            
            DATA_FOUND = 1
        Wend
        
        rs1.Close
        Set rs1 = Nothing
        
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

If DATA_FOUND = 1 Then
    Call Frm103_kira_untung_rugi
    
    Frm103.Frame2.Visible = True
    Frm103.Frame1.Visible = False
Else
    MsgBox "Tiada data jualan dijumpai dari " & TM & " hingga " & TA, vbInformation, "Info"
End If
End Sub
Sub Frm103_senarai_modal_jualan_header()
'on error resume next
With Frm103.LV1
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear

    Frm103.LV1.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "Tarikh", 1500, 2
    .ColumnHeaders.Add 5, , "No. invoice", 1500
    .ColumnHeaders.Add 6, , "No. siri produk", 1500
    .ColumnHeaders.Add 7, , "Modal tanpa GST (RM)", 1800, 1
    .ColumnHeaders.Add 8, , "GST modal (RM)", 1400, 1
    .ColumnHeaders.Add 9, , "Modal dengan GST (RM)", 2100, 1
    .ColumnHeaders.Add 10, , "Harga jualan tanpa GST (RM)", 2400, 1
    .ColumnHeaders.Add 11, , "GST jualan (RM)", 1500, 1
    .ColumnHeaders.Add 12, , "Harga jualan dengan GST (RM)", 2400, 1
    .ColumnHeaders.Add 13, , "Untung 1 (RM)", 1300, 1
    .ColumnHeaders.Add 14, , "Untung 2 (RM)", 1300, 1
    .ColumnHeaders.Add 15, , "Cawangan", 3000
    .ColumnHeaders.Add 16, , "Dulang", 1000, 2
    
End With
End Sub
Sub Frm103_senarai_modal_jualan()
'on error resume next
Dim TA As Date
Dim TM As Date
Dim Frm103_LM_TOTAL_PAGE As Double

TM = Frm103.L5_Text 'Tarikh mula
TA = Frm103.L6_Text 'Tarikh akhir

x = 0
Y = 0

Frm103_PAGE_SIZE = 45
Frm103_LM_TOTAL_PAGE = 0

LM_START_ROW = Frm103.L13_Text 'Start row

If Frm103.L31_Text = "Semua cawangan" Then

    Frm103_SEARCH_1 = Null
    Frm103_SEARCH_1_LOGIC = "<>"
    
Else

    Frm103_SEARCH_1 = Frm103.L31_Text
    Frm103_SEARCH_1_LOGIC = "="
    
End If

If Frm103.L35_Text = "Semua dulang" Then

    Frm103_SEARCH_2 = Null
    Frm103_SEARCH_2_LOGIC = "<>"
    
Else

    Frm103_SEARCH_2 = Frm103.L35_Text
    Frm103_SEARCH_2_LOGIC = "="
    
End If


If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm103_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm103.L14_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm103_PAGE_SIZE
        End If
    End If
End If

Frm103_LM_PAGE_FOUND = 0

Frm103.L8_Text = "Senarai modal dan harga jualan bagi cawangan [" & Frm103.L31_Text & "] dan dulang [" & Frm103.L35_Text & "] dari " & TM & " hingga " & TA & "."

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 23_senarai_jualan where dulang " & Frm103_SEARCH_2_LOGIC & "'" & Frm103_SEARCH_2 & "' AND (cawangan " & Frm103_SEARCH_1_LOGIC & "'" & Frm103_SEARCH_1 & "' OR cawangan " & Frm103_SEARCH_1_LOGIC & "'" & Frm103_SEARCH_1 & "') AND status_rekod = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC , ID ASC LIMIT " & LM_START_ROW & "," & Frm103_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm103_LM_PAGE_FOUND = 0 Then
        If Frm103.L14_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm103.L15_Text = Frm103.L15_Text + 1
                Frm103_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm103.L15_Text) Then
                    If Frm103.L15_Text <> 1 Then
                        Frm103.L15_Text = Frm103.L15_Text - 1
                        Frm103_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm103.L15_Text - 1) * Frm103_PAGE_SIZE) + x

    With Frm103.LV1.ListItems.Add(, , rs!ID)
     
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!tarikh) Then 'Tarikh invoice
            .ListSubItems.Add , , rs!tarikh
        Else
            .ListSubItems.Add , , ""
        End If
        If Not IsNull(rs!no_resit) Then 'No. invoice
            .ListSubItems.Add , , rs!no_resit
        Else
            .ListSubItems.Add , , ""
        End If
        If Not IsNull(rs!no_siri_Produk) Then 'No. siri produk
            .ListSubItems.Add , , rs!no_siri_Produk
        Else
            .ListSubItems.Add , , ""
        End If
        If Not IsNull(rs!harga_modal_excl_gst) Then 'Harga modal belian barang ini tidak termasuk GST
            .ListSubItems.Add , , Format(rs!harga_modal_excl_gst, "#,##0.00")
        Else
            .ListSubItems.Add , , "0.00"
        End If
        If Not IsNull(rs!harga_modal_gst) Then 'Jumlah GST modal barang ini
            .ListSubItems.Add , , Format(rs!harga_modal_gst, "#,##0.00")
        Else
            .ListSubItems.Add , , "0.00"
        End If
        If Not IsNull(rs!harga_modal_incl_gst) Then 'Jumlah modal belian barang ini dengan GST
            .ListSubItems.Add , , Format(rs!harga_modal_incl_gst, "#,##0.00")
        Else
            .ListSubItems.Add , , "0.00"
        End If
        If Not IsNull(rs!harga_jual_excl_gst) Then 'Harga jualan barang ini tidak termasuk GST
            .ListSubItems.Add , , Format(rs!harga_jual_excl_gst, "#,##0.00")
        Else
            .ListSubItems.Add , , "0.00"
        End If
        If Not IsNull(rs!jumlah_gst) Then 'Jumlah GST jualan barang ini
            .ListSubItems.Add , , Format(rs!jumlah_gst, "#,##0.00")
        Else
            .ListSubItems.Add , , "0.00"
        End If
        If Not IsNull(rs!harga_jualan_dengan_gst) Then 'Harga jualan dengan GST
            .ListSubItems.Add , , Format(rs!harga_jualan_dengan_gst, "#,##0.00")
        Else
            .ListSubItems.Add , , "0.00"
        End If
        If Not IsNull(rs!untung) Then 'Untung tanpa GST
            .ListSubItems.Add , , Format(rs!untung, "#,##0.00")
        Else
            .ListSubItems.Add , , "0.00"
        End If
        If Not IsNull(rs!untung2) Then 'Untung dengan GST
            .ListSubItems.Add , , Format(rs!untung2, "#,##0.00")
        Else
            .ListSubItems.Add , , "0.00"
        End If
        If Not IsNull(rs!cawangan) Then 'Cawangan
            .ListSubItems.Add , , rs!cawangan
        Else
            .ListSubItems.Add , , ""
        End If
        If Not IsNull(rs!dulang) Then 'Dulang
            .ListSubItems.Add , , rs!dulang
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
rs.Open "select COUNT(ID) from 23_senarai_jualan where dulang " & Frm103_SEARCH_2_LOGIC & "'" & Frm103_SEARCH_2 & "' AND (cawangan " & Frm103_SEARCH_1_LOGIC & "'" & Frm103_SEARCH_1 & "' OR cawangan " & Frm103_SEARCH_1_LOGIC & "'" & Frm103_SEARCH_1 & "') AND status_rekod = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm103_LM_TOTAL_PAGE = Format(rs(0) / Frm103_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm103_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm103_LM_PAGE = Split(Frm103_LM_TOTAL_PAGE, ".")(0)
        Frm103_LM_PAGE_LEBIHAN = Split(Frm103_LM_TOTAL_PAGE, ".")(1)
        
        If Frm103_LM_PAGE_LEBIHAN <> "00" Then
            Frm103.L16_Text = Frm103_LM_PAGE + 1
        Else
            Frm103.L16_Text = Frm103_LM_PAGE
        End If
        
    Else
    
        Frm103.L16_Text = Frm103_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm103.L16_Text = 0
    End If
Else
    Frm103.L16_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm103.L16_Text = vbNullString Then
    Frm103.L16_Text = 0
End If
'### Jumlah Data ### - End

If x <> 0 Then
    Frm103.L13_Text = LM_START_ROW
    Frm103.L14_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm103.L14_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub Frm103_kira_untung_rugi()
'on error resume next
Dim TA As Date
Dim TM As Date
Dim Frm103_LM_JUALAN As Double
Dim Frm103_LM_BELIAN As Double
Dim Frm103_LM_KOMISYEN As Double
Dim Frm103_LM_UNTUNG_RUGI As Double
Dim Frm103_LM_ADJUSTMENT As Double 'Jumlah adjustment yang telah dibuat
Dim Frm103_LM_ADJ_TEMPAHAN As Double 'Jumlah adjustment yang telah dibuat dari jualan secara tempahan
Dim Frm103_LM_DISKAUN As Double
Dim Frm103_LM_KUPON As Double
Dim Frm103_LM_TEBUS_MATA As Double

TM = Frm103.L5_Text 'Tarikh mula
TA = Frm103.L6_Text 'Tarikh akhir

If Frm103.L31_Text = "Semua cawangan" Then

    Frm103_SEARCH_1 = Null
    Frm103_SEARCH_1_LOGIC = "<>"
    
Else

    Frm103_SEARCH_1 = Frm103.L31_Text
    Frm103_SEARCH_1_LOGIC = "="
    
End If
If Frm103.L35_Text = "Semua dulang" Then

    Frm103_SEARCH_2 = Null
    Frm103_SEARCH_2_LOGIC = "<>"
    
Else

    Frm103_SEARCH_2 = Frm103.L35_Text
    Frm103_SEARCH_2_LOGIC = "="
    
End If

Frm103_LM_JUALAN = 0
Frm103_LM_BELIAN = 0
Frm103_LM_KOMISYEN = 0
Frm103_LM_UNTUNG_RUGI = 0
Frm103_LM_ADJUSTMENT = 0
Frm103_LM_ADJ_TEMPAHAN = 0 'Jumlah adjustment yang telah dibuat dari jualan secara tempahan
Frm103_LM_DISKAUN = 0
Frm103_LM_KUPON = 0
Frm103_LM_TEBUS_MATA = 0

Frm103.L25_Text = "Dari " & TM & " hingga " & TA & "." 'Header report

Frm103.L17_Text = "0.00"
Frm103.L18_Text = "0.00"
Frm103.L19_Text = "0.00"
Frm103.L20_Text = "0.00"
Frm103.L21_Text = "0.00"
Frm103.L22_Text = "0.00"
Frm103.L23_Text = "0.00"
Frm103.L28_Text = "0.00"
Frm103.L29_Text = "0.00"
Frm103.L30_Text = "0.00"

Dim Frm103_LM_JUALAN_DGN_GST As Double

Frm103_LM_JUALAN_DGN_GST = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_jualan_dengan_gst) , SUM(jumlah_gst) , SUM(harga_jual_excl_gst) , SUM(harga_modal_incl_gst) , SUM(harga_modal_gst) , SUM(harga_modal_excl_gst) from 23_senarai_jualan where dulang " & Frm103_SEARCH_2_LOGIC & "'" & Frm103_SEARCH_2 & "' AND cawangan " & Frm103_SEARCH_1_LOGIC & "'" & Frm103_SEARCH_1 & "' AND status_rekod = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then
    Frm103.L17_Text = Format(rs(0), "#,##0.00") 'Jualan (Termasuk GST)
    Frm103_LM_JUALAN_DGN_GST = rs(0)
End If
If Not IsNull(rs(1)) Then Frm103.L18_Text = Format(rs(1), "#,##0.00") 'Jumlah GST jualan barang ini
If Not IsNull(rs(2)) Then
    Frm103.L19_Text = Format(rs(2), "#,##0.00") 'Jumlah bersih jualan (Tanpa GST)
    Frm103_LM_JUALAN = rs(2)
End If
If Not IsNull(rs(3)) Then Frm103.L20_Text = Format(rs(3), "#,##0.00") 'Jumlah modal termasuk GST
If Not IsNull(rs(4)) Then Frm103.L21_Text = Format(rs(4), "#,##0.00") 'Jumlah GST modal
If Not IsNull(rs(5)) Then Frm103.L22_Text = Format(rs(5), "#,##0.00") 'Modal tidak termasuk GST

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(adjustment) , SUM(harga_barang_dengan_gst-harga_lepas_diskaun) , SUM(kupon_diskaun) , SUM(redeem_point) from 22_jualan where (cawangan " & Frm103_SEARCH_1_LOGIC & "'" & Frm103_SEARCH_1 & "' OR cawangan " & Frm103_SEARCH_1_LOGIC & "'" & Frm103_SEARCH_1 & "') AND status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then
    Frm103.L23_Text = Format(rs(0), "#,##0.00")
    Frm103_LM_ADJUSTMENT = rs(0) 'Jumlah adjustment terkumpul
End If
If Not IsNull(rs(1)) Then
    Frm103.L28_Text = Format(rs(1), "#,##0.00") 'Jumlah diskaun (RM)
    Frm103_LM_DISKAUN = rs(1)
End If
If Not IsNull(rs(2)) Then
    Frm103.L29_Text = Format(rs(2), "#,##0.00") 'Jumlah kupon diskaun (RM)
    Frm103_LM_KUPON = rs(2)
End If
If Not IsNull(rs(3)) Then
    Frm103.L30_Text = Format(rs(3), "#,##0.00") 'Jumlah tebusan mata ganjaran (RM)
    Frm103_LM_TEBUS_MATA = rs(3)
End If

rs.Close
Set rs = Nothing

Frm103.L26_Text = Format(Frm103_LM_ADJUSTMENT + Frm103_LM_ADJ_TEMPAHAN, "#,##0.00") 'Jumlah adjustment terkumpul
Frm103.L27_Text = Format(Frm103_LM_JUALAN - Frm103_LM_ADJUSTMENT - Frm103_LM_ADJ_TEMPAHAN - Frm103_LM_DISKAUN - Frm103_LM_KUPON - Frm103_LM_TEBUS_MATA, "#,##0.00") 'Jualan bersih (Setelah tolak adjustment)
End Sub
Sub Frm103_senarai_modal_jual_excel()
'on error resume next
Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
Dim TA As Date
Dim TM As Date

Note = "Sistem akan mengambil masa untuk mengeluarkan report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila tunggu sehingga sistem siap keluarkan report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    
    x = 0
    
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 15 'Tarikh
        .Columns("C").ColumnWidth = 15 'No. invoice
        .Columns("D").ColumnWidth = 15 'No. siri produk
        .Columns("E").ColumnWidth = 15 'Modal tanpa GST (RM)
        .Columns("F").ColumnWidth = 15 'GST modal (RM)
        .Columns("G").ColumnWidth = 15 'Modal dengan GST (RM)
        .Columns("H").ColumnWidth = 15 'Harga jualan tanpa GST (RM)
        .Columns("I").ColumnWidth = 15 'GST jualan (RM)
        .Columns("J").ColumnWidth = 15 'Harga jualan dengan GST (RM)
        .Columns("K").ColumnWidth = 15 'Untung 1 (RM)
        .Columns("L").ColumnWidth = 15 'Untung 2 (RM)
        .Columns("M").ColumnWidth = 30 'Cawangan
        .Columns("N").ColumnWidth = 10 'Dulang
        
        If Frm103.L5_Text <> vbNullString Then TM = Frm103.L5_Text 'Tarikh mula
        If Frm103.L6_Text <> vbNullString Then TA = Frm103.L6_Text 'Tarikh akhir

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
        
        '#### Header Report ###
        .Cells(7, 1) = Frm103.L8_Text 'Report Header"

        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh"
        .Cells(8, 3) = "No. invoice"
        .Cells(8, 4) = "No. siri produk"
        .Cells(8, 5) = "Modal tanpa GST (RM)"
        .Cells(8, 6) = "GST modal (RM)"
        .Cells(8, 7) = "Modal dengan GST (RM)"
        .Cells(8, 8) = "Harga jualan tanpa GST (RM)"
        .Cells(8, 9) = "GST jualan (RM)"
        .Cells(8, 10) = "Harga jualan dengan GST (RM)"
        .Cells(8, 11) = "Untung 1 (RM)"
        .Cells(8, 12) = "Untung 2 (RM)"
        .Cells(8, 13) = "Cawangan"
        .Cells(8, 14) = "Dulang"
        
        For i = 1 To 14
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
        
        If Frm103.L31_Text = "Semua cawangan" Then
        
            Frm103_SEARCH_1 = Null
            Frm103_SEARCH_1_LOGIC = "<>"
            
        Else
        
            Frm103_SEARCH_1 = Frm103.L31_Text
            Frm103_SEARCH_1_LOGIC = "="
            
        End If
        
        If Frm103.L35_Text = "Semua dulang" Then
        
            Frm103_SEARCH_2 = Null
            Frm103_SEARCH_2_LOGIC = "<>"
            
        Else
        
            Frm103_SEARCH_2 = Frm103.L35_Text
            Frm103_SEARCH_2_LOGIC = "="
            
        End If
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 23_senarai_jualan where dulang " & Frm103_SEARCH_2_LOGIC & "'" & Frm103_SEARCH_2 & "' AND (cawangan " & Frm103_SEARCH_1_LOGIC & "'" & Frm103_SEARCH_1 & "' OR cawangan " & Frm103_SEARCH_1_LOGIC & "'" & Frm103_SEARCH_1 & "') AND status_rekod = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC , ID ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh) Then .Cells(8 + x, 2) = "'" & rs!tarikh 'Tarikh
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_resit) Then .Cells(8 + x, 3) = rs!no_resit 'No. invoice
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_siri_Produk) Then .Cells(8 + x, 4) = rs!no_siri_Produk 'No. siri produk
            
            .Cells(8 + x, 5).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_modal_excl_gst) Then 'Harga modal belian barang ini tidak termasuk GST
                .Cells(8 + x, 5) = Format(rs!harga_modal_excl_gst, "#,##0.00")
            Else
                .Cells(8 + x, 5) = "0.00"
            End If
            .Cells(8 + x, 5).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 6).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_modal_gst) Then 'Jumlah GST modal barang ini
                .Cells(8 + x, 6) = Format(rs!harga_modal_gst, "#,##0.00")
            Else
                .Cells(8 + x, 6) = "0.00"
            End If
            .Cells(8 + x, 6).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 7).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_modal_incl_gst) Then 'Jumlah modal belian barang ini dengan GST
                .Cells(8 + x, 7) = Format(rs!harga_modal_incl_gst, "#,##0.00")
            Else
                .Cells(8 + x, 7) = "0.00"
            End If
            .Cells(8 + x, 7).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 8).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_jual_excl_gst) Then 'Harga jualan barang ini tidak termasuk GST
                .Cells(8 + x, 8) = Format(rs!harga_jual_excl_gst, "#,##0.00")
            Else
                .Cells(8 + x, 8) = "0.00"
            End If
            .Cells(8 + x, 8).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 9).HorizontalAlignment = xlRight
            If Not IsNull(rs!jumlah_gst) Then 'Jumlah GST jualan barang ini
                .Cells(8 + x, 9) = Format(rs!jumlah_gst, "#,##0.00")
            Else
                .Cells(8 + x, 9) = "0.00"
            End If
            .Cells(8 + x, 9).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 10).HorizontalAlignment = xlRight
            If Not IsNull(rs!harga_jualan_dengan_gst) Then 'Harga jualan dengan GST
                .Cells(8 + x, 10) = Format(rs!harga_jualan_dengan_gst, "#,##0.00")
            Else
                .Cells(8 + x, 10) = "0.00"
            End If
            .Cells(8 + x, 10).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 11).HorizontalAlignment = xlRight
            If Not IsNull(rs!untung) Then 'Untung
                .Cells(8 + x, 11) = Format(rs!untung, "#,##0.00")
            Else
                .Cells(8 + x, 11) = "0.00"
            End If
            .Cells(8 + x, 11).NumberFormat = "#,##0.00"
            
            .Cells(8 + x, 12).HorizontalAlignment = xlRight
            If Not IsNull(rs!untung2) Then 'Untung 2
                .Cells(8 + x, 12) = Format(rs!untung2, "#,##0.00")
            Else
                .Cells(8 + x, 12) = "0.00"
            End If
            .Cells(8 + x, 12).NumberFormat = "#,##0.00"

            If Not IsNull(rs!cawangan) Then .Cells(8 + x, 13) = rs!cawangan 'Cawangan
            
            If Not IsNull(rs!dulang) Then .Cells(8 + x, 14) = rs!dulang 'Dulang
            
            For Col = 1 To 14
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        Y = x + 2
        .Cells(8 + Y, 1) = "Harga Jualan Termasuk GST : RM " & Frm103.L17_Text
        
        Y = Y + 1
        .Cells(8 + Y, 1) = "GST Jualan : RM " & Frm103.L18_Text

        Y = Y + 1
        .Cells(8 + Y, 1) = "Harga Jualan Tanpa GST : RM " & Frm103.L19_Text

        Y = Y + 1
        .Cells(8 + Y, 1) = "Adjustment : RM " & Frm103.L26_Text
        
        Y = Y + 1
        .Cells(8 + Y, 1) = "Diskaun : RM " & Frm103.L28_Text
        
        Y = Y + 1
        .Cells(8 + Y, 1) = "Kupon Diskaun : RM " & Frm103.L29_Text
        
        Y = Y + 1
        .Cells(8 + Y, 1) = "Tebus Mata Ganjaran : RM " & Frm103.L30_Text
        
        Y = Y + 1
        .Cells(8 + Y, 1) = "Jualan Bersih (Tanpa GST) : RM " & Frm103.L27_Text
        
        Y = Y + 1
        .Cells(8 + Y, 1) = "Modal Termasuk GST : RM " & Frm103.L20_Text
        
        Y = Y + 1
        .Cells(8 + Y, 1) = "GST Modal : RM " & Frm103.L21_Text
    
        Y = Y + 1
        .Cells(8 + Y, 1) = "Modal Tanpa GST : RM " & Frm103.L22_Text
        
        Y = Y + 1
        .Cells(8 + Y, 1) = "Untung Bersih 1 : RM " & Frm103.L24_Text

        Y = Y + 1
        .Cells(8 + Y, 1) = "Untung Bersih 2 : RM " & Frm103.L34_Text
        
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
End Sub
Sub Frm103_cetak_penyata_untung_rugi()
'on error resume next
TM = Frm103.L5_Text 'Tarikh mula
TA = Frm103.L6_Text 'Tarikh akhir

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

PRINTER_FOUND = 1 '0 : Printer Not Found , 1 : Printer Found

'### Reset maklumat kedai ### - Start
Report60.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report60.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report60.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report60.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report60.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

If Frm103.L31_Text = "Semua cawangan" Then

    Frm103_SEARCH_1 = Null
    Frm103_SEARCH_1_LOGIC = "<>"
    
Else

    Frm103_SEARCH_1 = Frm103.L31_Text
    Frm103_SEARCH_1_LOGIC = "="
    
End If
If Frm103.L35_Text = "Semua dulang" Then

    Frm103_SEARCH_2 = Null
    Frm103_SEARCH_2_LOGIC = "<>"
    
Else

    Frm103_SEARCH_2 = Frm103.L35_Text
    Frm103_SEARCH_2_LOGIC = "="
    
End If

If MDI_frm1.L4_Text = "HQ" Then
    
    G_KEDAI = "HQ"
    
Else

    G_KEDAI = MDI_frm1.L20_Text
    
End If

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report60.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report60.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report60.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report60.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report60.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Report60.Sections("Section2").Controls("L1").Caption = Format(0, "#,##0.00") 'Jualan (Termasuk GST)
Report60.Sections("Section2").Controls("L2").Caption = Format(0, "#,##0.00") 'GST Jualan
Report60.Sections("Section2").Controls("L3").Caption = Format(0, "#,##0.00") 'Jualan Bersih (Belum tolak adjustment)
Report60.Sections("Section2").Controls("L4").Caption = Format(0, "#,##0.00") 'Kos Modal (Termasuk GST)
Report60.Sections("Section2").Controls("L5").Caption = Format(0, "#,##0.00") 'GST Modal
Report60.Sections("Section2").Controls("L6").Caption = Format(0, "#,##0.00") 'Kos Bersih (Modal)
Report60.Sections("Section2").Controls("L7").Caption = Format(0, "#,##0.00") 'Komisyen Staff
Report60.Sections("Section2").Controls("L8").Caption = Format(0, "#,##0.00") 'Untung Bersih
Report60.Sections("Section2").Controls("L9").Caption = vbNullString 'Header
Report60.Sections("Section2").Controls("L10").Caption = Format(0, "#,##0.00") 'Adjustment
Report60.Sections("Section2").Controls("L11").Caption = Format(0, "#,##0.00") 'Jualan Bersih
Report60.Sections("Section2").Controls("L12").Caption = Format(0, "#,##0.00") 'Jualan diskaun
Report60.Sections("Section2").Controls("L13").Caption = Format(0, "#,##0.00") 'Jualan kupon diskaun
Report60.Sections("Section2").Controls("L14").Caption = Format(0, "#,##0.00") 'Jualan mata ganjaran
Report60.Sections("Section2").Controls("L15").Caption = Format(0, "#,##0.00") 'Jualan mata ganjaran

Report60.Sections("Section2").Controls("L1").Caption = Frm103.L17_Text 'Jualan (Termasuk GST)
Report60.Sections("Section2").Controls("L2").Caption = Frm103.L18_Text 'GST Jualan
Report60.Sections("Section2").Controls("L3").Caption = Frm103.L19_Text 'Jualan Bersih (Belum tolak adjustment)
Report60.Sections("Section2").Controls("L4").Caption = Frm103.L20_Text 'Kos Modal (Termasuk GST)
Report60.Sections("Section2").Controls("L5").Caption = Frm103.L21_Text 'GST Modal
Report60.Sections("Section2").Controls("L6").Caption = Frm103.L22_Text 'Kos Bersih (Modal)
Report60.Sections("Section2").Controls("L7").Caption = Frm103.L23_Text 'Komisyen Staff
Report60.Sections("Section2").Controls("L8").Caption = Frm103.L24_Text 'Untung Bersih
Report60.Sections("Section2").Controls("L10").Caption = Frm103.L26_Text 'Adjustment
Report60.Sections("Section2").Controls("L11").Caption = Frm103.L27_Text 'Jualan Bersih
Report60.Sections("Section2").Controls("L12").Caption = Frm103.L28_Text 'Jualan diskaun
Report60.Sections("Section2").Controls("L13").Caption = Frm103.L29_Text 'Jualan kupon diskaun
Report60.Sections("Section2").Controls("L14").Caption = Frm103.L30_Text 'Jualan mata ganjaran
Report60.Sections("Section2").Controls("L15").Caption = Frm103.L34_Text

Report60.Sections("Section2").Controls("L9").Caption = Frm103.L25_Text 'Header

'### Paparan Penyata ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 23_senarai_jualan where dulang " & Frm103_SEARCH_2_LOGIC & "'" & Frm103_SEARCH_2 & "' AND (cawangan " & Frm103_SEARCH_1_LOGIC & "'" & Frm103_SEARCH_1 & "' OR cawangan " & Frm103_SEARCH_1_LOGIC & "'" & Frm103_SEARCH_1 & "') AND status_rekod = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC , ID ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report60.DataSource = rs
    Report60.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing
'### Paparan Penyata ### - End
End Sub
Sub Frm103_cetak_penyata_invoice()
'on error resume next
TM = Frm103.L5_Text 'Tarikh mula
TA = Frm103.L6_Text 'Tarikh akhir

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

Report58.Sections("Section5").Controls("L2").Caption = Format(0, "#,##0.00") 'Jualan Bersih (Belum tolak adjustment)
Report58.Sections("Section5").Controls("L3").Caption = Format(0, "#,##0.00") 'GST Jualan
Report58.Sections("Section5").Controls("L4").Caption = Format(0, "#,##0.00") 'Jualan (Termasuk GST)

Report58.Sections("Section4").Controls("L1").Caption = "Senarai pengeluaran invoice (jualan) dari " & TM & " hingga " & TA & "." 'Header
Report58.Sections("Section5").Controls("L2").Caption = Frm103.L19_Text 'Jualan Bersih (Belum tolak adjustment)
Report58.Sections("Section5").Controls("L3").Caption = Frm103.L18_Text 'GST Jualan
Report58.Sections("Section5").Controls("L4").Caption = Frm103.L17_Text 'Jualan (Termasuk GST)

'### Paparan Penyata ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 52_senarai_invoice_jualan", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report58.DataSource = rs
    Report58.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing
'### Paparan Penyata ### - End
End Sub
Sub Frm103_cetak_penyata_modal_jual()
'on error resume next
TM = Frm103.L5_Text 'Tarikh mula
TA = Frm103.L6_Text 'Tarikh akhir

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

'### Reset maklumat kedai ### - Start
Report59.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report59.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report59.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report59.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report59.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

If Frm103.L31_Text = "Semua cawangan" Then

    Frm103_SEARCH_1 = Null
    Frm103_SEARCH_1_LOGIC = "<>"
    
Else

    Frm103_SEARCH_1 = Frm103.L31_Text
    Frm103_SEARCH_1_LOGIC = "="
    
End If
If Frm103.L35_Text = "Semua dulang" Then

    Frm103_SEARCH_2 = Null
    Frm103_SEARCH_2_LOGIC = "<>"
    
Else

    Frm103_SEARCH_2 = Frm103.L35_Text
    Frm103_SEARCH_2_LOGIC = "="
    
End If

If MDI_frm1.L4_Text = "HQ" Then
    
    G_KEDAI = "HQ"
    
Else

    G_KEDAI = MDI_frm1.L20_Text
    
End If

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report59.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report59.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report59.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report59.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report59.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Report59.Sections("Section5").Controls("L2").Caption = "RM " & Format(0, "#,##0.00") 'Jualan Bersih (Belum tolak adjustment)
Report59.Sections("Section5").Controls("L3").Caption = "RM " & Format(0, "#,##0.00") 'GST Jualan
Report59.Sections("Section5").Controls("L4").Caption = "RM " & Format(0, "#,##0.00") 'Jualan (Termasuk GST)
Report59.Sections("Section5").Controls("L5").Caption = "RM " & Format(0, "#,##0.00") 'Kos Bersih (Modal)
Report59.Sections("Section5").Controls("L6").Caption = "RM " & Format(0, "#,##0.00") 'GST Modal
Report59.Sections("Section5").Controls("L7").Caption = "RM " & Format(0, "#,##0.00") 'Kos Modal (Termasuk GST)
Report59.Sections("Section5").Controls("L8").Caption = "RM " & Format(0, "#,##0.00") 'Keuntungan

Report59.Sections("Section5").Controls("L9").Caption = "RM " & Format(0, "#,##0.00")
Report59.Sections("Section5").Controls("L10").Caption = "RM " & Format(0, "#,##0.00")
Report59.Sections("Section5").Controls("L11").Caption = "RM " & Format(0, "#,##0.00")
Report59.Sections("Section5").Controls("L12").Caption = "RM " & Format(0, "#,##0.00")
Report59.Sections("Section5").Controls("L13").Caption = "RM " & Format(0, "#,##0.00")
Report59.Sections("Section5").Controls("L14").Caption = "RM " & Format(0, "#,##0.00")

Report59.Sections("Section4").Controls("L1").Caption = "Senarai modal dan harga jualan dari " & TM & " hingga " & TA & "." 'Header
Report59.Sections("Section5").Controls("L2").Caption = "RM " & Frm103.L19_Text 'Jualan Bersih (Belum tolak adjustment)
Report59.Sections("Section5").Controls("L3").Caption = "RM " & Frm103.L18_Text 'GST Jualan
Report59.Sections("Section5").Controls("L4").Caption = "RM " & Frm103.L17_Text 'Jualan (Termasuk GST)
Report59.Sections("Section5").Controls("L5").Caption = "RM " & Frm103.L22_Text 'Jualan Bersih (Belum tolak adjustment)
Report59.Sections("Section5").Controls("L6").Caption = "RM " & Frm103.L21_Text 'GST Jualan
Report59.Sections("Section5").Controls("L7").Caption = "RM " & Frm103.L20_Text 'Jualan (Termasuk GST)
Report59.Sections("Section5").Controls("L8").Caption = "RM " & Frm103.L23_Text 'Keuntungan

Report59.Sections("Section5").Controls("L9").Caption = "RM " & Frm103.L24_Text
Report59.Sections("Section5").Controls("L10").Caption = "RM " & Frm103.L34_Text
Report59.Sections("Section5").Controls("L11").Caption = "RM " & Frm103.L26_Text
Report59.Sections("Section5").Controls("L12").Caption = "RM " & Frm103.L28_Text
Report59.Sections("Section5").Controls("L13").Caption = "RM " & Frm103.L29_Text
Report59.Sections("Section5").Controls("L14").Caption = "RM " & Frm103.L30_Text

'### Paparan Penyata ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 23_senarai_jualan where dulang " & Frm103_SEARCH_2_LOGIC & "'" & Frm103_SEARCH_2 & "' AND (cawangan " & Frm103_SEARCH_1_LOGIC & "'" & Frm103_SEARCH_1 & "' OR cawangan " & Frm103_SEARCH_1_LOGIC & "'" & Frm103_SEARCH_1 & "') AND status_rekod = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC , ID ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report59.DataSource = rs
    Report59.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing
'### Paparan Penyata ### - End
End Sub
Sub Frm103_untung_rugi_summary()
'on error resume next
Dim Frm103_LM_JUALAN As Double 'Jumlah jualan bersih (tanpa GST)
Dim Frm103_LM_MODAL As Double 'Jumlah modal bersih (tanpa GST)
Dim Frm103_LM_ADJ As Double 'Jumlah Adjustment
Dim Frm103_LM_KOMISYEN As Double 'Jumlah komisyen
Dim Frm103_LM_UNTUNG_RUGI As Double 'Analisis untung rugi
Dim Frm103_LM_DISKAUN As Double 'Jumlah diskaun
Dim Frm103_LM_KUPON As Double 'Jumlah kupon
Dim Frm103_LM_TEBUS_MATA As Double 'Jumlah tebus mata
Dim Frm103_LM_UNTUNG_RUGI_GST As Double
Dim Frm103_LM_JUALAN_DGN_GST As Double
Dim Frm103_LM_MODAL_DGN_GST As Double

Frm103_LM_JUALAN = 0 'Jumlah jualan bersih (tanpa GST)
Frm103_LM_MODAL = 0 'Jumlah modal bersih (tanpa GST)
Frm103_LM_KOMISYEN = 0 'Jumlah komisyen
Frm103_LM_UNTUNG_RUGI = 0 'Analisis untung rugi
Frm103_LM_ADJ = 0 'Jumlah Adjustment
Frm103_LM_DISKAUN = 0 'Jumlah diskaun
Frm103_LM_KUPON = 0 'Jumlah kupon
Frm103_LM_TEBUS_MATA = 0 'Jumlah tebus mata
Frm103_LM_UNTUNG_RUGI_GST = 0
Frm103_LM_JUALAN_DGN_GST = 0
Frm103_LM_MODAL_DGN_GST = 0

If IsNumeric(Frm103.L19_Text) Then
    Frm103_LM_JUALAN = Frm103.L19_Text 'Jumlah jualan bersih (tanpa GST)
End If

If IsNumeric(Frm103.L17_Text) Then
    Frm103_LM_JUALAN_DGN_GST = Frm103.L17_Text 'Jumlah jualan bersih (dengan GST)
End If

If IsNumeric(Frm103.L20_Text) Then
    Frm103_LM_MODAL_DGN_GST = Frm103.L20_Text 'Jumlah modal bersih (dengan GST)
End If

If IsNumeric(Frm103.L22_Text) Then
    Frm103_LM_MODAL = Frm103.L22_Text 'Jumlah modal bersih (tanpa GST)
End If

If IsNumeric(Frm103.L23_Text) Then
    Frm103_LM_ADJ = Frm103.L23_Text 'Jumlah Adjustment
End If

If IsNumeric(Frm103.L26_Text) Then
    Frm103_LM_KOMISYEN = Frm103.L26_Text 'Jumlah komisyen
End If

If IsNumeric(Frm103.L28_Text) Then
    Frm103_LM_DISKAUN = Frm103.L28_Text 'Jumlah diskaun
End If

If IsNumeric(Frm103.L29_Text) Then
    Frm103_LM_KUPON = Frm103.L29_Text 'Jumlah kupon
End If

If IsNumeric(Frm103.L30_Text) Then
    Frm103_LM_TEBUS_MATA = Frm103.L30_Text 'Jumlah tebus mata
End If

'### Pengiraan untung bersih 1 ### - Start
Frm103_LM_UNTUNG_RUGI = Frm103_LM_JUALAN - Frm103_LM_MODAL - Frm103_LM_KOMISYEN - Frm103_LM_ADJ - Frm103_LM_DISKAUN - Frm103_LM_KUPON - Frm103_LM_TEBUS_MATA
Frm103_LM_UNTUNG_RUGI_GST = Frm103_LM_JUALAN_DGN_GST - Frm103_LM_MODAL_DGN_GST - Frm103_LM_KOMISYEN - Frm103_LM_ADJ - Frm103_LM_DISKAUN - Frm103_LM_KUPON - Frm103_LM_TEBUS_MATA

If Frm103_LM_UNTUNG_RUGI >= 0 Then
    Frm103.L24_Text = Format(Frm103_LM_UNTUNG_RUGI, "#,##0.00") 'Untung Bersih 1
Else
    Frm103.L24_Text = Format(Frm103_LM_UNTUNG_RUGI, "#,##0.00") 'Untung Bersih 1
End If
'### Pengiraan untung bersih 1 ### - Start

'### Pengiraan untung bersih 2 ### - Start
Frm103_LM_UNTUNG_RUGI = Frm103_LM_JUALAN - Frm103_LM_MODAL - Frm103_LM_KOMISYEN - Frm103_LM_ADJ - Frm103_LM_DISKAUN - Frm103_LM_KUPON - Frm103_LM_TEBUS_MATA
Frm103_LM_UNTUNG_RUGI_GST = Frm103_LM_JUALAN_DGN_GST - Frm103_LM_MODAL_DGN_GST - Frm103_LM_KOMISYEN - Frm103_LM_ADJ - Frm103_LM_DISKAUN - Frm103_LM_KUPON - Frm103_LM_TEBUS_MATA

If Frm103_LM_UNTUNG_RUGI_GST >= 0 Then
    Frm103.L34_Text = Format(Frm103_LM_UNTUNG_RUGI_GST, "#,##0.00") 'Untung Bersih 2
Else
    Frm103.L34_Text = Format(Frm103_LM_UNTUNG_RUGI_GST, "#,##0.00") 'Untung Bersih 2
End If
'### Pengiraan untung bersih 2 ### - Start
End Sub

