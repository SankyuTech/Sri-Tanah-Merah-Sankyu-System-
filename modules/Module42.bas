Attribute VB_Name = "Module42"
Sub Frm84_Call_Product_Detail()
'on error resume next
Dim Frm84_LM_BERAT As Double
Dim Frm84_LM_HARGA As Double
Dim Frm84_LM_GST As Double
Dim Frm84_LM_KENAIKAN_UPAH As Double
Dim Frm84_LM_UPAH_GRAM_ASAL As Double

Frm84_LM_DATA_FOUND = 0
Frm84_LM_BERAT = 1
Frm84_LM_READY_TO_SAVE = 0 'Flag : Ready To Save
Frm84_LM_UpdateList = 0
Frm84_LM_KOD_PURITY = vbNullString
Frm84_LM_PERMATA = 0
Frm84_LM_HARGA = 0
Frm84_LM_GST = 0
LM_KIRAAN_UPAH = 0
Frm84_LM_KENAIKAN_UPAH = 0
Frm84_LM_UPAH_GRAM_ASAL = 0

Frm84_LM_No_SIRI = UCase(Frm84.TB1) 'No. Siri Produk
Frm84.TB1 = vbNullString

Frm84.TB22 = "0.00"
LM_BARANG_KEMAS = 0

'###Periksa Samada Data Ini Telah Dimasukkan Ke Dalam Temp Table### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from " & G_JUALAN_TEMP & " where no_siri_produk='" & Frm84_LM_No_SIRI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Frm84.L41_Text = "0" Then 'Data Baru (Kemasukkan Baru)
        If rs!Status = "1" Or rs!Status = "4" Then
            MsgBox "Item Ini Telah Dimasukkan Ke Dalam Senarai Sebelum Ini.", vbInformation, "Info"
            
            Frm84.TB1 = vbNullString
            Frm84.TB1.SetFocus
        ElseIf rs!Status = 0 Then
            rs!Status = 1 '0 : Padam Data , 1 : Aktif (Tiada Perubahan) , 3 : Data Baru , 4 : Data Diedit , 5 : Keluarkan Data Dari Database Asal , 6 : Ignore Kemasukkan Data Ke Dalam Database
            rs.Update
            
            Frm84_LM_UpdateList = 1
        End If
    ElseIf Frm84.L41_Text = "1" Then 'Edit Data Lama + Kemasukkan Baru
        If rs!Status = "1" Or rs!Status = "4" Or rs!Status = "3" Then
            MsgBox "Item Ini Telah Dimasukkan Ke Dalam Senarai Sebelum Ini.", vbInformation, "Info"
            
            Frm84.TB1 = vbNullString
            Frm84.TB1.SetFocus
        ElseIf rs!Status = "5" Or rs!Status = "6" Then
            If rs!Status = "5" Then rs!Status = "4" '0 : Padam Data , 1 : Aktif (Tiada Perubahan) , 3 : Data Baru , 4 : Data Diedit , 5 : Keluarkan Data Dari Database Asal , 6 : Ignore Kemasukkan Data Ke Dalam Database
            If rs!Status = "6" Then rs!Status = "3" '0 : Padam Data , 1 : Aktif (Tiada Perubahan) , 3 : Data Baru , 4 : Data Diedit , 5 : Keluarkan Data Dari Database Asal , 6 : Ignore Kemasukkan Data Ke Dalam Database
            rs.Update
            
            Frm84.TB1 = vbNullString
            Frm84.TB1.SetFocus
            
            Frm84_LM_UpdateList = 1
        End If
    End If
    Frm84_LM_DATA_FOUND = 1
    If rs!Status = "0" Or rs!Status = "5" Then Frm84_LM_DATA_FOUND = 0
    
End If

rs.Close
Set rs = Nothing
'###Periksa Samada Data Ini Telah Dimasukkan Ke Dalam Temp Table### - End

LM_KIRAAN_UPAH = 1 '0 : Pengiraan upah mengikut berat barang , 1 : Pengiraan upah mengikut tetapan asal

'###Carian Data Basic Bagi Item Ini### - Start
If Frm84_LM_DATA_FOUND = 0 Then

    If G_UPAH_MODE = 1 Then
        LM_UPAH_MODE = 1
    Else
        LM_UPAH_MODE = 0
    End If
    If G_KIRAAN_UPAH = 0 Then
        LM_KIRAAN_UPAH = 0 '0 : Pengiraan upah mengikut berat barang , 1 : Pengiraan upah mengikut tetapan asal
    ElseIf G_KIRAAN_UPAH = 1 Then
        LM_KIRAAN_UPAH = 1 '0 : Pengiraan upah mengikut berat barang , 1 : Pengiraan upah mengikut tetapan asal
    End If

If LM_KIRAAN_UPAH = 0 Then
'### Carian pemalar kenaikan upah ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 73_tetapan_upah where default_setting='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        If Frm84.CB4 = 1 Then
            If Not IsNull(rs!pelanggan) Then
                
                If IsNumeric(rs!pelanggan) Then Frm84_LM_KENAIKAN_UPAH = rs!pelanggan 'Kenaikan upah : Pelanggan
                
            End If
        ElseIf Frm84.CB5 = 1 Then
            If Not IsNull(rs!Member) Then
            
                If IsNumeric(rs!Member) Then Frm84_LM_KENAIKAN_UPAH = rs!Member 'Kenaikan upah : Member
                
            End If
        ElseIf Frm84.CB9 = 1 Then
            If Not IsNull(rs!raf) Then
                
                If IsNumeric(rs!raf) Then Frm84_LM_KENAIKAN_UPAH = rs!raf 'Kenaikan upah : RAF
                
            End If
        ElseIf Frm84.CB6 = 1 Then
            If Not IsNull(rs!Pengedar) Then
        
                If IsNumeric(rs!Pengedar) Then Frm84_LM_KENAIKAN_UPAH = rs!Pengedar 'Kenaikan upah : Pengedar
                
            End If
        ElseIf Frm84.CB10 = 1 Then
            If Not IsNull(rs!normal_dealer) Then
        
                If IsNumeric(rs!normal_dealer) Then Frm84_LM_KENAIKAN_UPAH = rs!normal_dealer 'Kenaikan upah : Normal dealer
                
            End If
        'ElseIf Frm84.CB11 = 1 Then
        '    If Not IsNull(rs!master_dealer) Then
                
        '        If IsNumeric(rs!master_dealer) Then Frm84_LM_KENAIKAN_UPAH = rs!master_dealer 'Kenaikan upah : Master dealer
                
        '    End If
        End If

    End If
    
    rs.Close
    Set rs = Nothing
'### Carian pemalar kenaikan upah ### - End
    
End If

    LM_FLAG_GST_MODAL = 0 '0 : Ada cukai GST , 1 : Tiada cukai GST
    LM_FLAG_BARANG = 0 '0 : Barang yang belum pernah jual , 1 : Potong
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from Data_Database where no_siri_Produk='" & Frm84_LM_No_SIRI & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        If Not IsNull(rs!cawangan) Then
            
            If MDI_frm1.L20_Text <> rs!cawangan Then
                
                MsgBox "Stok ini adalah milik cawangan [" & rs!cawangan & "]. Anda tidak dibenarkan untuk jual barang ini.", vbExclamation, "Info"
                
                rs.Close
                Set rs = Nothing
                
                Exit Sub
                
            End If
        
        End If
    
        'If rs!StatusItem = "10" Or rs!StatusItem = "12" Or rs!StatusItem = "20" Or rs!StatusItem = "22" Or rs!StatusItem = "28" Then
        If rs!StatusItem = "10" Then
        
            If rs!StatusItem = "10" Then
                LM_FLAG_BARANG = 0 '0 : Barang yang belum pernah jual , 1 : Potong
            ElseIf rs!StatusItem = "12" Or rs!StatusItem = "20" Or rs!StatusItem = "22" Or rs!StatusItem = "28" Then
                LM_FLAG_BARANG = 1 '0 : Barang yang belum pernah jual , 1 : Potong
            End If
            
            LM_FOUND = 1
        
            If Not IsNull(rs!gst_ari_nashi) Then
            
                If rs!gst_ari_nashi = 0 Then
                    LM_FLAG_GST_MODAL = 1 '0 : Ada cukai GST , 1 : Tiada cukai GST
                End If
                
            End If
        
            If Not IsNull(rs!receiving_Status) Then
                If rs!receiving_Status = 0 Or rs!receiving_Status = 2 Or rs!receiving_Status = 4 Or rs!receiving_Status = 5 Then
                    
                    Frm84.TB7 = "0" 'Diskaun
                    Frm84.TB7.Locked = True
                    Frm84.TB7.BackColor = &H8000000A
                        
                    LM_BARANG_KEMAS = 1
                    
                    Frm84.Frame2.Visible = False
                    Frm84.Frame3.Visible = False
                    
                    Frm84.TB2 = Frm84_LM_No_SIRI 'No. Siri Produk
                    'Frm84.TB3 = Format(rs!beza_berat - rs!susut_berat, "0.00") 'Berat Asal (g)
                    Frm84.TB3 = Format(rs!beza_berat, "0.00") 'Berat Asal (g)
                    Frm84.TB4 = Format(rs!beza_berat, "0.00") 'Berat Jualan (g)
                    
                    If Not IsNull(rs!harga_Per_Gram_Item) Then
                        Frm84.L34_Text = Format(rs!harga_Per_Gram_Item, "0.00") 'Harga Per Gram Item (RM/g)
                    Else
                        Frm84.L34_Text = Format(0, "0.00") 'Harga Per Gram Item (RM/g)
                    End If
                    If Not IsNull(rs!harga_per_gram_tanpa_gst) Then
                        Frm84.L42_Text = Format(rs!harga_per_gram_tanpa_gst, "0.00") 'Harga Per Gram Item (RM/g)
                    Else
                        Frm84.L42_Text = Format(0, "0.00") 'Harga Per Gram Item (RM/g)
                    End If
                    
                    Frm84.L55_Text = rs!UPAH 'Upah modal asal
                    
                    If Not IsNull(rs!flag_upah) Then
                        
                        If rs!flag_upah = 0 Then
                        
                            LM_UPAH_TERIMA = 0 'Cara pengiraan upah semasa penerimaan stok , 0 : Pengiraan upah mengikut berat barang , 1 : Pengiraan upah mengikut tetapan asal
                            
                        ElseIf rs!flag_upah = 1 Then
                            
                            LM_UPAH_TERIMA = 1 'Cara pengiraan upah semasa penerimaan stok , 0 : Pengiraan upah mengikut berat barang , 1 : Pengiraan upah mengikut tetapan asal
                            
                        End If
                        
                    End If
                    
                    If Not IsNull(rs!upah_per_gram) Then Frm84_LM_UPAH_GRAM_ASAL = rs!upah_per_gram 'Upah per gram asal (dari supplier)
                    
                    Frm84.TB5.Locked = False
                    
                    If LM_FLAG_BARANG = 0 Then '0 : Barang yang belum pernah jual , 1 : Potong
                        Frm84.TB4.Locked = False
                        Frm84.TB4.BackColor = &HFFFFFF
                    Else
                        Frm84.TB4.Locked = True
                        Frm84.TB4.BackColor = &H8000000A
                    End If
                    
                    Frm84.TB15.Locked = False
                    Frm84.TB22.Locked = False
                    Frm84.TB6.Locked = True
                    
                    Frm84.L70_Text = 1 'Cara pengiraan upah jualan , 0 : Ikut berat , 1 : Ikut tetapan asal
                    
                    Frm84.L68_Text = "Modal (RM/g) :                   Jual (RM/g) :"

                    If Not IsNull(rs!harga_item) Then Frm84_LM_HARGA = rs!harga_item
                    If Not IsNull(rs!jumlah_gst) Then Frm84_LM_GST = rs!jumlah_gst
                    If Not IsNull(rs!Berat) Then Frm84_LM_BERAT = rs!Berat
                    
                    Frm84.L69_Text = Format((Frm84_LM_HARGA - Frm84_LM_GST) / Frm84_LM_BERAT, "#,##0.00")
                    
                    Frm84.TB5.BackColor = &HFFFFFF
                    
                    Frm84.TB15.BackColor = &HFFFFFF
                    Frm84.TB22.BackColor = &HFFFFFF
                    Frm84.TB6.BackColor = &H8000000A
                    
                    If Frm84.CB7 = 1 Then
                        Frm84.Frame2.Visible = True
                    ElseIf Frm84.CB7 = 0 Then
                        Frm84.Frame2.Visible = False
                    End If
                    Frm84_LM_READY_TO_SAVE = 1 'Flag : Ready To Save
                    
                Else
                
                    If G_DISC_ARI_NASHI = 1 Then
                    
                        Frm84.TB7 = Format(G_DISC_JUMLAH, "#,##0.00") 'Diskaun
                        Frm84.TB7.Locked = False
                        Frm84.TB7.BackColor = &HFFFFFF
                        
                    ElseIf G_DISC_ARI_NASHI = 0 Then
                    
                        Frm84.TB7 = "0" 'Diskaun
                        Frm84.TB7.Locked = True
                        Frm84.TB7.BackColor = &H8000000A
                    
                    End If
                    
                    Frm84.TB2 = Frm84_LM_No_SIRI 'No. Siri Produk
                    
                    'If Not IsNull(rs!beza_berat) Then Frm84.TB4 = Format(rs!beza_berat, "0.00") 'Berat Jualan (g)
                    
                    If Not IsNull(rs!harga_item) Then
                        Frm84.L34_Text = Format(rs!harga_item, "0.00") 'Harga Modal (RM)
                    Else
                        Frm84.L34_Text = Format(0, "0.00") 'Harga Modal (RM)
                    End If
                    If Not IsNull(rs!harga_tanpa_gst) Then
                        Frm84.L42_Text = Format(rs!harga_tanpa_gst, "0.00") 'Harga Modal (RM)
                    Else
                        Frm84.L42_Text = Format(0, "0.00") 'Harga Modal (RM)
                    End If
                    
                    If Not IsNull(rs!harga_item) Then Frm84_LM_HARGA = rs!harga_item
                    If Not IsNull(rs!jumlah_gst) Then Frm84_LM_GST = rs!jumlah_gst
                    
                    Frm84.L69_Text = Format(Frm84_LM_HARGA - Frm84_LM_GST, "#,##0.00")
                    
                    Frm84.Frame2.Visible = False
                    Frm84.Frame3.Visible = False
                    
                    Frm84.TB3 = vbNullString
                    Frm84.TB4 = vbNullString
                    Frm84.TB5 = vbNullString
                    Frm84.TB15 = vbNullString
                    Frm84.TB22 = vbNullString
                    
                    Frm84.L70_Text = 1 'Cara pengiraan upah jualan , 0 : Ikut berat , 1 : Ikut tetapan asal
                    
                    Frm84.TB5.Locked = True
                    Frm84.TB4.Locked = True
                    Frm84.TB15.Locked = True
                    Frm84.TB22.Locked = True
                    Frm84.TB6.Locked = False
                    
                    Frm84.L68_Text = "Modal (RM)   :                      Jual (RM) :"
                    
                    Frm84.TB5.BackColor = &H8000000A
                    Frm84.TB4.BackColor = &H8000000A
                    Frm84.TB15.BackColor = &H8000000A
                    Frm84.TB22.BackColor = &H8000000A
                    Frm84.TB6.BackColor = &HFFFFFF
                    If Frm84.CB7 = 1 Then
                        Frm84.Frame3.Visible = True
                    ElseIf Frm84.CB7 = 0 Then
                        Frm84.Frame3.Visible = False
                    End If
                    'If Not IsNull(rs!code_Supplier) Then Frm84_LM_HARGA_JUALAN = Format(rs!code_Supplier)
                    
                    If Frm84.CB4 = 1 Then
                        If Not IsNull(rs!code_Supplier) Then
                            Frm84_LM_HARGA_JUALAN = Format(rs!code_Supplier, "0.00")  'Harga Jualan Pelanggan
                            Frm84.L52_Text = Format(rs!code_Supplier, "0.00")  'Harga Jualan Pelanggan Asal
                        End If
                    ElseIf Frm84.CB5 = 1 Then
                        If Not IsNull(rs!HargaJualan_Member) Then
                            Frm84_LM_HARGA_JUALAN = Format(rs!HargaJualan_Member, "0.00") 'Harga Jualan Member
                            Frm84.L52_Text = Format(rs!HargaJualan_Member, "0.00") 'Harga Jualan Member Asal
                        End If
                    ElseIf Frm84.CB9 = 1 Then
                        If Not IsNull(rs!HargaJualan_RAF) Then
                            Frm84_LM_HARGA_JUALAN = Format(rs!HargaJualan_RAF, "0.00") 'Harga Jualan RAF
                            Frm84.L52_Text = Format(rs!HargaJualan_RAF, "0.00") 'Harga Jualan RAF Asal
                        End If
                    ElseIf Frm84.CB6 = 1 Then
                        If Not IsNull(rs!HargaJualan_Pengedar) Then
                            Frm84_LM_HARGA_JUALAN = Format(rs!HargaJualan_Pengedar, "0.00") 'Harga Jualan Pengedar
                            Frm84.L52_Text = Format(rs!HargaJualan_Pengedar, "0.00") 'Harga Jualan Pengedar Asal
                        End If
                    ElseIf Frm84.CB10 = 1 Then
                        If Not IsNull(rs!hargajualan_normal_dealer) Then
                            Frm84_LM_HARGA_JUALAN = Format(rs!hargajualan_normal_dealer, "0.00") 'Harga Jualan Normal Dealer
                            Frm84.L52_Text = Format(rs!hargajualan_normal_dealer, "0.00") 'Harga Jualan Normal Dealer Asal
                        End If
                    'ElseIf Frm84.CB11 = 1 Then
                    '    If Not IsNull(rs!hargajualan_master_dealer) Then
                    '        Frm84_LM_HARGA_JUALAN = Format(rs!hargajualan_master_dealer, "0.00") 'Harga Jualan Master Dealer
                    '        Frm84.L52_Text = Format(rs!hargajualan_master_dealer, "0.00") 'Harga Jualan Master Dealer Asal
                    '    End If
                    End If
                    
                    Frm84_LM_PERMATA = 1
                    
                End If
            End If
            
            If LM_BARANG_KEMAS = 1 Then
                If LM_UPAH_MODE = 1 Then
    
                    'If LM_UPAH_TERIMA = 0 And LM_KIRAAN_UPAH = 0 Then
                    If LM_KIRAAN_UPAH = 0 Then
                        
                        Frm84.TB22 = Format(Frm84_LM_KENAIKAN_UPAH + Frm84_LM_UPAH_GRAM_ASAL, "0.00") 'Upah per gram
                        
                        Frm84.L70_Text = 0 'Cara pengiraan upah jualan , 0 : Ikut berat , 1 : Ikut tetapan asal
                        
                        Frm84.TB22.Locked = False
                        Frm84.TB15.Locked = True
                        Frm84.TB22.BackColor = &HFFFFFF
                        Frm84.TB15.BackColor = &H8000000A
    
                    Else
                    
                        Frm84.TB15.Locked = False
                        Frm84.TB22.Locked = True
                        Frm84.TB15.BackColor = &HFFFFFF
                        Frm84.TB22.BackColor = &H8000000A
    
                        If Frm84.CB4 = 1 Then
                            If Not IsNull(rs!Upah_Jualan) Then
                                Frm84.TB15 = Format(rs!Upah_Jualan, "0.00") 'Upah Pelanggan
                                Frm84.L53_Text = Format(rs!Upah_Jualan, "0.00") 'Upah Pelanggan Asal
                            End If
                        ElseIf Frm84.CB5 = 1 Then
                            If Not IsNull(rs!Upah_Member) Then
                                Frm84.TB15 = Format(rs!Upah_Member, "0.00") 'Upah Member
                                Frm84.L53_Text = Format(rs!Upah_Member, "0.00") 'Upah Member Asal
                            End If
                        ElseIf Frm84.CB9 = 1 Then
                            If Not IsNull(rs!Upah_RAF) Then
                                Frm84.TB15 = Format(rs!Upah_RAF, "0.00") 'Upah RAF
                                Frm84.L53_Text = Format(rs!Upah_RAF, "0.00") 'Upah RAF Asal
                            End If
                        ElseIf Frm84.CB6 = 1 Then
                            If Not IsNull(rs!Upah_Pengedar) Then
                                Frm84.TB15 = Format(rs!Upah_Pengedar, "0.00") 'Upah Pengedar
                                Frm84.L53_Text = Format(rs!Upah_Pengedar, "0.00") 'Upah Pengedar Asal
                            End If
                        ElseIf Frm84.CB10 = 1 Then
                            If Not IsNull(rs!upah_normal_dealer) Then
                                Frm84.TB15 = Format(rs!upah_normal_dealer, "0.00") 'Upah Normal Dealer
                                Frm84.L53_Text = Format(rs!upah_normal_dealer, "0.00") 'Upah Normal Dealer Asal
                            End If
                        'ElseIf Frm84.CB11 = 1 Then
                        '    If Not IsNull(rs!upah_master_dealer) Then
                        '        Frm84.TB15 = Format(rs!upah_master_dealer, "0.00") 'Upah Master Dealer
                        '        Frm84.L53_Text = Format(rs!upah_master_dealer, "0.00") 'Upah Master Dealer Asal
                        '    End If
                        End If
                    
                    End If
                Else
                    Frm84.TB15 = Format(0, "0.00") 'Upah
                End If
            End If
            
            If Not IsNull(rs!kategori_Produk) Then Frm84.L12_Text = rs!kategori_Produk 'Kategori Produk
            If Not IsNull(rs!kod_Purity) Then
                Frm84_LM_KOD_PURITY = rs!kod_Purity 'Kod Purity
                Frm84.L13_Text = rs!kod_Purity 'Kod Purity
            End If
            
            If Frm84_LM_PERMATA = 1 Then
                Frm84.TB6 = Format(Frm84_LM_HARGA_JUALAN, "0.00")
            End If
        ElseIf rs!StatusItem = "11" Then
            MsgBox "Item Ini Telah Dijual. No. Siri Produk [" & Frm84_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm84.TB1 = vbNullString
            Frm84.TB1.SetFocus
        ElseIf rs!StatusItem = "12" Then
            MsgBox "Item Ini Telah Dijual Secara Potong. No. Siri Produk [" & Frm84_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm84.TB1 = vbNullString
            Frm84.TB1.SetFocus
        ElseIf rs!StatusItem = "13" Then
            MsgBox "Item Ini Telah Dijual Secara Potong. No. Siri Produk [" & Frm84_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm84.TB1 = vbNullString
            Frm84.TB1.SetFocus
        ElseIf rs!StatusItem = "14" Or rs!StatusItem = "21" Or rs!StatusItem = "22" Then
            MsgBox "Item Ini Telah Ditempah Oleh Pelanggan. No. Siri Produk [" & Frm84_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm84.TB1 = vbNullString
            Frm84.TB1.SetFocus
        ElseIf rs!StatusItem = "15" Or rs!StatusItem = "19" Or rs!StatusItem = "20" Then
            MsgBox "Item Ini Telah Dibeli Secara Ansuran. No. Siri Produk [" & Frm84_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm84.TB1 = vbNullString
            Frm84.TB1.SetFocus
        ElseIf rs!StatusItem = "16" Then
            MsgBox "Item Ini Telah Dihantar Ke Ar-Rahnu. No. Siri Produk [" & Frm84_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm84.TB1 = vbNullString
            Frm84.TB1.SetFocus
        ElseIf rs!StatusItem = "17" Then
            MsgBox "Item Ini Telah Dijual Secara ETA. No. Siri Produk [" & Frm84_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm84.TB1 = vbNullString
            Frm84.TB1.SetFocus
        ElseIf rs!StatusItem = "23" Then
            MsgBox "Item Ini Telah Dihantar Ke Supplier/Kilang. No. Siri Produk [" & Frm84_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm84.TB1 = vbNullString
            Frm84.TB1.SetFocus
        ElseIf rs!StatusItem = "24" Then
            MsgBox "Item Ini Telah Dihantar Ke Supplier/Kilang. No. Siri Produk [" & Frm84_LM_No_SIRI & "]", vbExclamation, "Info"
            
            Frm84.TB1 = vbNullString
            Frm84.TB1.SetFocus
        ElseIf rs!StatusItem = "25" Then
            MsgBox "Item Ini Telah Diagihkan Ke Cawangan. No. Siri Produk [" & Frm84_LM_No_SIRI & "]", vbExclamation, "Info"
    
            Frm84.TB1 = vbNullString
            Frm84.TB1.SetFocus
        ElseIf rs!StatusItem = "26" Then
            MsgBox "Item Ini Telah Dijual Oleh Cawangan. No. Siri Produk [" & Frm84_LM_No_SIRI & "]", vbExclamation, "Info"
    
            Frm84.TB1 = vbNullString
            Frm84.TB1.SetFocus
        ElseIf rs!StatusItem = "0" Then
            MsgBox "Item Ini Telah Dipadamkan Dari Sistem. No. Siri Produk [" & Frm84_LM_No_SIRI & "]", vbExclamation, "Info"
    
            Frm84.TB1 = vbNullString
            Frm84.TB1.SetFocus
        ElseIf rs!StatusItem = "27" Or rs!StatusItem = "28" Then
            MsgBox "Item Ini Telah Dijual Dari Menu GDN. No. Siri Produk [" & Frm84_LM_No_SIRI & "]", vbExclamation, "Info"
    
            Frm84.TB1 = vbNullString
            Frm84.TB1.SetFocus
        ElseIf rs!StatusItem = "29" Then
            MsgBox "Item Ini Telah Diubah Status Kepada Hilang , Dicuri Dan Sebagainya. No. Siri Produk [" & Frm84_LM_No_SIRI & "]", vbExclamation, "Info"
    
            Frm84.TB1 = vbNullString
            Frm84.TB1.SetFocus
        End If
    Else
        MsgBox "No. Siri Produk Ini [" & Frm84_LM_No_SIRI & "] Tidak Dijumpai.", vbExclamation, "Info"
        
        Frm84.TB1 = vbNullString
        Frm84.TB1.SetFocus
    End If
    
    rs.Close
    Set rs = Nothing
End If
'###Carian Data Basic Bagi Item Ini### - End

'###Periksa Data Produk### - Start
If Frm84_LM_READY_TO_SAVE = 1 Then 'Flag : Ready To Save
    If Frm84_LM_KOD_PURITY <> vbNullString Then

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from hargaemas where Purity='" & Frm84_LM_KOD_PURITY & "' AND cawangan='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!HargaDariSupplier) Then
                If IsNumeric(rs!HargaDariSupplier) Then
                    Frm84.L54_Text = rs!HargaDariSupplier
                Else
                    Frm84.L54_Text = 0
                End If
            Else
                Frm84.L54_Text = 0
            End If
            If Not IsNull(rs!harga_staff) Then
                If IsNumeric(rs!harga_staff) Then Frm84.L49_Text = Format(rs!harga_staff, "0.00") 'Harga jualan kepada staff
            End If
            If Frm84.CB4 = 1 Then
                If IsNumeric(rs!Harga_Pelanggan) Then Frm84.TB5 = Format(rs!Harga_Pelanggan, "0.00") 'Harga Emas Semasa Pelanggan
            ElseIf Frm84.CB5 = 1 Then
                If IsNumeric(rs!Harga_Member) Then Frm84.TB5 = Format(rs!Harga_Member, "0.00") 'Harga Emas Semasa Member
            ElseIf Frm84.CB6 = 1 Then
                If IsNumeric(rs!Harga_Pengedar) Then Frm84.TB5 = Format(rs!Harga_Pengedar, "0.00") 'Harga Emas Semasa Pengedar
            ElseIf Frm84.CB9 = 1 Then
                If IsNumeric(rs!Harga_RAF) Then Frm84.TB5 = Format(rs!Harga_RAF, "0.00") 'Harga Emas Semasa RAF
            ElseIf Frm84.CB10 = 1 Then
                If IsNumeric(rs!harga_nd) Then Frm84.TB5 = Format(rs!harga_nd, "0.00") 'Harga Emas Semasa Normal Dealer
            'ElseIf Frm84.CB11 = 1 Then
            '    If IsNumeric(rs!harga_md) Then Frm84.TB5 = Format(rs!harga_md, "0.00") 'Harga Emas Semasa Master Dealer
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        
        Frm84.TB9 = "0.00" 'Adjustment
    End If
    'If Frm84.CB1 = 1 Then Call Frm84_Auto_Update_List_Jualan
End If

Frm84.TB1 = vbNullString
'If Frm84.CB1 = 1 Then Call Frm84_auto_insert_data

If LM_FLAG_GST_MODAL = 1 Then '0 : Ada cukai GST , 1 : Tiada cukai GST

    If Frm84.L84_Text = "1" Then
      
        If Frm84.CB2 = 0 Then
        
            Note = "Tiada cukai GST dikenakan semasa pembelian barang ini dari supplier." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Adakah anda ingin menjual barang ini tanpa cukai GST?" & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila pilih [Yes] jika ingin menetapkan barang ini dijual tanpa cukai GST dan pilih [No] jika tiada perubahan yang diperlukan."
                    
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbYes Then
            
                Frm84.CB2 = 1
                
            End If
            
        End If
    
    End If
    
End If

If LM_FOUND = 1 Then
    Frm84.L83_Text = 0 '0 : Stok kedai , 1 : Barang trade in/potong
    Frm84.Pic8.Visible = False
End If

If G_AUTO_INSERT = "YES" Then

    Frm84.L89_Text = -1 'Titik Pencarian Data
    Frm84.L90_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm84.L87_Text = 0 'Paparan Page ke-xxx
    Frm84.L88_Text = 0
    
    GM_NEXT_PREV = 0
                
    Call tesutochu

End If

If Frm84_LM_UpdateList = 1 Then
    'Call Frm84_Senarai_Jualan_Header
    'Call Frm84_Senarai_Jualan
    Frm84.TB1.SetFocus
End If
'###Periksa Data Produk### - End
End Sub
Sub Frm84_Load_Form()
'on error resume next
GLOBAL_DISABLE = 0
Frm84.L4_Text = 0
Frm84.L5_Text = "0.00"
Frm84.L6_Text = "0.00"
Frm84.L7_Text = "0.00"
Frm84.L9_Text = "0.00"
Frm84.L10_Text = "0.00"
Frm84.L11_Text = "0.00"
Frm84.TB17 = "0.00"
Frm84.TB18 = vbNullString
Frm84.TB46 = vbNullString
Frm84.L16_Text = vbNullString
Frm84.L27_Text = vbNullString
Frm84.L28_Text = vbNullString
Frm84.L29_Text = vbNullString
Frm84.L46_Text = 0
Frm84.L14_Text = 0
Frm84.L15_Text = "0.00"
Frm84.TB41 = vbNullString
Frm84.CB19 = 0
Frm84.L84_Text = "0"

Frm84.L73_Text = "0.00" 'Nilai mata yang ditebus
Frm84.L74_Text = "0" 'Mata yang diperolehi

G_BIL_JUALAN = 0

For Z = 1 To 20
    G_PURITY_JUALAN(Z) = vbNullString
Next Z

Frm84.L75_Text = "0.00"
Frm84.TB35 = "0"
Frm84.L76_Text = "0"
Frm84.L77_Text = "0"
Frm84.TB36 = "0"
Frm84.TB37 = "0"
Frm84.L78_Text = "0.00"
Frm84.L79_Text = 0

Frm84.L17_Text = "0.00"
Frm84.L18_Text = "0.00"
Frm84.L19_Text = "0.00"
Frm84.L20_Text = "0.00"
Frm84.L21_Text = "0.00"
Frm84.L22_Text = "0.00"
Frm84.L23_Text = "0.00"
'Frm84.L26_Text = "0.00"
Frm84.L37_Text = "0.00"

'Frm84.L31_Text = "0.00"
'Frm84.L32_Text = "0.00"
'Frm84.L81_Text = "0.00"
'Frm84.L82_Text = "0.00"

Frm84.TB19 = "0.00"
Frm84.TB20 = "0.00"
'Frm84.TB21 = "0.00"
'Frm84.TB27 = "0.00"
'Frm84.TB28 = "0.00"
'Frm84.TB29 = "0.00"
'Frm84.TB32 = "0.00"
Frm84.TB33 = "0.00"
Frm84.TB42 = "0.00"
Frm84.TB45 = vbNullString

Frm84.L48_Text = "0.00"
Frm84.L49_Text = vbNullString
Frm84.L50_Text = "0.00"
Frm84.L51_Text = vbNullString
Frm84.L52_Text = vbNullString
Frm84.L53_Text = vbNullString

Frm84.L56_Text = 0 '0 : Tiada , 1 : Belian dengan trade in (sudah ada voucher) , 2 : Trade in
Frm84.L57_Text = vbNullString 'No. Voucher
Frm84.L58_Text = vbNullString 'Jumlah trade in
Frm84.L59_Text = 0 '0 : Barang baru , 1 : Edit
Frm84.L85_Text = "0" '0 : Barang baru , 1 : Edit
Frm84.L60_Text = vbNullString 'Memory : No. Voucher
Frm84.L61_Text = vbNullString 'Memory : No. rujukan belian

Frm84.CMD3.Visible = True
Frm84.CMD13.Visible = False
Frm84.CMD14.Visible = False

Frm84.L89_Text = -1 'Titik Pencarian Data
Frm84.L90_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Frm84.L87_Text = 0 'Paparan Page ke-xxx
Frm84.L88_Text = 0

GM_NEXT_PREV = 0
            
Frm84.Frame2.Visible = False
Frm84.Frame3.Visible = False
Frm84.Frame5.Visible = False
Frm84.Frame6.Visible = False

Frm84.TB49 = 0
Frm84.TB50 = G_TI_RATE_TI
Frm84.TB51 = G_TI_RATE_BB
Frm84.TB52 = G_TI_RATE_TUKAR
G_TI_MODE = 0

G_TI_MEMORY(0, 0) = 0
G_TI_MEMORY(1, 1) = 0
G_TI_MEMORY(1, 2) = 0
G_TI_MEMORY(1, 3) = 0

G_TI_MEMORY(2, 1) = 0
G_TI_MEMORY(2, 2) = 0
G_TI_MEMORY(2, 3) = 0

G_TI_MEMORY(3, 1) = 0
G_TI_MEMORY(3, 2) = 0
G_TI_MEMORY(3, 2) = 0

Call sys_config_membership
If G_MODE = "NO" Then
    Frm84.Frame7.Visible = False
Else
    Frm84.Frame7.Visible = True
End If

user_level = MDI_frm1.L4_Text

If user_level = "Administration" Or user_level = "Guest/User" Then

    Frm84.CB13.Visible = False
    Frm84.Label16.Visible = False
    
Else
    
    If G_GST_SYSTEM = "YES" Then
        Frm84.CB13.Visible = True
        Frm84.Label16.Visible = True
        
        If G_INVOICE_RASMI = 0 Then
            Frm84.CB13 = 1
        Else
            Frm84.CB13 = 0
        End If
        
    Else
        Frm84.CB13.Visible = False
        Frm84.Label16.Visible = False
    End If
End If

'GoTo skipa:
Call frm_kiraan_harga_selepas_ti
Call frm84_senarai_barang_purity

Frm84.CB14 = 0

GLOBAL_DISABLE = 1
Frm84.L8_Text = G_RATE_GST 'Jumlah Kadar GST
Frm84.L84_Text = G_FLAG_BIL_GST
Frm84.L38_Text = G_SPREAD_TI 'Potongan Harga Resit Trade in (%)
If G_GST_JUAL = 0 Then
    Frm84.CB2 = 1
    Frm84.CB3 = 0
    Frm84.CB18 = 0
ElseIf G_GST_JUAL = 1 Then
    Frm84.CB2 = 0
    Frm84.CB3 = 1
    Frm84.CB18 = 0
ElseIf G_GST_JUAL = 2 Then
    Frm84.CB2 = 0
    Frm84.CB3 = 0
    Frm84.CB18 = 1
End If
If G_GST_JUALAN_INC = 1 Then
    Frm84.CB18 = 1
ElseIf G_GST_JUALAN_INC = 0 Then
    Frm84.CB18 = 0
End If
If IsNumeric(G_J_DISC_UPAH) Then
    Frm84.L48_Text = G_J_DISC_UPAH 'Peratusan penurunan upah kepada staff
Else
    Frm84.L48_Text = 0
End If
If IsNumeric(G_J_DISC_PERMATA) Then
    Frm84.L50_Text = G_J_DISC_PERMATA 'Peratusan penurunan harga barang permata kepada staff
Else
    Frm84.L50_Text = 0
End If
If IsNumeric(G_KADAR_COMM_STAFF) Then 'Kadar komisyen upah kepada agen dropship (%)
    Frm84.TB43 = G_KADAR_COMM_STAFF
Else
    Frm84.TB43 = 0
End If
If G_DISC_ARI_NASHI = 1 Then
    Frm84.TB7 = Format(G_DISC_JUMLAH, "#,##0.00") 'Diskaun
    Frm84.TB7.Locked = False
    Frm84.TB7.BackColor = &HFFFFFF
Else
    Frm84.TB7 = "0" 'Diskaun
    Frm84.TB7.Locked = True
    Frm84.TB7.BackColor = &H8000000A
End If
If G_SCANNER_MODE = 1 Then
    Frm84.CB1 = 1
Else
    Frm84.CB1 = 0
End If
Frm84.L46_Text = G_LIMIT_INVOICE 'Jumlah Limit Jualan
Frm84.L80_Text = "RM " & Format(G_KUPON_DISC, "0.00") & " /g"
If Frm84.CB4 = 1 Or Frm84.CB5 = 1 Then

    Frm84.TB35 = G_PEMALAR_BONUS_BIASA
    Frm84.TB37 = G_PEMALAR_TEBUS_BIASA

ElseIf Frm84.CB6 = 1 Then

    Frm84.TB35 = G_PEMALAR_BONUS_SILVER
    Frm84.TB37 = G_PEMALAR_TEBUS_SILVER

ElseIf Frm84.CB9 = 1 Then

    Frm84.TB35 = G_PEMALAR_BONUS_GOLD
    Frm84.TB37 = G_PEMALAR_TEBUS_GOLD

ElseIf Frm84.CB10 = 1 Then

    Frm84.TB35 = G_PEMALAR_BONUS_PLATINUM
    Frm84.TB37 = G_PEMALAR_TEBUS_PLATINUM

End If
GLOBAL_DISABLE = 0

GoTo skip_a:

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Default1 = "Default" Then
        GLOBAL_DISABLE = 1
        If Not IsNull(rs!ResitNo) Then Frm84.L3_Text = rs!ResitNo 'No. invoice rasmi
        If Not IsNull(rs!no_rujukan_tak_rasmi) Then Frm84.L66_Text = rs!no_rujukan_tak_rasmi 'No. invoice tidak rasmi
        If Not IsNull(rs!gst_value) Then Frm84.L8_Text = rs!gst_value 'Jumlah Kadar GST
        If Not IsNull(rs!flag_bil_gst) Then Frm84.L84_Text = rs!flag_bil_gst
        'If Not IsNull(rs!cas_Kad_Kredit) Then Frm84.L31_Text = Format(rs!cas_Kad_Kredit, "0.00") 'Cas Kad Kredit
        'If Not IsNull(rs!cas_debit_kad) Then Frm84.L32_Text = Format(rs!cas_debit_kad, "0.00") 'Cas Debit Kredit
        If Not IsNull(rs!potongan_trade_in) Then Frm84.L38_Text = rs!potongan_trade_in 'Potongan Harga Resit Trade in (%)
        If Not IsNull(rs!gst_arinashi) Then
            If rs!gst_arinashi = 0 Then
                Frm84.CB2 = 1
                Frm84.CB3 = 0
                Frm84.CB18 = 0
            ElseIf rs!gst_arinashi = 1 Then
                Frm84.CB2 = 0
                Frm84.CB3 = 1
                Frm84.CB18 = 0
            ElseIf rs!gst_arinashi = 2 Then
                Frm84.CB2 = 0
                Frm84.CB3 = 0
                Frm84.CB18 = 1
            End If
        End If
        If Not IsNull(rs!gst_jualan_included) Then
            If rs!gst_jualan_included = 1 Then
                Frm84.CB18 = 1
            ElseIf rs!gst_jualan_included = 0 Then
                Frm84.CB18 = 0
            End If
        Else
            Frm84.CB18 = 0
        End If
        'Frm84.DTPicker1 = DateTime.Date
        
        If Not IsNull(rs!upah_staff) Then
            Frm84.L48_Text = rs!upah_staff 'Peratusan penurunan upah kepada staff
        Else
            Frm84.L48_Text = 0
        End If
        If Not IsNull(rs!diskaun_permata_staff) Then
            Frm84.L50_Text = rs!diskaun_permata_staff 'Peratusan penurunan harga barang permata kepada staff
        Else
            Frm84.L50_Text = 0
        End If
        
        If Not IsNull(rs!kadar_komisyen_upah) Then 'Kadar komisyen upah kepada agen dropship (%)
            Frm84.TB43 = rs!kadar_komisyen_upah
        Else
            Frm84.TB43 = 0
        End If
        
        If Not IsNull(rs!diskaun_ari_nashi) Then
            If rs!diskaun_ari_nashi = 1 Then
                Frm84.TB7 = Format(rs!diskaun, "0.00") 'Diskaun
                Frm84.TB7.Locked = False
                Frm84.TB7.BackColor = &HFFFFFF
            Else
                Frm84.TB7 = "0" 'Diskaun
                Frm84.TB7.Locked = True
                Frm84.TB7.BackColor = &H8000000A
            End If
        End If
        If rs!ScannerMode = 1 Then
            Frm84.CB1 = 1
        Else
            Frm84.CB1 = 0
        End If
        If Not IsNull(rs!invoice_type) Then
            Frm84.L46_Text = rs!invoice_type 'Jumlah Limit Jualan
        Else
            Frm84.L46_Text = 0
        End If
        If Not IsNull(rs!kupon_diskaun) Then
            If IsNumeric(rs!kupon_diskaun) Then
                Frm84.L80_Text = "RM " & Format(rs!kupon_diskaun, "0.00") & " /g"
            Else
                Frm84.L80_Text = "RM " & Format(0, "0.00") & " /g"
            End If
        Else
            Frm84.L80_Text = "RM " & Format(0, "0.00") & " /g"
        End If
        If Frm84.CB4 = 1 Or Frm84.CB5 = 1 Then
            If Not IsNull(rs!pemalar_bonus_biasa) Then
                Frm84.TB35 = rs!pemalar_bonus_biasa
            Else
                Frm84.TB35 = 0
            End If
            If Not IsNull(rs!pemalar_tebus_bonus_biasa) Then
                Frm84.TB37 = rs!pemalar_tebus_bonus_biasa
            Else
                Frm84.TB37 = 0
            End If
        ElseIf Frm84.CB6 = 1 Then
            If Not IsNull(rs!pemalar_bonus_silver) Then
                Frm84.TB35 = rs!pemalar_bonus_silver
            Else
                Frm84.TB35 = 0
            End If
            If Not IsNull(rs!pemalar_tebus_bonus_silver) Then
                Frm84.TB37 = rs!pemalar_tebus_bonus_silver
            Else
                Frm84.TB37 = 0
            End If
        ElseIf Frm84.CB9 = 1 Then
            If Not IsNull(rs!pemalar_bonus_gold) Then
                Frm84.TB35 = rs!pemalar_bonus_gold
            Else
                Frm84.TB35 = 0
            End If
            If Not IsNull(rs!pemalar_tebus_bonus_gold) Then
                Frm84.TB37 = rs!pemalar_tebus_bonus_gold
            Else
                Frm84.TB37 = 0
            End If
        ElseIf Frm84.CB10 = 1 Then
            If Not IsNull(rs!pemalar_bonus_platinum) Then
                Frm84.TB35 = rs!pemalar_bonus_platinum
            Else
                Frm84.TB35 = 0
            End If
            If Not IsNull(rs!pemalar_tebus_bonus_platinum) Then
                Frm84.TB37 = rs!pemalar_tebus_bonus_platinum
            Else
                Frm84.TB37 = 0
            End If
        End If
            
        GLOBAL_DISABLE = 0
    End If
End If

rs.Close
Set rs = Nothing

skip_a:

'###Senarai Nama Pekerja###
Frm84.CBB1.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where status='" & "Aktif" & "' AND InvestorSmall = 0 AND InvestorBig = 0", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    If rs!user_level = 1 Or rs!user_level = 2 Or rs!user_level = 3 Or rs!user_level = G_LEVEL_USER Then Frm84.CBB1.AddItem rs!Samaran & "  |  " & rs!NoPekerja
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'###Padam Temp Table###
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "TRUNCATE TABLE " & G_JUALAN_TEMP & ""

Set rs = cn.Execute(strsql)
Set rs = Nothing

Call Frm84_jurujual

'skipa:
End Sub
Sub Frm84_Reset()
'on error resume next
Frm84.Frame1.Left = 1600
Frm84.Frame1.Top = 150
Frm84.Frame4.Left = 1600
Frm84.Frame4.Top = 150
Frm84.Frame6.Left = 1600
Frm84.Frame6.Top = 150

Frm84.Pic3.Left = 1600
Frm84.Pic3.Top = 150
Frm84.Pic6.Left = 9720
Frm84.Pic6.Top = 3960

Frm84.Frame8.Left = 1600
Frm84.Frame8.Top = 150

Frm84.Pic8.Left = 240
Frm84.Pic8.Top = 1230

Frm84.Frame1.Visible = False
Frm84.Frame2.Visible = False
Frm84.Frame3.Visible = False
Frm84.Pic3.Visible = False
Frm84.Frame4.Visible = False
Frm84.Pic6.Visible = False
Frm84.Frame5.Visible = False
Frm84.Frame6.Visible = False
Frm84.Frame8.Visible = False

Frm84.TB1 = vbNullString
Frm84.TB2 = vbNullString
Frm84.L12_Text = vbNullString
Frm84.L13_Text = vbNullString
Frm84.L14_Text = 0
Frm84.L15_Text = "0.00"

Frm84.TB3 = "0.00"
Frm84.TB4 = "0.00"
Frm84.TB5 = "0.00"
Frm84.TB6 = "0.00"
'Frm84.TB7 = vbNullString
Frm84.TB8 = "0.00"
Frm84.TB9 = "0.00"
Frm84.TB10 = "0.00"
Frm84.TB11 = "0.00"
Frm84.TB12 = "0.00"
Frm84.TB13 = "0.00"
Frm84.TB14 = "0.00"
Frm84.TB15 = "0.00"
Frm84.TB22 = "0.00"
Frm84.TB16 = "0.00"
'Frm84.TB43 = 0 '% komisyen upah
Frm84.TB44 = "0.00" 'Komisyen upah
Frm84.L68_Text = "Modal (RM/g) :                   Jual (RM/g) :"

Frm84.L34_Text = "0.00"
Frm84.L42_Text = "0.00"
Frm84.L67_Text = "0.00"
Frm84.L69_Text = "0.00"

'Frm84.L31_Text = "0.00"
'Frm84.L32_Text = "0.00"
'Frm84.L81_Text = "0.00"
'Frm84.L82_Text = "0.00"

'Frm84.TB27 = "0.00"
'Frm84.TB28 = "0.00"
'Frm84.TB29 = "0.00"
'Frm84.TB32 = "0.00"
End Sub
Sub Frm84_Reset_Edit()
'on error resume next
Frm84.TB1 = vbNullString
Frm84.TB2 = vbNullString
Frm84.TB3 = "0.00"
Frm84.TB4 = "0.00"
Frm84.TB5 = "0.00"
Frm84.TB6 = "0.00"
Frm84.TB15 = "0.00"
Frm84.TB22 = "0.00"
'Frm84.TB7 = "0.00"
Frm84.TB8 = "0.00"
Frm84.TB9 = "0.00"
Frm84.TB10 = "0.00"
Frm84.TB12 = "0.00"
Frm84.TB13 = "0.00"
Frm84.TB11 = "0.00"
Frm84.TB14 = "0.00"
'Frm84.TB43 = 0 '% komisyen upah
Frm84.TB44 = "0.00" 'Komisyen upah
Frm84.L54_Text = vbNullString
Frm84.L55_Text = vbNullString

Frm84.L83_Text = "0" '0 : Stok kedai , 1 : Barang trade in/potong

Frm84.L48_Text = "0.00"
Frm84.L49_Text = vbNullString
Frm84.L50_Text = "0.00"
Frm84.L51_Text = vbNullString
Frm84.L52_Text = vbNullString
Frm84.L53_Text = vbNullString
Frm84.L12_Text = vbNullString
Frm84.L13_Text = vbNullString
Frm84.L34_Text = "0.00"
Frm84.L42_Text = "0.00"
Frm84.L68_Text = "Modal (RM/g) :                   Jual (RM/g) :"
Frm84.L67_Text = "0.00"
Frm84.L69_Text = "0.00"
Frm84.L39_Text = vbNullString

'Frm84.L75_Text = "0.00"
'Frm84.TB35 = "0"
'Frm84.L76_Text = "0"
''Frm84.L77_Text = "0"
'Frm84.TB36 = "0"
'Frm84.TB37 = "0"
'Frm84.L78_Text = "0.00"
''Frm84.L79_Text = 0

user_level = MDI_frm1.L4_Text

If user_level = "Administration" Or user_level = "Guest/User" Then

    Frm84.CB13.Visible = False
    Frm84.Label16.Visible = False
    
Else
    
    If G_GST_SYSTEM = "YES" Then
        Frm84.CB13.Visible = True
        Frm84.Label16.Visible = True
    Else
        Frm84.CB13.Visible = False
        Frm84.Label16.Visible = False
    End If
End If

Frm84.Frame2.Visible = False
Frm84.Frame3.Visible = False

Frm84.L48_Text = G_J_DISC_UPAH 'Peratusan penurunan upah kepada staff (Barang Kemas)
Frm84.L50_Text = G_J_DISC_PERMATA 'Peratusan penurunan harga barang permata kepada staff (Barang Permata)

'Set rs = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

'If Not rs.EOF Then
'    If rs!Default1 = "Default" Then
    
'        If Not IsNull(rs!upah_staff) Then
'            Frm84.L48_Text = rs!upah_staff 'Peratusan penurunan upah kepada staff
'        Else
'            Frm84.L48_Text = 0
'        End If
'        If Not IsNull(rs!diskaun_permata_staff) Then
'            Frm84.L50_Text = rs!diskaun_permata_staff 'Peratusan penurunan harga barang permata kepada staff
'        Else
'            Frm84.L50_Text = 0
'        End If
        
        'If Not IsNull(rs!pemalar_bonus) Then
        '    Frm84.TB35 = rs!pemalar_bonus
        'Else
        '    Frm84.TB35 = 0
        'End If
        'If Not IsNull(rs!pemalar_tebus_bonus) Then
        '    Frm84.TB37 = rs!pemalar_tebus_bonus
        'Else
        '    Frm84.TB37 = 0
        'End If

'    End If
'End If

'rs.Close
'Set rs = Nothing
End Sub
Sub Frm84_Senarai_Jualan_Header()
'on error resume next
With Frm84.ListView2
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear

    Frm84.ListView2.ListItems.Clear
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "No. Siri Produk", 2000
    .ColumnHeaders.Add 5, , "Kategori Produk", 3500
    .ColumnHeaders.Add 6, , "Purity", 1200
    .ColumnHeaders.Add 7, , "Berat Asal (g)", 1500, 1
    .ColumnHeaders.Add 8, , "Berat Jualan (g)", 1700, 1
    .ColumnHeaders.Add 9, , "Harga Semasa (RM/g)", 2200, 1
    .ColumnHeaders.Add 10, , "Upah (RM)", 1300, 1
    .ColumnHeaders.Add 11, , "Harga Asal (RM)", 1700, 1
    .ColumnHeaders.Add 12, , "Diskaun (%)", 1500, 1
    .ColumnHeaders.Add 13, , "Harga Lepas Diskaun (RM)", 2500, 1
    .ColumnHeaders.Add 14, , "Adjustment (RM)", 1700, 1
    .ColumnHeaders.Add 15, , "Harga Jualan (RM)", 1900, 1
    .ColumnHeaders.Add 16, , "Jenis GST", 1300, 2
    .ColumnHeaders.Add 17, , "Harga Termasuk GST", 2100, 2
    .ColumnHeaders.Add 18, , "Jumlah GST (RM)", 1700, 1
    .ColumnHeaders.Add 19, , "Harga Dengan GST (RM)", 2500, 1
    .ColumnHeaders.Add 20, , "Komisen Per Gram (RM/g)", 2500, 1
    .ColumnHeaders.Add 21, , "Komisyen Upah (RM)", 2200, 1
    .ColumnHeaders.Add 22, , "Jumlah Komisen (RM)", 2200, 1
    
End With
End Sub
Sub Frm84_Senarai_Jualan()
'on error resume next
Dim Frm84_LM_HARGA_TANPA_GST As Double 'Harga Jualan Tanpa Cukai GST
Dim Frm84_LM_HARGA_DENGAN_GST As Double 'Harga Jualan Dengan Cukai GST
Dim Frm84_LM_GST_SR As Double 'Kutipan GST : SR
Dim Frm84_LM_GST_ZR As Double 'Kutipan GST : ZR
Dim Frm84_LM_JUMLAH_HARGA_SR As Double 'Total Harga Yang Dikenakan GST SR
Dim Frm84_LM_JUMLAH_HARGA_ZR As Double 'Total Harga Yang Dikenakan GST ZR
Dim Frm84_LM_BERAT As Double 'Berat Jualan
Dim Frm84_LM_JUALAN_GST As Double
Dim Frm84_LM_JUALAN_DENGAN_GST As Double
Dim Frm84_LM_JUALAN_TANPA_GST As Double
Dim Frm84_LM_HARGA_JUALAN_TANPA_GST As Double 'Harga jualan barang kemas tanpa GST
Dim frm84_LM_TOTAL_PAGE As Double

x = 0
Frm84_LM_HARGA_TANPA_GST = 0
Frm84_LM_HARGA_DENGAN_GST = 0
Frm84_LM_GST_SR = 0
Frm84_LM_GST_ZR = 0
Frm84_LM_JUMLAH_HARGA_SR = 0
Frm84_LM_JUMLAH_HARGA_ZR = 0
Frm84_LM_HARGA_JUALAN_TANPA_GST = 0 'Harga jualan barang kemas tanpa GST

frm84_PAGE_SIZE = 29
frm84_LM_TOTAL_PAGE = 0

re_gen_report:

LM_START_ROW = Frm84.L89_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + frm84_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm84.L90_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - frm84_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        Frm84.L87_Text = 1
    End If
End If

frm84_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from " & G_JUALAN_TEMP & " LIMIT " & LM_START_ROW & "," & frm84_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    Frm84_LM_JUALAN_GST = 0
    Frm84_LM_JUALAN_DENGAN_GST = 0
    Frm84_LM_JUALAN_TANPA_GST = 0
    
    If rs!Status = 1 Or rs!Status = 3 Or rs!Status = 4 Then

        x = x + 1
        If frm84_LM_PAGE_FOUND = 0 Then
            If Frm84.L90_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
                If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                    Frm84.L87_Text = Frm84.L87_Text + 1 'Paparan Page ke-xxx
                    frm84_LM_PAGE_FOUND = 1
                ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                    If IsNumeric(Frm84.L87_Text) Then
                        If Frm84.L87_Text <> 1 Then
                            Frm84.L87_Text = Frm84.L87_Text - 1 'Paparan Page ke-xxx
                            frm84_LM_PAGE_FOUND = 1
                        End If
                    End If
                End If
            End If
        End If
    
        Y = ((Frm84.L87_Text - 1) * frm84_PAGE_SIZE) + x
    
        With Frm84.ListView2.ListItems.Add(, , rs!ID)
        
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
            
            If Not IsNull(rs!purity) Then 'Purity
                .ListSubItems.Add , , rs!purity
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!Berat_Asal) Then 'Berat Asal (g)
                .ListSubItems.Add , , Format(rs!Berat_Asal, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!berat_jualan) Then 'Berat Jualan (g)
                .ListSubItems.Add , , Format(rs!berat_jualan, "#,##0.00")
                If IsNumeric(rs!berat_jualan) Then Frm84_LM_BERAT = Frm84_LM_BERAT + rs!berat_jualan 'Jumlah Berat Jualan
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!harga_Semasa) Then 'Harga Semasa Bagi Purity Ini (RM/g)
                .ListSubItems.Add , , Format(rs!harga_Semasa, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!UPAH) Then 'Upah (RM)
                .ListSubItems.Add , , Format(rs!UPAH, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!harga_asal) Then 'Harga Asal Item (RM)
                .ListSubItems.Add , , Format(rs!harga_asal, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!diskaun) Then 'Diskaun (%)
                .ListSubItems.Add , , Format(rs!diskaun, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!harga_lepas_diskaun) Then 'Harga Lepas Diskaun (RM)
                .ListSubItems.Add , , Format(rs!harga_lepas_diskaun, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!adjustment) Then 'Adjustment (RM)
                .ListSubItems.Add , , Format(rs!adjustment, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!harga_jualan) Then 'Harga Jualan (RM)
                .ListSubItems.Add , , Format(rs!harga_jualan, "#,##0.00")
                If IsNumeric(rs!harga_tanpa_gst) Then Frm84_LM_HARGA_TANPA_GST = Frm84_LM_HARGA_TANPA_GST + rs!harga_tanpa_gst 'Harga Jualan Tanpa GST (RM)
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!gst_ari_nashi) Then 'Jenis GST
            
                If rs!gst_ari_nashi = "ZR (L)" Then
                    .ListSubItems.Add , , "ZR(L)"  'Jenis GST : Zero Rated
                    If IsNumeric(rs!jumlah_gst) Then Frm84_LM_GST_ZR = Frm84_LM_GST_ZR + rs!jumlah_gst 'Jumlah Kutipan GST ZR(L)
                    If IsNumeric(rs!harga_dengan_gst) Then Frm84_LM_JUMLAH_HARGA_ZR = Frm84_LM_JUMLAH_HARGA_ZR + rs!harga_dengan_gst 'Total Harga Yang Dikenakan GST ZR
                ElseIf rs!gst_ari_nashi = "SR" Then
                    .ListSubItems.Add , , "SR"  'Jenis GST : Standard Rated
                    If IsNumeric(rs!jumlah_gst) Then Frm84_LM_GST_SR = Frm84_LM_GST_SR + rs!jumlah_gst 'Jumlah Kutipan GST SR
                    If IsNumeric(rs!harga_tanpa_gst) Then Frm84_LM_JUMLAH_HARGA_SR = Frm84_LM_JUMLAH_HARGA_SR + rs!harga_tanpa_gst 'Total Harga Yang Dikenakan GST SR
                End If
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!gst_include) Then
                If rs!gst_include = 0 Then
                    .ListSubItems.Add , , "Tidak"  'Harga Termasuk GST
                Else
                    .ListSubItems.Add , , "Ya" 'Harga Termasuk GST
                End If
            Else
                .ListSubItems.Add , , "Tidak"  'Harga Termasuk GST
            End If
            
            If Not IsNull(rs!jumlah_gst) Then 'Jumlah GST (RM)
                .ListSubItems.Add , , Format(rs!jumlah_gst, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!harga_jualan_dengan_gst) Then 'Harga Dengan GST (RM)
                .ListSubItems.Add , , Format(rs!harga_jualan_dengan_gst, "#,##0.00")
                
                If IsNumeric(rs!harga_jualan_dengan_gst) And IsNumeric(rs!jumlah_gst) Then
                    Frm84_LM_HARGA_JUALAN_TANPA_GST = Frm84_LM_HARGA_JUALAN_TANPA_GST + (rs!harga_jualan_dengan_gst - rs!jumlah_gst) 'Harga jualan barang kemas tanpa GST
                End If
                
                If Not IsNull(rs!harga_jualan_dengan_gst) Then
                    If IsNumeric(rs!harga_dengan_gst) Then Frm84_LM_HARGA_DENGAN_GST = Frm84_LM_HARGA_DENGAN_GST + rs!harga_dengan_gst 'Harga Jualan Dengan GST (RM)
                End If
        
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!komisyen_per_gram) Then 'Komisen Per Gram (RM/g)
                .ListSubItems.Add , , Format(rs!komisyen_per_gram, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!komisyen_upah) Then 'Jumlah Komisyen Bagi Upah (RM)
                .ListSubItems.Add , , Format(rs!komisyen_upah, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If
            
            If Not IsNull(rs!jumlah_komisyen) Then 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini (RM)
                .ListSubItems.Add , , Format(rs!jumlah_komisyen, "#,##0.00")
            Else
                .ListSubItems.Add , , ""
            End If

        End With

    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

LM_X = 0

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from " & G_JUALAN_TEMP & " where (status = 1 OR status = 3 OR status = 4)", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    frm84_LM_TOTAL_PAGE = Format(rs(0) / frm84_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, frm84_LM_TOTAL_PAGE, ".") <> 0 Then
    
        frm84_LM_PAGE = Split(frm84_LM_TOTAL_PAGE, ".")(0)
        frm84_LM_PAGE_LEBIHAN = Split(frm84_LM_TOTAL_PAGE, ".")(1)
        
        If frm84_LM_PAGE_LEBIHAN <> "00" Then
            Frm84.L88_Text = frm84_LM_PAGE + 1
        Else
            Frm84.L88_Text = frm84_LM_PAGE
        End If
        
    Else
    
        Frm84.L88_Text = frm84_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm84.L88_Text = 0
    End If
Else
    Frm84.L88_Text = 0
End If

If Not IsNull(rs(0)) Then LM_X = rs(0)

rs.Close
Set rs = Nothing

Frm84.L4_Text = LM_X
Frm84.L14_Text = LM_X
Frm84.L5_Text = Format(Frm84_LM_HARGA_TANPA_GST, "#,##0.00") 'Harga Jualan Tanpa GST (RM)
Frm84.L17_Text = Format(Frm84_LM_HARGA_JUALAN_TANPA_GST, "#,##0.00") 'Harga Jualan Tanpa GST (RM)
Frm84.L6_Text = Format(Frm84_LM_HARGA_DENGAN_GST, "#,##0.00") 'Harga Jualan Dengan GST (RM)
Frm84.L18_Text = Format(Frm84_LM_GST_SR + Frm84_LM_GST_ZR, "#,##0.00") 'Jumlah Cukai GST (RM)
Frm84.L19_Text = Format(Frm84_LM_HARGA_JUALAN_TANPA_GST + Frm84_LM_GST_SR + Frm84_LM_GST_ZR, "#,##0.00") 'Harga Jualan Dengan GST (RM)
Frm84.L7_Text = Format(Frm84_LM_JUMLAH_HARGA_ZR, "#,##0.00")  'Total Harga Yang Dikenakan GST ZR
Frm84.L9_Text = Format(Frm84_LM_GST_ZR, "#,##0.00")  'Jumlah Kutipan GST ZR(L)
Frm84.L10_Text = Format(Frm84_LM_JUMLAH_HARGA_SR, "#,##0.00")  'Total Harga Yang Dikenakan GST SR
Frm84.L11_Text = Format(Frm84_LM_GST_SR, "#,##0.00")  'Jumlah Kutipan GST SR
Frm84.L15_Text = Format(Frm84_LM_BERAT, "#,##0.00") 'Jumlah Berat Jualan

If x <> 0 Then
    Frm84.L89_Text = LM_START_ROW
End If

If Frm84.L87_Text <> vbNullString And IsNumeric(Frm84.L87_Text) Then
    If Frm84.L88_Text <> vbNullString And IsNumeric(Frm84.L88_Text) Then
        frm84_LM_CURR_PAGE = Frm84.L87_Text
        frm84_LM_TOTAL_PAGE = Frm84.L88_Text
        
        If frm84_LM_CURR_PAGE > frm84_LM_TOTAL_PAGE Then
            
            Frm84.L87_Text = Frm84.L87_Text - 1
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            GoTo re_gen_report:
            
        End If
    End If
End If
End Sub
Sub Frm84_Resit_Jualan()
'on error resume next
Dim Frm79_LM_TOTAL_BERAT As Double
Dim LM_JUMLAH_BAYAR As Double
Dim LM_CAJ_KAD As Double
Dim LM_CAJ_GST As Double
Dim LM_QTY As Single
Dim LM_QTY_CHECK As Single
Dim LM_INV_CHECK As Single

LM_INV_CHECK = 0
LM_QTY = 0
LM_QTY_CHECK = 0

LM_JUMLAH_BAYAR = 0
LM_CAJ_KAD = 0
LM_CAJ_GST = 0
    
DATA_FOUND = 0
Frm84_DATA_PEKERJA_FOUND = 0
Frm84_DATA_CUST_FOUND = 0 '0 : Pelanggan Tidak Berdaftar , 1 : Pelanggan Berdaftar , 2 : Ahli
Frm84_LM_KATEGORI = 1

If Len(G_MODE) = 0 Or Len(G_MIN_LEN) = 0 Or Len(G_MAX_LEN) = 0 Or Len(G_CODE) = 0 Then

    Call sys_config_membership

End If

'### Reset Maklumat Pembeli ### - Start
Report38.Sections("Section2").Controls("L5").Caption = vbNullString 'Maklumat Pembeli : Nama
Report38.Sections("Section2").Controls("L7").Caption = vbNullString 'Maklumat Pembeli : No. Telefon
Report38.Sections("Section5").Controls("L25").Caption = "-"
'### Reset Maklumat Pembeli ### - End

Report38.Sections("Section2").Controls("L34").Caption = vbNullString 'No. Keahlian
Report38.Sections("Section2").Controls("L34").Visible = False 'No. Keahlian (Caption)
Report38.Sections("Section2").Controls("L35").Visible = False 'No. Keahlian (Caption)

'### Reset maklumat kedai ### - Start
Report38.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report38.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report38.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report38.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report38.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

Report38.Sections("Section5").Controls("L28").Caption = vbNullString 'Caption berkenaan mata ganjaran
Report38.Sections("Section5").Controls("L28").Visible = False 'Caption berkenaan mata ganjaran
Report38.Sections("Section5").Controls("L32").Caption = ": 0" 'Caption berkenaan mata ganjaran terkumpul
'Report38.Sections("Section5").Controls("L32").Visible = False 'Caption berkenaan mata ganjaran terkumpul

Report38.Sections("Section5").Controls("L33").Caption = "0.00" 'Kupon Diskaun
Report38.Sections("Section5").Controls("L38").Caption = vbNullString 'Remarks

Report38.Sections("Section5").Controls("L29").Caption = "0 @ RM 0.00"
Report38.Sections("Section5").Controls("L30").Caption = ": 0"
Report38.Sections("Section5").Controls("L41").Caption = "0.00" 'Tunai
Report38.Sections("Section5").Controls("L42").Caption = "0.00" 'Online Banking
Report38.Sections("Section5").Controls("L43").Caption = "0.00" 'Kad Kredit
Report38.Sections("Section5").Controls("L44").Caption = "0.00" 'Simpanan Di Kedai

Report38.Sections("Section4").Controls("L300").Caption = vbNullString
Report38.Sections("Section3").Controls("L301").Caption = vbNullString
Report38.Sections("Section3").Controls("L302").Caption = vbNullString
Report38.Sections("Section3").Controls("L303").Caption = vbNullString
Report38.Sections("Section3").Controls("L304").Caption = vbNullString
Report38.Sections("Section3").Controls("L305").Caption = vbNullString
Report38.Sections("Section3").Controls("L306").Caption = vbNullString
Report38.Sections("Section3").Controls("L307").Caption = vbNullString
Report38.Sections("Section3").Controls("L308").Caption = vbNullString
Report38.Sections("Section3").Controls("L309").Caption = vbNullString
Report38.Sections("Section3").Controls("L310").Caption = vbNullString
Report38.Sections("Section3").Controls("L311").Caption = vbNullString
Report38.Sections("Section3").Controls("L312").Caption = vbNullString
Report38.Sections("Section3").Controls("L313").Caption = vbNullString
Report38.Sections("Section3").Controls("L314").Caption = vbNullString
Report38.Sections("Section3").Controls("L315").Caption = vbNullString
Report38.Sections("Section3").Controls("L316").Caption = vbNullString
Report38.Sections("Section3").Controls("L317").Caption = vbNullString
Report38.Sections("Section3").Controls("L318").Caption = vbNullString
Report38.Sections("Section3").Controls("L319").Caption = vbNullString
Report38.Sections("Section3").Controls("L320").Caption = vbNullString

If G_JENIS_HEADER = 1 Then '0 : Pre Printed , 1 : Sistem

    '### Maklumat kedai ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!nama_kedai) Then Report38.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
        If Not IsNull(rs!no_pendaftaran) Then Report38.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
        If Not IsNull(rs!alamat) Then Report38.Sections("Section4").Controls("L202").Caption = rs!alamat
        If Not IsNull(rs!no_tel) Then Report38.Sections("Section4").Controls("L203").Caption = rs!no_tel
        If Not IsNull(rs!no_id_gst) Then Report38.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
        If Not IsNull(rs!gst_ari_nashi) Then
            If rs!gst_ari_nashi = 0 Then
                Report38.Sections("Section2").Controls("L205").Caption = "INVOICE"
                
                Report38.Sections("Section2").Controls("label19").Visible = False
                Report38.Sections("Section1").Controls("Text5").Visible = False
                
                Report38.Sections("Section5").Controls("label53").Visible = False
                Report38.Sections("Section5").Controls("label55").Visible = False
                Report38.Sections("Section5").Controls("label56").Visible = False
                Report38.Sections("Section5").Controls("label61").Visible = False
                Report38.Sections("Section5").Controls("label62").Visible = False
                Report38.Sections("Section5").Controls("L20").Visible = False
                Report38.Sections("Section5").Controls("L21").Visible = False
                Report38.Sections("Section5").Controls("L22").Visible = False
                Report38.Sections("Section5").Controls("L23").Visible = False
                Report38.Sections("Section5").Controls("Line2").Visible = False
                Report38.Sections("Section5").Controls("Shape1").Visible = False
                
                Report38.Sections("Section5").Controls("L28").Left = 3750
                Report38.Sections("Section5").Controls("L36").Left = 3750
                Report38.Sections("Section5").Controls("L37").Left = 3750
                Report38.Sections("Section5").Controls("L30").Left = 5535
                Report38.Sections("Section5").Controls("L32").Left = 5535
                
            ElseIf rs!gst_ari_nashi = 1 Then
                Report38.Sections("Section2").Controls("L205").Caption = "TAX INVOICE"
            End If
        Else
            Report38.Sections("Section2").Controls("L205").Caption = "INVOICE"
        End If
        If Not IsNull(rs!check_invoice) Then LM_INV_CHECK = rs!check_invoice
        If Not IsNull(rs!qty_item) Then LM_QTY_CHECK = rs!qty_item
        
        If Not IsNull(rs!cawangan_header) Then Report38.Sections("Section4").Controls("L300").Caption = rs!cawangan_header
        If Not IsNull(rs!cawagan_a_1) Then Report38.Sections("Section3").Controls("L301").Caption = rs!cawagan_a_1
        If Not IsNull(rs!cawagan_a_2) Then Report38.Sections("Section3").Controls("L302").Caption = rs!cawagan_a_2
        If Not IsNull(rs!cawagan_a_3) Then Report38.Sections("Section3").Controls("L303").Caption = rs!cawagan_a_3
        If Not IsNull(rs!cawagan_a_4) Then Report38.Sections("Section3").Controls("L304").Caption = rs!cawagan_a_4
        If Not IsNull(rs!cawagan_b_1) Then Report38.Sections("Section3").Controls("L305").Caption = rs!cawagan_b_1
        If Not IsNull(rs!cawagan_b_2) Then Report38.Sections("Section3").Controls("L306").Caption = rs!cawagan_b_2
        If Not IsNull(rs!cawagan_b_3) Then Report38.Sections("Section3").Controls("L307").Caption = rs!cawagan_b_3
        If Not IsNull(rs!cawagan_b_4) Then Report38.Sections("Section3").Controls("L308").Caption = rs!cawagan_b_4
        If Not IsNull(rs!cawagan_c_1) Then Report38.Sections("Section3").Controls("L309").Caption = rs!cawagan_c_1
        If Not IsNull(rs!cawagan_c_2) Then Report38.Sections("Section3").Controls("L310").Caption = rs!cawagan_c_2
        If Not IsNull(rs!cawagan_c_3) Then Report38.Sections("Section3").Controls("L311").Caption = rs!cawagan_c_3
        If Not IsNull(rs!cawagan_c_4) Then Report38.Sections("Section3").Controls("L312").Caption = rs!cawagan_c_4
        If Not IsNull(rs!cawagan_d_1) Then Report38.Sections("Section3").Controls("L313").Caption = rs!cawagan_d_1
        If Not IsNull(rs!cawagan_d_2) Then Report38.Sections("Section3").Controls("L314").Caption = rs!cawagan_d_2
        If Not IsNull(rs!cawagan_d_3) Then Report38.Sections("Section3").Controls("L315").Caption = rs!cawagan_d_3
        If Not IsNull(rs!cawagan_d_4) Then Report38.Sections("Section3").Controls("L316").Caption = rs!cawagan_d_4
        If Not IsNull(rs!cawagan_e_1) Then Report38.Sections("Section3").Controls("L317").Caption = rs!cawagan_e_1
        If Not IsNull(rs!cawagan_e_2) Then Report38.Sections("Section3").Controls("L318").Caption = rs!cawagan_e_2
        If Not IsNull(rs!cawagan_e_3) Then Report38.Sections("Section3").Controls("L319").Caption = rs!cawagan_e_3
        If Not IsNull(rs!cawagan_e_4) Then Report38.Sections("Section3").Controls("L320").Caption = rs!cawagan_e_4
        If Not IsNull(rs!cawagan_f_1) Then Report38.Sections("Section3").Controls("L321").Caption = rs!cawagan_f_1
        If Not IsNull(rs!cawagan_f_2) Then Report38.Sections("Section3").Controls("L322").Caption = rs!cawagan_f_2
        If Not IsNull(rs!cawagan_f_3) Then Report38.Sections("Section3").Controls("L323").Caption = rs!cawagan_f_3
        If Not IsNull(rs!cawagan_f_4) Then Report38.Sections("Section3").Controls("L324").Caption = rs!cawagan_f_4
        
    End If
    
    rs.Close
    Set rs = Nothing
    '### Maklumat kedai ### - End
    
    'Report38.Sections("Section4").Visible = True
    
    'Report38.Sections("Section4").Controls("Label6").Top = 1125
    'Report38.Sections("Section4").Controls("Label7").Top = 1125
    'Report38.Sections("Section4").Controls("L5").Top = 1125
    'Report38.Sections("Section4").Controls("Label13").Top = 930
    'Report38.Sections("Section4").Controls("Label15").Top = 1260
    'Report38.Sections("Section4").Controls("Label16").Top = 1260
    'Report38.Sections("Section4").Controls("L7").Top = 1260
    'Report38.Sections("Section4").Controls("L34").Top = 1260
    'Report38.Sections("Section4").Controls("L35").Top = 1260
    'Report38.Sections("Section4").Controls("Shape2").Top = 720
    'Report38.Sections("Section4").Controls("L205").Top = 705
    'Report38.Sections("Section4").Controls("Label4").Top = 990
    'Report38.Sections("Section4").Controls("Label11").Top = 990
    'Report38.Sections("Section4").Controls("L3").Top = 990
    
    'Report38.Sections("Section4").Controls("Label46").Top = 1290
    'Report38.Sections("Section4").Controls("Label12").Top = 1290
    'Report38.Sections("Section4").Controls("L17").Top = 1290
    
    'Report38.Sections("Section4").Height = 1400
    
Else
    'Report38.Sections("Section4").Visible = False
    'Report38.Sections("Section4").Controls("L200").Visible = False
    'Report38.Sections("Section4").Controls("L201").Visible = False
    'Report38.Sections("Section4").Controls("L202").Visible = False
    'Report38.Sections("Section4").Controls("L203").Visible = False
    'Report38.Sections("Section4").Controls("L204").Visible = False

    'Report38.Sections("Section4").Controls("Label6").Top = 510
    'Report38.Sections("Section4").Controls("Label7").Top = 510
    'Report38.Sections("Section4").Controls("L5").Top = 510
    'Report38.Sections("Section4").Controls("Label13").Top = 285
    'Report38.Sections("Section4").Controls("Label15").Top = 645
    'Report38.Sections("Section4").Controls("Label16").Top = 645
    'Report38.Sections("Section4").Controls("L7").Top = 645
    'Report38.Sections("Section4").Controls("L34").Top = 645
    'Report38.Sections("Section4").Controls("L35").Top = 645
    
    'Report38.Sections("Section4").Controls("Shape2").Top = 300
    'Report38.Sections("Section4").Controls("L205").Top = 285
    
    'Report38.Sections("Section4").Controls("Label4").Top = 550
    'Report38.Sections("Section4").Controls("Label11").Top = 550
    'Report38.Sections("Section4").Controls("L3").Top = 550
    
    'Report38.Sections("Section4").Controls("Label5").Top = 700
    'Report38.Sections("Section4").Controls("Label3").Top = 700
    'Report38.Sections("Section4").Controls("L4").Top = 700
    
    'Report38.Sections("Section4").Controls("Label46").Top = 850
    'Report38.Sections("Section4").Controls("Label12").Top = 850
    'Report38.Sections("Section4").Controls("L17").Top = 850
    
    'Report38.Sections("Section4").Height = 990
    Report38.Sections("Section4").Controls("L200").Visible = False
    Report38.Sections("Section4").Controls("L201").Visible = False
    Report38.Sections("Section4").Controls("L202").Visible = False
    Report38.Sections("Section4").Controls("L203").Visible = False
    Report38.Sections("Section4").Controls("L204").Visible = False

    Report38.Sections("Section4").Controls("L200").Height = 0
    Report38.Sections("Section4").Controls("L201").Height = 0
    Report38.Sections("Section4").Controls("L202").Height = 0
    Report38.Sections("Section4").Controls("L203").Height = 0
    Report38.Sections("Section4").Controls("L204").Height = 0
    
    Report38.Sections("Section4").Controls("L40").Height = G_TOP
    
End If

LM_INVOICE_RASMI = 0

user_level = MDI_frm1.L4_Text

If user_level = "Guest/User" Then

    LM_INVOICE_RASMI = 1

End If

LM_TI = vbNullString
LM_TI_SUSUT_NILAI_MODE = 0

Report38.Sections("Section5").Controls("L24").Caption = "0.00" 'Postage
LM_SUSUT_NILAI_MODE = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where no_resit='" & G_No_RESIT_JUALAN & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!jenis_trade_in) Then
        If rs!jenis_trade_in = 3 Then
            LM_SUSUT_NILAI_MODE = 1
        Else
            If Not IsNull(rs!remarks) Then Report38.Sections("Section5").Controls("L38").Caption = rs!remarks 'Remarks
        End If
    End If
    If LM_INVOICE_RASMI = 0 Then
        Report38.Sections("Section2").Controls("L3").Caption = G_No_RESIT_JUALAN 'No. Resit Jualan
    Else
        If Not IsNull(rs!no_invoice_r) Then
            Report38.Sections("Section2").Controls("L3").Caption = rs!no_invoice_r 'No. Resit Jualan
        Else
            Report38.Sections("Section2").Controls("L3").Caption = G_No_RESIT_JUALAN 'No. Resit Jualan
        End If
    End If
    If Not IsNull(rs!tarikh) Then Report38.Sections("Section2").Controls("L4").Caption = rs!tarikh 'Jumlah Harga Jualan (RM)
    If Not IsNull(rs!harga_barang_dengan_gst) Then Report38.Sections("Section5").Controls("L8").Caption = Format(rs!harga_barang_dengan_gst, "#,##0.00") 'Jumlah Harga Jualan (RM)
    If Not IsNull(rs!diskaun) Then Report38.Sections("Section5").Controls("L9").Caption = Format(rs!diskaun, "#,##0.00") 'Jumlah Diskaun (%)
    If Not IsNull(rs!harga_lepas_diskaun) Then Report38.Sections("Section5").Controls("L10").Caption = Format(rs!harga_lepas_diskaun, "#,##0.00") 'Harga Selepas Diskaun (RM)
    If Not IsNull(rs!adjustment) Then Report38.Sections("Section5").Controls("L11").Caption = Format(rs!adjustment, "#,##0.00") 'Adjustment (RM)
    If Not IsNull(rs!harga_jualan) Then Report38.Sections("Section5").Controls("L12").Caption = Format(rs!harga_jualan, "#,##0.00") 'Jumlah Harga Jualan (RM)
    
    If Not IsNull(rs!jumlah_trade_in) Then
        LM_HARGA_TI = rs!jumlah_trade_in 'Jumlah Resit Trade In (RM)
    Else
        LM_HARGA_TI = "0.00" 'Jumlah Resit Trade In (RM)
    End If
    
    'If Not IsNull(rs!jumlah_trade_in) Then
    '    Report38.Sections("Section5").Controls("L13").Caption = Format(rs!jumlah_trade_in, "#,##0.00") 'Jumlah Resit Trade In (RM)
    'Else
    '    Report38.Sections("Section5").Controls("L13").Caption = "0.00" 'Jumlah Resit Trade In (RM)
    'End If
    
    
    If Not IsNull(rs!flag_trade_in) Then
        If rs!flag_trade_in = 0 Then
            Report38.Sections("Section5").Controls("L14").Caption = "-" 'No. Resit Trade In
        Else
            LM_TI = rs!no_resit_trade_in
            If Not IsNull(rs!no_resit_trade_in) Then Report38.Sections("Section5").Controls("L14").Caption = rs!no_resit_trade_in 'No. Resit Trade In
        End If
    End If
    If Not IsNull(rs!jenis_trade_in) Then
        If rs!jenis_trade_in = 3 Then LM_TI_SUSUT_NILAI_MODE = 1
    End If
    If Not IsNull(rs!flag_bayaran) Then
        If rs!flag_bayaran = 0 Then
            Report38.Sections("Section5").Controls("L15").Caption = "Jumlah Bayaran"
        Else
            Report38.Sections("Section5").Controls("L15").Caption = "Lebihan Kedai Perlu Bayar"
        End If
    End If
    If Not IsNull(rs!jumlah_perlu_bayar) Then
        If IsNumeric(rs!jumlah_perlu_bayar) Then LM_JUMLAH_BAYAR = rs!jumlah_perlu_bayar
    End If
    If Not IsNull(rs!jumlah_cas_kad_kredit) Then
        If IsNumeric(rs!jumlah_cas_kad_kredit) Then LM_CAJ_KAD = rs!jumlah_cas_kad_kredit
    End If
    If Not IsNull(rs!gst_kad_kredit) Then
        If IsNumeric(rs!gst_kad_kredit) Then LM_CAJ_GST = rs!gst_kad_kredit
    End If

    Report38.Sections("Section5").Controls("L16").Caption = Format(LM_JUMLAH_BAYAR + LM_CAJ_KAD + LM_CAJ_GST, "#,##0.00") 'Jumlah Bayaran Yang Perlu Dibuat (RM)
    
    If Not IsNull(rs!kuantiti_barang) Then
        Report38.Sections("Section5").Controls("L18").Caption = rs!kuantiti_barang 'Kuantiti Barang Yang Dijual
        LM_QTY = rs!kuantiti_barang
    End If
    
    If Not IsNull(rs!JUMLAH_BERAT) Then Report38.Sections("Section5").Controls("L19").Caption = Format(rs!JUMLAH_BERAT, "#,##0.00 g") 'Jumlah Berat Barang Yang Dijual
    If Not IsNull(rs!gst_sr_harga) Then Report38.Sections("Section5").Controls("L20").Caption = Format(rs!gst_sr_harga, "#,##0.00") 'Harga Keseluruhan Bagi Barang SR
    If Not IsNull(rs!gst_sr_cukai) Then Report38.Sections("Section5").Controls("L21").Caption = Format(rs!gst_sr_cukai, "#,##0.00") 'Jumlah Cukai Bagi SR
    If Not IsNull(rs!gst_zr_harga) Then Report38.Sections("Section5").Controls("L22").Caption = Format(rs!gst_zr_harga, "#,##0.00") 'Harga Keseluruhan Bagi Barang ZR
    If Not IsNull(rs!gst_zr_cukai) Then Report38.Sections("Section5").Controls("L23").Caption = Format(rs!gst_zr_cukai, "#,##0.00") 'Jumlah Cukai Bagi ZR
    If Not IsNull(rs!caj_pos) Then Report38.Sections("Section5").Controls("L24").Caption = Format(rs!caj_pos, "#,##0.00") '
    If Not IsNull(rs!no_tracking) Then
        Report38.Sections("Section5").Controls("L25").Caption = rs!no_tracking
    Else
        Report38.Sections("Section5").Controls("L25").Caption = "-"
    End If
    If Not IsNull(rs!point_ari_nashi) Then
        If rs!point_ari_nashi = 0 And Not IsNull(rs!jumlah_point) Then
            Report38.Sections("Section5").Controls("L28").Caption = "Anda telah terlepas mata ganjaran sebanyak " & Format(rs!jumlah_point, "#,##0") & " daripada pembelian ini kerana tidak mempunyai kad keahlian kedai." & vbCrLf & _
                                                                    "Sila daftar keahlian dengan kedai untuk mendapat mata ganjaran bagi pembelian berikutnya." 'Caption berkenaan mata ganjaran
            Report38.Sections("Section5").Controls("L28").Visible = True 'Caption berkenaan mata ganjaran
            Report38.Sections("Section5").Controls("L32").Visible = True 'Caption berkenaan mata ganjaran terkumpul
        End If
    End If
    If Not IsNull(rs!kupon_diskaun) Then Report38.Sections("Section5").Controls("L33").Caption = Format(rs!kupon_diskaun, "#,##0.00") 'Kupon Diskaun
    
    If Not IsNull(rs!no_pekerja) Then
        Frm84_LM_No_PEKERJA = rs!no_pekerja
        Frm84_DATA_PEKERJA_FOUND = 1
    End If

'1:  Pelanggan
'2:  Member
'3:  RAF
'4:  Pengedar
'5:  Normal Dealer
'6:  Master Dealer

    If Not IsNull(rs!kategori_pembeli) Then 'Kategori Pembeli
        Frm84_LM_KATEGORI = rs!kategori_pembeli
    End If
    If Not IsNull(rs!no_rujukan_pembeli) Then
        Frm84_LM_No_CUST = rs!no_rujukan_pembeli
        Frm84_DATA_CUST_FOUND = 1 '0 : Pelanggan Tidak Berdaftar , 1 : Pelanggan Berdaftar , 2 : Ahli
    End If
    If Not IsNull(rs!tunai) Then Report38.Sections("Section5").Controls("L41").Caption = Format(rs!tunai, "#,##0.00") 'Tunai
    If Not IsNull(rs!bank_in) Then Report38.Sections("Section5").Controls("L42").Caption = Format(rs!bank_in, "#,##0.00") 'Online Banking
    If Not IsNull(rs!kad_kredit) Then Report38.Sections("Section5").Controls("L43").Caption = Format(rs!kad_kredit, "#,##0.00") 'Kad Kredit
    If Not IsNull(rs!duit_simpanan_kedai) Then Report38.Sections("Section5").Controls("L44").Caption = Format(rs!duit_simpanan_kedai, "#,##0.00") 'Simpanan Di Kedai
    
    
    DATA_FOUND = 1
End If

rs.Close
Set rs = Nothing

If LM_SUSUT_NILAI_MODE = 1 Then
    H1 = H1 & "Maklumat Trade In" & vbCrLf
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 93_trade_in_susut_niai where no_invoice='" & G_No_RESIT_JUALAN & "' AND status = 1 order by jenis ASC", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        LM_BERAT = 0
        LM_HARGA_SEMASA = 0
        LM_HARGA = 0
        LM_JENIS = vbNullString
        
        If Not IsNull(rs!Berat) Then LM_BERAT = rs!Berat
        If Not IsNull(rs!harga_Semasa) Then LM_HARGA_SEMASA = rs!harga_Semasa
        If Not IsNull(rs!harga) Then LM_HARGA = rs!harga
        If Not IsNull(rs!jenis) Then
            If rs!jenis = 0 Then LM_JENIS = "Trade In : "
            If rs!jenis = 1 Then LM_JENIS = "Buyback : "
            If rs!jenis = 2 Then LM_JENIS = "Caj Pertukaran : "
        End If
        If rs!jenis = 0 Or rs!jenis = 1 Then H1 = H1 & LM_JENIS & Format(LM_BERAT, "#,##0.00 g") & " X RM " & Format(LM_HARGA_SEMASA, "#,##0.00") & "/g = RM " & Format(LM_HARGA, "#,##0.00") & vbCrLf
        If rs!jenis = 2 Then H1 = H1 & LM_JENIS & " RM " & Format(LM_HARGA, "#,##0.00") & vbCrLf
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
    Report38.Sections("Section5").Controls("L38").Caption = H1 'Remarks
End If

If LM_TI <> vbNullString Then
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select SUM(berat) from data_database where bill_No_Trade_In='" & LM_TI & "' AND statusitem <> 0", cn, adOpenKeyset, adLockOptimistic

    If Not IsNull(rs(0)) Then LM_BERAT_TI = rs(0)
    
    rs.Close
    Set rs = Nothing
    
    Report38.Sections("Section5").Controls("L13").Caption = Format(LM_HARGA_TI, "#,##0.00") & " (" & Format(LM_BERAT_TI, "#,##0.00 g") & ")" 'Jumlah Resit Trade In (RM)
Else
    Report38.Sections("Section5").Controls("L13").Caption = "0.00 (0.00 g)" 'Jumlah Resit Trade In (RM)
End If

'GoTo skiplaa:

LM_JENIS_INVOICE = 0 '0 : Invoice tiada limit , 1 : Invoice dengan limit kuantiti

If LM_INV_CHECK = 1 Then

    If LM_QTY <= LM_QTY_CHECK Then
        
        Report38.Sections("Section1").Controls("label36").Visible = False
        Report38.Sections("Section1").Controls("Text6").Left = 135
        Report38.Sections("Section1").Controls("Text6").Width = 1800
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "TRUNCATE TABLE " & G_INVOICE_TEMP & ""
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "insert into " & G_INVOICE_TEMP & "(no_siri_produk,kategori_produk,purity,berat_jualan,harga_semasa,upah,harga_jualan_dengan_gst,gst_ari_nashi)" & _
                    "select no_siri_produk,kategori_produk,purity,berat_jualan,harga_semasa,upah,harga_jualan_dengan_gst,gst_ari_nashi from 23_senarai_jualan WHERE no_resit='" & G_No_RESIT_JUALAN & "' AND status_rekod = 1"
                    
        Set rs = cn.Execute(strsql)
        Set rs = Nothing

        For Z = 1 To LM_QTY_CHECK
            
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from " & G_INVOICE_TEMP & " where ID='" & Z & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!no_siri_Produk) Then
                
                    rs!no_siri_Produk = "No. Siri :  " & rs!no_siri_Produk
                    rs.Update
                    
                End If
                
            Else
            
                rs.AddNew
                rs.Update
        
            End If
            
        Next Z
        
        LM_JENIS_INVOICE = 1 '0 : Invoice tiada limit , 1 : Invoice dengan limit kuantiti
        
    End If
    
End If

'skiplaa:

If DATA_FOUND = 1 Then
    If Frm84_DATA_PEKERJA_FOUND = 1 Then
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where NoPekerja='" & Frm84_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!Samaran) Then Report38.Sections("Section2").Controls("L17").Caption = rs!Samaran 'Nama Samaran
        End If
        
        rs.Close
        Set rs = Nothing
    End If
    
'### Data jika pembeli TIDAK berdaftar ### - Start
    If Frm84_DATA_CUST_FOUND = 0 Then '0 : Pelanggan Tidak Berdaftar , 1 : Pelanggan Berdaftar , 2 : Ahli
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 44_senarai_pelanggan where no_resit='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!Nama) Then Report38.Sections("Section2").Controls("L5").Caption = rs!Nama 'Maklumat Pembeli : Nama
            If Not IsNull(rs!no_tel) Then Report38.Sections("Section2").Controls("L7").Caption = rs!no_tel 'Maklumat Pembeli : No. Telefon
        End If
        
        rs.Close
        Set rs = Nothing
    End If
'### Data jika pembeli TIDAK berdaftar ### - End

'### Data jika pembeli adalah berdaftar ### - Start
    If Frm84_DATA_CUST_FOUND = 1 Then '0 : Pelanggan Tidak Berdaftar , 1 : Pelanggan Berdaftar , 2 : Ahli
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm84_LM_No_CUST & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!Nama) Then Report38.Sections("Section2").Controls("L5").Caption = rs!Nama 'Maklumat Pembeli : Nama
            If Not IsNull(rs!no_tel) Then Report38.Sections("Section2").Controls("L7").Caption = rs!no_tel 'Maklumat Pembeli : No. Telefon
            If Not IsNull(rs!baki_point) Then
                Report38.Sections("Section5").Controls("L32").Caption = ": " & Format(rs!baki_point, "#,##0") 'Jumlah mata ganjaran terkumpul
                Report38.Sections("Section5").Controls("L32").Visible = True
            End If
            
            If Not IsNull(rs!no_pelanggan) Then
                If Not IsNull(rs!kategori_pelanggan) Then
                    If rs!kategori_pelanggan = 1 Then
                        Report38.Sections("Section2").Controls("L34").Caption = rs!no_pelanggan & " (Pelanggan Biasa)" 'No. Keahlian
                    ElseIf rs!kategori_pelanggan = 2 Then
                        Report38.Sections("Section2").Controls("L34").Caption = rs!no_pelanggan & " (Ahli Biasa)" 'No. Keahlian
                    ElseIf rs!kategori_pelanggan = 3 Then
                        Report38.Sections("Section2").Controls("L34").Caption = rs!no_pelanggan & " (Silver)" 'No. Keahlian
                    ElseIf rs!kategori_pelanggan = 4 Then
                        Report38.Sections("Section2").Controls("L34").Caption = rs!no_pelanggan & " (Gold)" 'No. Keahlian
                    ElseIf rs!kategori_pelanggan = 5 Then
                        Report38.Sections("Section2").Controls("L34").Caption = rs!no_pelanggan & " (Platinum)" 'No. Keahlian
                    End If
                    'Report38.Sections("Section2").Controls("L34").Visible = True 'No. Keahlian (Caption)
                    'Report38.Sections("Section2").Controls("L35").Visible = True 'No. Keahlian (Caption)
                End If
            End If
            
        End If
        
        rs.Close
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 71_tebus_agih_point where no_invoice='" & G_No_RESIT_JUALAN & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!jumlah_tebus_point) And Not IsNull(rs!nilaian_tebus_point) Then
                Report38.Sections("Section5").Controls("L29").Caption = rs!jumlah_tebus_point & " @ RM " & Format(rs!nilaian_tebus_point, "#,##0.00")
            End If
            If Not IsNull(rs!jumlah_peroleh_point) Then
                Report38.Sections("Section5").Controls("L30").Caption = ": " & Format(rs!jumlah_peroleh_point, "#,##0")
            End If
        
        End If
        
        rs.Close
        Set rs = Nothing

    End If
'### Data jika pembeli adalah berdaftar ### - End

    If G_MODE = "YES" Then
        Report38.Sections("Section5").Controls("L28").Visible = True
        Report38.Sections("Section5").Controls("L30").Visible = True
        Report38.Sections("Section5").Controls("L32").Visible = True
        Report38.Sections("Section5").Controls("L36").Visible = True
        Report38.Sections("Section5").Controls("L37").Visible = True
    Else
        Report38.Sections("Section5").Controls("L28").Visible = False
        Report38.Sections("Section5").Controls("L30").Visible = False
        Report38.Sections("Section5").Controls("L32").Visible = False
        Report38.Sections("Section5").Controls("L36").Visible = False
        Report38.Sections("Section5").Controls("L37").Visible = False
    End If
    
    '### Paparan Resit ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    If LM_JENIS_INVOICE = 0 Then '0 : Invoice tiada limit , 1 : Invoice dengan limit kuantiti
        
        rs.Open "select * from 23_senarai_jualan where no_resit='" & G_No_RESIT_JUALAN & "' AND status_rekod = 1", cn, adOpenKeyset, adLockOptimistic
    
    ElseIf LM_JENIS_INVOICE = 1 Then '0 : Invoice tiada limit , 1 : Invoice dengan limit kuantiti
        
        rs.Open "select * from " & G_INVOICE_TEMP & "", cn, adOpenKeyset, adLockOptimistic

    End If
    
    While rs.EOF = False
        Set Report38.DataSource = rs
        If G_PREVIEW = 1 Then Report38.Show
        rs.MoveNext
    Wend
    
    'rs.Close
    Set rs = Nothing
    
    'If LM_QTY <= 5 Then
    
    '    Report38.Sections("Section1").Height = 500
        
    'End If
    '### Paparan Resit ### - End
    
    If G_PREVIEW = 0 Then Report38.PrintReport
     
    G_No_RESIT_JUALAN = vbNullString
End If
End Sub

Sub Frm84_Resit_Jualan_ori()
'on error resume next
Dim Frm79_LM_TOTAL_BERAT As Double
Dim LM_JUMLAH_BAYAR As Double
Dim LM_CAJ_KAD As Double
Dim LM_CAJ_GST As Double
Dim LM_QTY As Single
Dim LM_QTY_CHECK As Single
Dim LM_INV_CHECK As Single

LM_INV_CHECK = 0
LM_QTY = 0
LM_QTY_CHECK = 0

LM_JUMLAH_BAYAR = 0
LM_CAJ_KAD = 0
LM_CAJ_GST = 0
    
DATA_FOUND = 0
Frm84_DATA_PEKERJA_FOUND = 0
Frm84_DATA_CUST_FOUND = 0 '0 : Pelanggan Tidak Berdaftar , 1 : Pelanggan Berdaftar , 2 : Ahli
Frm84_LM_KATEGORI = 1
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
'        PRINTER_FOUND = 1 '0 : Printer Not Found , 1 : Printer Found
'        Exit For
'    End If
'Next

If Len(G_MODE) = 0 Or Len(G_MIN_LEN) = 0 Or Len(G_MAX_LEN) = 0 Or Len(G_CODE) = 0 Then

    Call sys_config_membership

End If

'### Reset Maklumat Pembeli ### - Start
Report38.Sections("Section2").Controls("L5").Caption = vbNullString 'Maklumat Pembeli : Nama
Report38.Sections("Section2").Controls("L7").Caption = vbNullString 'Maklumat Pembeli : No. Telefon
Report38.Sections("Section5").Controls("L25").Caption = "-"
'### Reset Maklumat Pembeli ### - End

Report38.Sections("Section2").Controls("L34").Caption = vbNullString 'No. Keahlian
Report38.Sections("Section2").Controls("L34").Visible = False 'No. Keahlian (Caption)
Report38.Sections("Section2").Controls("L35").Visible = False 'No. Keahlian (Caption)

'### Reset maklumat kedai ### - Start
Report38.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report38.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report38.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report38.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report38.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

Report38.Sections("Section5").Controls("L28").Caption = vbNullString 'Caption berkenaan mata ganjaran
Report38.Sections("Section5").Controls("L28").Visible = False 'Caption berkenaan mata ganjaran
Report38.Sections("Section5").Controls("L32").Caption = ": 0" 'Caption berkenaan mata ganjaran terkumpul
'Report38.Sections("Section5").Controls("L32").Visible = False 'Caption berkenaan mata ganjaran terkumpul

Report38.Sections("Section5").Controls("L33").Caption = "0.00" 'Kupon Diskaun

Report38.Sections("Section5").Controls("L29").Caption = "0 @ RM 0.00"
Report38.Sections("Section5").Controls("L30").Caption = ": 0"

If LM_HEADER = 1 Then '0 : Pre Printed , 1 : Sistem
    '### Maklumat kedai ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!nama_kedai) Then Report38.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
        If Not IsNull(rs!no_pendaftaran) Then Report38.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
        If Not IsNull(rs!alamat) Then Report38.Sections("Section4").Controls("L202").Caption = rs!alamat
        If Not IsNull(rs!no_tel) Then Report38.Sections("Section4").Controls("L203").Caption = rs!no_tel
        If Not IsNull(rs!no_id_gst) Then Report38.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
        If Not IsNull(rs!gst_ari_nashi) Then
            If rs!gst_ari_nashi = 0 Then
                Report38.Sections("Section2").Controls("L205").Caption = "INVOICE"
                
                Report38.Sections("Section2").Controls("label19").Visible = False
                Report38.Sections("Section1").Controls("Text5").Visible = False
                
                Report38.Sections("Section5").Controls("label53").Visible = False
                Report38.Sections("Section5").Controls("label55").Visible = False
                Report38.Sections("Section5").Controls("label56").Visible = False
                Report38.Sections("Section5").Controls("label61").Visible = False
                Report38.Sections("Section5").Controls("label62").Visible = False
                Report38.Sections("Section5").Controls("L20").Visible = False
                Report38.Sections("Section5").Controls("L21").Visible = False
                Report38.Sections("Section5").Controls("L22").Visible = False
                Report38.Sections("Section5").Controls("L23").Visible = False
                Report38.Sections("Section5").Controls("Line2").Visible = False
                Report38.Sections("Section5").Controls("Shape1").Visible = False
                
                Report38.Sections("Section5").Controls("L28").Left = 3750
                Report38.Sections("Section5").Controls("L36").Left = 3750
                Report38.Sections("Section5").Controls("L37").Left = 3750
                Report38.Sections("Section5").Controls("L30").Left = 5535
                Report38.Sections("Section5").Controls("L32").Left = 5535
                
            ElseIf rs!gst_ari_nashi = 1 Then
                Report38.Sections("Section2").Controls("L205").Caption = "TAX INVOICE"
            End If
        Else
            Report38.Sections("Section2").Controls("L205").Caption = "INVOICE"
        End If
        If Not IsNull(rs!check_invoice) Then LM_INV_CHECK = rs!check_invoice
        If Not IsNull(rs!qty_item) Then LM_QTY_CHECK = rs!qty_item
        
    End If
    
    rs.Close
    Set rs = Nothing
    '### Maklumat kedai ### - End
    
    Report38.Sections("Section4").Visible = True
    
    'Report38.Sections("Section4").Controls("Label6").Top = 1125
    'Report38.Sections("Section4").Controls("Label7").Top = 1125
    'Report38.Sections("Section4").Controls("L5").Top = 1125
    'Report38.Sections("Section4").Controls("Label13").Top = 930
    'Report38.Sections("Section4").Controls("Label15").Top = 1260
    'Report38.Sections("Section4").Controls("Label16").Top = 1260
    'Report38.Sections("Section4").Controls("L7").Top = 1260
    'Report38.Sections("Section4").Controls("L34").Top = 1260
    'Report38.Sections("Section4").Controls("L35").Top = 1260
    'Report38.Sections("Section4").Controls("Shape2").Top = 720
    'Report38.Sections("Section4").Controls("L205").Top = 705
    'Report38.Sections("Section4").Controls("Label4").Top = 990
    'Report38.Sections("Section4").Controls("Label11").Top = 990
    'Report38.Sections("Section4").Controls("L3").Top = 990
    
    'Report38.Sections("Section4").Controls("Label46").Top = 1290
    'Report38.Sections("Section4").Controls("Label12").Top = 1290
    'Report38.Sections("Section4").Controls("L17").Top = 1290
    
    'Report38.Sections("Section4").Height = 1400
    
Else
    Report38.Sections("Section4").Visible = False
    'Report38.Sections("Section4").Controls("L200").Visible = False
    'Report38.Sections("Section4").Controls("L201").Visible = False
    'Report38.Sections("Section4").Controls("L202").Visible = False
    'Report38.Sections("Section4").Controls("L203").Visible = False
    'Report38.Sections("Section4").Controls("L204").Visible = False

    'Report38.Sections("Section4").Controls("Label6").Top = 510
    'Report38.Sections("Section4").Controls("Label7").Top = 510
    'Report38.Sections("Section4").Controls("L5").Top = 510
    'Report38.Sections("Section4").Controls("Label13").Top = 285
    'Report38.Sections("Section4").Controls("Label15").Top = 645
    'Report38.Sections("Section4").Controls("Label16").Top = 645
    'Report38.Sections("Section4").Controls("L7").Top = 645
    'Report38.Sections("Section4").Controls("L34").Top = 645
    'Report38.Sections("Section4").Controls("L35").Top = 645
    
    'Report38.Sections("Section4").Controls("Shape2").Top = 300
    'Report38.Sections("Section4").Controls("L205").Top = 285
    
    'Report38.Sections("Section4").Controls("Label4").Top = 550
    'Report38.Sections("Section4").Controls("Label11").Top = 550
    'Report38.Sections("Section4").Controls("L3").Top = 550
    
    'Report38.Sections("Section4").Controls("Label5").Top = 700
    'Report38.Sections("Section4").Controls("Label3").Top = 700
    'Report38.Sections("Section4").Controls("L4").Top = 700
    
    'Report38.Sections("Section4").Controls("Label46").Top = 850
    'Report38.Sections("Section4").Controls("Label12").Top = 850
    'Report38.Sections("Section4").Controls("L17").Top = 850
    
    'Report38.Sections("Section4").Height = 990
End If

LM_INVOICE_RASMI = 0

user_level = MDI_frm1.L4_Text

If user_level = "Guest/User" Then

    LM_INVOICE_RASMI = 1

End If

Report38.Sections("Section5").Controls("L24").Caption = "0.00" 'Postage

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where no_resit='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If LM_INVOICE_RASMI = 0 Then
        Report38.Sections("Section2").Controls("L3").Caption = G_No_RESIT_JUALAN 'No. Resit Jualan
    Else
        If Not IsNull(rs!no_invoice_r) Then
            Report38.Sections("Section2").Controls("L3").Caption = rs!no_invoice_r 'No. Resit Jualan
        Else
            Report38.Sections("Section2").Controls("L3").Caption = G_No_RESIT_JUALAN 'No. Resit Jualan
        End If
    End If
    If Not IsNull(rs!tarikh) Then Report38.Sections("Section2").Controls("L4").Caption = rs!tarikh 'Jumlah Harga Jualan (RM)
    If Not IsNull(rs!harga_barang_dengan_gst) Then Report38.Sections("Section5").Controls("L8").Caption = Format(rs!harga_barang_dengan_gst, "#,##0.00") 'Jumlah Harga Jualan (RM)
    If Not IsNull(rs!diskaun) Then Report38.Sections("Section5").Controls("L9").Caption = Format(rs!diskaun, "#,##0.00") 'Jumlah Diskaun (%)
    If Not IsNull(rs!harga_lepas_diskaun) Then Report38.Sections("Section5").Controls("L10").Caption = Format(rs!harga_lepas_diskaun, "#,##0.00") 'Harga Selepas Diskaun (RM)
    If Not IsNull(rs!adjustment) Then Report38.Sections("Section5").Controls("L11").Caption = Format(rs!adjustment, "#,##0.00") 'Adjustment (RM)
    If Not IsNull(rs!harga_jualan) Then Report38.Sections("Section5").Controls("L12").Caption = Format(rs!harga_jualan, "#,##0.00") 'Jumlah Harga Jualan (RM)
    If Not IsNull(rs!jumlah_trade_in) Then
        Report38.Sections("Section5").Controls("L13").Caption = Format(rs!jumlah_trade_in, "#,##0.00") 'Jumlah Resit Trade In (RM)
    Else
        Report38.Sections("Section5").Controls("L13").Caption = "0.00" 'Jumlah Resit Trade In (RM)
    End If
    If Not IsNull(rs!flag_trade_in) Then
        If rs!flag_trade_in = 0 Then
            Report38.Sections("Section5").Controls("L14").Caption = "-" 'No. Resit Trade In
        Else
            If Not IsNull(rs!no_resit_trade_in) Then Report38.Sections("Section5").Controls("L14").Caption = rs!no_resit_trade_in 'No. Resit Trade In
        End If
    End If
    If Not IsNull(rs!flag_bayaran) Then
        If rs!flag_bayaran = 0 Then
            Report38.Sections("Section5").Controls("L15").Caption = "Jumlah Bayaran"
        Else
            Report38.Sections("Section5").Controls("L15").Caption = "Lebihan Kedai Perlu Bayar"
        End If
    End If
    If Not IsNull(rs!jumlah_perlu_bayar) Then
        If IsNumeric(rs!jumlah_perlu_bayar) Then LM_JUMLAH_BAYAR = rs!jumlah_perlu_bayar
    End If
    If Not IsNull(rs!jumlah_cas_kad_kredit) Then
        If IsNumeric(rs!jumlah_cas_kad_kredit) Then LM_CAJ_KAD = rs!jumlah_cas_kad_kredit
    End If
    If Not IsNull(rs!gst_kad_kredit) Then
        If IsNumeric(rs!gst_kad_kredit) Then LM_CAJ_GST = rs!gst_kad_kredit
    End If

    Report38.Sections("Section5").Controls("L16").Caption = Format(LM_JUMLAH_BAYAR + LM_CAJ_KAD + LM_CAJ_GST, "#,##0.00") 'Jumlah Bayaran Yang Perlu Dibuat (RM)
    
    If Not IsNull(rs!kuantiti_barang) Then
        Report38.Sections("Section5").Controls("L18").Caption = rs!kuantiti_barang 'Kuantiti Barang Yang Dijual
        LM_QTY = rs!kuantiti_barang
    End If
    
    If Not IsNull(rs!JUMLAH_BERAT) Then Report38.Sections("Section5").Controls("L19").Caption = Format(rs!JUMLAH_BERAT, "#,##0.00 g") 'Jumlah Berat Barang Yang Dijual
    If Not IsNull(rs!gst_sr_harga) Then Report38.Sections("Section5").Controls("L20").Caption = Format(rs!gst_sr_harga, "#,##0.00") 'Harga Keseluruhan Bagi Barang SR
    If Not IsNull(rs!gst_sr_cukai) Then Report38.Sections("Section5").Controls("L21").Caption = Format(rs!gst_sr_cukai, "#,##0.00") 'Jumlah Cukai Bagi SR
    If Not IsNull(rs!gst_zr_harga) Then Report38.Sections("Section5").Controls("L22").Caption = Format(rs!gst_zr_harga, "#,##0.00") 'Harga Keseluruhan Bagi Barang ZR
    If Not IsNull(rs!gst_zr_cukai) Then Report38.Sections("Section5").Controls("L23").Caption = Format(rs!gst_zr_cukai, "#,##0.00") 'Jumlah Cukai Bagi ZR
    If Not IsNull(rs!caj_pos) Then Report38.Sections("Section5").Controls("L24").Caption = Format(rs!caj_pos, "#,##0.00") '
    If Not IsNull(rs!no_tracking) Then
        Report38.Sections("Section5").Controls("L25").Caption = rs!no_tracking
    Else
        Report38.Sections("Section5").Controls("L25").Caption = "-"
    End If
    If Not IsNull(rs!point_ari_nashi) Then
        If rs!point_ari_nashi = 0 And Not IsNull(rs!jumlah_point) Then
            Report38.Sections("Section5").Controls("L28").Caption = "Anda telah terlepas mata sebanyak " & rs!jumlah_point & " daripada pembelian ini kerana tidak mempunyai kad keahlian kedai." & vbCrLf & _
                                                                    "Sila daftar keahlian dengan kedai untuk mendapat mata ganjaran bagi pembelian berikutnya." 'Caption berkenaan mata ganjaran
            Report38.Sections("Section5").Controls("L28").Visible = True 'Caption berkenaan mata ganjaran
            Report38.Sections("Section5").Controls("L32").Visible = True 'Caption berkenaan mata ganjaran terkumpul
        End If
    End If
    If Not IsNull(rs!kupon_diskaun) Then Report38.Sections("Section5").Controls("L33").Caption = Format(rs!kupon_diskaun, "#,##0.00") 'Kupon Diskaun
    
    If Not IsNull(rs!no_pekerja) Then
        Frm84_LM_No_PEKERJA = rs!no_pekerja
        Frm84_DATA_PEKERJA_FOUND = 1
    End If

'1:  Pelanggan
'2:  Member
'3:  RAF
'4:  Pengedar
'5:  Normal Dealer
'6:  Master Dealer

    If Not IsNull(rs!kategori_pembeli) Then 'Kategori Pembeli
        Frm84_LM_KATEGORI = rs!kategori_pembeli
    End If
    If Not IsNull(rs!no_rujukan_pembeli) Then
        Frm84_LM_No_CUST = rs!no_rujukan_pembeli
        Frm84_DATA_CUST_FOUND = 1 '0 : Pelanggan Tidak Berdaftar , 1 : Pelanggan Berdaftar , 2 : Ahli
    End If
    DATA_FOUND = 1
End If

rs.Close
Set rs = Nothing

LM_JENIS_INVOICE = 0 '0 : Invoice tiada limit , 1 : Invoice dengan limit kuantiti

If LM_INV_CHECK = 1 Then

    If LM_QTY <= LM_QTY_CHECK Then
        
        Report38.Sections("Section1").Controls("label36").Visible = False
        Report38.Sections("Section1").Controls("Text6").Left = 135
        Report38.Sections("Section1").Controls("Text6").Width = 1800
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "TRUNCATE TABLE " & G_INVOICE_TEMP & ""
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "insert into " & G_INVOICE_TEMP & "(no_siri_produk,kategori_produk,purity,berat_jualan,harga_semasa,upah,harga_jualan_dengan_gst,gst_ari_nashi)" & _
                    "select no_siri_produk,kategori_produk,purity,berat_jualan,harga_semasa,upah,harga_jualan_dengan_gst,gst_ari_nashi from 23_senarai_jualan WHERE no_resit='" & G_No_RESIT_JUALAN & "' AND status_rekod = 1"
                    
        Set rs = cn.Execute(strsql)
        Set rs = Nothing

        For Z = 1 To LM_QTY_CHECK
            
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from " & G_INVOICE_TEMP & " where ID='" & Z & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                If Not IsNull(rs!no_siri_Produk) Then
                
                    rs!no_siri_Produk = "No. Siri :  " & rs!no_siri_Produk
                    rs.Update
                    
                End If
                
            Else
            
                rs.AddNew
                rs.Update
        
            End If
            
        Next Z
        
        LM_JENIS_INVOICE = 1 '0 : Invoice tiada limit , 1 : Invoice dengan limit kuantiti
        
    End If
    
End If

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 71_tebus_agih_point where no_invoice='" & G_No_RESIT_JUALAN & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Not IsNull(rs!jumlah_tebus_point) And Not IsNull(rs!nilaian_tebus_point) Then
        Report38.Sections("Section5").Controls("L29").Caption = rs!jumlah_tebus_point & " @ RM " & Format(rs!nilaian_tebus_point, "#,##0.00")
    End If
    If Not IsNull(rs!jumlah_peroleh_point) Then
        Report38.Sections("Section5").Controls("L30").Caption = ": " & rs!jumlah_peroleh_point
    End If

End If

rs.Close
Set rs = Nothing

If DATA_FOUND = 1 Then
    If Frm84_DATA_PEKERJA_FOUND = 1 Then
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where NoPekerja='" & Frm84_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!Samaran) Then Report38.Sections("Section2").Controls("L17").Caption = rs!Samaran 'Nama Samaran
        End If
        
        rs.Close
        Set rs = Nothing
    End If
    
'### Data jika pembeli TIDAK berdaftar ### - Start
    If Frm84_DATA_CUST_FOUND = 0 Then '0 : Pelanggan Tidak Berdaftar , 1 : Pelanggan Berdaftar , 2 : Ahli
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 44_senarai_pelanggan where no_resit='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!Nama) Then Report38.Sections("Section2").Controls("L5").Caption = rs!Nama 'Maklumat Pembeli : Nama
            If Not IsNull(rs!no_tel) Then Report38.Sections("Section2").Controls("L7").Caption = rs!no_tel 'Maklumat Pembeli : No. Telefon
        End If
        
        rs.Close
        Set rs = Nothing
    End If
'### Data jika pembeli TIDAK berdaftar ### - End

'### Data jika pembeli adalah berdaftar ### - Start
    If Frm84_DATA_CUST_FOUND = 1 Then '0 : Pelanggan Tidak Berdaftar , 1 : Pelanggan Berdaftar , 2 : Ahli
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm84_LM_No_CUST & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!Nama) Then Report38.Sections("Section2").Controls("L5").Caption = rs!Nama 'Maklumat Pembeli : Nama
            If Not IsNull(rs!no_tel) Then Report38.Sections("Section2").Controls("L7").Caption = rs!no_tel 'Maklumat Pembeli : No. Telefon
            If Not IsNull(rs!baki_point) Then
                Report38.Sections("Section5").Controls("L32").Caption = ": " & rs!baki_point 'Jumlah mata ganjaran terkumpul
                Report38.Sections("Section5").Controls("L32").Visible = True
            End If
            
            If Not IsNull(rs!no_pelanggan) Then
                If Not IsNull(rs!kategori_pelanggan) Then
                    If rs!kategori_pelanggan = 1 Then
                        Report38.Sections("Section2").Controls("L34").Caption = rs!no_pelanggan & " (Pelanggan Biasa)" 'No. Keahlian
                    ElseIf rs!kategori_pelanggan = 2 Then
                        Report38.Sections("Section2").Controls("L34").Caption = rs!no_pelanggan & " (Ahli Biasa)" 'No. Keahlian
                    ElseIf rs!kategori_pelanggan = 3 Then
                        Report38.Sections("Section2").Controls("L34").Caption = rs!no_pelanggan & " (Silver)" 'No. Keahlian
                    ElseIf rs!kategori_pelanggan = 4 Then
                        Report38.Sections("Section2").Controls("L34").Caption = rs!no_pelanggan & " (Gold)" 'No. Keahlian
                    ElseIf rs!kategori_pelanggan = 5 Then
                        Report38.Sections("Section2").Controls("L34").Caption = rs!no_pelanggan & " (Platinum)" 'No. Keahlian
                    End If
                    Report38.Sections("Section2").Controls("L34").Visible = True 'No. Keahlian (Caption)
                    Report38.Sections("Section2").Controls("L35").Visible = True 'No. Keahlian (Caption)
                End If
            End If
            
        End If
        
        rs.Close
        Set rs = Nothing
    End If
'### Data jika pembeli adalah berdaftar ### - End

    If G_MODE = "YES" Then
        Report38.Sections("Section5").Controls("L28").Visible = True
        Report38.Sections("Section5").Controls("L30").Visible = True
        Report38.Sections("Section5").Controls("L32").Visible = True
        Report38.Sections("Section5").Controls("L36").Visible = True
        Report38.Sections("Section5").Controls("L37").Visible = True
    Else
        Report38.Sections("Section5").Controls("L28").Visible = False
        Report38.Sections("Section5").Controls("L30").Visible = False
        Report38.Sections("Section5").Controls("L32").Visible = False
        Report38.Sections("Section5").Controls("L36").Visible = False
        Report38.Sections("Section5").Controls("L37").Visible = False
    End If

    '### Paparan Resit ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    If LM_JENIS_INVOICE = 0 Then '0 : Invoice tiada limit , 1 : Invoice dengan limit kuantiti
        
        rs.Open "select * from 23_senarai_jualan where no_resit='" & G_No_RESIT_JUALAN & "' AND status_rekod = 1", cn, adOpenKeyset, adLockOptimistic
    
    ElseIf LM_JENIS_INVOICE = 1 Then '0 : Invoice tiada limit , 1 : Invoice dengan limit kuantiti
        
        rs.Open "select * from " & G_INVOICE_TEMP & "", cn, adOpenKeyset, adLockOptimistic
    
    End If
    
    While rs.EOF = False
        Set Report38.DataSource = rs
        If G_PREVIEW = 1 Then Report38.Show
        rs.MoveNext
    Wend
    
    'rs.Close
    Set rs = Nothing
    '### Paparan Resit ### - End
    
    If G_PREVIEW = 0 Then Report38.PrintReport
     
    G_No_RESIT_JUALAN = vbNullString
End If
End Sub
Sub Frm84_cetak_invoice_rms()
'on error resume next
Dim Frm84_LM_TOTAL_BERAT As Double
Dim rs1 As ADODB.Recordset

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

x = 0
DATA_FOUND = 0
Frm84_DATA_PEKERJA_FOUND = 0
Frm84_DATA_CUST_FOUND = 0 '0 : Pelanggan Tidak Berdaftar , 1 : Pelanggan Berdaftar , 2 : Ahli

'### Reset maklumat kedai ### - Start
Report54.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report54.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report54.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report54.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report54.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report54.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report54.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report54.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report54.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report54.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

'### Reset Invoice ### - Start
Report54.Sections("Section4").Controls("L1").Caption = vbNullString 'Maklumat Pembeli : Nama
Report54.Sections("Section4").Controls("L2").Caption = vbNullString 'Maklumat Pembeli : No. Kad Pengenalan
Report54.Sections("Section4").Controls("L3").Caption = vbNullString 'Maklumat Pembeli : No. Telefon
Report54.Sections("Section4").Controls("L100").Caption = "-----------" 'No. ID GST
Report54.Sections("Section4").Controls("L5").Caption = vbNullString 'No. Resit
Report54.Sections("Section4").Controls("L6").Caption = vbNullString 'Tarikh Masa
Report54.Sections("Section4").Controls("L7").Caption = vbNullString 'Jurujual

Report54.Sections("Section5").Controls("L8").Caption = "RM 0.00" 'Total Sales
Report54.Sections("Section5").Controls("L9").Caption = "1" 'Bilangan Barang
Report54.Sections("Section5").Controls("L10").Caption = "RM 0.00" 'Amount bayaran yang dikenakan GST
Report54.Sections("Section5").Controls("L11").Caption = "RM 0.00" 'GST
'### Reset Invoice ### - End

'G_No_RESIT_JUALAN = "BK000021"

Report54.Sections("Section4").Controls("L5").Caption = G_No_RESIT_JUALAN 'No. Resit Jualan

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where no_resit='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!write_timestamp) Then Report54.Sections("Section4").Controls("L6").Caption = rs!write_timestamp 'Jumlah Harga Jualan (RM)
    If Not IsNull(rs!harga_jualan) Then Report54.Sections("Section5").Controls("L8").Caption = "RM " & Format(rs!harga_jualan, "#,##0.00") 'Total Sales
    If Not IsNull(rs!kuantiti_barang) Then Report54.Sections("Section5").Controls("L9").Caption = rs!kuantiti_barang 'Kuantiti Barang Yang Dijual
    If Not IsNull(rs!harga_jualan) Then Report54.Sections("Section5").Controls("L10").Caption = "RM " & Format(rs!harga_jualan, "#,##0.00") 'Harga Keseluruhan Bagi Barang SR
    If Not IsNull(rs!gst_sr_cukai) Then Report54.Sections("Section5").Controls("L11").Caption = "RM " & Format(rs!gst_sr_cukai, "#,##0.00") 'Jumlah Cukai Bagi SR
    
    If Not IsNull(rs!no_pekerja) Then
        Frm84_LM_No_PEKERJA = rs!no_pekerja
        Frm84_DATA_PEKERJA_FOUND = 1
    End If
    If Not IsNull(rs!no_rujukan_pembeli) Then
        Frm84_LM_No_CUST = rs!no_rujukan_pembeli
        Frm84_DATA_CUST_FOUND = 1 '0 : Pelanggan Tidak Berdaftar , 1 : Pelanggan Berdaftar , 2 : Ahli
    End If
    DATA_FOUND = 1
End If

rs.Close
Set rs = Nothing

If DATA_FOUND = 1 Then

'### Padam Table report3 #### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    
    strsql = "TRUNCATE TABLE report3"
    
    Set rs = cn.Execute(strsql)
    Set rs = Nothing
'### Padam Table report3 #### - End

    '### Paparan Senarai Jualan ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 23_senarai_jualan where no_resit='" & G_No_RESIT_JUALAN & "' AND status_rekod = 1", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        
        x = x + 1
        
        Set rs1 = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs1.Open "select * from report3", cn, adOpenKeyset, adLockOptimistic
        
        rs1.AddNew
        rs1!no = x 'No.
        If Not IsNull(rs!no_siri_Produk) Then rs1!no_siri_Produk = rs!no_siri_Produk 'No. Siri Produk
        If Not IsNull(rs!kategori_Produk) Then rs1!kategori_Produk = rs!kategori_Produk 'Kategori Produk
        If Not IsNull(rs!purity) Then rs1!purity = rs!purity 'Purity
        If Not IsNull(rs!berat_jualan) Then rs1!berat_jualan = rs!berat_jualan 'Berat Jualan
        If Not IsNull(rs!harga_Semasa) Then rs1!harga_Semasa = rs!harga_Semasa 'Harga Semasa
        If Not IsNull(rs!UPAH) Then rs1!UPAH = rs!UPAH 'Upah
        If Not IsNull(rs!harga_dengan_gst) Then rs1!harga_jualan = rs!harga_dengan_gst 'Harga Jualan
        rs1.Update
        
        rs1.Close
        Set rs1 = Nothing

        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    '### Paparan Senarai Jualan ### - End
    
    '### Limitkan paparan kepada 5 data sahaja ### - Start
    If x < 5 Then
        For Y = 1 To 5 - x
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from report3", cn, adOpenKeyset, adLockOptimistic
            
            rs.AddNew
            rs!no = x + 1
            rs.Update
            
            rs.Close
            Set rs = Nothing
        Next Y
    End If
    '### Limitkan paparan kepada 5 data sahaja ### - End

    If Frm84_DATA_PEKERJA_FOUND = 1 Then
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where NoPekerja='" & Frm84_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!Samaran) Then Report54.Sections("Section4").Controls("L7").Caption = rs!Samaran 'Nama Samaran
        End If
        
        rs.Close
        Set rs = Nothing
    End If
    
'### Data jika pembeli TIDAK berdaftar ### - Start
    If Frm84_DATA_CUST_FOUND = 0 Then '0 : Pelanggan Tidak Berdaftar , 1 : Pelanggan Berdaftar , 2 : Ahli
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 44_senarai_pelanggan where no_resit='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!Nama) Then Report54.Sections("Section4").Controls("L1").Caption = rs!Nama 'Maklumat Pembeli : Nama
            If Not IsNull(rs!no_tel) Then Report54.Sections("Section4").Controls("L3").Caption = rs!no_tel 'Maklumat Pembeli : No. Telefon
        End If
        
        rs.Close
        Set rs = Nothing
    End If
'### Data jika pembeli TIDAK berdaftar ### - End

'### Data jika pembeli adalah berdaftar ### - Start
    If Frm84_DATA_CUST_FOUND = 1 Then '0 : Pelanggan Tidak Berdaftar , 1 : Pelanggan Berdaftar , 2 : Ahli
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm84_LM_No_CUST & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!Nama) Then Report54.Sections("Section4").Controls("L1").Caption = rs!Nama 'Maklumat Pembeli : Nama
            If Not IsNull(rs!no_ic) Then Report54.Sections("Section4").Controls("L2").Caption = rs!no_ic 'Maklumat Pembeli : No. Kad Pengenalan
            If Not IsNull(rs!no_tel) Then Report54.Sections("Section4").Controls("L3").Caption = rs!no_tel 'Maklumat Pembeli : No. Telefon
        End If
        
        rs.Close
        Set rs = Nothing
    End If
'### Data jika pembeli adalah berdaftar ### - End

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If rs!Default1 = "Default" Then
            If Not IsNull(rs!id_gst) Then
                Report54.Sections("Section4").Controls("L100").Caption = rs!id_gst 'No. ID GST
            End If
        End If
    End If
    
    rs.Close
    Set rs = Nothing

    '### Paparan Resit ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from report3", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        Set Report54.DataSource = rs
        Report54.Show
        rs.MoveNext
    Wend
    
    'rs.Close
    Set rs = Nothing
    '### Paparan Resit ### - End
    
    G_No_RESIT_JUALAN = vbNullString
End If
End Sub
Sub Frm84_Resit_Buyback()
'on error resume next
Dim Frm84_LM_TOTAL_BERAT As Double

DATA_FOUND = 0
Frm84_DATA_PEKERJA_FOUND = 0
Frm84_DATA_CUST_FOUND = 0
Frm84_LM_TOTAL_BERAT = 0
LM_KATEGORI_PENJUAL = 0 '0 : Tidak Berdaftar & Berdaftar , 1 : Ahli
x = 0
Frm84_LM_No_CUST = vbNullString

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
    If Not IsNull(rs!default_printer) Then LM_PRINTER = rs!default_printer
    
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

'### Reset Maklumat Penjual #### - Start
Report39.Sections("Section2").Controls("L5").Caption = vbNullString 'Maklumat Pembeli : Nama
Report39.Sections("Section2").Controls("L7").Caption = vbNullString 'Maklumat Pembeli : No. Telefon
'### Reset Maklumat Penjual #### - End

Report39.Sections("Section5").Controls("L12").Caption = "0.00" 'Tunai
Report39.Sections("Section5").Controls("L13").Caption = "0.00" 'Bank In

'### Reset maklumat kedai ### - Start
Report39.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report39.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report39.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report39.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report39.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

If LM_HEADER = 1 Then '0 : Pre Printed , 1 : Sistem
    '### Maklumat kedai ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!nama_kedai) Then Report39.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
        If Not IsNull(rs!no_pendaftaran) Then Report39.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
        If Not IsNull(rs!alamat) Then Report39.Sections("Section4").Controls("L202").Caption = rs!alamat
        If Not IsNull(rs!no_tel) Then Report39.Sections("Section4").Controls("L203").Caption = rs!no_tel
        If Not IsNull(rs!no_id_gst) Then Report39.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
        
        If Not IsNull(rs!gst_ari_nashi) Then
        
            If rs!gst_ari_nashi = 0 Then

                Report39.Sections("Section5").Controls("Shape1").Visible = False
                Report39.Sections("Section5").Controls("Label53").Visible = False
                Report39.Sections("Section5").Controls("Label55").Visible = False
                Report39.Sections("Section5").Controls("Label56").Visible = False
                Report39.Sections("Section5").Controls("Label61").Visible = False
                Report39.Sections("Section5").Controls("Label62").Visible = False
                Report39.Sections("Section5").Controls("L20").Visible = False
                Report39.Sections("Section5").Controls("L21").Visible = False
                Report39.Sections("Section5").Controls("L22").Visible = False
                Report39.Sections("Section5").Controls("L23").Visible = False
                Report39.Sections("Section5").Controls("Line2").Visible = False
                
            End If
            
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    '### Maklumat kedai ### - End
    
    Report39.Sections("Section4").Visible = True
Else
    Report39.Sections("Section4").Visible = False
End If

Report39.Sections("Section2").Controls("L3").Caption = G_No_RESIT_JUALAN 'No. Resit Jualan

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!tarikh) Then Report39.Sections("Section2").Controls("L4").Caption = rs!tarikh 'Tarikh Belian
    If Not IsNull(rs!jumlah_dengan_gst) Then Report39.Sections("Section5").Controls("L9").Caption = "RM " & Format(rs!jumlah_dengan_gst, "#,##0.00") 'Jumlah Belian Dengan GST (RM)
    'If Not IsNull(rs!JUMLAH_BERAT) Then Report39.Sections("Section5").Controls("L19").Caption = rs!JUMLAH_BERAT 'Jumlah Berat Barang Yang Dijual
    If Not IsNull(rs!gst_sr_harga) Then Report39.Sections("Section5").Controls("L20").Caption = Format(rs!gst_sr_harga, "#,##0.00") 'Harga Keseluruhan Bagi Barang SR
    If Not IsNull(rs!gst_sr_cukai) Then Report39.Sections("Section5").Controls("L21").Caption = Format(rs!gst_sr_cukai, "#,##0.00") 'Jumlah Cukai Bagi SR
    If Not IsNull(rs!gst_zr_harga) Then Report39.Sections("Section5").Controls("L22").Caption = Format(rs!gst_zr_harga, "#,##0.00") 'Harga Keseluruhan Bagi Barang ZR
    If Not IsNull(rs!gst_zr_cukai) Then Report39.Sections("Section5").Controls("L23").Caption = Format(rs!gst_zr_cukai, "#,##0.00") 'Jumlah Cukai Bagi ZR
    If Not IsNull(rs!tunai) Then Report39.Sections("Section5").Controls("L12").Caption = Format(rs!tunai, "#,##0.00")  'Tunai
    If Not IsNull(rs!bank_in) Then Report39.Sections("Section5").Controls("L13").Caption = Format(rs!bank_in, "#,##0.00") 'Bank In
    If Not IsNull(rs!no_pekerja) Then
        Frm84_LM_No_PEKERJA = rs!no_pekerja
        Frm84_DATA_PEKERJA_FOUND = 1
    End If
    If Not IsNull(rs!no_rujukan_pelanggan_buyback) Then Frm84_LM_No_CUST = rs!no_rujukan_pelanggan_buyback
    DATA_FOUND = 1
End If

rs.Close
Set rs = Nothing

If DATA_FOUND = 1 Then
    If Frm84_DATA_PEKERJA_FOUND = 1 Then
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where NoPekerja='" & Frm84_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!Samaran) Then Report39.Sections("Section2").Controls("L8").Caption = rs!Samaran 'Nama Samaran
        End If
        
        rs.Close
        Set rs = Nothing
    End If
        
    If Frm84_LM_No_CUST = vbNullString Then
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 44_senarai_pelanggan where no_resit='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!Nama) Then Report39.Sections("Section2").Controls("L5").Caption = rs!Nama 'Maklumat Pembeli : Nama
            If Not IsNull(rs!no_tel) Then Report39.Sections("Section2").Controls("L7").Caption = rs!no_tel 'Maklumat Pembeli : No. Telefon
        End If
        
        rs.Close
        Set rs = Nothing
        
    ElseIf Frm84_LM_No_CUST <> vbNullString Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm84_LM_No_CUST & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!Nama) Then Report39.Sections("Section2").Controls("L5").Caption = rs!Nama 'Maklumat Pembeli : Nama
            If Not IsNull(rs!no_tel) Then Report39.Sections("Section2").Controls("L7").Caption = rs!no_tel 'Maklumat Pembeli : No. Telefon
        End If
        
        rs.Close
        Set rs = Nothing
        
    End If

    '### Paparan Resit ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from Data_Database where bill_No_Trade_In='" & G_No_RESIT_JUALAN & "' AND StatusItem <> 0", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        x = x + 1
        If Not IsNull(rs!Berat) Then
            If IsNumeric(rs!Berat) Then Frm84_LM_TOTAL_BERAT = Frm84_LM_TOTAL_BERAT + rs!Berat
        End If
        Set Report39.DataSource = rs
        Report39.Show
        rs.MoveNext
    Wend
    
    'rs.Close
    Set rs = Nothing
    '### Paparan Resit ### - End
    
    Report39.Sections("Section5").Controls("L10").Caption = x 'Bilangan Barang
    Report39.Sections("Section5").Controls("L11").Caption = Format(Frm84_LM_TOTAL_BERAT, "#,##0.00 g") 'Jumlah Berat (g)
    
    G_No_RESIT_JUALAN = vbNullString
End If
End Sub
Sub Frm84_auto_insert_data()
'On Error Resume Next
Dim Err(30)
Dim Frm84_LM_BERAT_ASAL As Double
Dim Frm84_LM_BERAT_JUAL As Double
Dim Frm84_LM_HARGA_MODAL As Double
Dim Frm84_LM_HARGA_JUAL As Double
Dim Frm84_LM_HARGA_SEMASA_MODAL As Double
Dim Frm84_LM_TETAPANHARGA As Double
Dim Frm84_LM_LIMIT As Double
Dim Frm84_LM_HARGA_STAFF As Double 'Tetapan harga jualan kepada staff
Dim Frm84_LM_HARGA_PELANGGAN As Double 'Tetapan harga jualan kepada pelanggan
Dim Frm84_LM_HARGA_SEMASA As Double 'Harga semasa (jualan)
Dim Frm84_LM_HARGA_SUPPLIER As Double 'Harga per gram (harga semasa) dari supplier (modal)
Dim Frm84_UPAH_MODAL As Double 'Upah modal
Dim Frm84_UPAH_JUAL As Double 'Upah jualan
Dim Frm84_LM_HARGA_JUALAN_CALC As Double 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
Dim Frm84_LM_GST_CALC As Double 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst

Frm84_LM_HARGA_JUALAN_CALC = 0 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
Frm84_LM_GST_CALC = 0 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
Frm84_LM_HARGA_SEMASA = 0 'Harga semasa (jualan)
Frm84_LM_HARGA_SUPPLIER = 0 'Harga per gram (harga semasa) dari supplier (modal)
x = 0
Frm84_LM_BERAT_ASAL = 0
Frm84_LM_BERAT_JUAL = 0
Frm84_LM_DATA_SAVE = 0
Frm84_LM_HARGA_MODAL = 0
Frm84_LM_HARGA_JUAL = 0
Frm84_LM_HARGA_SEMASA_MODAL = 0
Frm84_LM_PRICE_CHECK = 0 '0 : Tidak Perlu Periksa Harga Semasa Jualan , 1 : Perlu Periksa Harga Semasa Jualan
Frm84_LM_TETAPANHARGA = 0
Frm84_LM_LIMIT = 0
Frm84_LM_HARGA_STAFF = 0
Frm84_LM_HARGA_PELANGGAN = 0
Frm84_UPAH_MODAL = 0 'Upah modal
Frm84_UPAH_JUAL = 0 'Upah jualan

If Frm84.TB3 = vbNullString And Frm84.CB12 = 1 Then
    MsgBox "Tetapan GST ke atas UPAH hanya dibenarkan untuk barang kemas SAHAJA. Sila periksa tetapan GST anda.", vbExclamation, "Info"
    Exit Sub
End If

If Frm84.TB2 = vbNullString Then
    x = x + 1
    Err(x) = "Sila Masukkan [No. Siri Produk]."
End If
If (Frm84.TB14 <> vbNullString And IsNumeric(Frm84.TB14)) And (Frm84.L51_Text <> vbNullString And IsNumeric(Frm84.L51_Text)) Then
    Frm84_LM_HARGA_STAFF = Frm84.L51_Text
    Frm84_LM_HARGA_PELANGGAN = Frm84.TB14
    
    If Frm84_LM_HARGA_PELANGGAN < Frm84_LM_HARGA_STAFF Then
        x = x + 1
        Err(x) = "Harga Jualan Minimum Yang Dibenarkan Adalah RM " & Format(Frm84_LM_HARGA_STAFF, "#,##0.00")
    End If
End If

'### Error Bagi Item BK ### - Start
If Frm84.TB3 <> vbNullString Then
    If Frm84.TB3 = vbNullString Or (Frm84.TB3 <> vbNullString And Not IsNumeric(Frm84.TB3)) Then
        x = x + 1
        Err(x) = "Sila Maklumat [Berat Asal]. Sila Scan Item Sekali Lagi."
    End If
    If Frm84.TB4 = vbNullString Or (Frm84.TB4 <> vbNullString And Not IsNumeric(Frm84.TB4)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Berat Jualan]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
    End If
    If Frm84.TB5 = vbNullString Or (Frm84.TB5 <> vbNullString And Not IsNumeric(Frm84.TB5)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Harga Semasa]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
    End If
    If Frm84.TB15 = vbNullString Or (Frm84.TB15 <> vbNullString And Not IsNumeric(Frm84.TB15)) Then
        x = x + 1
        Err(x) = "Sila Masukkan [Upah]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
    End If
    If Frm84.CB7 = 1 Then
        If Frm84.TB12 = vbNullString Or (Frm84.TB12 <> vbNullString And Not IsNumeric(Frm84.TB12)) Then
            x = x + 1
            Err(x) = "Sila Masukkan [Komisen Per Gram]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
        End If
    End If
End If
'### Error Bagi Item BK ### - End

'### Error Bagi Item Permata ### - Start
If Frm84.TB3 = vbNullString Then
    If Frm84.CB7 = 1 Then
        If Frm84.TB16 = vbNullString Or (Frm84.TB16 <> vbNullString And Not IsNumeric(Frm84.TB16)) Then
            x = x + 1
            Err(x) = "Sila Masukkan [Jumlah Komisen]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
        End If
    End If
End If
'### Error Bagi Item Permata ### - End

If Frm84.TB7 = vbNullString Or (Frm84.TB7 <> vbNullString And Not IsNumeric(Frm84.TB7)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Diskaun]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
End If
If Frm84.TB9 = vbNullString Or (Frm84.TB9 <> vbNullString And Not IsNumeric(Frm84.TB9)) Then
    x = x + 1
    Err(x) = "Sila Masukkan [Adjustment]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
End If
If Frm84.TB10 = vbNullString Or (Frm84.TB10 <> vbNullString And Not IsNumeric(Frm84.TB10)) Then
    x = x + 1
    Err(x) = "Tiada maklumat [Harga Jualan]. Sila keluarkan item ini dari senarai dan scan sekali lagi."
End If
If Frm84.TB11 = vbNullString Or (Frm84.TB11 <> vbNullString And Not IsNumeric(Frm84.TB11)) Then
    x = x + 1
    Err(x) = "Tiada maklumat [Jumlah GST]. Sila keluarkan item ini dari senarai dan scan sekali lagi."
End If
If Frm84.CB4 = 0 And Frm84.CB5 = 0 And Frm84.CB6 = 0 And Frm84.CB9 = 0 And Frm84.CB10 = 0 Then
    x = x + 1
    Err(x) = "Sila Buat Pilihan Kategori Pembeli."
End If
If Frm84.CB2 = 0 And Frm84.CB3 = 0 Then
    x = x + 1
    Err(x) = "Sila Buat Pilihan Jenis GST."
End If
If (Frm84.TB3 <> vbNullString And IsNumeric(Frm84.TB3)) And (Frm84.TB4 <> vbNullString And IsNumeric(Frm84.TB4)) Then
    Frm84_LM_BERAT_ASAL = Frm84.TB3 'Berat Asal
    Frm84_LM_BERAT_JUAL = Frm84.TB4 'Berat Jualan
    
    If Frm84_LM_BERAT_JUAL > Frm84_LM_BERAT_ASAL Then
        x = x + 1
        Err(x) = "Berat Jualan Melebihi Berat Asal."
    End If
End If
If Frm84.TB3 <> vbNullString And Frm84.L54_Text = vbNullString Then
    x = x + 1
    Err(x) = "Tiada maklumat harga dari supplier bagi purity ini." & vbCrLf & _
                "Sila pastikan harga dari supplier bagi purity ini telah ditetapkan dalam TETAPAN HARIAN SISTEM."
End If

If x = 0 And Frm84.CB7 = 0 Then
'### Periksa Kadar Penurunan Harga ### - Start

    user = MDI_frm1.L3_Text
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from tblelogin where username='" & user & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!usertype) Then
            If rs!usertype = "Guest" Then
                Frm84_LM_PRICE_CHECK = 1 '0 : Tidak Perlu Periksa Harga Semasa Jualan , 1 : Perlu Periksa Harga Semasa Jualan
            End If
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    
'### Periksa Data Dulang ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from Data_Database where no_siri_Produk='" & Frm84.TB2 & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!dulang) Then Frm84_LM_DULANG = rs!dulang 'Dulang
    End If
    
    rs.Close
    Set rs = Nothing
'### Periksa Data Dulang ### - End
    
    If Frm84_LM_PRICE_CHECK = 1 Then '0 : Tidak Perlu Periksa Harga Semasa Jualan , 1 : Perlu Periksa Harga Semasa Jualan
        Frm84_LM_LIMIT_TYPE = 0 '1 : BK , 2 : Barang Permata
        
'### Periksa Purity Dan Tetapan Harga Jualan ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Data_Database where no_siri_Produk='" & Frm84.TB2 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!receiving_Status) Then
                If rs!receiving_Status = 0 Or rs!receiving_Status = 2 Then '0 : BK , 1 : Barang Permata , 2 : Buyback BK , 3 : Buyback Barang Permata
                    If Not IsNull(rs!kod_Purity) Then
                        Frm84_LM_PURITY = rs!kod_Purity 'Purity
                    End If
                    Frm84_LM_LIMIT_TYPE = 1 '1 : BK , 2 : Barang Permata
                End If
                If rs!receiving_Status = 1 Or rs!receiving_Status = 3 Then '0 : BK , 1 : Barang Permata , 2 : Buyback BK , 3 : Buyback Barang Permata
                    If Frm84.CB4 = 1 Then
                        If IsNumeric(rs!code_Supplier) Then Frm84_LM_TETAPANHARGA = Format(rs!code_Supplier, "0.00")  'Harga Pelanggan
                    ElseIf Frm84.CB5 = 1 Then
                        If IsNumeric(rs!HargaJualan_Member) Then Frm84_LM_TETAPANHARGA = Format(rs!HargaJualan_Member, "0.00") 'Harga Member
                    'ElseIf Frm84.CB15 = 1 Then
                    '    If IsNumeric(rs!HargaJualan_RAF) Then Frm84_LM_TETAPANHARGA = Format(rs!HargaJualan_RAF, "0.00") 'Harga RAF
                    ElseIf Frm84.CB6 = 1 Then
                        If IsNumeric(rs!HargaJualan_Pengedar) Then Frm84_LM_TETAPANHARGA = Format(rs!HargaJualan_Pengedar, "0.00") 'Harga Pengedar
                    'ElseIf Frm84.CB13 = 1 Then
                    '    If IsNumeric(rs!HargaJualan_Stokis) Then Frm84_LM_TETAPANHARGA = Format(rs!HargaJualan_Stokis, "0.00") 'Harga Stokis
                    End If
                    Frm84_LM_LIMIT_TYPE = 2 '1 : BK , 2 : Barang Permata
                End If
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        
'### Carian Harga Semasa Emas ### - Start
        If Frm84_LM_LIMIT_TYPE = 1 Then '1 : BK , 2 : Barang Permata
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from default_setting where Default1='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                'If rs!Default1 = "Default" Then
                    If IsNumeric(rs!limit_per_gram) Then Frm84_LM_LIMIT = rs!limit_per_gram
                'End If
            End If
            
            rs.Close
            Set rs = Nothing
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from hargaemas where Purity='" & Frm84_LM_PURITY & "' AND cawangan='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Frm84.CB4 = 1 Then
                    If IsNumeric(rs!Harga_Pelanggan) Then Frm84_LM_TETAPANHARGA = Format(rs!Harga_Pelanggan, "0.00") 'Harga Pelanggan
                ElseIf Frm84.CB5 = 1 Then
                    If IsNumeric(rs!Harga_Member) Then Frm84_LM_TETAPANHARGA = Format(rs!Harga_Member, "0.00") 'Harga Member
                'ElseIf Frm84.CB15 = 1 Then
                '    If IsNumeric(rs!Harga_RAF) Then Frm84_LM_TETAPANHARGA = Format(rs!Harga_RAF, "0.00") 'Harga RAF
                ElseIf Frm84.CB6 = 1 Then
                    If IsNumeric(rs!Harga_Pengedar) Then Frm84_LM_TETAPANHARGA = Format(rs!Harga_Pengedar, "0.00") 'Harga Pengedar
                'ElseIf Frm84.CB13 = 1 Then
                '    If IsNumeric(rs!Harga_Stokis) Then Frm84_LM_TETAPANHARGA = Format(rs!Harga_Stokis, "0.00") 'Harga Stokis
                End If
            End If
            
            rs.Close
            Set rs = Nothing
            
            If IsNumeric(Frm84.TB5) Then
                Frm84_LM_HARGA_JUALAN = Frm84.TB5 'Harga Semasa Jualan (RM/g)
            End If
            
            If Frm84_LM_TETAPANHARGA - Frm84_LM_HARGA_JUALAN > Frm84_LM_LIMIT Then
                MsgBox "Harga jualan tidak mengikut pengurangan harga minimum yang ditetapkan oleh kedai!." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Tetapan Asal Harga Jualan : RM " & Format(Frm84_LM_TETAPANHARGA, "0.00") & vbCrLf & _
                "Limit Diskaun Pengurangan Harga : RM " & Format(Frm84_LM_LIMIT, "0.00") & vbCrLf & _
                vbNullString & vbCrLf & _
                "Sila periksa data anda.", vbExclamation, "Info"
                Exit Sub
            End If
        End If
        
        If Frm84_LM_LIMIT_TYPE = 2 Then '1 : BK , 2 : Barang Permata
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If rs!Default1 = "Default" Then
                    If IsNumeric(rs!limit_per_item) Then Frm84_LM_LIMIT = rs!limit_per_item
                End If
            End If
            
            rs.Close
            Set rs = Nothing
            
            If IsNumeric(Frm84.TB10) Then
                Frm84_LM_HARGA_JUALAN = Frm84.TB10 'Harga Jualan (RM)
            End If
            
            If Frm84_LM_TETAPANHARGA - Frm84_LM_HARGA_JUALAN > Frm84_LM_LIMIT Then
                MsgBox "Harga jualan tidak mengikut pengurangan harga minimum yang ditetapkan oleh kedai!." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Tetapan Asal Harga Jualan : RM " & Format(Frm84_LM_TETAPANHARGA, "0.00") & vbCrLf & _
                "Limit Diskaun Pengurangan Harga : RM " & Format(Frm84_LM_LIMIT, "0.00") & vbCrLf & _
                vbNullString & vbCrLf & _
                "Sila periksa data anda.", vbExclamation, "Info"
                Exit Sub
            End If
        End If
'### Carian Harga Semasa Emas ### - End
    
'### Periksa Purity Dan Tetapan Harga Jualan ### - End
    End If
'### Periksa Kadar Penurunan Harga ### - End

'### Masukkan Data Ke Dalam Temp Table ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from " & G_JUALAN_TEMP & " where no_siri_Produk='" & Frm84.TB2 & "'", cn, adOpenKeyset, adLockOptimistic
    
    If rs.EOF Then
        rs.AddNew
        If Frm84.TB2 <> vbNullString Then
            rs!no_siri_Produk = Frm84.TB2 'No. Siri Produk
        Else
            rs!no_siri_Produk = Null 'No. Siri Produk
        End If
        If Frm84.L12_Text <> vbNullString Then
            rs!kategori_Produk = Frm84.L12_Text 'Kategori Produk
        Else
            rs!kategori_Produk = Null 'Kategori Produk
        End If
        If Frm84.L13_Text <> vbNullString Then
            rs!purity = Frm84.L13_Text 'Purity
        Else
            rs!purity = Null 'Purity
        End If
        If Frm84.TB3 <> vbNullString Then
            rs!Berat_Asal = Format(Frm84.TB3, "0.00") 'Berat Asal (g)
        Else
            rs!Berat_Asal = Null 'Berat Asal (g)
        End If
        If Frm84.TB4 <> vbNullString Then
            rs!berat_jualan = Format(Frm84.TB4, "0.00") 'Berat Jualan (g)
        Else
            rs!berat_jualan = Null 'Berat Jualan (g)
        End If
        If Frm84.TB5 <> vbNullString Then
            rs!harga_Semasa = Format(Frm84.TB5, "0.00") 'Harga Semasa (RM/g)
        Else
            rs!harga_Semasa = Null 'Harga Semasa (RM/g)
        End If
        If Frm84.TB15 <> vbNullString Then
            rs!UPAH = Format(Frm84.TB15, "0.00") 'Upah (RM)
        Else
            rs!UPAH = Null 'Upah (RM)
        End If
        If Frm84.TB6 <> vbNullString Then
            rs!harga_asal = Format(Frm84.TB6, "0.00") 'Harga Asal Item (RM)
        Else
            rs!harga_asal = Null 'Harga Asal Item (RM)
        End If
        If Frm84.TB7 <> vbNullString Then
            rs!diskaun = Format(Frm84.TB7, "0.00") 'Diskaun (%)
        Else
            rs!diskaun = Null 'Diskaun (%)
        End If
        If Frm84.TB8 <> vbNullString Then
            rs!harga_lepas_diskaun = Format(Frm84.TB8, "0.00") 'Harga Selepas Diskaun (RM)
        Else
            rs!harga_lepas_diskaun = Null 'Harga Selepas Diskaun (RM)
        End If
        If Frm84.TB9 <> vbNullString Then
            rs!adjustment = Format(Frm84.TB9, "0.00") 'Adjustment (RM)
        Else
            rs!adjustment = Null 'Adjustment (RM)
        End If
        If Frm84.TB10 <> vbNullString Then
            rs!harga_jualan = Format(Frm84.TB10, "0.00") 'Harga Jualan (RM)
        Else
            rs!harga_jualan = Null 'Harga Jualan (RM)
        End If
        If Frm84.CB2 = 1 Then
            rs!gst_ari_nashi = "ZR (L)" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
            rs!kadar_gst = Null 'Kadar Cukai GST (%)
            If Frm84.TB11 <> vbNullString Then
                rs!jumlah_gst = Format(Frm84.TB11, "0.00") 'Jumlah Cukai GST (RM)
            Else
                rs!jumlah_gst = "0.00" 'Jumlah Cukai GST (RM)
            End If
        ElseIf Frm84.CB3 = 1 Then
            rs!gst_ari_nashi = "SR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
            If Frm84.L8_Text <> vbNullString Then
                rs!kadar_gst = Frm84.L8_Text 'Kadar Cukai GST (%)
            Else
                rs!kadar_gst = "0" 'Jumlah Cukai GST (RM)
            End If
            If Frm84.TB11 <> vbNullString Then
                rs!jumlah_gst = Format(Frm84.TB11, "0.00") 'Jumlah Cukai GST (RM)
            Else
                rs!jumlah_gst = "0.00" 'Jumlah Cukai GST (RM)
            End If
            If Frm84.CB18 = 1 Then 'Jenis Cukai GST SR
                rs!gst_include = 1 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
            Else
                rs!gst_include = 0 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
            End If
        End If
        If Frm84.L44_Text <> vbNullString Then
            rs!harga_tanpa_gst = Format(Frm84.L44_Text, "0.00") 'Harga Jualan Tanpa GST (RM)
        Else
            rs!harga_tanpa_gst = Null 'Harga Jualan Tanpa GST (RM)
        End If
        If Frm84.TB14 <> vbNullString Then
            rs!harga_dengan_gst = Format(Frm84.TB14, "0.00") 'Harga Jualan Termasuk GST (RM)
        Else
            rs!harga_dengan_gst = Null 'Harga Jualan Termasuk GST (RM)
        End If
        If Frm84.CB7 = 1 Then
            rs!dropship = 1 '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
            If Frm84.Frame2.Visible = True Then 'Komisen Agen Dropship : BK
                If Frm84.TB12 <> vbNullString Then
                    rs!komisyen_per_gram = Format(Frm84.TB12, "0.00") 'Komisyen Per Gram Dropship (RM/g) : BK
                Else
                    rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g) : BK
                End If
                If Frm84.TB13 <> vbNullString Then
                    rs!jumlah_komisyen = Format(Frm84.TB13, "0.00") 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
                Else
                    rs!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
                End If
            End If
            If Frm84.Frame3.Visible = True Then 'Komisen Agen Dropship : Permata
                rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g) : BK
                If Frm84.TB16 <> vbNullString Then
                    rs!jumlah_komisyen = Format(Frm84.TB16, "0.00") 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : Permata
                Else
                    rs!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : Permata
                End If
            End If
        End If
            
        If Frm84.CB7 = 0 Then
            rs!dropship = 0 '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
            rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g)
            rs!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini
        End If
        
        If Frm84.L41_Text = "0" Then '0 : Menu Data Baru , 1 : Menu Edit Data
            rs!Status = 1
        ElseIf Frm84.L41_Text = "1" Then
            rs!Status = 3
        End If
        
        If Frm84.TB3 = vbNullString Then
            rs!Type = 1 '0 : BK , 1 : Barang Permata
            rs!harga_per_gram_modal = Null 'Harga Per Gram Bagi Modal (RM/g)
            If Frm84.L34_Text <> vbNullString Then
                rs!modal = Format(Frm84.L34_Text, "0.00") 'Harga Modal (RM)
            Else
                rs!modal = Null 'Harga Modal (RM)
            End If
            If IsNumeric(Frm84.TB14) And IsNumeric(Frm84.L34_Text) Then
                Frm84_LM_HARGA_MODAL = Frm84.L34_Text 'Harga Modal
                Frm84_LM_HARGA_JUAL = Frm84.TB14 'Harga Jualan
                
                rs!untung = Format(Frm84_LM_HARGA_JUAL - Frm84_LM_HARGA_MODAL, "0.00") 'Jumlah Keuntungan
            End If
        
            rs!harga_per_gram_supplier = Null 'Harga per gram (harga semasa) dari supplier (modal)
            rs!untung2 = Null 'Untung jika restok pada harga supplier ini
            
        Else
            rs!Type = 0 '0 : BK , 1 : Barang Permata
            
            If Frm84.L34_Text <> vbNullString Then
                rs!harga_per_gram_modal = Format(Frm84.L34_Text, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
                If IsNumeric(Frm84.L34_Text) Then
                    Frm84_LM_HARGA_SEMASA_MODAL = Frm84.L34_Text 'Harga Per Gram Bagi Modal (RM/g)
                    
                    rs!modal = Format(Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_SEMASA_MODAL, "0.00") 'Harga Modal (RM)
                End If
            Else
                rs!modal = Null 'Harga Modal (RM)
                rs!harga_per_gram_modal = Null 'Harga Per Gram Bagi Modal (RM/g)
                rs!upah_modal = Null 'Upah modal
            End If

            If IsNumeric(Frm84.TB14) And IsNumeric(Frm84.L34_Text) And IsNumeric(Frm84.TB10) And IsNumeric(Frm84.TB11) Then
                
                Frm84_LM_HARGA_MODAL = Frm84.L34_Text 'Harga Modal
                Frm84_LM_HARGA_JUAL = Frm84.TB14 'Harga Jualan
                Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
                Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
                
                If Frm84.CB12 = 0 Then
                
                    rs!untung = Format(Frm84_LM_HARGA_JUAL - (Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_MODAL), "0.00") 'Jumlah Keuntungan
                    
                ElseIf Frm84.CB12 = 1 Then
                    
                    If Frm84.CB3 = 1 And Frm84.CB18 = 0 Then
                        
                        rs!untung = Format((Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC) - (Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_MODAL), "0.00") 'Jumlah Keuntungan
                        
                    Else
                        
                        rs!untung = Format(Frm84_LM_HARGA_JUALAN_CALC - (Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_MODAL), "0.00") 'Jumlah Keuntungan
                        
                    End If
                    
                End If
                
            End If
            
            If IsNumeric(Frm84.TB4) And IsNumeric(Frm84.TB5) And IsNumeric(Frm84.L54_Text) And IsNumeric(Frm84.L55_Text) And IsNumeric(Frm84.TB15) And IsNumeric(Frm84.TB3) Then
                Frm84_LM_BERAT_JUAL = Frm84.TB4 'Berat Jualan
                Frm84_LM_HARGA_SEMASA = Frm84.TB5 'Harga semasa (jualan)
                Frm84_LM_HARGA_SUPPLIER = Frm84.L54_Text 'Harga per gram (harga semasa) dari supplier (modal)
                Frm84_UPAH_MODAL = Frm84.L55_Text 'Upah modal
                Frm84_UPAH_JUAL = Frm84.TB15 'Upah jualan
                Frm84_LM_BERAT_ASAL = Frm84.TB3 'Berat Asal
                
                rs!upah_modal = Frm84.L55_Text 'Upah modal
                rs!harga_per_gram_supplier = Frm84.L54_Text 'Harga per gram (harga semasa) dari supplier (modal)
                rs!untung2 = Format(((Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_SEMASA) + Frm84_UPAH_JUAL) - ((Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_SUPPLIER) + (Frm84_LM_BERAT_JUAL * Frm84_UPAH_MODAL / Frm84_LM_BERAT_ASAL)), "0.00") 'Untung jika restok pada harga supplier ini

            Else
                
                rs!harga_per_gram_supplier = "0.00" 'Harga per gram (harga semasa) dari supplier (modal)
                rs!untung2 = "0.00" 'Untung jika restok pada harga supplier ini
                rs!upah_modal = "0.00" 'Upah modal
                
            End If
            
        End If
        If Format(Frm84.TB3, "0.00") = Format(Frm84.TB4, "0.00") Then
            rs!potong_flag = 0 '0 : Tiada Potong , 1 : Ada Potong
        Else
            rs!potong_flag = 1 '0 : Tiada Potong , 1 : Ada Potong
        End If
        rs!dulang = Frm84_LM_DULANG 'Dulang
    
        '### Maklumat tetapan harga jualan kepada staff ### - Start
        If Frm84.L48_Text <> vbNullString Then
            rs!kadar_penurunan_upah = Format(Frm84.L48_Text, "0.00") 'Kadar peratusan penurunan harga upah kepada staff (%)
        Else
            rs!kadar_penurunan_upah = Null
        End If
        If Frm84.L49_Text <> vbNullString Then
            rs!harga_semasa_staff = Format(Frm84.L49_Text, "0.00") 'Harga emas semasa yang dijual kepada staff
        Else
            rs!harga_semasa_staff = Null
        End If
        If Frm84.L50_Text <> vbNullString Then
            rs!kadar_penurunan_bp = Format(Frm84.L50_Text, "0.00") 'Kadar peratusan penurunan harga barang permata kepada staff (%)
        Else
            rs!kadar_penurunan_bp = Null
        End If
        If Frm84.L51_Text <> vbNullString Then
            rs!harga_staff = Format(Frm84.L51_Text, "0.00") 'Harga yang dijual kepada staff (RM)
        Else
            rs!harga_staff = Null
        End If
        If Frm84.L52_Text <> vbNullString Then
            rs!harga_bp_asal = Format(Frm84.L52_Text, "0.00") 'Tetapan harga barang permata yang asal (RM)
        Else
            rs!harga_bp_asal = Null
        End If
        If Frm84.L53_Text <> vbNullString Then
            rs!upah_asal = Format(Frm84.L53_Text, "0.00") 'Tetapan upah asal (RM)
        Else
            rs!upah_asal = Null
        End If
        rs!komisyen_staff = Format(Frm84_LM_HARGA_PELANGGAN - Frm84_LM_HARGA_STAFF, "0.00") 'Jumlah Komisyen Staff (RM)
        '### Maklumat tetapan harga jualan kepada staff ### - End
        
        If Frm84.CB12 = 0 Then '0 : GST pada harga jualan , 1 : GST pada upah
            rs!gst_barang_atau_upah = 0
        Else
            rs!gst_barang_atau_upah = 1
        End If
        If Frm84.CB3 = 1 And Frm84.CB18 = 0 Then
            Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
            Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
            
            rs!harga_jualan_dengan_gst = Format(Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC, "0.00")
            'Field ini adalah lebih kurang kepada @harga_dengan_gst
            'Tetapi bagi gst yang dikenakan hanya pada UPAH akan membuatkan field ini hanya memaparkan harga dengan gst hanya pada upah.
            'Oleh itu field ini dibuat khas bagi mencampurkan gst pada upah dam harga barang.
        Else
            rs!harga_jualan_dengan_gst = Format(Frm84.TB10, "0.00") 'Harga Jualan (RM)
        End If
        If Frm84.L67_Text <> vbNullString Then 'Purata harga jualan per gram (RM/g) bagi barang kemas , Bagi barang permata adalah merujuk kepada harga jualan
            rs!jualan_per_gram = Format(Frm84.L67_Text, "0.00")
        Else
            rs!jualan_per_gram = Null
        End If
        If Frm84.L69_Text <> vbNullString Then 'Paparan modal per gram (tanpa GST)
            rs!modal_per_gram = Format(Frm84.L69_Text, "0.00")
        Else
            rs!modal_per_gram = Null
        End If
        
        rs.Update
        Frm84_LM_DATA_SAVE = 1
    Else
        If Frm84.TB2 <> vbNullString Then
            rs!no_siri_Produk = Frm84.TB2 'No. Siri Produk
        Else
            rs!no_siri_Produk = Null 'No. Siri Produk
        End If
        If Frm84.L12_Text <> vbNullString Then
            rs!kategori_Produk = Frm84.L12_Text 'Kategori Produk
        Else
            rs!kategori_Produk = Null 'Kategori Produk
        End If
        If Frm84.L13_Text <> vbNullString Then
            rs!purity = Frm84.L13_Text 'Purity
        Else
            rs!purity = Null 'Purity
        End If
        If Frm84.TB3 <> vbNullString Then
            rs!Berat_Asal = Format(Frm84.TB3, "0.00") 'Berat Asal (g)
        Else
            rs!Berat_Asal = Null 'Berat Asal (g)
        End If
        If Frm84.TB4 <> vbNullString Then
            rs!berat_jualan = Format(Frm84.TB4, "0.00") 'Berat Jualan (g)
        Else
            rs!berat_jualan = Null 'Berat Jualan (g)
        End If
        If Frm84.TB5 <> vbNullString Then
            rs!harga_Semasa = Format(Frm84.TB5, "0.00") 'Harga Semasa (RM/g)
        Else
            rs!harga_Semasa = Null 'Harga Semasa (RM/g)
        End If
        If Frm84.TB15 <> vbNullString Then
            rs!UPAH = Format(Frm84.TB15, "0.00") 'Upah (RM)
        Else
            rs!UPAH = Null 'Upah (RM)
        End If
        If Frm84.TB6 <> vbNullString Then
            rs!harga_asal = Format(Frm84.TB6, "0.00") 'Harga Asal Item (RM)
        Else
            rs!harga_asal = Null 'Harga Asal Item (RM)
        End If
        If Frm84.TB7 <> vbNullString Then
            rs!diskaun = Format(Frm84.TB7, "0.00") 'Diskaun (%)
        Else
            rs!diskaun = Null 'Diskaun (%)
        End If
        If Frm84.TB8 <> vbNullString Then
            rs!harga_lepas_diskaun = Format(Frm84.TB8, "0.00") 'Harga Selepas Diskaun (RM)
        Else
            rs!harga_lepas_diskaun = Null 'Harga Selepas Diskaun (RM)
        End If
        If Frm84.TB9 <> vbNullString Then
            rs!adjustment = Format(Frm84.TB9, "0.00") 'Adjustment (RM)
        Else
            rs!adjustment = Null 'Adjustment (RM)
        End If
        If Frm84.TB10 <> vbNullString Then
            rs!harga_jualan = Format(Frm84.TB10, "0.00") 'Harga Jualan (RM)
        Else
            rs!harga_jualan = Null 'Harga Jualan (RM)
        End If
        If Frm84.CB2 = 1 Then
            rs!gst_ari_nashi = "ZR (L)" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
            rs!kadar_gst = Null 'Kadar Cukai GST (%)
            
            If Frm84.TB11 <> vbNullString Then
                rs!jumlah_gst = Format(Frm84.TB11, "0.00") 'Jumlah Cukai GST (RM)
            Else
                rs!jumlah_gst = "0.00" 'Jumlah Cukai GST (RM)
            End If
        ElseIf Frm84.CB3 = 1 Then
            rs!gst_ari_nashi = "SR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
            If Frm84.L8_Text <> vbNullString Then
                rs!kadar_gst = Frm84.L8_Text 'Kadar Cukai GST (%)
            Else
                rs!kadar_gst = "0" 'Jumlah Cukai GST (RM)
            End If
            If Frm84.TB11 <> vbNullString Then
                rs!jumlah_gst = Format(Frm84.TB11, "0.00") 'Jumlah Cukai GST (RM)
            Else
                rs!jumlah_gst = "0.00" 'Jumlah Cukai GST (RM)
            End If
            If Frm84.CB18 = 1 Then 'Jenis Cukai GST SR
                rs!gst_include = 1 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
            Else
                rs!gst_include = 0 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
            End If
        End If
        If Frm84.L44_Text <> vbNullString Then
            rs!harga_tanpa_gst = Format(Frm84.L44_Text, "0.00") 'Harga Jualan Tanpa GST (RM)
        Else
            rs!harga_tanpa_gst = Null 'Harga Jualan Tanpa GST (RM)
        End If
        If Frm84.TB14 <> vbNullString Then
            rs!harga_dengan_gst = Format(Frm84.TB14, "0.00") 'Harga Jualan Termasuk GST (RM)
        Else
            rs!harga_dengan_gst = Null 'Harga Jualan Termasuk GST (RM)
        End If
        If Frm84.CB7 = 1 Then
            rs!dropship = 1 '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
            If Frm84.Frame2.Visible = True Then 'Komisen Agen Dropship : BK
                If Frm84.TB12 <> vbNullString Then
                    rs!komisyen_per_gram = Format(Frm84.TB12, "0.00") 'Komisyen Per Gram Dropship (RM/g) : BK
                Else
                    rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g) : BK
                End If
                If Frm84.TB13 <> vbNullString Then
                    rs!jumlah_komisyen = Format(Frm84.TB13, "0.00") 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
                Else
                    rs!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : BK
                End If
            End If
            If Frm84.Frame3.Visible = True Then 'Komisen Agen Dropship : Permata
                rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g) : BK
                If Frm84.TB16 <> vbNullString Then
                    rs!jumlah_komisyen = Format(Frm84.TB16, "0.00") 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : Permata
                Else
                    rs!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini : Permata
                End If
            End If
        End If
            
        If Frm84.CB7 = 0 Then
            rs!dropship = 0 '0 : Jualan Bukan Oleh Agen Dropship , 1 : Jualan Oleh Agen Dropship
            rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g)
            rs!jumlah_komisyen = Null 'Jumlah Komisyen Kepada Agen Dropship Bagi Item Ini
        End If
        
        If Frm84.L41_Text = "0" Then '0 : Menu Data Baru , 1 : Menu Edit Data
            rs!Status = 1
        ElseIf Frm84.L41_Text = "1" Then
            rs!Status = 3
        End If
        
        If Frm84.TB3 = vbNullString Then
            rs!Type = 1 '0 : BK , 1 : Barang Permata
            rs!harga_per_gram_modal = Null 'Harga Per Gram Bagi Modal (RM/g)
            If Frm84.L34_Text <> vbNullString Then
                rs!modal = Format(Frm84.L34_Text, "0.00") 'Harga Modal (RM)
            Else
                rs!modal = Null 'Harga Modal (RM)
            End If
            If IsNumeric(Frm84.TB14) And IsNumeric(Frm84.L34_Text) Then
                Frm84_LM_HARGA_MODAL = Frm84.L34_Text 'Harga Modal
                Frm84_LM_HARGA_JUAL = Frm84.TB14 'Harga Jualan
                
                rs!untung = Format(Frm84_LM_HARGA_JUAL - Frm84_LM_HARGA_MODAL, "0.00") 'Jumlah Keuntungan
            End If
            
        Else
            rs!Type = 0 '0 : BK , 1 : Barang Permata
            
            If Frm84.L34_Text <> vbNullString Then
                rs!harga_per_gram_modal = Format(Frm84.L34_Text, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
                If IsNumeric(Frm84.L34_Text) Then
                    Frm84_LM_HARGA_SEMASA_MODAL = Frm84.L34_Text 'Harga Per Gram Bagi Modal (RM/g)
                    
                    rs!modal = Format(Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_SEMASA_MODAL, "0.00") 'Harga Modal (RM)
                End If
            Else
                rs!modal = Null 'Harga Modal (RM)
                rs!harga_per_gram_modal = Null 'Harga Per Gram Bagi Modal (RM/g)
                rs!upah_modal = Null 'Upah modal
            End If

            If IsNumeric(Frm84.TB14) And IsNumeric(Frm84.L34_Text) And IsNumeric(Frm84.TB10) And IsNumeric(Frm84.TB11) Then
                
                Frm84_LM_HARGA_MODAL = Frm84.L34_Text 'Harga Modal
                Frm84_LM_HARGA_JUAL = Frm84.TB14 'Harga Jualan
                Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
                Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
                
                If Frm84.CB12 = 0 Then
                
                    rs!untung = Format(Frm84_LM_HARGA_JUAL - (Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_MODAL), "0.00") 'Jumlah Keuntungan
                    
                ElseIf Frm84.CB12 = 1 Then
                    
                    If Frm84.CB3 = 1 And Frm84.CB18 = 0 Then
                        
                        rs!untung = Format((Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC) - (Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_MODAL), "0.00") 'Jumlah Keuntungan
                        
                    Else
                        
                        rs!untung = Format(Frm84_LM_HARGA_JUALAN_CALC - (Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_MODAL), "0.00") 'Jumlah Keuntungan
                        
                    End If
                    
                End If
                
            End If
            
            If IsNumeric(Frm84.TB4) And IsNumeric(Frm84.TB5) And IsNumeric(Frm84.L54_Text) And IsNumeric(Frm84.L55_Text) And IsNumeric(Frm84.TB15) And IsNumeric(Frm84.TB3) Then
                Frm84_LM_BERAT_JUAL = Frm84.TB4 'Berat Jualan
                Frm84_LM_HARGA_SEMASA = Frm84.TB5 'Harga semasa (jualan)
                Frm84_LM_HARGA_SUPPLIER = Frm84.L54_Text 'Harga per gram (harga semasa) dari supplier (modal)
                Frm84_UPAH_MODAL = Frm84.L55_Text 'Upah modal
                Frm84_UPAH_JUAL = Frm84.TB15 'Upah jualan
                Frm84_LM_BERAT_ASAL = Frm84.TB3 'Berat Asal
                
                rs!upah_modal = Frm84.L55_Text 'Upah modal
                rs!harga_per_gram_supplier = Frm84.L54_Text 'Harga per gram (harga semasa) dari supplier (modal)
                rs!untung2 = Format(((Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_SEMASA) + Frm84_UPAH_JUAL) - ((Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_SUPPLIER) + (Frm84_LM_BERAT_JUAL * Frm84_UPAH_MODAL / Frm84_LM_BERAT_ASAL)), "0.00") 'Untung jika restok pada harga supplier ini

            Else
                
                rs!harga_per_gram_supplier = "0.00" 'Harga per gram (harga semasa) dari supplier (modal)
                rs!untung2 = "0.00" 'Untung jika restok pada harga supplier ini
                rs!upah_modal = "0.00" 'Upah modal
                
            End If
            
        End If
        If Format(Frm84.TB3, "0.00") = Format(Frm84.TB4, "0.00") Then
            rs!potong_flag = 0 '0 : Tiada Potong , 1 : Ada Potong
        Else
            rs!potong_flag = 1 '0 : Tiada Potong , 1 : Ada Potong
        End If
        
        rs!dulang = Frm84_LM_DULANG 'Dulang
        
        '### Maklumat tetapan harga jualan kepada staff ### - Start
        If Frm84.L48_Text <> vbNullString Then
            rs!kadar_penurunan_upah = Format(Frm84.L48_Text, "0.00") 'Kadar peratusan penurunan harga upah kepada staff (%)
        Else
            rs!kadar_penurunan_upah = Null
        End If
        If Frm84.L49_Text <> vbNullString Then
            rs!harga_semasa_staff = Format(Frm84.L49_Text, "0.00") 'Harga emas semasa yang dijual kepada staff
        Else
            rs!harga_semasa_staff = Null
        End If
        If Frm84.L50_Text <> vbNullString Then
            rs!kadar_penurunan_bp = Format(Frm84.L50_Text, "0.00") 'Kadar peratusan penurunan harga barang permata kepada staff (%)
        Else
            rs!kadar_penurunan_bp = Null
        End If
        If Frm84.L51_Text <> vbNullString Then
            rs!harga_staff = Format(Frm84.L51_Text, "0.00") 'Harga yang dijual kepada staff (RM)
        Else
            rs!harga_staff = Null
        End If
        If Frm84.L52_Text <> vbNullString Then
            rs!harga_bp_asal = Format(Frm84.L52_Text, "0.00") 'Tetapan harga barang permata yang asal (RM)
        Else
            rs!harga_bp_asal = Null
        End If
        If Frm84.L53_Text <> vbNullString Then
            rs!upah_asal = Format(Frm84.L53_Text, "0.00") 'Tetapan upah asal (RM)
        Else
            rs!upah_asal = Null
        End If
        rs!komisyen_staff = Format(Frm84_LM_HARGA_PELANGGAN - Frm84_LM_HARGA_STAFF, "0.00") 'Jumlah Komisyen Staff (RM)
        '### Maklumat tetapan harga jualan kepada staff ### - End
        
        If Frm84.CB12 = 0 Then '0 : GST pada harga jualan , 1 : GST pada upah
            rs!gst_barang_atau_upah = 0
        Else
            rs!gst_barang_atau_upah = 1
        End If
        If Frm84.CB3 = 1 And Frm84.CB18 = 0 Then
            Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
            Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
            
            rs!harga_jualan_dengan_gst = Format(Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC, "0.00")
            'Field ini adalah lebih kurang kepada @harga_dengan_gst
            'Tetapi bagi gst yang dikenakan hanya pada UPAH akan membuatkan field ini hanya memaparkan harga dengan gst hanya pada upah.
            'Oleh itu field ini dibuat khas bagi mencampurkan gst pada upah dam harga barang.
        Else
            rs!harga_jualan_dengan_gst = Format(Frm84.TB10, "0.00") 'Harga Jualan (RM)
        End If
        If Frm84.L67_Text <> vbNullString Then 'Purata harga jualan per gram (RM/g) bagi barang kemas , Bagi barang permata adalah merujuk kepada harga jualan
            rs!jualan_per_gram = Format(Frm84.L67_Text, "0.00")
        Else
            rs!jualan_per_gram = Null
        End If
        If Frm84.L69_Text <> vbNullString Then 'Paparan modal per gram (tanpa GST)
            rs!modal_per_gram = Format(Frm84.L69_Text, "0.00")
        Else
            rs!modal_per_gram = Null
        End If
        
        rs.Update
        Frm84_LM_DATA_SAVE = 1
    End If
    
    rs.Close
    Set rs = Nothing
'### Masukkan Data Ke Dalam Temp Table ### - End
    
    If Frm84_LM_DATA_SAVE = 1 Then
        'Call Frm84_Reset
        Call Frm84_Reset_Edit
        Call Frm84_Senarai_Jualan_Header
        Call Frm84_Senarai_Jualan
        
        MsgBox "Data Telah Berjaya Dimasukkan Ke Dalam Senarai Jualan.", vbInformation, "Info"
    End If
End If
End Sub
Sub Frm84_pengiraan_harga_staff()
'on error resume next
Dim Frm84_LM_BERAT As Double
Dim Frm84_LM_HARGA_SEMASA As Double
Dim Frm84_LM_UPAH As Double
Dim Frm84_LM_DISKAUN_UPAH As Double
Dim Frm84_LM_HARGA_BP As Double
Dim Frm84_LM_DISKAUN_BP As Double
Dim Frm84_LM_DISKAUN_OVERALL As Double
Dim Frm84_LM_HARGA_KESELURUHAN As Double

Frm84_LM_BERAT = 0
Frm84_LM_HARGA_SEMASA = 0
Frm84_LM_UPAH = 0
Frm84_LM_DISKAUN_UPAH = 0
Frm84_LM_HARGA_BP = 0
Frm84_LM_DISKAUN_BP = 0
Frm84_LM_DISKAUN_OVERALL = 0
Frm84_LM_HARGA_KESELURUHAN = 0

'### Pengiraan bagi barang kemas ### - Start
If ((Frm84.TB7 <> vbNullString And IsNumeric(Frm84.TB7)) And (Frm84.TB4 <> vbNullString And IsNumeric(Frm84.TB4)) And (Frm84.L53_Text <> vbNullString And IsNumeric(Frm84.L53_Text)) And (Frm84.L48_Text <> vbNullString And IsNumeric(Frm84.L48_Text)) And (Frm84.L49_Text <> vbNullString And IsNumeric(Frm84.L49_Text))) Then
    If IsNumeric(Frm84.TB4) Then Frm84_LM_BERAT = Frm84.TB4 'Berat
    If IsNumeric(Frm84.TB7) Then Frm84_LM_DISKAUN_OVERALL = Frm84.TB7 'Diskaun Keseluruhan
    If IsNumeric(Frm84.L49_Text) Then Frm84_LM_HARGA_SEMASA = Frm84.L49_Text 'Harga Per Gram
    If IsNumeric(Frm84.L53_Text) Then Frm84_LM_UPAH = Frm84.L53_Text 'Upah
    If IsNumeric(Frm84.L48_Text) Then Frm84_LM_DISKAUN_UPAH = Frm84.L48_Text 'Diskaun Upah
    
    Frm84_LM_HARGA_KESELURUHAN = (Frm84_LM_BERAT * Frm84_LM_HARGA_SEMASA) + ((1 - (Frm84_LM_DISKAUN_UPAH / 100)) * Frm84_LM_UPAH)
    Frm84.L51_Text = Format(Frm84_LM_HARGA_KESELURUHAN - (((Frm84_LM_DISKAUN_OVERALL / 100)) * Frm84_LM_HARGA_KESELURUHAN), "#,##0.00") 'Harga Staff
Else
    Frm84.L51_Text = "0.00" 'Harga Staff
End If
'### Pengiraan bagi barang kemas ### - End
End Sub
Sub Frm84_pengiraan_harga_bp_staff()
'on error resume next
Dim Frm84_LM_HARGA_BP As Double
Dim Frm84_LM_DISKAUN_BP As Double
Dim Frm84_LM_HARGA_KESELURUHAN As Double
Dim Frm84_LM_DISKAUN_OVERALL As Double

Frm84_LM_HARGA_BP = 0
Frm84_LM_DISKAUN_BP = 0
Frm84_LM_HARGA_KESELURUHAN = 0
Frm84_LM_DISKAUN_OVERALL = 0

'### Pengiraan bagi barang permata ### - Start
If ((Frm84.TB7 <> vbNullString And IsNumeric(Frm84.TB7)) And (Frm84.L50_Text <> vbNullString And IsNumeric(Frm84.L50_Text)) And (Frm84.L52_Text <> vbNullString And IsNumeric(Frm84.L52_Text))) Then
    If IsNumeric(Frm84.TB7) Then Frm84_LM_DISKAUN_OVERALL = Frm84.TB7 'Diskaun Keseluruhan
    If IsNumeric(Frm84.TB7) Then Frm84_LM_DISKAUN_OVERALL = Frm84.TB7 'Diskaun Keseluruhan
    If IsNumeric(Frm84.L52_Text) Then Frm84_LM_HARGA_BP = Frm84.L52_Text 'Harga Asal
    If IsNumeric(Frm84.L50_Text) Then Frm84_LM_DISKAUN_BP = Frm84.L50_Text 'Diskaun
    
    Frm84_LM_HARGA_KESELURUHAN = Frm84_LM_HARGA_BP - ((Frm84_LM_DISKAUN_BP / 100) * Frm84_LM_HARGA_BP)
    Frm84.L51_Text = Format(Frm84_LM_HARGA_KESELURUHAN - ((Frm84_LM_DISKAUN_OVERALL / 100) * Frm84_LM_HARGA_KESELURUHAN), "#,##0.00") 'Harga Staff
Else
    Frm84.L51_Text = "0.00" 'Harga Staff
End If
'### Pengiraan bagi barang permata ### - End
End Sub
Sub Frm84_pengiraan_harga_jualan()
'on error resume next
Dim Frm84_HARGA_LEPAS_DISKAUN As Double
Dim Frm84_ADJUSTMENT As Double
Dim Frm84_POSTAGE As Double
Dim Frm84_KUPON As Double
Dim Frm84_REDEEM As Double



If GLOBAL_DISABLE = 0 Then

    Frm84_HARGA_LEPAS_DISKAUN = 0
    Frm84_ADJUSTMENT = 0
    Frm84_POSTAGE = 0
    Frm84_KUPON = 0
    Frm84_REDEEM = 0
    
    If ((Frm84.L20_Text <> vbNullString And IsNumeric(Frm84.L20_Text)) And (Frm84.TB34 <> vbNullString And IsNumeric(Frm84.TB34)) And (Frm84.TB20 <> vbNullString And IsNumeric(Frm84.TB20)) And (Frm84.TB42 <> vbNullString And IsNumeric(Frm84.TB42)) And (Frm84.L73_Text <> vbNullString And IsNumeric(Frm84.L73_Text))) Then
        Frm84_HARGA_LEPAS_DISKAUN = Frm84.L20_Text 'Harga Lepas Diskaun
        Frm84_ADJUSTMENT = Frm84.TB20 'Adjustment
        Frm84_POSTAGE = Frm84.TB42 'Pos Laju
        Frm84_KUPON = Frm84.TB34 'Kupon
        Frm84_REDEEM = Frm84.L73_Text 'Tebus Mata Ganjaran

        Frm84.L21_Text = Format(Frm84_HARGA_LEPAS_DISKAUN + Frm84_POSTAGE - Frm84_ADJUSTMENT - Frm84_KUPON - Frm84_REDEEM, "#,##0.00") 'Harga Jualan
    Else
        Frm84.L21_Text = "0.00" 'Harga Jualan
    End If
    
End If
End Sub
Sub Frm84_pengiraan_komisyen_upah()
'on error resume next
Dim Frm84_UPAH As Double
Dim Frm84_KADAR_KOMISYEN As Double

If GLOBAL_DISABLE = 0 Then

    Frm84_UPAH = 0
    Frm84_KADAR_KOMISYEN = 0
    
    If ((Frm84.TB43 <> vbNullString And IsNumeric(Frm84.TB43)) And (Frm84.TB15 <> vbNullString And IsNumeric(Frm84.TB15))) Then
        Frm84_KADAR_KOMISYEN = Frm84.TB43 'Kadar komisyen (%)
        Frm84_UPAH = Frm84.TB15 'Upah jualan
        
        Frm84.TB44 = Format(((Frm84_KADAR_KOMISYEN / 100) * Frm84_UPAH), "#,##0.00") 'Jumlah komisyen bagi UPAH
    Else
        Frm84.TB44 = "0.00" 'Jumlah komisyen bagi UPAH
    End If
    
End If
End Sub
Sub Frm84_pengiraan_komisyen_dropship()
'on error resume next
Dim Frm84_UPAH As Double
Dim Frm84_BERAT As Double
Dim Frm84_RATE_KOMISYEN As Double

If GLOBAL_DISABLE = 0 Then

    Frm84_UPAH = 0
    Frm84_BERAT = 0
    Frm84_RATE_KOMISYEN = 0
    
    If ((Frm84.TB44 <> vbNullString And IsNumeric(Frm84.TB44)) And (Frm84.TB12 <> vbNullString And IsNumeric(Frm84.TB12)) And (Frm84.TB4 <> vbNullString And IsNumeric(Frm84.TB4))) Then
        Frm84_UPAH = Frm84.TB44 'Komisyen upah
        Frm84_BERAT = Frm84.TB4 'Berat jualan
        Frm84_RATE_KOMISYEN = Frm84.TB12 'Kadar komisyen per gram
        
        Frm84.TB13 = Format((Frm84_RATE_KOMISYEN * Frm84_BERAT) + Frm84_UPAH, "#,##0.00") 'Jumlah komisyen keseluruhan
    Else
        Frm84.TB13 = "0.00" 'Jumlah komisyen keseluruhan
    End If
    
End If
End Sub
Sub Frm84_trade_in_barang()
'on error resume next
'kakunin
Frm84.L16_Text = vbNullString 'No. Voucher
Frm84.TB17 = "0.00" 'Jumlah trade in

'Frm84.L59_Text = 0 '0 : Barang baru , 1 : Edit

If Frm84.L59_Text = 1 Then

    Frm83.L12_Text = Frm84.L60_Text
    Frm83.L9_Text = Frm84.L61_Text
    Frm84.Frame6.Visible = False
    Frm84.L57_Text = Frm84.L60_Text 'No. Voucher
    Frm84.L58_Text = Frm83.L26_Text 'Jumlah trade in
    Frm84.L56_Text = 2 '0 : Tiada , 1 : Belian dengan trade in (sudah ada voucher) , 2 : Trade in

End If

If (Frm84.L56_Text = 0 Or Frm84.L56_Text = 1) And Frm84.L59_Text = 0 Then '0 : Tiada , 1 : Belian dengan trade in (sudah ada voucher) , 2 : Trade in
'If Frm84.L59_Text = 0 Then '0 : Tiada , 1 : Belian dengan trade in (sudah ada voucher) , 2 : Trade in

    Frm83.CB10 = 0
    Frm83.CB8 = 1
    Frm83.CB9 = 1
    
    Frm83.CBB1.Enabled = False
    Frm83.CBB1.BackColor = &H8000000A

    Frm84.L56_Text = 2 '0 : Tiada , 1 : Belian dengan trade in (sudah ada voucher) , 2 : Trade in
    
    Frm84.L57_Text = vbNullString 'No. Voucher
    Frm84.L58_Text = vbNullString 'Jumlah trade in
    Frm84.Frame6.Visible = False

    Call Frm83_Initial_Setting
    Call Frm83_initial_setting2
    Call Frm83_Reset_Form
    
    Frm83_LM_No_RUJUKAN_BELIAN = Frm83.L9_Text 'No. Rujukan Belian
    GM_No_RUJUKAN_BELIAN = Format(Frm83.L9_Text, "000000") 'No. Rujukan Belian
    
    Frm83_LM_No_INVOICE = Frm83.L12_Text 'No. Invoice Trade In
    
    GoTo skip_a:
    
'### Periksa NO RUJUKAN SISTEM sebelum simpan data ke dalam database ### - Start
Re_gen_no_ruj:
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 16_gold_bar_belian where no_rujukan='" & Format(Frm83_LM_No_RUJUKAN_BELIAN, "000000") & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        Frm83_LM_No_RUJUKAN_BELIAN = Frm83_LM_No_RUJUKAN_BELIAN + 1
        Frm83.L9_Text = Frm83_LM_No_RUJUKAN_BELIAN
        
        rs.Close
        Set rs = Nothing
        
        GoTo Re_gen_no_ruj:
    End If
    
    rs.Close
    Set rs = Nothing
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from data_database where NoRujukanSistem='" & Format(Frm83_LM_No_RUJUKAN_BELIAN, "000000") & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        Frm83_LM_No_RUJUKAN_BELIAN = Frm83_LM_No_RUJUKAN_BELIAN + 1
        Frm83.L9_Text = Frm83_LM_No_RUJUKAN_BELIAN
        
        rs.Close
        Set rs = Nothing
        
        GoTo Re_gen_no_ruj:
    End If
    
    rs.Close
    Set rs = Nothing
'### Periksa NO RUJUKAN SISTEM sebelum simpan data ke dalam database ### - End
    
    
Re_gen_no_TI:
        
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & "TI" & Format(Frm83_LM_No_INVOICE, "000000") & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        Frm83_LM_No_INVOICE = Frm83_LM_No_INVOICE + 1
        Frm83.L12_Text = Frm83_LM_No_INVOICE
        
        rs.Close
        Set rs = Nothing
        
        GoTo Re_gen_no_TI:
    End If
    
    rs.Close
    Set rs = Nothing
        
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from data_database where bill_No_Trade_In='" & "TI" & Format(Frm83_LM_No_INVOICE, "000000") & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        Frm83_LM_No_INVOICE = Frm83_LM_No_INVOICE + 1
        Frm83.L12_Text = Frm83_LM_No_INVOICE
        
        rs.Close
        Set rs = Nothing
        
        GoTo Re_gen_no_TI:
    End If
    
    rs.Close
    Set rs = Nothing
    
skip_a:

    'Frm84.L57_Text = "TI" & Format(Frm83_LM_No_INVOICE, "000000")
    Frm84.L57_Text = Frm83.L12_Text

End If

'1:  Pelanggan
'2:  Member
'3:  RAF
'4:  Pengedar
'5:  Normal Dealer
'6:  Master Dealer
If Frm84.CB4 = 1 Then
    Frm83.L39_Text = 1
ElseIf Frm84.CB5 = 1 Then
    Frm83.L39_Text = 2
ElseIf Frm84.CB6 = 1 Then
    Frm83.L39_Text = 4
ElseIf Frm84.CB9 = 1 Then
    Frm83.L39_Text = 3
ElseIf Frm84.CB10 = 1 Then
    Frm83.L39_Text = 5
'ElseIf Frm84.CB11 = 1 Then
'    Frm83.L39_Text = 6
End If

Frm83.L41_Text = 1 '0 : Belian emas terpakai , 1 : Trade in , 2 : Barang baru

Frm83.CBB1.AddItem "Trade In"
Frm83.CBB1 = "Trade In"
Frm83.TB1 = "TI"

Frm83.CB2 = 1
Frm83.CB3 = 0
Frm83.CB11 = 0
Frm83.CB12 = 0

Frm83.CB2.Enabled = False
Frm83.CB3.Enabled = False
Frm83.CB11.Enabled = False
Frm83.CB12.Enabled = False

Frm83.CMD5.Visible = False
Frm83.CMD10.Visible = False
Frm83.CMD11.Visible = False
Frm83.CMD2.Caption = "Kembali ke menu jualan"

Frm83.Show
    
Frm84.Hide
End Sub
Sub Frm84_penerimaan_barang_trade_in()
'On Error Resume Next
Dim Err(5)
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
x = 0
Y = 0
DATA_SAVE = 0

'$$$ No. staff $$$ - Start
If InStr(1, Frm84.CBB1, "  |  ") <> 0 Then

    Frm84_LM_EMP_NO = Split(Frm84.CBB1, "  |  ")(1)
    
Else

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        
        If Not IsNull(rs!NoPekerja) Then Frm84_LM_EMP_NO = rs!NoPekerja

    End If
    
    rs.Close
    Set rs = Nothing

End If
'$$$ No. staff $$$ - End

'###Masukkan Data Belian Ke Dalam Database Akaun Belian### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 16_gold_bar_belian where no_rujukan='" & Format(Frm83.L9_Text, "000000") & "'", cn, adOpenKeyset, adLockOptimistic

If rs.EOF Then
    rs.AddNew
    If Frm83.L9_Text <> vbNullString Then
        rs!no_rujukan = Format(Frm83.L9_Text, "000000") 'No. Rujukan Belian
        GM_No_RUJUKAN_BELIAN = Format(Frm83.L9_Text, "000000") 'No. Rujukan Belian
    Else
        rs!no_rujukan = Null 'No. Rujukan Belian
    End If
    rs!tarikh = Frm84.DTPicker1
    'If Frm83.DTPicker1 <> vbNullString Then %%%%Default : Bagi sistem ini hanya belian secara CASH dibenarkan
        rs!cara_bayaran = 0 'Cara Belian // 0 : Cash @ Bank in @ Kad Kredit @ Kad Debit , 1 : Cheque
    'Else
    '    rs!cara_bayaran = Null 'Cara Belian // 0 : Cash @ Bank in @ Kad Kredit @ Kad Debit , 1 : Cheque
    'End If
    If Frm83.L26_Text <> vbNullString Then
        rs!tunai = Format(Frm83.L26_Text, "0.00") 'Cara Bayaran : Tunai (Jumlah Keseluruhan Dengan Cukai GST)
    Else
        rs!tunai = Null 'Cara Bayaran : Tunai (Jumlah Keseluruhan Dengan Cukai GST)
    End If
    If Frm83.L11_Text <> vbNullString Then
        rs!jumlah_asal = Format(Frm83.L11_Text, "0.00") 'Jumlah Bayaran Asal (Jumlah Tanpa Cukai GST)
    Else
        rs!jumlah_asal = Null 'Jumlah Bayaran Asal (Jumlah Tanpa Cukai GST)
    End If

    If Frm83.L8_Text <> vbNullString Then
        rs!gst_value = Frm83.L8_Text 'Jumlah Cukai GST (%)
    Else
        rs!gst_value = "0" 'Jumlah Cukai GST (%)
    End If
    If Frm83.L22_Text <> vbNullString Then
        rs!gst_zr_harga = Format(Frm83.L22_Text, "0.00") 'Jumlah Bayaran Yang Dikenakan Cukai GST ZR (RM)
    Else
        rs!gst_zr_harga = "0.00" 'Jumlah Bayaran Yang Dikenakan Cukai GST ZR (RM)
    End If
    If Frm83.L23_Text <> vbNullString Then
        rs!gst_zr_cukai = Format(Frm83.L23_Text, "0.00") 'Jumlah Bayaran Cukai GST ZR (RM)
    Else
        rs!gst_zr_cukai = "0.00" 'Jumlah Bayaran Cukai GST ZR (RM)
    End If
    If Frm83.L24_Text <> vbNullString Then
        rs!gst_sr_harga = Format(Frm83.L24_Text, "0.00") 'Jumlah Bayaran Yang Dikenakan Cukai GST SR (RM)
    Else
        rs!gst_sr_harga = "0.00" 'Jumlah Bayaran Yang Dikenakan Cukai GST SR (RM)
    End If
    If Frm83.L25_Text <> vbNullString Then
        rs!gst_sr_cukai = Format(Frm83.L25_Text, "0.00") 'Jumlah Bayaran Cukai GST SR (RM)
    Else
        rs!gst_sr_cukai = "0.00" 'Jumlah Bayaran Cukai GST SR (RM)
    End If
    If Frm83.TB28 <> vbNullString Then
        rs!no_id_gst_supplier = Frm83.TB28 'No. ID GST Supplier
    Else
        rs!no_id_gst_supplier = Null 'No. ID GST Supplier
    End If
    If Frm83.TB15 <> vbNullString Then
        rs!no_resit_supplier = UCase(Frm83.TB15) 'No. Resit Dari Supplier (Jika Ada)
    Else
        rs!no_resit_supplier = Null 'No. Resit Dari Supplier (Jika Ada)
    End If
    If Frm83.TB1 <> vbNullString Then
        rs!Kod_Supplier = UCase(Frm83.TB1) 'Kod Supplier
    Else
        rs!Kod_Supplier = Null 'Kod Supplier
    End If
    If Frm83.L11_Text <> vbNullString Then
        rs!jumlah_tanpa_gst = Format(Frm83.L11_Text, "0.00") 'Jumlah Bayaran Asal (Jumlah Tanpa Cukai GST)
    Else
        rs!jumlah_tanpa_gst = Null 'Jumlah Bayaran Asal (Jumlah Tanpa Cukai GST)
    End If
    If Frm83.L26_Text <> vbNullString Then
        rs!jumlah_dengan_gst = Format(Frm83.L26_Text, "0.00") 'Jumlah Bayaran Keseluruhan (Jumlah Dengan Cukai GST)
    Else
        rs!jumlah_dengan_gst = Null 'Jumlah Bayaran Keseluruhan (Jumlah Dengan Cukai GST)
    End If

    rs!flag_trade_in = 1 'Flag Trade In // 0 : Tiada , 1 : Ada
    rs!trade_in_status = 1 'Flag Samada Trade In Sudah Digunakan Atau Tidak , 0 : Tiada , 1 : Ada
    rs!no_resit_trade_in = "TI" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84.L57_Text, "000000") 'No. Resit Trade In
    G_No_RESIT_JUALAN = "TI" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84.L57_Text, "000000") 'No. Resit Trade In
    
    If Frm83.L39_Text <> vbNullString Then 'Pelanggan Biasa
        rs!kategori_penjual = Frm83.L39_Text
    Else
        rs!kategori_penjual = Null
    End If

    If Frm84.L28_Text <> vbNullString Then
        If Frm28.L5_Text <> vbNullString Then
            rs!no_rujukan_pelanggan_buyback = Frm28.L5_Text 'No. Rujukan Pembeli
        Else
            rs!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
        End If
    Else
        rs!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
    End If

    rs!no_pekerja = Frm84_LM_EMP_NO 'No. Pekerja
    rs!terminal = G_TERMINAL
    rs!write_timestamp = Now
    rs!remarks = "Penerimaan stok baru"
    rs!Status = 1
            
    rs.Update
End If

rs.Close
Set rs = Nothing
'###Masukkan Data Belian Ke Dalam Database Akaun Belian### - End

'###Masukkan Data Penjual Yang Tidak Berdaftar Ke Dalam Database### - Start
If Frm28.L5_Text = vbNullString And Frm26.TB1 <> vbNullString Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 44_senarai_pelanggan", cn, adOpenKeyset, adLockOptimistic
    
    rs.AddNew
    rs!tarikh = Frm84.DTPicker1 'Tarikh
    rs!no_resit = "TI" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84.L57_Text, "000000") 'No. Resit Trade In'Frm84.L57_Text 'No. Resit Trade In
    If Frm26.TB1 <> vbNullString Then 'Nama
        rs!Nama = UCase(Frm26.TB1)
    Else
        rs!Nama = Null
    End If
    If Frm26.TB2 <> vbNullString Then 'No. Telefon
        rs!no_tel = UCase(Frm26.TB2)
    Else
        rs!no_tel = Null
    End If
    rs!write_timestamp = Now
    rs!no_staff = Frm84_LM_EMP_NO 'No. Pekerja
    rs!terminal = G_TERMINAL
    rs!jenis_urusan = G_JENIS_URUSAN
    rs!cawangan = G_CAWANGAN
    rs.Update
            
    rs.Close
    Set rs = Nothing
    
End If
'###Masukkan Data Penjual Yang Tidak Berdaftar Ke Dalam Database### - End


'### Masukkan maklumat data barang ke dalam table #data_database ### - Start
LM_NO_RUJ_PEMBELI = vbNullString

If Frm28.L5_Text <> vbNullString Then
    LM_NO_RUJ_PEMBELI = Frm28.L5_Text 'No. Rujukan Pembeli
End If
            
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "insert into Data_Database(NoRujukanSistem,tarikh_belian,bill_No_Belian,terminal,supplier_ID,nama_Supplier,Kod_Supplier,purity_ID,purity,kod_Purity,kategori_produk_ID,kategori_Produk,Kod_Kategori_Produk,Berat,beza_berat,UPAH,Upah30,riyal,kos_Belian_Gram,kos_Belian_Item,SpreadValue,harga_lepas_spread,adjustment,kos_item_tanpa_tax,cara_Belian,dimension_Panjang,dimension_Lebar,dimension_Saiz,code1,code2,harga_Per_Gram_Item,dulang,no_cert,gst_barang_atau_upah,StatusItem,Upah_Jualan,Upah_Member,Upah_RAF,Upah_Pengedar,code_Supplier,HargaJualan_Member,HargaJualan_Pengedar,upah_normal_dealer,upah_master_dealer,HargaJualan_RAF,hargajualan_normal_dealer,hargajualan_master_dealer,remarks,gst_ari_nashi,kadar_gst,jumlah_gst,harga_item,receiving_Status,harga_tanpa_gst,gst_included,jenis_trade_in,flag_upah,upah_per_gram,flag_image,write_timestamp,no_id_gst,susut_berat,no_pekerja)" & _
            "select '" & Format(Frm83.L9_Text, "000000") & "',tarikh_belian,bill_No_Belian,terminal,supplier_ID,nama_Supplier,Kod_Supplier,purity_ID,purity,kod_Purity,kategori_produk_ID,kategori_Produk,Kod_Kategori_Produk,Berat,beza_berat,UPAH,Upah30,riyal,kos_Belian_Gram,kos_Belian_Item,Spread,harga_lepas_spread,adjustment,kos_item_tanpa_tax,cara_Belian,dimension_Panjang,dimension_Lebar,dimension_Saiz,code1,code2,harga_Per_Gram_Item,dulang,no_cert,gst_barang_atau_upah,10,Upah_Jualan,Upah_Member,Upah_RAF,Upah_Pengedar,code_Supplier,HargaJualan_Member,HargaJualan_Pengedar,upah_normal_dealer,upah_master_dealer,HargaJualan_RAF,hargajualan_normal_dealer,hargajualan_master_dealer,remarks,gst_ari_nashi,kadar_gst,jumlah_gst,harga_item,jenis,harga_tanpa_gst,gst_included,jenis_trade_in,flag_upah,upah_per_gram,flag_image,Now(),no_id_gst,0.00,'" & Frm84_LM_EMP_NO & "' from " & G_BELIAN_TEMP & " WHERE StatusItem='" & 10 & "'"
            
Set rs = cn.Execute(strsql)
Set rs = Nothing
'### Masukkan maklumat data barang ke dalam table #data_database ### - End

'### Update maklumat di bawah ke dalam maklumat barang ### - Start
'@no_siri_produk
'@Barcode
'@bill_no_trade_in
'@no_rujukan_pelanggan_buyback

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from data_database where NoRujukanSistem='" & Format(Frm83.L9_Text, "000000") & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    If Not IsNull(rs!ID) And Not IsNull(rs!Kod_Kategori_Produk) Then
        rs!no_siri_Produk = rs!Kod_Kategori_Produk & Format(rs!ID, "000000")
    Else
        rs!no_siri_Produk = Format(rs!ID, "000000")
    End If
    If Not IsNull(rs!ID) Then
        rs!Barcode = Format(rs!ID, "000000")
    Else
        rs!Barcode = Format(rs!ID, "000000")
    End If
    If Frm83.CB8 = 1 Then
        rs!bill_No_Trade_In = "TI" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84.L57_Text, "000000") '"TI" & Format(Frm83_LM_NO_TI, "000000") 'No. Resit Trade In

        If Frm83.L37_Text <> vbNullString Then
            If Frm28.L5_Text <> vbNullString Then
                rs!no_rujukan_pelanggan_buyback = Frm28.L5_Text 'No. Rujukan Pembeli
            Else
                rs!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
            End If
        Else
            rs!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
        End If
    Else
        rs!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
    End If
    rs!susut_berat = "0.00"
    rs.Update

    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'### Update maklumat di bawah ke dalam maklumat barang ### - End

DATA_SAVE = 1

If DATA_SAVE = 1 Then
    
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
    
skip_carian_user:

    'User = MDI_frm1.L3_Text
    LogAct_Memory = "[" & G_LOGIN_USER & "] Belian trade in [" & Frm84.L57_Text & "] , Bil Item [" & Y & "][Trade in 2]"
    LogDate_Memory = DateTime.Date & " " & DateTime.Time$
    Call UpdateLog_Database

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If rs!Default1 = "Default" Then
            If Frm83.CB8 = 1 Then rs!no_resit_trade_in = Frm83.L12_Text + 1 'No. Resit
            'rs!NoRujukanStock = Frm83.L3_Text 'Frm83.L3_Text + 1 'No. Siri Barcode
            'If Frm83.CB9 = 1 Then
            '    rs!NoRujukanStock = Frm83.L3_Text 'No. Siri Barcode
            'ElseIf Frm83.CB10 = 1 Then
            '    rs!no_siri_gb = Frm83.L3_Text 'No. Siri Barcode
            'End If
            rs!NoRujukanSistem = Frm83.L9_Text + 1 'No. Rujukan Sistem
            rs.Update
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    
    Call Frm83_reset_list
    'Call Frm83_Initial_Setting
    'Call Frm83_Reset_Form
    'Call Frm83_Reset_After_Save
    Call Frm83_Senarai_Belian
    
'### Print Barcode ### - Start
    Note = "Print barcode bagi barang yang telah di trade in ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        'Exit Sub
    End If
    If Answer = vbYes Then
        If Frm83.CB9 = 1 Then
            Call Print_All_Barcode
        ElseIf Frm83.CB10 = 1 Then
            Call cetak_barcode_gb_all
        End If
    End If
'### Print Barcode ### - End
    
    'MsgBox "Data Telah Berjaya Disimpan.", vbInformation, "Info"
End If

End Sub
Sub Frm84_penerimaan_barang_trade_in_edit()
'On Error Resume Next
Dim Err(5)
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset

x = 0
Y = 0 '0 : Tiada Perubahan Pada Data , 1 : Ada Perubahan Pada Data
DATA_SAVE = 0
    
'###Padam Data Penjual Yang Tidak Berdaftar Ke Dalam Database### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 44_senarai_pelanggan where no_resit='" & Frm83.L12_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            rs.Delete
            rs.Update
            
        End If
        
        rs.Close
        Set rs = Nothing
'###Padam Data Penjual Yang Tidak Berdaftar Ke Dalam Database### - End
    
        Frm83_LM_No_RUJUKAN_BELIAN = Frm83.L9_Text 'No. Rujukan Belian

'###Masukkan Data Belian Ke Dalam Database Akaun Belian### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 16_gold_bar_belian where no_rujukan='" & Frm83.L9_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Frm83.DTPicker1 <> vbNullString Then
                rs!tarikh = Frm84.DTPicker1 'Tarikh Belian
            Else
                rs!tarikh = Null 'Tarikh Belian
            End If
            'If Frm83.DTPicker1 <> vbNullString Then %%%%Default : Bagi sistem ini hanya belian secara CASH dibenarkan
                rs!cara_bayaran = 0 'Cara Belian // 0 : Cash @ Bank in @ Kad Kredit @ Kad Debit , 1 : Cheque
            'Else
            '    rs!cara_bayaran = Null 'Cara Belian // 0 : Cash @ Bank in @ Kad Kredit @ Kad Debit , 1 : Cheque
            'End If
            If Frm83.L26_Text <> vbNullString Then
                rs!tunai = Format(Frm83.L26_Text, "0.00") 'Cara Bayaran : Tunai (Jumlah Keseluruhan Dengan Cukai GST)
            Else
                rs!tunai = Null 'Cara Bayaran : Tunai (Jumlah Keseluruhan Dengan Cukai GST)
            End If
            If Frm83.L11_Text <> vbNullString Then
                rs!jumlah_asal = Format(Frm83.L11_Text, "0.00") 'Jumlah Bayaran Asal (Jumlah Tanpa Cukai GST)
            Else
                rs!jumlah_asal = Null 'Jumlah Bayaran Asal (Jumlah Tanpa Cukai GST)
            End If

            If Frm83.L8_Text <> vbNullString Then
                rs!gst_value = Frm83.L8_Text 'Jumlah Cukai GST (%)
            Else
                rs!gst_value = "0" 'Jumlah Cukai GST (%)
            End If
            If Frm83.L22_Text <> vbNullString Then
                rs!gst_zr_harga = Format(Frm83.L22_Text, "0.00") 'Jumlah Bayaran Yang Dikenakan Cukai GST ZR (RM)
            Else
                rs!gst_zr_harga = "0.00" 'Jumlah Bayaran Yang Dikenakan Cukai GST ZR (RM)
            End If
            If Frm83.L23_Text <> vbNullString Then
                rs!gst_zr_cukai = Format(Frm83.L23_Text, "0.00") 'Jumlah Bayaran Cukai GST ZR (RM)
            Else
                rs!gst_zr_cukai = "0.00" 'Jumlah Bayaran Cukai GST ZR (RM)
            End If
            If Frm83.L24_Text <> vbNullString Then
                rs!gst_sr_harga = Format(Frm83.L24_Text, "0.00") 'Jumlah Bayaran Yang Dikenakan Cukai GST SR (RM)
            Else
                rs!gst_sr_harga = "0.00" 'Jumlah Bayaran Yang Dikenakan Cukai GST SR (RM)
            End If
            If Frm83.L25_Text <> vbNullString Then
                rs!gst_sr_cukai = Format(Frm83.L25_Text, "0.00") 'Jumlah Bayaran Cukai GST SR (RM)
            Else
                rs!gst_sr_cukai = "0.00" 'Jumlah Bayaran Cukai GST SR (RM)
            End If
            If Frm83.TB28 <> vbNullString Then
                rs!no_id_gst_supplier = Frm83.TB28 'No. ID GST Supplier
            Else
                rs!no_id_gst_supplier = Null 'No. ID GST Supplier
            End If
            If Frm83.TB15 <> vbNullString Then
                rs!no_resit_supplier = UCase(Frm83.TB15) 'No. Resit Dari Supplier (Jika Ada)
            Else
                rs!no_resit_supplier = Null 'No. Resit Dari Supplier (Jika Ada)
            End If
            If Frm83.TB1 <> vbNullString Then
                rs!Kod_Supplier = UCase(Frm83.TB1) 'Kod Supplier
            Else
                rs!Kod_Supplier = Null 'Kod Supplier
            End If
            If Frm83.L11_Text <> vbNullString Then
                rs!jumlah_tanpa_gst = Format(Frm83.L11_Text, "0.00") 'Jumlah Bayaran Asal (Jumlah Tanpa Cukai GST)
            Else
                rs!jumlah_tanpa_gst = Null 'Jumlah Bayaran Asal (Jumlah Tanpa Cukai GST)
            End If
            If Frm83.L26_Text <> vbNullString Then
                rs!jumlah_dengan_gst = Format(Frm83.L26_Text, "0.00") 'Jumlah Bayaran Keseluruhan (Jumlah Dengan Cukai GST)
            Else
                rs!jumlah_dengan_gst = Null 'Jumlah Bayaran Keseluruhan (Jumlah Dengan Cukai GST)
            End If
            If Frm83.CB7 = 1 Then
                rs!flag_trade_in = 0 'Flag Trade In // 0 : Tiada , 1 : Ada
            ElseIf Frm83.CB8 = 1 Then
                rs!flag_trade_in = 1 'Flag Trade In // 0 : Tiada , 1 : Ada
                rs!trade_in_status = 0 'Flag Samada Trade In Sudah Digunakan Atau Tidak , 0 : Tiada , 1 : Ada
                rs!no_resit_trade_in = Frm83.L12_Text 'No. Resit Trade In
            End If
            If Frm83.CB8 = 1 Then
                If Frm83.L39_Text <> vbNullString Then 'Pelanggan Biasa
                    rs!kategori_penjual = Frm83.L39_Text
                Else
                    rs!kategori_penjual = Null
                End If
            Else
                rs!kategori_penjual = Null
            End If
            If Frm83.CB8 = 1 Then
                If Frm28.L5_Text <> vbNullString Then
                    rs!no_rujukan_pelanggan_buyback = Frm28.L5_Text 'No. Rujukan Pembeli
                Else
                    rs!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
                End If
            Else
                rs!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
            End If
            If Frm83.CBB6 <> vbNullString Then
                Frm83_LM_EMP_NO = Split(Frm83.CBB6, "  |  ")(1)
                rs!no_pekerja = Frm83_LM_EMP_NO 'No. Pekerja
            End If
            DATA_SAVE = 1
            rs.Update
        End If
        
        rs.Close
        Set rs = Nothing
        
'###Masukkan Data Penjual Yang Tidak Berdaftar Ke Dalam Database### - Start
        If Frm28.L5_Text = vbNullString And Frm26.TB1 <> vbNullString Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 44_senarai_pelanggan", cn, adOpenKeyset, adLockOptimistic

            rs.AddNew
            rs!tarikh = Frm83.DTPicker1 'Tarikh
            rs!no_resit = Frm83.L12_Text 'No. Resit Trade In
            If Frm26.TB1 <> vbNullString Then 'Nama
                rs!Nama = UCase(Frm26.TB1)
            Else
                rs!Nama = Null
            End If
            If Frm26.TB2 <> vbNullString Then 'No. Telefon
                rs!no_tel = UCase(Frm26.TB2)
            Else
                rs!no_tel = Null
            End If
            rs!write_timestamp = Now
            rs!cawangan = G_CAWANGAN
            rs.Update
            
            rs.Close
            Set rs = Nothing
            
        End If
'###Masukkan Data Penjual Yang Tidak Berdaftar Ke Dalam Database### - End

'###Masukkan Data Belian Ke Dalam Database### - Start
'Masukkan Data Ke Dalam Database
'0 : Data Yang Baru Dipadamkan (Dikeluarkan Dari Senarai)
'1 : Tiada Perubahan Pada Data
'2 : Item Sudah Terjual (Tidak Dibenarkan Untuk Diedit/Dipadamkan
'3 : Kemasukkan Data Baru
'4 : Data Telah Diedit
'--- Yang Terlibat Dalam Urusan Ini Adalah HANYA 0 , 3 Dan 4

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_BELIAN_TEMP & "", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
            Frm83_LM_IMAGE = 0
            
            If rs!StatusItem = "3" Then
    
                Set rs1 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs1.Open "select * from Data_Database where no_siri_Produk='" & rs!no_siri_Produk & "'", cn, adOpenKeyset, adLockOptimistic
            
                If rs1.EOF Then
                    rs1.AddNew
                    If Frm83.L9_Text <> vbNullString Then
                        rs1!NoRujukanSistem = Format(Frm83_LM_No_RUJUKAN_BELIAN, "000000") 'No. Rujukan Belian
                    Else
                        rs1!NoRujukanSistem = Null 'No. Rujukan Belian
                    End If
                    If Not IsNull(rs!supplier_ID) Then
                        rs1!supplier_ID = rs!supplier_ID 'No. ID Bagi Supplier
                    Else
                        rs1!supplier_ID = Null 'No. ID Bagi Supplier
                    End If
                    If Not IsNull(rs!nama_Supplier) Then
                        rs1!nama_Supplier = rs!nama_Supplier 'Nama Supplier
                    Else
                        rs1!nama_Supplier = Null 'Nama Supplier
                    End If
                    If Not IsNull(rs!Kod_Supplier) Then
                        rs1!Kod_Supplier = rs!Kod_Supplier 'Kod Supplier
                    Else
                        rs1!Kod_Supplier = Null 'Kod Supplier
                    End If
                    If Not IsNull(rs!purity_ID) Then
                        rs1!purity_ID = rs!purity_ID 'No. ID Bagi Purity
                    Else
                        rs1!purity_ID = Null 'No. ID Bagi Purity
                    End If
                    If Not IsNull(rs!purity) Then
                        rs1!purity = rs!purity 'Purity
                    Else
                        rs1!purity = Null 'Purity
                    End If
                    If Not IsNull(rs!kod_Purity) Then
                        rs1!kod_Purity = rs!kod_Purity 'Kod Purity
                    Else
                        rs1!kod_Purity = Null 'Kod Purity
                    End If
                    If Not IsNull(rs!kategori_produk_ID) Then
                        rs1!kategori_produk_ID = rs!kategori_produk_ID 'No. ID Bagi Purity
                    Else
                        rs1!kategori_produk_ID = Null 'No. ID Bagi Purity
                    End If
                    If Not IsNull(rs!kategori_Produk) Then
                        rs1!kategori_Produk = rs!kategori_Produk 'Kategori Produk
                    Else
                        rs1!kategori_Produk = Null 'Kategori Produk
                    End If
                    If Not IsNull(rs!Kod_Kategori_Produk) Then
                        rs1!Kod_Kategori_Produk = rs!Kod_Kategori_Produk 'Kod Kategori Produk
                    Else
                        rs1!Kod_Kategori_Produk = Null 'Kod Kategori Produk
                    End If
                    If Not IsNull(rs!Barcode) Then
                        rs1!Barcode = rs!Barcode 'No. Turutan Barcode
                    Else
                        rs1!Barcode = Null 'No. Turutan Barcode
                    End If
                    If Not IsNull(rs!no_siri_Produk) Then
                        rs1!no_siri_Produk = rs!no_siri_Produk 'No. Siri Produk
                        Frm83_LM_BARCODE = rs!no_siri_Produk
                    Else
                        rs1!no_siri_Produk = Null 'No. Siri Produk
                    End If
                    If Not IsNull(rs!Berat) Then
                        rs1!Berat = rs!Berat 'Berat
                    Else
                        rs1!Berat = Null 'Berat
                    End If
                    If Not IsNull(rs!beza_berat) Then
                        rs1!beza_berat = rs!beza_berat 'Beza Berat (Baki Berat Bagi Jualan Potong)
                    Else
                        rs1!beza_berat = Null 'Beza Berat (Baki Berat Bagi Jualan Potong)
                    End If
                    If Not IsNull(rs!UPAH) Then
                        rs1!UPAH = rs!UPAH 'Upah (RM) - Upah Dari Supplier
                    Else
                        rs1!UPAH = Null 'Upah (RM) - Upah Dari Supplier
                    End If
                    If Not IsNull(rs!Upah30) Then
                        rs1!Upah30 = rs!Upah30 'Upah (RM) - Upah Dari Supplier - Sudah Dinaikkan 30%
                    Else
                        rs1!Upah30 = Null 'Upah (RM) - Upah Dari Supplier - Sudah Dinaikkan 30%
                    End If
                    If Not IsNull(rs!riyal) Then
                        rs1!riyal = rs!riyal 'Berat Dalam Riyal
                    Else
                        rs1!riyal = Null 'Berat Dalam Riyal
                    End If
                    If Not IsNull(rs!kos_Belian_Gram) Then
                        rs1!kos_Belian_Gram = Format(rs!kos_Belian_Gram, "0.00") 'Harga Semasa Dari Supplier (RM/g)
                    Else
                        rs1!kos_Belian_Gram = Null 'Harga Semasa Dari Supplier (RM/g)
                    End If
                    If Not IsNull(rs!kos_Belian_Item) Then
                        rs1!kos_Belian_Item = Format(rs!kos_Belian_Item, "0.00") 'Kos belian item (campur upah)
                    Else
                        rs1!kos_Belian_Item = Null 'Kos belian item (campur upah)
                    End If
                    If Not IsNull(rs!Spread) Then
                        rs1!SpreadValue = Format(rs!Spread, "0.00") 'Spread Bagi Belian Trade In (%)
                    Else
                        rs1!SpreadValue = Null 'Spread Bagi Belian Trade In (%)
                    End If
                    If Not IsNull(rs!harga_lepas_spread) Then
                        rs1!harga_lepas_spread = Format(rs!harga_lepas_spread, "0.00") 'Harga Belian Selepas Spread (RM)
                    Else
                        rs1!harga_lepas_spread = Null 'Harga Belian Selepas Spread (RM)
                    End If
                    If Not IsNull(rs!adjustment) Then
                        rs1!adjustment = Format(rs!adjustment, "0.00") 'Adjustment (RM)
                    Else
                        rs1!adjustment = Null 'Adjustment (RM)
                    End If
                    If Not IsNull(rs!kos_item_tanpa_tax) Then
                        rs1!kos_item_tanpa_tax = Format(rs!kos_item_tanpa_tax, "0.00") 'Adjustment (RM)
                    Else
                        rs1!kos_item_tanpa_tax = Null 'Adjustment (RM)
                    End If
                    If Not IsNull(rs!cara_Belian) Then
                        rs1!cara_Belian = rs!cara_Belian 'Cara Belian , 0 : Tunai , 1 : Cek , 2 : Tukaran Barang
                    Else
                        rs1!cara_Belian = Null 'Cara Belian , 0 : Tunai , 1 : Cek , 2 : Tukaran Barang
                    End If
                    If Not IsNull(rs!dimension_Panjang) Then
                        rs1!dimension_Panjang = rs!dimension_Panjang 'Dimension : Panjang
                    Else
                        rs1!dimension_Panjang = Null 'Dimension : Panjang
                    End If
                    If Not IsNull(rs!dimension_Lebar) Then
                        rs1!dimension_Lebar = rs!dimension_Lebar 'Dimension : Lebar
                    Else
                        rs1!dimension_Lebar = Null 'Dimension : Lebar
                    End If
                    If Not IsNull(rs!dimension_Saiz) Then
                        rs1!dimension_Saiz = rs!dimension_Saiz 'Dimension : Size
                    Else
                        rs1!dimension_Saiz = Null 'Dimension : Size
                    End If
                    
                     If Not IsNull(rs!code1) Then 'Code 1
                        rs1!code1 = rs!code1
                    Else
                        rs1!code1 = Null
                    End If
                     If Not IsNull(rs!code2) Then 'Code 1
                        rs1!code2 = rs!code2
                    Else
                        rs1!code2 = Null
                    End If
                    
                    If Not IsNull(rs!harga_Per_Gram_Item) Then
                        rs1!harga_Per_Gram_Item = Format(rs!harga_Per_Gram_Item, "0.00") 'Kos belian per gram (average : sudah dicampur dengan upah) (Total)
                    Else
                        rs1!harga_Per_Gram_Item = Null 'Kos belian per gram (average : sudah dicampur dengan upah) (Total)
                    End If
                    If Not IsNull(rs!dulang) Then
                        rs1!dulang = rs!dulang 'Dulang
                    Else
                        rs1!dulang = Null 'Dulang
                    End If
                    If Not IsNull(rs!no_cert) Then
                        rs1!no_cert = rs!no_cert 'No. Cert
                    Else
                        rs1!no_cert = Null 'No. Cert
                    End If
                    If Frm83.CB8 = 1 Then rs1!bill_No_Trade_In = Frm83.L12_Text '"BK" & Format(Frm83.L12_Text, "000000") 'No. Resit Trade In
                    
'Status Item
'0 : In Stock
'1:  Sold
'2 : In Stock - Potong
'3:  Sold -Potong
'4:  Tempahan
'5:  Ansuran
'6:  Ar -Rahnu
'7:  ETA

                    If Not IsNull(rs!StatusItem) Then
                        rs1!StatusItem = 10 'Status Item
                    Else
                        rs1!StatusItem = Null 'Status Item
                    End If
                    If Not IsNull(rs!Upah_Jualan) Then
                        rs1!Upah_Jualan = Format(rs!Upah_Jualan, "0.00") 'Upah Kepada Pelanggan
                    Else
                        rs1!Upah_Jualan = Null 'Upah Kepada Pelanggan
                    End If
                    If Not IsNull(rs!Upah_Member) Then
                        rs1!Upah_Member = Format(rs!Upah_Member, "0.00") 'Upah Kepada Ahli / Member
                    Else
                        rs1!Upah_Member = Null 'Upah Kepada Ahli / Member
                    End If
                    If Not IsNull(rs!Upah_RAF) Then
                        rs1!Upah_RAF = Format(rs!Upah_RAF, "0.00") 'Upah Kepada RAF
                    Else
                        rs1!Upah_RAF = Null 'Upah Kepada RAF
                    End If
                    If Not IsNull(rs!Upah_Pengedar) Then
                        rs1!Upah_Pengedar = Format(rs!Upah_Pengedar, "0.00") 'Upah Kepada Pengedar
                    Else
                        rs1!Upah_Pengedar = Null 'Upah Kepada Pengedar
                    End If
                    If Not IsNull(rs!code_Supplier) Then
                        rs1!code_Supplier = Format(rs!code_Supplier, "0.00") 'Harga Jualan Kepada Pelanggan
                    Else
                        rs1!code_Supplier = Null 'Harga Jualan Kepada Pelanggan
                    End If
                    If Not IsNull(rs!HargaJualan_Member) Then
                        rs1!HargaJualan_Member = Format(rs!HargaJualan_Member, "0.00") 'Harga Jualan Kepada Ahli / Member
                    Else
                        rs1!HargaJualan_Member = Null 'Harga Jualan Kepada Ahli / Member
                    End If
                    If Not IsNull(rs!HargaJualan_Pengedar) Then
                        rs1!HargaJualan_Pengedar = Format(rs!HargaJualan_Pengedar, "0.00") 'Harga Jualan Kepada Pengedar
                    Else
                        rs1!HargaJualan_Pengedar = Null 'Harga Jualan Kepada Pengedar
                    End If
                    
                    If Not IsNull(rs!upah_normal_dealer) Then
                        rs1!upah_normal_dealer = Format(rs!upah_normal_dealer, "0.00") 'Upah Kepada N.Dealer
                    Else
                        rs1!upah_normal_dealer = Null 'Upah Kepada N.Dealer
                    End If
                    If Not IsNull(rs!upah_master_dealer) Then
                        rs1!upah_master_dealer = Format(rs!upah_master_dealer, "0.00") 'Upah Kepada M.Dealer
                    Else
                        rs1!upah_master_dealer = Null 'Upah Kepada M.Dealer
                    End If
                    If Not IsNull(rs!HargaJualan_RAF) Then
                        rs1!HargaJualan_RAF = Format(rs!HargaJualan_RAF, "0.00") 'Harga Jualan Kepada RAF
                    Else
                        rs1!HargaJualan_RAF = Null 'Harga Jualan Kepada Pengedar
                    End If
                    If Not IsNull(rs!hargajualan_normal_dealer) Then
                        rs1!hargajualan_normal_dealer = Format(rs!hargajualan_normal_dealer, "0.00") 'Harga Jualan Kepada N.Dealer
                    Else
                        rs1!hargajualan_normal_dealer = Null 'Harga Jualan Kepada N.Dealer
                    End If
                    If Not IsNull(rs!hargajualan_master_dealer) Then
                        rs1!hargajualan_master_dealer = Format(rs!hargajualan_master_dealer, "0.00") 'Harga Jualan Kepada M.Dealer
                    Else
                        rs1!hargajualan_master_dealer = Null 'Harga Jualan Kepada M.Dealer
                    End If
                    
                    If Not IsNull(rs!remarks) Then
                        rs1!remarks = rs!remarks 'Remarks
                    Else
                        rs1!remarks = Null 'Remarks
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
                    If Not IsNull(rs!harga_item) Then
                        rs1!harga_item = Format(rs!harga_item, "0.00") 'Jumlah Harga Keseluruhan Termasuk GST (RM)
                    Else
                        rs1!harga_item = Null 'Jumlah Harga Keseluruhan Termasuk GST (RM)
                    End If
                    rs!write_timestamp = Now 'Tarikh & Masa Data Dimasukkan
                    If Not IsNull(rs!kadar_gst) Then
                        rs1!kadar_gst = rs!kadar_gst 'Kadar Cukai GST (%)
                    Else
                        rs1!kadar_gst = Null 'Kadar Cukai GST (%)
                    End If
'Cara Penerimaan Stok
'0:  BK
'1:  Barang Permata
'2 : Trade In BK
'3 : Trade In Barang Permata
                    If Not IsNull(rs!jenis) Then
                        rs1!receiving_Status = rs!jenis 'Cara Penerimaan Stok
                    Else
                        rs1!receiving_Status = Null 'Cara Penerimaan Stok
                    End If
                    If Not IsNull(rs!jenis) Then
                        rs1!receiving_Status = rs!jenis 'Cara Penerimaan Stok
                    Else
                        rs1!receiving_Status = Null 'Cara Penerimaan Stok
                    End If
                    rs1!tarikh_belian = Frm83.DTPicker1 'Tarikh belian dibuat
                    If Frm83.TB15 <> vbNullString Then
                        rs1!bill_No_Belian = UCase(Frm83.TB15) 'No. Resit Dari Supplier (Jika Ada)
                    Else
                        rs1!bill_No_Belian = Null 'No. Resit Dari Supplier (Jika Ada)
                    End If
                    If Frm83.CB8 = 1 Then
                        If Frm28.L5_Text <> vbNullString Then
                            rs1!no_rujukan_pelanggan_buyback = Frm28.L5_Text 'No. Rujukan Pembeli
                        Else
                            rs1!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
                        End If
                    Else
                        rs1!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
                    End If
                    If Not IsNull(rs!harga_tanpa_gst) Then 'Harga modal barang ini tanpa GST (RM)
                        rs1!harga_tanpa_gst = Format(rs!harga_tanpa_gst, "0.00")
                    Else
                        rs1!harga_tanpa_gst = Null
                    End If
                    If Not IsNull(rs!gst_included) Then ''0 : Harga tidak termasuk GST , 1 : Harga termasuk GST
                        rs1!gst_included = rs!gst_included
                    Else
                        rs1!gst_included = Null
                    End If
                    If Not IsNull(rs!flag_image) Then
                        If rs!flag_image = 1 Then
                            Frm83_LM_IMAGE = 1
                            rs1!flag_image = 1
                        Else
                            rs1!flag_image = 0
                        End If
                    End If
                    
                    rs1.Update
                    DATA_SAVE = 1
                End If
                
                rs1.Close
                Set rs1 = Nothing
                
            ElseIf rs!StatusItem = "4" Then '0 : Padam Data , 1 : Aktif (Tiada Perubahan) , 2 : Sudah Terjual , 3 : Data Baru , 4 : Data Diedit
            
                Set rs1 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs1.Open "select * from Data_Database where id='" & rs!id_database & "'", cn, adOpenKeyset, adLockOptimistic
            
                If Not rs1.EOF Then
                    If Frm83.L9_Text <> vbNullString Then
                        rs1!NoRujukanSistem = Format(Frm83_LM_No_RUJUKAN_BELIAN, "000000") 'No. Rujukan Belian
                    Else
                        rs1!NoRujukanSistem = Null 'No. Rujukan Belian
                    End If
                    If Not IsNull(rs!supplier_ID) Then
                        rs1!supplier_ID = rs!supplier_ID 'No. ID Bagi Supplier
                    Else
                        rs1!supplier_ID = Null 'No. ID Bagi Supplier
                    End If
                    If Not IsNull(rs!nama_Supplier) Then
                        rs1!nama_Supplier = rs!nama_Supplier 'Nama Supplier
                    Else
                        rs1!nama_Supplier = Null 'Nama Supplier
                    End If
                    If Not IsNull(rs!Kod_Supplier) Then
                        rs1!Kod_Supplier = rs!Kod_Supplier 'Kod Supplier
                    Else
                        rs1!Kod_Supplier = Null 'Kod Supplier
                    End If
                    If Not IsNull(rs!purity_ID) Then
                        rs1!purity_ID = rs!purity_ID 'No. ID Bagi Purity
                    Else
                        rs1!purity_ID = Null 'No. ID Bagi Purity
                    End If
                    If Not IsNull(rs!purity) Then
                        rs1!purity = rs!purity 'Purity
                    Else
                        rs1!purity = Null 'Purity
                    End If
                    If Not IsNull(rs!kod_Purity) Then
                        rs1!kod_Purity = rs!kod_Purity 'Kod Purity
                    Else
                        rs1!kod_Purity = Null 'Kod Purity
                    End If
                    If Not IsNull(rs!kategori_produk_ID) Then
                        rs1!kategori_produk_ID = rs!kategori_produk_ID 'No. ID Bagi Purity
                    Else
                        rs1!kategori_produk_ID = Null 'No. ID Bagi Purity
                    End If
                    If Not IsNull(rs!kategori_Produk) Then
                        rs1!kategori_Produk = rs!kategori_Produk 'Kategori Produk
                    Else
                        rs1!kategori_Produk = Null 'Kategori Produk
                    End If
                    If Not IsNull(rs!Kod_Kategori_Produk) Then
                        rs1!Kod_Kategori_Produk = rs!Kod_Kategori_Produk 'Kod Kategori Produk
                    Else
                        rs1!Kod_Kategori_Produk = Null 'Kod Kategori Produk
                    End If
                    If Not IsNull(rs!Barcode) Then
                        rs1!Barcode = rs!Barcode 'No. Turutan Barcode
                    Else
                        rs1!Barcode = Null 'No. Turutan Barcode
                    End If
                    If Not IsNull(rs!no_siri_Produk) Then
                        rs1!no_siri_Produk = rs!no_siri_Produk 'No. Siri Produk
                        Frm83_LM_BARCODE = rs!no_siri_Produk
                    Else
                        rs1!no_siri_Produk = Null 'No. Siri Produk
                    End If
                    If Not IsNull(rs!Berat) Then
                        rs1!Berat = Format(rs!Berat, "0.00") 'Berat
                    Else
                        rs1!Berat = Null 'Berat
                    End If
                    If Not IsNull(rs!beza_berat) Then
                        rs1!beza_berat = Format(rs!beza_berat, "0.00") 'Beza Berat (Baki Berat Bagi Jualan Potong)
                    Else
                        rs1!beza_berat = Null 'Beza Berat (Baki Berat Bagi Jualan Potong)
                    End If
                    If Not IsNull(rs!UPAH) Then
                        rs1!UPAH = rs!UPAH 'Upah (RM) - Upah Dari Supplier
                    Else
                        rs1!UPAH = Null 'Upah (RM) - Upah Dari Supplier
                    End If
                    If Not IsNull(rs!Upah30) Then
                        rs1!Upah30 = rs!Upah30 'Upah (RM) - Upah Dari Supplier - Sudah Dinaikkan 30%
                    Else
                        rs1!Upah30 = Null 'Upah (RM) - Upah Dari Supplier - Sudah Dinaikkan 30%
                    End If
                    If Not IsNull(rs!riyal) Then
                        rs1!riyal = rs!riyal 'Berat Dalam Riyal
                    Else
                        rs1!riyal = Null 'Berat Dalam Riyal
                    End If
                    If Not IsNull(rs!kos_Belian_Gram) Then
                        rs1!kos_Belian_Gram = Format(rs!kos_Belian_Gram, "0.00") 'Harga Semasa Dari Supplier (RM/g)
                    Else
                        rs1!kos_Belian_Gram = Null 'Harga Semasa Dari Supplier (RM/g)
                    End If
                    If Not IsNull(rs!kos_Belian_Item) Then
                        rs1!kos_Belian_Item = Format(rs!kos_Belian_Item, "0.00") 'Kos belian item (campur upah)
                    Else
                        rs1!kos_Belian_Item = Null 'Kos belian item (campur upah)
                    End If
                    If Not IsNull(rs!Spread) Then
                        rs1!SpreadValue = Format(rs!Spread, "0.00") 'Spread Bagi Belian Trade In (%)
                    Else
                        rs1!SpreadValue = Null 'Spread Bagi Belian Trade In (%)
                    End If
                    If Not IsNull(rs!harga_lepas_spread) Then
                        rs1!harga_lepas_spread = Format(rs!harga_lepas_spread, "0.00") 'Harga Belian Selepas Spread (RM)
                    Else
                        rs1!harga_lepas_spread = Null 'Harga Belian Selepas Spread (RM)
                    End If
                    If Not IsNull(rs!adjustment) Then
                        rs1!adjustment = Format(rs!adjustment, "0.00") 'Adjustment (RM)
                    Else
                        rs1!adjustment = Null 'Adjustment (RM)
                    End If
                    If Not IsNull(rs!kos_item_tanpa_tax) Then
                        rs1!kos_item_tanpa_tax = Format(rs!kos_item_tanpa_tax, "0.00") 'Adjustment (RM)
                    Else
                        rs1!kos_item_tanpa_tax = Null 'Adjustment (RM)
                    End If
                    If Not IsNull(rs!cara_Belian) Then
                        rs1!cara_Belian = rs!cara_Belian 'Cara Belian , 0 : Tunai , 1 : Cek , 2 : Tukaran Barang
                    Else
                        rs1!cara_Belian = Null 'Cara Belian , 0 : Tunai , 1 : Cek , 2 : Tukaran Barang
                    End If
                    If Not IsNull(rs!dimension_Panjang) Then
                        rs1!dimension_Panjang = rs!dimension_Panjang 'Dimension : Panjang
                    Else
                        rs1!dimension_Panjang = Null 'Dimension : Panjang
                    End If
                    If Not IsNull(rs!dimension_Lebar) Then
                        rs1!dimension_Lebar = rs!dimension_Lebar 'Dimension : Lebar
                    Else
                        rs1!dimension_Lebar = Null 'Dimension : Lebar
                    End If
                    If Not IsNull(rs!dimension_Saiz) Then
                        rs1!dimension_Saiz = rs!dimension_Saiz 'Dimension : Size
                    Else
                        rs1!dimension_Saiz = Null 'Dimension : Size
                    End If
                    If Not IsNull(rs!code1) Then 'Code 1
                        rs1!code1 = rs!code1
                    Else
                        rs1!code1 = Null
                    End If
                     If Not IsNull(rs!code2) Then 'Code 1
                        rs1!code2 = rs!code2
                    Else
                        rs1!code2 = Null
                    End If
                    If Not IsNull(rs!harga_Per_Gram_Item) Then
                        rs1!harga_Per_Gram_Item = Format(rs!harga_Per_Gram_Item, "0.00") 'Kos belian per gram (average : sudah dicampur dengan upah) (Total)
                    Else
                        rs1!harga_Per_Gram_Item = Null 'Kos belian per gram (average : sudah dicampur dengan upah) (Total)
                    End If
                    If Not IsNull(rs!dulang) Then
                        rs1!dulang = rs!dulang 'Dulang
                    Else
                        rs1!dulang = Null 'Dulang
                    End If
                    If Not IsNull(rs!no_cert) Then
                        rs1!no_cert = rs!no_cert 'No. Cert
                    Else
                        rs1!no_cert = Null 'No. Cert
                    End If
                    If Frm83.CB8 = 1 Then rs1!bill_No_Trade_In = Frm83.L12_Text 'No. Resit Trade In
                    
'Status Item
'0 : In Stock
'1:  Sold
'2 : In Stock - Potong
'3:  Sold -Potong
'4:  Tempahan
'5:  Ansuran
'6:  Ar -Rahnu
'7:  ETA

                    If Not IsNull(rs!StatusItem) Then
                        rs1!StatusItem = 10 'Status Item
                    Else
                        rs1!StatusItem = Null 'Status Item
                    End If
                    If Not IsNull(rs!Upah_Jualan) Then
                        rs1!Upah_Jualan = Format(rs!Upah_Jualan, "0.00") 'Upah Kepada Pelanggan
                    Else
                        rs1!Upah_Jualan = Null 'Upah Kepada Pelanggan
                    End If
                    If Not IsNull(rs!Upah_Member) Then
                        rs1!Upah_Member = Format(rs!Upah_Member, "0.00") 'Upah Kepada Ahli / Member
                    Else
                        rs1!Upah_Member = Null 'Upah Kepada Ahli / Member
                    End If
                    If Not IsNull(rs!Upah_RAF) Then
                        rs1!Upah_RAF = Format(rs!Upah_RAF, "0.00") 'Upah Kepada RAF
                    Else
                        rs1!Upah_RAF = Null 'Upah Kepada RAF
                    End If
                    If Not IsNull(rs!Upah_Pengedar) Then
                        rs1!Upah_Pengedar = Format(rs!Upah_Pengedar, "0.00") 'Upah Kepada Pengedar
                    Else
                        rs1!Upah_Pengedar = Null 'Upah Kepada Pengedar
                    End If
                    
                    If Not IsNull(rs!code_Supplier) Then
                        rs1!code_Supplier = Format(rs!code_Supplier, "0.00") 'Harga Jualan Kepada Pelanggan
                    Else
                        rs1!code_Supplier = Null 'Harga Jualan Kepada Pelanggan
                    End If
                    If Not IsNull(rs!HargaJualan_Member) Then
                        rs1!HargaJualan_Member = Format(rs!HargaJualan_Member, "0.00") 'Harga Jualan Kepada Ahli / Member
                    Else
                        rs1!HargaJualan_Member = Null 'Harga Jualan Kepada Ahli / Member
                    End If
                    If Not IsNull(rs!HargaJualan_Pengedar) Then
                        rs1!HargaJualan_Pengedar = Format(rs!HargaJualan_Pengedar, "0.00") 'Harga Jualan Kepada Pengedar
                    Else
                        rs1!HargaJualan_Pengedar = Null 'Harga Jualan Kepada Pengedar
                    End If
                    

                    If Not IsNull(rs!upah_normal_dealer) Then
                        rs1!upah_normal_dealer = Format(rs!upah_normal_dealer, "0.00") 'Upah Kepada N.Dealer
                    Else
                        rs1!upah_normal_dealer = Null 'Upah Kepada N.Dealer
                    End If
                    If Not IsNull(rs!upah_master_dealer) Then
                        rs1!upah_master_dealer = Format(rs!upah_master_dealer, "0.00") 'Upah Kepada M.Dealer
                    Else
                        rs1!upah_master_dealer = Null 'Upah Kepada M.Dealer
                    End If
                    If Not IsNull(rs!HargaJualan_RAF) Then
                        rs1!HargaJualan_RAF = Format(rs!HargaJualan_RAF, "0.00") 'Harga Jualan Kepada RAF
                    Else
                        rs1!HargaJualan_RAF = Null 'Harga Jualan Kepada Pengedar
                    End If
                    If Not IsNull(rs!hargajualan_normal_dealer) Then
                        rs1!hargajualan_normal_dealer = Format(rs!hargajualan_normal_dealer, "0.00") 'Harga Jualan Kepada N.Dealer
                    Else
                        rs1!hargajualan_normal_dealer = Null 'Harga Jualan Kepada N.Dealer
                    End If
                    If Not IsNull(rs!hargajualan_master_dealer) Then
                        rs1!hargajualan_master_dealer = Format(rs!hargajualan_master_dealer, "0.00") 'Harga Jualan Kepada M.Dealer
                    Else
                        rs1!hargajualan_master_dealer = Null 'Harga Jualan Kepada M.Dealer
                    End If
                    
                    If Not IsNull(rs!remarks) Then
                        rs1!remarks = rs!remarks 'Remarks
                    Else
                        rs1!remarks = Null 'Remarks
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
                    If Not IsNull(rs!harga_item) Then
                        rs1!harga_item = Format(rs!harga_item, "0.00") 'Jumlah Harga Keseluruhan Termasuk GST (RM)
                    Else
                        rs1!harga_item = Null 'Jumlah Harga Keseluruhan Termasuk GST (RM)
                    End If
                    rs!write_timestamp = Now 'Tarikh & Masa Data Dimasukkan
                    If Not IsNull(rs!kadar_gst) Then
                        rs1!kadar_gst = rs!kadar_gst 'Kadar Cukai GST (%)
                    Else
                        rs1!kadar_gst = Null 'Kadar Cukai GST (%)
                    End If
'Cara Penerimaan Stok
'0:  BK
'1:  Barang Permata
'2 : Trade In BK
'3 : Trade In Barang Permata
                    If Not IsNull(rs!jenis) Then
                        rs1!receiving_Status = rs!jenis 'Cara Penerimaan Stok
                    Else
                        rs1!receiving_Status = Null 'Cara Penerimaan Stok
                    End If
                    If Not IsNull(rs!jenis) Then
                        rs1!receiving_Status = rs!jenis 'Cara Penerimaan Stok
                    Else
                        rs1!receiving_Status = Null 'Cara Penerimaan Stok
                    End If
                    rs1!tarikh_belian = Frm83.DTPicker1 'Tarikh belian dibuat
                    If Frm83.TB15 <> vbNullString Then
                        rs1!bill_No_Belian = UCase(Frm83.TB15) 'No. Resit Dari Supplier (Jika Ada)
                    Else
                        rs1!bill_No_Belian = Null 'No. Resit Dari Supplier (Jika Ada)
                    End If
                    If Frm83.CB8 = 1 Then
                        If Frm28.L5_Text <> vbNullString Then
                            rs1!no_rujukan_pelanggan_buyback = Frm28.L5_Text 'No. Rujukan Pembeli
                        Else
                            rs1!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
                        End If
                    Else
                        rs1!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
                    End If
                    If Not IsNull(rs!harga_tanpa_gst) Then 'Harga modal barang ini tanpa GST (RM)
                        rs1!harga_tanpa_gst = Format(rs!harga_tanpa_gst, "0.00")
                    Else
                        rs1!harga_tanpa_gst = Null
                    End If
                    If Not IsNull(rs!gst_included) Then ''0 : Harga tidak termasuk GST , 1 : Harga termasuk GST
                        rs1!gst_included = rs!gst_included
                    Else
                        rs1!gst_included = Null
                    End If
                    If Not IsNull(rs!flag_image) Then
                        If rs!flag_image = 1 Then
                            Frm83_LM_IMAGE = 1
                            rs1!flag_image = 1
                        Else
                            rs1!flag_image = 0
                        End If
                    End If
                    
                    rs1.Update
                    DATA_SAVE = 1
                End If

                rs1.Close
                Set rs1 = Nothing
                
            ElseIf rs!StatusItem = "5" Then
            
                Set rs1 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs1.Open "select * from Data_Database where no_siri_Produk='" & rs!no_siri_Produk & "'", cn, adOpenKeyset, adLockOptimistic
                    
                If Not rs1.EOF Then
                    rs1.Delete
                    rs1.Update
                    DATA_SAVE = 1
                End If
                    
                rs1.Close
                Set rs1 = Nothing
            End If

            If Frm83_LM_IMAGE = 1 Then
                Set rs2 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs2.Open "select * from 2_image_item where barcode='" & Frm83_LM_BARCODE & "'", cn, adOpenKeyset, adLockOptimistic
            
                If Not rs2.EOF Then
                    rs2!Image = rs!Image
                    rs2!write_timestamp = Now
                    rs2.Update
                Else
                    rs2.AddNew
                    rs2!Barcode = Frm83_LM_BARCODE
                    rs2!Image = rs!Image
                    rs2!write_timestamp = Now
                    rs2.Update
                End If
                
                rs2.Close
                Set rs2 = Nothing
            End If

            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        If DATA_SAVE = 1 Then
            If Frm83.TB15 <> vbNullString Then
                Frm83_LM_No_INVOICE_SUPPLIER = UCase(Frm83.TB15)
            Else
                Frm83_LM_No_INVOICE_SUPPLIER = Null
            End If
            
            '#### Update Maklumat Dulang Dalam Table Data_Database #### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

            strsql = "UPDATE Data_Database set bill_No_Belian='" & UCase(Frm83.TB15) & "'," _
            & "tarikh_belian='" & Frm83.DTPicker1 & "'" _
            & "WHERE NoRujukanSistem='" & Frm83.L9_Text & "'"
            
            Set rs = cn.Execute(strsql)
            Set rs = Nothing
            '#### Update Maklumat Dulang Dalam Table Data_Database #### - End
        
            If Frm83.CBB6 <> vbNullString Then
                Frm83_LM_EMP_NAME = Split(Frm83.CBB6, "  |  ")(0)
            End If
        
            'User = MDI_frm1.L3_Text
            If Frm83.CB7 = 1 Then LogAct_Memory = "[" & Frm83_LM_EMP_NAME & "] Edit Data Stok [" & Frm83.L9_Text & "]."
            If Frm83.CB8 = 1 Then LogAct_Memory = "[" & Frm83_LM_EMP_NAME & "] Edit Data Trade In [" & Frm83.L12_Text & "]."
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If rs!Default1 = "Default" Then
                    'If IsNumeric(Frm83.L3_Text) Then rs!NoRujukanStock = Frm83.L3_Text 'No. Siri Barcode
                    
                    If Frm83.CB9 = 1 Then
                        rs!NoRujukanStock = Frm83.L3_Text 'No. Siri Barcode
                    ElseIf Frm83.CB10 = 1 Then
                        rs!no_siri_gb = Frm83.L3_Text 'No. Siri Barcode
                    End If
                    
                    rs.Update
                End If
            End If
            
            rs.Close
            Set rs = Nothing
                
            Note = "Data Telah Berjaya Disimpan." & vbCrLf & _
                    "Refresh Data Anda ?"

            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbNo Then
                Frm85.Show
                Unload Frm83
                Unload Frm26
                Unload Frm27
                Unload Frm28
            End If
            If Answer = vbYes Then
                GM_NEXT_PREV = 2
                
                If Frm101.L33_Text = 0 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                    If Frm101.CB2 = 1 Then 'Report Belian
                        Call Frm85_Header_Report_Belian
                        Call Frm85_report_belian_page
                    End If
                    If Frm101.CB4 = 1 Then 'Report Buyback
                        Frm85_Header_Report_Buyback
                        Call Frm85_report_buyback_page
                    End If
                    If Frm101.CB11 = 1 Then 'Report Belian Gold Bar
                        Call Frm85_Header_Report_belian_gb
                        Call Frm85_report_belian_gb_page
                    End If
                    If Frm101.CB12 = 1 Then 'Report Buyback Gold Bar
                        Frm85_Header_Report_Buyback
                        Call Frm85_report_buyback_gb_page
                    End If
                ElseIf Frm101.L33_Text = 1 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                    Call Frm85_Header_Report_Belian
                    'Call Frm85_search_berat
                    Call Frm85_search_berat_page
                ElseIf Frm101.L33_Text = 3 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                    Call Frm85_Header_Report_buyback_gb
                    Call Frm85_carian_buyback_page
                ElseIf Frm101.L33_Text = 4 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                    Call Frm85_Header_Report_Belian
                    Call Frm85_search_invoice_supplier_page
                ElseIf Frm101.L33_Text = 5 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                    Call Frm85_Header_Report_Belian
                    Call Frm85_report_belian_barcode
                ElseIf Frm101.L33_Text = 6 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                    Call Frm85_Header_Report_Belian
                    Call Frm85_report_buyback_barcode
                ElseIf Frm101.L33_Text = 7 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                    Call Frm85_Header_Report_belian_gb
                    Call Frm85_report_belian_gb_barcode
                ElseIf Frm101.L33_Text = 8 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                    Call Frm85_Header_Report_belian_gb
                    Call Frm85_report_buyback_gb_barcode
                End If
                
                Frm85.Show
                Unload Frm83
                Unload Frm26
                Unload Frm27
                Unload Frm28
            End If
        End If

End Sub
Sub Frm84_recall_trade_in_data()
'on error resume next
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset

Frm83.CB9 = 0
Frm83.CB10 = 0

DATA_FOUND = 0
Frm_LM_DATA_PENJUAL_BUYBACK = 0
Frm85_LM_No_PENJUAL = vbNullString

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from Data_Database where NoRujukanSistem='" & Frm83.L9_Text & "' order by ID ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Not IsNull(rs!receiving_Status) Then
        
        If rs!receiving_Status = 0 Or rs!receiving_Status = 1 Or rs!receiving_Status = 2 Or rs!receiving_Status = 3 Or rs!receiving_Status = 6 Or rs!receiving_Status = 7 Then
            Frm83.CB9 = 1
            Frm83.CB10 = 0
        End If
        If rs!receiving_Status = 4 Or rs!receiving_Status = 5 Or rs!receiving_Status = 8 Then
            Frm83.CB9 = 0
            Frm83.CB10 = 1
        End If
        
    End If
        
End If

rs.Close
Set rs = Nothing

'### Masukkan maklumat data barang ke dalam table #data_database ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "insert into " & G_BELIAN_TEMP & "(id_database,terminal,tarikh_belian,no_siri_produk,bill_no_belian,supplier_ID,nama_Supplier,Kod_Supplier,purity_ID,purity,kod_Purity,kategori_produk_ID,kategori_Produk,Kod_Kategori_Produk,Berat,beza_berat,UPAH,Upah30,riyal,kos_Belian_Gram,kos_Belian_Item,Spread,harga_lepas_spread,adjustment,kos_item_tanpa_tax,cara_Belian,dimension_Panjang,dimension_Lebar,dimension_Saiz,code1,code2,harga_Per_Gram_Item,dulang,no_cert,gst_barang_atau_upah,StatusItem,Upah_Jualan,Upah_Member,Upah_RAF,Upah_Pengedar,code_Supplier,HargaJualan_Member,HargaJualan_Pengedar,upah_normal_dealer,upah_master_dealer,HargaJualan_RAF,hargajualan_normal_dealer,hargajualan_master_dealer,remarks,gst_ari_nashi,kadar_gst,jumlah_gst,harga_item,jenis,harga_tanpa_gst,gst_included,jenis_trade_in,flag_upah,upah_per_gram,flag_image,no_id_gst)" & _
                    "select ID,terminal,tarikh_belian,no_siri_produk,bill_no_belian,supplier_ID,nama_Supplier,Kod_Supplier,purity_ID,purity,kod_Purity,kategori_produk_ID,kategori_Produk,Kod_Kategori_Produk,Berat,beza_berat,UPAH,Upah30,riyal,kos_Belian_Gram,kos_Belian_Item,SpreadValue,harga_lepas_spread,adjustment,kos_item_tanpa_tax,cara_Belian,dimension_Panjang,dimension_Lebar,dimension_Saiz,code1,code2,harga_Per_Gram_Item,dulang,no_cert,gst_barang_atau_upah,StatusItem,Upah_Jualan,Upah_Member,Upah_RAF,Upah_Pengedar,code_Supplier,HargaJualan_Member,HargaJualan_Pengedar,upah_normal_dealer,upah_master_dealer,HargaJualan_RAF,hargajualan_normal_dealer,hargajualan_master_dealer,remarks,gst_ari_nashi,kadar_gst,jumlah_gst,harga_item,receiving_Status,harga_tanpa_gst,gst_included,jenis_trade_in,flag_upah,upah_per_gram,flag_image,no_id_gst from Data_Database WHERE norujukansistem='" & Frm83.L9_Text & "'"

        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Masukkan maklumat data barang ke dalam table #data_database ### - End

Call Frm83_Senarai_Belian_Header
Call Frm83_Senarai_Belian

'### Maklumat Belian / Buyback (Akaun) ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 16_gold_bar_belian where no_rujukan='" & Frm83.L9_Text & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!tarikh) Then Frm83.DTPicker1 = rs!tarikh 'Tarikh Belian
    If Not IsNull(rs!tunai) Then Frm83.L26_Text = rs!tunai 'Cara Bayaran : Tunai (Jumlah Keseluruhan Dengan Cukai GST)
    If Not IsNull(rs!jumlah_asal) Then Frm83.L11_Text = rs!jumlah_asal 'Jumlah Bayaran Asal (Jumlah Tanpa Cukai GST)
    If Not IsNull(rs!gst_value) Then Frm83.L8_Text = rs!gst_value ''Jumlah Cukai GST (%)'Jumlah Cukai GST (%)
    If Not IsNull(rs!gst_zr_harga) Then Frm83.L22_Text = rs!gst_zr_harga 'Jumlah Bayaran Yang Dikenakan Cukai GST ZR (RM)
    If Not IsNull(rs!gst_zr_cukai) Then Frm83.L23_Text = rs!gst_zr_cukai 'Jumlah Bayaran Cukai GST ZR (RM)
    If Not IsNull(rs!gst_sr_harga) Then Frm83.L24_Text = rs!gst_sr_harga 'Jumlah Bayaran Yang Dikenakan Cukai GST SR (RM)
    If Not IsNull(rs!gst_sr_cukai) Then Frm83.L25_Text = rs!gst_sr_cukai 'Jumlah Bayaran Cukai GST SR (RM)
    If Not IsNull(rs!no_id_gst_supplier) Then Frm83.TB28 = rs!no_id_gst_supplier 'No. ID GST Supplier
    If Not IsNull(rs!no_resit_supplier) Then Frm83.TB15 = rs!no_resit_supplier 'No. Resit Dari Supplier (Jika Ada)
    If Not IsNull(rs!jumlah_dengan_gst) Then Frm83.L26_Text = rs!jumlah_dengan_gst 'Jumlah Bayaran Keseluruhan (Jumlah Dengan Cukai GST)
    If Not IsNull(rs!flag_trade_in) Then
        If rs!flag_trade_in = 0 Then
            Frm83.CB7 = 1 'Flag Trade In // 0 : Tiada , 1 : Ada
        ElseIf rs!flag_trade_in = 1 Then
            Frm_LM_DATA_PENJUAL_BUYBACK = 1
            Frm83.CB8 = 1 'Flag Trade In // 0 : Tiada , 1 : Ada
            'Frm83.Pic3.Visible = True
            If Not IsNull(rs!kategori_penjual) Then Frm83.L39_Text = rs!kategori_penjual
            If Not IsNull(rs!no_rujukan_pelanggan_buyback) Then Frm85_LM_No_PENJUAL = rs!no_rujukan_pelanggan_buyback 'No. Rujukan Penjual (Penjual Buyback)
            If Not IsNull(rs!no_resit_trade_in) Then Frm83.L12_Text = rs!no_resit_trade_in 'No. Resit Trade In
        End If
    End If
    If Not IsNull(rs!no_pekerja) Then
        Frm83_LM_No_PEKERJA = rs!no_pekerja 'No. Pekerja
    End If
End If

rs.Close
Set rs = Nothing
'### Maklumat Belian / Buyback (Akaun) ### - End

'### Carian Maklumat Penjual Bagi Buyback ### - Start
If Frm85_LM_No_PENJUAL = vbNullString Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 44_senarai_pelanggan where no_resit='" & Frm83.L12_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        If Not IsNull(rs!Nama) Then
            Frm26.TB1 = rs!Nama 'Nama
            Frm83.L36_Text = rs!Nama 'Nama
        End If
        If Not IsNull(rs!no_tel) Then Frm26.TB2 = rs!no_tel 'No. Telefon

    End If
    
    rs.Close
    Set rs = Nothing
    
ElseIf Frm85_LM_No_PENJUAL <> vbNullString Then '0 : Penjual Tidak Berdaftar , 1 : Penjual Berdaftar , 2 : Ahli

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm85_LM_No_PENJUAL & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        If Not IsNull(rs!Nama) Then
            Frm28.L1_Text = rs!Nama 'Nama
            Frm83.L37_Text = rs!Nama 'Nama
        End If
        If Not IsNull(rs!no_ic) Then Frm28.L2_Text = rs!no_ic 'No. Kad Pengenalan
        If Not IsNull(rs!no_tel) Then Frm28.L3_Text = rs!no_tel 'No. Telefon
        If Not IsNull(rs!Email) Then Frm28.L4_Text = rs!Email 'E-mail
        If Not IsNull(rs!no_pelanggan) Then Frm28.L5_Text = rs!no_pelanggan 'No. Pelanggan

    End If
    
    rs.Close
    Set rs = Nothing
    
End If
'### Carian Maklumat Penjual Bagi Buyback ### - End

Frm83.CBB1.Enabled = False
Frm83.CBB1.BackColor = &H8000000A

'### Carian Maklumat Penjual (Data Pekerja) ### - Start
DATA_PEKERJA_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where NoPekerja='" & Frm83_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm83_LM_MAKLUMAT_PEKERJA = rs!Samaran & "  |  " & rs!NoPekerja
    DATA_PEKERJA_FOUND = 1
End If

rs.Close
Set rs = Nothing

If DATA_PEKERJA_FOUND = 1 Then
    On Error GoTo Err_A:
    Frm83.CBB6 = Frm83_LM_MAKLUMAT_PEKERJA
Restore_A:
End If
'### Carian Maklumat Penjual (Data Pekerja) ### - End

'Call Frm83_Senarai_Belian_Header
'Call Frm83_Senarai_Belian

Frm83.Frame9.Visible = True

Frm83.L21_Text = 1
Frm83.CMD1.Visible = False
Frm83.CMD12.Visible = True
Frm83.CMD13.Visible = False
Frm83.CMD14.Visible = False

Frm83.CMD2.Visible = True
Frm83.CMD2.Caption = "Kembali ke menu jualan"
'Frm83.CMD5.Visible = False
'Frm83.CMD10.Visible = True
'Frm83.CMD11.Visible = True

Exit Sub
Err_A:
Frm83.CBB6.AddItem Frm83_LM_MAKLUMAT_PEKERJA
Frm83.CBB6 = Frm83_LM_MAKLUMAT_PEKERJA
Resume Restore_A:
End Sub
Sub Frm84_save_edit_data_TI()
'On Error Resume Next
Dim Err(5)
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset

Call tesuto4

Exit Sub

x = 0
Y = 0 '0 : Tiada Perubahan Pada Data , 1 : Ada Perubahan Pada Data
DATA_SAVE = 0

'###Padam Data Penjual Yang Tidak Berdaftar Ke Dalam Database### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 44_senarai_pelanggan where no_resit='" & Frm84.L57_Text & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            G_ID = rs!ID
            Call recovery_44_senarai_pelanggan
                
            rs.Delete
            rs.Update
            
        End If
        
        rs.Close
        Set rs = Nothing
'###Padam Data Penjual Yang Tidak Berdaftar Ke Dalam Database### - End
    
        Frm83_LM_No_RUJUKAN_BELIAN = Frm83.L9_Text 'No. Rujukan Belian

'###Masukkan Data Belian Ke Dalam Database Akaun Belian### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 16_gold_bar_belian where no_rujukan='" & Format(Frm83.L9_Text, "000000") & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then

            If Frm84.DTPicker1 <> vbNullString Then
                rs!tarikh = Frm84.DTPicker1 'Tarikh Belian
            Else
                rs!tarikh = Null 'Tarikh Belian
            End If
            'If Frm83.DTPicker1 <> vbNullString Then %%%%Default : Bagi sistem ini hanya belian secara CASH dibenarkan
                rs!cara_bayaran = 0 'Cara Belian // 0 : Cash @ Bank in @ Kad Kredit @ Kad Debit , 1 : Cheque
            'Else
            '    rs!cara_bayaran = Null 'Cara Belian // 0 : Cash @ Bank in @ Kad Kredit @ Kad Debit , 1 : Cheque
            'End If
            If Frm83.L26_Text <> vbNullString Then
                rs!tunai = Format(Frm83.L26_Text, "0.00") 'Cara Bayaran : Tunai (Jumlah Keseluruhan Dengan Cukai GST)
            Else
                rs!tunai = Null 'Cara Bayaran : Tunai (Jumlah Keseluruhan Dengan Cukai GST)
            End If
            If Frm83.L11_Text <> vbNullString Then
                rs!jumlah_asal = Format(Frm83.L11_Text, "0.00") 'Jumlah Bayaran Asal (Jumlah Tanpa Cukai GST)
            Else
                rs!jumlah_asal = Null 'Jumlah Bayaran Asal (Jumlah Tanpa Cukai GST)
            End If

            If Frm83.L8_Text <> vbNullString Then
                rs!gst_value = Frm83.L8_Text 'Jumlah Cukai GST (%)
            Else
                rs!gst_value = "0" 'Jumlah Cukai GST (%)
            End If
            If Frm83.L22_Text <> vbNullString Then
                rs!gst_zr_harga = Format(Frm83.L22_Text, "0.00") 'Jumlah Bayaran Yang Dikenakan Cukai GST ZR (RM)
            Else
                rs!gst_zr_harga = "0.00" 'Jumlah Bayaran Yang Dikenakan Cukai GST ZR (RM)
            End If
            If Frm83.L23_Text <> vbNullString Then
                rs!gst_zr_cukai = Format(Frm83.L23_Text, "0.00") 'Jumlah Bayaran Cukai GST ZR (RM)
            Else
                rs!gst_zr_cukai = "0.00" 'Jumlah Bayaran Cukai GST ZR (RM)
            End If
            If Frm83.L24_Text <> vbNullString Then
                rs!gst_sr_harga = Format(Frm83.L24_Text, "0.00") 'Jumlah Bayaran Yang Dikenakan Cukai GST SR (RM)
            Else
                rs!gst_sr_harga = "0.00" 'Jumlah Bayaran Yang Dikenakan Cukai GST SR (RM)
            End If
            If Frm83.L25_Text <> vbNullString Then
                rs!gst_sr_cukai = Format(Frm83.L25_Text, "0.00") 'Jumlah Bayaran Cukai GST SR (RM)
            Else
                rs!gst_sr_cukai = "0.00" 'Jumlah Bayaran Cukai GST SR (RM)
            End If
            If Frm83.TB28 <> vbNullString Then
                rs!no_id_gst_supplier = Frm83.TB28 'No. ID GST Supplier
            Else
                rs!no_id_gst_supplier = Null 'No. ID GST Supplier
            End If
            If Frm83.TB15 <> vbNullString Then
                rs!no_resit_supplier = UCase(Frm83.TB15) 'No. Resit Dari Supplier (Jika Ada)
            Else
                rs!no_resit_supplier = Null 'No. Resit Dari Supplier (Jika Ada)
            End If
            If Frm83.TB1 <> vbNullString Then
                rs!Kod_Supplier = UCase(Frm83.TB1) 'Kod Supplier
            Else
                rs!Kod_Supplier = Null 'Kod Supplier
            End If
            If Frm83.L11_Text <> vbNullString Then
                rs!jumlah_tanpa_gst = Format(Frm83.L11_Text, "0.00") 'Jumlah Bayaran Asal (Jumlah Tanpa Cukai GST)
            Else
                rs!jumlah_tanpa_gst = Null 'Jumlah Bayaran Asal (Jumlah Tanpa Cukai GST)
            End If
            If Frm83.L26_Text <> vbNullString Then
                rs!jumlah_dengan_gst = Format(Frm83.L26_Text, "0.00") 'Jumlah Bayaran Keseluruhan (Jumlah Dengan Cukai GST)
            Else
                rs!jumlah_dengan_gst = Null 'Jumlah Bayaran Keseluruhan (Jumlah Dengan Cukai GST)
            End If
            If Frm84.L56_Text <> 0 Then
                Frm84_LM_Flag_TRADE_IN = 1 '0 : Tiada Urusan Trade In , 1 : Ada Urusan Trade In
                rs!flag_trade_in = 1 '0 : Tiada Urusan Trade in , 1 : Ada Urusan Trade In
                rs!trade_in_status = 1 'Flag Samada Trade In Sudah Digunakan Atau Tidak , 0 : Tiada , 1 : Ada
                'If Frm84.L56_Text = 1 Then
                '    rs!jenis_trade_in = 1 '1 : Trade in (Voucher) , 2 : Belian dengan trade in
                'ElseIf Frm84.L56_Text = 2 Then
                '    rs!jenis_trade_in = 2 '1 : Trade in (Voucher) , 2 : Belian dengan trade in
                'End If
                
                If Frm84.L57_Text <> vbNullString Then
                    rs!no_resit_trade_in = Frm84.L57_Text 'No. Resit Trade In
                Else
                    rs!no_resit_trade_in = Null 'No. Resit Trade In
                End If
                'If Frm84.L58_Text <> vbNullString Then
                '    rs!jumlah_trade_in = Format(Frm84.L58_Text, "0.00") 'No. Resit Trade In
                'Else
                '    rs!jumlah_trade_in = Null 'No. Resit Trade In
                'End If
            Else
                rs!flag_trade_in = 0 '0 : Tiada Urusan Trade in , 1 : Ada Urusan Trade In
                rs!no_resit_trade_in = Null 'No. Resit Trade In
                rs!jumlah_trade_in = Null 'Jumlah Resit Trade In (RM)
                rs!jenis_trade_in = Null '1 : Trade in (Voucher) , 2 : Belian dengan trade in
            End If
            
            If Frm83.CB8 = 1 Then
                If Frm83.L39_Text <> vbNullString Then 'Pelanggan Biasa
                    rs!kategori_penjual = Frm83.L39_Text
                Else
                    rs!kategori_penjual = Null
                End If
            Else
                rs!kategori_penjual = Null
            End If

            If Frm28.L5_Text <> vbNullString Then
                rs!no_rujukan_pelanggan_buyback = Frm28.L5_Text 'No. Rujukan Pembeli
            Else
                rs!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
            End If

            If Frm84.CBB1 <> vbNullString Then
                Frm83_LM_EMP_NO = Split(Frm84.CBB1, "  |  ")(1)
                rs!no_pekerja = Frm83_LM_EMP_NO 'No. Pekerja
            End If
            
            DATA_SAVE = 1
            rs.Update
        End If
        
        rs.Close
        Set rs = Nothing
        
'###Masukkan Data Penjual Yang Tidak Berdaftar Ke Dalam Database### - Start
        If Frm28.L5_Text = vbNullString And Frm26.TB1 <> vbNullString Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 44_senarai_pelanggan", cn, adOpenKeyset, adLockOptimistic

            rs.AddNew
            rs!tarikh = Frm84.DTPicker1 'Tarikh
            rs!no_resit = Frm84.L57_Text 'No. Resit Trade In
            If Frm26.TB1 <> vbNullString Then 'Nama
                rs!Nama = UCase(Frm26.TB1)
            Else
                rs!Nama = Null
            End If
            If Frm26.TB2 <> vbNullString Then 'No. Telefon
                rs!no_tel = UCase(Frm26.TB2)
            Else
                rs!no_tel = Null
            End If
            rs!cawangan = G_CAWANGAN
            rs!write_timestamp = Now
            rs.Update
            
            rs.Close
            Set rs = Nothing
            
        End If
'###Masukkan Data Penjual Yang Tidak Berdaftar Ke Dalam Database### - End

'###Masukkan Data Belian Ke Dalam Database### - Start
'Masukkan Data Ke Dalam Database
'0 : Data Yang Baru Dipadamkan (Dikeluarkan Dari Senarai)
'1 : Tiada Perubahan Pada Data
'2 : Item Sudah Terjual (Tidak Dibenarkan Untuk Diedit/Dipadamkan
'3 : Kemasukkan Data Baru
'4 : Data Telah Diedit
'--- Yang Terlibat Dalam Urusan Ini Adalah HANYA 0 , 3 Dan 4



'### Masukkan maklumat data barang ke dalam table #data_database ### - Start
'Barang / item baru

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "insert into Data_Database(NoRujukanSistem,tarikh_belian,bill_No_Belian,terminal,supplier_ID,nama_Supplier,Kod_Supplier,purity_ID,purity,kod_Purity,kategori_produk_ID,kategori_Produk,Kod_Kategori_Produk,Berat,beza_berat,UPAH,Upah30,riyal,kos_Belian_Gram,kos_Belian_Item,SpreadValue,harga_lepas_spread,adjustment,kos_item_tanpa_tax,cara_Belian,dimension_Panjang,dimension_Lebar,dimension_Saiz,code1,code2,harga_Per_Gram_Item,dulang,no_cert,gst_barang_atau_upah,StatusItem,Upah_Jualan,Upah_Member,Upah_RAF,Upah_Pengedar,code_Supplier,HargaJualan_Member,HargaJualan_Pengedar,upah_normal_dealer,upah_master_dealer,HargaJualan_RAF,hargajualan_normal_dealer,hargajualan_master_dealer,remarks,gst_ari_nashi,kadar_gst,jumlah_gst,harga_item,receiving_Status,harga_tanpa_gst,gst_included,jenis_trade_in,flag_upah,upah_per_gram,flag_image,write_timestamp,no_id_gst)" & _
                    "select '" & Format(Frm83_LM_No_RUJUKAN_BELIAN, "000000") & "',tarikh_belian,bill_No_Belian,terminal,supplier_ID,nama_Supplier,Kod_Supplier,purity_ID,purity,kod_Purity,kategori_produk_ID,kategori_Produk,Kod_Kategori_Produk,Berat,beza_berat,UPAH,Upah30,riyal,kos_Belian_Gram,kos_Belian_Item,Spread,harga_lepas_spread,adjustment,kos_item_tanpa_tax,cara_Belian,dimension_Panjang,dimension_Lebar,dimension_Saiz,code1,code2,harga_Per_Gram_Item,dulang,no_cert,gst_barang_atau_upah,10,Upah_Jualan,Upah_Member,Upah_RAF,Upah_Pengedar,code_Supplier,HargaJualan_Member,HargaJualan_Pengedar,upah_normal_dealer,upah_master_dealer,HargaJualan_RAF,hargajualan_normal_dealer,hargajualan_master_dealer,remarks,gst_ari_nashi,kadar_gst,jumlah_gst,harga_item,jenis,harga_tanpa_gst,gst_included,jenis_trade_in,flag_upah,upah_per_gram,flag_image,Now(),no_id_gst from " & G_BELIAN_TEMP & " WHERE StatusItem='" & 3 & "'"

        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Masukkan maklumat data barang ke dalam table #data_database ### - End

'### Update data barang ke dalam table #data_database ### - Start
'Barang sedia ada
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE Data_Database," & G_BELIAN_TEMP & " SET Data_Database.NoRujukanSistem='" & Format(Frm83_LM_No_RUJUKAN_BELIAN, "000000") & "'," _
        & "Data_Database.tarikh_belian = " & G_BELIAN_TEMP & ".tarikh_belian ," _
        & "Data_Database.bill_no_belian = " & G_BELIAN_TEMP & ".bill_no_belian , Data_Database.hargajualan_pengedar = " & G_BELIAN_TEMP & ".hargajualan_pengedar , Data_Database.upah_normal_dealer = " & G_BELIAN_TEMP & ".upah_normal_dealer , Data_Database.upah_master_dealer = " & G_BELIAN_TEMP & ".upah_master_dealer , Data_Database.hargajualan_raf = " & G_BELIAN_TEMP & ".hargajualan_raf ," _
        & "Data_Database.supplier_ID = " & G_BELIAN_TEMP & ".supplier_ID , Data_Database.hargajualan_normal_dealer = " & G_BELIAN_TEMP & ".hargajualan_normal_dealer , Data_Database.hargajualan_master_dealer = " & G_BELIAN_TEMP & ".hargajualan_master_dealer , Data_Database.remarks = " & G_BELIAN_TEMP & ".remarks ," _
        & "Data_Database.nama_Supplier = " & G_BELIAN_TEMP & ".nama_Supplier , Data_Database.gst_ari_nashi = " & G_BELIAN_TEMP & ".gst_ari_nashi , Data_Database.kadar_gst = " & G_BELIAN_TEMP & ".kadar_gst , Data_Database.jumlah_gst = " & G_BELIAN_TEMP & ".jumlah_gst , Data_Database.harga_item = " & G_BELIAN_TEMP & ".harga_item , Data_Database.receiving_status = " & G_BELIAN_TEMP & ".jenis , Data_Database.harga_tanpa_gst = " & G_BELIAN_TEMP & ".harga_tanpa_gst ," _
        & "Data_Database.Kod_Supplier = " & G_BELIAN_TEMP & ".Kod_Supplier , Data_Database.gst_included = " & G_BELIAN_TEMP & ".gst_included , Data_Database.jenis_trade_in = " & G_BELIAN_TEMP & ".jenis_trade_in , Data_Database.flag_upah = " & G_BELIAN_TEMP & ".flag_upah , Data_Database.upah_per_gram = " & G_BELIAN_TEMP & ".upah_per_gram , Data_Database.flag_image = " & G_BELIAN_TEMP & ".flag_image ," _
        & "Data_Database.purity_ID = " & G_BELIAN_TEMP & ".purity_ID ," _
        & "Data_Database.purity = " & G_BELIAN_TEMP & ".purity , Data_Database.code1 = " & G_BELIAN_TEMP & ".code1 , Data_Database.code2 = " & G_BELIAN_TEMP & ".code2 ," _
        & "Data_Database.kod_Purity = " & G_BELIAN_TEMP & ".kod_Purity ," _
        & "Data_Database.kategori_produk_ID = " & G_BELIAN_TEMP & ".kategori_produk_ID ," _
        & "Data_Database.kategori_Produk = " & G_BELIAN_TEMP & ".kategori_Produk ," _
        & "Data_Database.Kod_Kategori_Produk = " & G_BELIAN_TEMP & ".Kod_Kategori_Produk , Data_Database.terminal = " & G_BELIAN_TEMP & ".terminal ," _
        & "Data_Database.Berat = " & G_BELIAN_TEMP & ".Berat ," _
        & "Data_Database.beza_berat = " & G_BELIAN_TEMP & ".beza_berat ," _
        & "Data_Database.upah = " & G_BELIAN_TEMP & ".upah ," _
        & "Data_Database.upah30 = " & G_BELIAN_TEMP & ".upah30 ," _
        & "Data_Database.riyal = " & G_BELIAN_TEMP & ".riyal , Data_Database.no_id_gst = " & G_BELIAN_TEMP & ".no_id_gst ," _
        & "Data_Database.kos_belian_gram = " & G_BELIAN_TEMP & ".kos_belian_gram , Data_Database.kos_belian_item = " & G_BELIAN_TEMP & ".kos_belian_item , Data_Database.spreadvalue = " & G_BELIAN_TEMP & ".spread , Data_Database.harga_lepas_spread = " & G_BELIAN_TEMP & ".harga_lepas_spread , Data_Database.adjustment = " & G_BELIAN_TEMP & ".adjustment , Data_Database.kos_item_tanpa_tax = " & G_BELIAN_TEMP & ".kos_item_tanpa_tax , Data_Database.cara_belian = " & G_BELIAN_TEMP & ".cara_belian , Data_Database.dimension_panjang = " & G_BELIAN_TEMP & ".dimension_panjang , Data_Database.dimension_lebar = " & G_BELIAN_TEMP & ".dimension_lebar , Data_Database.dimension_saiz = " & G_BELIAN_TEMP & ".dimension_saiz ," _
        & "Data_Database.harga_per_gram_item = " & G_BELIAN_TEMP & ".harga_per_gram_item , Data_Database.dulang = " & G_BELIAN_TEMP & ".dulang , Data_Database.no_cert = " & G_BELIAN_TEMP & ".no_cert , Data_Database.gst_barang_atau_upah = " & G_BELIAN_TEMP & ".gst_barang_atau_upah , Data_Database.statusitem = 10 , Data_Database.upah_jualan = " & G_BELIAN_TEMP & ".upah_jualan , Data_Database.upah_member = " & G_BELIAN_TEMP & ".upah_member , Data_Database.upah_raf = " & G_BELIAN_TEMP & ".upah_raf , Data_Database.upah_pengedar = " & G_BELIAN_TEMP & ".upah_pengedar , Data_Database.code_supplier = " & G_BELIAN_TEMP & ".code_supplier , Data_Database.hargajualan_member = " & G_BELIAN_TEMP & ".hargajualan_member , " _
        & "Data_Database.write_timestamp2 = Now() WHERE " & G_BELIAN_TEMP & ".statusitem = 4 AND Data_Database.id = " & G_BELIAN_TEMP & ".id_database"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update data barang ke dalam table #data_database ### - End

'### Update data barang ke dalam table #data_database ### - Start
'Barang yang dipadamkan
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "UPDATE Data_Database," & G_BELIAN_TEMP & " SET Data_Database.statusitem='" & 0 & "', Data_Database.terminal = " & G_BELIAN_TEMP & ".terminal ," _
                & "Data_Database.write_timestamp2 = Now() WHERE " & G_BELIAN_TEMP & ".statusitem = 5 AND Data_Database.id = " & G_BELIAN_TEMP & ".id_database"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Update data barang ke dalam table #data_database ### - End

'### Update maklumat di bawah ke dalam maklumat barang ### - Start
'@no_siri_produk
'@Barcode
'@bill_no_trade_in
'@no_rujukan_pelanggan_buyback

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from data_database where NoRujukanSistem='" & Format(Frm83_LM_No_RUJUKAN_BELIAN, "000000") & "'", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
            
            If Not IsNull(rs!ID) And Not IsNull(rs!Kod_Kategori_Produk) Then
                rs!no_siri_Produk = rs!Kod_Kategori_Produk & Format(rs!ID, "000000")
            Else
                rs!no_siri_Produk = Format(rs!ID, "000000")
            End If
            If Not IsNull(rs!ID) Then
                rs!Barcode = Format(rs!ID, "000000")
            Else
                rs!Barcode = Format(rs!ID, "000000")
            End If
            If Frm83.CB8 = 1 Then
                rs!bill_No_Trade_In = Frm83.L12_Text '"TI" & Format(Frm83_LM_NO_TI, "000000") 'No. Resit Trade In

                If Frm83.L37_Text <> vbNullString Then
                    If Frm28.L5_Text <> vbNullString Then
                        rs!no_rujukan_pelanggan_buyback = Frm28.L5_Text 'No. Rujukan Pembeli
                    Else
                        rs!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
                    End If
                Else
                    rs!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
                End If
            Else
                rs!no_rujukan_pelanggan_buyback = Null 'No. Rujukan Pembeli
            End If
            rs.Update
        
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
'### Update maklumat di bawah ke dalam maklumat barang ### - End

DATA_SAVE = 1
        
        If DATA_SAVE = 1 Then
            If Frm83.TB15 <> vbNullString Then
                Frm83_LM_No_INVOICE_SUPPLIER = UCase(Frm83.TB15)
            Else
                Frm83_LM_No_INVOICE_SUPPLIER = Null
            End If
            
            '#### Update Maklumat Dulang Dalam Table Data_Database #### - Start
            'Set rs = New ADODB.Recordset
            'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

            'strsql = "UPDATE Data_Database set bill_No_Belian='" & UCase(Frm83.TB15) & "'," _
            '& "tarikh_belian='" & Frm83.DTPicker1 & "'" _
            '& "WHERE NoRujukanSistem='" & Frm83.L9_Text & "'"
            
            'Set rs = cn.Execute(strsql)
            'Set rs = Nothing
            '#### Update Maklumat Dulang Dalam Table Data_Database #### - End
        
            If Frm84.CBB1 <> vbNullString Then
                Frm83_LM_EMP_NAME = Split(Frm84.CBB1, "  |  ")(0)
            End If
        
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & Frm83_LM_EMP_NAME & "] Edit Data Trade In [" & Frm84.L57_Text & "]."
            LogDate_Memory = DateTime.Date & " " & DateTime.Time$
            Call UpdateLog_Database
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If rs!Default1 = "Default" Then
                    'If IsNumeric(Frm83.L3_Text) Then rs!NoRujukanStock = Frm83.L3_Text 'No. Siri Barcode
                    
                    If Frm83.CB9 = 1 Then
                    '    rs!NoRujukanStock = Frm83.L3_Text 'No. Siri Barcode
                    ElseIf Frm83.CB10 = 1 Then
                    '    rs!no_siri_gb = Frm83.L3_Text 'No. Siri Barcode
                    End If
                    
                    rs.Update
                End If
            End If
            
            rs.Close
            Set rs = Nothing

        End If

End Sub
Sub Frm84_modal_dan_jual()
'on error resume next
Dim Frm84_LM_HARGA As Double
Dim Frm84_LM_GST As Double
Dim Frm84_LM_BERAT As Double

Frm84_LM_HARGA = 0
Frm84_LM_BERAT = 0
Frm84_LM_GST = 0

If Frm84.L68_Text = "Modal (RM/g) :                   Jual (RM/g) :" Then
    
    If ((Frm84.TB4 <> vbNullString And IsNumeric(Frm84.TB4)) And (Frm84.TB10 <> vbNullString And IsNumeric(Frm84.TB10)) And (Frm84.TB11 <> vbNullString And IsNumeric(Frm84.TB11))) Then
    
        Frm84_LM_HARGA = Frm84.TB10
        Frm84_LM_GST = Frm84.TB11
        Frm84_LM_BERAT = Frm84.TB4
        
        If Frm84_LM_BERAT <> 0 Then
            
            If Frm84.CB3 = 1 And Frm84.CB18 = 0 Then
        
                Frm84.L67_Text = Format((Frm84_LM_HARGA + Frm84_LM_GST - Frm84_LM_GST) / Frm84_LM_BERAT, "#,##0.00")
                
            Else
            
                Frm84.L67_Text = Format((Frm84_LM_HARGA - Frm84_LM_GST) / Frm84_LM_BERAT, "#,##0.00")
            
            End If
        
        Else
        
            Frm84.L67_Text = "0.00"
            
        End If
            
    
    Else
        
        Frm84.L64_Text = "0.00"
        
    End If
    
End If

If Frm84.L68_Text = "Modal (RM)   :                      Jual (RM) :" Then

    If (Frm84.TB10 <> vbNullString And IsNumeric(Frm84.TB10)) And (Frm84.TB11 <> vbNullString And IsNumeric(Frm84.TB11)) Then

        Frm84_LM_HARGA = Frm84.TB10
        Frm84_LM_GST = Frm84.TB11
            
        If Frm84.CB3 = 1 And Frm84.CB18 = 0 Then
        

            Frm84.L67_Text = Format((Frm84_LM_HARGA + Frm84_LM_GST - Frm84_LM_GST), "#,##0.00")
            
            
        Else
        
            Frm84.L67_Text = Format((Frm84_LM_HARGA - Frm84_LM_GST), "#,##0.00")
            
        End If
    
    Else
        
        Frm84.L67_Text = "0.00"
        
    End If
    
End If
  
End Sub
Sub Frm84_kira_upah()
'on error resume next
Dim Frm84_LM_BERAT As Double
Dim Frm84_LM_UPAH_PER_GRAM As Double

Frm84_LM_BERAT = 0
Frm84_LM_UPAH_PER_GRAM = 0

If (Frm84.TB4 <> vbNullString And IsNumeric(Frm84.TB4)) And (Frm84.TB22 <> vbNullString And IsNumeric(Frm84.TB22)) Then

    Frm84_LM_BERAT = Frm84.TB4
    Frm84_LM_UPAH_PER_GRAM = Frm84.TB22
    
    Frm84.TB15 = Format(Frm84_LM_BERAT * Frm84_LM_UPAH_PER_GRAM, "0.00")
End If
End Sub
Sub Frm84_jurujual()
'On Error Resume Next
If MDI_frm1.L3_Text <> vbNullString Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        Frm84.CBB1 = rs!Samaran & "  |  " & rs!NoPekerja
        
    Else
        
        Frm84.CBB1.AddItem "" & "  |  " & rs!Samaran
        Frm84.CBB1 = "" & "  |  " & rs!Samaran
        
    End If
    
    rs.Close
    Set rs = Nothing
    
    If G_LOCK_JURUJUAL = "YES" Then
    
        Frm84.CBB1.Enabled = False
        Frm84.CBB1.BackColor = &H8000000A

    Else
    
        Frm84.CBB1.Enabled = True
        Frm84.CBB1.BackColor = &HFFFFFF

    End If

End If
End Sub
Sub Frm84_kiraan_potongan_kupon()
'On Error Resume Next
Dim Frm84_LM_BERAT As Double
Dim Frm84_LM_KUPON As Double

LM_RATE_KUPON_2 = vbNullString

Frm84_LM_BERAT = 0
Frm84_LM_KUPON = 0

If Frm84.L80_Text <> vbNullString Then
    If InStr(1, Frm84.L80_Text, " ") <> 0 Then
        LM_RATE_KUPON_1 = Split(Frm84.L80_Text, " ")(1)
        LM_RATE_KUPON_2 = Split(LM_RATE_KUPON_1, " ")(0)
    End If
End If

If LM_RATE_KUPON_2 <> vbNullString Then

    If IsNumeric(LM_RATE_KUPON_2) Then Frm84_LM_KUPON = LM_RATE_KUPON_2
    
End If

If Frm84.L15_Text <> vbNullString And IsNumeric(Frm84.L15_Text) Then

    Frm84_LM_BERAT = Frm84.L15_Text
    
End If

If Frm84.CB14 = 1 Then
    Frm84.TB34 = Format(Frm84_LM_BERAT * Frm84_LM_KUPON, "0.00")
Else
    Frm84.TB34 = Format(0, "0.00")
End If
End Sub
Sub Frm84_kira_harga_layak_mata()
'on error resume next
Dim Frm84_LM_HARGA_BARANG As Double
Dim Frm84_LM_ADJUSTMENT As Double
Dim Frm84_LM_KUPON As Double
Dim Frm84_LM_NILAI_TEBUS As Double
Dim Frm84_LM_LAYAK_MATA As Double

If GLOBAL_DISABLE = 0 Then

    Frm84_LM_HARGA_BARANG = 0
    Frm84_LM_ADJUSTMENT = 0
    Frm84_LM_NILAI_TEBUS = 0
    Frm84_LM_KUPON = 0
    
    If ((Frm84.L17_Text <> vbNullString And IsNumeric(Frm84.L17_Text)) And (Frm84.TB20 <> vbNullString And IsNumeric(Frm84.TB20)) And (Frm84.TB34 <> vbNullString And IsNumeric(Frm84.TB34)) And (Frm84.L78_Text <> vbNullString And IsNumeric(Frm84.L78_Text))) Then
        Frm84_LM_HARGA_BARANG = Frm84.L17_Text 'Harga Barang
        Frm84_LM_ADJUSTMENT = Frm84.TB20 'Adjustment
        Frm84_LM_KUPON = Frm84.TB34 'Kupon
        Frm84_LM_NILAI_TEBUS = Frm84.L78_Text 'Nilai mata yang ditebus
        
        Frm84_LM_LAYAK_MATA = Format(Frm84_LM_HARGA_BARANG - Frm84_LM_ADJUSTMENT - Frm84_LM_NILAI_TEBUS - Frm84_LM_KUPON, "#,##0.00") 'Harga barang yang layak dapat mata
        
        If Frm84_LM_LAYAK_MATA >= 0 Then
        
            Frm84.L75_Text = Format(Frm84_LM_LAYAK_MATA, "#,##0.00") 'Harga barang yang layak dapat mata
            
        Else
        
            Frm84.L75_Text = "0.00" 'Harga barang yang layak dapat mata
            
        End If
    Else
        Frm84.L75_Text = "0.00" 'Harga barang yang layak dapat mata
    End If
    
    Frm84.L73_Text = Frm84.L78_Text 'Harga barang yang layak dapat mata
    
End If
End Sub
Sub Frm84_kira_mata_ganjaran()
'on error resume next
Dim Frm84_LM_HARGA_LAYAK As Double
Dim Frm84_LM_KADAR As Double

If GLOBAL_DISABLE = 0 Then
    
    Frm84_LM_HARGA_LAYAK = 0
    Frm84_LM_KADAR = 0
    
    If ((Frm84.L75_Text <> vbNullString And IsNumeric(Frm84.L75_Text)) And (Frm84.TB35 <> vbNullString And IsNumeric(Frm84.TB35))) Then
        
        Frm84_LM_HARGA_LAYAK = Frm84.L75_Text
        Frm84_LM_KADAR = Frm84.TB35
        
        Frm84.L76_Text = Int(Frm84_LM_HARGA_LAYAK * Frm84_LM_KADAR)
        
    Else
    
        Frm84.L76_Text = "0"
    
    End If
    
End If
End Sub
Sub Frm84_nilai_mata_tebus()
'on error resume next
Dim Frm84_LM_MATA_TEBUS As Double
Dim Frm84_LM_KADAR As Double

If GLOBAL_DISABLE = 0 Then
    
    Frm84_LM_MATA_TEBUS = 0
    Frm84_LM_KADAR = 0
    
    If ((Frm84.TB36 <> vbNullString And IsNumeric(Frm84.TB36)) And (Frm84.TB37 <> vbNullString And IsNumeric(Frm84.TB37))) Then
        
        Frm84_LM_MATA_TEBUS = Frm84.TB36
        Frm84_LM_KADAR = Frm84.TB37
        
        Frm84.L78_Text = Format(Frm84_LM_MATA_TEBUS * Frm84_LM_KADAR, "0.00")
        
    Else
    
        Frm84.L78_Text = "0.00"
    
    End If
    
End If
End Sub




Sub frm84_harga_selepas_diskaun()
'on error resume next
Dim Frm84_HARGA_ASAL As Double
Dim Frm84_DISKAUN As Double

If GLOBAL_DISABLE = 0 Then

    Frm84_HARGA_ASAL = 0
    Frm84_DISKAUN = 0
    
    If ((Frm84.L19_Text <> vbNullString And IsNumeric(Frm84.L19_Text)) And (Frm84.TB19 <> vbNullString And IsNumeric(Frm84.TB19))) Then
        Frm84_HARGA_ASAL = Frm84.L19_Text 'Harga Asal
        Frm84_DISKAUN = Frm84.TB19 'Diskaun
        
        Frm84.L20_Text = Format(Frm84_HARGA_ASAL - ((Frm84_DISKAUN / 100) * Frm84_HARGA_ASAL), "#,##0.00") 'Harga Selepas Diskaun
    Else
        Frm84.L20_Text = "0.00" 'Harga Selepas Diskaun
    End If
    
End If
End Sub
Sub frm84_kiraan_harga_asal()
'on error resume next
Dim Frm84_BERAT As Double
Dim Frm84_HARGA_PER_GRAM As Double
Dim Frm84_UPAH As Double
Dim Frm84_KOMISEN_PER_GRAM As Double

Frm84_BERAT = 0
Frm84_HARGA_PER_GRAM = 0
Frm84_UPAH = 0

'If G_CALC_AUTO = 0 Then

    If (Frm84.TB4 <> vbNullString And IsNumeric(Frm84.TB4)) Then Frm84_BERAT = Frm84.TB4
    If (Frm84.TB5 <> vbNullString And IsNumeric(Frm84.TB5)) Then Frm84_HARGA_PER_GRAM = Frm84.TB5
    If (Frm84.TB15 <> vbNullString And IsNumeric(Frm84.TB15)) Then Frm84_UPAH = Frm84.TB15
    
    Frm84.TB6 = Format((Frm84_BERAT * Frm84_HARGA_PER_GRAM) + Frm84_UPAH, "#,##0.00") 'Harga Asal
    
    'Call Frm84_kira_upah
    Call Frm84_pengiraan_komisyen_dropship
    Call Frm84_modal_dan_jual
    
'End If
End Sub
Sub frm84_selepas_diskaun()
'on error resume next
Dim Frm84_HARGA_ASAL As Double
Dim Frm84_DISKAUN As Double
Dim Frm84_JUMLAH_DISKAUN As Double

Frm84_HARGA_ASAL = 0
Frm84_DISKAUN = 0
Frm84_JUMLAH_DISKAUN = 0

If (Frm84.TB6 <> vbNullString And IsNumeric(Frm84.TB6)) Then Frm84_HARGA_ASAL = Frm84.TB6
If (Frm84.TB7 <> vbNullString And IsNumeric(Frm84.TB7)) Then Frm84_DISKAUN = Frm84.TB7

Frm84_JUMLAH_DISKAUN = Frm84_HARGA_ASAL * (Frm84_DISKAUN / 100)

Frm84.TB8 = Format(Frm84_HARGA_ASAL - Frm84_JUMLAH_DISKAUN, "#,##0.00") 'Harga Asal
End Sub
Sub frm84_harga_jualan()
'on error resume next
Dim Frm84_HARGA_LEPAS_DISKAUN As Double
Dim Frm84_ADJUSTMENT As Double

Frm84_HARGA_LEPAS_DISKAUN = 0
Frm84_ADJUSTMENT = 0

If (Frm84.TB8 <> vbNullString And IsNumeric(Frm84.TB8)) Then Frm84_HARGA_LEPAS_DISKAUN = Frm84.TB8
If (Frm84.TB9 <> vbNullString And IsNumeric(Frm84.TB9)) Then Frm84_ADJUSTMENT = Frm84.TB9

Frm84.TB10 = Format(Frm84_HARGA_LEPAS_DISKAUN - Frm84_ADJUSTMENT, "#,##0.00")
End Sub
Sub frm84_kiraan_gst()
'On Error Resume Next
Dim Frm84_LM_HARGA As Double
Dim frm84_LM_KADAR_GST As Double
Dim frm84_TOTAL_GST As Double
Dim frm84_HARGA_TANPA_GST As Double

Frm84_LM_HARGA = 0
frm84_LM_KADAR_GST = 0
frm84_TOTAL_GST = 0
frm84_HARGA_TANPA_GST = 0

If Frm84.CB12 = 0 Then

    If Frm84.TB10 <> vbNullString And IsNumeric(Frm84.TB10) Then
        Frm84_LM_HARGA = Frm84.TB10
    End If
    
Else

    If Frm84.TB15 <> vbNullString And IsNumeric(Frm84.TB15) Then
        Frm84_LM_HARGA = Frm84.TB15
    End If
    
End If

If Frm84.L8_Text <> vbNullString And IsNumeric(Frm84.L8_Text) Then
    frm84_LM_KADAR_GST = Frm84.L8_Text
End If

If Frm84.CB2 = 1 Then
    
    Frm84.TB11 = Format(frm84_TOTAL_GST, "#,##0.00") 'Jumlah Cukai GST (RM)
    Frm84.L44_Text = Format(Frm84_LM_HARGA, "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)

ElseIf Frm84.CB3 = 1 Then

    Frm84.TB11 = Format((frm84_LM_KADAR_GST / 100) * Frm84_LM_HARGA, "#,##0.00") 'Jumlah Cukai GST (RM)
    Frm84.L44_Text = Format(Frm84_LM_HARGA, "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
    
ElseIf Frm84.CB18 = 1 Then

    Frm84.L44_Text = Format(Frm84_LM_HARGA / (1 + (frm84_LM_KADAR_GST / 100)), "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
    Frm84.TB11 = Format(Frm84_LM_HARGA - (Frm84_LM_HARGA / (1 + (frm84_LM_KADAR_GST / 100))), "#,##0.00") 'Jumlah Cukai GST (RM)
        
End If
End Sub
Sub frm84_senarai_barang_purity()
'on error resume next
Frm84.CBB3.Clear
Frm84.CBB4.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    If Not IsNull(rs!Metal_Purity) Then Frm84.CBB4.AddItem rs!Metal_Purity
    If Not IsNull(rs!kategori_Produk) Then Frm84.CBB3.AddItem rs!kategori_Produk

    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
End Sub
Sub frm84_call_edit_berat()
'on error resume next
Dim LM_BERAT_ASAL As Double
Dim LM_BERAT_GUNA As Double
Dim LM_BERAT_TEMP As Double
Dim LM_BERAT_TEMP_ASAL As Double

LM_BERAT_ASAL = 0
LM_BERAT_GUNA = 0
LM_BERAT_TEMP = 0
LM_BERAT_TEMP_ASAL = 0

If Frm84.TB4 <> vbNullString And IsNumeric(Frm84.TB4) Then LM_BERAT_TEMP_ASAL = Frm84.TB4

Set rs3 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs3.Open "select SUM(beza_berat) from data_database where Purity='" & Frm84.CBB4 & "' AND (((statusitem = 10 OR statusitem = 12 OR statusitem = 20 OR statusitem = 22) AND receiving_Status = 2) OR ((statusitem = 12 OR statusitem = 20 OR statusitem = 22) AND receiving_Status = 0))", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs3(0)) Then LM_BERAT_ASAL = Format(rs3(0), "#,##0.00")
    
rs3.Close
Set rs3 = Nothing

Set rs3 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs3.Open "select SUM(berat) from 85_penggunaan_ti where purity='" & Frm84.CBB4 & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
    
If Not IsNull(rs3(0)) Then LM_BERAT_GUNA = Format(rs3(0), "#,##0.00")
    
rs3.Close
Set rs3 = Nothing

Set rs3 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs3.Open "select SUM(Berat_Jualan) from " & G_JUALAN_TEMP & " where purity='" & Frm84.L13_Text & "' AND flag_barang = 1 AND (status = 1 OR Status = 2 OR Status = 3 OR Status = 4)", cn, adOpenKeyset, adLockOptimistic
    
If Not IsNull(rs3(0)) Then LM_BERAT_TEMP = Format(rs3(0), "#,##0.00")
    
rs3.Close
Set rs3 = Nothing

Frm84.TB3 = Format(LM_BERAT_ASAL - LM_BERAT_GUNA - LM_BERAT_TEMP + LM_BERAT_TEMP_ASAL, "#,##0.00")
End Sub
Sub frm84_berat_guna_dr_invoice_ini()
'on error resume next
Dim LM_BERAT_ASAL As Double
Dim LM_BERAT_GUNA As Double

LM_BERAT_ASAL = 0
LM_BERAT_GUNA = 0

If Frm84.TB3 <> vbNullString And IsNumeric(Frm84.TB3) Then LM_BERAT_ASAL = Frm84.TB3

Set rs3 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs3.Open "select SUM(berat) from 85_penggunaan_ti where purity='" & Frm84.CBB4 & "' AND status = 1 AND no_rujukan='" & Frm84.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
If Not IsNull(rs3(0)) Then LM_BERAT_GUNA = Format(rs3(0), "#,##0.00")
    
rs3.Close
Set rs3 = Nothing

Frm84.TB3 = Format(LM_BERAT_ASAL + LM_BERAT_GUNA, "#,##0.00")
End Sub
Sub a()
'On Error Resume Next
Dim Frm84_LM_HARGA As Double
Dim frm84_LM_KADAR_GST As Double
Dim frm84_TOTAL_GST As Double
Dim frm84_HARGA_TANPA_GST As Double

Frm84_LM_HARGA = 0
frm84_LM_KADAR_GST = 0
frm84_TOTAL_GST = 0
frm84_HARGA_TANPA_GST = 0

If Frm84.CB12 = 0 Then

    If Frm84.TB10 <> vbNullString And IsNumeric(Frm84.TB10) Then
        Frm84_LM_HARGA = Frm84.TB10
    End If
    
Else

    If Frm84.TB15 <> vbNullString And IsNumeric(Frm84.TB15) Then
        Frm84_LM_HARGA = Frm84.TB15
    End If
    
End If

If Frm84.L8_Text <> vbNullString And IsNumeric(Frm84.L8_Text) Then
    frm84_LM_KADAR_GST = Frm84.L8_Text
End If

If Frm84.CB2 = 1 Then
    
    Frm84.TB11 = Format(frm84_TOTAL_GST, "#,##0.00") 'Jumlah Cukai GST (RM)
    Frm84.L44_Text = Format(Frm84_LM_HARGA, "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)

ElseIf Frm84.CB3 = 1 Then

    Frm84.TB11 = Format((frm84_LM_KADAR_GST / 100) * Frm84_LM_HARGA, "#,##0.00") 'Jumlah Cukai GST (RM)
    Frm84.L44_Text = Format(Frm84_LM_HARGA, "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
    
ElseIf Frm84.CB18 = 1 Then

    Frm84.L44_Text = Format(Frm84_LM_HARGA / (1 + (frm84_LM_KADAR_GST / 100)), "#,##0.00") 'Jumlah Harga Barang Tanpa GST (RM)
    Frm84.TB11 = Format(Frm84_LM_HARGA - (Frm84_LM_HARGA / (1 + (frm84_LM_KADAR_GST / 100))), "#,##0.00") 'Jumlah Cukai GST (RM)
        
End If
End Sub
Sub frm84_disable_frame()
'on error resume next
Frm84.Frame1.Visible = False
Frm84.Frame2.Visible = False
Frm84.Frame3.Visible = False
Frm84.Frame4.Visible = False
Frm84.Frame5.Visible = False
Frm84.Frame6.Visible = False
Frm84.Pic3.Visible = False
Frm84.Pic6.Visible = False
Frm84.Frame8.Visible = False
End Sub
Sub cetak_invoice()
'on error resume next
Dim LM_JUMLAH_BAYAR As Double
Dim LM_CAJ_KAD As Double
Dim LM_CAJ_GST As Double
Dim A1 As String
Dim B1 As String
Dim C1 As String
Dim D1 As String
Dim E1 As String
Dim F1 As String
Dim G1 As String
Dim H1 As String
Dim I1 As String

LM_JUMLAH_BAYAR = 0
LM_CAJ_KAD = 0
LM_CAJ_GST = 0

Report83.Sections("Section1").Controls("L1").Caption = vbNullString 'Maklumat Pembeli : Nama
Report83.Sections("Section1").Controls("L2").Caption = vbNullString 'Maklumat Pembeli : No. Telefon
Report83.Sections("Section1").Controls("L3").Caption = vbNullString 'Maklumat Pembeli : No. Keahlian
Report83.Sections("Section1").Controls("L4").Caption = vbNullString 'Maklumat Kedai
Report83.Sections("Section1").Controls("L5").Caption = vbNullString 'No. Invoice
Report83.Sections("Section1").Controls("L6").Caption = vbNullString 'Tarikh
Report83.Sections("Section1").Controls("L7").Caption = vbNullString 'No.
Report83.Sections("Section1").Controls("L8").Caption = vbNullString 'Jenis Barang
Report83.Sections("Section1").Controls("L9").Caption = vbNullString 'Ketulenan
Report83.Sections("Section1").Controls("L10").Caption = vbNullString 'Berat
Report83.Sections("Section1").Controls("L11").Caption = vbNullString 'Harga Semasa
Report83.Sections("Section1").Controls("L12").Caption = vbNullString 'Jumlah
Report83.Sections("Section1").Controls("L13").Caption = vbNullString 'Remarks
Report83.Sections("Section1").Controls("L14").Caption = vbNullString 'Nama Jurujual
Report83.Sections("Section1").Controls("L15").Caption = vbNullString 'Maklumat Bayaran
Report83.Sections("Section1").Controls("L16").Caption = vbNullString 'Jumlah
Report83.Sections("Section1").Controls("L17").Visible = False

'L1
Report83.Sections("Section1").Controls("L1").Left = G_L1_LEFT
Report83.Sections("Section1").Controls("L1").Top = G_L1_TOP
Report83.Sections("Section1").Controls("L1").Font.Bold = G_L1_BOLD
Report83.Sections("Section1").Controls("L1").Font.Italic = G_L1_ITALIC
Report83.Sections("Section1").Controls("L1").Font.Size = G_L1_FONT
Report83.Sections("Section1").Controls("L1").Width = G_L1_WIDTH
Report83.Sections("Section1").Controls("L1").Height = G_L1_HEIGHT

'L2
Report83.Sections("Section1").Controls("L2").Left = G_L2_LEFT
Report83.Sections("Section1").Controls("L2").Top = G_L2_TOP
Report83.Sections("Section1").Controls("L2").Font.Bold = G_L2_BOLD
Report83.Sections("Section1").Controls("L2").Font.Italic = G_L2_ITALIC
Report83.Sections("Section1").Controls("L2").Font.Size = G_L2_FONT
Report83.Sections("Section1").Controls("L2").Width = G_L2_WIDTH
Report83.Sections("Section1").Controls("L2").Height = G_L2_HEIGHT

'L3
Report83.Sections("Section1").Controls("L3").Left = G_L3_LEFT
Report83.Sections("Section1").Controls("L3").Top = G_L3_TOP
Report83.Sections("Section1").Controls("L3").Font.Bold = G_L3_BOLD
Report83.Sections("Section1").Controls("L3").Font.Italic = G_L3_ITALIC
Report83.Sections("Section1").Controls("L3").Font.Size = G_L3_FONT
Report83.Sections("Section1").Controls("L3").Width = G_L3_WIDTH
Report83.Sections("Section1").Controls("L3").Height = G_L3_HEIGHT

'L4
Report83.Sections("Section1").Controls("L4").Left = G_L4_LEFT
Report83.Sections("Section1").Controls("L4").Top = G_L4_TOP
Report83.Sections("Section1").Controls("L4").Font.Bold = G_L4_BOLD
Report83.Sections("Section1").Controls("L4").Font.Italic = G_L4_ITALIC
Report83.Sections("Section1").Controls("L4").Font.Size = G_L4_FONT
Report83.Sections("Section1").Controls("L4").Width = G_L4_WIDTH
Report83.Sections("Section1").Controls("L4").Height = G_L4_HEIGHT

'L5
Report83.Sections("Section1").Controls("L5").Left = G_L5_LEFT
Report83.Sections("Section1").Controls("L5").Top = G_L5_TOP
Report83.Sections("Section1").Controls("L5").Font.Bold = G_L5_BOLD
Report83.Sections("Section1").Controls("L5").Font.Italic = G_L5_ITALIC
Report83.Sections("Section1").Controls("L5").Font.Size = G_L5_FONT
Report83.Sections("Section1").Controls("L5").Width = G_L5_WIDTH
Report83.Sections("Section1").Controls("L5").Height = G_L5_HEIGHT

'L6
Report83.Sections("Section1").Controls("L6").Left = G_L6_LEFT
Report83.Sections("Section1").Controls("L6").Top = G_L6_TOP
Report83.Sections("Section1").Controls("L6").Font.Bold = G_L6_BOLD
Report83.Sections("Section1").Controls("L6").Font.Italic = G_L6_ITALIC
Report83.Sections("Section1").Controls("L6").Font.Size = G_L6_FONT
Report83.Sections("Section1").Controls("L6").Width = G_L6_WIDTH
Report83.Sections("Section1").Controls("L6").Height = G_L6_HEIGHT

'L7
Report83.Sections("Section1").Controls("L7").Left = G_L7_LEFT
Report83.Sections("Section1").Controls("L7").Top = G_L7_TOP
Report83.Sections("Section1").Controls("L7").Font.Bold = G_L7_BOLD
Report83.Sections("Section1").Controls("L7").Font.Italic = G_L7_ITALIC
Report83.Sections("Section1").Controls("L7").Font.Size = G_L7_FONT
Report83.Sections("Section1").Controls("L7").Width = G_L7_WIDTH
Report83.Sections("Section1").Controls("L7").Height = G_L7_HEIGHT

'L8
Report83.Sections("Section1").Controls("L8").Left = G_L8_LEFT
Report83.Sections("Section1").Controls("L8").Top = G_L8_TOP
Report83.Sections("Section1").Controls("L8").Font.Bold = G_L8_BOLD
Report83.Sections("Section1").Controls("L8").Font.Italic = G_L8_ITALIC
Report83.Sections("Section1").Controls("L8").Font.Size = G_L8_FONT
Report83.Sections("Section1").Controls("L8").Width = G_L8_WIDTH
Report83.Sections("Section1").Controls("L8").Height = G_L8_HEIGHT

'L9
Report83.Sections("Section1").Controls("L9").Left = G_L9_LEFT
Report83.Sections("Section1").Controls("L9").Top = G_L9_TOP
Report83.Sections("Section1").Controls("L9").Font.Bold = G_L9_BOLD
Report83.Sections("Section1").Controls("L9").Font.Italic = G_L9_ITALIC
Report83.Sections("Section1").Controls("L9").Font.Size = G_L9_FONT
Report83.Sections("Section1").Controls("L9").Width = G_L9_WIDTH
Report83.Sections("Section1").Controls("L9").Height = G_L9_HEIGHT

'L10
Report83.Sections("Section1").Controls("L10").Left = G_L10_LEFT
Report83.Sections("Section1").Controls("L10").Top = G_L10_TOP
Report83.Sections("Section1").Controls("L10").Font.Bold = G_L10_BOLD
Report83.Sections("Section1").Controls("L10").Font.Italic = G_L10_ITALIC
Report83.Sections("Section1").Controls("L10").Font.Size = G_L10_FONT
Report83.Sections("Section1").Controls("L10").Width = G_L10_WIDTH
Report83.Sections("Section1").Controls("L10").Height = G_L10_HEIGHT

'L11
Report83.Sections("Section1").Controls("L11").Left = G_L11_LEFT
Report83.Sections("Section1").Controls("L11").Top = G_L11_TOP
Report83.Sections("Section1").Controls("L11").Font.Bold = G_L11_BOLD
Report83.Sections("Section1").Controls("L11").Font.Italic = G_L11_ITALIC
Report83.Sections("Section1").Controls("L11").Font.Size = G_L11_FONT
Report83.Sections("Section1").Controls("L11").Width = G_L11_WIDTH
Report83.Sections("Section1").Controls("L11").Height = G_L11_HEIGHT

'L12
Report83.Sections("Section1").Controls("L12").Left = G_L12_LEFT
Report83.Sections("Section1").Controls("L12").Top = G_L12_TOP
Report83.Sections("Section1").Controls("L12").Font.Bold = G_L12_BOLD
Report83.Sections("Section1").Controls("L12").Font.Italic = G_L12_ITALIC
Report83.Sections("Section1").Controls("L12").Font.Size = G_L12_FONT
Report83.Sections("Section1").Controls("L12").Width = G_L12_WIDTH
Report83.Sections("Section1").Controls("L12").Height = G_L12_HEIGHT

'L13
Report83.Sections("Section1").Controls("L13").Left = G_L13_LEFT
Report83.Sections("Section1").Controls("L13").Top = G_L13_TOP
Report83.Sections("Section1").Controls("L13").Font.Bold = G_L13_BOLD
Report83.Sections("Section1").Controls("L13").Font.Italic = G_L13_ITALIC
Report83.Sections("Section1").Controls("L13").Font.Size = G_L13_FONT
Report83.Sections("Section1").Controls("L13").Width = G_L13_WIDTH
Report83.Sections("Section1").Controls("L13").Height = G_L13_HEIGHT

'L14
Report83.Sections("Section1").Controls("L14").Left = G_L14_LEFT
Report83.Sections("Section1").Controls("L14").Top = G_L14_TOP
Report83.Sections("Section1").Controls("L14").Font.Bold = G_L14_BOLD
Report83.Sections("Section1").Controls("L14").Font.Italic = G_L14_ITALIC
Report83.Sections("Section1").Controls("L14").Font.Size = G_L14_FONT
Report83.Sections("Section1").Controls("L14").Width = G_L14_WIDTH
Report83.Sections("Section1").Controls("L14").Height = G_L14_HEIGHT

'L15
Report83.Sections("Section1").Controls("L15").Left = G_L15_LEFT
Report83.Sections("Section1").Controls("L15").Top = G_L15_TOP
Report83.Sections("Section1").Controls("L15").Font.Bold = G_L15_BOLD
Report83.Sections("Section1").Controls("L15").Font.Italic = G_L15_ITALIC
Report83.Sections("Section1").Controls("L15").Font.Size = G_L15_FONT
Report83.Sections("Section1").Controls("L15").Width = G_L15_WIDTH
Report83.Sections("Section1").Controls("L15").Height = G_L15_HEIGHT

'L16
Report83.Sections("Section1").Controls("L16").Left = G_L16_LEFT
Report83.Sections("Section1").Controls("L16").Top = G_L16_TOP
Report83.Sections("Section1").Controls("L16").Font.Bold = G_L16_BOLD
Report83.Sections("Section1").Controls("L16").Font.Italic = G_L16_ITALIC
Report83.Sections("Section1").Controls("L16").Font.Size = G_L16_FONT
Report83.Sections("Section1").Controls("L16").Width = G_L16_WIDTH
Report83.Sections("Section1").Controls("L16").Height = G_L16_HEIGHT

'L17
Report83.Sections("Section1").Controls("L17").Left = G_L17_LEFT
Report83.Sections("Section1").Controls("L17").Top = G_L17_TOP
Report83.Sections("Section1").Controls("L17").Font.Bold = G_L17_BOLD
Report83.Sections("Section1").Controls("L17").Font.Italic = G_L17_ITALIC
Report83.Sections("Section1").Controls("L17").Font.Size = G_L17_FONT
Report83.Sections("Section1").Controls("L17").Width = G_L17_WIDTH
Report83.Sections("Section1").Controls("L17").Height = G_L17_HEIGHT

'L18
Report83.Sections("Section1").Controls("L18").Left = G_L18_LEFT
Report83.Sections("Section1").Controls("L18").Top = G_L18_TOP
Report83.Sections("Section1").Controls("L18").Font.Bold = G_L18_BOLD
Report83.Sections("Section1").Controls("L18").Font.Italic = G_L18_ITALIC
Report83.Sections("Section1").Controls("L18").Font.Size = G_L18_FONT
Report83.Sections("Section1").Controls("L18").Width = G_L18_WIDTH
Report83.Sections("Section1").Controls("L18").Height = G_L18_HEIGHT

x = 0
A1 = vbNullString
B1 = vbNullString
C1 = vbNullString
D1 = vbNullString
E1 = vbNullString
F1 = vbNullString
G1 = vbNullString
H1 = vbNullString
I1 = vbNullString

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Not IsNull(rs!no_pendaftaran) Then I1 = I1 & rs!no_pendaftaran & vbCrLf
    If Not IsNull(rs!alamat) Then I1 = I1 & rs!alamat & vbCrLf
    If Not IsNull(rs!no_tel) Then I1 = I1 & rs!no_tel & vbCrLf
    
End If

rs.Close
Set rs = Nothing

Report83.Sections("Section1").Controls("L4").Caption = I1 'Maklumat Kedai

'Maklumat invoice
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 23_senarai_jualan where no_resit='" & G_No_RESIT_JUALAN & "' AND status_rekod = 1", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    A1 = A1 & x & ")" & vbCrLf
    
    If Not IsNull(rs!kategori_Produk) Then
        B1 = B1 & rs!kategori_Produk & vbCrLf
    Else
        B1 = B1 & vbNullString & vbCrLf
    End If
    If Not IsNull(rs!purity) Then
        C1 = C1 & rs!purity & vbCrLf
    Else
        C1 = C1 & vbNullString & vbCrLf
    End If
    If Not IsNull(rs!berat_jualan) Then
        D1 = D1 & Format(rs!berat_jualan, "#,##0.00 g") & vbCrLf
    Else
        D1 = D1 & vbNullString & vbCrLf
    End If
    If Not IsNull(rs!harga_Semasa) Then
        E1 = E1 & Format(rs!harga_Semasa, "#,##0.00") & vbCrLf
    Else
        E1 = E1 & vbNullString & vbCrLf
    End If
    If Not IsNull(rs!harga_dengan_gst) Then
        F1 = F1 & Format(rs!harga_dengan_gst, "#,##0.00") & vbCrLf
    Else
        F1 = F1 & vbNullString & vbCrLf
    End If
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

L7 = A1 'No.
L8 = B1 'Kategori Produk
L9 = C1 'Ketulenan
L10 = D1 'Berat
L11 = E1 'Harga Emas
L12 = F1 'Harga

Report83.Sections("Section1").Controls("L5").Caption = L5
Report83.Sections("Section1").Controls("L6").Caption = L6
Report83.Sections("Section1").Controls("L7").Caption = L7
Report83.Sections("Section1").Controls("L8").Caption = L8
Report83.Sections("Section1").Controls("L9").Caption = L9
Report83.Sections("Section1").Controls("L10").Caption = L10
Report83.Sections("Section1").Controls("L11").Caption = L11
Report83.Sections("Section1").Controls("L12").Caption = L12
Report83.Sections("Section1").Controls("L13").Caption = L13
Report83.Sections("Section1").Controls("L14").Caption = L14
Report83.Sections("Section1").Controls("L15").Caption = L15
Report83.Sections("Section1").Controls("L16").Caption = L16

LM_TI = vbNullString
LM_SUSUT_NILAI_MODE = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where no_resit='" & G_No_RESIT_JUALAN & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!jenis_trade_in) Then
        If rs!jenis_trade_in = 3 Then LM_SUSUT_NILAI_MODE = 1
    End If
    If Not IsNull(rs!JUMLAH_BERAT) Then Report83.Sections("Section1").Controls("L18").Caption = Format(rs!JUMLAH_BERAT, "#,##0.00 g") 'Berat
    If Not IsNull(rs!tarikh) Then Report83.Sections("Section1").Controls("L6").Caption = "Tarikh : " & rs!tarikh 'Tarikh
    If Not IsNull(rs!no_resit) Then Report83.Sections("Section1").Controls("L5").Caption = "No. Invoice : " & rs!no_resit 'No. Invoice
    If Not IsNull(rs!jumlah_perlu_bayar) Then
        If IsNumeric(rs!jumlah_perlu_bayar) Then LM_JUMLAH_BAYAR = rs!jumlah_perlu_bayar
    End If
    If Not IsNull(rs!jumlah_cas_kad_kredit) Then
        If IsNumeric(rs!jumlah_cas_kad_kredit) Then LM_CAJ_KAD = rs!jumlah_cas_kad_kredit
    End If
    If Not IsNull(rs!gst_kad_kredit) Then
        If IsNumeric(rs!gst_kad_kredit) Then LM_CAJ_GST = rs!gst_kad_kredit
    End If
    Report83.Sections("Section1").Controls("L16").Caption = Format(LM_JUMLAH_BAYAR + LM_CAJ_KAD + LM_CAJ_GST, "#,##0.00")  'Jumlah Bayaran Yang Perlu Dibuat (RM)
    If Not IsNull(rs!flag_trade_in) Then
        If rs!flag_trade_in = 1 Then
            LM_TI = rs!no_resit_trade_in
        End If
    End If
    If Not IsNull(rs!jumlah_trade_in) Then
        LM_HARGA_TI = rs!jumlah_trade_in 'Jumlah Resit Trade In (RM)
    Else
        LM_HARGA_TI = "0" 'Jumlah Resit Trade In (RM)
    End If
    If Not IsNull(rs!no_pekerja) Then
        Frm84_LM_No_PEKERJA = rs!no_pekerja
        Frm84_DATA_PEKERJA_FOUND = 1
    End If

'1:  Pelanggan
'2:  Member
'3:  RAF
'4:  Pengedar
'5:  Normal Dealer
'6:  Master Dealer

    If Not IsNull(rs!kategori_pembeli) Then 'Kategori Pembeli
        Frm84_LM_KATEGORI = rs!kategori_pembeli
    End If
    If Not IsNull(rs!no_rujukan_pembeli) Then
        Frm84_LM_No_CUST = rs!no_rujukan_pembeli
        Frm84_DATA_CUST_FOUND = 1 '0 : Pelanggan Tidak Berdaftar , 1 : Pelanggan Berdaftar , 2 : Ahli
    End If
    If Not IsNull(rs!tunai) Then
        G1 = G1 & "Tunai : RM " & Format(rs!tunai, "#,##0.00") & vbCrLf  'Tunai
    Else
        G1 = G1 & "Tunai : RM " & Format(0, "#,##0.00") & vbCrLf
    End If
    If Not IsNull(rs!bank_in) Then
        G1 = G1 & "Online Transfer : RM " & Format(rs!bank_in, "#,##0.00") & vbCrLf 'Online Banking
    Else
        G1 = G1 & "Online Transfer : RM " & Format(0, "#,##0.00") & vbCrLf 'Online Banking
    End If
    If Not IsNull(rs!kad_kredit) Then
        G1 = G1 & "Kad Kredit : RM " & Format(rs!kad_kredit, "#,##0.00") & vbCrLf 'Kad Kredit
    Else
        G1 = G1 & "Kad Kredit : RM " & Format(0, "#,##0.00") & vbCrLf
    End If

    If Not IsNull(rs!flag_bayaran) Then
        If rs!flag_bayaran = 1 Then
            Report83.Sections("Section1").Controls("L17").Visible = True
        End If
    End If
    
    DATA_FOUND = 1
End If

rs.Close
Set rs = Nothing

Report83.Sections("Section1").Controls("L15").Caption = G1 'Maklumat Bayaran
If LM_SUSUT_NILAI_MODE = 0 Then
    If LM_TI <> vbNullString Then
        LM_BERAT_TI = 0
        'LM_HARGA_TI = 0
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select SUM(berat) from data_database where bill_No_Trade_In='" & LM_TI & "' AND statusitem <> 0", cn, adOpenKeyset, adLockOptimistic
    
        If Not IsNull(rs(0)) Then LM_BERAT_TI = rs(0)
        
        rs.Close
        Set rs = Nothing
        
        H1 = H1 & "Maklumat Trade In" & vbCrLf
        H1 = H1 & "No. Payment Voucher : " & LM_TI & vbCrLf
        H1 = H1 & "Jumlah : RM " & Format(LM_HARGA_TI, "#,##0.00") & " (" & Format(LM_BERAT_TI, "#,##0.00 g") & ")" & vbCrLf
    End If
ElseIf LM_SUSUT_NILAI_MODE = 1 Then
    H1 = H1 & "Maklumat Trade In" & vbCrLf
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 93_trade_in_susut_niai where no_invoice='" & G_No_RESIT_JUALAN & "' AND status = 1 order by jenis ASC", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        LM_BERAT = 0
        LM_HARGA_SEMASA = 0
        LM_HARGA = 0
        LM_JENIS = vbNullString
        
        If Not IsNull(rs!Berat) Then LM_BERAT = rs!Berat
        If Not IsNull(rs!harga_Semasa) Then LM_HARGA_SEMASA = rs!harga_Semasa
        If Not IsNull(rs!harga) Then LM_HARGA = rs!harga
        If Not IsNull(rs!jenis) Then
            If rs!jenis = 0 Then LM_JENIS = "Trade In : "
            If rs!jenis = 1 Then LM_JENIS = "Buyback : "
            If rs!jenis = 2 Then LM_JENIS = "Caj Pertukaran : "
        End If
        If rs!jenis = 0 Or rs!jenis = 1 Then H1 = H1 & LM_JENIS & Format(LM_BERAT, "#,##0.00 g") & " X RM " & Format(LM_HARGA_SEMASA, "#,##0.00") & "/g = RM " & Format(LM_HARGA, "#,##0.00") & vbCrLf
        If rs!jenis = 2 Then H1 = H1 & LM_JENIS & " RM " & Format(LM_HARGA, "#,##0.00") & vbCrLf
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
End If

Report83.Sections("Section1").Controls("L13").Caption = H1 'Remarks

If DATA_FOUND = 1 Then
    If Frm84_DATA_PEKERJA_FOUND = 1 Then
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where nopekerja='" & Frm84_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!Samaran) Then Report83.Sections("Section1").Controls("L14").Caption = "Nama Jurujual : " & rs!Samaran  'Nama Samaran
        End If
        
        rs.Close
        Set rs = Nothing
    End If
    
'### Data jika pembeli TIDAK berdaftar ### - Start
    If Frm84_DATA_CUST_FOUND = 0 Then '0 : Pelanggan Tidak Berdaftar , 1 : Pelanggan Berdaftar , 2 : Ahli
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 44_senarai_pelanggan where no_resit='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!Nama) Then Report83.Sections("Section1").Controls("L1").Caption = rs!Nama 'Maklumat Pembeli : Nama
            If Not IsNull(rs!no_tel) Then Report83.Sections("Section1").Controls("L2").Caption = rs!no_tel 'Maklumat Pembeli : No. Telefon
        End If
        
        rs.Close
        Set rs = Nothing
    End If
'### Data jika pembeli TIDAK berdaftar ### - End

'### Data jika pembeli adalah berdaftar ### - Start
    If Frm84_DATA_CUST_FOUND = 1 Then '0 : Pelanggan Tidak Berdaftar , 1 : Pelanggan Berdaftar , 2 : Ahli
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm84_LM_No_CUST & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            If Not IsNull(rs!Nama) Then Report83.Sections("Section1").Controls("L1").Caption = rs!Nama 'Maklumat Pembeli : Nama
            If Not IsNull(rs!no_tel) Then Report83.Sections("Section1").Controls("L2").Caption = rs!no_tel 'Maklumat Pembeli : No. Telefon
           ' If Not IsNull(rs!no_pelanggan) Then Report83.Sections("Section1").Controls("L3").Caption = rs!no_pelanggan
            
        End If
        
        rs.Close
        Set rs = Nothing

    End If
'### Data jika pembeli adalah berdaftar ### - End

    '### Paparan Resit ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 22_jualan where no_resit='" & G_No_RESIT_JUALAN & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        Set Report83.DataSource = rs
        If G_PREVIEW = 1 Then Report83.Show
        rs.MoveNext
    Wend
    
    'rs.Close
    Set rs = Nothing
    '### Paparan Resit ### - End
    
    If G_PREVIEW = 0 Then Report83.PrintReport
     
    G_No_RESIT_JUALAN = vbNullString
End If
End Sub
Sub frm_kiraan_harga_selepas_ti()
'on error resume next
Dim Frm84_LM_HARGA As Double
Dim Frm84_LM_BUYBACK As Double
Dim Frm84_LM_HARGA_LAST As Double
Dim Frm84_LM_DEDUCT_RESIT As Double
Dim Frm84_LM_BERAT_JUALAN As Double
Dim Frm84_LM_BERAT_TI As Double
Dim Frm84_LM_BERAT_BAKI As Double
Dim Frm84_LM_HARGA_TI As Double
Dim Frm84_LM_HARGA_BUYBACK As Double
Dim Frm84_LM_CAJ As Double
Dim a As String
Dim Frm84_LM_TOTAL_TI As Double
Dim Frm84_LM_TOTAL_BUYBACK As Double
Dim Frm84_LM_TOTAL_CAJ As Double
Dim Frm84_LM_TOTAL_DEDUCT_TI As Double

Frm84_LM_TOTAL_TI = 0
Frm84_LM_TOTAL_BUYBACK = 0
Frm84_LM_TOTAL_CAJ = 0
Frm84_LM_TOTAL_DEDUCT_TI = 0

Frm84_LM_BERAT_JUALAN = 0
Frm84_LM_BERAT_TI = 0

If GLOBAL_DISABLE = 0 Then

    Frm84_LM_HARGA = 0
    Frm84_LM_BUYBACK = 0
    Frm84_LM_HARGA_LAST = 0
    Frm84_LM_TOLAKAN_RESIT = 0
    Frm84_LM_DEDUCT_RESIT = 0

    Frm84_LM_HARGA_TI = 0
    Frm84_LM_HARGA_BUYBACK = 0
    Frm84_LM_CAJ = 0
    
    G_TRADE_IN_TOTAL = 0
    G_TRADE_IN_CAJ = 0
    
    If G_TI_MODE = 3 Then
        
        If (Frm84.L15_Text <> vbNullString And IsNumeric(Frm84.L15_Text)) Then Frm84_LM_BERAT_JUALAN = Frm84.L15_Text
        Frm84_LM_BERAT_TI = G_TI_BERAT

        Frm84.L91_Text = vbNullString

        a = vbNullString
        a = "Maklumat Trade In (0%)" & vbCrLf
        a = a & "======================" & vbCrLf
        
        Frm84_LM_BERAT_BAKI = Frm84_LM_BERAT_TI - Frm84_LM_BERAT_JUALAN
        Frm84_LM_HARGA_TI = G_TI_TRADE_IN
        Frm84_LM_HARGA_BUYBACK = G_TI_BUYBACK
        Frm84_LM_CAJ = G_TI_CAJ
        
        If Frm84_LM_BERAT_BAKI > 0 Then
            
            Frm84_LM_TOTAL_TI = Frm84_LM_BERAT_JUALAN * Frm84_LM_HARGA_TI
            Frm84_LM_TOTAL_BUYBACK = Frm84_LM_BERAT_BAKI * Frm84_LM_HARGA_BUYBACK
            Frm84_LM_TOTAL_CAJ = Frm84_LM_CAJ 'Frm84_LM_BERAT_JUALAN * Frm84_LM_CAJ
            Frm84_LM_TOTAL_DEDUCT_TI = Frm84_LM_TOTAL_TI + Frm84_LM_TOTAL_BUYBACK
            
            a = a & "Trade In : " & Format(Frm84_LM_BERAT_JUALAN, "#,##0.00 g") & " X RM " & Format(Frm84_LM_HARGA_TI, "#,##0.00") & "/g = RM " & Format(Frm84_LM_BERAT_JUALAN * Frm84_LM_HARGA_TI, "#,##0.00") & vbCrLf
            a = a & "Buyback : " & Format(Frm84_LM_BERAT_BAKI, "#,##0.00 g") & " X RM " & Format(Frm84_LM_HARGA_BUYBACK, "#,##0.00") & "/g = RM " & Format(Frm84_LM_BERAT_BAKI * Frm84_LM_HARGA_BUYBACK, "#,##0.00") & vbCrLf
            a = a & "Caj Pertukaran : RM " & Format(Frm84_LM_CAJ, "#,##0.00") & vbCrLf
            
            G_TI_MEMORY(0, 0) = 0
            G_TI_MEMORY(1, 1) = Frm84_LM_BERAT_JUALAN
            G_TI_MEMORY(1, 2) = Frm84_LM_HARGA_TI
            G_TI_MEMORY(1, 3) = Frm84_LM_BERAT_JUALAN * Frm84_LM_HARGA_TI
            
            G_TI_MEMORY(2, 1) = Frm84_LM_BERAT_BAKI
            G_TI_MEMORY(2, 2) = Frm84_LM_HARGA_BUYBACK
            G_TI_MEMORY(2, 3) = Frm84_LM_BERAT_BAKI * Frm84_LM_HARGA_BUYBACK
            
            G_TI_MEMORY(3, 1) = 0 'Frm84_LM_BERAT_JUALAN
            G_TI_MEMORY(3, 2) = 0 'Frm84_LM_CAJ
            G_TI_MEMORY(3, 3) = Frm84_LM_CAJ 'Frm84_LM_BERAT_JUALAN * Frm84_LM_CAJ
            
        Else
            Frm84_LM_TOTAL_TI = Frm84_LM_BERAT_TI * Frm84_LM_HARGA_TI
            Frm84_LM_TOTAL_BUYBACK = 0
            Frm84_LM_TOTAL_CAJ = Frm84_LM_CAJ 'Frm84_LM_BERAT_TI * Frm84_LM_CAJ
            Frm84_LM_TOTAL_DEDUCT_TI = Frm84_LM_TOTAL_TI + Frm84_LM_TOTAL_BUYBACK
            
            a = a & "Trade In : " & Format(Frm84_LM_BERAT_TI, "#,##0.00 g") & " X RM " & Format(Frm84_LM_HARGA_TI, "#,##0.00") & "/g = RM " & Format(Frm84_LM_BERAT_TI * Frm84_LM_HARGA_TI, "#,##0.00") & vbCrLf
            a = a & "Caj Pertukaran : RM " & Format(Frm84_LM_CAJ, "#,##0.00") & vbCrLf

            G_TI_MEMORY(0, 0) = 1
            G_TI_MEMORY(1, 1) = Frm84_LM_BERAT_TI
            G_TI_MEMORY(1, 2) = Frm84_LM_HARGA_TI
            G_TI_MEMORY(1, 3) = Frm84_LM_BERAT_TI * Frm84_LM_HARGA_TI
            
            G_TI_MEMORY(2, 1) = 0 'Frm84_LM_BERAT_TI
            G_TI_MEMORY(2, 2) = 0 'Frm84_LM_CAJ
            G_TI_MEMORY(2, 3) = Frm84_LM_CAJ 'Frm84_LM_BERAT_TI * Frm84_LM_CAJ
            
        End If

        G_TRADE_IN_TOTAL = Frm84_LM_TOTAL_DEDUCT_TI
        G_TRADE_IN_CAJ = Frm84_LM_TOTAL_CAJ
            
        Frm84.L91_Text = a
        Frm84.L91_Text.Visible = True
    Else
        Frm84.L91_Text = vbNullString
        Frm84.L91_Text.Visible = False
    End If
    
    If ((Frm84.L21_Text <> vbNullString And IsNumeric(Frm84.L21_Text)) And (Frm84.L22_Text <> vbNullString And IsNumeric(Frm84.L22_Text))) Then
        Frm84_LM_HARGA = Frm84.L21_Text 'Harga Barang
        Frm84_LM_BUYBACK = Frm84.L22_Text 'Buyback
        
        Frm84_LM_HARGA_LAST = Frm84_LM_HARGA + Frm84_LM_TOTAL_CAJ - Frm84_LM_BUYBACK - Frm84_LM_TOTAL_DEDUCT_TI
        
        If Frm84_LM_HARGA_LAST >= 0 Then
            Frm84.L23_Text = Format(Frm84_LM_HARGA_LAST, "#,##0.00") 'Harga Perlu Bayar
            Frm84.TB33 = Format(Frm84_LM_HARGA_LAST, "#,##0.00") 'Harga Perlu Bayar
            Frm84.L37_Text = "0.00"
            Frm84.L24_Text = "Jumlah Bayaran"
            Frm84.L25_Text = "Jumlah Bayaran"
        Else
            Frm84.L23_Text = -Format(Frm84_LM_HARGA_LAST, "#,##0.00")  'Harga Perlu Bayar
            Frm84.TB33 = -Format(Frm84_LM_HARGA_LAST, "#,##0.00")  'Harga Perlu Bayar
            Frm84.L24_Text = "Harga Kedai Perlu Bayar Pelanggan"
            Frm84.L25_Text = "Harga Kedai Perlu Bayar Pelanggan"
            Frm84_LM_TOLAKAN_RESIT = 1
        End If
        
    Else
        Frm84.L23_Text = "0.00" 'Harga Perlu Bayar
        Frm84.TB33 = "0.00" 'Harga Perlu Bayar
        Frm84.L37_Text = "0.00"
        Frm84.L24_Text = "Jumlah Bayaran"
        Frm84.L25_Text = "Jumlah Bayaran"
    End If
    
    If Frm84_LM_TOLAKAN_RESIT = 1 Then
        If IsNumeric(Frm84.L38_Text) Then
            Frm84_LM_DEDUCT_RESIT = Frm84.L38_Text 'Potongan Harga Resit Trade in (%)
            
            Frm84_LM_JUMLAH_POTONG = (Frm84_LM_DEDUCT_RESIT / 100) * (-Frm84_LM_HARGA_LAST)
            
            Frm84.L37_Text = Format(Frm84_LM_JUMLAH_POTONG, "0.00")
            Frm84.L23_Text = Format((-Frm84_LM_HARGA_LAST) - Frm84_LM_JUMLAH_POTONG, "#,##0.00")
            Frm84.TB33 = Format((-Frm84_LM_HARGA_LAST) - Frm84_LM_JUMLAH_POTONG, "#,##0.00")
        End If
    End If
    
End If
End Sub

