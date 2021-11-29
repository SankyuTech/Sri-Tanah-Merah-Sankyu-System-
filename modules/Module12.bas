Attribute VB_Name = "Module12"
Sub tesutochu()
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
Dim Frm84_KOMISYEN_UPAH As Double 'Komisyen dari upah kepada agen dropship
Dim Frm84_LM_SUSUT_BERAT As Double
Dim Frm84_LM_BERAT_OVERALL As Double
Dim Frm84_LM_HARGA_JUALAN_DENGAN_GST As Double
Dim Frm84_LM_HARGA_JUALAN_TANPA_GST As Double
Dim Frm84_LM_MODAL_TANPA_GST As Double
Dim Frm84_LM_MODAL_DENGAN_GST As Double
Dim LM_HARGA_JUALAN_DGN_GST As Double
Dim LM_GST_JUAL As Double
Dim LM_MODAL_DGN_GST As Double
Dim LM_MODAL_TANPA_GST As Double
Dim LM_MODAL_TANPA_GST_GRAM As Double

LM_HARGA_JUALAN_DGN_GST = 0
LM_GST_JUAL = 0
LM_MODAL_DGN_GST = 0
LM_MODAL_TANPA_GST = 0
LM_MODAL_TANPA_GST_GRAM = 0

Frm84_LM_HARGA_JUALAN_DENGAN_GST = 0
Frm84_LM_MODAL_TANPA_GST = 0
Frm84_LM_MODAL_DENGAN_GST = 0
Frm84_LM_HARGA_JUALAN_TANPA_GST = 0
                
Frm84_LM_BERAT_OVERALL = 0
Frm84_LM_SUSUT_BERAT = 0
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
Frm84_LM_HARGA_JUALAN_CALC = 0 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
Frm84_LM_GST_CALC = 0 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
Frm84_KOMISYEN_UPAH = 0 'Komisyen dari upah kepada agen dropship

If Frm84.L83_Text = "0" Then '0 : Stok kedai , 1 : Barang trade in/potong
    If Frm84.TB2 = vbNullString Then
        x = x + 1
        Err(x) = "Sila Masukkan [No. Siri Produk]."
    End If
End If
If Frm84.TB2 <> vbNullString And Frm84.TB3 = vbNullString And Frm84.CB12 = 1 Then
    MsgBox "Tetapan GST ke atas UPAH hanya dibenarkan untuk barang kemas SAHAJA. Sila periksa tetapan GST anda.", vbExclamation, "Info"
    Exit Sub
End If
'If (Frm84.TB14 <> vbNullString And IsNumeric(Frm84.TB14)) And (Frm84.L51_Text <> vbNullString And IsNumeric(Frm84.L51_Text)) Then
'    Frm84_LM_HARGA_STAFF = Frm84.L51_Text
'    Frm84_LM_HARGA_PELANGGAN = Frm84.TB14
    
'    If Frm84_LM_HARGA_PELANGGAN < Frm84_LM_HARGA_STAFF Then
'        X = X + 1
'        Err(X) = "Harga Jualan Minimum Yang Dibenarkan Adalah RM " & Format(Frm84_LM_HARGA_STAFF, "#,##0.00")
'    End If
'End If
    
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
        If Frm84.TB43 = vbNullString Or (Frm84.TB43 <> vbNullString And Not IsNumeric(Frm84.TB43)) Then
            x = x + 1
            Err(x) = "Sila Masukkan [Kadar Komisyen Upah (%)]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
        End If
        If Frm84.TB44 = vbNullString Or (Frm84.TB44 <> vbNullString And Not IsNumeric(Frm84.TB44)) Then
            x = x + 1
            Err(x) = "Sila Masukkan [Jumlah Komisyen Bagi Upah]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
        End If
        If (Frm84.TB15 <> vbNullString And IsNumeric(Frm84.TB15)) And (Frm84.TB44 <> vbNullString And IsNumeric(Frm84.TB44)) Then
            Frm84_UPAH_JUAL = Frm84.TB15 'Upah jualan
            Frm84_KOMISYEN_UPAH = Frm84.TB44 'Komisyen Upah
            
            If Frm84_KOMISYEN_UPAH > Frm84_UPAH_JUAL Then
                x = x + 1
                Err(x) = "Komisyen upah bagi agen dropship adalah melebihi dari upah asal."
            End If
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
If Frm84.CB4 = 0 And Frm84.CB5 = 0 And Frm84.CB6 = 0 And Frm84.CB9 = 0 And Frm84.CB10 = 0 Then
    x = x + 1
    Err(x) = "Sila Buat Pilihan Kategori Pembeli."
End If
If Frm84.CB2 = 0 And Frm84.CB3 = 0 And Frm84.CB18 = 0 Then
    x = x + 1
    Err(x) = "Sila Buat Pilihan Jenis GST."
End If
If Frm84.TB10 = vbNullString Or (Frm84.TB10 <> vbNullString And Not IsNumeric(Frm84.TB10)) Then
    x = x + 1
    Err(x) = "Tiada maklumat [Harga Jualan]. Sila keluarkan item ini dari senarai dan scan sekali lagi."
End If
If Frm84.TB11 = vbNullString Or (Frm84.TB11 <> vbNullString And Not IsNumeric(Frm84.TB11)) Then
    x = x + 1
    Err(x) = "Tiada maklumat [Jumlah GST]. Sila keluarkan item ini dari senarai dan scan sekali lagi."
End If
If (Frm84.TB3 <> vbNullString And IsNumeric(Frm84.TB3)) And (Frm84.TB4 <> vbNullString And IsNumeric(Frm84.TB4)) Then
    Frm84_LM_BERAT_ASAL = Frm84.TB3 'Berat Asal
    Frm84_LM_BERAT_JUAL = Frm84.TB4 'Berat Jualan
    
    If Format(Frm84_LM_BERAT_JUAL, "0.00") = "0.00" Then
        x = x + 1
        Err(x) = "Berat jualan yang tidak sah Nilai 0 tidak dibenarkan di dalam ruangan ini."
    End If
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
If Frm84.L70_Text = "0" Then
    If Frm84.TB22 = vbNullString Or (Frm84.TB22 <> vbNullString And Not IsNumeric(Frm84.TB22)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Upah per gram]. Hanya NOMBOR Dibenarkan Dalam Ruangan Ini."
    End If
End If

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else
    Note = "Adakah Anda Ingin Masukkan Data Ini ?"
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbNo Then
        Exit Sub
    End If
    If Answer = vbYes Then
        
        LM_BARU_TI = 0 '0 : Barang baru , 1 : Barang Trade In
        
'### Periksa Data Dulang ### - Start
        If Frm84.L83_Text = "0" Then '0 : Stok kedai , 1 : Barang trade in/potong
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from Data_Database where no_siri_Produk='" & Frm84.TB2 & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
            
                If Not IsNull(rs!dulang) Then Frm84_LM_DULANG = rs!dulang 'Dulang
                If Not IsNull(rs!susut_berat) Then Frm84_LM_SUSUT_BERAT = rs!susut_berat 'Susut berat
                If Not IsNull(rs!receiving_Status) Then
                    
                    If rs!receiving_Status = "2" Or rs!receiving_Status = "3" Or rs!receiving_Status = "6" Or rs!receiving_Status = "7" Then LM_BARU_TI = 1 '0 : Barang baru , 1 : Barang Trade In
                    
                End If
                
            End If
            
            rs.Close
            Set rs = Nothing
            
        ElseIf Frm84.L83_Text = "1" Then
        
        
        
        
        End If
'### Periksa Data Dulang ### - End

        If Frm84.TB3 <> vbNullString And Frm84.TB4 <> vbNullString Then
        
            If IsNumeric(Frm84.TB3) Then Frm84_LM_BERAT_ASAL = Frm84.TB3 'Berat Asal
            If IsNumeric(Frm84.TB4) Then Frm84_LM_BERAT_JUAL = Frm84.TB4 'Berat Jualan
            
            Frm84_LM_BERAT_OVERALL = Frm84_LM_SUSUT_BERAT + Frm84_LM_BERAT_JUAL
            
            If Frm84_LM_BERAT_ASAL < Frm84_LM_BERAT_JUAL Then
            
                'MsgBox "Berat jualan melebihi berat jualan yang dibenarkan." & vbCrLf & _
                        "Berat asal : " & Format(Frm84_LM_BERAT_ASAL, "#,##0.00 g") & vbCrLf & _
                        "Susut berat : " & Format(Frm84_LM_SUSUT_BERAT, "#,##0.00 g") & vbCrLf & _
                        "Berat jualan maksimum yang dibenarkan adalah " & Format(Frm84_LM_BERAT_ASAL - Frm84_LM_SUSUT_BERAT, "#,##0.00 g"), vbInformation, "Info"
                        
                MsgBox "Berat jualan melebihi berat jualan yang dibenarkan." & vbCrLf & _
                        "Berat asal : " & Format(Frm84_LM_BERAT_ASAL, "#,##0.00 g") & vbCrLf & _
                        "Berat jualan maksimum yang dibenarkan adalah " & Format(Frm84_LM_BERAT_ASAL, "#,##0.00 g"), vbInformation, "Info"
                        
                Exit Sub
                        
            End If
            
        End If
    
'### Periksa Kadar Penurunan Harga ### - Start
'GoTo skip_periksa_harga:
        user = MDI_frm1.L3_Text
        
        If MDI_frm1.L4_Text <> vbNullString Then
            If MDI_frm1.L4_Text = "Staff" Then
                Frm84_LM_PRICE_CHECK = 1 '0 : Tidak Perlu Periksa Harga Semasa Jualan , 1 : Perlu Periksa Harga Semasa Jualan
            End If
        End If
'skip_periksa_harga:

'Frm84.L84_Text : 0 : Tiada tetapan untuk penentuan cukai GST ZR , 1 : Ada tetapan bagi penentuan cukai GST ZR

        If Frm84.CB13 = 0 And Frm84.CB2 = 1 And Frm84.L84_Text = "1" And Frm84.L85_Text = "0" Then
            
            Note = "Anda cuba menjual barang ini tanpa cukai GST." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sistem akan menukarkan jenis invoice jualan ini kepada TIDAK RASMI." & vbCrLf & _
                    "*** Invoice TIDAK RASMI adalah invoice yang tidak akan dikira sebagai jualan rasmi kedai." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila pilih [Yes] untuk meneruskan jualan ini dengan invoice tidak rasmi dan pilih [No] jika ingin meneruskan jualan dengan invoice rasmi."
                    
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbNo Then
                Exit Sub
            End If
            If Answer = vbYes Then

                Frm84.CB13 = 1
                
            End If
            
        End If
        
        If Frm84_LM_PRICE_CHECK = 1 Then '0 : Tidak Perlu Periksa Harga Semasa Jualan , 1 : Perlu Periksa Harga Semasa Jualan
            Frm84_LM_LIMIT_TYPE = 0 '1 : BK , 2 : Barang Permata
            
'### Periksa Purity Dan Tetapan Harga Jualan ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from Data_Database where no_siri_Produk='" & Frm84.TB2 & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!dulang) Then Frm84_LM_DULANG = rs!dulang 'Dulang
                
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
                        ElseIf Frm84.CB9 = 1 Then
                            If IsNumeric(rs!HargaJualan_RAF) Then Frm84_LM_TETAPANHARGA = Format(rs!HargaJualan_RAF, "0.00") 'Harga RAF
                        ElseIf Frm84.CB6 = 1 Then
                            If IsNumeric(rs!HargaJualan_Pengedar) Then Frm84_LM_TETAPANHARGA = Format(rs!HargaJualan_Pengedar, "0.00") 'Harga Pengedar
                        ElseIf Frm84.CB10 = 1 Then
                            If IsNumeric(rs!hargajualan_normal_dealer) Then Frm84_LM_TETAPANHARGA = Format(rs!hargajualan_normal_dealer, "0.00") 'Harga Normal Dealer
                        'ElseIf Frm84.CB11 = 1 Then
                        '    If IsNumeric(rs!hargajualan_master_dealer) Then Frm84_LM_TETAPANHARGA = Format(rs!hargajualan_master_dealer, "0.00") 'Harga Master Dealer
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
                    ElseIf Frm84.CB9 = 1 Then
                        If IsNumeric(rs!Harga_RAF) Then Frm84_LM_TETAPANHARGA = Format(rs!Harga_RAF, "0.00") 'Harga RAF
                    ElseIf Frm84.CB6 = 1 Then
                        If IsNumeric(rs!Harga_Pengedar) Then Frm84_LM_TETAPANHARGA = Format(rs!Harga_Pengedar, "0.00") 'Harga Pengedar
                    ElseIf Frm84.CB10 = 1 Then
                        If IsNumeric(rs!harga_normal_dealer) Then Frm84_LM_TETAPANHARGA = Format(rs!harga_normal_dealer, "0.00") 'Harga Normal Dealer
                    'ElseIf Frm84.CB11 = 1 Then
                    '    If IsNumeric(rs!harga_master_dealer) Then Frm84_LM_TETAPANHARGA = Format(rs!harga_master_dealer, "0.00") 'Harga Master Dealer
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
                rs.Open "select * from default_setting where Default1='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    'If rs!Default1 = "Default" Then
                        If IsNumeric(rs!limit_per_item) Then Frm84_LM_LIMIT = rs!limit_per_item
                    'End If
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

        If Frm84.L83_Text = "1" Then '0 : Stok kedai , 1 : Barang trade in/potong
            
            LM_PURITY_FOUND = 0
            
            For c = 1 To 20
            
                'If G_PURITY_JUALAN(c) Then
                
                    If Frm84.L13_Text = G_PURITY_JUALAN(c) Then
                    
                        LM_PURITY_FOUND = 1
                        GoTo skip_b:
                        
                    End If
                    
                'End If
                
                If LM_PURITY_FOUND = 0 Then
                    
                    If G_BIL_JUALAN < 20 Then
                    
                        G_BIL_JUALAN = G_BIL_JUALAN + 1
                    
                        G_PURITY_JUALAN(G_BIL_JUALAN) = Frm84.L13_Text
                    
                    End If
                    
                End If
                
            Next c
            
skip_b:

        End If

'### Masukkan Data Ke Dalam Temp Table ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_JUALAN_TEMP & " where no_siri_Produk='" & Frm84.TB2 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
            rs.AddNew
            If Frm84.L83_Text = "0" Then '0 : Stok kedai , 1 : Barang trade in/potong
                If Frm84.TB2 <> vbNullString Then
                    rs!no_siri_Produk = Frm84.TB2 'No. Siri Produk
                Else
                    rs!no_siri_Produk = Null 'No. Siri Produk
                End If
                rs!nama_purity = Null
                rs!dulang = Frm84_LM_DULANG 'Dulang
                
                If Frm84.TB3 <> vbNullString Then
                    rs!Berat_Asal = Format(Frm84.TB3, "0.00") 'Berat Asal (g)
                Else
                    rs!Berat_Asal = Null 'Berat Asal (g)
                End If
            
            Else
            
                rs!no_siri_Produk = "-" 'No. Siri Produk
                If Frm84.CBB4 <> vbNullString Then
                    rs!nama_purity = Frm84.CBB4
                Else
                    rs!nama_purity = Null
                End If
                rs!dulang = "-" 'Dulang
                
                If Frm84.TB4 <> vbNullString Then
                    rs!Berat_Asal = Format(Frm84.TB4, "0.00") 'Berat Asal (g)
                Else
                    rs!Berat_Asal = Null 'Berat Asal (g)
                End If
            
            End If
            If Frm84.L83_Text = "0" Then '0 : Stok kedai , 1 : Barang trade in/potong
                rs!flag_barang = 0
            ElseIf Frm84.L83_Text = "1" Then '0 : Stok kedai , 1 : Barang trade in/potong
                rs!flag_barang = 1
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
            'If Frm84.TB3 <> vbNullString Then
            '    rs!Berat_Asal = Format(Frm84.TB3, "0.00") 'Berat Asal (g)
            'Else
            '    rs!Berat_Asal = Null 'Berat Asal (g)
            'End If
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
            
            'If Frm84.CB2 = 1 Then
            '    rs!gst_ari_nashi = "ZR (L)" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
            '    rs!kadar_gst = Null 'Kadar Cukai GST (%)
            '    If Frm84.TB11 <> vbNullString Then
            '        rs!jumlah_gst = Format(Frm84.TB11, "0.00") 'Jumlah Cukai GST (RM)
            '    Else
            '        rs!jumlah_gst = "0.00" 'Jumlah Cukai GST (RM)
            '    End If
            'ElseIf Frm84.CB3 = 1 Then
            '    rs!gst_ari_nashi = "SR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
            '    If Frm84.L8_Text <> vbNullString Then
            '        rs!kadar_gst = Frm84.L8_Text 'Kadar Cukai GST (%)
            '    Else
            '        rs!kadar_gst = "0" 'Jumlah Cukai GST (RM)
            '    End If
            '    If Frm84.TB11 <> vbNullString Then
            '        rs!jumlah_gst = Format(Frm84.TB11, "0.00") 'Jumlah Cukai GST (RM)
            '    Else
            '        rs!jumlah_gst = "0.00" 'Jumlah Cukai GST (RM)
            '    End If
            '    If Frm84.CB18 = 1 Then 'Jenis Cukai GST SR
            '        rs!gst_include = "**Harga Termasuk GST" '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
            '    Else
            '        rs!gst_include = Null '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
            '    End If
            'End If
            
            If Frm84.CB2 = 1 Then
            
                rs!gst_ari_nashi = "ZR (L)" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                rs!kadar_gst = Null 'Kadar Cukai GST (%)
                If Frm84.TB11 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm84.TB11, "0.00") 'Jumlah Cukai GST (RM)
                    LM_GST_JUAL = Format(Frm84.TB11, "0.00")
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah Cukai GST (RM)
                End If
                
                rs!gst_include = Null '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                
                'If Frm84.L85_Text = "0" Then
                '    If Frm84.L84_Text = "1" Then Frm84.CB13 = 1
                'End If
                
            ElseIf Frm84.CB3 = 1 Then
            
                rs!gst_ari_nashi = "SR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                If Frm84.L8_Text <> vbNullString Then
                    rs!kadar_gst = Frm84.L8_Text 'Kadar Cukai GST (%)
                Else
                    rs!kadar_gst = "0" 'Jumlah Cukai GST (RM)
                End If
                If Frm84.TB11 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm84.TB11, "0.00") 'Jumlah Cukai GST (RM)
                    LM_GST_JUAL = Format(Frm84.TB11, "0.00")
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah Cukai GST (RM)
                End If

                rs!gst_include = Null '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang

            ElseIf Frm84.CB18 = 1 Then
            
                rs!gst_ari_nashi = "SR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                If Frm84.L8_Text <> vbNullString Then
                    rs!kadar_gst = Frm84.L8_Text 'Kadar Cukai GST (%)
                Else
                    rs!kadar_gst = "0" 'Jumlah Cukai GST (RM)
                End If
                If Frm84.TB11 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm84.TB11, "0.00") 'Jumlah Cukai GST (RM)
                    LM_GST_JUAL = Format(Frm84.TB11, "0.00")
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah Cukai GST (RM)
                End If
                
                rs!gst_include = 1 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang

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
                    If Frm84.TB43 <> vbNullString Then
                        rs!kadar_komisyen_upah = Frm84.TB43 'Kadar komisyen bagi upah kepada agen dropship
                    Else
                        rs!kadar_komisyen_upah = Null 'Kadar komisyen bagi upah kepada agen dropship
                    End If
                    If Frm84.TB44 <> vbNullString Then
                        rs!komisyen_upah = Format(Frm84.TB44, "0.00") 'Jumlah komisyen bagi upah kepada agen dropship
                    Else
                        rs!komisyen_upah = Null 'Jumlah komisyen bagi upah kepada agen dropship
                    End If
                End If
                If Frm84.Frame3.Visible = True Then 'Komisen Agen Dropship : Permata
                    rs!komisyen_per_gram = Null 'Komisyen Per Gram Dropship (RM/g) : BK
                    rs!kadar_komisyen_upah = Null 'Kadar komisyen bagi upah kepada agen dropship
                    rs!komisyen_upah = Null 'Jumlah komisyen bagi upah kepada agen dropship
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
                rs!kadar_komisyen_upah = Null 'Kadar komisyen bagi upah kepada agen dropship
                rs!komisyen_upah = Null 'Jumlah komisyen bagi upah kepada agen dropship
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
                    LM_MODAL_DGN_GST = Format(Frm84.L34_Text, "0.00")
                Else
                    rs!modal = Null 'Harga Modal (RM)
                End If
                If Frm84.L42_Text <> vbNullString Then
                    rs!modal_tanpa_gst = Format(Frm84.L42_Text, "0.00") 'Harga Modal Tanpa GST (RM)
                    LM_MODAL_TANPA_GST = Frm84.L42_Text
                Else
                    rs!modal_tanpa_gst = Null 'Harga Modal (RM)
                End If
                If Frm84.L44_Text <> vbNullString Then
                    If IsNumeric(Frm84.L44_Text) Then Frm84_LM_HARGA_JUALAN_TANPA_GST = Frm84.L44_Text
                End If
                
                If Frm84.CB3 = 1 And Frm84.CB18 = 0 Then
                    Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
                    Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
                    
                    Frm84_LM_HARGA_JUALAN_DENGAN_GST = Format(Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC, "0.00")
                    'Field ini adalah lebih kurang kepada @harga_dengan_gst
                    'Tetapi bagi gst yang dikenakan hanya pada UPAH akan membuatkan field ini hanya memaparkan harga dengan gst hanya pada upah.
                    'Oleh itu field ini dibuat khas bagi mencampurkan gst pada upah dan harga barang.
                Else
                    Frm84_LM_HARGA_JUALAN_DENGAN_GST = Format(Frm84.TB10, "0.00") 'Harga Jualan (RM)
                End If
                
                If IsNumeric(Frm84.L34_Text) Then Frm84_LM_MODAL_DENGAN_GST = Frm84.L34_Text 'Harga Modal
                
                rs!jualan_per_gram_dengan_gst = Null
                rs!untung = Format(Frm84_LM_HARGA_JUALAN_DENGAN_GST - Frm84_LM_MODAL_DENGAN_GST, "0.00") 'Jumlah Keuntungan (pengiraan termasuk GST)
                rs!untung2 = Format(Frm84_LM_HARGA_JUALAN_TANPA_GST - Frm84_LM_MODAL_TANPA_GST, "0.00") 'Jumlah Keuntungan (pengiraan TIDAK termasuk GST)
                
                rs!harga_per_gram_supplier = Null 'Harga per gram (harga semasa) dari supplier (modal)
                rs!upah_modal = Null 'Upah modal
                rs!harga_per_gram_tanpa_gst = Null 'Harga modal per gram tanpa GST (RM)
                
            Else
            
                rs!Type = 0 '0 : BK , 1 : Barang Permata
                
                If Frm84.L34_Text <> vbNullString Then
                    rs!harga_per_gram_modal = Format(Frm84.L34_Text, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
                    If IsNumeric(Frm84.L34_Text) Then
                        Frm84_LM_HARGA_SEMASA_MODAL = Frm84.L34_Text 'Harga Per Gram Bagi Modal (RM/g)
                        
                        rs!modal = Format(Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_SEMASA_MODAL, "0.00") 'Harga Modal (RM)
                        LM_MODAL_DGN_GST = Format(Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_SEMASA_MODAL, "0.00")
                    End If
                Else
                    rs!modal = Null 'Harga Modal (RM)
                    rs!harga_per_gram_modal = Null 'Harga Per Gram Bagi Modal (RM/g)
                End If
                If Frm84.L42_Text <> vbNullString Then
                    rs!harga_per_gram_tanpa_gst = Format(Frm84.L42_Text, "0.00") 'Harga modal per gram tanpa GST (RM)
                    LM_MODAL_TANPA_GST_GRAM = Frm84.L42_Text
                    LM_MODAL_TANPA_GST = Format(Frm84_LM_BERAT_JUAL * LM_MODAL_TANPA_GST_GRAM, "0.00")
                Else
                    rs!harga_per_gram_tanpa_gst = Null 'Harga modal per gram tanpa GST (RM)
                End If
                
                If IsNumeric(Frm84.TB14) And IsNumeric(Frm84.L34_Text) And IsNumeric(Frm84.TB10) And IsNumeric(Frm84.TB11) Then
                    
                    Frm84_LM_HARGA_MODAL = Frm84.L34_Text 'Harga Modal
                    Frm84_LM_HARGA_JUAL = Frm84.TB14 'Harga Jualan
                    Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
                    Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
                    
                    If Frm84.L34_Text <> vbNullString Then
                        If IsNumeric(Frm84.L34_Text) Then Frm84_LM_MODAL_DENGAN_GST = Frm84.L34_Text 'Harga Modal
                    End If
                    
                    If Frm84.L42_Text <> vbNullString Then
                        If IsNumeric(Frm84.L42_Text) Then Frm84_LM_MODAL_TANPA_GST = Frm84.L42_Text
                    End If
                        
                    If Frm84.CB12 = 0 Then
                        
                        If Frm84.TB14 <> vbNullString Then
                            If IsNumeric(Frm84.TB14) Then Frm84_LM_HARGA_JUALAN_DENGAN_GST = Frm84.TB14 'Harga Jualan
                        End If
                        
                        If Frm84.L44_Text <> vbNullString Then
                            If IsNumeric(Frm84.L44_Text) Then Frm84_LM_HARGA_JUALAN_TANPA_GST = Frm84.L44_Text
                        End If
                        
                        rs!jualan_per_gram_dengan_gst = Format(Frm84_LM_HARGA_JUALAN_DENGAN_GST / Frm84_LM_BERAT_JUAL, "0.00")
                        rs!untung = Format((Frm84_LM_HARGA_JUALAN_DENGAN_GST) - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_DENGAN_GST), "0.00") 'Jumlah Keuntungan (pengiraan termasuk GST)
                        rs!untung2 = Format(Frm84_LM_HARGA_JUALAN_TANPA_GST - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_TANPA_GST), "0.00") 'Jumlah Keuntungan (pengiraan TIDAK termasuk GST)

                    ElseIf Frm84.CB12 = 1 Then
                        
                        If Frm84.CB2 = 1 Then
                        
                            If Frm84.TB10 <> vbNullString Then
                                If IsNumeric(Frm84.TB10) Then Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
                            End If
                             
                            If Frm84.TB11 <> vbNullString Then
                                Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
                            End If
                            
                            rs!jualan_per_gram_dengan_gst = Format((Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC) / Frm84_LM_BERAT_JUAL, "0.00")
                            rs!untung = Format((Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC) - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_DENGAN_GST), "0.00") 'Jumlah Keuntungan (pengiraan termasuk GST)
                            rs!untung2 = Format(Frm84_LM_HARGA_JUALAN_CALC - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_TANPA_GST), "0.00") 'Jumlah Keuntungan (pengiraan TIDAK termasuk GST)
                            
                        ElseIf Frm84.CB3 = 1 Then
                            
                            If Frm84.TB10 <> vbNullString Then
                                If IsNumeric(Frm84.TB10) Then Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
                            End If
                             
                            If Frm84.TB11 <> vbNullString Then
                                Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
                            End If
                            
                            rs!jualan_per_gram_dengan_gst = Format((Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC) / Frm84_LM_BERAT_JUAL, "0.00")
                            rs!untung = Format((Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC) - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_DENGAN_GST), "0.00") 'Jumlah Keuntungan (pengiraan termasuk GST)
                            rs!untung2 = Format(Frm84_LM_HARGA_JUALAN_CALC - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_TANPA_GST), "0.00") 'Jumlah Keuntungan (pengiraan TIDAK termasuk GST)

                        ElseIf Frm84.CB18 = 1 Then
                        
                            If Frm84.TB10 <> vbNullString Then
                                If IsNumeric(Frm84.TB10) Then Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
                            End If
                             
                            If Frm84.TB11 <> vbNullString Then
                                Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
                            End If
                            
                            rs!jualan_per_gram_dengan_gst = Format((Frm84_LM_HARGA_JUALAN_CALC) / Frm84_LM_BERAT_JUAL, "0.00")
                            rs!untung = Format((Frm84_LM_HARGA_JUALAN_CALC) - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_DENGAN_GST), "0.00")  'Jumlah Keuntungan (pengiraan termasuk GST)
                            rs!untung2 = Format((Frm84_LM_HARGA_JUALAN_CALC - Frm84_LM_GST_CALC) - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_TANPA_GST), "0.00") 'Jumlah Keuntungan (pengiraan TIDAK termasuk GST)

                        End If
                        
                    End If
                    
                End If
                
                'If IsNumeric(Frm84.TB4) And IsNumeric(Frm84.TB5) And IsNumeric(Frm84.L54_Text) And IsNumeric(Frm84.L55_Text) And IsNumeric(Frm84.TB15) And IsNumeric(Frm84.TB3) Then
                    
                ''    Frm84_LM_BERAT_JUAL = Frm84.TB4 'Berat Jualan
                '    Frm84_LM_HARGA_SEMASA = Frm84.TB5 'Harga semasa (jualan)
                '    Frm84_LM_HARGA_SUPPLIER = Frm84.L54_Text 'Harga per gram (harga semasa) dari supplier (modal)
                '    Frm84_UPAH_MODAL = Frm84.L55_Text 'Upah modal
                '    Frm84_UPAH_JUAL = Frm84.TB15 'Upah jualan
                '    Frm84_LM_BERAT_ASAL = Frm84.TB3 'Berat Asal
                    
                '    rs!upah_modal = Frm84.L55_Text 'Upah modal
                '    rs!harga_per_gram_supplier = Frm84.L54_Text 'Harga per gram (harga semasa) dari supplier (modal)
                '    rs!untung2 = Format(((Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_SEMASA) + Frm84_UPAH_JUAL) - ((Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_SUPPLIER) + (Frm84_LM_BERAT_JUAL * Frm84_UPAH_MODAL / Frm84_LM_BERAT_ASAL)), "0.00") 'Untung jika restok pada harga supplier ini

                'Else
                    
                '    rs!harga_per_gram_supplier = "0.00" 'Harga per gram (harga semasa) dari supplier (modal)
                '    rs!untung2 = "0.00" 'Untung jika restok pada harga supplier ini
                '    rs!upah_modal = "0.00" 'Upah modal
                    
                'End If
                
            End If
            If Format(Frm84.TB3, "0.00") = Format(Frm84.TB4, "0.00") Then
                rs!potong_flag = 0 '0 : Tiada Potong , 1 : Ada Potong
                rs!status_jualan = 0 '0 : Tiada Potong , 1 : Ada Potong
            Else
                rs!potong_flag = 1 '0 : Tiada Potong , 1 : Ada Potong
                rs!status_jualan = 1 '0 : Tiada Potong , 1 : Ada Potong
            End If
            'rs!dulang = Frm84_LM_DULANG 'Dulang
            
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
                LM_HARGA_JUALAN_DGN_GST = Format(Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC, "0.00")
                'Field ini adalah lebih kurang kepada @harga_dengan_gst
                'Tetapi bagi gst yang dikenakan hanya pada UPAH akan membuatkan field ini hanya memaparkan harga dengan gst hanya pada upah.
                'Oleh itu field ini dibuat khas bagi mencampurkan gst pada upah dam harga barang.
            Else
                rs!harga_jualan_dengan_gst = Format(Frm84.TB10, "0.00") 'Harga Jualan (RM)
                LM_HARGA_JUALAN_DGN_GST = Format(Frm84.TB10, "0.00") 'Harga Jualan (RM)
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
            If Frm84.L70_Text = "0" Then
                
                rs!flag_upah = 0
                
                If Frm84.TB22 <> vbNullString Then
                
                    rs!upah_per_gram = Format(Frm84.TB22, "0.00")
                
                Else
                
                    rs!upah_per_gram = "0.00"
                    
                End If
            
            ElseIf Frm84.L70_Text = "1" Then
                
                rs!flag_upah = 1
                rs!upah_per_gram = Null
            
            End If
            rs!harga_jual_excl_gst = Format(LM_HARGA_JUALAN_DGN_GST - LM_GST_JUAL, "0.00")
            rs!harga_modal_gst = Format(LM_MODAL_DGN_GST - LM_MODAL_TANPA_GST, "0.00")
            rs!harga_modal_incl_gst = Format(LM_MODAL_DGN_GST, "0.00")
            rs!harga_modal_excl_gst = Format(LM_MODAL_TANPA_GST, "0.00")
            
            rs!untung = Format(LM_HARGA_JUALAN_DGN_GST - LM_GST_JUAL - LM_MODAL_TANPA_GST, "0.00")
            rs!untung2 = Format(LM_HARGA_JUALAN_DGN_GST - LM_MODAL_DGN_GST, "0.00")
            rs!baru_or_ti = LM_BARU_TI
            
            rs.Update
            Frm84_LM_DATA_SAVE = 1
        Else
            If Frm84.L83_Text = "0" Then '0 : Stok kedai , 1 : Barang trade in/potong
                If Frm84.TB2 <> vbNullString Then
                    rs!no_siri_Produk = Frm84.TB2 'No. Siri Produk
                Else
                    rs!no_siri_Produk = Null 'No. Siri Produk
                End If
                rs!nama_purity = Null
                rs!dulang = Frm84_LM_DULANG 'Dulang
                
                If Frm84.TB3 <> vbNullString Then
                    rs!Berat_Asal = Format(Frm84.TB3, "0.00") 'Berat Asal (g)
                Else
                    rs!Berat_Asal = Null 'Berat Asal (g)
                End If
                
            Else
                rs!no_siri_Produk = "-" 'No. Siri Produk
                If Frm84.CBB4 <> vbNullString Then
                    rs!nama_purity = Frm84.CBB4
                Else
                    rs!nama_purity = Null
                End If
                rs!dulang = "-" 'Dulang
                
                If Frm84.TB4 <> vbNullString Then
                    rs!Berat_Asal = Format(Frm84.TB4, "0.00") 'Berat Asal (g)
                Else
                    rs!Berat_Asal = Null 'Berat Asal (g)
                End If
                
            End If
            If Frm84.L83_Text = "0" Then '0 : Stok kedai , 1 : Barang trade in/potong
                rs!flag_barang = 0
            ElseIf Frm84.L83_Text = "1" Then '0 : Stok kedai , 1 : Barang trade in/potong
                rs!flag_barang = 1
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
            'If Frm84.TB3 <> vbNullString Then
            '    rs!Berat_Asal = Format(Frm84.TB3, "0.00") 'Berat Asal (g)
            'Else
            '    rs!Berat_Asal = Null 'Berat Asal (g)
            'End If
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
                    LM_GST_JUAL = Format(Frm84.TB11, "0.00")
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah Cukai GST (RM)
                End If
                
                rs!gst_include = Null '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang
                
            ElseIf Frm84.CB3 = 1 Then
            
                rs!gst_ari_nashi = "SR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                If Frm84.L8_Text <> vbNullString Then
                    rs!kadar_gst = Frm84.L8_Text 'Kadar Cukai GST (%)
                Else
                    rs!kadar_gst = "0" 'Jumlah Cukai GST (RM)
                End If
                If Frm84.TB11 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm84.TB11, "0.00") 'Jumlah Cukai GST (RM)
                    LM_GST_JUAL = Format(Frm84.TB11, "0.00")
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah Cukai GST (RM)
                End If

                rs!gst_include = Null '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang

            ElseIf Frm84.CB18 = 1 Then
            
                rs!gst_ari_nashi = "SR" '0 : Tiada GST : ZR(L) , 1 : Ada GST : SR
                If Frm84.L8_Text <> vbNullString Then
                    rs!kadar_gst = Frm84.L8_Text 'Kadar Cukai GST (%)
                Else
                    rs!kadar_gst = "0" 'Jumlah Cukai GST (RM)
                End If
                If Frm84.TB11 <> vbNullString Then
                    rs!jumlah_gst = Format(Frm84.TB11, "0.00") 'Jumlah Cukai GST (RM)
                    LM_GST_JUAL = Format(Frm84.TB11, "0.00")
                Else
                    rs!jumlah_gst = "0.00" 'Jumlah Cukai GST (RM)
                End If
                
                rs!gst_include = 1 '0 : GST Dibayar Oleh Pelanggan , 1 : GST Termasuk Dalam Harga Barang

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
                    LM_MODAL_DGN_GST = Format(Frm84.L34_Text, "0.00")
                Else
                    rs!modal = Null 'Harga Modal (RM)
                End If
                If Frm84.L42_Text <> vbNullString Then
                    rs!modal_tanpa_gst = Format(Frm84.L42_Text, "0.00") 'Harga Modal Tanpa GST (RM)
                    Frm84_LM_MODAL_TANPA_GST = Frm84.L42_Text
                Else
                    rs!modal_tanpa_gst = Null 'Harga Modal (RM)
                End If
                If Frm84.L44_Text <> vbNullString Then
                    If IsNumeric(Frm84.L44_Text) Then Frm84_LM_HARGA_JUALAN_TANPA_GST = Frm84.L44_Text
                End If
                
                If Frm84.CB3 = 1 And Frm84.CB18 = 0 Then
                    Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
                    Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
                    
                    Frm84_LM_HARGA_JUALAN_DENGAN_GST = Format(Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC, "0.00")
                    'Field ini adalah lebih kurang kepada @harga_dengan_gst
                    'Tetapi bagi gst yang dikenakan hanya pada UPAH akan membuatkan field ini hanya memaparkan harga dengan gst hanya pada upah.
                    'Oleh itu field ini dibuat khas bagi mencampurkan gst pada upah dan harga barang.
                Else
                    Frm84_LM_HARGA_JUALAN_DENGAN_GST = Format(Frm84.TB10, "0.00") 'Harga Jualan (RM)
                End If
                
                If IsNumeric(Frm84.L34_Text) Then Frm84_LM_MODAL_DENGAN_GST = Frm84.L34_Text 'Harga Modal
                
                rs!jualan_per_gram_dengan_gst = Null
                rs!untung = Format(Frm84_LM_HARGA_JUALAN_DENGAN_GST - Frm84_LM_MODAL_DENGAN_GST, "0.00") 'Jumlah Keuntungan (pengiraan termasuk GST)
                rs!untung2 = Format(Frm84_LM_HARGA_JUALAN_TANPA_GST - Frm84_LM_MODAL_TANPA_GST, "0.00") 'Jumlah Keuntungan (pengiraan TIDAK termasuk GST)
                
                rs!harga_per_gram_supplier = Null 'Harga per gram (harga semasa) dari supplier (modal)
                rs!upah_modal = Null 'Upah modal
                rs!harga_per_gram_tanpa_gst = Null 'Harga modal per gram tanpa GST (RM)
                
            Else
            
                rs!Type = 0 '0 : BK , 1 : Barang Permata
                
                If Frm84.L34_Text <> vbNullString Then
                    rs!harga_per_gram_modal = Format(Frm84.L34_Text, "0.00") 'Harga Per Gram Bagi Modal (RM/g)
                    If IsNumeric(Frm84.L34_Text) Then
                        Frm84_LM_HARGA_SEMASA_MODAL = Frm84.L34_Text 'Harga Per Gram Bagi Modal (RM/g)
                        
                        rs!modal = Format(Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_SEMASA_MODAL, "0.00") 'Harga Modal (RM)
                        LM_MODAL_DGN_GST = Format(Frm84_LM_BERAT_JUAL * Frm84_LM_HARGA_SEMASA_MODAL, "0.00")
                    End If
                Else
                    rs!modal = Null 'Harga Modal (RM)
                    rs!harga_per_gram_modal = Null 'Harga Per Gram Bagi Modal (RM/g)
                End If
                If Frm84.L42_Text <> vbNullString Then
                    rs!harga_per_gram_tanpa_gst = Format(Frm84.L42_Text, "0.00") 'Harga modal per gram tanpa GST (RM)
                    LM_MODAL_TANPA_GST_GRAM = Frm84.L42_Text
                    LM_MODAL_TANPA_GST = Format(Frm84_LM_BERAT_JUAL * LM_MODAL_TANPA_GST_GRAM, "0.00")
                Else
                    rs!harga_per_gram_tanpa_gst = Null 'Harga modal per gram tanpa GST (RM)
                End If
                
                If IsNumeric(Frm84.TB14) And IsNumeric(Frm84.L34_Text) And IsNumeric(Frm84.TB10) And IsNumeric(Frm84.TB11) Then
                    
                    Frm84_LM_HARGA_MODAL = Frm84.L34_Text 'Harga Modal
                    Frm84_LM_HARGA_JUAL = Frm84.TB14 'Harga Jualan
                    Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
                    Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
                    
                    If Frm84.L34_Text <> vbNullString Then
                        If IsNumeric(Frm84.L34_Text) Then Frm84_LM_MODAL_DENGAN_GST = Frm84.L34_Text 'Harga Modal
                    End If
                    
                    If Frm84.L42_Text <> vbNullString Then
                        If IsNumeric(Frm84.L42_Text) Then Frm84_LM_MODAL_TANPA_GST = Frm84.L42_Text
                    End If
                        
                    If Frm84.CB12 = 0 Then
                        
                        If Frm84.TB14 <> vbNullString Then
                            If IsNumeric(Frm84.TB14) Then Frm84_LM_HARGA_JUALAN_DENGAN_GST = Frm84.TB14 'Harga Jualan
                        End If
                        
                        If Frm84.L44_Text <> vbNullString Then
                            If IsNumeric(Frm84.L44_Text) Then Frm84_LM_HARGA_JUALAN_TANPA_GST = Frm84.L44_Text
                        End If
                        
                        rs!jualan_per_gram_dengan_gst = Format(Frm84_LM_HARGA_JUALAN_DENGAN_GST / Frm84_LM_BERAT_JUAL, "0.00")
                        rs!untung = Format((Frm84_LM_HARGA_JUALAN_DENGAN_GST) - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_DENGAN_GST), "0.00") 'Jumlah Keuntungan (pengiraan termasuk GST)
                        rs!untung2 = Format(Frm84_LM_HARGA_JUALAN_TANPA_GST - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_TANPA_GST), "0.00") 'Jumlah Keuntungan (pengiraan TIDAK termasuk GST)

                    ElseIf Frm84.CB12 = 1 Then
                        
                        If Frm84.CB2 = 1 Then
                        
                            If Frm84.TB10 <> vbNullString Then
                                If IsNumeric(Frm84.TB10) Then Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
                            End If
                             
                            If Frm84.TB11 <> vbNullString Then
                                Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
                            End If
                            
                            rs!jualan_per_gram_dengan_gst = Format((Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC) / Frm84_LM_BERAT_JUAL, "0.00")
                            rs!untung = Format((Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC) - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_DENGAN_GST), "0.00") 'Jumlah Keuntungan (pengiraan termasuk GST)
                            rs!untung2 = Format(Frm84_LM_HARGA_JUALAN_CALC - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_TANPA_GST), "0.00") 'Jumlah Keuntungan (pengiraan TIDAK termasuk GST)
                            
                        ElseIf Frm84.CB3 = 1 Then
                            
                            If Frm84.TB10 <> vbNullString Then
                                If IsNumeric(Frm84.TB10) Then Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
                            End If
                             
                            If Frm84.TB11 <> vbNullString Then
                                Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
                            End If
                            
                            rs!jualan_per_gram_dengan_gst = Format((Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC) / Frm84_LM_BERAT_JUAL, "0.00")
                            rs!untung = Format((Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC) - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_DENGAN_GST), "0.00") 'Jumlah Keuntungan (pengiraan termasuk GST)
                            rs!untung2 = Format(Frm84_LM_HARGA_JUALAN_CALC - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_TANPA_GST), "0.00") 'Jumlah Keuntungan (pengiraan TIDAK termasuk GST)

                        ElseIf Frm84.CB18 = 1 Then
                        
                            If Frm84.TB10 <> vbNullString Then
                                If IsNumeric(Frm84.TB10) Then Frm84_LM_HARGA_JUALAN_CALC = Frm84.TB10 'Harga jualan asal bagi pengiraan @harga_jualan_dengan_gst
                            End If
                             
                            If Frm84.TB11 <> vbNullString Then
                                Frm84_LM_GST_CALC = Frm84.TB11 'Jumlah GST bagi pengiraan @harga_jualan_dengan_gst
                            End If
                            
                            rs!jualan_per_gram_dengan_gst = Format((Frm84_LM_HARGA_JUALAN_CALC) / Frm84_LM_BERAT_JUAL, "0.00")
                            rs!untung = Format((Frm84_LM_HARGA_JUALAN_CALC) - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_DENGAN_GST), "0.00")  'Jumlah Keuntungan (pengiraan termasuk GST)
                            rs!untung2 = Format((Frm84_LM_HARGA_JUALAN_CALC - Frm84_LM_GST_CALC) - (Frm84_LM_BERAT_JUAL * Frm84_LM_MODAL_TANPA_GST), "0.00") 'Jumlah Keuntungan (pengiraan TIDAK termasuk GST)

                        End If
                        
                    End If
                    
                End If
                
            End If
            
            If Format(Frm84.TB3, "0.00") = Format(Frm84.TB4, "0.00") Then
                rs!potong_flag = 0 '0 : Tiada Potong , 1 : Ada Potong
                rs!status_jualan = 0 '0 : Tiada Potong , 1 : Ada Potong
            Else
                rs!potong_flag = 1 '0 : Tiada Potong , 1 : Ada Potong
                rs!status_jualan = 1 '0 : Tiada Potong , 1 : Ada Potong
            End If
            
            'rs!dulang = Frm84_LM_DULANG 'Dulang
            
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
                LM_HARGA_JUALAN_DGN_GST = Format(Frm84_LM_HARGA_JUALAN_CALC + Frm84_LM_GST_CALC, "0.00")
                'Field ini adalah lebih kurang kepada @harga_dengan_gst
                'Tetapi bagi gst yang dikenakan hanya pada UPAH akan membuatkan field ini hanya memaparkan harga dengan gst hanya pada upah.
                'Oleh itu field ini dibuat khas bagi mencampurkan gst pada upah dam harga barang.
            Else
                rs!harga_jualan_dengan_gst = Format(Frm84.TB10, "0.00") 'Harga Jualan (RM)
                LM_HARGA_JUALAN_DGN_GST = Format(Frm84.TB10, "0.00")
            End If
            'yabai
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
            If Frm84.L70_Text = 0 Then
                
                rs!flag_upah = 0
                
                If Frm84.TB22 <> vbNullString Then
                
                    rs!upah_per_gram = Format(Frm84.TB22, "0.00")
                
                Else
                
                    rs!upah_per_gram = "0.00"
                    
                End If
            
            ElseIf Frm84.L70_Text = "1" Then
                
                rs!flag_upah = 1
                rs!upah_per_gram = Null
            
            End If
            
            rs!harga_jual_excl_gst = Format(LM_HARGA_JUALAN_DGN_GST - LM_GST_JUAL, "0.00")
            rs!harga_modal_gst = Format(LM_MODAL_DGN_GST - LM_MODAL_TANPA_GST, "0.00")
            rs!harga_modal_incl_gst = Format(LM_MODAL_DGN_GST, "0.00")
            rs!harga_modal_excl_gst = Format(LM_MODAL_TANPA_GST, "0.00")
            
            rs!untung = Format(LM_HARGA_JUALAN_DGN_GST - LM_GST_JUAL - LM_MODAL_TANPA_GST, "0.00")
            rs!untung2 = Format(LM_HARGA_JUALAN_DGN_GST - LM_MODAL_DGN_GST, "0.00")
            rs!baru_or_ti = LM_BARU_TI
            
            rs.Update
            Frm84_LM_DATA_SAVE = 1
        End If
        
        rs.Close
        Set rs = Nothing
'### Masukkan Data Ke Dalam Temp Table ### - End
        
        If Frm84_LM_DATA_SAVE = 1 Then
            'Call Frm84_Reset
            Call Frm84_Reset_Edit
            
            GM_NEXT_PREV = 0 '0 : Next , 1 : Previous
                    
            Call Frm84_Senarai_Jualan_Header
            Call Frm84_Senarai_Jualan
            
            MsgBox "Barang ini telah berjaya dimasukkan ke dalam senarai jualan.", vbInformation, "Info"
            
            Frm84.TB1.SetFocus
        End If
    End If
End If
End Sub
Sub tesuto2()
'On Error Resume Next
Dim Err(30)
Dim Frm84_LM_HARGA As Double
Dim Frm84_LM_JUMLAH_BAYARAN As Double
Dim Frm84_LM_JUMLAH_SIMPANAN As Double
Dim Frm84_LM_GUNA_SIMPAN As Double
Dim Frm84_LM_BERAT_ASAL As Double 'Berat Asal (g)
Dim Frm84_LM_BERAT_JUALAN As Double 'Berat Jualan (g)
Dim Frm84_JUMLAH_SIMPAN_ASAL As Double 'Jumlah Simpanan Asal (RM)
Dim Frm84_JUMLAH_GUNA_SIMPANAN As Double 'Jumlah Penggunaan Duit Simpanan (RM)
Dim LM_SOLD As String
Dim Frm84_LM_QTY As Double

Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset

Dim Frm84_LM_MATA_ASAL As Double
Dim Frm84_LM_MATA_TEBUS As Double
Dim Frm84_LM_MATA_DAPAT As Double
Dim frm130_LM_JUMLAH_SIMPANAN As Double
Dim frm130_LM_GUNA_SIMPAN As Double

frm130_LM_JUMLAH_SIMPANAN = 0
frm130_LM_GUNA_SIMPAN = 0

DATA_SAVE = 0
Frm84_LM_NO_REG_CUST = 0
Frm84_LM_Flag_SIMPANAN = 0 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
Frm84_LM_Flag_TRADE_IN = 0 '0 : Tiada Urusan Trade In , 1 : Ada Urusan Trade In
Frm84_LM_No_RESIT_JUALAN = vbNullString
x = 0
Frm84_LM_QTY = 0
Frm84_LM_HARGA = 0
Frm84_LM_JUMLAH_BAYARAN = 0
Frm84_LM_JUMLAH_SIMPANAN = 0
Frm84_LM_GUNA_SIMPAN = 0
Frm84_JUMLAH_SIMPAN_ASAL = 0
Frm84_JUMLAH_GUNA_SIMPANAN = 0

Frm84_LM_MATA_ASAL = 0
Frm84_LM_MATA_TEBUS = 0
Frm84_LM_MATA_DAPAT = 0

G_JENIS_URUSAN = 0

If Frm84.L4_Text = 0 Then
    x = x + 1
    Err(x) = "Tiada senarai jualan."
End If
'If Frm84.L4_Text <> vbNullString And IsNumeric(Frm84.L4_Text) Then
'    Frm84_LM_QTY = Frm84.L4_Text
    
'    If Frm84_LM_QTY > 15 Then
'        x = x + 1
'        Err(x) = "Bilangan barang yang dibenarkan untuk dijual di dalam satu invoice adalah 15."
'    End If
'End If
If Frm84.CB7 = 1 Then
    If Frm84.L29_Text = vbNullString Then
        x = x + 1
        Err(x) = "Tiada maklumat agen dropship."
    End If
End If
If Frm84.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih nama pekerja yang membuat jualan."
End If
If Frm84.Pic6.Visible = True Then
    x = x + 1
    Err(x) = "Anda berada di dalam menu pilihan kategori pembeli. Sila tutup menu ini untuk teruskan jualan."
End If
If Frm84.TB19 = vbNullString Or (Frm84.TB19 <> vbNullString And Not IsNumeric(Frm84.TB19)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Diskaun]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm84.TB20 = vbNullString Or (Frm84.TB20 <> vbNullString And Not IsNumeric(Frm84.TB20)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Adjustment]. Hanya NOMBOR dibenarkan dalam ruangan Ini."
End If


If frm130.TB27 = vbNullString Or (frm130.TB27 <> vbNullString And Not IsNumeric(frm130.TB27)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara TUNAI. Sila masukkan 0 jika tiada bayaran secara tunai."
End If
If frm130.TB28 = vbNullString Or (frm130.TB28 <> vbNullString And Not IsNumeric(frm130.TB28)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara ONLINE TRANSFER. Sila masukkan 0 jka tiada bayaran secara online transfer."
End If
If frm130.TB29 = vbNullString Or (frm130.TB29 <> vbNullString And Not IsNumeric(frm130.TB29)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara KAD KREDIT. Sila masukkan 0 jika tiada bayaran secara kad kredit."
End If
If frm130.TB21 = vbNullString Or (frm130.TB21 <> vbNullString And Not IsNumeric(frm130.TB21)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara Duit Simpanan Di Kedai. Sila masukkan 0 jika tiada bayaran secara simpanan di kedai."
End If

'Error bagi penggunaan kad kredit - Start
If frm130.TB29 <> "0.00" And IsNumeric(frm130.TB29) Then

    If frm130.CBB2 = vbNullString Then
        x = x + 1
        Err(x) = "Sila pilih jenis kad kredit/debit"
    End If
    If frm130.L31_Text = vbNullString Or (frm130.L31_Text <> vbNullString And Not IsNumeric(frm130.L31_Text)) Then
        x = x + 1
        Err(x) = "Tiada maklumat bagi caj perkhidmatan kad kredit/debit."
    End If
    If frm130.L32_Text = vbNullString Or (frm130.L32_Text <> vbNullString And Not IsNumeric(frm130.L32_Text)) Then
        x = x + 1
        Err(x) = "Tiada maklumat bagi jumlah caj perkhidmatan kad kredit/debit."
    End If
    If frm130.L81_Text = vbNullString Or (frm130.L81_Text <> vbNullString And Not IsNumeric(frm130.L81_Text)) Then
        x = x + 1
        Err(x) = "Tiada maklumat bagi jumlah cukai GST bagi caj perkhidmatan kad kredit/debit."
    End If
    If frm130.L81_Text = vbNullString Or (frm130.L81_Text <> vbNullString And Not IsNumeric(frm130.L81_Text)) Then
        x = x + 1
        Err(x) = "Tiada maklumat bagi jumlah potongan kad kredit/debit."
    End If
    
End If
'Error bagi penggunaan kad kredit - End

If Frm84.L25_Text = "Jumlah Bayaran" Then
    If (frm130.TB32 <> vbNullString And IsNumeric(frm130.TB32)) And (frm130.TB33 <> vbNullString And IsNumeric(frm130.TB33)) Then
        frm130_LM_JUMLAH_BAYARAN = frm130.TB32 'Jumlah Bayaran
        frm130_LM_HARGA = frm130.TB33 'Harga Keseluruhan
        
        If frm130_LM_JUMLAH_BAYARAN <> frm130_LM_HARGA Then
            x = x + 1
            Err(x) = "Jumlah bayaran tidak sama dengan jumlah harga barang."
        End If
    End If
End If

If (frm130.TB21 <> vbNullString And IsNumeric(frm130.TB21)) And (frm130.L26_Text <> vbNullString And IsNumeric(frm130.L26_Text)) Then

    frm130_LM_JUMLAH_SIMPANAN = frm130.L26_Text  'Jumlah Simpanan Yang Ada
    frm130_LM_GUNA_SIMPAN = frm130.TB21  'Jumlah Simpanan Yang Hendak Digunakan
    
    If frm130_LM_GUNA_SIMPAN > frm130_LM_JUMLAH_SIMPANAN Then
        x = x + 1
        Err(x) = "Jumlah simpanan yang hendak digunakan melebihi simpanan terkumpul yang ada."
    End If
    
End If
If Frm84.TB42 = vbNullString Or (Frm84.TB42 <> vbNullString And Not IsNumeric(Frm84.TB42)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan [Caj Pos Laju]. Sila masukkan 0 jika tiada bayaran ini."
End If
If Frm84.CB4 = 0 And Frm84.CB5 = 0 And Frm84.CB6 = 0 And Frm84.CB9 = 0 And Frm84.CB10 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan kategori pembeli"
End If
If G_TI_MODE = 3 Then
    If Frm84.TB49 = vbNullString Or (Frm84.TB49 <> vbNullString And Not IsNumeric(Frm84.TB49)) Then
        x = x + 1
        Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan [Berat Trade In]."
    End If
    If Frm84.TB50 = vbNullString Or (Frm84.TB50 <> vbNullString And Not IsNumeric(Frm84.TB50)) Then
        x = x + 1
        Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan [Harga Semasa Trade In]."
    End If
    If Frm84.TB51 = vbNullString Or (Frm84.TB51 <> vbNullString And Not IsNumeric(Frm84.TB51)) Then
        x = x + 1
        Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan [Harga Semasa Buyback]."
    End If
    If Frm84.TB52 = vbNullString Or (Frm84.TB52 <> vbNullString And Not IsNumeric(Frm84.TB52)) Then
        x = x + 1
        Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan [Caj Pertukaran]."
    End If
End If
If Frm84.L56_Text = 2 Then
    If Frm83.L9_Text <> vbNullString Then
        If Not IsNumeric(Frm83.L9_Text) Then
            x = x + 1
            Err(x) = "Technical error. Sila keluar dari menu ini dan cuba sekali lagi."
        End If
    End If
    If Frm83.CB8 = 1 Then
        If Frm83.L12_Text <> vbNullString Then
            If Not IsNumeric(Frm83.L12_Text) Then
                x = x + 1
                Err(x) = "Technical error. Sila keluar dari menu ini dan cuba sekali lagi."
            End If
        End If
    End If

    'If Frm83.L9_Text = vbNullString Then
    '    X = X + 1
    '    Err(X) = "Tiada maklumat no. rujukan bagi trade in. Sila keluar dari menu jualan ini dan cuba sekali lagi."
    'End If
    If Frm84.L57_Text = vbNullString Then
        x = x + 1
        Err(x) = "Tiada maklumat no. voucher bagi trade in. Sila keluar dari menu jualan ini dan cuba sekali lagi."
    End If
End If
If Frm84.CB19 = 1 Then
    If Frm84.TB41 = vbNullString Then
        x = x + 1
        Err(x) = "Sila masukkan approval code bagi bayaran secara EPP."
    End If
End If

If Frm84.L56_Text = 1 Then 'Mode belian dengan trade in : 0 : Tiada , 1 : Belian dengan trade in (sudah ada voucher) , 2 : Trade in

    If Frm84.L57_Text = vbNullString Then
        x = x + 1
        Err(x) = "Tiada data bagi barang trade in (Sila masukkan maklumat No Voucher bagi trade in)."
    End If

End If
If Frm84.L56_Text = 2 Then 'Mode belian dengan trade in : 0 : Tiada , 1 : Belian dengan trade in (sudah ada voucher) , 2 : Trade in

    If Frm83.L10_Text = 0 Then
        x = x + 1
        Err(x) = "Tiada data bagi barang trade in."
    End If

End If

If Frm84.L76_Text <> 0 Then
    Frm84_LM_JUMLAH_POINT = Frm84.L76_Text
End If

'### Point
If Frm84.L79_Text = 1 Then
    
    Frm84_LM_MATA_ASAL = 0
    Frm84_LM_MATA_TEBUS = 0
    
    If Frm84.TB35 = vbNullString Or (Frm84.TB35 <> vbNullString And Not IsNumeric(Frm84.TB35)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Kadar perolehan mata ganjaran]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm84.TB36 = vbNullString Or (Frm84.TB36 <> vbNullString And Not IsNumeric(Frm84.TB36)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Jumlah tebusan mata ganjaran]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm84.TB37 = vbNullString Or (Frm84.TB37 <> vbNullString And Not IsNumeric(Frm84.TB37)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Kadar tebusan mata ganjaran]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm84.TB36 <> vbNullString And IsNumeric(Frm84.TB36) Then
        Frm84_LM_MATA_TEBUS = Frm84.TB36
    End If
    If Frm84.L77_Text <> vbNullString And IsNumeric(Frm84.L77_Text) Then
        Frm84_LM_MATA_ASAL = Frm84.L77_Text
    End If
    If Frm84_LM_MATA_TEBUS > Frm84_LM_MATA_ASAL Then
        x = x + 1
        Err(x) = "Mata yang ingin ditebus adalah melebihi dari mata terkumpul."
    End If

End If

'### Periksa Samada Maklumat Pembeli Diisi Dalam Kedua-dua Ruangan Berdaftar Dan Tidak Berdaftar ### - Start
If x = 0 Then
    If Frm84.L27_Text <> vbNullString And Frm84.L28_Text <> vbNullString Then
    
        MsgBox "Data bagi pembeli telah diisi bagi kedua-dua ruangan pembeli berdaftar dan tidak berdaftar." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila padam salah satu yang tidak berkenaan.", vbExclamation, "Info"
                    
        Exit Sub
          
    End If
End If
'### Periksa Samada Maklumat Pembeli Diisi Dalam Kedua-dua Ruangan Berdaftar Dan Tidak Berdaftar ### - End

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else

    If Frm84.L27_Text <> vbNullString And Frm84.L28_Text = vbNullString Then
        Note = "Adakah anda yakin untuk teruskan urusan jualan ini ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Data jualan akan disimpan ke dalam sistem." & vbCrLf & _
                vbNullString & vbCrLf & _
                "*** Oleh ini pembeli ini tidak berdaftar dengan sistem, oleh itu data dan rekod pembelian pembeli ini tidak akan disimpan di dalam sistem ***"
    End If
    If Frm84.L27_Text = vbNullString And Frm84.L28_Text <> vbNullString Then
        Note = "Adakah anda yakin untuk teruskan urusan jualan ini ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Data jualan akan disimpan ke dalam sistem."
    End If
    If Frm84.L27_Text = vbNullString And Frm84.L28_Text = vbNullString Then
    
        Note = "TIADA maklumat bagi pembeli telah diisi." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Maklumat pembeli tidak akan dicetak di dalam invoice pembeli." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Adakah anda yakin untuk teruskan urusan jualan ini ?"
                
        Frm84_LM_NO_REG_CUST = 1
        
    End If
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then
    
        '### Pop up confirmation bagi jualan bagi invoice tidak rasmi
        If Frm84.CB13 = 1 Then
        
            Note = "Jualan ini dibuat dengan pilihan INVOICE TIDAK RASMI." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Anda TIDAK BOLEH mengubah jenis invoice jika jualan ini telah dibuat." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Teruskan?"
                
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbNo Then
            
                Exit Sub
            
            End If
            
        End If
        
        LM_RATE_KUPON_2 = vbNullString
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
        
'$$$$ Periksa status terkini setiap item yang hendak dijual $$$$ - Start
        LM_TRANS_VOID = 0

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select data_database.no_siri_produk from " & G_JUALAN_TEMP & ",data_database where " & G_JUALAN_TEMP & ".no_siri_produk = data_database.no_siri_produk AND (data_database.statusitem <> 10 AND data_database.statusitem <> 12 AND data_database.statusitem <> 22 AND data_database.statusitem <> 28)", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
            If LM_TRANS_VOID = 1 Then
                LM_SOLD = LM_SOLD & " , " & rs!no_siri_Produk ' & vbCrLf
            End If
            If LM_TRANS_VOID = 0 Then
                LM_SOLD = LM_SOLD & rs!no_siri_Produk  ' & vbCrLf
                
                LM_TRANS_VOID = 1
            End If
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
        
        If LM_TRANS_VOID = 1 Then
            MsgBox " Barang-barang berikut tidak dibenarkan untuk dijual kerana kemungkinan barang tersebut telah terjual." & vbCrLf & _
                    "Senarai barang tersebut adalah seperti di bawah : " & vbCrLf & _
                    LM_SOLD & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila periksa data anda.", vbExclamation, "Periksa Data"
                    
            Exit Sub
        End If
'$$$$ Periksa status terkini setiap item yang hendak dijual $$$$ - End

        LM_NOW = Now
        
        '$$$ No. staff $$$ - Start
        If InStr(1, Frm84.CBB1, "  |  ") <> 0 Then
            Frm84_LM_EMP_NO = Split(Frm84.CBB1, "  |  ")(1)
            Frm84_LM_EMP_NAMA = Split(Frm84.CBB1, "  |  ")(0)
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
        
'---------------------------------------No. Invoice
        LM_NOW = Now

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        If Frm84.CB13 = 0 Then strsql = "insert into 4_senarai_invoice_rasmi(tarikh,terminal,write_timestamp,Status,nama_staff)" & _
                                "select '" & Frm84.DTPicker1 & "','" & G_TERMINAL & "','" & LM_NOW & "',1,'" & MDI_frm1.L3_Text & "'"
        If Frm84.CB13 = 1 Then strsql = "insert into 5_senarai_invoice_tidak_rasmi(tarikh,terminal,write_timestamp,Status,nama_staff)" & _
                                "select '" & Frm84.DTPicker1 & "','" & G_TERMINAL & "','" & LM_NOW & "',1,'" & MDI_frm1.L3_Text & "'"
                
        Set rs = cn2.Execute(strsql)
        Set rs = Nothing
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main3 Else Exit Sub
        If Frm84.CB13 = 0 Then rs.Open "select * from 4_senarai_invoice_rasmi where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & Frm84.DTPicker1 & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic
        If Frm84.CB13 = 1 Then rs.Open "select * from 5_senarai_invoice_tidak_rasmi where nama_staff='" & MDI_frm1.L3_Text & "' AND terminal='" & G_TERMINAL & "' AND write_timestamp='" & LM_NOW & "' AND tarikh='" & Frm84.DTPicker1 & "' AND status = 1 order by ID DESC", cn2, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!ID) Then
                If Frm84.CB13 = 0 Then rs!no_invoice = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(rs!ID, "000000")
                If Frm84.CB13 = 1 Then rs!no_invoice = "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(rs!ID, "000000")
                Frm84_LM_No_RESIT_JUALAN = rs!ID 'No. Rujukan Belian
                rs.Update
            End If
            rs.Update
        Else
        
            MsgBox "Berlaku ralat semasa data cuba disimpan. Sila keluar dari menu ini dan cuba lagi.", vbCritical, "Error"
            
            rs.Close
            Set rs = Nothing
            
            Exit Sub
            
        End If
        
        rs.Close
        Set rs = Nothing

'### Periksa NO INVOICE sebelum simpan data ke dalam database ### - Start
        
Re_Gen_No_Rujukan:
'###Masukkan Jualan Ke Dalam Table Akaun Jualan### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        If Frm84.CB13 = 0 Then rs.Open "select * from 22_jualan where no_resit='" & "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84_LM_No_RESIT_JUALAN, "000000") & "' AND bil_rasmi = 1 AND status = 1", cn, adOpenKeyset, adLockOptimistic
        If Frm84.CB13 = 1 Then rs.Open "select * from 22_jualan where no_resit='" & "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84_LM_No_RESIT_JUALAN, "000000") & "' AND bil_rasmi = 0 AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If rs.EOF Then
            rs.AddNew
            If Frm84.CB13 = 0 Then
                If Frm84.L3_Text <> vbNullString Then
                    rs!no_resit = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84_LM_No_RESIT_JUALAN, "000000") 'No. invoice rasmi
                Else
                    rs!no_resit = Null 'No. invoice rasmi
                End If
                rs!bil_rasmi = 1 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
            Else
                If Frm84.L66_Text <> vbNullString Then
                    rs!no_resit = "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84_LM_No_RESIT_JUALAN, "000000") 'No. invoice tidak rasmi
                Else
                    rs!no_resit = Null 'No. invoice tidak rasmi
                End If
                rs!bil_rasmi = 0 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
            End If
            
            rs!tarikh = Frm84.DTPicker1 'Tarikh Jualan
            
            rs!tunai = Format(0, "0.00")
            rs!bank_in = Format(0, "0.00") 'Cara Bayaran : Bank In
            rs!kad_kredit = Format(0, "0.00") 'Cara Bayaran : Kad Kredit
            rs!duit_simpanan_kedai = Format(0, "0.00") 'Cara Bayaran : Simpanan Duit Di Kedai
            rs!cek = Format(0, "0.00") 'Cara Bayaran : Cek
            
            If Frm84.L25_Text = "Jumlah Bayaran" Then
                If frm130.TB27 <> vbNullString Then
                    rs!tunai = Format(frm130.TB27, "0.00") 'Cara Bayaran : Tunai
                Else
                    rs!tunai = Null 'Cara Bayaran : Tunai
                End If
                If frm130.TB28 <> vbNullString Then
                    rs!bank_in = Format(frm130.TB28, "0.00") 'Cara Bayaran : Bank In
                Else
                    rs!bank_in = Null 'Cara Bayaran : Bank In
                End If
                If frm130.TB29 <> vbNullString Then
                    rs!kad_kredit = Format(frm130.TB29, "0.00") 'Cara Bayaran : Kad Kredit
                    If Format(frm130.TB29, "0.00") <> "0.00" Then
                        
                        If frm130.CBB2 <> vbNullString Then
                            rs!jenis_kad = frm130.CBB2
                        Else
                            rs!jenis_kad = Null
                        End If
                        If frm130.L31_Text <> vbNullString Then
                            rs!cas_Kad_Kredit = Format(frm130.L31_Text, "0.00") 'Cara Bayaran : Cas Kad Kredit (%)
                        Else
                            rs!cas_Kad_Kredit = "0.00" 'Cara Bayaran : Cas Kad Kredit (%)
                        End If
                        If frm130.L32_Text <> vbNullString Then
                            rs!jumlah_cas_kad_kredit = Format(frm130.L32_Text, "0.00") 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
                        Else
                            rs!jumlah_cas_kad_kredit = "0.00" 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
                        End If
                        If frm130.L81_Text <> vbNullString Then
                            rs!gst_kad_kredit = Format(frm130.L81_Text, "0.00") 'Cara Bayaran : Jumlah GST kad kredit (RM)
                        Else
                            rs!gst_kad_kredit = "0.00" 'Cara Bayaran : Jumlah GST kad kredit (RM)
                        End If
                        If frm130.L82_Text <> vbNullString Then
                            rs!jumlah_potongan_kad_kredit = Format(frm130.L82_Text, "0.00") 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
                        Else
                            rs!jumlah_potongan_kad_kredit = "0.00" 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
                        End If
                        If frm130.L8_Text <> vbNullString Then
                            rs!kadar_gst_kad_kredit = Format(frm130.L8_Text, "0.00") 'Cara Bayaran : Kadar GST bagi kad kredit
                        Else
                            rs!kadar_gst_kad_kredit = "0.00" 'Cara Bayaran : Kadar GST bagi kad kredit
                        End If
                        If Frm84.CB19 = 1 Then
                            rs!epp = 1 '0 : Bayaran selain dari EPP , 1 : Bayaran secara EPP
                            rs!approval_code_epp = UCase(Frm84.TB41) 'Approval Code (EPP)
                        Else
                            rs!epp = 0 '0 : Bayaran selain dari EPP , 1 : Bayaran secara EPP
                            rs!approval_code_epp = Null 'Approval Code (EPP)
                        End If
                    Else
                        rs!jenis_kad = Null
                        rs!cas_Kad_Kredit = Null 'Cara Bayaran : Cas Kad Kredit (%)
                        rs!jumlah_cas_kad_kredit = Null 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
                        rs!jumlah_potongan_kad_kredit = Null 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
                        rs!kadar_gst_kad_kredit = Null 'Cara Bayaran : Kadar GST bagi kad kredit
                        rs!gst_kad_kredit = Null 'Cara Bayaran : Jumlah GST kad kredit (RM)
                        
                        rs!epp = 0 '0 : Bayaran selain dari EPP , 1 : Bayaran secara EPP
                        rs!approval_code_epp = Null 'Approval Code (EPP)
                    End If
                Else
                    rs!jenis_kad = Null
                    rs!cas_Kad_Kredit = Null 'Cara Bayaran : Cas Kad Kredit (%)
                    rs!jumlah_cas_kad_kredit = Null 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
                    rs!jumlah_potongan_kad_kredit = Null 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
                    rs!kadar_gst_kad_kredit = Null 'Cara Bayaran : Kadar GST bagi kad kredit
                    rs!gst_kad_kredit = Null 'Cara Bayaran : Jumlah GST kad kredit (RM)
                    rs!epp = 0 '0 : Bayaran selain dari EPP , 1 : Bayaran secara EPP
                    rs!approval_code_epp = Null 'Approval Code (EPP)
                    'rs!kad_kredit = Null 'Cara Bayaran : Kad Kredit
                End If

                If frm130.TB21 <> vbNullString Then
                    If Format(frm130.TB21, "0.00") <> "0.00" Then
                        Frm84_LM_Flag_SIMPANAN = 1 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
                    End If
                    rs!duit_simpanan_kedai = Format(frm130.TB21, "0.00") 'Cara Bayaran : Simpanan Duit Di Kedai
                Else
                    rs!duit_simpanan_kedai = Null 'Cara Bayaran : Simpanan Duit Di Kedai
                End If
                If frm130.TB32 <> vbNullString Then
                    rs!jumlah_bayaran = Format(frm130.TB32, "0.00") 'Cara Bayaran : Jumlah Bayaran
                Else
                    rs!jumlah_bayaran = Null 'Cara Bayaran : Jumlah Bayaran
                End If
            Else
                rs!tunai = "0.00" 'Cara Bayaran : Tunai
                rs!bank_in = "0.00" 'Cara Bayaran : Bank In
                rs!kad_kredit = "0.00" 'Cara Bayaran : Kad Kredit
                If frm130.L31_Text <> vbNullString Then
                    rs!cas_Kad_Kredit = Format(frm130.L31_Text, "0.00") 'Cara Bayaran : Cas Kad Kredit (%)
                Else
                    rs!cas_Kad_Kredit = 0 'Cara Bayaran : Cas Kad Kredit (%)
                End If
                rs!epp = 0 '0 : Bayaran selain dari EPP , 1 : Bayaran secara EPP
                rs!approval_code_epp = Null 'Approval Code (EPP)
                rs!jumlah_cas_kad_kredit = "0.00" 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
                rs!jumlah_potongan_kad_kredit = "0.00" 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
                rs!duit_simpanan_kedai = "0.00" 'Cara Bayaran : Simpanan Duit Di Kedai
                rs!kad_debit = "0.00" 'Cara Bayaran : Kad Debit
                If frm130.L32_Text <> vbNullString Then
                    rs!cas_kad_debit = frm130.L32_Text 'Cara Bayaran : Jumlah Cas Kad Debit (%)
                Else
                    rs!cas_kad_debit = 0 'Cara Bayaran : Jumlah Cas Kad Debit (%)
                End If
                rs!jumlah_cas_kad_debit = "0.00" 'Cara Bayaran : Jumlah Cas Kad Debit (RM)
                rs!jumlah_potongan_kad_debit = "0.00" 'Cara Bayaran : Jumah Potongan Kad Debit (RM)
                rs!jumlah_bayaran = "0.00" 'Cara Bayaran : Jumlah Bayaran
                rs!jenis_kad = Null
                rs!cas_Kad_Kredit = Null 'Cara Bayaran : Cas Kad Kredit (%)
                rs!jumlah_cas_kad_kredit = Null 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
                rs!jumlah_potongan_kad_kredit = Null 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
                rs!kadar_gst_kad_kredit = Null 'Cara Bayaran : Kadar GST bagi kad kredit
                rs!gst_kad_kredit = Null 'Cara Bayaran : Jumlah GST kad kredit (RM)
                rs!epp = 0 '0 : Bayaran selain dari EPP , 1 : Bayaran secara EPP
                rs!approval_code_epp = Null 'Approval Code (EPP)
                'rs!kad_kredit = Null 'Cara Bayaran : Kad Kredit
            End If
            
            If Frm84.L17_Text <> vbNullString Then
                rs!harga_barang = Format(Frm84.L17_Text, "0.00") 'Jumlah Harga Barang Tanpa GST (RM)
            Else
                rs!harga_barang = Null 'Jumlah Harga Barang Tanpa GST (RM)
            End If
            If Frm84.L8_Text <> vbNullString Then
                rs!kadar_gst = Format(Frm84.L8_Text, "0.00")
            Else
                rs!kadar_gst = Null
            End If
            If Frm84.L18_Text <> vbNullString Then
                rs!jumlah_cukai_gst = Format(Frm84.L18_Text, "0.00") 'Jumlah Cukai GST (ZR + SR)
            Else
                rs!jumlah_cukai_gst = Null 'Jumlah Cukai GST (ZR + SR)
            End If
            If Frm84.L19_Text <> vbNullString Then
                rs!harga_barang_dengan_gst = Format(Frm84.L19_Text, "0.00") 'Jumlah Harga Barang Dengan GST (RM)
            Else
                rs!harga_barang_dengan_gst = Null 'Jumlah Harga Barang Dengan GST (RM)
            End If
            If Frm84.TB19 <> vbNullString Then
                rs!diskaun = Format(Frm84.TB19, "0.00") 'Jumlah Diskaun (%)
            Else
                rs!diskaun = Null 'Jumlah Diskaun (%)
            End If
            If Frm84.L20_Text <> vbNullString Then
                rs!harga_lepas_diskaun = Format(Frm84.L20_Text, "0.00") 'Harga Selepas Diskaun (RM)
            Else
                rs!harga_lepas_diskaun = Null 'Harga Selepas Diskaun (RM)
            End If
            If Frm84.TB20 <> vbNullString Then
                rs!adjustment = Format(Frm84.TB20, "0.00") 'Adjustment (RM)
            Else
                rs!adjustment = Null 'Adjustment (RM)
            End If
            If Frm84.L21_Text <> vbNullString Then
                rs!harga_jualan = Format(Frm84.L21_Text, "0.00") 'Jumlah Harga Jualan (RM)
            Else
                rs!harga_jualan = Null 'Jumlah Harga Jualan (RM)
            End If
            If Frm84.L38_Text <> vbNullString Then
                rs!loss_trade_in = Format(Frm84.L38_Text, "0.00") 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (%)
            Else
                rs!loss_trade_in = Null 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (%)
            End If
            If Frm84.L37_Text <> vbNullString Then
                rs!loss_trade_in_rm = Format(Frm84.L37_Text, "0.00") 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (RM)
            Else
                rs!loss_trade_in_rm = Null 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (RM)
            End If
            If Frm84.L24_Text = "Jumlah Bayaran" Then
                rs!flag_bayaran = 0 '0 : Pembeli Bayar , 1 : Kedai Bayar
            Else
                rs!flag_bayaran = 1 '0 : Pembeli Bayar , 1 : Kedai Bayar
            End If
            If Frm84.L23_Text <> vbNullString Then
                rs!jumlah_perlu_bayar = Format(Frm84.L23_Text, "0.00") 'Jumlah Bayaran Yang Perlu Dibuat (RM)
            Else
                rs!jumlah_perlu_bayar = Null 'Jumlah Bayaran Yang Perlu Dibuat (RM)
            End If
            If Frm84.L14_Text <> vbNullString Then
                rs!kuantiti_barang = Frm84.L14_Text 'Kuantiti Barang Yang Dijual
            Else
                rs!kuantiti_barang = Null 'Kuantiti Barang Yang Dijual
            End If
            If Frm84.L15_Text <> vbNullString Then
                rs!JUMLAH_BERAT = Frm84.L15_Text 'Jumlah Berat Barang Yang Dijual
            Else
                rs!JUMLAH_BERAT = Null 'Jumlah Berat Barang Yang Dijual
            End If
            If Frm84.L7_Text <> vbNullString Then
                rs!gst_zr_harga = Format(Frm84.L7_Text, "0.00") 'Harga Keseluruhan Bagi Barang ZR
            Else
                rs!gst_zr_harga = Null 'Harga Keseluruhan Bagi Barang ZR
            End If
            If Frm84.L9_Text <> vbNullString Then
                rs!gst_zr_cukai = Format(Frm84.L9_Text, "0.00") 'Jumlah Cukai Bagi ZR
            Else
                rs!gst_zr_cukai = Null 'Jumlah Cukai Bagi ZR
            End If
            If Frm84.L10_Text <> vbNullString Then
                rs!gst_sr_harga = Format(Frm84.L10_Text, "0.00") 'Harga Keseluruhan Bagi Barang SR
            Else
                rs!gst_sr_harga = Null 'Harga Keseluruhan Bagi Barang SR
            End If
            If Frm84.L11_Text <> vbNullString Then
                rs!gst_sr_cukai = Format(Frm84.L11_Text, "0.00") 'Jumlah Cukai Bagi SR
            Else
                rs!gst_sr_cukai = Null 'Jumlah Cukai Bagi SR
            End If

            If Frm84.TB42 <> vbNullString Then 'Jumlah caj pos laju (postage)
                rs!caj_pos = Format(Frm84.TB42, "0.00")
            Else
                rs!caj_pos = "0.00"
            End If
            If Frm84.TB45 <> vbNullString Then 'No. Tracking pos laju
                rs!no_tracking = UCase(Frm84.TB45)
            Else
                rs!no_tracking = Null
            End If
            
            rs!no_pekerja = Frm84_LM_EMP_NO 'No. Pekerja

            If Frm84.L28_Text <> vbNullString Then
                If Frm28.L5_Text <> vbNullString Then
                    rs!no_rujukan_pembeli = Frm28.L5_Text 'No. Rujukan Pembeli
                Else
                    rs!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
                End If
            Else
                rs!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
            End If
            If Frm84.CB7 = 1 Then
                'If Frm84.L29_Text <> vbNullString Then
                    If Frm27.L5_Text <> vbNullString Then
                        rs!no_rujukan_agen_dropship = Frm27.L5_Text 'No. Rujukan Agen Dropship
                    Else
                        rs!no_rujukan_agen_dropship = Null 'No. Rujukan Agen Dropship
                    End If
                'Else
                '    rs!no_rujukan_agen_dropship = Null 'No. Rujukan Agen Dropship
                'End If
            Else
                rs!no_rujukan_agen_dropship = Null 'No. Rujukan Agen Dropship
            End If
            If G_TI_MODE <> 3 Then
                If Frm84.L56_Text <> 0 Then
                    Frm84_LM_Flag_TRADE_IN = 1 '0 : Tiada Urusan Trade In , 1 : Ada Urusan Trade In
                    rs!flag_trade_in = 1 '0 : Tiada Urusan Trade in , 1 : Ada Urusan Trade In
                    
                    If Frm84.L56_Text = 1 Then
                        rs!jenis_trade_in = 1 '1 : Trade in (Voucher) , 2 : Belian dengan trade in
                        rs!no_resit_trade_in = Frm84.L57_Text 'Frm84.L57_Text 'No. Resit Trade In
                    ElseIf Frm84.L56_Text = 2 Then
                        rs!jenis_trade_in = 2 '1 : Trade in (Voucher) , 2 : Belian dengan trade in
                        rs!no_resit_trade_in = "TI" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84.L57_Text, "000000") 'Frm84.L57_Text 'No. Resit Trade In
                    Else
                        rs!no_resit_trade_in = Null 'No. Resit Trade In
                    End If
                    
                    'If Frm84.L57_Text <> vbNullString Then
                        'rs!no_resit_trade_in = "TI" & Format(Frm84.L57_Text, "000000") 'Frm84.L57_Text 'No. Resit Trade In
                        'rs!no_resit_trade_in = Frm84.L57_Text 'Frm84.L57_Text 'No. Resit Trade In
                    'Else
                    '    rs!no_resit_trade_in = Null 'No. Resit Trade In
                    'End If
                    If Frm84.L58_Text <> vbNullString Then
                        rs!jumlah_trade_in = Format(Frm84.L58_Text, "0.00") 'No. Resit Trade In
                    Else
                        rs!jumlah_trade_in = Null 'No. Resit Trade In
                    End If
                Else
                    rs!flag_trade_in = 0 '0 : Tiada Urusan Trade in , 1 : Ada Urusan Trade In
                    rs!no_resit_trade_in = Null 'No. Resit Trade In
                    rs!jumlah_trade_in = Null 'Jumlah Resit Trade In (RM)
                    rs!jenis_trade_in = Null '1 : Trade in (Voucher) , 2 : Belian dengan trade in
                End If
                rs!jumlah_caj_tukaran = Null
            Else
                rs!flag_trade_in = 1 '0 : Tiada Urusan Trade in , 1 : Ada Urusan Trade In
                rs!jenis_trade_in = 3 '1 : Trade in (Voucher) , 2 : Belian dengan trade in
                rs!jumlah_trade_in = G_TRADE_IN_TOTAL
                rs!jumlah_caj_tukaran = G_TRADE_IN_CAJ

                If Frm84.TB49 <> vbNullString Then
                    rs!berat_trade_in = Frm84.TB49
                Else
                    rs!berat_trade_in = 0
                End If
                If Frm84.TB50 <> vbNullString Then
                    rs!harga_semasa_trade_in = Frm84.TB50
                Else
                    rs!harga_semasa_trade_in = 0
                End If
                If Frm84.TB51 <> vbNullString Then
                    rs!harga_semasa_buyback = Frm84.TB51
                Else
                    rs!harga_semasa_buyback = 0
                End If
                If Frm84.TB52 <> vbNullString Then
                    rs!caj_pertukaran = Frm84.TB52
                Else
                    rs!caj_pertukaran = 0
                End If
            End If
            If Frm84.L46_Text <> vbNullString Then '0 : Unlimited , Selain 0 (Limited : Mengikut nombor yang dimasukkan)
                rs!invoice_type = Frm84.L46_Text
            Else
                rs!invoice_type = 0
            End If
                 
'1:  Pelanggan
'2:  Member
'3:  RAF
'4:  Pengedar
'5:  Normal Dealer
'6:  Master Dealer

'Zakaria&Sons
'1 : Pembeli biasa
'2 : Ahli biasa
'3 : Silver
'4 : Gold
'5 : Platinum

            If Frm84.CB4 = 1 Then
                rs!kategori_pembeli = 1
            ElseIf Frm84.CB5 = 1 Then
                rs!kategori_pembeli = 2
            ElseIf Frm84.CB6 = 1 Then
                rs!kategori_pembeli = 3
            ElseIf Frm84.CB9 = 1 Then
                rs!kategori_pembeli = 4
            ElseIf Frm84.CB10 = 1 Then
                rs!kategori_pembeli = 5
            'ElseIf Frm84.CB11 = 1 Then
            '    rs!kategori_pembeli = 6
            End If
            If Frm84.CB27 = 1 Then
                rs!jualan_online = 1
            Else
                rs!jualan_online = 0
            End If
            If Frm84.L79_Text = 0 Then
                rs!point_ari_nashi = 0
            ElseIf Frm84.L79_Text = 1 Then
                rs!point_ari_nashi = 1
            End If
            If Frm84.L73_Text <> vbNullString Then
                rs!redeem_point = Frm84.L73_Text
            Else
                rs!redeem_point = 0
            End If
            If Frm84.L76_Text <> vbNullString Then
                rs!jumlah_point = Frm84.L76_Text
            Else
                rs!jumlah_point = 0
            End If
            If Frm84.TB34 <> vbNullString Then
                rs!kupon_diskaun = Format(Frm84.TB34, "0.00")
            Else
                rs!kupon_diskaun = "0.00"
            End If
            If Frm84.TB35 <> vbNullString Then
                rs!kadar_peroleh_point = Frm84.TB35
            Else
                rs!kadar_peroleh_point = 0
            End If
            If Frm84.TB37 <> vbNullString Then
                rs!kadar_tebus_point = Frm84.TB37
            Else
                rs!kadar_tebus_point = 0
            End If
            rs!kadar_diskaun = Format(Frm84_LM_KUPON, "0.00") 'Kadar diskaun per gram
            rs!Status = 1
            rs!status_r = 0
            rs!terminal = G_TERMINAL
            rs!write_timestamp = LM_NOW
            rs!Menu = 0
            rs!nama_pekerja = Frm84_LM_EMP_NAMA
            rs!cawangan = G_CAWANGAN
            If Frm84.TB46 <> vbNullString Then
                rs!remarks = Frm84.TB46
            Else
                rs!remarks = Null
            End If
            
            DATA_SAVE = 1
            rs.Update
        Else
            Frm84_LM_No_RESIT_JUALAN = Frm84_LM_No_RESIT_JUALAN + 1
            If Frm84.CB13 = 0 Then Frm84.L3_Text = Frm84_LM_No_RESIT_JUALAN 'No. invoice rasmi
            If Frm84.CB13 = 1 Then Frm84.L66_Text = Frm84_LM_No_RESIT_JUALAN 'No. invoice tidak rasmi
            
            rs.Close
            Set rs = Nothing
            GoTo Re_Gen_No_Rujukan:
        End If
        
        rs.Close
        Set rs = Nothing
'###Masukkan Jualan Ke Dalam Table Akaun Jualan### - End

'###Update Data Simpanan Duit Pelanggan### - Start
        If Frm84_LM_Flag_SIMPANAN = 1 And Frm84.L24_Text = "Jumlah Bayaran" And Frm84.L28_Text <> vbNullString Then '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                
                G_ID = rs!ID
                Call recovery_senarai_pelanggan
                
                If Not IsNull(rs!baki_simpanan) Then Frm84_LM_JUMLAH_SIMPANAN = rs!baki_simpanan
                
                'Frm84_LM_JUMLAH_SIMPANAN = frm130.L26_Text  'Jumlah Simpanan Yang Ada
                Frm84_LM_GUNA_SIMPAN = frm130.TB21  'Jumlah Simpanan Yang Hendak Digunakan
                
                rs!baki_simpanan = Format(Frm84_LM_JUMLAH_SIMPANAN - Frm84_LM_GUNA_SIMPAN, "0.00") 'Baki Simpanan
                rs!write_timestamp2 = LM_NOW
                rs!no_staff = Frm84_LM_EMP_NO 'No. Pekerja
                rs!terminal = G_TERMINAL
                rs!jenis_urusan = G_JENIS_URUSAN
                rs!cawangan = G_CAWANGAN
                rs.Update
                
            End If
            
            rs.Close
            Set rs = Nothing
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 24_rekod_kewangan_pelanggan where id is null", cn, adOpenKeyset, adLockOptimistic
            
            If rs.EOF Then
                rs.AddNew
                rs!tarikh = Frm84.DTPicker1 'Tarikh
                rs!jenis = 1 '0 : Simpanan , 1 : Penggunaan Duit
                rs!no_rujukan_pelanggan = Frm28.L5_Text 'No. Rujukan Pelanggan
                If Frm84.CB13 = 0 Then rs!no_resit = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84_LM_No_RESIT_JUALAN, "000000") 'No. invoice rasmi
                If Frm84.CB13 = 1 Then rs!no_resit = "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84_LM_No_RESIT_JUALAN, "000000") 'No. invoice tidak rasmi
                rs!jumlah = Format(frm130.TB21, "0.00") 'Jumlah Simpanan Yang Digunakan (RM)
                rs!jenis_penggunaan = 0 '0 : Belian Barangan Kemas , 1 : Ansuran , 2 : Tempahan (Deposit) , 3 : Servis , 4 : Tempahan (Ambilan Barang)
                rs!no_rujukan_pekerja = Frm84_LM_EMP_NO 'No. Pekerja
                rs!terminal = G_TERMINAL
                rs!write_timestamp = LM_NOW
                rs!jenis_urusan = G_JENIS_URUSAN
                rs!cawangan = G_CAWANGAN
                rs!Status = 1
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
           
        End If
'###Update Data Simpanan Duit Pelanggan### - End

'### Maklumat Pembeli Yang Tidak Berdaftar Dengan Kedai ### - Start (08-07-2015)
        If Frm84.L27_Text <> vbNullString Then
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 44_senarai_pelanggan where id is null", cn, adOpenKeyset, adLockOptimistic
            
            If rs.EOF Then
                rs.AddNew
                rs!tarikh = Frm84.DTPicker1 'Tarikh
                If Frm84.CB13 = 0 Then rs!no_resit = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84_LM_No_RESIT_JUALAN, "000000") 'No. invoice rasmi
                If Frm84.CB13 = 1 Then rs!no_resit = "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84_LM_No_RESIT_JUALAN, "000000") 'No. invoice tidak rasmi
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
                rs!write_timestamp = LM_NOW
                rs!no_staff = Frm84_LM_EMP_NO 'No. Pekerja
                rs!terminal = G_TERMINAL
                rs!jenis_urusan = G_JENIS_URUSAN
                rs!cawangan = G_CAWANGAN
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
                
        End If
'### Maklumat Pembeli Yang Tidak Berdaftar Dengan Kedai ### - End

'### Maklumat agihan point ### - Start
        If Frm84.L79_Text = 1 Then
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            If Frm84.CB13 = 0 Then rs.Open "select * from 71_tebus_agih_point where no_invoice='" & "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84_LM_No_RESIT_JUALAN, "000000") & "'", cn, adOpenKeyset, adLockOptimistic
            If Frm84.CB13 = 1 Then rs.Open "select * from 71_tebus_agih_point where no_invoice='" & "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84_LM_No_RESIT_JUALAN, "000000") & "'", cn, adOpenKeyset, adLockOptimistic
            
            If rs.EOF Then
                rs.AddNew
                If Frm84.L3_Text <> vbNullString Then
                    If Frm84.CB13 = 0 Then rs!no_invoice = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84_LM_No_RESIT_JUALAN, "000000")  'No. Resit Jualan
                    If Frm84.CB13 = 1 Then rs!no_invoice = "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84_LM_No_RESIT_JUALAN, "000000")  'No. Resit Jualan
                Else
                    rs!no_invoice = Null 'No. Resit Jualan
                End If
                If Frm84.DTPicker1 <> vbNullString Then
                    rs!tarikh = Frm84.DTPicker1 'Tarikh Jualan
                Else
                    rs!tarikh = Null 'Tarikh Jualan
                End If
                If Frm28.L5_Text <> vbNullString Then 'No. Rujukan Pembeli
                    rs!no_ahli = Frm28.L5_Text
                Else
                    rs!no_ahli = Null
                End If
                If Frm84.L75_Text <> vbNullString Then 'Harga yang membolehkan untuk mendaparkan point
                    rs!harga_layak_bonus = Format(Frm84.L75_Text, "0.00")
                Else
                    rs!harga_layak_bonus = Null
                End If
                If Frm84.TB35 <> vbNullString Then 'Kadar perolehan point (eg. 0.5)
                    rs!kadar_peroleh_point = Frm84.TB35
                Else
                    rs!kadar_peroleh_point = Null
                End If
                If Frm84.L76_Text <> vbNullString Then 'Jumlah perolehan mata
                    rs!jumlah_peroleh_point = Frm84.L76_Text
                Else
                    rs!jumlah_peroleh_point = Null
                End If
                If Frm84.TB36 <> vbNullString Then 'Jumlah mata yang ditebus
                    rs!jumlah_tebus_point = Frm84.TB36
                Else
                    rs!jumlah_tebus_point = Null
                End If
                If Frm84.TB37 <> vbNullString Then 'Kadar tebusan mata
                    rs!kadar_tebus_point = Frm84.TB37
                Else
                    rs!kadar_tebus_point = Null
                End If
                If Frm84.L78_Text <> vbNullString Then 'Jumlah nilaian mata yang ditebus
                    rs!nilaian_tebus_point = Frm84.L78_Text
                Else
                    rs!nilaian_tebus_point = Null
                End If
                If Frm84.CB13 = 0 Then
                    rs!bil_rasmi = 1 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
                Else
                    rs!bil_rasmi = 0 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
                End If
                rs!no_pekerja = Frm84_LM_EMP_NO 'No. Pekerja
                rs!write_timestamp = LM_NOW
                rs!Status = 1 '0 : Tidak aktif , 1 : Aktif
                rs!Type = 1
                rs!terminal = G_TERMINAL
                rs!jenis_urusan = G_JENIS_URUSAN
                rs!cawangan = G_CAWANGAN
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
            
                G_ID = rs!ID
                Call recovery_senarai_pelanggan

                If Frm84.L76_Text <> vbNullString Then 'Jumlah perolehan mata
                    If IsNumeric(Frm84.L76_Text) Then Frm84_LM_MATA_DAPAT = Frm84.L76_Text
                End If
                If Frm84.TB36 <> vbNullString Then 'Jumlah mata yang ditebus
                    If IsNumeric(Frm84.TB36) Then Frm84_LM_MATA_TEBUS = Frm84.TB36
                End If
                If Frm84.L77_Text <> vbNullString Then 'Jumlah mata asal
                    If IsNumeric(Frm84.L77_Text) Then Frm84_LM_MATA_ASAL = Frm84.L77_Text
                End If
                rs!baki_point = Frm84_LM_MATA_ASAL - Frm84_LM_MATA_TEBUS + Frm84_LM_MATA_DAPAT
                rs!write_timestamp2 = LM_NOW
                rs!no_staff = Frm84_LM_EMP_NO 'No. Pekerja
                rs!terminal = G_TERMINAL
                rs!jenis_urusan = G_JENIS_URUSAN
                rs!cawangan = G_CAWANGAN
                rs.Update
                
            End If
            
            rs.Close
            Set rs = Nothing

        End If
'### Maklumat agihan point ### - End

'### Update Maklumat Trade In ### - Start
        If Frm84_LM_Flag_TRADE_IN = 1 Then '0 : Tiada Urusan Trade In , 1 : Ada Urusan Trade In
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & Frm84.L16_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then

                G_ID = rs!ID
                Call recovery_16_gold_bar_belian
                
                rs!trade_in_status = 1
                rs!no_staff = Frm84_LM_EMP_NO 'No. Pekerja
                rs!terminal = G_TERMINAL
                rs!write_timestamp2 = LM_NOW
                rs!jenis_urusan = G_JENIS_URUSAN
                rs!remarks = "Ubah status flag trade in bagi jualan kepada pelanggan"
                rs.Update
                
            End If
            
            rs.Close
            Set rs = Nothing
            
        End If
'### Update Maklumat Trade In ### - End

'### Masukkan Data Jualan Ke Dalam Table Jualan ### - Start
        If Frm84.CB13 = 0 Then LM_NO_INVOICE = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84_LM_No_RESIT_JUALAN, "000000") 'No. invoice rasmi
        If Frm84.CB13 = 1 Then LM_NO_INVOICE = "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84_LM_No_RESIT_JUALAN, "000000") 'No. invoice tidak rasmi
        
        If Frm84.CB4 = 1 Then
            LM_KATEGORI = 1
        ElseIf Frm84.CB5 = 1 Then
            LM_KATEGORI = 2
        ElseIf Frm84.CB6 = 1 Then
            LM_KATEGORI = 4
        ElseIf Frm84.CB9 = 1 Then
            LM_KATEGORI = 3
        ElseIf Frm84.CB10 = 1 Then
            LM_KATEGORI = 5
        'ElseIf Frm84.CB11 = 1 Then
        '    rs1!kategori_pembeli = 6
        End If
        If Frm84.L28_Text <> vbNullString Then
            If Frm28.L5_Text <> vbNullString Then
                LM_NO_PEMBELI = Frm28.L5_Text 'No. Rujukan Pembeli
            Else
                LM_NO_PEMBELI = Null
            End If
        Else
            LM_NO_PEMBELI = Null
        End If
        If Frm84.CB7 = 1 Then
            If Frm27.L5_Text <> vbNullString Then
                LM_NO_DROPSHIP = Frm27.L5_Text 'No. Rujukan Agen Dropship
            Else
                LM_NO_DROPSHIP = Null
            End If
        Else
            LM_NO_DROPSHIP = Null
        End If
        If Frm84.CB27 = 1 Then
            LM_ONLINE = 1
        Else
            LM_ONLINE = 0
        End If
        If Frm84.CB13 = 0 Then
            LM_BIL_RASMI = 1
        Else
            LM_BIL_RASMI = 0
        End If

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
        strsql = "insert into 23_senarai_jualan(no_resit,cawangan,nama_pekerja,jenis_jualan,tarikh,no_pekerja,no_rujukan_pembeli,no_rujukan_agen_dropship,kategori_pembeli,status_rekod,write_timestamp,jualan_online,bil_rasmi,status_r,no_siri_produk,flag_barang,nama_purity,kategori_produk,baru_or_ti,purity,berat_asal,berat_jualan,harga_semasa,upah,harga_asal,diskaun,harga_lepas_diskaun,adjustment,harga_jualan,gst_ari_nashi,kadar_gst,jumlah_gst,harga_dengan_gst," _
                    & "harga_jual_excl_gst,harga_modal_gst,harga_modal_incl_gst,harga_modal_excl_gst,dropship,komisyen_per_gram,jumlah_komisyen,status,type,potong_flag,modal_tanpa_gst,harga_per_gram_tanpa_gst,jualan_per_gram_dengan_gst,harga_per_gram_modal,modal,untung,harga_per_gram_supplier,untung2,dulang,gst_include,harga_tanpa_gst,harga_koperasi,kadar_penurunan_upah,kadar_penurunan_bp,harga_semasa_staff,harga_bp_asal,upah_asal,harga_staff,komisyen_staff,pemalar_tukaran_999,berat_999,upah_modal,gst_barang_atau_upah,harga_jualan_dengan_gst,kadar_komisyen_upah,komisyen_upah,jualan_per_gram,modal_per_gram,flag_upah,upah_per_gram,jenis_urusan,terminal)" & _
                    "select '" & LM_NO_INVOICE & "','" & G_CAWANGAN & "','" & Frm84_LM_EMP_NAMA & "',0,'" & Frm84.DTPicker1 & "','" & Frm84_LM_EMP_NO & "','" & LM_NO_PEMBELI & "','" _
                    & LM_NO_DROPSHIP & "','" & LM_KATEGORI & "',1,'" & LM_NOW & "','" & LM_ONLINE & "','" & LM_BIL_RASMI & "',0,no_siri_produk,flag_barang,nama_purity,kategori_produk," _
                    & "baru_or_ti,purity,berat_asal,berat_jualan,harga_semasa,upah,harga_asal,diskaun,harga_lepas_diskaun,adjustment,harga_jualan," _
                    & "gst_ari_nashi,kadar_gst,jumlah_gst,harga_dengan_gst,harga_jual_excl_gst,harga_modal_gst,harga_modal_incl_gst,harga_modal_excl_gst,dropship,komisyen_per_gram,jumlah_komisyen,status_jualan,type," _
                    & "potong_flag,modal_tanpa_gst,harga_per_gram_tanpa_gst,jualan_per_gram_dengan_gst,harga_per_gram_modal,modal,untung,harga_per_gram_supplier,untung2,dulang,gst_include,harga_tanpa_gst," _
                    & "harga_koperasi,kadar_penurunan_upah,kadar_penurunan_bp,harga_semasa_staff,harga_bp_asal,upah_asal,harga_staff," _
                    & "komisyen_staff,pemalar_tukaran_999,berat_999,upah_modal,gst_barang_atau_upah,harga_jualan_dengan_gst," _
                    & "kadar_komisyen_upah,komisyen_upah,jualan_per_gram,modal_per_gram,flag_upah,upah_per_gram,'" & G_JENIS_URUSAN & "','" & G_TERMINAL & "'" _
                    & "from " & G_JUALAN_TEMP & " WHERE status='" & 1 & "'"
            
        Set rs = cn.Execute(strsql)
        Set rs = Nothing

        If G_TI_MODE = 3 Then
        'masukkan senarai trade in
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            
            If G_TI_MEMORY(0, 0) = 0 Then strsql = "insert into 93_trade_in_susut_niai(no_invoice,tarikh,berat,harga_semasa,harga,status,write_timestamp,jenis,terminal,nama_pekerja) values ('" & LM_NO_INVOICE & "','" & Frm84.DTPicker1 & "','" & G_TI_MEMORY(1, 1) & "','" & G_TI_MEMORY(1, 2) & "','" & G_TI_MEMORY(1, 3) & "',1,'" & LM_NOW & "',0,'" & G_TERMINAL & "','" & Frm84_LM_EMP_NAMA & "') ,('" & LM_NO_INVOICE & "','" & Frm84.DTPicker1 & "','" & G_TI_MEMORY(2, 1) & "','" & G_TI_MEMORY(2, 2) & "','" & G_TI_MEMORY(2, 3) & "',1,'" & LM_NOW & "',1,'" & G_TERMINAL & "','" & Frm84_LM_EMP_NAMA & "'),('" & LM_NO_INVOICE & "','" & Frm84.DTPicker1 & "','" & G_TI_MEMORY(3, 1) & "','" & G_TI_MEMORY(3, 2) & "','" & G_TI_MEMORY(3, 3) & "',1,'" & LM_NOW & "',2,'" & G_TERMINAL & "','" & Frm84_LM_EMP_NAMA & "')"
            If G_TI_MEMORY(0, 0) = 1 Then strsql = "insert into 93_trade_in_susut_niai(no_invoice,tarikh,berat,harga_semasa,harga,status,write_timestamp,jenis,terminal,nama_pekerja) values ('" & LM_NO_INVOICE & "','" & Frm84.DTPicker1 & "','" & G_TI_MEMORY(1, 1) & "','" & G_TI_MEMORY(1, 2) & "','" & G_TI_MEMORY(1, 3) & "',1,'" & LM_NOW & "',0,'" & G_TERMINAL & "','" & Frm84_LM_EMP_NAMA & "') ,('" & LM_NO_INVOICE & "','" & Frm84.DTPicker1 & "','" & G_TI_MEMORY(2, 1) & "','" & G_TI_MEMORY(2, 2) & "','" & G_TI_MEMORY(2, 3) & "',1,'" & LM_NOW & "',2,'" & G_TERMINAL & "','" & Frm84_LM_EMP_NAMA & "')"

            Set rs = cn.Execute(strsql)
            Set rs = Nothing
            
        End If
'### Update status & info item yang terjual ### - Start

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_JUALAN_TEMP & " where status='" & "1" & "'", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
        
            Set rs2 = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs2.Open "select * from Data_Database where no_siri_produk='" & rs!no_siri_Produk & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs2.EOF Then
                
                G_ID = rs2!ID
                Call recovery_data_database
                
                If rs!Type = 0 Then
                    Frm84_LM_BERAT_ASAL = rs2!beza_berat 'Berat Asal (g)
                    Frm84_LM_BERAT_JUALAN = rs!berat_jualan 'Berat Jualan (g)
                    
                    If Frm84_LM_BERAT_JUALAN = Frm84_LM_BERAT_ASAL Then
                        rs2!beza_berat = "0.00" 'Baki Berat
                        rs2!susut_berat = "0.00" 'Susut berat
                        rs2!StatusItem = 11
                        rs2!tarikh_jualan1 = Null
                        rs2!nama_pekerja_potong = Null
                    Else
                        rs2!beza_berat = Format(Frm84_LM_BERAT_ASAL - Frm84_LM_BERAT_JUALAN, "0.00") 'Baki Berat
                        rs2!susut_berat = "0.00" 'Susut berat
                        rs2!StatusItem = 12
                        rs2!tarikh_jualan1 = Frm84.DTPicker1
                        rs2!nama_pekerja_potong = Frm84_LM_EMP_NAMA
                    End If
                Else
                    rs2!StatusItem = 11
                End If
                
                rs2!write_timestamp2 = LM_NOW
                rs2!no_pekerja = Frm84_LM_EMP_NO
                rs2!terminal = G_TERMINAL
                rs2!Menu = 0
                'rs2!cawangan = G_CAWANGAN

                rs2.Update
            End If
            
            rs2.Close
            Set rs2 = Nothing

            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
'### Update status & info item yang terjual ### - End
        
        If Frm84.L56_Text = 2 Then
            Call Frm84_penerimaan_barang_trade_in
        End If

'### Masukkan Data Jualan Ke Dalam Table Jualan ### - End

        If DATA_SAVE = 1 Then
'###Update No. Resit### - Start
            G_No_RESIT_JUALAN = vbNullString
            
            If Frm84.CB13 = 0 Then
                G_No_RESIT_JUALAN = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84_LM_No_RESIT_JUALAN, "000000")
            ElseIf Frm84.CB13 = 1 Then
                G_No_RESIT_JUALAN = "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84_LM_No_RESIT_JUALAN, "000000")
            End If
            
'#### Update Log Aktiviti Sistem #### - Start
            If Frm84.CB13 = 0 Then
                LogAct_Memory = "[" & MDI_frm1.L3_Text & "] Jualan barang kemas. No. Invoice [" & "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84_LM_No_RESIT_JUALAN, "000000") & "]."
            ElseIf Frm84.CB13 = 1 Then
                LogAct_Memory = "[" & MDI_frm1.L3_Text & "] Jualan barang kemas. No. Invoice [" & "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84_LM_No_RESIT_JUALAN, "000000") & "]."
            End If
            
            LogDate_Memory = LM_NOW
            Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End
            
            If G_SPKE_ME_MAIL = "YES" Then
            
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 88_sales_notfication where id is null", cn, adOpenKeyset, adLockOptimistic
                
                If rs.EOF Then
                    rs.AddNew
                    If Frm84.CB13 = 0 Then rs!no_invoice_asal = "BK" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84_LM_No_RESIT_JUALAN, "000000") 'No. invoice rasmi
                    If Frm84.CB13 = 1 Then rs!no_invoice_asal = "BKX" & G_TAHUN & "-" & G_KOD_KEDAI & "-" & Format(Frm84_LM_No_RESIT_JUALAN, "000000") 'No. invoice tidak rasmi
                    rs!jenis = 0
                    rs!jenis_report = 0 '0 : Jualan , 1 : Trade In
                    rs!write_timestamp = LM_NOW
                    rs!terminal = G_TERMINAL
                    rs!Status = 0
                    rs.Update
                End If
                
                rs.Close
                Set rs = Nothing

                Shell "cmd.exe /c " & G_SPKE_NE_PATH
                
            End If
'###Update No. Resit### - End
            
            Call Frm84_Load_Form
            Unload Frm26
            Unload Frm27
            Unload Frm28
            Unload Frm83
            
            G_PREVIEW = 1
            
            Note = "Data telah berjaya disimpan." & vbCrLf & _
                    "Adakah anda ingin cetak invoice jualan ?"
            
            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbYes Then
                
                G_KEDAI = G_CAWANGAN
                
                If Frm84.L46_Text = 0 Then
                    If G_INVOICE_TYPE = 0 Then '0 : Invoice Dari Sistem , 2 : Invoice Pre-printed
                        Call Frm84_Resit_Jualan
                    ElseIf G_INVOICE_TYPE = 1 Then '0 : Invoice Dari Sistem , 2 : Invoice Pre-printed
                        Call cetak_invoice
                    End If
                Else
                    Call Frm84_cetak_invoice_rms
                End If
            End If
        End If
    End If
End If
End Sub
Sub tesuto3()
'On Error Resume Next
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset

Dim Err(30)
Dim Frm84_LM_HARGA As Double
Dim Frm84_LM_JUMLAH_BAYARAN As Double
Dim Frm84_LM_JUMLAH_SIMPANAN As Double
Dim Frm84_LM_GUNA_SIMPAN As Double
Dim Frm84_LM_BERAT_ASAL As Double 'Berat Asal (g)
Dim Frm84_LM_BERAT_JUALAN As Double 'Berat Jualan (g)
Dim Frm84_JUMLAH_SIMPAN_ASAL As Double 'Jumlah Simpanan Asal (RM)
Dim Frm84_JUMLAH_GUNA_SIMPANAN As Double 'Jumlah Penggunaan Duit Simpanan (RM)
Dim Frm84_BERAT_JUALAN_BARU As Double
Dim Frm84_LM_BERAT_JUALAN_ASAL As Double
Dim Frm84_LM_BERAT_RETURN As Double
Dim Frm84_LM_REFUND_ASAL As Double 'Refund : Jumlah Simpanan Asal
Dim Frm84_LM_REFUND_GUNA As Double 'Refund : Jumlah Simpan Yang Telah Digunakan Sebelum Ini
Dim Frm84_LM_BEZA_BERAT As Double
Dim Frm84_LM_BAKI_BERAT As Double
Dim Frm84_LM_MATA_ASAL As Double
Dim Frm84_LM_MATA_TEBUS As Double
Dim Frm84_LM_MATA_DAPAT As Double
Dim Frm84_SUSUT_BERAT As Double
Dim Frm84_LM_BERAT_ASAL_COMP As Double
Dim Frm84_LM_BERAT_SELEPAS_COMP As Double
Dim Frm84_LM_QTY As Double
Dim frm130_LM_HARGA As Double

Dim frm130_LM_JUMLAH_SIMPANAN As Double
Dim frm130_LM_GUNA_SIMPAN As Double

frm130_LM_JUMLAH_SIMPANAN = 0
frm130_LM_GUNA_SIMPAN = 0

frm130_LM_HARGA = 0
Frm84_LM_BERAT_ASAL_COMP = 0
Frm84_LM_BERAT_SELEPAS_COMP = 0
DATA_SAVE = 0
Frm84_LM_Flag_SIMPANAN = 0 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
Frm84_LM_FLAG_PENGGUNAAN_DUIT_ASAL = 0 '0 : Tiada Penggunaan Duit Simpanan , 1 : Ada Pengunaan Duit Simpanan
Frm84_LM_No_RESIT_JUALAN = vbNullString
x = 0
Frm84_LM_HARGA = 0
Frm84_LM_JUMLAH_BAYARAN = 0
Frm84_LM_JUMLAH_SIMPANAN = 0
Frm84_LM_GUNA_SIMPAN = 0
Frm84_JUMLAH_SIMPAN_ASAL = 0
Frm84_JUMLAH_GUNA_SIMPANAN = 0
Frm84_LM_FLAG_TI_ASAL = vbNullString
Frm84_LM_No_PELANGGAN = vbNullString
Frm84_LM_No_PELANGGAN_MATA = vbNullString
Frm84_LM_Flag_TRADE_IN = 0 '0 : Tiada Urusan Trade In , 1 : Ada Urusan Trade In
Frm84_LM_BERAT_JUALAN_ASAL = 0
Frm84_LM_BERAT_RETURN = 0
Frm84_LM_JENIS_TRADE_IN = 0 '1 : Trade in (Voucher) , 2 : Belian dengan trade in
Frm84_LM_BAKI_BERAT = 0

Frm84_LM_QTY = 0
Frm84_LM_MATA_ASAL = 0
Frm84_LM_MATA_TEBUS = 0
Frm84_LM_MATA_DAPAT = 0
Frm84_LM_FLAG_MATA_ASAL = 0 '0 : Tiada perolehan mata , 1 : Ada perolehan mata (Pada awal pembelian asal)
Frm84_SUSUT_BERAT = 0

Frm84_LM_REFUND_ASAL = 0 'Refund : Jumlah Simpanan Asal
Frm84_LM_REFUND_GUNA = 0 'Refund : Jumlah Simpan Yang Telah Digunakan Sebelum Ini

If Frm84.L4_Text = 0 Then
    x = x + 1
    Err(x) = "Tiada senarai jualan."
End If
'If Frm84.L4_Text <> vbNullString And IsNumeric(Frm84.L4_Text) Then
'    Frm84_LM_QTY = Frm84.L4_Text
    
'    If Frm84_LM_QTY > 15 Then
'        x = x + 1
'        Err(x) = "Bilangan barang yang dibenarkan untuk dijual di dalam satu invoice adalah 15."
'    End If
'End If
If Frm84.CB7 = 1 Then
    If Frm84.L29_Text = vbNullString Then
        x = x + 1
        Err(x) = "Tiada maklumat agen dropship."
    End If
End If
If Frm84.L56_Text = 2 Then
    If Frm83.L9_Text = vbNullString Then
        x = x + 1
        Err(x) = "Tiada maklumat no. rujukan bagi trade in. Sila keluar dari menu jualan ini dan cuba sekali lagi."
    End If
    If Frm84.L57_Text = vbNullString Then
        x = x + 1
        Err(x) = "Tiada maklumat no. voucher bagi trade in. Sila keluar dari menu jualan ini dan cuba sekali lagi."
    End If
End If
If Frm84.CBB1 = vbNullString Then
    x = x + 1
    Err(x) = "Sila pilih nama pekerja yang membuat jualan."
End If
If Frm84.Pic6.Visible = True Then
    x = x + 1
    Err(x) = "Anda berada di dalam menu pilihan kategori pembeli. Sila tutup menu ini untuk teruskan jualan."
End If
If Frm84.TB19 = vbNullString Or (Frm84.TB19 <> vbNullString And Not IsNumeric(Frm84.TB19)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Diskaun]. Hanya NOMBOR dibenarkan dalam ruangan ini."
End If
If Frm84.TB20 = vbNullString Or (Frm84.TB20 <> vbNullString And Not IsNumeric(Frm84.TB20)) Then
    x = x + 1
    Err(x) = "Sila masukkan [Adjustment]. Hanya NOMBOR dibenarkan dalam ruangan Ini."
End If
If Frm84.CB4 = 0 And Frm84.CB5 = 0 And Frm84.CB6 = 0 And Frm84.CB9 = 0 And Frm84.CB10 = 0 Then
    x = x + 1
    Err(x) = "Sila buat pilihan kategori pembeli"
End If
If Frm84.CB19 = 1 Then
    If Frm84.TB41 = vbNullString Then
        x = x + 1
        Err(x) = "Sila masukkan approval code bagi bayaran secara EPP."
    End If
End If

If frm130.TB27 = vbNullString Or (frm130.TB27 <> vbNullString And Not IsNumeric(frm130.TB27)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara TUNAI. Sila masukkan 0 jika tiada bayaran secara tunai."
End If
If frm130.TB28 = vbNullString Or (frm130.TB28 <> vbNullString And Not IsNumeric(frm130.TB28)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara ONLINE TRANSFER. Sila masukkan 0 jka tiada bayaran secara online transfer."
End If
If frm130.TB29 = vbNullString Or (frm130.TB29 <> vbNullString And Not IsNumeric(frm130.TB29)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara KAD KREDIT. Sila masukkan 0 jika tiada bayaran secara kad kredit."
End If
If frm130.TB21 = vbNullString Or (frm130.TB21 <> vbNullString And Not IsNumeric(frm130.TB21)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan bayaran secara Duit Simpanan Di Kedai. Sila masukkan 0 jika tiada bayaran secara simpanan di kedai."
End If

'Error bagi penggunaan kad kredit - Start
If frm130.TB29 <> "0.00" And IsNumeric(frm130.TB29) Then

    If frm130.CBB2 = vbNullString Then
        x = x + 1
        Err(x) = "Sila pilih jenis kad kredit/debit"
    End If
    If frm130.L31_Text = vbNullString Or (frm130.L31_Text <> vbNullString And Not IsNumeric(frm130.L31_Text)) Then
        x = x + 1
        Err(x) = "Tiada maklumat bagi caj perkhidmatan kad kredit/debit."
    End If
    If frm130.L32_Text = vbNullString Or (frm130.L32_Text <> vbNullString And Not IsNumeric(frm130.L32_Text)) Then
        x = x + 1
        Err(x) = "Tiada maklumat bagi jumlah caj perkhidmatan kad kredit/debit."
    End If
    If frm130.L81_Text = vbNullString Or (frm130.L81_Text <> vbNullString And Not IsNumeric(frm130.L81_Text)) Then
        x = x + 1
        Err(x) = "Tiada maklumat bagi jumlah cukai GST bagi caj perkhidmatan kad kredit/debit."
    End If
    If frm130.L81_Text = vbNullString Or (frm130.L81_Text <> vbNullString And Not IsNumeric(frm130.L81_Text)) Then
        x = x + 1
        Err(x) = "Tiada maklumat bagi jumlah potongan kad kredit/debit."
    End If
    
End If
'Error bagi penggunaan kad kredit - End

If Frm84.L25_Text = "Jumlah Bayaran" Then
    If (frm130.TB32 <> vbNullString And IsNumeric(frm130.TB32)) And (frm130.TB33 <> vbNullString And IsNumeric(frm130.TB33)) Then
        frm130_LM_JUMLAH_BAYARAN = frm130.TB32 'Jumlah Bayaran
        frm130_LM_HARGA = frm130.TB33 'Harga Keseluruhan
        
        If frm130_LM_JUMLAH_BAYARAN <> frm130_LM_HARGA Then
            x = x + 1
            Err(x) = "Jumlah bayaran tidak sama dengan jumlah harga barang."
        End If
    End If
End If
If (frm130.TB21 <> vbNullString And IsNumeric(frm130.TB21)) And (frm130.L26_Text <> vbNullString And IsNumeric(frm130.L26_Text)) Then
    frm130_LM_JUMLAH_SIMPANAN = frm130.L26_Text  'Jumlah Simpanan Yang Ada
    frm130_LM_GUNA_SIMPAN = frm130.TB21  'Jumlah Simpanan Yang Hendak Digunakan
    
    If frm130_LM_GUNA_SIMPAN > frm130_LM_JUMLAH_SIMPANAN Then
        x = x + 1
        Err(x) = "Jumlah simpanan yang hendak digunakan melebihi simpanan terkumpul yang ada."
    End If
End If

If Frm84.TB42 = vbNullString Or (Frm84.TB42 <> vbNullString And Not IsNumeric(Frm84.TB42)) Then
    x = x + 1
    Err(x) = "Hanya NOMBOR dibenarkan dalam ruangan [Caj Pos Laju]. Sila masukkan 0 jika tiada bayaran ini."
End If

If Frm84.L3_Text <> vbNullString Then
    'If Not IsNumeric(Frm84.L3_Text) Then
    '    X = X + 1
    '    Err(X) = "Technical error. Sila keluar dari menu ini dan cuba sekali lagi."
    'End If
End If
If Frm84.L56_Text = 1 Then 'Mode belian dengan trade in : 0 : Tiada , 1 : Belian dengan trade in (sudah ada voucher) , 2 : Trade in

    If Frm84.L57_Text = vbNullString Then
        x = x + 1
        Err(x) = "Tiada data bagi barang trade in (Sila masukkan maklumat No Voucher bagi trade in)."
    End If

End If
If Frm84.L56_Text = 2 Then 'Mode belian dengan trade in : 0 : Tiada , 1 : Belian dengan trade in (sudah ada voucher) , 2 : Trade in
    'If Frm83.L9_Text <> vbNullString Then
    '    If Not IsNumeric(Frm83.L9_Text) Then
    '        X = X + 1
    '        Err(X) = "Technical error. Sila keluar dari menu ini dan cuba sekali lagi."
    '    End If
    'End If
    'If Frm83.CB8 = 1 Then
    '    If Frm83.L12_Text <> vbNullString Then
    '        If Not IsNumeric(Frm83.L12_Text) Then
    '            X = X + 1
    '            Err(X) = "Technical error. Sila keluar dari menu ini dan cuba sekali lagi."
    '        End If
    '    End If
    'End If
    If Frm83.L10_Text = 0 Then
        x = x + 1
        Err(x) = "Tiada data bagi barang trade in."
    End If
    If Frm84.L57_Text = vbNullString Then
        x = x + 1
        Err(x) = "Tiada maklumat no. voucher bagi trade in. Sila keluar dari menu jualan ini dan cuba sekali lagi."
    End If

End If

If Frm84.L76_Text <> 0 Then
    Frm84_LM_JUMLAH_POINT = Frm84.L76_Text
End If

'### Point
If Frm84.L79_Text = 1 Then
    
    Frm84_LM_MATA_ASAL = 0
    Frm84_LM_MATA_TEBUS = 0
    
    If Frm84.TB35 = vbNullString Or (Frm84.TB35 <> vbNullString And Not IsNumeric(Frm84.TB35)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Kadar perolehan mata ganjaran]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm84.TB36 = vbNullString Or (Frm84.TB36 <> vbNullString And Not IsNumeric(Frm84.TB36)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Jumlah tebusan mata ganjaran]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm84.TB37 = vbNullString Or (Frm84.TB37 <> vbNullString And Not IsNumeric(Frm84.TB37)) Then
        x = x + 1
        Err(x) = "Sila masukkan [Kadar tebusan mata ganjaran]. Hanya NOMBOR dibenarkan dalam ruangan ini."
    End If
    If Frm84.TB36 <> vbNullString And IsNumeric(Frm84.TB36) Then
        Frm84_LM_MATA_TEBUS = Frm84.TB36
    End If
    If Frm84.L77_Text <> vbNullString And IsNumeric(Frm84.L77_Text) Then
        Frm84_LM_MATA_ASAL = Frm84.L77_Text
    End If
    If Frm84_LM_MATA_TEBUS > Frm84_LM_MATA_ASAL Then
        x = x + 1
        Err(x) = "Mata yang ingin ditebus adalah melebihi dari mata terkumpul."
    End If

End If

'### Periksa Samada Maklumat Pembeli Diisi Dalam Kedua-dua Ruangan Berdaftar Dan Tidak Berdaftar ### - Start
If x = 0 Then
    If Frm84.L27_Text <> vbNullString And Frm84.L28_Text <> vbNullString Then
    
        MsgBox "Data bagi pembeli telah diisi bagi kedua-dua ruangan pembeli berdaftar dan tidak berdaftar." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila padam salah satu yang tidak berkenaan.", vbExclamation, "Info"
                    
        Exit Sub
          
    End If
End If
'### Periksa Samada Maklumat Pembeli Diisi Dalam Kedua-dua Ruangan Berdaftar Dan Tidak Berdaftar ### - End

If x <> 0 Then
    Frm6.Show
    Frm6.Pic1.Cls
    For Y = 1 To x
        Frm6.Pic1.Print Y & " - " & Err(Y)
    Next Y
    Exit Sub
Else

    If Frm84.L27_Text <> vbNullString And Frm84.L28_Text = vbNullString Then
        Note = "Adakah anda yakin untuk teruskan urusan jualan ini ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Data jualan akan disimpan ke dalam sistem." & vbCrLf & _
                vbNullString & vbCrLf & _
                "*** Oleh ini pembeli ini tidak berdaftar dengan sistem, oleh itu data dan rekod pembelian pembeli ini tidak akan disimpan di dalam sistem ***" & vbCrLf & _
                vbNullString & vbCrLf & _
                "***** Sistem mungkin akan mengambil masa untuk simpan semua data ini *****" & vbCrLf & _
                "Teruskan?"
    End If
    
    If Frm84.L27_Text = vbNullString And Frm84.L28_Text <> vbNullString Then
        Note = "Adakah anda yakin untuk teruskan urusan jualan ini ?" & vbCrLf & _
                vbNullString & vbCrLf & _
                "Data jualan akan disimpan ke dalam sistem." & vbCrLf & _
                vbNullString & vbCrLf & _
                "***** Sistem mungkin akan mengambil masa untuk simpan semua data ini *****" & vbCrLf & _
                "Teruskan?"
                
    End If
    
    If Frm84.L27_Text = vbNullString And Frm84.L28_Text = vbNullString Then
    
        Note = "TIADA maklumat bagi pembeli telah diisi." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Maklumat pembeli tidak akan dicetak di dalam invoice pembeli." & vbCrLf & _
                vbNullString & vbCrLf & _
                "***** Sistem mungkin akan mengambil masa untuk simpan semua data ini *****" & vbCrLf & _
                "Teruskan?"
        
    End If
    
    Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
    
    If Answer = vbYes Then
    
        LM_RATE_KUPON_2 = vbNullString
        Frm84_LM_KUPON = 0
        
        G_JENIS_URUSAN = 1
        
        If Frm84.L80_Text <> vbNullString Then
            If InStr(1, Frm84.L80_Text, " ") <> 0 Then
                LM_RATE_KUPON_1 = Split(Frm84.L80_Text, " ")(1)
                LM_RATE_KUPON_2 = Split(LM_RATE_KUPON_1, " ")(0)
            End If
        End If
        
        If LM_RATE_KUPON_2 <> vbNullString Then
            If IsNumeric(LM_RATE_KUPON_2) Then Frm84_LM_KUPON = LM_RATE_KUPON_2
        End If
        
        '$$$ No. staff $$$ - Start
        If InStr(1, Frm84.CBB1, "  |  ") <> 0 Then
            Frm84_LM_EMP_NO = Split(Frm84.CBB1, "  |  ")(1)
            Frm84_LM_EMP_NAMA = Split(Frm84.CBB1, "  |  ")(0)
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

'### Periksa status barang trade in ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 22_jualan where no_resit='" & Frm84.L3_Text & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!jenis_trade_in) Then '1 : Trade in (Voucher) , 2 : Belian dengan trade in
                If rs!jenis_trade_in = 1 Then '1 : Trade in (Voucher) , 2 : Belian dengan trade in
                    Frm84_LM_JENIS_TRADE_IN = 1 '1 : Trade in (Voucher) , 2 : Belian dengan trade in
                ElseIf rs!jenis_trade_in = 2 Then '1 : Trade in (Voucher) , 2 : Belian dengan trade in
                    Frm84_LM_JENIS_TRADE_IN = 2 '1 : Trade in (Voucher) , 2 : Belian dengan trade in
                ElseIf rs!jenis_trade_in = 3 Then '1 : Trade in (Voucher) , 2 : Belian dengan trade in
                    Frm84_LM_JENIS_TRADE_IN = 3 '1 : Trade in (Voucher) , 2 : Belian dengan trade in
                End If
            End If
            If Not IsNull(rs!no_resit_trade_in) Then Frm84_LM_No_VOUCHER_TI = rs!no_resit_trade_in
            If Not IsNull(rs!point_ari_nashi) Then
                If rs!point_ari_nashi = 1 Then Frm84_LM_FLAG_MATA_ASAL = 1 '0 : Tiada perolehan mata , 1 : Ada perolehan mata (Pada awal pembelian asal)
            End If
            If Not IsNull(rs!duit_simpanan_kedai) Then
                If rs!duit_simpanan_kedai <> "0.00" Then
                    If IsNumeric(rs!duit_simpanan_kedai) Then Frm84_LM_REFUND_ASAL = rs!duit_simpanan_kedai 'Refund : Jumlah Simpanan Asal
                    
                    If Not IsNull(rs!no_rujukan_pembeli) Then
                        Frm84_LM_No_PELANGGAN = rs!no_rujukan_pembeli 'No. Pelanggan
                    End If
                End If
            End If
        End If
        
        rs.Close
        Set rs = Nothing
'### Periksa status barang trade in ### - End
        
        LM_NOW = Now
        
'### Pulangkan point/mata kepada ahli ### - Start
        If Frm84_LM_FLAG_MATA_ASAL = 1 Then '0 : Tiada perolehan mata , 1 : Ada perolehan mata (Pada awal pembelian asal)
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 71_tebus_agih_point where no_invoice='" & Frm84.L3_Text & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If Not IsNull(rs!no_ahli) Then Frm84_LM_No_PELANGGAN_MATA = rs!no_ahli

                If Not IsNull(rs!jumlah_peroleh_point) Then 'Jumlah perolehan mata
                    If IsNumeric(rs!jumlah_peroleh_point) Then Frm84_LM_MATA_DAPAT = rs!jumlah_peroleh_point
                End If
                If Not IsNull(rs!jumlah_tebus_point) Then 'Jumlah mata yang ditebus
                    If IsNumeric(rs!jumlah_tebus_point) Then Frm84_LM_MATA_TEBUS = rs!jumlah_tebus_point
                End If
            End If
                
            rs.Close
            Set rs = Nothing
            
            If Frm84_LM_No_PELANGGAN_MATA <> vbNullString Then

                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm84_LM_No_PELANGGAN_MATA & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Not IsNull(rs!baki_point) Then 'Baki mata asal
                        If IsNumeric(rs!baki_point) Then Frm84_LM_MATA_ASAL = rs!baki_point
                    End If
                    rs!baki_point = Frm84_LM_MATA_ASAL + Frm84_LM_MATA_TEBUS - Frm84_LM_MATA_DAPAT
                    rs!remarks = "Pulangan mata ganjaran bagi tujuan edit data jualan"
                    rs!write_timestamp2 = LM_NOW
                    rs!no_staff = Frm84_LM_EMP_NO 'No. Pekerja
                    rs!terminal = G_TERMINAL
                    rs!jenis_urusan = G_JENIS_URUSAN
                    rs.Update
                End If
                
                rs.Close
                Set rs = Nothing
            
            End If
            
        End If
'### Pulangkan point/mata kepada ahli ### - End
        
'### Pulangkan status barang trade in ### - Start
        If Frm84_LM_JENIS_TRADE_IN = 1 Or Frm84_LM_JENIS_TRADE_IN = 2 Then '1 : Trade in (Voucher) , 2 : Belian dengan trade in
            '### Update Maklumat Trade In ### - Start

            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & Frm84_LM_No_VOUCHER_TI & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                G_ID = rs!ID
                Call recovery_16_gold_bar_belian
            
                rs!trade_in_status = 0
                rs!no_staff = Frm84_LM_EMP_NO 'No. Pekerja
                rs!terminal = G_TERMINAL
                rs!write_timestamp2 = LM_NOW
                rs!jenis_urusan = G_JENIS_URUSAN
                rs!remarks = "Pulangkan status trade in bagi tujuan edit data jualan"
                rs.Update
            End If

            rs.Close
            Set rs = Nothing

            '### Update Maklumat Trade In ### - End
        End If
'### Pulangkan status barang trade in ### - End
    
'### Periksa Samada Ada Pembayaran Menggunakan Simpanan Duit Di Kedai ### - Start 13/07/2015
        If Frm84_LM_No_PELANGGAN <> vbNullString Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm84_LM_No_PELANGGAN & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
                If IsNumeric(rs!baki_simpanan) Then Frm84_LM_REFUND_GUNA = rs!baki_simpanan 'Refund : Jumlah Simpan Yang Telah Digunakan Sebelum Ini
                
                rs!baki_simpanan = Format(Frm84_LM_REFUND_ASAL + Frm84_LM_REFUND_GUNA, "0.00") 'Baki Simpanan
                rs!write_timestamp2 = LM_NOW
                rs!no_staff = Frm84_LM_EMP_NO 'No. Pekerja
                rs!terminal = G_TERMINAL
                rs!jenis_urusan = G_JENIS_URUSAN
                
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
            
'### Padam Rekod Penggunaan Duit Pelanggan ### - Start
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            strsql = "DELETE from 24_rekod_kewangan_pelanggan where NO_RESIT='" & Frm84.L3_Text & "'"
            
            Set rs = cn.Execute(strsql)
            Set rs = Nothing
'### Padam Rekod Penggunaan Duit Pelanggan ### - End

        End If
'### Periksa Samada Ada Pembayaran Menggunakan Simpanan Duit Di Kedai ### - End
        
        LM_STATUS_R = 0
        
'###Masukkan Jualan Ke Dalam Table Akaun Jualan### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 22_jualan where no_resit='" & Frm84.L3_Text & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            G_ID = rs!ID
            Call recovery_22_jualan
            
            rs!tarikh = Frm84.DTPicker1 'Tarikh Jualan
            If Not IsNull(rs!status_r) Then LM_STATUS_R = rs!status_r
            rs!tunai = Format(0, "0.00")
            rs!bank_in = Format(0, "0.00") 'Cara Bayaran : Bank In
            rs!kad_kredit = Format(0, "0.00") 'Cara Bayaran : Kad Kredit
            rs!duit_simpanan_kedai = Format(0, "0.00") 'Cara Bayaran : Simpanan Duit Di Kedai
            rs!cek = Format(0, "0.00") 'Cara Bayaran : Cek
            
            If Frm84.L25_Text = "Jumlah Bayaran" Then
                If frm130.TB27 <> vbNullString Then
                    rs!tunai = Format(frm130.TB27, "0.00") 'Cara Bayaran : Tunai
                Else
                    rs!tunai = Null 'Cara Bayaran : Tunai
                End If
                If frm130.TB28 <> vbNullString Then
                    rs!bank_in = Format(frm130.TB28, "0.00") 'Cara Bayaran : Bank In
                Else
                    rs!bank_in = Null 'Cara Bayaran : Bank In
                End If
                If frm130.TB29 <> vbNullString Then
                    rs!kad_kredit = Format(frm130.TB29, "0.00") 'Cara Bayaran : Kad Kredit
                    If Format(frm130.TB29, "0.00") <> "0.00" Then
                        
                        If frm130.CBB2 <> vbNullString Then
                            rs!jenis_kad = frm130.CBB2
                        Else
                            rs!jenis_kad = Null
                        End If
                        If frm130.L31_Text <> vbNullString Then
                            rs!cas_Kad_Kredit = Format(frm130.L31_Text, "0.00") 'Cara Bayaran : Cas Kad Kredit (%)
                        Else
                            rs!cas_Kad_Kredit = "0.00" 'Cara Bayaran : Cas Kad Kredit (%)
                        End If
                        If frm130.L32_Text <> vbNullString Then
                            rs!jumlah_cas_kad_kredit = Format(frm130.L32_Text, "0.00") 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
                        Else
                            rs!jumlah_cas_kad_kredit = "0.00" 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
                        End If
                        If frm130.L81_Text <> vbNullString Then
                            rs!gst_kad_kredit = Format(frm130.L81_Text, "0.00") 'Cara Bayaran : Jumlah GST kad kredit (RM)
                        Else
                            rs!gst_kad_kredit = "0.00" 'Cara Bayaran : Jumlah GST kad kredit (RM)
                        End If
                        If frm130.L82_Text <> vbNullString Then
                            rs!jumlah_potongan_kad_kredit = Format(frm130.L82_Text, "0.00") 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
                        Else
                            rs!jumlah_potongan_kad_kredit = "0.00" 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
                        End If
                        If Frm84.L8_Text <> vbNullString Then
                            rs!kadar_gst_kad_kredit = Format(Frm84.L8_Text, "0.00") 'Cara Bayaran : Kadar GST bagi kad kredit
                        Else
                            rs!kadar_gst_kad_kredit = "0.00" 'Cara Bayaran : Kadar GST bagi kad kredit
                        End If
                    
                        If Frm84.CB19 = 1 Then
                            rs!epp = 1 '0 : Bayaran selain dari EPP , 1 : Bayaran secara EPP
                            rs!approval_code_epp = UCase(Frm84.TB41) 'Approval Code (EPP)
                        Else
                            rs!epp = 0 '0 : Bayaran selain dari EPP , 1 : Bayaran secara EPP
                            rs!approval_code_epp = Null 'Approval Code (EPP)
                        End If
                    Else
                        rs!jenis_kad = Null
                        rs!cas_Kad_Kredit = Null 'Cara Bayaran : Cas Kad Kredit (%)
                        rs!jumlah_cas_kad_kredit = Null 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
                        rs!jumlah_potongan_kad_kredit = Null 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
                        rs!kadar_gst_kad_kredit = Null 'Cara Bayaran : Kadar GST bagi kad kredit
                        rs!gst_kad_kredit = Null 'Cara Bayaran : Jumlah GST kad kredit (RM)
                        
                        rs!epp = 0 '0 : Bayaran selain dari EPP , 1 : Bayaran secara EPP
                        rs!approval_code_epp = Null 'Approval Code (EPP)
                    End If
                Else
                    rs!jenis_kad = Null
                    rs!cas_Kad_Kredit = Null 'Cara Bayaran : Cas Kad Kredit (%)
                    rs!jumlah_cas_kad_kredit = Null 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
                    rs!jumlah_potongan_kad_kredit = Null 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
                    rs!kadar_gst_kad_kredit = Null 'Cara Bayaran : Kadar GST bagi kad kredit
                    rs!gst_kad_kredit = Null 'Cara Bayaran : Jumlah GST kad kredit (RM)
                    rs!epp = 0 '0 : Bayaran selain dari EPP , 1 : Bayaran secara EPP
                    rs!approval_code_epp = Null 'Approval Code (EPP)
                    'rs!kad_kredit = Null 'Cara Bayaran : Kad Kredit
                End If

                If frm130.TB21 <> vbNullString Then
                    If Format(frm130.TB21, "0.00") <> "0.00" Then
                        Frm84_LM_Flag_SIMPANAN = 1 '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
                    End If
                    rs!duit_simpanan_kedai = Format(frm130.TB21, "0.00") 'Cara Bayaran : Simpanan Duit Di Kedai
                Else
                    rs!duit_simpanan_kedai = "0.00" 'Cara Bayaran : Simpanan Duit Di Kedai
                End If
                If frm130.TB32 <> vbNullString Then
                    rs!jumlah_bayaran = Format(frm130.TB32, "0.00") 'Cara Bayaran : Jumlah Bayaran
                Else
                    rs!jumlah_bayaran = Null 'Cara Bayaran : Jumlah Bayaran
                End If
            Else
                rs!tunai = "0.00" 'Cara Bayaran : Tunai
                rs!bank_in = "0.00" 'Cara Bayaran : Bank In
                rs!kad_kredit = "0.00" 'Cara Bayaran : Kad Kredit
                If frm130.L31_Text <> vbNullString Then
                    rs!cas_Kad_Kredit = Format(frm130.L31_Text, "0.00") 'Cara Bayaran : Cas Kad Kredit (%)
                Else
                    rs!cas_Kad_Kredit = 0 'Cara Bayaran : Cas Kad Kredit (%)
                End If
                rs!epp = 0 '0 : Bayaran selain dari EPP , 1 : Bayaran secara EPP
                rs!approval_code_epp = Null 'Approval Code (EPP)
                rs!jumlah_cas_kad_kredit = "0.00" 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
                rs!jumlah_potongan_kad_kredit = "0.00" 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
                rs!duit_simpanan_kedai = "0.00" 'Cara Bayaran : Simpanan Duit Di Kedai
                rs!kad_debit = "0.00" 'Cara Bayaran : Kad Debit
                If frm130.L32_Text <> vbNullString Then
                    rs!cas_kad_debit = frm130.L32_Text 'Cara Bayaran : Jumlah Cas Kad Debit (%)
                Else
                    rs!cas_kad_debit = 0 'Cara Bayaran : Jumlah Cas Kad Debit (%)
                End If
                rs!jumlah_cas_kad_debit = "0.00" 'Cara Bayaran : Jumlah Cas Kad Debit (RM)
                rs!jumlah_potongan_kad_debit = "0.00" 'Cara Bayaran : Jumah Potongan Kad Debit (RM)
                rs!jumlah_bayaran = "0.00" 'Cara Bayaran : Jumlah Bayaran
                rs!jenis_kad = Null
                rs!cas_Kad_Kredit = Null 'Cara Bayaran : Cas Kad Kredit (%)
                rs!jumlah_cas_kad_kredit = Null 'Cara Bayaran : Jumlah Cas Kad Kredit (RM)
                rs!jumlah_potongan_kad_kredit = Null 'Cara Bayaran : Jumlah Potongan Kad Kredit (RM)
                rs!kadar_gst_kad_kredit = Null 'Cara Bayaran : Kadar GST bagi kad kredit
                rs!gst_kad_kredit = Null 'Cara Bayaran : Jumlah GST kad kredit (RM)
                rs!epp = 0 '0 : Bayaran selain dari EPP , 1 : Bayaran secara EPP
                rs!approval_code_epp = Null 'Approval Code (EPP)
                'rs!kad_kredit = Null 'Cara Bayaran : Kad Kredit

            End If

            If Frm84.L17_Text <> vbNullString Then
                rs!harga_barang = Format(Frm84.L17_Text, "0.00") 'Jumlah Harga Barang Tanpa GST (RM)
            Else
                rs!harga_barang = Null 'Jumlah Harga Barang Tanpa GST (RM)
            End If
            If Frm84.L18_Text <> vbNullString Then
                rs!jumlah_cukai_gst = Format(Frm84.L18_Text, "0.00") 'Jumlah Cukai GST (ZR + SR)
            Else
                rs!jumlah_cukai_gst = Null 'Jumlah Cukai GST (ZR + SR)
            End If
            If Frm84.L19_Text <> vbNullString Then
                rs!harga_barang_dengan_gst = Format(Frm84.L19_Text, "0.00") 'Jumlah Harga Barang Dengan GST (RM)
            Else
                rs!harga_barang_dengan_gst = Null 'Jumlah Harga Barang Dengan GST (RM)
            End If
            If Frm84.TB19 <> vbNullString Then
                rs!diskaun = Format(Frm84.TB19, "0.00") 'Jumlah Diskaun (%)
            Else
                rs!diskaun = Null 'Jumlah Diskaun (%)
            End If
            If Frm84.L20_Text <> vbNullString Then
                rs!harga_lepas_diskaun = Format(Frm84.L20_Text, "0.00") 'Harga Selepas Diskaun (RM)
            Else
                rs!harga_lepas_diskaun = Null 'Harga Selepas Diskaun (RM)
            End If
            If Frm84.TB20 <> vbNullString Then
                rs!adjustment = Format(Frm84.TB20, "0.00") 'Adjustment (RM)
            Else
                rs!adjustment = Null 'Adjustment (RM)
            End If
            If Frm84.L21_Text <> vbNullString Then
                rs!harga_jualan = Format(Frm84.L21_Text, "0.00") 'Jumlah Harga Jualan (RM)
            Else
                rs!harga_jualan = Null 'Jumlah Harga Jualan (RM)
            End If
            If Frm84.L38_Text <> vbNullString Then
                rs!loss_trade_in = Format(Frm84.L38_Text, "0.00") 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (%)
            Else
                rs!loss_trade_in = Null 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (%)
            End If
            If Frm84.L37_Text <> vbNullString Then
                rs!loss_trade_in_rm = Format(Frm84.L37_Text, "0.00") 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (RM)
            Else
                rs!loss_trade_in_rm = Null 'Potongan Harga Bagi Trade In Jika Kedai Perlu Bayar (RM)
            End If
            If Frm84.L24_Text = "Jumlah Bayaran" Then
                rs!flag_bayaran = 0 '0 : Pembeli Bayar , 1 : Kedai Bayar
            Else
                rs!flag_bayaran = 1 '0 : Pembeli Bayar , 1 : Kedai Bayar
            End If
            If Frm84.L23_Text <> vbNullString Then
                rs!jumlah_perlu_bayar = Format(Frm84.L23_Text, "0.00") 'Jumlah Bayaran Yang Perlu Dibuat (RM)
            Else
                rs!jumlah_perlu_bayar = Null 'Jumlah Bayaran Yang Perlu Dibuat (RM)
            End If
            If Frm84.L14_Text <> vbNullString Then
                rs!kuantiti_barang = Frm84.L14_Text 'Kuantiti Barang Yang Dijual
            Else
                rs!kuantiti_barang = Null 'Kuantiti Barang Yang Dijual
            End If
            If Frm84.L15_Text <> vbNullString Then
                rs!JUMLAH_BERAT = Frm84.L15_Text 'Jumlah Berat Barang Yang Dijual
            Else
                rs!JUMLAH_BERAT = Null 'Jumlah Berat Barang Yang Dijual
            End If
            If Frm84.L7_Text <> vbNullString Then
                rs!gst_zr_harga = Format(Frm84.L7_Text, "0.00") 'Harga Keseluruhan Bagi Barang ZR
            Else
                rs!gst_zr_harga = Null 'Harga Keseluruhan Bagi Barang ZR
            End If
            If Frm84.L9_Text <> vbNullString Then
                rs!gst_zr_cukai = Format(Frm84.L9_Text, "0.00") 'Jumlah Cukai Bagi ZR
            Else
                rs!gst_zr_cukai = Null 'Jumlah Cukai Bagi ZR
            End If
            If Frm84.L10_Text <> vbNullString Then
                rs!gst_sr_harga = Format(Frm84.L10_Text, "0.00") 'Harga Keseluruhan Bagi Barang SR
            Else
                rs!gst_sr_harga = Null 'Harga Keseluruhan Bagi Barang SR
            End If
            If Frm84.L11_Text <> vbNullString Then
                rs!gst_sr_cukai = Format(Frm84.L11_Text, "0.00") 'Jumlah Cukai Bagi SR
            Else
                rs!gst_sr_cukai = Null 'Jumlah Cukai Bagi SR
            End If

            rs!no_pekerja = Frm84_LM_EMP_NO 'No. Pekerja

            If Frm84.L28_Text <> vbNullString Then
                If Frm28.L5_Text <> vbNullString Then
                    rs!no_rujukan_pembeli = Frm28.L5_Text 'No. Rujukan Pembeli
                Else
                    rs!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
                End If
            Else
                rs!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
            End If
            If Frm84.CB7 = 1 Then
                If Frm27.L5_Text <> vbNullString Then
                    rs!no_rujukan_agen_dropship = Frm27.L5_Text 'No. Rujukan Agen Dropship
                Else
                    rs!no_rujukan_agen_dropship = Null 'No. Rujukan Agen Dropship
                End If
            Else
                rs!no_rujukan_agen_dropship = Null 'No. Rujukan Agen Dropship
            End If
            
            If G_TI_MODE <> 3 Then
                If Frm84.L56_Text <> 0 Then
                    Frm84_LM_Flag_TRADE_IN = 1 '0 : Tiada Urusan Trade In , 1 : Ada Urusan Trade In
                    rs!flag_trade_in = 1 '0 : Tiada Urusan Trade in , 1 : Ada Urusan Trade In
                    
                    If Frm84.L56_Text = 1 Then
                        rs!jenis_trade_in = 1 '1 : Trade in (Voucher) , 2 : Belian dengan trade in
                    ElseIf Frm84.L56_Text = 2 Then
                        rs!jenis_trade_in = 2 '1 : Trade in (Voucher) , 2 : Belian dengan trade in
                    End If
                    
                    If Frm84.L57_Text <> vbNullString Then
                        rs!no_resit_trade_in = Frm84.L57_Text 'No. Resit Trade In
                    Else
                        rs!no_resit_trade_in = Null 'No. Resit Trade In
                    End If
                    If Frm84.L58_Text <> vbNullString Then
                        rs!jumlah_trade_in = Format(Frm84.L58_Text, "0.00") 'No. Resit Trade In
                    Else
                        rs!jumlah_trade_in = Null 'No. Resit Trade In
                    End If
                Else
                    rs!flag_trade_in = 0 '0 : Tiada Urusan Trade in , 1 : Ada Urusan Trade In
                    rs!no_resit_trade_in = Null 'No. Resit Trade In
                    rs!jumlah_trade_in = Null 'Jumlah Resit Trade In (RM)
                    rs!jenis_trade_in = Null '1 : Trade in (Voucher) , 2 : Belian dengan trade in
                End If
                rs!jumlah_caj_tukaran = Null
            Else
                rs!flag_trade_in = 1 '0 : Tiada Urusan Trade in , 1 : Ada Urusan Trade In
                rs!jenis_trade_in = 3 '1 : Trade in (Voucher) , 2 : Belian dengan trade in
                rs!jumlah_trade_in = G_TRADE_IN_TOTAL
                rs!jumlah_caj_tukaran = G_TRADE_IN_CAJ

                If Frm84.TB49 <> vbNullString Then
                    rs!berat_trade_in = Frm84.TB49
                Else
                    rs!berat_trade_in = 0
                End If
                If Frm84.TB50 <> vbNullString Then
                    rs!harga_semasa_trade_in = Frm84.TB50
                Else
                    rs!harga_semasa_trade_in = 0
                End If
                If Frm84.TB51 <> vbNullString Then
                    rs!harga_semasa_buyback = Frm84.TB51
                Else
                    rs!harga_semasa_buyback = 0
                End If
                If Frm84.TB52 <> vbNullString Then
                    rs!caj_pertukaran = Frm84.TB52
                Else
                    rs!caj_pertukaran = 0
                End If
            End If
            If Frm84.CB4 = 1 Then
                rs!kategori_pembeli = 1
            ElseIf Frm84.CB5 = 1 Then
                rs!kategori_pembeli = 2
            ElseIf Frm84.CB6 = 1 Then
                rs!kategori_pembeli = 3
            ElseIf Frm84.CB9 = 1 Then
                rs!kategori_pembeli = 4
            ElseIf Frm84.CB10 = 1 Then
                rs!kategori_pembeli = 5
            'ElseIf Frm84.CB11 = 1 Then
            '    rs!kategori_pembeli = 6
            End If
                        
            If Frm84.CB27 = 1 Then
                rs!jualan_online = 1
            Else
                rs!jualan_online = 0
            End If
            If Frm84.TB42 <> vbNullString Then 'Jumlah caj pos laju (postage)
                rs!caj_pos = Format(Frm84.TB42, "0.00")
            Else
                rs!caj_pos = "0.00"
            End If
            If Frm84.TB45 <> vbNullString Then 'No. Tracking pos laju
                rs!no_tracking = UCase(Frm84.TB45)
            Else
                rs!no_tracking = Null
            End If
            If Frm84.L79_Text = 0 Then
                rs!point_ari_nashi = 0
            ElseIf Frm84.L79_Text = 1 Then
                rs!point_ari_nashi = 1
            End If
            If Frm84.L76_Text <> vbNullString Then
                rs!jumlah_point = Frm84.L76_Text
            Else
                rs!jumlah_point = 0
            End If
            If Frm84.TB34 <> vbNullString Then
                rs!kupon_diskaun = Format(Frm84.TB34, "0.00")
            Else
                rs!kupon_diskaun = "0.00"
            End If
            If Frm84.TB35 <> vbNullString Then
                rs!kadar_peroleh_point = Frm84.TB35
            Else
                rs!kadar_peroleh_point = 0
            End If
            If Frm84.TB37 <> vbNullString Then
                rs!kadar_tebus_point = Frm84.TB37
            Else
                rs!kadar_tebus_point = 0
            End If
            rs!kadar_diskaun = Format(Frm84_LM_KUPON, "0.00") 'Kadar diskaun per gram
            rs!Status = 1
            rs!terminal = G_TERMINAL
            rs!no_staff = G_LOGIN_USER
            rs!write_timestamp2 = LM_NOW
            If Frm84.L73_Text <> vbNullString Then
                rs!redeem_point = Frm84.L73_Text
            Else
                rs!redeem_point = 0
            End If
            rs!Menu = 0
            'rs!cawangan = G_CAWANGAN
            rs!nama_pekerja = Frm84_LM_EMP_NAMA
            If Not IsNull(rs!cawangan) Then LM_CAWANGAN = rs!cawangan
            If Frm84.TB46 <> vbNullString Then
                rs!remarks = Frm84.TB46
            Else
                rs!remarks = Null
            End If
            
            DATA_SAVE = 1
            rs.Update
        End If
        
        rs.Close
        Set rs = Nothing
'###Masukkan Jualan Ke Dalam Table Akaun Jualan### - End

'###Update Data Simpanan Duit Pelanggan### - Start
        If Frm84_LM_Flag_SIMPANAN = 1 And Frm84.L24_Text = "Jumlah Bayaran" And Frm84.L28_Text <> vbNullString Then '0 : Tiada Penggunakan Duit Simpanan Kedai , 1 : Ada Penggunakan Duit Simpanan Kedai
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If Not rs.EOF Then
            
                G_ID = rs!ID
                Call recovery_senarai_pelanggan
                
                If IsNumeric(frm130.L26_Text) Then Frm84_LM_JUMLAH_SIMPANAN = frm130.L26_Text  'Jumlah Simpanan Yang Ada
                If IsNumeric(frm130.TB21) Then Frm84_LM_GUNA_SIMPAN = frm130.TB21  'Jumlah Simpanan Yang Hendak Digunakan
                
                rs!baki_simpanan = Format(Frm84_LM_JUMLAH_SIMPANAN - Frm84_LM_GUNA_SIMPAN, "0.00") 'Baki Simpanan
                rs!write_timestamp2 = LM_NOW
                rs!no_staff = Frm84_LM_EMP_NO 'No. Pekerja
                rs!terminal = G_TERMINAL
                rs!jenis_urusan = G_JENIS_URUSAN
                
                rs.Update
                
            End If
            
            rs.Close
            Set rs = Nothing
            
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 24_rekod_kewangan_pelanggan where id is null", cn, adOpenKeyset, adLockOptimistic
            
            If rs.EOF Then
                rs.AddNew
                rs!tarikh = Frm84.DTPicker1 'Tarikh
                rs!jenis = 1 '0 : Simpanan , 1 : Penggunaan Duit
                rs!no_rujukan_pelanggan = Frm28.L5_Text 'No. Rujukan Pelanggan
                rs!no_resit = Frm84.L3_Text 'No. Resit Jualan
                rs!jumlah = Format(frm130.TB21, "0.00") 'Jumlah Simpanan Yang Digunakan (RM)
                rs!jenis_penggunaan = 0 '0 : Belian Barangan Kemas , 1 : Ansuran , 2 : Tempahan (Deposit) , 3 : Servis , 4 : Tempahan (Ambilan Barang)
                rs!no_rujukan_pekerja = Frm84_LM_EMP_NO 'No. Pekerja
                rs!terminal = G_TERMINAL
                rs!write_timestamp = LM_NOW
                rs!jenis_urusan = G_JENIS_URUSAN
                rs!cawangan = G_CAWANGAN
                rs!Status = 1
                rs.Update
            End If
            
            rs.Close
            Set rs = Nothing
           
        End If
        
'### Padam Rekod Belian Pembeli Yang Tidak Berdaftar Dengan Kedai ### - Start (08-07-2015)
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        strsql = "DELETE from 44_senarai_pelanggan where no_resit='" & Frm84.L3_Text & "'"
        
        Set rs = cn.Execute(strsql)
        Set rs = Nothing
'### Padam Rekod Belian Pembeli Yang Tidak Berdaftar Dengan Kedai ### - End (08-07-2015)
        
'### Maklumat Pembeli Yang Tidak Berdaftar Dengan Kedai ### - Start (08-07-2015)
        If Frm84.L27_Text <> vbNullString Then
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 44_senarai_pelanggan where no_resit='" & Frm84.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
            
            If rs.EOF Then
                rs.AddNew
                rs!tarikh = Frm84.DTPicker1 'Tarikh
                rs!no_resit = Frm84.L3_Text 'No. Resit Jualan
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
                rs!write_timestamp = LM_NOW
                rs!no_staff = Frm84_LM_EMP_NO 'No. Pekerja
                rs!terminal = G_TERMINAL
                rs!jenis_urusan = G_JENIS_URUSAN
                rs!cawangan = G_CAWANGAN
                rs.Update
            Else
                G_ID = rs!ID
                Call recovery_44_senarai_pelanggan
            
                rs!tarikh = Frm84.DTPicker1 'Tarikh
                rs!no_resit = Frm84.L3_Text 'No. Resit Jualan
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
                rs!write_timestamp = LM_NOW
                rs!no_staff = Frm84_LM_EMP_NO 'No. Pekerja
                rs!terminal = G_TERMINAL
                rs!jenis_urusan = G_JENIS_URUSAN
                rs!cawangan = G_CAWANGAN
                rs.Update
            
            End If
            
            rs.Close
            Set rs = Nothing
            
        End If
        
'### Maklumat agihan point ### - Start
        If Frm84.L79_Text = 1 Then

            If Frm84_LM_FLAG_MATA_ASAL = 1 Then '0 : Tiada perolehan mata , 1 : Ada perolehan mata (Pada awal pembelian asal)
            
'@ Edit data jualan. Jualan asal adalah kepada ahli kedai.

                'Jika data ahli lama dan ahli baru adalah orang yang sama
                If Frm28.L5_Text = Frm84_LM_No_PELANGGAN_MATA Then

                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from 71_tebus_agih_point where no_invoice='" & Frm84.L3_Text & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                        
                        G_ID = rs!ID
                        Call recovery_71_tebus_agih_point
                        
                        rs!tarikh = Frm84.DTPicker1 'Tarikh Jualan
                        If Frm28.L5_Text <> vbNullString Then 'No. Rujukan Pembeli
                            rs!no_ahli = Frm28.L5_Text
                        Else
                            rs!no_ahli = Null
                        End If
                        If Frm84.L75_Text <> vbNullString Then 'Harga yang membolehkan untuk mendaparkan point
                            rs!harga_layak_bonus = Format(Frm84.L75_Text, "0.00")
                        Else
                            rs!harga_layak_bonus = Null
                        End If
                        If Frm84.TB35 <> vbNullString Then 'Kadar perolehan point (eg. 0.5)
                            rs!kadar_peroleh_point = Frm84.TB35
                        Else
                            rs!kadar_peroleh_point = Null
                        End If
                        If Frm84.L76_Text <> vbNullString Then 'Jumlah perolehan mata
                            rs!jumlah_peroleh_point = Frm84.L76_Text
                        Else
                            rs!jumlah_peroleh_point = Null
                        End If
                        If Frm84.TB36 <> vbNullString Then 'Jumlah mata yang ditebus
                            rs!jumlah_tebus_point = Frm84.TB36
                        Else
                            rs!jumlah_tebus_point = Null
                        End If
                        If Frm84.TB37 <> vbNullString Then 'Kadar tebusan mata
                            rs!kadar_tebus_point = Frm84.TB37
                        Else
                            rs!kadar_tebus_point = Null
                        End If
                        If Frm84.L78_Text <> vbNullString Then 'Jumlah nilaian mata yang ditebus
                            rs!nilaian_tebus_point = Frm84.L78_Text
                        Else
                            rs!nilaian_tebus_point = Null
                        End If
                        If Frm84.CB13 = 0 Then
                            rs!bil_rasmi = 1 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
                        Else
                            rs!bil_rasmi = 0 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
                        End If
                        rs!write_timestamp2 = LM_NOW
                        rs!Status = 1 '0 : Tidak aktif , 1 : Aktif
                        rs!Type = 1
                        rs!terminal = G_TERMINAL
                        rs!jenis_urusan = G_JENIS_URUSAN
                        
                        rs.Update
                    End If
                        
                    rs.Close
                    Set rs = Nothing
                
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                        G_ID = rs!ID
                        Call recovery_senarai_pelanggan
                        If Frm84.L76_Text <> vbNullString Then 'Jumlah perolehan mata
                            If IsNumeric(Frm84.L76_Text) Then Frm84_LM_MATA_DAPAT = Frm84.L76_Text
                        End If
                        If Frm84.TB36 <> vbNullString Then 'Jumlah mata yang ditebus
                            If IsNumeric(Frm84.TB36) Then Frm84_LM_MATA_TEBUS = Frm84.TB36
                        End If
                        If Frm84.L77_Text <> vbNullString Then 'Jumlah mata asal
                            If IsNumeric(Frm84.L77_Text) Then Frm84_LM_MATA_ASAL = Frm84.L77_Text
                        End If
                        rs!baki_point = Frm84_LM_MATA_ASAL - Frm84_LM_MATA_TEBUS + Frm84_LM_MATA_DAPAT
                        rs!write_timestamp2 = LM_NOW
                        rs!no_staff = Frm84_LM_EMP_NO 'No. Pekerja
                        rs!terminal = G_TERMINAL
                        rs!jenis_urusan = G_JENIS_URUSAN

                        rs.Update
                        
                    End If
                    
                    rs.Close
                    Set rs = Nothing
                    
                Else
                
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from 71_tebus_agih_point where no_invoice='" & Frm84.L3_Text & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                        rs!Status = 0
                        rs!write_timestamp3 = LM_NOW
                        rs!terminal = G_TERMINAL
                        rs!jenis_urusan = G_JENIS_URUSAN
                        rs.Update
                    End If
                    
                    rs.Close
                    Set rs = Nothing

                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from 71_tebus_agih_point where no_invoice='" & Frm84.L3_Text & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
                    
                    If rs.EOF Then
                        rs.AddNew
                        If Frm84.L3_Text <> vbNullString Then
                            rs!no_invoice = Frm84.L3_Text 'No. Resit Jualan
                        Else
                            rs!no_invoice = Null 'No. Resit Jualan
                        End If
                        If Frm84.DTPicker1 <> vbNullString Then
                            rs!tarikh = Frm84.DTPicker1 'Tarikh Jualan
                        Else
                            rs!tarikh = Null 'Tarikh Jualan
                        End If
                        If Frm28.L5_Text <> vbNullString Then 'No. Rujukan Pembeli
                            rs!no_ahli = Frm28.L5_Text
                        Else
                            rs!no_ahli = Null
                        End If
                        If Frm84.L75_Text <> vbNullString Then 'Harga yang membolehkan untuk mendaparkan point
                            rs!harga_layak_bonus = Format(Frm84.L75_Text, "0.00")
                        Else
                            rs!harga_layak_bonus = Null
                        End If
                        If Frm84.TB35 <> vbNullString Then 'Kadar perolehan point (eg. 0.5)
                            rs!kadar_peroleh_point = Frm84.TB35
                        Else
                            rs!kadar_peroleh_point = Null
                        End If
                        If Frm84.L76_Text <> vbNullString Then 'Jumlah perolehan mata
                            rs!jumlah_peroleh_point = Frm84.L76_Text
                        Else
                            rs!jumlah_peroleh_point = Null
                        End If
                        If Frm84.TB36 <> vbNullString Then 'Jumlah mata yang ditebus
                            rs!jumlah_tebus_point = Frm84.TB36
                        Else
                            rs!jumlah_tebus_point = Null
                        End If
                        If Frm84.TB37 <> vbNullString Then 'Kadar tebusan mata
                            rs!kadar_tebus_point = Frm84.TB37
                        Else
                            rs!kadar_tebus_point = Null
                        End If
                        If Frm84.L78_Text <> vbNullString Then 'Jumlah nilaian mata yang ditebus
                            rs!nilaian_tebus_point = Frm84.L78_Text
                        Else
                            rs!nilaian_tebus_point = Null
                        End If
                        If Frm84.CB13 = 0 Then
                            rs!bil_rasmi = 1 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
                        Else
                            rs!bil_rasmi = 0 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
                        End If
                        rs!write_timestamp = LM_NOW
                        rs!Status = 1 '0 : Tidak aktif , 1 : Aktif
                        rs!Type = 1
                        rs!terminal = G_TERMINAL
                        rs!jenis_urusan = G_JENIS_URUSAN
                        rs!cawangan = G_CAWANGAN
                    
                        rs.Update
                    End If
                    
                    rs.Close
                    Set rs = Nothing
                    
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs.EOF Then
                        If Frm84.L76_Text <> vbNullString Then 'Jumlah perolehan mata
                            If IsNumeric(Frm84.L76_Text) Then Frm84_LM_MATA_DAPAT = Frm84.L76_Text
                        End If
                        If Frm84.TB36 <> vbNullString Then 'Jumlah mata yang ditebus
                            If IsNumeric(Frm84.TB36) Then Frm84_LM_MATA_TEBUS = Frm84.TB36
                        End If
                        If Frm84.L77_Text <> vbNullString Then 'Jumlah mata asal
                            If IsNumeric(Frm84.L77_Text) Then Frm84_LM_MATA_ASAL = Frm84.L77_Text
                        End If
                        rs!baki_point = Frm84_LM_MATA_ASAL - Frm84_LM_MATA_TEBUS + Frm84_LM_MATA_DAPAT
                        rs!write_timestamp2 = LM_NOW
                        rs!no_staff = Frm84_LM_EMP_NO 'No. Pekerja
                        rs!terminal = G_TERMINAL
                        rs!jenis_urusan = G_JENIS_URUSAN
                        rs.Update
                    End If
                    
                    rs.Close
                    Set rs = Nothing
                

                End If
                
            ElseIf Frm84_LM_FLAG_MATA_ASAL = 0 Then '0 : Tiada perolehan mata , 1 : Ada perolehan mata (Pada awal pembelian asal)
            
'@ Kemasukkan data baru. Pembelian asal adalah dari bukan ahli kedai tetapi ditukar kepada ahli yang mempunyai kad.

                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 71_tebus_agih_point where no_invoice='" & Frm84.L3_Text & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
                
                If rs.EOF Then
                    rs.AddNew
                    If Frm84.L3_Text <> vbNullString Then
                        rs!no_invoice = Frm84.L3_Text 'No. Resit Jualan
                    Else
                        rs!no_invoice = Null 'No. Resit Jualan
                    End If
                    If Frm84.DTPicker1 <> vbNullString Then
                        rs!tarikh = Frm84.DTPicker1 'Tarikh Jualan
                    Else
                        rs!tarikh = Null 'Tarikh Jualan
                    End If
                    If Frm28.L5_Text <> vbNullString Then 'No. Rujukan Pembeli
                        rs!no_ahli = Frm28.L5_Text
                    Else
                        rs!no_ahli = Null
                    End If
                    If Frm84.L75_Text <> vbNullString Then 'Harga yang membolehkan untuk mendaparkan point
                        rs!harga_layak_bonus = Format(Frm84.L75_Text, "0.00")
                    Else
                        rs!harga_layak_bonus = Null
                    End If
                    If Frm84.TB35 <> vbNullString Then 'Kadar perolehan point (eg. 0.5)
                        rs!kadar_peroleh_point = Frm84.TB35
                    Else
                        rs!kadar_peroleh_point = Null
                    End If
                    If Frm84.L76_Text <> vbNullString Then 'Jumlah perolehan mata
                        rs!jumlah_peroleh_point = Frm84.L76_Text
                    Else
                        rs!jumlah_peroleh_point = Null
                    End If
                    If Frm84.TB36 <> vbNullString Then 'Jumlah mata yang ditebus
                        rs!jumlah_tebus_point = Frm84.TB36
                    Else
                        rs!jumlah_tebus_point = Null
                    End If
                    If Frm84.TB37 <> vbNullString Then 'Kadar tebusan mata
                        rs!kadar_tebus_point = Frm84.TB37
                    Else
                        rs!kadar_tebus_point = Null
                    End If
                    If Frm84.L78_Text <> vbNullString Then 'Jumlah nilaian mata yang ditebus
                        rs!nilaian_tebus_point = Frm84.L78_Text
                    Else
                        rs!nilaian_tebus_point = Null
                    End If
                    If Frm84.CB13 = 0 Then
                        rs!bil_rasmi = 1 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
                    Else
                        rs!bil_rasmi = 0 'Jenis invoice , 0 : Tidak rasmi , 1 : Rasmi
                    End If
                    rs!write_timestamp = LM_NOW
                    rs!Status = 1 '0 : Tidak aktif , 1 : Aktif
                    rs!Type = 1
                    rs!terminal = G_TERMINAL
                    rs!jenis_urusan = G_JENIS_URUSAN
                    rs!cawangan = G_CAWANGAN
                
                    rs.Update
                End If
                
                rs.Close
                Set rs = Nothing
                
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm28.L5_Text & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    If Frm84.L76_Text <> vbNullString Then 'Jumlah perolehan mata
                        If IsNumeric(Frm84.L76_Text) Then Frm84_LM_MATA_DAPAT = Frm84.L76_Text
                    End If
                    If Frm84.TB36 <> vbNullString Then 'Jumlah mata yang ditebus
                        If IsNumeric(Frm84.TB36) Then Frm84_LM_MATA_TEBUS = Frm84.TB36
                    End If
                    If Frm84.L77_Text <> vbNullString Then 'Jumlah mata asal
                        If IsNumeric(Frm84.L77_Text) Then Frm84_LM_MATA_ASAL = Frm84.L77_Text
                    End If
                    rs!baki_point = Frm84_LM_MATA_ASAL - Frm84_LM_MATA_TEBUS + Frm84_LM_MATA_DAPAT
                    rs!write_timestamp2 = LM_NOW
                    rs!no_staff = Frm84_LM_EMP_NO 'No. Pekerja
                    rs!terminal = G_TERMINAL
                    rs!jenis_urusan = G_JENIS_URUSAN
                    rs.Update
                End If
                
                rs.Close
                Set rs = Nothing
                
            End If
            
        Else
        
'@ Jika data yang diedit aadalah dijual kepada bukan ahli kedai (Jualan asal kepada ahli)

            If Frm84_LM_FLAG_MATA_ASAL = 1 Then '0 : Tiada perolehan mata , 1 : Ada perolehan mata (Pada awal pembelian asal)

                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 71_tebus_agih_point where no_invoice='" & Frm84.L3_Text & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs.EOF Then
                    rs!Status = 0
                    rs!write_timestamp3 = LM_NOW
                    rs!terminal = G_TERMINAL
                    rs!jenis_urusan = G_JENIS_URUSAN
                    rs.Update
                End If
                
                rs.Close
                Set rs = Nothing
            
            End If

        End If
'### Maklumat agihan point ### - End
        
        '### Update Maklumat Trade In ### - Start
        If Frm84_LM_Flag_TRADE_IN = 1 Then '0 : Tiada Urusan Trade In , 1 : Ada Urusan Trade In
        
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            rs.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & Frm84.L57_Text & "'", cn, adOpenKeyset, adLockOptimistic

            If Not rs.EOF Then

                G_ID = rs!ID
                Call recovery_16_gold_bar_belian
                
                rs!trade_in_status = 1
                rs!no_staff = Frm84_LM_EMP_NO 'No. Pekerja
                rs!terminal = G_TERMINAL
                rs!write_timestamp2 = LM_NOW
                rs!jenis_urusan = G_JENIS_URUSAN
                rs!remarks = "Ubah flag trade in selepas edit data jualan [Ada trade in]"
                rs.Update
                
            End If
            
            rs.Close
            Set rs = Nothing
            
        End If
        '### Update Maklumat Trade In ### - End
                        
'### Masukkan Data Jualan Ke Dalam Table Jualan ### - Start
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from " & G_JUALAN_TEMP & " where (status = 3 OR status = 4 OR status = 5)", cn, adOpenKeyset, adLockOptimistic
        
        While rs.EOF = False
        
            Frm84_LM_BERAT_ASAL = 0
            Frm84_LM_BEZA_BERAT = 0
            Frm84_BERAT_JUALAN_BARU = 0
            
            If rs!Status = "3" Then
            
'### Kemasukkan Data Baru ### - Start
                Set rs1 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs1.Open "select * from 23_senarai_jualan where id is null", cn, adOpenKeyset, adLockOptimistic
                
                If rs1.EOF Then
                    rs1.AddNew
                    rs1!tarikh = Frm84.DTPicker1 'Tarikh Jualan
                    rs1!no_resit = Frm84.L3_Text 'No. Resit Jualan
                    If Not IsNull(rs!flag_barang) Then
                        rs1!flag_barang = rs!flag_barang
                    Else
                        rs1!flag_barang = Null
                    End If
                    If Not IsNull(rs!nama_purity) Then
                        rs1!nama_purity = rs!nama_purity
                    Else
                        rs1!nama_purity = Null
                    End If
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
                        rs1!berat_jualan = rs!berat_jualan 'Berat Jualan (g)
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
                    If Not IsNull(rs!kadar_komisyen_upah) Then 'Kadar komisyen bagi upah kepada agen dropship
                        rs1!kadar_komisyen_upah = rs!kadar_komisyen_upah
                    Else
                        rs1!kadar_komisyen_upah = Null
                    End If
                    If Not IsNull(rs!komisyen_upah) Then 'Jumlah komisyen bagi upah kepada agen dropship
                        rs1!komisyen_upah = Format(rs!komisyen_upah, "0.00")
                    Else
                        rs1!komisyen_upah = Null
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
                        rs1!untung2 = Format(rs!untung2, "0.00") 'Untung jika restok pada harga supplier ini
                    Else
                        rs1!untung2 = Null 'Untung jika restok pada harga supplier ini
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
                    If Not IsNull(rs!status_jualan) Then
                        rs1!Status = rs!status_jualan '0 : Tiada Potong , 1 : Ada Potong
                    Else
                        rs1!Status = Null '0 : Tiada Potong , 1 : Ada Potong
                    End If
                    If Not IsNull(rs!Type) Then
                        rs1!Type = rs!Type '0 : BK , 1 : Barang Permata
                    Else
                        rs1!Type = Null '0 : BK , 1 : Barang Permata
                    End If
                    rs1!no_pekerja = Frm84_LM_EMP_NO 'No. Pekerja
                    If Frm84.L28_Text <> vbNullString Then
                        If Frm28.L5_Text <> vbNullString Then
                            rs1!no_rujukan_pembeli = Frm28.L5_Text 'No. Rujukan Pembeli
                        Else
                            rs1!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
                        End If
                    Else
                        rs1!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
                    End If
                    If Frm84.CB7 = 1 Then
                        If Frm27.L5_Text <> vbNullString Then
                            rs1!no_rujukan_agen_dropship = Frm27.L5_Text 'No. Rujukan Agen Dropship
                        Else
                            rs1!no_rujukan_agen_dropship = Null 'No. Rujukan Agen Dropship
                        End If
                    Else
                        rs1!no_rujukan_agen_dropship = Null 'No. Rujukan Agen Dropship
                    End If
                
'2:  Member
'3:  RAF
'4:  Pengedar
'5:  Normal Dealer
'6:  Master Dealer
    
                    If Frm84.CB4 = 1 Then
                        rs1!kategori_pembeli = 1
                    ElseIf Frm84.CB5 = 1 Then
                        rs1!kategori_pembeli = 2
                    ElseIf Frm84.CB6 = 1 Then
                        rs1!kategori_pembeli = 4
                    ElseIf Frm84.CB9 = 1 Then
                        rs1!kategori_pembeli = 3
                    ElseIf Frm84.CB10 = 1 Then
                        rs1!kategori_pembeli = 5
                    'ElseIf Frm84.CB11 = 1 Then
                    '    rs1!kategori_pembeli = 6
                    End If
                    
                    If Frm84.CB13 = 0 Then
                        rs1!bil_rasmi = 1
                    Else
                        rs1!bil_rasmi = 0
                    End If
                    
                    If Not IsNull(rs!gst_include) Then 'Pilihan Cukai GST (SR) Samada Pelanggan Bayar Atau Kedai Bayar
                        rs1!gst_include = rs!gst_include
                    Else
                        rs1!gst_include = Null
                    End If
                    If Not IsNull(rs!harga_tanpa_gst) Then
                        rs1!harga_tanpa_gst = Format(rs!harga_tanpa_gst, "0.00") 'Harga Semasa (RM/g)
                    Else
                        rs1!harga_tanpa_gst = Null 'Harga Semasa (RM/g)
                    End If
                    
                    If Not IsNull(rs!modal_tanpa_gst) Then
                        rs1!modal_tanpa_gst = Format(rs!modal_tanpa_gst, "0.00")
                    Else
                        rs1!modal_tanpa_gst = Null
                    End If
                    If Not IsNull(rs!harga_per_gram_tanpa_gst) Then
                        rs1!harga_per_gram_tanpa_gst = Format(rs!harga_per_gram_tanpa_gst, "0.00")
                    Else
                        rs1!harga_per_gram_tanpa_gst = Null
                    End If
                    If Not IsNull(rs!jualan_per_gram_dengan_gst) Then
                        rs1!jualan_per_gram_dengan_gst = Format(rs!jualan_per_gram_dengan_gst, "0.00")
                    Else
                        rs1!jualan_per_gram_dengan_gst = Null
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
                    rs1!jenis_jualan = 0 '0 : Jualan biasa kepada pelanggan , 1 : Jualan secara tukaran barang kepada agen
                    If Not IsNull(rs!gst_barang_atau_upah) Then '0 : GST pada harga jualan , 1 : GST pada upah
                        rs1!gst_barang_atau_upah = rs!gst_barang_atau_upah
                    Else
                        rs1!gst_barang_atau_upah = 0
                    End If
                    If Not IsNull(rs!harga_jualan_dengan_gst) Then
                        rs1!harga_jualan_dengan_gst = Format(rs!harga_jualan_dengan_gst, "0.00")
                    Else
                        rs1!harga_jualan_dengan_gst = 0
                    End If
                    If Frm84.CB27 = 1 Then
                        rs1!jualan_online = 1
                    Else
                        rs1!jualan_online = 0
                    End If
                    rs1!status_rekod = 1
                    If Not IsNull(rs!jualan_per_gram) Then
                        rs1!jualan_per_gram = Format(rs!jualan_per_gram, "0.00")
                    Else
                        rs1!jualan_per_gram = 0
                    End If
                    If Not IsNull(rs!modal_per_gram) Then
                        rs1!modal_per_gram = Format(rs!modal_per_gram, "0.00")
                    Else
                        rs1!modal_per_gram = 0
                    End If
                    If Not IsNull(rs!flag_upah) Then
                        rs1!flag_upah = rs!flag_upah
                    Else
                        rs1!flag_upah = 1
                    End If
                    If Not IsNull(rs!upah_per_gram) Then
                        rs1!upah_per_gram = Format(rs!upah_per_gram, "0.00")
                    Else
                        rs1!upah_per_gram = Null
                    End If
                    rs1!status_r = LM_STATUS_R
                    rs1!cawangan = LM_CAWANGAN
                    
                    If Not IsNull(rs!harga_jual_excl_gst) Then
                        rs1!harga_jual_excl_gst = Format(rs!harga_jual_excl_gst, "0.00")
                    Else
                        rs1!harga_jual_excl_gst = Null
                    End If
                    If Not IsNull(rs!harga_modal_gst) Then
                        rs1!harga_modal_gst = Format(rs!harga_modal_gst, "0.00")
                    Else
                        rs1!harga_modal_gst = Null
                    End If
                    If Not IsNull(rs!harga_modal_incl_gst) Then
                        rs1!harga_modal_incl_gst = Format(rs!harga_modal_incl_gst, "0.00")
                    Else
                        rs1!harga_modal_incl_gst = Null
                    End If
                    If Not IsNull(rs!harga_modal_excl_gst) Then
                        rs1!harga_modal_excl_gst = Format(rs!harga_modal_excl_gst, "0.00")
                    Else
                        rs1!harga_modal_excl_gst = Null
                    End If
                    If Not IsNull(rs!baru_or_ti) Then
                        rs1!baru_or_ti = rs!baru_or_ti
                    Else
                        rs1!baru_or_ti = Null
                    End If
                    rs1!nama_pekerja = Frm84_LM_EMP_NAMA
                    rs1!write_timestamp = LM_NOW
                    rs1.Update
                End If
                
                rs1.Close
                Set rs1 = Nothing

'### Update Table Database Bagi Item Ini ### - Start
                Set rs2 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs2.Open "select * from Data_Database where no_siri_produk='" & rs!no_siri_Produk & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs2.EOF Then
                
                    G_ID = rs2!ID
                    Call recovery_data_database

                    If rs!Type = 0 Then
                        If Not IsNull(rs2!Berat) Then Frm84_LM_BERAT_ASAL = rs2!Berat 'Berat Asal (g)
                        If Not IsNull(rs!berat_jualan) Then Frm84_BERAT_JUALAN_BARU = rs!berat_jualan 'Berat Jualan (g)
                        If Not IsNull(rs!berat_jualan) Then Frm84_LM_BERAT_JUALAN = rs!berat_jualan 'Berat Jualan (g)
                        
                        If Format(Frm84_LM_BERAT_JUALAN, "0.00") = Format(Frm84_LM_BERAT_ASAL, "0.00") Then
                            rs2!beza_berat = "0.00" 'Baki Berat
                            'rs2!susut_berat = "0.00" 'Susut berat
                            rs2!StatusItem = 11
                            rs2!tarikh_jualan1 = Null
                            rs2!nama_pekerja_potong = Null
                        Else
                            'rs2!beza_berat = Format(Frm84_LM_BERAT_JUALAN - Frm84_LM_BERAT_ASAL, "0.00") 'Baki Berat
                            rs2!beza_berat = Format(Frm84_LM_BERAT_ASAL - Frm84_LM_BERAT_JUALAN, "0.00") 'Baki Berat
                            'rs2!susut_berat = "0.00" 'Susut berat
                            rs2!StatusItem = 12
                            rs2!tarikh_jualan1 = Frm84.DTPicker1
                            rs2!nama_pekerja_potong = Frm84_LM_EMP_NAMA
                        End If
                    Else
                        rs2!StatusItem = 11
                    End If
                    
                    rs2!write_timestamp2 = LM_NOW
                    rs2!no_pekerja = Frm84_LM_EMP_NO
                    rs2!terminal = G_TERMINAL
                    'rs2!cawangan = LM_CAWANGAN
                    rs2!Menu = 1
                
                    rs2.Update
                End If
                
                rs2.Close
                Set rs2 = Nothing
'### Update Table Database Bagi Item Ini ### - End

'### Kemasukkan Data Baru ### - End

'### Edit Data Sedia Ada ### - Start

            ElseIf rs!Status = "4" Then
            
                '### Masukkan data lama ke dalam table #72_data_amendment ### - Start
                If rs!Type = 0 Then
                
                    Set rs3 = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    
                    strsql = "insert into 72_data_amendment(id_asal,no_siri_produk,kategori_produk,berat,upah,nama_supplier,jenis_barang,jenis,menu)" & _
                                "select ID,no_siri_produk,kategori_produk,Berat_Jualan,upah,harga_Semasa,0,0,1 from 23_senarai_jualan WHERE id='" & rs!id_database & "'"
                    
                    Set rs3 = cn.Execute(strsql)
                    Set rs3 = Nothing
                    
                ElseIf rs!Type = 1 Then
                
                    Set rs3 = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    
                    strsql = "insert into 72_data_amendment(id_asal,no_siri_produk,kategori_produk,upah,jenis_barang,jenis,menu)" & _
                                "select ID,no_siri_produk,kategori_produk,harga_jualan,1,0,1 from 23_senarai_jualan WHERE id='" & rs!id_database & "'"
                    
                    Set rs3 = cn.Execute(strsql)
                    Set rs3 = Nothing
                
                End If
                '### Masukkan data lama ke dalam table #72_data_amendment ### - End
            
                Set rs1 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs1.Open "select * from 23_senarai_jualan where ID='" & rs!id_database & "'", cn, adOpenKeyset, adLockOptimistic
            
                If Not rs1.EOF Then

                    G_ID = rs1!ID
                    Call recovery_23_senarai_jualan
                    
                    rs1!tarikh = Frm84.DTPicker1 'Tarikh Jualan
                    rs1!no_resit = Frm84.L3_Text 'No. Resit Jualan
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
                    If Not IsNull(rs!flag_barang) Then
                        rs1!flag_barang = rs!flag_barang
                    Else
                        rs1!flag_barang = Null
                    End If
                    If Not IsNull(rs!nama_purity) Then
                        rs1!nama_purity = rs!nama_purity
                    Else
                        rs1!nama_purity = Null
                    End If
                    If Not IsNull(rs!Berat_Asal) Then
                        rs1!Berat_Asal = rs!Berat_Asal 'Berat Asal (g)
                    Else
                        rs1!Berat_Asal = Null 'Berat Asal (g)
                    End If
                    
                    If Not IsNull(rs1!berat_jualan) Then
                        If IsNumeric(rs1!berat_jualan) Then Frm84_LM_BERAT_JUALAN_ASAL = Format(rs1!berat_jualan, "0.00")
                    End If
                    
                    If Not IsNull(rs!berat_jualan) Then
                        rs1!berat_jualan = rs!berat_jualan 'Berat Jualan (g)
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
                    If Not IsNull(rs!kadar_komisyen_upah) Then 'Kadar komisyen bagi upah kepada agen dropship
                        rs1!kadar_komisyen_upah = rs!kadar_komisyen_upah
                    Else
                        rs1!kadar_komisyen_upah = Null
                    End If
                    If Not IsNull(rs!komisyen_upah) Then 'Jumlah komisyen bagi upah kepada agen dropship
                        rs1!komisyen_upah = Format(rs!komisyen_upah, "0.00")
                    Else
                        rs1!komisyen_upah = Null
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
                        rs1!untung2 = Format(rs!untung2, "0.00") 'Untung jika restok pada harga supplier ini
                    Else
                        rs1!untung2 = Null 'Untung jika restok pada harga supplier ini
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
                    If Not IsNull(rs!status_jualan) Then
                        rs1!Status = rs!status_jualan '0 : Tiada Potong , 1 : Ada Potong
                    Else
                        rs1!Status = Null '0 : Tiada Potong , 1 : Ada Potong
                    End If
                    
                    If Not IsNull(rs!Type) Then
                        rs1!Type = rs!Type '0 : BK , 1 : Barang Permata
                    Else
                        rs1!Type = Null '0 : BK , 1 : Barang Permata
                    End If
                    rs1!no_pekerja = Frm84_LM_EMP_NO 'No. Pekerja

                    If Frm84.L28_Text <> vbNullString Then
                        If Frm28.L5_Text <> vbNullString Then
                            rs1!no_rujukan_pembeli = Frm28.L5_Text 'No. Rujukan Pembeli
                        Else
                            rs1!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
                        End If
                    Else
                        rs1!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
                    End If
                    If Frm84.CB7 = 1 Then
                        If Frm27.L5_Text <> vbNullString Then
                            rs1!no_rujukan_agen_dropship = Frm27.L5_Text 'No. Rujukan Agen Dropship
                        Else
                            rs1!no_rujukan_agen_dropship = Null 'No. Rujukan Agen Dropship
                        End If
                    Else
                        rs1!no_rujukan_agen_dropship = Null 'No. Rujukan Agen Dropship
                    End If
'2:  Member
'3:  RAF
'4:  Pengedar
'5:  Normal Dealer
'6:  Master Dealer

                    If Frm84.CB4 = 1 Then
                        rs1!kategori_pembeli = 1
                    ElseIf Frm84.CB5 = 1 Then
                        rs1!kategori_pembeli = 2
                    ElseIf Frm84.CB6 = 1 Then
                        rs1!kategori_pembeli = 4
                    ElseIf Frm84.CB9 = 1 Then
                        rs1!kategori_pembeli = 3
                    ElseIf Frm84.CB10 = 1 Then
                        rs1!kategori_pembeli = 5
                    'ElseIf Frm84.CB11 = 1 Then
                    '    rs1!kategori_pembeli = 6
                    End If
                    
                    If Frm84.CB13 = 0 Then
                        rs1!bil_rasmi = 1
                    Else
                        rs1!bil_rasmi = 0
                    End If
                
                    If Not IsNull(rs!gst_include) Then 'Pilihan Cukai GST (SR) Samada Pelanggan Bayar Atau Kedai Bayar
                        rs1!gst_include = rs!gst_include
                    Else
                        rs1!gst_include = Null
                    End If
                    If Not IsNull(rs!harga_tanpa_gst) Then
                        rs1!harga_tanpa_gst = Format(rs!harga_tanpa_gst, "0.00") 'Harga Semasa (RM/g)
                    Else
                        rs1!harga_tanpa_gst = Null 'Harga Semasa (RM/g)
                    End If
                    If Not IsNull(rs!modal_tanpa_gst) Then
                        rs1!modal_tanpa_gst = Format(rs!modal_tanpa_gst, "0.00")
                    Else
                        rs1!modal_tanpa_gst = Null
                    End If
                    If Not IsNull(rs!harga_per_gram_tanpa_gst) Then
                        rs1!harga_per_gram_tanpa_gst = Format(rs!harga_per_gram_tanpa_gst, "0.00")
                    Else
                        rs1!harga_per_gram_tanpa_gst = Null
                    End If
                    If Not IsNull(rs!jualan_per_gram_dengan_gst) Then
                        rs1!jualan_per_gram_dengan_gst = Format(rs!jualan_per_gram_dengan_gst, "0.00")
                    Else
                        rs1!jualan_per_gram_dengan_gst = Null
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
                    
                    rs1!jenis_jualan = 0 '0 : Jualan biasa kepada pelanggan , 1 : Jualan secara tukaran barang kepada agen
                    If Not IsNull(rs!gst_barang_atau_upah) Then '0 : GST pada harga jualan , 1 : GST pada upah
                        rs1!gst_barang_atau_upah = rs!gst_barang_atau_upah
                    Else
                        rs1!gst_barang_atau_upah = 0
                    End If
                    If Not IsNull(rs!harga_jualan_dengan_gst) Then
                        rs1!harga_jualan_dengan_gst = Format(rs!harga_jualan_dengan_gst, "0.00")
                    Else
                        rs1!harga_jualan_dengan_gst = 0
                    End If
                    If Frm84.CB27 = 1 Then
                        rs1!jualan_online = 1
                    Else
                        rs1!jualan_online = 0
                    End If
                    rs1!status_rekod = 1
                    If Not IsNull(rs!jualan_per_gram) Then
                        rs1!jualan_per_gram = Format(rs!jualan_per_gram, "0.00")
                    Else
                        rs1!jualan_per_gram = 0
                    End If
                    If Not IsNull(rs!modal_per_gram) Then
                        rs1!modal_per_gram = Format(rs!modal_per_gram, "0.00")
                    Else
                        rs1!modal_per_gram = 0
                    End If
                    If Not IsNull(rs!flag_upah) Then
                        rs1!flag_upah = rs!flag_upah
                    Else
                        rs1!flag_upah = 1
                    End If
                    If Not IsNull(rs!upah_per_gram) Then
                        rs1!upah_per_gram = Format(rs!upah_per_gram, "0.00")
                    Else
                        rs1!upah_per_gram = Null
                    End If
                    rs1!no_staff = G_LOGIN_USER
                    rs1!write_timestamp2 = LM_NOW
                    'rs1!cawangan = G_CAWANGAN
                    If Not IsNull(rs!harga_jual_excl_gst) Then
                        rs1!harga_jual_excl_gst = Format(rs!harga_jual_excl_gst, "0.00")
                    Else
                        rs1!harga_jual_excl_gst = Null
                    End If
                    If Not IsNull(rs!harga_modal_gst) Then
                        rs1!harga_modal_gst = Format(rs!harga_modal_gst, "0.00")
                    Else
                        rs1!harga_modal_gst = Null
                    End If
                    If Not IsNull(rs!harga_modal_incl_gst) Then
                        rs1!harga_modal_incl_gst = Format(rs!harga_modal_incl_gst, "0.00")
                    Else
                        rs1!harga_modal_incl_gst = Null
                    End If
                    If Not IsNull(rs!harga_modal_excl_gst) Then
                        rs1!harga_modal_excl_gst = Format(rs!harga_modal_excl_gst, "0.00")
                    Else
                        rs1!harga_modal_excl_gst = Null
                    End If
                    If Not IsNull(rs!baru_or_ti) Then
                        rs1!baru_or_ti = rs!baru_or_ti
                    Else
                        rs1!baru_or_ti = Null
                    End If
                    rs1!nama_pekerja = Frm84_LM_EMP_NAMA
                    rs1.Update
                    
                    rs1.Close
                    Set rs1 = Nothing
                
'### Update Table Database Bagi Item Ini ### - Start
                    Set rs2 = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs2.Open "select * from Data_Database where no_siri_produk='" & rs!no_siri_Produk & "'", cn, adOpenKeyset, adLockOptimistic
                    
                    If Not rs2.EOF Then

                        G_ID = rs2!ID
                        Call recovery_data_database
                
                        If rs!Type = 0 Then
                            If Not IsNull(rs2!Berat) Then Frm84_LM_BERAT_ASAL = Format(rs2!Berat, "0.00") 'Berat Asal (g)
                            If Not IsNull(rs2!beza_berat) Then Frm84_LM_BEZA_BERAT = Format(rs2!beza_berat, "0.00") 'Berat Asal (g)
                            If Not IsNull(rs!berat_jualan) Then Frm84_BERAT_JUALAN_BARU = Format(rs!berat_jualan, "0.00") 'Berat Jualan (g)
                            If Not IsNull(rs2!susut_berat) Then Frm84_SUSUT_BERAT = Format(rs2!susut_berat, "0.00") 'Susut berat
                            
                            Frm84_LM_BAKI_BERAT = Frm84_BERAT_JUALAN_BARU - Format((Frm84_LM_BERAT_JUALAN_ASAL + Frm84_LM_BEZA_BERAT), "0.00") - Frm84_SUSUT_BERAT
                            
                            If Frm84_LM_BAKI_BERAT = 0 Then
                                rs2!beza_berat = "0.00" 'Baki Berat
                                rs2!StatusItem = 11
                                rs2!tarikh_jualan1 = Null
                                rs2!nama_pekerja_potong = Null
                            Else
                                rs2!beza_berat = Format(Frm84_LM_BERAT_ASAL - Frm84_BERAT_JUALAN_BARU - Frm84_SUSUT_BERAT, "0.00") 'Baki Berat
                                rs2!StatusItem = 12
                                rs2!tarikh_jualan1 = Frm84.DTPicker1
                                rs2!nama_pekerja_potong = Frm84_LM_EMP_NAMA
                            End If
                        Else
                            rs2!StatusItem = 11
                        End If

                        rs2!write_timestamp2 = LM_NOW
                        rs2!no_pekerja = Frm84_LM_EMP_NO
                        rs2!terminal = G_TERMINAL
                        rs2!Menu = 1
                    
                        rs2.Update
                        
                    End If
                    
                    rs2.Close
                    Set rs2 = Nothing
'### Update Table Database Bagi Item Ini ### - End

                End If

                If rs!Type = 0 Then
                
                    Set rs3 = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
                    strsql = "UPDATE 72_data_amendment set no_siri_produk_new='" & rs!no_siri_Produk & "'," _
                    & "kategori_produk_new='" & rs!kategori_Produk & "'," _
                    & "berat_new='" & rs!berat_jualan & "'," _
                    & "upah_new='" & rs!UPAH & "'," _
                    & "nama_supplier_new='" & rs!harga_Semasa & "'," _
                    & "nama_pic='" & Frm84_LM_EMP_NAMA & "'," _
                    & "terminal='" & G_TERMINAL & "'," _
                    & "write_timestamp='" & LM_NOW & "'" _
                    & "WHERE id_asal='" & rs!id_database & "'"

                    Set rs3 = cn.Execute(strsql)
                    Set rs3 = Nothing
                    
                ElseIf rs!Type = 1 Then

                    Set rs3 = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        
                    strsql = "UPDATE 72_data_amendment set no_siri_produk_new='" & rs!no_siri_Produk & "'," _
                    & "kategori_produk_new='" & rs!kategori_Produk & "'," _
                    & "upah_new='" & rs!harga_jualan & "'," _
                    & "nama_pic='" & Frm84_LM_EMP_NAMA & "'," _
                    & "terminal='" & G_TERMINAL & "'," _
                    & "write_timestamp='" & LM_NOW & "'" _
                    & "WHERE id_asal='" & rs!id_database & "'"
                    
                    Set rs3 = cn.Execute(strsql)
                    Set rs3 = Nothing
                    
                End If
'### Edit Data Sedia Ada ### - End

            ElseIf rs!Status = "5" Then
            
                '### Masukkan data lama ke dalam table #72_data_amendment ### - Start
                If rs!Type = 0 Then
                
                    Set rs3 = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    
                    strsql = "insert into 72_data_amendment(id_asal,no_siri_produk,kategori_produk,berat,upah,nama_supplier,jenis_barang,jenis,menu,write_timestamp)" & _
                                "select ID,no_siri_produk,kategori_produk,Berat_Jualan,upah,harga_Semasa,0,1,1,'" & LM_NOW & "' from 23_senarai_jualan WHERE id='" & rs!id_database & "'"
                    
                    Set rs3 = cn.Execute(strsql)
                    Set rs3 = Nothing
                    
                ElseIf rs!Type = 1 Then
                
                    Set rs3 = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    
                    strsql = "insert into 72_data_amendment(id_asal,no_siri_produk,kategori_produk,upah,jenis_barang,jenis,menu,write_timestamp)" & _
                                "select ID,no_siri_produk,kategori_produk,harga_jualan,1,1,1,'" & LM_NOW & "' from 23_senarai_jualan WHERE id='" & rs!id_database & "'"
                    
                    Set rs3 = cn.Execute(strsql)
                    Set rs3 = Nothing
                
                End If
                '### Masukkan data lama ke dalam table #72_data_amendment ### - End

                Set rs1 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs1.Open "select * from 23_senarai_jualan where ID='" & rs!id_database & "'", cn, adOpenKeyset, adLockOptimistic
            
                If Not rs1.EOF Then
                    
                    G_ID = rs1!ID
                    Call recovery_23_senarai_jualan
                
                    If Not IsNull(rs1!berat_jualan) Then
                        Frm84_LM_BERAT_RETURN = rs1!berat_jualan
                    End If
                    rs1!no_staff = G_LOGIN_USER
                    rs1!write_timestamp2 = LM_NOW
                    rs1!status_rekod = 0
                    
                    rs1.Update
                End If
                
                rs1.Close
                Set rs1 = Nothing
                
'### Update Table Database Bagi Item Ini ### - Start
                Set rs2 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs2.Open "select * from Data_Database where no_siri_produk='" & rs!no_siri_Produk & "'", cn, adOpenKeyset, adLockOptimistic
                
                If Not rs2.EOF Then
                
                    G_ID = rs2!ID
                    Call recovery_data_database
                
                    If rs!Type = 0 Then
                    
                        If Not IsNull(rs2!Berat) Then Frm84_LM_BERAT_ASAL = Format(rs2!Berat, "0.00") 'Berat Asal (g)
                        If Not IsNull(rs2!beza_berat) Then Frm84_LM_BEZA_BERAT = Format(rs2!beza_berat, "0.00") 'Berat Asal (g)
                        If Not IsNull(rs2!susut_berat) Then Frm84_SUSUT_BERAT = Format(rs2!susut_berat, "0.00") 'Susut berat
                        
                        Frm84_LM_BERAT_ASAL_COMP = Format(Frm84_LM_BERAT_ASAL, "0.00")
                        Frm84_LM_BERAT_SELEPAS_COMP = Format(Frm84_LM_BERAT_RETURN + Frm84_LM_BEZA_BERAT - Frm84_SUSUT_BERAT, "0.00")

                        If Frm84_LM_BERAT_ASAL_COMP = Frm84_LM_BERAT_SELEPAS_COMP Then
                            rs2!beza_berat = Format(Frm84_LM_BERAT_RETURN + Frm84_LM_BEZA_BERAT, "0.00")  'Baki Berat
                            rs2!StatusItem = 10
                            rs2!tarikh_jualan1 = Null
                            rs2!nama_pekerja_potong = Null
                        Else
                            rs2!beza_berat = Format(Frm84_LM_BEZA_BERAT + Frm84_LM_BERAT_RETURN, "0.00") 'Baki Berat
                            rs2!StatusItem = 12
                            rs2!tarikh_jualan1 = Frm84.DTPicker1
                            'rs2!beza_berat = Format(Frm84_LM_BERAT_ASAL - Frm84_LM_BERAT_RETURN - Frm84_SUSUT_BERAT, "0.00") 'Baki Berat
                        End If
                                
                    Else
                        rs2!StatusItem = 10
                    End If
                    
                    rs2!write_timestamp2 = LM_NOW
                    rs2!no_pekerja = Frm84_LM_EMP_NO
                    rs2!terminal = G_TERMINAL
                    rs2!Menu = 1
                    
                    rs2.Update
                End If
                
                rs2.Close
                Set rs2 = Nothing
'### Update Table Database Bagi Item Ini ### - End
            
            End If

            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing

'### Masukkan Data Jualan Ke Dalam Table Jualan ### - End
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 23_senarai_jualan where no_resit='" & Frm84.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
        While rs.EOF = False
        
            rs!tarikh = Frm84.DTPicker1 'Tarikh Jualan
            rs!no_pekerja = Frm84_LM_EMP_NO 'No. Pekerja
            If Frm84.L28_Text <> vbNullString Then
                If Frm28.L5_Text <> vbNullString Then
                    rs!no_rujukan_pembeli = Frm28.L5_Text 'No. Rujukan Pembeli
                Else
                    rs!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
                End If
            Else
                rs!no_rujukan_pembeli = Null 'No. Rujukan Pembeli
            End If
            If Frm84.CB7 = 1 Then
                If Frm27.L5_Text <> vbNullString Then
                    rs!no_rujukan_agen_dropship = Frm27.L5_Text 'No. Rujukan Agen Dropship
                Else
                    rs!no_rujukan_agen_dropship = Null 'No. Rujukan Agen Dropship
                End If
            Else
                rs!no_rujukan_agen_dropship = Null 'No. Rujukan Agen Dropship
            End If
            If Frm84.CB27 = 1 Then
                rs!jualan_online = 1
            Else
                rs!jualan_online = 0
            End If
            
            rs.Update
            rs.MoveNext
            
        Wend
        
        rs.Close
        Set rs = Nothing

        If Frm84_LM_JENIS_TRADE_IN = 3 Then
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    
            strsql = "UPDATE 93_trade_in_susut_niai set status = 0 , write_timestamp2='" & LM_NOW & "' WHERE no_invoice='" & Frm84.L3_Text & "' AND status = 1"
    
            Set rs = cn.Execute(strsql)
            Set rs = Nothing
        End If
        
        If G_TI_MODE = 3 Then
        'masukkan senarai trade in
   
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            
            If G_TI_MEMORY(0, 0) = 0 Then strsql = "insert into 93_trade_in_susut_niai(no_invoice,tarikh,berat,harga_semasa,harga,status,write_timestamp,jenis,terminal,nama_pekerja) values ('" & Frm84.L3_Text & "','" & Frm84.DTPicker1 & "','" & G_TI_MEMORY(1, 1) & "','" & G_TI_MEMORY(1, 2) & "','" & G_TI_MEMORY(1, 3) & "',1,'" & LM_NOW & "',0,'" & G_TERMINAL & "','" & Frm84_LM_EMP_NAMA & "') ,('" & Frm84.L3_Text & "','" & Frm84.DTPicker1 & "','" & G_TI_MEMORY(2, 1) & "','" & G_TI_MEMORY(2, 2) & "','" & G_TI_MEMORY(2, 3) & "',1,'" & LM_NOW & "',1,'" & G_TERMINAL & "','" & Frm84_LM_EMP_NAMA & "'),('" & Frm84.L3_Text & "','" & Frm84.DTPicker1 & "','" & G_TI_MEMORY(3, 1) & "','" & G_TI_MEMORY(3, 2) & "','" & G_TI_MEMORY(3, 3) & "',1,'" & LM_NOW & "',2,'" & G_TERMINAL & "','" & Frm84_LM_EMP_NAMA & "')"
            If G_TI_MEMORY(0, 0) = 1 Then strsql = "insert into 93_trade_in_susut_niai(no_invoice,tarikh,berat,harga_semasa,harga,status,write_timestamp,jenis,terminal,nama_pekerja) values ('" & Frm84.L3_Text & "','" & Frm84.DTPicker1 & "','" & G_TI_MEMORY(1, 1) & "','" & G_TI_MEMORY(1, 2) & "','" & G_TI_MEMORY(1, 3) & "',1,'" & LM_NOW & "',0,'" & G_TERMINAL & "','" & Frm84_LM_EMP_NAMA & "') ,('" & Frm84.L3_Text & "','" & Frm84.DTPicker1 & "','" & G_TI_MEMORY(2, 1) & "','" & G_TI_MEMORY(2, 2) & "','" & G_TI_MEMORY(2, 3) & "',1,'" & LM_NOW & "',2,'" & G_TERMINAL & "','" & Frm84_LM_EMP_NAMA & "')"

            Set rs = cn.Execute(strsql)
            Set rs = Nothing
            
        End If

        If Frm84_LM_JENIS_TRADE_IN = 1 Then '1 : Trade in (Voucher) , 2 : Belian dengan trade in
        
            If Frm84.L56_Text = 2 Then '1 : Trade in (Voucher) , 2 : Belian dengan trade in
                Call Frm84_penerimaan_barang_trade_in
            End If

        End If
        If Frm84_LM_JENIS_TRADE_IN = 2 Then '1 : Trade in (Voucher) , 2 : Belian dengan trade in
        
            'tukar status barang trade in kepada 0
            If Frm84.L56_Text = 1 Then '1 : Trade in (Voucher) , 2 : Belian dengan trade in
                '### Update Maklumat Trade In ### - Start
                If Frm84_LM_Flag_TRADE_IN = 1 Then '0 : Tiada Urusan Trade In , 1 : Ada Urusan Trade In
                
                    Set rs = New ADODB.Recordset
                    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                    rs.Open "select * from 16_gold_bar_belian where no_resit_trade_in='" & Frm84_LM_No_VOUCHER_TI & "'", cn, adOpenKeyset, adLockOptimistic
        
                    If Not rs.EOF Then
                    
                        G_ID = rs!ID
                        Call recovery_16_gold_bar_belian
                    
                        rs!trade_in_status = 0
                        rs!no_staff = Frm84_LM_EMP_NO 'No. Pekerja
                        rs!terminal = G_TERMINAL
                        rs!write_timestamp2 = LM_NOW
                        rs!jenis_urusan = G_JENIS_URUSAN
                        rs!remarks = "Pulangkan status trade in kerana tukar jualan dengan trade in 2 kepada jualan dengan trade in 1"

                        rs.Update
                    End If
                    
                    rs.Close
                    Set rs = Nothing
                    
                End If
                '### Update Maklumat Trade In ### - End
            End If
            
            If Frm84.L56_Text = 2 Then '1 : Trade in (Voucher) , 2 : Belian dengan trade in
                Call Frm84_save_edit_data_TI
            End If
        
        End If

        If DATA_SAVE = 1 Then
        
'#### Update Log Aktiviti Sistem #### - Start
            user = MDI_frm1.L3_Text
            LogAct_Memory = "[" & G_LOGIN_USER & "] Edit jualan barang kemas. No. Invoice [" & Frm84.L3_Text & "]."
            LogDate_Memory = LM_NOW
            Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End

            Call amendment_email_check
            
            If G_SPKE_ME_MAIL = "YES" Then
            
                Set rs = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs.Open "select * from 88_sales_notfication where id is null", cn, adOpenKeyset, adLockOptimistic
                
                If rs.EOF Then
                    rs.AddNew
                    rs!no_invoice_asal = Frm84.L3_Text 'No. invoice rasmi
                    rs!jenis = 1
                    rs!jenis_report = 0 '0 : Jualan , 1 : Trade In
                    rs!write_timestamp = LM_NOW
                    rs!terminal = G_TERMINAL
                    rs!Status = 0
                    rs.Update
                End If
                
                rs.Close
                Set rs = Nothing

                Shell "cmd.exe /c " & G_SPKE_NE_PATH
                
            End If

            Note = "Data telah berjaya disimpan." & vbCrLf & _
                    "Sistem akan refresh data."

            Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
            
            If Answer = vbNo Or Answer = vbYes Then
            
                GM_NEXT_PREV = 2
                
                If Frm101.L33_Text = 0 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                    If Frm101.CB3 = 1 Then 'Report Jualan
                        Call Frm85_Header_Report_Jualan
                        Call Frm85_Report_Jualan_page
                    End If
                ElseIf Frm101.L33_Text = 2 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                    Call Frm85_Header_Report_Jualan
                    Call Frm85_carian_jualan_page
                ElseIf Frm101.L33_Text = 5 Then '0 : Carian Report Biasa ,  1 : Carian Ikut Berat , 2 : Carian Ikut No. Resit Jualan , 3 : Carian Ikut No. Resit Buyback / Trade In , 4 : Carian mengikut No. Invoice Supplier , 5 : Carian mengikut No. Siri Produk (Belian BK) , 6 : Carian mengikut No. Siri Produk (Buyback BK) , 7 : Carian mengikut No. Siri Produk (Belian GB) , 8 : Carian mengikut No. Siri Produk (Buyback GB)
                    Call Frm85_Header_Report_Jualan
                    Call Frm85_Report_Jualan_barcode
                End If

                Frm85.Show
                Unload Frm84
                Unload Frm26
                Unload Frm27
                Unload Frm28
                Unload Frm83
                MDI_frm1.L5_Text = 12
                
            End If
            
            MsgBox "Data Jualan Telah Berjaya Disimpan.", vbInformation, "Info"
        End If
    End If
End If
End Sub
Sub tesuto4()
'On Error Resume Next
Dim Err(5)
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset

x = 0
Y = 0 '0 : Tiada Perubahan Pada Data , 1 : Ada Perubahan Pada Data
DATA_SAVE = 0

G_JENIS_URUSAN = 1

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

    G_ID = rs!ID
    Call recovery_16_gold_bar_belian

    rs!tarikh = Frm84.DTPicker1 'Tarikh Belian
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
    rs!remarks = "Edit data jualan stok baru dari trade in 2"
    rs!Status = 1
    
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

strsql = "insert into Data_Database(NoRujukanSistem,tarikh_belian,bill_No_Belian,terminal,supplier_ID,nama_Supplier,Kod_Supplier,purity_ID,purity,kod_Purity,kategori_produk_ID,kategori_Produk,Kod_Kategori_Produk,Berat,beza_berat,UPAH,Upah30,riyal,kos_Belian_Gram,kos_Belian_Item,SpreadValue,harga_lepas_spread,adjustment,kos_item_tanpa_tax,cara_Belian,dimension_Panjang,dimension_Lebar,dimension_Saiz,code1,code2,harga_Per_Gram_Item,dulang,no_cert,gst_barang_atau_upah,StatusItem,Upah_Jualan,Upah_Member,Upah_RAF,Upah_Pengedar,code_Supplier,HargaJualan_Member,HargaJualan_Pengedar,upah_normal_dealer,upah_master_dealer,HargaJualan_RAF,hargajualan_normal_dealer,hargajualan_master_dealer,remarks,gst_ari_nashi,kadar_gst,jumlah_gst,harga_item,receiving_Status,harga_tanpa_gst,gst_included,jenis_trade_in,flag_upah,upah_per_gram,flag_image,write_timestamp,no_id_gst,susut_berat,no_pekerja)" & _
            "select '" & Format(Frm83_LM_No_RUJUKAN_BELIAN, "000000") & "',tarikh_belian,bill_No_Belian,terminal,supplier_ID,nama_Supplier,Kod_Supplier,purity_ID,purity,kod_Purity,kategori_produk_ID,kategori_Produk,Kod_Kategori_Produk,Berat,beza_berat,UPAH,Upah30,riyal,kos_Belian_Gram,kos_Belian_Item,Spread,harga_lepas_spread,adjustment,kos_item_tanpa_tax,cara_Belian,dimension_Panjang,dimension_Lebar,dimension_Saiz,code1,code2,harga_Per_Gram_Item,dulang,no_cert,gst_barang_atau_upah,10,Upah_Jualan,Upah_Member,Upah_RAF,Upah_Pengedar,code_Supplier,HargaJualan_Member,HargaJualan_Pengedar,upah_normal_dealer,upah_master_dealer,HargaJualan_RAF,hargajualan_normal_dealer,hargajualan_master_dealer,remarks,gst_ari_nashi,kadar_gst,jumlah_gst,harga_item,jenis,harga_tanpa_gst,gst_included,jenis_trade_in,flag_upah,upah_per_gram,flag_image,Now(),no_id_gst,0.00,'" & Frm84_LM_EMP_NO & "' from " & G_BELIAN_TEMP & " WHERE StatusItem='" & 3 & "'"
        
Set rs = cn.Execute(strsql)
Set rs = Nothing
'### Masukkan maklumat data barang ke dalam table #data_database ### - End

'### Update data barang ke dalam table #data_database ### - Start
'Barang sedia ada
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "UPDATE data_database," & G_BELIAN_TEMP & " SET Data_Database.NoRujukanSistem='" & Format(Frm83_LM_No_RUJUKAN_BELIAN, "000000") & "'," _
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
& "Data_Database.no_pekerja='" & Frm84_LM_EMP_NO & "'," _
& "Data_Database.menu='" & G_JENIS_URUSAN & "'," _
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
        & "Data_Database.no_pekerja='" & Frm84_LM_EMP_NO & "'," _
        & "Data_Database.menu='" & G_JENIS_URUSAN & "'," _
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
    
    LM_ID = rs!ID
    
    'If LM_ID > 562 Then
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
    'Else
    '    rs!no_siri_Produk = rs!Kod_Kategori_Produk & rs!Barcode
    'End If
    
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

    user = MDI_frm1.L3_Text
    LogAct_Memory = "[" & G_LOGIN_USER & "] Edit data trade in [" & Frm84.L57_Text & "]."
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

