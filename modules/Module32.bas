Attribute VB_Name = "Module32"
Sub Frm54_ClearAllField()
'On Error Resume Next
Frm54.TB1 = "0.00" 'Harga MKS
Frm54.TB12 = "0.00" 'Harga Dari Supplier
Frm54.TB2 = "0.00" 'Pemalar Pelanggan
Frm54.TB3 = "0.00" 'Pemalar Member
Frm54.TB4 = "0.00" 'Pemalar VVIP
Frm54.TB5 = "0.00" 'Pemalar Pengedar
Frm54.TB6 = "0.00" 'Harga jualan kepada pekerja kedai
Frm54.TB30 = "0.00" 'Pemalar N.Dealer
Frm54.TB31 = "0.00" 'Pemalar M.Dealer

Frm54.TB20 = "0.00" 'Kenaikan Harga Pelanggan
Frm54.TB21 = "0.00" 'Kenaikan Harga Member
Frm54.TB22 = "0.00" 'Kenaikan Harga Pengedar
Frm54.TB23 = "0.00" 'Kenaikan Harga RAF
Frm54.TB24 = "0.00" 'Kenaikan Harga Normal Dealer
Frm54.TB25 = "0.00" 'Kenaikan Harga Master Dealer

Frm54.L3_Text = "0.00" 'Harga Pelanggan
Frm54.L4_Text = "0.00" 'Harga Member
Frm54.L5_Text = "0.00" 'Harga RAF
Frm54.L6_Text = "0.00" 'Harga Pengedar
Frm54.L13_Text = "0.00" 'Harga N.Dealer
Frm54.L14_Text = "0.00" 'Harga M.Dealer
End Sub
Sub frm54_initial_location()
'On Error Resume Next
Frm54.Frame2.Top = 360
Frm54.Frame1.Top = 360
Frm54.Frame3.Top = 360
Frm54.Pic5.Top = 360
Frm54.Frame2.Left = 120
Frm54.Frame1.Left = 120
Frm54.Frame3.Left = 120
Frm54.Pic5.Left = 120

Frm54.Frame1.Visible = False
Frm54.Frame2.Visible = False
Frm54.Frame3.Visible = False
End Sub
Sub Frm54_call_setting_upah()
'On Error Resume Next
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 73_tetapan_upah where default_setting='" & MDI_frm1.L20_Text & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    
    If Not IsNull(rs!pelanggan) Then Frm54.TB20 = Format(rs!pelanggan, "0.00") 'Kenaikan upah bagi pelanggan
    If Not IsNull(rs!Member) Then Frm54.TB21 = Format(rs!Member, "0.00") 'Kenaikan upah bagi Member
    If Not IsNull(rs!Pengedar) Then Frm54.TB22 = Format(rs!Pengedar, "0.00") 'Kenaikan upah bagi Pengedar
    If Not IsNull(rs!raf) Then Frm54.TB23 = Format(rs!raf, "0.00") 'Kenaikan upah bagi RAF
    If Not IsNull(rs!normal_dealer) Then Frm54.TB24 = Format(rs!normal_dealer, "0.00") 'Kenaikan upah bagi Normal Dealer
    If Not IsNull(rs!master_dealer) Then Frm54.TB25 = Format(rs!master_dealer, "0.00") 'Kenaikan upah bagi Master Dealer
    
End If

rs.Close
Set rs = Nothing
End Sub
Sub frm54_senarai_harga_header()
'on error resume next
With Frm54.LV1
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear

    Frm54.LV1.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 800, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "Cawangan", 3000
    .ColumnHeaders.Add 5, , "Purity", 1700, 2
    .ColumnHeaders.Add 6, , "Pelanggan", 2100, 1
    .ColumnHeaders.Add 7, , "Ahli", 2100, 1
    .ColumnHeaders.Add 8, , "Silver", 2100, 1
    .ColumnHeaders.Add 9, , "Gold", 2100, 1
    .ColumnHeaders.Add 10, , "Platinum", 2100, 1
    .ColumnHeaders.Add 11, , "Update", 4500

End With
End Sub
Sub frm54_senarai_harga()
'on error resume next
Dim Frm54_LM_TOTAL_PAGE As Double

Frm54_PAGE_SIZE = 20
Frm54_LM_TOTAL_PAGE = 0
x = 0

re_gen_report:

LM_START_ROW = Frm54.L63_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm54_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm54.L64_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm54_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        Frm54.L61_Text = 1
    End If
End If

If MDI_frm1.L20_Text = "Semua cawangan" Then

    Frm54_SEARCH_3 = Null
    Frm54_SEARCH_3_LOGIC = "<>"
    
Else

    Frm54_SEARCH_3 = MDI_frm1.L20_Text
    Frm54_SEARCH_3_LOGIC = "="
    
End If

Frm54_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from hargaemas where cawangan " & Frm54_SEARCH_3_LOGIC & "'" & Frm54_SEARCH_3 & "' order by purity ASC , ID ASC LIMIT " & LM_START_ROW & "," & Frm54_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1
    If Frm54_LM_PAGE_FOUND = 0 Then
        If Frm54.L64_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm54.L61_Text = Frm54.L61_Text + 1 'Paparan Page ke-xxx
                Frm54_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm54.L61_Text) Then
                    If Frm54.L61_Text <> 1 Then
                        Frm54.L61_Text = Frm54.L61_Text - 1 'Paparan Page ke-xxx
                        Frm54_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If

    Y = ((Frm54.L61_Text - 1) * Frm54_PAGE_SIZE) + x

    With Frm54.LV1.ListItems.Add(, , rs!ID)
     
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!cawangan) Then 'Cawangan
            .ListSubItems.Add , , rs!cawangan
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!purity) Then 'Purity
            .ListSubItems.Add , , rs!purity
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!Harga_Pelanggan) Then
            .ListSubItems.Add , , Format(rs!Harga_Pelanggan, "#,##0.00")
        Else
            .ListSubItems.Add , , "0.00"
        End If
        
        If Not IsNull(rs!Harga_Member) Then
            .ListSubItems.Add , , Format(rs!Harga_Member, "#,##0.00")
        Else
            .ListSubItems.Add , , "0.00"
        End If
        
        If Not IsNull(rs!Harga_Pengedar) Then
            .ListSubItems.Add , , Format(rs!Harga_Pengedar, "#,##0.00")
        Else
            .ListSubItems.Add , , "0.00"
        End If
        
        If Not IsNull(rs!Harga_RAF) Then
            .ListSubItems.Add , , Format(rs!Harga_RAF, "#,##0.00")
        Else
            .ListSubItems.Add , , "0.00"
        End If
        
        If Not IsNull(rs!harga_nd) Then
            .ListSubItems.Add , , Format(rs!harga_nd, "#,##0.00")
        Else
            .ListSubItems.Add , , "0.00"
        End If
        
        If Not IsNull(rs!write_timestamp) Then .ListSubItems.Add , , rs!write_timestamp
        
    End With
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from hargaemas where cawangan " & Frm54_SEARCH_3_LOGIC & "'" & Frm54_SEARCH_3 & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    Frm54_LM_TOTAL_PAGE = Format(rs(0) / Frm54_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm54_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm54_LM_PAGE = Split(Frm54_LM_TOTAL_PAGE, ".")(0)
        Frm54_LM_PAGE_LEBIHAN = Split(Frm54_LM_TOTAL_PAGE, ".")(1)
        
        If Frm54_LM_PAGE_LEBIHAN <> "00" Then
            Frm54.L62_Text = Frm54_LM_PAGE + 1
        Else
            Frm54.L62_Text = Frm54_LM_PAGE
        End If
        
    Else
    
        Frm54.L62_Text = Frm54_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm54.L62_Text = 0
    End If
Else
    Frm54.L62_Text = 0
End If

rs.Close
Set rs = Nothing

If x <> 0 Then
    Frm54.L63_Text = LM_START_ROW
End If

If Frm54.L61_Text <> vbNullString And IsNumeric(Frm54.L61_Text) Then
    If Frm54.L62_Text <> vbNullString And IsNumeric(Frm54.L62_Text) Then
        Frm54_LM_CURR_PAGE = Frm54.L61_Text
        Frm54_LM_TOTAL_PAGE = Frm54.L62_Text
        
        If Frm54_LM_CURR_PAGE > Frm54_LM_TOTAL_PAGE Then
            
            Frm54.L61_Text = Frm54.L61_Text - 1
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            GoTo re_gen_report:
            
        End If
    End If
End If
End Sub
Sub TetapanHargaJualan()
'On Error Resume Next
Dim a As Double
Dim b As Double
Dim ab As Single
'Dim aa As Double
Dim ad As Double

'If IsNumeric(Frm54.TB2) Then
'    a = Frm54.TB1 'Harga MKS
'    b = Frm54.TB2 'Pemalar Harga Pelanggan
'    aa = Format(a * b, "0.00")
'    ac = Len(aa)
'    If InStr(1, aa, ".") <> 0 Then
'        ae = InStr(1, aa, ".")
'        ab = Right(aa, ac - ae)
'        ad = Left(aa, ae - 1)
'    End If
    
'    If 99 >= ab And ab > 70 Then
'        Frm54.L3_Text = Format(ad + 1, "0.00")
'    End If
'    If 30 < ab And ab <= 70 Then
'        Frm54.L3_Text = Format(ad + 0.5, "0.00")
'    End If
'    If 30 >= ab And ab > 0 Then
'        Frm54.L3_Text = Format(ad, "0.00")
'    End If
'Else
'    Frm54.L3_Text = "XXX.XX"
'End If

If IsNumeric(Frm54.TB1) And IsNumeric(Frm54.TB2) Then
    a = Frm54.TB1
    b = Frm54.TB2
    Frm54.L3_Text = Format((a - b), "0.00") 'Harga Jualan Bagi Pelanggan
Else
    Frm54.L3_Text = "XXX.XX"
End If
Call TetapanHargaJualan1
End Sub
Sub TetapanHargaJualan1()
'On Error Resume Next
Dim a As Double
Dim b As Double
'If IsNumeric(Frm54.TB3) And IsNumeric(Frm54.L3_Text) Then
'    Frm54.L4_Text = Format(Frm54.L3_Text - Frm54.TB3, "0.00")
'Else
'    Frm54.L4_Text = "XXX.XX"
'End If

If IsNumeric(Frm54.TB1) And IsNumeric(Frm54.TB3) Then
    a = Frm54.TB1
    b = Frm54.TB3
    Frm54.L4_Text = Format((a - b), "0.00") 'Harga Jualan Bagi Member
Else
    Frm54.L4_Text = "XXX.XX"
End If

Call TetapanHargaJualan2
End Sub
Sub TetapanHargaJualan2()
'On Error Resume Next
Dim a As Double
Dim b As Double
Dim ab As Single
'Dim aa As Double
Dim ad As Double

'If IsNumeric(Frm54.TB4) Then
'    a = Frm54.TB1 'Harga MKS
'    b = Frm54.TB4 'Pemalar Harga VVIP
'    aa = Format(a * b, "0.00")
'    ac = Len(aa)
'    If InStr(1, aa, ".") <> 0 Then
'        ae = InStr(1, aa, ".")
'        ab = Right(aa, 1)
'        ad = Left(aa, ac - 1)
'    End If
    
'    Frm54.L5_Text = Format(ad, "0.00")
'
'    If 5 <= ab And ab <= 9 Then
'        Frm54.L5_Text = Format(ad + 0.1, "0.00")
'    End If
    'If 0 < ab And ab < 5 Then
    '    Frm54.L5_Text = Format(ad - (ab / 100), "0.00")
    'End If
'Else
'    Frm54.L5_Text = "XXX.XX"
'End If

If IsNumeric(Frm54.TB1) And IsNumeric(Frm54.TB5) Then
    a = Frm54.TB1
    b = Frm54.TB5
    Frm54.L6_Text = Format((a - b), "0.00") 'Harga Jualan Bagi Pengedar
Else
    Frm54.L6_Text = "XXX.XX"
End If
End Sub
Sub TetapanHargaJualan6()
'On Error Resume Next
Dim a As Double
Dim b As Double
Dim ab As Single
'Dim aa As Double
Dim ad As Double

'If IsNumeric(Frm54.TB8) Then
'    a = Frm54.TB1 'Harga MKS
'    b = Frm54.TB8 'Pemalar Harga Ansuran : Pelanggan
'    aa = Format(a + b, "0.00")
'    ac = Len(aa)
'    If InStr(1, aa, ".") <> 0 Then
'        ae = InStr(1, aa, ".")
'        ab = Right(aa, 1)
'        ad = Left(aa, ac - 1)
'        ae = Right(ad, 1)
'    End If
    
'    Frm54.L9_Text = Format(ad, "0.00")
'
'    If 5 < ae And ae <= 9 Then
 '       Frm54.L9_Text = Format(ad - ae + 10, "0.00")
'    End If
'    If 0 < ae And ae < 5 Then
'        Frm54.L9_Text = Format(ad - ae, "0.00")
'    End If
'Else
'    Frm54.L9_Text = "XXX.XX"
'End If
If IsNumeric(Frm54.TB1) And IsNumeric(Frm54.TB4) Then
    a = Frm54.TB1
    b = Frm54.TB4
    Frm54.L5_Text = Format((a - b), "0.00") 'Harga Jualan Ansuran Bagi Pengedar
Else
    Frm54.L5_Text = "XXX.XX"
End If

'Call TetapanHargaJualan7
End Sub
Sub Frm54_rekod_tetapan_harga()
'On Error Resume Next
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 66_rekod_tetapan_harga", cn, adOpenKeyset, adLockOptimistic

rs.AddNew
If Frm54.CBB1 <> vbNullString Then 'Purity
    rs!purity = Frm54.CBB1
Else
    rs!purity = Null
End If
If Frm54.TB1 <> vbNullString Then 'Harga MKS
    rs!harga_jualan = Format(Frm54.TB1, "0.00")
Else
    rs!harga_jualan = Null
End If
If Frm54.TB12 <> vbNullString Then 'Harga Dari Supplier
    rs!harga_supplier = Format(Frm54.TB12, "0.00")
Else
    rs!harga_supplier = Null
End If
If Frm54.L3_Text <> vbNullString Then 'Harga Pelanggan
    rs!Harga_Pelanggan = Format(Frm54.L3_Text, "0.00")
Else
    rs!Harga_Pelanggan = Null
End If
If Frm54.L4_Text <> vbNullString Then 'Harga Member
    rs!Harga_Member = Format(Frm54.L4_Text, "0.00")
Else
    rs!Harga_Member = Null
End If
If Frm54.L5_Text <> vbNullString Then 'Harga RAF
    rs!Harga_RAF = Format(Frm54.L5_Text, "0.00")
Else
    rs!Harga_RAF = Null
End If
If Frm54.L6_Text <> vbNullString Then 'Harga Pengedar
    rs!Harga_Pengedar = Format(Frm54.L6_Text, "0.00")
Else
    rs!Harga_Pengedar = Null
End If
If Frm54.L13_Text <> vbNullString Then 'Harga ND
    rs!harga_normal_dealer = Format(Frm54.L13_Text, "0.00")
Else
    rs!harga_normal_dealer = Null
End If
If Frm54.L14_Text <> vbNullString Then 'Harga MD
    rs!harga_master_dealer = Format(Frm54.L14_Text, "0.00")
Else
    rs!harga_master_dealer = Null
End If
If MDI_frm1.L3_Text <> vbNullString Then
    rs!pic_name = MDI_frm1.L3_Text
Else
    rs!pic_name = Null
End If
rs!write_timestamp = Now
rs.Update
    
rs.Close
Set rs = Nothing
End Sub
