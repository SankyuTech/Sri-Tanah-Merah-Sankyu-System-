Attribute VB_Name = "Module24"
Sub frm118_initial_setting()
'On Error Resume Next
frm118.TB7 = vbNullString
frm118.TB8 = vbNullString

frm118.TB7.BackColor = &H8000000A
frm118.TB7.Locked = True

frm118.CB1 = 1
frm118.CB2 = 0
frm118.CB3 = 1
frm118.CB4 = 0
frm118.CB5 = 0
frm118.CB6 = 1
frm118.CB7 = 0

frm118.CB1.Enabled = True
frm118.CB2.Enabled = True

frm118.DTPicker1 = DateTime.Date

frm118.TB1 = vbNullString
frm118.TB2 = "0.00"
frm118.TB3 = "0.00"
frm118.TB4 = "0.00"
frm118.TB5 = "0.00"
frm118.TB6 = "0.00"

frm118.L1_Text = vbNullString
frm118.L2_Text = "6.00"
frm118.L3_Text = "0.00"

frm118.CMD1.Visible = True
frm118.CMD2.Visible = False
frm118.CMD3.Visible = False

frm118.CBB1.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where status='" & 1 & "' order by supplier ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!supplier) Then frm118.CBB1.AddItem rs!supplier
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

frm118.L2_Text = G_RATE_GST
'Set rs = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from default_setting where Default1='" & "Default" & "'", cn, adOpenKeyset, adLockOptimistic

'If Not rs.EOF Then

'    GLOBAL_DISABLE = 1

'    If Not IsNull(rs!gst_value) Then frm118.L2_Text = rs!gst_value 'Jumlah Kadar GST

'    GLOBAL_DISABLE = 0

'End If

'rs.Close
'Set rs = Nothing

'###Senarai Nama Pekerja###
frm118.CBB2.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where status='" & "Aktif" & "' AND InvestorSmall = 0 AND InvestorBig = 0", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    If rs!user_level = 1 Or rs!user_level = 2 Or rs!user_level = 3 Or rs!user_level = G_LEVEL_USER Then frm118.CBB2.AddItem rs!Samaran & "  |  " & rs!NoPekerja
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Call Frm118_jurujual
End Sub
Sub frm118_calc_gst_1()
'On Error Resume Next
Dim frm118_LM_KADAR_GST As Double
Dim frm118_LM_UPAH As Double

frm118_LM_KADAR_GST = 0
frm118_LM_UPAH = 0

If IsNumeric(frm118.L2_Text) Then frm118_LM_KADAR_GST = frm118.L2_Text 'Kadar gst (%)
If IsNumeric(frm118.TB2) Then frm118_LM_UPAH = frm118.TB2 'Harga (RM)

If frm118.L2_Text <> vbNullString And IsNumeric(frm118.L2_Text) Then

    If frm118.TB2 <> vbNullString And IsNumeric(frm118.TB2) Then
        
        If frm118.CB6 = 1 Then
        
            frm118.L3_Text = Format(frm118_LM_UPAH, "#,##0.00") 'Harga upah tanpa GST
            frm118.TB3 = Format(frm118_LM_UPAH * (frm118_LM_KADAR_GST / 100), "#,##0.00") 'Jumlah GST (Bagi jualan setiap item)
            
        End If
        
        If frm118.CB7 = 1 Then
    
            frm118.L3_Text = Format(frm118_LM_UPAH / (1 + (frm118_LM_KADAR_GST / 100)), "#,##0.00") 'Harga upah tanpa GST
            frm118.TB3 = Format(frm118_LM_UPAH - (frm118_LM_UPAH / (1 + (frm118_LM_KADAR_GST / 100))), "#,##0.00") 'Jumlah GST (Bagi jualan setiap item)
                
        End If

    Else
    
        frm118.L3_Text = Format(0, "#,##0.00") 'Harga upah tanpa GST
        frm118.TB3 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
        
    End If

Else

    If IsNumeric(frm118.TB2) Then
    
        frm118.L3_Text = Format(frm118.TB2, "#,##0.00") 'Harga upah tanpa GST
        frm118.TB3 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
        
    Else
        
        frm118.L3_Text = Format(0, "#,##0.00") 'Harga upah tanpa GST
        frm118.TB3 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
        
    End If
    
End If
End Sub
Sub frm118_calc_gst_2()
'On Error Resume Next
Dim frm118_LM_HARGA_TANPA_GST As Double
Dim frm118_LM_GST As Double

frm118_LM_HARGA_TANPA_GST = 0
frm118_LM_GST = 0

If IsNumeric(frm118.L3_Text) Then frm118_LM_HARGA_TANPA_GST = frm118.L3_Text 'Harga tanpa GST
If IsNumeric(frm118.TB3) Then frm118_LM_GST = frm118.TB3 'GST (RM)

If (frm118.L3_Text <> vbNullString And IsNumeric(frm118.L3_Text)) And (frm118.TB3 <> vbNullString And IsNumeric(frm118.TB3)) Then
    
    frm118.TB4 = Format(frm118_LM_HARGA_TANPA_GST + frm118_LM_GST, "#,##0.00")

Else

    frm118.TB4 = "0.00"
    
End If
End Sub
Sub frm118_calc_gst_3()
'On Error Resume Next
Dim frm118_LM_HARGA_GST_SR As Double
Dim frm118_LM_HARGA_GST_ZR As Double

frm118_LM_HARGA_GST_SR = 0
frm118_LM_HARGA_GST_ZR = 0

If IsNumeric(frm118.TB4) Then frm118_LM_HARGA_GST_SR = frm118.TB4 'Harga tanpa GST
If IsNumeric(frm118.TB5) Then frm118_LM_HARGA_GST_ZR = frm118.TB5 'GST (RM)

If (frm118.TB4 <> vbNullString And IsNumeric(frm118.TB4)) And (frm118.TB5 <> vbNullString And IsNumeric(frm118.TB5)) Then
    
    frm118.TB6 = Format(frm118_LM_HARGA_GST_SR + frm118_LM_HARGA_GST_ZR, "#,##0.00")

Else

    frm118.TB6 = "0.00"
    
End If
End Sub
Sub Frm118_jurujual()
'On Error Resume Next
If MDI_frm1.L3_Text <> vbNullString Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        frm118.CBB2 = rs!Samaran & "  |  " & rs!NoPekerja
        
    Else
        
        frm118.CBB2.AddItem "" & "  |  " & rs!Samaran
        frm118.CBB2 = "" & "  |  " & rs!Samaran
        
    End If
    
    rs.Close
    Set rs = Nothing
    
    If G_LOCK_JURUJUAL = "YES" Then
    
        frm118.CBB2.Enabled = False
        frm118.CBB2.BackColor = &H8000000A

    Else
    
        frm118.CBB2.Enabled = True
        frm118.CBB2.BackColor = &HFFFFFF

    End If

End If
End Sub
Sub frm118_kiraan_berat_bayaran()
'On Error Resume Next
Dim frm118_LM_HARGA As Double
Dim frm118_LM_HARGA_SEMASA As Double

frm118_LM_HARGA = 0
frm118_LM_HARGA_SEMASA = 0

If frm118.TB6 <> vbNullString And IsNumeric(frm118.TB6) Then frm118_LM_HARGA = frm118.TB6 'Jumlah bayaran
If frm118.TB7 <> vbNullString And IsNumeric(frm118.TB7) Then frm118_LM_HARGA_SEMASA = frm118.TB7 'Harga semasa

If frm118_LM_HARGA_SEMASA <> 0 Then
    frm118.TB8 = Format(frm118_LM_HARGA / frm118_LM_HARGA_SEMASA, "#,##0.00")
Else
    frm118.TB8 = "0.00"
End If
End Sub
Sub frm118_save_data_expenses()
'### Update Akaun Bagi Expense ### - Start
Dim LM_JUMLAH_ALL As Double
Dim LM_JUMLAH_GST As Double
Dim LM_HARGA_DAN_GST As Double

LM_JUMLAH_ALL = 0
LM_JUMLAH_GST = 0
LM_HARGA_DAN_GST = 0

If frm118.TB6 <> vbNullString And IsNumeric(frm118.TB6) Then LM_JUMLAH_ALL = frm118.TB6
If frm118.TB3 <> vbNullString And IsNumeric(frm118.TB3) Then LM_JUMLAH_GST = frm118.TB3
If frm118.TB4 <> vbNullString And IsNumeric(frm118.TB4) Then LM_HARGA_DAN_GST = frm118.TB4

LM_NOW = Now

LM_ID_GST = vbNullString

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where Supplier='" & frm118.CBB1 & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!no_id_gst) Then LM_ID_GST = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
        
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 39_akaun_expense", cn, adOpenKeyset, adLockOptimistic

rs.AddNew
rs!no_rujukan_expense = G_No_RESIT_JUALAN
If frm118.CBB1 <> vbNullString Then 'Nama Kedai / Supplier
    rs!nama_kedai = UCase(frm118.CBB1)
Else
    rs!nama_kedai = Null
End If
If frm118.TB1 <> vbNullString Then 'No. Invoice
    rs!no_resit = UCase(frm118.TB1)
Else
    rs!no_resit = Null
End If
rs!tujuan = "Belian stok emas" 'Tujuan
If LM_ID_GST <> vbNullString Then
    rs!no_id_gst = LM_ID_GST 'No. ID GST
Else
    rs!no_id_gst = Null
End If
rs!tarikh = frm118.DTPicker1 'Tarikh
rs!jumlah_tanpa_gst = Format(LM_JUMLAH_ALL - LM_JUMLAH_GST, "0.00") 'Jumlah Tanpa GST (RM)
If frm118.TB6 <> vbNullString Then 'Jumlah Dengan GST (RM)
    rs!harga_dengan_gst = Format(frm118.TB6, "0.00")
Else
    rs!harga_dengan_gst = Null
End If
If frm118.TB5 <> vbNullString Then 'Harga Keseluruhan Bagi Barang ZR
    rs!gst_zr_harga = Format(frm118.TB5, "0.00")
Else
    rs!gst_zr_harga = Null
End If
rs!gst_zr_cukai = Format(0, "0.00") 'Jumlah Cukai Bagi ZR
rs!gst_sr_harga = Format(LM_HARGA_DAN_GST - LM_JUMLAH_GST, "0.00") 'Harga Keseluruhan Bagi Barang SR
If frm118.TB3 <> vbNullString Then 'Jumlah Cukai Bagi SR
    rs!gst_sr_cukai = Format(frm118.TB3, "0.00")
Else
    rs!gst_sr_cukai = Null
End If
If frm118.L2_Text <> vbNullString Then '% Cukai GST
    rs!gst_value = frm118.L2_Text
Else
    rs!gst_value = Null
End If
If frm118.CBB2 <> vbNullString Then
    frm118_LM_EMP_NO = Split(frm118.CBB2, "  |  ")(1)
    rs!no_pekerja = frm118_LM_EMP_NO 'No. Pekerja
End If
rs!write_timestamp = LM_NOW
rs!terminal = G_TERMINAL
rs!Menu = 2
rs!Status = 1
If frm118.CB3 = 1 Then
    rs!cara_bayaran = 0
ElseIf frm118.CB4 = 1 Then
    rs!cara_bayaran = 1
ElseIf frm118.CB5 = 1 Then
    rs!cara_bayaran = 2
End If
rs!cawangan = G_KEDAI
rs.Update

rs.Close
Set rs = Nothing
'### Update Akaun Bagi Expense ### - End
End Sub
Sub frm118_save_data_expenses_edit()
'### Update Akaun Bagi Expense ### - Start
Dim LM_JUMLAH_ALL As Double
Dim LM_JUMLAH_GST As Double
Dim LM_HARGA_DAN_GST As Double

LM_JUMLAH_ALL = 0
LM_JUMLAH_GST = 0
LM_HARGA_DAN_GST = 0

If frm118.TB6 <> vbNullString And IsNumeric(frm118.TB6) Then LM_JUMLAH_ALL = frm118.TB6
If frm118.TB3 <> vbNullString And IsNumeric(frm118.TB3) Then LM_JUMLAH_GST = frm118.TB3
If frm118.TB4 <> vbNullString And IsNumeric(frm118.TB4) Then LM_HARGA_DAN_GST = frm118.TB4

LM_NOW = Now

LM_ID_GST = vbNullString

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where Supplier='" & frm118.CBB1 & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!no_id_gst) Then LM_ID_GST = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
        
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 39_akaun_expense where no_rujukan_expense='" & G_No_RESIT_JUALAN & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    rs!no_rujukan_expense = G_No_RESIT_JUALAN
    If frm118.CBB1 <> vbNullString Then 'Nama Kedai / Supplier
        rs!nama_kedai = UCase(frm118.CBB1)
    Else
        rs!nama_kedai = Null
    End If
    If frm118.TB1 <> vbNullString Then 'No. Invoice
        rs!no_resit = UCase(frm118.TB1)
    Else
        rs!no_resit = Null
    End If
    rs!tujuan = "Belian stok emas" 'Tujuan
    If LM_ID_GST <> vbNullString Then
        rs!no_id_gst = LM_ID_GST 'No. ID GST
    Else
        rs!no_id_gst = Null
    End If
    rs!tarikh = frm118.DTPicker1 'Tarikh
    rs!jumlah_tanpa_gst = Format(LM_JUMLAH_ALL - LM_JUMLAH_GST, "0.00") 'Jumlah Tanpa GST (RM)
    If frm118.TB6 <> vbNullString Then 'Jumlah Dengan GST (RM)
        rs!harga_dengan_gst = Format(frm118.TB6, "0.00")
    Else
        rs!harga_dengan_gst = Null
    End If
    If frm118.TB5 <> vbNullString Then 'Harga Keseluruhan Bagi Barang ZR
        rs!gst_zr_harga = Format(frm118.TB5, "0.00")
    Else
        rs!gst_zr_harga = Null
    End If
    rs!gst_zr_cukai = Format(0, "0.00") 'Jumlah Cukai Bagi ZR
    rs!gst_sr_harga = Format(LM_HARGA_DAN_GST - LM_JUMLAH_GST, "0.00") 'Harga Keseluruhan Bagi Barang SR
    If frm118.TB3 <> vbNullString Then 'Jumlah Cukai Bagi SR
        rs!gst_sr_cukai = Format(frm118.TB3, "0.00")
    Else
        rs!gst_sr_cukai = Null
    End If
    If frm118.L2_Text <> vbNullString Then '% Cukai GST
        rs!gst_value = frm118.L2_Text
    Else
        rs!gst_value = Null
    End If
    If frm118.CBB2 <> vbNullString Then
        frm118_LM_EMP_NO = Split(frm118.CBB2, "  |  ")(1)
        rs!no_pekerja = frm118_LM_EMP_NO 'No. Pekerja
    End If
    rs!write_timestamp = LM_NOW
    rs!terminal = G_TERMINAL
    rs!Menu = 2
    rs!Status = 1
    If frm118.CB3 = 1 Then
        rs!cara_bayaran = 0
    ElseIf frm118.CB4 = 1 Then
        rs!cara_bayaran = 1
    ElseIf frm118.CB5 = 1 Then
        rs!cara_bayaran = 2
    End If
    rs.Update

End If

rs.Close
Set rs = Nothing
'### Update Akaun Bagi Expense ### - End
End Sub
