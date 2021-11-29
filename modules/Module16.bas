Attribute VB_Name = "mod_cash_in_out"
Sub frm100_initial_setting()
'on error resume next
Frm100.Pic1.Left = 120
Frm100.Pic1.Top = 360
Frm100.Pic2.Left = 120
Frm100.Pic2.Top = 360
Frm100.Pic3.Left = 120
Frm100.Pic3.Top = 360

Frm100.TB1 = vbNullString
Frm100.TB2 = vbNullString
Frm100.TB3 = vbNullString
Frm100.TB4 = vbNullString
Frm100.TB5 = vbNullString
Frm100.CBB1.Clear
Frm100.DTPicker1 = DateTime.Date$
Frm100.DTPicker2 = DateTime.Date$
Frm100.DTPicker3 = DateTime.Date$

Frm100.Pic1.Visible = False
Frm100.Pic2.Visible = False
Frm100.Pic3.Visible = False

Frm100.CMD1.Visible = True
Frm100.CMD2.Visible = False
Frm100.CMD3.Visible = False

Frm100.L13_Text = vbNullString

Frm100.CB1 = 1
Frm100.CB2 = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where status='" & "Aktif" & "' AND InvestorSmall = 0 AND InvestorBig = 0", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    If rs!user_level = 1 Or rs!user_level = 2 Or rs!user_level = 3 Or rs!user_level = G_LEVEL_USER Then Frm100.CBB1.AddItem rs!Samaran & "  |  " & rs!NoPekerja
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Call Frm100_jurujual
End Sub
Sub frm100_initial_setting2()
'on error resume next
Frm100.TB1 = vbNullString
Frm100.TB2 = vbNullString
Frm100.CBB1.Clear
Frm100.DTPicker1 = DateTime.Date$

Frm100.CB1 = 1
Frm100.CB2 = 0

'###Senarai Nama Pekerja###
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where status='" & "Aktif" & "' AND InvestorSmall = 0 AND InvestorBig = 0", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    If rs!user_level = 1 Or rs!user_level = 2 Or rs!user_level = 3 Or rs!user_level = G_LEVEL_USER Then Frm100.CBB1.AddItem rs!Samaran & "  |  " & rs!NoPekerja
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
End Sub
Sub frm100_cash_in_out_header()
'on error resume next
'#### Header Report #### - Start
Frm100.MSFlexGrid1.Clear
Frm100.MSFlexGrid1.Rows = 1
Frm100.MSFlexGrid1.RowHeight(0) = 700
Frm100.MSFlexGrid1.FormatString = "No.|<No.|<No. ID|<Tarikh|<Jenis|<Jumlah (RM)|<Nama Pekerja"

Frm100.MSFlexGrid1.ColWidth(0) = 0 'No.
Frm100.MSFlexGrid1.ColWidth(1) = 600 'No.
Frm100.MSFlexGrid1.ColWidth(2) = 0 'No. ID
Frm100.MSFlexGrid1.ColWidth(3) = 2500 'Tarikh
Frm100.MSFlexGrid1.ColWidth(4) = 3000 'Jenis
Frm100.MSFlexGrid1.ColWidth(5) = 2400 'Jumlah (RM)
Frm100.MSFlexGrid1.ColWidth(6) = 2400 'Nama Pekerja
'#### Header Report #### - End
End Sub
Sub frm100_cash_in_out_report()
'on error resume next
Dim TM As Date
Dim TA As Date

Frm100_PAGE_SIZE = 29

TM = Frm100.DTPicker2 'Tarikh Mula
TA = Frm100.DTPicker3 'Tarikh Akhir

LM_START_ROW = Frm100.L8_Text

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm100_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm100.L9_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm100_PAGE_SIZE
        End If
    End If
End If

Frm100_LM_PAGE_FOUND = 0
Frm100.L7_Text = "Rekod kemasukkan atau ambilan tunai kedai dari " & Frm100.DTPicker2 & " hingga " & Frm100.DTPicker3 & "."

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 47_account_close where status='" & 1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm100_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm100_LM_PAGE_FOUND = 0 Then
        If Frm100.L9_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm100.L10_Text = Frm100.L10_Text + 1
                Frm100_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm100.L10_Text) Then
                    If Frm100.L10_Text <> 1 Then
                        Frm100.L10_Text = Frm100.L10_Text - 1
                        Frm100_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm100.L10_Text - 1) * Frm100_PAGE_SIZE) + x
    Frm100.MSFlexGrid1.Rows = x + 1
    Frm100.MSFlexGrid1.TextMatrix(x, 0) = x 'No.
    Frm100.MSFlexGrid1.TextMatrix(x, 1) = Y 'No.
    Frm100.MSFlexGrid1.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm100.MSFlexGrid1.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!jenis) Then 'Jenis
        If rs!jenis = 0 Then
            Frm100.MSFlexGrid1.TextMatrix(x, 4) = "Kemasukkan tunai"
        ElseIf rs!jenis = 1 Then
            Frm100.MSFlexGrid1.TextMatrix(x, 4) = "Pengeluaran tunai"
        End If
    End If
    If Not IsNull(rs!jumlah) Then 'Jumlah (RM)
        Frm100.MSFlexGrid1.TextMatrix(x, 5) = Format(rs!jumlah, "#,##0.00")
    Else
        Frm100.MSFlexGrid1.TextMatrix(x, 5) = "0.00"
    End If
    If Not IsNull(rs!staff_name) Then Frm100.MSFlexGrid1.TextMatrix(x, 6) = rs!staff_name 'Nama Pekerja
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Kemasukkan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah) from 47_account_close where status='" & 1 & "' AND jenis='" & 0 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm100.L11_Text = Format(rs(0), "#,##0.00") 'Jumlah (RM)
    If rs(0) = vbNullString Then
        Frm100.L11_Text = "0.00"
    End If
Else
    Frm100.L11_Text = "0.00"
End If

rs.Close
Set rs = Nothing

If Frm100.L11_Text = vbNullString Then
    Frm100.L11_Text = "0.00"
End If
'### Jumlah Kemasukkan ### - End

'### Jumlah Ambilan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah) from 47_account_close where status='" & 1 & "' AND jenis='" & 1 & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm100.L12_Text = Format(rs(0), "#,##0.00") 'Jumlah (RM)
    If rs(0) = vbNullString Then
        Frm100.L12_Text = "0.00"
    End If
Else
    Frm100.L12_Text = "0.00"
End If

rs.Close
Set rs = Nothing

If Frm100.L12_Text = vbNullString Then
    Frm100.L12_Text = "0.00"
End If
'### Jumlah Ambilan ### - End

'Frm100.Pic2.Visible = True

If x <> 0 Then
    Frm100.L8_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm100.Pic2.Visible = True
    Frm100.Pic3.Visible = False
    
    Frm100.L9_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm100.L9_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
    
    MsgBox "Tiada rekod dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm100_cetak_voucher()
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

Report75.Sections("Section4").Controls("L1").Caption = vbNullString 'Nama
Report75.Sections("Section4").Controls("L2").Caption = vbNullString 'No. Kad Pengenalan
Report75.Sections("Section4").Controls("L3").Caption = vbNullString 'No. Telefon
Report75.Sections("Section4").Controls("L4").Caption = "RM 0.00" 'Jumlah
Report75.Sections("Section4").Controls("L5").Caption = vbNullString 'No. Voucher
Report75.Sections("Section4").Controls("L6").Caption = vbNullString 'Tarikh
Report75.Sections("Section4").Controls("L7").Caption = vbNullString 'Nama Pekerja
Report75.Sections("Section4").Controls("L8").Caption = vbNullString 'Remarks

'### Reset maklumat kedai ### - Start
Report75.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report75.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report75.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report75.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report75.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report75.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report75.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report75.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report75.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report75.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

'### Maklumat voucher ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 47_account_close where no_voucher='" & G_VOUCHER & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    
    If Not IsNull(rs!Nama) Then Report75.Sections("Section4").Controls("L1").Caption = rs!Nama 'Nama
    If Not IsNull(rs!no_ic) Then Report75.Sections("Section4").Controls("L2").Caption = rs!no_ic 'No. Kad Pengenalan
    If Not IsNull(rs!no_tel) Then Report75.Sections("Section4").Controls("L3").Caption = rs!no_tel 'No. Telefon
    If Not IsNull(rs!jumlah) Then Report75.Sections("Section4").Controls("L4").Caption = "RM " & Format(rs!jumlah, "#,##0.00") 'Jumlah
    If Not IsNull(rs!no_voucher) Then Report75.Sections("Section4").Controls("L5").Caption = rs!no_voucher 'No. Voucher
    If Not IsNull(rs!tarikh) Then Report75.Sections("Section4").Controls("L6").Caption = rs!tarikh 'Tarikh
    If Not IsNull(rs!staff_name) Then Report75.Sections("Section4").Controls("L7").Caption = rs!staff_name 'Nama Pekerja
    If Not IsNull(rs!remarks) Then Report75.Sections("Section4").Controls("L8").Caption = rs!remarks 'Remarks

End If

rs.Close
Set rs = Nothing
'### Maklumat voucher ### - End

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 47_account_close where no_voucher='" & G_VOUCHER & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report75.DataSource = rs
    Report75.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing
End Sub
Sub Frm100_jurujual()
'On Error Resume Next
If MDI_frm1.L3_Text <> vbNullString Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        Frm100.CBB1 = rs!Samaran & "  |  " & rs!NoPekerja
        
    Else
        
        Frm100.CBB1.AddItem "" & "  |  " & rs!Samaran
        Frm100.CBB1 = "" & "  |  " & rs!Samaran
        
    End If
    
    rs.Close
    Set rs = Nothing
    
    If G_LOCK_JURUJUAL = "YES" Then
    
        Frm100.CBB1.Enabled = False
        Frm100.CBB1.BackColor = &H8000000A

    Else
    
        Frm100.CBB1.Enabled = True
        Frm100.CBB1.BackColor = &HFFFFFF

    End If

End If
End Sub
