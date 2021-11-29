Attribute VB_Name = "Module7"
Sub Frm92_Initial_Setting()
'on error resume next
GLOBAL_DISABLE = 0
Frm92.TB1 = vbNullString
Frm92.TB2 = "0.00"
'Frm92.TB3 = vbNullString

'Frm92.TB21 = "0.00"

'Frm92.TB27 = "0.00"
'Frm92.TB28 = "0.00"
'Frm92.TB29 = "0.00"
'Frm92.TB30 = "0.00"
'Frm92.TB31 = "0.00"
'Frm92.TB32 = "0.00"
'Frm92.TB38 = "0.00"
'Frm92.TB39 = "0.00"
'Frm92.TB40 = "0.00"
Frm92.TB41 = vbNullString
Frm92.TB42 = vbNullString
Frm92.TB43 = vbNullString
Frm92.TB44 = vbNullString
Frm92.TB46 = "0.00"
Frm92.TB47 = "0.00"
Frm92.TB48 = "0.00"
Frm92.TB49 = "0.00"
Frm92.TB45 = "0.00"

Frm92.L43_Text = vbNullString

Frm92.L7_Text = "0.00"
Frm92.L8_Text = "0.00"
Frm92.L9_Text = "0.00"
Frm92.L10_Text = "0.00"
Frm92.L11_Text = "0.00"
Frm92.L12_Text = "0.00"
Frm92.L13_Text = "0.00"
Frm92.L14_Text = "0.00"
'Frm92.L19_Text = "0.00"
Frm92.L20_Text = 0

'Frm92.L27_Text = "0.00"
'Frm92.L31_Text = "0.00"
'Frm92.L32_Text = "0.00"
'Frm92.L33_Text = "0.00"

'Frm92.L34_Text = "0.00"
'Frm92.L35_Text = 0
'Frm92.L36_Text = "0.00"
'Frm92.L37_Text = "0.00"
'Frm92.L38_Text = "0.00"
'Frm92.L39_Text = "0.00"
'Frm92.L40_Text = "0.00"
'Frm92.L41_Text = "0.00"
Frm92.L42_Text = 0

'Frm92.L47_Text = 0
'Frm92.L48_Text = vbNullString
'Frm92.L49_Text = vbNullString
Frm92.L51_Text = vbNullString
Frm92.L52_Text = vbNullString
Frm92.L53_Text = 0

Frm92.L15_Text.BackStyle = 0
Frm92.L18_Text.BackStyle = 0
Frm92.L42_Text.BackStyle = 0
'Frm92.L44_Text.BackStyle = 0

Frm92.CMD1.Visible = True
Frm92.CMD2.Visible = False
Frm92.CMD3.Visible = False
Frm92.CMD4.Visible = True
Frm92.CMD5.Visible = False
Frm92.CMD8.Visible = False
'Frm92.CMD9.Visible = True
'Frm92.CMD10.Visible = False
'Frm92.CMD13.Visible = False
'Frm92.CMD21.Visible = True
Frm92.CMD22.Visible = True

Frm92.L18_Text.Visible = False
'Frm92.L44_Text.Visible = False

Frm92.DTPicker1 = DateTime.Date
Frm92.DTPicker2 = DateTime.Date
Frm92.DTPicker3 = DateTime.Date
Frm92.DTPicker4 = DateTime.Date
Frm92.DTPicker5 = DateTime.Date
Frm92.DTPicker6 = DateTime.Date

Frm92.CB1 = 0
Frm92.CB2 = 1
Frm92.CB10 = 1
Frm92.CB3 = 0
Frm92.CB4 = 0
'Frm92.CB5 = 1
'Frm92.CB6 = 0
'Frm92.CB7 = 0

Frm92.CMD29.Visible = False
Frm92.CMD30.Visible = False
Frm92.CMD28.Visible = True

Frm92.CB9 = 0
Frm92.CB9.Enabled = True

Frm92.L15_Text = G_RATE_GST 'Jumlah Kadar GST
Frm92.L42_Text = G_RATE_GST 'Jumlah Kadar GST
If G_GST_JUAL = 1 Then
    Frm92.CB1 = 0
    Frm92.CB2 = 1
Else
    Frm92.CB1 = 1
    Frm92.CB2 = 0
End If

GoTo skip_aaa:

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Default1 = "Default" Then
        If Not IsNull(rs!gst_value) Then
            Frm92.L15_Text = rs!gst_value 'Jumlah Kadar GST
            Frm92.L42_Text = rs!gst_value 'Jumlah Kadar GST
        Else
            Frm92.L15_Text = 0
            Frm92.L42_Text = 0
        End If
        If Not IsNull(rs!gst_arinashi) Then
            If rs!gst_arinashi = 1 Then
                Frm92.CB1 = 0
                Frm92.CB2 = 1
                'Frm92.CB3 = 0
                'Frm92.CB4 = 1
            Else
                Frm92.CB1 = 1
                Frm92.CB2 = 0
                'Frm92.CB3 = 1
                'Frm92.CB4 = 0
            End If
        End If
        
        If Not IsNull(rs!ResitNo) Then
            Frm92.L17_Text = rs!ResitNo 'No. invoice
        Else
            Frm92.L17_Text = 1 'No. invoice
        End If
        If Not IsNull(rs!no_rujukan_tak_rasmi) Then
            Frm92.L28_Text = rs!no_rujukan_tak_rasmi 'No. invoice tidak rasmi
        Else
            Frm92.L28_Text = 1 'No. invoice tidak rasmi
        End If
        If Not IsNull(rs!no_rujukan_expense) Then
            Frm92.L45_Text = rs!no_rujukan_expense 'No. Rujukan Perbelanjaan
        Else
            Frm92.L45_Text = 1 'No. Rujukan Perbelanjaan
        End If
    End If
End If

rs.Close
Set rs = Nothing

skip_aaa:
'Frm92.CBB3.Clear

'Set rs = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from 74_cas_kad_kredit where status = 1 order by jenis_kad ASC", cn, adOpenKeyset, adLockOptimistic

'While rs.EOF = False
'    If Not IsNull(rs!jenis_kad) Then Frm92.CBB3.AddItem rs!jenis_kad
'    rs.MoveNext
'Wend

'rs.Close
'Set rs = Nothing

'### Padam Temp Database ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "TRUNCATE TABLE " & G_SERVICE_TEMP & ""

Set rs = cn.Execute(strsql)
Set rs = Nothing
'### Padam Temp Database ### - End

'### Padam Temp Database ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 37_expense_temp", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    rs.Delete
    rs.Update
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'### Padam Temp Database ### - End

user_level = MDI_frm1.L4_Text

If user_level = "Administration" Or user_level = "Guest/User" Then

    Frm92.CB9.Visible = False
    Frm92.Label33.Visible = False
    Frm92.Label6.Visible = False
    
Else
    
    If G_GST_SYSTEM = "YES" Then
        Frm92.CB9.Visible = True
        Frm92.Label33.Visible = True
        Frm92.Label6.Visible = True
        
        If G_INVOICE_RASMI = 0 Then
            Frm92.CB9 = 1
        Else
            Frm92.CB9 = 0
        End If
        
    Else
        Frm92.CB9.Visible = False
        Frm92.Label33.Visible = False
        Frm92.Label6.Visible = False
    End If
End If

Frm92.CBB1.Clear
Frm92.CBB2.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where status='" & "Aktif" & "' AND InvestorSmall = 0 AND InvestorBig = 0", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    If rs!user_level = 1 Or rs!user_level = 2 Or rs!user_level = 3 Or rs!user_level = G_LEVEL_USER Then
        Frm92.CBB1.AddItem rs!Samaran & "  |  " & rs!NoPekerja
        Frm92.CBB2.AddItem rs!Samaran & "  |  " & rs!NoPekerja
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'Call frm92_senarai_servis_header
'Call frm92_senarai_service
'Call Frm92_Header_Expense
'Call Frm92_senarai_expense
Call Frm92_jurujual
End Sub
Sub frm92_pic_visible()
'on error resume next
Frm92.Frame1.Left = 120
Frm92.Frame1.Top = 240
Frm92.Frame4.Left = 120
Frm92.Frame4.Top = 240
Frm92.Frame5.Left = 120
Frm92.Frame5.Top = 240
Frm92.Frame6.Left = 120
Frm92.Frame6.Top = 240
Frm92.Frame3.Left = 120
Frm92.Frame3.Top = 240

Frm92.Frame1.Visible = False
Frm92.Frame4.Visible = False
Frm92.Frame5.Visible = False
Frm92.Frame6.Visible = False
Frm92.Frame3.Visible = False
End Sub
Sub frm92_setting_report()
'on error resume next
Frm92.L70_Text = vbNullString
Frm92.L71_Text = vbNullString
Frm92.L72_Text = vbNullString
Frm92.L73_Text = vbNullString
Frm92.L74_Text = vbNullString

Frm92.CBB4.Clear

Frm92.CBB4.AddItem "Semua senarai servis"
Frm92.CBB4.AddItem "No. invoice"

Frm92.CBB4 = "Semua senarai servis"

Frm92.CBB6.Clear

Frm92.CBB6.AddItem "Semua cawangan"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where status='" & 1 & "' order by cawangan ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!cawangan) Then Frm92.CBB6.AddItem rs!cawangan
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm92.CBB6 = "Semua cawangan"

If MDI_frm1.L4_Text <> "HQ" And MDI_frm1.L4_Text <> "Developer" Then

    Frm92.CBB6 = MDI_frm1.L20_Text
    Frm92.CBB6.Enabled = False
    
Else
    
    Frm92.CBB6.Enabled = True
    
End If
End Sub
Sub frm92_senarai_service_header()
'on error resume next
With Frm92.LV1
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear

    Frm92.LV1.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "Maklumat Servis", 3000
    .ColumnHeaders.Add 5, , "Jumlah (RM)", 1400, 1
    .ColumnHeaders.Add 6, , "Jenis GST", 1200, 2
    .ColumnHeaders.Add 7, , "Jumlah GST (RM)", 1400, 1
    .ColumnHeaders.Add 8, , "Harga Dengan GST (RM)", 2350, 1

End With
End Sub
Sub frm92_senarai_service()
'on error resume next
Dim Frm92_LM_SR_TANPA As Double
Dim Frm92_LM_ZR_TANPA As Double
Dim Frm92_LM_SR_DENGAN As Double
Dim Frm92_LM_ZR_DENGAN As Double
Dim Frm92_LM_GST_SR As Double
Dim Frm92_LM_GST_ZR As Double

Frm92_LM_SR_TANPA = 0
Frm92_LM_ZR_TANPA = 0
Frm92_LM_SR_DENGAN = 0
Frm92_LM_ZR_DENGAN = 0
Frm92_LM_GST_SR = 0
Frm92_LM_GST_ZR = 0
x = 0

'Maklumat Servis
'Jumlah (RM)
'Jenis GST
'Jumlah GST (RM)
'Harga Dengan GST (RM)

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from " & G_SERVICE_TEMP & "", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    
    With Frm92.LV1.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , x
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!Detail) Then 'Maklumat Servis
            .ListSubItems.Add , , rs!Detail
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!jumlah) Then 'Jumlah (RM)
            .ListSubItems.Add , , Format(rs!jumlah, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!jenis_gst) Then 'Jenis GST
            .ListSubItems.Add , , rs!jenis_gst
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!jumlah_gst) Then 'Jumlah GST (RM)
            .ListSubItems.Add , , Format(rs!jumlah_gst, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_dengan_gst) Then 'Jumlah dengan GST (RM)
            .ListSubItems.Add , , Format(rs!harga_dengan_gst, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
    End With
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_tanpa_gst) from " & G_SERVICE_TEMP & " where kod_gst = 0", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm92_LM_ZR_TANPA = rs(0)

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_gst) from " & G_SERVICE_TEMP & " where kod_gst = 0", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm92_LM_GST_ZR = rs(0)

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_tanpa_gst) from " & G_SERVICE_TEMP & " where kod_gst = 1 OR kod_gst = 2", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm92_LM_SR_TANPA = rs(0)

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_gst) from " & G_SERVICE_TEMP & " where kod_gst = 1 OR kod_gst = 2", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm92_LM_GST_SR = rs(0)

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_dengan_gst) from " & G_SERVICE_TEMP & " where kod_gst = 0", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm92_LM_ZR_DENGAN = rs(0)

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(harga_dengan_gst) from " & G_SERVICE_TEMP & " where kod_gst = 1 OR kod_gst = 2", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm92_LM_SR_DENGAN = rs(0)

rs.Close
Set rs = Nothing

Frm92.L7_Text = Format(Frm92_LM_GST_ZR + Frm92_LM_GST_SR, "#,##0.00")  'Jumlah GST
Frm92.L9_Text = Format(Frm92_LM_SR_TANPA + Frm92_LM_ZR_TANPA, "#,##0.00")  'Jumlah Tanpa GST
Frm92.L10_Text = Format(Frm92_LM_SR_DENGAN + Frm92_LM_ZR_DENGAN, "#,##0.00")  'Jumlah Dengan GST
Frm92.L11_Text = Format(Frm92_LM_ZR_TANPA, "#,##0.00")  'Jumlah Harga ZR
Frm92.L12_Text = Format(Frm92_LM_GST_ZR, "#,##0.00")  'Jumlah Cukai ZR
Frm92.L13_Text = Format(Frm92_LM_SR_TANPA, "#,##0.00")  'Jumlah Harga SR
Frm92.L14_Text = Format(Frm92_LM_GST_SR, "#,##0.00")  'Jumlah Cukai SR
Frm92.L20_Text = x 'Bilangan
End Sub
Sub frm92_senarai_servis_header()
'on error resume next
With Frm92.LV2
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear

    Frm92.LV2.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "Tarikh", 1700, 2
    .ColumnHeaders.Add 5, , "No. Invoice", 2000
    .ColumnHeaders.Add 6, , "Bilangan Servis", 1600
    .ColumnHeaders.Add 7, , "Jumlah Tanpa GST (RM)", 2700, 1
    .ColumnHeaders.Add 8, , "Jumlah GST (RM)", 2000, 1
    .ColumnHeaders.Add 9, , "Jumlah Dengan GST (RM)", 2700, 1
    .ColumnHeaders.Add 10, , "Cawangan", 4800
    
End With
End Sub
Sub frm92_senarai_servis()
'on error resume next
Dim frm92_LM_TOTAL_PAGE As Double
Dim TM As Date
Dim TA As Date

frm92_LM_TOTAL_PAGE = 0
x = 0
Y = 0
frm92_PAGE_SIZE = 35

Frm92.L22_Text = 0
Frm92.L23_Text = Format(0, "#,##0.00")

If Frm92.CB16 = 0 Then
    Frm92.L70_Text = 0 '0 : Tiada carian mengikut tarikh , 1 : Carian mengikut tarikh
Else
    Frm92.L70_Text = 1 '0 : Tiada carian mengikut tarikh , 1 : Carian mengikut tarikh
    Frm92.L71_Text = Frm92.DTPicker2 'Tarik mula
    Frm92.L72_Text = Frm92.DTPicker3 'Tarikh akhir
End If

If Frm92.L70_Text = 1 Then '0 : Tiada carian mengikut tarikh , 1 : Carian mengikut tarikh
    TM = Frm92.L71_Text 'Tarikh mula
    TA = Frm92.L72_Text 'Tarikh akhir
End If

If Frm92.L73_Text = "Semua senarai servis" Then 'Semua senarai servis
    frm92_LM_SEARCH_1 = Null
    frm92_LM_SEARCH_1_LOGIC = "<>"
Else
    frm92_LM_SEARCH_1 = Frm92.L74_Text
    frm92_LM_SEARCH_1_LOGIC = "="
End If
If Frm92.L81_Text = "Semua cawangan" Then
    frm92_LM_SEARCH_2 = Null
    frm92_LM_SEARCH_2_LOGIC = "<>"
Else
    frm92_LM_SEARCH_2 = Frm92.L81_Text
    frm92_LM_SEARCH_2_LOGIC = "="
End If

If Frm92.L70_Text = 0 Then Frm92.L21_Text = "Senarai invoice servis kepada pelanggan , cawangan [" & Frm92.L81_Text & "]." 'Header
If Frm92.L70_Text = 1 Then Frm92.L21_Text = "Senarai invoice servis kepada pelanggan , cawangan [" & Frm92.L81_Text & "] dari " & TM & " hingga " & TA & "." 'Header

LM_START_ROW = Frm92.L62_Text

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + frm92_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm92.L63_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - frm92_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        Frm92.L60_Text = 1
    End If
End If

frm92_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm92.L70_Text = 0 Then rs.Open "select * from 22_jualan where cawangan " & frm92_LM_SEARCH_2_LOGIC & "'" & frm92_LM_SEARCH_2 & "' AND no_resit " & frm92_LM_SEARCH_1_LOGIC & "'" & frm92_LM_SEARCH_1 & "' AND status = 1 AND menu = 1 order by tarikh ASC LIMIT " & LM_START_ROW & "," & frm92_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If Frm92.L70_Text = 1 Then rs.Open "select * from 22_jualan where cawangan " & frm92_LM_SEARCH_2_LOGIC & "'" & frm92_LM_SEARCH_2 & "' AND no_resit " & frm92_LM_SEARCH_1_LOGIC & "'" & frm92_LM_SEARCH_1 & "' AND status = 1 AND menu = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'order by tarikh ASC LIMIT " & LM_START_ROW & "," & frm92_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If frm92_LM_PAGE_FOUND = 0 Then
        If Frm92.L63_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm92.L60_Text = Frm92.L60_Text + 1
                frm92_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm92.L60_Text) Then
                    If Frm92.L60_Text <> 1 Then
                        Frm92.L60_Text = Frm92.L60_Text - 1
                        frm92_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm92.L60_Text - 1) * frm92_PAGE_SIZE) + x
        
    With Frm92.LV2.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!tarikh) Then 'Tarikh
            .ListSubItems.Add , , rs!tarikh
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_resit) Then 'No. invoice
            .ListSubItems.Add , , rs!no_resit
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!kuantiti_barang) Then 'Bilangan servis
            .ListSubItems.Add , , rs!kuantiti_barang
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_barang) Then 'Jumlah Tanpa GST (RM)
            .ListSubItems.Add , , Format(rs!harga_barang, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!jumlah_cukai_gst) Then 'Jumlah GST (RM)
            .ListSubItems.Add , , Format(rs!jumlah_cukai_gst, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_barang_dengan_gst) Then 'Jumlah dengan GST (RM)
            .ListSubItems.Add , , Format(rs!harga_barang_dengan_gst, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        If Not IsNull(rs!cawangan) Then 'Cawangan
            .ListSubItems.Add , , rs!cawangan
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
If Frm92.L70_Text = 0 Then rs.Open "select COUNT(ID) from 22_jualan where cawangan " & frm92_LM_SEARCH_2_LOGIC & "'" & frm92_LM_SEARCH_2 & "' AND no_resit " & frm92_LM_SEARCH_1_LOGIC & "'" & frm92_LM_SEARCH_1 & "' AND status = 1 AND menu = 1", cn, adOpenKeyset, adLockOptimistic
If Frm92.L70_Text = 1 Then rs.Open "select COUNT(ID) from 22_jualan where cawangan " & frm92_LM_SEARCH_2_LOGIC & "'" & frm92_LM_SEARCH_2 & "' AND no_resit " & frm92_LM_SEARCH_1_LOGIC & "'" & frm92_LM_SEARCH_1 & "' AND status = 1 AND menu = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs(0)) Then
        frm92_LM_TOTAL_PAGE = Format(rs(0) / frm92_PAGE_SIZE, "0.00") 'Jumlah Page
        
        'Periksa Samada ada titik perpuluhan atau tidak
        If InStr(1, frm92_LM_TOTAL_PAGE, ".") <> 0 Then
        
            Frm85_LM_PAGE = Split(frm92_LM_TOTAL_PAGE, ".")(0)
            Frm85_LM_PAGE_LEBIHAN = Split(frm92_LM_TOTAL_PAGE, ".")(1)
            
            If Frm85_LM_PAGE_LEBIHAN <> "00" Then
                Frm92.L61_Text = Frm85_LM_PAGE + 1 'Total Page
            Else
                Frm92.L61_Text = Frm85_LM_PAGE
            End If
            
        Else
        
            Frm92.L61_Text = frm92_LM_TOTAL_PAGE
            
        End If
    
        If rs(0) = vbNullString Then
            Frm92.L61_Text = 0
        End If
    End If
Else
    Frm92.L61_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm92.L61_Text = vbNullString Then
    Frm92.L61_Text = 0
End If
'### Jumlah Data ### - End

'#### Jumlah Data #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm92.L70_Text = 0 Then rs.Open "select COUNT(ID) , SUM(harga_barang_dengan_gst) from 22_jualan where cawangan " & frm92_LM_SEARCH_2_LOGIC & "'" & frm92_LM_SEARCH_2 & "' AND no_resit " & frm92_LM_SEARCH_1_LOGIC & "'" & frm92_LM_SEARCH_1 & "' AND status = 1 AND menu = 1", cn, adOpenKeyset, adLockOptimistic
If Frm92.L70_Text = 1 Then rs.Open "select COUNT(ID) , SUM(harga_barang_dengan_gst) from 22_jualan where cawangan " & frm92_LM_SEARCH_2_LOGIC & "'" & frm92_LM_SEARCH_2 & "' AND no_resit " & frm92_LM_SEARCH_1_LOGIC & "'" & frm92_LM_SEARCH_1 & "' AND status = 1 AND menu = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm92.L22_Text = rs(0)
If Not IsNull(rs(1)) Then Frm92.L23_Text = Format(rs(1), "#,##0.00")

rs.Close
Set rs = Nothing
'#### Jumlah Data#### - End

If x <> 0 Then
    Frm92.L63_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm92.L62_Text = LM_START_ROW
Else
    Frm92.L63_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If

Frm92.Frame4.Visible = True
Frm92.Frame3.Visible = False
End Sub
Sub Frm92_Resit_Servis()
'on error resume next
DATA_FOUND = 0
Frm92_LM_KATEGORI_PEMBELI = 0 '0 : Pembeli Tidak Berdaftar , 1 : Pembeli Berdaftar , 2 : Ahli
Frm92_NO_PELANGGAN = vbNullString
Frm92_LM_KATEGORI = 0

Report42.Sections("Section5").Controls("L10").Caption = "0.00" 'Harga Keseluruhan Bagi Barang SR
Report42.Sections("Section5").Controls("L11").Caption = "0.00" 'Jumlah Cukai Bagi SR
Report42.Sections("Section5").Controls("L12").Caption = "0.00" 'Harga Keseluruhan Bagi Barang ZR
Report42.Sections("Section5").Controls("L13").Caption = "0.00" 'Jumlah Cukai Bagi ZR
Report42.Sections("Section2").Controls("L3").Caption = vbNullString 'No. Invoice
Report42.Sections("Section2").Controls("L4").Caption = vbNullString 'Tarikh
Report42.Sections("Section2").Controls("L5").Caption = vbNullString 'Maklumat Pembeli : Nama
Report42.Sections("Section2").Controls("L7").Caption = vbNullString 'Maklumat Pembeli : No. Telefon
Report42.Sections("Section2").Controls("L8").Caption = vbNullString 'Jurujual

'### Reset maklumat kedai ### - Start
Report42.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report42.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report42.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report42.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report42.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
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
    
        If Not IsNull(rs!nama_kedai) Then Report42.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
        If Not IsNull(rs!no_pendaftaran) Then Report42.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
        If Not IsNull(rs!alamat) Then Report42.Sections("Section4").Controls("L202").Caption = rs!alamat
        If Not IsNull(rs!no_tel) Then Report42.Sections("Section4").Controls("L203").Caption = rs!no_tel
        If Not IsNull(rs!no_id_gst) Then Report42.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
        If Not IsNull(rs!gst_ari_nashi) Then
            If rs!gst_ari_nashi = 0 Then
                Report42.Sections("Section2").Controls("L205").Caption = "INVOICE"
            ElseIf rs!gst_ari_nashi = 1 Then
                Report42.Sections("Section2").Controls("L205").Caption = "TAX INVOICE"
            End If
        Else
            Report42.Sections("Section2").Controls("L205").Caption = "INVOICE"
        End If
        
    End If
    
    rs.Close
    Set rs = Nothing
    '### Maklumat kedai ### - End
    
    Report42.Sections("Section4").Visible = True
Else
    Report42.Sections("Section4").Visible = False
End If

Report42.Sections("Section2").Controls("L3").Caption = G_No_RESIT_SERVIS 'No. Resit Ansuran

Frm92_DATA_CUST_FOUND = 0 '0 : Pelanggan Tidak Berdaftar , 1 : Pelanggan Berdaftar , 2 : Ahli

'### Maklumat Bayaran Dan GST ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where no_resit='" & G_No_RESIT_SERVIS & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!tarikh) Then Report42.Sections("Section2").Controls("L4").Caption = rs!tarikh 'Tarikh
    If Not IsNull(rs!gst_sr_harga) Then Report42.Sections("Section5").Controls("L10").Caption = Format(rs!gst_sr_harga, "#,##0.00") 'Harga Keseluruhan Bagi Barang SR
    If Not IsNull(rs!gst_sr_cukai) Then Report42.Sections("Section5").Controls("L11").Caption = Format(rs!gst_sr_cukai, "#,##0.00") 'Jumlah Cukai Bagi SR
    If Not IsNull(rs!gst_zr_harga) Then Report42.Sections("Section5").Controls("L12").Caption = Format(rs!gst_zr_harga, "#,##0.00") 'Harga Keseluruhan Bagi Barang ZR
    If Not IsNull(rs!gst_zr_cukai) Then Report42.Sections("Section5").Controls("L13").Caption = Format(rs!gst_zr_cukai, "#,##0.00") 'Jumlah Cukai Bagi ZR
    If Not IsNull(rs!harga_jualan) Then Report42.Sections("Section5").Controls("L9").Caption = Format(rs!harga_jualan, "#,##0.00") 'Jumlah Bayaran

    If Not IsNull(rs!no_rujukan_pembeli) Then
        Frm92_NO_PELANGGAN = rs!no_rujukan_pembeli
        Frm92_DATA_CUST_FOUND = 1 '0 : Pelanggan Tidak Berdaftar , 1 : Pelanggan Berdaftar , 2 : Ahli
    End If
    
    If Not IsNull(rs!no_rujukan_pembeli) Then Frm92_NO_PELANGGAN = rs!no_rujukan_pembeli
    If Not IsNull(rs!kategori_pembeli) Then Frm92_LM_KATEGORI = rs!kategori_pembeli
    If Not IsNull(rs!no_pekerja) Then Frm92_LM_No_PEKERJA = rs!no_pekerja
    
End If

rs.Close
Set rs = Nothing
'### Maklumat Bayaran Dan GST ### - End

'### Nama Pekerja ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where NoPekerja='" & Frm92_LM_No_PEKERJA & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!Samaran) Then Report42.Sections("Section2").Controls("L8").Caption = rs!Samaran 'Nama Samaran
End If

rs.Close
Set rs = Nothing
'### Nama Pekerja ### - End

'### Maklumat Pelanggan ### - Start
'If Frm92_LM_KATEGORI <> 0 Then
If Frm92_DATA_CUST_FOUND = 1 Then '0 : Pelanggan Tidak Berdaftar , 1 : Pelanggan Berdaftar , 2 : Ahli

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from senarai_pelanggan where no_pelanggan='" & Frm92_NO_PELANGGAN & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!Nama) Then Report42.Sections("Section2").Controls("L5").Caption = rs!Nama 'Maklumat Pembeli : Nama
        If Not IsNull(rs!no_tel) Then Report42.Sections("Section2").Controls("L7").Caption = rs!no_tel 'Maklumat Pembeli : No. Telefon
    End If
    
    rs.Close
    Set rs = Nothing
End If
'### Maklumat Pelanggan ### - End

'If Frm92_LM_KATEGORI = 0 Then
If Frm92_DATA_CUST_FOUND = 0 Then '0 : Pelanggan Tidak Berdaftar , 1 : Pelanggan Berdaftar , 2 : Ahli
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 44_senarai_pelanggan where no_resit='" & G_No_RESIT_SERVIS & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        
        If Not IsNull(rs!Nama) Then Report42.Sections("Section2").Controls("L5").Caption = rs!Nama 'Maklumat Pembeli : Nama
        If Not IsNull(rs!no_tel) Then Report42.Sections("Section2").Controls("L7").Caption = rs!no_tel 'Maklumat Pembeli : No. Telefon

    End If
    
    rs.Close
    Set rs = Nothing

End If

'### Paparan Resit Servis ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 35_senarai_servis where no_resit_servis='" & G_No_RESIT_SERVIS & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report42.DataSource = rs
    If G_PREVIEW = 1 Then Report42.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing

If G_PREVIEW = 0 Then Report42.PrintReport
'### Paparan Resit Servis ### - End
    
G_No_RESIT_SERVIS = vbNullString

End Sub
Sub Frm92_report_expenses_header()
'on error resume next

With Frm92.LV3
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear

    Frm92.LV3.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "Tarikh", 1500, 2
    .ColumnHeaders.Add 5, , "No. Invoice", 2000
    .ColumnHeaders.Add 6, , "No. Voucher", 2000
    .ColumnHeaders.Add 7, , "Nama Kedai", 3000
    .ColumnHeaders.Add 8, , "No. ID GST", 2300
    .ColumnHeaders.Add 9, , "Tujuan", 6000
    .ColumnHeaders.Add 10, , "Jumlah (RM)", 1500, 1
    .ColumnHeaders.Add 11, , "Jumlah GST (RM)", 1800, 1
    .ColumnHeaders.Add 12, , "Jenis", 2500
    .ColumnHeaders.Add 13, , "Cawangan", 2500

End With
End Sub
Sub Frm92_report_expenses()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim frm92_LM_TOTAL_PAGE As Double

frm92_PAGE_SIZE = 33
frm92_LM_TOTAL_PAGE = 0
x = 0

TM = Frm92.L76_Text 'Tarikh Mula
TA = Frm92.L77_Text 'Tarikh Akhir

Frm92.L79_Text = "Bilangan : 0"
Frm92.L80_Text = "Jumlah   : RM 0.00"

re_gen_report:

If Frm92.L82_Text = "Semua jenis" Then 'Jenis
    frm92_LM_SEARCH_1 = Null
    frm92_LM_SEARCH_1_LOGIC = "<>"
Else
    frm92_LM_SEARCH_1 = Frm92.L82_Text
    frm92_LM_SEARCH_1_LOGIC = "="
End If

If Frm92.L83_Text = "Semua cawangan" Then
    frm92_LM_SEARCH_2 = Null
    frm92_LM_SEARCH_2_LOGIC = "<>"
Else
    frm92_LM_SEARCH_2 = Frm92.L83_Text
    frm92_LM_SEARCH_2_LOGIC = "="
End If

If Frm92.L78_Text = 0 Then Frm92.L46_Text = "Senarai perbelanjaan kedai , jenis [" & Frm92.L82_Text & "] dan cawangan [" & Frm92.L83_Text & "]."  'Header
If Frm92.L78_Text = 1 Then Frm92.L46_Text = "Senarai perbelanjaan kedai , jenis [" & Frm92.L82_Text & "] dan cawangan [" & Frm92.L83_Text & "] dari " & TM & " hingga " & TA & "."  'Header

LM_START_ROW = Frm92.L69_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + frm92_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm92.L75_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - frm92_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        Frm92.L67_Text = 1
    End If
End If

frm92_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If Frm92.L78_Text = 0 Then rs.Open "select * from 39_akaun_expense where cawangan " & frm92_LM_SEARCH_2_LOGIC & "'" & frm92_LM_SEARCH_2 & "' AND jenis_expense " & frm92_LM_SEARCH_1_LOGIC & "'" & frm92_LM_SEARCH_1 & "' AND menu = 1 AND status = 1 order by tarikh ASC LIMIT " & LM_START_ROW & "," & frm92_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If Frm92.L78_Text = 1 Then rs.Open "select * from 39_akaun_expense where cawangan " & frm92_LM_SEARCH_2_LOGIC & "'" & frm92_LM_SEARCH_2 & "' AND jenis_expense " & frm92_LM_SEARCH_1_LOGIC & "'" & frm92_LM_SEARCH_1 & "'AND menu = 1 AND status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & frm92_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1
    If frm92_LM_PAGE_FOUND = 0 Then
        If Frm92.L75_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm92.L67_Text = Frm92.L67_Text + 1 'Paparan Page ke-xxx
                frm92_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm92.L67_Text) Then
                    If Frm92.L67_Text <> 1 Then
                        Frm92.L67_Text = Frm92.L67_Text - 1 'Paparan Page ke-xxx
                        frm92_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If

    Y = ((Frm92.L67_Text - 1) * frm92_PAGE_SIZE) + x
    
    With Frm92.LV3.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!tarikh) Then 'Tarikh
            .ListSubItems.Add , , rs!tarikh
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_resit) Then 'No. Invoice
            .ListSubItems.Add , , rs!no_resit
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_voucher) Then 'No. Voucher
            .ListSubItems.Add , , rs!no_voucher
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!nama_kedai) Then 'Nama Kedai
            .ListSubItems.Add , , rs!nama_kedai
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_id_gst) Then 'No ID GST
            .ListSubItems.Add , , rs!no_id_gst
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!tujuan) Then 'Tujuan
            .ListSubItems.Add , , rs!tujuan
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!harga_dengan_gst) Then 'Jumlah (RM)
            .ListSubItems.Add , , Format(rs!harga_dengan_gst, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!gst_sr_cukai) Then 'Jumlah GST (RM)
            .ListSubItems.Add , , Format(rs!gst_sr_cukai, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!jenis_expense) Then 'Jenis
            .ListSubItems.Add , , rs!jenis_expense
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!cawangan) Then 'Cawangan
            .ListSubItems.Add , , rs!cawangan
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
If Frm92.L78_Text = 0 Then rs.Open "select COUNT(ID) , SUM(harga_dengan_gst) from 39_akaun_expense where cawangan " & frm92_LM_SEARCH_2_LOGIC & "'" & frm92_LM_SEARCH_2 & "' AND jenis_expense " & frm92_LM_SEARCH_1_LOGIC & "'" & frm92_LM_SEARCH_1 & "' AND menu = 1 AND status = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic
If Frm92.L78_Text = 1 Then rs.Open "select COUNT(ID) , SUM(harga_dengan_gst) from 39_akaun_expense where cawangan " & frm92_LM_SEARCH_2_LOGIC & "'" & frm92_LM_SEARCH_2 & "' AND jenis_expense " & frm92_LM_SEARCH_1_LOGIC & "'" & frm92_LM_SEARCH_1 & "'AND menu = 1 AND status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    frm92_LM_TOTAL_PAGE = Format(rs(0) / frm92_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, frm92_LM_TOTAL_PAGE, ".") <> 0 Then
    
        frm92_LM_PAGE = Split(frm92_LM_TOTAL_PAGE, ".")(0)
        frm92_LM_PAGE_LEBIHAN = Split(frm92_LM_TOTAL_PAGE, ".")(1)
        
        If frm92_LM_PAGE_LEBIHAN <> "00" Then
            Frm92.L68_Text = frm92_LM_PAGE + 1
        Else
            Frm92.L68_Text = frm92_LM_PAGE
        End If
        
    Else
    
        Frm92.L68_Text = frm92_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm92.L68_Text = 0
    End If
Else
    Frm92.L68_Text = 0
End If

If Not IsNull(rs(0)) Then Frm92.L79_Text = "Bilangan : " & rs(0) 'Jumlah bilangan barang jualan
If Not IsNull(rs(1)) Then Frm92.L80_Text = "Jumlah   : RM " & Format(rs(1), "#,##0.00")

rs.Close
Set rs = Nothing

If x <> 0 Then
    Frm92.L69_Text = LM_START_ROW
End If

If Frm92.L67_Text <> vbNullString And IsNumeric(Frm92.L67_Text) Then
    If Frm92.L68_Text <> vbNullString And IsNumeric(Frm92.L68_Text) Then
        frm92_LM_CURR_PAGE = Frm92.L67_Text
        frm92_LM_TOTAL_PAGE = Frm92.L68_Text
        
        If frm92_LM_CURR_PAGE > frm92_LM_TOTAL_PAGE Then
            
            Frm92.L67_Text = Frm92.L67_Text - 1
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            GoTo re_gen_report:
            
        End If
    End If
End If
End Sub
Sub Frm92_excel_overall()
'On Error Resume Next
Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
       
x = 0

Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh
        .Columns("C").ColumnWidth = 20 'No. Invoice
        .Columns("D").ColumnWidth = 20 'Bilangan
        .Columns("E").ColumnWidth = 20 'Jumlah Tanpa GST (RM)
        .Columns("F").ColumnWidth = 20 'Jumlah Dengan GST (RM)
    
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
                .Cells(1, 4) = rs!nama_kedai
                .Cells(1, 4).Font.Name = "Times New Roman"
            End If
            If Not IsNull(rs!no_pendaftaran) Then .Cells(2, 4) = rs!no_pendaftaran
            If Not IsNull(rs!alamat) Then .Cells(3, 4) = rs!alamat
            If Not IsNull(rs!no_tel) Then .Cells(4, 4) = rs!no_tel
            If Not IsNull(rs!no_id_gst) Then .Cells(5, 4) = rs!no_id_gst
        End If
        
        rs.Close
        Set rs = Nothing
        '### Maklumat kedai ### - End
        
        .Cells(1, 4).Font.Bold = True
        .Cells(1, 4).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 4).HorizontalAlignment = xlCenter
        Next Row
        
        .Cells(7, 1) = Frm92.L21_Text 'Header Report
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh"
        .Cells(8, 3) = "No. Invoice"
        .Cells(8, 4) = "Bilangan"
        .Cells(8, 5) = "Jumlah Tanpa GST (RM)"
        .Cells(8, 6) = "Jumlah Dengan GST (RM)"
        
        For i = 1 To 6
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
                
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 36_akaun_servis order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh) Then .Cells(8 + x, 2) = "'" & rs!tarikh 'Tarikh Jualan
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_resit_servis) Then .Cells(8 + x, 3) = rs!no_resit_servis 'No. Invoice
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!bilangan) Then .Cells(8 + x, 4) = rs!bilangan 'Bilangan
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            .Cells(8 + x, 5).NumberFormat = "#,##0.00"
            .Cells(8 + x, 5).HorizontalAlignment = xlCenter
            If Not IsNull(rs!jumlah_tanpa_gst) Then 'Jumlah Tanpa GST (RM)
                .Cells(8 + x, 5) = rs!jumlah_tanpa_gst
            Else
                .Cells(8 + x, 5) = "0.00"
            End If

            .Cells(8 + x, 6).NumberFormat = "#,##0.00"
            .Cells(8 + x, 6).HorizontalAlignment = xlCenter
            If Not IsNull(rs!harga_dengan_gst) Then 'Jumlah Dengan GST (RM)
                .Cells(8 + x, 6) = rs!harga_dengan_gst
            Else
                .Cells(8 + x, 6) = "0.00"
            End If

            For Col = 1 To 6
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
    
        x = x + 2
        .Cells(8 + x, 1) = "Bilangan : " & Frm92.L22_Text 'Bilangan
        x = x + 1
        .Cells(8 + x, 1) = "Jumlah : RM " & Frm92.L23_Text 'Jumlah
        
        x = x + 4
        .Cells(8 + x, 1).Font.Bold = True
        .Cells(8 + x, 1) = "Report Generated By Sankyu System [ +6010 - 900 4788 , sankyusystem@gmail.com ]" 'Watermark Sankyu System
    End With
    
    ' This makes Excel visible
    xlObject.Visible = True
    xlObject.EnableEvents = True
End If
End Sub
Sub Frm92_excel_overall_tarikh()
'On Error Resume Next
Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
Dim TM As Date
Dim TA As Date

TM = Frm92.L25_Text 'Tarikh Mula
TA = Frm92.L26_Text 'Tarikh Akhir

x = 0

Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh
        .Columns("C").ColumnWidth = 20 'No. Invoice
        .Columns("D").ColumnWidth = 20 'Bilangan
        .Columns("E").ColumnWidth = 20 'Jumlah Tanpa GST (RM)
        .Columns("F").ColumnWidth = 20 'Jumlah Dengan GST (RM)
        
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
                .Cells(1, 4) = rs!nama_kedai
                .Cells(1, 4).Font.Name = "Times New Roman"
            End If
            If Not IsNull(rs!no_pendaftaran) Then .Cells(2, 4) = rs!no_pendaftaran
            If Not IsNull(rs!alamat) Then .Cells(3, 4) = rs!alamat
            If Not IsNull(rs!no_tel) Then .Cells(4, 4) = rs!no_tel
            If Not IsNull(rs!no_id_gst) Then .Cells(5, 4) = rs!no_id_gst
        End If
        
        rs.Close
        Set rs = Nothing
        '### Maklumat kedai ### - End
        
        .Cells(1, 4).Font.Bold = True
        .Cells(1, 4).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 4).HorizontalAlignment = xlCenter
        Next Row
        
        .Cells(7, 1) = Frm92.L21_Text 'Header Report
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh"
        .Cells(8, 3) = "No. Invoice"
        .Cells(8, 4) = "Bilangan"
        .Cells(8, 5) = "Jumlah Tanpa GST (RM)"
        .Cells(8, 6) = "Jumlah Dengan GST (RM)"
        
        For i = 1 To 6
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i
                
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 36_akaun_servis where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
        
            x = x + 1
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh) Then .Cells(8 + x, 2) = "'" & rs!tarikh 'Tarikh Jualan
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_resit_servis) Then .Cells(8 + x, 3) = rs!no_resit_servis 'No. Invoice
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!bilangan) Then .Cells(8 + x, 4) = rs!bilangan 'Bilangan
            .Cells(8 + x, 4).HorizontalAlignment = xlCenter
            
            .Cells(8 + x, 5).NumberFormat = "#,##0.00"
            .Cells(8 + x, 5).HorizontalAlignment = xlCenter
            If Not IsNull(rs!jumlah_tanpa_gst) Then 'Jumlah Tanpa GST (RM)
                .Cells(8 + x, 5) = rs!jumlah_tanpa_gst
            Else
                .Cells(8 + x, 5) = "0.00"
            End If

            .Cells(8 + x, 6).NumberFormat = "#,##0.00"
            .Cells(8 + x, 6).HorizontalAlignment = xlCenter
            If Not IsNull(rs!harga_dengan_gst) Then 'Jumlah Dengan GST (RM)
                .Cells(8 + x, 6) = rs!harga_dengan_gst
            Else
                .Cells(8 + x, 6) = "0.00"
            End If

            For Col = 1 To 6
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
    
        x = x + 2
        .Cells(8 + x, 1) = "Bilangan : " & Frm92.L22_Text 'Bilangan
        x = x + 1
        .Cells(8 + x, 1) = "Jumlah : RM " & Frm92.L23_Text 'Jumlah
        
        x = x + 4
        .Cells(8 + x, 1).Font.Bold = True
        .Cells(8 + x, 1) = "Report Generated By Sankyu System [ +6010 - 900 4788 , sankyusystem@gmail.com ]" 'Watermark Sankyu System
    End With
    
    ' This makes Excel visible
    xlObject.Visible = True
    xlObject.EnableEvents = True
End If
End Sub
Sub Frm92_excel_detail()
'On Error Resume Next
Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
Dim TM As Date
Dim TA As Date
Dim Frm92_LM_TANPA_GST As Double
Dim Frm92_LM_GST As Double

TM = Frm92.L25_Text 'Tarikh Mula
TA = Frm92.L26_Text 'Tarikh Akhir

x = 0

Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh
        .Columns("C").ColumnWidth = 20 'No. Invoice
        .Columns("D").ColumnWidth = 40 'Maklumat Servis
        .Columns("E").ColumnWidth = 20 'Jumlah Tanpa GST (RM)
        .Columns("F").ColumnWidth = 20 'Jumlah Dengan GST (RM)
    
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
                .Cells(1, 4) = rs!nama_kedai
                .Cells(1, 4).Font.Name = "Times New Roman"
            End If
            If Not IsNull(rs!no_pendaftaran) Then .Cells(2, 4) = rs!no_pendaftaran
            If Not IsNull(rs!alamat) Then .Cells(3, 4) = rs!alamat
            If Not IsNull(rs!no_tel) Then .Cells(4, 4) = rs!no_tel
            If Not IsNull(rs!no_id_gst) Then .Cells(5, 4) = rs!no_id_gst
        End If
        
        rs.Close
        Set rs = Nothing
        '### Maklumat kedai ### - End
        
        .Cells(1, 4).Font.Bold = True
        .Cells(1, 4).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 4).HorizontalAlignment = xlCenter
        Next Row
        
        .Cells(7, 1) = Frm92.L21_Text 'Header Report
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh"
        .Cells(8, 3) = "No. Invoice"
        .Cells(8, 4) = "Maklumat Servis"
        .Cells(8, 5) = "Jumlah Tanpa GST (RM)"
        .Cells(8, 6) = "Jumlah Dengan GST (RM)"
        
        For i = 1 To 6
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 35_senarai_servis order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
            Frm92_LM_TANPA_GST = 0
            Frm92_LM_GST = 0
            x = x + 1
            
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh) Then .Cells(8 + x, 2) = "'" & rs!tarikh 'Tarikh Jualan
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_resit_servis) Then .Cells(8 + x, 3) = rs!no_resit_servis 'No. Invoice
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!Detail) Then .Cells(8 + x, 4) = rs!Detail 'Detail
            
            .Cells(8 + x, 5).NumberFormat = "#,##0.00"
            .Cells(8 + x, 5).HorizontalAlignment = xlCenter
            If Not IsNull(rs!harga_tanpa_gst) Then 'Jumlah Tanpa GST (RM)
                .Cells(8 + x, 5) = rs!harga_tanpa_gst
                If IsNumeric(rs!harga_tanpa_gst) Then Frm92_LM_TANPA_GST = rs!harga_tanpa_gst
            Else
                .Cells(8 + x, 5) = "0.00"
            End If
            
            If Not IsNull(rs!jumlah_gst) Then
                If IsNumeric(rs!jumlah_gst) Then Frm92_LM_GST = rs!jumlah_gst
            End If
            
            .Cells(8 + x, 6).NumberFormat = "#,##0.00"
            .Cells(8 + x, 6).HorizontalAlignment = xlCenter
            .Cells(8 + x, 6) = Format(Frm92_LM_TANPA_GST + Frm92_LM_GST, "#,##0.00") 'Jumlah Dengan GST (RM)

            For Col = 1 To 6
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
    
        x = x + 2
        .Cells(8 + x, 1) = "Bilangan : " & Frm92.L22_Text 'Bilangan
        x = x + 1
        .Cells(8 + x, 1) = "Jumlah : RM " & Frm92.L23_Text 'Jumlah
        
        x = x + 4
        .Cells(8 + x, 1).Font.Bold = True
        .Cells(8 + x, 1) = "Report Generated By Sankyu System [ +6010 - 900 4788 , sankyusystem@gmail.com ]" 'Watermark Sankyu System
    End With
    
    ' This makes Excel visible
    xlObject.Visible = True
    xlObject.EnableEvents = True
End If
End Sub
Sub Frm92_excel_detail_tarikh()
'On Error Resume Next
Dim xlObject As Excel.Application
Dim xlWB As Excel.Workbook
Dim TM As Date
Dim TA As Date
Dim Frm92_LM_TANPA_GST As Double
Dim Frm92_LM_GST As Double

TM = Frm92.L25_Text 'Tarikh Mula
TA = Frm92.L26_Text 'Tarikh Akhir

x = 0

Note = "Sistem Akan Mengambil Masa Untuk Mengeluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Sila Tunggu Sehingga Sistem Siap Keluarkan Report." & vbCrLf & _
        "" & vbCrLf & _
        "Teruskan ?"
        
Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")
If Answer = vbNo Then
    Exit Sub
End If
If Answer = vbYes Then
    Set xlObject = New Excel.Application
    Set xlWB = xlObject.Workbooks.Add
               
    'xlObject.Visible = True
    With xlObject.ActiveWorkbook.ActiveSheet
        .Cells.VerticalAlignment = xlCenter
        .Columns("A").ColumnWidth = 5 'No.
        .Columns("B").ColumnWidth = 20 'Tarikh
        .Columns("C").ColumnWidth = 20 'No. Invoice
        .Columns("D").ColumnWidth = 40 'Maklumat Servis
        .Columns("E").ColumnWidth = 20 'Jumlah Tanpa GST (RM)
        .Columns("F").ColumnWidth = 20 'Jumlah Dengan GST (RM)
    
        '### Maklumat kedai ### - Start
        If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
            
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
                .Cells(1, 4) = rs!nama_kedai
                .Cells(1, 4).Font.Name = "Times New Roman"
            End If
            If Not IsNull(rs!no_pendaftaran) Then .Cells(2, 4) = rs!no_pendaftaran
            If Not IsNull(rs!alamat) Then .Cells(3, 4) = rs!alamat
            If Not IsNull(rs!no_tel) Then .Cells(4, 4) = rs!no_tel
            If Not IsNull(rs!no_id_gst) Then .Cells(5, 4) = rs!no_id_gst
        End If
        
        rs.Close
        Set rs = Nothing
        '### Maklumat kedai ### - End
        
        .Cells(1, 4).Font.Bold = True
        .Cells(1, 4).Font.Size = 30
        
        For Row = 1 To 5
            .Cells(Row, 4).HorizontalAlignment = xlCenter
        Next Row
        
        .Cells(7, 1) = Frm92.L21_Text 'Header Report
        
        .Cells(8, 1) = "No."
        .Cells(8, 2) = "Tarikh"
        .Cells(8, 3) = "No. Invoice"
        .Cells(8, 4) = "Maklumat Servis"
        .Cells(8, 5) = "Jumlah Tanpa GST (RM)"
        .Cells(8, 6) = "Jumlah Dengan GST (RM)"
        
        For i = 1 To 6
            .Cells(8, i).HorizontalAlignment = xlCenter
            .Cells(8, i).Interior.ColorIndex = 15
            .Cells(8, i).WrapText = True
            .Cells(8, i).Borders.LineStyle = xlContinuous
        Next i

        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 35_senarai_servis where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

        While rs.EOF = False
            Frm92_LM_TANPA_GST = 0
            Frm92_LM_GST = 0
            x = x + 1
            
            .Cells(8 + x, 1) = x 'No.
            .Cells(8 + x, 1).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!tarikh) Then .Cells(8 + x, 2) = "'" & rs!tarikh 'Tarikh Jualan
            .Cells(8 + x, 2).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!no_resit_servis) Then .Cells(8 + x, 3) = rs!no_resit_servis 'No. Invoice
            .Cells(8 + x, 3).HorizontalAlignment = xlCenter
            
            If Not IsNull(rs!Detail) Then .Cells(8 + x, 4) = rs!Detail 'Detail
            
            .Cells(8 + x, 5).NumberFormat = "#,##0.00"
            .Cells(8 + x, 5).HorizontalAlignment = xlCenter
            If Not IsNull(rs!harga_tanpa_gst) Then 'Jumlah Tanpa GST (RM)
                .Cells(8 + x, 5) = rs!harga_tanpa_gst
                If IsNumeric(rs!harga_tanpa_gst) Then Frm92_LM_TANPA_GST = rs!harga_tanpa_gst
            Else
                .Cells(8 + x, 5) = "0.00"
            End If
            
            If Not IsNull(rs!jumlah_gst) Then
                If IsNumeric(rs!jumlah_gst) Then Frm92_LM_GST = rs!jumlah_gst
            End If
            
            .Cells(8 + x, 6).NumberFormat = "#,##0.00"
            .Cells(8 + x, 6).HorizontalAlignment = xlCenter
            .Cells(8 + x, 6) = Format(Frm92_LM_TANPA_GST + Frm92_LM_GST, "#,##0.00") 'Jumlah Dengan GST (RM)

            For Col = 1 To 6
                .Cells(8 + x, Col).Borders.LineStyle = xlContinuous
            Next Col
            
            rs.MoveNext
        Wend
        
        rs.Close
        Set rs = Nothing
    
        x = x + 2
        .Cells(8 + x, 1) = "Bilangan : " & Frm92.L22_Text 'Bilangan
        x = x + 1
        .Cells(8 + x, 1) = "Jumlah : RM " & Frm92.L23_Text 'Jumlah
        
        x = x + 4
        .Cells(8 + x, 1).Font.Bold = True
        .Cells(8 + x, 1) = "Report Generated By Sankyu System [ +6010 - 900 4788 , sankyusystem@gmail.com ]" 'Watermark Sankyu System
    End With
    
    ' This makes Excel visible
    xlObject.Visible = True
    xlObject.EnableEvents = True
End If
End Sub
Sub Frm92_jurujual()
'On Error Resume Next
If MDI_frm1.L3_Text <> vbNullString Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        Frm92.CBB1 = rs!Samaran & "  |  " & rs!NoPekerja
        Frm92.CBB2 = rs!Samaran & "  |  " & rs!NoPekerja
        
    Else
        
        Frm92.CBB1.AddItem "" & "  |  " & rs!Samaran
        Frm92.CBB1 = "" & "  |  " & rs!Samaran
        
        Frm92.CBB2.AddItem "" & "  |  " & rs!Samaran
        Frm92.CBB2 = "" & "  |  " & rs!Samaran
        
    End If
    
    rs.Close
    Set rs = Nothing
    
    If G_LOCK_JURUJUAL = "YES" Then
    
        Frm92.CBB1.Enabled = False
        Frm92.CBB1.BackColor = &H8000000A

        Frm92.CBB2.Enabled = False
        Frm92.CBB2.BackColor = &H8000000A
        
    Else
    
        Frm92.CBB1.Enabled = True
        Frm92.CBB1.BackColor = &HFFFFFF
        
        Frm92.CBB2.Enabled = True
        Frm92.CBB2.BackColor = &HFFFFFF

    End If

End If
End Sub
Sub frm92_kiraan_gst()
'on error resume next
Dim Frm92_LM_KADAR_GST As Double
Dim Frm92_LM_HARGA As Double
Dim Frm92_LM_GST As Double

Frm92_LM_KADAR_GST = 0
Frm92_LM_HARGA = 0
Frm92_LM_GST = 0

If GLOBAL_DISABLE = 0 Then

    If (Frm92.TB2 <> vbNullString And IsNumeric(Frm92.TB2)) And (Frm92.L15_Text <> vbNullString And IsNumeric(Frm92.L15_Text)) Then

        If IsNumeric(Frm92.L15_Text) Then Frm92_LM_KADAR_GST = Frm92.L15_Text 'Jumlah Kadar GST (%)
        If IsNumeric(Frm92.TB2) Then Frm92_LM_HARGA = Frm92.TB2 'Jumlah Bayaran (RM)
        
        If Frm92.CB1 = 1 Then
        
            Frm92_LM_GST = 0
            
            Frm92.L8_Text = Format(Frm92_LM_GST, "#,##0.00") 'Jumlah cukai GST
            Frm92.L50_Text = Format(Frm92_LM_HARGA + Frm92_LM_GST, "#,##0.00") 'Jumlah harga tanpa GST
            Frm92.L55_Text = Format(Frm92_LM_HARGA + Frm92_LM_GST, "#,##0.00") 'Jumlah harga dengan GST
            
        ElseIf Frm92.CB2 = 1 Then
        
            Frm92_LM_GST = Frm92_LM_HARGA * (Frm92_LM_KADAR_GST / 100)
            
            Frm92.L8_Text = Format(Frm92_LM_GST, "#,##0.00") 'Jumlah cukai GST
            Frm92.L50_Text = Format(Frm92_LM_HARGA, "#,##0.00")  'Jumlah harga tanpa GST
            Frm92.L55_Text = Format(Frm92_LM_HARGA + Frm92_LM_GST, "#,##0.00") 'Jumlah harga dengan GST
        
        ElseIf Frm92.CB8 = 1 Then
        
            Frm92_LM_GST = Frm92_LM_HARGA - (Frm92_LM_HARGA / (1 + (Frm92_LM_KADAR_GST / 100)))
            
            Frm92.L8_Text = Format(Frm92_LM_GST, "#,##0.00") 'Jumlah cukai GST
            Frm92.L50_Text = Format((Frm92_LM_HARGA / (1 + (Frm92_LM_KADAR_GST / 100))), "#,##0.00")  'Jumlah harga tanpa GST
            Frm92.L55_Text = Format(Frm92_LM_HARGA, "#,##0.00")  'Jumlah harga dengan GST
        
        Else
        
            Frm92.L8_Text = Format(0, "#,##0.00") 'Jumlah cukai GST
            Frm92.L50_Text = Format(0, "#,##0.00")  'Jumlah harga tanpa GST
            Frm92.L55_Text = Format(0, "#,##0.00") 'Jumlah harga dengan GST
        
        End If
        
        
    Else
    
        Frm92.L8_Text = Format(0, "#,##0.00") 'Jumlah cukai GST
        Frm92.L50_Text = Format(0, "#,##0.00")  'Jumlah harga tanpa GST
        Frm92.L55_Text = Format(0, "#,##0.00") 'Jumlah harga dengan GST
        
    End If
    
End If
End Sub
Sub frm92_kiraan_harga_belanja()
'on error resume next
Dim LM_HARGA_SR As Double
Dim LM_CUKAI_SR As Double
Dim LM_HARGA_ZR As Double
Dim LM_CUKAI_ZR As Double

LM_HARGA_SR = 0
LM_CUKAI_SR = 0
LM_HARGA_ZR = 0
LM_CUKAI_ZR = 0

If Frm92.TB46 <> vbNullString And IsNumeric(Frm92.TB46) Then LM_HARGA_SR = Frm92.TB46
If Frm92.TB47 <> vbNullString And IsNumeric(Frm92.TB47) Then LM_CUKAI_SR = Frm92.TB47
If Frm92.TB48 <> vbNullString And IsNumeric(Frm92.TB48) Then LM_HARGA_ZR = Frm92.TB48
If Frm92.TB49 <> vbNullString And IsNumeric(Frm92.TB49) Then LM_CUKAI_ZR = Frm92.TB49

Frm92.TB45 = Format(LM_HARGA_SR + LM_CUKAI_SR + LM_HARGA_ZR + LM_CUKAI_ZR, "#,##0.00")
End Sub
Sub frm92_kiraan_cukai_sr_belanja()
'on error resume next
Dim LM_HARGA_SR As Double
Dim LM_KADAR As Double

LM_HARGA_SR = 0
LM_KADAR = 0

If Frm92.TB46 <> vbNullString And IsNumeric(Frm92.TB46) Then LM_HARGA_SR = Frm92.TB46
If Frm92.L42_Text <> vbNullString And IsNumeric(Frm92.L42_Text) Then LM_KADAR = Frm92.L42_Text

Frm92.TB47 = Format(((LM_KADAR / 100) * LM_HARGA_SR), "#,##0.00")
End Sub
Sub frm92_kiraan_cukai_zr_belanja()
'on error resume next
Dim LM_HARGA_ZR As Double
Dim LM_KADAR As Double

LM_HARGA_ZR = 0
LM_KADAR = 0

If Frm92.TB48 <> vbNullString And IsNumeric(Frm92.TB48) Then LM_HARGA_ZR = Frm92.TB48
If Frm92.L42_Text <> vbNullString And IsNumeric(Frm92.L42_Text) Then LM_KADAR = Frm92.L42_Text

LM_KADAR = 0

Frm92.TB49 = Format(((LM_KADAR / 100) * LM_HARGA_ZR), "#,##0.00")
End Sub
Sub frm92_initial_one_time()
'on error resume next
Frm92.CBB5.Clear
Frm92.CBB7.Clear
Frm92.CBB8.Clear
Frm92.CBB9.Clear

Frm92.CBB8.AddItem "Semua jenis"

Set rs3 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs3.Open "select * from setting_database where (Supplier <> '" & Null & "' OR jenis_expense <> '" & Null & "') AND status='" & 1 & "'", cn, adOpenKeyset, adLockOptimistic

While rs3.EOF = False
    If Not IsNull(rs3!supplier) Then Frm92.CBB5.AddItem rs3!supplier
    If Not IsNull(rs3!jenis_expense) Then Frm92.CBB7.AddItem rs3!jenis_expense
    If Not IsNull(rs3!jenis_expense) Then Frm92.CBB8.AddItem rs3!jenis_expense
    rs3.MoveNext
Wend

rs3.Close
Set rs3 = Nothing

Frm92.CBB5.AddItem "Lain-lain"

Frm92.CBB5 = "Lain-lain"

Frm92.CBB8 = "Semua jenis"

Frm92.CBB9.AddItem "Semua cawangan"

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where status='" & 1 & "' order by cawangan ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!cawangan) Then Frm92.CBB9.AddItem rs!cawangan
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm92.CBB9 = "Semua cawangan"

If MDI_frm1.L4_Text <> "HQ" And MDI_frm1.L4_Text <> "Developer" Then

    Frm92.CBB9 = MDI_frm1.L20_Text
    Frm92.CBB9.Enabled = False
    
Else
    
    Frm92.CBB9.Enabled = True
    
End If
End Sub

Sub frm92_cetak_pv()
'on error resume next
DATA_FOUND = 0
Frm84_DATA_PEKERJA_FOUND = 0


Report81.Sections("Section2").Controls("L5").Caption = vbNullString 'Nama kedai
Report81.Sections("Section2").Controls("L7").Caption = "No. ID GST : -" 'No ID GST
Report81.Sections("Section2").Controls("L3").Caption = vbNullString 'No. Payment Voucher
Report81.Sections("Section2").Controls("L4").Caption = vbNullString 'Date
Report81.Sections("Section2").Controls("L17").Caption = vbNullString 'Staff

Report81.Sections("Section1").Controls("L14").Caption = vbNullString
Report81.Sections("Section1").Controls("L8").Caption = "1" 'Kuantiti
Report81.Sections("Section1").Controls("L9").Caption = "0.00" 'Unit price
Report81.Sections("Section1").Controls("L10").Caption = "0.00" 'Total

Report81.Sections("Section1").Controls("L11").Caption = "0.00" 'Exclude GST
Report81.Sections("Section1").Controls("L12").Caption = "0.00" 'GST
Report81.Sections("Section1").Controls("L13").Caption = "0.00" 'Include GST

'### Reset maklumat kedai ### - Start
Report81.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report81.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report81.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report81.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report81.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report81.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report81.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report81.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report81.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report81.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 39_akaun_expense where status = 1 AND no_voucher='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Not IsNull(rs!nama_kedai) Then Report81.Sections("Section2").Controls("L5").Caption = rs!nama_kedai
    If Not IsNull(rs!no_id_gst) Then Report81.Sections("Section2").Controls("L7").Caption = "No. ID GST : " & rs!no_id_gst
    Report81.Sections("Section2").Controls("L3").Caption = G_No_RESIT_JUALAN
    If Not IsNull(rs!tarikh) Then Report81.Sections("Section2").Controls("L4").Caption = rs!tarikh
    If Not IsNull(rs!tujuan) Then Report81.Sections("Section1").Controls("L14").Caption = rs!tujuan
    If Not IsNull(rs!jumlah_tanpa_gst) Then Report81.Sections("Section1").Controls("L9").Caption = Format(rs!harga_dengan_gst - rs!gst_sr_cukai, "#,##0.00")
    If Not IsNull(rs!jumlah_tanpa_gst) Then Report81.Sections("Section1").Controls("L10").Caption = Format(rs!harga_dengan_gst - rs!gst_sr_cukai, "#,##0.00")
    If Not IsNull(rs!jumlah_tanpa_gst) Then Report81.Sections("Section1").Controls("L11").Caption = Format(rs!harga_dengan_gst - rs!gst_sr_cukai, "#,##0.00")
    If Not IsNull(rs!gst_sr_cukai) Then Report81.Sections("Section1").Controls("L12").Caption = Format(rs!gst_sr_cukai, "#,##0.00")
    If Not IsNull(rs!harga_dengan_gst) Then Report81.Sections("Section1").Controls("L13").Caption = Format(rs!harga_dengan_gst, "#,##0.00")

    If Not IsNull(rs!no_pekerja) Then
        Frm84_LM_No_PEKERJA = rs!no_pekerja
        Frm84_DATA_PEKERJA_FOUND = 1
    End If

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
            If Not IsNull(rs!Samaran) Then Report81.Sections("Section2").Controls("L17").Caption = rs!Samaran 'Nama Samaran
        End If

        rs.Close
        Set rs = Nothing
    End If

    '### Paparan Resit ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 39_akaun_expense where status = 1 AND no_voucher='" & G_No_RESIT_JUALAN & "'", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        Set Report81.DataSource = rs
        If G_PREVIEW = 1 Then Report81.Show
        rs.MoveNext
    Wend
    
    'rs.Close
    Set rs = Nothing
    '### Paparan Resit ### - End
    
    If G_PREVIEW = 0 Then Report81.PrintReport
     
    G_No_RESIT_JUALAN = vbNullString
End If
End Sub

