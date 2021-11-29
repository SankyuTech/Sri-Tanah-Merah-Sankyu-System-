Attribute VB_Name = "Module37"
Sub Frm75_Initial_Setting()
'on error resume next
Frm75.Pic1.Visible = False
Frm75.Pic2.Visible = False
Frm75.Pic3.Visible = False

Frm75.Pic1.Left = 120
Frm75.Pic1.Top = 480
Frm75.Pic2.Left = 120
Frm75.Pic2.Top = 480
Frm75.Pic3.Left = 120
Frm75.Pic3.Top = 480

Frm75.L8_Text.BackStyle = 0
Frm75.L10_Text.BackStyle = 0

Frm75.L12_Text.BackStyle = 0
Frm75.L13_Text.BackStyle = 0
Frm75.L14_Text.BackStyle = 0
Frm75.L15_Text.BackStyle = 0
Frm75.L17_Text.BackStyle = 0
Frm75.L18_Text.BackStyle = 0
Frm75.L19_Text.BackStyle = 0
End Sub
Sub Frm75_Report_GST_Header()
'on error resume next
Frm75.MSFlexGrid1.Clear
Frm75.MSFlexGrid1.RowHeight(0) = 800
Frm75.MSFlexGrid1.FormatString = "No.|<No.|<No. ID|<Tarikh|<No. Resit|<No. ID GST|<Perkara|<Jumlah Harga SR (RM)|<Jumlah Cukai SR (RM)|<Jumlah Harga ZR(L) (RM)|<Jumlah Cukai ZR (RM)"

Frm75.MSFlexGrid1.Rows = 1
Frm75.MSFlexGrid1.ColWidth(0) = 600
Frm75.MSFlexGrid1.ColWidth(1) = 0
Frm75.MSFlexGrid1.ColWidth(2) = 0 'No. ID
Frm75.MSFlexGrid1.ColWidth(3) = 1000 'Tarikh
Frm75.MSFlexGrid1.ColWidth(4) = 1500 'No. Resit
Frm75.MSFlexGrid1.ColWidth(5) = 1500 'No. ID GST
Frm75.MSFlexGrid1.ColWidth(6) = 4800 'Perkara
Frm75.MSFlexGrid1.ColWidth(7) = 1700 'Jumlah Harga SR (RM)
Frm75.MSFlexGrid1.ColWidth(8) = 1700 'Jumlah Cukai SR (RM)
Frm75.MSFlexGrid1.ColWidth(9) = 1700 'Jumlah Harga ZR(L) (RM)
Frm75.MSFlexGrid1.ColWidth(10) = 1700 'Jumlah Cukai SR (RM)

Frm75.MSFlexGrid3.Clear
Frm75.MSFlexGrid3.RowHeight(0) = 800
Frm75.MSFlexGrid3.FormatString = "No.|<No.|<No. ID|<Tarikh|<No. Resit|<Perkara|<Jumlah Harga SR (RM)|<Jumlah Cukai SR (RM)|<Jumlah Harga ZR(L) (RM)|<Jumlah Cukai ZR (RM)"

Frm75.MSFlexGrid3.Rows = 1
Frm75.MSFlexGrid3.ColWidth(0) = 600
Frm75.MSFlexGrid3.ColWidth(1) = 0
Frm75.MSFlexGrid3.ColWidth(2) = 0 'No. ID
Frm75.MSFlexGrid3.ColWidth(3) = 1500 'Tarikh
Frm75.MSFlexGrid3.ColWidth(4) = 1800 'No. Resit
Frm75.MSFlexGrid3.ColWidth(5) = 5000 'Perkara
Frm75.MSFlexGrid3.ColWidth(6) = 1800 'Jumlah Harga SR (RM)
Frm75.MSFlexGrid3.ColWidth(7) = 1800 'Jumlah Cukai SR (RM)
Frm75.MSFlexGrid3.ColWidth(8) = 1800 'Jumlah Harga ZR(L) (RM)
Frm75.MSFlexGrid3.ColWidth(9) = 1800 'Jumlah Cukai SR (RM)
End Sub
Sub Frm75_report_gst_kutip_header()
'on error resume next
Frm75.MSFlexGrid3.Clear
Frm75.MSFlexGrid3.RowHeight(0) = 800
Frm75.MSFlexGrid3.FormatString = "No.|<No.|<No. ID|<Tarikh|<No. Invoice|<Perkara|<Jumlah Harga SR (RM)|<Jumlah Cukai SR (RM)|<Jumlah Harga ZR (RM)|<Jumlah Cukai ZR (RM)"

Frm75.MSFlexGrid3.Rows = 1
Frm75.MSFlexGrid3.ColWidth(0) = 700
Frm75.MSFlexGrid3.ColAlignment(0) = 4

Frm75.MSFlexGrid3.ColWidth(1) = 0
Frm75.MSFlexGrid3.ColWidth(2) = 0 'No. ID

Frm75.MSFlexGrid3.ColWidth(3) = 1500 'Tarikh
Frm75.MSFlexGrid3.ColAlignment(3) = 4

Frm75.MSFlexGrid3.ColWidth(4) = 1500 'No. Invoice

Frm75.MSFlexGrid3.ColWidth(5) = 5000 'Perkara

Frm75.MSFlexGrid3.ColWidth(6) = 1300 'Jumlah Harga SR (RM)
Frm75.MSFlexGrid3.ColAlignment(6) = 7

Frm75.MSFlexGrid3.ColWidth(7) = 1300 'Jumlah Cukai SR (RM)
Frm75.MSFlexGrid3.ColAlignment(7) = 7

Frm75.MSFlexGrid3.ColWidth(8) = 1300 'Jumlah Harga ZR(L) (RM)
Frm75.MSFlexGrid3.ColAlignment(8) = 7

Frm75.MSFlexGrid3.ColWidth(9) = 1300 'Jumlah Cukai SR (RM)
Frm75.MSFlexGrid3.ColAlignment(9) = 7
End Sub
Sub Frm75_report_gst_kutip()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim frm75_LM_TOTAL_PAGE As Double

Frm75_PAGE_SIZE = 38
frm75_LM_TOTAL_PAGE = 0
x = 0

Frm75.L14_Text = Format(0, "#,##0.00")
Frm75.L15_Text = Format(0, "#,##0.00")
Frm75.L22_Text = Format(0, "#,##0.00")
Frm75.L23_Text = Format(0, "#,##0.00")

TM = Frm75.L6_Text 'Tarikh Mula
TA = Frm75.L7_Text 'Tarikh Akhir

user_level = MDI_frm1.L4_Text

LM_INVOICE_RASMI = 0

Frm110_LM_SEARCH_1 = 0
Frm110_LM_SEARCH_2 = 1

If user_level = "Guest/User" Then
    Frm85_LM_SEARCH_6 = 1
    Frm85_LM_SEARCH_6_LOGIC = "="
    LM_INVOICE_RASMI = 1
    Frm85_LM_SEARCH_7 = 1
    Frm85_LM_SEARCH_7_LOGIC = "="
    
Else
    Frm85_LM_SEARCH_6 = 0
    Frm85_LM_SEARCH_6_LOGIC = "="
    
    Frm85_LM_SEARCH_7 = 0
    Frm85_LM_SEARCH_7_LOGIC = "="
End If

If user_level = "Administration" Then

    Frm110_LM_SEARCH_1 = 1
    Frm110_LM_SEARCH_2 = 1
    
End If

Frm75.L10_Text = "Report terperinci kutipan cukai GST dari " & TM & " hingga " & TA & "." 'Header Report

re_gen_report:

LM_START_ROW = Frm75.L69_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm75_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm75.L75_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm75_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        Frm75.L67_Text = 1
    End If
End If

Frm75_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND status = 1 AND (bil_rasmi='" & Frm110_LM_SEARCH_1 & "' OR bil_rasmi='" & Frm110_LM_SEARCH_2 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm75_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1
    If Frm75_LM_PAGE_FOUND = 0 Then
        If Frm75.L75_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm75.L67_Text = Frm75.L67_Text + 1 'Paparan Page ke-xxx
                Frm75_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm75.L67_Text) Then
                    If Frm75.L67_Text <> 1 Then
                        Frm75.L67_Text = Frm75.L67_Text - 1 'Paparan Page ke-xxx
                        Frm75_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If

    Y = ((Frm75.L67_Text - 1) * Frm75_PAGE_SIZE) + x
    Frm75.MSFlexGrid3.Rows = x + 1
    Frm75.MSFlexGrid3.TextMatrix(x, 0) = Y 'No.
    Frm75.MSFlexGrid3.TextMatrix(x, 1) = x 'No.
    Frm75.MSFlexGrid3.ColAlignment(1) = 4
    Frm75.MSFlexGrid3.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!tarikh) Then Frm75.MSFlexGrid3.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    'If Not IsNull(rs!no_resit) Then Frm75.MSFlexGrid3.TextMatrix(x, 4) = rs!no_resit 'No. Resit
    
    If LM_INVOICE_RASMI = 0 Then
        If Not IsNull(rs!no_resit) Then Frm75.MSFlexGrid3.TextMatrix(x, 4) = rs!no_resit 'No. Invoice
    Else
        If Not IsNull(rs!no_invoice_r) Then Frm75.MSFlexGrid3.TextMatrix(x, 4) = rs!no_invoice_r 'No. Invoice
    End If
    
    If Not IsNull(rs!Menu) Then 'Detail
        If rs!Menu = 0 Then Frm75.MSFlexGrid3.TextMatrix(x, 5) = "Jualan emas kepada pelanggan"
        If rs!Menu = 1 Then Frm75.MSFlexGrid3.TextMatrix(x, 5) = "Servis kepada pelanggan"
        If rs!Menu = 2 Then Frm75.MSFlexGrid3.TextMatrix(x, 5) = "Deposit tempahan emas"
        If rs!Menu = 3 Then Frm75.MSFlexGrid3.TextMatrix(x, 5) = "Tempahan siap"
        If rs!Menu = 4 Then Frm75.MSFlexGrid3.TextMatrix(x, 5) = "Jualan kepada agen/reseller"
    End If
    
    If Not IsNull(rs!gst_sr_harga) Then Frm75.MSFlexGrid3.TextMatrix(x, 6) = Format(rs!gst_sr_harga, "#,##0.00") 'Jumlah Harga SR(RM)
    If Not IsNull(rs!gst_sr_cukai) Then Frm75.MSFlexGrid3.TextMatrix(x, 7) = Format(rs!gst_sr_cukai, "#,##0.00") 'Jumlah Cukai SR(RM)
    If Not IsNull(rs!gst_zr_harga) Then Frm75.MSFlexGrid3.TextMatrix(x, 8) = Format(rs!gst_zr_harga, "#,##0.00") 'Jumlah Harga ZR(RM)
    If Not IsNull(rs!gst_zr_cukai) Then Frm75.MSFlexGrid3.TextMatrix(x, 9) = Format(rs!gst_zr_cukai, "#,##0.00") 'Jumlah Cukai SR(RM)
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 22_jualan where tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND status = 1 AND (bil_rasmi='" & Frm110_LM_SEARCH_1 & "' OR bil_rasmi='" & Frm110_LM_SEARCH_2 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    frm75_LM_TOTAL_PAGE = Format(rs(0) / Frm75_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, frm75_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm75_LM_PAGE = Split(frm75_LM_TOTAL_PAGE, ".")(0)
        Frm75_LM_PAGE_LEBIHAN = Split(frm75_LM_TOTAL_PAGE, ".")(1)
        
        If Frm75_LM_PAGE_LEBIHAN <> "00" Then
            Frm75.L68_Text = Frm75_LM_PAGE + 1
        Else
            Frm75.L68_Text = Frm75_LM_PAGE
        End If
        
    Else
    
        Frm75.L68_Text = frm75_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm75.L68_Text = 0
    End If
Else
    Frm75.L68_Text = 0
End If

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(gst_sr_harga) , SUM(gst_sr_cukai) , SUM(gst_zr_harga) , SUM(gst_zr_cukai) from 22_jualan where tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND status = 1 AND (bil_rasmi='" & Frm110_LM_SEARCH_1 & "' OR bil_rasmi='" & Frm110_LM_SEARCH_2 & "') AND (status_r " & Frm85_LM_SEARCH_6_LOGIC & "'" & Frm85_LM_SEARCH_6 & "' OR status_r " & Frm85_LM_SEARCH_7_LOGIC & "'" & Frm85_LM_SEARCH_7 & "') order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm75.L14_Text = Format(rs(0), "#,##0.00")
If Not IsNull(rs(1)) Then Frm75.L15_Text = Format(rs(1), "#,##0.00")
If Not IsNull(rs(2)) Then Frm75.L22_Text = Format(rs(2), "#,##0.00")
If Not IsNull(rs(3)) Then Frm75.L23_Text = Format(rs(3), "#,##0.00")

rs.Close
Set rs = Nothing

If x <> 0 Then
    Frm75.L69_Text = LM_START_ROW
End If

If Frm75.L67_Text <> vbNullString And IsNumeric(Frm75.L67_Text) Then
    If Frm75.L68_Text <> vbNullString And IsNumeric(Frm75.L68_Text) Then
        frm75_LM_CURR_PAGE = Frm75.L67_Text
        frm75_LM_TOTAL_PAGE = Frm75.L68_Text
        
        If frm75_LM_CURR_PAGE > frm75_LM_TOTAL_PAGE Then
            
            Frm75.L67_Text = Frm75.L67_Text - 1
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            GoTo re_gen_report:
            
        End If
    End If
End If
End Sub
Sub Frm75_report_gst_bayar_header()
'on error resume next
Frm75.MSFlexGrid1.Clear
Frm75.MSFlexGrid1.RowHeight(0) = 800
Frm75.MSFlexGrid1.FormatString = "No.|<No.|<No. ID|<Tarikh|<No. Invoice|<Nama Supplier/Kedai|<No. ID GST|<Perkara|<Jumlah Harga SR (RM)|<Jumlah Cukai SR (RM)|<Jumlah Harga ZR (RM)|<Jumlah Cukai ZR (RM)"

Frm75.MSFlexGrid1.Rows = 1
Frm75.MSFlexGrid1.ColWidth(0) = 800
Frm75.MSFlexGrid1.ColAlignment(0) = 4

Frm75.MSFlexGrid1.ColWidth(1) = 0
Frm75.MSFlexGrid1.ColWidth(2) = 0 'No. ID
Frm75.MSFlexGrid1.ColWidth(3) = 1200 'Tarikh
Frm75.MSFlexGrid1.ColAlignment(3) = 4

Frm75.MSFlexGrid1.ColWidth(4) = 1500 'No. Invoice

Frm75.MSFlexGrid1.ColWidth(5) = 4600 'Nama Supplier/Kedai

Frm75.MSFlexGrid1.ColWidth(6) = 1500 'No. ID GST

Frm75.MSFlexGrid1.ColWidth(7) = 4500 'Perkara

Frm75.MSFlexGrid1.ColWidth(8) = 1300 'Jumlah Harga SR (RM)
Frm75.MSFlexGrid1.ColAlignment(8) = 7

Frm75.MSFlexGrid1.ColWidth(9) = 1300 'Jumlah Cukai SR (RM)
Frm75.MSFlexGrid1.ColAlignment(9) = 7

Frm75.MSFlexGrid1.ColWidth(10) = 1300 'Jumlah Harga ZR (RM)
Frm75.MSFlexGrid1.ColAlignment(10) = 7

Frm75.MSFlexGrid1.ColWidth(11) = 1300 'Jumlah Cukai SR (RM)
Frm75.MSFlexGrid1.ColAlignment(11) = 7
End Sub
Sub Frm75_report_gst_bayar()
'on error resume next
Dim frm75_LM_TOTAL_PAGE As Double
Dim TM As Date
Dim TA As Date

frm75_LM_TOTAL_PAGE = 0
x = 0
Y = 0
Frm75_PAGE_SIZE = 38

Frm75.L12_Text = Format(0, "#,##0.00")
Frm75.L13_Text = Format(0, "#,##0.00")
Frm75.L20_Text = Format(0, "#,##0.00")
Frm75.L21_Text = Format(0, "#,##0.00")

TM = Frm75.L6_Text 'Tarikh Mula
TA = Frm75.L7_Text 'Tarikh Akhir

Frm75.L8_Text = "Report terperinci bayaran cukai GST dari " & TM & " hingga " & TA & "(Perbelanjaan kedai)" 'Header Report

LM_START_ROW = Frm75.L62_Text

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm75_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm75.L63_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm75_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        Frm75.L60_Text = 1
    End If
End If

Frm75_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 39_akaun_expense where status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm75_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm75_LM_PAGE_FOUND = 0 Then
        If Frm75.L63_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm75.L60_Text = Frm75.L60_Text + 1
                Frm75_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm75.L60_Text) Then
                    If Frm75.L60_Text <> 1 Then
                        Frm75.L60_Text = Frm75.L60_Text - 1
                        Frm75_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm75.L60_Text - 1) * Frm75_PAGE_SIZE) + x
        
    Frm75.MSFlexGrid1.Rows = x + 1
    Frm75.MSFlexGrid1.TextMatrix(x, 0) = Y 'No.
    Frm75.MSFlexGrid1.TextMatrix(x, 1) = x 'No.
    If Not IsNull(rs!ID) Then Frm75.MSFlexGrid1.TextMatrix(x, 2) = rs!ID 'No. ID
    
    If Not IsNull(rs!tarikh) Then Frm75.MSFlexGrid1.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    
    If Not IsNull(rs!no_resit) Then Frm75.MSFlexGrid1.TextMatrix(x, 4) = rs!no_resit 'No. Invoice
    
    If Not IsNull(rs!nama_kedai) Then Frm75.MSFlexGrid1.TextMatrix(x, 5) = rs!nama_kedai 'Nama Supplier/Kedai
    
    If Not IsNull(rs!no_id_gst) Then Frm75.MSFlexGrid1.TextMatrix(x, 6) = rs!no_id_gst 'No. GST ID Supplier
    
    If Not IsNull(rs!tujuan) Then Frm75.MSFlexGrid1.TextMatrix(x, 7) = rs!tujuan 'Detail
    
    If Not IsNull(rs!gst_sr_harga) Then Frm75.MSFlexGrid1.TextMatrix(x, 8) = Format(rs!gst_sr_harga, "#,##0.00") 'Jumlah Harga SR(RM)

    If Not IsNull(rs!gst_sr_cukai) Then Frm75.MSFlexGrid1.TextMatrix(x, 9) = Format(rs!gst_sr_cukai, "#,##0.00") 'Jumlah Cukai SR(RM)

    If Not IsNull(rs!gst_zr_harga) Then Frm75.MSFlexGrid1.TextMatrix(x, 10) = Format(rs!gst_zr_harga, "#,##0.00") 'Jumlah Harga ZR(RM)
        
    If Not IsNull(rs!gst_zr_cukai) Then Frm75.MSFlexGrid1.TextMatrix(x, 11) = Format(rs!gst_zr_cukai, "#,##0.00") 'Jumlah Cukai SR(RM)
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 39_akaun_expense where status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs(0)) Then
        frm75_LM_TOTAL_PAGE = Format(rs(0) / Frm75_PAGE_SIZE, "0.00") 'Jumlah Page
        
        'Periksa Samada ada titik perpuluhan atau tidak
        If InStr(1, frm75_LM_TOTAL_PAGE, ".") <> 0 Then
        
            Frm85_LM_PAGE = Split(frm75_LM_TOTAL_PAGE, ".")(0)
            Frm85_LM_PAGE_LEBIHAN = Split(frm75_LM_TOTAL_PAGE, ".")(1)
            
            If Frm85_LM_PAGE_LEBIHAN <> "00" Then
                Frm75.L61_Text = Frm85_LM_PAGE + 1 'Total Page
            Else
                Frm75.L61_Text = Frm85_LM_PAGE
            End If
            
        Else
        
            Frm75.L61_Text = frm75_LM_TOTAL_PAGE
            
        End If
    
        If rs(0) = vbNullString Then
            Frm75.L61_Text = 0
        End If
    End If
Else
    Frm75.L61_Text = 0
End If

rs.Close
Set rs = Nothing

If Frm75.L61_Text = vbNullString Then
    Frm75.L61_Text = 0
End If
'### Jumlah Data ### - End

'#### Jumlah Data #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(gst_sr_harga) , SUM(gst_sr_cukai) , SUM(gst_zr_harga) , SUM(gst_zr_cukai) from 39_akaun_expense where status = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm75.L12_Text = Format(rs(0), "#,##0.00")
If Not IsNull(rs(1)) Then Frm75.L13_Text = Format(rs(1), "#,##0.00")
If Not IsNull(rs(2)) Then Frm75.L20_Text = Format(rs(2), "#,##0.00")
If Not IsNull(rs(3)) Then Frm75.L21_Text = Format(rs(3), "#,##0.00")

rs.Close
Set rs = Nothing
'#### Jumlah Data#### - End

If x <> 0 Then
    Frm75.L63_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    Frm75.L62_Text = LM_START_ROW
Else
    Frm75.L63_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If
End Sub
Sub Frm75_Report_GST_BAYARAN()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim Frm75_JUMLAH_HARGA_SR As Double
Dim Frm75_JUMLAH_CUKAI_SR As Double
Dim Frm75_JUMLAH_HARGA_ZR As Double
Dim Frm75_JUMLAH_CUKAI_ZR As Double

Frm75_JUMLAH_HARGA_SR = 0
Frm75_JUMLAH_CUKAI_SR = 0
Frm75_JUMLAH_HARGA_ZR = 0
Frm75_JUMLAH_CUKAI_ZR = 0

TM = Frm75.L6_Text 'Tarikh Mula
TA = Frm75.L7_Text 'Tarikh Akhir

Frm75.L8_Text = "Report Terperinci Bayaran GST Dari " & TM & " Hingga " & TA 'Header Report

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 77_gdn_grn where status = 1 AND jenis_urusan = 3 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm75.MSFlexGrid1.Rows = x + 1
    Frm75.MSFlexGrid1.TextMatrix(x, 0) = x
    Frm75.MSFlexGrid1.TextMatrix(x, 1) = x 'No.
    If Not IsNull(rs!ID) Then Frm75.MSFlexGrid1.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm75.MSFlexGrid1.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!no_rujukan_supplier) Then Frm75.MSFlexGrid1.TextMatrix(x, 4) = rs!no_rujukan_supplier 'No. Resit Dari Supplier
    'If Not IsNull(rs!no_id_gst_supplier) Then Frm75.MSFlexGrid1.TextMatrix(x, 5) = rs!no_id_gst_supplier 'No. GST ID Supplier
    Frm75.MSFlexGrid1.TextMatrix(x, 6) = "Belian Stok Emas Kedai" 'Detail
    If Not IsNull(rs!gst_sr_harga) Then
        Frm75.MSFlexGrid1.TextMatrix(x, 7) = Format(rs!gst_sr_harga, "#,##0.00") 'Jumlah Harga SR(RM)
        If IsNumeric(rs!gst_sr_harga) Then Frm75_JUMLAH_HARGA_SR = Frm75_JUMLAH_HARGA_SR + rs!gst_sr_harga
    End If
    If Not IsNull(rs!gst_sr_cukai) Then
        Frm75.MSFlexGrid1.TextMatrix(x, 8) = Format(rs!gst_sr_cukai, "#,##0.00") 'Jumlah Cukai SR(RM)
        If IsNumeric(rs!gst_sr_cukai) Then Frm75_JUMLAH_CUKAI_SR = Frm75_JUMLAH_CUKAI_SR + rs!gst_sr_cukai
    End If
    If Not IsNull(rs!gst_zr_harga) Then
        Frm75.MSFlexGrid1.TextMatrix(x, 9) = Format(rs!gst_zr_harga, "#,##0.00") 'Jumlah Harga ZR(RM)
        If IsNumeric(rs!gst_zr_harga) Then Frm75_JUMLAH_HARGA_ZR = Frm75_JUMLAH_HARGA_ZR + rs!gst_zr_harga
    End If
    If Not IsNull(rs!gst_zr_cukai) Then
        Frm75.MSFlexGrid1.TextMatrix(x, 10) = Format(rs!gst_zr_cukai, "#,##0.00") 'Jumlah Cukai SR(RM)
        If IsNumeric(rs!gst_zr_cukai) Then Frm75_JUMLAH_CUKAI_ZR = Frm75_JUMLAH_CUKAI_ZR + rs!gst_zr_cukai
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Kutipan GST Dari Perbelanjaan Kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 39_akaun_expense where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm75.MSFlexGrid1.Rows = x + 1
    Frm75.MSFlexGrid1.TextMatrix(x, 0) = x
    Frm75.MSFlexGrid1.TextMatrix(x, 1) = x 'No.
    If Not IsNull(rs!ID) Then Frm75.MSFlexGrid1.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm75.MSFlexGrid1.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!no_resit) Then Frm75.MSFlexGrid1.TextMatrix(x, 4) = rs!no_resit 'No. Resit Dari Supplier
    If Not IsNull(rs!no_id_gst) Then Frm75.MSFlexGrid1.TextMatrix(x, 5) = rs!no_id_gst 'No. GST ID Supplier
    Frm75.MSFlexGrid1.TextMatrix(x, 6) = "Perbelanjaan Kedai" 'Detail
    If Not IsNull(rs!gst_sr_harga) Then
        Frm75.MSFlexGrid1.TextMatrix(x, 7) = Format(rs!gst_sr_harga, "0.00") 'Jumlah Harga SR(RM)
        If IsNumeric(rs!gst_sr_harga) Then Frm75_JUMLAH_HARGA_SR = Frm75_JUMLAH_HARGA_SR + rs!gst_sr_harga
    End If
    If Not IsNull(rs!gst_sr_cukai) Then
        Frm75.MSFlexGrid1.TextMatrix(x, 8) = Format(rs!gst_sr_cukai, "0.00") 'Jumlah Cukai SR(RM)
        If IsNumeric(rs!gst_sr_cukai) Then Frm75_JUMLAH_CUKAI_SR = Frm75_JUMLAH_CUKAI_SR + rs!gst_sr_cukai
    End If
    If Not IsNull(rs!gst_zr_harga) Then
        Frm75.MSFlexGrid1.TextMatrix(x, 9) = Format(rs!gst_zr_harga, "0.00") 'Jumlah Harga ZR(RM)
        If IsNumeric(rs!gst_zr_harga) Then Frm75_JUMLAH_HARGA_ZR = Frm75_JUMLAH_HARGA_ZR + rs!gst_zr_harga
    End If
    If Not IsNull(rs!gst_zr_cukai) Then
        Frm75.MSFlexGrid1.TextMatrix(x, 10) = Format(rs!gst_zr_cukai, "0.00") 'Jumlah Cukai SR(RM)
        If IsNumeric(rs!gst_zr_cukai) Then Frm75_JUMLAH_CUKAI_ZR = Frm75_JUMLAH_CUKAI_ZR + rs!gst_zr_cukai
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'### Kutipan GST Dari Perbelanjaan Kedai ### - End

Frm75.L12_Text = Format(Frm75_JUMLAH_HARGA_SR, "0.00") 'jumlah SR
Frm75.L13_Text = Format(Frm75_JUMLAH_CUKAI_SR, "0.00") 'Cukai SR
Frm75.L20_Text = Format(Frm75_JUMLAH_HARGA_ZR, "0.00") 'Harga ZR
Frm75.L21_Text = Format(Frm75_JUMLAH_CUKAI_ZR, "0.00") 'Cukai ZR
End Sub
Sub Frm75_Report_GST_KUTIPAN()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim Frm75_JUMLAH_HARGA_SR As Double
Dim Frm75_JUMLAH_CUKAI_SR As Double
Dim Frm75_JUMLAH_HARGA_ZR As Double
Dim Frm75_JUMLAH_CUKAI_ZR As Double

Frm75_JUMLAH_HARGA_SR = 0
Frm75_JUMLAH_CUKAI_SR = 0
Frm75_JUMLAH_HARGA_ZR = 0
Frm75_JUMLAH_CUKAI_ZR = 0

TM = Frm75.L6_Text 'Tarikh Mula
TA = Frm75.L7_Text 'Tarikh Akhir

Frm75.L10_Text = "Report Terperinci Kutipan GST Dari " & TM & " Hingga " & TA 'Header Report

'### Kutipan GST Dari Jualan Barang Kemas ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 22_jualan where tarikh BETWEEN '" & TM & "' AND '" & TA & "' AND status = 1 AND bil_rasmi = 1 order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm75.MSFlexGrid3.Rows = x + 1
    Frm75.MSFlexGrid3.TextMatrix(x, 0) = x
    Frm75.MSFlexGrid3.TextMatrix(x, 1) = x 'No.
    If Not IsNull(rs!ID) Then Frm75.MSFlexGrid3.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm75.MSFlexGrid3.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!no_resit) Then Frm75.MSFlexGrid3.TextMatrix(x, 4) = rs!no_resit 'No. Resit
    Frm75.MSFlexGrid3.TextMatrix(x, 5) = "Jualan Emas Kedai Kepada Pelanggan" 'Detail
    If Not IsNull(rs!gst_sr_harga) Then
        Frm75.MSFlexGrid3.TextMatrix(x, 6) = Format(rs!gst_sr_harga, "0.00") 'Jumlah Harga SR(RM)
        If IsNumeric(rs!gst_sr_harga) Then Frm75_JUMLAH_HARGA_SR = Frm75_JUMLAH_HARGA_SR + rs!gst_sr_harga
    End If
    If Not IsNull(rs!gst_sr_cukai) Then
        Frm75.MSFlexGrid3.TextMatrix(x, 7) = Format(rs!gst_sr_cukai, "0.00") 'Jumlah Cukai SR(RM)
        If IsNumeric(rs!gst_sr_cukai) Then Frm75_JUMLAH_CUKAI_SR = Frm75_JUMLAH_CUKAI_SR + rs!gst_sr_cukai
    End If
    If Not IsNull(rs!gst_zr_harga) Then
        Frm75.MSFlexGrid3.TextMatrix(x, 8) = Format(rs!gst_zr_harga, "0.00") 'Jumlah Harga ZR(RM)
        If IsNumeric(rs!gst_zr_harga) Then Frm75_JUMLAH_HARGA_ZR = Frm75_JUMLAH_HARGA_ZR + rs!gst_zr_harga
    End If
    If Not IsNull(rs!gst_zr_cukai) Then
        Frm75.MSFlexGrid3.TextMatrix(x, 9) = Format(rs!gst_zr_cukai, "0.00") 'Jumlah Cukai SR(RM)
        If IsNumeric(rs!gst_zr_cukai) Then Frm75_JUMLAH_CUKAI_ZR = Frm75_JUMLAH_CUKAI_ZR + rs!gst_zr_cukai
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'### Kutipan GST Dari Jualan Barang Kemas ### - End

'### Kutipan GST Dari Servis ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 36_akaun_servis where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm75.MSFlexGrid3.Rows = x + 1
    Frm75.MSFlexGrid3.TextMatrix(x, 0) = x
    Frm75.MSFlexGrid3.TextMatrix(x, 1) = x 'No.
    If Not IsNull(rs!ID) Then Frm75.MSFlexGrid3.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm75.MSFlexGrid3.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!no_resit_servis) Then Frm75.MSFlexGrid3.TextMatrix(x, 4) = rs!no_resit_servis 'No. Resit
    Frm75.MSFlexGrid3.TextMatrix(x, 5) = "Servis Kepada Pelanggan" 'Detail
    If Not IsNull(rs!gst_sr_harga) Then
        Frm75.MSFlexGrid3.TextMatrix(x, 6) = Format(rs!gst_sr_harga, "0.00") 'Jumlah Harga SR(RM)
        If IsNumeric(rs!gst_sr_harga) Then Frm75_JUMLAH_HARGA_SR = Frm75_JUMLAH_HARGA_SR + rs!gst_sr_harga
    End If
    If Not IsNull(rs!gst_sr_cukai) Then
        Frm75.MSFlexGrid3.TextMatrix(x, 7) = Format(rs!gst_sr_cukai, "0.00") 'Jumlah Cukai SR(RM)
        If IsNumeric(rs!gst_sr_cukai) Then Frm75_JUMLAH_CUKAI_SR = Frm75_JUMLAH_CUKAI_SR + rs!gst_sr_cukai
    End If
    If Not IsNull(rs!gst_zr_harga) Then
        Frm75.MSFlexGrid3.TextMatrix(x, 8) = Format(rs!gst_zr_harga, "0.00") 'Jumlah Harga ZR(RM)
        If IsNumeric(rs!gst_zr_harga) Then Frm75_JUMLAH_HARGA_ZR = Frm75_JUMLAH_HARGA_ZR + rs!gst_zr_harga
    End If
    If Not IsNull(rs!gst_zr_cukai) Then
        Frm75.MSFlexGrid3.TextMatrix(x, 9) = Format(rs!gst_zr_cukai, "0.00") 'Jumlah Cukai SR(RM)
        If IsNumeric(rs!gst_zr_cukai) Then Frm75_JUMLAH_CUKAI_ZR = Frm75_JUMLAH_CUKAI_ZR + rs!gst_zr_cukai
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'### Kutipan GST Dari Servis ### - End

'### Kutipan GST Dari Ansuran ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 29_akaun_ansuran where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm75.MSFlexGrid3.Rows = x + 1
    Frm75.MSFlexGrid3.TextMatrix(x, 0) = x
    Frm75.MSFlexGrid3.TextMatrix(x, 1) = x 'No.
    If Not IsNull(rs!ID) Then Frm75.MSFlexGrid3.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm75.MSFlexGrid3.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!no_resit) Then Frm75.MSFlexGrid3.TextMatrix(x, 4) = rs!no_resit 'No. Resit
    Frm75.MSFlexGrid3.TextMatrix(x, 5) = "Bayaran Ansuran Belian Emas" 'Detail
    If Not IsNull(rs!gst_sr_harga) Then
        Frm75.MSFlexGrid3.TextMatrix(x, 6) = Format(rs!gst_sr_harga, "0.00") 'Jumlah Harga SR(RM)
        If IsNumeric(rs!gst_sr_harga) Then Frm75_JUMLAH_HARGA_SR = Frm75_JUMLAH_HARGA_SR + rs!gst_sr_harga
    End If
    If Not IsNull(rs!gst_sr_cukai) Then
        Frm75.MSFlexGrid3.TextMatrix(x, 7) = Format(rs!gst_sr_cukai, "0.00") 'Jumlah Cukai SR(RM)
        If IsNumeric(rs!gst_sr_cukai) Then Frm75_JUMLAH_CUKAI_SR = Frm75_JUMLAH_CUKAI_SR + rs!gst_sr_cukai
    End If
    If Not IsNull(rs!gst_zr_harga) Then
        Frm75.MSFlexGrid3.TextMatrix(x, 8) = Format(rs!gst_zr_harga, "0.00") 'Jumlah Harga ZR(RM)
        If IsNumeric(rs!gst_zr_harga) Then Frm75_JUMLAH_HARGA_ZR = Frm75_JUMLAH_HARGA_ZR + rs!gst_zr_harga
    End If
    If Not IsNull(rs!gst_zr_cukai) Then
        Frm75.MSFlexGrid3.TextMatrix(x, 9) = Format(rs!gst_zr_cukai, "0.00") 'Jumlah Cukai SR(RM)
        If IsNumeric(rs!gst_zr_cukai) Then Frm75_JUMLAH_CUKAI_ZR = Frm75_JUMLAH_CUKAI_ZR + rs!gst_zr_cukai
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'### Kutipan GST Dari Ansuran ### - End

'### Kutipan GST Dari Tempahan ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 41_akaun_tempahan where tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm75.MSFlexGrid3.Rows = x + 1
    Frm75.MSFlexGrid3.TextMatrix(x, 0) = x
    Frm75.MSFlexGrid3.TextMatrix(x, 1) = x 'No.
    If Not IsNull(rs!ID) Then Frm75.MSFlexGrid3.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm75.MSFlexGrid3.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!no_resit_tempahan) Then Frm75.MSFlexGrid3.TextMatrix(x, 4) = rs!no_resit_tempahan 'No. Resit
    Frm75.MSFlexGrid3.TextMatrix(x, 5) = "Bayaran Tempahan Emas" 'Detail
    If Not IsNull(rs!gst_sr_harga) Then
        Frm75.MSFlexGrid3.TextMatrix(x, 6) = Format(rs!gst_sr_harga, "0.00") 'Jumlah Harga SR(RM)
        If IsNumeric(rs!gst_sr_harga) Then Frm75_JUMLAH_HARGA_SR = Frm75_JUMLAH_HARGA_SR + rs!gst_sr_harga
    End If
    If Not IsNull(rs!gst_sr_cukai) Then
        Frm75.MSFlexGrid3.TextMatrix(x, 7) = Format(rs!gst_sr_cukai, "0.00") 'Jumlah Cukai SR(RM)
        If IsNumeric(rs!gst_sr_cukai) Then Frm75_JUMLAH_CUKAI_SR = Frm75_JUMLAH_CUKAI_SR + rs!gst_sr_cukai
    End If
    If Not IsNull(rs!gst_zr_harga) Then
        Frm75.MSFlexGrid3.TextMatrix(x, 8) = Format(rs!gst_zr_harga, "0.00") 'Jumlah Harga ZR(RM)
        If IsNumeric(rs!gst_zr_harga) Then Frm75_JUMLAH_HARGA_ZR = Frm75_JUMLAH_HARGA_ZR + rs!gst_zr_harga
    End If
    If Not IsNull(rs!gst_zr_cukai) Then
        Frm75.MSFlexGrid3.TextMatrix(x, 9) = Format(rs!gst_zr_cukai, "0.00") 'Jumlah Cukai SR(RM)
        If IsNumeric(rs!gst_zr_cukai) Then Frm75_JUMLAH_CUKAI_ZR = Frm75_JUMLAH_CUKAI_ZR + rs!gst_zr_cukai
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
'### Kutipan GST Dari Tempahan ### - End

Frm75.L14_Text = Format(Frm75_JUMLAH_HARGA_SR, "0.00") 'jumlah SR
Frm75.L15_Text = Format(Frm75_JUMLAH_CUKAI_SR, "0.00") 'Cukai SR
Frm75.L22_Text = Format(Frm75_JUMLAH_HARGA_ZR, "0.00") 'Harga ZR
Frm75.L23_Text = Format(Frm75_JUMLAH_CUKAI_ZR, "0.00") 'Cukai ZR
End Sub
Sub frm75_kiraan_summary_gst()
'on error resume next
Dim LM_KUTIP As Double
Dim LM_BAYAR As Double
Dim LM_GST As Double

LM_KUTIP = 0
LM_BAYAR = 0
LM_GST = 0

If Frm75.L17_Text <> vbNullString And IsNumeric(Frm75.L17_Text) Then LM_KUTIP = Frm75.L17_Text
If Frm75.L18_Text <> vbNullString And IsNumeric(Frm75.L18_Text) Then LM_BAYAR = Frm75.L18_Text

LM_GST = LM_BAYAR - LM_KUTIP

If LM_GST > 0 Then
    Frm75.L19_Text = "Anda layak buat tuntutan dari kastam sebanyak RM " & Format(LM_GST, "#,##0.00") & "."
ElseIf LM_GST < 0 Then
    Frm75.L19_Text = "Anda perlu bayar kepada kastam sebanyak RM " & Format(-LM_GST, "#,##0.00") & "."
ElseIf LM_GST = 0 Then
    Frm75.L19_Text = vbNullString
End If
End Sub
