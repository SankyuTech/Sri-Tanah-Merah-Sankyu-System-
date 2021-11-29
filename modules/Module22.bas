Attribute VB_Name = "Module22"
Sub frm116_one_time_reset()
'on error resume next
frm116.TB2 = "1.00" 'Kadar tukaran mutu
frm116.TB6 = "0.00"
frm116.DTPicker1 = DateTime.Date$

frm116.L39_Text = 0
frm116.L40_Text = 0
frm116.L41_Text = 0
frm116.L42_Text = 0
frm116.TB8 = "1.00" 'Kadar tukaran emas (mutu) - Urusan keseluruhan
frm116.L71_Text = vbNullString

frm116.Pic1.Visible = False
frm116.Pic1.Left = 10080
frm116.Pic1.Top = 240

frm116.CMD8.Visible = True
frm116.CMD9.Visible = True
frm116.CMD10.Visible = False
frm116.CMD11.Visible = False

frm116.CBB1.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where Metal_Purity<>'" & Null & "' order by Metal_Purity ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!Metal_Purity) Then frm116.CBB1.AddItem rs!Metal_Purity
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'###Senarai Nama Pekerja###
frm116.CBB4.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where status='" & "Aktif" & "' AND InvestorSmall = 0 AND InvestorBig = 0", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    If rs!user_level = 1 Or rs!user_level = 2 Or rs!user_level = 3 Or rs!user_level = G_LEVEL_USER Then frm116.CBB4.AddItem rs!Samaran & "  |  " & rs!NoPekerja
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

frm116.CBB2.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where status='" & 1 & "' order by supplier ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!supplier) Then frm116.CBB2.AddItem rs!supplier
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

frm116.L22_Text = G_RATE_GST 'Jumlah Kadar GST
If G_GST_JUAL = 1 Then
    frm116.CB3 = 1
    frm116.CB2 = 0
Else
    frm116.CB2 = 1
    frm116.CB3 = 0
End If
If G_GST_JUALAN_INC = 1 Then
    frm116.CB4 = 1
Else
    frm116.CB4 = 0
End If
frm116.TB6 = Format(G_HARGA_999, "0.00")

GoTo skip_daaa:

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Default1 = "Default" Then
        GLOBAL_DISABLE = 1

        If Not IsNull(rs!gst_value) Then frm116.L22_Text = rs!gst_value 'Jumlah Kadar GST
        If Not IsNull(rs!gst_arinashi) Then 'Tetapan GST , ZR atau SR
            If rs!gst_arinashi = 1 Then 'SR
                frm116.CB3 = 1
                frm116.CB2 = 0
                If Not IsNull(rs!gst_jualan_included) Then
                    If rs!gst_jualan_included = 1 Then
                    
                        frm116.CB4 = 1
                        
                    Else
                        
                        frm116.CB4 = 0
                        
                    End If
                End If
            Else 'ZR
                frm116.CB2 = 1
                frm116.CB3 = 0
                frm116.CB4 = 0
            End If
        End If

        If Not IsNull(rs!no_grn) Then frm116.L23_Text = rs!no_grn 'No. GRN

        If Not IsNull(rs!harga_999) Then
            frm116.TB6 = Format(rs!harga_beli_999, "0.00")
            frm116.TB6 = Format(rs!harga_beli_999, "0.00")
        Else
            frm116.TB6 = "0.00"
            frm116.TB6 = "0.00"
        End If

        GLOBAL_DISABLE = 0
    End If
End If

rs.Close
Set rs = Nothing
skip_daaa:

'###Padam Table Belian Temp### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "TRUNCATE TABLE " & G_GRN_TEMP & ""

Set rs = cn.Execute(strsql)
Set rs = Nothing
'###Padam Table Belian Temp### - End

Call Frm116_jurujual
End Sub
Sub frm116_reset_1()
'on error resume next
'Reset maklumat penerimaan barang
frm116.L2_Text = vbNullString 'Memory ID
frm116.TB1 = vbNullString 'Berat asal
frm116.L1_Text = "0.00" 'Berat 999.9
frm116.TB3 = "0.00" 'Upah
frm116.TB4 = "0.00" 'Jumlah GST
frm116.TB5 = "0.00" 'Upah Dengan GST

frm116.CMD1.Visible = True
frm116.CMD2.Visible = False
frm116.CMD3.Visible = False
End Sub
Sub Frm116_reset_3()
'on error resume next
'### Digunakan untuk reset paparan / komponen semua komponen transaksi
frm116.L9_Text = "0.00" 'Berat jualan 999.9
frm116.L12_Text = "0.00" 'Harga emas

frm116.L15_Text = "0.00" 'Maklumat GST : Jumlah harga tanpa GST
frm116.L16_Text = "0.00" 'Maklumat GST : Jumlah harga dengan GST
frm116.L17_Text = "0.00" 'Maklumat GST : Jumlah harga ZR
frm116.L18_Text = "0.00" 'Maklumat GST : Jumlah harga SR
frm116.L19_Text = "0.00" 'Maklumat GST : Jumlah GST ZR
frm116.L20_Text = "0.00" 'Maklumat GST : Jumlah GST SR
frm116.L24_Text = 0 'No. id jualan
frm116.L25_Text = 0 'No. id trade in

frm116.L43_Text = 0
frm116.L48_Text = "0.00" 'Berat Asal
frm116.TB9 = vbNullString 'No. rujukan dari supplier

frm116.L51_Text = "0.00" 'Upah tanpa GST
frm116.L52_Text = "0.00" 'Jumlah GST
frm116.L53_Text = "0.00" 'Upah dengan GST
'Frm116.TB2 = "0.00" 'Harga emas semasa 999.9 (Overall)
End Sub
Sub Frm116_jurujual()
'On Error Resume Next
If MDI_frm1.L3_Text <> vbNullString Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        frm116.CBB4 = rs!Samaran & "  |  " & rs!NoPekerja
        
    Else
        
        frm116.CBB4.AddItem "" & "  |  " & rs!Samaran
        frm116.CBB4 = "" & "  |  " & rs!Samaran
        
    End If
    
    rs.Close
    Set rs = Nothing

    If G_LOCK_JURUJUAL = "YES" Then
    
        frm116.CBB4.Enabled = False
        frm116.CBB4.BackColor = &H8000000A

    Else
    
        frm116.CBB4.Enabled = True
        frm116.CBB4.BackColor = &HFFFFFF

    End If
End If
End Sub
Sub Frm116_calc1()
'On Error Resume Next
Dim Frm116_LM_BERAT As Double
Dim frm116_LM_KADAR_TUKARAN As Double

Frm116_LM_BERAT = 0 'Berat jualan (g)
frm116_LM_KADAR_TUKARAN = 0 'Kadar tukaran kepada purity 999.9

If ((frm116.TB1 <> vbNullString And IsNumeric(frm116.TB1)) And (frm116.TB2 <> vbNullString And IsNumeric(frm116.TB2))) Then

    Frm116_LM_BERAT = frm116.TB1 'Berat jualan (g)
    frm116_LM_KADAR_TUKARAN = frm116.TB2 'Kadar tukaran kepada purity 999.9
    
    frm116.L1_Text = Format(Frm116_LM_BERAT * frm116_LM_KADAR_TUKARAN, "#,##0.00") 'Berat 999.9
    
Else

    frm116.L1_Text = "0.00" 'Berat 999.9
    
End If
End Sub
Sub frm116_calc2()
'On Error Resume Next
Dim frm116_LM_KADAR_GST As Double
Dim frm116_LM_UPAH As Double

frm116_LM_KADAR_GST = 0
frm116_LM_UPAH = 0

If IsNumeric(frm116.L22_Text) Then frm116_LM_KADAR_GST = frm116.L22_Text 'Kadar gst (%)
If IsNumeric(frm116.TB3) Then frm116_LM_UPAH = frm116.TB3 'Upah (RM)

If frm116.L22_Text <> vbNullString And IsNumeric(frm116.L22_Text) Then

    If frm116.TB3 <> vbNullString And IsNumeric(frm116.TB3) Then
    
        If frm116.CB2 = 1 Then 'Upah : GST ZR
        
            frm116.L30_Text = Format(frm116.TB3, "#,##0.00") 'Harga upah tanpa GST
            frm116.TB4 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
            
        End If
        
        If frm116.CB3 = 1 Then
        
            frm116.L30_Text = Format(frm116_LM_UPAH, "#,##0.00") 'Harga upah tanpa GST
            frm116.TB4 = Format(frm116_LM_UPAH * (frm116_LM_KADAR_GST / 100), "#,##0.00") 'Jumlah GST (Bagi jualan setiap item)
            
        End If
        
        
        If frm116.CB4 = 1 Then
    
            frm116.L30_Text = Format(frm116_LM_UPAH / (1 + (frm116_LM_KADAR_GST / 100)), "#,##0.00") 'Harga upah tanpa GST
            frm116.TB4 = Format(frm116_LM_UPAH - (frm116_LM_UPAH / (1 + (frm116_LM_KADAR_GST / 100))), "#,##0.00") 'Jumlah GST (Bagi jualan setiap item)
                
        End If

    Else
    
        frm116.L30_Text = Format(0, "#,##0.00") 'Harga upah tanpa GST
        frm116.TB4 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
        
    End If

Else

    If IsNumeric(frm116.TB3) Then
    
        frm116.L30_Text = Format(frm116.TB3, "#,##0.00") 'Harga upah tanpa GST
        frm116.TB4 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
        
    Else
        
        frm116.L30_Text = Format(0, "#,##0.00") 'Harga upah tanpa GST
        frm116.TB4 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
        
    End If
    
End If
End Sub
Sub frm116_calc3()
'On Error Resume Next
Dim Frm116_LM_UPAH_TANPA_GST As Double
Dim frm116_LM_GST As Double

Frm116_LM_UPAH_TANPA_GST = 0 'Jumlah upah tanpa GST
frm116_LM_GST = 0 'Jumlah GST

If ((frm116.TB4 <> vbNullString And IsNumeric(frm116.TB4)) And (frm116.L30_Text <> vbNullString And IsNumeric(frm116.L30_Text))) Then

    frm116_LM_GST = frm116.TB4 'Jumlah GST (Bagi jualan setiap item)
    Frm116_LM_UPAH_TANPA_GST = frm116.L30_Text 'Harga upah tanpa GST
    
    frm116.TB5 = Format(frm116_LM_GST + Frm116_LM_UPAH_TANPA_GST, "#,##0.00") 'Jumlah Upah + GST (Bagi jualan setiap item)
    
Else

    frm116.TB5 = "0.00" 'Jumlah Upah + GST (Bagi jualan setiap item)
    
End If
End Sub
Sub Frm116_calc5()
'On Error Resume Next
Dim Frm116_LM_BEZA_BERAT As Double
Dim Frm116_LM_HARGA_SEMASA As Double

Frm116_LM_BEZA_BERAT = 0 'Beza berat (g)
Frm116_LM_HARGA_SEMASA = 0 'Harga semasa (RM/g)

If ((frm116.L9_Text <> vbNullString And IsNumeric(frm116.L9_Text)) And (frm116.TB6 <> vbNullString And IsNumeric(frm116.TB6))) Then
    Frm116_LM_BEZA_BERAT = frm116.L9_Text 'Berat jualan (g)
    Frm116_LM_HARGA_SEMASA = frm116.TB6 'Kadar belian (g)
    
    frm116.L12_Text = Format(Frm116_LM_BEZA_BERAT * Frm116_LM_HARGA_SEMASA, "#,##0.00") 'Harga jualan
Else
    frm116.L12_Text = "0.00" 'Harga jualan
End If
End Sub
Sub Frm116_Senarai_Belian_Header()
'on error resume next
frm116.MSFlexGrid1.Clear
frm116.MSFlexGrid1.RowHeight(0) = 700
frm116.MSFlexGrid1.FormatString = "No.|<No.|<ID|<Purity|<Berat Asal (g)|<Mutu|<Berat 999.9 (g)|<Upah (RM)|<Jenis GST|<Jumlah GST (RM)|<Upah + GST (RM)"

frm116.MSFlexGrid1.Rows = 1
frm116.MSFlexGrid1.ColWidth(0) = 600 'No.
frm116.MSFlexGrid1.ColAlignment(0) = 4

frm116.MSFlexGrid1.ColWidth(1) = 0 'No.
frm116.MSFlexGrid1.ColWidth(2) = 0 'No. ID
frm116.MSFlexGrid1.ColWidth(3) = 1800 'Purity
frm116.MSFlexGrid1.ColWidth(4) = 1100 'Berat Asal (g)
frm116.MSFlexGrid1.ColAlignment(4) = 7

frm116.MSFlexGrid1.ColWidth(5) = 1000 'Mutu
frm116.MSFlexGrid1.ColAlignment(5) = 7

frm116.MSFlexGrid1.ColWidth(6) = 1000 'Berat 999.9 (g)
frm116.MSFlexGrid1.ColAlignment(6) = 7

frm116.MSFlexGrid1.ColWidth(7) = 1100 'Upah (RM)
frm116.MSFlexGrid1.ColAlignment(7) = 7

frm116.MSFlexGrid1.ColWidth(8) = 800 'Jenis GST
frm116.MSFlexGrid1.ColAlignment(8) = 4

frm116.MSFlexGrid1.ColWidth(9) = 1100 'Jumlah GST (RM)
frm116.MSFlexGrid1.ColAlignment(9) = 7

frm116.MSFlexGrid1.ColWidth(10) = 1100 'Upah + GST (RM)
frm116.MSFlexGrid1.ColAlignment(10) = 7
End Sub
Sub Frm116_Senarai_Belian()
'on error resume next
Dim Frm116_LM_TOTAL_PAGE As Double
Dim Frm116_LM_FIELD As String
Dim Frm116_LM_UPAH_TANPA_GST As Double 'Harga Jualan Tanpa Cukai GST
Dim Frm116_LM_UPAH_DENGAN_GST As Double 'Harga Jualan Dengan Cukai GST
Dim Frm116_LM_GST_SR As Double 'Kutipan GST : SR
Dim Frm116_LM_GST_ZR As Double 'Kutipan GST : ZR
Dim Frm116_LM_JUMLAH_UPAH_SR As Double 'Total Harga Yang Dikenakan GST SR
Dim Frm116_LM_JUMLAH_UPAH_ZR As Double 'Total Harga Yang Dikenakan GST ZR
Dim Frm116_LM_BERAT As Double 'Berat Jualan
Dim Frm116_LM_BERAT_ASAL As Double 'Berat Asal (Sebelum tukar kepada purity 999.9)

Frm116_PAGE_SIZE = 26
Frm116_LM_TOTAL_PAGE = 0
x = 0
Frm116_LM_UPAH_TANPA_GST = 0
Frm116_LM_UPAH_DENGAN_GST = 0
Frm116_LM_GST_SR = 0
Frm116_LM_GST_ZR = 0
Frm116_LM_JUMLAH_UPAH_SR = 0
Frm116_LM_JUMLAH_UPAH_ZR = 0
Frm116_LM_BERAT = 0
Frm116_LM_BERAT_ASAL = 0 'Berat Asal (Sebelum tukar kepada purity 999.9)

re_gen_report:

frm116.L43_Text = x 'Jumlah bilangan barang jualan
frm116.L48_Text = Format(0, "#,##0.00") 'Jumlah berat jualan
frm116.L35_Text = Format(0, "#,##0.00") 'Maklumat GST : Jumlah harga ZR
frm116.L37_Text = Format(0, "#,##0.00") 'Maklumat GST : Jumlah harga SR
frm116.L36_Text = Format(0, "#,##0.00") 'Maklumat GST : Jumlah GST ZR
frm116.L38_Text = Format(0, "#,##0.00")  'Maklumat GST : Jumlah GST SR
frm116.L9_Text = Format(0, "#,##0.00") 'Berat jualan 999.9
frm116.L51_Text = Format(0, "#,##0.00") 'Jumlah upah tanpa GST (keseluruhan)
frm116.L52_Text = Format(0, "#,##0.00") 'Jumlah GST (Keseluruhan)
frm116.L53_Text = Format(0, "#,##0.00") 'Jumlah Upah + GST (Keseluruhan)

LM_START_ROW = frm116.L69_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm116_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If frm116.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm116_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        frm116.L67_Text = 1
    End If
End If

Frm116_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from " & G_GRN_TEMP & " where Status='" & 1 & "' Or Status='" & 2 & "' Or Status='" & 3 & "' Or Status='" & 4 & "' order by purity ASC LIMIT " & LM_START_ROW & "," & Frm116_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1
    If Frm116_LM_PAGE_FOUND = 0 Then
        If frm116.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                frm116.L67_Text = frm116.L67_Text + 1 'Paparan Page ke-xxx
                Frm116_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(frm116.L67_Text) Then
                    If frm116.L67_Text <> 1 Then
                        frm116.L67_Text = frm116.L67_Text - 1 'Paparan Page ke-xxx
                        Frm116_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    
    Y = ((frm116.L67_Text - 1) * Frm116_PAGE_SIZE) + x
    frm116.MSFlexGrid1.Rows = x + 1
    frm116.MSFlexGrid1.TextMatrix(x, 0) = x 'No.
    frm116.MSFlexGrid1.TextMatrix(x, 1) = Y 'No.
    frm116.MSFlexGrid1.ColAlignment(1) = 4
    frm116.MSFlexGrid1.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!purity) Then frm116.MSFlexGrid1.TextMatrix(x, 3) = rs!purity 'Purity
    If Not IsNull(rs!Berat_Asal) Then frm116.MSFlexGrid1.TextMatrix(x, 4) = Format(rs!Berat_Asal, "#,##0.00") 'Berat Asal (g)
    If Not IsNull(rs!kadar_tukaran) Then frm116.MSFlexGrid1.TextMatrix(x, 5) = rs!kadar_tukaran 'Mutu
    If Not IsNull(rs!berat_tukaran_grn) Then frm116.MSFlexGrid1.TextMatrix(x, 6) = Format(rs!berat_tukaran_grn, "#,##0.00") 'Berat 999.9 (g)
    If Not IsNull(rs!UPAH) Then frm116.MSFlexGrid1.TextMatrix(x, 7) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
    If Not IsNull(rs!gst_ari_nashi) Then frm116.MSFlexGrid1.TextMatrix(x, 8) = rs!gst_ari_nashi 'Jenis GST
    If Not IsNull(rs!jumlah_gst) Then frm116.MSFlexGrid1.TextMatrix(x, 9) = Format(rs!jumlah_gst, "#,##0.00") 'Jumlah GST (RM)
    If Not IsNull(rs!harga_dengan_gst_grn) Then frm116.MSFlexGrid1.TextMatrix(x, 10) = Format(rs!harga_dengan_gst_grn, "#,##0.00") 'Upah + GST (RM)
    
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from " & G_GRN_TEMP & " where Status='" & 1 & "' Or Status='" & 2 & "' Or Status='" & 3 & "' Or Status='" & 4 & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    LM_BILANGAN_AHLI = rs(0)
    Frm116_LM_TOTAL_PAGE = Format(rs(0) / Frm116_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm116_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm116_LM_PAGE = Split(Frm116_LM_TOTAL_PAGE, ".")(0)
        Frm116_LM_PAGE_LEBIHAN = Split(Frm116_LM_TOTAL_PAGE, ".")(1)
        
        If Frm116_LM_PAGE_LEBIHAN <> "00" Then
            frm116.L68_Text = Frm116_LM_PAGE + 1
        Else
            frm116.L68_Text = Frm116_LM_PAGE
        End If
        
    Else
    
        frm116.L68_Text = Frm116_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        frm116.L68_Text = 0
    End If
Else
    frm116.L68_Text = 0
End If

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) , SUM(berat_tukaran_grn) , SUM(berat_tukaran_grn) , SUM(harga_tanpa_gst_grn) , SUM(jumlah_gst) , SUM(harga_dengan_gst_grn) from " & G_GRN_TEMP & " where Status='" & 1 & "' Or Status='" & 2 & "' Or Status='" & 3 & "' Or Status='" & 4 & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then frm116.L43_Text = rs(0) 'Jumlah bilangan barang jualan
If Not IsNull(rs(1)) Then frm116.L48_Text = Format(rs(1), "#,##0.00") 'Jumlah berat jualan
'If Not IsNull(rs(2)) Then frm116.L9_Text = Format(rs(2), "#,##0.00") 'Berat jualan 999.9
If Not IsNull(rs(3)) Then frm116.L51_Text = Format(rs(3), "#,##0.00") 'Jumlah upah tanpa GST (keseluruhan)
If Not IsNull(rs(4)) Then frm116.L52_Text = Format(rs(4), "#,##0.00") 'Jumlah GST (Keseluruhan)
If Not IsNull(rs(5)) Then frm116.L53_Text = Format(rs(5), "#,##0.00") 'Jumlah Upah + GST (Keseluruhan)

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_gst) , SUM(harga_tanpa_gst_grn) from " & G_GRN_TEMP & " where (Status='" & 1 & "' Or Status='" & 2 & "' Or Status='" & 3 & "' Or Status='" & 4 & "') AND gst_ari_nashi='" & "ZR" & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then frm116.L36_Text = Format(rs(0), "#,##0.00") 'Maklumat GST : Jumlah GST ZR
If Not IsNull(rs(1)) Then frm116.L35_Text = Format(rs(1), "#,##0.00") 'Maklumat GST : Jumlah harga ZR

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_gst) , SUM(harga_tanpa_gst_grn) from " & G_GRN_TEMP & " where (Status='" & 1 & "' Or Status='" & 2 & "' Or Status='" & 3 & "' Or Status='" & 4 & "') AND gst_ari_nashi='" & "SR" & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then frm116.L38_Text = Format(rs(0), "#,##0.00") 'Maklumat GST : Jumlah GST SR
If Not IsNull(rs(1)) Then frm116.L37_Text = Format(rs(1), "#,##0.00") 'Maklumat GST : Jumlah harga SR

rs.Close
Set rs = Nothing

If x <> 0 Then
    frm116.L69_Text = LM_START_ROW
End If

If frm116.L67_Text <> vbNullString And IsNumeric(frm116.L67_Text) Then
    If frm116.L68_Text <> vbNullString And IsNumeric(frm116.L68_Text) Then
        Frm116_LM_CURR_PAGE = frm116.L67_Text
        Frm116_LM_TOTAL_PAGE = frm116.L68_Text
        
        If Frm116_LM_CURR_PAGE > Frm116_LM_TOTAL_PAGE Then
            
            frm116.L67_Text = frm116.L67_Text - 1
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            GoTo re_gen_report:
            
        End If
    End If
End If

End Sub
Sub Frm116_calc10()
'On Error Resume Next
Dim Frm116_LM_HARGA_ZR_UPAH As Double
Dim Frm116_LM_HARGA_SR_UPAH As Double
Dim Frm116_LM_HARGA_ZR_EMAS As Double
Dim Frm116_LM_HARGA_SR_EMAS As Double
Dim Frm116_LM_GST_ZR_UPAH As Double
Dim Frm116_LM_GST_SR_UPAH As Double
Dim Frm116_LM_GST_ZR_EMAS As Double
Dim Frm116_LM_GST_SR_EMAS As Double

Frm116_LM_HARGA_ZR_UPAH = 0
Frm116_LM_HARGA_SR_UPAH = 0
Frm116_LM_HARGA_ZR_EMAS = 0
Frm116_LM_HARGA_SR_EMAS = 0
Frm116_LM_GST_ZR_UPAH = 0
Frm116_LM_GST_SR_UPAH = 0
Frm116_LM_GST_ZR_EMAS = 0
Frm116_LM_GST_SR_EMAS = 0

If ((frm116.L35_Text <> vbNullString And IsNumeric(frm116.L35_Text)) And (frm116.L39_Text <> vbNullString And IsNumeric(frm116.L39_Text))) Then

    Frm116_LM_HARGA_ZR_UPAH = frm116.L35_Text 'Harga ZR (Upah)
    'Frm116_LM_HARGA_ZR_EMAS = Frm116.L39_Text 'Harga ZR (Emas)
    
    frm116.L17_Text = Format(Frm116_LM_HARGA_ZR_UPAH + Frm116_LM_HARGA_ZR_EMAS, "#,##0.00") 'Jumlah Harga ZR
    
Else

    frm116.L17_Text = "0.00" 'Jumlah Harga ZR
    
End If

If ((frm116.L37_Text <> vbNullString And IsNumeric(frm116.L37_Text)) And (frm116.L41_Text <> vbNullString And IsNumeric(frm116.L41_Text))) Then

    Frm116_LM_HARGA_SR_UPAH = frm116.L37_Text 'Harga SR (Upah)
    'Frm116_LM_HARGA_SR_EMAS = Frm116.L41_Text 'Harga SR (Emas)
    
    frm116.L18_Text = Format(Frm116_LM_HARGA_SR_UPAH + Frm116_LM_HARGA_SR_EMAS, "#,##0.00") 'Jumlah Harga SR
    
Else

    frm116.L18_Text = "0.00" 'Jumlah Harga SR
    
End If

If ((frm116.L36_Text <> vbNullString And IsNumeric(frm116.L36_Text)) And (frm116.L40_Text <> vbNullString And IsNumeric(frm116.L40_Text))) Then

    Frm116_LM_GST_SR_UPAH = frm116.L36_Text 'GST ZR (Upah)
    'Frm116_LM_GST_SR_EMAS = Frm116.L40_Text 'GST ZR (Emas)
    
    frm116.L20_Text = Format(Frm116_LM_GST_SR_UPAH + Frm116_LM_GST_SR_EMAS, "#,##0.00") 'Jumlah GST ZR
    
Else

    frm116.L20_Text = "0.00" 'Jumlah GST ZR
    
End If

If ((frm116.L38_Text <> vbNullString And IsNumeric(frm116.L38_Text)) And (frm116.L42_Text <> vbNullString And IsNumeric(frm116.L42_Text))) Then

    Frm116_LM_GST_ZR_UPAH = frm116.L38_Text 'GST SR (Upah)
    'Frm116_LM_GST_ZR_EMAS = Frm116.L42_Text 'GST SR (Emas)
    
    frm116.L20_Text = Format(Frm116_LM_GST_ZR_UPAH + Frm116_LM_GST_ZR_EMAS, "#,##0.00") 'Jumlah GST SR
    
Else

    frm116.L20_Text = "0.00" 'Jumlah GST SR
    
End If

frm116.L15_Text = Format(Frm116_LM_HARGA_ZR_UPAH + Frm116_LM_HARGA_ZR_EMAS + Frm116_LM_HARGA_SR_UPAH + Frm116_LM_HARGA_SR_EMAS, "#,##0.00") 'Jumlah harga tanpa GST
frm116.L16_Text = Format(Frm116_LM_HARGA_ZR_UPAH + Frm116_LM_HARGA_ZR_EMAS + Frm116_LM_HARGA_SR_UPAH + Frm116_LM_HARGA_SR_EMAS + Frm116_LM_GST_SR_UPAH + Frm116_LM_GST_SR_EMAS + Frm116_LM_GST_ZR_UPAH + Frm116_LM_GST_ZR_EMAS, "#,##0.00") 'Jumlah harga dengan GST
End Sub
Sub Frm116_cetak_grn()
'on error resume next
Frm115_LM_CUST = vbNullString

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

Report78.Sections("Section4").Controls("L205").Caption = "Goods Received Note"

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
    If Not IsNull(rs!berat_tukaran_grn) Then Report78.Sections("Section5").Controls("L20").Caption = Format(rs!berat_tukaran_grn, "#,##0.00 g") 'Berat 999.9 (g)
    If Not IsNull(rs!jumlah_gst) Then Report78.Sections("Section5").Controls("L13").Caption = Format(rs!jumlah_gst, "#,##0.00") 'Jumlah GST
    If Not IsNull(rs!harga_dengan_gst_grn) Then Report78.Sections("Section5").Controls("L14").Caption = Format(rs!harga_dengan_gst_grn, "#,##0.00") 'Jumlah keseluruhan (Upah + GST)
    If Not IsNull(rs!gst_sr_harga) Then Report78.Sections("Section5").Controls("L8").Caption = Format(rs!gst_sr_harga, "#,##0.00") 'Jumlah harga SR
    If Not IsNull(rs!gst_zr_harga) Then Report78.Sections("Section5").Controls("L9").Caption = Format(rs!gst_zr_harga, "#,##0.00") 'Jumlah harga ZR
    If Not IsNull(rs!gst_sr_cukai) Then Report78.Sections("Section5").Controls("L10").Caption = Format(rs!gst_sr_cukai, "#,##0.00") 'Jumlah GST SR
    If Not IsNull(rs!gst_zr_cukai) Then Report78.Sections("Section5").Controls("L11").Caption = Format(rs!gst_zr_cukai, "#,##0.00") 'Jumlah GST ZR
    If Not IsNull(rs!supplier_agen) Then Frm115_LM_CUST = rs!supplier_agen
    If Not IsNull(rs!no_rujukan_supplier) Then Report78.Sections("Section4").Controls("L21").Caption = "No. Rujukan Supplier         : " & rs!no_rujukan_supplier 'No. Rujukan Dari Supplier

End If

rs.Close
Set rs = Nothing

If Frm115_LM_CUST <> vbNullString Then
 
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from setting_database where Supplier='" & Frm115_LM_CUST & "'", cn, adOpenKeyset, adLockOptimistic
    
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
rs.Open "select * from 79_grn where no_rujukan='" & G_No_RESIT_JUALAN & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    Set Report78.DataSource = rs
    Report78.Show
    
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing

End Sub
Sub frm116_calc11()
'On Error Resume Next
Dim Frm116_LM_BERAT_ASAL As Double
Dim Frm116_LM_MUTU As Double

Frm116_LM_BERAT_ASAL = 0 'Beza berat (g)
Frm116_LM_MUTU = 0 'Harga semasa (RM/g)

If ((frm116.L48_Text <> vbNullString And IsNumeric(frm116.L48_Text)) And (frm116.TB8 <> vbNullString And IsNumeric(frm116.TB8))) Then
    Frm116_LM_BERAT_ASAL = frm116.L48_Text 'Berat jualan (g)
    Frm116_LM_MUTU = frm116.TB8 'Kadar belian (g)
    
    frm116.L9_Text = Format(Frm116_LM_BERAT_ASAL * Frm116_LM_MUTU, "#,##0.00") 'Harga jualan
Else
    frm116.L9_Text = "0.00" 'Harga jualan
End If
End Sub

