Attribute VB_Name = "Module29"
Sub Frm123_one_time_reset()
'on error resume next
frm123.TB2 = "1.00" 'Kadar tukaran mutu
frm123.TB6 = "0.00"
frm123.DTPicker1 = DateTime.Date$

frm123.L39_Text = 0
frm123.L40_Text = 0
frm123.L41_Text = 0
frm123.L42_Text = 0
frm123.TB8 = "1.00" 'Kadar tukaran emas (mutu) - Urusan keseluruhan
frm123.L71_Text = vbNullString

frm123.Pic1.Visible = False
frm123.Pic1.Left = 10080
frm123.Pic1.Top = 240

frm123.CMD8.Visible = True
frm123.CMD9.Visible = True
frm123.CMD10.Visible = False
frm123.CMD11.Visible = False

frm123.CBB1.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where Metal_Purity<>'" & Null & "' order by Metal_Purity ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!Metal_Purity) Then frm123.CBB1.AddItem rs!Metal_Purity
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'###Senarai Nama Pekerja###
frm123.CBB4.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where status='" & "Aktif" & "' AND InvestorSmall = 0 AND InvestorBig = 0", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    If rs!user_level = 1 Or rs!user_level = 2 Or rs!user_level = 3 Or rs!user_level = G_LEVEL_USER Then frm123.CBB4.AddItem rs!Samaran & "  |  " & rs!NoPekerja
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

frm123.CBB2.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where status='" & 1 & "' order by supplier ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!supplier) Then frm123.CBB2.AddItem rs!supplier
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

frm123.L22_Text = G_RATE_GST 'Jumlah Kadar GST

If G_GST_JUALAN_INC = 1 Then
    frm123.CB4 = 1
Else
    frm123.CB4 = 0
End If
If G_GST_JUAL = 1 Then
    frm123.CB3 = 1
    frm123.CB2 = 0
Else
    frm123.CB2 = 1
    frm123.CB3 = 0
    frm123.CB4 = 0
End If
frm123.TB6 = Format(G_HARGA_999, "0.00")

GoTo skip_oi:

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Default1 = "Default" Then
        GLOBAL_DISABLE = 1

        If Not IsNull(rs!gst_value) Then frm123.L22_Text = rs!gst_value 'Jumlah Kadar GST
        If Not IsNull(rs!gst_arinashi) Then 'Tetapan GST , ZR atau SR
            If rs!gst_arinashi = 1 Then 'SR
                frm123.CB3 = 1
                frm123.CB2 = 0
                If Not IsNull(rs!gst_jualan_included) Then
                    If rs!gst_jualan_included = 1 Then
                    
                        frm123.CB4 = 1
                        
                    Else
                        
                        frm123.CB4 = 0
                        
                    End If
                End If
            Else 'ZR
                frm123.CB2 = 1
                frm123.CB3 = 0
                frm123.CB4 = 0
            End If
        End If

        If Not IsNull(rs!no_grn) Then frm123.L23_Text = rs!no_grn 'No. GRN

        If Not IsNull(rs!harga_999) Then
            frm123.TB6 = Format(rs!harga_999, "0.00")
            frm123.TB6 = Format(rs!harga_999, "0.00")
        Else
            frm123.TB6 = "0.00"
            frm123.TB6 = "0.00"
        End If

        GLOBAL_DISABLE = 0
    End If
End If

rs.Close
Set rs = Nothing

skip_oi:

'###Padam Table Belian Temp### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "TRUNCATE TABLE " & G_GRN_TEMP & ""

Set rs = cn.Execute(strsql)
Set rs = Nothing
'###Padam Table Belian Temp### - End

Call Frm123_jurujual
End Sub
Sub Frm123_reset_1()
'on error resume next
'Reset maklumat penerimaan barang
GLOBAL_DISABLE = 0
frm123.L2_Text = vbNullString 'Memory ID
frm123.TB1 = vbNullString 'Berat asal
frm123.L1_Text = "0.00" 'Berat 999.9
frm123.TB3 = "0.00" 'Upah
frm123.TB4 = "0.00" 'Jumlah GST
frm123.TB5 = "0.00" 'Upah Dengan GST
frm123.TB10 = "0.00" 'Baki berat

frm123.CMD1.Visible = True
frm123.CMD2.Visible = False
frm123.CMD3.Visible = False

frm123.CBB1.Enabled = True
frm123.CBB1.BackColor = &HFFFFFF
End Sub
Sub Frm123_reset_3()
'on error resume next
'### Digunakan untuk reset paparan / komponen semua komponen transaksi
frm123.L9_Text = "0.00" 'Berat jualan 999.9
frm123.L12_Text = "0.00" 'Harga emas

frm123.L15_Text = "0.00" 'Maklumat GST : Jumlah harga tanpa GST
frm123.L16_Text = "0.00" 'Maklumat GST : Jumlah harga dengan GST
frm123.L17_Text = "0.00" 'Maklumat GST : Jumlah harga ZR
frm123.L18_Text = "0.00" 'Maklumat GST : Jumlah harga SR
frm123.L19_Text = "0.00" 'Maklumat GST : Jumlah GST ZR
frm123.L20_Text = "0.00" 'Maklumat GST : Jumlah GST SR
frm123.L24_Text = 0 'No. id jualan
frm123.L25_Text = 0 'No. id trade in

frm123.L43_Text = 0
frm123.L48_Text = "0.00" 'Berat Asal
frm123.TB9 = vbNullString 'No. rujukan dari supplier

frm123.L51_Text = "0.00" 'Upah tanpa GST
frm123.L52_Text = "0.00" 'Jumlah GST
frm123.L53_Text = "0.00" 'Upah dengan GST
'Frm123.TB2 = "0.00" 'Harga emas semasa 999.9 (Overall)
End Sub
Sub Frm123_jurujual()
'On Error Resume Next
If MDI_frm1.L3_Text <> vbNullString Then

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where Samaran='" & MDI_frm1.L3_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
    
        frm123.CBB4 = rs!Samaran & "  |  " & rs!NoPekerja
        
    Else
        
        frm123.CBB4.AddItem "" & "  |  " & rs!Samaran
        frm123.CBB4 = "" & "  |  " & rs!Samaran
        
    End If
    
    rs.Close
    Set rs = Nothing
    
    If G_LOCK_JURUJUAL = "YES" Then
    
        frm123.CBB4.Enabled = False
        frm123.CBB4.BackColor = &H8000000A

    Else
    
        frm123.CBB4.Enabled = True
        frm123.CBB4.BackColor = &HFFFFFF

    End If

End If
End Sub
Sub Frm123_calc1()
'On Error Resume Next
Dim Frm123_LM_BERAT As Double
Dim Frm123_LM_KADAR_TUKARAN As Double

Frm123_LM_BERAT = 0 'Berat jualan (g)
Frm123_LM_KADAR_TUKARAN = 0 'Kadar tukaran kepada purity 999.9

If ((frm123.TB1 <> vbNullString And IsNumeric(frm123.TB1)) And (frm123.TB2 <> vbNullString And IsNumeric(frm123.TB2))) Then

    Frm123_LM_BERAT = frm123.TB1 'Berat jualan (g)
    Frm123_LM_KADAR_TUKARAN = frm123.TB2 'Kadar tukaran kepada purity 999.9
    
    frm123.L1_Text = Format(Frm123_LM_BERAT * Frm123_LM_KADAR_TUKARAN, "#,##0.00") 'Berat 999.9
    
Else

    frm123.L1_Text = "0.00" 'Berat 999.9
    
End If
End Sub
Sub Frm123_calc2()
'On Error Resume Next
Dim Frm123_LM_KADAR_GST As Double
Dim Frm123_LM_UPAH As Double

Frm123_LM_KADAR_GST = 0
Frm123_LM_UPAH = 0

If IsNumeric(frm123.L22_Text) Then Frm123_LM_KADAR_GST = frm123.L22_Text 'Kadar gst (%)
If IsNumeric(frm123.TB3) Then Frm123_LM_UPAH = frm123.TB3 'Upah (RM)

If frm123.L22_Text <> vbNullString And IsNumeric(frm123.L22_Text) Then

    If frm123.TB3 <> vbNullString And IsNumeric(frm123.TB3) Then
    
        If frm123.CB2 = 1 Then 'Upah : GST ZR
        
            frm123.L30_Text = Format(frm123.TB3, "#,##0.00") 'Harga upah tanpa GST
            frm123.TB4 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
            
        End If
        
        If frm123.CB3 = 1 Then
        
            frm123.L30_Text = Format(Frm123_LM_UPAH, "#,##0.00") 'Harga upah tanpa GST
            frm123.TB4 = Format(Frm123_LM_UPAH * (Frm123_LM_KADAR_GST / 100), "#,##0.00") 'Jumlah GST (Bagi jualan setiap item)
            
        End If
        
        
        If frm123.CB4 = 1 Then
    
            frm123.L30_Text = Format(Frm123_LM_UPAH / (1 + (Frm123_LM_KADAR_GST / 100)), "#,##0.00") 'Harga upah tanpa GST
            frm123.TB4 = Format(Frm123_LM_UPAH - (Frm123_LM_UPAH / (1 + (Frm123_LM_KADAR_GST / 100))), "#,##0.00") 'Jumlah GST (Bagi jualan setiap item)
                
        End If

    Else
    
        frm123.L30_Text = Format(0, "#,##0.00") 'Harga upah tanpa GST
        frm123.TB4 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
        
    End If

Else

    If IsNumeric(frm123.TB3) Then
    
        frm123.L30_Text = Format(frm123.TB3, "#,##0.00") 'Harga upah tanpa GST
        frm123.TB4 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
        
    Else
        
        frm123.L30_Text = Format(0, "#,##0.00") 'Harga upah tanpa GST
        frm123.TB4 = "0.00" 'Jumlah GST (Bagi jualan setiap item)
        
    End If
    
End If
End Sub
Sub Frm123_calc3()
'On Error Resume Next
Dim Frm123_LM_UPAH_TANPA_GST As Double
Dim Frm123_LM_GST As Double

Frm123_LM_UPAH_TANPA_GST = 0 'Jumlah upah tanpa GST
Frm123_LM_GST = 0 'Jumlah GST

If ((frm123.TB4 <> vbNullString And IsNumeric(frm123.TB4)) And (frm123.L30_Text <> vbNullString And IsNumeric(frm123.L30_Text))) Then

    Frm123_LM_GST = frm123.TB4 'Jumlah GST (Bagi jualan setiap item)
    Frm123_LM_UPAH_TANPA_GST = frm123.L30_Text 'Harga upah tanpa GST
    
    frm123.TB5 = Format(Frm123_LM_GST + Frm123_LM_UPAH_TANPA_GST, "#,##0.00") 'Jumlah Upah + GST (Bagi jualan setiap item)
    
Else

    frm123.TB5 = "0.00" 'Jumlah Upah + GST (Bagi jualan setiap item)
    
End If
End Sub
Sub Frm123_calc5()
'On Error Resume Next
Dim Frm123_LM_BEZA_BERAT As Double
Dim Frm123_LM_HARGA_SEMASA As Double

Frm123_LM_BEZA_BERAT = 0 'Beza berat (g)
Frm123_LM_HARGA_SEMASA = 0 'Harga semasa (RM/g)

If ((frm123.L9_Text <> vbNullString And IsNumeric(frm123.L9_Text)) And (frm123.TB6 <> vbNullString And IsNumeric(frm123.TB6))) Then
    Frm123_LM_BEZA_BERAT = frm123.L9_Text 'Berat jualan (g)
    Frm123_LM_HARGA_SEMASA = frm123.TB6 'Kadar belian (g)
    
    frm123.L12_Text = Format(Frm123_LM_BEZA_BERAT * Frm123_LM_HARGA_SEMASA, "#,##0.00") 'Harga jualan
Else
    frm123.L12_Text = "0.00" 'Harga jualan
End If
End Sub
Sub Frm123_Senarai_Belian_Header()
'on error resume next
frm123.MSFlexGrid1.Clear
frm123.MSFlexGrid1.RowHeight(0) = 700
frm123.MSFlexGrid1.FormatString = "No.|<No.|<ID|<Purity|<Berat Asal (g)|<Mutu|<Berat 999.9 (g)|<Upah (RM)|<Jenis GST|<Jumlah GST (RM)|<Upah + GST (RM)"

frm123.MSFlexGrid1.Rows = 1
frm123.MSFlexGrid1.ColWidth(0) = 600 'No.
frm123.MSFlexGrid1.ColAlignment(0) = 4

frm123.MSFlexGrid1.ColWidth(1) = 0 'No.
frm123.MSFlexGrid1.ColWidth(2) = 0 'No. ID
frm123.MSFlexGrid1.ColWidth(3) = 1800 'Purity
frm123.MSFlexGrid1.ColWidth(4) = 1100 'Berat Asal (g)
frm123.MSFlexGrid1.ColAlignment(4) = 7

frm123.MSFlexGrid1.ColWidth(5) = 1000 'Mutu
frm123.MSFlexGrid1.ColAlignment(5) = 7

frm123.MSFlexGrid1.ColWidth(6) = 1000 'Berat 999.9 (g)
frm123.MSFlexGrid1.ColAlignment(6) = 7

frm123.MSFlexGrid1.ColWidth(7) = 1100 'Upah (RM)
frm123.MSFlexGrid1.ColAlignment(7) = 7

frm123.MSFlexGrid1.ColWidth(8) = 800 'Jenis GST
frm123.MSFlexGrid1.ColAlignment(8) = 4

frm123.MSFlexGrid1.ColWidth(9) = 1100 'Jumlah GST (RM)
frm123.MSFlexGrid1.ColAlignment(9) = 7

frm123.MSFlexGrid1.ColWidth(10) = 1100 'Upah + GST (RM)
frm123.MSFlexGrid1.ColAlignment(10) = 7
End Sub
Sub Frm123_Senarai_Belian()
'on error resume next
Dim Frm123_LM_TOTAL_PAGE As Double
Dim Frm123_LM_FIELD As String
Dim Frm123_LM_UPAH_TANPA_GST As Double 'Harga Jualan Tanpa Cukai GST
Dim Frm123_LM_UPAH_DENGAN_GST As Double 'Harga Jualan Dengan Cukai GST
Dim Frm123_LM_GST_SR As Double 'Kutipan GST : SR
Dim Frm123_LM_GST_ZR As Double 'Kutipan GST : ZR
Dim Frm123_LM_JUMLAH_UPAH_SR As Double 'Total Harga Yang Dikenakan GST SR
Dim Frm123_LM_JUMLAH_UPAH_ZR As Double 'Total Harga Yang Dikenakan GST ZR
Dim Frm123_LM_BERAT As Double 'Berat Jualan
Dim frm123_LM_BERAT_ASAL As Double 'Berat Asal (Sebelum tukar kepada purity 999.9)

Frm123_PAGE_SIZE = 26
Frm123_LM_TOTAL_PAGE = 0
x = 0
Frm123_LM_UPAH_TANPA_GST = 0
Frm123_LM_UPAH_DENGAN_GST = 0
Frm123_LM_GST_SR = 0
Frm123_LM_GST_ZR = 0
Frm123_LM_JUMLAH_UPAH_SR = 0
Frm123_LM_JUMLAH_UPAH_ZR = 0
Frm123_LM_BERAT = 0
frm123_LM_BERAT_ASAL = 0 'Berat Asal (Sebelum tukar kepada purity 999.9)

re_gen_report:

frm123.L43_Text = x 'Jumlah bilangan barang jualan
frm123.L48_Text = Format(0, "#,##0.00") 'Jumlah berat jualan
frm123.L35_Text = Format(0, "#,##0.00") 'Maklumat GST : Jumlah harga ZR
frm123.L37_Text = Format(0, "#,##0.00") 'Maklumat GST : Jumlah harga SR
frm123.L36_Text = Format(0, "#,##0.00") 'Maklumat GST : Jumlah GST ZR
frm123.L38_Text = Format(0, "#,##0.00")  'Maklumat GST : Jumlah GST SR
frm123.L9_Text = Format(0, "#,##0.00") 'Berat jualan 999.9
frm123.L51_Text = Format(0, "#,##0.00") 'Jumlah upah tanpa GST (keseluruhan)
frm123.L52_Text = Format(0, "#,##0.00") 'Jumlah GST (Keseluruhan)
frm123.L53_Text = Format(0, "#,##0.00") 'Jumlah Upah + GST (Keseluruhan)

LM_START_ROW = frm123.L69_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm123_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If frm123.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm123_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        frm123.L67_Text = 1
    End If
End If

Frm123_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from " & G_GRN_TEMP & " where Status='" & 1 & "' Or Status='" & 2 & "' Or Status='" & 3 & "' Or Status='" & 4 & "' order by purity ASC LIMIT " & LM_START_ROW & "," & Frm123_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1
    If Frm123_LM_PAGE_FOUND = 0 Then
        If frm123.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                frm123.L67_Text = frm123.L67_Text + 1 'Paparan Page ke-xxx
                Frm123_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(frm123.L67_Text) Then
                    If frm123.L67_Text <> 1 Then
                        frm123.L67_Text = frm123.L67_Text - 1 'Paparan Page ke-xxx
                        Frm123_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    
    Y = ((frm123.L67_Text - 1) * Frm123_PAGE_SIZE) + x
    frm123.MSFlexGrid1.Rows = x + 1
    frm123.MSFlexGrid1.TextMatrix(x, 0) = x 'No.
    frm123.MSFlexGrid1.TextMatrix(x, 1) = Y 'No.
    frm123.MSFlexGrid1.ColAlignment(1) = 4
    frm123.MSFlexGrid1.TextMatrix(x, 2) = rs!ID 'No. ID

    If Not IsNull(rs!purity) Then frm123.MSFlexGrid1.TextMatrix(x, 3) = rs!purity 'Purity
    If Not IsNull(rs!Berat_Asal) Then frm123.MSFlexGrid1.TextMatrix(x, 4) = Format(rs!Berat_Asal, "#,##0.00") 'Berat Asal (g)
    If Not IsNull(rs!kadar_tukaran) Then frm123.MSFlexGrid1.TextMatrix(x, 5) = rs!kadar_tukaran 'Mutu
    If Not IsNull(rs!berat_tukaran_grn) Then frm123.MSFlexGrid1.TextMatrix(x, 6) = Format(rs!berat_tukaran_grn, "#,##0.00") 'Berat 999.9 (g)
    If Not IsNull(rs!UPAH) Then frm123.MSFlexGrid1.TextMatrix(x, 7) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
    If Not IsNull(rs!gst_ari_nashi) Then frm123.MSFlexGrid1.TextMatrix(x, 8) = rs!gst_ari_nashi 'Jenis GST
    If Not IsNull(rs!jumlah_gst) Then frm123.MSFlexGrid1.TextMatrix(x, 9) = Format(rs!jumlah_gst, "#,##0.00") 'Jumlah GST (RM)
    If Not IsNull(rs!harga_dengan_gst_grn) Then frm123.MSFlexGrid1.TextMatrix(x, 10) = Format(rs!harga_dengan_gst_grn, "#,##0.00") 'Upah + GST (RM)
    
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
    Frm123_LM_TOTAL_PAGE = Format(rs(0) / Frm123_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm123_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm123_LM_PAGE = Split(Frm123_LM_TOTAL_PAGE, ".")(0)
        Frm123_LM_PAGE_LEBIHAN = Split(Frm123_LM_TOTAL_PAGE, ".")(1)
        
        If Frm123_LM_PAGE_LEBIHAN <> "00" Then
            frm123.L68_Text = Frm123_LM_PAGE + 1
        Else
            frm123.L68_Text = Frm123_LM_PAGE
        End If
        
    Else
    
        frm123.L68_Text = Frm123_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        frm123.L68_Text = 0
    End If
Else
    frm123.L68_Text = 0
End If

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) , SUM(berat_tukaran_grn) , SUM(berat_tukaran_grn) , SUM(harga_tanpa_gst_grn) , SUM(jumlah_gst) , SUM(harga_dengan_gst_grn) from " & G_GRN_TEMP & " where Status='" & 1 & "' Or Status='" & 2 & "' Or Status='" & 3 & "' Or Status='" & 4 & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then frm123.L43_Text = rs(0) 'Jumlah bilangan barang jualan
If Not IsNull(rs(1)) Then frm123.L48_Text = Format(rs(1), "#,##0.00") 'Jumlah berat jualan
'If Not IsNull(rs(2)) Then Frm123.L9_Text = Format(rs(2), "#,##0.00") 'Berat jualan 999.9
If Not IsNull(rs(3)) Then frm123.L51_Text = Format(rs(3), "#,##0.00") 'Jumlah upah tanpa GST (keseluruhan)
If Not IsNull(rs(4)) Then frm123.L52_Text = Format(rs(4), "#,##0.00") 'Jumlah GST (Keseluruhan)
If Not IsNull(rs(5)) Then frm123.L53_Text = Format(rs(5), "#,##0.00") 'Jumlah Upah + GST (Keseluruhan)

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_gst) , SUM(harga_tanpa_gst_grn) from " & G_GRN_TEMP & " where (Status='" & 1 & "' Or Status='" & 2 & "' Or Status='" & 3 & "' Or Status='" & 4 & "') AND gst_ari_nashi='" & "ZR" & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then frm123.L36_Text = Format(rs(0), "#,##0.00") 'Maklumat GST : Jumlah GST ZR
If Not IsNull(rs(1)) Then frm123.L35_Text = Format(rs(1), "#,##0.00") 'Maklumat GST : Jumlah harga ZR

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(jumlah_gst) , SUM(harga_tanpa_gst_grn) from " & G_GRN_TEMP & " where (Status='" & 1 & "' Or Status='" & 2 & "' Or Status='" & 3 & "' Or Status='" & 4 & "') AND gst_ari_nashi='" & "SR" & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then frm123.L38_Text = Format(rs(0), "#,##0.00") 'Maklumat GST : Jumlah GST SR
If Not IsNull(rs(1)) Then frm123.L37_Text = Format(rs(1), "#,##0.00") 'Maklumat GST : Jumlah harga SR

rs.Close
Set rs = Nothing

If x <> 0 Then
    frm123.L69_Text = LM_START_ROW
End If

If frm123.L67_Text <> vbNullString And IsNumeric(frm123.L67_Text) Then
    If frm123.L68_Text <> vbNullString And IsNumeric(frm123.L68_Text) Then
        Frm123_LM_CURR_PAGE = frm123.L67_Text
        Frm123_LM_TOTAL_PAGE = frm123.L68_Text
        
        If Frm123_LM_CURR_PAGE > Frm123_LM_TOTAL_PAGE Then
            
            frm123.L67_Text = frm123.L67_Text - 1
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            GoTo re_gen_report:
            
        End If
    End If
End If

End Sub
Sub Frm123_calc10()
'On Error Resume Next
Dim Frm123_LM_HARGA_ZR_UPAH As Double
Dim Frm123_LM_HARGA_SR_UPAH As Double
Dim Frm123_LM_HARGA_ZR_EMAS As Double
Dim Frm123_LM_HARGA_SR_EMAS As Double
Dim Frm123_LM_GST_ZR_UPAH As Double
Dim Frm123_LM_GST_SR_UPAH As Double
Dim Frm123_LM_GST_ZR_EMAS As Double
Dim Frm123_LM_GST_SR_EMAS As Double

Frm123_LM_HARGA_ZR_UPAH = 0
Frm123_LM_HARGA_SR_UPAH = 0
Frm123_LM_HARGA_ZR_EMAS = 0
Frm123_LM_HARGA_SR_EMAS = 0
Frm123_LM_GST_ZR_UPAH = 0
Frm123_LM_GST_SR_UPAH = 0
Frm123_LM_GST_ZR_EMAS = 0
Frm123_LM_GST_SR_EMAS = 0

If ((frm123.L35_Text <> vbNullString And IsNumeric(frm123.L35_Text)) And (frm123.L39_Text <> vbNullString And IsNumeric(frm123.L39_Text))) Then

    Frm123_LM_HARGA_ZR_UPAH = frm123.L35_Text 'Harga ZR (Upah)
    'Frm123_LM_HARGA_ZR_EMAS = Frm123.L39_Text 'Harga ZR (Emas)
    
    frm123.L17_Text = Format(Frm123_LM_HARGA_ZR_UPAH + Frm123_LM_HARGA_ZR_EMAS, "#,##0.00") 'Jumlah Harga ZR
    
Else

    frm123.L17_Text = "0.00" 'Jumlah Harga ZR
    
End If

If ((frm123.L37_Text <> vbNullString And IsNumeric(frm123.L37_Text)) And (frm123.L41_Text <> vbNullString And IsNumeric(frm123.L41_Text))) Then

    Frm123_LM_HARGA_SR_UPAH = frm123.L37_Text 'Harga SR (Upah)
    'Frm123_LM_HARGA_SR_EMAS = Frm123.L41_Text 'Harga SR (Emas)
    
    frm123.L18_Text = Format(Frm123_LM_HARGA_SR_UPAH + Frm123_LM_HARGA_SR_EMAS, "#,##0.00") 'Jumlah Harga SR
    
Else

    frm123.L18_Text = "0.00" 'Jumlah Harga SR
    
End If

If ((frm123.L36_Text <> vbNullString And IsNumeric(frm123.L36_Text)) And (frm123.L40_Text <> vbNullString And IsNumeric(frm123.L40_Text))) Then

    Frm123_LM_GST_SR_UPAH = frm123.L36_Text 'GST ZR (Upah)
    'Frm123_LM_GST_SR_EMAS = Frm123.L40_Text 'GST ZR (Emas)
    
    frm123.L20_Text = Format(Frm123_LM_GST_SR_UPAH + Frm123_LM_GST_SR_EMAS, "#,##0.00") 'Jumlah GST ZR
    
Else

    frm123.L20_Text = "0.00" 'Jumlah GST ZR
    
End If

If ((frm123.L38_Text <> vbNullString And IsNumeric(frm123.L38_Text)) And (frm123.L42_Text <> vbNullString And IsNumeric(frm123.L42_Text))) Then

    Frm123_LM_GST_ZR_UPAH = frm123.L38_Text 'GST SR (Upah)
    'Frm123_LM_GST_ZR_EMAS = Frm123.L42_Text 'GST SR (Emas)
    
    frm123.L20_Text = Format(Frm123_LM_GST_ZR_UPAH + Frm123_LM_GST_ZR_EMAS, "#,##0.00") 'Jumlah GST SR
    
Else

    frm123.L20_Text = "0.00" 'Jumlah GST SR
    
End If

frm123.L15_Text = Format(Frm123_LM_HARGA_ZR_UPAH + Frm123_LM_HARGA_ZR_EMAS + Frm123_LM_HARGA_SR_UPAH + Frm123_LM_HARGA_SR_EMAS, "#,##0.00") 'Jumlah harga tanpa GST
frm123.L16_Text = Format(Frm123_LM_HARGA_ZR_UPAH + Frm123_LM_HARGA_ZR_EMAS + Frm123_LM_HARGA_SR_UPAH + Frm123_LM_HARGA_SR_EMAS + Frm123_LM_GST_SR_UPAH + Frm123_LM_GST_SR_EMAS + Frm123_LM_GST_ZR_UPAH + Frm123_LM_GST_ZR_EMAS, "#,##0.00") 'Jumlah harga dengan GST
End Sub
Sub Frm123_calc11()
'On Error Resume Next
Dim frm123_LM_BERAT_ASAL As Double
Dim Frm123_LM_MUTU As Double

frm123_LM_BERAT_ASAL = 0 'Beza berat (g)
Frm123_LM_MUTU = 0 'Harga semasa (RM/g)

If ((frm123.L48_Text <> vbNullString And IsNumeric(frm123.L48_Text)) And (frm123.TB8 <> vbNullString And IsNumeric(frm123.TB8))) Then
    frm123_LM_BERAT_ASAL = frm123.L48_Text 'Berat jualan (g)
    Frm123_LM_MUTU = frm123.TB8 'Kadar belian (g)
    
    frm123.L9_Text = Format(frm123_LM_BERAT_ASAL * Frm123_LM_MUTU, "#,##0.00") 'Harga jualan
Else
    frm123.L9_Text = "0.00" 'Harga jualan
End If
End Sub
Sub frm123_periksa_baki_berat()
'On Error Resume Next
Dim LM_BERAT_ASAL As Double
Dim LM_BERAT_GUNA As Double
Dim LM_BERAT_TEMP As Double
Dim LM_BERAT_TEMP_ASAL As Double

LM_BERAT_ASAL = 0
LM_BERAT_GUNA = 0
LM_BERAT_TEMP = 0
LM_BERAT_TEMP_ASAL = 0

frm123.TB10 = Format(0, "#,##0.00")

Set rs3 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs3.Open "select SUM(beza_berat) from data_database where Purity='" & frm123.CBB1 & "' AND (((statusitem = 10 OR statusitem = 12 OR statusitem = 20 OR statusitem = 22) AND receiving_Status = 2) OR ((statusitem = 12 OR statusitem = 20 OR statusitem = 22) AND receiving_Status = 0)) AND cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs3(0)) Then LM_BERAT_ASAL = Format(rs3(0), "#,##0.00")
    
rs3.Close
Set rs3 = Nothing

Set rs3 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs3.Open "select SUM(berat) from 85_penggunaan_ti where purity='" & frm123.CBB1 & "' AND status = 1 AND cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic
    
If Not IsNull(rs3(0)) Then LM_BERAT_GUNA = Format(rs3(0), "#,##0.00")
    
rs3.Close
Set rs3 = Nothing

'Set rs3 = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs3.Open "select SUM(Berat_Asal) from " & G_GRN_TEMP & " where purity='" & frm123.CBB1 & "' AND (status = 2 OR status = 3 OR status = 4)", cn, adOpenKeyset, adLockOptimistic
    
'If Not IsNull(rs3(0)) Then LM_BERAT_TEMP = Format(rs3(0), "#,##0.00")
    
'rs3.Close
'Set rs3 = Nothing

Set rs3 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs3.Open "select SUM(Berat_Asal) , SUM(berat_bef_edit) from " & G_GRN_TEMP & " where purity='" & frm123.CBB1 & "' AND (status = 1 OR status = 2 OR status = 3 OR status = 4)", cn, adOpenKeyset, adLockOptimistic
    
If Not IsNull(rs3(0)) Then LM_BERAT_TEMP = Format(rs3(0), "#,##0.00")
If Not IsNull(rs3(1)) Then LM_BERAT_TEMP_ASAL = Format(rs3(1), "#,##0.00")

rs3.Close
Set rs3 = Nothing

frm123.TB10 = Format(LM_BERAT_ASAL - LM_BERAT_GUNA - LM_BERAT_TEMP + LM_BERAT_TEMP_ASAL, "#,##0.00")
End Sub
Sub frm123_berat_edit()
'On Error Resume Next
Dim LM_BERAT_TOTAL As Double
Dim LM_BERAT_GUNA As Double

LM_BERAT_TOTAL = 0
LM_BERAT_GUNA = 0

Set rs3 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs3.Open "select SUM(Berat_Asal) from " & G_GRN_TEMP & " where ID='" & frm123.L2_Text & "'", cn, adOpenKeyset, adLockOptimistic
    
If Not IsNull(rs3(0)) Then LM_BERAT_GUNA = Format(rs3(0), "#,##0.00")
    
rs3.Close
Set rs3 = Nothing

If frm123.TB10 <> vbNullString And IsNumeric(frm123.TB10) Then LM_BERAT_TOTAL = frm123.TB10

frm123.TB10 = Format(LM_BERAT_TOTAL + LM_BERAT_GUNA, "#,##0.00")
End Sub
Sub frm123_edit_data_gdn_bulk()
'on error resume next
LM_FOUND = 0

If G_No_RESIT_JUALAN <> vbNullString Then
    
    frm123_LM_USER = vbNullString
    
    frm123.L69_Text = -1 'Titik Pencarian Data
    frm123.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
    frm123.L67_Text = 0 'Paparan Page ke-xxx
    
    Call Frm123_one_time_reset
    Call Frm123_reset_1
    Call Frm123_reset_3

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    
    strsql = "insert into " & G_GRN_TEMP & "(id_database,purity,berat_asal," _
                & "berat_tukaran_grn,upah,harga_tanpa_gst_grn,gst_ari_nashi," _
                & "gst_include,jumlah_gst,kadar_gst,harga_dengan_gst_grn,nilaian_harga_emas," _
                & "kadar_tukaran,berat_bef_edit,Status)" & _
                "select ID,purity,berat_asal," _
                & "berat_tukaran_grn,upah,harga_tanpa_gst_grn,gst_ari_nashi," _
                & "gst_include,jumlah_gst,kadar_gst,harga_dengan_gst_grn,nilaian_harga_emas," _
                & "kadar_tukaran,berat_asal,2 " _
                & "from 79_grn WHERE no_rujukan='" & G_No_RESIT_JUALAN & "' AND status = 1"

    Set rs = cn.Execute(strsql)
    Set rs = Nothing
    
    Call Frm123_Senarai_Belian_Header
    Call Frm123_Senarai_Belian

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 77_gdn_grn where no_rujukan='" & G_No_RESIT_JUALAN & "' AND jenis_urusan = 4 AND status = 1", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        
        If Not IsNull(rs!ID) Then frm123.L71_Text = rs!ID
        If Not IsNull(rs!tarikh) Then frm123.DTPicker1 = rs!tarikh
        If Not IsNull(rs!Berat_Asal) Then frm123.L48_Text = Format(rs!Berat_Asal, "#,##0.00") 'Berat asal sebelum tukaran mutu
        If Not IsNull(rs!kadar_tukaran) Then frm123.TB8 = rs!kadar_tukaran
        If Not IsNull(rs!berat_tukaran) Then frm123.L9_Text = Format(rs!berat_tukaran, "#,##0.00")
        If Not IsNull(rs!harga_tanpa_gst) Then frm123.L51_Text = Format(rs!harga_tanpa_gst, "#,##0.00")
        If Not IsNull(rs!jumlah_gst) Then frm123.L52_Text = Format(rs!jumlah_gst, "#,##0.00")
        If Not IsNull(rs!kadar_gst) Then frm123.L22_Text = Format(rs!kadar_gst, "#,##0.00")
        If Not IsNull(rs!harga_dengan_gst) Then frm123.L53_Text = Format(rs!harga_dengan_gst, "#,##0.00")
        If Not IsNull(rs!harga_999) Then frm123.TB6 = Format(rs!harga_999, "#,##0.00")
        If Not IsNull(rs!nilaian_harga_emas) Then frm123.L12_Text = Format(rs!nilaian_harga_emas, "#,##0.00")
        If Not IsNull(rs!gst_zr_harga) Then frm123.L17_Text = Format(rs!gst_zr_harga, "#,##0.00")
        If Not IsNull(rs!gst_sr_harga) Then frm123.L18_Text = Format(rs!gst_sr_harga, "#,##0.00")
        If Not IsNull(rs!gst_zr_cukai) Then frm123.L19_Text = Format(rs!gst_zr_cukai, "#,##0.00")
        If Not IsNull(rs!gst_sr_cukai) Then frm123.L20_Text = Format(rs!gst_sr_cukai, "#,##0.00")
        If Not IsNull(rs!bil_barang) Then frm123.L43_Text = rs!bil_barang
        If Not IsNull(rs!no_rujukan_supplier) Then frm123.TB9 = rs!no_rujukan_supplier
        
        If Not IsNull(rs!supplier_agen) Then
            'on error goto Err_A:
            frm123_LM_SUPPLIER = rs!supplier_agen
            frm123.CBB2 = frm123_LM_SUPPLIER
        
Restore_A:
        End If
        
        If Not IsNull(rs!user) Then

            frm123_LM_USER = rs!user

        End If
        'on error resume next
        LM_FOUND = 1
        
    End If
    
    rs.Close
    Set rs = Nothing
'### Masukkan data voucher / invoice bagi belian agen ini ### - End

'### Carian Maklumat Penjual (Data Pekerja) ### - Start
    If frm123_LM_USER <> vbNullString Then
    
        DATA_PEKERJA_FOUND = 0
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where Samaran='" & frm123_LM_USER & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
        
            frm123_LM_MAKLUMAT_PEKERJA = rs!Samaran & "  |  " & rs!NoPekerja
            DATA_PEKERJA_FOUND = 1
            
        End If
        
        rs.Close
        Set rs = Nothing
    
        If DATA_PEKERJA_FOUND = 1 Then
            'On Error GoTo Err_B:
            frm123.CBB4 = frm123_LM_MAKLUMAT_PEKERJA
            
Restore_B:
        End If
        
        'on error resume next
    End If
'### Carian Maklumat Penjual (Data Pekerja) ### - End

    If LM_FOUND = 1 Then
    
        frm123.CBB4.Enabled = True
        frm123.CBB4.BackColor = &HFFFFFF
    
        frm123.CMD8.Visible = False
        frm123.CMD9.Visible = False
        frm123.CMD10.Visible = True
        frm123.CMD11.Visible = True
        
        frm123.Show
        frm117.Hide
        
    End If

End If
     
Exit Sub

Err_A:

frm123.CBB2.AddItem frm123_LM_SUPPLIER
frm123.CBB2 = frm123_LM_SUPPLIER
            
Resume Restore_A:

Exit Sub
Err_B:
frm123.CBB4.AddItem frm123_LM_MAKLUMAT_PEKERJA
frm123.CBB4 = frm123_LM_MAKLUMAT_PEKERJA
Resume Restore_B:
End Sub
Sub frm123_padam_gdn_bulk()
'On Error Resume Next
'### Masukkan maklumat Good Delivery Note (GRN) ### - Start
DATA_SAVE = 0

LM_NOW = Now
LM_TARIKH = DateTime.Date$
LM_MASA = DateTime.Time$
LM_NO_RUJUKAN = vbNullString

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 77_gdn_grn where ID='" & G_ID & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    G_ID = rs!ID
    Call recovery_77_gdn_grn
    
    rs!Status = 0
    'rs!jenis_urusan = 1
    rs!terminal = G_TERMINAL
    rs!user = MDI_frm1.L3_Text 'Nama Pekerja
    rs.Update
    DATA_SAVE = 1
    
End If

rs.Close
Set rs = Nothing
'### Masukkan data voucher / invoice bagi belian agen ini ### - End

'### Masukkan Data Jualan Ke Dalam Table Jualan ### - Start

If DATA_SAVE = 1 Then
    
    '### Transfer data kepada recovery database ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

    strsql = "insert into " & G_RECOVERY_DATABASE & ".79_grn(id_asal,tarikh,masa,write_timestamp,no_rujukan,no_rujukan_rev,purity,berat_asal," _
                & "berat_tukaran_grn,upah,harga_tanpa_gst_grn,gst_ari_nashi," _
                & "gst_include,jumlah_gst,kadar_gst,harga_dengan_gst_grn,nilaian_harga_emas," _
                & "kadar_tukaran,Status,terminal,user)" & _
                "select ID,tarikh,masa,write_timestamp,no_rujukan,no_rujukan_rev,purity,berat_asal," _
                & "berat_tukaran_grn,upah,harga_tanpa_gst_grn,gst_ari_nashi," _
                & "gst_include,jumlah_gst,kadar_gst,harga_dengan_gst_grn,nilaian_harga_emas," _
                & "kadar_tukaran,Status,terminal,user " _
                & "from " & G_SERVER_DATABASE & ".79_grn WHERE no_rujukan='" & G_No_RESIT_JUALAN & "' AND status = 1"
                
    Set rs = cn.Execute(strsql)
    Set rs = Nothing
    '### Transfer data kepada recovery database ### - End
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

    strsql = "UPDATE 79_grn set status='" & 0 & "'," _
    & "user='" & MDI_frm1.L3_Text & "'," _
    & "terminal='" & G_TERMINAL & "'" _
    & "WHERE no_rujukan='" & G_No_RESIT_JUALAN & "' AND status = 1"
    
    Set rs = cn.Execute(strsql)
    Set rs = Nothing
    
    '### Transfer data kepada recovery database ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

    strsql = "insert into " & G_RECOVERY_DATABASE & ".85_penggunaan_ti(id_asal,tarikh,no_rujukan,purity,berat,write_timestamp,terminal,Status)" & _
                "select ID,tarikh,no_rujukan,purity,berat,write_timestamp,terminal,Status " _
                & "from " & G_SERVER_DATABASE & ".85_penggunaan_ti WHERE no_rujukan='" & LM_NO_RUJUKAN & "' AND status = 1"
                
    Set rs = cn.Execute(strsql)
    Set rs = Nothing
    '### Transfer data kepada recovery database ### - End
            
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

    strsql = "UPDATE 85_penggunaan_ti set status='" & 0 & "'," _
    & "write_timestamp='" & LM_NOW & "'," _
    & "terminal='" & G_TERMINAL & "'" _
    & "WHERE no_rujukan='" & G_No_RESIT_JUALAN & "' AND status = 1"
    
    Set rs = cn.Execute(strsql)
    Set rs = Nothing

'#### Update Log Aktiviti Sistem #### - Start
    'User = MDI_frm1.L3_Text
    LogAct_Memory = "[" & MDI_frm1.L3_Text & "] Padam data GDN kepada agen/supplier (bulk). No. Rujukan [" & G_No_RESIT_JUALAN & "]."
    LogDate_Memory = LM_NOW
    Call UpdateLog_Database
'#### Update Log Aktiviti Sistem #### - End
    
    GM_NEXT_PREV = 2
    
    Call frm117_report_gdn_grn_header
    Call frm117_report_gdn_grn

    MsgBox "Data GDN telah berjaya dipadamkan.", vbInformation, "Info"

End If
End Sub




