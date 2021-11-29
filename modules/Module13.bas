Attribute VB_Name = "Module13"
Sub Frm109_analisa_harga_emas()
'On Error Resume Next
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim LM_BERAT As Double
Dim LM_HARGA As Double
Dim a As String
Dim b As String
Dim c As String

Frm109.L1_Text = vbNullString
Frm109.L2_Text = vbNullString
Frm109.L3_Text = vbNullString

Set rs1 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs1.Open "select * from hargaemas", cn, adOpenKeyset, adLockOptimistic

While rs1.EOF = False

    LM_BERAT = 0
    LM_HARGA = 0

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select SUM(berat) from Data_Database where Kod_Purity='" & rs1!purity & "' AND StatusItem='" & "10" & "' AND (receiving_Status='" & 0 & "' OR receiving_Status='" & 4 & "')", cn, adOpenKeyset, adLockOptimistic
    
    If Not IsNull(rs(0)) Then
        If IsNumeric(rs(0)) Then LM_BERAT = rs(0)
    End If
    
    rs.Close
    Set rs = Nothing
    
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select SUM(harga_item) from Data_Database where Kod_Purity='" & rs1!purity & "' AND StatusItem='" & "10" & "' AND (receiving_Status='" & 0 & "' OR receiving_Status='" & 4 & "')", cn, adOpenKeyset, adLockOptimistic
    
    If Not IsNull(rs(0)) Then
        If IsNumeric(rs(0)) Then LM_HARGA = rs(0)
    End If
    
    rs.Close
    Set rs = Nothing
    
    a = a & rs1!purity & vbCrLf
    
    b = b & Format(LM_BERAT, "#,##0.00 g") & vbCrLf
    
    If LM_BERAT <> 0 Then
        If IsNumeric(LM_BERAT) Then
            c = c & "RM " & Format(LM_HARGA / LM_BERAT, "#,##0.00") & "/g" & vbCrLf
        Else
            c = c & "RM 0.00/g" & vbCrLf
        End If
    Else
        c = c & "RM 0.00/g" & vbCrLf
    End If
    

    rs1.MoveNext
Wend

rs1.Close
Set rs1 = Nothing

Frm109.L1_Text = a

Frm109.L2_Text = b

Frm109.L3_Text = c
End Sub
Sub Frm2_AnalystSpot()
'On Error Resume Next
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset

Dim Frm2_LM_Frm2_LM_BERAT_STOK As Double
Dim Frm2_LM_HARGA_PERITEM As Double
Dim Frm2_LM_TOTAL_BERAT As Double
Dim Frm2_LM_TOTAL_HARGA As Double
Dim DATA1(100) 'Purity
Dim DATA2(100) 'Total Berat Stok
Dim DATA3(100) 'Total Harga
Dim DATA4(100) 'Harga MKS

Set rs1 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs1.Open "select * from hargaemas", cn, adOpenKeyset, adLockOptimistic

While rs1.EOF = False
    i = i + 1
    Frm2_LM_BERAT_STOK = 0
    Frm2_LM_HARGA_PERITEM = 0
    Frm2_LM_TOTAL_BERAT = 0
    Frm2_LM_TOTAL_HARGA = 0

    Set rs = New ADODB.Recordset
    rs.Open "select * from Data_Database where Kod_Purity='" & rs1!purity & "' AND StatusItem='" & "10" & "' AND (receiving_Status='" & 0 & "' OR receiving_Status='" & 4 & "')", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        'If rs!receiving_Status = "0" Or rs!receiving_Status = "2" Then
            'If IsNumeric(rs!beza_berat) Then Frm2_LM_BERAT_STOK = rs!beza_berat
            If IsNumeric(rs!Berat) Then Frm2_LM_BERAT_STOK = rs!Berat
            If IsNumeric(rs!harga_Per_Gram_Item) Then Frm2_LM_HARGA_PERITEM = rs!harga_Per_Gram_Item
            Frm2_LM_TOTAL_BERAT = Frm2_LM_TOTAL_BERAT + Frm2_LM_BERAT_STOK
            Frm2_LM_TOTAL_HARGA = Frm2_LM_TOTAL_HARGA + (Frm2_LM_BERAT_STOK * Frm2_LM_HARGA_PERITEM)
        'End If
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
    DATA1(i) = rs1!purity 'Purity
    DATA2(i) = Format(Frm2_LM_TOTAL_BERAT, "#,##0.00") 'Total Berat
    DATA3(i) = Format(Frm2_LM_TOTAL_HARGA, "#,##0.00") 'Total Harga
    DATA4(i) = Format(rs1!HargaMKS, "0.00") 'Total Harga
    rs1.MoveNext
Wend

rs1.Close
Set rs1 = Nothing

Frm2.MSFlexGrid2.Clear
Frm2.MSFlexGrid2.RowHeight(0) = 500
Frm2.MSFlexGrid2.FormatString = "<No.|<Purity|<Harga Jualan (RM/g)|<Berat (g)|<Kos Per Gram (RM/g)"

Frm2.MSFlexGrid2.Rows = 1
Frm2.MSFlexGrid2.ColWidth(0) = 500
Frm2.MSFlexGrid2.ColWidth(1) = 2200
Frm2.MSFlexGrid2.ColWidth(2) = 1500
Frm2.MSFlexGrid2.ColWidth(3) = 1700
Frm2.MSFlexGrid2.ColWidth(4) = 1500

For k = 1 To i
    x = x + 1
    Frm2.MSFlexGrid2.Rows = x + 1
    Frm2.MSFlexGrid2.TextMatrix(x, 0) = x
    Frm2.MSFlexGrid2.TextMatrix(x, 1) = DATA1(k) 'Purity
    Frm2.MSFlexGrid2.TextMatrix(x, 2) = DATA4(k) 'Harga MKS
    Frm2.MSFlexGrid2.TextMatrix(x, 3) = DATA2(k) 'Berat
    If DATA2(k) = 0 Then
        ANALISIS_HARGA = Format(0, "0.00")
    Else
        ANALISIS_HARGA = Format(DATA3(k) / DATA2(k), "#,##0.00")
    End If
    Frm2.MSFlexGrid2.TextMatrix(x, 4) = ANALISIS_HARGA 'Kos Per Gram
Next k

Frm2.L8_Text = "Update Terkini : " & DateTime.Date & " " & DateTime.Time
End Sub
Sub frm114_kalkulator_ti()
'On Error Resume Next
Dim LM_HARGA_BESAR As Double
Dim LM_TAEL As Double
Dim LM_SA As Double
Dim LM_PURITY As Double
Dim LM_PUBLIC As Double
Dim LM_ASSAY As Double

Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim f As String

LM_HARGA_BESAR = 0
LM_TAEL = 0
LM_SA = 0
LM_PUBLIC = 0
LM_ASSAY = 0

If IsNumeric(Frm114.TB1) Then LM_HARGA_BESAR = Frm114.TB1
If IsNumeric(Frm114.L2_Text) Then LM_SA = Frm114.L2_Text
If IsNumeric(Frm114.L3_Text) Then LM_TAEL = Frm114.L3_Text
If IsNumeric(Frm114.L6_Text) Then LM_PUBLIC = Frm114.L6_Text

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from setting_database where status = 1 AND trade_in <>'" & Null & "'", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    LM_PURITY = 0
    LM_ASSAY = 0
    
    If Not IsNull(rs!Kod_Metal_Purity) Then

        If Not IsNull(rs!trade_in) And Not IsNull(rs!assay) Then
            If IsNumeric(rs!trade_in) Then
                LM_PURITY = rs!trade_in * 100
                LM_ASSAY = rs!assay * 100
                
                LM_CALC_SA = (((LM_HARGA_BESAR - LM_SA) * LM_TAEL) * LM_PURITY) / 100
                LM_CALC_PUBLIC = (((LM_HARGA_BESAR - LM_PUBLIC) * LM_TAEL) * LM_PURITY) / 100
                LM_CALC_UNTUNG_TOLAK = LM_CALC_SA - LM_CALC_PUBLIC
                LM_CALC_UNTUNG_MUTU = (((LM_HARGA_BESAR * LM_TAEL) * LM_ASSAY) / 100) - (((LM_HARGA_BESAR * LM_TAEL) * LM_PURITY) / 100)
                
                a = a & rs!Kod_Metal_Purity & vbCrLf
                
                b = b & "RM " & Format(LM_CALC_SA, "#,##0.00") & " /g" & vbCrLf
                
                c = c & "RM " & Format(LM_CALC_PUBLIC, "#,##0.00") & " /g" & vbCrLf
                
                d = d & "RM " & Format(LM_CALC_UNTUNG_TOLAK, "#,##0.00") & " /g" & vbCrLf
                
                e = e & "RM " & Format(LM_CALC_UNTUNG_MUTU, "#,##0.00") & " /g" & vbCrLf
                
                f = f & "RM " & Format(LM_CALC_UNTUNG_TOLAK + LM_CALC_UNTUNG_MUTU, "#,##0.00") & " /g" & vbCrLf
                
            End If
            
        End If
        
    End If

    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm114.L4_Text = a

Frm114.L5_Text = b

Frm114.L7_Text = c

Frm114.L10_Text = d

Frm114.L12_Text = e

Frm114.L14_Text = f
End Sub
