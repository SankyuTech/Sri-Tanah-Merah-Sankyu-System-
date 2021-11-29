Attribute VB_Name = "DrawBar"
Public BARCODE_BERAT
Public BARCODE_SUPPLIER
Public BARCODE_HARGA
Public BARCODE_UPAH
Public BARCODE_NOSIRI
Public BARCODE_DULANG
Public BARCODE_MARKET
Public BARCODE_TYPE
Public BARCODE_MODAL
Public BARCODE_UPAH2
Public BARCODE_TARIKH
Public BARCODE_Panjang
Public BARCODE_Lebar
Public BARCODE_Dia
Public BARCODE_Saiz
Public BARCODE_PURITY
Public BARCODE_UPAH30
Public BARCODE_RIYAL
Public BARCODE_CODE1
Public BARCODE_CODE2
Public G_No_Inv
Public G_VOUCHER
Public G_MAIL
Public G_INVOICE_AHLI
Public G_No_DO
Public G_No_RESIT_GB
Public G_No_RESIT_BUYBACK_GB
Public GB_BARCODE
Public GB_BERAT
Public GB_DULANG
Public GB_PANJANG
Public GB_LEBAR
Public GB_TEBAL
Public GB_PURITY
Public GB_CERT
Public GM_No_RUJUKAN_BELIAN
Public G_No_RESIT_JUALAN
Public G_No_RESIT_ANSURAN
Public G_No_RESIT_SERVIS
Public G_No_INV_BOOK
Public G_No_RUJ_HIBAH
Public G_PAYSLIP_BULAN
Public G_PAYSLIP_IC
Public G_No_STATMENT_FORM
Public G_No_RUJUKAN_FORM
Public G_PENYATA_AMBILAN
Public G_PENYATA_PULANGAN
Public G_TARIKH_HIBAH As Date
Public GM_KATEGORI As Integer
Public G_RANKING_FIELD As String

Public G_Frm56_LM_SIZE(2)
Public G_Frm56_LM_SIZE_2
Public G_Frm56_LM_SIZE_3
Public G_Frm56_LM_SIZE_4
Public G_Frm56_LM_SIZE_0
Public G_Frm56_LM_TYPE_0
Public G_SKU_LINE(2)
Public G_BAROCDE_LINE_2
Public G_BAROCDE_LINE_3
Public G_BAROCDE_LINE_4
Public G_Frm56_LM_TYPE(2)
Public G_Frm56_LM_TYPE_2
Public G_Frm56_LM_TYPE_3
Public G_Frm56_LM_TYPE_4
Public G_Frm56_LM_TYPE_5
Public G_Frm56_LM_SIZE_5
Public G_SKU_SIZE(2)
Public G_SKU_FONT(2)
Public G_SKU_POS_X(2)
Public G_SKU_POS_Y(2)
Public G_SKU_BOLD(2)
Public G_SKU_ITALIC(2)
Public G_SKU_DATA(2)
Public G_SKU_PRE_DATA(2)
Public G_FOOTER_A
Public G_HEADER_A
Public G_FIELD
Sub ClearPrinterMemory()
'On Error Resume Next
BARCODE_SUPPLIER = vbNullString
BARCODE_BERAT = vbNullString
BARCODE_HARGA = vbNullString
BARCODE_NOSIRI = vbNullString
BARCODE_UPAH = vbNullString
BARCODE_DULANG = vbNullString
BARCODE_MARKET = vbNullString
BARCODE_TYPE = vbNullString
BARCODE_MODAL = vbNullString
BARCODE_UPAH2 = vbNullString
BARCODE_Panjang = vbNullString
BARCODE_Lebar = vbNullString
BARCODE_Dia = vbNullString
BARCODE_Saiz = vbNullString
BARCODE_TARIKH = vbNullString
BARCODE_PURITY = vbNullString
BARCODE_UPAH30 = vbNullString
BARCODE_RIYAL = vbNullString
BARCODE_CODE1 = vbNullString
BARCODE_CODE2 = vbNullString
GB_BARCODE = vbNullString
GB_BERAT = vbNullString
GB_DULANG = vbNullString
GB_PANJANG = vbNullString
GB_LEBAR = vbNullString
GB_TEBAL = vbNullString
GB_PURITY = vbNullString
GB_CERT = vbNullString
End Sub
Sub Print_All_Barcode()
'On Error Resume Next
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset

Dim JENIS_BARCODE(12)
Dim LOC_(12)
Dim VALUE_BARCODE(12)
Dim LM_NAMA_KEDAI As String

Frm56x_LM_SIZE_1 = 6 'Saiz tulisan bagi barisan pertama
Frm56x_LM_SIZE_2 = 6 'Saiz tulisan bagi barisan kedua
Frm56x_LM_SIZE_3 = 6 'Saiz tulisan bagi barisan ketiga
Frm56x_LM_SIZE_4 = 6 'Saiz tulisan bagi barisan keempat

PRINTER_FOUND = 0 '0 : Barcode Label Printer Not Found , 1 : Barcode Label Printer Found
For Each oPrn In Printers
    'If oPrn.DeviceName = "ZDesigner GT800 (EPL)" Then
    If oPrn.DeviceName = G_PRINTER_BARCODE Then
        Set Printer = oPrn
        PRINTER_FOUND = 1 '0 : Barcode Label Printer Not Found , 1 : Barcode Label Printer Found
        Exit For
    End If
Next

LM_HEAD = "Sankyu System"
LM_NAMA_KEDAI = G_LM_NAMA_KEDAI
LM_HEAD = G_LM_HEAD

'### Maklumat kedai ### - Start
'Set rs = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

'If Not rs.EOF Then
'    If Not IsNull(rs!nama_kedai_3) Then LM_NAMA_KEDAI = rs!nama_kedai_3
'    If Not IsNull(rs!bar_header) Then LM_HEAD = rs!bar_header
'End If

'rs.Close
'Set rs = Nothing
'### Maklumat kedai ### - End

'##########Layout Barcode##################
    
L_JENIS_BARCODE = 0

Frm56x_LM_SIZE_1 = G_Frm56_LM_SIZE_1
Frm56x_LM_SIZE_2 = G_Frm56_LM_SIZE_2
Frm56x_LM_SIZE_3 = G_Frm56_LM_SIZE_3
Frm56x_LM_SIZE_4 = G_Frm56_LM_SIZE_4
L_JENIS_BARCODE = G_L_JENIS_BARCODE

'Set rs1 = New ADODB.Recordset
'If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
'rs1.Open "select * from layout_Barcode where perkara='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

'If Not rs1.EOF Then
'    If rs1!Line1 = 1 Then G_BAROCDE_LINE_1 = 1
'    If rs1!Line2 = 1 Then G_BAROCDE_LINE_2 = 1
'    If rs1!Line3 = 1 Then G_BAROCDE_LINE_3 = 1
'    If rs1!Line4 = 1 Then G_BAROCDE_LINE_4 = 1
    
'    If Not IsNull(rs1!font_size_1) Then Frm56x_LM_SIZE_1 = rs1!font_size_1 'Saiz tulisan bagi barisan pertama
'    If Not IsNull(rs1!font_size_2) Then Frm56x_LM_SIZE_2 = rs1!font_size_2 'Saiz tulisan bagi barisan kedua
'    If Not IsNull(rs1!font_size_3) Then Frm56x_LM_SIZE_3 = rs1!font_size_3 'Saiz tulisan bagi barisan ketiga
'    If Not IsNull(rs1!font_size_4) Then Frm56x_LM_SIZE_4 = rs1!font_size_4 'Saiz tulisan bagi barisan keempat
    
'    If Not IsNull(rs1!BARCODE_TYPE) Then
        
'        If rs1!BARCODE_TYPE = 0 Then
'            L_JENIS_BARCODE = 0
'        ElseIf rs1!BARCODE_TYPE = 1 Then
'            L_JENIS_BARCODE = 1
'        End If
    
'    End If
'
'End If

'rs1.Close
'Set rs1 = Nothing
            
If PRINTER_FOUND = 1 Then '0 : Barcode Label Printer Not Found , 1 : Barcode Label Printer Found

    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from Data_Database where NoRujukanSistem='" & GM_No_RUJUKAN_BELIAN & "' AND StatusItem <> 0", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
    
        Call ClearPrinterMemory
        
        LM_UPAH = 0
        L_BERAT = vbNullString
        L_TRAY = vbNullString
        L_PURITY = purity
        L_HARGA = 0
        
        LM_GST = 0
        
        If Not IsNull(rs!gst_ari_nashi) Then
            If rs!gst_ari_nashi = 0 Then
                LM_GST = 0
            ElseIf rs!gst_ari_nashi = 1 Then
                LM_GST = 1
            End If
        End If
        
        If Not IsNull(rs!no_siri_Produk) Then
            BARCODE_NOSIRI = rs!no_siri_Produk 'No. Siri Produk
        Else
            BARCODE_NOSIRI = "-" 'No. Siri Produk
        End If
        If Not IsNull(rs!Berat) Then
            BARCODE_BERAT = "Wg:" & Format(rs!Berat, "0.00g") 'Berat
            L_BERAT = Format(rs!Berat, "0.00g") 'Berat
        Else
            BARCODE_BERAT = "Wg:" & "-" 'Berat
        End If
        If Not IsNull(rs!Kod_Supplier) Then
            BARCODE_SUPPLIER = "S:" & rs!Kod_Supplier 'Kod Bagi Supplier
        Else
            BARCODE_SUPPLIER = "S:" & "-" 'Kod Bagi Supplier
        End If
        If Not IsNull(rs!UPAH) Then
            If IsNumeric(rs!UPAH) Then
                LM_UPAH = rs!UPAH
            Else
                LM_UPAH = rs!UPAH
            End If
            
            BARCODE_UPAH = "Wo:" & "1" & LM_UPAH & "2" 'Upah belian dari supplier
        Else
            BARCODE_UPAH = "Wo:" & "-" 'Upah belian dari supplier
        End If
        If Not IsNull(rs!kod_Purity) Then
            BARCODE_PURITY = "P:" & rs!kod_Purity 'Kod Purity Produk
            L_PURITY = rs!kod_Purity 'Kod Purity Produk
        Else
            BARCODE_PURITY = "P:" & "-" 'Kod Purity Produk
        End If
        If Not IsNull(rs!harga_item) Then
            
            If InStr(1, rs!harga_item, ".") <> 0 Then
            
                BARCODE_MODAL = "A" & Split(rs!harga_item, ".")(0) & "A"
                
            End If
            
        End If
        If Not IsNull(rs!dulang) Then
            BARCODE_DULANG = "T:" & rs!dulang 'dulang
            L_TRAY = rs!dulang
        Else
            BARCODE_DULANG = "T:" & "-" 'dulang
        End If
        If Not IsNull(rs!dimension_Panjang) Then
            BARCODE_Panjang = "L:" & rs!dimension_Panjang 'Dimension : Panjang
        Else
            BARCODE_Panjang = "L:" & "-" 'Dimension : Panjang
        End If
        If Not IsNull(rs!dimension_Lebar) Then
            BARCODE_Lebar = "W:" & rs!dimension_Lebar 'Dimension : Lebar
        Else
            BARCODE_Lebar = "W:" & "-" 'Dimension : Lebar
        End If
        If Not IsNull(rs!dimension_Saiz) Then
            BARCODE_Saiz = "Sz:" & rs!dimension_Saiz 'Dimension : Saiz
        Else
            BARCODE_Saiz = "Sz:" & "-" 'Dimension : Saiz
        End If
        If Not IsNull(rs!Upah_Jualan) Then
            BARCODE_UPAH2 = "Ws:" & rs!Upah_Jualan 'Upah Jualan (RM)
        Else
            BARCODE_UPAH2 = "Ws:" & "-" 'Upah Jualan (RM)
        End If
        If Not IsNull(rs!code_Supplier) Then
            BARCODE_HARGA = "RM" & Format(rs!code_Supplier, "#,##0.00") 'Harga Jualan Jualan (RM)
            L_HARGA = "RM" & Format(rs!code_Supplier, "#,##0.00")
        Else
            BARCODE_HARGA = "Pr:" & "-" 'Upah Jualan (RM)
        End If
        If Not IsNull(rs!riyal) Then
            BARCODE_RIYAL = rs!riyal & "R" 'Barcode Riyal (Berat Amah)
        Else
            BARCODE_RIYAL = "-" 'Barcode Riyal (Berat Amah)
        End If
        If Not IsNull(rs!code1) Then
            BARCODE_CODE1 = "C1:" & rs!code1
        End If
        If Not IsNull(rs!code2) Then
            BARCODE_CODE2 = "C2:" & rs!code2
        End If
        
        JENIS_BARCODE(1) = "BARCODE_SUPPLIER"
        JENIS_BARCODE(2) = "BARCODE_BERAT"
        JENIS_BARCODE(3) = "BARCODE_UPAH"
        JENIS_BARCODE(4) = "BARCODE_DULANG"
        JENIS_BARCODE(5) = "BARCODE_UPAH2"
        JENIS_BARCODE(6) = "BARCODE_Panjang"
        JENIS_BARCODE(7) = "BARCODE_Lebar"
        JENIS_BARCODE(8) = "BARCODE_Saiz"
        JENIS_BARCODE(9) = "BARCODE_PURITY"
        JENIS_BARCODE(10) = "BARCODE_CODE1"
        JENIS_BARCODE(11) = "BARCODE_CODE2"
        JENIS_BARCODE(12) = "BARCODE_RIYAL"
        
        VALUE_BARCODE(1) = BARCODE_SUPPLIER
        VALUE_BARCODE(2) = BARCODE_BERAT
        VALUE_BARCODE(3) = BARCODE_UPAH
        VALUE_BARCODE(4) = BARCODE_DULANG
        VALUE_BARCODE(5) = BARCODE_UPAH2
        VALUE_BARCODE(6) = BARCODE_Panjang
        VALUE_BARCODE(7) = BARCODE_Lebar
        VALUE_BARCODE(8) = BARCODE_Saiz
        VALUE_BARCODE(9) = BARCODE_PURITY
        VALUE_BARCODE(10) = BARCODE_CODE1
        VALUE_BARCODE(11) = BARCODE_CODE2
        VALUE_BARCODE(12) = BARCODE_RIYAL
        
        If L_JENIS_BARCODE = 0 Then
        
            Printer.FontName = "Code128"
            Printer.FontSize = 24
            Printer.CurrentX = 4
            Printer.CurrentY = 4
            Printer.Print BARCODE_NOSIRI
            'Printer.Print vbNullString
            
            Printer.FontSize = 7 '''Asal : Digunakan Kebanyakkan Kedai
            Printer.FontName = "Text"
            Printer.FontBold = True
            Printer.Print BARCODE_NOSIRI
            'Printer.Print vbNullString '''Asal : Digunakan Kebanyakkan Kedai
            
            Printer.FontName = "Andalus"
            Printer.FontSize = 6
            Printer.FontBold = True
            Printer.Print LM_HEAD
            Printer.FontBold = False
            Printer.FontName = "Text"
            Printer.FontSize = 6
            Printer.Print vbNullString
            
            Printer.FontSize = 6
            Printer.FontBold = True
            
            If rs!receiving_Status = 0 Or rs!receiving_Status = 2 Then
                
                x = 0
                
                For x = 1 To 12
                    
                    For Y = 1 To 12
                    
                        If G_NAMA_BARCODE(x) = JENIS_BARCODE(Y) Then
                        
                            LOC_(x) = VALUE_BARCODE(Y)
                            GoTo Skip_Carian3:
                        
                        End If
                        
                    Next Y
                    
Skip_Carian3:
                    
                Next x
                
                GoTo skip_this:
                
                '##########Masukkan Nilai Setiap Barcode##########
                Set rs2 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs2.Open "select * from tetapan_barcode where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic
                
                While rs2.EOF = False
                    x = x + 1
                    If Not IsNull(rs2!jenis) Then
                        For i = 1 To 12
                            If Not IsNull(rs2!Nama) Then
                                If rs2!Nama = JENIS_BARCODE(i) Then
                                    LOC_(x) = VALUE_BARCODE(i)
    
                                    GoTo Skip_Carian:
                                End If
                            End If
                        Next i
Skip_Carian:
                    End If
                    rs2.MoveNext
                Wend
                
                rs2.Close
                Set rs2 = Nothing
                
skip_this:
                
                If G_BAROCDE_LINE_1 = 1 Then
                
                    Printer.FontSize = Frm56x_LM_SIZE_1 'Saiz tulisan bagi barisan pertama
                    Printer.Print LOC_(1) & "  " & LOC_(2) & "  " & LOC_(10)
                    'Printer.Print BARCODE_BERAT & "  " & BARCODE_PURITY
                    
                Else
                
                    Printer.FontSize = Frm56x_LM_SIZE_1 'Saiz tulisan bagi barisan pertama
                    Printer.Print vbNullString
                    
                End If
                If G_BAROCDE_LINE_2 = 1 Then
                
                    Printer.FontSize = Frm56x_LM_SIZE_2 'Saiz tulisan bagi barisan kedua
                    Printer.Print LOC_(3) & "  " & LOC_(4) & "  " & LOC_(5)
                    'Printer.Print BARCODE_Panjang & "  " & BARCODE_Lebar & "  " & BARCODE_Saiz
                    
                Else
                
                    Printer.FontSize = Frm56x_LM_SIZE_2 'Saiz tulisan bagi barisan kedua
                    Printer.Print vbNullString
                    
                End If
                If G_BAROCDE_LINE_3 = 1 Then
                    
                    Printer.FontSize = Frm56x_LM_SIZE_3 'Saiz tulisan bagi barisan ketiga
                    Printer.Print LOC_(6) & "  " & LOC_(7) & "  " & LOC_(11)
                    'Printer.Print BARCODE_UPAH2 & "  " & BARCODE_CODE1
                    
                Else
                
                    Printer.FontSize = Frm56x_LM_SIZE_3 'Saiz tulisan bagi barisan ketiga
                    Printer.Print vbNullString
                    
                End If
                If G_BAROCDE_LINE_4 = 1 Then
                    
                    Printer.FontSize = Frm56x_LM_SIZE_4 'Saiz tulisan bagi barisan keempat
                    Printer.Print LOC_(8) & "  " & LOC_(9) & "  " & LOC_(12)
                    'Printer.Print BARCODE_SUPPLIER & "  " & BARCODE_DULANG
                    
                Else
                    
                    Printer.FontSize = Frm56x_LM_SIZE_4 'Saiz tulisan bagi barisan keempat
                    Printer.Print vbNullString
                    
                End If
    
            End If
            
            If rs!receiving_Status = 1 Or rs!receiving_Status = 3 Then
                
                'Printer.FontSize = 7
                'Printer.Print BARCODE_HARGA & " " & BARCODE_DULANG
                'Printer.Print BARCODE_PURITY & " " & BARCODE_CODE1 & " " & BARCODE_CODE2
                'Printer.Print BARCODE_Panjang & " " & BARCODE_Lebar & " " & BARCODE_Saiz
                'Printer.Print BARCODE_SUPPLIER & " " & BARCODE_MODAL
                
                Printer.FontSize = 7
                
                Printer.FontSize = Frm56x_LM_SIZE_1 'Saiz tulisan bagi barisan pertama
                Printer.Print BARCODE_HARGA
                
                Printer.FontSize = Frm56x_LM_SIZE_2 'Saiz tulisan bagi barisan kedua
                'Printer.Print BARCODE_PURITY & " " & BARCODE_CODE1 & " " & BARCODE_CODE2
                Printer.Print BARCODE_PURITY & " " & BARCODE_Saiz
                
                Printer.FontSize = Frm56x_LM_SIZE_3 'Saiz tulisan bagi barisan ketiga
                'Printer.Print BARCODE_Panjang & " " & BARCODE_Lebar & " " & BARCODE_Saiz
                Printer.Print BARCODE_SUPPLIER & " " & BARCODE_MODAL
                
                Printer.FontSize = Frm56x_LM_SIZE_4 'Saiz tulisan bagi barisan keempat
                Printer.Print BARCODE_DULANG
                
            End If
        
        ElseIf L_JENIS_BARCODE = 1 Then
        
            Printer.FontName = "Andalus"
            Printer.FontSize = 8
            'Printer.CurrentX = 500
            'Printer.CurrentY = 0
            Printer.FontBold = True
            Printer.Print LM_HEAD
            Printer.FontBold = False
            Printer.FontName = "Text"
            Printer.FontSize = 6
            Printer.Print vbNullString
            
            'Printer.FontName = "Code128"
            Printer.FontName = "Code39"
            Printer.FontSize = 24
            Printer.CurrentX = 10
            Printer.CurrentY = 220
            Printer.FontBold = True
            Printer.Print BARCODE_NOSIRI
            'Printer.Print vbNullString
            
            Printer.FontName = "Text"
        
            Printer.FontSize = 10
            Printer.CurrentX = 10
            Printer.CurrentY = 700
            Printer.FontBold = True
            If G_BARCODE_READABLE = "YES" Then
                Printer.Print BARCODE_NOSIRI
            Else
                Printer.Print vbNullString
            End If

            Printer.FontSize = 10
            Printer.CurrentX = G_TRAY_X
            Printer.CurrentY = 0
            Printer.FontBold = True
            Printer.Print L_TRAY
            Printer.FontBold = False
            
            Printer.FontSize = 6
            Printer.FontBold = True
            
            If rs!receiving_Status = 0 Or rs!receiving_Status = 2 Then
            
                x = 0
                
                For x = 1 To 12
                    
                    For Y = 1 To 12
                    
                        If G_NAMA_BARCODE(x) = JENIS_BARCODE(Y) Then
                        
                            LOC_(x) = VALUE_BARCODE(Y)
                            GoTo Skip_Carian4:
                        
                        End If
                        
                    Next Y
                    
Skip_Carian4:
                    
                Next x
                
                GoTo skip_this2:
                
                '##########Masukkan Nilai Setiap Barcode##########
                Set rs2 = New ADODB.Recordset
                If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
                rs2.Open "select * from tetapan_barcode where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic
                
                While rs2.EOF = False
                    x = x + 1
                    If Not IsNull(rs2!jenis) Then
                        For i = 1 To 12
                            If Not IsNull(rs2!Nama) Then
                                If rs2!Nama = JENIS_BARCODE(i) Then
                                    LOC_(x) = VALUE_BARCODE(i)
                                    GoTo Skip_Carian2:
                                End If
                            End If
                        Next i
Skip_Carian2:
                    End If
                    rs2.MoveNext
                Wend
                
                rs2.Close
                Set rs2 = Nothing
                
skip_this2:
                
                Printer.FontName = "Text"
        
                Printer.FontSize = 10
                Printer.CurrentX = 1400
                Printer.CurrentY = 300
                Printer.FontBold = True
                Printer.Print "Berat"
                
                Printer.FontSize = 8
                Printer.CurrentX = 1400
                Printer.CurrentY = 500
                Printer.FontBold = True
                Printer.Print L_BERAT
                
                Printer.FontSize = 8
                Printer.CurrentX = 1400
                Printer.CurrentY = 700
                'Printer.FontBold = False
                Printer.Print L_PURITY
                
                Printer.FontSize = 6
                Printer.Print vbNullString
            
                If G_BAROCDE_LINE_1 = 1 Then
                
                    Printer.FontSize = Frm56x_LM_SIZE_1 'Saiz tulisan bagi barisan pertama
                    Printer.Print LOC_(1) & "  " & LOC_(2) & "  " & LOC_(10)
                    'Printer.Print BARCODE_BERAT & "  " & BARCODE_PURITY
                    
                Else
                
                    Printer.FontSize = Frm56x_LM_SIZE_1 'Saiz tulisan bagi barisan pertama
                    Printer.Print vbNullString
                    
                End If
                If G_BAROCDE_LINE_2 = 1 Then
                
                    Printer.FontSize = Frm56x_LM_SIZE_2 'Saiz tulisan bagi barisan kedua
                    Printer.Print LOC_(3) & "  " & LOC_(4) & "  " & LOC_(5)
                    'Printer.Print BARCODE_Panjang & "  " & BARCODE_Lebar & "  " & BARCODE_Saiz
                    
                Else
                
                    Printer.FontSize = Frm56x_LM_SIZE_2 'Saiz tulisan bagi barisan kedua
                    Printer.Print vbNullString
                    
                End If
                If G_BAROCDE_LINE_3 = 1 Then
                    
                    Printer.FontSize = Frm56x_LM_SIZE_3 'Saiz tulisan bagi barisan ketiga
                    Printer.Print LOC_(6) & "  " & LOC_(7) & "  " & LOC_(11)
                    'Printer.Print BARCODE_UPAH2 & "  " & BARCODE_CODE1
                    
                Else
                
                    Printer.FontSize = Frm56x_LM_SIZE_3 'Saiz tulisan bagi barisan ketiga
                    Printer.Print vbNullString
                    
                End If
                If G_BAROCDE_LINE_4 = 1 Then
                    
                    Printer.FontSize = Frm56x_LM_SIZE_4 'Saiz tulisan bagi barisan keempat
                    Printer.Print LOC_(8) & "  " & LOC_(9) & "  " & LOC_(12)
                    'Printer.Print BARCODE_SUPPLIER & "  " & BARCODE_DULANG
                    
                Else
                    
                    Printer.FontSize = Frm56x_LM_SIZE_4 'Saiz tulisan bagi barisan keempat
                    Printer.Print vbNullString
                    
                End If
                
            End If
        
            If rs!receiving_Status = 1 Or rs!receiving_Status = 3 Then
    
                Printer.FontName = "Text"
        
                Printer.FontSize = 8
                Printer.CurrentX = 1250
                Printer.CurrentY = 300
                Printer.FontBold = True
                Printer.Print "Harga"
                
                Printer.FontSize = 7
                Printer.CurrentX = 1250
                Printer.CurrentY = 520
                Printer.FontBold = True
                Printer.Print L_HARGA
                
                Printer.FontSize = 7
                Printer.CurrentX = 1250
                Printer.CurrentY = 700
                'Printer.FontBold = False
                Printer.Print L_PURITY
                
                Printer.FontSize = 4
                Printer.Print vbNullString
                
                Printer.FontSize = 7
                
                Printer.FontSize = Frm56x_LM_SIZE_1 'Saiz tulisan bagi barisan pertama
                Printer.Print BARCODE_HARGA
                
                Printer.FontSize = Frm56x_LM_SIZE_2 'Saiz tulisan bagi barisan kedua
                'Printer.Print BARCODE_PURITY & " " & BARCODE_CODE1 & " " & BARCODE_CODE2
                Printer.Print BARCODE_PURITY & " " & BARCODE_Saiz
                
                Printer.FontSize = Frm56x_LM_SIZE_3 'Saiz tulisan bagi barisan ketiga
                'Printer.Print BARCODE_Panjang & " " & BARCODE_Lebar & " " & BARCODE_Saiz
                Printer.Print BARCODE_SUPPLIER & " " & BARCODE_MODAL
                
                Printer.FontSize = Frm56x_LM_SIZE_4 'Saiz tulisan bagi barisan keempat
                Printer.Print BARCODE_DULANG
                
            End If
            
        End If
        
        Printer.FontSize = 7
        Printer.FontSize = G_Frm56_LM_SIZE_NAMA_KEDAI
        'Printer.Print vbNullString
        
        If LM_GST = 1 Then
            Printer.Print LM_NAMA_KEDAI
        ElseIf LM_GST = 0 Then
            Printer.Print LM_NAMA_KEDAI & "*"
        End If
        
        'Printer.Print LM_NAMA_KEDAI
        Printer.FontBold = False
        Printer.FontName = "Text"
        Printer.FontSize = 7 '''Asal : Digunakan Kebanyakkan Kedai
        
        Printer.EndDoc
        
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
Else
    'MsgBox "Barcode Label Printer [ZDesigner GT800 (EPL)] Tidak Dijumpai.", vbCritical, "Sila Install Software Untuk Printer"
    MsgBox "Barcode Label Printer [" & G_PRINTER_BARCODE & "] Tidak Dijumpai.", vbCritical, "Sila Install Software Untuk Printer"
End If
End Sub

Sub Print_All_Barcode2()
'On Error GoTo logging:
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset

Dim JENIS_BARCODE(17)
Dim LOC_(12)
Dim VALUE_BARCODE(17)
Dim LM_NAMA_KEDAI As String
Dim LM_BERAT_STOK As Double
Dim LM_UPAH_JUALAN As Double

Frm56_LM_SIZE_0 = 6 'Saiz tulisan bagi barcode
Frm56_LM_SIZE_1 = 6 'Saiz tulisan bagi barisan pertama
Frm56_LM_SIZE_2 = 6 'Saiz tulisan bagi barisan kedua
Frm56_LM_SIZE_3 = 6 'Saiz tulisan bagi barisan ketiga
Frm56_LM_SIZE_4 = 6 'Saiz tulisan bagi barisan keempat
Frm56_LM_SIZE_5 = 6 'Saiz tulisan bagi nama kedai

PRINTER_FOUND = 0 '0 : Barcode Label Printer Not Found , 1 : Barcode Label Printer Found
For Each oPrn In Printers
    'If oPrn.DeviceName = "ZDesigner GT800 (EPL)" Then
    If oPrn.DeviceName = G_PRINTER_BARCODE Then
        Set Printer = oPrn
        PRINTER_FOUND = 1 '0 : Barcode Label Printer Not Found , 1 : Barcode Label Printer Found
        Exit For
    End If
Next

LM_HEAD = "Sankyu System"
LM_NAMA_KEDAI = G_LM_NAMA_KEDAI
LM_HEAD = G_LM_HEAD

L_JENIS_BARCODE = 0
L_JENIS_BARCODE = G_L_JENIS_BARCODE

If PRINTER_FOUND = 1 Then '0 : Barcode Label Printer Not Found , 1 : Barcode Label Printer Found
LM_CONN = 1
re_conn_1:
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from data_database where " & G_FIELD & "='" & GM_No_RUJUKAN_BELIAN & "' AND StatusItem <> 0", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
    
        Call ClearPrinterMemory
        
        LM_UPAH = 0
        L_BERAT = vbNullString
        L_TRAY = vbNullString
        L_PURITY = purity
        L_HARGA = 0
        LM_BERAT_STOK = 0
        LM_UPAH_JUALAN = 0
        LM_GST = 0
        
        If Not IsNull(rs!gst_ari_nashi) Then
            If rs!gst_ari_nashi = 0 Then
                LM_GST = 0
            ElseIf rs!gst_ari_nashi = 1 Then
                LM_GST = 1
            End If
        End If
        
        If Not IsNull(rs!no_siri_Produk) Then
            BARCODE_NOSIRI = rs!no_siri_Produk 'No. Siri Produk
        Else
            BARCODE_NOSIRI = "-" 'No. Siri Produk
        End If
        If Not IsNull(rs!Berat) Then
            BARCODE_BERAT = "Wg:" & Format(rs!beza_berat, "0.00g") 'Berat
            L_BERAT = Format(rs!beza_berat, "0.00g") 'Berat
            If IsNumeric(rs!Berat) Then LM_BERAT_STOK = rs!Berat
        Else
            BARCODE_BERAT = "Wg:" & "-" 'Berat
        End If
        If Not IsNull(rs!Kod_Supplier) Then
            BARCODE_SUPPLIER = "S:" & rs!Kod_Supplier 'Kod Bagi Supplier
        Else
            BARCODE_SUPPLIER = "S:" & "-" 'Kod Bagi Supplier
        End If
        If Not IsNull(rs!UPAH) Then
            If IsNumeric(rs!UPAH) Then
                LM_UPAH = rs!UPAH
            Else
                LM_UPAH = rs!UPAH
            End If
            
            BARCODE_UPAH = "Wo:" & "1" & LM_UPAH & "2" 'Upah belian dari supplier
        Else
            BARCODE_UPAH = "Wo:" & "-" 'Upah belian dari supplier
        End If
        If Not IsNull(rs!harga_item) Then
            If InStr(1, rs!harga_item, ".") <> 0 Then
                BARCODE_MODAL = "A" & Split(rs!harga_item, ".")(0) & "A"
            Else
                BARCODE_MODAL = "A" & rs!harga_item & "A"
            End If
        End If
        If Not IsNull(rs!code1) Then
            If G_JENIS_KOD_UPAH = 1 Then BARCODE_MODAL = rs!code1
        End If
        If Not IsNull(rs!kod_Purity) Then
            BARCODE_PURITY = "P:" & rs!kod_Purity 'Kod Purity Produk
            L_PURITY = rs!kod_Purity 'Kod Purity Produk
        Else
            BARCODE_PURITY = "P:" & "-" 'Kod Purity Produk
        End If
        If Not IsNull(rs!dulang) Then
            BARCODE_DULANG = "T:" & rs!dulang 'dulang
            L_TRAY = rs!dulang
        Else
            BARCODE_DULANG = "T:" & "-" 'dulang
        End If
        If Not IsNull(rs!dimension_Panjang) Then
            BARCODE_Panjang = "L:" & rs!dimension_Panjang 'Dimension : Panjang
        Else
            BARCODE_Panjang = "L:" & "-" 'Dimension : Panjang
        End If
        If Not IsNull(rs!dimension_Lebar) Then
            BARCODE_Lebar = "W:" & rs!dimension_Lebar 'Dimension : Lebar
        Else
            BARCODE_Lebar = "W:" & "-" 'Dimension : Lebar
        End If
        If Not IsNull(rs!dimension_Saiz) Then
            BARCODE_Saiz = "Sz:" & rs!dimension_Saiz 'Dimension : Saiz
        Else
            BARCODE_Saiz = "Sz:" & "-" 'Dimension : Saiz
        End If
'Paparan upah jualan pada tagging
        If Not IsNull(rs!Upah_Jualan) Then
            BARCODE_UPAH2 = "Ws:" & rs!Upah_Jualan 'Upah Jualan (RM)
        Else
            BARCODE_UPAH2 = "Ws:" & "-" 'Upah Jualan (RM)
        End If
        If Not IsNull(rs!code_Supplier) Then
            BARCODE_HARGA = "RM" & Format(rs!code_Supplier, "#,##0.00") 'Harga Jualan Jualan (RM)
            L_HARGA = "RM" & Format(rs!code_Supplier, "#,##0.00")
        Else
            BARCODE_HARGA = "Pr:" & "-" 'Upah Jualan (RM)
        End If
        If Not IsNull(rs!riyal) Then
            BARCODE_RIYAL = rs!riyal & "R" 'Barcode Riyal (Berat Amah)
        Else
            BARCODE_RIYAL = "-" 'Barcode Riyal (Berat Amah)
        End If
        If Not IsNull(rs!code1) Then
            BARCODE_CODE1 = "C1:" & rs!code1
        End If
        If Not IsNull(rs!code2) Then
            BARCODE_CODE2 = "C2:" & rs!code2
        End If
        
        JENIS_BARCODE(1) = "BARCODE_SUPPLIER"
        JENIS_BARCODE(2) = "BARCODE_BERAT"
        JENIS_BARCODE(3) = "BARCODE_UPAH"
        JENIS_BARCODE(4) = "BARCODE_DULANG"
        JENIS_BARCODE(5) = "BARCODE_UPAH2"
        JENIS_BARCODE(6) = "BARCODE_Panjang"
        JENIS_BARCODE(7) = "BARCODE_Lebar"
        JENIS_BARCODE(8) = "BARCODE_Saiz"
        JENIS_BARCODE(9) = "BARCODE_PURITY"
        JENIS_BARCODE(10) = "BARCODE_CODE1"
        JENIS_BARCODE(11) = "BARCODE_CODE2"
        JENIS_BARCODE(12) = "BARCODE_RIYAL"
        JENIS_BARCODE(13) = "BARCODE_BARCODE"
        JENIS_BARCODE(14) = "BARCODE_HARGA"
        JENIS_BARCODE(15) = "BARCODE_MODAL"
        JENIS_BARCODE(16) = "BARCODE_DIAMOND"
        JENIS_BARCODE(17) = "BARCODE_DESIGN"

        VALUE_BARCODE(1) = BARCODE_SUPPLIER
        VALUE_BARCODE(2) = BARCODE_BERAT
        VALUE_BARCODE(3) = BARCODE_UPAH
        VALUE_BARCODE(4) = BARCODE_DULANG
        VALUE_BARCODE(5) = BARCODE_UPAH2
        VALUE_BARCODE(6) = BARCODE_Panjang
        VALUE_BARCODE(7) = BARCODE_Lebar
        VALUE_BARCODE(8) = BARCODE_Saiz
        VALUE_BARCODE(9) = BARCODE_PURITY
        VALUE_BARCODE(10) = BARCODE_CODE1
        VALUE_BARCODE(11) = BARCODE_CODE2
        VALUE_BARCODE(12) = BARCODE_RIYAL
        VALUE_BARCODE(13) = BARCODE_NOSIRI
        VALUE_BARCODE(14) = BARCODE_HARGA
        VALUE_BARCODE(15) = BARCODE_MODAL
        VALUE_BARCODE(16) = BARCODE_DIAMOND
        VALUE_BARCODE(17) = BARCODE_DESIGN
        
        If L_JENIS_BARCODE = 0 Then
'TYPE A
            Printer.FontName = Split(G_SKU_FONT(0), ",")(0)
            Printer.FontSize = Split(G_SKU_SIZE(0), ",")(0)
            If Split(G_SKU_POS_X(0), ",")(0) <> "/" Then Printer.CurrentX = Split(G_SKU_POS_X(0), ",")(0)
            If Split(G_SKU_POS_Y(0), ",")(0) <> "/" Then Printer.CurrentY = Split(G_SKU_POS_Y(0), ",")(0)
            Printer.FontBold = Split(G_SKU_BOLD(0), ",")(0)
            Printer.FontItalic = Split(G_SKU_ITALIC(0), ",")(0)
            Printer.Print BARCODE_NOSIRI
            
            Printer.FontName = Split(G_SKU_FONT(0), ",")(1)
            Printer.FontSize = Split(G_SKU_SIZE(0), ",")(1)
            If Split(G_SKU_POS_X(0), ",")(1) <> "/" Then Printer.CurrentX = Split(G_SKU_POS_X(0), ",")(1)
            If Split(G_SKU_POS_Y(0), ",")(1) <> "/" Then Printer.CurrentY = Split(G_SKU_POS_Y(0), ",")(1)
            Printer.FontBold = Split(G_SKU_BOLD(0), ",")(1)
            Printer.FontItalic = Split(G_SKU_ITALIC(0), ",")(1)
            Printer.Print BARCODE_NOSIRI

            Printer.FontName = Split(G_SKU_FONT(0), ",")(7)
            Printer.FontSize = Split(G_SKU_SIZE(0), ",")(7)
            If Split(G_SKU_POS_X(0), ",")(7) <> "/" Then Printer.CurrentX = Split(G_SKU_POS_X(0), ",")(7)
            If Split(G_SKU_POS_Y(0), ",")(7) <> "/" Then Printer.CurrentY = Split(G_SKU_POS_Y(0), ",")(7)
            Printer.FontBold = Split(G_SKU_BOLD(0), ",")(7)
            Printer.FontItalic = Split(G_SKU_ITALIC(0), ",")(7)
            Printer.Print LM_HEAD
            
            Printer.FontBold = False
            Printer.FontName = "Text"
            Printer.FontSize = 6
            Printer.Print vbNullString
            
            If rs!receiving_Status = 0 Or rs!receiving_Status = 2 Then
                LM_JENIS = 0
            ElseIf rs!receiving_Status = 1 Or rs!receiving_Status = 3 Then
                LM_JENIS = 1
            End If
            
            x = 0
            
            For x = 0 To 12
                For Y = 1 To 15
                    If Split(G_SKU_DATA(LM_JENIS), ",")(x) = JENIS_BARCODE(Y) Then
                        LOC_(x) = VALUE_BARCODE(Y)
                        GoTo Skip_Carian4:
                    End If
                Next Y
Skip_Carian4:
            Next x
            
            Printer.FontBold = False
            
            If Split(G_SKU_LINE(0), ",")(0) = 1 Then

                Printer.FontName = Split(G_SKU_FONT(0), ",")(2)
                Printer.FontSize = Split(G_SKU_SIZE(0), ",")(2)
                If Split(G_SKU_POS_X(0), ",")(2) <> "/" Then Printer.CurrentX = Split(G_SKU_POS_X(0), ",")(2)
                If Split(G_SKU_POS_Y(0), ",")(2) <> "/" Then Printer.CurrentY = Split(G_SKU_POS_Y(0), ",")(2)
                Printer.FontBold = Split(G_SKU_BOLD(0), ",")(2)
                Printer.FontItalic = Split(G_SKU_ITALIC(0), ",")(2)
                Printer.Print LOC_(0) & "  " & LOC_(1) & "  " & LOC_(2)
            Else
                Printer.FontSize = Split(G_SKU_SIZE(0), ",")(2) 'Saiz tulisan bagi barisan pertama
                Printer.Print vbNullString
            End If
            If Split(G_SKU_LINE(0), ",")(1) = 1 Then
                Printer.FontName = Split(G_SKU_FONT(0), ",")(3)
                Printer.FontSize = Split(G_SKU_SIZE(0), ",")(3)
                If Split(G_SKU_POS_X(0), ",")(3) <> "/" Then Printer.CurrentX = Split(G_SKU_POS_X(0), ",")(3)
                If Split(G_SKU_POS_Y(0), ",")(3) <> "/" Then Printer.CurrentY = Split(G_SKU_POS_Y(0), ",")(3)
                Printer.FontBold = Split(G_SKU_BOLD(0), ",")(3)
                Printer.FontItalic = Split(G_SKU_ITALIC(0), ",")(3)
                Printer.Print LOC_(3) & "  " & LOC_(4) & "  " & LOC_(5)
            Else
                Printer.FontSize = Split(G_SKU_SIZE(0), ",")(3) 'Saiz tulisan bagi barisan kedua
                Printer.Print vbNullString
            End If
            If Split(G_SKU_LINE(0), ",")(2) = 1 Then
                Printer.FontName = Split(G_SKU_FONT(0), ",")(4)
                Printer.FontSize = Split(G_SKU_SIZE(0), ",")(4)
                If Split(G_SKU_POS_X(0), ",")(4) <> "/" Then Printer.CurrentX = Split(G_SKU_POS_X(0), ",")(4)
                If Split(G_SKU_POS_Y(0), ",")(4) <> "/" Then Printer.CurrentY = Split(G_SKU_POS_Y(0), ",")(4)
                Printer.FontBold = Split(G_SKU_BOLD(0), ",")(4)
                Printer.FontItalic = Split(G_SKU_ITALIC(0), ",")(4)
                Printer.Print LOC_(6) & "  " & LOC_(7) & "  " & LOC_(8)
            Else
                Printer.FontSize = Split(G_SKU_SIZE(0), ",")(4) 'Saiz tulisan bagi barisan ketiga
                Printer.Print vbNullString
                
            End If
            If Split(G_SKU_LINE(0), ",")(3) = 1 Then
                Printer.FontName = Split(G_SKU_FONT(0), ",")(5)
                Printer.FontSize = Split(G_SKU_SIZE(0), ",")(5)
                If Split(G_SKU_POS_X(0), ",")(5) <> "/" Then Printer.CurrentX = Split(G_SKU_POS_X(0), ",")(5)
                If Split(G_SKU_POS_Y(0), ",")(5) <> "/" Then Printer.CurrentY = Split(G_SKU_POS_Y(0), ",")(5)
                Printer.FontBold = Split(G_SKU_BOLD(0), ",")(5)
                Printer.FontItalic = Split(G_SKU_ITALIC(0), ",")(5)
                Printer.Print LOC_(9) & "  " & LOC_(10) & "  " & LOC_(11)
            Else
                Printer.FontSize = Split(G_SKU_SIZE(0), ",")(5) 'Saiz tulisan bagi barisan keempat
                Printer.Print vbNullString
            End If

            Printer.FontName = Split(G_SKU_FONT(0), ",")(6)
            Printer.FontSize = Split(G_SKU_SIZE(0), ",")(6)
            If Split(G_SKU_POS_X(0), ",")(6) <> "/" Then Printer.CurrentX = Split(G_SKU_POS_X(0), ",")(6)
            If Split(G_SKU_POS_Y(0), ",")(6) <> "/" Then Printer.CurrentY = Split(G_SKU_POS_Y(0), ",")(6)
            Printer.FontBold = Split(G_SKU_BOLD(0), ",")(6)
            Printer.FontItalic = Split(G_SKU_ITALIC(0), ",")(6)
                        
            If L_JENIS_BARCODE = 0 Then LM_NAMA_KEDAI = G_FOOTER_A
            
            If LM_GST = 1 Then
                Printer.Print LM_NAMA_KEDAI
            ElseIf LM_GST = 0 Then
                Printer.Print LM_NAMA_KEDAI & "*"
            End If
        
        ElseIf L_JENIS_BARCODE = 1 Or L_JENIS_BARCODE = 2 Then
'TYPE B
            Printer.FontName = Split(G_SKU_FONT(L_JENIS_BARCODE), ",")(7)
            Printer.FontSize = Split(G_SKU_SIZE(L_JENIS_BARCODE), ",")(7)
            If Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(7) <> "/" Then Printer.CurrentX = Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(7)
            If Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(7) <> "/" Then Printer.CurrentY = Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(7)
            Printer.FontBold = Split(G_SKU_BOLD(L_JENIS_BARCODE), ",")(7)
            Printer.FontItalic = Split(G_SKU_ITALIC(L_JENIS_BARCODE), ",")(7)
            Printer.Print LM_HEAD
            
            Printer.FontBold = False
            Printer.FontName = "Text"
            Printer.FontSize = 6
            Printer.Print vbNullString

            Printer.FontName = Split(G_SKU_FONT(L_JENIS_BARCODE), ",")(0)
            Printer.FontSize = Split(G_SKU_SIZE(L_JENIS_BARCODE), ",")(0)
            If Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(0) <> "/" Then Printer.CurrentX = Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(0)
            If Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(0) <> "/" Then Printer.CurrentY = Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(0)
            Printer.FontBold = Split(G_SKU_BOLD(L_JENIS_BARCODE), ",")(0)
            Printer.FontItalic = Split(G_SKU_ITALIC(L_JENIS_BARCODE), ",")(0)
            Printer.Print BARCODE_NOSIRI

            Printer.FontName = Split(G_SKU_FONT(L_JENIS_BARCODE), ",")(1)
            Printer.FontSize = Split(G_SKU_SIZE(L_JENIS_BARCODE), ",")(1)
            If Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(1) <> "/" Then Printer.CurrentX = Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(1)
            If Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(1) <> "/" Then Printer.CurrentY = Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(1)
            Printer.FontBold = Split(G_SKU_BOLD(L_JENIS_BARCODE), ",")(1)
            Printer.FontItalic = Split(G_SKU_ITALIC(L_JENIS_BARCODE), ",")(1)
            If G_BARCODE_READABLE = "YES" Then
                Printer.Print BARCODE_NOSIRI
            Else
                Printer.Print vbNullString
            End If
            'done/////
            Printer.FontName = Split(G_SKU_FONT(L_JENIS_BARCODE), ",")(8)
            Printer.FontSize = Split(G_SKU_SIZE(L_JENIS_BARCODE), ",")(8)
            If Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(8) <> "/" Then Printer.CurrentX = Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(8)
            If Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(8) <> "/" Then Printer.CurrentY = Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(8)
            Printer.FontBold = Split(G_SKU_BOLD(L_JENIS_BARCODE), ",")(8)
            Printer.FontItalic = Split(G_SKU_ITALIC(L_JENIS_BARCODE), ",")(8)
            Printer.Print L_TRAY
            
            If rs!receiving_Status = 0 Or rs!receiving_Status = 2 Then
                LM_JENIS = 0
                Printer.FontName = Split(G_SKU_FONT(L_JENIS_BARCODE), ",")(9)
                Printer.FontSize = Split(G_SKU_SIZE(L_JENIS_BARCODE), ",")(9)
                If Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(9) <> "/" Then Printer.CurrentX = Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(9)
                If Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(9) <> "/" Then Printer.CurrentY = Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(9)
                Printer.FontBold = Split(G_SKU_BOLD(L_JENIS_BARCODE), ",")(9)
                Printer.FontItalic = Split(G_SKU_ITALIC(L_JENIS_BARCODE), ",")(9)
                Printer.Print "Berat"
                
                Printer.FontName = Split(G_SKU_FONT(L_JENIS_BARCODE), ",")(10)
                Printer.FontSize = Split(G_SKU_SIZE(L_JENIS_BARCODE), ",")(10)
                If Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(10) <> "/" Then Printer.CurrentX = Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(10)
                If Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(10) <> "/" Then Printer.CurrentY = Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(10)
                Printer.FontBold = Split(G_SKU_BOLD(L_JENIS_BARCODE), ",")(10)
                Printer.FontItalic = Split(G_SKU_ITALIC(L_JENIS_BARCODE), ",")(10)
                Printer.Print L_BERAT
                
                Printer.FontName = Split(G_SKU_FONT(L_JENIS_BARCODE), ",")(11)
                Printer.FontSize = Split(G_SKU_SIZE(L_JENIS_BARCODE), ",")(11)
                If Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(11) <> "/" Then Printer.CurrentX = Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(11)
                If Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(11) <> "/" Then Printer.CurrentY = Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(11)
                Printer.FontBold = Split(G_SKU_BOLD(L_JENIS_BARCODE), ",")(11)
                Printer.FontItalic = Split(G_SKU_ITALIC(L_JENIS_BARCODE), ",")(11)
                Printer.Print L_PURITY
                
                Printer.FontSize = 4
                Printer.FontBold = False
                Printer.Print vbNullString

            ElseIf rs!receiving_Status = 1 Or rs!receiving_Status = 3 Then
                LM_JENIS = 1
                Printer.FontName = Split(G_SKU_FONT(L_JENIS_BARCODE), ",")(9)
                Printer.FontSize = Split(G_SKU_SIZE(L_JENIS_BARCODE), ",")(9)
                If Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(9) <> "/" Then Printer.CurrentX = Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(9)
                If Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(9) <> "/" Then Printer.CurrentY = Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(9)
                Printer.FontBold = Split(G_SKU_BOLD(L_JENIS_BARCODE), ",")(9)
                Printer.FontItalic = Split(G_SKU_ITALIC(L_JENIS_BARCODE), ",")(9)
                Printer.Print "Harga"
                
                Printer.FontName = Split(G_SKU_FONT(L_JENIS_BARCODE), ",")(10)
                Printer.FontSize = Split(G_SKU_SIZE(L_JENIS_BARCODE), ",")(10)
                If Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(10) <> "/" Then Printer.CurrentX = Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(10)
                If Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(10) <> "/" Then Printer.CurrentY = Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(10)
                Printer.FontBold = Split(G_SKU_BOLD(L_JENIS_BARCODE), ",")(10)
                Printer.FontItalic = Split(G_SKU_ITALIC(L_JENIS_BARCODE), ",")(10)
                Printer.Print L_HARGA
                
                Printer.FontName = Split(G_SKU_FONT(L_JENIS_BARCODE), ",")(11)
                Printer.FontSize = Split(G_SKU_SIZE(L_JENIS_BARCODE), ",")(11)
                If Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(11) <> "/" Then Printer.CurrentX = Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(11)
                If Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(11) <> "/" Then Printer.CurrentY = Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(11)
                Printer.FontBold = Split(G_SKU_BOLD(L_JENIS_BARCODE), ",")(11)
                Printer.FontItalic = Split(G_SKU_ITALIC(L_JENIS_BARCODE), ",")(11)
                Printer.Print L_PURITY
                
                Printer.FontSize = 4
                Printer.FontBold = False
                Printer.Print vbNullString
            End If
            
            x = 0
            
            For x = 0 To 12
                For Y = 1 To 17
                    If Split(G_SKU_DATA(LM_JENIS), ",")(x) = JENIS_BARCODE(Y) Then
                        LOC_(x) = VALUE_BARCODE(Y)
                        GoTo Skip_Carian5:
                    End If
                Next Y
Skip_Carian5:
            Next x
                
            Printer.FontBold = False
            
            If Split(G_SKU_LINE(L_JENIS_BARCODE), ",")(0) = 1 Then

                Printer.FontName = Split(G_SKU_FONT(L_JENIS_BARCODE), ",")(2)
                Printer.FontSize = Split(G_SKU_SIZE(L_JENIS_BARCODE), ",")(2)
                If Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(2) <> "/" Then Printer.CurrentX = Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(2)
                If Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(2) <> "/" Then Printer.CurrentY = Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(2)
                Printer.FontBold = Split(G_SKU_BOLD(L_JENIS_BARCODE), ",")(2)
                Printer.FontItalic = Split(G_SKU_ITALIC(L_JENIS_BARCODE), ",")(2)
                Printer.Print LOC_(0) & "  " & LOC_(1) & "  " & LOC_(2)
            Else
                Printer.FontSize = Split(G_SKU_SIZE(L_JENIS_BARCODE), ",")(2) 'Saiz tulisan bagi barisan pertama
                Printer.Print vbNullString
            End If
            If Split(G_SKU_LINE(L_JENIS_BARCODE), ",")(1) = 1 Then
                Printer.FontName = Split(G_SKU_FONT(L_JENIS_BARCODE), ",")(3)
                Printer.FontSize = Split(G_SKU_SIZE(L_JENIS_BARCODE), ",")(3)
                If Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(3) <> "/" Then Printer.CurrentX = Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(3)
                If Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(3) <> "/" Then Printer.CurrentY = Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(3)
                Printer.FontBold = Split(G_SKU_BOLD(L_JENIS_BARCODE), ",")(3)
                Printer.FontItalic = Split(G_SKU_ITALIC(L_JENIS_BARCODE), ",")(3)
                Printer.Print LOC_(3) & "  " & LOC_(4) & "  " & LOC_(5)
            Else
                Printer.FontSize = Split(G_SKU_SIZE(L_JENIS_BARCODE), ",")(3) 'Saiz tulisan bagi barisan kedua
                Printer.Print vbNullString
            End If
            If Split(G_SKU_LINE(L_JENIS_BARCODE), ",")(2) = 1 Then
                Printer.FontName = Split(G_SKU_FONT(L_JENIS_BARCODE), ",")(4)
                Printer.FontSize = Split(G_SKU_SIZE(L_JENIS_BARCODE), ",")(4)
                If Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(4) <> "/" Then Printer.CurrentX = Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(4)
                If Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(4) <> "/" Then Printer.CurrentY = Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(4)
                Printer.FontBold = Split(G_SKU_BOLD(L_JENIS_BARCODE), ",")(4)
                Printer.FontItalic = Split(G_SKU_ITALIC(L_JENIS_BARCODE), ",")(4)
                Printer.Print LOC_(6) & "  " & LOC_(7) & "  " & LOC_(8)
            Else
                Printer.FontSize = Split(G_SKU_SIZE(L_JENIS_BARCODE), ",")(4) 'Saiz tulisan bagi barisan ketiga
                Printer.Print vbNullString
                
            End If
            If Split(G_SKU_LINE(L_JENIS_BARCODE), ",")(3) = 1 Then
                Printer.FontName = Split(G_SKU_FONT(L_JENIS_BARCODE), ",")(5)
                Printer.FontSize = Split(G_SKU_SIZE(L_JENIS_BARCODE), ",")(5)
                If Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(5) <> "/" Then Printer.CurrentX = Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(5)
                If Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(5) <> "/" Then Printer.CurrentY = Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(5)
                Printer.FontBold = Split(G_SKU_BOLD(L_JENIS_BARCODE), ",")(5)
                Printer.FontItalic = Split(G_SKU_ITALIC(L_JENIS_BARCODE), ",")(5)
                Printer.Print LOC_(9) & "  " & LOC_(10) & "  " & LOC_(11)
            Else
                Printer.FontSize = Split(G_SKU_SIZE(L_JENIS_BARCODE), ",")(5) 'Saiz tulisan bagi barisan keempat
                Printer.Print vbNullString
            End If
            
            Printer.FontName = Split(G_SKU_FONT(L_JENIS_BARCODE), ",")(6)
            Printer.FontSize = Split(G_SKU_SIZE(L_JENIS_BARCODE), ",")(6)
            If Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(6) <> "/" Then Printer.CurrentX = Split(G_SKU_POS_X(L_JENIS_BARCODE), ",")(6)
            If Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(6) <> "/" Then Printer.CurrentY = Split(G_SKU_POS_Y(L_JENIS_BARCODE), ",")(6)
            Printer.FontBold = Split(G_SKU_BOLD(L_JENIS_BARCODE), ",")(6)
            Printer.FontItalic = Split(G_SKU_ITALIC(L_JENIS_BARCODE), ",")(6)
                        
            If L_JENIS_BARCODE = 0 Then LM_NAMA_KEDAI = G_FOOTER_A
            
            If LM_GST = 1 Then
                Printer.Print LM_NAMA_KEDAI
            ElseIf LM_GST = 0 Then
                Printer.Print LM_NAMA_KEDAI & "*"
            End If
        
        End If
        
        'Printer.Print LM_NAMA_KEDAI
        Printer.FontBold = False
        Printer.FontName = "Text"
        Printer.FontSize = 7 '''Asal : Digunakan Kebanyakkan Kedai
        
        Printer.EndDoc
        
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
Else
    If G_PRINTER_BARCODE = "NO SDK" Then
        MsgBox "Terminal ini tiada SDK untuk cetak barcode. Sila hubungi pihak Sankyu System bagi pembelian SDK.", vbCritical, "SDK"
    Else
        MsgBox "Barcode Label Printer [" & G_PRINTER_BARCODE & "] Tidak Dijumpai.", vbCritical, "Sila Install Software Untuk Printer"
    End If
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " DrawBar : Print_All_Barcode2" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If LM_ERR_NO = "3704" Or LM_ERR_NO = "-2147467259" Or LM_ERR_NO = "-2147217887" Then
    Call Main
    
    If LM_CONN = 1 Then
        Resume re_conn_1:
    End If
Else
    Resume Next
End If
End Sub
Sub cetak_barcode_gb()
'On Error Resume Next
'Hanya untuk cetak HANYA SATU barcode bagi item ini (bukan mengikut batch kemasukkan data ke dalam sistem)
Dim oPrn As Printer
Dim LM_NAMA_KEDAI As String

PRINTER_FOUND = 0 '0 : Barcode Label Printer Not Found , 1 : Barcode Label Printer Found

'Frm56x_LM_SIZE_1 = 6 'Saiz tulisan bagi barisan pertama
'Frm56x_LM_SIZE_2 = 6 'Saiz tulisan bagi barisan kedua
'Frm56x_LM_SIZE_3 = 6 'Saiz tulisan bagi barisan ketiga
'Frm56x_LM_SIZE_4 = 6 'Saiz tulisan bagi barisan keempat

For Each oPrn In Printers
    'If oPrn.DeviceName = "ZDesigner GT800 (EPL)" Then
    If oPrn.DeviceName = G_PRINTER_BARCODE Then
        Set Printer = oPrn
        PRINTER_FOUND = 1 '0 : Barcode Label Printer Not Found , 1 : Barcode Label Printer Found
        Exit For
    End If
Next

LM_HEAD = "Sankyu System"

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai_3) Then LM_NAMA_KEDAI = rs!nama_kedai_3
    If Not IsNull(rs!bar_header) Then LM_HEAD = rs!bar_header
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End
        
If PRINTER_FOUND = 1 Then '0 : Barcode Label Printer Not Found , 1 : Barcode Label Printer Found
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from Data_Database where no_siri_Produk='" & GM_No_RUJUKAN_BELIAN & "' AND StatusItem <> 0", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        LM_NO_SIRI = vbNullString
        LM_BERAT = vbNullString
        LM_PURITY = vbNullString
        LM_CERT = vbNullString
        LM_DULANG = vbNullString
        
        LM_GST = 0
        
        If Not IsNull(rs!gst_ari_nashi) Then
            If rs!gst_ari_nashi = 0 Then
                LM_GST = 0
            ElseIf rs!gst_ari_nashi = 1 Then
                LM_GST = 1
            End If
        End If
        
        If Not IsNull(rs!no_siri_Produk) Then LM_NO_SIRI = rs!no_siri_Produk 'No. Siri Produk
        If Not IsNull(rs!Berat) Then LM_BERAT = Format(rs!Berat, "#,##0.00 g") 'Berat (g)
        If Not IsNull(rs!kod_Purity) Then LM_PURITY = rs!kod_Purity 'Purity
        If Not IsNull(rs!dulang) Then LM_DULANG = rs!dulang 'Dulang
        If Not IsNull(rs!no_cert) Then 'No. Cert
            LM_CERT = rs!no_cert
        Else
            LM_CERT = "-"
        End If
        
        Printer.FontName = "Andalus"
        Printer.FontSize = 8
        'Printer.CurrentX = 500
        'Printer.CurrentY = 0
        Printer.FontBold = True
        Printer.Print LM_HEAD
        Printer.FontBold = False
        Printer.FontName = "Text"
        Printer.FontSize = 6
        Printer.Print vbNullString
        
        'Printer.FontName = "Code128"
        Printer.FontName = "Code39"
        Printer.FontSize = 24
        Printer.CurrentX = 10
        Printer.CurrentY = 220
        Printer.FontBold = True
        Printer.Print LM_NO_SIRI
        'Printer.Print vbNullString
        
        Printer.FontName = "Text"
        
        Printer.FontSize = 10
        Printer.CurrentX = 10
        Printer.CurrentY = 700
        Printer.FontBold = True
        'Printer.Print LM_NO_SIRI
        If G_BARCODE_READABLE = "YES" Then
            Printer.Print LM_NO_SIRI
        Else
            Printer.Print vbNullString
        End If

        Printer.FontSize = 12
        Printer.CurrentX = 1600
        Printer.CurrentY = 0
        Printer.FontBold = True
        Printer.Print LM_DULANG
        Printer.FontBold = False
        
        Printer.FontSize = 6
        Printer.FontBold = True
        
        Printer.FontName = "Text"

        Printer.FontSize = 10
        Printer.CurrentX = 1250
        Printer.CurrentY = 300
        Printer.FontBold = True
        Printer.Print "Berat"
        
        Printer.FontSize = 8
        Printer.CurrentX = 1250
        Printer.CurrentY = 500
        Printer.FontBold = True
        Printer.Print LM_BERAT
        
        Printer.FontSize = 8
        Printer.CurrentX = 1250
        Printer.CurrentY = 700
        'Printer.FontBold = False
        Printer.Print LM_PURITY
        
        Printer.FontSize = 4
        Printer.Print vbNullString
        
        Printer.FontSize = 7
        
        Printer.Print "Wg:" & LM_BERAT
        Printer.Print "P:" & LM_PURITY
        Printer.Print "No. Cert:" & LM_CERT
        Printer.Print "T:" & LM_DULANG
        
        Printer.FontSize = 7
        Printer.FontSize = G_Frm56_LM_SIZE_NAMA_KEDAI
        'Printer.Print vbNullString
        
        LM_GST = 0
        
        If Not IsNull(rs!gst_ari_nashi) Then
            If rs!gst_ari_nashi = 0 Then
                LM_GST = 0
            ElseIf rs!gst_ari_nashi = 1 Then
                LM_GST = 1
            End If
        End If
        
        If LM_GST = 1 Then
            Printer.Print LM_NAMA_KEDAI
        ElseIf LM_GST = 0 Then
            Printer.Print LM_NAMA_KEDAI & "*"
        End If
        
        'Printer.Print LM_NAMA_KEDAI
        Printer.FontBold = False
        Printer.FontName = "Text"
        Printer.FontSize = 7 '''Asal : Digunakan Kebanyakkan Kedai
        
        Printer.EndDoc
        
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
Else
    'MsgBox "Barcode Label Printer [ZDesigner GT800 (EPL)] Tidak Dijumpai.", vbCritical, "Sila Install Software Untuk Printer"
    MsgBox "Barcode Label Printer [" & G_PRINTER_BARCODE & "] Tidak Dijumpai.", vbCritical, "Sila Install Software Untuk Printer"
End If
End Sub
Sub cetak_barcode_gb_all()
'On Error Resume Next
'Print semua barcode bagi item yang dimasukkan mengikut batch kemasukkan data ke dalam sistem
Dim oPrn As Printer
Dim LM_NAMA_KEDAI As String

PRINTER_FOUND = 0 '0 : Barcode Label Printer Not Found , 1 : Barcode Label Printer Found

For Each oPrn In Printers
    'If oPrn.DeviceName = "ZDesigner GT800 (EPL)" Then
    If oPrn.DeviceName = G_PRINTER_BARCODE Then
        Set Printer = oPrn
        PRINTER_FOUND = 1 '0 : Barcode Label Printer Not Found , 1 : Barcode Label Printer Found
        Exit For
    End If
Next

LM_HEAD = "Sankyu System"

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai_3) Then LM_NAMA_KEDAI = rs!nama_kedai_3
    If Not IsNull(rs!bar_header) Then LM_HEAD = rs!bar_header
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End
        
If PRINTER_FOUND = 1 Then '0 : Barcode Label Printer Not Found , 1 : Barcode Label Printer Found
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from Data_Database where NoRujukanSistem='" & GM_No_RUJUKAN_BELIAN & "' AND StatusItem <> 0", cn, adOpenKeyset, adLockOptimistic
    
    While rs.EOF = False
        LM_NO_SIRI = vbNullString
        LM_BERAT = vbNullString
        LM_PURITY = vbNullString
        LM_CERT = vbNullString
        LM_DULANG = vbNullString
        
        LM_GST = 0
        
        If Not IsNull(rs!gst_ari_nashi) Then
            If rs!gst_ari_nashi = 0 Then
                LM_GST = 0
            ElseIf rs!gst_ari_nashi = 1 Then
                LM_GST = 1
            End If
        End If
        
        If Not IsNull(rs!no_siri_Produk) Then LM_NO_SIRI = rs!no_siri_Produk 'No. Siri Produk
        If Not IsNull(rs!Berat) Then LM_BERAT = Format(rs!Berat, "#,##0.00 g") 'Berat (g)
        If Not IsNull(rs!kod_Purity) Then LM_PURITY = rs!kod_Purity 'Purity
        If Not IsNull(rs!dulang) Then LM_DULANG = rs!dulang 'Dulang
        If Not IsNull(rs!no_cert) Then 'No. Cert
            LM_CERT = rs!no_cert
        Else
            LM_CERT = "-"
        End If
        
        Printer.FontName = "Andalus"
        Printer.FontSize = 8
        'Printer.CurrentX = 500
        'Printer.CurrentY = 0
        Printer.FontBold = True
        Printer.Print LM_HEAD
        Printer.FontBold = False
        Printer.FontName = "Text"
        Printer.FontSize = 6
        Printer.Print vbNullString
        
        'Printer.FontName = "Code128"
        Printer.FontName = "Code39"
        Printer.FontSize = 24
        Printer.CurrentX = 10
        Printer.CurrentY = 220
        Printer.FontBold = True
        Printer.Print LM_NO_SIRI
        'Printer.Print vbNullString
        
        Printer.FontName = "Text"
        
        Printer.FontSize = 10
        Printer.CurrentX = 10
        Printer.CurrentY = 700
        Printer.FontBold = True
        'Printer.Print LM_NO_SIRI
        If G_BARCODE_READABLE = "YES" Then
            Printer.Print LM_NO_SIRI
        Else
            Printer.Print vbNullString
        End If

        Printer.FontSize = 12
        Printer.CurrentX = 1600
        Printer.CurrentY = 0
        Printer.FontBold = True
        Printer.Print LM_DULANG
        Printer.FontBold = False
        
        Printer.FontSize = 6
        Printer.FontBold = True
        
        Printer.FontName = "Text"

        Printer.FontSize = 10
        Printer.CurrentX = 1250
        Printer.CurrentY = 300
        Printer.FontBold = True
        Printer.Print "Berat"
        
        Printer.FontSize = 8
        Printer.CurrentX = 1250
        Printer.CurrentY = 500
        Printer.FontBold = True
        Printer.Print LM_BERAT
        
        Printer.FontSize = 8
        Printer.CurrentX = 1250
        Printer.CurrentY = 700
        'Printer.FontBold = False
        Printer.Print LM_PURITY
        
        Printer.FontSize = 4
        Printer.Print vbNullString
        
        Printer.FontSize = 7
        
        Printer.Print "Wg:" & LM_BERAT
        Printer.Print "P:" & LM_PURITY
        Printer.Print "No. Cert:" & LM_CERT
        Printer.Print "T:" & LM_DULANG
        
        Printer.FontSize = 7
        Printer.FontSize = G_Frm56_LM_SIZE_NAMA_KEDAI
        'Printer.Print vbNullString
        'Printer.Print LM_NAMA_KEDAI
        
        If LM_GST = 1 Then
            Printer.Print LM_NAMA_KEDAI
        ElseIf LM_GST = 0 Then
            Printer.Print LM_NAMA_KEDAI & "*"
        End If
        
        Printer.FontBold = False
        Printer.FontName = "Text"
        Printer.FontSize = 7 '''Asal : Digunakan Kebanyakkan Kedai
        
        Printer.EndDoc
        
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
Else
    'MsgBox "Barcode Label Printer [ZDesigner GT800 (EPL)] Tidak Dijumpai.", vbCritical, "Sila Install Software Untuk Printer"
    MsgBox "Barcode Label Printer [" & G_PRINTER_BARCODE & "] Tidak Dijumpai.", vbCritical, "Sila Install Software Untuk Printer"
End If
End Sub
Sub frm56_font_type()
'On Error GoTo logging:
Frm56.CBB20.Clear
Frm56.CBB21.Clear
Frm56.CBB22.Clear
Frm56.CBB23.Clear

LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 92_barcode_font order by font_type ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!font_type) Then
        Frm56.CBB20.AddItem rs!font_type
        Frm56.CBB21.AddItem rs!font_type
        Frm56.CBB22.AddItem rs!font_type
        Frm56.CBB23.AddItem rs!font_type
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
    
Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " DrawBar : frm56_font_type" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If LM_ERR_NO = "3704" Or LM_ERR_NO = "-2147467259" Or LM_ERR_NO = "-2147217887" Then
    Call Main
    If LM_CONN = 1 Then
        Resume re_conn_1:
    End If
Else
    Resume Next
End If
End Sub
Sub frm56_initial_frame()
'On Error GoTo logging:
Frm56.Frame1.Visible = False
Frm56.Frame2.Visible = False
Frm56.Frame3.Visible = False

Frm56.Frame1.Top = 120
Frm56.Frame1.Left = 1680
Frm56.Frame2.Top = 120
Frm56.Frame2.Left = 1680
Frm56.Frame3.Top = 120
Frm56.Frame3.Left = 1680

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_log : frm56_initial_frame" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Sub frm56_initial_frame2()
'On Error GoTo logging:
Frm56.Frame5.Visible = False

Frm56.Frame5.Top = 2160
Frm56.Frame5.Left = 120

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " mod_log : frm56_initial_frame2" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Sub frm56_senarai_printer()
'On Error GoTo logging:
Frm56.LV2.ListItems.Clear

LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select perkara from layout_barcode order by perkara ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!perkara) Then
        With Frm56.LV2
            Set .SmallIcons = Frm56.ImageList4
            Set .Icons = Frm56.ImageList4
            
            .ListItems.Add , rs!perkara, rs!perkara, 41
        End With
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
    
Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " DrawBar : frm56_senarai_printer" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If LM_ERR_NO = "3704" Or LM_ERR_NO = "-2147467259" Or LM_ERR_NO = "-2147217887" Then
    Call Main
    If LM_CONN = 1 Then
        Resume re_conn_1:
    End If
Else
    Resume Next
End If
End Sub
Sub frm56_reset_element()
'On Error GoTo logging:
For x = 0 To 11
    Frm56.TB1(x) = 0
    Frm56.TB2(x) = vbNullString
    Frm56.TB3(x) = 0
    Frm56.TB4(x) = 0
    Frm56.CB10(x) = 0
    Frm56.CB11(x) = 0
Next x

If Frm56.L15_Text = "0" Then '0 : Type A , 1 : Type B , 2 : Type C
    LM_JENIS = 0
ElseIf Frm56.L15_Text = "1" Then '0 : Type A , 1 : Type B , 2 : Type C
    LM_JENIS = 1
ElseIf Frm56.L15_Text = "2" Then '0 : Type A , 1 : Type B , 2 : Type C
    LM_JENIS = 2
End If

For x = 0 To 11
    Frm56.TB1(x) = Split(G_SKU_SIZE(LM_JENIS), ",")(x)
    Frm56.TB2(x) = Split(G_SKU_FONT(LM_JENIS), ",")(x)
    Frm56.TB3(x) = Split(G_SKU_POS_X(LM_JENIS), ",")(x)
    Frm56.TB4(x) = Split(G_SKU_POS_Y(LM_JENIS), ",")(x)
    Frm56.CB10(x).Value = Split(G_SKU_BOLD(LM_JENIS), ",")(x)
    Frm56.CB11(x).Value = Split(G_SKU_ITALIC(LM_JENIS), ",")(x)
Next x

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " DrawBar : frm56_reset_element" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Sub frm56_recall_setting_barcode()
'On Error GoTo logging:
Dim LM_SKU_SIZE(2)
Dim LM_SKU_FONT(2)
Dim LM_SKU_POS_X(2)
Dim LM_SKU_POS_Y(2)
Dim LM_SKU_BOLD(2)
Dim LM_SKU_ITALIC(2)

For x = 0 To 11
    Frm56.TB1(x) = 0
    Frm56.TB2(x) = vbNullString
    Frm56.TB3(x) = 0
    Frm56.TB4(x) = 0
    Frm56.CB10(x) = 0
    Frm56.CB11(x) = 0
Next x

If Frm56.L15_Text = "0" Then '0 : Type A , 1 : Type B , 2 : Type C
    LM_JENIS = 0
ElseIf Frm56.L15_Text = "1" Then '0 : Type A , 1 : Type B , 2 : Type C
    LM_JENIS = 1
ElseIf Frm56.L15_Text = "2" Then '0 : Type A , 1 : Type B , 2 : Type C
    LM_JENIS = 2
End If


LM_CONN = 1
re_conn_1:
Set rs1 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs1.Open "select * from layout_barcode where perkara='" & G_ID & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs1.EOF Then
    If Not IsNull(rs1!a_font_size) Then LM_SKU_SIZE(0) = rs1!a_font_size
    If Not IsNull(rs1!a_font_type) Then LM_SKU_FONT(0) = rs1!a_font_type
    If Not IsNull(rs1!a_position_x) Then LM_SKU_POS_X(0) = rs1!a_position_x
    If Not IsNull(rs1!a_position_y) Then LM_SKU_POS_Y(0) = rs1!a_position_y
    If Not IsNull(rs1!a_bold) Then LM_SKU_BOLD(0) = rs1!a_bold
    If Not IsNull(rs1!a_italic) Then LM_SKU_ITALIC(0) = rs1!a_italic
    
    If Not IsNull(rs1!b_font_size) Then LM_SKU_SIZE(1) = rs1!b_font_size
    If Not IsNull(rs1!b_font_type) Then LM_SKU_FONT(1) = rs1!b_font_type
    If Not IsNull(rs1!b_position_x) Then LM_SKU_POS_X(1) = rs1!b_position_x
    If Not IsNull(rs1!b_position_y) Then LM_SKU_POS_Y(1) = rs1!b_position_y
    If Not IsNull(rs1!b_bold) Then LM_SKU_BOLD(1) = rs1!b_bold
    If Not IsNull(rs1!b_italic) Then LM_SKU_ITALIC(1) = rs1!b_italic
    
    If Not IsNull(rs1!c_font_size) Then LM_SKU_SIZE(2) = rs1!c_font_size
    If Not IsNull(rs1!c_font_type) Then LM_SKU_FONT(2) = rs1!c_font_type
    If Not IsNull(rs1!c_position_x) Then LM_SKU_POS_X(2) = rs1!c_position_x
    If Not IsNull(rs1!c_position_y) Then LM_SKU_POS_Y(2) = rs1!c_position_y
    If Not IsNull(rs1!c_bold) Then LM_SKU_BOLD(2) = rs1!c_bold
    If Not IsNull(rs1!c_italic) Then LM_SKU_ITALIC(2) = rs1!c_italic
End If

rs1.Close
Set rs1 = Nothing

For x = 0 To 11
    Frm56.TB1(x) = Split(LM_SKU_SIZE(LM_JENIS), ",")(x)
    Frm56.TB2(x) = Split(LM_SKU_FONT(LM_JENIS), ",")(x)
    Frm56.TB3(x) = Split(LM_SKU_POS_X(LM_JENIS), ",")(x)
    Frm56.TB4(x) = Split(LM_SKU_POS_Y(LM_JENIS), ",")(x)
    Frm56.CB10(x).Value = Split(LM_SKU_BOLD(LM_JENIS), ",")(x)
    Frm56.CB11(x).Value = Split(LM_SKU_ITALIC(LM_JENIS), ",")(x)
Next x

Exit Sub
logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " DrawBar : frm56_recall_setting_barcode" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If LM_ERR_NO = "3704" Or LM_ERR_NO = "-2147467259" Or LM_ERR_NO = "-2147217887" Then
    Call Main
    If LM_CONN = 1 Then
        Resume re_conn_1:
    End If
Else
    Resume Next
End If
End Sub
