Attribute VB_Name = "Module1"
Public cn As ADODB.Connection
Public cn2 As ADODB.Connection
Public cn3 As ADODB.Connection
Public rs As ADODB.Recordset
Public LogDate_Memory As String
Public LogAct_Memory As String
Public GLOBAL_DISABLE As String
Public G_SERVER_IP
Public G_SERVER_USER
Public G_SERVER_PASS
Public G_SERVER_DATABASE
Public G_SERVER_PORT
Public G_PRINTER_BARCODE
Public G_TERMINAL
Public G_AUTO_BACKUP
Public G_RECOVERY_DATABASE
Public G_BELIAN_TEMP As String
Public G_JUALAN_TEMP As String
Public G_NE_PATH As String
Public G_ID
Public G_TAHUN
Public G_LOGIN_USER
Public G_SERVICE_TEMP As String
Public G_JENIS_URUSAN
Public GM_NO_SIRI
Public G_TEMPAHAN '0 : Padam data , 1 : Tukar status kepada belum siap
Public GM_NEXT_PREV As Single '0 : Next , 1 : Previous
Public GM_REPORT_MODE As Single '0 : Report belian , 1 : Report mengikut berat , 2 : Report jualan , 3 : Report mengikut invoice (jualan) , 4 : Report mengikut invoice (buyback) , 5 : Report buyback , 6 : Report Belian Gold Bar , 7 : Report Buyback Gold Bar , 8 : Report mengikut no invoice supplier , 9 : Report mengikut no. siri produk , 10 : Report stok , 11 : Report potong , 12 : Report ansuran , 13 : Report tempahan
Public G_FORM_OUT_DESC
Public G_FORM_LIST
Public G_GDN_TEMP As String
Public G_GRN_TEMP As String
Public G_BARCODE_READABLE
Public G_KOD_KEDAI
Public G_BIL_JUALAN As Integer
Public G_PURITY_JUALAN(20)
Public G_LOCK_JURUJUAL
Public G_PREVIEW
Public G_INVOICE_TEMP As String
Public G_GST_SYSTEM
Public G_DATABASE_SETTING
Public G_AGIHAN_TEMP As String
Public G_PULANGAN_TEMP As String
Public G_SPKE_ME_MAIL
Public G_SPKE_NE_PATH
Public G_SYSTEM_TYPE
Public G_CAWANGAN
Public G_KEDAI
Public G_JENIS_HEADER '0 : Pre Printed , 1 : Sistem
Public G_NAMA_KEDAI
Public G_NO_PENDAFTARAN
Public G_ALAMAT
Public G_NO_TEL
Public G_NO_GST
Public G_FLAG_GST
Public G_RATE_GST
Public G_RIYAL
Public G_UPAH_SUPPLIER
Public G_SPREAD
Public G_PRINT_BARCODE '0 : Tidak Cetak Barcode , 1 : Cetak Barcode
Public G_GST_INCOMING
Public G_GST_JUAL
Public G_FLAG_BIL_GST
Public G_SPREAD_TI
Public G_GST_JUALAN_INC
Public G_KADAR_COMM_STAFF
Public G_DISC_ARI_NASHI
Public G_DISC_JUMLAH
Public G_LIMIT_INVOICE
Public G_KUPON_DISC
Public G_TOP
Public G_HARGA_999
Public G_SCANNER_MODE '0 : Scanner tidak digunakan , 1 : Scanner digunakan
Public G_UPAH_MODE
Public G_KIRAAN_UPAH
Public G_LM_NAMA_KEDAI
Public G_LM_HEAD
Public G_Frm56_LM_SIZE_1
Public G_Frm56_LM_SIZE_2
Public G_Frm56_LM_SIZE_3
Public G_Frm56_LM_SIZE_4
Public G_Frm56_LM_SIZE_NAMA_KEDAI
Public G_L_JENIS_BARCODE
Public G_BAROCDE_LINE_1
Public G_BAROCDE_LINE_2
Public G_BAROCDE_LINE_3
Public G_BAROCDE_LINE_4
Public G_PEMALAR_BONUS_BIASA As Double
Public G_PEMALAR_BONUS_SILVER As Double
Public G_PEMALAR_BONUS_GOLD As Double
Public G_PEMALAR_BONUS_PLATINUM As Double
Public G_PEMALAR_TEBUS_BIASA As Double
Public G_PEMALAR_TEBUS_SILVER As Double
Public G_PEMALAR_TEBUS_GOLD As Double
Public G_PEMALAR_TEBUS_PLATINUM As Double
Public G_LEVEL_USER As Single
Public G_NAMA_BARCODE(12)
Public G_EXPIRED
Public G_GDN_SUBSCRIBE
Public G_INVOICE_RASMI
Public G_TRAY_X
Public G_JENIS_JUALAN
Public G_AUTO_INSERT
Public G_CALC_AUTO
Public G_MENU_ADMIN

Public G_L1_LEFT
Public G_L1_TOP
Public G_L1_BOLD
Public G_L1_ITALIC
Public G_L1_FONT
Public G_L2_LEFT
Public G_L2_TOP
Public G_L2_BOLD
Public G_L2_ITALIC
Public G_L2_FONT
Public G_L3_LEFT
Public G_L3_TOP
Public G_L3_BOLD
Public G_L3_ITALIC
Public G_L3_FONT
Public G_L4_LEFT
Public G_L4_TOP
Public G_L4_BOLD
Public G_L4_ITALIC
Public G_L4_FONT
Public G_L5_LEFT
Public G_L5_TOP
Public G_L5_BOLD
Public G_L5_ITALIC
Public G_L5_FONT
Public G_L6_LEFT
Public G_L6_TOP
Public G_L6_BOLD
Public G_L6_ITALIC
Public G_L6_FONT
Public G_L7_LEFT
Public G_L7_TOP
Public G_L7_BOLD
Public G_L7_ITALIC
Public G_L7_FONT
Public G_L8_LEFT
Public G_L8_TOP
Public G_L8_BOLD
Public G_L8_ITALIC
Public G_L8_FONT
Public G_L9_LEFT
Public G_L9_TOP
Public G_L9_BOLD
Public G_L9_ITALIC
Public G_L9_FONT
Public G_L10_LEFT
Public G_L10_TOP
Public G_L10_BOLD
Public G_L10_ITALIC
Public G_L10_FONT
Public G_L11_LEFT
Public G_L11_TOP
Public G_L11_BOLD
Public G_L11_ITALIC
Public G_L11_FONT
Public G_L12_LEFT
Public G_L12_TOP
Public G_L12_BOLD
Public G_L12_ITALIC
Public G_L12_FONT
Public G_L13_LEFT
Public G_L13_TOP
Public G_L13_BOLD
Public G_L13_ITALIC
Public G_L13_FONT
Public G_L14_LEFT
Public G_L14_TOP
Public G_L14_BOLD
Public G_L14_ITALIC
Public G_L14_FONT
Public G_L15_LEFT
Public G_L15_TOP
Public G_L15_BOLD
Public G_L15_ITALIC
Public G_L15_FONT
Public G_L16_LEFT
Public G_L16_TOP
Public G_L16_BOLD
Public G_L16_ITALIC
Public G_L16_FONT
Public G_L17_LEFT
Public G_L17_TOP
Public G_L17_BOLD
Public G_L17_ITALIC
Public G_L17_FONT
Public G_L18_LEFT
Public G_L18_TOP
Public G_L18_BOLD
Public G_L18_ITALIC
Public G_L18_FONT

Public G_L1_WIDTH
Public G_L1_HEIGHT
Public G_L2_WIDTH
Public G_L2_HEIGHT
Public G_L3_WIDTH
Public G_L3_HEIGHT
Public G_L4_WIDTH
Public G_L4_HEIGHT
Public G_L5_WIDTH
Public G_L5_HEIGHT
Public G_L6_WIDTH
Public G_L6_HEIGHT
Public G_L7_WIDTH
Public G_L7_HEIGHT
Public G_L8_WIDTH
Public G_L8_HEIGHT
Public G_L9_WIDTH
Public G_L9_HEIGHT
Public G_L10_WIDTH
Public G_L10_HEIGHT
Public G_L11_WIDTH
Public G_L11_HEIGHT
Public G_L12_WIDTH
Public G_L12_HEIGHT
Public G_L13_WIDTH
Public G_L13_HEIGHT
Public G_L14_WIDTH
Public G_L14_HEIGHT
Public G_L15_WIDTH
Public G_L15_HEIGHT
Public G_L16_WIDTH
Public G_L16_HEIGHT
Public G_L17_WIDTH
Public G_L17_HEIGHT
Public G_L18_WIDTH
Public G_L18_HEIGHT
Public G_X
Public G_DEV_PASS
Public G_DEV_PASS_DEFAULT
Public G_INVOICE_TYPE
Public G_PRINTER_TI_MODE

Public G_TI_TRADE_IN
Public G_TI_BUYBACK
Public G_TI_CAJ
Public G_TI_MODE
Public G_TI_BERAT
Public G_TI_RATE_TI
Public G_TI_RATE_BB
Public G_TI_RATE_TUKAR
Public G_TRADE_IN_TOTAL
Public G_TRADE_IN_CAJ
Public G_TI_MEMORY(3, 3)
Public G_JENIS_BARCODE_PRINTER
Sub Main()
'On Error Resume Next
Dim LM_SERVER As String
DATABASEMODE = 1 '0 : Access , 1 : MySQL
LM_OPEN = 0

If Len(G_AUTO_INSERT) = 0 Or Len(G_GDN_SUBSCRIBE) = 0 Or Len(G_SYSTEM_TYPE) = 0 Or Len(G_AGIHAN_TEMP) = 0 Or Len(G_PULANGAN_TEMP) = 0 Or Len(G_GST_SYSTEM) = 0 Or Len(G_LOCK_JURUJUAL) = 0 Or Len(G_BARCODE_READABLE) = 0 Or Len(G_GDN_TEMP) = 0 Or Len(G_NE_PATH) = 0 Or Len(G_AUTO_BACKUP) = 0 Or Len(G_BELIAN_TEMP) = 0 Or Len(G_JUALAN_TEMP) = 0 Or Len(G_TERMINAL) = 0 Or Len(G_SERVER_IP) = 0 Or Len(G_SERVER_USER) = 0 Or Len(G_SERVER_PASS) = 0 Or Len(G_SERVER_DATABASE) = 0 Or Len(G_SERVER_PORT) = 0 Or Len(G_PRINTER_BARCODE) = 0 Or Len(G_RECOVERY_DATABASE) = 0 Or Len(G_TAHUN) = 0 Or Len(G_SERVICE_TEMP) = 0 Then

    Call system_configuration

End If

'### Set date format
Call SetDateTime

'If Not cn Is Nothing Then

'    Select Case cn.State
        
'        Case adStateClosed
'            MsgBox "close"
'            LM_OPEN = 0
'        Case adStateOpen
'            MsgBox "open"
'            LM_OPEN = 1
'    End Select
    
'Else

'End If
If DATABASEMODE = 0 Then
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    cn.ConnectionString = "Provider=microsoft.jet.oledb.4.0 ; data source = " & (App.Path & "\Database.mdb") & ";"
    cn.Open
End If

'MsgBox cn
'Exit Sub

If DATABASEMODE = 1 And LM_OPEN = 0 Then
    Set cn = New ADODB.Connection
    cn.ConnectionString = "Driver={MySQL ODBC 3.51 Driver}; Server=" & G_SERVER_IP & ";port=" & G_SERVER_PORT & "; database=" & G_SERVER_DATABASE & "; user=" & G_SERVER_USER & "; password=" & G_SERVER_PASS & "; option=3;"
    
    cn.Open

    If cn.State = adStateOpen Then
        'MsgBox "Connected"
    Else
        MsgBox "Tiada connection antara sistem dan database , Sila pastikan XAMPP diaktifkan!", vbCritical, "Error"
        End
    End If
    
    MDI_frm1.L18_Text = "1"
End If
End Sub
Sub Main3()
'On Error Resume Next
Dim LM_SERVER As String

If Len(G_AUTO_INSERT) = 0 Or Len(G_GDN_SUBSCRIBE) = 0 Or Len(G_SYSTEM_TYPE) = 0 Or Len(G_AGIHAN_TEMP) = 0 Or Len(G_PULANGAN_TEMP) = 0 Or Len(G_DATABASE_SETTING) = 0 Or Len(G_GST_SYSTEM) = 0 Or Len(G_LOCK_JURUJUAL) = 0 Or Len(G_BARCODE_READABLE) = 0 Or Len(G_GDN_TEMP) = 0 Or Len(G_NE_PATH) = 0 Or Len(G_AUTO_BACKUP) = 0 Or Len(G_BELIAN_TEMP) = 0 Or Len(G_JUALAN_TEMP) = 0 Or Len(G_TERMINAL) = 0 Or Len(G_SERVER_IP) = 0 Or Len(G_SERVER_USER) = 0 Or Len(G_SERVER_PASS) = 0 Or Len(G_SERVER_DATABASE) = 0 Or Len(G_SERVER_PORT) = 0 Or Len(G_PRINTER_BARCODE) = 0 Or Len(G_RECOVERY_DATABASE) = 0 Or Len(G_TAHUN) = 0 Or Len(G_SERVICE_TEMP) = 0 Then

    Call system_configuration

End If

'### Set date format
Call SetDateTime

Set cn2 = New ADODB.Connection
cn2.ConnectionString = "Driver={MySQL ODBC 3.51 Driver}; Server=" & G_SERVER_IP & ";port=" & G_SERVER_PORT & "; database=" & G_DATABASE_SETTING & "; user=" & G_SERVER_USER & "; password=" & G_SERVER_PASS & "; option=3;"

cn2.Open

If cn2.State = adStateOpen Then
    'MsgBox "Connected"
Else
    MsgBox "Tiada connection antara sistem dan database , Sila pastikan XAMPP diaktifkan!", vbCritical, "Error"
    End
End If

MDI_frm1.L19_Text = "1"
End Sub
Sub Main2()
'On Error Resume Next
Dim LM_SERVER As String

If Len(G_AUTO_INSERT) = 0 Or Len(G_GDN_SUBSCRIBE) = 0 Or Len(G_SYSTEM_TYPE) = 0 Or Len(G_AGIHAN_TEMP) = 0 Or Len(G_PULANGAN_TEMP) = 0 Or Len(G_GST_SYSTEM) = 0 Or Len(G_LOCK_JURUJUAL) = 0 Or Len(G_BARCODE_READABLE) = 0 Or Len(G_GDN_TEMP) = 0 Or Len(G_NE_PATH) = 0 Or Len(G_AUTO_BACKUP) = 0 Or Len(G_BELIAN_TEMP) = 0 Or Len(G_JUALAN_TEMP) = 0 Or Len(G_TERMINAL) = 0 Or Len(G_SERVER_IP) = 0 Or Len(G_SERVER_USER) = 0 Or Len(G_SERVER_PASS) = 0 Or Len(G_SERVER_DATABASE) = 0 Or Len(G_SERVER_PORT) = 0 Or Len(G_PRINTER_BARCODE) = 0 Or Len(G_RECOVERY_DATABASE) = 0 Or Len(G_TAHUN) = 0 Or Len(G_SERVICE_TEMP) = 0 Then

    Call system_configuration

End If

'### Set date format
Call SetDateTime

Set cn3 = New ADODB.Connection
cn3.ConnectionString = "Driver={MySQL ODBC 3.51 Driver}; Server=" & G_SERVER_IP & ";port=" & G_SERVER_PORT & "; database=" & G_RECOVERY_DATABASE & "; user=" & G_SERVER_USER & "; password=" & G_SERVER_PASS & "; option=3;"

cn3.Open

If cn3.State = adStateOpen Then
    'MsgBox "Connected"
Else
    MsgBox "Tiada connection antara sistem dan database , Sila pastikan XAMPP diaktifkan!", vbCritical, "Error"
    End
End If

MDI_frm1.L22_Text = "1"
End Sub
Sub UnloadLoading()
'On Error Resume Next
'Application.ScreenUpdating = False
Frm2.MSFlexGrid1.Clear
Frm2.MSFlexGrid1.FormatString = "< Tarikh Dan Masa |< Log Aktiviti"
Frm2.MSFlexGrid1.ColWidth(0) = 2000
Frm2.MSFlexGrid1.ColWidth(1) = 7000
    
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from Log order by ID DESC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm2.MSFlexGrid1.Rows = x + 1
    If Not IsNull(rs!Log_Tarikh) Then Frm2.MSFlexGrid1.TextMatrix(x, 0) = rs!Log_Tarikh
    If Not IsNull(rs!Log_Aktiviti) Then Frm2.MSFlexGrid1.TextMatrix(x, 1) = rs!Log_Aktiviti
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm2.Show
'Application.ScreenUpdating = True
End Sub
Sub check_db_conn_main()
'On Error Resume Next
LM_OPEN = 0

If G_SYSTEM_TYPE = "ONLINE" Then

    If MDI_frm1.L17_Text = "OFFLINE" Then
        
        MsgBox "Tiada sambungan internet. Sila pastikan komputer anda disambungkan dengan internet.", vbCritical, "Connection Failed"
        
        Exit Sub
        
    End If

End If

'If Not cn Is Nothing Then
'    LM_OPEN = 1
'Else
'    LM_OPEN = 0
'End If

If MDI_frm1.L18_Text = "0" Then
    If MDI_frm1.L17_Text = "ONLINE" And G_SYSTEM_TYPE = "ONLINE" Then Call Main
End If

If G_SYSTEM_TYPE = "OFFLINE" Then Call Main
End Sub
Sub check_db_conn_main2()
'On Error Resume Next
If G_SYSTEM_TYPE = "ONLINE" Then

    If MDI_frm1.L17_Text = "OFFLINE" Then
        
        MsgBox "Tiada sambungan internet. Sila pastikan komputer anda disambungkan dengan internet.", vbCritical, "Connection Failed"
        
        Exit Sub
        
    End If

End If

If MDI_frm1.L22_Text = "0" Then
    If MDI_frm1.L17_Text = "ONLINE" And G_SYSTEM_TYPE = "ONLINE" Then Call Main2
End If

If G_SYSTEM_TYPE = "OFFLINE" Then Call Main2
End Sub
Sub check_db_conn_main3()
'On Error Resume Next
If G_SYSTEM_TYPE = "ONLINE" Then

    If MDI_frm1.L17_Text = "OFFLINE" Then
        
        MsgBox "Tiada sambungan internet. Sila pastikan komputer anda disambungkan dengan internet.", vbCritical, "Connection Failed"
        
        Exit Sub
        
    End If

End If

If MDI_frm1.L19_Text = "0" Then
    If MDI_frm1.L17_Text = "ONLINE" And G_SYSTEM_TYPE = "ONLINE" Then Call Main3
End If

If G_SYSTEM_TYPE = "OFFLINE" Then Call Main3
End Sub
Sub system_configuration()
'On Error Resume Next
Dim File_Path As String
File_Path = App.Path & "\system_configuration.txt"
Open File_Path For Input As #1

Line Input #1, G_SERVER_IP
Line Input #1, G_SERVER_USER
Line Input #1, G_SERVER_PASS
Line Input #1, G_SERVER_DATABASE
Line Input #1, G_SERVER_PORT
Line Input #1, G_PRINTER_BARCODE
Line Input #1, G_TERMINAL
Line Input #1, G_BELIAN_TEMP
Line Input #1, G_JUALAN_TEMP
Line Input #1, G_AUTO_BACKUP
Line Input #1, G_NE_PATH
Line Input #1, G_RECOVERY_DATABASE
Line Input #1, G_TAHUN
Line Input #1, G_SERVICE_TEMP
Line Input #1, G_FORM_OUT_DESC
Line Input #1, G_FORM_LIST
Line Input #1, G_GDN_TEMP
Line Input #1, G_GRN_TEMP
Line Input #1, G_BARCODE_READABLE
Line Input #1, G_LOCK_JURUJUAL
Line Input #1, G_INVOICE_TEMP
Line Input #1, G_GST_SYSTEM
Line Input #1, G_AGIHAN_TEMP
Line Input #1, G_PULANGAN_TEMP
Line Input #1, G_SPKE_ME_MAIL
Line Input #1, G_SPKE_NE_PATH
Line Input #1, G_SYSTEM_TYPE
Line Input #1, G_GDN_SUBSCRIBE
Line Input #1, G_AUTO_INSERT
Line Input #1, G_MENU_ADMIN

Close #1

'Call trial
End Sub
Sub trial()
'on error resume next
Dim DATE_TODAY As Date
Dim DATE_EXP As Date

DATE_TODAY = DateTime.Date
DATE_EXP = G_EXPIRED

If DATE_TODAY > DATE_EXP Then

    MsgBox "Sistem anda telah disekat. Sila hubungi pihak Sankyu System.", vbCritical, "Info"
    
    End
    
End If
End Sub
Sub main_setting_kedai()
'on error resume next
If MDI_frm1.L20_Text = "Semua cawangan" Then
    
    LM_CAWANGAN = "HQ"
    
Else
    
    LM_CAWANGAN = MDI_frm1.L20_Text
    
End If

G_FLAG_GST = 0 '0 : Tiada id gst , 1 : Ada id gst

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan ='" & LM_CAWANGAN & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    G_CAWANGAN = MDI_frm1.L20_Text
    MDI_frm1.L20_Text = G_CAWANGAN

    If Not IsNull(rs!kod_cawangan) Then G_KOD_KEDAI = rs!kod_cawangan
    If Not IsNull(rs!setting_database) Then G_DATABASE_SETTING = rs!setting_database
    
    If Not IsNull(rs!nama_kedai) Then G_NAMA_KEDAI = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then G_NO_PENDAFTARAN = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then G_ALAMAT = rs!alamat
    If Not IsNull(rs!no_tel) Then G_NO_TEL = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then G_NO_GST = rs!no_id_gst
    If Not IsNull(rs!gst_ari_nashi) Then
        If rs!gst_ari_nashi = 0 Then
            
            G_FLAG_GST = 0 '0 : Tiada id gst , 1 : Ada id gst
            
        ElseIf rs!gst_ari_nashi = 1 Then
            
            G_FLAG_GST = 1 '0 : Tiada id gst , 1 : Ada id gst
            
        End If
    Else
        
        G_FLAG_GST = 0 '0 : Tiada id gst , 1 : Ada id gst
        
    End If
    If Not IsNull(rs!check_invoice) Then LM_INV_CHECK = rs!check_invoice
    If Not IsNull(rs!qty_item) Then LM_QTY_CHECK = rs!qty_item
    
    LM_FOUND = 1
    
End If

rs.Close
Set rs = Nothing

Call Main3
End Sub
Sub main_setting()
'on error resume next
G_JENIS_HEADER = 1 '0 : Pre Printed , 1 : Sistem

If MDI_frm1.L4_Text = "Semua cawangan" Then
    LM_KEDAI = "HQ"
Else
    LM_KEDAI = MDI_frm1.L20_Text
End If
        
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting where Default1='" & LM_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Not IsNull(rs!top_margin) Then
        G_TOP = rs!top_margin
    Else
        G_TOP = 0
    End If
    If Not IsNull(rs!upah_staff) Then
        G_J_DISC_UPAH = rs!upah_staff 'Peratusan penurunan upah kepada staff
    Else
        G_J_DISC_UPAH = 0
    End If
    If Not IsNull(rs!diskaun_permata_staff) Then
        G_J_DISC_PERMATA = rs!diskaun_permata_staff 'Peratusan penurunan harga barang permata kepada staff
    Else
        G_J_DISC_PERMATA = 0
    End If

    If Not IsNull(rs!jenis_header) Then
        If rs!jenis_header = 0 Then
            G_JENIS_HEADER = 0 '0 : Pre Printed , 1 : Sistem
        ElseIf rs!jenis_header = 1 Then
            G_JENIS_HEADER = 1 '0 : Pre Printed , 1 : Sistem
        End If
    Else
        G_JENIS_HEADER = 1 '0 : Pre Printed , 1 : Sistem
    End If
    If Not IsNull(rs!gst_value) Then
        G_RATE_GST = rs!gst_value
    Else
        G_RATE_GST = 6
    End If
    If Not IsNull(rs!riyal) Then
        G_RIYAL = rs!riyal
    Else
        G_RIYAL = 1
    End If
    If Not IsNull(rs!upah_supplier) Then
        If rs!upah_supplier = 0 Then
            G_UPAH_SUPPLIER = 0 'Tetapan upah dari supplier , 0 : Lump sum , 1 : Upah per gram
        ElseIf rs!upah_supplier = 1 Then
            G_UPAH_SUPPLIER = 1 'Tetapan upah dari supplier , 0 : Lump sum , 1 : Upah per gram
        Else
            G_UPAH_SUPPLIER = 0 'Tetapan upah dari supplier , 0 : Lump sum , 1 : Upah per gram
        End If
    Else
        G_UPAH_SUPPLIER = 0 'Tetapan upah dari supplier , 0 : Lump sum , 1 : Upah per gram
    End If
    If Not IsNull(rs!spread_Cash_Trade_In) Then
        G_SPREAD = rs!spread_Cash_Trade_In 'Spread Trade In %
    Else
        G_SPREAD = 0
    End If
    If rs!BarcodeYesNo = 0 Then
        G_PRINT_BARCODE = 0
    Else
        G_PRINT_BARCODE = 1
    End If
    If rs!ScannerMode = 1 Then
        G_SCANNER_MODE = 1
    Else
        G_SCANNER_MODE = 0
    End If
    If rs!printer_mode_ti = 1 Then
        G_PRINTER_TI_MODE = 1
    Else
        G_PRINTER_TI_MODE = 0
    End If
    If Not IsNull(rs!gst_arinashi_belian) Then
        If rs!gst_arinashi_belian = 1 Then
            G_GST_INCOMING = 1
        Else
            G_GST_INCOMING = 0
        End If
    Else
        G_GST_INCOMING = 0
    End If
    If Not IsNull(rs!gst_arinashi) Then
        If rs!gst_arinashi = 1 Then
            G_GST_JUAL = 1
        Else
            G_GST_JUAL = 0
        End If
    Else
        G_GST_JUAL = 0
    End If
    If Not IsNull(rs!gst_jualan_included) Then
        If rs!gst_jualan_included = 1 Then
            G_GST_JUALAN_INC = 1
        ElseIf rs!gst_jualan_included = 0 Then
            G_GST_JUALAN_INC = 0
        End If
    Else
        G_GST_JUALAN_INC = 0
    End If
    If Not IsNull(rs!flag_bil_gst) Then G_FLAG_BIL_GST = rs!flag_bil_gst
    If Not IsNull(rs!potongan_trade_in) Then G_SPREAD_TI = rs!potongan_trade_in 'Potongan Harga Resit Trade in (%)
    If Not IsNull(rs!kadar_komisyen_upah) Then 'Kadar komisyen upah kepada agen dropship (%)
        G_KADAR_COMM_STAFF = rs!kadar_komisyen_upah
    Else
        G_KADAR_COMM_STAFF = 0
    End If
    If Not IsNull(rs!diskaun_ari_nashi) Then
        If rs!diskaun_ari_nashi = 1 Then
            G_DISC_ARI_NASHI = 1
        Else
            G_DISC_ARI_NASHI = 0
        End If
    Else
        G_DISC_ARI_NASHI = 0
    End If
    If Not IsNull(rs!diskaun) Then
        If IsNumeric(rs!diskaun) Then
        
            G_DISC_JUMLAH = rs!diskaun
        
        Else
            
            G_DISC_JUMLAH = 0
            
        End If
        
    Else
    
        G_DISC_JUMLAH = 0
        
    End If
    If Not IsNull(rs!invoice_type) Then
        G_LIMIT_INVOICE = rs!invoice_type 'Jumlah Limit Jualan
    Else
        G_LIMIT_INVOICE = 0
    End If
    If Not IsNull(rs!kupon_diskaun) Then
        If IsNumeric(rs!kupon_diskaun) Then
            G_KUPON_DISC = rs!kupon_diskaun
        Else
            G_KUPON_DISC = 0
        End If
    Else
        G_KUPON_DISC = 0
    End If
    If Not IsNull(rs!pemalar_bonus_biasa) Then
        G_PEMALAR_BONUS_BIASA = rs!pemalar_bonus_biasa
    Else
        G_PEMALAR_BONUS_BIASA = 0
    End If
    If Not IsNull(rs!pemalar_tebus_bonus_biasa) Then
        G_PEMALAR_TEBUS_BIASA = rs!pemalar_tebus_bonus_biasa
    Else
        G_PEMALAR_TEBUS_BIASA = 0
    End If
    If Not IsNull(rs!pemalar_bonus_silver) Then
        G_PEMALAR_BONUS_SILVER = rs!pemalar_bonus_silver
    Else
        G_PEMALAR_BONUS_SILVER = 0
    End If
    If Not IsNull(rs!pemalar_tebus_bonus_silver) Then
        G_PEMALAR_TEBUS_SILVER = rs!pemalar_tebus_bonus_silver
    Else
        G_PEMALAR_TEBUS_SILVER = 0
    End If
    
    If Not IsNull(rs!pemalar_bonus_gold) Then
        G_PEMALAR_BONUS_GOLD = rs!pemalar_bonus_gold
    Else
        G_PEMALAR_BONUS_GOLD = 0
    End If
    If Not IsNull(rs!pemalar_tebus_bonus_gold) Then
        G_PEMALAR_TEBUS_GOLD = rs!pemalar_tebus_bonus_gold
    Else
        G_PEMALAR_TEBUS_GOLD = 0
    End If
    If Not IsNull(rs!pemalar_bonus_platinum) Then
        G_PEMALAR_BONUS_PLATINUM = rs!pemalar_bonus_platinum
    Else
        G_PEMALAR_BONUS_PLATINUM = 0
    End If
    If Not IsNull(rs!pemalar_tebus_bonus_platinum) Then
        G_PEMALAR_TEBUS_PLATINUM = rs!pemalar_tebus_bonus_platinum
    Else
        G_PEMALAR_TEBUS_PLATINUM = 0
    End If
    If Not IsNull(rs!harga_999) Then
        If IsNumeric(rs!harga_999) Then
            G_HARGA_999 = rs!harga_999
        Else
            G_HARGA_999 = 0
        End If
    Else
        G_HARGA_999 = 0
    End If
    If Not IsNull(rs!flag_upah) Then
        If rs!flag_upah = 1 Then
            G_UPAH_MODE = 1
        Else
            G_UPAH_MODE = 0
        End If
    End If
    If Not IsNull(rs!kiraan_upah) Then '0 : Pengiraan upah mengikut berat barang , 1 : Pengiraan upah mengikut tetapan asal
        
        If rs!kiraan_upah = 0 Then
            G_KIRAAN_UPAH = 0 '0 : Pengiraan upah mengikut berat barang , 1 : Pengiraan upah mengikut tetapan asal
        ElseIf rs!kiraan_upah = 1 Then
            G_KIRAAN_UPAH = 1 '0 : Pengiraan upah mengikut berat barang , 1 : Pengiraan upah mengikut tetapan asal
        End If
        
    End If

    If Not IsNull(rs!pemalar_bonus_biasa) Then 'Kadar perolehan mata ganjaran (ahli biasa)
        G_PEMALAR_BONUS_BIASA = rs!pemalar_bonus_biasa
    Else
        G_PEMALAR_BONUS_BIASA = 0
    End If
    If Not IsNull(rs!pemalar_tebus_bonus_biasa) Then 'Kadar tebusan mata ganjaran (ahli biasa)
        G_PEMALAR_TEBUS_BIASA = rs!pemalar_tebus_bonus_biasa
    Else
        G_PEMALAR_TEBUS_BIASA = 0
    End If
    If Not IsNull(rs!pemalar_bonus_silver) Then 'Kadar perolehan mata ganjaran (silver)
        G_PEMALAR_BONUS_SILVER = rs!pemalar_bonus_silver
    Else
        G_PEMALAR_BONUS_SILVER = 0
    End If
    If Not IsNull(rs!pemalar_tebus_bonus_silver) Then 'Kadar tebusan mata ganjaran (silver)
        G_PEMALAR_TEBUS_SILVER = rs!pemalar_tebus_bonus_silver
    Else
        G_PEMALAR_TEBUS_SILVER = 0
    End If
    If Not IsNull(rs!pemalar_bonus_gold) Then 'Kadar perolehan mata ganjaran (gold)
        G_PEMALAR_BONUS_GOLD = rs!pemalar_bonus_gold
    Else
        G_PEMALAR_BONUS_GOLD = 0
    End If
    If Not IsNull(rs!pemalar_tebus_bonus_gold) Then 'Kadar tebusan mata ganjaran (gold)
        G_PEMALAR_TEBUS_GOLD = rs!pemalar_tebus_bonus_gold
    Else
        G_PEMALAR_TEBUS_GOLD = 0
    End If
    
    If Not IsNull(rs!pemalar_bonus_platinum) Then 'Kadar perolehan mata ganjaran (platinum)
        G_PEMALAR_BONUS_PLATINUM = rs!pemalar_bonus_platinum
    Else
        G_PEMALAR_BONUS_PLATINUM = 0
    End If
    If Not IsNull(rs!pemalar_tebus_bonus_platinum) Then 'Kadar tebusan mata ganjaran (platinum)
        G_PEMALAR_TEBUS_PLATINUM = rs!pemalar_tebus_bonus_platinum
    Else
        G_PEMALAR_TEBUS_PLATINUM = 0
    End If
    If Not IsNull(rs!invoice_tak_rasmi) Then '0 : Tidak Rasmi , 1 : Rasmi
        If rs!invoice_tak_rasmi = 0 Then
            G_INVOICE_RASMI = 0
        ElseIf rs!invoice_tak_rasmi = 1 Then
            G_INVOICE_RASMI = 1
        End If
    Else
        G_INVOICE_RASMI = 0
    End If
    If Not IsNull(rs!tray_x_loc) Then
        If IsNumeric(rs!tray_x_loc) Then
            G_TRAY_X = rs!tray_x_loc
        Else
            G_TRAY_X = 1800
        End If
    Else
        G_TRAY_X = 1800
    End If
    If Not IsNull(rs!dev_p) Then
        G_DEV_PASS_DEFAULT = rs!dev_p
    Else
        G_DEV_PASS_DEFAULT = "Sankyu1234567890-"
    End If
    If Not IsNull(rs!inv_type) Then
        G_INVOICE_TYPE = rs!inv_type
    Else
        G_INVOICE_TYPE = 0
    End If
    If Not IsNull(rs!rate_trade_in) Then
        G_TI_RATE_TI = rs!rate_trade_in
    Else
        G_TI_RATE_TI = 0
    End If
    If Not IsNull(rs!rate_buyback) Then
        G_TI_RATE_BB = rs!rate_buyback
    Else
        G_TI_RATE_BB = 0
    End If
    If Not IsNull(rs!rate_caj_pertukaran) Then
        G_TI_RATE_TUKAR = rs!rate_caj_pertukaran
    Else
        G_TI_RATE_TUKAR = 0
    End If
End If

rs.Close
Set rs = Nothing

Call setting_invoice
End Sub
Sub setting_barcode()
'On Error GoTo logging:
If MDI_frm1.L20_Text = "Semua cawangan" Then
    LM_KEDAI = "HQ"
Else
    LM_KEDAI = MDI_frm1.L20_Text
End If

'### Maklumat kedai ### - Start
LM_CONN = 1
re_conn_1:
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select nama_kedai_3,bar_header,header_a,footer_a from 56_maklumat_kedai where cawangan='" & LM_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai_3) Then G_LM_NAMA_KEDAI = rs!nama_kedai_3
    If Not IsNull(rs!bar_header) Then G_LM_HEAD = rs!bar_header
    If Not IsNull(rs!header_a) Then G_HEADER_A = rs!header_a
    If Not IsNull(rs!footer_a) Then G_FOOTER_A = rs!footer_a
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Erase G_SKU_LINE
Erase G_Frm56_LM_TYPE
Erase G_Frm56_LM_SIZE
Erase G_NAMA_BARCODE

LM_CONN = 2
re_conn_2:
Set rs1 = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs1.Open "select * from layout_barcode where perkara='" & G_JENIS_BARCODE_PRINTER & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs1.EOF Then
    If Not IsNull(rs1!a_data) Then G_SKU_DATA(0) = rs1!a_data
    If Not IsNull(rs1!a_pre_data) Then G_SKU_PRE_DATA(0) = rs1!a_pre_data
    If Not IsNull(rs1!a_line) Then G_SKU_LINE(0) = rs1!a_line
    If Not IsNull(rs1!a_font_size) Then G_SKU_SIZE(0) = rs1!a_font_size
    If Not IsNull(rs1!a_font_type) Then G_SKU_FONT(0) = rs1!a_font_type
    If Not IsNull(rs1!a_position_x) Then G_SKU_POS_X(0) = rs1!a_position_x
    If Not IsNull(rs1!a_position_y) Then G_SKU_POS_Y(0) = rs1!a_position_y
    If Not IsNull(rs1!a_bold) Then G_SKU_BOLD(0) = rs1!a_bold
    If Not IsNull(rs1!a_italic) Then G_SKU_ITALIC(0) = rs1!a_italic

    If Not IsNull(rs1!b_data) Then G_SKU_DATA(1) = rs1!b_data
    If Not IsNull(rs1!b_pre_data) Then G_SKU_PRE_DATA(1) = rs1!b_pre_data
    If Not IsNull(rs1!b_line) Then G_SKU_LINE(1) = rs1!b_line
    If Not IsNull(rs1!b_font_size) Then G_SKU_SIZE(1) = rs1!b_font_size
    If Not IsNull(rs1!b_font_type) Then G_SKU_FONT(1) = rs1!b_font_type
    If Not IsNull(rs1!b_position_x) Then G_SKU_POS_X(1) = rs1!b_position_x
    If Not IsNull(rs1!b_position_y) Then G_SKU_POS_Y(1) = rs1!b_position_y
    If Not IsNull(rs1!b_bold) Then G_SKU_BOLD(1) = rs1!b_bold
    If Not IsNull(rs1!b_italic) Then G_SKU_ITALIC(1) = rs1!b_italic

    If Not IsNull(rs1!c_data) Then G_SKU_DATA(2) = rs1!c_data
    If Not IsNull(rs1!c_pre_data) Then G_SKU_PRE_DATA(2) = rs1!c_pre_data
    If Not IsNull(rs1!c_line) Then G_SKU_LINE(2) = rs1!c_line
    If Not IsNull(rs1!c_font_size) Then G_SKU_SIZE(2) = rs1!c_font_size
    If Not IsNull(rs1!c_font_type) Then G_SKU_FONT(2) = rs1!c_font_type
    If Not IsNull(rs1!c_position_x) Then G_SKU_POS_X(2) = rs1!c_position_x
    If Not IsNull(rs1!c_position_y) Then G_SKU_POS_Y(2) = rs1!c_position_y
    If Not IsNull(rs1!c_bold) Then G_SKU_BOLD(2) = rs1!c_bold
    If Not IsNull(rs1!c_italic) Then G_SKU_ITALIC(2) = rs1!c_italic
    
    If Not IsNull(rs1!BARCODE_TYPE) Then
        If rs1!BARCODE_TYPE = 0 Then
            G_L_JENIS_BARCODE = 0
        ElseIf rs1!BARCODE_TYPE = 1 Then
            G_L_JENIS_BARCODE = 1
        ElseIf rs1!BARCODE_TYPE = 2 Then
            G_L_JENIS_BARCODE = 2
        End If
    End If
End If

rs1.Close
Set rs1 = Nothing

Exit Sub
logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Module1 : setting_barcode" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If LM_ERR_NO = "3704" Or LM_ERR_NO = "-2147467259" Or LM_ERR_NO = "-2147217887" Then
    Call Main
    If LM_CONN = 1 Then
        Resume re_conn_1:
    ElseIf LM_CONN = 2 Then
        Resume re_conn_2:
    End If
Else
    Resume Next
End If
End Sub
Sub setting_invoice()
'on error resume next

If MDI_frm1.L4_Text = "Semua cawangan" Then
    
    LM_KEDAI = "HQ"
    
Else
    
    LM_KEDAI = MDI_frm1.L20_Text
    
End If

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 92_setting_inv where cawangan='" & LM_KEDAI & "' AND terminal='" & G_TERMINAL & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!x_1) Then G_L1_LEFT = rs!x_1
    If Not IsNull(rs!y_1) Then G_L1_TOP = rs!y_1
    If Not IsNull(rs!bold_1) Then G_L1_BOLD = rs!bold_1
    If Not IsNull(rs!italic_1) Then G_L1_ITALIC = rs!italic_1
    If Not IsNull(rs!font_size_1) Then G_L1_FONT = rs!font_size_1
    If Not IsNull(rs!width_1) Then G_L1_WIDTH = rs!width_1
    If Not IsNull(rs!height_1) Then G_L1_HEIGHT = rs!height_1

    If Not IsNull(rs!x_2) Then G_L2_LEFT = rs!x_2
    If Not IsNull(rs!y_2) Then G_L2_TOP = rs!y_2
    If Not IsNull(rs!bold_2) Then G_L2_BOLD = rs!bold_2
    If Not IsNull(rs!italic_2) Then G_L2_ITALIC = rs!italic_2
    If Not IsNull(rs!font_size_2) Then G_L2_FONT = rs!font_size_2
    If Not IsNull(rs!width_2) Then G_L2_WIDTH = rs!width_2
    If Not IsNull(rs!height_2) Then G_L2_HEIGHT = rs!height_2

    If Not IsNull(rs!x_3) Then G_L3_LEFT = rs!x_3
    If Not IsNull(rs!y_3) Then G_L3_TOP = rs!y_3
    If Not IsNull(rs!bold_3) Then G_L3_BOLD = rs!bold_3
    If Not IsNull(rs!italic_3) Then G_L3_ITALIC = rs!italic_3
    If Not IsNull(rs!font_size_3) Then G_L3_FONT = rs!font_size_3
    If Not IsNull(rs!width_3) Then G_L3_WIDTH = rs!width_3
    If Not IsNull(rs!height_3) Then G_L3_HEIGHT = rs!height_3
    
    If Not IsNull(rs!x_4) Then G_L4_LEFT = rs!x_4
    If Not IsNull(rs!y_4) Then G_L4_TOP = rs!y_4
    If Not IsNull(rs!bold_4) Then G_L4_BOLD = rs!bold_4
    If Not IsNull(rs!italic_4) Then G_L4_ITALIC = rs!italic_4
    If Not IsNull(rs!font_size_4) Then G_L4_FONT = rs!font_size_4
    If Not IsNull(rs!width_4) Then G_L4_WIDTH = rs!width_4
    If Not IsNull(rs!height_4) Then G_L4_HEIGHT = rs!height_4
    
    If Not IsNull(rs!x_5) Then G_L5_LEFT = rs!x_5
    If Not IsNull(rs!y_5) Then G_L5_TOP = rs!y_5
    If Not IsNull(rs!bold_5) Then G_L5_BOLD = rs!bold_5
    If Not IsNull(rs!italic_5) Then G_L5_ITALIC = rs!italic_5
    If Not IsNull(rs!font_size_5) Then G_L5_FONT = rs!font_size_5
    If Not IsNull(rs!width_5) Then G_L5_WIDTH = rs!width_5
    If Not IsNull(rs!height_5) Then G_L5_HEIGHT = rs!height_5
    
    If Not IsNull(rs!x_6) Then G_L6_LEFT = rs!x_6
    If Not IsNull(rs!y_6) Then G_L6_TOP = rs!y_6
    If Not IsNull(rs!bold_6) Then G_L6_BOLD = rs!bold_6
    If Not IsNull(rs!italic_6) Then G_L6_ITALIC = rs!italic_6
    If Not IsNull(rs!font_size_6) Then G_L6_FONT = rs!font_size_6
    If Not IsNull(rs!width_6) Then G_L6_WIDTH = rs!width_6
    If Not IsNull(rs!height_6) Then G_L6_HEIGHT = rs!height_6
    
    If Not IsNull(rs!x_7) Then G_L7_LEFT = rs!x_7
    If Not IsNull(rs!y_7) Then G_L7_TOP = rs!y_7
    If Not IsNull(rs!bold_7) Then G_L7_BOLD = rs!bold_7
    If Not IsNull(rs!italic_7) Then G_L7_ITALIC = rs!italic_7
    If Not IsNull(rs!font_size_7) Then G_L7_FONT = rs!font_size_7
    If Not IsNull(rs!width_7) Then G_L7_WIDTH = rs!width_7
    If Not IsNull(rs!height_7) Then G_L7_HEIGHT = rs!height_7
    
    If Not IsNull(rs!x_8) Then G_L8_LEFT = rs!x_8
    If Not IsNull(rs!y_8) Then G_L8_TOP = rs!y_8
    If Not IsNull(rs!bold_8) Then G_L8_BOLD = rs!bold_8
    If Not IsNull(rs!italic_8) Then G_L8_ITALIC = rs!italic_8
    If Not IsNull(rs!font_size_8) Then G_L8_FONT = rs!font_size_8
    If Not IsNull(rs!width_8) Then G_L8_WIDTH = rs!width_8
    If Not IsNull(rs!height_8) Then G_L8_HEIGHT = rs!height_8
    
    If Not IsNull(rs!x_9) Then G_L9_LEFT = rs!x_9
    If Not IsNull(rs!y_9) Then G_L9_TOP = rs!y_9
    If Not IsNull(rs!bold_9) Then G_L9_BOLD = rs!bold_9
    If Not IsNull(rs!italic_9) Then G_L9_ITALIC = rs!italic_9
    If Not IsNull(rs!font_size_9) Then G_L9_FONT = rs!font_size_9
    If Not IsNull(rs!width_9) Then G_L9_WIDTH = rs!width_9
    If Not IsNull(rs!height_9) Then G_L9_HEIGHT = rs!height_9
    
    If Not IsNull(rs!x_10) Then G_L10_LEFT = rs!x_10
    If Not IsNull(rs!y_10) Then G_L10_TOP = rs!y_10
    If Not IsNull(rs!bold_10) Then G_L10_BOLD = rs!bold_10
    If Not IsNull(rs!italic_10) Then G_L10_ITALIC = rs!italic_10
    If Not IsNull(rs!font_size_10) Then G_L10_FONT = rs!font_size_10
    If Not IsNull(rs!width_10) Then G_L10_WIDTH = rs!width_10
    If Not IsNull(rs!height_10) Then G_L10_HEIGHT = rs!height_10
    
    If Not IsNull(rs!x_11) Then G_L11_LEFT = rs!x_11
    If Not IsNull(rs!y_11) Then G_L11_TOP = rs!y_11
    If Not IsNull(rs!bold_11) Then G_L11_BOLD = rs!bold_11
    If Not IsNull(rs!italic_11) Then G_L11_ITALIC = rs!italic_11
    If Not IsNull(rs!font_size_11) Then G_L11_FONT = rs!font_size_11
    If Not IsNull(rs!width_11) Then G_L11_WIDTH = rs!width_11
    If Not IsNull(rs!height_11) Then G_L11_HEIGHT = rs!height_11
    
    If Not IsNull(rs!x_12) Then G_L12_LEFT = rs!x_12
    If Not IsNull(rs!y_12) Then G_L12_TOP = rs!y_12
    If Not IsNull(rs!bold_12) Then G_L12_BOLD = rs!bold_12
    If Not IsNull(rs!italic_12) Then G_L12_ITALIC = rs!italic_12
    If Not IsNull(rs!font_size_12) Then G_L12_FONT = rs!font_size_12
    If Not IsNull(rs!width_12) Then G_L12_WIDTH = rs!width_12
    If Not IsNull(rs!height_12) Then G_L12_HEIGHT = rs!height_12
    
    If Not IsNull(rs!x_13) Then G_L13_LEFT = rs!x_13
    If Not IsNull(rs!y_13) Then G_L13_TOP = rs!y_13
    If Not IsNull(rs!bold_13) Then G_L13_BOLD = rs!bold_13
    If Not IsNull(rs!italic_13) Then G_L13_ITALIC = rs!italic_13
    If Not IsNull(rs!font_size_13) Then G_L13_FONT = rs!font_size_13
    If Not IsNull(rs!width_13) Then G_L13_WIDTH = rs!width_13
    If Not IsNull(rs!height_13) Then G_L13_HEIGHT = rs!height_13
    
    If Not IsNull(rs!x_14) Then G_L14_LEFT = rs!x_14
    If Not IsNull(rs!y_14) Then G_L14_TOP = rs!y_14
    If Not IsNull(rs!bold_14) Then G_L14_BOLD = rs!bold_14
    If Not IsNull(rs!italic_14) Then G_L14_ITALIC = rs!italic_14
    If Not IsNull(rs!font_size_14) Then G_L14_FONT = rs!font_size_14
    If Not IsNull(rs!width_14) Then G_L14_WIDTH = rs!width_14
    If Not IsNull(rs!height_14) Then G_L14_HEIGHT = rs!height_14
    
    If Not IsNull(rs!x_15) Then G_L15_LEFT = rs!x_15
    If Not IsNull(rs!y_15) Then G_L15_TOP = rs!y_15
    If Not IsNull(rs!bold_15) Then G_L15_BOLD = rs!bold_15
    If Not IsNull(rs!italic_15) Then G_L15_ITALIC = rs!italic_15
    If Not IsNull(rs!font_size_15) Then G_L15_FONT = rs!font_size_15
    If Not IsNull(rs!width_15) Then G_L15_WIDTH = rs!width_15
    If Not IsNull(rs!height_15) Then G_L15_HEIGHT = rs!height_15
    
    If Not IsNull(rs!x_16) Then G_L16_LEFT = rs!x_16
    If Not IsNull(rs!y_16) Then G_L16_TOP = rs!y_16
    If Not IsNull(rs!bold_16) Then G_L16_BOLD = rs!bold_16
    If Not IsNull(rs!italic_16) Then G_L16_ITALIC = rs!italic_16
    If Not IsNull(rs!font_size_16) Then G_L16_FONT = rs!font_size_16
    If Not IsNull(rs!width_16) Then G_L16_WIDTH = rs!width_16
    If Not IsNull(rs!height_16) Then G_L16_HEIGHT = rs!height_16
    
    If Not IsNull(rs!x_17) Then G_L17_LEFT = rs!x_17
    If Not IsNull(rs!y_17) Then G_L17_TOP = rs!y_17
    If Not IsNull(rs!bold_17) Then G_L17_BOLD = rs!bold_17
    If Not IsNull(rs!italic_17) Then G_L17_ITALIC = rs!italic_17
    If Not IsNull(rs!font_size_17) Then G_L17_FONT = rs!font_size_17
    If Not IsNull(rs!width_17) Then G_L17_WIDTH = rs!width_17
    If Not IsNull(rs!height_17) Then G_L17_HEIGHT = rs!height_17
    
    If Not IsNull(rs!x_18) Then G_L18_LEFT = rs!x_18
    If Not IsNull(rs!y_18) Then G_L18_TOP = rs!y_18
    If Not IsNull(rs!bold_18) Then G_L18_BOLD = rs!bold_18
    If Not IsNull(rs!italic_18) Then G_L18_ITALIC = rs!italic_18
    If Not IsNull(rs!font_size_18) Then G_L18_FONT = rs!font_size_18
    If Not IsNull(rs!width_18) Then G_L18_WIDTH = rs!width_18
    If Not IsNull(rs!height_18) Then G_L18_HEIGHT = rs!height_18
End If

rs.Close
Set rs = Nothing

End Sub
Sub terminal_memory()
'On Error resume next
G_JENIS_BARCODE_PRINTER = 0
DATA_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 91_senarai_terminal where terminal='" & G_TERMINAL & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!jenis_barcode_printer) Then G_JENIS_BARCODE_PRINTER = rs!jenis_barcode_printer
    DATA_FOUND = 1
End If

rs.Close
Set rs = Nothing

If DATA_FOUND = 1 Then
    Call setting_barcode
Else
    MsgBox "Maklumat Terminal Ini Tidak Dijumpai.", vbCritical, "Error"
    End
End If
End Sub


