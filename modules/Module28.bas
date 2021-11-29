Attribute VB_Name = "Module28"
Global GB_RESIT
Global GB_VOUCHER
Public G_MODE 'YES : Ada kad keahlian , NO : Tiada kad keahlian
Public G_MIN_LEN As String 'Panjang minimum
Public G_MAX_LEN As String 'Panjang maksimum
Public G_CODE 'Kod kedai
Sub Frm68_Reset_All()
'on error resume next
Frm68.TB1 = vbNullString
Frm68.TB2 = vbNullString
Frm68.TB3 = vbNullString
Frm68.TB4 = vbNullString
Frm68.TB5 = vbNullString
Frm68.TB6 = vbNullString
Frm68.TB7 = vbNullString
Frm68.TB8 = vbNullString
Frm68.TB9 = vbNullString
Frm68.TB10 = vbNullString
Frm68.TB11 = vbNullString
Frm68.TB12 = vbNullString
Frm68.TB14 = vbNullString
'Frm68.L11_Text.Visible = False
'Frm68.TB12.Visible = False

Frm68.TB2.Locked = False
Frm68.TB2.BackColor = &H80000005

Frm68.TB17 = "0.00"
Frm68.TB18 = vbNullString
Frm68.TB19 = "0.00"

'Frm68.CB3 = 0
'Frm68.CB4 = 0
'Frm68.CB5 = 0
'Frm68.CB6 = 0
'Frm68.CB7 = 0
'Frm68.CB8 = 0
Frm68.CB9 = 0
Frm68.CB10 = 0
Frm68.CB11 = 0
Frm68.CB12 = 0
Frm68.CB14 = 0
Frm68.CB15 = 0
Frm68.CB16 = 0
Frm68.CB17 = 0
Frm68.CB19 = 1
Frm68.CB20 = 0

Frm68.CB13.Enabled = True
Frm68.CB13 = 0

user_level = MDI_frm1.L4_Text

If user_level = "Administration" Or user_level = "Guest/User" Then

    Frm68.CB13.Visible = False
    Frm68.Label8.Visible = False
    
Else
    
    If G_GST_SYSTEM = "YES" Then
    '    Frm68.CB13.Visible = True
    '    Frm68.Label8.Visible = True
    
        If G_INVOICE_RASMI = 0 Then
            Frm68.CB13 = 1
        Else
            Frm68.CB13 = 0
        End If
        
    Else
        Frm68.CB13.Visible = False
        Frm68.Label8.Visible = False
    End If
End If

Frm68.Frame1.Left = 120
Frm68.Frame1.Top = 240
'Frm68.Pic2.Left = 120
'Frm68.Pic2.Top = 2400
Frm68.Frame10.Left = 120
Frm68.Frame10.Top = 240
Frm68.Frame9.Left = 3120
Frm68.Frame9.Top = 240
Frm68.Frame5.Left = 120
Frm68.Frame5.Top = 240
'Frm68.Pic7.Left = 120
'Frm68.Pic7.Top = 2400
Frm68.Pic8.Left = 120
Frm68.Pic8.Top = 240
Frm68.Pic9.Left = 120
Frm68.Pic9.Top = 240
Frm68.Frame4.Left = 120
Frm68.Frame4.Top = 240
Frm68.Frame7.Left = 120
Frm68.Frame7.Top = 240
Frm68.Frame6.Left = 120
Frm68.Frame6.Top = 240
Frm68.Frame2.Left = 120
Frm68.Frame2.Top = 240
Frm68.Frame8.Left = 120
Frm68.Frame8.Top = 240

Frm68.L16_Text = vbNullString
Frm68.L17_Text = vbNullString
Frm68.L18_Text = vbNullString
Frm68.L19_Text = vbNullString
Frm68.L26_Text = vbNullString
Frm68.L29_Text = vbNullString
Frm68.L30_Text = vbNullString
Frm68.L31_Text = vbNullString
Frm68.L32_Text = vbNullString
Frm68.L5_Text = vbNullString
Frm68.L42_Text = vbNullString

Frm68.L43_Text = vbNullString
Frm68.L44_Text = vbNullString
Frm68.L45_Text = vbNullString
Frm68.L46_Text = vbNullString
Frm68.L47_Text = vbNullString
Frm68.L48_Text = vbNullString
Frm68.L49_Text = vbNullString
Frm68.L55_Text = vbNullString
Frm68.L57_Text = 0
Frm68.L66_Text = 0

'Frm68.L39_Text = vbNullString
'Frm68.L40_Text = vbNullString

Frm68.L33_Text = "0.00"
Frm68.L34_Text = "0.00"
Frm68.L35_Text = "0.00"

Frm68.L26_Text.BackStyle = 0
Frm68.L27_Text.BackStyle = 0
Frm68.L10_Text.BackStyle = 0
Frm68.L14_Text.BackStyle = 0
Frm68.DTPicker1 = DateTime.Date
Frm68.DTPicker2 = DateTime.Date
Frm68.DTPicker3 = DateTime.Date
Frm68.DTPicker4 = DateTime.Date
Frm68.DTPicker5 = DateTime.Date
Frm68.DTPicker6 = DateTime.Date
Frm68.DTPicker7 = DateTime.Date
Frm68.DTPicker8 = DateTime.Date

Frm68.CMD1.Visible = False
Frm68.CMD2.Visible = False
Frm68.CMD29.Visible = False
Frm68.Frame1.Visible = False
'Frm68.Pic2.Visible = False
Frm68.Frame10.Visible = False
Frm68.Frame9.Visible = False
Frm68.Frame5.Visible = False
'Frm68.Pic7.Visible = False
Frm68.Pic8.Visible = False
Frm68.Pic9.Visible = False
Frm68.Frame4.Visible = False
Frm68.Frame7.Visible = False
Frm68.Frame6.Visible = False
Frm68.Frame2.Visible = False
Frm68.Frame8.Visible = False

Frm68.CMD14.Visible = True
Frm68.CMD15.Visible = True
Frm68.CMD16.Visible = False
Frm68.CMD17.Visible = False

Frm68.MSFlexGrid6.Visible = False
Frm68.MSFlexGrid7.Visible = False
Frm68.MSFlexGrid8.Visible = False
Frm68.MSFlexGrid9.Visible = False
Frm68.MSFlexGrid10.Visible = False
Frm68.L64_Text.Visible = False

Frm68.MSFlexGrid6.Left = 5280
Frm68.MSFlexGrid6.Top = 720
Frm68.MSFlexGrid7.Left = 5280
Frm68.MSFlexGrid7.Top = 720
Frm68.MSFlexGrid8.Left = 5280
Frm68.MSFlexGrid8.Top = 720
Frm68.MSFlexGrid9.Left = 5280
Frm68.MSFlexGrid9.Top = 720
Frm68.MSFlexGrid10.Left = 5280
Frm68.MSFlexGrid10.Top = 720

'Frm68.CB9.Enabled = True
'Frm68.CB10.Enabled = True
'Frm68.CB11.Enabled = True
'Frm68.CB12.Enabled = True

Frm68.CBB1.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where status='" & "Aktif" & "' AND InvestorSmall = 0 AND InvestorBig = 0", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    
    If rs!user_level = 1 Or rs!user_level = 2 Or rs!user_level = 3 Or rs!user_level = G_LEVEL_USER Then Frm68.CBB1.AddItem rs!Samaran & "  |  " & rs!NoPekerja
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Default1 = "Default" Then
        If Not IsNull(rs!no_resit_simpanan) Then Frm68.L26_Text = rs!no_resit_simpanan 'No. Resit Simpanan
        If Not IsNull(rs!no_customer) And Not IsNull(rs!kod_customer) Then
            Frm68.TB12 = rs!kod_customer & Format(rs!no_customer, "0000") 'No. Pelanggan
        Else
            MsgBox "Tiada maklumat tentang Kod Syarikat atau No. Agen Dropship dijumpai." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Sila hubungi Sankyu System untuk langkah seterusnya.", vbExclamation, "Error"
        End If
        'If Not IsNull(rs!membership_flag) Then
        '    If rs!membership_flag = 0 Then
        '        Frm68.CB16 = 1
        '    ElseIf rs!membership_flag = 1 Then
        '        Frm68.CB15 = 1
        '    End If
        'Else
        '    Frm68.CB16 = 1
        'End If
    End If
End If

rs.Close
Set rs = Nothing

If G_MODE = "NO" Then
    Frm68.CB18 = 1
    Frm68.CB15 = 0
ElseIf G_MODE = "YES" Then
    Frm68.CB18 = 0
    Frm68.CB15 = 1
End If
End Sub
Sub Frm68_hide_report()
'on error resume next
Frm68.MSFlexGrid6.Visible = False
Frm68.MSFlexGrid7.Visible = False
Frm68.MSFlexGrid8.Visible = False
Frm68.MSFlexGrid9.Visible = False
Frm68.MSFlexGrid10.Visible = False
End Sub
Sub frm68_senarai_pelanggan_header()
'on error resume next
With Frm68.LV1
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear

    Frm68.LV1.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 600, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "Kategori", 1700
    .ColumnHeaders.Add 5, , "No. Pelanggan", 1500
    .ColumnHeaders.Add 6, , "Nama", 7000
    .ColumnHeaders.Add 7, , "No. Kad Pengenalan", 2000
    .ColumnHeaders.Add 8, , "No. Telefon", 1800
    .ColumnHeaders.Add 9, , "E-mail", 3500
    .ColumnHeaders.Add 10, , "Jumlah Simpanan (RM)", 2500, 1
    .ColumnHeaders.Add 11, , "Agen Dropship", 1700, 2
    .ColumnHeaders.Add 12, , "Kad Ahli", 1500, 2
    .ColumnHeaders.Add 13, , "Mata Terkumpul", 1700, 2
    .ColumnHeaders.Add 14, , "Status", 1700

'No.
'No.
'No. ID
'Kategori
'No. Pelanggan
'Nama
'No. Kad Pengenalan
'No. Telefon
'E-mail
'Jumlah Simpanan (RM)
'Agen Dropship
'Kad Ahli
'Mata Terkumpul
'Status

End With
End Sub
Sub frm68_senarai_pelanggan()
'on error resume next
Dim Frm68_LM_TOTAL_PAGE As Double
Dim Frm68_LM_FIELD As String
'Dim Frm68_LM_SEARCH_1 As String
Dim Frm68_LM_SEARCH_1_LOGIC As String

Frm68_PAGE_SIZE = 37
Frm68_LM_TOTAL_PAGE = 0
x = 0
LM_BILANGAN_AHLI = 0

If Frm68.L39_Text = "Semua senarai" Then
    Frm68_LM_SEARCH_1 = Null
    Frm68_LM_SEARCH_1_LOGIC = "<>"
    Frm68_LM_FIELD = "kategori_pelanggan"
    
    Frm68.L71_Text = "Senarai semua pelanggan."
End If
If Frm68.L39_Text = "Semua pelanggan biasa" Then
    Frm68_LM_SEARCH_1 = "1"
    Frm68_LM_SEARCH_1_LOGIC = "="
    Frm68_LM_FIELD = "kategori_pelanggan"
    
    Frm68.L71_Text = "Senarai semua pelanggan biasa sahaja."
End If
If Frm68.L39_Text = "Semua ahli biasa" Then
    Frm68_LM_SEARCH_1 = "2"
    Frm68_LM_SEARCH_1_LOGIC = "="
    Frm68_LM_FIELD = "kategori_pelanggan"
    
    Frm68.L71_Text = "Senarai semua ahli biasa sahaja."
End If
If Frm68.L39_Text = "Semua silver" Then
    Frm68_LM_SEARCH_1 = "3"
    Frm68_LM_SEARCH_1_LOGIC = "="
    Frm68_LM_FIELD = "kategori_pelanggan"
    
    Frm68.L71_Text = "Senarai semua silver sahaja."
End If
If Frm68.L39_Text = "Semua gold" Then
    Frm68_LM_SEARCH_1 = "4"
    Frm68_LM_SEARCH_1_LOGIC = "="
    Frm68_LM_FIELD = "kategori_pelanggan"
    
    Frm68.L71_Text = "Senarai semua gold sahaja."
End If
If Frm68.L39_Text = "Semua platinum" Then
    Frm68_LM_SEARCH_1 = "5"
    Frm68_LM_SEARCH_1_LOGIC = "="
    Frm68_LM_FIELD = "kategori_pelanggan"
    
    Frm68.L71_Text = "Senarai semua platinum sahaja."
End If
If Frm68.L39_Text = "Semua agen dropship" Then
    Frm68_LM_SEARCH_1 = "1"
    Frm68_LM_SEARCH_1_LOGIC = "="
    Frm68_LM_FIELD = "dropship"
    
    Frm68.L71_Text = "Senarai semua agen dropship sahaja."
End If
If Frm68.L39_Text = "Nama" Then
    Frm68_LM_SEARCH_1 = UCase(Frm68.L40_Text)
    Frm68_LM_SEARCH_1_LOGIC = "="
    Frm68_LM_FIELD = "nama"
    
    Frm68.L71_Text = "Senarai ahli dengan nama " & UCase(Frm68.L40_Text) & "."
End If
If Frm68.L39_Text = "No. kad pengenalan" Then
    Frm68_LM_SEARCH_1 = UCase(Frm68.L40_Text)
    Frm68_LM_SEARCH_1_LOGIC = "="
    Frm68_LM_FIELD = "no_ic"
    
    Frm68.L71_Text = "Senarai ahli dengan no kad pengenalan " & UCase(Frm68.L40_Text) & "."
End If
If Frm68.L39_Text = "No. keahlian" Then
    Frm68_LM_SEARCH_1 = UCase(Frm68.L40_Text)
    Frm68_LM_SEARCH_1_LOGIC = "="
    Frm68_LM_FIELD = "no_pelanggan"
    
    Frm68.L71_Text = "Senarai ahli dengan no keahlian " & UCase(Frm68.L40_Text) & "."
End If
If Frm68.L39_Text = "no_tel_hp" Then
    Frm68_LM_SEARCH_1 = UCase(Frm68.L40_Text)
    Frm68_LM_SEARCH_1_LOGIC = "="
    Frm68_LM_FIELD = "no_tel"
    
    Frm68.L71_Text = "Senarai ahli dengan no telefon " & UCase(Frm68.L40_Text) & "."
End If

LM_START_ROW = Frm68.L69_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm68_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If Frm68.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - Frm68_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        Frm68.L67_Text = 1
    End If
End If

Frm68_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from senarai_pelanggan where (status = 0 OR status = 1 OR status = 2) AND " & Frm68_LM_FIELD & " " & Frm68_LM_SEARCH_1_LOGIC & " '" & Frm68_LM_SEARCH_1 & "' order by nama ASC LIMIT " & LM_START_ROW & "," & Frm68_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    If Frm68_LM_PAGE_FOUND = 0 Then
        If Frm68.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                Frm68.L67_Text = Frm68.L67_Text + 1 'Paparan Page ke-xxx
                Frm68_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(Frm68.L67_Text) Then
                    If Frm68.L67_Text <> 1 Then
                        Frm68.L67_Text = Frm68.L67_Text - 1 'Paparan Page ke-xxx
                        Frm68_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If
    Y = ((Frm68.L67_Text - 1) * Frm68_PAGE_SIZE) + x
    
    With Frm68.LV1.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
   
        If Not IsNull(rs!kategori_pelanggan) Then 'Kategori
        
            If rs!kategori_pelanggan = 1 Then .ListSubItems.Add , , "Pelanggan Biasa"
            If rs!kategori_pelanggan = 2 Then .ListSubItems.Add , , "Ahli Biasa"
            If rs!kategori_pelanggan = 3 Then .ListSubItems.Add , , "Silver"
            If rs!kategori_pelanggan = 4 Then .ListSubItems.Add , , "Gold"
            If rs!kategori_pelanggan = 5 Then .ListSubItems.Add , , "Platinum"
        
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_pelanggan) Then 'No. Pelanggan
            .ListSubItems.Add , , rs!no_pelanggan
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Nama) Then 'Nama
            .ListSubItems.Add , , rs!Nama
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_ic) Then 'No. Kad Pengenalan
            .ListSubItems.Add , , rs!no_ic
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!no_tel) Then 'No. Telefon
            .ListSubItems.Add , , rs!no_tel
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Email) Then 'E-mail
            .ListSubItems.Add , , rs!Email
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!baki_simpanan) Then 'Jumlah Simpanan (RM)
            .ListSubItems.Add , , Format(rs!baki_simpanan, "#,##0.00")
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!dropship) Then 'Agen Dropship

            If rs!dropship = 0 Then
                .ListSubItems.Add , , "Tidak"
            ElseIf rs!dropship = 1 Then
                .ListSubItems.Add , , "Ya"
            End If
        
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!membership_card) Then 'Kad Ahli

            If rs!membership_card = 0 Then
                .ListSubItems.Add , , "Tidak"
            ElseIf rs!membership_card = 1 Then
                .ListSubItems.Add , , "Ya"
            End If
        
        Else
            .ListSubItems.Add , , ""
        End If

        If Not IsNull(rs!baki_point) Then 'Mata Terkumpul
            .ListSubItems.Add , , Format(rs!baki_point, "#,##0")
        Else
            .ListSubItems.Add , , "0"
        End If
        
        If Not IsNull(rs!Status) Then 'Status
            If rs!Status = 1 Then
                .ListSubItems.Add , , "Aktif"  'Status
            ElseIf rs!Status = 0 Then
                .ListSubItems.Add , , "Tidak Aktif"  'Status
            End If
        Else
            .ListSubItems.Add , , "Tidak Aktif"  'Status
        End If

    End With

    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'### Jumlah Data ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from senarai_pelanggan where (status = 0 OR status = 1 OR status = 2) AND " & Frm68_LM_FIELD & " " & Frm68_LM_SEARCH_1_LOGIC & " '" & Frm68_LM_SEARCH_1 & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    LM_BILANGAN_AHLI = rs(0)
    Frm68_LM_TOTAL_PAGE = Format(rs(0) / Frm68_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, Frm68_LM_TOTAL_PAGE, ".") <> 0 Then
    
        Frm68_LM_PAGE = Split(Frm68_LM_TOTAL_PAGE, ".")(0)
        Frm68_LM_PAGE_LEBIHAN = Split(Frm68_LM_TOTAL_PAGE, ".")(1)
        
        If Frm68_LM_PAGE_LEBIHAN <> "00" Then
            Frm68.L68_Text = Frm68_LM_PAGE + 1
        Else
            Frm68.L68_Text = Frm68_LM_PAGE
        End If
        
    Else
    
        Frm68.L68_Text = Frm68_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        Frm68.L68_Text = 0
    End If
Else
    Frm68.L68_Text = 0
End If

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from senarai_pelanggan where status = 1 AND " & Frm68_LM_FIELD & " " & Frm68_LM_SEARCH_1_LOGIC & " '" & Frm68_LM_SEARCH_1 & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm68.L72_Text = rs(0)

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from senarai_pelanggan where status = 0 AND " & Frm68_LM_FIELD & " " & Frm68_LM_SEARCH_1_LOGIC & " '" & Frm68_LM_SEARCH_1 & "'", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm68.L73_Text = rs(0)

rs.Close
Set rs = Nothing

If Frm68.L68_Text = vbNullString Then
    Frm68.L68_Text = 0
End If
'### Jumlah Data ### - End

If x <> 0 Then
    Frm68.L69_Text = LM_START_ROW 'Titik Pencarian Data
    
    Frm68.Frame5.Visible = True
    Frm68.Frame4.Visible = False
Else
    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
    'If Frm68_LM_JUMLAH = 0 Then MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm68.L70_Text = 0 '0 : Bukan page terakhir , 1 : Page Terakhir
Else
    Frm68.L70_Text = 1 '0 : Bukan page terakhir , 1 : Page Terakhir
End If

Frm68.L57_Text = LM_BILANGAN_AHLI
End Sub
Sub Frm68_ListSimpanan_Header()
'on error resume next
Frm68.MSFlexGrid4.Clear
Frm68.MSFlexGrid4.RowHeight(0) = 600
Frm68.MSFlexGrid4.FormatString = "No.|<No.|<No. ID|<No. Rujukan|<Tarikh|<Jumlah (RM)"

Frm68.MSFlexGrid4.Rows = 1
Frm68.MSFlexGrid4.ColWidth(0) = 600
Frm68.MSFlexGrid4.ColWidth(1) = 0
Frm68.MSFlexGrid4.ColWidth(2) = 0 'No. ID
Frm68.MSFlexGrid4.ColWidth(3) = 1800 'No. Rujukan
Frm68.MSFlexGrid4.ColWidth(4) = 1800 'Tarikh
Frm68.MSFlexGrid4.ColWidth(5) = 1900 'Jumlah (RM)
Frm68.MSFlexGrid4.ColAlignment(5) = 7
End Sub
Sub Frm68_List_Simpanan()
'on error resume next
Dim Frm68_LM_JUMLAH_SIMPANAN As Double

Frm68_LM_JUMLAH_SIMPANAN = 0
x = 0
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 24_rekod_kewangan_pelanggan where no_rujukan_pelanggan='" & Frm68.L32_Text & "' AND jenis='" & "0" & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm68.MSFlexGrid4.Rows = x + 1
    Frm68.MSFlexGrid4.TextMatrix(x, 0) = x
    Frm68.MSFlexGrid4.TextMatrix(x, 1) = x
    If Not IsNull(rs!ID) Then Frm68.MSFlexGrid4.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!no_resit) Then Frm68.MSFlexGrid4.TextMatrix(x, 3) = rs!no_resit 'No. Rujukan
    If Not IsNull(rs!tarikh) Then Frm68.MSFlexGrid4.TextMatrix(x, 4) = rs!tarikh 'Tarikh
    If Not IsNull(rs!jumlah) Then
        Frm68.MSFlexGrid4.TextMatrix(x, 5) = rs!jumlah 'Jumlah Simpanan (RM)
        If IsNumeric(rs!jumlah) Then Frm68_LM_JUMLAH_SIMPANAN = Frm68_LM_JUMLAH_SIMPANAN + rs!jumlah 'Jumlah Simpanan (RM)
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm68.L33_Text = Format(Frm68_LM_JUMLAH_SIMPANAN, "0.00") 'Jumlah Simpanan (RM)
End Sub
Sub Frm68_ListPenggunaan_Header()
'on error resume next
Frm68.MSFlexGrid5.Clear
Frm68.MSFlexGrid5.RowHeight(0) = 600
Frm68.MSFlexGrid5.FormatString = "No.|<No.|<No. ID|<No. Resit|<Tarikh|<Tujuan Penggunaan|<Jumlah (RM)"

Frm68.MSFlexGrid5.Rows = 1
Frm68.MSFlexGrid5.ColWidth(0) = 600
Frm68.MSFlexGrid5.ColWidth(1) = 0
Frm68.MSFlexGrid5.ColWidth(2) = 0 'No. ID
Frm68.MSFlexGrid5.ColWidth(3) = 1800 'No. Rujukan
Frm68.MSFlexGrid5.ColWidth(4) = 1800 'Tarikh
Frm68.MSFlexGrid5.ColWidth(5) = 5000 'Tujuan Penggunaan
Frm68.MSFlexGrid5.ColWidth(6) = 2000 'Jumlah (RM)
End Sub
Sub Frm68_List_Penggunaan()
'on error resume next
Dim Frm68_LM_JUMLAH_PENGGUNAAN As Double

Frm68_LM_JUMLAH_PENGGUNAAN = 0
x = 0
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 24_rekod_kewangan_pelanggan where no_rujukan_pelanggan='" & Frm68.L32_Text & "' AND jenis='" & "1" & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm68.MSFlexGrid5.Rows = x + 1
    Frm68.MSFlexGrid5.TextMatrix(x, 0) = x
    Frm68.MSFlexGrid5.TextMatrix(x, 1) = x
    If Not IsNull(rs!ID) Then Frm68.MSFlexGrid5.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!no_resit) Then Frm68.MSFlexGrid5.TextMatrix(x, 3) = rs!no_resit 'No. Rujukan
    If Not IsNull(rs!tarikh) Then Frm68.MSFlexGrid5.TextMatrix(x, 4) = rs!tarikh 'Tarikh
    If Not IsNull(rs!jenis_penggunaan) Then
        If rs!jenis_penggunaan = 0 Then
            Frm68.MSFlexGrid5.TextMatrix(x, 5) = "Belian Barangan Kemas" 'Tujuaan Penggunaan
        ElseIf rs!jenis_penggunaan = 1 Then
            Frm68.MSFlexGrid5.TextMatrix(x, 5) = "Bayaran Ansuran Emas" 'Tujuaan Penggunaan
        ElseIf rs!jenis_penggunaan = 2 Then
            Frm68.MSFlexGrid5.TextMatrix(x, 5) = "Bayaran Deposit Tempahan Emas" 'Tujuaan Penggunaan
        ElseIf rs!jenis_penggunaan = 3 Then
            Frm68.MSFlexGrid5.TextMatrix(x, 5) = "Bayaran Servis" 'Tujuaan Penggunaan
        ElseIf rs!jenis_penggunaan = 4 Then
            Frm68.MSFlexGrid5.TextMatrix(x, 5) = "Bayaran Ambilan Tempahan Emas" 'Tujuaan Penggunaan
        End If
    End If
    If Not IsNull(rs!jumlah) Then
        Frm68.MSFlexGrid5.TextMatrix(x, 6) = Format(rs!jumlah, "0.00") 'Jumlah Penggunaan (RM)
        If IsNumeric(rs!jumlah) Then Frm68_LM_JUMLAH_PENGGUNAAN = Frm68_LM_JUMLAH_PENGGUNAAN + rs!jumlah 'Jumlah Penggunaan (RM)
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm68.L34_Text = Format(Frm68_LM_JUMLAH_PENGGUNAAN, "0.00") 'Jumlah Penggunaan (RM)
End Sub
Sub Frm68_SenaraiKomisyen_Header()
'on error resume next
Frm68.MSFlexGrid2.Clear
Frm68.MSFlexGrid2.RowHeight(0) = 900
Frm68.MSFlexGrid2.FormatString = "No.|<No.|<Tarikh|<No. Invoice|<No. Siri Produk|<Nama Item|<Berat (g)|<Komisyen Per Gram (RM/g)|<Upah (RM)|<Jumlah (RM)"

Frm68.MSFlexGrid2.Rows = 1
Frm68.MSFlexGrid2.ColWidth(0) = 600
Frm68.MSFlexGrid2.ColWidth(1) = 0
Frm68.MSFlexGrid2.ColWidth(2) = 1500 'Tarikh
Frm68.MSFlexGrid2.ColWidth(3) = 1700 'No. Invoice
Frm68.MSFlexGrid2.ColWidth(4) = 1500 'No. Siri Produk
Frm68.MSFlexGrid2.ColWidth(5) = 4300 'Nama Item
Frm68.MSFlexGrid2.ColWidth(6) = 1200 'Berat (g)
Frm68.MSFlexGrid2.ColAlignment(6) = 7
Frm68.MSFlexGrid2.ColWidth(7) = 1200 'Komisyen Per Gram (RM/g)
Frm68.MSFlexGrid2.ColAlignment(7) = 7
Frm68.MSFlexGrid2.ColWidth(8) = 1200 'Upah (RM)
Frm68.MSFlexGrid2.ColAlignment(8) = 7
Frm68.MSFlexGrid2.ColWidth(9) = 1200 'Jumlah (RM)
Frm68.MSFlexGrid2.ColAlignment(9) = 7
End Sub
Sub Frm68_SenaraiKomisyen_page()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim Frm68_JUMLAH_KOMISEN As Double

TM = Frm68.L6_Text
TA = Frm68.L7_Text

x = 0
Frm68_PAGE_SIZE = 28
Frm68_JUMLAH_KOMISEN = 0

LM_START_ROW = Frm68.L56_Text

If GM_NEXT_PREV = 0 Then
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm68_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        LM_START_ROW = LM_START_ROW - Frm68_PAGE_SIZE
    End If
End If

Frm68.L4_Text = "Rekod komisyen dari " & TM & " hingga " & TA

'###Senarai komisyen bagi Agen / Staff###
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 23_senarai_jualan where  status_rekod = 1 AND no_rujukan_agen_dropship='" & Frm68.L5_Text & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm68_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm68.MSFlexGrid2.Rows = x + 1
    Frm68.MSFlexGrid2.TextMatrix(x, 0) = x
    Frm68.MSFlexGrid2.TextMatrix(x, 1) = x
    If Not IsNull(rs!tarikh) Then Frm68.MSFlexGrid2.TextMatrix(x, 2) = rs!tarikh 'Tarikh
    If Not IsNull(rs!no_resit) Then Frm68.MSFlexGrid2.TextMatrix(x, 3) = rs!no_resit 'No Invoice
    If Not IsNull(rs!no_siri_Produk) Then Frm68.MSFlexGrid2.TextMatrix(x, 4) = rs!no_siri_Produk 'No Siri Produk
    If Not IsNull(rs!kategori_Produk) Then Frm68.MSFlexGrid2.TextMatrix(x, 5) = rs!kategori_Produk 'Nama Item
    If Not IsNull(rs!berat_jualan) Then Frm68.MSFlexGrid2.TextMatrix(x, 6) = Format(rs!berat_jualan, "#,##0.00") 'Berat
    If Not IsNull(rs!komisyen_per_gram) Then Frm68.MSFlexGrid2.TextMatrix(x, 7) = Format(rs!komisyen_per_gram, "#,##0.00") 'Komisyen Per Gram
    If Not IsNull(rs!komisyen_upah) Then Frm68.MSFlexGrid2.TextMatrix(x, 8) = Format(rs!komisyen_upah, "#,##0.00") 'Komisyen Upah
    If Not IsNull(rs!jumlah_komisyen) Then
        Frm68.MSFlexGrid2.TextMatrix(x, 9) = rs!jumlah_komisyen 'Jumlah Komisyen Agen
        If IsNumeric(rs!jumlah_komisyen) Then Frm68_JUMLAH_KOMISEN = Frm68_JUMLAH_KOMISEN + rs!jumlah_komisyen
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

'Frm68.L8_Text = x 'Bil Data
'Frm68.L9_Text = Format(Frm68_JUMLAH_KOMISEN, "#,##0.00") 'Jumlah Komisyen

'#### Jumlah Bilangan Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) , SUM(jumlah_komisyen) from 23_senarai_jualan where  status_rekod = 1 AND no_rujukan_agen_dropship='" & Frm68.L5_Text & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Frm68.L61_Text = Format(rs(0), "#,##0")
If Not IsNull(rs(1)) Then Frm68.L62_Text = "RM " & Format(rs(1), "#,##0.00")

rs.Close
Set rs = Nothing
'#### Jumlah Bilangan Keseluruhan #### - End

If x <> 0 Then
    Frm68.L56_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If

If x <> 0 Then
    Frm68.Frame10.Visible = True
    Frm68.Frame9.Visible = False
Else
    MsgBox "Tiada rekod dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm68_report_belian_header()
'on error resume next
'#### Header Report Belian #### - Start
Frm68.MSFlexGrid6.Clear
Frm68.MSFlexGrid6.RowHeight(0) = 900
Frm68.MSFlexGrid6.FormatString = "No.|<No.|<ID|<Tarikh|<No. Invoice|<No. Siri Produk|<Nama Item|<Berat (g)|<Harga Semasa (RM/g)|<Upah (RM)|<Harga (RM)"

Frm68.MSFlexGrid6.Rows = 1
Frm68.MSFlexGrid6.ColWidth(0) = 600
Frm68.MSFlexGrid6.ColWidth(1) = 0
Frm68.MSFlexGrid6.ColWidth(2) = 0
Frm68.MSFlexGrid6.ColWidth(3) = 1500 'Tarikh
Frm68.MSFlexGrid6.ColAlignment(3) = 4
Frm68.MSFlexGrid6.ColWidth(4) = 1500 'No. Invoice
Frm68.MSFlexGrid6.ColAlignment(4) = 4
Frm68.MSFlexGrid6.ColWidth(5) = 1600 'No. Siri Produk
Frm68.MSFlexGrid6.ColAlignment(5) = 4
Frm68.MSFlexGrid6.ColWidth(6) = 3900 'Nama Item
Frm68.MSFlexGrid6.ColWidth(7) = 1000 'Berat (g)
Frm68.MSFlexGrid6.ColAlignment(7) = 7
Frm68.MSFlexGrid6.ColWidth(8) = 1500 'Harga Semasa (RM/g)
Frm68.MSFlexGrid6.ColAlignment(8) = 7
Frm68.MSFlexGrid6.ColWidth(9) = 1500 'Upah (RM)
Frm68.MSFlexGrid6.ColAlignment(9) = 7
Frm68.MSFlexGrid6.ColWidth(10) = 1500 'Harga (RM)
Frm68.MSFlexGrid6.ColAlignment(10) = 7
'#### Header Report Belian #### - End
End Sub
Sub Frm68_report_buyback_header()
'on error resume next
'#### Header Report Trade In #### - Start
Frm68.MSFlexGrid7.Clear
Frm68.MSFlexGrid7.RowHeight(0) = 900
Frm68.MSFlexGrid7.FormatString = "No.|<No.|<ID|<Tarikh|<No. Invoice|<No. Siri Produk|<Nama Item|<Berat (g)|<Harga Semasa (RM/g)|<Spread (%)|<Harga (RM)"

Frm68.MSFlexGrid7.Rows = 1
Frm68.MSFlexGrid7.ColWidth(0) = 600
Frm68.MSFlexGrid7.ColWidth(1) = 0
Frm68.MSFlexGrid7.ColWidth(2) = 0
Frm68.MSFlexGrid7.ColWidth(3) = 1500 'Tarikh
Frm68.MSFlexGrid7.ColAlignment(3) = 4
Frm68.MSFlexGrid7.ColWidth(4) = 1500 'No. Invoice
Frm68.MSFlexGrid7.ColAlignment(4) = 4
Frm68.MSFlexGrid7.ColWidth(5) = 1500 'No. Siri Produk
Frm68.MSFlexGrid7.ColAlignment(5) = 4
Frm68.MSFlexGrid7.ColWidth(6) = 3800 'Nama Item
Frm68.MSFlexGrid7.ColWidth(7) = 1000 'Berat (g)
Frm68.MSFlexGrid7.ColAlignment(7) = 7
Frm68.MSFlexGrid7.ColWidth(8) = 1500 'Harga Semasa (RM/g)
Frm68.MSFlexGrid7.ColAlignment(8) = 7
Frm68.MSFlexGrid7.ColWidth(9) = 1500 'Spread (%)
Frm68.MSFlexGrid7.ColAlignment(9) = 7
Frm68.MSFlexGrid7.ColWidth(10) = 1500 'Harga (RM)
Frm68.MSFlexGrid7.ColAlignment(10) = 7
'#### Header Report Trade In #### - End
End Sub
Sub Frm68_report_tempahan_header()
'on error resume next
'#### Header Report Tempahan #### - Start
Frm68.MSFlexGrid8.Clear
Frm68.MSFlexGrid8.RowHeight(0) = 900
Frm68.MSFlexGrid8.FormatString = "No.|<No.|<ID|<Tarikh Tempahan|<No. Siri Produk|<Nama Item|<Berat / Anggaran Berat (g)|<Status"

Frm68.MSFlexGrid8.Rows = 1
Frm68.MSFlexGrid8.ColWidth(0) = 600
Frm68.MSFlexGrid8.ColWidth(1) = 0
Frm68.MSFlexGrid8.ColWidth(2) = 0
Frm68.MSFlexGrid8.ColWidth(3) = 1500 'Tarikh Tempahan
Frm68.MSFlexGrid8.ColAlignment(3) = 4
Frm68.MSFlexGrid8.ColWidth(4) = 1500 'No. Siri Produk
Frm68.MSFlexGrid8.ColAlignment(4) = 4
Frm68.MSFlexGrid8.ColWidth(5) = 3800 'Nama Item
Frm68.MSFlexGrid8.ColWidth(6) = 1000 'Berat (g)
Frm68.MSFlexGrid8.ColAlignment(6) = 7
Frm68.MSFlexGrid8.ColWidth(7) = 1500 'Status
'#### Header Report Tempahan #### - End
End Sub
Sub Frm68_report_ansuran_header()
'on error resume next
'#### Header Report Ansuran #### - Start
Frm68.MSFlexGrid9.Clear
Frm68.MSFlexGrid9.RowHeight(0) = 900
Frm68.MSFlexGrid9.FormatString = "No.|<No.|<ID|<Tarikh Pendaftaran|<No. Siri Produk|<Nama Item|<Berat (g)|<Jumlah Bayaran Terkumpul (RM)|<Status"

Frm68.MSFlexGrid9.Rows = 1
Frm68.MSFlexGrid9.ColWidth(0) = 600
Frm68.MSFlexGrid9.ColWidth(1) = 0
Frm68.MSFlexGrid9.ColWidth(2) = 0
Frm68.MSFlexGrid9.ColWidth(3) = 1500 'Tarikh Pendaftaran
Frm68.MSFlexGrid9.ColAlignment(3) = 4
Frm68.MSFlexGrid9.ColWidth(4) = 1500 'No. Siri Produk
Frm68.MSFlexGrid9.ColAlignment(4) = 4
Frm68.MSFlexGrid9.ColWidth(5) = 3800 'Nama Item
Frm68.MSFlexGrid9.ColWidth(6) = 1000 'Berat (g)
Frm68.MSFlexGrid9.ColAlignment(6) = 7
Frm68.MSFlexGrid9.ColWidth(7) = 1500 'Jumlah Bayaran Terkumpul (RM)
Frm68.MSFlexGrid9.ColAlignment(7) = 7
Frm68.MSFlexGrid9.ColWidth(8) = 1500 'Status
'#### Header Report Ansuran #### - End
End Sub
Sub Frm68_rekod_servis_header()
'on error resume next
'#### Header Report Servis #### - Start
Frm68.MSFlexGrid10.Clear
Frm68.MSFlexGrid10.RowHeight(0) = 900
Frm68.MSFlexGrid10.FormatString = "No.|<No.|<ID|<Tarikh|<No. Invoice|<Jumlah (RM)"

Frm68.MSFlexGrid10.Rows = 1
Frm68.MSFlexGrid10.ColWidth(0) = 600
Frm68.MSFlexGrid10.ColWidth(1) = 0
Frm68.MSFlexGrid10.ColWidth(2) = 0
Frm68.MSFlexGrid10.ColWidth(3) = 1500 'Tarikh
Frm68.MSFlexGrid10.ColAlignment(3) = 4
Frm68.MSFlexGrid10.ColWidth(4) = 1500 'No. Invoice
Frm68.MSFlexGrid10.ColAlignment(4) = 4
Frm68.MSFlexGrid10.ColWidth(5) = 1500 'Jumlah (RM)
Frm68.MSFlexGrid10.ColAlignment(5) = 7
Frm68.MSFlexGrid10.ColAlignment(1) = 1
'#### Header Report Servis #### - End
End Sub
Sub Frm68_report_belian_page()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim Frm68_LM_No_PELANGGAN As String

x = 0
Frm68_PAGE_SIZE = 32

If Frm68.L47_Text <> vbNullString Then
    TM = Frm68.L47_Text 'Tarikh Mula
Else
    TM = DateTime.Date
End If
If Frm68.L48_Text <> vbNullString Then
    TA = Frm68.L48_Text 'Tarikh Akhir
Else
    TA = DateTime.Date
End If
If Frm68.L46_Text <> vbNullString Then
    Frm68_LM_No_PELANGGAN = Frm68.L46_Text 'No. Pelanggan
Else
    Frm68_LM_No_PELANGGAN = 1
End If

LM_START_ROW = Frm68.L56_Text

If GM_NEXT_PREV = 0 Then
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm68_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        LM_START_ROW = LM_START_ROW - Frm68_PAGE_SIZE
    End If
End If

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 23_senarai_jualan where no_rujukan_pembeli='" & Frm68_LM_No_PELANGGAN & "' AND status_rekod = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm68_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm68.MSFlexGrid6.Rows = x + 1
    Frm68.MSFlexGrid6.TextMatrix(x, 0) = x 'No.
    Frm68.MSFlexGrid6.TextMatrix(x, 1) = x 'No.
    Frm68.MSFlexGrid6.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm68.MSFlexGrid6.TextMatrix(x, 3) = rs!tarikh 'Tarikh Jualan
    If Not IsNull(rs!no_resit) Then Frm68.MSFlexGrid6.TextMatrix(x, 4) = rs!no_resit 'No. Resit
    If Not IsNull(rs!no_siri_Produk) Then Frm68.MSFlexGrid6.TextMatrix(x, 5) = rs!no_siri_Produk 'No. Siri Produk
    If Not IsNull(rs!kategori_Produk) Then Frm68.MSFlexGrid6.TextMatrix(x, 6) = rs!kategori_Produk 'Kategori Produk
    If Not IsNull(rs!berat_jualan) Then 'Berat Jualan (g)
        Frm68.MSFlexGrid6.TextMatrix(x, 7) = Format(rs!berat_jualan, "#,##0.00") 'Berat Jualan (g)
        If IsNumeric(rs!berat_jualan) Then Frm68_LM_BERAT = Frm68_LM_BERAT + rs!berat_jualan 'Total Berat Jualan (g)
    End If
    If Not IsNull(rs!harga_Semasa) Then Frm68.MSFlexGrid6.TextMatrix(x, 8) = Format(rs!harga_Semasa, "#,##0.00") 'Harga Semasa (RM/g)
    If Not IsNull(rs!UPAH) Then Frm68.MSFlexGrid6.TextMatrix(x, 9) = Format(rs!UPAH, "#,##0.00") 'Upah (RM)
    If Not IsNull(rs!harga_dengan_gst) Then
        Frm68.MSFlexGrid6.TextMatrix(x, 10) = Format(rs!harga_dengan_gst, "#,##0.00") 'Harga Jualan (RM)
        If IsNumeric(rs!harga_dengan_gst) Then Frm68_LM_HARGA = Frm68_LM_HARGA + rs!harga_dengan_gst 'Total Harga Jualan (RM)
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm68.L58_Text = x

'#### Jumlah Bilangan Barang Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(no_siri_Produk) from 23_senarai_jualan where no_rujukan_pembeli='" & Frm68_LM_No_PELANGGAN & "' AND status_rekod = 1 AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm68.L59_Text = rs(0)
    If rs(0) = vbNullString Then
        Frm68.L59_Text = 0
    End If
Else
    Frm68.L59_Text = 0
End If

rs.Close
Set rs = Nothing
'#### Jumlah Bilangan Barang Keseluruhan #### - End

If x <> 0 Then
    Frm68.L56_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm68_report_buyback_page()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim Frm68_LM_No_PELANGGAN As String

x = 0
Frm68_PAGE_SIZE = 32

If Frm68.L47_Text <> vbNullString Then
    TM = Frm68.L47_Text 'Tarikh Mula
Else
    TM = DateTime.Date
End If
If Frm68.L48_Text <> vbNullString Then
    TA = Frm68.L48_Text 'Tarikh Akhir
Else
    TA = DateTime.Date
End If
If Frm68.L46_Text <> vbNullString Then
    Frm68_LM_No_PELANGGAN = Frm68.L46_Text 'No. Pelanggan
Else
    Frm68_LM_No_PELANGGAN = 1
End If

LM_START_ROW = Frm68.L56_Text

If GM_NEXT_PREV = 0 Then
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm68_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        LM_START_ROW = LM_START_ROW - Frm68_PAGE_SIZE
    End If
End If

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from Data_Database where no_rujukan_pelanggan_buyback='" & Frm68_LM_No_PELANGGAN & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh_belian ASC LIMIT " & LM_START_ROW & "," & Frm68_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm68.MSFlexGrid7.Rows = x + 1
    Frm68.MSFlexGrid7.TextMatrix(x, 0) = x 'No.
    Frm68.MSFlexGrid7.TextMatrix(x, 1) = x 'No.
    Frm68.MSFlexGrid7.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh_belian) Then Frm68.MSFlexGrid7.TextMatrix(x, 3) = rs!tarikh_belian 'Tarikh Belian
    If Not IsNull(rs!bill_No_Trade_In) Then Frm68.MSFlexGrid7.TextMatrix(x, 4) = rs!bill_No_Trade_In 'No. Invoice
    If Not IsNull(rs!no_siri_Produk) Then Frm68.MSFlexGrid7.TextMatrix(x, 5) = rs!no_siri_Produk 'No. Siri Produk
    If Not IsNull(rs!kategori_Produk) Then Frm68.MSFlexGrid7.TextMatrix(x, 6) = rs!kategori_Produk 'Kategori Produk
    If Not IsNull(rs!Berat) Then Frm68.MSFlexGrid7.TextMatrix(x, 7) = Format(rs!Berat, "#,##0.00") 'Berat (g)
    If Not IsNull(rs!kos_Belian_Gram) Then Frm68.MSFlexGrid7.TextMatrix(x, 8) = Format(rs!kos_Belian_Gram, "#,##0.00") 'Rate Penerimaan (RM/g)
    If Not IsNull(rs!SpreadValue) Then Frm68.MSFlexGrid7.TextMatrix(x, 9) = Format(rs!SpreadValue, "#,##0.00") 'Spread (%)
    If Not IsNull(rs!kos_item_tanpa_tax) Then Frm68.MSFlexGrid7.TextMatrix(x, 10) = Format(rs!kos_item_tanpa_tax, "#,##0.00") 'Harga Belian (RM) : Tidak Campur Cukai GST
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm68.L58_Text = x

'#### Jumlah Bilangan Barang Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(no_rujukan_pelanggan_buyback) from Data_Database where no_rujukan_pelanggan_buyback='" & Frm68_LM_No_PELANGGAN & "' AND StatusItem <> 0 AND tarikh_belian BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh_belian ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm68.L59_Text = rs(0)
    If rs(0) = vbNullString Then
        Frm68.L59_Text = 0
    End If
Else
    Frm68.L59_Text = 0
End If

rs.Close
Set rs = Nothing
'#### Jumlah Bilangan Barang Keseluruhan #### - End

If x <> 0 Then
    Frm68.L56_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm68_report_tempahan_page()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim Frm68_LM_No_PELANGGAN As String

x = 0
Frm68_PAGE_SIZE = 32

If Frm68.L47_Text <> vbNullString Then
    TM = Frm68.L47_Text 'Tarikh Mula
Else
    TM = DateTime.Date
End If
If Frm68.L48_Text <> vbNullString Then
    TA = Frm68.L48_Text 'Tarikh Akhir
Else
    TA = DateTime.Date
End If
If Frm68.L46_Text <> vbNullString Then
    Frm68_LM_No_PELANGGAN = Frm68.L46_Text 'No. Pelanggan
Else
    Frm68_LM_No_PELANGGAN = 1
End If

LM_START_ROW = Frm68.L56_Text

If GM_NEXT_PREV = 0 Then
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm68_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        LM_START_ROW = LM_START_ROW - Frm68_PAGE_SIZE
    End If
End If

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 40_tempahan_deposit where no_rujukan_pelanggan='" & Frm68_LM_No_PELANGGAN & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm68_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm68.MSFlexGrid8.Rows = x + 1
    Frm68.MSFlexGrid8.TextMatrix(x, 0) = x 'No.
    Frm68.MSFlexGrid8.TextMatrix(x, 1) = x 'No.
    If Not IsNull(rs!ID) Then Frm68.MSFlexGrid8.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm68.MSFlexGrid8.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!no_siri_Produk) Then Frm68.MSFlexGrid8.TextMatrix(x, 4) = rs!no_siri_Produk 'No. Siri Produk
    If Not IsNull(rs!kategori_Produk) Then Frm68.MSFlexGrid8.TextMatrix(x, 5) = rs!kategori_Produk 'Kategori Produk
    If Not IsNull(rs!jenis_tempahan) Then
        If rs!jenis_tempahan = 0 Then
            If Not IsNull(rs!anggaran_berat) Then Frm68.MSFlexGrid8.TextMatrix(x, 6) = Format(rs!anggaran_berat, "#,##0.00") 'Anggaran Berat
        ElseIf rs!jenis_tempahan = 1 Then
            If Not IsNull(rs!berat_jualan) Then Frm68.MSFlexGrid8.TextMatrix(x, 6) = Format(rs!berat_jualan, "#,##0.00") 'Berat
        End If
    End If
    If Not IsNull(rs!Status) Then Frm68.MSFlexGrid8.TextMatrix(x, 7) = rs!Status 'Status
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm68.L58_Text = x

'#### Jumlah Bilangan Barang Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 40_tempahan_deposit where no_rujukan_pelanggan='" & Frm68_LM_No_PELANGGAN & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm68.L59_Text = rs(0)
    If rs(0) = vbNullString Then
        Frm68.L59_Text = 0
    End If
Else
    Frm68.L59_Text = 0
End If

rs.Close
Set rs = Nothing
'#### Jumlah Bilangan Barang Keseluruhan #### - End

If x <> 0 Then
    Frm68.L56_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm68_report_ansuran_page()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim Frm68_LM_No_PELANGGAN As String

x = 0
Frm68_PAGE_SIZE = 32

If Frm68.L47_Text <> vbNullString Then
    TM = Frm68.L47_Text 'Tarikh Mula
Else
    TM = DateTime.Date
End If
If Frm68.L48_Text <> vbNullString Then
    TA = Frm68.L48_Text 'Tarikh Akhir
Else
    TA = DateTime.Date
End If
If Frm68.L46_Text <> vbNullString Then
    Frm68_LM_No_PELANGGAN = Frm68.L46_Text 'No. Pelanggan
Else
    Frm68_LM_No_PELANGGAN = 1
End If

LM_START_ROW = Frm68.L56_Text

If GM_NEXT_PREV = 0 Then
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm68_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        LM_START_ROW = LM_START_ROW - Frm68_PAGE_SIZE
    End If
End If

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 27_senarai_ansuran where no_rujukan_pelanggan='" & Frm68_LM_No_PELANGGAN & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm68_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm68.MSFlexGrid9.Rows = x + 1
    Frm68.MSFlexGrid9.TextMatrix(x, 0) = x 'No.
    Frm68.MSFlexGrid9.TextMatrix(x, 1) = x 'No.
    If Not IsNull(rs!ID) Then Frm68.MSFlexGrid9.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm68.MSFlexGrid9.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!no_siri_Produk) Then Frm68.MSFlexGrid9.TextMatrix(x, 4) = rs!no_siri_Produk 'No. Siri Produk
    If Not IsNull(rs!kategori_Produk) Then Frm68.MSFlexGrid9.TextMatrix(x, 5) = rs!kategori_Produk 'Kategori Produk
    If Not IsNull(rs!berat_jualan) Then Frm68.MSFlexGrid9.TextMatrix(x, 6) = Format(rs!berat_jualan, "#,##0.00") 'Berat Jualan (g)
    If Not IsNull(rs!jumlah_bayaran) Then Frm68.MSFlexGrid9.TextMatrix(x, 7) = Format(rs!jumlah_bayaran, "#,##0.00") 'Jumlah Bayaran Terkumpul (RM)
    If Not IsNull(rs!Status) Then Frm68.MSFlexGrid9.TextMatrix(x, 8) = rs!Status 'Status
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm68.L58_Text = x

'#### Jumlah Bilangan Barang Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 27_senarai_ansuran where no_rujukan_pelanggan='" & Frm68_LM_No_PELANGGAN & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm68.L59_Text = rs(0)
    If rs(0) = vbNullString Then
        Frm68.L59_Text = 0
    End If
Else
    Frm68.L59_Text = 0
End If

rs.Close
Set rs = Nothing
'#### Jumlah Bilangan Barang Keseluruhan #### - End

If x <> 0 Then
    Frm68.L56_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm68_rekod_servis_page()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim Frm68_LM_No_PELANGGAN As String

x = 0
Frm68_PAGE_SIZE = 32

If Frm68.L47_Text <> vbNullString Then
    TM = Frm68.L47_Text 'Tarikh Mula
Else
    TM = DateTime.Date
End If
If Frm68.L48_Text <> vbNullString Then
    TA = Frm68.L48_Text 'Tarikh Akhir
Else
    TA = DateTime.Date
End If
If Frm68.L46_Text <> vbNullString Then
    Frm68_LM_No_PELANGGAN = Frm68.L46_Text 'No. Pelanggan
Else
    Frm68_LM_No_PELANGGAN = 1
End If

LM_START_ROW = Frm68.L56_Text

If GM_NEXT_PREV = 0 Then
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + Frm68_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        LM_START_ROW = LM_START_ROW - Frm68_PAGE_SIZE
    End If
End If


Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 36_akaun_servis where no_rujukan_pembeli='" & Frm68_LM_No_PELANGGAN & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC LIMIT " & LM_START_ROW & "," & Frm68_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm68.MSFlexGrid10.Rows = x + 1
    Frm68.MSFlexGrid10.TextMatrix(x, 0) = x 'No.
    Frm68.MSFlexGrid10.TextMatrix(x, 1) = x 'No.
    If Not IsNull(rs!ID) Then Frm68.MSFlexGrid10.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm68.MSFlexGrid10.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!no_resit_servis) Then Frm68.MSFlexGrid10.TextMatrix(x, 4) = rs!no_resit_servis 'No. Resit Servis
    If Not IsNull(rs!harga_dengan_gst) Then Frm68.MSFlexGrid10.TextMatrix(x, 5) = Format(rs!harga_dengan_gst, "#,##0.00") 'Jumlah Dengan GST (RM)
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Frm68.L58_Text = x

'#### Jumlah Bilangan Barang Keseluruhan #### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 36_akaun_servis where no_rujukan_pembeli='" & Frm68_LM_No_PELANGGAN & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    Frm68.L59_Text = rs(0)
    If rs(0) = vbNullString Then
        Frm68.L59_Text = 0
    End If
Else
    Frm68.L59_Text = 0
End If

rs.Close
Set rs = Nothing
'#### Jumlah Bilangan Barang Keseluruhan #### - End

If x <> 0 Then
    Frm68.L56_Text = LM_START_ROW
Else
'    MsgBox "Tiada data dijumpai.", vbInformation, "Info"
End If
End Sub
Sub Frm68_rekod_servis()
'on error resume next
Dim TM As Date
Dim TA As Date
Dim Frm68_LM_No_PELANGGAN As String

If Frm68.L47_Text <> vbNullString Then
    TM = Frm68.L47_Text 'Tarikh Mula
Else
    TM = DateTime.Date
End If
If Frm68.L48_Text <> vbNullString Then
    TA = Frm68.L48_Text 'Tarikh Akhir
Else
    TA = DateTime.Date
End If
If Frm68.L46_Text <> vbNullString Then
    Frm68_LM_No_PELANGGAN = Frm68.L46_Text 'No. Pelanggan
Else
    Frm68_LM_No_PELANGGAN = 1
End If

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 36_akaun_servis where no_rujukan_pembeli='" & Frm68_LM_No_PELANGGAN & "' AND tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by tarikh ASC", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    x = x + 1
    Frm68.MSFlexGrid10.Rows = x + 1
    Frm68.MSFlexGrid10.TextMatrix(x, 0) = x 'No.
    Frm68.MSFlexGrid10.TextMatrix(x, 1) = x 'No.
    If Not IsNull(rs!ID) Then Frm68.MSFlexGrid10.TextMatrix(x, 2) = rs!ID 'No. ID
    If Not IsNull(rs!tarikh) Then Frm68.MSFlexGrid10.TextMatrix(x, 3) = rs!tarikh 'Tarikh
    If Not IsNull(rs!no_resit_servis) Then Frm68.MSFlexGrid10.TextMatrix(x, 4) = rs!no_resit_servis 'No. Resit Servis
    If Not IsNull(rs!harga_dengan_gst) Then Frm68.MSFlexGrid10.TextMatrix(x, 5) = Format(rs!harga_dengan_gst, "#,##0.00") 'Jumlah Dengan GST (RM)
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing
End Sub
Sub Frm68_invoice_yuran_ahli()
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


Report70.Sections("Section1").Controls("L1").Caption = vbNullString 'Nama
Report70.Sections("Section1").Controls("L2").Caption = vbNullString 'No. kad pengenalan
Report70.Sections("Section1").Controls("L3").Caption = vbNullString 'No. telefon
Report70.Sections("Section1").Controls("L4").Caption = vbNullString 'No. keahlian
Report70.Sections("Section1").Controls("L5").Caption = vbNullString 'E-mail
Report70.Sections("Section1").Controls("L6").Caption = vbNullString 'Alamat
Report70.Sections("Section1").Controls("L7").Caption = "RM 0.00" 'Jumlah bayaran
Report70.Sections("Section1").Controls("L8").Caption = vbNullString 'Kategori
Report70.Sections("Section1").Controls("L9").Caption = vbNullString 'Tarikh
Report70.Sections("Section1").Controls("L10").Caption = vbNullString 'No. invoice

'### Reset maklumat kedai ### - Start
Report70.Sections("Section4").Controls("L200").Caption = vbNullString 'Nama kedai
Report70.Sections("Section4").Controls("L201").Caption = vbNullString 'No. pendaftaran kedai
Report70.Sections("Section4").Controls("L202").Caption = vbNullString 'Alamat
Report70.Sections("Section4").Controls("L203").Caption = vbNullString 'No. telefon kedai
Report70.Sections("Section4").Controls("L204").Caption = vbNullString 'Maklumat GST
'### Reset maklumat kedai ### - End

'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report70.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report70.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report70.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report70.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report70.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

'### Paparan Penyata ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from senarai_pelanggan where no_invoice='" & G_INVOICE_AHLI & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    If Not IsNull(rs!Nama) Then Report70.Sections("Section1").Controls("L1").Caption = rs!Nama 'Nama
    If Not IsNull(rs!no_ic) Then Report70.Sections("Section1").Controls("L2").Caption = rs!no_ic 'No. kad pengenalan
    If Not IsNull(rs!no_tel) Then Report70.Sections("Section1").Controls("L3").Caption = rs!no_tel 'No. telefon
    If Not IsNull(rs!no_pelanggan) Then Report70.Sections("Section1").Controls("L4").Caption = rs!no_pelanggan 'No. keahlian
    If Not IsNull(rs!Email) Then Report70.Sections("Section1").Controls("L5").Caption = rs!Email 'E-mail
    If Not IsNull(rs!alamat) Then Report70.Sections("Section1").Controls("L6").Caption = rs!alamat 'Alamat
    If Not IsNull(rs!jumlah_yuran) Then Report70.Sections("Section1").Controls("L7").Caption = "RM " & Format(rs!jumlah_yuran, "#,##0.00") 'Jumlah bayaran
    
    If Not IsNull(rs!kategori_pelanggan) Then
        If rs!kategori_pelanggan = 1 Then Report70.Sections("Section1").Controls("L8").Caption = "Pelanggan Biasa"
        If rs!kategori_pelanggan = 2 Then Report70.Sections("Section1").Controls("L8").Caption = "Ahli Biasa"
        If rs!kategori_pelanggan = 3 Then Report70.Sections("Section1").Controls("L8").Caption = "Silver"
        If rs!kategori_pelanggan = 4 Then Report70.Sections("Section1").Controls("L8").Caption = "Gold"
        If rs!kategori_pelanggan = 5 Then Report70.Sections("Section1").Controls("L8").Caption = "Platinum"
    End If
        
    If Not IsNull(rs!tarikh) Then Report70.Sections("Section1").Controls("L9").Caption = rs!tarikh 'Tarikh
    If Not IsNull(rs!no_invoice) Then Report70.Sections("Section1").Controls("L10").Caption = rs!no_invoice 'No. invoice

    Set Report70.DataSource = rs
    Report70.Show
End If

'rs.Close
Set rs = Nothing
'### Paparan Penyata ### - End
End Sub
Sub sys_config_membership()
'On Error Resume Next
Dim File_Path As String
File_Path = App.Path & "\sys_config_membership.txt"
Open File_Path For Input As #1

Line Input #1, G_MODE 'YES : Ada kad keahlian , NO : Tiada kad keahlian
Line Input #1, G_MIN_LEN 'Panjang minimum
Line Input #1, G_MAX_LEN 'Panjang maksimum
Line Input #1, G_CODE 'Kod kedai
Close #1
End Sub
Sub frm68_statement_komisen()
'On Error Resume Next
Dim rs1 As ADODB.Recordset

If G_RANKING_FIELD = "bil_barang" Then

    LM_RANKING = "bilangan barang yang dijual."
    
ElseIf G_RANKING_FIELD = "jumlah_berat" Then
    
    LM_RANKING = "jumlah berat yang dijual."
    
ElseIf G_RANKING_FIELD = "jumlah_harga" Then
    
    LM_RANKING = "jumlah harga yang terjual."

ElseIf G_RANKING_FIELD = "jumlah_komisen" Then
    
    LM_RANKING = "jumlah komisen yang diperolehi."
    
End If
            
Report77.Sections("Section4").Controls("L1").Caption = vbNullString
Report77.Sections("Section5").Controls("L6").Caption = 0
Report77.Sections("Section5").Controls("L7").Caption = Format(0, "#,##0.00 g")
Report77.Sections("Section5").Controls("L8").Caption = "RM " & Format(0, "#,##0.00")
Report77.Sections("Section5").Controls("L9").Caption = "RM " & Format(0, "#,##0.00")
Report77.Sections("Section5").Controls("L10").Caption = vbNullString

Report77.Sections("Section5").Controls("L10").Caption = "Report ini dikira dan dikeluarkan pada " & Now
Report77.Sections("Section4").Controls("L1").Caption = "Senarai ranking jualan oleh agen dropship dari " & Frm68.DTPicker7 & " hingga " & Frm68.DTPicker8 & " yang disusun mengikut " & LM_RANKING

            
If MDI_frm1.L4_Text = "HQ" Then
    
    LM_NAMA_HEADER = "HQ"
    
Else
    
    LM_NAMA_HEADER = MDI_frm1.L20_Text
    
End If
        
'### Maklumat kedai ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 56_maklumat_kedai where cawangan='" & LM_NAMA_HEADER & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If Not IsNull(rs!nama_kedai) Then Report77.Sections("Section4").Controls("L200").Caption = rs!nama_kedai
    If Not IsNull(rs!no_pendaftaran) Then Report77.Sections("Section4").Controls("L201").Caption = rs!no_pendaftaran
    If Not IsNull(rs!alamat) Then Report77.Sections("Section4").Controls("L202").Caption = rs!alamat
    If Not IsNull(rs!no_tel) Then Report77.Sections("Section4").Controls("L203").Caption = rs!no_tel
    If Not IsNull(rs!no_id_gst) Then Report77.Sections("Section4").Controls("L204").Caption = rs!no_id_gst
End If

rs.Close
Set rs = Nothing
'### Maklumat kedai ### - End

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select SUM(bil_barang) , SUM(jumlah_berat) , SUM(jumlah_harga) , SUM(jumlah_komisen) from 75_senarai_komisen_agen", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then Report77.Sections("Section5").Controls("L6").Caption = rs(0) 'Jumlah bilang barang
If Not IsNull(rs(1)) Then Report77.Sections("Section5").Controls("L7").Caption = Format(rs(1), "#,##0.00 g") 'Jumlah berat jualan
If Not IsNull(rs(2)) Then Report77.Sections("Section5").Controls("L8").Caption = "RM " & Format(rs(2), "#,##0.00") 'Jumlah harga jualan
If Not IsNull(rs(3)) Then Report77.Sections("Section5").Controls("L9").Caption = "RM " & Format(rs(3), "#,##0.00") 'Jumlah komisen jualan

rs.Close
Set rs = Nothing

'###Senarai komisyen bagi Agen / Staff###
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 75_senarai_komisen_agen order by ranking ASC", cn, adOpenKeyset, adLockOptimistic
    
While rs.EOF = False
    Set Report77.DataSource = rs
    Report77.Show
    rs.MoveNext
Wend

'rs.Close
Set rs = Nothing
End Sub
