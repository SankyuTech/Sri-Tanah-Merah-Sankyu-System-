Attribute VB_Name = "Module48"
Sub frm127_log_header()
'on error resume next
With frm127.LV1
    
    'Alignment : 2 : Center , 1 Right
    .ColumnHeaders.Clear

    frm127.LV1.ListItems.Clear
    
    .ColumnHeaders.Add 1, , "No.", 0
    .ColumnHeaders.Add 2, , "No.", 800, 1
    .ColumnHeaders.Add 3, , "No. ID", 0, 1
    .ColumnHeaders.Add 4, , "Tarikh & Masa", 2200
    .ColumnHeaders.Add 5, , "Detail", 11700
    .ColumnHeaders.Add 6, , "User", 1400
    .ColumnHeaders.Add 7, , "Terminal", 1200, 2
    .ColumnHeaders.Add 8, , "Cawangan", 1900
    
End With
End Sub
Sub frm127_log1()
'on error resume next
Dim frm127_LM_TOTAL_PAGE As Double
Dim TM As Date
Dim TA As Date

frm127_PAGE_SIZE = 31
frm127_LM_TOTAL_PAGE = 0
x = 0

re_gen_report:

If frm127.L5_Text = "1" Then
    
    TM = frm127.L6_Text & " " & "00:00:00"
    TA = frm127.L7_Text & " " & "23:59:59"
    
End If

'frm127.L5_Text '0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh
'frm127.L6_Text = frm127.DTPicker1 'Tarikh Mula
'frm127.L7_Text = frm127.DTPicker2 'Tarikh Akhir
'frm127.L8_Text = frm127.TB1 'Keyword

If frm127.L9_Text = "Semua cawangan" Then

    Frm127_LM_SEARCH_1 = Null
    Frm127_LM_SEARCH_1_LOGIC = "<>"
    
Else

    Frm127_LM_SEARCH_1 = frm127.L9_Text
    Frm127_LM_SEARCH_1_LOGIC = "="
    
End If

If frm127.L10_Text = "Semua terminal" Then

    Frm127_LM_SEARCH_2 = Null
    Frm127_LM_SEARCH_2_LOGIC = "<>"
    
Else

    Frm127_LM_SEARCH_2 = frm127.L10_Text
    Frm127_LM_SEARCH_2_LOGIC = "="
    
End If

If frm127.L11_Text = "Semua user" Then

    Frm127_LM_SEARCH_3 = Null
    Frm127_LM_SEARCH_3_LOGIC = "<>"
    
Else

    Frm127_LM_SEARCH_3 = frm127.L11_Text
    Frm127_LM_SEARCH_3_LOGIC = "="
    
End If

If frm127.L8_Text = "" Then

    Frm127_LM_SEARCH_4 = Null
    Frm127_LM_SEARCH_4_LOGIC = "<>"
    
Else

    Frm127_LM_SEARCH_4 = "%" & frm127.L8_Text & "%"
    Frm127_LM_SEARCH_4_LOGIC = "LIKE"
    
End If

LM_START_ROW = frm127.L69_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + frm127_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If frm127.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - frm127_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        frm127.L67_Text = 1
    End If
End If

frm127_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm127.L5_Text = 0 Then rs.Open "select * from " & G_RECOVERY_DATABASE & ".log where Log_Aktiviti " & Frm127_LM_SEARCH_4_LOGIC & "'" & Frm127_LM_SEARCH_4 & "' AND cawangan " & Frm127_LM_SEARCH_1_LOGIC & "'" & Frm127_LM_SEARCH_1 & "' AND terminal " & Frm127_LM_SEARCH_2_LOGIC & "'" & Frm127_LM_SEARCH_2 & "' AND username " & Frm127_LM_SEARCH_3_LOGIC & "'" & Frm127_LM_SEARCH_3 & "' order by ID ASC LIMIT " & LM_START_ROW & "," & frm127_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If frm127.L5_Text = 1 Then rs.Open "select * from " & G_RECOVERY_DATABASE & ".log where Log_Aktiviti " & Frm127_LM_SEARCH_4_LOGIC & "'" & Frm127_LM_SEARCH_4 & "' AND cawangan " & Frm127_LM_SEARCH_1_LOGIC & "'" & Frm127_LM_SEARCH_1 & "' AND terminal " & Frm127_LM_SEARCH_2_LOGIC & "'" & Frm127_LM_SEARCH_2 & "' AND username " & Frm127_LM_SEARCH_3_LOGIC & "'" & Frm127_LM_SEARCH_3 & "' AND Log_Tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by ID ASC LIMIT " & LM_START_ROW & "," & frm127_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1
    If frm127_LM_PAGE_FOUND = 0 Then
        If frm127.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                frm127.L67_Text = frm127.L67_Text + 1 'Paparan Page ke-xxx
                frm127_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(frm127.L67_Text) Then
                    If frm127.L67_Text <> 1 Then
                        frm127.L67_Text = frm127.L67_Text - 1 'Paparan Page ke-xxx
                        frm127_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If

    Y = ((frm127.L67_Text - 1) * frm127_PAGE_SIZE) + x

    With frm127.LV1.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!Log_Tarikh) Then 'Tarikh & Masa
            .ListSubItems.Add , , rs!Log_Tarikh
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Log_Aktiviti) Then 'Details
            .ListSubItems.Add , , rs!Log_Aktiviti
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!UserName) Then 'User
            .ListSubItems.Add , , rs!UserName
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!terminal) Then 'Terminal
            .ListSubItems.Add , , rs!terminal
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
If frm127.L5_Text = 0 Then rs.Open "select COUNT(ID) from " & G_RECOVERY_DATABASE & ".log where Log_Aktiviti " & Frm127_LM_SEARCH_4_LOGIC & "'" & Frm127_LM_SEARCH_4 & "' AND cawangan " & Frm127_LM_SEARCH_1_LOGIC & "'" & Frm127_LM_SEARCH_1 & "' AND terminal " & Frm127_LM_SEARCH_2_LOGIC & "'" & Frm127_LM_SEARCH_2 & "' AND username " & Frm127_LM_SEARCH_3_LOGIC & "'" & Frm127_LM_SEARCH_3 & "' order by ID ASC LIMIT " & LM_START_ROW & "," & frm127_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If frm127.L5_Text = 1 Then rs.Open "select COUNT(ID) from " & G_RECOVERY_DATABASE & ".log where Log_Aktiviti " & Frm127_LM_SEARCH_4_LOGIC & "'" & Frm127_LM_SEARCH_4 & "' AND cawangan " & Frm127_LM_SEARCH_1_LOGIC & "'" & Frm127_LM_SEARCH_1 & "' AND terminal " & Frm127_LM_SEARCH_2_LOGIC & "'" & Frm127_LM_SEARCH_2 & "' AND username " & Frm127_LM_SEARCH_3_LOGIC & "'" & Frm127_LM_SEARCH_3 & "' AND Log_Tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by ID ASC LIMIT " & LM_START_ROW & "," & frm127_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    frm127_LM_TOTAL_PAGE = Format(rs(0) / frm127_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, frm127_LM_TOTAL_PAGE, ".") <> 0 Then
    
        frm127_LM_PAGE = Split(frm127_LM_TOTAL_PAGE, ".")(0)
        frm127_LM_PAGE_LEBIHAN = Split(frm127_LM_TOTAL_PAGE, ".")(1)
        
        If frm127_LM_PAGE_LEBIHAN <> "00" Then
            frm127.L68_Text = frm127_LM_PAGE + 1
        Else
            frm127.L68_Text = frm127_LM_PAGE
        End If
        
    Else
    
        frm127.L68_Text = frm127_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        frm127.L68_Text = 0
    End If
Else
    frm127.L68_Text = 0
End If

rs.Close
Set rs = Nothing

If x <> 0 Then
    frm127.L69_Text = LM_START_ROW
End If

'If frm127.L67_Text <> vbNullString And IsNumeric(frm127.L67_Text) Then
'    If frm127.L68_Text <> vbNullString And IsNumeric(frm127.L68_Text) Then
'        frm127_LM_CURR_PAGE = frm127.L67_Text
'        frm127_LM_TOTAL_PAGE = frm127.L68_Text
        
'        If frm127_LM_CURR_PAGE > frm127_LM_TOTAL_PAGE Then
            
'            frm127.L67_Text = frm127.L67_Text - 1
'            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
'            GoTo re_gen_report:
            
'        End If
'    End If
'End If

End Sub
Sub frm127_log()
'On Error Resume Next
Dim frm127_LM_TOTAL_PAGE As Double
Dim frm127_LM_CURR_PAGE As Double

Dim TM As Date
Dim TA As Date

frm127_PAGE_SIZE = 31
frm127_LM_TOTAL_PAGE = 0
x = 0

If frm127.L5_Text = "1" Then
    
    TM = frm127.L6_Text & " " & "00:00:00"
    TA = frm127.L7_Text & " " & "23:59:59"
    
End If

'frm127.L5_Text '0 : Tiada pilihan tarikh , 1 : Ada pilihan tarikh
'frm127.L6_Text = frm127.DTPicker1 'Tarikh Mula
'frm127.L7_Text = frm127.DTPicker2 'Tarikh Akhir
'frm127.L8_Text = frm127.TB1 'Keyword

If frm127.L9_Text = "Semua cawangan" Then

    Frm127_LM_SEARCH_1 = Null
    Frm127_LM_SEARCH_1_LOGIC = "<>"
    
Else

    Frm127_LM_SEARCH_1 = frm127.L9_Text
    Frm127_LM_SEARCH_1_LOGIC = "="
    
End If

If frm127.L10_Text = "Semua terminal" Then

    Frm127_LM_SEARCH_2 = Null
    Frm127_LM_SEARCH_2_LOGIC = "<>"
    
Else

    Frm127_LM_SEARCH_2 = frm127.L10_Text
    Frm127_LM_SEARCH_2_LOGIC = "="
    
End If

If frm127.L11_Text = "Semua user" Then

    Frm127_LM_SEARCH_3 = Null
    Frm127_LM_SEARCH_3_LOGIC = "<>"
    
Else

    Frm127_LM_SEARCH_3 = frm127.L11_Text
    Frm127_LM_SEARCH_3_LOGIC = "="
    
End If

If frm127.L8_Text = "" Then

    Frm127_LM_SEARCH_4 = Null
    Frm127_LM_SEARCH_4_LOGIC = "<>"
    
Else

    Frm127_LM_SEARCH_4 = "%" & frm127.L8_Text & "%"
    Frm127_LM_SEARCH_4_LOGIC = "LIKE"
    
End If

re_gen_report:

LM_START_ROW = frm127.L69_Text 'Titik Pencarian Data

If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
    If LM_START_ROW <> -1 Then
        LM_START_ROW = LM_START_ROW + frm127_PAGE_SIZE
    Else
        LM_START_ROW = 0
    End If
ElseIf GM_NEXT_PREV = 1 Then
    If LM_START_ROW <= 0 Then
        LM_START_ROW = 0
    Else
        If frm127.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            LM_START_ROW = LM_START_ROW - frm127_PAGE_SIZE
        End If
    End If
ElseIf GM_NEXT_PREV = 2 Then
    If LM_START_ROW = -1 Then
        LM_START_ROW = 0
        frm127.L67_Text = 1
    End If
End If

frm127_LM_PAGE_FOUND = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
If frm127.L5_Text = 0 Then rs.Open "select * from " & G_RECOVERY_DATABASE & ".log where Log_Aktiviti " & Frm127_LM_SEARCH_4_LOGIC & "'" & Frm127_LM_SEARCH_4 & "' AND cawangan " & Frm127_LM_SEARCH_1_LOGIC & "'" & Frm127_LM_SEARCH_1 & "' AND terminal " & Frm127_LM_SEARCH_2_LOGIC & "'" & Frm127_LM_SEARCH_2 & "' AND username " & Frm127_LM_SEARCH_3_LOGIC & "'" & Frm127_LM_SEARCH_3 & "' order by ID DESC LIMIT " & LM_START_ROW & "," & frm127_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic
If frm127.L5_Text = 1 Then rs.Open "select * from " & G_RECOVERY_DATABASE & ".log where Log_Aktiviti " & Frm127_LM_SEARCH_4_LOGIC & "'" & Frm127_LM_SEARCH_4 & "' AND cawangan " & Frm127_LM_SEARCH_1_LOGIC & "'" & Frm127_LM_SEARCH_1 & "' AND terminal " & Frm127_LM_SEARCH_2_LOGIC & "'" & Frm127_LM_SEARCH_2 & "' AND username " & Frm127_LM_SEARCH_3_LOGIC & "'" & Frm127_LM_SEARCH_3 & "' AND Log_Tarikh BETWEEN '" & TM & "' AND '" & TA & "' order by ID DESC LIMIT " & LM_START_ROW & "," & frm127_PAGE_SIZE, cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False

    x = x + 1
    If frm127_LM_PAGE_FOUND = 0 Then
        If frm127.L70_Text = 0 Then '0 : Bukan page terakhir , 1 : Page Terakhir
            If GM_NEXT_PREV = 0 Then '0 : Next , 1 : Previous
                frm127.L67_Text = frm127.L67_Text + 1 'Paparan Page ke-xxx
                frm127_LM_PAGE_FOUND = 1
            ElseIf GM_NEXT_PREV = 1 Then '0 : Next , 1 : Previous
                If IsNumeric(frm127.L67_Text) Then
                    If frm127.L67_Text <> 1 Then
                        frm127.L67_Text = frm127.L67_Text - 1 'Paparan Page ke-xxx
                        frm127_LM_PAGE_FOUND = 1
                    End If
                End If
            End If
        End If
    End If

    Y = ((frm127.L67_Text - 1) * frm127_PAGE_SIZE) + x

    With frm127.LV1.ListItems.Add(, , rs!ID)
    
        .ListSubItems.Add , , Y
        
        If Not IsNull(rs!ID) Then .ListSubItems.Add , , rs!ID
        
        If Not IsNull(rs!Log_Tarikh) Then 'Tarikh & Masa
            .ListSubItems.Add , , rs!Log_Tarikh
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!Log_Aktiviti) Then 'Details
            .ListSubItems.Add , , rs!Log_Aktiviti
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!UserName) Then 'User
            .ListSubItems.Add , , rs!UserName
        Else
            .ListSubItems.Add , , ""
        End If
        
        If Not IsNull(rs!terminal) Then 'Terminal
            .ListSubItems.Add , , rs!terminal
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
If frm127.L5_Text = 0 Then rs.Open "select COUNT(ID) from " & G_RECOVERY_DATABASE & ".log where Log_Aktiviti " & Frm127_LM_SEARCH_4_LOGIC & "'" & Frm127_LM_SEARCH_4 & "' AND cawangan " & Frm127_LM_SEARCH_1_LOGIC & "'" & Frm127_LM_SEARCH_1 & "' AND terminal " & Frm127_LM_SEARCH_2_LOGIC & "'" & Frm127_LM_SEARCH_2 & "' AND username " & Frm127_LM_SEARCH_3_LOGIC & "'" & Frm127_LM_SEARCH_3 & "'", cn, adOpenKeyset, adLockOptimistic
If frm127.L5_Text = 1 Then rs.Open "select COUNT(ID) from " & G_RECOVERY_DATABASE & ".log where Log_Aktiviti " & Frm127_LM_SEARCH_4_LOGIC & "'" & Frm127_LM_SEARCH_4 & "' AND cawangan " & Frm127_LM_SEARCH_1_LOGIC & "'" & Frm127_LM_SEARCH_1 & "' AND terminal " & Frm127_LM_SEARCH_2_LOGIC & "'" & Frm127_LM_SEARCH_2 & "' AND username " & Frm127_LM_SEARCH_3_LOGIC & "'" & Frm127_LM_SEARCH_3 & "' AND Log_Tarikh BETWEEN '" & TM & "' AND '" & TA & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    frm127_LM_TOTAL_PAGE = Format(rs(0) / frm127_PAGE_SIZE, "0.00") 'Jumlah Page
    
    'Periksa Samada ada titik perpuluhan atau tidak
    If InStr(1, frm127_LM_TOTAL_PAGE, ".") <> 0 Then
    
        frm127_LM_PAGE = Split(frm127_LM_TOTAL_PAGE, ".")(0)
        frm127_LM_PAGE_LEBIHAN = Split(frm127_LM_TOTAL_PAGE, ".")(1)
        
        If frm127_LM_PAGE_LEBIHAN <> "00" Then
            frm127.L68_Text = frm127_LM_PAGE + 1
        Else
            frm127.L68_Text = frm127_LM_PAGE
        End If
        
    Else
    
        frm127.L68_Text = frm127_LM_TOTAL_PAGE
        
    End If

    If rs(0) = vbNullString Then
        frm127.L68_Text = 0
    End If
Else
    frm127.L68_Text = 0
End If

rs.Close
Set rs = Nothing

If x <> 0 Then
    frm127.L69_Text = LM_START_ROW
End If

If frm127.L67_Text <> vbNullString And IsNumeric(frm127.L67_Text) Then
    If frm127.L68_Text <> vbNullString And IsNumeric(frm127.L68_Text) Then
        frm127_LM_CURR_PAGE = frm127.L67_Text
        frm127_LM_TOTAL_PAGE = frm127.L68_Text
        
        If frm127_LM_CURR_PAGE > frm127_LM_TOTAL_PAGE Then
            
            frm127.L67_Text = frm127.L67_Text - 1
            GM_NEXT_PREV = 1 '0 : Next , 1 : Previous
            
            GoTo re_gen_report:
            
        End If
    End If
End If
End Sub

