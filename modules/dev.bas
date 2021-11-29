Attribute VB_Name = "dev"
Sub reset_database()
'On Error GoTo logging:
LM_FOUND = 0

Note = "Adakah anda ingin RESET DATABASE ini?" & vbCrLf & _
        vbNullString & vbCrLf & _
        "Sila masukkan password bagi memadamkan data ini."

LM_PASSWORD = InputBox(Note, "RESET DATABASE", "Sila masukkan password anda")

If StrPtr(LM_PASSWORD) = 0 Then
    Exit Sub
End If
    
If StrPtr(LM_PASSWORD) <> 0 Then
    If InStr(1, LM_PASSWORD, "*") <> 0 Or InStr(1, LM_PASSWORD, "&") <> 0 Or InStr(1, LM_PASSWORD, "/") <> 0 Or InStr(1, LM_PASSWORD, "\") <> 0 Or InStr(1, LM_PASSWORD, "'") <> 0 Or InStr(1, LM_PASSWORD, "`") <> 0 Then
        MsgBox "Password mengandungi simbol yang tidak sah.", vbExclamation, "Error"
        Exit Sub
    End If
    
    If MDI_frm1.L3_Text <> vbNullString Then
LM_CONN = 1
re_conn_1:
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where samaran='" & MDI_frm1.L3_Text & "' and password='" & LM_PASSWORD & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            LM_USER_FOUND = 1
        Else
            MsgBox "Password yang dimasukkan tidak betul/sah." & vbCrLf & _
                    vbNullString & vbCrLf & _
                    "Sila cuba sekali lagi.", vbExclamation, "Info"
        End If
        
        rs.Close
        Set rs = Nothing
        
        If LM_USER_FOUND = 1 Then
 
            Note = "Sila masukkan CAWANGAN ASAL!"
            
            LM_CAWANGAN_OLD = InputBox(Note, "CAWANGAN LAMA", "")
            
            If StrPtr(LM_CAWANGAN_OLD) = 0 Then
                Exit Sub
            End If
            
            Note = "Sila masukkan CAWANGAN BARU!"
            
            LM_CAWANGAN_NEW = InputBox(Note, "CAWANGAN BARU", "")
            
            If StrPtr(LM_CAWANGAN_NEW) = 0 Then
                Exit Sub
            End If
            
'Mulakan reset database
LM_CONN = 2
re_conn_2:
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            
            strsql = "UPDATE 56_maklumat_kedai set cawangan='" & LM_CAWANGAN_NEW & "' where cawangan='" & G_CAWANGAN & "'"
                        
            Set rs = cn.Execute(strsql)
            Set rs = Nothing
            
LM_CONN = 3
re_conn_3:
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            
            strsql = "UPDATE 73_tetapan_upah set default_setting='" & LM_CAWANGAN_NEW & "' where default_setting='" & G_CAWANGAN & "'"
                        
            Set rs = cn.Execute(strsql)
            Set rs = Nothing
            
LM_CONN = 4
re_conn_4:
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            
            strsql = "UPDATE 89_jenis_expense set jenis_expense='" & LM_CAWANGAN_NEW & "' where jenis_expense='" & G_CAWANGAN & "'"
                        
            Set rs = cn.Execute(strsql)
            Set rs = Nothing
            
LM_CONN = 5
re_conn_5:
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            
            strsql = "UPDATE default_setting set Default1='" & LM_CAWANGAN_NEW & "' where Default1='" & G_CAWANGAN & "'"
                        
            Set rs = cn.Execute(strsql)
            Set rs = Nothing

LM_CONN = 6
re_conn_6:
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            
            strsql = "UPDATE employee set cawangan='" & LM_CAWANGAN_NEW & "' where cawangan='" & G_CAWANGAN & "'"
                        
            Set rs = cn.Execute(strsql)
            Set rs = Nothing
            
LM_CONN = 7
re_conn_7:
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            
            strsql = "DELETE from employee where id > 15"
            
            Set rs = cn.Execute(strsql)
            Set rs = Nothing
            
LM_CONN = 8
re_conn_8:
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            
            strsql = "UPDATE layout_barcode set perkara='" & LM_CAWANGAN_NEW & "' where perkara='" & G_CAWANGAN & "'"
                        
            Set rs = cn.Execute(strsql)
            Set rs = Nothing
            
LM_CONN = 9
re_conn_9:
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            
            strsql = "DELETE from setting_database where id > 4"
            
            Set rs = cn.Execute(strsql)
            Set rs = Nothing
            
LM_CONN = 10
re_conn_10:
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            
            strsql = "UPDATE tetapan_barcode set cawangan='" & LM_CAWANGAN_NEW & "' where cawangan='" & G_CAWANGAN & "'"
                        
            Set rs = cn.Execute(strsql)
            Set rs = Nothing
            
LM_CONN = 11
re_conn_11:
            Set rs = New ADODB.Recordset
            If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
            
            strsql = "UPDATE 92_setting_inv set cawangan='" & LM_CAWANGAN_NEW & "' where cawangan='" & G_CAWANGAN & "'"
                        
            Set rs = cn.Execute(strsql)
            Set rs = Nothing
            
            MsgBox "Selesai.", vbInformation, "Info"
        End If
    End If
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " dev : reset_database" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

If LM_ERR_NO = "3704" Or LM_ERR_NO = "-2147467259" Or LM_ERR_NO = "-2147217887" Then
    
    Call Main
    
    If LM_CONN = 1 Then
        Resume re_conn_1:
    ElseIf LM_CONN = 2 Then
        Resume re_conn_2:
    ElseIf LM_CONN = 3 Then
        Resume re_conn_3:
    ElseIf LM_CONN = 4 Then
        Resume re_conn_4:
    ElseIf LM_CONN = 5 Then
        Resume re_conn_5:
    ElseIf LM_CONN = 6 Then
        Resume re_conn_6:
    ElseIf LM_CONN = 7 Then
        Resume re_conn_7:
    ElseIf LM_CONN = 8 Then
        Resume re_conn_8:
    ElseIf LM_CONN = 9 Then
        Resume re_conn_9:
    ElseIf LM_CONN = 10 Then
        Resume re_conn_10:
    ElseIf LM_CONN = 11 Then
        Resume re_conn_11:
    End If
Else
    Resume Next
End If
End Sub


