Attribute VB_Name = "Module17"
Sub Expiry()
'On Error Resume Next
Dim TARIKH1 As Date
Dim TARIKH2 As Date

EXP_SYS = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from Emas where a='" & 1000 & "'", cn, adOpenKeyset, adLockOptimistic

'If rs!u = 1 Then
'    MsgBox "Sila hubungi developer [Insan 012-6488553] untuk terus menggunakan sistem ini.", vbExclamation, "Info"
'    End
'End If

If Not rs.EOF Then
    'TARIKH1 = Format(rs!d, "00") & "-" & Format(rs!l, "00") & "-" & Format(rs!j, "0000")
    'TARIKH1 = "02-28-2014"
    'TARIKH2 = DateTime.Date
    
    'If TARIKH2 > TARIKH1 Then
    '    rs!u = 1
    '    rs.Update
    '    MsgBox "Sila hubungi developer [Insan 012-6488553] untuk terus menggunakan sistem ini.", vbInformation, "Info"
    '    End
    'End If
    If rs!t = 1122 Then
        Frm3.Show
        Frm3.TxtUsername.SetFocus
    End If
    If IsNull(rs!t) Or rs!t = vbNullString Or rs!t <> 1122 Then
        EXP_SYS = 1
    End If
Else
    MsgBox "Sila hubungi developer [Insan 012-6488553] untuk terus menggunakan sistem ini.", vbExclamation, "Info"
    End
End If

rs.Close
Set rs = Nothing

If EXP_SYS = 1 Then
    Call GenerateLicense
End If

End Sub
Sub GenerateLicense()
'On Error Resume Next
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from Emas where a='" & 1000 & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!p = 0 Or IsNull(rs!p) Or rs!p = vbNullString Then
        InstallationTime = DateTime.Time$
        I1 = Left(InstallationTime, 1)
        If Not IsNumeric(I1) Then
            I1 = 3
        End If
        i2 = Mid(InstallationTime, 2, 1)
        If Not IsNumeric(i2) Then
            i2 = 4
        End If
        i3 = Mid(InstallationTime, 4, 1)
        If Not IsNumeric(i3) Then
            i3 = 5
        End If
        i4 = Mid(InstallationTime, 5, 1)
        If Not IsNumeric(i4) Then
            i4 = 6
        End If
        i5 = Mid(InstallationTime, 7, 1)
        If Not IsNumeric(i5) Then
            i5 = 8
        End If
        i6 = Right(InstallationTime, 1)
        If Not IsNumeric(i6) Then
            i6 = 3
        End If
        
        'a = Right(Int(Rnd * (i1 * 50)), 1)
        'b = Right(Int(Rnd * (i2 * 10)), 1)
        'c = Right(Int(Rnd * (i3 * 100)), 1)
        'd = Right(Int(Rnd * (i4 * 60)), 1)
        'e = Right(Int(Rnd * (i5 * 30)), 1)
        'f = Right(Int(Rnd * (i6 * 26)), 1)
        
        'rs!p = a & b & c & d & e & f
        'rs.Update
        'Frm24.L1_Text = "5.0.0E" & a & b & c & d & e & f
    End If
    If rs!p <> 0 And Len(rs!p) = 6 Then
        Frm24.L1_Text = "5.0.0E" & rs!p
    End If
    'If (rs!p <> 0 And Len(rs!p) <> 6) Or IsNull(rs!p) Then
    '    MsgBox "Sila hubungi developer [Insan 012-6488553] untuk terus menggunakan sistem ini.", vbInformation, "Info"
    '    Exit Sub
    'End If
End If

rs.Close
Set rs = Nothing

Frm24.TB1 = vbNullString
Frm24.Show
End Sub
Sub LicenseCheck()
'On Error Resume Next
NoSistemAnda = Right(Frm24.L1_Text, 6)

If Frm24.TB1 <> vbNullString And IsNumeric(Frm24.TB1) And Len(Frm24.TB1) = 6 Then
    'a = Left(NoSistemAnda, 1)
    'b = Mid(NoSistemAnda, 2, 1)
    'c = Mid(NoSistemAnda, 3, 1)
    'd = Mid(NoSistemAnda, 4, 1)
    'e = Mid(NoSistemAnda, 5, 1)
    'f = Right(NoSistemAnda, 1)
    
    'aa = Right(a * 3, 1)
    'bb = Right(b + 5, 1)
    'CC = Right(c - 7, 1)
    'dd = Right(d * 2, 1)
    'ee = Right(e + 3, 1)
    'ff = Right(f * 4, 1)
    
    InstallationCode = aa & bb & CC & dd & ee & ff
    
    If Frm24.TB1 = InstallationCode Then
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from Emas where a='" & 1000 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            rs!t = 1122
            rs.Update
        End If
        
        rs.Close
        Set rs = Nothing
        
        Unload Frm24
        Frm3.Show
        Frm3.TxtUsername.SetFocus
    Else
        MsgBox "[Installation Code] TIDAK betul!", vbExclamation, "Error"
        End
    End If
Else
    MsgBox "Sila masukkan [Installation Code].", vbExclamation, "Error"
    Exit Sub
End If

End Sub

