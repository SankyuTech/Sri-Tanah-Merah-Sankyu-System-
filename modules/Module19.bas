Attribute VB_Name = "Module19"
Public GLOBALRESITNO As Double
Sub NewGenerateResitNo()
'On Error Resume Next
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Default1 = "Default" Then
        RESIT_OLD = rs!ResitNo
        rs!ResitNo = RESIT_OLD + 1
        rs.Update
    End If
End If

rs.Close
Set rs = Nothing
End Sub
Sub NewGenerateNoRujukan()
'On Error Resume Next
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Default1 = "Default" Then
        NORUJUKAN_OLD = rs!NoRujukanSistem
        rs!NoRujukanSistem = NORUJUKAN_OLD + 1
        rs.Update
    End If
End If

rs.Close
Set rs = Nothing
End Sub
Sub NewGenerateBarcode()
'On Error Resume Next
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Default1 = "Default" Then
        BARCODE_OLD = rs!Barcode
        rs!Barcode = BARCODE_OLD + 1
        rs.Update
    End If
End If

rs.Close
Set rs = Nothing
End Sub
Sub NewGenerateDelayNo()

'On Error Resume Next
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Default1 = "Default" Then
        'a = rs!RujukanDelay
        'rs!RujukanDelay = a + 1
        'rs.Update
    End If
End If

rs.Close
Set rs = Nothing
    
End Sub
Sub NewGenerateLabaDelayNo()
'On Error Resume Next
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Default1 = "Default" Then
        'a = rs!NoRujukanBayaranDelay
        'rs!NoRujukanBayaranDelay = a + 1
        'rs.Update
    End If
End If

rs.Close
Set rs = Nothing
    
End Sub
Sub NewGenerateEmpNo()
'On Error Resume Next
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
Set rs = New ADODB.Recordset
rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Default1 = "Default" Then
        EmpNoSeq = rs!EmpNo
        EmpNoSeq = EmpNoSeq + 1
        rs!EmpNo = EmpNoSeq
        rs.Update
    End If
End If

rs.Close
Set rs = Nothing
End Sub
Sub NewGenerateStockNo()
'On Error Resume Next
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from default_setting", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Default1 = "Default" Then
        OldNo = rs!NoRujukanStock
        NewNo = OldNo + 1
        rs!NoRujukanStock = NewNo
        rs.Update
    End If
End If
rs.Close
Set rs = Nothing
    
End Sub
