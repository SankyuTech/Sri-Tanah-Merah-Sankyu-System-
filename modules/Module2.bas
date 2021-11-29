Attribute VB_Name = "Module2"
Sub UpdateLog_Main()
'On Error Resume Next
'Application.ScreenUpdating = False
Frm2.MSFlexGrid1.Clear
Frm2.MSFlexGrid1.FormatString = "< Tarikh Dan Masa |< Log Aktiviti"
Frm2.MSFlexGrid1.ColWidth(0) = 2300
Frm2.MSFlexGrid1.ColWidth(1) = 8200
    
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

Frm2.L9_Text = "Update Terkini : " & DateTime.Date & " " & DateTime.Time
'Application.ScreenUpdating = True
End Sub
