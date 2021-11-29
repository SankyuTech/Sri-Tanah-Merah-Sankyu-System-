Attribute VB_Name = "Module5"
Sub Frm95_initial_setting()
'on error resume next
GLOBAL_DISABLE = 0

Frm95.Frame1.Left = 120
Frm95.Frame1.Top = 2520

Frm95.Frame1.Visible = False

Frm95.TB1 = vbNullString
Frm95.TB2 = vbNullString

Frm95.CMD1.Visible = True
Frm95.CMD2.Visible = False
Frm95.CMD3.Visible = False

Frm95.L12_Text.Visible = False
End Sub
Sub Frm97_initial_setting()
'on error resume next
Frm97.Pic1.Left = 120
Frm97.Pic1.Top = 240
Frm97.Pic2.Left = 120
Frm97.Pic2.Top = 240

Frm97.Pic1.Visible = False
Frm97.Pic2.Visible = False

Frm97.TB1 = vbNullString
Frm97.TB2 = vbNullString

Frm97.L6_Text = vbNullString
Frm97.L7_Text = vbNullString

Frm97.L6_Text.BackStyle = 0
Frm97.L7_Text.BackStyle = 0
End Sub
Sub amendment_email_check()
'On Error Resume Next
Dim strsql As String
Dim DATA_QTY As Double

DATA_QTY = 0

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select COUNT(ID) from 72_data_amendment", cn, adOpenKeyset, adLockOptimistic

If Not IsNull(rs(0)) Then DATA_QTY = rs(0)

rs.Close
Set rs = Nothing


If DATA_QTY <> 0 Then

    Frm111.Show
    
    Shell "cmd.exe /c " & G_NE_PATH
    
    Unload Frm111
    
End If

Exit Sub

'###Padam Table Jualan Temp### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub

strsql = "TRUNCATE TABLE 72_data_amendment"

Set rs = cn.Execute(strsql)
Set rs = Nothing
'###Padam Table Jualan Temp### - End

End Sub
