Attribute VB_Name = "mod_hantar_mail_user"
Public Function SendMail(sTo As String, sSubject As String, sFrom As String, _
    sBody As String, sSmtpServer As String, iSmtpPort As Integer, _
    sSmtpUser As String, sSmtpPword As String, _
    sFilePath As String, bSmtpSSL As Boolean) As String
      
    On Error GoTo SendMail_Error:
    Dim lobj_cdomsg      As CDO.Message
    Set lobj_cdomsg = New CDO.Message
    lobj_cdomsg.Configuration.Fields(cdoSMTPServer) = sSmtpServer
    lobj_cdomsg.Configuration.Fields(cdoSMTPServerPort) = iSmtpPort
    lobj_cdomsg.Configuration.Fields(cdoSMTPUseSSL) = bSmtpSSL
    lobj_cdomsg.Configuration.Fields(cdoSMTPAuthenticate) = cdoBasic
    lobj_cdomsg.Configuration.Fields(cdoSendUserName) = sSmtpUser
    lobj_cdomsg.Configuration.Fields(cdoSendPassword) = sSmtpPword
    lobj_cdomsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
    lobj_cdomsg.Configuration.Fields(cdoSendUsingMethod) = cdoSendUsingPort
    lobj_cdomsg.Configuration.Fields.Update
    lobj_cdomsg.To = sTo
    lobj_cdomsg.From = sFrom
    lobj_cdomsg.Subject = sSubject
    lobj_cdomsg.TextBody = sBody
    If Trim$(sFilePath) <> vbNullString Then
        lobj_cdomsg.AddAttachment (sFilePath)
    End If
    lobj_cdomsg.Send
    Set lobj_cdomsg = Nothing
    SendMail = "ok"
    Exit Function
          
SendMail_Error:
    SendMail = Err.Description
End Function
Sub Frm3_hantar_mail()
'on error resume next
Dim A1 As String
Dim B1 As String
Dim C1 As String
Dim D1 As String
Dim E1 As String

Dim retVal          As String
Dim objControl      As Control

DATA_FOUND = 0

'### Periksa kewujudan email ### - Start
Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where mail='" & G_MAIL & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    
    If Not IsNull(rs!Nama) Then Frm3_LM_NAMA = rs!Nama
    If Not IsNull(rs!NoIC) Then Frm3_LM_IC = rs!NoIC
    If Not IsNull(rs!Samaran) Then Frm3_LM_USER = rs!Samaran
    If Not IsNull(rs!Password) Then Frm3_LM_PWD = rs!Password
    
    DATA_FOUND = 1
End If

rs.Close
Set rs = Nothing
'### Periksa kewujudan email ### - End

If DATA_FOUND = 1 Then

    Frm3.Hide

    A1 = A1 & "Assalamualaikum & selamat sejahtera." & vbCrLf
    A1 = A1 & vbNullString & vbCrLf
    A1 = A1 & "Sila gunakan maklumat di bawah untuk memasuki sistem." & vbCrLf
    
    A1 = A1 & "-----------------------------------------------------------" & vbCrLf
    A1 = A1 & "Nama               : " & Frm3_LM_NAMA & vbCrLf
    A1 = A1 & "No. Kad Pengenalan : " & Frm3_LM_IC & vbCrLf
    A1 = A1 & "Username           : " & Frm3_LM_USER & vbCrLf
    A1 = A1 & "Password           : " & Frm3_LM_PWD & vbCrLf
    A1 = A1 & "-----------------------------------------------------------" & vbCrLf
    
    A1 = A1 & "Terima kasih." & vbCrLf

    C1 = C1 & vbNullString & vbCrLf
    C1 = C1 & "This is an auto-generated e-mail by Sankyu System. Please do not reply to this e-mail address." & vbCrLf
    C1 = C1 & vbNullString & vbCrLf
    C1 = C1 & "Any enquiries please feel free to contact Sankyu System at +6010 - 900 4788 , sankyusystem@gmail.com ." & vbCrLf
    
    ' Set email body
    B1 = A1 & vbCrLf & _
        C1

    '### Maklumat kedai ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from 56_maklumat_kedai where cawangan='" & G_KEDAI & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        If Not IsNull(rs!nama_kedai) Then Frm3_LM_NAMA_KEDAI = rs!nama_kedai
        
        If Not IsNull(rs!m_add) Then Frm3_LM_ADDR = rs!m_add
        If Not IsNull(rs!m_pass) Then Frm3_LM_M_PWD = rs!m_pass
        If Not IsNull(rs!Port) Then Frm3_LM_PORT = rs!Port
        If Not IsNull(rs!server) Then Frm3_LM_SERVER = rs!server
    End If
    
    rs.Close
    Set rs = Nothing
    '### Maklumat kedai ### - End
    
    E_SUBJECT = "[" & Frm3_LM_NAMA_KEDAI & "] Username & Password bagi memasuki sistem pengurusan kedai emas."

    'cmdSend.Enabled = False
    retVal = SendMail(Trim$(G_MAIL), _
        Trim$(E_SUBJECT), _
        Trim$("Sankyu System (Admin)") & "<" & Trim$(Frm3_LM_ADDR) & ">", _
        Trim$(B1), _
        Trim$(Frm3_LM_SERVER), _
        CInt(Trim$(Frm3_LM_PORT)), _
        Trim$(Frm3_LM_ADDR), _
        Trim$(Frm3_LM_M_PWD), _
        Trim$(txtAttach), _
        CBool(1))
        
        'CBool(Frm1.chkSSL.Value))
    'cmdSend.Enabled = True
    
    Frm3.Show
    
    MsgBox "Maklumat bagi username dan password telah dihantar ke email anda." & vbCrLf & _
            vbNullString & vbCrLf & _
            G_MAIL

End If
End Sub

