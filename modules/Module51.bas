Attribute VB_Name = "Module51"
'Working with registry declarations and constants
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Const ERROR_SUCCESS = 0&
Private Const APINULL = 0&
Private Const HKEY_LOCAL_MACHINE = &H80000002
'Working with wininet.dll declarations and constants
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" Alias "InternetGetConnectedStateExA" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Long, ByVal dwReserved As Long) As Long 'Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long 'this function used with IE4
'Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long 'this function used with IE4
Private Const INTERNET_CONNECTION_MODEM = &H1&
Private Const INTERNET_CONNECTION_LAN = &H2&
Private Const INTERNET_CONNECTION_PROXY = &H4&
Private Const INTERNET_RAS_INSTALLED = &H10&
Private Const INTERNET_CONNECTION_OFFLINE = &H20&
Private Const INTERNET_CONNECTION_CONFIGURED = &H40&
'Declares for direct ping
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInet As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long

Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Dim checkType As Integer
Dim remMsg(2) As String
Sub check_internet_dev(Optional ByRef ConnectionInfo As Long, Optional ByRef sConnectionName As String)
'On Error GoTo logging:
Dim dwFlags As Long
Dim sNameBuf As String, msg As String
Dim lPos As Long
sNameBuf = String$(513, 0)

G_DEV_PASS = G_DEV_PASS_DEFAULT

If InternetGetConnectedStateEx(dwFlags, sNameBuf, 512, 0&) Then
    Call generate_dev_pass
'online mode
Else
    G_DEV_PASS = G_DEV_PASS_DEFAULT
'offline mode
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Module51 : check_internet_dev" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Sub generate_dev_pass()
'On Error GoTo logging:
all_chars = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z")

Randomize
For i = 1 To 10
    random_index = Int(Rnd() * 61)
    clave = clave & all_chars(random_index)
Next

G_DEV_PASS = clave

Call send_login_tele

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Module51 : generate_dev_pass" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub
Sub check_license()
'On Error GoTo logging:
LM_FOUND = 0

Enter = Chr$(13) + Chr$(10)
'Os Information
Set SystemSet = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")

For Each System20 In SystemSet
    LM_NAME = System20.SerialNumber
Next

LM_CONN = 1
re_conn_1:

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 91_senarai_terminal where station_id='" & LM_NAME & "' AND terminal='" & G_TERMINAL & "' AND status = 1", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then

    LM_FOUND = 1

End If

rs.Close
Set rs = Nothing

If LM_FOUND = 0 Then
    MsgBox "LICENSE TIDAK SAH. SILA HUBUNGI PIHAK SANKYU SYSTEM UNTUK URUSAN PEMBELIAN LICENSE. [INSAN : +6010 - 900 4788]", vbCritical, "Critical"
    End
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Module51 : check_license" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
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
Sub register_license()
'On Error GoTo logging:
LM_FOUND = 0

Enter = Chr$(13) + Chr$(10)
'Os Information
Set SystemSet = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")

For Each System20 In SystemSet
    LM_NAME = System20.SerialNumber
Next

LM_NOW = Now

LM_CONN = 1
re_conn_1:

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 91_senarai_terminal where (station_id='" & LM_NAME & "' OR terminal='" & G_TERMINAL & "') AND status = 1", cn, adOpenKeyset, adLockOptimistic

If rs.EOF Then
    rs.AddNew
    rs!station_id = LM_NAME
    rs!terminal = G_TERMINAL
    rs!write_timestamp = LM_NOW
    rs!Status = 1
    rs.Update
    LM_FOUND = 1
Else
    MsgBox "TERMIMAL atau STATION ID telah wujud. Anda tidak dibenarkan untuk daftarkan LICENSE ini.", vbExclamation, "Info"
End If

rs.Close
Set rs = Nothing

If LM_FOUND = 1 Then
    
    '** Send telegram - Start
    Dim A1 As String
    Dim File_Path As String
    
    File_Path = App.Path & "\logging_configuration.txt"
    Open File_Path For Input As #1
    
    Line Input #1, LM_TELE_TOKEN
    Line Input #1, LM_CHAT_ID
    Line Input #1, LM_NAMA_KEDAI
    Line Input #1, LM_VERSION_SYSTEM
    Line Input #1, LM_VERSION_DATABASE
    Line Input #1, LM_CAWANGAN
    Line Input #1, LM_STATION
    Line Input #1, LM_LIMIT_ERROR
    Line Input #1, LM_CHAT_ID_LOGIN
    
    Close #1
    
    App.LogEvent "License registration for station " & LM_NAME & " , terminal " & G_TERMINAL & ".", vbLogEventTypeInformation
    
    
    A1 = A1 & "** License Registration **" & "%0A" & vbCrLf
    A1 = A1 & vbNullString & vbCrLf
    A1 = A1 & "Nama Kedai : " & LM_NAMA_KEDAI & "%0A" & vbCrLf
    A1 = A1 & "Cawangan : " & LM_CAWANGAN & "%0A" & vbCrLf
    A1 = A1 & "Station : " & LM_STATION & "%0A" & vbCrLf
    A1 = A1 & "System Version : " & LM_VERSION_SYSTEM & "%0A" & vbCrLf
    A1 = A1 & "Database Version : " & LM_VERSION_DATABASE & "%0A" & vbCrLf
    A1 = A1 & vbNullString & vbCrLf
    A1 = A1 & "Station ID : " & LM_NAME & "%0A" & vbCrLf
    
    strURL = LM_TELE_TOKEN & "/" & "sendmessage?chat_id=" & LM_CHAT_ID_LOGIN & "&text=" & A1 & ""
    
    Set XMLHttpRequest = New MSXML2.XMLHTTP60
    XMLHttpRequest.Open "GET", strURL, False
    'XMLHttpRequest.setRequestHeader "Content-Type", "text/xml"
    XMLHttpRequest.Send
    
    strResponse = XMLHttpRequest.responseText
    Set XMLHttpRequest = Nothing

    '** Send telegram - End
    
    user = MDI_frm1.L3_Text
    LogAct_Memory = "[" & user & "] Register LICENSE bagi station ID [" & LM_NAME & "] , Terminal [" & G_TERMINAL & "]."
    LogDate_Memory = LM_NOW
    Call UpdateLog_Database
    
    MsgBox "LICENSE bagi station ini telah berjaya didaftarkan.", vbInformation, "Info"
    
End If

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Module51 : register_license" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
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
Sub log_rekod()
'On Error Resume Next
Dim sFilename As String
sFilename = App.Path & "\logging.txt"

' Archive file at certain size
If FileLen(sFilename) > 200000 Then
    FileCopy sFilename _
        , Replace(sFilename, ".txt", Format(Now, "ddmmyyyy hhmmss.txt"))
    Kill sFilename
End If

' Open the file to write
Dim filenumber As Variant
filenumber = FreeFile
Open sFilename For Append As #filenumber

Print #filenumber, G_ERROR_NAIYO

Close #filenumber

Call send_report_tele
End Sub
Sub send_report_tele()
'On Error Resume Next
Dim A1 As String
Dim File_Path As String
Dim LM_LIMIT As Integer

File_Path = App.Path & "\logging_configuration.txt"
Open File_Path For Input As #1

Line Input #1, LM_TELE_TOKEN
Line Input #1, LM_CHAT_ID
Line Input #1, LM_NAMA_KEDAI
Line Input #1, LM_VERSION_SYSTEM
Line Input #1, LM_VERSION_DATABASE
Line Input #1, LM_CAWANGAN
Line Input #1, LM_STATION
Line Input #1, LM_LIMIT_ERROR

Close #1

LM_LIMIT = LM_LIMIT_ERROR

G_X = G_X + 1

If G_X >= LM_LIMIT Then
    
    MsgBox "Telah berlaku error di dalam sistem. Sila hubungi pihak Sankyu System." & vbCrLf & _
            vbNullString & vbCrLf & _
            "Sistem akan ditutup.", vbCritical, "Error"
            
    End
    
End If

A1 = A1 & "Error detected :" & "%0A" & vbCrLf
A1 = A1 & vbNullString & vbCrLf
A1 = A1 & "Nama Kedai : " & LM_NAMA_KEDAI & "%0A" & vbCrLf
A1 = A1 & "Cawangan : " & LM_CAWANGAN & "%0A" & vbCrLf
A1 = A1 & "Station : " & LM_STATION & "%0A" & vbCrLf
A1 = A1 & "System Version : " & LM_VERSION_SYSTEM & "%0A" & vbCrLf
A1 = A1 & "Database Version : " & LM_VERSION_DATABASE & "%0A" & vbCrLf
A1 = A1 & "Error : " & G_ERROR_NAIYO & "%0A" & vbCrLf
A1 = A1 & "Error Number : " & Err.Number & "%0A" & vbCrLf
A1 = A1 & "Error Description : " & Err.Description & "%0A" & vbCrLf

strURL = LM_TELE_TOKEN & "/" & "sendmessage?chat_id=" & LM_CHAT_ID & "&text=" & A1 & ""

Set XMLHttpRequest = New MSXML2.XMLHTTP60
XMLHttpRequest.Open "GET", strURL, False
'XMLHttpRequest.setRequestHeader "Content-Type", "text/xml"
XMLHttpRequest.Send

strResponse = XMLHttpRequest.responseText
Set XMLHttpRequest = Nothing
End Sub
Sub send_login_tele()
'On Error GoTo logging:
Dim A1 As String
Dim File_Path As String
Dim LM_LIMIT As Integer
Dim LM_ERR_DETAIL As String

File_Path = App.Path & "\logging_configuration.txt"
Open File_Path For Input As #1

Line Input #1, LM_TELE_TOKEN
Line Input #1, LM_CHAT_ID
Line Input #1, LM_NAMA_KEDAI
Line Input #1, LM_VERSION_SYSTEM
Line Input #1, LM_VERSION_DATABASE
Line Input #1, LM_CAWANGAN
Line Input #1, LM_STATION
Line Input #1, LM_LIMIT_ERROR
Line Input #1, LM_CHAT_ID_LOGIN

Close #1

App.LogEvent "Request Developer Pass", vbLogEventTypeInformation


A1 = A1 & "**Request Developer Pass**" & "%0A" & vbCrLf
A1 = A1 & vbNullString & vbCrLf
A1 = A1 & "Nama Kedai : " & LM_NAMA_KEDAI & "%0A" & vbCrLf
A1 = A1 & "Cawangan : " & LM_CAWANGAN & "%0A" & vbCrLf
A1 = A1 & "Station : " & LM_STATION & "%0A" & vbCrLf
A1 = A1 & "System Version : " & LM_VERSION_SYSTEM & "%0A" & vbCrLf
A1 = A1 & "Database Version : " & LM_VERSION_DATABASE & "%0A" & vbCrLf
A1 = A1 & vbNullString & vbCrLf
A1 = A1 & "Pass : **" & G_DEV_PASS & "**" & "%0A" & vbCrLf

strURL = LM_TELE_TOKEN & "/" & "sendmessage?chat_id=" & LM_CHAT_ID_LOGIN & "&text=" & A1 & ""

Set XMLHttpRequest = New MSXML2.XMLHTTP60
XMLHttpRequest.Open "GET", strURL, False
'XMLHttpRequest.setRequestHeader "Content-Type", "text/xml"
XMLHttpRequest.Send

strResponse = XMLHttpRequest.responseText
Set XMLHttpRequest = Nothing

Exit Sub

logging:

LM_ERR_NO = Err.Number

G_ERROR_NAIYO = CStr(Now) & " Module51 : send_login_tele" & " / " & LM_CONN & " / " & Err.Number & " / " & Err.Description
Call log_rekod

Resume Next
End Sub

