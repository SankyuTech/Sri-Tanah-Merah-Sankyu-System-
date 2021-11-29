VERSION 5.00
Begin VB.Form Frm7 
   Caption         =   "Tukar Password"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5475
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm7.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3060
   ScaleWidth      =   5475
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD1 
      Caption         =   "Tukar Password"
      Height          =   375
      Left            =   1920
      MouseIcon       =   "Frm7.frx":0ECA
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox TB4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1900
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox TB3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1900
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox TB1 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1900
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox TB2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1900
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   2070
      Width           =   255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   1720
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   1380
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   1000
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Sila masukkan password yang lama dan masukkan password yang baru yang ingin ditukar."
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Password baru *"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   2070
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password baru *"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1720
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password lama *"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1380
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User *"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1000
      Width           =   1455
   End
End
Attribute VB_Name = "Frm7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD1_Click()
'On Error Resume Next
DATA_FOUND = 0

If InStr(1, Frm7.TB2, "*") <> 0 Or InStr(1, Frm7.TB2, "/") <> 0 Or InStr(1, Frm7.TB2, "\") <> 0 Or InStr(1, Frm7.TB2, "'") <> 0 Then
    MsgBox "Password lama mengandungi simbol yang tidak dibenarkan.", vbExclamation, "Error"
    Exit Sub
End If

If InStr(1, Frm7.TB3, "&") <> 0 Or InStr(1, Frm7.TB3, "*") <> 0 Or InStr(1, Frm7.TB3, "/") <> 0 Or InStr(1, Frm7.TB3, "\") <> 0 Or InStr(1, Frm7.TB3, "'") <> 0 Then
    MsgBox "Password baru mengandungi simbol yang tidak dibenarkan.", vbExclamation, "Error"
    Exit Sub
End If

If InStr(1, Frm7.TB4, "&") <> 0 Or InStr(1, Frm7.TB4, "*") <> 0 Or InStr(1, Frm7.TB4, "/") <> 0 Or InStr(1, Frm7.TB4, "\") <> 0 Or InStr(1, Frm7.TB4, "'") <> 0 Then
    MsgBox "Password baru mengandungi simbol yang tidak dibenarkan.", vbExclamation, "Error"
    Exit Sub
End If

If UCase(Frm7.TB3) <> UCase(Frm7.TB4) Then
    MsgBox "Password yang baru tidak sama.", vbInformation, "Error"
    Exit Sub
End If

If Frm7.TB2 = vbNullString Or Frm7.TB3 = vbNullString Or Frm7.TB4 = vbNullString Then
    MsgBox "Sila masukkan semua ruangan yang wajib.", vbInformation, "Error"
    Frm7.TB2.SetFocus
    Exit Sub
End If

If Frm7.TB1 = vbNullString Then
    MsgBox "Tiada maklumat bagi USER.", vbInformation, "Error"
    Exit Sub
End If

Note = "Adakah anda ingin tukar password ?"

Answer = MsgBox(Note, vbQuestion + vbYesNo, "Confirmation")

If Answer = vbYes Then

'### Periksa password lama betul atau tidak ### - Start
    Set rs = New ADODB.Recordset
    If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
    rs.Open "select * from employee where Samaran='" & Frm7.TB1 & "'", cn, adOpenKeyset, adLockOptimistic
    
    If Not rs.EOF Then
        
        If Not IsNull(rs!Password) Then
            If UCase(rs!Password) <> UCase(Frm7.TB2) Then
                
                MsgBox "Password lama yang dimasukkan tidak sama dengan sistem.", vbExclamation, "Error"
                
                rs.Close
                Set rs = Nothing
                
                Exit Sub
            End If
        End If
        
        DATA_FOUND = 1
        
    End If
    
    rs.Close
    Set rs = Nothing
'### Periksa password lama betul atau tidak ### - End

    If DATA_FOUND = 1 Then
        
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from employee where Samaran='" & Frm7.TB1 & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            
            rs!Password = Frm7.TB3 'Password baru
            rs.Update
            
            DATA_FOUND = 2
            
        End If
        
        rs.Close
        Set rs = Nothing
        
    End If
    
    If DATA_FOUND = 2 Then

        '##########################################################
        '#                      Update Log                        #
        '##########################################################
        user = MDI_frm1.L3_Text
        LogDate_Memory = DateTime.Date & " " & DateTime.Time$
        LogAct_Memory = "[" & user & "] Tukar password."
        Call UpdateLog_Database
        '##########################################################
        '#                     @Update Log                        #
        '##########################################################
        
        Unload Frm7
        
        MsgBox "Password telah BERJAYA ditukar." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Sila gunakan password yang baru untuk memasuki sistem.", vbInformation, "Tukar Password"
                
    Else
    
        MsgBox "Anda tidak berjaya untuk menukar password anda." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Sila keluar dari sistem dan cuba sekali lagi.", vbExclamation, "Error"
        
    End If


End If
End Sub
Private Sub TB1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Frm7.TB2.SetFocus
End If
End Sub
Private Sub TB2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Frm7.TB3.SetFocus
End If
End Sub
Private Sub TB3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Frm7.TB4.SetFocus
End If
End Sub
Private Sub TB4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Frm7.CMD1.SetFocus
End If
End Sub
