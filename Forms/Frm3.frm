VERSION 5.00
Begin VB.Form Frm3 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Login"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10335
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
   ForeColor       =   &H8000000B&
   Icon            =   "Frm3.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CMD3 
      BackColor       =   &H000080FF&
      Caption         =   "Batal"
      Height          =   405
      Left            =   3600
      MaskColor       =   &H00400000&
      MouseIcon       =   "Frm3.frx":0ECA
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   5400
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.ComboBox CBB1 
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
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4920
      Width           =   2900
   End
   Begin VB.CommandButton CMD1 
      BackColor       =   &H000080FF&
      Caption         =   "Batal"
      Height          =   405
      Left            =   3600
      MaskColor       =   &H00400000&
      MouseIcon       =   "Frm3.frx":11D4
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   5400
      Width           =   1305
   End
   Begin VB.TextBox TxtUsername 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      TabIndex        =   0
      Top             =   4200
      Width           =   2895
   End
   Begin VB.TextBox Txtpassword 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4560
      Width           =   2895
   End
   Begin VB.CommandButton CMD2 
      BackColor       =   &H000080FF&
      Caption         =   "Login"
      Height          =   405
      Left            =   2040
      MaskColor       =   &H00400000&
      MouseIcon       =   "Frm3.frx":14DE
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   5400
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton cmdlogin 
      BackColor       =   &H000080FF&
      Caption         =   "Login"
      Height          =   405
      Left            =   2040
      MaskColor       =   &H00400000&
      MouseIcon       =   "Frm3.frx":17E8
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   5400
      Width           =   1305
   End
   Begin VB.Label L1_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "Sila klik di sini jika anda lupa username atau password."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   240
      MouseIcon       =   "Frm3.frx":1AF2
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   5880
      Width           =   5655
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1920
      TabIndex        =   12
      Top             =   4935
      Width           =   135
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Wallpaper *"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      TabIndex        =   11
      Top             =   4935
      Width           =   1395
   End
   Begin VB.Label Label44 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "sankyusystem@gmail.com / 010 - 900 4788"
      BeginProperty Font 
         Name            =   "Berlin Sans FB"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   4800
      TabIndex        =   7
      Top             =   6120
      Width           =   5385
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username *"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   360
      Left            =   240
      TabIndex        =   6
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   360
      Left            =   1920
      TabIndex        =   5
      Top             =   4200
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password *"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   360
      Left            =   240
      TabIndex        =   4
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   360
      Left            =   1920
      TabIndex        =   3
      Top             =   4560
      Width           =   135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Sila masukkan Username dan Password untuk memasuki sistem ini."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   240
      TabIndex        =   2
      Top             =   3840
      Width           =   7815
   End
   Begin VB.Image Image1 
      Height          =   7500
      Left            =   0
      Picture         =   "Frm3.frx":1DFC
      Top             =   0
      Width           =   10365
   End
End
Attribute VB_Name = "Frm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMD1_Click()
End
End Sub
Private Sub CMD2_Click()
'On Error Resume Next
USER_FOUND = 0
If Frm3.TxtUsername = vbNullString Or Frm3.Txtpassword = vbNullString Then
    MsgBox "Sila masukkan semua ruangan yang wajib.", vbInformation, "Error"
    Frm3.TxtUsername.SetFocus
    Exit Sub
End If

If InStr(1, Frm3.TxtUsername, "*") <> 0 Or InStr(1, Frm3.TxtUsername, "/") <> 0 Or InStr(1, Frm3.TxtUsername, "\") <> 0 Or InStr(1, Frm3.TxtUsername, "'") <> 0 Then
    MsgBox "Simbol Yang Tidak Dibenarkan Di Dalam Username.", vbExclamation, "Login"
    Exit Sub
End If

If InStr(1, Frm3.Txtpassword, "*") <> 0 Or InStr(1, Frm3.Txtpassword, "/") <> 0 Or InStr(1, Frm3.Txtpassword, "\") <> 0 Or InStr(1, Frm3.Txtpassword, "'") <> 0 Then
    MsgBox "Simbol Yang Tidak Dibenarkan Di Dalam Password.", vbExclamation, "Login"
    Exit Sub
End If

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where Samaran='" & TxtUsername.Text & "' and password='" & Txtpassword.Text & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Status = "Aktif" Then
        If Not IsNull(rs!user_level) Then
            If rs!user_level = 1 Then
                User_Type = "Admin"
            ElseIf rs!user_level = 2 Then
                User_Type = "Manager"
            ElseIf rs!user_level = 3 Then
                User_Type = "Staff"
            ElseIf rs!user_level = 4 Then
                User_Type = "Guest/User"
            ElseIf rs!user_level = 5 Then
                User_Type = "Administration"
            ElseIf rs!user_level = 6 Then
                User_Type = "HQ"
            ElseIf rs!user_level = 7 Then
                User_Type = "Developer"
                G_LEVEL_USER = 7
            End If
        End If
        
'user_level
'1 : Admin
'2 : Manager
'3 : Staff
'4 : Guest/User -> Audit (Bagi menggunakan back end system 1 (external system)
'5 : Administration -> Audit (Bagi menggunakan back end system 2 (internal system)
'6 : HQ

        If User_Type <> "HQ" Then
            If Not IsNull(rs!cawangan) Then
                MDI_frm1.L20_Text = rs!cawangan
                G_CAWANGAN = rs!cawangan
                G_KEDAI = rs!cawangan
            End If
        End If
        
        If Not IsNull(rs!Samaran) Then User_Name = rs!Samaran
        USER_FOUND = 1
    Else
        MsgBox "Status anda adalah telah BERHENTI." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Oleh itu anda tidak dibenarkan untuk memasuki sistem ini." & vbCrLf & _
                "Sila hubungi pihak admin sistem.", vbExclamation, "Error"
    End If
Else
    MsgBox "User Tidak Wujud / Username & Password Tidak Sama.", vbInformation, "Login Tidak Berjaya"
    TxtUsername = vbNullString
    Txtpassword = vbNullString
    TxtUsername.SetFocus
End If

rs.Close
Set rs = Nothing

If USER_FOUND = 1 Then
GoTo bypass_dev:
    If G_LEVEL_USER = 7 Then
    
        Call check_internet_dev
        
        Note = "Sila masukkan Developer Pass."
    
        LM_PASSWORD = InputBox(Note, "Developer", "")
        
        If StrPtr(LM_PASSWORD) = 0 Then
            Exit Sub
        End If
        
        If LM_PASSWORD = G_DEV_PASS Then
        
        Else
        
            MsgBox "Pass Key yang dimasukkan TIDAK BETUL.", vbCritical, "Critical"
            Exit Sub
            
        End If
    
    End If
bypass_dev:
    MDI_frm1.L4_Text = "Staff"

    If Frm3.CBB1 <> vbNullString Then
        MDI_frm1.Picture = LoadPicture(App.Path & "\Backgrounds\" & Frm3.CBB1 & ".jpg")
    End If
    
    Unload Frm3
    MDI_frm1.L3_Text = User_Name 'User
    G_LOGIN_USER = User_Name 'User
    MDI_frm1.L4_Text = User_Type 'Level
    MDI_frm1.Pic5.Visible = False
    Call MDI_frm1_unload_admin_menu
    
    If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
        
        MDI_frm1.CMD44.Enabled = True
        
    Else
        
        MDI_frm1.CMD44.Enabled = False
    
    End If

    If User_Type = "HQ" Or User_Type = "Developer" Then

        Call Frm96_initial
        
        Frm96.Show vbModal
       
    Else

        Call main_setting_kedai
        Call main_setting
        
    End If
    
    If MDI_frm1.L4_Text = "Admin" Then
        G_LOCK_JURUJUAL = "NO"
    End If
    Call terminal_memory
    'If User_Type <> "Developer" Then Call check_license
    
End If
End Sub
Private Sub CMD3_Click()
'On Error Resume Next
Unload Frm3
End Sub
Private Sub cmdlogin_Click()
'On Error Resume Next
USER_FOUND = 0
If Frm3.TxtUsername = vbNullString Or Frm3.Txtpassword = vbNullString Then
    MsgBox "Sila masukkan semua ruangan yang wajib.", vbInformation, "Error"
    Frm3.TxtUsername.SetFocus
    Exit Sub
End If

If InStr(1, Frm3.TxtUsername, "*") <> 0 Or InStr(1, Frm3.TxtUsername, "/") <> 0 Or InStr(1, Frm3.TxtUsername, "\") <> 0 Or InStr(1, Frm3.TxtUsername, "'") <> 0 Then
    MsgBox "Simbol Yang Tidak Dibenarkan Di Dalam Username.", vbExclamation, "Login"
    Exit Sub
End If

If InStr(1, Frm3.Txtpassword, "*") <> 0 Or InStr(1, Frm3.Txtpassword, "/") <> 0 Or InStr(1, Frm3.Txtpassword, "\") <> 0 Or InStr(1, Frm3.Txtpassword, "'") <> 0 Then
    MsgBox "Simbol Yang Tidak Dibenarkan Di Dalam Password.", vbExclamation, "Login"
    Exit Sub
End If

MDI_frm1.L20_Text = vbNullString

G_LEVEL_USER = 3

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from employee where Samaran='" & TxtUsername.Text & "' and password='" & Txtpassword.Text & "'", cn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    If rs!Status = "Aktif" Then
        If Not IsNull(rs!user_level) Then
            If rs!user_level = 1 Then
                User_Type = "Admin"
                G_LEVEL_USER = 1
            ElseIf rs!user_level = 2 Then
                User_Type = "Manager"
                G_LEVEL_USER = 2
            ElseIf rs!user_level = 3 Then
                User_Type = "Staff"
                G_LEVEL_USER = 3
            ElseIf rs!user_level = 4 Then
                User_Type = "Guest/User"
                G_LEVEL_USER = 4
            ElseIf rs!user_level = 5 Then
                User_Type = "Administration"
                G_LEVEL_USER = 5
            ElseIf rs!user_level = 6 Then
                User_Type = "HQ"
                G_LEVEL_USER = 6
            ElseIf rs!user_level = 7 Then
                User_Type = "Developer"
                G_LEVEL_USER = 7
            End If
        End If
        
'user_level
'1 : Admin
'2 : Manager
'3 : Staff
'4 : Guest/User -> Audit (Bagi menggunakan back end system 1 (external system)
'5 : Administration -> Audit (Bagi menggunakan back end system 2 (internal system)
'6 : HQ
        
        If User_Type <> "HQ" Then
            If Not IsNull(rs!cawangan) Then
                MDI_frm1.L20_Text = rs!cawangan
                G_CAWANGAN = rs!cawangan
                G_KEDAI = rs!cawangan
            End If
        End If
        
        If Not IsNull(rs!Samaran) Then User_Name = rs!Samaran
        USER_FOUND = 1
    Else
        MsgBox "Status anda adalah telah BERHENTI." & vbCrLf & _
                vbNullString & vbCrLf & _
                "Oleh itu anda tidak dibenarkan untuk memasuki sistem ini." & vbCrLf & _
                "Sila hubungi pihak admin sistem.", vbExclamation, "Error"
    End If
Else
    MsgBox "User Tidak Wujud / Username & Password Tidak Sama.", vbInformation, "Login Tidak Berjaya"
    TxtUsername = vbNullString
    Txtpassword = vbNullString
    TxtUsername.SetFocus
End If

rs.Close
Set rs = Nothing

If USER_FOUND = 1 Then
    
    MDI_frm1.L16_Text = vbNullString 'Harga Besar
    'Call amendment_email_check
    Call main_setting
    GoTo bypass_dev:
    If G_LEVEL_USER = 7 Then
    
        Call check_internet_dev
        
        Note = "Sila masukkan Developer Pass."
    
        LM_PASSWORD = InputBox(Note, "Developer", "")
        
        If StrPtr(LM_PASSWORD) = 0 Then
            Exit Sub
        End If
        
        If LM_PASSWORD = G_DEV_PASS Then
        
        Else
        
            MsgBox "Pass Key yang dimasukkan TIDAK BETUL.", vbCritical, "Critical"
            Exit Sub
            
        End If
    
    End If
bypass_dev:
    MDI_frm1.Caption = "[Sistem Pengurusan Kedai Emas (Sankyu System)] SPKE106.1.22"
    MDI_frm1.L4_Text = "Staff"
    MDI_frm1.Show

    If Frm3.CBB1 <> vbNullString Then
        MDI_frm1.Picture = LoadPicture(App.Path & "\Backgrounds\" & Frm3.CBB1 & ".jpg")
    End If
    
    Unload Frm3
    MDI_frm1.L3_Text = User_Name 'User
    G_LOGIN_USER = User_Name 'User
    MDI_frm1.L4_Text = User_Type 'Level
    MDI_frm1.L16_Text = 0 'Harga Besar
    
    Frm96.CMD1.Visible = True
    Frm96.CMD2.Visible = False
    
    If MDI_frm1.L4_Text = "HQ" Or MDI_frm1.L4_Text = "Developer" Then
        
        MDI_frm1.CMD44.Enabled = True
        
    Else
        
        MDI_frm1.CMD44.Enabled = False
    
    End If
    
    If User_Type = "HQ" Or User_Type = "Developer" Then

        Call Frm96_initial
        
        Frm96.Show vbModal
        
    Else
    
        Call main_setting_kedai
        Call main_setting
        
    End If
    Call terminal_memory
    If MDI_frm1.L4_Text = "Admin" Then
        G_LOCK_JURUJUAL = "NO"
    End If
    
    'If User_Type <> "Developer" Then Call check_license
    
End If
End Sub
Private Sub Form_Load()
'On Error Resume Next
LM_FOUND = 0

Call system_configuration

Frm3.CBB1.Clear

Set rs = New ADODB.Recordset
If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
rs.Open "select * from 55_wallpaper", cn, adOpenKeyset, adLockOptimistic

While rs.EOF = False
    If Not IsNull(rs!wallpaper_no) Then
        Frm3.CBB1.AddItem rs!wallpaper_no
        If LM_FOUND = 0 Then
            LM_WALLPAPER = rs!wallpaper_no
        End If
        LM_FOUND = 1
    End If
    rs.MoveNext
Wend

rs.Close
Set rs = Nothing

If LM_FOUND = 1 Then
    Frm3.CBB1 = LM_WALLPAPER
End If
End Sub

Private Sub L1_Text_Click()
'On Error Resume Next

Note = "Sila masukkan e-mail anda yang didaftarkan ke dalam sistem ini." & vbCrLf & _
        vbNullString & vbCrLf & _
        "Sistem akan menghantar USERNAME dan PASSWORD anda ke email ini."

G_MAIL = InputBox(Note, "Username dan password", "Masukkan e-mail anda")

If StrPtr(G_MAIL) = 0 Then
    Exit Sub
End If

If StrPtr(G_MAIL) <> 0 Then

    myAt = InStr(1, G_MAIL, "@", vbTextCompare)
    myDot = InStr(myAt + 2, G_MAIL, ".", vbTextCompare)
    myDotDot = InStr(myAt + 2, G_MAIL, "..", vbTextCompare)
    
    If myAt = 0 Or myDot = 0 Or Not myDotDot = 0 Or Right(G_MAIL, 1) = "." Then
        MsgBox "E-mail yang tidak sah.", vbExclamation, "Info"
        
        Exit Sub
    End If

    Call Frm3_check_internet
    
End If

End Sub
Private Sub TxtUsername_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txtpassword.SetFocus
End If
End Sub
Private Sub Txtpassword_KeyPress(KeyAscii As Integer)
If Frm3.cmdlogin.Visible = True Then
    If KeyAscii = 13 Then
        Frm3.cmdlogin.SetFocus
    End If
Else
    If KeyAscii = 13 Then
        Frm3.CMD2.SetFocus
    End If
End If
End Sub
