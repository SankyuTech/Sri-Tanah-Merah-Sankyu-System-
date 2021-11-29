VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Frm1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selamat Datang"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7035
   ControlBox      =   0   'False
   FillColor       =   &H0000FF00&
   Icon            =   "Frm1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   PaletteMode     =   2  'Custom
   ScaleHeight     =   558
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   469
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   6240
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   661
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      Max             =   10
   End
   Begin VB.Timer tmrInfo 
      Interval        =   150
      Left            =   1080
      Top             =   120
   End
   Begin VB.Timer tmrLoading 
      Interval        =   100
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer Tmr1 
      Interval        =   900
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "http://sankyutech-visualbasic.weebly.com"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   375
      Left            =   240
      MouseIcon       =   "Frm1.frx":0ECA
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   8040
      Width           =   6375
   End
   Begin VB.Label Period 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lblSplash 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait ..."
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   5880
      Width           =   6495
   End
   Begin VB.Label L7 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading System"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   5400
      Width           =   6375
   End
   Begin VB.Label L6 
      BackStyle       =   0  'Transparent
      Caption         =   "https://www.facebook.com/ExcelVisualBasicApplicationVba"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   255
      Left            =   240
      MouseIcon       =   "Frm1.frx":11D4
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   7800
      Width           =   6615
   End
   Begin VB.Label L5 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail : sankyusystem@gmail.com"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   7560
      Width           =   5055
   End
   Begin VB.Label L4 
      BackStyle       =   0  'Transparent
      Caption         =   "No Telefon : +6010 - 900 4788"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   7320
      Width           =   6135
   End
   Begin VB.Label L3 
      BackStyle       =   0  'Transparent
      Caption         =   "Developer : Sankyu System"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   6960
      Width           =   5055
   End
   Begin VB.Label L2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "COPYRIGHT © Sankyu System 2013"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   6720
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   7365
      Left            =   0
      Picture         =   "Frm1.frx":14DE
      Top             =   -600
      Width           =   8415
   End
End
Attribute VB_Name = "Frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                (ByVal hwnd As Long, _
                ByVal lpOperation As String, _
                ByVal lpFile As String, _
                ByVal lpParameters As String, _
                ByVal lpDirectory As String, _
                ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
'on error resume next
'Check for previous instance and exit if found.
Dim rc As Long

If App.PrevInstance Then

    rc = MsgBox("Sistem Pengurusan Kedai Emas Telah Dibuka Sebelum Ini.", vbCritical, App.Title)
    
    End

End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next
If ProgressBar1.Value < ProgressBar1.Max Then Cancel = True
End Sub
Private Sub L6_Click()
'On Error Resume Next
BrowseTo "https://www.facebook.com/ExcelVisualBasicApplicationVba"
End Sub
Private Sub BrowseTo(ByRef pstrURL As String)
'On Error Resume Next
' Opens users default web browser and navigates to the selected URL
Call ShellExecute(Me.hwnd, "Open", pstrURL, "", "", True)
End Sub
Private Sub Label1_Click()
'On Error Resume Next
BrowseTo "http://sankyutech-visualbasic.weebly.com"
End Sub

Private Sub Tmr1_Timer()
'On Error Resume Next
ProgressBar1.Max = 180
ProgressBar1.Value = Int(ProgressBar1.Value) + 20
If ProgressBar1.Value >= ProgressBar1.Max Then
    Unload Me
    Frm3.Show
    'Call Expiry
    'Frm3.Show
    'Frm3.TxtUsername.SetFocus
End If

End Sub
Private Sub tmrInfo_Timer()
Static counting As Integer
'On Error Resume Next
counting = counting + 1

If counting = 3 Then
    lblSplash.Caption = "Preparing ..."
ElseIf counting = 10 Then
    lblSplash.Caption = "Loading forms ..."
ElseIf counting = 20 Then
    lblSplash.Caption = "Checking Connection System And Databases ..."
    
    Call system_configuration
    Call check_internet_connection_main
    If G_SYSTEM_TYPE = "ONLINE" Then
        If MDI_frm1.L17_Text = "ONLINE" Then
        
            Call Main
            'Call main_setting
            
        Else
        
            MsgBox "Tiada sambungan internet. Sila pastikan komputer anda disambungkan dengan internet bagi membolehkan sistem beroperasi.", vbCritical, App.Title
            
            End
            
        End If
    End If
'### Periksa conn database dan mode bagi auto backup ### - Start
    'Sila masukkan path bagi auto backup ke dalam field "ab_link"
    'Pastikan jika ada terdapat update pada sistem auto backup atau perubahan pada path auto backup , perlu update juga link / path dalam field ini.
    
    GoTo aa:
    
    If G_AUTO_BACKUP = "YES" Then
    
        Set rs = New ADODB.Recordset
        If (MDI_frm1.L17_Text = "ONLINE" Or G_SYSTEM_TYPE = "OFFLINE") Then Call check_db_conn_main Else Exit Sub
        rs.Open "select * from 56_maklumat_kedai where default_setting='" & "default" & "'", cn, adOpenKeyset, adLockOptimistic
        
        If Not rs.EOF Then
            If Not IsNull(rs!auto_backup) Then
                If rs!auto_backup = 1 Then
                    If Not IsNull(rs!ab_link) Then Shell rs!ab_link
                End If
            End If
        End If
        
        rs.Close
        Set rs = Nothing
        
    End If
    
aa:
'### Periksa conn database dan mode bagi auto backup ### - End
ElseIf counting = 30 Then
    lblSplash.Caption = "Loading Databases ..."
ElseIf counting = 40 Then
    tmrLoading.Enabled = False
    lblSplash.Caption = "Done ..."
ElseIf counting = 45 Then
    lblSplash.Caption = "Starting Sistem Pengurusan Kedai Emas ..."
ElseIf counting = 60 Then
    'tmrInfo.Enabled = False
    'Unload Me
    'frmLogIn.Show
End If
End Sub
Private Sub tmrLoading_Timer()
'On Error Resume Next
'If Disk1.Visible = True Then
'    Disk1.Visible = False
'    Disk2.Visible = True
'Else
'    Disk1.Visible = True
'    Disk2.Visible = False
'End If
End Sub

